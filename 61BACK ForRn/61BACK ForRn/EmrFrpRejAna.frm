VERSION 5.00
Object = "{8767A745-088E-4CA6-8594-073D6D2DE57A}#9.2#0"; "crviewer9.dll"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "tabctl32.ocx"
Object = "{AB2CC32A-E652-42A2-B341-26ECE02F6386}#2.1#0"; "mwfctlforspeed9.ocx"
Begin VB.Form EmrFrpRejAna 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Rejection Analysis Report"
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
      Left            =   11370
      TabIndex        =   63
      TabStop         =   0   'False
      Top             =   9600
      Visible         =   0   'False
      Width           =   495
      _ExtentX        =   873
      _ExtentY        =   873
   End
   Begin MwfCtl.MWCTL_DST DST_REP 
      Height          =   570
      Left            =   4455
      TabIndex        =   62
      TabStop         =   0   'False
      Top             =   9630
      Width           =   6330
      _ExtentX        =   11165
      _ExtentY        =   1005
   End
   Begin VB.Frame fra_keys 
      BackColor       =   &H00C0C0C0&
      BorderStyle     =   0  'None
      Height          =   465
      Left            =   5325
      TabIndex        =   68
      Top             =   9630
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
      Left            =   6540
      TabIndex        =   67
      Top             =   9645
      Visible         =   0   'False
      Width           =   855
      Begin CRVIEWER9LibCtl.CRViewer9 CRV_REP 
         Height          =   495
         Left            =   480
         TabIndex        =   66
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
      TabIndex        =   64
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
      TabIndex        =   65
      Top             =   60
      Width           =   15255
      Begin TabDlg.SSTab TAB_REP 
         Height          =   9465
         Left            =   60
         TabIndex        =   71
         Top             =   30
         Width           =   15195
         _ExtentX        =   26802
         _ExtentY        =   16695
         _Version        =   393216
         Tabs            =   2
         TabsPerRow      =   2
         TabHeight       =   520
         TabCaption(0)   =   "&Data Range"
         TabPicture(0)   =   "EmrFrpRejAna.frx":0000
         Tab(0).ControlEnabled=   -1  'True
         Tab(0).Control(0)=   "fra_tabrep(0)"
         Tab(0).Control(0).Enabled=   0   'False
         Tab(0).ControlCount=   1
         TabCaption(1)   =   "&Group Sort"
         TabPicture(1)   =   "EmrFrpRejAna.frx":001C
         Tab(1).ControlEnabled=   0   'False
         Tab(1).Control(0)=   "fra_tabrep(1)"
         Tab(1).ControlCount=   1
         Begin VB.Frame fra_tabrep 
            BorderStyle     =   0  'None
            Height          =   9070
            Index           =   0
            Left            =   60
            TabIndex        =   74
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
               TabIndex        =   108
               ToolTipText     =   "Select Production or Process"
               Top             =   1440
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
               TabIndex        =   107
               ToolTipText     =   "Select Production or Process"
               Top             =   1440
               Value           =   -1  'True
               Width           =   2085
            End
            Begin VB.CheckBox ChkBoxArr 
               Alignment       =   1  'Right Justify
               Caption         =   "Show All Prd Lines            "
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
               Index           =   3
               Left            =   60
               TabIndex        =   61
               Tag             =   "ShowPrdLine"
               ToolTipText     =   "Check To All Production Lines irrespective Of Rejected Qty or Wt"
               Top             =   6060
               Width           =   2820
            End
            Begin VB.CheckBox ChkBoxArr 
               Alignment       =   1  'Right Justify
               Caption         =   "Show Box Det                     "
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
               TabIndex        =   58
               Tag             =   "ShowBoxDet"
               ToolTipText     =   "Check To Show Box Details"
               Top             =   5490
               Width           =   2850
            End
            Begin VB.Frame FraOptQW 
               BorderStyle     =   0  'None
               Height          =   285
               Left            =   2685
               TabIndex        =   106
               ToolTipText     =   "Select either Quantity Or Weight"
               Top             =   5205
               Width           =   4785
               Begin VB.OptionButton OptQW 
                  Caption         =   "Wt"
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
                  Left            =   2670
                  TabIndex        =   56
                  Top             =   0
                  Width           =   1365
               End
               Begin VB.OptionButton OptQW 
                  Caption         =   "Qty"
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
                  TabIndex        =   55
                  Top             =   0
                  Width           =   1365
               End
            End
            Begin MwfCtl.MWCTL_MED atxt 
               Height          =   285
               Index           =   14
               Left            =   5355
               TabIndex        =   6
               ToolTipText     =   "Enter To Bag Year"
               Top             =   570
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
               Index           =   17
               Left            =   2685
               TabIndex        =   3
               ToolTipText     =   "Enter From Bag Year"
               Top             =   570
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
               Index           =   19
               Left            =   3105
               TabIndex        =   4
               ToolTipText     =   "Enter From Bag Character"
               Top             =   570
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
               Index           =   20
               Left            =   3885
               TabIndex        =   5
               ToolTipText     =   "Enter From Bag Number"
               Top             =   570
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
               Index           =   21
               Left            =   5775
               TabIndex        =   7
               ToolTipText     =   "Enter To Bag Character"
               Top             =   570
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
               Left            =   6555
               TabIndex        =   8
               ToolTipText     =   "Enter To Bag Number"
               Top             =   570
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
               Index           =   23
               Left            =   5355
               TabIndex        =   10
               ToolTipText     =   "Enter To Transaction Date"
               Top             =   855
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
               Index           =   25
               Left            =   2685
               TabIndex        =   9
               ToolTipText     =   "Enter From Transaction Date"
               Top             =   855
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
               Index           =   1
               Left            =   5355
               TabIndex        =   15
               ToolTipText     =   "Enter To FrBag Location"
               Top             =   1785
               Width           =   1485
               _ExtentX        =   2619
               _ExtentY        =   503
               MaxLength       =   8
               DataField       =   "UOISSLOCTO"
               IdName          =   "UOISSLOCTO"
            End
            Begin MwfCtl.MWCTL_MED atxt 
               Height          =   285
               Index           =   2
               Left            =   8010
               TabIndex        =   16
               ToolTipText     =   "Enter FrBag Location Selection"
               Top             =   1785
               Width           =   6870
               _ExtentX        =   12118
               _ExtentY        =   503
               MaxLength       =   255
               DataField       =   "UOISSLOCSEL"
               IdName          =   "UOISSLOCSEL"
            End
            Begin MwfCtl.MWCTL_MED atxt 
               Height          =   285
               Index           =   0
               Left            =   2685
               TabIndex        =   14
               ToolTipText     =   "Enter From FrBag Location"
               Top             =   1785
               Width           =   1485
               _ExtentX        =   2619
               _ExtentY        =   503
               MaxLength       =   8
               DataField       =   "UOISSLOCFR"
               IdName          =   "UOISSLOCFR"
            End
            Begin MwfCtl.MWCTL_MED atxt 
               Height          =   285
               Index           =   28
               Left            =   5355
               TabIndex        =   21
               ToolTipText     =   "Enter To Design Category"
               Top             =   2355
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
               Left            =   8010
               TabIndex        =   22
               ToolTipText     =   "Enter Design Category Selection"
               Top             =   2355
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
               Left            =   2685
               TabIndex        =   20
               ToolTipText     =   "Enter From Design Category"
               Top             =   2355
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
               Left            =   5355
               TabIndex        =   27
               ToolTipText     =   "Enter To Design Code"
               Top             =   2925
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
               Index           =   32
               Left            =   8010
               TabIndex        =   28
               ToolTipText     =   "Enter Design Code Selection"
               Top             =   2925
               Width           =   6870
               _ExtentX        =   12118
               _ExtentY        =   503
               MaxLength       =   255
               DataField       =   "UODMCDSEL"
               IdName          =   "UODMCDSEL"
               CmpStr          =   "BOdDmCd In"
            End
            Begin MwfCtl.MWCTL_MED atxt 
               Height          =   285
               Index           =   33
               Left            =   2685
               TabIndex        =   26
               ToolTipText     =   "Enter From Design Code"
               Top             =   2925
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
               Index           =   40
               Left            =   5355
               TabIndex        =   41
               ToolTipText     =   "Enter To Customer"
               Top             =   4065
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
               Index           =   41
               Left            =   8010
               TabIndex        =   42
               ToolTipText     =   "Enter Customer Selection"
               Top             =   4065
               Width           =   6870
               _ExtentX        =   12118
               _ExtentY        =   503
               MaxLength       =   255
               DataType        =   4
               DataField       =   "UOCMCDSEL"
               IdName          =   "UOCMCDSEL"
               CmpStr          =   "OmCmCd In"
            End
            Begin MwfCtl.MWCTL_MED atxt 
               Height          =   285
               Index           =   42
               Left            =   2685
               TabIndex        =   40
               ToolTipText     =   "Enter From Customer"
               Top             =   4065
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
               Index           =   5
               Left            =   6135
               TabIndex        =   36
               ToolTipText     =   "Enter To Order Year"
               Top             =   3495
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
               Index           =   3
               Left            =   3465
               TabIndex        =   33
               ToolTipText     =   "Enter From Order Year"
               Top             =   3495
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
               Index           =   10
               Left            =   2685
               TabIndex        =   32
               ToolTipText     =   "Enter From Order Tc"
               Top             =   3495
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
               Index           =   13
               Left            =   3885
               TabIndex        =   34
               ToolTipText     =   "Enter From Order Character"
               Top             =   3495
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
               Index           =   4
               Left            =   5355
               TabIndex        =   35
               ToolTipText     =   "Enter To Order Tc"
               Top             =   3495
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
               Index           =   6
               Left            =   6555
               TabIndex        =   37
               ToolTipText     =   "Enter To Order Character"
               Top             =   3495
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
               Index           =   53
               Left            =   2685
               TabIndex        =   38
               ToolTipText     =   "Enter From Order Number"
               Top             =   3780
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
               Index           =   54
               Left            =   5355
               TabIndex        =   39
               ToolTipText     =   "Enter To Order Number"
               Top             =   3780
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
               Index           =   7
               Left            =   5355
               TabIndex        =   18
               ToolTipText     =   "Enter To By Worker"
               Top             =   2070
               Width           =   1500
               _ExtentX        =   2646
               _ExtentY        =   503
               MaxLength       =   8
               DataField       =   "UOWRKTO"
               IdName          =   "UOWRKTO"
               CmpStr          =   "TdByWrk <="
            End
            Begin MwfCtl.MWCTL_MED atxt 
               Height          =   285
               Index           =   8
               Left            =   8010
               TabIndex        =   19
               ToolTipText     =   "Enter By Worker Selection"
               Top             =   2070
               Width           =   6870
               _ExtentX        =   12118
               _ExtentY        =   503
               MaxLength       =   255
               DataField       =   "UOWRKSEL"
               IdName          =   "UOWRKSEL"
               CmpStr          =   "TdByWrk In"
            End
            Begin MwfCtl.MWCTL_MED atxt 
               Height          =   285
               Index           =   9
               Left            =   2685
               TabIndex        =   17
               ToolTipText     =   "Enter From By Worker"
               Top             =   2070
               Width           =   1515
               _ExtentX        =   2672
               _ExtentY        =   503
               MaxLength       =   8
               DataField       =   "UOWRKFR"
               IdName          =   "UOWRKFR"
               CmpStr          =   "TdByWrk >="
            End
            Begin MwfCtl.MWCTL_MED atxt 
               Height          =   285
               Index           =   15
               Left            =   5355
               TabIndex        =   30
               ToolTipText     =   "Enter To Karat"
               Top             =   3210
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
               Index           =   18
               Left            =   2685
               TabIndex        =   29
               ToolTipText     =   "Enter From Karat"
               Top             =   3210
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
               Index           =   43
               Left            =   5355
               TabIndex        =   44
               ToolTipText     =   "Enter To Purchase Order "
               Top             =   4350
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
               Index           =   44
               Left            =   8010
               TabIndex        =   45
               ToolTipText     =   "Enter Purchase Order Selection"
               Top             =   4350
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
               Index           =   45
               Left            =   2685
               TabIndex        =   43
               ToolTipText     =   "Enter From Purchase Order "
               Top             =   4350
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
               Index           =   11
               Left            =   5355
               TabIndex        =   24
               ToolTipText     =   "Enter To Design Production Category"
               Top             =   2640
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
               Left            =   8010
               TabIndex        =   25
               ToolTipText     =   "Enter Design Production Category Selection"
               Top             =   2640
               Width           =   6870
               _ExtentX        =   12118
               _ExtentY        =   503
               MaxLength       =   255
               DataType        =   4
               DataField       =   "UODMPRDCTGSEL"
               IdName          =   "UODMPRDCTGSEL"
               CmpStr          =   "DmPrdCtg In"
            End
            Begin MwfCtl.MWCTL_MED atxt 
               Height          =   285
               Index           =   16
               Left            =   2685
               TabIndex        =   23
               ToolTipText     =   "Enter From Design Production Category"
               Top             =   2640
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
               Index           =   26
               Left            =   8010
               TabIndex        =   31
               ToolTipText     =   "Enter Karat Selection"
               Top             =   3210
               Width           =   6870
               _ExtentX        =   12118
               _ExtentY        =   503
               MaxLength       =   255
               DataField       =   "UOKTSEL"
               IdName          =   "UOKTSEL"
               CmpStr          =   "OdKt In"
            End
            Begin MwfCtl.MWCTL_MED atxt 
               Height          =   285
               Index           =   37
               Left            =   5355
               TabIndex        =   52
               ToolTipText     =   "Enter To Rejection Code"
               Top             =   4920
               Width           =   1335
               _ExtentX        =   2355
               _ExtentY        =   503
               MaxLength       =   8
               DataField       =   "UOREJCDTO"
               IdName          =   "UOREJCDTO"
               CmpStr          =   "TdRjCd<= "
            End
            Begin MwfCtl.MWCTL_MED atxt 
               Height          =   285
               Index           =   39
               Left            =   8010
               TabIndex        =   53
               ToolTipText     =   "Enter Rejection Code Selection"
               Top             =   4920
               Width           =   6870
               _ExtentX        =   12118
               _ExtentY        =   503
               MaxLength       =   255
               DataField       =   "UOREJCDSEL"
               IdName          =   "UOREJCDSEL"
               CmpStr          =   "TdRjCd In "
            End
            Begin MwfCtl.MWCTL_MED atxt 
               Height          =   285
               Index           =   27
               Left            =   2685
               TabIndex        =   51
               ToolTipText     =   "Enter From Rejection Code"
               Top             =   4920
               Width           =   1335
               _ExtentX        =   2355
               _ExtentY        =   503
               MaxLength       =   8
               DataField       =   "UOREJCDFR"
               IdName          =   "UOREJCDFR"
               CmpStr          =   "TdRjCd>= "
            End
            Begin MwfCtl.MWCTL_MED atxt 
               Height          =   285
               Index           =   34
               Left            =   3255
               TabIndex        =   57
               ToolTipText     =   "Show Box Detail (Yes / No)"
               Top             =   5490
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
               Index           =   52
               Left            =   2685
               TabIndex        =   0
               ToolTipText     =   "Enter Company Code"
               Top             =   285
               Width           =   705
               _ExtentX        =   1244
               _ExtentY        =   503
               MaxLength       =   3
               DataType        =   4
               DataField       =   "UOCOCDFR"
               IdName          =   "UOCOCDFR"
               CmpStr          =   "TdCoCd >= "
            End
            Begin MwfCtl.MWCTL_MED atxt 
               Height          =   285
               Index           =   24
               Left            =   2685
               TabIndex        =   54
               ToolTipText     =   "Enter 'Y' To Show Quantity Or 'N' To Show Weight"
               Top             =   5205
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
               Index           =   35
               Left            =   2685
               TabIndex        =   59
               ToolTipText     =   "Specify Bag Type or ' ' for All"
               Top             =   5775
               Width           =   285
               _ExtentX        =   503
               _ExtentY        =   503
               DataType        =   4
               DataField       =   "UOBAGTYP"
               IdName          =   "UOBAGTYP"
            End
            Begin MwfCtl.MWCTL_MED atxt 
               Height          =   285
               Index           =   36
               Left            =   3255
               TabIndex        =   60
               ToolTipText     =   "Show All Production Lines? (Enter 'Y' for All Or 'N' For Only Production Lines With Rejection Qty Or Wt)"
               Top             =   6060
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
               Left            =   5355
               TabIndex        =   1
               ToolTipText     =   "Enter To Company Code"
               Top             =   285
               Width           =   705
               _ExtentX        =   1244
               _ExtentY        =   503
               MaxLength       =   3
               DataType        =   4
               DataField       =   "UOCOCDTO"
               IdName          =   "UOCOCDTO"
               CmpStr          =   "TdCoCd <= "
            End
            Begin MwfCtl.MWCTL_MED atxt 
               Height          =   285
               Index           =   75
               Left            =   8010
               TabIndex        =   2
               ToolTipText     =   "Enter Design Code Selection"
               Top             =   285
               Width           =   6870
               _ExtentX        =   12118
               _ExtentY        =   503
               MaxLength       =   255
               DataField       =   "UOCOCDSEL"
               IdName          =   "UOCOCDSEL"
               CmpStr          =   "TdCoCd In"
            End
            Begin MwfCtl.MWCTL_MED atxt 
               Height          =   285
               Index           =   55
               Left            =   5355
               TabIndex        =   12
               ToolTipText     =   "Enter To Parent Code"
               Top             =   1140
               Width           =   1485
               _ExtentX        =   2619
               _ExtentY        =   503
               MaxLength       =   8
               DataField       =   "UOPRNTCDTO"
               IdName          =   "UOPRNTCDTO"
               CmpStr          =   "Lc.LocPrntCd <="
            End
            Begin MwfCtl.MWCTL_MED atxt 
               Height          =   285
               Index           =   56
               Left            =   8010
               TabIndex        =   13
               ToolTipText     =   "Enter Parent Code Selection"
               Top             =   1140
               Width           =   6870
               _ExtentX        =   12118
               _ExtentY        =   503
               MaxLength       =   255
               DataField       =   "UOPRNTCDSEL"
               IdName          =   "UOPRNTCDSEL"
               CmpStr          =   "Lc.LocPrntCd In"
            End
            Begin MwfCtl.MWCTL_MED atxt 
               Height          =   285
               Index           =   57
               Left            =   2685
               TabIndex        =   11
               ToolTipText     =   "Enter From Parent Code"
               Top             =   1140
               Width           =   1485
               _ExtentX        =   2619
               _ExtentY        =   503
               MaxLength       =   8
               DataField       =   "UOPRNTCDFR"
               IdName          =   "UOPRNTCDFR"
               CmpStr          =   "Lc.LocPrntCd >="
            End
            Begin MwfCtl.MWCTL_MED atxt 
               Height          =   285
               Index           =   69
               Left            =   3975
               TabIndex        =   47
               ToolTipText     =   "Enter From Sub Purchase Order"
               Top             =   4635
               Width           =   1335
               _ExtentX        =   2355
               _ExtentY        =   503
               MaxLength       =   20
               DataField       =   "UOODPONOFR"
               IdName          =   "UOODPONOFR"
               CmpStr          =   "OdPoNo >="
            End
            Begin MwfCtl.MWCTL_MED atxt 
               Height          =   285
               Index           =   70
               Left            =   6645
               TabIndex        =   49
               ToolTipText     =   "Enter To Sub Purchase Order "
               Top             =   4635
               Width           =   1335
               _ExtentX        =   2355
               _ExtentY        =   503
               MaxLength       =   20
               DataField       =   "UOODPONOTO"
               IdName          =   "UOODPONOTO"
               CmpStr          =   "OdPoNo <="
            End
            Begin MwfCtl.MWCTL_MED atxt 
               Height          =   285
               Index           =   71
               Left            =   8010
               TabIndex        =   50
               ToolTipText     =   "Enter Sub Purchase Order Selection"
               Top             =   4635
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
               Index           =   90
               Left            =   8040
               TabIndex        =   109
               ToolTipText     =   "Show Production or Process Location"
               Top             =   1440
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
               Index           =   38
               Left            =   2685
               TabIndex        =   46
               ToolTipText     =   "Enter From Sub Customer"
               Top             =   4635
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
               Index           =   46
               Left            =   5355
               TabIndex        =   48
               ToolTipText     =   "Enter To Sub Customer"
               Top             =   4635
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
               Index           =   23
               Left            =   6585
               TabIndex        =   111
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
               Index           =   16
               Left            =   3915
               TabIndex        =   110
               Top             =   4680
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
               Left            =   60
               TabIndex        =   105
               Top             =   4635
               Width           =   1425
            End
            Begin VB.Label LblPrntCd 
               BackStyle       =   0  'Transparent
               Caption         =   "Parent Code"
               BeginProperty Font 
                  Name            =   "Arial"
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
               Top             =   1140
               Width           =   1425
            End
            Begin VB.Label LblBagTyp 
               BackStyle       =   0  'Transparent
               Caption         =   "Bags Type"
               BeginProperty Font 
                  Name            =   "Arial"
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
               TabIndex        =   103
               ToolTipText     =   "Location"
               Top             =   5775
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
               Left            =   60
               TabIndex        =   102
               Top             =   6060
               Visible         =   0   'False
               Width           =   1965
            End
            Begin VB.Label ALBL 
               BackStyle       =   0  'Transparent
               Caption         =   "Show Qty/Wt"
               BeginProperty Font 
                  Name            =   "Arial"
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
               TabIndex        =   101
               ToolTipText     =   "Location"
               Top             =   5205
               Width           =   1485
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
               TabIndex        =   100
               Top             =   285
               Width           =   1335
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
               Index           =   9
               Left            =   60
               TabIndex        =   99
               ToolTipText     =   "Location"
               Top             =   5490
               Visible         =   0   'False
               Width           =   1485
            End
            Begin VB.Label LblRejCd 
               BackStyle       =   0  'Transparent
               Caption         =   "Rejection Cd"
               BeginProperty Font 
                  Name            =   "Arial"
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
               Top             =   4920
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
               Left            =   60
               TabIndex        =   97
               Top             =   2640
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
               Index           =   8
               Left            =   60
               TabIndex        =   96
               Top             =   1785
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
               TabIndex        =   95
               Top             =   2355
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
               TabIndex        =   94
               Top             =   2925
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
               TabIndex        =   93
               Top             =   4065
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
               Index           =   27
               Left            =   3405
               TabIndex        =   92
               Top             =   3495
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
               Left            =   6075
               TabIndex        =   91
               Top             =   3495
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
               Left            =   3825
               TabIndex        =   90
               Top             =   3495
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
               Left            =   6495
               TabIndex        =   89
               Top             =   3495
               Width           =   105
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
               Left            =   60
               TabIndex        =   88
               Top             =   3780
               Width           =   1605
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
               Left            =   60
               TabIndex        =   87
               Top             =   3495
               Width           =   1725
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
               Left            =   60
               TabIndex        =   86
               Top             =   2070
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
               TabIndex        =   85
               Top             =   3210
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
               Index           =   0
               Left            =   60
               TabIndex        =   84
               Top             =   4350
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
               Index           =   11
               Left            =   60
               TabIndex        =   83
               Top             =   855
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
               Index           =   10
               Left            =   60
               TabIndex        =   82
               Top             =   570
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
               Index           =   18
               Left            =   3045
               TabIndex        =   81
               Top             =   570
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
               Left            =   5715
               TabIndex        =   80
               Top             =   570
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
               Left            =   3825
               TabIndex        =   79
               Top             =   570
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
               Left            =   6495
               TabIndex        =   78
               Top             =   570
               Width           =   105
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
               Left            =   2805
               TabIndex        =   77
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
               Left            =   5325
               TabIndex        =   76
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
               TabIndex        =   75
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
            TabIndex        =   72
            Top             =   360
            Width           =   15105
            Begin MwfCtl.MWCTL_GRP GRP_REP 
               Height          =   360
               Left            =   240
               TabIndex        =   73
               Top             =   360
               Width           =   8445
               _ExtentX        =   14896
               _ExtentY        =   635
            End
         End
      End
   End
End
Attribute VB_Name = "EmrFrpRejAna"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim Rep As New EmrRepRejAna
Dim moCn As MwfLib.MDOConnection   '*** (Bef speed) bef 24/11/06

'****** Sachin 2.14.0 - [25-11-2006] - [Multi-Company] ******
Dim ws_MultiCoMod As Boolean
'****** Sachin 2.14.0 - [25-11-2006] - [Multi-Company] ******

Private Sub SetGroupSort()
  GRP_REP.Add "Location", "TdCoCd+TdByLoc", "'['+TdCoCd+'] '+TdByLoc", "TdByLoc", "hTdByLoc", "Company Code,Loc Prnt Code", "", "", "(Select LocDesc From Loc Where LocCoCd = TdCoCd and LocTyp= 'P' and LocCd= TdByLoc) "
  GRP_REP.Add "(Location)Worker", "TdCoCd+TdByLoc+TdByWrk", "'('+TdCoCd+'/'+ RTrim(TdByLoc)+ ') '+ TdByWrk", "TdByLoc,TdByWrk", "hTdByLoc,hTdByWrk", "Company Code,Process", "", "", "(Select vPDesc From vParam Where vPCoCd = TdCoCd and vPTyp= 'WORK' and vPMCd= TdByWrk) "
  GRP_REP.Add "Process", "TdCoCd+LocPrcs", "'['+TdCoCd+'] '+Lc.LocPrcs", "TdByLoc", "hTdByLoc", "Company Code,Loc Prnt Code", "", "", "(Select LocDesc From Loc Where LocCoCd = TdCoCd and LocTyp= 'R' and LocCd= Lc.LocPrcs) "
  GRP_REP.Add "(Process)Worker", "TdCoCd+LocPrcs+TdByWrk", "'('+TdCoCd+'/'+ RTrim(Lc.LocPrcs)+ ') '+ TdByWrk", "TdByLoc,TdByWrk", "hTdByLoc,hTdByWrk", "Company Code,Process", "", "", "(Select vPDesc From vParam Where vPCoCd = TdCoCd and vPTyp= 'WORK' and vPMCd= TdByWrk) "
  
  GRP_REP.Add "Design Ctg", "DmCtg", "", "DmCtg", "hDmCtg", "", "", "", "(Select PDesc from Param where PTyp= 'DMCTG' and PMCd= DmCtg) "
  GRP_REP.Add "Prd Ctg", "DmPrdCtg", "", "DmPrdCtg", "hDmPrdCtg", "", "", "", "(Select PDesc from Param where PTyp= 'PRDCTG' and PMCd= DmPrdCtg) "
  'indigo.164 DmTcTyp= 'DM' changed as DmTcTyp=BOdDmTcTyp in DsgMst where condition
  GRP_REP.Add "Design Code", "BOdDmCd", "", "BOdDmCd", "hBOdDmCd", "Design Ctg,Prd Ctg,(DsgCd)Colour", "", "", "(Select Substring(DmDesc,1,120) From DsgMst Where " + IIF(gs_Partition = ctCurrPrtn, " DmPrtKey='" + ctCurrPrtn + "' and ", "") + " DmTcTyp=BOdDmTcTyp and DmCd= BOdDmCd and DmSz= '') "
  GRP_REP.Add "(DsgCd)Colour", "OdDmCd+OdDmCol", "'('+OdDmCd+') '+OdDmCol", "BOdDmCd", "hBOdDmCd", "Design Ctg,Prd Ctg,Design Code", "", "", "(Select PDesc From Param Where PTyp= 'DMCOL' and PMCd = OdDmCol) "
  GRP_REP.Add "Ord Kt", "OdKt", "", "OdKt", "hOdKt", "", "", "", "(Select PDesc from Param where PTyp= 'KT' and PMCd= OdKt) "
  GRP_REP.Add "OrderNo", "TdCoCd+BOdTc+BOdYy+BOdChr+str(BOdNo)", "TdCoCd+'/'+BOdTc+'/'+BOdYy+'/'+BOdChr+'/'+LTrim(Str(BOdNo))", "wOrdNo,OmCmCd,OmPoNo", "hOrdNo,hOmCmCd,hOmPoNo", "Company Code,Customer Code,Customer PO", "", "", "'('+OmCmCd+ '/'+ OmPONo+ ')' "
  '******Geeta*****************Emr206****05/04/04*******
  '**********If UserLevel>6 then don't display the Customer Name
  GRP_REP.Add "Customer Code", "OmCmCd", "", "OmCmCd", "hOmCmCd", "", "", "", "(Select (case when (Select PNum from Param where Ptyp = 'USR' and PmCd = '" + Trim(EmrFrmMDI.TXT_UCD.text) + "')  <= 5 then  CmName else CmCd end) From CustMst Where CmCtg= 'C' and CmCd= OmCmCd) "
  GRP_REP.Add "Customer PO", "OmCmCd+OmPONo", "'('+ OmCmCd+ ') '+ LTrim(OmPONo)", "OmCmCd,OmPoNo", "hOmCmCd,hOmPONo", "Company Code", "", "", "BOdTc+'/'+BOdYy+'/'+BOdChr+'/'+LTrim(Str(BOdNo))"
  '*** Jay 3.2.0 [OdPoNo]
  'vk.19 sub customer added
  GRP_REP.Add "Cust SUB PO", "OmCmCd+OmPONo+OdSubCust+OdPoNo", "'('+ OmCmCd+ ') '+ LTrim(OmPONo)", "OmCmCd,OmPoNo,OdPoNo", "hOmCmCd,hOmPONo,hOdPONo", "Company Code,Customer PO", "", "", "BOdTc+'/'+BOdYy+'/'+BOdChr+'/'+LTrim(Str(BOdNo))+'/'+LTrim(Str(BOdSr))+(case when OdPONo<> '' or OdSubCust<>'' then ' - '+ LTrim(OdSubCust)+' : '+LTrim(OdPONo) else '' end )"
      
  '*** Jay 3.2.0 [OdPoNo]
  '### Check This Grouping
  GRP_REP.Add "Rej Code", "TdRjCd", "", "", "", "", "", "", "(Select PDesc From Param Where PTyp= 'REJCD' and PMCd= TdRjCd) "
  '### Check This Grouping
'*** Jay 2.13(CT) ***
  GRP_REP.Add "Bag Type", "(Select (case PValue when 'F' then 'Flute Bag' when 'P' then 'Primary Bag' when 'C' then 'Component Bag' end) from Param where PTyp= 'BCHR' and PMCd= BChr and PSCd= '')", _
    "(Select case PValue when 'F' then 'Flute Bag' when 'P' then 'Primary Bag' when 'C' then 'Component Bag' end from Param where PTyp= 'BCHR' and PMCd= BChr and PSCd= '')", "", "", "Bag Typ+Chr", "", "", "(Select case PValue when 'F' then 'Flute Bag' when 'P' then 'Primary Bag' when 'C' then 'Component Bag' end from Param where PTyp= 'BCHR' and PMCd= BChr and PSCd= '')"
  GRP_REP.Add "Bag Character", "BChr", "BChr", "", "", "Bag Type,Bag Typ+Chr", "", "", "(Select PDesc from Param where PTyp ='BCHR' and PMCd= BChr and PSCd= '')"
  GRP_REP.Add "Bag Typ+Chr", "(Select (case PValue when 'F' then 'Flute Bag' when 'P' then 'Primary Bag' when 'C' then 'Component Bag' end) from Param where PTyp= 'BCHR' and PMCd= BChr and PSCd= '')+BChr", _
    "(Select (case PValue when 'F' then 'Flute Bag' when 'P' then 'Primary Bag' when 'C' then 'Component Bag' end) from Param where PTyp= 'BCHR' and PMCd= BChr and PSCd= '')+'/ '+ BChr", "", "", "Bag Type,Bag Character", "", "", "(Select PDesc from Param where PTyp ='BCHR' and PMCd= BChr and PSCd= '')"

  '****** Sachin 2.14.0 - [23-11-2006] - [Multi-Company] - ,Process,OrderNo,OrderSr,Customer PO,FgSubLoc ******
  If ws_MultiCoMod = True Then
    GRP_REP.Add "Company Code", "TdCoCd", "TdCoCd", "", "", "", "", "", "(Select HName from Head where HCd= '" + ctSelfCmCd + "' and HCoCd=TdCoCd) "
    GRP_REP.Add "Loc Prnt Code", "IsNull(Lc.LocPrntCd, '')", "IsNull(Lc.LocPrntCd, '')", "", "", "", "", "", "IsNull((Select LocDesc from Loc Ln where Ln.LocCoCd= '" + ctSelfCoCd + "' and Ln.LocCd= IsNull(Lc.LocPrntCd, '')), '') "
  End If
  '****** Sachin 2.14.0 - [23-11-2006] - [Multi-Company] - ******

'' **** Zubin **** '
'  GRP_REP.Add "Bag Type", "(case PValue when 'Y' then 'Flute Bag' when 'N' then 'Primary Bag' end)", "(case PValue when 'Y' then 'Flute Bag' when 'N' then 'Primary Bag' end )", "", "", "Bag Typ+Chr", "", "", "(case PValue when 'Y' then 'Flute Bag' when 'N' then 'Primary Bag' end)"
'  GRP_REP.Add "Bag Character", "BChr", "BChr", "", "", "Bag Type,Bag Typ+Chr", "", "", "(Select PDesc from param where PTyp ='BCHR' and PMCd= BChr and PSCd= '')"
'  GRP_REP.Add "Bag Typ+Chr", "(case PValue when 'Y' then 'Flute Bag' when 'N' then 'Primary Bag' end)+BChr", "(case PValue when 'Y' then 'Flute Bag' when 'N' then 'Primary Bag' end)+ '/ ' + BChr", "", "", "Bag Type,Bag Character", "", "", "(Select PDesc from param where PTyp ='BCHR' and PMCd= BChr and PSCd= '')"
'' **** Zubin **** '
'*** Jay 2.13(CT) ***
End Sub
Private Sub adc_setreprecsource()
  '*** Report Sql ***
'  Select space(30) as Grp1, space(30) as Grp2, space(30) as Grp3,
'  space(30) as DGrp1, space(30) as DGrp2, space(30) as DGrp3,
'  space(30) as G1Desc, space(30) as G2Desc, space(30) as G3Desc,
'  space(1) as qTag, TdCoCd, TdTc, TdYy, TdChr, TdNo, TdSr, TdBYy, TdBChr, TdBNo,
'  TdRjQty as qRjQty, TdRjPtQty as qRjPtQty, TdBQty as qPrdQty, TdPtQty as qPrdPtQty,
'  0 as qMul, TdRjCd as qRjCd, TdByWrk, TdByLoc as LocPrntCd, TdByLoc, TDt, DmCtg, DmPrdCtg, BOdDmCd, OdKt,
'  BOdTc , BOdYy, BOdChr, BOdNo, OmCmCd, OmPoNo
'  From Txnd, Txn, Bag, OrdDsg, OrdMst, DsgMst Where 1 = 2
  '*** Report Sql ***

  Dim grpflds As String, wrepcnd As String, wCnd As String, wSqlStrg1 As String, wSqlStrg2 As String, wSqlStrg As String
' **** Zubin 211 **** '
  Dim ws_BagJoin As String
' **** Zubin 211 **** '
  
  Dim ws_TmpTblA As String, ws_ShowAllPrd As String   '*** (Jen 2.13)
  
  Dim wAddlCnd As String  '****** Sachin 3.02.0 - New fields In Where Clause
  
  Set Rep = Nothing
  Set adc.RepSource = Rep
  Call DispCoNm
  Call SetMwName(Rep)   ' ***** Manali 3.6.1 - Microway footer
  Call GRP_REP.Gen3LRep
  grpflds = GRP_REP.GrpFldLst
    
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
  
  '*** If Detail not selected then suppress Box Header, Detail and Footer ***
  If wDetPos = 0 Then
    Rep.BoxHd.Suppress = True: Rep.BoxDet.Suppress = True: Rep.BoxFt.Suppress = True
    Rep.DetSec2.Suppress = True
  End If
    
  '****** Sachin 3.02.0 - Additional Conditions For New Fields
  wAddlCnd = wAddlCnd + IIF(wAddlCnd <> "" And Trim(adc("UOCMCDFR")) <> "", " And ", "") + IIF(Trim(adc("UOCMCDFR")) <> "", " OdOmCmCd >= '" + Trim(adc("UOCMCDFR")) + "'", "")
  wAddlCnd = wAddlCnd + IIF(wAddlCnd <> "" And Trim(adc("UOCMCDTO")) <> "", " And ", "") + IIF(Trim(adc("UOCMCDTO")) <> "", " OdOmCmCd <= '" + Trim(adc("UOCMCDTO")) + "'", "")
  wAddlCnd = wAddlCnd + IIF(wAddlCnd <> "" And Trim(adc("UOCMCDSEL")) <> "", " And ", "") + IIF(Trim(adc("UOCMCDSEL")) <> "", " OdOmCmCd In (" + Trim(adc("UOCMCDSEL")) + ")", "")
  
  wAddlCnd = wAddlCnd + IIF(wAddlCnd <> "" And Trim(adc("UODTFR")) <> "01/01/80", " And ", "") + IIF(Trim(adc("UODTFR")) <> "01/01/80", " TdDt >= '" + Trim(adc("UODTFR")) + "'", "")
  wAddlCnd = wAddlCnd + IIF(wAddlCnd <> "" And Trim(adc("UODTTO")) <> "01/01/80", " And ", "") + IIF(Trim(adc("UODTTO")) <> "01/01/80", " TdDt <= '" + Trim(adc("UODTTO")) + "'", "")
  '****** Sachin 3.02.0 - Additional Conditions For New Fields
    
  adc("UoRejCdFr").CmpStr = "TdRjCd>= ": adc("UoRejCdTo").CmpStr = "TdRjCd<= ": adc("UoRejCdSel").CmpStr = "TdRjCd In "
  wrepcnd = adc.RepCond
  wCnd = IIF(wrepcnd <> "", " and " + wrepcnd, "")


  If OptLoc(0).Value = True Then
      If adc("UoIssLocFr") <> "" Then wCnd = wCnd + " and TdByLoc >= '" + adc("UoIssLocFr") + "'"
      If adc("UoIssLocTo") <> "" Then wCnd = wCnd + " and TdByLoc <= '" + adc("UoIssLocTo") + "'"
      If adc("UoIssLocSel") <> "" Then wCnd = wCnd + " and TdByLoc In (" + adc("UoIssLocSel") + ")"
  ElseIf OptLoc(1).Value = True Then
      If adc("UoIssLocFr") <> "" Then wCnd = wCnd + " and Lc.LocPrcs >= '" + adc("UoIssLocFr") + "'"
      If adc("UoIssLocTo") <> "" Then wCnd = wCnd + " and Lc.LocPrcs <= '" + adc("UoIssLocTo") + "'"
      If adc("UoIssLocSel") <> "" Then wCnd = wCnd + " and Lc.LocPrcs In (" + adc("UoIssLocSel") + ")"
  End If
  
  
'*** Jay 2.13(CT) ***
  ws_BagJoin = " Join Param On PTyp= 'BCHR' and PMCd= BChr and PSCd= '' " + _
                  IIF(adc("UoBagTyp") <> "", " and PValue= '" + adc("UoBagTyp") + "' ", " ")
'' **** Zubin 211 **** '
'  If UCase(adc("UOYN2")) = "N" Then
'    ws_BagJoin = " Join Param On PTyp= 'BCHR' and PMCd= BChr and PSCd= '' and PValue = 'N' "
'  ElseIf UCase(adc("UOYN2")) = "Y" Then
'    ws_BagJoin = " Join Param On PTyp= 'BCHR' and PMCd= BChr and PSCd= '' and PValue = 'Y' "
'  Else
'    ws_BagJoin = " Join Param On PTyp= 'BCHR' and PMCd= BChr and PSCd= '' "
'  End If
'*** Jay 2.13(CT) ***
  
  Dim ws_RjQtyWt As String, ws_RjPtQty As String, ws_PrdQty As String, ws_PrdPtQty As String
  
  ws_PrdQty = " 0.0 ": ws_PrdPtQty = " 0 "
  If UCase(adc("UOYN1")) = "Y" Then
    ws_RjQtyWt = " TdRjQty ": ws_RjPtQty = " TdRjPtQty "
  Else
    ws_RjQtyWt = " TdBGrWt ": ws_RjPtQty = " 0 "
  End If
  ' **** Zubin 211 **** '

  '*** (Bef 2.13)
  'wSqlStrg1 = "Select " + grpflds + ", " + _
  '           "'D' as qTag, TdCoCd, TdTc, TdYy, TdChr, TdNo, TdSr, TdBYy, TdBChr, TdBNo, " + _
  '           ws_RjQtyWt + " as qRjQty, " + ws_RjPtQty + " as qRjPtQty, " + ws_PrdQty + " as qPrdQty, " + _
  '           ws_PrdPtQty + " as qPrdPtQty, " + _
  '           "(Case when TdRjDc = 'D' then 1 when TdRjDc = 'C' then -1 else 0 end) as qMul, " + _
  '           "TdRjCd as qRjCd, TdByWrk, TdByLoc, TDt, DmCtg, DmPrdCtg, BOdDmCd, OdKt, " + _
  '           "BOdTc, BOdYy, BOdChr, BOdNo, OmCmCd, OmPoNo " + _
  '           "From Txnd Join Txn on TCoCd= TdCoCd and TTc= TdTc and TYy= TdYy " + _
  '           "and TChr= TdChr and TNo= TdNo " + _
  '           "Join Bag on BCoCd= TdCoCd and BYy= TdBYy and BChr= TdBChr and BNo= TdBNo " + _
  '           ws_BagJoin + _
  '           "Join OrdDsg on OdCoCd= BCoCd and OdTc= BOdTc and OdYy= BOdYy " + _
  '           "and OdChr= BOdChr and OdNo= BOdNo and OdSr= BOdSr " + _
  '           "Join OrdMst on OmCoCd= BCoCd and OmTc= BOdTc and OmYy= BOdYy " + _
  '           "and OmChr= BOdChr and OmNo= BOdNo " + _
  '           "Join DsgMst on DmTcTyp= 'DM' and DmCd= BOdDmCd and DmSz= '' " + _
  '           "where (Select PValue from Param where PTyp= 'Tc' and PMcd= TdTc and PScd= '')= 'REJ' " + _
  '           "and TdSrNo= 0 " + wCnd
  '*** (Bef 2.13)

  
  '****** (Jen 3.2.0) added: "left outer Join Loc Lc on LocCoCd= TdCoCd and LocCd= TdByLoc "
  '*** (Jen 2.13) (introduced #TMPA)
  '****** Sachin 3.02 - Id fields in Joins
  'indigo.164 DmTcTyp= 'DM' changed as DmTcTyp=BOdDmTcTyp in DsgMst join
  '6.1
  ws_TmpTblA = "Set dateformat " + DtFmtStr() + " Select " + grpflds + ", " + _
             "'D' as qTag, TdCoCd, TdTc, TdYy, TdChr, TdNo, TdSr, TdBYy, TdBChr, TdBNo, " + _
             ws_RjQtyWt + " as qRjQty, " + ws_RjPtQty + " as qRjPtQty, " + ws_PrdQty + " as qPrdQty, " + _
             ws_PrdPtQty + " as qPrdPtQty, " + _
             "(Case when TdRjDc = 'D' then 1 when TdRjDc = 'C' then -1 else 0 end) as qMul, " + _
             "TdRjCd as qRjCd, TdByWrk, LocPrntCd," + IIF(OptLoc(0).Value = True, "TdByLoc", "Lc.LocPrcs") + " as TdByLoc , TDt, DmCtg, DmPrdCtg, BOdDmCd, OdKt, " + _
             "BOdTc, BOdYy, BOdChr, BOdNo, OmCmCd, OmPoNo " + _
             " into #TMPA From Txnd Join Txn on TIdNo=TdTIdNo and TCoCd= TdCoCd and TTc= TdTc and TYy= TdYy " + _
             "and TChr= TdChr and TNo= TdNo " + IIF(gs_Partition = ctCurrPrtn, " and TPrtKey=TdPrtKey ", "") + _
             "Join Bag on BIdNo=TdBIdNo and BCoCd= TdCoCd and BYy= TdBYy and BChr= TdBChr and BNo= TdBNo " + IIF(gs_Partition = ctCurrPrtn, " and BPrtKey=TdPrtKey ", "") + _
             ws_BagJoin + _
             "Join OrdDsg on OdIdNo=BOdIdNo and OdCoCd= BCoCd and OdTc= BOdTc and OdYy= BOdYy " + _
             "and OdChr= BOdChr and OdNo= BOdNo and OdSr= BOdSr " + IIF(gs_Partition = ctCurrPrtn, " and OdPrtKey=BPrtKey ", "") + _
             "Join OrdMst on OmCoCd= BCoCd and OmTc= BOdTc and OmYy= BOdYy " + _
             "and OmChr= BOdChr and OmNo= BOdNo " + IIF(gs_Partition = ctCurrPrtn, " and OmPrtKey=BPrtKey ", "") + _
             "Join DsgMst on DmIdNo=BDmIdNo and DmTcTyp=BOdDmTcTyp and DmCd= BOdDmCd and DmSz= '' " + IIF(gs_Partition = ctCurrPrtn, " and DmPrtKey=BPrtKey ", "") + _
             "left outer Join Loc Lc on LocCoCd= TdCoCd and LocCd= TdByLoc " + _
             "where (Select PValue from Param where PTyp= 'Tc' and PMcd= TdTc and PScd= '')= 'REJ' " + _
             "and TdSrNo= 0 " + IIF(gs_Partition = ctCurrPrtn, " and TdPrtKey='" + ctCurrPrtn + "' ", "") + wCnd + IIF(wAddlCnd <> "", " And " + wAddlCnd, "")
  
             '*** (Bef 3.2.0)
             '"Join Loc Lc on LocCoCd= TdCoCd and LocCd=TdByLoc "
             '*** (Bef 3.2.0)
  
  moCn.CreateTmpTable "#TMPA", ws_TmpTblA
  wSqlStrg1 = "Set dateformat " + DtFmtStr() + " Select * from #TMPA "  '6.1
  '*** (Jen 2.13)

  adc("UoRejCdFr").CmpStr = "": adc("UoRejCdTo").CmpStr = "": adc("UoRejCdSel").CmpStr = ""
  wrepcnd = adc.RepCond
  wCnd = IIF(wrepcnd <> "", " and " + wrepcnd, "")
  If OptLoc(0).Value = True Then
      If adc("UoIssLocFr") <> "" Then wCnd = wCnd + " and TdByLoc >= '" + adc("UoIssLocFr") + "'"
      If adc("UoIssLocTo") <> "" Then wCnd = wCnd + " and TdByLoc <= '" + adc("UoIssLocTo") + "'"
      If adc("UoIssLocSel") <> "" Then wCnd = wCnd + " and TdByLoc In (" + adc("UoIssLocSel") + ")"
  ElseIf OptLoc(1).Value = True Then
      If adc("UoIssLocFr") <> "" Then wCnd = wCnd + " and Lc.LocPrcs >= '" + adc("UoIssLocFr") + "'"
      If adc("UoIssLocTo") <> "" Then wCnd = wCnd + " and Lc.LocPrcs <= '" + adc("UoIssLocTo") + "'"
      If adc("UoIssLocSel") <> "" Then wCnd = wCnd + " and Lc.LocPrcs In (" + adc("UoIssLocSel") + ")"
  End If
  
  ' **** Zubin 211 **** '
  ws_RjQtyWt = " 0.0 ": ws_RjPtQty = " 0 "
  If UCase(adc("UOYN1")) = "Y" Then
    ws_PrdQty = " TdBQty ": ws_PrdPtQty = " TdPtQty "
  Else
    ws_PrdQty = " TdBGrWt ": ws_PrdPtQty = " 0 "
  End If
  ' **** Zubin 211 **** '
  
  '*** (Bef 2.13)
  'wSqlStrg2 = "Select " + grpflds + ", " + _
  '           "'P' as qTag, TdCoCd, TdTc, TdYy, TdChr, TdNo, TdSr, TdBYy, TdBChr, TdBNo, " + _
  '           ws_RjQtyWt + " as qRjQty, " + ws_RjPtQty + " as qRjPtQty, " + ws_PrdQty + " as qPrdQty, " + _
  '           ws_PrdPtQty + " as qPrdPtQty, " + _
  '           "0 as qMul, '' as qRjCd, TdByWrk, TdByLoc, TDt, DmCtg, DmPrdCtg, BOdDmCd, OdKt, " + _
  '           "BOdTc, BOdYy, BOdChr, BOdNo, OmCmCd, OmPoNo " + _
  '           "From Txnd Join Txn on TCoCd= TdCoCd and TTc= TdTc and TYy= TdYy " + _
  '           "and TChr= TdChr and TNo= TdNo " + _
  '           "Join Bag on BCoCd= TdCoCd and BYy= TdBYy and BChr= TdBChr and BNo= TdBNo " + _
  '           ws_BagJoin + _
  '           "Join OrdDsg on OdCoCd= BCoCd and OdTc= BOdTc and OdYy= BOdYy " + _
  '           "and OdChr= BOdChr and OdNo= BOdNo and OdSr= BOdSr " + _
  '           "Join OrdMst on OmCoCd= BCoCd and OmTc= BOdTc and OmYy= BOdYy " + _
  '           "and OmChr= BOdChr and OmNo= BOdNo " + _
  '           "Join DsgMst on DmTcTyp= 'DM' and DmCd= BOdDmCd and DmSz= '' " + _
  '           "where TdPrdYN= 'Y' and TdSrNo=0 " + wCnd
  '*** (Bef 2.13)
  
  '*** (Jen 2.13)
  If adc("UoYN3") = "N" Then
    ws_ShowAllPrd = " where exists (Select 'x' from #TMPA A where A.DGrp1= Tbl.DGrp1 and A.DGrp2= Tbl.DGrp2 and A.DGrp3= Tbl.DGrp3 " + _
                    " and A.TdBYy= Tbl.TdBYy and A.TdBChr= Tbl.TdBChr and A.TdBNo= Tbl.TdBNo) "
  Else
    ws_ShowAllPrd = " "
  End If
  
  '****** (Jen 3.2.0) added: "left outer Join Loc Lc on LocCoCd= TdCoCd and LocCd= TdByLoc "
  '****** Sachin 3.02 - Id fields in Join
  'indigo.164 DmTcTyp= 'DM' changed as DmTcTyp=BOdDmTcTyp in DsgMst join
  wSqlStrg2 = " Select Grp1, Grp2, Grp3, DGrp1, DGrp2, DGrp3, G1Desc, G2Desc, G3Desc, " + _
              "'P' as qTag, TdCoCd, TdTc, TdYy, TdChr, TdNo, TdSr, TdBYy, TdBChr, TdBNo, " + _
              "qRjQty, qRjPtQty, qPrdQty, qPrdPtQty, 0 as qMul, '' as qRjCd, TdByWrk, LocPrntCd, TdByLoc, TDt, DmCtg, " + _
              "DmPrdCtg, BOdDmCd, OdKt, BOdTc, BOdYy, BOdChr, BOdNo, OmCmCd, OmPoNo " + _
              " From (Select " + grpflds + ", " + _
              "TdCoCd, TdTc, TdYy, TdChr, TdNo, TdSr, TdBYy, TdBChr, TdBNo, " + _
              ws_RjQtyWt + " as qRjQty, " + ws_RjPtQty + " as qRjPtQty, " + ws_PrdQty + " as qPrdQty, " + _
              ws_PrdPtQty + " as qPrdPtQty, TdByWrk, IsNull(Lc.LocPrntCd, '') as LocPrntCd," + IIF(OptLoc(0).Value = True, "TdByLoc", "Lc.LocPrcs") + " as TdByLoc, TDt, DmCtg, DmPrdCtg, BOdDmCd, OdKt, " + _
              "BOdTc, BOdYy, BOdChr, BOdNo, OmCmCd, OmPoNo " + _
              "From Txnd Join Txn on TCoCd= TdCoCd and TTc= TdTc and TYy= TdYy and TChr= TdChr and TNo= TdNo " + IIF(gs_Partition = ctCurrPrtn, " and TPrtKey=TdPrtKey ", "") + _
              "Join Bag on BCoCd= TdCoCd and BYy= TdBYy and BChr= TdBChr and BNo= TdBNo " + IIF(gs_Partition = ctCurrPrtn, " and BPrtKey=TdPrtKey ", "") + _
              ws_BagJoin + _
              "Join OrdDsg on OdIdNo=BOdIdNo and OdCoCd= BCoCd and OdTc= BOdTc and OdYy= BOdYy and OdChr= BOdChr and OdNo= BOdNo and OdSr= BOdSr " + IIF(gs_Partition = ctCurrPrtn, " and OdPrtKey=BPrtKey ", "") + _
              "Join OrdMst on OmIdNo=BOmIdNo and OmCoCd= BCoCd and OmTc= BOdTc and OmYy= BOdYy and OmChr= BOdChr and OmNo= BOdNo " + IIF(gs_Partition = ctCurrPrtn, " and OmPrtKey=BPrtKey ", "") + _
              "Join DsgMst on DmIdNo=BDmIdNo and DmTcTyp=BOdDmTcTyp and DmCd= BOdDmCd and DmSz= '' " + IIF(gs_Partition = ctCurrPrtn, " and DmPrtKey=BPrtKey ", "") + _
              "left outer Join Loc Lc on LocCoCd= TdCoCd and LocCd= TdByLoc " + _
              "where TdPrdYN= 'Y' and TdSrNo= 0 " + IIF(gs_Partition = ctCurrPrtn, " and TdPrtKey='" + ctCurrPrtn + "' ", "") + wCnd + IIF(wAddlCnd <> "", " And " + wAddlCnd, "") + ") Tbl " + ws_ShowAllPrd
  '*** (Jen 2.13)

              '*** (Bef 3.2.0)
              '"Join Loc Lc on LocCoCd= TdCoCd and LocCd= TdByLoc "
              '*** (Bef 3.2.0)

  wSqlStrg = wSqlStrg1 + " Union All " + wSqlStrg2
  adc.RepRecSource = wSqlStrg + ctMaxDopOpt  '****** Sachin 3.01 added ctMaxDopOpt
  Rep.TxtFrDt.SetText (adc("UoDtFr"))
  Rep.TxtToDt.SetText (adc("UoDtTo"))
  
  If adc("UoYN") = "N" Then
    Rep.BoxHd.Suppress = True
    Rep.BoxDet.Suppress = True
    Rep.BoxFt.Suppress = True
    Rep.DET.Suppress = True
    Rep.FldDet2BagNo.Suppress = False
  Else
    Rep.FldDet2BagNo.Suppress = True
  End If
    
  ' **** Zubin 211 **** '
  With Rep
    If UCase(adc("UOYN1")) = "Y" Then
      .h1RjQty.SetText "RjQty": .h2RjQty.SetText "RjQty"
      .h1RjPtQty.Suppress = False: .h2RjPtQty.Suppress = False
      .h1PrdQty.SetText "PrdQty": .h2PrdQty.SetText "PrdQty"
      .h1PrdPtQty.Suppress = False: .h2PrdPtQty.Suppress = False
      .hDetQtyWt.SetText "Qty": .hDetPtQty.Suppress = False
      
      .FldDetQtyWt.DecimalPlaces = 1: .FldDetQtyWt.RoundingType = crRoundToTenth
      
      .g4_wRjQtyWt.DecimalPlaces = 1: .g4_wRjQtyWt.RoundingType = crRoundToTenth
      .g3_wRjQtyWt.DecimalPlaces = 1: .g3_wRjQtyWt.RoundingType = crRoundToTenth
      .g2_wRjQtyWt.DecimalPlaces = 1: .g2_wRjQtyWt.RoundingType = crRoundToTenth
      .g1_wRjQtyWt.DecimalPlaces = 1: .g1_wRjQtyWt.RoundingType = crRoundToTenth
      .g_wRjQtyWt.DecimalPlaces = 1: .g_wRjQtyWt.RoundingType = crRoundToTenth
      .g4_wPrdQtyWt.DecimalPlaces = 1: .g4_wPrdQtyWt.RoundingType = crRoundToTenth
      .g3_wPrdQtyWt.DecimalPlaces = 1: .g3_wPrdQtyWt.RoundingType = crRoundToTenth
      .g2_wPrdQtyWt.DecimalPlaces = 1: .g2_wPrdQtyWt.RoundingType = crRoundToTenth
      .g1_wPrdQtyWt.DecimalPlaces = 1: .g1_wPrdQtyWt.RoundingType = crRoundToTenth
      .g_wPrdQtyWt.DecimalPlaces = 1: .g_wPrdQtyWt.RoundingType = crRoundToTenth
    Else
      .h1RjQty.SetText "RjGrsWt": .h2RjQty.SetText "RjGrsWt"
      .h1RjPtQty.Suppress = True: .h2RjPtQty.Suppress = True
      .h1PrdQty.SetText "PrdGrsWt": .h2PrdQty.SetText "PrdGrsWt"
      .h1PrdPtQty.Suppress = True: .h2PrdPtQty.Suppress = True
      .hDetQtyWt.SetText "GrsWt": .hDetPtQty.Suppress = True
      
      .FldDetQtyWt.DecimalPlaces = 3: .FldDetQtyWt.RoundingType = crRoundToThousandth
      
      .g4_wRjQtyWt.DecimalPlaces = 3: .g4_wRjQtyWt.RoundingType = crRoundToThousandth
      .g3_wRjQtyWt.DecimalPlaces = 3: .g3_wRjQtyWt.RoundingType = crRoundToThousandth
      .g2_wRjQtyWt.DecimalPlaces = 3: .g2_wRjQtyWt.RoundingType = crRoundToThousandth
      .g1_wRjQtyWt.DecimalPlaces = 3: .g1_wRjQtyWt.RoundingType = crRoundToThousandth
      .g_wRjQtyWt.DecimalPlaces = 3: .g_wRjQtyWt.RoundingType = crRoundToThousandth
      .g4_wPrdQtyWt.DecimalPlaces = 3: .g4_wPrdQtyWt.RoundingType = crRoundToThousandth
      .g3_wPrdQtyWt.DecimalPlaces = 3: .g3_wPrdQtyWt.RoundingType = crRoundToThousandth
      .g2_wPrdQtyWt.DecimalPlaces = 3: .g2_wPrdQtyWt.RoundingType = crRoundToThousandth
      .g1_wPrdQtyWt.DecimalPlaces = 3: .g1_wPrdQtyWt.RoundingType = crRoundToThousandth
      .g_wPrdQtyWt.DecimalPlaces = 3: .g_wPrdQtyWt.RoundingType = crRoundToThousandth
    End If
  End With
  ' **** Zubin 211 **** '
    '*** Jay 2.13(CT) ***
  If adc("UoBagTyp") = "F" Then
    Rep.TxtBagTypAndChr.SetText "(Flute Bags)"
  ElseIf adc("UoBagTyp") = "P" Then
    Rep.TxtBagTypAndChr.SetText "(Primary Bags)"
  ElseIf adc("UoBagTyp") = "C" Then
    Rep.TxtBagTypAndChr.SetText "(Component Bags)"
  End If
'  '***(Jen 2.11)
'  If adc("UoYn2") <> "" Then Rep.TxtBagTypAndChr.SetText IIF(adc("UoYn2") = "Y", "(Flute Bags", "(Primary Bags") & ")"
'  '***(Jen 2.11)
  '*** Jay 2.13(CT) ***
  
  CRV_REP.DisplayGroupTree = False
  
  Call DispReport(adc) '****** Sachin 3.02 - Display Report ******
  
End Sub

' ****** Manali 3.5.0 - 19/11/08 - Check Boxes added
Private Sub ChkBoxArr_Click(Index As Integer)
With ChkBoxArr(Index)
  Select Case Index
  Case Is = 0     'Tag = "ShowBoxDet"
    If .Value = Checked Then adc("UoYN") = "Y" Else adc("UoYN") = "N"
  Case Is = 3     'Tag = "ShowPrdLine"
    If .Value = Checked Then adc("UoYN3") = "Y" Else adc("UoYN3") = "N"
  End Select
End With
End Sub
' ****** Manali 3.5.0 - 19/11/08 - Check Boxes added

' ****** Manali 3.5.0 - 19/11/08 - Check Boxes added
Private Sub ChkBoxArr_GotFocus(Index As Integer)
  DispMsg ChkBoxArr(Index).ToolTipText, etInfo
End Sub
' ****** Manali 3.5.0 - 19/11/08 - Check Boxes added

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
    Set adc.FirNKeyCtl = adc("UoBYyFr")
  End If
  '****** Sachin 2.14.0 - [23-11-2006] - [Multi-Company] ******

  Call InitProp(Me)
  
  '*** Parameters Used For Setting Properties Of UsrOpt Fields ***
  gs_LocTyp = "'P'"
  'gb_RmCtgFor = True
  'gb_RmZ = False
  'indigo.164 sketch design tctype and order tc added
  gs_DmTcTyp = "DM','SM"
  gs_OmTcTyp = "SO','OS"
  'gs_TxnTcTyp = "BV"
  
  gs_CmCtg = "C"
  'gs_RmGrp = "ORDDET"
  'gs_LabGrp = "ORDDET"
  '*** Parameters Used For Setting Properties Of UsrOpt Fields ***
  
  '****** Sachin 2.14.0 - [23-11-2006] - [Multi-Company] ******
  If ws_MultiCoMod = True Then
      adc("UOCOCDFR").Enabled = True: adc("UOCOCDTO").Visible = True: adc("UOCOCDSEL").Visible = True
      LblPrntCd.Visible = True: adc("UOPRNTCDFR").Visible = True: adc("UOPRNTCDTO").Visible = True: adc("UOPRNTCDSEL").Visible = True
  Else
      adc("UOCOCDFR").Enabled = False: adc("UOCOCDTO").Visible = False: adc("UOCOCDSEL").Visible = False
      LblPrntCd.Visible = False: adc("UOPRNTCDFR").Visible = False: adc("UOPRNTCDTO").Visible = False: adc("UOPRNTCDSEL").Visible = False
  End If
  '****** Sachin 2.14.0 - [23-11-2006] - [Multi-Company] ******
  
  Call SetGroupSort
  
End Sub

Private Sub ATXT_FldChange(Index As Integer, ByVal IdName As String, Cancel As Boolean, ErrMsg As String, ByVal pv_NewValue As Variant)
  gb_CoCdFor = False  '****** Sachin 2.14.0 - [Multi-Company]
  Call SetProp(Me, IdName, Valid, Cancel, ErrMsg)
End Sub
Private Sub ATXT_FldWhen(Index As Integer, ByVal IdName As String, Cancel As Boolean, ErrMsg As String)
  gb_CoCdFor = False  '****** Sachin 2.14.0 - [Multi-Company]
  gs_LocTyp = "'P'"
  'indigo.164 sketch design tctype and order tc added
  gs_DmTcTyp = "DM','SM"
  gs_OmTcTyp = "SO','OS"
   Select Case UCase(IdName)
      Case Is = UCase("UOISSLOCFR"), UCase("UOISSLOCTO"), UCase("UOISSLOCSEL")
        If adc("UOLOCTYPFR") = "R" Then gs_LocTyp = "'R'"
        If adc("UOLOCTYPFR") = "P" Then gs_LocTyp = "'P'"
       
   End Select
  SetProp Me, IdName, When
End Sub

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

' ****** Manali 3.5.0 - 03/11/08 - Radio Buttons Added
Private Sub OptQW_Click(Index As Integer)
Select Case Index
Case Is = 0
  adc("UoYN1") = "Y"
Case Is = 1
  adc("UoYN1") = "N"
End Select
End Sub
' ****** Manali 3.5.0 - 03/11/08 - Radio Buttons Added

' ****** Manali 3.5.0 - 05/11/08 - Radio Buttons Added
Private Sub OptQW_GotFocus(Index As Integer)
  DispMsg FraOptQW.ToolTipText, etInfo
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
  
  '****** Sachin 2.14.0 - [25-11-2006] - [Multi-Company] ******
  If ws_MultiCoMod = False Then _
      adc("UoCoCdFr") = gs_CoCd: adc("UoCoCdTo") = gs_CoCd: adc("UoCoCdSel") = "": adc("UoCoCdFr").Enabled = False
  '****** Sachin 2.14.0 - [25-11-2006] - [Multi-Company] ******
  
  ' **** Zubin 211 **** '
  If adc("UOYN1") = "" Then adc("UOYN1") = "Y"
  ' **** Zubin 211 **** '
   '*** Jay 2.13(CT) ***
  Dim wHeadChk As MwfLib.MDORowSet
  Set wHeadChk = moCn.OpenRes("Select hFluteBagYN, hCompBagsYn from Head where HCoCd= '" + gs_CoCd + "' and " + _
                                                    "HCd= '" + ctSelfCmCd + "' ")
  If wHeadChk!hFluteBagYN = "N" And wHeadChk!hCompBagsYn = "N" Then
    LblBagTyp.Visible = False
    adc("UoBagTyp").Visible = False
    adc("UoBagTyp") = "P"
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

  ' ****** Manali 3.5.0 - 03/11/08 - Radio Buttons Added
  If adc("UoYN1") = "Y" Then OptQW(0).Value = True
  If adc("UoYN1") = "N" Then OptQW(1).Value = True
  ' ****** Manali 3.5.0 - 03/11/08 - Radio Buttons Added

  ' ****** Manali 3.5.0 - 19/11/08 - Check Boxes added
  If adc("UoYN") = "" Then adc("UoYN") = "N"
  If adc("UoYN3") = "" Then adc("UoYN3") = "N"
  
  If adc("UoYN") = "Y" Then ChkBoxArr(0).Value = Checked Else ChkBoxArr(0).Value = Unchecked
  If adc("UoYN3") = "Y" Then ChkBoxArr(3).Value = Checked Else ChkBoxArr(3).Value = Unchecked
  ' ****** Manali 3.5.0 - 19/11/08 - Check Boxes added
  If adc("UOLOCTYPFR") = "P" Then
    OptLoc(0).Value = True
  ElseIf adc("UOLOCTYPFR") = "R" Then
    OptLoc(1).Value = True
  End If

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
