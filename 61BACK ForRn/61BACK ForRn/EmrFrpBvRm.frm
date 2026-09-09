VERSION 5.00
Object = "{8767A745-088E-4CA6-8594-073D6D2DE57A}#9.2#0"; "crviewer9.dll"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "tabctl32.ocx"
Object = "{AB2CC32A-E652-42A2-B341-26ECE02F6386}#2.1#0"; "mwfctlforspeed9.ocx"
Begin VB.Form EmrFrpBvRm 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Bag Movement RM"
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
      Left            =   12330
      TabIndex        =   68
      TabStop         =   0   'False
      Top             =   9720
      Visible         =   0   'False
      Width           =   495
      _ExtentX        =   873
      _ExtentY        =   873
   End
   Begin MwfCtl.MWCTL_DST DST_REP 
      Height          =   570
      Left            =   4455
      TabIndex        =   67
      TabStop         =   0   'False
      Top             =   9630
      Width           =   6330
      _ExtentX        =   11165
      _ExtentY        =   1005
   End
   Begin VB.Frame fra_keys 
      BorderStyle     =   0  'None
      Height          =   510
      Left            =   6285
      TabIndex        =   73
      Top             =   9585
      Width           =   3555
      Begin MwfCtl.MWCTL_MED aktxt 
         Height          =   345
         Index           =   0
         Left            =   210
         TabIndex        =   74
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
         TabIndex        =   75
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
      Left            =   7500
      TabIndex        =   72
      Top             =   9660
      Visible         =   0   'False
      Width           =   855
      Begin CRVIEWER9LibCtl.CRViewer9 CRV_REP 
         Height          =   495
         Left            =   480
         TabIndex        =   71
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
      TabIndex        =   69
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
      TabIndex        =   70
      Top             =   60
      Width           =   15255
      Begin TabDlg.SSTab TAB_REP 
         Height          =   9465
         Left            =   60
         TabIndex        =   76
         Top             =   30
         Width           =   15195
         _ExtentX        =   26802
         _ExtentY        =   16695
         _Version        =   393216
         Tabs            =   2
         TabsPerRow      =   2
         TabHeight       =   520
         TabCaption(0)   =   "&Data Range"
         TabPicture(0)   =   "EmrFrpBvRm.frx":0000
         Tab(0).ControlEnabled=   -1  'True
         Tab(0).Control(0)=   "fra_tabrep(0)"
         Tab(0).Control(0).Enabled=   0   'False
         Tab(0).ControlCount=   1
         TabCaption(1)   =   "&Group Sort"
         TabPicture(1)   =   "EmrFrpBvRm.frx":001C
         Tab(1).ControlEnabled=   0   'False
         Tab(1).Control(0)=   "fra_tabrep(1)"
         Tab(1).ControlCount=   1
         Begin VB.Frame fra_tabrep 
            BorderStyle     =   0  'None
            Height          =   9070
            Index           =   0
            Left            =   120
            TabIndex        =   79
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
               Index           =   1
               Left            =   30
               TabIndex        =   124
               Tag             =   "ShowPic"
               ToolTipText     =   "Check To Show Stock Rate"
               Top             =   6270
               Width           =   1950
            End
            Begin VB.Frame FraLoc 
               BorderStyle     =   0  'None
               Height          =   285
               Left            =   1800
               TabIndex        =   121
               Top             =   720
               Width           =   4305
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
                  Left            =   2500
                  TabIndex        =   123
                  ToolTipText     =   "Select Production or Process"
                  Top             =   0
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
                  Left            =   0
                  TabIndex        =   122
                  ToolTipText     =   "Select Production or Process"
                  Top             =   0
                  Value           =   -1  'True
                  Width           =   2085
               End
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
               Index           =   0
               Left            =   30
               TabIndex        =   65
               Tag             =   "ShowBoxDet"
               ToolTipText     =   "Check To Show box Details"
               Top             =   7755
               Width           =   1945
            End
            Begin VB.Frame FraOptBag 
               BorderStyle     =   0  'None
               Height          =   285
               Left            =   1785
               TabIndex        =   111
               ToolTipText     =   "Show Open Or Closed Or All Bags"
               Top             =   7470
               Width           =   4395
               Begin VB.OptionButton OptBag 
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
                  TabIndex        =   63
                  ToolTipText     =   "Show Open Or Closed Or All Bags"
                  Top             =   0
                  Width           =   1005
               End
               Begin VB.OptionButton OptBag 
                  Caption         =   "Closed"
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
                  TabIndex        =   62
                  ToolTipText     =   "Show Open Or Closed Or All Bags"
                  Top             =   0
                  Width           =   945
               End
               Begin VB.OptionButton OptBag 
                  Caption         =   "Open"
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
                  TabIndex        =   61
                  ToolTipText     =   "Show Open Or Closed Or All Bags"
                  Top             =   0
                  Width           =   1005
               End
            End
            Begin VB.TextBox TxtPwd 
               Height          =   285
               IMEMode         =   3  'DISABLE
               Left            =   1785
               MaxLength       =   30
               PasswordChar    =   "*"
               TabIndex        =   54
               Top             =   5670
               Width           =   2340
            End
            Begin MwfCtl.MWCTL_MED atxt 
               Height          =   285
               Index           =   1
               Left            =   4215
               TabIndex        =   2
               ToolTipText     =   "Enter Issue Location To Range"
               Top             =   1080
               Width           =   1830
               _ExtentX        =   3228
               _ExtentY        =   503
               MaxLength       =   8
               DataField       =   "UOISSLOCTO"
               IdName          =   "UOISSLOCTO"
            End
            Begin MwfCtl.MWCTL_MED atxt 
               Height          =   285
               Index           =   2
               Left            =   6630
               TabIndex        =   3
               ToolTipText     =   "Enter Issue Location Selection"
               Top             =   1080
               Width           =   7950
               _ExtentX        =   14023
               _ExtentY        =   503
               MaxLength       =   255
               DataField       =   "UOISSLOCSEL"
               IdName          =   "UOISSLOCSEL"
            End
            Begin MwfCtl.MWCTL_MED atxt 
               Height          =   285
               Index           =   0
               Left            =   1785
               TabIndex        =   1
               ToolTipText     =   "Enter Issue Location From Range"
               Top             =   1080
               Width           =   1830
               _ExtentX        =   3228
               _ExtentY        =   503
               MaxLength       =   8
               DataField       =   "UOISSLOCFR"
               IdName          =   "UOISSLOCFR"
            End
            Begin MwfCtl.MWCTL_MED atxt 
               Height          =   285
               Index           =   8
               Left            =   4215
               TabIndex        =   29
               ToolTipText     =   "Enter Raw Material Code To Range"
               Top             =   3390
               Width           =   2385
               _ExtentX        =   4207
               _ExtentY        =   503
               MaxLength       =   16
               DataField       =   "UORMCDTO"
               IdName          =   "UORMCDTO"
               CmpStr          =   "TdRmCd <="
            End
            Begin MwfCtl.MWCTL_MED atxt 
               Height          =   285
               Index           =   15
               Left            =   6630
               TabIndex        =   30
               ToolTipText     =   "Enter Raw Material Code Selection"
               Top             =   3390
               Width           =   7950
               _ExtentX        =   14023
               _ExtentY        =   503
               MaxLength       =   255
               DataField       =   "UORMCDSEL"
               IdName          =   "UORMCDSEL"
               CmpStr          =   "TdRmCd In"
            End
            Begin MwfCtl.MWCTL_MED atxt 
               Height          =   285
               Index           =   4
               Left            =   4215
               TabIndex        =   23
               ToolTipText     =   "Enter Raw Material Sub Category To Range"
               Top             =   2820
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
               Index           =   5
               Left            =   6630
               TabIndex        =   24
               ToolTipText     =   "Enter Raw Material Sub Category Selection"
               Top             =   2820
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
               Left            =   1785
               TabIndex        =   19
               ToolTipText     =   "Enter From Raw Material Category"
               Top             =   2535
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
               Index           =   7
               Left            =   1785
               TabIndex        =   22
               ToolTipText     =   "Enter Raw Material Sub Category From Range"
               Top             =   2820
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
               Index           =   9
               Left            =   1785
               TabIndex        =   28
               ToolTipText     =   "Enter Raw Material Code From Range"
               Top             =   3390
               Width           =   2385
               _ExtentX        =   4207
               _ExtentY        =   503
               MaxLength       =   16
               DataField       =   "UORMCDFR"
               IdName          =   "UORMCDFR"
               CmpStr          =   "TdRmCd >="
            End
            Begin MwfCtl.MWCTL_MED atxt 
               Height          =   285
               Index           =   10
               Left            =   4215
               TabIndex        =   35
               ToolTipText     =   "Enter To Raw Material Len/Sv"
               Top             =   3960
               Width           =   1035
               _ExtentX        =   1826
               _ExtentY        =   503
               Alignment       =   2
               Mask            =   "#0.0000"
               MaxLength       =   7
               DataType        =   2
               DataField       =   "UOLN1TO"
               IdName          =   "UOLN1TO"
               CmpStr          =   "TdRmSz <="
            End
            Begin MwfCtl.MWCTL_MED atxt 
               Height          =   285
               Index           =   11
               Left            =   1785
               TabIndex        =   34
               ToolTipText     =   "Enter From Raw Material Len/Sv"
               Top             =   3960
               Width           =   1035
               _ExtentX        =   1826
               _ExtentY        =   503
               Alignment       =   2
               Mask            =   "#0.0000"
               MaxLength       =   7
               DataType        =   2
               DataField       =   "UOLN1FR"
               IdName          =   "UOLN1FR"
               CmpStr          =   "TdRmSz >="
            End
            Begin MwfCtl.MWCTL_MED atxt 
               Height          =   285
               Index           =   12
               Left            =   4215
               TabIndex        =   20
               ToolTipText     =   "Enter To Raw Material Category"
               Top             =   2535
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
               Index           =   13
               Left            =   6630
               TabIndex        =   21
               ToolTipText     =   "Enter Raw Material Category Selection"
               Top             =   2535
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
               Index           =   16
               Left            =   4215
               TabIndex        =   37
               ToolTipText     =   "Enter To Cost Rate"
               Top             =   4245
               Width           =   1245
               _ExtentX        =   2196
               _ExtentY        =   503
               Alignment       =   2
               MaxLength       =   0
               DataType        =   2
               DataField       =   "UOCSTRTTO"
               IdName          =   "UOCSTRTTO"
               CmpStr          =   "TdRmStkRt <="
            End
            Begin MwfCtl.MWCTL_MED atxt 
               Height          =   285
               Index           =   18
               Left            =   1785
               TabIndex        =   36
               ToolTipText     =   "Enter From Cost Rate"
               Top             =   4245
               Width           =   1230
               _ExtentX        =   2170
               _ExtentY        =   503
               Alignment       =   2
               MaxLength       =   0
               DataType        =   2
               DataField       =   "UOCSTRTFR"
               IdName          =   "UOCSTRTFR"
               CmpStr          =   "TdRmStkRt >="
            End
            Begin MwfCtl.MWCTL_MED atxt 
               Height          =   285
               Index           =   14
               Left            =   2115
               TabIndex        =   64
               ToolTipText     =   "Show Box Detail (Yes / No)"
               Top             =   7755
               Visible         =   0   'False
               Width           =   315
               _ExtentX        =   556
               _ExtentY        =   503
               Enabled         =   0   'False
               DataType        =   4
               DataField       =   "UOYN"
               IdName          =   "UOYN"
            End
            Begin MwfCtl.MWCTL_MED atxt 
               Height          =   285
               Index           =   31
               Left            =   4215
               TabIndex        =   39
               ToolTipText     =   "Enter To Design Code"
               Top             =   4530
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
               Left            =   6630
               TabIndex        =   40
               ToolTipText     =   "Enter Design Code Selection"
               Top             =   4530
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
               Index           =   33
               Left            =   1785
               TabIndex        =   38
               ToolTipText     =   "Enter From Design Code"
               Top             =   4530
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
               Left            =   4215
               TabIndex        =   52
               ToolTipText     =   "Enter To Customer"
               Top             =   5385
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
               Left            =   6630
               TabIndex        =   53
               ToolTipText     =   "Enter Customer Selection"
               Top             =   5385
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
               Index           =   17
               Left            =   4995
               TabIndex        =   45
               ToolTipText     =   "Enter To Order Year"
               Top             =   4815
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
               Index           =   20
               Left            =   2565
               TabIndex        =   42
               ToolTipText     =   "Enter From Order Year"
               Top             =   4815
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
               Index           =   19
               Left            =   1785
               TabIndex        =   41
               ToolTipText     =   "Enter From Order Tc"
               Top             =   4815
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
               Index           =   21
               Left            =   2985
               TabIndex        =   43
               ToolTipText     =   "Enter From Order Character"
               Top             =   4815
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
               Index           =   22
               Left            =   4215
               TabIndex        =   44
               ToolTipText     =   "Enter To Order Tc"
               Top             =   4815
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
               Index           =   23
               Left            =   5415
               TabIndex        =   46
               ToolTipText     =   "Enter To Order Character"
               Top             =   4815
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
               Left            =   1785
               TabIndex        =   47
               ToolTipText     =   "Enter From Order Number"
               Top             =   5100
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
               Left            =   4215
               TabIndex        =   49
               ToolTipText     =   "Enter To Order Number"
               Top             =   5100
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
               Left            =   5100
               TabIndex        =   50
               ToolTipText     =   "Enter To Order Serial"
               Top             =   5100
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
               Index           =   25
               Left            =   2670
               TabIndex        =   48
               ToolTipText     =   "Enter From Order Serial"
               Top             =   5100
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
               Index           =   26
               Left            =   4215
               TabIndex        =   26
               ToolTipText     =   "Enter To Parent Code"
               Top             =   3105
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
               Index           =   27
               Left            =   6630
               TabIndex        =   27
               ToolTipText     =   "Enter Parent Code Selection"
               Top             =   3105
               Width           =   7950
               _ExtentX        =   14023
               _ExtentY        =   503
               MaxLength       =   255
               DataField       =   "UOPARENTCDSEL"
               IdName          =   "UOPARENTCDSEL"
               CmpStr          =   "RmParentCd In"
            End
            Begin MwfCtl.MWCTL_MED atxt 
               Height          =   285
               Index           =   28
               Left            =   1785
               TabIndex        =   25
               ToolTipText     =   "Enter From Parent Code"
               Top             =   3105
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
               Index           =   29
               Left            =   1785
               TabIndex        =   55
               ToolTipText     =   "Specify If The Value To Be Shown Should Be Stock Value, Cost Value Or Sale Value Or No Value"
               Top             =   5955
               Width           =   1305
               _ExtentX        =   2302
               _ExtentY        =   503
               MaxLength       =   8
               DataField       =   "UOPMCDFR"
               IdName          =   "UOPMCDFR"
            End
            Begin MwfCtl.MWCTL_MED atxt 
               Height          =   285
               Index           =   30
               Left            =   1785
               TabIndex        =   57
               ToolTipText     =   "Enter The Percentage By Which Value Has To Be Incremented"
               Top             =   6600
               Width           =   855
               _ExtentX        =   1508
               _ExtentY        =   503
               Alignment       =   2
               Mask            =   "-##0.000"
               MaxLength       =   8
               DataType        =   2
               DataField       =   "UOPRC"
               IdName          =   "UOPRC"
            End
            Begin MwfCtl.MWCTL_MED atxt 
               Height          =   285
               Index           =   34
               Left            =   1785
               TabIndex        =   0
               ToolTipText     =   "Enter Company Code"
               Top             =   435
               Width           =   705
               _ExtentX        =   1244
               _ExtentY        =   503
               Enabled         =   0   'False
               MaxLength       =   3
               DataType        =   4
               DataField       =   "UOCOCDFR"
               IdName          =   "UOCOCDFR"
               CmpStr          =   "BCoCd = "
            End
            Begin MwfCtl.MWCTL_MED atxt 
               Height          =   285
               Index           =   35
               Left            =   6210
               TabIndex        =   60
               ToolTipText     =   "Show Open Or Closed Or All Bags As On Today? (Open - 'Y', Closed - 'N', All - Blank)"
               Top             =   7470
               Visible         =   0   'False
               Width           =   315
               _ExtentX        =   556
               _ExtentY        =   503
               Enabled         =   0   'False
               DataField       =   "UOYN1"
               IdName          =   "UOYN1"
            End
            Begin MwfCtl.MWCTL_MED atxt 
               Height          =   285
               Index           =   36
               Left            =   4215
               TabIndex        =   5
               ToolTipText     =   "Enter Receiving Location To Range"
               Top             =   1365
               Width           =   1830
               _ExtentX        =   3228
               _ExtentY        =   503
               MaxLength       =   8
               DataField       =   "UORECLOCTO"
               IdName          =   "UORECLOCTO"
            End
            Begin MwfCtl.MWCTL_MED atxt 
               Height          =   285
               Index           =   37
               Left            =   6630
               TabIndex        =   6
               ToolTipText     =   "Enter Receiving Location Selection"
               Top             =   1365
               Width           =   7950
               _ExtentX        =   14023
               _ExtentY        =   503
               MaxLength       =   255
               DataField       =   "UORECLOCSEL"
               IdName          =   "UORECLOCSEL"
            End
            Begin MwfCtl.MWCTL_MED atxt 
               Height          =   285
               Index           =   38
               Left            =   1785
               TabIndex        =   4
               ToolTipText     =   "Enter Receiving Location From Range"
               Top             =   1365
               Width           =   1830
               _ExtentX        =   3228
               _ExtentY        =   503
               MaxLength       =   8
               DataField       =   "UORECLOCFR"
               IdName          =   "UORECLOCFR"
            End
            Begin MwfCtl.MWCTL_MED atxt 
               Height          =   285
               Index           =   42
               Left            =   1785
               TabIndex        =   51
               ToolTipText     =   "Enter From Customer"
               Top             =   5385
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
               Index           =   39
               Left            =   4215
               TabIndex        =   18
               ToolTipText     =   "Enter To Voucher Date"
               Top             =   2250
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
               Index           =   43
               Left            =   1785
               TabIndex        =   17
               ToolTipText     =   "Enter From Voucher Date"
               Top             =   2250
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
               Index           =   57
               Left            =   1785
               TabIndex        =   58
               ToolTipText     =   "Specify Bag Type or ' ' for All"
               Top             =   7170
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
               Index           =   45
               Left            =   3165
               TabIndex        =   59
               ToolTipText     =   "Enter Bag Character"
               Top             =   7170
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
               Index           =   44
               Left            =   1785
               TabIndex        =   31
               ToolTipText     =   "Enter From Lot No."
               Top             =   3675
               Width           =   1905
               _ExtentX        =   3360
               _ExtentY        =   503
               MaxLength       =   12
               DataType        =   4
               DataField       =   "UOLOTNOFR"
               IdName          =   "UOLOTNOFR"
               CmpStr          =   "Td.TdLotNo>= "
            End
            Begin MwfCtl.MWCTL_MED atxt 
               Height          =   285
               Index           =   46
               Left            =   4215
               TabIndex        =   32
               ToolTipText     =   "Enter To Lot No."
               Top             =   3675
               Width           =   1905
               _ExtentX        =   3360
               _ExtentY        =   503
               MaxLength       =   0
               DataType        =   4
               DataField       =   "UOLOTNOTO"
               IdName          =   "UOLOTNOTO"
               CmpStr          =   "Td.TdLotNo<= "
            End
            Begin MwfCtl.MWCTL_MED atxt 
               Height          =   285
               Index           =   47
               Left            =   6630
               TabIndex        =   33
               ToolTipText     =   "Enter Lot No. Selection"
               Top             =   3675
               Width           =   7950
               _ExtentX        =   14023
               _ExtentY        =   503
               MaxLength       =   255
               DataType        =   4
               DataField       =   "UOLOTNOSEL"
               IdName          =   "UOLOTNOSEL"
               CmpStr          =   "Td.TdLotNo In "
            End
            Begin MwfCtl.MWCTL_MED atxt 
               Height          =   285
               Index           =   3
               Left            =   5025
               TabIndex        =   11
               ToolTipText     =   "Enter To Transaction Year"
               Top             =   1665
               Width           =   345
               _ExtentX        =   609
               _ExtentY        =   503
               MaxLength       =   2
               DataField       =   "UOTYYTO"
               IdName          =   "UOTYYTO"
               CmpStr          =   "Bv.TdYy <="
            End
            Begin MwfCtl.MWCTL_MED atxt 
               Height          =   285
               Index           =   48
               Left            =   2595
               TabIndex        =   8
               ToolTipText     =   "Enter From Transaction Year"
               Top             =   1665
               Width           =   345
               _ExtentX        =   609
               _ExtentY        =   503
               MaxLength       =   2
               DataField       =   "UOTYYFR"
               IdName          =   "UOTYYFR"
               CmpStr          =   "Bv.TdYy >="
            End
            Begin MwfCtl.MWCTL_MED atxt 
               Height          =   285
               Index           =   49
               Left            =   1785
               TabIndex        =   7
               ToolTipText     =   "Enter From Transaction Tc"
               Top             =   1665
               Width           =   705
               _ExtentX        =   1244
               _ExtentY        =   503
               Enabled         =   0   'False
               MaxLength       =   3
               DataField       =   "UOTTCFR"
               IdName          =   "UOTTCFR"
               CmpStr          =   "Bv.TdTc >="
            End
            Begin MwfCtl.MWCTL_MED atxt 
               Height          =   285
               Index           =   50
               Left            =   3015
               TabIndex        =   9
               ToolTipText     =   "Enter From Transaction Character"
               Top             =   1665
               Width           =   705
               _ExtentX        =   1244
               _ExtentY        =   503
               MaxLength       =   3
               DataType        =   4
               DataField       =   "UOTCHRFR"
               IdName          =   "UOTCHRFR"
               CmpStr          =   "Bv.TdChr >="
            End
            Begin MwfCtl.MWCTL_MED atxt 
               Height          =   285
               Index           =   51
               Left            =   4215
               TabIndex        =   10
               ToolTipText     =   "Enter To Transaction Tc"
               Top             =   1665
               Width           =   705
               _ExtentX        =   1244
               _ExtentY        =   503
               Enabled         =   0   'False
               MaxLength       =   3
               DataField       =   "UOTTCTO"
               IdName          =   "UOTTCTO"
               CmpStr          =   "Bv.TdTc <="
            End
            Begin MwfCtl.MWCTL_MED atxt 
               Height          =   285
               Index           =   52
               Left            =   5490
               TabIndex        =   12
               ToolTipText     =   "Enter To Transaction Character"
               Top             =   1665
               Width           =   705
               _ExtentX        =   1244
               _ExtentY        =   503
               MaxLength       =   3
               DataType        =   4
               DataField       =   "UOTCHRTO"
               IdName          =   "UOTCHRTO"
               CmpStr          =   "Bv.TdChr <="
            End
            Begin MwfCtl.MWCTL_MED atxt 
               Height          =   285
               Index           =   55
               Left            =   5145
               TabIndex        =   16
               ToolTipText     =   "Enter To Transaction Serial"
               Top             =   1965
               Width           =   525
               _ExtentX        =   926
               _ExtentY        =   503
               MaxLength       =   0
               DataType        =   1
               DataField       =   "UOTDSRTO"
               IdName          =   "UOTDSRTO"
               CmpStr          =   "Bv.TdSr <="
            End
            Begin MwfCtl.MWCTL_MED atxt 
               Height          =   285
               Index           =   56
               Left            =   2715
               TabIndex        =   14
               ToolTipText     =   "Enter From Transaction Serial"
               Top             =   1965
               Width           =   525
               _ExtentX        =   926
               _ExtentY        =   503
               MaxLength       =   0
               DataType        =   1
               DataField       =   "UOTDSRFR"
               IdName          =   "UOTDSRFR"
               CmpStr          =   "Bv.TdSr >="
            End
            Begin MwfCtl.MWCTL_MED atxt 
               Height          =   285
               Index           =   58
               Left            =   1785
               TabIndex        =   13
               ToolTipText     =   "Enter From Transaction Number"
               Top             =   1965
               Width           =   825
               _ExtentX        =   1455
               _ExtentY        =   503
               MaxLength       =   0
               DataType        =   1
               DataField       =   "UOTNOFR"
               IdName          =   "UOTNOFR"
               CmpStr          =   "Bv.TdNo >="
            End
            Begin MwfCtl.MWCTL_MED atxt 
               Height          =   285
               Index           =   59
               Left            =   4215
               TabIndex        =   15
               ToolTipText     =   "Enter To Transaction Number"
               Top             =   1965
               Width           =   825
               _ExtentX        =   1455
               _ExtentY        =   503
               MaxLength       =   0
               DataType        =   1
               DataField       =   "UOTNOTO"
               IdName          =   "UOTNOTO"
               CmpStr          =   "Bv.TdNo <="
            End
            Begin MwfCtl.MWCTL_MED atxt 
               Height          =   285
               Index           =   84
               Left            =   5505
               TabIndex        =   56
               ToolTipText     =   "Specify If The Other Value To Be Shown Should Be Stock Value, Cost Value Or Sale Value Or No Value"
               Top             =   5955
               Width           =   1305
               _ExtentX        =   2302
               _ExtentY        =   503
               MaxLength       =   8
               DataField       =   "UOPMCDTO"
               IdName          =   "UOPMCDTO"
            End
            Begin MwfCtl.MWCTL_MED atxt 
               Height          =   240
               Index           =   89
               Left            =   2400
               TabIndex        =   125
               ToolTipText     =   "Show Custom Pure Wt? (Y/N)"
               Top             =   6240
               Visible         =   0   'False
               Width           =   285
               _ExtentX        =   503
               _ExtentY        =   423
               Enabled         =   0   'False
               DataField       =   "UOYN3"
               IdName          =   "UOYN3"
            End
            Begin MwfCtl.MWCTL_MED atxt 
               Height          =   285
               Index           =   60
               Left            =   6720
               TabIndex        =   126
               ToolTipText     =   "Show Production or Process Location"
               Top             =   720
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
               Left            =   1785
               TabIndex        =   66
               ToolTipText     =   "Show Production Bags (Yes / No)"
               Top             =   8040
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
               TabIndex        =   128
               Top             =   8040
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
               Index           =   18
               Left            =   60
               TabIndex        =   127
               Top             =   8050
               Width           =   1605
            End
            Begin VB.Label Label3 
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
               Left            =   3480
               TabIndex        =   120
               ToolTipText     =   "Location"
               Top             =   5955
               Width           =   1845
            End
            Begin VB.Label ALBL 
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
               Left            =   2655
               TabIndex        =   119
               Top             =   1965
               Width           =   105
            End
            Begin VB.Label ALBL 
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
               Left            =   5055
               TabIndex        =   118
               Top             =   1965
               Width           =   105
            End
            Begin VB.Label ALBL 
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
               Left            =   2535
               TabIndex        =   117
               Top             =   1680
               Width           =   105
            End
            Begin VB.Label ALBL 
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
               Left            =   4920
               TabIndex        =   116
               Top             =   1680
               Width           =   105
            End
            Begin VB.Label ALBL 
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
               Left            =   2955
               TabIndex        =   115
               Top             =   1680
               Width           =   105
            End
            Begin VB.Label ALBL 
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
               TabIndex        =   114
               Top             =   1680
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
               TabIndex        =   113
               Top             =   1965
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
               TabIndex        =   112
               Top             =   1680
               Width           =   1605
            End
            Begin VB.Label LblBagTyp 
               BackStyle       =   0  'Transparent
               Caption         =   "Enter Bag Typ"
               BeginProperty Font 
                  Name            =   "Arial"
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
               TabIndex        =   110
               ToolTipText     =   "Location"
               Top             =   7170
               Width           =   1395
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
               TabIndex        =   109
               Top             =   3675
               Width           =   1425
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
               Left            =   2310
               TabIndex        =   108
               Top             =   7170
               Width           =   885
            End
            Begin VB.Label LblVchDt 
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
               Left            =   60
               TabIndex        =   107
               Top             =   2265
               Width           =   1425
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
               Index           =   29
               Left            =   60
               TabIndex        =   106
               Top             =   5670
               Width           =   1575
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
               TabIndex        =   105
               Top             =   5385
               Width           =   1425
            End
            Begin VB.Label LbRecLoc 
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
               Left            =   60
               TabIndex        =   104
               Top             =   1380
               Width           =   1515
            End
            Begin VB.Label ALBL 
               BackStyle       =   0  'Transparent
               Caption         =   "Open Bags"
               BeginProperty Font 
                  Name            =   "Arial"
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
               TabIndex        =   103
               Top             =   7470
               Width           =   1245
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
               TabIndex        =   102
               Top             =   450
               Width           =   1335
            End
            Begin VB.Label ALBL 
               BackStyle       =   0  'Transparent
               Caption         =   "Value Incremental %"
               BeginProperty Font 
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
               Index           =   0
               Left            =   60
               TabIndex        =   101
               Top             =   6600
               Width           =   1515
            End
            Begin VB.Label Lbl2ndLine 
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
               Left            =   60
               TabIndex        =   100
               ToolTipText     =   "Location"
               Top             =   5955
               Width           =   1485
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
               Index           =   13
               Left            =   60
               TabIndex        =   99
               Top             =   3120
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
               TabIndex        =   98
               Top             =   4530
               Width           =   1515
            End
            Begin VB.Label ALBL 
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
               Left            =   2505
               TabIndex        =   97
               Top             =   4815
               Width           =   105
            End
            Begin VB.Label ALBL 
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
               Left            =   4935
               TabIndex        =   96
               Top             =   4815
               Width           =   105
            End
            Begin VB.Label ALBL 
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
               Left            =   2925
               TabIndex        =   95
               Top             =   4815
               Width           =   105
            End
            Begin VB.Label ALBL 
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
               Left            =   5355
               TabIndex        =   94
               Top             =   4815
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
               Index           =   21
               Left            =   60
               TabIndex        =   93
               Top             =   5100
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
               TabIndex        =   92
               Top             =   4815
               Width           =   1725
            End
            Begin VB.Label ALBL 
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
               Left            =   5040
               TabIndex        =   91
               Top             =   5100
               Width           =   105
            End
            Begin VB.Label ALBL 
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
               Left            =   2610
               TabIndex        =   90
               Top             =   5100
               Width           =   105
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
               Index           =   10
               Left            =   60
               TabIndex        =   89
               ToolTipText     =   "Location"
               Top             =   7740
               Visible         =   0   'False
               Width           =   1485
            End
            Begin VB.Label ALBL 
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
               Height          =   300
               Index           =   9
               Left            =   60
               TabIndex        =   88
               Top             =   4245
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
               Index           =   8
               Left            =   60
               TabIndex        =   87
               Top             =   3405
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
               Index           =   7
               Left            =   60
               TabIndex        =   86
               Top             =   2835
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
               Index           =   6
               Left            =   60
               TabIndex        =   85
               ToolTipText     =   "Location"
               Top             =   2550
               Width           =   1425
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
               Index           =   5
               Left            =   60
               TabIndex        =   84
               Top             =   3960
               Width           =   1425
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
               TabIndex        =   83
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
               Left            =   4215
               TabIndex        =   82
               Top             =   90
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
               Left            =   6630
               TabIndex        =   81
               Top             =   90
               Width           =   1260
            End
            Begin VB.Label LbIssLoc 
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
               Left            =   60
               TabIndex        =   80
               Top             =   1095
               Width           =   1515
            End
         End
         Begin VB.Frame fra_tabrep 
            BorderStyle     =   0  'None
            Enabled         =   0   'False
            Height          =   9070
            Index           =   1
            Left            =   -74940
            TabIndex        =   77
            Top             =   360
            Width           =   15105
            Begin MwfCtl.MWCTL_GRP GRP_REP 
               Height          =   360
               Left            =   240
               TabIndex        =   78
               Top             =   360
               Width           =   9555
               _ExtentX        =   16854
               _ExtentY        =   635
            End
         End
      End
   End
End
Attribute VB_Name = "EmrFrpBvRm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim Rep As New EmrRepBvRm
Dim moCn As MwfLib.MDOConnection   '*** (Bef speed) bef 24/11/06

'****** Sachin 2.14.0 - [25-11-2006] - [Multi-Company] ******
Dim ws_MultiCoMod As Boolean
'****** Sachin 2.14.0 - [25-11-2006] - [Multi-Company] ******

Private Sub SetGroupSort()
  
  GRP_REP.Add "Location", "BLoc", "", "", "", "", "", "", "(Select LocDesc From Loc Where LocCoCd = BCoCd and LocTyp= 'P' and LocCd= BLoc) "
  GRP_REP.Add "Process", "(select LocPrcs from Loc where LocCoCd= BCoCd and LocCd=BLoc)", "", "", "", "", "", "", "(Select LocDesc From Loc Where LocCoCd = BCoCd and LocCd=(Select LocPrcs From Loc Where LocCoCd = BCoCd and LocCd=BLoc)) "

  GRP_REP.Add "Rm Ctg", "Rtrim(IsNull(RmCtg, ''))", "", "", "", "", "", "", "(Select PDesc From Param Where PTyp= 'RMCTG' and PMCd = IsNull(RmCtg, '')) "
  GRP_REP.Add "Rm SubCtg", "Rtrim(IsNull(RmCtg, ''))+Rtrim(IsNull(RmSCtg, ''))", "'('+ RTrim(IsNull(RmCtg, ''))+ ') '+ IsNull(RmSCtg, '')", "", "", "Rm Ctg", "", "", "(Select PDesc From Param Where PTyp= 'RMSCTG' and PMCd = IsNull(RmCtg, '') and PSCd= IsNull(RmSCtg, ''))"
 
  'GRP_REP.Add "Rm Code", "RmCtg+RmSCtg+RmCd", "'('+ RTrim(RmCtg)+'/'+ RmSCtg+ ') '+ RmCd", "", "", "Rm Ctg,Rm SubCtg", "", "", Array("(Select RmDesc From RmMst Where RmCd= IsNull(Td.TdRmCd, ''))", "(Select RmDesc From RmMst Where RmCd= FrRmCd)")
  GRP_REP.Add "Rm Code", "Rtrim(IsNull(RmCtg, ''))+Rtrim(IsNull(RmSCtg, ''))+Rtrim(IsNull(RmCd, ''))", "'('+ RTrim(IsNull(RmCtg, ''))+'/'+ IsNull(RmSCtg, '')+ ') '+ IsNull(RmCd, '') ", "", "", "Rm Ctg,Rm SubCtg,BaseCd + RmCd,Parent Cd", "", "", "SubString(IsNull(RmDesc, ''), 1, 30)"
  GRP_REP.Add "BaseCd + RmCd", "IsNull(RmBaseCd, '')+IsNull(RmCd, '')", "'('+ RTrim(IsNull(RmBaseCd, '')) + ') '+ IsNull(RmCd, '')", "", "", "Rm Ctg,Rm SubCtg,Rm Code,Parent Cd", "", "", "SubString(IsNull(RmDesc, ''), 1, 30)"
  GRP_REP.Add "Parent Cd", "Rtrim(IsNull(RmCtg, ''))+Rtrim(IsNull(RmSCtg, ''))+(case when IsNull(RmParentCd, '')= '' then IsNull(RmCd, '') else IsNull(RmParentCd, '') end)", "'('+ RTrim(IsNull(RmCtg, ''))+'/'+ IsNull(RmSCtg, '')+ ') '+ (case when IsNull(RmParentCd, '')= '' then IsNull(RmCd, '') else IsNull(RmParentCd, '') end)", "RmCtg,RmSCtg", "hRmCtg,hRmSCtg", "Rm Ctg,Rm SubCtg,BaseCd + RmCd", "", "", "(Select PDesc from Param where PTyp= 'PARENTCD' and PMCd= IsNull(RmParentCd, '')) "
  
  '4.1.0.0
  GRP_REP.Add "Vch No", "Bv.TdTc+'/'+Bv.TdYy+'/'+Bv.TdChr+'/'+Str(Bv.TdNo)", "Bv.TdTc+'/'+Bv.TdYy+'/'+Bv.TdChr+'/'+LTrim(Str(Bv.TdNo))", "", "", "", "", "", ""
  
  '****************Geeta**************Emr206*********12/04/04
  Dim wsPtrYN As String
  wsPtrYN = "(Select PValue3 from Param where PTyp= 'RMSCTG' and PMCd= IsNull(RmCtg, '') and PSCd= IsNull(RmSCtg, ''))"
'  GRP_REP.Add "Rm Sv/Sz", Array("case when " + wsPtrYN + "='Y' then Str(IsNull(Td.TdRmSz, 0),7,4) else Str(IsNull(Td.TdRmSz, 0),6,3) end ", "case when " + wsPtrYN + "='Y' then Str(FrRmSz,7,4) else Str(FrRmSz,6,3) end "), "", "", "", "Rm Ptr,Rm StkRt,(RmCd) Sv/Sz,(RmCd) StkRt", "", "", ""
 
  GRP_REP.Add "Rm Sv/Sz", Array(GetRmSzDBFldNm("Td.TdRmSz", "Td.TdRmSz2", "Td.TdRmSz3"), _
                                GetRmSzDBFldNm("FrRmSz", "FrRmSz2", "FrRmSz3")), "", "", "", "Rm Ptr,Rm StkRt,(RmCd) Sv/Sz,(RmCd) StkRt", "", "", ""
  
  GRP_REP.Add "Rm StkRt", Array("Str(IsNull(Td.TdRmStkRt, 0),9,2)", "Str(FrRmStkRt,9,2)"), "", "", "", "Rm Ptr,Rm Sv/Sz,(RmCd) Sv/Sz,(RmCd) StkRt", "", "", ""
  '****************Geeta**************Emr206*********12/04/04
  'GRP_REP.Add "(RmCd) Sv/Sz", Array("TdRmCd+Str(TdRmSz,6,3)", "FrRmCd+Str(FrRmSz,6,3)"), Array("'('+TdRmCd+') '+LTrim(str(TdRmSz,6,3))", "'('+FrRmCd+') '+LTrim(str(FrRmSz,6,3))"), "", "", "Rm Ctg,Rm SubCtg,Rm Code,BaseCd + RmCd,(RmCd) StkRt,Rm Sv/Sz,Rm Ptr,Rm StkRt", "", "", " '('+ RTrim(RmCtg)+ '/'+ RmSCtg+ ')' "
'  GRP_REP.Add "(RmCd) Sv/Sz", Array("IsNull(Td.TdRmCd, '') +case when " + wsPtrYN + "='Y' then Str(IsNull(Td.TdRmSz, 0),7,4) else Str(IsNull(Td.TdRmSz, 0),6,3) end ", "FrRmCd+case when " + wsPtrYN + "='Y' then Str(FrRmSz,7,4) else Str(FrRmSz,6,3) end "), _
              Array("'('+IsNull(Td.TdRmCd, '')+') '+LTrim(case when " + wsPtrYN + "='Y' then Str(IsNull(Td.TdRmSz, 0),7,4) else Str(IsNull(Td.TdRmSz, 0),6,3) end )", "'('+FrRmCd+') '+LTrim(case when " + wsPtrYN + "='Y' then Str(FrRmSz,7,4) else Str(FrRmSz,6,3) end )"), "", "", "Rm Ctg,Rm SubCtg,Rm Code,BaseCd + RmCd,(RmCd) StkRt,Rm Sv/Sz,Rm Ptr,Rm StkRt", "", "", " '('+ RTrim(IsNull(RmCtg, ''))+ '/'+ IsNull(RmSCtg, '')+ ')' "
  GRP_REP.Add "(RmCd) Sv/Sz", Array("IsNull(Td.TdRmCd, '') + " + GetRmSzDBFldNm("Td.TdRmSz", "Td.TdRmSz2", "Td.TdRmSz3"), _
                                     "FrRmCd+ " + GetRmSzDBFldNm("FrRmSz", "FrRmSz2", "FrRmSz3")), _
                               Array("'('+IsNull(Td.TdRmCd, '')+') ' + " + GetRmSzDispFldNm("Td.TdRmSz", "Td.TdRmSz2", "Td.TdRmSz3"), _
                                     "'('+FrRmCd+') '+ " + GetRmSzDispFldNm("FrRmSz", "FrRmSz2", "FrRmSz3")), _
                                     "", "", "Rm Ctg,Rm SubCtg,Rm Code,BaseCd + RmCd,(RmCd) StkRt,Rm Sv/Sz,Rm Ptr,Rm StkRt", _
                                     "", "", " '('+ RTrim(IsNull(RmCtg, ''))+ '/'+ IsNull(RmSCtg, '')+ ')' "
              
  GRP_REP.Add "(RmCd) StkRt", Array("IsNull(Td.TdRmCd, '')+Str(IsNull(Td.TdRmStkRt, 0),9,2)", "FrRmCd+Str(FrRmStkRt,9,2)"), _
                    Array("'('+IsNull(Td.TdRmCd, '')+') '+LTrim(str(IsNull(Td.TdRmStkRt, 0),9,2))", "'('+FrRmCd+') '+LTrim(str(FrRmStkRt,9,2))"), _
                    "", "", "Rm Ctg,Rm SubCtg,Rm Code,BaseCd + RmCd,(RmCd) Sv/Sz,Rm Sv/Sz,Rm Ptr,Rm StkRt", "", "", " '('+ RTrim(IsNull(RmCtg, ''))+ '/'+ IsNull(RmSCtg, '')+')' "
  
 '4.1.4
 Dim wTdRate As String, wFrRate As String
   wTdRate = " (Case When RSCtg.PValue1='N' Then " + RmRate("RrCstRt", "IsNull(Td.TdRmSz, 0)", "RmQwCst") + " Else Td.TdRmStkRt End)"
   wFrRate = " (Case When RSCtg.PValue1='N' Then " + RmRate("RrCstRt", "FrRmSz", "RmQwCst") + " Else Td.TdRmStkRt End)"
   GRP_REP.Add "(RmCd) Stk-Cst", _
        Array("Td.TdRmCd+ Str(" + wTdRate + ", 9, 2)", "FrRmCd+ Str(" + wFrRate + ", 9, 2)"), _
        Array("'('+Td.TdRmCd+') '+LTrim(str(" + wTdRate + ",9,2))", "'('+FrRmCd+') '+LTrim(str(" + wFrRate + ",9,2))"), _
        "", "", "Rm Ctg,Rm SubCtg,Rm Code,BaseCd + RmCd,(RmCd) Sv/Sz,Rm Sv/Sz,Rm Ptr,Rm StkRt", "", "", " '('+ RTrim(RmCtg)+ '/'+ RmSCtg+')' "


   wTdRate = " (Case When RSCtg.PValue1='N' Then " + RmRate("RrSalRt", "Td.TdRmSz", "RmQw") + " Else Td.TdRmStkRt End)"
   wFrRate = " (Case When RSCtg.PValue1='N' Then " + RmRate("RrSalRt", "FrRmSz", "RmQw") + " Else Td.FrRmStkRt End)"
   GRP_REP.Add "(RmCd) Stk-Sal", _
        Array("Td.TdRmCd+ Str(" + wTdRate + ", 9, 2)", "FrRmCd+ Str(" + wFrRate + ", 9, 2)"), _
        Array("'('+Td.TdRmCd+') '+LTrim(str(" + wTdRate + ",9,2))", "'('+FrRmCd+') '+LTrim(str(" + wFrRate + ",9,2))"), _
        "", "", "Rm Ctg,Rm SubCtg,Rm Code,BaseCd + RmCd,(RmCd) Sv/Sz,Rm Sv/Sz,Rm Ptr,Rm StkRt", "", "", " '('+ RTrim(RmCtg)+ '/'+ RmSCtg+')' "
  
  
  If moCn.GetFldVal("Select HPtrYN from Head where HCoCd ='" + ctSelfCoCd + "' and HCd='" + ctSelfCmCd + "'") = "Y" Then _
     GRP_REP.Add "Rm Ptr", Array("Str(IsNull(Td.TdRmSz, 0),6,3)", "Str(FrRmSz,6,3)"), "", "", "", "Rm Sv/Sz,Rm StkRt,(RmCd) Sv/Sz,(RmCd) StkRt", "", "", ""
   
  GRP_REP.Add "Design", "BOdDmCd", "", "", "", "(DsgCd)Colour", "", "", ""     'Substring(DmDesc,1,120)
  GRP_REP.Add "(DsgCd)Colour", "BOdDmCd+OdDmCol", "'('+BOdDmCd+') '+OdDmCol", "", "", "Design", "", "", ""
  GRP_REP.Add "OrderNo", "BOdTc+BOdYy+BOdChr+str(BOdNo)", "BOdTc+'/'+BOdYy+'/'+BOdChr+'/'+LTrim(Str(BOdNo))", "", "", "Customer", "", "", "'('+OmCmCd+ '/'+ OmPONo+ ')' "
  
  '***** Manali 3.8.0
  GRP_REP.Add "Ord Kt", "OdKt", "", "", "", "", "", "", "(Select PDesc from Param Where PTyp='KT' and PMCd=OdKt) "     'Substring(DmDesc,1,120)

  GRP_REP.Add "OrderSr", "BOdTc+BOdYy+BOdChr+str(BOdNo)+str(BOdSr)", "BOdTc+'/'+BOdYy+'/'+BOdChr+'/'+LTrim(Str(BOdNo))+'/'+LTrim(Str(BOdSr))", "", "", "OrderNo,Design,(DsgCd)Colour,Customer", "", "", "'('+OmCmCd+ '/'+ OmPONo+ ')' "
  GRP_REP.Add "Customer", "OmCmCd", "", "", "", "", "", "", "(Select (case when (Select PNum from Param where Ptyp = 'USR' and PmCd = '" + gs_UsrCd + "')  <= 5 then  CmName else CmCd end) From CustMst Where CmCtg= 'C' and CmCd= OmCmCd) "
  
  GRP_REP.Add "FgSubLoc", "BFgSubLoc", "", "", "", "", "", "", "(Select vPDesc From vParam Where vPCoCd = BCoCd and vPTyp= 'FGSUBLOC' and vPMCd= BFgSubLoc) "
  '**** Shilpa Emr 206 ****
  GRP_REP.Add "BagNo", "BYy+BChr+str(BNo)", "BYy+'/'+BChr+'/'+LTrim(Str(BNo))", "", "", "Design Code,Bag Type,Bag Character", "", "", ""
  '**** Shilpa Emr 206 ****
  
  GRP_REP.Add "Issuing Loc", Array("Bv.TdFrBLoc", "FgFrBLoc"), "", "TdFrBLoc", "hTdFrBLoc", "", "", "", Array("(Select LocDesc From Loc Where LocCoCd= Bv.TdCoCd and LocTyp= 'P' and LocCd = Bv.TdFrBLoc)", "(Select LocDesc From Loc Where LocCoCd= FgCoCd and LocTyp= 'P' and LocCd = FgFrBLoc)")
  GRP_REP.Add "Issuing Prcs", Array("(select LocPrcs from Loc where LocCoCd= Bv.TdCoCd and LocCd=Bv.TdFrBLoc)", "FgFrBLoc"), "", "TdFrBLoc", "hTdFrBLoc", "", "", "", Array("(Select LocDesc From Loc Where LocCoCd= Bv.TdCoCd and LocCd =(Select LocPrcs from Loc where LocCoCd= Bv.TdCoCd and LocCd=Bv.TdFrBLoc))", "(Select LocDesc From Loc Where LocCoCd= FgCoCd and LocTyp= 'P' and LocCd = FgFrBLoc)")
  GRP_REP.Add "Receiving Loc", Array("Bv.TdToBLoc", "FgToBLoc"), "", "TdToBLoc", "hTdToBLoc", "", "", "", Array("(Select LocDesc From Loc Where LocCoCd= Bv.TdCoCd and LocTyp= 'P' and LocCd = Bv.TdToBLoc)", "(Select LocDesc From Loc Where LocCoCd= FgCoCd and LocTyp= 'P' and LocCd = FgToBLoc)")
  GRP_REP.Add "Receiving Prcs", Array("(select LocPrcs from Loc where LocCoCd= Bv.TdCoCd and LocCd=Bv.TdToBLoc)", "FgToBLoc"), "", "TdToBLoc", "hTdToBLoc", "", "", "", Array("(Select LocDesc From Loc Where LocCoCd= Bv.TdCoCd and LocCd =(Select LocPrcs from Loc where LocCoCd= Bv.TdCoCd and LocCd=Bv.TdToBLoc))", "(Select LocDesc From Loc Where LocCoCd= FgCoCd and LocTyp= 'P' and LocCd = FgToBLoc)")
  '6.1
  GRP_REP.Add "Date", "convert(char(10),Bv.TdDt,112)", "convert(varchar, bv.TdDt, " + CStr(DtFmtSlahVal()) + ")", "", "", "", "", "", ""    '4.1.2
'  GRP_REP.Add "Issuing Loc", Array("Bv.TdFrBLoc", "FgFrBLoc"), "TdFrBLoc", "", "", "", "", "", ""

'*** Jay 2.13(CT) ***
  GRP_REP.Add "Bag Type", "(Select (case PValue when 'F' then 'Flute Bag' when 'P' then 'Primary Bag' when 'C' then 'Component Bag' end) from Param where PTyp= 'BCHR' and PMCd= BChr and PSCd= '')", _
    "(Select case PValue when 'F' then 'Flute Bag' when 'P' then 'Primary Bag' when 'C' then 'Component Bag' end from Param where PTyp= 'BCHR' and PMCd= BChr and PSCd= '')", "", "", "Bag Typ+Chr", "", "", "(Select case PValue when 'F' then 'Flute Bag' when 'P' then 'Primary Bag' when 'C' then 'Component Bag' end from Param where PTyp= 'BCHR' and PMCd= BChr and PSCd= '')"
  GRP_REP.Add "Bag Character", "BChr", "BChr", "", "", "Bag Type,Bag Typ+Chr", "", "", "(Select PDesc from Param where PTyp ='BCHR' and PMCd= BChr and PSCd= '')"
  GRP_REP.Add "Bag Typ+Chr", "(Select (case PValue when 'F' then 'Flute Bag' when 'P' then 'Primary Bag' when 'C' then 'Component Bag' end) from Param where PTyp= 'BCHR' and PMCd= BChr and PSCd= '')+BChr", _
    "(Select (case PValue when 'F' then 'Flute Bag' when 'P' then 'Primary Bag' when 'C' then 'Component Bag' end) from Param where PTyp= 'BCHR' and PMCd= BChr and PSCd= '')+'/ '+ BChr", "", "", "Bag Type,Bag Character", "", "", "(Select PDesc from Param where PTyp ='BCHR' and PMCd= BChr and PSCd= '')"
'' **** Zubin 211 **** '
'  GRP_REP.Add "Bag Type", "(case PValue when 'Y' then 'Flute Bag' when 'N' then 'Primary Bag' end)", "(case PValue when 'Y' then 'Flute Bag' when 'N' then 'Primary Bag' end )", "", "", "Bag Typ+Chr", "", "", "(case PValue when 'Y' then 'Flute Bag' when 'N' then 'Primary Bag' end)"
'  GRP_REP.Add "Bag Character", "BChr", "BChr", "", "", "Bag Type,Bag Typ+Chr", "", "", "(Select PDesc from param where PTyp ='BCHR' and PMCd= BChr and PSCd= '')"
'  GRP_REP.Add "Bag Typ+Chr", "(case PValue when 'Y' then 'Flute Bag' when 'N' then 'Primary Bag' end)+BChr", "(case PValue when 'Y' then 'Flute Bag' when 'N' then 'Primary Bag' end)+ '/ ' + BChr", "", "", "Bag Type,Bag Character", "", "", "(Select PDesc from param where PTyp ='BCHR' and PMCd= BChr and PSCd= '')"
'' **** Zubin 211 **** '
'*** Jay 2.13(CT) ***

' Zubin 212
  If moCn.RecSeek("select 'x' from Head where HCd= '" + ctSelfCmCd + "' and HCoCd= '" + adc("UoCoCdFr") + "' and HLotNoYN= 'Y' ") Then
    GRP_REP.Add "RM Lot No.", "Td.TdLotNo", "", "TdLotNo", "hLotNo", "", "", "", "Td.TdLotNo"
  End If
' Zubin 212
End Sub

Private Sub adc_setreprecsource()
On Error GoTo RepErr

'***Geeta***Emr208***New Flds DmTcTyp, BOdDmCd, BOdDmSz and DsgMst Join to Display Picture
  '*** Report Sql ***
' Zubin 212
'  Select space(30) as Grp1, space(30) as Grp2, space(30) as Grp3,
'  space(30) as DGrp1, space(30) as DGrp2, space(30) as DGrp3,
'  space(30) as G1Desc, space(30) as G2Desc, space(30) as G3Desc,
'  BYy, BChr, BNo,DmTcTyp, BOdDmCd, BOdDmSz, BLoc, RmCtg, RmSCtg, TdRmCd, TdLotNo, TdRmSz, TdRmStkRt,
'  PValue3 as qPtrYN, TdRmQty, TdRmWt, 0 as qMul, TdRmWt as qGmWt,
'  TdRmWt as qGldWt, BQty , BPtQty, TDt, TdTc, TdYy, TdChr, TdNo,
'  TdSr, TdSrNo, TdRmStkRt as qCalcRt, RmQW, RmQWCst, TdFrBLoc, TdToBLoc
'  From Bag, Txnd, Txn, RmMst,DsgMst, Param where 1= 2
  '*** Report Sql ***
  
  Dim grpflds0 As String, grpflds1 As String, wrepcnd As String, wCnd As String, wSqlStrg As String
  Dim wPurityWt As String, wSelCnd As String, ws_PrdBag As String
  Dim wTdRate As String, wFrRate As String, wBagClsCnd As String
  Dim wRtStr As String, wStkRtStr As String '4.1.4
  
' **** Zubin 211 **** '
  Dim ws_BagJoin As String
' **** Zubin 211 **** '
  
'****** Sachin 2.13.0 - 05/06/2006 ******
  Dim wQwChk As String
'****** Sachin 2.13.0 - 05/06/2006 ******

  Dim wAddlCnd As String '****** Sachin 3.02.0  - New fields in Where Clause
  
  Set Rep = Nothing
  Set adc.RepSource = Rep
  Call DispCoNm
  Call SetMwName(Rep)   ' ***** Manali 3.6.1 - Microway footer
  Call GRP_REP.Gen3LRep
  'grpflds = GRP_REP.GrpFldLst
  
  grpflds0 = GRP_REP.GrpFldLst(0)
  grpflds1 = GRP_REP.GrpFldLst(1)
  
  '*** In This Rep Grouping On Raw Material Sv/Len or Stk Rt Should Only Be Taken Under Group RM Code ***
  
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
  
    '****** Sachin 3.02.0 - Additional Fields in Where Clause
    wAddlCnd = wAddlCnd + IIF(wAddlCnd <> "" And Trim(adc("UOCMCDFR")) <> "", " And ", "") + IIF(Trim(adc("UOCMCDFR")) <> "", " OdOmCmCd >= '" + Trim(adc("UOCMCDFR")) + "'", "")
    wAddlCnd = wAddlCnd + IIF(wAddlCnd <> "" And Trim(adc("UOCMCDTO")) <> "", " And ", "") + IIF(Trim(adc("UOCMCDTO")) <> "", " OdOmCmCd <= '" + Trim(adc("UOCMCDTO")) + "'", "")
    wAddlCnd = wAddlCnd + IIF(wAddlCnd <> "" And Trim(adc("UOCMCDSEL")) <> "", " And ", "") + IIF(Trim(adc("UOCMCDSEL")) <> "", " OdOmCmCd In (" + Trim(adc("UOCMCDSEL")) + ")", "")
    
    wAddlCnd = wAddlCnd + IIF(wAddlCnd <> "" And Trim(adc("UORMCTGFR")) <> "", " And ", "") + IIF(Trim(adc("UORMCTGFR")) <> "", " Td.TdRmCtg >= '" + Trim(adc("UORMCTGFR")) + "'", "")
    wAddlCnd = wAddlCnd + IIF(wAddlCnd <> "" And Trim(adc("UORMCTGTO")) <> "", " And ", "") + IIF(Trim(adc("UORMCTGTO")) <> "", " Td.TdRmCtg <= '" + Trim(adc("UORMCTGTO")) + "'", "")
    wAddlCnd = wAddlCnd + IIF(wAddlCnd <> "" And Trim(adc("UORMCTGSEL")) <> "", " And ", "") + IIF(Trim(adc("UORMCTGSEL")) <> "", " Td.TdRmCtg In (" + Trim(adc("UORMCTGSEL")) + ")", "")

    wAddlCnd = wAddlCnd + IIF(wAddlCnd <> "" And Trim(adc("UORMSCTGFR")) <> "", " And ", "") + IIF(Trim(adc("UORMSCTGFR")) <> "", " Td.TdRmSCtg >= '" + Trim(adc("UORMSCTGFR")) + "'", "")
    wAddlCnd = wAddlCnd + IIF(wAddlCnd <> "" And Trim(adc("UORMSCTGTO")) <> "", " And ", "") + IIF(Trim(adc("UORMSCTGTO")) <> "", " Td.TdRmSCtg <= '" + Trim(adc("UORMSCTGTO")) + "'", "")
    wAddlCnd = wAddlCnd + IIF(wAddlCnd <> "" And Trim(adc("UORMSCTGSEL")) <> "", " And ", "") + IIF(Trim(adc("UORMSCTGSEL")) <> "", " Td.TdRmSCtg In (" + Trim(adc("UORMSCTGSEL")) + ")", "")

    wAddlCnd = wAddlCnd + IIF(wAddlCnd <> "" And Trim(adc("UODTFR")) <> "01/01/80", " And ", "") + IIF(Trim(adc("UODTFR")) <> "01/01/80", " Bv.TdDt >= '" + Trim(adc("UODTFR")) + "'", "")
    wAddlCnd = wAddlCnd + IIF(wAddlCnd <> "" And Trim(adc("UODTTO")) <> "01/01/80", " And ", "") + IIF(Trim(adc("UODTTO")) <> "01/01/80", " Bv.TdDt <= '" + Trim(adc("UODTTO")) + "'", "")
    '****** Sachin 3.02.0 - Additional Fields in Where Clause
  
  If UCase(adc("UoPMCdFr")) = "STKVAL" Then
    wTdRate = " 0 as qCalcRt "
    wFrRate = " 0 as qCalcRt "
    Rep.TxtValTyp.SetText "(Value = Stock)"
    
'   GRP_REP.SetFormula Rep, "wRmVal", "{@wRmWt} * {rdo.TdRmStkRt}"
'   GRP_REP.SetFormula Rep, "wRmVal", "(If UpperCase({rdo.RmQW}) = UpperCase(""Q"") AND UpperCase({rdo.RmCtg}) = UpperCase(""X"") Then {@wRmQty} Else {@wRmWt}) * {rdo.TdRmStkRt}"
'   ****** Sachin 2.13.0 - 05/06/2006 - Stock Value should always be Weight Wise ******
      
    wStkRtStr = "If Round({@wPureWt}, 3)<> 0 then {@wPureWt}*{rdo.TdRmStkRt} Else {@wRmWt}*{rdo.TdRmStkRt}"
      
    '4.1.4
    If UCase(adc("UoPMCdTo")) = "CSTVAL" Then
        wTdRate = RmRate("RrCstRt", "IsNull(Td.TdRmSz, 0)", "RmQwCst") + " as qCalcRt "
        wFrRate = RmRate("RrCstRt", "FrRmSz", "RmQwCst") + " as qCalcRt "
        Rep.TxtValTyp.SetText "(Value=Stk-Cst)"
        wRtStr = "(IF UpperCase({rdo.qQWCst}) = UpperCase(""Q"") Then {@wRmQty} Else (If Round({@wPureWt}, 3)<> 0 then {@wPureWt} Else {@wRmWt}))*{rdo.qCalcRt}"
        GRP_REP.SetFormula Rep, "wRmVal", " If {rdo.qStkbyRt}='Y' then " + wStkRtStr + " else " + wRtStr + IIF(Round(adc("UoPrc"), 3) = 0, "", "* (1 + (" + CStr(adc("UoPrc")) + "/ 100))")
        
    ElseIf UCase(adc("UoPMCdTo")) = "SALVAL" Then
        wTdRate = RmRate("RrSalRt", "IsNull(Td.TdRmSz, 0)", "RmQw") + " as qCalcRt "
        wFrRate = RmRate("RrSalRt", "FrRmSz", "RmQw") + " as qCalcRt "
        Rep.TxtValTyp.SetText "(Value=Stk-Sal)"
        wRtStr = "(IF UpperCase({rdo.qQW}) = UpperCase(""Q"") Then {@wRmQty} Else (If Round({@wPureWt}, 3)<> 0 then {@wPureWt} Else {@wRmWt})) * {rdo.qCalcRt}"
        GRP_REP.SetFormula Rep, "wRmVal", " If {rdo.qStkbyRt}='Y' then " + wStkRtStr + " else " + wRtStr + IIF(Round(adc("UoPrc"), 3) = 0, "", "* (1 + (" + CStr(adc("UoPrc")) + "/ 100))")
    Else
      GRP_REP.SetFormula Rep, "wRmVal", "If Round({@wPureWt}, 3)<> 0 then {@wPureWt}*{rdo.TdRmStkRt} Else {@wRmWt}*{rdo.TdRmStkRt}"
    End If
      
'   ****** Sachin 2.13.0 - 05/06/2006 - Stock Value should always be Weight Wise ******
  ElseIf UCase(adc("UoPMCdFr")) = "CSTVAL" Then
    wTdRate = RmRate("RrCstRt", "IsNull(Td.TdRmSz, 0)", "RmQwCst") + " as qCalcRt "
    wFrRate = RmRate("RrCstRt", "FrRmSz", "RmQwCst") + " as qCalcRt "
    Rep.TxtValTyp.SetText "(Value = Cost)"
    
  '****** Sachin 2.13.0 - 01-06-2006 - For Q/W Update ******
    If Round(adc("UoPrc"), 3) = 0 Then
      GRP_REP.SetFormula Rep, "wRmVal", "(IF UpperCase({rdo.qQWCst}) = UpperCase(""Q"") Then {@wRmQty} Else (If Round({@wPureWt}, 3)<> 0 then {@wPureWt} Else {@wRmWt})) * {rdo.qCalcRt}"
    Else
      GRP_REP.SetFormula Rep, "wRmVal", "(IF UpperCase({rdo.qQWCst}) = UpperCase(""Q"") Then {@wRmQty} Else (If Round({@wPureWt}, 3)<> 0 then {@wPureWt} Else {@wRmWt})) * {rdo.qCalcRt} * (1 + (" + CStr(adc("UoPrc")) + "/ 100))"
    End If
      
    '****** Sachin 2.13.0 - 01-06-2006 ******
  ElseIf UCase(adc("UoPMCdFr")) = "SALVAL" Then
    wTdRate = RmRate("RrSalRt", "IsNull(Td.TdRmSz, 0)", "RmQw") + " as qCalcRt "
    wFrRate = RmRate("RrSalRt", "FrRmSz", "RmQw") + " as qCalcRt "
    Rep.TxtValTyp.SetText "(Value = Sale)"
    
    If Round(adc("UoPrc"), 3) = 0 Then
      GRP_REP.SetFormula Rep, "wRmVal", "(IF UpperCase({rdo.qQW}) = UpperCase(""Q"") Then {@wRmQty} Else (If Round({@wPureWt}, 3)<> 0 then {@wPureWt} Else {@wRmWt})) * {rdo.qCalcRt}"
    Else
      GRP_REP.SetFormula Rep, "wRmVal", "(IF UpperCase({rdo.qQW}) = UpperCase(""Q"") Then {@wRmQty} Else (If Round({@wPureWt}, 3)<> 0 then {@wPureWt} Else {@wRmWt})) * {rdo.qCalcRt} * (1 + (" + CStr(adc("UoPrc")) + "/ 100))"
    End If
  Else
    wTdRate = " 0 as qCalcRt "
    wFrRate = " 0 as qCalcRt "
    Rep.TxtValTyp.SetText ""
    
    GRP_REP.SetFormula Rep, "wRmVal", "{@wRmWt}* 0"
    Rep.hRmVal.Suppress = True
    Rep.hSecRmVal.Suppress = True
  End If
  
  '*** If Detail not selected then suppress Box Header, Detail and Footer ***
  If wDetPos = 0 Then Rep.BoxHd.Suppress = True: Rep.BoxDet.Suppress = True: Rep.BoxFt.Suppress = True
    
  '*** This Inner Select Gives An Error
  'wPurityWt = " /(case when Rm.RmCtg in ('G', 'P') then " + _
                            "IsNull((Select RmPurityWt from Param join RmMst on " + _
                            "RmCd= PValue where PTyp= 'RMCTG' and PMCd= Rm.RmCtg), 1) " + _
                  "else 1 end) "
  
  '*** (Bef Rel 1.11)
  'wPurityWt = " /(case when Rm.RmCtg in ('G', 'P') then " + _
  '                          "IsNull((Select RmPurityWt from Param, RmMst " + _
  '                          "where RmCd= PValue and PTyp= 'RMCTG' and PMCd= Rm.RmCtg), 1) " + _
  '                "else 1 end) "
  
  ' ***** Manali 3.03 - 23/06/08 - Pd Mod - 'L' added
  wPurityWt = " /(case when IsNull(Rm.RmCtg, '') in ('G', 'P', 'S', 'L') then " + _
                          "IsNull((Select RmPurityWt from Param, RmMst " + _
                          "Where " + IIF(gs_Partition = ctCurrPrtn, " RmPrtKey='" + ctCurrPrtn + "' and ", "") + " RmCd= PValue and PTyp= 'RMCTG' and PMCd= IsNull(Rm.RmCtg, '')), 1) " + _
                 "when IsNull(Rm.RmCtg, '')= 'X' then " + _
                          "IsNull((Select RmPurityWt from Param, RmMst Base " + _
                          "where " + IIF(gs_Partition = ctCurrPrtn, " RmPrtKey='" + ctCurrPrtn + "' and ", "") + " RmCd= PValue and PTyp= 'RMCTG' and " + _
                          "PMCd= IsNull((Select RmCtg From RmMst where " + IIF(gs_Partition = ctCurrPrtn, " RmPrtKey='" + ctCurrPrtn + "' and ", "") + " RmCd= IsNull(Rm.RmBaseCd, '')), '') ), 1) " + _
                "else 1 end) "
  
  ' ****** Sachin 2.13.0 - 05/06/2006 ******
  ' ***** Manali 3.03 - 23/06/08 - Pd Mod - 'L' added
    wQwChk = " IsNull((Case when (IsNull(RmCtg,'') in ('G', 'P', 'S', 'L') And IsNull(RmQw,'')='W') then (Select RmPr.RmQw From RmMst RmPr " + _
              "   Where " + IIF(gs_Partition = ctCurrPrtn, " RmPr.RmPrtKey='" + ctCurrPrtn + "' and ", "") + " RmPr.RmCd=(Select PValue from param " + _
              "       where PTyp= 'RMCTG' and PMCd= IsNull(Rm.RmCtg,''))) Else RmQw End),'')  as qQw," + _
              " IsNull((Case when (IsNull(RmCtg,'') in ('G', 'P', 'S', 'L') And IsNull(RmQwCst,'')='W') then (Select RmPr.RmQwCst From RmMst RmPr " + _
              "   Where " + IIF(gs_Partition = ctCurrPrtn, " RmPr.RmPrtKey='" + ctCurrPrtn + "' and ", "") + " RmPr.RmCd=(Select PValue from param " + _
              "       where PTyp= 'RMCTG' and PMCd= IsNull(Rm.RmCtg,''))) Else RmQwCst End),'') as qQwCst "
  ' ****** Sachin 2.13.0 - 05/06/2006 ******
  
  '*** Selection condition for normal or rejected pieces ('N'- Normal, 'R'-Rejected) ***
  '### wSelCnd = IIf(adc("UoBagPcs") = "N", " and (BRjQty= 0 and BRjPtQty= 0) ", " and (BRjQty<> 0 or BRjPtQty<> 0) ")
  'wSelCnd = IIf(adc("UoBagPcs") = "N", " and (BQty<> 0 or BPtQty<> 0 or (BQty= 0 and BPtQty= 0 and BRjQty= 0 and BRjPtQty= 0 and Round(BGrWt, 3)<> 0)) ", " and (BRjQty<> 0 or BRjPtQty<> 0) ")
  '### wSelCnd = IIf(ADC("UoBagPcs") = "N", " and (BQty<> 0 or BPtQty<> 0) ", " and (BRjQty<> 0 or BRjPtQty<> 0) ")   '(This condition should have given the same result as above but it does not give)

'*** Jay 2.14Next *** [remove N/R pcs scop]
'  wQty = IIF(adc("UoBagPcs") = "N", " Bv.TdBQty as BQty", " Bv.TdRjQty as BQty")
'  wPtQty = IIF(adc("UoBagPcs") = "N", " Bv.TdPtQty as BPtQty", " Bv.TdRjPtQty as BPtQty")
'  If adc("UoBagPcs") = "N" Then Rep.TxtHead.SetText "Bag Movement RM (Normal)"
'  If adc("UoBagPcs") = "R" Then Rep.TxtHead.SetText "Bag Movement RM (Rejected)"
  Rep.TxtHead.SetText "Bag Movement RM (Normal)"
'*** Jay 2.14Next *** [remove N/R pcs scop]

  '*** Selection condition for normal or rejected pieces ('N'- Normal, 'R'-Rejected) ***
  
  ' **** Use Below Condition if Empty bag that is moved is also to be shown
  'wSelCnd = IIF(adc("UoBagPcs") = "N", " (BQty<> 0 or BPtQty<> 0 or (BQty= 0 and BPtQty= 0 and BRjQty= 0 and BRjPtQty= 0 and Round(BGrWt, 3)<> 0)) ", " (BRjQty<> 0 or BRjPtQty<> 0) ")
  
  '************** sachin **********
  ''wSelCnd = IIF(adc("UoBagPcs") = "N", " (BQty<> 0 or BPtQty<> 0 or (BQty= 0 and BPtQty= 0 and BRjQty= 0 and BRjPtQty= 0 and Round(BGrWt, 3)<> 0)) ", " (BRjQty<> 0 or BRjPtQty<> 0) ")
  '************** sachin **********

  If adc("UoDtFr") = "01/01/80" Then
    Rep.TxtVchDtFr.Suppress = True
    Rep.LblVchDtFr.Suppress = True
    Rep.LblFrVchDtCol.Suppress = True
  Else
      Rep.TxtVchDtFr.SetText adc("UoDtFr")
  End If
  
  If adc("UoDtTo") = "01/01/80" Then
    Rep.TxtVchDtTo.Suppress = True
    Rep.LblVchDtTo.Suppress = True
    Rep.LblToVchDtCol.Suppress = True
  Else
    Rep.TxtVchDtTo.SetText adc("UoDtTo")
  End If

  Call CmpStr
  
  wrepcnd = adc.RepCond
  'wCnd = IIF(Trim(wrepcnd) <> "", " and " + wrepcnd, "")
  wCnd = IIF(wrepcnd <> "", wrepcnd, "")
  If OptLoc(0).Value = True Then
      If adc("UoIssLocFr") <> "" Then wCnd = wCnd + " and Bv.TdFrBLoc >= '" + adc("UoIssLocFr") + "'"
      If adc("UoIssLocTo") <> "" Then wCnd = wCnd + " and Bv.TdFrBLoc <= '" + adc("UoIssLocTo") + "'"
      If adc("UoIssLocSel") <> "" Then wCnd = wCnd + " and Bv.TdFrBLoc In (" + adc("UoIssLocSel") + ")"
      If adc("UoRecLocFr") <> "" Then wCnd = wCnd + " and Bv.TdToBLoc >= '" + adc("UoRecLocFr") + "'"
      If adc("UoRecLocTo") <> "" Then wCnd = wCnd + " and Bv.TdToBLoc <= '" + adc("UoRecLocTo") + "'"
      If adc("UoRecLocSel") <> "" Then wCnd = wCnd + " and Bv.TdToBLoc In (" + adc("UoRecLocSel") + ")"
    ElseIf OptLoc(1).Value = True Then
      If adc("UoIssLocFr") <> "" And adc("UoIssLocTo") <> "" Then wCnd = wCnd + " and Bv.TdFrBLoc in (select LocCd from Loc where LocPrcs >= '" + adc("UoIssLocFr") + "'"
      If adc("UoIssLocFr") <> "" And adc("UoIssLocTo") = "" Then wCnd = wCnd + " and Bv.TdFrBLoc in (select LocCd from Loc where LocPrcs >= '" + adc("UoIssLocFr") + "')"
      If adc("UoIssLocTo") <> "" And adc("UoIssLocFr") <> "" Then wCnd = wCnd + " and LocPrcs <= '" + adc("UoIssLocTo") + "')"
      If adc("UoIssLocTo") <> "" And adc("UoIssLocFr") = "" Then wCnd = wCnd + " and Bv.TdFrBLoc in (select LocCd from Loc where LocPrcs <= '" + adc("UoIssLocTo") + "')"
      If adc("UoIssLocSel") <> "" Then wCnd = wCnd + " and Bv.TdFrBLoc in (select LocCd from Loc where LocPrcs In (" + adc("UoIssLocSel") + "))"
      If adc("UoRecLocFr") <> "" And adc("UoRecLocTo") <> "" Then wCnd = wCnd + " and Bv.TdToBLoc in (select LocCd from Loc where LocPrcs >= '" + adc("UoRecLocFr") + "'"
      If adc("UoRecLocFr") <> "" And adc("UoRecLocTo") = "" Then wCnd = wCnd + " and Bv.TdToBLoc in (select LocCd from Loc where LocPrcs >= '" + adc("UoRecLocFr") + "')"
      If adc("UoRecLocTo") <> "" And adc("UoRecLocFr") <> "" Then wCnd = wCnd + " and LocPrcs <= '" + adc("UoRecLocTo") + "')"
      If adc("UoRecLocTo") <> "" And adc("UoRecLocFr") = "" Then wCnd = wCnd + " and Bv.TdToBLoc in (select LocCd from Loc where LocPrcs <= '" + adc("UoRecLocTo") + "')"
      If adc("UoRecLocSel") <> "" Then wCnd = wCnd + " and Bv.TdToBLoc in (select LocCd from Loc where LocPrcs In (" + adc("UoRecLocSel") + "))"
    End If
    
  
  '***************** urmi ******************
  wBagClsCnd = IIF(adc("UoYN1") = "Y", " and BCls= 'N' ", IIF(adc("UoYN1") = "N", " and BCls= 'Y' ", ""))
  '***************** urmi ******************
  
  '"TdRmWt * (Case when RmCtg in ('G', 'P') then RmPurityWt" + wPurityWt + " else 0 end) as qGldWt, "

'*** Jay 2.13(CT) ***
   ws_BagJoin = " Join Param On PTyp= 'BCHR' and PMCd= BChr and PSCd= '' " + _
                    IIF(adc("UoBagTyp") <> "", " and PValue= '" + adc("UoBagTyp") + "' ", " ") + _
                    IIF(adc("UOBCHRFR") <> "", " and PMCd= '" + adc("UOBCHRFR") + "' ", " ")
'' **** Zubin 211 **** '
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
'' **** Zubin 211 **** '
'*** Jay 2.13(CT) ***
  'kbs.136-selects only Production=Y records from txnd, it will execute only if PrdBag is Yes
  If adc("wPrdBagsYN") = "Y" Then
    ws_PrdBag = "and exists (Select 'a' from Txnd ForPrd where ForPrd.TdByy = Bv.TdByy and ForPrd.TdBChr = Bv.TdBChr " + _
              "and ForPrd.TdBNo = Bv.TdBNo and ForPrd.TdPrdYn = 'Y' and ForPrd.TdByLoc = Bv.TdFrBLoc and " + _
              "ForPrd.TdYyKey < Bv.TdYyKey and ForPrd.TdYyKey > isnull((Select max(TdYyKey) from Txnd BagIssueRecTxnd where " + _
              "BagIssueRecTxnd.TdByy = Bv.TdByy And BagIssueRecTxnd.TdBChr = Bv.TdBChr And BagIssueRecTxnd.TdBNo = Bv.TdBNo and " + _
              "BagIssueRecTxnd.TdToBLoc = Bv.TdFrBLoc and BagIssueRecTxnd.TdTc = 'BV' and BagIssueRecTxnd.TdKey < Bv.TdKey),0)) "
  End If
  ' Zubin 212
  '****** Sachin 2.13.0 - added RmQWCst
  ' ***** Manali 3.03 - 23/06/08 - Pd Mod - 'L' added
  'sv.81 BvCoCd,BvTc, BvYy, BvChr, BvNo, BvSr added
  '6.1
  wSqlStrg = " Set DateFormat " + DtFmtStr() + _
              " Select " + grpflds0 + ", " + _
              " BYy, BChr, BNo, DmTcTyp, BOdDmCd, BOdDmSz, BLoc, RmCtg, RmSCtg, IsNull(Td.TdRmCd, '') as TdRmCd, " + _
              " IsNull(Td.TdLotNo, '') as TdLotNo, IsNull(Td.TdRmSz, '') as TdRmSz, IsNull(Td.TdRmSz2, '') as TdRmSz2, " + _
              " IsNull(Td.TdRmSz3, '') as TdRmSz3,IsNull(Td.TdRmStkRt, 0) as TdRmStkRt, " + _
              " (Select PValue3 from Param where PTyp= 'RMSCTG' and PMCd= IsNull(RmCtg, '') and " + _
              " PSCd= IsNull(RmSCtg, '')) as qPtrYN, " + _
              " IsNull(Td.TdRmQty, 0) as TdRmQty, IsNull(Td.TdRmWt, 0) as TdRmWt, " + _
              " (Case when IsNull(Td.TdFrRmLoc, '')='W' and IsNull(Td.TdFrRmDc, '')='D' then 1 " + _
              " when IsNull(Td.TdFrRmLoc, '')='W' and IsNull(Td.TdFrRmDc, '')='C' then -1 " + _
              " when IsNull(Td.TdToRmLoc, '')='W' and IsNull(Td.TdToRmDc, '')='D' then 1 " + _
              " when IsNull(Td.TdToRmLoc, '')='W' and IsNull(Td.TdToRmDc, '')='C' then -1 Else 0 end) as qMul, " + _
              " IsNull(Td.TdRmWt, 0) / (Case when IsNull(RmCtg, '') in ('D', 'C') then 5 else 1 end) as qGmWt, " + _
              " IsNull(Td.TdRmWt, 0) * (Case when IsNull(RmCtg, '') in ('G', 'P', 'S', 'L') then IsNull(RmPurityWt, 0) " + _
              "   when IsNull(RmCtg, '')= 'X' then IsNull(RmBasePurityWt, 0) else 0 end) " + wPurityWt + " as qGldWt, " + _
              " Bv.TdBQty as BQty,  Bv.TdPtQty as BPtQty, TDt, IsNull(Td.TdTc, '') as TdTc, IsNull(Td.TdYy, '') as TdYy, " + _
              " IsNull(Td.TdChr, '') as TdChr, " + _
              " IsNull(Td.TdNo, 0) as TdNo, IsNull(Td.TdSr, 0) as TdSr, IsNull(Td.TdSrNo, 0) as TdSrNo, " + wTdRate + ", " + wQwChk + "," + _
              IIF(OptLoc(0).Value = True, "Bv.TdFrBLoc", "(select LocPrcs from Loc where LocCd =Bv.TdFrBLoc and LocCoCd =Bv.TdCoCd)") + " as TdFrBLoc, " + _
              IIF(OptLoc(0).Value = True, "Bv.TdToBLoc", "(select LocPrcs from Loc where LocCd =Bv.TdToBLoc and LocCoCd =Bv.TdCoCd)") + " as TdToBLoc, " + _
              "RSCtg.PValue1 as qStkbyRt, Bv.TdIdNo " + _
              ",IsNull(Bv.TdCoCd, '') as BvCoCd, IsNull(Bv.TdTc, '') as BvTc, IsNull(Bv.TdYy, '') as BvYy, " + _
              "IsNull(Bv.TdChr, '') as BvChr, IsNull(Bv.TdNo, 0) as BvNo, IsNull(Bv.TdSr, 0) as BvSr "

              '" Bv.TdFrBLoc as TdFrBLoc, Bv.TdToBLoc as TdToBLoc, RSCtg.PValue1 as qStkbyRt "

' **** Zubin 211 () **** '
  '*** (Jen 2.12 corrected join of Txnd Bv)
  '****** Sachin 3.02 - Id fields in Joins
  'kbs.136- ws_PrdBag added
  wSqlStrg = wSqlStrg + " From Bag " + ws_BagJoin + _
              " Join OrdDsg on OdIdNo=BOdIdNo and OdCoCd= BCoCd and OdTc= BOdTc and OdYy= BOdYy and " + _
              "     OdChr= BOdChr and OdNo= BOdNo and OdSr= BOdSr " + IIF(gs_Partition = ctCurrPrtn, " And OdPrtKey=BPrtKey ", "") + _
              " Join OrdMst On OmIdNo=OdOmIdNo and OmCoCd= OdCoCd and  OmTc= OdTc and OmYy= OdYy and OmChr= OdChr and OmNo= OdNo " + IIF(gs_Partition = ctCurrPrtn, " And OmPrtKey=OdPrtKey ", "") + _
              " /*Left Outer */Join Txnd Td on Td.TdBIdNo=BIdNo and Td.TdCoCd= BCoCd and Td.TdBYy=BYy and Td.TdBChr=BChr and " + _
              "     Td.TdBNo=BNo and Td.TdSrNo<> 0 " + IIF(gs_Partition = ctCurrPrtn, " And Td.TdPrtKey=BPrtKey ", "") + _
              " Join Txnd Bv on Bv.TdBIdNo=BIdNo and Bv.TdCoCd= BCoCd and Bv.TdBYy=BYy and Bv.TdBChr=BChr and Bv.TdBNo=BNo and " + IIF(gs_Partition = ctCurrPrtn, " Bv.TdPrtKey=BPrtKey And ", "") + _
              "   Bv.TdSrNo= 0 and (Select PValue from Param where PTyp= 'TC' and PMCd= Bv.TdTc)= 'BV' and " + _
              "  ( " + _
              "     (IsNull(Td.TdYy, '')<= Bv.TdPrevYy And " + _
              "     (IsNull(Td.TdYy, '')< Bv.TdPrevYy Or IsNull(Td.TdKey, 0)<= Bv.TdPrevKey) ) " + _
              "    /*OR (Bv.TdPrevYy= '' and Bv.TdPrevKey= 0) */ " + _
              "  ) " + _
              " Join Txn on TIdNo=Bv.TdTIdNo and TCoCd= Bv.TdCoCd and TTc= Bv.TdTc and TYy= Bv.TdYy and " + _
              "   TChr= Bv.TdChr and TNo= Bv.TdNo " + IIF(gs_Partition = ctCurrPrtn, " And TPrtKey=Bv.TdPrtKey ", "") + _
              " Join DsgMst on DmIdNo=BDmIdNo and DmTcTyp= BOdDmTcTyp and DmCd= BOdDmCd and DmSz= '' " + IIF(gs_Partition = ctCurrPrtn, " And DmPrtKey=BPrtKey ", "") + _
              " Join RmMst Rm on RmCd= IsNull(Td.TdRmCd, '') " + IIF(gs_Partition = ctCurrPrtn, " And RmPrtKey=Td.TdPrtKey ", "") + _
              " Join Param RSCtg On RSCtg.PTyp='RMSCTG' And RSCtg.PMCd=RmCtg And RSCtg.PSCd=RmSCtg " + _
              " where " + IIF(gs_Partition = ctCurrPrtn, " BPrtKey='" + ctCurrPrtn + "' And ", "") + _
              wSelCnd + wCnd + IIF(wAddlCnd <> "", IIF(wCnd <> "", " And ", "") + wAddlCnd, "") + wBagClsCnd + _
              " and Bv.TdAck = 'Y'" + ws_PrdBag '*** Jay 2.14 *** (Added Check for Ack Bags Only.)
  '*** (Jen 2.12 corrected join of Txnd Bv)
  
  '*** (Bef 2.11)
  'wSqlStrg = wSqlStrg + " From Bag " + ws_BagJoin + _
  '            "Join OrdDsg on OdCoCd= BCoCd and OdTc= BOdTc and OdYy= BOdYy and " + _
  '            "OdChr= BOdChr and OdNo= BOdNo and OdSr= BOdSr Join OrdMst On OmCoCd= OdCoCd and " + _
  '            "OmTc= OdTc and OmYy= OdYy and OmChr= OdChr and OmNo= OdNo " + _
  '            "/*Left Outer */Join Txnd Td on Td.TdCoCd= BCoCd and Td.TdBYy=BYy and Td.TdBChr=BChr and " + _
  '            "   Td.TdBNo=BNo and Td.TdSrNo<> 0 " + _
  '            "Join Txnd Bv on Bv.TdCoCd= BCoCd and Bv.TdBYy=BYy and Bv.TdBChr=BChr and Bv.TdBNo=BNo and " + _
  '            "   Bv.TdSrNo= 0 and (Select PValue from Param where PTyp= 'TC' and PMCd= Bv.TdTc)= 'BV' and " + _
  '            "  ( " + _
  '            "     (IsNull(Td.TdYy, 0)<= Bv.TdPrevYy And IsNull(Td.TdKey, 0)<= Bv.TdPrevKey) " + _
  '            "    /*OR (Bv.TdPrevYy= '' and Bv.TdPrevKey= 0) */ " + _
  '            "  ) " + _
  '           "Join Txn on TCoCd= Bv.TdCoCd and TTc= Bv.TdTc and TYy= Bv.TdYy and " + _
  '            "   TChr= Bv.TdChr and TNo= Bv.TdNo " + _
  '            "Join DsgMst on DmTcTyp= 'DM' and DmCd= BOdDmCd and DmSz= '' " + _
  '            "/*Left Outer */Join RmMst Rm on RmCd= IsNull(Td.TdRmCd, '') where " + _
  '            wSelCnd + wCnd + wBagClsCnd
  '*** (Bef 2.11)
  
  adc.RepRecSource = wSqlStrg + ctMaxDopOpt  '****** Sachin 3.01 ctMaxDopOpt added

  If adc("UoYN") = "N" Then
    Rep.BoxHd.Suppress = True
    Rep.BoxDet.Suppress = True
    Rep.BoxFt.Suppress = True
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
'  If adc("UoBagTyp") <> "" Then Rep.TxtBagTypAndChr.SetText IIF(adc("UoBagTyp") = "Y", "(Flute Bags", "(Primary Bags") & IIF(adc("UoBChrFr") <> "", ",  Chr = " & adc("UoBChrFr"), "") & ")"
'  '***(Jen 2.11)
  '*** Jay 2.13(CT) ***
  
  '***Geeta**Emr208 **Picture
'  If adc("UoYN2") = "N" Then
'    Rep.PicSec.Suppress = True
'  Else
'    Rep.PicSec.Suppress = False
'  End If
  ' Zubin 212
  If Not moCn.RecSeek("select 'x' from Head where HCd= '" + ctSelfCmCd + "' and HCoCd= '" + adc("UoCoCdFr") + "' and HLotNoYN= 'Y' ") Then
    Rep.HLotNo.Suppress = True: Rep.FldLotNo.Suppress = True
  End If
  ' Zubin 212
  
  
  'KC.3- StkRt-when Show Stock Rate=YES then Stock Rate is visible in report
  'if it is NO then Stock Rate is invisible in report
 
  If adc("UoYN3") = "Y" Then
    Rep.hTdRmStkRt.SetText "Stock Rt"
    Rep.FldRmStkRt.Suppress = False
  Else
    Rep.hTdRmStkRt.SetText ""
    Rep.FldRmStkRt.Suppress = True
  End If
  '578.19 calling function to change Rmcode font size
  GRP_REP.SetFormula Rep, "wRmCdFontSz", GetRmcdFontSz("{rdo.TdRmCd}", 8)
    
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
  Case Is = 0               ' Tag = ShowBoxDet
    If .Value = Checked Then adc("UoYN") = "Y" Else adc("UoYN") = "N"
  'StkRt-To select this option(Show Stock rate) either Seo user can give their password or
  ' can enter SEOPWD password otherwise this option can't not select
  Case Is = 1
    If .Value = Checked Then
       If Trim(TxtPwd.text) <> moCn.GetFldVal("select IsNull ((select PValue from Param Join UsrAccess on UaUsrCd= PMCd " + _
            "where PTyp= 'USR' and PMCd= '" + gs_UsrCd + "'  and UaMnuCd= '" + adc.MenuCd + "' and " + _
            "UaSeoYn= 'Y' and PValue<> '' and PValue= '" + Trim(TxtPwd.text) + "' ), " + _
            "Isnull ((Select PValue1 from Param where " + _
            "Ptyp= 'SEOPWD' and PmCd= 'SEOPWD' ), ''))") Then _
            .Value = Unchecked
    End If
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
  gb_CoCdFor = True
  Set adc.FirNKeyCtl = IIF(ws_MultiCoMod = True, adc("UoCoCdFr"), adc("UoIssLocFr"))
  adc("UOCOCDFR").Enabled = ws_MultiCoMod
  '****** Sachin 2.14.0 - [23-11-2006] - [Multi-Company] ******
  
  Call InitProp(Me)
  
  '*** Parameters Used For Setting Properties Of UsrOpt Fields ***
  gs_LocTyp = "'P'"
  gb_RmCtgFor = False
  gb_RmZ = False
  gs_DmTcTyp = "DM"
  'gs_TxnTcTyp = "BV"
  gs_OmTcTyp = "SO"
  gs_PTyp = "WIPVAL"
  'gs_RmGrp = "ORDDET"
  'gs_LabGrp = "ORDDET"
  gs_Tbl = "Param"
  '*** Parameters Used For Setting Properties Of UsrOpt Fields ***

  Call SetGroupSort
End Sub
Private Sub ATXT_FldChange(Index As Integer, ByVal IdName As String, Cancel As Boolean, ErrMsg As String, ByVal pv_NewValue As Variant)
  gs_PTyp = "WIPVAL"
  gs_Tbl = "Param"
  gb_CoCdFor = True '****** Sachin 2.14.0 - [Multi-Company]
    
  Select Case UCase(IdName)
  ' urmi 2.06 Pwd for Val
  Case Is = UCase("UoPmCdFr"), UCase("UoPmCdTo")     '4.1.4
    '*** Jay 2.14 *** (SEO Change)
    If UCase(IdName) = UCase("UoPmCdFr") And pv_NewValue = "" Then adc("UoPMCdTo") = "" '4.1.4
    
    If pv_NewValue <> "" Then
       If Trim(TxtPwd.text) <> moCn.GetFldVal("select IsNull ((select PValue from Param Join UsrAccess on UaUsrCd= PMCd " + _
            "where PTyp= 'USR' and PMCd= '" + gs_UsrCd + "'  and UaMnuCd= '" + adc.MenuCd + "' and " + _
            "UaSeoYn= 'Y' and PValue<> '' and PValue= '" + Trim(TxtPwd.text) + "' ), " + _
            "Isnull ((Select PValue1 from Param where " + _
            "Ptyp= 'SEOPWD' and PmCd= 'SEOPWD' ), ''))") Then _
            Cancel = True: ErrMsg = "Enter SEO Password to View Values ": Exit Sub
    End If
    
  '****** Sachin 2.14.0 - [24-11-2006] - [Multi-Company] ******
  Case Is = UCase("UoCoCdFr")
    If Not moCn.RecSeek("Select 'x' From Head Where HCd='" + ctSelfCmCd + "' And HCoCd='" + pv_NewValue + "'") Then _
            Cancel = True: ErrMsg = "Invalid Company Code": Exit Sub
    Me.Caption = GetFrmCaption(adc("UOCOCDFR"), "Bag Movement RM")
  '****** Sachin 2.14.0 - [24-11-2006] - [Multi-Company] ******
  
    ' Bef 2.14
    'If pv_NewValue <> "" And Trim(TxtPwd.text) <> moCn.GetFldVal("Select PValue1 from Param where " + _
                                   "Ptyp = 'SEOPWD' and PmCd = 'SEOPWD'") Then
    '  Cancel = True: DispMsg "Enter SEO Password to View Values", etError
    'End If
    ' Bef 2.14
    '*** Jay 2.14 *** (SEO Change)
  'kbs.136- checking valid option
  Case Is = UCase("wPrdBagsYN")
    Cancel = Not moCn.RecSeek(" Select PMCd from Param where PTyp= 'YN' and PMCd = '" + pv_NewValue + "' ")
    If Cancel = True Then ErrMsg = "Invalid Production Bags Option. Enter(Y/N)": Exit Sub
        
  End Select

  Call SetProp(Me, IdName, Valid, Cancel, ErrMsg)
End Sub
Private Sub ATXT_FldWhen(Index As Integer, ByVal IdName As String, Cancel As Boolean, ErrMsg As String)
  gs_PTyp = "WIPVAL"
  gs_Tbl = "Param"
  gb_CoCdFor = True '****** Sachin 2.14.0 - [Multi-Company]
  
  '*** (Jen 2.11)
  Select Case UCase(IdName)
    Case Is = UCase("UoBChrFr")
      If adc.Mode = xNorm Then
        '*** Jay 2.13(CT) ***
        If adc("UoBagTyp") = "" Then Cancel = True: ErrMsg = "First Specify Bag Type": Exit Sub
        HlpList.PMCd "BCHR", "'" + adc("UoBagTyp") + "'": Exit Sub
        '*** Jay 2.13(CT) ***
      End If
    '*** Jay 3.01 **** [LotNo Err]
    Case Is = UCase("UoLotNoFr"), UCase("UoLotNoTo"), UCase("UoLotNoSel")
      gs_RmLocTyp = "W":      gs_Loc = "W"
    '*** Jay 3.01 **** [LotNo Err]
    
    Case Is = UCase("UoPMCdTo")     '4.1.4
      If adc.Mode = xNorm Then _
        If UCase(adc("UoPMCdFr")) <> UCase("STKVAL") Then Cancel = True: ErrMsg = "Can Enter Only if Value Type is 'STKVAL'": Exit Sub
    
    Case Is = UCase("UOISSLOCFR"), UCase("UOISSLOCTO"), UCase("UOISSLOCSEL"), UCase("UORECLOCFR"), UCase("UORECLOCTO"), UCase("UORECLOCSEL")
        If adc("UOLOCTYPFR") = "R" Then gs_LocTyp = "'R'"
        If adc("UOLOCTYPFR") = "P" Then gs_LocTyp = "'P'"
    'kbs.136-calling help
    Case Is = UCase("wPrdBagsYN")
      Call HlpList.PMCd("YN")

'    Case Is = UCase("UOISSLOCTO")
'                   If adc("UOLOCTYPFR") = "R" Then gs_LocTyp = "'R'"
'        If adc("UOLOCTYPFR") = "P" Then gs_LocTyp = "'P'"
        
'        Case Is = UCase("UOISSLOCSEL")
'        If OptLoc(0).Value = True Then
'        Call HlpList.Loc("", "'P'")
'        Else
'        Call HlpList.Loc("", "'R'")
'        Exit Sub
'       End If
'
'
'        Case Is = UCase("UORECLOCFR")
'        If OptLoc(0).Value = True Then
'        Call HlpList.Loc("", "'P'")
'
'        Else
'        Call HlpList.Loc("", "'R'")
'        Exit Sub
'        End If
'
'
'        Case Is = UCase("UORECLOCTO")
'        If OptLoc(0).Value = True Then
'        Call HlpList.Loc("", "'P'")
'        Else
'        Call HlpList.Loc("", "'R'")
'       Exit Sub
'        End If
'
'        Case Is = UCase("UORECLOCSEL")
'        If OptLoc(0).Value = True Then
'        Call HlpList.Loc("", "'P'")
'        Else
'        HlpList.MultiSelect = True
'        Call HlpList.Loc("", "'R'")
'        Exit Sub
'       End If
    
  End Select
  '*** (Jen 2.11)
  
  SetProp Me, IdName, When
End Sub
Private Sub atxt_ReCalculate(Index As Integer, ByVal IdName As String, ByVal OldVal As Variant, ByVal RowNum As Integer, ByVal SourceName As String, ByVal SourceColName As String)
' **** Zubin 211 **** '
  Select Case UCase(IdName)
    Case Is = UCase("UOBCHRFR")
      'If ADC("UoBagTyp") = "N" Then
      adc("UOBCHRFR") = ""
  End Select
' **** Zubin 211 **** '
End Sub

' ****** Manali 3.5.0 - 06/11/08 - Radio Buttons Added
Private Sub OptBag_Click(Index As Integer)
Select Case Index
Case Is = 0
  adc("UoYN1") = "Y"
Case Is = 1
  adc("UoYN1") = "N"
Case Is = 2
  adc("UoYN1") = ""
End Select
End Sub
' ****** Manali 3.5.0 - 06/11/08 - Radio Buttons Added

' ****** Manali 3.5.0 - 06/11/08 - Radio Buttons Added
Private Sub OptBag_GotFocus(Index As Integer)
  DispMsg FraOptBag.ToolTipText, etInfo
End Sub
' ****** Manali 3.5.0 - 06/11/08 - Radio Buttons Added

Private Sub OptLoc_GotFocus(Index As Integer)
  DispMsg FraLoc.ToolTipText, etInfo
Select Case Index
Case Is = 0
  adc("UOISSLOCFR") = "":   adc("UOISSLOCTO") = "":   adc("UOISSLOCsel") = ""
  adc("UORECLOCFR") = "":   adc("UORECLOCTO") = "":   adc("UORECLOCSEL") = ""
  adc("UoLocTypFr") = "P"
Case Is = 1
  adc("UOISSLOCFR") = "":   adc("UOISSLOCTO") = "":   adc("UOISSLOCsel") = ""
  adc("UORECLOCFR") = "":   adc("UORECLOCTO") = "":   adc("UORECLOCSEL") = ""
  adc("UoLocTypFr") = "R"
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
  
  adc("UoCoCdFr") = gs_CoCd
  'Call SetGroupSort
  
  If adc("UoDtFr") = "" Then adc("UoDtFr") = moCn.SrvrDate
  If adc("UoDtTo") = "" Then adc("UoDtTo") = moCn.SrvrDate
  adc("UoPmCdFr") = "": adc("UoPmCdTo") = ""    '4.1.4

  '****** Sachin 2.14.0 - [24-11-2006] - [Multi-Company] ******
  Me.Caption = GetFrmCaption(adc("UOCOCDFR"), "Bag Movement RM")
  '****** Sachin 2.14.0 - [24-11-2006] - [Multi-Company] ******

  '**Geeta*Emr208*Default Show picture is always no
  'adc("UoYN2") = "N"
  
  '*** Jay 2.13(CT) ***
  Dim wHeadChk As MwfLib.MDORowSet
  Set wHeadChk = moCn.OpenRes("Select hFluteBagYN, hCompBagsYn from Head where HCoCd= '" + ctSelfCoCd + "' and " + _
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

  ' ****** Manali 3.5.0 - 06/11/08 - Radio Buttons Added
  If adc("UoYN1") = "Y" Then
    OptBag(0).Value = True
  ElseIf adc("UoYN1") = "N" Then
    OptBag(1).Value = True
  ElseIf adc("UoYN1") = "" Then
    OptBag(2).Value = True
  End If
  ' ****** Manali 3.5.0 - 06/11/08 - Radio Buttons Added
  
  If adc("UOLOCTYPFR") = "P" Then
    OptLoc(0).Value = True
  ElseIf adc("UOLOCTYPFR") = "R" Then
    OptLoc(1).Value = True
  End If
 
  ' ***** Manali 3.5.0 - 24/11/08 - Check Boxes Added
  If adc("UoYN") = "" Then adc("UoYN") = "N"

  If adc("UoYN") = "Y" Then ChkBoxArr(0).Value = Checked Else ChkBoxArr(0).Value = Unchecked
  ' ****** Manali 3.5.0 - 24/11/08 - Check Boxes Added
  
  '4.1.0.0
  adc("UOTTCFR") = "BV"
  adc("UOTTCTO") = "BV"
  
  'StkRt-Set Show StockRate option is No
  adc("UoYN3") = "N"
  'kbs.136- default value for Only PrdBag is No
  adc("wPrdBagsYN") = "N"

End Sub
Private Sub DispCoNm()
'Rep.wCoCd.SetText gs_CoNm
  Dim ws_HName As String
  ws_HName = moCn.GetFldVal("Select HName from Head where HCoCd ='" + adc("UoCoCdFr") + "' and " + _
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
'*** Jay 2.14 *** (SEO Change)
Private Sub TxtPwd_GotFocus()
  DispMsg "Enter SEO Password", etInfo
End Sub
'*** Jay 2.14 *** (SEO Change)

'urmi 2.06
Private Sub TxtPwd_Validate(Cancel As Boolean)
'*** Jay 2.14 *** (SEO Change)
  Dim ws_VldPwd As String
  ws_VldPwd = moCn.GetFldVal("select isnull ((select PValue from Param Join UsrAccess on UaUsrCd = PMCd " + _
                   "where PTyp= 'USR' and PMCd= '" + gs_UsrCd + "' and UaMnuCd = '" + adc.MenuCd + "' and " + _
                   "UaSeoYn= 'Y' and PValue= '" + Trim(TxtPwd.text) + "'), " + _
                   "(Select PValue1 from Param where Ptyp= 'SEOPWD' and PmCd= 'SEOPWD' )) ")
    
  If Trim(TxtPwd.text) = "" And ws_VldPwd <> "" Then adc("UoPMCdFr") = "": Exit Sub
  If Trim(TxtPwd.text) <> "" And Trim(TxtPwd.text) <> ws_VldPwd Then
      Cancel = True:  DispMsg "Invalid Password / User Does Not Have Seo Rights ", etError
  End If
' Bef 2.14
'  If Trim(TxtPwd.text) = "" Then adc("UoPmCDFr") = ""
'  If Trim(TxtPwd.text) = "" Or Trim(TxtPwd.text) = moCn.GetFldVal("Select PValue1 from Param where " + _
'                         "Ptyp = 'SEOPWD' and PmCd = 'SEOPWD'") Then
'    Exit Sub
'  Else
'    Cancel = True
'    DispMsg "Invalid Password", etError
'  End If
' Bef 2.14
'*** Jay 2.14 *** (SEO Change)
End Sub

Private Function RmRate(ByVal mRtFld As String, ByVal mSzFld As String, ByVal mQw As String) As String
  Dim wPureGldCd As String, wPurePlCd As String, wPureSlCd As String
  Dim wGRt As String, wPRt As String, wSRt As String, wDCXARt As String
  Dim wPurePdCd As String, wLRt As String ' *** Manali 3.03 - 23/06/08 - Pd Mod - 'L'
  
  wPureGldCd = moCn.GetFldVal("Select PValue from param where PTyp= 'RMCTG' and PMCd= 'G'")
  wPurePlCd = moCn.GetFldVal("Select PValue from param where PTyp= 'RMCTG' and PMCd= 'P'")
  wPureSlCd = moCn.GetFldVal("Select PValue from param where PTyp= 'RMCTG' and PMCd= 'S'")
  ' ***** Manali 3.03 - 23/06/08 - Pd Mod - 'L'
  wPurePdCd = moCn.GetFldVal("Select PValue from param where PTyp= 'RMCTG' and PMCd= 'L'")
  
  ' Zubin 211 (Patch 03)
  'Prev 01/02/05 Geeta  2.10
  ' Goes Wrong when RmCtg='G' and RmSCtg='GX'
  ' And RmCtg='G' and RmSCtg='-' for RmCd='G24' PureMetal Because RmSCtg does not Match
  
'  wGRt = " isnull((select " + mRtFld + " from RmRt where RrTcTyp= 'RM' and RrCmCtg= 'C' and " + _
'         "RrCtg= RmCtg and RrSCtg= RmSCtg and RrCd= '" + wPureGldCd + "' and " + _
'         "RrFrLn <= " + mSzFld + " And RrToLn >= " + mSzFld + " and RrCmCd = '" + ctValueCmCd + "' and " + _
'         "RrDmCtg= '' and RrLabMCd= '') , 0) "
'  wPRt = " isnull((select " + mRtFld + " from RmRt where RrTcTyp= 'RM' and RrCmCtg= 'C' and " + _
'         "RrCtg= RmCtg and RrSCtg= RmSCtg and RrCd= '" + wPurePlCd + "' and " + _
'         "RrFrLn <= " + mSzFld + " And RrToLn >= " + mSzFld + " and RrCmCd = '" + ctValueCmCd + "' and " + _
'         "RrDmCtg= '' and RrLabMCd= '') , 0) "
'  wSRt = " isnull((select " + mRtFld + " from RmRt where RrTcTyp= 'RM' and RrCmCtg= 'C' and " + _
'         "RrCtg= RmCtg and RrSCtg= RmSCtg and RrCd= '" + wPureSlCd + "' and " + _
'         "RrFrLn <= " + mSzFld + " And RrToLn >= " + mSzFld + " and RrCmCd = '" + ctValueCmCd + "' and " + _
'         "RrDmCtg= '' and RrLabMCd= '') , 0) "

  ' Zubin 213
  Dim ws_BaseCurCd As String
  ws_BaseCurCd = moCn.GetFldVal("Select HBaseCurCd from Head where HCd= '" + ctSelfCmCd + "' and " + _
                  "HCoCd= '" + adc("UoCoCdFr") + "'")

  ' Zubin 213 (ws_BaseCurCd cnd added)
  wGRt = " isnull((select " + mRtFld + " from RmRt where " + IIF(gs_Partition = ctCurrPrtn, " RrPrtKey='" + ctCurrPrtn + "' and ", "") + " RrTcTyp= 'RM' and RrCmCtg= 'C' and " + _
         "RrCd= '" + wPureGldCd + "' and " + _
         "RrFrLn <= " + mSzFld + " And RrToLn >= " + mSzFld + " and RrCmCd = '" + ctValueCmCd + "' and " + _
         "RrCmCurCd= '" + ws_BaseCurCd + "' and RrDmCtg= '' and RrLabMCd= '') , 0) "
  wPRt = " isnull((select " + mRtFld + " from RmRt where " + IIF(gs_Partition = ctCurrPrtn, " RrPrtKey='" + ctCurrPrtn + "' and ", "") + " RrTcTyp= 'RM' and RrCmCtg= 'C' and " + _
         "RrCd= '" + wPurePlCd + "' and " + _
         "RrFrLn <= " + mSzFld + " And RrToLn >= " + mSzFld + " and RrCmCd = '" + ctValueCmCd + "' and " + _
         "RrCmCurCd= '" + ws_BaseCurCd + "' and RrDmCtg= '' and RrLabMCd= '') , 0) "
  wSRt = " isnull((select " + mRtFld + " from RmRt where " + IIF(gs_Partition = ctCurrPrtn, " RrPrtKey='" + ctCurrPrtn + "' and ", "") + " RrTcTyp= 'RM' and RrCmCtg= 'C' and " + _
         "RrCd= '" + wPureSlCd + "' and " + _
         "RrFrLn <= " + mSzFld + " And RrToLn >= " + mSzFld + " and RrCmCd = '" + ctValueCmCd + "' and " + _
         "RrCmCurCd= '" + ws_BaseCurCd + "' and RrDmCtg= '' and RrLabMCd= '') , 0) "
  ' ***** Manali 3.03 - 23/06/08 - Pd Mod - 'L' string
  wLRt = " isnull((select " + mRtFld + " from RmRt where " + IIF(gs_Partition = ctCurrPrtn, " RrPrtKey='" + ctCurrPrtn + "' and ", "") + " RrTcTyp= 'RM' and RrCmCtg= 'C' and " + _
         "RrCd= '" + wPurePdCd + "' and " + _
         "RrFrLn <= " + mSzFld + " And RrToLn >= " + mSzFld + " and RrCmCd = '" + ctValueCmCd + "' and " + _
         "RrCmCurCd= '" + ws_BaseCurCd + "' and RrDmCtg= '' and RrLabMCd= '') , 0) "
  ' ***** Manali 3.03 - 23/06/08 - Pd Mod - 'L' string
'*********
  ' Zubin 213
  wDCXARt = " isnull((Select " + mRtFld + " from RmRt where " + IIF(gs_Partition = ctCurrPrtn, " RrPrtKey='" + ctCurrPrtn + "' and ", "") + " RrTcTyp='RM' and RrCmCtg= 'C' and " + _
         "RrCtg= RmCtg and RrSCtg= RmSCtg and RrCd= RmCd and RrCmCd = '" + ctValueCmCd + "' and " + _
         "RrCmCurCd= '" + ws_BaseCurCd + "' and " + _
         "RrFrLn <= " + mSzFld + " And RrToLn >= " + mSzFld + " and RrDmCtg= '' and RrLabMCd= '' and " + _
         mRtFld + " <> 0), 0) "
  '****** Before 2.13.0
  'RmRate = " (case when RmCtg= 'G' then " + wGRt + " when RmCtg= 'P' then " + wPRt + " " + _
           " when RmCtg= 'S' then " + wSRt + " when RmCtg in ('D', 'C', 'X', 'A') then " + wDCXARt + " else 0 end) "
           
  '****** Sachin 2.13.0 [12-08-2006] [Q/W Module]
  ' ***** Manali 3.03 - 23/06/08 - Pd Mod - 'L' added
  RmRate = " (case when (RmCtg= 'G' And " + mQw + "='W')  then " + wGRt + " when (RmCtg= 'P' And " + mQw + "='W') then " + wPRt + " " + _
           " when (RmCtg= 'S' And " + mQw + "='W') then " + wSRt + "when (RmCtg= 'L' And " + mQw + "='W') then " + wLRt + " " + _
           " when (RmCtg In ('G','P','S', 'L') And " + mQw + "='Q') Or  (RmCtg in ('D', 'C', 'X', 'A')) then " + wDCXARt + " else 0 end) "
  '****** Sachin 2.13.0 [12-08-2006] [Q/W Module]
           
End Function
Private Sub CmpStr()
  'adc("UoIssLocFr").CmpStr = "Bv.TdFrBLoc >= "
  'adc("UoIssLocTo").CmpStr = "Bv.TdFrBLoc <= "
  'adc("UoIssLocSel").CmpStr = "Bv.TdFrBLoc In "
  'adc("UoRecLocFr").CmpStr = "Bv.TdToBLoc >= "
  'adc("UoRecLocTo").CmpStr = "Bv.TdToBLoc <= "
  'adc("UoRecLocSel").CmpStr = "Bv.TdToBLoc In "
  
  adc("UoRmCdFr").CmpStr = "Td.TdRmCd >= "
  adc("UoRmCdTo").CmpStr = "Td.TdRmCd <= "
  adc("UoRmCdSel").CmpStr = "Td.TdRmCd in "
  adc("UoLn1Fr").CmpStr = "Td.TdRmSz >= "
  adc("UoLn1To").CmpStr = "Td.TdRmSz <= "
  adc("UoCstRtFr").CmpStr = "Td.TdRmStkRt >= "
  adc("UoCstRtTo").CmpStr = "TdRmStkRt <= "
  adc("UoDtFr").CmpStr = "TDt>= "
  adc("UoDtTo").CmpStr = "TDt<= "
End Sub

'StkRt-if password is change to blank, Show StockRate check box will turn into blank
Private Sub TxtPwd_Change()
    If Trim(TxtPwd.text) = "" Then
        ChkBoxArr(1).Value = Unchecked
    End If
End Sub
