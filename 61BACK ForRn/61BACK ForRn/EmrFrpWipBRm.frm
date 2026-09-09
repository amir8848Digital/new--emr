VERSION 5.00
Object = "{8767A745-088E-4CA6-8594-073D6D2DE57A}#9.2#0"; "crviewer9.dll"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "tabctl32.ocx"
Object = "{AB2CC32A-E652-42A2-B341-26ECE02F6386}#2.1#0"; "MwfCtlForSpeed7.ocx"
Begin VB.Form EmrFrpWipBRm 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "WIP Bag Raw Material Report"
   ClientHeight    =   10185
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   15330
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   10185
   ScaleWidth      =   15330
   WindowState     =   2  'Maximized
   Begin MwfCtl.MWCTL_HLP MWCTL_HLP1 
      Height          =   495
      Left            =   13260
      TabIndex        =   96
      TabStop         =   0   'False
      Top             =   9630
      Visible         =   0   'False
      Width           =   495
      _ExtentX        =   873
      _ExtentY        =   873
   End
   Begin MwfCtl.MWCTL_DST DST_REP 
      Height          =   570
      Left            =   4455
      TabIndex        =   95
      TabStop         =   0   'False
      Top             =   9630
      Width           =   6330
      _ExtentX        =   11165
      _ExtentY        =   1005
   End
   Begin VB.Frame fra_keys 
      BackColor       =   &H00C0C0C0&
      BorderStyle     =   0  'None
      Height          =   510
      Left            =   6300
      TabIndex        =   101
      Top             =   9630
      Width           =   3555
      Begin MwfCtl.MWCTL_MED aktxt 
         Height          =   345
         Index           =   0
         Left            =   210
         TabIndex        =   102
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
         TabIndex        =   103
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
      Left            =   9360
      TabIndex        =   100
      Top             =   9630
      Visible         =   0   'False
      Width           =   855
      Begin CRVIEWER9LibCtl.CRViewer9 CRV_REP 
         Height          =   495
         Left            =   480
         TabIndex        =   99
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
      TabIndex        =   97
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
      TabIndex        =   98
      Top             =   60
      Width           =   15270
      Begin TabDlg.SSTab TAB_REP 
         Height          =   9465
         Left            =   60
         TabIndex        =   104
         Top             =   60
         Width           =   15120
         _ExtentX        =   26670
         _ExtentY        =   16695
         _Version        =   393216
         Tabs            =   2
         TabsPerRow      =   2
         TabHeight       =   520
         TabCaption(0)   =   "&Data Range"
         TabPicture(0)   =   "EmrFrpWipBRm.frx":0000
         Tab(0).ControlEnabled=   -1  'True
         Tab(0).Control(0)=   "fra_tabrep(0)"
         Tab(0).Control(0).Enabled=   0   'False
         Tab(0).ControlCount=   1
         TabCaption(1)   =   "&Group Sort"
         TabPicture(1)   =   "EmrFrpWipBRm.frx":001C
         Tab(1).ControlEnabled=   0   'False
         Tab(1).Control(0)=   "fra_tabrep(1)"
         Tab(1).ControlCount=   1
         Begin VB.Frame fra_tabrep 
            BorderStyle     =   0  'None
            Height          =   9070
            Index           =   0
            Left            =   120
            TabIndex        =   107
            Top             =   360
            Width           =   15025
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
               Index           =   4
               Left            =   4440
               TabIndex        =   173
               Tag             =   "ShowPic"
               ToolTipText     =   "Check To Show Stock Rate"
               Top             =   6750
               Width           =   2220
            End
            Begin VB.Frame FraLoc 
               BorderStyle     =   0  'None
               Height          =   230
               Left            =   1920
               TabIndex        =   169
               Top             =   760
               Width           =   4065
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
                  Left            =   100
                  TabIndex        =   172
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
                  Left            =   2550
                  TabIndex        =   171
                  ToolTipText     =   "Select Production or Process"
                  Top             =   0
                  Width           =   1545
               End
            End
            Begin VB.Frame FraOptBagIn 
               BorderStyle     =   0  'None
               Height          =   285
               Left            =   2025
               TabIndex        =   161
               ToolTipText     =   "Show Bags In FG / INV / Both"
               Top             =   8760
               Width           =   3465
               Begin VB.OptionButton OptBagIn 
                  Caption         =   "FG"
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
                  TabIndex        =   164
                  ToolTipText     =   "Show Bags In FG / INV / Both"
                  Top             =   0
                  Width           =   1005
               End
               Begin VB.OptionButton OptBagIn 
                  Caption         =   "INV"
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
                  TabIndex        =   163
                  ToolTipText     =   "Show Bags In FG / INV / Both"
                  Top             =   0
                  Width           =   945
               End
               Begin VB.OptionButton OptBagIn 
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
                  TabIndex        =   162
                  ToolTipText     =   "Show Bags In FG / INV / Both"
                  Top             =   0
                  Width           =   1005
               End
            End
            Begin VB.CheckBox ChkBoxArr 
               Alignment       =   1  'Right Justify
               Caption         =   "Show Custom Pure Wt"
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
               Left            =   8520
               TabIndex        =   170
               Tag             =   "ShowZWt"
               ToolTipText     =   "Check To Show ZPurity Wt"
               Top             =   8800
               Width           =   2505
            End
            Begin VB.CheckBox ChkBoxArr 
               Alignment       =   1  'Right Justify
               Caption         =   "Show Box Details     "
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
               Left            =   8520
               TabIndex        =   166
               Tag             =   "ShowBoxDet"
               ToolTipText     =   "Check To Show Box Details"
               Top             =   8205
               Width           =   2505
            End
            Begin VB.CheckBox ChkBoxArr 
               Alignment       =   1  'Right Justify
               Caption         =   "Show Picture           "
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
               Left            =   8520
               TabIndex        =   168
               Tag             =   "ShowPic"
               ToolTipText     =   "Check To Show Picture"
               Top             =   8520
               Width           =   2505
            End
            Begin VB.Frame FraOptBagPcs 
               BorderStyle     =   0  'None
               Height          =   285
               Left            =   2025
               TabIndex        =   151
               ToolTipText     =   "Show Normal Or Rejected Bag Pieces"
               Top             =   8745
               Width           =   2595
               Begin VB.OptionButton OptBagPcs 
                  Caption         =   "Normal"
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
                  TabIndex        =   91
                  ToolTipText     =   "Show Normal Or Rejected Bag Pieces"
                  Top             =   0
                  Width           =   1005
               End
               Begin VB.OptionButton OptBagPcs 
                  Caption         =   "Rejected"
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
                  TabIndex        =   92
                  ToolTipText     =   "Show Normal Or Rejected Bag Pieces"
                  Top             =   0
                  Width           =   1185
               End
            End
            Begin VB.Frame FraOptBag 
               BorderStyle     =   0  'None
               Height          =   285
               Left            =   2025
               TabIndex        =   149
               ToolTipText     =   "Show Open Or Closed Or All Bags"
               Top             =   8445
               Width           =   3465
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
                  TabIndex        =   89
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
                  TabIndex        =   88
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
                  TabIndex        =   87
                  ToolTipText     =   "Show Open Or Closed Or All Bags"
                  Top             =   0
                  Width           =   1005
               End
            End
            Begin VB.TextBox TxtPwd 
               Height          =   285
               IMEMode         =   3  'DISABLE
               Left            =   2025
               MaxLength       =   30
               PasswordChar    =   "*"
               TabIndex        =   69
               Top             =   6750
               Width           =   2340
            End
            Begin MwfCtl.MWCTL_MED atxt 
               Height          =   285
               Index           =   1
               Left            =   4485
               TabIndex        =   7
               ToolTipText     =   "Enter To FrBag Location"
               Top             =   1035
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
               Left            =   6960
               TabIndex        =   8
               ToolTipText     =   "Enter FrBag Location Selection"
               Top             =   1035
               Width           =   8010
               _ExtentX        =   14129
               _ExtentY        =   503
               MaxLength       =   255
               DataField       =   "UOISSLOCSEL"
               IdName          =   "UOISSLOCSEL"
            End
            Begin MwfCtl.MWCTL_MED atxt 
               Height          =   285
               Index           =   0
               Left            =   2025
               TabIndex        =   6
               ToolTipText     =   "Enter From FrBag Location"
               Top             =   1035
               Width           =   1830
               _ExtentX        =   3228
               _ExtentY        =   503
               MaxLength       =   8
               DataField       =   "UOISSLOCFR"
               IdName          =   "UOISSLOCFR"
            End
            Begin MwfCtl.MWCTL_MED atxt 
               Height          =   285
               Index           =   58
               Left            =   5625
               TabIndex        =   90
               ToolTipText     =   "Show Normal Or Rejected Pieces ?"
               Top             =   8745
               Visible         =   0   'False
               Width           =   285
               _ExtentX        =   503
               _ExtentY        =   503
               Enabled         =   0   'False
               DataField       =   "UOBAGPCS"
               IdName          =   "UOBAGPCS"
            End
            Begin MwfCtl.MWCTL_MED atxt 
               Height          =   285
               Index           =   32
               Left            =   4485
               TabIndex        =   36
               ToolTipText     =   "Enter Raw Material Code To Range"
               Top             =   3600
               Width           =   2385
               _ExtentX        =   4207
               _ExtentY        =   503
               MaxLength       =   12
               DataField       =   "UORMCDTO"
               IdName          =   "UORMCDTO"
               CmpStr          =   "TdRmCd <="
            End
            Begin MwfCtl.MWCTL_MED atxt 
               Height          =   285
               Index           =   33
               Left            =   6960
               TabIndex        =   37
               ToolTipText     =   "Enter Raw Material Code Selection"
               Top             =   3600
               Width           =   8010
               _ExtentX        =   14129
               _ExtentY        =   503
               MaxLength       =   255
               DataField       =   "UORMCDSEL"
               IdName          =   "UORMCDSEL"
               CmpStr          =   "TdRmCd In"
            End
            Begin MwfCtl.MWCTL_MED atxt 
               Height          =   285
               Index           =   38
               Left            =   4485
               TabIndex        =   30
               ToolTipText     =   "Enter Raw Material Sub Category To Range"
               Top             =   3030
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
               Index           =   39
               Left            =   6960
               TabIndex        =   31
               ToolTipText     =   "Enter Raw Material Sub Category Selection"
               Top             =   3030
               Width           =   8010
               _ExtentX        =   14129
               _ExtentY        =   503
               MaxLength       =   255
               DataField       =   "UORMSCTGSEL"
               IdName          =   "UORMSCTGSEL"
               CmpStr          =   "RmSCtg In"
            End
            Begin MwfCtl.MWCTL_MED atxt 
               Height          =   285
               Index           =   34
               Left            =   2025
               TabIndex        =   23
               ToolTipText     =   "Enter From Raw Material Category"
               Top             =   2460
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
               Index           =   37
               Left            =   2025
               TabIndex        =   29
               ToolTipText     =   "Enter Raw Material Sub Category From Range"
               Top             =   3030
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
               Index           =   31
               Left            =   2025
               TabIndex        =   35
               ToolTipText     =   "Enter Raw Material Code From Range"
               Top             =   3600
               Width           =   2385
               _ExtentX        =   4207
               _ExtentY        =   503
               MaxLength       =   12
               DataField       =   "UORMCDFR"
               IdName          =   "UORMCDFR"
               CmpStr          =   "TdRmCd >="
            End
            Begin MwfCtl.MWCTL_MED atxt 
               Height          =   285
               Index           =   35
               Left            =   4485
               TabIndex        =   24
               ToolTipText     =   "Enter To Raw Material Category"
               Top             =   2460
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
               Index           =   36
               Left            =   6960
               TabIndex        =   25
               ToolTipText     =   "Enter Raw Material Category Selection"
               Top             =   2460
               Width           =   8010
               _ExtentX        =   14129
               _ExtentY        =   503
               MaxLength       =   255
               DataField       =   "UORMCTGSEL"
               IdName          =   "UORMCTGSEL"
               CmpStr          =   "RmCtg In"
            End
            Begin MwfCtl.MWCTL_MED atxt 
               Height          =   285
               Index           =   56
               Left            =   11415
               TabIndex        =   94
               ToolTipText     =   "Show Box Detail (Yes / No)"
               Top             =   8205
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
               Index           =   41
               Left            =   4485
               TabIndex        =   33
               ToolTipText     =   "Enter To Parent Code"
               Top             =   3315
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
               Index           =   42
               Left            =   6960
               TabIndex        =   34
               ToolTipText     =   "Enter Parent Code Selection"
               Top             =   3315
               Width           =   8010
               _ExtentX        =   14129
               _ExtentY        =   503
               MaxLength       =   255
               DataField       =   "UOPARENTCDSEL"
               IdName          =   "UOPARENTCDSEL"
               CmpStr          =   "RmParentCd In"
            End
            Begin MwfCtl.MWCTL_MED atxt 
               Height          =   285
               Index           =   40
               Left            =   2025
               TabIndex        =   32
               ToolTipText     =   "Enter From Parent Code"
               Top             =   3315
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
               Index           =   18
               Left            =   2025
               TabIndex        =   0
               ToolTipText     =   "Enter Company Code"
               Top             =   210
               Width           =   705
               _ExtentX        =   1244
               _ExtentY        =   503
               MaxLength       =   3
               DataType        =   4
               DataField       =   "UOCOCDFR"
               IdName          =   "UOCOCDFR"
               CmpStr          =   "BCoCd >= "
            End
            Begin MwfCtl.MWCTL_MED atxt 
               Height          =   285
               Index           =   57
               Left            =   5625
               TabIndex        =   86
               ToolTipText     =   "Show Open Or Closed Or All Bags ? (Open - 'Y', Closed - 'N', All - Blank)"
               Top             =   8445
               Visible         =   0   'False
               Width           =   285
               _ExtentX        =   503
               _ExtentY        =   503
               Enabled         =   0   'False
               DataField       =   "UOYN1"
               IdName          =   "UOYN1"
            End
            Begin MwfCtl.MWCTL_MED atxt 
               Height          =   285
               Index           =   23
               Left            =   4485
               TabIndex        =   13
               ToolTipText     =   "Enter To FgSubLoc Location"
               Top             =   1605
               Width           =   1830
               _ExtentX        =   3228
               _ExtentY        =   503
               MaxLength       =   5
               DataField       =   "UOFGSUBLOCTO"
               IdName          =   "UOFGSUBLOCTO"
               CmpStr          =   "BFgSubLoc <="
            End
            Begin MwfCtl.MWCTL_MED atxt 
               Height          =   285
               Index           =   24
               Left            =   6960
               TabIndex        =   14
               ToolTipText     =   "Enter FgSubLoc Location Selection"
               Top             =   1605
               Width           =   8010
               _ExtentX        =   14129
               _ExtentY        =   503
               MaxLength       =   255
               DataField       =   "UOFGSUBLOCSEL"
               IdName          =   "UOFGSUBLOCSEL"
               CmpStr          =   "BFgSubLoc In"
            End
            Begin MwfCtl.MWCTL_MED atxt 
               Height          =   285
               Index           =   22
               Left            =   2025
               TabIndex        =   12
               ToolTipText     =   "Enter From FgSubLoc Location"
               Top             =   1605
               Width           =   1830
               _ExtentX        =   3228
               _ExtentY        =   503
               MaxLength       =   8
               DataField       =   "UOFGSUBLOCFR"
               IdName          =   "UOFGSUBLOCFR"
               CmpStr          =   "BFgSubLoc >="
            End
            Begin MwfCtl.MWCTL_MED atxt 
               Height          =   285
               Index           =   4
               Left            =   4485
               TabIndex        =   16
               ToolTipText     =   "Enter To Fg Date"
               Top             =   1890
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
               Index           =   3
               Left            =   2025
               TabIndex        =   15
               ToolTipText     =   "Enter From Fg Date"
               Top             =   1890
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
               Index           =   55
               Left            =   11385
               TabIndex        =   93
               ToolTipText     =   "Show Picture? ( Enter 'Y'/'N')"
               Top             =   8520
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
               Index           =   25
               Left            =   2025
               TabIndex        =   17
               ToolTipText     =   "Enter From Packing List Year"
               Top             =   2175
               Width           =   345
               _ExtentX        =   609
               _ExtentY        =   503
               MaxLength       =   2
               DataField       =   "UOFDPLYYFR"
               IdName          =   "UOFDPLYYFR"
            End
            Begin MwfCtl.MWCTL_MED atxt 
               Height          =   285
               Index           =   27
               Left            =   2445
               TabIndex        =   18
               ToolTipText     =   "Enter From Packing List Character"
               Top             =   2175
               Width           =   705
               _ExtentX        =   1244
               _ExtentY        =   503
               MaxLength       =   3
               DataType        =   4
               DataField       =   "UOFDPLCHRFR"
               IdName          =   "UOFDPLCHRFR"
            End
            Begin MwfCtl.MWCTL_MED atxt 
               Height          =   285
               Index           =   29
               Left            =   3225
               TabIndex        =   19
               ToolTipText     =   "Enter From Packing List Number"
               Top             =   2175
               Width           =   825
               _ExtentX        =   1455
               _ExtentY        =   503
               MaxLength       =   0
               DataType        =   1
               DataField       =   "UOFDPLNOFR"
               IdName          =   "UOFDPLNOFR"
            End
            Begin MwfCtl.MWCTL_MED atxt 
               Height          =   285
               Index           =   26
               Left            =   4485
               TabIndex        =   20
               ToolTipText     =   "Enter To Packing List Year"
               Top             =   2175
               Width           =   345
               _ExtentX        =   609
               _ExtentY        =   503
               MaxLength       =   2
               DataField       =   "UOFDPLYYTO"
               IdName          =   "UOFDPLYYTO"
            End
            Begin MwfCtl.MWCTL_MED atxt 
               Height          =   285
               Index           =   28
               Left            =   4905
               TabIndex        =   21
               ToolTipText     =   "Enter To Packing List Character"
               Top             =   2175
               Width           =   705
               _ExtentX        =   1244
               _ExtentY        =   503
               MaxLength       =   3
               DataType        =   4
               DataField       =   "UOFDPLCHRTO"
               IdName          =   "UOFDPLCHRTO"
               CmpStr          =   "."
            End
            Begin MwfCtl.MWCTL_MED atxt 
               Height          =   285
               Index           =   30
               Left            =   5685
               TabIndex        =   22
               ToolTipText     =   "Enter To Packing List Number"
               Top             =   2175
               Width           =   825
               _ExtentX        =   1455
               _ExtentY        =   503
               MaxLength       =   0
               DataType        =   1
               DataField       =   "UOFDPLNOTO"
               IdName          =   "UOFDPLNOTO"
            End
            Begin MwfCtl.MWCTL_MED atxt 
               Height          =   260
               Index           =   59
               Left            =   2025
               TabIndex        =   73
               ToolTipText     =   "Specify Bag Type or ' ' for All"
               Top             =   7300
               Width           =   285
               _ExtentX        =   503
               _ExtentY        =   450
               DataType        =   4
               DataField       =   "UOBAGTYP"
               ReCalcParent    =   "UOBCHRFR,UOBCHR,UOBCHR2FR,UOBCHR2TO,UOBCHR2SEL"
               IdName          =   "UOBAGTYP"
            End
            Begin MwfCtl.MWCTL_MED atxt 
               Height          =   285
               Index           =   60
               Left            =   2025
               TabIndex        =   174
               ToolTipText     =   "Enter Bag Character"
               Top             =   7560
               Visible         =   0   'False
               Width           =   60
               _ExtentX        =   106
               _ExtentY        =   503
               MaxLength       =   3
               DataField       =   "UOBCHR"
               ReCalcOn        =   "UOBAGTYP"
               IdName          =   "UOBCHR"
            End
            Begin MwfCtl.MWCTL_MED atxt 
               Height          =   285
               Index           =   43
               Left            =   2025
               TabIndex        =   38
               ToolTipText     =   "Enter From Lot No."
               Top             =   3885
               Width           =   1905
               _ExtentX        =   3360
               _ExtentY        =   503
               MaxLength       =   12
               DataType        =   4
               DataField       =   "UOLOTNOFR"
               IdName          =   "UOLOTNOFR"
               CmpStr          =   "TdLotNo>= "
            End
            Begin MwfCtl.MWCTL_MED atxt 
               Height          =   285
               Index           =   44
               Left            =   4485
               TabIndex        =   39
               ToolTipText     =   "Enter To Lot No."
               Top             =   3885
               Width           =   1905
               _ExtentX        =   3360
               _ExtentY        =   503
               MaxLength       =   12
               DataType        =   4
               DataField       =   "UOLOTNOTO"
               IdName          =   "UOLOTNOTO"
               CmpStr          =   "TdLotNo<= "
            End
            Begin MwfCtl.MWCTL_MED atxt 
               Height          =   285
               Index           =   45
               Left            =   6960
               TabIndex        =   40
               ToolTipText     =   "Enter Lot No. Selection"
               Top             =   3885
               Width           =   8010
               _ExtentX        =   14129
               _ExtentY        =   503
               MaxLength       =   255
               DataType        =   4
               DataField       =   "UOLOTNOSEL"
               IdName          =   "UOLOTNOSEL"
               CmpStr          =   "TdLotNo In "
            End
            Begin MwfCtl.MWCTL_MED atxt 
               Height          =   285
               Index           =   47
               Left            =   4485
               TabIndex        =   42
               ToolTipText     =   "Enter To Range Size"
               Top             =   4170
               Width           =   1905
               _ExtentX        =   3360
               _ExtentY        =   503
               MaxLength       =   10
               DataField       =   "UORNGSZTO"
               IdName          =   "UORNGSZTO"
               CmpStr          =   "RrRngSz<= "
            End
            Begin MwfCtl.MWCTL_MED atxt 
               Height          =   285
               Index           =   48
               Left            =   6960
               TabIndex        =   43
               ToolTipText     =   "Enter Range Size Selection"
               Top             =   4170
               Width           =   8010
               _ExtentX        =   14129
               _ExtentY        =   503
               MaxLength       =   255
               DataField       =   "UORNGSZSEL"
               IdName          =   "UORNGSZSEL"
               CmpStr          =   "RrRngSz In "
            End
            Begin MwfCtl.MWCTL_MED atxt 
               Height          =   285
               Index           =   20
               Left            =   4485
               TabIndex        =   65
               ToolTipText     =   "Enter To Customer"
               Top             =   6180
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
               Index           =   21
               Left            =   6960
               TabIndex        =   66
               ToolTipText     =   "Enter Customer Selection"
               Top             =   6180
               Width           =   8010
               _ExtentX        =   14129
               _ExtentY        =   503
               MaxLength       =   255
               DataType        =   4
               DataField       =   "UOCMCDSEL"
               IdName          =   "UOCMCDSEL"
               CmpStr          =   "OmCmCd In"
            End
            Begin MwfCtl.MWCTL_MED atxt 
               Height          =   285
               Index           =   8
               Left            =   5265
               TabIndex        =   58
               ToolTipText     =   "Enter To Order Year"
               Top             =   5595
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
               Index           =   6
               Left            =   4485
               TabIndex        =   57
               ToolTipText     =   "Enter To Order Tc"
               Top             =   5595
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
               Index           =   10
               Left            =   5685
               TabIndex        =   59
               ToolTipText     =   "Enter To Order Character"
               Top             =   5595
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
               Index           =   12
               Left            =   4485
               TabIndex        =   62
               ToolTipText     =   "Enter To Order Number"
               Top             =   5895
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
               Index           =   14
               Left            =   5370
               TabIndex        =   63
               ToolTipText     =   "Enter To Order Serial"
               Top             =   5895
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
               Index           =   52
               Left            =   4485
               TabIndex        =   47
               ToolTipText     =   "Enter To Cost Rate"
               Top             =   4740
               Width           =   1350
               _ExtentX        =   2381
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
               Index           =   16
               Left            =   4485
               TabIndex        =   52
               ToolTipText     =   "Enter To Design Code"
               Top             =   5310
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
               Index           =   17
               Left            =   6960
               TabIndex        =   53
               ToolTipText     =   "Enter Design Code Selection"
               Top             =   5310
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
               Index           =   50
               Left            =   4485
               TabIndex        =   45
               ToolTipText     =   "Enter To Raw Material Len/Sv"
               Top             =   4455
               Width           =   1185
               _ExtentX        =   2090
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
               Index           =   7
               Left            =   2805
               TabIndex        =   55
               ToolTipText     =   "Enter From Order Year"
               Top             =   5595
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
               Index           =   5
               Left            =   2040
               TabIndex        =   54
               ToolTipText     =   "Enter From Order Tc"
               Top             =   5595
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
               Index           =   9
               Left            =   3225
               TabIndex        =   56
               ToolTipText     =   "Enter From Order Character"
               Top             =   5595
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
               Index           =   11
               Left            =   2025
               TabIndex        =   60
               ToolTipText     =   "Enter From Order Number"
               Top             =   5895
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
               Index           =   13
               Left            =   2910
               TabIndex        =   61
               ToolTipText     =   "Enter From Order Serial"
               Top             =   5895
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
               Index           =   53
               Left            =   2025
               TabIndex        =   70
               ToolTipText     =   "Specify If The Value To Be Shown Should Be Stock Value, Cost Value Or Sale Value Or No Value"
               Top             =   7025
               Width           =   1305
               _ExtentX        =   2302
               _ExtentY        =   503
               MaxLength       =   8
               DataField       =   "UOPMCDFR"
               IdName          =   "UOPMCDFR"
            End
            Begin MwfCtl.MWCTL_MED atxt 
               Height          =   285
               Index           =   19
               Left            =   2025
               TabIndex        =   64
               ToolTipText     =   "Enter From Customer"
               Top             =   6180
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
               Index           =   54
               Left            =   10545
               TabIndex        =   72
               ToolTipText     =   "Enter The Percentage By Which Value Has To Be Incremented"
               Top             =   6965
               Width           =   885
               _ExtentX        =   1561
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
               Index           =   51
               Left            =   2025
               TabIndex        =   46
               ToolTipText     =   "Enter From Cost Rate"
               Top             =   4740
               Width           =   1350
               _ExtentX        =   2381
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
               Index           =   15
               Left            =   2025
               TabIndex        =   51
               ToolTipText     =   "Enter From Design Code"
               Top             =   5310
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
               Index           =   49
               Left            =   2025
               TabIndex        =   44
               ToolTipText     =   "Enter From Raw Material Len/Sv"
               Top             =   4455
               Width           =   1245
               _ExtentX        =   2196
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
               Index           =   46
               Left            =   2025
               TabIndex        =   41
               ToolTipText     =   "Enter From Range Size"
               Top             =   4170
               Width           =   1905
               _ExtentX        =   3360
               _ExtentY        =   503
               MaxLength       =   10
               DataField       =   "UORNGSZFR"
               IdName          =   "UORNGSZFR"
               CmpStr          =   "RrRngSz>= "
            End
            Begin MwfCtl.MWCTL_MED atxt 
               Height          =   285
               Index           =   61
               Left            =   4485
               TabIndex        =   27
               ToolTipText     =   "Enter Raw Material Stone Class To Range"
               Top             =   2745
               Width           =   1365
               _ExtentX        =   2408
               _ExtentY        =   503
               MaxLength       =   5
               DataField       =   "UOSTNCLSTO"
               IdName          =   "UOSTNCLSTO"
               CmpStr          =   "IsNull(SCls.PSCd,'') <="
            End
            Begin MwfCtl.MWCTL_MED atxt 
               Height          =   285
               Index           =   62
               Left            =   6960
               TabIndex        =   28
               ToolTipText     =   "Enter Raw Material Stone Class Selection"
               Top             =   2745
               Width           =   8010
               _ExtentX        =   14129
               _ExtentY        =   503
               MaxLength       =   255
               DataField       =   "UOSTNCLSSEL"
               IdName          =   "UOSTNCLSSEL"
               CmpStr          =   "IsNull(SCls.PSCd,'') In"
            End
            Begin MwfCtl.MWCTL_MED atxt 
               Height          =   285
               Index           =   63
               Left            =   2025
               TabIndex        =   26
               ToolTipText     =   "Enter Raw Material Stone Class From Range"
               Top             =   2745
               Width           =   1365
               _ExtentX        =   2408
               _ExtentY        =   503
               MaxLength       =   5
               DataField       =   "UOSTNCLSFR"
               IdName          =   "UOSTNCLSFR"
               CmpStr          =   "IsNull(SCls.PSCd,'') >="
            End
            Begin MwfCtl.MWCTL_MED atxt 
               Height          =   285
               Index           =   64
               Left            =   4485
               TabIndex        =   10
               ToolTipText     =   "Enter To By Worker"
               Top             =   1320
               Width           =   1275
               _ExtentX        =   2249
               _ExtentY        =   503
               MaxLength       =   8
               DataField       =   "UOWRKTO"
               IdName          =   "UOWRKTO"
               CmpStr          =   "BWrk <="
            End
            Begin MwfCtl.MWCTL_MED atxt 
               Height          =   285
               Index           =   65
               Left            =   6960
               TabIndex        =   11
               ToolTipText     =   "Enter By Worker Selection"
               Top             =   1320
               Width           =   8010
               _ExtentX        =   14129
               _ExtentY        =   503
               MaxLength       =   255
               DataField       =   "UOWRKSEL"
               IdName          =   "UOWRKSEL"
               CmpStr          =   "BWrk In"
            End
            Begin MwfCtl.MWCTL_MED atxt 
               Height          =   285
               Index           =   66
               Left            =   2025
               TabIndex        =   9
               ToolTipText     =   "Enter From By Worker"
               Top             =   1320
               Width           =   1275
               _ExtentX        =   2249
               _ExtentY        =   503
               MaxLength       =   8
               DataField       =   "UOWRKFR"
               IdName          =   "UOWRKFR"
               CmpStr          =   "BWrk >="
            End
            Begin MwfCtl.MWCTL_MED atxt 
               Height          =   285
               Index           =   67
               Left            =   4485
               TabIndex        =   1
               ToolTipText     =   "Enter To Company Code"
               Top             =   210
               Width           =   705
               _ExtentX        =   1244
               _ExtentY        =   503
               MaxLength       =   3
               DataType        =   4
               DataField       =   "UOCOCDTO"
               IdName          =   "UOCOCDTO"
               CmpStr          =   "BCoCd <= "
            End
            Begin MwfCtl.MWCTL_MED atxt 
               Height          =   285
               Index           =   68
               Left            =   6960
               TabIndex        =   2
               ToolTipText     =   "Enter Design Code Selection"
               Top             =   210
               Width           =   8010
               _ExtentX        =   14129
               _ExtentY        =   503
               MaxLength       =   255
               DataField       =   "UOCOCDSEL"
               IdName          =   "UOCOCDSEL"
               CmpStr          =   "BCoCd In"
            End
            Begin MwfCtl.MWCTL_MED atxt 
               Height          =   285
               Index           =   69
               Left            =   4485
               TabIndex        =   4
               ToolTipText     =   "Enter To Parent Code"
               Top             =   495
               Width           =   1710
               _ExtentX        =   3016
               _ExtentY        =   503
               MaxLength       =   8
               DataField       =   "UOPRNTCDTO"
               IdName          =   "UOPRNTCDTO"
               CmpStr          =   "LocPrntCd <="
            End
            Begin MwfCtl.MWCTL_MED atxt 
               Height          =   285
               Index           =   70
               Left            =   6960
               TabIndex        =   5
               ToolTipText     =   "Enter Parent Code Selection"
               Top             =   510
               Width           =   8010
               _ExtentX        =   14129
               _ExtentY        =   503
               MaxLength       =   255
               DataField       =   "UOPRNTCDSEL"
               IdName          =   "UOPRNTCDSEL"
               CmpStr          =   "LocPrntCd In"
            End
            Begin MwfCtl.MWCTL_MED atxt 
               Height          =   285
               Index           =   71
               Left            =   2025
               TabIndex        =   3
               ToolTipText     =   "Enter From Parent Code"
               Top             =   495
               Width           =   1710
               _ExtentX        =   3016
               _ExtentY        =   503
               MaxLength       =   8
               DataField       =   "UOPRNTCDFR"
               IdName          =   "UOPRNTCDFR"
               CmpStr          =   "LocPrntCd >="
            End
            Begin MwfCtl.MWCTL_MED atxt 
               Height          =   285
               Index           =   72
               Left            =   10545
               TabIndex        =   74
               ToolTipText     =   "Enter WIP As Of Date"
               Top             =   7260
               Width           =   885
               _ExtentX        =   1561
               _ExtentY        =   503
               MaxLength       =   8
               DataType        =   3
               DataField       =   "UOASOFDT"
               IdName          =   "UOASOFDT"
            End
            Begin MwfCtl.MWCTL_MED atxt 
               Height          =   285
               Index           =   73
               Left            =   2025
               TabIndex        =   67
               ToolTipText     =   "Show Bags which are Dalayed By More than or equal to Entered Days: (Enter zero to see all Bags)"
               Top             =   6465
               Width           =   645
               _ExtentX        =   1138
               _ExtentY        =   503
               Alignment       =   2
               Mask            =   "##0"
               MaxLength       =   3
               DataType        =   1
               DataField       =   "UOBDLDAYS"
               IdName          =   "UOBDLDAYS"
            End
            Begin MwfCtl.MWCTL_MED atxt 
               Height          =   285
               Index           =   74
               Left            =   4485
               TabIndex        =   68
               ToolTipText     =   "Show Bags which are Dalayed By Less  than or equal to Entered Days: (Enter zero to see all Bags)"
               Top             =   6465
               Width           =   645
               _ExtentX        =   1138
               _ExtentY        =   503
               Alignment       =   2
               Mask            =   "##0"
               MaxLength       =   3
               DataType        =   1
               DataField       =   "UOBDLDAYSTO"
               IdName          =   "UOBDLDAYSTO"
            End
            Begin MwfCtl.MWCTL_MED atxt 
               Height          =   285
               Index           =   75
               Left            =   4485
               TabIndex        =   49
               ToolTipText     =   "Enter To Design Category"
               Top             =   5025
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
               Index           =   76
               Left            =   6960
               TabIndex        =   50
               ToolTipText     =   "Enter Design Category Selection"
               Top             =   5025
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
               Index           =   77
               Left            =   2025
               TabIndex        =   48
               ToolTipText     =   "Enter From Design Category"
               Top             =   5025
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
               Index           =   78
               Left            =   4500
               TabIndex        =   83
               ToolTipText     =   "Enter To Bag Year"
               Top             =   8145
               Visible         =   0   'False
               Width           =   345
               _ExtentX        =   609
               _ExtentY        =   503
               MaxLength       =   2
               DataField       =   "UOBYYTO"
               IdName          =   "UOBYYTO"
               CmpStr          =   "BYy <="
            End
            Begin MwfCtl.MWCTL_MED atxt 
               Height          =   285
               Index           =   79
               Left            =   2025
               TabIndex        =   80
               ToolTipText     =   "Enter From Bag Year"
               Top             =   8145
               Visible         =   0   'False
               Width           =   345
               _ExtentX        =   609
               _ExtentY        =   503
               MaxLength       =   2
               DataField       =   "UOBYYFR"
               IdName          =   "UOBYYFR"
               CmpStr          =   "BYy >="
            End
            Begin MwfCtl.MWCTL_MED atxt 
               Height          =   285
               Index           =   80
               Left            =   2445
               TabIndex        =   81
               ToolTipText     =   "Enter From Bag Character"
               Top             =   8145
               Visible         =   0   'False
               Width           =   705
               _ExtentX        =   1244
               _ExtentY        =   503
               MaxLength       =   3
               DataField       =   "UOBCHRFR"
               IdName          =   "UOBCHRFR"
               CmpStr          =   "BChr >="
            End
            Begin MwfCtl.MWCTL_MED atxt 
               Height          =   285
               Index           =   81
               Left            =   3225
               TabIndex        =   82
               ToolTipText     =   "Enter From Bag Number"
               Top             =   8145
               Visible         =   0   'False
               Width           =   825
               _ExtentX        =   1455
               _ExtentY        =   503
               MaxLength       =   0
               DataType        =   1
               DataField       =   "UOBNOFR"
               IdName          =   "UOBNOFR"
               CmpStr          =   "BNo >="
            End
            Begin MwfCtl.MWCTL_MED atxt 
               Height          =   285
               Index           =   82
               Left            =   4905
               TabIndex        =   84
               ToolTipText     =   "Enter To Bag Character"
               Top             =   8145
               Visible         =   0   'False
               Width           =   705
               _ExtentX        =   1244
               _ExtentY        =   503
               MaxLength       =   3
               DataField       =   "UOBCHRTO"
               IdName          =   "UOBCHRTO"
               CmpStr          =   "BChr <="
            End
            Begin MwfCtl.MWCTL_MED atxt 
               Height          =   285
               Index           =   83
               Left            =   5655
               TabIndex        =   85
               ToolTipText     =   "Enter To Bag Number"
               Top             =   8145
               Visible         =   0   'False
               Width           =   825
               _ExtentX        =   1455
               _ExtentY        =   503
               MaxLength       =   0
               DataType        =   1
               DataField       =   "UOBNOTO"
               IdName          =   "UOBNOTO"
               CmpStr          =   "BNo <="
            End
            Begin MwfCtl.MWCTL_MED atxt 
               Height          =   285
               Index           =   84
               Left            =   6105
               TabIndex        =   71
               ToolTipText     =   "Specify If The Other Value To Be Shown Should Be Stock Value, Cost Value Or Sale Value Or No Value"
               Top             =   7035
               Width           =   1305
               _ExtentX        =   2302
               _ExtentY        =   503
               MaxLength       =   8
               DataField       =   "UOPMCDTO"
               IdName          =   "UOPMCDTO"
            End
            Begin MwfCtl.MWCTL_MED atxt 
               Height          =   285
               Index           =   85
               Left            =   4500
               TabIndex        =   79
               ToolTipText     =   "Enter To Design Code"
               Top             =   7845
               Width           =   2385
               _ExtentX        =   4207
               _ExtentY        =   503
               MaxLength       =   15
               DataField       =   "UOBOETO"
               IdName          =   "UOBOETO"
               CmpStr          =   "DmSrcDsgCd <="
            End
            Begin MwfCtl.MWCTL_MED atxt 
               Height          =   285
               Index           =   86
               Left            =   2025
               TabIndex        =   78
               ToolTipText     =   "Enter From Design Code"
               Top             =   7845
               Width           =   2385
               _ExtentX        =   4207
               _ExtentY        =   503
               MaxLength       =   15
               DataField       =   "UOBOEFR"
               IdName          =   "UOBOEFR"
               CmpStr          =   "DmSrcDsgCd >="
            End
            Begin MwfCtl.MWCTL_MED atxt 
               Height          =   240
               Index           =   87
               Left            =   11370
               TabIndex        =   160
               ToolTipText     =   "Show Custom Pure Wt? (Y/N)"
               Top             =   8805
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
               Index           =   88
               Left            =   6015
               TabIndex        =   165
               ToolTipText     =   "Show Bags In FG / INV / Both"
               Top             =   8760
               Visible         =   0   'False
               Width           =   315
               _ExtentX        =   556
               _ExtentY        =   503
               Enabled         =   0   'False
               DataField       =   "UOYN4"
               IdName          =   "UOYN4"
            End
            Begin MwfCtl.MWCTL_MED atxt 
               Height          =   240
               Index           =   89
               Left            =   6840
               TabIndex        =   175
               ToolTipText     =   "Show Custom Pure Wt? (Y/N)"
               Top             =   6720
               Visible         =   0   'False
               Width           =   285
               _ExtentX        =   503
               _ExtentY        =   423
               Enabled         =   0   'False
               DataField       =   "UOYN5"
               IdName          =   "UOYN5"
            End
            Begin MwfCtl.MWCTL_MED atxt 
               Height          =   285
               Index           =   90
               Left            =   6960
               TabIndex        =   176
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
               Index           =   91
               Left            =   2025
               TabIndex        =   75
               ToolTipText     =   "Enter Bag Character From"
               Top             =   7560
               Width           =   795
               _ExtentX        =   1402
               _ExtentY        =   503
               MaxLength       =   3
               DataField       =   "UOBCHR2FR"
               ReCalcOn        =   "UOBAGTYP"
               IdName          =   "UOBCHR2FR"
            End
            Begin MwfCtl.MWCTL_MED atxt 
               Height          =   285
               Index           =   92
               Left            =   4500
               TabIndex        =   76
               ToolTipText     =   "Enter Bag Character To"
               Top             =   7560
               Width           =   800
               _ExtentX        =   1402
               _ExtentY        =   503
               MaxLength       =   3
               DataField       =   "UOBCHR2TO"
               ReCalcOn        =   "UOBAGTYP"
               IdName          =   "UOBCHR2TO"
            End
            Begin MwfCtl.MWCTL_MED atxt 
               Height          =   285
               Index           =   93
               Left            =   6960
               TabIndex        =   77
               ToolTipText     =   "Enter Bag Character Selection"
               Top             =   7560
               Width           =   8025
               _ExtentX        =   14155
               _ExtentY        =   503
               MaxLength       =   255
               DataField       =   "UOBCHR2SEL"
               ReCalcOn        =   "UOBAGTYP"
               IdName          =   "UOBCHR2SEL"
            End
            Begin VB.Label LblBagIn 
               BackStyle       =   0  'Transparent
               Caption         =   "Bags In"
               BeginProperty Font 
                  Name            =   "Arial"
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
               TabIndex        =   167
               Top             =   8760
               Width           =   1185
            End
            Begin VB.Label LblSrcDsgCd 
               BackStyle       =   0  'Transparent
               Caption         =   "Src Sk/ Dsg Cd"
               BeginProperty Font 
                  Name            =   "Arial"
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
               TabIndex        =   159
               ToolTipText     =   "Location"
               Top             =   7845
               Width           =   1395
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
               Left            =   4200
               TabIndex        =   158
               ToolTipText     =   "Location"
               Top             =   7035
               Width           =   1845
            End
            Begin VB.Label LblBagNo 
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
               Left            =   60
               TabIndex        =   157
               Top             =   8145
               Visible         =   0   'False
               Width           =   1605
            End
            Begin VB.Label LblBgSlash1 
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
               Left            =   2385
               TabIndex        =   156
               Top             =   8145
               Visible         =   0   'False
               Width           =   105
            End
            Begin VB.Label LblBgSlash3 
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
               Left            =   4830
               TabIndex        =   155
               Top             =   8145
               Visible         =   0   'False
               Width           =   105
            End
            Begin VB.Label LblBgSlash2 
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
               Left            =   3165
               TabIndex        =   154
               Top             =   8145
               Visible         =   0   'False
               Width           =   105
            End
            Begin VB.Label LblBgSlash4 
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
               Left            =   5595
               TabIndex        =   153
               Top             =   8160
               Visible         =   0   'False
               Width           =   105
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
               Index           =   4
               Left            =   60
               TabIndex        =   152
               Top             =   5040
               Width           =   1425
            End
            Begin VB.Label Label2 
               BackStyle       =   0  'Transparent
               Caption         =   "Bag Delay Days"
               BeginProperty Font 
                  Name            =   "Arial"
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
               TabIndex        =   150
               Top             =   6465
               Width           =   1485
            End
            Begin VB.Label LblAsOfDt 
               BackColor       =   &H00C0C0C0&
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
               Left            =   8580
               TabIndex        =   148
               Top             =   7260
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
               Left            =   60
               TabIndex        =   147
               Top             =   495
               Width           =   1425
            End
            Begin VB.Label LblWrk 
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
               Left            =   60
               TabIndex        =   146
               Top             =   1320
               Width           =   1425
            End
            Begin VB.Label LblBagTyp 
               BackStyle       =   0  'Transparent
               Caption         =   "Enter Bag Type"
               BeginProperty Font 
                  Name            =   "Arial"
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
               TabIndex        =   145
               ToolTipText     =   "Location"
               Top             =   7320
               Width           =   1875
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
               TabIndex        =   144
               Top             =   2745
               Width           =   1425
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
               Left            =   2850
               TabIndex        =   143
               Top             =   5895
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
               Left            =   3165
               TabIndex        =   142
               Top             =   5595
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
               Index           =   27
               Left            =   2745
               TabIndex        =   141
               Top             =   5595
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
               Index           =   30
               Left            =   5310
               TabIndex        =   140
               Top             =   5895
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
               Left            =   5625
               TabIndex        =   139
               Top             =   5595
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
               Left            =   5205
               TabIndex        =   138
               Top             =   5595
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
               TabIndex        =   137
               Top             =   6195
               Width           =   1425
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
               TabIndex        =   136
               ToolTipText     =   "Location"
               Top             =   7035
               Width           =   1485
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
               TabIndex        =   135
               Top             =   5340
               Width           =   1515
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
               TabIndex        =   134
               Top             =   5910
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
               TabIndex        =   133
               Top             =   5625
               Width           =   1395
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
               TabIndex        =   132
               Top             =   4740
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
               TabIndex        =   131
               Top             =   4455
               Width           =   1425
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
               Height          =   285
               Index           =   0
               Left            =   8580
               TabIndex        =   130
               Top             =   6965
               Width           =   1965
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
               TabIndex        =   129
               Top             =   6750
               Width           =   1575
            End
            Begin VB.Label Label1 
               BackStyle       =   0  'Transparent
               Caption         =   "Range Size"
               BeginProperty Font 
                  Name            =   "Arial"
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
               TabIndex        =   128
               Top             =   4160
               Width           =   1425
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
               TabIndex        =   127
               Top             =   3885
               Width           =   1425
            End
            Begin VB.Label LblProcess 
               BackStyle       =   0  'Transparent
               Caption         =   "Location"
               BeginProperty Font 
                  Name            =   "Arial"
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
               Top             =   1035
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
               TabIndex        =   125
               ToolTipText     =   "Location"
               Top             =   2460
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
               TabIndex        =   124
               Top             =   3030
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
               TabIndex        =   123
               Top             =   3600
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
               Index           =   13
               Left            =   60
               TabIndex        =   122
               Top             =   3315
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
               Index           =   15
               Left            =   60
               TabIndex        =   121
               Top             =   210
               Width           =   1335
            End
            Begin VB.Label lblFgSubLoc 
               BackStyle       =   0  'Transparent
               Caption         =   "Fg Sub Loc"
               BeginProperty Font 
                  Name            =   "Arial"
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
               TabIndex        =   120
               Top             =   1605
               Width           =   1425
            End
            Begin VB.Label LblFgDt 
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
               Left            =   60
               TabIndex        =   119
               Top             =   1890
               Width           =   1425
            End
            Begin VB.Label LblPL 
               BackStyle       =   0  'Transparent
               Caption         =   "Packing List"
               BeginProperty Font 
                  Name            =   "Arial"
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
               TabIndex        =   118
               Top             =   2175
               Width           =   1545
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
               Left            =   60
               TabIndex        =   117
               Top             =   7560
               Width           =   1665
            End
            Begin VB.Label LblPlFrSlsh1 
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
               Left            =   2385
               TabIndex        =   116
               Top             =   2175
               Width           =   105
            End
            Begin VB.Label LblPlFrSlsh2 
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
               Left            =   3165
               TabIndex        =   115
               Top             =   2175
               Width           =   105
            End
            Begin VB.Label LblPlFrSlsh3 
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
               Left            =   4845
               TabIndex        =   114
               Top             =   2175
               Width           =   105
            End
            Begin VB.Label LblPlFrSlsh4 
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
               Left            =   5625
               TabIndex        =   113
               Top             =   2175
               Width           =   105
            End
            Begin VB.Label LblOpnCls 
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
               Left            =   60
               TabIndex        =   112
               Top             =   8445
               Width           =   1095
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
               Height          =   240
               Index           =   1
               Left            =   2025
               TabIndex        =   111
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
               Height          =   195
               Index           =   2
               Left            =   4485
               TabIndex        =   110
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
               Left            =   6960
               TabIndex        =   109
               Top             =   0
               Width           =   1260
            End
            Begin VB.Label LblBagPcs 
               BackStyle       =   0  'Transparent
               Caption         =   "Bag Pcs"
               BeginProperty Font 
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
               Left            =   60
               TabIndex        =   108
               Top             =   8745
               Width           =   1680
            End
         End
         Begin VB.Frame fra_tabrep 
            BorderStyle     =   0  'None
            Enabled         =   0   'False
            Height          =   9070
            Index           =   1
            Left            =   -74940
            TabIndex        =   105
            Top             =   360
            Width           =   15025
            Begin MwfCtl.MWCTL_GRP GRP_REP 
               Height          =   360
               Left            =   420
               TabIndex        =   106
               Top             =   390
               Width           =   10125
               _ExtentX        =   17859
               _ExtentY        =   635
            End
         End
      End
   End
End
Attribute VB_Name = "EmrFrpWipBRm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim Rep As New EmrRepWipBRm
Dim moCn As MwfLib.MDOConnection   '*** (Bef speed) bef 24/11/06

'****** Sachin 2.14.0 - [25-11-2006] - [Multi-Company] ******
Dim ws_MultiCoMod As Boolean
'****** Sachin 2.14.0 - [25-11-2006] - [Multi-Company] ******

Private Sub SetGroupSort()
  '****** (Jen 2.12 Patch 3) Added third element in arrays of group options ******
  '*** (Bef 2.14)
  'GRP_REP.Add "Process", "BCoCd+BLoc", "'['+BCoCd+'] '+BLoc", "", "", "Company Code,Loc Prnt Code", "", "", "(Select LocDesc From Loc Where LocCoCd = BCoCd and LocTyp= 'P' and LocCd= BLoc) "
  '*** (Bef 2.14)
  
  '*** (Jen 2.14)
  GRP_REP.Add "Location", "BCoCd+IsNull(Lc.LocCd, '')", "'['+BCoCd+'] '+IsNull(Lc.LocCd, '')", "", "", "Company Code,Loc Prnt Code", "", "", "IsNull(Lc.LocDesc, '') "
  GRP_REP.Add "Process", "BCoCd+IsNull(Lc.LocPrcs, '')", "'['+BCoCd+'] '+IsNull(Lc.LocPrcs, '')", "", "", "Company Code,Loc Prnt Code", "", "", "(Select LocDesc From Loc Where LocCoCd = BCoCd and LocTyp= 'R' and LocCd= Lc.LocPrcs) "
  '*** (Jen 2.14)
  
  GRP_REP.Add "Rm Ctg", "Rtrim(RmCtg)", "", "", "", "", "", "", "(Select PDesc From Param Where PTyp= 'RMCTG' and PMCd = RmCtg) "
  GRP_REP.Add "Rm SubCtg", "Rtrim(RmCtg)+Rtrim(RmSCtg)", "'('+ RTrim(RmCtg)+ ') '+ RmSCtg", "", "", "Rm Ctg", "", "", "(Select PDesc From Param Where PTyp= 'RMSCTG' and PMCd = RmCtg and PSCd= RmSCtg)"
 
  'GRP_REP.Add "Rm Code", "RmCtg+RmSCtg+RmCd", "'('+ RTrim(RmCtg)+'/'+ RmSCtg+ ') '+ RmCd", "", "", "Rm Ctg,Rm SubCtg", "", "", Array("(Select RmDesc From RmMst Where RmCd= TdRmCd)", "(Select RmDesc From RmMst Where RmCd= FrRmCd)")
  GRP_REP.Add "Rm Code", "Rtrim(RmCtg)+Rtrim(RmSCtg)+Rtrim(RmCd)", "'('+ RTrim(RmCtg)+'/'+ RmSCtg+ ') '+ RmCd", "", "", "Rm Ctg,Rm SubCtg,BaseCd + RmCd,Parent Cd", "", "", "SubString(RmDesc, 1, 30)"
  GRP_REP.Add "BaseCd + RmCd", "RmBaseCd+RmCd", "'('+ RTrim(RmBaseCd) + ') '+ RmCd", "", "", "Rm Ctg,Rm SubCtg,Rm Code,Parent Cd", "", "", "SubString(RmDesc, 1, 30)"
  GRP_REP.Add "Parent Cd", "Rtrim(RmCtg)+Rtrim(RmSCtg)+(case when RmParentCd= '' then RmCd else RmParentCd end)", "'('+ RTrim(RmCtg)+'/'+ RmSCtg+ ') '+ (case when RmParentCd= '' then RmCd else RmParentCd end)", "RmCtg,RmSCtg", "hRmCtg,hRmSCtg", "Rm Ctg,Rm SubCtg,BaseCd + RmCd", "", "", "(Select PDesc from Param where PTyp= 'PARENTCD' and PMCd= RmParentCd) "
  
  '****************Geeta**************Emr206*********12/04/04
  Dim wsPtrYN As String, wStkBySz As String
  wsPtrYN = "(Select PValue3 from Param where PTyp= 'RMSCTG' and PMCd= RmCtg and PSCd= RmSCtg)"
  
  ' **** Manali 3.03 - 07/06/08 - RmSzDesc added
  GRP_REP.Add "Rm Sv/Sz", Array("case when " + wsPtrYN + "= 'Y' then " + GetRmSzDBFldNm("TdRmSz", "TdRmSz2", "TdRmSz3") + " else (case when RisRmSzDesc is null then " + GetRmSzDBFldNm("TdRmSz", "TdRmSz2", "TdRmSz3") + " else RisRmSzDesc end) end ", _
            "case when " + wsPtrYN + "= 'Y' then " + GetRmSzDBFldNm("FrRmSz", "FrRmSz2", "FrRmSz3") + " else (case when RisRmSzDesc is null then " + GetRmSzDBFldNm("FrRmSz", "FrRmSz2", "FrRmSz3") + "  else RisRmSzDesc end) end ", _
            "case when " + wsPtrYN + "= 'Y' then " + GetRmSzDBFldNm("FrRmSz", "FrRmSz2", "FrRmSz3") + " else (case when RisRmSzDesc is null then " + GetRmSzDBFldNm("FrRmSz", "FrRmSz2", "FrRmSz3") + "  else RisRmSzDesc end) end "), _
            "", "wRmSz", "hTdRmSz", "Rm Ptr,Rm StkRt,(RmCd) Sv/Sz,(RmCd) StkRt", "", "", ""

 GRP_REP.Add "(RmCd) Sv/Sz", Array("TdRmCd +case when " + wsPtrYN + "='Y' then " + GetRmSzDBFldNm("TdRmSz", "TdRmSz2", "TdRmSz3") + " else (case when RisRmSzDesc is null then " + GetRmSzDBFldNm("TdRmSz", "TdRmSz2", "TdRmSz3") + "  else RisRmSzDesc end) end ", _
            "FrRmCd+case when " + wsPtrYN + "='Y' then " + GetRmSzDBFldNm("FrRmSz", "FrRmSz2", "FrRmSz3") + " else (case when RisRmSzDesc is null then " + GetRmSzDBFldNm("FrRmSz", "FrRmSz2", "FrRmSz3") + "  else RisRmSzDesc end) end ", _
            "FrRmCd+case when " + wsPtrYN + "='Y' then " + GetRmSzDBFldNm("FrRmSz", "FrRmSz2", "FrRmSz3") + " else (case when RisRmSzDesc is null then " + GetRmSzDBFldNm("FrRmSz", "FrRmSz2", "FrRmSz3") + "  else RisRmSzDesc end) end ", _
            "TR.TdRmCd +case when " + wsPtrYN + "= 'Y' then " + GetRmSzDBFldNm("TR.TdRmSz", "TR.TdRmSz2", "TR.TdRmSz3") + " else (case when RisRmSzDesc is null then " + GetRmSzDBFldNm("TR.TdRmSz", "TR.TdRmSz2", "TR.TdRmSz3") + " else RisRmSzDesc end) end "), _
            Array("'('+TdRmCd+') '+LTrim(case when " + wsPtrYN + "='Y' then " + GetRmSzDispFldNm("TdRmSz", "TdRmSz2", "TdRmSz3") + " else (case when RisRmSzDesc is null then " + GetRmSzDispFldNm("TdRmSz", "TdRmSz2", "TdRmSz3") + " else RisRmSzDesc end) end )", _
            "'('+FrRmCd+') '+LTrim(case when " + wsPtrYN + "='Y' then " + GetRmSzDispFldNm("FrRmSz", "FrRmSz2", "FrRmSz3") + " else (case when RisRmSzDesc is null then " + GetRmSzDispFldNm("FrRmSz", "FrRmSz2", "FrRmSz3") + " else RisRmSzDesc end) end )", _
            "'('+FrRmCd+') '+LTrim(case when " + wsPtrYN + "='Y' then " + GetRmSzDispFldNm("FrRmSz", "FrRmSz2", "FrRmSz3") + " else (case when RisRmSzDesc is null then " + GetRmSzDispFldNm("FrRmSz", "FrRmSz2", "FrRmSz3") + " else RisRmSzDesc end) end )", _
            "'('+TR.TdRmCd+') '+LTrim(case when " + wsPtrYN + "='Y' then " + GetRmSzDispFldNm("TR.TdRmSz", "TR.TdRmSz2", "TR.TdRmSz3") + " else (case when RisRmSzDesc is null then " + GetRmSzDispFldNm("TR.TdRmSz", "TR.TdRmSz2", "TR.TdRmSz3") + " else RisRmSzDesc end) end )"), _
            "TdRmCd,wRmSz", "hTdRmCd,hTdRmSz", "Rm Ctg,Rm SubCtg,Rm Code,BaseCd + RmCd,(RmCd) StkRt,Rm Sv/Sz,Rm Ptr,Rm StkRt", "", "", " '('+ RTrim(RmCtg)+ '/'+ RmSCtg+ ')' "
 
  ' ****** sachin 2.13.0 - 23/05/06 ****** '
  GRP_REP.Add "Range Size", "RrRngSz", "", "RrRngSz", "hRrRngSz", "", "", "", ""
  ' ****** sachin 2.13.0 - 23/05/06 ****** '
 
 
  GRP_REP.Add "Rm StkRt", Array("Str(TdRmStkRt,9,2)", "Str(FrRmStkRt,9,2)", "Str(FrRmStkRt,9,2)", "Str(TR.TdRmStkRt, 9, 2)"), "", "", "", "Rm Ptr,Rm Sv/Sz,(RmCd) Sv/Sz,(RmCd) StkRt", "", "", ""
  '****************Geeta**************Emr206*********12/04/04
            
  Dim wTdRate As String
  Dim wFrRate As String
  Dim wBkDtRate As String
    
 GRP_REP.Add "(RmCd) StkRt", Array("TdRmCd+ Str(TdRmStkRt, 9, 2)", "FrRmCd+ Str(FrRmStkRt, 9, 2)", "FrRmCd+ Str(FrRmStkRt, 9, 2)", "TR.TdRmCd+ Str(TR.TdRmStkRt, 9, 2)"), Array("'('+TdRmCd+') '+LTrim(str(TdRmStkRt,9,2))", "'('+FrRmCd+') '+LTrim(str(FrRmStkRt,9,2))", "'('+FrRmCd+') '+LTrim(str(FrRmStkRt,9,2))", "'('+TR.TdRmCd+') '+LTrim(str(TR.TdRmStkRt,9,2))"), "", "", "Rm Ctg,Rm SubCtg,Rm Code,BaseCd + RmCd,(RmCd) Sv/Sz,Rm Sv/Sz,Rm Ptr,Rm StkRt", "", "", " '('+ RTrim(RmCtg)+ '/'+ RmSCtg+')' "
    
 '4.1.4
 'BS error
    wTdRate = " (Case When RSCtg.PValue1='N' Then " + RmRate("RrCstRt", "TdRmSz", "TdRmSz2", "RmQwCst") + " Else TdRmStkRt End)"
    wFrRate = " (Case When RSCtg.PValue1='N' Then " + RmRate("RrCstRt", "FrRmSz", "FrRmSz2", "RmQwCst") + " Else FrRmStkRt End)"
    wBkDtRate = " (Case When RSCtg.PValue1='N' Then " + RmRate("RrCstRt", "TR.TdRmSz", "TR.TdRmSz2", "RmQwCst") + " Else TR.TdRmStkRt End)"
  GRP_REP.Add "(RmCd) Stk-CstRt", Array("TdRmCd+ Str(" + wTdRate + ", 9, 2)", "FrRmCd+ Str(" + wFrRate + ", 9, 2)", "FrRmCd+ Str(" + wFrRate + ", 9, 2)", "TR.TdRmCd+ Str(" + wBkDtRate + ", 9, 2)"), Array("'('+TdRmCd+') '+LTrim(str(" + wTdRate + ",9,2))", "'('+FrRmCd+') '+LTrim(str(" + wFrRate + ",9,2))", "'('+FrRmCd+') '+LTrim(str(" + wFrRate + ",9,2))", "'('+Str(" + wBkDtRate + ",9,2)+') '+LTrim(str(" + wBkDtRate + ",9,2))"), "", "", "Rm Ctg,Rm SubCtg,Rm Code,BaseCd + RmCd,(RmCd) Sv/Sz,Rm Sv/Sz,Rm Ptr,Rm SaleRt", "", "", " '('+ RTrim(RmCtg)+ '/'+ RmSCtg+')' "
  'BS error
    wTdRate = " (Case When RSCtg.PValue1='N' Then " + RmRate("RrSalRt", "TdRmSz", "TdRmSz2", "RmQw") + " Else TdRmStkRt End)"
    wFrRate = " (Case When RSCtg.PValue1='N' Then " + RmRate("RrSalRt", "FrRmSz", "FrRmSz2", "RmQw") + " Else FrRmStkRt End)"
    wBkDtRate = " (Case When RSCtg.PValue1='N' Then " + RmRate("RrSalRt", "TR.TdRmSz", "TR.TdRmSz2", "RmQw") + " Else TR.TdRmStkRt End)"
  GRP_REP.Add "(RmCd) Stk-SalRt", Array("TdRmCd+ Str(" + wTdRate + ", 9, 2)", "FrRmCd+ Str(" + wFrRate + ", 9, 2)", "FrRmCd+ Str(" + wFrRate + ", 9, 2)", "TR.TdRmCd+ Str(" + wBkDtRate + ", 9, 2)"), Array("'('+TdRmCd+') '+LTrim(str(" + wTdRate + ",9,2))", "'('+FrRmCd+') '+LTrim(str(" + wFrRate + ",9,2))", "'('+FrRmCd+') '+LTrim(str(" + wFrRate + ",9,2))", "'('+Str(" + wBkDtRate + ",9,2)+') '+LTrim(str(" + wBkDtRate + ",9,2))"), "", "", "Rm Ctg,Rm SubCtg,Rm Code,BaseCd + RmCd,(RmCd) Sv/Sz,Rm Sv/Sz,Rm Ptr,Rm SaleRt", "", "", " '('+ RTrim(RmCtg)+ '/'+ RmSCtg+')' "
  'BS error
    wTdRate = RmRate("RrCstRt", "TdRmSz", "TdRmSz2", "RmQwCst")
    wFrRate = RmRate("RrCstRt", "FrRmSz", "FrRmSz2", "RmQwCst")
    wBkDtRate = RmRate("RrCstRt", "TR.TdRmSz", "TR.TdRmSz2", "RmQwCst")
  GRP_REP.Add "(RmCd) CostRt", Array("TdRmCd+ Str(" + wTdRate + ", 9, 2)", "FrRmCd+ Str(" + wFrRate + ", 9, 2)", "FrRmCd+ Str(" + wFrRate + ", 9, 2)", "TR.TdRmCd+ Str(" + wBkDtRate + ", 9, 2)"), Array("'('+TdRmCd+') '+LTrim(str(" + wTdRate + ",9,2))", "'('+FrRmCd+') '+LTrim(str(" + wFrRate + ",9,2))", "'('+FrRmCd+') '+LTrim(str(" + wFrRate + ",9,2))", "'('+Str(" + wBkDtRate + ",9,2)+') '+LTrim(str(" + wBkDtRate + ",9,2))"), "", "", "Rm Ctg,Rm SubCtg,Rm Code,BaseCd + RmCd,(RmCd) Sv/Sz,Rm Sv/Sz,Rm Ptr,Rm CostRt", "", "", " '('+ RTrim(RmCtg)+ '/'+ RmSCtg+')' "
  'BS error
    wTdRate = RmRate("RrSalRt", "TdRmSz", "TdRmSz2", "RmQw")
    wFrRate = RmRate("RrSalRt", "FrRmSz", "FrRmSz2", "RmQw")
    wBkDtRate = RmRate("RrSalRt", "TR.TdRmSz", "TR.TdRmSz2", "RmQw")
  GRP_REP.Add "(RmCd) SaleRt", Array("TdRmCd+ Str(" + wTdRate + ", 9, 2)", "FrRmCd+ Str(" + wFrRate + ", 9, 2)", "FrRmCd+ Str(" + wFrRate + ", 9, 2)", "TR.TdRmCd+ Str(" + wBkDtRate + ", 9, 2)"), Array("'('+TdRmCd+') '+LTrim(str(" + wTdRate + ",9,2))", "'('+FrRmCd+') '+LTrim(str(" + wFrRate + ",9,2))", "'('+FrRmCd+') '+LTrim(str(" + wFrRate + ",9,2))", "'('+Str(" + wBkDtRate + ",9,2)+') '+LTrim(str(" + wBkDtRate + ",9,2))"), "", "", "Rm Ctg,Rm SubCtg,Rm Code,BaseCd + RmCd,(RmCd) Sv/Sz,Rm Sv/Sz,Rm Ptr,Rm SaleRt", "", "", " '('+ RTrim(RmCtg)+ '/'+ RmSCtg+')' "
     
  
  ' *** Manali 3.03 - RmSzDesc added
  If moCn.GetFldVal("Select HPtrYN from Head where HCoCd ='" + ctSelfCoCd + "' and HCd='" + ctSelfCmCd + "'") = "Y" Then _
     GRP_REP.Add "Rm Ptr", Array("IsNull(RisRmSzDesc, Str(TdRmSz,6,3))", "IsNull(RisRmSzDesc, Str(FrRmSz,6,3))", "IsNull(RisRmSzDesc, Str(FrRmSz,6,3))", "IsNull(RisRmSzDesc, Str(TR.TdRmSz,6,3))"), "", "", "", "Rm Sv/Sz,Rm StkRt,(RmCd) Sv/Sz,(RmCd) StkRt", "", "", ""
     'GRP_REP.Add "Rm Ptr", Array("Str(TdRmSz,6,3)", "Str(FrRmSz,6,3)", "Str(FrRmSz,6,3)", "Str(TR.TdRmSz,6,3)"), "", "", "", "Rm Sv/Sz,Rm StkRt,(RmCd) Sv/Sz,(RmCd) StkRt", "", "", ""
  '*** (Jen 2.14)
   
  GRP_REP.Add "Design", "BOdDmCd", "", "", "", "(DsgCd)Colour", "", "", ""    'Substring(DmDesc,1,120)
  ' ***** Manali 3.10.0 - 07/03/12 - DmCtg added in Group Sort
  GRP_REP.Add "Design Ctg", "DmCtg", "", "", "", "", "", "", "(Select PDesc from Param where PTyp= 'DMCTG' and PMCd= DmCtg) "
  
  GRP_REP.Add "(DsgCd)Colour", "BOdDmCd+OdDmCol", "'('+BOdDmCd+') '+OdDmCol", "", "", "Design", "", "", ""
  GRP_REP.Add "OrderNo", "BCoCd+BOdTc+BOdYy+BOdChr+str(BOdNo)", "BCoCd+'/'+BOdTc+'/'+BOdYy+'/'+BOdChr+'/'+LTrim(Str(BOdNo))", "", "", "Company Code,Customer", "", "", "'('+OmCmCd+ '/'+ OmPONo+ ')' "
'????? check
  ' Zubin 212 (Ord Kt added)
  GRP_REP.Add "OrderSr", "BCoCd+BOdTc+BOdYy+BOdChr+str(BOdNo)+str(BOdSr)", "BCoCd+'/'+BOdTc+'/'+BOdYy+'/'+BOdChr+'/'+LTrim(Str(BOdNo))+'/'+LTrim(Str(BOdSr))", "", "", "Company Code,OrderNo,Design,(DsgCd)Colour,Customer,Ord Kt", "", "", "'('+OmCmCd+ '/'+ OmPONo+ ')' "
  'GRP_REP.Add "OrderSr", "BOdTc+BOdYy+BOdChr+str(BOdNo)+str(BOdSr)", "BOdTc+'/'+BOdYy+'/'+BOdChr+'/'+LTrim(Str(BOdNo))+'/'+LTrim(Str(BOdSr))", "OmCmCd,OmPoNo", "hOmCmCd,hOmPoNo", "Customer Code,Customer PO", "", "", "'('+OmCmCd+ '/'+ OmPONo+ ')' "
  '******Geeta*****************Emr206****05/04/04*******
  '**********If UserLevel>6 then don't display the Customer Name
  'GRP_REP.Add "Customer", "OmCmCd", "", "", "", "", "", "", "(Select CmName From CustMst Where CmCtg= 'C' and CmCd= OmCmCd) "
  GRP_REP.Add "Customer", "OmCmCd", "", "", "", "", "", "", "(Select (case when (Select PNum from Param where Ptyp = 'USR' and PmCd = '" + gs_UsrCd + "')  <= 5 then  CmName else CmCd end) From CustMst Where CmCtg= 'C' and CmCd= OmCmCd) "
  
'  GRP_REP.Add "Design Ctg", "DmCtg", "", "DmCtg", "hDmCtg", "", "", "", "(Select PDesc from Param where PTyp= 'DMCTG' and PMCd= DmCtg) "
'  GRP_REP.Add "Design Code", "BOdDmCd", "", "BOdDmCd", "hBOdDmCd", "Design Ctg", "", "", "(Select Substring(DmDesc,1,120) From DsgMst Where DmTcTyp= 'DM' and DmCd= BOdDmCd and DmSz= '') "
'  GRP_REP.Add "Ord Kt", "OdKt", "", "OdKt", "hOdKt", "", "", "", "(Select PDesc from Param where PTyp= 'KT' and PMCd= OdKt) "
'  GRP_REP.Add "OrderNo", "BOdTc+BOdYy+BOdChr+str(BOdNo)", "BOdTc+'/'+BOdYy+'/'+BOdChr+'/'+LTrim(Str(BOdNo))", "wOrdNo,OmCmCd,OmPoNo", "hOrdNo,hOmCmCd,hOmPoNo", "Customer Code,Customer PO", "", "", "'('+OmCmCd+ '/'+ OmPONo+ ')' "
'  GRP_REP.Add "Customer Code", "OmCmCd", "", "OmCmCd", "hOmCmCd", "", "", "", "(Select CmName From CustMst Where CmCtg= 'C' and CmCd= OmCmCd) "
'  GRP_REP.Add "Customer PO", "OmCmCd+OmPONo", "'('+ OmCmCd+ ') '+ LTrim(OmPONo)", "OmCmCd,OmPoNo", "hOmCmCd,hOmPONo", "", "", "", "BOdTc+'/'+BOdYy+'/'+BOdChr+'/'+LTrim(Str(BOdNo))"
  
  If UCase(adc.MenuCd) = UCase("RepWipFgBRm") Then
    GRP_REP.Add "Pck List", "FdPlYy+FdPlChr+str(FdPlNo)", "FdPlYy+'/'+FdPlChr+'/'+LTrim(Str(FdPlNo))", "", "", "", "", "", ""
    '*** (Jen 2.12 Patch 3)
    GRP_REP.Add "FgSubLoc", Array("BFgSubLoc", "BFgSubLoc", "qFgSubLoc"), "", "", "", "Company Code", "", "", Array("(Select vPDesc From vParam Where vPCoCd = BCoCd and vPTyp= 'FGSUBLOC' and vPMCd= BFgSubLoc) ", "(Select vPDesc From vParam Where vPCoCd = BCoCd and vPTyp= 'FGSUBLOC' and vPMCd= BFgSubLoc) ", "(Select vPDesc From vParam Where vPCoCd = BCoCd and vPTyp= 'FGSUBLOC' and vPMCd= qFgSubLoc) ")
    '*** (Jen 2.12 Patch 3)
    GRP_REP.Add "FG / INV", " (Case when exists (Select 'x' From InvFgd where IfBIdNo = BIdNo and IfFdBYy = BYy and IfFdBChr = BChr and IfFdBNo = BNo and IfCoCd = BCoCd) then 'INV' else 'FG' end)", " (Case when exists (Select 'x' From InvFgd where IfBIdNo = BIdNo and IfFdBYy = BYy and IfFdBChr = BChr and IfFdBNo = BNo and IfCoCd = BCoCd) then 'INV' else 'FG' end)", "", "", "", "", "", ""

  End If
  '*** (Bef 2.12 Patch 3) GRP_REP.Add "FgSubLoc", "BFgSubLoc", "", "", "", "", "", "", "(Select vPDesc From vParam Where vPCoCd = BCoCd and vPTyp= 'FGSUBLOC' and vPMCd= BFgSubLoc) "
  
  '**** Shilpa Emr 206 ****
  GRP_REP.Add "BagNo", "BCoCd+BYy+BChr+str(BNo)", "BCoCd+'/'+BYy+'/'+BChr+'/'+LTrim(Str(BNo))", "", "", "Company Code,Design Code,Bag Type,Bag Character", "", "", ""
  '**** Shilpa Emr 206 ****
  
  '****** Sachin 2.13.0 - [07-07-2006] ******
  GRP_REP.Add "Stone Class", "RmCtg+IsNull(SCls.PSCd, '')", "'('+ RTrim(RmCtg)+') '+IsNull(SCls.PSCd, '')", "RmCtg", "hRmCtg", "Rm Ctg", "", "", "SCls.PDesc"
  '****** Sachin 2.13.0 - [07-07-2006] ******
  
  ' **** Zubin 211 **** '
  'If UCase(adc.MenuCd) = UCase("RepWipBRm") Then   '*** (Bef 2.14)
  '*** (Jen 2.14)
  'Uma *** new 'D' btyp
  If UCase(adc.MenuCd) = UCase("RepWipBRm") Or UCase(adc.MenuCd) = UCase("RepWipBRmBkDt") Then
    GRP_REP.Add "Bag Type", "(case BTyp.PValue when 'F' then 'Flute Bag' when 'P' then 'Primary Bag' when 'C' then 'Component Bag' when 'D' then 'PD Bag' end) ", _
      "", "", "", "Bag Typ+Chr", "", "", ""
    GRP_REP.Add "Bag Character", "BChr", "BChr", "", "", "Bag Type,Bag Typ+Chr", "", "", "BTyp.PDesc"
    GRP_REP.Add "Bag Typ+Chr", "(case BTyp.PValue when 'F' then 'Flute Bag' when 'P' then 'Primary Bag' when 'C' then 'Component Bag' when 'D' then 'PD Bag' end)+ BChr", _
      "(case BTyp.PValue when 'F' then 'Flute Bag' when 'P' then 'Primary Bag' when 'C' then 'Component Bag' when 'D' then 'PD Bag' end)+'/ '+ BChr", "", "", "Bag Type,Bag Character", "", "", "BTyp.PDesc"
    
    GRP_REP.Add "Worker", Array("BWrk", "BWrk", "BWrk", "qTdWrk"), "", "", "", "", "", "", Array("(Select w.vPDesc From vParam w Where w.vPCoCd= BCoCd and w.vPTyp= 'WORK' and w.vPMCd= BWrk) ", "(Select w.vPDesc From vParam w Where w.vPCoCd= BCoCd and w.vPTyp= 'WORK' and w.vPMCd= BWrk) ", "(Select w.vPDesc From vParam w Where w.vPCoCd= BCoCd and w.vPTyp= 'WORK' and w.vPMCd= BWrk) ", "(Select w.vPDesc From vParam w Where w.vPCoCd= BCoCd and w.vPTyp= 'WORK' and w.vPMCd= IsNull(qTdWrk, '')) ")
    '*** (Jen 2.14)
    
        
  End If
  ' Zubin 212
  '****** Sachin 2.14.0 - [MultiCompany] - All companies considered for Lot NO Group sort
  
  'Dim wb_CoCdSel As Boolean
  'If adc("UoInExpNoSel") <> "" Then
    'Cancel = Not (ChkVal("Select 'x' from InvHd where InExpNo in (" + pv_NewValue + ")", ErrMsg))
    'If Cancel = True Then ErrMsg = "Invalid Inv Exp Number(s) Scope": Exit Sub
  'End If
  
  '*** (Bef 2.14)
  'If moCn.GetFldVal("select Max(HLotNoYn) from Head where HCd= '" + ctSelfCmCd + "' " + IIF(adc("UoCoCdFr") <> "", " and HCoCd>= '" + adc("UoCoCdFr") + "' ", " ") + IIF(adc("UoCoCdTo") <> "", " and HCoCd<= '" + adc("UoCoCdTo") + "' ", " ") + IIF(adc("UoCoCdSel") <> "", " and HCoCd In (" + adc("UoCoCdSel") + ") ", " ")) = "Y" Then
  '  GRP_REP.Add "RM Lot No.", Array("TdLotNo", "FrLotNo", "FrLotNo"), "", "TdLotNo", "hLotNo", "", "", "", Array("TdLotNo", "FrLotNo", "FrLotNo")
  'End If
  '*** (Bef 2.14)
  
  '*** (Jen 2.14)
  If (moCn.GetFldVal("select Max(HLotNoYn) from Head where HCd= '" + ctSelfCmCd + "' " + IIF(adc("UoCoCdFr") <> "", " and HCoCd>= '" + adc("UoCoCdFr") + "' ", " ") + IIF(adc("UoCoCdTo") <> "", " and HCoCd<= '" + adc("UoCoCdTo") + "' ", " ") + IIF(adc("UoCoCdSel") <> "", " and HCoCd In (" + adc("UoCoCdSel") + ") ", " ")) = "Y" And ws_MultiCoMod = False) Or (ws_MultiCoMod = True) Then
    GRP_REP.Add "RM Lot No.", Array("TdLotNo", "FrLotNo", "FrLotNo", "TR.TdLotNo"), "", "TdLotNo", "hLotNo", "", "", "", Array("TdLotNo", "FrLotNo", "FrLotNo", "TR.TdLotNo")
  End If
  '*** (Jen 2.14)
  
  GRP_REP.Add "Ord Kt", "OdKt", "", "", "", "", "", "", "(Select PDesc from Param where PTyp= 'KT' and PMCd= OdKt) "
  ' Zubin 212
  ' **** Zubin 211 **** '

  '****** Sachin 2.14.0 - [23-11-2006] - [Multi-Company] - ,Process,OrderNo,OrderSr,Customer PO,FgSubLoc ******
  If ws_MultiCoMod = True Then
    GRP_REP.Add "Company Code", "BCoCd", "BCoCd", "", "", "", "", "", "(Select HName from Head where HCd= '" + ctSelfCmCd + "' and HCoCd=OdCoCd) "
    If UCase(adc.MenuCd) = UCase("RepWipBRm") Then _
      GRP_REP.Add "Parent Location", "Lc.LocPrntCd", "Lc.LocPrntCd", "", "", "", "", "", "(Select LocDesc from Loc Ln where Ln.LocCoCd= '" + ctSelfCoCd + "' and Ln.LocCd= Lc.LocPrntCd) "
  End If
  '****** Sachin 2.14.0 - [23-11-2006] - [Multi-Company] - ******

End Sub
Private Sub adc_setreprecsource()
'***Geeta***Emr208***New Flds DmTcTyp, BOdDmCd, BOdDmSz and DsgMst Join to Display Picture
  '*** Report Sql ***
  ' Zubin 212 (OrdDsg was not included)
  '*** Jay 3.2.0 [PicNm]
  ' **** Manali 3.03 - qRmSzDesc Added
'  Select space(30) as Grp1, space(30) as Grp2, space(30) as Grp3,
'  space(30) as DGrp1, space(30) as DGrp2, space(30) as DGrp3,
'  space(30) as G1Desc, space(30) as G2Desc, space(30) as G3Desc,
'  BCoCd, BYy, BChr, BNo, DmTcTyp, BOdDmCd, BOdDmSz, BLoc, RmCtg, PSCd As StnCls, RmSCtg, TdRmCd, TdLotNo, TdRmSz, '' as qRmSzDesc, RrRngSz, TdRmStkRt,
'  PValue3 as qPtrYN, TdRmQty, TdRmWt, 0 as qMul, TdRmWt as qGmWt,
'  TdRmWt as qGldWt, BQty , BPtQty, TDt, TdTc, TdYy, TdChr, TdNo,
'  TdSr, TdSrNo, TdRmStkRt as qCalcRt, RmQW, RmQWCst, OdKt, OdPicNm
'  From Bag, Txnd, Txn, RmMst, RmRt, DsgMst, Param, OrdDsg where 1= 2
  '*** Report Sql ***
  
  Dim grpflds0 As String, grpflds1 As String, wrepcnd As String, wCnd As String, wSqlStrg As String
  Dim wPurityWt As String, wQty As String, wPtQty As String, wSelCnd As String, wFgCnd As String
  Dim wPureWtOrZWt As String, wXPureWtOrZwt As String
  Dim wSqlStrg1 As String, wSqlStrg2 As String, wFgSel As String, wFgJoin As String
  Dim wTdRate As String, wFrRate As String, wBagClsCnd As String, wRmQWStr As String
  Dim ws_FGINVChk As String
  ws_FGINVChk = ""
  
  '*** Jen (2.12 Patch)
  Dim ws_AddFgCnd As String
  '*** Jen (2.12 Patch)
  
  ' **** Zubin 211 **** '
  Dim ws_BagJoin As String
  ' **** Zubin 211 **** '
  
  '*** Manali 3.03 - 07/07/08 - RmSzDesc
  Dim wRmSzDesc As String
  '*** Manali 3.03 - 07/07/08 - RmSzDesc
  
  ' ***** Manali 3.5.0 - 08/11/08 -  Bag Delay Days Cond Added
  Dim wBagDelayCnd As String
  ' ***** Manali 3.5.0 - 08/11/08 -  Bag Delay Days Cond Added
  
  Dim ws_FgTbl As String, ws_TmpTblCnd As String   '*** (Jen 2.12 Patch 3)
  
  Dim wSqlStrg3 As String, wBkDtRate As String, grpflds2 As String     '*** (Jen 2.14)
  
  Dim wAsOfDt As String   ' ***** Manali 3.4.1 - 02/12/08
  
  '****** Sachin 3.02.0 - Check on YYYYMM Entered by user ******
  '****** Manali 3.4.1 - 02/12/08 - "01/01/80" excluded from the check
  Dim ws_PurgeDt As String
  'sv.9 getting max(PurgeDate) from vParam for given company code
  ws_PurgeDt = moCn.GetFldVal("Select max(convert(date,vPDesc225,3)) from vParam " + _
                          " where vPCoCd >='" + adc("UoCoCdFr") + "' and " + _
                          " vPCoCd <='" + IIF(adc("UoCoCdTo") <> "", adc("UoCoCdTo"), "ZZZ") + "'" + _
                          IIF(adc("UoCoCdSel") <> "", " and vPCoCd In(" + adc("UoCoCdSel") + ")", "") + _
                          " and vPTyp= 'PURGEDT' and vPMCd= 'PURGEDT' and vPSCd =''")
  
  If ws_PurgeDt = "" Then ws_PurgeDt = "01/01/80"
  
  
  'sv.9 gs_PurgeDt replaced with ws_PurgeDt
  If gs_Partition = ctCurrPrtn And CDate(ws_PurgeDt) <> "01/01/80" And _
      (UCase(adc.MenuCd) = UCase("RepWipBRmBkDt") And _
      CDate(adc("UOASOFDT")) <= CDate(ws_PurgeDt) And CDate(adc("UoAsOfDt")) <> MWLib.EmptyDate) Then
      DispMsg "From/To/As Of Date Can Be Either [01/01/80] Or Greater Than Purge Date [" + CStr(CDate(ws_PurgeDt)) + "]", etError: Exit Sub
      'DispMsg "From/To/As Of Date Cannot Be Less Than Purge Date [" + CStr(gs_PurgeDt) + "]", etError: Exit Sub
  End If
  
  If (adc("UoStnClsFr") <> "" Or adc("UoStnClsTo") <> "" Or adc("UoStnClsSel") <> "") And ((adc("UoRmCtgFr") <> adc("UoRmCtgTo")) Or (adc("UoRmCtgFr") = "")) Then
      DispMsg "Please select an RmCtg to provide a filter of Stone Class", etError: Exit Sub
  End If
  
'  If gs_Partition = ctCurrPrtn And gs_PurgeDt <> "01/01/80" And _
'      (UCase(adc.MenuCd) = UCase("RepWipFgBRm") And (CDate(adc("UODTFR")) <= CDate(gs_PurgeDt) Or CDate(adc("UODTTO")) <= CDate(gs_PurgeDt))) Or _
'      (UCase(adc.MenuCd) = UCase("RepWipBRmBkDt") And CDate(adc("UOASOFDT")) <= CDate(gs_PurgeDt)) Then
'      DispMsg "From/To/As Of Date Cannot Be Less Than Purge Date [" + CStr(gs_PurgeDt) + "]", etError: Exit Sub
'  End If
  '****** Sachin 3.02.0 - Check on YYYYMM Entered by user ******
 
  Set Rep = Nothing
  Set adc.RepSource = Rep
  Call DispCoNm
  Call SetMwName(Rep)   ' ***** Manali 3.6.1 - Microway footer
  Call GRP_REP.Gen3LRep
  'grpflds = GRP_REP.GrpFldLst
  
  grpflds0 = GRP_REP.GrpFldLst(0)
  '*** (Bef 2.12 Patch 3) grpflds1 = GRP_REP.GrpFldLst(1)
  
  '*** (Jen 2.12 Patch 3)
  If adc("UoDtFr") = "01/01/80" And adc("UoDtTo") = "01/01/80" Then
    grpflds1 = GRP_REP.GrpFldLst(1)
  Else
    grpflds1 = GRP_REP.GrpFldLst(2)
  End If
  '*** (Jen 2.12 Patch 3)
  
  grpflds2 = GRP_REP.GrpFldLst(3)   '*** (Jen 2.14)
  
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
  
 If UCase(adc("UoYN3")) = UCase("Y") Then
    Rep.h1PureWt.SetText ("Z Pure Wt")
    Rep.h2PureWt.SetText ("Z Pure Wt")
    wPureWtOrZWt = "RmPurityZ"
    wXPureWtOrZwt = "(Select XZ.RmPurityZ from RmMst XZ where XZ.RmCd=Rm.RmBaseCd " + IIF(gs_Partition = ctCurrPrtn, " And XZ.RmPrtKey=Rm.RmPrtKey)", ")")
  ElseIf UCase(adc("UoYN3")) = UCase("N") Then
    wPureWtOrZWt = "RmPurityWt"
    wXPureWtOrZwt = "RmBasePurityWt"
  End If

  If UCase(adc("UoPMCdFr")) = "STKVAL" Then
    'BS error
    wTdRate = IIF(UCase(adc("UoPMCdTo")) = "CSTVAL", " (Case When RSCtg.PValue1='N' Then " + RmRate("RrCstRt", "TdRmSz", "TdRmSz2", "RmQwCst") + " Else TdRmStkRt End)", _
                IIF(UCase(adc("UoPMCdTo")) = "SALVAL", " (Case When RSCtg.PValue1='N' Then " + RmRate("RrSalRt", "TdRmSz", "TdRmSz2", "RmQw") + " Else TdRmStkRt End)", " TdRmStkRt ")) + " as qCalcRt "
    'BS error
    wFrRate = IIF(UCase(adc("UoPMCdTo")) = "CSTVAL", " (Case When RSCtg.PValue1='N' Then " + RmRate("RrCstRt", "FrRmSz", "FrRmSz2", "RmQwCst") + " Else FrRmStkRt End)", _
                IIF(UCase(adc("UoPMCdTo")) = "SALVAL", " (Case When RSCtg.PValue1='N' Then " + RmRate("RrSalRt", "FrRmSz", "FrRmSz2", "RmQw") + " Else FrRmStkRt End)", " FrRmStkRt ")) + " as qCalcRt "   '" FrRmStkRt as qCalcRt "
    'BS error
    wBkDtRate = IIF(UCase(adc("UoPMCdTo")) = "CSTVAL", " (Case When RSCtg.PValue1='N' Then " + RmRate("RrCstRt", "TR.TdRmSz", "TR.TdRmSz2", "RmQwCst") + " Else TR.TdRmStkRt End)", _
                IIF(UCase(adc("UoPMCdTo")) = "SALVAL", " (Case When RSCtg.PValue1='N' Then " + RmRate("RrSalRt", "TR.TdRmSz", "TR.TdRmSz2", "RmQw") + " Else TR.TdRmStkRt End)", " TR.TdRmStkRt ")) + " as qCalcRt "   '" TR.TdRmStkRt as qCalcRt "
    
    Rep.hTdRmStkRt.SetText IIF(UCase(adc("UoPMCdTo")) = "CSTVAL", "Stk/Cst Rt", IIF(UCase(adc("UoPMCdTo")) = "SALVAL", "Stk/Sal Rt", "Stock Rt"))
    Rep.TxtValTyp.SetText IIF(UCase(adc("UoPMCdTo")) = "CSTVAL", "(Value=Stk-Cst)", IIF(UCase(adc("UoPMCdTo")) = "SALVAL", "(Value=Stk-Sal)", "(Value = Stock)"))
    GRP_REP.SetFormula Rep, "wRmVal", "{rdo.TdRmWt}* {rdo.qMul}* {rdo.qCalcRt}"
    
  ElseIf UCase(adc("UoPMCdFr")) = "CSTVAL" Then
    'BS error
    wTdRate = RmRate("RrCstRt", "TdRmSz", "TdRmSz2", "RmQwCst") + " as qCalcRt "
    wFrRate = RmRate("RrCstRt", "FrRmSz", "FrRmSz2", "RmQwCst") + " as qCalcRt "
    wBkDtRate = RmRate("RrCstRt", "TR.TdRmSz", "TR.TdRmSz2", "RmQwCst") + " as qCalcRt "   '*** (Jen 2.14)
    Rep.TxtValTyp.SetText "(Value = Cost)"
    Rep.hTdRmStkRt.SetText "Cost Rt"
    
    If Round(adc("UoPrc"), 3) = 0 Then
      GRP_REP.SetFormula Rep, "wRmVal", "(IF UpperCase({rdo.RmQWCst}) = UpperCase(""Q"") Then {@wRmQty} Else (If UpperCase({rdo.RmCtg}) = UpperCase(""X"") Then {@wGmWt} Else (if Round({@wPureWt}, 3)<> 0 then {@wPureWt} Else {@wRmWt}))) * {rdo.qCalcRt}"
    Else
      GRP_REP.SetFormula Rep, "wRmVal", "(IF UpperCase({rdo.RmQWCst}) = UpperCase(""Q"") Then {@wRmQty} Else (If UpperCase({rdo.RmCtg}) = UpperCase(""X"") Then {@wGmWt} Else (if Round({@wPureWt}, 3)<> 0 then {@wPureWt} Else {@wRmWt}))) * {rdo.qCalcRt} * (1 + (" + CStr(adc("UoPrc")) + "/ 100))"
    End If
    
  ElseIf UCase(adc("UoPMCdFr")) = "SALVAL" Then
    'BS error
    wTdRate = RmRate("RrSalRt", "TdRmSz", "TdRmSz2", "RmQw") + " as qCalcRt "
    wFrRate = RmRate("RrSalRt", "FrRmSz", "FrRmSz2", "RmQw") + " as qCalcRt "
    wBkDtRate = RmRate("RrSalRt", "TR.TdRmSz", "TR.TdRmSz2", "RmQw") + " as qCalcRt "  '*** (Jen 2.14)
    
    Rep.TxtValTyp.SetText "(Value = Sale)"
    Rep.hTdRmStkRt.SetText "Sale Rt"
    If Round(adc("UoPrc"), 3) = 0 Then
      GRP_REP.SetFormula Rep, "wRmVal", "(IF UpperCase({rdo.RmQW}) = UpperCase(""Q"") Then {@wRmQty} Else (If UpperCase({rdo.RmCtg}) = UpperCase(""X"") Then {@wGmWt} Else (if Round({@wPureWt}, 3)<> 0 then {@wPureWt} Else {@wRmWt}))) * {rdo.qCalcRt}"
    Else
      GRP_REP.SetFormula Rep, "wRmVal", "(IF UpperCase({rdo.RmQW}) = UpperCase(""Q"") Then {@wRmQty} Else (If UpperCase({rdo.RmCtg}) = UpperCase(""X"") Then {@wGmWt} Else (if Round({@wPureWt}, 3)<> 0 then {@wPureWt} Else {@wRmWt}))) * {rdo.qCalcRt} * (1 + (" + CStr(adc("UoPrc")) + "/ 100))"
    End If
 ElseIf (UCase(adc("UoPMCdFr")) = "LOTV-BAS" Or UCase(adc("UoPMCdFr")) = "LOTV-INR") Then
    'BS error
    wTdRate = RmRate("RrSalRt", "TdRmSz", "TdRmSz2", "RmQw", IIF(UCase(adc("UoPMCDFr")) = "LOTV-INR", "RS", ""), True) + " as qCalcRt "
    wFrRate = RmRate("RrSalRt", "FrRmSz", "FrRmSz2", "RmQw", IIF(UCase(adc("UoPMCDFr")) = "LOTV-INR", "RS", ""), True) + " as qCalcRt "
    wBkDtRate = RmRate("RrSalRt", "TR.TdRmSz", "TR.TdRmSz2", "RmQw", IIF(UCase(adc("UoPMCDFr")) = "LOTV-INR", "RS", ""), True) + " as qCalcRt "  '*** (Jen 2.14)
    If UCase(adc("UoPMCdFr")) = "LOTV-INR" Then
      Rep.TxtValTyp.SetText "(Value = LINR)"
      Rep.hTdRmStkRt.SetText "Lot INR"
    Else
      Rep.TxtValTyp.SetText "(Value = LBAS)"
      Rep.hTdRmStkRt.SetText "Lot BAS"
    End If
    If Round(adc("UoPrc"), 3) = 0 Then
      GRP_REP.SetFormula Rep, "wRmVal", "(IF UpperCase({rdo.RmQW}) = UpperCase(""Q"") Then {@wRmQty} Else (If UpperCase({rdo.RmCtg}) = UpperCase(""X"") Then {@wGmWt} Else (if Round({@wPureWt}, 3)<> 0 then {@wPureWt} Else {@wRmWt}))) * {rdo.qCalcRt}"
    Else
      GRP_REP.SetFormula Rep, "wRmVal", "(IF UpperCase({rdo.RmQW}) = UpperCase(""Q"") Then {@wRmQty} Else (If UpperCase({rdo.RmCtg}) = UpperCase(""X"") Then {@wGmWt} Else (if Round({@wPureWt}, 3)<> 0 then {@wPureWt} Else {@wRmWt}))) * {rdo.qCalcRt} * (1 + (" + CStr(adc("UoPrc")) + "/ 100))"
    End If
  Else
    wTdRate = " 0 as qCalcRt "
    wFrRate = " 0 as qCalcRt "
    wBkDtRate = " 0 as qCalcRt "  '*** (Jen 2.14)
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
  '*** This Inner Select Gives An Error
  
  ' ***** Manali 3.03 - 24/06/08 - Pd Mod - 'L' added
  wPurityWt = " /(case when Rm.RmCtg in ('G', 'P', 'S', 'L', 'M') then " + _
                          "IsNull((Select " + wPureWtOrZWt + " from Param, RmMst " + _
                          "where " + IIF(gs_Partition = ctCurrPrtn, " RmPrtKey='" + ctCurrPrtn + "' and ", "") + " RmCd= PValue and PTyp= 'RMCTG' and PMCd= Rm.RmCtg), 1) " + _
                "when Rm.RmCtg= 'X' then " + _
                          "IsNull((Select " + wPureWtOrZWt + " from Param, RmMst Base " + _
                          "where " + IIF(gs_Partition = ctCurrPrtn, " RmPrtKey='" + ctCurrPrtn + "' and ", "") + " RmCd= PValue and PTyp= 'RMCTG' and " + _
                          "PMCd= IsNull((Select RmCtg From RmMst where " + IIF(gs_Partition = ctCurrPrtn, " RmPrtKey='" + ctCurrPrtn + "' and ", "") + " RmCd= Rm.RmBaseCd), '') ), 1) " + _
                "else 1 end) "
  
  '*** Selection condition for normal or rejected pieces ('N'- Normal, 'R'-Rejected) ***
  '### wSelCnd = IIf(adc("UoBagPcs") = "N", " and (BRjQty= 0 and BRjPtQty= 0) ", " and (BRjQty<> 0 or BRjPtQty<> 0) ")
  'wSelCnd = IIf(adc("UoBagPcs") = "N", " and (BQty<> 0 or BPtQty<> 0 or (BQty= 0 and BPtQty= 0 and BRjQty= 0 and BRjPtQty= 0 and Round(BGrWt, 3)<> 0)) ", " and (BRjQty<> 0 or BRjPtQty<> 0) ")
  '### wSelCnd = IIf(ADC("UoBagPcs") = "N", " and (BQty<> 0 or BPtQty<> 0) ", " and (BRjQty<> 0 or BRjPtQty<> 0) ")   '(This condition should have given the same result as above but it does not give)
           
  '*** Manali 3.03 - 07/07/08 - RmSzDesc
  wRmSzDesc = "case when (Select PDesc225 from Param where PTyp= 'RMSCTG' and PMCd= IsNull(RisRmCtg, '') and PSCd= RmSCtg)= 'Y' then IsNull(RisRmSzDesc, '') Else '' End "
  '*** Manali 3.03 - 07/07/08 - RmSzDesc
           
  If UCase(adc.MenuCd) = UCase("RepWipBRm") Then
    wQty = IIF(adc("UoBagPcs") = "N", " BQty as BQty", " BRjQty as BQty")
    wPtQty = IIF(adc("UoBagPcs") = "N", " BPtQty as BPtQty", " BRjPtQty as BPtQty")
       
    '*** Selection condition for normal or rejected pieces ('N'- Normal, 'R'-Rejected) ***
    wSelCnd = IIF(adc("UoBagPcs") = "N", " (BQty<> 0 or BPtQty<> 0 or (BQty= 0 and BPtQty= 0 and BRjQty= 0 and BRjPtQty= 0 and Round(BGrWt, 3)<> 0)) ", " (BRjQty<> 0 or BRjPtQty<> 0) ")
    wFgCnd = " and BLoc<> 'PFG' "
    wFgSel = "(Case when FrFrRmLoc='W' and FrFrRmDc='D' then 1 " + _
              "when FrFrRmLoc='W' and FrFrRmDc='C' then -1 " + _
              "when FrToRmLoc='W' and FrToRmDc='D' then 1 " + _
              "when FrToRmLoc='W' and FrToRmDc='C' then -1 " + _
              "Else 0 end) as qMul, "
    wFgJoin = ""
    If adc("UoBagPcs") = "N" Then Rep.TxtHead.SetText "WIP Bag Raw Material Report (Normal)"
    If adc("UoBagPcs") = "R" Then Rep.TxtHead.SetText "WIP Bag Raw Material Report (Rejected)"
    Rep.LblFgDtFr.Suppress = True
    Rep.LblFgDtTo.Suppress = True
    Rep.TxtFgDtTo.Suppress = True
    Rep.TxtFgDtFr.Suppress = True
    ' Zubin 212 (Patch01 11/03/06)
    Rep.TxtLocFr.SetText adc("UoIssLocFr")
    Rep.TxtLocTo.SetText adc("UoIssLocTo")
    'Rep.TxtLocSel.SetText adc("UoIssLocSel") 22 Loc sel removed from header
    ' Zubin 212 (Patch01 11/03/06)
        
    '*** (Jen 2.12 Patch)
    wBagClsCnd = IIF(adc("UoYN1") = "Y", " and BCls= 'N' ", IIF(adc("UoYN1") = "N", " and BCls= 'Y' ", ""))
    '*** (Jen 2.12 Patch)
    
    '*** (Jen 2.14)
  ElseIf UCase(adc.MenuCd) = UCase("RepWipBRmBkDt") Then
    wQty = IIF(adc("UoBagPcs") = "N", " BOpnQty+ qTdBQty as BQty", " qTdRjQty as BQty")
    wPtQty = IIF(adc("UoBagPcs") = "N", " BOpnPtQty+ qTdBPtQty as BPtQty", " qTdRjPtQty as BPtQty")
    wSelCnd = IIF(adc("UoBagPcs") = "N", " and qTTcTyp<> 'MLT' and Not (qTTcTyp= 'REJ' and qTdBDC= 'C') and ((BOpnQty+ qTdBQty)<> 0 Or (BOpnPtQty+ qTdBPtQty)<> 0 Or (BOpnQty+ qTdBQty= 0 and BOpnPtQty+ qTdBPtQty= 0 and qTdRjQty= 0 and qTdRjPtQty= 0 and qTdBGrWt<> 0)) ", " and (qTTcTyp= 'REJ' and qTdBDC= 'C') and (qTdRjQty<> 0 or qTdRjPtQty<> 0) ")
  
    If adc("UoBagPcs") = "N" Then Rep.TxtHead.SetText "WIP Bag RM Report (Normal) As On " + CStr(adc("UoAsOfDt"))
    If adc("UoBagPcs") = "R" Then Rep.TxtHead.SetText "WIP Bag RM Report (Rejected) As On " + CStr(adc("UoAsOfDt"))
    Rep.LblFgDtFr.Suppress = True
    Rep.LblFgDtTo.Suppress = True
    Rep.TxtFgDtTo.Suppress = True
    Rep.TxtFgDtFr.Suppress = True
    Rep.TxtLocFr.SetText adc("UoIssLocFr")
    Rep.TxtLocTo.SetText adc("UoIssLocTo")
    'Rep.TxtLocSel.SetText adc("UoIssLocSel") 22 Loc sel removed from header
  '*** (Jen 2.14)
    
  ElseIf UCase(adc.MenuCd) = UCase("RepWipFgBRm") Then
    '*** Jen (2.12 Patch)
    If adc("UoDtTo") = "01/01/80" Then
    '*** Jen (2.12 Patch)
      wQty = " BQty as BQty"
      wPtQty = " BPtQty as BPtQty"
      wSelCnd = " "
      wFgCnd = " BLoc= 'PFG' "
      '*** (Jen 2.12 Patch)
      wBagClsCnd = IIF(adc("UoYN1") = "Y", " and BCls= 'N' ", IIF(adc("UoYN1") = "N", " and BCls= 'Y' ", ""))
      '*** (Jen 2.12 Patch)
      
    '*** Jen (2.12 Patch)
    Else
      wQty = " FdQty as BQty"
      wPtQty = " 0 as BPtQty"
      wSelCnd = " "
      '*** (Bef 2.12 Patch 3) wFgCnd = " exists(Select 'x' from Param where PTyp= 'TC' and PMCd= FgTc and PValue= 'FB') "
      '*** (Jen 2.12 Patch 3)
      wFgCnd = " exists(Select 'x' from Param where PTyp= 'TC' and PMCd= FdTc and PValue= 'FB') "
      '*** (Jen 2.12 Patch 3)
      If adc("UoYN1") = "Y" Then
        'wBagClsCnd = " and (BCls= 'N' or BClsDt >= '" + CStr(adc("UoDtTo")) + "') "      '*** (Bef 2.14 Next)
        wBagClsCnd = " and (BCls= 'N' or BClsDt > '" + CStr(adc("UoDtTo")) + "') "      '*** (Jen 2.14 Next)
      ElseIf adc("UoYN1") = "N" Then
        wBagClsCnd = " and (BCls= 'Y' and BClsDt <= '" + CStr(adc("UoDtTo")) + "') "  ' ***** Manali 3.04 - 13/10/08 - [BCls= 'Y'] added
      Else
        wBagClsCnd = " "
      End If
    End If
    
    If adc("UoYn4") <> "" Then
        ws_FGINVChk = IIF(adc("UoYn4") = "Y", " Not ", "") + _
            "exists (Select 'x' From InvFgd where IfBIdNo = BIdNo and " + _
            "IfFdBYy = BYy and IfFdBChr = BChr and IfFdBNo = BNo and IfCoCd = BCoCd)"
    End If

    If adc("UoDtFr") = "01/01/80" And adc("UoDtTo") = "01/01/80" Then
      'ws_AddFgCnd = " "  '*** (Bef 2.12 Patch 3)
      
      '*** (Bef 2.14 Next Patch 1)
      'ws_FgTbl = " Join Fgd Fd on FdCoCd= BCoCd and FdBYy= BYy and FdBChr= BChr and FdBNo= BNo " + _
                 "Join Fg F on FgCoCd= FdCoCd and FgTc= FdTc and FgYy= FdYy and FgChr= FdChr and FgNo= FdNo "
      '*** (Bef 2.14 Next Patch 1)
      
      '*** (Jen 2.14 Next Patch 1 23/05/07)
      '****** Sachin 3.02 - Id fields In Join
      ws_FgTbl = " Join Fgd Fd on FdCoCd= BCoCd and FdBYy= BYy and FdBChr= BChr and FdBNo= BNo " + _
                 "Join Fg F on FgIdNo=FdFgIdNo and FgCoCd= FdCoCd and FgTc= FdTc and FgYy= FdYy and FgChr= FdChr and FgNo= FdNo " + IIF(gs_Partition = ctCurrPrtn, " and FGPrtKey=FdPrtKey ", "") + _
                 "    and exists (select 'x' from Param where PTyp= 'TC' and PMCd= FdTc and PValue= 'FB') " + _
                 "    and (FdDespQty= 0 or " + _
                 "         exists (Select 'x' From InvFgd where IfFdIdNo=FdIdNo and IfCoCd= FdCoCd and IfFdTc= FdTc and IfFdYy= FdYy " + _
                 "                 and IfFdChr= FdChr and IfFdNo= FdNo and IfFdSr= FdSr and (Select PValue from Param Where PTyp='TC' and PMCd=IfTc)='" + ctInTcTyp + "' " + IIF(gs_Partition = ctCurrPrtn, " and IfPrtKey=FdPrtKey ", "") + ")) "
      '*** (Jen 2.14 Next Patch 1 23/05/07)
      
    Else
      '*** (Jen 2.12 Patch 3)
      ws_TmpTblCnd = GetCnd("SubQuery")
      ws_TmpTblCnd = IIF(ws_TmpTblCnd <> "", " where " + ws_TmpTblCnd, " ")
      
      '****** Sachin 3.02 - Id fields in Join
      ' Bef FgRm Speed Issue in Shrenuj
'''      ws_AddFgCnd = " and FdYy = (Select max(FdYy) from Fg F1 join Fgd Fd1 " + _
'''                    "   on F1.FgIdNo=Fd1.FdFgIdNo and F1.FgCoCd= Fd1.FdCoCd and F1.FgTc= Fd1.FdTc and " + _
'''                    "   F1.FgYy= Fd1.FdYy and F1.FgChr= Fd1.FdChr and F1.FgNo= Fd1.FdNo " + IIF(gs_Partition = ctCurrPrtn, " and F1.FgPrtKey=Fd1.FdPrtKey ", "") + _
'''                    "   Where Fd1.FdBIdNo=FdMax.FdBIdNo and Fd1.FdCoCd = FdMax.FdCoCd And Fd1.FdBYy = FdMax.FdBYy And Fd1.FdBChr = FdMax.FdBChr " + _
'''                    "   and Fd1.FdBNo= FdMax.FdBNo " + IIF(gs_Partition = ctCurrPrtn, " and Fd1.FdPrtKey=FdMax.FdPrtKey ", "") + _
'''                        IIF(ADC("UoDtFr") = "01/01/80", " ", " and F1.FgDt>= '" + CStr(ADC("UoDtFr")) + "' ") + _
'''                        IIF(ADC("UoDtTo") = "01/01/80", " ", " and F1.FgDt<= '" + CStr(ADC("UoDtTo")) + "' ") + _
'''                        IIF(ADC("UoDtFr") = "01/01/80", " ", " and Fd1.FdDt>= '" + CStr(ADC("UoDtFr")) + "' ") + _
'''                        IIF(ADC("UoDtTo") = "01/01/80", " ", " and Fd1.FdDt<= '" + CStr(ADC("UoDtTo")) + "' ") + _
'''                        ") "
'''
'''      '****** Sachin 3.02 - Id fields in Joins
'''      ws_AddFgCnd = ws_AddFgCnd + " and FdKey = (Select max(FdKey) from Fg F1 join Fgd Fd1 " + _
'''                    "   on F1.FgIdNo=Fd1.FdFgIdNo and F1.FgCoCd= Fd1.FdCoCd and F1.FgTc= Fd1.FdTc and " + _
'''                    "   F1.FgYy= Fd1.FdYy and F1.FgChr= Fd1.FdChr and F1.FgNo= Fd1.FdNo " + IIF(gs_Partition = ctCurrPrtn, " and F1.FgPrtKey=Fd1.FdPrtKey ", "") + _
'''                    "   Where Fd1.FdBIdNo=FdMax.FdBIdNo and  Fd1.FdCoCd = FdMax.FdCoCd And Fd1.FdBYy = FdMax.FdBYy And Fd1.FdBChr = FdMax.FdBChr " + _
'''                    "   and Fd1.FdBNo= FdMax.FdBNo " + IIF(gs_Partition = ctCurrPrtn, " and Fd1.FdPrtKey=FdMax.FdPrtKey ", "") + _
'''                        IIF(ADC("UoDtFr") = "01/01/80", " ", " and F1.FgDt>= '" + CStr(ADC("UoDtFr")) + "' ") + _
'''                        IIF(ADC("UoDtTo") = "01/01/80", " ", " and F1.FgDt<= '" + CStr(ADC("UoDtTo")) + "' ") + _
'''                        IIF(ADC("UoDtFr") = "01/01/80", " ", " and Fd1.FdDt>= '" + CStr(ADC("UoDtFr")) + "' ") + _
'''                        IIF(ADC("UoDtTo") = "01/01/80", " ", " and Fd1.FdDt<= '" + CStr(ADC("UoDtTo")) + "' ") + _
'''                        " " + ws_AddFgCnd + ") "
       
      ' ***** Manali FgRm Speed
      '578-34 STR(FdKey, 7) changed to STR(FdKey, 8)
      ws_AddFgCnd = " and (FdYy + STR(FdKey, 8)) = (Select max(FdYy + STR(FdKey, 8)) from Fgd Fd1 " + _
                    "   Where Fd1.FdBIdNo=FdMax.FdBIdNo and Fd1.FdCoCd = FdMax.FdCoCd And Fd1.FdBYy = FdMax.FdBYy And Fd1.FdBChr = FdMax.FdBChr " + _
                    "   and Fd1.FdBNo= FdMax.FdBNo " + IIF(gs_Partition = ctCurrPrtn, " and Fd1.FdPrtKey=FdMax.FdPrtKey ", "") + _
                        IIF(adc("UoDtFr") = "01/01/80", " ", " and Fd1.FdDt>= '" + CStr(adc("UoDtFr")) + "' ") + _
                        IIF(adc("UoDtTo") = "01/01/80", " ", " and Fd1.FdDt<= '" + CStr(adc("UoDtTo")) + "' ") + _
                        ") "
      
      'ws_FgTbl = " Join ( Select FdCoCd, FdBYy, FdBChr, FdBNo, FdTc, FdYy, FdChr, FdNo, FdSr, max(FdPlYy) as FdPlYy, " + _
                 "max(FdPlChr) as FdPlChr, max(FdPlNo) as FdPlNo, max(FdQty) as FdQty, max(FgDt) as FgDt, " + _
                 "max(FgToBLoc) as FgToBLoc, SUBSTRING (max(IsNull(FmdYy, '  ')+ Str(IsNull(FmdKey, 0) )+ " + _
                 "IsNull(FmdToFgSubLoc, FdSubLoc)), 13, len(max(IsNull(FmdYy, '  ')+ Str(IsNull(FmdKey, 0) )+ " + _
                 "IsNull(FmdToFgSubLoc, FdSubLoc)))- 12) as qFgSubLoc " + _
                 "From FG FgMax " + _
                 "join Fgd FdMax on FdCoCd= FgCoCd and FdTc= FgTc and FdYy= FgYy and FdChr= FgChr and FdNo= FgNo " + _
                 "left outer join Fmd on FmdCoCd= FdCoCd and FmdFdYy= FdYy and FmdFdKey= FdKey " + ws_TmpTblCnd + " " + _
                 ws_AddFgCnd + " group By FdCoCd, FdBYy, FdBChr, FdBNo, FdTc, FdYy, FdChr, FdNo, FdSr) FgTbl " + _
                 " on FdCoCd= BCoCd and FdBYy= BYy and FdBChr= BChr and FdBNo= BNo "
      
      'SUBSTRING (max(IsNull(FmdYy, '  ')+ Str(IsNull(FmdKey, 0) )+ " + _
                 "IsNull(FmdToFgSubLoc, FdSubLoc)), 13, len(max(IsNull(FmdYy, '  ')+ Str(IsNull(FmdKey, 0) )+ " + _
                 "IsNull(FmdToFgSubLoc, FdSubLoc)))- 12) as qFgSubLoc
      
      '****** (Jen 3.2.0) added FdBIdNo, FdIdNo, FdPrtKey in the select list
      '****** Sachin 3.02 - Id fields in Joins
      ws_FgTbl = " Join (Select FdCoCd, FdBYy, FdBChr, FdBNo, FdTc, FdYy, FdChr, FdNo, FdSr, max(FdPlYy) as FdPlYy, " + _
                 " max(FdPlChr) as FdPlChr, max(FdPlNo) as FdPlNo, max(FdQty) as FdQty, max(FgDt) as FgDt, " + _
                 " max(FgToBLoc) as FgToBLoc, SUBSTRING (max(IsNull(FmdYy, '00')+ Str(IsNull(FmdKey, 0), 10, 0)+ " + _
                 " IsNull(FmdToFgSubLoc, FdSubLoc)), 13, len(max(IsNull(FmdYy, '00')+ Str(IsNull(FmdKey, 0), 10, 0)+ " + _
                 " IsNull(FmdToFgSubLoc, FdSubLoc)))- 12) as qFgSubLoc, max(FdBIdNo) as FdBIdNo, max(FdIdNo) as FdIdNo, " + _
                 " max(FdPrtKey) as FdPrtKey " + _
                 " From FG FgMax " + _
                 " join Fgd FdMax on FdFgIdNo=FgIdNo and FdCoCd= FgCoCd and FdTc= FgTc and FdYy= FgYy and FdChr= FgChr and FdNo= FgNo " + IIF(gs_Partition = ctCurrPrtn, " and FdPrtKey=FgPrtKey ", "") + _
                 " left outer join (Select FmdPrtKey, FmdCoCd, FmdFdYy, FmdFdKey, FmdYy, FmdKey, FmdToFgSubLoc From Fmd " + _
                 "                 join Fm on FmCoCd= FmdCoCd and FmTc= FmdTc and FmYy= FmdYy and FmChr= FmdChr and FmNo= FmdNo " + IIF(gs_Partition = ctCurrPrtn, " and FmPrtKey=FmdPrtKey ", "") + _
                 "                 where " + IIF(gs_Partition = ctCurrPrtn, " FmPrtKey='" + ctCurrPrtn + "' and ", "") + IIF(adc("UoDtFr") = "01/01/80", " ", " FmDt>= '" + CStr(adc("UoDtFr")) + "' ") + _
                                             IIF(adc("UoDtTo") = "01/01/80", " ", IIF(adc("UoDtFr") = "01/01/80", " ", " and ") + " FmDt<= '" + CStr(adc("UoDtTo")) + "' ") + " " + _
                 ") FmTbl on FmdCoCd= FdCoCd and FmdFdYy= FdYy and FmdFdKey= FdKey " + IIF(gs_Partition = ctCurrPrtn, " and FmdPrtKey=FdPrtKey ", "") + ws_TmpTblCnd + " " + _
                 ws_AddFgCnd + " group By FdCoCd, FdBYy, FdBChr, FdBNo, FdTc, FdYy, FdChr, FdNo, FdSr) FgTbl " + _
                 " on " + IIF(gs_Partition = ctCurrPrtn, " FdPrtKey=BPrtKey and ", "") + " FdBIdNo=BIdNo and FdCoCd= BCoCd and FdBYy= BYy and FdBChr= BChr and FdBNo= BNo "
      
      '''ws_FgTbl = ws_FgTbl + _
      '''           " left outer join Fmd Fm on FmdCoCd= FdCoCd and FmdYy= qFmdYy and FmdKey= qFmdKey "
                 
      '*** (Jen 2.12 Patch 3)
      
    End If
    '*** Jen (2.12 Patch)
    
    wFgSel = "(Case when LFr.LocTyp= 'F' and FrFrRmDc= 'D' then 1 " + _
              "when LFr.LocTyp= 'F' and FrFrRmDc= 'C' then -1 " + _
              "when LTo.LocTyp= 'F' and FrToRmDc= 'D' then 1 " + _
              "when LTo.LocTyp= 'F' and FrToRmDc= 'C' then -1 " + _
              "Else 0 end) as qMul, "
    wFgJoin = " Join Loc LFr on LFr.LocCoCd= FrCoCd and LFr.LocCd= FrFrRmLoc " + _
              " Join Loc LTo on LTo.LocCoCd= FrCoCd and LTo.LocCd= FrToRmLoc "
    Rep.TxtHead.SetText "Fg Bag Raw Material Report"
'    Rep.TxthLocFr.SetText "Loc :": Rep.TxthLocTo.SetText " ": Rep.TxthLocSel.SetText " "
'    Rep.TxtLocFr.SetText "PFG"
    Rep.TxthLocFr.SetText "Loc :": Rep.TxthLocTo.SetText "Bags In": 'Rep.TxthLocSel.SetText " " 22 Loc sel removed from header
    Rep.TxtLocFr.SetText "PFG"
    Rep.TxtLocTo.SetText IIF(adc("UoYn4") = "Y", " FG ", IIF(adc("UoYn4") = "N", " INV ", "FG & INV"))

    If adc("UoDtFr") = "01/01/80" Then
      Rep.TxtFgDtFr.Suppress = True
      Rep.LblFgDtFr.Suppress = True
    Else
        Rep.TxtFgDtFr.SetText adc("UoDtFr")
    End If
    If adc("UoDtTo") = "01/01/80" Then
      Rep.TxtFgDtTo.Suppress = True
      Rep.LblFgDtTo.Suppress = True
    Else
      Rep.TxtFgDtTo.SetText adc("UoDtTo")
    End If
  
  End If
  '22 Customer From & To added to header
  Rep.TxtCustFr.SetText adc("UoCmCdFr"): Rep.TxtCustTo.SetText adc("UoCmCdTo")
  
  '*** (Bef 2.12 Patch)
  ''***************** urmi ******************
  'wBagClsCnd = IIF(adc("UoYN1") = "Y", " and BCls= 'N' ", IIF(adc("UoYN1") = "N", " and BCls= 'Y' ", ""))
  ''***************** urmi ******************
  '*** (Bef 2.12 Patch)
  
  '****** Sachin 2.13.0 - 30-05-2006 ******
  'If UCase(adc.MenuCd) = UCase("RepWipFgBRm") Then
    'wRmQwStr = ", RmQW as RmQW, RmQWCst as RmQWCst "
    ' ***** Manali 3.03 - 24/06/08 - Pd Mod - 'L' added
  wRmQWStr = ", IsNull((Case when (RmCtg in ('G', 'P', 'S', 'L', 'M') And RmQw='W') then (Select RmPr.RmQw From RmMst RmPr " + _
              "   Where " + IIF(gs_Partition = ctCurrPrtn, " RmPr.RmPrtKey='" + ctCurrPrtn + "' and ", "") + " RmPr.RmCd=(Select PValue from param " + _
              "       Where PTyp= 'RMCTG' and PMCd= Rm.RmCtg)) Else RmQw End),'')  as RmQw," + _
              " IsNull((Case when (RmCtg in ('G', 'P', 'S', 'L', 'M') And RmQw='W') then (Select RmPr.RmQwCst From RmMst RmPr " + _
              "   Where " + IIF(gs_Partition = ctCurrPrtn, " RmPr.RmPrtKey='" + ctCurrPrtn + "' and ", "") + " RmPr.RmCd=(Select PValue from param " + _
              "       Where PTyp= 'RMCTG' and PMCd= Rm.RmCtg)) Else RmQwCst End),'') as RmQwCst "
  'Else
  '  wRmQwStr = ", '' as RmQw, '' as RmQWCst "
  'End If
  '****** Sachin 2.13.0 - 30-05-2006 ******
  
 '*** Jay 2.13(CT) ***
 ' **** Zubin 211 **** '
 'Uma *** blank <> 'D'
 'IG.88-BChr Fr,To & Sel added in Param Join to get BChr list
    ws_BagJoin = " Join Param BTyp On BTyp.PTyp= 'BCHR' and BTyp.PMCd= BChr and BTyp.PSCd= '' " + _
                 IIF(adc("UoBagTyp") <> "", " and BTyp.PValue= '" + adc("UoBagTyp") + "' ", " and BTyp.PValue <> 'D' ") + _
                 IIF(adc("UoBChr2Fr") <> "", " and BTyp.PMCd>= '" + adc("UoBChr2Fr") + "' ", " ") + _
                 IIF(adc("UoBChr2To") <> "", " and BTyp.PMCd<= '" + adc("UoBChr2To") + "' ", " ") + _
                 IIF(adc("UoBChr2Sel") <> "", " and BTyp.PMCd IN (" + adc("UoBChr2Sel") + ") ", " ")
'  If UCase(adc("UoBagTyp")) = "N" Then
'    ws_FluteBagJoin = " Join Param On PTyp= 'BCHR' and PMCd= BChr and PSCd= '' and PValue = 'N' " + _
'            IIF(adc("UOBCHRFR") <> "", " and PMCd= '" + adc("UOBCHRFR") + "' ", " ")
'  ElseIf UCase(adc("UoBagTyp")) = "Y" Then
'    ws_FluteBagJoin = " Join Param On PTyp= 'BCHR' and PMCd= BChr and PSCd= '' and PValue = 'Y' " + _
'            IIF(adc("UOBCHRFR") <> "", " and PMCd= '" + adc("UOBCHRFR") + "' ", " ")
'  Else
'    ws_FluteBagJoin = " Join Param On PTyp= 'BCHR' and PMCd= BChr and PSCd= '' " + _
'            IIF(adc("UOBCHRFR") <> "", " and PMCd= '" + adc("UOBCHRFR") + "' ", " ")
'  End If
  ' **** Zubin 211 **** '
  '*** Jay 2.13(CT) ***
  
  ' ***** Manali 3.5.0 - 08/11/08 -  Bag Delay Days Cond Added
  wBagDelayCnd = IIF(adc("UoBDlDays") = 0, "", " and datediff(dd, BRecvDt, '" + CStr(moCn.SrvrDate) + "') >= " + CStr(adc("UoBDlDays")))
  wBagDelayCnd = wBagDelayCnd + IIF(adc("UoBDlDaysTo") = 0, "", " and  datediff(dd, BRecvDt, '" + CStr(moCn.SrvrDate) + "') <= " + CStr(adc("UoBDlDaysTo")))
  ' ***** Manali 3.5.0 - 08/11/08 -  Bag Delay Days Cond Added

  ' ***** Manali 3.4.1 - 02/12/08 - ["01/01/80" converted to purge date ]
  'sv.9 gs_PurgeDt replaced with ws_PurgeDt
  If gs_Partition = ctCurrPrtn And CDate(ws_PurgeDt) <> "01/01/80" And _
    (UCase(adc.MenuCd) = UCase("RepWipBRmBkDt") And CDate(adc("UoAsOfDt")) = MWLib.EmptyDate) Then
      wAsOfDt = CDate(ws_PurgeDt)
  Else
      wAsOfDt = adc("UoAsOfDt")
  End If

  'wrepcnd = ADC.RepCond
  'wCnd = IIF(wrepcnd <> "", " and " + wrepcnd, "")
      
  wrepcnd = GetCnd("Query")
  wCnd = IIF(wrepcnd <> "", " and " + wrepcnd, "")
  
  
  
  If UCase(adc.MenuCd) = UCase("RepWipBRm") Then
    If OptLoc(0).Value = True Then
      If adc("UoIssLocFr") <> "" Then wCnd = wCnd + " and BLoc >= '" + adc("UoIssLocFr") + "'"
      If adc("UoIssLocTo") <> "" Then wCnd = wCnd + " and BLoc <= '" + adc("UoIssLocTo") + "'"
      If adc("UoIssLocSel") <> "" Then wCnd = wCnd + " and BLoc In (" + adc("UoIssLocSel") + ")"
    ElseIf OptLoc(1).Value = True Then
      If adc("UoIssLocFr") <> "" Then wCnd = wCnd + " and LocPrcs >= '" + adc("UoIssLocFr") + "'"
      If adc("UoIssLocTo") <> "" Then wCnd = wCnd + " and LocPrcs <= '" + adc("UoIssLocTo") + "'"
      If adc("UoIssLocSel") <> "" Then wCnd = wCnd + " and LocPrcs In (" + adc("UoIssLocSel") + ")"
    End If
    '****** Sachin 2.13.0 - Left Outer Join added to Param for StnCls [07-07-2006] ******
    '****** Sachin 2.14.0 - Join added to Loc for Parent Code [01-12-2006] ******
    '****** Sachin 3.02.0 - Id fields in Joins
    
    '****** Sachin 3.02.0 - Additional Conditions For New Fields
    wAddlCnd = ""
    wAddlCnd = wAddlCnd + IIF(wAddlCnd <> "" And Trim(adc("UORMCTGFR")) <> "", " And ", "") + IIF(Trim(adc("UORMCTGFR")) <> "", " TdRmCtg >= '" + Trim(adc("UORMCTGFR")) + "'", "")
    wAddlCnd = wAddlCnd + IIF(wAddlCnd <> "" And Trim(adc("UORMCTGTO")) <> "", " And ", "") + IIF(Trim(adc("UORMCTGTO")) <> "", " TdRmCtg <= '" + Trim(adc("UORMCTGTO")) + "'", "")
    wAddlCnd = wAddlCnd + IIF(wAddlCnd <> "" And Trim(adc("UORMCTGSEL")) <> "", " And ", "") + IIF(Trim(adc("UORMCTGSEL")) <> "", " TdRmCtg In (" + Trim(adc("UORMCTGSEL")) + ")", "")
    
    wAddlCnd = wAddlCnd + IIF(wAddlCnd <> "" And Trim(adc("UORMSCTGFR")) <> "", " And ", "") + IIF(Trim(adc("UORMSCTGFR")) <> "", " TdRmSCtg >= '" + Trim(adc("UORMSCTGFR")) + "'", "")
    wAddlCnd = wAddlCnd + IIF(wAddlCnd <> "" And Trim(adc("UORMSCTGTO")) <> "", " And ", "") + IIF(Trim(adc("UORMSCTGTO")) <> "", " TdRmSCtg <= '" + Trim(adc("UORMSCTGTO")) + "'", "")
    wAddlCnd = wAddlCnd + IIF(wAddlCnd <> "" And Trim(adc("UORMSCTGSEL")) <> "", " And ", "") + IIF(Trim(adc("UORMSCTGSEL")) <> "", " TdRmSCtg In (" + Trim(adc("UORMSCTGSEL")) + ")", "")
    '****** Sachin 3.02.0 - Additional Conditions For New Fields
    '*** Jay 3.2.0 [PicNm]
    ' ***** Manali 3.03 - 24/06/08 - Pd Mod - 'L' added
    ' ***** Manali 3.03 - 07/07/08 - RmSzDesc added, join with RmIdSz added
    ' ***** Manali 3.5.0 - 08/11/08 -  Bag Delay Days Cond Added
    
    wSqlStrg1 = "Select " + grpflds0 + ", " + _
            "BCoCd, BYy, BChr, BNo, DmTcTyp, BOdDmCd, BOdDmSz, /* LocPrntCd, */" + IIF(OptLoc(0).Value = True, "BLoc", "LocPrcs") + " as BLoc, RmCtg, IsNull(SCls.PSCd, '') As StnCls, " + _
            "RmSCtg, TdRmCd, TdLotNo, TdRmSz, TdRmSz2, TdRmSz3, " + wRmSzDesc + " as qRmSzDesc, IsNull(RrRngSz,'') as RrRngSz, TdRmStkRt, " + _
            "(Select PValue3 from Param where PTyp= 'RMSCTG' and PMCd= RmCtg and PSCd= RmSCtg) as qPtrYN, " + _
            "TdRmQty, TdRmWt, " + _
            "(Case when TdFrRmLoc= 'W' and TdFrRmDc= 'D' then 1 " + _
            "when TdFrRmLoc= 'W' and TdFrRmDc= 'C' then -1 " + _
            "when TdToRmLoc= 'W' and TdToRmDc= 'D' then 1 " + _
            "when TdToRmLoc= 'W' and TdToRmDc= 'C' then -1 " + _
            "Else 0 end) as qMul, " + _
            "TdRmWt / (Case when RmCtg in ('D', 'C') then 5 else 1 end) as qGmWt, " + _
            "TdRmWt * (Case when RmCtg in ('G', 'P', 'S', 'L', 'M') then " + wPureWtOrZWt + " when RmCtg= 'X' then " + wXPureWtOrZwt + " else 0 end)" + wPurityWt + " as qGldWt, " + _
            wQty + "," + wPtQty + ", TDt , TdTc, TdYy, TdChr, TdNo, TdSr, TdSrNo, " + wTdRate + wRmQWStr + ", OdKt, OdPicNm " + _
            "From Bag " + ws_BagJoin + " Join OrdDsg on " + IIF(gs_Partition = ctCurrPrtn, " OdPrtKey=BPrtKey and ", "") + " OdIdNo=BOdIdNo and OdCoCd= BCoCd and OdTc= BOdTc and OdYy= BOdYy and " + _
            "OdChr= BOdChr and OdNo= BOdNo and OdSr= BOdSr Join OrdMst On " + IIF(gs_Partition = ctCurrPrtn, " OmPrtKey=OdPrtKey and ", "") + " OmIdNo=OdOmIdNo and OmCoCd= OdCoCd and " + _
            "OmTc= OdTc and OmYy= OdYy and OmChr= OdChr and OmNo= OdNo " + _
            "Join Txnd on TdBIdNo=BIdNo and TdCoCd= BCoCd and TdBYy=BYy and TdBChr=BChr and TdBNo=BNo and TdSrNo<> 0 " + IIF(gs_Partition = ctCurrPrtn, " and TdPrtKey=BPrtKey ", "") + _
            "Join Txn on TIdNo=TdTIdNo and TCoCd= TdCoCd and TTc= TdTc and TYy= TdYy and TChr= TdChr and TNo= TdNo " + IIF(gs_Partition = ctCurrPrtn, " and TPrtKey=TdPrtKey ", "") + _
            "Join Loc Lc on LocCoCd= BCoCd and LocCd= BLoc " + _
            "Join DsgMst on DmIdNo=BDmIdNo and DmTcTyp= BOdDmTcTyp and DmCd= BOdDmCd and DmSz= '' " + IIF(gs_Partition = ctCurrPrtn, " and DmPrtKey=BPrtKey ", "") + _
            "Join RmMst Rm on RmCd= TdRmCd " + IIF(gs_Partition = ctCurrPrtn, " and RmPrtKey=TdPrtKey ", "") + _
            " Left Outer Join LotMst on LotNo = TdLotNo " + _
            "Join Param RSCtg On RSCtg.PTyp='RMSCTG' And RSCtg.PMCd=Rm.RmCtg And RSCtg.PSCd=Rm.RmSCtg "
' "TdRmWt * (Case when RmCtg in ('G', 'P', 'S', 'L', 'M') then RmPurityWt when RmCtg= 'X' then RmBasePurityWt else 0 end)" + wPurityWt + " as qGldWt, "

     'JewelSmith.2 ',' + SCls.PDesc225 + ',' like '%,'+ RmSCtg +',%' changed as SCls.PSCd = RmStnCls
     wSqlStrg1 = wSqlStrg1 + " Left Outer Join Param SCls on SCls.PTyp= 'STNCLS' and SCls.PMCD=RmCtg and Scls.PSCd=RmStnCls " + _
            "Left Outer Join RmRt On RrTcTyp='RSZ' and RrCmCtg= 'C' And RrCmCd='" + ctSelfCmCd + "' And RrCtg=RmCtg And RrSCtg=RmSCtg and RrSTWGrp= '' and RrCd= '' and RrDmCtg= '' and " + _
            " RrLsCtg= '' and RrLabMCd= '' And RrFrLn<= TdRmSz and RrToLn>= TdRmSz " + _
            "Left Outer Join RmIdSz on RisRmCtg=RmCtg and RisRmId=RmId and RisSz1=TdRmSz Where" + _
            wSelCnd + wFgCnd + wCnd + IIF(gs_Partition = ctCurrPrtn, " and BPrtKey='" + ctCurrPrtn + "' ", "") + wBagClsCnd + IIF(wAddlCnd <> "", " And " + wAddlCnd, "") + wBagDelayCnd

    wSqlStrg = wSqlStrg1
 
  ElseIf UCase(adc.MenuCd) = UCase("RepWipFgBRm") Then
    '*** recs from FgRm ***
    '"FrRmWt * (Case when RmCtg in ('G', 'P') then RmPurityWt" + wPurityWt + " else 0 end) as qGldWt, "
    
    '"BLoc, "   '*** Bef 2.12 Patch
    
    '****** Sachin 3.02.0 - Additional Conditions For New Fields
    wAddlCnd = ""
    wAddlCnd = wAddlCnd + IIF(wAddlCnd <> "" And Trim(adc("UORMCTGFR")) <> "", " And ", "") + IIF(Trim(adc("UORMCTGFR")) <> "", " FrRmCtg >= '" + Trim(adc("UORMCTGFR")) + "'", "")
    wAddlCnd = wAddlCnd + IIF(wAddlCnd <> "" And Trim(adc("UORMCTGTO")) <> "", " And ", "") + IIF(Trim(adc("UORMCTGTO")) <> "", " FrRmCtg <= '" + Trim(adc("UORMCTGTO")) + "'", "")
    wAddlCnd = wAddlCnd + IIF(wAddlCnd <> "" And Trim(adc("UORMCTGSEL")) <> "", " And ", "") + IIF(Trim(adc("UORMCTGSEL")) <> "", " FrRmCtg In (" + Trim(adc("UORMCTGSEL")) + ")", "")
    
    wAddlCnd = wAddlCnd + IIF(wAddlCnd <> "" And Trim(adc("UORMSCTGFR")) <> "", " And ", "") + IIF(Trim(adc("UORMSCTGFR")) <> "", " FrRmSCtg >= '" + Trim(adc("UORMSCTGFR")) + "'", "")
    wAddlCnd = wAddlCnd + IIF(wAddlCnd <> "" And Trim(adc("UORMSCTGTO")) <> "", " And ", "") + IIF(Trim(adc("UORMSCTGTO")) <> "", " FrRmSCtg <= '" + Trim(adc("UORMSCTGTO")) + "'", "")
    wAddlCnd = wAddlCnd + IIF(wAddlCnd <> "" And Trim(adc("UORMSCTGSEL")) <> "", " And ", "") + IIF(Trim(adc("UORMSCTGSEL")) <> "", " FrRmSCtg In (" + Trim(adc("UORMSCTGSEL")) + ")", "")
    '****** Sachin 3.02.0 - Additional Conditions For New Fields
    
    '*** (Jen 2.12 Patch 3 changed the Fg& Fgd Table joins)
    '****** Sachin 2.13.0 - Left Outer Join added to Param for StnCls [07-07-2006] ******
    '****** Sachin 2.14.0 - Join added to Loc for Parent Code [01-12-2006] ******
    '****** Sachin 3.02.0 - Id fields in Joins
    '*** Jay 3.2.0 [PicNm]
    ' ***** Manali 3.03 - 24/06/08 - Pd Mod - 'L' added
    ' ***** Manali 3.03 - 07/07/08 - RmSzDesc added, join with RmIdSz added
    ' ***** Manali 3.5.0 - 08/11/08 -  Bag Delay Days Cond Added
    '***********Bhavna added DmSrcDsgCd, though not used in DSR now..
    wSqlStrg2 = "set dateformat dmy Select " + grpflds1 + ", " + _
                "BCoCd, BYy, BChr, BNo, DmTcTyp, DmSrcDsgCd, BOdDmCd, BOdDmSz, /* '' As LocPrntCd, */ " + IIF(adc("UoDtTo") = "01/01/80", "BLoc, ", " FgToBLoc, ") + " RmCtg, IsNull(SCls.PSCd,'') As StnCls, RmSCtg, FrRmCd as TdRmCd, FrLotNo as TdLotNo, " + _
                "FrRmSz as TdRmSz,FrRmSz2 as TdRmSz2,FrRmSz3 as TdRmSz3, " + wRmSzDesc + " as qRmSzDesc, IsNull(RrRngSz,'') as RrRngSz, FrRmStkRt as TdRmStkRt, " + _
                "(Select PValue3 from Param where PTyp= 'RMSCTG' and PMCd= RmCtg and PSCd= RmSCtg) as qPtrYN, " + _
                "FrRmQty as TdRmQty, FrRmWt as TdRmWt, " + wFgSel + _
                "FrRmWt / (Case when RmCtg in ('D', 'C') then 5 else 1 end) as qGmWt, " + _
                "FrRmWt * (Case when RmCtg in ('G', 'P', 'S', 'L', 'M') then " + wPureWtOrZWt + " when RmCtg= 'X' then " + wXPureWtOrZwt + " else 0 end)" + wPurityWt + " as qGldWt, " + _
                wQty + "," + wPtQty + ", FgDt as TDt, FrTc as TdTc, FrYy as TdYy, FrChr as TdChr, FrNo as TdNo, " + _
                "FrSr as TdSr, FrSrNo as TdSrNo, " + wFrRate + wRmQWStr + ", OdKt, OdPicNm " + _
                " From Bag Join OrdDsg on OdIdNo=BOdIdNo and OdCoCd= BCoCd and OdTc= BOdTc and OdYy= BOdYy and " + _
                "OdChr= BOdChr and OdNo= BOdNo and OdSr= BOdSr " + IIF(gs_Partition = ctCurrPrtn, " and OdPrtKey=BPrtKey ", "") + " Join OrdMst On OmCoCd= OdCoCd and " + _
                "OmTc= OdTc and OmYy= OdYy and OmChr= OdChr and OmNo= OdNo " + IIF(gs_Partition = ctCurrPrtn, " and OmPrtKey=OdPrtKey ", "") + ws_FgTbl + _
                " Join FgRm on FrFdIdNo=FdIdNo and  FrCoCd= FdCoCd and FrTc= FdTc and FrYy= FdYy " + _
                "and FrChr= FdChr and FrNo= FdNo and FrSr= FdSr " + IIF(gs_Partition = ctCurrPrtn, " and FrPrtKey=FdPrtKey ", "") + _
                "Join Loc Lc on LocCoCd= BCoCd and LocCd= " + IIF(adc("UoDtTo") = "01/01/80", "BLoc ", " FgToBLoc ") + _
                "Join DsgMst on DmIdNo=BDmIdNo and DmTcTyp= BOdDmTcTyp and DmCd= BOdDmCd and DmSz= '' " + IIF(gs_Partition = ctCurrPrtn, " and DmPrtKey=BPrtKey ", "") + _
                "Join RmMst Rm on RmCd= FrRmCd " + IIF(gs_Partition = ctCurrPrtn, " and RmPrtKey=FrPrtKey ", "") + _
                " Left Outer Join LotMst on LotNo = FrLotNo " + _
                "Join Param RSCtg On RSCtg.PTyp='RMSCTG' And RSCtg.PMCd=Rm.RmCtg And RSCtg.PSCd=Rm.RmSCtg "

'"FrRmWt * (Case when RmCtg in ('G', 'P', 'S', 'L', 'M') then RmPurityWt when RmCtg= 'X' then RmBasePurityWt else 0 end)" + wPurityWt + " as qGldWt, "
'JewelSmith.2 ',' + SCls.PDesc225 + ',' like '%,'+ RmSCtg +',%' changed as SCls.PSCd = RmStnCls
       wSqlStrg2 = wSqlStrg2 + " Left Outer Join Param SCls on SCls.PTyp= 'STNCLS' and SCls.PMCD=RmCtg and Scls.PSCd=RmStnCls " + _
                "Left Outer Join RmRt On RrTcTyp= 'RSZ' and RrCmCtg= 'C' And RrCmCd= '" + ctSelfCmCd + "' And RrCtg= RmCtg And RrSCtg= RmSCtg and RrSTWGrp= '' and RrCd= '' and RrDmCtg= '' and " + _
                " RrLsCtg= '' and RrLabMCd= '' And RrFrLn<= FrRmSz and RrToLn>= FrRmSz " + wFgJoin + _
                "Left Outer Join RmIdSz on RisRmCtg=RmCtg and RisRmId=RmId and RisSz1=FrRmSz " + _
                "Where " + IIF(gs_Partition = ctCurrPrtn, " BPrtKey='" + ctCurrPrtn + "' And ", "") + wSelCnd + wFgCnd + wCnd + wBagClsCnd + IIF(wAddlCnd <> "", " And " + wAddlCnd, "") + _
                wBagDelayCnd + IIF(ws_FGINVChk <> "", " And " + ws_FGINVChk, "")
    '*** (Jen 2.12 Patch 3 changed the Fg& Fgd Table joins)

    wSqlStrg = wSqlStrg2

  '*** (Jen 2.14)
  ElseIf UCase(adc.MenuCd) = UCase("RepWipBRmBkDt") Then
    '*** This query is only to view Open Bags As Of A Particular Date
    'Loc-Prcs selection on option button
    If OptLoc(0).Value = True Then
      If adc("UoIssLocFr") <> "" Then wCnd = wCnd + " and IsNull(Lc.LocCd, '') >= '" + adc("UoIssLocFr") + "'"
      If adc("UoIssLocTo") <> "" Then wCnd = wCnd + " and IsNull(Lc.LocCd, '') <= '" + adc("UoIssLocTo") + "'"
      If adc("UoIssLocSel") <> "" Then wCnd = wCnd + " and IsNull(Lc.LocCd, '') In (" + adc("UoIssLocSel") + ")"
    ElseIf OptLoc(1).Value = True Then
      If adc("UoIssLocFr") <> "" Then wCnd = wCnd + " and Lc.LocPrcs >= '" + adc("UoIssLocFr") + "'"
      If adc("UoIssLocTo") <> "" Then wCnd = wCnd + " and Lc.LocPrcs <= '" + adc("UoIssLocTo") + "'"
      If adc("UoIssLocSel") <> "" Then wCnd = wCnd + " and Lc.LocPrcs In (" + adc("UoIssLocSel") + ")"
    End If
    
    '****** Sachin 3.02.0 - Additional Conditions For New Fields
    wAddlCnd = ""
    wAddlCnd = wAddlCnd + IIF(wAddlCnd <> "" And Trim(adc("UORMCTGFR")) <> "", " And ", "") + IIF(Trim(adc("UORMCTGFR")) <> "", " Tr.TdRmCtg >= '" + Trim(adc("UORMCTGFR")) + "'", "")
    wAddlCnd = wAddlCnd + IIF(wAddlCnd <> "" And Trim(adc("UORMCTGTO")) <> "", " And ", "") + IIF(Trim(adc("UORMCTGTO")) <> "", " Tr.TdRmCtg <= '" + Trim(adc("UORMCTGTO")) + "'", "")
    wAddlCnd = wAddlCnd + IIF(wAddlCnd <> "" And Trim(adc("UORMCTGSEL")) <> "", " And ", "") + IIF(Trim(adc("UORMCTGSEL")) <> "", " Tr.TdRmCtg In (" + Trim(adc("UORMCTGSEL")) + ")", "")
    
    wAddlCnd = wAddlCnd + IIF(wAddlCnd <> "" And Trim(adc("UORMSCTGFR")) <> "", " And ", "") + IIF(Trim(adc("UORMSCTGFR")) <> "", " Tr.TdRmSCtg >= '" + Trim(adc("UORMSCTGFR")) + "'", "")
    wAddlCnd = wAddlCnd + IIF(wAddlCnd <> "" And Trim(adc("UORMSCTGTO")) <> "", " And ", "") + IIF(Trim(adc("UORMSCTGTO")) <> "", " Tr.TdRmSCtg <= '" + Trim(adc("UORMSCTGTO")) + "'", "")
    wAddlCnd = wAddlCnd + IIF(wAddlCnd <> "" And Trim(adc("UORMSCTGSEL")) <> "", " And ", "") + IIF(Trim(adc("UORMSCTGSEL")) <> "", " Tr.TdRmSCtg In (" + Trim(adc("UORMSCTGSEL")) + ")", "")
    '****** Sachin 3.02.0 - Additional Conditions For New Fields
  
    '****** Sachin 3.02.0 - Id fields in Joins
    '*** Jay 3.2.0 [PicNm]
    ' ***** Manali 3.03 - 24/06/08 - Pd Mod - 'L' added
    ' ***** Manali 3.03 - 07/07/08 - RmSzDesc added, join with RmIdSz added
    ' ***** Manali 3.4.1 - 02/12/08 - [UoAsOfDt replaced by wAsOfDt]
    wSqlStrg3 = "Select " + grpflds2 + ", " + _
                "   BCoCd, BYy, BChr, BNo, DmTcTyp, BOdDmCd, BOdDmSz, /* IsNull(Lc.LocPrntCd, '') as LocPrntCd, */ " + _
                IIF(OptLoc(0).Value = True, "IsNull(Lc.LocCd, '')", "IsNull(Lc.LocPrcs,'')") + " as BLoc, /* IsNull(qBLoc, (case when qTdToBLoc= '' then (case when qTdFrBLoc= '' then BOpnLoc else qTdFrBLoc end) else qTdToBLoc end)) as BLoc, */ " + _
                "   RmCtg, IsNull(SCls.PSCd, '') As StnCls, RmSCtg, TR.TdRmCd as TdRmCd, TR.TdLotNo as TdLotNo, " + _
                "   TR.TdRmSz as TdRmSz,TR.TdRmSz2 as TdRmSz2,TR.TdRmSz3 as TdRmSz3, " + wRmSzDesc + " as qRmSzDesc, IsNull(RrRngSz, '') as RrRngSz, TR.TdRmStkRt as TdRmStkRt, " + _
                "   (Select PValue3 from Param where PTyp= 'RMSCTG' and PMCd= RmCtg and PSCd= RmSCtg) as qPtrYN, " + _
                "   TR.TdRmQty as TdRmQty, TR.TdRmWt as TdRmWt, " + _
                "   (Case when TR.TdFrRmLoc= 'W' and TR.TdFrRmDc= 'D' then 1 " + _
                "      when TR.TdFrRmLoc= 'W' and TR.TdFrRmDc= 'C' then -1 " + _
                "      when TR.TdToRmLoc= 'W' and TR.TdToRmDc= 'D' then 1 " + _
                "      when TR.TdToRmLoc= 'W' and TR.TdToRmDc= 'C' then -1 Else 0 end) as qMul, " + _
                "   TR.TdRmWt / (Case when RmCtg in ('D', 'C') then 5 else 1 end) as qGmWt, " + _
                "   TR.TdRmWt * (Case when RmCtg in ('G', 'P', 'S', 'L', 'M') then " + wPureWtOrZWt + " when RmCtg= 'X' then " + wXPureWtOrZwt + " else 0 end)" + wPurityWt + " as qGldWt, " + _
                    wQty + ", " + wPtQty + ", /* qTDt */ T.TDt as TDt, TR.TdTc as TdTc, TR.TdYy as TdYy, TR.TdChr as TdChr, " + _
                "   TR.TdNo as TdNo, TR.TdSr as TdSr, TR.TdSrNo as TdSrNo, " + wBkDtRate + wRmQWStr + ", OdKt, OdPicNm " + _
                "from (Select BPrtKey, BIdNo, BCoCd, BYy, BChr, BNo, BOdDmTcTyp, BDmIdNo, BOdDmCd, BOdDmSz, BOdIdNo, BOdTc, BOdYy, BOdChr, BOdNo, BOdSr, BOpnDt, " + _
                "     BClsDt, BOpnLoc , BOpnQty, BOpnPtQty From Bag Where BOpnDt<= '" + CStr(wAsOfDt) + "' /* '" + CStr(adc("UoAsOfDt")) + "' */ and " + _
                "     (case when BClsDt= '01/01/80' then '31/12/28' else BClsDt end)> '" + CStr(wAsOfDt) + "' /* '" + CStr(adc("UoAsOfDt")) + "' */ ) B " + _
                ws_BagJoin + " " + _
                "Join OrdDsg on OdIdNo=BOdIdNo and OdCoCd= BCoCd and OdTc= BOdTc and OdYy= BOdYy and OdChr= BOdChr and OdNo= BOdNo and OdSr= BOdSr " + IIF(gs_Partition = ctCurrPrtn, " and OdPrtKey=BPrtKey ", "") + _
                "Join OrdMst On OmIdNo=OdOmIdNo and OmCoCd= OdCoCd and OmTc= OdTc and OmYy= OdYy and OmChr= OdChr and OmNo= OdNo " + IIF(gs_Partition = ctCurrPrtn, " and OmPrtKey=OdPrtKey ", "") + _
                "Join DsgMst on DmIdNo=BDmIdNo and DmTcTyp= BOdDmTcTyp and DmCd= BOdDmCd and DmSz= '' " + IIF(gs_Partition = ctCurrPrtn, " and DmPrtKey=BPrtKey ", "")
      
                '" Join Param BTyp On BTyp.PTyp= 'BCHR' and BTyp.PMCd= BChr and BTyp.PSCd= '' and BTyp.PValue = 'N' "
          '"   TR.TdRmWt * (Case when RmCtg in ('G', 'P', 'S', 'L', 'M') then RmPurityWt when RmCtg= 'X' then RmBasePurityWt else 0 end) " + wPurityWt + " as qGldWt, "
    '*** *** ***
    '(Bef 09/01/07) this works too
    'wSqlStrg3 = wSqlStrg3 + _
                " join (select TdCoCd, TdBYy, TdBChr, TdBNo, max(TdYyKey) as qTdYyKey, " + _
                "       SubString (max(TdYyKey+ PValue), 10, 5) as qTTcTyp, " + _
                "       SubString (max(TdYyKey+ TdFrBLoc), 10, 5) as qTdFrBLoc, " + _
                "       SubString (max(TdYyKey+ TdToBLoc), 10, 5) as qTdToBLoc, " + _
                "       SubString (max(TdYyKey+ TdBDC), 10, 1) as qTdBDC, " + _
                "       SubString (max(TdYyKey+ TdWrk), 10, 8) as qTdWrk, " + _
                "       /* SubString (max(TdYyKey+ convert (varchar (8), TDt, 3)), 10, 8) as qTDt, */ " + _
                "       convert (decimal (16, 3), SubString (max(TdYyKey+ convert (varchar (16), TdBGrWt)), 10, 16)) as qTdBGrWt, " + _
                "       sum (TdBQty* case when TdBDC= 'C' then -1 when TdBDC= 'D' and TdTc<> 'BM' then 1 else 0 end) qTdBQty, " + _
                "       sum(TdPtQty* case when TdBDC= 'C' then -1 when TdBDC= 'D' and TdTc<> 'BM' then 1 else 0 end) qTdBPtQty, " + _
                "       sum(TdRjQty* case when TdBDC= 'C' then -1 when TdBDC= 'D' then 1 else 0 end) qTdRjQty, " + _
                "       sum(TdRjPtQty* case when TdBDC= 'C' then -1 when TdBDC= 'D' then 1 else 0 end) qTdRjPtQty " + _
                "   From Txn " + _
                "   join (select PMCd, PValue FROM Param where PTyp= 'TC') Tc on PMCd= TTc " + _
                "   Join Txnd on TdCoCd= TCoCd and TdTc= TTc and TdYy= TYy and TdChr= TChr and TdNo= TNo " + _
                "   where TDt<= '" + CStr(adc("UoAsOfDt")) + "' and TdSrNo= 0 " + _
                "   group by TdCoCd, TdBYy, TdBChr, TdBNo) TB " + _
                "   on BYy= TB.TdBYy and BChr= TB.TdBChr and BNo= TB.TdBNo and BCoCd= TB.TdCoCd " + _
                "join Txnd TR on TR.TdCoCd= TB.TdCoCd and TR.TdBYy= TB.TdBYy and TR.TdBChr= TB.TdBChr and TR.TdBNo= TB.TdBNo and TR.TdSrNo<> 0 and TR.TdYyKey<= TB.qTdYyKey " + _
                "Join RmMst Rm on RmCd= TR.TdRmCd " + _
                "Join Txn T on TCoCd= TR.TdCoCd and TTc= TR.TdTc and TYy= TR.TdYy and TChr= TR.TdChr and TNo= TR.TdNo "
    '(Bef 09/01/07) this works too
    
    '(09/01/07)
    '****** Sachin 3.02.0 - Id fields in Joins
    '578-34 substring starting position changed from 10 to 11
     wSqlStrg3 = wSqlStrg3 + _
                " join (select TdPrtKey, TdBIdNo, TdCoCd, TdBYy, TdBChr, TdBNo, max(TdYyKey) as qTdYyKey, " + _
                "       SubString (max(TdYyKey+ PValue), 11, 5) as qTTcTyp, " + _
                "       SubString (max(TdYyKey+ TdFrBLoc), 11, 8) as qTdFrBLoc, " + _
                "       SubString (max(TdYyKey+ TdToBLoc), 11, 8) as qTdToBLoc, " + _
                "       SubString (max(TdYyKey+ TdBDC), 11, 1) as qTdBDC, " + _
                "       SubString (max(TdYyKey+ TdWrk), 11, 8) as qTdWrk, " + _
                "       /* SubString (max(TdYyKey+ convert (varchar (8), TDt, 3)), 10, 8) as qTDt, */ " + _
                "       convert (decimal (16, 3), SubString (max(TdYyKey+ convert (varchar (16), TdBGrWt)), 11, 16)) as qTdBGrWt, " + _
              "       sum (TdBQty* case when TTrfQty=0 and TdBDC= 'C' then -1 when  TTrfQty=0 and TdBDC= 'D' and TdTc<> 'BM' then 1 else 0 end) qTdBQty, " + _
                "       sum(TdPtQty* case when  TTrfQty=0 and TdBDC= 'C' then -1 when  TTrfQty=0 and TdBDC= 'D' and TdTc<> 'BM' then 1 else 0 end) qTdBPtQty, " + _
                "       sum(TdRjQty* case when  TTrfQty=0 and TdBDC= 'C' then -1 when  TTrfQty=0 and TdBDC= 'D' then 1 else 0 end) qTdRjQty, " + _
                "       sum(TdRjPtQty* case when  TTrfQty=0 and TdBDC= 'C' then -1 when  TTrfQty=0 and TdBDC= 'D' then 1 else 0 end) qTdRjPtQty " + _
                "   From Txnd " + _
                "   join Bag on BIdNo=TdBIdNo and BCoCd= TdCoCd and BYy= TdBYy and BChr= TdBChr and BNo= TdBNo " + IIF(gs_Partition = ctCurrPrtn, " and BPrtKey=TdPrtKey ", "") + " and BOpnDt<= '" + CStr(wAsOfDt) + "' /* '" + CStr(adc("UoAsOfDt")) + "'*/ and " + _
                "       (case when BClsDt= '01/01/80' then '31/12/28' else BClsDt end)> '" + CStr(wAsOfDt) + "' /* '" + CStr(adc("UoAsOfDt")) + "' */ " + _
                "   Left Outer Join LotMst on LotNo = TdLotNo " + _
                "   Join Txn on TIdNo=TdTIdNo and TCoCd= TdCoCd and TTc= TdTc and TYy= TdYy and TChr= TdChr and TNo= TdNo " + IIF(gs_Partition = ctCurrPrtn, " and TPrtKey=TdPrtKey ", "") + _
                "   join (select PMCd, PValue FROM Param where PTyp= 'TC') Tc on PMCd= TTc " + _
                "   where " + IIF(gs_Partition = ctCurrPrtn, " TPrtKey='" + ctCurrPrtn + "' and ", "") + " TDt<= '" + CStr(wAsOfDt) + "' /*'" + CStr(adc("UoAsOfDt")) + "'*/ And TdDt<= '" + CStr(wAsOfDt) + "' /*'" + CStr(adc("UoAsOfDt")) + "'*/ and TdSrNo= 0 " + _
                "   group by TdPrtKey, TdBIdNo, TdCoCd, TdBYy, TdBChr, TdBNo) TB " + _
                "   on BIdNo=TB.TdBIdNo and BCoCd= TB.TdCoCd and BYy= TB.TdBYy and BChr= TB.TdBChr and BNo= TB.TdBNo " + IIF(gs_Partition = ctCurrPrtn, " and BPrtKey=TB.TdPrtKey ", "") + _
                "join Txnd TR on TR.TdBIdNo=TB.TdBIdNo and TR.TdCoCd= TB.TdCoCd and TR.TdBYy= TB.TdBYy and TR.TdBChr= TB.TdBChr and TR.TdBNo= TB.TdBNo " + IIF(gs_Partition = ctCurrPrtn, " and TR.TdPrtKey=TB.TdPrtKey ", "") + " and TR.TdSrNo<> 0 and TR.TdYyKey<= TB.qTdYyKey " + _
                "Join RmMst Rm on RmCd= TR.TdRmCd " + IIF(gs_Partition = ctCurrPrtn, " and RmPrtKey=TR.TdPrtKey ", "") + _
                " Left Outer Join LotMst on LotNo = TdLotNo Join Txn T on TIdNo=TR.TdTIdNo and TCoCd= TR.TdCoCd and TTc= TR.TdTc and TYy= TR.TdYy and TChr= TR.TdChr and TNo= TR.TdNo " + IIF(gs_Partition = ctCurrPrtn, " and TPrtKey=TR.TdPrtKey ", "")
    '(09/01/07)
    '*** ***
    
    '****** Sachin 3.02.0 - Id fields in Joins
    'JewelSmith.2 ',' + SCls.PDesc225 + ',' like '%,'+ RmSCtg +',%' changed as SCls.PSCd = RmStnCls
    'BS error Join Param RSCtg added
    wSqlStrg3 = wSqlStrg3 + _
                " Join Param RSCtg On RSCtg.PTyp='RMSCTG' And RSCtg.PMCd=Rm.RmCtg And RSCtg.PSCd=Rm.RmSCtg " + _
                " left outer join (select FdPrtKey, FdBIdNo, FdCoCd, FdBYy, FdBChr, FdBNo, max(FdBYyKey) as FdBYyKey, " + _
                "       max(FdYy+ str(FdKey, 10, 0)) as qFdYyKey, " + _
                "       max(case when PValue= 'FB' then FdYy+ Str(FdKey, 10, 0) else '' end) as qFbYyKey, " + _
                "       SubString(max(FdYy + Str(FdKey, 10, 0) + FgToBLoc), 13, 8) As qBLoc " + _
                "   from Fg F join Fgd on " + IIF(gs_Partition = ctCurrPrtn, " FdPrtKey=FgPrtKey and ", "") + " FdFgIdNo=FgIdNo and FdCoCd= FgCoCd and FdTc= FgTc and FdYy= FgYy and FdChr= FgChr and FdNo= FgNo " + _
                "   join Param on PTyp= 'TC' and PMCd= FgTc " + _
                "   Where " + IIF(gs_Partition = ctCurrPrtn, " FgPrtKey='" + ctCurrPrtn + "' and ", "") + " FgDt<= '" + CStr(wAsOfDt) + "' /*'" + CStr(adc("UoAsOfDt")) + "'*/ And FdDt<= '" + CStr(wAsOfDt) + "' /*'" + CStr(adc("UoAsOfDt")) + "'*/ " + _
                "   group By FdPrtKey, FdBIdNo, FdCoCd, FdBYy, FdBChr, FdBNo " + _
                "   ) TmpFg On FdBIdNo=BIdNo and FdCoCd= BCoCd and FdBYy= BYy and FdBChr= BChr and FdBNo= BNo and FdBYyKey= TB.qTdYyKey " + IIF(gs_Partition = ctCurrPrtn, " and FdPrtKey=BPrtKey ", "") + _
                "left outer join Loc Lc on Lc.LocCoCd= BCoCd and Lc.LocCd= IsNull(qBLoc, (case when qTdToBLoc= '' then (case when qTdFrBLoc= '' then BOpnLoc else qTdFrBLoc end) else qTdToBLoc end)) " + _
                "Left Outer Join Param SCls on SCls.PTyp= 'STNCLS' and SCls.PMCD=RmCtg and Scls.PSCd=RmStnCls " + _
                "left Outer Join RmRt On RrTcTyp= 'RSZ' and RrCmCtg= 'C' And RrCmCd= '" + ctSelfCmCd + "' " + _
                "   and RrCtg= RmCtg and RrSCtg= RmSCtg and RrSTWGrp= '' and RrCd= '' and RrDmCtg= '' " + _
                "   and RrLsCtg= '' and RrLabMCd= '' and RrFrLn<= TdRmSz and RrToLn>= TdRmSz " + _
                "Left Outer Join RmIdSz on RisRmCtg=RmCtg and RisRmId=RmId and RisSz1=TdRmSz " + _
                "where " + IIF(gs_Partition = ctCurrPrtn, " BPrtKey='" + ctCurrPrtn + "' and  ", "") + " IsNull(Lc.LocCd, '')<> 'PFG' " + _
                "   and IsNull (qFdYyKey, 'A')<> IsNull(qFbYyKey, 'B') " + wSelCnd + wCnd + IIF(wAddlCnd <> "", " And " + wAddlCnd, "") + _
                " Order By BYy, BChr, BNo "
      wSqlStrg = wSqlStrg3
  End If
  '*** (Jen 2.14)

  adc.RepRecSource = wSqlStrg + ctMaxDopOpt '****** Sachin 3.01 added ctMaxDopOpt
    
  If adc("UoYN") = "N" Then
    Rep.BoxHd.Suppress = True
    Rep.BoxDet.Suppress = True
    Rep.BoxFt.Suppress = True
  End If
  
  '***Geeta**Emr208 **Picture
  If adc("UoYN2") = "N" Or wDetPos = 0 Then
    Rep.PICSEC.Suppress = True
  Else
    Rep.PICSEC.Suppress = False
  End If
  
    
  '*** Jay 2.13(CT) ***
  'Uma *** blank <> 'D'
  'IG.88-BChr Fr,To & Sel added in Bag Chr range to show in report
  If adc("UoBagTyp") = "F" Then
    Rep.TxtBagTypAndChr.SetText "(Flute Bags" & IIF(adc("UoBChr2Fr") <> "", ", FrChr = " & adc("UoBChr2Fr"), "") & IIF(adc("UoBChr2To") <> "", "  ToChr:" & adc("UoBChr2To"), "") & IIF(adc("UoBChr2Sel") <> "", " SelChr:" & adc("UoBChr2Sel"), "") & ")"
  ElseIf adc("UoBagTyp") = "P" Then
    Rep.TxtBagTypAndChr.SetText "(Primary Bags" & IIF(adc("UoBChr2Fr") <> "", ", FrChr:" & adc("UoBChr2Fr"), "") & IIF(adc("UoBChr2To") <> "", "  ToChr:" & adc("UoBChr2To"), "") & IIF(adc("UoBChr2Sel") <> "", " SelChr:" & adc("UoBChr2Sel"), "") & ")"
  ElseIf adc("UoBagTyp") = "C" Then
    Rep.TxtBagTypAndChr.SetText "(Component Bags" & IIF(adc("UoBChr2Fr") <> "", ", FrChr = " & adc("UoBChr2Fr"), "") & IIF(adc("UoBChr2To") <> "", "  ToChr:" & adc("UoBChr2To"), "") & IIF(adc("UoBChr2Sel") <> "", " SelChr:" & adc("UoBChr2Sel"), "") & ")"
  ElseIf adc("UoBagTyp") = "D" Then
    Rep.TxtBagTypAndChr.SetText "(PD Bags" & IIF(adc("UoBChr2Fr") <> "", ", FrChr = " & adc("UoBChr2Fr"), "") & IIF(adc("UoBChr2To") <> "", "  ToChr:" & adc("UoBChr2To"), "") & IIF(adc("UoBChr2Sel") <> "", " SelChr:" & adc("UoBChr2Sel"), "") & ")"
  End If
  
  ' **** Manali 3.03 - 07/06/08 - RmIdSzDesc
  GRP_REP.SetFormula Rep, "wTdRmSz", "If {rdo.qRmSzDesc}<> '' Then {rdo.qRmSzDesc} " + _
                                     "Else (If {rdo.qPtrYN}= 'Y' then ToText({rdo.TdRmSz}, 4) Else ToText({rdo.TdRmSz}, 3))"
  ' **** Manali 3.03 - 07/06/08 - RmIdSzDesc
  
  '***(Jen 2.11)
  'If adc("UoYn2") <> "" Then Rep.TxtBagTypAndChr.SetText IIF(adc("UoYn2") = "Y", "(Flute Bags", "(Primary Bags") & IIF(adc("UoBChrFr") <> "", ",  Chr = " & adc("UoBChrFr"), "") & ")"
  '***(Jen 2.11)
  '*** Jay 2.13(CT) ***

  ' Zubin 212
  '****** Sachin 2.14.0 - [MultiCompany] ******
  If moCn.GetFldVal("select Max(HLotNoYn) from Head where HCd= '" + ctSelfCmCd + "' " + IIF(adc("UoCoCdFr") <> "", " and HCoCd>= '" + adc("UoCoCdFr") + "' ", " ") + IIF(adc("UoCoCdTo") <> "", " and HCoCd<= '" + adc("UoCoCdTo") + "' ", " ") + IIF(adc("UoCoCdSel") <> "", " and HCoCd In (" + adc("UoCoCdSel") + ") ", " ")) = "Y" Then
  'If Not moCn.GetFldVal("select max(HLotNoYn) from Head where HCd= '" + ctSelfCmCd + "' and HCoCd >= '" + adc("UoCoCdFr") + "' And HCoCd<='" + adc("UoCoCdTo") + "' And HCoCd In (" + adc("UoCoCdSel") + ")") = "Y" Then
    Rep.HLotNo.Suppress = True: Rep.FldLotNo.Suppress = True
  End If
  ' Zubin 212
  
 'KC.3- StkRt-when Show Stock Rate=YES then Stock Rate is visible in report
 'if it is NO then Stock Rate is invisible in report
 
  If adc("UoYN5") = "Y" Then
    Rep.hTdRmStkRt.SetText "Stock Rt"
    Rep.qCalcRt1.Suppress = False
  Else
    Rep.hTdRmStkRt.SetText ""
    Rep.qCalcRt1.Suppress = True
  End If
  '578.19 calling function to change Rmcode font size
  GRP_REP.SetFormula Rep, "wRmCdFontSz", GetRmcdFontSz("{rdo.TdRmCd}", 6)
    
  CRV_REP.DisplayGroupTree = False
  
  Call DispReport(adc) '****** Sachin 3.02 - Display Report ******
  
End Sub

' ****** Manali 3.5.0 - 21/11/08 - Check Boxes added
Private Sub ChkBoxArr_Click(Index As Integer)
With ChkBoxArr(Index)
  Select Case Index
  Case Is = 0     'Tag = "ShowBoxDet"
    If .Value = Checked Then adc("UoYN") = "Y" Else adc("UoYN") = "N"
  Case Is = 2     'Tag = "ShowPic"
    If .Value = Checked Then adc("UoYN2") = "Y" Else adc("UoYN2") = "N"
  Case Is = 3     'Tag = "ShowZWt"
    If .Value = Checked Then adc("UoYN3") = "Y" Else adc("UoYN3") = "N"
  'StkRt-To select this option(Show Stock rate) either Seo user can give their password or
  ' can enter SEOPWD password otherwise this option can't not select
  Case Is = 4
    If .Value = Checked Then
       If Trim(TxtPwd.text) <> moCn.GetFldVal("select IsNull ((select PValue from Param Join UsrAccess on UaUsrCd= PMCd " + _
            "where PTyp= 'USR' and PMCd= '" + gs_UsrCd + "'  and UaMnuCd= '" + adc.MenuCd + "' and " + _
            "UaSeoYn= 'Y' and PValue<> '' and PValue= '" + Trim(TxtPwd.text) + "' ), " + _
            "Isnull ((Select PValue1 from Param where " + _
            "Ptyp= 'SEOPWD' and PmCd= 'SEOPWD' ), ''))") Then _
            .Value = Unchecked
    End If
    If .Value = Checked Then adc("UoYN5") = "Y" Else adc("UoYN5") = "N"
    
  End Select
End With
End Sub
' ****** Manali 3.5.0 - 21/11/08 - Check Boxes added

' ****** Manali 3.5.0 - 21/11/08 - Check Boxes added
Private Sub ChkBoxArr_GotFocus(Index As Integer)
  DispMsg ChkBoxArr(Index).ToolTipText, etInfo
End Sub
' ****** Manali 3.5.0 - 21/11/08 - Check Boxes added



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

  'Set adc.FirNKeyCtl = adc("UoCoCdFr")
  
   '****** Sachin 2.14.0 - [23-11-2006] - [Multi-Company] ******
    ws_MultiCoMod = IIF(moCn.GetFldVal("Select hImpComLoc From Head Where HCd='" + ctSelfCmCd + "' And HCoCd='" + ctSelfCoCd + "'") = "Y", True, False)
    gb_CoCdFor = False
    
    
  '****** Sachin 2.14.0 - [23-11-2006] - [Multi-Company] ******
  'If adc("UoCoCdFr") = "" Then adc("UoCoCdFr") = gs_CoCd
  
  If UCase(adc.MenuCd) = UCase("RepWipBRm") Then
    
    '****** Sachin 2.14.0 - [23-11-2006] - [Multi-Company] ******
    'Set adc.FirNKeyCtl = adc("UoIssLocFr")
    
    If ws_MultiCoMod = True Then
      Set adc.FirNKeyCtl = adc("UoCoCdFr")
      LblPrntCd.Visible = True: adc("UoPrntCdFr").Visible = True: adc("UoPrntCdTo").Visible = True: adc("UoPrntCdSel").Visible = True
    Else
      Set adc.FirNKeyCtl = adc("UoIssLocFr")
      LblPrntCd.Visible = False: adc("UoPrntCdFr").Visible = False: adc("UoPrntCdTo").Visible = False: adc("UoPrntCdSel").Visible = False
    End If
    '****** Sachin 2.14.0 - [23-11-2006] - [Multi-Company] ******
  
    '*** (Jen 2.14)
    LblWrk.Visible = True: adc("UoWrkFr").Visible = True: adc("UoWrkTo").Visible = True: adc("UoWrkSel").Visible = True
    '*** (Jen 2.14)
    
    'LblBagPcs.Visible = True
    'adc("UoBagPcs").Visible = True: adc("UoBagPcs").Enabled = True
    ' ****** Manali 3.5.0 - 21/11/08 - Radio Buttons Added
    FraOptBagPcs.Enabled = True: FraOptBagPcs.Visible = True
    ' ****** Manali 3.5.0 - 21/11/08 - Radio Buttons Added

    
    lblFgSubLoc.Visible = False
    adc("UoFgSubLocFr").Visible = False: adc("UoFgSubLocFr").Enabled = False
    adc("UoFgSubLocTo").Visible = False: adc("UoFgSubLocTo").Enabled = False
    adc("UoFgSubLocSel").Visible = False: adc("UoFgSubLocSel").Enabled = False
    '*** Shilpa ***
    LblFgDt.Visible = False
    adc("UoDtFr").Visible = False: adc("UoDtFr").Enabled = False
    adc("UoDtTo").Visible = False: adc("UoDtTo").Enabled = False
    '*** Shilpa ***
    
    ' **** Zubin 211 **** '
    LblPL.Visible = False
    adc("UoFdPlYyFr").Visible = False: adc("UoFdPlYyTo").Visible = False
    adc("UoFdPlChrFr").Visible = False: adc("UoFdPlChrTo").Visible = False
    adc("UoFdPlNoFr").Visible = False: adc("UoFdPlNoTo").Visible = False
    LblPlFrSlsh1.Visible = False: LblPlFrSlsh2.Visible = False
    LblPlFrSlsh3.Visible = False: LblPlFrSlsh4.Visible = False
    ' **** Zubin 211 **** '
  
    '*** (Jen 2.14)
    LblAsOfDt.Visible = False: adc("UoAsOfDt").Visible = False
    'LblOpnCls.Visible = True: adc("UoYN1").Visible = True
    '*** (Jen 2.14)
    
    '4.1.0.0
    LblBagNo.Visible = True
    adc("UoBYyFr").Visible = True: adc("UoBYyTo").Visible = True: LblBgSlash1.Visible = True
    adc("UoBChrFr").Visible = True: adc("UoBChrTo").Visible = True: LblBgSlash2.Visible = True
    adc("UoBNoFr").Visible = True: adc("UoBNoTo").Visible = True: LblBgSlash3.Visible = True: LblBgSlash3.Visible = True
    
    '**********Bhavna src dsg cd for Fg bag Rm pcs
    LblSrcDsgCd.Visible = False
    adc("UoBOEFr").Visible = False: adc("UoBOEFr").Enabled = False
    adc("UoBOETo").Visible = False: adc("UoBOETo").Enabled = False
    '**********Bhavna src dsg cd for Fg bag pcs
    LblBagIn.Visible = False
    FraOptBagIn.Visible = False
    OptBagIn(0).Value = False: OptBagIn(1).Value = False: OptBagIn(2).Value = False

  ElseIf UCase(adc.MenuCd) = UCase("RepWipFgBRm") Then
    
    '****** Sachin 2.14.0 - [23-11-2006] - [Multi-Company] ******
    If ws_MultiCoMod = True Then
      Set adc.FirNKeyCtl = adc("UoCoCdFr")
    Else
      Set adc.FirNKeyCtl = adc("UoFgSubLocFr")
    End If
    
    LblPrntCd.Visible = False: adc("UoPrntCdFr").Visible = False: adc("UoPrntCdTo").Visible = False: adc("UoPrntCdSel").Visible = False
    '****** Sachin 2.14.0 - [23-11-2006] - [Multi-Company] ******
        
    '*** (Jen 2.14)
    LblWrk.Visible = False: adc("UoWrkFr").Visible = False: adc("UoWrkTo").Visible = False: adc("UoWrkSel").Visible = False
    '*** (Jen 2.14)
    
    LblProcess.Visible = False
    
    FraLoc.Visible = False: FraLoc.Enabled = False
    adc("UoIssLocFr").Visible = False: adc("UoIssLocFr").Enabled = False
    adc("UoIssLocTo").Visible = False: adc("UoIssLocTo").Enabled = False
    adc("UoIssLocSel").Visible = False: adc("UoIssLocSel").Enabled = False
    
    LblBagPcs.Visible = False
    adc("UoBagPcs").Visible = False: adc("UoBagPcs").Enabled = False
    ' ****** Manali 3.5.0 - 21/11/08 - Radio Buttons Added
    FraOptBagPcs.Enabled = False: FraOptBagPcs.Visible = False
    OptBagPcs(0).Value = False: OptBagPcs(1).Value = False
    ' ****** Manali 3.5.0 - 21/11/08 - Radio Buttons Added

    LblBagIn.Visible = True
    FraOptBagIn.Visible = True

    lblFgSubLoc.Visible = True
    adc("UoFgSubLocFr").Visible = True: adc("UoFgSubLocFr").Enabled = True
    adc("UoFgSubLocTo").Visible = True: adc("UoFgSubLocTo").Enabled = True
    adc("UoFgSubLocSel").Visible = True: adc("UoFgSubLocSel").Enabled = True
    LblFgDt.Visible = True
    adc("UoDtFr").Visible = True: adc("UoDtFr").Enabled = True
    adc("UoDtTo").Visible = True: adc("UoDtTo").Enabled = True
  
    adc("UoBagTyp").Visible = False
    LblBagTyp.Visible = False
    LblUoBChrFr.Visible = False
    adc("UoBChr2Fr").Visible = False: adc("UoBChr2To").Visible = False: adc("UoBChr2Sel").Visible = False
    ' **** Zubin 211 **** '
    
    '**********Bhavna src dsg cd for Fg bag Rm pcs
    LblSrcDsgCd.Visible = True
    adc("UoBOEFr").Visible = True: adc("UoBOEFr").Enabled = True
    adc("UoBOETo").Visible = True: adc("UoBOETo").Enabled = True
    '**********Bhavna src dsg cd for Fg bag pcs

    '*** (Jen 2.14)
    'LblAsOfDt.Visible = False: adc("UoAsOfDt").Visible = False
    'LblOpnCls.Visible = True:
    '*** (Jen 2.14)
    
    ' ****** Manali 3.5.0 - 06/11/08 - Radio Buttons Added
    'adc("UoYN1").Visible = True
    FraOptBag.Visible = True
    ' ****** Manali 3.5.0 - 06/11/08 - Radio Buttons Added
  '*** (Jen 2.14)
  
    '4.1.0.0
    LblBagNo.Visible = True
    adc("UoBYyFr").Visible = True: adc("UoBYyTo").Visible = True: LblBgSlash1.Visible = True
    adc("UoBChrFr").Visible = True: adc("UoBChrTo").Visible = True: LblBgSlash2.Visible = True
    adc("UoBNoFr").Visible = True: adc("UoBNoTo").Visible = True: LblBgSlash3.Visible = True: LblBgSlash3.Visible = True
    adc("UODTTO") = "01/01/99"
  ElseIf UCase(adc.MenuCd) = UCase("RepWipBRmBkDt") Then
  
    '****** Sachin 2.14.0 - [23-11-2006] - [Multi-Company] ******
    If ws_MultiCoMod = True Then
      Set adc.FirNKeyCtl = adc("UoCoCdFr")
      LblPrntCd.Visible = True: adc("UoPrntCdFr").Visible = True: adc("UoPrntCdTo").Visible = True: adc("UoPrntCdSel").Visible = True
    Else
      Set adc.FirNKeyCtl = adc("UoIssLocFr")
      LblPrntCd.Visible = False: adc("UoPrntCdFr").Visible = False: adc("UoPrntCdTo").Visible = False: adc("UoPrntCdSel").Visible = False
    End If
    '****** Sachin 2.14.0 - [23-11-2006] - [Multi-Company] ******
  
    LblWrk.Visible = True: adc("UoWrkFr").Visible = True: adc("UoWrkTo").Visible = True: adc("UoWrkSel").Visible = True
    'LblBagPcs.Visible = True
    'adc("UoBagPcs").Visible = True: adc("UoBagPcs").Enabled = True
    ' ****** Manali 3.5.0 - 21/11/08 - Radio Buttons Added
    FraOptBagPcs.Enabled = True: FraOptBagPcs.Visible = True
    ' ****** Manali 3.5.0 - 21/11/08 - Radio Buttons Added

    lblFgSubLoc.Visible = False
    adc("UoFgSubLocFr").Visible = False: adc("UoFgSubLocFr").Enabled = False
    adc("UoFgSubLocTo").Visible = False: adc("UoFgSubLocTo").Enabled = False
    adc("UoFgSubLocSel").Visible = False: adc("UoFgSubLocSel").Enabled = False
    LblFgDt.Visible = False
    adc("UoDtFr").Visible = False: adc("UoDtFr").Enabled = False
    adc("UoDtTo").Visible = False: adc("UoDtTo").Enabled = False
    
    LblPL.Visible = False
    adc("UoFdPlYyFr").Visible = False: adc("UoFdPlYyTo").Visible = False
    adc("UoFdPlChrFr").Visible = False: adc("UoFdPlChrTo").Visible = False
    adc("UoFdPlNoFr").Visible = False: adc("UoFdPlNoTo").Visible = False
    LblPlFrSlsh1.Visible = False: LblPlFrSlsh2.Visible = False
    LblPlFrSlsh3.Visible = False: LblPlFrSlsh4.Visible = False
    
    LblAsOfDt.Visible = True: adc("UoAsOfDt").Visible = True
    LblOpnCls.Visible = False: adc("UoYN1").Visible = False
  '*** (Jen 2.14)
  
    '**********Bhavna src dsg cd for Fg bag Rm pcs
    LblSrcDsgCd.Visible = False
    adc("UoBOEFr").Visible = False: adc("UoBOEFr").Enabled = False
    adc("UoBOETo").Visible = False: adc("UoBOETo").Enabled = False
    '**********Bhavna src dsg cd for Fg bag pcs

  
    ' ****** Manali 3.5.0 - 06/11/08 - Radio Buttons Added
    FraOptBag.Visible = False
    ' ****** Manali 3.5.0 - 06/11/08 - Radio Buttons Added
    LblBagIn.Visible = False
    FraOptBagIn.Visible = False
    OptBagIn(0).Value = False: OptBagIn(1).Value = False: OptBagIn(2).Value = False

  End If
   
  Call InitProp(Me)
  
  '****** Sachin 2.14.0 - [23-11-2006] - [Multi-Company] ******
  If ws_MultiCoMod = True Then
      adc("UOCOCDTO").Visible = True: adc("UOCOCDSEL").Visible = True
  Else
      adc("UOCOCDTO").Visible = False: adc("UOCOCDSEL").Visible = False
  End If
  '****** Sachin 2.14.0 - [23-11-2006] - [Multi-Company] ******
  
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
Private Sub atxt_FldChange(Index As Integer, ByVal IdName As String, Cancel As Boolean, ErrMsg As String, ByVal pv_NewValue As Variant)
  gs_PTyp = "WIPVAL"
  gs_Tbl = "Param"
  gb_CoCdFor = False '****** Sachin 2.14.0 - [25-11-2006] - [Multi-Company] ******
  
  Select Case UCase(IdName)
  ' urmi 2.06 Pwd for Val
  Case Is = UCase("UoPmCdFr"), UCase("UoPmCdTo")        '4.1.4
    '*** Jay 2.14 *** (SEO Change)
    
    If UCase(IdName) = UCase("UoPmCdFr") And pv_NewValue <> "STKVAL" Then adc("UoPMCdTo") = "" '4.1.4
    
    If pv_NewValue <> "" Then
          If Trim(TxtPwd.text) <> moCn.GetFldVal("select IsNull ((select PValue from Param Join UsrAccess on UaUsrCd= PMCd " + _
              "where PTyp= 'USR' and PMCd= '" + gs_UsrCd + "'  and UaMnuCd= '" + adc.MenuCd + "' and " + _
              "UaSeoYn= 'Y' and PValue<> '' and PValue= '" + Trim(TxtPwd.text) + "' ), " + _
              "(Select PValue1 from Param where " + _
              "Ptyp= 'SEOPWD' and PmCd= 'SEOPWD' ) )") Then _
              Cancel = True: ErrMsg = "Enter SEO Password to View Values ": Exit Sub
    End If
    
    '*** Jay 2.14 *** (SEO Change)
' Bef 2.14
'    If pv_NewValue <> "" And Trim(TxtPwd.text) <> moCn.GetFldVal("Select PValue1 from Param where " + _
'                                   "Ptyp = 'SEOPWD' and PmCd = 'SEOPWD'") Then
'      Cancel = True: DispMsg "Enter SEO Password to View Values", etError
'    End If
' Bef 2.14
    '*********Bhavna changed to remove error of Uodtto when new user given wip reports right
  Case Is = UCase("UoDtTo")
'       If (pv_NewValue = "01/01/80") Or Trim(pv_NewValue) = "" Then
'        Cancel = True: ErrMsg = "Please enter a Valid To Date": Exit Sub
'      End If
      If Trim(pv_NewValue) = "" Then
        Cancel = True: ErrMsg = "Please enter a Valid To Date": Exit Sub
      End If

  End Select
  Call SetProp(Me, IdName, Valid, Cancel, ErrMsg)
End Sub
Private Sub ATXT_FldWhen(Index As Integer, ByVal IdName As String, Cancel As Boolean, ErrMsg As String)
  gs_PTyp = "WIPVAL"
  gs_Tbl = "Param"
  gb_CoCdFor = False
  gs_LocTyp = "'P'"
  '*** (Jen 2.11)
  Select Case UCase(IdName)
    'IG.88- BChr Fr,TO & Sel for help added
    Case Is = UCase("UoBChr2Fr"), UCase("UoBChr2To")
      If adc.Mode = xNorm Then
        If adc("UoBagTyp") = "" Then Cancel = True: ErrMsg = "First Specify Bag Type": Exit Sub
        HlpList.PMCd "BCHR", "'" + adc("UoBagTyp") + "'": Exit Sub
      End If
    Case Is = UCase("UoBChr2Sel")
      If adc.Mode = xNorm Then
        If adc("UoBagTyp") = "" Then Cancel = True: ErrMsg = "First Specify Bag Type.": Exit Sub
        HlpList.MultiSelect = True
        HlpList.PMCd "BCHR", "'" + adc("UoBagTyp") + "'": Exit Sub
      End If
      
    '*** Jay 3.01 **** [LotNo Err]
    Case Is = UCase("UoLotNoFr"), UCase("UoLotNoTo"), UCase("UoLotNoSel")
      gs_RmLocTyp = "W": gs_Loc = "W"
    '*** Jay 3.01 **** [LotNo Err]
    
    Case Is = UCase("UoPMCdTo")     '4.1.4
      If adc.Mode = xNorm Then _
        If UCase(adc("UoPMCdFr")) <> UCase("STKVAL") Then Cancel = True: ErrMsg = "Can Enter Only if Value Type is 'STKVAL'": Exit Sub
    
    Case Is = UCase("UOISSLOCFR"), UCase("UOISSLOCTO"), UCase("UOISSLOCSEL")
      If adc("UOLOCTYPFR") = "R" Then gs_LocTyp = "'R'"
      If adc("UOLOCTYPFR") = "P" Then gs_LocTyp = "'P'"
   
  End Select
  '*** (Jen 2.11)
  
  SetProp Me, IdName, When
End Sub
Private Sub atxt_ReCalculate(Index As Integer, ByVal IdName As String, ByVal OldVal As Variant, ByVal RowNum As Integer, ByVal SourceName As String, ByVal SourceColName As String)
  ' **** Zubin 211 **** '
  Select Case UCase(IdName)
'    Case Is = UCase("UOBCHR")
'      'If ADC("UoBagTyp") = "N" Then
'      adc("UOBCHR") = ""
    Case Is = UCase("UoBChr2Fr")
      adc("UoBChr2Fr") = ""
    Case Is = UCase("UoBChr2To")
      adc("UoBChr2To") = ""
    Case Is = UCase("UoBChr2Sel")
      adc("UoBChr2Sel") = ""
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
Private Sub OptBagIn_Click(Index As Integer)
Select Case Index
Case Is = 0
  adc("UoYN4") = "Y"
Case Is = 1
  adc("UoYN4") = "N"
Case Is = 2
  adc("UoYN4") = ""
End Select
End Sub

' ****** Manali 3.5.0 - 06/11/08 - Radio Buttons Added
Private Sub OptBag_GotFocus(Index As Integer)
  DispMsg FraOptBag.ToolTipText, etInfo
End Sub
' ****** Manali 3.5.0 - 06/11/08 - Radio Buttons Added
Private Sub OptBagIn_GotFocus(Index As Integer)
  DispMsg FraOptBagIn.ToolTipText, etInfo
End Sub

' ****** Manali 3.5.0 - 21/11/08 - Radio Buttons Added
Private Sub OptBagPcs_Click(Index As Integer)
Select Case Index
Case Is = 0
  adc("UoBagPcs") = "N"
Case Is = 1
  adc("UoBagPcs") = "R"
End Select
End Sub
' ****** Manali 3.5.0 - 21/11/08 - Radio Buttons Added

' ****** Manali 3.5.0 - 21/11/08 - Radio Buttons Added
Private Sub OptBagPcs_GotFocus(Index As Integer)
DispMsg FraOptBagPcs.ToolTipText, etInfo
End Sub
' ****** Manali 3.5.0 - 21/11/08 - Radio Buttons Added


Private Sub OptProduction_Click()
OptProduction.Value = True
  OptProcess.Value = False
  
 ' UOISSLOCFR.text = ""
'UOISSLOCTO.text = ""
'UOISSLOCSEL.text = ""

End Sub

Private Sub OptProcess_Click()
OptProduction.Value = False
  OptProcess.Value = True
 '  UOISSLOCFR.text = ""
'UOISSLOCTO.text = ""
'UOISSLOCSEL.text = ""
End Sub


Private Sub OptLoc_GotFocus(Index As Integer)
DispMsg FraLoc.ToolTipText, etInfo
Select Case Index
Case Is = 0
  adc("UOISSLOCFR") = "":   adc("UOISSLOCTO") = "":   adc("UOISSLOCsel") = ""
  adc("UoLocTypFr") = "P"
Case Is = 1
  adc("UOISSLOCFR") = "":   adc("UOISSLOCTO") = "":   adc("UOISSLOCsel") = ""
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
   'Call SetGroupSort

  '****** Sachin 2.14.0 - [25-11-2006] - [Multi-Company] ******
  If ws_MultiCoMod = False Then _
      adc("UoCoCdFr") = gs_CoCd: adc("UoCoCdTo") = gs_CoCd: adc("UoCoCdSel") = "": adc("UoCoCdFr").Enabled = False
  '****** Sachin 2.14.0 - [25-11-2006] - [Multi-Company] ******
  
  If UCase(adc.MenuCd) = UCase("RepWipBRm") Then
    adc("UoDtFr") = "01/01/80"
    adc("UoDtTo") = "01/01/80"
  
  '*** (Jen 2.14)
    adc("UoAsOfDt") = "01/01/80"
  ElseIf UCase(adc.MenuCd) = UCase("RepWipBRmBkDt") Then
    adc("UoDtFr") = "01/01/80"
    adc("UoDtTo") = "01/01/80"
    adc("UoAsOfDt") = moCn.SrvrDate
  ElseIf UCase(adc.MenuCd) = UCase("RepWipFgBRm") Then
    adc("UoAsOfDt") = "01/01/80"
    adc("UoDtTo") = moCn.SrvrDate
  '*** (Jen 2.14)
  End If
  adc("UoPmCdFr") = "": adc("UoPmCdTo") = ""    '4.1.4
   '**Geeta*Emr208*Default Show picture is always no
  adc("UoYN2") = "N"
    
  '*** Jay 2.13(CT) ***
  Dim wHeadChk As MwfLib.MDORowSet
  Set wHeadChk = moCn.OpenRes("Select hFluteBagYN, hCompBagsYn from Head where HCoCd= '" + gs_CoCd + "' and " + _
                                                    "HCd= '" + ctSelfCmCd + "' ")
  If wHeadChk!hFluteBagYN = "N" And wHeadChk!hCompBagsYn = "N" Then
    LblBagTyp.Visible = False
    adc("UoBagTyp").Visible = False
    LblUoBChrFr.Visible = False
    adc("UoBChr2Fr").Visible = False: adc("UoBChr2To").Visible = False: adc("UoBChr2Sel").Visible = False
    adc("UoBagTyp") = "P": adc("UoBChr2Fr") = "": adc("UoBChr2To") = "": adc("UoBChr2Sel") = ""
  End If
  Set wHeadChk = Nothing
  '*** Jay 2.13(CT) ***
  
'*** Jay 2.13(CT) ***
'  '*** (Jen 2.11)
'  If moCn.GetFldVal("Select hFluteBagYN from Head where HCoCd= '" + adc("UoCoCdFr") + "' and " + _
'                                                    "HCd= '" + ctSelfCmCd + "' ") = "N" Then
'    LblBagTyp.Visible = False
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
  
  If adc("UoYN4") = "Y" Then
    OptBagIn(0).Value = True
  ElseIf adc("UoYN4") = "N" Then
    OptBagIn(1).Value = True
  ElseIf adc("UoYN4") = "" Then
    OptBagIn(2).Value = True
  End If
  
  If adc("UOBAGPCS") = "N" Then
    OptBagPcs(0).Value = True
  ElseIf adc("UOBAGPCS") = "R" Then
    OptBagPcs(1).Value = True
  End If
  ' ****** Manali 3.5.0 - 06/11/08 - Radio Buttons Added

  ' ****** Manali 3.5.0 - 21/11/08 - Check Boxes added
  If adc("UoYN") = "" Then adc("UoYN") = "N"
  If adc("UoYN2") = "" Then adc("UoYN2") = "N"
  If adc("UoYN3") = "" Then adc("UoYN3") = "N"
  
  If adc("UoYN") = "Y" Then ChkBoxArr(0).Value = Checked Else ChkBoxArr(0).Value = Unchecked
  If adc("UoYN2") = "Y" Then ChkBoxArr(2).Value = Checked Else ChkBoxArr(2).Value = Unchecked
  If adc("UoYN3") = "Y" Then ChkBoxArr(3).Value = Checked Else ChkBoxArr(3).Value = Unchecked
  ' ****** Manali 3.5.0 - 21/11/08 - Check Boxes added
  'StkRt-Set Show StockRate option is No
  adc("UoYN5") = "N"
  
  If adc("UOLOCTYPFR") = "P" Then
    OptLoc(0).Value = True
  ElseIf adc("UOLOCTYPFR") = "R" Then
    OptLoc(1).Value = True
  End If

End Sub
Private Sub DispCoNm()
  'Rep.wCoCd.SetText gs_CoNm
  Dim ws_HName As String
  'ws_HName = moCn.GetFldVal("Select HName from Head where HCoCd ='" + adc("UoCoCdFr") + "' and " + _
                                                    "HCd = '" + ctSelfCmCd + "' ")
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
      Rep.TxtCoCdToLbl.Suppress = False ': Rep.TxtCoCdSelLbl.Suppress = False  '578.99.d CoCd selection removed
      Rep.TxtCoCdFr.SetText adc("UOCOCDFR"): Rep.TxtCoCdTo.SetText adc("UOCOCDTO") ': Rep.TxtCoCdSel.SetText ADC("UOCOCDSEL")   '578.99.d CoCd selection removed
  Else
      Rep.TxtCoCdTo.Suppress = True ': Rep.TxtCoCdSel.Suppress = True   '578.99.d CoCd selection removed
      Rep.TxtCoCdToLbl.Suppress = True: 'Rep.TxtCoCdSelLbl.Suppress = True    '578.99.d CoCd selection removed
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
    ws_Dt = Format$(moCn.SrvrDate, "DD-MMM-YYYY")
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
'    If Trim(TxtPwd.text) = "" Then adc("UoPmCDFr") = ""
'    If Trim(TxtPwd.text) = "" Or Trim(TxtPwd.text) = moCn.GetFldVal("Select PValue1 from Param where " + _
'                           "Ptyp = 'SEOPWD' and PmCd = 'SEOPWD'") Then
'
'      Exit Sub
'    Else
'      Cancel = True
'      DispMsg "Invalid Password", etError
'    End If
' Bef 2.14
'*** Jay 2.14 *** (SEO Change)
End Sub

Private Function RmRate(ByVal mRtFld As String, ByVal mSzFld As String, ByVal mSzFld2 As String, ByVal mQw As String, Optional ByVal mCurrenCy As String = "", Optional ByVal LotValYN As Boolean = False) As String
  Dim wPureGldCd As String, wPurePlCd As String, wPureSlCd As String
  Dim wGRt As String, wPRt As String, wSRt As String, wDCXARt As String
  ' ***** Manali 3.03 - 24/06/08 - Pd Mod - 'L'
  Dim wPurePdCd As String, wLRt As String
  ' ***** Manali 3.03 - 24/06/08 - Pd Mod - 'L'
  
  wPureGldCd = moCn.GetFldVal("Select PValue from param where PTyp= 'RMCTG' and PMCd= 'G'")
  wPurePlCd = moCn.GetFldVal("Select PValue from param where PTyp= 'RMCTG' and PMCd= 'P'")
  wPureSlCd = moCn.GetFldVal("Select PValue from param where PTyp= 'RMCTG' and PMCd= 'S'")
  ' ***** Manali 3.03 - 24/06/08 - Pd Mod - 'L' added
  wPurePdCd = moCn.GetFldVal("Select PValue from param where PTyp= 'RMCTG' and PMCd= 'L'")
  ' ***** Manali 3.03 - 24/06/08 - Pd Mod - 'L' added

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
  If mCurrenCy = "" Then
    ws_BaseCurCd = moCn.GetFldVal("Select HBaseCurCd from Head where HCd= '" + ctSelfCmCd + "' and " + _
                  "HCoCd= '" + adc("UoCoCdFr") + "'")
  Else
    ws_BaseCurCd = mCurrenCy
  End If
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
  ' ***** Manali 3.03 - 24/06/08 - Pd Mod - 'L'
  wLRt = " isnull((select " + mRtFld + " from RmRt where " + IIF(gs_Partition = ctCurrPrtn, " RrPrtKey='" + ctCurrPrtn + "' and ", "") + " RrTcTyp= 'RM' and RrCmCtg= 'C' and " + _
         "RrCd= '" + wPurePdCd + "' and " + _
         "RrFrLn <= " + mSzFld + " And RrToLn >= " + mSzFld + " and RrCmCd = '" + ctValueCmCd + "' and " + _
         "RrCmCurCd= '" + ws_BaseCurCd + "' and RrDmCtg= '' and RrLabMCd= '') , 0) "
  ' ***** Manali 3.03 - 24/06/08 - Pd Mod - 'L'
'*********
 Dim wLotValFldNm As String
  If ws_BaseCurCd = "RS" Then
    wLotValFldNm = "(case when RmQWCst='Q' then LotRsRtByQ  when RmQWCst='W' then LotRsRtByW end) "
  Else
    wLotValFldNm = "(case when RmQWCst='Q' then LotBaseRtByQ when RmQWCst='W' THEN LotBaseRtByW end) "
  End If
  
  ' ***** Manali 3.8.0 - min added - [rectified error - subquery return more than 1 value - in case of rate entered for both sides]
  If LotValYN = True Then
    'BS error
    wDCXARt = " (Case when RSCtg.PValue ='BS' then ( " + _
         " isnull( " + wLotValFldNm + ",isnull((Select " + mRtFld + " from RmRt where " + IIF(gs_Partition = ctCurrPrtn, " RrPrtKey='" + ctCurrPrtn + "' and ", "") + " RrTcTyp='RM' and RrCmCtg= 'C' and " + _
         "RrCtg= RmCtg and RrSCtg= RmSCtg and RrCd= RmCd and RrCmCd = '" + ctValueCmCd + "' and " + _
         "RrCmCurCd= '" + ws_BaseCurCd + "' and " + _
         "RrFrLn = " + mSzFld + " And RrToLn = " + mSzFld2 + " and RrDmCtg= '' and RrLabMCd= '' and " + _
         mRtFld + " <> 0), 0)) ) else " + _
         " (isnull( " + wLotValFldNm + ",isnull((Select " + mRtFld + " from RmRt where " + IIF(gs_Partition = ctCurrPrtn, " RrPrtKey='" + ctCurrPrtn + "' and ", "") + " RrTcTyp='RM' and RrCmCtg= 'C' and " + _
         "RrCtg= RmCtg and RrSCtg= RmSCtg and RrCd= RmCd and RrCmCd = '" + ctValueCmCd + "' and " + _
         "RrCmCurCd= '" + ws_BaseCurCd + "' and " + _
         "RrFrLn <= " + mSzFld + " And RrToLn >= " + mSzFld + " and RrDmCtg= '' and RrLabMCd= '' and " + _
         mRtFld + " <> 0), 0)) ) end) "
    'wDCXARt = " isnull( " + wLotValFldNm + ",isnull((Select " + mRtFld + " from RmRt where " + IIF(gs_Partition = ctCurrPrtn, " RrPrtKey='" + ctCurrPrtn + "' and ", "") + " RrTcTyp='RM' and RrCmCtg= 'C' and " + _
         "RrCtg= RmCtg and RrSCtg= RmSCtg and RrCd= RmCd and RrCmCd = '" + ctValueCmCd + "' and " + _
         "RrCmCurCd= '" + ws_BaseCurCd + "' and " + _
         "RrFrLn <= " + mSzFld + " And RrToLn >= " + mSzFld + " and RrDmCtg= '' and RrLabMCd= '' and " + _
         mRtFld + " <> 0), 0)) "
  Else
    'BS error
    wDCXARt = " (Case when RSCtg.PValue ='BS' then ( " + _
         " isnull((Select " + mRtFld + " from RmRt where " + IIF(gs_Partition = ctCurrPrtn, " RrPrtKey='" + ctCurrPrtn + "' and ", "") + " RrTcTyp='RM' and RrCmCtg= 'C' and " + _
         "RrCtg= RmCtg and RrSCtg= RmSCtg and RrCd= RmCd and RrCmCd = '" + ctValueCmCd + "' and " + _
         "RrCmCurCd= '" + ws_BaseCurCd + "' and " + _
         "RrFrLn = " + mSzFld + " And RrToLn = " + mSzFld2 + " and RrDmCtg= '' and RrLabMCd= '' and " + _
         mRtFld + " <> 0), 0)) else " + _
         " (isnull((Select " + mRtFld + " from RmRt where " + IIF(gs_Partition = ctCurrPrtn, " RrPrtKey='" + ctCurrPrtn + "' and ", "") + " RrTcTyp='RM' and RrCmCtg= 'C' and " + _
         "RrCtg= RmCtg and RrSCtg= RmSCtg and RrCd= RmCd and RrCmCd = '" + ctValueCmCd + "' and " + _
         "RrCmCurCd= '" + ws_BaseCurCd + "' and " + _
         "RrFrLn <= " + mSzFld + " And RrToLn >= " + mSzFld + " and RrDmCtg= '' and RrLabMCd= '' and " + _
         mRtFld + " <> 0), 0)) end) "
    'wDCXARt = " isnull((Select " + mRtFld + " from RmRt where " + IIF(gs_Partition = ctCurrPrtn, " RrPrtKey='" + ctCurrPrtn + "' and ", "") + " RrTcTyp='RM' and RrCmCtg= 'C' and " + _
         "RrCtg= RmCtg and RrSCtg= RmSCtg and RrCd= RmCd and RrCmCd = '" + ctValueCmCd + "' and " + _
         "RrCmCurCd= '" + ws_BaseCurCd + "' and " + _
         "RrFrLn <= " + mSzFld + " And RrToLn >= " + mSzFld + " and RrDmCtg= '' and RrLabMCd= '' and " + _
         mRtFld + " <> 0), 0) "
  End If
  '***** Before 2.13.0
  'RmRate = " (case when RmCtg= 'G' then " + wGRt + " when RmCtg= 'P' then " + wPRt + " " + _
           " when RmCtg= 'S' then " + wSRt + " when RmCtg in ('D', 'C', 'X', 'A') then " + wDCXARt + " else 0 end) "
           
  '****** Sachin 2.13.0 [12-08-2006] [Q/W Module]
  ' ***** Manali 3.03 - 24/06/08 - Pd Mod - 'L' added
  RmRate = " (Case When (RmCtg= 'G' And " + mQw + "='W') Then " + wGRt + " When (RmCtg= 'P' And " + mQw + "='W') then " + wPRt + " " + _
           " when (RmCtg= 'S' And " + mQw + "='W') then " + wSRt + " when (RmCtg= 'L' And " + mQw + "='W') then " + wLRt + " " + _
           " when (RmCtg In ('G','P','S', 'L') And " + mQw + "='Q') Or  (RmCtg in ('D', 'C', 'X', 'A')) then " + wDCXARt + " else 0 end) "
  '****** Sachin 2.13.0 [12-08-2006] [Q/W Module]
           
End Function
Private Function GetCnd(ByVal ws_Typ As String) As String
  '*** (Jen 2.12 Patch 3)
  If UCase(adc.MenuCd) = UCase("RepWipBRm") Then
    'adc("UoIssLocFr").CmpStr = "BLoc >= "
    'adc("UoIssLocTo").CmpStr = "BLoc <= "
    'adc("UoIssLocSel").CmpStr = "BLoc In "
    
    '*** (Jen 2.14)
    adc("UoWrkFr").CmpStr = "BWrk >= "
    adc("UoWrkTo").CmpStr = "BWrk <= "
    adc("UoWrkSel").CmpStr = "BWrk In "
    '*** (Jen 2.14)
    
    adc("UoDtFr").CmpStr = ""
    adc("UoDtTo").CmpStr = ""
    
    '****** Sachin 2.13.0 - [07-07-2006] ******
    adc("UoStnClsFr").CmpStr = "SCls.PSCd>="
    adc("UoStnClsTo").CmpStr = "SCls.PSCd<="
    adc("UoStnClsSel").CmpStr = "SCls.PSCd In"
    '****** Sachin 2.13.0 - [07-07-2006] ******
    
    '****** Sachin 2.14.0 - [Multi-Company] ******
    adc("UoCoCdFr").CmpStr = "BCoCd >="
    adc("UoCoCdTo").CmpStr = "BCoCd <="
    adc("UoCoCdSel").CmpStr = "BCoCd In "
    adc("UoPrntCdFr").CmpStr = "LocPrntCd >="
    adc("UoPrntCdTo").CmpStr = "LocPrntCd <="
    adc("UoPrntCdSel").CmpStr = "LocPrntCd In "
    '****** Sachin 2.14.0 - [Multi-Company] ******
    
    adc("UoRmCtgFr").CmpStr = "RmCtg >="
    adc("UoRmCtgTo").CmpStr = "RmCtg <="
    adc("UoRmCtgSel").CmpStr = "RmCtg In"
    
    adc("UoRmSCtgFr").CmpStr = "RmSCtg >="
    adc("UoRmSCtgTo").CmpStr = "RmSCtg <="
    adc("UoRmSCtgSel").CmpStr = "RmSCtg In"
    
    adc("UoParentCdFr").CmpStr = "RmParentCd >="
    adc("UoParentCdTo").CmpStr = "RmParentCd <="
    adc("UoParentCdSel").CmpStr = "RmParentCd In"
    
    '*** (Jen 2.13)
    adc("UoRngSzFr").CmpStr = "RrRngSz >="
    adc("UoRngSzTo").CmpStr = "RrRngSz <="
    adc("UoRngSzSel").CmpStr = "RrRngSz In"
    '*** (Jen 2.13)
    
    ' ***** Manali 3.10.0 - 07/03/12 - DmCtg added in Scope
    adc("UoDmCtgFr").CmpStr = "DmCtg >="
    adc("UoDmCtgTo").CmpStr = "DmCtg <="
    adc("UoDmCtgSel").CmpStr = "DmCtg In"
    ' ***** Manali 3.10.0 - 07/03/12 - DmCtg added in Scope
    adc("UoDmCdFr").CmpStr = "BOdDmCd >="
    adc("UoDmCdTo").CmpStr = "BOdDmCd <="
    adc("UoDmCdSel").CmpStr = "BOdDmCd In"
    adc("UoOmTcFr").CmpStr = "OdTc >="
    adc("UoOmTcTo").CmpStr = "OdTc <="
    adc("UoOmYyFr").CmpStr = "OdYy >="
    adc("UoOmYyTo").CmpStr = "OdYy <="
    adc("UoOmChrFr").CmpStr = "OdChr >="
    adc("UoOmChrTo").CmpStr = "OdChr <="
    adc("UoOmNoFr").CmpStr = "OdNo >="
    adc("UoOmNoTo").CmpStr = "OdNo <="
    adc("UoOdSrFr").CmpStr = "OdSr >="
    adc("UoOdSrTo").CmpStr = "OdSr <="
    adc("UoCmCdFr").CmpStr = "OmCmCd >="
    adc("UoCmCdTo").CmpStr = "OmCmCd <="
    adc("UoCmCdSel").CmpStr = "OmCmCd In"
    
    adc("UoFgSubLocFr").CmpStr = ""
    adc("UoFgSubLocTo").CmpStr = ""
    adc("UoFgSubLocSel").CmpStr = ""

    adc("UoFdPlYyFr").CmpStr = "": adc("UoFdPlYyTo").CmpStr = ""
    adc("UoFdPlChrFr").CmpStr = "": adc("UoFdPlChrTo").CmpStr = ""
    adc("UoFdPlNoFr").CmpStr = "": adc("UoFdPlNoTo").CmpStr = ""
    
    adc("UoRmCdFr").CmpStr = "TdRmCd >= "
    adc("UoRmCdTo").CmpStr = "TdRmCd <= "
    adc("UoRmCdSel").CmpStr = "TdRmCd in "
    adc("UoLn1Fr").CmpStr = "TdRmSz >= "
    adc("UoLn1To").CmpStr = "TdRmSz <= "
    adc("UoCstRtFr").CmpStr = "TdRmStkRt >= "
    adc("UoCstRtTo").CmpStr = "TdRmStkRt <= "
    ' Zubin 212
    adc("UoLotNoFr").CmpStr = "TdLotNo >= "
    adc("UoLotNoTo").CmpStr = "TdLotNo <= "
    adc("UoLotNoSel").CmpStr = "TdLotNo In "
    ' Zubin 212
    '**********Bhavna src dsg cd
      adc("UoBOEFr").CmpStr = ""
      adc("UoBOETo").CmpStr = ""
      '**********Bhavna src dsg cd
      
  '*** (Jen 2.14)
  ElseIf UCase(adc.MenuCd) = UCase("RepWipBRmBkDt") Then
  'Loc-Prcs- added this consition in query section
    'adc("UoIssLocFr").CmpStr = "IsNull(Lc.LocCd, '') >= "
    'adc("UoIssLocTo").CmpStr = "IsNull(Lc.LocCd, '') <= "
    'adc("UoIssLocSel").CmpStr = "IsNull(Lc.LocCd, '') In "
    
    adc("UoWrkFr").CmpStr = "qTdWrk >= "
    adc("UoWrkTo").CmpStr = "qTdWrk <= "
    adc("UoWrkSel").CmpStr = "qTdWrk In "
    
    adc("UoDtFr").CmpStr = ""
    adc("UoDtTo").CmpStr = ""
    
    adc("UoStnClsFr").CmpStr = "SCls.PSCd >="
    adc("UoStnClsTo").CmpStr = "SCls.PSCd <="
    adc("UoStnClsSel").CmpStr = "SCls.PSCd In"
    
    '****** Sachin 2.14.0 - [Multi-Company] ******
    adc("UoCoCdFr").CmpStr = "BCoCd >="
    adc("UoCoCdTo").CmpStr = "BCoCd <="
    adc("UoCoCdSel").CmpStr = "BCoCd In "
    adc("UoPrntCdFr").CmpStr = "IsNull(Lc.LocPrntCd, '') >="
    adc("UoPrntCdTo").CmpStr = "IsNull(Lc.LocPrntCd, '') <="
    adc("UoPrntCdSel").CmpStr = "IsNull(Lc.LocPrntCd, '') In "
    '****** Sachin 2.14.0 - [Multi-Company] ******
    
    adc("UoRmCtgFr").CmpStr = "RmCtg >="
    adc("UoRmCtgTo").CmpStr = "RmCtg <="
    adc("UoRmCtgSel").CmpStr = "RmCtg In"
    
    adc("UoRmSCtgFr").CmpStr = "RmSCtg >="
    adc("UoRmSCtgTo").CmpStr = "RmSCtg <="
    adc("UoRmSCtgSel").CmpStr = "RmSCtg In"
    
    adc("UoParentCdFr").CmpStr = "RmParentCd >="
    adc("UoParentCdTo").CmpStr = "RmParentCd <="
    adc("UoParentCdSel").CmpStr = "RmParentCd In"
    
    adc("UoRngSzFr").CmpStr = "RrRngSz >="
    adc("UoRngSzTo").CmpStr = "RrRngSz <="
    adc("UoRngSzSel").CmpStr = "RrRngSz In"
    
    ' ***** Manali 3.10.0 - 07/03/12 - DmCtg added in Scope
    adc("UoDmCtgFr").CmpStr = "DmCtg >="
    adc("UoDmCtgTo").CmpStr = "DmCtg <="
    adc("UoDmCtgSel").CmpStr = "DmCtg In"
    ' ***** Manali 3.10.0 - 07/03/12 - DmCtg added in Scope
    adc("UoDmCdFr").CmpStr = "BOdDmCd >="
    adc("UoDmCdTo").CmpStr = "BOdDmCd <="
    adc("UoDmCdSel").CmpStr = "BOdDmCd In"
    adc("UoOmTcFr").CmpStr = "OdTc >="
    adc("UoOmTcTo").CmpStr = "OdTc <="
    adc("UoOmYyFr").CmpStr = "OdYy >="
    adc("UoOmYyTo").CmpStr = "OdYy <="
    adc("UoOmChrFr").CmpStr = "OdChr >="
    adc("UoOmChrTo").CmpStr = "OdChr <="
    adc("UoOmNoFr").CmpStr = "OdNo >="
    adc("UoOmNoTo").CmpStr = "OdNo <="
    adc("UoOdSrFr").CmpStr = "OdSr >="
    adc("UoOdSrTo").CmpStr = "OdSr <="
    adc("UoCmCdFr").CmpStr = "OmCmCd >="
    adc("UoCmCdTo").CmpStr = "OmCmCd <="
    adc("UoCmCdSel").CmpStr = "OmCmCd In"
    
    adc("UoFgSubLocFr").CmpStr = ""
    adc("UoFgSubLocTo").CmpStr = ""
    adc("UoFgSubLocSel").CmpStr = ""

    adc("UoFdPlYyFr").CmpStr = "": adc("UoFdPlYyTo").CmpStr = ""
    adc("UoFdPlChrFr").CmpStr = "": adc("UoFdPlChrTo").CmpStr = ""
    adc("UoFdPlNoFr").CmpStr = "": adc("UoFdPlNoTo").CmpStr = ""
    
    adc("UoRmCdFr").CmpStr = "TR.TdRmCd >= "
    adc("UoRmCdTo").CmpStr = "TR.TdRmCd <= "
    adc("UoRmCdSel").CmpStr = "TR.TdRmCd in "
    adc("UoLn1Fr").CmpStr = "TR.TdRmSz >= "
    adc("UoLn1To").CmpStr = "TR.TdRmSz <= "
    adc("UoCstRtFr").CmpStr = "TR.TdRmStkRt >= "
    adc("UoCstRtTo").CmpStr = "TR.TdRmStkRt <= "
    adc("UoLotNoFr").CmpStr = "TR.TdLotNo >= "
    adc("UoLotNoTo").CmpStr = "TR.TdLotNo <= "
    adc("UoLotNoSel").CmpStr = "TR.TdLotNo In "
  '*** (Jen 2.14)
    '**********Bhavna src dsg cd
      adc("UoBOEFr").CmpStr = ""
      adc("UoBOETo").CmpStr = ""
      '**********Bhavna src dsg cd
      
  ElseIf UCase(adc.MenuCd) = UCase("RepWipFgBRm") Then
    If UCase(ws_Typ) = UCase("Query") Then
      adc("UoIssLocFr").CmpStr = ""
      adc("UoIssLocTo").CmpStr = ""
      adc("UoIssLocSel").CmpStr = ""
      
      '*** (Jen 2.14)
      adc("UoWrkFr").CmpStr = "BWrk >= "
      adc("UoWrkTo").CmpStr = "BWrk <= "
      adc("UoWrkSel").CmpStr = "BWrk In "
      '*** (Jen 2.14)
      
      '****** Sachin 2.13.0 - [07-07-2006] ******
      adc("UoStnClsFr").CmpStr = "SCls.PSCd>="
      adc("UoStnClsTo").CmpStr = "SCls.PSCd<="
      adc("UoStnClsSel").CmpStr = "SCls.PSCd In"
      '****** Sachin 2.13.0 - [07-07-2006] ******
      
      adc("UoDtFr").CmpStr = ""
      adc("UoDtTo").CmpStr = ""
      
      '****** Sachin 2.14.0 - [Multi-Company] ******
      adc("UoCoCdFr").CmpStr = "BCoCd >="
      adc("UoCoCdTo").CmpStr = "BCoCd <="
      adc("UoCoCdSel").CmpStr = "BCoCd In "
      adc("UoPrntCdFr").CmpStr = ""
      adc("UoPrntCdTo").CmpStr = ""
      adc("UoPrntCdSel").CmpStr = ""
      '****** Sachin 2.14.0 - [Multi-Company] ******
      
      adc("UoRmCtgFr").CmpStr = "RmCtg >="
      adc("UoRmCtgTo").CmpStr = "RmCtg <="
      adc("UoRmCtgSel").CmpStr = "RmCtg In"
      
      adc("UoRmSCtgFr").CmpStr = "RmSCtg >="
      adc("UoRmSCtgTo").CmpStr = "RmSCtg <="
      adc("UoRmSCtgSel").CmpStr = "RmSCtg In"
      
      adc("UoParentCdFr").CmpStr = "RmParentCd >="
      adc("UoParentCdTo").CmpStr = "RmParentCd <="
      adc("UoParentCdSel").CmpStr = "RmParentCd In"
      
      '*** (Jen 2.13)
      adc("UoRngSzFr").CmpStr = "RrRngSz >="
      adc("UoRngSzTo").CmpStr = "RrRngSz <="
      adc("UoRngSzSel").CmpStr = "RrRngSz In"
      '*** (Jen 2.13)
      
      ' ***** Manali 3.10.0 - 07/03/12 - DmCtg added in Scope
      adc("UoDmCtgFr").CmpStr = "DmCtg >="
      adc("UoDmCtgTo").CmpStr = "DmCtg <="
      adc("UoDmCtgSel").CmpStr = "DmCtg In"
      ' ***** Manali 3.10.0 - 07/03/12 - DmCtg added in Scope
      adc("UoDmCdFr").CmpStr = "BOdDmCd >="
      adc("UoDmCdTo").CmpStr = "BOdDmCd <="
      adc("UoDmCdSel").CmpStr = "BOdDmCd In"
      adc("UoOmTcFr").CmpStr = "OdTc >="
      adc("UoOmTcTo").CmpStr = "OdTc <="
      adc("UoOmYyFr").CmpStr = "OdYy >="
      adc("UoOmYyTo").CmpStr = "OdYy <="
      adc("UoOmChrFr").CmpStr = "OdChr >="
      adc("UoOmChrTo").CmpStr = "OdChr <="
      adc("UoOmNoFr").CmpStr = "OdNo >="
      adc("UoOmNoTo").CmpStr = "OdNo <="
      adc("UoOdSrFr").CmpStr = "OdSr >="
      adc("UoOdSrTo").CmpStr = "OdSr <="
      adc("UoCmCdFr").CmpStr = "OmCmCd >="
      adc("UoCmCdTo").CmpStr = "OmCmCd <="
      adc("UoCmCdSel").CmpStr = "OmCmCd In"
      
      If adc("UoDtFr") = "01/01/80" And adc("UoDtTo") = "01/01/80" Then
        adc("UoFgSubLocFr").CmpStr = "BFgSubLoc >="
        adc("UoFgSubLocTo").CmpStr = "BFgSubLoc <="
        adc("UoFgSubLocSel").CmpStr = "BFgSubLoc In "
      Else
        adc("UoFgSubLocFr").CmpStr = "qFgSubLoc >="
        adc("UoFgSubLocTo").CmpStr = "qFgSubLoc <="
        adc("UoFgSubLocSel").CmpStr = "qFgSubLoc In "
      End If
      
      adc("UoFdPlYyFr").CmpStr = "FdPlYy >=": adc("UoFdPlYyTo").CmpStr = "FdPlYy <="
      adc("UoFdPlChrFr").CmpStr = "FdPlChr >=": adc("UoFdPlChrTo").CmpStr = "FdPlChr <="
      adc("UoFdPlNoFr").CmpStr = "FdPlNo >=": adc("UoFdPlNoTo").CmpStr = "FdPlNo <="
      
      adc("UoRmCdFr").CmpStr = "FrRmCd >= "
      adc("UoRmCdTo").CmpStr = "FrRmCd <= "
      adc("UoRmCdSel").CmpStr = "FrRmCd in "
      adc("UoLn1Fr").CmpStr = "FrRmSz >= "
      adc("UoLn1To").CmpStr = "FrRmSz <= "
      adc("UoCstRtFr").CmpStr = "FrRmStkRt >= "
      adc("UoCstRtTo").CmpStr = "FrRmStkRt <= "
      ' Zubin 212
      adc("UoLotNoFr").CmpStr = "FrLotNo >= "
      adc("UoLotNoTo").CmpStr = "FrLotNo <= "
      adc("UoLotNoSel").CmpStr = "FrLotNo In "
      ' Zubin 212
      '**********Bhavna src dsg cd
      adc("UoBOEFr").CmpStr = "DmSrcDsgCd >= "
      adc("UoBOETo").CmpStr = "DmSrcDsgCd <= "
      '**********Bhavna src dsg cd
      
    ElseIf UCase(ws_Typ) = UCase("SubQuery") Then
      adc("UoIssLocFr").CmpStr = ""
      adc("UoIssLocTo").CmpStr = ""
      adc("UoIssLocSel").CmpStr = ""
      
      '*** (Jen 2.14)
      adc("UoWrkFr").CmpStr = ""
      adc("UoWrkTo").CmpStr = ""
      adc("UoWrkSel").CmpStr = ""
      '*** (Jen 2.14)
      
      '****** Sachin 2.13.0 - [07-07-2006] ******
      adc("UoStnClsFr").CmpStr = ""
      adc("UoStnClsTo").CmpStr = ""
      adc("UoStnClsSel").CmpStr = ""
      '****** Sachin 2.13.0 - [07-07-2006] ******
      
      adc("UoDtFr").CmpStr = "FgDt>= "
      adc("UoDtTo").CmpStr = "FgDt<= "
      
      '****** Sachin 2.14.0 - [Multi-Company] ******
      adc("UoCoCdFr").CmpStr = "FdCoCd >="
      adc("UoCoCdTo").CmpStr = "FdCoCd <="
      adc("UoCoCdSel").CmpStr = "FdCoCd In "
      adc("UoPrntCdFr").CmpStr = ""
      adc("UoPrntCdTo").CmpStr = ""
      adc("UoPrntCdSel").CmpStr = ""
      '****** Sachin 2.14.0 - [Multi-Company] ******
      
      adc("UoRmCtgFr").CmpStr = ""
      adc("UoRmCtgTo").CmpStr = ""
      adc("UoRmCtgSel").CmpStr = ""
      
      adc("UoRmSCtgFr").CmpStr = ""
      adc("UoRmSCtgTo").CmpStr = ""
      adc("UoRmSCtgSel").CmpStr = ""
      
      adc("UoParentCdFr").CmpStr = ""
      adc("UoParentCdTo").CmpStr = ""
      adc("UoParentCdSel").CmpStr = ""
      
      '*** (Jen 2.13)
      adc("UoRngSzFr").CmpStr = ""
      adc("UoRngSzTo").CmpStr = ""
      adc("UoRngSzSel").CmpStr = ""
      '*** (Jen 2.13)
      
      ' ***** Manali 3.10.0 - 07/03/12 - DmCtg added in Scope
      adc("UoDmCtgFr").CmpStr = ""
      adc("UoDmCtgTo").CmpStr = ""
      adc("UoDmCtgSel").CmpStr = ""
      ' ***** Manali 3.10.0 - 07/03/12 - DmCtg added in Scope
      
      adc("UoDmCdFr").CmpStr = "FdDmCd >="
      adc("UoDmCdTo").CmpStr = "FdDmCd <="
      adc("UoDmCdSel").CmpStr = "FdDmCd In"
      adc("UoOmTcFr").CmpStr = "FdPrdOdTc >="
      adc("UoOmTcTo").CmpStr = "FdPrdOdTc <="
      adc("UoOmYyFr").CmpStr = "FdPrdOdYy >="
      adc("UoOmYyTo").CmpStr = "FdPrdOdYy <="
      adc("UoOmChrFr").CmpStr = "FdPrdOdChr >="
      adc("UoOmChrTo").CmpStr = "FdPrdOdChr <="
      adc("UoOmNoFr").CmpStr = "FdPrdOdNo >="
      adc("UoOmNoTo").CmpStr = "FdPrdOdNo <="
      adc("UoOdSrFr").CmpStr = "FdPrdOdSr >="
      adc("UoOdSrTo").CmpStr = "FdPrdOdSr <="
      adc("UoCmCdFr").CmpStr = "FdPrdCmCd >="
      adc("UoCmCdTo").CmpStr = "FdPrdCmCd <="
      adc("UoCmCdSel").CmpStr = "FdPrdCmCd In"
      
      adc("UoFgSubLocFr").CmpStr = ""
      adc("UoFgSubLocTo").CmpStr = ""
      adc("UoFgSubLocSel").CmpStr = ""
      
      adc("UoFdPlYyFr").CmpStr = "": adc("UoFdPlYyTo").CmpStr = ""
      adc("UoFdPlChrFr").CmpStr = "": adc("UoFdPlChrTo").CmpStr = ""
      adc("UoFdPlNoFr").CmpStr = "": adc("UoFdPlNoTo").CmpStr = ""
      
      adc("UoRmCdFr").CmpStr = ""
      adc("UoRmCdTo").CmpStr = ""
      adc("UoRmCdSel").CmpStr = ""
      adc("UoLn1Fr").CmpStr = ""
      adc("UoLn1To").CmpStr = ""
      adc("UoCstRtFr").CmpStr = ""
      adc("UoCstRtTo").CmpStr = ""
      adc("UoLotNoFr").CmpStr = ""
      adc("UoLotNoTo").CmpStr = ""
      adc("UoLotNoSel").CmpStr = ""
            
      adc("UoBYyFr").CmpStr = "FdBYy>="
      adc("UoBYyTo").CmpStr = "FdBYy<="
      adc("UoBChrFr").CmpStr = "FdBChr>="
      adc("UoBChrTo").CmpStr = "FdBChr<="
      adc("UoBNoFr").CmpStr = "FdBNo>="
      adc("UoBNoTo").CmpStr = "FdBNo<="
      
      '**********Bhavna src dsg cd
      adc("UoBOEFr").CmpStr = ""
      adc("UoBOETo").CmpStr = ""
      '**********Bhavna src dsg cd
      
    End If
  End If

  GetCnd = adc.RepCond
  'IIF(ws_RepCnd <> "", " and " + ws_RepCnd, "")

  '*** (Jen 2.12 Patch 3)
End Function

'StkRt-if password is change to blank, Show StockRate check box will turn into blank
Private Sub TxtPwd_Change()
    If Trim(TxtPwd.text) = "" Then
        ChkBoxArr(4).Value = Unchecked
    End If
End Sub




