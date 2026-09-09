VERSION 5.00
Object = "{8767A745-088E-4CA6-8594-073D6D2DE57A}#9.2#0"; "crviewer9.dll"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "tabctl32.ocx"
Object = "{AB2CC32A-E652-42A2-B341-26ECE02F6386}#2.1#0"; "mwfctlforspeed9.ocx"
Begin VB.Form EmrFrpFgBagLed 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Finished Goods Bag ledger"
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
   Begin MwfCtl.MWCTL_ADC adc 
      Height          =   555
      Left            =   60
      TabIndex        =   29
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
   Begin MwfCtl.MWCTL_HLP MWCTL_HLP1 
      Height          =   495
      Left            =   11850
      TabIndex        =   28
      TabStop         =   0   'False
      Top             =   9690
      Visible         =   0   'False
      Width           =   495
      _ExtentX        =   873
      _ExtentY        =   873
   End
   Begin MwfCtl.MWCTL_DST DST_REP 
      Height          =   570
      Left            =   4455
      TabIndex        =   27
      TabStop         =   0   'False
      Top             =   9630
      Width           =   6330
      _ExtentX        =   11165
      _ExtentY        =   1005
   End
   Begin VB.Frame fra_keys 
      BorderStyle     =   0  'None
      Height          =   375
      Left            =   5805
      TabIndex        =   37
      Top             =   9690
      Width           =   3555
      Begin MwfCtl.MWCTL_MED aktxt 
         Height          =   345
         Index           =   0
         Left            =   210
         TabIndex        =   38
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
         TabIndex        =   39
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
      Left            =   7020
      TabIndex        =   36
      Top             =   9615
      Visible         =   0   'False
      Width           =   855
      Begin CRVIEWER9LibCtl.CRViewer9 CRV_REP 
         Height          =   495
         Left            =   480
         TabIndex        =   33
         Top             =   -165
         Visible         =   0   'False
         Width           =   495
         lastProp        =   600
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
      TabIndex        =   32
      Top             =   60
      Width           =   15255
      Begin TabDlg.SSTab TAB_REP 
         Height          =   9465
         Left            =   60
         TabIndex        =   31
         Top             =   30
         Width           =   15195
         _ExtentX        =   26802
         _ExtentY        =   16695
         _Version        =   393216
         Tabs            =   2
         TabsPerRow      =   2
         TabHeight       =   520
         TabCaption(0)   =   "&Data Range"
         TabPicture(0)   =   "EmrFrpFgBagLed.frx":0000
         Tab(0).ControlEnabled=   -1  'True
         Tab(0).Control(0)=   "fra_tabrep(0)"
         Tab(0).Control(0).Enabled=   0   'False
         Tab(0).ControlCount=   1
         TabCaption(1)   =   "&Group Sort"
         TabPicture(1)   =   "EmrFrpFgBagLed.frx":001C
         Tab(1).ControlEnabled=   0   'False
         Tab(1).Control(0)=   "fra_tabrep(1)"
         Tab(1).ControlCount=   1
         Begin VB.Frame fra_tabrep 
            BorderStyle     =   0  'None
            Enabled         =   0   'False
            Height          =   9070
            Index           =   1
            Left            =   -74940
            TabIndex        =   35
            Top             =   360
            Width           =   15105
            Begin MwfCtl.MWCTL_GRP GRP_REP 
               Height          =   360
               Left            =   240
               TabIndex        =   30
               Top             =   360
               Width           =   8295
               _ExtentX        =   14631
               _ExtentY        =   635
            End
         End
         Begin VB.Frame fra_tabrep 
            BorderStyle     =   0  'None
            Caption         =   "S"
            Height          =   9070
            Index           =   0
            Left            =   60
            TabIndex        =   34
            Top             =   360
            Width           =   15105
            Begin VB.CheckBox ChkBoxArr 
               Alignment       =   1  'Right Justify
               Caption         =   "Show Box        "
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
               TabIndex        =   25
               Tag             =   "ShowBoxDet"
               ToolTipText     =   "Check To Show box Details"
               Top             =   3060
               Width           =   1710
            End
            Begin MwfCtl.MWCTL_MED atxt 
               Height          =   285
               Index           =   28
               Left            =   4215
               TabIndex        =   2
               ToolTipText     =   "Enter To Design Category"
               Top             =   1065
               Width           =   885
               _ExtentX        =   1561
               _ExtentY        =   503
               MaxLength       =   5
               DataType        =   4
               DataField       =   "UODMCTGTO"
               IdName          =   "UODMCTGTO"
            End
            Begin MwfCtl.MWCTL_MED atxt 
               Height          =   285
               Index           =   29
               Left            =   6900
               TabIndex        =   3
               ToolTipText     =   "Enter Design Category Selection"
               Top             =   1065
               Width           =   7950
               _ExtentX        =   14023
               _ExtentY        =   503
               MaxLength       =   255
               DataType        =   4
               DataField       =   "UODMCTGSEL"
               IdName          =   "UODMCTGSEL"
            End
            Begin MwfCtl.MWCTL_MED atxt 
               Height          =   285
               Index           =   31
               Left            =   4215
               TabIndex        =   5
               ToolTipText     =   "Enter To Design Code"
               Top             =   1350
               Width           =   2385
               _ExtentX        =   4207
               _ExtentY        =   503
               MaxLength       =   15
               DataField       =   "UODMCDTO"
               IdName          =   "UODMCDTO"
            End
            Begin MwfCtl.MWCTL_MED atxt 
               Height          =   285
               Index           =   32
               Left            =   6900
               TabIndex        =   6
               ToolTipText     =   "Enter Design Code Selection"
               Top             =   1350
               Width           =   7950
               _ExtentX        =   14023
               _ExtentY        =   503
               MaxLength       =   255
               DataField       =   "UODMCDSEL"
               IdName          =   "UODMCDSEL"
            End
            Begin MwfCtl.MWCTL_MED atxt 
               Height          =   285
               Index           =   33
               Left            =   1515
               TabIndex        =   4
               ToolTipText     =   "Enter From Design Code"
               Top             =   1350
               Width           =   2385
               _ExtentX        =   4207
               _ExtentY        =   503
               MaxLength       =   15
               DataField       =   "UODMCDFR"
               IdName          =   "UODMCDFR"
            End
            Begin MwfCtl.MWCTL_MED atxt 
               Height          =   285
               Index           =   40
               Left            =   4215
               TabIndex        =   21
               ToolTipText     =   "Enter To Customer"
               Top             =   2490
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
               Left            =   6900
               TabIndex        =   22
               ToolTipText     =   "Enter Customer Selection"
               Top             =   2490
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
               Left            =   1515
               TabIndex        =   20
               ToolTipText     =   "Enter From Customer"
               Top             =   2490
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
               Index           =   5
               Left            =   4995
               TabIndex        =   14
               ToolTipText     =   "Enter To Order Year"
               Top             =   1920
               Width           =   345
               _ExtentX        =   609
               _ExtentY        =   503
               MaxLength       =   2
               DataField       =   "UOOMYYTO"
               IdName          =   "UOOMYYTO"
            End
            Begin MwfCtl.MWCTL_MED atxt 
               Height          =   285
               Index           =   20
               Left            =   2295
               TabIndex        =   11
               ToolTipText     =   "Enter From Order Year"
               Top             =   1920
               Width           =   345
               _ExtentX        =   609
               _ExtentY        =   503
               MaxLength       =   2
               DataField       =   "UOOMYYFR"
               IdName          =   "UOOMYYFR"
            End
            Begin MwfCtl.MWCTL_MED atxt 
               Height          =   285
               Index           =   10
               Left            =   1515
               TabIndex        =   10
               ToolTipText     =   "Enter From Order Tc"
               Top             =   1920
               Width           =   705
               _ExtentX        =   1244
               _ExtentY        =   503
               MaxLength       =   3
               DataField       =   "UOOMTCFR"
               IdName          =   "UOOMTCFR"
            End
            Begin MwfCtl.MWCTL_MED atxt 
               Height          =   285
               Index           =   13
               Left            =   2715
               TabIndex        =   12
               ToolTipText     =   "Enter From Order Character"
               Top             =   1920
               Width           =   705
               _ExtentX        =   1244
               _ExtentY        =   503
               MaxLength       =   3
               DataType        =   4
               DataField       =   "UOOMCHRFR"
               IdName          =   "UOOMCHRFR"
            End
            Begin MwfCtl.MWCTL_MED atxt 
               Height          =   285
               Index           =   17
               Left            =   4215
               TabIndex        =   13
               ToolTipText     =   "Enter To Order Tc"
               Top             =   1920
               Width           =   705
               _ExtentX        =   1244
               _ExtentY        =   503
               MaxLength       =   3
               DataField       =   "UOOMTCTO"
               IdName          =   "UOOMTCTO"
            End
            Begin MwfCtl.MWCTL_MED atxt 
               Height          =   285
               Index           =   23
               Left            =   5415
               TabIndex        =   15
               ToolTipText     =   "Enter To Order Character"
               Top             =   1920
               Width           =   705
               _ExtentX        =   1244
               _ExtentY        =   503
               MaxLength       =   3
               DataType        =   4
               DataField       =   "UOOMCHRTO"
               IdName          =   "UOOMCHRTO"
            End
            Begin MwfCtl.MWCTL_MED atxt 
               Height          =   285
               Index           =   53
               Left            =   1515
               TabIndex        =   16
               ToolTipText     =   "Enter From Order Number"
               Top             =   2205
               Width           =   825
               _ExtentX        =   1455
               _ExtentY        =   503
               MaxLength       =   0
               DataType        =   1
               DataField       =   "UOOMNOFR"
               IdName          =   "UOOMNOFR"
            End
            Begin MwfCtl.MWCTL_MED atxt 
               Height          =   285
               Index           =   54
               Left            =   4215
               TabIndex        =   18
               ToolTipText     =   "Enter To Order Number"
               Top             =   2205
               Width           =   825
               _ExtentX        =   1455
               _ExtentY        =   503
               MaxLength       =   0
               DataType        =   1
               DataField       =   "UOOMNOTO"
               IdName          =   "UOOMNOTO"
            End
            Begin MwfCtl.MWCTL_MED atxt 
               Height          =   285
               Index           =   15
               Left            =   4215
               TabIndex        =   8
               ToolTipText     =   "Enter To Karat"
               Top             =   1635
               Width           =   885
               _ExtentX        =   1561
               _ExtentY        =   503
               MaxLength       =   5
               DataField       =   "UOKTTO"
               IdName          =   "UOKTTO"
            End
            Begin MwfCtl.MWCTL_MED atxt 
               Height          =   285
               Index           =   16
               Left            =   6900
               TabIndex        =   9
               ToolTipText     =   "Enter Karat Selection"
               Top             =   1635
               Width           =   7950
               _ExtentX        =   14023
               _ExtentY        =   503
               MaxLength       =   255
               DataField       =   "UOKTSEL"
               IdName          =   "UOKTSEL"
            End
            Begin MwfCtl.MWCTL_MED atxt 
               Height          =   285
               Index           =   18
               Left            =   1515
               TabIndex        =   7
               ToolTipText     =   "Enter From Karat"
               Top             =   1635
               Width           =   885
               _ExtentX        =   1561
               _ExtentY        =   503
               MaxLength       =   5
               DataField       =   "UOKTFR"
               IdName          =   "UOKTFR"
            End
            Begin MwfCtl.MWCTL_MED atxt 
               Height          =   285
               Index           =   21
               Left            =   5100
               TabIndex        =   19
               ToolTipText     =   "Enter To Order Serial"
               Top             =   2205
               Width           =   525
               _ExtentX        =   926
               _ExtentY        =   503
               MaxLength       =   0
               DataType        =   1
               DataField       =   "UOODSRTO"
               IdName          =   "UOODSRTO"
            End
            Begin MwfCtl.MWCTL_MED atxt 
               Height          =   285
               Index           =   4
               Left            =   2400
               TabIndex        =   17
               ToolTipText     =   "Enter From Order Serial"
               Top             =   2205
               Width           =   525
               _ExtentX        =   926
               _ExtentY        =   503
               MaxLength       =   0
               DataType        =   1
               DataField       =   "UOODSRFR"
               IdName          =   "UOODSRFR"
            End
            Begin MwfCtl.MWCTL_MED atxt 
               Height          =   285
               Index           =   48
               Left            =   1515
               TabIndex        =   23
               ToolTipText     =   "Enter From  Date"
               Top             =   2775
               Width           =   885
               _ExtentX        =   1561
               _ExtentY        =   503
               MaxLength       =   8
               DataType        =   3
               DataField       =   "UODOCDTFR"
               IdName          =   "UODOCDTFR"
            End
            Begin MwfCtl.MWCTL_MED atxt 
               Height          =   285
               Index           =   2
               Left            =   2670
               TabIndex        =   24
               ToolTipText     =   "Show Box Detail (Yes / No)"
               Top             =   3060
               Visible         =   0   'False
               Width           =   285
               _ExtentX        =   503
               _ExtentY        =   503
               Enabled         =   0   'False
               DataType        =   4
               DataField       =   "UOYN"
               ReCalcParent    =   "UOPMCDFR"
               IdName          =   "UOYN"
            End
            Begin MwfCtl.MWCTL_MED atxt 
               Height          =   285
               Index           =   30
               Left            =   1515
               TabIndex        =   1
               ToolTipText     =   "Enter From Design Category"
               Top             =   1065
               Width           =   885
               _ExtentX        =   1561
               _ExtentY        =   503
               MaxLength       =   5
               DataType        =   4
               DataField       =   "UODMCTGFR"
               IdName          =   "UODMCTGFR"
            End
            Begin MwfCtl.MWCTL_MED atxt 
               Height          =   285
               Index           =   0
               Left            =   1515
               TabIndex        =   0
               ToolTipText     =   "Enter Company Code"
               Top             =   780
               Width           =   705
               _ExtentX        =   1244
               _ExtentY        =   503
               Enabled         =   0   'False
               MaxLength       =   3
               DataType        =   4
               DataField       =   "UOCOCDFR"
               IdName          =   "UOCOCDFR"
            End
            Begin MwfCtl.MWCTL_MED atxt 
               Height          =   285
               Index           =   1
               Left            =   1515
               TabIndex        =   26
               ToolTipText     =   "Show Box Detail  Bag Wise /Transaction Wise"
               Top             =   3345
               Width           =   885
               _ExtentX        =   1561
               _ExtentY        =   503
               DataType        =   4
               DataField       =   "UOPMCDFR"
               ReCalcOn        =   "UOYN"
               IdName          =   "UOPMCDFR"
            End
            Begin VB.Label ALBL 
               BackStyle       =   0  'Transparent
               Caption         =   "Box Sort"
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00000000&
               Height          =   345
               Index           =   0
               Left            =   60
               TabIndex        =   58
               Top             =   3345
               Width           =   1545
            End
            Begin VB.Label ALBL 
               BackStyle       =   0  'Transparent
               Caption         =   "Show Box"
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
               Height          =   255
               Index           =   4
               Left            =   60
               TabIndex        =   57
               Top             =   3060
               Visible         =   0   'False
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
               Index           =   15
               Left            =   60
               TabIndex        =   56
               Top             =   780
               Width           =   1545
            End
            Begin VB.Label ALBL 
               BackColor       =   &H00C0C0C0&
               BackStyle       =   0  'Transparent
               Caption         =   "From Date"
               BeginProperty Font 
                  Name            =   "Arial"
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
               TabIndex        =   55
               Top             =   2775
               Width           =   1545
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
               Left            =   2340
               TabIndex        =   54
               Top             =   2205
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
               Left            =   5040
               TabIndex        =   53
               Top             =   2205
               Width           =   105
            End
            Begin VB.Label ALBL 
               BackColor       =   &H00C0C0C0&
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
               TabIndex        =   52
               Top             =   1635
               Width           =   1545
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
               TabIndex        =   51
               Top             =   1920
               Width           =   1545
            End
            Begin VB.Label ALBL 
               BackColor       =   &H00C0C0C0&
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
               TabIndex        =   50
               Top             =   2205
               Width           =   1545
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
               Left            =   5355
               TabIndex        =   49
               Top             =   1920
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
               Left            =   2655
               TabIndex        =   48
               Top             =   1920
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
               Left            =   4935
               TabIndex        =   47
               Top             =   1920
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
               Left            =   2235
               TabIndex        =   46
               Top             =   1920
               Width           =   105
            End
            Begin VB.Label ALBL 
               BackColor       =   &H00C0C0C0&
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
               TabIndex        =   45
               Top             =   2490
               Width           =   1545
            End
            Begin VB.Label ALBL 
               BackColor       =   &H00C0C0C0&
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
               TabIndex        =   44
               Top             =   1350
               Width           =   1545
            End
            Begin VB.Label ALBL 
               BackColor       =   &H00C0C0C0&
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
               TabIndex        =   43
               Top             =   1035
               Width           =   1545
            End
            Begin VB.Label ALBL 
               BackColor       =   &H00C0C0C0&
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
               TabIndex        =   42
               Top             =   120
               Width           =   1575
            End
            Begin VB.Label ALBL 
               BackColor       =   &H00C0C0C0&
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
               TabIndex        =   41
               Top             =   120
               Width           =   1275
            End
            Begin VB.Label ALBL 
               BackColor       =   &H00C0C0C0&
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
               Left            =   1515
               TabIndex        =   40
               Top             =   120
               Width           =   1335
            End
         End
      End
   End
End
Attribute VB_Name = "EmrFrpFgBagLed"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim Rep As New EmrRepFgBagLed
Dim moCn As MwfLib.MDOConnection   '*** (Bef speed) bef 24/11/06
Private Sub SetGroupSort()
  GRP_REP.Add "Design Ctg", "DmCtg", "", "DmCtg", "hDmCtg", "", "", "", "(Select PDesc from Param where PTyp= 'DMCTG' and PMCd= DmCtg) "
  GRP_REP.Add "Design Code", "OdDmCd", "", "OdDmCd", "hOdDmCd", "Design Ctg", "", "", "Substring(DmDesc,1,120)"
  GRP_REP.Add "Ord Kt", "OdKt", "", "OdKt", "hOdKt", "", "", "", "(Select PDesc from Param where PTyp= 'KT' and PMCd= OdKt) "
  GRP_REP.Add "Customer Code", "OmCmCd", "", "OmCmCd", "hOmCmCd", "", "", "", "(Select CmName From CustMst Where CmCtg= 'C' and CmCd= OmCmCd) "
  GRP_REP.Add "Colour", "OdDmCol", "OdDmCol", "OdDmCol", "hOdDmCol", "", "", "", "(Select PDesc From Param Where PTyp= 'DMCOL' and PMCd = OdDmCol)"
End Sub
Private Sub ADC_NonKeyWhen(pr_Cancel As Boolean, pr_ErrMsg As String)
adc("UoCoCdFr") = gs_CoCd

 ' ***** Manali 3.5.0 - 21/11/08 - Check Boxes Added
 If adc("UoYN") = "Y" Then ChkBoxArr(0).Value = Checked Else ChkBoxArr(0).Value = Unchecked
 ' ***** Manali 3.5.0 - 21/11/08 - Check Boxes Added
End Sub
Private Sub adc_setreprecsource()
On Error GoTo RepErr
  '*** Report Sql ***
  'Select space(30) as Grp1, space(30) as Grp2, space(30) as Grp3,
  'space(30) as DGrp1, space(30) as DGrp2, space(30) as DGrp3,
  'space(30) as G1Desc, space(30) as G2Desc, space(30) as G3Desc,
  'space(5) as qTag,PValue as qTcTyp,FgDt as qDt,BCoCd,fdTc as qTc,FdYy as qYy,
  'FdChr as qChr,FdNo as qNo,FdSr as qSr,Byy,BChr,BNo,BQty,FdYy + STR(FdKey, 7) as RefKey,
  'BOdTc , BOdYy,BOdChr,BOdNo, BOdSr,DmCtg,OdDmCd,DmDesc, OdKt, OdDmCol, OmCmCd
  'from Bag,Fgd,InvFgd,Fg,InvHd,Param,OrdDsg,OrdMst,DsgMst where 1=2
  
  Dim grpflds As String, wrepcnd As String, wCnd As String
  Dim wCnd1 As String, wSqlStrg As String
  Dim wStr As String
  
  Dim wAddlCnd As String '****** Sachin 3.02.0 - Additional Fields In Where Clause
  
  Set Rep = Nothing
  Set adc.RepSource = Rep
  Call DispCoNm
  Call SetMwName(Rep)   ' ***** Manali 3.6.1 - Microway footer
  Call GRP_REP.Gen3LRep
  grpflds = GRP_REP.GrpFldLst
  
  '*** To check if Grp1, Grp2 is selected and also to check if Detail is selected ***
  Dim wGrp(6) As String, i As Integer, wDetPos As Integer, wNonePos As Integer
  wGrp(1) = GRP_REP.Value(gltGroup1)
  wGrp(2) = GRP_REP.Value(gltGroup2)
  wGrp(3) = GRP_REP.Value(gltGroup3)
  wDetPos = 0: wNonePos = 0
  For i = 1 To 6
    If wGrp(i) = "(Detail)" Then
      wDetPos = i: Exit For
    End If
    If wGrp(i) = "(None)" Then
      wNonePos = i: Exit For
    End If
  Next i
  
  If wDetPos = 0 Then
    Rep.BagHd.Suppress = True
    Rep.Bag.Suppress = True
    Rep.BagDet.Suppress = True
    Rep.BagFt.Suppress = True
  End If
  
  adc("UoCoCdFr").CmpStr = "BCoCd= "
  adc("UoOmTcFr").CmpStr = "BOdTc>= ": adc("UoOmTcTo").CmpStr = "BOdTc<= "
  adc("UoOmYyFr").CmpStr = "BOdYy>= ": adc("UoOmYyTo").CmpStr = "BOdYy<= "
  adc("UoOmChrFr").CmpStr = "BOdChr>=": adc("UoOmChrTo").CmpStr = "BOdChr<= "
  adc("UoOmNoFr").CmpStr = "BOdNo>= ": adc("UoOmNoTo").CmpStr = "BOdNo<= "
  adc("UoOdSrFr").CmpStr = "BOdSr>= ": adc("UoOdSrTo").CmpStr = "BOdSr<= "
  adc("UoCmCdFr").CmpStr = "": adc("UoCmCdTo").CmpStr = "": adc("UoCmCdSel").CmpStr = ""
  adc("UoDmCtgFr").CmpStr = "": adc("UoDmCtgTo").CmpStr = "": adc("UoDmCtgSel").CmpStr = ""
  adc("UoKtFr").CmpStr = "": adc("UoKtTo").CmpStr = "": adc("UoKtSel").CmpStr = ""
  adc("UoDmCdFr").CmpStr = "": adc("UoDmCdTo").CmpStr = "": adc("UoDmCdSel").CmpStr = ""
  
  wrepcnd = adc.RepCond
  wCnd = IIF(wrepcnd <> "", " and " + wrepcnd, "")
  
  adc("UoCoCdFr").CmpStr = "FdCoCd= "
  adc("UoOmTcFr").CmpStr = "FdPrdOdTc>= "
  adc("UoOmTcTo").CmpStr = "FdPrdOdTc<= "
  adc("UoOmYyFr").CmpStr = "FdPrdOdYy>= "
  adc("UoOmYyTo").CmpStr = "FdPrdOdYy<= "
  adc("UoOmChrFr").CmpStr = "FdPrdOdChr>= "
  adc("UoOmChrTo").CmpStr = "FdPrdOdChr<= "
  adc("UoOmNoFr").CmpStr = "FdPrdOdNo>= "
  adc("UoOmNoTo").CmpStr = "FdPrdOdNo<= "
  adc("UoOdSrFr").CmpStr = "FdPrdOdSr>= "
  adc("UoOdSrTo").CmpStr = "FdPrdOdSr<= "
  wrepcnd = adc.RepCond
  wCnd1 = IIF(wrepcnd <> "", " and " + wrepcnd, "")
      
  Call ApportnTbl
     
 ' wStr = "Set DateFormat DMY select * into  " + wTmpTbl + "  from  "
  '****** Sachin 3.02 - Id fields in Joins
  ' ***** Manali 3.9.1 - IfTc='IN' clause added for CTB
  '578-34 STR(FdKey, 7) changed to STR(FdKey, 8)
  '6.1
  wStr = "Set DateFormat " + DtFmtStr() + "  insert into #TmpFgBag " + _
          " select '4Cls' as Tag,BClsDt as Dt,BCoCd as BCoCd ,'' as Tc,'' as Yy,'' as Chr," + _
          "0 as TNo,0 as Sr,BYy as Byy,BChr as BChr,BNo as BNo ,BQty,'' as RefKey,BPrtKey,BOdTc,BOdYy,BOdChr, " + _
          "BOdNo,BOdSr from  Bag where " + IIF(gs_Partition = ctCurrPrtn, " BPrtKey='" + ctCurrPrtn + "' And ", "") + " BLoc='PFG' and BCls ='N'  " + wCnd + _
          " Union All " + _
          "select '2Fg',FgDt,FgCoCd,FgTc,FgYy,FgChr,FgNo,FdSr,FdBYy,FdBChr,FdBNo,FdQty,FdYy + STR(FdKey, 8), " + _
          "FdPrtKey,FdPrdOdTc,FdPrdOdYy,FdPrdOdChr,FdPrdOdNo,FdPrdOdSr from Fgd " + _
          "JOIN  Fg ON FgIdNo=FdFgIdNo and FgCoCd = FdCoCd And FgTc = FdTc And FgYy = FdYy And FgChr = FdChr And FgNo = FdNo " + IIF(gs_Partition = ctCurrPrtn, " and FgPrtKey=FdPrtKey ", "") + _
          " WHERE " + IIF(gs_Partition = ctCurrPrtn, " FgPrtKey='" + ctCurrPrtn + "' And ", "") + " FgDt  Between '" + CStr(adc("UoDocDtFr")) + "' and getdate() " + _
          " And FdDt  Between '" + CStr(adc("UoDocDtFr")) + "' and getdate() " + wCnd1 + _
          " Union All " + _
          "select '3Inv',InDt,IfCocd,IfTc,IfYy,IfChr,IfNo,IfSr,IfFdByy,IfFdBChr,IfFdBNo,IfQty,'', " + _
          "BPrtKey,BOdTc,BOdYy,BOdChr,BOdNo,BOdSr from InvFgd " + _
          "Join InvHd On IfInIdNo=InIdNo and IfCoCd=InCocd And IfYy=InYy and IfChr=InChr And IfNo=InNo " + IIF(gs_Partition = ctCurrPrtn, " and IfPrtKey=InPrtKey ", "") + _
          "Join Bag on BIdNo=IfBIdNo and BCocd=IfCocd and BYy=IfFdBYy and BChr=IfFdBChr and BNo=IfFdBNo " + IIF(gs_Partition = ctCurrPrtn, " and BPrtKey=IfPrtKey ", "") + _
          " where (Select PValue From Param Where PTyp='TC' and PMCd=IfTc)='" + ctInTcTyp + "' and " + IIF(gs_Partition = ctCurrPrtn, " IfPrtKey='" + ctCurrPrtn + "' And ", "") + " InDt  between '" + CStr(adc("UoDocDtFr")) + "' and getdate() " + wCnd

  moCn.Execute (wStr + ctMaxDopOpt)  '****** Sachin 3.01 ctMaxDopOpt added
 ' moCn.CreateTmpTable "#TmpFgBag", wStr
 
  adc("UoCoCdFr").CmpStr = "": adc("UoOmTcFr").CmpStr = "": adc("UoOmTcTo").CmpStr = ""
  adc("UoOmYyFr").CmpStr = "": adc("UoOmYyTo").CmpStr = "": adc("UoOmChrFr").CmpStr = ""
  adc("UoOmChrTo").CmpStr = "": adc("UoOmNoFr").CmpStr = "": adc("UoOmNoTo").CmpStr = ""
  adc("UoOdSrFr").CmpStr = "": adc("UoOdSrTo").CmpStr = ""
  adc("UoCmCdFr").CmpStr = "OmCmCd>=": adc("UoCmCdTo").CmpStr = "OmCmCd<=": adc("UoCmCdSel").CmpStr = "OmCmCd In"
  adc("UoDmCtgFr").CmpStr = "DmCtg>=": adc("UoDmCtgTo").CmpStr = "DmCtg<=": adc("UoDmCtgSel").CmpStr = "DmCtg In"
  adc("UoKtFr").CmpStr = "OdKt>=": adc("UoKtTo").CmpStr = "OdKt<=": adc("UoKtSel").CmpStr = "OdKt In"
  adc("UoDmCdFr").CmpStr = "OdDmCd>=": adc("UoDmCdTo").CmpStr = "OdDmCd<=": adc("UoDmCdSel").CmpStr = "OdDmCd In"
  
  '****** Sachin 3.02.0 - New Fields In Where Clause
  wAddlCnd = wAddlCnd + IIF(wAddlCnd <> "" And Trim(adc("UOCMCDFR")) <> "", " And ", "") + IIF(Trim(adc("UOCMCDFR")) <> "", " OdOmCmCd >= '" + Trim(adc("UOCMCDFR")) + "'", "")
  wAddlCnd = wAddlCnd + IIF(wAddlCnd <> "" And Trim(adc("UOCMCDTO")) <> "", " And ", "") + IIF(Trim(adc("UOCMCDTO")) <> "", " OdOmCmCd <= '" + Trim(adc("UOCMCDTO")) + "'", "")
  wAddlCnd = wAddlCnd + IIF(wAddlCnd <> "" And Trim(adc("UOCMCDSEL")) <> "", " And ", "") + IIF(Trim(adc("UOCMCDSEL")) <> "", " OdOmCmCd In (" + Trim(adc("UOCMCDSEL")) + ")", "")

  wrepcnd = adc.RepCond
  wCnd = IIF(wrepcnd <> "", " Where " + wrepcnd, "")
 ' wCnd1 = IIf(wrepcnd <> "", " And " + wrepcnd, "")
 '6.1
   wStr = "Set DateFormat " + DtFmtStr() + "  insert into #TmpFgBag " + _
          " select '1Opn' as Tag,Min(Dt),BCoCd,Min(Tc),Min(Yy),Min(Chr)," + _
          "Min(TNo),Min(Sr),BYy,BChr,BNo ,Min(BQty),Min(RefKey),Min(BPrtKey),Min(BOdTc),Min(BOdYy),Min(BOdChr), " + _
          "Min(BOdNo),Min(BOdSr) from  #TmpFgBag  as a " + _
          "Group By BCoCd,BYy,BChr,BNo " + _
          "Having  (not exists(select * from #TmpFgBag a1 join Param on PTyp= 'TC' and PmCd= Tc and PSCd= '' and PValue ='FB' and " + _
                               "refKey = (Select Min(refkey) from #TmpFgBag  a2 " + _
                                           "Where a.BCoCd = a2.BCoCd And a.Byy = a2.Byy And " + _
                                           "a.BChr = a2.BChr and a.BNo=a2.BNo and Tag='2Fg') " + _
          "where a.BCoCd=a1.BCoCd and a.BYy=a1.BYy and a.BChr=a1.BChr and a.BNo=a1.BNo and Tag='2Fg') ) "

    moCn.Execute (wStr + ctMaxDopOpt)  '****** Sachin 3.01 ctMaxDopOpt added

  'qTag, qTcTyp, qDt, BCoCd, qTc, qYy, qChr, qNo, qSr, Byy, BChr, BNo, BQty, RefKey,
  'BOdTc, BOdYy, BOdChr, BOdNo, BOdSr, DmCtg, OdDmCd, DmDesc, OdKt, OdDmCol, OmCmCd
  
  '****** Sachin 3.02 - Id fields in Joins
  wSqlStrg = "select " + grpflds + ",Tag as qTag, IsNull((select PValue from Param where PTyp= 'TC' and " + _
             "PmCd= Tc and PSCd= ''),'') as qTcTyp, Dt as qDt, BCoCd, Tc as qTc, Yy as qYy, Chr as qChr, " + _
             "TNo as qNo, Sr as qSr, Byy, BChr, BNo, BQty, RefKey, BPrtKey, BOdTc, BOdYy, BOdChr, BOdNo, BOdSr, " + _
             "DmCtg, OdDmCd, DmDesc, OdKt, OdDmCol, OmCmCd from #TmpFgBag as A " + _
             "join OrdDsg on OdCoCd=BCoCd and OdTc=BOdTc and OdYy=BOdYy  " + _
             "and OdChr=BOdChr and OdNo=BOdNo and OdSr=BOdSr " + IIF(gs_Partition = ctCurrPrtn, " and OdPrtKey=BPrtKey", "") + _
             " Join OrdMst on OmIdNo=OdOmIdNo and OmCoCd=OdCoCd and OmTc=OdTc and OmYy=OdYy and OmChr=OdChr and OmNo=OdNo " + IIF(gs_Partition = ctCurrPrtn, " and OmPrtKey=OdPrtKey", "") + _
             " Join DsgMst On DmIdNo=OdDmIdNo and DmTcTyp=OdDmTcTyp and DmCd=OdDmCd and DmSz='' " + IIF(gs_Partition = ctCurrPrtn, " and DmPrtKey=OdPrtKey", "") + wCnd + _
             IIF(wAddlCnd <> "", IIF(wCnd <> "", " And ", " Where ") + wAddlCnd + IIF(gs_Partition = ctCurrPrtn, " and OdPrtKey='" + ctCurrPrtn + "' ", ""), "") + _
             " Order by Tag, RefKey, BCocd, BYy, BChr, BNo "

'"select " + grpflds + ",'1Opn' as Tag,IsNull((select PValue from Param where PTyp= 'TC' and " + _
       "PmCd= Tc and PSCd= ''),'') TcTyp,Dt,BCoCd,Tc,Yy,Chr,TNo,Sr,Byy,BChr,BNo,BQty, " + _
       "RefKey,BOdTc,BOdYy,BOdChr,BOdNo,BOdSr,DmCtg,OdDmCd,DmDesc,OdKt,OdDmCol,OmCmCd from #TmpFgBag  as a " + _
       "join OrdDsg on OdCoCd=BCoCd and OdTc=BOdTc and OdYy=BOdYy  " + _
       "and OdChr=BOdChr and OdNo=BOdNo and OdSr=BOdSr " + _
       "join OrdMst on OmCoCd=OdCoCd and OmTc=OdTc and OmYy=OdYy and OmChr=OdChr and OmNo=OdNo " + _
       "Join DsgMst On DmTcTyp='DM' and DmCd=OdDmCd and DmSz='' " + _
       "where (not exists(select * from #TmpFgBag a1 join Param on PTyp= 'TC' and PmCd= Tc and PSCd= '' and PValue ='FB' and " + _
       "refKey = (Select Min(refkey) from #TmpFgBag  a2 " + _
       "Where a.BCoCd = a2.BCoCd And a.Byy = a2.Byy And a.BChr = a2.BChr " + _
       "and a.BNo=a2.BNo and Tag='2Fg') " + _
       "where a.BCoCd=a1.BCoCd and a.BYy=a1.BYy and a.BChr=a1.BChr and a.BNo=a1.BNo and Tag='2Fg') ) " + wCnd1 + _
       "Order by  Tag,RefKey,BCocd,BYy,BChr,BNo "
 ' wSqlStrg = wSqlStrg1 + " Union All  " + wSqlStrg2
     
  adc.RepRecSource = wSqlStrg + ctMaxDopOpt  '****** Sachin 3.01 ctMaxDopOpt added
  If adc("UoYN") = "N" Then
    Rep.BagDet.Suppress = True
    Rep.Bag.Suppress = True
    Rep.BagHd.Suppress = True
    Rep.BagFt.Suppress = True
  ElseIf adc("UoYN") = "Y" And wDetPos <> 0 Then
    Rep.BagDet.Suppress = False
    Rep.Bag.Suppress = False
    Rep.BagHd.Suppress = False
    Rep.BagFt.Suppress = False
  End If
 '*******Detail Sort is Bag Wise/Transaction Wise******
   GRP_REP.SetFormula Rep, "Grp5", IIF(adc("UoPMCdFr") = "BAG", "{@wBagNo}", "{@wTag}")
   GRP_REP.SetFormula Rep, "g5_Desc1", IIF(adc("UoPMCdFr") = "BAG", "{@wBagNo}", "{@wTransaction}")
   GRP_REP.SetFormula Rep, "g5_Desc2", IIF(adc("UoPMCdFr") = "BAG", "{@wTransaction}", "{@wBagNo}")
   Rep.hDesc1.SetText (IIF(adc("UoPMCdFr") = "BAG", "Bag No", "Transaction Type"))
   Rep.hDesc2.SetText (IIF(adc("UoPMCdFr") = "BAG", "Transaction Type", "Bag No"))
'  If adc("UoYN1") = "" Then
'    Rep.BagHd.Suppress = True
'    Rep.Bag.Suppress = True
'    Rep.BagDet.Suppress = True
'    Rep.BagFt.Suppress = True
'  End If
      
 Rep.TxtFrDt.SetText adc("UoDocDtFr")
 Rep.TxtToDt.SetText moCn.SrvrDate
 CRV_REP.DisplayGroupTree = False
 
 Call DispReport(adc) '****** Sachin 3.02 - Display Report ******
 
Exit Sub
RepErr:
  Call ChkSelScope(Err.Number, Err.Description)
End Sub

Private Sub ApportnTbl()
  Dim wStr As String
  '*** Creating Temporary Table #TmpFgBag which is the main Apportioning Sql ******
  '578-34 RefKey VarChar(9) changed to VarChar(10)
  wStr = "Create Table  #TmpFgBag " + _
          "(Tag  VarChar(4)," + _
          "Dt smallDateTime, " + _
          "BCocd  VarChar(3)," + _
          "Tc VarChar(3)," + _
          "Yy VarChar(2)," + _
          "Chr VarChar(3)," + _
          "TNo integer," + _
          "Sr Integer," + _
          "BYy varchar(2)," + _
          "BChr  varchar(3)," + _
          "BNo Integer," + _
          "BQty Float, " + _
          "RefKey VarChar(10)," + _
          "BPrtKey VarChar(1)," + _
          "BOdTc VarChar(3)," + _
          "BOdYy VarChar(2)," + _
          "BOdChr VarChar(3)," + _
          "BOdNo integer," + _
          "BOdSr Integer" + _
          ")"
 
   moCn.CreateTmpTable "#TmpFgBag", wStr
  
End Sub

' ***** Manali 3.5.0 - 21/11/08 - Check Boxes Added
Private Sub ChkBoxArr_Click(Index As Integer)
With ChkBoxArr(Index)
  Select Case Index
  Case Is = 0     ' Tag = "ShowLine"
    If .Value = Checked Then adc("UoYN") = "Y" Else adc("UoYN") = "N"
  End Select
End With
End Sub
' ***** Manali 3.5.0 - 21/11/08 - Check Boxes Added

' ***** Manali 3.5.0 - 21/11/08 - Check Boxes Added
Private Sub ChkBoxArr_GotFocus(Index As Integer)
  DispMsg ChkBoxArr(Index).ToolTipText, etInfo
End Sub
' ***** Manali 3.5.0 - 21/11/08 - Check Boxes Added

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
  
  Set adc.FirNKeyCtl = adc("UoDmCtgFr")
  Call InitProp(Me)
  
  '*** Parameters Used For Setting Properties Of UsrOpt Fields ***
  gs_LocTyp = "'P'"
  gs_DmTcTyp = "DM"
  gs_OmTcTyp = "SO"
  gs_CmCtg = "C"
  gs_PTyp = "SRTFGBAG"
  '*** Parameters Used For Setting Properties Of UsrOpt Fields ***
  
  Call SetGroupSort
End Sub

Private Sub ATXT_FldChange(Index As Integer, ByVal IdName As String, Cancel As Boolean, ErrMsg As String, ByVal pv_NewValue As Variant)
  Call SetProp(Me, IdName, Valid, Cancel, ErrMsg)
End Sub
Private Sub ATXT_FldWhen(Index As Integer, ByVal IdName As String, Cancel As Boolean, ErrMsg As String)
 gb_CoCdFor = True         '*** (Jen 2.14 Next Patch 1)
 
 Select Case IdName
  Case Is = UCase("UOPMCDFR")
    If adc("UoYN") = "N" Then
      Cancel = True: ErrMsg = "Cannot Edit This Value": Exit Sub
    End If
 End Select
   SetProp Me, IdName, When
End Sub

Private Sub atxt_ReCalculate(Index As Integer, ByVal IdName As String, ByVal OldVal As Variant, ByVal RowNum As Integer, ByVal SourceName As String, ByVal SourceColName As String)
Select Case IdName
Case Is = UCase("UOPMCDFR")
  If adc("UoYN") = "N" Then
    adc("UOPMCDFR") = ""
  End If
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
Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
  If UnloadMode <> vbFormCode Then Cancel = True: Exit Sub
    
    '*** (09/08/05)
    Set moCn = Nothing   '*** (Bef speed) bef 24/11/06
    Set Rep = Nothing
    '*** (09/08/05)
End Sub
