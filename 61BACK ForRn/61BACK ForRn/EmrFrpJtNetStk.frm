VERSION 5.00
Object = "{8767A745-088E-4CA6-8594-073D6D2DE57A}#9.2#0"; "crviewer9.dll"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "tabctl32.ocx"
Object = "{AB2CC32A-E652-42A2-B341-26ECE02F6386}#2.1#0"; "mwfctlforspeed9.ocx"
Begin VB.Form EmrFrpJtNetStk 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Net Stock Requirement (Jwlry)"
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
      Left            =   11295
      TabIndex        =   44
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
      TabIndex        =   43
      TabStop         =   0   'False
      Top             =   9630
      Width           =   6330
      _ExtentX        =   11165
      _ExtentY        =   1005
   End
   Begin VB.Frame fra_keys 
      BackColor       =   &H00C0C0C0&
      BorderStyle     =   0  'None
      Height          =   525
      Left            =   5295
      TabIndex        =   53
      Top             =   9615
      Width           =   3555
      Begin MwfCtl.MWCTL_MED aktxt 
         Height          =   345
         Index           =   0
         Left            =   210
         TabIndex        =   54
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
         TabIndex        =   55
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
   Begin MwfCtl.MWCTL_ADC adc 
      Height          =   555
      Left            =   60
      TabIndex        =   45
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
   Begin VB.Frame FRA_REPOBJ 
      BorderStyle     =   0  'None
      Height          =   450
      Left            =   6510
      TabIndex        =   52
      Top             =   9585
      Visible         =   0   'False
      Width           =   735
      Begin CRVIEWER9LibCtl.CRViewer9 CRV_REP 
         Height          =   495
         Left            =   120
         TabIndex        =   49
         Top             =   30
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
      Height          =   9660
      Left            =   0
      TabIndex        =   48
      Top             =   60
      Width           =   15255
      Begin TabDlg.SSTab TAB_REP 
         Height          =   9465
         Left            =   15
         TabIndex        =   47
         Top             =   60
         Width           =   15195
         _ExtentX        =   26802
         _ExtentY        =   16695
         _Version        =   393216
         Tabs            =   2
         TabsPerRow      =   2
         TabHeight       =   520
         TabCaption(0)   =   "&Data Range"
         TabPicture(0)   =   "EmrFrpJtNetStk.frx":0000
         Tab(0).ControlEnabled=   -1  'True
         Tab(0).Control(0)=   "fra_tabrep(0)"
         Tab(0).Control(0).Enabled=   0   'False
         Tab(0).ControlCount=   1
         TabCaption(1)   =   "&Group Sort"
         TabPicture(1)   =   "EmrFrpJtNetStk.frx":001C
         Tab(1).ControlEnabled=   0   'False
         Tab(1).Control(0)=   "fra_tabrep(1)"
         Tab(1).ControlCount=   1
         Begin VB.Frame fra_tabrep 
            BorderStyle     =   0  'None
            Enabled         =   0   'False
            Height          =   9070
            Index           =   1
            Left            =   -74940
            TabIndex        =   51
            Top             =   360
            Width           =   15105
            Begin MwfCtl.MWCTL_GRP GRP_REP 
               Height          =   360
               Left            =   240
               TabIndex        =   46
               Top             =   360
               Width           =   9855
               _ExtentX        =   17383
               _ExtentY        =   635
            End
         End
         Begin VB.Frame fra_tabrep 
            BorderStyle     =   0  'None
            Height          =   9070
            Index           =   0
            Left            =   60
            TabIndex        =   50
            Top             =   360
            Width           =   15105
            Begin VB.CheckBox ChkBoxArr 
               Alignment       =   1  'Right Justify
               Caption         =   "Show All"
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
               Tag             =   "ShowDet"
               ToolTipText     =   "Check To Show All "
               Top             =   4590
               Width           =   1680
            End
            Begin VB.CheckBox ChkBoxArr 
               Alignment       =   1  'Right Justify
               Caption         =   "Show Detail    "
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
               TabIndex        =   40
               Tag             =   "ShowDet"
               ToolTipText     =   "Check To Show Location And order Details"
               Top             =   4305
               Width           =   1680
            End
            Begin MwfCtl.MWCTL_MED atxt 
               Height          =   285
               Index           =   28
               Left            =   4500
               TabIndex        =   12
               ToolTipText     =   "Enter To Design Category"
               Top             =   1470
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
               Left            =   7440
               TabIndex        =   13
               ToolTipText     =   "Enter Design Category Selection"
               Top             =   1470
               Width           =   7590
               _ExtentX        =   13388
               _ExtentY        =   503
               MaxLength       =   255
               DataType        =   4
               DataField       =   "UODMCTGSEL"
               IdName          =   "UODMCTGSEL"
            End
            Begin MwfCtl.MWCTL_MED atxt 
               Height          =   285
               Index           =   30
               Left            =   1485
               TabIndex        =   11
               ToolTipText     =   "Enter From Design Category"
               Top             =   1470
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
               Index           =   31
               Left            =   4500
               TabIndex        =   15
               ToolTipText     =   "Enter To Design Code"
               Top             =   1755
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
               Left            =   7440
               TabIndex        =   16
               ToolTipText     =   "Enter Design Code Selection"
               Top             =   1755
               Width           =   7590
               _ExtentX        =   13388
               _ExtentY        =   503
               MaxLength       =   255
               DataField       =   "UODMCDSEL"
               IdName          =   "UODMCDSEL"
            End
            Begin MwfCtl.MWCTL_MED atxt 
               Height          =   285
               Index           =   33
               Left            =   1485
               TabIndex        =   14
               ToolTipText     =   "Enter From Design Code"
               Top             =   1755
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
               Left            =   4500
               TabIndex        =   22
               ToolTipText     =   "Enter To Customer"
               Top             =   2610
               Width           =   1305
               _ExtentX        =   2302
               _ExtentY        =   503
               MaxLength       =   8
               DataField       =   "UOCMCDTO"
               IdName          =   "UOCMCDTO"
            End
            Begin MwfCtl.MWCTL_MED atxt 
               Height          =   285
               Index           =   41
               Left            =   7440
               TabIndex        =   23
               ToolTipText     =   "Enter Customer Selection"
               Top             =   2610
               Width           =   7590
               _ExtentX        =   13388
               _ExtentY        =   503
               MaxLength       =   255
               DataField       =   "UOCMCDSEL"
               IdName          =   "UOCMCDSEL"
            End
            Begin MwfCtl.MWCTL_MED atxt 
               Height          =   285
               Index           =   42
               Left            =   1485
               TabIndex        =   21
               ToolTipText     =   "Enter From Customer"
               Top             =   2610
               Width           =   1305
               _ExtentX        =   2302
               _ExtentY        =   503
               MaxLength       =   8
               DataField       =   "UOCMCDFR"
               IdName          =   "UOCMCDFR"
            End
            Begin MwfCtl.MWCTL_MED atxt 
               Height          =   285
               Index           =   43
               Left            =   4500
               TabIndex        =   25
               ToolTipText     =   "Enter To Purchase Order "
               Top             =   2895
               Width           =   2625
               _ExtentX        =   4630
               _ExtentY        =   503
               MaxLength       =   20
               DataField       =   "UOPONOTO"
               IdName          =   "UOPONOTO"
            End
            Begin MwfCtl.MWCTL_MED atxt 
               Height          =   285
               Index           =   44
               Left            =   7440
               TabIndex        =   26
               ToolTipText     =   "Enter Purchase Order Selection"
               Top             =   2895
               Width           =   7590
               _ExtentX        =   13388
               _ExtentY        =   503
               MaxLength       =   255
               DataField       =   "UOPONOSEL"
               IdName          =   "UOPONOSEL"
            End
            Begin MwfCtl.MWCTL_MED atxt 
               Height          =   285
               Index           =   45
               Left            =   1485
               TabIndex        =   24
               ToolTipText     =   "Enter From Purchase Order "
               Top             =   2895
               Width           =   2625
               _ExtentX        =   4630
               _ExtentY        =   503
               MaxLength       =   20
               DataField       =   "UOPONOFR"
               IdName          =   "UOPONOFR"
            End
            Begin MwfCtl.MWCTL_MED atxt 
               Height          =   285
               Index           =   46
               Left            =   4500
               TabIndex        =   33
               ToolTipText     =   "Enter To Order Date"
               Top             =   3465
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
               Index           =   48
               Left            =   1485
               TabIndex        =   32
               ToolTipText     =   "Enter From Order Date"
               Top             =   3465
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
               Index           =   5
               Left            =   5265
               TabIndex        =   5
               ToolTipText     =   "Enter To Order Year"
               Top             =   900
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
               Left            =   2280
               TabIndex        =   2
               ToolTipText     =   "Enter From Order Year"
               Top             =   900
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
               Left            =   1485
               TabIndex        =   1
               ToolTipText     =   "Enter From Order Tc"
               Top             =   900
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
               Left            =   2685
               TabIndex        =   3
               ToolTipText     =   "Enter From Order Character"
               Top             =   900
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
               Left            =   4500
               TabIndex        =   4
               ToolTipText     =   "Enter To Order Tc"
               Top             =   900
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
               Left            =   5685
               TabIndex        =   6
               ToolTipText     =   "Enter To Order Character"
               Top             =   900
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
               Index           =   26
               Left            =   5385
               TabIndex        =   10
               ToolTipText     =   "Enter To Order Serial"
               Top             =   1185
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
               Index           =   52
               Left            =   2385
               TabIndex        =   8
               ToolTipText     =   "Enter From Order Serial"
               Top             =   1185
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
               Index           =   53
               Left            =   1485
               TabIndex        =   7
               ToolTipText     =   "Enter From Order Number"
               Top             =   1185
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
               Left            =   4500
               TabIndex        =   9
               ToolTipText     =   "Enter To Order Number"
               Top             =   1185
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
               Index           =   47
               Left            =   1485
               TabIndex        =   0
               ToolTipText     =   "Enter Company Code"
               Top             =   645
               Width           =   705
               _ExtentX        =   1244
               _ExtentY        =   503
               MaxLength       =   3
               DataType        =   4
               DataField       =   "UOCOCDFR"
               IdName          =   "UOCOCDFR"
            End
            Begin MwfCtl.MWCTL_MED atxt 
               Height          =   285
               Index           =   9
               Left            =   4500
               TabIndex        =   37
               ToolTipText     =   "Enter To FrBag Location"
               Top             =   4020
               Width           =   885
               _ExtentX        =   1561
               _ExtentY        =   503
               MaxLength       =   5
               DataField       =   "UOISSLOCTO"
               IdName          =   "UOISSLOCTO"
            End
            Begin MwfCtl.MWCTL_MED atxt 
               Height          =   285
               Index           =   11
               Left            =   7440
               TabIndex        =   38
               ToolTipText     =   "Enter FrBag Location Selection"
               Top             =   4020
               Width           =   7590
               _ExtentX        =   13388
               _ExtentY        =   503
               MaxLength       =   255
               DataField       =   "UOISSLOCSEL"
               IdName          =   "UOISSLOCSEL"
            End
            Begin MwfCtl.MWCTL_MED atxt 
               Height          =   285
               Index           =   12
               Left            =   1485
               TabIndex        =   36
               ToolTipText     =   "Enter From FrBag Location"
               Top             =   4020
               Width           =   885
               _ExtentX        =   1561
               _ExtentY        =   503
               MaxLength       =   5
               DataField       =   "UOISSLOCFR"
               IdName          =   "UOISSLOCFR"
            End
            Begin MwfCtl.MWCTL_MED atxt 
               Height          =   285
               Index           =   24
               Left            =   1905
               TabIndex        =   39
               ToolTipText     =   "Show Location And Order Detail ? (Enter 'Y'/ 'N')"
               Top             =   4305
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
               Index           =   16
               Left            =   3075
               TabIndex        =   41
               ToolTipText     =   "Show All Or Order Specific RM (Enter 'Y' For All Or 'N' For Order Specific RM)"
               Top             =   4560
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
               Index           =   4
               Left            =   2770
               TabIndex        =   28
               ToolTipText     =   "Enter From Sub Purchase Order"
               Top             =   3180
               Width           =   1665
               _ExtentX        =   2937
               _ExtentY        =   503
               MaxLength       =   20
               DataField       =   "UOODPONOFR"
               IdName          =   "UOODPONOFR"
            End
            Begin MwfCtl.MWCTL_MED atxt 
               Height          =   285
               Index           =   6
               Left            =   5790
               TabIndex        =   30
               ToolTipText     =   "Enter To Sub Purchase Order "
               Top             =   3180
               Width           =   1545
               _ExtentX        =   2725
               _ExtentY        =   503
               MaxLength       =   20
               DataField       =   "UOODPONOTO"
               IdName          =   "UOODPONOTO"
            End
            Begin MwfCtl.MWCTL_MED atxt 
               Height          =   285
               Index           =   14
               Left            =   7440
               TabIndex        =   31
               ToolTipText     =   "Enter Sub Purchase Order Selection"
               Top             =   3180
               Visible         =   0   'False
               Width           =   7590
               _ExtentX        =   13388
               _ExtentY        =   503
               MaxLength       =   255
               DataField       =   "UOODPONOSEL"
               IdName          =   "UOODPONOSEL"
            End
            Begin MwfCtl.MWCTL_MED atxt 
               Height          =   285
               Index           =   34
               Left            =   4500
               TabIndex        =   20
               ToolTipText     =   "Enter To Design Size"
               Top             =   2325
               Width           =   585
               _ExtentX        =   1032
               _ExtentY        =   503
               MaxLength       =   5
               DataField       =   "UODMSZTO"
               IdName          =   "UODMSZTO"
            End
            Begin MwfCtl.MWCTL_MED atxt 
               Height          =   285
               Index           =   36
               Left            =   1485
               TabIndex        =   19
               ToolTipText     =   "Enter From Design Size"
               Top             =   2325
               Width           =   585
               _ExtentX        =   1032
               _ExtentY        =   503
               MaxLength       =   5
               DataField       =   "UODMSZFR"
               IdName          =   "UODMSZFR"
            End
            Begin MwfCtl.MWCTL_MED atxt 
               Height          =   285
               Index           =   37
               Left            =   4500
               TabIndex        =   18
               ToolTipText     =   "Enter To Suffix"
               Top             =   2040
               Width           =   1605
               _ExtentX        =   2831
               _ExtentY        =   503
               MaxLength       =   10
               DataField       =   "UOODSFXTO"
               IdName          =   "UOODSFXTO"
            End
            Begin MwfCtl.MWCTL_MED atxt 
               Height          =   285
               Index           =   39
               Left            =   1485
               TabIndex        =   17
               ToolTipText     =   "Enter From Suffix"
               Top             =   2040
               Width           =   1605
               _ExtentX        =   2831
               _ExtentY        =   503
               MaxLength       =   10
               DataField       =   "UOODSFXFR"
               IdName          =   "UOODSFXFR"
            End
            Begin MwfCtl.MWCTL_MED atxt 
               Height          =   285
               Index           =   0
               Left            =   4500
               TabIndex        =   35
               ToolTipText     =   "Enter To Export Delivery Date"
               Top             =   3750
               Width           =   885
               _ExtentX        =   1561
               _ExtentY        =   503
               MaxLength       =   8
               DataType        =   3
               DataField       =   "UODELDTTO"
               IdName          =   "UODELDTTO"
            End
            Begin MwfCtl.MWCTL_MED atxt 
               Height          =   285
               Index           =   1
               Left            =   1485
               TabIndex        =   34
               ToolTipText     =   "Enter From Export Delievery Date"
               Top             =   3750
               Width           =   885
               _ExtentX        =   1561
               _ExtentY        =   503
               MaxLength       =   8
               DataType        =   3
               DataField       =   "UODELDTFR"
               IdName          =   "UODELDTFR"
            End
            Begin MwfCtl.MWCTL_MED atxt 
               Height          =   285
               Index           =   2
               Left            =   4500
               TabIndex        =   79
               ToolTipText     =   "Enter Company Code"
               Top             =   645
               Width           =   705
               _ExtentX        =   1244
               _ExtentY        =   503
               MaxLength       =   3
               DataType        =   4
               DataField       =   "UOCOCDTO"
               IdName          =   "UOCOCDTO"
            End
            Begin MwfCtl.MWCTL_MED atxt 
               Height          =   285
               Index           =   50
               Left            =   7440
               TabIndex        =   80
               ToolTipText     =   "Enter Design Code Selection"
               Top             =   645
               Width           =   7590
               _ExtentX        =   13388
               _ExtentY        =   503
               MaxLength       =   255
               DataField       =   "UOCOCDSEL"
               IdName          =   "UOCOCDSEL"
            End
            Begin MwfCtl.MWCTL_MED atxt 
               Height          =   285
               Index           =   3
               Left            =   1485
               TabIndex        =   27
               ToolTipText     =   "Enter From Sub Customer"
               Top             =   3180
               Width           =   1215
               _ExtentX        =   2143
               _ExtentY        =   503
               MaxLength       =   8
               DataType        =   4
               DataField       =   "UOODSUBCUSTFR"
               IdName          =   "UOODSUBCUSTFR"
            End
            Begin MwfCtl.MWCTL_MED atxt 
               Height          =   285
               Index           =   35
               Left            =   4500
               TabIndex        =   29
               ToolTipText     =   "Enter To Sub Customer"
               Top             =   3180
               Width           =   1215
               _ExtentX        =   2143
               _ExtentY        =   503
               MaxLength       =   8
               DataType        =   4
               DataField       =   "UOODSUBCUSTTO"
               IdName          =   "UOODSUBCUSTTO"
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
               Left            =   5720
               TabIndex        =   82
               Top             =   3180
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
               Index           =   7
               Left            =   2700
               TabIndex        =   81
               Top             =   3180
               Width           =   105
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
               Index           =   0
               Left            =   30
               TabIndex        =   78
               Top             =   3750
               Width           =   1485
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
               Left            =   30
               TabIndex        =   77
               Top             =   2040
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
               Left            =   30
               TabIndex        =   76
               Top             =   2325
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
               Index           =   5
               Left            =   30
               TabIndex        =   75
               Top             =   3180
               Width           =   1425
            End
            Begin VB.Label Label3 
               BackStyle       =   0  'Transparent
               Caption         =   "Show Detail"
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
               TabIndex        =   74
               Top             =   4260
               Visible         =   0   'False
               Width           =   1485
            End
            Begin VB.Label ALBL 
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
               Index           =   6
               Left            =   30
               TabIndex        =   73
               Top             =   4005
               Width           =   1485
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
               Left            =   30
               TabIndex        =   72
               Top             =   900
               Width           =   1485
            End
            Begin VB.Label ALBL 
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
               Left            =   30
               TabIndex        =   71
               Top             =   1185
               Width           =   1485
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
               Left            =   5625
               TabIndex        =   70
               Top             =   900
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
               Left            =   2625
               TabIndex        =   69
               Top             =   900
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
               Left            =   5205
               TabIndex        =   68
               Top             =   900
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
               Left            =   2205
               TabIndex        =   67
               Top             =   900
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
               Left            =   5325
               TabIndex        =   66
               Top             =   1185
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
               Left            =   2325
               TabIndex        =   65
               Top             =   1185
               Width           =   105
            End
            Begin VB.Label ALBL 
               BackStyle       =   0  'Transparent
               Caption         =   "Ord Dt"
               BeginProperty Font 
                  Name            =   "Arial"
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
               Left            =   30
               TabIndex        =   64
               Top             =   3465
               Width           =   1485
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
               Left            =   30
               TabIndex        =   63
               Top             =   2895
               Width           =   1485
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
               Left            =   30
               TabIndex        =   62
               Top             =   2610
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
               Left            =   30
               TabIndex        =   61
               Top             =   1755
               Width           =   1485
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
               Left            =   30
               TabIndex        =   60
               Top             =   1470
               Width           =   1485
            End
            Begin VB.Label ALBL 
               BackStyle       =   0  'Transparent
               Caption         =   "Co Cd"
               BeginProperty Font 
                  Name            =   "Arial"
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
               Left            =   30
               TabIndex        =   59
               ToolTipText     =   "Location"
               Top             =   645
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
               Left            =   6840
               TabIndex        =   58
               Top             =   135
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
               Left            =   4140
               TabIndex        =   57
               Top             =   135
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
               Left            =   1485
               TabIndex        =   56
               Top             =   135
               Width           =   1335
            End
         End
      End
   End
End
Attribute VB_Name = "EmrFrpJtNetStk"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim Rep As New EmrRepJtNetStk
Dim moCn As MwfLib.MDOConnection   '*** (Bef speed) bef 24/11/06
Dim ms_MultiCoMod As Boolean

Private Sub SetGroupSort()
  GRP_REP.Add "Design Ctg", "max(DmCtg)", "", "", "", "Sales Ctg", "", "", "(Select PDesc From Param Where PTyp= 'DMCTG' and PmCd= Max(DmCtg)) "
  GRP_REP.Add "Sales Ctg", "Max(DmSalCtg)", "", "", "", "Design Ctg", "", "", "(Select PDesc From Param Where PTyp= 'SALCTG' and Pmcd=max(DmCtg) and PSCd= Max(DmSalCtg)) "
  
  GRP_REP.Add "Design Code", "So.OdDmCd", "", "OdDmCd", "hOdDmCd", "Design Ctg,Sales Ctg", "", "", ""
  GRP_REP.Add "Design Size", "So.OdDmSz", "", "OdDmSz", "hOdDmSz", "", "", "", "(Select PDesc From Param Where PTyp= 'DMSZ' and PmCd= So.OdDmSz)"
  GRP_REP.Add "Suffix", "So.OdSfx", "", "OdSfx", "hOdSfx", "", "", "", ""

 If ms_MultiCoMod = True Then _
   GRP_REP.Add "Company Code", "So.OdCoCd", "", "OdCoCd", "", "", "", "", ""
 
  
'  GRP_REP.Add "Design Ctg", Array("max(DmCtg)", "DmCtg"), "", "", "", "Sales Ctg", "", "", Array("(Select PDesc From Param Where PTyp= 'DMCTG' and PmCd= Max(DmCtg)) ", "(Select PDesc From Param Where PTyp= 'DMCTG' and PmCd= DmCtg) ")
'  GRP_REP.Add "Sales Ctg", Array("Max(DmSalCtg)", "DmSalCtg"), "", "", "", "Design Ctg", "", "", Array("(Select PDesc From Param Where PTyp= 'SALCTG' and Pmcd=max(DmCtg) and PSCd= Max(DmSalCtg)) ", "(Select PDesc From Param Where PTyp= 'SALCTG' and Pmcd=DmCtg and PSCd= DmSalCtg) ")
'
'  GRP_REP.Add "Design Code", Array("DDmCd", "So.OdDmCd"), "", "OdDmCd", "hOdDmCd", "Design Ctg,Sales Ctg", "", "", ""
'  GRP_REP.Add "Design Size", Array("DDmSz", "So.OdDmSz"), "", "OdDmSz", "hOdDmSz", "", "", "", Array("(Select PDesc From Param Where PTyp= 'DMSZ' and PmCd= DDmSz)", "(Select PDesc From Param Where PTyp= 'DMSZ' and PmCd= SO.OdDmSz)")
'  GRP_REP.Add "Suffix", Array("DOdSfx", "So.OdSfx"), "", "OdSfx", "hOdSfx", "", "", "", ""
  
  'GRP_REP.Add "Rm StkRt", Array("str(SRmStkRt,9,2)", "str(TdRmStkRt,9,2)", "str(FrRmStkRt,9,2)", "str(IrStkRt,9,2)"), "", "qStkRt", "hqStkRt", "(RmCd) StkRt", "", "", ""

End Sub
Private Sub adc_setreprecsource()
On Err GoTo Errhnd
  
  '*** Report Sql ***
  '  Select space(30) as Grp1, space(30) as Grp2, space(30) as Grp3,
  '  space(30) as DGrp1, space(30) as DGrp2, space(30) as DGrp3,
  '  space(30) as G1Desc, space(30) as G2Desc, space(30) as G3Desc,
  '  DLoc, DDmCd as OdDmCd, DOdSfx as OdSfx, DDmSz as OdDmSz, DDmId as OdDmId,
  '  DDmDrQty as qAvlDmQty, OdOrdQty as qBalOrdQty,
  '  OdCoCd , OdTc, OdYy, OdChr, OdNo, OdSr
  '  From DmYyMm, OrdDsg, OrdMst
  '  Where 1 = 2
  '*** Report Sql ***

  Dim grpflds0 As String, grpflds1 As String
  Dim wSetTmpTbl As String
  Dim wSetTmpSo As String, wSetTmpPo As String, wSetTmpStk As String
  Dim wSoSql As String, wPoSql As String, wStkSql As String
  Dim wDmYyMmSql As String, wOrdSql As String
  Dim ws_RepCnd As String, ws_Cnd As String, ws_SqlStrg As String
  Dim ws_Cnd1 As String, ws_Cnd2 As String
  Dim wAddlCnd As String
  Dim ws_All As String, ws_OrdQTY As String, ws_Order As String

  Set Rep = Nothing
  Set adc.RepSource = Rep
  Call DispCoNm
  Call GRP_REP.Gen3LRep
  
  grpflds0 = GRP_REP.GrpFldLst(0)
  grpflds1 = GRP_REP.GrpFldLst(1)
    
  '*** To check if Detail is selected ***
  Dim ws_Grp(6) As String, wi_i As Integer, wi_DetPos As Integer
  ws_Grp(1) = GRP_REP.Value(gltGroup1)
  ws_Grp(2) = GRP_REP.Value(gltGroup2)
  ws_Grp(3) = GRP_REP.Value(gltGroup3)
  wi_DetPos = 0
  For wi_i = 1 To 6
    If ws_Grp(wi_i) = "(Detail)" Then
      wi_DetPos = wi_i: Exit For
    End If
  Next wi_i
    
  wSetTmpSo = "Select OdCoCd,OdTc,OdYy,OdChr ,OdNo,OdSr,OdDmCd,OdSfx,OdDmSz,OdOmCmCd as OmCmCd,OdOmDt as OmDt, 0 as OdIdNo,DmCtg,DmSalCtg, " + _
              " convert(decimal(12, 1), (OdOrdQty - OdExpQty)) As qOrdQty " + _
              " Into #TmpSo " + _
              "From OrdDsg, DsgMst Where 1=2 "

  moCn.CreateTmpTable "#TmpSo", wSetTmpSo
  
  wSetTmpPo = "Select OdCoCd,OdTc,OdYy,OdChr,OdNo,OdSr,OdDmCd,OdSfx,OdDmSz,OdOmCmCd as OmCmCd,OdOmDt as OmDt, 0 as OdIdNo, " + _
              "convert(decimal(12, 1), (OdOrdQty-OdExpQty)) as qPoQty,OdRefIdKey " + _
              "Into #TmpPo From OrdDsg Where 1=2 "

  moCn.CreateTmpTable "#TmpPo", wSetTmpPo

  wSetTmpStk = "Select DCoCd,DLoc,DDmCd,DOdSfx,DDmSz, convert(decimal(12, 1), DDmDrQty) as qStkQty " + _
                "Into #TmpStk " + _
                "From DmYyMm Where 1=2 "
                
  moCn.CreateTmpTable "#TmpStk", wSetTmpStk
  
  
  wAddlCnd = wAddlCnd + IIF(wAddlCnd <> "" And Trim(adc("UOCMCDFR")) <> "", " And ", "") + IIF(Trim(adc("UOCMCDFR")) <> "", " OdOmCmCd >= '" + Trim(adc("UOCMCDFR")) + "'", "")
  wAddlCnd = wAddlCnd + IIF(wAddlCnd <> "" And Trim(adc("UOCMCDTO")) <> "", " And ", "") + IIF(Trim(adc("UOCMCDTO")) <> "", " OdOmCmCd <= '" + Trim(adc("UOCMCDTO")) + "'", "")
  wAddlCnd = wAddlCnd + IIF(wAddlCnd <> "" And Trim(adc("UOCMCDSEL")) <> "", " And ", "") + IIF(Trim(adc("UOCMCDSEL")) <> "", " OdOmCmCd In (" + Trim(adc("UOCMCDSEL")) + ")", "")
  
  wAddlCnd = wAddlCnd + IIF(wAddlCnd <> "" And Trim(adc("UODTFR")) <> "01/01/80", " And ", "") + IIF(Trim(adc("UODTFR")) <> "01/01/80", " OdOmDt >= '" + Trim(adc("UODTFR")) + "'", "")
  wAddlCnd = wAddlCnd + IIF(wAddlCnd <> "" And Trim(adc("UODTTO")) <> "01/01/80", " And ", "") + IIF(Trim(adc("UODTTO")) <> "01/01/80", " OdOmDt <= '" + Trim(adc("UODTTO")) + "'", "")
  
  ws_Cnd1 = GetCnd("R")
  ws_Cnd2 = GetCnd("D")
  
  moCn.Execute "Select * from #TmpStk"
  wSoSql = "Insert Into #TmpSo " + _
          "Select OdCoCd,OdTc,OdYy,OdChr ,OdNo,OdSr,OdDmCd,OdSfx,OdDmSz,OdOmCmCd as OmCmCd,OdOmDt as OmDt,OdIdNo,DmCtg,DmSalCtg, " + _
          "convert(decimal(12, 1), (OdOrdQty - OdExpQty)) As qOrdQty " + _
          "From OrdDsg " + _
          "Join OrdMst on OmIdNo=OdOmIdNo and OmCoCd= OdCoCd and OmTc= OdTc  and OmYy= OdYy and OmChr= OdChr and OmNo= OdNo  " + IIF(gs_Partition = ctCurrPrtn, " and OmPrtKey=OdPrtKey ", "") + _
          "Join DsgMst on DmIdNo=OdDmIdNo and DmTcTyp= 'DM' and DmCd= OdDmCd and DmSz= ''  " + IIF(gs_Partition = ctCurrPrtn, " and OdPrtKey=DmPrtKey ", "") + _
          "where  " + IIF(gs_Partition = ctCurrPrtn, " OdPrtKey='" + ctCurrPrtn + "' and ", "") + _
          " convert(decimal(12, 1), (OdOrdQty-OdExpQty))> 0 " + ws_Cnd2 + IIF(wAddlCnd <> "", " And " + wAddlCnd, "")
  moCn.Execute wSoSql

  wPoSql = "Insert Into #TmpPo " + _
            "Select OdCoCd,OdTc,OdYy,OdChr,OdNo,OdSr,OdDmCd,OdSfx,OdDmSz,OdOmCmCd as OmCmCd,OdOmDt as OmDt,OdIdNo, " + _
            "convert(decimal(12, 1), (OdOrdQty-OdExpQty)) as qPoQty,OdRefIdKey " + _
            "From OrdDsg Po where " + IIF(gs_Partition = ctCurrPrtn, " OdPrtKey='" + ctCurrPrtn + "' and ", "") + _
            "OdTc='JPO' and exists(Select 'x' from #TmpSo So Where So.OdCoCd=Po.OdCoCd and So.OdDmCd=Po.OdDmCd and So.OdSfx=Po.OdSfx and So.OdDmSz=Po.OdDmSz) " + _
            "and convert(decimal(12, 1), (OdOrdQty-OdExpQty))>0  "
  moCn.Execute wPoSql
          
  wStkSql = "Insert Into #TmpStk  " + _
            "Select DCoCd,DLoc,DDmCd,DOdSfx,DDmSz, convert(decimal(12, 1),Sum(DDmDrQty-DDmCrQty)) as qStkQty " + _
            "From DmYyMm " + _
            "/* Join #TmpSo On DDmCd=OdDmCd and DOdSfx=OdSfx and DDmSz=OdDmSz */ " + _
            "Join Loc On DCoCd=LocCoCd and DLoc=LocCd " + _
            "where " + IIF(gs_Partition = ctCurrPrtn, " DPrtKey='" + ctCurrPrtn + "' and ", "") + _
            " LocTyp= 'S' /* and DCoCd= '" + adc("UoCoCdFr") + "' */ " + ws_Cnd1 + _
            " and Exists(Select 'x' From  #TmpSo where DCocd=OdCoCd and DDmCd=OdDmCd and DOdSfx=OdSfx and DDmSz=OdDmSz) " + _
            "Group By DCoCd,DLoc,DDmCd,DOdSfx,DDmSz  " + _
            "Having convert(decimal(12, 1), Sum(DDmDrQty - DDmCrQty)) > 0 "
  moCn.Execute wStkSql
  
    ws_All = ""
  If adc("UoYn1") = "N" Then ws_All = "Having Sum(qOrdQty) > IsNull((Select Sum(qPoQty) From #TmpPo Po where So.OdDmCd=Po.OdDmCd and So.OdSfx=Po.OdSfx and So.OdDmSz=Po.OdDmSz),0) + IsNull((Select Sum(qStkQty) From #TmpStk where DCoCd=So.OdCoCd and DDmCd=So.OdDmCd and DOdSfx=So.OdSfx and DDmSz=So.OdDmSz),0) "

   adc.RepRecSource = " Select " + grpflds1 + _
          ",So.OdCoCd, So.OdDmCd as OdDmCd, So.OdSfx as OdSfx, So.OdDmSz as OdDmSz, Sum(qOrdQty) as qOrdQty, " + _
          "IsNull((Select Sum(qPoQty) From #TmpPo Po where so.OdCoCd=po.OdCoCd and So.OdDmCd=Po.OdDmCd and So.OdSfx=Po.OdSfx and So.OdDmSz=Po.OdDmSz),0) as qPoQty, " + _
          "IsNull((Select Sum(qStkQty) From #TmpStk where DCoCd=So.OdCoCd and DDmCd=So.OdDmCd and DOdSfx=So.OdSfx and DDmSz=So.OdDmSz),0) as qStkQty " + _
          "From #TmpSo So " + _
          "Group By So.OdCoCd,So.OdDmCd,So.OdSfx,So.OdDmSz " + ws_All

   
  If Not (wi_DetPos = 0 Or adc("UoYn") = "N") Then
    Dim ws_SubRepord As String, ws_SubRepPo As String, ws_SubRepStk As String
    
    ws_SubRepord = "Select *, 0 as OdRefIdKey From #TmpSo"
    Call moCn.RepRes(Rep.SubRepOpnOrd.OpenSubreport, ws_SubRepord + ctMaxDopOpt)
    
    ws_SubRepStk = "Select * From #TmpStk"
    Call moCn.RepRes(Rep.SubRepStk.OpenSubreport, ws_SubRepStk + ctMaxDopOpt)
  
    ws_SubRepPo = "Select * From #TmpPo"
    Call moCn.RepRes(Rep.SubRepOpnPo.OpenSubreport, ws_SubRepPo + ctMaxDopOpt)
  Else
    Rep.SecSubRep1.Suppress = True: Rep.SecSubRep2.Suppress = True
  End If
  
  If wi_DetPos = 0 Then Rep.DetLine.Suppress = True

  Dim ws_LocScope As String
  ws_LocScope = IIF(adc("UoIssLocFr") = "", "", "From: " + adc("UoIssLocFr")) + IIF(adc("UoIssLocTo") = "", "", "   To: " + adc("UoIssLocFr")) + IIF(adc("UoIssLocSel") = "", "", "    Selection: " + adc("UoIssLocSel"))
  If ws_LocScope <> "" Then ws_LocScope = "Loc " + ws_LocScope
  Rep.TxtLocScope.SetText ws_LocScope
  Rep.TxtForCoCd.SetText adc("UoCoCdFr")
  CRV_REP.DisplayGroupTree = False
  
Call DispReport(adc) '****** Sachin 3.02 - Display Report ******
Exit Sub

Errhnd:
moCn.Execute "Delete From #TmpSo"
moCn.Execute "Delete From #TmpPo"
moCn.Execute "Delete From #TmpStk"
Call ChkSelScope(Err.Number, Err.Description)
End Sub


Private Function GetCnd(ByVal ws_Typ As String) As String
Dim ws_RepCnd As String
If ws_Typ = "R" Then
    adc("UoOmTcFr").CmpStr = "": adc("UoOmTcTo").CmpStr = ""
    adc("UoOmYyFr").CmpStr = "": adc("UoOmYyTo").CmpStr = ""
    adc("UoOmChrFr").CmpStr = "": adc("UoOmChrTo").CmpStr = ""
    adc("UoOmNoFr").CmpStr = "": adc("UoOmNoTo").CmpStr = ""
    adc("UoOdSrFr").CmpStr = "": adc("UoOdSrTo").CmpStr = ""
    adc("UoDmCtgFr").CmpStr = "": adc("UoDmCtgTo").CmpStr = "": adc("UoDmCtgSel").CmpStr = ""
    adc("UoDmCdFr").CmpStr = "": adc("UoDmCdTo").CmpStr = "": adc("UoDmCdSel").CmpStr = ""
    adc("UoCmCdFr").CmpStr = "": adc("UoCmCdTo").CmpStr = "": adc("UoCmCdSel").CmpStr = ""
    adc("UoPoNoFr").CmpStr = "": adc("UoPoNoTo").CmpStr = "": adc("UoPoNoSel").CmpStr = ""
    adc("UoOdPoNoFr").CmpStr = "": adc("UoOdPoNoTo").CmpStr = "": adc("UoOdPoNoSel").CmpStr = ""
    'vk.19 sub customer added
    adc("UoOdSubCustFr").CmpStr = "": adc("UoOdSubCustTo").CmpStr = ""
    adc("UoDtFr").CmpStr = "": adc("UoDtTo").CmpStr = ""
    adc("UoDelDtFr").CmpStr = "": adc("UoDelDtTo").CmpStr = ""
    adc("UoIssLocFr").CmpStr = "DLoc >=": adc("UoIssLocTo").CmpStr = "DLoc <=": adc("UoIssLocSel").CmpStr = "DLoc In"
    adc("UoCoCdFr").CmpStr = "DCoCd >=": adc("UoCoCdTo").CmpStr = "DCoCd <=": adc("UoCoCdSel").CmpStr = "DCoCd In"
    ws_RepCnd = adc.RepCond
    GetCnd = IIF(ws_RepCnd <> "", " and " + ws_RepCnd, "")
    
ElseIf ws_Typ = "D" Then
    adc("UoOmTcFr").CmpStr = "OdTc>= ": adc("UoOmTcTo").CmpStr = "OdTc<= "
    adc("UoOmYyFr").CmpStr = "OdYy>= ": adc("UoOmYyTo").CmpStr = "OdYy<= "
    adc("UoOmChrFr").CmpStr = "OdChr>= ": adc("UoOmChrTo").CmpStr = "OdChr<= "
    adc("UoOmNoFr").CmpStr = "OdNo>= ": adc("UoOmNoTo").CmpStr = "OdNo<= "
    adc("UoOdSrFr").CmpStr = "OdSr>= ": adc("UoOdSrTo").CmpStr = "OdSr<= "
    adc("UoDmCtgFr").CmpStr = "DmCtg >=": adc("UoDmCtgTo").CmpStr = "DmCtg <=": adc("UoDmCtgSel").CmpStr = "DmCtg In"
    adc("UoDmCdFr").CmpStr = "OdDmCd >=": adc("UoDmCdTo").CmpStr = "OdDmCd <=": adc("UoDmCdSel").CmpStr = "OdDmCd In"
    adc("UoCmCdFr").CmpStr = "OmCmCd >=": adc("UoCmCdTo").CmpStr = "OmCmCd <=": adc("UoCmCdSel").CmpStr = "OmCmCd In"
    adc("UoPoNoFr").CmpStr = "OmPoNo >=": adc("UoPoNoTo").CmpStr = "OmPoNo <=": adc("UoPoNoSel").CmpStr = "OmPoNo In"
    adc("UoOdPoNoFr").CmpStr = "OdPoNo >=": adc("UoOdPoNoTo").CmpStr = "OdPoNo <=": adc("UoOdPoNoSel").CmpStr = "OdPoNo in"
    'vk.19 sub customer added
    adc("UoOdSubCustFr").CmpStr = "OdSubCust >=": adc("UoOdSubCustTo").CmpStr = "OdSubCust <="
    adc("UoDtFr").CmpStr = "OmDt >=": adc("UoDtTo").CmpStr = "OmDt <="
    adc("UoDelDtFr").CmpStr = "OdDelDt >=": adc("UoDelDtTo").CmpStr = "OdDelDt <="
    adc("UoIssLocFr").CmpStr = "": adc("UoIssLocTo").CmpStr = "": adc("UoIssLocSel").CmpStr = ""
    adc("UoCoCdFr").CmpStr = "OdCoCd >=": adc("UoCoCdTo").CmpStr = "OdCoCd <=": adc("UoCoCdSel").CmpStr = "OdCoCd In"
    ws_RepCnd = adc.RepCond
    GetCnd = IIF(ws_RepCnd <> "", " and " + ws_RepCnd, "")
  End If

End Function

Private Sub ChkBoxArr_Click(Index As Integer)
With ChkBoxArr(Index)
  Select Case Index
  Case Is = 0       ' Tag = "ShowDet"
    If ChkBoxArr(0).Value = Checked Then adc("UoYN") = "Y" Else adc("UoYN") = "N"
  Case Is = 1       ' Tag = "ShowAll"
    If ChkBoxArr(1).Value = Checked Then adc("UoYN1") = "Y" Else adc("UoYN1") = "N"
  End Select
End With
End Sub

Private Sub ChkBoxArr_GotFocus(Index As Integer)
  DispMsg ChkBoxArr(Index).ToolTipText, etInfo
End Sub

'*** std code not to be changed
Private Sub Form_GotFocus()
  '*** Jenny Color (For Setting System Colors)
  '(Bef 13/08/05)
  'SetMDISysColorsAfterLogin
  '*** Jenny Color (For Setting System Colors)
End Sub
Private Sub Form_Load()
  '*** Jenny Colour
  SetMDISysColorsAfterLogin   '*** Jenny Color (For Setting System Colors)
  
  Dim ws_BtnStr As String
  ws_BtnStr = ""
  Call SetObjColors(Me, ws_BtnStr)
  '*** Jenny Colour
End Sub
Private Sub ADC_Load()
    Set moCn = adc.Connection    '*** (Bef speed) bef 24/11/06
    
     
    ms_MultiCoMod = IIF(moCn.GetFldVal("Select hImpComLoc From Head Where HCd='" + ctSelfCmCd + "' And HCoCd= '" + ctSelfCoCd + "'") = "Y", True, False)
    gb_CoCdFor = False
    
    If ms_MultiCoMod = True Then
      Set adc.FirNKeyCtl = adc("UoCoCdFr")
      adc("UoCoCdFr").Enabled = True: adc("UoCoCdTo").Visible = True: adc("UoCoCdSel").Visible = True
    Else
      Set adc.FirNKeyCtl = adc("UoOmTcFr")
      adc("UoCoCdFr").Enabled = False: adc("UoCoCdTo").Visible = False: adc("UoCoCdSel").Visible = False
    End If
    
    Call InitProp(Me)
    
    '*** Parameters Used For Setting Properties Of UsrOpt Fields ***
    gs_DmTcTyp = "DM"
    gs_OmTcTyp = "JT"
    gs_LocTyp = "'S'"
    '*** Parameters Used For Setting Properties Of UsrOpt Fields ***
    
    Call SetGroupSort
End Sub
Private Sub ATXT_FldChange(Index As Integer, ByVal IdName As String, Cancel As Boolean, ErrMsg As String, ByVal pv_NewValue As Variant)
  '***  Extra Validation Code Here
  
  Call SetProp(Me, IdName, Valid, Cancel, ErrMsg)
End Sub
Private Sub ATXT_FldWhen(Index As Integer, ByVal IdName As String, Cancel As Boolean, ErrMsg As String)
  '***  Extra Field When Code Here
  gb_CoCdFor = True         '*** (Jen 2.14 Next Patch 1)
  
  SetProp Me, IdName, When
End Sub
Private Sub atxt_ReCalculate(Index As Integer, ByVal IdName As String, ByVal OldVal As Variant, ByVal RowNum As Integer, ByVal SourceName As String, ByVal SourceColName As String)
  '*** Recalc Code Here
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
  If adc("UOYN1") = "" Then adc("UOYN1") = "Y"
  
  If adc("UoYN") = "" Then adc("UoYN") = "N"
  
  If adc("UoYN") = "Y" Then ChkBoxArr(0).Value = Checked Else ChkBoxArr(0).Value = Unchecked
  If adc("UoYN1") = "Y" Then ChkBoxArr(1).Value = Checked Else ChkBoxArr(1).Value = Unchecked
End Sub
Private Sub DispCoNm()
  Dim ws_HName As String
  ws_HName = moCn.GetFldVal("Select HName from Head where HCoCd = '" + adc("UoCoCdFr") + "' and " + _
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
