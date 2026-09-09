VERSION 5.00
Object = "{8767A745-088E-4CA6-8594-073D6D2DE57A}#9.2#0"; "crviewer9.dll"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "tabctl32.ocx"
Object = "{AB2CC32A-E652-42A2-B341-26ECE02F6386}#2.1#0"; "mwfctlforspeed9.ocx"
Begin VB.Form EmrFrpPpcCntrlCpy 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Controller Copy"
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
      TabIndex        =   28
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
      Left            =   11250
      TabIndex        =   27
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
      TabIndex        =   26
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
      Left            =   5370
      TabIndex        =   36
      Top             =   9600
      Width           =   3555
      Begin MwfCtl.MWCTL_MED aktxt 
         Height          =   345
         Index           =   0
         Left            =   210
         TabIndex        =   37
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
         TabIndex        =   38
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
      Left            =   6570
      TabIndex        =   35
      Top             =   9615
      Visible         =   0   'False
      Width           =   855
      Begin CRVIEWER9LibCtl.CRViewer9 CRV_REP 
         Height          =   495
         Left            =   480
         TabIndex        =   32
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
      TabIndex        =   31
      Top             =   60
      Width           =   15255
      Begin TabDlg.SSTab TAB_REP 
         Height          =   9465
         Left            =   15
         TabIndex        =   30
         Top             =   45
         Width           =   15195
         _ExtentX        =   26802
         _ExtentY        =   16695
         _Version        =   393216
         Tabs            =   2
         TabsPerRow      =   2
         TabHeight       =   520
         TabCaption(0)   =   "&Data Range"
         TabPicture(0)   =   "EmrFrpPpcCntrlCpy.frx":0000
         Tab(0).ControlEnabled=   -1  'True
         Tab(0).Control(0)=   "fra_tabrep(0)"
         Tab(0).Control(0).Enabled=   0   'False
         Tab(0).ControlCount=   1
         TabCaption(1)   =   "&Group Sort"
         TabPicture(1)   =   "EmrFrpPpcCntrlCpy.frx":001C
         Tab(1).ControlEnabled=   0   'False
         Tab(1).Control(0)=   "fra_tabrep(1)"
         Tab(1).ControlCount=   1
         Begin VB.Frame fra_tabrep 
            BorderStyle     =   0  'None
            Enabled         =   0   'False
            Height          =   9070
            Index           =   1
            Left            =   -74940
            TabIndex        =   34
            Top             =   360
            Width           =   15105
            Begin MwfCtl.MWCTL_GRP GRP_REP 
               Height          =   360
               Left            =   240
               TabIndex        =   29
               Top             =   360
               Width           =   9285
               _ExtentX        =   16378
               _ExtentY        =   635
            End
         End
         Begin VB.Frame fra_tabrep 
            BorderStyle     =   0  'None
            Height          =   9070
            Index           =   0
            Left            =   60
            TabIndex        =   33
            Top             =   360
            Width           =   15105
            Begin MwfCtl.MWCTL_MED atxt 
               Height          =   285
               Index           =   4
               Left            =   6450
               TabIndex        =   23
               ToolTipText     =   "EnterDesign Code Selection"
               Top             =   2505
               Width           =   7950
               _ExtentX        =   14023
               _ExtentY        =   503
               MaxLength       =   255
               DataType        =   4
               DataField       =   "UODMCDSEL"
               IdName          =   "UODMCDSEL"
               CmpStr          =   "OdDmCd In"
            End
            Begin MwfCtl.MWCTL_MED atxt 
               Height          =   285
               Index           =   1
               Left            =   5055
               TabIndex        =   8
               ToolTipText     =   "Enter To Order Year"
               Top             =   1080
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
               Left            =   2835
               TabIndex        =   5
               ToolTipText     =   "Enter From Order Year"
               Top             =   1080
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
               Left            =   2040
               TabIndex        =   4
               ToolTipText     =   "Enter From Order Tc"
               Top             =   1080
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
               Left            =   3285
               TabIndex        =   6
               ToolTipText     =   "Enter From Order Character"
               Top             =   1080
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
               Index           =   17
               Left            =   4260
               TabIndex        =   7
               ToolTipText     =   "Enter To Order Tc"
               Top             =   1080
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
               Left            =   5475
               TabIndex        =   9
               ToolTipText     =   "Enter To Order Character"
               Top             =   1080
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
               Left            =   2040
               TabIndex        =   10
               ToolTipText     =   "Enter From Order Number"
               Top             =   1365
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
               Left            =   4260
               TabIndex        =   12
               ToolTipText     =   "Enter To Order Number"
               Top             =   1365
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
               Index           =   26
               Left            =   5190
               TabIndex        =   13
               ToolTipText     =   "Enter To Order Serial"
               Top             =   1365
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
               Index           =   34
               Left            =   2970
               TabIndex        =   11
               ToolTipText     =   "Enter From Order Serial"
               Top             =   1365
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
               Index           =   6
               Left            =   2040
               TabIndex        =   0
               ToolTipText     =   "Enter Company Code"
               Top             =   510
               Width           =   705
               _ExtentX        =   1244
               _ExtentY        =   503
               Enabled         =   0   'False
               MaxLength       =   3
               DataType        =   4
               DataField       =   "UOCOCDFR"
               IdName          =   "UOCOCDFR"
               CmpStr          =   "OpdCoCd = "
            End
            Begin MwfCtl.MWCTL_MED atxt 
               Height          =   285
               Index           =   7
               Left            =   4260
               TabIndex        =   2
               ToolTipText     =   "Enter To Customer Code"
               Top             =   795
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
               Index           =   8
               Left            =   2040
               TabIndex        =   1
               ToolTipText     =   "Enter From Customer Code"
               Top             =   795
               Width           =   1305
               _ExtentX        =   2302
               _ExtentY        =   503
               MaxLength       =   8
               DataType        =   4
               DataField       =   "UOCMCDFR"
               IdName          =   "UOCMCDFR"
               CmpStr          =   "OmCmCd>="
            End
            Begin MwfCtl.MWCTL_MED atxt 
               Height          =   285
               Index           =   12
               Left            =   4260
               TabIndex        =   15
               ToolTipText     =   "Enter To Delivery Date"
               Top             =   1650
               Width           =   885
               _ExtentX        =   1561
               _ExtentY        =   503
               MaxLength       =   8
               DataType        =   3
               DataField       =   "UODELDTTO"
               IdName          =   "UODELDTTO"
               CmpStr          =   "OdDelDt <="
            End
            Begin MwfCtl.MWCTL_MED atxt 
               Height          =   285
               Index           =   14
               Left            =   2040
               TabIndex        =   14
               ToolTipText     =   "Enter From Delivery Date"
               Top             =   1650
               Width           =   885
               _ExtentX        =   1561
               _ExtentY        =   503
               MaxLength       =   8
               DataType        =   3
               DataField       =   "UODELDTFR"
               IdName          =   "UODELDTFR"
               CmpStr          =   "OdDelDt >="
            End
            Begin MwfCtl.MWCTL_MED atxt 
               Height          =   285
               Index           =   15
               Left            =   4260
               TabIndex        =   19
               ToolTipText     =   "Enter To Design Category To"
               Top             =   2220
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
               Index           =   16
               Left            =   2040
               TabIndex        =   18
               ToolTipText     =   "Enter From Design Category"
               Top             =   2220
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
               Index           =   18
               Left            =   4260
               TabIndex        =   22
               ToolTipText     =   "Enter To Design Code"
               Top             =   2505
               Width           =   2115
               _ExtentX        =   3731
               _ExtentY        =   503
               MaxLength       =   15
               DataType        =   4
               DataField       =   "UODMCDTO"
               IdName          =   "UODMCDTO"
               CmpStr          =   "OdDmCd <="
            End
            Begin MwfCtl.MWCTL_MED atxt 
               Height          =   285
               Index           =   19
               Left            =   2040
               TabIndex        =   21
               ToolTipText     =   "Enter From Design Code"
               Top             =   2505
               Width           =   2115
               _ExtentX        =   3731
               _ExtentY        =   503
               MaxLength       =   15
               DataType        =   4
               DataField       =   "UODMCDFR"
               IdName          =   "UODMCDFR"
               CmpStr          =   "OdDmCd >="
            End
            Begin MwfCtl.MWCTL_MED atxt 
               Height          =   285
               Index           =   22
               Left            =   6450
               TabIndex        =   3
               ToolTipText     =   "Ente Customer Selection"
               Top             =   795
               Width           =   7950
               _ExtentX        =   14023
               _ExtentY        =   503
               MaxLength       =   255
               DataType        =   4
               DataField       =   "UOCMCDSEL"
               IdName          =   "UOCMCDSEL"
               CmpStr          =   "OmCmCD In"
            End
            Begin MwfCtl.MWCTL_MED atxt 
               Height          =   285
               Index           =   25
               Left            =   6450
               TabIndex        =   20
               ToolTipText     =   "Enter Line Selection"
               Top             =   2220
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
               Index           =   27
               Left            =   4260
               TabIndex        =   17
               ToolTipText     =   "Enter To Export Delivery Date"
               Top             =   1935
               Width           =   885
               _ExtentX        =   1561
               _ExtentY        =   503
               MaxLength       =   8
               DataType        =   3
               DataField       =   "UODELDT1TO"
               IdName          =   "UODELDT1TO"
               CmpStr          =   "OdExpDelDt <="
            End
            Begin MwfCtl.MWCTL_MED atxt 
               Height          =   285
               Index           =   28
               Left            =   2040
               TabIndex        =   16
               ToolTipText     =   "Enter From Export Delivery Date"
               Top             =   1935
               Width           =   885
               _ExtentX        =   1561
               _ExtentY        =   503
               MaxLength       =   8
               DataType        =   3
               DataField       =   "UODELDT1FR"
               IdName          =   "UODELDT1FR"
               CmpStr          =   "OdExpDelDt >="
            End
            Begin MwfCtl.MWCTL_MED atxt 
               Height          =   285
               Index           =   3
               Left            =   4260
               TabIndex        =   25
               ToolTipText     =   "Enter To Line"
               Top             =   2790
               Width           =   855
               _ExtentX        =   1508
               _ExtentY        =   503
               MaxLength       =   5
               DataType        =   4
               DataField       =   "UOLINETO"
               IdName          =   "UOLINETO"
               CmpStr          =   "OpLine<="
            End
            Begin MwfCtl.MWCTL_MED atxt 
               Height          =   285
               Index           =   5
               Left            =   2040
               TabIndex        =   24
               ToolTipText     =   "Enter From Line"
               Top             =   2790
               Width           =   855
               _ExtentX        =   1508
               _ExtentY        =   503
               MaxLength       =   5
               DataType        =   4
               DataField       =   "UOLINEFR"
               IdName          =   "UOLINEFR"
               CmpStr          =   "OpLine>="
            End
            Begin VB.Label ALBL 
               BackStyle       =   0  'Transparent
               Caption         =   "Line"
               BeginProperty Font 
                  Name            =   "Arial"
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
               TabIndex        =   56
               Top             =   2790
               Width           =   1995
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
               Left            =   5100
               TabIndex        =   55
               Top             =   1365
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
               Left            =   5400
               TabIndex        =   54
               Top             =   1080
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
               Index           =   11
               Left            =   4980
               TabIndex        =   53
               Top             =   1080
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
               Index           =   6
               Left            =   2880
               TabIndex        =   52
               Top             =   1365
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
               Index           =   4
               Left            =   3210
               TabIndex        =   51
               Top             =   1080
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
               Index           =   0
               Left            =   2760
               TabIndex        =   50
               Top             =   1080
               Width           =   105
            End
            Begin VB.Label ALBL 
               BackStyle       =   0  'Transparent
               Caption         =   "Exp Del Date"
               BeginProperty Font 
                  Name            =   "Arial"
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
               TabIndex        =   49
               Top             =   1935
               Width           =   1995
            End
            Begin VB.Label ALBL 
               BackStyle       =   0  'Transparent
               Caption         =   "Dsg Cd"
               BeginProperty Font 
                  Name            =   "Arial"
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
               TabIndex        =   48
               Top             =   2505
               Width           =   1995
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
               Index           =   8
               Left            =   60
               TabIndex        =   47
               ToolTipText     =   " --------------------------------------------------------"
               Top             =   2220
               Width           =   1995
            End
            Begin VB.Label ALBL 
               BackStyle       =   0  'Transparent
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
               ForeColor       =   &H00000000&
               Height          =   285
               Index           =   7
               Left            =   60
               TabIndex        =   46
               Top             =   1650
               Width           =   1995
            End
            Begin VB.Label ALBL 
               BackStyle       =   0  'Transparent
               Caption         =   "Customer Cd"
               BeginProperty Font 
                  Name            =   "Arial"
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
               TabIndex        =   45
               Top             =   795
               Width           =   1995
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
               TabIndex        =   44
               Top             =   510
               Width           =   1995
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
               TabIndex        =   43
               Top             =   1080
               Width           =   1995
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
               Left            =   60
               TabIndex        =   42
               Top             =   1365
               Width           =   1995
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
               Left            =   6450
               TabIndex        =   41
               Top             =   60
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
               TabIndex        =   40
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
               Left            =   2040
               TabIndex        =   39
               Top             =   60
               Width           =   1335
            End
         End
      End
   End
End
Attribute VB_Name = "EmrFrpPpcCntrlCpy"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim Rep As New EmrRepPpcCntrlCpy
Dim moCn As MwfLib.MDOConnection   '*** (Bef speed) bef 24/11/06

Private Sub SetGroupSort()
  GRP_REP.Add "OrderNo", "max(OdTc+OdYy+OdChr+str(OdNo))", "max(OdTc+'/'+OdYy+'/'+OdChr+'/'+LTrim(Str(OdNo)))", "", "", "Customer Code,Customer PO", "", "", "'('+max(OmCmCd)+ '/'+ max(OmPONo)+ ')' "
  GRP_REP.Add "Customer PO", "max(OmPoNo)", "", "", "", "", "", "", ""
  '******Geeta*****************Emr206****05/04/04*******
  '**********If UserLevel>6 then don't display the Customer Name
  GRP_REP.Add "Customer Code", "max(OmCmCd)", "", "", "", "", "", "", "(Select (case when (Select PNum from Param where Ptyp = 'USR' and PmCd = '" + Trim(EmrFrmMDI.TXT_UCD.text) + "')  <= 5 then  CmName else CmCd end ) From CustMst Where CmCtg= 'C' and CmCd= max(OmCmCd)) "
  GRP_REP.Add "Design Ctg", "max(DmCtg)", "", "", "", "", "", "", "(Select PDesc from Param where PTyp= 'DMCTG' and PMCd= max(DmCtg))"
  GRP_REP.Add "Design Code", "max(OdDmCd)", "", "", "", "Design Ctg", "", "", "(Select Substring(DmDesc,1,120) From DsgMst Where DmTcTyp= 'DM' and DmCd= max(OdDmCd) and DmSz= '') "
  GRP_REP.Add "Karatage", "max(OdKt)", "", "", "", "", "", "", ""
  '6.1 Exp Del Date,Prd Del Date,Order Date
  GRP_REP.Add "Exp Del Date", "convert(varchar(8), max(OdExpDelDt), 112)", "convert(varchar, max(OdExpDelDt), " + CStr(DtFmtSlahVal()) + ")", "", "", "", "", "", ""
  GRP_REP.Add "Prd Del Date", "convert(varchar(8), max(OdDelDt), 112)", "convert(varchar, max(OdDelDt), " + CStr(DtFmtSlahVal()) + ")", "", "", "", "", "", ""
  GRP_REP.Add "Order Date", "convert(varchar(8), max(OmDt), 112)", "convert(varchar, max(OmDt), " + CStr(DtFmtSlahVal()) + ")", "", "", "", "", "", ""
  GRP_REP.Add "Line", "max(OpLine)", "", "", "", "", "", "", "(Select vPDesc from vParam where vPCoCd= max(OpCoCd) and vPTyp= 'LINE' and vPmCd = max(OpLine))"
End Sub
Private Sub ADC_KeyWhen()
  adc("UoUsrCd") = adc.UsrCd
  adc("UoMnuCd") = adc.MenuCd
End Sub
Private Sub ADC_NonKeyWhen(pr_Cancel As Boolean, pr_ErrMsg As String)
    adc("UoCoCdFr") = gs_CoCd
    'adc ("UoDelDt1Fr")
    'adc("UoExpDelDtFr") = "01/01/80"
    'adc("UoExpDelDtTo") = "01/01/80"
    'adc("UoDelDtFr") = "01/01/80"
    'adc("UoDelDtTo") = "01/01/80"
End Sub
Private Sub adc_setreprecsource()
  '*** Report Sql ***
  
'**** Before 090205 ****'
  'Select space(30) as Grp1, space(30) as Grp2, space(30) as Grp3,
  'space(30) as DGrp1, space(30) as DGrp2, space(30) as DGrp3,
  'space(30) as G1Desc, space(30) as G2Desc, space(30) as G3Desc,
  'OpdTc, OpdYy, OpdChr, OpdNo, OpdSr, OmCmCd , OdDmCd,
  'OdSfx, OdDmSz,OddMCol,OdKt,OdExpDelDt, OdExpQty, OdDelDt,OdPrdQty,
  'OpdPlnDT As qMinDt1, OpdPlnDT As qMaxDt1, OpdPlnDT As qMinDt2, OpdPlnDT As qMaxDt2, OpdPlnDT As qMinDt3,
  'OpdPlnDT As qMaxDt3, OpdPlnDT As qMinDt4, OpdPlnDT As qMaxDt4, OpdPlnDT As qMinDt5, OpdPlnDT As qMaxDt5
  'From OrdPlnDt, OrdDsg, OrdMst where 1 = 2
'**** Before 090205 ****'

  '****** (Jen 2.14 Next) (changed the fields WSet, Cas, Fin, HSet, Ei to P1, P2, P3, P4, P5) ******
  '********* Zubin(090205) *********'
  'Select space(30) as Grp1, space(30) as Grp2, space(30) as Grp3,
  'space(30) as DGrp1, space(30) as DGrp2, space(30) as DGrp3,
  'space(30) as G1Desc, space(30) as G2Desc, space(30) as G3Desc,
  'OpdTc, OpdYy, OpdChr, OpdNo, OpdSr, OmCmCd , OdDmCd,
  'OdSfx, OdDmSz, OdDmCol,OdKt,OdExpDelDt, OdExpQty, OdDelDt,OdPrdQty,
  'OpdPlnDT As qP1MinDt, OpdPlnDT As qP1MaxDt, OpdPlnDT As qP2MinDt, OpdPlnDT As qP2MaxDt, OpdPlnDT As qP3MinDt,
  'OpdPlnDT As qP3MaxDt, OpdPlnDT As qP4MinDt, OpdPlnDT As qP4MaxDt, OpdPlnDT As qP5MinDt, OpdPlnDT As qP5MaxDt
  'From OrdPlnDt, OrdDsg, OrdMst where 1 = 2
  '********* Zubin(090205) *********'

  '*** Variables Displayed for the  Group Fields and Conditions and Sql String
  Dim grpflds As String, wCnd As String, wSqlStrg As String
    
  Set Rep = Nothing
  Set adc.RepSource = Rep
  Call DispCoNm
  Call SetMwName(Rep)   ' ***** Manali 3.6.1 - Microway footer
  Call GRP_REP.Gen3LRep
  grpflds = GRP_REP.GrpFldLst

  '*** To check if Detail is selected ***
  Dim wGrp(3) As String, i As Integer, wDetPos As Integer
  wGrp(1) = GRP_REP.Value(gltGroup1)
  wGrp(2) = GRP_REP.Value(gltGroup2)
  wGrp(3) = GRP_REP.Value(gltGroup3)
  wDetPos = 0
     
  Rep.TxtHead.SetText ("Controller Copy")
  wCnd = adc.RepCond
  wCnd = IIF(wCnd <> "", " and " + wCnd, "")

  '****** (Jen 2.14 Next) (changed the fields WSet, Cas, Fin, HSet, Ei to P1, P2, P3, P4, P5) ******
  '6.1
  '6.1 all date format changed
  wSqlStrg = "Set dateformat " + DtFmtStr() + " " + _
             "Select " + grpflds + ", OpdTc, OpdYy, OpdChr, OpdNo, OpdSr, max(OmCmCd) as OmCmCd, max(OdDmCd) as OdDmCd, " + _
             " max(OdSfx)as OdSfx, max(OdDmSz) as OdDmSz, max(OdDmCol) as OdDmCol, max(OdKt) as OdKt, " + _
             " max(OdExpDelDt) as OdExpDelDt, max(OdExpQty) as OdExpQty, max(OdDelDt) as OdDelDt, " + _
             " max(OdPrdQty) as OdPrdQty, " + _
             " min((Case when Str(OpdP1, 16, 1) <> Str(0, 16, 1) then OpdPlnDt else '" + IIF(UCase(MWLib.DtSetting) = "MM/DD/YY", "12/01/28", "01/12/28") + "' end)) as qP1MinDt, " + _
             " max((Case when Str(OpdP1, 16, 1) <> Str(0, 16, 1) then OpdPlnDt else '01/01/80' end)) as qP1MaxDt, " + _
             " min((Case when Str(OpdP2, 16, 1) <> Str(0, 16, 1) then OpdPlnDt else '" + IIF(UCase(MWLib.DtSetting) = "MM/DD/YY", "12/01/28", "01/12/28") + "' end)) as qP2MinDt, " + _
             " max((Case when Str(OpdP2, 16, 1) <> Str(0, 16, 1) then OpdPlnDt else '01/01/80' end)) as qP2MaxDt, " + _
             " min((Case when Str(OpdP3, 16, 1) <> Str(0, 16, 1) then OpdPlnDt else '" + IIF(UCase(MWLib.DtSetting) = "MM/DD/YY", "12/01/28", "01/12/28") + "' end)) as qP3MinDt, " + _
             " max((Case when Str(OpdP3, 16, 1) <> Str(0, 16, 1) then OpdPlnDt else '01/01/80' end)) as qP3MaxDt, " + _
             " min((Case when Str(OpdP4, 16, 1) <> Str(0, 16, 1) then OpdPlnDt else '" + IIF(UCase(MWLib.DtSetting) = "MM/DD/YY", "12/01/28", "01/12/28") + "' end)) as qP4MinDt, " + _
             " max((Case when Str(OpdP4, 16, 1) <> Str(0, 16, 1) then OpdPlnDt else '01/01/80' end)) as qP4MaxDt, " + _
             " min((Case when Str(OpdP5, 16, 1) <> Str(0, 16, 1) then OpdPlnDt else '" + IIF(UCase(MWLib.DtSetting) = "MM/DD/YY", "12/01/28", "01/12/28") + "' end)) as qP5MinDt, " + _
             " max((Case when Str(OpdP5, 16, 1) <> Str(0, 16, 1) then OpdPlnDt else '01/01/80' end)) as qP5MaxDt " + _
             " ,'" + UCase(MWLib.DtSetting) + "' as qDtsetting From OrdPlnDt " + _
             " join OrdPln on OpTc= OpdTc and OpYy= OpdYy and OpChr= OpdChr and OpNo= OpdNo and OpSr= OpdSr " + _
             " join OrdDsg on OdTc= OpdTc and OdYy= OpdYy and OdChr= OpdChr and OdNo= OpdNo and OdSr= OpdSr " + _
             " join OrdMst on OmTc= OpdTc and OmYy= OpdYy and OmChr= OpdChr and OmNo= OpdNo " + _
             " Join DsgMst on /* DmIdNo=OdDmIdNo and */ DmTcTyp= 'DM' and DmCd= OdDmCd and DmSz= ''" + _
             " where OpdCoCd= '" + gs_CoCd + "' " + wCnd + _
             " group by OpdTc, OpdYy, OpdChr, OpdNo, OpdSr "

             '" min((Case when Str(OpdWSet, 16, 1) <> Str(0, 16, 1) then convert(varchar(8), OpdPlnDt, 112) else convert(varchar(8), '01/12/28', 112) end)) as qWSetMinDt, " + _
             '" max((Case when Str(OpdWSet, 16, 1) <> Str(0, 16, 1) then convert(varchar(8), OpdPlnDt, 112) else convert(varchar(8), '01/01/80', 112) end)) as qWSetMaxDt " + _

  adc.RepRecSource = wSqlStrg + ctMaxDopOpt '****** Sachin 3.01 added ctMaxDopOpt
  
  '*** (Jen 2.14 Next)
  Dim wo_RsColHd As MwfLib.MDORowSet
  Set wo_RsColHd = moCn.OpenRes("select  max(case when PNum= 1 then PValue else '' end) as qP1, " + _
                                "        max(case when PNum= 2 then PValue else '' end) as qP2, " + _
                                "        max(case when PNum= 3 then PValue else '' end) as qP3, " + _
                                "        max(case when PNum= 4 then PValue else '' end) as qP4, " + _
                                "        max(case when PNum= 5 then PValue else '' end) as qP5 " + _
                                "From Param where PTyp= 'PRDSTG'")
  With Rep
    Do While Not (wo_RsColHd.EOF Or wo_RsColHd.BOF)
      .hP1.SetText wo_RsColHd!qP1: .h1P1.SetText wo_RsColHd!qP1
      .hP2.SetText wo_RsColHd!qP2: .h1P2.SetText wo_RsColHd!qP2
      .hP3.SetText wo_RsColHd!qP3: .h1P3.SetText wo_RsColHd!qP3
      .hP4.SetText wo_RsColHd!qP4: .h1P4.SetText wo_RsColHd!qP4
      .hP5.SetText wo_RsColHd!qP5: .h1P5.SetText wo_RsColHd!qP5
      wo_RsColHd.MoveNext
    Loop
  End With
  '*** (Jen 2.14 Next)
  
  CRV_REP.DisplayGroupTree = False
  Call DispReport(adc) '****** Sachin 3.02 - Display Report ******
  
End Sub

Private Sub ADC_Load()
  Set moCn = adc.Connection    '*** (Bef speed) bef 24/11/06
  Set adc.FirNKeyCtl = adc("UoCmCdFr")
  Call InitProp(Me)
  Call SetGroupSort
End Sub
Private Sub ATXT_FldChange(Index As Integer, ByVal IdName As String, Cancel As Boolean, ErrMsg As String, ByVal pv_NewValue As Variant)
  Select Case IdName
  Case Is = UCase("UoCmCdFr"), UCase("UoCmCdTo"), UCase("UoCmCdSel")
    gs_CmCtg = "C"
  End Select
  
  Call SetProp(Me, IdName, Valid, Cancel, ErrMsg)
End Sub
Private Sub ATXT_FldWhen(Index As Integer, ByVal IdName As String, Cancel As Boolean, ErrMsg As String)
  Select Case IdName
  Case Is = UCase("UoCmCdFr"), UCase("UoCmCdTo"), UCase("UoCmCdSel")
    gs_CmCtg = "C"
  End Select
  
  SetProp Me, IdName, When
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

'*** std code not to be changed ***
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
Private Sub tab_REP_Click(PreviousTab As Integer)
    fra_tabrep(TAB_REP.Tab).Enabled = True
    fra_tabrep(PreviousTab).Enabled = False
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
