VERSION 5.00
Object = "{8767A745-088E-4CA6-8594-073D6D2DE57A}#9.2#0"; "crviewer9.dll"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "tabctl32.ocx"
Object = "{AB2CC32A-E652-42A2-B341-26ECE02F6386}#2.1#0"; "mwfctlforspeed9.ocx"
Begin VB.Form EmrFrpRmRt 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Raw Material Rate Master Listing"
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
      Left            =   10620
      TabIndex        =   36
      TabStop         =   0   'False
      Top             =   9600
      Visible         =   0   'False
      Width           =   495
      _ExtentX        =   873
      _ExtentY        =   873
   End
   Begin MwfCtl.MWCTL_DST DST_REP 
      Height          =   570
      Left            =   4395
      TabIndex        =   35
      TabStop         =   0   'False
      Top             =   9630
      Width           =   6450
      _ExtentX        =   11377
      _ExtentY        =   1005
   End
   Begin VB.Frame fra_keys 
      BackColor       =   &H00C0C0C0&
      BorderStyle     =   0  'None
      Height          =   465
      Left            =   4785
      TabIndex        =   45
      Top             =   9630
      Width           =   3555
      Begin MwfCtl.MWCTL_MED aktxt 
         Height          =   345
         Index           =   0
         Left            =   210
         TabIndex        =   46
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
         TabIndex        =   47
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
      Left            =   6000
      TabIndex        =   44
      Top             =   9645
      Visible         =   0   'False
      Width           =   855
      Begin CRVIEWER9LibCtl.CRViewer9 CRV_REP 
         Height          =   495
         Left            =   480
         TabIndex        =   41
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
      TabIndex        =   37
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
      TabIndex        =   40
      Top             =   60
      Width           =   15255
      Begin TabDlg.SSTab TAB_REP 
         Height          =   9465
         Left            =   15
         TabIndex        =   39
         Top             =   45
         Width           =   15195
         _ExtentX        =   26802
         _ExtentY        =   16695
         _Version        =   393216
         Tabs            =   2
         TabsPerRow      =   2
         TabHeight       =   520
         TabCaption(0)   =   "&Data Range"
         TabPicture(0)   =   "EmrFrpRmRt.frx":0000
         Tab(0).ControlEnabled=   -1  'True
         Tab(0).Control(0)=   "fra_tabrep(0)"
         Tab(0).Control(0).Enabled=   0   'False
         Tab(0).ControlCount=   1
         TabCaption(1)   =   "&Group Sort"
         TabPicture(1)   =   "EmrFrpRmRt.frx":001C
         Tab(1).ControlEnabled=   0   'False
         Tab(1).Control(0)=   "fra_tabrep(1)"
         Tab(1).ControlCount=   1
         Begin VB.Frame fra_tabrep 
            BorderStyle     =   0  'None
            Enabled         =   0   'False
            Height          =   9070
            Index           =   1
            Left            =   -74940
            TabIndex        =   43
            Top             =   360
            Width           =   15105
            Begin MwfCtl.MWCTL_GRP GRP_REP 
               Height          =   360
               Left            =   240
               TabIndex        =   38
               Top             =   360
               Width           =   9765
               _ExtentX        =   17224
               _ExtentY        =   635
            End
         End
         Begin VB.Frame fra_tabrep 
            BorderStyle     =   0  'None
            Height          =   9070
            Index           =   0
            Left            =   60
            TabIndex        =   42
            Top             =   360
            Width           =   15105
            Begin MwfCtl.MWCTL_MED atxt 
               Height          =   285
               Index           =   1
               Left            =   3645
               TabIndex        =   3
               ToolTipText     =   "Enter To Rate Type "
               Top             =   1395
               Width           =   1305
               _ExtentX        =   2302
               _ExtentY        =   503
               MaxLength       =   8
               DataField       =   "UOPMCDTO"
               ReCalcParent    =   "UOCURCDFR"
               IdName          =   "UOPMCDTO"
            End
            Begin MwfCtl.MWCTL_MED atxt 
               Height          =   285
               Index           =   2
               Left            =   6120
               TabIndex        =   4
               ToolTipText     =   "Enter Rate Type Selection"
               Top             =   1395
               Width           =   8790
               _ExtentX        =   15505
               _ExtentY        =   503
               MaxLength       =   255
               DataField       =   "UOPMCDSEL"
               ReCalcParent    =   "UOCURCDFR"
               IdName          =   "UOPMCDSEL"
            End
            Begin MwfCtl.MWCTL_MED atxt 
               Height          =   285
               Index           =   0
               Left            =   1215
               TabIndex        =   2
               ToolTipText     =   "Enter From Rate Type "
               Top             =   1395
               Width           =   1305
               _ExtentX        =   2302
               _ExtentY        =   503
               MaxLength       =   8
               DataField       =   "UOPMCDFR"
               ReCalcParent    =   "UOCURCDFR"
               IdName          =   "UOPMCDFR"
            End
            Begin MwfCtl.MWCTL_MED atxt 
               Height          =   285
               Index           =   8
               Left            =   3645
               TabIndex        =   15
               ToolTipText     =   "Enter Raw Material Code To Range"
               Top             =   2535
               Width           =   2385
               _ExtentX        =   4207
               _ExtentY        =   503
               MaxLength       =   16
               DataField       =   "UORMCDTO"
               IdName          =   "UORMCDTO"
            End
            Begin MwfCtl.MWCTL_MED atxt 
               Height          =   285
               Index           =   15
               Left            =   6120
               TabIndex        =   16
               ToolTipText     =   "Enter Raw Material Code Selection"
               Top             =   2535
               Width           =   8790
               _ExtentX        =   15505
               _ExtentY        =   503
               MaxLength       =   255
               DataField       =   "UORMCDSEL"
               IdName          =   "UORMCDSEL"
            End
            Begin MwfCtl.MWCTL_MED atxt 
               Height          =   285
               Index           =   4
               Left            =   3645
               TabIndex        =   9
               ToolTipText     =   "Enter Raw Material Sub Category To Range"
               Top             =   1965
               Width           =   885
               _ExtentX        =   1561
               _ExtentY        =   503
               MaxLength       =   5
               DataField       =   "UORMSCTGTO"
               IdName          =   "UORMSCTGTO"
            End
            Begin MwfCtl.MWCTL_MED atxt 
               Height          =   285
               Index           =   5
               Left            =   6120
               TabIndex        =   10
               ToolTipText     =   "Enter Raw Material Sub Category Selection"
               Top             =   1965
               Width           =   8790
               _ExtentX        =   15505
               _ExtentY        =   503
               MaxLength       =   255
               DataField       =   "UORMSCTGSEL"
               IdName          =   "UORMSCTGSEL"
            End
            Begin MwfCtl.MWCTL_MED atxt 
               Height          =   285
               Index           =   6
               Left            =   1215
               TabIndex        =   5
               ToolTipText     =   "Enter From Raw Material Category"
               Top             =   1680
               Width           =   465
               _ExtentX        =   820
               _ExtentY        =   503
               MaxLength       =   2
               DataField       =   "UORMCTGFR"
               IdName          =   "UORMCTGFR"
            End
            Begin MwfCtl.MWCTL_MED atxt 
               Height          =   285
               Index           =   7
               Left            =   1215
               TabIndex        =   8
               ToolTipText     =   "Enter Raw Material Sub Category From Range"
               Top             =   1965
               Width           =   885
               _ExtentX        =   1561
               _ExtentY        =   503
               MaxLength       =   5
               DataField       =   "UORMSCTGFR"
               IdName          =   "UORMSCTGFR"
            End
            Begin MwfCtl.MWCTL_MED atxt 
               Height          =   285
               Index           =   9
               Left            =   1215
               TabIndex        =   14
               ToolTipText     =   "Enter Raw Material Code From Range"
               Top             =   2535
               Width           =   2385
               _ExtentX        =   4207
               _ExtentY        =   503
               MaxLength       =   16
               DataField       =   "UORMCDFR"
               IdName          =   "UORMCDFR"
            End
            Begin MwfCtl.MWCTL_MED atxt 
               Height          =   285
               Index           =   12
               Left            =   3645
               TabIndex        =   6
               ToolTipText     =   "Enter To Raw Material Category"
               Top             =   1680
               Width           =   465
               _ExtentX        =   820
               _ExtentY        =   503
               MaxLength       =   2
               DataField       =   "UORMCTGTO"
               IdName          =   "UORMCTGTO"
            End
            Begin MwfCtl.MWCTL_MED atxt 
               Height          =   285
               Index           =   13
               Left            =   6120
               TabIndex        =   7
               ToolTipText     =   "Enter Raw Material Category Selection"
               Top             =   1680
               Width           =   8790
               _ExtentX        =   15505
               _ExtentY        =   503
               MaxLength       =   255
               DataField       =   "UORMCTGSEL"
               IdName          =   "UORMCTGSEL"
            End
            Begin MwfCtl.MWCTL_MED atxt 
               Height          =   285
               Index           =   28
               Left            =   3645
               TabIndex        =   18
               ToolTipText     =   "Enter To Design Category"
               Top             =   2820
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
               Left            =   6120
               TabIndex        =   19
               ToolTipText     =   "Enter Design Category Selection"
               Top             =   2820
               Width           =   8790
               _ExtentX        =   15505
               _ExtentY        =   503
               MaxLength       =   255
               DataType        =   4
               DataField       =   "UODMCTGSEL"
               IdName          =   "UODMCTGSEL"
            End
            Begin MwfCtl.MWCTL_MED atxt 
               Height          =   285
               Index           =   30
               Left            =   1215
               TabIndex        =   17
               ToolTipText     =   "Enter From Design Category"
               Top             =   2820
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
               Index           =   40
               Left            =   3645
               TabIndex        =   21
               ToolTipText     =   "Enter To Customer"
               Top             =   3105
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
               Left            =   6120
               TabIndex        =   22
               ToolTipText     =   "Enter Customer Selection"
               Top             =   3105
               Width           =   8790
               _ExtentX        =   15505
               _ExtentY        =   503
               MaxLength       =   255
               DataType        =   4
               DataField       =   "UOCMCDSEL"
               IdName          =   "UOCMCDSEL"
            End
            Begin MwfCtl.MWCTL_MED atxt 
               Height          =   285
               Index           =   42
               Left            =   1215
               TabIndex        =   20
               ToolTipText     =   "Enter From Customer"
               Top             =   3105
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
               Index           =   3
               Left            =   1215
               TabIndex        =   0
               ToolTipText     =   "Enter Customer Category"
               Top             =   840
               Width           =   465
               _ExtentX        =   820
               _ExtentY        =   503
               MaxLength       =   2
               DataField       =   "UOCMCTGFR"
               IdName          =   "UOCMCTGFR"
            End
            Begin MwfCtl.MWCTL_MED atxt 
               Height          =   285
               Index           =   48
               Left            =   1215
               TabIndex        =   1
               ToolTipText     =   "Enter From YearMonth"
               Top             =   1125
               Width           =   765
               _ExtentX        =   1349
               _ExtentY        =   503
               Mask            =   "#####0"
               MaxLength       =   6
               DataType        =   1
               DataField       =   "UOYYMMFR"
               IdName          =   "UOYYMMFR"
            End
            Begin MwfCtl.MWCTL_MED atxt 
               Height          =   285
               Index           =   14
               Left            =   3645
               TabIndex        =   24
               ToolTipText     =   "Enter To Raw Material Range Size"
               Top             =   3390
               Width           =   1605
               _ExtentX        =   2831
               _ExtentY        =   503
               MaxLength       =   10
               DataField       =   "UORNGSZTO"
               IdName          =   "UORNGSZTO"
            End
            Begin MwfCtl.MWCTL_MED atxt 
               Height          =   285
               Index           =   17
               Left            =   6120
               TabIndex        =   25
               ToolTipText     =   "Enter Raw Material Range Size Selection"
               Top             =   3390
               Width           =   8790
               _ExtentX        =   15505
               _ExtentY        =   503
               MaxLength       =   255
               DataField       =   "UORNGSZSEL"
               IdName          =   "UORNGSZSEL"
            End
            Begin MwfCtl.MWCTL_MED atxt 
               Height          =   285
               Index           =   19
               Left            =   1215
               TabIndex        =   23
               ToolTipText     =   "Enter From Raw Material Range Size"
               Top             =   3390
               Width           =   1605
               _ExtentX        =   2831
               _ExtentY        =   503
               MaxLength       =   10
               DataField       =   "UORNGSZFR"
               IdName          =   "UORNGSZFR"
            End
            Begin MwfCtl.MWCTL_MED atxt 
               Height          =   285
               Index           =   10
               Left            =   3645
               TabIndex        =   30
               ToolTipText     =   "Enter To Size To Weight Group"
               Top             =   3960
               Width           =   885
               _ExtentX        =   1561
               _ExtentY        =   503
               MaxLength       =   5
               DataField       =   "UOSTWGRPTO"
               IdName          =   "UOSTWGRPTO"
            End
            Begin MwfCtl.MWCTL_MED atxt 
               Height          =   285
               Index           =   11
               Left            =   6120
               TabIndex        =   31
               ToolTipText     =   "Enter Size To Weight Group Selection"
               Top             =   3960
               Width           =   8790
               _ExtentX        =   15505
               _ExtentY        =   503
               MaxLength       =   255
               DataField       =   "UOSTWGRPSEL"
               IdName          =   "UOSTWGRPSEL"
            End
            Begin MwfCtl.MWCTL_MED atxt 
               Height          =   285
               Index           =   16
               Left            =   1215
               TabIndex        =   29
               ToolTipText     =   "Enter From Size To Weight Group"
               Top             =   3960
               Width           =   885
               _ExtentX        =   1561
               _ExtentY        =   503
               MaxLength       =   5
               DataField       =   "UOSTWGRPFR"
               IdName          =   "UOSTWGRPFR"
            End
            Begin MwfCtl.MWCTL_MED atxt 
               Height          =   285
               Index           =   18
               Left            =   3645
               TabIndex        =   27
               ToolTipText     =   "Enter To Loss Category "
               Top             =   3675
               Width           =   885
               _ExtentX        =   1561
               _ExtentY        =   503
               MaxLength       =   5
               DataField       =   "UOLSCTGTO"
               IdName          =   "UOLSCTGTO"
            End
            Begin MwfCtl.MWCTL_MED atxt 
               Height          =   285
               Index           =   20
               Left            =   6120
               TabIndex        =   28
               ToolTipText     =   "Enter Loss Category Selection"
               Top             =   3675
               Width           =   8790
               _ExtentX        =   15505
               _ExtentY        =   503
               MaxLength       =   255
               DataField       =   "UOLSCTGSEL"
               IdName          =   "UOLSCTGSEL"
            End
            Begin MwfCtl.MWCTL_MED atxt 
               Height          =   285
               Index           =   21
               Left            =   1215
               TabIndex        =   26
               ToolTipText     =   "Enter From Loss Category "
               Top             =   3675
               Width           =   885
               _ExtentX        =   1561
               _ExtentY        =   503
               MaxLength       =   5
               DataField       =   "UOLSCTGFR"
               IdName          =   "UOLSCTGFR"
            End
            Begin MwfCtl.MWCTL_MED atxt 
               Height          =   285
               Index           =   44
               Left            =   1215
               TabIndex        =   32
               ToolTipText     =   "Currency to be entered only for TcTyp = ‘RM"
               Top             =   4245
               Width           =   885
               _ExtentX        =   1561
               _ExtentY        =   503
               MaxLength       =   5
               DataType        =   4
               DataField       =   "UOCURCDFR"
               ReCalcOn        =   "UOPMCDFR,UOPMCDTO,UOPMCDSEL"
               IdName          =   "UOCURCDFR"
            End
            Begin MwfCtl.MWCTL_MED atxt 
               Height          =   285
               Index           =   22
               Left            =   3645
               TabIndex        =   33
               ToolTipText     =   "Currency to be entered only for TcTyp = ‘RM"
               Top             =   4245
               Width           =   885
               _ExtentX        =   1561
               _ExtentY        =   503
               MaxLength       =   5
               DataType        =   4
               DataField       =   "UOCURCDTO"
               IdName          =   "UOCURCDTO"
            End
            Begin MwfCtl.MWCTL_MED atxt 
               Height          =   285
               Index           =   23
               Left            =   6120
               TabIndex        =   34
               ToolTipText     =   "Currency to be entered only for TcTyp = ‘RM"
               Top             =   4245
               Width           =   8805
               _ExtentX        =   15531
               _ExtentY        =   503
               MaxLength       =   255
               DataType        =   4
               DataField       =   "UOCURCDSEL"
               IdName          =   "UOCURCDSEL"
            End
            Begin MwfCtl.MWCTL_MED atxt 
               Height          =   285
               Index           =   76
               Left            =   3645
               TabIndex        =   12
               ToolTipText     =   "Enter To Parent Code"
               Top             =   2250
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
               Index           =   75
               Left            =   1215
               TabIndex        =   11
               ToolTipText     =   "Enter From Parent Code"
               Top             =   2250
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
               Index           =   77
               Left            =   6120
               TabIndex        =   13
               ToolTipText     =   "Enter Parent Code Selection"
               Top             =   2250
               Width           =   8790
               _ExtentX        =   15505
               _ExtentY        =   503
               MaxLength       =   255
               DataField       =   "UOPARENTCDSEL"
               IdName          =   "UOPARENTCDSEL"
            End
            Begin VB.Label LblRmPrntCd 
               AutoSize        =   -1  'True
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
               Height          =   240
               Left            =   60
               TabIndex        =   63
               Top             =   2250
               Width           =   945
            End
            Begin VB.Label LblCurrCd 
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
               Left            =   60
               TabIndex        =   62
               Top             =   4245
               Width           =   1005
            End
            Begin VB.Label ALBL 
               BackStyle       =   0  'Transparent
               Caption         =   "Loss Ctg"
               BeginProperty Font 
                  Name            =   "Arial"
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
               TabIndex        =   61
               Top             =   3675
               Width           =   1245
            End
            Begin VB.Label ALBL 
               BackStyle       =   0  'Transparent
               Caption         =   "STW Grp"
               BeginProperty Font 
                  Name            =   "Arial"
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
               TabIndex        =   60
               Top             =   3960
               Width           =   1245
            End
            Begin VB.Label ALBL 
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
               Index           =   12
               Left            =   60
               TabIndex        =   59
               Top             =   3390
               Width           =   1245
            End
            Begin VB.Label ALBL 
               BackStyle       =   0  'Transparent
               Caption         =   "Cust Ctg"
               BeginProperty Font 
                  Name            =   "Arial"
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
               TabIndex        =   58
               ToolTipText     =   "Location"
               Top             =   840
               Width           =   1245
            End
            Begin VB.Label LblRhYyyyMm 
               BackStyle       =   0  'Transparent
               Caption         =   "YyyyMm"
               BeginProperty Font 
                  Name            =   "Arial"
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
               TabIndex        =   57
               Top             =   1125
               Width           =   1245
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
               TabIndex        =   56
               Top             =   3105
               Width           =   1245
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
               TabIndex        =   55
               Top             =   2820
               Width           =   1245
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
               TabIndex        =   54
               ToolTipText     =   "Location"
               Top             =   1680
               Width           =   1245
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
               TabIndex        =   53
               Top             =   1965
               Width           =   1245
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
               TabIndex        =   52
               Top             =   2535
               Width           =   1245
            End
            Begin VB.Label ALBL 
               BackStyle       =   0  'Transparent
               Caption         =   "Tc Type"
               BeginProperty Font 
                  Name            =   "Arial"
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
               TabIndex        =   51
               Top             =   1395
               Width           =   1245
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
               Left            =   6000
               TabIndex        =   50
               Top             =   480
               Width           =   1215
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
               Left            =   3645
               TabIndex        =   49
               Top             =   480
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
               Left            =   1215
               TabIndex        =   48
               Top             =   480
               Width           =   1335
            End
         End
      End
   End
End
Attribute VB_Name = "EmrFrpRmRt"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim Rep As New EmrRepRmRt
Dim moCn As MwfLib.MDOConnection   '*** (Bef speed) bef 24/11/06
Dim ws_Opt As String
Private Sub SetGroupSort()
  'GRP_REP.Add "Design Ctg", "RrDmCtg", "", "", "", "", "", "", "(Select PDesc from Param where PTyp= 'DMCTG' and PMCd= RrDmCtg) "
  'GRP_REP.Add "Customer Code", "RrCmCd", "", "", "", "", "", "", "(Select CmName From CustMst Where CmCtg= 'C' and CmCd= RrCmCd) "
  'GRP_REP.Add "Lab Main Code", "RrLabMCd", "", "", "", "", "", "", "(Select PDesc from Param where PTyp= 'DMCTG' and PMCd= RrLabMCd) "
End Sub

Private Sub adc_setreprecsource()
  '*** Report Sql ***
  ' Zubin 213 (CurCd added)
  '*** Jay 3.01 *** [ILC] [RrXWtPer, RrFixMinTol, RrFixMaxTol added]
  'Select space(1) as qTag, space(20) as qHdRrSCtg, space(20) as qHdRrCd,
  'space(20) as qHdRrDmCtg, space(20) as qHdRrLabMCd, space(20) as qHdRrRngSz,
  'space(20) as qHdRrLsCtg, space(20) as qHdRrSTWGrp, space(20) as qHdRrFrLn,
  'space(20) as qHdRrToLn, space(20) as qHdRrCstRt, space(20) as qHdRrSalRt,
  'space(20) as qHdRrQw, space(20) as qHdRrQwCst, space(20) as qHdRrCmCurCd,
  'space(20) as qHdRrXWtPer, space(20) as qHdRrFixMinTol, space(20) as qHdRrFixMaxTol,
  'RrTcTyp, RrCtg, RrCmCd, RrSCtg, RrCd, RrDmCtg, RrLabMCd, RrRngSz, RrLsCtg,
  'RrSTWGrp, RrFrLn, RrToLn, RrCstRt, RrSalRt, space(1) as qRrQw, space(1) as qRrQwCst,
  'RrCmCurCd, RrXWtPer, RrFixMinTol, RrFixMaxTol from RmRt where 1= 2
  '*** Jay 3.01 *** [ILC] [RrXWtPer, RrFixMinTol, RrFixMaxTol added]
  '*** Report Sql ***

  Dim wrepcnd As String, wCnd As String, wsql1 As String, wsql2 As String
  Dim wsql3 As String, wSqlStrg As String, wHdCnd As String, wHdJoin As String
  Set Rep = Nothing
  Set adc.RepSource = Rep
  Call DispCoNm
  Call SetMwName(Rep)   ' ***** Manali 3.6.1 - Microway footer
  'Call GRP_REP.Gen3LRep
  'grpflds = GRP_REP.GrpFldLst
  
  ws_Opt = IIF(adc("UoYyMmFr") = 0, "RepRmRt", "RepRmRtHist")
  
  If UCase(ws_Opt) = UCase("RepRmRt") Then
    adc("UoCmCtgFr").CmpStr = "RrCmCtg = "
    adc("UoYyMmFr").CmpStr = " "
    adc("UoPmCdFr").CmpStr = "RrTcTyp >= ": adc("UoPmCdTo").CmpStr = "RrTcTyp <= ": adc("UoPmCdSel").CmpStr = "RrTcTyp In "
    adc("UoRmCtgFr").CmpStr = "RrCtg >= ": adc("UoRmCtgTo").CmpStr = "RrCtg <= ": adc("UoRmCtgSel").CmpStr = "RrCtg In "
    adc("UoRmSCtgFr").CmpStr = "RrSCtg >= ": adc("UoRmSCtgTo").CmpStr = "RrSCtg <= ": adc("UoRmSCtgSel").CmpStr = "RrSCtg In "
    adc("UoRmCdFr").CmpStr = "RrCd >= ": adc("UoRmCdTo").CmpStr = "RrCd <= ": adc("UoRmCdSel").CmpStr = "RrCd In "
    adc("UoDmCtgFr").CmpStr = "RrDmCtg >= ": adc("UoDmCtgTo").CmpStr = "RrDmCtg <= ": adc("UoDmCtgSel").CmpStr = "RrDmCtg In "
    adc("UoCmCdFr").CmpStr = "RrCmCd >= ": adc("UoCmCdTo").CmpStr = "RrCmCd <= ": adc("UoCmCdSel").CmpStr = "RrCmCd In "
    adc("UoRngSzFr").CmpStr = "RrRngSz >= ": adc("UoRngSzTo").CmpStr = "RrRngSz <= ": adc("UoRngSzSel").CmpStr = "RrRngSz In "
    adc("UoLsCtgFr").CmpStr = "RrLsCtg >= ": adc("UoLsCtgTo").CmpStr = "RrLsCtg <= ": adc("UoLsCtgSel").CmpStr = "RrLsCtg In "
    adc("UoSTWGrpFr").CmpStr = "RrSTWGrp >= ": adc("UoSTWGrpTo").CmpStr = "RrSTWGrp <= ": adc("UoSTWGrpSel").CmpStr = "RrSTWGrp In "
    'Uni.20- Currency code added
    adc("UoCurCdFr").CmpStr = "RrCmCurCd >= ": adc("UoCurCdTo").CmpStr = "RrCmCurCd <= ": adc("UoCurCdSel").CmpStr = "RrCmCurCd In "
    'VK.28-Parent code scope added
    adc("UoParentCdFr").CmpStr = "isnull(RmParentCd,'') >=": adc("UoParentCdTo").CmpStr = "isnull(RmParentCd,'') <=": adc("UoParentCdSel").CmpStr = "isnull(RmParentCd,'') In":
    wrepcnd = adc.RepCond
    wCnd = IIF(wrepcnd <> "", " where ", "") + wrepcnd
    'vk.28 set values for wHdCnd ,wHdJoin.  to avoid blank page in output
    wHdCnd = IIF(wrepcnd <> "", " and " + wrepcnd, "")
    If adc("UoParentCdFr") <> "" Or adc("UoParentCdTo") <> "" Or adc("UoParentCdSel") <> "" Then _
      wHdJoin = " left outer join RmMst on RmCd=RrCd "
      
    ' Zubin 213 (CurCd added)
    '*** Jay 3.01 *** [ILC] [RrXWtPer, RrFixMinTol, RrFixMaxTol added]
    wsql1 = "Select 'R' as qTag, " + _
            "'' as qHdRrSCtg, '' as qHdRrCd, '' as qHdRrDmCtg, '' as qHdRrLabMCd, " + _
            "'' as qHdRrRngSz, '' as qHdRrLsCtg, '' as qHdRrSTWGrp, '' as qHdRrFrLn, " + _
            "'' as qHdRrToLn, '' as qHdRrCstRt, '' as qHdRrSalRt, '' as qHdRrQw, '' as qHdRrQwCst, '' as qHdRrCmCurCd, " + _
            "'' as qHdRrXWtPer, '' as qHdRrFixMinTol, '' as qHdRrFixMaxTol, " + _
            "Upper(RrTcTyp) as RrTcTyp, Upper(RrCtg) as RrCtg, RrCmCd, RrSCtg, RrCd, " + _
            "RrDmCtg, RrLabMCd, RrRngSz, RrLsCtg, RrSTWGrp, RrFrLn, RrToLn, RrCstRt, " + _
            "RrSalRt, " + _
            "Isnull((Select RmQw from RmMst where RmCd= RrCd " + IIF(gs_Partition = ctCurrPrtn, " and RmPrtKey=RrPrtKey ", "") + "), '') as qRrQw, " + _
            "Isnull((Select RmQwCst from RmMst where RmCd= RrCd " + IIF(gs_Partition = ctCurrPrtn, " and RmPrtKey=RrPrtKey ", "") + "),'') as qRrQwCst, RrCmCurCd, " + _
            "RrXWtPer, RrFixMinTol, RrFixMaxTol " + _
            "From RmRt left outer join RmMst on RmCd=RrCd " + wCnd + IIF(gs_Partition = ctCurrPrtn, IIF(wCnd <> "", " and ", " Where ") + " RrPrtKey='" + ctCurrPrtn + "' ", "")
             
    Rep.TxtHead.SetText "Raw Material Rate Listing"
    Rep.TxtYyyyMm.SetText ""
    Rep.TxtYyyyMmLbl.Suppress = True
  ElseIf UCase(ws_Opt) = UCase("RepRmRtHist") Then
    adc("UoCmCtgFr").CmpStr = "RhCmCtg = "
    adc("UoYyMmFr").CmpStr = "RhYyyyMm = "
    adc("UoPmCdFr").CmpStr = "RhTcTyp >= ": adc("UoPmCdTo").CmpStr = "RhTcTyp <= ": adc("UoPmCdSel").CmpStr = "RhTcTyp In "
    adc("UoRmCtgFr").CmpStr = "RhCtg >= ": adc("UoRmCtgTo").CmpStr = "RhCtg <= ": adc("UoRmCtgSel").CmpStr = "RhCtg In "
    adc("UoRmSCtgFr").CmpStr = "RhSCtg >= ": adc("UoRmSCtgTo").CmpStr = "RhSCtg <= ": adc("UoRmSCtgSel").CmpStr = "RhSCtg In "
    adc("UoRmCdFr").CmpStr = "RhCd >= ": adc("UoRmCdTo").CmpStr = "RhCd <= ": adc("UoRmCdSel").CmpStr = "RhCd In "
    adc("UoDmCtgFr").CmpStr = "RhDmCtg >= ": adc("UoDmCtgTo").CmpStr = "RhDmCtg <= ": adc("UoDmCtgSel").CmpStr = "RhDmCtg In "
    adc("UoCmCdFr").CmpStr = "RhCmCd >= ": adc("UoCmCdTo").CmpStr = "RhCmCd <= ": adc("UoCmCdSel").CmpStr = "RhCmCd In "
    adc("UoRngSzFr").CmpStr = "RhRngSz >= ": adc("UoRngSzTo").CmpStr = "RhRngSz <= ": adc("UoRngSzSel").CmpStr = "RhRngSz In "
    adc("UoLsCtgFr").CmpStr = "RhLsCtg >= ": adc("UoLsCtgTo").CmpStr = "RhLsCtg <= ": adc("UoLsCtgSel").CmpStr = "RhLsCtg In "
    adc("UoSTWGrpFr").CmpStr = "RhSTWGrp >= ": adc("UoSTWGrpTo").CmpStr = "RhSTWGrp <= ": adc("UoSTWGrpSel").CmpStr = "RhSTWGrp In "
    'VK.28-Parent code scope added
    adc("UoParentCdFr").CmpStr = "isnull(RmParentCd,'') >=": adc("UoParentCdTo").CmpStr = "isnull(RmParentCd,'') <=": adc("UoParentCdSel").CmpStr = "isnull(RmParentCd,'') In":
  
    wrepcnd = adc.RepCond
    wCnd = IIF(wrepcnd <> "", " where ", "") + wrepcnd
    'vk.28 set values for wHdCnd ,wHdJoin.  to avoid blank page in output
    wHdCnd = IIF(wrepcnd <> "", " and " + wrepcnd, "")
    If adc("UoParentCdFr") <> "" Or adc("UoParentCdTo") <> "" Or adc("UoParentCdSel") <> "" Then _
      wHdJoin = " left outer join RmMst on RmCd=RhCd "
        
    ' Zubin 213 (CurCd added)
    '*** Jay 3.01 *** [ILC] [RrXWtPer, RrFixMinTol, RrFixMaxTol added]
    wsql1 = "Select 'R' as qTag, " + _
            "'' as qHdRrSCtg, '' as qHdRrCd, '' as qHdRrDmCtg, '' as qHdRrLabMCd, " + _
            "'' as qHdRrRngSz, '' as qHdRrLsCtg, '' as qHdRrSTWGrp, '' as qHdRrFrLn, " + _
            "'' as qHdRrToLn, '' as qHdRrCstRt, '' as qHdRrSalRt, '' as qHdRrQw, '' as qHdRrQwCst, '' as qHdRrCmCurCd, " + _
            "'' as qHdRrXWtPer, '' as qHdRrFixMinTol, '' as qHdRrFixMaxTol, " + _
            "Upper(RhTcTyp) as RrTcTyp, Upper(RhCtg) as RrCtg, RhCmCd as RrCmCd, " + _
            "RhSCtg as RrSCtg, RhCd as RrCd, RhDmCtg as RrDmCtg, RhLabMCd as RrLabMCd, " + _
            "RhRngSz as RrRngSz, RhLsCtg as RrLsCtg, RhSTWGrp as RrSTWGrp, " + _
            "RhFrLn as RrFrLn, RhToLn as RrToLn, RhCstRt as RrCstRt, RhSalRt as RrSalRt, " + _
            "(Select RmQw from RmMst where RmCd= RhCd " + IIF(gs_Partition = ctCurrPrtn, " and RmPrtKey=RhPrtKey ", "") + " ) as qRrQw, " + _
            "(Select RmQwCst from RmMst where RmCd= RhCd" + IIF(gs_Partition = ctCurrPrtn, " and RmPrtKey=RhPrtKey ", "") + ") as qRrQwCst, RhCmCurCd as RrCmCurCd, " + _
            "RhXWtPer as RrXWtPer, RhFixMinTol as RrFixMinTol, RhFixMaxTol as RrFixMaxTol " + _
            "From RmRtHist left outer join RmMst on RmCd=RhCd " + wCnd + IIF(gs_Partition = ctCurrPrtn, IIF(wCnd <> "", " and ", " Where ") + " RhPrtKey='" + ctCurrPrtn + "' ", "")
    Rep.TxtHead.SetText "Raw Material Rate History"
    Rep.TxtYyyyMm.SetText adc("UoYyMmFr")
    Rep.TxtYyyyMmLbl.Suppress = False
  End If
  
  adc("UoCmCtgFr").CmpStr = "": adc("UoYyMmFr").CmpStr = ""
  adc("UoPMCdFr").CmpStr = "TsTyp>= ": adc("UoPMCdTo").CmpStr = "TsTyp<= ": adc("UoPMCdSel").CmpStr = "TsTyp In "
  adc("UoRmCtgFr").CmpStr = "TsSTyp>= ": adc("UoRmCtgTo").CmpStr = "TsSTyp<= ": adc("UoRmCtgSel").CmpStr = "TsSTyp In "
  adc("UoRmSCtgFr").CmpStr = "": adc("UoRmSCtgTo").CmpStr = "": adc("UoRmSCtgSel").CmpStr = ""
  adc("UoRmCdFr").CmpStr = "": adc("UoRmCdTo").CmpStr = "": adc("UoRmCdSel").CmpStr = ""
  adc("UoDmCtgFr").CmpStr = "": adc("UoDmCtgTo").CmpStr = "": adc("UoDmCtgSel").CmpStr = ""
  adc("UoCmCdFr").CmpStr = "": adc("UoCmCdTo").CmpStr = "": adc("UoCmCdSel").CmpStr = ""
  adc("UoRngSzFr").CmpStr = "": adc("UoRngSzTo").CmpStr = "": adc("UoRngSzSel").CmpStr = ""
  adc("UoLsCtgFr").CmpStr = "": adc("UoLsCtgTo").CmpStr = "": adc("UoLsCtgSel").CmpStr = ""
  adc("UoSTWGrpFr").CmpStr = "": adc("UoSTWGrpTo").CmpStr = "": adc("UoSTWGrpSel").CmpStr = ""
  'Uni.20-Currency code added
  adc("UoCurCdFr").CmpStr = "": adc("UoCurCdTo").CmpStr = "": adc("UoCurCdSel").CmpStr = ""
  'VK.28-Parent code scope added
  adc("UoParentCdFr").CmpStr = "": adc("UoParentCdTo").CmpStr = "": adc("UoParentCdSel").CmpStr = "":
  wrepcnd = adc.RepCond
  wCnd = IIF(wrepcnd <> "", " and ", "") + wrepcnd
  
  
  'RrRngSz, RrLsCtg, RrSTWGrp,
  wsql2 = " Union All " + _
         "Select 'H' as qTag, " + _
         "(Select RTrim(TsLbl) From TcStr where TsTbl= 'RmRt' and TsTyp= TS.TsTyp and " + _
         "TsSTyp= TS.TsSTyp and TsIdName= 'RrSCtg') as qHdRrSCtg, " + _
         "(Select RTrim(TsLbl) From TcStr where TsTbl= 'RmRt' and TsTyp= TS.TsTyp and " + _
         "TsSTyp= TS.TsSTyp and TsIdName= 'RrCd') as qHdRrCd, " + _
         "(Select RTrim(TsLbl) From TcStr where TsTbl= 'RmRt' and TsTyp= TS.TsTyp and " + _
         "TsSTyp= TS.TsSTyp and TsIdName= 'RrDmCtg') as qHdRrDmCtg, " + _
         "(Select RTrim(TsLbl) From TcStr where TsTbl= 'RmRt' and TsTyp= TS.TsTyp and " + _
         "TsSTyp= TS.TsSTyp and TsIdName= 'RrLabMCd') as qHdRrLabMCd, " + _
         "(Select RTrim(TsLbl) From TcStr where TsTbl= 'RmRt' and TsTyp= TS.TsTyp and " + _
         "TsSTyp= TS.TsSTyp and TsIdName= 'RrRngSz') as qHdRrRngSz, " + _
         "(Select RTrim(TsLbl) From TcStr where TsTbl= 'RmRt' and TsTyp= TS.TsTyp and " + _
         "TsSTyp= TS.TsSTyp and TsIdName= 'RrLsCtg') as qHdRrLsCtg, " + _
         "(Select RTrim(TsLbl) From TcStr where TsTbl= 'RmRt' and TsTyp= TS.TsTyp and " + _
         "TsSTyp= TS.TsSTyp and TsIdName= 'RrSTWGrp') as qHdRrSTWGrp, "
  ' Zubin 213 (CurCd added)
  '*** Jay 3.01 *** [ILC] [RrXWtPer, RrFixMinTol, RrFixMaxTol added]
  wsql2 = wsql2 + "(Select RTrim(TsLbl) From TcStr where TsTbl= 'RmRt' and TsTyp= TS.TsTyp and " + _
         "TsSTyp= TS.TsSTyp and TsIdName= 'RrFrLn') as qHdRrFrLn, " + _
         "(Select RTrim(TsLbl) From TcStr where TsTbl= 'RmRt' and TsTyp= TS.TsTyp and " + _
         "TsSTyp= TS.TsSTyp and TsIdName= 'RrToLn') as qHdRrToLn, " + _
         "(Select RTrim(TsLbl) From TcStr where TsTbl= 'RmRt' and TsTyp= TS.TsTyp and " + _
         "TsSTyp= TS.TsSTyp and TsIdName= 'RrCstRt') as qHdRrCstRt, " + _
         "(Select RTrim(TsLbl) From TcStr where TsTbl= 'RmRt' and TsTyp= TS.TsTyp and " + _
         "TsSTyp= TS.TsSTyp and TsIdName= 'RrSalRt') as qHdRrSalRt, " + _
         "(Select RTrim(TsLbl) From TcStr where TsTbl= 'RmRt' and TsTyp= TS.TsTyp and " + _
         "TsSTyp= TS.TsSTyp and TsIdName= 'wRrQw') as qHdRrQw, " + _
         "(Select RTrim(TsLbl) From TcStr where TsTbl= 'RmRt' and TsTyp= TS.TsTyp and " + _
         "TsSTyp= TS.TsSTyp and TsIdName= 'wRrQwCst') as qHdRrQwCst, " + _
         "(Select RTrim(TsLbl) From TcStr where TsTbl= 'RmRt' and TsTyp= TS.TsTyp and " + _
         "TsSTyp= TS.TsSTyp and TsIdName= 'RrCmCurCd') as qHdRrCmCurCd, " + _
         "(Select RTrim(TsLbl) From TcStr where TsTbl= 'RmRt' and TsTyp= TS.TsTyp and " + _
         "TsSTyp= TS.TsSTyp and TsIdName= 'RrXWtPer') as qHdRrXWtPer, " + _
         "(Select RTrim(TsLbl) From TcStr where TsTbl= 'RmRt' and TsTyp= TS.TsTyp and " + _
         "TsSTyp= TS.TsSTyp and TsIdName= 'RrFixMinTol') as qHdRrFixMinTol, " + _
         "(Select RTrim(TsLbl) From TcStr where TsTbl= 'RmRt' and TsTyp= TS.TsTyp and " + _
         "TsSTyp= TS.TsSTyp and TsIdName= 'RrFixMaxTol') as qHdRrFixMaxTol, " + _
         "Upper(TsTyp) as RrTcTyp, Upper(TsSTyp) as RrCtg, '' as RrCmCd, '' as RrSctg, '' as RrCd, " + _
         "'' as RrDmCtg, '' as RrLabMCd, '' as RrRngSz, '' as RrLsCtg, '' as RrSTWGrp, " + _
         "0 as RrFrLn, 0 as RrToLn, 0 as RrCstRt, 0 as RrSalRt, '' as qRrQw, '' as qRrQwCst, '' as RrCmCurCd, " + _
         "'' as RrXWtPer, '' as RrFixMinTol, '' as RrFixMaxTol " + _
         "from TcStr TS where TsTbl= 'RmRt' and "
        
' urmi 2.04-3 One Clause for custCtg has been added RrCmCtg= UoCmCtg in both the cases
' to avoid the blank page coming for p Customers when there is no rec in rmRt for them
  'vk.28 wHdCnd added to RmRt table to get link with TcStr table and get headings,
  ' to avoid blank page in output
  If UCase(ws_Opt) = UCase("RepRmRt") Then
    wsql3 = " Exists(Select 'x' from RmRt " + wHdJoin + " where " + IIF(gs_Partition = ctCurrPrtn, " RrPrtKey='" + ctCurrPrtn + "' and ", "") + " RrTcTyp= TsTyp and RrCtg= TsSTyp and " + _
            " RrCmCtg= '" + adc("UoCmCtgFr") + "' " + wHdCnd + " ) " + wCnd + _
            " group by TsTbl, TsTyp, TsSTyp "
  ElseIf UCase(ws_Opt) = UCase("RepRmRtHist") Then
    wsql3 = " Exists(Select 'x' from RmRtHist " + wHdJoin + " where " + IIF(gs_Partition = ctCurrPrtn, " RhPrtKey='" + ctCurrPrtn + "' and ", "") + " RhYyyyMm= '" + CStr(adc("UoYyMmFr")) + "' and " + _
            " RhTcTyp= TsTyp and RhCtg= TsSTyp and RhCmCtg= '" + adc("UoCmCtgFr") + "' " + wHdCnd + " ) " + wCnd + _
            " group by TsTbl, TsTyp, TsSTyp "
  End If
  'Debug.Print wsql1 + wsql2 + wsql3
  wSqlStrg = wsql1 + wsql2 + wsql3
  adc.RepRecSource = wSqlStrg + ctMaxDopOpt '****** Sachin 3.01 added ctMaxDopOpt
  
  If UCase(adc("UoCmCtgFr")) = "C" Then
    Rep.TxtHead1.SetText "(Customer)"
  ElseIf UCase(adc("UoCmCtgFr")) = "P" Then
    Rep.TxtHead1.SetText "(Price List)"
  ElseIf UCase(adc("UoCmCtgFr")) = "S" Then
    Rep.TxtHead1.SetText "(Supplier)"
  ElseIf UCase(adc("UoCmCtgFr")) = "T" Then
    Rep.TxtHead1.SetText "(Contractor)"
  End If
  '578.19 calling function to change Rmcode font size
  GRP_REP.SetFormula Rep, "wRmCdFontSz", GetRmcdFontSz("{rdo.RrCd}", 8)
  
  CRV_REP.DisplayGroupTree = False
  
  Call DispReport(adc) '****** Sachin 3.02 - Display Report ******
  
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
    Set moCn = adc.Connection    '*** (Bef speed) bef 24/11/06
    Set adc.FirNKeyCtl = adc("UoCmCtgFr")
    Call InitProp(Me)
    '*** Parameters Used For Setting Properties Of UsrOpt Fields ***
    gs_PTyp = "RRTCTYP"
    gs_Tbl = "Param"
    gb_ShowYyMmHlp = True     '*** Emr2-06
    '*** Parameters Used For Setting Properties Of UsrOpt Fields ***
    
    '*** Emr2-06
    'If UCase(ADC.MenuCd) = UCase("RepRmRt") Then
    '  ADC("UoYyMmFr").Visible = False: ADC("UoYyMmFr").Enabled = False
    '  LblRhYyyyMm.Visible = False
    'ElseIf UCase(ADC.MenuCd) = UCase("RepRmRtHist") Then
    '  ADC("UoYyMmFr").Visible = True: ADC("UoYyMmFr").Enabled = True
    '  LblRhYyyyMm.Visible = True
    '  gb_ShowYyMmHlp = True
    'End If
    '*** Emr2-06
    
    
    Call SetGroupSort
End Sub
Private Sub ADC_NonKeyWhen(pr_Cancel As Boolean, pr_ErrMsg As String)
  gs_CmCtg = adc("UoCmCtgFr")
End Sub
Private Sub ATXT_FldChange(Index As Integer, ByVal IdName As String, Cancel As Boolean, ErrMsg As String, ByVal pv_NewValue As Variant)
  gs_Tbl = "Param"
  
  Select Case UCase(IdName)
  Case Is = UCase("UoCmCtgFr")
    gs_CmCtg = pv_NewValue
  End Select
  
  Call SetProp(Me, IdName, Valid, Cancel, ErrMsg)
End Sub
Private Sub ATXT_FldWhen(Index As Integer, ByVal IdName As String, Cancel As Boolean, ErrMsg As String)
  gs_Tbl = "Param"
  SetProp Me, IdName, When
End Sub

Private Sub tab_REP_Click(PreviousTab As Integer)
    fra_tabrep(TAB_REP.Tab).Enabled = True
    fra_tabrep(PreviousTab).Enabled = False
End Sub
Private Sub ADC_KeyWhen()
    adc("UOUSRCD") = adc.UsrCd
    adc("UOMNUCD") = adc.MenuCd
    adc("UoYyMmFr") = 0
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

