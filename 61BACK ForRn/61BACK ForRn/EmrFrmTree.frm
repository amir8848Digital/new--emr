VERSION 5.00
Object = "{AB2CC32A-E652-42A2-B341-26ECE02F6386}#2.1#0"; "MwfCtlforspeed8.ocx"
Begin VB.Form EmrFrmTree 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Tree  Entry"
   ClientHeight    =   9810
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   13200
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   9810
   ScaleWidth      =   13200
   WindowState     =   2  'Maximized
   Begin MwfCtl.MWCTL_BTN1 CmdShowTot 
      Height          =   495
      Left            =   7200
      TabIndex        =   122
      ToolTipText     =   "Show Totals"
      Top             =   9270
      Width           =   735
      _ExtentX        =   1296
      _ExtentY        =   873
      BackColor       =   -2147483633
      MaskColor       =   8454143
      Caption         =   "S&how Totals"
      BeginProperty LblFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin MwfCtl.MWCTL_BTN1 CmdCB 
      Height          =   495
      Left            =   6465
      TabIndex        =   28
      ToolTipText     =   "Copy Bags From One Loc To Another "
      Top             =   9270
      Width           =   735
      _ExtentX        =   1296
      _ExtentY        =   873
      BackColor       =   -2147483633
      MaskColor       =   8454143
      Caption         =   "C&opy Bags"
      BeginProperty LblFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin MwfCtl.MWCTL_HLP MWCTL_HLP1 
      Height          =   405
      Left            =   8100
      TabIndex        =   29
      TabStop         =   0   'False
      Top             =   9300
      Visible         =   0   'False
      Width           =   555
      _ExtentX        =   979
      _ExtentY        =   714
   End
   Begin MwfCtl.MWCTL_ADC ADC 
      CausesValidation=   0   'False
      Height          =   555
      Left            =   60
      TabIndex        =   30
      Top             =   9255
      Width           =   11865
      _ExtentX        =   20929
      _ExtentY        =   979
      BtnWidth        =   885
      BtnHeight       =   435
      BtnMaskColor    =   8454143
      FontSz          =   8.25
   End
   Begin VB.Frame FRA_KEYS 
      BorderStyle     =   0  'None
      Height          =   280
      Left            =   0
      TabIndex        =   33
      Top             =   30
      Width           =   4020
      Begin MwfCtl.MWCTL_MED ATXT 
         Height          =   285
         Index           =   4
         Left            =   2925
         TabIndex        =   2
         ToolTipText     =   "Enter Voucher Number"
         Top             =   0
         Width           =   930
         _ExtentX        =   1640
         _ExtentY        =   503
         Alignment       =   2
         Mask            =   "######0"
         MaxLength       =   7
         DataType        =   1
         DataField       =   "TrNo"
         IdName          =   "TRNO"
      End
      Begin MwfCtl.MWCTL_MED ATXT 
         Height          =   285
         Index           =   6
         Left            =   1260
         TabIndex        =   0
         ToolTipText     =   "Enter Voucher YyMmdd"
         Top             =   0
         Width           =   855
         _ExtentX        =   1508
         _ExtentY        =   503
         MaxLength       =   6
         DataType        =   4
         DataField       =   "TrYmd"
         IdName          =   "TRYMD"
      End
      Begin MwfCtl.MWCTL_MED ATXT 
         Height          =   285
         Index           =   94
         Left            =   3870
         TabIndex        =   41
         Top             =   0
         Visible         =   0   'False
         Width           =   135
         _ExtentX        =   238
         _ExtentY        =   503
         Enabled         =   0   'False
         MaxLength       =   3
         DataType        =   4
         DataField       =   "TrCoCd"
         IdName          =   "TRCOCD"
      End
      Begin MwfCtl.MWCTL_MED ATXT 
         Height          =   285
         Index           =   1
         Left            =   2220
         TabIndex        =   1
         ToolTipText     =   "Enter Voucher Year"
         Top             =   0
         Width           =   570
         _ExtentX        =   1005
         _ExtentY        =   503
         MaxLength       =   3
         DataType        =   4
         DataField       =   "TrChr"
         IdName          =   "TRCHR"
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
         Height          =   285
         Index           =   14
         Left            =   1710
         TabIndex        =   44
         Top             =   0
         Width           =   135
      End
      Begin VB.Label Label3 
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
         Height          =   285
         Left            =   2130
         TabIndex        =   39
         Top             =   0
         Width           =   135
      End
      Begin VB.Label Label2 
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
         Height          =   285
         Left            =   2835
         TabIndex        =   38
         Top             =   0
         Width           =   135
      End
      Begin VB.Label ALBL 
         BackStyle       =   0  'Transparent
         Caption         =   "Voucher"
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
         Left            =   60
         TabIndex        =   34
         Top             =   0
         Width           =   1065
      End
   End
   Begin VB.Frame FRA_NKEYS 
      BorderStyle     =   0  'None
      Enabled         =   0   'False
      Height          =   9570
      Left            =   -120
      TabIndex        =   32
      Top             =   0
      Width           =   13500
      Begin VB.Frame FraNKeyAll 
         BorderStyle     =   0  'None
         Height          =   9375
         Left            =   180
         TabIndex        =   35
         Top             =   30
         Width           =   13065
         Begin MwfCtl.MWCTL_MED ATXT 
            Height          =   285
            Index           =   30
            Left            =   4605
            TabIndex        =   3
            ToolTipText     =   "Enter Tree Date"
            Top             =   15
            Width           =   855
            _ExtentX        =   1508
            _ExtentY        =   503
            MaxLength       =   8
            DataType        =   3
            DataField       =   "TrDt"
            IdName          =   "TRDT"
         End
         Begin MwfCtl.MWCTL_MED ATXT 
            Height          =   285
            Index           =   0
            Left            =   1200
            TabIndex        =   6
            ToolTipText     =   "Enter Wax Specific Gravity"
            Top             =   285
            Width           =   1035
            _ExtentX        =   1826
            _ExtentY        =   503
            Alignment       =   2
            Mask            =   "##0.00"
            MaxLength       =   6
            DataType        =   2
            DataField       =   "TrWaxSpGr"
            IdName          =   "TRWAXSPGR"
         End
         Begin MwfCtl.MWCTL_MED ATXT 
            Height          =   285
            Index           =   25
            Left            =   615
            TabIndex        =   27
            ToolTipText     =   "Enter Raw Material Description"
            Top             =   8820
            Width           =   11250
            _ExtentX        =   19844
            _ExtentY        =   503
            MaxLength       =   60
            DataField       =   "TrDesc"
            IdName          =   "TRDESC"
         End
         Begin MwfCtl.MWCTL_MED ATXT 
            Height          =   285
            Index           =   2
            Left            =   6435
            TabIndex        =   4
            ToolTipText     =   "Enter Tree Color"
            Top             =   15
            Width           =   840
            _ExtentX        =   1482
            _ExtentY        =   503
            MaxLength       =   5
            DataType        =   4
            DataField       =   "TrCol"
            IdName          =   "TRCOL"
         End
         Begin MwfCtl.MWCTL_MED ATXT 
            Height          =   285
            Index           =   99
            Left            =   8190
            TabIndex        =   5
            ToolTipText     =   "Enter Tree Kt"
            Top             =   15
            Width           =   2730
            _ExtentX        =   4815
            _ExtentY        =   503
            MaxLength       =   16
            DataType        =   4
            DataField       =   "TrRmcd"
            IdName          =   "TRRMCD"
         End
         Begin MwfCtl.MWCTL_MED ATXT 
            Height          =   285
            Index           =   3
            Left            =   3600
            TabIndex        =   7
            ToolTipText     =   "Enter Wax Gross Wt"
            Top             =   300
            Width           =   1035
            _ExtentX        =   1826
            _ExtentY        =   503
            Alignment       =   2
            Mask            =   "###0.000"
            MaxLength       =   8
            DataType        =   2
            DataField       =   "TrGrsWaxWt"
            IdName          =   "TRGRSWAXWT"
         End
         Begin MwfCtl.MWCTL_MED ATXT 
            Height          =   285
            Index           =   7
            Left            =   9880
            TabIndex        =   9
            ToolTipText     =   "Enter Metal additional Wt"
            Top             =   300
            Width           =   1035
            _ExtentX        =   1826
            _ExtentY        =   503
            Alignment       =   2
            Mask            =   "-###0.000"
            MaxLength       =   9
            DataType        =   2
            DataField       =   "TrAddMetWt"
            IdName          =   "TRADDMETWT"
         End
         Begin VB.Frame FraCast 
            Height          =   1395
            Left            =   0
            TabIndex        =   47
            Top             =   7410
            Width           =   11835
            Begin MwfCtl.MWCTL_MED ATXT 
               Height          =   285
               Index           =   17
               Left            =   3675
               TabIndex        =   13
               ToolTipText     =   "Enter Received Metal Wt"
               Top             =   990
               Width           =   915
               _ExtentX        =   1614
               _ExtentY        =   503
               Alignment       =   2
               Mask            =   "###0.000"
               MaxLength       =   8
               DataType        =   2
               DataField       =   "TrRecdMetWt"
               ReCalcParent    =   "WTRDIFF"
               IdName          =   "TRRECDMETWT"
            End
            Begin MwfCtl.MWCTL_MED ATXT 
               Height          =   285
               Index           =   9
               Left            =   1035
               TabIndex        =   10
               ToolTipText     =   "Enter Casting Date "
               Top             =   420
               Width           =   855
               _ExtentX        =   1508
               _ExtentY        =   503
               MaxLength       =   8
               DataType        =   3
               DataField       =   "TrCstDt"
               IdName          =   "TRCSTDT"
            End
            Begin MwfCtl.MWCTL_MED ATXT 
               Height          =   285
               Index           =   10
               Left            =   3675
               TabIndex        =   11
               ToolTipText     =   "Enter Fresh Metal Weight"
               Top             =   420
               Width           =   915
               _ExtentX        =   1614
               _ExtentY        =   503
               Alignment       =   2
               Mask            =   "###0.000"
               MaxLength       =   8
               DataType        =   2
               DataField       =   "TrFrMetWt"
               ReCalcParent    =   "WTRDIFF"
               IdName          =   "TRFRMETWT"
            End
            Begin MwfCtl.MWCTL_MED ATXT 
               Height          =   285
               Index           =   11
               Left            =   3675
               TabIndex        =   12
               ToolTipText     =   "Enter ReUsed Metal Weight"
               Top             =   705
               Width           =   915
               _ExtentX        =   1614
               _ExtentY        =   503
               Alignment       =   2
               Mask            =   "###0.000"
               MaxLength       =   8
               DataType        =   2
               DataField       =   "TrUsdMetWt"
               ReCalcParent    =   "WTRDIFF"
               IdName          =   "TRUSDMETWT"
            End
            Begin MwfCtl.MWCTL_MED ATXT 
               Height          =   285
               Index           =   8
               Left            =   6075
               TabIndex        =   16
               ToolTipText     =   "Enter Scrap"
               Top             =   990
               Width           =   915
               _ExtentX        =   1614
               _ExtentY        =   503
               Alignment       =   2
               Mask            =   "###0.000"
               MaxLength       =   8
               DataType        =   2
               DataField       =   "TrScrap"
               ReCalcParent    =   "WTRDIFF"
               IdName          =   "TRSCRAP"
            End
            Begin MwfCtl.MWCTL_MED ATXT 
               Height          =   285
               Index           =   12
               Left            =   6075
               TabIndex        =   14
               ToolTipText     =   "Enter Casting Loss"
               Top             =   420
               Width           =   915
               _ExtentX        =   1614
               _ExtentY        =   503
               Alignment       =   2
               Mask            =   "###0.000"
               MaxLength       =   8
               DataType        =   2
               DataField       =   "TrCastLoss"
               ReCalcParent    =   "WTRDIFF"
               IdName          =   "TRCASTLOSS"
            End
            Begin MwfCtl.MWCTL_MED ATXT 
               Height          =   285
               Index           =   13
               Left            =   6075
               TabIndex        =   15
               ToolTipText     =   "EnterCutting Loss"
               Top             =   705
               Width           =   915
               _ExtentX        =   1614
               _ExtentY        =   503
               Alignment       =   2
               Mask            =   "###0.000"
               MaxLength       =   8
               DataType        =   2
               DataField       =   "TrCutLoss"
               ReCalcParent    =   "WTRDIFF"
               IdName          =   "TRCUTLOSS"
            End
            Begin MwfCtl.MWCTL_MED ATXT 
               Height          =   285
               Index           =   14
               Left            =   8595
               TabIndex        =   19
               ToolTipText     =   "Enter Peices Gross Weight"
               Top             =   990
               Width           =   915
               _ExtentX        =   1614
               _ExtentY        =   503
               Alignment       =   2
               Mask            =   "###0.000"
               MaxLength       =   8
               DataType        =   2
               DataField       =   "TrGrsPcWt"
               ReCalcParent    =   "TRNETPCWT"
               IdName          =   "TRGRSPCWT"
            End
            Begin MwfCtl.MWCTL_MED ATXT 
               Height          =   285
               Index           =   15
               Left            =   8595
               TabIndex        =   17
               ToolTipText     =   "Enter Grinding Loss"
               Top             =   420
               Width           =   915
               _ExtentX        =   1614
               _ExtentY        =   503
               Alignment       =   2
               Mask            =   "###0.000"
               MaxLength       =   8
               DataType        =   2
               DataField       =   "TrGrindLoss"
               ReCalcParent    =   "WTRDIFF"
               IdName          =   "TRGRINDLOSS"
            End
            Begin MwfCtl.MWCTL_MED ATXT 
               Height          =   285
               Index           =   16
               Left            =   8595
               TabIndex        =   18
               ToolTipText     =   "EnterStrip Loss"
               Top             =   705
               Width           =   915
               _ExtentX        =   1614
               _ExtentY        =   503
               Alignment       =   2
               Mask            =   "###0.000"
               MaxLength       =   8
               DataType        =   2
               DataField       =   "TrStripLoss"
               ReCalcParent    =   "WTRDIFF"
               IdName          =   "TRSTRIPLOSS"
            End
            Begin MwfCtl.MWCTL_MED ATXT 
               Height          =   285
               Index           =   20
               Left            =   10755
               TabIndex        =   24
               Top             =   705
               Width           =   945
               _ExtentX        =   1667
               _ExtentY        =   503
               Alignment       =   2
               Enabled         =   0   'False
               Mask            =   "###0.000"
               MaxLength       =   8
               DataType        =   2
               ReCalcOn        =   "TRFRMETWT,TRUSDMETWT,TRRECDMETWT,TRCASTLOSS,TRCUTLOSS,TRGRINDLOSS,TRSCRAP,TRSTRIPLOSS,TRNETPCWT"
               IdName          =   "WTRDIFF"
            End
            Begin MwfCtl.MWCTL_BTN1 CmdRec 
               Height          =   495
               Left            =   1005
               TabIndex        =   23
               ToolTipText     =   "Generate Receipt Voucher"
               Top             =   780
               Width           =   885
               _ExtentX        =   1561
               _ExtentY        =   873
               BackColor       =   -2147483633
               MaskColor       =   8454143
               Caption         =   "&Rec Vch"
               BeginProperty LblFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
            End
            Begin MwfCtl.MWCTL_BTN1 CmdIss 
               Height          =   495
               Left            =   120
               TabIndex        =   22
               ToolTipText     =   "Generate Issue Voucher"
               Top             =   780
               Width           =   885
               _ExtentX        =   1561
               _ExtentY        =   873
               BackColor       =   -2147483633
               MaskColor       =   8454143
               Caption         =   "&Iss Vch"
               BeginProperty LblFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
            End
            Begin MwfCtl.MWCTL_MED ATXT 
               Height          =   285
               Index           =   29
               Left            =   10755
               TabIndex        =   20
               ToolTipText     =   "Enter Net Pc Wt"
               Top             =   420
               Width           =   945
               _ExtentX        =   1667
               _ExtentY        =   503
               Alignment       =   2
               Enabled         =   0   'False
               Mask            =   "###0.000"
               MaxLength       =   8
               DataType        =   2
               DataField       =   "TrNetPcWt"
               ReCalcParent    =   "WTRDIFF"
               ReCalcOn        =   "TRGRSPCWT"
               IdName          =   "TRNETPCWT"
            End
            Begin MwfCtl.MWCTL_MED ATXT 
               Height          =   285
               Index           =   31
               Left            =   10755
               TabIndex        =   21
               ToolTipText     =   "Enter Tree Dust Wt"
               Top             =   990
               Width           =   945
               _ExtentX        =   1667
               _ExtentY        =   503
               Alignment       =   2
               Mask            =   "###0.000"
               MaxLength       =   8
               DataType        =   2
               DataField       =   "TrDustWt"
               IdName          =   "TRDUSTWT"
            End
            Begin VB.Label Label19 
               BackStyle       =   0  'Transparent
               Caption         =   "Dust Wt"
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
               Left            =   9720
               TabIndex        =   66
               Top             =   990
               Width           =   1035
            End
            Begin VB.Label Label18 
               BackStyle       =   0  'Transparent
               Caption         =   "Net Pc Wt"
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
               Left            =   9690
               TabIndex        =   65
               Top             =   420
               Width           =   1065
            End
            Begin VB.Label LblMain 
               BackStyle       =   0  'Transparent
               BorderStyle     =   1  'Fixed Single
               Caption         =   "  Casting Details"
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
               TabIndex        =   59
               Top             =   90
               Width           =   11835
            End
            Begin VB.Label Label14 
               BackStyle       =   0  'Transparent
               Caption         =   "Difference"
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
               Left            =   9690
               TabIndex        =   58
               Top             =   705
               Width           =   1035
            End
            Begin VB.Label Label11 
               BackStyle       =   0  'Transparent
               Caption         =   "Grinding Loss"
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
               Left            =   7140
               TabIndex        =   57
               Top             =   420
               Width           =   1455
            End
            Begin VB.Label Label10 
               BackStyle       =   0  'Transparent
               Caption         =   "Strip Loss"
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
               Left            =   7140
               TabIndex        =   56
               Top             =   705
               Width           =   1455
            End
            Begin VB.Label ALBL 
               BackStyle       =   0  'Transparent
               Caption         =   "Gross Pc Wt"
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
               Index           =   53
               Left            =   7140
               TabIndex        =   55
               Top             =   990
               Width           =   1455
            End
            Begin VB.Label ALBL 
               BackStyle       =   0  'Transparent
               Caption         =   "Scrap"
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
               Index           =   21
               Left            =   4770
               TabIndex        =   54
               Top             =   990
               Width           =   1335
            End
            Begin VB.Label Label9 
               BackStyle       =   0  'Transparent
               Caption         =   "Cutting Loss"
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
               Left            =   4770
               TabIndex        =   53
               Top             =   705
               Width           =   1335
            End
            Begin VB.Label Label8 
               BackStyle       =   0  'Transparent
               Caption         =   "Casting Loss"
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
               Left            =   4770
               TabIndex        =   52
               Top             =   420
               Width           =   1335
            End
            Begin VB.Label Label7 
               BackStyle       =   0  'Transparent
               Caption         =   "Fresh Met Wt"
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
               Left            =   2070
               TabIndex        =   51
               Top             =   420
               Width           =   1635
            End
            Begin VB.Label Label6 
               BackStyle       =   0  'Transparent
               Caption         =   "ReUsed  Met Wt"
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
               Left            =   2070
               TabIndex        =   50
               Top             =   705
               Width           =   1635
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
               Height          =   285
               Index           =   55
               Left            =   150
               TabIndex        =   49
               Top             =   420
               Width           =   765
            End
            Begin VB.Label ALBL 
               BackStyle       =   0  'Transparent
               Caption         =   "Recvd  Met Wt"
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
               Index           =   36
               Left            =   2070
               TabIndex        =   48
               Top             =   990
               Width           =   1635
            End
         End
         Begin MwfCtl.MWCTL_MED ATXT 
            Height          =   285
            Index           =   27
            Left            =   1410
            TabIndex        =   63
            ToolTipText     =   "AutoGenerated Issue Voucher No"
            Top             =   7140
            Width           =   4140
            _ExtentX        =   7303
            _ExtentY        =   503
            Enabled         =   0   'False
            MaxLength       =   30
            DataType        =   4
            DataField       =   "TrIssVchNo"
            IdName          =   "TRISSVCHNO"
         End
         Begin MwfCtl.MWCTL_MED ATXT 
            Height          =   285
            Index           =   28
            Left            =   6900
            TabIndex        =   64
            ToolTipText     =   "AutoGenerated Receipt Voucher No"
            Top             =   7140
            Width           =   4080
            _ExtentX        =   7197
            _ExtentY        =   503
            Enabled         =   0   'False
            MaxLength       =   30
            DataType        =   4
            DataField       =   "TrRecVchNo"
            IdName          =   "TRRECVCHNO"
         End
         Begin MwfCtl.MWCTL_FLX GrdTreeDet 
            Height          =   6255
            Left            =   0
            TabIndex        =   26
            Top             =   600
            Width           =   13095
            _ExtentX        =   23098
            _ExtentY        =   11033
            Cols            =   15
            colname1        =   "TRDSR"
            heading1        =   "Sr          "
            datafld1        =   "TrdSr"
            datatype1       =   1
            mask1           =   "###0"
            maxlength1      =   4
            style1          =   2
            tooltiptext1    =   "Enter Sr No."
            colname2        =   "WBAGNO"
            datatype2       =   4
            maxlength2      =   14
            tooltiptext2    =   "Enter Bag No using Barcode Reader"
            colname3        =   "TRDBYY"
            heading3        =   "Yy"
            datafld3        =   "TrdBYy"
            datatype3       =   4
            recalcparent3   =   "GrdTreeDet(TRDBNO),GrdTreeDet(TRDBAGPCS),GrdTreeDet(WDMCD),GrdTreeDet(TRDDCWT)"
            maxlength3      =   2
            tooltiptext3    =   "Enter Bag Year"
            colname4        =   "TRDBCHR"
            heading4        =   "Chr   "
            datafld4        =   "TrdBChr"
            datatype4       =   4
            recalcparent4   =   "GrdTreeDet(TRDBNO),GrdTreeDet(TRDBAGPCS),GrdTreeDet(WDMCD),GrdTreeDet(TRDDCWT)"
            maxlength4      =   3
            tooltiptext4    =   "Enter Bag Character"
            colname5        =   "TRDBNO"
            heading5        =   "No         "
            datafld5        =   "TrdBNo"
            datatype5       =   1
            mask5           =   "######0"
            recalcon5       =   "GrdTreeDet(TRDBYY),GrdTreeDet(TRDBCHR)"
            recalcparent5   =   "GrdTreeDet(TRDBAGPCS),GrdTreeDet(WDMCD),GrdTreeDet(TRDDCWT)"
            maxlength5      =   7
            tooltiptext5    =   "Enter Bag No."
            colname6        =   "TRDBAGPCS"
            heading6        =   "Bag Pcs "
            datafld6        =   "TrdBagPcs"
            datatype6       =   2
            mask6           =   "##0.0"
            recalcon6       =   "GrdTreeDet(TRDBYY),GrdTreeDet(TRDBCHR),GrdTreeDet(TRDBNO)"
            recalcparent6   =   "GrdTreeDet(TRDDCWT)"
            maxlength6      =   5
            tooltiptext6    =   "Enter Bag Quantity"
            colname7        =   "TRDDCWT"
            heading7        =   "Dia/Col   Wt "
            datafld7        =   "TrdDcWt"
            datatype7       =   2
            mask7           =   "####0.000"
            recalcon7       =   "GrdTreeDet(TRDBYY),GrdTreeDet(TRDBCHR),GrdTreeDet(TRDBNO),GrdTreeDet(TRDBAGPCS)"
            maxlength7      =   9
            tooltiptext7    =   "Enter Bag Gross Weight"
            colname8        =   "TRDMETXWT"
            heading8        =   "Metal Wt     "
            datafld8        =   "TrdMetXWt"
            datatype8       =   2
            mask8           =   "####0.000"
            maxlength8      =   9
            colname9        =   "WDMCD"
            heading9        =   "Dsg Cd                      "
            datatype9       =   4
            recalcon9       =   "GrdTreeDet(TRDBYY),GrdTreeDet(TRDBCHR),GrdTreeDet(TRDBNO)"
            maxlength9      =   15
            style9          =   2
            tooltiptext9    =   "Design Code"
            colname10       =   "WDMCOL"
            heading10       =   "Dsg Color "
            datatype10      =   4
            maxlength10     =   5
            style10         =   2
            tooltiptext10   =   "Design Color"
            colname11       =   "WDMKT"
            heading11       =   "KT         "
            datatype11      =   4
            maxlength11     =   5
            style11         =   2
            tooltiptext11   =   "KT"
            colname12       =   "WORDNO"
            heading12       =   "Order No                        "
            datatype12      =   4
            maxlength12     =   15
            style12         =   2
            tooltiptext12   =   "Order No"
            colname13       =   "WCMCD"
            heading13       =   "Customer      "
            datatype13      =   4
            maxlength13     =   8
            style13         =   2
            tooltiptext13   =   "Customer Code"
            colname14       =   "TRDDESC"
            datafld14       =   "TrdDesc"
            datatype14      =   4
            recalcon14      =   "WTRDDESC"
            maxlength14     =   60
            style14         =   0
         End
         Begin MwfCtl.MWCTL_MED ATXT 
            Height          =   285
            Index           =   48
            Left            =   1410
            TabIndex        =   67
            ToolTipText     =   "Enter Description"
            Top             =   6870
            Width           =   9270
            _ExtentX        =   16351
            _ExtentY        =   503
            MaxLength       =   60
            ReCalcParent    =   "GrdTreeDet(TRDDESC)"
            IdName          =   "WTRDDESC"
         End
         Begin MwfCtl.MWCTL_MED ATXT 
            Height          =   285
            Index           =   33
            Left            =   6240
            TabIndex        =   8
            ToolTipText     =   "Enter Rubber Wt"
            Top             =   300
            Width           =   1035
            _ExtentX        =   1826
            _ExtentY        =   503
            Alignment       =   2
            Mask            =   "-###0.000"
            MaxLength       =   9
            DataType        =   2
            DataField       =   "TRRUBWT"
            ReCalcParent    =   "WTRGRSMETREQ"
            IdName          =   "TRRUBWT"
         End
         Begin VB.Label Label22 
            BackStyle       =   0  'Transparent
            Caption         =   "Rubber Wt"
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
            Left            =   5040
            TabIndex        =   69
            Top             =   300
            Width           =   1095
         End
         Begin VB.Label Label20 
            BackStyle       =   0  'Transparent
            Caption         =   "Desc"
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
            Left            =   0
            TabIndex        =   68
            Top             =   6870
            Width           =   1275
         End
         Begin VB.Label Label17 
            BackStyle       =   0  'Transparent
            Caption         =   "Receipt Vch"
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
            Left            =   5700
            TabIndex        =   62
            Top             =   7170
            Width           =   1155
         End
         Begin VB.Label Label16 
            BackStyle       =   0  'Transparent
            Caption         =   "Issue Vch"
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
            Left            =   30
            TabIndex        =   61
            Top             =   7140
            Width           =   1275
         End
         Begin VB.Label Label5 
            BackStyle       =   0  'Transparent
            Caption         =   "Add MetWt"
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
            Left            =   8640
            TabIndex        =   46
            Top             =   300
            Width           =   1215
         End
         Begin VB.Label Label1 
            BackStyle       =   0  'Transparent
            Caption         =   "Tree GrsWt"
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
            Left            =   2400
            TabIndex        =   45
            Top             =   300
            Width           =   1185
         End
         Begin VB.Label LblTrfQty 
            BackStyle       =   0  'Transparent
            Caption         =   "Tree Kt"
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
            Left            =   7365
            TabIndex        =   43
            Top             =   15
            Width           =   795
         End
         Begin VB.Label LblTFrBLoc 
            BackStyle       =   0  'Transparent
            Caption         =   "Tree Col"
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
            Left            =   5550
            TabIndex        =   42
            Top             =   15
            Width           =   960
         End
         Begin VB.Label LblTdDesc 
            BackStyle       =   0  'Transparent
            Caption         =   "Desc"
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
            Left            =   0
            TabIndex        =   40
            Top             =   8820
            Width           =   660
         End
         Begin VB.Label LblTFrRmLoc 
            BackStyle       =   0  'Transparent
            Caption         =   "Wax Sp Gr"
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
            Left            =   0
            TabIndex        =   37
            Top             =   285
            Width           =   1185
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
            Height          =   285
            Index           =   2
            Left            =   4080
            TabIndex        =   36
            Top             =   15
            Width           =   555
         End
      End
      Begin VB.Frame FraCB 
         Height          =   3675
         Left            =   2085
         TabIndex        =   31
         Top             =   5460
         Width           =   9300
         Begin VB.OptionButton OptCpyBagFr 
            Caption         =   "Scope"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   225
            Index           =   0
            Left            =   1290
            TabIndex        =   120
            ToolTipText     =   "Specify Copy From Scope / Bag Alias No."
            Top             =   450
            Width           =   855
         End
         Begin VB.OptionButton OptCpyBagFr 
            Caption         =   "Bag Alias No."
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   225
            Index           =   1
            Left            =   2340
            TabIndex        =   119
            ToolTipText     =   "Specify Copy From Scope / Bag Alias No."
            Top             =   450
            Width           =   1575
         End
         Begin VB.Frame FraCBScope 
            Caption         =   "Scope"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   2895
            Left            =   60
            TabIndex        =   72
            Top             =   720
            Width           =   4575
            Begin MwfCtl.MWCTL_MED ATXT 
               Height          =   285
               Index           =   61
               Left            =   1245
               TabIndex        =   74
               ToolTipText     =   "Enter From Bag Year"
               Top             =   540
               Width           =   360
               _ExtentX        =   635
               _ExtentY        =   503
               MaxLength       =   2
               DataType        =   4
               IdName          =   "WCBBYYFR"
            End
            Begin MwfCtl.MWCTL_MED ATXT 
               Height          =   285
               Index           =   70
               Left            =   2400
               TabIndex        =   76
               ToolTipText     =   "Enter From Bag Number"
               Top             =   540
               Width           =   780
               _ExtentX        =   1376
               _ExtentY        =   503
               Alignment       =   2
               Mask            =   "######0"
               MaxLength       =   7
               DataType        =   1
               IdName          =   "WCBBNOFR"
            End
            Begin MwfCtl.MWCTL_MED ATXT 
               Height          =   285
               Index           =   71
               Left            =   1695
               TabIndex        =   75
               ToolTipText     =   "Enter From Bag Character"
               Top             =   540
               Width           =   600
               _ExtentX        =   1058
               _ExtentY        =   503
               MaxLength       =   3
               DataType        =   4
               IdName          =   "WCBBCHRFR"
            End
            Begin MwfCtl.MWCTL_MED ATXT 
               Height          =   285
               Index           =   80
               Left            =   1245
               TabIndex        =   77
               ToolTipText     =   "Enter To Bag Year"
               Top             =   825
               Width           =   360
               _ExtentX        =   635
               _ExtentY        =   503
               MaxLength       =   2
               DataType        =   4
               IdName          =   "WCBBYYTO"
            End
            Begin MwfCtl.MWCTL_MED ATXT 
               Height          =   285
               Index           =   81
               Left            =   2400
               TabIndex        =   79
               ToolTipText     =   "Enter To Bag Number"
               Top             =   825
               Width           =   780
               _ExtentX        =   1376
               _ExtentY        =   503
               Alignment       =   2
               Mask            =   "######0"
               MaxLength       =   7
               DataType        =   1
               IdName          =   "WCBBNOTO"
            End
            Begin MwfCtl.MWCTL_MED ATXT 
               Height          =   285
               Index           =   82
               Left            =   1695
               TabIndex        =   78
               ToolTipText     =   "Enter To Bag Character"
               Top             =   825
               Width           =   600
               _ExtentX        =   1058
               _ExtentY        =   503
               MaxLength       =   3
               DataType        =   4
               IdName          =   "WCBBCHRTO"
            End
            Begin MwfCtl.MWCTL_MED ATXT 
               Height          =   285
               Index           =   77
               Left            =   1950
               TabIndex        =   81
               ToolTipText     =   "Enter From Bag Order Year "
               Top             =   1110
               Width           =   360
               _ExtentX        =   635
               _ExtentY        =   503
               MaxLength       =   2
               DataType        =   4
               IdName          =   "WCBODYYFR"
            End
            Begin MwfCtl.MWCTL_MED ATXT 
               Height          =   285
               Index           =   79
               Left            =   3090
               TabIndex        =   83
               ToolTipText     =   "Enter From Bag Order Number"
               Top             =   1110
               Width           =   780
               _ExtentX        =   1376
               _ExtentY        =   503
               Alignment       =   2
               Mask            =   "#####0"
               MaxLength       =   6
               DataType        =   1
               IdName          =   "WCBODNOFR"
            End
            Begin MwfCtl.MWCTL_MED ATXT 
               Height          =   285
               Index           =   83
               Left            =   1245
               TabIndex        =   80
               ToolTipText     =   "Enter From Bag Order Tc "
               Top             =   1110
               Width           =   600
               _ExtentX        =   1058
               _ExtentY        =   503
               MaxLength       =   3
               DataType        =   4
               IdName          =   "WCBODTCFR"
            End
            Begin MwfCtl.MWCTL_MED ATXT 
               Height          =   285
               Index           =   84
               Left            =   2400
               TabIndex        =   82
               ToolTipText     =   "Enter From Bag Order Character"
               Top             =   1110
               Width           =   600
               _ExtentX        =   1058
               _ExtentY        =   503
               MaxLength       =   3
               DataType        =   4
               IdName          =   "WCBODCHRFR"
            End
            Begin MwfCtl.MWCTL_MED ATXT 
               Height          =   285
               Index           =   85
               Left            =   3945
               TabIndex        =   84
               ToolTipText     =   "Enter From Bag Order Serial Number"
               Top             =   1110
               Width           =   570
               _ExtentX        =   1005
               _ExtentY        =   503
               Alignment       =   2
               Mask            =   "###0"
               MaxLength       =   4
               DataType        =   1
               IdName          =   "WCBODSRFR"
            End
            Begin MwfCtl.MWCTL_MED ATXT 
               Height          =   285
               Index           =   86
               Left            =   1950
               TabIndex        =   86
               ToolTipText     =   "Enter To Bag Order Year"
               Top             =   1395
               Width           =   360
               _ExtentX        =   635
               _ExtentY        =   503
               MaxLength       =   2
               DataType        =   4
               IdName          =   "WCBODYYTO"
            End
            Begin MwfCtl.MWCTL_MED ATXT 
               Height          =   285
               Index           =   87
               Left            =   3090
               TabIndex        =   88
               ToolTipText     =   "Enter To Bag Order Number"
               Top             =   1395
               Width           =   780
               _ExtentX        =   1376
               _ExtentY        =   503
               Alignment       =   2
               Mask            =   "#####0"
               MaxLength       =   6
               DataType        =   1
               IdName          =   "WCBODNOTO"
            End
            Begin MwfCtl.MWCTL_MED ATXT 
               Height          =   285
               Index           =   88
               Left            =   1245
               TabIndex        =   85
               ToolTipText     =   "Enter To Bag Order Tc"
               Top             =   1395
               Width           =   600
               _ExtentX        =   1058
               _ExtentY        =   503
               MaxLength       =   3
               DataType        =   4
               IdName          =   "WCBODTCTO"
            End
            Begin MwfCtl.MWCTL_MED ATXT 
               Height          =   285
               Index           =   89
               Left            =   2400
               TabIndex        =   87
               ToolTipText     =   "Enter To Bag Order Character"
               Top             =   1395
               Width           =   600
               _ExtentX        =   1058
               _ExtentY        =   503
               MaxLength       =   3
               DataType        =   4
               IdName          =   "WCBODCHRTO"
            End
            Begin MwfCtl.MWCTL_MED ATXT 
               Height          =   285
               Index           =   90
               Left            =   3945
               TabIndex        =   89
               ToolTipText     =   "Enter To Bag Order Serial Number"
               Top             =   1395
               Width           =   570
               _ExtentX        =   1005
               _ExtentY        =   503
               Alignment       =   2
               Mask            =   "###0"
               MaxLength       =   4
               DataType        =   1
               IdName          =   "WCBODSRTO"
            End
            Begin MwfCtl.MWCTL_MED ATXT 
               Height          =   285
               Index           =   91
               Left            =   1245
               TabIndex        =   90
               ToolTipText     =   "Enter From Design Code"
               Top             =   1680
               Width           =   2415
               _ExtentX        =   4260
               _ExtentY        =   503
               MaxLength       =   15
               DataType        =   4
               IdName          =   "WCBDMCDFR"
            End
            Begin MwfCtl.MWCTL_MED ATXT 
               Height          =   285
               Index           =   92
               Left            =   1245
               TabIndex        =   91
               ToolTipText     =   "Enter To Design Code"
               Top             =   1965
               Width           =   2415
               _ExtentX        =   4260
               _ExtentY        =   503
               MaxLength       =   15
               DataType        =   4
               IdName          =   "WCBDMCDTO"
            End
            Begin MwfCtl.MWCTL_MED ATXT 
               Height          =   285
               Index           =   60
               Left            =   1245
               TabIndex        =   73
               ToolTipText     =   "Enter The Location To Which The Bags Have To Be Copied"
               Top             =   240
               Width           =   1305
               _ExtentX        =   2302
               _ExtentY        =   503
               MaxLength       =   12
               DataType        =   4
               IdName          =   "WCBBLOC"
            End
            Begin MwfCtl.MWCTL_MED ATXT 
               Height          =   285
               Index           =   22
               Left            =   1245
               TabIndex        =   92
               ToolTipText     =   "Enter From Design KT"
               Top             =   2250
               Width           =   825
               _ExtentX        =   1455
               _ExtentY        =   503
               MaxLength       =   5
               DataType        =   4
               IdName          =   "WCBODKTFR"
            End
            Begin MwfCtl.MWCTL_MED ATXT 
               Height          =   285
               Index           =   23
               Left            =   2835
               TabIndex        =   93
               ToolTipText     =   "Enter To Design KT"
               Top             =   2250
               Width           =   825
               _ExtentX        =   1455
               _ExtentY        =   503
               MaxLength       =   5
               DataType        =   4
               IdName          =   "WCBODKTTO"
            End
            Begin MwfCtl.MWCTL_MED ATXT 
               Height          =   285
               Index           =   24
               Left            =   1245
               TabIndex        =   94
               ToolTipText     =   "Enter From Design Color"
               Top             =   2535
               Width           =   825
               _ExtentX        =   1455
               _ExtentY        =   503
               MaxLength       =   5
               DataType        =   4
               IdName          =   "WCBODCOLFR"
            End
            Begin MwfCtl.MWCTL_MED ATXT 
               Height          =   285
               Index           =   26
               Left            =   2835
               TabIndex        =   95
               ToolTipText     =   "Enter To Design Color"
               Top             =   2535
               Width           =   825
               _ExtentX        =   1455
               _ExtentY        =   503
               MaxLength       =   5
               DataType        =   4
               IdName          =   "WCBODCOLTO"
            End
            Begin VB.Label ALBL 
               BackStyle       =   0  'Transparent
               Caption         =   "Bag Fr"
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
               Index           =   22
               Left            =   120
               TabIndex        =   118
               Top             =   540
               Width           =   885
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
               Height          =   285
               Index           =   23
               Left            =   1620
               TabIndex        =   117
               Top             =   540
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
               Height          =   285
               Index           =   24
               Left            =   2325
               TabIndex        =   116
               Top             =   540
               Width           =   105
            End
            Begin VB.Label ALBL 
               BackStyle       =   0  'Transparent
               Caption         =   "Bag To"
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
               Index           =   37
               Left            =   120
               TabIndex        =   115
               Top             =   825
               Width           =   885
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
               Height          =   285
               Index           =   38
               Left            =   1620
               TabIndex        =   114
               Top             =   825
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
               Height          =   285
               Index           =   42
               Left            =   2325
               TabIndex        =   113
               Top             =   825
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
               Height          =   285
               Index           =   20
               Left            =   3870
               TabIndex        =   112
               Top             =   1110
               Width           =   105
            End
            Begin VB.Label ALBL 
               BackStyle       =   0  'Transparent
               Caption         =   "Order Fr"
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
               Index           =   25
               Left            =   120
               TabIndex        =   111
               Top             =   1110
               Width           =   1065
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
               Height          =   285
               Index           =   26
               Left            =   1860
               TabIndex        =   110
               Top             =   1110
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
               Height          =   285
               Index           =   43
               Left            =   2325
               TabIndex        =   109
               Top             =   1110
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
               Height          =   285
               Index           =   44
               Left            =   3015
               TabIndex        =   108
               Top             =   1110
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
               Height          =   285
               Index           =   45
               Left            =   3870
               TabIndex        =   107
               Top             =   1395
               Width           =   105
            End
            Begin VB.Label ALBL 
               BackStyle       =   0  'Transparent
               Caption         =   "Order To"
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
               Index           =   46
               Left            =   120
               TabIndex        =   106
               Top             =   1395
               Width           =   1065
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
               Height          =   285
               Index           =   47
               Left            =   1860
               TabIndex        =   105
               Top             =   1395
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
               Height          =   285
               Index           =   48
               Left            =   2325
               TabIndex        =   104
               Top             =   1395
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
               Height          =   285
               Index           =   49
               Left            =   3015
               TabIndex        =   103
               Top             =   1395
               Width           =   105
            End
            Begin VB.Label ALBL 
               BackStyle       =   0  'Transparent
               Caption         =   "Design To"
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
               Index           =   51
               Left            =   120
               TabIndex        =   102
               Top             =   1965
               Width           =   1125
            End
            Begin VB.Label ALBL 
               BackStyle       =   0  'Transparent
               Caption         =   "Design Fr"
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
               Index           =   50
               Left            =   120
               TabIndex        =   101
               Top             =   1680
               Width           =   1125
            End
            Begin VB.Label ALBL 
               BackStyle       =   0  'Transparent
               Caption         =   "Bag Loc"
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
               Left            =   120
               TabIndex        =   100
               Top             =   240
               Width           =   885
            End
            Begin VB.Label ALBL 
               BackStyle       =   0  'Transparent
               Caption         =   "Dsg KT Fr"
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
               Left            =   120
               TabIndex        =   99
               Top             =   2250
               Width           =   1125
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
               Height          =   285
               Index           =   3
               Left            =   2370
               TabIndex        =   98
               Top             =   2250
               Width           =   615
            End
            Begin VB.Label ALBL 
               BackStyle       =   0  'Transparent
               Caption         =   "Dsg Col Fr"
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
               Left            =   120
               TabIndex        =   97
               Top             =   2550
               Width           =   1125
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
               Height          =   285
               Index           =   6
               Left            =   2370
               TabIndex        =   96
               Top             =   2550
               Width           =   435
            End
         End
         Begin VB.Frame FraCBBag 
            Caption         =   "Bag Alias No."
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   3435
            Left            =   4680
            TabIndex        =   70
            Top             =   165
            Width           =   4575
            Begin MwfCtl.MWCTL_MED ATXT 
               Height          =   3165
               Index           =   478
               Left            =   60
               TabIndex        =   71
               ToolTipText     =   "Copy Bag Alias No."
               Top             =   195
               Width           =   4440
               _ExtentX        =   7832
               _ExtentY        =   5583
               Alignment       =   3
               MaxLength       =   0
               IdName          =   "WRFIDSTR"
            End
         End
         Begin MwfCtl.MWCTL_BTN1 CmdCBGo 
            Height          =   285
            Left            =   4050
            TabIndex        =   25
            Top             =   420
            Width           =   585
            _ExtentX        =   0
            _ExtentY        =   0
            BackColor       =   -2147483633
            MaskColor       =   8454143
            Caption         =   "&Go"
            BeginProperty LblFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
         End
         Begin VB.Label Label23 
            BackStyle       =   0  'Transparent
            Caption         =   "Copy from "
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
            Left            =   120
            TabIndex        =   121
            Top             =   420
            Width           =   1005
         End
         Begin VB.Label LblMain 
            BackStyle       =   0  'Transparent
            BorderStyle     =   1  'Fixed Single
            Caption         =   "  Copy Bags"
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
            Index           =   1
            Left            =   0
            TabIndex        =   60
            Top             =   90
            Width           =   4620
         End
      End
      Begin VB.Frame FraShowTot 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1275
         Left            =   1200
         TabIndex        =   123
         Top             =   7920
         Width           =   11175
         Begin MwfCtl.MWCTL_MED ATXT 
            Height          =   285
            Index           =   21
            Left            =   1620
            TabIndex        =   125
            Top             =   480
            Width           =   1305
            _ExtentX        =   2302
            _ExtentY        =   503
            Alignment       =   2
            Enabled         =   0   'False
            Mask            =   "###0.0"
            MaxLength       =   6
            DataType        =   2
            IdName          =   "WBAGPCS"
         End
         Begin MwfCtl.MWCTL_MED ATXT 
            Height          =   285
            Index           =   18
            Left            =   4680
            TabIndex        =   127
            Top             =   480
            Width           =   1725
            _ExtentX        =   3043
            _ExtentY        =   503
            Alignment       =   2
            Enabled         =   0   'False
            Mask            =   "###0.000"
            MaxLength       =   8
            DataType        =   2
            IdName          =   "WTOTDCWT"
         End
         Begin MwfCtl.MWCTL_MED ATXT 
            Height          =   285
            Index           =   32
            Left            =   7800
            TabIndex        =   128
            ToolTipText     =   "Other Wt"
            Top             =   480
            Width           =   2205
            _ExtentX        =   3889
            _ExtentY        =   503
            Alignment       =   2
            Enabled         =   0   'False
            Mask            =   "#######0.000"
            MaxLength       =   12
            DataType        =   2
            IdName          =   "WTOTMETXWT"
         End
         Begin MwfCtl.MWCTL_MED ATXT 
            Height          =   285
            Index           =   5
            Left            =   1620
            TabIndex        =   131
            ToolTipText     =   "Enter Metal Net Wt"
            Top             =   765
            Width           =   1305
            _ExtentX        =   2302
            _ExtentY        =   503
            Alignment       =   2
            Mask            =   "###0.000"
            MaxLength       =   8
            DataType        =   2
            DataField       =   "TrNetMetWt"
            IdName          =   "TRNETMETWT"
         End
         Begin MwfCtl.MWCTL_MED ATXT 
            Height          =   285
            Index           =   19
            Left            =   5100
            TabIndex        =   133
            ToolTipText     =   "Gross Metal Required"
            Top             =   765
            Width           =   1305
            _ExtentX        =   2302
            _ExtentY        =   503
            Alignment       =   2
            Enabled         =   0   'False
            Mask            =   "###0.000"
            MaxLength       =   8
            DataType        =   2
            IdName          =   "WTRGRSMETREQ"
         End
         Begin MwfCtl.MWCTL_BTN1 CmdShowTotOk 
            Height          =   375
            Left            =   10320
            TabIndex        =   135
            ToolTipText     =   "Show Totals"
            Top             =   480
            Width           =   495
            _ExtentX        =   873
            _ExtentY        =   661
            BackColor       =   -2147483633
            MaskColor       =   8454143
            Caption         =   "&OK"
            BeginProperty LblFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
         End
         Begin VB.Label Label13 
            BackStyle       =   0  'Transparent
            Caption         =   "Grs Met Reqd"
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
            Left            =   3240
            TabIndex        =   134
            Top             =   765
            Width           =   1485
         End
         Begin VB.Label Label4 
            BackStyle       =   0  'Transparent
            Caption         =   "Metal NetWt"
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
            Left            =   120
            TabIndex        =   132
            Top             =   765
            Width           =   1335
         End
         Begin VB.Label Label12 
            BackStyle       =   0  'Transparent
            Caption         =   "Dia/Col Wt"
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
            Left            =   3240
            TabIndex        =   130
            Top             =   480
            Width           =   1275
         End
         Begin VB.Label Label21 
            BackStyle       =   0  'Transparent
            Caption         =   "Other Wt"
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
            Left            =   6720
            TabIndex        =   129
            Top             =   480
            Width           =   915
         End
         Begin VB.Label Label15 
            BackStyle       =   0  'Transparent
            Caption         =   "Total Bag Pcs"
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
            Left            =   120
            TabIndex        =   126
            Top             =   480
            Width           =   1875
         End
         Begin VB.Label LblMain 
            BackStyle       =   0  'Transparent
            BorderStyle     =   1  'Fixed Single
            Caption         =   "  Totals"
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
            Index           =   2
            Left            =   0
            TabIndex        =   124
            Top             =   120
            Width           =   11115
         End
      End
      Begin VB.Label Label24 
         BackStyle       =   0  'Transparent
         Caption         =   "Click on Show Totals to see the Totals"
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
         Left            =   2880
         TabIndex        =   136
         Top             =   9360
         Width           =   3675
      End
   End
End
Attribute VB_Name = "EmrFrmTree"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim moCn As MwfLib.MDOConnection   '*** (Bef speed) bef 24/11/06
Dim ms_TrCoCd As String, ms_TrYmd As String, ms_TrChr As String
'*** For the Toggle Options of Command Buttons
Enum en_TreeFra
    CpyBag = 1
    ShowTot = 2 '578-33
End Enum

'*** Declare the variables ms_LockDt, mf_DtLocked for the Lock Date Check
Dim ms_LockDt As String
Dim ms_FutureDt As String
Dim mf_DtLocked As Boolean
'*** Arr_Cv stores the causes validation state, Arr_Ed the Ena/Disa state of Cmd Buttons
Dim mArr_Cv As Variant, mArr_Ed As Variant
Dim mb_AddRow As Boolean
  
Private Sub ADC_Load()
  '*** Redim the arrays mArr_Cv and mArr_Ed to the number of command buttons
  '*** Set the First Key Control property to the field TChr
  '*** Set the First Non Key Control property to TrDt
  '*** Set the Previous and Next Control Properties for the Grids GrdTreeDet
  '*** Set the Hot Keys for the Grid GrdTreeDet'A'
  '*** Bring the FraNKeyAll frame to the top
  
  ReDim mArr_Cv(10), mArr_Ed(10)
  
  Set moCn = ADC.Connection    '*** (Bef speed) bef 24/11/06
  Set ADC.FirKeyCtl = ADC("TrYmd")
  Set ADC.FirNKeyCtl = ADC("TrDt")
  ADC.Child = "GrdTreeDet"
  Set GrdTreeDet.PrevCtl = ADC("TrCol")
  Set GrdTreeDet.NextCtl = ADC

  GrdTreeDet.HotKey = "A"
  FraNKeyAll.ZOrder (0)

  ms_TrCoCd = gs_CoCd
  
  ' 4.1.0.0 - RFID Module Check
  If Not moCn.RecSeek("Select 'x' from Head Where HCoCd='" + gs_CoCd + "' and HCd='" + ctSelfCmCd + "' and HRfIdMod='Y' ") Then _
    ADC("WRFIDSTR").Enabled = False ': FraCB.Width = 4630
  
  mb_AddRow = True
      
End Sub
Private Sub ADC_KeyWhen()
  '*** Disable All command buttons
  '*** Hide all frames except Frame FraNKeyAll
  '*** setting the default values for TrCoCd,TrYmd and TrChr from the previous entry
  
  mf_DtLocked = False 'IG.192-Set LockDt as false
 
  Call EnaDisaCmds(True)
  Call HideAllFras

 
  If ms_TrCoCd <> "" Then ADC("TrCoCd") = ms_TrCoCd
  If ms_TrYmd <> "" Then ADC("TrYmd") = ms_TrYmd
  If ms_TrChr <> "" Then ADC("TrChr") = ms_TrChr
  
  '****** Sachin 2.14.0 - [24-11-2006] - [Multi-Company] ******
  Me.Caption = GetFrmCaption(gs_CoCd, "Tree Entry")
  '****** Sachin 2.14.0 - [24-11-2006] - [Multi-Company] ******

End Sub
Private Sub ADC_KeyValid(Cancel As Boolean, ErrMsg As String)
  '*** Check whether the TrYmd,TrChr are valid
  
  If ADC("TrYmd") = "" Then Cancel = True: ErrMsg = "Enter Voucher Year": Exit Sub
  If ADC("TrChr") = "" Then Cancel = True: ErrMsg = "Enter Voucher Character": Exit Sub
End Sub

Private Sub ADC_NonKeyValid(pr_Cancel As Boolean, pr_ErrMsg As String)
    ' **** Manali 3.8.0
    If ADC("TrCstDt") <> "01/01/80" And ADC("TRNETPCWT") < 0 Then pr_Cancel = True: pr_ErrMsg = "Net Pc Wt Cannot Be Less Than Zero": Exit Sub
    Call GetTotals     '578-33 calculating total
End Sub

Private Sub ADC_NonKeyWhen(Cancel As Boolean, ErrMsg As String)
  '*** Store the TrYmd and TrChr values for the next entry
  '*** Give default value for TrDt and TrCst   blank (Empty date gives '01/01/1980')
  
  Call EnaDisaCmds(False)
  
  
'  ms_LockDt = moCn.GetFldVal("Select PDesc225 from Param where PTyp= 'LOCKDT' and PMCd= 'LOCKDT'")
'  ms_FutureDt = moCn.GetFldVal("Select PDesc225 from Param where PTyp= 'LOCKDT' and PMCd= 'FUTUREDT'")
'  If ms_LockDt = "" Then ms_LockDt = "01/01/80"
'  If ms_FutureDt = "" Then ms_FutureDt = "31/12/28"
'
  If ADC.AddFndMode = xaddmode Then

'     Dim ws_Dt As String
'    ws_Dt = IIF((moCn.SrvrDate < CDate(ms_LockDt)) _
'                               Or _
'                (moCn.SrvrDate > CDate(ms_FutureDt)), _
'                "", moCn.SrvrDate)
'    If ws_Dt <> "" Then ADC("TrDt") = ws_Dt
     '53 company code parameter added
     ADC("TrDt") = GetNextDate(ADC("TrCoCd"))  'moCn.SrvrDate
     ADC("TrCstDt") = MWLib.EmptyDate
    ADC("TrWaxSpGr") = moCn.GetFldVal("select PNum from param where PTyp='WAXSPGR' and PMCd= 'WAXSPGR' and PSCd=''")
  End If


'  mf_DtLocked = IIF((ADC("TrDt") < CDate(ms_LockDt)) _
'                               Or _
'                   (ADC("TrDt") > CDate(ms_FutureDt)), _
'                True, False)
'  If mf_DtLocked = True Then
'    GrdTreeDet.AllowAdd = False: GrdTreeDet.AllowDelete = False
'  Else
'    GrdTreeDet.AllowAdd = True: GrdTreeDet.AllowDelete = True
'  End If
'IG.127-If Issue voucher no. in tree entry is not exist then TrIssVchNo will be blank
Dim wIssVchNo, wRecVchNo As String
wIssVchNo = moCn.GetFldVal("select TrIssVchNo from TreeMst where TrYmd='" + ADC("TRYMD") + "' and TrChr='" + ADC("TRCHR") + "' and TrNo=" + CStr(ADC("TrNo")))
If wIssVchNo <> "" Then
  Dim wIssArr() As String
  wIssArr = Split(wIssVchNo, "/")
  If UBound(wIssArr) = 4 Then
    If Not (moCn.RecSeek("Select 'x' from Txn where TCocd= '" + wIssArr(0) + "' and TTc= '" + wIssArr(1) + "' and TYy= '" + wIssArr(2) + "' and TChr= '" + wIssArr(3) + "' and TNo= " + wIssArr(4) + "  ")) Then
      ADC("TRISSVCHNO") = ""
      ADC.SaveRec
    End If
  End If
End If

'IG.127-If receipt voucher no. in tree entry is not exist then TrRecVchNo will be blank
wRecVchNo = moCn.GetFldVal("select TrRecVchNo from TreeMst where TrYmd='" + ADC("TRYMD") + "' and TrChr='" + ADC("TRCHR") + "' and TrNo=" + CStr(ADC("TrNo")))
If wRecVchNo <> "" Then
  Dim wRecArr() As String
  wRecArr = Split(wRecVchNo, "/")
  If UBound(wRecArr) = 4 Then
    If Not (moCn.RecSeek("Select 'x' from Txn where TCocd= '" + wRecArr(0) + "' and TTc= '" + wRecArr(1) + "' and TYy= '" + wRecArr(2) + "' and TChr= '" + wRecArr(3) + "' and TNo= " + wRecArr(4) + "  ")) Then
        ADC("TRRECVCHNO") = ""
        ADC.SaveRec
    End If
  End If
End If
'vk.4 setting default value
OptCpyBagFr(0).Value = True
 
End Sub

Private Sub ATXT_FldWhen(Index As Integer, ByVal IdName As String, _
                          Cancel As Boolean, ErrMsg As String)
  '*** Set Helps For Each Field
  '*** Do not allow the user to edit the voucher date if record exists in the 'A' grid
   Dim wOthLocBag As String

'  If mf_DtLocked = True Then
'    Cancel = True: ErrMsg = "Cannot Edit Locked Entry": Exit Sub
'  End If
  If GrdTreeDet.Rows - 1 > 0 Then
   Select Case UCase(IdName)
      Case Is = UCase("TrDt"), UCase("TrRmCd"), UCase("TrCol")
         Cancel = True
        ErrMsg = "Cannot Edit Value when Recs exist in Detail Box"
        Exit Sub
   End Select
  End If
  If ADC("TrCstdt") = MWLib.EmptyDate Then
     Select Case UCase(IdName)
       Case Is = UCase("TrUsdMetWt"), UCase("TrRecdMetWt"), UCase("TrCastLoss"), _
                  UCase("TrCutLoss"), UCase("TrScrap"), UCase("TrGrindLoss"), _
                  UCase("TrStripLoss"), UCase("TrGrsPcWt"), UCase("TrFrMetWt")
          Cancel = True: ErrMsg = "Cannot Edit Value,when casting date is blank": Exit Sub
     End Select
  End If
  'IG -If Issue voucher exists can't change frsh&reused met.wt & X values
  If ADC("TrIssVchNo") <> "" Then
     Select Case UCase(IdName)
       Case Is = UCase("TrFrMetWt"), UCase("TrUsdMetWt"), UCase("TrDt"), _
                 UCase("TrCol"), UCase("TrRmcd"), UCase("TrWaxSpGr"), UCase("TrGrsWaxWt"), _
                 UCase("TRRUBWT"), UCase("TrAddMetWt"), UCase("TrNetMetWt")
          Cancel = True: ErrMsg = "Cannot Edit Value,When Issue Voucher Exists": Exit Sub
     End Select
  End If
  'IG -If Receipt voucher exists can't change any values
  If ADC("TrRecVchNo") <> "" Then
     Select Case UCase(IdName)
       Case Is = UCase("TrFrMetWt"), UCase("TrUsdMetWt"), UCase("TrDt"), _
                 UCase("TrCol"), UCase("TrRmcd"), UCase("TrWaxSpGr"), UCase("TrGrsWaxWt"), _
                 UCase("TRRUBWT"), UCase("TrAddMetWt"), UCase("TrNetMetWt"), _
                 UCase("TrCstDt"), UCase("TrRecdMetWt"), UCase("TrCastLoss"), UCase("TrCutLoss"), _
                 UCase("TrScrap"), UCase("TrGrindLoss"), UCase("TrStripLoss"), UCase("TrGrsPcWt"), _
                 UCase("TrDustWt"), UCase("WTRDDESC"), UCase("TrDesc")
          Cancel = True: ErrMsg = "Cannot Edit Value,When Receipt Voucher Exists": Exit Sub
     End Select
  End If
  
 Select Case UCase(IdName)
  Case Is = UCase("TrYmd")
   Call HlpList.TrYmd(ADC("TrCoCd"))
  Case Is = UCase("TrChr")
    Call HlpList.PMCd("TRCHR")
  Case Is = UCase("TrNo")
    Call HlpList.TrNo(ADC("TrCoCd"), ADC("TrYmd"), ADC("TrChr"))
  Case Is = UCase("TrRmCd")
    Call HlpList.RmCd("N", "'G','P','S','L','M'")
  Case Is = UCase("TrCol")
    Call HlpList.PMCd("DMCOL")
  Case Is = UCase("wCbBLoc")
    Call HlpList.Loc(ADC("TrCoCd"), "'P'")
  Case Is = UCase("wCbBYyFr"), UCase("wCbBYyTo")
    Call HlpList.PMCd("YY")
  Case Is = UCase("wCbBChrFr"), UCase("wCbBChrTo")
    Call HlpList.PMCd("BCHR")
  Case Is = UCase("wCbBNoFr")
    wOthLocBag = moCn.GetFldVal("Select hOthLocBag from Head where HCoCd= '" + ADC("TrCoCd") + "' and HCd='" + ctSelfCmCd + "'")
    Call HlpList.BNo(ADC("TrCoCd"), ADC("wCbBYyFr"), ADC("wCbBChrFr"), wOthLocBag, ADC("wCbBLoc"))
  Case Is = UCase("wCbBNoTo")
    wOthLocBag = moCn.GetFldVal("Select hOthLocBag from Head where HCoCd= '" + ADC("TrCoCd") + "' and HCd='" + ctSelfCmCd + "'")
    Call HlpList.BNo(ADC("TrCoCd"), ADC("wCbBYyTo"), ADC("wCbBChrTo"), wOthLocBag)
  Case Is = UCase("wCbOdTcFr"), UCase("wCbOdTcTo")
    Call HlpList.PMCd("TC", "'SO'")
  Case Is = UCase("wCbOdYyFr"), UCase("wCbOdYyTo")
    Call HlpList.PMCd("YY")
  Case Is = UCase("wCbOdChrFr")
    Call HlpList.vPSCd(ADC("TrCoCd"), "CHR", ADC("wCbOdTcFr"))
  Case Is = UCase("wCbOdChrTo")
    Call HlpList.vPSCd(ADC("TrCoCd"), "CHR", ADC("wCbOdTcTo"))
  Case Is = UCase("wCbOdNoFr")
    Call HlpList.OmNo(ADC("TrCoCd"), ADC("wCbOdTcFr"), ADC("wCbOdYyFr"), ADC("wCbOdChrFr"))
  Case Is = UCase("wCbOdNoTo")
    Call HlpList.OmNo(ADC("TrCoCd"), ADC("wCbOdTcFr"), ADC("wCbOdYyFr"), ADC("wCbOdChrFr"))
  Case Is = UCase("wCbOdSrFr")
    Call HlpList.OdSr(ADC("TrCoCd"), ADC("wCbOdTcFr"), ADC("wCbOdYyFr"), ADC("wCbOdChrFr"), ADC("wCbOdNoFr"))
  Case Is = UCase("wCbOdSrTo")
    Call HlpList.OdSr(ADC("TrCoCd"), ADC("wCbOdTcTo"), ADC("wCbOdYyTo"), ADC("wCbOdChrTo"), ADC("wCbOdNoTo"))
  Case Is = UCase("wCbDmCdFr"), UCase("wCbDmCdTo")
    Call HlpList.DmCd("DM")
  End Select
End Sub

Private Sub atxt_ReCalculate(Index As Integer, ByVal IdName As String, ByVal OldVal As Variant, ByVal RowNum As Integer, ByVal SourceName As String, ByVal SourceColName As String)
  Dim i As Integer
  
  Select Case UCase(IdName)
  Case Is = UCase("TrUsdMetWt")
    If ADC.Mode = xNorm Then _
       ADC("TrUsdMetWt") = ADC("WTrGrsMetReq") - ADC("TrFrMetWt")
  Case Is = UCase("wTrDiff")
    '*** (Jen 3.2.0)
    ' ***** Manali 3.9.0 - xInit added below
    If ADC.Mode = xNorm Or ADC.Mode = xCopy Or ADC.Mode = xInit Then
    '*** (Jen 3.2.0)
    ' ****** Manali 350Nxt - TrNetPcWt used for Calculations instead of TrGrsPcWt
      ADC("wTrDiff") = ADC("TrFrMetWt") + ADC("TrUsdMetWt") - (ADC("TrRecdMetWt") + ADC("TrCastLoss") + _
                     ADC("TrCutLoss") + ADC("TrGrindLoss") + ADC("TrScrap") + ADC("TrStripLoss") + ADC("TrNetPcWt"))
                     '''+ ADC("TrGrsPcWt"))
    '*** (Jen 3.2.0)
    End If
    '*** (Jen 3.2.0)
  
  ' ****** Manali 350Nxt - TrNetPcWt
  Case Is = UCase("TrNetPcWt")
    If ADC.Mode = xNorm Then
      ' ***** Manali 3.8.0 - MetX Wt Deducted in calculation
      If ADC("TrGrsPcWt") <> 0 Then
        ADC("TrNetPcWt") = ADC("TrGrsPcWt") - (MWLib.Div(ADC("wTotDCWt"), 5) + ADC("wTotMetXWt"))
      Else
        ADC("TrNetPcWt") = 0
      End If
    End If
  ' ****** Manali 350Nxt - TrNetPcWt
    
  End Select
End Sub
Private Sub ATXT_FldChange(Index As Integer, ByVal IdName As String, Cancel As Boolean, ErrMsg As String, ByVal pv_NewValue As Variant)
  '*** TrYmd        Not Blank
  '*** TrChr        PMCd('TrCHR')
  '*** TrDt         Not Blank
  '*** TrRmCd       Valid RmCd from RmMst
  Select Case UCase(IdName)
  Case Is = UCase("TrYmd")
    If pv_NewValue = "" Then Cancel = True: ErrMsg = "Transaction YyMmDd Cannot Be Blank": Exit Sub
    If Not IsNumeric(pv_NewValue) Then Cancel = True: ErrMsg = "Transaction YyMmDd Must Be a Numeric Value": Exit Sub
    If Mid(pv_NewValue, 1, 2) < "00" Or Mid(pv_NewValue, 1, 2) > "99" Then Cancel = True: ErrMsg = "Year of YyMmDd  must be in between(00-99)": Exit Sub
    If Mid(pv_NewValue, 3, 2) <> "" And (Mid(pv_NewValue, 3, 2) < "01" Or Mid(pv_NewValue, 3, 2) > "12") Then Cancel = True: ErrMsg = "Month of YyMmDd must be in between(01-12)": Exit Sub
    If Mid(pv_NewValue, 5, 2) <> "" And (Mid(pv_NewValue, 5, 2) < "01" Or Mid(pv_NewValue, 5, 2) > "31") Then Cancel = True: ErrMsg = "Date of YyMmDd must be in between(01-31)": Exit Sub
   ' If Not IsDate("20" + Mid(pv_NewValue, 1, 2) + "/" + Mid(pv_NewValue, 3, 2) + "/" + Mid(pv_NewValue, 5, 2)) Then Cancel = True: ErrMsg = "Invalid Date": Exit Sub
  Case Is = UCase("TrChr")
    Cancel = (pv_NewValue <> "") And (Not moCn.RecSeek("select PMCd from Param where " + _
                         "PTyp= 'TrCHR' and PMCd = '" + pv_NewValue + "'"))
    ErrMsg = "Invalid Voucher Character": Exit Sub
  Case Is = UCase("TrDt")
    If pv_NewValue = "" Then Cancel = True: ErrMsg = "Transaction Date Cannot Be Blank": Exit Sub
    Dim wPurgeDt As String
    wPurgeDt = GetPurgeDt(gs_CoCd, "")
    If pv_NewValue < CDate(wPurgeDt) Then Cancel = True: ErrMsg = "Transaction Date Cannot Be Less Than Purge Date": Exit Sub
'    If pv_NewValue < CDate(ms_LockDt) Then Cancel = True: ErrMsg = "Transaction Date Cannot Be Less Than Lock Date": Exit Sub
'    If pv_NewValue > CDate(ms_FutureDt) Then Cancel = True: ErrMsg = "Transaction Date Cannot Be Greater Than Future Date": Exit Sub
    If moCn.GetFldVal("Select hHolChkYN From Head where HCd='ZSELF' and HCoCd='" + gs_CoCd + "'") = "Y" Then
        '53 company code added
        'If moCn.RecSeek("Select top 1 From DtTbl where DtTag='HOL' and DtCoCd in('" + ctSelfCoCd + "','" + ADC("TCoCd") + "')" + _
          " and convert(varchar(15),DtDt,3)= '" + CStr(pv_NewValue) + "' and DtHolYN='Y'")
        If IsHOLYN(ADC("TrCoCd"), pv_NewValue) Then Cancel = True: ErrMsg = "Voucher Cannot be Added on Holidays": Exit Sub
    End If
  Case Is = UCase("TrRmCd")
    'Sachin 3.11.0 - 'M' added below
    Cancel = (Not moCn.RecSeek("Select RmCd from RmMst where " + IIF(gs_Partition = ctCurrPrtn, " RmPrtKey='" + ctCurrPrtn + "' And ", "") + _
                                          "RmCtg in ('G','P','S','L','M') and RmCd= '" + pv_NewValue + "' and RmValidYN='Y' "))
                                          
        'If Cancel = True Then ErrMsg = "Invalid Tree Kt ": Exit Sub
  Case Is = UCase("TrCol")
    Cancel = (Not moCn.RecSeek("Select PMCd from Param  where " + _
                                          "PTyp='DMCOL' and PmCd= '" + pv_NewValue + "'"))
                                          
        If Cancel = True Then ErrMsg = "Invalid Tree Color ": Exit Sub
  'IG -Cant enter holiday as casting date
  Case Is = UCase("TrCstDt")
    If moCn.GetFldVal("Select hHolChkYN From Head where HCd='ZSELF' and HCoCd='" + gs_CoCd + "'") = "Y" Then
      If IsHOLYN(ADC("TrCoCd"), pv_NewValue) Then Cancel = True: ErrMsg = "Voucher Cannot be Added on Holidays": Exit Sub
    End If
    Call GetTotals     '578-33 calculating total
        
'        If IsHOLYN(ADC("TrCoCd"), pv_NewValue) Then Cancel = True:
        
 '       If moCn.RecSeek("Select 'x' From DtTbl where DtTag='HOL' and DtCoCd in('" + ctSelfCoCd + "','" + ADC("TrCoCd") + "')" + _
          " and convert(varchar(15),DtDt,3)= '" + CStr(pv_NewValue) + "' and DtHolYN='Y'") Then Cancel = True: ErrMsg = "Casting Date Cannot be Added on Holidays": Exit Sub
  '  End If
    
  ' ****** Manali 350Nxt - Check added for TrNetPcWt > 0
''  Case Is = UCase("TrGrsPcWt")
''    ' **** Manali 3.8.0 - wTotMetXWt added in check
''    Cancel = ((pv_NewValue - (MWLib.Div(adc("wTotDCWt"), 5)) + adc("wTotMetXWt")) < 0)
''    If Cancel = True Then ErrMsg = "Gross Wt Cannot Be Less Than Total DC + Metal Wt": Exit Sub
  ' ****** Manali 350Nxt - Check added for TrNetPcWt > 0
  
  
 End Select
End Sub

Private Sub CmdCB_Click()
  If Not ADC.SaveRec Then
    DispMsg "X Save Failed", etInfo
    Exit Sub
  End If
  Call DispFra(CpyBag)
End Sub

Private Sub CmdCBGo_Click()
' If ADC("wCbBYyFr") = "" And ADC("wCbBChrFr") = "" And ADC("wCbBNoFr") = "" And _
  ADC("wCbBYyTo") = "" And ADC("wCbBChrTo") = "" And ADC("wCbBNoTo") = "" And _
  ADC("wCbOdTcFr") = "" And ADC("wCbOdYyFr") = "" And ADC("wCbOdChrFr") = "" And _
  ADC("wCbOdNoFr") = "" And ADC("wCbOdSrFr") = "" And _
  ADC("wCbOdTcTo") = "" And ADC("wCbOdYyTo") = "" And ADC("wCbOdChrTo") = "" And _
  ADC("wCbOdNoTo") = 0 And ADC("wCbOdSrTo") = 0 And _
  ADC("wCbDmCdFr") = "" And ADC("wCbDmCdTo") = "" And ADC("wCbBLoc") = "" Then _
  DispMsg "Enter the scope ", etError: Exit Sub
  
  'IG.213 - if issue/receipt voucher exists cant add new row to grid while copy bags
  Call TrIRVch_Check("IR")
    If mb_AddRow = True Then
      Call CopyBag
    End If
    Call DispFra(CpyBag)
End Sub
'*** Jay 3.2.0 [AutoTree]
Private Sub CmdIss_Click()
  If TrIRVch_Check("I") <> True Then
    InsertRecords True
  End If
End Sub
'*** Jay 3.2.0 [AutoTree]
Private Sub CmdRec_Click()
  If TrIRVch_Check("R") <> True Then
    InsertRecords False
  End If
End Sub
'578-33 showing totals in a new frame
Private Sub CmdShowTot_Click()
  If Not ADC.SaveRec Then
    DispMsg "X Save Failed", etInfo
    Exit Sub
  End If
     
  Call GetTotals
  Call DispFra(ShowTot)
  
End Sub
'578-33
Private Sub CmdShowTotOk_Click()
  Call DispFra(ShowTot)
End Sub

Private Sub GrdTreeDet_RowWhen(ByVal RowNum As Integer)
  '*** If GrdTreeDet.IsNew Then TdSr = Max + 1
  If GrdTreeDet.IsNew(RowNum) Then GrdTreeDet.Value(RowNum, "TrdSr") = GrdTreeDet.MaxVal("TrdSr") + 1
  ADC("wTrdDesc") = GrdTreeDet.Value(RowNum, "TrdDesc")   ' ***** Manali 3.8.0 - Description at Bag Level
  'IG -if issue/receipt voucher exists cant add data to grid
  Call TrIRVch_Check("IR")
End Sub
Private Sub GrdTreeDet_ColWhen(ByVal RowNum As Integer, ByVal ColName As String, Cancel As Boolean, ErrMsg As String, ByVal Action As MwfCtl.en_fgAction)
  '*** Set Helps For Each Field
  '*** If 1st row, then get the default values for TrdBYy and TrdBChr from the Head file
      'If Row> 1 then get the defa values for TrdBYy and TrdBChr from the previous row of the grid
      'If not GrdTreeDet.IsNew the cannot edit TrdBYy, TrdBChr, TrdBNo
  
 ' If mf_DtLocked = True Then Cancel = True: ErrMsg = "Cannot Edit Locked Entry": Exit Sub
  With GrdTreeDet
    'IG -If issue/receipt voucher exists, cant change values in Grid
    If ADC("TrIssVchNo") <> "" Or ADC("TrRecVchNo") <> "" Then
       Select Case ColName
         Case Is = UCase("TrdBYy"), UCase("TrdBChr"), UCase("TrdBNo"), UCase("wBagNo"), _
                   UCase("TrdBagPcs"), UCase("TrdDcWt"), UCase("TrdMetXWt")
            Cancel = True: ErrMsg = "Cannot Edit Value,When Issue/Receipt Voucher Exists": Exit Sub
       End Select
    End If
  
    Select Case ColName
    '********Bhavna added bag no barcode
    Case Is = UCase("wBagNo")
    If Not (.IsNew(RowNum)) Then
      Cancel = True: ErrMsg = "Cannot Edit Bag No": Exit Sub
    End If
    '********Bhavna added bag no barcode
    
    Case Is = UCase("TrdBYy")
     If .Value(RowNum, "TrdBYy") = "" Then
        If RowNum = 1 Then .Value(RowNum, "TrdBYy") = _
           moCn.GetFldVal("Select hYy from Head where HCoCd= '" + ADC("TrCoCd") + "' and HCd='" + ctSelfCmCd + "'")
        If RowNum > 1 Then .Value(RowNum, "TrdBYy") = .Value(RowNum - 1, "TrdBYy")
     ElseIf Not (.IsNew(RowNum)) Or .Value(RowNum, "WBAGNO") <> "" Then
      Cancel = True: ErrMsg = "Cannot Edit Bag Year": Exit Sub
     End If
     Call HlpList.PMCd("YY")
    Case Is = UCase("TrdBChr")
      If .Value(RowNum, "TrdBChr") = "" Then
        If RowNum = 1 Then .Value(RowNum, "TrdBChr") = _
           moCn.GetFldVal("Select hBChr from Head where HCoCd= '" + ADC("TrCoCd") + "' and HCd='" + ctSelfCmCd + "'")
        If RowNum > 1 Then .Value(RowNum, "TrdBChr") = .Value(RowNum - 1, "TrdBChr")
      ElseIf Not (.IsNew(RowNum)) Or .Value(RowNum, "WBAGNO") <> "" Then
       Cancel = True: ErrMsg = "Cannot Edit Bag Character": Exit Sub
      End If
      Call HlpList.PMCd("BCHR")
    Case Is = UCase("TrdBNo")
      If Not (.IsNew(RowNum)) Or .Value(RowNum, "WBAGNO") <> "" Then _
        Cancel = True: ErrMsg = "Cannot Edit Bag No": Exit Sub
      Dim wOthLocBag As String
      wOthLocBag = moCn.GetFldVal("Select hOthLocBag from Head where HCoCd= '" + ADC("TrCoCd") + "' and HCd='" + ctSelfCmCd + "'")
      Call HlpList.BNo(ADC("TrCoCd"), .Value(RowNum, "TrdBYy"), .Value(RowNum, "TrdBChr"), wOthLocBag)
    Case Is = UCase("TrdBagPcs"), UCase("TrdDcWt"), UCase("TrdMetXWt")    ' ***** Manali 3.8.0
      If ADC("TrCstDt") <> MWLib.EmptyDate Then _
        Cancel = True: ErrMsg = "Cannot Edit Value,When Casting Details Entered ": Exit Sub
    End Select
End With
End Sub
Private Sub GrdTreeDet_ReCalc(ByVal RowNum As Integer, ByVal ColName As String, ByVal OldVal As Variant, ByVal SourceName As String, ByVal SourceColName As String)
  With GrdTreeDet
    Select Case UCase(ColName)
    Case Is = UCase("TrdBNo")
      If .Mode = fgmnorm Then .Value(RowNum, "TrdBNo") = 0
    Case Is = UCase("TrdBagPcs")
      If .Mode = fgmnorm Then
     .Value(RowNum, "TrdBagPcs") = moCn.GetFldVal("Select BQty from Bag where " + IIF(gs_Partition = ctCurrPrtn, " BPrtKey='" + ctCurrPrtn + "' And ", "") + _
                                     "BCoCd='" + ADC("TrCoCd") + "' and " + _
                                     "BYy='" + .Value(RowNum, "TrdBYy") + "' and " + _
                                     "BChr='" + .Value(RowNum, "TrdBChr") + "' and " + _
                                     "BNo= " + CStr(.Value(RowNum, "TrdBNo")))
                                                
      End If
    Case Is = UCase("TrdDCWt")
      Dim wd_DcWt As Double
      Dim wRs_BgRm As MDORowSet
      If .Mode = fgmnorm Or .Mode = fgmCopy Then
        ' ***** Manali 3.6.0 - 12/02/10 - decmal changed for wd_DcWt (10, 3) -> (14, 6)
        ' ***** Manali 3.8.0 - if BQty = 0 Then no Proportionating
         Set wRs_BgRm = moCn.OpenRes("Select IsNull(max(BQty),0) as BQty, " + _
                                                  "isnull(cast((sum((Case When TdRmCtg In ('D','C') Then 1 Else 0 End) * TdRmWt* " + _
                                                  " (case when TdFrRmLoc= 'W' and TdFrRmDC= 'C' then -1 " + _
                                                   "when TdFrRmLoc= 'W' and TdFrRmDC= 'D' then 1 " + _
                                                   "when TdToRmLoc= 'W' and TdToRmDC= 'C' then -1 " + _
                                                   "when  TdToRmLoc= 'W' and TdToRmDC= 'D' then 1 else 0 end))/ (Case When max(BQty)=0 Then 1 Else max(BQty) End)) as decimal(14,6)),0) as qDcWt, " + _
                                                  "isnull(cast((sum((Case When TdRmCtg Not In ('D','C') Then 1 Else 0 End) * TdRmWt* " + _
                                                  " (case when TdFrRmLoc= 'W' and TdFrRmDC= 'C' then -1 " + _
                                                   "when TdFrRmLoc= 'W' and TdFrRmDC= 'D' then 1 " + _
                                                   "when TdToRmLoc= 'W' and TdToRmDC= 'C' then -1 " + _
                                                   "when  TdToRmLoc= 'W' and TdToRmDC= 'D' then 1 else 0 end))/(Case When max(BQty)=0 Then 1 Else max(BQty) End)) as decimal(14,6)),0) as qMetXWt " + _
                                                   "from Bag join Txnd  " + _
                                                   "on BPrtKey=TdPrtKey and BIdNo=TdBIdNo And BCoCd= TdCoCd and BYy= TdBYy and BChr= TdBChr  and BNo= TdBNo  " + _
                                                   " and (BQty>0 or BPtQty>0) and BCls='N' and BCoCd='" + ADC("TrCoCd") + "' and " + _
                                                    "BYy='" + .Value(RowNum, "TrdBYy") + "' and " + _
                                                    "BChr='" + .Value(RowNum, "TrdBChr") + "' and " + _
                                                    "BNo= " + CStr(.Value(RowNum, "TrdBNo")))
        Do While Not (wRs_BgRm.EOF Or wRs_BgRm.BOF)
          .Value(RowNum, "TrdDCWt") = wRs_BgRm!qDcWt * IIF(wRs_BgRm!BQty = 0 And .Value(RowNum, "TrdBagPcs") = 0, 1, .Value(RowNum, "TrdBagPcs"))
          .Value(RowNum, "TrdMetXwt") = wRs_BgRm!qMetXWt * IIF(wRs_BgRm!BQty = 0 And .Value(RowNum, "TrdBagPcs") = 0, 1, .Value(RowNum, "TrdBagPcs"))
          wRs_BgRm.MoveNext
        Loop
                                                    
        
''         wd_DcWt = moCn.GetFldVal("Select isnull(cast(sum(TdRmWt* case when TdFrRmLoc= 'W' and TdFrRmDC= 'C' then -1 " + _
''                                                   "when TdFrRmLoc= 'W' and TdFrRmDC= 'D' then 1 " + _
''                                                   "when TdToRmLoc= 'W' and TdToRmDC= 'C' then -1 " + _
''                                                   "when  TdToRmLoc= 'W' and TdToRmDC= 'D' then 1 else 0 end)/max(BQty) as decimal(14,6)),0) " + _
''                                                   "from Bag join Txnd  " + _
''                                                   "on BPrtKey=TdPrtKey and BIdNo=TdBIdNo And BCoCd= TdCoCd and BYy= TdBYy and BChr= TdBChr  and BNo= TdBNo  " + _
''                                                    "join RmMst on RmPrtKey=TdPrtKey And RmCd=TdRmCd where " + IIF(gs_Partition = ctCurrPrtn, " BPrtKey='" + ctCurrPrtn + "' And ", "") + " BQty>0 and BCls='N' and RmCtg in ('D','C') and BCoCd='" + adc("TrCoCd") + "' and " + _
''                                                    "BYy='" + .Value(RowNum, "TrdBYy") + "' and " + _
''                                                    "BChr='" + .Value(RowNum, "TrdBChr") + "' and " + _
''                                                    "BNo= " + CStr(.Value(RowNum, "TrdBNo")))
''        .Value(RowNum, "TrdDCWt") = wd_DcWt * .Value(RowNum, "TrdBagPcs")
      End If
    Case Is = UCase("wDmCd")
      Dim wo_rsSrcOrd As MDORowSet
        If ADC.Mode = xNorm And (.Mode = fgminit Or .Mode = fgmnorm Or .Mode = fgmCopy) Then
          Set wo_rsSrcOrd = moCn.OpenRes("select OdDmCd,OdDmCol,OdTc+'/'+OdYy+'/'+OdChr+'/'+LTrim(Str(OdNo))+'/'+LTrim(Str(OdSr)) as qOrdNo, " + _
                                       "OmCmCd,OdKt from Bag " + _
                                       "join OrdDsg on BPrtKey=OdPrtKey And BOdIdNo=OdIdNo And BCocd=OdCoCd and BOdTc=OdTc and BOdYy=OdYy " + _
                                       "and BOdChr=OdChr and BOdNo=OdNo and BOdSr=OdSr " + _
                                       "join OrdMst on OmPrtKey=OdPrtKey And OmIdNo=OdOmIdNo And OmCoCd=OdCoCd and OmTc=OdTc and OmYy=OdYy and OmChr=OdChr " + _
                                       "and OmNo=OdNo where " + IIF(gs_Partition = ctCurrPrtn, " BPrtKey='" + ctCurrPrtn + "' And ", "") + " BCoCd='" + ADC("TrCoCd") + "' and " + _
                                                    "BYy='" + .Value(RowNum, "TrdBYy") + "' and " + _
                                                    "BChr='" + .Value(RowNum, "TrdBChr") + "' and " + _
                                                    "BNo= " + CStr(.Value(RowNum, "TrdBNo")))
        Do While Not (wo_rsSrcOrd.EOF Or wo_rsSrcOrd.BOF)
         .Value(RowNum, "wDmCd") = wo_rsSrcOrd!OdDmCd
         .Value(RowNum, "wDmCol") = wo_rsSrcOrd!OdDmCol
         .Value(RowNum, "wDmKt") = wo_rsSrcOrd!OdKt
         .Value(RowNum, "wOrdNo") = wo_rsSrcOrd!qOrdNo
         .Value(RowNum, "wCmcd") = wo_rsSrcOrd!OmCmCd
         wo_rsSrcOrd.MoveNext
        Loop
      End If
      
      
    Case Is = UCase("TrdMetXwt")
    ' ***** Manali 3.8.0 - Metal and X Wt
'      Dim wd_MetXWt As Double
'      If .Mode = fgmnorm Or .Mode = fgmCopy Then
'         wd_MetXWt = moCn.GetFldVal("Select isnull(cast(sum(TdRmWt* case when TdFrRmLoc= 'W' and TdFrRmDC= 'C' then -1 " + _
'                                                   "when TdFrRmLoc= 'W' and TdFrRmDC= 'D' then 1 " + _
'                                                   "when TdToRmLoc= 'W' and TdToRmDC= 'C' then -1 " + _
'                                                   "when  TdToRmLoc= 'W' and TdToRmDC= 'D' then 1 else 0 end)/" + IIF(.Value(RowNum, "TrdBagPcs") = 0, "1", "max(BQty)") + " as decimal(14,6)),0) " + _
'                                                   "from Bag join Txnd  " + _
'                                                   "on BPrtKey=TdPrtKey and BIdNo=TdBIdNo And BCoCd= TdCoCd and BYy= TdBYy and BChr= TdBChr  and BNo= TdBNo  " + _
'                                                    "join RmMst on RmPrtKey=TdPrtKey And RmCd=TdRmCd where " + IIF(gs_Partition = ctCurrPrtn, " BPrtKey='" + ctCurrPrtn + "' And ", "") + " BQty>0 and BCls='N' and RmCtg Not in ('D','C') and BCoCd='" + adc("TrCoCd") + "' and " + _
'                                                    "BYy='" + .Value(RowNum, "TrdBYy") + "' and " + _
'                                                    "BChr='" + .Value(RowNum, "TrdBChr") + "' and " + _
'                                                    "BNo= " + CStr(.Value(RowNum, "TrdBNo")))
'        .Value(RowNum, "TrdMetXwt") = wd_MetXWt * IIF(.Value(RowNum, "TrdBagPcs") = 0, 1, .Value(RowNum, "TrdBagPcs"))
'      End If
      
    Case Is = UCase("TrdDesc")
    ' ***** Manali 3.8.0 - Desc at Bag Level
      If .Mode = fgmnorm Then
        If .Row > 0 Then
          If .Value(.Row, "TrdDesc") <> ADC("wTrdDesc") Then
            .Value(.Row, "TrdDesc") = ADC("wTrdDesc")
            .SaveRec .Row
          End If
        End If
      End If
      
     End Select
  End With
End Sub
Private Sub GrdTreeDet_ColChange(ByVal RowNum As Integer, ByVal ColName As String, ByVal NewValue As Variant, Cancel As Boolean, ErrMsg As String)
  '*** TrdBYy            valid PMCd('YY') or ''
  '*** TrdBChr           valid PMCd('BCHR') or ''
  '*** TrdBNo            valid BNo from Bag(BYy= TdBYy, BChr= TdBChr) or 0
  Dim ws As String
  '*** Bhavna added bag no barcode
  Dim ws_BNo() As String
  Dim ws_OldYy As String, ws_OldChr As String, ws_OldNo As Single
  '*** Bhavna added bag no barcode
  
  With GrdTreeDet
    Select Case ColName
    '*********Bhavna added bag no using barcode reader
    Case Is = UCase("WBagNo")
      If NewValue <> "" Then
        ws_BNo = Split(NewValue, "/")
        If UBound(ws_BNo) <> 2 Then Cancel = True: ErrMsg = "Invalid Bag No": Exit Sub
        If Not IsNumeric(ws_BNo(2)) Then Cancel = True: ErrMsg = "Invalid Bag No": Exit Sub
        
        ws_OldYy = .Value(RowNum, "TrdBYY")
        ws_OldChr = .Value(RowNum, "TrdBCHR")
        ws_OldNo = .Value(RowNum, "TrdBNO")
        
        .Value(RowNum, "TrdBYY") = UCase(ws_BNo(0))
        Call GrdTreeDet_ColChange(RowNum, "TRDBYY", ws_BNo(0), Cancel, "")
        If Cancel = True Then ErrMsg = "Invalid Bag Year": .Value(RowNum, "TRDBYY") = ws_OldYy: Exit Sub
        
        .Value(RowNum, "TRDBCHR") = UCase(ws_BNo(1))
        Call GrdTreeDet_ColChange(RowNum, "TRDBCHR", ws_BNo(1), Cancel, "")
        If Cancel = True Then ErrMsg = "Invalid Bag Character ": .Value(RowNum, "TRDBCHR") = ws_OldChr: Exit Sub
        
        .Value(RowNum, "TRDBNO") = ws_BNo(2)
        Call GrdTreeDet_ColChange(RowNum, "TRDBNO", ws_BNo(2), Cancel, "")
        If Cancel = True Then ErrMsg = "Invalid Bag Number": .Value(RowNum, "TRDBNO") = ws_OldNo: Exit Sub
        .Col = 6
        
'        If .Rows - 1 = RowNum Then
'          .AddItem
'          .Row = RowNum + 1
'          .Col = 6
'        End If
        
        'HK.2- Adding a new row automatically
        If mb_AddRow Then
          If .Rows - 1 = RowNum Then
             .AddItem
             .Row = RowNum + 1
             .Col = 2
          End If
        End If

      End If
    
    Case Is = UCase("TrdBYy")
      Cancel = (NewValue <> "") And (Not moCn.RecSeek("Select PMCd From Param Where PTyp= 'YY' And " + _
                                     "PMCd = '" + NewValue + "'"))
      ErrMsg = "Invalid Bag Year": Exit Sub
    Case Is = UCase("TrdBChr")
      Cancel = (NewValue <> "") And (Not moCn.RecSeek("Select PMCd From Param Where PTyp= 'BCHR' And " + _
                                     "PMCd = '" + NewValue + "'"))
      ErrMsg = "Invalid Bag Character": Exit Sub
    
    Case Is = UCase("TrdBNo")
      Cancel = (NewValue <> 0) And (Not moCn.RecSeek("Select BNo From Bag Where " + IIF(gs_Partition = ctCurrPrtn, " BPrtKey='" + ctCurrPrtn + "' And ", "") + _
                                    "BCoCd='" + ADC("TrCoCd") + "' and " + _
                                    "BYy= '" + .Value(RowNum, "TrdBYy") + "' And " + _
                                    "BChr = '" + .Value(RowNum, "TrdBChr") + "' And " + _
                                    "BNo= " + CStr(NewValue) + " and BCls='N' and BLoc<>'PFG' and (BQty>0 or BptQty>0)"))
      If Cancel = True Then ErrMsg = "Invalid Bag Number": Exit Sub
      
      ' ***** Manali 3.6.0 - 01/09/09 - Join with OrdRm added - OrRmCd checked against TrRmCd
      Cancel = (NewValue <> 0) And (Not moCn.RecSeek("Select '*' from OrdRm Join Bag " + _
                                    "On OrCoCd=BCoCd and OrTC=BOdTc and OrYy=BOdYy and OrChr=BOdChr and OrNo=BOdNo and OrSr=BOdSr " + _
                                    "Where " + IIF(gs_Partition = ctCurrPrtn, " BPrtKey='" + ctCurrPrtn + "' And ", "") + _
                                    "BCoCd='" + ADC("TrCoCd") + "' and " + _
                                    "BYy= '" + .Value(RowNum, "TrdBYy") + "' And " + _
                                    "BChr = '" + .Value(RowNum, "TrdBChr") + "' And " + _
                                    "BNo= " + CStr(NewValue) + " and OrRmCd='" + ADC("TrRmcd") + "' "))
      If Cancel = True Then ErrMsg = "Order RmCd of Bag not Matching with Tree RmCd": Exit Sub
      
   ' Case Is = UCase("TrdBagPcs")
     '   Cancel = .Value(RowNum, "TrdBNo") <> 0 And (moCn.RecSeek("Select BQty From Bag Where " + _
                                    "BCoCd='" + ADC("TrCoCd") + "' and " + _
                                    "BYy= '" + .Value(RowNum, "TrdBYy") + "' And " + _
                                    "BChr = '" + .Value(RowNum, "TrdBChr") + "' And " + _
                                    "BNo= " + CStr(.Value(RowNum, "TrdBNo")) + " and BQty< " + CStr(NewValue)))
      ' If Cancel = True Then ErrMsg = "Bag Peices cannot be greater than Bag Quantity": Exit Sub
    End Select
  End With
End Sub
Private Sub GrdTreeDet_InitKey(ByVal RowNum As Integer)
  '*** Initialise Key Fields and all the fields which are not included in the grid
 With GrdTreeDet
    .Store "TrdCoCd", ADC("TrCoCd")
    .Store "TrdYmd", ADC("TrYmd")
    .Store "TrdChr", ADC("TrChr")
    .Store "TrdNo", ADC("TrNo")
  End With
End Sub
Private Sub ADC_SetRecSource()
  '*** Set the RecSource for ADC (Table TreeMst)
  ADC.RecSource = " Select * from TreeMst " + _
                  " Where TrCoCd ='" + ADC("TrCoCd") + "'" + _
                  " And TrYmd = '" + ADC("TrYmd") + "'" + _
                  " And TrChr ='" + ADC("TrChr") + "' " + _
                  " And TrNo =" + CStr(ADC("TrNo"))
  
End Sub
Private Sub GrdTreeDet_SetRecSource()
  '*** Set the RecSource for GrdTreeDet (Table TreeDet )
    GrdTreeDet.RecSource = " Select * from TreeDet " + _
                         " Where TrdCoCd ='" + ADC("TrCoCd") + "'" + _
                         " And TrdYmd = '" + ADC("TrYmd") + "'" + _
                         " And TrdChr ='" + ADC("TrChr") + "' " + _
                         " And TrdNo =" + CStr(ADC("TrNo")) + " Order By TrdSr"
 
 End Sub
Private Sub adc_AutoGen()
  
  '*** write code for AutoGenerating the Transaction Voucher Number
  If ADC("TrNo") = 0 Then
    If moCn.RecSeek("select * from TreeMst where TrCoCd='" + ADC("TrCoCd") + "' and TrYmd='" + ADC("TrYmd") + "'" + _
                                " and TrChr='" + ADC("TrChr") + "'") Then
            ADC("TrNo") = moCn.GetFldVal("select max(TrNo)+1 from TreeMst where TrCoCd='" + ADC("TrCoCd") + "' and TrYmd='" + ADC("TrYmd") + "'" + _
                                        " and TrChr='" + ADC("TrChr") + "'")
    Else
     ADC("TrNo") = 1
    End If
  End If
End Sub
Private Function ColSty(ByVal mColSty As String) As en_fgColStyle
  '*** Function to return the Style of columns in the Grid
  Select Case mColSty
    Case Is = "N"
      ColSty = en_fgColStyle.fgcNormal
    Case Is = "I"
      ColSty = en_fgColStyle.fgcinvisible
    Case Is = "R"
      ColSty = en_fgColStyle.fgcReadOnly
  End Select
End Function
Private Function FldEna(ByVal mEna As String) As Boolean
  '*** Function to return the Ena/Disa state of the controls on the form
  Select Case mEna
    Case Is = "E", "B"
      FldEna = True
    Case Else
      FldEna = False
  End Select
End Function
Private Function FldVis(ByVal mVis As String) As Boolean
  '*** Function to return the Visible property of the controls on the form
  Select Case mVis
    Case Is = "V", "B"
      FldVis = True
    Case Else
      FldVis = False
  End Select
End Function
Private Sub CopyBag()
  '*** Procedure to Copy Bags
  Dim wRsCB As MwfLib.MDORowSet, wTotRows As Single, wRow As Single, wCnd As String
  Dim wRfIdStr As String, wRfIdNoArr() As String, wRfIdNo As String, i As Integer, wTotRfid As Integer
  Dim wMinSr As Integer
  
  ADC("wCbBYyFr").CmpStr = "BYy>=": ADC("wCbBChrFr").CmpStr = "BChr>=": ADC("wCbBNoFr").CmpStr = "BNo>="
  ADC("wCbBYyTo").CmpStr = "BYy<=": ADC("wCbBChrTo").CmpStr = "BChr<=": ADC("wCbBNoTo").CmpStr = "BNo<="
  ADC("wCbOdTcFr").CmpStr = "BOdTc>=": ADC("wCbOdYyFr").CmpStr = "BOdYy>=": ADC("wCbOdChrFr").CmpStr = "BOdChr>=": ADC("wCbOdNoFr").CmpStr = "BOdNo>=": ADC("wCbOdSrFr").CmpStr = "BOdSr>="
  ADC("wCbOdTcTo").CmpStr = "BOdTc<=": ADC("wCbOdYyTo").CmpStr = "BOdYy<=": ADC("wCbOdChrTo").CmpStr = "BOdChr<=": ADC("wCbOdNoTo").CmpStr = "BOdNo<=": ADC("wCbOdSrTo").CmpStr = "BOdSr<="
  ADC("wCbDmCdFr").CmpStr = "BOdDmCd>=": ADC("wCbDmCdTo").CmpStr = "BOdDmCd<="
  ADC("wCbBLoc").CmpStr = "BLoc="
  ADC("wCbOdKtFr").CmpStr = "OdKt>=": ADC("wCbOdKtTo").CmpStr = "OdKt<="
  ADC("wCbOdColFr").CmpStr = "OdDmCol>=": ADC("wCbOdColTo").CmpStr = "OdDmCol<="
  
  If LTrim(ADC("WRFIDSTR")) <> "" Then
           Dim wInvalRfIdStr As String
               'vk.4 calling function to replace enter with ,
               wRfIdStr = PrepareRfIdString(ADC("WRFIDSTR"))

               wInvalRfIdStr = ""
              wRfIdNoArr = Split(wRfIdStr, ",")
        
            If Left(wRfIdStr, 1) = "," Or Right(wRfIdStr, 1) = "," Then
            'vk.4 Tag Input changed as Bag Alias No.
                DispMsg "Invalid Bag Alias No.", etError: Exit Sub
            End If
            wRfIdStr = ""
            For i = 0 To UBound(wRfIdNoArr)
              wRfIdNo = wRfIdNoArr(i)
              'sv.17 Tag not numeric message changed
              'vk.4 RFID Number changed as Bag Alias No.
              If Not IsNumeric(wRfIdNo) Then DispMsg "Enter Valid Bag Alias No., Given Value is Not Numeric", etError: Exit Sub
              If Not moCn.RecSeek("Select 'x' from RfIdMst Where RimNo=" + wRfIdNo + " ") Then
                wInvalRfIdStr = wInvalRfIdStr + IIF(wInvalRfIdStr = "", "", ",") + wRfIdNo
              End If
              wRfIdStr = wRfIdStr + IIF(wRfIdStr = "", "", ",") + wRfIdNo
            Next
            wTotRfid = UBound(wRfIdNoArr) + 1
            'vk.4 Tag changed as Bag Alias No.
            If wInvalRfIdStr <> "" Then DispMsg "Bag Alias No. Not Defined:" + wInvalRfIdStr, etError: Exit Sub
                'vk.4 adc("WRFIDSTR") changed as wRfIdStr
                wCnd = IIF(LTrim(wRfIdStr) <> "", "  Exists(Select 'x' From RfIdMst where RimNo in (" + wRfIdStr + ") and BCoCd=RimBCoCd and BYy=RimBYy and BChr=RimBChr and BNo=RimBNo) and BLoc<>'PFG' and BCls='N' ", "")
  Else
    wCnd = ADC.RepCond
  End If
  
  wCnd = IIF(wCnd <> "", " and " + wCnd, "")
  ' ***** Manali 3.6.0 - 01/09/09 - Join with OrdRm added - OrRmCd checked against TrRmCd
  Set wRsCB = moCn.OpenRes("Select BCoCd, BYy, BChr, BNo, max(BQty) as BQty  from Bag join OrdDsg on BPrtKey=OdPrtKey And BOdIdNo=OdIdNo And BCocd=OdCoCd and BOdTc=OdTc and  " + _
                           "BOdYy=OdYy and BOdChr=OdChr and BOdNo=OdNo and BOdSr=OdSr  " + _
                           "Where " + IIF(gs_Partition = ctCurrPrtn, " BPrtKey='" + ctCurrPrtn + "' And ", "") + " BCoCd= '" + ADC("TrCoCd") + "' and BCls= 'N' and " + _
                            "(BQty>0 or BPtQty>0) and BLoc<>'PFG' " + _
                            "and (Select Count('x') from OrdRm Where OrCoCd=OdCoCd and OrTc=OdTc and OrYy=OdYy and OrChr=OdChr and OrNo=OdNo and OrSr=OdSr " + IIF(gs_Partition = ctCurrPrtn, " and OrPrtKey=OdPrtKey ", "") + _
                            " and OrRmCd='" + ADC("TrRmCd") + "') > 0 " + _
                            wCnd + " Group By BCoCd, BYy, BChr, BNo " + _
                            "Order By BYy, BChr, BNo ")
  
  If wRsCB.RecCount = 0 Then
    DispMsg "No Records found for this scope ", etError
    Exit Sub
  ElseIf wRsCB.RecCount > 999 Then
    DispMsg "More Than 999 records selected ,Reduce the Scope", etError
    Exit Sub
  End If
  
  GrdTreeDet.StartCopy
  MWLib.BeginProcess Me, "Copying Bag Details ..."
  With GrdTreeDet
    Do While Not (wRsCB.EOF Or wRsCB.BOF)
      .AddItem
      wRow = .Rows - 1
      .Value(wRow, "TrdSr") = .MaxVal("TrdSr") + 1
      .Value(wRow, "TrdBYy") = wRsCB!BYy
      .Value(wRow, "TrdBChr") = wRsCB!BChr
      .Value(wRow, "TrdBNo") = wRsCB!BNo
      .Value(wRow, "TrdBagPcs") = wRsCB!BQty
      .Value(wRow, "TrdDesc") = ""      ' ***** Manali 3.8.0 - Desc at Bag Level
      .SaveRec (wRow)
      wRsCB.MoveNext
    Loop
    MWLib.EndProcess Me
  GrdTreeDet.EndCopy
  '*** End of the copy process
  
    If ADC("WRFIDSTR") <> "" Then
        Dim wRsRfId As MwfLib.MDORowSet
        'vk.4 adc("WRFIDSTR") changed as wRfIdStr
        Set wRsRfId = moCn.OpenRes("Select RimNo from RfIdMst where Not Exists(Select 'x' From TreeDet where " + _
                                        " TrdCoCd=RimBCoCd and TrdBYy=RimBYY and TrdBChr=RimBChr and TrdBNo=RimBNo and " + _
                                        " TrdCoCd='" + ADC("TrCoCd") + "' and TrdYmd='" + ADC("TrYmd") + "' and TrdChr='" + ADC("TrChr") + "' and TrdNo=" + CStr(ADC("TrNo")) + _
                                        " and  TrdSr>=" + CStr(wMinSr) + " and TrdSr<=" + CStr(.MaxVal("TrdSr")) + ")" + _
                                        " and RimNo in (" + wRfIdStr + ")")
      
      
          wRfIdStr = ""
          Do While Not (wRsRfId.EOF Or wRsRfId.BOF)
                wRfIdStr = wRfIdStr + IIF(wRfIdStr = "", "", ",") + CStr(wRsRfId!RimNo)
            wRsRfId.MoveNext
          Loop
          'vk.4 RfId changed as Bag Alias No.
            If wRfIdStr <> "" Then DispMsg "Total Bag Alias No. {" + CStr(wTotRfid) + " } :Bag Alias No. Not Copied :  " + wRfIdStr, etError: Exit Sub
    End If
  
  End With
  Set wRsCB = Nothing
End Sub
'*** Jay 3.2.0 [AutoTree]
Public Sub InsertRecords(ByVal pb_Iss As Boolean)
  Dim wsql As String, wi_TdSr As Integer
  Dim wl_TdKey As Long, wi_TdKeyLen As Integer, ws_TdYyKey As String
  Dim wRs_TrDef As MDORowSet
  Dim ws_TTc As String, ws_TYy As String, ws_TChr As String, wd_TNo As Double
  Dim ws_RmCtg As String, ws_RmSCtg As String, ws_Desc As String
  Dim wd_TIdNo As Double
  
  Dim ws_CurCd As String    '*** (Jen 3.2.0)
  
  Dim wd_StkRt As Double, wTdRmStkRtYn As String, wb_AvgSRtYn As Boolean    ' ***** Manali 3.8.0 - Avg StkRt in Update
  Dim wRmLocTyp As String, wb_RmLocTyp As Boolean
  
  If moCn.RecSeek("Select 'x' from vParam where vPCoCd= '" + ms_TrCoCd + "' and vPTyp= 'TrDef' and vPValue= ''") Then
      DispMsg "Tree Default Locations Not Specified In Some Record(s) of vParam 'TRDEF'. Please Enter The Defaults Before Proceeding Further.", etError: Exit Sub
  End If
  
  On Error GoTo ErrHndlr
  
  '*** (Jen 3.2.0)
  ws_CurCd = moCn.GetFldVal("Select HBaseCurCd From Head Where HCd= '" + ctSelfCmCd + "' And HCoCd= '" + ms_TrCoCd + "'")
  '*** (Jen 3.2.0)
  
  Set wRs_TrDef = moCn.OpenRes("Select " + _
    "Max(case when vPmcd = 'CasStk' then vPValue else '' end) as qCasStk, " + _
    "(select LocTyp from Loc where LocCd= Max(case when vPmcd = 'CasStk' then vPValue else '' end) and LocCOCd= Max(vPCoCd)) as qCasStkTyp, " + _
    "Max(case vPmcd when 'CstLs' then vPValue else '' end) as qCstLs, " + _
    "(select LocTyp from Loc where LocCd= Max(case when vPmcd = 'CstLs' then vPValue else '' end) and LocCOCd= Max(vPCoCd)) as qCstLsTyp, " + _
    "Max(case vPmcd when 'CutLs' then vPValue else '' end) as qCutLs, " + _
    "(select LocTyp from Loc where LocCd= Max(case when vPmcd = 'CutLs' then vPValue else '' end) and LocCOCd= Max(vPCoCd)) as qCutLsTyp, " + _
    "Max(case vPmcd when 'FrsRMStk' then vPValue else '' end) as qFrsRMStk, " + _
    "(select LocTyp from Loc where LocCd= Max(case when vPmcd = 'FrsRMStk' then vPValue else '' end) and LocCOCd= Max(vPCoCd)) as qFrsRMStkTyp, " + _
    "Max(case vPmcd when 'GrindLs' then vPValue else '' end) as qGrindLs, " + _
    "(select LocTyp from Loc where LocCd= Max(case when vPmcd = 'GrindLs' then vPValue else '' end) and LocCOCd= Max(vPCoCd)) as qGrindLsTyp, " + _
    "Max(case vPmcd when 'IssVch' then vPValue else '' end) as qIssVch, " + _
    "Max(case vPmcd when 'PcsStk' then vPValue else '' end) as qPcsStk, " + _
    "(select LocTyp from Loc where LocCd= Max(case when vPmcd = 'PcsStk' then vPValue else '' end) and LocCOCd= Max(vPCoCd)) as qPcsStkTyp, " + _
    "Max(case vPmcd when 'RecVch' then vPValue else '' end) as qRecVch, " + _
    "Max(case vPmcd when 'RusRMStk' then vPValue else '' end) as qRusRMStk, " + _
    "(select LocTyp from Loc where LocCd= Max(case when vPmcd = 'RusRMStk' then vPValue else '' end) and LocCOCd= Max(vPCoCd)) as qRusRMStkTyp, " + _
    "Max(case vPmcd when 'ScrapLs' then vPValue else '' end) as qScrapLs, " + _
    "(select LocTyp from Loc where LocCd= Max(case when vPmcd = 'ScrapLs' then vPValue else '' end) and LocCOCd= Max(vPCoCd)) as qScrapLsTyp, " + _
    "Max(case vPmcd when 'SprueStk' then vPValue else '' end) as qSprueStk, " + _
    "(select LocTyp from Loc where LocCd= Max(case when vPmcd = 'SprueStk' then vPValue else '' end) and LocCOCd= Max(vPCoCd)) as qSprueStkTyp, " + _
    "Max(case vPmcd when 'StripLs' then vPValue else '' end) as qStripLs, " + _
    "(select LocTyp from Loc where LocCd= Max(case when vPmcd = 'StripLs' then vPValue else '' end) and LocCOCd= Max(vPCoCd)) as qStripLsTyp " + _
    "From vParam where vPCocd= '" + ms_TrCoCd + "' and vPtyp= 'TrDef'")
  
  ws_TTc = "MV"
  ws_TYy = moCn.GetFldVal("Select HYy from Head where HCoCd= '" + ms_TrCoCd + "' ")
  ws_RmCtg = moCn.GetFldVal("Select RmCtg from RmMst where RmCd= '" + ADC("TrRmcd") + "' " + IIF(gs_Partition = ctCurrPrtn, " And RmPrtKey='" + ctCurrPrtn + "' ", ""))
  ws_RmSCtg = moCn.GetFldVal("Select RmSCtg from RmMst where RmCd= '" + ADC("TrRmcd") + "' " + IIF(gs_Partition = ctCurrPrtn, " And RmPrtKey='" + ctCurrPrtn + "' ", ""))
  wi_TdSr = 0
  
  
  If ADC("TrFrMetWt") = 0# And ADC("TrUsdMetWt") = 0# Then
      DispMsg "Fresh and Reused both Metal Weight are ZERO. Cannot proceed further.", etError: Exit Sub
  End If
    
  'IG.192-Getting back/future LockDt. If lockdate is true then can't generate vouchers
  Call GetLockDt(ms_TrCoCd, "TC", ws_TTc, ms_LockDt, ms_FutureDt)
  
  mf_DtLocked = IIF((ADC("TRCSTDT") < CDate(ms_LockDt)) _
                               Or _
                   (ADC("TRCSTDT") > CDate(ms_FutureDt)), _
                True, False)
                                  
  If mf_DtLocked = True Then
    DispMsg "Cannot Generate Voucher for Locked Entry", etError: Exit Sub
  End If
  
  GetTotals
  
  '############# Check weather Difference exists or Not
  If Not pb_Iss And ADC("WTRDIFF") <> 0# Then _
          DispMsg "Weight Difference Exists, Cannot proceed further.", etError: Exit Sub
  
  If pb_Iss Then
    ws_TChr = wRs_TrDef!qIssVch
  Else
    ws_TChr = wRs_TrDef!qRecVch
  End If
  
  '***** To Generate new TNo *****
  wd_TNo = moCn.AutoGen("Txn", "TNO", ms_TrCoCd + ws_TTc + ws_TYy + ws_TChr)
  
  '***** Validate Generated new TNo *****
  If moCn.RecSeek("Select 'x' from Txn Where " + IIF(gs_Partition = ctCurrPrtn, " TPrtKey='" + ctCurrPrtn + "' And ", "") + " TCocd= '" + ms_TrCoCd + "' and TTc= '" + ws_TTc + "' and TYy= '" + ws_TYy + "' " + _
          "and TChr= '" + ws_TChr + "' and TNo= " + CStr(wd_TNo)) Then _
    DispMsg "Voucher Number Already Exists. Run Reset Autogen First.", etError: Exit Sub
  
  ' ***** Manali 3.8.0 - Avg StkRt in Update
  wTdRmStkRtYn = moCn.GetFldVal("Select PValue1 from Param join RmMst on RmCd='" + ADC("TrRmCd") + "' and RmPrtKey='" + ctCurrPrtn + "' where PTyp='RmSCtg' and PmCd=RmCtg and " + _
             " PScd = RmSCtg ")
               
  wb_AvgSRtYn = moCn.RecSeek("Select 'x' from Head where HCoCd= '" + ms_TrCoCd + "' and " + _
                      "HCd='" + ctSelfCmCd + "' and hAvgSRtYn='Y'")
  
  If wTdRmStkRtYn = "N" Or Not wb_AvgSRtYn Then wd_StkRt = 0
  ' ***** Manali 3.8.0 - Avg StkRt in Update
  Dim wChr_Ack As String, wAckYn As String, gs_ValidOthLocSeq As String
  'Ack- Getting AckYN from vParam for given Tc & Chr
  wChr_Ack = IIF(moCn.GetFldVal("Select vPNum from vParam where " + _
                                     "vPCoCd= '" + ms_TrCoCd + "' and vPTyp='CHR' and " + _
                                     "vPMCd='" + ws_TTc + "' and vPSCd='" + ws_TChr + "'") = 1, "Y", "N")
  
  'Ack- If acknowledgement is Yes then gives message & stop the change else continue the process
  If wChr_Ack = "Y" Then _
    DispMsg "Cannot Create Voucher, Requires Acknowledgement", etError: Exit Sub
    
  'Ack- Getting valid Non Production Locations for user
  gs_ValidOthLocSeq = Trim(moCn.GetFldVal("Select PValue1 from param where PTyp= 'USR' and " + _
                                     "PMCd= '" + EmrFrmMDI.TXT_UCD + "'"))
  'Ack-if Ack field in vparam Chr type is 'N' then Ack for this voucher is Yes
  'if Ack field in vparam Chr type is 'Y' then system will look for user access location for From & To location
  'if user can access both From & To location then Acknoledgement for this voucher is 'Y' else 'N'
  'when Ack=Y then stock will update else stock will not update
  If ws_TTc = "MV" And wChr_Ack = "N" Then
    wAckYn = "Y"
  Else
    If ws_TTc = "MV" And wChr_Ack = "Y" Then
    Dim wb_FrFlag As Boolean, wb_ToFlag As Boolean
      If gs_ValidOthLocSeq <> "" And wChr_Ack = "Y" Then
        wb_FrFlag = IIF((InStr(1, "," + UCase(gs_ValidOthLocSeq) + ",", "," + UCase(wRs_TrDef!qFrsRMStk) + ",") <= 0), True, False)
        wb_ToFlag = IIF((InStr(1, "," + UCase(gs_ValidOthLocSeq) + ",", "," + UCase(wRs_TrDef!qCasStk) + ",") <= 0), True, False)
        If wb_FrFlag = False And wb_ToFlag = False Then
          wAckYn = "Y"
        Else
          wAckYn = "N"
        End If
      Else
        wAckYn = "Y"
      End If
    ElseIf ws_TTc <> "MV" Then wAckYn = ""
    End If
  End If

  If pb_Iss Then
    '**** Insert Records for Casting Issue
    ' ***** Manali 3.8.0 - [TMultiIr, TImpCtg] added
    'IG.127 -Description for voucher is added while creating issue voucher
    ws_Desc = "Tree No = " + ADC("TRYMD") + "/" + ADC("TRCHR") + "/" + CStr(ADC("TRNO")) + ", Col - " + ADC("TRCOL") + ", Kt - " + ADC("TRRMCD")
    '" convert(varchar(8),getdate(),3), '" + wRs_TrDef!qFrsRMStk
    
    'ADC ("TRCSTDT")
    wsql = " Set dateformat DMY " + _
         "Insert into Txn (TTc, TYy, TChr, TNo, TDt, TFrRmLoc, TToRmLoc, TFrBLoc, TLsLoc, TLsFrDt, TLsToDt, TGldInLs, " + _
         "TGldRec, TDustWt, TSuppCd, TBillNo, TBillDt, TBOENo, TBOEDt, TCnvRt, TNetAmt, TAddAmt1, TAddAmt2, TAddAmt3, " + _
         "TAddAmt4, TTotAmt, TDesc, ModUsr, ModDt, ModTime, TUseBefDt, TCoCd, TImpTyp, TTrfQty, TSeoPwd, TLockYN, " + _
         "TLotNo, TCurrCd, TMulDiv, TCnvRtRs, TMulDivRs, TPrtKey, TCmCd, TMultiIr, TImpCtg, TSuppPayTerms, TAck) " + _
         "values ( '" + ws_TTc + "', '" + ws_TYy + "', '" + ws_TChr + "', " + CStr(wd_TNo) + ", " + _
         "'" + CStr(ADC("TRCSTDT")) + "','" + wRs_TrDef!qFrsRMStk + "', '" + wRs_TrDef!qCasStk + "', '', '', '" + CStr(MWLib.EmptyDate) + "', " + _
         "'" + CStr(MWLib.EmptyDate) + "',  0, 0, 0, '', '', '" + CStr(MWLib.EmptyDate) + "', '', '" + CStr(MWLib.EmptyDate) + "', 1, 0, 0, 0, 0, 0, 0, " + _
         "'" + ws_Desc + "', 'MIC', convert(varchar(8),getdate(),3), 0, '" + CStr(MWLib.EmptyDate) + "', '" + ms_TrCoCd + "', '', 0, '', 'N', " + _
         "'', '" + ws_CurCd + "', 'M', " + CStr(CnvRt(ws_CurCd, ctIndCurncy, moCn.SrvrDate)) + ", '" + _
         MulDiv(ws_CurCd, ctIndCurncy) + "', '" + ctCurrPrtn + "','', '', '', '','" + wAckYn + "') "
                  

    moCn.Execute wsql
    wd_TIdNo = moCn.GetFldVal("Select TIdNo from Txn where " + IIF(gs_Partition = ctCurrPrtn, " TPrtKey='" + ctCurrPrtn + "' And ", "") + " TCocd= '" + ms_TrCoCd + "' and TTc= '" + ws_TTc + "' and TYy= '" + ws_TYy + "' " + _
          "and TChr= '" + ws_TChr + "' and TNo= " + CStr(wd_TNo))
    If ADC("TrFrMetWt") <> 0# Then
      wi_TdSr = wi_TdSr + 1
      wl_TdKey = moCn.AutoGen("Txnd", "TdKey", ms_TrCoCd + ws_TYy)
      
      wi_TdKeyLen = Len(Trim(Str(wl_TdKey)))
      ws_TdYyKey = ws_TYy + Space(8 - wi_TdKeyLen) + Trim(Str(wl_TdKey))  '578-34 space 7 changed to 8
      
      ' ***** Manali 3.8.0 - Avg StkRt in Update
      wb_RmLocTyp = moCn.RecSeek("Select LocTyp from Loc where LocCoCd= '" + ms_TrCoCd + "' " + _
                                "and LocCd= '" + wRs_TrDef!qFrsRMStk + "'  and LocTyp in " + ctAvgLocTypSeq)
      'Calculate Avg Rt
      wRmLocTyp = wRs_TrDef!qFrsRMStkTyp
      
      If wTdRmStkRtYn = "Y" And wb_AvgSRtYn = True And wb_RmLocTyp = True And wRmLocTyp <> "W" Then _
        wd_StkRt = GetAvgStkRt(ms_TrCoCd, wRs_TrDef!qFrsRMStk, ADC("TrRmCd"), "", 0, 0, 0)
      ' ***** Manali 3.8.0 - Avg StkRt in Update
      
      
      ' ***** Manali 3.8.0 - TdRtByQW added in Txnd
      wsql = " Insert into Txnd " + _
        "(TdTc, TdYy, TdChr, TdNo, TdSr, TdSrNo, TdBYy, TdBChr, TdBNo, TdBQty, TdBDc, TdBGrWt, TdFrBLoc, TdToBLoc, " + _
        "TdFrRmLoc, TdFrRmDc, TdRmCd, TdRmSz, TdRmStkRt, TdRmQty, TdRmWt, TdToRmLoc, TdToRmDc, TdAck, " + _
        "TdWrk, TdPrdYN, TdByLoc, TdByWrk, TdDustWt, TdKey, TdRefYy, TdRefKey, TdRjCd, TdRjQty, TdRjDc, TdPtQty, " + _
        "TdRjPtQty, TdDesc, TdPurRt, TdPurAmt, TdPrevYy, TdPrevKey, ModUsr, ModDt, ModTime, TdYyKey, TdPcPerCt, " + _
        "TdCoCd, TdFBRmQty, TdFBRmWt, TdLotNo, TdDt, TdRmCtg, TdRmSCtg, TdFrRmLocTyp, TdToRmLocTyp, TdTIdNo, " + _
        "TdBIdNo, TdPrtKey, InsUsr, InsDt, TdRtByQW,TdBstOdTc,TdBstOdYy,TdBstOdChr,TdBstOdNo,TdBstOdSr,TdBstDmCd,TdWaxWt) values ( " + _
        "'" + ws_TTc + "', '" + ws_TYy + "', '" + ws_TChr + "', " + CStr(wd_TNo) + ", " + CStr(wi_TdSr) + ", 0, " + _
        "'', '', 0, 0, '', 0, '', '', '" + wRs_TrDef!qFrsRMStk + "', 'C', '" + ADC("TrRmCd") + "', 0, " + CStr(wd_StkRt) + " /*0*/, 0, " + _
        CStr(ADC("TrFrMetWt")) + ", '" + wRs_TrDef!qCasStk + "', 'D', '" + wAckYn + "', '', '', '', '', 0, " + CStr(wl_TdKey) + ", '', 0, " + _
        "'', 0, '', 0, 0, '', 0, 0, '', 0, 'MIC', convert(varchar(8),getdate(),3) , 0, '" + ws_TdYyKey + "', 0, " + _
        "'" + ms_TrCoCd + "', '', '', ''," + "'" + CStr(ADC("TRCSTDT")) + "','" + ws_RmCtg + "','" + ws_RmSCtg + "'," + _
        "'" + wRs_TrDef!qFrsRMStkTyp + "','" + wRs_TrDef!qCasStkTyp + "'," + CStr(wd_TIdNo) + ",0,'" + ctCurrPrtn + "', " + _
        "'" + gs_UsrCd + "', getdate(), '','','','',0,0,'',0)"
      moCn.Execute wsql
    End If
    If ADC("TrUsdMetWt") <> 0# Then
      wi_TdSr = wi_TdSr + 1
      wl_TdKey = moCn.AutoGen("Txnd", "TdKey", ms_TrCoCd + ws_TYy)
      
      wi_TdKeyLen = Len(Trim(Str(wl_TdKey)))
      ws_TdYyKey = ws_TYy + Space(8 - wi_TdKeyLen) + Trim(Str(wl_TdKey))  '578-34 space 7 changed to 8
    
      ' ***** Manali 3.8.0 - Avg StkRt in Update
      wb_RmLocTyp = moCn.RecSeek("Select LocTyp from Loc where LocCoCd= '" + ms_TrCoCd + "' " + _
                                "and LocCd= '" + wRs_TrDef!qRusRMStk + "'  and LocTyp in " + ctAvgLocTypSeq)
      'Calculate Avg Rt
      wRmLocTyp = wRs_TrDef!qRusRMStkTyp
      
      If wTdRmStkRtYn = "Y" And wb_AvgSRtYn = True And wb_RmLocTyp = True And wRmLocTyp <> "W" Then _
        wd_StkRt = GetAvgStkRt(ms_TrCoCd, wRs_TrDef!qRusRMStk, ADC("TrRmCd"), "", 0, 0, 0)
      ' ***** Manali 3.8.0 - Avg StkRt in Update
      
     ' ***** Manali 3.8.0 - TdRtByQW added in Txnd
     wsql = " Insert into Txnd " + _
        "(TdTc, TdYy, TdChr, TdNo, TdSr, TdSrNo, TdBYy, TdBChr, TdBNo, TdBQty, TdBDc, TdBGrWt, TdFrBLoc, TdToBLoc, " + _
        "TdFrRmLoc, TdFrRmDc, TdRmCd, TdRmSz, TdRmStkRt, TdRmQty, TdRmWt, TdToRmLoc, TdToRmDc, TdAck, " + _
        "TdWrk, TdPrdYN, TdByLoc, TdByWrk, TdDustWt, TdKey, TdRefYy, TdRefKey, TdRjCd, TdRjQty, TdRjDc, TdPtQty, " + _
        "TdRjPtQty, TdDesc, TdPurRt, TdPurAmt, TdPrevYy, TdPrevKey, ModUsr, ModDt, ModTime, TdYyKey, TdPcPerCt, " + _
        "TdCoCd, TdFBRmQty, TdFBRmWt, TdLotNo, TdDt, TdRmCtg, TdRmSCtg, TdFrRmLocTyp, TdToRmLocTyp, TdTIdNo, " + _
        "TdBIdNo, TdPrtKey, InsUsr, InsDt, TdRtByQW,TdBstOdTc,TdBstOdYy,TdBstOdChr,TdBstOdNo,TdBstOdSr,TdBstDmCd,TdWaxWt) values ( " + _
        "'" + ws_TTc + "', '" + ws_TYy + "', '" + ws_TChr + "', " + CStr(wd_TNo) + ", " + CStr(wi_TdSr) + ", 0, " + _
        "'', '', 0, 0, '', 0, '', '', '" + wRs_TrDef!qRusRMStk + "', 'C', '" + ADC("TrRmCd") + "', 0, " + CStr(wd_StkRt) + " /*0*/, 0, " + _
        CStr(ADC("TrUsdMetWt")) + ", '" + wRs_TrDef!qCasStk + "', 'D', '" + wAckYn + "', '', '', '', '', 0, " + CStr(wl_TdKey) + ", '', 0, " + _
        "'', 0, '', 0, 0, '', 0, 0, '', 0, 'MIC', convert(varchar(8),getdate(),3) , 0, '" + ws_TdYyKey + "', 0, " + _
        "'" + ms_TrCoCd + "', '', '', '','" + CStr(ADC("TRCSTDT")) + "','" + ws_RmCtg + "','" + ws_RmSCtg + "'," + _
        "'" + wRs_TrDef!qRusRMStkTyp + "','" + wRs_TrDef!qCasStkTyp + "'," + CStr(wd_TIdNo) + ",0,'" + ctCurrPrtn + "', " + _
        "'" + gs_UsrCd + "', getdate(), '','','','',0,0,'',0 )"
      moCn.Execute wsql
    End If
    ADC("TrIssVchNo") = ms_TrCoCd + "/" + ws_TTc + "/" + ws_TYy + "/" + ws_TChr + "/" + CStr(wd_TNo)
    
  Else
    '**** Insert Records from Casting Receipt
    ' **** Manali 3.8.0 - TrDustWt in TDesc -
    ' ***** Manali 3.8.0 - [TMultiIr, TImpCtg] added
    'IG.127 -Description for voucher is added while creating receipt voucher
    ws_Desc = "Tree No=" + ADC("TRYMD") + "/" + ADC("TRCHR") + "/" + CStr(ADC("TRNO")) + ", Col-" + ADC("TRCOL") + ", Kt-" + ADC("TRRMCD")
    '578.19 TDEsc size is 60. left 60+DustWt: cannot insert into TDesc field
    'is is giving error. so left 60 changed to 53
    wsql = " Set dateformat DMY " + _
        "Insert into Txn (TTc, TYy, TChr, TNo, TDt, TFrRmLoc, TToRmLoc, TFrBLoc, TLsLoc, TLsFrDt, TLsToDt, TGldInLs, " + _
        "TGldRec, TDustWt, TSuppCd, TBillNo, TBillDt, TBOENo, TBOEDt, TCnvRt, TNetAmt, TAddAmt1, TAddAmt2, TAddAmt3, " + _
        "TAddAmt4, TTotAmt, TDesc, ModUsr, ModDt, ModTime, TUseBefDt, TCoCd, TImpTyp, TTrfQty, TSeoPwd, TLockYN, " + _
        "TLotNo, TCurrCd, TMulDiv, TCnvRtRs, TMulDivRs, TPrtKey, TCmCd, TMultiIr, TImpCtg, TSuppPayTerms, TAck) " + _
        "values ( '" + ws_TTc + "', '" + ws_TYy + "', '" + ws_TChr + "', " + CStr(wd_TNo) + ", " + _
        " '" + CStr(ADC("TRCSTDT")) + "', '" + wRs_TrDef!qFrsRMStk + "', '" + wRs_TrDef!qCasStk + "', '', '', '" + CStr(MWLib.EmptyDate) + "', " + _
        "'" + CStr(MWLib.EmptyDate) + "',  0, 0, 0, '', '', '" + CStr(MWLib.EmptyDate) + "', '', '" + CStr(MWLib.EmptyDate) + "', 1, 0, 0, 0, 0, 0, 0, " + _
        "'DustWt:'+Left('" + CStr(ADC("TrDustWt")) + "Gms, " + ws_Desc + "', 53), 'MIC', convert(varchar(8),getdate(),3), 0, '" + CStr(MWLib.EmptyDate) + "', '" + ms_TrCoCd + "', '', 0, '', 'N', " + _
        "'', '" + ws_CurCd + "', 'M', " + CStr(CnvRt(ws_CurCd, ctIndCurncy, moCn.SrvrDate)) + ", '" + _
         MulDiv(ws_CurCd, ctIndCurncy) + "', '" + ctCurrPrtn + "', '', '', '', '','" + wAckYn + "') "
            
    moCn.Execute wsql
    wd_TIdNo = moCn.GetFldVal("Select TIdNo from Txn Where " + IIF(gs_Partition = ctCurrPrtn, " TPrtKey='" + ctCurrPrtn + "' And ", "") + " TCocd= '" + ms_TrCoCd + "' and TTc= '" + ws_TTc + "' and TYy= '" + ws_TYy + "' " + _
          "and TChr= '" + ws_TChr + "' and TNo= " + CStr(wd_TNo))
          
    If ADC("TrRecdMetWt") <> 0# Then
      wi_TdSr = wi_TdSr + 1
      wl_TdKey = moCn.AutoGen("Txnd", "TdKey", ms_TrCoCd + ws_TYy)
      wi_TdKeyLen = Len(Trim(Str(wl_TdKey)))
      ws_TdYyKey = ws_TYy + Space(8 - wi_TdKeyLen) + Trim(Str(wl_TdKey))  '578-34 space 7 changed to 8

      ' ***** Manali 3.8.0 - Avg StkRt in Update
      wb_RmLocTyp = moCn.RecSeek("Select LocTyp from Loc where LocCoCd= '" + ms_TrCoCd + "' " + _
                                "and LocCd= '" + wRs_TrDef!qCasStk + "'  and LocTyp in " + ctAvgLocTypSeq)
      'Calculate Avg Rt
      wRmLocTyp = wRs_TrDef!qCasStkTyp
      
      If wTdRmStkRtYn = "Y" And wb_AvgSRtYn = True And wb_RmLocTyp = True And wRmLocTyp <> "W" Then _
        wd_StkRt = GetAvgStkRt(ms_TrCoCd, wRs_TrDef!qCasStk, ADC("TrRmCd"), "", 0, 0, 0)
      ' ***** Manali 3.8.0 - Avg StkRt in Update

      ' ***** Manali 3.8.0 - TdRtByQW added in Txnd
      wsql = " Insert into Txnd " + _
        "(TdTc, TdYy, TdChr, TdNo, TdSr, TdSrNo, TdBYy, TdBChr, TdBNo, TdBQty, TdBDc, TdBGrWt, TdFrBLoc, TdToBLoc, " + _
        "TdFrRmLoc, TdFrRmDc, TdRmCd, TdRmSz, TdRmStkRt, TdRmQty, TdRmWt, TdToRmLoc, TdToRmDc, TdAck, " + _
        "TdWrk, TdPrdYN, TdByLoc, TdByWrk, TdDustWt, TdKey, TdRefYy, TdRefKey, TdRjCd, TdRjQty, TdRjDc, TdPtQty, " + _
        "TdRjPtQty, TdDesc, TdPurRt, TdPurAmt, TdPrevYy, TdPrevKey, ModUsr, ModDt, ModTime, TdYyKey, TdPcPerCt, " + _
        "TdCoCd, TdFBRmQty, TdFBRmWt, TdLotNo, TdDt, TdRmCtg, TdRmSCtg, TdFrRmLocTyp, TdToRmLocTyp, TdTIdNo, " + _
        "TdBIdNo, TdPrtKey, InsUsr, InsDt, TdRtByQW,TdBstOdTc,TdBstOdYy,TdBstOdChr,TdBstOdNo,TdBstOdSr,TdBstDmCd,TdWaxWt) values ( " + _
        "'" + ws_TTc + "', '" + ws_TYy + "', '" + ws_TChr + "', " + CStr(wd_TNo) + ", " + CStr(wi_TdSr) + ", 0, " + _
        "'', '', 0, 0, '', 0, '', '', '" + wRs_TrDef!qCasStk + "', 'C', '" + ADC("TrRmCd") + "', 0, " + CStr(wd_StkRt) + " /*0*/, 0, " + _
        CStr(ADC("TrRecdMetWt")) + ", '" + wRs_TrDef!qSprueStk + "', 'D', '" + wAckYn + "', '', '', '', '', 0, " + CStr(wl_TdKey) + ", '', 0, " + _
        "'', 0, '', 0, 0, '', 0, 0, '', 0, 'MIC', convert(varchar(8),getdate(),3) , 0, '" + ws_TdYyKey + "', 0, " + _
        "'" + ms_TrCoCd + "', '', '', '','" + CStr(ADC("TRCSTDT")) + "','" + ws_RmCtg + "','" + ws_RmSCtg + "'," + _
        "'" + wRs_TrDef!qCasStkTyp + "','" + wRs_TrDef!qSprueStkTyp + "'," + CStr(wd_TIdNo) + ",0,'" + ctCurrPrtn + "', " + _
        "'" + gs_UsrCd + "', getdate(), '','','','',0,0,'',0  )"
      moCn.Execute wsql
    End If
    
    ' ****** Manali 350Nxt - TrNetPcWt used for Calculations instead of TrGrsPcWt
    'If ADC("TrGrsPcWt") <> 0# Then
    If ADC("TrNetPcWt") <> 0# Then
      wi_TdSr = wi_TdSr + 1
      wl_TdKey = moCn.AutoGen("Txnd", "TdKey", ms_TrCoCd + ws_TYy)
      wi_TdKeyLen = Len(Trim(Str(wl_TdKey)))
      ws_TdYyKey = ws_TYy + Space(8 - wi_TdKeyLen) + Trim(Str(wl_TdKey))  '578-34 space 7 changed to 8

      ' ***** Manali 3.8.0 - Avg StkRt in Update
      wb_RmLocTyp = moCn.RecSeek("Select LocTyp from Loc where LocCoCd= '" + ms_TrCoCd + "' " + _
                                "and LocCd= '" + wRs_TrDef!qCasStk + "'  and LocTyp in " + ctAvgLocTypSeq)
      'Calculate Avg Rt
      wRmLocTyp = wRs_TrDef!qCasStkTyp
      
      If wTdRmStkRtYn = "Y" And wb_AvgSRtYn = True And wb_RmLocTyp = True And wRmLocTyp <> "W" Then _
        wd_StkRt = GetAvgStkRt(ms_TrCoCd, wRs_TrDef!qCasStk, ADC("TrRmCd"), "", 0)
      ' ***** Manali 3.8.0 - Avg StkRt in Update

      ' ***** Manali 3.8.0 - TdRtByQW added in Txnd
      wsql = " Insert into Txnd " + _
        "(TdTc, TdYy, TdChr, TdNo, TdSr, TdSrNo, TdBYy, TdBChr, TdBNo, TdBQty, TdBDc, TdBGrWt, TdFrBLoc, TdToBLoc, " + _
        "TdFrRmLoc, TdFrRmDc, TdRmCd, TdRmSz, TdRmStkRt, TdRmQty, TdRmWt, TdToRmLoc, TdToRmDc, TdAck, " + _
        "TdWrk, TdPrdYN, TdByLoc, TdByWrk, TdDustWt, TdKey, TdRefYy, TdRefKey, TdRjCd, TdRjQty, TdRjDc, TdPtQty, " + _
        "TdRjPtQty, TdDesc, TdPurRt, TdPurAmt, TdPrevYy, TdPrevKey, ModUsr, ModDt, ModTime, TdYyKey, TdPcPerCt, " + _
        "TdCoCd, TdFBRmQty, TdFBRmWt, TdLotNo, TdDt, TdRmCtg, TdRmSCtg, TdFrRmLocTyp, TdToRmLocTyp, TdTIdNo, " + _
        "TdBIdNo, TdPrtKey, InsUsr, InsDt, TdRtByQW,TdBstOdTc,TdBstOdYy,TdBstOdChr,TdBstOdNo,TdBstOdSr,TdBstDmCd,TdWaxWt) values ( " + _
        "'" + ws_TTc + "', '" + ws_TYy + "', '" + ws_TChr + "', " + CStr(wd_TNo) + ", " + CStr(wi_TdSr) + ", 0, " + _
        "'', '', 0, 0, '', 0, '', '', '" + wRs_TrDef!qCasStk + "', 'C', '" + ADC("TrRmCd") + "', 0, " + CStr(wd_StkRt) + " /*0*/, 0, /* " + _
        CStr(ADC("TrGrsPcWt")) + " */ " + CStr(ADC("TrNetPcWt")) + ", '" + wRs_TrDef!qPcsStk + "', 'D', '" + wAckYn + "', '', '', '', '', 0, " + CStr(wl_TdKey) + ", '', 0, " + _
        "'', 0, '', 0, 0, '', 0, 0, '', 0, 'MIC', convert(varchar(8),getdate(),3) , 0, '" + ws_TdYyKey + "', 0, " + _
        "'" + ms_TrCoCd + "', '', '', ''," + "'" + CStr(ADC("TRCSTDT")) + "','" + ws_RmCtg + "','" + ws_RmSCtg + "'," + _
        "'" + wRs_TrDef!qCasStkTyp + "','" + wRs_TrDef!qPcsStkTyp + "'," + CStr(wd_TIdNo) + ",0,'" + ctCurrPrtn + "', " + _
        "'" + gs_UsrCd + "', getdate(), '','','','',0,0,'',0  )"
      ' ****** Manali 350Nxt - TrNetPcWt used for Calculations instead of TrGrsPcWt
      moCn.Execute wsql
    End If
    If ADC("TrCastLoss") <> 0# Then
      wi_TdSr = wi_TdSr + 1
      wl_TdKey = moCn.AutoGen("Txnd", "TdKey", ms_TrCoCd + ws_TYy)
      wi_TdKeyLen = Len(Trim(Str(wl_TdKey)))
      ws_TdYyKey = ws_TYy + Space(8 - wi_TdKeyLen) + Trim(Str(wl_TdKey))  '578-34 space 7 changed to 8

      ' ***** Manali 3.8.0 - Avg StkRt in Update
      wb_RmLocTyp = moCn.RecSeek("Select LocTyp from Loc where LocCoCd= '" + ms_TrCoCd + "' " + _
                                "and LocCd= '" + wRs_TrDef!qCasStk + "'  and LocTyp in " + ctAvgLocTypSeq)
      'Calculate Avg Rt
      wRmLocTyp = wRs_TrDef!qCasStkTyp
      
      If wTdRmStkRtYn = "Y" And wb_AvgSRtYn = True And wb_RmLocTyp = True And wRmLocTyp <> "W" Then _
        wd_StkRt = GetAvgStkRt(ms_TrCoCd, wRs_TrDef!qCasStk, ADC("TrRmCd"), "", 0)
      ' ***** Manali 3.8.0 - Avg StkRt in Update

      ' ***** Manali 3.8.0 - TdRtByQW added in Txnd
      wsql = " Insert into Txnd " + _
        "(TdTc, TdYy, TdChr, TdNo, TdSr, TdSrNo, TdBYy, TdBChr, TdBNo, TdBQty, TdBDc, TdBGrWt, TdFrBLoc, TdToBLoc, " + _
        "TdFrRmLoc, TdFrRmDc, TdRmCd, TdRmSz, TdRmStkRt, TdRmQty, TdRmWt, TdToRmLoc, TdToRmDc, TdAck, " + _
        "TdWrk, TdPrdYN, TdByLoc, TdByWrk, TdDustWt, TdKey, TdRefYy, TdRefKey, TdRjCd, TdRjQty, TdRjDc, TdPtQty, " + _
        "TdRjPtQty, TdDesc, TdPurRt, TdPurAmt, TdPrevYy, TdPrevKey, ModUsr, ModDt, ModTime, TdYyKey, TdPcPerCt, " + _
        "TdCoCd, TdFBRmQty, TdFBRmWt, TdLotNo, TdDt, TdRmCtg, TdRmSCtg, TdFrRmLocTyp, TdToRmLocTyp, TdTIdNo, " + _
        "TdBIdNo, TdPrtKey, InsUsr, InsDt, TdRtByQW,TdBstOdTc,TdBstOdYy,TdBstOdChr,TdBstOdNo,TdBstOdSr,TdBstDmCd,TdWaxWt) values ( " + _
        "'" + ws_TTc + "', '" + ws_TYy + "', '" + ws_TChr + "', " + CStr(wd_TNo) + ", " + CStr(wi_TdSr) + ", 0, " + _
        "'', '', 0, 0, '', 0, '', '', '" + wRs_TrDef!qCasStk + "', 'C', '" + ADC("TrRmCd") + "', 0, " + CStr(wd_StkRt) + " /*0*/, 0, " + _
        CStr(ADC("TrCastLoss")) + ", '" + wRs_TrDef!qCstLs + "', 'D', '" + wAckYn + "', '', '', '', '', 0, " + CStr(wl_TdKey) + ", '', 0, " + _
        "'', 0, '', 0, 0, '', 0, 0, '', 0, 'MIC', convert(varchar(8),getdate(),3) , 0, '" + ws_TdYyKey + "', 0, " + _
        "'" + ms_TrCoCd + "', '', '', '','" + CStr(ADC("TRCSTDT")) + "','" + ws_RmCtg + "','" + ws_RmSCtg + "'," + _
        "'" + wRs_TrDef!qCasStkTyp + "','" + wRs_TrDef!qCstLsTyp + "'," + CStr(wd_TIdNo) + ",0,'" + ctCurrPrtn + "', " + _
        "'" + gs_UsrCd + "', getdate(), '','','','',0,0,'',0)"
      moCn.Execute wsql
    End If
    If ADC("TrCutLoss") <> 0# Then
      wi_TdSr = wi_TdSr + 1
      wl_TdKey = moCn.AutoGen("Txnd", "TdKey", ms_TrCoCd + ws_TYy)
      wi_TdKeyLen = Len(Trim(Str(wl_TdKey)))
      ws_TdYyKey = ws_TYy + Space(8 - wi_TdKeyLen) + Trim(Str(wl_TdKey))  '578-34 space 7 changed to 8

      ' ***** Manali 3.8.0 - Avg StkRt in Update
      wb_RmLocTyp = moCn.RecSeek("Select LocTyp from Loc where LocCoCd= '" + ms_TrCoCd + "' " + _
                                "and LocCd= '" + wRs_TrDef!qCasStk + "'  and LocTyp in " + ctAvgLocTypSeq)
      'Calculate Avg Rt
      wRmLocTyp = wRs_TrDef!qCasStkTyp
      
      If wTdRmStkRtYn = "Y" And wb_AvgSRtYn = True And wb_RmLocTyp = True And wRmLocTyp <> "W" Then _
        wd_StkRt = GetAvgStkRt(ms_TrCoCd, wRs_TrDef!qCasStk, ADC("TrRmCd"), "", 0)
      ' ***** Manali 3.8.0 - Avg StkRt in Update

      ' ***** Manali 3.8.0 - TdRtByQW added in Txnd
      wsql = " Insert into Txnd " + _
        "(TdTc, TdYy, TdChr, TdNo, TdSr, TdSrNo, TdBYy, TdBChr, TdBNo, TdBQty, TdBDc, TdBGrWt, TdFrBLoc, TdToBLoc, " + _
        "TdFrRmLoc, TdFrRmDc, TdRmCd, TdRmSz, TdRmStkRt, TdRmQty, TdRmWt, TdToRmLoc, TdToRmDc, TdAck, " + _
        "TdWrk, TdPrdYN, TdByLoc, TdByWrk, TdDustWt, TdKey, TdRefYy, TdRefKey, TdRjCd, TdRjQty, TdRjDc, TdPtQty, " + _
        "TdRjPtQty, TdDesc, TdPurRt, TdPurAmt, TdPrevYy, TdPrevKey, ModUsr, ModDt, ModTime, TdYyKey, TdPcPerCt, " + _
        "TdCoCd, TdFBRmQty, TdFBRmWt, TdLotNo, TdDt, TdRmCtg, TdRmSCtg, TdFrRmLocTyp, TdToRmLocTyp, TdTIdNo, " + _
        "TdBIdNo, TdPrtKey, InsUsr, InsDt, TdRtByQW,TdBstOdTc,TdBstOdYy,TdBstOdChr,TdBstOdNo,TdBstOdSr,TdBstDmCd,TdWaxWt) values ( " + _
        "'" + ws_TTc + "', '" + ws_TYy + "', '" + ws_TChr + "', " + CStr(wd_TNo) + ", " + CStr(wi_TdSr) + ", 0, " + _
        "'', '', 0, 0, '', 0, '', '', '" + wRs_TrDef!qCasStk + "', 'C', '" + ADC("TrRmCd") + "', 0, " + CStr(wd_StkRt) + " /*0*/, 0, " + _
        CStr(ADC("TrCutLoss")) + ", '" + wRs_TrDef!qCutLs + "', 'D', '" + wAckYn + "', '', '', '', '', 0, " + CStr(wl_TdKey) + ", '', 0, " + _
        "'', 0, '', 0, 0, '', 0, 0, '', 0, 'MIC', convert(varchar(8),getdate(),3) , 0, '" + ws_TdYyKey + "', 0, " + _
        "'" + ms_TrCoCd + "', '', '', '','" + CStr(ADC("TRCSTDT")) + "','" + ws_RmCtg + "','" + ws_RmSCtg + "'," + _
        "'" + wRs_TrDef!qCasStkTyp + "','" + wRs_TrDef!qCutLsTyp + "'," + CStr(wd_TIdNo) + ",0,'" + ctCurrPrtn + "', " + _
        "'" + gs_UsrCd + "', getdate(), '','','','',0,0,'',0)"
      moCn.Execute wsql
    End If
    If ADC("TrScrap") <> 0# Then
      wi_TdSr = wi_TdSr + 1
      wl_TdKey = moCn.AutoGen("Txnd", "TdKey", ms_TrCoCd + ws_TYy)
      wi_TdKeyLen = Len(Trim(Str(wl_TdKey)))
      ws_TdYyKey = ws_TYy + Space(8 - wi_TdKeyLen) + Trim(Str(wl_TdKey))  '578-34 space 7 changed to 8

      ' ***** Manali 3.8.0 - Avg StkRt in Update
      wb_RmLocTyp = moCn.RecSeek("Select LocTyp from Loc where LocCoCd= '" + ms_TrCoCd + "' " + _
                                "and LocCd= '" + wRs_TrDef!qCasStk + "'  and LocTyp in " + ctAvgLocTypSeq)
      'Calculate Avg Rt
      wRmLocTyp = wRs_TrDef!qCasStkTyp
      
      If wTdRmStkRtYn = "Y" And wb_AvgSRtYn = True And wb_RmLocTyp = True And wRmLocTyp <> "W" Then _
        wd_StkRt = GetAvgStkRt(ms_TrCoCd, wRs_TrDef!qCasStk, ADC("TrRmCd"), "", 0)
      ' ***** Manali 3.8.0 - Avg StkRt in Update

      ' ***** Manali 3.8.0 - TdRtByQW added in Txnd
      wsql = " Insert into Txnd " + _
        "(TdTc, TdYy, TdChr, TdNo, TdSr, TdSrNo, TdBYy, TdBChr, TdBNo, TdBQty, TdBDc, TdBGrWt, TdFrBLoc, TdToBLoc, " + _
        "TdFrRmLoc, TdFrRmDc, TdRmCd, TdRmSz, TdRmStkRt, TdRmQty, TdRmWt, TdToRmLoc, TdToRmDc, TdAck, " + _
        "TdWrk, TdPrdYN, TdByLoc, TdByWrk, TdDustWt, TdKey, TdRefYy, TdRefKey, TdRjCd, TdRjQty, TdRjDc, TdPtQty, " + _
        "TdRjPtQty, TdDesc, TdPurRt, TdPurAmt, TdPrevYy, TdPrevKey, ModUsr, ModDt, ModTime, TdYyKey, TdPcPerCt, " + _
        "TdCoCd, TdFBRmQty, TdFBRmWt, TdLotNo, TdDt, TdRmCtg, TdRmSCtg, TdFrRmLocTyp, TdToRmLocTyp, TdTIdNo, " + _
        "TdBIdNo, TdPrtKey, InsUsr, InsDt, TdRtByQW,TdBstOdTc,TdBstOdYy,TdBstOdChr,TdBstOdNo,TdBstOdSr,TdBstDmCd,TdWaxWt) values ( " + _
        "'" + ws_TTc + "', '" + ws_TYy + "', '" + ws_TChr + "', " + CStr(wd_TNo) + ", " + CStr(wi_TdSr) + ", 0, " + _
        "'', '', 0, 0, '', 0, '', '', '" + wRs_TrDef!qCasStk + "', 'C', '" + ADC("TrRmCd") + "', 0, " + CStr(wd_StkRt) + " /*0*/, 0, " + _
        CStr(ADC("TrScrap")) + ", '" + wRs_TrDef!qScrapLs + "', 'D', '" + wAckYn + "', '', '', '', '', 0, " + CStr(wl_TdKey) + ", '', 0, " + _
        "'', 0, '', 0, 0, '', 0, 0, '', 0, 'MIC', convert(varchar(8),getdate(),3) , 0, '" + ws_TdYyKey + "', 0, " + _
        "'" + ms_TrCoCd + "', '', '', '','" + CStr(ADC("TRCSTDT")) + "','" + ws_RmCtg + "','" + ws_RmSCtg + "'," + _
        "'" + wRs_TrDef!qCasStkTyp + "','" + wRs_TrDef!qScrapLsTyp + "'," + CStr(wd_TIdNo) + ",0,'" + ctCurrPrtn + "', " + _
        "'" + gs_UsrCd + "', getdate(), '','','','',0,0,'',0)"
      moCn.Execute wsql
    End If
    If ADC("TrGrindLoss") <> 0# Then
      wi_TdSr = wi_TdSr + 1
      wl_TdKey = moCn.AutoGen("Txnd", "TdKey", ms_TrCoCd + ws_TYy)
      wi_TdKeyLen = Len(Trim(Str(wl_TdKey)))
      ws_TdYyKey = ws_TYy + Space(8 - wi_TdKeyLen) + Trim(Str(wl_TdKey))  '578-34 space 7 changed to 8

      ' ***** Manali 3.8.0 - Avg StkRt in Update
      wb_RmLocTyp = moCn.RecSeek("Select LocTyp from Loc where LocCoCd= '" + ms_TrCoCd + "' " + _
                                "and LocCd= '" + wRs_TrDef!qCasStk + "'  and LocTyp in " + ctAvgLocTypSeq)
      'Calculate Avg Rt
      wRmLocTyp = wRs_TrDef!qCasStkTyp
      
      If wTdRmStkRtYn = "Y" And wb_AvgSRtYn = True And wb_RmLocTyp = True And wRmLocTyp <> "W" Then _
        wd_StkRt = GetAvgStkRt(ms_TrCoCd, wRs_TrDef!qCasStk, ADC("TrRmCd"), "", 0)
      ' ***** Manali 3.8.0 - Avg StkRt in Update

      ' ***** Manali 3.8.0 - TdRtByQW added in Txnd
      wsql = " Insert into Txnd " + _
        "(TdTc, TdYy, TdChr, TdNo, TdSr, TdSrNo, TdBYy, TdBChr, TdBNo, TdBQty, TdBDc, TdBGrWt, TdFrBLoc, TdToBLoc, " + _
        "TdFrRmLoc, TdFrRmDc, TdRmCd, TdRmSz, TdRmStkRt, TdRmQty, TdRmWt, TdToRmLoc, TdToRmDc, TdAck, " + _
        "TdWrk, TdPrdYN, TdByLoc, TdByWrk, TdDustWt, TdKey, TdRefYy, TdRefKey, TdRjCd, TdRjQty, TdRjDc, TdPtQty, " + _
        "TdRjPtQty, TdDesc, TdPurRt, TdPurAmt, TdPrevYy, TdPrevKey, ModUsr, ModDt, ModTime, TdYyKey, TdPcPerCt, " + _
        "TdCoCd, TdFBRmQty, TdFBRmWt, TdLotNo, TdDt, TdRmCtg, TdRmSCtg, TdFrRmLocTyp, TdToRmLocTyp, TdTIdNo, " + _
        "TdBIdNo, TdPrtKey, InsUsr, InsDt, TdRtByQW,TdBstOdTc,TdBstOdYy,TdBstOdChr,TdBstOdNo,TdBstOdSr,TdBstDmCd,TdWaxWt) values ( " + _
        "'" + ws_TTc + "', '" + ws_TYy + "', '" + ws_TChr + "', " + CStr(wd_TNo) + ", " + CStr(wi_TdSr) + ", 0, " + _
        "'', '', 0, 0, '', 0, '', '', '" + wRs_TrDef!qCasStk + "', 'C', '" + ADC("TrRmCd") + "', 0, " + CStr(wd_StkRt) + " /*0*/, 0, " + _
        CStr(ADC("TrGrindLoss")) + ", '" + wRs_TrDef!qGrindLs + "', 'D', '" + wAckYn + "', '', '', '', '', 0, " + CStr(wl_TdKey) + ", '', 0, " + _
        "'', 0, '', 0, 0, '', 0, 0, '', 0, 'MIC', convert(varchar(8),getdate(),3) , 0, '" + ws_TdYyKey + "', 0, " + _
        "'" + ms_TrCoCd + "', '', '', '','" + CStr(ADC("TRCSTDT")) + "','" + ws_RmCtg + "','" + ws_RmSCtg + "'," + _
        "'" + wRs_TrDef!qCasStkTyp + "','" + wRs_TrDef!qGrindLsTyp + "'," + CStr(wd_TIdNo) + ",0,'" + ctCurrPrtn + "', " + _
        "'" + gs_UsrCd + "', getdate(), '','','','',0,0,'',0)"
      moCn.Execute wsql
    End If
    If ADC("TrStripLoss") <> 0# Then
      wi_TdSr = wi_TdSr + 1
      wl_TdKey = moCn.AutoGen("Txnd", "TdKey", ms_TrCoCd + ws_TYy)
      wi_TdKeyLen = Len(Trim(Str(wl_TdKey)))
      ws_TdYyKey = ws_TYy + Space(8 - wi_TdKeyLen) + Trim(Str(wl_TdKey))  '578-34 space 7 changed to 8

      ' ***** Manali 3.8.0 - Avg StkRt in Update
      wb_RmLocTyp = moCn.RecSeek("Select LocTyp from Loc where LocCoCd= '" + ms_TrCoCd + "' " + _
                                "and LocCd= '" + wRs_TrDef!qCasStk + "'  and LocTyp in " + ctAvgLocTypSeq)
      'Calculate Avg Rt
      wRmLocTyp = wRs_TrDef!qCasStkTyp
      
      If wTdRmStkRtYn = "Y" And wb_AvgSRtYn = True And wb_RmLocTyp = True And wRmLocTyp <> "W" Then _
        wd_StkRt = GetAvgStkRt(ms_TrCoCd, wRs_TrDef!qCasStk, ADC("TrRmCd"), "", 0)
      ' ***** Manali 3.8.0 - Avg StkRt in Update

      ' ***** Manali 3.8.0 - TdRtByQW added in Txnd
      wsql = " Insert into Txnd " + _
        "(TdTc, TdYy, TdChr, TdNo, TdSr, TdSrNo, TdBYy, TdBChr, TdBNo, TdBQty, TdBDc, TdBGrWt, TdFrBLoc, TdToBLoc, " + _
        "TdFrRmLoc, TdFrRmDc, TdRmCd, TdRmSz, TdRmStkRt, TdRmQty, TdRmWt, TdToRmLoc, TdToRmDc, TdAck, " + _
        "TdWrk, TdPrdYN, TdByLoc, TdByWrk, TdDustWt, TdKey, TdRefYy, TdRefKey, TdRjCd, TdRjQty, TdRjDc, TdPtQty, " + _
        "TdRjPtQty, TdDesc, TdPurRt, TdPurAmt, TdPrevYy, TdPrevKey, ModUsr, ModDt, ModTime, TdYyKey, TdPcPerCt, " + _
        "TdCoCd, TdFBRmQty, TdFBRmWt, TdLotNo, TdDt, TdRmCtg, TdRmSCtg, TdFrRmLocTyp, TdToRmLocTyp, TdTIdNo, " + _
        "TdBIdNo, TdPrtKey, InsUsr, InsDt, TdRtByQW,TdBstOdTc,TdBstOdYy,TdBstOdChr,TdBstOdNo,TdBstOdSr,TdBstDmCd,TdWaxWt) values ( " + _
        "'" + ws_TTc + "', '" + ws_TYy + "', '" + ws_TChr + "', " + CStr(wd_TNo) + ", " + CStr(wi_TdSr) + ", 0, " + _
        "'', '', 0, 0, '', 0, '', '', '" + wRs_TrDef!qCasStk + "', 'C', '" + ADC("TrRmCd") + "', 0, " + CStr(wd_StkRt) + " /*0*/, 0, " + _
        CStr(ADC("TrStripLoss")) + ", '" + wRs_TrDef!qStripLs + "', 'D', '" + wAckYn + "', '', '', '', '', 0, " + CStr(wl_TdKey) + ", '', 0, " + _
        "'', 0, '', 0, 0, '', 0, 0, '', 0, 'MIC', convert(varchar(8),getdate(),3) , 0, '" + ws_TdYyKey + "', 0, " + _
        "'" + ms_TrCoCd + "', '', '', '','" + CStr(ADC("TRCSTDT")) + "','" + ws_RmCtg + "','" + ws_RmSCtg + "'," + _
        "'" + wRs_TrDef!qCasStkTyp + "','" + wRs_TrDef!qStripLsTyp + "'," + CStr(wd_TIdNo) + ",0,'" + ctCurrPrtn + "', " + _
        "'" + gs_UsrCd + "', getdate(), '','','','',0,0,'',0)"
      moCn.Execute wsql
    End If
    ADC("TrRecVchNo") = ms_TrCoCd + "/" + ws_TTc + "/" + ws_TYy + "/" + ws_TChr + "/" + CStr(wd_TNo)
        
  End If

  If moCn.RecSeek("Select 'x' from Txn where " + IIF(gs_Partition = ctCurrPrtn, " TPrtKey='" + ctCurrPrtn + "' And ", "") + " TCocd= '" + ms_TrCoCd + "' and TTc= '" + ws_TTc + "' and TYy= '" + ws_TYy + "' and TChr= '" + ws_TChr + "' and TNo= " + CStr(wd_TNo) + "  ") Then
    DispMsg "Update Completed Successfully To : " + ms_TrCoCd + "/" + ws_TTc + "/" + ws_TYy + "/" + ws_TChr + "/" + CStr(wd_TNo), etInfo
  End If
  
Exit Sub
ErrHndlr:
  wsql = "Delete from Txnd where " + IIF(gs_Partition = ctCurrPrtn, " TdPrtKey='" + ctCurrPrtn + "' And ", "") + " TdCocd= '" + ms_TrCoCd + "' and TdTc= '" + ws_TTc + "' and TdYy= '" + ws_TYy + "' and TdChr= '" + ws_TChr + "' and TdNo= " + CStr(wd_TNo)
  wsql = wsql + "Delete from Txn where " + IIF(gs_Partition = ctCurrPrtn, " TPrtKey='" + ctCurrPrtn + "' And ", "") + " TCocd= '" + ms_TrCoCd + "' and TTc= '" + ws_TTc + "' and TYy= '" + ws_TYy + "' and TChr= '" + ws_TChr + "' and TNo= " + CStr(wd_TNo)
  moCn.Execute wsql
  DispMsg Err.Description, etError
End Sub
    
Private Sub DispFra(ByVal pv_TreeFra As en_TreeFra)
  '*** This is the common procedure which is used to toggle between hiding and displaying
      'the frames (i.e. if the frame is visible then make it invisible and viseversa)
      'associated with a command button
  
  Select Case pv_TreeFra
  Case Is = CpyBag
    If FraCB.Visible = True Then
      FraCB.Visible = False
      FraCB.Enabled = False
      Call EnaDisaCmds(False, CmdCB)
      CmdCB.SetFocus
  Else
      FraCB.Visible = True
      FraCB.Enabled = True
      'vk.4 cursor setfocus will change as per scope/bag alias no. frame enable/disable status
      If FraCBScope.Enabled = True Then ADC("WCBBLOC").SetFocus Else ADC("wRfIdStr").SetFocus
      FraCB.ZOrder
      Call EnaDisaCmds(True, CmdCB)
    End If
  '578-33 frame visible/invisible for show totals
  Case Is = ShowTot
    If FraShowTot.Visible = True Then
      FraShowTot.Visible = False
      FraShowTot.Enabled = False
      Call EnaDisaCmds(False, CmdShowTot)
      CmdShowTot.SetFocus
    Else
      FraShowTot.Visible = True
      FraShowTot.Enabled = True
      ADC("TrNetMetWt").SetFocus
      FraShowTot.ZOrder
      Call EnaDisaCmds(True, CmdShowTot)
    End If
  End Select
End Sub
Private Sub EnaDisaCmds(ByVal pv_ShowFra As Boolean, Optional ByVal pv_CmdBut As MWCTL_BTN1, Optional ByVal pv_TreeFra As en_TreeFra)
  '*** This procedure is used to give the toggle option (enable and disable) to
      'the command buttons
  '*** storing the state of cmd buttons, when the frame is going to be vis.
  
  Dim X As Integer
  If pv_ShowFra = True Then
    Call StoreState
    
  End If
  
  If Not pv_CmdBut Is Nothing Then
    If pv_ShowFra = True Then
      '*** Jenny (Old Code Bef Color)
'      pv_CmdBut.BackColor = vbWhite
'      pv_CmdBut.FontBold = True
      '*** Jenny (Old Code Bef Color)
      If pv_CmdBut.Pic = 0 Then
        pv_CmdBut.BackColor = vbWhite
        'pv_CmdBut.FontBold = True
      End If
      pv_CmdBut.CausesValidation = True
    Else
      '*** Jenny (Old Code Bef Color)
'      pv_CmdBut.BackColor = &H8000000F
'      pv_CmdBut.FontBold = False
      '*** Jenny (Old Code Bef Color)
      If pv_CmdBut.Pic = 0 Then
        pv_CmdBut.BackColor = &H8000000F
        'pv_CmdBut.FontBold = False
      End If
      '### pv_CmdBut.CausesValidation = False
      '*** restoring the 'Causes Validation' state of the current pv_cmdbutton
      If Not IsMissing(pv_TreeFra) Then
        pv_CmdBut.CausesValidation = mArr_Cv(pv_TreeFra)
      End If
    End If
  End If
  FraNKeyAll.Enabled = Not pv_ShowFra
  ADC.AllowSave = Not pv_ShowFra
  ADC.AllowDelete = Not pv_ShowFra
  '578-33 command buttons will enable/disable
  If pv_ShowFra = True Then
    '** All command buttons are disabled
    CmdShowTot.Enabled = Not pv_ShowFra
    CmdCB.Enabled = Not pv_ShowFra
  Else
    '*** Restoring Ena/Disa state of all Command Buttons
    CmdShowTot.Enabled = mArr_Ed(ShowTot)
    CmdCB.Enabled = mArr_Ed(CpyBag)
  End If
  
  If Not pv_CmdBut Is Nothing Then pv_CmdBut.Enabled = True
 End Sub
Private Sub StoreState()
  '*** storing the Enabled and CausesValidation state of cmd buttons
  mArr_Cv(CpyBag) = CmdCB.CausesValidation
  mArr_Cv(ShowTot) = CmdShowTot.CausesValidation  '578-33
  mArr_Ed(CpyBag) = CmdCB.Enabled
  mArr_Ed(ShowTot) = CmdShowTot.Enabled '578-33
End Sub
Private Sub HideAllFras()
  '*** This procedure hides all the frames other than FraNKeyAll
  FraNKeyAll.Visible = True
  FraCB.Visible = False
  FraCB.Enabled = False
  FraShowTot.Visible = False  '578-33
  End Sub
Private Function TrIRVch_Check(ByVal pv_IssRec As String) As Boolean
'*** This procedure checks Issue or Receipt voucher exists or not
  Dim wIssArr() As String
  Dim wRecArr() As String

  wIssArr = IIF(ADC("TrIssVchNo") <> "" And pv_IssRec <> "R", Split(ADC("TrIssVchNo"), "/"), Split("0/0/0/0/0", "/"))
  wRecArr = IIF(ADC("TrRecVchNo") <> "" And pv_IssRec <> "I", Split(ADC("TrRecVchNo"), "/"), Split("0/0/0/0/0", "/"))

  If ADC("TrIssVchNo") <> "" Or ADC("TrRecVchNo") <> "" Then
    If moCn.RecSeek("Select 'x' from Txn where " + IIF(gs_Partition = ctCurrPrtn, " TPrtKey='" + ctCurrPrtn + "' And ", "") + " TCocd= '" + wIssArr(0) + "' and TTc= '" + wIssArr(1) + "' and TYy= '" + wIssArr(2) + "' and TChr= '" + wIssArr(3) + "' and TNo= " + wIssArr(4) + "  ") Or _
       moCn.RecSeek("Select 'x' from Txn where " + IIF(gs_Partition = ctCurrPrtn, " TPrtKey='" + ctCurrPrtn + "' And ", "") + " TCocd= '" + wRecArr(0) + "' and TTc= '" + wRecArr(1) + "' and TYy= '" + wRecArr(2) + "' and TChr= '" + wRecArr(3) + "' and TNo= " + wRecArr(4) + "  ") Then
          GrdTreeDet.AllowAdd = False
          mb_AddRow = False
          TrIRVch_Check = True
          If pv_IssRec = "I" Then
            DispMsg "Issue Voucher Already Exists.Please Delete the Issue Voucher and Continue..", etInfo: Exit Function
          ElseIf pv_IssRec = "R" Then
            DispMsg "Receipt Voucher Already Exists.Please Delete the Receipt Voucher and Continue..", etInfo: Exit Function
          Else
            DispMsg "Cannot Edit Value When Issue/Receipt Voucher Exists.", etInfo: Exit Function
          End If
    End If
  End If
  
  GrdTreeDet.AllowAdd = True
  mb_AddRow = True
End Function

'''  ***** Std Code not to be changed **********************************
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
  ws_BtnStr = "CmdCBGo"
  Call SetObjColors(Me, ws_BtnStr)
  '*** Jenny Colour
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
    
    Set moCn = Nothing    '*** (09/08/05)    '*** (Bef speed) bef 24/11/06
End Sub
Private Sub GrdTreeDet_Validate(Cancel As Boolean)
  Cancel = GrdTreeDet.Validate
End Sub
'Ig.217-adding new line after each scan
Private Sub Form_KeyPress(KeyAscii As Integer)
  If KeyAscii <> vbKeyReturn Then Exit Sub
  If UCase(Me.ActiveControl.Name) = UCase("GrdTreeDet") Then
    With Me.ActiveControl
        If .Col = 2 Then
          If .Rows - 1 = .Row Then
            .Col = 3
          End If
      End If
    End With
  End If
End Sub

Private Sub OptCpyBagFr_Click(Index As Integer)
'vk.4 enable/disable the frames as per the scope/bag alias no. radio button selected
  Select Case Index
  Case Is = 0
    FraCBScope.Enabled = True
    FraCBBag.Enabled = False
    ADC("wRfIdStr") = ""
  Case Is = 1
    FraCBScope.Enabled = False
    FraCBBag.Enabled = True
    ADC("wCbBLoc") = ""
    ADC("wCbBYyFr") = "": ADC("wCbBChrFr") = "": ADC("wCbBNoFr") = ""
    ADC("wCbBYyTo") = "": ADC("wCbBChrTo") = "": ADC("wCbBNoTo") = ""
    ADC("wCbOdTcFr") = "": ADC("wCbOdYyFr") = "": ADC("wCbOdChrFr") = "": ADC("wCbOdNoFr") = "": ADC("wCbOdSrFr") = ""
    ADC("wCbOdTcTo") = "": ADC("wCbOdYyTo") = "": ADC("wCbOdChrTo") = "": ADC("wCbOdNoTo") = "": ADC("wCbOdSrTo") = ""
    ADC("wCbDmCdFr") = "": ADC("wCbDmCdTo") = ""
    ADC("wCbOdKtFr") = "": ADC("wCbOdKtTo") = "": ADC("wCbOdColFr") = "": ADC("wCbOdColTo") = ""
  End Select
End Sub
'vk.4 tooltip for copy bag options
Private Sub OptCpyBagFr_GotFocus(Index As Integer)
  DispMsg OptCpyBagFr(Index).ToolTipText, etInfo
End Sub
'578-33 calculating totals
Private Sub GetTotals()
  Dim wd_BagPcs  As Double, i As Integer
  Dim wd_DcWt  As Double, wd_MetXWt As Double
  For i = 1 To GrdTreeDet.Rows - 1
    If Not GrdTreeDet.IsDel(i) Then
      wd_BagPcs = wd_BagPcs + GrdTreeDet.Value(i, "TrdBagPcs")
      wd_DcWt = wd_DcWt + GrdTreeDet.Value(i, "TrdDcWt")
      wd_MetXWt = wd_MetXWt + GrdTreeDet.Value(i, "TrdMetXWt")
    End If
  Next i
  ADC("wBagPcs") = wd_BagPcs
  ADC("wTotDcWt") = wd_DcWt
  ADC("wTotMetXWt") = wd_MetXWt
  '578-33 Metal NetW, Grs Met Reqd moved from header into totals frame
  ADC("TrNetMetWt") = (ADC("TrGrsWaxWt") - ADC("TRRUBWT") - (ADC("wTotDcWt") / 5 + ADC("wTotMetXWt"))) _
                         * MWLib.Div(moCn.GetFldVal(" select RmSpGrav from RmMst where " + IIF(gs_Partition = ctCurrPrtn, " RmPrtKey='" + ctCurrPrtn + "' And ", "") + " RmCd= '" + ADC("TrRmCd") + "'"), ADC("TrWaxSpGr"))
    
  ADC("WTrGrsMetReq") = ADC("TrNetMetWt") + ADC("TrAddMetWt")
  
End Sub


