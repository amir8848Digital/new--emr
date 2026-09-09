VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Object = "{AB2CC32A-E652-42A2-B341-26ECE02F6386}#2.1#0"; "MwfCtlforspeed8.ocx"
Begin VB.Form EmrFrmTxnRm 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Transaction Rm Entry"
   ClientHeight    =   9810
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   14250
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   9810
   ScaleWidth      =   14250
   WindowState     =   2  'Maximized
   Begin MwfCtl.MWCTL_BTN1 CmdShowTot 
      Height          =   465
      Left            =   2670
      TabIndex        =   267
      ToolTipText     =   "Show Totals"
      Top             =   9270
      Width           =   675
      _ExtentX        =   1191
      _ExtentY        =   820
      BackColor       =   -2147483633
      MaskColor       =   8454143
      ForeColor       =   11842740
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
   Begin MwfCtl.MWCTL_BTN1 CmdInvHd 
      Height          =   465
      Left            =   3330
      TabIndex        =   266
      ToolTipText     =   "Enter invoice header details"
      Top             =   9270
      Width           =   705
      _ExtentX        =   1244
      _ExtentY        =   820
      BackColor       =   -2147483633
      MaskColor       =   8454143
      ForeColor       =   11842740
      Caption         =   "&3 Inv Head"
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
   Begin MwfCtl.MWCTL_BTN1 CmdChgDt 
      Height          =   465
      Left            =   10725
      TabIndex        =   195
      ToolTipText     =   "Change Voucher Date"
      Top             =   9270
      Width           =   765
      _ExtentX        =   1349
      _ExtentY        =   820
      BackColor       =   -2147483633
      MaskColor       =   8454143
      Caption         =   "Chan&ge Date"
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
   Begin MSComDlg.CommonDialog CdImpExc 
      Left            =   3195
      Top             =   9270
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin MwfCtl.MWCTL_BTN1 CmdImpExc 
      Height          =   465
      Left            =   9930
      TabIndex        =   194
      ToolTipText     =   "Copy Option"
      Top             =   9270
      Width           =   795
      _ExtentX        =   1402
      _ExtentY        =   820
      BackColor       =   -2147483633
      MaskColor       =   8454143
      Caption         =   "&Import Excel"
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
   Begin MwfCtl.MWCTL_BTN1 CmdCpLocRm 
      Height          =   465
      Left            =   9150
      TabIndex        =   193
      ToolTipText     =   "Copy Option"
      Top             =   9270
      Width           =   795
      _ExtentX        =   1402
      _ExtentY        =   820
      BackColor       =   -2147483633
      MaskColor       =   8454143
      Caption         =   "&Move  Rm"
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
   Begin MwfCtl.MWCTL_BTN1 CmdCpyDt 
      Height          =   465
      Left            =   8295
      TabIndex        =   192
      ToolTipText     =   "Copy From DT"
      Top             =   9270
      Width           =   855
      _ExtentX        =   1508
      _ExtentY        =   820
      BackColor       =   -2147483633
      MaskColor       =   8454143
      Caption         =   "Copy Fr &DT/Mlt"
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
   Begin MwfCtl.MWCTL_BTN1 CmdContra 
      Height          =   465
      Left            =   4815
      TabIndex        =   123
      ToolTipText     =   "Copy Option"
      Top             =   9270
      Width           =   765
      _ExtentX        =   1349
      _ExtentY        =   820
      BackColor       =   -2147483633
      MaskColor       =   8454143
      Caption         =   "Con&tra"
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
   Begin MwfCtl.MWCTL_BTN1 CmdCpyRet 
      Height          =   465
      Left            =   7200
      TabIndex        =   127
      ToolTipText     =   "Copy RM Details From Return Memo"
      Top             =   9270
      Width           =   1095
      _ExtentX        =   1931
      _ExtentY        =   820
      BackColor       =   -2147483633
      MaskColor       =   8454143
      Caption         =   "C&opy From Ret Memo"
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
   Begin MwfCtl.MWCTL_BTN1 CmdLN 
      CausesValidation=   0   'False
      Height          =   465
      Left            =   6375
      TabIndex        =   126
      ToolTipText     =   "Voucher Raw Material Summary"
      Top             =   9270
      Width           =   825
      _ExtentX        =   1455
      _ExtentY        =   820
      BackColor       =   -2147483633
      MaskColor       =   8454143
      Caption         =   "&New Lot No"
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
      Left            =   9450
      TabIndex        =   37
      TabStop         =   0   'False
      Top             =   9300
      Visible         =   0   'False
      Width           =   555
      _ExtentX        =   979
      _ExtentY        =   714
   End
   Begin VB.Frame FRA_KEYS 
      BorderStyle     =   0  'None
      Height          =   285
      Left            =   0
      TabIndex        =   74
      Top             =   0
      Width           =   4830
      Begin MwfCtl.MWCTL_MED ATXT 
         Height          =   285
         Index           =   1
         Left            =   1605
         TabIndex        =   0
         ToolTipText     =   "Enter Voucher Tc"
         Top             =   0
         Width           =   570
         _ExtentX        =   1005
         _ExtentY        =   503
         Enabled         =   0   'False
         MaxLength       =   3
         DataType        =   4
         DataField       =   "TTc"
         ReCalcParent    =   "TCHR"
         IdName          =   "TTC"
      End
      Begin MwfCtl.MWCTL_MED ATXT 
         Height          =   285
         Index           =   4
         Left            =   3495
         TabIndex        =   3
         ToolTipText     =   "Enter Voucher Number"
         Top             =   0
         Width           =   930
         _ExtentX        =   1640
         _ExtentY        =   503
         Alignment       =   2
         Mask            =   "######0"
         MaxLength       =   7
         DataType        =   1
         DataField       =   "TNo"
         IdName          =   "TNO"
      End
      Begin MwfCtl.MWCTL_MED ATXT 
         Height          =   285
         Index           =   6
         Left            =   2280
         TabIndex        =   1
         ToolTipText     =   "Enter Voucher Year"
         Top             =   0
         Width           =   435
         _ExtentX        =   767
         _ExtentY        =   503
         MaxLength       =   2
         DataType        =   4
         DataField       =   "TYy"
         IdName          =   "TYY"
      End
      Begin MwfCtl.MWCTL_MED ATXT 
         Height          =   285
         Index           =   41
         Left            =   2820
         TabIndex        =   2
         ToolTipText     =   "Enter Voucher Character"
         Top             =   0
         Width           =   570
         _ExtentX        =   1005
         _ExtentY        =   503
         MaxLength       =   3
         DataType        =   4
         DataField       =   "TChr"
         ReCalcParent    =   "TFRRMLOC,TTORMLOC,TFRBLOC,TACK"
         ReCalcOn        =   "TTC"
         IdName          =   "TCHR"
      End
      Begin MwfCtl.MWCTL_MED ATXT 
         Height          =   285
         Index           =   35
         Left            =   938
         TabIndex        =   109
         Top             =   0
         Width           =   585
         _ExtentX        =   1032
         _ExtentY        =   503
         MaxLength       =   3
         DataType        =   4
         DataField       =   "TCoCd"
         IdName          =   "TCOCD"
      End
      Begin MwfCtl.MWCTL_MED ATXT 
         Height          =   285
         Index           =   139
         Left            =   4440
         TabIndex        =   189
         Top             =   0
         Width           =   345
         _ExtentX        =   609
         _ExtentY        =   503
         Enabled         =   0   'False
         DataField       =   "TPrtKey"
         IdName          =   "TPRTKEY"
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
         Index           =   30
         Left            =   1530
         TabIndex        =   158
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
         Left            =   0
         TabIndex        =   90
         Top             =   0
         Width           =   855
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
         Left            =   2730
         TabIndex        =   81
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
         Left            =   3405
         TabIndex        =   80
         Top             =   0
         Width           =   135
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
         Left            =   2190
         TabIndex        =   76
         Top             =   0
         Width           =   135
      End
   End
   Begin MwfCtl.MWCTL_BTN1 CmdVS 
      Height          =   465
      Left            =   4935
      TabIndex        =   124
      ToolTipText     =   "Voucher Raw Material Summary"
      Top             =   9270
      Width           =   765
      _ExtentX        =   1349
      _ExtentY        =   820
      BackColor       =   -2147483633
      MaskColor       =   8454143
      Caption         =   "&Vch Summ"
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
   Begin MwfCtl.MWCTL_BTN1 CmdCP 
      Height          =   465
      Left            =   5580
      TabIndex        =   125
      ToolTipText     =   "Copy Option"
      Top             =   9270
      Width           =   795
      _ExtentX        =   1402
      _ExtentY        =   820
      BackColor       =   -2147483633
      MaskColor       =   8454143
      Caption         =   "&Copy"
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
   Begin MwfCtl.MWCTL_BTN1 CmdGV 
      Height          =   465
      Left            =   3990
      TabIndex        =   122
      ToolTipText     =   "Generate Voucher"
      Top             =   9270
      Width           =   825
      _ExtentX        =   1455
      _ExtentY        =   820
      BackColor       =   -2147483633
      MaskColor       =   8454143
      Caption         =   "G&en Vch"
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
   Begin MwfCtl.MWCTL_ADC ADC 
      CausesValidation=   0   'False
      Height          =   555
      Left            =   0
      TabIndex        =   38
      Top             =   9270
      Width           =   14115
      _ExtentX        =   24897
      _ExtentY        =   979
      BtnWidth        =   885
      BtnHeight       =   435
      BtnMaskColor    =   8454143
      FontSz          =   8.25
   End
   Begin VB.Frame FRA_NKEYS 
      BorderStyle     =   0  'None
      Enabled         =   0   'False
      Height          =   9750
      Left            =   -120
      TabIndex        =   73
      Top             =   -120
      Width           =   15345
      Begin VB.Frame FraGrdTxndADets 
         Height          =   675
         Left            =   120
         TabIndex        =   226
         Top             =   7560
         Width           =   14115
         Begin VB.Frame FraGrdTxndADetsDisp 
            Height          =   675
            Left            =   1290
            TabIndex        =   227
            Top             =   0
            Width           =   12825
            Begin MwfCtl.MWCTL_MED ATXT 
               Height          =   285
               Index           =   90
               Left            =   1470
               TabIndex        =   230
               ToolTipText     =   "Enter Order Year"
               Top             =   225
               Width           =   330
               _ExtentX        =   582
               _ExtentY        =   503
               MaxLength       =   2
               DataType        =   4
               IdName          =   "WTDBSTODYY"
            End
            Begin MwfCtl.MWCTL_MED ATXT 
               Height          =   285
               Index           =   91
               Left            =   2550
               TabIndex        =   232
               ToolTipText     =   "Enter Order Number"
               Top             =   225
               Width           =   750
               _ExtentX        =   1323
               _ExtentY        =   503
               Alignment       =   2
               Mask            =   "#####0"
               MaxLength       =   6
               DataType        =   1
               IdName          =   "WTDBSTODNO"
            End
            Begin MwfCtl.MWCTL_MED ATXT 
               Height          =   285
               Index           =   92
               Left            =   810
               TabIndex        =   229
               ToolTipText     =   "Enter Order Tc"
               Top             =   225
               Width           =   570
               _ExtentX        =   1005
               _ExtentY        =   503
               MaxLength       =   3
               DataType        =   4
               IdName          =   "WTDBSTODTC"
            End
            Begin MwfCtl.MWCTL_MED ATXT 
               Height          =   285
               Index           =   93
               Left            =   1890
               TabIndex        =   231
               ToolTipText     =   "Enter Order Character"
               Top             =   225
               Width           =   570
               _ExtentX        =   1005
               _ExtentY        =   503
               MaxLength       =   3
               DataType        =   4
               IdName          =   "WTDBSTODCHR"
            End
            Begin MwfCtl.MWCTL_MED ATXT 
               Height          =   285
               Index           =   94
               Left            =   5520
               TabIndex        =   234
               ToolTipText     =   "Enter Design Code"
               Top             =   225
               Width           =   2385
               _ExtentX        =   4207
               _ExtentY        =   503
               MaxLength       =   15
               DataType        =   4
               IdName          =   "WTDBSTDMCD"
            End
            Begin MwfCtl.MWCTL_MED ATXT 
               Height          =   285
               Index           =   105
               Left            =   3375
               TabIndex        =   233
               ToolTipText     =   "Enter Order Serial Number – ONLY for reference will not be used for BEST Used in Customs Matching"
               Top             =   225
               Width           =   750
               _ExtentX        =   1323
               _ExtentY        =   503
               Alignment       =   2
               Mask            =   "####0"
               MaxLength       =   5
               DataType        =   1
               IdName          =   "WTDBSTODSR"
            End
            Begin VB.Label LblOrdSlash4 
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
               Left            =   3300
               TabIndex        =   265
               Top             =   225
               Width           =   105
            End
            Begin VB.Label LblOrdSlash1 
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
               Left            =   1410
               TabIndex        =   239
               Top             =   225
               Width           =   105
            End
            Begin VB.Label LblOrdSlash2 
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
               Left            =   1815
               TabIndex        =   238
               Top             =   225
               Width           =   105
            End
            Begin VB.Label LblOrdSlash3 
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
               Left            =   2475
               TabIndex        =   237
               Top             =   225
               Width           =   105
            End
            Begin VB.Label LblwTdOrd 
               BackStyle       =   0  'Transparent
               Caption         =   "Order"
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
               TabIndex        =   236
               Top             =   225
               Width           =   615
            End
            Begin VB.Label LblwTdDmCd 
               BackStyle       =   0  'Transparent
               Caption         =   "Design"
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
               Left            =   4710
               TabIndex        =   235
               Top             =   225
               Width           =   750
            End
         End
         Begin MwfCtl.MWCTL_BTN1 CmdGrdTxndADets 
            Height          =   350
            Left            =   135
            TabIndex        =   228
            Top             =   180
            Width           =   1065
            _ExtentX        =   1879
            _ExtentY        =   609
            BackColor       =   -2147483633
            MaskColor       =   8454143
            Caption         =   "&Enter Dets"
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
      End
      Begin VB.Frame FraNKeyAll 
         BorderStyle     =   0  'None
         Height          =   9255
         Left            =   120
         TabIndex        =   75
         Top             =   120
         Width           =   14175
         Begin MwfCtl.MWCTL_MED ATXT 
            Height          =   285
            Index           =   27
            Left            =   795
            TabIndex        =   30
            ToolTipText     =   "Enter Raw Material Description"
            Top             =   7155
            Width           =   12210
            _ExtentX        =   21537
            _ExtentY        =   503
            MaxLength       =   60
            ReCalcParent    =   "GrdTxndA(TDDESC)"
            IdName          =   "WTDDESC"
         End
         Begin MwfCtl.MWCTL_MED ATXT 
            Height          =   285
            Index           =   22
            Left            =   7425
            TabIndex        =   6
            ToolTipText     =   "Enter Raw Material Default To Location"
            Top             =   270
            Width           =   1410
            _ExtentX        =   2487
            _ExtentY        =   503
            MaxLength       =   8
            DataType        =   4
            DataField       =   "TToRmLoc"
            ReCalcParent    =   "TACK"
            ReCalcOn        =   "TCHR"
            IdName          =   "TTORMLOC"
         End
         Begin MwfCtl.MWCTL_MED ATXT 
            Height          =   285
            Index           =   30
            Left            =   5355
            TabIndex        =   4
            ToolTipText     =   "Enter Voucher Date"
            Top             =   0
            Width           =   855
            _ExtentX        =   1508
            _ExtentY        =   503
            MaxLength       =   8
            DataType        =   3
            DataField       =   "TDt"
            ReCalcParent    =   "TCNVRT,TCNVRTRS"
            IdName          =   "TDT"
         End
         Begin MwfCtl.MWCTL_MED ATXT 
            Height          =   285
            Index           =   0
            Left            =   7425
            TabIndex        =   5
            ToolTipText     =   "Enter Raw Material Default From Location"
            Top             =   0
            Width           =   1410
            _ExtentX        =   2487
            _ExtentY        =   503
            MaxLength       =   8
            DataType        =   4
            DataField       =   "TFrRmLoc"
            ReCalcParent    =   "TLSLOC,TACK"
            ReCalcOn        =   "TCHR"
            IdName          =   "TFRRMLOC"
         End
         Begin MwfCtl.MWCTL_MED ATXT 
            Height          =   285
            Index           =   2
            Left            =   12000
            TabIndex        =   13
            ToolTipText     =   "Bag Location"
            Top             =   0
            Visible         =   0   'False
            Width           =   210
            _ExtentX        =   370
            _ExtentY        =   503
            Enabled         =   0   'False
            MaxLength       =   5
            DataType        =   4
            DataField       =   "TFrBLoc"
            ReCalcOn        =   "TCHR"
            IdName          =   "TFRBLOC"
         End
         Begin MwfCtl.MWCTL_MED ATXT 
            Height          =   285
            Index           =   48
            Left            =   945
            TabIndex        =   11
            ToolTipText     =   "Enter Description"
            Top             =   570
            Width           =   9270
            _ExtentX        =   16351
            _ExtentY        =   503
            MaxLength       =   60
            DataField       =   "TDesc"
            IdName          =   "TDESC"
         End
         Begin MwfCtl.MWCTL_MED ATXT 
            Height          =   285
            Index           =   25
            Left            =   9885
            TabIndex        =   7
            ToolTipText     =   "Enter Default I/R"
            Top             =   0
            Width           =   330
            _ExtentX        =   582
            _ExtentY        =   503
            DataType        =   4
            IdName          =   "WFRRMIR"
         End
         Begin MwfCtl.MWCTL_MED ATXT 
            Height          =   285
            Index           =   5
            Left            =   12210
            TabIndex        =   107
            Top             =   0
            Visible         =   0   'False
            Width           =   165
            _ExtentX        =   291
            _ExtentY        =   503
            Enabled         =   0   'False
            MaxLength       =   8
            DataType        =   3
            DataField       =   "TUseBefDt"
            IdName          =   "TUSEBEFDT"
         End
         Begin MwfCtl.MWCTL_MED ATXT 
            Height          =   285
            Index           =   96
            Left            =   12390
            TabIndex        =   110
            ToolTipText     =   "Enter Import Type"
            Top             =   0
            Visible         =   0   'False
            Width           =   165
            _ExtentX        =   291
            _ExtentY        =   503
            Enabled         =   0   'False
            MaxLength       =   5
            DataType        =   4
            DataField       =   "TImpTyp"
            IdName          =   "TIMPTYP"
         End
         Begin MwfCtl.MWCTL_MED ATXT 
            Height          =   285
            Index           =   99
            Left            =   12570
            TabIndex        =   111
            ToolTipText     =   "Enter Quantity To Be Transferred"
            Top             =   0
            Visible         =   0   'False
            Width           =   195
            _ExtentX        =   344
            _ExtentY        =   503
            Enabled         =   0   'False
            Mask            =   "##0.0"
            MaxLength       =   5
            DataType        =   2
            DataField       =   "TTRFQTY"
            IdName          =   "TTRFQTY"
         End
         Begin MwfCtl.MWCTL_MED ATXT 
            Height          =   285
            Index           =   37
            Left            =   12780
            TabIndex        =   114
            Top             =   0
            Visible         =   0   'False
            Width           =   240
            _ExtentX        =   423
            _ExtentY        =   503
            Enabled         =   0   'False
            MaxLength       =   30
            DataField       =   "TSEOPWD"
            IdName          =   "TSEOPWD"
         End
         Begin MwfCtl.MWCTL_MED ATXT 
            Height          =   285
            Index           =   109
            Left            =   945
            TabIndex        =   9
            ToolTipText     =   "Enter Order Lock (Y/N)"
            Top             =   285
            Width           =   345
            _ExtentX        =   609
            _ExtentY        =   503
            DataType        =   4
            DataField       =   "TLockYN"
            IdName          =   "TLOCKYN"
         End
         Begin MwfCtl.MWCTL_MED ATXT 
            Height          =   285
            Index           =   52
            Left            =   11220
            TabIndex        =   12
            ToolTipText     =   "Enter Lot No."
            Top             =   570
            Width           =   2595
            _ExtentX        =   4577
            _ExtentY        =   503
            MaxLength       =   16
            DataType        =   4
            DataField       =   "TLOTNO"
            IdName          =   "TLOTNO"
         End
         Begin VB.TextBox txtPwd 
            Height          =   295
            IMEMode         =   3  'DISABLE
            Left            =   3495
            MaxLength       =   30
            PasswordChar    =   "*"
            TabIndex        =   10
            Top             =   285
            Width           =   2715
         End
         Begin VB.Frame FraPurHd 
            Height          =   1395
            Left            =   0
            TabIndex        =   91
            Top             =   780
            Width           =   14130
            Begin MwfCtl.MWCTL_MED ATXT 
               Height          =   285
               Index           =   11
               Left            =   1560
               TabIndex        =   17
               ToolTipText     =   "Enter Supplier/Customer Code"
               Top             =   390
               Visible         =   0   'False
               Width           =   1335
               _ExtentX        =   2355
               _ExtentY        =   503
               MaxLength       =   8
               DataType        =   4
               DataField       =   "TSuppCd"
               ReCalcParent    =   "WCMNAME,TCURRCD"
               ReCalcOn        =   "WCMNAME"
               IdName          =   "TSUPPCD"
            End
            Begin MwfCtl.MWCTL_MED ATXT 
               Height          =   285
               Index           =   13
               Left            =   9930
               TabIndex        =   24
               ToolTipText     =   "Enter Bill No"
               Top             =   660
               Visible         =   0   'False
               Width           =   3165
               _ExtentX        =   5583
               _ExtentY        =   503
               MaxLength       =   20
               DataType        =   4
               DataField       =   "TBillNo"
               IdName          =   "TBILLNO"
            End
            Begin MwfCtl.MWCTL_MED ATXT 
               Height          =   285
               Index           =   15
               Left            =   13200
               TabIndex        =   25
               ToolTipText     =   "Enter Bill Date"
               Top             =   675
               Visible         =   0   'False
               Width           =   855
               _ExtentX        =   1508
               _ExtentY        =   503
               MaxLength       =   8
               DataType        =   3
               DataField       =   "TBillDt"
               IdName          =   "TBILLDT"
            End
            Begin MwfCtl.MWCTL_MED ATXT 
               Height          =   285
               Index           =   16
               Left            =   13200
               TabIndex        =   28
               ToolTipText     =   "Enter Bill Of Entry Date"
               Top             =   960
               Visible         =   0   'False
               Width           =   855
               _ExtentX        =   1508
               _ExtentY        =   503
               MaxLength       =   8
               DataType        =   3
               DataField       =   "TBOEDt"
               IdName          =   "TBOEDT"
            End
            Begin MwfCtl.MWCTL_MED ATXT 
               Height          =   285
               Index           =   17
               Left            =   5580
               TabIndex        =   19
               ToolTipText     =   "Enter Conversion Rate"
               Top             =   390
               Visible         =   0   'False
               Width           =   1035
               _ExtentX        =   1826
               _ExtentY        =   503
               Alignment       =   2
               Mask            =   "###0.0000"
               MaxLength       =   9
               DataType        =   2
               DataField       =   "TCnvRt"
               ReCalcOn        =   "TCURRCD,TDT"
               IdName          =   "TCNVRT"
            End
            Begin MwfCtl.MWCTL_MED ATXT 
               Height          =   285
               Index           =   14
               Left            =   9930
               TabIndex        =   27
               ToolTipText     =   "Enter Bill Of Entry Number"
               Top             =   960
               Visible         =   0   'False
               Width           =   3165
               _ExtentX        =   5583
               _ExtentY        =   503
               MaxLength       =   20
               DataType        =   4
               DataField       =   "TBOENo"
               IdName          =   "TBOENO"
            End
            Begin MwfCtl.MWCTL_MED ATXT 
               Height          =   285
               Index           =   33
               Left            =   1560
               TabIndex        =   23
               ToolTipText     =   "Enter Supplier/Customer Name"
               Top             =   675
               Visible         =   0   'False
               Width           =   6285
               _ExtentX        =   11086
               _ExtentY        =   503
               MaxLength       =   45
               ReCalcParent    =   "TSUPPCD"
               ReCalcOn        =   "TSUPPCD"
               IdName          =   "WCMNAME"
            End
            Begin MwfCtl.MWCTL_MED ATXT 
               Height          =   285
               Index           =   39
               Left            =   3870
               TabIndex        =   18
               ToolTipText     =   "Enter Currency Code"
               Top             =   390
               Width           =   825
               _ExtentX        =   1455
               _ExtentY        =   503
               Enabled         =   0   'False
               MaxLength       =   5
               DataType        =   4
               DataField       =   "TCURRCD"
               ReCalcParent    =   "TMULDIV,TMULDIVRS,TCNVRT,TCNVRTRS"
               ReCalcOn        =   "TSUPPCD"
               IdName          =   "TCURRCD"
            End
            Begin MwfCtl.MWCTL_MED ATXT 
               Height          =   285
               Index           =   55
               Left            =   7500
               TabIndex        =   20
               ToolTipText     =   "Enter Multiply/Divide (M/D)"
               Top             =   390
               Width           =   345
               _ExtentX        =   609
               _ExtentY        =   503
               Enabled         =   0   'False
               DataType        =   4
               DataField       =   "TMulDiv"
               ReCalcOn        =   "TCURRCD"
               IdName          =   "TMULDIV"
            End
            Begin MwfCtl.MWCTL_MED ATXT 
               Height          =   285
               Index           =   65
               Left            =   9930
               TabIndex        =   21
               ToolTipText     =   "Enter Rupees Conversion Rate"
               Top             =   390
               Width           =   1035
               _ExtentX        =   1826
               _ExtentY        =   503
               Alignment       =   2
               Mask            =   "###0.0000"
               MaxLength       =   9
               DataType        =   2
               DataField       =   "TCnvRtRs"
               ReCalcOn        =   "TCURRCD,TDT"
               IdName          =   "TCNVRTRS"
            End
            Begin MwfCtl.MWCTL_MED ATXT 
               Height          =   285
               Index           =   66
               Left            =   12750
               TabIndex        =   22
               ToolTipText     =   "Enter Rupees Multiply/Divide (M/D)"
               Top             =   390
               Width           =   345
               _ExtentX        =   609
               _ExtentY        =   503
               Enabled         =   0   'False
               DataType        =   4
               DataField       =   "TMulDivRs"
               ReCalcOn        =   "TCURRCD"
               IdName          =   "TMULDIVRS"
            End
            Begin MwfCtl.MWCTL_MED ATXT 
               Height          =   285
               Index           =   86
               Left            =   1560
               TabIndex        =   26
               ToolTipText     =   "Enter Pay Terms"
               Top             =   960
               Width           =   7035
               _ExtentX        =   12409
               _ExtentY        =   503
               MaxLength       =   45
               DataField       =   "TSuppPayTerms"
               IdName          =   "TSUPPPAYTERMS"
            End
            Begin VB.Label ALBL 
               BackStyle       =   0  'Transparent
               Caption         =   "Pay Terms"
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
               Left            =   90
               TabIndex        =   220
               Top             =   990
               Width           =   1575
            End
            Begin VB.Label ALBL 
               BackStyle       =   0  'Transparent
               Caption         =   "Rs. Conv Rt"
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
               Index           =   29
               Left            =   8805
               TabIndex        =   157
               Top             =   390
               Width           =   1200
            End
            Begin VB.Label Label1 
               BackStyle       =   0  'Transparent
               Caption         =   "Rs.  Mul/Div"
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
               Left            =   11580
               TabIndex        =   156
               Top             =   390
               Width           =   1185
            End
            Begin VB.Label LblMulDiv 
               BackStyle       =   0  'Transparent
               Caption         =   "Mul/Div"
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
               TabIndex        =   153
               Top             =   390
               Width           =   735
            End
            Begin VB.Label LblMain 
               BackStyle       =   0  'Transparent
               BorderStyle     =   1  'Fixed Single
               Caption         =   "  Purchase Header"
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
               Left            =   -30
               TabIndex        =   134
               Top             =   90
               Width           =   14160
            End
            Begin VB.Label LblCurrCd 
               BackStyle       =   0  'Transparent
               Caption         =   "Curr Cd."
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
               Left            =   3060
               TabIndex        =   132
               Top             =   390
               Width           =   825
            End
            Begin VB.Label ALBL 
               BackStyle       =   0  'Transparent
               Caption         =   "Supplier Name"
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
               Index           =   13
               Left            =   90
               TabIndex        =   108
               Top             =   675
               Width           =   1590
            End
            Begin VB.Label lblSlash2 
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
               Left            =   13110
               TabIndex        =   97
               Top             =   960
               Width           =   120
            End
            Begin VB.Label lblSlash1 
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
               Left            =   13110
               TabIndex        =   96
               Top             =   675
               Width           =   120
            End
            Begin VB.Label ALBL 
               BackStyle       =   0  'Transparent
               Caption         =   "Conv Rt"
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
               Left            =   4740
               TabIndex        =   95
               Top             =   390
               Width           =   930
            End
            Begin VB.Label lblBOENo 
               BackStyle       =   0  'Transparent
               Caption         =   "BOE No/ Dt"
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
               Left            =   8805
               TabIndex        =   94
               Top             =   960
               Width           =   1290
            End
            Begin VB.Label lblBillNo 
               BackStyle       =   0  'Transparent
               Caption         =   "Bill No/ Dt"
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
               Left            =   8805
               TabIndex        =   93
               Top             =   675
               Width           =   1290
            End
            Begin VB.Label ALBL 
               BackStyle       =   0  'Transparent
               Caption         =   "Supplier Cd"
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
               Left            =   90
               TabIndex        =   92
               Top             =   390
               Width           =   1740
            End
         End
         Begin VB.Frame FraRecHd 
            Height          =   975
            Left            =   0
            TabIndex        =   98
            Top             =   900
            Width           =   13470
            Begin MwfCtl.MWCTL_MED ATXT 
               Height          =   285
               Index           =   9
               Left            =   3510
               TabIndex        =   15
               ToolTipText     =   "Enter Loss From Date"
               Top             =   390
               Width           =   855
               _ExtentX        =   1508
               _ExtentY        =   503
               MaxLength       =   8
               DataType        =   3
               DataField       =   "TLsFrDt"
               ReCalcParent    =   "TLSTODT"
               ReCalcOn        =   "TLSLOC"
               IdName          =   "TLSFRDT"
            End
            Begin MwfCtl.MWCTL_MED ATXT 
               Height          =   285
               Index           =   10
               Left            =   3510
               TabIndex        =   16
               ToolTipText     =   "Enter Loss To Date"
               Top             =   675
               Width           =   855
               _ExtentX        =   1508
               _ExtentY        =   503
               MaxLength       =   8
               DataType        =   3
               DataField       =   "TLsToDt"
               ReCalcOn        =   "TLSFRDT"
               IdName          =   "TLSTODT"
            End
            Begin MwfCtl.MWCTL_MED ATXT 
               Height          =   285
               Index           =   7
               Left            =   1140
               TabIndex        =   14
               ToolTipText     =   "Enter Loss Location"
               Top             =   390
               Width           =   885
               _ExtentX        =   1561
               _ExtentY        =   503
               MaxLength       =   5
               DataType        =   4
               DataField       =   "TLsLoc"
               ReCalcParent    =   "TLSFRDT"
               ReCalcOn        =   "TFRRMLOC"
               IdName          =   "TLSLOC"
            End
            Begin VB.Label LblMain 
               BackStyle       =   0  'Transparent
               BorderStyle     =   1  'Fixed Single
               Caption         =   "  Recovery Header"
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
               Left            =   -30
               TabIndex        =   135
               Top             =   90
               Width           =   13470
            End
            Begin VB.Label ALBL 
               BackStyle       =   0  'Transparent
               Caption         =   "Loss Loc"
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
               Index           =   19
               Left            =   120
               TabIndex        =   101
               Top             =   390
               Width           =   1035
            End
            Begin VB.Label ALBL 
               BackStyle       =   0  'Transparent
               Caption         =   "        To Dt"
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
               Left            =   2370
               TabIndex        =   100
               Top             =   645
               Width           =   1125
            End
            Begin VB.Label ALBL 
               BackStyle       =   0  'Transparent
               Caption         =   "Loss Fr Dt"
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
               Left            =   2370
               TabIndex        =   99
               Top             =   390
               Width           =   1125
            End
         End
         Begin MwfCtl.MWCTL_MED ATXT 
            Height          =   285
            Index           =   84
            Left            =   11760
            TabIndex        =   190
            ToolTipText     =   "Enter Customer Code"
            Top             =   0
            Visible         =   0   'False
            Width           =   255
            _ExtentX        =   450
            _ExtentY        =   503
            Enabled         =   0   'False
            MaxLength       =   8
            DataType        =   4
            DataField       =   "TCmCd"
            IdName          =   "TCMCD"
         End
         Begin MwfCtl.MWCTL_MED ATXT 
            Height          =   285
            Index           =   50
            Left            =   12960
            TabIndex        =   206
            ToolTipText     =   "Enter Multi Issue/ receipt"
            Top             =   0
            Visible         =   0   'False
            Width           =   195
            _ExtentX        =   344
            _ExtentY        =   503
            Enabled         =   0   'False
            DataType        =   4
            DataField       =   "TMultiIr"
            IdName          =   "TMULTIIR"
         End
         Begin MwfCtl.MWCTL_MED ATXT 
            Height          =   285
            Index           =   85
            Left            =   13200
            TabIndex        =   207
            ToolTipText     =   "Enter Import Category"
            Top             =   0
            Visible         =   0   'False
            Width           =   165
            _ExtentX        =   291
            _ExtentY        =   503
            Enabled         =   0   'False
            MaxLength       =   3
            DataType        =   4
            DataField       =   "TImpCtg"
            IdName          =   "TIMPCTG"
         End
         Begin MwfCtl.MWCTL_MED ATXT 
            Height          =   285
            Index           =   89
            Left            =   10800
            TabIndex        =   8
            ToolTipText     =   "Enter Raw Material Default To Location"
            Top             =   240
            Width           =   330
            _ExtentX        =   582
            _ExtentY        =   503
            DataType        =   4
            DataField       =   "TAck"
            ReCalcOn        =   "TCHR,TFRRMLOC,TTORMLOC"
            IdName          =   "TACK"
         End
         Begin MwfCtl.MWCTL_FLX GrdTxndA 
            Height          =   4980
            Left            =   0
            TabIndex        =   29
            Top             =   2160
            Width           =   14130
            _ExtentX        =   24924
            _ExtentY        =   8784
            Cols            =   28
            colname1        =   "TDSR"
            heading1        =   "Sr. "
            datafld1        =   "TdSr"
            datatype1       =   1
            mask1           =   "###0"
            recalcparent1   =   "GrdTxndA(TDLOTNO)"
            maxlength1      =   4
            tooltiptext1    =   "Enter Sr No."
            colname2        =   "TDFRRMLOC"
            heading2        =   "Fr Loc                 "
            datafld2        =   "TdFrRmLoc"
            datatype2       =   4
            recalcparent2   =   "GrdTxndA(TDDUSTWT),GrdTxndA(TDFRRMLOCTYP)"
            maxlength2      =   8
            tooltiptext2    =   "Enter Raw Material From Location"
            colname3        =   "WTDFRRMIR"
            heading3        =   "I/R "
            datatype3       =   4
            recalcon3       =   "GrdTxndA(TDFRRMDC)"
            recalcparent3   =   "GrdTxndA(TDFRRMDC)"
            maxlength3      =   1
            tooltiptext3    =   "Enter Issue Or Receipt"
            colname4        =   "TDFRRMDC"
            heading4        =   "D/C"
            datafld4        =   "TdFrRmDc"
            datatype4       =   4
            recalcon4       =   "GrdTxndA(WTDFRRMIR)"
            recalcparent4   =   "GrdTxndA(WTDFRRMIR),WCNVBASERTW,WCNVTOTWT,WCNVTOTQTY,WCNVTOTVALUE,WCNVTOTRSVALUE"
            maxlength4      =   1
            style4          =   0
            tooltiptext4    =   "Enter Debit Or Credit (D/C)"
            colname5        =   "TDRMCD"
            heading5        =   "Rm Code                                  "
            datafld5        =   "TdRmCd"
            datatype5       =   4
            recalcparent5   =   $"EmrFrmTxnRm.frx":0000
            maxlength5      =   16
            tooltiptext5    =   "Enter Raw Material Code"
            colname6        =   "TDLOTNO"
            heading6        =   "Lot No.                          "
            datafld6        =   "TdLotNo"
            datatype6       =   4
            recalcon6       =   "GrdTxndA(TDRMCD),GrdTxndA(TDSR)"
            recalcparent6   =   "GrdTxndA(TDRTBYQW),WCNVTOTRSVALUE,GrdTxndA(TDPURRT)"
            maxlength6      =   16
            tooltiptext6    =   "Enter Lot No."
            colname7        =   "TDRMSZ"
            heading7        =   "Ln/Sv/Ptr    "
            datafld7        =   "TdRmSz"
            datatype7       =   2
            mask7           =   "##0.0000"
            recalcon7       =   "GrdTxndA(TDRMCD)"
            maxlength7      =   8
            tooltiptext7    =   "Enter Raw Material Size"
            colname8        =   "TDRMSZ2"
            heading8        =   "Brdth   "
            datafld8        =   "TdRmSz2"
            datatype8       =   2
            mask8           =   "#0.000"
            maxlength8      =   6
            colname9        =   "TDRMSZ3"
            heading9        =   "Dpth    "
            datafld9        =   "TdRmSz3"
            datatype9       =   2
            mask9           =   "#0.000"
            maxlength9      =   6
            colname10       =   "TDRMSTKRT"
            heading10       =   "Stock Rate     "
            datafld10       =   "TdRmStkRt"
            datatype10      =   2
            mask10          =   "-######0.000"
            recalcon10      =   "GrdTxndA(TDRMCD),GrdTxndA(TDPURRT),GrdTxndA(TDRTBYQW)"
            maxlength10     =   12
            tooltiptext10   =   "Enter Raw Material Stock Rate"
            colname11       =   "TDPCPERCT"
            heading11       =   "Pc/Ct "
            datafld11       =   "TdPcPerCt"
            datatype11      =   1
            mask11          =   "##0"
            recalcon11      =   "GrdTxndA(TDRMCD),GrdTxndA(TDRMQTY),GrdTxndA(TDRMWT)"
            recalcparent11  =   "GrdTxndA(TDRMQTY)"
            maxlength11     =   3
            tooltiptext11   =   "Enter Average Piece Per Carat"
            colname12       =   "TDRMQTY"
            heading12       =   "Qty     "
            datafld12       =   "TdRmQty"
            datatype12      =   1
            mask12          =   "####0"
            recalcon12      =   "GrdTxndA(TDPCPERCT),GrdTxndA(TDRMWT),GrdTxndA(TDRMCD)"
            recalcparent12  =   "GrdTxndA(TDPCPERCT),GrdTxndA(TDPURAMT),WCNVTOTQTY,WCNVTOTRSVALUE"
            maxlength12     =   5
            tooltiptext12   =   "Enter Raw Material Quantity"
            colname13       =   "TDRMWT"
            heading13       =   "Rm Wt        "
            datafld13       =   "TdRmWt"
            datatype13      =   2
            mask13          =   "####0.000"
            recalcon13      =   "GrdTxndA(TDRMCD)"
            recalcparent13  =   "GrdTxndA(TDRMQTY),GrdTxndA(TDPCPERCT),GrdTxndA(TDPURAMT),WCNVBASERTW,WCNVTOTWT,WCNVTOTRSVALUE"
            maxlength13     =   9
            tooltiptext13   =   "Enter Raw Material Weight"
            colname14       =   "TDTORMLOC"
            heading14       =   "To Loc               "
            datafld14       =   "TdToRmLoc"
            datatype14      =   4
            recalcparent14  =   "GrdTxndA(TDDUSTWT),GrdTxndA(TDTORMLOCTYP)"
            maxlength14     =   8
            tooltiptext14   =   "Enter To Location"
            colname15       =   "TDRTBYQW"
            heading15       =   "Q/W"
            datafld15       =   "TdRtByQW"
            datatype15      =   4
            recalcon15      =   "GrdTxndA(TDRMCD),GrdTxndA(TDLOTNO)"
            recalcparent15  =   "GrdTxndA(TDPURAMT),GrdTxndA(TDRMSTKRT),GrdTxndA(TDPURRT)"
            maxlength15     =   1
            colname16       =   "TDPURRT"
            heading16       =   "Pur Rate          "
            datafld16       =   "TdPurRt"
            datatype16      =   2
            mask16          =   "######0.000"
            recalcon16      =   "GrdTxndA(TDRMCD),GrdTxndA(TDLOTNO),GrdTxndA(TDRTBYQW)"
            recalcparent16  =   "GrdTxndA(TDPURAMT),GrdTxndA(TDRMSTKRT),WCNVTOTRSVALUE"
            maxlength16     =   11
            tooltiptext16   =   "Enter Pur Rate"
            colname17       =   "TDPURAMT"
            heading17       =   "Pur Amount             "
            datafld17       =   "TdPurAmt"
            datatype17      =   2
            mask17          =   "#########0.00"
            recalcon17      =   "GrdTxndA(TDPURRT),GrdTxndA(TDRMWT),GrdTxndA(TDRMQTY),GrdTxndA(TDRTBYQW)"
            recalcparent17  =   "WCNVTOTVALUE"
            maxlength17     =   13
            tooltiptext17   =   "Enter Pur Amount"
            colname18       =   "TDDESC"
            datafld18       =   "TdDesc"
            recalcon18      =   "WTDDESC"
            maxlength18     =   60
            style18         =   0
            tooltiptext18   =   "Enter Description for the Raw Material"
            colname19       =   "TDDUSTWT"
            heading19       =   "Dust Wt      "
            datafld19       =   "TdDustWt"
            datatype19      =   2
            mask19          =   "####0.000"
            recalcon19      =   "GrdTxndA(TDFRRMLOC),GrdTxndA(TDTORMLOC)"
            maxlength19     =   9
            style19         =   2
            tooltiptext19   =   "Enter Dust Weight"
            colname20       =   "TDPRTKEY"
            heading20       =   "Prtn  "
            datafld20       =   "TdPrtKey"
            maxlength20     =   1
            style20         =   0
            colname21       =   "TDBSTODTC"
            datafld21       =   "TdBstOdTc"
            datatype21      =   4
            maxlength21     =   3
            style21         =   0
            colname22       =   "TDBSTODYY"
            datafld22       =   "TdBstOdYy"
            datatype22      =   4
            maxlength22     =   2
            style22         =   0
            colname23       =   "TDBSTODCHR"
            datafld23       =   "TdBstOdChr"
            datatype23      =   4
            maxlength23     =   3
            style23         =   0
            colname24       =   "TDBSTODNO"
            datafld24       =   "TdBstOdNo"
            datatype24      =   1
            mask24          =   "######0"
            maxlength24     =   7
            style24         =   0
            colname25       =   "TDBSTODSR"
            datafld25       =   "TdBstOdSr"
            datatype25      =   1
            mask25          =   "####0"
            maxlength25     =   5
            style25         =   0
            colname26       =   "TDBSTDMCD"
            datafld26       =   "TdBstDmCd"
            datatype26      =   4
            maxlength26     =   15
            style26         =   0
            colname27       =   "TDACK"
            heading27       =   "Ack"
            datafld27       =   "TdAck"
            datatype27      =   4
            maxlength27     =   1
            tooltiptext27   =   "Enter Acknowledgement Y/N"
         End
         Begin VB.Frame FraCNVOLOR 
            Height          =   1125
            Left            =   0
            TabIndex        =   249
            Top             =   8100
            Visible         =   0   'False
            Width           =   14130
            Begin MwfCtl.MWCTL_MED ATXT 
               Height          =   285
               Index           =   97
               Left            =   1830
               TabIndex        =   250
               ToolTipText     =   "Total Amount"
               Top             =   690
               Width           =   2145
               _ExtentX        =   3784
               _ExtentY        =   503
               Alignment       =   2
               Enabled         =   0   'False
               Mask            =   "-########0.0000"
               MaxLength       =   15
               DataType        =   2
               ReCalcParent    =   "WCNVRSRT"
               ReCalcOn        =   "GrdTxndA(TDRMQTY),GrdTxndA(TDRMWT),GrdTxndA(TDFRRMDC),GrdTxndA(TDRMCD),GrdTxndA(TdLotNo),GrdTxndA(TDPURRT)"
               IdName          =   "WCNVTOTRSVALUE"
            End
            Begin MwfCtl.MWCTL_MED ATXT 
               Height          =   285
               Index           =   103
               Left            =   1830
               TabIndex        =   251
               ToolTipText     =   "Net Amount"
               Top             =   390
               Width           =   2145
               _ExtentX        =   3784
               _ExtentY        =   503
               Alignment       =   2
               Enabled         =   0   'False
               Mask            =   "-########0.0000"
               MaxLength       =   15
               DataType        =   2
               ReCalcParent    =   "TTOTAMT,WCNVBASERTQ,WCNVBASERTW,WCNVRSRT"
               ReCalcOn        =   "GrdTxndA(TDFRRMDC),GrdTxndA(TDRMCD),GrdTxndA(TDPURAMT)"
               IdName          =   "WCNVTOTVALUE"
            End
            Begin MwfCtl.MWCTL_MED ATXT 
               Height          =   285
               Index           =   100
               Left            =   5070
               TabIndex        =   257
               Top             =   390
               Width           =   1215
               _ExtentX        =   2143
               _ExtentY        =   503
               Alignment       =   2
               Enabled         =   0   'False
               Mask            =   "######0"
               MaxLength       =   7
               DataType        =   2
               ReCalcParent    =   "WCNVBASERTQ,WCNVBASERTQ"
               ReCalcOn        =   "GrdTxndA(TDRMQTY),GrdTxndA(TDFRRMDC),GrdTxndA(TDRMCD)"
               IdName          =   "WCNVTOTQTY"
            End
            Begin MwfCtl.MWCTL_MED ATXT 
               Height          =   285
               Index           =   101
               Left            =   5070
               TabIndex        =   258
               Top             =   690
               Width           =   1215
               _ExtentX        =   2143
               _ExtentY        =   503
               Alignment       =   2
               Enabled         =   0   'False
               Mask            =   "######0.000"
               MaxLength       =   11
               DataType        =   2
               ReCalcParent    =   "WCNVBASERTW"
               ReCalcOn        =   "GrdTxndA(TDRMWT),GrdTxndA(TDFRRMDC),GrdTxndA(TDRMCD)"
               IdName          =   "WCNVTOTWT"
            End
            Begin MwfCtl.MWCTL_MED ATXT 
               Height          =   285
               Index           =   104
               Left            =   8250
               TabIndex        =   259
               ToolTipText     =   "Total Stk Rate"
               Top             =   360
               Width           =   1215
               _ExtentX        =   2143
               _ExtentY        =   503
               Alignment       =   2
               Enabled         =   0   'False
               Mask            =   "-#######0.00"
               MaxLength       =   12
               DataType        =   2
               ReCalcOn        =   "WCNVTOTQTY,WCNVTOTVALUE"
               IdName          =   "WCNVBASERTQ"
            End
            Begin MwfCtl.MWCTL_MED ATXT 
               Height          =   285
               Index           =   98
               Left            =   8250
               TabIndex        =   261
               ToolTipText     =   "Total Stk Rate"
               Top             =   660
               Width           =   1215
               _ExtentX        =   2143
               _ExtentY        =   503
               Alignment       =   2
               Enabled         =   0   'False
               Mask            =   "-#######0.00"
               MaxLength       =   12
               DataType        =   2
               ReCalcOn        =   "WCNVTOTWT,WCNVTOTVALUE"
               IdName          =   "WCNVBASERTW"
            End
            Begin MwfCtl.MWCTL_MED ATXT 
               Height          =   285
               Index           =   102
               Left            =   11580
               TabIndex        =   263
               ToolTipText     =   "Enter Rupees Conversion Rate"
               Top             =   360
               Width           =   1035
               _ExtentX        =   1826
               _ExtentY        =   503
               Alignment       =   2
               Enabled         =   0   'False
               Mask            =   "###0.0000"
               MaxLength       =   9
               DataType        =   2
               ReCalcOn        =   "WCNVTOTRSVALUE,WCNVTOTVALUE"
               IdName          =   "WCNVRSRT"
            End
            Begin VB.Label ALBL 
               BackStyle       =   0  'Transparent
               Caption         =   "Rs. Conv Rt"
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
               Left            =   10410
               TabIndex        =   264
               Top             =   360
               Width           =   1200
            End
            Begin VB.Label Label10 
               BackStyle       =   0  'Transparent
               Caption         =   "Base Rt / W"
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
               Left            =   6750
               TabIndex        =   262
               Top             =   720
               Width           =   1485
            End
            Begin VB.Label Label11 
               BackStyle       =   0  'Transparent
               Caption         =   "Base Rt / Q"
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
               Left            =   6750
               TabIndex        =   260
               Top             =   390
               Width           =   1485
            End
            Begin VB.Label Label9 
               BackStyle       =   0  'Transparent
               Caption         =   "Tot Wt"
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
               Left            =   4350
               TabIndex        =   256
               Top             =   690
               Width           =   615
            End
            Begin VB.Label Label8 
               BackStyle       =   0  'Transparent
               Caption         =   "Tot Qty"
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
               Left            =   4260
               TabIndex        =   255
               Top             =   390
               Width           =   705
            End
            Begin VB.Label Label7 
               BackStyle       =   0  'Transparent
               Caption         =   "Rs Value"
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
               TabIndex        =   254
               Top             =   675
               Width           =   1635
            End
            Begin VB.Label ALBL 
               BackStyle       =   0  'Transparent
               Caption         =   "Base Value"
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
               Left            =   120
               TabIndex        =   253
               Top             =   390
               Width           =   1635
            End
            Begin VB.Label LblMain 
               BackStyle       =   0  'Transparent
               BorderStyle     =   1  'Fixed Single
               Caption         =   "Conversion Footer"
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
               Index           =   13
               Left            =   -30
               TabIndex        =   252
               Top             =   90
               Width           =   14460
            End
         End
         Begin VB.Frame FraRecFot 
            Height          =   795
            Left            =   0
            TabIndex        =   102
            Top             =   8400
            Width           =   14130
            Begin MwfCtl.MWCTL_MED ATXT 
               Height          =   285
               Index           =   12
               Left            =   1440
               TabIndex        =   39
               ToolTipText     =   "Enter Gold In Loss"
               Top             =   390
               Width           =   1065
               _ExtentX        =   1879
               _ExtentY        =   503
               Alignment       =   2
               Mask            =   "####0.000"
               MaxLength       =   9
               DataType        =   2
               DataField       =   "TGldInLs"
               ReCalcParent    =   "WRECPRC"
               IdName          =   "TGLDINLS"
               Data            =   "99999.999"
            End
            Begin MwfCtl.MWCTL_MED ATXT 
               Height          =   285
               Index           =   8
               Left            =   4500
               TabIndex        =   40
               ToolTipText     =   "Enter Gold Recovered"
               Top             =   390
               Width           =   1065
               _ExtentX        =   1879
               _ExtentY        =   503
               Alignment       =   2
               Mask            =   "####0.000"
               MaxLength       =   9
               DataType        =   2
               DataField       =   "TGldRec"
               ReCalcParent    =   "WRECPRC"
               IdName          =   "TGLDREC"
            End
            Begin MwfCtl.MWCTL_MED ATXT 
               Height          =   285
               Index           =   3
               Left            =   10110
               TabIndex        =   48
               ToolTipText     =   "Recovery Percentage"
               Top             =   390
               Width           =   825
               _ExtentX        =   1455
               _ExtentY        =   503
               Alignment       =   2
               Mask            =   "##0.000"
               MaxLength       =   7
               DataType        =   2
               ReCalcOn        =   "TGLDINLS,TGLDREC"
               IdName          =   "WRECPRC"
            End
            Begin MwfCtl.MWCTL_MED ATXT 
               Height          =   285
               Index           =   26
               Left            =   7080
               TabIndex        =   45
               ToolTipText     =   "Enter Dust Weight"
               Top             =   390
               Width           =   1065
               _ExtentX        =   1879
               _ExtentY        =   503
               Alignment       =   2
               Mask            =   "####0.000"
               MaxLength       =   9
               DataType        =   2
               DataField       =   "TDustWt"
               IdName          =   "TDUSTWT"
            End
            Begin VB.Label LblMain 
               BackStyle       =   0  'Transparent
               BorderStyle     =   1  'Fixed Single
               Caption         =   "  Recovery Footer"
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
               Index           =   4
               Left            =   -30
               TabIndex        =   136
               Top             =   90
               Width           =   14190
            End
            Begin VB.Label ALBL 
               BackStyle       =   0  'Transparent
               Caption         =   "Gold In Loss"
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
               Index           =   9
               Left            =   120
               TabIndex        =   106
               Top             =   390
               Width           =   1365
            End
            Begin VB.Label ALBL 
               BackStyle       =   0  'Transparent
               Caption         =   "Gold Recovered"
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
               Index           =   10
               Left            =   2730
               TabIndex        =   105
               Top             =   390
               Width           =   1755
            End
            Begin VB.Label ALBL 
               BackStyle       =   0  'Transparent
               Caption         =   "% age Recovery"
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
               Index           =   11
               Left            =   8340
               TabIndex        =   104
               Top             =   390
               Width           =   1785
            End
            Begin VB.Label ALBL 
               BackStyle       =   0  'Transparent
               Caption         =   "Dust Weight"
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
               Left            =   5760
               TabIndex        =   103
               Top             =   390
               Width           =   1335
            End
         End
         Begin VB.Label ALBL 
            BackStyle       =   0  'Transparent
            BorderStyle     =   1  'Fixed Single
            Caption         =   "Please click on ""Show Totals"" button below to view the voucher Totals"
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
            Left            =   0
            TabIndex        =   299
            Top             =   8160
            Width           =   7035
         End
         Begin VB.Label LblAck 
            BackStyle       =   0  'Transparent
            Caption         =   "Acknowledgement"
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
            Left            =   8880
            TabIndex        =   248
            Top             =   270
            Width           =   1815
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
            Height          =   285
            Left            =   10470
            TabIndex        =   140
            Top             =   570
            Width           =   780
         End
         Begin VB.Label ALBL 
            BackStyle       =   0  'Transparent
            Caption         =   "Vch Lock"
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
            Index           =   7
            Left            =   0
            TabIndex        =   117
            Top             =   285
            Width           =   1035
         End
         Begin VB.Label ALBL 
            BackStyle       =   0  'Transparent
            Caption         =   "Vchr Lock"
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
            Left            =   0
            TabIndex        =   116
            Top             =   0
            Width           =   1035
         End
         Begin VB.Label ALBL 
            BackStyle       =   0  'Transparent
            Caption         =   "Vchr Lock"
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
            Left            =   0
            TabIndex        =   115
            Top             =   0
            Width           =   1035
         End
         Begin VB.Label ALBL 
            BackStyle       =   0  'Transparent
            Caption         =   "Seo Pwd"
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
            Index           =   16
            Left            =   2490
            TabIndex        =   113
            Top             =   285
            Width           =   1005
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
            TabIndex        =   85
            Top             =   7155
            Width           =   660
         End
         Begin VB.Label ALBL 
            BackStyle       =   0  'Transparent
            Caption         =   "Rm Fr I/R"
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
            Index           =   8
            Left            =   8880
            TabIndex        =   84
            Top             =   0
            Width           =   990
         End
         Begin VB.Label LblTDesc 
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
            TabIndex        =   82
            Top             =   570
            Width           =   615
         End
         Begin VB.Label LblTToRmLoc 
            BackStyle       =   0  'Transparent
            Caption         =   "To Loc"
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
            Left            =   6660
            TabIndex        =   79
            Top             =   270
            Width           =   795
         End
         Begin VB.Label LblTFrRmLoc 
            BackStyle       =   0  'Transparent
            Caption         =   "Rm Fr Loc"
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
            Left            =   6330
            TabIndex        =   78
            Top             =   0
            Width           =   1125
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
            Left            =   4830
            TabIndex        =   77
            Top             =   0
            Width           =   855
         End
      End
      Begin VB.Frame FraCpyRet 
         Height          =   1665
         Left            =   3930
         TabIndex        =   147
         Top             =   7710
         Visible         =   0   'False
         Width           =   5625
         Begin MwfCtl.MWCTL_BTN1 CmdCpyRetGo 
            Height          =   495
            Left            =   4770
            TabIndex        =   36
            ToolTipText     =   "Click Here To Copy"
            Top             =   555
            Width           =   555
            _ExtentX        =   979
            _ExtentY        =   873
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
         Begin MwfCtl.MWCTL_MED ATXT 
            Height          =   285
            Index           =   231
            Left            =   1890
            TabIndex        =   31
            ToolTipText     =   "Enter Return Memo Year"
            Top             =   555
            Width           =   555
            _ExtentX        =   979
            _ExtentY        =   503
            MaxLength       =   2
            DataType        =   4
            IdName          =   "WCRREYY"
         End
         Begin MwfCtl.MWCTL_MED ATXT 
            Height          =   285
            Index           =   232
            Left            =   2505
            TabIndex        =   32
            ToolTipText     =   "Enter Return Memo Character"
            Top             =   555
            Width           =   705
            _ExtentX        =   1244
            _ExtentY        =   503
            MaxLength       =   3
            DataType        =   4
            IdName          =   "WCRRECHR"
         End
         Begin MwfCtl.MWCTL_MED ATXT 
            Height          =   285
            Index           =   233
            Left            =   3270
            TabIndex        =   33
            ToolTipText     =   "Enter Return Memo Number"
            Top             =   555
            Width           =   1065
            _ExtentX        =   1879
            _ExtentY        =   503
            Alignment       =   2
            Mask            =   "#####0"
            MaxLength       =   6
            DataType        =   1
            IdName          =   "WCRRENO"
         End
         Begin MwfCtl.MWCTL_MED ATXT 
            Height          =   285
            Index           =   56
            Left            =   1890
            TabIndex        =   34
            ToolTipText     =   "Enter From RM Category"
            Top             =   840
            Width           =   495
            _ExtentX        =   873
            _ExtentY        =   503
            MaxLength       =   2
            DataType        =   4
            IdName          =   "WCRRMCTGFR"
         End
         Begin MwfCtl.MWCTL_MED ATXT 
            Height          =   285
            Index           =   64
            Left            =   1890
            TabIndex        =   35
            ToolTipText     =   "Enter To RM Category"
            Top             =   1125
            Width           =   495
            _ExtentX        =   873
            _ExtentY        =   503
            MaxLength       =   2
            DataType        =   4
            IdName          =   "WCRRMCTGTO"
         End
         Begin VB.Label ALBL 
            BackStyle       =   0  'Transparent
            Caption         =   "RM Ctg To"
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
            Index           =   27
            Left            =   90
            TabIndex        =   155
            Top             =   1125
            Width           =   1695
         End
         Begin VB.Label ALBL 
            BackStyle       =   0  'Transparent
            Caption         =   "RM Ctg Fr"
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
            Index           =   17
            Left            =   90
            TabIndex        =   154
            Top             =   840
            Width           =   1695
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
            Index           =   183
            Left            =   2445
            TabIndex        =   152
            Top             =   585
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
            Index           =   184
            Left            =   3210
            TabIndex        =   151
            Top             =   555
            Width           =   105
         End
         Begin VB.Label ALBL 
            BackStyle       =   0  'Transparent
            Caption         =   "Return Memo No."
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
            Index           =   185
            Left            =   90
            TabIndex        =   150
            Top             =   555
            Width           =   1695
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
            Index           =   189
            Left            =   2460
            TabIndex        =   149
            Top             =   540
            Width           =   105
         End
         Begin VB.Label LblMain 
            BackStyle       =   0  'Transparent
            BorderStyle     =   1  'Fixed Single
            Caption         =   " Copy From Return Memo"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   300
            Index           =   19
            Left            =   -30
            TabIndex        =   148
            Top             =   90
            Width           =   5640
         End
      End
      Begin VB.Frame FraCpyDt 
         Height          =   1695
         Left            =   4800
         TabIndex        =   174
         Top             =   7650
         Width           =   4950
         Begin VB.OptionButton OptVch 
            Caption         =   "DT"
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
            Left            =   1560
            TabIndex        =   176
            Top             =   840
            Width           =   825
         End
         Begin VB.OptionButton OptVch 
            Caption         =   "Melting"
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
            Left            =   2400
            TabIndex        =   177
            Top             =   840
            Width           =   1425
         End
         Begin MwfCtl.MWCTL_BTN1 CmdCpyDtGo 
            Height          =   405
            Left            =   4215
            TabIndex        =   182
            ToolTipText     =   "Copy From DT"
            Top             =   1155
            Width           =   555
            _ExtentX        =   979
            _ExtentY        =   714
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
         Begin MwfCtl.MWCTL_MED ATXT 
            Height          =   285
            Index           =   78
            Left            =   1575
            TabIndex        =   178
            ToolTipText     =   "Enter Voucher Tc"
            Top             =   1155
            Width           =   570
            _ExtentX        =   1005
            _ExtentY        =   503
            Enabled         =   0   'False
            MaxLength       =   3
            DataType        =   4
            IdName          =   "WDTTC"
         End
         Begin MwfCtl.MWCTL_MED ATXT 
            Height          =   285
            Index           =   79
            Left            =   2250
            TabIndex        =   179
            ToolTipText     =   "Enter Voucher Year"
            Top             =   1155
            Width           =   330
            _ExtentX        =   582
            _ExtentY        =   503
            MaxLength       =   2
            DataType        =   4
            IdName          =   "WDTYY"
         End
         Begin MwfCtl.MWCTL_MED ATXT 
            Height          =   285
            Index           =   80
            Left            =   2685
            TabIndex        =   180
            ToolTipText     =   "Enter Voucher Character"
            Top             =   1155
            Width           =   570
            _ExtentX        =   1005
            _ExtentY        =   503
            MaxLength       =   3
            DataType        =   4
            IdName          =   "WDTCHR"
         End
         Begin MwfCtl.MWCTL_MED ATXT 
            Height          =   285
            Index           =   81
            Left            =   3360
            TabIndex        =   181
            ToolTipText     =   "Enter Voucher No."
            Top             =   1155
            Width           =   750
            _ExtentX        =   1323
            _ExtentY        =   503
            Alignment       =   2
            Mask            =   "#####0"
            MaxLength       =   6
            DataType        =   1
            IdName          =   "WDTNO"
         End
         Begin MwfCtl.MWCTL_MED ATXT 
            Height          =   285
            Index           =   82
            Left            =   1575
            TabIndex        =   175
            ToolTipText     =   "Enter Voucher Company Code"
            Top             =   510
            Width           =   750
            _ExtentX        =   1323
            _ExtentY        =   503
            MaxLength       =   3
            DataType        =   4
            IdName          =   "WDTCOCD"
         End
         Begin VB.Label ALBL 
            BackStyle       =   0  'Transparent
            Caption         =   "Source Vch"
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
            Left            =   90
            TabIndex        =   225
            Top             =   840
            Width           =   1425
         End
         Begin VB.Label Label6 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Company Code"
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
            Left            =   90
            TabIndex        =   188
            Top             =   510
            Width           =   1455
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
            Index           =   37
            Left            =   90
            TabIndex        =   187
            Top             =   1155
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
            Index           =   38
            Left            =   2145
            TabIndex        =   186
            Top             =   1155
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
            Index           =   39
            Left            =   2580
            TabIndex        =   185
            Top             =   1155
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
            Index           =   40
            Left            =   3255
            TabIndex        =   184
            Top             =   1155
            Width           =   105
         End
         Begin VB.Label LblMain 
            BackStyle       =   0  'Transparent
            BorderStyle     =   1  'Fixed Single
            Caption         =   "  Copy From Voucher"
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
            Index           =   10
            Left            =   -30
            TabIndex        =   183
            Top             =   90
            Width           =   4980
         End
      End
      Begin VB.Frame FraContra 
         Height          =   1545
         Left            =   4350
         TabIndex        =   166
         Top             =   7830
         Width           =   4410
         Begin MwfCtl.MWCTL_BTN1 CmdCntGo 
            Height          =   405
            Left            =   3705
            TabIndex        =   165
            Top             =   1035
            Width           =   555
            _ExtentX        =   979
            _ExtentY        =   714
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
         Begin MwfCtl.MWCTL_MED ATXT 
            Height          =   285
            Index           =   72
            Left            =   1665
            TabIndex        =   160
            ToolTipText     =   "Enter Voucher Tc"
            Top             =   720
            Width           =   570
            _ExtentX        =   1005
            _ExtentY        =   503
            MaxLength       =   3
            DataType        =   4
            IdName          =   "WCNTTC"
         End
         Begin MwfCtl.MWCTL_MED ATXT 
            Height          =   285
            Index           =   73
            Left            =   2340
            TabIndex        =   161
            ToolTipText     =   "Enter Voucher Year"
            Top             =   720
            Width           =   330
            _ExtentX        =   582
            _ExtentY        =   503
            MaxLength       =   2
            DataType        =   4
            IdName          =   "WCNTYY"
         End
         Begin MwfCtl.MWCTL_MED ATXT 
            Height          =   285
            Index           =   74
            Left            =   2790
            TabIndex        =   162
            ToolTipText     =   "Enter Voucher Character"
            Top             =   720
            Width           =   570
            _ExtentX        =   1005
            _ExtentY        =   503
            MaxLength       =   3
            DataType        =   4
            IdName          =   "WCNTCHR"
         End
         Begin MwfCtl.MWCTL_MED ATXT 
            Height          =   285
            Index           =   75
            Left            =   3450
            TabIndex        =   163
            ToolTipText     =   "Enter Voucher No."
            Top             =   720
            Width           =   840
            _ExtentX        =   1482
            _ExtentY        =   503
            Alignment       =   2
            Mask            =   "######0"
            MaxLength       =   7
            DataType        =   1
            IdName          =   "WCNTNO"
         End
         Begin MwfCtl.MWCTL_MED ATXT 
            Height          =   285
            Index           =   76
            Left            =   1665
            TabIndex        =   164
            ToolTipText     =   "Enter Voucher Tc"
            Top             =   1005
            Width           =   1140
            _ExtentX        =   2011
            _ExtentY        =   503
            MaxLength       =   5
            DataType        =   4
            IdName          =   "WCNLOCCD"
         End
         Begin MwfCtl.MWCTL_MED ATXT 
            Height          =   285
            Index           =   77
            Left            =   1665
            TabIndex        =   159
            ToolTipText     =   "Enter Voucher Tc"
            Top             =   420
            Width           =   750
            _ExtentX        =   1323
            _ExtentY        =   503
            MaxLength       =   3
            DataType        =   4
            IdName          =   "WCNCOCD"
         End
         Begin VB.Label Label4 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Location Code"
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
            Left            =   90
            TabIndex        =   173
            Top             =   1005
            Width           =   1380
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
            Index           =   35
            Left            =   90
            TabIndex        =   172
            Top             =   750
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
            Index           =   34
            Left            =   2250
            TabIndex        =   171
            Top             =   720
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
            Index           =   32
            Left            =   2685
            TabIndex        =   170
            Top             =   720
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
            Index           =   31
            Left            =   3375
            TabIndex        =   169
            Top             =   720
            Width           =   105
         End
         Begin VB.Label Label5 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Company Code"
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
            Left            =   90
            TabIndex        =   168
            Top             =   420
            Width           =   1455
         End
         Begin VB.Label LblMain 
            BackStyle       =   0  'Transparent
            BorderStyle     =   1  'Fixed Single
            Caption         =   "  Contra Voucher Details"
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
            Index           =   9
            Left            =   -30
            TabIndex        =   167
            Top             =   90
            Width           =   4410
         End
      End
      Begin VB.Frame FraCP 
         Height          =   2055
         Left            =   4560
         TabIndex        =   62
         Top             =   7320
         Width           =   6705
         Begin VB.OptionButton OptRt 
            Caption         =   "Sale Rate"
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
            Left            =   3120
            TabIndex        =   71
            Top             =   1560
            Width           =   1305
         End
         Begin VB.OptionButton OptRt 
            Caption         =   "Cost Rate"
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
            Left            =   1785
            TabIndex        =   247
            Top             =   1560
            Width           =   1305
         End
         Begin MwfCtl.MWCTL_BTN1 CmdCPGo 
            Height          =   405
            Left            =   5880
            TabIndex        =   72
            Top             =   1440
            Width           =   555
            _ExtentX        =   979
            _ExtentY        =   714
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
         Begin MwfCtl.MWCTL_MED ATXT 
            Height          =   285
            Index           =   57
            Left            =   1785
            TabIndex        =   64
            ToolTipText     =   "Enter Voucher Tc"
            Top             =   705
            Width           =   570
            _ExtentX        =   1005
            _ExtentY        =   503
            MaxLength       =   3
            DataType        =   4
            IdName          =   "WCPTTC"
         End
         Begin MwfCtl.MWCTL_MED ATXT 
            Height          =   285
            Index           =   58
            Left            =   2460
            TabIndex        =   65
            ToolTipText     =   "Enter Voucher Year"
            Top             =   705
            Width           =   330
            _ExtentX        =   582
            _ExtentY        =   503
            MaxLength       =   2
            DataType        =   4
            IdName          =   "WCPTYY"
         End
         Begin MwfCtl.MWCTL_MED ATXT 
            Height          =   285
            Index           =   59
            Left            =   2910
            TabIndex        =   66
            ToolTipText     =   "Enter Voucher Character"
            Top             =   705
            Width           =   570
            _ExtentX        =   1005
            _ExtentY        =   503
            MaxLength       =   3
            DataType        =   4
            IdName          =   "WCPTCHR"
         End
         Begin MwfCtl.MWCTL_MED ATXT 
            Height          =   285
            Index           =   60
            Left            =   3570
            TabIndex        =   67
            ToolTipText     =   "Enter Voucher No."
            Top             =   705
            Width           =   840
            _ExtentX        =   1482
            _ExtentY        =   503
            Alignment       =   2
            Mask            =   "######0"
            MaxLength       =   7
            DataType        =   1
            IdName          =   "WCPTNO"
         End
         Begin MwfCtl.MWCTL_MED ATXT 
            Height          =   285
            Index           =   36
            Left            =   1785
            TabIndex        =   68
            ToolTipText     =   "Enter Voucher Tc"
            Top             =   990
            Width           =   1680
            _ExtentX        =   2963
            _ExtentY        =   503
            MaxLength       =   8
            DataType        =   4
            IdName          =   "WCPGRPOPT"
         End
         Begin MwfCtl.MWCTL_MED ATXT 
            Height          =   285
            Index           =   51
            Left            =   5265
            TabIndex        =   69
            ToolTipText     =   "Copy with details"
            Top             =   1035
            Width           =   300
            _ExtentX        =   529
            _ExtentY        =   503
            DataType        =   4
            IdName          =   "WCPDETCHK"
         End
         Begin MwfCtl.MWCTL_MED ATXT 
            Height          =   285
            Index           =   83
            Left            =   1785
            TabIndex        =   63
            ToolTipText     =   "Enter Voucher Tc"
            Top             =   420
            Width           =   570
            _ExtentX        =   1005
            _ExtentY        =   503
            MaxLength       =   3
            DataType        =   4
            IdName          =   "WCPTCOCD"
         End
         Begin MwfCtl.MWCTL_MED ATXT 
            Height          =   285
            Index           =   95
            Left            =   1785
            TabIndex        =   70
            ToolTipText     =   "Enter Voucher Tc"
            Top             =   1280
            Width           =   1170
            _ExtentX        =   2064
            _ExtentY        =   503
            MaxLength       =   8
            DataType        =   4
            IdName          =   "WCPRMCTG"
         End
         Begin VB.Label ALBL 
            BackStyle       =   0  'Transparent
            Caption         =   "Src RmCtg"
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
            TabIndex        =   246
            Top             =   1280
            Width           =   1065
         End
         Begin VB.Label ALBL 
            BackStyle       =   0  'Transparent
            Caption         =   "Rate As"
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
            Left            =   120
            TabIndex        =   245
            Top             =   1560
            Width           =   1065
         End
         Begin VB.Label ALBL 
            BackStyle       =   0  'Transparent
            Caption         =   "Comp Cd"
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
            Index           =   41
            Left            =   90
            TabIndex        =   191
            Top             =   420
            Width           =   1065
         End
         Begin VB.Label LblMain 
            BackStyle       =   0  'Transparent
            BorderStyle     =   1  'Fixed Single
            Caption         =   "  Copy From"
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
            Index           =   7
            Left            =   -30
            TabIndex        =   139
            Top             =   90
            Width           =   4590
         End
         Begin VB.Label LblCpdChk 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Copy With Detail"
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
            Left            =   3570
            TabIndex        =   133
            Top             =   1035
            Width           =   1740
         End
         Begin VB.Label LblGrpOpt 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Group Option"
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
            Left            =   90
            TabIndex        =   112
            Top             =   990
            Width           =   1380
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
            Index           =   54
            Left            =   3495
            TabIndex        =   89
            Top             =   705
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
            Index           =   57
            Left            =   2805
            TabIndex        =   88
            Top             =   705
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
            Index           =   58
            Left            =   2370
            TabIndex        =   87
            Top             =   705
            Width           =   105
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
            Index           =   59
            Left            =   90
            TabIndex        =   86
            Top             =   705
            Width           =   1065
         End
      End
      Begin VB.Frame FraImpExc 
         Height          =   1515
         Left            =   5190
         TabIndex        =   208
         Top             =   7830
         Width           =   7515
         Begin MwfCtl.MWCTL_MED ATXT 
            Height          =   285
            Index           =   248
            Left            =   660
            TabIndex        =   209
            ToolTipText     =   "Excel File Name and Path"
            Top             =   555
            Width           =   6075
            _ExtentX        =   10716
            _ExtentY        =   503
            Enabled         =   0   'False
            MaxLength       =   250
            IdName          =   "WXLFILE"
         End
         Begin MwfCtl.MWCTL_MED ATXT 
            Height          =   285
            Index           =   249
            Left            =   750
            TabIndex        =   211
            ToolTipText     =   "Specify the Excel Sheet Number"
            Top             =   1065
            Width           =   615
            _ExtentX        =   1085
            _ExtentY        =   503
            Alignment       =   2
            Mask            =   "#0"
            MaxLength       =   2
            DataType        =   1
            IdName          =   "WXLSHTNO"
         End
         Begin MwfCtl.MWCTL_MED ATXT 
            Height          =   285
            Index           =   250
            Left            =   2220
            TabIndex        =   212
            ToolTipText     =   "Specify Excel From Row Number"
            Top             =   1065
            Width           =   675
            _ExtentX        =   1191
            _ExtentY        =   503
            Alignment       =   2
            Mask            =   "###0"
            MaxLength       =   4
            DataType        =   1
            IdName          =   "WXLFRROW"
         End
         Begin MwfCtl.MWCTL_MED ATXT 
            Height          =   285
            Index           =   251
            Left            =   3870
            TabIndex        =   213
            ToolTipText     =   "Specify Excel To Row Number"
            Top             =   1065
            Width           =   705
            _ExtentX        =   1244
            _ExtentY        =   503
            Alignment       =   2
            Mask            =   "###0"
            MaxLength       =   4
            DataType        =   1
            IdName          =   "WXLTOROW"
         End
         Begin MwfCtl.MWCTL_BTN1 CmdImpExcGo 
            Height          =   495
            Left            =   6840
            TabIndex        =   218
            ToolTipText     =   "Click Here To Copy"
            Top             =   960
            Width           =   555
            _ExtentX        =   979
            _ExtentY        =   873
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
         Begin MwfCtl.MWCTL_BTN1 CmdImpExcPath 
            Height          =   495
            Left            =   6840
            TabIndex        =   210
            ToolTipText     =   "Click Here To Copy"
            Top             =   450
            Width           =   555
            _ExtentX        =   979
            _ExtentY        =   873
            BackColor       =   -2147483633
            MaskColor       =   8454143
            Caption         =   "&Imp"
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
         Begin VB.Label LblMain 
            BackStyle       =   0  'Transparent
            BorderStyle     =   1  'Fixed Single
            Caption         =   " Import From Excel"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   300
            Index           =   11
            Left            =   -30
            TabIndex        =   219
            Top             =   90
            Width           =   7440
         End
         Begin VB.Label ALBL 
            BackStyle       =   0  'Transparent
            Caption         =   "File"
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
            Index           =   201
            Left            =   30
            TabIndex        =   217
            Top             =   555
            Width           =   615
         End
         Begin VB.Label ALBL 
            BackStyle       =   0  'Transparent
            Caption         =   "Sht No"
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
            Index           =   202
            Left            =   30
            TabIndex        =   216
            Top             =   1065
            Width           =   765
         End
         Begin VB.Label ALBL 
            BackStyle       =   0  'Transparent
            Caption         =   "Fr Rw"
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
            Index           =   203
            Left            =   1560
            TabIndex        =   215
            Top             =   1065
            Width           =   675
         End
         Begin VB.Label ALBL 
            BackStyle       =   0  'Transparent
            Caption         =   "To Rw"
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
            Index           =   204
            Left            =   3030
            TabIndex        =   214
            Top             =   1065
            Width           =   645
         End
      End
      Begin VB.Frame FraCpLocRm 
         Height          =   1395
         Left            =   4020
         TabIndex        =   196
         ToolTipText     =   "#0.0000"
         Top             =   7950
         Visible         =   0   'False
         Width           =   6975
         Begin MwfCtl.MWCTL_MED ATXT 
            Height          =   285
            Index           =   151
            Left            =   990
            TabIndex        =   198
            ToolTipText     =   "Enter From Raw Material Code"
            Top             =   945
            Width           =   2460
            _ExtentX        =   4339
            _ExtentY        =   503
            MaxLength       =   16
            DataType        =   4
            IdName          =   "WLOCRMCDFR"
         End
         Begin MwfCtl.MWCTL_MED ATXT 
            Height          =   285
            Index           =   152
            Left            =   3600
            TabIndex        =   199
            ToolTipText     =   "Enter To Raw Material Code"
            Top             =   945
            Width           =   2460
            _ExtentX        =   4339
            _ExtentY        =   503
            MaxLength       =   16
            DataType        =   4
            IdName          =   "WLOCRMCDTO"
         End
         Begin MwfCtl.MWCTL_MED ATXT 
            Height          =   285
            Index           =   153
            Left            =   990
            TabIndex        =   197
            ToolTipText     =   "Enter From Raw Material Category"
            Top             =   645
            Width           =   480
            _ExtentX        =   847
            _ExtentY        =   503
            MaxLength       =   3
            DataType        =   4
            IdName          =   "WLOCRMCTG"
         End
         Begin MwfCtl.MWCTL_BTN1 CmdCpLocRmGo 
            Height          =   495
            Left            =   6210
            TabIndex        =   200
            ToolTipText     =   "Press This Button To Copy"
            Top             =   735
            Width           =   555
            _ExtentX        =   979
            _ExtentY        =   873
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
            Height          =   285
            Index           =   42
            Left            =   990
            TabIndex        =   205
            Top             =   360
            Width           =   855
         End
         Begin VB.Label ALBL 
            BackStyle       =   0  'Transparent
            Caption         =   "Rm Code  "
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
            Index           =   99
            Left            =   30
            TabIndex        =   204
            Top             =   945
            Width           =   885
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
            Index           =   100
            Left            =   3600
            TabIndex        =   203
            Top             =   360
            Width           =   765
         End
         Begin VB.Label ALBL 
            BackStyle       =   0  'Transparent
            Caption         =   "RmCtg"
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
            Index           =   101
            Left            =   30
            TabIndex        =   202
            Top             =   645
            Width           =   795
         End
         Begin VB.Label LblMain 
            BackStyle       =   0  'Transparent
            BorderStyle     =   1  'Fixed Single
            Caption         =   "  Copy Rm"
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
            Index           =   16
            Left            =   -30
            TabIndex        =   201
            Top             =   90
            Width           =   7005
         End
      End
      Begin VB.Frame FraGV 
         Height          =   3075
         Left            =   3300
         TabIndex        =   118
         ToolTipText     =   "#0.0000"
         Top             =   6300
         Visible         =   0   'False
         Width           =   7455
         Begin VB.CheckBox ChkBoxArr 
            Alignment       =   1  'Right Justify
            Caption         =   "Remove Stk Val"
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
            Left            =   0
            TabIndex        =   223
            Tag             =   "ShowPnd"
            ToolTipText     =   "Check To Show Pending Records"
            Top             =   2720
            Width           =   1935
         End
         Begin VB.CheckBox ChkBoxArr 
            Alignment       =   1  'Right Justify
            Caption         =   "Valuation"
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
            Left            =   0
            TabIndex        =   221
            Tag             =   "ShowPnd"
            ToolTipText     =   "Check To Show Pending Records"
            Top             =   2415
            Width           =   1935
         End
         Begin MwfCtl.MWCTL_BTN1 CmdGVGo 
            Height          =   495
            Left            =   6750
            TabIndex        =   61
            ToolTipText     =   "Press This Button To Copy"
            Top             =   2505
            Width           =   555
            _ExtentX        =   979
            _ExtentY        =   873
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
         Begin MwfCtl.MWCTL_MED ATXT 
            Height          =   285
            Index           =   40
            Left            =   1740
            TabIndex        =   53
            ToolTipText     =   "Enter From Raw Material Sub Category"
            Top             =   1260
            Width           =   885
            _ExtentX        =   1561
            _ExtentY        =   503
            MaxLength       =   5
            DataType        =   4
            IdName          =   "WRMSCTGFR"
            CmpStr          =   "RmSCtg>="
         End
         Begin MwfCtl.MWCTL_MED ATXT 
            Height          =   285
            Index           =   42
            Left            =   1740
            TabIndex        =   55
            ToolTipText     =   "Enter From Raw Material Code"
            Top             =   1545
            Width           =   2460
            _ExtentX        =   4339
            _ExtentY        =   503
            MaxLength       =   16
            DataType        =   4
            IdName          =   "WRMCDFR"
            CmpStr          =   "RmCd>="
         End
         Begin MwfCtl.MWCTL_MED ATXT 
            Height          =   285
            Index           =   43
            Left            =   4440
            TabIndex        =   54
            ToolTipText     =   "Enter To Raw Material Sub Category"
            Top             =   1260
            Width           =   885
            _ExtentX        =   1561
            _ExtentY        =   503
            MaxLength       =   5
            DataType        =   4
            IdName          =   "WRMSCTGTO"
            CmpStr          =   "RmSCtg<="
         End
         Begin MwfCtl.MWCTL_MED ATXT 
            Height          =   285
            Index           =   44
            Left            =   4440
            TabIndex        =   56
            ToolTipText     =   "Enter To Raw Material Code"
            Top             =   1545
            Width           =   2460
            _ExtentX        =   4339
            _ExtentY        =   503
            MaxLength       =   16
            DataType        =   4
            IdName          =   "WRMCDTO"
            CmpStr          =   "RmCd<="
         End
         Begin MwfCtl.MWCTL_MED ATXT 
            Height          =   285
            Index           =   45
            Left            =   1740
            TabIndex        =   50
            ToolTipText     =   "Enter From Location"
            Top             =   690
            Width           =   1830
            _ExtentX        =   3228
            _ExtentY        =   503
            MaxLength       =   8
            DataType        =   4
            IdName          =   "WSLOCFR"
            CmpStr          =   "SLoc>="
         End
         Begin MwfCtl.MWCTL_MED ATXT 
            Height          =   285
            Index           =   38
            Left            =   1740
            TabIndex        =   52
            ToolTipText     =   "Enter From Raw Material Category"
            Top             =   975
            Width           =   480
            _ExtentX        =   847
            _ExtentY        =   503
            MaxLength       =   3
            DataType        =   4
            IdName          =   "WRMCTGFR"
            CmpStr          =   "RmCtg>="
         End
         Begin MwfCtl.MWCTL_MED ATXT 
            Height          =   285
            Index           =   46
            Left            =   4440
            TabIndex        =   51
            ToolTipText     =   "Enter To Location"
            Top             =   690
            Width           =   1815
            _ExtentX        =   3201
            _ExtentY        =   503
            MaxLength       =   8
            DataType        =   4
            IdName          =   "WSLOCTO"
            CmpStr          =   "SLoc<="
         End
         Begin MwfCtl.MWCTL_MED ATXT 
            Height          =   285
            Index           =   47
            Left            =   1740
            TabIndex        =   59
            ToolTipText     =   "Enter From Raw Material Size"
            Top             =   2115
            Width           =   900
            _ExtentX        =   1588
            _ExtentY        =   503
            Mask            =   "##0.0000"
            MaxLength       =   8
            DataType        =   2
            IdName          =   "WRMSZFR"
            CmpStr          =   "SRmSz>="
         End
         Begin MwfCtl.MWCTL_MED ATXT 
            Height          =   285
            Index           =   49
            Left            =   4440
            TabIndex        =   60
            ToolTipText     =   "Enter To Raw Material Size"
            Top             =   2115
            Width           =   900
            _ExtentX        =   1588
            _ExtentY        =   503
            Mask            =   "##0.0000"
            MaxLength       =   8
            DataType        =   2
            IdName          =   "WRMSZTO"
            CmpStr          =   "SRmSz<="
         End
         Begin MwfCtl.MWCTL_MED ATXT 
            Height          =   285
            Index           =   53
            Left            =   1740
            TabIndex        =   57
            ToolTipText     =   "Enter From Lot No"
            Top             =   1830
            Width           =   1860
            _ExtentX        =   3281
            _ExtentY        =   503
            MaxLength       =   12
            DataType        =   4
            IdName          =   "WGVLOTNOFR"
            CmpStr          =   "SLotNo>= "
         End
         Begin MwfCtl.MWCTL_MED ATXT 
            Height          =   285
            Index           =   54
            Left            =   4440
            TabIndex        =   58
            ToolTipText     =   "Enter To Lot No"
            Top             =   1830
            Width           =   1860
            _ExtentX        =   3281
            _ExtentY        =   503
            MaxLength       =   12
            DataType        =   4
            IdName          =   "WGVLOTNOTO"
            CmpStr          =   "SLotNo<= "
         End
         Begin MwfCtl.MWCTL_MED ATXT 
            Height          =   285
            Index           =   87
            Left            =   2040
            TabIndex        =   222
            ToolTipText     =   "Valuation ? (Yes / No)"
            Top             =   2400
            Width           =   285
            _ExtentX        =   503
            _ExtentY        =   503
            Enabled         =   0   'False
            DataType        =   4
            IdName          =   "WVALYN"
         End
         Begin MwfCtl.MWCTL_MED ATXT 
            Height          =   285
            Index           =   88
            Left            =   2040
            TabIndex        =   224
            ToolTipText     =   "Remove Stock Value with Zero Wt Or Qty ? (Yes / No)"
            Top             =   2720
            Width           =   285
            _ExtentX        =   503
            _ExtentY        =   503
            Enabled         =   0   'False
            DataType        =   4
            IdName          =   "WREMSTKVALYN"
         End
         Begin VB.Label LblGVLotNo 
            BackStyle       =   0  'Transparent
            Caption         =   "Lot No"
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
            TabIndex        =   146
            Top             =   1830
            Width           =   1305
         End
         Begin VB.Label LblMain 
            BackStyle       =   0  'Transparent
            BorderStyle     =   1  'Fixed Single
            Caption         =   "  Generate Voucher"
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
            Index           =   5
            Left            =   -30
            TabIndex        =   137
            Top             =   90
            Width           =   7455
         End
         Begin VB.Label ALBL 
            BackStyle       =   0  'Transparent
            Caption         =   "Rm Size"
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
            Left            =   30
            TabIndex        =   131
            Top             =   2115
            Width           =   1305
         End
         Begin VB.Label ALBL 
            BackStyle       =   0  'Transparent
            Caption         =   "RmCtg"
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
            Index           =   28
            Left            =   30
            TabIndex        =   130
            Top             =   975
            Width           =   1425
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
            Index           =   26
            Left            =   4440
            TabIndex        =   129
            Top             =   360
            Width           =   765
         End
         Begin VB.Label ALBL 
            BackStyle       =   0  'Transparent
            Caption         =   "Sub Ctg  "
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
            Left            =   30
            TabIndex        =   128
            Top             =   1260
            Width           =   1185
         End
         Begin VB.Label ALBL 
            BackStyle       =   0  'Transparent
            Caption         =   "Rm Code  "
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
            Left            =   30
            TabIndex        =   121
            Top             =   1545
            Width           =   1305
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
            Height          =   285
            Index           =   21
            Left            =   1740
            TabIndex        =   120
            Top             =   360
            Width           =   855
         End
         Begin VB.Label ALBL 
            BackStyle       =   0  'Transparent
            Caption         =   "Loc"
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
            Index           =   15
            Left            =   30
            TabIndex        =   119
            Top             =   690
            Width           =   1725
         End
      End
      Begin VB.Frame FraVS 
         Enabled         =   0   'False
         Height          =   3165
         Left            =   120
         TabIndex        =   49
         Top             =   6210
         Width           =   13470
         Begin MwfCtl.MWCTL_MED ATXT 
            Height          =   285
            Index           =   62
            Left            =   5160
            TabIndex        =   41
            ToolTipText     =   "Total Raw Material Quantity Issued"
            Top             =   2790
            Width           =   825
            _ExtentX        =   1455
            _ExtentY        =   503
            Alignment       =   2
            Enabled         =   0   'False
            Mask            =   "######0"
            MaxLength       =   7
            DataType        =   2
            IdName          =   "WVSISSTOTQTY"
         End
         Begin MwfCtl.MWCTL_FLX GrdVS 
            Height          =   2400
            Left            =   90
            TabIndex        =   44
            ToolTipText     =   "Raw Material Summary For The Entire Voucher"
            Top             =   390
            Width           =   12555
            _ExtentX        =   22146
            _ExtentY        =   4233
            Cols            =   11
            colname1        =   "WVSLOC"
            heading1        =   "Loc                     "
            datatype1       =   4
            maxlength1      =   8
            tooltiptext1    =   "Location"
            colname2        =   "WVSRMCTG"
            heading2        =   "Ctg"
            datatype2       =   4
            maxlength2      =   2
            tooltiptext2    =   "Raw Materila Category"
            colname3        =   "WVSRMSCTG"
            heading3        =   "Sub Ctg"
            datatype3       =   4
            maxlength3      =   5
            tooltiptext3    =   "Raw Material Sub Category"
            colname4        =   "WVSRMCD"
            heading4        =   "Rm Code                                 "
            datatype4       =   4
            maxlength4      =   16
            tooltiptext4    =   "Raw Material Code"
            colname5        =   "WVSLOTNO"
            heading5        =   "Lot No.                      "
            datatype5       =   4
            maxlength5      =   12
            tooltiptext5    =   "Lot No."
            colname6        =   "WVSRMSZ"
            heading6        =   "Size            "
            datatype6       =   2
            mask6           =   "##0.0000"
            maxlength6      =   8
            tooltiptext6    =   "Raw Material Size"
            colname7        =   "WVSISSRMQTY"
            heading7        =   "Iss Qty   "
            datatype7       =   1
            mask7           =   "#####0"
            maxlength7      =   6
            tooltiptext7    =   "Issued Raw Material Quantity"
            colname8        =   "WVSISSRMWT"
            heading8        =   "Iss Wt.              "
            datatype8       =   2
            mask8           =   "######0.000"
            maxlength8      =   11
            tooltiptext8    =   "Issued Raw Material Weight"
            colname9        =   "WVSRECRMQTY"
            heading9        =   "Rec Qty "
            datatype9       =   1
            mask9           =   "#####0"
            maxlength9      =   6
            tooltiptext9    =   "Received Raw Material Quantity"
            colname10       =   "WVSRECRMWT"
            heading10       =   "Rec Wt            "
            datatype10      =   2
            mask10          =   "######0.000"
            maxlength10     =   11
            tooltiptext10   =   "Received Raw Material Weight"
         End
         Begin MwfCtl.MWCTL_MED ATXT 
            Height          =   285
            Index           =   63
            Left            =   5970
            TabIndex        =   42
            ToolTipText     =   "Total Raw Material Weight Issued"
            Top             =   2790
            Width           =   1275
            _ExtentX        =   2249
            _ExtentY        =   503
            Alignment       =   2
            Enabled         =   0   'False
            Mask            =   "#######0.000"
            MaxLength       =   12
            DataType        =   2
            IdName          =   "WVSISSTOTWT"
         End
         Begin MwfCtl.MWCTL_MED ATXT 
            Height          =   285
            Index           =   68
            Left            =   1095
            TabIndex        =   46
            ToolTipText     =   "Enter Group By Field"
            Top             =   2790
            Width           =   1365
            _ExtentX        =   2408
            _ExtentY        =   503
            MaxLength       =   8
            DataType        =   4
            IdName          =   "WVSGRP"
         End
         Begin MwfCtl.MWCTL_MED ATXT 
            Height          =   285
            Index           =   67
            Left            =   7230
            TabIndex        =   43
            ToolTipText     =   "Total Raw Material Quantity Received"
            Top             =   2790
            Width           =   855
            _ExtentX        =   1508
            _ExtentY        =   503
            Alignment       =   2
            Enabled         =   0   'False
            Mask            =   "######0"
            MaxLength       =   7
            DataType        =   2
            IdName          =   "WVSRECTOTQTY"
         End
         Begin MwfCtl.MWCTL_MED ATXT 
            Height          =   285
            Index           =   69
            Left            =   8070
            TabIndex        =   47
            ToolTipText     =   "Total Raw Material Weight Received"
            Top             =   2790
            Width           =   1335
            _ExtentX        =   2355
            _ExtentY        =   503
            Alignment       =   2
            Enabled         =   0   'False
            Mask            =   "#######0.000"
            MaxLength       =   12
            DataType        =   2
            IdName          =   "WVSRECTOTWT"
         End
         Begin VB.Label LblMain 
            BackStyle       =   0  'Transparent
            BorderStyle     =   1  'Fixed Single
            Caption         =   "  Voucher Summary"
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
            Index           =   6
            Left            =   -30
            TabIndex        =   138
            Top             =   90
            Width           =   13500
         End
         Begin VB.Label ALBL 
            BackStyle       =   0  'Transparent
            Caption         =   "Group By"
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
            Left            =   90
            TabIndex        =   83
            Top             =   2790
            Width           =   1005
         End
      End
   End
   Begin VB.Frame FraLotNo 
      Height          =   1125
      Left            =   4860
      TabIndex        =   141
      ToolTipText     =   "#0.0000"
      Top             =   8130
      Visible         =   0   'False
      Width           =   4575
      Begin MwfCtl.MWCTL_BTN1 CmdLNGo 
         Height          =   465
         Left            =   3840
         TabIndex        =   144
         ToolTipText     =   "Press This Button To Generate New Lot No."
         Top             =   480
         Width           =   525
         _ExtentX        =   926
         _ExtentY        =   820
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
      Begin MwfCtl.MWCTL_MED ATXT 
         Height          =   285
         Index           =   61
         Left            =   840
         TabIndex        =   142
         ToolTipText     =   "Enter New Lot Number To Be Generated"
         Top             =   480
         Width           =   2475
         _ExtentX        =   4366
         _ExtentY        =   503
         MaxLength       =   16
         DataType        =   4
         IdName          =   "WLOTNO"
      End
      Begin VB.Label ALBL 
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
         Height          =   285
         Index           =   33
         Left            =   60
         TabIndex        =   145
         Top             =   480
         Width           =   945
      End
      Begin VB.Label LblMain 
         BackStyle       =   0  'Transparent
         BorderStyle     =   1  'Fixed Single
         Caption         =   "  Generate Lot No."
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
         Index           =   8
         Left            =   -30
         TabIndex        =   143
         Top             =   90
         Width           =   3615
      End
   End
   Begin VB.Frame FraCD 
      Height          =   1035
      Left            =   7050
      TabIndex        =   240
      ToolTipText     =   "#0.0000"
      Top             =   8200
      Visible         =   0   'False
      Width           =   3255
      Begin MwfCtl.MWCTL_MED ATXT 
         Height          =   285
         Index           =   212
         Left            =   1350
         TabIndex        =   241
         ToolTipText     =   "Enter New Voucher Date"
         Top             =   465
         Width           =   855
         _ExtentX        =   1508
         _ExtentY        =   503
         MaxLength       =   8
         DataType        =   3
         IdName          =   "WTDT"
      End
      Begin MwfCtl.MWCTL_BTN1 CmdCDGo 
         Height          =   495
         Left            =   2520
         TabIndex        =   242
         ToolTipText     =   "Press This Button To Change Date"
         Top             =   360
         Width           =   555
         _ExtentX        =   979
         _ExtentY        =   873
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
      Begin VB.Label ALBL 
         BackStyle       =   0  'Transparent
         Caption         =   "New Date"
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
         Index           =   169
         Left            =   150
         TabIndex        =   244
         Top             =   465
         Width           =   1245
      End
      Begin VB.Label LblMain 
         BackStyle       =   0  'Transparent
         BorderStyle     =   1  'Fixed Single
         Caption         =   "  Change Date"
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
         Index           =   12
         Left            =   -30
         TabIndex        =   243
         Top             =   90
         Width           =   3285
      End
   End
   Begin VB.Frame FraShowTot 
      Height          =   2115
      Left            =   0
      TabIndex        =   268
      Top             =   7080
      Width           =   14130
      Begin VB.Frame FraPurFot 
         Height          =   1125
         Left            =   0
         TabIndex        =   284
         Top             =   840
         Width           =   13050
         Begin MwfCtl.MWCTL_MED ATXT 
            Height          =   285
            Index           =   18
            Left            =   8310
            TabIndex        =   289
            ToolTipText     =   "Total Amount"
            Top             =   675
            Width           =   1695
            _ExtentX        =   2990
            _ExtentY        =   503
            Alignment       =   2
            Mask            =   "#########0.00"
            MaxLength       =   13
            DataType        =   2
            DataField       =   "TTotAmt"
            ReCalcOn        =   "TNETAMT,TADDAMT1,TADDAMT2,TADDAMT3,TADDAMT4"
            IdName          =   "TTOTAMT"
         End
         Begin MwfCtl.MWCTL_MED ATXT 
            Height          =   285
            Index           =   19
            Left            =   8310
            TabIndex        =   288
            ToolTipText     =   "Additional Amount 4"
            Top             =   390
            Width           =   1005
            _ExtentX        =   1773
            _ExtentY        =   503
            Alignment       =   2
            Mask            =   "-#####0.00"
            MaxLength       =   10
            DataType        =   2
            DataField       =   "TAddAmt4"
            ReCalcParent    =   "TTOTAMT"
            IdName          =   "TADDAMT4"
         End
         Begin MwfCtl.MWCTL_MED ATXT 
            Height          =   285
            Index           =   20
            Left            =   5160
            TabIndex        =   287
            ToolTipText     =   "Additional Amount 3"
            Top             =   675
            Width           =   1005
            _ExtentX        =   1773
            _ExtentY        =   503
            Alignment       =   2
            Mask            =   "-#####0.00"
            MaxLength       =   10
            DataType        =   2
            DataField       =   "TAddAmt3"
            ReCalcParent    =   "TTOTAMT"
            IdName          =   "TADDAMT3"
         End
         Begin MwfCtl.MWCTL_MED ATXT 
            Height          =   285
            Index           =   21
            Left            =   5160
            TabIndex        =   286
            ToolTipText     =   "Additional Amount 2"
            Top             =   390
            Width           =   1005
            _ExtentX        =   1773
            _ExtentY        =   503
            Alignment       =   2
            Mask            =   "#####0.00"
            MaxLength       =   9
            DataType        =   2
            DataField       =   "TAddAmt2"
            ReCalcParent    =   "TTOTAMT"
            IdName          =   "TADDAMT2"
         End
         Begin MwfCtl.MWCTL_MED ATXT 
            Height          =   285
            Index           =   23
            Left            =   1830
            TabIndex        =   285
            ToolTipText     =   "Additional Amount 1"
            Top             =   675
            Width           =   1005
            _ExtentX        =   1773
            _ExtentY        =   503
            Alignment       =   2
            Mask            =   "#####0.00"
            MaxLength       =   9
            DataType        =   2
            DataField       =   "TAddAmt1"
            ReCalcParent    =   "TTOTAMT"
            IdName          =   "TADDAMT1"
         End
         Begin MwfCtl.MWCTL_MED ATXT 
            Height          =   285
            Index           =   24
            Left            =   1830
            TabIndex        =   290
            ToolTipText     =   "Net Amount"
            Top             =   390
            Width           =   1575
            _ExtentX        =   2778
            _ExtentY        =   503
            Alignment       =   2
            Mask            =   "#########0.00"
            MaxLength       =   13
            DataType        =   2
            DataField       =   "TNetAmt"
            ReCalcParent    =   "TTOTAMT"
            IdName          =   "TNETAMT"
         End
         Begin VB.Label ALBL 
            BackStyle       =   0  'Transparent
            Caption         =   "Tot Amount"
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
            Index           =   18
            Left            =   6630
            TabIndex        =   297
            Top             =   675
            Width           =   1635
         End
         Begin VB.Label LblTAddAmt4 
            BackStyle       =   0  'Transparent
            Caption         =   "Addn Amount 4"
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
            Left            =   6630
            TabIndex        =   296
            Top             =   390
            Width           =   1635
         End
         Begin VB.Label LblTAddAmt3 
            BackStyle       =   0  'Transparent
            Caption         =   "Addn Amount 3"
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
            Left            =   3510
            TabIndex        =   295
            Top             =   675
            Width           =   1635
         End
         Begin VB.Label LblTAddAmt2 
            BackStyle       =   0  'Transparent
            Caption         =   "Addn Amount 2"
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
            Left            =   3510
            TabIndex        =   294
            Top             =   390
            Width           =   1635
         End
         Begin VB.Label LblTAddAmt1 
            BackStyle       =   0  'Transparent
            Caption         =   "Addn Amount 1"
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
            TabIndex        =   293
            Top             =   675
            Width           =   1635
         End
         Begin VB.Label ALBL 
            BackStyle       =   0  'Transparent
            Caption         =   "Net Amount"
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
            Index           =   12
            Left            =   120
            TabIndex        =   292
            Top             =   390
            Width           =   1635
         End
         Begin VB.Label LblMain 
            BackStyle       =   0  'Transparent
            BorderStyle     =   1  'Fixed Single
            Caption         =   "  Purchase Footer"
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
            Index           =   3
            Left            =   -30
            TabIndex        =   291
            Top             =   90
            Width           =   13020
         End
      End
      Begin MwfCtl.MWCTL_MED ATXT 
         Height          =   285
         Index           =   28
         Left            =   990
         TabIndex        =   277
         ToolTipText     =   "Enter Voucher Character"
         Top             =   435
         Width           =   330
         _ExtentX        =   582
         _ExtentY        =   503
         MaxLength       =   3
         DataType        =   4
         IdName          =   "WCTG"
      End
      Begin MwfCtl.MWCTL_MED ATXT 
         Height          =   285
         Index           =   29
         Left            =   2610
         TabIndex        =   278
         Top             =   435
         Width           =   1215
         _ExtentX        =   2143
         _ExtentY        =   503
         Alignment       =   2
         Mask            =   "######0.000"
         MaxLength       =   11
         DataType        =   2
         IdName          =   "WTOTNETWT"
      End
      Begin MwfCtl.MWCTL_MED ATXT 
         Height          =   285
         Index           =   31
         Left            =   10890
         TabIndex        =   279
         Top             =   435
         Width           =   1215
         _ExtentX        =   2143
         _ExtentY        =   503
         Alignment       =   2
         Mask            =   "######0"
         MaxLength       =   7
         DataType        =   2
         IdName          =   "WTOTQTY"
      End
      Begin MwfCtl.MWCTL_MED ATXT 
         Height          =   285
         Index           =   32
         Left            =   12840
         TabIndex        =   280
         Top             =   435
         Width           =   1215
         _ExtentX        =   2143
         _ExtentY        =   503
         Alignment       =   2
         Mask            =   "######0.000"
         MaxLength       =   11
         DataType        =   2
         IdName          =   "WTOTWT"
      End
      Begin MwfCtl.MWCTL_MED ATXT 
         Height          =   285
         Index           =   34
         Left            =   4620
         TabIndex        =   281
         ToolTipText     =   "Net Weight In Grams For The Rm Code Of The Current Row "
         Top             =   435
         Width           =   1215
         _ExtentX        =   2143
         _ExtentY        =   503
         Alignment       =   2
         Mask            =   "######0.000"
         MaxLength       =   11
         DataType        =   2
         IdName          =   "WNETWT"
      End
      Begin MwfCtl.MWCTL_MED ATXT 
         Height          =   285
         Index           =   70
         Left            =   6525
         TabIndex        =   282
         ToolTipText     =   "Total Stk Rate"
         Top             =   435
         Visible         =   0   'False
         Width           =   1425
         _ExtentX        =   2514
         _ExtentY        =   503
         Alignment       =   2
         Mask            =   "-########0.00"
         MaxLength       =   13
         DataType        =   2
         IdName          =   "WTOTSTKRT"
      End
      Begin MwfCtl.MWCTL_MED ATXT 
         Height          =   285
         Index           =   71
         Left            =   8700
         TabIndex        =   283
         ToolTipText     =   "Total Stock Amount"
         Top             =   435
         Visible         =   0   'False
         Width           =   1455
         _ExtentX        =   2566
         _ExtentY        =   503
         Alignment       =   2
         Mask            =   "#########0.00"
         MaxLength       =   13
         DataType        =   2
         IdName          =   "WTOTSTKAMT"
      End
      Begin MwfCtl.MWCTL_BTN1 CmdShowTotOk 
         Height          =   465
         Left            =   13250
         TabIndex        =   298
         Top             =   1125
         Width           =   615
         _ExtentX        =   1085
         _ExtentY        =   820
         BackColor       =   -2147483633
         MaskColor       =   8454143
         ForeColor       =   11842740
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
      Begin VB.Label LblTotWt 
         BackStyle       =   0  'Transparent
         Caption         =   "Tot Wt"
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
         Left            =   12120
         TabIndex        =   276
         Top             =   435
         Width           =   615
      End
      Begin VB.Label LblTotQty 
         BackStyle       =   0  'Transparent
         Caption         =   "Tot Qty"
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
         Left            =   10200
         TabIndex        =   275
         Top             =   435
         Width           =   705
      End
      Begin VB.Label LblTotNetWt 
         BackStyle       =   0  'Transparent
         Caption         =   "Pure Net Wt"
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
         Left            =   1395
         TabIndex        =   274
         Top             =   435
         Width           =   1155
      End
      Begin VB.Label LblCtg 
         BackStyle       =   0  'Transparent
         Caption         =   "Category"
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
         TabIndex        =   273
         Top             =   435
         Width           =   855
      End
      Begin VB.Label LblNetWt 
         BackStyle       =   0  'Transparent
         Caption         =   "Net Wt"
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
         Left            =   3915
         TabIndex        =   272
         Top             =   435
         Width           =   705
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
         Index           =   14
         Left            =   -30
         TabIndex        =   271
         Top             =   90
         Width           =   14160
      End
      Begin VB.Label LblTotStkAmt 
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
         Height          =   285
         Left            =   8055
         TabIndex        =   270
         Top             =   435
         Width           =   585
      End
      Begin VB.Label LblTotStkRt 
         BackStyle       =   0  'Transparent
         Caption         =   "Stk Rt"
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
         Left            =   5895
         TabIndex        =   269
         Top             =   435
         Width           =   585
      End
   End
End
Attribute VB_Name = "EmrFrmTxnRm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'****** Sachin - Completed Form With IdNos and Partition Key and QC on 17-01-08 ******
'---- General Comments on the Entry --------'
' 6 Menus. All the menus use the same form EmrFrmTxnRm.
'   1. 'PR' - Purchase Entry
'      This Entry is used to enter the RM Purchase Transaction.
'   2. 'SA' - Sales Entry
'      This Entry is used to enter the RM sales Transaction.
'   3. 'TZ' - Custom Purchase
'      This Entry is used to enter the RM Custom Purchase Transaction.
'   4. 'MV' - RM Issue
'      This Entry is used to Issue Raw Materials to the various Stock Locations
'   5. 'REC' - Recovery Transaction
'      This Entry is used to enter details of the Raw Materials recovered from loss locations
'         i.e. for example out of the total 10 gms of 14Kt gold in the Factory Loss location,
'         4 is recovered as 24KT gold and put in a factory stock location and the rest
'         is entered as loss in 24Kt gold in a loss location outside the factory.
'      The Issued and received weights must match.
'      The RmToLoc is always a location of LocTyp 'Xc'.
'      The recovery footer frame shows the rm wt in the factory loss loc and the rm wt
'         recovered. It also shows the recovery percentage.
'   6. 'CNV'- Conversion Entry
'      This entry is used normally for converting 24Kt gold to golds of other karats.
'      The issued and received wts must match
'
' X-Txn, A-Txnd (TdSrNo= 0)
'
'CmdVS - To show the Voucher Summary
'        It shows the Location wise Issue & Receipt Wt for each RmCtg, SubCtg or RmCode
'           for the entire Voucher
'        Private sub VouchSumm
'CmdCP - Copy option is used to mainly copy the structure from another voucher (in 'REC'
'        and 'CNV'). The main use is, it helps the user to make the entry easily.
'        Private sub CpyRm
'--------------------

'*** Declare the variables ms_LockDt, mf_DtLocked for the Lock Date Check
Dim ms_LockDt As String
Dim ms_FutureDt As String
Dim ms_PurgeDt As String    '****** Sachin 3.02 24-12-07
Dim mf_DtLocked As Boolean
Dim mf_ValidLocSeq As Boolean
Dim mf_AvgRtFldChg As Boolean
Dim ms_ChkPurStkVal As String
Dim mf_OLORYN As Boolean
Option Explicit
Dim moCn As MwfLib.MDOConnection   '*** (Bef speed) bef 24/11/06
Dim ms_TcTyp As String, ms_TCoCd As String
Dim ms_TTc As String, ms_TYy As String, ms_TChr As String
Private Enum en_TxnRmFra
  Summ = 0
  Cpy = 1
  vch = 2
  CpISS = 3
  LotNo = 4
  CpyRet = 5      '*** (Jen 2.13)
  Contra = 6      '****** Sachin 2.14.0
  CpyDT = 7       '*** (Jen 2.14 Next)
  CpyLocRm = 8    ' ***** Manali 3.6.0 - 23/10/09 - Copy Rm From Loc
  ImpExc = 9    ' ***** Manali 3.9.0 - Import from Excel
  DET = 10
  ChgDt = 11    '4.3.0.0
  ShowTot = 12    '578-33 Show totals
  InvHd = 13
End Enum
'*** Jay 2.14 *** (SEO Change)
Dim ms_PwdFlag As Boolean
'Dim ms_Pwd As String
'*** Jay 2.14 *** (SEO Change)

'****** Sachin 2.14.0 - [24-11-2006] - [Multi-Company]
Dim mb_MultiCoCd As Boolean, ms_FrmCaption As String
'****** Sachin 2.14.0 - [24-11-2006] - [Multi-Company]

' ***** Manali 3.03 -20/06/08 - Pd Mod - 'L'
Dim mb_PdMod As Boolean
' ***** Manali 3.03 -20/06/08 - Pd Mod - 'L'

' Zubin 212
Dim mArr_Cv As Variant, mArr_Ed As Variant
'Dim wb_GenLNFlg As Boolean
Dim wCpGldAs As String, wTcTyp As String, wLooseTc As String
Dim mb_BestUse As Boolean       '4.1.2
Dim wStkNoDisp As String
Dim mb_SeoUsr As Boolean
Dim ms_MVAckIsReqdYN As String
Private Sub ADC_Load()
  Set m_ObjCmdGetRmRt = ADC.Connection.OpenQuery("ForCpy_GetRmRt")
  Set m_ObjCmdGetLabRt = ADC.Connection.OpenQuery("ForCpy_GetLabRt")
  '*** Set the First Key Control property to Transaction Character (TChr)
  '*** Set the First Non Key Control property to Transaction Date (TDt)
  '*** Set the Child Property Of ADC to GrdTxndA
  '*** Set the Previous and Next Control Properties for the Grid GrdTxndA
  '*** Set the Hot Keys for the Grid GrdTxndA as 'A'
  '*** Store the type of transaction in the variable ms_TcTyp
  '*** Get the initial values for TTc and TYy
  '*** Call the Private Sub Procedure PropSet to set the enable and visible properties
  '    for all the controls on the form based on the TcStructure file
  '*** When the Purchase & Recovery Header frames are not visible then shift the 'A' grid
  '    a little to the top and also increase the height
    
  '****** (Jen 2.13) ******
  ' Zubin 212
  'ReDim mArr_Cv(7), mArr_Ed(7)   '*** (Bef 2.14 Next)
  'ReDim mArr_Cv(8), mArr_Ed(8)    '*** (Jen 2.14 Next)
  'ReDim mArr_Cv(9), mArr_Ed(9)      ' ***** Manali 3.6.0 - 23/10/09 - Copy Rm From Loc
  ReDim mArr_Cv(13), mArr_Ed(13)      ' ***** Manali 3.9.0 - Import from Excel

  Set moCn = ADC.Connection    '*** (Bef speed) bef 24/11/06
  
  '****** Sachin 2.14.0 - [Multi-Company] ******
  mb_MultiCoCd = IIF(moCn.GetFldVal("Select hImpComLoc From Head Where HCd='" + ctSelfCmCd + "' And HCoCd='" + ctSelfCoCd + "'") = "Y", True, False)
  'Set ADC.FirKeyCtl = IIF(mb_MultiCoCd = True, ADC("TCoCd"), ADC("TYy"))     '*** (Bef 3.2.0)
  '****** Sachin 2.14.0 - [Multi-Company]
  
  '*** (Jen 3.2.0)
  Set ADC.FirKeyCtl = IIF(mb_MultiCoCd = True And Not (UCase(ADC.MenuCd) = UCase("PR") Or UCase(ADC.MenuCd) = UCase("SA")), ADC("TCoCd"), ADC("TYy"))
  '*** (Jen 3.2.0)
  
  Set ADC.FirNKeyCtl = ADC("TDt")
  ADC.Child = "GrdTxndA"
  Set GrdTxndA.PrevCtl = ADC("TDesc")
  Set GrdTxndA.NextCtl = ADC
  Set GrdVS.NextCtl = ADC("wVSGrp")
  GrdTxndA.HotKey = "A"
  FraNKeyAll.ZOrder (0)
  
  ms_TCoCd = gs_CoCd
  If UCase(ADC.MenuCd) <> UCase("SA") Then
    'MMAck-if menucd is MA then assign MV as TC type else assign menucd as TC type
    ms_TcTyp = IIF((UCase(ADC.MenuCd) <> UCase("MA")), Trim(UCase(ADC.MenuCd)), "MV")
    ms_TTc = Trim(UCase(moCn.GetFldVal("Select PValue from Param where PTyp='TCTYP' and " + _
              "PMCd='" + ms_TcTyp + "' ")))
  Else
    ms_TcTyp = "PR"
    ms_TTc = "SA"
  End If

  ms_TYy = moCn.GetFldVal("Select hYy from Head where HCoCd= '" + ms_TCoCd + "' and HCd='" + ctSelfCmCd + "'")
  
  '****** Sachin 2.14.0 - [24-11-2006] - [Multi-Company] ******
  ms_FrmCaption = Me.Caption
  '*** Bef Jay 3.1.1
  'If mb_MultiCoCd = True And ms_TTc = "MV" Then
  '  CmdContra.Visible = mb_MultiCoCd
  'Else
  '  CmdContra.Visible = False
  'End If
  '*** Bef Jay 3.1.1
  CmdContra.Visible = False     '*** Jay 3.1.1
  '****** Sachin 2.14.0 - [24-11-2006] - [Multi-Company] ******
  
  Call PropSet
  
  'MMAck-Acknowledgement Lable& Textbox will display only if menu type is MV. For other vouchers visible property is false
  'Grid A Acknowledgement option visible for MV & Ack entry, for other vouchers it is invisible
  If UCase(ms_TcTyp) = UCase("MV") Then
    LblAck.Visible = True:    ADC("TACK").Visible = True
    
  Else
    LblAck.Visible = False:    ADC("TACK").Visible = False
    GrdTxndA.ColProp("TdAck").Style = fgcinvisible
  End If
   
  ' Zubin 212
  If moCn.RecSeek("select 'x' from Head where HCd= '" + ctSelfCmCd + "' and " + _
                            "HCoCd= '" + gs_CoCd + "' and HLotNoYN= 'N' ") Then
    LblLotNo.Visible = False: ADC("TLotNo").Visible = False
    GrdTxndA.ColProp("TdLotNo").Style = fgcinvisible: CmdLN.Visible = False
    GrdVS.ColProp("wVSLotNo").Style = fgcinvisible
    LblGVLotNo.Visible = False: ADC("wGVLotNoFr").Visible = False: ADC("wGVLotNoTo").Visible = False
  ElseIf UCase(ADC.MenuCd) <> UCase("PR") And UCase(ADC.MenuCd) <> UCase("MV") And UCase(ADC.MenuCd) <> UCase("MA") Then
    LblLotNo.Visible = False: ADC("TLotNo").Visible = False: CmdLN.Visible = False
  End If
  ' Zubin 212
  
  If FraPurHd.Visible = False And FraRecHd.Visible = False Then
    ' Zubin 213 (Before 300506)
    'Call GrdTxndA.Move(0, 1040, GrdTxndA.Width,  GrdTxndA.Height + 660)
    Call GrdTxndA.Move(0, 900, GrdTxndA.Width, GrdTxndA.Height + 1180)
  End If
  
 '*******************Geeta***********Emr206*********23/04/04
 If moCn.GetFldVal("Select hAvgSRtYn from Head where HCoCd= '" + ms_TCoCd + "' and " + _
                      "HCd='" + ctSelfCmCd + "'") = "Y" And UCase(ADC.MenuCd) = UCase("CNV") Then
     CmdGV.Visible = True
 Else
     CmdGV.Visible = False
 End If

 CmdChgDt.Visible = IIF(UCase(ADC.MenuCd) = UCase("PR"), True, False)    '4.3.0.0

 '**************Geeta****************
 '*************Geeta*********Emr207***************
  ' Zubin 213 ("SA" added)
  If UCase(ADC.MenuCd) = UCase("PR") Or UCase(ADC.MenuCd) = UCase("MV") Or UCase(ADC.MenuCd) = UCase("SA") Or UCase(ADC.MenuCd) = UCase("CNV") Then
    '****** Sachin 2.14.0 - 20/11/06 - [Multi-Company] ******
    ADC("TCoCd").Enabled = False
    '****** Sachin 2.14.0 - 20/11/06 - [Multi-Company] ******
    LblCurrCd.Visible = True
    ADC("TCurrCd").Visible = True
       
' ########################  Manoj #### Ver: 2.0.9 Patch #### Date: 01/12/2004  ########################
'    If gb_SingleCoCd = True Or UCase$(Left$(gs_CoNm, 8)) <> "DIAJEWEL" Then CmdCP.Visible = False
' ###########################################  Manoj  ###########################################
  Else
  '****** Sachin 2.14.0 - 20/11/06 - [Multi-Company] ******
    ADC("TCoCd").Enabled = True
    '****** Sachin 2.14.0 - 20/11/06 - [Multi-Company] ******
    LblCurrCd.Visible = False
    ADC("TCurrCd").Visible = False
  End If
  '*****************************************
  'Manoj 2.10.0
' ########################  Manoj #### Ver: 2.0.10 Patch #### Date: 09/05/2005  ########################
'  If gb_SingleCoCd = True Or UCase$(Left$(gs_CoNm, 8)) <> "DIAJEWEL" Then
  If gb_SingleCoCd = True Or _
    moCn.RecSeek("Select * from Param Where PTyp = 'MWOPT' And PMCD = 'MULCOTRF' And PValue = 'Y'") = False Then
' ###########################################  Manoj  ###########################################
    If UCase$(ADC.MenuCd) = "PR" Then
      CmdCP.Visible = False
    ElseIf UCase$(ADC.MenuCd) = "CNV" Then
      LblCpdChk.Visible = False
      ADC("wCpDetChk").Visible = False
    End If
  End If
  'Manoj 2.10.0
  
  '*** (Jen 2.13)
  If UCase$(ADC.MenuCd) <> "PR" Or moCn.GetFldVal("Select HAutomateReturns from Head where HCoCd= '" + gs_CoCd + "' and HCd= '" + ctSelfCmCd + "'") = "N" Then
    CmdCpyRet.Visible = False: CmdCpyRet.Enabled = False
  End If
  '*** (Jen 2.13)
  
  '*** (Jen 2.14 Next)
  'If UCase$(ADC.MenuCd) <> "PR" Then
  'If ms_TcTyp <> "PR" And ms_TcTyp <> "MV" Then
  If ms_TTc <> "PR" And ms_TTc <> "MV" Then
    CmdCpyDt.Visible = False: CmdCpyDt.Enabled = False
  End If
  
  If ms_TTc = "MV" Then
    ADC("wDtCoCd").Enabled = True
  Else
    ADC("wDtCoCd").Enabled = False
  End If
  '*** (Jen 2.14 Next)
  
  '****** Sachin 3.02 24-12-07 - Setting the Purge Date
    
  '****** Sachin 3.02.0 - Hiding the Partition Key Column based on Partition Selected
  ADC("TPRTKEY").Visible = IIF(gs_Partition = ctCurrPrtn, False, True)
  GrdTxndA.ColProp("TdPrtKey").Style = IIF(gs_Partition = ctCurrPrtn, fgcinvisible, fgcReadOnly)
  
  ' ***** Manali 3.03 -20/06/08 - Pd Mod - 'L'
  mb_PdMod = moCn.RecSeek("Select 'x' from Head  where HCoCd= '" + ctSelfCoCd + "' and HCd= '" + ctSelfCmCd + "' And HPdModYN='Y'")
  ' ***** Manali 3.03 -20/06/08 - Pd Mod - 'L'
 
  ' ***** Manali 3.8.0 - Rate By QW
 ' If UCase(ADC.MenuCd) = UCase("SA") Then GrdTxndA.ColProp("TdRtByQW").Style = fgcinvisible

 '4.1.2
 If UCase(ADC.MenuCd) = UCase("PR") Or UCase(ADC.MenuCd) = UCase("SA") Then
    mb_BestUse = moCn.RecSeek("Select 'x' from Head where HCoCd= '" + gs_CoCd + "' and " + "HCd= '" + ctSelfCmCd + "' and hBestUse= 'Y'")
    If mb_BestUse Then
        FraGrdTxndADets.Visible = True
        FraGrdTxndADets.ZOrder (0)  '4.1.2
    Else
        GrdTxndA.Height = GrdTxndA.Height + 675
        LblTdDesc.Top = LblTdDesc.Top + 675: ADC("wTdDesc").Top = ADC("wTdDesc").Top + 675
        FraGrdTxndADets.Visible = False
    End If
 Else
    GrdTxndA.Height = GrdTxndA.Height + 675
    LblTdDesc.Top = LblTdDesc.Top + 675: ADC("wTdDesc").Top = ADC("wTdDesc").Top + 675
    FraGrdTxndADets.Visible = False
 End If
 
 If GrdTxndA.ColProp("TDRMSZ").Style <> fgcinvisible Then
      Call SetStockSizeVisible(2, GrdTxndA.ColProp("TDRMSZ2"), ADC.Connection)
      Call SetStockSizeVisible(3, GrdTxndA.ColProp("TDRMSZ3"), ADC.Connection)
 Else
    GrdTxndA.ColProp("TDRMSZ2").Style = fgcinvisible
    GrdTxndA.ColProp("TDRMSZ3").Style = fgcinvisible
 End If
 'MMAck-set column property of GridA for Acknowledge entry
 If UCase(ADC.MenuCd) = UCase("MA") Then
    GrdTxndA.ColProp("TDFRRMLOC").Style = fgcReadOnly
    GrdTxndA.ColProp("WTDFRRMIR").Style = fgcReadOnly
    GrdTxndA.ColProp("TDRMCD").Style = fgcReadOnly
    GrdTxndA.ColProp("TDLOTNO").Style = fgcReadOnly
    GrdTxndA.ColProp("TDRMSZ").Style = fgcReadOnly
    GrdTxndA.ColProp("TDRMSTKRT").Style = fgcReadOnly
    GrdTxndA.ColProp("TDRMQTY").Style = fgcReadOnly
    GrdTxndA.ColProp("TDRMWT").Style = fgcReadOnly
    GrdTxndA.ColProp("TDTORMLOC").Style = fgcReadOnly
    GrdTxndA.ColProp("TDPURRT").Style = fgcReadOnly
    GrdTxndA.ColProp("TDPURAMT").Style = fgcReadOnly
    GrdTxndA.ColProp("TDRMSZ2").Style = fgcReadOnly
    GrdTxndA.ColProp("TDRMSZ3").Style = fgcReadOnly
 End If
  
 If moCn.GetFldVal("Select hAvgSRtYn from Head where HCoCd= '" + gs_CoCd + "'") = "Y" Then
 'StkRt-Getting stock rate display Y\N from head table
  wStkNoDisp = moCn.GetFldVal("Select hStkRtNoDsp from Head where hcocd='" + gs_CoCd + "'")
  
  'StkRt-when NoDisplay =NO then stock rate is visible,if NoDisplay =yes &AvgStkRate=yes then user donot have
  'SEO rights for the form then stock rate column is Invisible else visible for Seo users.
  If wStkNoDisp = "N" Then
    GrdTxndA.ColProp("TDRMSTKRT").Style = fgcNormal
    If UCase(ADC.MenuCd) = UCase("CNV") Then
    ADC("WTOTSTKRT").Visible = True: ADC("WTOTSTKAMT").Visible = True
    LblTotStkRt.Visible = True: LblTotStkAmt.Visible = True
    End If
  ElseIf wStkNoDisp = "Y" Then
    mb_SeoUsr = moCn.RecSeek("Select 'x' From UsrAccess where UaUsrCd = '" + gs_UsrCd + "' " + _
                        " and UaMnuCd = '" + ADC.MenuCd + "' and UaSeoYn= 'Y' ")
    If mb_SeoUsr = True Then
      GrdTxndA.ColProp("TDRMSTKRT").Style = fgcNormal
      If UCase(ADC.MenuCd) = UCase("CNV") Then
      ADC("WTOTSTKRT").Visible = True: ADC("WTOTSTKAMT").Visible = True
      LblTotStkRt.Visible = True: LblTotStkAmt.Visible = True
      End If
    End If
    If mb_SeoUsr = False Then
      GrdTxndA.ColProp("TDRMSTKRT").Style = fgcinvisible
      If UCase(ADC.MenuCd) = UCase("CNV") Then
      ADC("WTOTSTKRT").Visible = False: ADC("WTOTSTKAMT").Visible = False
      LblTotStkRt.Visible = False: LblTotStkAmt.Visible = False
      End If
    End If
  End If
 End If
       
 mf_OLORYN = IIF(moCn.GetFldVal("Select hOLORYN from Head where HCoCd= '" + ms_TCoCd + "' and HCd='" + ctSelfCmCd + "'") = "Y", True, False)
 If mf_OLORYN Then
  CmdLN.Visible = False
  CmdCpyRet.Visible = False
  CmdCpLocRm.Visible = False
 End If
 If UCase(ADC.MenuCd) = UCase("MA") Then
    ADC.AllowAdd = False
    ADC.AllowDelete = False
 End If
 'sv.9 getting purge date from vparam
 ms_PurgeDt = GetPurgeDt(ms_TCoCd, "")
 ms_MVAckIsReqdYN = ""
 If UCase(ADC.MenuCd) = UCase("CNV") Then CmdShowTot.Visible = True   '578-33 showtotal command button will be visible only for cnv entry
End Sub
Private Sub ADC_KeyWhen()
  '*** Disable All command buttons
  '*** set the default values for TTc,TYy and TChr from the previous entry
  '*** Hide all frames except Frame FraNKeyAll
  mf_DtLocked = False
  mf_ValidLocSeq = True
  
  Call EnaDisaCmds(True)
  Call HideAllFras
  
  '****** Sachin 2.14.0 - 17-11-2006 [Multi-Company] ******
  If ms_TCoCd <> "" Then
     ADC("TCoCd") = ms_TCoCd
  Else
     ADC("TCoCd") = gs_CoCd
  End If
  
  If mb_MultiCoCd = False Or _
       (UCase(ADC.MenuCd) = UCase("PR") Or UCase(ADC.MenuCd) = UCase("SA")) Then
        ADC("TCoCd").Enabled = False  ': adc("wCpTCoCd").Enabled = False '*** Jay 3.1.1 [ADC("wCpTCoCd").Enabled ]
  Else
        ADC("TCoCd").Enabled = True   ': adc("wCpTCoCd").Enabled = True   '*** Jay 3.1.1 [ADC("wCpTCoCd").Enabled ]
  End If
  
  'If ms_TCoCd <> "" Then ADC("TCoCd") = ms_TCoCd
  '****** Sachin 2.14.0 - 17-11-2006 [Multi-Company] ******
  
  If ms_TTc <> "" Then ADC("TTc") = ms_TTc
  If ms_TYy <> "" Then ADC("TYy") = ms_TYy
  If ms_TChr <> "" Then ADC("TChr") = ms_TChr
  
  '****** Sachin 2.14.0 - [24-11-2006] - [Multi-Company] ******
  Me.Caption = GetFrmCaption(ADC("TCoCd"), ms_FrmCaption)
  '****** Sachin 2.14.0 - [24-11-2006] - [Multi-Company] ******
  

  '********* Urmila sales **********
If UCase(ADC.MenuCd) = UCase("MV") Or UCase(ADC.MenuCd) = UCase("MA") Or UCase(ADC.MenuCd) = UCase("CNV") Then
  ALBL(1).Caption = "Customer Cd"
  ALBL(13).Caption = "Customer Name"
End If

If UCase(ADC.MenuCd) = UCase("SA") Then
  LblMain(1).Caption = "  Sales Header"
  LblMain(3).Caption = "  Sales Footer"
  ALBL(1).Caption = "Customer Cd"
  ALBL(13).Caption = "Customer Name"
  GrdTxndA.ColProp("TdPurRt").Heading = "Sal Rt            "
  GrdTxndA.ColProp("TdPurAmt").Heading = "Sal Amount     "
  ADC("TBILLNO").Visible = False: ADC("TBILLNO").Enabled = False
  ADC("TBILLDT").Visible = False: ADC("TBILLDT").Enabled = False
  ADC("TBOENO").Visible = False: ADC("TBOENO").Enabled = False
  ADC("TBOEDT").Visible = False: ADC("TBOEDT").Enabled = False
  lblSlash1.Visible = False: lblSlash2.Visible = False
  lblBillNo.Visible = False: lblBOENo.Visible = False
End If
If UCase(ADC.MenuCd) = UCase("CNV") Then
  LblMain(1).Caption = "  Conversion Header"
  LblMain(3).Caption = "  Conversion Footer"
  ' After ClrScr
  
  ALBL(1).Caption = "Customer Cd"
  ALBL(13).Caption = "Customer Name"
  GrdTxndA.ColProp("TdPurRt").Heading = "Rate            "
  GrdTxndA.ColProp("TdPurAmt").Heading = "Amount     "
  ADC("TBILLNO").Visible = False: ADC("TBILLNO").Enabled = False
  ADC("TBILLDT").Visible = False: ADC("TBILLDT").Enabled = False
  ADC("TBOENO").Visible = False: ADC("TBOENO").Enabled = False
  ADC("TBOEDT").Visible = False: ADC("TBOEDT").Enabled = False
  lblSlash1.Visible = False: lblSlash2.Visible = False
  lblBillNo.Visible = False: lblBOENo.Visible = False
End If
If UCase(ADC.MenuCd) = UCase("SA") Then
  'Changing the Captions Of the Frames, Labels and Grid Columns to make applicable to Sales
  ' After ClrScr
  'FraPurHd.Caption = "Sales Header"
  'FraPurFot.Caption = "Sales Footer"
  LblMain(1).Caption = "  Sales Header"
  LblMain(3).Caption = "  Sales Footer"
  ' After ClrScr
  
  ALBL(1).Caption = "Customer Cd"
  ALBL(13).Caption = "Customer Name"
  GrdTxndA.ColProp("TdPurRt").Heading = "Sal Rt            "
  GrdTxndA.ColProp("TdPurAmt").Heading = "Sal Amount     "
  ADC("TBILLNO").Visible = False: ADC("TBILLNO").Enabled = False
  ADC("TBILLDT").Visible = False: ADC("TBILLDT").Enabled = False
  ADC("TBOENO").Visible = False: ADC("TBOENO").Enabled = False
  ADC("TBOEDT").Visible = False: ADC("TBOEDT").Enabled = False
  lblSlash1.Visible = False: lblSlash2.Visible = False
  lblBillNo.Visible = False: lblBOENo.Visible = False
End If

'**********************************
' ########################  Manoj #### Ver: 2.0.6 #### Date: 25/12/2003  ########################
  txtPwd.Text = ""
' ###########################################  Manoj  ###########################################
ms_ChkPurStkVal = ""
If UCase(ADC.MenuCd) = UCase("PR") Then _
    ms_ChkPurStkVal = moCn.GetFldVal("Select vPDesc From vParam where vPTyp='CHR' and vPCoCd='" + gs_CoCd + "' and vPMCd='PR' and vPSCd='ZYX'")
End Sub
Private Sub ADC_KeyValid(Cancel As Boolean, ErrMsg As String)
  '*** Make sure the user enters some value in the TYy & TChr fields
  
  If ADC("TYy") = "" Then Cancel = True: ErrMsg = "Enter Voucher Year": Exit Sub
  If ADC("TChr") = "" Then Cancel = True: ErrMsg = "Enter Voucher Character": Exit Sub
  
End Sub

Private Sub ADC_NonKeyWhen(Cancel As Boolean, ErrMsg As String)
   '*** Store the TTc,TYy and TChr values for the next entry
  '*** Call the Private Sub Procedure CmdPropSet to set the enable and visible properties
  '    for all the commandbuttons on the form based on the TcStructure file
  '*** Give default values fro TDt, TFrRmLoc, TToRmLoc, TLsFrDt, TLsToDt, TBillDt,
      'TBOEDt, TUseBefDt, wFrRmIr if they are blank (Empty date gives '01/01/1980')
  '*** If TAddAmt1 is visible then set the Next Control property of the 'A' grid as
      'TAddAmt1
  '*** If TDustWt is visible then set the Next Control property of the 'A' grid as
      'TDustWt
      If ADC.IsNew And (Not moCn.RecSeek("select vPSCd from vParam where " + _
                         "vPCoCd= '" + ADC("TCoCd") + "' and vPTyp= 'CHR' and " + _
                         "vPMCd = '" + ADC("TTc") + "' and vPSCd='" + ADC("TChr") + "' and vPValidYn<>'N'")) Then
       Cancel = True: ErrMsg = "Voucher Character No More Valid": ADC.AllowSave = False: Exit Sub
      Else
        ADC.AllowSave = True
      End If
 
 ' ms_LockDt = moCn.GetFldVal("Select PDesc225 from Param where PTyp= 'LOCKDT' and PMCd= 'LOCKDT'")
 ' ms_FutureDt = moCn.GetFldVal("Select PDesc225 from Param where PTyp= 'LOCKDT' and PMCd= 'FUTUREDT'")
  ms_TCoCd = ADC("TCoCd")
  ms_TTc = ADC("TTc")
  ms_TYy = ADC("TYy")
  ms_TChr = ADC("TChr")
  
  Call GetLockDt(ms_TCoCd, "TC", ms_TTc, ms_LockDt, ms_FutureDt)
  
  'If ms_LockDt = "" Then ms_LockDt = "01/01/80"
  'If ms_FutureDt = "" Then ms_FutureDt = "31/12/28"
 
  '****** Sachin 3.02 24-12-07
  If ADC.AddFndMode = xaddmode And ADC("TPrtKey") <> ctCurrPrtn Then ADC("TPrtKey") = ctCurrPrtn
  
  '*** Parameters for checking the Valid User Locs in the Txn Entries
  gs_ValidPrdLocSeq = Trim(moCn.GetFldVal("Select PDesc225 from param where PTyp= 'USR' and " + _
                                     "PMCd= '" + EmrFrmMDI.TXT_UCD + "'"))
  gs_ValidOthLocSeq = Trim(moCn.GetFldVal("Select PValue1 from param where PTyp= 'USR' and " + _
                                     "PMCd= '" + EmrFrmMDI.TXT_UCD + "'"))
  '*** Parameters for checking the Valid User Locs in the Txn Entries
  
  '****** Sachin 2.14.0 - [17/11/2006] - [Multi-Company] ******
'  ms_TCoCd = adc("TCoCd")
  '****** Sachin 2.14.0 - [17/11/2006] - [Multi-Company] ******
  
'  ms_TTc = adc("TTc")
'  ms_TYy = adc("TYy")
'  ms_TChr = adc("TChr")
  Call EnaDisaCmds(False)
  Call CmdPropSet
  
  '****** Sachin 2.14.0 - [17/11/2006] - [Multi-Company] ******
  ' Zubin 212
  If moCn.RecSeek("select 'x' from Head where HCd= '" + ctSelfCmCd + "' and " + _
                            "HCoCd= '" + ADC("TCoCd") + "' and HLotNoYN= 'N' ") Then
    LblLotNo.Visible = False: ADC("TLotNo").Visible = False
    GrdTxndA.ColProp("TdLotNo").Style = fgcinvisible: CmdLN.Visible = False
    GrdVS.ColProp("wVSLotNo").Style = fgcinvisible
    LblGVLotNo.Visible = False: ADC("wGVLotNoFr").Visible = False: ADC("wGVLotNoTo").Visible = False
  ElseIf UCase(ADC.MenuCd) <> UCase("PR") And UCase(ADC.MenuCd) <> UCase("MV") And UCase(ADC.MenuCd) <> UCase("MA") Then
    LblLotNo.Visible = False: ADC("TLotNo").Visible = False: CmdLN.Visible = False
  ElseIf UCase(ADC.MenuCd) = UCase("MA") Then
    GrdTxndA.ColProp("TdLotNo").Style = fgcReadOnly
    CmdLN.Visible = False
  Else
    LblLotNo.Visible = True: ADC("TLotNo").Visible = True
    GrdTxndA.ColProp("TdLotNo").Style = fgcNormal: CmdLN.Visible = True
    GrdVS.ColProp("wVSLotNo").Style = fgcNormal
    LblGVLotNo.Visible = True: ADC("wGVLotNoFr").Visible = True: ADC("wGVLotNoTo").Visible = True
  End If
  
  'MV
  Dim wAckYn As Integer
  ms_MVAckIsReqdYN = ""
  wAckYn = moCn.GetFldVal("Select vPNum from vParam where vPCoCd= '" + ADC("TCoCd") + "' and vPTyp= 'CHR' and " + _
                             "vPMCd='" + ADC("TTc") + "' and vPSCd='" + ADC("TChr") + "'")
  
  If UCase(ms_TcTyp) = UCase("MV") And wAckYn = 0 Then
    ADC("WFRRMIR").Enabled = True
    ms_MVAckIsReqdYN = "N"
    LblAck.Visible = False:  ADC("TACK").Visible = False
    GrdTxndA.ColProp("TdAck").Style = fgcinvisible
  ElseIf UCase(ms_TcTyp) = UCase("MV") And wAckYn = 1 Then
    LblAck.Visible = True:  ADC("TACK").Visible = True
    ms_MVAckIsReqdYN = "Y"
    ADC("WFRRMIR").Enabled = False
    GrdTxndA.ColProp("TdAck").Style = fgcNormal
  End If
  
  ADC("TAck").Enabled = False
  If UCase(ADC.MenuCd) = UCase("MA") Then
    ADC("TAck").Enabled = True
  End If
  
  
  If moCn.GetFldVal("Select hAvgSRtYn from Head where HCoCd= '" + ADC("TCoCd") + "' and " + _
                       "HCd='" + ctSelfCmCd + "'") = "Y" And UCase(ADC.MenuCd) = UCase("CNV") Then
   CmdGV.Visible = True
  Else
   CmdGV.Visible = False
  End If
  
  
  
  '*** (Jen 2.13)
  If UCase$(ADC.MenuCd) <> "PR" Or moCn.GetFldVal("Select HAutomateReturns from Head where HCoCd= '" + ADC("TCoCd") + "' and HCd= '" + ctSelfCmCd + "'") = "N" Then
    CmdCpyRet.Visible = False: CmdCpyRet.Enabled = False
  Else
    CmdCpyRet.Visible = True: CmdCpyRet.Enabled = True
  End If
  '*** (Jen 2.13)
  
  '*** (Jen 2.14 Next)
  ''If UCase$(ADC.MenuCd) <> "PR" Then
  ''  CmdCpyDt.Visible = False: CmdCpyDt.Enabled = False
  ''Else
  ''  CmdCpyDt.Visible = True: CmdCpyDt.Enabled = True
  ''End If
  '*** (Jen 2.14 Next)
  
  '****** Sachin 2.14.0 - [17/11/2006] - [Multi-Company] ******
  
  ' Zubin 212
  If moCn.RecSeek("select 'x' from Head where HCd='" + ctSelfCmCd + "' and HCoCd='" + ADC("TCoCd") + "' and HLotNoYN='N'") _
              Or (UCase(ADC.MenuCd) <> UCase("PR") And UCase(ADC.MenuCd) <> UCase("MV")) Then
    CmdLN.Visible = False
  Else
    CmdLN.Visible = True
  End If
  ' Zubin 212
  
  If ADC.AddFndMode = xaddmode Then
    Dim ws_Dt As String
    ws_Dt = GetNextDate(ADC("TCoCd"))  '53 company code parameter added
    ws_Dt = IIF((ws_Dt < CDate(ms_LockDt)) _
                               Or _
                (ws_Dt > CDate(ms_FutureDt)), _
                "", ws_Dt)
  
    If ws_Dt <> "" Then ADC("TDt") = ws_Dt
  
    'ADC("TDt") = moCn.SrvrDate
    ADC("TLsFrDt") = MWLib.EmptyDate
    ADC("TLsToDt") = MWLib.EmptyDate
    ADC("TBillDt") = MWLib.EmptyDate
    ADC("TBOEDt") = MWLib.EmptyDate
    ADC("TUseBefDt") = MWLib.EmptyDate
    '***********Geeta****Emr207*************
    'ADC("tUSDYN") = "Y"          Commented as field removed from Db [Sachin]
    '***********Geeta****Emr207*************
    '****** Sachin 2.13.0 - [17-07-2006] Multi-Currency ******
    ADC("TCurrCd") = moCn.GetFldVal("Select HBaseCurCd From Head Where HCd='" + ctSelfCmCd + "' And HCoCd='" + ADC("TCoCd") + "'")
    '****** Sachin 2.13.0 - [17-07-2006] Multi-Currency ******
    ADC("TFrRmLoc") = moCn.GetFldVal("Select vPValue from vParam where " + _
                                     "vPCoCd= '" + ADC("TCoCd") + "' and vPTyp='CHR' and " + _
                                     "vPMCd='" + ADC("TTc") + "' and vPSCd='" + ADC("TChr") + "'")
    ADC("TToRmLoc") = moCn.GetFldVal("Select vPDesc225 from vParam where " + _
                                     "vPCoCd= '" + ADC("TCoCd") + "' and vPTyp='CHR' and " + _
                                     "vPMCd='" + ADC("TTc") + "' and vPSCd='" + ADC("TChr") + "'")
    
  End If
  
  If ADC("wFrRmIr") = "" Then
    Dim wFrRmIr As String
    wFrRmIr = moCn.GetFldVal("Select vPValue2 from vParam where " + _
                             "vPCoCd= '" + ADC("TCoCd") + "' and vPTyp= 'CHR' and " + _
                             "vPMCd='" + ADC("TTc") + "' and vPSCd='" + ADC("TChr") + "'")
    wFrRmIr = IIF(wFrRmIr = "C", "I", "R")
    If wFrRmIr <> "" Then ADC("wFrRmIr") = wFrRmIr
  End If
  
  mf_DtLocked = IIF((ADC("TDt") <> "" And ADC("TDt") < CDate(ms_LockDt)) _
                                Or _
                   (ADC("TDt") <> "" And ADC("TDt") > CDate(ms_FutureDt)) _
                                Or _
                  (ADC("TDt") <= CDate(ms_PurgeDt)), _
                True, False)
  
  '******************Geeta********Emr206******02/04/04****
  '*****************added adc("TLockYn") check*************
  If ADC("TPrtKey") <> ctCurrPrtn Or mf_DtLocked = True Or ADC("TLockYN") = "Y" Then
    GrdTxndA.AllowAdd = False: GrdTxndA.AllowDelete = False
  Else
    If gs_ValidPrdLocSeq <> "" Then
      If ADC("TFrBLoc") <> "" And InStr(1, "," + UCase(gs_ValidPrdLocSeq) + ",", "," + UCase(ADC("TFrBLoc")) + ",") <= 0 Then
        mf_ValidLocSeq = False
        GrdTxndA.AllowAdd = False: GrdTxndA.AllowDelete = False
      Else
        '***
        mf_ValidLocSeq = True
        GrdTxndA.AllowAdd = True: GrdTxndA.AllowDelete = True
        '***
      End If
    Else
      mf_ValidLocSeq = True
      GrdTxndA.AllowAdd = True: GrdTxndA.AllowDelete = True
    End If
    
    If gs_ValidOthLocSeq <> "" And ms_TcTyp <> "MV" And ms_TcTyp <> "MA" Then
      If (ADC("TFrRmLoc") <> "" And InStr(1, "," + UCase(gs_ValidOthLocSeq) + ",", "," + UCase(ADC("TFrRmLoc")) + ",") <= 0) Or _
         (ADC("TToRmLoc") <> "" And InStr(1, "," + UCase(gs_ValidOthLocSeq) + ",", "," + UCase(ADC("TToRmLoc")) + ",") <= 0) Then
        mf_ValidLocSeq = False
        GrdTxndA.AllowAdd = False: GrdTxndA.AllowDelete = False
      Else
        '***
        If mf_ValidLocSeq = True Then GrdTxndA.AllowAdd = True: GrdTxndA.AllowDelete = True
        '***
      End If
    ElseIf gs_ValidOthLocSeq <> "" And UCase$(ADC.MenuCd) = "MV" And ms_MVAckIsReqdYN = "Y" And ADC("TAck") = "N" Then
      If ADC("TFrRmLoc") <> "" And InStr(1, "," + UCase(gs_ValidOthLocSeq) + ",", "," + UCase(ADC("TFrRmLoc")) + ",") <= 0 Then
           mf_ValidLocSeq = False: GrdTxndA.AllowAdd = False: GrdTxndA.AllowDelete = False
           ErrMsg = "No permission to From Location"
      Else
        If mf_ValidLocSeq = True Then GrdTxndA.AllowAdd = True: GrdTxndA.AllowDelete = True
      End If
    ElseIf gs_ValidOthLocSeq <> "" And ms_MVAckIsReqdYN = "N" And UCase$(ADC.MenuCd) = "MV" Then
      If ADC("TToRmLoc") <> "" And InStr(1, "," + UCase(gs_ValidOthLocSeq) + ",", "," + UCase(ADC("TToRmLoc")) + ",") <= 0 Then
        mf_ValidLocSeq = False:         GrdTxndA.AllowAdd = False: GrdTxndA.AllowDelete = False
        ErrMsg = "No permission to To Location"
      Else
        If ADC("TFrRmLoc") <> "" And InStr(1, "," + UCase(gs_ValidOthLocSeq) + ",", "," + UCase(ADC("TFrRmLoc")) + ",") <= 0 Then
           mf_ValidLocSeq = False:         GrdTxndA.AllowAdd = False: GrdTxndA.AllowDelete = False
            ErrMsg = "No permission to From Location"
        Else
          If mf_ValidLocSeq = True Then mf_ValidLocSeq = True: GrdTxndA.AllowAdd = True: GrdTxndA.AllowDelete = True
        End If
      End If
    ElseIf gs_ValidOthLocSeq <> "" And ms_MVAckIsReqdYN = "Y" And ADC("TAck") = "Y" Then
        mf_ValidLocSeq = False:         GrdTxndA.AllowAdd = False: GrdTxndA.AllowDelete = False
    ElseIf gs_ValidOthLocSeq <> "" And UCase$(ADC.MenuCd) = "MA" And ADC("TAck") = "N" Then
      If ADC("TToRmLoc") <> "" And InStr(1, "," + UCase(gs_ValidOthLocSeq) + ",", "," + UCase(ADC("TToRmLoc")) + ",") <= 0 Then
           mf_ValidLocSeq = False: GrdTxndA.AllowAdd = False: GrdTxndA.AllowDelete = False: Cancel = True
           ErrMsg = "No permission to To Location"
      End If
    Else
      If mf_ValidLocSeq = True Then GrdTxndA.AllowAdd = True: GrdTxndA.AllowDelete = True
    End If
  End If
  
  'MMAck-if menu code is Ack. then user can't add or delete row entry
  If UCase(ADC.MenuCd) = UCase("MA") Then
    GrdTxndA.AllowAdd = False
    GrdTxndA.AllowDelete = False
    CmdLN.Visible = False
    CmdCP.Visible = False:      CmdCpyDt.Visible = False
    CmdCpLocRm.Visible = False:   CmdImpExc.Visible = False
  End If
  
  If ADC("TAddAmt1").Visible Then Set GrdTxndA.NextCtl = ADC("TAddAmt1")
  If ADC("TDustWt").Visible Then Set GrdTxndA.NextCtl = ADC("TDustWt")
      
  '**************** urmi NegStk ****************************'
  ADC("TSeoPwd") = ""
  txtPwd.Text = ""
  '*********************************************************

' ########################  Manoj #### Ver: 2.0.6 #### Date: 25/12/2003  ########################
  If ADC("TLockYN") = "" Then ADC("TLockYN") = "N"
  '*** Jay 2.14 *** (SEO Change)
  If moCn.GetFldVal("Select PValue1 from Param where Ptyp = 'SEOPWD' and PmCd = 'SEOPWD'") = "" Then ms_PwdFlag = True
  'ms_Pwd = moCn.GetFldVal("Select PValue1 from Param where Ptyp = 'SEOPWD' and PmCd = 'SEOPWD'")
  '*** Jay 2.14 *** (SEO Change)
' ###########################################  Manoj  ###########################################
'******************Geeta********Emr206******02/04/04****
  '*****************added adc("TLockYn") check*************
  If ADC("TLockYN") = "Y" Then
    ADC.AllowDelete = False
  Else
    If UCase(ADC.MenuCd) = UCase("MA") Then
      ADC.AllowDelete = False
    Else
      ADC.AllowDelete = True
    End If
  End If
' ########################  Manoj #### Ver: 2.0.10 Patch #### Date: 09/05/2005  ########################
  If gb_SingleCoCd = True Or _
    moCn.RecSeek("Select * from Param Where PTyp = 'MWOPT' And PMCD = 'MULCOTRF' And PValue = 'Y'") = False Then
    If UCase$(ADC.MenuCd) = "PR" Then
      CmdCP.Visible = False
    ElseIf UCase$(ADC.MenuCd) = "CNV" Then
      LblCpdChk.Visible = False
      ADC("wCpDetChk").Visible = False
    End If
  End If
' ###########################################  Manoj  ###########################################
  ' Zubin 212
  If GrdTxndA.Row = 0 Then CmdLN.Enabled = False
  ' Zubin 212
  'MMACK-In vParam Chr, if Acknowledgement =N then TAck field in X and TdAck in GridA are invisible.
  'otherwise Ack option in X and GridA are visible

  If mf_OLORYN Then
    CmdLN.Visible = False
  Else
    If ADC.MenuCd <> "MA" Then CmdLN.Visible = True
  End If
  If UCase(ADC.MenuCd) = UCase("CNV") Then
    If ADC("TSuppCd") = "" Then ADC("TSuppCd") = "ZSELF"
  End If
  If Not mf_OLORYN Then
    FraCNVOLOR.Visible = False
  End If
  If ms_TcTyp = "MV" Then
      If ms_MVAckIsReqdYN = "Y" And ADC("TAck") = "Y" Then
        Cancel = True
        GrdTxndA.AllowAdd = False
        GrdTxndA.AllowDelete = False
'        ADC.AllowCancel = True
        ADC.AllowDelete = False
        CmdCP.Visible = False:      CmdCpyDt.Visible = False
        CmdCpLocRm.Visible = False:   CmdImpExc.Visible = False
        MsgBox "Acknowledged Voucher is not  editable"
        ', etInfo
      End If
  End If
End Sub
Private Sub ADC_NonKeyValid(Cancel As Boolean, ErrMsg As String)
  '*** In case of 'REC' and 'CNV' entries, the
  '    Issued Rm Wt for a particular Ctg = Received Rm Wt for that particular Ctg
  '    Issued Rm Qty for a particular Ctg = Received Rm Qty for that particular Ctg
  '*** If Not the do not allow the user to save the entry

  Dim wCount, i As Integer
  ' **** Zubn 211 **** '
  Dim wRs_Cnv As MwfLib.MDORowSet
  Dim wb_CnvChk As Boolean
  Dim wd_EffDt As Date
  ' **** Zubn 211 **** '
  
  ' ****** Sachin 2.12 - 06/12/05 - SJM
    Dim wb_SlvMod As Boolean, wb_AlyCnt As Boolean
  ' ****** Sachin 2.12 - 06/12/05 - SJM
    
    ' **** Zubin 211 **** '
        '** Net of RmWt, RmQty for each RmCtg =0 _
            considering purity, & TdFrRmDc, 'D'+ve **
    '''''    wCount = moCn.GetFldVal("Select count('x') from Txnd join RmMst on RmCd=TdRmCd " + _
                                "where TdCoCd='" + ADC("TCoCd") + "' and TdTc='" + ADC("TTc") + "' and " + _
                                "TdYy='" + ADC("TYy") + "' and TdChr='" + ADC("TChr") + "' and " + _
                                "TdNo=" + CStr(ADC("TNo")) + _
                                "group by RmCtg " + _
                                "having ROUND(ABS(sum((case when TdFrRmDc='C' then -TdRmWt " + _
                                "when TdFrRmDc='D' then TdRmWt else 0 end) * " + _
                                "(case when RmCtg='G' or RmCtg='P' then RmPurityWt else 1 end))), 3) > 0 ")
    '''''    If wCount <> 0 Then Cancel = True: ErrMsg = "Net Pure Rm Weight Should Be Zero For Each Category": Exit Sub
    
    '******Zubin ***2.11
    'Set wRs_Cnv = moCn.OpenRes("Select convert(Decimal(16, 2), HAlyCnvDiff) as HAlyCnvDiff, HCnvDiffEffDt " + _
                            "from Head where HCoCd= '" + ADC("TCoCd") + "' and HCd= '" + ctSelfCmCd + "'")
    '    If Not (wRs_Cnv.EOF And wRs_Cnv.BOF) Then
    '      wb_CnvChk = IIF(wRs_Cnv!HAlyCnvDiff > 0 And wRs_Cnv!HCnvDiffEffDt <= ADC("TDt"), True, False)
    '      wd_EffDt = IIF(wb_CnvChk = True, wRs_Cnv!HCnvDiffEffDt, "31/12/2028")
    '    Else
    '      wb_CnvChk = False
    '    End If
    '******Geeta ***2.11***
    wb_CnvChk = moCn.RecSeek("Select HCoCd from Head  where HCoCd= '" + ADC("TCoCd") + "' and HCd= '" + ctSelfCmCd + "' " + _
                              " and  HCnvDiffEffDt <> '01/01/80' and HCnvDiffEffDt<= '" + Str(ADC("TDt")) + "'")
                              
    ' ****** Sachin 2.12 - 06/12/05 - SJM
    wb_SlvMod = moCn.RecSeek("Select 'x' from Head  where HCoCd= '" + ctSelfCoCd + "' and HCd= '" + ctSelfCmCd + "' And HSlvModYN='Y'")
    ' ****** Sachin 2.14.0 Next - 12/02/07 ******
    wb_AlyCnt = moCn.RecSeek("Select 'x' From Txnd Join RmMst On RmCd=TdRmCd and RmPrtKey=TdPrtKey Where TdCoCd='" + ADC("TCoCd") + "' And TdTc='" + ADC("TTc") + "' And TdYy='" + ADC("TYy") + "' And " + _
                              "TdChr='" + ADC("TChr") + "' And TdNo=" + CStr(ADC("TNo")) + " And TdPrtKey='" + ctCurrPrtn + "' And RmCtg='A'")
    ' ****** Sachin 2.14.0 Next - 12/02/07 ******
    
    ' ****** Sachin 2.12 - 06/12/05 - SJM
    ' ****** Sachin 2.12 - 09/12/05 - 'XC' Net Wt to be zero for all Entries ****** '
    ' ****** Sachin 2.14.0 - If Multi-Company Module is Enabled then The check below will not fire in 'MV'
    '****** Sachin 3.02 24-12-07 - Current Partition
    ' ***** Manali 3.03 -20/06/08 - Pd Mod - 'L' included
    If (ADC("TPrtKey") = ctCurrPrtn) Then
'    And ((mb_MultiCoCd = False And ms_TcTyp <> "CNV") Or (mb_MultiCoCd = True And (ms_TcTyp <> "CNV" And ms_TcTyp <> "MV"))) Then
      wCount = moCn.GetFldVal("Select count('x') from Txnd join RmMst on RmCd=TdRmCd and RmPrtKey=TdPrtKey " + _
                              "Where TdCoCd='" + ADC("TCoCd") + "' and TdTc='" + ADC("TTc") + "' and " + _
                              "TdYy='" + ADC("TYy") + "' and TdChr='" + ADC("TChr") + "' and " + _
                              "TdNo=" + CStr(ADC("TNo")) + " and TdPrtKey='" + ctCurrPrtn + "'" + IIF(wb_CnvChk, " AND  RmCtg<> 'A'  ", " ") + _
                              "And (TdFrRmLoc = 'XC' Or TdToRmLoc = 'XC') group by RmCtg " + _
                              "Having  ROUND(ABS(sum((case when TdFrRmDc='C' then -TdRmWt " + _
                              "when TdFrRmDc='D' then TdRmWt else 0 end) * " + _
                              "(case when RmCtg='G' or RmCtg='P' " + IIF(wb_SlvMod = True, " or RmCtg='S'", "") + _
                              IIF(mb_PdMod = True, " or RmCtg='L'", "") + " then RmPurityWt else 1 end))), 3) > 0 ")  ' ****** Silver 2.12 - 02/12/05 - SJM
      If wCount <> 0 Then Cancel = True: ErrMsg = "Net Pure Rm Weight Should Be Zero For Each Category For 'XC' Location": Exit Sub
          
      wCount = moCn.GetFldVal("Select count('x') from Txnd join RmMst on RmCd=TdRmCd and RmPrtKey=TdPrtKey " + _
                              "where TdCoCd='" + ADC("TCoCd") + "' and TdTc='" + ADC("TTc") + "' and " + _
                              "TdYy='" + ADC("TYy") + "' and TdChr='" + ADC("TChr") + "' and " + _
                              "TdNo=" + CStr(ADC("TNo")) + " and TdPrtKey='" + ctCurrPrtn + "' " + _
                              "And (TdFrRmLoc = 'XC' Or TdToRmLoc = 'XC') group by RmCtg " + _
                              "having sum(case when TdFrRmDc='C' then -TdRmQty " + _
                              "when TdFrRmDc='D' then TdRmQty else 0 end) <> 0 ")
      If wCount <> 0 Then Cancel = True: ErrMsg = "Net Rm Quantity Should Be Zero For Each Category": Exit Sub
    End If
    ' ****** Sachin 2.12 - 09/12/05 - 'XC' Net Wt to be zero for all Entries ****** '

    '******Geeta ***2.11***
    '****** Sachin 3.02 24-12-07 - Current Partition
  If ADC("TPrtKey") = ctCurrPrtn And (ms_TcTyp = "REC" Or ms_TcTyp = "CNV") Then
        ' ***** Manali 3.03 -20/06/08 - Pd Mod - 'L' included
         wCount = moCn.GetFldVal("Select count('x') from Txnd join RmMst on RmCd=TdRmCd and RmPrtKey=TdPrtKey " + _
                              "Where TdCoCd='" + ADC("TCoCd") + "' and TdTc='" + ADC("TTc") + "' and " + _
                              "TdYy='" + ADC("TYy") + "' and TdChr='" + ADC("TChr") + "' and " + _
                              "TdNo=" + CStr(ADC("TNo")) + " and TdPrtKey='" + ctCurrPrtn + "' " + IIF(wb_CnvChk, " AND  RmCtg<> 'A'  ", " ") + _
                              "Group by RmCtg " + _
                              "Having  ROUND(ABS(sum((case when TdFrRmDc='C' then -TdRmWt " + _
                              "When TdFrRmDc='D' then TdRmWt else 0 end) * " + _
                              "(Case when RmCtg='G' or RmCtg='P' " + IIF(wb_SlvMod = True, " or RmCtg='S'", "") + _
                              IIF(mb_PdMod = True, " or RmCtg='L'", "") + " then RmPurityWt else 1 end))), 3) > 0 ")  ' ****** Silver 2.12 - 02/12/05 - SJM
      If wCount <> 0 Then Cancel = True: ErrMsg = "Net Pure Rm Weight Should Be Zero For Each Category " + IIF(wb_CnvChk, "Other Than Alloy", ""): Exit Sub
    
      ' ****** Sachin 2.14.0 Next - 12/02/07 - wb_AlyCnt Condition added below ******
      If wb_CnvChk = True And wb_AlyCnt = True Then
        wCount = moCn.GetFldVal(" Select Case When Convert(Decimal(16,3), abs(Sum((Case TdFrRmDc when 'D' then TdRmWt " + _
                              " Else 0 End))) * (select Convert(Decimal(16,2), HAlyCnvDiff) from Head where " + _
                              " HCoCd= '" + ADC("TCoCd") + "' and HCd= '" + ctSelfCmCd + "') / 100 )< " + _
                              "Abs(Convert(Decimal(16,3), Sum((Case TdFrRmDc when 'C' then -TdRmWt " + _
                              "When 'D' then TdRmWt else 0 End)))) then 1 else 0 end " + _
                              "From Txnd " + _
                              "join RmMst on RmCd=TdRmCd and RmPrtKey=TdPrtKey " + _
                              "where TdCoCd='" + ADC("TCoCd") + "' and TdTc='" + ADC("TTc") + "' and " + _
                              "TdYy='" + ADC("TYy") + "' and TdChr='" + ADC("TChr") + "' and " + _
                              "TdNo=" + CStr(ADC("TNo")) + " and TdPrtKey='" + ctCurrPrtn + "' ")
        If wCount <> 0 Then Cancel = True: ErrMsg = "Net Rm Weight Difference Should Be Zero Or Within Permissible Range. (Check Permissible Weight % In Head Master) ": Exit Sub
      End If
  ' **** Zubin 211 **** '
  
      wCount = moCn.GetFldVal("Select count('x') from Txnd join RmMst on RmCd=TdRmCd and RmPrtKey=TdPrtKey " + _
                              "Where TdCoCd='" + ADC("TCoCd") + "' and TdTc='" + ADC("TTc") + "' and " + _
                              "TdYy='" + ADC("TYy") + "' and TdChr='" + ADC("TChr") + "' and " + _
                              "TdNo=" + CStr(ADC("TNo")) + " and TdPrtKey='" + ctCurrPrtn + "'  " + _
                              "Group by RmCtg " + _
                              "Having sum(case when TdFrRmDc='C' then -TdRmQty " + _
                              "When TdFrRmDc='D' then TdRmQty else 0 end) <> 0 ")
      If wCount <> 0 Then Cancel = True: ErrMsg = "Net Rm Quantity Should Be Zero For Each Category": Exit Sub
    End If
    If (Not ms_PwdFlag) And moCn.RecSeek("Select 'Y' From Param where PTyp='CNVCDCHK' and PMcd='CNVCDCHK' and (PDesc225 ='' or charindex(','+'" + ADC("TChr") + "'+',',','+PDesc225+',')>0)") Then
      wCount = moCn.GetFldVal("Select count('x') from Txnd join RmMst on RmCd=TdRmCd and RmPrtKey=TdPrtKey " + _
                                "Where TdCoCd='" + ADC("TCoCd") + "' and TdTc='" + ADC("TTc") + "' and " + _
                                "TdYy='" + ADC("TYy") + "' and TdChr='" + ADC("TChr") + "' and " + _
                                "TdNo=" + CStr(ADC("TNo")) + " and TdPrtKey='" + ctCurrPrtn + "'" + " AND  RmCtg= 'D' " + _
                                "And (TdFrRmLoc = 'XC' Or TdToRmLoc = 'XC') group by RmCd " + _
                                "Having  ROUND(ABS(sum((case when TdFrRmDc='C' then -TdRmWt " + _
                                "when TdFrRmDc='D' then TdRmWt else 0 end) * " + _
                                "(case when RmCtg='G' or RmCtg='P' " + IIF(wb_SlvMod = True, " or RmCtg='S'", "") + _
                                IIF(mb_PdMod = True, " or RmCtg='L'", "") + " then RmPurityWt else 1 end))), 3) > 0 ")  ' ****** Silver 2.12 - 02/12/05 - SJM
      If wCount <> 0 Then Cancel = True: ErrMsg = "Net Rm Weight Should Be Zero For Each RmCd For 'XC' Location": Exit Sub
    End If
    
    If (Not ms_PwdFlag) And moCn.RecSeek("Select 'Y' From Param where PTyp='CNVCDCHK' and PMcd='CNVCDCHK' and (PDesc225 ='' or charindex(','+'" + ADC("TChr") + "'+',',','+PDesc225+',')>0)") Then
      wCount = moCn.GetFldVal("Select count('x') from Txnd join RmMst on RmCd=TdRmCd and RmPrtKey=TdPrtKey " + _
                                "where TdCoCd='" + ADC("TCoCd") + "' and TdTc='" + ADC("TTc") + "' and " + _
                                "TdYy='" + ADC("TYy") + "' and TdChr='" + ADC("TChr") + "' and " + _
                                "TdNo=" + CStr(ADC("TNo")) + " and TdPrtKey='" + ctCurrPrtn + "' " + _
                                "And (TdFrRmLoc = 'XC' Or TdToRmLoc = 'XC') and TdRmCtg = 'D' group by TdRmCd " + _
                                "having sum(case when TdFrRmDc='C' then -TdRmQty " + _
                                "when TdFrRmDc='D' then TdRmQty else 0 end) <> 0 ")
      If wCount <> 0 Then Cancel = True: ErrMsg = "Net Rm Quantity Should Be Zero For Each RmCd": Exit Sub
    End If

    '*** Jay 2.14 ***
    If ADC("wTotStkAmt") < -1 Or ADC("wTotStkAmt") > 1 Then
      If MsgBox("Total Amount Is Not In Acceptable Limit,Do You Want To Save The Record.", _
         vbYesNo + vbInformation + vbDefaultButton2, gs_CoNm) = vbNo Then Cancel = True
    End If
    '*** Jay 2.14 ***
    
   If UCase(ADC.MenuCd) = UCase("PR") And (UCase(ms_ChkPurStkVal) = UCase("VCH") Or UCase(ms_ChkPurStkVal) = UCase("SR")) Then
   Dim wf_Cancel As Boolean, wf_errmsg As String
    TxtPwd_Validate (wf_Cancel)
   If ms_PwdFlag = False Then
       Dim wTdSr As Integer
'       wTdSr = moCn.GetFldVal("Select Min(TdSr) From Txnd Where TdCoCd='" + ADC("TCoCd") + "' And TdTc='" + ADC("TTc") + "' And TdYy='" + ADC("TYy") + "' And " + _
'                        "TdChr='" + ADC("TChr") + "' And TdNo=" + CStr(ADC("TNo")) + " And TdPrtKey='" + ctCurrPrtn + "'" + _
'                        " Group By TdCoCd,TdYy,TdChr,TdNo" + IIF(UCase(ms_ChkPurStkVal) = UCase("SR"), ",TdSr", "") + _
'                        " Having Cast(Sum(Cast(TdRmStkRt*TdRmWt as Decimal(10,2))) as Decimal(10,2))<>0 and  " + _
'                                "Cast(Sum(Cast(TdRmStkRt*TdRmWt as Decimal(10,2)))" + IIF(ADC("TCnvRt") <> 0 And ADC("TMulDiv") = "M", "*", "/") + Str(ADC("TCnvRt")) + "  as Decimal(10,0))<>Cast(Sum(TdPurAmt) as Decimal(10,0)) ")
'
                                
     If UCase(ms_ChkPurStkVal) = UCase("SR") Then
       wTdSr = moCn.GetFldVal("Select Min(TdSr) From Txnd Where TdCoCd='" + ADC("TCoCd") + "' And TdTc='" + ADC("TTc") + "' And TdYy='" + ADC("TYy") + "' And " + _
                        "TdChr='" + ADC("TChr") + "' And TdNo=" + CStr(ADC("TNo")) + " And TdPrtKey='" + ctCurrPrtn + "'" + _
                        " and  Cast(TdRmStkRt*TdRmWt as Decimal(10,2))<>0 and  " + _
                                " Cast((TdRmStkRt*TdRmWt)" + IIF(ADC("TCnvRt") <> 0 And ADC("TMulDiv") = "M", "*", "/") + Str(ADC("TCnvRt")) + "  as Decimal(10,2))<>Cast(TdPurAmt as Decimal(10,2)) ")
     
       If wTdSr <> 0 Then Cancel = True: ErrMsg = "Stock Value doesnot match Purchase Amount " + IIF(UCase(ms_ChkPurStkVal) = UCase("SR"), " for Sr:'" + Str(wTdSr), "") + "."
     End If
       
     If UCase(ms_ChkPurStkVal) = UCase("VCH") And moCn.GetFldVal("Select Cast(Sum(Cast(TdRmStkRt*TdRmWt as Decimal(10,2)))" + IIF(ADC("TCnvRt") <> 0 And ADC("TMulDiv") = "M", "*", "/") + Str(ADC("TCnvRt")) + "  as Decimal(10,2))  From Txnd Where TdCoCd='" + ADC("TCoCd") + "' And TdTc='" + ADC("TTc") + "' And TdYy='" + ADC("TYy") + "' And " + _
                        "TdChr='" + ADC("TChr") + "' And TdNo=" + CStr(ADC("TNo")) + " And TdPrtKey='" + ctCurrPrtn + "'") <> Round(ADC("TTotAmt"), 0) Then _
                Cancel = True: ErrMsg = "Stock Value doesnot match Purchase Amount " + IIF(UCase(ms_ChkPurStkVal) = UCase("SR"), " for Sr:'" + Str(wTdSr), "") + "."
                       
       
       
    End If
   End If
  
 
  'MMAck-If Menu code is MA and GridA has rows then GridA Acknowledge will update as per
  ' the value of Acknowledge in X. if Acknowledge values in X and GridA are not equal then
  'this change will occur else no chnage
  If UCase(ADC.MenuCd) = UCase("MA") And ADC("TACK") <> "" And GrdTxndA.Rows - 1 > 0 Then
      For i = 1 To GrdTxndA.Rows - 1
        If GrdTxndA.Value(i, "TDACK") <> ADC("TACK") Then
           GrdTxndA.Value(i, "TDACK") = ADC("TACK")
           If Not GrdTxndA.SaveRec(i) Then
              ADC("TAck") = "N"
              MsgBox "Error while saving, cannot Acknowledge this voucher"
          End If
        End If
      Next i
  End If
      
  If mf_OLORYN Then
    If ADC("WCNVTOTVALUE") <> 0 Then
      Cancel = True
      ErrMsg = "Cannot exit the voucher when Total Base Value is not equal to zero"
    End If
  End If
  '578-33 to calculate totals
  If UCase(ADC.MenuCd) <> UCase("CNV") Then Call GetTotals
End Sub
Private Sub atxt_FldWhen(Index As Integer, ByVal IdName As String, Cancel As Boolean, ErrMsg As String)
  '*** Set Helps For Each Field
  '*** Do not allow the user to edit the voucher date if record exists in the 'A' grid
  '*** The help of TFrRmLoc should show all the locations other than Process locations
  '*** Similarly the help of TToRmLoc should show only the locs of type 'Xc' for 'REC'
  '       and 'CNV' entries
  '*** The help of TToRmLoc should show all the locs other than Process locs for entries
  '       'PR' and 'MV' entries
  '*** wVSGrp
      'It is a working field used for entering the various Rm group options according
      'to which the Iss/Recpt Rm Qty & Wt totals will be displayed in the
      'Voucher Summary Option
  
  '****** Sachin 3.02.0 - Locking Atxt based on Partition Key
  If (ADC("TPrtKey") <> ctCurrPrtn Or mf_DtLocked = True) _
                        And (IdName <> UCase("TCoCd") And IdName <> UCase("TTc") And IdName <> UCase("TTc") And _
                              IdName <> UCase("TYy") And IdName <> UCase("TChr") And IdName <> UCase("TNo") And _
                              IdName <> UCase("wVSGrp") And IdName <> UCase("wCpTTc") And _
                              IdName <> UCase("wCpTYy") And IdName <> UCase("wCpTChr") And _
                              IdName <> UCase("wCpTNo")) And IdName <> UCase("wCpGrpOpt") Then
    Cancel = True: ErrMsg = "Cannot Edit Locked/Previous Partition Entry": Exit Sub
  End If
' ########################  Manoj #### Ver: 2.0.6 #### Date: 25/12/2003  ########################
  If UCase$(ADC("TLockYN")) = UCase$("Y") And UCase$(IdName) <> UCase$("TLockYN") And _
        (IdName <> UCase("wVSGrp") And IdName <> UCase("wCpTTc") And _
        IdName <> UCase("wCpTYy") And IdName <> UCase("wCpTChr") And _
        IdName <> UCase("wCpTNo")) And IdName <> UCase("wCpGrpOpt") Then
    Cancel = True: ErrMsg = "Cannot Add/Edit/Delete Locked Record": Exit Sub
  End If
' ###########################################  Manoj  ###########################################
'
  If mf_ValidLocSeq = False And UCase(IdName) <> UCase("wVSGrp") Then Cancel = True: ErrMsg = "The User Does Not Have Access To One Of The Locs": Exit Sub
  
  Select Case IdName
  Case Is = UCase("TCoCd"), UCase("wCnCoCd"), UCase("wDtCoCd")
    '****** (Jen 2.14 Next) (added wDtCoCd)
    '****** Sachin 2.14.0 - [Multi-Company] - Contra Entry Fields ******
    Call HlpList.hCoCd
    
  Case Is = UCase("wCnTTc")
    Call HlpList.PMCd("TC", "'MV'")
    
  Case Is = UCase("wCnTChr")
      Call HlpList.vPSCd(ADC("wCnCoCd"), "CHR", ADC("wCnTTc"))
      
  Case Is = UCase("wCnTNo")
      Call HlpList.TNo(ADC("wCnCoCd"), ADC("wCnTTc"), ADC("wCnTYy"), ADC("wCnTChr"))
  '****** Sachin 2.14.0 - [Multi-Company] - Contra Entry Fields ******
    
  Case Is = UCase("wCpTTc")
 ' Sri has put back this filter for 5.4.0 as we could not properly complete the Copy Rm Feature for timely release
 '      Here is the line that allows all Tcs to be copied - Call HlpList.PMCd("TC", "'MV','PR','TZ','CNV','JT'")
    If ms_TcTyp = "MV" Or ms_TTc = "PR" Then
        Call HlpList.PMCd("TC", "'MV','PR'")
    Else
        Call HlpList.PMCd("TC", "'" + ms_TcTyp + "'")
    End If
  Case Is = UCase("TYy"), UCase("wCpTYy"), UCase("wCnTYy")
      Call HlpList.PMCd("YY")
  Case Is = UCase("wCpTChr")
    If UCase(ms_TcTyp) = "CNV" Then
      'CopyData-to select source copy TC
      ' Call HlpList.vPSCd(adc("wCpTCoCd"), "CHR", adc("TTc"))
      Call HlpList.vPSCd(ADC("wCpTCoCd"), "CHR", ADC("wCpTTc"))
    '*** Jay 3.1.1
    ElseIf UCase$(ms_TTc) = "PR" Or UCase(ms_TcTyp) = "MV" Or UCase(ms_TcTyp) = "PR" Then
      Call HlpList.vPSCd(ADC("wCPTCoCd"), "CHR", ADC("wCpTTc"))
    '*** Jay 3.1.1
    'Bef Jay 3.1.1
    'ElseIf UCase(ms_TcTyp) = "MV" Or UCase(ms_TcTyp) = "PR" Then
    '  Call HlpList.vPSCd(ADC("TCoCd"), "CHR", ADC("wCpTTc"))
    'Bef Jay 3.1.1
    End If
  Case Is = UCase("TChr")
  ' **** Zubin 211 **** '
    'Call HlpList.vPSCd(ADC("TCoCd"), "CHR", ADC("TTc"))
    If UCase(ADC.MenuCd) = UCase("MA") Then
      Call HlpList.vPSCd(ADC("TCoCd"), "CHR", ADC("TTc"))
    Else
      Call HlpList.vPSCd(ADC("TCoCd"), "CHR", ADC("TTc"), "Y")
    End If
  ' **** Zubin 211 **** '
  Case Is = UCase("TNo")
    '****** Sachin 3.02 24-12-07 - Help based on selected partition
    Call HlpList.TNo(ADC("TCoCd"), ADC("TTc"), ADC("TYy"), ADC("TChr"))
  Case Is = UCase("wCpTNo")
    '*** Jay 3.1.1 [added mv]
    If UCase$(ms_TTc) = "PR" Or UCase$(ms_TTc) = "MV" Or UCase$(ms_TTc) = "CNV" Then
      '****** Sachin 3.02.0 - Only Current Partition Vouchers Displayed in Copy Scope
      'CopyData- jEmr included
      If wTcTyp = "JT" Then
      Call HlpList.OmNo(ADC("wCpTCoCd"), ADC("wCpTTc"), ADC("wCpTYy"), ADC("wCpTChr"))
      Else
      Call HlpList.TNo(ADC("wCPTCoCd"), ADC("wCpTTc"), ADC("wCpTYy"), ADC("wCpTChr"), ctCurrPrtn)
      End If
    Else
      If wTcTyp = "JT" Then
      Call HlpList.OmNo(ADC("wCpTCoCd"), ADC("wCpTTc"), ADC("wCpTYy"), ADC("wCpTChr"))
      Else
      Call HlpList.TNo(ADC("TCoCd"), ADC("wCpTTc"), ADC("wCpTYy"), ADC("wCpTChr"), ctCurrPrtn)
      End If
    End If
      
  Case Is = UCase("TDt")
    If GrdTxndA.Rows - 1 > 0 Then Cancel = True: ErrMsg = "Cannot Edit Date, Child Records Exist": Exit Sub
    'MMAck-In acknowledge menucode can't change the value
    If UCase(ADC.MenuCd) = UCase("MA") Then Cancel = True: ErrMsg = "Cannot Modify Readonly Column": Exit Sub

   
  Case Is = UCase("TSuppCd")
    'MMAck-In acknowledge menucode can't change the value
   
    
    If UCase(ADC.MenuCd) = UCase("MA") Or UCase(ADC.MenuCd) = UCase("CNV") Then Cancel = True: ErrMsg = "Cannot Modify Readonly Column": Exit Sub
  '**** urmi Sales ****
    If UCase(ms_TTc) = "SA" Or UCase(ms_TTc) = "MV" Or UCase(ms_TTc) = "CNV" Then
      Call HlpList.CustCd("C")
    Else
      Call HlpList.CustCd("S")
    End If
  '********************
  Case Is = UCase("TFrRmLoc"), UCase("WCNLOCCD")
    'MMAck-Cannot change Location value if records exist in GridA
    If GrdTxndA.Rows - 1 > 0 Then Cancel = True: ErrMsg = "Cannot Edit RmLoc, Child Records Exist": Exit Sub
    'MMAck-In acknowledge menucode can't change the value
    If UCase(ADC.MenuCd) = UCase("MA") Then Cancel = True: ErrMsg = "Cannot Modify Readonly Column": Exit Sub
    'sv.10 R location added
    Call HlpList.Loc(ADC("TCoCd"), , "'P','R','XC'")
    
  Case Is = UCase("TToRmLoc")
    'MMAck-Cannot change Location value if records exist in GridA
    If GrdTxndA.Rows - 1 > 0 Then Cancel = True: ErrMsg = "Cannot Edit RmLoc, Child Records Exist": Exit Sub
    'MMAck-In acknowledge menucode can't change the value
    If UCase(ADC.MenuCd) = UCase("MA") Then Cancel = True: ErrMsg = "Cannot Modify Readonly Column": Exit Sub

    If ms_TcTyp = "REC" Or ms_TcTyp = "CNV" Then
      Call HlpList.Loc(ADC("TCoCd"), "'XC'")
    Else
      Call HlpList.Loc(ADC("TCoCd"), , "'P','R'") 'sv.10 R location added
    End If
  ' ########################  Manoj #### Ver: 2.0.6 #### Date: 25/12/2003  ########################
  Case Is = UCase("TLockYN")
    'MMAck-In acknowledge menucode can't change the value
    If UCase(ADC.MenuCd) = UCase("MA") Then Cancel = True: ErrMsg = "Cannot Modify Readonly Column": Exit Sub
    '*** Jay 2.14 *** (SEO Change)
    'If Trim$(TxtPwd.text) <> Trim$(ms_Pwd) And ADC("TLockYN") = "Y" Then
    If ADC("TLockYN") = "Y" And ms_PwdFlag = False Then _
        Cancel = True: ErrMsg = "Enter SEO Password to Edit Value": Exit Sub
    '*** Jay 2.14 *** (SEO Change)
    Call HlpList.PMCd("YN")
  ' ###########################################  Manoj  ###########################################
  Case Is = UCase("wFrRmIr")
    'MMAck-In acknowledge menucode can't change the value
    If UCase(ADC.MenuCd) = UCase("MA") Then Cancel = True: ErrMsg = "Cannot Modify Readonly Column": Exit Sub
    Call HlpList.PMCd("IR")
  Case Is = UCase("TLsLoc")
    Call HlpList.Loc(ADC("TCoCd"), "'L'")
  Case Is = UCase("wVSGrp")
    Call HlpList.PMCd("TBAGVSGR")
  Case Is = UCase("wCmName")
    'MMAck-In acknowledge menucode can't change the value
    If UCase(ADC.MenuCd) = UCase("MA") Or UCase(ADC.MenuCd) = UCase("CNV") Then Cancel = True: ErrMsg = "Cannot Modify Readonly Column": Exit Sub
  '**** urmi Sales ****
    If UCase(ms_TTc) = "SA" Or UCase(ms_TTc) = "CNV" Then
      Call HlpList.CmName("C")
    Else
      Call HlpList.CmName("S")
    End If
  '********************

  '*** Shilpa Emr 2.05 ****
  Case Is = UCase("wCpGrpOpt")
    Call HlpList.PMCd("VCHCPYGR")
  '*** Shilpa Emr 2.05 ****
  '*****************Geeta***Emr206 ***28/04/04***********
  Case Is = UCase("wSLocFr"), UCase("wSLocTo")
    Call HlpList.Loc(ADC("TCoCd"), "'W', 'S', 'L', 'M', 'B'")  'ctAvgLocTypSeq)
  Case Is = UCase("wRmCtgFr")
    Call HlpList.PMCd("RMCTG")
  Case Is = UCase("wRmSCtgFr"), UCase("wRmSCtgTo")
    Call HlpList.PSCd("RMSCTG", ADC("wRmCtgFr"))
  Case Is = UCase("wRmCdFr")
    Call HlpList.RmCd("N", "'" + ADC("wRmCtgFr") + "'", ADC("wRmSCtgFr"))
  Case Is = UCase("wRmCdTo")
    Call HlpList.RmCd("N", "'" + ADC("wRmCtgFr") + "'", ADC("wRmSCtgTo"))
  '*****************Geeta***Emr206 ***28/04/04***********
  '*********Geeta****Emr207***********
  Case Is = UCase("TCurrCd")
    If GrdTxndA.Rows - 1 > 0 Then Cancel = True: ErrMsg = "Cannot Edit Currency Code, Child Records Exist": Exit Sub
    Call HlpList.PMCd("CURNCY")
  
  ' ****** Sachin 2.13.0 [18-07-2006] - Multi-Currency ******
  Case Is = UCase("TCnvRt")
    If GrdTxndA.Rows - 1 > 0 Then Cancel = True: ErrMsg = "Cannot Edit Conversion Factor, Child Records Exist": Exit Sub
    'MMAck-In acknowledge menucode can't change the value
    If UCase(ADC.MenuCd) = UCase("MA") Then Cancel = True: ErrMsg = "Cannot Modify Readonly Column": Exit Sub

  ' ****** Sachin 2.13.0 [18-07-2006] - Multi-Currency ******
  Case Is = UCase("TCnvRtRs")
   If mf_OLORYN Then
    If moCn.RecSeek("Select 'a' from Txnd where TdLotMstRecYN='Y' and TdCoCd='" + ADC("TCoCd") + "' And TdTc='" + ADC("TTc") + "' And TdYy='" + ADC("TYy") + "' And " + _
                              "TdChr='" + ADC("TChr") + "' And TdNo=" + CStr(ADC("TNo"))) Then
      
      Cancel = True: ErrMsg = "Cannot Edit Conversion Factor Rs, records of new LotMst exist in Grid": Exit Sub
    End If
    End If
    'MMAck-In acknowledge menucode can't change the value
     If UCase(ADC.MenuCd) = UCase("MA") Then Cancel = True: ErrMsg = "Cannot Modify Readonly Column": Exit Sub
  'Manoj
  Case Is = UCase$("wCpTCoCd")
    Call HlpList.hCoCd
  Case Is = UCase$("WCCpdChk")
    Call HlpList.PMCd("YN")
  'Manoj
  ' Zubin 212
  Case Is = UCase("TLotNo")
    If GrdTxndA.Rows > 1 Then Cancel = True: ErrMsg = "Cannot Edit Lot No. As Child Record(s) Exist": Exit Sub
    
    'MMAck-In acknowledge menucode can't change the value
    If UCase(ADC.MenuCd) = UCase("MA") Then Cancel = True: ErrMsg = "Cannot Modify Readonly Column": Exit Sub

    Call HlpList.LotNo(ADC("TCoCd"), IIF(ADC("wFrRmIR") = "I", ADC("TFrRmLoc"), ADC("TToRmLoc")))
  Case Is = UCase("wGVLotNoFr")
    Call HlpList.LotNo(ADC("TCoCd"), ADC("wSLocFr"))
  Case Is = UCase("wGVLotNoTo")
    Call HlpList.LotNo(ADC("TCoCd"), ADC("wSLocTo"))
  ' Zubin 212
 
  '*** (Jen 2.13)
  Case Is = UCase("wCrReYy")
    Call HlpList.PMCd("YY")
  Case Is = UCase("wCrReChr")
    Call HlpList.PMCd("RETCHR")
  Case Is = UCase("wCrReNo")
    Call HlpList.ReNo(ADC("TCoCd"), ADC("wCrReYy"), ADC("wCrReChr"))
  Case Is = UCase("wCrRmCtgFr"), UCase("wCrRmCtgTo")
    Call HlpList.PMCd("RMCTG")
  '*** (Jen 2.13)
 
  '*** (Jen 2.14 Next)
  Case Is = UCase("wDtYy")
    Call HlpList.PMCd("YY")
  Case Is = UCase("wDtChr")
    Call HlpList.vPSCd(ADC("wDtCoCd"), "CHR", ADC("wDtTc"))
  Case Is = UCase("wDtNo")
    '****** Sachin 3.02.0 - Only Current Partition Vouchers in Help
    Call HlpList.TNo(ADC("wDtCoCd"), ADC("wDtTc"), ADC("wDtYy"), ADC("wDtChr"), ctCurrPrtn)
  '*** (Jen 2.14 Next)
 
  ' ***** Manali 3.6.0 - 23/10/09 - Copy Rm From Loc
  Case Is = UCase("wLocRmCtg")
    Call HlpList.PMCd("RMCTG")
    
  Case Is = UCase("wLocRmCdFr"), UCase("wLocRmCdTo")
    Call HlpList.RmCd("N", "'" + ADC("wLocRmCtg") + "'")
  ' ***** Manali 3.6.0 - 23/10/09 - Copy Rm From Loc
  
  '4.1.2
  Case Is = UCase("wTdBstOdTc")   'indigo.164 OS tc added
    Call HlpList.PMCd("TC", "'SO','OS'")
  Case Is = UCase("wTdBstOdYy")
    Call HlpList.PMCd("YY")
  Case Is = UCase("wTdBstOdChr")
    Call HlpList.vPSCd(ADC("TCoCd"), "CHR", ADC("wTdBstOdTc"), "Y")
  Case Is = UCase("wTdBstOdNo")
    Call HlpList.OmNo(ADC("TCoCd"), ADC("wTdBstOdTc"), ADC("wTdBstOdYy"), ADC("wTdBstOdChr"))
  'kbs.132 help for Odsr
  Case Is = UCase("wTdBstOdSr")
    Call HlpList.OdSr(ADC("TCoCd"), ADC("wTdBstOdTc"), ADC("wTdBstOdYy"), ADC("wTdBstOdChr"), ADC("wTdBstOdNo"))
  Case Is = UCase("wTdBstDmCd") 'indigo.164 SM design type added
    Call HlpList.DmCd("DM','SM")
  'CopyData - for Source RmCtg
  Case Is = UCase("WCPRMCTG")
    Call HlpList.PMCd("GLDAS")
  
  Case Is = UCase("TACK")
  'MMAck-In Rm movement entry, user can change TAck value in X only if no record exists in Grid A
    If GrdTxndA.Rows - 1 > 0 And UCase(ADC.MenuCd) <> UCase("MA") Then Cancel = True: ErrMsg = "Cannot Edit Acknowledgement, Child Records Exist": Exit Sub
    Call HlpList.PMCd("YN")
  Case Is = UCase("TDesc")
    'MMAck-In acknowledge menucode can't change the value
    If UCase(ADC.MenuCd) = UCase("MA") Then Cancel = True: ErrMsg = "Cannot Modify Readonly Column": Exit Sub
  Case Is = UCase("TSUPPPAYTERMS")
    'MMAck-In acknowledge menucode can't change the value
    If UCase(ADC.MenuCd) = UCase("MA") Then Cancel = True: ErrMsg = "Cannot Modify Readonly Column": Exit Sub
  Case Is = UCase("TCNVRTRS")
    'MMAck-In acknowledge menucode can't change the value
    If UCase(ADC.MenuCd) = UCase("MA") Then Cancel = True: ErrMsg = "Cannot Modify Readonly Column": Exit Sub
  Case Is = UCase("TBILLNO")
    'MMAck-In acknowledge menucode can't change the value
    If UCase(ADC.MenuCd) = UCase("MA") Then Cancel = True: ErrMsg = "Cannot Modify Readonly Column": Exit Sub
  Case Is = UCase("TBILLDT")
    'MMAck-In acknowledge menucode can't change the value
    If UCase(ADC.MenuCd) = UCase("MA") Then Cancel = True: ErrMsg = "Cannot Modify Readonly Column": Exit Sub
  Case Is = UCase("TBOENO")
    'MMAck-In acknowledge menucode can't change the value
    If UCase(ADC.MenuCd) = UCase("MA") Then Cancel = True: ErrMsg = "Cannot Modify Readonly Column": Exit Sub
  Case Is = UCase("TBOEDT")
    'MMAck-In acknowledge menucode can't change the value
    If UCase(ADC.MenuCd) = UCase("MA") Then Cancel = True: ErrMsg = "Cannot Modify Readonly Column": Exit Sub
  Case Is = UCase("TADDAMT1")
    'MMAck-In acknowledge menucode can't change the value
    If UCase(ADC.MenuCd) = UCase("MA") Then Cancel = True: ErrMsg = "Cannot Modify Readonly Column": Exit Sub
  Case Is = UCase("TADDAMT2")
    'MMAck-In acknowledge menucode can't change the value
    If UCase(ADC.MenuCd) = UCase("MA") Then Cancel = True: ErrMsg = "Cannot Modify Readonly Column": Exit Sub
  Case Is = UCase("TADDAMT3")
    'MMAck-In acknowledge menucode can't change the value
    If UCase(ADC.MenuCd) = UCase("MA") Then Cancel = True: ErrMsg = "Cannot Modify Readonly Column": Exit Sub
  Case Is = UCase("TADDAMT4")
    'MMAck-In acknowledge menucode can't change the value
    If UCase(ADC.MenuCd) = UCase("MA") Then Cancel = True: ErrMsg = "Cannot Modify Readonly Column": Exit Sub
  
  Case Is = UCase("WTDDESC")
    'MMAck-In acknowledge menucode can't change the value
    If UCase(ADC.MenuCd) = UCase("MA") Then Cancel = True: ErrMsg = "Cannot Modify Readonly Column": Exit Sub
  
      
  'Case Is = UCase("wTdBstRmCd")
  '  Call HlpList.RmCd("N", , , True)
'    If GrdTxndA.Row > 0 Then
'      If GrdTxndA.Value(GrdTxndA.Row, "TdRmCd") <> "" Then
'        Dim ws_Arr() As String
'        ws_Arr = Split(moCn.GetFldVal("Select RmCtg+','+RmSCtg From RmMst where " + IIF(gs_Partition = ctCurrPrtn, " RmPrtKey='" + ctCurrPrtn + "' And ", "") + " RmCd= '" + GrdTxndZ.Value(GrdTxndZ.Row, "TzRmZCd") + "'"), ",")
'        Call HlpList.RmCd("N", "'" + ws_Arr(0) + "'", ws_Arr(1), True)
'      Else
'        Call HlpList.RmCd("N", , , True)
'      End If
'    Else
'      Call HlpList.RmCd("N", , , True)
'    End If
  '*** (Jen 3.1.2)
 
  End Select
End Sub

Private Sub atxt_ReCalculate(Index As Integer, ByVal IdName As String, ByVal OldVal As Variant, ByVal RowNum As Integer, ByVal SourceName As String, ByVal SourceColName As String)
  '*** TChr on TTc                                                Normal
      'PDesc225('TC', adc(TTc))
  '*** TCnvRt on TDt                                              Normal
      'DtValue corresponding to max(DtDt) from DtTbl(DtTag= 'USD' and DtDt<= TDt)
      'USD stands for 'US Dollars'.
  '*** TLsLoc on TFrRmLoc                                         Normal
      'If TLsLoc is visible then TLsLoc= TFrRmLoc
  '*** TLsFrDt on TLsLoc                                          Normal
      'If TLsFrDt is visible and TLsLoc<> Blank then
      '   TLsFrDt= max(TLsToDt) from Txn(TTc= adc(TTc) and TLsLoc= adc(TLsLoc)) plus one day
      '   If rec does not exist then TLsFrDt= Current Date - 1
  '*** TLsToDt on TLsFrDt                                         Normal
      'if TLsToDt is visible and TLsFrDt<> '' then TLsToDt= Current Date - 1
  '*** wCtg on TdRmCd                                             Normal
      'If wCtg is visible then wCtg= RmCtg from RmMst(RmCd= TdRmCd)
  '*** wTotNetWt on TdFrRmDc,TdRmWt,wCtg                          Normal
      'If wTotNetWt is visible then
      'wTotNetWt= Sum of all Wts Issued - Sum of all Wts Received
      '           The RmPurityWt of all the RmCodes have to be considered while
      '           taking the sum.
  '*** wTotQty on TdFrRmDc,TdRmQty                                Normal
      'If wTotQty is visible then
      'wTotQty= Sum of all Qty Issued - Sum of all Qty Received
  '*** wTotWt on TdFrRmDc,TdRmWt                                  Normal
      'If wTotWt is visible then
      'wTotWt= Sum of all Wt Issued - Sum of all Wt Received
  '*** wRecPrc on TGldRec,TGldInLs                                Normal
      'If wRecPrc is visible then
      'wRecPrc= TGldRec / TGldInLs * 100
  '*** TNetAmt on TdFrRmDc,TdPurAmt                               Normal
      'If TNetAmt is visible then
      'TNetAmt= Sum of all Amt Issued - Sum of all Amt Received
  '*** TTotAmt on TNetAmt,TAddAmt1,TAddAmt2,TAddAmt3,TAddAmt4     Normal
      'If TTotAmt is visible then
      'TTotAmt= TNetAmt+ TAddAmt1+ TAddAmt2+ TAddAmt3+ TAddAmt4
  '*** wNetWt on TdRmCd,wTotNetWt                                 Normal, delete of GrdTxndA
      'If wNetWt is visible then
      'wNetWt= wTotNetWt/ PurityWt of the RmCode Of The current Row Of The grid GrdTxndA
  
  Dim i As Integer
  Dim wCmCd As String, wCustName As String
  Dim wSuppPayTerms As String     ' ***** Manali 3.10.0 - 12/03/12 - Supplier Payment Terms
  Dim wCmCurCd As String
  Select Case UCase(IdName)
  Case Is = UCase("TChr")
    If ADC.Mode = xNorm Then _
       ADC("TChr") = moCn.GetFldVal("Select vPValue from vParam where " + _
                                   "vPCoCd= '" + ADC("TCoCd") + "' and vPTyp='DEFCHR' " + _
                                   "and vPMCd='" + ADC("TTc") + "'")
                                   
  '****** Sachin 2.13.0 - [17-07-2006] Multi-Currency ******
  Case Is = UCase("TCnvRt")
    If ADC.Mode = xNorm Then
       ADC("TCnvRt") = CnvRt(moCn.GetFldVal("Select hBaseCurCd From Head Where HCd='" + ctSelfCmCd + "' And HCoCd='" + ADC("TCoCd") + "'"), ADC("TCurrCd"), ADC("TDt"))
    End If
  
  Case Is = UCase("TMulDiv")
    If ADC.Mode = xNorm Then
       ADC("TMulDiv") = MulDiv(moCn.GetFldVal("Select hBaseCurCd From Head Where HCd='" + ctSelfCmCd + "' And HCoCd='" + ADC("TCoCd") + "'"), ADC("TCurrCd"))
    End If
  '****** Sachin 2.13.0 - [17-07-2006] Multi-Currency ******
' Zubin 213
  Case Is = UCase("TCnvRtRs")
    If ADC.Mode = xNorm Then
       ADC("TCnvRtRs") = CnvRt(ADC("TCurrCd"), ctIndCurncy, ADC("TDt"))
    End If
    
  Case Is = UCase("TMulDivRs")
    If ADC.Mode = xNorm Then
       ADC("TMulDivRs") = MulDiv(ADC("TCurrCd"), ctIndCurncy)
    End If
  ' Zubin 213
  
  Case Is = UCase("TLsLoc")
    If ADC("TLsLoc").Visible And ADC.Mode = xNorm Then _
        ADC("TLsLoc") = ADC("TFrRmLoc")
        
  '### What exactly are TLsFrDt and TLsToDt ?
  Case Is = UCase("TLsFrDt")
    Dim wLsToDt As Variant
    '****** Sachin 3.02 24-12-07 - Current Partition
    If ADC.Mode = xNorm And ADC("TLsFrDt").Visible And ADC("TLsLoc") <> "" Then
        wLsToDt = moCn.GetFldVal("Select convert(varchar(8), max(TLsToDt), 3) from Txn where " + _
                                  "TCoCd='" + ADC("TCoCd") + "' and " + _
                                  "TTc='" + ADC("TTc") + "' and " + _
                                  "TLsLoc='" + ADC("TLsLoc") + IIF(gs_Partition = ctCurrPrtn, "' and TPrtKey='" + ADC("TPrtKey") + "'", ""))
        If wLsToDt = "" Then
            ADC("TLsFrDt") = moCn.SrvrDate - 1
        Else
            ADC("TLsFrDt") = CDate(wLsToDt) + 1
        End If
    End If
  Case Is = UCase("TLsToDt")
    If ADC.Mode = xNorm And ADC("TLsToDt").Visible And ADC("TLsFrDt") <> "" Then
        ADC("TLsToDt") = IIF(ADC("TLsFrDt") <= moCn.SrvrDate - 1, moCn.SrvrDate - 1, ADC("TLsFrDt"))
    End If
    
  Case Is = UCase("wRecPrc")
    If ADC.Mode = xNorm And ADC("wRecPrc").Visible Then _
        ADC("wRecPrc") = MWLib.Div(ADC("TGldRec"), ADC("TGldInLs")) * 100
  
  Case Is = UCase("TTotAmt")
    If ADC.Mode = xNorm And ADC("TTotAmt").Visible Then _
        ADC("TTotAmt") = ADC("TNetAmt") + ADC("TAddAmt1") + ADC("TAddAmt2") + ADC("TAddAmt3") + ADC("TAddAmt4")
  Case Is = UCase("TSuppCd")
    If ADC.Mode = xNorm Then
      '** Urmi Sales ***********
      ' ***** Manali 3.10.0 - 12/03/12 - Supplier Payment Terms
      If UCase(ms_TTc) = "SA" Or UCase(ms_TTc) = "MV" Then
        wCmCd = moCn.GetFldVal("Select CmCd from CustMst where CmCtg='C' and " + _
                               "CmName= '" + ADC("wCmName") + "'")
        wSuppPayTerms = moCn.GetFldVal("Select CmPayTerms from CustMst where CmCtg='C' and " + _
                               "CmName= '" + ADC("wCmName") + "'")
      ElseIf UCase(ms_TTc) = "CNV" Then
        wCmCd = "ZSELF"
        wSuppPayTerms = moCn.GetFldVal("Select CmPayTerms from CustMst where CmCtg='S' and " + _
                               "CmName= '" + ADC("wCmName") + "'")
      Else
      
        wCmCd = moCn.GetFldVal("Select CmCd from CustMst where CmCtg='S' and " + _
                               "CmName= '" + ADC("wCmName") + "'")
        wSuppPayTerms = moCn.GetFldVal("Select CmPayTerms from CustMst where CmCtg='S' and " + _
                               "CmName= '" + ADC("wCmName") + "'")
      End If
      '***********************
      If wCmCd <> ADC("TSuppCd") Then ADC("TSuppCd") = wCmCd: ADC("TSuppPayTerms") = wSuppPayTerms
    End If
  Case Is = UCase("wCmName")
    If ADC.Mode = xNorm Then
      '**Urmi Sales ***********
      ' ***** Manali 3.10.0 - 12/03/12 - Supplier Payment Terms
       If UCase(ms_TTc) = "SA" Or UCase(ms_TTc) = "MV" Or UCase(ms_TTc) = "CNV" Then
          wCustName = moCn.GetFldVal("Select CmName from CustMst where CmCtg='C' and " + _
                                     "CmCd= '" + ADC("TSuppCd") + "'")
          wSuppPayTerms = moCn.GetFldVal("Select CmPayTerms from CustMst where CmCtg='C' and " + _
                                 "CmCd= '" + ADC("TSuppCd") + "'")
       Else
         wCustName = moCn.GetFldVal("Select CmName from CustMst where CmCtg='S' and " + _
                                    "CmCd= '" + ADC("TSuppCd") + "'")
          wSuppPayTerms = moCn.GetFldVal("Select CmPayTerms from CustMst where CmCtg='S' and " + _
                                 "CmCd= '" + ADC("TSuppCd") + "'")
       End If
       
       If wCustName <> ADC("wCmName") Then ADC("wCmName") = wCustName: ADC("TSuppPayTerms") = wSuppPayTerms
    
    ElseIf ADC.Mode = xInit Then
       If UCase(ms_TTc) = "SA" Or UCase(ms_TTc) = "MV" Or UCase(ms_TTc) = "CNV" Then
          ADC("wCmName") = moCn.GetFldVal("Select CmName from CustMst where CmCtg='C' and " + _
                                          "CmCd= '" + ADC("TSuppCd") + "'")
       Else
          ADC("wCmName") = moCn.GetFldVal("Select CmName from CustMst where CmCtg='S' and " + _
                                          "CmCd= '" + ADC("TSuppCd") + "'")
       End If
    '********************
    End If
  
  '******************Geeta*Emr207************************
  Case Is = UCase("TCurrCd")
    If ADC.Mode = xNorm Then
      If UCase(ms_TTc) = "CNV" Then
        ADC("TCurrCd") = moCn.GetFldVal("Select HBaseCurCd  from Head where HCoCd = '" + ADC("TCoCd") + "'")
        
      End If
      If UCase(ms_TTc) = "PR" Or UCase(ms_TTc) = "SA" Or UCase(ms_TTc) = "MV" Then
        ADC("TCurrCd") = moCn.GetFldVal("Select CmCurCd from CustMst where CmCtg='" + IIF(UCase(ms_TTc) = "PR", "S", "C") + "' and " + _
                                           "CmCd= '" + ADC("TSuppCd") + "'")
        'ADC("TUSDYn") = IIF(UCase(wCmCurCd) = ctDefCurncy, "Y", "N")
      'Else
       ' ADC("TUSDYn") = "Y"
     End If
     
   End If
   '******************Geeta*Emr207************************
  'MMAck-if CHR AckYn is Y then Acknowledge in X can change into Y
  'if CHR AckYn is Y and User rights Loc is empty then Acknowledge in X can change into Y
  ' else RM from & To loc compare with user rights loc, if both are true then Ackyn=y
  'else AckYn=N
  Case Is = UCase("TAck")
  If ADC.Mode = xNorm And ADC("TChr") <> "" And ADC("TNO") > 0 Then
    If ms_TcTyp = "MV" Then
      If ms_MVAckIsReqdYN = "Y" Then ADC("TAck") = "N"
      If ms_MVAckIsReqdYN = "N" Then ADC("TAck") = "Y"
    Else
      ADC("TAck") = ""
    End If
  End If

    '      wb_FrFlag = IIF((InStr(1, "," + UCase(gs_ValidOthLocSeq) + ",", "," + UCase(ADC("TFRRMLOC")) + ",") <= 0), True, False)
    '      wb_ToFlag = IIF((InStr(1, "," + UCase(gs_ValidOthLocSeq) + ",", "," + UCase(ADC("TTORMLOC")) + ",") <= 0), True, False)
    '      If wb_FrFlag = False And wb_ToFlag = False Then
            'new Sri
     '       ADC("TAck") = "N"
      '    Else
       '     ADC("TAck") = "N"
        '  End If
       ' Else
        '  ADC("TAck") = "Y"
        'End If
'      End If
 '   End If
  Case Is = UCase("WCNVTOTRSVALUE")
    If Not mf_OLORYN Then Exit Sub
    Dim wRmCtg As String
    Dim wTot As Double
    If ADC.Mode = xNorm And ADC("WCNVTOTRSVALUE").Visible Then
      If GrdTxndA.Mode = fgmnorm Or GrdTxndA.Mode = fgmInitOver Or GrdTxndA.Mode = fgmdel Or GrdTxndA.Mode = fgmCopyOver Then
        wTot = 0
        For i = 1 To (GrdTxndA.Rows - 1)
          wRmCtg = moCn.GetFldVal("Select RmCtg from RmMst where RmCd='" + GrdTxndA.Value(i, "TdRmCd") + "'")
          If wRmCtg = "D" Or wRmCtg = "C" Or wRmCtg = "X" Then
            If Not GrdTxndA.IsDel(i) Then
               
            wTot = wTot + GetLotRsAmt(GrdTxndA.Value(i, "TdPurRt"), GrdTxndA.Value(i, "TdLotNo"), GrdTxndA.Value(i, "TdRtByQW"), GrdTxndA.Value(i, "TdRmQty"), GrdTxndA.Value(i, "TdRmWt")) * _
                                        IIF(GrdTxndA.Value(i, "TdFrRmDc") = "D", 1, _
                                        IIF(GrdTxndA.Value(i, "TdFrRmDc") = "C", -1, 0))
            End If
          End If
        Next i
        ADC("WCNVTOTRSVALUE") = wTot
      End If
    End If
  Case Is = UCase("WCNVTOTVALUE")
    If Not mf_OLORYN Then Exit Sub
    If ADC.Mode = xNorm And ADC("WCNVTOTVALUE").Visible Then
      If GrdTxndA.Mode = fgmnorm Or GrdTxndA.Mode = fgmInitOver Or GrdTxndA.Mode = fgmdel Or GrdTxndA.Mode = fgmCopyOver Then
        wTot = 0
        Dim wReqVal As Double
        For i = 1 To (GrdTxndA.Rows - 1)
          wRmCtg = moCn.GetFldVal("select RmCtg from RmMst where RmCd='" + GrdTxndA.Value(i, "TdRmCd") + "'")
          If wRmCtg = "D" Or wRmCtg = "C" Or wRmCtg = "X" Then
            wReqVal = GrdTxndA.Value(i, "TdPurRt") * IIF(GrdTxndA.Value(i, "TdRtByQW") = "Q", GrdTxndA.Value(i, "TdRmQty"), GrdTxndA.Value(i, "TdRmWt"))
            If Not GrdTxndA.IsDel(i) Then wTot = wTot + wReqVal * _
                                        IIF(GrdTxndA.Value(i, "TdFrRmDc") = "D", 1, _
                                        IIF(GrdTxndA.Value(i, "TdFrRmDc") = "C", -1, 0))
            
          End If
        Next i
        ADC("WCNVTOTVALUE") = wTot
      End If
    End If
  Case Is = UCase("WCNVTOTQTY")
    If Not mf_OLORYN Then Exit Sub
   
    If ADC.Mode = xNorm And ADC("WCNVTOTQTY").Visible Then
      If GrdTxndA.Mode = fgmnorm Or GrdTxndA.Mode = fgmInitOver Or GrdTxndA.Mode = fgmdel Or GrdTxndA.Mode = fgmCopyOver Then
        wTot = 0
        For i = 1 To (GrdTxndA.Rows - 1)
          wRmCtg = moCn.GetFldVal("select RmCtg from RmMst where RmCd='" + GrdTxndA.Value(i, "TdRmCd") + "'")
          If wRmCtg = "D" Or wRmCtg = "C" Or wRmCtg = "X" Then
            If Not GrdTxndA.IsDel(i) Then wTot = wTot + GrdTxndA.Value(i, "TdRmQty") * _
                                        IIF(GrdTxndA.Value(i, "TdFrRmDc") = "D", 1, _
                                        IIF(GrdTxndA.Value(i, "TdFrRmDc") = "C", -1, 0))
            
          End If
        Next i
        ADC("WCNVTOTQTY") = wTot
      End If
    End If
   Case Is = UCase("WCNVTOTWT")
    If Not mf_OLORYN Then Exit Sub
   
    If ADC.Mode = xNorm And ADC("WCNVTOTWT").Visible Then
      If GrdTxndA.Mode = fgmnorm Or GrdTxndA.Mode = fgmInitOver Or GrdTxndA.Mode = fgmdel Or GrdTxndA.Mode = fgmCopyOver Then
        wTot = 0
        For i = 1 To (GrdTxndA.Rows - 1)
          wRmCtg = moCn.GetFldVal("select RmCtg from RmMst where RmCd='" + GrdTxndA.Value(i, "TdRmCd") + "'")
          If wRmCtg = "D" Or wRmCtg = "C" Or wRmCtg = "X" Then
            If Not GrdTxndA.IsDel(i) Then wTot = wTot + GrdTxndA.Value(i, "TdRmWt") * _
                                        IIF(GrdTxndA.Value(i, "TdFrRmDc") = "D", 1, _
                                        IIF(GrdTxndA.Value(i, "TdFrRmDc") = "C", -1, 0))
            
          End If
        Next i
        ADC("WCNVTOTWT") = wTot
      End If
    End If
  Case Is = UCase("WCNVRSRT")
    ADC("wCnvRsRt") = MWLib.Div(ADC("WCNVTOTRSVALUE"), ADC("wcnvtotvalue"))
  End Select
End Sub
Private Function GetLotRsAmt(pTdPurRt As Double, pTdLotNo As String, pTdRtByQW As String, pRmQty As Double, pRmWt As Double) As Double
    Dim wRsLotNo As MDORowSet
    Dim wReqVal As Double
    Set wRsLotNo = moCn.OpenRes("Select * from LotMst where LotNo = '" + pTdLotNo + "'")
    If (wRsLotNo.RecCount > 0) Then
        Dim wReqRt As Double
        wReqRt = IIF(pTdRtByQW = "Q", wRsLotNo!LotRsRtByQ, _
        wRsLotNo!LotRsRtByW)
    Else
        wReqRt = IIF(ADC("TcnvRtRs") <> 0, IIF(ADC("TMulDivRs") = "D", _
        MWLib.Div(pTdPurRt, ADC("TCnvRtRs")), pTdPurRt * ADC("TCnvRtRs")), 0)
    End If
    wReqVal = IIF(pTdRtByQW = "Q", pRmQty, pRmWt) * wReqRt
    GetLotRsAmt = wReqVal
End Function

Private Sub ATXT_FldChange(Index As Integer, ByVal IdName As String, Cancel As Boolean, ErrMsg As String, ByVal pv_NewValue As Variant)
  '*** TTc,wCpTTc     PMCd('TC', PValue= ms_TcTyp) and the type of entry should be 'TxnRm'
  '*** TYy,wCpTYy     PMCd('YY')
  '*** TChr,wCpTChr   PSCd('CHR', adc(TTc)) or ''
  '*** TDt            Not Blank
  '*** TFrRmLoc       LocCd from Loc(LocTyp<> 'P') or ''
  '*** TToRmLoc       If ms_TcTyp= 'REC' Or ms_TcTyp= 'CNV' then LocCd from Loc(LocTyp= 'Xc')
  '                      or ''
  '                   else LocCd from Loc(LocTyp<> 'P') or ''
  '*** wFrRmIr        PMCd('IR')
  '*** TSuppCd        CmCd from CustMst(CmCtg= 'S')
  '*** TLsToDt        >= TLsFrDt or ''
  '*** wVSGrp         PMCd('TBAGVSGR')
  '                   Call the sub Procedure VouchSumm
  
  Select Case UCase(IdName)
  '****** Sachin 2.14.0 - 17-11-2006 [Multi-Company] ******
  Dim wToRmLoc As Boolean
  Case Is = UCase("TCoCd")
      Cancel = (pv_NewValue = ctSelfCoCd)
      If Cancel = True Then ErrMsg = "Company Code Cannot be 'ZZZ'": Exit Sub
  
      Cancel = Not moCn.RecSeek("Select 'x' From Head Where HCd='" + ctSelfCmCd + "' And HCoCd='" + pv_NewValue + "'")
      If Cancel = True Then ErrMsg = "Invalid Company Code": Exit Sub
  
      Cancel = Not moCn.RecSeek("Select 'x' From UsrAccess Where UaMnuCd= '" + ADC.MenuCd + "' And " + _
                         "UaUsrCd = '" + gs_UsrCd + "' And UaMnuForm='MDI' And (UaValidCoCd='' Or ''','+UaValidCoCd+',''' Like '%,'+'" + pv_NewValue + "'+',%')")
      If Cancel = True Then ErrMsg = "User Doesn't Have Access To This Company Code": Exit Sub
      
      If pv_NewValue <> ADC("TCoCd") Then ADC("TChr") = ""
      Me.Caption = GetFrmCaption(pv_NewValue, ms_FrmCaption)
      'sv.9 if login company code changed, getting purge date from vparam
      ms_PurgeDt = GetPurgeDt(pv_NewValue, "")
  '****** Sachin 2.14.0 - 17-11-2006 [Multi-Company] ******
  
  Case Is = UCase("TTc")
    If ms_TcTyp = "MV" Then
        Cancel = Not moCn.RecSeek("select PMCd from Param where PTyp= 'TC' and " + _
                         "PMCd = '" + pv_NewValue + "' and " + _
                         "PValue= 'MV' or PValue = 'PR' and " + _
                         "(Select Pvalue1 from Param where PTyp='TCTYP' and " + _
                         "PMCd= '" + pv_NewValue + "') ='TXNRM'")
        If Cancel = True Then ErrMsg = "Invalid Voucher Tc": Exit Sub
    Else
        Cancel = Not moCn.RecSeek("select PMCd from Param where PTyp= 'TC' and " + _
                         "PMCd = '" + pv_NewValue + "' and " + _
                         "PValue= '" + ms_TcTyp + "' and " + _
                         "(Select Pvalue1 from Param where PTyp='TCTYP' and " + _
                         "PMCd= '" + ms_TcTyp + "') ='TXNRM'")
        If Cancel = True Then ErrMsg = "Invalid Voucher Tc": Exit Sub
    End If
  Case Is = UCase("wCpTTc")
  'CopyData- TC validation
    Cancel = Not moCn.RecSeek("select PMCd from Param where PTyp= 'TC' and " + _
                         "PMCd = '" + pv_NewValue + "' and (PValue= 'MV' or " + _
                         "PValue = 'PR' or PValue = 'TZ' or PValue = 'CNV' or PValue = 'JT')")
    If Cancel = True Then ErrMsg = "Invalid Voucher Tc": Exit Sub
    wTcTyp = moCn.GetFldVal("Select PValue from Param where PTyp='TC' and PMCd='" + pv_NewValue + "'")
    wLooseTc = moCn.GetFldVal("Select PValue1 from Param where PTyp='TC' and PMCd='" + pv_NewValue + "' and PValue='JT'")
    If wTcTyp = "JT" And wLooseTc = "N" Then
      OptRt(0).Enabled = True: OptRt(1).Enabled = True: OptRt(1).Value = True
    Else
      OptRt(0).Enabled = False: OptRt(1).Enabled = False
    End If
    
  Case Is = UCase("TYy"), UCase("wCpTYy")
    Cancel = Not moCn.RecSeek("select PMCd from Param where PTyp= 'YY' and " + _
                         "PMCd = '" + pv_NewValue + "'")
    If Cancel = True Then ErrMsg = "Invalid Voucher Year": Exit Sub
  Case Is = UCase("TChr")
    Cancel = (pv_NewValue <> "") And ((Not moCn.RecSeek("select vPSCd from vParam where " + _
             "vPCoCd= '" + ADC("TCoCd") + "' and vPTyp= 'CHR' and " + _
             "vPMCd = '" + ADC("TTc") + "' and vPSCd='" + pv_NewValue + "'")))
' **** Zubin 211 **** '
'''''    ErrMsg = "Invalid Voucher Character": Exit Sub
    If Cancel = True Then ErrMsg = "Invalid Voucher Character": Exit Sub
    
  If UCase(ADC.MenuCd) <> UCase("MA") Then
    Dim ws_UsrArr() As String, wi_Index As Integer
    ws_UsrArr = Split(moCn.GetFldVal("Select vPValue4 from vParam where vPCoCd= '" + ADC("TCoCd") + "' and " + _
                                        "vPTyp='CHR' and vPMCd='" + ADC("TTc") + "' and vPSCd= '" + pv_NewValue + "'"), ",")
    If IsArray(ws_UsrArr) Then
      For wi_Index = 0 To UBound(ws_UsrArr)
        Cancel = IIF(UCase(gs_UsrCd) = UCase(ws_UsrArr(wi_Index)), False, True)
        If Cancel = False Then: Exit For
      Next wi_Index
    End If
   If Cancel = True Then ErrMsg = "Access To This Voucher Character Is Denied To User " + gs_UsrCd: Exit Sub
  End If
' **** Zubin 211 **** '
    
' ########################  Manoj #### Ver: 2.0.6 #### Date: 25/12/2003  ########################
  Case Is = UCase("TLockYN")
    Cancel = (pv_NewValue <> "") And ((Not moCn.RecSeek("select PMCd from Param where " + _
            "PTyp= 'YN' and PMCd = '" + pv_NewValue + "'")))
    If Cancel = True Then
      ErrMsg = "Invalid Custom Purchase Lock (Y/N)": Exit Sub
    Else
      '******************Geeta********Emr206******02/04/04****
      If pv_NewValue = "Y" Then
        ADC.AllowDelete = False: GrdTxndA.AllowDelete = False
        GrdTxndA.AllowAdd = False
      Else
        ADC.AllowDelete = True: GrdTxndA.AllowDelete = True
        GrdTxndA.AllowAdd = True
      End If
   End If
   Case Is = UCase("TLotNo")
    If mf_OLORYN Then
      If (Len(pv_NewValue) > 11) Then
        Cancel = True
        ErrMsg = "Length of Lot No must be less than 11 when OLOR is set to Yes"
      End If
    End If
    '****************************Geeta****************************
' ###########################################  Manoj  ###########################################
  Case Is = UCase("wCpTChr")
    'Bef Jay 3.1.1
    'If ms_TcTyp = "MV" Or ms_TTc = "PR" Then
    '  Cancel = (pv_NewValue <> "") And ((Not moCn.RecSeek("select vPSCd from vParam where " + _
    '         "vPCoCd= '" + ADC("TCoCd") + "' and vPTyp= 'CHR' and " + _
    '         "vPMCd = '" + ADC("wCPTTc") + "' and vPSCd='" + pv_NewValue + "'")))
    '  ErrMsg = "Invalid Voucher Character": Exit Sub
    'Bef Jay 3.1.1
    '*** Jay 3.1.1 [MV Added]
   'CopyData- CHR validation for user input
   'If ms_TcTyp = "MV" Or ms_TTc = "PR" Then
    '*** Jay 3.1.1 [MV Added]
      Cancel = ((Not moCn.RecSeek("select vPSCd from vParam where " + _
             "vPCoCd= '" + ADC("wCPTCoCd") + "' and vPTyp= 'CHR' and " + _
             "vPMCd = '" + ADC("wCPTTc") + "' and vPSCd='" + pv_NewValue + "'")))
      If Cancel = True Then ErrMsg = "Invalid Voucher Character": Exit Sub
    'Else   'CopyData- CHR validation for user input
      'Cancel = (pv_NewValue <> "") And ((Not moCn.RecSeek("select vPSCd from vParam where " + _
             "vPCoCd= '" + adc("TCoCd") + "' and vPTyp= 'CHR' and " + _
             "vPMCd = '" + adc("TTc") + "' and vPSCd='" + pv_NewValue + "'")))
      'ErrMsg = "Invalid Voucher Character": Exit Sub
    'End If
      
  Case Is = UCase("TDt")
    If pv_NewValue = "" Then Cancel = True: ErrMsg = "Transaction Date Cannot Be Blank": Exit Sub
    If pv_NewValue < CDate(ms_LockDt) Then Cancel = True: ErrMsg = "Transaction Date Cannot Be Less Than Lock Date": Exit Sub
    If pv_NewValue > CDate(ms_FutureDt) Then Cancel = True: ErrMsg = "Transaction Date Cannot Be Greater Than Future Date": Exit Sub
    If pv_NewValue <= CDate(ms_PurgeDt) Then Cancel = True: ErrMsg = "Transaction Date Cannot Be Less Than Or Equal To Purge Date": Exit Sub
    
        '53 company code added
    If moCn.GetFldVal("Select hHolChkYN From Head where HCd='ZSELF' and HCoCd='" + gs_CoCd + "'") = "Y" Then
        '53 company code added
        'If moCn.RecSeek("Select top 1 From DtTbl where DtTag='HOL' and DtCoCd in('" + ctSelfCoCd + "','" + ADC("TCoCd") + "')" + _
          " and convert(varchar(15),DtDt,3)= '" + CStr(pv_NewValue) + "' and DtHolYN='Y'")
        If IsHOLYN(ADC("TCoCd"), pv_NewValue) Then Cancel = True: ErrMsg = "Voucher Cannot be Added on Holidays": Exit Sub
    End If
    
    '4.3.0.0
    If pv_NewValue <> ADC("TDt") Then
      If GrdTxndA.Mode = fgmnorm And ADC.Mode = xNorm And GrdTxndA.Rows > 1 Then
        ADC.AddSave "GrdTxndA", asAll
      End If
    End If
  
  Case Is = UCase("wCpGrpOpt")
    Cancel = Not moCn.RecSeek("select PMCd from Param where PTyp= 'VCHCPYGR' and " + _
                         "PMCd = '" + pv_NewValue + "' and PScd = '' ")
    ErrMsg = "Invalid Group Option"
  Case Is = UCase("TFrRmLoc")
  'sv.10 R location type added
    Cancel = (pv_NewValue <> "") And (Not moCn.RecSeek("Select LocCd from Loc where " + _
                                          "LocCoCd = '" + ADC("TCoCd") + "' and " + _
                                          "LocCd = '" + pv_NewValue + "' and LocTyp<>'P' and LocTyp<>'R' "))
    If Cancel = True Then ErrMsg = "Invalid Raw Material From Location": Exit Sub
    If ms_TcTyp = "MV" Then
'      mf_ValidLocSeq = True
      If (gs_ValidOthLocSeq <> "") And ((InStr(1, "," + UCase(gs_ValidOthLocSeq) + ",", "," + pv_NewValue + ",") <= 0)) Then
        Cancel = True: ErrMsg = "No permission to this Location": Exit Sub
      ': mf_ValidLocSeq = False
      End If
    End If
    
  Case Is = UCase("TToRmLoc")
    If ms_TcTyp = "REC" Or ms_TcTyp = "CNV" Then
      Cancel = (pv_NewValue <> "") And (Not moCn.RecSeek("Select LocCd from Loc where " + _
                "LocCoCd = '" + ADC("TCoCd") + "' and " + _
                "LocCd = '" + pv_NewValue + "' and LocTyp= 'XC'"))
    Else
      'sv.10 R location type added
      Cancel = (pv_NewValue <> "") And (Not moCn.RecSeek("Select LocCd from Loc where " + _
                "LocCoCd = '" + ADC("TCoCd") + "' and " + _
                "LocCd = '" + pv_NewValue + "' and LocTyp<>'P' "))
      If Cancel = True Then ErrMsg = "Invalid Raw Material To Location": Exit Sub
    
      If ms_TcTyp <> "MV" Or (ms_TcTyp = "MV" And ms_MVAckIsReqdYN = "N") Then
        If (gs_ValidOthLocSeq <> "") And ((InStr(1, "," + UCase(gs_ValidOthLocSeq) + ",", "," + pv_NewValue + ",") <= 0)) Then
          Cancel = True: ErrMsg = "No permission to this Location":  Exit Sub
          ': mf_ValidLocSeq = False
        End If
      End If
    End If
    
  Case Is = UCase("wFrRmIr")
    Cancel = Not moCn.RecSeek("select PMCd from Param where PTyp= 'IR' and " + _
                         "PMCd = '" + pv_NewValue + "'")
    If Cancel = True Then ErrMsg = "Invalid From Rm IR": Exit Sub
  Case Is = UCase("TSuppCd")
  '******** urmi Sales **********
    If UCase(ms_TTc = "SA") Or UCase(ms_TTc) = "MV" Or UCase(ms_TTc) = "CNV" Then
      Cancel = Not moCn.RecSeek("Select CmCd From CustMst Where " + _
                                "CmCtg = 'C' and CmCd = '" + pv_NewValue + "'")
      If Cancel = True Then ErrMsg = "Invalid Customer Code": Exit Sub
    Else
      Cancel = Not moCn.RecSeek("Select CmCd From CustMst Where " + _
                                "CmCtg = 'S' and CmCd = '" + pv_NewValue + "'")
      If Cancel = True Then ErrMsg = "Invalid Supplier Code": Exit Sub
    End If
  '******************************
  
  Case Is = UCase("TLsToDt")
    Cancel = pv_NewValue <> "" And pv_NewValue < ADC("TLsFrDt")
    If Cancel = True Then ErrMsg = "Invalid Loss To Date": Exit Sub
  Case Is = UCase("wVSGrp")
    Cancel = Not moCn.RecSeek("select PMCd from Param where PTyp= 'TBAGVSGR' and " + _
                         "PMCd = '" + pv_NewValue + "'")
    If Cancel = True Then ErrMsg = "Invalid Group By Option": Exit Sub
    Call VouchSumm(pv_NewValue)
  Case Is = UCase("wCmName")
  '******** urmi Sales **********
    If UCase(ms_TTc = "SA") Or UCase(ms_TTc = "MV") Or UCase(ms_TTc = "CNV") Then
      Cancel = Not moCn.RecSeek("select CmName from CustMst where CmCtg= 'C' and CmName = '" + pv_NewValue + "'")
      If Cancel = True Then ErrMsg = "Invalid Customer Name": Exit Sub
    Else
      Cancel = Not moCn.RecSeek("select CmName from CustMst where CmCtg= 'S' and CmName = '" + pv_NewValue + "'")
      If Cancel = True Then ErrMsg = "Invalid Supplier Name": Exit Sub
    End If
  '******************************
  '*********Geeta****Emr207***********
  Case Is = UCase("TCurrCd")
    If UCase(ms_TTc = "PR") Then
      Cancel = Not moCn.RecSeek("select * from Param where PTyp= 'CURNCY' and PMCd = '" + pv_NewValue + "' and PsCd='' ")
      If Cancel = True Then ErrMsg = "Enter a Valid Currency Code": Exit Sub
    End If
  'Manoj
  Case Is = UCase$("wCpTCoCd")
    If UCase(ms_TTc) = "PR" Or UCase(ms_TTc) = "CNV" Then
      Cancel = Not moCn.RecSeek("Select hCoCd From Head Where HCd ='Zself' And HCoCd = '" & pv_NewValue & "'")
      ErrMsg = "Company Code Not Valid"
    End If
  Case Is = UCase$("wCpDetChk")
    If UCase(ms_TTc) = "CNV" Then
      Cancel = Not moCn.RecSeek("Select 1 From Param Where PTyp ='YN' And PMCd = '" & pv_NewValue & "'")
      ErrMsg = "Copy Detail Check Value Not Valid."
    End If
    If Cancel = False Then
      Dim wbChk As Boolean
      wbChk = IIF(UCase$(pv_NewValue) = "Y", True, False)
      'LblGrpOpt.Enabled = wbChk    'CopyData
      'LblGrpOpt.Visible = wbChk    'CopyData
      'adc("wCpTCoCd").Enabled = wbChk   'CopyData
      'adc("wCpTCoCd").Visible = wbChk   'CopyData
    End If
  'Manoj
 
  '*** (Jen 2.13)
  Case Is = UCase("wCrReYy")
    Cancel = (pv_NewValue <> "") And (Not moCn.RecSeek("select PMCd from Param where PTyp= 'YY' and " + _
                         "PMCd = '" + pv_NewValue + "'"))
    If Cancel = True Then ErrMsg = "Invalid Voucher Year": Exit Sub
  Case Is = UCase("wCrReChr")
    Cancel = (pv_NewValue <> "") And (Not moCn.RecSeek("select PMCd from Param where PTyp= 'RETCHR' and PMCd = '" + pv_NewValue + "'"))
    If Cancel = True Then ErrMsg = "Invalid Return Memo Voucher Character": Exit Sub
  Case Is = UCase("wCrReNo")
    Cancel = (pv_NewValue <> 0) And (Not moCn.RecSeek("select ReNo from RetMem where ReCoCd= '" + ADC("TCoCd") + "' and " + _
             "ReYy= '" + ADC("wCrReYy") + "' and ReChr= '" + ADC("wCrReChr") + "' and ReNo= " + CStr(pv_NewValue) + " "))
    If Cancel = True Then ErrMsg = "Invalid Return Memo Voucher Number": Exit Sub
  '*** (Jen 2.13)
  
  '*** (Jen 2.14 Next)
  Case Is = UCase("wDtCoCd")
    Cancel = Not moCn.RecSeek("Select 'x' From Head Where HCd= '" + ctSelfCmCd + "' And HCoCd= '" + pv_NewValue + "'")
    If Cancel = True Then ErrMsg = "Invalid Company Code": Exit Sub
  Case Is = UCase("wDtYy")
    Cancel = (pv_NewValue <> "") And (Not moCn.RecSeek("select PMCd from Param where PTyp= 'YY' and " + _
                                     "PMCd = '" + pv_NewValue + "'"))
    If Cancel = True Then ErrMsg = "Invalid Voucher Year": Exit Sub
  Case Is = UCase("wDtChr")
    Cancel = (pv_NewValue <> "") And (Not moCn.RecSeek("select vPSCd from vParam where vPCoCd= '" + ADC("wDtCoCd") + "' " + _
                                     "and vPTyp= 'CHR' and vPMCd= '" + ADC("wDtTc") + "' and vPSCd= '" + pv_NewValue + "'"))
    If Cancel = True Then ErrMsg = "Invalid Source Voucher Character": Exit Sub
  Case Is = UCase("wDtNo")
    '****** Sachin 3.02 24-12-07 - Current Partition
    Cancel = (pv_NewValue <> 0) And (Not moCn.RecSeek("select TNo from Txn where TCoCd= '" + ADC("wDtCoCd") + "' " + _
             "and TTc= '" + ADC("wDtTc") + "' and TYy= '" + ADC("wDtYy") + "' and TChr= '" + ADC("wDtChr") + "' " + _
             "and TNo= " + CStr(pv_NewValue) + " and TPrtKey='" + ctCurrPrtn + "' "))
    If Cancel = True Then ErrMsg = "Invalid Source Voucher Number": Exit Sub
  '*** (Jen 2.14 Next)
  
  Case Is = UCase("WCPTNO")
  'CopyData- jEmr validation added
    If wTcTyp = "JT" Then
    Cancel = (pv_NewValue <> 0) And (Not moCn.RecSeek("select OmNo from OrdMst where OmCoCd= '" + ADC("wCpTCoCd") + "' " + _
             "and OmTc= '" + ADC("WCPTTC") + "' and OmYy= '" + ADC("WCPTYY") + "' and OmChr= '" + ADC("WCPTCHR") + "' " + _
             "and OmNo= " + CStr(pv_NewValue) + " and OmPrtKey='" + ctCurrPrtn + "' "))
    If Cancel = True Then ErrMsg = "Invalid Voucher Number": Exit Sub
    Else
  '****** Sachin 3.02 24-12-07 - Current Partition
    Cancel = (pv_NewValue <> 0) And (Not moCn.RecSeek("select TNo from Txn where TCoCd= '" + ADC("wCpTCoCd") + "' " + _
             "and TTc= '" + ADC("WCPTTC") + "' and TYy= '" + ADC("WCPTYY") + "' and TChr= '" + ADC("WCPTCHR") + "' " + _
             "and TNo= " + CStr(pv_NewValue) + " and TPrtKey='" + ctCurrPrtn + "' "))
    If Cancel = True Then ErrMsg = "Invalid Voucher Number": Exit Sub
 '****** Sachin 2.14.0 - [Multi-Company]
    End If
 Case Is = UCase("WCNCOCD")
   If pv_NewValue = ADC("TCOCD") Or pv_NewValue = ctSelfCoCd Then
    Cancel = True: ErrMsg = "Contra Entry Cannot be Done in Same Company or 'ZZZ' Company": Exit Sub
   End If
 Case Is = UCase("wCnCoCd")
      Cancel = (pv_NewValue = ctSelfCoCd)
      If Cancel = True Then ErrMsg = "Company Code Cannot be 'ZZZ'": Exit Sub
  
      Cancel = Not moCn.RecSeek("Select 'x' From Head Where HCd='" + ctSelfCmCd + "' And HCoCd='" + pv_NewValue + "'")
      If Cancel = True Then ErrMsg = "Invalid Company Code": Exit Sub
      
 Case Is = UCase("wCnTTc")
    If pv_NewValue <> "MV" Then Cancel = True
    If Cancel = True Then ErrMsg = "Invalid Source Voucher Tc": Exit Sub
 '****** Sachin 2.14.0 - [Multi-Company]
 
  '4.1.2
  Case Is = UCase("wTdBstOdTc")
  'indigo.164 PValue= 'SO' changed
    Cancel = pv_NewValue <> "" And (Not moCn.RecSeek("select PMCd from Param where PTyp= 'TC' and " + _
                         "PMCd = '" + pv_NewValue + "' and PValue In('SO','OS') "))
    If Cancel = True Then ErrMsg = "Invalid Order Tc": Exit Sub
  Case Is = UCase("wTdBstOdYy")
    Cancel = pv_NewValue <> "" And (Not moCn.RecSeek("select PMCd from Param where PTyp= 'YY' and " + _
                         "PMCd = '" + pv_NewValue + "'"))
    If Cancel = True Then ErrMsg = "Invalid Year": Exit Sub
  Case Is = UCase("wTdBstOdChr")
    Cancel = pv_NewValue <> "" And (Not moCn.RecSeek("select vPSCd from vParam where " + _
                         "vPCoCd= '" + ADC("TCoCd") + "' and vPTyp= 'CHR' and " + _
                         "vPMCd= '" + ADC("wTdBstOdTc") + "' and vPSCd= '" + pv_NewValue + "'"))
    If Cancel = True Then ErrMsg = "Invalid Order Character": Exit Sub
  Case Is = UCase("wTdBstOdNo")
    Cancel = pv_NewValue <> 0 And (Not moCn.RecSeek("select OmNo from OrdMst where OmPrtKey='" + ctCurrPrtn + "' and " + _
                         "OmCoCd= '" + ADC("TCoCd") + "' and OmTc= '" + ADC("wTdBstOdTc") + "' " + _
                         "and OmYy= '" + ADC("wTdBstOdYy") + "' and OmChr= '" + ADC("wTdBstOdChr") + "' " + _
                         "and OmNo='" + pv_NewValue + "'"))
    If Cancel = True Then ErrMsg = "Invalid Order Number": Exit Sub
  'kbs.132 OdSr validation
  Case Is = UCase("wTdBstOdSr")
    Cancel = pv_NewValue <> 0 And (Not moCn.RecSeek("select OdSr from OrdDsg where OdPrtKey='" + ctCurrPrtn + "' and " + _
                         "OdCoCd= '" + ADC("TCoCd") + "' and OdTc= '" + ADC("wTdBstOdTc") + "' " + _
                         "and OdYy= '" + ADC("wTdBstOdYy") + "' and OdChr= '" + ADC("wTdBstOdChr") + "' " + _
                         "and OdNo=" + CStr(ADC("wTdBstOdNo")) + " And OdSr=" + CStr(pv_NewValue)))
    If Cancel = True Then ErrMsg = "Invalid Order Serial Number": Exit Sub
    
  Case Is = UCase("wTdBstDmCd")


  'indigo.164 DmTcTyp= 'DM' removed
    Cancel = pv_NewValue <> "" And (Not moCn.RecSeek("Select DmCd From DsgMst Where DmPrtKey='" + ctCurrPrtn + "' " + _
                              "and DmTcTyp='DM' and DmCd= '" + pv_NewValue + "' and DmSz= '' /* and DmHld= 'N' */ "))
    If Cancel = True Then ErrMsg = "Invalid Design Code": Exit Sub
  'CopyData - validation for source rmctg
  Case Is = UCase("WCPRMCTG")
    Cancel = pv_NewValue <> "" And (Not moCn.RecSeek("Select PMCd From Param Where PTyp ='GLDAS'" + _
                              "and PMCd= '" + pv_NewValue + "' and PSCd= ''"))
    If Cancel = True Then ErrMsg = "Invalid Rm Category": Exit Sub
  'MMAck-if user wants to change AckYN then it checks RmToLoc with User rights NonProd Loc
  'if user have permission then it will change otherwise gives error msg
  Case Is = UCase("TACK")
    If ms_TcTyp = "MV" Then
        Cancel = (Not moCn.RecSeek("Select PMCd From Param Where PTyp= 'YN' And " + _
                           "PMCd = '" + pv_NewValue + "' and PSCd= ''"))
        If Cancel = True Then ErrMsg = "Enter Acknowledgement Yes Or No (Y/N)": Exit Sub
         
       If pv_NewValue = "Y" Then
          If ms_MVAckIsReqdYN = "Y" And gs_ValidOthLocSeq <> "" Then
              Cancel = IIF((InStr(1, "," + UCase(gs_ValidOthLocSeq) + ",", "," + UCase(ADC("TTORMLOC")) + ",") <= 0), True, False)
          Else
            Cancel = False
          End If
       End If
       If Cancel = True Then ErrMsg = "Since No Access to To Location, cannot change the Ackd flag ": Exit Sub
    End If
   
   'MMACK-When user want to change TAck value from N to Y, acknowledgement value will pass from X to GridA.
   'System will check whether user has access to To Location or not for each record in GridA.
   'If user has access to all records in GridA then TdAck value will change else block all, no replacement
    If pv_NewValue = "Y" And GrdTxndA.Rows - 1 > 0 Then
    Dim i As Integer
      For i = 1 To GrdTxndA.Rows - 1
        If gs_ValidOthLocSeq <> "" Then
          Cancel = InStr(1, "," + UCase(gs_ValidOthLocSeq) + ",", "," + GrdTxndA.Value(i, "TdToRmLoc") + ",") <= 0
          If Cancel = True Then ErrMsg = "Rm To Location in Row:" + CStr(i) + " is Not Accessible to the User": Exit Sub
        End If
      Next i
    End If

  'Case Is = UCase("wTdBstRmCd")
  '  Cancel = pv_NewValue <> "" And (Not moCn.RecSeek("Select RmCd From RmMst Where RmPrtKey='" + ctCurrPrtn + "' and RmCd = '" + pv_NewValue + "' " + _
  '           "and RmZ= 'N' "))
  '  If Cancel = True Then ErrMsg = "Invalid Raw Material Code": Exit Sub
  '4.1.2

 End Select
End Sub

Private Sub ChkBoxArr_Click(Index As Integer)
    With ChkBoxArr(Index)
        Select Case Index
            Case Is = 0
                If .Value = Checked Then ADC("WVALYN") = "Y" Else ADC("WVALYN") = "N"
            Case Is = 1
                If .Value = Checked Then ADC("WREMSTKVALYN") = "Y" Else ADC("WREMSTKVALYN") = "N"
        End Select
    End With
End Sub

Private Sub CmdCDGo_Click()
    '4.3.0.0
    If ADC("wTDt") <> "01/01/80" And ADC("wTDt") <> ADC("TDt") Then ADC("TDt") = ADC("wTDt")
'    If Not ADC.SaveRec Then
'        DispMsg "X Save Failed", etInfo
'        Exit Sub
'    End If
    Call DispFra(ChgDt)
End Sub

Private Sub CmdChgDt_Click()
    If ADC("TLockYN") = "N" Then
        Call DispFra(ChgDt)       '4.3.0.0
    Else
        DispMsg "Cannot Change Date as Voucher is Locked", etError
    End If
End Sub

Private Sub CmdCntGo_Click()
'****** Sachin 2.14.0 - [Multi-Company]
  If GrdTxndA.Rows > 1 Then DispMsg "Cannot Copy As Record Already Exists", etError: Exit Sub
  If ms_TcTyp = "MV" Then
    If ADC("wCnCoCd") = ADC("TCoCd") Then
      DispMsg "A Contra Entry Can be Done only between Two Different Companies Other than 'ZZZ'", etError: Exit Sub
    Else
      If ADC("wCnCoCd") <> "" And ADC("wCnTTc") <> "" And ADC("wCnTYy") <> "" And ADC("wCnTChr") <> "" And ADC("wCnTNo") > 0 And ADC("wCnLocCd") <> "" Then
        Call GenContra
      Else
        DispMsg "All the Fields Should be Entered To Generate Contra Voucher", etError: Exit Sub
      End If
    End If
    '****** Sachin 2.14.0 - [Multi-Company]
  End If
  Call CmdContra_Click
'****** Sachin 2.14.0 - [Multi-Company]
End Sub

Private Sub CmdContra_Click()

  '****** Sachin 2.14.0 - [Multi-Company]
  '****** Sachin 3.02.0 - Partition Locking
  If ADC("TPrtKey") <> ctCurrPrtn Or mf_DtLocked = True Or UCase$(ADC("TLockYN")) = UCase$("Y") Then _
          DispMsg "Contra Not Allowed In Locked/Previous Partition Entry", etError: Exit Sub

  If mf_ValidLocSeq = False Then DispMsg "The User Does Not Have Access To One Of The Locs", etError: Exit Sub
  
  '*** Display or hide the Copy Frame depending on the toggle state
    Call DispFra(Contra)
  '****** Sachin 2.14.0 - [Multi-Company]

End Sub

Private Sub CmdCpLocRm_Click()
  If Not ADC.SaveRec Then
    DispMsg "X Save Failed", etInfo
    Exit Sub
  End If
  If mf_ValidLocSeq = False Then DispMsg "The User Does Not Have Access To One Of The Locs", etError: Exit Sub

  Call DispFra(CpyLocRm)      ' ***** Manali 3.6.0 - 23/10/09 - Copy Rm From Loc
End Sub

Private Sub CmdCpLocRmGo_Click()
' ***** Manali 3.6.0 - 23/10/09 - Copy Rm From Loc
If Not moCn.RecSeek("select PMCd from Param where PTyp= 'RMCTG' and " + _
                         "PMCd = '" + ADC("wLocRmCtg") + "'") Then
    DispMsg "Invalid RmCtg", etError
    Exit Sub
End If
Call CpyLocRmDet
Call DispFra(CpyLocRm)
' ***** Manali 3.6.0 - 23/10/09 - Copy Rm From Loc
End Sub
Private Sub CmdGrdTxndADets_Click()
  If GrdTxndA.Row = 0 Then Exit Sub
  Call DispFra(DET)
  With GrdTxndA
    If FraGrdTxndADetsDisp.Enabled = False Then
      '.Value(.Row, "TzDesc") = ADC("wTzDesc")
      .Value(.Row, "TdBstOdTc") = ADC("wTdBstOdTc")
      .Value(.Row, "TdBstOdYy") = ADC("wTdBstOdYy")
      .Value(.Row, "TdBstOdChr") = ADC("wTdBstOdChr")
      .Value(.Row, "TdBstOdNo") = ADC("wTdBstOdNo")
      'kbs.132 grid getting values from bset use OdSr textbox
      .Value(.Row, "TdBstOdSr") = ADC("wTdBstOdSr")
      .Value(.Row, "TdBstDmCd") = ADC("wTdBstDmCd")
    End If
  End With
End Sub

'**************Geeta***********Emr206*********23/04/04****
Private Sub CmdGV_Click()
ChkBoxArr(0).Value = Checked: ADC("WVALYN") = "Y"                '3.11.0
ChkBoxArr(1).Value = Unchecked: ADC("WREMSTKVALYN") = "N"        '3.11.0
Call DispFra(vch)
End Sub
Private Sub CmdGVGo_Click()
'****************************Geeta**********************Emr206**************
'Dim wo_rsSrcSYM As MDORowSet
'Dim wCnd As String
'Dim wRow As Integer
If Not moCn.RecSeek("select PMCd from Param where PTyp= 'RMCTG' and " + _
                         "PMCd = '" + ADC("wRmCtgfr") + "'") Then
    DispMsg "Invalid RmCtg", etError
    Exit Sub
End If

'3.11.0
If UCase(ADC("WVALYN")) = "N" And UCase(ADC("WREMSTKVALYN")) = "N" Then
    DispMsg "Please Select atleast One Option from Valuation Or Remove Stk Option, to Continue..", etError
    Exit Sub
End If

Call GenVchdet
Call DispFra(vch)
End Sub
'***********Geeta******************************

Private Sub CmdImpExc_Click()
  If Not ADC.SaveRec Then
    DispMsg "X Save Failed", etInfo
    Exit Sub
  End If
  If mf_ValidLocSeq = False Then DispMsg "The User Does Not Have Access To One Of The Locs", etError: Exit Sub

  Call DispFra(ImpExc)      ' ***** Manali 3.9.0 - Import From Excel

End Sub

Private Sub CmdImpExcGo_Click()
If Trim(ADC("WXLFILE")) = "" Then DispMsg "Please Specify a Proper Excel File Name !!", etError: Exit Sub
If ADC("WXLSHTNO") = 0 Or ADC("WXLFRROW") = 0 Or ADC("WXLTOROW") = 0 Then DispMsg "Please Specify a Proper Sheet No./From Row/To Row !!", etError: Exit Sub

'**********Bhavna - for imp excel cnv
If UCase(ADC.MenuCd) = "CNV" Then
    Call ImportExcelCnv
Else
    Call ImportExcel
End If
Call DispFra(ImpExc)

End Sub

Private Sub CmdImpExcPath_Click()
  CdImpExc.FileName = ""
  '**********Bhavna - for imp excel cnv
  If UCase(ADC.MenuCd) = "PR" Then
    CdImpExc.DialogTitle = "Select Purchase File "
  Else
    CdImpExc.DialogTitle = "Select Conversion Excel File "
  End If
  CdImpExc.Filter = "Excel Files(*.xls;*.xlsx)|*.xls;*.xlsx"
  CdImpExc.FilterIndex = 1
  CdImpExc.ShowOpen
  ADC("WXLFILE") = CdImpExc.FileName

End Sub

Private Sub CmdInvHd_Click()
  If Not ADC.SaveRec Then
    DispMsg "X Save Failed", etInfo
    Exit Sub
  End If
  If moCn.RecSeek("Select 'x' From UsrAccess where UaUsrCd='" + ADC.UsrCd + "'  and UaMnuCd='In'") Then
      If gs_InNo = 0 Then
        gs_InCoCd = ADC("TCoCd")
        gs_InTc = ADC("TTc")
        gs_InYy = ADC("TYy")
        gs_InChr = ADC("TChr")
        gs_InNo = ADC("TNo")
        FormSelect ("In")
      Else
       DispMsg "Invoice Header already open.", etError
      End If
  Else
      DispMsg "User has no access to Invoice Header menu.", etError
  End If

End Sub
Private Sub CmdLN_Click()
  If GrdTxndA.Rows = 1 Then Exit Sub
  If mf_ValidLocSeq = False Then DispMsg "The User Does Not Have Access To One Of The Locs", etError: Exit Sub
 
  If FraLotNo.Visible = False Then
    If UCase(GrdTxndA.ColProp(GrdTxndA.Col).ColName) = UCase("TdLotNo") And GrdTxndA.TxtText <> "" Then
      ADC("wLotNo") = GrdTxndA.TxtText
    End If
    If GrdTxndA.TxtText <> GrdTxndA.Value(GrdTxndA.Row, "TdLotNo") Then GrdTxndA.Value(GrdTxndA.Row, "TdLotNo") = ""
  End If
  Call DispFra(LotNo)
End Sub
Private Sub CmdLNGo_Click()
  If ADC("wLotNo") = "" Then DispMsg "New Lot No. Cannot Be Blank", etError: ADC("wLotNo").SetFocus: Exit Sub
  If GrdTxndA.Value(GrdTxndA.Row, "TdRmCd") <> "" And Not moCn.RecSeek("select 'x' from RmMst where RmZ= 'N' and " + _
          "RmCd= '" + GrdTxndA.Value(GrdTxndA.Row, "TdRmCd") + "' and RmPrtKey='" + ctCurrPrtn + "' and RmCtg IN " + ctLotNoRmCtg) Then
    DispMsg "Lot No. Allowed For 'D', 'C','X','G','P','S','L','M' Only", etInfo
    Exit Sub
  End If
  If moCn.RecSeek("select 'x' from SYyMm where SCoCd= '" + ADC("TCoCd") + "' and " + _
                              "SLotNo= '" + ADC("wLotNo") + "' and SPrtKey='" + ctCurrPrtn + "' ") Then
    DispMsg "Lot No. Already Exists", etError: ADC("wLotNo").SetFocus: Exit Sub
  End If
  If GrdTxndA.Rows > 1 And ADC("wLotNo") <> "" Then GrdTxndA.Value(GrdTxndA.Row, "TdLotNo") = ADC("wLotNo")
  Call CmdLN_Click
End Sub

Private Sub CmdShowTot_Click()
'578-33 calculating & showing total fields
  If Not ADC.SaveRec Then
    DispMsg "X Save Failed", etInfo
    Exit Sub
  End If
  If UCase(ADC.MenuCd) = UCase("CNV") Then _
    If GrdTxndA.Row = 0 Then DispMsg "Select Anyone Row from Grid", etInfo: Exit Sub
  
  Call GetTotals
  Call DispFra(ShowTot)
End Sub
'578-33
Private Sub CmdShowTotOk_Click()
  Call DispFra(ShowTot)
End Sub

Private Sub GrdTxndA_RowPostValid(ByVal RowNum As Integer, Cancel As Boolean, ErrMsg As String, ChangedColName As String)
          If Not moCn.RecSeek("Select TYy from Txn where " + _
                            "TCoCd='" + ADC("TCoCd") + "' and " + _
                            "TTc='" + ADC("TTc") + "' and TYy='" + ADC("TYy") + "' and " + _
                            "TChr='" + ADC("TChr") + "' and TNo=" + CStr(ADC("TNo"))) Then
            Cancel = True: ErrMsg = "||Invalid Voucher"
            Exit Sub
          End If
End Sub

Private Sub GrdTxndA_RowSave(ByVal RowNum As Integer, Cancel As Boolean, ErrMsg As String)
      Cancel = ValidateStkSize(GrdTxndA.Value(RowNum, "TDRMSZ"), GrdTxndA.Value(RowNum, "TDRMSZ2"), GrdTxndA.Value(RowNum, "TDRMSZ3"), ErrMsg)
End Sub
Private Sub GrdTxndA_RowWhen(ByVal RowNum As Integer)
  With GrdTxndA
    If .IsNew(RowNum) Then
      If .Value(RowNum, "TdSr") = 0 Then .Value(RowNum, "TdSr") = .MaxVal("TdSr") + 1
      .Value(RowNum, "TdPrtKey") = ctCurrPrtn
      ADC("wCtg") = ""
      If .Value(RowNum, "TdToRmLoc") = "" Then
        If RowNum = 1 Then .Value(RowNum, "TdToRmLoc") = ADC("TToRmLoc")
        If RowNum > 1 Then .Value(RowNum, "TdToRmLoc") = .Value(RowNum - 1, "TdToRmLoc")
      End If
      If ms_TTc = "PR" Then
          .Value(RowNum, "TdRtByQW") = "W"
      End If
      If UCase(ms_TcTyp) <> "MV" Then .Value(RowNum, "TdAck") = "Y"
      If UCase(ms_TcTyp) = "MV" Then
        .Value(RowNum, "TdAck") = ADC("TAck")
      End If
      
  
    End If
    ADC("wTdDesc") = .Value(RowNum, "TdDesc")
    
    
    If Not (mf_DtLocked = True Or ADC("TLockYN") = "Y") Then
      If UCase(ADC.MenuCd) = UCase("MV") And Not .IsNew(RowNum) And UCase$(.Value(RowNum, "TdAck")) = "Y" Then
        .AllowDelete = (gs_ValidOthLocSeq = "") Or ((InStr(1, "," + UCase(gs_ValidOthLocSeq) + ",", "," + UCase(ADC("TFRRMLOC")) + ",") > 0) And _
                        (InStr(1, "," + UCase(gs_ValidOthLocSeq) + ",", "," + UCase(ADC("TTORMLOC")) + ",") > 0))
      ElseIf UCase(ADC.MenuCd) = UCase("MV") And Not .IsNew(RowNum) And UCase$(.Value(RowNum, "TdAck")) = "N" Then
        .AllowDelete = (gs_ValidOthLocSeq = "") Or (InStr(1, "," + UCase(gs_ValidOthLocSeq) + ",", "," + UCase(ADC("TFRRMLOC")) + ",") > 0)
      ElseIf UCase(ADC.MenuCd) = UCase("MA") Then
        .AllowDelete = False
      ElseIf UCase(ADC.MenuCd) <> UCase("MV") Then
        .AllowDelete = (gs_ValidOthLocSeq = "") Or ((InStr(1, "," + UCase(gs_ValidOthLocSeq) + ",", "," + UCase(ADC("TFRRMLOC")) + ",") > 0) And _
                        (InStr(1, "," + UCase(gs_ValidOthLocSeq) + ",", "," + UCase(ADC("TTORMLOC")) + ",") > 0))
      End If
    Else
      .AllowDelete = False
    End If
        
     
  End With
  
  ' ***** Manali 3.6.0 - 19/11/09 - Pure Net Wt for 'X'
    
  ' Manoj 2.9.0 Patch
  mf_AvgRtFldChg = False
  ' Manoj 2.9.0 Patch
  ' Zubin 212
  CmdLN.Enabled = True
  ' Zubin 212
  
  '4.1.2
    ADC("wTdBstOdTc") = GrdTxndA.Value(RowNum, "TdBstOdTc")
    ADC("wTdBstOdYy") = GrdTxndA.Value(RowNum, "TdBstOdYy")
    ADC("wTdBstOdChr") = GrdTxndA.Value(RowNum, "TdBstOdChr")
    ADC("wTdBstOdNo") = GrdTxndA.Value(RowNum, "TdBstOdNo")
    'kbs.132 OdSr textbox getting values from grid
    ADC("wTdBstOdSr") = GrdTxndA.Value(RowNum, "TdBstOdSr")
    ADC("wTdBstDmCd") = GrdTxndA.Value(RowNum, "TdBstDmCd")
  '4.1.2
  
End Sub
Private Sub GrdTxndA_ColWhen(ByVal RowNum As Integer, ByVal ColName As String, Cancel As Boolean, ErrMsg As String, ByVal Action As MwfCtl.en_fgAction)
  '*** Set Helps For Each Field
  '*** If 1st row, then get the default values as
      '     TdFrRmLoc= TFrRmLoc
      '     wTdFrRmIr= wFrRmDc
      '     TdToRmLoc= TToRmLoc
      'If Row> 1 then get the defa values for TdFrRmLoc, wTdFrRmIr and TdToRmLoc from the
      'previous row of the grid.
  '*** Check the 'Sz Y/N' option in param('RMSCTG', RmCtg of TdRmCd)
      'If the option is 'N' do not allow the user to enter value for TdRmSz
  '*** Check the 'StkRt Y/N' option in param('RMSCTG', RmCtg of TdRmCd)
      'If the option is 'N' do not allow the user to enter value for TdRmStkRt
  
  '****** Sachin 3.02.0 - Locking Previous Partition Records
  If (GrdTxndA.Value(RowNum, "TdPrtKey") <> ctCurrPrtn Or mf_DtLocked = True) Then Cancel = True: ErrMsg = "Cannot Edit Locked/Previous Partition Entry": Exit Sub
  
' ########################  Manoj #### Ver: 2.0.6 #### Date: 25/12/2003  ########################
  If UCase$(ADC("TLockYN")) = UCase$("Y") Then Cancel = True: ErrMsg = "Cannot Add/Edit/Delete Locked Record": Exit Sub
' ###########################################  Manoj  ###########################################
  If mf_ValidLocSeq = False Then Cancel = True: ErrMsg = "The User Does Not Have Access To One Of The Locs": Exit Sub
  Dim wRmCtg As String
  Dim wRmSCtg As String
  Dim wRsLotNo As MDORowSet
  With GrdTxndA
  
    If UCase$(ADC.MenuCd) = "MV" And .IsNew(RowNum) <> True And UCase$(.Value(RowNum, "TdAck")) = "N" Then
      If gs_ValidOthLocSeq <> "" Then
        Cancel = IIF((InStr(1, "," + UCase(gs_ValidOthLocSeq) + ",", "," + UCase(ADC("TFRRMLOC")) + ",") <= 0), True, False)
      Else
        Cancel = False
      End If
    ElseIf UCase$(ADC.MenuCd) = "MV" And .IsNew(RowNum) <> True And UCase$(.Value(RowNum, "TdAck")) = "Y" Then
      If gs_ValidOthLocSeq <> "" Then
        Cancel = IIF((InStr(1, "," + UCase(gs_ValidOthLocSeq) + ",", "," + UCase(ADC("TTORMLOC")) + ",") <= 0), True, False)
      Else
        Cancel = False
      End If
    End If
    If Cancel = True Then ErrMsg = "Cannot Edit record As Loc. is Not Accessible to the User": Exit Sub
    
  
    If UCase(ADC.MenuCd) = UCase("MV") And Not .IsNew(RowNum) And UCase$(.Value(RowNum, "TdAck")) = "Y" Then
        .AllowDelete = (gs_ValidOthLocSeq = "") Or ((InStr(1, "," + UCase(gs_ValidOthLocSeq) + ",", "," + UCase(ADC("TTORMLOC")) + ",") > 0))
    ElseIf UCase(ADC.MenuCd) = UCase("MV") And Not .IsNew(RowNum) And UCase$(.Value(RowNum, "TdAck")) = "N" Then
        .AllowDelete = (gs_ValidOthLocSeq = "") Or (InStr(1, "," + UCase(gs_ValidOthLocSeq) + ",", "," + UCase(ADC("TFRRMLOC")) + ",") > 0)
    ElseIf UCase(ADC.MenuCd) = UCase("MA") Then
        .AllowDelete = False
    Else
        .AllowDelete = True
    End If
   
    Select Case ColName
    Case Is = UCase("TdFrRmLoc")
      ' SRI REMOVED cos its not enough time to correct this bug in 5.4.0 release
'      If moCn.GetFldVal("Select hStkRtNoEdit from Head where hcocd='" + gs_CoCd + "'") = "Y" And .Value(RowNum, "TdRmStkRt") <> 0 Then _
 '     ErrMsg = "Cannot Change From Location If StockRate <> 0": Cancel = True
            
      If (.IsNew(RowNum) And .Value(RowNum, "TdFrRmLoc") = "") Then
        If RowNum = 1 Then .Value(RowNum, "TdFrRmLoc") = ADC("TFrRmLoc")
        If RowNum > 1 Then .Value(RowNum, "TdFrRmLoc") = .Value(RowNum - 1, "TdFrRmLoc")
      End If
      Call HlpList.Loc(ADC("TCoCd"), , "'P','R','XC'") 'sv.10 R location added
      

'      sri
      If UCase(ADC.MenuCd) = "MV" Then
        If ms_MVAckIsReqdYN = "Y" Then
           Cancel = True: ErrMsg = "RM Fr Location is Not Editable when Acknowledgement is required": Exit Sub
        End If
      End If

    Case Is = UCase("wTdFrRmIr")
      ' SRI REMOVED cos its not enough time to correct this bug in 5.4.0 release
      'If moCn.GetFldVal("Select hStkRtNoEdit from Head where hcocd='" + gs_CoCd + "'") = "Y" And .Value(RowNum, "TdRmStkRt") <> 0 Then _
      'ErrMsg = "Cannot Change I/R If StockRate <> 0": Cancel = True

      If .IsNew(RowNum) And .Value(RowNum, "wTdFrRmIr") = "" Then
        If RowNum = 1 Then .Value(RowNum, "wTdFrRmIr") = ADC("wFrRmIr")
        If RowNum > 1 Then .Value(RowNum, "wTdFrRmIr") = .Value(RowNum - 1, "wTdFrRmIr")
      End If
      Call HlpList.PMCd("IR")
      If UCase(ADC.MenuCd) = "MV" Then
        If ms_MVAckIsReqdYN = "Y" Then
           Cancel = True: ErrMsg = "This value is Not editable when Acknowledgement is required": Exit Sub
        End If
      End If
    Case Is = UCase("TdRmCd")
      Call HlpList.RmCd("N")
    Case Is = UCase("TdRmSz")
      Dim wTdRmSzYn As String, wTdRmPtrYn As String
      '****** Sachin 3.02 24-12-07 - Current Partition
      wTdRmSzYn = moCn.GetFldVal("Select Pdesc225 from Param join RmMst on RmCd='" + _
                 .Value(RowNum, "TdRmCd") + "' and RmPrtKey='" + ctCurrPrtn + "' where PTyp='RmSCtg' and PmCd=RmCtg and " + _
                 " PScd = RmSCtg ")
      '****** Sachin 3.02 24-12-07 - Current Partition
      wTdRmPtrYn = moCn.GetFldVal("Select PValue3 from Param join RmMst on RmCd='" + _
                 .Value(RowNum, "TdRmCd") + "' and RmPrtKey='" + ctCurrPrtn + "' where PTyp='RmSCtg' and PmCd=RmCtg and " + _
                 " PScd = RmSCtg ")
      If wTdRmSzYn = "N" And wTdRmPtrYn = "N" Then ErrMsg = "Size/ Ptr Should Not Be Entered For This Rm SubCtg": Cancel = True
    Case Is = UCase("TdRmSz2")
      wRmCtg = moCn.GetFldVal("Select RmCtg from RmMst where RmCd= '" + .Value(RowNum, "TdRmCd") + "' and RmPrtKey='" + ctCurrPrtn + "' ")
      wRmSCtg = moCn.GetFldVal("Select RmSCtg from RmMst where RmCd= '" + .Value(RowNum, "TdRmCd") + "' and RmPrtKey='" + ctCurrPrtn + "' ")
      Cancel = ColWhenStkSz(2, wRmCtg, wRmSCtg, ErrMsg)
    Case Is = UCase("TdRmSz3")
      wRmCtg = moCn.GetFldVal("Select RmCtg from RmMst where RmCd= '" + .Value(RowNum, "TdRmCd") + "' and RmPrtKey='" + ctCurrPrtn + "' ")
      wRmSCtg = moCn.GetFldVal("Select RmSCtg from RmMst where RmCd= '" + .Value(RowNum, "TdRmCd") + "' and RmPrtKey='" + ctCurrPrtn + "' ")
      Cancel = ColWhenStkSz(3, wRmCtg, wRmSCtg, ErrMsg)
    Case Is = UCase("TdRmStkRt")
      Dim wTdRmStkRtYn  As String, wStkNoEdit As String
      wStkNoEdit = moCn.GetFldVal("Select hStkRtNoEdit from Head where hcocd='" + gs_CoCd + "'")
      wTdRmStkRtYn = moCn.GetFldVal("Select PValue1 from Param join RmMst on RmCd='" + _
                 .Value(RowNum, "TdRmCd") + "' and RmPrtKey='" + ctCurrPrtn + "' where PTyp='RmSCtg' and PmCd=RmCtg and " + _
                 " PScd = RmSCtg ")
      If wTdRmStkRtYn = "N" Then ErrMsg = "Stock Rate Should Not Be Entered For This Rm SubCtg": Cancel = True
      Dim wRmLoc As String, wRmLocTyp As String, wb_RmLocTyp As Boolean
      wRmLoc = IIF(.Value(RowNum, "wTdFrRmIr") = "I", _
                             .Value(RowNum, "TdFrRmLoc"), _
                             .Value(RowNum, "TdToRmLoc"))
                                                 
      wb_RmLocTyp = moCn.RecSeek("Select LocTyp from Loc where LocCoCd= '" + ADC("TCoCd") + "' " + _
                                    "and LocCd= '" + wRmLoc + "' " + _
                                    "and LocTyp in " + ctAvgLocTypSeq)
      wRmLocTyp = Trim(moCn.GetFldVal("Select LocTyp from Loc where LocCoCd= '" + ADC("TCoCd") + "' " + _
                                    "and LocCd= '" + wRmLoc + "' "))

      
      If moCn.GetFldVal("Select hAvgSRtYn from Head where HCoCd= '" + ms_TCoCd + "' and " + _
                      "HCd='" + ctSelfCmCd + "'") = "Y" And (.IsNew(RowNum) = True Or mf_AvgRtFldChg = True) _
                      And UCase(ADC("TTc")) <> "PR" Then
                      'SRI And And .Value(RowNum, "TdRmStkRt") = 0#
        If wb_RmLocTyp = True And wRmLocTyp <> "W" Then
              .Value(RowNum, "TdRmStkRt") = GetAvgStkRt(ADC("TCoCd"), wRmLoc, _
                                          .Value(RowNum, "TdRmCd"), .Value(RowNum, "TdLotNo"), .Value(RowNum, "TdRmSz"), .Value(RowNum, "TdRmSz2"), .Value(RowNum, "TdRmSz3"))
        End If
      End If

      mf_AvgRtFldChg = False

      If moCn.GetFldVal("Select hAvgSRtYn from Head where HCoCd= '" + gs_CoCd + "'") = "Y" And wStkNoEdit = "Y" Then
        'StkRt-if Stock rate NoEdit is yes and AvgStkRate is yes then can edit only the values for X loctyp
        If wRmLocTyp Like "X*" Then
          Call HlpList.StkRt(ADC("TCoCd"), IIF(.Value(RowNum, "wTdFrRmIr") = "I", .Value(RowNum, "TdFrRmLoc"), IIF(.Value(RowNum, "wTdFrRmIr") = "", ADC("TFrRmLoc"), .Value(RowNum, ("TdToRmLoc")))), .Value(RowNum, "TdRmCd"), .Value(RowNum, "TdRmSz"), .Value(RowNum, "TdRmSz2"), .Value(RowNum, "TdRmSz3"))
        Else
          Cancel = True
        End If
      Else
        If moCn.GetFldVal("Select hStkRtPassReq from Head where HCoCd= '" + ms_TCoCd + "' and " + _
                      "HCd='" + ctSelfCmCd + "'") = "Y" And UCase(ADC("TTc")) <> "PR" And wb_RmLocTyp Then
          If ms_PwdFlag = False Then
            Cancel = True
            ErrMsg = "SEO Password Is Require To Edit Stock Rate."
          Else
           If moCn.GetFldVal("Select hAvgSRtYn from Head where HCoCd= '" + ms_TCoCd + "' and HCd='" + ctSelfCmCd + "'") <> "Y" Then
              Call HlpList.StkRt(ADC("TCoCd"), IIF(.Value(RowNum, "wTdFrRmIr") = "I", .Value(RowNum, "TdFrRmLoc"), IIF(.Value(RowNum, "wTdFrRmIr") = "", ADC("TFrRmLoc"), .Value(RowNum, ("TdToRmLoc")))), .Value(RowNum, "TdRmCd"), .Value(RowNum, "TdRmSz"), .Value(RowNum, "TdRmSz2"), .Value(RowNum, "TdRmSz3"))
           End If
          End If
        Else
          If UCase(ADC("TTc")) <> "PR" And moCn.GetFldVal("Select hAvgSRtYn from Head where HCoCd= '" + ms_TCoCd + "' and HCd='" + ctSelfCmCd + "'") <> "Y" Then
            Call HlpList.StkRt(ADC("TCoCd"), IIF(.Value(RowNum, "wTdFrRmIr") = "I", .Value(RowNum, "TdFrRmLoc"), IIF(.Value(RowNum, "wTdFrRmIr") = "", ADC("TFrRmLoc"), .Value(RowNum, ("TdToRmLoc")))), .Value(RowNum, "TdRmCd"), .Value(RowNum, "TdRmSz"), .Value(RowNum, "TdRmSz2"), .Value(RowNum, "TdRmSz3"))
          End If
        End If
      End If
      
    Case Is = UCase("TdToRmLoc")
      ' SRI REMOVED cos its not enough time to correct this bug in 5.4.0 release      ' SRI REMOVED cos its not enough time to correct this bug in 5.4.0 release
'      If moCn.GetFldVal("Select hStkRtNoEdit from Head where hcocd='" + gs_CoCd + "'") = "Y" And .Value(RowNum, "TdRmStkRt") <> 0 Then _
 '     ErrMsg = "Cannot Change To Location If StockRate <> 0": Cancel = True
      
      If .IsNew(RowNum) And .Value(RowNum, "TdToRmLoc") = "" Then
        If RowNum = 1 Then .Value(RowNum, "TdToRmLoc") = ADC("TToRmLoc")
        If RowNum > 1 Then .Value(RowNum, "TdToRmLoc") = .Value(RowNum - 1, "TdToRmLoc")
      End If
      Call HlpList.Loc(ADC("TCoCd"), , "'P','R','XC'") 'sv.10 R location added
    ' Manoj 2.10.0
      If UCase(ADC.MenuCd) = "MV" Then
        If ms_MVAckIsReqdYN = "Y" Then
           Cancel = True: ErrMsg = "RM To Location is Not Editable when Acknowledgement is required": Exit Sub
        End If
      End If
    Case Is = UCase("TdDustWt")
      If moCn.GetFldVal("select LocTyp from Loc Where LocCoCd = '" & ADC("TCoCd") & "' And LocCd = '" & GrdTxndA.Value(RowNum, "TdFrRmLoc") & "'") <> "L" And _
          moCn.GetFldVal("select LocTyp from Loc Where LocCoCd = '" & ADC("TCoCd") & "' And LocCd = '" & GrdTxndA.Value(RowNum, "TdToRmLoc") & "'") <> "L" Then
        ErrMsg = "Without Any Loss Location Can not Edit Value."
        Cancel = True
      End If
    ' Manoj 2.10.0
    ' Zubin 212
    Case Is = UCase("TdPurRt")
        If Not mf_OLORYN Then Exit Sub

        wRmCtg = moCn.GetFldVal("Select RmCtg from RmMst where RmCd='" + GrdTxndA.Value(RowNum, "TdRmCd") + "'")
        If wRmCtg = "D" Or wRmCtg = "C" Or wRmCtg = "X" Then
           Set wRsLotNo = moCn.OpenRes("Select * from LotMst where LotNo = '" + .Value(RowNum, "TdLotNo") + "'")
           If (wRsLotNo.RecCount > 0) Then
            If (wRsLotNo!LotRmCtg = wRmCtg) Then
             If wRsLotNo!LotVchCoCd <> ADC("TCoCd") Or wRsLotNo!LotVchYy <> ADC("TYy") Or wRsLotNo!LotVchTc <> ADC("TTc") _
              Or wRsLotNo!LotVchChr <> ADC("TChr") Or wRsLotNo!LotVchNo <> ADC("TNo") Or wRsLotNo!LotVchSr <> GrdTxndA.Value(RowNum, "TdSr") Then
                 Cancel = True
                 ErrMsg = "Cannot edit rate as it is a child record"
             End If
            End If
           End If
        End If
    Case Is = UCase("TdLotNo")
      If .Value(RowNum, "TdRmCd") = "" Or moCn.RecSeek("select 'x' from RmMst where RmZ= 'N' and RmCd= '" + .Value(RowNum, "TdRmCd") + "' and RmPrtKey='" + ctCurrPrtn + "' and RmCtg IN " + ctLotNoRmCtg) Then
        Call HlpList.LotNo(ADC("TCoCd"), IIF(.Value(RowNum, "wTdFrRmIr") = "I", .Value(RowNum, "TdFrRmLoc"), .Value(RowNum, ("TdToRmLoc"))), .Value(RowNum, "TdRmCd"), .Value(RowNum, "TdRmSz"), , .Value(RowNum, "TdRmSz2"), .Value(RowNum, "TdRmSz3"))
      Else
        Cancel = True: ErrMsg = "Lot No. Allowed For 'D', 'C','X','G','P','S','L','M' Only": Exit Sub
      End If
    Case Is = UCase("TdRtByQW")
     If mf_OLORYN Then
      'SRITEMP
      'Rajeshbhai has changed specs to always allow edit of QW
'       wRmCtg = moCn.GetFldVal("Select RmCtg from RmMst where RmCd='" + GrdTxndA.Value(RowNum, "TdRmCd") + "'")
'       If wRmCtg = "D" Or wRmCtg = "C" Or wRmCtg = "X" Then
'           Set wRsLotNo = moCn.OpenRes("Select * from LotMst where LotNo = '" + .Value(RowNum, "TdLotNo") + "'")
 '          If (wRsLotNo.RecCount > 0) Then
  '          If (wRsLotNo!LotRmCtg = wRmCtg) Then
   '          Cancel = True
    '         ErrMsg = "Cannot edit TdRtByQW for a child LotMst record"
     '        Exit Sub
      '      End If
      '    End If
       'End If
       Cancel = False
     Else
      If Not moCn.RecSeek("Select 'x' from RmMst where RmZ= 'N' and RmCd= '" + .Value(RowNum, "TdRmCd") + "' and RmPrtKey='" + ctCurrPrtn + "' and RmCtg IN ('X','D','C')") Then
        Cancel = True: ErrMsg = "Cannot Edit Value": Exit Sub
      End If
     End If
     Call HlpList.PMCd("LABQW")    ' ***** Manali 3.8.0 - QW
    'MMAck-User can move into this field, can't change the value.Value get from X
    Case Is = UCase("TdAck")
      Cancel = True
    End Select
  End With
End Sub
Private Sub GrdTxndA_ReCalc(ByVal RowNum As Integer, ByVal ColName As String, ByVal OldVal As Variant, ByVal SourceName As String, ByVal SourceColName As String)
  '*** wTdFrRmIr on TdFrRmDc                    Init
      'If TdFrRmDc='C' then 'I'
      'If TdFrRmDc='D' then 'R'
  '*** TdFrRmDc on wTdFrRmIr                    Normal
      'If wTdFrRmIr= 'I' then TdFrRmDc= 'C'
      'If wTdFrRmIr= 'R' then TdFrRmDc= 'D'
  '*** TdRmSz on TdRmCd                         Normal
      ' If the Ctg or SubCtg of the old value of TdRmCd is different from the
      ' Ctg or SubCtg of the new value of TdRmCd then TdRmSz = 0
  '*** TdRmStkRt on TdRmCd                      Normal
      ' If the Ctg or SubCtg of the old value of TdRmCd is different from the
      ' Ctg or SubCtg of the new value of TdRmCd then TdRmStkRt= 0
  '*** TdRmQty on TdRmCd,TdPcPerCt,TdRmWt       Normal
      ' If TcType = 'PR' then
      '     Recalculation for TdRmQty has to be done only when TdPcPerCt <> TdRmQty/TdRmWt
      '     TdRmQty = TdPcPerCt* TdRmWt
      ' Else
      '     If SourceColName = TdRmCd then
      '     if the Ctg or SubCtg of the old value of TdRmCd is different from the
      '     Ctg or SubCtg of the new value of TdRmCd then TdRmQty= 0
  '*** TdPcPerCt on TdRmCd,TdRmQty              Normal
      ' If the Ctg or SubCtg of the old value of TdRmCd is different from the
      ' Ctg or SubCtg of the new value of TdRmCd then TdPcPerCt= 0
      ' TdPcPerCt = TdRmQty/TdRmWt
  '*** TdRmWt on TdRmCd                         Normal
      ' If the Ctg or SubCtg of the old value of TdRmCd is different from the
      ' Ctg or SubCtg of the new value of TdRmCd then TdRmWt= 0
  '*** TdPurRt on TdRmCd                        Normal
      'TdPurRt= 0
  '*** TdPurAmt on TdPurRt,TdRmWt               Normal
      'TdPurAmt= TdPurRt* TdRmWt
  '*** TdDesc on wTdDesc                        Normal
      'If TdDesc<> wTdDesc then TdDesc= wTdDesc
      'Save the grid row
  
  Dim wRmCtg As String
  Dim wTdFrRmLocTyp As String
  Dim wTdToRmLocTyp As String
   Dim wRsLotNo As MDORowSet
  With GrdTxndA
    Select Case ColName
    Case Is = UCase("wTdFrRmIr")
      If .Mode = fgminit Then
        If .Value(RowNum, "TdFrRmDc") = "C" Then .Value(RowNum, "wTdFrRmIr") = "I"
        If .Value(RowNum, "TdFrRmDc") = "D" Then .Value(RowNum, "wTdFrRmIr") = "R"
      End If
    Case Is = UCase("TdFrRmDc")
      If .Mode = fgmnorm Then
        If .Value(RowNum, "wTdFrRmIr") = "I" Then .Value(RowNum, "TdFrRmDc") = "C"
        If .Value(RowNum, "wTdFrRmIr") = "R" Then .Value(RowNum, "TdFrRmDc") = "D"
      End If
    Case Is = UCase("TdRmSz")
      If .Mode = fgmnorm Then
        If FnRmCtg(.Value(RowNum, "TdRmCd")) <> FnRmCtg(OldVal) Or _
           FnRmSCtg(.Value(RowNum, "TdRmCd")) <> FnRmSCtg(OldVal) Then _
           .Value(RowNum, "TdRmSz") = 0
      End If
    Case Is = UCase("TdRmStkRt")
      If .Mode = fgmnorm Then
        Dim wTdRmStkRtYn  As String, wQwFact As Double    ' ***** Manali 3.8.0
        wTdRmStkRtYn = moCn.GetFldVal("Select PValue1 from Param join RmMst on RmCd='" + _
                   .Value(RowNum, "TdRmCd") + "' and RmPrtKey='" + ctCurrPrtn + "' where PTyp='RmSCtg' and PmCd=RmCtg and " + _
                   " PScd = RmSCtg ")
        
        If wTdRmStkRtYn = "Y" And UCase(SourceColName) = UCase("TdPurRt") Then
          If UCase(ADC("TTc")) = "PR" And .Value(RowNum, "TdRmStkRt") = 0 Then
              wQwFact = IIF(.Value(RowNum, "TdRtByQW") = "Q", MWLib.Div(.Value(RowNum, "TdRmQty"), .Value(RowNum, "TdRmWt")), 1)
              .Value(RowNum, "TdRmStkRt") = IIF(ADC("TMulDiv") = "M", _
               MWLib.Div(.Value(RowNum, "TdPurRt"), ADC("TCnvRt")), .Value(RowNum, "TdPurRt") * ADC("TCnvRt")) * wQwFact
          End If
        End If
        If UCase(SourceColName) = UCase("TdRmCd") Then
          If FnRmCtg(.Value(RowNum, "TdRmCd")) <> FnRmCtg(OldVal) Or _
            FnRmSCtg(.Value(RowNum, "TdRmCd")) <> FnRmSCtg(OldVal) Then _
            .Value(RowNum, "TdRmStkRt") = 0
          End If
      End If
    Case Is = UCase("TdRmQty")
      If .Mode = fgmnorm Then
        If ms_TcTyp = "PR" Then
          If FnRmCtg(.Value(RowNum, "TdRmCd")) = "D" Or FnRmCtg(.Value(RowNum, "TdRmCd")) = "C" Then
            If .Value(RowNum, "TdPcPerCt") <> Round(MWLib.Div(.Value(RowNum, "TdRmQty"), _
               .Value(RowNum, "TdRmWt")), 0) And .Value(RowNum, "TdRmWt") Then _
               .Value(RowNum, "TdRmQty") = Round(.Value(RowNum, "TdPcPerCt") * .Value(RowNum, "TdRmWt"), 0)
          End If
        Else
          If UCase(SourceColName) = UCase("TdRmCd") Then
            If FnRmCtg(.Value(RowNum, "TdRmCd")) <> FnRmCtg(OldVal) Or _
               FnRmSCtg(.Value(RowNum, "TdRmCd")) <> FnRmSCtg(OldVal) Then _
               .Value(RowNum, "TdRmQty") = 0
          End If
        End If
      End If
    Case Is = UCase("TdPcPerCt")
      If .Mode = fgmnorm Then
        If ms_TcTyp = "PR" Then
          If UCase(SourceColName) = UCase("TdRmCd") Then
            If FnRmCtg(.Value(RowNum, "TdRmCd")) <> FnRmCtg(OldVal) Or _
               FnRmSCtg(.Value(RowNum, "TdRmCd")) <> FnRmSCtg(OldVal) Then _
               .Value(RowNum, "TdPcPerCt") = 0
          End If
          If .Value(RowNum, "TdRmQty") <> 0 And .Value(RowNum, "TdRmWt") <> 0 Then .Value(RowNum, "TdPcPerCt") = Round(MWLib.Div(.Value(RowNum, "TdRmQty"), .Value(RowNum, "TdRmWt")), 0)
        End If
      End If
    Case Is = UCase("TdRmWt")
      If .Mode = fgmnorm Then
        If FnRmCtg(.Value(RowNum, "TdRmCd")) <> FnRmCtg(OldVal) Or _
           FnRmSCtg(.Value(RowNum, "TdRmCd")) <> FnRmSCtg(OldVal) Then _
           .Value(RowNum, "TdRmWt") = 0
      End If
    Case Is = UCase("TdPurRt")
       If mf_OLORYN And .Mode = fgmnorm Then
          UpdateTdPurRt (RowNum)
       Else
        If .Mode = fgmnorm Then .Value(RowNum, "TdPurRt") = 0
       End If
    Case Is = UCase("TdPurAmt")
        '*** (Jen 2.14 Next)
        If .Mode = fgmnorm Or .Mode = fgmCopy Then
          ' **** Manali 3.8.0
          'If ms_TTc = "PR" Then
            .Value(RowNum, "TdPurAmt") = .Value(RowNum, "TdPurRt") * IIF(.Value(RowNum, "TdRtByQW") = "Q", .Value(RowNum, "TdRmQty"), .Value(RowNum, "TdRmWt"))
          'Else
          ' **** Manali 3.8.0
          '  .Value(RowNum, "TdPurAmt") = .Value(RowNum, "TdPurRt") * .Value(RowNum, "TdRmWt")
          'End If
        End If
        '*** (Jen 2.14 Next)
        
        '*** (Bef 2.14 Next)
        'If .Mode = fgmnorm Then .Value(RowNum, "TdPurAmt") = _
        '                       .Value(RowNum, "TdPurRt") * .Value(RowNum, "TdRmWt")
        '*** (Bef 2.14 Next)
    Case Is = UCase("TdDesc")
      If .Mode = fgmnorm Then
        If .Row > 0 Then
          If .Value(.Row, "TdDesc") <> ADC("wTdDesc") Then
            .Value(.Row, "TdDesc") = ADC("wTdDesc")
            .SaveRec .Row
          End If
        End If
      End If
    ' Manoj 2.10.0
    Case Is = UCase("TdDustWt")
      If .Mode = fgmnorm Then
        If .Row > 0 Then
          .Value(.Row, "TdDustWt") = 0
        End If
      End If
    ' Manoj 2.10.0
' Zubin 212
    Case Is = UCase("TdLotNo")
      If moCn.RecSeek("Select 'x' from Head where " + _
                                      "HCd= '" + ctSelfCmCd + "' and HCoCd= '" + ADC("TCoCd") + "' and HLotNoYN= 'Y'") Then
        If .Row > 0 And .Mode = fgmnorm Then
        
          '****** Sachin 3.02 24-12-07 - Current Partition
          ' **** Manali 3.03 - 03/07/08 - Lot No Open for 'X'
          If mf_OLORYN Then
                   wTdFrRmLocTyp = moCn.GetFldVal("Select LocTyp From Loc Where LocCoCd='" + ADC("TCoCd") + "' And LocCd='" + .Value(RowNum, "TdFrRmLoc") + "'")
                   wTdToRmLocTyp = moCn.GetFldVal("Select LocTyp From Loc Where LocCoCd='" + ADC("TCoCd") + "' And LocCd='" + .Value(RowNum, "TdToRmLoc") + "'")
                   wRmCtg = moCn.GetFldVal("Select RmCtg From RmMst Where RmCd='" + .Value(RowNum, "TdRmCd") + "'")
                   If ChkForOLORYN(wTdFrRmLocTyp, wTdToRmLocTyp, wRmCtg, GrdTxndA.Value(RowNum, "TdFrRmDc")) And ADC("TLOTNO") <> "" Then
                    .Value(RowNum, "TdLotNo") = ADC("TLotNo") + "." + Format(.Value(RowNum, "TdSr"), "0000")
                   End If
               
          Else
            If Not moCn.RecSeek("select 'x' from RmMst where RmZ= 'N' and RmCd= '" + .Value(RowNum, "TdRmCd") + IIF(gs_Partition = ctCurrPrtn, "' and RmPrtKey='" + ctCurrPrtn + "' ", "") + _
                                      "and RmCtg IN " + ctLotNoRmCtg) Then
              .Value(RowNum, "TdLotNo") = ""
            Else
              If .IsNew(RowNum) Then
                If .Row = 1 And .Value(RowNum, "TdLotNo") = "" Then
        
                    .Value(RowNum, "TdLotNo") = ADC("TLotNo")
        
                ElseIf .Value(RowNum, "TdLotNo") = "" Then
                  .Value(RowNum, "TdLotNo") = .Value(RowNum - 1, "TdLotNo")
                End If
              End If
            End If
          End If
        End If
      Else
        If RowNum > 0 And .Mode = fgmnorm Then .Value(.Row, "TdLotNo") = ""
      End If
    Case Is = UCase("TdRtByQW")
      If .Mode = fgmnorm Then
        UpdateTdRtByQw (RowNum)
      End If
    End Select
  End With
End Sub

Private Function ChkForOLORYN(wTdFrRmLocTyp As String, wTdToRmLocTyp As String, wRmCtg As String, wTdFrRmDc As String) As Boolean
       If ((wRmCtg = "D" Or wRmCtg = "C" Or wRmCtg = "X") And ((Left(wTdFrRmLocTyp, 1) = "X" And wTdFrRmDc = "C") Or _
         (Left(wTdToRmLocTyp, 1) = "X" And wTdFrRmDc = "D"))) Then
          ChkForOLORYN = True
        Else
          ChkForOLORYN = False
       End If
End Function
Private Sub GrdTxndA_ColChange(ByVal RowNum As Integer, ByVal ColName As String, ByVal NewValue As Variant, Cancel As Boolean, ErrMsg As String)
  '*** TdFrRmLoc        valid LocCd from Loc(LocTyp<> 'P')
  '*** wTdFrRmIr        valid PMCd('IR')
  '*** TdRmCd           Blank or valid RmCd From RmMst where RmCd not a custom code
  '*** TdToRmLoc        If ms_TcTyp= 'REC' Or ms_TcTyp= 'CNV' then LocCd from
  '                        Loc(LocTyp= 'Xc')
  '                     else LocCd from Loc(LocTyp<> 'P')
   Dim wRmCtg As String
   Dim wTdFrRmLocTyp As String
   Dim wTdToRmLocTyp As String
  
  With GrdTxndA
    Select Case ColName
    Case Is = UCase("TdFrRmLoc")
    'sv.10 R location type added
      
      Cancel = Not moCn.RecSeek("Select LocCd From Loc Where " + _
                           "LocCoCd = '" + ADC("TCoCd") + "' and " + _
                           "LocTyp<>'P' and LocTyp<>'R' and LocCd = '" + NewValue + "' And LocValidYN='Y'") '****** Sachin 2.13.0 - ValidYn added ******
      If Cancel = True Then ErrMsg = "Invalid Raw Material From Location": Exit Sub

      If ms_MVAckIsReqdYN = "Y" Then
      '****** Sachin 2.13.0 - 24/06/2006 ******
        If gs_ValidOthLocSeq <> "" Then
           Cancel = InStr(1, "," + UCase(gs_ValidOthLocSeq) + ",", "," + _
                        UCase(NewValue) + ",") <= 0
           If Cancel = True Then ErrMsg = "RM From Location is Not Accessible to the User": Exit Sub
        End If
      Else
        'check  for permissions based on ADC("wIR")....
        
      End If
        
      '****** Sachin 2.13.0 - 24/06/2006 ******

    '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    '        If gs_ValidOthLocSeq <> "" Then
    '           Cancel = IIF((InStr(1, "," + UCase(gs_ValidOthLocSeq) + ",", "," + _
    '                        UCase(.Value(RowNum, "TdToRmLoc")) + ",") <= 0), True, False)
    '           Cancel = Cancel And (UCase(NewValue) = "R")
    '        Else
    '          Cancel = False
    '        End If
    '        If Cancel = True Then ErrMsg = "RM To Location is Not Accessible to the User": Exit Sub
    '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
      
    Case Is = UCase("wTdFrRmIr")
      Cancel = Not moCn.RecSeek("Select PMCd From Param Where PTyp= 'IR' And " + _
                           "PMCd = '" + NewValue + "'")
      If Cancel = True Then ErrMsg = "Enter Issue Or Receipt (I/R)": Exit Sub
      
    Case Is = UCase("TdRmCd")
      Cancel = (Not moCn.RecSeek("Select RmCd From RmMst where RmZ='N' and " + _
                           "RmCd ='" + NewValue + "' and RmPrtKey='" + ctCurrPrtn + "'")) And NewValue <> ""
      ErrMsg = "Invalid Raw Material Code": Exit Sub
      
    Case Is = UCase("TdToRmLoc")
      If ms_TcTyp = "REC" Or ms_TcTyp = "CNV" Then
        Cancel = Not moCn.RecSeek("Select LocCd From Loc Where " + _
                             "LocCoCd = '" + ADC("TCoCd") + "' and " + _
                             " LocCd = '" + NewValue + "' And LocValidYN='Y'") '****** Sachin 2.13.0 - ValidYn added ******
      Else
      'sv.10 R location type added
        Cancel = Not moCn.RecSeek("Select LocCd From Loc Where " + _
                                    "LocCoCd = '" + ADC("TCoCd") + "' and " + _
                                    "LocTyp<>'P' and LocTyp<>'R' and LocCd = '" + NewValue + "' And LocValidYN='Y'") '****** Sachin 2.13.0 - ValidYn added ******
      End If
      If Cancel = True Then ErrMsg = "Invalid Raw Material To Location": Exit Sub
    ' Zubin 212
    
      '****** Sachin 2.13.0 - 24/06/2006 ******
      'MMAck-if user have no access to Rm To location then can't change To location
      'but in MV entry user can make entry, NO update in stock if Ack=N
      If UCase(ADC.MenuCd) = "MV" Then
        If gs_ValidOthLocSeq <> "" And ADC("TAck") = "Y" Then
           Cancel = InStr(1, "," + UCase(gs_ValidOthLocSeq) + ",", "," + _
                        UCase(NewValue) + ",") <= 0
           If Cancel = True Then ErrMsg = "RM To Location is Not Accessible to the User": Exit Sub
        End If
      Else
        If gs_ValidOthLocSeq <> "" Then
           Cancel = InStr(1, "," + UCase(gs_ValidOthLocSeq) + ",", "," + _
                        UCase(NewValue) + ",") <= 0
           If Cancel = True Then ErrMsg = "RM To Location is Not Accessible to the User": Exit Sub
        End If
      End If
      '****** Sachin 2.13.0 - 24/06/2006 ******
        
        
    Case Is = UCase("TdLotNo")
     
      wTdFrRmLocTyp = moCn.GetFldVal("Select LocTyp From Loc Where LocCoCd='" + ADC("TCoCd") + "' And LocCd='" + .Value(RowNum, "TdFrRmLoc") + "'")
      wTdToRmLocTyp = moCn.GetFldVal("Select LocTyp From Loc Where LocCoCd='" + ADC("TCoCd") + "' And LocCd='" + .Value(RowNum, "TdToRmLoc") + "'")
      wRmCtg = moCn.GetFldVal("Select RmCtg From RmMst Where RmCd='" + .Value(RowNum, "TdRmCd") + "'")
      
     If mf_OLORYN Then
        If wRmCtg = "D" Or wRmCtg = "C" Or wRmCtg = "X" Then
         Cancel = (NewValue = "")
         ErrMsg = "LotNo is reqd for RmCtg D,C,X"
        Else
         Cancel = NewValue <> ""
         ErrMsg = "LotNo must be empty"
        End If
        If Cancel Then Exit Sub
        
      
      If _
        ChkForOLORYN(wTdFrRmLocTyp, wTdToRmLocTyp, wRmCtg, GrdTxndA.Value(RowNum, "TdFrRmDc")) Then
         Dim wReqRmCtg As String
         wReqRmCtg = moCn.GetFldVal("Select LotRmCtg from LotMst where " + _
                "LotNo= '" + NewValue + "'")
         If wReqRmCtg <> "" And wReqRmCtg <> wRmCtg Then
          Cancel = True
          ErrMsg = "Lot No already exists and belongs to a different Rm Ctg"
         End If
        If ms_TcTyp = "CNV" And ((wRmCtg = "D" Or wRmCtg = "C" Or wRmCtg = "X") And (((wTdFrRmLocTyp) = "XC" And GrdTxndA.Value(RowNum, "TdFrRmDc") = "C") Or _
         ((wTdToRmLocTyp) = "XC" And GrdTxndA.Value(RowNum, "TdFrRmDc") = "D"))) Then
          If wReqRmCtg <> "" Then
            Cancel = True
            ErrMsg = "Whenever XC is credited it has to be a new Lot No"
          End If
        End If
      Else
        If wRmCtg = "D" Or wRmCtg = "C" Or wRmCtg = "X" Then
         Cancel = (Not moCn.RecSeek("Select 'x' from LotMst where " + _
                "LotNo= '" + NewValue + "' and '" + wRmCtg + "' = LotRmCtg "))
         ErrMsg = "Lot No does not exist or Lot No belongs to another RmCtg"
        End If
      End If
     Else
        Cancel = NewValue <> "" And (Not moCn.RecSeek("Select 'x' from SYyMm where " + IIF(gs_Partition = ctCurrPrtn, " SPrtKey='" + ctCurrPrtn + "' And ", "") + " SCoCd= '" + ADC("TCoCd") + "' and " + _
                "SLotNo= '" + NewValue + "'")) And NewValue <> ADC("wLotNo") And NewValue <> ADC("TLotNo")
        ErrMsg = "Invalid Lot No.": Exit Sub
     End If
     
    End Select

    ' Manoj 2.9.0 Patch
        'For AvgRtFldChg
    Select Case ColName
      Case Is = UCase$("TdRmCd"), UCase$("TdRmSz"), UCase$("wTdFrRmIr")
        mf_AvgRtFldChg = True
      Case Is = UCase$("TdFrRmLoc")
        If UCase$(GrdTxndA.Value(RowNum, "wTdFrRmIr")) = "I" Then mf_AvgRtFldChg = True
      Case Is = UCase$("TdToRmLoc")
        If UCase$(GrdTxndA.Value(RowNum, "wTdFrRmIr")) = "R" Then mf_AvgRtFldChg = True
    ' Manoj 2.9.0 Patch
    
    ' ***** Manali 3.8.0
    Case Is = UCase("TdRtByQW")
     ' If ms_TTc = "PR" Then
        Cancel = Not moCn.RecSeek("Select 'x' from Param where PTyp='LABQW' and PMCd='" + NewValue + "' ")
        ErrMsg = "Invalid QW Option": Exit Sub
     
      'End If
  
  End Select
  End With
End Sub
Private Sub GrdTxndA_Valid()
  '*** if the TcType is 'REC' then show how much pure gold wt in Factory Loss Location
      'and how much pure gold has been recovered

  Dim wTGldInLs As Double, wTGldRec As Double, i As Integer
  Dim wctg  As String, wPurWt As Double, wLocTyp As String

  If ms_TcTyp = "REC" Then
      wTGldInLs = 0
      wTGldRec = 0
      With GrdTxndA
        For i = 1 To (.Rows - 1)
          '****** Sachin 3.02 24-12-07 - Current Partition
          wctg = moCn.GetFldVal("select RmCtg from RmMst where RmCd='" + .Value(i, "TdRmCd") + "' and RmPrtKey='" + ctCurrPrtn + "'")
          wPurWt = moCn.GetFldVal("select RmPurityWt from RmMst where RmCd='" + .Value(i, "TdRmCd") + "' and RmPrtKey='" + ctCurrPrtn + "'")
          wLocTyp = moCn.GetFldVal("select LocTyp from Loc where " + _
                                  "LocCoCd = '" + ADC("TCoCd") + "' and LocCd='" + .Value(i, "TdFrRmLoc") + "'")
          If Trim(wctg) = "G" And Trim(wLocTyp) = "L" Then
              If Not .IsDel(i) Then wTGldInLs = wTGldInLs + .Value(i, "TdRmWt") * wPurWt * _
                                      IIF(.Value(i, "TdFrRmDc") = "C", 1, _
                                      IIF(GrdTxndA.Value(i, "TdFrRmDc") = "D", -1, 0))
          End If
          If Trim(wctg) = "G" And Trim(wLocTyp) = "S" Then
              If Not .IsDel(i) Then wTGldRec = wTGldRec + .Value(i, "TdRmWt") * wPurWt * _
                                      IIF(.Value(i, "TdFrRmDc") = "D", 1, _
                                      IIF(GrdTxndA.Value(i, "TdFrRmDc") = "C", -1, 0))
          End If
        Next i
      End With
      ADC("TGldInLs") = wTGldInLs
      ADC("TGldRec") = wTGldRec
  End If
  ' Zubin 212
  '*** Disable all the command buttons associated only with the Grid GrdTxndA
  'CmdLN.Enabled = False
  ' Zubin 212
End Sub

Private Sub GrdTxndA_InitKey(ByVal RowNum As Integer)

  '*** Initialise Key Fields and all the fields which are not included in the grid
  Dim ws_TdKey As Long, ws_TdYyKey As String, wf_TdKeyLen As Single
  With GrdTxndA
    If .IsNew(RowNum) Then
      .Store "TdCoCd", ADC("TCoCd")
      .Store "TdTc", ADC("TTc")
      .Store "TdYy", ADC("TYy")
      .Store "TdChr", ADC("TChr")
      .Store "TdNo", ADC("TNo")
      .Store "TdSrNo", 0

      ws_TdKey = moCn.AutoGen("Txnd", "TdKey", ADC("TCoCd") + ADC("TYy"))
      .Store "TdKey", ws_TdKey
      wf_TdKeyLen = Len(Trim(Str(ws_TdKey)))
      ws_TdYyKey = ADC("TYy") + Space(8 - wf_TdKeyLen) + Trim(Str(ws_TdKey))  '578-34 space 7 changed to 8
      .Store "TdYyKey", ws_TdYyKey
      
      .Store "TdBYy", ""
      .Store "TdBChr", ""
      .Store "TdBNo", 0
      .Store "TdBQty", 0
      .Store "TdBDc", ""
      .Store "TdBGrWt", 0
      .Store "TdWrk", ""
      .Store "TdFrBLoc", ""
      .Store "TdToBLoc", ""
      '.Store "TdAck", ""
      .Store "TdPrdYN", ""
      .Store "TdByLoc", ""
      .Store "TdByWrk", ""
      '.Store "TdDustWt", 0
      .Store "TdRefYy", ""
      .Store "TdRefKey", 0
      .Store "TdRjCd", ""
      .Store "TdRjQty", 0
      .Store "TdRjDc", ""
      .Store "TdPtQty", 0
      .Store "TdRjPtQty", 0
      .Store "TdPrevYy", ""
      .Store "TdPrevKey", 0
    ' **** Zubin 211 **** '
      .Store "TdFBRmQty", ""
      .Store "TdFBRmWt", ""
    ' **** Zubin 211 **** '

    '****** Sachin 3.02 (14/09/07) Id Fields
      .Store "TdBIdNo", 0
      .Store "TdTIdNo", moCn.GetFldVal("Select TIdNo From Txn " + _
                        "Where TCoCd='" + ADC("TCoCd") + "' And TTc='" + ADC("TTc") + "' And TYy='" + ADC("TYy") + "' And Tchr='" + ADC("TChr") + "' And TNo=" + CStr(ADC("TNo")) + " And TPrtKey='" + ctCurrPrtn + "'")
      '.Store "TdDt", adc("TDt")
    '********* Manali 3.2.2 - Inserted user and date in Txnd Grids
      .Store "InsUsr", UCase(Trim(gs_UsrCd))
      .Store "InsDt", moCn.GetFldVal("Select getdate()")
      .Store "TdWaxWt", 0#
    End If
        
      .Store "TdDt", ADC("TDt")
      .Store "TdRmCtg", moCn.GetFldVal("Select RmCtg From RmMst Where RmCd='" + .Value(RowNum, "TdRmCd") + "' and RmPrtKey='" + ctCurrPrtn + "'")
      .Store "TdRmSCtg", moCn.GetFldVal("Select RmSCtg From RmMst Where RmCd='" + .Value(RowNum, "TdRmCd") + "' and RmPrtKey='" + ctCurrPrtn + "'")
      .Store "TdFrRmLocTyp", moCn.GetFldVal("Select LocTyp From Loc Where LocCoCd='" + ADC("TCoCd") + "' And LocCd='" + .Value(RowNum, "TdFrRmLoc") + "'")
      .Store "TdToRmLocTyp", moCn.GetFldVal("Select LocTyp From Loc Where LocCoCd='" + ADC("TCoCd") + "' And LocCd='" + .Value(RowNum, "TdToRmLoc") + "'")

    '****** Sachin 3.02 (14/09/07) Id Fields

    If GrdTxndA.Value(RowNum, "TdFrRmDc") = "C" Then
      .Store "TdToRmDc", "D"
    ElseIf GrdTxndA.Value(RowNum, "TdFrRmDc") = "D" Then
      .Store "TdToRmDc", "C"
    Else
      .Store "TdToRmDc", ""
    End If
    If UCase(ms_TcTyp) <> "MV" Then .Store "TdAck", "Y"
 End With
End Sub

Private Sub ADC_SetRecSource()
  '*** Set the RecSource for ADC (Table Txn)
  '****** Sachin 3.02 24-12-07 - Corresponding Partition
  ADC.RecSource = " Select * from Txn " + _
                  " Where TCoCd ='" + ADC("TCoCd") + "'" + _
                  " And TTc ='" + ADC("TTc") + "'" + _
                  " And TYy = '" + ADC("TYy") + "'" + _
                  " And TChr ='" + ADC("TChr") + "' " + _
                  " And TNo =" + CStr(ADC("TNo")) + " And TPrtKey=(Case When '" + gs_Partition + "'='" + ctCurrPrtn + "' Then '" + ctCurrPrtn + "' Else TPrtKey End)" + _
                  IIF(gs_Partition = ctCurrPrtn, " And Tdt>'" + ms_PurgeDt + "' ", "")
End Sub
Private Sub GrdTxndA_SetRecSource()
  '5715 existing check removed and now added TIdNo only
  Dim wTIdNo As Long
  wTIdNo = moCn.GetFldVal("Select TIdNo From Txn " + _
                        "Where TCoCd='" + ADC("TCoCd") + "' And TTc='" + ADC("TTc") + "' And TYy='" + ADC("TYy") + "' And Tchr='" + ADC("TChr") + "' And TNo=" + CStr(ADC("TNo")))
                        
  GrdTxndA.RecSource = " Select * from Txnd Where TdTIdNo =" + CStr(wTIdNo)
End Sub
Private Sub GrdTxndA_When()
  CmdLN.Enabled = True
End Sub

Private Sub GrdVS_ColWhen(ByVal RowNum As Integer, ByVal ColName As String, Cancel As Boolean, ErrMsg As String, ByVal Action As MwfCtl.en_fgAction)
  '*** No change allowed in the fields of the working grid GrdVS of the Voucher Summary frame
  Cancel = True
End Sub
Private Sub adc_AutoGen()
  '*** write code for AutoGenerating the Transaction Voucher Number
  If ADC("TNo") = 0 Then
    ADC("TNo") = moCn.AutoGen("Txn", "TNo", ADC("TCoCd") + ADC("TTc") + ADC("TYy") + ADC("TChr"))
  Else
    Call moCn.AutoGen("Txn", "TNo", ADC("TCoCd") + ADC("TTc") + ADC("TYy") + ADC("TChr"), ADC("TNo"))
  End If
End Sub
Private Sub CmdVS_Click()
  '*** Display or hide the Voucher Summary frame depending on the toggle state
      'and call the Sub Procedure VouchSumm in case the frame is visible
  
  Call DispFra(Summ)
  
  If FraVS.Visible = True Then
    If ADC("wVSGrp") = "" Then
      '*** when wVsGrp value changes the ColChange event calls the VouchSumm procedure
      ADC("wVSGrp") = "RMCTG"
    Else
      Call VouchSumm(ADC("wVSGrp"))
    End If
  End If
End Sub
Private Sub CmdCP_Click()
' ########################  Manoj #### Ver: 2.0.6 #### Date: 25/12/2003  ########################
  If ADC("TPrtKey") <> ctCurrPrtn Or mf_DtLocked = True Or UCase$(ADC("TLockYN")) = UCase$("Y") Then Exit Sub
'  If mf_DtLocked = True Then Exit Sub
' ###########################################  Manoj  ###########################################

  If mf_ValidLocSeq = False Then DispMsg "The User Does Not Have Access To One Of The Locs", etError: Exit Sub
  
  '*** Display or hide the Copy Frame depending on the toggle state
  If FraCP.Visible = False Then
    If wTcTyp <> "JT" Then OptRt(0).Enabled = False: OptRt(1).Enabled = False   'CopyData- Rate option only for jEmr
    Call DispFra(Cpy)
      '/************** Shilpa Emr205 **********
      'Manoj
        ADC("wCPTCoCd") = ADC("TCoCd")
        If ADC("wCpDetChk").Visible = True Then ADC("wCpDetChk") = "N" '*** Jay 3.1.1
        'adc("wCpDetChk") = "N"     '*** Bef Jay 3.1.1
      'Manoj
    '/************** Shilpa Emr205 **********
  Else
    Call DispFra(Cpy)
  End If
End Sub
Private Sub CmdCpGo_Click()
  '*** Call the Sub Procedure CpyRm
  If ms_TcTyp <> "CNV" Or ADC("wCpTTc") <> "CNV" Then
    If GrdTxndA.Rows > 1 Then DispMsg "Cannot Copy As Record Already Exists", etError: Exit Sub
  End If
  'CopyData- to get source Rmctg
  wCpGldAs = ""
  If ADC("WCPRMCTG") <> "" Then
    Dim wCpRmChr
    For wCpRmChr = 1 To Len(ADC("WCPRMCTG"))
      wCpGldAs = wCpGldAs + Mid(ADC("WCPRMCTG"), wCpRmChr, 1)
      If wCpRmChr <> Len(ADC("WCPRMCTG")) Then wCpGldAs = wCpGldAs + "','"
    Next
  Else
  wCpGldAs = "G','P','S','L','M','X','D','C"
  End If
  
  'CopyData- various source TC added
  If ms_TcTyp = "MV" Then
    'Or ms_TcTyp = "PR"
    If ADC("wCpTTc") <> "" And ADC("wCpTYy") <> "" And ADC("wCpTChr") <> "" And ADC("wCpTNo") > 0 And ADC("wCpGrpOpt") <> "" Then
        If wTcTyp = "JT" Then
          Call JtCopy
        ElseIf ADC("WCPTTC") = "TZ" Then
          Call CustzCopy
        Else
          Call CpyVchdet
        End If
    Else
      DispMsg "Enter The Voucher No and Group Option", etError: Exit Sub
    End If
  ElseIf ms_TcTyp = "PR" Or ms_TcTyp = "SA" Then
    If ADC("wCpTTc") <> "" And ADC("wCpTYy") <> "" And ADC("wCpTChr") <> "" And ADC("wCpTNo") > 0 And ADC("wCpTCoCd") <> "" Then
        If wTcTyp = "JT" Then
          Call JtCopy
        ElseIf ADC("WCPTTC") = "TZ" Then
          Call CustzCopy
        Else
          Call CpyMVRm
        End If
    Else
      DispMsg "Enter The Voucher No and Company code", etError: Exit Sub
    End If
  ElseIf ms_TcTyp = "CNV" And ADC("wCpDetChk") = "Y" Then
    If ADC("wCpTTc") <> "" And ADC("wCpTYy") <> "" And ADC("wCpTChr") <> "" And ADC("wCpTNo") > 0 And ADC("wCpTCoCd") <> "" Then
        If wTcTyp = "JT" Then
          Call JtCopy
        ElseIf ADC("WCPTTC") = "TZ" Then
          Call CustzCopy
        Else
          Call CpyRmDet
        End If
    Else
      DispMsg "Enter The Voucher No and Company code", etError: Exit Sub
    End If
  Else
    Call CpyRm
  End If
  Call CmdCP_Click
End Sub

'*** (Jen 2.13)
Private Sub CmdCpyRet_Click()

  '****** Sachin 3.02.0 - TPrtKey added below
  If ADC("TPrtKey") <> ctCurrPrtn Or mf_DtLocked = True Or UCase$(ADC("TLockYN")) = UCase$("Y") Then Exit Sub
  
  If ms_TTc <> "PR" Then Exit Sub

  If Not ADC.SaveRec Then
    DispMsg "X Save Failed", etInfo
    Exit Sub
  End If

  If mf_ValidLocSeq = False Then DispMsg "The User Does Not Have Access To One Of The Locs", etError: Exit Sub
  
  If FraCpyRet.Visible = False Then
    If ADC("wCrReYy") = "" Then ADC("wCrReYy") = ADC("TYy")
  End If
  '*** Display or hide the Return Copy Frame depending on the toggle state
  Call DispFra(CpyRet)
End Sub
Private Sub CmdCpyRetGo_Click()
  '*** Call the Sub Procedure CpyReturn
  
  Call CpyReturn
  
  Call CmdCpyRet_Click
End Sub
'*** (Jen 2.13)

'*** (Jen 2.14 Next)
Private Sub CmdCpyDt_Click()
  '****** Sachin 3.02.0 - TPrtKey added below
  If ADC("TPrtKey") <> ctCurrPrtn Or mf_DtLocked = True Or UCase$(ADC("TLockYN")) = UCase$("Y") Then Exit Sub
  
  If ms_TTc <> "PR" And ms_TTc <> "MV" Then Exit Sub

  If Not ADC.SaveRec Then
    DispMsg "X Save Failed", etInfo
    Exit Sub
  End If

  If mf_ValidLocSeq = False Then DispMsg "The User Does Not Have Access To One Of The Locs", etError: Exit Sub
  
  If FraCpyDt.Visible = False Then
    'If ADC("wDtCoCd") = "" Then
    OptVch(0).Value = True          '3.11.0
    If ms_TTc = "PR" Then ADC("wDtCoCd") = ADC("TCoCd")
    If ADC("wDtTc") = "" Then ADC("wDtTc") = "DT"
    If ADC("wDtYy") = "" Then ADC("wDtYy") = ADC("TYy")
  End If
  '*** Display or hide the DT Copy Frame depending on the toggle state
  Call DispFra(CpyDT)
End Sub
Private Sub CmdCpyDtGo_Click()
  Call CpyFromDT
  Call CmdCpyDt_Click
End Sub
'*** (Jen 2.14 Next)

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
Private Sub PropSet()
  '*** Setting Vis and Ena Properties for X Fields ***
  Dim wRsX As MwfLib.MDORowSet
'  If UCase(ADC.MenuCd) = "SA" Then
'    Set wRsX = moCn.OpenRes("Select * from TcStr where TsTbl='TxnRm' and " + _
                           "TsTyp='" + ms_TTc + "' and TsSTyp='X'")
'  Else
    Set wRsX = moCn.OpenRes("Select * from TcStr where TsTbl='TxnRm' and " + _
                           "TsTyp='" + ms_TcTyp + "' and TsSTyp='X'")
'  End If

  Do While Not (wRsX.EOF Or wRsX.BOF)
    ADC(wRsX!TsIdName).Enabled = FldEna(wRsX!TsStyle)
    ADC(wRsX!TsIdName).Visible = FldVis(wRsX!TsStyle)
    wRsX.MoveNext
  Loop
  
  '*** Setting Vis and Ena Properties for A Fields ***
  Call GrdPropSet(GrdTxndA, "A")

  '*** Setting Vis and Ena Properties for Label Conrols ***
  Dim wRsLbl As MwfLib.MDORowSet
  
 ' If UCase(ADC.MenuCd) = "SA" Then
 '   Set wRsLbl = moCn.OpenRes("Select * from TcStr where TsTbl='TxnRm' and " + _
                           "TsTyp='" + ms_TTc + "' and TsSTyp='Lbl'")
 ' Else
    Set wRsLbl = moCn.OpenRes("Select * from TcStr where TsTbl='TxnRm' and " + _
                           "TsTyp='" + ms_TcTyp + "' and TsSTyp='Lbl'")
 ' End If
  
  Dim wLblCtrl As Control
  Do While Not (wRsLbl.EOF Or wRsLbl.BOF)
    For Each wLblCtrl In Controls
      If TypeOf wLblCtrl Is Label Then
        If UCase(wLblCtrl.Name) = UCase(wRsLbl!TsIdName) Then
          wLblCtrl.Visible = FldVis(wRsLbl!TsStyle)
          If wRsLbl!TsLbl <> "" Then wLblCtrl.Caption = wRsLbl!TsLbl
'          MsgBox wLblCtrl.Name & " : " & wLblCtrl.Caption & " : " & wLblCtrl.Visible
        End If
      End If
    Next
    wRsLbl.MoveNext
  Loop
  
  '*** Setting Vis and Ena Properties for Frame Conrols ***
  Dim wRsFra As MwfLib.MDORowSet
  
 ' If UCase(ADC.MenuCd) = "SA" Then
 '   Set wRsFra = moCn.OpenRes("Select * from TcStr where TsTbl='TxnRm' and " + _
                             "TsTyp='" + ms_TTc + "' and TsSTyp='Fra'")
 ' Else
    Set wRsFra = moCn.OpenRes("Select * from TcStr where TsTbl='TxnRm' and " + _
                             "TsTyp='" + ms_TcTyp + "' and TsSTyp='Fra'")
 ' End If
  
  Dim wFraCtrl As Control
  Do While Not (wRsFra.EOF Or wRsFra.BOF)
    For Each wFraCtrl In Controls
      ' After ClrScr
      '*** (Jen 2.14 Next)
      If TypeOf wFraCtrl Is Frame Then
      '*** (Jen 2.14 Next)
      'If TypeOf wFraCtrl Is PictureBox Then    '*** (Bef 2.14 Next)
      ' After ClrScr
        If UCase(wFraCtrl.Name) = UCase(wRsFra!TsIdName) Then
          wFraCtrl.Visible = FldVis(wRsFra!TsStyle)
        End If
      End If
    Next
    wRsFra.MoveNext
  Loop
  
  '*** Setting Vis and Ena Properties for CommandButton Conrols ***
  Call CmdPropSet
  
  
  Set wRsX = Nothing
  Set wRsLbl = Nothing
  Set wRsFra = Nothing
End Sub
Private Sub GrdPropSet(ByVal mGrd As Object, ByVal mTsSTyp As String)
  '*** Procedure to set the Style of the columns in the grids
  
  Dim wGrdVis  As String
  With mGrd
    '*** Checking If The Grid Is Visible ***
  '  If UCase(ADC.MenuCd) = "SA" Then
  '    wGrdVis = moCn.GetFldVal("Select TsStyle from TcStr where TsTbl='TxnRm' and " + _
                              "TsTyp='" + ms_TTc + "' and TsSTyp= '" + mTsSTyp + "' and " + _
                              "TsIdName=''")
  '  Else
      wGrdVis = moCn.GetFldVal("Select TsStyle from TcStr where TsTbl='TxnRm' and " + _
                              "TsTyp='" + ms_TcTyp + "' and TsSTyp= '" + mTsSTyp + "' and " + _
                              "TsIdName=''")
  '  End If
    
    .Visible = FldVis(wGrdVis)

    '*** If The Grid Is Visible Then Set The Vis And Ena Properties For The Columns For
    '*** Which Record Exists in TcStructure Table
    If FldVis(wGrdVis) = True Then
      Dim wRsGrd As MwfLib.MDORowSet
   '   If UCase(ADC.MenuCd) = "SA" Then
   '     Set wRsGrd = moCn.OpenRes("Select * from TcStr where TsTbl='TxnRm' and " + _
                                   "TsTyp='" + ms_TTc + "' and TsSTyp= '" + mTsSTyp + "' and " + _
                                   "TsIdName<>''")
   '   Else
        Set wRsGrd = moCn.OpenRes("Select * from TcStr where TsTbl='TxnRm' and " + _
                                   "TsTyp='" + ms_TcTyp + "' and TsSTyp= '" + mTsSTyp + "' and " + _
                                   "TsIdName<>''")
   '   End If
      
      Do While Not (wRsGrd.EOF Or wRsGrd.BOF)
        .ColProp(wRsGrd!TsIdName).Style = ColSty(wRsGrd!TsStyle)
        wRsGrd.MoveNext
      Loop
    End If
  End With
  Set wRsGrd = Nothing
End Sub
Private Sub CmdPropSet()
  '*** Setting Vis and Ena Properties for CommandButton Conrols ***
  Dim wRsCmd As MwfLib.MDORowSet
  
 ' If UCase(ADC.MenuCd) = "SA" Then
 '   Set wRsCmd = moCn.OpenRes("Select * from TcStr where TsTbl='TxnRm' and " + _
                             "TsTyp='" + ms_TTc + "' and TsSTyp='Cmd'")
 ' Else
    Set wRsCmd = moCn.OpenRes("Select * from TcStr where TsTbl='TxnRm' and " + _
                             "TsTyp='" + ms_TcTyp + "' and TsSTyp='Cmd'")
 ' End If
  
  Dim wCmdCtrl As Control
  Do While Not (wRsCmd.EOF Or wRsCmd.BOF)
    For Each wCmdCtrl In Controls
      ' After ClrScr
      'If TypeOf wCmdCtrl Is CommandButton Then
      If TypeOf wCmdCtrl Is MWCTL_BTN1 Then
      ' After ClrScr
        If UCase(wCmdCtrl.Name) = UCase(wRsCmd!TsIdName) Then
          wCmdCtrl.Enabled = FldEna(wRsCmd!TsStyle)
          wCmdCtrl.Visible = FldVis(wRsCmd!TsStyle)
        End If
      End If
    Next
    wRsCmd.MoveNext
  Loop
  Set wRsCmd = Nothing
End Sub
Private Sub CpyRm()
  '*** Procedure to Copy the structure from another Voucher of the same TcType
      'in case of 'REC' or 'CNV' entry
      'Copying the structure means not copying the RmSz, StkRt, RmQty and RmWt
      'This copy helps the user in making the entry if he is not very well versed in
      'the way he has to proceed with the entry.
  
  Dim wTxndARow  As Single, wRow As Single
  wTxndARow = IIF(GrdTxndA.Row, GrdTxndA.Row, 1)
  Dim wRsCp As MwfLib.MDORowSet
 
  
 'CopyData- jEMr &custom purchase vouchers included
 If wTcTyp = "JT" Then
    If Not moCn.RecSeek("Select 'x' From OrdMst where " + _
       "OmCoCd= '" + ADC("wCpTCoCd") + "' and OmTc= '" + ADC("wCpTTc") + "' and " + _
      "OmYy= '" + ADC("wCpTYy") + "' and OmChr= '" + ADC("wCpTChr") + "' and " + _
      "OmNo= '" + CStr(ADC("wCpTNo")) + "' and OmPrtKey='" + ctCurrPrtn + "'") Then DispMsg "Invalid Order Voucher", etError: Exit Sub

    If Not moCn.RecSeek("Select 'x' From OrdMst where " + _
      "OmCoCd= '" + ADC("wCpTCoCd") + "' and OmTc= '" + ADC("wCpTTc") + "' and " + _
      "OmYy= '" + ADC("wCpTYy") + "' and OmChr= '" + ADC("wCpTChr") + "' and " + _
      "OmNo= '" + CStr(ADC("wCpTNo")) + "' and OmPrtKey='" + ctCurrPrtn + "' and OmCmCurCd ='" + ADC("TCurrCd") + "'") Then DispMsg "Source && Target Currency Does Not Match", etError: Exit Sub
    
    If wLooseTc = "Y" Then _
      Set wRsCp = moCn.OpenRes(" Select " + IIF(UCase(ADC("wCpGrpOpt")) = "SR", "max(OdJLRmCd)", "OdJLRmCd") + " as TdRmCd from OrdDsg " + _
                            " Join RmMst on RmCd=OdJLRmCd and RmCtg in ('" + wCpGldAs + "')" + _
                            " Where OdJLRmCd <> '' and OdCoCd= '" + ADC("wCpTCoCd") + "' and " + _
                            " OdTc= '" + ADC("wCpTTc") + "' and OdYy= '" + ADC("wCpTYy") + "' and " + _
                            " OdChr= '" + ADC("wCpTChr") + "' and OdNo= '" + CStr(ADC("wCpTNo")) + "' and " + _
                            " OdPrtKey='" + ctCurrPrtn + "' Group by " + IIF(UCase(ADC("wCpGrpOpt")) = "SR", "OdSr", "OdJLRmCd"))
  
    If wLooseTc = "N" Then _
      Set wRsCp = moCn.OpenRes(" Select " + IIF(UCase(ADC("wCpGrpOpt")) = "SR", "max(OrRmCd)", "OrRmCd") + " as TdRmCd " + _
                            " from OrdRm Where OrRmCd <> '' and OrRmCtg in ('" + wCpGldAs + "') and " + _
                            " OrCoCd= '" + ADC("wCpTCoCd") + "' and  OrTc= '" + ADC("wCpTTc") + "' and " + _
                            " OrYy= '" + ADC("wCpTYy") + "' and OrChr= '" + ADC("wCpTChr") + "' and " + _
                            " OrNo= '" + CStr(ADC("wCpTNo")) + "' and OrPrtKey='" + ctCurrPrtn + "'" + _
                            " Group by " + IIF(UCase(ADC("wCpGrpOpt")) = "SR", "OrSr,OrSrNo", "OrRmCd"))
  
 Else
   If Not moCn.RecSeek("Select 'x' From Txn where " + _
     "TCoCd= '" + ADC("wCpTCoCd") + "' and TTc= '" + ADC("wCpTTc") + "' and " + _
     "TYy= '" + ADC("wCpTYy") + "' and TChr= '" + ADC("wCpTChr") + "' and " + _
     "TNo= '" + CStr(ADC("wCpTNo")) + "' and TPrtKey='" + ctCurrPrtn + "' and TCurrCd ='" + ADC("TCurrCd") + "'") Then DispMsg "Source && Target Currency Does Not Match", etError: Exit Sub
    
   If ADC("wCpTTc") = "TZ" Then
     Set wRsCp = moCn.OpenRes(" Select " + IIF(UCase(ADC("wCpGrpOpt")) = "SR", "max(TzRmCdDesc)", "TzRmCdDesc") + " as TdRmCd " + _
                           " From Txndz Where TzRmCdDesc <> '' and TzRmCtg in ('" + wCpGldAs + "') and " + _
                           "TzCoCd= '" + ADC("wCpTCoCd") + "' and TzTc= '" + ADC("wCpTTc") + "' and " + _
                           "TzYy= '" + ADC("wCpTYy") + "' and TzChr= '" + ADC("wCpTChr") + "' and " + _
                           "TzNo= " + CStr(ADC("wCpTNo")) + " and TzPrtKey='" + ctCurrPrtn + "'" + _
                           " Group by " + IIF(UCase(ADC("wCpGrpOpt")) = "SR", "TzSr", "TzRmCdDesc"))
   Else
  '****** Sachin 2.13.0 - [17-07-2006] - Multi-Currency ******
  '****** Sachin 3.02.0 24-12-07 - Current Partition
    '*** Jay 3.1.1 [adc("TCoCd") replaced with adc("wCpTCoCd")]
   ' Set wRsCp = moCn.OpenRes("Select TdFrRmLoc, TdFrRmDc, TdRmCd, TdToRmLoc "
     Set wRsCp = moCn.OpenRes("Select TdFrRmLoc, TdFrRmDc, TdRmCd, TdToRmLoc " + _
                           " From Txnd Where TdRmCd <> '' and TdRmCtg in ('" + wCpGldAs + "') and " + _
                           "TdCoCd= '" + ADC("wCpTCoCd") + "' and TdTc= '" + ADC("wCpTTc") + "' and " + _
                           "TdYy= '" + ADC("wCpTYy") + "' and TdChr= '" + ADC("wCpTChr") + "' and " + _
                           "TdNo= " + CStr(ADC("wCpTNo")) + " and TdPrtKey='" + ctCurrPrtn + "'")
   End If
 End If

  If wRsCp.RecCount <= 0 Then DispMsg "No Record Exists", etError: Exit Sub
  If ADC.SaveRec = False Then Exit Sub
  GrdTxndA.StartCopy
  MWLib.BeginProcess Me, "Copying The Structure ..."
  With GrdTxndA
    Do While Not (wRsCp.EOF Or wRsCp.BOF)
      .AddItem
      wRow = .Rows - 1
      .Row = wRow
      ' Manali 3.6.0
      If .Value(wRow, "TdSr") = 0 Then .Value(wRow, "TdSr") = .MaxVal("TdSr") + 1
      .Value(wRow, "TdSr") = wRow
      .Value(wRow, "TdFrRmLoc") = wRsCp!TdFrRmLoc              '### ADC("TFrRmLoc")
      '.Value(wRow, "TdFrRmLoc") = ADC("TFrRmLoc")
      If UCase(ADC("wCpTTc")) = "CNV" Then
        .Value(wRow, "TdFrRmDc") = wRsCp!TdFrRmDc
        If wRsCp!TdFrRmDc = "C" Then
          .Value(wRow, "wTdFrRmIr") = "I"
        ElseIf wRsCp!TdFrRmDc = "D" Then
          .Value(wRow, "wTdFrRmIr") = "R"
        End If
      Else
        .Value(wRow, "wTdFrRmIr") = ADC("WFrRmIR")
        If ADC("WFrRmIR") = "R" Then
          .Value(wRow, "TdFrRmDc") = "D"
        ElseIf ADC("WFrRmIR") = "I" Then
          .Value(wRow, "TdFrRmDc") = "C"
        End If
      End If
      .Value(wRow, "TdRmCd") = wRsCp!TdRmCd
' Zubin 212
      .Value(wRow, "TdLotNo") = ""
' Zubin 212
      .Value(wRow, "TdRmSz") = 0
      .Value(wRow, "TdRmSz2") = 0
      .Value(wRow, "TdRmSz3") = 0
      .Value(wRow, "TdRmStkRt") = 0
      .Value(wRow, "TdRmQty") = 0
      .Value(wRow, "TdRmWt") = 0
      .Value(wRow, "TdToRmLoc") = wRsCp!TdToRmLoc             '### ADC("TToRmLoc")
      '.Value(wRow, "TdToRmLoc") = ADC("TToRmLoc")
      .Value(wRow, "TdDesc") = ""
      .Value(wRow, "TdPurRt") = 0
      .Value(wRow, "TdPurAmt") = 0
      'Manoj
      .Value(wRow, "TdDustWt") = 0
      .Value(wRow, "TdAck") = "Y"
      'Manoj
      .Value(wRow, "TdPrtKey") = ctCurrPrtn   '****** Sachin 3.02 24-12-07
      .SaveRec (wRow)
      wRsCp.MoveNext
    Loop
  End With
  MWLib.EndProcess Me
  GrdTxndA.EndCopy
  Set wRsCp = Nothing
  '*** End of Copying the structure
End Sub
'**********************************Geeta***********************Emr206**********************

Private Sub GenVchdet()
Dim wValnSql As String   '3.11.0
Dim wUnion As String   '3.11.0
Dim wStkValSql As String   '3.11.0
Dim wFinalSql As String
Dim wo_rsSrcSYM As MDORowSet
Dim wCnd As String
Dim wRow As Integer
If ADC.SaveRec = True Then
wCnd = ADC.RepCond
wCnd = IIF(wCnd <> "", " and " + wCnd, "")
 
' Zubin 212 (SLotNo added)
' Manoj 2.9.0 Patch
'****** Sachin 3.02.0 24-12-07 - Current Partition

wValnSql = " Select 1 as qTag ,SCoCd, SLoc, SRmCd, SLotNo, SRmSz,SRmSz2,SRmSz3,sum(SRmDrQty-SRmDrQty) as qStkQty,Convert(decimal(10,3),sum((SRmDrWt-SRmCrWt)*SRmStkRt)) as qStkVal, " + _
                              " Convert(decimal(10,3),sum(SRmDrWt-SRmCrWt)) as qStkWt,max(RrCstRt) as RrCstRt from SYyMm " + _
                              " join RmRt " + _
                              " on RrTcTyp='RM' and RrCmCtg='C' and RrCmCd='ZValue' and RrCd=SRmCd and RrPrtKey=SPrtKey and " + _
                              " convert(Decimal(10,4),RrFrLn)<=convert(Decimal(10,4),SRmSz) and " + _
                              " convert(Decimal(10,4),RrToLn) >=convert(Decimal(10,4),SRmSz) " + _
                              " join Loc on LocCoCd= '" + ADC("TCoCd") + "' and LocCd= SLoc and LocTyp = 'S' " + _
                              " join RmMst on RmCd=SRmCd and RmPrtKey=SPrtKey Where SCoCd= '" + ADC("TCOCd") + "' and SPrtKey='" + ctCurrPrtn + "' and (Select PValue1 from Param where  " + _
                              " PTyp='RMSCTG' and PMCD=RMCTG and PSCd=RmSCtg)='Y' and " + _
                              " (select LocValidYN From Loc where LocCd=SLoc and LocCoCd=SCoCd)='Y' " + wCnd + "   group by SCoCd,SLoc,SRmCd,SLotNo,SRmSz,SRmSz2,SRmSz3 " + _
                              " Having Convert(decimal(10,3),sum(SRmDrWt-SRmCrWt))>0  and Convert(Decimal(10,2),Convert(decimal(10,3),sum((SRmDrWt-SRmCrWt)*SRmStkRt))/  " + _
                              " (case when Convert(decimal(10,3),sum(SRmDrWt-SRmCrWt))=0 then 1 else Convert(decimal(10,3),sum(SRmDrWt-SRmCrWt)) end )) <>convert(decimal(10,2),max(RrCstRt)) "
wUnion = " Union All "
wStkValSql = " Select 2 as qTag ,SCoCd, SLoc, SRmCd, SLotNo, SRmSz,SRmSz2,SRmSz3,sum(SRmDrQty-SRmDrQty) as qStkQty, " + _
                              " Convert(decimal(10,3),sum((SRmDrWt-SRmCrWt)*SRmStkRt)) as qStkVal, " + _
                              " 1 as qStkWt,0 as RrCstRt from SYyMm " + _
                               "join Loc on LocCoCd= '" + ADC("TCoCd") + "' and LocCd= SLoc and LocTyp = 'S' " + _
                              "join RmMst on RmCd=SRmCd and RmPrtKey=SPrtKey Where SCoCd= '" + ADC("TCOCd") + "' and SPrtKey='" + ctCurrPrtn + "' and (Select PValue1 from Param where  " + _
                              "PTyp='RMSCTG' and PMCD=RMCTG and PSCd=RmSCtg)='Y' and " + _
                              "(select LocValidYN From Loc where LocCd=SLoc and LocCoCd=SCoCd)='Y' " + wCnd + "   group by SCoCd,SLoc,SRmCd,SLotNo,SRmSz,SRmSz2,SRmSz3 " + _
                              " Having Convert(decimal(16, 3), Sum(SRmDrWt - SRmCrWt)) = 0 And Convert(decimal(16, 3), Sum(SRmDrQty - SRmCrQty)) = 0 " + _
                              " and Convert(decimal(10,2),sum((SRmDrWt-SRmCrWt)*SRmStkRt))<>0 " + _
                              "Order by qTag, SCoCd,SLoc,SRmCd,SLotNo,SRmSz,SRmSz2,SRmSz3 "

'3.11.0
If ADC("WVALYN") = "Y" And ADC("WREMSTKVALYN") = "N" Then
    Set wo_rsSrcSYM = moCn.OpenRes(wValnSql)
ElseIf ADC("WVALYN") = "N" And ADC("WREMSTKVALYN") = "Y" Then
    Set wo_rsSrcSYM = moCn.OpenRes(wStkValSql)
Else
    Set wo_rsSrcSYM = moCn.OpenRes(wValnSql + wUnion + wStkValSql)
End If

'Set wo_rsSrcSYM = moCn.OpenRes("select SCoCd, SLoc, SRmCd, SRmSz,sum(SRmDrQty-SRmDrQty) as qStkQty,Convert(decimal(10,3),sum((SRmDrWt-SRmCrWt)*SRmStkRt)) as qStkVal, " + _
'                              "Convert(decimal(10,3),sum(SRmDrWt-SRmCrWt)) as qStkWt,max(RrCstRt) as RrCstRt from SYyMm " + _
'                              "join RmRt " + _
'                              "on RrTcTyp='RM' and RrCmCtg='C' and RrCmCd='ZValue' and RrCd=SRmCd and " + _
'                              "convert(Decimal(10,3),RrFrLn)<=convert(Decimal(10,3),SRmSz) and " + _
'                              "convert(Decimal(10,3),RrToLn) >=convert(Decimal(10,3),SRmSz) " + _
'                              "join Loc on LocCoCd= '" + ADC("TCoCd") + "' and LocCd= SLoc and LocTyp in " + ctAvgLocTypSeq + " " + _
'                              "join RmMst on RmCd=SRmCd  where SCoCd=  '" + ADC("TCOCd") + "' and (Select PValue1 from Param where  " + _
'                              "PTyp='RMSCTG' and PMCD=RMCTG and PSCd=RmSCtg)='Y' and " + _
'                              "(select LocValidYN From Loc where LocCd=SLoc and LocCoCd=SCoCd)='Y' " + wCnd + "   group by SCoCd,SLoc,SRmCd,SRmSz  " + _
'                              "Having Convert(decimal(10,3),sum(SRmDrWt-SRmCrWt))>=0  and  Convert(decimal(10,3),sum((SRmDrWt-SRmCrWt)*SRmStkRt))/  " + _
'                              "(case when Convert(decimal(10,3),sum(SRmDrWt-SRmCrWt))=0 then 1 else Convert(decimal(10,3),sum(SRmDrWt-SRmCrWt)) end ) <>max(RrCstRt) " + _
'                              "Order by SCoCd,SLoc,SRmCd,SRmSz ")
' Manoj 2.9.0 Patch

  If wo_rsSrcSYM.RecCount = 0 Then
    DispMsg "No Records found for this scope in Stock Rate Table/Invalid Location /Rm SubCategory is not Stock By Rate ", etError
    Exit Sub
  ElseIf wo_rsSrcSYM.RecCount > 499 Then
    DispMsg "More Than 998 records selected ,Reduce the Scope", etError
    Exit Sub
  End If
    
'  If GrdTxndA.Rows >= 2 Then
'    DispMsg "Can Not Generate Vouchers As Records Already Exist in the Grid", etError
'    Exit Sub
'  End If
MWLib.BeginProcess Me, "Generating Vouchers"
 With wo_rsSrcSYM
    GrdTxndA.StartCopy
    Do While Not (.EOF Or .BOF)
    MWLib.RunningProcess "Generating Vouchers  for Sr " + CStr(GrdTxndA.MaxVal("TdSr") + 1)
      GrdTxndA.AddItem
      wRow = GrdTxndA.Rows - 1
      GrdTxndA.Row = wRow
'      GrdTxndA.Value(wRow, "TdSr") = GrdTxndA.MaxVal("TdSr") + 1
      GrdTxndA.Value(wRow, "TdFrRmLoc") = !SLoc
      GrdTxndA.Value(wRow, "wTdFrRmIR") = "I"
      GrdTxndA.Value(wRow, "TdFrRmDC") = "C"
      GrdTxndA.Value(wRow, "TDRMCD") = !SRmCd
      ' Zubin 212
      GrdTxndA.Value(wRow, "TDLotNo") = !SLotNo
      ' Zubin 212
      GrdTxndA.Value(wRow, "TDRMSZ") = !SRmSz
      GrdTxndA.Value(wRow, "TDRMSZ2") = !SRmSz2
      GrdTxndA.Value(wRow, "TDRMSZ3") = !SRmSz3
      GrdTxndA.Value(wRow, "TDRMSTKRT") = !qStkVal / IIF(!qStkWt = 0, 1, !qStkWt)
      GrdTxndA.Value(wRow, "TDPCPERCT") = 0
      GrdTxndA.Value(wRow, "TDRMQTY") = !qStkQty
      GrdTxndA.Value(wRow, "TDRMWT") = !qStkWt
      GrdTxndA.Value(wRow, "TDTORMLOC") = "XC"
     ' GrdTxndA.Value(wRow, "TDPURRT") = GrdTxndA.Value(wRow, "TDRMSTKRT")
     ' GrdTxndA.Value(wRow, "TDPURAMT") = GrdTxndA.Value(wRow, "TDRMSTKRT") * GrdTxndA.Value(wRow, "TDRMWT")
      If GrdTxndA.SaveRec(wRow) = False Then GoTo NextRec
      
      MWLib.RunningProcess "Generating Vouchers  for Sr " + CStr(GrdTxndA.MaxVal("TdSr") + 1)
      GrdTxndA.AddItem
      wRow = GrdTxndA.Rows - 1
      GrdTxndA.Row = wRow
'      GrdTxndA.Value(wRow, "TdSr") = GrdTxndA.MaxVal("TdSr") + 1
      GrdTxndA.Value(wRow, "TdFRRMLOC") = !SLoc
      GrdTxndA.Value(wRow, "WTDFRRMIR") = "R"
      GrdTxndA.Value(wRow, "TDFRRMDC") = "D"
      GrdTxndA.Value(wRow, "TDRMCD") = !SRmCd
      ' Zubin 212
      GrdTxndA.Value(wRow, "TDLotNo") = !SLotNo
      ' Zubin 212
      GrdTxndA.Value(wRow, "TDRMSZ") = !SRmSz
      GrdTxndA.Value(wRow, "TDRMSZ2") = !SRmSz2
      GrdTxndA.Value(wRow, "TDRMSZ3") = !SRmSz3
      GrdTxndA.Value(wRow, "TDRMSTKRT") = !RrCstRt
      GrdTxndA.Value(wRow, "TDPCPERCT") = 0
      GrdTxndA.Value(wRow, "TDRMQTY") = !qStkQty
      GrdTxndA.Value(wRow, "TDRMWT") = !qStkWt
      GrdTxndA.Value(wRow, "TDTORMLOC") = "XC"
     ' GrdTxndA.Value(wRow, "TDPURRT") = GrdTxndA.Value(wRow, "TDRMSTKRT")
     ' GrdTxndA.Value(wRow, "TDPURAMT") = GrdTxndA.Value(wRow, "TDRMSTKRT") * GrdTxndA.Value(wRow, "TDRMWT")
     'Manoj
     GrdTxndA.Value(wRow, "TdDustWt") = 0
     GrdTxndA.Value(wRow, "TdPrtKey") = ctCurrPrtn    '****** Sachin 3.02 24-12-07
     GrdTxndA.Value(wRow, "TdAck") = "Y"
     'Manoj
      GrdTxndA.SaveRec (wRow)
NextRec:
      .MoveNext
    Loop
    GrdTxndA.EndCopy
  End With
 Else
  DispMsg "X Save Failed", etError
 End If
MWLib.EndProcess Me
End Sub
Private Sub CpyVchdet()
  Dim wTxndARow  As Single, wRow As Single
  Dim wTxndFlds As String, wGroupBy As String, wTdRmStkRtYn As String
  Dim wRsCp As MwfLib.MDORowSet
  
  wTxndARow = IIF(GrdTxndA.Row, GrdTxndA.Row, 1)
  
  If Not moCn.RecSeek("Select 'x' From Txn " + _
                           "Where TCoCd= '" + ADC("wCpTCoCd") + "' and " + _
                           "TTc= '" + ADC("wCpTTc") + "' and " + _
                           "TYy= '" + ADC("wCpTYy") + "' and " + _
                           "TChr= '" + ADC("wCpTChr") + "' and " + _
                           "TNo= " + CStr(ADC("wCpTNo")) + " And TCurrCd ='" + ADC("TCurrCd") + "' And TPrtKey='" + ctCurrPrtn + "' ") Then
    DispMsg "Currency Mismatch in Source And Target Vouchers", etError: Exit Sub
  End If
  
  ' Zubin 212 (TdLotNo added)
  If UCase(ADC("wCpGrpOpt")) = "RMCDRT" Then
    wTxndFlds = "TdRmCd, TdRmStkRt, 0 as TdRmSz, 0 as TdRmSz2,0 as TdRmSz3, '' as TdLotNo,  TdPurRt, TdRtByQw, Sum(TdPurAmt)  "
    wGroupBy = "TdRmCd, TdRmStkRt, TdPurRt, TdRtByQw "
  ElseIf UCase(ADC("wCpGrpOpt")) = "RMCDSZ" Then
    wTxndFlds = "TdRmCd, 0 as TdRmStkRt, TdRmSz, TdRmSz2, TdRmSz3, '' as TdLotNo, 0 as TdPurRt, max(RmQw) as TdRtByQw, 0 as TdPurAmt "
    wGroupBy = "TdRmCd, TdRmSz, TdRmSz2, TdRmSz3"
  ElseIf UCase(ADC("wCpGrpOpt")) = "SR" Then
    wTxndFlds = "max(TdRmCd) as TdRmCd, max(TdRmStkRt) as TdRmStkRt, max(TdRmSz) as TdRmSz, max(TdRmSz2) as TdRmSz2, max(TdRmSz3) as TdRmSz3, max(TdLotNo) as TdLotNo, max(TdPurRt) as TdPurRt, max(TdRtByQw) as TdRtByQw, sum(TdPurAmt) as TdPurAmt "
    wGroupBy = "TdSr "
  ElseIf UCase(ADC("wCpGrpOpt")) = "RMCDSZRT" Then
    wTxndFlds = "TdRmCd, TdRmStkRt, TdRmSz,TdRmSz2, TdRmSz3, '' as TdLotNo, TdPurRt, TdRtByQw, sum(TdPurAmt) "
    wGroupBy = "TdRmCd,TdRmSz,TdRmSz2, TdRmSz3,TdRmStkRt,TdPurRt,TdRtByQw "
  ' Zubin 212
  ElseIf UCase(ADC("wCpGrpOpt")) = "CDRTLT" Then
    wTxndFlds = "TdRmCd, TdRmStkRt, TdLotNo, 0 as TdRmSz,0 as TdRmSz2,0 as TdRmSz3, TdPurRt,TdRtByQw, Sum(TdPurAmt) "
    wGroupBy = "TdRmCd, TdRmStkRt, TdPurRt, TdLotNo,TdRtByQw "
  ElseIf UCase(ADC("wCpGrpOpt")) = "CDSZLT" Then
    wTxndFlds = "TdRmCd, TdRmSz, TdRmSz2, TdRmSz3, TdLotNo, 0 as TdRmStkRt, 0 as TdPurRt, max(RmQw) as TdRtByQw, 0 as TdPurAmt "
    wGroupBy = "TdRmCd, TdRmSz, TdRmSz2, TdRmSz3, TdLotNo "
  ElseIf UCase(ADC("wCpGrpOpt")) = "CDSZRTLT" Then
    wTxndFlds = "TdRmCd, TdRmSz,  TdRmSz2, TdRmSz3, TdRmStkRt, TdLotNo, TdPurRt, TdRtByQw, sum(TdPurAmt) "
    wGroupBy = "TdRmCd, TdRmSz, TdRmSz2,  TdRmSz3,TdRmStkRt,TdPurRt, TdLotNo, TdRtByQw "
  End If
     
  Set wRsCp = moCn.OpenRes("Select " + wTxndFlds + ", Sum(TdRmQty) as TotQty, " + _
                           "Sum(TdRmWt) as TotWt From Txnd join Rmmst on RmCd = TdRmCd " + _
                           "Where TdRmCd <> '' and TdRmCtg in ('" + wCpGldAs + "') and " + _
                           "TdCoCd= '" + ADC("wCpTCoCd") + "' and " + _
                           "TdTc= '" + ADC("wCpTTc") + "' and " + _
                           "TdYy= '" + ADC("wCpTYy") + "' and " + _
                           "TdChr= '" + ADC("wCpTChr") + "' and " + _
                           "TdNo= " + CStr(ADC("wCpTNo")) + " and TdPrtKey='" + ctCurrPrtn + "'" + _
                           "Group by " + wGroupBy)
  If wRsCp.RecCount <= 0 Then DispMsg "No Record Exists", etError: Exit Sub
  If ADC.SaveRec = False Then Exit Sub
  GrdTxndA.StartCopy
  MWLib.BeginProcess Me, "Copying The Voucher Details ..."
  With GrdTxndA
    Do While Not (wRsCp.EOF Or wRsCp.BOF)
      .AddItem
      wRow = .Rows - 1
      .Row = wRow
      If .Value(wRow, "TdSr") = 0 Then .Value(wRow, "TdSr") = .MaxVal("TdSr") + 1
      .Value(wRow, "TdFrRmLoc") = ADC("TFrRmLoc")
      .Value(wRow, "wTdFrRmIr") = ADC("wFrRmIr")
      If ADC("wFrRmIr") = "R" Then
        .Value(wRow, "TdFrRmDc") = "D"
      Else
        .Value(wRow, "TdFrRmDc") = "C"
      End If
      .Value(wRow, "TdToRmLoc") = ADC("TToRmLoc")
      
      .Value(wRow, "TdRmCd") = wRsCp!TdRmCd
      .Value(wRow, "TdLotNo") = wRsCp!TdLotNo
      .Value(wRow, "TdRmSz") = wRsCp!TdRmSz
      .Value(wRow, "TdRmSz2") = wRsCp!TdRmSz2
      .Value(wRow, "TdRmSz3") = wRsCp!TdRmSz3
      .Value(wRow, "TdPurRt") = wRsCp!TdPurRt
      .Value(wRow, "TdRtByQw") = wRsCp!TdRtByQW
       wTdRmStkRtYn = moCn.GetFldVal("Select PValue1 from Param join RmMst on RmCd= '" + _
                 .Value(wRow, "TdRmCd") + "' and RmPrtKey='" + ctCurrPrtn + "' where PTyp= 'RmSCtg' and PmCd=RmCtg and " + _
                 " PScd = RmSCtg ")
                 
      If wTdRmStkRtYn = "Y" Then
          If moCn.GetFldVal("Select hAvgSRtYn from Head where HCoCd= '" + ADC("TCoCd") + "' and " + _
                            "HCd='" + ctSelfCmCd + "'") = "Y" And ms_TTc <> "PR" Then
          '*** (Jay 3.1.1)
          Dim wRmLocTyp As String, wb_RmLocTyp As Boolean
            wb_RmLocTyp = moCn.RecSeek("Select LocTyp from Loc where LocCoCd= '" + ADC("TCoCd") + "' " + _
                                       "and LocCd= '" + ADC("TFrRmLoc") + "' " + _
                                       "and LocTyp in " + ctAvgLocTypSeq)
            'Calculate Avg Rt
            wRmLocTyp = Trim(moCn.GetFldVal("Select LocTyp from Loc where LocCoCd= '" + ADC("TCoCd") + "' " + _
                                       "and LocCd= '" + ADC("TFrRmLoc") + "' "))

            If wb_RmLocTyp = True Then
              ' Zubin 212
              .Value(wRow, "TdRmStkRt") = GetAvgStkRt(ADC("TCoCd"), ADC("TFrRmLoc"), _
                                          wRsCp!TdRmCd, wRsCp!TdLotNo, wRsCp!TdRmSz, wRsCp!TdRmSz2, wRsCp!TdRmSz3)
            Else
              .Value(wRow, "TdRmStkRt") = wRsCp!TDRMSTKRT
            End If
          Else
            .Value(wRow, "TdRmStkRt") = wRsCp!TDRMSTKRT
          End If
      Else
        .Value(wRow, "TdRmStkRt") = 0
      End If
    '*************************************************
      .Value(wRow, "TdRmQty") = wRsCp!TotQty
      .Value(wRow, "TdRmWt") = wRsCp!TotWt
      .Value(wRow, "TdDesc") = ADC("WCPTCOCD") + "/" + ADC("WCPTTC") + "/" + ADC("WCPTYY") + "/" + ADC("WCPTCHR") + "/" + CStr(ADC("WCPTNO"))
      .Value(wRow, "TdDustWt") = 0
      .Value(wRow, "TdAck") = ADC("TAck")
      .Value(wRow, "TdPrtKey") = ctCurrPrtn
      .SaveRec (wRow)
      wRsCp.MoveNext
    Loop
  End With
  MWLib.EndProcess Me
  GrdTxndA.EndCopy
  Set wRsCp = Nothing
End Sub

Private Sub GenContra()
  
  '****** Sachin 2.14.0 [Multi-Company]
  '****** Procedure to Generate the Contra Voucher - From One Company to Other
  '****** Both the Source And Target Companies should have same Base Currency Codes
  '****** Contra can only be done between two different companies other than 'ZZZ'
  
  Dim wTxndARow  As Single, wRow As Single
  Dim wTxndFlds As String, wGroupBy As String, wTdRmStkRtYn As String
  Dim wRsCp As MwfLib.MDORowSet, wBaseCurMatch As Boolean
  
  wTxndARow = IIF(GrdTxndA.Row, GrdTxndA.Row, 1)
  
  '****** Base Currencies of Both the Companies should match to do this entry
  wBaseCurMatch = IIF(Not moCn.GetFldVal("Select HBaseCurCd From Head  Where HCd='" + ctSelfCmCd + "' And HCoCd='" + _
                   ADC("TCoCd") + "'") = moCn.GetFldVal("Select HBaseCurCd From Head  Where HCd='" + ctSelfCmCd + "' And HCoCd='" + _
                   ADC("wCnCoCd") + "'"), False, True)
                    
  Set wRsCp = moCn.OpenRes("Select qLoc, TdRmCd, TdRmSz, TdRmSz2, TdRmSz3, TdRmStkRt, Sum(RmWt) as RmWt, Sum(RmQty) as RmQty " + _
              " From " + _
              " (Select TdFrRmLoc qLoc, TdRmCd, TdRmSz,TdRmSz2,TdRmSz3, TdRmStkRt, " + _
              "   Sum(TdRmWt*(Case When LFr.LocTyp = 'XC' And TdFrRmDC='D' Then 1 " + _
              "                    When LFr.LocTyp = 'XC' And TdFrRmDC='C' Then -1 End)) RmWt, " + _
              "   Sum(TdRmQty*(Case When LFr.LocTyp = 'XC' And TdFrRmDC='D' Then 1 " + _
              "                     When LFr.LocTyp = 'XC' And TdFrRmDC='C' Then -1 End)) RmQty " + _
              "    From Txnd " + _
              " Join Loc LFr On LFr.LocCoCd=TdCoCd And LFr.LocCd=TdFrRmLoc " + _
              " Where TdCoCd='" + ADC("wCnCoCd") + "' And TdTc='" + ADC("wCnTTc") + "' And TdYy='" + ADC("wCnTYy") + "' And TdChr='" + ADC("wCnTChr") + "' And TdNo=" + CStr(ADC("wCnTNo")) + " And TdPrtKey='" + ctCurrPrtn + "' And LFr.LocTyp='XC' " + _
              " Group By TdCoCd, TdTc, TdYy, TdChr, TdNo, TdFrRmLoc,TdRmCd, TdRmSz,TdRmSz2,TdRmSz3, TdRmStkRt " + _
              " Union " + _
              " Select TdToRmLoc qLoc, TdRmCd, TdRmSz,TdRmSz2,TdRmSz3, TdRmStkRt, " + _
              "   Sum(TdRmWt*(Case When LTo.LocTyp = 'XC' And TdToRmDC='D' Then 1 " + _
              "                    When LTo.LocTyp = 'XC' And TdToRmDC='C' Then -1 End)) RmWt, " + _
              "   Sum(TdRmQty*(Case When LTo.LocTyp = 'XC' And TdToRmDC='D' Then 1 " + _
              "                     When LTo.LocTyp = 'XC' And TdToRmDC='C' Then -1 End)) RmQty " + _
              " From Txnd " + _
              "   Join Loc LTo On LTo.LocCoCd=TdCoCd And LTo.LocCd=TdToRmLoc " + _
              " Where TdCoCd='" + ADC("wCnCoCd") + "' And TdTc='" + ADC("wCnTTc") + "' And TdYy='" + ADC("wCnTYy") + "' And TdChr='" + ADC("wCnTChr") + "' And TdNo=" + CStr(ADC("wCnTNo")) + " And TdPrtKey='" + ctCurrPrtn + "' And LTo.LocTyp='XC' " + _
              " Group By TdCoCd, TdTc, TdYy, TdChr, TdNo, TdToRmLoc,TdRmCd, TdRmSz,TdRmSz2,TdRmSz3, TdRmStkRt) TmpTab " + _
              "  Where TdRmCd<> '' and RmWt > 0 " + _
              "  Group By qLoc, TdRmCd, TdRmSz,TdRmSz2,TdRmSz3, TdRmStkRt " + _
              " Having Sum(rmwt) > 0 ")
        
  If wRsCp.RecCount <= 0 Then DispMsg "No Record Exists", etError: Exit Sub
  If ADC.SaveRec = False Then Exit Sub
    GrdTxndA.StartCopy
    MWLib.BeginProcess Me, "Copying The Voucher Details ..."
  With GrdTxndA
    Do While Not (wRsCp.EOF Or wRsCp.BOF)
      .AddItem
      wRow = .Rows - 1
      .Value(wRow, "TdSr") = wRow
      .Value(wRow, "TdFrRmLoc") = wRsCp!qLoc
      .Value(wRow, "TdFrRmDc") = "C"
      .Value(wRow, "TdToRmLoc") = ADC("wCnLocCd")
      .Value(wRow, "wTdFrRmIr") = "I"
      .Value(wRow, "TdRmCd") = wRsCp!TdRmCd
      .Value(wRow, "TdLotNo") = ""
      .Value(wRow, "TdRmSz") = wRsCp!TdRmSz
      .Value(wRow, "TdRmSz2") = wRsCp!TdRmSz2
      .Value(wRow, "TdRmSz3") = wRsCp!TdRmSz3
            
      '****** Sachin - Calculation of Stock Rate
      '****** If a Particular RmsCtg Stock is maintained Stock Rate Wise then
      '       -> The System Checks whether the Average Stock Rate has been enabled in the Head Master
      '       -> If Enabled then the system checks whether the location belongs to valid location types fo Stock Rate Calculation
      '       -> If Yes then the Average Stock Rate is Calculated accordingly
      '       -> Else the Stock Rate From the Source is Copied to Target
      '       -> If the stock of a particular Rm SubCtg is not mailtained Rate Wise then the Stock Rate Copied is 0.
      
      '****** Sachin 3.02.0 24-12-07 - Current Partition
      wTdRmStkRtYn = moCn.GetFldVal("Select PValue1 from Param Join RmMst on RmCd= '" + _
                 .Value(wRow, "TdRmCd") + "' and RmPrtKey='" + ctCurrPrtn + "' where PTyp= 'RmSCtg' and PmCd=RmCtg and " + _
                 " PScd = RmSCtg ")
            
      '****** If Base Currencies of the two companies dont match then Stock Rate rate copied will be zero
      If wTdRmStkRtYn = "Y" And wBaseCurMatch = True Then
          If moCn.GetFldVal("Select hAvgSRtYn from Head where HCoCd= '" + ms_TCoCd + "' and " + _
                            "HCd='" + ctSelfCmCd + "'") = "Y" And ms_TTc <> "PR" Then
            Dim wRmLocTyp As String, wb_RmLocTyp As Boolean
            wb_RmLocTyp = moCn.RecSeek("Select LocTyp from Loc where LocCoCd= '" + ADC("wCnCoCd") + "' " + _
                                       "and LocCd= '" + ADC("WCNLOCCD") + "' " + _
                                       "and LocTyp in " + ctAvgLocTypSeq)
            'Calculate Average Stock Rate
            wRmLocTyp = Trim(moCn.GetFldVal("Select LocTyp from Loc where LocCoCd= '" + ADC("wCnCoCd") + "' " + _
                                       "and LocCd= '" + ADC("WCNLOCCD") + "' "))

            If wb_RmLocTyp = True Then
              .Value(wRow, "TdRmStkRt") = GetAvgStkRt(ADC("wCnCoCd"), ADC("WCNLOCCD"), _
                                          wRsCp!TdRmCd, "", wRsCp!TdRmSz, wRsCp!TdRmSz2, wRsCp!TdRmSz3)
            Else
              .Value(wRow, "TdRmStkRt") = wRsCp!TDRMSTKRT
            End If
          Else
            .Value(wRow, "TdRmStkRt") = wRsCp!TDRMSTKRT
          End If
      Else
        .Value(wRow, "TdRmStkRt") = 0
      End If
    '*************************************************
      .Value(wRow, "TdRmQty") = wRsCp!RmQty
      .Value(wRow, "TdRmWt") = wRsCp!rmwt
      .Value(wRow, "TdDesc") = ""
      .Value(wRow, "TdDustWt") = 0
      .Value(wRow, "TdPrtKey") = ctCurrPrtn   '****** Sachin 3.02 24-12-07
      .SaveRec (wRow)
      wRsCp.MoveNext
    Loop
  End With
  MWLib.EndProcess Me
  GrdTxndA.EndCopy
  Set wRsCp = Nothing
  '****** Sachin 2.14.0 [Multi-Company]
  
End Sub

Private Sub CpyReturn()
  '*** (Jen 2.13)
  '*** Procedure to Copy from Return Memo in case of 'PR' entry
  
  Dim wRow As Single, wRsCr As MwfLib.MDORowSet
  'Dim wTdRmStkRtYn As String
  
  '****** Sachin 3.02 24-12-07 - Current Partition
  If Not moCn.RecSeek("select 'x' From RetMem where ReCoCd= '" + ADC("TCoCd") + "' and ReYy= '" + ADC("wCrReYy") + "' " + _
                      "and ReChr= '" + ADC("wCrReChr") + "' and ReNo= " + CStr(ADC("wCrReNo"))) Then
    DispMsg "Invalid Return Memo Voucher", etError: Exit Sub
  End If
  
  '****** Sachin 2.13.0 - [17-07-2006] - Multi-Currency ******
  '****** Sachin 3.02 24-12-07 - Current Partition
  If Not moCn.RecSeek("Select 'x' From RetMem Where ReCoCd='" + ADC("TCoCd") + "' And ReYy= '" + ADC("wCrReYy") + "' " + _
                      "and ReChr= '" + ADC("wCrReChr") + "' and ReNo= " + CStr(ADC("wCrReNo")) + _
                      "and (Select CmCurCd from CustMst where CmCtg= 'C' and CmCd= ReCmCd) = '" + ADC("TCurrCd") + "'") Then
    DispMsg "Currency Mismatch in Return Memo And Purchase Voucher", etError: Exit Sub
  End If
  '****** Sachin 2.13.0 - [17-07-2006] - Multi-Currency ******
  
  'Set wRsCr = moCn.OpenRes("Select IrRmCd, IrStkRt, IrRmSz, IrLotNo, sum(IrRmQty* (RedQty/ IdQty)) as TotQty, " + _
  '                         "sum((case when IrMainMet= 'Y' then IdCustmMtchWt else IrRmZWt* (1 + " + _
  '                         "(case when RmCtg in ('G', 'P', 'S') then IdGldLs/100 else 0 end) ) end) " + _
  '                         "* (RedQty/ IdQty)) as TotWt " + _
  '                         "From RetMemDet " + _
  '                         "join InvDsg on IdCoCd= RedCoCd and IdTc= RedIdTc and IdYy= RedIdYy and IdChr= RedIdChr " + _
  '                         "and IdNo= RedIdNo and IdSr= RedIdSr " + _
  '                         "join InvRm on IrCoCd= IdCoCd and IrTc= IdTc and IrYy= IdYy and IrChr= IdChr " + _
  '                         "and IrNo= IdNo and IrSr= IdSr " + _
  '                         "join RmMst on RmCd= IrRmCd " + _
  '                         "where RedCoCd= '" + ADC("TCoCd") + "' and RedYy= '" + ADC("wCrReYy") + "' " + _
  '                         "and RedChr= '" + ADC("wCrReChr") + "' and RedNo= " + CStr(ADC("wCrReNo")) + " " + _
  '                         "and IdQty> 0 Group by IrRmCd, IrRmSz, IrStkRt, IrLotNo")
  
  
  '*** (Bef 2.14 Next)
  'Set wRsCr = moCn.OpenRes("Select FrRmCd, FrRmStkRt, FrRmSz, " + _
                           "/* IrLotNo, */ sum(FrRmQty* (RedQty/ IdQty)) as TotQty, sum(FrRmWt* (RedQty/ IdQty)) as TotWt " + _
                           "From RetMemDet " + _
                           "join InvDsg on IdCoCd= RedCoCd and IdTc= RedIdTc and IdYy= RedIdYy and IdChr= RedIdChr " + _
                           "and IdNo= RedIdNo and IdSr= RedIdSr " + _
                           "join InvFgd on IfCoCd= IdCoCd and IfTc= IdTc and IfYy= IdYy and IfChr= IdChr " + _
                           "and IfNo= IdNo and IfSr= IdSr " + _
                           "join FgRm on FrCoCd= IfCoCd and FrTc= IfFdTc and FrYy= IfFdYy and FrChr= IfFdChr " + _
                           "and FrNo= IfFdNo and FrSr= IfFdSr " + _
                           "join RmMst on RmCd= FrRmCd " + _
                           "where RedCoCd= '" + ADC("TCoCd") + "' and RedYy= '" + ADC("wCrReYy") + "' " + _
                           "and RedChr= '" + ADC("wCrReChr") + "' and RedNo= " + CStr(ADC("wCrReNo")) + " " + _
                           IIF(ADC("wCrRmCtgFr") = "", " ", "and RmCtg>= '" + ADC("wCrRmCtgFr") + "' ") + _
                           IIF(ADC("wCrRmCtgTo") = "", " ", "and RmCtg<= '" + ADC("wCrRmCtgTo") + "' ") + _
                           "and IdQty> 0 Group by FrRmCd, FrRmSz, FrRmStkRt /* , IrLotNo */ ")
  '*** (Bef 2.14 Next)
  
  '*** (Jen 2.14 Next)
  ' ***** Manali 3.03 -20/06/08 - Pd Mod - 'L', 'LmlCst' added
  Dim wd_LmgCst As Double, wd_LmpCst As Double, wd_LmsCst As Double, wd_LmlCst As Double
  Call GetDtVal("LMG", ADC("TDt"), wd_LmgCst, , ADC("TCurrCd"))
  Call GetDtVal("LMP", ADC("TDt"), wd_LmpCst, , ADC("TCurrCd"))
  Call GetDtVal("LMS", ADC("TDt"), wd_LmsCst, , ADC("TCurrCd"))
  Call GetDtVal("LML", ADC("TDt"), wd_LmlCst, , ADC("TCurrCd"))
  ' ***** Manali 3.03 -20/06/08 - Pd Mod - 'L', 'LmlCst' added
  
  '****** Sachin 3.02 - Id fields added to joins
  '****** Sachin 3.02 24-12-07 - Current Partition
  ' ***** Manali 3.03 -20/06/08 - Pd Mod - 'L' included
  ' ***** Manali 3.8.1 - RmValidYN = 'Y' clause added
  Set wRsCr = moCn.OpenRes("Select max(Rm.RmCtg) as qRmCtg, max(Rm.RmSCtg) as qRmSCtg, max(IsNull(B.RmCtg, '')) as qBaseRmCtg, " + _
                           "FrRmCd, FrRmStkRt, FrRmSz, FrRmSz2,FrRmSz3,  /* IrLotNo, */ " + _
                           "(case max(IsNull(B.RmCtg, '')) when 'G' then " + CStr(wd_LmgCst) + " " + _
                           "when 'P' then " + CStr(wd_LmgCst) + " when 'S' then " + CStr(wd_LmgCst) + " when 'L' then " + CStr(wd_LmgCst) + " else 0 end) as qXLme, " + _
                           "sum(FrRmQty* (RedQty/ IdQty)) as TotQty, " + _
                           "sum((case Rm.RmCtg when 'G' then (FrRmWt + RedGrsWtDiff) else FrRmWt end) * (RedQty/ IdQty)) as TotWt " + _
                           "From RetMemDet " + _
                           "join InvDsg on IdCoCd= RedCoCd and IdTc= RedIdTc and IdYy= RedIdYy and IdChr= RedIdChr " + _
                           "and IdNo= RedIdNo and IdSr= RedIdSr " + _
                           "join InvFgd on IfCoCd= IdCoCd and IfTc= IdTc and IfYy= IdYy and IfChr= IdChr " + _
                           "and IfNo= IdNo and IfSr= IdSr and IfPrtKey=IdPrtKey " + _
                           "join FgRm on FrFdIdNo=IfFdIdNo and FrCoCd= IfCoCd and FrTc= IfFdTc and FrYy= IfFdYy and FrChr= IfFdChr " + _
                           "and FrNo= IfFdNo and FrSr= IfFdSr and FrPrtKey=IfPrtKey " + _
                           "join RmMst Rm on Rm.RmCd= FrRmCd and Rm.RmPrtKey='" + ctCurrPrtn + "' " + _
                           "left outer join RmMst B on B.RmCd= Rm.RmBaseCd and B.RmPrtKey='" + ctCurrPrtn + "' " + _
                           "where RedCoCd= '" + ADC("TCoCd") + "' and RedYy= '" + ADC("wCrReYy") + "' " + _
                           "and RedChr= '" + ADC("wCrReChr") + "' and RedNo= " + CStr(ADC("wCrReNo")) + _
                           IIF(ADC("wCrRmCtgFr") = "", " ", "and Rm.RmCtg>= '" + ADC("wCrRmCtgFr") + "' ") + _
                           IIF(ADC("wCrRmCtgTo") = "", " ", "and Rm.RmCtg<= '" + ADC("wCrRmCtgTo") + "' ") + _
                           "and IdQty> 0 and Rm.RmValidYN = 'Y' Group by FrRmCd, FrRmSz, FrRmSz2,FrRmSz3,  FrRmStkRt /* , IrLotNo */ " + _
                           "Order By qRmCtg, qRmSCtg, FrRmCd, FrRmSz, FrRmSz2, FrRmSz3, FrRmStkRt ")
                           
  '*************Bhavna  Main metal wt = RedGrsWt - sum of (FrRmWt) for MainMet <> 'Y'
  
  If wRsCr.RecCount = 0 Then DispMsg "No Record Exists", etError: Exit Sub
  
  GrdTxndA.StartCopy
  MWLib.BeginProcess Me, "Copying The Voucher Details ..."
  With GrdTxndA
    Do While Not (wRsCr.EOF Or wRsCr.BOF)
      .AddItem
      wRow = .Rows - 1
      .Row = wRow
      ' Bef 3.6.0
      '.Value(wRow, "TdSr") = wRow       ' Manali 3.2.2 uncommented
      ' ****** Manali 3.6.0
      If .Value(wRow, "TdSr") = 0 Then .Value(wRow, "TdSr") = .MaxVal("TdSr") + 1
      .Value(wRow, "TdFrRmLoc") = ADC("TFrRmLoc")
      If ADC("wFrRmIr") = "R" Then
        .Value(wRow, "TdFrRmDc") = "D"
      Else
        .Value(wRow, "TdFrRmDc") = "C"
      End If
      .Value(wRow, "TdPrtKey") = ctCurrPrtn   '****** Sachin 3.02.0
      .Value(wRow, "TdToRmLoc") = ADC("TToRmLoc")
      .Value(wRow, "wTdFrRmIr") = ADC("wFrRmIr")
      .Value(wRow, "TdRmCd") = wRsCr!FrRmCd
      .Value(wRow, "TdLotNo") = ""    'wRsCr!IrLotNo
      .Value(wRow, "TdRmSz") = wRsCr!FrRmSz
      .Value(wRow, "TdRmSz2") = wRsCr!FrRmSz2
      .Value(wRow, "TdRmSz3") = wRsCr!FrRmSz3
      .Value(wRow, "TdRmStkRt") = wRsCr!FrRmStkRt
      
      .Value(wRow, "TdRmQty") = wRsCr!TotQty
      .Value(wRow, "TdRmWt") = wRsCr!TotWt
      .Value(wRow, "TdDesc") = ""
      .Value(wRow, "TdDustWt") = 0
      If ADC("TTc") = "MV" Then
        .Value(wRow, "TdAck") = ADC("TAck")
      Else
        .Value(wRow, "TdAck") = "Y"
      End If
      
      '*** (Jen 2.14 Next)
      Dim ws_RmCtg As String, wf_SlvMod As Boolean

      wf_SlvMod = IIF(moCn.GetFldVal("Select HSlvModYN From Head Where HCoCd = '" + gs_CoCd + "' and HCd = '" + ctSelfCmCd + "'") = "Y", True, False)
      
      ws_RmCtg = UCase(Trim(wRsCr!qRmCtg))
      
      If ws_RmCtg = "G" Then
        .Value(wRow, "TdPurRt") = GetGldRt(ctSelfCmCd, wd_LmgCst, _
                                  wRsCr!FrRmCd, 0, "", "", "C", "C", , ADC("TCurrCd"), "", "", "", "Y")
      ElseIf ws_RmCtg = "P" Then
        .Value(wRow, "TdPurRt") = GetGldRt(ctSelfCmCd, wd_LmpCst, _
                                    wRsCr!FrRmCd, 0, "", "", "C", "C", , ADC("TCurrCd"))
      ElseIf ws_RmCtg = "S" Then
        wf_SlvMod = IIF(moCn.GetFldVal("Select HSlvModYN From Head Where HCoCd = '" + gs_CoCd + "' and HCd = '" + ctSelfCmCd + "'") = "Y", True, False)
      
        If wf_SlvMod = True Then
          .Value(wRow, "TdPurRt") = GetGldRt(ctSelfCmCd, wd_LmsCst, _
                                      wRsCr!FrRmCd, 0, "", "", "C", "C", , ADC("TCurrCd"))
        Else
          .Value(wRow, "TdPurRt") = 0
        End If
      ' ***** Manali 3.03 -20/06/08 - Pd Mod - "L"
      ElseIf ws_RmCtg = "L" Then
        If mb_PdMod = True Then
          .Value(wRow, "TdPurRt") = GetGldRt(ctSelfCmCd, wd_LmlCst, _
                                      wRsCr!FrRmCd, 0, "", "", "C", "C", , ADC("TCurrCd"))
        Else
          .Value(wRow, "TdPurRt") = 0
        End If
      ' ***** Manali 3.03 -20/06/08 - Pd Mod - 'L'
      Else
        ' ***** Manali 3.10.0 - 18/05/12 - Round for mWtPerDia changed from 4 to 5 (vb rounds of (0.00055, 4) to 0.0005 instead of 0.0006)
        .Value(wRow, "TdPurRt") = GetRmRt(ctSelfCmCd, wRsCr!FrRmCd, _
                                  wRsCr!FrRmSz, wRsCr!FrRmSz, _
                                  Round(MWLib.Div(wRsCr!TotWt, wRsCr!TotQty), 5), _
                                  "C", wRsCr!qXLme, , , , , ADC("TCurrCd"))
      End If
      '*** (Jen 2.14 Next)
      
      .SaveRec (wRow)
      wRsCr.MoveNext
    Loop
  End With
  MWLib.EndProcess Me
  GrdTxndA.EndCopy
  Set wRsCr = Nothing
  
  '*** (Jen 2.13)
End Sub

Private Sub CpyFromDT()
  '*** (Jen 2.14 Next)
  '*** Procedure to Copy from 'DT' in case of 'PR' & 'MV' entry
  
  ' **** Manali 3.8.0 - Location copied From Header
  
  Dim wRow As Single, wRsDt As MwfLib.MDORowSet
  Dim ws_FrLocTyp As String, ws_FrLoc As String, ws_ToLocTyp As String, ws_ToLoc As String
  
  If ms_TTc = "PR" Then
    ws_FrLocTyp = "XP"
  ElseIf ms_TTc = "MV" And ADC("wDtTc") = "DT" Then
    '*** Bef 2.14 Patch 2) If ADC("wDtCoCd") = ADC("TCoCd") Then DispMsg "Copy Voucher Company Code Cannot Be Same As Movement Company Code", etError: Exit Sub
    ws_FrLocTyp = "S"
    ws_ToLocTyp = "XS"
  End If
  
  If ms_TTc = "PR" Or (ms_TTc = "MV" And ADC("wDtTc") = "DT") Then       '3.11.0
    ws_FrLoc = moCn.GetFldVal("select (case when LocCd= '" + ADC("TFrRmLoc") + "' then '" + ADC("TFrRmLoc") + "' " + _
                               "else '" + ADC("TToRmLoc") + "' end) from Loc " + _
                               "where LocTyp= '" + ws_FrLocTyp + "' " + _
                               "and (LocCd= '" + ADC("TFrRmLoc") + "' Or LocCd= '" + ADC("TToRmLoc") + "')")
    
    ws_ToLoc = moCn.GetFldVal("select (case when LocCd= '" + ADC("TFrRmLoc") + "' then '" + ADC("TFrRmLoc") + "' " + _
                               "else '" + ADC("TToRmLoc") + "' end) from Loc " + _
                               "where LocTyp= '" + ws_ToLocTyp + "' " + _
                               "and (LocCd= '" + ADC("TFrRmLoc") + "' Or LocCd= '" + ADC("TToRmLoc") + "')")
  End If
  
  If ms_TTc = "PR" Then
    If ws_FrLoc = "" Then DispMsg "One Of The Default RM Locations Should Be An External Purchase Location", etError: Exit Sub
  ElseIf ms_TTc = "MV" And ADC("wDtTc") = "DT" Then
    If ws_FrLoc = "" Then DispMsg "One Of The Default RM Locations Should Be A Sales Location", etError: Exit Sub
    If ws_ToLoc = "" Then DispMsg "One Of The Default RM Locations Should Be An External Sales Location", etError: Exit Sub
    
    ' **** Manali 3.8.0 - Location copied From Header
    If (ws_FrLoc = ADC("TFrRmLoc") And ADC("wFrRmIr") <> "I") Or (ws_FrLoc = ADC("TToRmLoc") And ADC("wFrRmIr") <> "R") Then
      DispMsg "Stock Location Should Be Issuing Location", etError: Exit Sub
    End If
    ' **** Manali 3.8.0 - Location copied From Header
  End If
  
  '****** Sachin 3.02 24-12-07 - Current Partition
  '3.11.0
  If Not moCn.RecSeek("select 'x' From Txn where TCoCd= '" + ADC("wDtCoCd") + "' and TTc= '" + ADC("wDtTc") + "' " + _
                      "and TYy= '" + ADC("wDtYy") + "' and TChr= '" + ADC("wDtChr") + "' " + _
                      "and TNo= " + CStr(ADC("wDtNo")) + " and TPrtKey='" + ctCurrPrtn + "' ") Then
    DispMsg "Invalid Source Voucher", etError: Exit Sub
  End If
 
  '****** Sachin 3.02 24-12-07 - Current Partition
  ' ***** Manali 3.8.1 - RmValidYN = 'Y' clause added for 'PR'
  '3.11.0
  Set wRsDt = moCn.OpenRes("Select TdRmCd, TdRmSz,TdRmSz2,TdRmSz3, TdRmStkRt,  TdLotNo,  " + _
                           IIF(ms_TTc = "PR", " IsNull (L1.LocCd, L2.LocCd) ", " '' ") + " as qLoc, " + _
                           "sum(TdRmQty* (case (case when IsNull (L1.LocCd, '')= '' then TdToRmDc else TdFrRmDc end) when 'C' then 1 when 'D' then -1 else 0 end) ) as TotQty, " + _
                           "sum(TdRmWt* (case (case when IsNull (L1.LocCd, '')= '' then TdToRmDc else TdFrRmDc end) when 'C' then 1 when 'D' then -1 else 0 end) ) as TotWt " + _
                           "From Txnd join RmMst Rm on Rm.RmCd= TdRmCd and Rm.RmPrtKey=TdPrtKey " + _
                           "left outer join Loc L1 on L1.LocCoCd= TdCoCd and L1.LocCd= TdFrRmLoc " + IIF(ADC("wDtTc") <> "MLT", " and L1.LocTyp= 'S' ", "") + _
                           "left outer join Loc L2 on L2.LocCoCd= TdCoCd and L2.LocCd= TdToRmLoc " + IIF(ADC("wDtTc") <> "MLT", " and L2.LocTyp= 'S' ", "") + _
                           "where TdCoCd= '" + ADC("wDtCoCd") + "' and TdTc= '" + ADC("wDtTc") + "' " + _
                           "and TdYy= '" + ADC("wDtYy") + "' and TdChr= '" + ADC("wDtChr") + "' " + _
                           "and TdNo= " + CStr(ADC("wDtNo")) + " and TdPrtKey='" + ctCurrPrtn + "' " + _
                           IIF(ms_TTc = "PR", " and Rm.RmValidYN = 'Y' ", "") + _
                           IIF(ADC("wDtTc") <> "MLT", " and IsNull(L1.LocTyp, L2.LocTyp)= 'S'", "") + " and IsNull(L1.LocTyp, '')<> IsNull(L2.LocTyp, '') " + _
                           "Group by TdRmCd, TdRmSz,TdRmSz2,TdRmSz3, TdRmStkRt, TdLotNo,  " + _
                           IIF(ms_TTc = "PR", " IsNull (L1.LocCd, L2.LocCd) ", " IsNull (L1.LocTyp, L2.LocTyp) ") + _
                           "having convert ( decimal (16, 3), sum(TdRmWt* (case (case when IsNull (L1.LocCd, '')= '' then TdToRmDc else TdFrRmDc end) when 'C' then 1 when 'D' then -1 else 0 end) ) )> 0 " + _
                           "Order By qLoc, TdRmCd, TdRmSz,TdRmSz2,TdRmSz3, TdRmStkRt ")

  If wRsDt.RecCount = 0 Then DispMsg "No Record Exists", etError: Exit Sub
  
  GrdTxndA.StartCopy
  MWLib.BeginProcess Me, "Copying The Voucher Details ..."
  With GrdTxndA
    Do While Not (wRsDt.EOF Or wRsDt.BOF)
      .AddItem
      wRow = .Rows - 1
      .Row = wRow
      '.Value(wRow, "TdSr") = wRow
      ' Manali 3.6.0
      If .Value(wRow, "TdSr") = 0 Then .Value(wRow, "TdSr") = .MaxVal("TdSr") + 1
      .Value(wRow, "TdFrRmLoc") = ADC("TFrRmLoc")     'ws_FrLoc     'ADC("TFrRmLoc")
      '.Value(wRow, "TdFrRmDc") = "C"
      
      
      If ADC("wFrRmIr") = "R" Then
        .Value(wRow, "TdFrRmDc") = "D"
      Else
        .Value(wRow, "TdFrRmDc") = "C"
      End If
      If ms_TTc = "PR" Then
        .Value(wRow, "TdToRmLoc") = ADC("TToRmLoc")       'wRsDt!qLoc    'ADC("TToRmLoc")
      ElseIf ms_TTc = "MV" Then
        .Value(wRow, "TdToRmLoc") = ADC("TToRmLoc")       'ws_ToLoc
      End If
      .Value(wRow, "wTdFrRmIr") = ADC("wFrRmIr")      '"I"           'ADC("wFrRmIr")
      .Value(wRow, "TdRmCd") = wRsDt!TdRmCd
      .Value(wRow, "TdLotNo") = wRsDt!TdLotNo
      .Value(wRow, "TdRmSz") = wRsDt!TdRmSz
      .Value(wRow, "TdRmSz2") = wRsDt!TdRmSz2
      .Value(wRow, "TdRmSz3") = wRsDt!TdRmSz3
      .Value(wRow, "TdRmStkRt") = wRsDt!TDRMSTKRT
      
      .Value(wRow, "TdRmQty") = wRsDt!TotQty
      .Value(wRow, "TdRmWt") = wRsDt!TotWt
      .Value(wRow, "TdDesc") = ""
      .Value(wRow, "TdDustWt") = 0
      .Value(wRow, "TdPurRt") = 0
      .Value(wRow, "TdPrtKey") = ctCurrPrtn   '****** Sachin 3.02 24-12-07
      If ADC("TTc") = "MV" Then
        .Value(wRow, "TdAck") = ADC("TAck")
      Else
        .Value(wRow, "TdAck") = "Y"
      End If
      
      .SaveRec (wRow)
      wRsDt.MoveNext
    Loop
  End With
  MWLib.EndProcess Me
  GrdTxndA.EndCopy
  Set wRsDt = Nothing
  
  '*** (Jen 2.14 Next)
End Sub

Private Sub VouchSumm(ByVal mGrpByCd As String)
  '*** Procedure to get the total Wt Issued and Received in various locations
      'for the entire voucher

  Dim wGrpByCd  As String, wGrpBy As String, wTmpTbl As String, wCommStr1 As String
  Dim wCommStr2 As String, wStr1 As String, wStr2 As String, wSqlStr As String
  Dim wStr As String, wIssTotQty As Single, wIssTotWt As Double, wRecTotQty As Single
  Dim wRecTotWt As Double, wRow As Single
  wGrpByCd = mGrpByCd     'ADC("wVSGrp")
  If wGrpByCd = "RMCTG" Then wGrpBy = "qLoc,RmCtg"
  If wGrpByCd = "RMSCTG" Then wGrpBy = "qLoc,RmCtg,RmSCtg"
  If wGrpByCd = "RMCD" Then wGrpBy = "qLoc,RmCtg,RmSCtg,RmCd"
  If wGrpByCd = "RMSZ" Then wGrpBy = "qLoc,RmCtg,RmSCtg,RmCd,RmSz"
  ' Zubin 212
  If wGrpByCd = "CD+LT" Then wGrpBy = "qLoc,RmCtg,RmSCtg,RmCd,qLotNo"
  If wGrpByCd = "CTG+LT" Then wGrpBy = "qLoc,RmCtg,qLotNo"
  If wGrpByCd = "SCTG+LT" Then wGrpBy = "qLoc,RmCtg,RmSCtg,qLotNo"
  If wGrpByCd = "LT+SZ" Then wGrpBy = "qLoc,RmCtg,RmSCtg,RmCd,RmSz,qLotNo"
  ' Zubin 212
  
  '### If wGrpByCd = "" Then wGrpBy = moCn.GetFldVal("Select hTBagVsGr from Head where HCd='" + ctSelfCmCd + "'")
  'If wGrpByCd = "" Then wGrpBy = "qLoc,RmCtg"
  
  wCommStr1 = "": wCommStr2 = "": wStr1 = "": wStr2 = "": wSqlStr = "": wStr = ""
  wIssTotQty = 0: wIssTotWt = 0: wRecTotQty = 0: wRecTotWt = 0
  wTmpTbl = "#TmpVS"
  '****** Sachin 3.02 24-12-07 - Corresponding Partition
  wCommStr1 = "(case when TdFrRmDc='C' then TdRmQty else 0 end) as qIssRmQty, " + _
              "(case when TdFrRmDc='C' then TdRmWt else 0 end) as qIssRmWt, " + _
              "(case when TdFrRmDc='D' then TdRmQty else 0 end) as qRecRmQty, " + _
              "(case when TdFrRmDc='D' then TdRmWt else 0 end) as qRecRmWt into " + wTmpTbl + " " + _
              "from Txnd Join RmMst On RmCd= TdRmCd " + IIF(gs_Partition = ctCurrPrtn, " and RmPrtKey=TdPrtKey ", "") + _
              "Where TdCoCd ='" + ADC("TCoCd") + "' And " + _
              "TdTc ='" + ADC("TTc") + "' And TdYy ='" + ADC("TYy") + "' And " + _
              "TdChr ='" + ADC("TChr") + "' And TdNo =" + CStr(ADC("TNo")) + IIF(gs_Partition = ctCurrPrtn, " And TdPrtKey='" + ctCurrPrtn + "' ", "")
              '(index=RmMst)
  wCommStr2 = "(case when TdToRmDc='C' then TdRmQty else 0 end) as qIssRmQty, " + _
              "(case when TdToRmDc='C' then TdRmWt else 0 end) as qIssRmWt, " + _
              "(case when TdToRmDc='D' then TdRmQty else 0 end) as qRecRmQty, " + _
              "(case when TdToRmDc='D' then TdRmWt else 0 end) as qRecRmWt " + _
              "from Txnd Join RmMst On RmCd= TdRmCd " + IIF(gs_Partition = ctCurrPrtn, " and RmPrtKey=TdPrtKey ", "") + _
              " Where TdCoCd ='" + ADC("TCoCd") + "' And " + _
              "TdTc ='" + ADC("TTc") + "' And TdYy ='" + ADC("TYy") + "' And " + _
              "TdChr ='" + ADC("TChr") + "' And TdNo =" + CStr(ADC("TNo")) + IIF(gs_Partition = ctCurrPrtn, " And TdPrtKey='" + ctCurrPrtn + "' ", "")
              '(index=RmMst)
' Zubin 212
  If wGrpBy = "qLoc,RmCtg" Then
    wStr1 = "Select TdFrRmLoc as qLoc, RmCtg, '*' as RmSCtg, '*' as RmCd, '*' as qLotNo, '*' as RmSz, " + wCommStr1
    wStr2 = "Select TdToRmLoc as qLoc, RmCtg, '*' as RmSCtg, '*' as RmCd, '*' as qLotNo, '*' as RmSz, " + wCommStr2
    wSqlStr = " Select qLoc, RmCtg, max(RmSCtg) as RmSCtg, max(RmCd) as RmCd, max(qLotNo) as qLotNo, max(RmSz) as RmSz, " + _
              "sum(qIssRmQty) as qIssRmQty, sum(qIssRmWt) as qIssRmWt, " + _
              "sum(qRecRmQty) as qRecRmQty, sum(qRecRmWt) as qRecRmWt " + _
              "from " + wTmpTbl + " Group By " + wGrpBy
  ElseIf wGrpBy = "qLoc,RmCtg,RmSCtg" Then
    wStr1 = "Select TdFrRmLoc as qLoc, RmCtg, RmSCtg, '*' as RmCd, '*' as qLotNo, '*' as RmSz, " + wCommStr1
    wStr2 = "Select TdToRmLoc as qLoc, RmCtg, RmSCtg, '*' as RmCd, '*' as qLotNo, '*' as RmSz, " + wCommStr2
    wSqlStr = " Select qLoc, RmCtg, RmSCtg, max(RmCd) as RmCd, max(qLotNo) as qLotNo, max(RmSz) as RmSz, " + _
              "sum(qIssRmQty) as qIssRmQty, sum(qIssRmWt) as qIssRmWt, " + _
              "sum(qRecRmQty) as qRecRmQty, sum(qRecRmWt) as qRecRmWt " + _
              "from " + wTmpTbl + " Group By " + wGrpBy
  ElseIf wGrpBy = "qLoc,RmCtg,RmSCtg,RmCd" Then
    wStr1 = "Select TdFrRmLoc as qLoc, RmCtg, RmSCtg, RmCd, '*' as qLotNo, '*' as RmSz, " + wCommStr1
    wStr2 = "Select TdToRmLoc as qLoc, RmCtg, RmSCtg, RmCd, '*' as qLotNo, '*' as RmSz, " + wCommStr2
    wSqlStr = " Select qLoc, RmCtg, RmSCtg, RmCd, max(qLotNo) as qLotNo, max(RmSz) as RmSz, " + _
              "sum(qIssRmQty) as qIssRmQty, sum(qIssRmWt) as qIssRmWt, " + _
              "sum(qRecRmQty) as qRecRmQty, sum(qRecRmWt) as qRecRmWt " + _
              "from " + wTmpTbl + " Group By " + wGrpBy
  ElseIf wGrpBy = "qLoc,RmCtg,RmSCtg,RmCd,RmSz" Then
    wStr1 = "Select TdFrRmLoc as qLoc, RmCtg, RmSCtg, RmCd, '*' as qLotNo, TdRmSz as RmSz, " + wCommStr1
    wStr2 = "Select TdToRmLoc as qLoc, RmCtg, RmSCtg, RmCd, '*' as qLotNo, TdRmSz as RmSz, " + wCommStr2
    wSqlStr = " Select qLoc, RmCtg, RmSCtg, RmCd, max(qLotNo) as qLotNo, RmSz, sum(qIssRmQty) as qIssRmQty, " + _
              "sum(qIssRmWt) as qIssRmWt, sum(qRecRmQty) as qRecRmQty, " + _
              "sum(qRecRmWt) as qRecRmWt from " + wTmpTbl + " Group By " + wGrpBy
  ' Zubin 212
  ElseIf wGrpBy = "qLoc,RmCtg,RmSCtg,RmCd,qLotNo" Then
    wStr1 = "Select TdFrRmLoc as qLoc, RmCtg, RmSCtg, RmCd, TdLotNo as qLotNo, '*' as RmSz, " + wCommStr1
    wStr2 = "Select TdToRmLoc as qLoc, RmCtg, RmSCtg, RmCd, TdLotNo as qLotNo, '*' as RmSz, " + wCommStr2
    wSqlStr = " Select qLoc, RmCtg, RmSCtg, RmCd, qLotNo, max(RmSz) as RmSz, sum(qIssRmQty) as qIssRmQty, " + _
              "sum(qIssRmWt) as qIssRmWt, sum(qRecRmQty) as qRecRmQty, " + _
              "sum(qRecRmWt) as qRecRmWt from " + wTmpTbl + " Group By " + wGrpBy
  ElseIf wGrpBy = "qLoc,RmCtg,qLotNo" Then
    wStr1 = "Select TdFrRmLoc as qLoc, RmCtg, '*' as RmSCtg, '*' as RmCd, TdLotNo as qLotNo, '*' as RmSz, " + wCommStr1
    wStr2 = "Select TdToRmLoc as qLoc, RmCtg, '*' as RmSCtg, '*' as RmCd, TdLotNo as qLotNo, '*' as RmSz, " + wCommStr2
    wSqlStr = " Select qLoc, RmCtg, max(RmSCtg) as RmSCtg, max(RmCd) as RmCd, qLotNo, max(RmSz) as RmSz, sum(qIssRmQty) as qIssRmQty, " + _
              "sum(qIssRmWt) as qIssRmWt, sum(qRecRmQty) as qRecRmQty, " + _
              "sum(qRecRmWt) as qRecRmWt from " + wTmpTbl + " Group By " + wGrpBy
  ElseIf wGrpBy = "qLoc,RmCtg,RmSCtg,qLotNo" Then
    wStr1 = "Select TdFrRmLoc as qLoc, RmCtg, RmSCtg, '*' as RmCd, TdLotNo as qLotNo, '*' as RmSz, " + wCommStr1
    wStr2 = "Select TdToRmLoc as qLoc, RmCtg, RmSCtg, '*' as RmCd, TdLotNo as qLotNo, '*' as RmSz, " + wCommStr2
    wSqlStr = " Select qLoc, RmCtg, RmSCtg, max(RmCd) as RmCd, qLotNo, max(RmSz) as RmSz, sum(qIssRmQty) as qIssRmQty, " + _
              "sum(qIssRmWt) as qIssRmWt, sum(qRecRmQty) as qRecRmQty, " + _
              "sum(qRecRmWt) as qRecRmWt from " + wTmpTbl + " Group By " + wGrpBy
  ElseIf wGrpBy = "qLoc,RmCtg,RmSCtg,RmCd,RmSz,qLotNo" Then
    wStr1 = "Select TdFrRmLoc as qLoc, RmCtg, RmSCtg, RmCd, TdLotNo as qLotNo, TdRmSz as RmSz, " + wCommStr1
    wStr2 = "Select TdToRmLoc as qLoc, RmCtg, RmSCtg, RmCd, TdLotNo as qLotNo, TdRmSz as RmSz, " + wCommStr2
    wSqlStr = " Select qLoc, RmCtg, RmSCtg, RmCd, qLotNo, RmSz, sum(qIssRmQty) as qIssRmQty, " + _
              "sum(qIssRmWt) as qIssRmWt, sum(qRecRmQty) as qRecRmQty, " + _
              "sum(qRecRmWt) as qRecRmWt from " + wTmpTbl + " Group By " + wGrpBy
  ' Zubin 212
  End If
  
  wStr = wStr1 + " Union All " + wStr2

  '*** Creating Temporary Table #TmpVS ***
  moCn.CreateTmpTable wTmpTbl, wStr
  
  Dim wRsVS As MwfLib.MDORowSet
  Set wRsVS = moCn.OpenRes(wSqlStr)
   
  GrdVS.AllowAdd = True
  GrdVS.AllowDelete = True

  GrdVS.Rows = 1
  With GrdVS
    Do While Not (wRsVS.EOF Or wRsVS.BOF)
      .AddItem
      wRow = .Rows - 1
      .Value(wRow, "wVSLoc") = wRsVS!qLoc
      .Value(wRow, "wVSRmCtg") = wRsVS!RmCtg
      .Value(wRow, "wVSRmSCtg") = wRsVS!RmSCtg
      .Value(wRow, "wVSRmCd") = wRsVS!RmCd
      ' Zubin 212
      .Value(wRow, "wVSLotNo") = wRsVS!qLotNo
      ' Zubin 212
      .Value(wRow, "wVSRmSz") = wRsVS!RmSz
      .Value(wRow, "wVSIssRmQty") = wRsVS!qIssRmQty
      .Value(wRow, "wVSIssRmWt") = wRsVS!qIssRmWt
      .Value(wRow, "wVSRecRmQty") = wRsVS!qRecRmQty
      .Value(wRow, "wVSRecRmWt") = wRsVS!qRecRmWt
      Select Case Trim(wRsVS!RmCtg)
      ' ***** Manali 3.03 -20/06/08 - Pd Mod - "L" included
      Case Is = "G", "P", "S", "L", "M", "X"
        wIssTotWt = wIssTotWt + wRsVS!qIssRmWt
        wRecTotWt = wRecTotWt + wRsVS!qRecRmWt
      Case Is = "D", "C"
        wIssTotWt = wIssTotWt + MWLib.Div(wRsVS!qIssRmWt, 5)
        wRecTotWt = wRecTotWt + MWLib.Div(wRsVS!qRecRmWt, 5)
      End Select
      wIssTotQty = wIssTotQty + wRsVS!qIssRmQty
      wRecTotQty = wRecTotQty + wRsVS!qRecRmQty
      wRsVS.MoveNext
    Loop
  End With
  ADC("wVsIssTotWt") = wIssTotWt
  ADC("wVsIssTotQty") = wIssTotQty
  ADC("wVsRecTotWt") = wRecTotWt
  ADC("wVsRecTotQty") = wRecTotQty
  GrdVS.AllowAdd = False
  GrdVS.AllowDelete = False
  Set wRsVS = Nothing
End Sub
Private Function FnRmCtg(ByVal mRmCd As String) As String
  '*** Function which returns the RmCtg of the RmCode passed as the parameter
  '*** The reason for writing this function is that RmCtg is frequently needed in this entry
  FnRmCtg = Trim(moCn.GetFldVal("Select RmCtg from RmMst where RmCd='" + mRmCd + "'" + IIF(gs_Partition = ctCurrPrtn, " and RmPrtKey='" + ctCurrPrtn + "'", "")))
End Function
Private Function FnRmSCtg(ByVal mRmCd As String) As String
  '*** Function which returns the Sub Ctg of the RmCode passed as the parameter
  FnRmSCtg = Trim(moCn.GetFldVal("Select RmSCtg from RmMst where RmCd='" + mRmCd + "'" + IIF(gs_Partition = ctCurrPrtn, " and RmPrtKey='" + ctCurrPrtn + "'", "")))
End Function
Private Function FnwTotNetWt(Optional ByVal pb_NetWt = False) As Double
  ' ***** Manali 3.6.0 - 19/11/09 - Pure Net Wt for 'X' - Function Changed
  '*** Function which is used to calculate the Net Wt
  Dim wTotNetWt As Double, i As Integer, wctg As String, wRmWt As Double
  wTotNetWt = 0
  For i = 1 To (GrdTxndA.Rows - 1)
    wctg = FnRmCtg(GrdTxndA.Value(i, "TdRmCd"))
    If wctg = ADC("wCtg") Then
    ' ***** Manali 3.03 -20/06/08 - Pd Mod - "L" included
      wRmWt = GrdTxndA.Value(i, "TdRmWt") * _
            (IIF(wctg = "G" Or wctg = "P" Or wctg = "S" Or wctg = "L", moCn.GetFldVal("select RmPurityWt from RmMst where RmCd='" + GrdTxndA.Value(i, "TdRmCd") + "' and RmPrtKey='" + ctCurrPrtn + "'"), _
              IIF(wctg = "X" And Not pb_NetWt, moCn.GetFldVal("select RmBasePurityWt from RmMst where RmCd='" + GrdTxndA.Value(i, "TdRmCd") + "' and RmPrtKey='" + ctCurrPrtn + "'"), 1)))
    Else
      wRmWt = 0
    End If
    If Not GrdTxndA.IsDel(i) Then wTotNetWt = wTotNetWt + wRmWt * _
                                  IIF(GrdTxndA.Value(i, "TdFrRmDc") = "D", 1, _
                                  IIF(GrdTxndA.Value(i, "TdFrRmDc") = "C", -1, 0))
  Next i
  FnwTotNetWt = Round(wTotNetWt, 3)
End Function


Private Function FnwTotNetWt_Bef360() As Double
'  '*** Function which is used to calculate the Net Wt
'  Dim wTotNetWt as double, i As Integer, wctg As String, wRmWt as double
'  wTotNetWt = 0
'  For i = 1 To (GrdTxndA.Rows - 1)
'    wctg = FnRmCtg(GrdTxndA.Value(i, "TdRmCd"))
'    If wctg = ADC("wCtg") Then
'    ' ***** Manali 3.03 -20/06/08 - Pd Mod - "L" included
'      wRmWt = GrdTxndA.Value(i, "TdRmWt") * _
'            (IIF(wctg = "G" Or wctg = "P" Or wctg = "S" Or wctg = "L", moCn.GetFldVal("select RmPurityWt from RmMst where RmCd='" + GrdTxndA.Value(i, "TdRmCd") + "' and RmPrtKey='" + ctCurrPrtn + "'"), 1))
'    Else
'      wRmWt = 0
'    End If
'    If Not GrdTxndA.IsDel(i) Then wTotNetWt = wTotNetWt + wRmWt * _
'                                  IIF(GrdTxndA.Value(i, "TdFrRmDc") = "D", 1, _
'                                  IIF(GrdTxndA.Value(i, "TdFrRmDc") = "C", -1, 0))
'  Next i
'  FnwTotNetWt = Round(wTotNetWt, 3)
End Function

Private Sub DispFra(ByVal pv_TxnRmFra As en_TxnRmFra)
  '*** This is the common procedure which is used to toggle between hiding and displaying
      'the frames (i.e. if the frame is visible then make it invisible and viseversa)
      'associated with a command button

  Select Case pv_TxnRmFra
  Case Is = Summ
    If FraVS.Visible = True Then
      FraVS.Visible = False
      FraVS.Enabled = False
      Call EnaDisaCmds(False, CmdVS, Summ)
      CmdVS.SetFocus
    Else
      FraVS.Visible = True
      FraVS.Enabled = True
      CmdVS.SetFocus
      FraVS.ZOrder
      Call EnaDisaCmds(True, CmdVS)
    End If
  Case Is = Cpy
    If FraCP.Visible = True Then
      FraCP.Visible = False
      FraCP.Enabled = False
      Call EnaDisaCmds(False, CmdCP, Cpy)
      CmdCP.SetFocus
    Else
      FraCP.Visible = True
      FraCP.Enabled = True
      CmdCP.SetFocus
      FraCP.ZOrder
      Call EnaDisaCmds(True, CmdCP)
    End If
  '****** Sachin 2.14.0 - [Multi-Company]
  Case Is = Contra
    If FraContra.Visible = True Then
      FraContra.Visible = False
      FraContra.Enabled = False
      Call EnaDisaCmds(False, CmdContra, Contra)
      CmdContra.SetFocus
    Else
      FraContra.Visible = True
      FraContra.Enabled = True
      CmdContra.SetFocus
      FraContra.ZOrder
      Call EnaDisaCmds(True, CmdContra)
    End If
   '****** Sachin 2.14.0 - [Multi-Company]
  '***************************Geeta**************Emr206**********23/04/04
  Case Is = vch
    If FraGV.Visible = True Then
      FraGV.Visible = False
      FraGV.Enabled = False
      Call EnaDisaCmds(False, CmdGV, vch)
      CmdGV.SetFocus
    Else
      FraGV.Visible = True
      FraGV.Enabled = True
      CmdGV.SetFocus
      FraGV.ZOrder
      Call EnaDisaCmds(True, CmdGV)
    End If
 '***********************Geeta***********
 ' Zubin 212
  Case Is = LotNo
    If FraLotNo.Visible = True Then
      FraLotNo.Visible = False
      Call EnaDisaCmds(False, CmdLN, LotNo)
      GrdTxndA.SetFocus
    Else
      FraLotNo.Visible = True
      ADC("wLotNo").SetFocus
      FraLotNo.ZOrder
      Call EnaDisaCmds(True, CmdLN)
    End If
 ' Zubin 212
  
  '*** (Jen 2.13)
  Case Is = CpyRet
    If FraCpyRet.Visible = True Then
      FraCpyRet.Visible = False
      FraCpyRet.Enabled = False
      Call EnaDisaCmds(False, CmdCpyRet, CpyRet)
      CmdCpyRet.SetFocus
    Else
      FraCpyRet.Visible = True
      FraCpyRet.Enabled = True
      ADC("wCrReYy").SetFocus
      FraCpyRet.ZOrder
      Call EnaDisaCmds(True, CmdCpyRet)
    End If
  '*** (Jen 2.13)
  
  '*** (Jen 2.14 Next)
  Case Is = CpyDT
    If FraCpyDt.Visible = True Then
      FraCpyDt.Visible = False
      FraCpyDt.Enabled = False
      Call EnaDisaCmds(False, CmdCpyDt, CpyDT)
      CmdCpyDt.SetFocus
    Else
      FraCpyDt.Visible = True
      FraCpyDt.Enabled = True
       If ms_TTc = "PR" Then ADC("wDtYy").SetFocus Else ADC("wDtCoCd").SetFocus '3.11.0
      FraCpyDt.ZOrder
      Call EnaDisaCmds(True, CmdCpyDt)
    End If
  '*** (Jen 2.14 Next)
  

  ' ***** Manali 3.6.0 - 23/10/09 - Copy Rm From Loc
  Case Is = CpyLocRm
    If FraCpLocRm.Visible = True Then
      FraCpLocRm.Visible = False
      FraCpLocRm.Enabled = False
      Call EnaDisaCmds(False, CmdCpLocRm, CpyLocRm)
      CmdCpLocRm.SetFocus
    Else
      FraCpLocRm.Visible = True
      FraCpLocRm.Enabled = True
      ADC("wLocRmCtg").SetFocus
      FraCpLocRm.ZOrder
      Call EnaDisaCmds(True, CmdCpLocRm)
    End If
  ' ***** Manali 3.6.0 - 23/10/09 - Copy Rm From Loc
  
  ' ***** Manali 3.6.0 - 23/10/09 - Copy Rm From Loc
  Case Is = ImpExc
    If FraImpExc.Visible = True Then
      FraImpExc.Visible = False
      FraImpExc.Enabled = False
      Call EnaDisaCmds(False, CmdImpExc, ImpExc)
      CmdImpExc.SetFocus
    Else
      FraImpExc.Visible = True
      FraImpExc.Enabled = True
      CmdImpExcPath.SetFocus
      FraImpExc.ZOrder
      Call EnaDisaCmds(True, CmdImpExc)
    End If
  ' ***** Manali 3.6.0 - 23/10/09 - Copy Rm From Loc
  
    '4.1.2
    Case Is = DET
    If FraGrdTxndADetsDisp.Enabled = True Then
      FraGrdTxndADetsDisp.Enabled = False
      Call EnaDisaCmds(False, CmdGrdTxndADets, DET)
      GrdTxndA.SetFocus
    Else
      FraGrdTxndADetsDisp.Enabled = True
      FraGrdTxndADetsDisp.ZOrder
      ADC("WTDBSTODTC").SetFocus
      Call EnaDisaCmds(True, CmdGrdTxndADets)
    End If
    
    '4.3.0.0
    Case Is = ChgDt
    If FraCD.Visible = True Then
      FraCD.Visible = False
      FraCD.Enabled = False
      Call EnaDisaCmds(False, CmdChgDt, ChgDt)
      CmdChgDt.SetFocus
    Else
      FraCD.Visible = True
      FraCD.Enabled = True
      CmdChgDt.SetFocus
      FraCD.ZOrder
      ADC("wTDt") = ADC("TDt")
      Call EnaDisaCmds(True, CmdChgDt)
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
      If UCase(ADC.MenuCd) = UCase("CNV") Then
        CmdShowTotOk.SetFocus
      Else
        ADC("TADDAMT1").SetFocus
      End If
      FraShowTot.ZOrder
      Call EnaDisaCmds(True, CmdShowTot)
    End If

  End Select
End Sub

' Zubin 212
'Private Sub EnaDisaCmds(ByVal pv_ShowFra As Boolean, Optional ByVal pv_CmdBut As MWCTL_BTN1)
Private Sub EnaDisaCmds(ByVal pv_ShowFra As Boolean, Optional ByVal pv_CmdBut As MWCTL_BTN1, Optional ByVal pv_DispFra As en_DispFra)
  '*** This procedure is used to give the toggle option (enable and disable) to
      'the command buttons
  
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
      'pv_CmdBut.CausesValidation = False
' Zubin 212
      '*** restoring the 'Causes Validation' state of the current pv_cmdbutton
      If Not IsMissing(pv_DispFra) Then
        pv_CmdBut.CausesValidation = mArr_Cv(pv_DispFra)
      End If
' Zubin 212
    End If
  End If
  FraNKeyAll.Enabled = Not pv_ShowFra
  ADC.AllowSave = Not pv_ShowFra
  ADC.AllowDelete = Not pv_ShowFra
  ADC.AllowCancel = Not pv_ShowFra
' Zubin 212
'  CmdVS.Enabled = Not pv_ShowFra
'  CmdCP.Enabled = Not pv_ShowFra
'  CmdGV.Enabled = Not pv_ShowFra
'  CmdLN.Enabled = Not pv_ShowFra
  
  If pv_ShowFra = True Then
    '** All command buttons are disabled
    CmdGV.Enabled = Not pv_ShowFra
    CmdVS.Enabled = Not pv_ShowFra
    CmdCP.Enabled = Not pv_ShowFra
    CmdContra.Enabled = Not pv_ShowFra    '****** Sachin 2.14.0
    CmdLN.Enabled = Not pv_ShowFra
    CmdCpyRet.Enabled = Not pv_ShowFra    '*** (Jen 2.13)
    CmdCpyDt.Enabled = Not pv_ShowFra   '*** (Jen 2.14 Next)
    CmdCpLocRm.Enabled = Not pv_ShowFra   ' ***** Manali 3.6.0 - 23/10/09 - Copy Rm From Loc
    CmdImpExc.Enabled = Not pv_ShowFra   ' ***** Manali 3.9.0 - Import from Excel
    CmdGrdTxndADets.Enabled = Not pv_ShowFra    '4.1.2
    CmdChgDt.Enabled = Not pv_ShowFra           '4.3.0.0
    CmdShowTot.Enabled = Not pv_ShowFra   '578-33
    CmdInvHd.Enabled = Not pv_ShowFra
  Else
    '*** Restoring Ena/Disa state of all Command Buttons
    CmdGV.Enabled = mArr_Ed(vch)
    CmdVS.Enabled = mArr_Ed(Summ)
    CmdCP.Enabled = mArr_Ed(Cpy)
    CmdContra.Enabled = mArr_Ed(Contra)
    CmdLN.Enabled = mArr_Ed(LotNo)
    CmdCpyRet.Enabled = mArr_Ed(CpyRet)   '*** (Jen 2.13)
    CmdCpyDt.Enabled = mArr_Ed(CpyDT)   '*** (Jen 2.14 Next)
    CmdCpLocRm.Enabled = mArr_Ed(CpyLocRm)    ' ***** Manali 3.6.0 - 23/10/09 - Copy Rm From Loc
    CmdImpExc.Enabled = mArr_Ed(ImpExc)    ' ***** Manali 3.9.0 - Import from Excel
    CmdGrdTxndADets.Enabled = mArr_Ed(ImpExc)   '4.1.2
    CmdChgDt.Enabled = mArr_Ed(ChgDt)           '4.3.0.0
    CmdShowTot.Enabled = mArr_Ed(ShowTot)   '578-33
    CmdInvHd.Enabled = mArr_Ed(InvHd)
  End If
' Zubin 212
  If Not pv_CmdBut Is Nothing Then pv_CmdBut.Enabled = True
End Sub
Private Sub HideAllFras()
  '*** This procedure hides all the frames other than FraNKeyAll
  FraNKeyAll.Visible = True
  FraVS.Visible = False
  FraVS.Enabled = False
  FraCP.Visible = False
  FraCP.Enabled = False
  
  FraCD.Visible = False         '4.3.0.0
  FraCD.Enabled = False         '4.3.0.0
  
  FraContra.Visible = False
  FraContra.Enabled = False

  '*** (Jen 2.13)
  FraCpyRet.Visible = False
  FraCpyRet.Enabled = False
  '*** (Jen 2.13)
  
  '*** (Jen 2.14 Next)
  FraCpyDt.Visible = False
  FraCpyDt.Enabled = False
  '*** (Jen 2.14 Next)
  
  ' ***** Manali 3.6.0 - 23/10/09 - Copy Rm From Loc
  FraCpLocRm.Visible = False
  FraCpLocRm.Enabled = False
  ' ***** Manali 3.6.0 - 23/10/09 - Copy Rm From Loc
  
  ' ***** Manali 3.6.0 - 23/10/09 - Copy Rm From Loc
  FraImpExc.Visible = False
  FraImpExc.Enabled = False
  ' ***** Manali 3.6.0 - 23/10/09 - Copy Rm From Loc
  '578-33 hide show totals frame
  FraShowTot.Visible = False
  FraShowTot.Enabled = False
  
  FraGrdTxndADetsDisp.Enabled = False       '4.1.2
  
End Sub

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
  ws_BtnStr = "CmdCPGo,CmdGVGo,CmdLNGo"
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
  Set m_ObjCmdGetRmRt = ADC.Connection.OpenQuery("ForCpy_GetRmRt")
  Set m_ObjCmdGetLabRt = ADC.Connection.OpenQuery("ForCpy_GetLabRt")
    Set m_ObjCmdGetRmRt_IPD = ADC.Connection.OpenQuery("ForCpy_GetRmRt_IPD")
    Set m_objCmdGetRmRt_Base = ADC.Connection.OpenQuery("ForCpy_GetRmRt_base")

  Call FrmActivate(Me)
End Sub
Private Sub Form_Deactivate()
  If Not (m_ObjCmdGetRmRt Is Nothing) Then m_ObjCmdGetRmRt.QryClose
  If Not (m_ObjCmdGetLabRt Is Nothing) Then m_ObjCmdGetLabRt.QryClose
  If Not (m_ObjCmdGetRmRt_IPD Is Nothing) Then m_ObjCmdGetRmRt_IPD.QryClose
  If Not (m_objCmdGetRmRt_Base Is Nothing) Then m_objCmdGetRmRt_Base.QryClose

  Call FrmDeActivate(Me)
End Sub
Private Sub ATXT_Validate(Index As Integer, Cancel As Boolean)
  Cancel = ATXT(Index).Validate
End Sub
Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    If UnloadMode <> vbFormCode Then Cancel = True: Exit Sub
    If Not m_ObjCmdGetRmRt Is Nothing Then m_ObjCmdGetRmRt.QryClose
    If Not m_ObjCmdGetLabRt Is Nothing Then m_ObjCmdGetLabRt.QryClose
    
    Set moCn = Nothing    '*** (09/08/05)    '*** (Bef speed) bef 24/11/06
End Sub
Private Sub GrdTxndA_Validate(Cancel As Boolean)
  Cancel = GrdTxndA.Validate
End Sub


Private Sub OptVch_Click(Index As Integer)
'3.11.0
Select Case Index
Case 0
    ADC("wDtTc") = "DT"
Case 1
    ADC("wDtTc") = "MLT"
End Select
ADC("wDtChr") = "": ADC("wDtNo") = ""
End Sub

' urmila NegStk
Private Sub TxtPwd_Change()
ADC("TSeoPWd") = txtPwd.Text
If ADC.IsNew = False Then ADC.SaveRec
End Sub
' urmila NegStk
' ########################  Manoj #### Ver: 2.0.6 #### Date: 25/12/2003  ########################
Private Sub TxtPwd_GotFocus()
  DispMsg "Enter SEO Password", etInfo
  
End Sub
Private Sub TxtPwd_LostFocus()
Call TxtPwd_Validate(False)
End Sub
Private Sub TxtPwd_Validate(Cancel As Boolean)
  '*** Jay 2.14 *** (SEO Change)*
  'If ADC.AddFndMode = xFndMode Then
'  mf_DtLocked = IIF((ADC("TDt") <> "" And ADC("TDt") < CDate(ms_LockDt)) _
'                                Or _
'                   (ADC("TDt") <> "" And ADC("TDt") > CDate(ms_FutureDt)) _
'                                Or _
'                  (ADC("TDt") <= CDate(ms_PurgeDt)), _
'                True, False)
      
    If Trim(txtPwd.Text) <> "" Then
    'MMAck-In acknowledge menucode can't change the value
      'If UCase(ADC.MenuCd) = UCase("MA") Then Cancel = True: DispMsg "Cannot Modify Readonly Column", etError: Exit Sub
      If Trim(txtPwd.Text) = moCn.GetFldVal("select IsNull ((select PValue from Param Join UsrAccess on UaUsrCd = PMCd " + _
                   "where PTyp= 'USR' and PMCd= '" + gs_UsrCd + "' and UaMnuCd = '" + ADC.MenuCd + "' and " + _
                   "UaSeoYn= 'Y' and PValue= '" + Trim(txtPwd.Text) + "'), " + _
                   "(Select PValue1 from Param where Ptyp= 'SEOPWD' and PmCd= 'SEOPWD' )) ") Then
          ms_PwdFlag = True: ADC("TSeoPwd") = txtPwd.Text
      Else
          ms_PwdFlag = False: Cancel = True: DispMsg "Invalid Edit Password / User Does Not Have SEO Rights ", etError
      End If
    Else
        ms_PwdFlag = moCn.RecSeek("Select PValue1 from Param where PTyp= 'SEOPWD' and PmCd= 'SEOPWD' and PValue1= '' ")
        If ms_PwdFlag = True Then ADC("TSeoPwd") = txtPwd.Text
    End If
  'End If
' Bef 2.14
'  If Trim$(txtPwd.text) = Trim$(ms_Pwd) Then
'    ADC("TSeoPwd") = txtPwd.text
'  Else
'    If Trim(txtPwd.text) <> "" Then
'      Cancel = True
'      DispMsg "Invalid SEO Password", etError
'    End If
'  End If
' Bef 2.14
'*** Jay 2.14 *** (SEO Change)*
End Sub
' ###########################################  Manoj  ###########################################

'Manoj
Private Sub CpyMVRm()
           
  Dim wTxndARow  As Single, wRow As Single, wLmeVal As Double, wDtVal As Double
  Dim wRmCtg As String, wRmSCtg As String
  wTxndARow = IIF(GrdTxndA.Row, GrdTxndA.Row, 1)
  Dim wRsCp As MwfLib.MDORowSet
  Dim wTxndFlds As String, wGroupBy As String
  ' Zubin 212 (TdLotNo added)
    
  '****** Sachin 2.13.0 - [17-07-2006] - Multi-Currency ******
  '****** Sachin 3.02 24-12-07 - Current Partition
  '*** Jay 3.1.1 [ms_TCoCd replaced with adc("wCpTCoCd")]
  If Not moCn.RecSeek("Select 'x' From Txn " + _
                           "Where TCoCd= '" + ADC("wCpTCoCd") + "' and " + _
                           "TTc= '" + ADC("wCpTTc") + "' and " + _
                           "TYy= '" + ADC("wCpTYy") + "' and " + _
                           "TChr= '" + ADC("wCpTChr") + "' and " + _
                           "TNo= " + CStr(ADC("wCpTNo")) + " And TCurrCd ='" + ADC("TCurrCd") + "' And TPrtKey='" + ctCurrPrtn + "' ") Then
    DispMsg "Currency Mismatch in Source And Target Vouchers", etError: Exit Sub
  End If
  '****** Sachin 2.13.0 - [17-07-2006] - Multi-Currency ******
  
  If UCase(ADC("wCpGrpOpt")) = "RMCDRT" Then
    wTxndFlds = "TdRmCd, TdRmStkRt, 0 as TdRmSz, 0 as TdRmSz2,0 as TdRmSz3, '' as TdLotNo, TdPurRt, TdRtByQW "
    wGroupBy = "TdRmCd, TdRmStkRt ,TdPurRt, TdRtByQw "
  ElseIf UCase(ADC("wCpGrpOpt")) = "RMCDSZ" Then
    wTxndFlds = "TdRmCd, 0 as TdRmStkRt, TdRmSz, TdRmSz2, TdRmSz3, '' as TdLotNo, 0 as TdPurRt, max(RmQwCst) as TdRtByQw "
    wGroupBy = "TdRmCd, TdRmSz, TdRmSz2, TdRmSz3"
  ElseIf UCase(ADC("wCpGrpOpt")) = "SR" Then
    wTxndFlds = "max(TdRmCd) as TdRmCd, max(TdRmStkRt) as TdRmStkRt, max(TdRmSz) as TdRmSz, max(TdRmSz2) as TdRmSz2, max(TdRmSz3) as TdRmSz3, max(TdLotNo) as TdLotNo, max(TdPurRt) as TdPurRt, max(TdRtByQw) as TdRtByQw "
    wGroupBy = "TdSr "
  ElseIf UCase(ADC("wCpGrpOpt")) = "RMCDSZRT" Then
    wTxndFlds = "TdRmCd, TdRmStkRt, TdRmSz,TdRmSz2, TdRmSz3, '' as TdLotNo, TdPurRt, TdRtByQw "
    wGroupBy = "TdRmCd, TdRmSz,TdRmSz2, TdRmSz3, TdRmStkRt,TdPurRt,TdRtByQw "
  ElseIf UCase(ADC("wCpGrpOpt")) = "CDRTLT" Then
    wTxndFlds = "TdRmCd, TdRmStkRt, TdLotNo, 0 as TdRmSz,0 as TdRmSz2,0 as TdRmSz3, TdPurRt, TdRtByQw "
    wGroupBy = "TdRmCd, TdRmStkRt, TdPurRt, TdLotNo, TdRtByQw "
  ElseIf UCase(ADC("wCpGrpOpt")) = "CDSZLT" Then
    wTxndFlds = "TdRmCd, TdRmSz, TdRmSz2, TdRmSz3, TdLotNo, 0 as TdRmStkRt, 0 as TdPurRt, max(RmQwCst) as TdRtByQw"
    wGroupBy = "TdRmCd, TdRmSz, TdRmSz2, TdRmSz3, TdLotNo "
  ElseIf UCase(ADC("wCpGrpOpt")) = "CDSZRTLT" Then
    wTxndFlds = "TdRmCd, TdRmSz,  TdRmSz2, TdRmSz3, TdRmStkRt, TdLotNo, TdPurRt, TdRtByQw "
    wGroupBy = "TdRmCd, TdRmSz, TdRmSz2, TdRmSz3,TdRmStkRt, TdPurRt, TdLotNo, TdRtByQw "
  End If
  
  '****** Sachin 3.02 24-12-07 - Current Partition
  '*** Jay 3.1.1 [ADC("WCoCd") replaced with adc("wCpTCoCd")]
  'Set wRsCp = moCn.OpenRes("Select TdFrRmLoc, TdFrRmDc, TdToRmLoc, TdRmCd, TdLotNo, TdRmSz,TdRmSz2,TdRmSz3, TdRmQty, TdRmStkRt, TdRmWt "
  Set wRsCp = moCn.OpenRes("Select " + wTxndFlds + ",Sum(TdRmQty) as TotQty, " + _
                           "Sum(TdRmWt) as TotWt From Txnd Join RmMst on TdRmCd  = RmCd " + _
                           "Where TdRmCd <> '' and TdRmCtg in ('" + wCpGldAs + "') and " + _
                           "TdCoCd= '" + ADC("wCpTCoCd") + "' and " + _
                           "TdTc= '" + ADC("wCpTTc") + "' and " + _
                           "TdYy= '" + ADC("wCpTYy") + "' and " + _
                           "TdChr= '" + ADC("wCpTChr") + "' and " + _
                           "TdNo= " + CStr(ADC("wCpTNo")) + " and TdPrtKey='" + ctCurrPrtn + "'" + _
                           "Group by " + wGroupBy)

  If wRsCp.RecCount <= 0 Then DispMsg "No Record Exists", etError: Exit Sub
  If ADC.SaveRec = False Then Exit Sub
  GrdTxndA.StartCopy
  MWLib.BeginProcess Me, "Copying The Structure ..."
  With GrdTxndA
    Do While Not (wRsCp.EOF Or wRsCp.BOF)
      .AddItem
      wRow = .Rows - 1
      .Row = wRow
      '.Value(wRow, "TdSr") = wRow
      ' Manali 3.6.0
      If .Value(wRow, "TdSr") = 0 Then .Value(wRow, "TdSr") = .MaxVal("TdSr") + 1
      .Value(wRow, "TdFrRmLoc") = ADC("TFrRmLoc")              '### ADC("TFrRmLoc")
      .Value(wRow, "wTdFrRmIr") = ADC("WFrRmIR")
      If ADC("WFrRmIR") = "R" Then
        .Value(wRow, "TdFrRmDc") = "D"
      ElseIf ADC("WFrRmIR") = "I" Then
        .Value(wRow, "TdFrRmDc") = "C"
      End If
      .Value(wRow, "TdRmCd") = wRsCp!TdRmCd
' Zubin 212
      .Value(wRow, "TdLotNo") = wRsCp!TdLotNo
' Zubin 212
      .Value(wRow, "TdRmSz") = wRsCp!TdRmSz
      .Value(wRow, "TdRmSz2") = wRsCp!TdRmSz2
      .Value(wRow, "TdRmSz3") = wRsCp!TdRmSz3
      .Value(wRow, "TdRmQty") = wRsCp!TotQty
      .Value(wRow, "TdRmWt") = wRsCp!TotWt
      .Value(wRow, "TdToRmLoc") = ADC("TToRmLoc")             '### ADC("TToRmLoc")
      '.Value(wRow, "TdDesc") = ""
      .Value(wRow, "TdDesc") = ADC("WCPTCOCD") + "/" + ADC("WCPTTC") + "/" + ADC("WCPTYY") + "/" + ADC("WCPTCHR") + "/" + CStr(ADC("WCPTNO"))
      .Value(wRow, "TdPcPerCt") = wRsCp!TotQty / wRsCp!TotWt
      .Value(wRow, "TdPurRt") = wRsCp!TdPurRt
      .Value(wRow, "TdRtByQw") = wRsCp!TdRtByQW
      '.Value(wRow, "TdRmStkRt") = .Value(wRow, "TdPurRt")
      .Value(wRow, "TdRmStkRt") = wRsCp!TDRMSTKRT
      
      'Manoj
      .Value(wRow, "TdDustWt") = 0
      'Manoj
      .Value(wRow, "TdPrtKey") = ctCurrPrtn   '****** Sachin 3.02 24-12-07
      .Value(wRow, "TdAck") = "Y"
         
      If Not .SaveRec(wRow) Then
        GrdTxndA.StartCopy
      End If
      wRsCp.MoveNext
    Loop
  End With
'      '****** Sachin 3.02 24-12-07 - Current Partition
'      wRmCtg = UCase$(moCn.GetFldVal("Select RmCtg From RmMst Where RmCd = '" & wRsCp!TdRmCd & "' and RmPrtKey='" + ctCurrPrtn + "'"))
'      ' ***** Manali 3.03 -20/06/08 - Pd Mod - "L" included
'      If wRmCtg <> "G" And wRmCtg <> "P" And wRmCtg <> "S" And wRmCtg <> "L" Then
'        '.Value(wRow, "TdPurRt") = GetRmRt(ADC("TSuppCd"), wRsCp!TdRmCd, 0, 0, 0, "S", "S")   '*** (Bef 2.13)
'
'        '*** (Jen 2.13)
'        Dim ws_BaseRmCtg As String, wf_BaseLme As Single, wd_Dt As Variant
'        If wRmCtg = "X" Then
'          ws_BaseRmCtg = UCase(Trim(moCn.GetFldVal("Select B.RmCtg from RmMst B where RmPrtKey='" + ctCurrPrtn + "' And " + _
'                         "RmCd= (Select Rm.RmBaseCd from RmMst Rm where Rm.RmCd= '" + .Value(wRow, "TdRmCd") + "' And Rm.RmPrtKey='" + ctCurrPrtn + "') ")))
'          wf_BaseLme = 0
'
'          If ws_BaseRmCtg = "G" Then
'            Call GetDtVal("LMG", adc("TDt"), wd_Dt, , adc("TCURRCD"))
'            wf_BaseLme = GetLmgSal("LMG", adc("TSUPPCD"), wd_Dt, "S", adc("TCURRCD"))
'          ElseIf ws_BaseRmCtg = "P" Then
'            Call GetDtVal("LMP", adc("TDt"), wd_Dt, , adc("TCURRCD"))
'            wf_BaseLme = GetLmgSal("LMP", adc("TSUPPCD"), wd_Dt, "S", adc("TCURRCD"))
'          ElseIf ws_BaseRmCtg = "S" Then
'            Call GetDtVal("LMS", adc("TDt"), wd_Dt, , adc("TCURRCD"))
'            wf_BaseLme = GetLmgSal("LMS", adc("TSUPPCD"), wd_Dt, "S", adc("TCURRCD"))
'          ' ***** Manali 3.03 -20/06/08 - Pd Mod - "L"
'          ElseIf ws_BaseRmCtg = "L" Then
'            Call GetDtVal("LML", adc("TDt"), wd_Dt, , adc("TCURRCD"))
'            wf_BaseLme = GetLmgSal("LML", adc("TSUPPCD"), wd_Dt, "S", adc("TCURRCD"))
'          ' ***** Manali 3.03 -20/06/08 - Pd Mod - "L"
'          End If
'        Else
'          wf_BaseLme = 0
'        End If
        
'        .Value(wRow, "TdPurRt") = GetRmRt(adc("TSuppCd"), wRsCp!TdRmCd, 0, 0, 0, "S", wf_BaseLme, "S", , , , adc("TCurrCd"))
        '*** (Jen 2.13)
'      ElseIf wRmCtg = "G" Then
'        Call GetDtVal("LMG", adc("TDt"), wDtVal, , adc("TCURRCD"))
'        .Value(wRow, "TdPurRt") = GetGldRt(adc("TSUPPCD"), _
                        GetLmgSal("LMG", adc("TSUPPCD"), wDtVal, "S", adc("TCURRCD")), _
                        wRsCp!TdRmCd, 0, "DM", "", "S", "S", , adc("TCURRCD"))
'      ElseIf wRmCtg = "P" Then
'        Call GetDtVal("LMP", adc("TDt"), wDtVal, , adc("TCURRCD"))
'        .Value(wRow, "TdPurRt") = GetGldRt(adc("TSUPPCD"), _
                        GetLmgSal("LMP", adc("TSUPPCD"), wDtVal, "S", adc("TCURRCD")), _
                        wRsCp!TdRmCd, 0, "DM", "", "S", "S", , adc("TCURRCD"))
      ' ****** Sachin 2.12 - 02/12/05 - SJM
'      ElseIf wRmCtg = "S" Then
'        Call GetDtVal("LMS", adc("TDt"), wDtVal, , adc("TCURRCD"))
'        .Value(wRow, "TdPurRt") = GetGldRt(adc("TSUPPCD"), _
                        GetLmgSal("LMS", adc("TSUPPCD"), wDtVal, "S", adc("TCURRCD")), _
                        wRsCp!TdRmCd, 0, "DM", "", "S", "S", , adc("TCURRCD"))
      ' ****** Sachin 2.12 - 02/12/05 - SJM
      ' ***** Manali 3.03 -20/06/08 - Pd Mod - 'L'
'      ElseIf wRmCtg = "L" Then
'        Call GetDtVal("LML", adc("TDt"), wDtVal, , adc("TCURRCD"))
'        .Value(wRow, "TdPurRt") = GetGldRt(adc("TSUPPCD"), _
                        GetLmgSal("LML", adc("TSUPPCD"), wDtVal, "S", adc("TCURRCD")), _
                        wRsCp!TdRmCd, 0, "DM", "", "S", "S", , adc("TCURRCD"))
      ' ***** Manali 3.03 -20/06/08 - Pd Mod - 'L'
 '     End If
      
  MWLib.EndProcess Me
  GrdTxndA.EndCopy
  Set wRsCp = Nothing
  '*** End of Copying the structure
End Sub
'Manoj
Private Sub CpyRmDet()
  '*** Procedure to Copy the structure from another Voucher of the same TcType
      'in case of 'REC' or 'CNV' entry
      'Copying the structure means not copying the RmSz, StkRt, RmQty and RmWt
      'This copy helps the user in making the entry if he is not very well versed in
      'the way he has to proceed with the entry.
  
  Dim wTxndARow  As Single, wRow As Single
  wTxndARow = IIF(GrdTxndA.Row, GrdTxndA.Row, 1)
  Dim wRsCp As MwfLib.MDORowSet
  Dim wTxndFlds As String, wGroupBy As String
  ' Zubin 212
  '****** Sachin 3.02 24-12-07 - Current Partition
  If Not moCn.RecSeek("Select 'x' From Txn " + _
                           "Where TCoCd= '" + ADC("wCpTCoCd") + "' and " + _
                           "TTc= '" + ADC("wCpTTc") + "' and " + _
                           "TYy= '" + ADC("wCpTYy") + "' and " + _
                           "TChr= '" + ADC("wCpTChr") + "' and " + _
                           "TNo= " + CStr(ADC("wCpTNo")) + " And TCurrCd ='" + ADC("TCurrCd") + "' And TPrtKey='" + ctCurrPrtn + "' ") Then
    DispMsg "Currency Mismatch in Source And Target Vouchers", etError: Exit Sub
  End If
  'CopyData- Copy from Txnd table
  If UCase(ADC("wCpGrpOpt")) = "RMCDRT" Then
    wTxndFlds = "TdRmCd, TdRmStkRt, 0 as TdRmSz, 0 as TdRmSz2,0 as TdRmSz3, '' as TdLotNo, TdPurRt, TdRtByQw, sum(TdPurAmt) "
    wGroupBy = "TdRmCd, TdRmStkRt,TdPurRt,TdRtByQw "
  ElseIf UCase(ADC("wCpGrpOpt")) = "RMCDSZ" Then
    wTxndFlds = "TdRmCd, 0 as TdRmStkRt, TdRmSz, TdRmSz2, TdRmSz3, '' as TdLotNo, 0 as TdPurRt,max(RmQw) as TdRtByQw, 0 as TdPurAmt "
    wGroupBy = "TdRmCd, TdRmSz, TdRmSz2, TdRmSz3"
  ElseIf UCase(ADC("wCpGrpOpt")) = "SR" Then
    wTxndFlds = "max(TdRmCd) as TdRmCd, max(TdRmStkRt) as TdRmStkRt, max(TdRmSz) as TdRmSz, max(TdRmSz2) as TdRmSz2, max(TdRmSz3) as TdRmSz3, max(TdLotNo) as TdLotNo, 0 as TdPurRt, 0 as TdPurAmt, max(TdRtByQw) as TdRtByQw "
    wGroupBy = "TdSr "
  ElseIf UCase(ADC("wCpGrpOpt")) = "RMCDSZRT" Then
    wTxndFlds = "TdRmCd, TdRmStkRt, TdRmSz,TdRmSz2, TdRmSz3, '' as TdLotNo, TdPurRt, TdRtByQw, sum(TdPurAmt) "
    wGroupBy = "TdRmCd, TdRmSz,TdRmSz2, TdRmSz3, TdRmStkRt,TdPurRt, TdRtByQw "
  ElseIf UCase(ADC("wCpGrpOpt")) = "CDRTLT" Then
    wTxndFlds = "TdRmCd, TdRmStkRt, TdLotNo, 0 TdRmSz,0 as TdRmSz2,0 as TdRmSz3, TdPurRt, TdRtByQw, sum(TdPurAmt) "
    wGroupBy = "TdRmCd, TdRmStkRt, TdPurRt, TdLotNo, TdRtByQw "
  ElseIf UCase(ADC("wCpGrpOpt")) = "CDSZLT" Then
    wTxndFlds = "TdRmCd, TdRmSz, TdRmSz2, TdRmSz3, TdLotNo, 0 as TdRmStkRt, 0 as TdPurRt, max(RmQw) as TdRtByQw, 0 as TdPurAmt "
    wGroupBy = "TdRmCd, TdRmSz, TdRmSz2, TdRmSz3, TdLotNo "
  ElseIf UCase(ADC("wCpGrpOpt")) = "CDSZRTLT" Then
    wTxndFlds = "TdRmCd, TdRmSz,  TdRmSz2, TdRmSz3, TdRmStkRt, TdLotNo, TdPurRt, TdRtByQw as TdRtByQw, sum(TdPurAmt) "
    wGroupBy = "TdRmCd, TdRmSz, TdRmSz2,  TdRmSz3,TdRmStkRt,TdPurRt, TdLotNo,TdRtByQw "
  End If
    
  Set wRsCp = moCn.OpenRes("Select " + wTxndFlds + ", Sum(TdRmQty) as TotQty, " + _
                           "Sum(TdRmWt)  as TotWt, Sum(TdDustWt) as TdDustWt,max(TdFrRmDc) as TdFrRmDc " + _
                           "From Txnd join RmMst on RmCd = TdRmCd " + _
                           "Where TdRmCd <> '' and TdRmCtg in ('" + wCpGldAs + "') and " + _
                           "TdCoCd= '" + ADC("wCpTCoCd") + "' and " + _
                           "TdTc= '" + ADC("wCpTTc") + "' and " + _
                           "TdYy= '" + ADC("wCpTYy") + "' and " + _
                           "TdChr= '" + ADC("wCpTChr") + "' and " + _
                           "TdNo= " + CStr(ADC("wCpTNo")) + " and TdPrtKey='" + ctCurrPrtn + "'" + _
                           "Group by " + wGroupBy)
                             
  If wRsCp.RecCount <= 0 Then DispMsg "No Record Exists", etError: Exit Sub
  If ADC.SaveRec = False Then Exit Sub
  GrdTxndA.StartCopy
  MWLib.BeginProcess Me, "Copying The Structure ..."
  With GrdTxndA
    Do While Not (wRsCp.EOF Or wRsCp.BOF)
      .AddItem
      wRow = .Rows - 1
      .Row = wRow
      If .Value(wRow, "TdSr") = 0 Then .Value(wRow, "TdSr") = .MaxVal("TdSr") + 1
      .Value(wRow, "TdFrRmLoc") = ADC("TFrRmLoc")              '### ADC("TFrRmLoc")
      If UCase(ADC("wCpTTc")) = "CNV" Then
        .Value(wRow, "TdFrRmDc") = wRsCp!TdFrRmDc
        If wRsCp!TdFrRmDc = "C" Then
          .Value(wRow, "wTdFrRmIr") = "I"
        ElseIf wRsCp!TdFrRmDc = "D" Then
          .Value(wRow, "wTdFrRmIr") = "R"
        End If
      Else
        .Value(wRow, "wTdFrRmIr") = ADC("wFrRmIr")
        If ADC("wFrRmIr") = "R" Then
          .Value(wRow, "TdFrRmDc") = "D"
        Else
          .Value(wRow, "TdFrRmDc") = "C"
        End If
      End If
      .Value(wRow, "TdRmCd") = wRsCp!TdRmCd
      ' Zubin 212
      .Value(wRow, "TdLotNo") = wRsCp!TdLotNo
      ' Zubin 212
      .Value(wRow, "TdRmSz") = wRsCp!TdRmSz
      .Value(wRow, "TdRmSz2") = wRsCp!TdRmSz2
      .Value(wRow, "TdRmSz3") = wRsCp!TdRmSz3
      .Value(wRow, "TdRmStkRt") = wRsCp!TDRMSTKRT
      .Value(wRow, "TdRmQty") = wRsCp!TotQty
      .Value(wRow, "TdRmWt") = wRsCp!TotWt
      .Value(wRow, "TdToRmLoc") = ADC("TToRmLoc")             '### ADC("TToRmLoc")
      '.Value(wRow, "TdDesc") = ""
      .Value(wRow, "TdDesc") = ADC("WCPTCOCD") + "/" + ADC("WCPTTC") + "/" + ADC("WCPTYY") + "/" + ADC("WCPTCHR") + "/" + CStr(ADC("WCPTNO"))
      .Value(wRow, "TdPurRt") = wRsCp!TdPurRt
      .Value(wRow, "TdRtByQw") = wRsCp!TdRtByQW
      .Value(wRow, "TdAck") = "Y"
      'Manoj
      .Value(wRow, "TdDustWt") = wRsCp!TdDustWt
      'Manoj
      .Value(wRow, "TdPrtKey") = ctCurrPrtn
      '****** Sachin 3.02 24-12-07
      .SaveRec (wRow)
      wRsCp.MoveNext
    Loop
  End With
  MWLib.EndProcess Me
  GrdTxndA.EndCopy
  Set wRsCp = Nothing
  '*** End of Copying the structure
End Sub

' Zubin 212
Private Sub StoreState()
  '*** This procedure is used to store the CausesValidation State and the Enabled/Disabled
      'states of all the Command Buttons
  '*** This is used for restoring the original states in the future
  
  mArr_Cv(Cpy) = CmdCP.CausesValidation
  mArr_Cv(Contra) = CmdContra.CausesValidation
  mArr_Cv(Summ) = CmdVS.CausesValidation
  mArr_Cv(LotNo) = CmdLN.CausesValidation
  mArr_Cv(vch) = CmdGV.CausesValidation
  mArr_Cv(CpyRet) = CmdCpyRet.CausesValidation
  mArr_Cv(CpyDT) = CmdCpyDt.CausesValidation    '*** (Jen 2.4 Next)
  mArr_Cv(CpyLocRm) = CmdCpLocRm.CausesValidation     ' ***** Manali 3.6.0 - 23/10/09 - Copy Rm From Loc
  mArr_Cv(ImpExc) = CmdImpExc.CausesValidation     ' ***** Manali 3.9.0 - Import From Excel
  mArr_Cv(ChgDt) = CmdChgDt.CausesValidation        '4.3.0.0
  mArr_Cv(ShowTot) = CmdShowTot.CausesValidation    '578-33
  mArr_Cv(InvHd) = CmdInvHd.CausesValidation
    
  mArr_Ed(Cpy) = CmdCP.Enabled
  mArr_Ed(Contra) = CmdContra.Enabled   '****** Sachin 2.14.0
  mArr_Ed(Summ) = CmdVS.Enabled
  mArr_Ed(LotNo) = CmdLN.Enabled
  mArr_Ed(vch) = CmdGV.Enabled
  mArr_Ed(CpyRet) = CmdCpyRet.Enabled
  mArr_Ed(CpyDT) = CmdCpyDt.Enabled   '*** (Jen 2.4 Next)
  mArr_Ed(CpyLocRm) = CmdCpLocRm.Enabled      ' ***** Manali 3.6.0 - 23/10/09 - Copy Rm From Loc
  mArr_Ed(ImpExc) = CmdImpExc.Enabled      ' ***** Manali 3.9.0 - Import From Excel
  mArr_Ed(ChgDt) = CmdChgDt.Enabled                 '4.3.0.0
  mArr_Ed(ShowTot) = CmdShowTot.Enabled   '578-33
  mArr_Ed(InvHd) = CmdInvHd.Enabled

End Sub


Private Sub CpyLocRmDet()
' ***** Manali 3.6.0 - 23/10/09 - Copy Rm From Loc
Dim wSqlStrg As String, wRsLocRm As MDORowSet, wRow As Integer
Dim wFrLoc As String
Dim wbLotNo As Boolean

  wbLotNo = moCn.RecSeek("select 'x' from Head where HCd= '" + ctSelfCmCd + "' and HCoCd= '" + ADC("TCoCd") + "' and HLotNoYN= 'Y' ")

  wFrLoc = IIF(ADC("wFrRmIR") = "I", ADC("TFrRmLoc"), ADC("TToRmLoc"))
  
  ' ***** Manali 3.8.1/3.9.0 - StkRt removed from Group By, Avg Stk Rt Calculated always
  wSqlStrg = "Select SCoCd, SLoc, SRmCd, " + IIF(wbLotNo, "SLotNo, ", "'' as SLotNo, ") + _
             "SRmSz,SRmSz2,SRmSz3, " + _
             "(Case Max(PValue1) When 'Y' Then " + _
             "     (Case When convert (decimal (10, 3), IsNull(sum(SRmDrWt- SRmCrWt), 0)) >0 Then " + _
             "        (convert (decimal (10, 2), IsNull(sum((SRmDrWt-SRmCrWt)* SRmStkRt) /sum(SRmDrWt- SRmCrWt), 0))) " + _
             "      Else 0 End) " + _
             "Else 0 End) as SRmStkRt, " + _
              "Sum(SRmDrQty - SRmCrQty) as qTotQty, Sum(SRmDrWt-SRmCrWt) as qTotWt from SyyMm " + _
              "Join RmMst on RmCd=SRmCd " + IIF(gs_Partition = ctCurrPrtn, " and SPrtKey=RmPrtKey ", "") + _
              "Join Param on PTyp='RMSCTG' and PMCd=RmCtg and PSCd=RmSCtg " + _
              "Where SCoCd='" + ADC("TCoCd") + "' and SLoc='" + wFrLoc + "' and RmCtg='" + ADC("wLocRmCtg") + "' " + _
              IIF(ADC("wLocRmCdFr") = "", " ", "and SRmCd>= '" + ADC("wLocRmCdFr") + "' ") + _
              IIF(ADC("wLocRmCdTo") = "", " ", "and SRmCd<= '" + ADC("wLocRmCdTo") + "' ") + _
              "Group By SCoCd, SLoc, SRmCd, " + IIF(wbLotNo, "SLotNo, ", "") + " SRmSz,SRmSz2,SRmSz3 " + _
              "Having Convert(decimal(10,3),sum(SRmDrWt-SRmCrWt))>0  " + _
              "Order By SCoCd, SLoc, SRmCd, " + IIF(wbLotNo, "SLotNo, ", "") + "SRmSz,SRmSz2,SRmSz3  "
  
  ' Bef 3.8.1/3.9.0
'  wSqlStrg = "Select SCoCd, SLoc, SRmCd, " + IIF(wbLotNo, "SLotNo, ", "'' as SLotNo, ") + _
'             " SRmSz, Case PValue1 When 'Y' Then SRmStkRt Else 0 End as SRmStkRt, " + _
'              "Sum(SRmDrQty - SRmCrQty) as qTotQty, Sum(SRmDrWt-SRmCrWt) as qTotWt from SyyMm " + _
'              "Join RmMst on RmCd=SRmCd " + IIF(gs_Partition = ctCurrPrtn, " and SPrtKey=RmPrtKey ", "") + _
'              "Join Param on PTyp='RMSCTG' and PMCd=RmCtg and PSCd=RmSCtg " + _
'              "Where SCoCd='" + adc("TCoCd") + "' and SLoc='" + wFrLoc + "' and RmCtg='" + adc("wLocRmCtg") + "' " + _
'              IIF(adc("wLocRmCdFr") = "", " ", "and SRmCd>= '" + adc("wLocRmCdFr") + "' ") + _
'              IIF(adc("wLocRmCdTo") = "", " ", "and SRmCd<= '" + adc("wLocRmCdTo") + "' ") + _
'              "Group By SCoCd, SLoc, SRmCd, " + IIF(wbLotNo, "SLotNo, ", "") + " SRmSz, Case PValue1 When 'Y' Then SRmStkRt Else 0 End " + _
'              "Having Convert(decimal(10,3),sum(SRmDrWt-SRmCrWt))>0  " + _
'              "Order By SCoCd, SLoc, SRmCd, " + IIF(wbLotNo, "SLotNo, ", "") + "SRmSz, Case PValue1 When 'Y' Then SRmStkRt Else 0 End "
  
  Set wRsLocRm = moCn.OpenRes(wSqlStrg)
  
  If wRsLocRm.RecCount = 0 Then DispMsg "No Record Exists", etError: Exit Sub
  
  If wRsLocRm.RecCount > 999 Then DispMsg "No of Record is greater than 999, reduce the scope", etError: Exit Sub
  
  GrdTxndA.StartCopy
  MWLib.BeginProcess Me, "Copying RM Details ..."
  With GrdTxndA
    Do While Not (wRsLocRm.EOF Or wRsLocRm.BOF)
      .AddItem
      wRow = .Rows - 1
      .Row = wRow
      '.Value(wRow, "TdSr") = wRow
      ' Manali 3.6.0
      If .Value(wRow, "TdSr") = 0 Then .Value(wRow, "TdSr") = .MaxVal("TdSr") + 1
      .Value(wRow, "TdFrRmLoc") = ADC("TFrRmLoc")
      If ADC("wFrRmIr") = "R" Then
        .Value(wRow, "TdFrRmDc") = "D"
      Else
        .Value(wRow, "TdFrRmDc") = "C"
      End If
      .Value(wRow, "TdPrtKey") = ctCurrPrtn
      .Value(wRow, "TdToRmLoc") = ADC("TToRmLoc")
      .Value(wRow, "wTdFrRmIr") = ADC("wFrRmIr")
      .Value(wRow, "TdRmCd") = wRsLocRm!SRmCd
      .Value(wRow, "TdLotNo") = wRsLocRm!SLotNo
      .Value(wRow, "TdRmSz") = wRsLocRm!SRmSz
      .Value(wRow, "TdRmSz2") = wRsLocRm!SRmSz2
      .Value(wRow, "TdRmSz3") = wRsLocRm!SRmSz3
      .Value(wRow, "TdRmStkRt") = wRsLocRm!SRmStkRt
      
      .Value(wRow, "TdRmQty") = wRsLocRm!qTotQty
      .Value(wRow, "TdRmWt") = wRsLocRm!qTotWt
      .Value(wRow, "TdDesc") = ""
      .Value(wRow, "TdDustWt") = 0
      
      .Value(wRow, "TdPurRt") = 0

      If ADC("TTc") = "MV" Then
        .Value(wRow, "TdAck") = ADC("TAck")
      Else
        .Value(wRow, "TdAck") = "Y"
      End If

      .SaveRec (wRow)
      wRsLocRm.MoveNext
    Loop
  End With
  MWLib.EndProcess Me
  GrdTxndA.EndCopy
  
  Set wRsLocRm = Nothing
' ***** Manali 3.6.0 - 23/10/09 - Copy Rm From Loc
End Sub

Private Sub UpdateTdRtByQw(ByVal RowNum As Integer)
    Dim wTdFrRmLocTyp As String
    Dim wTdToRmLocTyp As String
    Dim wRmCtg As String
    Dim wRsLotNo As MDORowSet
    With GrdTxndA
    If mf_OLORYN Then
       wTdFrRmLocTyp = moCn.GetFldVal("Select LocTyp From Loc Where LocCoCd='" + ADC("TCoCd") + "' And LocCd='" + .Value(RowNum, "TdFrRmLoc") + "'")
       wTdToRmLocTyp = moCn.GetFldVal("Select LocTyp From Loc Where LocCoCd='" + ADC("TCoCd") + "' And LocCd='" + .Value(RowNum, "TdToRmLoc") + "'")
       wRmCtg = moCn.GetFldVal("Select RmCtg From RmMst Where RmCd='" + .Value(RowNum, "TdRmCd") + "'")
      
       If (wRmCtg = "D" Or wRmCtg = "C" Or wRmCtg = "X") Then
         Set wRsLotNo = moCn.OpenRes("Select * from LotMst where LotNo = '" + .Value(RowNum, "TdLotNo") + "'")
         If (wRsLotNo.RecCount > 0) Then
           If (wRsLotNo!LotRmCtg = wRmCtg) Then
             .Value(RowNum, "TdRtByQW") = wRsLotNo!LotQW
           End If
         End If
       End If
    End If
    If .Value(RowNum, "TdRtByQW") = "" Then
      If ms_TTc = "PR" And RowNum > 1 Then
        If .Value(RowNum, "TdRmCd") = .Value(RowNum - 1, "TdRmCd") Or _
          moCn.GetFldVal("Select Count(Distinct RmCtg) from RmMst where RmZ= 'N' and RmCd In ('" + .Value(RowNum, "TdRmCd") + "', '" + .Value(RowNum - 1, "TdRmCd") + "') and RmPrtKey='" + ctCurrPrtn + "' ") = 1 Then
          .Value(RowNum, "TdRtByQW") = .Value(RowNum - 1, "TdRtByQW")
        Else
          If .Value(RowNum, "TdRtByQW") = "" Then .Value(RowNum, "TdRtByQW") = "W"
        End If
      End If
    End If
    If .Value(RowNum, "TdRtByQW") = "" Then
      .Value(RowNum, "TdRtByQW") = moCn.GetFldVal("Select " + IIF(UCase(ms_TTc) = "SA" Or UCase(ms_TTc) = "MV" Or UCase(ms_TTc) = "CNV", "RmQw", "RmQwCst") + " from RmMst where RmCd = '" + .Value(RowNum, "TdRmCd") + "' ")
    End If
    End With
End Sub
Private Sub UpdateTdPurRt(ByVal RowNum As Integer)
    Dim wTdFrRmLocTyp As String
    Dim wTdToRmLocTyp As String
    Dim wRmCtg As String
    Dim wRsLotNo As MDORowSet
          With GrdTxndA
        wTdFrRmLocTyp = moCn.GetFldVal("Select LocTyp From Loc Where LocCoCd='" + ADC("TCoCd") + "' And LocCd='" + .Value(RowNum, "TdFrRmLoc") + "'")
        wTdToRmLocTyp = moCn.GetFldVal("Select LocTyp From Loc Where LocCoCd='" + ADC("TCoCd") + "' And LocCd='" + .Value(RowNum, "TdToRmLoc") + "'")
        wRmCtg = moCn.GetFldVal("Select RmCtg From RmMst Where RmCd='" + .Value(RowNum, "TdRmCd") + "'")
      
        If (wRmCtg = "D" Or wRmCtg = "C" Or wRmCtg = "X") Then
          Set wRsLotNo = moCn.OpenRes("Select * from LotMst where LotNo = '" + .Value(RowNum, "TdLotNo") + "'")
          If (wRsLotNo.RecCount > 0) Then
            If (wRsLotNo!LotRmCtg = wRmCtg) Then
              Dim wReqRt As Double
              wReqRt = IIF(GrdTxndA.Value(RowNum, "TdRtByQw") = "Q", wRsLotNo!LotBaseRtByQ, _
              wRsLotNo!LotBaseRtByW)
              .Value(RowNum, "TdPurRt") = IIF(ADC("TcnvRt") <> 0, IIF(ADC("TMulDiv") = "D", _
               MWLib.Div(wReqRt, ADC("TCnvRt")), wReqRt * ADC("TCnvRt")), 0)
            End If
          End If
        End If
    End With
End Sub
Private Sub ImportExcel()
On Error GoTo ErrHdlr
    
    Dim wo_rsExcel As New MDORowSet, i As Integer
    
    Dim wRmCtg As String, wStkByRt As String
    Dim wRmCd As String, wRmSz As Double, wStkRt As Double, wPcPerCts As Integer, wRmQty As Integer
    Dim wLotNo As String, wRmSz2 As Double, wRmSz3 As Double
    '****** Checking For Valid Inputs From the User ******
    ' ***** Manali 3.9.0 - modified to accept Excel 2007 (.xlsx) files
    If Not OpenExcelFile(ADC("WXLFILE")) Then DispMsg "Improper File Format", etError: Exit Sub
    
    'Call OpenExcelFile(adc("WXLFILE"))
    If ChkSheet(ADC("WXLSHTNO")) = False Then DispMsg "This WorkSheet Number Does Not Exist in the specified Excel File", etError: CloseExcel: Exit Sub
    '****** Checking For Valid Inputs From the User ******

    ' ***** Manali 3.9.1 - 21/02/12 - qSr added for Order By according to data entered in Excel
    '6.1b size2,3 added
    moCn.CreateTmpTable "#TmpXls", "Select 0 as qSr, TdRmCd, TdRmSz,TdRmSz2,TdRmSz3, TdLotNo,  TdRmStkRt, TdPcPerCt, TdRmQty, TdRmWt, TdPurRt, TdPurAmt, TdDesc,TdBstOdTc,TdBstOdYy,TdBstOdChr,TdBstOdNo,TdBstOdSr,TdBstDmCd Into #TmpXls From Txnd Where 1=2"
        
     ADC("WXLFRROW") = IIF(ADC("WXLFRROW") <= 1, 2, ADC("WXLFRROW"))
    For i = ADC("WXLFRROW") To ADC("WXLTOROW")
        '****** Valid RmCode
        wRmCd = Trim(ReadCell(ADC("WXLSHTNO"), i, 2))
        If Not moCn.RecSeek("Select 'x' from RmMst Where RmCd='" + wRmCd + "' And RmPrtKey='" + ctCurrPrtn + "' and RmValidYN='Y' and RmZ='N' ") Then _
          DispMsg "Cannot Import as RmCode is Not Valid [Row :" + CStr(i) + "]", etError: CloseExcel: Exit Sub
    
        wRmCtg = moCn.GetFldVal("Select RmCtg from RmMst Where RmCd='" + wRmCd + "' ")
        wStkByRt = moCn.GetFldVal("Select PValue1 from RmMst Join Param On PTyp='RMSCTG' and PMCd=RmCtg and PSCd=RmSCTg Where RmCd='" + wRmCd + "' ")
        
        '****** Checking for valid RmCd Size in the Excel File
        'Debug.Print Val(ReadCell(ADC("WXLSHTNO"), i, 3))
        If ReadCell(ADC("WXLSHTNO"), i, 3) <> "" And Not IsNumeric(ReadCell(ADC("WXLSHTNO"), i, 3)) Then
            DispMsg "Cannot Copy as Ln/Sv/Ptr Value is not Number [Row :" + CStr(i) + "]", etError: CloseExcel: Exit Sub
        End If
        
        wRmSz = Val(ReadCell(ADC("WXLSHTNO"), i, 3))
        '6.1 new
        'If wRmSz > 0 And Not (wRmCtg = "C" Or wRmCtg = "D") Then _
            DispMsg "Cannot Copy as Ln/Sv/Ptr Should be 0 RmCtg Other Than 'D' and 'C' [Row :" + CStr(i) + "]", etError: CloseExcel: Exit Sub
        
        If wRmSz < 0 Or wRmSz > 999.9999 Then _
            DispMsg "Cannot Copy as Invalid Ln/Sv/Ptr Entered, Should be Between 0 and 999.9999 [Row :" + CStr(i) + "]", etError: CloseExcel: Exit Sub
        
        
        '6.1b size2 added
        If ReadCell(ADC("WXLSHTNO"), i, 4) <> "" And Not IsNumeric(ReadCell(ADC("WXLSHTNO"), i, 4)) Then
            DispMsg "Cannot Copy as Breadth Value is not Number [Row :" + CStr(i) + "]", etError: CloseExcel: Exit Sub
        End If
        

        wRmSz2 = Val(ReadCell(ADC("WXLSHTNO"), i, 4))
        '6.1
        'If wRmSz2 > 0 And Not (wRmCtg = "C" Or wRmCtg = "D") Then _
            DispMsg "Cannot Copy as Breadth Should be 0 RmCtg Other Than 'D' and 'C' [Row :" + CStr(i) + "]", etError: CloseExcel: Exit Sub
        
        If wRmSz2 < 0 Or wRmSz2 > 999.9999 Then _
            DispMsg "Cannot Copy as Invalid Breadth Entered, Should be Between 0 and 999.9999 [Row :" + CStr(i) + "]", etError: CloseExcel: Exit Sub
        Dim wRmSCtg As String
        wRmCtg = moCn.GetFldVal("Select RmCtg from RmMst where RmCd= '" + wRmCd + "' and RmPrtKey='" + ctCurrPrtn + "' ")
        wRmSCtg = moCn.GetFldVal("Select RmSCtg from RmMst where RmCd= '" + wRmCd + "' and RmPrtKey='" + ctCurrPrtn + "' ")
        Dim ErrMsg As String
        Dim Cancel As Boolean
        Cancel = ColWhenStkSz(2, wRmCtg, wRmSCtg, ErrMsg)
        
        If Cancel = True Then
            DispMsg "Cannot Copy, Breadth value incorrect. " + ErrMsg + " [Row :" + CStr(i) + "]", etError: CloseExcel: Exit Sub
        
        End If
        
        '6.1b size3 added
        If ReadCell(ADC("WXLSHTNO"), i, 5) <> "" And Not IsNumeric(ReadCell(ADC("WXLSHTNO"), i, 5)) Then
            DispMsg "Cannot Copy as Depth Value is not Number [Row :" + CStr(i) + "]", etError: CloseExcel: Exit Sub
        End If
        
        wRmSz3 = Val(ReadCell(ADC("WXLSHTNO"), i, 5))
        '6.1 new
        'If wRmSz3 > 0 And Not (wRmCtg = "C" Or wRmCtg = "D") Then _
            DispMsg "Cannot Copy as Depth Should be 0 RmCtg Other Than 'D' and 'C' [Row :" + CStr(i) + "]", etError: CloseExcel: Exit Sub
        
        If wRmSz3 < 0 Or wRmSz3 > 999.9999 Then _
            DispMsg "Cannot Copy as Invalid Depth Entered, Should be Between 0 and 999.9999 [Row :" + CStr(i) + "]", etError: CloseExcel: Exit Sub
        
        Cancel = ColWhenStkSz(3, wRmCtg, wRmSCtg, ErrMsg)
        
        If Cancel = True Then
            DispMsg "Cannot Copy, Breadth value incorrect. " + ErrMsg + " [Row :" + CStr(i) + "]", etError: CloseExcel: Exit Sub
        
        End If
        
        '577-err if LotNo module is yes then it reads the lotno value from excel otherwise lotno will be blank
        If moCn.RecSeek("select 'x' from Head where HCd= '" + ctSelfCmCd + "' and HCoCd= '" + ADC("TCoCd") + "' and HLotNoYN= 'N' ") Then
          wLotNo = ""
        Else
          wLotNo = ReadCell(ADC("WXLSHTNO"), i, 6)
        End If
        If Len(wLotNo) > 16 Then _
          DispMsg "Cannot Copy as LotNo Length Should be Lessthan Or Equal to 16 [Row :" + CStr(i) + "]", etError: CloseExcel: Exit Sub
        
        '****** Checking for valid RmCd StkRt in the Excel File
        If ReadCell(ADC("WXLSHTNO"), i, 7) <> "" And Not IsNumeric(ReadCell(ADC("WXLSHTNO"), i, 7)) Then
            DispMsg "Cannot Copy as Stk Rate Value is not Number [Row :" + CStr(i) + "]", etError: CloseExcel: Exit Sub
        End If
        
        wStkRt = Val(ReadCell(ADC("WXLSHTNO"), i, 7))
        
        If wStkRt > 0 And wStkByRt = "N" Then _
            DispMsg "Cannot Copy as StkRt Should be 0 For This RmCd [Row :" + CStr(i) + "]", etError: CloseExcel: Exit Sub
        
        If wStkRt < -999999.99 Or wStkRt > 999999.99 Then _
            DispMsg "Cannot Copy as Invalid StkRt Entered, Should be Between -999999.99 and 999999.99 [Row :" + CStr(i) + "]", etError: CloseExcel: Exit Sub
            
        '****** Checking for valid Pcs/Cts in the Excel File
        If (Not IsNumeric(ReadCell(ADC("WXLSHTNO"), i, 8)) And ReadCell(ADC("WXLSHTNO"), i, 8) <> "") Or Val(ReadCell(ADC("WXLSHTNO"), i, 8)) < 0 Or Val(ReadCell(ADC("WXLSHTNO"), i, 8)) > 999 Then _
            DispMsg "Cannot Copy as Invalid Pcs/Cts Entered, Should be Between 0 and 999 [Row :" + CStr(i) + "]", etError: CloseExcel: Exit Sub
        
        '****** Checking for valid RmQty in the Excel File
        If (Not IsNumeric(ReadCell(ADC("WXLSHTNO"), i, 9)) And ReadCell(ADC("WXLSHTNO"), i, 9) <> "") Or Val(ReadCell(ADC("WXLSHTNO"), i, 9)) < 0 Or Val(ReadCell(ADC("WXLSHTNO"), i, 9)) > 99999 Then _
            DispMsg "Cannot Copy as Invalid Rm Qty Entered, Should be Between 0 and 99999 [Row :" + CStr(i) + "]", etError: CloseExcel: Exit Sub
        
        '****** Checking for valid RmWt in the Excel File
        If (Not IsNumeric(ReadCell(ADC("WXLSHTNO"), i, 10)) And ReadCell(ADC("WXLSHTNO"), i, 10) <> "") Or Val(ReadCell(ADC("WXLSHTNO"), i, 10)) < 0 Or Val(ReadCell(ADC("WXLSHTNO"), i, 10)) > 99999.999 Then _
            DispMsg "Cannot Copy as Invalid Rm Wt Entered, Should be Between 0 and 99999.999 [Row :" + CStr(i) + "]", etError: CloseExcel: Exit Sub
        
        '****** Checking for valid Purchase Rate in the Excel File
        If (Not IsNumeric(ReadCell(ADC("WXLSHTNO"), i, 11)) And ReadCell(ADC("WXLSHTNO"), i, 11) <> "") Or Val(ReadCell(ADC("WXLSHTNO"), i, 11)) < 0 Or Val(ReadCell(ADC("WXLSHTNO"), i, 11)) > 999999.99 Then _
            DispMsg "Cannot Copy as Invalid Purchase Rate Entered, Should be Between 0 and 999999.99 [Row :" + CStr(i) + "]", etError: CloseExcel: Exit Sub
        
        '****** Checking for valid Purchase Amt in the Excel File
        If (Not IsNumeric(ReadCell(ADC("WXLSHTNO"), i, 12)) And ReadCell(ADC("WXLSHTNO"), i, 12) <> "") Or Val(ReadCell(ADC("WXLSHTNO"), i, 12)) < 0 Or Val(ReadCell(ADC("WXLSHTNO"), i, 12)) > 9999999.99 Then _
            DispMsg "Cannot Copy as Invalid Purchase Amount Entered, Should be Between 0 and 9999999.99 [Row :" + CStr(i) + "]", etError: CloseExcel: Exit Sub
        
        If Trim(ReadCell(ADC("WXLSHTNO"), i, 17)) <> "" Then
           If Not IsNumeric(ReadCell(ADC("WXLSHTNO"), i, 17)) Then DispMsg "Cannot Copy as Best Use OdNo should be a numeric value [Row :" + CStr(i) + "]", etError: CloseExcel: Exit Sub
           'kbs.132 OdSr added in excel template and checking numeric values
           If Not IsNumeric(ReadCell(ADC("WXLSHTNO"), i, 18)) Then DispMsg "Cannot Copy as Best Use OdSr should be a numeric value [Row :" + CStr(i) + "]", etError: CloseExcel: Exit Sub
                      
           If ReadCell(ADC("WXLSHTNO"), i, 15) > 0 Then
           'indigo.164 Trim(ReadCell(ADC("WXLSHTNO"), i, 12)) <> "OS" added
                If Trim(ReadCell(ADC("WXLSHTNO"), i, 14)) <> "" And Trim(ReadCell(ADC("WXLSHTNO"), i, 14)) <> "SO" And Trim(ReadCell(ADC("WXLSHTNO"), i, 14)) <> "OS" Then _
                    DispMsg "Cannot Copy as Best Use OdTc should be 'SO' Or 'OS' Only [Row :" + CStr(i) + "]", etError: CloseExcel: Exit Sub
           'indigo.164 'SO' replaced with ReadCell(ADC("WXLSHTNO"), i, 12)
                If Not moCn.RecSeek("Select 'x' From OrdMst where OmCoCd='" + ADC("TCoCd") + "' and OmTc='" + ReadCell(ADC("WXLSHTNO"), i, 14) + "' and OmYy='" + ReadCell(ADC("WXLSHTNO"), i, 15) + _
                        "' and OmChr='" + ReadCell(ADC("WXLSHTNO"), i, 16) + "' and OmNo=" + ReadCell(ADC("WXLSHTNO"), i, 17)) Then _
                    DispMsg "Cannot Copy as Best Use Voucher  does not exist  [Row :" + CStr(i) + "]", etError: CloseExcel: Exit Sub
                'kbs.132 checking valid OdSr
                'indigo.164 'SO' replaced with ReadCell(ADC("WXLSHTNO"), i, 12)
                If Not moCn.RecSeek("Select 'x' From OrdDsg where OdCoCd='" + ADC("TCoCd") + "' and OdTc='" + ReadCell(ADC("WXLSHTNO"), i, 14) + "' and OdYy='" + ReadCell(ADC("WXLSHTNO"), i, 15) + _
                        "' and OdChr='" + ReadCell(ADC("WXLSHTNO"), i, 16) + "' and OdNo=" + ReadCell(ADC("WXLSHTNO"), i, 17) + " and OdSr=" + ReadCell(ADC("WXLSHTNO"), i, 18)) Then _
                    DispMsg "Cannot Copy as Best Use Voucher, Sr. No. does not exist  [Row :" + CStr(i) + "]", etError: CloseExcel: Exit Sub
                'indigo.164 DmTcTyp='DM' removed

                If ReadCell(ADC("WXLSHTNO"), i, 19) <> "" And Not moCn.RecSeek("Select 'x' From DsgMst where DmTcTyp='DM' and DmCd='" + ReadCell(ADC("WXLSHTNO"), i, 19) + "' and DmSz=''") Then _
                    DispMsg "Cannot Copy as Best Use Design Code does not exist  [Row :" + CStr(i) + "]", etError: CloseExcel: Exit Sub
           End If
            
        End If
       
        ' ***** Manali 3.9.1 - 21/02/12 - qSr added for Order By according to data entered in Excel
       ' moCn.Execute ("Insert Into #TmpXls Values (" + CStr(i) + ", '" + wRmCd + "', " + Str(wRmSz) + ", '" + CStr(ReadCell(adc("WXLSHTNO"), i, 4)) + "'," + Str(wStkRt) + ", " + Str(Val(ReadCell(adc("WXLSHTNO"), i, 6))) + ", " + Str(Val(ReadCell(adc("WXLSHTNO"), i, 7))) + ", " + Str(Val(ReadCell(adc("WXLSHTNO"), i, 8))) + ", " + Str(Val(ReadCell(adc("WXLSHTNO"), i, 9))) + " , " + Str(Val(ReadCell(adc("WXLSHTNO"), i, 10))) + ", Left('" + ReadCell(adc("WXLSHTNO"), i, 11) + "', 60))")
       'kbs.132 inserting OdSr value
       '577-err ReadCell(ADC("WXLSHTNO"), i, 4) changed as wLotNo
       moCn.Execute ("Insert Into #TmpXls Values (" + CStr(i) + ", '" + wRmCd + "', " + Str(wRmSz) + "," + Str(wRmSz2) + "," + Str(wRmSz3) + ", '" + wLotNo + _
                        "'," + Str(wStkRt) + ", " + Str(Val(ReadCell(ADC("WXLSHTNO"), i, 8))) + ", " + Str(Val(ReadCell(ADC("WXLSHTNO"), i, 9))) + ", " + _
                        Str(Val(ReadCell(ADC("WXLSHTNO"), i, 10))) + ", " + Str(Val(ReadCell(ADC("WXLSHTNO"), i, 11))) + " , " + Str(Val(ReadCell(ADC("WXLSHTNO"), i, 12))) + _
                        ", Left('" + ReadCell(ADC("WXLSHTNO"), i, 13) + "', 60)" + ",'" + ReadCell(ADC("WXLSHTNO"), i, 14) + "','" + ReadCell(ADC("WXLSHTNO"), i, 15) + _
                        "','" + ReadCell(ADC("WXLSHTNO"), i, 16) + "'," + Str(Val(ReadCell(ADC("WXLSHTNO"), i, 17))) + "," + Str(Val(ReadCell(ADC("WXLSHTNO"), i, 18))) + ",'" + ReadCell(ADC("WXLSHTNO"), i, 19) + "')")
               
    Next

    CloseExcel
    ' ***** Manali 3.9.1 - 21/02/12 - qSr added for Order By according to data entered in Excel
    Set wo_rsExcel = moCn.OpenRes("Select * From #TmpXls Order By qSr")
    wo_rsExcel.MoveFirst
    
    Dim wRow As Double
    
    If wo_rsExcel.RecCount > 0 Then
      With GrdTxndA
        .StartCopy
        MWLib.BeginProcess Me, "Copying RM Details ..."
        Do While Not wo_rsExcel.EOF
          .AddItem
          wRow = .Rows - 1
          .Row = wRow
          If .Value(wRow, "TdSr") = 0 Then .Value(wRow, "TdSr") = .MaxVal("TdSr") + 1
          .Value(wRow, "TdFrRmLoc") = ADC("TFrRmLoc")
          If ADC("wFrRmIr") = "R" Then
            .Value(wRow, "TdFrRmDc") = "D"
          Else
            .Value(wRow, "TdFrRmDc") = "C"
          End If
          .Value(wRow, "TdPrtKey") = ctCurrPrtn
          .Value(wRow, "TdToRmLoc") = ADC("TToRmLoc")
          .Value(wRow, "wTdFrRmIr") = ADC("wFrRmIr")
          .Value(wRow, "TdRmCd") = wo_rsExcel!TdRmCd
          .Value(wRow, "TdLotNo") = wo_rsExcel!TdLotNo    'Sachin 3.11.0
          .Value(wRow, "TdRmSz") = wo_rsExcel!TdRmSz
          .Value(wRow, "TdRmSz2") = wo_rsExcel!TdRmSz2  '6.1b
          .Value(wRow, "TdRmSz3") = wo_rsExcel!TdRmSz3  '6.1b
          .Value(wRow, "TdRmStkRt") = wo_rsExcel!TDRMSTKRT
          .Value(wRow, "TdPcPerCt") = wo_rsExcel!TdPcPerCt
          .Value(wRow, "TdRmQty") = IIF(wo_rsExcel!TdRmQty = 0 And (FnRmCtg(.Value(wRow, "TdRmCd")) = "D" Or FnRmCtg(.Value(wRow, "TdRmCd")) = "C"), Round(.Value(wRow, "TdPcPerCt") * wo_rsExcel!TdRmWt, 0), wo_rsExcel!TdRmQty)
          .Value(wRow, "TdRmWt") = wo_rsExcel!TdRmWt
          .Value(wRow, "TdDesc") = wo_rsExcel!TdDesc
          .Value(wRow, "TdDustWt") = 0
          If Not mf_OLORYN Then
            .Value(wRow, "TdPurRt") = wo_rsExcel!TdPurRt
          Else
            UpdateTdPurRt (wRow)
          End If
          .Value(wRow, "TdAck") = ADC("TAck")
          UpdateTdRtByQw (wRow)
          If wo_rsExcel!TdPurAmt <> 0 Then
            .Value(wRow, "TdPurAmt") = wo_rsExcel!TdPurAmt
          End If
          If wo_rsExcel!TdBstOdNo <> 0 Then
            'indigo.164 "SO" replaced with wo_rsExcel!TdBstOdTc
            .Value(wRow, "TdBstOdTc") = wo_rsExcel!TdBstOdTc
            .Value(wRow, "TdBstOdYY") = wo_rsExcel!TdBstOdYy
            .Value(wRow, "TdBstOdChr") = wo_rsExcel!TdBstOdChr
            .Value(wRow, "TdBstOdNo") = wo_rsExcel!TdBstOdNo
            .Value(wRow, "TdBstOdSr") = wo_rsExcel!TdBstOdSr
            .Value(wRow, "TdBstDmCd") = wo_rsExcel!TdBstDmCd
          End If
        .SaveRec (wRow)
         wo_rsExcel.MoveNext
        Loop
        
        MWLib.EndProcess Me
        .EndCopy
      End With
    End If

Exit Sub

ErrHdlr:
  If Err.Number = 429 Then
    DispMsg "MS Excel is either Corrupt Or Missing on this PC. Please Install/Re-Install MS Excel ! ", etError
  ElseIf Err.Number = 70 Then       '4.1.4
    DispMsg "Cannot Access Template File. Close If Already Open and Continue.", etError
  Else
    DispMsg Err.Description, etError
  End If


End Sub

Private Sub ImportExcelOld()
On Error GoTo ErrHdlr
    
    Dim wo_rsExcel As New MDORowSet, i As Integer
    
    Dim wRmCtg As String, wStkByRt As String
    Dim wRmCd As String, wRmSz As Double, wStkRt As Double, wPcPerCts As Integer, wRmQty As Integer

    '****** Checking For Valid Inputs From the User ******
    ' ***** Manali 3.9.0 - modified to accept Excel 2007 (.xlsx) files
    If Not OpenExcelFile(ADC("WXLFILE")) Then DispMsg "Improper File Format", etError: Exit Sub
    
    'Call OpenExcelFile(adc("WXLFILE"))
    If ChkSheet(ADC("WXLSHTNO")) = False Then DispMsg "This WorkSheet Number Does Not Exist in the specified Excel File", etError: CloseExcel: Exit Sub
    '****** Checking For Valid Inputs From the User ******

    ' ***** Manali 3.9.1 - 21/02/12 - qSr added for Order By according to data entered in Excel
    moCn.CreateTmpTable "#TmpXls", "Select 0 as qSr, TdRmCd, TdRmSz, TdLotNo,  TdRmStkRt, TdPcPerCt, TdRmQty, TdRmWt, TdPurRt, TdPurAmt, TdDesc,TdBstOdTc,TdBstOdYy,TdBstOdChr,TdBstOdNo,TdBstOdSr,TdBstDmCd Into #TmpXls From Txnd Where 1=2"
        
     ADC("WXLFRROW") = IIF(ADC("WXLFRROW") <= 1, 2, ADC("WXLFRROW"))
    For i = ADC("WXLFRROW") To ADC("WXLTOROW")
        '****** Valid RmCode
        wRmCd = Trim(ReadCell(ADC("WXLSHTNO"), i, 2))
        If Not moCn.RecSeek("Select 'x' from RmMst Where RmCd='" + wRmCd + "' And RmPrtKey='" + ctCurrPrtn + "' and RmValidYN='Y' and RmZ='N' ") Then _
          DispMsg "Cannot Import as RmCode is Not Valid [Row :" + CStr(i) + "]", etError: CloseExcel: Exit Sub
    
        wRmCtg = moCn.GetFldVal("Select RmCtg from RmMst Where RmCd='" + wRmCd + "' ")
        wStkByRt = moCn.GetFldVal("Select PValue1 from RmMst Join Param On PTyp='RMSCTG' and PMCd=RmCtg and PSCd=RmSCTg Where RmCd='" + wRmCd + "' ")
        
        '****** Checking for valid RmCd Size in the Excel File
        'Debug.Print Val(ReadCell(ADC("WXLSHTNO"), i, 3))
        If ReadCell(ADC("WXLSHTNO"), i, 3) <> "" And Not IsNumeric(ReadCell(ADC("WXLSHTNO"), i, 3)) Then
            DispMsg "Cannot Copy as Ln/Sv/Ptr Value is not Number [Row :" + CStr(i) + "]", etError: CloseExcel: Exit Sub
        End If
        
        wRmSz = Val(ReadCell(ADC("WXLSHTNO"), i, 3))
        
'        If wRmSz > 0 And Not (wRmCtg = "C" Or wRmCtg = "D") Then _
 '           DispMsg "Cannot Copy as Ln/Sv/Ptr Should be 0 RmCtg Other Than 'D' and 'C' [Row :" + CStr(i) + "]", etError: CloseExcel: Exit Sub
        
        If wRmSz < 0 Or wRmSz > 999.9999 Then _
            DispMsg "Cannot Copy as Invalid Ln/Sv/Ptr Entered, Should be Between 0 and 999.9999 [Row :" + CStr(i) + "]", etError: CloseExcel: Exit Sub
        Dim wLotNo As String
  '577-err if LotNo module is yes then it reads the lotno value from excel otherwise lotno will be blank
        If moCn.RecSeek("select 'x' from Head where HCd= '" + ctSelfCmCd + "' and HCoCd= '" + ADC("TCoCd") + "' and HLotNoYN= 'N' ") Then
          wLotNo = ""
        Else
          wLotNo = ReadCell(ADC("WXLSHTNO"), i, 4)
        End If
        If Len(wLotNo) > 16 Then _
          DispMsg "Cannot Copy as LotNo Length Should be Lessthan Or Equal to 16 [Row :" + CStr(i) + "]", etError: CloseExcel: Exit Sub


        '****** Checking for valid RmCd StkRt in the Excel File
        If ReadCell(ADC("WXLSHTNO"), i, 5) <> "" And Not IsNumeric(ReadCell(ADC("WXLSHTNO"), i, 5)) Then
            DispMsg "Cannot Copy as Stk Rate Value is not Number [Row :" + CStr(i) + "]", etError: CloseExcel: Exit Sub
        End If
        
        wStkRt = Val(ReadCell(ADC("WXLSHTNO"), i, 5))
        
        If wStkRt > 0 And wStkByRt = "N" Then _
            DispMsg "Cannot Copy as StkRt Should be 0 For This RmCd [Row :" + CStr(i) + "]", etError: CloseExcel: Exit Sub
        
        If wStkRt < -999999.99 Or wStkRt > 999999.99 Then _
            DispMsg "Cannot Copy as Invalid StkRt Entered, Should be Between -999999.99 and 999999.99 [Row :" + CStr(i) + "]", etError: CloseExcel: Exit Sub
            
        '****** Checking for valid Pcs/Cts in the Excel File
        If (Not IsNumeric(ReadCell(ADC("WXLSHTNO"), i, 6)) And ReadCell(ADC("WXLSHTNO"), i, 6) <> "") Or Val(ReadCell(ADC("WXLSHTNO"), i, 6)) < 0 Or Val(ReadCell(ADC("WXLSHTNO"), i, 6)) > 999 Then _
            DispMsg "Cannot Copy as Invalid Pcs/Cts Entered, Should be Between 0 and 999 [Row :" + CStr(i) + "]", etError: CloseExcel: Exit Sub
        
        '****** Checking for valid RmQty in the Excel File
        If (Not IsNumeric(ReadCell(ADC("WXLSHTNO"), i, 7)) And ReadCell(ADC("WXLSHTNO"), i, 7) <> "") Or Val(ReadCell(ADC("WXLSHTNO"), i, 7)) < 0 Or Val(ReadCell(ADC("WXLSHTNO"), i, 7)) > 99999 Then _
            DispMsg "Cannot Copy as Invalid Rm Qty Entered, Should be Between 0 and 99999 [Row :" + CStr(i) + "]", etError: CloseExcel: Exit Sub
        
        '****** Checking for valid RmWt in the Excel File
        If (Not IsNumeric(ReadCell(ADC("WXLSHTNO"), i, 8)) And ReadCell(ADC("WXLSHTNO"), i, 8) <> "") Or Val(ReadCell(ADC("WXLSHTNO"), i, 8)) < 0 Or Val(ReadCell(ADC("WXLSHTNO"), i, 8)) > 99999.999 Then _
            DispMsg "Cannot Copy as Invalid Rm Wt Entered, Should be Between 0 and 99999.999 [Row :" + CStr(i) + "]", etError: CloseExcel: Exit Sub
        
        '****** Checking for valid Purchase Rate in the Excel File
        If (Not IsNumeric(ReadCell(ADC("WXLSHTNO"), i, 9)) And ReadCell(ADC("WXLSHTNO"), i, 9) <> "") Or Val(ReadCell(ADC("WXLSHTNO"), i, 9)) < 0 Or Val(ReadCell(ADC("WXLSHTNO"), i, 9)) > 999999.99 Then _
            DispMsg "Cannot Copy as Invalid Purchase Rate Entered, Should be Between 0 and 999999.99 [Row :" + CStr(i) + "]", etError: CloseExcel: Exit Sub
        
        '****** Checking for valid Purchase Amt in the Excel File
        If (Not IsNumeric(ReadCell(ADC("WXLSHTNO"), i, 10)) And ReadCell(ADC("WXLSHTNO"), i, 10) <> "") Or Val(ReadCell(ADC("WXLSHTNO"), i, 10)) < 0 Or Val(ReadCell(ADC("WXLSHTNO"), i, 10)) > 9999999.99 Then _
            DispMsg "Cannot Copy as Invalid Purchase Amount Entered, Should be Between 0 and 9999999.99 [Row :" + CStr(i) + "]", etError: CloseExcel: Exit Sub
        
        If Trim(ReadCell(ADC("WXLSHTNO"), i, 15)) <> "" Then
           If Not IsNumeric(ReadCell(ADC("WXLSHTNO"), i, 15)) Then DispMsg "Cannot Copy as Best Use OdNo should be a numeric value [Row :" + CStr(i) + "]", etError: CloseExcel: Exit Sub
           'kbs.132 OdSr added in excel template and checking numeric values
           If Not IsNumeric(ReadCell(ADC("WXLSHTNO"), i, 16)) Then DispMsg "Cannot Copy as Best Use OdSr should be a numeric value [Row :" + CStr(i) + "]", etError: CloseExcel: Exit Sub
                      
           If ReadCell(ADC("WXLSHTNO"), i, 15) > 0 Then
           'indigo.164 Trim(ReadCell(ADC("WXLSHTNO"), i, 12)) <> "OS" added
                If Trim(ReadCell(ADC("WXLSHTNO"), i, 12)) <> "" And Trim(ReadCell(ADC("WXLSHTNO"), i, 12)) <> "SO" And Trim(ReadCell(ADC("WXLSHTNO"), i, 12)) <> "OS" Then _
                    DispMsg "Cannot Copy as Best Use OdTc should be 'SO' Or 'OS' Only [Row :" + CStr(i) + "]", etError: CloseExcel: Exit Sub
           'indigo.164 'SO' replaced with ReadCell(ADC("WXLSHTNO"), i, 12)
                If Not moCn.RecSeek("Select 'x' From OrdMst where OmCoCd='" + ADC("TCoCd") + "' and OmTc='" + ReadCell(ADC("WXLSHTNO"), i, 12) + "' and OmYy='" + ReadCell(ADC("WXLSHTNO"), i, 13) + _
                        "' and OmChr='" + ReadCell(ADC("WXLSHTNO"), i, 14) + "' and OmNo=" + ReadCell(ADC("WXLSHTNO"), i, 15)) Then _
                    DispMsg "Cannot Copy as Best Use Voucher  does not exist  [Row :" + CStr(i) + "]", etError: CloseExcel: Exit Sub
                'kbs.132 checking valid OdSr
                'indigo.164 'SO' replaced with ReadCell(ADC("WXLSHTNO"), i, 12)
                If Not moCn.RecSeek("Select 'x' From OrdDsg where OdCoCd='" + ADC("TCoCd") + "' and OdTc='" + ReadCell(ADC("WXLSHTNO"), i, 12) + "' and OdYy='" + ReadCell(ADC("WXLSHTNO"), i, 13) + _
                        "' and OdChr='" + ReadCell(ADC("WXLSHTNO"), i, 14) + "' and OdNo=" + ReadCell(ADC("WXLSHTNO"), i, 15) + " and OdSr=" + ReadCell(ADC("WXLSHTNO"), i, 16)) Then _
                    DispMsg "Cannot Copy as Best Use Voucher, Sr. No. does not exist  [Row :" + CStr(i) + "]", etError: CloseExcel: Exit Sub
                'indigo.164 DmTcTyp='DM' removed

                If ReadCell(ADC("WXLSHTNO"), i, 17) <> "" And Not moCn.RecSeek("Select 'x' From DsgMst where DmTcTyp='DM' and DmCd='" + ReadCell(ADC("WXLSHTNO"), i, 17) + "' and DmSz=''") Then _
                    DispMsg "Cannot Copy as Best Use Design Code does not exist  [Row :" + CStr(i) + "]", etError: CloseExcel: Exit Sub
           End If
            
        End If
       
        ' ***** Manali 3.9.1 - 21/02/12 - qSr added for Order By according to data entered in Excel
       ' moCn.Execute ("Insert Into #TmpXls Values (" + CStr(i) + ", '" + wRmCd + "', " + Str(wRmSz) + ", '" + CStr(ReadCell(adc("WXLSHTNO"), i, 4)) + "'," + Str(wStkRt) + ", " + Str(Val(ReadCell(adc("WXLSHTNO"), i, 6))) + ", " + Str(Val(ReadCell(adc("WXLSHTNO"), i, 7))) + ", " + Str(Val(ReadCell(adc("WXLSHTNO"), i, 8))) + ", " + Str(Val(ReadCell(adc("WXLSHTNO"), i, 9))) + " , " + Str(Val(ReadCell(adc("WXLSHTNO"), i, 10))) + ", Left('" + ReadCell(adc("WXLSHTNO"), i, 11) + "', 60))")
       'kbs.132 inserting OdSr value
       moCn.Execute ("Insert Into #TmpXls Values (" + CStr(i) + ", '" + wRmCd + "', " + Str(wRmSz) + ", '" + CStr(ReadCell(ADC("WXLSHTNO"), i, 4)) + _
                        "'," + Str(wStkRt) + ", " + Str(Val(ReadCell(ADC("WXLSHTNO"), i, 6))) + ", " + Str(Val(ReadCell(ADC("WXLSHTNO"), i, 7))) + ", " + _
                        Str(Val(ReadCell(ADC("WXLSHTNO"), i, 8))) + ", " + Str(Val(ReadCell(ADC("WXLSHTNO"), i, 9))) + " , " + Str(Val(ReadCell(ADC("WXLSHTNO"), i, 10))) + _
                        ", Left('" + ReadCell(ADC("WXLSHTNO"), i, 11) + "', 60)" + ",'" + ReadCell(ADC("WXLSHTNO"), i, 12) + "','" + ReadCell(ADC("WXLSHTNO"), i, 13) + _
                        "','" + ReadCell(ADC("WXLSHTNO"), i, 14) + "'," + Str(Val(ReadCell(ADC("WXLSHTNO"), i, 15))) + "," + Str(Val(ReadCell(ADC("WXLSHTNO"), i, 16))) + ",'" + ReadCell(ADC("WXLSHTNO"), i, 17) + "')")
        
    Next

    CloseExcel
    ' ***** Manali 3.9.1 - 21/02/12 - qSr added for Order By according to data entered in Excel
    Set wo_rsExcel = moCn.OpenRes("Select * From #TmpXls Order By qSr")
    wo_rsExcel.MoveFirst
    
    Dim wRow As Double
    
    If wo_rsExcel.RecCount > 0 Then
      With GrdTxndA
        .StartCopy
        MWLib.BeginProcess Me, "Copying RM Details ..."
        Do While Not wo_rsExcel.EOF
          .AddItem
          wRow = .Rows - 1
          .Row = wRow
          If .Value(wRow, "TdSr") = 0 Then .Value(wRow, "TdSr") = .MaxVal("TdSr") + 1
          .Value(wRow, "TdFrRmLoc") = ADC("TFrRmLoc")
          If ADC("wFrRmIr") = "R" Then
            .Value(wRow, "TdFrRmDc") = "D"
          Else
            .Value(wRow, "TdFrRmDc") = "C"
          End If
          .Value(wRow, "TdPrtKey") = ctCurrPrtn
          .Value(wRow, "TdToRmLoc") = ADC("TToRmLoc")
          .Value(wRow, "wTdFrRmIr") = ADC("wFrRmIr")
          .Value(wRow, "TdRmCd") = wo_rsExcel!TdRmCd
          .Value(wRow, "TdLotNo") = wo_rsExcel!TdLotNo    'Sachin 3.11.0
          .Value(wRow, "TdRmSz") = wo_rsExcel!TdRmSz
          .Value(wRow, "TdRmStkRt") = wo_rsExcel!TDRMSTKRT
          .Value(wRow, "TdPcPerCt") = wo_rsExcel!TdPcPerCt
          .Value(wRow, "TdRmQty") = IIF(wo_rsExcel!TdRmQty = 0 And (FnRmCtg(.Value(wRow, "TdRmCd")) = "D" Or FnRmCtg(.Value(wRow, "TdRmCd")) = "C"), Round(.Value(wRow, "TdPcPerCt") * wo_rsExcel!TdRmWt, 0), wo_rsExcel!TdRmQty)
          .Value(wRow, "TdRmWt") = wo_rsExcel!TdRmWt
          .Value(wRow, "TdDesc") = wo_rsExcel!TdDesc
          .Value(wRow, "TdDustWt") = 0
          If Not mf_OLORYN Then
            .Value(wRow, "TdPurRt") = wo_rsExcel!TdPurRt
          Else
            UpdateTdPurRt (wRow)
          End If
          .Value(wRow, "TdAck") = ADC("TAck")
          UpdateTdRtByQw (wRow)
          If wo_rsExcel!TdPurAmt <> 0 Then
            .Value(wRow, "TdPurAmt") = wo_rsExcel!TdPurAmt
          End If
          If wo_rsExcel!TdBstOdNo <> 0 Then
            'indigo.164 "SO" replaced with wo_rsExcel!TdBstOdTc
            .Value(wRow, "TdBstOdTc") = wo_rsExcel!TdBstOdTc
            .Value(wRow, "TdBstOdYY") = wo_rsExcel!TdBstOdYy
            .Value(wRow, "TdBstOdChr") = wo_rsExcel!TdBstOdChr
            .Value(wRow, "TdBstOdNo") = wo_rsExcel!TdBstOdNo
            .Value(wRow, "TdBstOdSr") = wo_rsExcel!TdBstOdSr
            .Value(wRow, "TdBstDmCd") = wo_rsExcel!TdBstDmCd
          End If
        .SaveRec (wRow)
         wo_rsExcel.MoveNext
        Loop
        
        MWLib.EndProcess Me
        .EndCopy
      End With
    End If

Exit Sub

ErrHdlr:
  If Err.Number = 429 Then
    DispMsg "MS Excel is either Corrupt Or Missing on this PC. Please Install/Re-Install MS Excel ! ", etError
  ElseIf Err.Number = 70 Then       '4.1.4
    DispMsg "Cannot Access Template File. Close If Already Open and Continue.", etError
  Else
    DispMsg Err.Description, etError
  End If


End Sub
'*************bhavna - for imp excel for cnv
Private Sub ImportExcelCnv()
On Error GoTo ErrHdlr
    
    Dim wo_rsExcel As New MDORowSet, i As Integer
    
    Dim wRmCtg As String, wStkByRt As String
    Dim wRmCd As String, wRmSz As Double, wStkRt As Double, wPcPerCts As Integer, wRmQty As Integer

    '****** Checking For Valid Inputs From the User ******
    ' ***** Manali 3.9.0 - modified to accept Excel 2007 (.xlsx) files
    If Not OpenExcelFile(ADC("WXLFILE")) Then DispMsg "Improper File Format", etError: Exit Sub
    
    'Call OpenExcelFile(adc("WXLFILE"))
    If ChkSheet(ADC("WXLSHTNO")) = False Then DispMsg "This WorkSheet Number Does Not Exist in the specified Excel File", etError: CloseExcel: Exit Sub
    '****** Checking For Valid Inputs From the User ******

    ' ***** Manali 3.9.1 - 21/02/12 - qSr added for Order By according to data entered in Excel
    moCn.CreateTmpTable "#TmpXls", "Select 0 as qSr, TdFrRmLoc, '' as wTdFrRmIr, TdRmCd, TdLotNo, TdRmSz, TdRmStkRt, TdRmQty, TdRmWt, TdDustWt Into #TmpXls From Txnd Where 1=2"
        
     ADC("WXLFRROW") = IIF(ADC("WXLFRROW") <= 1, 2, ADC("WXLFRROW"))
    For i = ADC("WXLFRROW") To ADC("WXLTOROW")
    
        '****** valid fr rm loc 'sv.10 R location type added
        If Not moCn.RecSeek("Select LocCd From Loc Where " + _
                           "LocCoCd = '" + ADC("TCoCd") + "' and " + _
                           "LocTyp<>'P' and LocTyp<>'R' and LocCd = '" + Trim(ReadCell(ADC("WXLSHTNO"), i, 2)) + "' And LocValidYN='Y'") Then _
                            DispMsg "Cannot Import as Invalid Raw Material From Location [Row :" + CStr(i) + "]", etError: CloseExcel: Exit Sub

        '****** valid rm I/R
        If Not (Trim(ReadCell(ADC("WXLSHTNO"), i, 3)) = "I" Or Trim(ReadCell(ADC("WXLSHTNO"), i, 3)) = "R") Then _
                            DispMsg "Cannot Import as Invalid Raw Material I/R [Row :" + CStr(i) + "]", etError: CloseExcel: Exit Sub
        '****** Valid RmCode
        wRmCd = Trim(ReadCell(ADC("WXLSHTNO"), i, 4))
        If Not moCn.RecSeek("Select 'x' from RmMst Where RmCd='" + wRmCd + "' And RmPrtKey='" + ctCurrPrtn + "' and RmValidYN='Y' and RmZ='N' ") Then _
          DispMsg "Cannot Import as RmCode is Not Valid [Row :" + CStr(i) + "]", etError: CloseExcel: Exit Sub
    
        wRmCtg = moCn.GetFldVal("Select RmCtg from RmMst Where RmCd='" + wRmCd + "' ")
'        wStkByRt = moCn.GetFldVal("Select PValue1 from RmMst Join Param On PTyp='RMSCTG' and PMCd=RmCtg and PSCd=RmSCTg Where RmCd='" + wRmCd + "' ")
        
        '****** Checking for valid RmCd Size in the Excel File
        'Debug.Print Val(ReadCell(ADC("WXLSHTNO"), i, 3))
        If ReadCell(ADC("WXLSHTNO"), i, 6) <> "" And Not IsNumeric(ReadCell(ADC("WXLSHTNO"), i, 6)) Then
            DispMsg "Cannot Copy as Ln/Sv/Ptr Value is not Number [Row :" + CStr(i) + "]", etError: CloseExcel: Exit Sub
        End If
        
        wRmSz = Val(ReadCell(ADC("WXLSHTNO"), i, 6))
        
        If wRmSz > 0 And Not (wRmCtg = "C" Or wRmCtg = "D" Or wRmCtg = "X") Then _
            DispMsg "Cannot Copy as Ln/Sv/Ptr Should be 0 RmCtg Other Than 'D' 'X' and 'C' [Row :" + CStr(i) + "]", etError: CloseExcel: Exit Sub
        
        If wRmSz < 0 Or wRmSz > 999.9999 Then _
            DispMsg "Cannot Copy as Invalid Ln/Sv/Ptr Entered, Should be Between 0 and 999.9999 [Row :" + CStr(i) + "]", etError: CloseExcel: Exit Sub
        
        
        '****** Checking for valid RmCd StkRt in the Excel File
        If ReadCell(ADC("WXLSHTNO"), i, 7) <> "" And Not IsNumeric(ReadCell(ADC("WXLSHTNO"), i, 7)) Then
            DispMsg "Cannot Copy as Stk Rate Value is not Number [Row :" + CStr(i) + "]", etError: CloseExcel: Exit Sub
        End If
        
        wStkRt = Val(ReadCell(ADC("WXLSHTNO"), i, 7))
        
        If wStkRt > 0 And wStkByRt = "N" Then _
            DispMsg "Cannot Copy as StkRt Should be 0 For This RmCd [Row :" + CStr(i) + "]", etError: CloseExcel: Exit Sub
        
        If wStkRt < -999999.99 Or wStkRt > 999999.99 Then _
            DispMsg "Cannot Copy as Invalid StkRt Entered, Should be Between -999999.99 and 999999.99 [Row :" + CStr(i) + "]", etError: CloseExcel: Exit Sub
            
        '****** Checking for valid RmQty in the Excel File
        If (Not IsNumeric(ReadCell(ADC("WXLSHTNO"), i, 8)) And ReadCell(ADC("WXLSHTNO"), i, 8) <> "") Or Val(ReadCell(ADC("WXLSHTNO"), i, 8)) < 0 Or Val(ReadCell(ADC("WXLSHTNO"), i, 8)) > 99999 Then _
            DispMsg "Cannot Copy as Invalid Rm Qty Entered, Should be Between 0 and 99999 [Row :" + CStr(i) + "]", etError: CloseExcel: Exit Sub
        
        '****** Checking for valid RmWt in the Excel File
        If (Not IsNumeric(ReadCell(ADC("WXLSHTNO"), i, 9)) And ReadCell(ADC("WXLSHTNO"), i, 9) <> "") Or Val(ReadCell(ADC("WXLSHTNO"), i, 9)) < 0 Or Val(ReadCell(ADC("WXLSHTNO"), i, 9)) > 99999.999 Then _
            DispMsg "Cannot Copy as Invalid Rm Wt Entered, Should be Between 0 and 99999.999 [Row :" + CStr(i) + "]", etError: CloseExcel: Exit Sub
        
        '****** Checking for valid Dust Wt in the Excel File
        If (Not IsNumeric(ReadCell(ADC("WXLSHTNO"), i, 10)) And ReadCell(ADC("WXLSHTNO"), i, 10) <> "") Or Val(ReadCell(ADC("WXLSHTNO"), i, 10)) < 0 Or Val(ReadCell(ADC("WXLSHTNO"), i, 10)) > 999999.99 Then _
            DispMsg "Cannot Copy as Invalid Dust Wt Entered, Should be Between 0 and 9999.99 [Row :" + CStr(i) + "]", etError: CloseExcel: Exit Sub
        
        '******* Checking user rights to the location
        If gs_ValidOthLocSeq <> "" Then
           If InStr(1, "," + UCase(gs_ValidOthLocSeq) + ",", "," + _
                        Trim(ReadCell(ADC("WXLSHTNO"), i, 2)) + ",") <= 0 _
           Then DispMsg "RM From Location " + Trim(ReadCell(ADC("WXLSHTNO"), i, 2)) + " is Not Accessible to the User", etError: CloseExcel: Exit Sub
        End If
        If gs_ValidOthLocSeq <> "" Then
           If InStr(1, "," + UCase(gs_ValidOthLocSeq) + ",", "," + _
                        "XC" + ",") <= 0 _
           Then DispMsg "RM To Location 'XC' is Not Accessible to the User", etError: CloseExcel: Exit Sub
        End If
                
       moCn.Execute ("Insert Into #TmpXls Values (" + CStr(i) + ", '" + Trim(ReadCell(ADC("WXLSHTNO"), i, 2)) + "', '" + Trim(ReadCell(ADC("WXLSHTNO"), i, 3)) + _
                        "', '" + wRmCd + "', '" + Trim(ReadCell(ADC("WXLSHTNO"), i, 5)) + "', " + Str(wRmSz) + _
                        ", " + Str(wStkRt) + ", " + Str(Val(ReadCell(ADC("WXLSHTNO"), i, 8))) + ", " + Str(Val(ReadCell(ADC("WXLSHTNO"), i, 9))) + ", " + _
                        Str(Val(ReadCell(ADC("WXLSHTNO"), i, 10))) + ")")
        
    Next

    CloseExcel
    ' ***** Manali 3.9.1 - 21/02/12 - qSr added for Order By according to data entered in Excel
    Set wo_rsExcel = moCn.OpenRes("Select * From #TmpXls Order By qSr")
    wo_rsExcel.MoveFirst
    
    Dim wRow As Double
    
    If wo_rsExcel.RecCount > 0 Then
      With GrdTxndA
        .StartCopy
        MWLib.BeginProcess Me, "Copying RM Details ..."
        Do While Not wo_rsExcel.EOF
          .AddItem
          wRow = .Rows - 1
          .Row = wRow
'          .Value(wRow, "TdCoCd") = adc("TCoCd")
'          .Value(wRow, "TdTc") = adc("TTc")
'          .Value(wRow, "TdYy") = adc("TYy")
'          .Value(wRow, "TdChr") = adc("TChr")
'          .Value(wRow, "TdNo") = adc("TNo")
          If .Value(wRow, "TdSr") = 0 Then .Value(wRow, "TdSr") = .MaxVal("TdSr") + 1
          .Value(wRow, "TdFrRmLoc") = wo_rsExcel!TdFrRmLoc
          If wo_rsExcel!wTdFrRmIr = "R" Then
            .Value(wRow, "TdFrRmDc") = "D"
          Else
            .Value(wRow, "TdFrRmDc") = "C"
          End If
          .Value(wRow, "TdPrtKey") = ctCurrPrtn
          .Value(wRow, "TdToRmLoc") = "XC"
          .Value(wRow, "wTdFrRmIr") = wo_rsExcel!wTdFrRmIr
          .Value(wRow, "TdRmCd") = wo_rsExcel!TdRmCd
          .Value(wRow, "TdLotNo") = wo_rsExcel!TdLotNo
          .Value(wRow, "TdRmSz") = wo_rsExcel!TdRmSz
          .Value(wRow, "TdRmStkRt") = wo_rsExcel!TDRMSTKRT
          .Value(wRow, "TdRmQty") = wo_rsExcel!TdRmQty
          .Value(wRow, "TdRmWt") = wo_rsExcel!TdRmWt
          .Value(wRow, "TdDustWt") = wo_rsExcel!TdDustWt
          .Value(wRow, "TdAck") = "Y"
          UpdateTdRtByQw (wRow)
          If mf_OLORYN Then
            UpdateTdPurRt (wRow)
          End If
          .SaveRec (wRow)
         wo_rsExcel.MoveNext
        Loop
        
        MWLib.EndProcess Me
        .EndCopy
      End With
    End If

Exit Sub

ErrHdlr:
  If Err.Number = 429 Then
    DispMsg "MS Excel is either Corrupt Or Missing on this PC. Please Install/Re-Install MS Excel ! ", etError
  ElseIf Err.Number = 70 Then       '4.1.4
    DispMsg "Cannot Access Template File. Close If Already Open and Continue.", etError
  Else
    DispMsg Err.Description, etError
  End If


End Sub

Private Sub CustzCopy()
'CopyData- copying data from custom purchase
  Dim wRsCustz As MwfLib.MDORowSet
  Dim wRow As Single, wSr As Single
  Dim wTxndFlds As String, wGroupBy As String, wTdRmStkRtYn As String
        
  If Not moCn.RecSeek("Select 'x' From Txn where " + _
     "TCoCd= '" + ADC("wCpTCoCd") + "' and TTc= '" + ADC("wCpTTc") + "' and " + _
     "TYy= '" + ADC("wCpTYy") + "' and TChr= '" + ADC("wCpTChr") + "' and " + _
     "TNo= '" + CStr(ADC("wCpTNo")) + "' and TPrtKey='" + ctCurrPrtn + "'") Then DispMsg "Invalid Customs Voucher", etError: Exit Sub
  
  If Not moCn.RecSeek("Select 'x' From Txn where " + _
     "TCoCd= '" + ADC("wCpTCoCd") + "' and TTc= '" + ADC("wCpTTc") + "' and " + _
     "TYy= '" + ADC("wCpTYy") + "' and TChr= '" + ADC("wCpTChr") + "' and " + _
     "TNo= '" + CStr(ADC("wCpTNo")) + "' and TPrtKey='" + ctCurrPrtn + "' and TCurrCd ='" + ADC("TCurrCd") + "'") Then DispMsg "Source && Target Currency Does Not Match", etError: Exit Sub
         

  If UCase(ADC("wCpGrpOpt")) = "SR" Then
    wTxndFlds = "max(TzRmCdDesc) as TzRmCdDesc, max(TzPurRt) as TzPurRt "
    wGroupBy = "TzSr "
  ElseIf UCase(ADC("wCpGrpOpt")) = "RMCDSZ" Or UCase(ADC("wCpGrpOpt")) = "CDSZLT" Then
    wTxndFlds = "TzRmCdDesc, 0 as TzPurRt "
    wGroupBy = "TzRmCdDesc "
  Else
    wTxndFlds = "TzRmCdDesc, TzPurRt "
    wGroupBy = "TzRmCdDesc, TzPurRt "
  End If
  
  Set wRsCustz = moCn.OpenRes(" Select " + wTxndFlds + ", Sum(TzRmQty) as TotQty, Sum(TzRmWt) as TotWt " + _
                           " From Txndz Where TzRmCdDesc <> '' and TzRmCtg in ('" + wCpGldAs + "') and " + _
                           "TzCoCd= '" + ADC("wCpTCoCd") + "' and " + _
                           "TzTc= '" + ADC("wCpTTc") + "' and " + _
                           "TzYy= '" + ADC("wCpTYy") + "' and " + _
                           "TzChr= '" + ADC("wCpTChr") + "' and " + _
                           "TzNo= " + CStr(ADC("wCpTNo")) + " and TzPrtKey='" + ctCurrPrtn + "'" + _
                           " Group by " + wGroupBy)

  '*** Beginning the copy process
  GrdTxndA.StartCopy
  MWLib.BeginProcess Me, "Copying Purchase Details ..."
    With wRsCustz
    wSr = GrdTxndA.MaxVal("TdSr") + 1
    Do While Not (.EOF Or .BOF)
      MWLib.RunningProcess "Copying Details "
      GrdTxndA.AddItem
      wRow = GrdTxndA.Rows - 1
      GrdTxndA.Row = wRow
      GrdTxndA.Value(wRow, "TdSr") = wSr
      GrdTxndA.Value(wRow, "TdFrRmLoc") = ADC("TFrRmLoc")
      GrdTxndA.Value(wRow, "wTdFrRmIr") = ADC("WFrRmIR")
      If ADC("WFrRmIR") = "R" Then
        GrdTxndA.Value(wRow, "TdFrRmDc") = "D"
      ElseIf ADC("WFrRmIR") = "I" Then
        GrdTxndA.Value(wRow, "TdFrRmDc") = "C"
      End If

      GrdTxndA.Value(wRow, "TdToRmLoc") = ADC("TToRmLoc")
      GrdTxndA.Value(wRow, "TdRmCd") = !TzRmCdDesc
      GrdTxndA.Value(wRow, "TdLotNo") = ""
      GrdTxndA.Value(wRow, "TdRmSz") = ""
      GrdTxndA.Value(wRow, "TdRmSz2") = ""
      GrdTxndA.Value(wRow, "TdRmSz3") = ""
      GrdTxndA.Value(wRow, "TdPurRt") = !TzPurRt
      
      GrdTxndA.Value(wRow, "TdRmQty") = !TotQty
      GrdTxndA.Value(wRow, "TdRmWt") = !TotWt
      GrdTxndA.Value(wRow, "TdPcPerCt") = !TotQty / !TotWt
      GrdTxndA.Value(wRow, "TdDesc") = ADC("WCPTCOCD") + "/" + ADC("WCPTTC") + "/" + ADC("WCPTYY") + "/" + ADC("WCPTCHR") + "/" + CStr(ADC("WCPTNO"))
      'GrdTxndA.Value(wRow, "TdDesc") = ""
      If ms_TcTyp = "MV" Then
      wTdRmStkRtYn = moCn.GetFldVal("Select PValue1 from Param join RmMst on RmCd= '" + _
                 GrdTxndA.Value(wRow, "TdRmCd") + "' and RmPrtKey='" + ctCurrPrtn + "' where PTyp= 'RmSCtg' and PmCd=RmCtg and PScd = RmSCtg ")
                 
      If wTdRmStkRtYn = "Y" Then
          If moCn.GetFldVal("Select hAvgSRtYn from Head where HCoCd= '" + ADC("TCoCd") + "' and " + _
                            "HCd='" + ctSelfCmCd + "'") = "Y" And ms_TTc <> "PR" Then
          Dim wRmLocTyp As String, wb_RmLocTyp As Boolean
            wb_RmLocTyp = moCn.RecSeek("Select LocTyp from Loc where LocCoCd= '" + ADC("TCoCd") + "' " + _
                                       "and LocCd= '" + ADC("TFrRmLoc") + "' " + _
                                       "and LocTyp in " + ctAvgLocTypSeq)
            
            wRmLocTyp = Trim(moCn.GetFldVal("Select LocTyp from Loc where LocCoCd= '" + ADC("TCoCd") + "' " + _
                                       "and LocCd= '" + ADC("TFrRmLoc") + "' "))

            If wb_RmLocTyp = True Then
              GrdTxndA.Value(wRow, "TdRmStkRt") = GetAvgStkRt(ADC("TCoCd"), ADC("TFrRmLoc"), _
                                          !TzRmCdDesc, 0, 0, 0, 0)
            Else
              GrdTxndA.Value(wRow, "TdRmStkRt") = 0
            End If
          Else
            GrdTxndA.Value(wRow, "TdRmStkRt") = 0
          End If
      Else
        GrdTxndA.Value(wRow, "TdRmStkRt") = 0
      End If
      Else
      GrdTxndA.Value(wRow, "TdRmStkRt") = 0
      End If
      'GrdTxndA.Value(wRow, "TdRmStkRt") = 0
      GrdTxndA.Value(wRow, "TdPurAmt") = wRsCustz!TotWt * GrdTxndA.Value(wRow, "TdPurRt")
      GrdTxndA.Value(wRow, "TdDustWt") = 0
      GrdTxndA.Value(wRow, "TdPrtKey") = ctCurrPrtn
      
      GrdTxndA.SaveRec (wRow)
      wSr = wSr + 1
      .MoveNext
    Loop
  End With

  MWLib.EndProcess Me
  GrdTxndA.EndCopy
  Set wRsCustz = Nothing
  
End Sub

Private Sub JtCopy()
'CopyData- Coping Data from jEmr Orders
  Dim wRsOrd As MwfLib.MDORowSet
  Dim wRow As Single, wSr As Single
  Dim wTxndFlds As String, wGroupBy As String, wTdRmStkRtYn As String
  
  If Not moCn.RecSeek("Select 'x' From OrdMst where " + _
     "OmCoCd= '" + ADC("wCpTCoCd") + "' and OmTc= '" + ADC("wCpTTc") + "' and " + _
     "OmYy= '" + ADC("wCpTYy") + "' and OmChr= '" + ADC("wCpTChr") + "' and " + _
     "OmNo= '" + CStr(ADC("wCpTNo")) + "' and OmPrtKey='" + ctCurrPrtn + "'") Then DispMsg "Invalid Order Voucher", etError: Exit Sub
  
  If Not moCn.RecSeek("Select 'x' From OrdMst where " + _
     "OmCoCd= '" + ADC("wCpTCoCd") + "' and OmTc= '" + ADC("wCpTTc") + "' and " + _
     "OmYy= '" + ADC("wCpTYy") + "' and OmChr= '" + ADC("wCpTChr") + "' and " + _
     "OmNo= '" + CStr(ADC("wCpTNo")) + "' and OmPrtKey='" + ctCurrPrtn + "' and OmCmCurCd ='" + ADC("TCurrCd") + "'") Then DispMsg "Source && Target Currency Does Not Match", etError: Exit Sub
   

If wLooseTc = "Y" Then
  If UCase(ADC("wCpGrpOpt")) = "RMCDRT" Then
    wTxndFlds = "OdJLRmCd as TdRmCd, 0 as TdRmStkRt, 0 as TdRmSz, 0 as TdRmSz2,0 as TdRmSz3, '' as TdLotNo, OdJLVchRt as TdPurRt "
    wGroupBy = "OdJLRmCd,OdJLVchRt "
  ElseIf UCase(ADC("wCpGrpOpt")) = "RMCDSZ" Then
    wTxndFlds = "OdJLRmCd as TdRmCd, 0 as TdRmStkRt, OdJLLn1 as TdRmSz,OdJLLn2 as TdRmSz2,OdJLLn3 as TdRmSz3, '' as TdLotNo, 0 as TdPurRt "
    wGroupBy = "OdJLRmCd, OdJLLn1, OdJLLn2, OdJLLn3 "
  ElseIf UCase(ADC("wCpGrpOpt")) = "SR" Then
    wTxndFlds = "max(OdJLRmCd) as TdRmCd, 0 as TdRmStkRt, max(OdJLLn1) as TdRmSz, max(OdJLLn2) as TdRmSz2, max(OdJLLn3) as TdRmSz3, max(OdJLLotNo) as TdLotNo, max(OdJLVchRt) as TdPurRt "
    wGroupBy = "OdSr "
  ElseIf UCase(ADC("wCpGrpOpt")) = "RMCDSZRT" Then
    wTxndFlds = "OdJLRmCd as TdRmCd, 0 as TdRmStkRt, OdJLLn1 as TdRmSz,OdJLLn2 as TdRmSz2, OdJLLn3 as TdRmSz3, '' as TdLotNo, OdJLVchRt as TdPurRt "
    wGroupBy = "OdJLRmCd, OdJLLn1,OdJLLn2, OdJLLn3,OdJLVchRt "
  ElseIf UCase(ADC("wCpGrpOpt")) = "CDRTLT" Then
    wTxndFlds = "OdJLRmCd as TdRmCd, 0 as TdRmStkRt, OdJLLotNo as TdLotNo, 0 as TdRmSz,0 as TdRmSz2,0 as TdRmSz3, OdJLVchRt as TdPurRt "
    wGroupBy = "OdJLRmCd, OdJLVchRt, OdJLLotNo "
  ElseIf UCase(ADC("wCpGrpOpt")) = "CDSZLT" Then
    wTxndFlds = "OdJLRmCd as TdRmCd, OdJLLn1 as TdRmSz,OdJLLn2 as TdRmSz2, OdJLLn3 as TdRmSz3, OdJLLotNo as TdLotNo, 0 as TdRmStkRt, 0 as TdPurRt "
    wGroupBy = "OdJLRmCd, OdJLLn1, OdJLLn2, OdJLLn3, OdJLLotNo "
  ElseIf UCase(ADC("wCpGrpOpt")) = "CDSZRTLT" Then
    wTxndFlds = "OdJLRmCd as TdRmCd, OdJLLn1 as TdRmSz,OdJLLn2 as TdRmSz2, OdJLLn3 as TdRmSz3, 0 as TdRmStkRt, OdJLLotNo as TdLotNo, OdJLVchRt as TdPurRt "
    wGroupBy = "OdJLRmCd, OdJLLn1, OdJLLn2, OdJLLn3,OdJLVchRt, OdJLLotNo "
  End If
 
  Set wRsOrd = moCn.OpenRes(" Select " + wTxndFlds + ", Sum(OdJLQty) as TotQty,Sum(OdJLWt) as TotWt " + _
                            " from OrdDsg Join RmMst on RmCd=OdJLRmCd and RmCtg in ('" + wCpGldAs + "')" + _
                            " Where OdJLRmCd <> '' and " + _
                            " OdCoCd= '" + ADC("wCpTCoCd") + "' and " + _
                            "   OdTc= '" + ADC("wCpTTc") + "' and " + _
                            "   OdYy= '" + ADC("wCpTYy") + "' and " + _
                            "   OdChr= '" + ADC("wCpTChr") + "' and " + _
                            "   OdNo= '" + CStr(ADC("wCpTNo")) + "' and OdPrtKey='" + ctCurrPrtn + "'" + _
                            " Group by " + wGroupBy)
    
End If

If wLooseTc = "N" Then
'Copy- While copying from 'J' entry, if for the given RmCd the Stock By is "Ptr" then get the TRmSz from the OrRmPtr
  If UCase(ADC("wCpGrpOpt")) = "RMCDRT" Then
    wTxndFlds = "OrRmCd as TdRmCd, 0 as TdRmStkRt, 0 as TdRmSz,0 as TdRmSz2,0 as TdRmSz3, '' as TdLotNo, " + IIF(OptRt(0).Value = True, "OrCstRt", "OrSalRt") + " as TdPurRt "
    wGroupBy = "OrRmCd, " + IIF(OptRt(0).Value = True, "OrCstRt", "OrSalRt")
  ElseIf UCase(ADC("wCpGrpOpt")) = "RMCDSZ" Then
    wTxndFlds = "OrRmCd as TdRmCd, 0 as TdRmStkRt, (case when max(PValue3)='Y' then max(OrRmPtr) else max(OrLn1) end) as TdRmSz,(case when max(PValue3)='Y' then 0 else max(OrLn2) end) as TdRmSz2,(case when max(PValue3)='Y' then 0 else max(OrLn3) end) as TdRmSz3, '' as TdLotNo, 0 as TdPurRt "
    wGroupBy = "OrRmCd, (case when PValue3='Y' then OrRmPtr else OrLn1 end), (case when PValue3='Y' then 0 else OrLn2 end), (case when PValue3='Y' then 0 else OrLn3 end)"
  ElseIf UCase(ADC("wCpGrpOpt")) = "SR" Then
    wTxndFlds = "max(OrRmCd) as TdRmCd, 0 as TdRmStkRt, (case when max(PValue3)='Y' then max(OrRmPtr) else max(OrLn1) end) as TdRmSz, (case when max(PValue3)='Y' then 0 else max(OrLn2) end) as TdRmSz2, (case when max(PValue3)='Y' then 0 else max(OrLn3) end) as TdRmSz3, '' as TdLotNo, " + IIF(OptRt(0).Value = True, "max(OrCstRt)", "max(OrSalRt)") + " as TdPurRt "
    wGroupBy = "OrSr,OrSrNo "
  ElseIf UCase(ADC("wCpGrpOpt")) = "RMCDSZRT" Then
    wTxndFlds = "OrRmCd as TdRmCd, 0 as TdRmStkRt, (case when max(PValue3)='Y' then max(OrRmPtr) else max(OrLn1) end) as TdRmSz,(case when max(PValue3)='Y' then 0 else max(OrLn2) end) as TdRmSz2, (case when max(PValue3)='Y' then 0 else max(OrLn3) end) as TdRmSz3, '' as TdLotNo, " + IIF(OptRt(0).Value = True, "OrCstRt", "OrSalRt") + " as TdPurRt "
    wGroupBy = "OrRmCd, (case when PValue3='Y' then OrRmPtr else OrLn1 end),(case when PValue3='Y' then 0 else OrLn2 end), (case when PValue3='Y' then 0 else OrLn3 end), " + IIF(OptRt(0).Value = True, "OrCstRt", "OrSalRt")
  ElseIf UCase(ADC("wCpGrpOpt")) = "CDRTLT" Then
    wTxndFlds = "OrRmCd as TdRmCd, 0 as TdRmStkRt, '' as TdLotNo, 0 as TdRmSz,0 as TdRmSz2,0 as TdRmSz3, " + IIF(OptRt(0).Value = True, "OrCstRt", "OrSalRt") + " as TdPurRt "
    wGroupBy = "OrRmCd, " + IIF(OptRt(0).Value = True, "OrCstRt", "OrSalRt")
  ElseIf UCase(ADC("wCpGrpOpt")) = "CDSZLT" Then
    wTxndFlds = "OrRmCd as TdRmCd, (case when max(PValue3)='Y' then max(OrRmPtr) else max(OrLn1) end) as TdRmSz,(case when max(PValue3)='Y' then 0 else max(OrLn2) end) as TdRmSz2,(case when max(PValue3)='Y' then 0 else max(OrLn3) end) as TdRmSz3, '' as TdLotNo, 0 as TdRmStkRt, 0 as TdPurRt "
    wGroupBy = "OrRmCd, (case when PValue3='Y' then OrRmPtr else OrLn1 end), (case when PValue3='Y' then 0 else OrLn2 end), (case when PValue3='Y' then 0 else OrLn3 end) "
  ElseIf UCase(ADC("wCpGrpOpt")) = "CDSZRTLT" Then
    wTxndFlds = "OrRmCd as TdRmCd, (case when max(PValue3)='Y' then max(OrRmPtr) else max(OrLn1) end) as TdRmSz,(case when max(PValue3)='Y' then 0 else max(OrLn2) end) as TdRmSz2, (case when max(PValue3)='Y' then 0 else max(OrLn3) end) as TdRmSz3, 0 as TdRmStkRt, '' as TdLotNo, " + IIF(OptRt(0).Value = True, "OrCstRt", "OrSalRt") + " as TdPurRt "
    wGroupBy = "OrRmCd, (case when PValue3='Y' then OrRmPtr else OrLn1 end), (case when PValue3='Y' then 0 else OrLn2 end), (case when PValue3='Y' then 0 else OrLn3 end), " + IIF(OptRt(0).Value = True, "OrCstRt", "OrSalRt")
  End If
  
  Set wRsOrd = moCn.OpenRes(" Select " + wTxndFlds + ", Sum(OrQty*OdOrdQty) as TotQty,Sum(OrWt*OdOrdQty) as TotWt " + _
                            " from OrdRm " + _
                            " Join Param on PTyp ='RMSCTG' AND PMCd = OrRmCtg AND PSCd=OrRmSCtg " + _
                            " Join OrdDsg on ODTc=OrTc and ODYy=OrYy and ODChr=OrChr and ODNo=OrNo AND ODSR=OrSr " + _
                            " Where OrRmCd <> '' and OrRmCtg in ('" + wCpGldAs + "') and " + _
                            " OrCoCd= '" + ADC("wCpTCoCd") + "' and " + _
                            "   OrTc= '" + ADC("wCpTTc") + "' and " + _
                            "   OrYy= '" + ADC("wCpTYy") + "' and " + _
                            "   OrChr= '" + ADC("wCpTChr") + "' and " + _
                            "   OrNo= '" + CStr(ADC("wCpTNo")) + "' and OrPrtKey='" + ctCurrPrtn + "'" + _
                            " Group by " + wGroupBy)
      
End If
  
  GrdTxndA.StartCopy
  MWLib.BeginProcess Me, "Copying Purchase Details ..."
  With wRsOrd
    wSr = GrdTxndA.MaxVal("TdSr") + 1
    Do While Not (.EOF Or .BOF)
      MWLib.RunningProcess "Copying  Details "
      GrdTxndA.AddItem
      wRow = GrdTxndA.Rows - 1
      GrdTxndA.Row = wRow
      GrdTxndA.Value(wRow, "TdSr") = wSr
      GrdTxndA.Value(wRow, "TdFrRmLoc") = ADC("TFrRmLoc")
      GrdTxndA.Value(wRow, "wTdFrRmIr") = ADC("WFrRmIR")
      If ADC("WFrRmIR") = "R" Then
        GrdTxndA.Value(wRow, "TdFrRmDc") = "D"
      ElseIf ADC("WFrRmIR") = "I" Then
        GrdTxndA.Value(wRow, "TdFrRmDc") = "C"
      End If
      
      GrdTxndA.Value(wRow, "TdToRmLoc") = ADC("TToRmLoc")
      GrdTxndA.Value(wRow, "TdRmCd") = !TdRmCd
      GrdTxndA.Value(wRow, "TdLotNo") = !TdLotNo
      GrdTxndA.Value(wRow, "TdRmSz") = !TdRmSz
      GrdTxndA.Value(wRow, "TdRmSz2") = !TdRmSz2
      GrdTxndA.Value(wRow, "TdRmSz3") = !TdRmSz3
      If ms_TcTyp = "CNV" Then
      GrdTxndA.Value(wRow, "TdPurRt") = 0
      Else
      GrdTxndA.Value(wRow, "TdPurRt") = !TdPurRt
      End If
      GrdTxndA.Value(wRow, "TdRmQty") = !TotQty
      GrdTxndA.Value(wRow, "TdRmWt") = !TotWt
      GrdTxndA.Value(wRow, "TdPcPerCt") = !TotQty / !TotWt
      GrdTxndA.Value(wRow, "TdDesc") = ADC("WCPTCOCD") + "/" + ADC("WCPTTC") + "/" + ADC("WCPTYY") + "/" + ADC("WCPTCHR") + "/" + CStr(ADC("WCPTNO"))
      'GrdTxndA.Value(wRow, "TdDesc") = ""
      'GrdTxndA.Value(wRow, "TdRmStkRt") = !TdRmStkRt
      If ms_TcTyp = "MV" Then
      wTdRmStkRtYn = moCn.GetFldVal("Select PValue1 from Param join RmMst on RmCd= '" + _
                 GrdTxndA.Value(wRow, "TdRmCd") + "' and RmPrtKey='" + ctCurrPrtn + "' where PTyp= 'RmSCtg' and PmCd=RmCtg and PScd = RmSCtg ")
                 
      If wTdRmStkRtYn = "Y" Then
          If moCn.GetFldVal("Select hAvgSRtYn from Head where HCoCd= '" + ADC("TCoCd") + "' and " + _
                            "HCd='" + ctSelfCmCd + "'") = "Y" And ms_TTc <> "PR" Then
          Dim wRmLocTyp As String, wb_RmLocTyp As Boolean
            wb_RmLocTyp = moCn.RecSeek("Select LocTyp from Loc where LocCoCd= '" + ADC("TCoCd") + "' " + _
                                       "and LocCd= '" + ADC("TFrRmLoc") + "' " + _
                                       "and LocTyp in " + ctAvgLocTypSeq)
            
            wRmLocTyp = Trim(moCn.GetFldVal("Select LocTyp from Loc where LocCoCd= '" + ADC("TCoCd") + "' " + _
                                       "and LocCd= '" + ADC("TFrRmLoc") + "' "))

            If wb_RmLocTyp = True Then
              GrdTxndA.Value(wRow, "TdRmStkRt") = GetAvgStkRt(ADC("TCoCd"), ADC("TFrRmLoc"), _
                                          wRsOrd!TdRmCd, wRsOrd!TdLotNo, wRsOrd!TdRmSz, wRsOrd!TdRmSz2, wRsOrd!TdRmSz3)
            Else
              GrdTxndA.Value(wRow, "TdRmStkRt") = wRsOrd!TDRMSTKRT
            End If
          Else
            GrdTxndA.Value(wRow, "TdRmStkRt") = wRsOrd!TDRMSTKRT
          End If
      Else
        GrdTxndA.Value(wRow, "TdRmStkRt") = 0
      End If
      Else
      GrdTxndA.Value(wRow, "TdRmStkRt") = !TDRMSTKRT
      End If
      GrdTxndA.Value(wRow, "TdPurAmt") = wRsOrd!TotWt * GrdTxndA.Value(wRow, "TdPurRt")
      GrdTxndA.Value(wRow, "TdDustWt") = 0
      GrdTxndA.Value(wRow, "TdPrtKey") = ctCurrPrtn
                        
      GrdTxndA.SaveRec (wRow)
      wSr = wSr + 1
      .MoveNext
    Loop
  End With

  MWLib.EndProcess Me
  GrdTxndA.EndCopy
  Set wRsOrd = Nothing
End Sub

Private Sub GetTotals()
  Dim wTotStkAmt As Double, i As Integer, wTotStkRt As Double
  '578-33 total values in conversion entry
  If UCase(ADC.MenuCd) = UCase("CNV") Then
    ADC("wCtg") = FnRmCtg(GrdTxndA.Value(GrdTxndA.Row, "TdRmCd"))
  
    ADC("wTotNetWt") = FnwTotNetWt
  
    If moCn.GetFldVal("select RmCtg from RmMst where RmCd='" + GrdTxndA.Value(GrdTxndA.Row, "TdRmCd") + "'" + IIF(gs_Partition = ctCurrPrtn, " And RmPrtKey='" + ctCurrPrtn + "'", "")) = "X" Then
      ADC("wNetWt") = FnwTotNetWt(True)
    Else
      ADC("wNetWt") = Round(MWLib.Div(ADC("wTotNetWt"), moCn.GetFldVal("select (case when Round(RmPurityWt, 4)> 0 then RmPurityWt else 1 end)      from RmMst where RmCd='" + GrdTxndA.Value(GrdTxndA.Row, "TdRmCd") + "'" + IIF(gs_Partition = ctCurrPrtn, " And RmPrtKey='" + ctCurrPrtn + "'", ""))), 3)
    End If
  
    
    wTotStkAmt = 0
    For i = 1 To (GrdTxndA.Rows - 1)
      If Not GrdTxndA.IsDel(i) Then wTotStkAmt = wTotStkAmt + GrdTxndA.Value(i, "TdRmStkRt") * GrdTxndA.Value(i, "TdRmWt") * _
                                          IIF(GrdTxndA.Value(i, "TdFrRmDc") = "D", 1, _
                                          IIF(GrdTxndA.Value(i, "TdFrRmDc") = "C", -1, 0))
    Next i
    ADC("wTotStkAmt") = wTotStkAmt
  
  
    ADC("wTotStkRt") = MWLib.Div(ADC("wTotStkAmt"), ADC("wNetWt"))
  End If
    
  '578-33 values not related conversion entry
  Dim wTotQty As Double, wTotWt As Double, wTNetAmt As Double
  If UCase(ADC.MenuCd) <> UCase("CNV") Then
    wTotQty = 0: wTotWt = 0: wTNetAmt = 0
    
    For i = 1 To (GrdTxndA.Rows - 1)
      If Not GrdTxndA.IsDel(i) Then wTotQty = wTotQty + GrdTxndA.Value(i, "TdRmQty") * _
                                              IIF(GrdTxndA.Value(i, "TdFrRmDc") = "D", 1, _
                                              IIF(GrdTxndA.Value(i, "TdFrRmDc") = "C", -1, 0))
            
      If Not GrdTxndA.IsDel(i) Then wTotWt = wTotWt + GrdTxndA.Value(i, "TdRmWt") * _
                                             IIF(GrdTxndA.Value(i, "TdFrRmDc") = "D", 1, _
                                             IIF(GrdTxndA.Value(i, "TdFrRmDc") = "C", -1, 0))
    
      If Not GrdTxndA.IsDel(i) Then wTNetAmt = wTNetAmt + GrdTxndA.Value(i, "TdPurAmt") * _
                                              IIF(GrdTxndA.Value(i, "TdFrRmDc") = "D", 1, _
                                              IIF(GrdTxndA.Value(i, "TdFrRmDc") = "C", -1, 0))
    
    Next i
    
    ADC("wTotQty") = wTotQty
    ADC("wTotWt") = wTotWt
    ADC("TNetAmt") = wTNetAmt
    ADC("TTotAmt") = ADC("TNetAmt") + ADC("TAddAmt1") + ADC("TAddAmt2") + ADC("TAddAmt3") + ADC("TAddAmt4")
  End If
End Sub





