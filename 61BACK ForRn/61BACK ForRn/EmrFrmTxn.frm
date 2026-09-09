VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Object = "{AB2CC32A-E652-42A2-B341-26ECE02F6386}#2.1#0"; "mwfctlforspeed9.ocx"
Begin VB.Form EmrFrmTxn 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Transaction Entry"
   ClientHeight    =   9810
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   15210
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   11040
   ScaleWidth      =   20400
   WindowState     =   2  'Maximized
   Begin MwfCtl.MWCTL_BTN1 CmdTR 
      Height          =   465
      Left            =   10200
      TabIndex        =   232
      ToolTipText     =   "Copy Rm From One Bag To Another "
      Top             =   9270
      Width           =   885
      _ExtentX        =   1561
      _ExtentY        =   820
      BackColor       =   -2147483633
      MaskColor       =   8454143
      Caption         =   "Trans&fer Rm"
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
   Begin MwfCtl.MWCTL_BTN1 CmdHlpSrt 
      Height          =   465
      Left            =   2700
      TabIndex        =   371
      ToolTipText     =   "Option To Enter Help Sort Options"
      Top             =   9270
      Width           =   1005
      _ExtentX        =   1773
      _ExtentY        =   820
      BackColor       =   -2147483633
      MaskColor       =   8454143
      Caption         =   "Help Sort &Option"
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
      Left            =   11820
      TabIndex        =   356
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
   Begin MwfCtl.MWCTL_BTN1 CmdRwrk 
      Height          =   465
      Left            =   8535
      TabIndex        =   310
      ToolTipText     =   "Option To Enter Rework Module"
      Top             =   9270
      Width           =   765
      _ExtentX        =   1349
      _ExtentY        =   820
      BackColor       =   -2147483633
      MaskColor       =   8454143
      Caption         =   "&Rework Entry"
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
   Begin MwfCtl.MWCTL_BTN1 CmdIE 
      Height          =   465
      Left            =   2520
      TabIndex        =   337
      ToolTipText     =   "Copy Bags From One Loc To Another "
      Top             =   9270
      Width           =   735
      _ExtentX        =   1296
      _ExtentY        =   820
      BackColor       =   -2147483633
      MaskColor       =   8454143
      Caption         =   "&2 Import"
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
      Left            =   60
      TabIndex        =   20
      Top             =   9240
      Width           =   15075
      _ExtentX        =   26591
      _ExtentY        =   979
      BtnWidth        =   870
      BtnHeight       =   435
      BtnMaskColor    =   8454143
      FontSz          =   8.25
   End
   Begin MwfCtl.MWCTL_BTN1 CmdCRM 
      Height          =   465
      Left            =   10155
      TabIndex        =   237
      ToolTipText     =   "Copy Raw Material From FG Bag"
      Top             =   9270
      Width           =   795
      _ExtentX        =   1402
      _ExtentY        =   820
      BackColor       =   -2147483633
      MaskColor       =   8454143
      Caption         =   "Cpy R&M FG/Inv"
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
   Begin MwfCtl.MWCTL_BTN1 CmdCert 
      Height          =   465
      Left            =   5520
      TabIndex        =   320
      ToolTipText     =   "Bag Raw Material Summary"
      Top             =   9270
      Width           =   705
      _ExtentX        =   1244
      _ExtentY        =   820
      BackColor       =   -2147483633
      MaskColor       =   8454143
      Caption         =   "&1 Cert Details"
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
   Begin MwfCtl.MWCTL_BTN1 CmdQW 
      CausesValidation=   0   'False
      Height          =   465
      Left            =   6960
      TabIndex        =   235
      ToolTipText     =   "Voucher Summary"
      Top             =   9270
      Width           =   945
      _ExtentX        =   1667
      _ExtentY        =   820
      BackColor       =   -2147483633
      MaskColor       =   8454143
      Caption         =   "&FluteBag Qty/Wt"
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
   Begin MwfCtl.MWCTL_BTN1 CmdGBD 
      Height          =   465
      Left            =   7920
      TabIndex        =   236
      ToolTipText     =   "Gen Bag with Dets"
      Top             =   9270
      Width           =   585
      _ExtentX        =   1032
      _ExtentY        =   820
      BackColor       =   -2147483633
      MaskColor       =   8454143
      Caption         =   "&Gen Bag"
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
   Begin MwfCtl.MWCTL_BTN1 CmdCR 
      Height          =   465
      Left            =   7125
      TabIndex        =   246
      ToolTipText     =   "Copy Raw Material"
      Top             =   9270
      Width           =   840
      _ExtentX        =   1482
      _ExtentY        =   820
      BackColor       =   -2147483633
      MaskColor       =   8454143
      Caption         =   " &Copy  Rm"
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
      Left            =   5520
      TabIndex        =   314
      ToolTipText     =   "Voucher Raw Material Summary"
      Top             =   9270
      Width           =   705
      _ExtentX        =   1244
      _ExtentY        =   820
      Enabled         =   0   'False
      BackColor       =   -2147483633
      MaskColor       =   8454143
      ForeColor       =   11842740
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
   Begin MwfCtl.MWCTL_BTN1 CmdCB 
      Height          =   465
      Left            =   2640
      TabIndex        =   240
      ToolTipText     =   "Copy Bags From One Loc To Another "
      Top             =   9270
      Width           =   615
      _ExtentX        =   1085
      _ExtentY        =   820
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
   Begin MwfCtl.MWCTL_BTN1 CmdLs 
      Height          =   465
      Left            =   9300
      TabIndex        =   238
      ToolTipText     =   "Option To Enter The Loss Record"
      Top             =   9270
      Width           =   615
      _ExtentX        =   1085
      _ExtentY        =   820
      BackColor       =   -2147483633
      MaskColor       =   8454143
      Caption         =   "&Loss Entry"
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
   Begin MwfCtl.MWCTL_BTN1 CmdTot 
      Height          =   465
      Left            =   9540
      TabIndex        =   233
      Top             =   9270
      Width           =   615
      _ExtentX        =   1085
      _ExtentY        =   820
      BackColor       =   -2147483633
      MaskColor       =   8454143
      Caption         =   "&Totals"
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
   Begin MwfCtl.MWCTL_BTN1 CmdTP 
      Height          =   465
      Left            =   3240
      TabIndex        =   234
      ToolTipText     =   "Production Points "
      Top             =   9270
      Width           =   495
      _ExtentX        =   873
      _ExtentY        =   820
      BackColor       =   -2147483633
      MaskColor       =   8454143
      Caption         =   "&Prd Pts"
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
      Left            =   10950
      TabIndex        =   239
      ToolTipText     =   "Copy Bags From One Loc To Another "
      Top             =   9270
      Visible         =   0   'False
      Width           =   975
      _ExtentX        =   1720
      _ExtentY        =   820
      BackColor       =   -2147483633
      MaskColor       =   8454143
      Caption         =   "R&eturn BagsCpy "
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
   Begin MwfCtl.MWCTL_BTN1 CmdER 
      Height          =   465
      Left            =   6840
      TabIndex        =   245
      ToolTipText     =   "Empty The Bag Raw Materials"
      Top             =   9270
      Width           =   840
      _ExtentX        =   1482
      _ExtentY        =   820
      BackColor       =   -2147483633
      MaskColor       =   8454143
      Caption         =   "&Empty Rm"
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
   Begin MwfCtl.MWCTL_BTN1 CmdIR 
      Height          =   465
      Left            =   4200
      TabIndex        =   244
      ToolTipText     =   "Raw Material Issue"
      Top             =   9270
      Width           =   630
      _ExtentX        =   1111
      _ExtentY        =   820
      BackColor       =   -2147483633
      MaskColor       =   8454143
      Caption         =   "&Issue Rm"
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
   Begin MwfCtl.MWCTL_BTN1 CmdBS 
      Height          =   465
      Left            =   4800
      TabIndex        =   243
      ToolTipText     =   "Bag Raw Material Summary"
      Top             =   9270
      Width           =   750
      _ExtentX        =   1323
      _ExtentY        =   820
      BackColor       =   -2147483633
      MaskColor       =   8454143
      Caption         =   "BagRm &Summ"
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
   Begin MwfCtl.MWCTL_BTN1 CmdVS 
      Height          =   465
      Left            =   6240
      TabIndex        =   242
      ToolTipText     =   "Voucher Summary"
      Top             =   9270
      Width           =   735
      _ExtentX        =   1296
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
   Begin VB.Frame FRA_KEYS 
      BorderStyle     =   0  'None
      Height          =   280
      Left            =   -36
      TabIndex        =   52
      Top             =   0
      Width           =   6090
      Begin MwfCtl.MWCTL_MED ATXT 
         Height          =   285
         Index           =   1
         Left            =   1755
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
         Left            =   3645
         TabIndex        =   4
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
         Index           =   5
         Left            =   2970
         TabIndex        =   3
         ToolTipText     =   "Enter Voucher Tc"
         Top             =   0
         Width           =   570
         _ExtentX        =   1005
         _ExtentY        =   503
         MaxLength       =   3
         DataType        =   4
         DataField       =   "TChr"
         ReCalcParent    =   "TFRRMLOC,TTORMLOC,TFRBLOC"
         ReCalcOn        =   "TTC"
         IdName          =   "TCHR"
      End
      Begin MwfCtl.MWCTL_MED ATXT 
         Height          =   285
         Index           =   6
         Left            =   2430
         TabIndex        =   2
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
         Index           =   94
         Left            =   1005
         TabIndex        =   1
         Top             =   0
         Width           =   555
         _ExtentX        =   979
         _ExtentY        =   503
         MaxLength       =   3
         DataType        =   4
         DataField       =   "TCoCd"
         IdName          =   "TCOCD"
      End
      Begin MwfCtl.MWCTL_MED ATXT 
         Height          =   285
         Index           =   95
         Left            =   5745
         TabIndex        =   5
         ToolTipText     =   "Enter Y for your unacknowledged Bags / N for all the unacknowledged bags/ leave Blank to see all the Bags "
         Top             =   0
         Visible         =   0   'False
         Width           =   315
         _ExtentX        =   556
         _ExtentY        =   503
         DataType        =   4
         IdName          =   "WBACKOPT"
      End
      Begin MwfCtl.MWCTL_MED ATXT 
         Height          =   285
         Index           =   139
         Left            =   4590
         TabIndex        =   270
         Top             =   0
         Visible         =   0   'False
         Width           =   345
         _ExtentX        =   609
         _ExtentY        =   503
         Enabled         =   0   'False
         DataField       =   "TPrtKey"
         IdName          =   "TPRTKEY"
      End
      Begin VB.Label LblBAckOpt 
         BackStyle       =   0  'Transparent
         Caption         =   "AckOpt"
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
         Left            =   4950
         TabIndex        =   269
         Top             =   0
         Visible         =   0   'False
         Width           =   750
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
         Index           =   75
         Left            =   1680
         TabIndex        =   268
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
         Left            =   2880
         TabIndex        =   56
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
         Left            =   3555
         TabIndex        =   55
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
         Left            =   2340
         TabIndex        =   54
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
         TabIndex        =   53
         Top             =   0
         Width           =   945
      End
   End
   Begin MwfCtl.MWCTL_HLP MWCTL_HLP1 
      Height          =   405
      Left            =   2580
      TabIndex        =   19
      TabStop         =   0   'False
      Top             =   9390
      Visible         =   0   'False
      Width           =   405
      _ExtentX        =   714
      _ExtentY        =   714
   End
   Begin MwfCtl.MWCTL_BTN1 CmdBD 
      Height          =   465
      Left            =   3720
      TabIndex        =   241
      ToolTipText     =   "Bag Details "
      Top             =   9270
      Width           =   540
      _ExtentX        =   953
      _ExtentY        =   820
      BackColor       =   -2147483633
      MaskColor       =   8454143
      Caption         =   "Bag &Dets"
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
   Begin VB.Frame FRA_NKEYS 
      BorderStyle     =   0  'None
      Enabled         =   0   'False
      Height          =   9570
      Left            =   0
      TabIndex        =   51
      Top             =   0
      Width           =   15315
      Begin VB.Frame FraBD 
         Height          =   3285
         Left            =   0
         TabIndex        =   40
         Top             =   5745
         Width           =   15105
         Begin VB.Frame FraBagDt 
            Height          =   1335
            Left            =   3660
            TabIndex        =   75
            Top             =   360
            Width           =   1755
            Begin MwfCtl.MWCTL_MED ATXT 
               Height          =   285
               Index           =   34
               Left            =   795
               TabIndex        =   76
               ToolTipText     =   "Bag Quantity"
               Top             =   360
               Width           =   600
               _ExtentX        =   1058
               _ExtentY        =   503
               Alignment       =   2
               Enabled         =   0   'False
               Mask            =   "##0.0"
               MaxLength       =   5
               DataType        =   2
               IdName          =   "WBDBQTY"
            End
            Begin MwfCtl.MWCTL_MED ATXT 
               Height          =   285
               Index           =   35
               Left            =   795
               TabIndex        =   78
               ToolTipText     =   "Bag Part Quantity"
               Top             =   645
               Width           =   360
               _ExtentX        =   635
               _ExtentY        =   503
               Enabled         =   0   'False
               Mask            =   "-#0"
               MaxLength       =   3
               DataType        =   1
               IdName          =   "WBDPTQTY"
            End
            Begin MwfCtl.MWCTL_MED ATXT 
               Height          =   285
               Index           =   45
               Left            =   795
               TabIndex        =   80
               ToolTipText     =   "Bag Gross Weight"
               Top             =   930
               Width           =   870
               _ExtentX        =   1535
               _ExtentY        =   503
               Alignment       =   2
               Enabled         =   0   'False
               Mask            =   "####0.000"
               MaxLength       =   9
               DataType        =   2
               IdName          =   "WBDGRSWT"
            End
            Begin VB.Label LblMain 
               BackStyle       =   0  'Transparent
               BorderStyle     =   1  'Fixed Single
               Caption         =   "  Bag"
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
               TabIndex        =   184
               Top             =   90
               Width           =   1755
            End
            Begin VB.Label ALBL 
               BackStyle       =   0  'Transparent
               Caption         =   "Grs Wt"
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
               Left            =   60
               TabIndex        =   81
               Top             =   930
               Width           =   735
            End
            Begin VB.Label ALBL 
               BackStyle       =   0  'Transparent
               Caption         =   "Prt Qty"
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
               Left            =   60
               TabIndex        =   79
               Top             =   645
               Width           =   705
            End
            Begin VB.Label ALBL 
               BackStyle       =   0  'Transparent
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
               Index           =   6
               Left            =   60
               TabIndex        =   77
               Top             =   360
               Width           =   495
            End
         End
         Begin VB.Frame FraPic3D 
            BorderStyle     =   0  'None
            Height          =   1935
            Left            =   12840
            TabIndex        =   336
            Top             =   720
            Width           =   1452
            Begin VB.Image Pic3D 
               Height          =   1830
               Left            =   30
               Top             =   30
               Width           =   1350
            End
         End
         Begin VB.Frame FraPicHsk 
            BorderStyle     =   0  'None
            Height          =   1935
            Left            =   11040
            TabIndex        =   335
            Top             =   720
            Width           =   1452
            Begin VB.Image PicHsk 
               Height          =   1830
               Left            =   30
               Top             =   30
               Width           =   1350
            End
         End
         Begin VB.Frame FraRm 
            Height          =   2925
            Left            =   7410
            TabIndex        =   87
            Top             =   360
            Width           =   3195
            Begin MwfCtl.MWCTL_MED ATXT 
               Height          =   285
               Index           =   49
               Left            =   2040
               TabIndex        =   88
               ToolTipText     =   "Gold Weight"
               Top             =   615
               Width           =   1020
               _ExtentX        =   1799
               _ExtentY        =   503
               Alignment       =   2
               Enabled         =   0   'False
               Mask            =   "####0.000"
               MaxLength       =   9
               DataType        =   2
               IdName          =   "WBDGWT"
            End
            Begin MwfCtl.MWCTL_MED ATXT 
               Height          =   285
               Index           =   50
               Left            =   2040
               TabIndex        =   89
               ToolTipText     =   "Other Metal Weight"
               Top             =   1725
               Width           =   1020
               _ExtentX        =   1799
               _ExtentY        =   503
               Alignment       =   2
               Enabled         =   0   'False
               Mask            =   "####0.000"
               MaxLength       =   9
               DataType        =   2
               IdName          =   "WBDMWT"
            End
            Begin MwfCtl.MWCTL_MED ATXT 
               Height          =   285
               Index           =   51
               Left            =   2040
               TabIndex        =   90
               ToolTipText     =   "Diamond Weight"
               Top             =   2010
               Width           =   1020
               _ExtentX        =   1799
               _ExtentY        =   503
               Alignment       =   2
               Enabled         =   0   'False
               Mask            =   "###0.000"
               MaxLength       =   8
               DataType        =   2
               IdName          =   "WBDDWT"
            End
            Begin MwfCtl.MWCTL_MED ATXT 
               Height          =   285
               Index           =   52
               Left            =   2040
               TabIndex        =   91
               ToolTipText     =   "Colour Stone Weight"
               Top             =   2295
               Width           =   1020
               _ExtentX        =   1799
               _ExtentY        =   503
               Alignment       =   2
               Enabled         =   0   'False
               Mask            =   "###0.000"
               MaxLength       =   8
               DataType        =   2
               IdName          =   "WBDCWT"
            End
            Begin MwfCtl.MWCTL_MED ATXT 
               Height          =   285
               Index           =   53
               Left            =   2040
               TabIndex        =   92
               ToolTipText     =   "Accessories Weight"
               Top             =   2580
               Width           =   1020
               _ExtentX        =   1799
               _ExtentY        =   503
               Alignment       =   2
               Enabled         =   0   'False
               Mask            =   "###0.000"
               MaxLength       =   8
               DataType        =   2
               IdName          =   "WBDXWT"
            End
            Begin MwfCtl.MWCTL_MED ATXT 
               Height          =   285
               Index           =   54
               Left            =   2040
               TabIndex        =   93
               ToolTipText     =   "Platinum Weight"
               Top             =   900
               Width           =   1020
               _ExtentX        =   1799
               _ExtentY        =   503
               Alignment       =   2
               Enabled         =   0   'False
               Mask            =   "####0.000"
               MaxLength       =   9
               DataType        =   2
               IdName          =   "WBDPWT"
            End
            Begin MwfCtl.MWCTL_MED ATXT 
               Height          =   285
               Index           =   56
               Left            =   1425
               TabIndex        =   94
               ToolTipText     =   "Diamond Quantity"
               Top             =   2010
               Width           =   615
               _ExtentX        =   1085
               _ExtentY        =   503
               Alignment       =   2
               Enabled         =   0   'False
               Mask            =   "####0"
               MaxLength       =   5
               DataType        =   2
               IdName          =   "WBDDQTY"
            End
            Begin MwfCtl.MWCTL_MED ATXT 
               Height          =   285
               Index           =   57
               Left            =   1425
               TabIndex        =   95
               ToolTipText     =   "Colour Stone Quantity"
               Top             =   2295
               Width           =   615
               _ExtentX        =   1085
               _ExtentY        =   503
               Alignment       =   2
               Enabled         =   0   'False
               Mask            =   "####0"
               MaxLength       =   5
               DataType        =   2
               IdName          =   "WBDCQTY"
            End
            Begin MwfCtl.MWCTL_MED ATXT 
               Height          =   285
               Index           =   58
               Left            =   1425
               TabIndex        =   96
               ToolTipText     =   "Accessories Quantity"
               Top             =   2580
               Width           =   615
               _ExtentX        =   1085
               _ExtentY        =   503
               Alignment       =   2
               Enabled         =   0   'False
               Mask            =   "####0"
               MaxLength       =   5
               DataType        =   2
               IdName          =   "WBDXQTY"
            End
            Begin MwfCtl.MWCTL_MED ATXT 
               Height          =   285
               Index           =   118
               Left            =   2040
               TabIndex        =   230
               ToolTipText     =   "Silver Weight"
               Top             =   1170
               Width           =   1020
               _ExtentX        =   1799
               _ExtentY        =   503
               Alignment       =   2
               Enabled         =   0   'False
               Mask            =   "####0.000"
               MaxLength       =   9
               DataType        =   2
               IdName          =   "WBDSWT"
            End
            Begin MwfCtl.MWCTL_MED ATXT 
               Height          =   285
               Index           =   148
               Left            =   2040
               TabIndex        =   295
               ToolTipText     =   "Palladium Weight"
               Top             =   1455
               Width           =   1020
               _ExtentX        =   1799
               _ExtentY        =   503
               Alignment       =   2
               Enabled         =   0   'False
               Mask            =   "####0.000"
               MaxLength       =   9
               DataType        =   2
               IdName          =   "WBDLWT"
            End
            Begin MwfCtl.MWCTL_MED ATXT 
               Height          =   285
               Index           =   176
               Left            =   1020
               TabIndex        =   350
               ToolTipText     =   "Gold Weight"
               Top             =   615
               Width           =   1020
               _ExtentX        =   1799
               _ExtentY        =   503
               Alignment       =   2
               Enabled         =   0   'False
               Mask            =   "####0.000"
               MaxLength       =   9
               DataType        =   2
               IdName          =   "WBDFNDGWT"
            End
            Begin MwfCtl.MWCTL_MED ATXT 
               Height          =   285
               Index           =   177
               Left            =   1020
               TabIndex        =   351
               ToolTipText     =   "Gold Weight"
               Top             =   900
               Width           =   1020
               _ExtentX        =   1799
               _ExtentY        =   503
               Alignment       =   2
               Enabled         =   0   'False
               Mask            =   "####0.000"
               MaxLength       =   9
               DataType        =   2
               IdName          =   "WBDFNDPWT"
            End
            Begin MwfCtl.MWCTL_MED ATXT 
               Height          =   285
               Index           =   178
               Left            =   1020
               TabIndex        =   352
               ToolTipText     =   "Gold Weight"
               Top             =   1185
               Width           =   1020
               _ExtentX        =   1799
               _ExtentY        =   503
               Alignment       =   2
               Enabled         =   0   'False
               Mask            =   "####0.000"
               MaxLength       =   9
               DataType        =   2
               IdName          =   "WBDFNDSWT"
            End
            Begin MwfCtl.MWCTL_MED ATXT 
               Height          =   285
               Index           =   179
               Left            =   1020
               TabIndex        =   353
               ToolTipText     =   "Gold Weight"
               Top             =   1470
               Width           =   1020
               _ExtentX        =   1799
               _ExtentY        =   503
               Alignment       =   2
               Enabled         =   0   'False
               Mask            =   "####0.000"
               MaxLength       =   9
               DataType        =   2
               IdName          =   "WBDFNDLWT"
            End
            Begin MwfCtl.MWCTL_MED ATXT 
               Height          =   285
               Index           =   180
               Left            =   1020
               TabIndex        =   354
               ToolTipText     =   "Gold Weight"
               Top             =   1755
               Width           =   1020
               _ExtentX        =   1799
               _ExtentY        =   503
               Alignment       =   2
               Enabled         =   0   'False
               Mask            =   "####0.000"
               MaxLength       =   9
               DataType        =   2
               IdName          =   "WBDFNDMWT"
            End
            Begin VB.Label ALBL 
               BackStyle       =   0  'Transparent
               Caption         =   "Palladium"
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
               Index           =   97
               Left            =   60
               TabIndex        =   296
               Top             =   1455
               Width           =   975
            End
            Begin VB.Label ALBL 
               BackStyle       =   0  'Transparent
               Caption         =   "Silver"
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
               Index           =   68
               Left            =   60
               TabIndex        =   231
               Top             =   1170
               Width           =   1395
            End
            Begin VB.Label LblMain 
               BackStyle       =   0  'Transparent
               BorderStyle     =   1  'Fixed Single
               Caption         =   "  Raw Material"
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
               TabIndex        =   185
               Top             =   90
               Width           =   3195
            End
            Begin VB.Label ALBL 
               BackStyle       =   0  'Transparent
               Caption         =   "Gold"
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
               Left            =   60
               TabIndex        =   104
               Top             =   615
               Width           =   1395
            End
            Begin VB.Label ALBL 
               BackStyle       =   0  'Transparent
               Caption         =   "Oth Mt"
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
               Left            =   60
               TabIndex        =   103
               Top             =   1725
               Width           =   735
            End
            Begin VB.Label ALBL 
               BackStyle       =   0  'Transparent
               Caption         =   "Dia"
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
               Left            =   60
               TabIndex        =   102
               Top             =   2010
               Width           =   1395
            End
            Begin VB.Label ALBL 
               BackStyle       =   0  'Transparent
               Caption         =   "Col Stn"
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
               Left            =   60
               TabIndex        =   101
               Top             =   2295
               Width           =   1395
            End
            Begin VB.Label ALBL 
               BackStyle       =   0  'Transparent
               Caption         =   "Accessories"
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
               Left            =   60
               TabIndex        =   100
               Top             =   2580
               Width           =   1395
            End
            Begin VB.Label ALBL 
               BackStyle       =   0  'Transparent
               Caption         =   "Qty / Met-X"
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
               Left            =   960
               TabIndex        =   99
               ToolTipText     =   "Qty for DCX / Non ‘-‘ Wt for GPSLM"
               Top             =   360
               Width           =   1095
            End
            Begin VB.Label ALBL 
               BackStyle       =   0  'Transparent
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
               Index           =   34
               Left            =   2370
               TabIndex        =   98
               ToolTipText     =   "Total Wt inclusive of ‘-‘ + Non ‘-‘"
               Top             =   360
               Width           =   345
            End
            Begin VB.Label ALBL 
               BackStyle       =   0  'Transparent
               Caption         =   "Platinum"
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
               Left            =   60
               TabIndex        =   97
               Top             =   900
               Width           =   1395
            End
         End
         Begin VB.Frame FraRj 
            Height          =   1335
            Left            =   5715
            TabIndex        =   82
            Top             =   360
            Width           =   1455
            Begin MwfCtl.MWCTL_MED ATXT 
               Height          =   285
               Index           =   37
               Left            =   765
               TabIndex        =   83
               ToolTipText     =   "Bag Quantity"
               Top             =   360
               Width           =   600
               _ExtentX        =   1058
               _ExtentY        =   503
               Alignment       =   2
               Enabled         =   0   'False
               Mask            =   "##0.0"
               MaxLength       =   5
               DataType        =   2
               IdName          =   "WBDRJQTY"
            End
            Begin MwfCtl.MWCTL_MED ATXT 
               Height          =   285
               Index           =   38
               Left            =   765
               TabIndex        =   85
               ToolTipText     =   "Bag Part Quantity"
               Top             =   645
               Width           =   360
               _ExtentX        =   635
               _ExtentY        =   503
               Enabled         =   0   'False
               Mask            =   "-#0"
               MaxLength       =   3
               DataType        =   1
               IdName          =   "WBDRJPTQTY"
            End
            Begin VB.Label LblMain 
               BackStyle       =   0  'Transparent
               BorderStyle     =   1  'Fixed Single
               Caption         =   "  Rejection"
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
               TabIndex        =   186
               Top             =   90
               Width           =   1455
            End
            Begin VB.Label ALBL 
               BackStyle       =   0  'Transparent
               Caption         =   "Prt Qty"
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
               Left            =   60
               TabIndex        =   86
               Top             =   645
               Width           =   735
            End
            Begin VB.Label ALBL 
               BackStyle       =   0  'Transparent
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
               Index           =   9
               Left            =   60
               TabIndex        =   84
               Top             =   360
               Width           =   675
            End
         End
         Begin MwfCtl.MWCTL_MED ATXT 
            Height          =   285
            Index           =   33
            Left            =   1065
            TabIndex        =   57
            ToolTipText     =   "Bag Number"
            Top             =   600
            Width           =   1680
            _ExtentX        =   2963
            _ExtentY        =   503
            Enabled         =   0   'False
            MaxLength       =   14
            DataType        =   4
            IdName          =   "WBDBAG"
         End
         Begin MwfCtl.MWCTL_MED ATXT 
            Height          =   285
            Index           =   36
            Left            =   1065
            TabIndex        =   59
            ToolTipText     =   "Location"
            Top             =   885
            Width           =   1470
            _ExtentX        =   2593
            _ExtentY        =   503
            Enabled         =   0   'False
            MaxLength       =   8
            DataType        =   4
            IdName          =   "WBDLOC"
         End
         Begin MwfCtl.MWCTL_MED ATXT 
            Height          =   285
            Index           =   39
            Left            =   1065
            TabIndex        =   61
            ToolTipText     =   "Customer Code"
            Top             =   1455
            Width           =   1470
            _ExtentX        =   2593
            _ExtentY        =   503
            Enabled         =   0   'False
            MaxLength       =   8
            DataType        =   4
            IdName          =   "WBDCMCD"
         End
         Begin MwfCtl.MWCTL_MED ATXT 
            Height          =   285
            Index           =   40
            Left            =   4455
            TabIndex        =   63
            ToolTipText     =   "Design Size"
            Top             =   1740
            Width           =   705
            _ExtentX        =   1244
            _ExtentY        =   503
            Enabled         =   0   'False
            MaxLength       =   5
            DataType        =   4
            IdName          =   "WBDDMSZ"
         End
         Begin MwfCtl.MWCTL_MED ATXT 
            Height          =   285
            Index           =   41
            Left            =   1065
            TabIndex        =   64
            ToolTipText     =   "Design Code"
            Top             =   1740
            Width           =   2190
            _ExtentX        =   3863
            _ExtentY        =   503
            Enabled         =   0   'False
            MaxLength       =   15
            DataType        =   4
            IdName          =   "WBDDMCD"
         End
         Begin MwfCtl.MWCTL_MED ATXT 
            Height          =   285
            Index           =   42
            Left            =   5580
            TabIndex        =   68
            ToolTipText     =   "Design Suffix"
            Top             =   1740
            Width           =   1605
            _ExtentX        =   2831
            _ExtentY        =   503
            Enabled         =   0   'False
            MaxLength       =   15
            DataType        =   4
            IdName          =   "WBDDMSFX"
         End
         Begin MwfCtl.MWCTL_MED ATXT 
            Height          =   285
            Index           =   43
            Left            =   1065
            TabIndex        =   69
            ToolTipText     =   "Sales Order "
            Top             =   2025
            Width           =   2670
            _ExtentX        =   4710
            _ExtentY        =   503
            Enabled         =   0   'False
            MaxLength       =   23
            DataType        =   4
            IdName          =   "WBDSO"
         End
         Begin MwfCtl.MWCTL_MED ATXT 
            Height          =   285
            Index           =   44
            Left            =   1065
            TabIndex        =   71
            ToolTipText     =   "Worker"
            Top             =   1170
            Width           =   1470
            _ExtentX        =   2593
            _ExtentY        =   503
            Enabled         =   0   'False
            MaxLength       =   8
            DataType        =   4
            IdName          =   "WBDWRK"
         End
         Begin MwfCtl.MWCTL_MED ATXT 
            Height          =   285
            Index           =   46
            Left            =   1065
            TabIndex        =   73
            ToolTipText     =   "Purchase Order"
            Top             =   2310
            Width           =   2670
            _ExtentX        =   4710
            _ExtentY        =   503
            Enabled         =   0   'False
            MaxLength       =   20
            DataType        =   4
            IdName          =   "WBDPONO"
         End
         Begin MwfCtl.MWCTL_MED ATXT 
            Height          =   285
            Index           =   93
            Left            =   6255
            TabIndex        =   152
            ToolTipText     =   "Worker"
            Top             =   2025
            Width           =   930
            _ExtentX        =   1640
            _ExtentY        =   503
            Enabled         =   0   'False
            MaxLength       =   5
            DataType        =   4
            IdName          =   "WBDKT"
         End
         Begin MwfCtl.MWCTL_MED ATXT 
            Height          =   285
            Index           =   181
            Left            =   4455
            TabIndex        =   358
            ToolTipText     =   "Design Colour"
            Top             =   2025
            Width           =   945
            _ExtentX        =   1667
            _ExtentY        =   503
            Enabled         =   0   'False
            MaxLength       =   5
            DataType        =   4
            IdName          =   "WBDDMCOL"
         End
         Begin VB.Label ALBL 
            BackStyle       =   0  'Transparent
            Caption         =   "DmCol"
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
            Index           =   108
            Left            =   3765
            TabIndex        =   357
            Top             =   2025
            Width           =   645
         End
         Begin VB.Label ALBL 
            BackStyle       =   0  'Transparent
            Caption         =   "Line Sketch"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   -1  'True
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Index           =   107
            Left            =   11400
            TabIndex        =   334
            Top             =   2760
            Width           =   885
         End
         Begin VB.Label ALBL 
            BackStyle       =   0  'Transparent
            Caption         =   "3D Image"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   -1  'True
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Index           =   106
            Left            =   13320
            TabIndex        =   333
            Top             =   2760
            Width           =   795
         End
         Begin VB.Label LblMain 
            BackStyle       =   0  'Transparent
            BorderStyle     =   1  'Fixed Single
            Caption         =   "  Bag Details"
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
            TabIndex        =   183
            Top             =   90
            Width           =   15135
         End
         Begin VB.Label ALBL 
            BackStyle       =   0  'Transparent
            Caption         =   "Karat"
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
            Index           =   52
            Left            =   5670
            TabIndex        =   153
            Top             =   2025
            Width           =   675
         End
         Begin VB.Label ALBL 
            BackStyle       =   0  'Transparent
            Caption         =   "PO No."
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
            Left            =   60
            TabIndex        =   74
            Top             =   2310
            Width           =   765
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
            Height          =   285
            Index           =   17
            Left            =   60
            TabIndex        =   72
            Top             =   1170
            Width           =   945
         End
         Begin VB.Label ALBL 
            BackStyle       =   0  'Transparent
            Caption         =   "SO"
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
            Left            =   60
            TabIndex        =   70
            Top             =   2025
            Width           =   375
         End
         Begin VB.Label ALBL 
            BackStyle       =   0  'Transparent
            Caption         =   "Sfx"
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
            Left            =   5250
            TabIndex        =   67
            Top             =   1740
            Width           =   345
         End
         Begin VB.Label ALBL 
            BackStyle       =   0  'Transparent
            Caption         =   "Sz"
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
            Left            =   3765
            TabIndex        =   66
            Top             =   1740
            Width           =   285
         End
         Begin VB.Label ALBL 
            BackStyle       =   0  'Transparent
            Caption         =   "Dsg"
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
            Left            =   60
            TabIndex        =   65
            Top             =   1740
            Width           =   735
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
            Height          =   285
            Index           =   11
            Left            =   60
            TabIndex        =   62
            Top             =   1455
            Width           =   1005
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
            Index           =   8
            Left            =   60
            TabIndex        =   60
            Top             =   885
            Width           =   705
         End
         Begin VB.Label ALBL 
            BackStyle       =   0  'Transparent
            Caption         =   "Bag No."
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
            Left            =   60
            TabIndex        =   58
            Top             =   600
            Width           =   945
         End
      End
      Begin VB.Frame FraLs 
         Height          =   1935
         Left            =   5505
         TabIndex        =   142
         Top             =   7095
         Width           =   5280
         Begin MwfCtl.MWCTL_BTN1 CmdLsGo 
            Height          =   495
            Left            =   4590
            TabIndex        =   148
            ToolTipText     =   "Press This Button To Loss Record"
            Top             =   1125
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
            Index           =   55
            Left            =   1515
            TabIndex        =   146
            ToolTipText     =   "Enter Gross Wt"
            Top             =   1050
            Width           =   1050
            _ExtentX        =   1852
            _ExtentY        =   503
            Alignment       =   2
            Mask            =   "####0.000"
            MaxLength       =   9
            DataType        =   2
            IdName          =   "WLSGRSWT"
         End
         Begin MwfCtl.MWCTL_MED ATXT 
            Height          =   285
            Index           =   47
            Left            =   1515
            TabIndex        =   147
            ToolTipText     =   "Enter Dust Wt"
            Top             =   1335
            Width           =   1050
            _ExtentX        =   1852
            _ExtentY        =   503
            Alignment       =   2
            Mask            =   "##0.000"
            MaxLength       =   7
            DataType        =   2
            IdName          =   "WLSDUSTWT"
         End
         Begin MwfCtl.MWCTL_MED ATXT 
            Height          =   285
            Index           =   105
            Left            =   1515
            TabIndex        =   144
            ToolTipText     =   "Enter Dust Wt"
            Top             =   765
            Width           =   2250
            _ExtentX        =   3969
            _ExtentY        =   503
            MaxLength       =   16
            DataType        =   4
            ReCalcOn        =   "WLSSECMETALLYN"
            IdName          =   "WLSSECMETCD"
         End
         Begin MwfCtl.MWCTL_MED ATXT 
            Height          =   285
            Index           =   157
            Left            =   4815
            TabIndex        =   145
            ToolTipText     =   "Enter 'Y' For entering Loss For All RmCodes"
            Top             =   765
            Width           =   330
            _ExtentX        =   582
            _ExtentY        =   503
            MaxLength       =   12
            DataType        =   4
            ReCalcParent    =   "WLSSECMETCD"
            IdName          =   "WLSSECMETALLYN"
         End
         Begin MwfCtl.MWCTL_MED ATXT 
            Height          =   285
            Index           =   104
            Left            =   1515
            TabIndex        =   151
            ToolTipText     =   "Enter Gross Wt"
            Top             =   480
            Width           =   2250
            _ExtentX        =   3969
            _ExtentY        =   503
            Enabled         =   0   'False
            MaxLength       =   16
            DataType        =   4
            ReCalcOn        =   "WLSMAINMETALLYN"
            IdName          =   "WLSMAINMETCD"
         End
         Begin MwfCtl.MWCTL_MED ATXT 
            Height          =   285
            Index           =   156
            Left            =   4815
            TabIndex        =   143
            ToolTipText     =   "Enter 'Y' For entering Loss For All RmCodes"
            Top             =   480
            Width           =   330
            _ExtentX        =   582
            _ExtentY        =   503
            MaxLength       =   12
            DataType        =   4
            ReCalcParent    =   "WLSMAINMETCD"
            IdName          =   "WLSMAINMETALLYN"
         End
         Begin VB.Label ALBL 
            BackStyle       =   0  'Transparent
            Caption         =   "All Cd Y/N"
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
            Left            =   3840
            TabIndex        =   301
            Top             =   780
            Width           =   945
         End
         Begin VB.Label ALBL 
            BackStyle       =   0  'Transparent
            Caption         =   "All Cd Y/N"
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
            Left            =   3840
            TabIndex        =   300
            Top             =   480
            Width           =   945
         End
         Begin VB.Label LblMain 
            BackStyle       =   0  'Transparent
            BorderStyle     =   1  'Fixed Single
            Caption         =   "  Main Metal Loss"
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
            TabIndex        =   188
            Top             =   90
            Width           =   5070
         End
         Begin VB.Label ALBL 
            BackStyle       =   0  'Transparent
            Caption         =   "Main Metal Cd"
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
            Index           =   65
            Left            =   60
            TabIndex        =   169
            Top             =   480
            Width           =   1395
         End
         Begin VB.Label ALBL 
            BackStyle       =   0  'Transparent
            Caption         =   "Sec Metal Cd"
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
            Index           =   64
            Left            =   60
            TabIndex        =   168
            Top             =   765
            Width           =   1305
         End
         Begin VB.Label ALBL 
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
            Index           =   1
            Left            =   60
            TabIndex        =   150
            Top             =   1335
            Width           =   1005
         End
         Begin VB.Label ALBL 
            BackStyle       =   0  'Transparent
            Caption         =   "Gross Wt"
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
            Left            =   60
            TabIndex        =   149
            Top             =   1050
            Width           =   1005
         End
      End
      Begin VB.Frame FraBRS 
         Enabled         =   0   'False
         Height          =   3615
         Left            =   165
         TabIndex        =   125
         Top             =   5415
         Width           =   15105
         Begin MwfCtl.MWCTL_MED ATXT 
            Height          =   285
            Index           =   75
            Left            =   3150
            TabIndex        =   126
            Top             =   3210
            Width           =   645
            _ExtentX        =   1138
            _ExtentY        =   503
            Alignment       =   2
            Enabled         =   0   'False
            Mask            =   "#####0"
            MaxLength       =   6
            DataType        =   2
            IdName          =   "WBRSRMTOTQTY"
         End
         Begin MwfCtl.MWCTL_MED ATXT 
            Height          =   285
            Index           =   76
            Left            =   3780
            TabIndex        =   128
            Top             =   3210
            Width           =   915
            _ExtentX        =   1614
            _ExtentY        =   503
            Alignment       =   2
            Enabled         =   0   'False
            Mask            =   "#####0.000"
            MaxLength       =   10
            DataType        =   2
            IdName          =   "WBRSRMTOTWT"
         End
         Begin MwfCtl.MWCTL_MED ATXT 
            Height          =   285
            Index           =   78
            Left            =   8445
            TabIndex        =   129
            ToolTipText     =   "Enter Group By Option"
            Top             =   450
            Width           =   1365
            _ExtentX        =   2408
            _ExtentY        =   503
            MaxLength       =   8
            DataType        =   4
            IdName          =   "WBRSGRP"
         End
         Begin MwfCtl.MWCTL_FLX GrdBRS 
            Height          =   2850
            Left            =   90
            TabIndex        =   127
            Top             =   360
            Width           =   6975
            _ExtentX        =   12303
            _ExtentY        =   5027
            Cols            =   7
            AllowAdd        =   0   'False
            AllowDelete     =   0   'False
            colname1        =   "WBRSRMCTG"
            heading1        =   "Ctg"
            datatype1       =   4
            maxlength1      =   2
            colname2        =   "WBRSRMSCTG"
            heading2        =   "Sub Ctg"
            datatype2       =   4
            maxlength2      =   5
            colname3        =   "WBRSRMCD"
            heading3        =   "Rm Code                                 "
            datatype3       =   4
            maxlength3      =   16
            colname4        =   "WBRSRMQTY"
            heading4        =   "Qty     "
            datatype4       =   1
            mask4           =   "####0"
            maxlength4      =   5
            colname5        =   "WBRSRMWT"
            heading5        =   "Wt.             "
            datatype5       =   2
            mask5           =   "#####0.000"
            maxlength5      =   10
            colname6        =   "WBRSRMVAL"
            heading6        =   "Value            "
            datatype6       =   2
            mask6           =   "######0.00"
            maxlength6      =   10
            style6          =   2
         End
         Begin MwfCtl.MWCTL_MED ATXT 
            Height          =   285
            Index           =   164
            Left            =   8445
            TabIndex        =   130
            ToolTipText     =   "Enter Group By Option"
            Top             =   810
            Width           =   1365
            _ExtentX        =   2408
            _ExtentY        =   503
            MaxLength       =   8
            DataType        =   4
            IdName          =   "WBRSVALTYP"
         End
         Begin MwfCtl.MWCTL_MED ATXT 
            Height          =   285
            Index           =   170
            Left            =   4680
            TabIndex        =   309
            Top             =   3210
            Width           =   1275
            _ExtentX        =   2249
            _ExtentY        =   503
            Alignment       =   2
            Enabled         =   0   'False
            Mask            =   "-#####0.000"
            MaxLength       =   11
            DataType        =   2
            IdName          =   "WBRSRMTOTVAL"
         End
         Begin VB.Label ALBL 
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
            Height          =   285
            Index           =   99
            Left            =   7200
            TabIndex        =   306
            Top             =   810
            Width           =   1245
         End
         Begin VB.Label LblMain 
            BackStyle       =   0  'Transparent
            BorderStyle     =   1  'Fixed Single
            Caption         =   "  Bag Raw Material Summary"
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
            TabIndex        =   180
            Top             =   90
            Width           =   15135
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
            Index           =   29
            Left            =   7200
            TabIndex        =   131
            Top             =   450
            Width           =   1005
         End
      End
      Begin VB.Frame FraVS 
         Enabled         =   0   'False
         Height          =   4455
         Left            =   0
         TabIndex        =   47
         Top             =   4800
         Width           =   15195
         Begin MwfCtl.MWCTL_MED ATXT 
            Height          =   285
            Index           =   62
            Left            =   4380
            TabIndex        =   114
            ToolTipText     =   "Total Raw Material Quantity Issued"
            Top             =   4080
            Width           =   735
            _ExtentX        =   1296
            _ExtentY        =   503
            Alignment       =   2
            Enabled         =   0   'False
            Mask            =   "#####0"
            MaxLength       =   6
            DataType        =   2
            IdName          =   "WVSISSTOTQTY"
         End
         Begin MwfCtl.MWCTL_MED ATXT 
            Height          =   285
            Index           =   63
            Left            =   5100
            TabIndex        =   115
            ToolTipText     =   "Total Raw Material Weight Issued"
            Top             =   4080
            Width           =   975
            _ExtentX        =   1720
            _ExtentY        =   503
            Alignment       =   2
            Enabled         =   0   'False
            Mask            =   "####0.000"
            MaxLength       =   9
            DataType        =   2
            IdName          =   "WVSISSTOTWT"
         End
         Begin MwfCtl.MWCTL_MED ATXT 
            Height          =   285
            Index           =   68
            Left            =   13725
            TabIndex        =   49
            ToolTipText     =   "Enter Group By Field"
            Top             =   600
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
            Left            =   6060
            TabIndex        =   117
            ToolTipText     =   "Total Raw Material Quantity Received"
            Top             =   4080
            Width           =   765
            _ExtentX        =   1349
            _ExtentY        =   503
            Alignment       =   2
            Enabled         =   0   'False
            Mask            =   "#####0"
            MaxLength       =   6
            DataType        =   2
            IdName          =   "WVSRECTOTQTY"
         End
         Begin MwfCtl.MWCTL_MED ATXT 
            Height          =   285
            Index           =   69
            Left            =   6810
            TabIndex        =   118
            ToolTipText     =   "Total Raw Material Weight Received"
            Top             =   4080
            Width           =   1035
            _ExtentX        =   1826
            _ExtentY        =   503
            Alignment       =   2
            Enabled         =   0   'False
            Mask            =   "####0.000"
            MaxLength       =   9
            DataType        =   2
            IdName          =   "WVSRECTOTWT"
         End
         Begin MwfCtl.MWCTL_FLX GrdVS 
            Height          =   3720
            Left            =   120
            TabIndex        =   48
            Top             =   360
            Width           =   12345
            _ExtentX        =   21775
            _ExtentY        =   6562
            Cols            =   12
            AllowAdd        =   0   'False
            AllowDelete     =   0   'False
            colname1        =   "WVSLOC"
            heading1        =   "Loc         "
            datatype1       =   4
            maxlength1      =   5
            tooltiptext1    =   "Location"
            colname2        =   "WVSRMCTG"
            heading2        =   "Ctg "
            datatype2       =   4
            maxlength2      =   2
            tooltiptext2    =   "Raw Materila Category"
            colname3        =   "WVSRMSCTG"
            heading3        =   "Sub Ctg"
            datatype3       =   4
            maxlength3      =   5
            tooltiptext3    =   "Raw Material Sub Category"
            colname4        =   "WVSRMCD"
            heading4        =   "Rm Code                                "
            datatype4       =   4
            maxlength4      =   16
            tooltiptext4    =   "Raw Material Code"
            colname5        =   "WVSLOTNO"
            heading5        =   "Lot No.                       "
            datatype5       =   4
            maxlength5      =   12
            tooltiptext5    =   "Lot No."
            colname6        =   "WVSRMSZ"
            heading6        =   "Size           "
            datatype6       =   2
            mask6           =   "##0.0000"
            maxlength6      =   8
            tooltiptext6    =   "Raw Material Size"
            colname7        =   "WVSISSRMQTY"
            heading7        =   "Iss Qty "
            datatype7       =   1
            mask7           =   "####0"
            maxlength7      =   5
            tooltiptext7    =   "Issued Raw Material Quantity"
            colname8        =   "WVSISSRMWT"
            heading8        =   "Iss Wt.        "
            datatype8       =   2
            mask8           =   "####0.000"
            maxlength8      =   9
            tooltiptext8    =   "Issued Raw Material Weight"
            colname9        =   "WVSRECRMQTY"
            heading9        =   "Rec Qty"
            datatype9       =   1
            mask9           =   "####0"
            maxlength9      =   5
            tooltiptext9    =   "Received Raw Material Quantity"
            colname10       =   "WVSRECRMWT"
            heading10       =   "Rec Wt      "
            datatype10      =   2
            mask10          =   "####0.000"
            maxlength10     =   9
            tooltiptext10   =   "Received Raw Material Weight"
            colname11       =   "WVSRMVAL"
            heading11       =   "Value             "
            datatype11      =   2
            mask11          =   "-######0.00"
            maxlength11     =   11
            style11         =   2
         End
         Begin MwfCtl.MWCTL_MED ATXT 
            Height          =   285
            Index           =   168
            Left            =   13725
            TabIndex        =   50
            ToolTipText     =   "Enter Group By Option"
            Top             =   1080
            Width           =   1365
            _ExtentX        =   2408
            _ExtentY        =   503
            MaxLength       =   8
            DataType        =   4
            IdName          =   "WVSVALTYP"
         End
         Begin MwfCtl.MWCTL_MED ATXT 
            Height          =   285
            Index           =   169
            Left            =   7845
            TabIndex        =   308
            ToolTipText     =   "Total Raw Material Value Received"
            Top             =   4080
            Width           =   1395
            _ExtentX        =   2461
            _ExtentY        =   503
            Alignment       =   2
            Enabled         =   0   'False
            Mask            =   "#######0.00"
            MaxLength       =   11
            DataType        =   2
            IdName          =   "WVSTOTVAL"
         End
         Begin VB.Label ALBL 
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
            Height          =   285
            Index           =   101
            Left            =   12570
            TabIndex        =   307
            Top             =   1080
            Width           =   1245
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
            Index           =   7
            Left            =   -30
            TabIndex        =   187
            Top             =   90
            Width           =   15105
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
            Left            =   12570
            TabIndex        =   116
            Top             =   600
            Width           =   1005
         End
      End
      Begin VB.Frame FraIR 
         Height          =   2325
         Left            =   2700
         TabIndex        =   132
         Top             =   6840
         Width           =   7980
         Begin MwfCtl.MWCTL_BTN1 CmdIRGo 
            Height          =   495
            Left            =   3360
            TabIndex        =   138
            Top             =   1080
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
            Index           =   74
            Left            =   1215
            TabIndex        =   137
            ToolTipText     =   "Enter Raw Material Gross Weight"
            Top             =   1260
            Width           =   810
            _ExtentX        =   1429
            _ExtentY        =   503
            Mask            =   "##0.000"
            MaxLength       =   7
            DataType        =   2
            IdName          =   "WIRRMGRWT"
         End
         Begin MwfCtl.MWCTL_MED ATXT 
            Height          =   285
            Index           =   116
            Left            =   1215
            TabIndex        =   134
            ToolTipText     =   "Enter Lot No."
            Top             =   690
            Width           =   2730
            _ExtentX        =   4815
            _ExtentY        =   503
            MaxLength       =   16
            DataType        =   4
            ReCalcOn        =   "WIRRMCD"
            IdName          =   "WIRLOTNO"
         End
         Begin MwfCtl.MWCTL_MED ATXT 
            Height          =   285
            Index           =   73
            Left            =   1215
            TabIndex        =   135
            ToolTipText     =   "Enter Raw Material Wt"
            Top             =   975
            Width           =   810
            _ExtentX        =   1429
            _ExtentY        =   503
            Mask            =   "##0.000"
            MaxLength       =   7
            DataType        =   2
            IdName          =   "WIRRMADDWT"
         End
         Begin MwfCtl.MWCTL_MED ATXT 
            Height          =   285
            Index           =   150
            Left            =   2025
            TabIndex        =   136
            ToolTipText     =   "Enter Raw Material Qty"
            Top             =   975
            Width           =   840
            _ExtentX        =   1482
            _ExtentY        =   503
            Mask            =   "####0"
            MaxLength       =   5
            DataType        =   1
            IdName          =   "WIRRMADDQTY"
         End
         Begin MwfCtl.MWCTL_MED ATXT 
            Height          =   285
            Index           =   72
            Left            =   1215
            TabIndex        =   133
            ToolTipText     =   "Enter Raw Material Sub Category"
            Top             =   390
            Width           =   2730
            _ExtentX        =   4815
            _ExtentY        =   503
            MaxLength       =   16
            DataType        =   4
            IdName          =   "WIRRMCD"
         End
         Begin MwfCtl.MWCTL_FLX GrdOrLot 
            Height          =   1800
            Left            =   4200
            TabIndex        =   355
            Top             =   360
            Width           =   3675
            _ExtentX        =   6482
            _ExtentY        =   3175
            Cols            =   2
            AllowAdd        =   0   'False
            AllowDelete     =   0   'False
            colname1        =   "WORLOTNO"
            heading1        =   "Rm Lot                                                   "
            datatype1       =   4
            maxlength1      =   16
            style1          =   2
         End
         Begin VB.Label LblLotNo 
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
            Left            =   60
            TabIndex        =   193
            Top             =   690
            Width           =   1185
         End
         Begin VB.Label LblMain 
            BackStyle       =   0  'Transparent
            BorderStyle     =   1  'Fixed Single
            Caption         =   "  Raw Material Issue"
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
            TabIndex        =   182
            Top             =   90
            Width           =   7875
         End
         Begin VB.Label ALBL 
            BackStyle       =   0  'Transparent
            Caption         =   "Rm Cd"
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
            Left            =   60
            TabIndex        =   141
            Top             =   405
            Width           =   1185
         End
         Begin VB.Label ALBL 
            BackStyle       =   0  'Transparent
            Caption         =   "Add Wt/Qty"
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
            Left            =   60
            TabIndex        =   140
            Top             =   975
            Width           =   1185
         End
         Begin VB.Label ALBL 
            BackStyle       =   0  'Transparent
            Caption         =   "Grs Wt"
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
            Left            =   60
            TabIndex        =   139
            Top             =   1260
            Width           =   1185
         End
      End
      Begin VB.Frame FraCB 
         Height          =   3720
         Left            =   360
         TabIndex        =   36
         Top             =   5280
         Width           =   14445
         Begin VB.Frame FraCBBag 
            Caption         =   "Bag Nos / Bag Alias No."
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   3255
            Left            =   4800
            TabIndex        =   434
            Top             =   430
            Width           =   9580
            Begin MwfCtl.MWCTL_MED ATXT 
               Height          =   2895
               Index           =   478
               Left            =   60
               TabIndex        =   435
               ToolTipText     =   "Copy Bag Alias No. / Bag Nos (Yy/Chr/No,Yy/Chr....,Yy/Chr/No)  String"
               Top             =   240
               Width           =   9435
               _ExtentX        =   16642
               _ExtentY        =   5106
               Alignment       =   3
               MaxLength       =   0
               IdName          =   "WRFIDSTR"
            End
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
            TabIndex        =   377
            Top             =   360
            Width           =   4695
            Begin VB.Frame FraCbRej 
               BorderStyle     =   0  'None
               Height          =   585
               Left            =   30
               TabIndex        =   378
               Top             =   2235
               Width           =   3855
               Begin MwfCtl.MWCTL_MED ATXT 
                  Height          =   285
                  Index           =   159
                  Left            =   1890
                  TabIndex        =   380
                  ToolTipText     =   "Enter From Rej Year "
                  Top             =   0
                  Width           =   360
                  _ExtentX        =   635
                  _ExtentY        =   503
                  MaxLength       =   2
                  DataType        =   4
                  IdName          =   "WCBTDYY"
               End
               Begin MwfCtl.MWCTL_MED ATXT 
                  Height          =   285
                  Index           =   160
                  Left            =   3030
                  TabIndex        =   382
                  ToolTipText     =   "Enter From Rej Vch Number"
                  Top             =   0
                  Width           =   780
                  _ExtentX        =   1376
                  _ExtentY        =   503
                  Alignment       =   2
                  Mask            =   "#####0"
                  MaxLength       =   6
                  DataType        =   1
                  IdName          =   "WCBTDNO"
               End
               Begin MwfCtl.MWCTL_MED ATXT 
                  Height          =   285
                  Index           =   161
                  Left            =   1125
                  TabIndex        =   379
                  ToolTipText     =   "Enter From Rej Tc "
                  Top             =   0
                  Width           =   600
                  _ExtentX        =   1058
                  _ExtentY        =   503
                  MaxLength       =   3
                  DataType        =   4
                  IdName          =   "WCBTDTC"
               End
               Begin MwfCtl.MWCTL_MED ATXT 
                  Height          =   285
                  Index           =   162
                  Left            =   2340
                  TabIndex        =   381
                  ToolTipText     =   "Enter From Rej Character"
                  Top             =   0
                  Width           =   600
                  _ExtentX        =   1058
                  _ExtentY        =   503
                  MaxLength       =   3
                  DataType        =   4
                  IdName          =   "WCBTDCHR"
               End
               Begin MwfCtl.MWCTL_MED ATXT 
                  Height          =   285
                  Index           =   163
                  Left            =   1125
                  TabIndex        =   383
                  ToolTipText     =   "Enter From Rejectiion Vch Serial Number"
                  Top             =   285
                  Width           =   570
                  _ExtentX        =   1005
                  _ExtentY        =   503
                  Alignment       =   2
                  Mask            =   "###0"
                  MaxLength       =   4
                  DataType        =   1
                  IdName          =   "WCBTDSRFR"
               End
               Begin MwfCtl.MWCTL_MED LblSrTo 
                  Height          =   285
                  Index           =   164
                  Left            =   2340
                  TabIndex        =   384
                  ToolTipText     =   "Enter From Rejection Vch Serial Number"
                  Top             =   285
                  Width           =   570
                  _ExtentX        =   1005
                  _ExtentY        =   503
                  Alignment       =   2
                  Mask            =   "###0"
                  MaxLength       =   4
                  DataType        =   1
                  IdName          =   "WCBTDSRTO"
               End
               Begin VB.Label LblToSr 
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
                  Left            =   1920
                  TabIndex        =   390
                  Top             =   285
                  Width           =   615
               End
               Begin VB.Label LblSrFr 
                  BackStyle       =   0  'Transparent
                  Caption         =   "Sr Fr"
                  BeginProperty Font 
                     Name            =   "Arial"
                     Size            =   9.75
                     Charset         =   0
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   375
                  Left            =   60
                  TabIndex        =   389
                  Top             =   240
                  Width           =   1095
               End
               Begin VB.Label LblSlash3 
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
                  Left            =   2955
                  TabIndex        =   388
                  Top             =   0
                  Width           =   105
               End
               Begin VB.Label LblSlash2 
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
                  Left            =   2265
                  TabIndex        =   387
                  Top             =   0
                  Width           =   105
               End
               Begin VB.Label LblSlash1 
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
                  Left            =   1800
                  TabIndex        =   386
                  Top             =   0
                  Width           =   105
               End
               Begin VB.Label LblRejVchNo 
                  BackStyle       =   0  'Transparent
                  Caption         =   "Rej Vch No"
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
                  TabIndex        =   385
                  Top             =   15
                  Width           =   1065
               End
            End
            Begin MwfCtl.MWCTL_MED ATXT 
               Height          =   285
               Index           =   60
               Left            =   1155
               TabIndex        =   391
               ToolTipText     =   "Enter The Location To Which The Bags Have To Be Copied"
               Top             =   240
               Width           =   1245
               _ExtentX        =   2196
               _ExtentY        =   503
               MaxLength       =   8
               DataType        =   4
               IdName          =   "WCBTOBLOC"
            End
            Begin MwfCtl.MWCTL_MED ATXT 
               Height          =   285
               Index           =   61
               Left            =   1155
               TabIndex        =   393
               ToolTipText     =   "Enter From Bag Year"
               Top             =   525
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
               Left            =   2310
               TabIndex        =   395
               ToolTipText     =   "Enter From Bag Number"
               Top             =   525
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
               Left            =   1605
               TabIndex        =   394
               ToolTipText     =   "Enter From Bag Character"
               Top             =   525
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
               Left            =   1155
               TabIndex        =   396
               ToolTipText     =   "Enter To Bag Year"
               Top             =   810
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
               Left            =   2310
               TabIndex        =   398
               ToolTipText     =   "Enter To Bag Number"
               Top             =   810
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
               Left            =   1605
               TabIndex        =   397
               ToolTipText     =   "Enter To Bag Character"
               Top             =   810
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
               Left            =   1860
               TabIndex        =   400
               ToolTipText     =   "Enter From Bag Order Year "
               Top             =   1095
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
               Left            =   3000
               TabIndex        =   402
               ToolTipText     =   "Enter From Bag Order Number"
               Top             =   1095
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
               Left            =   1155
               TabIndex        =   399
               ToolTipText     =   "Enter From Bag Order Tc "
               Top             =   1095
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
               Left            =   2310
               TabIndex        =   401
               ToolTipText     =   "Enter From Bag Order Character"
               Top             =   1095
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
               Left            =   3855
               TabIndex        =   403
               ToolTipText     =   "Enter From Bag Order Serial Number"
               Top             =   1095
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
               Left            =   1860
               TabIndex        =   405
               ToolTipText     =   "Enter To Bag Order Year"
               Top             =   1380
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
               Left            =   3000
               TabIndex        =   407
               ToolTipText     =   "Enter To Bag Order Number"
               Top             =   1380
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
               Left            =   1155
               TabIndex        =   404
               ToolTipText     =   "Enter To Bag Order Tc"
               Top             =   1380
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
               Left            =   2310
               TabIndex        =   406
               ToolTipText     =   "Enter To Bag Order Character"
               Top             =   1380
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
               Left            =   3855
               TabIndex        =   408
               ToolTipText     =   "Enter To Bag Order Serial Number"
               Top             =   1380
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
               Left            =   1155
               TabIndex        =   409
               ToolTipText     =   "Enter From Design Code"
               Top             =   1665
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
               Left            =   1155
               TabIndex        =   410
               ToolTipText     =   "Enter To Design Code"
               Top             =   1950
               Width           =   2415
               _ExtentX        =   4260
               _ExtentY        =   503
               MaxLength       =   15
               DataType        =   4
               IdName          =   "WCBDMCDTO"
            End
            Begin MwfCtl.MWCTL_MED ATXT 
               Height          =   285
               Index           =   151
               Left            =   2490
               TabIndex        =   392
               ToolTipText     =   "Enter The Location From Which The Bags Have To Be Copied "
               Top             =   240
               Width           =   1245
               _ExtentX        =   2196
               _ExtentY        =   503
               MaxLength       =   8
               DataType        =   4
               IdName          =   "WCBFRBLOC"
            End
            Begin MwfCtl.MWCTL_MED ATXT 
               Height          =   285
               Index           =   154
               Left            =   1155
               TabIndex        =   411
               ToolTipText     =   "Enter Worker Code"
               Top             =   2240
               Width           =   1290
               _ExtentX        =   2275
               _ExtentY        =   503
               MaxLength       =   8
               DataType        =   4
               IdName          =   "WCBWRK"
            End
            Begin MwfCtl.MWCTL_MED ATXT 
               Height          =   285
               Index           =   155
               Left            =   3315
               TabIndex        =   412
               ToolTipText     =   "Enter Y for Bag With Production Yes/ N for All Bags"
               Top             =   2240
               Width           =   315
               _ExtentX        =   556
               _ExtentY        =   503
               DataType        =   4
               IdName          =   "WCBPRDYN"
            End
            Begin VB.Label LblCbWrk 
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
               Height          =   285
               Left            =   60
               TabIndex        =   433
               Top             =   2240
               Width           =   885
            End
            Begin VB.Label LblCbPrdYn 
               BackStyle       =   0  'Transparent
               Caption         =   "Prd Y/N"
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
               Left            =   2550
               TabIndex        =   432
               Top             =   2240
               Width           =   885
            End
            Begin VB.Label LblCbLoc 
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
               Left            =   60
               TabIndex        =   431
               Top             =   240
               Width           =   885
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
               Left            =   60
               TabIndex        =   430
               Top             =   525
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
               Left            =   1530
               TabIndex        =   429
               Top             =   525
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
               Left            =   2235
               TabIndex        =   428
               Top             =   525
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
               Left            =   60
               TabIndex        =   427
               Top             =   810
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
               Left            =   1530
               TabIndex        =   426
               Top             =   810
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
               Left            =   2235
               TabIndex        =   425
               Top             =   810
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
               Left            =   3780
               TabIndex        =   424
               Top             =   1095
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
               Left            =   60
               TabIndex        =   423
               Top             =   1095
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
               Left            =   1770
               TabIndex        =   422
               Top             =   1095
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
               Left            =   2235
               TabIndex        =   421
               Top             =   1095
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
               Left            =   2925
               TabIndex        =   420
               Top             =   1095
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
               Left            =   3780
               TabIndex        =   419
               Top             =   1380
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
               Left            =   60
               TabIndex        =   418
               Top             =   1380
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
               Left            =   1770
               TabIndex        =   417
               Top             =   1380
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
               Left            =   2235
               TabIndex        =   416
               Top             =   1380
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
               Left            =   2925
               TabIndex        =   415
               Top             =   1380
               Width           =   105
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
               Left            =   60
               TabIndex        =   414
               Top             =   1665
               Width           =   1125
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
               Left            =   60
               TabIndex        =   413
               Top             =   1950
               Width           =   1125
            End
         End
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
            Left            =   6240
            TabIndex        =   376
            ToolTipText     =   "Specify Copy From Scope / Bag Nos / Bag Alias No."
            Top             =   160
            Width           =   1215
         End
         Begin VB.OptionButton OptCpyBagFr 
            Caption         =   "Bag Nos"
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
            Left            =   7650
            TabIndex        =   375
            ToolTipText     =   "Specify Copy From Scope / Bag Nos / Bag Alias No."
            Top             =   160
            Width           =   1335
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
            Index           =   2
            Left            =   9180
            TabIndex        =   374
            ToolTipText     =   "Specify Copy From Scope / Bag Nos / Bag Alias No."
            Top             =   160
            Width           =   2055
         End
         Begin MwfCtl.MWCTL_BTN1 CmdCBGo 
            Height          =   375
            Left            =   4200
            TabIndex        =   37
            Top             =   3280
            Width           =   555
            _ExtentX        =   979
            _ExtentY        =   661
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
            Index           =   297
            Left            =   13530
            TabIndex        =   331
            ToolTipText     =   "Enter 'Y' to Copy Thru Bag Nos and  'N' to Copy thru RfId"
            Top             =   120
            Visible         =   0   'False
            Width           =   315
            _ExtentX        =   556
            _ExtentY        =   503
            DataType        =   4
            IdName          =   "WCPYTHRUBAGNO"
         End
         Begin VB.Label lblCbRfIdCpy 
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
            Left            =   4920
            TabIndex        =   436
            Top             =   140
            Width           =   1365
         End
         Begin VB.Label LblCpyThruBagNo 
            BackStyle       =   0  'Transparent
            Caption         =   "Copy thru Bag No Y/N"
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
            Left            =   11040
            TabIndex        =   332
            Top             =   600
            Width           =   2325
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
            Index           =   10
            Left            =   -30
            TabIndex        =   190
            Top             =   90
            Width           =   4740
         End
      End
      Begin VB.Frame FraExImp 
         Height          =   1725
         Left            =   720
         TabIndex        =   338
         Top             =   7440
         Visible         =   0   'False
         Width           =   8775
         Begin MSComDlg.CommonDialog CdSelPic 
            Left            =   8160
            Top             =   120
            _ExtentX        =   847
            _ExtentY        =   847
            _Version        =   393216
            DefaultExt      =   "*.jpg"
            DialogTitle     =   "Select Picture Name"
            PrinterDefault  =   0   'False
         End
         Begin MwfCtl.MWCTL_MED ATXT 
            Height          =   285
            Index           =   248
            Left            =   600
            TabIndex        =   339
            ToolTipText     =   "Excel File Name and Path"
            Top             =   600
            Width           =   7455
            _ExtentX        =   13150
            _ExtentY        =   503
            Enabled         =   0   'False
            MaxLength       =   250
            IdName          =   "WXLFILE"
         End
         Begin MwfCtl.MWCTL_BTN1 CmdXlImport 
            Height          =   375
            Left            =   8040
            TabIndex        =   340
            ToolTipText     =   "Click Here To Import Excel File"
            Top             =   600
            Width           =   555
            _ExtentX        =   979
            _ExtentY        =   661
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
         Begin MwfCtl.MWCTL_BTN1 CmdXlImpGo 
            Height          =   375
            Left            =   4680
            TabIndex        =   344
            ToolTipText     =   "Click Here To Import Excel File"
            Top             =   1080
            Width           =   555
            _ExtentX        =   979
            _ExtentY        =   661
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
            Index           =   250
            Left            =   2040
            TabIndex        =   342
            ToolTipText     =   "Specify Excel From Row Number"
            Top             =   1080
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
            Left            =   3600
            TabIndex        =   343
            ToolTipText     =   "Specify Excel To Row Number"
            Top             =   1080
            Width           =   705
            _ExtentX        =   1244
            _ExtentY        =   503
            Alignment       =   2
            Mask            =   "###0"
            MaxLength       =   4
            DataType        =   1
            IdName          =   "WXLTOROW"
         End
         Begin MwfCtl.MWCTL_MED ATXT 
            Height          =   285
            Index           =   249
            Left            =   795
            TabIndex        =   341
            ToolTipText     =   "Specify the Excel Sheet Number"
            Top             =   1080
            Width           =   465
            _ExtentX        =   820
            _ExtentY        =   503
            Alignment       =   2
            Mask            =   "#0"
            MaxLength       =   2
            DataType        =   1
            IdName          =   "WXLSHTNO"
         End
         Begin VB.Label LblMain 
            BackStyle       =   0  'Transparent
            BorderStyle     =   1  'Fixed Single
            Caption         =   "  Raw Material Issue from Excel"
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
            Index           =   19
            Left            =   0
            TabIndex        =   349
            Top             =   120
            Width           =   2940
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
            Left            =   120
            TabIndex        =   348
            Top             =   600
            Width           =   525
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
            Left            =   120
            TabIndex        =   347
            Top             =   1080
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
            Left            =   1410
            TabIndex        =   346
            Top             =   1080
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
            Left            =   2910
            TabIndex        =   345
            Top             =   1080
            Width           =   765
         End
      End
      Begin VB.Frame FraNKeyAll 
         BorderStyle     =   0  'None
         Height          =   9285
         Left            =   0
         TabIndex        =   194
         Top             =   0
         Width           =   15375
         Begin MwfCtl.MWCTL_FLX GrdTxndB 
            Height          =   2625
            Left            =   0
            TabIndex        =   195
            Top             =   6150
            Width           =   15075
            _ExtentX        =   26591
            _ExtentY        =   4630
            Cols            =   24
            colname1        =   "TDSRNO"
            heading1        =   "Sr.  "
            datafld1        =   "TdSrNo"
            datatype1       =   1
            mask1           =   "##0"
            maxlength1      =   3
            style1          =   2
            tooltiptext1    =   "Enter Sr No."
            colname2        =   "TDFRRMLOC"
            heading2        =   "Fr Loc               "
            datafld2        =   "TdFrRmLoc"
            datatype2       =   4
            recalcparent2   =   "GrdTxndB(TDBMRSNCD)"
            maxlength2      =   8
            tooltiptext2    =   "Enter Raw Material From Location"
            colname3        =   "WTDFRRMIR"
            heading3        =   "I/R "
            datatype3       =   4
            recalcon3       =   "GrdTxndB(TDFRRMDC)"
            recalcparent3   =   "GrdTxndB(TDFRRMDC)"
            maxlength3      =   1
            tooltiptext3    =   "Enter Issue Or Receipt"
            colname4        =   "TDFRRMDC"
            heading4        =   "D/C"
            datafld4        =   "TdFrRmDc"
            datatype4       =   4
            recalcon4       =   "GrdTxndB(WTDFRRMIR)"
            recalcparent4   =   "GrdTxndB(WTDFRRMIR)"
            maxlength4      =   1
            style4          =   0
            tooltiptext4    =   "Enter Debit Or Credit (D/C)"
            colname5        =   "TDRMCD"
            heading5        =   "Rm Code                                  "
            datafld5        =   "TdRmCd"
            datatype5       =   4
            recalcparent5   =   "GrdTxndB(TDRMSZ),GrdTxndB(TDRMSTKRT),GrdTxndB(TDRMQTY),GrdTxndB(TDRMWT),GrdTxndB(TDLOTNO)"
            maxlength5      =   16
            tooltiptext5    =   "Enter Raw Material Code"
            colname6        =   "TDLOTNO"
            heading6        =   "Lot No.                            "
            datafld6        =   "TdLotNo"
            datatype6       =   4
            recalcon6       =   "GrdTxndB(TDRMCD)"
            maxlength6      =   16
            tooltiptext6    =   "Enter Lot No."
            colname7        =   "TDRMSZ"
            heading7        =   "Ln/Sv/Ptr   "
            datafld7        =   "TdRmSz"
            datatype7       =   2
            mask7           =   "##0.0000"
            recalcon7       =   "GrdTxndB(TDRMCD)"
            maxlength7      =   8
            tooltiptext7    =   "Enter Raw Material Size"
            colname8        =   "TDRMSZ2"
            heading8        =   "Brdth  "
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
            heading10       =   "Stock Rate         "
            datafld10       =   "TdRmStkRt"
            datatype10      =   2
            mask10          =   "-######0.000"
            recalcon10      =   "GrdTxndB(TDRMCD)"
            maxlength10     =   12
            tooltiptext10   =   "Enter Raw Material Stock Rate"
            colname11       =   "TDRMQTY"
            heading11       =   "Qty      "
            datafld11       =   "TdRmQty"
            datatype11      =   1
            mask11          =   "####0"
            recalcon11      =   "GrdTxndB(TDRMCD)"
            maxlength11     =   5
            tooltiptext11   =   "Enter Raw Material Quantity"
            colname12       =   "TDRMWT"
            heading12       =   "Rm Wt           "
            datafld12       =   "TdRmWt"
            datatype12      =   2
            mask12          =   "-####0.000"
            recalcon12      =   "GrdTxndB(TDRMCD)"
            maxlength12     =   10
            tooltiptext12   =   "Enter Raw Material Weight"
            colname13       =   "WTDBRKMSG"
            heading13       =   "B/M"
            datatype13      =   4
            recalcon13      =   "GrdTxndB(TDTORMLOC)"
            recalcparent13  =   "GrdTxndB(TDTORMLOC)"
            maxlength13     =   1
            tooltiptext13   =   "Enter Broken Or Missing (B/M)"
            colname14       =   "TDTORMLOC"
            heading14       =   "To Loc               "
            datafld14       =   "TdToRmLoc"
            datatype14      =   4
            recalcon14      =   "GrdTxndB(WTDBRKMSG)"
            recalcparent14  =   "GrdTxndB(TDBYLOC),GrdTxndB(TDBYWRK),GrdTxndB(WTDBRKMSG),GrdTxndB(TDBMRSNCD)"
            maxlength14     =   8
            tooltiptext14   =   "Enter To Location"
            colname15       =   "TDBYLOC"
            heading15       =   "By Loc                "
            datafld15       =   "TdByLoc"
            datatype15      =   4
            recalcon15      =   "GrdTxndB(TDTORMLOC)"
            maxlength15     =   8
            tooltiptext15   =   "Enter Location"
            colname16       =   "TDBYWRK"
            heading16       =   "By Wrk              "
            datafld16       =   "TdByWrk"
            datatype16      =   4
            recalcon16      =   "GrdTxndB(TDTORMLOC)"
            maxlength16     =   8
            tooltiptext16   =   "Enter Worker"
            colname17       =   "TDBMRSNCD"
            heading17       =   "Reason Code"
            datafld17       =   "TdBmRsnCd"
            datatype17      =   4
            recalcon17      =   "GrdTxndB(TDTORMLOC),GrdTxndB(TDFRRMLOC)"
            maxlength17     =   5
            style17         =   0
            tooltiptext17   =   "Enter B/M Reason Code"
            colname18       =   "TDDUSTWT"
            heading18       =   "Dust Wt"
            datafld18       =   "TdDustWt"
            datatype18      =   2
            mask18          =   "##0.000"
            maxlength18     =   7
            tooltiptext18   =   "Enter Dust Weight"
            colname19       =   "TDDESC"
            datafld19       =   "TdDesc"
            recalcon19      =   "WTDDESC"
            maxlength19     =   60
            style19         =   0
            tooltiptext19   =   "Enter Description for the Raw Material"
            colname20       =   "TDFBRMQTY"
            heading20       =   "TdFBRmQty"
            datafld20       =   "TdFBRmQty"
            datatype20      =   4
            maxlength20     =   255
            style20         =   0
            tooltiptext20   =   "Enter Rm Quantity For Flute Bags"
            colname21       =   "TDFBRMWT"
            heading21       =   "TdFBRmWt"
            datafld21       =   "TdFBRmWt"
            datatype21      =   4
            maxlength21     =   255
            style21         =   0
            tooltiptext21   =   "Enter Rm Weight For Flute Bags"
            colname22       =   "TDPRTKEY"
            datafld22       =   "TdPrtKey"
            maxlength22     =   1
            style22         =   0
            colname23       =   "TDACK"
            datafld23       =   "TdAck"
            maxlength23     =   1
            style23         =   0
         End
         Begin MwfCtl.MWCTL_MED ATXT 
            Height          =   285
            Index           =   117
            Left            =   12570
            TabIndex        =   229
            TabStop         =   0   'False
            Top             =   270
            Visible         =   0   'False
            Width           =   195
            _ExtentX        =   344
            _ExtentY        =   503
            Enabled         =   0   'False
            MaxLength       =   12
            DataType        =   4
            DataField       =   "TLotNo"
            IdName          =   "TLOTNO"
         End
         Begin MwfCtl.MWCTL_MED ATXT 
            Height          =   285
            Index           =   96
            Left            =   12345
            TabIndex        =   196
            ToolTipText     =   "Enter Import Type"
            Top             =   270
            Visible         =   0   'False
            Width           =   225
            _ExtentX        =   397
            _ExtentY        =   503
            Enabled         =   0   'False
            MaxLength       =   5
            DataType        =   4
            DataField       =   "TImpTyp"
            IdName          =   "TIMPTYP"
         End
         Begin MwfCtl.MWCTL_MED ATXT 
            Height          =   285
            Index           =   59
            Left            =   12150
            TabIndex        =   197
            Top             =   270
            Visible         =   0   'False
            Width           =   195
            _ExtentX        =   344
            _ExtentY        =   503
            Enabled         =   0   'False
            MaxLength       =   8
            DataType        =   3
            DataField       =   "TUseBefDt"
            IdName          =   "TUSEBEFDT"
         End
         Begin MwfCtl.MWCTL_MED ATXT 
            Height          =   285
            Index           =   3
            Left            =   11865
            TabIndex        =   198
            Top             =   0
            Visible         =   0   'False
            Width           =   195
            _ExtentX        =   344
            _ExtentY        =   503
            Alignment       =   2
            Enabled         =   0   'False
            Mask            =   "####0.00"
            MaxLength       =   8
            DataType        =   2
            DataField       =   "TDustWt"
            IdName          =   "TDUSTWT"
         End
         Begin MwfCtl.MWCTL_MED ATXT 
            Height          =   285
            Index           =   7
            Left            =   11670
            TabIndex        =   199
            Top             =   0
            Visible         =   0   'False
            Width           =   195
            _ExtentX        =   344
            _ExtentY        =   503
            Enabled         =   0   'False
            MaxLength       =   5
            DataType        =   4
            DataField       =   "TLsLoc"
            IdName          =   "TLSLOC"
         End
         Begin MwfCtl.MWCTL_MED ATXT 
            Height          =   285
            Index           =   8
            Left            =   11475
            TabIndex        =   200
            Top             =   0
            Visible         =   0   'False
            Width           =   195
            _ExtentX        =   344
            _ExtentY        =   503
            Alignment       =   2
            Enabled         =   0   'False
            Mask            =   "####0.00"
            MaxLength       =   8
            DataType        =   2
            DataField       =   "TGldRec"
            IdName          =   "TGLDREC"
         End
         Begin MwfCtl.MWCTL_MED ATXT 
            Height          =   285
            Index           =   9
            Left            =   11280
            TabIndex        =   201
            Top             =   0
            Visible         =   0   'False
            Width           =   195
            _ExtentX        =   344
            _ExtentY        =   503
            Enabled         =   0   'False
            MaxLength       =   8
            DataType        =   3
            DataField       =   "TLsFrDt"
            IdName          =   "TLSFRDT"
         End
         Begin MwfCtl.MWCTL_MED ATXT 
            Height          =   285
            Index           =   10
            Left            =   11085
            TabIndex        =   202
            Top             =   0
            Visible         =   0   'False
            Width           =   195
            _ExtentX        =   344
            _ExtentY        =   503
            Enabled         =   0   'False
            MaxLength       =   8
            DataType        =   3
            DataField       =   "TLsToDt"
            IdName          =   "TLSTODT"
         End
         Begin MwfCtl.MWCTL_MED ATXT 
            Height          =   285
            Index           =   11
            Left            =   10890
            TabIndex        =   203
            Top             =   300
            Visible         =   0   'False
            Width           =   195
            _ExtentX        =   344
            _ExtentY        =   503
            Enabled         =   0   'False
            MaxLength       =   8
            DataType        =   4
            DataField       =   "TSuppCd"
            IdName          =   "TSUPPCD"
         End
         Begin MwfCtl.MWCTL_MED ATXT 
            Height          =   285
            Index           =   12
            Left            =   10890
            TabIndex        =   204
            Top             =   0
            Visible         =   0   'False
            Width           =   195
            _ExtentX        =   344
            _ExtentY        =   503
            Alignment       =   2
            Enabled         =   0   'False
            Mask            =   "####0.00"
            MaxLength       =   8
            DataType        =   2
            DataField       =   "TGldInLs"
            IdName          =   "TGLDINLS"
         End
         Begin MwfCtl.MWCTL_MED ATXT 
            Height          =   285
            Index           =   13
            Left            =   11085
            TabIndex        =   205
            Top             =   300
            Visible         =   0   'False
            Width           =   195
            _ExtentX        =   344
            _ExtentY        =   503
            Enabled         =   0   'False
            MaxLength       =   20
            DataType        =   4
            DataField       =   "TBillNo"
            IdName          =   "TBILLNO"
         End
         Begin MwfCtl.MWCTL_MED ATXT 
            Height          =   285
            Index           =   14
            Left            =   11475
            TabIndex        =   206
            Top             =   300
            Visible         =   0   'False
            Width           =   195
            _ExtentX        =   344
            _ExtentY        =   503
            Enabled         =   0   'False
            MaxLength       =   20
            DataType        =   4
            DataField       =   "TBOENo"
            IdName          =   "TBOENO"
         End
         Begin MwfCtl.MWCTL_MED ATXT 
            Height          =   285
            Index           =   15
            Left            =   11280
            TabIndex        =   207
            Top             =   300
            Visible         =   0   'False
            Width           =   195
            _ExtentX        =   344
            _ExtentY        =   503
            Enabled         =   0   'False
            MaxLength       =   8
            DataType        =   3
            DataField       =   "TBillDt"
            IdName          =   "TBILLDT"
         End
         Begin MwfCtl.MWCTL_MED ATXT 
            Height          =   285
            Index           =   16
            Left            =   11670
            TabIndex        =   208
            Top             =   300
            Visible         =   0   'False
            Width           =   195
            _ExtentX        =   344
            _ExtentY        =   503
            Enabled         =   0   'False
            MaxLength       =   8
            DataType        =   3
            DataField       =   "TBOEDt"
            IdName          =   "TBOEDT"
         End
         Begin MwfCtl.MWCTL_MED ATXT 
            Height          =   285
            Index           =   17
            Left            =   11865
            TabIndex        =   209
            Top             =   300
            Visible         =   0   'False
            Width           =   195
            _ExtentX        =   344
            _ExtentY        =   503
            Alignment       =   2
            Enabled         =   0   'False
            Mask            =   "##0.000"
            MaxLength       =   7
            DataType        =   2
            DataField       =   "TCnvRt"
            IdName          =   "TCNVRT"
         End
         Begin MwfCtl.MWCTL_MED ATXT 
            Height          =   285
            Index           =   18
            Left            =   13125
            TabIndex        =   210
            Top             =   0
            Visible         =   0   'False
            Width           =   195
            _ExtentX        =   344
            _ExtentY        =   503
            Alignment       =   2
            Enabled         =   0   'False
            Mask            =   "#######0.00"
            MaxLength       =   11
            DataType        =   2
            DataField       =   "TTotAmt"
            IdName          =   "TTOTAMT"
         End
         Begin MwfCtl.MWCTL_MED ATXT 
            Height          =   285
            Index           =   19
            Left            =   12930
            TabIndex        =   211
            Top             =   0
            Visible         =   0   'False
            Width           =   195
            _ExtentX        =   344
            _ExtentY        =   503
            Alignment       =   2
            Enabled         =   0   'False
            Mask            =   "#####0.00"
            MaxLength       =   9
            DataType        =   2
            DataField       =   "TAddAmt4"
            IdName          =   "TADDAMT4"
         End
         Begin MwfCtl.MWCTL_MED ATXT 
            Height          =   285
            Index           =   20
            Left            =   12735
            TabIndex        =   212
            Top             =   0
            Visible         =   0   'False
            Width           =   195
            _ExtentX        =   344
            _ExtentY        =   503
            Alignment       =   2
            Enabled         =   0   'False
            Mask            =   "#####0.00"
            MaxLength       =   9
            DataType        =   2
            DataField       =   "TAddAmt3"
            IdName          =   "TADDAMT3"
         End
         Begin MwfCtl.MWCTL_MED ATXT 
            Height          =   285
            Index           =   21
            Left            =   12540
            TabIndex        =   213
            Top             =   0
            Visible         =   0   'False
            Width           =   195
            _ExtentX        =   344
            _ExtentY        =   503
            Alignment       =   2
            Enabled         =   0   'False
            Mask            =   "#####0.00"
            MaxLength       =   9
            DataType        =   2
            DataField       =   "TAddAmt2"
            IdName          =   "TADDAMT2"
         End
         Begin MwfCtl.MWCTL_MED ATXT 
            Height          =   285
            Index           =   23
            Left            =   12345
            TabIndex        =   214
            Top             =   0
            Visible         =   0   'False
            Width           =   195
            _ExtentX        =   344
            _ExtentY        =   503
            Alignment       =   2
            Enabled         =   0   'False
            Mask            =   "#####0.00"
            MaxLength       =   9
            DataType        =   2
            DataField       =   "TAddAmt1"
            IdName          =   "TADDAMT1"
         End
         Begin MwfCtl.MWCTL_MED ATXT 
            Height          =   285
            Index           =   24
            Left            =   12150
            TabIndex        =   215
            Top             =   0
            Visible         =   0   'False
            Width           =   195
            _ExtentX        =   344
            _ExtentY        =   503
            Alignment       =   2
            Enabled         =   0   'False
            Mask            =   "#######0.00"
            MaxLength       =   11
            DataType        =   2
            DataField       =   "TNetAmt"
            IdName          =   "TNETAMT"
         End
         Begin MwfCtl.MWCTL_MED ATXT 
            Height          =   285
            Index           =   100
            Left            =   10860
            TabIndex        =   217
            Top             =   570
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
            Index           =   135
            Left            =   13320
            TabIndex        =   262
            ToolTipText     =   "Enter Transaction in US$ (Y/N)"
            Top             =   0
            Visible         =   0   'False
            Width           =   225
            _ExtentX        =   397
            _ExtentY        =   503
            Enabled         =   0   'False
            DataType        =   4
            DataField       =   "TMULDIV"
            IdName          =   "TMULDIV"
         End
         Begin MwfCtl.MWCTL_MED ATXT 
            Height          =   285
            Index           =   136
            Left            =   13530
            TabIndex        =   263
            Top             =   0
            Visible         =   0   'False
            Width           =   225
            _ExtentX        =   397
            _ExtentY        =   503
            Enabled         =   0   'False
            DataType        =   4
            DataField       =   "TCURRCD"
            IdName          =   "TCURRCD"
         End
         Begin MwfCtl.MWCTL_MED ATXT 
            Height          =   285
            Index           =   137
            Left            =   13800
            TabIndex        =   266
            Top             =   0
            Visible         =   0   'False
            Width           =   225
            _ExtentX        =   397
            _ExtentY        =   503
            Enabled         =   0   'False
            DataType        =   4
            DataField       =   "TMULDIVRS"
            IdName          =   "TMULDIVRS"
         End
         Begin MwfCtl.MWCTL_MED ATXT 
            Height          =   285
            Index           =   138
            Left            =   14010
            TabIndex        =   267
            Top             =   0
            Visible         =   0   'False
            Width           =   225
            _ExtentX        =   397
            _ExtentY        =   503
            Enabled         =   0   'False
            Mask            =   "###0.0000"
            MaxLength       =   9
            DataType        =   2
            DataField       =   "TCnvRtRs"
            IdName          =   "TCNVRTRS"
         End
         Begin MwfCtl.MWCTL_MED ATXT 
            Height          =   285
            Index           =   140
            Left            =   14220
            TabIndex        =   271
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
            Index           =   153
            Left            =   14520
            TabIndex        =   299
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
            Index           =   158
            Left            =   13380
            TabIndex        =   302
            TabStop         =   0   'False
            Top             =   300
            Visible         =   0   'False
            Width           =   195
            _ExtentX        =   344
            _ExtentY        =   503
            Enabled         =   0   'False
            MaxLength       =   45
            DataType        =   4
            DataField       =   "TSuppPayTerms"
            IdName          =   "TSUPPPAYTERMS"
         End
         Begin MwfCtl.MWCTL_MED ATXT 
            Height          =   285
            Index           =   102
            Left            =   11490
            TabIndex        =   264
            ToolTipText     =   "Set Number"
            Top             =   8715
            Visible         =   0   'False
            Width           =   885
            _ExtentX        =   1561
            _ExtentY        =   503
            Enabled         =   0   'False
            Mask            =   "######0"
            MaxLength       =   7
            DataType        =   1
            IdName          =   "WTDSETNO"
         End
         Begin MwfCtl.MWCTL_MED ATXT 
            Height          =   285
            Index           =   30
            Left            =   960
            TabIndex        =   7
            Top             =   285
            Width           =   855
            _ExtentX        =   1508
            _ExtentY        =   503
            MaxLength       =   8
            DataType        =   3
            DataField       =   "TDt"
            IdName          =   "TDT"
         End
         Begin MwfCtl.MWCTL_MED ATXT 
            Height          =   285
            Index           =   22
            Left            =   3195
            TabIndex        =   14
            ToolTipText     =   "Enter Raw Material Default To Location"
            Top             =   570
            Width           =   1365
            _ExtentX        =   2408
            _ExtentY        =   503
            MaxLength       =   8
            DataType        =   4
            DataField       =   "TToRmLoc"
            ReCalcOn        =   "TCHR"
            IdName          =   "TTORMLOC"
         End
         Begin MwfCtl.MWCTL_MED ATXT 
            Height          =   285
            Index           =   0
            Left            =   990
            TabIndex        =   13
            ToolTipText     =   "Enter Raw Material Default From Location"
            Top             =   570
            Width           =   1350
            _ExtentX        =   2381
            _ExtentY        =   503
            MaxLength       =   8
            DataType        =   4
            DataField       =   "TFrRmLoc"
            ReCalcOn        =   "TCHR"
            IdName          =   "TFRRMLOC"
         End
         Begin MwfCtl.MWCTL_MED ATXT 
            Height          =   285
            Index           =   25
            Left            =   615
            TabIndex        =   216
            ToolTipText     =   "Enter Raw Material Description"
            Top             =   8715
            Width           =   10050
            _ExtentX        =   17727
            _ExtentY        =   503
            MaxLength       =   60
            ReCalcParent    =   "GrdTxndB(TDDESC)"
            IdName          =   "WTDDESC"
         End
         Begin MwfCtl.MWCTL_MED ATXT 
            Height          =   285
            Index           =   2
            Left            =   3195
            TabIndex        =   8
            ToolTipText     =   "Enter Bag Location"
            Top             =   285
            Width           =   1365
            _ExtentX        =   2408
            _ExtentY        =   503
            MaxLength       =   8
            DataType        =   4
            DataField       =   "TFrBLoc"
            ReCalcOn        =   "TCHR"
            IdName          =   "TFRBLOC"
         End
         Begin MwfCtl.MWCTL_MED ATXT 
            Height          =   285
            Index           =   99
            Left            =   8580
            TabIndex        =   373
            ToolTipText     =   "Enter Quantity To Be Cross Transferred"
            Top             =   540
            Visible         =   0   'False
            Width           =   435
            _ExtentX        =   767
            _ExtentY        =   503
            Alignment       =   2
            Mask            =   "##0.0"
            MaxLength       =   5
            DataType        =   2
            DataField       =   "TTRFQTY"
            ReCalcParent    =   "TMULTIIR"
            ReCalcOn        =   "WTRFYN"
            IdName          =   "TTRFQTY"
         End
         Begin MwfCtl.MWCTL_MED ATXT 
            Height          =   285
            Index           =   101
            Left            =   14205
            TabIndex        =   218
            Top             =   570
            Visible         =   0   'False
            Width           =   855
            _ExtentX        =   1508
            _ExtentY        =   503
            Enabled         =   0   'False
            MaxLength       =   8
            DataType        =   3
            IdName          =   "WMODDT"
         End
         Begin MwfCtl.MWCTL_MED ATXT 
            Height          =   285
            Index           =   109
            Left            =   7155
            TabIndex        =   6
            ToolTipText     =   "Enter Order Lock (Y/N)"
            Top             =   0
            Width           =   345
            _ExtentX        =   609
            _ExtentY        =   503
            DataType        =   4
            DataField       =   "TLockYN"
            IdName          =   "TLOCKYN"
         End
         Begin MwfCtl.MWCTL_MED ATXT 
            Height          =   570
            Index           =   48
            Left            =   10065
            TabIndex        =   12
            ToolTipText     =   "Enter Description"
            Top             =   0
            Width           =   4980
            _ExtentX        =   8784
            _ExtentY        =   1005
            Alignment       =   3
            MaxLength       =   60
            DataField       =   "TDesc"
            IdName          =   "TDESC"
         End
         Begin VB.TextBox txtPwd 
            Height          =   285
            IMEMode         =   3  'DISABLE
            Left            =   10065
            MaxLength       =   30
            PasswordChar    =   "*"
            TabIndex        =   17
            Top             =   570
            Width           =   2535
         End
         Begin MwfCtl.MWCTL_MED ATXT 
            Height          =   285
            Index           =   152
            Left            =   8625
            TabIndex        =   11
            ToolTipText     =   "Enter I for Multi-Issue Voucher/ R for Multi-Receive Voucher"
            Top             =   270
            Visible         =   0   'False
            Width           =   345
            _ExtentX        =   609
            _ExtentY        =   503
            DataType        =   4
            DataField       =   "TMultiIr"
            ReCalcOn        =   "TTRFQTY"
            IdName          =   "TMULTIIR"
         End
         Begin MwfCtl.MWCTL_MED ATXT 
            Height          =   285
            Index           =   165
            Left            =   5715
            TabIndex        =   9
            ToolTipText     =   "Enter To Bag Location"
            Top             =   285
            Visible         =   0   'False
            Width           =   1485
            _ExtentX        =   2619
            _ExtentY        =   503
            MaxLength       =   8
            DataType        =   4
            IdName          =   "WTTOBLOC"
         End
         Begin MwfCtl.MWCTL_MED ATXT 
            Height          =   285
            Index           =   166
            Left            =   5715
            TabIndex        =   15
            ToolTipText     =   "Enter To Default Rejection Code"
            Top             =   570
            Visible         =   0   'False
            Width           =   960
            _ExtentX        =   1693
            _ExtentY        =   503
            MaxLength       =   8
            DataType        =   4
            IdName          =   "WTRJCD"
         End
         Begin MwfCtl.MWCTL_MED ATXT 
            Height          =   285
            Index           =   167
            Left            =   7995
            TabIndex        =   16
            ToolTipText     =   "Enter Order Lock (Y/N)"
            Top             =   570
            Visible         =   0   'False
            Width           =   345
            _ExtentX        =   609
            _ExtentY        =   503
            DataType        =   4
            IdName          =   "WTBIR"
         End
         Begin MwfCtl.MWCTL_FLX GrdTxndA 
            Height          =   5220
            Left            =   0
            TabIndex        =   18
            Top             =   840
            Width           =   15135
            _ExtentX        =   26696
            _ExtentY        =   9208
            Cols            =   36
            colname1        =   "TDSR"
            heading1        =   "Sr.  "
            datafld1        =   "TdSr"
            datatype1       =   1
            mask1           =   "##0"
            maxlength1      =   3
            style1          =   2
            tooltiptext1    =   "Enter Sr. No"
            colname2        =   "TDSRNO"
            datafld2        =   "TdSrNo"
            datatype2       =   1
            mask2           =   "#0"
            maxlength2      =   2
            style2          =   0
            tooltiptext2    =   "Enter Sr No."
            colname3        =   "WBAGNO"
            datatype3       =   4
            maxlength3      =   18
            tooltiptext3    =   "Enter Bag Number Using Barcode Reader"
            colname4        =   "TDBYY"
            heading4        =   "Yy "
            datafld4        =   "TdBYy"
            datatype4       =   4
            recalcparent4   =   $"EmrFrmTxn.frx":0000
            maxlength4      =   2
            tooltiptext4    =   "Enter Bag Year"
            colname5        =   "TDBCHR"
            heading5        =   "Chr        "
            datafld5        =   "TdBChr"
            datatype5       =   4
            recalcparent5   =   $"EmrFrmTxn.frx":0156
            maxlength5      =   3
            tooltiptext5    =   "Enter Bag Character"
            colname6        =   "TDBNO"
            heading6        =   "No.            "
            datafld6        =   "TdBNo"
            datatype6       =   1
            mask6           =   "######0"
            recalcon6       =   "GrdTxndA(TDBYY),GrdTxndA(TDBCHR)"
            recalcparent6   =   $"EmrFrmTxn.frx":02AC
            maxlength6      =   7
            tooltiptext6    =   "Enter Bag No."
            colname7        =   "WTDBIR"
            heading7        =   "I/R "
            datatype7       =   4
            recalcon7       =   "GrdTxndA(TDBDC)"
            recalcparent7   =   "GrdTxndA(TDBDC),GrdTxndA(TDBQTY),GrdTxndA(TDPTQTY)"
            maxlength7      =   1
            style7          =   0
            tooltiptext7    =   "Enter Issue Or Receipt (I/R)"
            colname8        =   "TDBDC"
            heading8        =   "D/C"
            datafld8        =   "TdBDc"
            datatype8       =   4
            recalcon8       =   "GrdTxndA(WTDBIR)"
            recalcparent8   =   "GrdTxndA(WTDBIR),GrdTxndA(TDRJDC)"
            maxlength8      =   1
            style8          =   0
            tooltiptext8    =   "Enter Debit Or Credit (D/C)"
            colname9        =   "TDBQTY"
            heading9        =   "Qty    "
            datafld9        =   "TdBQty"
            datatype9       =   2
            mask9           =   "##0.0"
            recalcon9       =   "GrdTxndA(TDBYY),GrdTxndA(TDBCHR),GrdTxndA(TDBNO),GrdTxndA(WTDBIR)"
            recalcparent9   =   "GrdTxndA(TDRJQTY)"
            maxlength9      =   5
            style9          =   0
            tooltiptext9    =   "Enter Bag Quantity"
            colname10       =   "TDPTQTY"
            heading10       =   "PrtQty"
            datafld10       =   "TdPtQty"
            datatype10      =   1
            mask10          =   "-###0"
            recalcon10      =   "GrdTxndA(TDBYY),GrdTxndA(TDBCHR),GrdTxndA(TDBNO),GrdTxndA(WTDBIR)"
            recalcparent10  =   "GrdTxndA(TDRJPTQTY)"
            maxlength10     =   5
            style10         =   2
            tooltiptext10   =   "Enter Part Quantity"
            colname11       =   "TDBGRWT"
            heading11       =   "GrossWt     "
            datafld11       =   "TdBGrWt"
            datatype11      =   2
            mask11          =   "####0.000"
            recalcon11      =   "GrdTxndA(TDBYY),GrdTxndA(TDBCHR),GrdTxndA(TDBNO)"
            maxlength11     =   9
            tooltiptext11   =   "Enter Bag Gross Weight"
            colname12       =   "TDFRBLOC"
            heading12       =   "Fr Loc             "
            datafld12       =   "TdFrBLoc"
            datatype12      =   4
            recalcon12      =   "GrdTxndA(TDBYY),GrdTxndA(TDBCHR),GrdTxndA(TDBNO)"
            recalcparent12  =   "GrdTxndA(TDTOBLOC),GrdTxndA(TDACK)"
            maxlength12     =   8
            style12         =   0
            tooltiptext12   =   "Enter From Bag Location"
            colname13       =   "TDTOBLOC"
            heading13       =   "To Loc             "
            datafld13       =   "TdToBLoc"
            datatype13      =   4
            recalcon13      =   "GrdTxndA(TDBYY),GrdTxndA(TDBCHR),GrdTxndA(TDBNO),GrdTxndA(TDFRBLOC)"
            recalcparent13  =   "GrdTxndA(TDACK)"
            maxlength13     =   8
            style13         =   0
            tooltiptext13   =   "Enter To Bag Location"
            colname14       =   "TDRJDC"
            datafld14       =   "TdRjDc"
            datatype14      =   4
            recalcon14      =   "GrdTxndA(TDBYY),GrdTxndA(TDBCHR),GrdTxndA(TDBNO),GrdTxndA(TDBDC)"
            recalcparent14  =   "GrdTxndA(TDRJCD)"
            maxlength14     =   1
            style14         =   0
            tooltiptext14   =   "Enter Debit Or Credit (D/C) For Rejection"
            colname15       =   "TDRJQTY"
            heading15       =   "RjQty"
            datafld15       =   "TdRjQty"
            datatype15      =   2
            mask15          =   "##0.0"
            recalcon15      =   "GrdTxndA(TDBYY),GrdTxndA(TDBCHR),GrdTxndA(TDBNO),GrdTxndA(TDBQTY)"
            maxlength15     =   5
            style15         =   0
            tooltiptext15   =   "Enter Rejection Quantity"
            colname16       =   "TDRJPTQTY"
            heading16       =   "RjPrtQty"
            datafld16       =   "TdRjPtQty"
            datatype16      =   1
            mask16          =   "-#0"
            recalcon16      =   "GrdTxndA(TDBYY),GrdTxndA(TDBCHR),GrdTxndA(TDBNO),GrdTxndA(TDPTQTY)"
            maxlength16     =   3
            style16         =   0
            tooltiptext16   =   "Enter Rejection Part Quantity"
            colname17       =   "TDRJCD"
            heading17       =   "Rj Code       "
            datafld17       =   "TdRjCd"
            datatype17      =   4
            recalcon17      =   "GrdTxndA(TDRJDC)"
            maxlength17     =   8
            style17         =   0
            tooltiptext17   =   "Enter Rejection Code"
            colname18       =   "TDWRK"
            heading18       =   "Worker         "
            datafld18       =   "TdWrk"
            datatype18      =   4
            recalcon18      =   "GrdTxndA(TDBYY),GrdTxndA(TDBCHR),GrdTxndA(TDBNO)"
            recalcparent18  =   "GrdTxndA(TDPRDYN)"
            maxlength18     =   8
            tooltiptext18   =   "Enter Worker"
            colname19       =   "TDACK"
            heading19       =   "Ack"
            datafld19       =   "TdAck"
            datatype19      =   4
            recalcon19      =   "GrdTxndA(TDBYY),GrdTxndA(TDBCHR),GrdTxndA(TDBNO),GrdTxndA(TDFRBLOC),GrdTxndA(TDTOBLOC)"
            maxlength19     =   1
            style19         =   0
            tooltiptext19   =   "Enter Acknowledgement (Y/N)"
            colname20       =   "TDPRDYN"
            heading20       =   "Prd"
            datafld20       =   "TdPrdYN"
            datatype20      =   4
            recalcon20      =   "GrdTxndA(TDBYY),GrdTxndA(TDBCHR),GrdTxndA(TDBNO),GrdTxndA(TDWRK)"
            recalcparent20  =   "GrdTxndA(TDBYLOC),GrdTxndA(TDBYWRK)"
            maxlength20     =   1
            tooltiptext20   =   "Enter Production Completed (Y/N)"
            colname21       =   "TDPDQTY"
            heading21       =   "PDQty"
            datafld21       =   "TdPDQty"
            datatype21      =   2
            mask21          =   "##0.0"
            maxlength21     =   5
            style21         =   0
            tooltiptext21   =   "Enter PD Bag Qty"
            colname22       =   "WTDPRDPTS"
            tooltiptext22   =   "Column To Go To The PrdPts Button"
            colname23       =   "TDBYLOC"
            heading23       =   "By Loc              "
            datafld23       =   "TdByLoc"
            datatype23      =   4
            recalcon23      =   "GrdTxndA(TDBYY),GrdTxndA(TDBCHR),GrdTxndA(TDBNO),GrdTxndA(TDPRDYN)"
            maxlength23     =   8
            tooltiptext23   =   "Enter Location"
            colname24       =   "TDBYWRK"
            heading24       =   "By Wrk          "
            datafld24       =   "TdByWrk"
            datatype24      =   4
            recalcon24      =   "GrdTxndA(TDBYY),GrdTxndA(TDBCHR),GrdTxndA(TDBNO),GrdTxndA(TDPRDYN)"
            maxlength24     =   8
            tooltiptext24   =   "Enter Worker"
            colname25       =   "TDKEY"
            datafld25       =   "TdKey"
            datatype25      =   1
            mask25          =   "#######0"
            maxlength25     =   8
            style25         =   2
            tooltiptext25   =   "Enter Key"
            colname26       =   "TDWAXWT"
            heading26       =   "Wax Wt    "
            datafld26       =   "TdWaxWt"
            datatype26      =   2
            mask26          =   "###0.000"
            maxlength26     =   8
            tooltiptext26   =   "Enter Wax Pullout Weight."
            colname27       =   "TDDESC"
            heading27       =   "Desc                                                                                "
            datafld27       =   "TdDesc"
            maxlength27     =   60
            tooltiptext27   =   "Enter Description"
            colname28       =   "TDPURRT"
            heading28       =   "Pur Rt         "
            datafld28       =   "TdPurRt"
            datatype28      =   2
            mask28          =   "#####0.00"
            maxlength28     =   9
            style28         =   0
            tooltiptext28   =   "Enter Pure Rate"
            colname29       =   "TDPURAMT"
            heading29       =   "Pur Amt         "
            datafld29       =   "TdPurAmt"
            datatype29      =   2
            mask29          =   "######0.00"
            maxlength29     =   10
            style29         =   0
            tooltiptext29   =   "Enter Pure Amount"
            colname30       =   "TDPREVYY"
            datafld30       =   "TdPrevYy"
            datatype30      =   4
            recalcon30      =   "GrdTxndA(TDBYY),GrdTxndA(TDBCHR),GrdTxndA(TDBNO)"
            maxlength30     =   2
            style30         =   0
            tooltiptext30   =   "Enter Previous Year"
            colname31       =   "TDPREVKEY"
            datafld31       =   "TdPrevKey"
            datatype31      =   1
            mask31          =   "#######0"
            recalcon31      =   "GrdTxndA(TDBYY),GrdTxndA(TDBCHR),GrdTxndA(TDBNO)"
            maxlength31     =   8
            style31         =   0
            tooltiptext31   =   "Enter Previous Key"
            colname32       =   "TDPRTKEY"
            heading32       =   "Prtn "
            datafld32       =   "TdPrtKey"
            maxlength32     =   1
            style32         =   0
            colname33       =   "WTDCMCD"
            heading33       =   "CmCd"
            datatype33      =   4
            recalcon33      =   "GrdTxndA(TDBYY),GrdTxndA(TDBCHR),GrdTxndA(TDBNO)"
            maxlength33     =   8
            style33         =   2
            style34         =   0
            style35         =   0
         End
         Begin MwfCtl.MWCTL_MED ATXT 
            Height          =   285
            Index           =   184
            Left            =   8625
            TabIndex        =   10
            ToolTipText     =   "Enter Rm Transfer (Yes/No)"
            Top             =   -15
            Visible         =   0   'False
            Width           =   345
            _ExtentX        =   609
            _ExtentY        =   503
            DataType        =   4
            ReCalcParent    =   "TTRFQTY"
            IdName          =   "WTRFYN"
         End
         Begin VB.Image PicDp 
            Height          =   3285
            Left            =   -120
            Top             =   0
            Width           =   3105
         End
         Begin VB.Label LblTBIr 
            BackStyle       =   0  'Transparent
            Caption         =   "Bag I/R"
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
            Left            =   7215
            TabIndex        =   305
            Top             =   570
            Visible         =   0   'False
            Width           =   1035
         End
         Begin VB.Label LblTRjCd 
            BackStyle       =   0  'Transparent
            Caption         =   "Rej Code"
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
            Left            =   4650
            TabIndex        =   304
            Top             =   585
            Visible         =   0   'False
            Width           =   1020
         End
         Begin VB.Label LblMultiIr 
            BackStyle       =   0  'Transparent
            Caption         =   "Multi I/R"
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
            Left            =   7680
            TabIndex        =   298
            Top             =   270
            Visible         =   0   'False
            Width           =   870
         End
         Begin VB.Label LblSetNo 
            BackStyle       =   0  'Transparent
            Caption         =   "Set"
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
            Left            =   11070
            TabIndex        =   265
            Top             =   8715
            Visible         =   0   'False
            Width           =   495
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
            Left            =   30
            TabIndex        =   228
            Top             =   285
            Width           =   615
         End
         Begin VB.Label LblTFrRmLoc 
            BackStyle       =   0  'Transparent
            Caption         =   "RmLoc Fr"
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
            TabIndex        =   227
            Top             =   570
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
            Left            =   30
            TabIndex        =   226
            Top             =   8715
            Width           =   660
         End
         Begin VB.Label ALBL 
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
            Index           =   21
            Left            =   9480
            TabIndex        =   225
            Top             =   0
            Width           =   615
         End
         Begin VB.Label LblTToRmLoc 
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
            Left            =   2385
            TabIndex        =   224
            Top             =   570
            Width           =   270
         End
         Begin VB.Label LblTFrBLoc 
            BackStyle       =   0  'Transparent
            Caption         =   "BagLoc"
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
            Left            =   2370
            TabIndex        =   223
            Top             =   285
            Width           =   810
         End
         Begin VB.Label LblTrfQty 
            BackStyle       =   0  'Transparent
            Caption         =   "Trf YN"
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
            Left            =   7935
            TabIndex        =   222
            Top             =   15
            Visible         =   0   'False
            Width           =   735
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
            Index           =   55
            Left            =   9120
            TabIndex        =   221
            Top             =   570
            Width           =   975
         End
         Begin VB.Label LblModDt 
            BackStyle       =   0  'Transparent
            Caption         =   "Last Mod Dt"
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
            Left            =   12930
            TabIndex        =   220
            Top             =   570
            Visible         =   0   'False
            Width           =   1275
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
            Index           =   58
            Left            =   6135
            TabIndex        =   219
            Top             =   0
            Width           =   1035
         End
         Begin VB.Label LblTToBLoc 
            BackStyle       =   0  'Transparent
            Caption         =   "To BagLoc"
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
            Left            =   4620
            TabIndex        =   303
            Top             =   285
            Visible         =   0   'False
            Width           =   1050
         End
      End
      Begin VB.Frame FraBS 
         Height          =   1005
         Left            =   5085
         TabIndex        =   41
         Top             =   8025
         Width           =   4230
         Begin MwfCtl.MWCTL_BTN1 CmdBSGo 
            Height          =   495
            Left            =   3420
            TabIndex        =   45
            Top             =   390
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
            Index           =   64
            Left            =   990
            TabIndex        =   42
            ToolTipText     =   "Enter Bag Year"
            Top             =   480
            Width           =   330
            _ExtentX        =   582
            _ExtentY        =   503
            MaxLength       =   2
            DataType        =   4
            ReCalcParent    =   "WBSBNO"
            IdName          =   "WBSBYY"
         End
         Begin MwfCtl.MWCTL_MED ATXT 
            Height          =   285
            Index           =   65
            Left            =   1440
            TabIndex        =   43
            ToolTipText     =   "Enter Bag Character"
            Top             =   480
            Width           =   570
            _ExtentX        =   1005
            _ExtentY        =   503
            MaxLength       =   3
            DataType        =   4
            ReCalcParent    =   "WBSBNO"
            IdName          =   "WBSBCHR"
         End
         Begin MwfCtl.MWCTL_MED ATXT 
            Height          =   285
            Index           =   66
            Left            =   2130
            TabIndex        =   44
            ToolTipText     =   "Enter Bag No."
            Top             =   480
            Width           =   900
            _ExtentX        =   1588
            _ExtentY        =   503
            Alignment       =   2
            Mask            =   "######0"
            MaxLength       =   7
            DataType        =   1
            ReCalcOn        =   "WBSBYY,WBSBCHR"
            IdName          =   "WBSBNO"
         End
         Begin VB.Label LblMain 
            BackStyle       =   0  'Transparent
            BorderStyle     =   1  'Fixed Single
            Caption         =   "  Bag Raw Material Summary"
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
            TabIndex        =   189
            Top             =   90
            Width           =   4260
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
            Left            =   2025
            TabIndex        =   107
            Top             =   480
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
            Left            =   1335
            TabIndex        =   106
            Top             =   480
            Width           =   105
         End
         Begin VB.Label ALBL 
            BackStyle       =   0  'Transparent
            Caption         =   "Bag No."
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
            TabIndex        =   105
            Top             =   480
            Width           =   1065
         End
      End
      Begin VB.Frame FraCpyRet 
         Height          =   2685
         Left            =   5070
         TabIndex        =   247
         Top             =   6345
         Width           =   4230
         Begin MwfCtl.MWCTL_BTN1 CmdCpyRetGo 
            Height          =   495
            Left            =   3510
            TabIndex        =   35
            Top             =   1950
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
            Index           =   119
            Left            =   1155
            TabIndex        =   21
            ToolTipText     =   "Enter The Location From Which The Bags Have To Be Copied"
            Top             =   450
            Width           =   1485
            _ExtentX        =   2619
            _ExtentY        =   503
            MaxLength       =   12
            DataType        =   4
            IdName          =   "WCRBLOC"
         End
         Begin MwfCtl.MWCTL_MED ATXT 
            Height          =   285
            Index           =   123
            Left            =   1155
            TabIndex        =   28
            ToolTipText     =   "Enter From Bag Year"
            Top             =   1590
            Width           =   360
            _ExtentX        =   635
            _ExtentY        =   503
            MaxLength       =   2
            DataType        =   4
            IdName          =   "WCRBYYFR"
         End
         Begin MwfCtl.MWCTL_MED ATXT 
            Height          =   285
            Index           =   124
            Left            =   2310
            TabIndex        =   30
            ToolTipText     =   "Enter From Bag Number"
            Top             =   1590
            Width           =   1020
            _ExtentX        =   1799
            _ExtentY        =   503
            Alignment       =   2
            Mask            =   "######0"
            MaxLength       =   7
            DataType        =   1
            IdName          =   "WCRBNOFR"
         End
         Begin MwfCtl.MWCTL_MED ATXT 
            Height          =   285
            Index           =   125
            Left            =   1605
            TabIndex        =   29
            ToolTipText     =   "Enter From Bag Character"
            Top             =   1590
            Width           =   600
            _ExtentX        =   1058
            _ExtentY        =   503
            MaxLength       =   3
            DataType        =   4
            IdName          =   "WCRBCHRFR"
         End
         Begin MwfCtl.MWCTL_MED ATXT 
            Height          =   285
            Index           =   126
            Left            =   1155
            TabIndex        =   31
            ToolTipText     =   "Enter To Bag Year"
            Top             =   1875
            Width           =   360
            _ExtentX        =   635
            _ExtentY        =   503
            MaxLength       =   2
            DataType        =   4
            IdName          =   "WCRBYYTO"
         End
         Begin MwfCtl.MWCTL_MED ATXT 
            Height          =   285
            Index           =   127
            Left            =   2310
            TabIndex        =   33
            ToolTipText     =   "Enter To Bag Number"
            Top             =   1875
            Width           =   1020
            _ExtentX        =   1799
            _ExtentY        =   503
            Alignment       =   2
            Mask            =   "######0"
            MaxLength       =   7
            DataType        =   1
            IdName          =   "WCRBNOTO"
         End
         Begin MwfCtl.MWCTL_MED ATXT 
            Height          =   285
            Index           =   128
            Left            =   1605
            TabIndex        =   32
            ToolTipText     =   "Enter To Bag Character"
            Top             =   1875
            Width           =   600
            _ExtentX        =   1058
            _ExtentY        =   503
            MaxLength       =   3
            DataType        =   4
            IdName          =   "WCRBCHRTO"
         End
         Begin MwfCtl.MWCTL_MED ATXT 
            Height          =   285
            Index           =   129
            Left            =   1860
            TabIndex        =   23
            ToolTipText     =   "Enter Bag Order Year "
            Top             =   735
            Width           =   360
            _ExtentX        =   635
            _ExtentY        =   503
            MaxLength       =   2
            DataType        =   4
            IdName          =   "WCRODYY"
         End
         Begin MwfCtl.MWCTL_MED ATXT 
            Height          =   285
            Index           =   130
            Left            =   3000
            TabIndex        =   25
            ToolTipText     =   "Enter Bag Order Number"
            Top             =   735
            Width           =   840
            _ExtentX        =   1482
            _ExtentY        =   503
            Alignment       =   2
            Mask            =   "#####0"
            MaxLength       =   6
            DataType        =   1
            IdName          =   "WCRODNO"
         End
         Begin MwfCtl.MWCTL_MED ATXT 
            Height          =   285
            Index           =   131
            Left            =   1155
            TabIndex        =   22
            ToolTipText     =   "Enter Bag Order Tc "
            Top             =   735
            Width           =   600
            _ExtentX        =   1058
            _ExtentY        =   503
            MaxLength       =   3
            DataType        =   4
            ReCalcParent    =   "WCRODCHR"
            IdName          =   "WCRODTC"
         End
         Begin MwfCtl.MWCTL_MED ATXT 
            Height          =   285
            Index           =   132
            Left            =   2310
            TabIndex        =   24
            ToolTipText     =   "Enter Bag Order Character"
            Top             =   735
            Width           =   600
            _ExtentX        =   1058
            _ExtentY        =   503
            MaxLength       =   3
            DataType        =   4
            ReCalcOn        =   "WCRODTC"
            IdName          =   "WCRODCHR"
         End
         Begin MwfCtl.MWCTL_MED ATXT 
            Height          =   285
            Index           =   133
            Left            =   1155
            TabIndex        =   26
            ToolTipText     =   "Enter From Bag Order Serial Number"
            Top             =   1020
            Width           =   570
            _ExtentX        =   1005
            _ExtentY        =   503
            Alignment       =   2
            Mask            =   "###0"
            MaxLength       =   4
            DataType        =   1
            IdName          =   "WCRODSRFR"
         End
         Begin MwfCtl.MWCTL_MED ATXT 
            Height          =   285
            Index           =   134
            Left            =   1155
            TabIndex        =   27
            ToolTipText     =   "Enter To Bag Order Serial Number"
            Top             =   1305
            Width           =   570
            _ExtentX        =   1005
            _ExtentY        =   503
            Alignment       =   2
            Mask            =   "###0"
            MaxLength       =   4
            DataType        =   1
            IdName          =   "WCRODSRTO"
         End
         Begin MwfCtl.MWCTL_MED ATXT 
            Height          =   285
            Index           =   149
            Left            =   2955
            TabIndex        =   34
            ToolTipText     =   "Enter 'Y' to copy wts from Order and 'N' for Actual wts"
            Top             =   2160
            Width           =   360
            _ExtentX        =   635
            _ExtentY        =   503
            DataType        =   4
            IdName          =   "WCPYORDRMYN"
         End
         Begin VB.Label ALBL 
            BackStyle       =   0  'Transparent
            Caption         =   "Copy Wt from Order Or Actual"
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
            Index           =   98
            Left            =   60
            TabIndex        =   297
            Top             =   2220
            Width           =   2835
         End
         Begin VB.Label ALBL 
            BackStyle       =   0  'Transparent
            Caption         =   "Sr To"
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
            Index           =   74
            Left            =   90
            TabIndex        =   261
            Top             =   1305
            Width           =   885
         End
         Begin VB.Label ALBL 
            BackStyle       =   0  'Transparent
            Caption         =   "Sr Fr"
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
            Index           =   73
            Left            =   90
            TabIndex        =   260
            Top             =   1020
            Width           =   885
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
            Index           =   94
            Left            =   90
            TabIndex        =   259
            Top             =   450
            Width           =   735
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
            Index           =   93
            Left            =   90
            TabIndex        =   258
            Top             =   1590
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
            Index           =   92
            Left            =   1530
            TabIndex        =   257
            Top             =   1590
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
            Index           =   91
            Left            =   2235
            TabIndex        =   256
            Top             =   1590
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
            Index           =   87
            Left            =   90
            TabIndex        =   255
            Top             =   1875
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
            Index           =   86
            Left            =   1530
            TabIndex        =   254
            Top             =   1875
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
            Index           =   85
            Left            =   2235
            TabIndex        =   253
            Top             =   1875
            Width           =   105
         End
         Begin VB.Label ALBL 
            BackStyle       =   0  'Transparent
            Caption         =   "Order "
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
            Index           =   83
            Left            =   90
            TabIndex        =   252
            Top             =   735
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
            Index           =   82
            Left            =   1770
            TabIndex        =   251
            Top             =   735
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
            Index           =   81
            Left            =   2235
            TabIndex        =   250
            Top             =   735
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
            Index           =   80
            Left            =   2925
            TabIndex        =   249
            Top             =   735
            Width           =   105
         End
         Begin VB.Label LblMain 
            BackStyle       =   0  'Transparent
            BorderStyle     =   1  'Fixed Single
            Caption         =   "  Copy RM Details Of Return Pieces"
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
            Index           =   15
            Left            =   -30
            TabIndex        =   248
            Top             =   90
            Width           =   4260
         End
      End
      Begin VB.Frame FraTot 
         Height          =   2250
         Left            =   5828
         TabIndex        =   46
         Top             =   6780
         Width           =   2745
         Begin MwfCtl.MWCTL_MED ATXT 
            Height          =   285
            Index           =   26
            Left            =   1560
            TabIndex        =   108
            ToolTipText     =   "Total Quantiyt In The Bags"
            Top             =   645
            Width           =   780
            _ExtentX        =   1376
            _ExtentY        =   503
            Alignment       =   2
            Enabled         =   0   'False
            Mask            =   "###0.0"
            MaxLength       =   6
            DataType        =   2
            IdName          =   "WTOTBQTY"
         End
         Begin MwfCtl.MWCTL_MED ATXT 
            Height          =   285
            Index           =   27
            Left            =   1560
            TabIndex        =   109
            ToolTipText     =   "Total Rejection Quantity"
            Top             =   1500
            Width           =   750
            _ExtentX        =   1323
            _ExtentY        =   503
            Alignment       =   2
            Enabled         =   0   'False
            Mask            =   "-###0.0"
            MaxLength       =   7
            DataType        =   2
            IdName          =   "WTOTRJQTY"
         End
         Begin MwfCtl.MWCTL_MED ATXT 
            Height          =   285
            Index           =   28
            Left            =   1560
            TabIndex        =   110
            ToolTipText     =   "Total Number Of Bags"
            Top             =   360
            Width           =   930
            _ExtentX        =   1640
            _ExtentY        =   503
            Alignment       =   2
            Enabled         =   0   'False
            Mask            =   "#######0"
            MaxLength       =   8
            DataType        =   1
            IdName          =   "WTOTBAGS"
         End
         Begin MwfCtl.MWCTL_MED ATXT 
            Height          =   285
            Index           =   29
            Left            =   1560
            TabIndex        =   111
            ToolTipText     =   "Total Gross Weight"
            Top             =   1215
            Width           =   1125
            _ExtentX        =   1984
            _ExtentY        =   503
            Alignment       =   2
            Enabled         =   0   'False
            Mask            =   "#####0.000"
            MaxLength       =   10
            DataType        =   2
            IdName          =   "WTOTBGRWT"
         End
         Begin MwfCtl.MWCTL_MED ATXT 
            Height          =   285
            Index           =   31
            Left            =   1560
            TabIndex        =   112
            ToolTipText     =   "Total Rejection Part Quantity"
            Top             =   1785
            Width           =   450
            _ExtentX        =   794
            _ExtentY        =   503
            Alignment       =   2
            Enabled         =   0   'False
            Mask            =   "-####0"
            MaxLength       =   6
            DataType        =   1
            IdName          =   "WTOTRJPTQTY"
         End
         Begin MwfCtl.MWCTL_MED ATXT 
            Height          =   285
            Index           =   32
            Left            =   1560
            TabIndex        =   113
            ToolTipText     =   "Total Part Quantity In The Bags"
            Top             =   930
            Width           =   450
            _ExtentX        =   794
            _ExtentY        =   503
            Alignment       =   2
            Enabled         =   0   'False
            Mask            =   "-####0"
            MaxLength       =   6
            DataType        =   1
            IdName          =   "WTOTBPTQTY"
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
            Index           =   1
            Left            =   -30
            TabIndex        =   181
            Top             =   90
            Width           =   2745
         End
         Begin VB.Label LblTotRjPtQty 
            BackStyle       =   0  'Transparent
            Caption         =   "Tot Rj Prt Qty"
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
            TabIndex        =   124
            Top             =   1785
            Width           =   1425
         End
         Begin VB.Label LBLTotRjQty 
            BackStyle       =   0  'Transparent
            Caption         =   "Tot Rj Qty"
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
            TabIndex        =   123
            Top             =   1500
            Width           =   1425
         End
         Begin VB.Label LblTotBGrWt 
            BackStyle       =   0  'Transparent
            Caption         =   "Tot Grs Wt"
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
            TabIndex        =   122
            Top             =   1215
            Width           =   1425
         End
         Begin VB.Label LblTotBPtQty 
            BackStyle       =   0  'Transparent
            Caption         =   "Tot Prt Qty"
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
            TabIndex        =   121
            Top             =   930
            Width           =   1425
         End
         Begin VB.Label LblTotBQty 
            BackStyle       =   0  'Transparent
            Caption         =   "Tot Bag Qty"
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
            TabIndex        =   120
            Top             =   645
            Width           =   1425
         End
         Begin VB.Label LblTotBags 
            BackStyle       =   0  'Transparent
            Caption         =   "No Of Bags"
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
            TabIndex        =   119
            Top             =   360
            Width           =   1425
         End
      End
      Begin VB.Frame FraCRM 
         Height          =   2505
         Left            =   4860
         TabIndex        =   272
         Top             =   6510
         Width           =   4815
         Begin MwfCtl.MWCTL_BTN1 CmdCRMGo 
            Height          =   495
            Left            =   4140
            TabIndex        =   273
            Top             =   1860
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
            Index           =   120
            Left            =   1305
            TabIndex        =   275
            ToolTipText     =   "Enter From Bag Year"
            Top             =   705
            Width           =   360
            _ExtentX        =   635
            _ExtentY        =   503
            MaxLength       =   2
            DataType        =   4
            IdName          =   "WFGBYY"
         End
         Begin MwfCtl.MWCTL_MED ATXT 
            Height          =   285
            Index           =   121
            Left            =   2460
            TabIndex        =   277
            ToolTipText     =   "Enter From Bag Number"
            Top             =   705
            Width           =   780
            _ExtentX        =   1376
            _ExtentY        =   503
            Alignment       =   2
            Mask            =   "######0"
            MaxLength       =   7
            DataType        =   1
            IdName          =   "WFGBNO"
         End
         Begin MwfCtl.MWCTL_MED ATXT 
            Height          =   285
            Index           =   122
            Left            =   1755
            TabIndex        =   276
            ToolTipText     =   "Enter From Bag Character"
            Top             =   705
            Width           =   600
            _ExtentX        =   1058
            _ExtentY        =   503
            MaxLength       =   3
            DataType        =   4
            IdName          =   "WFGBCHR"
         End
         Begin MwfCtl.MWCTL_MED ATXT 
            Height          =   285
            Index           =   141
            Left            =   1320
            TabIndex        =   278
            ToolTipText     =   "Enter The Source Company Code"
            Top             =   1350
            Width           =   930
            _ExtentX        =   1640
            _ExtentY        =   503
            MaxLength       =   3
            DataType        =   4
            IdName          =   "WINCOCD"
         End
         Begin MwfCtl.MWCTL_MED ATXT 
            Height          =   285
            Index           =   142
            Left            =   2025
            TabIndex        =   280
            ToolTipText     =   "Enter Source Invoice Year"
            Top             =   1635
            Width           =   360
            _ExtentX        =   635
            _ExtentY        =   503
            MaxLength       =   2
            DataType        =   4
            IdName          =   "WINYY"
         End
         Begin MwfCtl.MWCTL_MED ATXT 
            Height          =   285
            Index           =   143
            Left            =   3165
            TabIndex        =   282
            ToolTipText     =   "Enter Source Invoice Number"
            Top             =   1635
            Width           =   780
            _ExtentX        =   1376
            _ExtentY        =   503
            Alignment       =   2
            Mask            =   "#####0"
            MaxLength       =   6
            DataType        =   1
            IdName          =   "WINNO"
         End
         Begin MwfCtl.MWCTL_MED ATXT 
            Height          =   285
            Index           =   144
            Left            =   1320
            TabIndex        =   279
            ToolTipText     =   "Enter Source Invoice Tc"
            Top             =   1635
            Width           =   600
            _ExtentX        =   1058
            _ExtentY        =   503
            MaxLength       =   3
            DataType        =   4
            IdName          =   "WINTC"
         End
         Begin MwfCtl.MWCTL_MED ATXT 
            Height          =   285
            Index           =   145
            Left            =   2475
            TabIndex        =   281
            ToolTipText     =   "Enter Source Invoice Character"
            Top             =   1635
            Width           =   600
            _ExtentX        =   1058
            _ExtentY        =   503
            MaxLength       =   3
            DataType        =   4
            IdName          =   "WINCHR"
         End
         Begin MwfCtl.MWCTL_MED ATXT 
            Height          =   285
            Index           =   146
            Left            =   1305
            TabIndex        =   274
            ToolTipText     =   "Specify Whether To Copy RM From Invoice Design Rm Or FG Bag ('Y' for Invoice, 'N' for FG)"
            Top             =   420
            Width           =   360
            _ExtentX        =   635
            _ExtentY        =   503
            DataType        =   4
            IdName          =   "WYN"
         End
         Begin MwfCtl.MWCTL_MED ATXT 
            Height          =   285
            Index           =   147
            Left            =   1320
            TabIndex        =   283
            ToolTipText     =   "Enter Source Invoice Serial Number"
            Top             =   1920
            Width           =   780
            _ExtentX        =   1376
            _ExtentY        =   503
            Alignment       =   2
            Mask            =   "#####0"
            MaxLength       =   6
            DataType        =   1
            IdName          =   "WINSR"
         End
         Begin VB.Label ALBL 
            BackStyle       =   0  'Transparent
            Caption         =   "Bag No"
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
            Index           =   90
            Left            =   90
            TabIndex        =   294
            Top             =   705
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
            Index           =   89
            Left            =   1680
            TabIndex        =   293
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
            Index           =   88
            Left            =   2385
            TabIndex        =   292
            Top             =   705
            Width           =   105
         End
         Begin VB.Label LblMain 
            BackStyle       =   0  'Transparent
            BorderStyle     =   1  'Fixed Single
            Caption         =   "  Copy RM From FG Bag / Invoice Design RM"
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
            TabIndex        =   291
            Top             =   90
            Width           =   4860
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
            Index           =   76
            Left            =   1935
            TabIndex        =   290
            Top             =   1635
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
            Index           =   77
            Left            =   2400
            TabIndex        =   289
            Top             =   1635
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
            Index           =   78
            Left            =   3090
            TabIndex        =   288
            Top             =   1635
            Width           =   105
         End
         Begin VB.Label ALBL 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Src CoCd"
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
            Index           =   79
            Left            =   90
            TabIndex        =   287
            Top             =   1350
            Width           =   885
         End
         Begin VB.Label ALBL 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Src Inv No"
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
            Index           =   84
            Left            =   90
            TabIndex        =   286
            Top             =   1635
            Width           =   975
         End
         Begin VB.Label ALBL 
            BackStyle       =   0  'Transparent
            Caption         =   "Invoice / FG"
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
            Index           =   95
            Left            =   120
            TabIndex        =   285
            Top             =   420
            Width           =   1245
         End
         Begin VB.Label ALBL 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Src Inv Sr"
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
            Index           =   96
            Left            =   90
            TabIndex        =   284
            Top             =   1920
            Width           =   930
         End
      End
      Begin VB.Frame FraCert 
         Height          =   1725
         Left            =   1320
         TabIndex        =   321
         Top             =   7440
         Width           =   12180
         Begin MwfCtl.MWCTL_MED ATXT 
            Height          =   285
            Index           =   172
            Left            =   2415
            TabIndex        =   322
            ToolTipText     =   "Enter Lab Certificate No"
            Top             =   405
            Width           =   5000
            _ExtentX        =   8811
            _ExtentY        =   503
            MaxLength       =   30
            ReCalcParent    =   "WIRLOTNO"
            IdName          =   "WBLABCERT"
         End
         Begin MwfCtl.MWCTL_MED ATXT 
            Height          =   285
            Index           =   174
            Left            =   2415
            TabIndex        =   323
            ToolTipText     =   "Enter Lab Comments"
            Top             =   690
            Width           =   5000
            _ExtentX        =   8811
            _ExtentY        =   503
            MaxLength       =   30
            IdName          =   "WBLABCMNT"
         End
         Begin MwfCtl.MWCTL_MED ATXT 
            Height          =   285
            Index           =   173
            Left            =   2415
            TabIndex        =   329
            ToolTipText     =   "Enter Lab Bill No"
            Top             =   975
            Width           =   7000
            _ExtentX        =   12356
            _ExtentY        =   503
            MaxLength       =   45
            ReCalcParent    =   "WIRLOTNO"
            IdName          =   "WBLABDOCNO"
         End
         Begin MwfCtl.MWCTL_MED ATXT 
            Height          =   285
            Index           =   175
            Left            =   2415
            TabIndex        =   330
            ToolTipText     =   "Enter Remarks"
            Top             =   1260
            Width           =   9500
            _ExtentX        =   16748
            _ExtentY        =   503
            MaxLength       =   120
            IdName          =   "WBLABREM"
         End
         Begin VB.Label Label1 
            BackStyle       =   0  'Transparent
            Caption         =   "Lab Comments"
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
            TabIndex        =   328
            Top             =   690
            Width           =   1905
         End
         Begin VB.Label LblMain 
            BackStyle       =   0  'Transparent
            BorderStyle     =   1  'Fixed Single
            Caption         =   "  Certification Details"
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
            Index           =   18
            Left            =   -30
            TabIndex        =   327
            Top             =   90
            Width           =   12165
         End
         Begin VB.Label ALBL 
            BackStyle       =   0  'Transparent
            Caption         =   "Lab Certificate No"
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
            Index           =   103
            Left            =   60
            TabIndex        =   326
            Top             =   405
            Width           =   1905
         End
         Begin VB.Label ALBL 
            BackStyle       =   0  'Transparent
            Caption         =   "Bill Number"
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
            Index           =   104
            Left            =   60
            TabIndex        =   325
            Top             =   975
            Width           =   1185
         End
         Begin VB.Label ALBL 
            BackStyle       =   0  'Transparent
            Caption         =   "Remarks"
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
            Index           =   105
            Left            =   60
            TabIndex        =   324
            Top             =   1260
            Width           =   1185
         End
      End
      Begin VB.Frame FraLotNo 
         Height          =   1125
         Left            =   4800
         TabIndex        =   315
         ToolTipText     =   "#0.0000"
         Top             =   8040
         Visible         =   0   'False
         Width           =   5415
         Begin MwfCtl.MWCTL_BTN1 CmdLNGo 
            Height          =   465
            Left            =   4320
            TabIndex        =   316
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
            Index           =   171
            Left            =   1140
            TabIndex        =   317
            ToolTipText     =   "Enter New Lot Number To Be Generated"
            Top             =   480
            Width           =   2715
            _ExtentX        =   4789
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
            Index           =   102
            Left            =   60
            TabIndex        =   319
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
            Index           =   17
            Left            =   -30
            TabIndex        =   318
            Top             =   90
            Width           =   3615
         End
      End
      Begin VB.Frame FraRwrk 
         Height          =   3105
         Left            =   120
         TabIndex        =   311
         Top             =   6000
         Width           =   15085
         Begin MwfCtl.MWCTL_FLX GrdRwrk 
            Height          =   2640
            Left            =   0
            TabIndex        =   312
            Top             =   360
            Width           =   9675
            _ExtentX        =   17066
            _ExtentY        =   4657
            Cols            =   7
            colname1        =   "TRWSRNO"
            heading1        =   "Sr."
            datafld1        =   "TrwSrNo"
            datatype1       =   1
            mask1           =   "#0"
            maxlength1      =   2
            style1          =   2
            tooltiptext1    =   "Enter Rework Serial Number"
            colname2        =   "TRWBYWRK"
            heading2        =   "By Wrk         "
            datafld2        =   "TrwByWrk"
            datatype2       =   4
            maxlength2      =   8
            tooltiptext2    =   "Enter Worker"
            colname3        =   "TRWQTY"
            heading3        =   "    Qty"
            datafld3        =   "TrwQty"
            datatype3       =   2
            mask3           =   "-##0.0"
            maxlength3      =   6
            tooltiptext3    =   "Enter Rework Qty"
            colname4        =   "TRWRWRKCD"
            heading4        =   "Rework Code  "
            datafld4        =   "TrwRwrkCd"
            datatype4       =   4
            recalcparent4   =   "GrdRwrk(WRWRKDESC)"
            maxlength4      =   8
            tooltiptext4    =   "Enter Rework Code"
            colname5        =   "WRWRKDESC"
            heading5        =   "Rework Reason Description                                 "
            recalcon5       =   "GrdRwrk(TRWRWRKCD)"
            maxlength5      =   50
            colname6        =   "TRWPRTKEY"
            heading6        =   "Prtn"
            datafld6        =   "TrwPrtKey"
            maxlength6      =   1
            style6          =   0
         End
         Begin VB.Label LblMain 
            BackStyle       =   0  'Transparent
            BorderStyle     =   1  'Fixed Single
            Caption         =   "  Rework Entry     "
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
            TabIndex        =   313
            Top             =   90
            Width           =   15135
         End
      End
      Begin VB.Frame FraTP 
         Height          =   3105
         Left            =   165
         TabIndex        =   38
         Top             =   5925
         Width           =   15085
         Begin MwfCtl.MWCTL_FLX GrdTP 
            Height          =   2640
            Left            =   90
            TabIndex        =   39
            Top             =   360
            Width           =   9675
            _ExtentX        =   17066
            _ExtentY        =   4657
            Cols            =   12
            colname1        =   "TPSRNO"
            heading1        =   "Sr."
            datafld1        =   "TpSrNo"
            datatype1       =   1
            mask1           =   "#0"
            maxlength1      =   2
            style1          =   2
            tooltiptext1    =   "Enter Production Serial Number"
            colname2        =   "TPBYWRK"
            heading2        =   "By Wrk         "
            datafld2        =   "TpByWrk"
            datatype2       =   4
            maxlength2      =   8
            tooltiptext2    =   "Enter Worker"
            colname3        =   "TPSUBPRC"
            heading3        =   "SubPrc  "
            datafld3        =   "TpSubPrc"
            datatype3       =   4
            recalcparent3   =   "GrdTP(TPPTS)"
            maxlength3      =   5
            tooltiptext3    =   "Enter Sub Process"
            colname4        =   "TPRMCTG"
            heading4        =   "Ctg"
            datafld4        =   "TpRmCtg"
            datatype4       =   4
            recalcparent4   =   "GrdTP(TPPTS)"
            maxlength4      =   2
            tooltiptext4    =   "Enter Raw Material Category"
            colname5        =   "TPRMSCTG"
            heading5        =   "Sub Ctg "
            datafld5        =   "TpRmSCtg"
            datatype5       =   4
            recalcparent5   =   "GrdTP(TPPTS)"
            maxlength5      =   5
            tooltiptext5    =   "Enter Raw Material Sub Category"
            colname6        =   "TPPRDCTG"
            heading6        =   "Prd Ctg "
            datafld6        =   "TpPrdCtg"
            datatype6       =   4
            recalcparent6   =   "GrdTP(TPPTS)"
            maxlength6      =   5
            tooltiptext6    =   "Enter Design Production Category"
            colname7        =   "TPSETTYP"
            heading7        =   "Set Typ           "
            datafld7        =   "TpSetTyp"
            datatype7       =   4
            recalcparent7   =   "GrdTP(TPPTS)"
            maxlength7      =   8
            tooltiptext7    =   "Enter Setting Type"
            colname8        =   "TPBQTY"
            heading8        =   "BagQty"
            datafld8        =   "TpBQty"
            datatype8       =   2
            mask8           =   "-##0.0"
            recalcparent8   =   "GrdTP(TPPTS)"
            maxlength8      =   6
            tooltiptext8    =   "Enter Quantity"
            colname9        =   "TPRMQTY"
            heading9        =   "RmQty"
            datafld9        =   "TpRmQty"
            datatype9       =   1
            mask9           =   "-###0"
            recalcparent9   =   "GrdTP(TPPTS)"
            maxlength9      =   5
            tooltiptext9    =   "Enter Raw Material Quantity"
            colname10       =   "TPPTS"
            heading10       =   "Points        "
            datafld10       =   "TpPts"
            datatype10      =   2
            mask10          =   "-####0.00"
            recalcon10      =   "GrdTP(TPSUBPRC),GrdTP(TPRMCTG),GrdTP(TPRMSCTG),GrdTP(TPSETTYP),GrdTP(TPPRDCTG),GrdTP(TPBQTY),GrdTP(TPRMQTY)"
            maxlength10     =   9
            tooltiptext10   =   "Enter Production Points"
            colname11       =   "TPPRTKEY"
            datafld11       =   "TpPrtKey"
            maxlength11     =   1
            style11         =   0
         End
         Begin VB.Label LblMain 
            BackStyle       =   0  'Transparent
            BorderStyle     =   1  'Fixed Single
            Caption         =   "  Production"
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
            Index           =   11
            Left            =   -30
            TabIndex        =   191
            Top             =   90
            Width           =   15135
         End
      End
      Begin VB.Frame FraQW 
         Height          =   1185
         Left            =   3405
         TabIndex        =   172
         Top             =   7845
         Width           =   7590
         Begin MwfCtl.MWCTL_BTN1 CmdQWOk 
            Height          =   495
            Left            =   6855
            TabIndex        =   173
            Top             =   510
            Width           =   555
            _ExtentX        =   979
            _ExtentY        =   873
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
         Begin MwfCtl.MWCTL_MED ATXT 
            Height          =   285
            Index           =   108
            Left            =   1590
            TabIndex        =   174
            ToolTipText     =   "Rm Code"
            Top             =   150
            Width           =   3015
            _ExtentX        =   5318
            _ExtentY        =   503
            Enabled         =   0   'False
            MaxLength       =   0
            DataType        =   4
            ReCalcParent    =   "GrdTxndB(TDDESC)"
            IdName          =   "WTDRMCD"
         End
         Begin MwfCtl.MWCTL_MED ATXT 
            Height          =   285
            Index           =   110
            Left            =   1590
            TabIndex        =   175
            ToolTipText     =   "Enter Raw Material Quantity For Flute Bags"
            Top             =   435
            Width           =   5175
            _ExtentX        =   9128
            _ExtentY        =   503
            MaxLength       =   255
            DataType        =   4
            ReCalcParent    =   "GrdTxndB(TDDESC)"
            IdName          =   "WTDFBRMQTY"
         End
         Begin MwfCtl.MWCTL_MED ATXT 
            Height          =   285
            Index           =   115
            Left            =   1590
            TabIndex        =   176
            ToolTipText     =   "Enter Raw Material Weight For Flute Bags"
            Top             =   720
            Width           =   5175
            _ExtentX        =   9128
            _ExtentY        =   503
            MaxLength       =   500
            DataType        =   4
            ReCalcParent    =   "GrdTxndB(TDDESC)"
            IdName          =   "WTDFBRMWT"
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
            Height          =   285
            Index           =   69
            Left            =   90
            TabIndex        =   179
            Top             =   150
            Width           =   1275
         End
         Begin VB.Label ALBL 
            BackStyle       =   0  'Transparent
            Caption         =   "Weight"
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
            Index           =   70
            Left            =   90
            TabIndex        =   178
            Top             =   720
            Width           =   1275
         End
         Begin VB.Label ALBL 
            BackStyle       =   0  'Transparent
            Caption         =   "Quantity"
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
            Index           =   72
            Left            =   90
            TabIndex        =   177
            Top             =   435
            Width           =   1275
         End
      End
   End
   Begin VB.Frame FraGBD 
      Height          =   1815
      Left            =   5100
      TabIndex        =   154
      Top             =   7215
      Width           =   3960
      Begin MwfCtl.MWCTL_BTN1 CmdGBDGo 
         Height          =   285
         Left            =   3225
         TabIndex        =   162
         Top             =   1380
         Width           =   585
         _ExtentX        =   0
         _ExtentY        =   0
         BackColor       =   -2147483633
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
         Index           =   103
         Left            =   1215
         TabIndex        =   155
         ToolTipText     =   "Enter The Location To Which The Bags Have To Be Copied"
         Top             =   360
         Width           =   930
         _ExtentX        =   1640
         _ExtentY        =   503
         MaxLength       =   3
         DataType        =   4
         IdName          =   "WGBDCOCD"
      End
      Begin MwfCtl.MWCTL_MED ATXT 
         Height          =   285
         Index           =   111
         Left            =   1920
         TabIndex        =   157
         ToolTipText     =   "Enter From Bag Order Year "
         Top             =   705
         Width           =   360
         _ExtentX        =   635
         _ExtentY        =   503
         MaxLength       =   2
         DataType        =   4
         IdName          =   "WGBDINYY"
      End
      Begin MwfCtl.MWCTL_MED ATXT 
         Height          =   285
         Index           =   112
         Left            =   3060
         TabIndex        =   159
         ToolTipText     =   "Enter From Bag Order Number"
         Top             =   705
         Width           =   780
         _ExtentX        =   1376
         _ExtentY        =   503
         Alignment       =   2
         Mask            =   "#####0"
         MaxLength       =   6
         DataType        =   1
         IdName          =   "WGBDINNO"
      End
      Begin MwfCtl.MWCTL_MED ATXT 
         Height          =   285
         Index           =   113
         Left            =   1215
         TabIndex        =   156
         ToolTipText     =   "Enter From Bag Order Tc "
         Top             =   705
         Width           =   600
         _ExtentX        =   1058
         _ExtentY        =   503
         MaxLength       =   3
         DataType        =   4
         IdName          =   "WGBDINTC"
      End
      Begin MwfCtl.MWCTL_MED ATXT 
         Height          =   285
         Index           =   114
         Left            =   2370
         TabIndex        =   158
         ToolTipText     =   "Enter From Bag Order Character"
         Top             =   705
         Width           =   600
         _ExtentX        =   1058
         _ExtentY        =   503
         MaxLength       =   3
         DataType        =   4
         IdName          =   "WGBDINCHR"
      End
      Begin MwfCtl.MWCTL_MED ATXT 
         Height          =   285
         Index           =   106
         Left            =   1215
         TabIndex        =   160
         ToolTipText     =   "Enter Location from where Diamonds and Color Stones are to be Issues to Bags"
         Top             =   1050
         Width           =   930
         _ExtentX        =   1640
         _ExtentY        =   503
         MaxLength       =   5
         DataType        =   4
         IdName          =   "WGBDDCFRLOC"
      End
      Begin MwfCtl.MWCTL_MED ATXT 
         Height          =   285
         Index           =   107
         Left            =   1215
         TabIndex        =   161
         ToolTipText     =   "Enter The Location To Which The Bags Have To Be Copied"
         Top             =   1380
         Width           =   930
         _ExtentX        =   1640
         _ExtentY        =   503
         Enabled         =   0   'False
         MaxLength       =   5
         DataType        =   4
         IdName          =   "WGBDDCTOLOC"
      End
      Begin VB.Label LblMain 
         BackStyle       =   0  'Transparent
         BorderStyle     =   1  'Fixed Single
         Caption         =   "  Generate Bag with Details"
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
         TabIndex        =   192
         Top             =   90
         Width           =   3960
      End
      Begin VB.Label ALBL 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "DC To Loc"
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
         Index           =   67
         Left            =   90
         TabIndex        =   171
         Top             =   1380
         Width           =   1110
      End
      Begin VB.Label ALBL 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "DC Fr Loc"
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
         Index           =   66
         Left            =   90
         TabIndex        =   170
         Top             =   1050
         Width           =   1035
      End
      Begin VB.Label ALBL 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Fr CoCd"
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
         Index           =   71
         Left            =   90
         TabIndex        =   163
         Top             =   360
         Width           =   855
      End
      Begin VB.Label ALBL 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Invoice No"
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
         Index           =   63
         Left            =   90
         TabIndex        =   167
         Top             =   705
         Width           =   1125
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
         Index           =   62
         Left            =   1830
         TabIndex        =   166
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
         Index           =   61
         Left            =   2295
         TabIndex        =   165
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
         Index           =   60
         Left            =   2985
         TabIndex        =   164
         Top             =   705
         Width           =   105
      End
   End
   Begin VB.Frame FraHlpSrt 
      Height          =   1785
      Left            =   4620
      TabIndex        =   359
      Top             =   7230
      Width           =   6390
      Begin VB.CheckBox ChkHSMtchOrdNo 
         Alignment       =   1  'Right Justify
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
         Height          =   240
         Left            =   180
         TabIndex        =   365
         Tag             =   "ShowPnd"
         ToolTipText     =   "Check To Match Order"
         Top             =   1470
         Width           =   1455
      End
      Begin VB.CheckBox ChkHSMtchCmCd 
         Alignment       =   1  'Right Justify
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
         Height          =   240
         Left            =   180
         TabIndex        =   364
         Tag             =   "ShowPnd"
         ToolTipText     =   "Check To Match Customer"
         Top             =   1230
         Width           =   1455
      End
      Begin VB.CheckBox ChkHSMtchSz 
         Alignment       =   1  'Right Justify
         Caption         =   "Size"
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
         Left            =   180
         TabIndex        =   363
         Tag             =   "ShowPnd"
         ToolTipText     =   "Check To Match Design Size"
         Top             =   990
         Width           =   1455
      End
      Begin VB.CheckBox ChkHSMtchSfx 
         Alignment       =   1  'Right Justify
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
         Height          =   240
         Left            =   180
         TabIndex        =   362
         Tag             =   "ShowPnd"
         ToolTipText     =   "Check To Match Design Suffix"
         Top             =   750
         Width           =   1455
      End
      Begin MwfCtl.MWCTL_MED ATXT 
         Height          =   285
         Index           =   186
         Left            =   3900
         TabIndex        =   369
         ToolTipText     =   "Enter Location Code"
         Top             =   1290
         Width           =   1530
         _ExtentX        =   2699
         _ExtentY        =   503
         MaxLength       =   8
         DataType        =   4
         ReCalcOn        =   "WHSPR"
         IdName          =   "WHSLOCCD"
      End
      Begin MwfCtl.MWCTL_MED ATXT 
         Height          =   285
         Index           =   182
         Left            =   3900
         TabIndex        =   366
         ToolTipText     =   "Enter P for Location / R for Process"
         Top             =   720
         Width           =   450
         _ExtentX        =   794
         _ExtentY        =   503
         MaxLength       =   12
         DataType        =   4
         ReCalcParent    =   "WHSLOCCD,WHSSEQ"
         IdName          =   "WHSPR"
      End
      Begin MwfCtl.MWCTL_MED ATXT 
         Height          =   285
         Index           =   183
         Left            =   3900
         TabIndex        =   368
         ToolTipText     =   "Enter Production / Process Sequence"
         Top             =   1005
         Width           =   1530
         _ExtentX        =   2699
         _ExtentY        =   503
         MaxLength       =   8
         DataType        =   4
         ReCalcOn        =   "WHSPR"
         IdName          =   "WHSSEQ"
      End
      Begin VB.Label ALBL 
         BackStyle       =   0  'Transparent
         Caption         =   "Loc/Prcs Seq"
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
         Index           =   112
         Left            =   2310
         TabIndex        =   372
         Top             =   1005
         Width           =   1425
      End
      Begin VB.Label ALBL 
         BackStyle       =   0  'Transparent
         Caption         =   "Bags Ahead Of"
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
         Index           =   111
         Left            =   2310
         TabIndex        =   370
         Top             =   1290
         Width           =   1545
      End
      Begin VB.Label ALBL 
         BackStyle       =   0  'Transparent
         Caption         =   "By Loc/Prcs"
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
         Index           =   110
         Left            =   2310
         TabIndex        =   367
         Top             =   720
         Width           =   1305
      End
      Begin VB.Label ALBL 
         BackStyle       =   0  'Transparent
         Caption         =   "Match"
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
         Index           =   109
         Left            =   1410
         TabIndex        =   361
         Top             =   450
         Width           =   705
      End
      Begin VB.Label LblMain 
         BackStyle       =   0  'Transparent
         BorderStyle     =   1  'Fixed Single
         Caption         =   "  Help Sort Option"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   270
         Index           =   20
         Left            =   -30
         TabIndex        =   360
         Top             =   90
         Width           =   6375
      End
   End
End
Attribute VB_Name = "EmrFrmTxn"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'---- General Comments on the Entry --------'
' 5 Menus. All the menus use the same form EmrFrmTxn.
'   1. 'DT' - Daily Transaction
'      This Entry is used to enter the daily RM traansactions. It is used to issue
'      Raw Materials into Bags.
'   2. 'BV' - Bag Movement
'      This Entry is used to move the bags from one department to another
'   3. 'BS' - Bag Split
'      This Entry is used to move pieces from one bag to another
'   4. 'REJ'- Rejection
'      This entry is used for rejecting full or part pieces from a bag
'      Rejection has to be done for all the pieces in the bag (This is helpful in the WIP Reports)
'   5. 'MLT'- Melting
'      This entry is used for melting an entire bag. i.e. a bag from which all the
'      pieces have been rejected.
'
' X-Txn, A-Txnd (TdxlNo= 0), B-Txnd (TdSrNo> 0), C-TxndPrd
'
'CmdTP - To show the Production Points details
'        This option also automatically enters the Production Points record. The user is
'        allowed to make changes.
'        Private sub PrdPoints
'CmdBD - To show the Bag Details
'        Private sub BagDetails
'CmdIR - To Issue Raw Materials into the B-Grid in case of a 'DT' Entry
'        Private sub IssueRm
'CmdBS - To show the Bag Raw Material Summary
'        It shows the RM Qty & Wt for each RmCtg, SubCtg or RmCode in the specified Bag
'        Private sub BagRmSummary
'CmdVS - To show the Voucher Summary
'        It shows the Location wise Issue & Receipt Wt for each RmCtg, SubCtg or RmCode
'           for the entire Voucher
'        Private sub VouchSumm
'CmdCR - Copy Raw Material option is used to fill the B-Grid in case of a 'BS' Entry
'        Private sub CopyRm
'CmdER - Empty Raw Material from Bag option fill the B-Grid in case of 'MLT' Entry
'        Private sub EmptyRm
'CmdTot- To show the totals (i.e. Total Bags, Tot BagQty, Tot PartQty, Tot GrossWt)
'        Private sub Totals
'CmdLs - To enter the Loss Record in B-Grid in 'DT' Entry
'        Here the user has to just enter the final Gross weight and the dust weight if any
'        Private sub LossEntry
'CmdCB - To Copy Bags From One Loc To Another (B-Grid) in case of a 'BV' Entry
'        Private sub CopyBag
'CmdTR - To Copy Rm From One Bag To Another (B-Grid) in case of a 'BS' Entry
'        Private sub TransferRm
'CmdRwrk - To Enter Rework reason codes
'CmdLC - To
'
'--------------------

Option Explicit
Dim moCn As MwfLib.MDOConnection   '*** (Bef speed) bef 24/11/06
Dim ms_TCoCd As String, ms_TTc As String, ms_TYy As String, ms_TChr As String
Dim ms_TcTyp As String
Dim ms_hTdAck As String
Dim wb_Valid As Boolean
Dim ms_RwrkModYN As Boolean    '*********Bhavna added rework mod yn
Dim mf_AvgRtFldChg As Boolean   ' Manoj 2.9.0 Patch
Dim mf_OLORYN  As Boolean
Dim wPic3DHt As Single
Dim wPic3DWdt As Single
Dim wPicHskHt As Single
Dim wPicHskWdt As Single

Dim wPicDPHt As Single, wPicDPWdt As Single

Dim wFraPic3DHt As Single
Dim wFraPic3DWdt As Single
Dim wFraPicHskHt As Single
Dim wFraPicHskWdt As Single

Dim wFraPic3DLft As Single
Dim wFraPic3DTop As Single
Dim wFraPicHskLft As Single
Dim wFraPicHskTop As Single
Dim wPic3DDblClk As Boolean
Dim wPicHskDblClk As Boolean

Dim wHskPth As String
Dim w3DPth As String


'*** For the Toggle Options of Command Buttons
Enum en_TxnBagFra
  PRDPTS = 0
  BagDet = 1
  IssRm = 2
  BagRmSumm = 3
  VchSumm = 4
  Tot = 5
  Loss = 6
  EmpRm = 7
  CpyRm = 8
  CpyBag = 9
  TrfRm = 10
  GenBagDet = 11
  FBQtyWt = 12    ' **** Zubin 211 **** '
  CpyFgBagRm = 13 ' Zubin 212
  CpyRet = 14     '*** (Jen 2.13)
  Rwrk = 15    '*********Bhavna
  LotNo = 16   '****Bhavna added new lot
  Cert = 17
  ImpExcel = 18
  HlpSrt = 19
  InvHd = 20  'Sv.22
End Enum

'*** Declare the variables ms_LockDt, mf_DtLocked for the Lock Date Check
Dim ms_LockDt As String
Dim ms_FutureDt As String
Dim mf_DtLocked As Boolean
Dim mf_ValidLocSeq As Boolean

Dim ms_PurgeDt As String  '****** Sachin 3.02 21-12-07 - Checking on Purge Date

Dim mi_Count As Integer
Dim wShGrsWtYN As String
'*** Arr_Cv stores the causes validation state, Arr_Ed the Ena/Disa state of Cmd Buttons
Dim mArr_Cv As Variant, mArr_Ed As Variant
'*** Jay 2.14 *** (SEO Change)
Dim ms_PwdFlag As Boolean
'Dim ms_Pwd As String
'*** Jay 2.14 *** (SEO Change)
Dim mb_PrdPtsFlg As Boolean
Dim mb_RwrkFlg As Boolean     '*******Bhavna


Dim ms_FrmCaption As String '****** Sachin 2.14.0
Dim mb_MultiCoCd As Boolean '****** Sachin 2.14.0
Dim mb_ChkTol As Boolean, mb_SeoUsr As Boolean, mb_VldSeoPwd As Boolean '****** Sachin 3.5.0 - Tolerance
Dim mb_ChkMetTol As Boolean     ' 3.11.0
Dim mb_ChkTolQty As Boolean
'Dim mb_ChkCstTol As Boolean, mb_ChkWxtTol As Boolean

Dim ms_CurrentBtn As String  '*** Jenny Color
Dim mb_CopyBag As Boolean
Dim mi_TdKey As Long
Dim ms_RmCtgs As String                 '***** Sachin 3.11.0 - Tolerance
Dim ms_SetTypOnlyfrOrdYN As String      '***** Sachin 3.11.0 - Setting Typ only from Order
Dim ms_BTyp As String
Dim mb_AddRow As Boolean      ' ***** Manali 3.10.0 - 09/05/12 - Bar Code (Ctrl+N)
Dim wStkNoDisp As String
Dim wStkNoEdit As String
Dim mb_OrLot As Boolean, ms_DmTcTyp As String
Dim ml_bsnoofcreditbags As Long
Dim ml_bsnoofdebitbags As Long

Private Sub ADC_Load()
  '*** Redim the arrays mArr_Cv and mArr_Ed to the number of command buttons
  '*** Set the First Key Control property to the field TChr
  '*** Set the First Non Key Control property to TDt
  '*** Set the Child Property Of ADC to GrdTxndA
  '*** Set the Child Property Of GrdTxndA to GrdTxndB,GrdTP
  '*** Set the Previous and Next Control Properties for the Grids GrdTxndA and GrdTxndB
  '*** Set the Hot Keys for the Grids GrdTxndA, GrdTxndB,GrdTP and grdRwrk as 'A', 'B', 'C' and 'G'
  '*** Bring the FraNKeyAll frame to the top
  '*** Call the Private Sub Procedure PropSet to set the enable and visible properties
      'for all the controls on the form based on the TcStructure file
  
  'ReDim mArr_Cv(11), mArr_Ed(11)
  'ReDim mArr_Cv(12), mArr_Ed(12)   ' **** Zubin 211 **** '
  ReDim mArr_Cv(20), mArr_Ed(20)    ' **** Bhavna 5.1.1 **** '
  
  Set moCn = ADC.Connection    '*** (Bef speed) bef 24/11/06
  
  'Sachin 2.14.0 - [Multi-Company]
  mb_MultiCoCd = IIF(moCn.GetFldVal("Select hImpComLoc From Head Where HCoCd= '" + ctSelfCoCd + "' and HCd= '" + ctSelfCmCd + "'") = "Y", True, False)
  Set ADC.FirKeyCtl = IIF(mb_MultiCoCd = True, ADC("TCoCd"), ADC("TYy"))
  'Sachin 2.14.0 - [Multi-Company]
  
  Set ADC.FirNKeyCtl = ADC("TDt")
  ADC.Child = "GrdTxndA"
'  GrdTxndA.Child = "GrdTxndB,GrdTP"
  GrdTxndA.Child = "GrdTxndB,GrdTP,GrdRwrk"   '*************bhavna
  Set GrdTxndA.PrevCtl = ADC("TDesc")
  Set GrdTxndA.NextCtl = ADC
  Set GrdTxndB.PrevCtl = GrdTxndA
  Set GrdTxndB.NextCtl = GrdTxndA
  
  Set GrdBRS.NextCtl = ADC("wBrsGrp")
  Set GrdVS.NextCtl = ADC("wVSGrp")
  
  GrdTxndA.HotKey = "A"
  GrdTxndB.HotKey = "B"
  GrdTP.HotKey = "C"
  GrdRwrk.HotKey = "G"
  FraNKeyAll.ZOrder (0)

 '*** Storing The Original Size And Position Of The Picture Controls ***
  wPic3DHt = Pic3D.Height
  wPic3DWdt = Pic3D.Width
  wPicHskHt = PicHsk.Height
  wPicHskWdt = PicHsk.Width
  
  wPicDPHt = PicDP.Height: wPicDPWdt = PicDP.Width
  
  wFraPic3DHt = FraPic3D.Height
  wFraPic3DWdt = FraPic3D.Width
  wFraPicHskHt = FraPicHsk.Height
  wFraPicHskWdt = FraPicHsk.Width
 
  wFraPic3DLft = FraPic3D.Left
  wFraPic3DTop = FraPic3D.Top
  wFraPicHskLft = FraPicHsk.Left
  wFraPicHskTop = FraPicHsk.Top
  
  wPic3DDblClk = False
  wPicHskDblClk = False
  '*** Storing The Original Size And Position Of The Picture Controls ***
  ms_TCoCd = gs_CoCd
  ms_TcTyp = IIF((UCase(ADC.MenuCd) <> UCase("BA")), ADC.MenuCd, "BV")
  
  'uma *** Repair TC
  If UCase(ADC.MenuCd) = UCase("RDT") Then
  ms_TcTyp = "DT"
  ElseIf UCase(ADC.MenuCd) = UCase("RBV") Then
  ms_TcTyp = "BV"
  End If
  
  '****** Sachin 3.02 21-12-07 - Partition Key Value to Current
  ms_TTc = moCn.GetFldVal("Select PValue from Param where PTyp='TCTYP' and " + _
               "PMCd='" + ms_TcTyp + "'")
  ms_TYy = moCn.GetFldVal("Select hYy from Head where HCoCd= '" + ms_TCoCd + "' and HCd='" + ctSelfCmCd + "'")
  ms_FrmCaption = Me.Caption
'****** Sachin 2.14.0 - Code Below also copied to ADC_NONKEYWHEN for Multi-Company

  '3.11.0
  ms_SetTypOnlyfrOrdYN = moCn.GetFldVal("Select hSetTypOnlyFrOrdYN from Head where HCoCd= '" + ms_TCoCd + "' and HCd='" + ctSelfCmCd + "'")
  
  '****Bhavna added rwrk mod yn
  ms_RwrkModYN = IIF(moCn.GetFldVal("Select hRwrkModYn From Head Where HCd='" + ctSelfCmCd + "' And HCoCd='" + ctSelfCoCd + "'") = "Y", True, False)
  If ms_RwrkModYN Then
    CmdRwrk.Visible = True
  Else
    CmdRwrk.Visible = False
  End If
' **** Zubin 211 (ms_TcTyp = "BM" added) **** '
  If ms_TcTyp = "BS" Or ms_TcTyp = "BM" Then
' **** Zubin 211**** '
    wShGrsWtYN = moCn.GetFldVal("Select hGrsWtInBsYN from Head where HCoCd= '" + ms_TCoCd + "' and HCd='" + ctSelfCmCd + "'")
    If wShGrsWtYN = "Y" Then
      moCn.Execute ("update TcStr set TsStyle = 'R' where TsTbl='TxnBag' and TsTyp in ('BS','BM') and TsSTyp = 'A' and TsIdName ='TdBGrWt'")
    ElseIf wShGrsWtYN = "N" Then
      moCn.Execute ("update TcStr set TsStyle = 'I' where TsTbl='TxnBag' and TsTyp in ('BS','BM') and TsSTyp = 'A' and TsIdName ='TdBGrWt'")
    End If
  End If
  
  Call PropSet(ms_TcTyp)
  
'/********************************* urmi ***************************/
  ms_hTdAck = moCn.GetFldVal("Select hTdAck from Head where HCoCd= '" + gs_CoCd + "' and HCd='" + ctSelfCmCd + "'")

  If UCase(ADC.MenuCd) = UCase("BV") And UCase(ms_hTdAck) = "N" Then
    GrdTxndA.ColProp("TdAck").Style = fgcinvisible
  ElseIf UCase(ADC.MenuCd) = UCase("BV") And UCase(ms_hTdAck) = "Y" Then
    GrdTxndA.ColProp("TdAck").Style = fgcReadOnly
  End If

  If UCase(ADC.MenuCd) = UCase("BV") Then
    GrdTxndA.ColProp("TdWaxWt").Style = fgcNormal
  Else
    GrdTxndA.ColProp("TdWaxWt").Style = fgcinvisible
  End If

  If UCase(ADC.MenuCd) = UCase("BA") Then
    GrdTxndA.ColProp("TdBYy").Style = fgcReadOnly
    GrdTxndA.ColProp("TdBChr").Style = fgcReadOnly
    GrdTxndA.ColProp("TdBNo").Style = fgcReadOnly
    GrdTxndA.ColProp("TdBQty").Style = fgcReadOnly
    GrdTxndA.ColProp("TdPtQty").Style = fgcReadOnly
    GrdTxndA.ColProp("TdBGrWt").Style = fgcReadOnly
    GrdTxndA.ColProp("TdFrBLoc").Style = fgcReadOnly
    GrdTxndA.ColProp("TdToBLoc").Style = fgcReadOnly
    GrdTxndA.ColProp("TdWrk").Style = fgcReadOnly
    GrdTxndA.ColProp("TdAck").Style = fgcNormal
    GrdTxndA.ColProp("TdDesc").Style = fgcReadOnly

    If UCase(ms_hTdAck) = UCase("Y") Then
        ADC("wBAckOpt").Enabled = True
        ADC("wBAckOpt").Visible = True
        LblBAckOpt.Visible = True
    Else
        GrdTxndA.ColProp("TdAck").Style = fgcReadOnly
    End If

    'urmi ModDt
    GrdTxndA.Height = GrdTxndA.Height + GrdTxndB.Height
    LblModDt.Visible = True: ADC("wModDt").Visible = True

  End If
' ####  Manoj #### Ver: 2.0.9 Patch #### Date: 01/12/2004  ###
' ####  Manoj #### Ver: 2.0.10 Patch #### Date: 09/05/2005  ###
'    If gb_SingleCoCd = True Or UCase$(Left$(gs_CoNm, 8)) <> "DIAJEWEL" Then CmdGBD.Visible = False
    '****** Sachin 3.02 21-12-07 - Current Partition
    If gb_SingleCoCd = True Or _
      moCn.RecSeek("Select * from Param Where PTyp = 'MWOPT' And PMCD = 'MULCOTRF' And PValue = 'Y' ") = False Then _
        CmdGBD.Visible = False
' ###  Manoj  ###
' ###  Manoj  ###

'/********************************* urmi ***************************/

'/********************************* Shilpa ***************************/
'??? If ADC("TtrfQty").Enabled = True Then ADC("TtrfQty") = 0#
'/********************************* Shilpa ***************************/

  ' **** Zubin 211 **** '
  If UCase(ADC.MenuCd) = "BA" Then GrdTxndA.ColProp("wBagNo").Style = fgcinvisible
  If UCase(ADC.MenuCd) <> "DT" Or moCn.GetFldVal("Select HFluteBagYN from Head where HCoCd= '" + gs_CoCd + "' and HCd= '" + ctSelfCmCd + "'") = "N" Then
    CmdQW.Visible = False
  Else
    CmdQW.Visible = True
  End If

  mf_OLORYN = IIF(moCn.GetFldVal("Select hOLORYN from Head where HCoCd= '" + ms_TCoCd + "' and HCd='" + ctSelfCmCd + "'") = "Y", True, False)

  If UCase(ADC.MenuCd) <> "DT" Or moCn.GetFldVal("Select HLotNoYN from Head where HCoCd= '" + gs_CoCd + "' and HCd= '" + ctSelfCmCd + "'") = "N" Then
    CmdLN.Visible = False
  Else
    If mf_OLORYN Then
       CmdLN.Visible = False
    Else
      CmdLN.Visible = True
    End If
  End If
  If UCase(ADC.MenuCd) <> "BV" Then
    CmdCert.Visible = False
  Else
    CmdCert.Visible = True
  End If
  
'*********Bhavna
  '*** (Jen 2.13)
  If UCase(ADC.MenuCd) <> "DT" Or moCn.GetFldVal("Select HAutomateReturns from Head where HCoCd= '" + gs_CoCd + "' and HCd= '" + ctSelfCmCd + "'") = "N" Then
    CmdCpyRet.Visible = False
  Else
    CmdCpyRet.Visible = True
  End If
  '*** (Jen 2.13)

  ' Zubin 212
  If moCn.RecSeek("select 'x' from Head where HCd='" + ctSelfCmCd + "' and HCoCd='" + gs_CoCd + "' and HLotNoYN='N'") Then
    GrdTxndB.ColProp("TdLotNo").Style = fgcinvisible
    LblLotNo.Visible = False: ADC("wIrLotNo").Visible = False
    CmdLN.Visible = False       '******bhavna added new lotno cmd btn
  ElseIf UCase(ADC.MenuCd) = UCase("BV") Or UCase(ADC.MenuCd) = UCase("BA") Or UCase(ADC.MenuCd) = UCase("REJ") Then
    GrdTxndB.ColProp("TdLotNo").Style = fgcinvisible
    LblLotNo.Visible = False: ADC("wIrLotNo").Visible = False
    CmdLN.Visible = False       '******bhavna added new lotno cmd btn
  Else
    GrdTxndB.ColProp("TdLotNo").Style = fgcNormal
    LblLotNo.Visible = True: ADC("wIrLotNo").Visible = True
  End If
  ' Zubin 212
  
  '****** Sachin 3.02.0 - Displaying Partition Key Column/Field based on Login Partition
  GrdTxndB.ColProp("TdPrtKey").Style = IIF(gs_Partition = ctCurrPrtn, fgcinvisible, fgcReadOnly)
  ADC("TPrtKey").Visible = IIF(gs_Partition = ctCurrPrtn, False, True)


  ' ***** Manali 3.10.0 - 12/04/12 - Bar Code
  LblTToBLoc.Visible = False: ADC("wTToBLoc").Enabled = False: ADC("wTToBLoc").Visible = False
  LblTRjCd.Visible = False: ADC("wTRjCd").Enabled = False: ADC("wTRjCd").Visible = False
  LblTBIr.Visible = False: ADC("wTBIr").Enabled = False: ADC("wTBIr").Visible = False
  If UCase(ADC.MenuCd) = UCase("BV") Then
    LblTToBLoc.Visible = True: ADC("wTToBLoc").Enabled = True: ADC("wTToBLoc").Visible = True
  End If
  
  If UCase(ADC.MenuCd) = UCase("REJ") Then
    LblTRjCd.Visible = True: ADC("wTRjCd").Enabled = True: ADC("wTRjCd").Visible = True
    LblTBIr.Visible = True: ADC("wTBIr").Enabled = True: ADC("wTBIr").Visible = True
    
    '4.1.0.0
    LblRejVchNo.Visible = False: LblSrFr.Visible = False: LblToSr.Visible = False
    ADC("WCBTDTC").Visible = False: ADC("WCBTDYY").Visible = False: ADC("WCBTDCHR").Visible = False: ADC("WCBTDNO").Visible = False
    ADC("WCBTDSRFR").Visible = False: ADC("WCBTDSRTO").Visible = False
    LblSlash1.Visible = False: LblSlash2.Visible = False: LblSlash3.Visible = False
  End If
  ' ***** Manali 3.10.0 - 12/04/12 - Bar Code

  mb_AddRow = True      ' ***** Manali 3.10.0 - 09/05/12 - Bar Code (Ctrl+N)
  
  ' 4.1.0.0 - RFID Module Check
  If Not ((UCase(ADC.MenuCd) = UCase("BV") Or UCase(ADC.MenuCd) = UCase("MLT") Or UCase(ADC.MenuCd) = UCase("REJ")) And _
        moCn.RecSeek("Select 'x' from Head Where HCoCd='" + gs_CoCd + "' and HCd='" + ctSelfCmCd + "' and HRfIdMod='Y' ")) Then _
    ADC("WRFIDSTR").Enabled = False ': FraCB.Width = 4630
      
  If GrdTxndB.ColProp("TDRMSZ").Style <> fgcinvisible Then
    Call SetStockSizeVisible(2, GrdTxndB.ColProp("TDRMSZ2"), ADC.Connection)
    Call SetStockSizeVisible(3, GrdTxndB.ColProp("TDRMSZ3"), ADC.Connection)
  Else
    GrdTxndB.ColProp("TDRMSZ2").Style = fgcinvisible
    GrdTxndB.ColProp("TDRMSZ3").Style = fgcinvisible
  End If
 
 If moCn.GetFldVal("Select hAvgSRtYn from Head where HCoCd= '" + gs_CoCd + "'") = "Y" Then
  'StkRt-Getting stock rate display Y\N & Edit Y\N from head table
  wStkNoDisp = moCn.GetFldVal("Select hStkRtNoDsp from Head where hcocd='" + gs_CoCd + "'")
  wStkNoEdit = moCn.GetFldVal("Select hStkRtNoEdit from Head where hcocd='" + gs_CoCd + "'")
  
  'StkRt-when NoDisplay =NO then stock rate is visible in grids,if NoDisplay =yes and AvgStkRate=yes & user donot have
  'SEO rights for the form then stock rate column is Invisible else visible for Seo user.
  If wStkNoDisp = "N" Then
    GrdTxndB.ColProp("TDRMSTKRT").Style = fgcNormal
    GrdBRS.ColProp("WBRSRMVAL").Style = fgcNormal
    GrdVS.ColProp("WVSRMVAL").Style = fgcNormal
  ElseIf wStkNoDisp = "Y" Then
    mb_SeoUsr = moCn.RecSeek("Select 'x' From UsrAccess where UaUsrCd = '" + gs_UsrCd + "' " + _
                        " and UaMnuCd = '" + ADC.MenuCd + "' and UaSeoYn= 'Y' ")
    If mb_SeoUsr = True Then
      GrdTxndB.ColProp("TDRMSTKRT").Style = fgcNormal
      GrdBRS.ColProp("WBRSRMVAL").Style = fgcNormal
      GrdVS.ColProp("WVSRMVAL").Style = fgcNormal
    End If
    If mb_SeoUsr = False Then
      GrdTxndB.ColProp("TDRMSTKRT").Style = fgcinvisible
      GrdBRS.ColProp("WBRSRMVAL").Style = fgcinvisible
      GrdVS.ColProp("WVSRMVAL").Style = fgcinvisible
    End If
  End If
 End If
 'sv.9 getting purge date from vparam
 ms_PurgeDt = GetPurgeDt(ms_TCoCd, "")
End Sub
Private Sub ADC_KeyWhen()
  '*** Disable All command buttons
  '*** Hide all frames except Frame FraNKeyAll
  '*** setting the default values for TTc,TYy and TChr from the previous entry

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
  
  If mb_MultiCoCd = False Then
    ADC("TCoCd").Enabled = False
  Else
    ADC("TCoCd").Enabled = True
  End If
  
  'If ms_TCoCd <> "" Then ADC("TCoCd") = ms_TCoCd
  '****** Sachin 2.14.0 - 17-11-2006 [Multi-Company] ******
 
  If ms_TTc <> "" Then ADC("TTc") = ms_TTc
    
  If ms_TYy <> "" Then ADC("TYy") = ms_TYy
  If ms_TChr <> "" Then ADC("TChr") = ms_TChr

  '****** Sachin 2.14.0 - [24-11-2006] - [Multi-Company] ******
  Me.Caption = GetFrmCaption(ADC("TCoCd"), ms_FrmCaption)
  '****** Sachin 2.14.0 - [24-11-2006] - [Multi-Company] ******

' ########################  Manoj #### Ver: 2.0.6 #### Date: 25/12/2003  ########################
  TxtPwd.Text = ""
  ADC("TSeoPwd") = ""     ' ***** Manali 3.8.1
' ###########################################  Manoj  ###########################################
End Sub
Private Sub ADC_KeyValid(Cancel As Boolean, ErrMsg As String)
  '*** Check whether the TYy,TChr are valid
  If ADC("TYy") = "" Then Cancel = True: ErrMsg = "Enter Voucher Year": Exit Sub
  If ADC("TChr") = "" Then Cancel = True: ErrMsg = "Enter Voucher Character": Exit Sub
End Sub
Private Sub ADC_NonKeyWhen(Cancel As Boolean, ErrMsg As String)
  '*** Store the TTc,TYy and TChr values for the next entry
  '*** Call the Private Sub Procedure CmdPropSet to set the Enabled and Visible properties
  '    for all the commandbuttons on the form based on the TcStructure file
  '*** Disable all the commandbuttons associated with the 'A' Grid
  '*** Give default values fro TDt, TFrRmLoc, TToRmLoc, TFrBLoc, TLsFrDt, TLsToDt, TBillDt,
      'TBOEDt, TUseBefDt if they are blank (Empty date gives '01/01/1980')
     
      If ADC.IsNew And (Not moCn.RecSeek("select vPSCd from vParam where " + _
                         "vPCoCd= '" + ADC("TCoCd") + "' and vPTyp= 'CHR' and " + _
                         "vPMCd = '" + ADC("TTc") + "' and vPSCd='" + ADC("TChr") + "' and vPValidYn<>'N'")) Then
       Cancel = True: ErrMsg = "Voucher Character No More Valid": ADC.AllowSave = False: Exit Sub
      Else
        ADC.AllowSave = True
      End If

  Call EnaDisaCmds(False)
  
  'ms_LockDt = moCn.GetFldVal("Select PDesc225 from Param where PTyp= 'LOCKDT' and PMCd= 'LOCKDT'")
  'ms_FutureDt = moCn.GetFldVal("Select PDesc225 from Param where PTyp= 'LOCKDT' and PMCd= 'FUTUREDT'")
  ms_TCoCd = ADC("TCoCd")
  ms_TTc = ADC("TTc")
  ms_TYy = ADC("TYy")
  ms_TChr = ADC("TChr")
  
  Call GetLockDt(ms_TCoCd, "TC", ms_TTc, ms_LockDt, ms_FutureDt)
  
 ' If ms_LockDt = "" Then ms_LockDt = "01/01/80"
 ' If ms_FutureDt = "" Then ms_FutureDt = "31/12/28"
  
  '****** Sachin 3.02.0 - Setting Partition Key to Current for new records
  If ADC.IsNew And ADC("TPrtKey") <> ctCurrPrtn Then ADC("TPrtKey") = ctCurrPrtn
  
  '*** Parameters for checking the Valid User Locs in the Txn Entries
  gs_ValidPrdLocSeq = Trim(moCn.GetFldVal("Select PDesc225 from param where PTyp= 'USR' and " + _
                                     "PMCd= '" + EmrFrmMDI.TXT_UCD + "'"))
  gs_ValidOthLocSeq = Trim(moCn.GetFldVal("Select PValue1 from param where PTyp= 'USR' and " + _
                                     "PMCd= '" + EmrFrmMDI.TXT_UCD + "'"))
  '*** Parameters for checking the Valid User Locs in the Txn Entries
  
  '****** Sachin 2.14.0 - [Multi-Company] ******
  'ms_TCoCd = adc("TCoCd")
  '****** Sachin 2.14.0 - [Multi-Company] ******
 ' ms_TTc = adc("TTc")
 ' ms_TYy = adc("TYy")
 ' ms_TChr = adc("TChr")
  Call CmdPropSet
  
  '****** Sachin 2.14.0 - [Multi-Company] - Code Below shifted From ADC_Load
  ' **** Zubin 211 (ms_TcTyp = "BM" added) **** '
  If ms_TcTyp = "BS" Or ms_TcTyp = "BM" Then
' **** Zubin 211**** '
    wShGrsWtYN = moCn.GetFldVal("Select hGrsWtInBsYN from Head where HCoCd= '" + ADC("TCoCd") + "' and HCd='" + ctSelfCmCd + "'")
    If wShGrsWtYN = "Y" Then
      moCn.Execute ("update TcStr set TsStyle = 'R' where TsTbl='TxnBag' and TsTyp in ('BS','BM') and TsSTyp = 'A' and TsIdName ='TdBGrWt'")
    ElseIf wShGrsWtYN = "N" Then
      moCn.Execute ("update TcStr set TsStyle = 'I' where TsTbl='TxnBag' and TsTyp in ('BS','BM') and TsSTyp = 'A' and TsIdName ='TdBGrWt'")
    End If
  End If
  
  If UCase(ADC.MenuCd) = UCase("BV") And UCase(ms_hTdAck) = "N" Then
    GrdTxndA.ColProp("TdAck").Style = fgcinvisible
  ElseIf UCase(ADC.MenuCd) = UCase("BV") And UCase(ms_hTdAck) = "Y" Then
    GrdTxndA.ColProp("TdAck").Style = fgcReadOnly
  End If
  
  ' **** Zubin 211 **** '
  If UCase(ADC.MenuCd) <> "DT" Or moCn.GetFldVal("Select HFluteBagYN from Head where HCoCd= '" + ADC("TCoCd") + "' and HCd= '" + ctSelfCmCd + "'") = "N" Then
    CmdQW.Visible = False
  Else
    CmdQW.Visible = True
  End If
' **** Zubin 211 **** '

'*********Bhavna
  If UCase(ADC.MenuCd) <> "DT" Or moCn.GetFldVal("Select HLotNoYN from Head where HCoCd= '" + gs_CoCd + "' and HCd= '" + ctSelfCmCd + "'") = "N" Then
    CmdLN.Visible = False
  Else
    If Not mf_OLORYN Then
      CmdLN.Visible = True
    Else
      CmdLN.Visible = False
    End If
  End If
  If UCase(ADC.MenuCd) <> "BV" Then
    CmdCert.Visible = False
  Else
    CmdCert.Visible = True
  End If
'*********Bhavna

  '*** (Jen 2.13)
  If UCase(ADC.MenuCd) <> "DT" Or moCn.GetFldVal("Select HAutomateReturns from Head where HCoCd= '" + ADC("TCoCd") + "' and HCd= '" + ctSelfCmCd + "'") = "N" Then
    CmdCpyRet.Visible = False
  Else
    CmdCpyRet.Visible = True
  End If
  '*** (Jen 2.13)
  
  '********Bhavna added new lot no
  If moCn.RecSeek("select 'x' from Head where HCd='" + ctSelfCmCd + "' and HCoCd='" + ADC("TCoCd") + "' and HLotNoYN='N'") _
              Or UCase(ADC.MenuCd) <> UCase("DT") Then
    CmdLN.Visible = False
  Else
    If Not mf_OLORYN Then
      CmdLN.Visible = True
    Else
      CmdLN.Visible = False
      
    End If
  End If
  '********Bhavna added new lot no
  
  ' Zubin 212
  If moCn.RecSeek("select 'x' from Head where HCd='" + ctSelfCmCd + "' and HCoCd='" + ADC("TCoCd") + "' and HLotNoYN='N'") Then
    GrdTxndB.ColProp("TdLotNo").Style = fgcinvisible
    LblLotNo.Visible = False: ADC("wIrLotNo").Visible = False
    CmdLN.Visible = False       '******bhavna added new lotno cmd btn
  ElseIf UCase(ADC.MenuCd) = UCase("BV") Or UCase(ADC.MenuCd) = UCase("BA") Or UCase(ADC.MenuCd) = UCase("REJ") Then
    GrdTxndB.ColProp("TdLotNo").Style = fgcinvisible
    LblLotNo.Visible = False: ADC("wIrLotNo").Visible = False
    CmdLN.Visible = False       '******bhavna added new lotno cmd btn
  Else
    GrdTxndB.ColProp("TdLotNo").Style = fgcNormal
    LblLotNo.Visible = True: ADC("wIrLotNo").Visible = True
  End If
  ' Zubin 212
  'Uma KBS
  'If UCase(adc.MenuCd) = "DT" Then CmdIE.Enabled = False: CmdIE.Visible = True
  If UCase(ADC.MenuCd) <> "DT" Then CmdIE.Visible = False: CmdIE.Enabled = False 'Uma KBS
  
  
  '****** Sachin 2.14.0 - [Multi-Company] - Code Above copied From ADC_Load
  CmdTP.Enabled = False
  CmdBD.Enabled = False
  CmdIR.Enabled = False
  
  CmdBS.Enabled = False
  CmdCR.Enabled = False
  CmdER.Enabled = False
  CmdVS.Enabled = False
  CmdTot.Enabled = False
  CmdLs.Enabled = False
  'CmdCB.Enabled = False
  CmdTR.Enabled = False
  
  CmdRwrk.Enabled = False  '******Bhavna
  CmdCert.Enabled = False  '******Bhavna
  'CmdCpyRet.Enabled = False
  
  ' **** Zubin 211  **** '
  CmdQW.Enabled = False
  mb_PrdPtsFlg = False
  mb_RwrkFlg = False     '*********Bhavna
  ' **** Zubin 211  **** '
  ' Zubin 212
  CmdCRM.Enabled = False
  ' Zubin 212
  
  ADC("TLsFrDt") = MWLib.EmptyDate
  ADC("TLsToDt") = MWLib.EmptyDate
  ADC("TBillDt") = MWLib.EmptyDate
  ADC("TBOEDt") = MWLib.EmptyDate
  ADC("TUseBefDt") = MWLib.EmptyDate
  '****************Geeta*******Emr207******
  '****** Sachin 2.13.0 [15-07-2006] - TMulDiv added ******
  'adc("TUSDYN") = ""
  ADC("TCURRCD") = ""
  ADC("TMULDIV") = ""
  '****** Sachin 2.13.0 [15-07-2006] - TMulDiv added ******
  '****************Geeta*******Emr207******
   If ADC.AddFndMode = xaddmode Then
    Dim ws_Dt As String
    ws_Dt = GetNextDate(ADC("TCoCd")) '53 company code parameter added
    ws_Dt = IIF((ws_Dt < CDate(ms_LockDt)) _
                               Or _
                (ws_Dt > CDate(ms_FutureDt)), _
                "", ws_Dt)
  
    If ws_Dt <> "" Then ADC("TDt") = ws_Dt
  
    'ADC("TDt") = moCn.SrvrDate
    ADC("TFrRmLoc") = moCn.GetFldVal("Select vPValue from vParam where " + _
                                     "vPCoCd= '" + ADC("TCoCd") + "' and vPTyp='CHR' " + _
                                     "and vPMCd= '" + ADC("TTc") + "' and vPSCd='" + ADC("TChr") + "'")
    ADC("TToRmLoc") = moCn.GetFldVal("Select vPDesc225 from vParam where " + _
                                     "vPCoCd= '" + ADC("TCoCd") + "' and vPTyp='CHR' " + _
                                     "and vPMCd= '" + ADC("TTc") + "' and vPSCd='" + ADC("TChr") + "'")
    ADC("TFrBLoc") = moCn.GetFldVal("Select vPValue1 from vParam where " + _
                                    "vPCoCd= '" + ADC("TCoCd") + "' and vPTyp='CHR' " + _
                                    "and vPMCd= '" + ADC("TTc") + "' and vPSCd='" + ADC("TChr") + "'")
                                    
    ' ***** Manali 3.8.0 -  Multi Issue-Receive
    If ADC("TTc") = "BS" Then ADC("TMultiIr") = "R"
                                    
  End If
      
' ########################  Manoj #### Ver: 2.0.6 #### Date: 25/12/2003  ########################
  If ADC("TLockYN") = "" Then ADC("TLockYN") = "N"
  '*** Jay 2.14 *** (SEO Change)
  ms_PwdFlag = moCn.RecSeek("Select PValue1 from Param where PTyp= 'SEOPWD' and PmCd= 'SEOPWD' and PValue1= '' ")
  'ms_Pwd = moCn.GetFldVal("Select PValue1 from Param where " + "Ptyp = 'SEOPWD' and PmCd = 'SEOPWD'")
  '*** Jay 2.14 *** (SEO Change)
' ###########################################  Manoj  ###########################################
 
  mf_DtLocked = IIF((ADC("TDt") <> "" And ADC("TDt") < CDate(ms_LockDt)) _
                                Or _
                   (ADC("TDt") <> "" And ADC("TDt") > CDate(ms_FutureDt)) _
                                Or _
                  (ADC("TDt") <= CDate(ms_PurgeDt)), _
                True, False)

 '*************Geeta******Emr206**  02/04/04***********
 '*************TLockYN check added*******************
  If mf_DtLocked = True Or ADC("TLockYN") = "Y" Then
    GrdTxndA.AllowAdd = False: GrdTxndA.AllowDelete = False
    GrdTxndB.AllowAdd = False: GrdTxndB.AllowDelete = False
    GrdTP.AllowAdd = False: GrdTP.AllowDelete = False
    GrdRwrk.AllowAdd = False: GrdRwrk.AllowDelete = False  '*********Bhavna
  Else
    If gs_ValidPrdLocSeq <> "" And UCase$(ADC.MenuCd) <> "BA" Then
      If ADC("TFrBLoc") <> "" And InStr(1, "," + UCase(gs_ValidPrdLocSeq) + ",", "," + UCase(ADC("TFrBLoc")) + ",") <= 0 Then
        mf_ValidLocSeq = False
        GrdTxndA.AllowAdd = False: GrdTxndA.AllowDelete = False
        GrdTxndB.AllowAdd = False: GrdTxndB.AllowDelete = False
        GrdTP.AllowAdd = False: GrdTP.AllowDelete = False
        GrdRwrk.AllowAdd = False: GrdRwrk.AllowDelete = False   '***********Bhavna
       Else
        '***
        mf_ValidLocSeq = True
        GrdTxndA.AllowAdd = True: GrdTxndA.AllowDelete = True
        GrdTxndB.AllowAdd = True: GrdTxndB.AllowDelete = True
        GrdTP.AllowAdd = True: GrdTP.AllowDelete = True
        If ms_RwrkModYN Then
            GrdRwrk.AllowAdd = True: GrdRwrk.AllowDelete = True  '***Bhavna
        End If
        '***
      End If
    ElseIf UCase$(ADC.MenuCd) <> "BA" Then
      mf_ValidLocSeq = True
      GrdTxndA.AllowAdd = True: GrdTxndA.AllowDelete = True
      GrdTxndB.AllowAdd = True: GrdTxndB.AllowDelete = True
      GrdTP.AllowAdd = True: GrdTP.AllowDelete = True
      If ms_RwrkModYN Then
        GrdRwrk.AllowAdd = True: GrdRwrk.AllowDelete = True   '**********Bhavna
      End If
    End If
  
    'Manoj 2.10.0
    If gs_ValidOthLocSeq <> "" And ms_TcTyp <> "BV" Then
    'If gs_ValidOthLocSeq <> "" Then
    'Manoj 2.10.0
      If (ADC("TFrRmLoc") <> "" And InStr(1, "," + UCase(gs_ValidOthLocSeq) + ",", "," + UCase(ADC("TFrRmLoc")) + ",") <= 0) Or _
         (ADC("TToRmLoc") <> "" And InStr(1, "," + UCase(gs_ValidOthLocSeq) + ",", "," + UCase(ADC("TToRmLoc")) + ",") <= 0) Then
        mf_ValidLocSeq = False
        GrdTxndA.AllowAdd = False: GrdTxndA.AllowDelete = False
        GrdTxndB.AllowAdd = False: GrdTxndB.AllowDelete = False
        GrdTP.AllowAdd = False: GrdTP.AllowDelete = False
        GrdRwrk.AllowAdd = False: GrdRwrk.AllowDelete = False
      Else
        '***
        If mf_ValidLocSeq = True Then
          GrdTxndA.AllowAdd = True: GrdTxndA.AllowDelete = True
          GrdTxndB.AllowAdd = True: GrdTxndB.AllowDelete = True
          GrdTP.AllowAdd = True: GrdTP.AllowDelete = True
          If ms_RwrkModYN Then
            GrdRwrk.AllowAdd = True: GrdRwrk.AllowDelete = True
          End If
        End If
        '***
      End If
    ElseIf ms_TcTyp <> "BV" Then
      If mf_ValidLocSeq = True Then
        GrdTxndA.AllowAdd = True: GrdTxndA.AllowDelete = True
        GrdTxndB.AllowAdd = True: GrdTxndB.AllowDelete = True
        GrdTP.AllowAdd = True: GrdTP.AllowDelete = True
        If ms_RwrkModYN Then
            GrdRwrk.AllowAdd = True: GrdRwrk.AllowDelete = True  '***********Bhavna
        End If
      End If
    End If
  End If
  
  '************************ urmi **********************************
  If UCase(ADC.MenuCd) = UCase("BA") Then
    GrdTxndA.AllowAdd = False
    GrdTxndA.AllowDelete = False
    CmdCB.Enabled = False
  End If
  '************************ urmi **********************************
  
  '**************** urmi NegStk ****************************'
  'adc("OmSeoPwd") = ""
  TxtPwd.Text = ""
  ADC("TSeoPwd") = ""
  '*********************************************************
  
  '*************Geeta******Emr206**  02/04/04***********
  If ADC("TLockYN") = "Y" Then
    ADC.AllowDelete = False
  Else
    ADC.AllowDelete = True
  End If
 '******************Geeta*******************************
  '@@@ mb_RmToBeEntered = True
' ########################  Manoj #### Ver: 2.0.9 Patch #### Date: 01/12/2004  ########################
' ########################  Manoj #### Ver: 2.0.10 Patch #### Date: 09/05/2005  ########################
'    If gb_SingleCoCd = True Or UCase$(Left$(gs_CoNm, 8)) <> "DIAJEWEL" Then CmdGBD.Visible = False
    If gb_SingleCoCd = True Or _
      moCn.RecSeek("Select * from Param Where PTyp = 'MWOPT' And PMCD = 'MULCOTRF' And PValue = 'Y' ") = False Then _
        CmdGBD.Visible = False
' ###########################################  Manoj  ###########################################
' ###########################################  Manoj  ###########################################

  '****** Sachin 3.5.0 - Tolerance
  mb_ChkTol = moCn.RecSeek("Select 'x' From Head Where HCoCd= '" + ADC("TCoCd") + "' and HCd='" + ctSelfCmCd + "' and HChkDCWtAtIssYN='Y'")
  mb_ChkTolQty = moCn.RecSeek("Select 'x' From Head Where HCoCd= '" + ADC("TCoCd") + "' and HCd='" + ctSelfCmCd + "' and HChkDCQtyAtIssYN='Y'")
  mb_ChkMetTol = moCn.RecSeek("Select 'x' From Head Where HCoCd= '" + ADC("TCoCd") + "' and HCd='" + ctSelfCmCd + "' and HChkGPWtAtIssYN='Y'")  '3.11.0

  '**********Bhavna
  If GrdTxndB.Row = 0 Then CmdLN.Enabled = False
  '**********Bhavna
  If ADC.MenuCd = "RDT" Then CmdBS.Visible = False
  If ADC.MenuCd = "RDT" Then CmdRwrk.Visible = False
  If ADC.MenuCd = "RBV" Then CmdBS.Visible = False
  If ADC.MenuCd = "RBV" Then CmdRwrk.Visible = False
  If ADC.MenuCd = "RBV" Then CmdCB.Visible = False
  
  'mw.84- if HOrLotYN=Y in head then RmLot grid will be visible else cant be view
  mb_OrLot = moCn.RecSeek("select 'x' from Head where HCd='" + ctSelfCmCd + "' and HCoCd='" + ADC("TCoCd") + "' and HOrLotYN='Y'")
  If mb_OrLot Then
    GrdOrLot.Enabled = True
    GrdOrLot.Visible = True
  Else
    GrdOrLot.Enabled = False
    GrdOrLot.Visible = True
  End If
  'ec.4- setting default values
  ChkHSMtchSfx.Value = Unchecked:   ChkHSMtchSz.Value = Unchecked
  ChkHSMtchCmCd.Value = Unchecked:  ChkHSMtchOrdNo.Value = Unchecked
  ADC("wHsPR") = "R":  ADC("wHsLocCd") = "":   ADC("wHsSeq") = ""
  If ADC("TtrfQty") > 0 Then
    CmdHlpSrt.Enabled = True
    'ec.4-if transfer qty>0 then TransferRmYN will be Yes else No
    ADC("wTrfYN") = "Y"
  Else
    CmdHlpSrt.Enabled = False
    ADC("wTrfYN") = "N"
  End If
  'vk.4 setting default values
  OptCpyBagFr(0).Value = True
  ADC("wCpyThruBagNo") = ""
  
  If ADC("TTc") = "BS" Then
    ml_bsnoofcreditbags = moCn.GetFldVal("Select count('a') from Txnd where " + _
                              "TdCoCd= '" + ADC("TCoCd") + "' and TdTc= '" + ADC("TTc") + "' and " + _
                              "TdYy='" + ADC("TYy") + "' and TdChr='" + ADC("TChr") + "' and " + _
                              "TdNo=" + CStr(ADC("TNo")) + " and TdSrNo= 0 and TdBDc='C'")
                              

    ml_bsnoofdebitbags = moCn.GetFldVal("Select count('a') from Txnd where " + _
                              "TdCoCd= '" + ADC("TCoCd") + "' and TdTc= '" + ADC("TTc") + "' and " + _
                              "TdYy='" + ADC("TYy") + "' and TdChr='" + ADC("TChr") + "' and " + _
                              "TdNo=" + CStr(ADC("TNo")) + " and TdSrNo= 0 and TdBDc='D'")
  End If
End Sub

Private Sub ADC_NonKeyValid(Cancel As Boolean, ErrMsg As String)
  '*** In case of a Bag Split Entry, check whether the
  '    Total Issued Bag Qty = Total Received Bag Qty
  '    Total Issued Bag Part Qty = Total Received Bag Part Qty
  '    Total Issued Rm Wt = Total Received Rm Wt
  '    Total Issued Rm Qty = Total Received Rm Qty
  '*** If Not the do not allow the user to save the entry
  
  '*** In Melting Entry all the bag should be compulsorily emptied of its contents
      'i.e. there should be records in the 'B' grid if the bag wt is not zero initially
  
  ' ****** Sachin 2.12 - 06/12/05 - SJM
  Dim wb_SlvMod As Boolean
  ' ****** Sachin 2.12 - 06/12/05 - SJM
  
  ' ****** Sachin 2.12 - 06/12/05 - SJM
  wb_SlvMod = moCn.RecSeek("Select 'x' from Head  where HCoCd= '" + ctSelfCoCd + "' and HCd= '" + ctSelfCmCd + "' And HSlvModYN='Y'")
  ' ****** Sachin 2.12 - 06/12/05 - SJM
  
  ' ***** Manali 3.03 -20/06/08 - Pd Mod - 'L'
  Dim Wb_PdMod As Boolean
  Wb_PdMod = moCn.RecSeek("Select 'x' from Head  where HCoCd= '" + ctSelfCoCd + "' and HCd= '" + ctSelfCmCd + "' And HPdModYN='Y'")
  ' ***** Manali 3.03 -20/06/08 - Pd Mod - 'L'
  '*** (Jen 2.13)
  ' **** Geeta***2.11 (ms_TcTyp = "BM" added)
  '****** Sachin 3.02 21-12-07
  If ms_TcTyp = "BS" Then
  '*** (Jen 2.13)
  '****** (Bef 2.13) If ms_TcTyp = "BS" Or ms_TcTyp = "BM" Then
  
    '??? And (ADC("TtrfQty") = "" Or ADC("TtrfQty") = 0)
    Dim wNetTdBQty As Single, wNetTdPtQty As Single, wCounter As Single
    
    '****** Net TdBQty = 0 and TdPtQty = 0 for TdSrNo=0 ***
    '****** Sachin 3.02 21-12-07 - Current Partition
    wNetTdBQty = moCn.GetFldVal("Select sum(case when TdBDc= 'C' then -TdBQty " + _
                               "when TdBDc= 'D' then TdBQty else 0 end) from Txnd where " + _
                               "TdCoCd= '" + ADC("TCoCd") + "' and TdTc='" + ADC("TTc") + "' and " + _
                               "TdYy= '" + ADC("TYy") + "' and TdChr='" + ADC("TChr") + "' and " + _
                               "TdNo= " + CStr(ADC("TNo")) + " and TdSrNo= 0 And TdPrtKey='" + ctCurrPrtn + "'")
    If Round(wNetTdBQty, 1) <> 0 Then Cancel = True: ErrMsg = "Issue, Receipt Of Bag Quantity Not Matching": Exit Sub
    
    '****** Sachin 3.02 21-12-07 - Current Partition
    wNetTdPtQty = moCn.GetFldVal("Select sum(case when TdBDc= 'C' then -TdPtQty " + _
                             "when TdBDc= 'D' then TdPtQty else 0 end) from Txnd where " + _
                             "TdCoCd= '" + ADC("TCoCd") + "' and TdTc= '" + ADC("TTc") + "' and " + _
                             "TdYy= '" + ADC("TYy") + "' and TdChr= '" + ADC("TChr") + "' and " + _
                             "TdNo= " + CStr(ADC("TNo")) + " and TdSrNo= 0 And TdPrtKey='" + ctCurrPrtn + "'")
    If wNetTdPtQty <> 0 Then Cancel = True: ErrMsg = "Issue, Receipt Of Bag Part Quantity Not Matching": Exit Sub
     
    ' **** Zubin 211 **** '
    ' Either Primary or only secondary bags can be used in Bag Split
     
    'Geeta ****2.11 Distinct PValue
    '****** Sachin 3.02 21-12-07 - Current Partition
    If moCn.GetFldVal("select count(Distinct PValue) from Param where PTyp= 'BCHR' and " + _
                            "PMCd In (Select distinct TdBChr from Txnd where TdCoCd='" + ADC("TCoCd") + "' and " + _
                            "TdTc='" + ADC("TTc") + "' and TdYy='" + ADC("TYy") + "' and " + _
                            "TdChr='" + ADC("TChr") + "' and TdNo=" + CStr(ADC("TNo")) + " and " + _
                            "TdSrNo= 0 And TdPrtKey='" + ctCurrPrtn + "') and PSCd= '' ") > 1 Then
      ErrMsg = "Bag Split Should Be Between Same Type Of Bags"
      Cancel = True
    End If
    ' **** Zubin 211 **** '
  
    '****** Net of RmWt = 0 and RmQty = 0 for each RmCd,RmSz,RmStkRt,LotNo  and for TdFrRmLoc='W' ***
    ' Zubin 212
    '****** Sachin 3.02 21-12-07 - Current Partition
    wCounter = moCn.GetFldVal("Select count('x') from Txnd where " + _
                            "TdCoCd='" + ADC("TCoCd") + "' and TdTc='" + ADC("TTc") + "' and " + _
                            "TdYy='" + ADC("TYy") + "' and TdChr='" + ADC("TChr") + "' and " + _
                            "TdNo=" + CStr(ADC("TNo")) + " and TdFrRmLoc= 'W' And TdPrtKey='" + ctCurrPrtn + "' " + _
                            "group by TdRmCd, TdLotNo, TdRmSz, TdRmSz2, TdRmSz3, TdRmStkRt " + _
                            "having Round(sum(case when TdFrRmDc='C' then -TdRmWt " + _
                            "when TdFrRmDc='D' then TdRmWt else 0 end), 4) <> 0 ")
    If wCounter <> 0 Then Cancel = True: ErrMsg = "Issue, Receipt Of Rm Weight Not Matching": Exit Sub
  
    ' Zubin 212
    '****** Sachin 3.02 21-12-07 - Current Partition
    wCounter = moCn.GetFldVal("Select count('x') from Txnd where " + _
                            "TdCoCd='" + ADC("TCoCd") + "' and TdTc='" + ADC("TTc") + "' and " + _
                            "TdYy='" + ADC("TYy") + "' and TdChr='" + ADC("TChr") + "' and " + _
                            "TdNo=" + CStr(ADC("TNo")) + " and TdFrRmLoc= 'W' And TdPrtKey='" + ctCurrPrtn + "' " + _
                            "group by TdRmCd, TdLotNo, TdRmSz, TdRmSz2, TdRmSz3, TdRmStkRt " + _
                            "having sum(case when TdFrRmDc='C' then -TdRmQty " + _
                            "when TdFrRmDc='D' then TdRmQty else 0 end) <> 0")
    If wCounter <> 0 Then Cancel = True: ErrMsg = "Issue, Receipt Of Rm Quantity Not Matching": Exit Sub
  
  '???ElseIf ms_TcTyp = "BS" And ADC("TtrfQty") <> "" And ADC("TtrfQty") <> 0# Then
  End If
 
  '*** (Jen 2.13)
  '****** Sachin 3.02 21-12-07
  If ms_TcTyp = "BM" Then
    Dim wf_Set As Single
    wf_Set = 0
    
    '*** (Bef 16/09/2006) This code is applicable only if one 'BChr' is permitted for the Issuing Bags
    'wf_Set = moCn.GetFldVal("Select min(TdRefKey) from Txnd where " + _
                  "TdCoCd= '" + ADC("TCoCd") + "' and TdTc='" + ADC("TTc") + "' and " + _
                  "TdYy= '" + ADC("TYy") + "' and TdChr='" + ADC("TChr") + "' and " + _
                  "TdNo= " + CStr(ADC("TNo")) + " and TdSrNo= 0 group By TdRefKey " + _
                  "having convert (decimal (16, 1), sum(case when TdBDc= 'C' then -TdBQty " + _
                  "when TdBDc= 'D' then TdBQty else 0 end)) <> convert (decimal (16, 1), 0) ")
    '*** (Bef 16/09/2006) This code is applicable only if one 'BChr' is permitted for the Issuing Bags

    '*** (16/09/2006) This code is applicable for Multiple BChr's for the Issuing Bags
    '****** Sachin 3.02 21-12-07 - Current Partition
    wf_Set = moCn.GetFldVal("Select min(TdRefKey) from Txnd Tc where TdCoCd= '" + ADC("TCoCd") + "' " + _
                  "and TdTc= '" + ADC("TTc") + "' and TdYy= '" + ADC("TYy") + "' and TdChr= '" + ADC("TChr") + "' " + _
                  "and TdNo= " + CStr(ADC("TNo")) + " and TdSrNo= 0 and TdBDc= 'C' And TdPrtKey='" + ctCurrPrtn + "' " + _
                  "group By TdRefKey, TdBChr, TdPrtKey, TdCoCd, TdTc, TdYy, TdChr, TdNo " + _
                  "having convert (decimal (16, 1), sum(TdBQty)) <> convert (decimal (16, 1), " + _
                  "IsNull((Select sum (Td.TdBQty) From Txnd Td where Td.TdCoCd= Tc.TdCoCd and Td.TdTc= Tc.TdTc " + _
                  "and Td.TdYy= Tc.TdYy and Td.TdChr= Tc.TdChr and Td.TdNo= Tc.TdNo and Td.TdSrNo= 0 " + _
                  "and Td.TdRefKey= Tc.TdRefKey and Td.TdBDc= 'D' And Td.TdPrtKey=Tc.TdPrtKey), 0)) ")
    '*** (16/09/2006) This code is applicable for Multiple BChr's for the Issuing Bags
    If wf_Set <> 0 Then Cancel = True: ErrMsg = "Issue, Receipt Of Bag Quantity Not Matching For Set: " & CStr(wf_Set): Exit Sub
  
    '****** Net of RmWt = 0 and RmQty = 0 for each RmCd, RmSz, RmStkRt, LotNo and for TdFrRmLoc= 'W' ***
    ' ****** Sachin 3.02 (14/09/07) Id fields added
    '****** Sachin 3.02 21-12-07 - Current Partition
    wf_Set = moCn.GetFldVal("Select min(TA.TdRefKey) from Txnd TA " + _
                            "join Txnd TB on TB.TdCoCd= TA.TdCoCd and TB.TdTc= TA.TdTc and TB.TdYy= TA.TdYy " + _
                            "and TB.TdChr= TA.TdChr and TB.TdNo= TA.TdNo and TB.TdSr= TA.TdSr and TB.TdPrtKey=TA.TdPrtKey " + _
                            "where TA.TdCoCd= '" + ADC("TCoCd") + "' and TA.TdTc= '" + ADC("TTc") + "' and " + _
                            "TA.TdYy= '" + ADC("TYy") + "' and TA.TdChr= '" + ADC("TChr") + "' and " + _
                            "TA.TdNo= " + CStr(ADC("TNo")) + " and TA.TdSrNo= 0 and TB.TdSrNo<> 0 and TB.TdFrRmLoc= 'W' And TB.TdPrtKey='" + ctCurrPrtn + "' " + _
                            "group by TA.TdRefKey, TB.TdRmCd, TB.TdLotNo, TB.TdRmSz,TB.TdRmSz2,TB.TdRmSz3, TB.TdRmStkRt " + _
                            "having convert (decimal (16, 3), sum(case when TB.TdFrRmDc= 'C' then -TB.TdRmWt " + _
                            "when TB.TdFrRmDc= 'D' then TB.TdRmWt else 0 end)) <> convert (decimal (16, 3), 0) ")
    If wf_Set <> 0 Then Cancel = True: ErrMsg = "Issue, Receipt Of Rm Weight Not Matching For Set: " & CStr(wf_Set): Exit Sub
  
    '****** Sachin 3.02 (14/09/07) Id fields added
    '****** Sachin 3.02 21-12-07 - Current Partition
    wf_Set = moCn.GetFldVal("Select min(TA.TdRefKey) from Txnd TA " + _
                            "join Txnd TB on TB.TdCoCd= TA.TdCoCd and TB.TdTc= TA.TdTc and TB.TdYy= TA.TdYy " + _
                            "and TB.TdChr= TA.TdChr and TB.TdNo= TA.TdNo and TB.TdSr= TA.TdSr and TB.TdPrtKey=TA.TdPrtKey " + _
                            "where TA.TdCoCd= '" + ADC("TCoCd") + "' and TA.TdTc= '" + ADC("TTc") + "' and " + _
                            "TA.TdYy= '" + ADC("TYy") + "' and TA.TdChr= '" + ADC("TChr") + "' and " + _
                            "TA.TdNo= " + CStr(ADC("TNo")) + " and TA.TdSrNo= 0 and TB.TdSrNo<> 0 and TB.TdFrRmLoc= 'W' and TB.TdPrtKey='" + ctCurrPrtn + "' " + _
                            "group by TA.TdRefKey, TB.TdRmCd, TB.TdLotNo, TB.TdRmSz,TB.TdRmSz2,TB.TdRmSz3, TB.TdRmStkRt " + _
                            "having sum(case when TB.TdFrRmDc= 'C' then -TB.TdRmQty " + _
                            "when TB.TdFrRmDc= 'D' then TB.TdRmQty else 0 end) <> 0 ")
  
    If wf_Set <> 0 Then Cancel = True: ErrMsg = "Issue, Receipt Of Rm Quantity Not Matching For Set: " & CStr(wf_Set): Exit Sub
  End If
  '*** (Jen 2.13)
  If (ms_TcTyp = "BM" Or ms_TcTyp = "BS") Then
    Dim wTxndASr As Single
    '****** For All entries in Txnd where TdSrNo= 0 there should be atleast one row of Txnd where TdSrNo<> 0 ***
    wTxndASr = 0

    wTxndASr = moCn.GetFldVal("Select min(TdSr) from Txnd Orig where " + _
                              "TdCoCd= '" + ADC("TCoCd") + "' and TdTc= '" + ADC("TTc") + "' and " + _
                              "TdYy='" + ADC("TYy") + "' and TdChr='" + ADC("TChr") + "' and " + _
                              "TdNo=" + CStr(ADC("TNo")) + " and TdSrNo= 0 and ((TdTc in ('BS','BM') and TdBDc='C') or TdTc='MLT') and TdBGrWt > 0 and TdPrtKey='" + ctCurrPrtn + "' and " + _
                              "not exists(Select 'x' from Txnd ForRmRows where " + _
                              "ForRmRows.TdCoCd = Orig.TdCoCd and ForRmRows.TdYy = Orig.TdYy and ForRmRows.TdTc = Orig.TdTc and ForRmRows.TdChr = Orig.TdChr and ForRmRows.TdNo = Orig.TdNo and ForRmRows.TdSr = Orig.TdSr and ForRmRows.TdSrNo <> 0 " + _
                              ")")
                          
    If wTxndASr <> 0 Then Cancel = True: ErrMsg = "Raw Materials Not Emptied For Sr: " + CStr(wTxndASr): Exit Sub
  
    'wCnt = moCn.GetFldVal("Select count(TdTc) from Txnd P where TdTc= '" + adc("TTc") + "' and " + _
                          "TdYy='" + adc("TYy") + "' and TdChr='" + adc("TChr") + "' and " + _
                          "TdNo=" + CStr(adc("TNo")) + " and TdSrNo= 0 and " + _
                          "not exists(Select 'x' from Txnd Ch where Ch.TdTc= P.TdTc and " + _
                          "Ch.TdYy= P.TdYy and Ch.TdChr= P.TdChr and Ch.TdNo= P.TdNo and " + _
                          "Ch.TdSr= P.TdSr and Ch.TdSrNo<> 0)")
    'If wCnt <> 0 Then Cancel = True: ErrMsg = "Raw Materials Not Emptied For Bag/s ": Exit Sub
  
  End If
 
  If (ms_TcTyp = "MLT") Then
    '****** For All entries in Txnd where TdSrNo= 0 there should be atleast one row of Txnd where TdSrNo<> 0 ***
    wTxndASr = 0

    wTxndASr = moCn.GetFldVal("Select min(TdSr) from Txnd where " + _
                              "TdCoCd= '" + ADC("TCoCd") + "' and TdTc= '" + ADC("TTc") + "' and " + _
                              "TdYy='" + ADC("TYy") + "' and TdChr='" + ADC("TChr") + "' and " + _
                              "TdNo=" + CStr(ADC("TNo")) + " and TdSrNo= 0 and ((TdTc in ('BS','BM') and TdBDc='C') or TdTc='MLT') and TdPrtKey='" + ctCurrPrtn + "' and " + _
                              "not exists(Select 'x' from Bag where " + _
                              "BCoCd= TdCoCd and BYy= TdBYy and " + _
                              "BChr= TdBChr and BNo= TdBNo and Str(ABS(BGrWt), 16, 3) = Str(0, 16, 3) and BPrtKey=TdPrtKey)")
                          
    If wTxndASr <> 0 Then Cancel = True: ErrMsg = "Raw Materials Not Emptied For Sr: " + CStr(wTxndASr): Exit Sub
  
    'wCnt = moCn.GetFldVal("Select count(TdTc) from Txnd P where TdTc= '" + adc("TTc") + "' and " + _
                          "TdYy='" + adc("TYy") + "' and TdChr='" + adc("TChr") + "' and " + _
                          "TdNo=" + CStr(adc("TNo")) + " and TdSrNo= 0 and " + _
                          "not exists(Select 'x' from Txnd Ch where Ch.TdTc= P.TdTc and " + _
                          "Ch.TdYy= P.TdYy and Ch.TdChr= P.TdChr and Ch.TdNo= P.TdNo and " + _
                          "Ch.TdSr= P.TdSr and Ch.TdSrNo<> 0)")
    'If wCnt <> 0 Then Cancel = True: ErrMsg = "Raw Materials Not Emptied For Bag/s ": Exit Sub
  End If
  
  ' ****** Sachin 2.12 - 09/12/05 - 'XC' Net Wt to be zero for all Entries ****** '
  '****** Sachin 3.02 21-12-07 - Checks Only on vouchers of current partition
  If (ms_TcTyp <> "BV" Or ms_TcTyp <> "REJ") Then
    ' ****** Manali 3.03 - 20/06/08 - Pd Mod - 'L' added
    wCounter = moCn.GetFldVal("Select count('x') from Txnd join RmMst on RmCd=TdRmCd " + IIF(gs_Partition = ctCurrPrtn, " And RmPrtKey=TdPrtKey ", "") + _
                            "where TdCoCd='" + ADC("TCoCd") + "' and TdTc='" + ADC("TTc") + "' and " + _
                            "TdYy='" + ADC("TYy") + "' and TdChr='" + ADC("TChr") + "' and " + _
                            "TdNo=" + CStr(ADC("TNo")) + _
                            "And TdPrtKey='" + ctCurrPrtn + "' And (TdFrRmLoc = 'XC' Or TdToRmLoc = 'XC') group by RmCtg " + _
                            "having  ROUND(ABS(sum((case when TdFrRmDc='C' then -TdRmWt " + _
                            "when TdFrRmDc='D' then TdRmWt else 0 end) * " + _
                            "(case when RmCtg='G' or RmCtg='P' " + IIF(wb_SlvMod = True, " or RmCtg='S'", "") + IIF(Wb_PdMod = True, " or RmCtg='L'", "") + _
                            " then RmPurityWt else 1 end))), 3) > 0 ")  ' ****** Silver 2.12 - 02/12/05 - SJM
    If wCounter <> 0 Then Cancel = True: ErrMsg = "Net Pure Rm Weight Should Be Zero For Each Category For 'XC' Location": Exit Sub
        
    wCounter = moCn.GetFldVal("Select count('x') from Txnd join RmMst on RmCd=TdRmCd " + IIF(gs_Partition = ctCurrPrtn, " And RmPrtKey=TdPrtKey ", "") + _
                            "where TdCoCd='" + ADC("TCoCd") + "' and TdTc='" + ADC("TTc") + "' and " + _
                            "TdYy='" + ADC("TYy") + "' and TdChr='" + ADC("TChr") + "' and " + _
                            "TdNo=" + CStr(ADC("TNo")) + _
                            "And TdPrtKey='" + ctCurrPrtn + "' And (TdFrRmLoc = 'XC' Or TdToRmLoc = 'XC') group by RmCtg " + _
                            "having sum(case when TdFrRmDc='C' then -TdRmQty " + _
                            "when TdFrRmDc='D' then TdRmQty else 0 end) <> 0 ")
    If wCounter <> 0 Then Cancel = True: ErrMsg = "Net Rm Quantity Should Be Zero For Each Category For 'XC' Location": Exit Sub
    ' ****** Sachin 2.12 - 09/12/05 - 'XC' Net Wt to be zero for all Entries ****** '
  End If
  
  ' **** Manali 3.8.0
  mi_TdKey = 0
End Sub


Private Sub ATXT_FldWhen(Index As Integer, ByVal IdName As String, _
                          Cancel As Boolean, ErrMsg As String)
  '*** Set Helps For Each Field
  '*** Do not allow the user to edit the voucher date if record exists in the 'A' grid
  '*** The help of TFrRmLoc should show all the locations other than Process locations
  '*** Similarly for the help of TToRmLoc
  '*** wBrsGrp
      'It is a working field used for entering the various Raw Material group options
      'according to which the Rm Qty & Wt totals will be displayed in the Bag Rm Summary
      'Option
  '*** wVSGrp
      'It is a working field used for entering the various Rm group options according
      'to which the Iss/Recpt Rm Qty & Wt totals will be displayed in the
      'Voucher Summary Option
    
  If mf_DtLocked = True And (IdName <> UCase("wBrsGrp") And IdName <> UCase("wVSGrp") And _
                             IdName <> UCase("wBsBYy") And IdName <> UCase("wBsBChr") And _
                             IdName <> UCase("wBsBNo")) Then
    Cancel = True: ErrMsg = "Cannot Edit Locked/Previous Partition Entry": Exit Sub
  End If
  
' ########################  Manoj #### Ver: 2.0.6 #### Date: 25/12/2003  ########################
  If UCase$(ADC("TLockYN")) = UCase$("Y") And UCase$(IdName) <> UCase$("TLockYN") And _
        (IdName <> UCase("wBrsGrp") And IdName <> UCase("wVSGrp") And _
        IdName <> UCase("wBsBYy") And IdName <> UCase("wBsBChr") And IdName <> UCase("wBsBNo")) Then
    Cancel = True: ErrMsg = "Cannot Add/Edit/Delete Locked/Previous Partition Entry": Exit Sub
  End If
' ###########################################  Manoj  ###########################################
  
  If mf_ValidLocSeq = False Then Cancel = True: ErrMsg = "The User Does Not Have Access To One Of The Locs": Exit Sub
  
  Dim wOthLocBag As String
  Select Case UCase(IdName)
  
  '****** Sachin 2.14.0 - [Multi-Company] ******
  Case Is = UCase("TCoCd")
    Call HlpList.hCoCd
  '****** Sachin 2.14.0 - [Multi-Company] ******
  
  Case Is = UCase("TYy")
    Call HlpList.PMCd("YY")
  Case Is = UCase("TChr")
  ' **** Zubin 211 **** '
    'Call HlpList.vPSCd(ADC("TCoCd"), "CHR", ADC("TTc"))
  ' Zubin 211 (Patch 03)
  'uma *** Repair Chr
    If UCase(ADC.MenuCd) = UCase("BA") Then
      Call HlpList.vPSCd(ADC("TCoCd"), "CHR", ADC("TTc"))
    ElseIf (UCase(ADC.MenuCd) = UCase("RDT") Or UCase(ADC.MenuCd) = UCase("RBV")) Then
      Call HlpList.vPSCd(ADC("TCoCd"), "CHR", ADC("TTc"), "Y", , "Y")
    ElseIf (UCase(ADC.MenuCd) = UCase("DT") Or UCase(ADC.MenuCd) = UCase("BV")) Then
      Call HlpList.vPSCd(ADC("TCoCd"), "CHR", ADC("TTc"), "Y", , "N")
    Else
      Call HlpList.vPSCd(ADC("TCoCd"), "CHR", ADC("TTc"), "Y")
    End If
  ' **** Zubin 211 **** '
  Case Is = UCase("TNo")
    Call HlpList.TNo(ADC("TCoCd"), ADC("TTc"), ADC("TYy"), ADC("TChr"))
  Case Is = UCase("TDt")
    If GrdTxndA.Rows - 1 > 0 Then Cancel = True: ErrMsg = "Cannot Edit Date": Exit Sub
  Case Is = UCase("TFrRmLoc"), UCase("WGBDDCFRLoc")
    Call HlpList.Loc(ADC("TCoCd"), , "'P'")
  Case Is = UCase("TToRmLoc"), UCase("WGBDDCTOLoc")
    Call HlpList.Loc(ADC("TCoCd"), , "'P'")
  Case Is = UCase("TFrBLoc")
    Call HlpList.Loc(ADC("TCoCd"), "'P'")
' ########################  Manoj #### Ver: 2.0.6 #### Date: 25/12/2003  ########################
  Case Is = UCase("TLockYN")
    '*** Jay 2.14 *** (SEO Change)
    'If Trim$(TxtPwd.text) <> Trim$(ms_Pwd) And ADC("TLockYN") = "Y" Then
    If ADC("TLockYN") = "Y" And ms_PwdFlag = False Then
        Cancel = True: ErrMsg = "Enter SEO Password to Edit Value": Exit Sub
    End If
    '*** Jay 2.14 *** (SEO Change)
    Call HlpList.PMCd("YN")
' ###########################################  Manoj  ###########################################
  Case Is = UCase("wBrsGrp")
    Call HlpList.PMCd("TBAGBSGR")
  Case Is = UCase("wVSGrp")
    Call HlpList.PMCd("TBAGVSGR")
  Case Is = UCase("wIrRmCd")
    Dim wOrdSr As String, wOrd() As String
    '****** Sachin 3.02 21-12-07 - Current Partition
    wOrdSr = moCn.GetFldVal("Select BOdTc+','+BOdYy+','+BOdChr+','+Ltrim(str(BOdNo))+" + _
                            "','+LTrim(str(BOdSr)) from Bag where " + _
                            "BCoCd='" + ADC("TCoCd") + "' and " + _
                            "BYy='" + GrdTxndA.Value(GrdTxndA.Row, "TdBYy") + "' and " + _
                            "BChr='" + GrdTxndA.Value(GrdTxndA.Row, "TdBChr") + "' and " + _
                            "BNo= " + CStr(GrdTxndA.Value(GrdTxndA.Row, "TdBNo")) + " And BPrtKey='" + ctCurrPrtn + "'")
    wOrd() = Split(wOrdSr, ",")
    '****** here before displaying the help you would have had to check whether five values
           'have been entered in the wOrd Array (by checking if UBound(wBaseVchSr) = 4) but
           'it is not necessary as there are trigger checks which ensure that a valid order

'????wOrd?  'Sr would be returned.
    Call HlpList.OrRmCdForIssueRm(ADC("TCoCd"), wOrd(0), wOrd(1), wOrd(2), wOrd(3), wOrd(4), , , GrdTxndA.Value(GrdTxndA.Row, "TdBChr"))
  
  Case Is = UCase("wCbToBLoc"), UCase("wCbFrBLoc")      ' **** Manali 3.6.0 - 26/10/09 - Copy Multi Bags into Melting - wCbFrBLoc
    Call HlpList.Loc(ADC("TCoCd"), "'P'")
  Case Is = UCase("wCbBYyFr"), UCase("wCbBYyTo")
    Call HlpList.PMCd("YY")
  Case Is = UCase("wCbBChrFr"), UCase("wCbBChrTo")
    Call HlpList.PMCd("BCHR")
  Case Is = UCase("wCbBNoFr")
    wOthLocBag = moCn.GetFldVal("Select hOthLocBag from Head where HCoCd= '" + ADC("TCoCd") + "' and HCd='" + ctSelfCmCd + "'")
    '****** Sachin 3.02.0 bags from current partiton only in help
    Call HlpList.BNo(ADC("TCoCd"), ADC("wCbBYyFr"), ADC("wCbBChrFr"), wOthLocBag, ADC("TFrBLoc"), , , , , , , ctCurrPrtn)
  Case Is = UCase("wCbBNoTo")
    wOthLocBag = moCn.GetFldVal("Select hOthLocBag from Head where HCoCd= '" + ADC("TCoCd") + "' and HCd='" + ctSelfCmCd + "'")
    '****** Sachin 3.02.0 bags from current partiton only in help
    Call HlpList.BNo(ADC("TCoCd"), ADC("wCbBYyTo"), ADC("wCbBChrTo"), wOthLocBag, ADC("TFrBLoc"), , , , , , , ctCurrPrtn)
  Case Is = UCase("wCbOdTcFr"), UCase("wCbOdTcTo")  'indigo.164 OS tc added
    Call HlpList.PMCd("TC", "'SO','OS'")
  Case Is = UCase("wCbOdYyFr"), UCase("wCbOdYyTo")
    Call HlpList.PMCd("YY")
  Case Is = UCase("wCbOdChrFr")
    Call HlpList.vPSCd(ADC("TCoCd"), "CHR", ADC("wCbOdTcFr"))
  Case Is = UCase("wCbOdChrTo")
    Call HlpList.vPSCd(ADC("TCoCd"), "CHR", ADC("wCbOdTcTo"))
  Case Is = UCase("wCbOdNoFr")
    '****** Sachin 3.02 24-12-07 - Current Partition Orders only
    Call HlpList.OmNo(ADC("TCoCd"), ADC("wCbOdTcFr"), ADC("wCbOdYyFr"), ADC("wCbOdChrFr"), , ctCurrPrtn)
  Case Is = UCase("wCbOdNoTo")
    '****** Sachin 3.02 24-12-07 - Order Number filter
    Call HlpList.OmNo(ADC("TCoCd"), ADC("wCbOdTcFr"), ADC("wCbOdYyFr"), ADC("wCbOdChrFr"), , ctCurrPrtn)
  Case Is = UCase("wCbOdSrFr")
    Call HlpList.OdSr(ADC("TCoCd"), ADC("wCbOdTcFr"), ADC("wCbOdYyFr"), ADC("wCbOdChrFr"), ADC("wCbOdNoFr"), , ctCurrPrtn)
  Case Is = UCase("wCbOdSrTo")
    Call HlpList.OdSr(ADC("TCoCd"), ADC("wCbOdTcTo"), ADC("wCbOdYyTo"), ADC("wCbOdChrTo"), ADC("wCbOdNoTo"), , ctCurrPrtn)
  'indigo.164 SM design added
  Case Is = UCase("wCbDmCdFr"), UCase("wCbDmCdTo")
    '****** Sachin 3.02 24-12-07 - Design Code filter
    Call HlpList.DmCd("DM','SM")
  
  ' ***** Manali 3.10.0 - 30/03/12 - Melting Copy from Rejection Vch
  Case Is = UCase("wCbTdTc")
    Call HlpList.PMCd("TC", "'REJ'")
  Case Is = UCase("wCbTdYy")
    Call HlpList.PMCd("YY")
  Case Is = UCase("wCbTdChr")
    Call HlpList.vPSCd(ADC("TCoCd"), "CHR", ADC("wCbTdTc"))
  Case Is = UCase("wCbTdNo")
    Call HlpList.TNo(ADC("TCoCd"), ADC("wCbTdTc"), ADC("wCbTdYy"), ADC("wCbTdChr"), ctCurrPrtn)
  Case Is = UCase("wCbTdSrFr"), UCase("wCbTdSrTo")
    Call HlpList.TdSr(ADC("TCoCd"), ADC("wCbTdTc"), ADC("wCbTdYy"), ADC("wCbTdChr"), ADC("wCbTdNo"))
  ' ***** Manali 3.10.0 - 30/03/12 - Melting Copy from Rejection Vch

'/********************************* urmi ***************************/
  '*** Jay 3.1.3
  Case Is = UCase("wBAckOpt"), UCase("wYn")
  '*** Jay 3.1.3
    Call HlpList.PMCd("YN")
 '/********************************* urmi ***************************/
  
  Case Is = UCase("wGBDCoCd")
    Call HlpList.hCoCd
  Case Is = UCase("wGBDInTc")
    Call HlpList.PMCd("TC", "'IN'")
  Case Is = UCase("wGBDInYy")
    Call HlpList.PMCd("YY")
  Case Is = UCase("wGBDInChr")
    Call HlpList.vPSCd(ADC("wGBDCoCd"), "CHR", ADC("wGBDInTc"))
  Case Is = UCase("wGBDInNo")
    Call HlpList.InNo(ADC("wGBDCoCd"), ADC("wGBDInTc"), ADC("wGBDInYy"), ADC("wGBDInChr"))
  
  '*** Jay 3.1.3
  Case Is = UCase("wInCoCd")
    If ADC("wInCoCd") = "N" Then Cancel = True: ErrMsg = "This Field Should Not Be Edited": Exit Sub
    Call HlpList.hCoCd
  Case Is = UCase("wInTc")
    If ADC("wInCoCd") = "N" Then Cancel = True: ErrMsg = "This Field Should Not Be Edited": Exit Sub
    Call HlpList.PMCd("TC", "'IN'")
  Case Is = UCase("wInYy")
    If ADC("wInCoCd") = "N" Then Cancel = True: ErrMsg = "This Field Should Not Be Edited": Exit Sub
    Call HlpList.PMCd("YY")
  Case Is = UCase("wInChr")
    If ADC("wInCoCd") = "N" Then Cancel = True: ErrMsg = "This Field Should Not Be Edited": Exit Sub
    Call HlpList.vPSCd(ADC("wInCoCd"), "CHR", ADC("wInTc"))
  Case Is = UCase("wInNo")
    If ADC("wInCoCd") = "N" Then Cancel = True: ErrMsg = "This Field Should Not Be Edited": Exit Sub
    Call HlpList.InNo(ADC("wInCoCd"), ADC("wInTc"), ADC("wInYy"), ADC("wInChr"))
  Case Is = UCase("wInSr")
    If UCase(IdName) = UCase("wInSr") And ADC("wInCoCd") = "N" Then Cancel = True: ErrMsg = "This Field Should Not Be Edited": Exit Sub
    Call HlpList.InSr(ADC("wInCoCd"), ADC("wInTc"), ADC("wInYy"), ADC("wInChr"), ADC("wInNo"))
  '*** Jay 3.1.3
  
  '****** Sachin 3.4.0 ******
  Case Is = UCase("WSRCCOCD")
    Call HlpList.hCoCd
  Case Is = UCase("WSRCINTC")
    Call HlpList.PMCd("TC", "'IN'")
  Case Is = UCase("WSRCINYY")
    Call HlpList.PMCd("YY")
  Case Is = UCase("WSRCINCHR")
    Call HlpList.vPSCd(ADC("wSrcCoCd"), "CHR", ADC("wSrcInTc"))
  Case Is = UCase("WSRCINNO")
    Call HlpList.InNo(ADC("wSrcCoCd"), ADC("wSrcInTc"), ADC("wSrcInYy"), ADC("wSrcInChr"))
  Case Is = UCase("WSRCFRINSR"), UCase("WSRCTOINSR")
    Call HlpList.InSr(ADC("wSrcCoCd"), ADC("wSrcInTc"), ADC("wSrcInYy"), ADC("wSrcInChr"), ADC("wSrcInNo"))
  '****** Sachin 3.4.0 ******
    
  'Manoj 2.10.0
  Case Is = UCase("wLsSecMetCd")
    ' ***** Manali 3.8.0
    If ADC("wLsSecMetAllYn") <> "N" Then Cancel = True: ErrMsg = "Loss For All RmCode For Secondary Metal Option Should Be No ": Exit Sub
    Dim wRmCtg As String
    wRmCtg = moCn.GetFldVal("Select RmCtg From RmMst Where " + IIF(gs_Partition = ctCurrPrtn, " RmPrtKey='" + ctCurrPrtn + "' and ", "") + " RmCd = '" & ADC("wLsMainMetCd") & "'")
    ' ***** Manali 3.03 -20/06/08 - Pd Mod - 'L' included
    If wRmCtg = "G" Then
      wRmCtg = "'P','S','L'"
    ElseIf wRmCtg = "P" Then
      wRmCtg = "'G','S','L'"
    ElseIf wRmCtg = "S" Then
      wRmCtg = "'G','P','L'"
    ElseIf wRmCtg = "L" Then
      wRmCtg = "'G','P','S'"
    Else
      wRmCtg = "'G','P','S','L'"
    End If
    ' ***** Manali 3.03 -20/06/08 - Pd Mod - 'L' included
    
'    Call HlpList.RmCd("N", wRmCtg)
    '****** Sachin 3.02 21-12-07 - Current Partition
    wOrdSr = moCn.GetFldVal("Select BOdTc+','+BOdYy+','+BOdChr+','+Ltrim(str(BOdNo))+" + _
                            "','+LTrim(str(BOdSr)) from Bag where " + _
                            "BCoCd='" + ADC("TCoCd") + "' and " + _
                            "BYy='" + GrdTxndA.Value(GrdTxndA.Row, "TdBYy") + "' and " + _
                            "BChr='" + GrdTxndA.Value(GrdTxndA.Row, "TdBChr") + "' and " + _
                            "BNo= " + CStr(GrdTxndA.Value(GrdTxndA.Row, "TdBNo")) + " and BPrtKey='" + ctCurrPrtn + "' ")
    wOrd() = Split(wOrdSr, ",")
    Call HlpList.OrRmCd(ADC("TCoCd"), wOrd(0), wOrd(1), wOrd(2), wOrd(3), wOrd(4), wRmCtg)
  'Manoj 2.10.0
' Zubin 212
  Case Is = UCase("wIrLotNo")
    '****** Sachin 3.02 21-12-07 - Current Partition
    ' **** Manali 3.03 - 03/07/08 - Lot No Open for 'X'
    If ADC("wIrRmCd") = "" Or moCn.RecSeek("select 'x' from RmMst where RmZ= 'N' and RmCd= '" + ADC("wIrRmCd") + "' and RmCtg IN " + ctLotNoRmCtg + " And RmPrtKey='" + ctCurrPrtn + "'") Then
      Call HlpList.LotNo(ADC("TCoCd"), ADC("TFrRmLoc"), ADC("wIrRmCd"))
    Else
      Cancel = True: ErrMsg = "Lot No. Allowed For 'D', 'C','X','G','P','S','L','M' Only": Exit Sub
    End If
  Case Is = UCase("wFgBYy")
    Call HlpList.PMCd("YY")
  Case Is = UCase("wFgBChr")
    Call HlpList.PMCd("BCHR", "'P'")
  Case Is = UCase("wFgBNo")
    '****** Sachin 3.02.0 - Bag No Help only from current partition in copy
    Call HlpList.BNo(ADC("TCoCd"), ADC("wFgBYy"), ADC("wFgBChr"), "", ctFgLoc, , , , , , , ctCurrPrtn)
' Zubin 212
  
  '***  (Jen 2.13)
  Case Is = UCase("wCrOdTc")  'indigo.164 OS tc added
    Call HlpList.PMCd("TC", "'SO','OS'")
  Case Is = UCase("wCrOdYy")
    Call HlpList.PMCd("YY")
  Case Is = UCase("wCrOdChr")
    Call HlpList.vPSCd(ADC("TCoCd"), "CHR", ADC("wCrOdTc"), "Y")
  Case Is = UCase("wCrOdNo")
    Call HlpList.OmNo(ADC("TCoCd"), ADC("wCrOdTc"), ADC("wCrOdYy"), ADC("wCrOdChr"))
  Case Is = UCase("wCrOdSrFr"), UCase("wCrOdSrTo")
    Call HlpList.OdSr(ADC("TCoCd"), ADC("wCrOdTc"), ADC("wCrOdYy"), ADC("wCrOdChr"), ADC("wCrOdNo"), False)
  Case Is = UCase("wCrBYyFr"), UCase("wCrBYyTo")
    Call HlpList.PMCd("YY")
  Case Is = UCase("wCrBChrFr"), UCase("wCrBChrTo")
    Call HlpList.PMCd("BCHR", "'P'")
  Case Is = UCase("wCrBNoFr")
    'Call HlpList.BNo(ADC("TCoCd"), ADC("wCrBYyFr"), ADC("wCrBChrFr"), "", , ADC("wCrOdTc"), ADC("wCrOdYy"), ADC("wCrOdChr"), ADC("wCrOdNo"))
    '****** Sachin 3.02.0 - Bag No Help only from current partition in copy
    Call HlpList.BNo(ADC("TCoCd"), ADC("wCrBYyFr"), ADC("wCrBChrFr"), "", , , , , , , , ctCurrPrtn)
  Case Is = UCase("wCrBNoTo")
    'Call HlpList.BNo(ADC("TCoCd"), ADC("wCrBYyTo"), ADC("wCrBChrTo"), "", , ADC("wCrOdTc"), ADC("wCrOdYy"), ADC("wCrOdChr"), ADC("wCrOdNo"))
    '****** Sachin 3.02.0 - Bag No Help only from current partition in copy
    Call HlpList.BNo(ADC("TCoCd"), ADC("wCrBYyTo"), ADC("wCrBChrTo"), "", , , , , , , , ctCurrPrtn)
  Case Is = UCase("wCrBLoc")
    Call HlpList.Loc(ADC("TCoCd"), "'P'")
  '***  (Jen 2.13)
  
  ' ******* Manali 3.5.0 - Change in Return Memo Copy
  Case Is = UCase("wCpyOrdRmYn")
    Call HlpList.PMCd("YN")

  ' **** Manali 3.8.0 - Multiple Issue/ Receipts
  Case Is = UCase("TMultiIr")
    Call HlpList.PMCd("IR")
  ' **** Manali 3.8.0 - Multiple Issue/ Receipts

  ' ***** Manali 3.8.0 - Wrk and Prd Y/N Scope added
  Case Is = UCase("wCbWrk")
    Call HlpList.vPMCd(ADC("TCoCd"), "WORK")
    
  Case Is = UCase("wCbPrdYn")
    Call HlpList.PMCd("YN")
  ' ***** Manali 3.8.0 - Wrk and Prd Y/N Scope added

  Case Is = UCase("wLsMainMetAllYn"), UCase("wLsSecMetAllYn")
    Call HlpList.PMCd("YN")
    
  ' ***** Manali 3.10.0 - 12/04/12 - Bar Code
  Case Is = UCase("wTToBLoc")
    Call HlpList.Loc(ADC("TCoCd"), "'P'")
    
  Case Is = UCase("wTRjCd")
    Call HlpList.PMCd("REJCD")
    
  Case Is = UCase("wTBIr")
    Call HlpList.PMCd("IR")
  ' ***** Manali 3.10.0 - 12/04/12 - Bar Code
    
  Case Is = UCase("WBRSVALTYP"), UCase("WVSVALTYP")     '4.1.2
    Call HlpList.PMCd("WIPVAL")
        
    '*********Bhavna cpy thru bag no
    Case Is = UCase("wCPYTHRUBAGNO")
        If ADC("WCPYTHRUBAGNO") = "" Then
            ADC("WCPYTHRUBAGNO") = "N"
        End If
  Case Is = UCase("wRfIdStr")
    If ADC("wCpyThruBagNo") = "Y" Then
        wOthLocBag = moCn.GetFldVal("Select hOthLocBag from Head where HCoCd= '" + ADC("TCoCd") + "' and HCd='" + ctSelfCmCd + "'")
        HlpList.MultiSelect = True
'        Call HlpList.BagNo(adc("TCoCd"), wOthLocBag, adc("TFrBLoc"), , , , , , , ctCurrPrtn, "")
    End If
 '*********Bhavna cpy thru bag no
 'ec.4- help for LocType, Location code, sequence
  Case Is = UCase("wHsPR")
    Call HlpList.PMCd("LOCTYP", , , " PMCd in ('P','R')")
  Case Is = UCase("wHsLocCd")
    Call HlpList.Loc(ADC("TCoCd"), "'" + ADC("wHsPR") + "'")
  Case Is = UCase("wHsSeq")
    Call HlpList.vPMCd(ADC("TCoCd"), IIF(ADC("wHsPR") = "P", "PRDSEQ", "PRCSSEQ"))
  'ec.4-help for transferRm YN, if records exist in grid then cannot change RmTransfer value
  Case Is = UCase("wTrfYN")
    If GrdTxndA.Rows > 1 Then Cancel = True: ErrMsg = "Cannot Edit, Child Records Exist": Exit Sub
    Call HlpList.PMCd("YN")
  End Select
End Sub
Private Sub atxt_ReCalculate(Index As Integer, ByVal IdName As String, ByVal OldVal As Variant, ByVal RowNum As Integer, ByVal SourceName As String, ByVal SourceColName As String)
  '*** TChr on TTc                  Normal
      'PDesc225('TC', adc(TTc))
  '*** wBsBNo on wBsBYy,wBsBChr     Normal
      '0
  
  Select Case UCase(IdName)
  Case Is = UCase("TChr")
    If ADC.Mode = xNorm Then _
       ADC("TChr") = moCn.GetFldVal("Select vPValue from vParam where " + _
                                    "vPCoCd= '" + ADC("TCoCd") + "' and vPTyp='DEFCHR' " + _
                                    "and vPMCd='" + ADC("TTc") + "'")
  Case Is = UCase("wBsBNo")
    If ADC.Mode = xNorm Then ADC("wBsBNo") = 0
  ' Zubin 212
  Case Is = UCase("wIrLotNo")
    If moCn.RecSeek("Select 'x' from Head where HCd= '" + ctSelfCmCd + "' and " + _
                                    "HCoCd= '" + ADC("TCoCd") + "' and HLotNoYN= 'Y'") Then
      If ADC.Mode = xNorm Then
        '****** Sachin 3.02 21-12-07 - Current Partition
        ' **** Manali 3.03 - 03/07/08 - Lot No Open for 'X'
        If ADC("wIrRmCd") <> "" And Not moCn.RecSeek("select 'x' from RmMst where RmZ= 'N' and RmCd= '" + ADC("wIrRmCd") + "' " + _
                                    "and RmCtg IN " + ctLotNoRmCtg + " and RmPrtKey='" + ctCurrPrtn + "' ") Then
          ADC("wIrLotNo") = ""
        End If
      End If
    'Else
    '  If ADC.Mode = xNorm Then ADC("wIrLotNo") = ""
    End If
  
  ' Zubin 212
  
  '***  (Jen 2.13)
  Case Is = UCase("wCrOdChr")
    If ADC.Mode = xNorm Then _
       ADC("wCrOdChr") = moCn.GetFldVal("Select vPValue from vParam where " + _
                         "vPCoCd= '" + ADC("TCoCd") + "' and vPTyp= 'DEFCHR' " + _
                         "and vPMCd= '" + ADC("wCrOdTc") + "' and vPSCd= ''")
  '***  (Jen 2.13)
  
  ' **** Manali 3.8.0
  Case Is = UCase("TMultiIr")
    If ADC.Mode = xNorm Then
        If ADC("TTc") = "BS" Then ADC("TMultiIr") = IIF(ADC("TTrfQty") > 0, "", "R")
    End If
  ' **** Manali 3.8.0
  
  ' ***** Manali 3.8.0 - Loss in more than 2 codes
  Case Is = UCase("wLsMainMetCd")
    If ADC.Mode = xNorm Then
        If ADC("TTc") = "DT" Then
            ADC("wLsMainMetCd") = IIF(ADC("wLsMainMetAllYn") = "Y", "", _
                                moCn.GetFldVal("Select OrRmCd From OrdRm Join Bag On OrOdIdNo=BOdIdNo And OrCoCd = BCoCd And OrTc = BOdTc And " & _
                                "OrYy = BOdYy And OrChr = BOdChr And OrNo = BOdNo And OrSr = BOdSr And OrPrtKey=BPrtKey " & _
                                "Where OrMainMet = 'Y' And BCoCd = '" & ADC("TCoCd") & "' And BYy = '" & GrdTxndA.Value(GrdTxndA.Row, "TdBYy") & "' And " & _
                                "BChr = '" & GrdTxndA.Value(GrdTxndA.Row, "TdBChr") & "' And BNo = " & GrdTxndA.Value(GrdTxndA.Row, "TdBNo") & " And BPrtKey='" + ctCurrPrtn + "'"))
                          
        End If
    End If
  
  Case Is = UCase("wLsSecMetCd")
    If ADC.Mode = xNorm Then
        If ADC("TTc") = "DT" And ADC("wLsSecMetAllYn") = "Y" Then ADC("wLsSecMetCd") = ""
    End If
  'EC.4- if Loc/Prcs is changed then Location code & sequence will be blank
  Case Is = UCase("wHsLocCd")
    If ADC.Mode = xNorm Then ADC("wHsLocCd") = ""
  Case Is = UCase("wHsSeq")
    If ADC.Mode = xNorm Then ADC("wHsSeq") = ""
  'ec.4- if Rm transferYN is yes then store TrfQty as 1 else 0
  Case Is = UCase("TtrfQty")
    If ADC.Mode = xNorm Then
      If ADC("wTrfYN") = "Y" Then ADC("TtrfQty") = 1 Else ADC("TtrfQty") = 0
    End If
    
  End Select
End Sub
Private Sub ATXT_FldChange(Index As Integer, ByVal IdName As String, Cancel As Boolean, ErrMsg As String, ByVal pv_NewValue As Variant)
  '*** TCoCd        Not 'ZZZ' Or Valid HCoCd From Head which the user has access to
  '*** TTc          PMCd('TC', PValue= ms_TcTyp) and the type of entry should be 'TxnBag'
  '*** TYy          PMCd('YY')
  '*** TChr         PSCd('CHR', TTc) or ''
  '*** TDt          Not Blank
  '*** TFrRmLoc     LocCd from Loc(LocTyp<> 'P') or ''
  '*** TToRmLoc     LocCd from Loc(LocTyp<> 'P') or ''
  '*** TFrBLoc      LocCd from Loc(LocTyp= 'P') or ''
  '*** wBsBYy       PMCd('YY') or ''
  '*** wBsBChr      PMCd('BCHR') or ''
  '*** wBsBNo       BNo from Bag(BYy= adc(wBsBYy) and BChr= adc(wBsBChr)) or 0
  '*** wBrsGrp      PMCd('TBAGBSGR')
  '*** wVSGrp       PMCd('TBAGVSGR')
  '*** wIrRmCd      valid OrRmCd from OrdRm(OrderNo = OrderNo of the corresponding Bag) or ''
                    'i.e. Weight can be issued for only those Rm Codes mentioned in the
                    'order responsible for the bag
  
  Select Case UCase(IdName)
  
  '****** Sachin 2.14.0 - 17-11-2006 [Multi-Company] ******
  Case Is = UCase("TCoCd")
      Cancel = (pv_NewValue = ctSelfCoCd)
      If Cancel = True Then ErrMsg = "Company Code Cannot be 'ZZZ'": Exit Sub
  
      Cancel = Not moCn.RecSeek("Select 'x' From Head Where HCd='" + ctSelfCmCd + "' And HCoCd='" + pv_NewValue + "'")
      If Cancel = True Then ErrMsg = "Invalid Company Code": Exit Sub
      If UCase(ADC.MenuCd) = "RDT" Or UCase(ADC.MenuCd) = "RBV" Then
        Cancel = Not moCn.RecSeek("Select 'x' From UsrAccess Where UaMnuCd= '" + ADC.MenuCd + "' And " + _
                         "UaUsrCd = '" + gs_UsrCd + "' And UaMnuForm='LST' And (UaValidCoCd='' Or ''','+UaValidCoCd+',''' Like '%,'+'" + pv_NewValue + "'+',%')")
      Else
        Cancel = Not moCn.RecSeek("Select 'x' From UsrAccess Where UaMnuCd= '" + ADC.MenuCd + "' And " + _
                         "UaUsrCd = '" + gs_UsrCd + "' And UaMnuForm='MDI' And (UaValidCoCd='' Or ''','+UaValidCoCd+',''' Like '%,'+'" + pv_NewValue + "'+',%')")
      End If
      If Cancel = True Then ErrMsg = "User Doesn't Have Access To This Company Code": Exit Sub
      
      If pv_NewValue <> ADC("TCoCd") Then ADC("TChr") = ""
      
      If UCase(ADC.MenuCd) = UCase("BA") Then
        ms_hTdAck = moCn.GetFldVal("Select hTdAck from Head where HCoCd= '" + pv_NewValue + "' and HCd='" + ctSelfCmCd + "'")
        If UCase(ms_hTdAck) = UCase("Y") Then
          ADC("wBAckOpt").Enabled = True: ADC("wBAckOpt").Visible = True
          LblBAckOpt.Visible = True
          GrdTxndA.ColProp("TdAck").Style = fgcNormal
        Else
          ADC("wBAckOpt").Enabled = False: ADC("wBAckOpt").Visible = False
          LblBAckOpt.Visible = False
          GrdTxndA.ColProp("TdAck").Style = fgcReadOnly
        End If
      End If
      
      Me.Caption = GetFrmCaption(pv_NewValue, ms_FrmCaption)
      'sv.9 if login company code changed, getting purge date from vparam
      ms_PurgeDt = GetPurgeDt(pv_NewValue, "")
  '****** Sachin 2.14.0 - 17-11-2006 [Multi-Company] ******
  
  Case Is = UCase("TTc")
    Cancel = Not moCn.RecSeek("select PMCd from Param where PTyp= 'TC' and " + _
                         "PMCd = '" + pv_NewValue + "' and " + _
                         "PValue= '" + ms_TcTyp + "' and " + _
                         "(Select Pvalue1 from Param where PTyp='TCTYP' and " + _
                         "PMCd= '" + ms_TcTyp + "') ='TXNBAG'")
    If Cancel = True Then ErrMsg = "Invalid Voucher Tc": Exit Sub
  
  Case Is = UCase("TYy")
    Cancel = Not moCn.RecSeek("select PMCd from Param where PTyp= 'YY' and " + _
                         "PMCd = '" + pv_NewValue + "'")
    If Cancel = True Then ErrMsg = "Invalid Voucher Year": Exit Sub
  
  Case Is = UCase("TChr")
    If UCase(ADC.MenuCd) = UCase("RDT") Or UCase(ADC.MenuCd) = UCase("RBV") Then
        Cancel = (pv_NewValue <> "") And (Not moCn.RecSeek("select vPSCd from vParam where " + _
                         "vPCoCd= '" + ADC("TCoCd") + "' and vPTyp= 'CHR' and " + _
                         "vPMCd = '" + ADC("TTc") + "' and vPSCd='" + pv_NewValue + "' and vPValue3 = 'Y'"))
        If Cancel = True Then ErrMsg = "Invalid Voucher Character or Voucher Character not assigned for Repair": Exit Sub
    Else
      If UCase(ADC.MenuCd) = UCase("DT") Or UCase(ADC.MenuCd) = UCase("BV") Then
        Cancel = (pv_NewValue <> "") And (Not moCn.RecSeek("select vPSCd from vParam where " + _
                         "vPCoCd= '" + ADC("TCoCd") + "' and vPTyp= 'CHR' and " + _
                         "vPMCd = '" + ADC("TTc") + "' and vPSCd='" + pv_NewValue + "' and vPValue3='N'"))
        If Cancel = True Then ErrMsg = "Invalid Voucher Character or Voucher Character is assigned for Repair": Exit Sub
      Else
        Cancel = (pv_NewValue <> "") And (Not moCn.RecSeek("select vPSCd from vParam where " + _
                         "vPCoCd= '" + ADC("TCoCd") + "' and vPTyp= 'CHR' and " + _
                         "vPMCd = '" + ADC("TTc") + "' and vPSCd='" + pv_NewValue + "'"))
        If Cancel = True Then ErrMsg = "Invalid Voucher Character": Exit Sub
      End If
    End If

' **** Zubin 211 **** '
'''''    ErrMsg = "Invalid Voucher Character": Exit Sub
    
   
   ' Zubin 211 (Patch 03)
    If UCase(ADC.MenuCd) <> UCase("BA") Then
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
  
  Case Is = UCase("TDt")
    If pv_NewValue = "" Then Cancel = True: ErrMsg = "Transaction Date Cannot Be Blank": Exit Sub

    If pv_NewValue < CDate(ms_LockDt) Then Cancel = True: ErrMsg = "Transaction Date Cannot Be Less Than Lock Date": Exit Sub
    If pv_NewValue > CDate(ms_FutureDt) Then Cancel = True: ErrMsg = "Transaction Date Cannot Be Greater Than Future Date": Exit Sub
    '578.39 this function not needed to check lockdate.
    'If IsDateValidAgainstLockDt(ms_TCoCd, "TC", ms_TTc, pv_NewValue) Then Cancel = True: ErrMsg = "Transaction Date Cannot Be Less Than Server Lock Date": Exit Sub

    'CC.999.13 check for purge date added
    If pv_NewValue <= CDate(ms_PurgeDt) Then Cancel = True: ErrMsg = "Transaction Date Cannot Be Less Than Or Equal To Purge Date": Exit Sub
    
    
    If moCn.GetFldVal("Select hHolChkYN From Head where HCd='ZSELF' and HCoCd='" + gs_CoCd + "'") = "Y" Then
        '53 company code added
        'If moCn.RecSeek("Select top 1 From DtTbl where DtTag='HOL' and DtCoCd in('" + ctSelfCoCd + "','" + ADC("TCoCd") + "')" + _
          " and convert(varchar(15),DtDt,3)= '" + CStr(pv_NewValue) + "' and DtHolYN='Y'")
        If IsHOLYN(ADC("TCoCd"), pv_NewValue) Then Cancel = True: ErrMsg = "Voucher Cannot be Added on Holidays": Exit Sub
    End If
  Case Is = UCase("TFrRmLoc")
    '****** Sachin 2.13.0 [16-08-2006] - ValidYN added below
    
      Cancel = pv_NewValue <> "" And (Not moCn.RecSeek("Select LocCd From Loc Where LocTyp<>'P' and LocTyp<>'R' and " + _
                          "LocTyp<>'Y' and LocTyp<>'XT' and " + _
                           "LocCoCd= '" + ADC("TCoCd") + "' and " + _
                           "LocCd = '" + pv_NewValue + "' And LocValidYN='Y'"))
      If Cancel = True Then ErrMsg = "Invalid Raw Material From Location": Exit Sub

              If ms_TcTyp = "DT" Then
              Cancel = (Not moCn.RecSeek("Select LocCd From Loc Where LocTyp<>'XC'  and  " + _
                           "LocCoCd= '" + ADC("TCoCd") + "' and " + _
                           "LocCd = '" + pv_NewValue + "' And LocValidYN='Y'"))
              If Cancel = True Then ErrMsg = "Invalid Raw Material From Location": Exit Sub

              End If


    If ADC.MenuCd = "RDT" Then
          Cancel = (pv_NewValue <> "") And (Not moCn.RecSeek("Select LocCd from Loc where " + _
              "LocCoCd= '" + ADC("TCoCd") + "' and " + _
              "LocCd = '" + pv_NewValue + "' and (LocTyp = 'XR' or LocTyp = 'S')  And LocValidYN='Y'"))
    If Cancel = True Then ErrMsg = "Raw Material From Loc Should Be XR/S LocType": Exit Sub
    End If
    Case Is = UCase("WGBDDCFRLoc")
    Cancel = (pv_NewValue <> "") And (Not moCn.RecSeek("Select LocCd from Loc where " + _
                                          "LocCoCd= '" + ADC("TCoCd") + "' and " + _
                                          "LocCd = '" + pv_NewValue + "' and LocTyp<>'P' And LocValidYn='Y'"))
    
    If Cancel = True Then ErrMsg = "Invalid Raw Material From Location": Exit Sub

    If ADC.MenuCd = "RDT" Then
          Cancel = (pv_NewValue <> "") And (Not moCn.RecSeek("Select LocCd from Loc where " + _
              "LocCoCd= '" + ADC("TCoCd") + "' and " + _
              "LocCd = '" + pv_NewValue + "' and (LocTyp = 'XR' or LocTyp = 'S')  And LocValidYN='Y'"))
    If Cancel = True Then ErrMsg = "Raw Material From Loc Should Be XR/S LocType": Exit Sub
    End If
  Case Is = UCase("TToRmLoc")
      Cancel = pv_NewValue <> "" And (Not moCn.RecSeek("Select LocCd From Loc Where LocTyp<>'P' and LocTyp<>'R' and " + _
                          "LocTyp<>'Y' and LocTyp<>'XT' and " + _
                           "LocCoCd= '" + ADC("TCoCd") + "' and " + _
                           "LocCd = '" + pv_NewValue + "' And LocValidYN='Y'"))
      If Cancel = True Then ErrMsg = "Invalid Raw Material To Location": Exit Sub

              If ms_TcTyp = "DT" Then
                  Cancel = (Not moCn.RecSeek("Select LocCd From Loc Where LocTyp<>'XC' and    " + _
                           "LocCoCd= '" + ADC("TCoCd") + "' and " + _
                           "LocCd = '" + pv_NewValue + "' And LocValidYN='Y'"))
                  If Cancel = True Then ErrMsg = "Invalid Raw Material To Location": Exit Sub

              End If

    If ADC.MenuCd = "RDT" Then
          Cancel = (pv_NewValue <> "") And (Not moCn.RecSeek("Select LocCd from Loc where " + _
              "LocCoCd= '" + ADC("TCoCd") + "' and " + _
              "LocCd = '" + pv_NewValue + "' and (LocTyp = 'XR' or LocTyp = 'S')  And LocValidYN='Y'"))
    If Cancel = True Then ErrMsg = "Raw Material To Loc Should Be XR/S LocType": Exit Sub
    End If
  
  Case Is = UCase("WGBDDCTOLoc")
    '****** Sachin 2.13.0 [16-08-2006] - ValidYN added below
    Cancel = (pv_NewValue <> "") And (Not moCn.RecSeek("Select LocCd from Loc where " + _
              "LocCoCd= '" + ADC("TCoCd") + "' and " + _
              "LocCd = '" + pv_NewValue + "' and LocTyp<>'P' And LocValidYN='Y'"))
    If Cancel = True Then ErrMsg = "Invalid Raw Material To Location": Exit Sub

    If ADC.MenuCd = "RDT" Then
          Cancel = (pv_NewValue <> "") And (Not moCn.RecSeek("Select LocCd from Loc where " + _
              "LocCoCd= '" + ADC("TCoCd") + "' and " + _
              "LocCd = '" + pv_NewValue + "' and (LocTyp = 'XR' or LocTyp = 'S')  And LocValidYN='Y'"))
    If Cancel = True Then ErrMsg = "Raw Material To Loc Should Be XR/S LocType": Exit Sub
    End If
    
  Case Is = UCase("TFrBLoc")
    If pv_NewValue = ctFgLoc Then Cancel = True: ErrMsg = "Bag Cannot Be Moved From FG Location": Exit Sub
    Cancel = (pv_NewValue <> "") And (Not moCn.RecSeek("Select LocCd from Loc where " + _
              "LocCoCd= '" + ADC("TCoCd") + "' and " + _
              "LocCd = '" + pv_NewValue + "' and LocTyp= 'P'"))
    If Cancel = True Then ErrMsg = "Invalid Bag From Location": Exit Sub
  
' ########################  Manoj #### Ver: 2.0.6 #### Date: 25/12/2003  ########################
  Case Is = UCase("TLockYN")
    Cancel = (pv_NewValue <> "") And ((Not moCn.RecSeek("select PMCd from Param where " + _
            "PTyp= 'YN' and PMCd = '" + pv_NewValue + "'")))
    If Cancel = True Then
      ErrMsg = "Invalid Custom Purchase Lock (Y/N)": Exit Sub
    '*************Geeta******Emr206**  02/04/04***********
    ElseIf Cancel = False Then
      If pv_NewValue = "Y" Then
        GrdTxndA.AllowAdd = False: GrdTxndA.AllowDelete = False
        GrdTxndB.AllowAdd = False: GrdTxndB.AllowDelete = False
        GrdTP.AllowAdd = False: GrdTP.AllowDelete = False
        GrdRwrk.AllowAdd = False: GrdRwrk.AllowDelete = False  '**********Bhavna
        ADC.AllowDelete = False
      Else
        GrdTxndA.AllowAdd = True: GrdTxndA.AllowDelete = True
        GrdTxndB.AllowAdd = True: GrdTxndB.AllowDelete = True
        GrdTP.AllowAdd = True: GrdTP.AllowDelete = True
        If ms_RwrkModYN Then
            GrdRwrk.AllowAdd = True: GrdRwrk.AllowDelete = True  '**********Bhavna
        End If
        ADC.AllowDelete = True
      End If
      '*************Geeta*******************
    End If
' ###########################################  Manoj  ###########################################
  ' Zubin 212 (wFgBYy added)
  Case Is = UCase("wBsBYy"), UCase("wFgBYy")
    Cancel = (pv_NewValue <> "") And (Not moCn.RecSeek("Select PMCd From Param Where PTyp= 'YY' And " + _
                                   "PMCd = '" + pv_NewValue + "'"))
    ErrMsg = "Invalid Bag Year": Exit Sub
  Case Is = UCase("wBsBChr")
    Cancel = (pv_NewValue <> "") And (Not moCn.RecSeek("Select PMCd From Param Where PTyp= 'BCHR' And " + _
                                   "PMCd = '" + pv_NewValue + "'"))
    ErrMsg = "Invalid Bag Character": Exit Sub
  ' Zubin 212
  Case Is = UCase("wFgBChr")
    '*** (Bef 2.13 CT)
    'Cancel = (pv_NewValue <> "") And (Not moCn.RecSeek("Select PMCd From Param Where PTyp= 'BCHR' And " + _
    '                               "PMCd = '" + pv_NewValue + "' and PSCd= '' and PValue= 'N'"))
    '*** (Bef 2.13 CT)
    
    '*** (Jen 2.13 CT)
    Cancel = (pv_NewValue <> "") And (Not moCn.RecSeek("Select PMCd From Param Where PTyp= 'BCHR' And " + _
                                   "PMCd = '" + pv_NewValue + "' and PSCd= '' and PValue= 'P'"))
    '*** (Jen 2.13 CT)
    ErrMsg = "Invalid Bag Character": Exit Sub
    
  ' Zubin 212
  Case Is = UCase("wBsBNo")
    '****** Sachin 3.02 21-12-07 - Current Partition
    Cancel = (pv_NewValue <> 0) And (Not moCn.RecSeek("Select BNo From Bag Where " + _
                                  "BCoCd= '" + ADC("TCoCd") + "' And " + _
                                  "BYy= '" + ADC("wBsBYy") + "' And " + _
                                  "BChr = '" + ADC("wBsBChr") + "' And " + _
                                  "BNo= " + CStr(pv_NewValue) + " And BPrtKey=(Case When '" + gs_Partition + "'='" + ctCurrPrtn + "' then '" + ctCurrPrtn + "' Else BPrtKey End)"))
    ErrMsg = "Invalid Bag Number Or Bag Closed": Exit Sub
  
  Case Is = UCase("wBrsGrp")
    Cancel = Not moCn.RecSeek("select PMCd from Param where PTyp= 'TBAGBSGR' and " + _
                         "PMCd = '" + pv_NewValue + "'")
    If Cancel = True Then ErrMsg = "Invalid Group By Option": Exit Sub
    Call BagRmSummary(pv_NewValue, ADC("WBRSVALTYP"))
  
  Case Is = UCase("wBrsValTyp")
    Cancel = pv_NewValue <> "" And Not moCn.RecSeek("select PMCd from Param where PTyp= 'WIPVAL' and " + _
                         " PMCd = '" + pv_NewValue + "'")
    If Cancel = True Then ErrMsg = "Invalid Value Type Option": Exit Sub
    Call BagRmSummary(ADC("WBRSGRP"), pv_NewValue)
  
  Case Is = UCase("wVSGrp")
    Cancel = Not moCn.RecSeek("select PMCd from Param where PTyp= 'TBAGVSGR' and " + _
                         "PMCd = '" + pv_NewValue + "'")
    If Cancel = True Then ErrMsg = "Invalid Group By Option": Exit Sub
    Call VouchSumm(pv_NewValue, ADC("wVsValTyp"))
  
  Case Is = UCase("wVSValTyp")
    Cancel = pv_NewValue <> "" And Not moCn.RecSeek("select PMCd from Param where PTyp= 'WIPVAL' and " + _
                         "PMCd = '" + pv_NewValue + "'")
    If Cancel = True Then ErrMsg = "Invalid Value Type Option": Exit Sub
    Call VouchSumm(ADC("wVSGrp"), pv_NewValue)
  
  Case Is = UCase("wIrRmCd")
    Dim wTxndARow As Single
    wTxndARow = IIF(GrdTxndA.Row > 0, GrdTxndA.Row, 0)
    If wTxndARow = 0 Then Exit Sub
    '****** Sachin 3.02 21-12-07 - Current Partition
    Cancel = (pv_NewValue <> "") And (Not moCn.RecSeek("Select OrRmCd From OrdRm " + _
                "join Bag on OrOdIdNo=BOdIdNo And OrCoCd= BCoCd and " + _
                "OrTc= BOdTc and OrYy= BOdYy and OrChr= BOdChr and OrNo= BOdNo And " + _
                "OrSr = BOdSr and OrPrtKey=BPrtKey where BCoCd= '" + ADC("TCoCd") + "' and " + _
                "BYy= '" + GrdTxndA.Value(wTxndARow, "TdBYy") + "' and " + _
                "BChr= '" + GrdTxndA.Value(wTxndARow, "TdBChr") + "' and " + _
                "BNo= " + CStr(GrdTxndA.Value(wTxndARow, "TdBNo")) + " and " + _
                "OrRmCd= '" + pv_NewValue + "' and OrPrtKey='" + ctCurrPrtn + "'"))
                '(index=Bag)
    If Cancel = True Then ErrMsg = "Invalid Raw Material Code": Exit Sub
    'mw.84- if user enter rmcd then it will get lotno from Ordrm and show that in RmLot grid
    If mb_OrLot Then
      Dim wRsOrLot As MwfLib.MDORowSet, wLotRow As Integer
      Set wRsOrLot = moCn.OpenRes("Select distinct OrLotNo From OrdRm " + _
                  "join Bag on OrOdIdNo=BOdIdNo And OrCoCd= BCoCd and " + _
                  "OrTc= BOdTc and OrYy= BOdYy and OrChr= BOdChr and OrNo= BOdNo And " + _
                  "OrSr = BOdSr and OrPrtKey=BPrtKey where BCoCd= '" + ADC("TCoCd") + "' and " + _
                  "BYy= '" + GrdTxndA.Value(wTxndARow, "TdBYy") + "' and " + _
                  "BChr= '" + GrdTxndA.Value(wTxndARow, "TdBChr") + "' and " + _
                  "BNo= " + CStr(GrdTxndA.Value(wTxndARow, "TdBNo")) + " and " + _
                  "OrRmCd= '" + pv_NewValue + "' and OrLotNo<>'' and OrPrtKey='" + ctCurrPrtn + "'")
      GrdOrLot.AllowAdd = True
      GrdOrLot.Rows = 1
      With GrdOrLot
        Do While Not (wRsOrLot.EOF Or wRsOrLot.BOF)
          .AddItem
          wLotRow = .Rows - 1
          .Value(wLotRow, "wOrLotNo") = wRsOrLot!OrLotNo
          wRsOrLot.MoveNext
        Loop
      End With
      GrdOrLot.AllowAdd = False
    End If
  Case Is = UCase("wCbToBLoc")
    Cancel = (pv_NewValue <> "") And (Not moCn.RecSeek("Select LocCd From Loc Where " + _
                              "LocCoCd= '" + ADC("TCoCd") + "' and " + _
                              "LocTyp= 'P' and LocCd<> 'PFG' and LocCd= '" + pv_NewValue + "'"))
    ErrMsg = "Invalid To Bag Location/ To Bag Loc Cannot Be 'PFG'": Exit Sub

  ' **** Manali 3.6.0 - 26/10/09 - Copy Multi Bags into Melting - wCbFrBLoc
  Case Is = UCase("wCbFrBLoc")
    Cancel = (pv_NewValue <> "") And (Not moCn.RecSeek("Select LocCd From Loc Where " + _
                              "LocCoCd= '" + ADC("TCoCd") + "' and " + _
                              "LocTyp= 'P' and LocCd<> 'PFG' and LocCd= '" + pv_NewValue + "'"))
    ErrMsg = "Invalid From Bag Location/ From Bag Loc Cannot Be 'PFG'": Exit Sub
  ' **** Manali 3.6.0 - 26/10/09 - Copy Multi Bags into Melting - wCbFrBLoc
  
'/********************************* urmi ***************************/
  Case Is = UCase("wBAckOpt")
    Cancel = (pv_NewValue <> "") And (Not moCn.RecSeek("select PMCd from Param where PTyp= 'YN' and " + _
                         "PMCd = '" + pv_NewValue + "'"))
                              
    ErrMsg = "Invalid Bag Acknowledgement Option": Exit Sub
 '/********************************* urmi ***************************/
 
  'Manoj 2.10.0
  Case Is = UCase("wLsSecMetCd")
    Dim wRmCtg As String, wRow As Integer
    wRow = GrdTxndA.Row
    wRmCtg = moCn.GetFldVal("Select RmCtg From RmMst Where RmCd ='" & ADC("wLsMainMetCd") & "'")
    '****** Sachin 2.12 - 12/12/05  - SJM - Silver Included below
    '****** Sachin 3.02 21-12-07 - Current Partition
    ' ***** Manali 3.03 -20/06/08 - Pd Mod - 'L' included
    Cancel = (pv_NewValue <> "") And _
          (Not moCn.RecSeek("Select 1 from Bag Join OrdRm On BOdIdNo=OrOdIdNo And BCoCd = OrCoCd And BodTc = OrTc And " + _
              "BOdYy = OrYy And BOdChr  = OrChr And BodNo = OrNo And BOdSr = OrSr And BPrtKey=OrPrtKey Join RmMst On OrRmCd = RmCd and OrPrtKey=RmPrtKey " + _
               "where BCoCd='" + ADC("TCoCd") + "' and BYy ='" + GrdTxndA.Value(wRow, "TdBYy") + "' and " + _
               "BChr = '" + GrdTxndA.Value(wRow, "TdBChr") + "' and BNo = '" + CStr(GrdTxndA.Value(wRow, "TdBNo")) + "' and " + _
               "RmCtg <> '" + wRmCtg + "' and RmCtg in ('G','P','S','L') And BPrtKey='" + ctCurrPrtn + "' And RmCd= '" + pv_NewValue + "' "))
    ErrMsg = "Invalid Second Metal RM Code for Loss Calculation": Exit Sub
  'Manoj 2.10.0
' Zubin 212
  Case Is = UCase("wIrLotNo")
    Cancel = pv_NewValue <> "" And Not moCn.RecSeek("Select 'x' from SYyMm where SCoCd= '" + ADC("TCoCd") + "' and SLotNo= '" + pv_NewValue + "'" + IIF(gs_Partition = ctCurrPrtn, " And SPrtKey='" + ctCurrPrtn + "'", ""))
    ErrMsg = "Invalid Lot No.": Exit Sub
' Zubin 212
  
  '***  (Jen 2.13)
  Case Is = UCase("wCrOdTc")
  'indigo.164 PValue= 'SO' changed
    Cancel = (pv_NewValue <> "") And (Not moCn.RecSeek("select PMCd from Param where PTyp= 'TC' and " + _
                                      "PMCd = '" + pv_NewValue + "' and PValue In('SO','OS') "))
    If Cancel = True Then ErrMsg = "Invalid Order Tc": Exit Sub
  Case Is = UCase("wCrOdYy")
    Cancel = (pv_NewValue <> "") And (Not moCn.RecSeek("select PMCd from Param where PTyp= 'YY' and " + _
                                      "PMCd = '" + pv_NewValue + "'"))
    If Cancel = True Then ErrMsg = "Invalid Order Year": Exit Sub
  Case Is = UCase("wCrOdChr")
    Cancel = pv_NewValue <> "" And (Not moCn.RecSeek("select vPSCd from vParam where " + _
                                    "vPCoCd= '" + ADC("TCoCd") + "' and vPTyp= 'CHR' and " + _
                                    "vPMCd = '" + ADC("wCrOdTc") + "' and vPSCd= '" + pv_NewValue + "' and " + _
                                    "(vPValue4= '' or (', '+vPValue4+', ' like '%,'+ '" + gs_UsrCd + "'+ ',%') ) "))
    If Cancel = True Then ErrMsg = "Invalid Order Character/ Access To This Caharacter Denied TO The User": Exit Sub
  Case Is = UCase("wCrOdNo")
    '****** Sachin 3.02 21-12-07 - Current Partition
    Cancel = (pv_NewValue <> 0) And (Not moCn.RecSeek("select OmNo from OrdMst where " + _
                                     "OmCoCd= '" + ADC("TCoCd") + "' and " + _
                                     "OmTc= '" + ADC("wCrOdTc") + "' and " + _
                                     "OmYy= '" + ADC("wCrOdYy") + "' and " + _
                                     "OmChr= '" + ADC("wCrOdChr") + "' and " + _
                                     "OmNo= " + CStr(pv_NewValue) + " and OmReNo> 0 And OmPrtKey='" + ctCurrPrtn + "' "))
    ErrMsg = "Invalid Order Number/ This Is Not A Return Order": Exit Sub
  Case Is = UCase("wCrBLoc")
    If pv_NewValue = ctFgLoc Then Cancel = True: ErrMsg = "Bag Cannot Be Moved From FG Location": Exit Sub
    Cancel = (pv_NewValue <> "") And (Not moCn.RecSeek("Select LocCd from Loc where LocCoCd= '" + ADC("TCoCd") + "' " + _
                                      "and LocCd = '" + pv_NewValue + "' and LocTyp= 'P'"))
    If Cancel = True Then ErrMsg = "Invalid Bag Location": Exit Sub
  '***  (Jen 2.13)
  '*** Jay 3.1.3 [InvDsg Rm]
  Case Is = UCase("wYn")
    Cancel = Not moCn.RecSeek("select PMCd from Param where PTyp= 'YN' and " + _
                         "PMCd = '" + pv_NewValue + "'")
    If Cancel = True Then ErrMsg = "Enter ''Y'' for Invoice Desing Serial RM Or Enter ''N'' for Copy From FG Bag RM": Exit Sub
    If pv_NewValue = "N" Then
      ADC("wInCoCd") = "": ADC("wInTc") = "": ADC("wInYy") = "": ADC("wInChr") = "": ADC("wInNo") = "": ADC("wInSr") = ""
      If ADC("WFGBYY") = "" Then _
        ADC("WFGBYY") = moCn.GetFldVal("Select hYy from Head where HCoCd= '" + ADC("TCoCd") + "' and HCd='" + ctSelfCmCd + "'")
      If ADC("WFGBCHR") = "" Then _
        ADC("WFGBCHR") = moCn.GetFldVal("Select hBChr from Head where HCoCd= '" + ADC("TCoCd") + "' and HCd='" + ctSelfCmCd + "'")
    Else
      ADC("WFGBYY") = "": ADC("WFGBCHR") = "": ADC("WFGBNO") = ""
      If ADC("wInCoCd") = "" Then ADC("wInCoCd") = gs_CoCd
      If ADC("WInTc") = "" Then _
        ADC("WInTc") = moCn.GetFldVal("Select PValue from Param where PTyp= 'TCTYP' and PMCd= '" + ctInTcTyp + "' and PSCd= ''")
      If ADC("WInYY") = "" Then _
        ADC("WInYY") = moCn.GetFldVal("Select hYy from Head where HCoCd= '" + ADC("TCoCd") + "' and HCd='" + ctSelfCmCd + "'")
      If ADC("WInCHR") = "" Then _
        ADC("WInCHR") = moCn.GetFldVal("Select vPValue from vParam where vPTyp= 'DEFCHR' and vPCoCd= '" + ADC("wInCoCd") + "' and vPMCd='" + ADC("wInTc") + "'")
    End If
  '*** Jay 3.1.3 [InvDsg Rm]
  
  ' ******* Manali 3.5.0 - Change in Return Memo Copy
  Case Is = UCase("wCpyOrdRmYn")
    Cancel = Not moCn.RecSeek("select PMCd from Param where PTyp= 'YN' and PMCd = '" + pv_NewValue + "'")
    If Cancel = True Then ErrMsg = "Enter 'Y' To Copy Rm Wts from Order and 'N' to Copy Actual Wts": Exit Sub

  ' ***** Manali 3.8.0
  Case Is = UCase("TMultiIr")
    If ms_TcTyp = "BS" And ADC("TTrfQty") = 0 Then
        Cancel = Not moCn.RecSeek("select PMCd from Param where PTyp= 'IR' and PMCd = '" + pv_NewValue + "'")
        If Cancel = True Then ErrMsg = "Enter 'I' for Multiple Issue or 'R' for Multiple Receipt ": Exit Sub
    Else
        Cancel = (pv_NewValue <> "")
        If Cancel = True Then ErrMsg = "Multiple I/R Should be Blank": Exit Sub
    End If

  ' ***** Manali 3.8.0 - Wrk and Prd Y/N Scope added
  Case Is = UCase("wCbWrk")
    Cancel = pv_NewValue <> "" And Not moCn.RecSeek("Select vPMCd From vParam Where " + _
                           "vPCoCd = '" + ADC("TCoCd") + "' and vPTyp= 'WORK' And " + _
                           "vPMCd = '" + pv_NewValue + "'")
    ErrMsg = "Invalid Worker": Exit Sub
    
  Case Is = UCase("wCbPrdYn")
    Cancel = Not moCn.RecSeek("select PMCd from Param where PTyp= 'YN' and PMCd = '" + pv_NewValue + "'")
    If Cancel = True Then ErrMsg = "Enter 'Y' To Copy Bags With Production Yes Only and 'N' to Copy All Bags": Exit Sub
  ' ***** Manali 3.8.0 - Wrk and Prd Y/N Scope added
  
  ' ***** Manali 3.8.0 - Loss For More Than 2 Metals
  Case Is = UCase("wLsMainMetAllYn")
    Cancel = Not moCn.RecSeek("select PMCd from Param where PTyp= 'YN' and PMCd = '" + pv_NewValue + "'")
    If Cancel = True Then ErrMsg = "Enter 'Y' To Enter Loss For All Rm Code with Main Metal Category or 'N' For Main Metal Rm Code Only ": Exit Sub
  
  Case Is = UCase("wLsSecMetAllYn")
    Cancel = Not moCn.RecSeek("select PMCd from Param where PTyp= 'YN' and PMCd = '" + pv_NewValue + "'")
    If Cancel = True Then ErrMsg = "Enter 'Y' To Enter Loss For All Rm Code with Secondary Metal Category or 'N' For Specified Rm Code Only ": Exit Sub
  ' ***** Manali 3.8.0 - Loss For More Than 2 Metals
  
  ' ***** Manali 3.10.0 - 12/04/12 - Bar Code
  Case Is = UCase("wTToBLoc")
    If pv_NewValue = ctFgLoc Then Cancel = True: ErrMsg = "Bag Cannot Be Moved To FG Location": Exit Sub
    Cancel = (pv_NewValue <> "") And (Not moCn.RecSeek("Select LocCd from Loc where " + _
              "LocCoCd= '" + ADC("TCoCd") + "' and " + _
              "LocCd = '" + pv_NewValue + "' and LocTyp= 'P'"))
    If Cancel = True Then ErrMsg = "Invalid Bag To Location": Exit Sub
  
    Case Is = UCase("wTRjCd")
      Cancel = (pv_NewValue <> "") And (Not moCn.RecSeek("Select PMCd From Param Where " + _
                                     "PTyp= 'REJCD' And PMCd = '" + pv_NewValue + "'"))
      ErrMsg = "Invalid Default Rejection Code": Exit Sub
  
    Case Is = UCase("wTBIr")
      Cancel = (pv_NewValue <> "") And (Not moCn.RecSeek("Select PMCd From Param Where PTyp= 'IR' And " + _
                                     "PMCd = '" + pv_NewValue + "'"))
      ErrMsg = "Enter Issue Or Receipt (I/R)": Exit Sub
  ' ***** Manali 3.10.0 - 12/04/12 - Bar Code
  '**** Bhavna added cpy thru bag no
  Case Is = UCase("wCpyThruBagNo")
    If ADC.Mode = xNorm Then
      'Cancel = Not moCn.RecSeek("select PMCd from Param where PTyp= 'YN' and PMCd = '" + pv_NewValue + "'")
      'ErrMsg = "Enter 'Y' to Copy thru Bag No Or 'N' to Copy thru RfId No ": Exit Sub
    End If
    'Uma KBS
 Case Is = UCase("WXLSHTNO")
     If mf_ValidLocSeq = False Then
        Exit Sub
    End If
    Cancel = pv_NewValue = "" Or pv_NewValue = 0
    If Cancel = True Then ErrMsg = "Sheet No. Should Not Be Blank Or Zero": Exit Sub
  Case Is = UCase("WXLFRROW")
    If mf_ValidLocSeq = False Then
        Exit Sub
    End If
    Cancel = pv_NewValue = "" Or pv_NewValue = 0
    If Cancel = True Then ErrMsg = "Row No. Should Not Be Blank Or Zero": Exit Sub
  Case Is = UCase("WXLTOROW")
    If mf_ValidLocSeq = False Then
        Exit Sub
    End If
    Cancel = pv_NewValue = "" Or pv_NewValue = 0
    If Cancel = True Then ErrMsg = "Row No. Should Not Be Blank Or Zero": Exit Sub
  'ec.4- validating help for LocType, Location code, sequence
  Case Is = UCase("wHsPR")
    Cancel = (pv_NewValue <> "") And (Not moCn.RecSeek("Select PMCd From Param Where " + _
              "PTyp= 'LOCTYP' and PMCd = '" + pv_NewValue + "' and PMCd in ('P','R')"))
    ErrMsg = "Invalid LocType, LocType Should Be 'P' or 'R'": Exit Sub
  Case Is = UCase("wHsLocCd")
    Cancel = (pv_NewValue <> "") And (Not moCn.RecSeek("Select LocCd From Loc Where " + _
              "LocCoCd= '" + ADC("TCoCd") + "' and LocCd='" + pv_NewValue + "' and LocTyp='" + ADC("wHsPR") + "'"))
    ErrMsg = "Invalid Location Code": Exit Sub
  Case Is = UCase("wHsSeq")
    Cancel = (pv_NewValue <> "") And (Not moCn.RecSeek("Select vPMCd From vParam Where " + _
              "vPCoCd='" + ADC("TCoCd") + "' and vPTyp='" + IIF(ADC("wHsPR") = "P", "PRDSEQ", "PRCSSEQ") + "' and vPMCd = '" + pv_NewValue + "'"))
    ErrMsg = "Invalid Sequence": Exit Sub
  'ec.4-if trfqty is >0 then hlpsort command button will be enable else disable
  Case Is = UCase("TtrfQty")
    If ms_TcTyp = "BS" And pv_NewValue > 0 Then
      CmdHlpSrt.Enabled = True
    Else
      CmdHlpSrt.Enabled = False
    End If
'ec.4- validate transfer YN value
  Case Is = UCase("wTrfYN")
    Cancel = Not moCn.RecSeek("select PMCd from Param where PTyp= 'YN' and PMCd = '" + pv_NewValue + "'")
    ErrMsg = "Enter 'Y' Or 'N' For Transfer Rm ": Exit Sub
End Select
End Sub

Private Sub CmdBrsGo_Click()

End Sub

'*** Jay 3.1.3
Private Sub CmdCRM_Click()
  If GrdTxndA.Row <= 0 Then Exit Sub
  If FraCRM.Visible = False And GrdTxndB.Rows - 1 > 0 Then DispMsg "Rm Grid Should Be Blank", etError: Exit Sub
  If ADC("wYn") = "" Then ADC("wYn") = "N"
  Call DispFra(CpyFgBagRm)
End Sub
'' Zubin 212
'Private Sub CmdCRM_Click()
'  If GrdTxndA.Row <= 0 Then Exit Sub
'  If FraCRM.Visible = False And GrdTxndB.Rows - 1 > 0 Then Exit Sub
'  If ADC("WFGBYY") = "" Then _
'    ADC("WFGBYY") = moCn.GetFldVal("Select hYy from Head where HCoCd= '" + ADC("TCoCd") + "' and HCd='" + ctSelfCmCd + "'")
'  If ADC("WFGBCHR") = "" Then _
'    ADC("WFGBCHR") = moCn.GetFldVal("Select hBChr from Head where HCoCd= '" + ADC("TCoCd") + "' and HCd='" + ctSelfCmCd + "'")
'  Call DispFra(CpyFgBagRm)
'End Sub
'*** Jay 3.1.3

' Zubin 212
'*** Jay 3.1.3
Private Sub CmdCRMGo_Click()
'*** Jay 3.1.3
  If ADC("wYn") = "N" Then
  '*** Jay 3.1.3
  ' BYy, BChr, BNo cannot be blank
    If ADC("wFgBYy") = "" Then DispMsg "Bag Year Cannot Be Blank", etWarning: Exit Sub
    If ADC("wFgBChr") = "" Then DispMsg "Bag Character Cannot Be Blank", etWarning: Exit Sub
    If ADC("wFgBNo") = 0 Then DispMsg "Bag Number Cannot Be Blank", etWarning: Exit Sub

  ' Src Bag shld be in PFG
    '****** Sachin 3.02 21-12-07 - Current Partition
    If Not moCn.RecSeek("Select 'x' from Fgd B " + _
                "Where FdCoCd = '" + ADC("TCoCd") + "' and (Select PValue from Param where PTyp= 'TC' and PMCd= B.FdTc)= 'FB' and " + _
                "not exists (select 'x' from Fgd M where M.FdCoCd= B.FdCoCd and M.FdRefYy= B.FdYy and " + IIF(gs_Partition = ctCurrPrtn, " M.FdPrtKey=B.FdPrtKey and ", "") + _
                "M.FdRefKey= B.FdKey and (Select PValue from Param where PTyp= 'TC' and PMCd= M.FdTc)= 'FWB') And " + _
                "FdBYy= '" + ADC("wFgBYy") + "' and FdBChr= '" + ADC("wFgBChr") + "' and " + _
                "FdBNo = " + CStr(ADC("wFgBNo")) + " And FdPrtKey='" + ctCurrPrtn + "' ") Then
      DispMsg "Source Bag Should Be In PFG Location", etWarning
      Exit Sub
    End If
  '*** Jay 3.1.3
  Else
    'Invdsg Sr Check cannot be blank
    If ADC("wInCoCd") = "" Or ADC("wInTc") = "" Or ADC("wInYy") = "" Or ADC("wInChr") = "" Or ADC("wInNo") = "" Or ADC("wInSr") = "" Then _
        DispMsg "Invoice Details Cannot Be Blank", etWarning: Exit Sub
  End If
  '*** Jay 3.1.3

  With GrdTxndA
    ' Tgt bag shld be empty
    '****** Sachin 3.02 21-12-07 - Current Partition
    If Not moCn.RecSeek("Select 'x' from Bag where BCoCd= '" + ADC("TCoCd") + "' and " + _
              "BYy= '" + .Value(.Row, "TdBYy") + "' and BChr= '" + .Value(.Row, "TdBChr") + "' and " + _
              "BNo= " + CStr(.Value(.Row, "TdBNo")) + " and Round(BGrWt, 3)= 0 and BPrtKey='" + ctCurrPrtn + "' ") Then
      DispMsg "Target Bag Should Empty", etWarning
      'DispMsg "Target Bag Already Has Raw Material", etWarning
      Call DispFra(CpyFgBagRm)
      Exit Sub
    End If
    
    '*** Jay 3.1.3
    If ADC("wYn") = "N" Then
    '*** Jay 3.1.3
      ' Src OrdDsg and Tgt OrdDsg shld be same
      '****** Sachin 3.02 21-12-07 - Current Partition
      If Not moCn.RecSeek("Select 'x' from Bag where BPrtKey='" + ctCurrPrtn + "' And BCoCd= '" + ADC("TCoCd") + "' and " + _
                "BYy= '" + .Value(.Row, "TdBYy") + "' and BChr= '" + .Value(.Row, "TdBChr") + "' and " + _
                "BNo= '" + CStr(.Value(.Row, "TdBNo")) + "' and BOdDmCd= (Select BOdDmCd from Bag where " + _
                "BCoCd= '" + ADC("TCoCd") + "' and BYy= '" + ADC("wFgBYy") + "' and BChr= '" + ADC("wFgBChr") + "' and " + _
                "BNo= " + CStr(ADC("wFgBNo")) + " And BPrtKey='" + ctCurrPrtn + "' )") Then
        DispMsg "Source And Target Bag Design Codes Should Be Same", etWarning
        Exit Sub
      End If

      ' Target Bag Qty <= Source Bag Qty
      If GrdTxndA.Value(GrdTxndA.Row, "TdBQty") > _
                moCn.GetFldVal("Select FdQty from Fgd where FdCoCd= '" + ADC("TCoCd") + "' and " + _
                    "FdBYy= '" + ADC("wFgBYy") + "' and FdBChr= '" + ADC("wFgBChr") + "' and " + _
                    "FdBNo= " + CStr(ADC("wFgBNo")) + " And FdPrtKey='" + ctCurrPrtn + "' ") Then
        DispMsg "Target Bag Qty Cannot Be > FG Bag Qty", etError
        Exit Sub
      End If

      ' Source Bag RmCd and Tgt Bag OrRmCd should match
      '****** Sachin 3.02 21-12-07 - Current Partition
      If moCn.RecSeek("Select 'x' from Fgd " + _
                " Join FgRm On FdIdNo=FrFdIdNo And FdCoCd= FrCoCd and FdTc= FrTc and FdYy= FrYy and FdChr= FrChr and FdNo= FrNo and FdSr= FrSr and FdPrtKey=FrPrtKey " + _
                " Join Loc LFr on LFr.LocCoCd= FrCoCd and LFr.LocCd= FrFrRmLoc " + _
                " Join Loc LTo on LTo.LocCoCd= FrCoCd and LTo.LocCd= FrToRmLoc " + _
                "where FrCoCd= '" + ADC("TCoCd") + "' and FdBYy= '" + ADC("wFgBYy") + "' and " + _
                "     FdBChr= '" + ADC("wFgBChr") + "' and FdBNo= " + CStr(ADC("wFgBNo")) + " and FdPrtKey='" + ctCurrPrtn + "' and " + _
                "     (Case when LFr.LocTyp= 'F' and FrFrRmDc='D' then 1 " + _
                "           when LTo.LocTyp= 'F' and FrToRmDc='D' then 1 " + _
                "       Else 0 End)= 1 and " + _
                "     FrRmCd Not In (Select OrRmCd from Bag join OrdRm on BOdIdNo=OrOdIdNo And BCoCd= OrCoCd and BOdTc= OrTc and " + _
                "     BOdYy= OrYy and BOdChr = OrChr And BOdNo = OrNo And BOdSr = OrSr And BPrtKey=OrPrtKey " + _
                "Where BCoCd= '" + ADC("TCoCd") + "' and BYy= '" + .Value(.Row, "TdBYy") + "' and " + _
                "   BChr= '" + .Value(.Row, "TdBChr") + "' and BNo= " + CStr(.Value(.Row, "TdBNo")) + " and BPrtKey='" + ctCurrPrtn + "')") Then
        DispMsg "RM Code Of Source Bag Is Not Same As That Of Target Bag Order RM", etError
        Exit Sub
      End If
  '*** Jay 3.1.3
      Call CopyFgRm
    Else
      ' Src OrdDsg and Tgt OrdDsg shld be same
      If Not moCn.RecSeek("Select 'x' from Bag where BPrtKey='" + ctCurrPrtn + "' And BCoCd= '" + ADC("TCoCd") + "' and " + _
                "BYy= '" + .Value(.Row, "TdBYy") + "' and BChr= '" + .Value(.Row, "TdBChr") + "' and " + _
                "BNo= '" + CStr(.Value(.Row, "TdBNo")) + "' and BOdDmCd= (Select IdDmCd from InvDsg " + _
                "where IdPrtKey='" + ctCurrPrtn + "' And IdCoCd= '" + ADC("wInCoCd") + "' and IdTc= '" + ADC("wInTc") + "' and " + _
                "IdYy= '" + ADC("wInYy") + "' and IdChr= '" + ADC("wInChr") + "' and " + _
                "IdNo= " + CStr(ADC("wInNo")) + " and IdSr= " + CStr(ADC("wInSr")) + " )") Then
        DispMsg "Source And Target Design Codes Should Be Same", etWarning
        Exit Sub
      End If
      
      ' Source Bag RmCd and Tgt Bag OrRmCd should match
      If moCn.RecSeek("Select 'x' from InvRm " + _
                "where IrPrtKey='" + ctCurrPrtn + "' And IrCoCd= '" + ADC("wInCoCd") + "' and IrTc= '" + ADC("wInTc") + "' and IrYy= '" + ADC("wInYy") + "' and " + _
                "     IrChr= '" + ADC("wInChr") + "' and IrNo= " + CStr(ADC("wInNo")) + " and IrSr= " + CStr(ADC("wInSr")) + " and " + _
                "     IrRmCd Not In (Select OrRmCd from Bag join OrdRm on BPrtKey=OrPrtKey And BCoCd= OrCoCd and BOdTc= OrTc and " + _
                "     BOdYy= OrYy and BOdChr = OrChr And BOdNo = OrNo And BOdSr = OrSr " + _
                "where BPrtKey='" + ctCurrPrtn + "' and BCoCd= '" + ADC("TCoCd") + "' and BYy= '" + .Value(.Row, "TdBYy") + "' and " + _
                "   BChr= '" + .Value(.Row, "TdBChr") + "' and BNo= " + CStr(.Value(.Row, "TdBNo")) + ")") Then
        DispMsg "RM Code Of Source Design Is Not In Target Order Design RM", etError
        Exit Sub
      End If
      Call CopyInvRm
    End If
  End With
  '*** Jay 3.1.3
  'Call CopyFgRm  '*** Jay 3.1.3 (Bef)
  Call DispFra(CpyFgBagRm)
  GrdTxndA.SetFocus
End Sub

Private Sub CmdGBD_Click()
  '*** The 'X' record (Txn record) has to be saved before the 'A' records can be entered
  '*** Display or hide the Copy Bag frame depending on the toggle state
  
  If Not ADC.SaveRec Then
    DispMsg "X Save Failed", etInfo
    Exit Sub
  End If
  
  ADC("WGBDDCTOLOC") = "W"
  Call DispFra(GenBagDet)
  
  If FraGBD.Visible = True Then
    ADC("WGBDCOCD") = ADC("TCoCd")
  End If
End Sub

Private Sub CmdGBDGo_Click()
' ###  Manoj #### Ver: 2.0.6 #### Date: 25/12/2003  ###
  If mf_DtLocked = True Or UCase$(ADC("TLockYN")) = UCase$("Y") Then
' ###  Manoj  ###
    DispMsg "Cannot Generate Bags with Details For Locked/Previous Partition Entry", etError
  Else
    If mf_ValidLocSeq = False Then
      DispMsg "The User Does Not Have Access To One Of The Locs", etError
    Else
      Call GenBagDetails
    End If
  End If
  Call CmdGBD_Click
End Sub

Private Sub CmdHlpSrt_Click()
'ec.4
  Call DispFra(HlpSrt)
  
End Sub

Private Sub CmdIE_Click()
'Uma KBS
'*** Display or hide the Import Excel frame depending on the toggle state
  'If GrdTxndA.Row <= 0 Then Exit Sub
  'cc.9999.4 - first it save X part then calling Import excel frame
  If Not ADC.SaveRec Then
    DispMsg "X Save Failed", etInfo
    Exit Sub
  End If
  Call DispFra(ImpExcel)
  
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

'************Bhavna added new lot
Private Sub CmdLN_Click()
  If GrdTxndB.Rows = 1 Then Exit Sub
  If FraLotNo.Visible = False Then
    If UCase(GrdTxndB.ColProp(GrdTxndB.Col).ColName) = UCase("TdLotNo") And GrdTxndB.TxtText <> "" Then
      ADC("wLotNo") = GrdTxndB.TxtText
    End If
    If GrdTxndB.TxtText <> GrdTxndB.Value(GrdTxndB.Row, "TdLotNo") Then GrdTxndB.Value(GrdTxndB.Row, "TdLotNo") = ""
  End If
  Call DispFra(LotNo)
End Sub

Private Sub CmdLNGo_Click()
  If ADC("wLotNo") = "" Then DispMsg "New Lot No. Cannot Be Blank", etError: ADC("wLotNo").SetFocus: Exit Sub
  If GrdTxndB.Value(GrdTxndB.Row, "TdRmCd") <> "" And Not moCn.RecSeek("select 'x' from RmMst where RmZ= 'N' and " + _
          "RmCd= '" + GrdTxndB.Value(GrdTxndB.Row, "TdRmCd") + "' and RmPrtKey='" + ctCurrPrtn + "' and RmCtg IN " + ctLotNoRmCtg) Then
    DispMsg "Lot No. Allowed For 'D', 'C','X','G','P','S','L','M' Only", etInfo
    Exit Sub
  End If
  '****** Sachin 3.02 24-12-07 - Current Partition
  If moCn.RecSeek("select 'x' from SYyMm where SCoCd= '" + ADC("TCoCd") + "' and " + _
                              "SLotNo= '" + ADC("wLotNo") + "' and SPrtKey='" + ctCurrPrtn + "' ") Then
    DispMsg "Lot No. Already Exists", etError: ADC("wLotNo").SetFocus: Exit Sub
  End If
  If GrdTxndB.Rows > 1 And ADC("wLotNo") <> "" Then GrdTxndB.Value(GrdTxndB.Row, "TdLotNo") = ADC("wLotNo")
  Call CmdLN_Click
End Sub

'************Bhavna added new lot

Private Sub CmdQW_Click()
' **** Zubin 211 **** '
  '*** Display or hide the Flute Bag Quantity/Weight frame depending on the toggle state

  If GrdTxndB.HalfValidate Then
    GrdTxndB.SetFocus
    Exit Sub
  End If

  Call DispFra(FBQtyWt)
  If FraQW.Visible = True Then
    ADC("wTdRmCd") = GrdTxndB.Value(GrdTxndB.Row, "TdRmCd")
    ADC("wTdFbRmQty") = GrdTxndB.Value(GrdTxndB.Row, "TdFbRmQty")
    ADC("wTdFbRmWt") = GrdTxndB.Value(GrdTxndB.Row, "TdFbRmWt")
  End If
' **** Zubin 211 **** '
End Sub

Private Sub CmdQWOk_Click()
' **** Zubin 211 **** '
On Error GoTo ErrHndlr
  
  '*** (Bef 22/06/05)
  'GrdTxndB.Value(GrdTxndB.Row, "TdRmQty") = moCn.GetFldVal("Select convert(int, " + ADC("wTdFbRmQty") + ")")
  'GrdTxndB.Value(GrdTxndB.Row, "TdRmWt") = moCn.GetFldVal("Select convert(decimal(16,3), " + ADC("wTdFbRmWt") + ")")
  '*** (Bef 22/06/05)
  
  '*** (Jen 22/06/05)
  If Trim(ADC("wTdFbRmQty")) <> "" Then GrdTxndB.Value(GrdTxndB.Row, "TdRmQty") = moCn.GetFldVal("Select convert(int, " + ADC("wTdFbRmQty") + ")")
  If Trim(ADC("wTdFbRmWt")) <> "" Then GrdTxndB.Value(GrdTxndB.Row, "TdRmWt") = moCn.GetFldVal("Select convert(decimal(16,3), " + ADC("wTdFbRmWt") + ")")
  '*** (Jen 22/06/05)
  
  GrdTxndB.Value(GrdTxndB.Row, "TdFbRmQty") = ADC("wTdFbRmQty")
  GrdTxndB.Value(GrdTxndB.Row, "TdFbRmWt") = ADC("wTdFbRmWt")
  
  Call DispFra(FBQtyWt)
  
Exit Sub

ErrHndlr:
  DispMsg "Invalid Value In Rm Quantity/Weight", etError
' **** Zubin 211 **** '
End Sub

Private Sub CmdRwrk_Click()

  '*** Rework entry
  '*** TdByLoc and TdByWrk should be entered in Grid 'A'
  
  If mb_RwrkFlg Or Not ms_RwrkModYN Then Exit Sub
    
  If GrdTxndA.Row <= 0 Then Exit Sub
  
  If GrdTxndA.Value(GrdTxndA.Row, "TdByLoc") = "" Or _
     GrdTxndA.Value(GrdTxndA.Row, "TdByWrk") = "" Then _
     DispMsg "By Location/Worker Should Be Entered", etWarning: GrdTxndA.SetFocus: Exit Sub
    
  Call DispFra(Rwrk)
  If FraRwrk.Visible = True Then GrdRwrk.SetFocus
End Sub

' urmila NegStk
'Private Sub TxtPwd_Change()
' ADC("TSeoPWd") = TxtPwd.Text
' ADC.SaveRec
'End Sub
' urmila NegStk

Private Sub CmdTR_Click()
  If GrdTxndA.Row <= 0 Then Exit Sub
  
  If mf_DtLocked = True Or UCase$(ADC("TLockYN")) = UCase$("Y") Then
    DispMsg "Cannot Transfer Rm For Date Locked/Locked/Previous Partition Entry", etError: Exit Sub
  Else
    If mf_ValidLocSeq = False Then
      DispMsg "The User Does Not Have Access To One Of The Locs", etError: Exit Sub
    Else
      Call TransferRm
    End If
  End If
    
End Sub

Private Sub CmdXlImport_Click()
  'Uma KBS
  CdSelPic.FileName = ""
  CdSelPic.DialogTitle = "Select Raw Material File "
  CdSelPic.Filter = "Excel Files(*.xls;*.xlsx)|*.xls;*.xlsx"
  CdSelPic.FilterIndex = 1
  CdSelPic.ShowOpen
  ADC("WXLFILE") = CdSelPic.FileName
  
End Sub
Private Sub CmdXlImpGo_Click()
'Uma KBS
    Dim wo_rsExcel As New MDORowSet, mf_ExcelMod As Boolean, i As Integer, wf_CpyFrDsgMst As Boolean
    Dim wCdDmCd As String, wCdSfx As String, wCdSz As String, wo_rsCDsg As New MDORowSet
    Dim wb_CpyOdSrFrXl As Boolean
    Dim wb_OnlyFrPDC As Boolean, wb_OnlyChk As Boolean
    Dim errorinfiles As Boolean
    errorinfiles = False
    If mf_ValidLocSeq = False Then
        DispMsg "The User Does Not Have Access To One Of The Locs", etError: Exit Sub
    Else
    If Trim(ADC("WXLFILE")) = "" Then DispMsg "Please Specify a Proper Excel File Name !!", etError: Exit Sub
    'If ADC("WXLSHTNO") = 0 Or ADC("WXLFRROW") = 0 Or ADC("WXLTOROW") = 0 Then DispMsg "Please Specify a Proper Sheet No./From Row/To Row !!", etError: Exit Sub
    If ADC("WXLSHTNO") = 0 Then DispMsg "Please Specify a Proper Sheet No.!!", etError: Exit Sub
    If Not OpenExcelFile(ADC("WXLFILE")) Then DispMsg "Improper File Format", etError: Exit Sub
    If IsExcelFileReadOnly() Then
      CloseExcel
      MsgBox ("File Is Already open. Cannot Import. Close the file and try again")
      Exit Sub
    End If
    
    If ChkSheet(ADC("WXLSHTNO")) = False Then DispMsg "This WorkSheet Number Does Not Exist in the specified Excel File", etError: CloseExcel: Exit Sub
    moCn.CreateTmpTable "#TmpXls", "Select Space(15) as qBagNo, RmCd as qRmCd, Space(16) as qRmLot," + _
                              " Cast(0.0 as Float) qRmSz,Cast(0.0 as Float) qRmSz2,Cast(0.0 as Float) qRmSz3," + _
                              " Cast(0.0 as Float) qRmStk,Cast(0.0 as Float) qRmQty, Cast(0.0 as Float) qRmWt, space(60) as qDesc, space(100) as qErr  into #TmpXls from RmMst where 1=2"

'                              " Space(8) as qBFrLoc," + _
                              " Space(1) as qBM, Space(8) as qByLoc , " + _
                              " Space(50) as qByWrk, Cast(0.0 as Float) qDWt into #TmpXls from RmMst where 1=2 "
                    
                              '" Cast(0.0 as Float) qBQty, Cast(0.0 as Float) qBWt, Space(8) as qBFrLoc," + _

'"Select 0 as qRow, 0 as qOdSr, CdCustDmCd qCustDmCd, "
'CdOurDmCd qOurDmCd, CdOurSfx qCdOurSfx, CdOurDmSz qOurDmSz, Cast(0.0 as Float) qQty,
'C 'ast(0.0 as Float) qPrdQty, Space(5) as qCol, Space(100) qOdSalRem, Space(45) qStmpInstr,
'Space(20) qOdPoNo, Space(255) as qSfxDesc, Cast(0.00 as Float) qOdSalPrc,OdDelDt as qOdDelDt,
'OdExpDelDt as qOdExpDelDt,OdLine as qOdLine  Into #TmpXls Where 1=2"
    Dim wErrMsg As String
    Dim j As Integer

    Dim Cancel As Boolean
  
    Dim RowNumA, RowNumB  As Integer
    Dim wBagDt() As String, wBagDt1 As String, wBagDt2 As String, wBagDt3 As String
    Dim wRmCd As String, wRmSz As Double, wRmSz2 As Double, wRmSz3 As Double
    Dim wRmStkRt As Double, wRQty As Long, wRWt As Double
       wBagDt1 = ""
       wBagDt2 = ""
       wBagDt3 = ""
    Dim xColRowNum As Integer
    For i = IIF(ADC("WXLFRROW") <= 1, 2, ADC("WXLFRROW")) To IIF(ADC("WXLTOROW") = 0, 9999, ADC("WXLTOROW"))
        UpdateEmrProcess "Validating row:" + CStr(i) + "/" + "Sr:" + CStr(1)
        
        wErrMsg = ""
        wBagDt() = Split(ReadCell(ADC("WXLSHTNO"), i, 1), "/")
        If UBound(wBagDt) < 2 Then DispMsg "No Record Found for Given Scope.", etInfo: CloseExcel: Exit Sub
        If wBagDt(0) <> "" And wBagDt(1) <> "" And wBagDt(2) <> "" Then
          If Not moCn.RecSeek("Select 'x' from Bag Where BYy ='" + wBagDt(0) + "' And BChr='" + wBagDt(1) + "' and Bno=" + _
                wBagDt(2) + " and BCoCd ='" + ADC("TCoCd") + "' and BPrtKey='" + ctCurrPrtn + "'") Then _
          wErrMsg = wErrMsg + IIF(wErrMsg <> "", "/", "") + " Invalid Bag Number"
        Else
        DispMsg "No Record Found for Given Scope.", etInfo: CloseExcel: Exit Sub
        End If
        
        wRmCd = Trim(ReadCell(ADC("WXLSHTNO"), i, 2))
        If Not moCn.RecSeek("Select 'x' from RmMst Where RmCd='" + wRmCd + "' And RmPrtKey='" + ctCurrPrtn + "' and RmValidYN='Y' and RmZ='N' ") Then _
          wErrMsg = wErrMsg + IIF(wErrMsg <> "", "/", "") + " Invalid RmCode"
          
        If ReadCell(ADC("WXLSHTNO"), i, 4) = "" Or Not IsNumeric(ReadCell(ADC("WXLSHTNO"), i, 4)) Then
            wErrMsg = wErrMsg + IIF(wErrMsg <> "", "/", "") + " Ln/Sv/Ptr is not Number"
        End If
        wRmSz = Val(ReadCell(ADC("WXLSHTNO"), i, 4))
        If wRmSz < 0 Or wRmSz > 999.9999 Then _
            wErrMsg = wErrMsg + IIF(wErrMsg <> "", "/", "") + " Ln/Sv/Ptr is not Correct Number"
        
        If ReadCell(ADC("WXLSHTNO"), i, 5) = "" Or Not IsNumeric(ReadCell(ADC("WXLSHTNO"), i, 5)) Then
            wErrMsg = wErrMsg + IIF(wErrMsg <> "", "/", "") + " Size2 is not Number"
        End If
        wRmSz2 = Val(ReadCell(ADC("WXLSHTNO"), i, 5))
        If wRmSz2 < 0 Or wRmSz2 > 999.9999 Then _
            wErrMsg = wErrMsg + IIF(wErrMsg <> "", "/", "") + " Size2 is not Correct Number"
            
        If ReadCell(ADC("WXLSHTNO"), i, 6) = "" Or Not IsNumeric(ReadCell(ADC("WXLSHTNO"), i, 6)) Then
            wErrMsg = wErrMsg + IIF(wErrMsg <> "", "/", "") + " Size3 is not Number"
        End If
        wRmSz3 = Val(ReadCell(ADC("WXLSHTNO"), i, 6))
        If wRmSz3 < 0 Or wRmSz3 > 999.9999 Then _
            wErrMsg = wErrMsg + IIF(wErrMsg <> "", "/", "") + " Size3 is not Correct Number"
         
        If ReadCell(ADC("WXLSHTNO"), i, 7) = "" Or Not IsNumeric(ReadCell(ADC("WXLSHTNO"), i, 7)) Then
            wErrMsg = wErrMsg + IIF(wErrMsg <> "", "/", "") + " Invalid Stock Rate no no."
        End If
        wRmStkRt = Val(ReadCell(ADC("WXLSHTNO"), i, 7))
        If wRmStkRt < 0 Then wErrMsg = wErrMsg + IIF(wErrMsg <> "", "/", "") + " Invalid Stock Rate < 0"
        Dim wRmCtg As String, wRmSCtg As String, wGrpByRt As String
        If wRmStkRt > 0 Then
          wRmCtg = moCn.GetFldVal("Select RmCtg from RmMst where RmCd= '" + wRmCd + "' and RmPrtKey='" + ctCurrPrtn + "' ")
          wRmSCtg = moCn.GetFldVal("Select RmSCtg from RmMst where RmCd= '" + wRmCd + "' and RmPrtKey='" + ctCurrPrtn + "' ")
          wGrpByRt = moCn.GetFldVal("Select PValue1 from Param where PTyp= 'RMSCTG' and " + _
                 "PMCd= '" + wRmCtg + "' and PSCd= '" + wRmSCtg + "' ")
          If wGrpByRt = "N" Then wErrMsg = wErrMsg + IIF(wErrMsg <> "", "/", "") + "Stock Rate Should Not Be Entered"
        End If
        
        If ReadCell(ADC("WXLSHTNO"), i, 8) = "" Or Not IsNumeric(ReadCell(ADC("WXLSHTNO"), i, 8)) Then
            wErrMsg = wErrMsg + IIF(wErrMsg <> "", "/", "") + " Invalid RmQty"
        End If
        wRQty = Val(ReadCell(ADC("WXLSHTNO"), i, 8))
        If wRQty < 0 Then wErrMsg = wErrMsg + IIF(wErrMsg <> "", "/", "") + " Invalid RmQty"
            
        If ReadCell(ADC("WXLSHTNO"), i, 9) = "" Or Not IsNumeric(ReadCell(ADC("WXLSHTNO"), i, 9)) Then
            wErrMsg = wErrMsg + IIF(wErrMsg <> "", "/", "") + " Invalid Rm Weight"
        End If
        wRWt = Val(ReadCell(ADC("WXLSHTNO"), i, 9))
        If wRWt < 0 Then wErrMsg = wErrMsg + IIF(wErrMsg <> "", "/", "") + " Invalid Rm Weight"
        mf_AvgRtFldChg = True
        
        If Not moCn.RecSeek("Select OrRmCd From OrdRm join Bag on " + _
                           " OrOdIdNo=BOdIdNo And OrCoCd= BCoCd and OrTc= BOdTc and OrYy= BOdYy and " + _
                           "OrChr= BOdChr and OrNo= BOdNo And OrSr = BOdSr And OrPrtKey=BPrtKey " + _
                           "where BCoCd= '" + ADC("TCoCd") + "' and " + _
                           "BYy= '" + wBagDt(0) + "' and " + _
                           "BChr= '" + wBagDt(1) + "' and " + _
                           "BNo= " + wBagDt(2) + " and " + _
                           "OrRmCd= '" + ReadCell(ADC("WXLSHTNO"), i, 2) + "' and OrPrtKey='" + ctCurrPrtn + "'") Then
                           '(index=Bag)
          wErrMsg = wErrMsg + IIF(wErrMsg <> "", "/", "") + "Invalid Raw Material Code"
        End If
    Cancel = False
    Dim ErrMsg As String
    ErrMsg = ""
    
    If (moCn.GetFldVal("Select hAvgSRtYn from Head where HCoCd= '" + ms_TCoCd + "' and " + _
                      "HCd='" + ctSelfCmCd + "'") = "Y") Or _
        (moCn.GetFldVal("Select hSizeStkRt from Head where HCoCd= '" + ms_TCoCd + "' and " + _
                      "HCd='" + ctSelfCmCd + "'") = "Y") Then
        
        '******* urmi to chk for 'Trial order Chrs' ******************
        '****** Sachin 3.02 21-12-07 - Current Partition
        If moCn.GetFldVal("Select vPValue3 from vParam where " + _
                    "vPCoCd= '" + ms_TCoCd + "' and vPTyp= 'CHR' and " + _
                    "vPmCd= (select BOdTc from Bag where BCoCd= '" + ms_TCoCd + "' and " + _
                        "BYy= '" + wBagDt(0) + "' and " + _
                        "BChr= '" + wBagDt(1) + "' and " + _
                        "BNo= " + wBagDt(2) + " and BPrtKey='" + ctCurrPrtn + "') and " + _
                    "vPSCd= (select BOdChr from Bag where BCoCd= '" + ms_TCoCd + "' and " + _
                        "BYy= '" + wBagDt(0) + "' and " + _
                        "BChr= '" + wBagDt(1) + "' and " + _
                        "BNo= " + wBagDt(2) + " and BPrtKey='" + ctCurrPrtn + "') ") = "N" Then
    '******* urmi to chk for 'Trial order Chrs' ******************
     If moCn.GetFldVal("Select PDesc225 from Param join RmMst on PTyp='RMSCTG' and PMCd=RmCtg and PSCd=RmSCtg and RmPrtKey='" + ctCurrPrtn + "' " + _
                               " where RmCd='" + ReadCell(ADC("WXLSHTNO"), i, 2) + "' ") = "Y" Then
          
          Dim wb_OrRmSz As Boolean, wb_RSz As Boolean
          '****** Sachin 3.02 21-12-07 - Current Partition
          wb_OrRmSz = moCn.RecSeek("Select OrLn1 From OrdRm join Bag on " + _
                       " OrOdIdNo=BOdIdNo And OrCoCd= BCoCd and OrTc= BOdTc and OrYy= BOdYy and " + _
                       "OrChr= BOdChr and OrNo= BOdNo And OrSr = BOdSr and OrPrtKey=BPrtKey " + _
                       "where BCoCd= '" + ADC("TCoCd") + "' and " + _
                       "BYy= '" + wBagDt(0) + "' and " + _
                       "BChr= '" + wBagDt(1) + "' and " + _
                       "BNo= " + wBagDt(2) + " and " + _
                       "OrRmCd= '" + ReadCell(ADC("WXLSHTNO"), i, 2) + "' and " + _
                       "OrLn1= " + ReadCell(ADC("WXLSHTNO"), i, 4) + " and OrPrtKey='" + ctCurrPrtn + "'")
          If wb_OrRmSz = True Then Cancel = False: GoTo ValCont
          
          '****** Sachin 3.02 21-12-07 - Current Partition
          wb_RSz = moCn.RecSeek("Select RrRngSz From OrdRm join Bag on " + _
                    "OrOdIdNo=BOdIdNo And OrCoCd= BCoCd and OrTc= BOdTc and OrYy= BOdYy and " + _
                    "OrChr= BOdChr and OrNo= BOdNo And OrSr = BOdSr and OrPrtKey=BPrtKey " + _
                    "and OrRmCd= '" + ReadCell(ADC("WXLSHTNO"), i, 2) + "' " + _
                    "join RmRt on RrTcTyp= 'Rsz' and RrCmCd= '" + ctSelfCmCd + "' and RrCtg= OrRmCtg and " + _
                    "RrSCtg= OrRmSCtg and RrPrtKey=OrPrtKey " + _
                    "where BCoCd= '" + ADC("TCoCd") + "' and " + _
                    "BYy= '" + wBagDt(0) + "' and " + _
                    "BChr= '" + wBagDt(1) + "' and " + _
                    "BNo= " + wBagDt(2) + " and " + _
                    "OrLn1 >= RrFrLn and OrLn1 <= RrToLn and OrPrtKey='" + ctCurrPrtn + "' and " + _
                    ReadCell(ADC("WXLSHTNO"), i, 4) + " >= RrFrLn and " + ReadCell(ADC("WXLSHTNO"), i, 4) + " <= RrToLn")
          If wb_RSz = True Then Cancel = False: GoTo ValCont
          ' Manoj 2.9.0. Patch
          mf_AvgRtFldChg = False
          ' Manoj 2.9.0. Patch
          Cancel = True: ErrMsg = "Invalid Raw Material Size"
        
     Else
      If moCn.GetFldVal("Select PValue3 from Param join RmMst on PTyp='RMSCTG' and PMCd=RmCtg  and PSCd=RmSCtg " + _
                               " where RmCd='" + ReadCell(ADC("WXLSHTNO"), i, 2) + "' and RmPrtKey='" + ctCurrPrtn + "' ") = "Y" Then
                               
        Dim ws_OrRmPtr As String
        Dim wb_OrRmPtr As Boolean, wb_RPtr As Boolean
          ws_OrRmPtr = IIF(moCn.GetFldVal("Select PValue2 from Param join RmMst on PTyp='RMSCTG' and PMCd=RmCtg  and PSCd=RmSCtg " + _
                               " where RmCd='" + ReadCell(ADC("WXLSHTNO"), i, 2) + "' And RmPrtKey='" + ctCurrPrtn + "' ") = "Y", "OrRmPtr", "Case when OrPrdQty=0 then 0 else OrPrdWt/OrPrdQty end")
         
          wb_OrRmPtr = moCn.RecSeek("Select OrRmPtr From OrdRm join Bag on " + _
                       "OrOdIdNo=BOdIdNo And OrCoCd= BCoCd and OrTc= BOdTc and OrYy= BOdYy and " + _
                       "OrChr= BOdChr and OrNo= BOdNo And OrSr = BOdSr And OrPrtKey=BPrtKey " + _
                       "where BCoCd= '" + ADC("TCoCd") + "' and " + _
                       "BYy= '" + wBagDt(0) + "' and " + _
                       "BChr= '" + wBagDt(1) + "' and " + _
                       "BNo= " + wBagDt(2) + " and " + _
                       "OrRmCd= '" + ReadCell(ADC("WXLSHTNO"), i, 2) + "' and " + _
                       "OrPrtKey= '" + ctCurrPrtn + "' and " + _
                       ws_OrRmPtr + "= " + ReadCell(ADC("WXLSHTNO"), i, 4))
          If wb_OrRmPtr = True Then Cancel = False: GoTo ValCont
          
          wb_RPtr = moCn.RecSeek("Select RrRngSz From OrdRm join Bag on " + _
                    "OrOdIdNo=BOdIdNo And OrCoCd= BCoCd and OrTc= BOdTc and OrYy=BOdYy and " + _
                    "OrChr= BOdChr and OrNo= BOdNo And OrSr = BOdSr and OrPrtKey=BPrtKey " + _
                    "and OrRmCd= '" + ReadCell(ADC("WXLSHTNO"), i, 2) + "' " + _
                    "join RmRt on RrTcTyp= 'Rsz' and RrCmCd= '" + ctSelfCmCd + "' and RrCtg= OrRmCtg and " + _
                    "RrSCtg= OrRmSCtg and RrPrtKey=OrPrtKey " + _
                    "where BCoCd= '" + ADC("TCoCd") + "' and " + _
                    "BYy= '" + wBagDt(0) + "' and " + _
                    "BChr= '" + wBagDt(1) + "' and " + _
                    "BNo= " + wBagDt(2) + " and " + _
                    "BPrtKey= '" + ctCurrPrtn + "' and " + _
                    ws_OrRmPtr + " >= RrFrLn and " + ws_OrRmPtr + " <= RrToLn and " + _
                    ReadCell(ADC("WXLSHTNO"), i, 4) + " >= RrFrLn and " + ReadCell(ADC("WXLSHTNO"), i, 4) + " <= RrToLn")
          If wb_RPtr = True Then Cancel = False: GoTo ValCont
          mf_AvgRtFldChg = True
          Cancel = True: ErrMsg = "Invalid Raw Material Pointer"
        End If
       End If
      End If
    End If
ValCont:
    If Cancel = True Then wErrMsg = wErrMsg + IIF(wErrMsg <> "", "/", "") + ErrMsg
    If wErrMsg <> "" Then
      errorinfiles = True
      Call SetCellVal(ADC("WXLSHTNO"), i, 11, wErrMsg)
    End If
        
    If wErrMsg <> "" Then
      moCn.Execute ("Insert Into #TmpXls Values ('" + ReadCell(ADC("WXLSHTNO"), i, 1) + "','" + ReadCell(ADC("WXLSHTNO"), i, 2) + "','" + _
                  ReadCell(ADC("WXLSHTNO"), i, 3) + "',0,0," + _
                  "0,0,0,0,''," + IIF(wErrMsg <> "", "'E'", "''") + _
                  ")")
  Else
    moCn.Execute ("Insert Into #TmpXls Values ('" + ReadCell(ADC("WXLSHTNO"), i, 1) + "','" + ReadCell(ADC("WXLSHTNO"), i, 2) + "','" + _
                  ReadCell(ADC("WXLSHTNO"), i, 3) + "'," + ReadCell(ADC("WXLSHTNO"), i, 4) + "," + ReadCell(ADC("WXLSHTNO"), i, 5) + "," + _
                  ReadCell(ADC("WXLSHTNO"), i, 6) + "," + ReadCell(ADC("WXLSHTNO"), i, 7) + ", " + ReadCell(ADC("WXLSHTNO"), i, 8) + ", " + _
                  ReadCell(ADC("WXLSHTNO"), i, 9) + ",'" + Mid(ReadCell(ADC("WXLSHTNO"), i, 10), 1, 60) + "'," + IIF(wErrMsg <> "", "'E'", "''") + _
                  ")")
  End If
  Next
    
  Set wo_rsExcel = moCn.OpenRes("Select * From #TmpXls")
  If (errorinfiles) Then
    DispMsg "Errors reported in Excel File, other rows will be imported", etError
  End If
  If wo_rsExcel.RecCount <= 0 Then DispMsg "No Record Found for Given Scope.", etInfo: Exit Sub
  If wo_rsExcel.RecCount > 0 Then
  wo_rsExcel.MoveFirst
  MWLib.BeginProcess Me, "Copying Raw Material Details ..."
  BeginEmrProcess Me, "Copying"
  xColRowNum = IIF(ADC("WXLFRROW") <= 1, 2, ADC("WXLFRROW"))
  ADC.StartCopy
  GrdTxndA.StartCopy
  GrdTxndB.StartCopy
  ADC.StartCopy
  Do While Not wo_rsExcel.EOF
      If (wo_rsExcel!qErr = "E") Then
         'Call SetCellVal(ADC("WXLSHTNO"), xColRowNum, 11, "System encountered an error while importing this row, this row is skipped")
      Else
'        ADC.StartCopy
        wBagDt() = Split(wo_rsExcel!qBagNo, "/")
        If wBagDt1 <> wBagDt(0) Or wBagDt2 <> wBagDt(1) Or wBagDt3 <> wBagDt(2) Then
           With GrdTxndA
          .AddItem
          RowNumA = .Rows - 1
          .Value(RowNumA, "TdSr") = RowNumA
          .Value(RowNumA, "TDBYY") = wBagDt(0)
          .Value(RowNumA, "TDBCHR") = wBagDt(1)
          .Value(RowNumA, "TDBNO") = wBagDt(2)
          .Value(RowNumA, "TDBQTY") = moCn.GetFldVal("Select isnull(BQty,0) from Bag where " + _
                              " BPrtKey='" + ctCurrPrtn + "' And BCoCd='" + ADC("TCoCd") + "' and " + _
                              "BYy='" + GrdTxndA.Value(RowNumA, "TdBYy") + "' and " + _
                              "BChr='" + GrdTxndA.Value(RowNumA, "TdBChr") + "' and " + _
                              "BNo= " + CStr(GrdTxndA.Value(RowNumA, "TdBNo")))
                              
          .Value(RowNumA, "TDPTQTY") = 0
          .Value(RowNumA, "TDBGRWT") = moCn.GetFldVal("Select isnull(BGrWt,0) from Bag where " + _
                              " BPrtKey='" + ctCurrPrtn + "' And BCoCd='" + ADC("TCoCd") + "' and " + _
                              "BYy='" + GrdTxndA.Value(RowNumA, "TdBYy") + "' and " + _
                              "BChr='" + GrdTxndA.Value(RowNumA, "TdBChr") + "' and " + _
                              "BNo= " + CStr(GrdTxndA.Value(RowNumA, "TdBNo")))
                              
          .Value(RowNumA, "TDFRBLOC") = moCn.GetFldVal("Select isnull(BLoc,'') from Bag where " + _
                              " BPrtKey='" + ctCurrPrtn + "' And BCoCd='" + ADC("TCoCd") + "' and " + _
                              "BYy='" + GrdTxndA.Value(RowNumA, "TdBYy") + "' and " + _
                              "BChr='" + GrdTxndA.Value(RowNumA, "TdBChr") + "' and " + _
                              "BNo= " + CStr(GrdTxndA.Value(RowNumA, "TdBNo")))
                              
          .Value(RowNumA, "TDWRK") = "ZSELF"
          .Value(RowNumA, "TDPRDYN") = "N"
          .Value(RowNumA, "TDPRTKEY") = "C"
          .Value(RowNumA, "TDKEY") = moCn.AutoGen("Txnd", "TdKey", ADC("TCoCd") + ADC("TYy"))
          
          Dim blnTxndASavedYN As Boolean
          
          blnTxndASavedYN = False
          
          On Error GoTo NextOd
          blnTxndASavedYN = .SaveRec(RowNumA)
          If Not blnTxndASavedYN Then
            Call SetCellVal(ADC("WXLSHTNO"), xColRowNum, 11, "System encountered an error while importing this row, this row is skipped")
            GoTo NextOd
          End If
        End With
        'GrdTxndA.EndCopy
      End If
      
      'GrdTxndB.StartCopy
      GrdTxndA.Row = RowNumA
      With GrdTxndB
      .AddItem
              UpdateEmrProcess "Copying row:" + CStr(RowNumA) + "/" + "Sr:" + CStr(RowNumB)

              RowNumB = .Rows - 1
              .Value(RowNumB, "TDSRNO") = RowNumB
              .Value(RowNumB, "TDFRRMLOC") = ADC("TFRRMLOC")
              .Value(RowNumB, "WTDFRRMIR") = "I"
              .Value(RowNumB, "TDFRRMDC") = "C"
              .Value(RowNumB, "TDRMCD") = wo_rsExcel!qRmCd
              .Value(RowNumB, "TDLOTNO") = wo_rsExcel!qRmLot
              .Value(RowNumB, "TDRMSZ") = wo_rsExcel!qRmSz
              .Value(RowNumB, "TDRMSZ2") = wo_rsExcel!qRmSz2
              .Value(RowNumB, "TDRMSZ3") = wo_rsExcel!qRmSz3
              .Value(RowNumB, "TDRMSTKRT") = wo_rsExcel!qRmStk
              .Value(RowNumB, "TDDESC") = wo_rsExcel!qDesc
              
              'DJ.11 - SRS requested to copy AvgStkRt from Excel into the Grid when it is non zero in Excel irrespective of
              'the AvgStkRt settings etc.
              
              If wo_rsExcel!qRmStk = 0 Then
                If moCn.GetFldVal("Select hAvgSRtYn from Head where HCoCd= '" + ms_TCoCd + "' and " + _
                      "HCd='" + ctSelfCmCd + "'") = "Y" Then
                Dim wRmLocTyp As String, wb_RmLocTyp As Boolean
                wb_RmLocTyp = moCn.RecSeek("Select LocTyp from Loc where LocCoCd= '" + ADC("TCoCd") + "' " + _
                                   "and LocCd= '" + ADC("TFrRmLoc") + "' " + _
                                   "and LocTyp in " + ctAvgLocTypSeq)
                                   
                wRmCtg = moCn.GetFldVal("Select RmCtg from RmMst where RmCd= '" + wo_rsExcel!qRmCd + "' and RmPrtKey='" + ctCurrPrtn + "' ")
                wRmSCtg = moCn.GetFldVal("Select RmSCtg from RmMst where RmCd= '" + wo_rsExcel!qRmCd + "' and RmPrtKey='" + ctCurrPrtn + "' ")
                                   
               'Calculate Avg Rt
                wGrpByRt = moCn.GetFldVal("Select PValue1 from Param where PTyp= 'RMSCTG' and " + _
                   "PMCd= '" + wRmCtg + "' and PSCd= '" + wRmSCtg + "' ")

                wRmLocTyp = Trim(moCn.GetFldVal("Select LocTyp from Loc where LocCoCd= '" + ADC("TCoCd") + "' " + _
                                        "and LocCd= '" + ADC("TFrRmLoc") + "' "))

                  If wb_RmLocTyp = True And wGrpByRt = "Y" Then
                      If wRmLocTyp <> "W" Then
                        .Value(RowNumB, "TDRMSTKRT") = GetAvgStkRt(ADC("TCoCd"), ADC("TFrRmLoc"), _
                                        wo_rsExcel!qRmCd, wo_rsExcel!qRmLot, wo_rsExcel!qRmSz, wo_rsExcel!qRmSz2, wo_rsExcel!qRmSz3)
                      Else
                        .Value(RowNumB, "TDRMSTKRT") = GetAvgStkRt(ADC("TCoCd"), ADC("TFrRmLoc"), _
                                          wo_rsExcel!qRmCd, wo_rsExcel!qRmLot, wo_rsExcel!qRmSz, wo_rsExcel!qRmSz2, wo_rsExcel!qRmSz3, _
                                          wBagDt(0), wBagDt(1), wBagDt(2))
                      End If
                  Else
                    .Value(RowNumB, "TDRMSTKRT") = 0
                  End If
                End If
              End If
              .Value(RowNumB, "TDRMQTY") = wo_rsExcel!qRmQty
              .Value(RowNumB, "TDRMWT") = wo_rsExcel!qRmWt
              .Value(RowNumB, "WTDBRKMSG") = ""
              .Value(RowNumB, "TDTORMLOC") = ADC("TTORMLOC")
              .Value(RowNumB, "TDBYLOC") = ""
              .Value(RowNumB, "TDBYWRK") = ""
              .Value(RowNumB, "TDDUSTWT") = 0

              .Value(RowNumB, "TDFBRMQTY") = ""
              .Value(RowNumB, "TDFBRMWT") = ""
              .Value(RowNumB, "TDPRTKEY") = "C"
              
          Dim blnTxndBSavedYN As Boolean
          blnTxndBSavedYN = False
          Dim GrdTxndBCancel As Boolean
          Dim GrdTxndBErrMsg As String
          GrdTxndBCancel = False
          GrdTxndBErrMsg = ""
          Call GrdTxndB_RowSave(RowNumB, GrdTxndBCancel, GrdTxndBErrMsg)
          
          If GrdTxndBCancel Then
              GrdTxndB.Del_Rec False, RowNumB
                Call SetCellVal(ADC("WXLSHTNO"), xColRowNum, 11, IIF(GrdTxndBErrMsg = "", "System encountered an error while importing this row, this row is skipped", GrdTxndBErrMsg))
                GoTo NextOd
          End If
          
          On Error GoTo NextOd
          blnTxndBSavedYN = .SaveRec(RowNumB)
                   
          If Not blnTxndBSavedYN Then
                Call SetCellVal(ADC("WXLSHTNO"), xColRowNum, 11, "System encountered an error while importing this row, this row is skipped")
                GoTo NextOd
          End If
          
          GrdTxndA.Value(RowNumA, "TDBGRWT") = moCn.GetFldVal("Select isnull(BGrWt,0) from Bag where " + _
                              " BPrtKey='" + ctCurrPrtn + "' And BCoCd='" + ADC("TCoCd") + "' and " + _
                              "BYy='" + GrdTxndA.Value(RowNumA, "TdBYy") + "' and " + _
                              "BChr='" + GrdTxndA.Value(RowNumA, "TdBChr") + "' and " + _
                              "BNo= " + CStr(GrdTxndA.Value(RowNumA, "TdBNo")))
                              
          blnTxndASavedYN = GrdTxndA.SaveRec(RowNumA)
          
          If Not blnTxndASavedYN Then
            Call SetCellVal(ADC("WXLSHTNO"), xColRowNum, 11, "System encountered an error while importing this row, this row is skipped")
            GoTo NextOd
          End If
        End With
      wBagDt1 = wBagDt(0)
      wBagDt2 = wBagDt(1)
      wBagDt3 = wBagDt(2)
                              
      'GrdTxndB.EndCopy
' SRINOW      ADC.EndCopy
'NxtBag:
      'GrdTxndA.EndCopy
NextOd:
      End If
      wo_rsExcel.MoveNext

      xColRowNum = xColRowNum + 1
      Loop
      GrdTxndB.EndCopy
      GrdTxndA.EndCopy
      ADC.EndCopy
      MWLib.EndProcess Me
      EndEmrProcess Me
      If wo_rsExcel.RecCount = (ADC("WXLToROW") - ADC("WXLFrROW") + 1) Then DispMsg "Copy completed Successfully.", etInfo
      CloseExcel
      
  End If
Call CmdIE_Click
End If
End Sub




Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
' ***** Manali 3.10.0 - 09/05/12 - Bar Code (Ctrl+N)
  If Not (KeyCode = vbKeyN And Shift = 2) Then Exit Sub
  If UCase(Me.ActiveControl.Name) = UCase("GrdTxndA") Then
    With Me.ActiveControl
      If .Col = 3 Then
        If (ms_TcTyp = "BV" Or ms_TcTyp = "DT" Or ms_TcTyp = "MLT" Or ms_TcTyp = "REJ") Then
            mb_AddRow = False
        End If
      End If
    End With
  End If
' ***** Manali 3.10.0 - 09/05/12 - Bar Code (Ctrl+N)
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
' ***** Manali 3.10.0 - 12/04/12 - Bar Code
If KeyAscii <> vbKeyReturn Then Exit Sub
If UCase(Me.ActiveControl.Name) = UCase("GrdTxndA") Then
  With Me.ActiveControl
      If .Col = 3 Then
      If (ms_TcTyp = "BV" Or ms_TcTyp = "DT" Or ms_TcTyp = "MLT" Or ms_TcTyp = "REJ") Then
          If .Rows - 1 = .Row Then
            .Col = 4
          End If
        End If
    End If
  End With
End If
' ***** Manali 3.10.0 - 12/04/12 - Bar Code
End Sub



Private Sub GrdRwrk_ColChange(ByVal RowNum As Integer, ByVal ColName As String, ByVal NewValue As Variant, Cancel As Boolean, ErrMsg As String)
'**********Bhavna
  Dim wCorrespondingPrcs As String
  wCorrespondingPrcs = moCn.GetFldVal("Select LocPrcs from Loc where LocCoCd = '" + ADC("TCoCd") + "' and LocTyp='P' and LocCd='" + GrdTxndA.Value(GrdTxndA.Row, "TdByLoc") + "'")
  Select Case ColName
  Case Is = UCase("TrwByWrk")
    Cancel = Not moCn.RecSeek("Select vPMCd From vParam Where vPCoCd = '" + ADC("TCoCd") + "' And " + _
             "vPTyp= 'WORK' and vPMCd = '" + NewValue + "'")
    If Cancel = True Then ErrMsg = "Invalid By Worker": Exit Sub
  Case Is = UCase("TrwRwrkCd")
    If wCorrespondingPrcs <> "" Then
      Cancel = Not moCn.RecSeek("Select PSCd From Param Where PTyp= 'RWRK' And " + _
             "PMCd= '" + wCorrespondingPrcs + "' and PSCd= '" + NewValue + "'")
      If Cancel = True Then ErrMsg = "Invalid Rework Code": Exit Sub
    End If
  End Select
End Sub

Private Sub GrdRwrk_ColWhen(ByVal RowNum As Integer, ByVal ColName As String, Cancel As Boolean, ErrMsg As String, ByVal Action As MwfCtl.en_fgAction)
  Dim wPrcLoc As String
  '*** Set Helps For Each Field
  '****** Bhavna
  If GrdRwrk.Value(RowNum, "TrwPrtKey") <> ctCurrPrtn Or mf_DtLocked = True Then Cancel = True: ErrMsg = "Cannot Edit Locked/Previous Partition Entry": Exit Sub
  
  If UCase$(ADC("TLockYN")) = UCase$("Y") Then Cancel = True: ErrMsg = "Cannot Add/Edit/Delete Locked Record": Exit Sub
  
  If mf_ValidLocSeq = False Then Cancel = True: ErrMsg = "The User Does Not Have Access To One Of The Locs": Exit Sub
  'uma xxxxxxxx Getting R for P - LocPrcs
  'wPrcLoc = moCn.GetFldVal("Select LocPrdTyp from Loc where LocCoCd = '" + adc("TCoCd") + "' and LocTyp= 'P' and " + _
               "LocCd = '" + GrdTxndA.Value(GrdTxndA.Row, "TdByLoc") + "'")
  wPrcLoc = moCn.GetFldVal("Select LocPrcs from Loc where LocCoCd = '" + ADC("TCoCd") + "' and LocTyp= 'P' and " + _
               "LocCd = '" + GrdTxndA.Value(GrdTxndA.Row, "TdByLoc") + "'")

  Select Case ColName
  Case Is = UCase("TrwByWrk")
    Call HlpList.vPMCd(ADC("TCoCd"), "WORK")
  Case Is = UCase("TrwRwrkCd")
    'Call HlpList.PSCd("RWRK", GrdTxndA.Value(GrdTxndA.Row, "TdByLoc"))
    Call HlpList.PSCd("RWRK", wPrcLoc)
  End Select
End Sub

Private Sub GrdRwrk_GotFocus()
'**********Bhavna
  If mf_DtLocked = True Or GrdTxndA.Value(GrdTxndA.Row, "TdPrtKey") <> ctCurrPrtn Or Not ms_RwrkModYN Then
     GrdRwrk.AllowAdd = False: GrdRwrk.AllowDelete = False
  Else
     GrdRwrk.AllowAdd = True: GrdRwrk.AllowDelete = True
  End If
End Sub

Private Sub GrdRwrk_InitKey(ByVal RowNum As Integer)
'************Bhavna
  '*** Initialise Key Fields
  With GrdRwrk
    .Store "TrwCoCd", ADC("TCoCd")
    .Store "TrwTc", ADC("TTc")
    .Store "TrwYy", ADC("TYy")
    .Store "TrwChr", ADC("TChr")
    .Store "TrwNo", ADC("TNo")
    .Store "TrwSr", IIF(GrdTxndA.Value(GrdTxndA.Row, "TdSr") > 0, GrdTxndA.Value(GrdTxndA.Row, "TdSr"), GrdTxndA.Value(1, "TdSr"))
    If .IsNew(RowNum) Then
      .Store "TrwTdIdNo", moCn.GetFldVal("Select TdIdNo From Txnd Where " + _
                                        " TdCoCd='" + ADC("TCoCd") + "' And TdTc='" + ADC("TTc") + "' And TdYy='" + ADC("TYy") + "' " + _
                                        " And TdChr='" + ADC("TChr") + "' And TdNo=" + CStr(ADC("TNo")) + " And TdSr = " + _
                                        IIF(GrdTxndA.Value(GrdTxndA.Row, "TdSr") > 0, CStr(GrdTxndA.Value(GrdTxndA.Row, "TdSr")), CStr(GrdTxndA.Value(1, "TdSr"))) + " And TdSrNo=0 and TdPrtKey='" + ctCurrPrtn + "'")
      .Store "TrwByLoc", GrdTxndA.Value(GrdTxndA.Row, "TdByLoc")
      .Store "TrwDt", ADC("TDt")
    End If
  End With

End Sub
'**********Bhavna
Private Sub GrdRwrk_ReCalc(ByVal RowNum As Integer, ByVal ColName As String, ByVal OldVal As Variant, ByVal SourceName As String, ByVal SourceColName As String)
  With GrdRwrk
    Select Case UCase(ColName)
    Case Is = UCase("wRwrkDesc")
      If (.Mode = fgminit Or .Mode = fgmnorm) Then
        If GrdTxndA.Value(GrdTxndA.Row, "TdBNo") <> 0 Then
        Dim wPrcLoc As String
  'uma xxxxxxxx Getting R for P - LocPrcs
        'wPrcLoc = moCn.GetFldVal("Select LocPrdTyp from Loc where LocCoCd = '" + adc("TCoCd") + "' and LocTyp= 'P' and " + _
               "LocCd = '" + GrdTxndA.Value(GrdTxndA.Row, "TdByLoc") + "'")
        wPrcLoc = moCn.GetFldVal("Select LocPrcs from Loc where LocCoCd = '" + ADC("TCoCd") + "' and LocTyp= 'P' and " + _
               "LocCd = '" + GrdTxndA.Value(GrdTxndA.Row, "TdByLoc") + "'")
               
        .Value(RowNum, "wRwrkDesc") = moCn.GetFldVal("Select PDesc From Param Where PTyp= 'RWRK' And " + _
             "PMCd= '" + wPrcLoc + "' and PSCd= '" + .Value(RowNum, "TrwRwrkCd") + "'")
             
        End If
      End If
      
    End Select
  End With
End Sub

Private Sub GrdRwrk_RowWhen(ByVal RowNum As Integer)
'**************Bhavna
  With GrdRwrk
    If GrdRwrk.IsNew(RowNum) Then
      .Value(RowNum, "TrwSrNo") = .MaxVal("TrwSrNo") + 1
      .Value(RowNum, "TrwByWrk") = GrdTxndA.Value(GrdTxndA.Row, "TdByWrk")
      .Value(RowNum, "TrwPrtKey") = ctCurrPrtn
      ' ******  default values from Previous rows
      If RowNum > 1 Then
        .Value(RowNum, "TrwByWrk") = .Value(RowNum - 1, "TrwByWrk")
      End If
    End If
  End With
End Sub

Private Sub GrdRwrk_SetRecSource()
'************Bhavna
  GrdRwrk.RecSource = " Select * from TxndRwrk " + _
                         " Where TrwCoCd ='" + ADC("TCoCd") + "'" + _
                         " And TrwTc ='" + ADC("TTc") + "'" + _
                         " And TrwYy = '" + ADC("TYy") + "'" + _
                         " And TrwChr ='" + ADC("TChr") + "' " + _
                         " And TrwNo =" + CStr(ADC("TNo")) + _
                         " And TrwSr =" + CStr(GrdTxndA.Value(GrdTxndA.Row, "TdSr")) + _
                         " And TrwPrtKey ='" + CStr(GrdTxndA.Value(GrdTxndA.Row, "TdPrtKey")) + "'" + _
                         " Order By TrwSrNo"

End Sub
    '***********Bhavna
Private Sub GrdRwrk_Valid()
    '*******we check that First time pass ('-' code for FTP in rework) peices do not exceed bag pcs
    Dim wBQty As Single, wRwrkQty As Single
    Dim wTxndARow As Single
  
  'If moCn.GetFldVal("Select  hRwrkModYn From Head where HCoCd= '" + ADC("TCoCd") + "' and HCd= '" + ctSelfCmCd + "'") = "Y" Then
  If ms_RwrkModYN Then
    wTxndARow = IIF(GrdTxndA.Row, GrdTxndA.Row, 0)
    
    wBQty = moCn.GetFldVal("Select Sum(BQty) from Bag where " + _
                        "BCoCd='" + ADC("TCoCd") + "' and " + _
                        "BYy='" + GrdTxndA.Value(wTxndARow, "TdBYy") + "' and " + _
                        "BChr='" + GrdTxndA.Value(wTxndARow, "TdBChr") + "' and " + _
                        "BNo= " + CStr(GrdTxndA.Value(wTxndARow, "TdBNo")) + " and BPrtKey='" + ctCurrPrtn + "' and " + _
                        "BLstYy='" + ADC("TYy") + "' and " + _
                        "BLstKey=" + CStr(GrdTxndA.Value(wTxndARow, "TdKey")))
    'uma *** TdPDQty for PD bag
    'If GrdTxndA.Value(wTxndARow, "TdBChr") = "PD2" Then wBQty = GrdTxndA.Value(wTxndARow, "TdPDQty")
    ms_BTyp = moCn.GetFldVal("Select PValue from Param where PTyp ='BCHR' and PMCd='" + GrdTxndA.Value(wTxndARow, "TdBChr") + "' and PSCd=''")
    If ms_BTyp = ctBTyp Then wBQty = GrdTxndA.Value(wTxndARow, "TdPDQty")
    
    If wBQty <= 0 Then Exit Sub

    wRwrkQty = moCn.GetFldVal("Select sum(TrwQty) from TxndRwrk join Txnd on TdCoCd= TrwCoCd and " + _
                  "TdTc= TrwTc and TdYy= TrwYy and TdChr= TrwChr and TdNo= TrwNo and TdSr= TrwSr and TdPrtKey=TrwPrtKey and TdSrNo= 0 " + _
                  " and TdByLoc= '" + GrdTxndA.Value(wTxndARow, "TdByLoc") + "' " + _
                  " where TdCoCd= '" + ADC("TCoCd") + "' and " + _
                  "TdBYy= '" + GrdTxndA.Value(wTxndARow, "TdBYy") + "' and " + _
                  "TdBChr= '" + GrdTxndA.Value(wTxndARow, "TdBChr") + "' and " + _
                  "TdBNo= " + CStr(GrdTxndA.Value(wTxndARow, "TdBNo")) + " and TdPrtKey='" + ctCurrPrtn + "' and TrwRwrkCd='-'")
                  
        If wRwrkQty > wBQty Then
          DispMsg "First time Pass Quantity " + Str(wRwrkQty) + " Cannot Exceed Bag Quantity. " + Str(wBQty), etWarning
          mb_RwrkFlg = True
          GrdRwrk.SetFocus
          Exit Sub
        Else
          mb_RwrkFlg = False
        End If
        
   End If
End Sub

'**************Bhavna
Private Sub GrdRwrk_Validate(Cancel As Boolean)
  Cancel = GrdRwrk.Validate
End Sub

Private Sub GrdTP_GotFocus()

'****** Sachin 3.02 24-12-07
  If mf_DtLocked = True Or GrdTxndA.Value(GrdTxndA.Row, "TdPrtKey") <> ctCurrPrtn Then
     GrdTP.AllowAdd = False: GrdTP.AllowDelete = False
  Else
     GrdTP.AllowAdd = True: GrdTP.AllowDelete = True
  End If

End Sub

Private Sub GrdTP_Valid()
' **** Zubin 211 **** '
  Dim wTxndARow As Single, wWHQty As Single
  Dim wRsRm As MwfLib.MDORowSet
  Dim wPrdTyp As String, wPrcLoc As String

  mb_PrdPtsFlg = False
  
  
  If moCn.GetFldVal("Select HOrdPrdPts From Head where HCoCd= '" + ADC("TCoCd") + "' and HCd= '" + ctSelfCmCd + "'") = "Y" Then
    wTxndARow = IIF(GrdTxndA.Row, GrdTxndA.Row, 0)
'Manoj 2.11.0 Corrections
    '****** Sachin 3.02 21-12-07 - Current Partition
    If Not moCn.RecSeek("Select BYy from Bag where " + _
                      "BCoCd='" + ADC("TCoCd") + "' and " + _
                      "BYy='" + GrdTxndA.Value(wTxndARow, "TdBYy") + "' and " + _
                      "BChr='" + GrdTxndA.Value(wTxndARow, "TdBChr") + "' and " + _
                      "BNo= " + CStr(GrdTxndA.Value(wTxndARow, "TdBNo")) + " and BPrtKey='" + ctCurrPrtn + "' and " + _
                      "BLstYy='" + ADC("TYy") + "' and " + _
                      "BLstKey=" + CStr(GrdTxndA.Value(wTxndARow, "TdKey"))) Then Exit Sub

    ' keep this check if we want to see if PrdPts have been given for RmCtg-SCtg that do not exist in the bag
'    If moCn.RecSeek("Select 'x' From TxndPrd join Txnd Td1 on TdCoCd= TpCoCd and TdTc= TpTc and TdYy= TpYy and " + _
'                  "TdChr= TpChr and TdNo= TpNo and TdSr= TpSr and TdSrNo= 0  And " + _
'                  "(Select LocPrdTyp From Loc Where LocCoCd = TdCoCd And LocCd = TdByLoc) in ('GS', 'WS', 'GP', 'WP') " + _
'                  "where TdCoCd= '" + adc("TCoCd") + "' and TdBYy= '" + GrdTxndA.Value(wTxndARow, "TdBYy") + "' and " + _
'                  "TdBChr= '" + GrdTxndA.Value(wTxndARow, "TdBChr") + "' and " + _
'                  "TdBNo= " + CStr(GrdTxndA.Value(wTxndARow, "TdBNo")) + " and " + _
'                  "Not Exists ( Select 'x' From Txnd Td2 join Bag on BCoCd= TdCoCd and BYy= TdBYy and " + _
'                  "BChr= TdBChr and BNo= TdBNo " + _
'                  "join RmMst on RmCd= TdRmCd where Td2.TdCoCd=Td1.TdCoCd and Td2.TdBYy= Td1.TdBYy and " + _
'                  "Td2.TdBChr= Td1.TdBChr and Td2.TdBNo= Td1.TdBNo and TpRmCtg= RmCtg and TpRmSCtg= RmSCtg) ") Then
'        DispMsg "Production Prd Quantity Cannot Exceed Diamond/Color Stone Quantity In Bag", etWarning
'        mb_PrdPtsFlg = True
'        GrdTP.SetFocus
'        Exit Sub
'    Else
'      mb_PrdPtsFlg = False
'    End If
'
'    wPrdTyp = moCn.GetFldVal("Select LocPrdTyp from Loc where " + _
'              "LocCoCd= '" + adc("TCoCd") + "' and " + _
'              "LocCd= '" + GrdTxndA.Value(wTxndARow, "TdByLoc") + "'")
'    'wStr = IIF(UCase(Trim(wPrdTyp)) = "WS" Or UCase(Trim(wPrdTyp)) = "WP", " sum(OrWsQty) as qRmQty ", " sum(OrHsQty) as qRmQty ")
'    Set wRsRm = moCn.OpenRes("Select RmCtg, RmSCtg, sum((case (case when TdFrRmLoc='W' " + _
'                "then TdFrRmDc when TdToRmLoc='W' then TdToRmDc end) when 'C' then -1 " + _
'                "when 'D' then +1 else 0 end)* TdRmQty) as qRmQty, max(BOdTc) as BOdTc, max(BOdYy) as BOdYy, max(BOdChr) as BOdChr, " + _
'                "max(BOdNo) as BOdNo, max(BOdSr) as BOdSr From Txnd join Bag on BCoCd= TdCoCd and " + _
'                "BYy= TdBYy and BChr= TdBChr and BNo= TdBNo join RmMst on RmCd= TdRmCd " + _
'                "where TdCoCd='" + ADC("TCoCd") + "' and " + _
'                "TdBYy= '" + GrdTxndA.Value(wTxndARow, "TdBYy") + "' and " + _
'                "TdBChr= '" + GrdTxndA.Value(wTxndARow, "TdBChr") + "' and " + _
'                "TdBNo= " + CStr(GrdTxndA.Value(wTxndARow, "TdBNo")) + " " + _
'                "group By RmCtg, RmSCtg order By RmCtg, RmSCtg ")
'    Do While Not (wRsRm.EOF Or wRsRm.BOF)
'      'wBagRmQty = wRsRm!qRmQty
'      If wRsRm!qRmQty > 0 Then
'' ###################################  Manoj #### Ver: 2.0.6  ###################################
'' Date: 05/12/2003
'' Change: To Check Already given Prdpts of Same Type i.e. Setting or Picking
'        '     Filter for PrdTyp is Added
'        wWHQty = moCn.GetFldVal("Select sum(TpRmQty) from TxndPrd join Txnd on TdCoCd= TpCoCd and " + _
'                "TdTc= TpTc and TdYy= TpYy and TdChr= TpChr and TdNo= TpNo and TdSr= TpSr and TdSrNo= 0 " + _
'                " And (Select LocPrdTyp From Loc Where LocCoCd = TdCoCd And LocCd = TdByLoc) in (" + _
'                IIF(wPrdTyp = "GS" Or wPrdTyp = "WS", "'GS', 'WS'", "'GP', 'WP'") + _
'                ") where TdCoCd= '" + ADC("TCoCd") + "' and " + _
'                "TdBYy= '" + GrdTxndA.Value(wTxndARow, "TdBYy") + "' and " + _
'                "TdBChr= '" + GrdTxndA.Value(wTxndARow, "TdBChr") + "' and " + _
'                "TdBNo= " + CStr(GrdTxndA.Value(wTxndARow, "TdBNo")) + " and " + _
'                "TpRmCtg= '" + wRsRm!RmCtg + "' and TpRmSCtg= '" + wRsRm!RmSCtg + "' ")
'        'wBagRmQty = wRsRm!qRmQty - wWHQty
'        If wRsRm!qRmQty < wWHQty Then
'          DispMsg "Production Prd Quantity Cannot Exceed Diamond/Color Stone Quantity In Bag", etWarning
'          mb_PrdPtsFlg = True
'          GrdTP.SetFocus
'          Exit Sub
'        Else
'          mb_PrdPtsFlg = False
'        End If
'      End If
'      wRsRm.MoveNext
'    Loop
'uma xxxxxxxx Getting R for P - LocPrcs
      wPrcLoc = moCn.GetFldVal("Select LocPrcs from Loc where LocCoCd = '" + ADC("TCoCd") + "' and LocTyp= 'P' and " + _
               "LocCd = '" + GrdTxndA.Value(GrdTxndA.Row, "TdByLoc") + "'")

          Dim wPrcLocQuery As String
          wPrcLocQuery = "(Select LocPrcs from Loc where LocCoCd = '" + ADC("TCoCd") + "' and LocTyp= 'P' and " + _
               "LocCd = Txnd.TdByLoc)"
    
    wPrdTyp = moCn.GetFldVal("Select LocPrdTyp from Loc where " + _
              "LocCoCd= '" + ADC("TCoCd") + "' and " + _
              "LocCd= '" + wPrcLoc + "'")

    ' ****** Sachin 3.02 (14/09/07) Id fields added
    '****** Sachin 3.02 21-12-07 - Current Partition
    Set wRsRm = moCn.OpenRes("Select Tp1.TpRmCtg, Tp1.TpRmSCtg, sum(Tp.TpRmQty) qRmQty from TxndPrd Tp join Txnd on TdIdNo=TpTdIdNo And TdCoCd= TpCoCd and " + _
                "TdTc= Tp.TpTc and TdYy= Tp.TpYy and TdChr= Tp.TpChr and TdNo= Tp.TpNo and TdSr= Tp.TpSr and TdSrNo= 0 and TdPrtKey=Tp.TpPrtKey " + _
                " And (Select LocPrdTyp From Loc Where LocCoCd = TdCoCd And LocCd = " + wPrcLocQuery + ") in (" + _
                IIF(wPrdTyp = "GS" Or wPrdTyp = "WS", "'GS', 'WS'", "'GP', 'WP'") + _
                ") Join (Select TpRmCtg, TpRmSCtg From TxndPrd Where " + _
                "TpCoCd= '" + ADC("TCoCd") + "' and TpTc = '" + ADC("TTc") + "' and " + _
                "TpYy = '" + ADC("TYy") + "' and TpChr = '" + ADC("TChr") + "' and " + _
                "TpNo = '" + CStr(ADC("TNo")) + "' and TpSr = '" + CStr(GrdTxndA.Value(wTxndARow, "TdSr")) + "'" + _
                " and TpPrtKey='" + ctCurrPrtn + "'" + _
                "Group By TpRmCtg, TpRmSCtg) Tp1 On " + _
                "Tp.TpRmCtg = Tp1.TpRmCtg And Tp.TpRmSCtg =Tp1.TpRmSCtg " + _
                "where TdCoCd= '" + ADC("TCoCd") + "' and " + _
                "TdBYy= '" + GrdTxndA.Value(wTxndARow, "TdBYy") + "' and " + _
                "TdBChr= '" + GrdTxndA.Value(wTxndARow, "TdBChr") + "' and " + _
                "TdBNo= " + CStr(GrdTxndA.Value(wTxndARow, "TdBNo")) + " and TdPrtKey='" + ctCurrPrtn + "' " + _
                "Group By Tp1.TpRmCtg, Tp1.TpRmSCtg order By Tp1.TpRmCtg, Tp1.TpRmSCtg ")
    Do While Not (wRsRm.EOF Or wRsRm.BOF)
      'wBagRmQty = wRsRm!qRmQty
      If wRsRm!qRmQty > 0 Then
' ###################################  Manoj #### Ver: 2.0.6  ###################################
' Date: 05/12/2003
' Change: To Check Already given Prdpts of Same Type i.e. Setting or Picking
        '     Filter for PrdTyp is Added
        '****** Sachin 3.02 21-12-07 - Current Partition
        wWHQty = moCn.GetFldVal("Select sum((case (case when TdFrRmLoc='W' " + _
                "then TdFrRmDc when TdToRmLoc='W' then TdToRmDc end) when 'C' then -1 " + _
                "when 'D' then +1 else 0 end)* TdRmQty) as qRmQty " + _
                "From Txnd join RmMst on RmCd= TdRmCd And RmPrtKey=TdPrtKey " + _
                " where TdCoCd='" + ADC("TCoCd") + "' and " + _
                "TdBYy= '" + GrdTxndA.Value(wTxndARow, "TdBYy") + "' and " + _
                "TdBChr= '" + GrdTxndA.Value(wTxndARow, "TdBChr") + "' and " + _
                "TdBNo= " + CStr(GrdTxndA.Value(wTxndARow, "TdBNo")) + " And TdPrtKey='" + ctCurrPrtn + "' And " + _
                "RmCtg = '" + wRsRm!TpRmCtg + "' And RmSCtg = '" + wRsRm!TpRmSCtg + "' ")
        'wBagRmQty = wRsRm!qRmQty - wWHQty
        If wRsRm!qRmQty > wWHQty Then
          DispMsg "Production Prd Quantity Cannot Exceed Diamond/Color Stone Quantity In Bag", etWarning
          mb_PrdPtsFlg = True
          GrdTP.SetFocus
          Exit Sub
        Else
          mb_PrdPtsFlg = False
        End If
      End If
      wRsRm.MoveNext
    Loop
'Manoj 2.11.0 Corrections
  End If
' **** Zubin 211 **** '
 wTxndARow = IIF(GrdTxndA.Row, GrdTxndA.Row, 0)
'uma xxxxxxxx Getting R for P - LocPrcs
 wPrcLoc = moCn.GetFldVal("Select LocPrcs from Loc where LocCoCd = '" + ADC("TCoCd") + "' and LocTyp= 'P' and " + _
               "LocCd = '" + GrdTxndA.Value(GrdTxndA.Row, "TdByLoc") + "'")

 wPrdTyp = moCn.GetFldVal("Select LocPrdTyp from Loc where " + _
              "LocCoCd= '" + ADC("TCoCd") + "' and " + _
              "LocCd= '" + wPrcLoc + "'")
              
    If moCn.GetFldVal("Select hPrdPtsOnBQtyYN From Head where HCoCd= '" + ADC("TCoCd") + "' and HCd= '" + ctSelfCmCd + "'") = "Y" And UCase(wPrdTyp) = "GN" Then
   
    Dim wBQty As Single, wTpBQty As Single
    
    wBQty = moCn.GetFldVal("Select Sum(case when BRepYN='Y' then BRepQty else BQty end) from Bag where " + _
                        "BCoCd='" + ADC("TCoCd") + "' and " + _
                        "BYy='" + GrdTxndA.Value(wTxndARow, "TdBYy") + "' and " + _
                        "BChr='" + GrdTxndA.Value(wTxndARow, "TdBChr") + "' and " + _
                        "BNo= " + CStr(GrdTxndA.Value(wTxndARow, "TdBNo")) + " and BPrtKey='" + ctCurrPrtn + "' and " + _
                        "BLstYy='" + ADC("TYy") + "' and " + _
                        "BLstKey=" + CStr(GrdTxndA.Value(wTxndARow, "TdKey")))

    'uma *** BQty to TdPDQty
    'If GrdTxndA.Value(GrdTxndA.Row, "TdBChr") = "PD2" Then wBQty = GrdTxndA.Value(GrdTxndA.Row, "TdPDQty")
    ms_BTyp = moCn.GetFldVal("Select PValue from Param where PTyp ='BCHR' and PMCd='" + GrdTxndA.Value(GrdTxndA.Row, "TdBChr") + "' and PSCd=''")
    If ms_BTyp = ctBTyp Then wBQty = GrdTxndA.Value(GrdTxndA.Row, "TdPDQty")
    
    If wBQty <= 0 Then Exit Sub

    Dim wTpSubPrc As String
  
    
    wTpSubPrc = moCn.GetFldVal("Select isnull(TpSubPrc,'') from TxndPrd join Txnd on TdCoCd= TpCoCd and " + _
                  " TdTc= TpTc and TdYy= TpYy and TdChr= TpChr and TdNo= TpNo and TdSr= TpSr and TdPrtKey=TpPrtKey and TdSrNo= 0 " + _
                  " join Loc on LocCoCd = TdCoCd  " + _
                  " and LocCd = TdByLoc " + _
                  " where TdCoCd= '" + ADC("TCoCd") + "' and " + _
                  " TdBYy= '" + GrdTxndA.Value(wTxndARow, "TdBYy") + "' and " + _
                  " TdBChr= '" + GrdTxndA.Value(wTxndARow, "TdBChr") + "' and " + _
                  " Loc.LocPrcs = '" + wPrcLoc + "' and " + _
                  " TdBNo= " + CStr(GrdTxndA.Value(wTxndARow, "TdBNo")) + " and TdPrtKey='" + ctCurrPrtn + "' group by TpSubPrc having sum(TpBQty) > " + CStr(wBQty))
                  
        If wTpSubPrc <> "" Then
          'If wTpBQty > wBQty Then
          DispMsg "Production Prd Quantity Cannot Exceed PD/Bag Quantity.", etWarning
          mb_PrdPtsFlg = True
          GrdTP.SetFocus
          Exit Sub
        Else
          mb_PrdPtsFlg = False
        End If
    End If
    
End Sub

Private Sub GrdTxndA_GotFocus()
' New Code
  '*** Enable all the command buttons associated with the Grid 'A'
  Dim wTRChr As String
  If CmdTP.Visible = True Then CmdTP.Enabled = True
  If CmdBD.Visible = True Then CmdBD.Enabled = True
  If CmdIR.Visible = True Then CmdIR.Enabled = True
  If CmdBS.Visible = True Then CmdBS.Enabled = True
  If CmdER.Visible = True Then CmdER.Enabled = True
  If CmdIE.Visible = True Then CmdIE.Enabled = True ' Uma KBS
  If CmdTot.Visible = True Then CmdTot.Enabled = True
  If CmdLs.Visible = True Then CmdLs.Enabled = True
  If CmdCert.Visible = True Then CmdCert.Enabled = True    '*****Bhavna
  
  ' **** Zubin 211 **** '
  If CmdQW.Visible = True Then CmdQW.Enabled = False
  ' **** Zubin 211 **** '
  If CmdLN.Visible = True Then CmdLN.Enabled = False   '********Bhavna
  
  ' Zubin 212
  If CmdCRM.Visible = True Then CmdCRM.Enabled = True
  ' Zubin 212
  
 ' *Bhavna
 If ms_RwrkModYN Then
  If CmdRwrk.Visible = True Then CmdRwrk.Enabled = True
 End If
  ' *******Bhavna
   
  '???
  'If ADC("TtrfQty") > 0 Then
  '  If CmdTR.Visible = True Then CmdTR.Enabled = True
  '  If CmdCR.Visible = False Then CmdCR.Enabled = False
  'Else
  '  If CmdTR.Visible = False Then CmdTR.Enabled = False
  '  If CmdCR.Visible = True Then CmdCR.Enabled = True
  'End If
  '???
  
  If ADC("TtrfQty") > 0 Then
    If CmdTR.Visible = True Then CmdTR.Enabled = True
    CmdCR.Enabled = False
  Else
    CmdTR.Enabled = False
    If CmdCR.Visible = True Then CmdCR.Enabled = True
  End If

  'uma *** Repair Tc
  If CmdVS.Visible = True Then CmdVS.Enabled = True
  wTRChr = moCn.GetFldVal("Select vPValue3 from vParam where vPCoCd='" + ADC("TCoCd") + "' and vPTyp='CHR' and vPMCD='" + ADC("TTc") + "' and vPSCd='" + ADC("TChr") + "'")
If wTRChr = "Y" Then
'  CmdTP.Enabled = False
'  CmdBD.Enabled = False
  CmdIR.Enabled = False
 ' CmdIE.Enabled = False 'Uma KBS
'  CmdBS.Enabled = False
  CmdGBD.Enabled = False
 ' CmdRwrk.Enabled = False
  CmdLs.Enabled = False
'  CmdCB.Enabled = False
  CmdCRM.Enabled = False
'  CmdTot.Enabled = True
End If

   'If CmdCB.Visible = True Then CmdCB.Enabled = True
End Sub
Private Sub GrdTxndA_OnDelete(ByVal RowNum As Integer, Cancel As Boolean, ErrMsg As String)
  'If wb_Valid = False Then DispMsg "From Location Not Accessible", etError
  If ADC("Ttc") = "BS" Then
    If ADC("TMultiIR") = "R" And GrdTxndA.Value(RowNum, "TdBDC") = "C" Then
    ml_bsnoofcreditbags = ml_bsnoofcreditbags - 1
    End If
  If ADC("TMultiIR") = "I" And GrdTxndA.Value(RowNum, "TdBDC") = "D" Then
    ml_bsnoofdebitbags = ml_bsnoofdebitbags - 1
  End If
End If
End Sub
'Private Sub GrdTxndA_OnDelete(ByVal RowNum As Integer, Cancel As Boolean, errmsg As String)
'End Sub

Private Sub GrdTxndA_When()
CmdLN.Enabled = False

' Old Code
  '*** Enable all the command buttons associated with the Grid 'A'
  
'  If CmdTP.Visible = True Then CmdTP.Enabled = True
'  If CmdBD.Visible = True Then CmdBD.Enabled = True
'  If CmdIR.Visible = True Then CmdIR.Enabled = True
'  If CmdBS.Visible = True Then CmdBS.Enabled = True
'  If CmdCR.Visible = True Then CmdCR.Enabled = True
'  If CmdER.Visible = True Then CmdER.Enabled = True
'  If CmdVS.Visible = True Then CmdVS.Enabled = True
'  If CmdTot.Visible = True Then CmdTot.Enabled = True
'  If CmdLs.Visible = True Then CmdLs.Enabled = True
End Sub

Private Sub GrdTxndA_RowWhen(ByVal RowNum As Integer)
  '*** If GrdTxndA.IsNew Then TdSr = Max + 1
  '*** Autogenerate the TdKey with TYy as the Pre Key
  
  With GrdTxndA
    '****** Sachin 3.02 20-12-07 - Setting the default value in the 'A' grid to Current Partition
    If .IsNew(RowNum) Then .Value(RowNum, "TdPrtKey") = ctCurrPrtn
           
    'If adc("TTc") = "BS" And adc("TtrfQty") > 0 Then
    If ADC("TTc") = "BS" Then
      If ADC("TtrfQty") > 0 Then
        'If .Rows - 1 >= 2 Then
        '  .AllowAdd = False
        'Else
          .AllowAdd = True
        'End If
      ' ***** Manali 3.8.0
      ElseIf ADC("TMultiIr") = "I" Then
        If RowNum > 1 Then
          If .Value(.Rows - 1, "wTdBIr") = "R" Then
            .AllowAdd = False
          Else
            .AllowAdd = True
          End If
        ElseIf .Rows - 1 = 1 Then
          .AllowAdd = True
        End If
      End If
      ' ***** Manali 3.8.0
    End If
  
    ' **** Manali 3.8.0
'    If adc("TTc") = "BS" And adc("TMultiIr") = "I" And RowNum = 1 And .IsNew(RowNum) Then
'      mi_TdKey = moCn.AutoGen("Txnd", "TdKey", adc("TCoCd") + adc("TYy"))
'    End If
  
    If ADC("TTc") = "BS" And ADC("TMultiIr") = "I" Then
      'moCn.GetFldVal ("Select * fro bag")
      If RowNum = 1 And .IsNew(RowNum) Then
        mi_TdKey = moCn.AutoGen("Txnd", "TdKey", ADC("TCoCd") + ADC("TYy"))
      Else
        mi_TdKey = IIF(mi_TdKey <> 0, mi_TdKey, _
                    moCn.GetFldVal("Select Top 1 TdRefKey From Txnd Where TdCoCd='" + ADC("TCoCd") + "' and TdTc='" + ADC("TTc") + "' and TdYy='" + ADC("TYy") + "' " + _
                                  " and TdChr='" + ADC("TChr") + "' and TdNo=" + CStr(ADC("TNo")) + " /*and TdSr=" + CStr(.Value(1, "TdSr")) + "*/ and TdSrNo=0 and TdBDc='C' Order By TdKey "))
      End If
      'Debug.Print mi_TdKey
    End If
  
    If ADC("TTc") = "BS" Then
      If ADC("wTrfYN") = "Y" Then
        If RowNum Mod 2 = 0 Then
          .Value(RowNum, "wTdBIr") = "R"
        Else
          .Value(RowNum, "wTdBIr") = "I"
        End If
      End If
      ' ***** Manali 3.8.0
    End If
    
    mi_Count = 1
    If UCase(ADC.MenuCd) <> UCase("BA") Then .AllowDelete = Not mf_DtLocked
    If .Value(RowNum, "TdSr") = 0 Then .Value(RowNum, "TdSr") = .MaxVal("TdSr") + 1
    .Value(RowNum, "TdSrNo") = 0
    If .Value(RowNum, "TdKey") = 0 Then _
       .Value(RowNum, "TdKey") = moCn.AutoGen("Txnd", "TdKey", ADC("TCoCd") + ADC("TYy"))
  
    '4.1.4
    '****************  urmi 19/07/03  ***************************************'
'    Dim wBFrLoc As String
'    If (UCase(adc.MenuCd) = "BV" Or UCase(adc.MenuCd) = "DT") And .IsNew(RowNum) <> True And (Not mf_DtLocked) Then
'      If gs_ValidPrdLocSeq <> "" Then
'        wBFrLoc = .Value(RowNum, "TdFrBLoc")
'        wb_Valid = IIF((InStr(1, "," + UCase(gs_ValidPrdLocSeq) + ",", "," + UCase(wBFrLoc) + ",") <= 0), False, True)
'      Else
'        wb_Valid = True
'      End If
'      .AllowDelete = wb_Valid
'    End If
    '**************************************************************************'
    
'    Dim wBFrLoc As String
'    If .IsNew(RowNum) <> True And Not mf_DtLocked And UCase(adc.MenuCd) <> "BA" Then
'      If gs_ValidPrdLocSeq <> "" Then
'        wBFrLoc = moCn.GetFldVal("Select BLoc from Bag where " + _
'                            "BCoCd='" + adc("TCoCd") + "' and " + _
'                            "BYy='" + .Value(RowNum, "TdBYy") + "' and " + _
'                            "BChr='" + .Value(RowNum, "TdBChr") + "' and " + _
'                            "BNo= " + CStr(.Value(RowNum, "TdBNo")) + " And BPrtKey='" + ctCurrPrtn + "'")
'        wb_Valid = IIF((InStr(1, "," + UCase(gs_ValidPrdLocSeq) + ",", "," + UCase(wBFrLoc) + ",") <= 0), False, True)
'      Else
'        wb_Valid = True
'      End If
'      .AllowDelete = wb_Valid
'    End If
    
    'Manoj 2.10.0
    '****** Sachin 2.14.0 Next [03-02-2007] Added DtLock and Voucher Lock Condition ******
    If Not (mf_DtLocked = True Or ADC("TLockYN") = "Y") Then
      If UCase(ADC.MenuCd) = UCase("BV") And Not .IsNew(RowNum) And UCase$(.Value(RowNum, "TdAck")) = "Y" And UCase$(ms_hTdAck) = "Y" Then
        .AllowDelete = (gs_ValidPrdLocSeq = "") Or ((InStr(1, "," + UCase(gs_ValidPrdLocSeq) + ",", "," + UCase(.Value(RowNum, "TdFrBLoc")) + ",") > 0) And _
                        (InStr(1, "," + UCase(gs_ValidPrdLocSeq) + ",", "," + UCase(.Value(RowNum, "TdToBLoc")) + ",") > 0))
      ElseIf UCase(ADC.MenuCd) = UCase("BV") And Not .IsNew(RowNum) And UCase$(.Value(RowNum, "TdAck")) = "N" And UCase$(ms_hTdAck) = "Y" Then
        .AllowDelete = (gs_ValidPrdLocSeq = "") Or (InStr(1, "," + UCase(gs_ValidPrdLocSeq) + ",", "," + UCase(.Value(RowNum, "TdFrBLoc")) + ",") > 0)
      ElseIf UCase(ADC.MenuCd) = UCase("BA") Then
        .AllowDelete = False
      Else
        .AllowDelete = True
      End If
     Else
        .AllowDelete = False
     End If
    'Manoj 2.10.0
    ' urmi ModDt of BA
    
    '****** Sachin 3.02 21-12-07 - Current Partition
    If UCase(ADC.MenuCd) = UCase("BA") And .IsNew(RowNum) = False Then
      ADC("wModDt") = moCn.GetFldVal("select ModDt from TxnD where " + _
                      "TdCoCd= '" + ADC("TCoCd") + "' and TdTc= '" + ADC("TTc") + _
                      "' and TdYy= '" + ADC("TYy") + "' and TdChr='" + ADC("TChr") + _
                      "' and TdNo= " + CStr(ADC("TNo")) + " and TdSr= " + CStr(GrdTxndA.Value(RowNum, "TdSr")) + " And TdPrtKey='" + ctCurrPrtn + "'")
    ElseIf .IsNew(RowNum) = True Then
      ADC("wModDt") = moCn.SrvrDate
    End If
    
    ' ########################  Manoj #### Ver: 2.0.10 #### Date: 25/12/2003  ########################
    If UCase(ADC.MenuCd) = "MLT" Or UCase(ADC.MenuCd) = "DT" Or UCase(ADC.MenuCd) = "BS" Or UCase(ADC.MenuCd) = "BM" Then
      If GrdTxndB.Rows <= 1 Then
        ADC("wTDDESC") = ""
      Else
        GrdTxndB.Row = 1
      End If
    End If
    ' ############################################  Manoj  ###########################################
  
    '*** (Jen 2.13)
    '****** Sachin 3.02 21-12-07 - Current Partition
    If UCase(ADC.MenuCd) = "BM" And Not .IsNew(RowNum) Then
      ADC("wTdSetNo") = moCn.GetFldVal("Select TdRefKey from Txnd where TdCoCd= '" + ADC("TCoCd") + "' " + _
                        "and TdTc= '" + ADC("TTc") + "' and TdYy= '" + ADC("TYy") + "' " + _
                        "and TdChr='" + ADC("TChr") + "' and TdNo= " + CStr(ADC("TNo")) + " " + _
                        "and TdSr= " + CStr(.Value(RowNum, "TdSr")) + " And TdPrtKey='" + ctCurrPrtn + "' ")
    Else
      ADC("wTdSetNo") = 0
    End If
    '
    '*** (Jen 2.13)
    'IG.89-When new comes it gets value for rejection code & Bag I/R from X
    If UCase(ADC.MenuCd) = "REJ" And .IsNew(RowNum) Then
      .Value(RowNum, "wTdBIr") = ADC("wTBIr")
      .Value(RowNum, "TdRjCd") = ADC("wTRjCd")
    End If
    
    mb_AddRow = True    ' ***** Manali 3.10.0 - 09/05/12 - Bar Code (Ctrl+N)
  End With
  CmdLN.Enabled = False
End Sub
Private Sub GrdTxndA_ColWhen(ByVal RowNum As Integer, ByVal ColName As String, Cancel As Boolean, ErrMsg As String, ByVal Action As MwfCtl.en_fgAction)
  '*** Check if last transaction of the bag.
      'If not then do no allow the user to edit any field in the grid
  '*** Set Helps For Each Field
  '*** If 1st row, then get the default values for TdBYy and TdBChr from the Head file
      'If Row> 1 then get the defa values for TdBYy and TdBChr from the previous row of the grid
      'If not GrdTxndA.IsNew the cannot edit TdBYy, TdBChr, TdBNo and TdBQty
  '*** TdByLoc and TdByWrk can be entered only if TdPrdYN= 'Y' (i.e. Production = 'Yes')
  '*** wTdPrdPts
      'when the user comes into this column check if there are any Production Points records.
      'If no then set the focus to CmdTP command button (this procedure automatically
      'generates the Production Points records in the TxndPrd Grid.
  
  With GrdTxndA
  If .Value(RowNum, "TdPrtKey") <> ctCurrPrtn Or mf_DtLocked = True Then Cancel = True: ErrMsg = "Cannot Edit Locked/Previous Partition Entry": Exit Sub
  
' ########################  Manoj #### Ver: 2.0.6 #### Date: 25/12/2003  ########################
  If UCase$(ADC("TLockYN")) = UCase$("Y") Then Cancel = True: ErrMsg = "Cannot Add/Edit/Delete Locked Entry": Exit Sub
' ###########################################  Manoj  ###########################################
  
  If mf_ValidLocSeq = False Then Cancel = True: ErrMsg = "The User Does Not Have Access To One Of The Locs": Exit Sub
  If UCase(ColName) <> UCase("wTdPrdPts") Then mi_Count = 1     '*** rel 1.10
  
    '*** Check If Last Transaction Of Bag ***
    '****** Sachin 3.02 21-12-07 - Current Partition
    If moCn.RecSeek("Select TdTc From Txnd Where TdCoCd='" + ADC("TCoCd") + "' and " + _
               "TdTc='" + ADC("TTc") + "' and TdYy='" + ADC("TYy") + "' and " + _
               "TdChr='" + ADC("TChr") + "' and TdNo=" + CStr(ADC("TNo")) + " and " + _
               "TdSr=" + CStr(.Value(RowNum, "TdSr")) + " and " + _
               "TdSrNo=0 And TdPrtKey='" + ctCurrPrtn + "'") And _
       (Not moCn.RecSeek("Select BYy from Bag where " + _
                    "BCoCd='" + ADC("TCoCd") + "' and " + _
                    "BYy='" + .Value(RowNum, "TdBYy") + "' and " + _
                    "BChr='" + .Value(RowNum, "TdBChr") + "' and " + _
                    "BNo= " + CStr(.Value(RowNum, "TdBNo")) + " and " + _
                    "BPrtKey= '" + ctCurrPrtn + "' and " + _
                    "BLstYy='" + ADC("TYy") + "' and " + _
                    "BLstKey=" + CStr(.Value(RowNum, "TdKey")))) Then
      Cancel = True: ErrMsg = "Cannot Edit Since Not The Last Transaction": Exit Sub
    End If
    
    '**************************************** urmi **************************************
'    Dim wBFrLoc As String, wBToLoc As String
'    'Manoj 2.10.0
'   ' If (UCase$(adc.MenuCd) = "BV" Or ms_TcTyp = "DT") And .IsNew(RowNum) <> True Then
'    'If (ms_TcTyp = "BV" Or ms_TcTyp = "DT") And .IsNew(RowNum) <> True Then
'    'Manoj 2.10.0
'    If .IsNew(RowNum) <> True And UCase(adc.MenuCd) <> UCase("BA") Then
'      If gs_ValidPrdLocSeq <> "" Then
'        '****** Sachin 3.02 21-12-07 - Current Partition
'        wBFrLoc = moCn.GetFldVal("Select BLoc from Bag where " + _
'                                 "BCoCd='" + adc("TCoCd") + "' and " + _
'                                 "BYy='" + .Value(RowNum, "TdBYy") + "' and " + _
'                                 "BChr='" + .Value(RowNum, "TdBChr") + "' and " + _
'                                 "BNo= " + CStr(.Value(RowNum, "TdBNo")) + " And BPrtKey='" + ctCurrPrtn + "' ")
'        Cancel = IIF((InStr(1, "," + UCase(gs_ValidPrdLocSeq) + ",", "," + UCase(wBFrLoc) + ",") <= 0), True, False)
'      Else
'        Cancel = False
'      End If
'      If Cancel = True Then
'        ErrMsg = "Cannot Edit record As Bag Loc. is Not Accessible to the User"
'        If ms_TcTyp = "DT" Then GrdTxndB.AllowAdd = False: GrdTxndB.AllowDelete = False: _
'                                CmdLs.Enabled = False: CmdIR.Enabled = False: CmdTP.Enabled = False
'        Exit Sub
'      ElseIf Not mf_DtLocked Then
'        If ms_TcTyp = "DT" Then GrdTxndB.AllowAdd = True: GrdTxndB.AllowDelete = True: _
'                                CmdLs.Enabled = True: CmdIR.Enabled = True: CmdTP.Enabled = True
'      End If
'    End If
'    If UCase(adc.MenuCd) = UCase("BA") Then .AllowDelete = False

    
    
  Dim wBFrLoc As String, wBToLoc As String
    'Manoj 2.10.0
    If (UCase$(ADC.MenuCd) = "BV" Or ms_TcTyp = "DT") And .IsNew(RowNum) <> True Then
    'If (ms_TcTyp = "BV" Or ms_TcTyp = "DT") And .IsNew(RowNum) <> True Then
    'Manoj 2.10.0
      If gs_ValidPrdLocSeq <> "" Then
        '****** Sachin 3.02 21-12-07 - Current Partition
        wBFrLoc = moCn.GetFldVal("Select BLoc from Bag where " + _
                                 "BCoCd='" + ADC("TCoCd") + "' and " + _
                                 "BYy='" + .Value(RowNum, "TdBYy") + "' and " + _
                                 "BChr='" + .Value(RowNum, "TdBChr") + "' and " + _
                                 "BNo= " + CStr(.Value(RowNum, "TdBNo")) + " And BPrtKey='" + ctCurrPrtn + "' ")
        Cancel = IIF((InStr(1, "," + UCase(gs_ValidPrdLocSeq) + ",", "," + UCase(wBFrLoc) + ",") <= 0), True, False)
      Else
        Cancel = False
      End If
      If Cancel = True Then
        ErrMsg = "Cannot Edit record As Bag Loc. is Not Accessible to the User"
        If ms_TcTyp = "DT" Then GrdTxndB.AllowAdd = False: GrdTxndB.AllowDelete = False: _
                                CmdLs.Enabled = False: CmdIR.Enabled = False: CmdTP.Enabled = False: CmdRwrk.Enabled = False
        Exit Sub
      ElseIf Not mf_DtLocked Then
        If ms_TcTyp = "DT" And ADC.MenuCd <> "RDT" Then
            GrdTxndB.AllowAdd = True: GrdTxndB.AllowDelete = True: _
                                CmdLs.Enabled = True: CmdIR.Enabled = True: CmdTP.Enabled = True
            If ms_RwrkModYN Then CmdRwrk.Enabled = True
        End If
      End If
    End If
      '**************************************** urmi **************************************
    'Manoj 2.10.0
    If UCase(ADC.MenuCd) = UCase("BV") And Not .IsNew(RowNum) And UCase$(.Value(RowNum, "TdAck")) = "Y" And UCase$(ms_hTdAck) = "Y" Then
      .AllowDelete = (gs_ValidPrdLocSeq = "") Or ((InStr(1, "," + UCase(gs_ValidPrdLocSeq) + ",", "," + UCase(.Value(RowNum, "TdFrBLoc")) + ",") > 0) And _
                      (InStr(1, "," + UCase(gs_ValidPrdLocSeq) + ",", "," + UCase(.Value(RowNum, "TdToBLoc")) + ",") > 0))
        If UCase(ColName) <> UCase("TdWaxWt") Then Cancel = True: ErrMsg = "Cannot Edit Value as Bag is Acknowledge ..": Exit Sub

    ElseIf UCase(ADC.MenuCd) = UCase("BV") And Not .IsNew(RowNum) And UCase$(.Value(RowNum, "TdAck")) = "N" And UCase$(ms_hTdAck) = "Y" Then
      .AllowDelete = (gs_ValidPrdLocSeq = "") Or (InStr(1, "," + UCase(gs_ValidPrdLocSeq) + ",", "," + UCase(.Value(RowNum, "TdFrBLoc")) + ",") > 0)
      
    ElseIf UCase(ADC.MenuCd) = UCase("BA") Then
      .AllowDelete = False
      
    Else
      .AllowDelete = True

    End If
    'Manoj 2.10.0

    Select Case ColName
    ' **** Zubin 211 **** '
    Case Is = UCase("wBagNo")
      '****** Sachin 3.02 21-12-07 - Current Partition
      If moCn.RecSeek("Select TdTc From Txnd Where TdCoCd='" + ADC("TCoCd") + "' and " + _
                      "TdTc='" + ADC("TTc") + "' and TdYy='" + ADC("TYy") + "' and " + _
                      "TdChr='" + ADC("TChr") + "' and TdNo=" + CStr(ADC("TNo")) + " and " + _
                      "TdSr=" + CStr(.Value(RowNum, "TdSr")) + " and TdSrNo=0 And TdPrtKey='" + ctCurrPrtn + "'") Then _
         Cancel = True: ErrMsg = "Cannot Edit Bag Number": Exit Sub
    ' **** Zubin 211 **** '
    
    Case Is = UCase("TdBYy")
    ' **** Zubin 211 **** '
    '''''      If .Value(RowNum, "wBagNo") <> "" Then Cancel = True: ErrMsg = "Cannot Edit Bag Year When Bag Number Is Read Using Barcode Reader": Exit Sub
    ' **** Zubin 211 **** '
      If .Value(RowNum, "TdBYy") = "" Then
        If RowNum = 1 Then .Value(RowNum, "TdBYy") = _
           moCn.GetFldVal("Select hYy from Head where HCoCd= '" + ADC("TCoCd") + "' and HCd='" + ctSelfCmCd + "'")
        If RowNum > 1 Then .Value(RowNum, "TdBYy") = .Value(RowNum - 1, "TdBYy")
      Else
        '****** Sachin 3.02 21-12-07 - Current Partition
        If moCn.RecSeek("Select TdTc From Txnd Where " + _
                        "TdCoCd='" + ADC("TCoCd") + "' and TdTc='" + ADC("TTc") + "' and " + _
                        "TdYy='" + ADC("TYy") + "' and TdChr='" + ADC("TChr") + "' and " + _
                        "TdNo=" + CStr(ADC("TNo")) + " and " + _
                        "TdSr=" + CStr(.Value(RowNum, "TdSr")) + " and TdSrNo=0 and TdPrtKey='" + ctCurrPrtn + "'") Then _
           Cancel = True: ErrMsg = "Cannot Edit Bag Year": Exit Sub
      End If
      Call HlpList.PMCd("YY")
    Case Is = UCase("TdBChr")
' **** Zubin 211 **** '
'''''      If .Value(RowNum, "wBagNo") <> "" Then Cancel = True: ErrMsg = "Cannot Edit Bag Character When Bag Number Is Read Using Barcode Reader": Exit Sub
' **** Zubin 211 **** '
      If .Value(RowNum, "TdBChr") = "" Then
        If RowNum = 1 Then .Value(RowNum, "TdBChr") = _
           moCn.GetFldVal("Select hBChr from Head where HCoCd= '" + ADC("TCoCd") + "' and HCd='" + ctSelfCmCd + "'")
        If RowNum > 1 Then .Value(RowNum, "TdBChr") = .Value(RowNum - 1, "TdBChr")
      Else
        '****** Sachin 3.02 21-12-07 - Current Partition
        If moCn.RecSeek("Select TdTc From Txnd Where TdCoCd='" + ADC("TCoCd") + "' and " + _
                        "TdTc='" + ADC("TTc") + "' and TdYy='" + ADC("TYy") + "' and " + _
                        "TdChr='" + ADC("TChr") + "' and TdNo=" + CStr(ADC("TNo")) + " and " + _
                        "TdSr=" + CStr(.Value(RowNum, "TdSr")) + " and TdSrNo=0 and TdPrtKey='" + ctCurrPrtn + "'") Then _
           Cancel = True: ErrMsg = "Cannot Edit Bag Character": Exit Sub
      End If
      Call HlpList.PMCd("BCHR")

    Case Is = UCase("TdBNo")
' **** Zubin 211 **** '
'''''    If .Value(RowNum, "wBagNo") <> "" Then Cancel = True: ErrMsg = "Cannot Edit Bag Number When Bag Number Is Read Using Barcode Reader": Exit Sub
' **** Zubin 211 **** '
      '****** Sachin 3.02 21-12-07 - Current Partition
      If moCn.RecSeek("Select TdTc From Txnd Where TdCoCd= '" + ADC("TCoCd") + "' and " + _
                      "TdTc= '" + ADC("TTc") + "' and TdYy= '" + ADC("TYy") + "' and " + _
                      "TdChr= '" + ADC("TChr") + "' and TdNo= " + CStr(ADC("TNo")) + " and " + _
                      "TdSr= " + CStr(.Value(RowNum, "TdSr")) + " and TdSrNo=0 and TdPrtKey='" + ctCurrPrtn + "'") Then _
         Cancel = True: ErrMsg = "Cannot Edit Bag Number": Exit Sub
      Dim whOthLocBag As String
      whOthLocBag = moCn.GetFldVal("Select hOthLocBag from Head where HCoCd= '" + ADC("TCoCd") + "' and HCd='" + ctSelfCmCd + "'")
      'Call HlpList.BNo(.Value(RowNum, "TdBYy"), .Value(RowNum, "TdBChr"), whOthLocBag, ADC("TFrBLoc"))
      Dim wDecPos As Integer
      If ms_TcTyp = "BS" And ADC("TtrfQty") > 0 Then wDecPos = InStr(1, MWLib.Div(RowNum, 2), ".")
      Dim wOd As String, wOdArr() As String
      If RowNum > 1 And ms_TcTyp = "BS" And ADC("TtrfQty") = 0 Then
        
                
        '****** Sachin 3.02 21-12-07 - Current Partition
        wOd = moCn.GetFldVal("Select BOdTc+'|'+BOdYy+'|'+BOdChr+'|'+Ltrim(str(BOdNo))+" + _
                             "'|'+BOdDmCd From Bag where " + _
                             "BCoCd= '" + ADC("TCoCd") + "' and " + _
                             "BYy= '" + .Value(1, "TdBYy") + "' and " + _
                             "BChr= '" + .Value(1, "TdBChr") + "' and " + _
                             "BNo= " + CStr(.Value(1, "TdBNo")) + " And BPrtKey='" + ctCurrPrtn + "' ")
        
        wOdArr = Split(wOd, "|")
        If UBound(wOdArr) = 4 Then
          Call HlpList.BNo(ADC("TCoCd"), .Value(RowNum, "TdBYy"), .Value(RowNum, "TdBChr"), whOthLocBag, ADC("TFrBLoc"), wOdArr(0), wOdArr(1), wOdArr(2), wOdArr(3), wOdArr(4))
        End If
                
      ElseIf ms_TcTyp = "BS" And ADC("TtrfQty") > 0 And wDecPos = 0 Then
      'ec.4- if bag transfer then suffix, size, order and customer check is included
        wOd = moCn.GetFldVal("Select BOdTc+'|'+BOdYy+'|'+BOdChr+'|'+Ltrim(str(BOdNo))+" + _
                             "'|'+BOdDmCd+'|'+BOdSfx+'|'+BOdDmSz+'|'+Ltrim(str(BOdSr)) From Bag where " + _
                             "BCoCd= '" + ADC("TCoCd") + "' and " + _
                             "BYy= '" + .Value(RowNum - 1, "TdBYy") + "' and " + _
                             "BChr= '" + .Value(RowNum - 1, "TdBChr") + "' and " + _
                             "BNo= " + CStr(.Value(RowNum - 1, "TdBNo")) + " And BPrtKey='" + ctCurrPrtn + "' ")

        wOdArr = Split(wOd, "|")
        
        Dim wOmCmCd As String, wBagAheadLocPrcs As String, wChkRm As String
        wOmCmCd = IIF(ChkHSMtchCmCd.Value = Checked, moCn.GetFldVal("Select OmCmCd from OrdMst where OmCoCd='" + ADC("TCoCd") + "' and OmTc='" + _
                  wOdArr(0) + "' and OmYy='" + wOdArr(1) + "' and OmChr='" + wOdArr(2) + "' and OmNo=" + CStr(wOdArr(3))), "")
        
        'ec.4 if bag ahead loc/process is given then it select valid locations/process from sequence entered.
        wBagAheadLocPrcs = ""
        If ADC("wHsLocCd") <> "" And ADC("wHsSeq") <> "" Then
          Dim wLocPrcsSeq As String, wPos As Integer, wSeq As String, i As Integer, wSeqArr() As String
          wLocPrcsSeq = moCn.GetFldVal("select vPDesc225 from vParam where vPCoCd ='" + ADC("TCoCd") + "' " + _
                  "and vPTyp ='" + IIF(ADC("wHsPR") = "P", "PRDSEQ", "PRCSSEQ") + "' and vPMCd ='" + ADC("wHsSeq") + "' ")
          
          wPos = InStr(1, wLocPrcsSeq, ADC("wHsLocCd"))
          If wPos > 0 Then
            wSeq = Mid(wLocPrcsSeq, wPos, Len(wLocPrcsSeq))
            wSeqArr = Split(wSeq, ",")
              For i = 0 To UBound(wSeqArr)
              wBagAheadLocPrcs = wBagAheadLocPrcs + IIF(wBagAheadLocPrcs = "", "", "','") + wSeqArr(i)
              Next i
          End If
        End If
        
        wOdArr(0) = IIF(ChkHSMtchOrdNo.Value = Checked, wOdArr(0), "")
        wOdArr(1) = IIF(ChkHSMtchOrdNo.Value = Checked, wOdArr(1), "")
        wOdArr(2) = IIF(ChkHSMtchOrdNo.Value = Checked, wOdArr(2), "")
        wOdArr(3) = IIF(ChkHSMtchOrdNo.Value = Checked, wOdArr(3), "")
        wOdArr(5) = IIF(ChkHSMtchSfx.Value = Checked, wOdArr(5), "")
        wOdArr(6) = IIF(ChkHSMtchSz.Value = Checked, wOdArr(6), "")
        
        Call HlpList.BNo(ADC("TCoCd"), .Value(RowNum, "TdBYy"), .Value(RowNum, "TdBChr"), whOthLocBag, ADC("TFrBLoc"), wOdArr(0), wOdArr(1), wOdArr(2), wOdArr(3), wOdArr(4), , , , wOdArr(5), wOdArr(6), wOmCmCd, wBagAheadLocPrcs, ADC("wHsPR"))
       
      Else
      'uma **** Repair added RepairYN
      If UCase(ADC.MenuCd) = "RDT" Or UCase(ADC.MenuCd) = "RBV" Then
        Call HlpList.BNo(ADC("TCoCd"), .Value(RowNum, "TdBYy"), .Value(RowNum, "TdBChr"), whOthLocBag, ADC("TFrBLoc"), , , , , , , , "Y")
        Else
        Call HlpList.BNo(ADC("TCoCd"), .Value(RowNum, "TdBYy"), .Value(RowNum, "TdBChr"), whOthLocBag, ADC("TFrBLoc"))
      End If
      End If
      
    Case Is = UCase("wTdBIr")
      If ADC("TTc") = "BS" And ADC("wTrfYN") = "Y" Then
        Cancel = True
        ErrMsg = "No need to enter Issue/Receipt for transfer"
      Else
        If ADC("TMultiIr") = "I" And GrdTxndA.Value(RowNum, "wTdBIr") <> "" Then Cancel = True: ErrMsg = "Cannot Edit Value": Exit Sub
        If ADC("TMultiIr") = "R" And GrdTxndA.Value(RowNum, "wTdBIr") <> "" Then Cancel = True: ErrMsg = "Cannot Edit Value": Exit Sub

        Call HlpList.PMCd("IR")
      End If
    Case Is = UCase("TdBQty")
        '/****************** Shilpa *******************/
        'ec.4- if bag transfer receiving bag qty will get from issue bag in previous row and
        'cannot modify the reciving bag qty
          If ms_TcTyp = "BS" And ADC("TtrfQty") <> 0 And GrdTxndA.Value(RowNum, "wTdBIR") = "R" Then
            If RowNum > 1 Then GrdTxndA.Value(RowNum, "TdBQty") = GrdTxndA.Value(RowNum - 1, "TdBQty")
            Cancel = True: ErrMsg = "Cannot Edit Quantity Of Receiving Bags": Exit Sub
          End If
          ' **** Zubin 211 **** '
          If ms_TcTyp = "BM" And GrdTxndA.Value(RowNum, "wTdBIR") = "R" Then _
              Cancel = True: ErrMsg = "Cannot Edit Quantity Of Receiving Bags": Exit Sub
          ' **** Zubin 211 **** '
          '****** Sachin 3.02 21-12-07 - Current Partition
          If moCn.RecSeek("Select TdTc From Txnd Where TdCoCd='" + ADC("TCoCd") + "' and " + _
                      "TdTc='" + ADC("TTc") + "' and TdYy='" + ADC("TYy") + "' and " + _
                      "TdChr='" + ADC("TChr") + "' and TdNo=" + CStr(ADC("TNo")) + " and " + _
                      "TdSr=" + CStr(.Value(RowNum, "TdSr")) + " and TdSrNo<>0 And TdPrtKey='" + ctCurrPrtn + "'") Then _
            Cancel = True: ErrMsg = "Cannot Edit Quantity": Exit Sub

    Case Is = UCase("TdPtQty")
      If ADC("TtrfQty") > 0 Then
        '*** Jay 2.13(CT) ***
        GrdTxndA.Value(RowNum, "TdPtQty") = 0
        '*** Jay 2.13(CT) ***
        Cancel = True: ErrMsg = "Cannot Edit Part Quantity in the Transfer": Exit Sub
      Else
        '*** (Jen 2.13 CT)
        '*** (Bef 27/11/06)
        'If ms_TcTyp = "BM" And moCn.GetFldVal("Select PValue from Param where PTYp= 'BCHR' " + _
        '                       "and PMCd= '" + .Value(RowNum, "TdBChr") + "' and PSCd= ''") <> "C" Then _
        '                       Cancel = True: ErrMsg = "Cannot Edit Part Quantity": Exit Sub
        '*** (Bef 27/11/06)
        '*** (27/11/06)
        Dim ws_BTyp As String
        ws_BTyp = moCn.GetFldVal("Select PValue from Param where PTYp= 'BCHR' " + _
                               "and PMCd= '" + .Value(RowNum, "TdBChr") + "' and PSCd= ''")
        If ms_TcTyp = "BM" And ws_BTyp <> "C" And ws_BTyp <> "F" Then _
                               Cancel = True: ErrMsg = "Cannot Edit Part Quantity": Exit Sub
        '*** (27/11/06)
        '*** (Jen 2.13 CT)
        
        '*** (Bef 2.13 CT)
        '' **** Zubin 211 **** '
        'If ms_TcTyp = "BM" Then Cancel = True: ErrMsg = "Cannot Edit Part Quantity": Exit Sub
        '' **** Zubin 211 **** '
        '*** (Bef 2.13 CT)
        
        '*** (Bef 27/11/06)
        ''*** Jay 2.13(CT) ***
        'If ms_TcTyp = "BS" Or ms_TcTyp = "REJ" Then
        '  If moCn.GetFldVal("Select PValue from Param where PTYp = 'BCHR' and PMCd= '" + .Value(RowNum, "TdBChr") + "' " + _
        '                    "and PSCd =''") <> "C" Then _
        '                    Cancel = True: ErrMsg = "Cannot Edit Part Quantity for Primary/ Flute Bags": Exit Sub
        'End If
        ''*** Jay 2.13(CT) ***
        '*** (Bef 27/11/06)
        '****** Sachin 3.02 24-12-07 - Current Partition
        If moCn.RecSeek("Select TdTc From Txnd Where TdCoCd='" + ADC("TCoCd") + "' and " + _
                      "TdTc='" + ADC("TTc") + "' and TdYy='" + ADC("TYy") + "' and " + _
                      "TdChr='" + ADC("TChr") + "' and TdNo=" + CStr(ADC("TNo")) + " and " + _
                      "TdSr=" + CStr(.Value(RowNum, "TdSr")) + " and TdSrNo<>0 and TdPrtKey='" + ctCurrPrtn + "'") Then _
         Cancel = True: ErrMsg = "Cannot Edit Part Quantity": Exit Sub
      End If
    '/****************** Shilpa *******************/
    Case Is = UCase("TdFrBLoc")
      Call HlpList.Loc(ADC("TCoCd"), "'P'")
    Case Is = UCase("TdToBLoc")
      Call HlpList.Loc(ADC("TCoCd"), "'P'")
    'uma *** new PRc
    'Case Is = UCase("TdBPrc")
    '  Call HlpList.Loc(ADC("TCoCd"), "'R'")
                
    Case Is = UCase("TdRjCd")
      Call HlpList.PMCd("REJCD")
    Case Is = UCase("TdWrk")
     Call HlpList.vPMCd(ADC("TCoCd"), "WORK", , , , , .Value(RowNum, "TdFrBLoc"))
    ' ***** Manali 3.6.0 - Worker Help Sorted On Process Loc
    Case Is = UCase("TdAck")
      If Not GrdTxndA.IsNew(RowNum) And GrdTxndA.Value(RowNum, "TdAck") = "Y" Then
        Cancel = True
        ErrMsg = "Once Bag is Acknoledged ('Y'), Can not change Acknowledgment Status."
      Else
        Call HlpList.PMCd("YN")
      End If
    Case Is = UCase("TdPrdYN")
      Call HlpList.PMCd("YN")
      Cancel = moCn.RecSeek("Select 'x' from TxndPrd where TpCoCd= '" + ADC("TCoCd") + "' and TpTc= '" + ADC("TTC") + _
                  "' and TpYy= '" + ADC("TYY") + "' and TpChr= '" + ADC("TCHR") + "' and TpNo=" + CStr(ADC("TNO")) + _
                  " and TpSr =" + CStr(GrdTxndA.Value(GrdTxndA.Row, "TdSr")))
      ErrMsg = "Cannot edit as Production points exist"
    Case Is = UCase("TdPDQty")
      'If (GrdTxndA.Value(RowNum, "TdBChr") <> "PD2" Or GrdTxndA.Value(RowNum, "TDPRDYN") = "N") Then Cancel = True: ErrMsg = "Cannot Edit PD Bag Qty": Exit Sub
      ms_BTyp = moCn.GetFldVal("Select PValue from Param where PTyp ='BCHR' and PMCd='" + GrdTxndA.Value(RowNum, "TdBChr") + "' and PSCd=''")
      If (ms_BTyp <> ctBTyp Or GrdTxndA.Value(RowNum, "TDPRDYN") = "N") Then Cancel = True: ErrMsg = "Cannot Edit PD Bag Qty": Exit Sub
      Cancel = moCn.RecSeek("Select 'x' from TxndPrd where TpCoCd= '" + ADC("TCoCd") + "' and TpTc= '" + ADC("TTC") + _
                  "' and TpYy= '" + ADC("TYY") + "' and TpChr= '" + ADC("TCHR") + "' and TpNo=" + CStr(ADC("TNO")) + _
                  " and TpSr =" + CStr(GrdTxndA.Value(GrdTxndA.Row, "TdSr")))
      If Cancel = True Then ErrMsg = "Cannot Edit PD Bag Qty": Exit Sub
    Case Is = UCase("TdByLoc")
      'If UCase(ms_TcTyp) = "DT" And GrdTxndA.Value(RowNum, "TdPrdYN") = "N" Then _
      '   Cancel = True: ErrMsg = "Cannot Enter ByLoc When Prd= 'N'"
      Call HlpList.Loc(ADC("TCoCd"), "'P'")
    Case Is = UCase("TdByWrk")
      'If UCase(ms_TcTyp) = "DT" And GrdTxndA.Value(RowNum, "TdPrdYN") = "N" Then _
      '   Cancel = True: ErrMsg = "Cannot Enter ByWorker When Prd= 'N'"
         Call HlpList.vPMCd(ADC("TCoCd"), "WORK", , , , , .Value(RowNum, "TdByLoc"))
      ' ***** Manali 3.6.0 - Worker Help Sorted On Process Loc
    Case Is = UCase("wTdPrdPts")
      If mi_Count = 1 And GrdTxndA.Value(RowNum, "TdPrdYN") = "Y" And GrdTP.Rows - 1 = 0 Then mi_Count = mi_Count + 1: CmdTP.SetFocus
    
    End Select
End With
End Sub
Private Sub GrdTxndA_ReCalc(ByVal RowNum As Integer, ByVal ColName As String, ByVal OldVal As Variant, ByVal SourceName As String, ByVal SourceColName As String)
  '*** TdBNo on TdBYy, TdBChr                   Normal
      'TdBNo = 0
  '*** wTdBIr on TdBDc                          Init
      'If TdBDc='C' then 'I'
      'If TdBDc='D' then 'R'
  '*** TdBDc on wTdBIr                          Normal
      'If wTdBIr= 'I' then TdBDc= 'C'
      'If wTdBIr= 'R' then TdBDc= 'D'
  '*** TdBQty on TdBYy, TdBChr, TdBNo           Normal
      'If (TcType= 'DT' or 'BV' or 'REJ') or (TcTyp = 'BS' and wTdBIr = 'I')
      'then BQty from Bag(BYy= TdByy, BChr= TdBChr, BNo= TdBNo)
  '*** TdPtQty on TdBYy, TdBChr, TdBNo          Normal
      'If (TcType= 'DT' or 'BV' or 'REJ') or (TcTyp = 'BS' and wTdBIr = 'I')
      'then BPtQty from Bag(BYy= TdByy, BChr= TdBChr, BNo= TdBNo)
  '*** TdBGrWt on TdBYy, TdBChr, TdBNo          Normal
      'BGrWt from Bag(BYy= TdByy, BChr= TdBChr, BNo= TdBNo)
      '(also calculated in Grd.B.Valid & in Copy Rm Options)
  '*** TdFrBLoc on TdBYy, TdBChr, TdBNo         Normal
      'BLoc from Bag(BYy= TdByy, BChr= TdBChr, BNo= TdBNo)
  '*** TdToBLoc on TdBYy, TdBChr, TdBNo         Normal
      'If TcType= 'BV' then Next location after the TdFrBLoc in the production Sequence
      '   mentioned in the OrdDsg file of the corresponding Bag
      'else ''
  '*** TdRjDc on TdBYy,TdBChr,TdBNo,TdBDc       Normal
      'If TcType= 'REJ' then
      '   If TdBDc= 'C' then TdRjDc= 'D'
      '   If TdBDc= 'D' then TdRjDc= 'C'
      'ElseIf TcType= 'MLT' then
      '   If TdBDc= 'C' then TdRjDc= 'C'
      'else ''
  '*** TdRjQty on TdBYy,TdBChr,TdBNo,TdBQty     Normal
      'If TcType= 'REJ' then
      '   TdRjQty= TdBQty
      'If TcType= 'MLT' then
      '   TdRjQty= BRjQty from Bag(BYy= TdBYy, BChr= TdBChr, BNo= TdBNo)
      'else 0
  '*** TdRjPtQty on TdBYy,TdBChr,TdBNo,TdPtQty  Normal
      'If TcType= 'REJ' then
      '   TdRjPtQty= TdPtQty
      'If TcType= 'MLT' then
      '   TdRjPtQty= BRjPtQty from Bag(BYy= TdBYy, BChr= TdBChr, BNo= TdBNo)
      'else 0
  '*** TdRjCd on TdRjDc                         Normal
      'If TdRjDc<> 'D' then TdRjCd= ''
  '*** TdWrk on TdBYy,TdBChr,TdBNo              Normal
      'If TcType= 'BV' then
      '   TdWrk= ctTdWrk
      'else
      '   TdWrk= BWrk from Bag(BYy= TdBYy, BChr= TdBChr, BNo= TdBNo)
  '*** TdAck on TdBYy,TdBChr,TdBNo              Normal
      'If TcType= 'BV' then
      '   TdAck= 'Y'
      'else
      '   Blank
  '*** TdPrdYN on TdBYy,TdBChr,TdBNo,TdWrk      Normal
      'If TcType= 'DT' and TdWrk= 'ZSELF' and Previous Worker <> 'ZSELF' then
      '   TdPrdYN= 'Y'
      'else
      '   Blank
  '*** TdByLoc on TdBYy,TdBChr,TdBNo,TdPrdYN    Normal
      'If (TcType= 'DT' and TdPrdYN= 'Y') or TcType= 'REJ' then
      '   TdByLoc= TdFrBLoc
      'else
      '   Blank
  '*** TdByWrk on TdBYy,TdBChr,TdBNo,TdPrdYN    Normal
      'If (TcType= 'DT' and TdPrdYN= 'Y') or TcType= 'REJ' then
      '   TdByWrk = BWrk from Bag(BYy= TdBYy, BChr= TdBChr, BNo= TdBNo)
      'else
      '   Blank
  '*** TdPrevYy on TdBYy,TdBChr,TdBNo           Normal
      'TdPrevYy = BLstYy from Bag(BYy= TdBYy, BChr= TdBChr, BNo= TdBNo)
  '*** TdPrevKey on TdBYy,TdBChr,TdBNo          Normal
      'TdPrevKey = BLstKey from Bag(BYy= TdBYy, BChr= TdBChr, BNo= TdBNo)
  
  ' **** Zubin 211 **** '
  Dim ws_BagNo(), wRBYN As String
     
  With GrdTxndA
    Select Case UCase(ColName)
'    Case Is = UCase("TdBIdNo")
'****** Sachin 3.02.0 Shifted to InitKey
'      If .Mode = fgmnorm Or .Mode = fgmCopy Or .Mode = fgminit Then
'      '****** Sachin 3.02 (14/09/07) Id Fields
'        .Value(RowNum, "TdBIdNo") = moCn.GetFldVal("Select BIdNo From Bag " + _
'                      " Where BCoCd='" + ADC("TCoCd") + "' And BYy='" + .Value(RowNum, "TdBYy") + "' And BChr='" + .Value(RowNum, "TdBChr") + "' And BNo=" + CStr(.Value(RowNum, "TdBNo")))
'      '****** Sachin 3.02 (14/09/07) Id Fields
'      End If
    
    Case Is = UCase("wTdBIr")
      If .Mode = fgminit Then
        If .Value(RowNum, "TdBDc") = "C" Then .Value(RowNum, "wTdBIr") = "I"
        If .Value(RowNum, "TdBDc") = "D" Then .Value(RowNum, "wTdBIr") = "R"
      End If
    Case Is = UCase("TdBDc")
      If .Mode = fgmnorm Or .Mode = fgmCopy Then
        If .Value(RowNum, "wTdBIr") = "I" Then .Value(RowNum, "TdBDc") = "C"
        If .Value(RowNum, "wTdBIr") = "R" Then .Value(RowNum, "TdBDc") = "D"
        ' **** Manali 3.8.0
        If ADC("TMultiIr") = "I" Then
          If .Value(RowNum, "wTdBIr") = "R" Then
            .Value(RowNum, "TdKey") = mi_TdKey
            .AllowAdd = False
'          Else
'            .AllowAdd = True
          End If
        End If
        ' **** Manali 3.8.0
      End If
    Case Is = UCase("TdBQty")
      If .Mode = fgmnorm Then
        Select Case UCase(ms_TcTyp)
        Case Is = "DT", "BV", "REJ"
          '****** Sachin 3.02 21-12-07 - Current Partition
          'uma *** Repair BRepQty/ BQty
          wRBYN = moCn.GetFldVal("Select BRepYN from Bag where BCoCd='" + ADC("TCoCd") + "' and BYy='" + .Value(RowNum, "TdBYy") + "' and " + _
                                     "BChr='" + .Value(RowNum, "TdBChr") + "' and BNo= " + CStr(.Value(RowNum, "TdBNo")) + IIF(gs_Partition = ctCurrPrtn, " And BPrtKey='" + ctCurrPrtn + "'", ""))
  
          .Value(RowNum, "TdBQty") = moCn.GetFldVal("Select case when '" + wRBYN + "'='Y' then BRepQty else BQty end from Bag where " + _
                                     "BCoCd='" + ADC("TCoCd") + "' and " + _
                                     "BYy='" + .Value(RowNum, "TdBYy") + "' and " + _
                                     "BChr='" + .Value(RowNum, "TdBChr") + "' and " + _
                                    "BNo= " + CStr(.Value(RowNum, "TdBNo")) + IIF(gs_Partition = ctCurrPrtn, " And BPrtKey='" + ctCurrPrtn + "'", ""))
    
        ms_BTyp = moCn.GetFldVal("Select PValue from Param where PTyp ='BCHR' and PMCd='" + GrdTxndA.Value(RowNum, "TdBChr") + "' and PSCd=''")
        If ms_BTyp = ctBTyp Then
          .Value(RowNum, "TdPDQty") = .Value(RowNum, "TdBQty")
        End If
          
        Case Is = "BS"
          If ADC("TtrfQty") = 0# Then
            'recalc on Bag No, wTdBIr
            If .Value(RowNum, "wTdBIr") = "I" Then
              '****** Sachin 3.02 21-12-07 - Current Partition
              .Value(RowNum, "TdBQty") = moCn.GetFldVal("Select BQty from Bag where " + _
                                       "BCoCd='" + ADC("TCoCd") + "' and " + _
                                       "BYy='" + .Value(RowNum, "TdBYy") + "' and " + _
                                       "BChr='" + .Value(RowNum, "TdBChr") + "' and " + _
                                       "BNo= " + CStr(.Value(RowNum, "TdBNo")) + IIF(gs_Partition = ctCurrPrtn, " And BPrtKey='" + ctCurrPrtn + "'", ""))
            ' ***** Manali 3.8.0 - Multil Issue in bag Split
            ElseIf ADC("TMultiIr") = "I" And .Value(RowNum, "wTdBIr") = "R" Then
              .Value(RowNum, "TdBQty") = moCn.GetFldVal("Select Sum(TdBQty) From Txnd where " + _
                                       "TdCoCd='" + ADC("TCoCd") + "' and " + _
                                       "TdTc='" + ADC("TTc") + "' and TdYy='" + ADC("TYy") + "' and " + _
                                       "TdChr= '" + ADC("TChr") + "' and TdNo=" + CStr(ADC("TNo")) + " and TdBDc='C' and TdSrNo=0 " + _
                                       IIF(gs_Partition = ctCurrPrtn, " And TdPrtKey='" + ctCurrPrtn + "'", ""))
            
            End If
          End If
' **** Zubin 211 **** '
        Case Is = "BM"
        '****** Sachin 3.02 21-12-07 - Current Partition
        .Value(RowNum, "TdBQty") = moCn.GetFldVal("Select BQty from Bag where " + _
                                  "BCoCd='" + ADC("TCoCd") + "' and " + _
                                  "BYy='" + .Value(RowNum, "TdBYy") + "' and " + _
                                  "BChr='" + .Value(RowNum, "TdBChr") + "' and " + _
                                  "BNo= " + CStr(.Value(RowNum, "TdBNo")) + IIF(gs_Partition = ctCurrPrtn, " And BPrtKey='" + ctCurrPrtn + "'", ""))
' **** Zubin 211 **** '
        End Select
      End If
    Case Is = UCase("TdPtQty")
      If .Mode = fgmnorm Then
        Select Case UCase(ms_TcTyp)
        Case Is = "DT", "BV", "REJ"
          '****** Sachin 3.02 21-12-07 - Current Partition
          .Value(RowNum, "TdPtQty") = moCn.GetFldVal("Select BPtQty from Bag where " + _
                                      "BCoCd='" + ADC("TCoCd") + "' and " + _
                                      "BYy='" + .Value(RowNum, "TdBYy") + "' and " + _
                                      "BChr='" + .Value(RowNum, "TdBChr") + "' and " + _
                                      "BNo= " + CStr(.Value(RowNum, "TdBNo")) + IIF(gs_Partition = ctCurrPrtn, " And BPrtKey='" + ctCurrPrtn + "'", ""))
        Case Is = "BS"
          '*** Jay 2.13(CT) ***
          If ADC("TtrfQty") = 0# Then
            'recalc on Bag No, wTdBIr
            If .Value(RowNum, "wTdBIr") = "I" Then
              '****** Sachin 3.02 21-12-07 - Current Partition
              .Value(RowNum, "TdPtQty") = moCn.GetFldVal("Select BPtQty from Bag where " + _
                                          "BCoCd='" + ADC("TCoCd") + "' and " + _
                                          "BYy='" + .Value(RowNum, "TdBYy") + "' and " + _
                                          "BChr='" + .Value(RowNum, "TdBChr") + "' and " + _
                                          "BNo= " + CStr(.Value(RowNum, "TdBNo")) + IIF(gs_Partition = ctCurrPrtn, " And BPrtKey='" + ctCurrPrtn + "'", ""))
            End If
          End If
          '*** Jay 2.13(CT) ***
' **** Zubin 211 **** '
        Case Is = "BM"
          '****** Sachin 3.02 21-12-07 - Current Partition
          .Value(RowNum, "TdPtQty") = moCn.GetFldVal("Select BPtQty from Bag where " + _
                                      "BCoCd='" + ADC("TCoCd") + "' and " + _
                                      "BYy='" + .Value(RowNum, "TdBYy") + "' and " + _
                                      "BChr='" + .Value(RowNum, "TdBChr") + "' and " + _
                                      "BNo= " + CStr(.Value(RowNum, "TdBNo")) + IIF(gs_Partition = ctCurrPrtn, " And BPrtKey='" + ctCurrPrtn + "'", ""))
' **** Zubin 211 **** '
        
        End Select
      End If
    Case Is = UCase("TdBGrWt")
      If .Mode = fgmnorm Then
        '****** Sachin 3.02 21-12-07 - Current Partition
        'uma *** Repair BRepQty/ BQty
        wRBYN = moCn.GetFldVal("Select BRepYN from Bag where BCoCd='" + ADC("TCoCd") + "' and BYy='" + .Value(RowNum, "TdBYy") + "' and " + _
                                     "BChr='" + .Value(RowNum, "TdBChr") + "' and BNo= " + CStr(.Value(RowNum, "TdBNo")) + IIF(gs_Partition = ctCurrPrtn, " And BPrtKey='" + ctCurrPrtn + "'", ""))
                                     
      .Value(RowNum, "TdBGrWt") = moCn.GetFldVal("Select case when '" + wRBYN + "'='Y' then isnull(BRepGrWt, 0) else isnull(BGrWt, 0) end from Bag where " + _
                                    "BCoCd='" + ADC("TCoCd") + "' and " + _
                                    "BYy='" + .Value(RowNum, "TdBYy") + "' and " + _
                                    "BChr='" + .Value(RowNum, "TdBChr") + "' and " + _
                                    "BNo= " + CStr(.Value(RowNum, "TdBNo")) + IIF(gs_Partition = ctCurrPrtn, " And BPrtKey='" + ctCurrPrtn + "'", ""))
      End If
      '***
      'Recalc Of TdBGrWt on BYy, BChr, BNo,    Normal
      '       TdBGrWt = vBGrWt (also calculated in Grd.B.Valid & in Copy Rm OPtions)
      '***
    Case Is = UCase("TdFrBLoc")
      If .Mode = fgmnorm Then
        '****** Sachin 3.02 21-12-07 - Current Partition
        'uma *** Repair BRepQty/ BQty
         wRBYN = moCn.GetFldVal("Select BRepYN from Bag where BCoCd='" + ADC("TCoCd") + "' and BYy='" + .Value(RowNum, "TdBYy") + "' and " + _
                                     "BChr='" + .Value(RowNum, "TdBChr") + "' and BNo= " + CStr(.Value(RowNum, "TdBNo")) + IIF(gs_Partition = ctCurrPrtn, " And BPrtKey='" + ctCurrPrtn + "'", ""))
                                     
        .Value(RowNum, "TdFrBLoc") = moCn.GetFldVal("Select case when '" + wRBYN + "'='Y' then BRepLoc else BLoc end from Bag where " + _
                                     "BCoCd='" + ADC("TCoCd") + "' and " + _
                                     "BYy='" + .Value(RowNum, "TdBYy") + "' and " + _
                                     "BChr='" + .Value(RowNum, "TdBChr") + "' and " + _
                                     "BNo= " + CStr(.Value(RowNum, "TdBNo")) + IIF(gs_Partition = ctCurrPrtn, " And BPrtKey='" + ctCurrPrtn + "'", ""))
      End If
    Case Is = UCase("TdToBLoc")
      If .Mode = fgmnorm Then
        Dim wLocSeq As String, wPos1 As Single, wPos2 As Single, wPos3 As Integer
        'wLenFrLoc As Single,
        If ms_TcTyp = "BV" Then
        '*** Jay 2.13(CT) ***
          If moCn.RecSeek("Select 'x' from Param where PTYp = 'BCHR' and PMCd= '" + .Value(RowNum, "TdBChr") + "' " + _
                            "and PSCd ='' and PValue = 'C'") Then
            '****** Sachin 3.02 21-12-07 - Current Partition
            wLocSeq = moCn.GetFldVal("Select vPDesc225 from vParam where " + _
                                  "vPCoCd= '" + ADC("TCoCd") + "' and vPTyp='PRDSEQ' and " + _
                                  "vPMCd= (Select Max(OctPrdSeq) from OrdCT " + _
                                  "join Bag on OctCoCd=BCoCd and OctTc=BOdTc and OctYy= BOdYy " + _
                                  "and OctChr=BOdChr and OctNo= BOdNo and OctSr= BOdSr and OctCtChr= BChr " + IIF(gs_Partition = ctCurrPrtn, " and OctPrtKey=BPrtKey ", "") + " where " + _
                                  "BCoCd='" + ADC("TCoCd") + "' and " + _
                                  "BYy= '" + .Value(RowNum, "TdBYy") + "' and " + _
                                  "BChr= '" + .Value(RowNum, "TdBChr") + "' and " + _
                                  "BNo= " + CStr(.Value(RowNum, "TdBNo")) + IIF(gs_Partition = ctCurrPrtn, " And BPrtKey='" + ctCurrPrtn + "')", ")"))
          Else
            '****** Sachin 3.02 21-12-07 - Current Partition
            wLocSeq = moCn.GetFldVal("Select vPDesc225 from vParam where " + _
                                  "vPCoCd= '" + ADC("TCoCd") + "' and vPTyp='PRDSEQ' and " + _
                                  "vPMCd= (Select OdPrdSeq from OrdDsg " + _
                                  "join Bag on OdIdNo=BOdIdNo And OdCoCd=BCoCd and OdTc=BOdTc and OdYy= BOdYy " + _
                                  "and OdChr=BOdChr and OdNo= BOdNo and OdSr= BOdSr " + IIF(gs_Partition = ctCurrPrtn, " And OdPrtKey=BPrtKey ", "") + " where " + _
                                  "BCoCd='" + ADC("TCoCd") + "' and " + _
                                  "BYy= '" + .Value(RowNum, "TdBYy") + "' and " + _
                                  "BChr= '" + .Value(RowNum, "TdBChr") + "' and " + _
                                  "BNo= " + CStr(.Value(RowNum, "TdBNo")) + IIF(gs_Partition = ctCurrPrtn, " And BPrtKey='" + ctCurrPrtn + "')", ")"))
                                 '(index=Bag)
          End If
        '*** Jay 2.13(CT) ***
        
          'wLocSeq = "," + wLocSeq + ","
          'wPos1 = InStr(1, wLocSeq, "," + .Value(RowNum, "TdFrBLoc") + ",")
          'wLenFrLoc = Len(.Value(RowNum, "TdFrBLoc"))
          'If wPos1 > 0 Then wPos2 = InStr(wPos1 + wLenFrLoc, wLocSeq, ",")
          'If wPos2 > 0 And wPos2 <> wPos1 Then _
          '  .Value(RowNum, "TdToBLoc") = Mid(wLocSeq, wPos1 + wLenFrLoc + 2, wPos2 - wPos1 - 1)
        
          wLocSeq = "," + wLocSeq + ","
          wPos1 = InStr(1, wLocSeq, "," + .Value(RowNum, "TdFrBLoc") + ",")
          If wPos1 > 0 Then wPos2 = InStr(wPos1 + 1, wLocSeq, ",")
          If wPos2 > 0 Then wPos3 = InStr(wPos2 + 1, wLocSeq, ",")
          If wPos3 > 0 And wPos3 <> wPos2 Then .Value(RowNum, "TdToBLoc") = Mid(wLocSeq, wPos2 + 1, wPos3 - wPos2 - 1)
        Else
          .Value(RowNum, "TdToBLoc") = ""
        End If
      End If
    Case Is = UCase("TdRjDc")
      If .Mode = fgmnorm Or .Mode = fgmCopy Then
        Select Case UCase(ms_TcTyp)
        Case Is = "REJ"
          If .Value(RowNum, "TdBDc") = "C" Then
            .Value(RowNum, "TdRjDc") = "D"
          ElseIf .Value(RowNum, "TdBDc") = "D" Then
            .Value(RowNum, "TdRjDc") = "C"
          Else
            .Value(RowNum, "TdRjDc") = ""
          End If
        Case Is = "MLT"
          .Value(RowNum, "TdRjDc") = "C"
        End Select
      End If
    Case Is = UCase("TdRjQty")
      If .Mode = fgmnorm Then
        Select Case UCase(ms_TcTyp)
        Case Is = "REJ"
          .Value(RowNum, "TdRjQty") = .Value(RowNum, "TdBQty")
        Case Is = "MLT"
          '****** Sachin 3.02 21-12-07 - Current Partition
          .Value(RowNum, "TdRjQty") = moCn.GetFldVal("Select BRjQty from Bag where " + _
                                      "BCoCd='" + ADC("TCoCd") + "' and " + _
                                      "BYy='" + .Value(RowNum, "TdBYy") + "' and " + _
                                      "BChr='" + .Value(RowNum, "TdBChr") + "' and " + _
                                      "BNo= " + CStr(.Value(RowNum, "TdBNo")) + IIF(gs_Partition = ctCurrPrtn, " And BPrtKey='" + ctCurrPrtn + "'", ""))
        End Select
      End If
    Case Is = UCase("TdRjPtQty")
      If .Mode = fgmnorm Or .Mode = fgmCopy Then
        Select Case UCase(ms_TcTyp)
        Case Is = "REJ"
          .Value(RowNum, "TdRjPtQty") = .Value(RowNum, "TdPtQty")
        Case Is = "MLT"
          '****** Sachin 3.02 21-12-07 - Current Partition
          .Value(RowNum, "TdRjPtQty") = moCn.GetFldVal("Select BRjPtQty from Bag where " + _
                                        "BCoCd='" + ADC("TCoCd") + "' and " + _
                                        "BYy='" + .Value(RowNum, "TdBYy") + "' and " + _
                                        "BChr='" + .Value(RowNum, "TdBChr") + "' and " + _
                                        "BNo= " + CStr(.Value(RowNum, "TdBNo")) + IIF(gs_Partition = ctCurrPrtn, " And BPrtKey='" + ctCurrPrtn + "'", ""))
        End Select
      End If
    Case Is = UCase("TdRjCd")
      If .Mode = fgmnorm Then
        If .Value(RowNum, "TdRjDc") <> "D" Then .Value(RowNum, "TdRjCd") = ""
      End If
    Case Is = UCase("TdWrk")
      If .Mode = fgmnorm Then
        If ms_TcTyp = "BV" Then
          .Value(RowNum, "TdWrk") = ctTdWrk
        Else
          '****** Sachin 3.02 21-12-07 - Current Partition
          .Value(RowNum, "TdWrk") = moCn.GetFldVal("Select BWrk from Bag where " + _
                                    "BCoCd='" + ADC("TCoCd") + "' and " + _
                                    "BYy='" + .Value(RowNum, "TdBYy") + "' and " + _
                                    "BChr='" + .Value(RowNum, "TdBChr") + "' and " + _
                                    "BNo= " + CStr(.Value(RowNum, "TdBNo")) + IIF(gs_Partition = ctCurrPrtn, " and BPrtKey='" + ctCurrPrtn + "'", ""))
        End If
      End If
    
    Case Is = UCase("TdAck")
      If .Mode = fgmnorm Then
        If ms_TcTyp = "BV" And ms_hTdAck = "N" Then
          .Value(RowNum, "TdAck") = "Y"
        Else
          If ms_TcTyp = "BV" And ms_hTdAck = "Y" Then
            Dim wFrBLoc As String, wToBLoc As String, wb_FrFlag As Boolean, wb_ToFlag As Boolean
            If gs_ValidPrdLocSeq <> "" Then
              wFrBLoc = .Value(RowNum, "TdFrBLoc")
              wb_FrFlag = IIF((InStr(1, "," + UCase(gs_ValidPrdLocSeq) + ",", "," + UCase(wFrBLoc) + ",") <= 0), True, False)
              wToBLoc = .Value(RowNum, "TdToBLoc")
              wb_ToFlag = IIF((InStr(1, "," + UCase(gs_ValidPrdLocSeq) + ",", "," + UCase(wToBLoc) + ",") <= 0), True, False)
              If wb_FrFlag = False And wb_ToFlag = False Then
                .Value(RowNum, "TdAck") = "Y"
              Else
                .Value(RowNum, "TdAck") = "N"
              End If
            Else
                .Value(RowNum, "TdAck") = "Y"
            End If
          ElseIf ms_TcTyp <> "BV" Then .Value(RowNum, "TdAck") = ""
          End If
        End If
      End If
    
    Case Is = UCase("TdPrdYN")
      If .Mode = fgmnorm Then
        If ms_TcTyp = "DT" Then
          '****** Sachin 3.02 21-12-07 - Current Partition
          ' ****** Manali 3.10.0 - 12/04/12 - PRDYn='Y' only once in a process for a bag by default
          .Value(RowNum, "TdPrdYN") = IIF(.Value(RowNum, "TdWrk") = ctTdWrk And _
                                      moCn.GetFldVal("Select BWrk from Bag where " + _
                                      "BCoCd='" + ADC("TCoCd") + "' and " + _
                                      "BYy='" + .Value(RowNum, "TdBYy") + "' and " + _
                                      "BChr='" + .Value(RowNum, "TdBChr") + "' and " + _
                                      "BNo= " + CStr(.Value(RowNum, "TdBNo")) + IIF(gs_Partition = ctCurrPrtn, " and BPrtKey='" + ctCurrPrtn + "'", "")) <> "ZSELF" And _
                                      Not moCn.RecSeek("Select 'x' from Txnd " + _
                                      "Join Bag On BCoCd=TdCoCd and BYy=TdBYY and BChr=TdBChr and BNo=TdBNo and BIdNo=TdBIdNo and BPrtKey=TdPrtKey " + _
                                      "Where TdCoCd='" + ADC("TCoCd") + "'  and TdBYY='" + .Value(RowNum, "TdBYy") + "' and TdBChr='" + .Value(RowNum, "TdBChr") + "' " + _
                                      "and TdBNo= " + CStr(.Value(RowNum, "TdBNo")) + " and TdFrBLoc='" + .Value(RowNum, "TdFrBLoc") + "' " + _
                                      "and TdYYKey<" + IIF(.IsNew(RowNum), "=", "") + " BLstYyKey and TdPrdYn='Y' "), _
                                      "Y", "N")
        Else
          .Value(RowNum, "TdPrdYN") = ""
        End If
      End If
    'umaaaaa
    Case Is = UCase("TdByLoc")
      If .Mode = fgmnorm Then
        If (ms_TcTyp = "DT" And .Value(RowNum, "TdPrdYN") = "Y") Or ms_TcTyp = "REJ" Then
          .Value(RowNum, "TdByLoc") = .Value(RowNum, "TdFrBLoc")
        Else
          .Value(RowNum, "TdByLoc") = ""
        End If
      End If
    Case Is = UCase("TdByWrk")
      If .Mode = fgmnorm Then
        If (ms_TcTyp = "DT" And .Value(RowNum, "TdPrdYN") = "Y") Or ms_TcTyp = "REJ" Then
          '****** Sachin 3.02 21-12-07 - Current Partition
          .Value(RowNum, "TdByWrk") = moCn.GetFldVal("Select BWrk from Bag where " + _
                                      "BCoCd='" + ADC("TCoCd") + "' and " + _
                                      "BYy='" + .Value(RowNum, "TdBYy") + "' and " + _
                                      "BChr='" + .Value(RowNum, "TdBChr") + "' and " + _
                                      "BNo= " + CStr(.Value(RowNum, "TdBNo")) + IIF(gs_Partition = ctCurrPrtn, " And BPrtKey='" + ctCurrPrtn + "'", ""))
        Else
          .Value(RowNum, "TdByWrk") = ""
        End If
      End If
    Case Is = UCase("TdPrevYy")
      If .Mode = fgmnorm Then
          '****** Sachin 3.02 21-12-07 - Current Partition
          .Value(RowNum, "TdPrevYy") = moCn.GetFldVal("Select BLstYy from Bag where " + _
                                     "BCoCd='" + ADC("TCoCd") + "' and " + _
                                     "BYy='" + .Value(RowNum, "TdBYy") + "' and " + _
                                     "BChr='" + .Value(RowNum, "TdBChr") + "' and " + _
                                     "BNo= " + CStr(.Value(RowNum, "TdBNo")) + IIF(gs_Partition = ctCurrPrtn, " And BPrtKey='" + ctCurrPrtn + "'", ""))
      End If
    Case Is = UCase("TdPrevKey")
      If .Mode = fgmnorm Then
        '****** Sachin 3.02 21-12-07 - Current Partition
        .Value(RowNum, "TdPrevKey") = moCn.GetFldVal("Select BLstKey from Bag where " + _
                                     "BCoCd='" + ADC("TCoCd") + "' and " + _
                                     "BYy='" + .Value(RowNum, "TdBYy") + "' and " + _
                                     "BChr='" + .Value(RowNum, "TdBChr") + "' and " + _
                                     "BNo= " + CStr(.Value(RowNum, "TdBNo")) + IIF(gs_Partition = ctCurrPrtn, " And BPrtKey='" + ctCurrPrtn + "'", ""))
      End If
      
    Case Is = UCase("wTdCmCd")
      If (.Mode = fgminit Or .Mode = fgmnorm) Then
        If GrdTxndA.Value(RowNum, "TdBNo") <> 0 Then
        .Value(RowNum, "wTdCmCd") = moCn.GetFldVal("Select OdOmCmCd From OrdDsg " + _
                                                    "join Bag on OdIdNo=BOdIdNo And OdCoCd= BCoCd and " + _
                                                    "OdTc= BOdTc and OdYy= BOdYy and OdChr= BOdChr and OdNo= BOdNo and OdSr= BOdSr " + IIF(gs_Partition = ctCurrPrtn, " And OdPrtKey=BPrtKey ", "") + _
                                                    "Where BCoCd='" + ADC("TCoCd") + "' and " + _
                                                    "BYy='" + GrdTxndA.Value(RowNum, "TdBYy") + "' and " + _
                                                    "BChr='" + GrdTxndA.Value(RowNum, "TdBChr") + "' and " + _
                                                    "BNo= " + CStr(GrdTxndA.Value(RowNum, "TdBNo")) + " and BPrtKey=(Case When '" + gs_Partition + "'='" + ctCurrPrtn + "' Then '" + ctCurrPrtn + "' Else BPrtKey End) ")
                                                    
          
        End If
      End If
      
    End Select
  End With
End Sub
Private Sub GrdTxndA_ColChange(ByVal RowNum As Integer, ByVal ColName As String, ByVal NewValue As Variant, Cancel As Boolean, ErrMsg As String)
  '*** TdBYy            valid PMCd('YY') or ''
  '*** TdBChr           valid PMCd('BCHR') or ''
  '*** TdBNo            valid BNo from Bag(BYy= TdBYy, BChr= TdBChr) or 0
  '*** wTdBIr           valid PMCd('IR') or ''
  '*** TdFrBLoc         valid LocCd from Loc(LocTyp= 'P')
  '*** TdToBLoc         If TcType = 'BV' then valid LocCd from Loc(LocTyp= 'P') else ''
  '*** TdRjCd           valid PMCd('REJCD') or ''
  '*** TdWrk            valid PMCd('WORK')
  '*** TdAck            If TcType = 'BV' then valid PMCd('YN') else ''
  '*** TdPrdYN          If TcType = 'DT' then valid PMCd('YN') or '' else ''
  '*** TdByLoc          valid LocCd from Loc(LocTyp= 'P')
  '*** TdByWrk          '' or valid PMCd('WORK')

  Dim wfhOthLocBag As Boolean
' **** Zubin 211 **** '
  Dim ws_BagNo() As String, ws_BagNoBarCd As String
  Dim ws_OldYy As String, ws_OldChr As String, ws_OldNo As Single
  Dim ws_BagTyp As String
' **** Zubin 211 **** '
  With GrdTxndA
    Select Case ColName
    Case Is = UCase("TdBYy")
      Cancel = (NewValue <> "") And (Not moCn.RecSeek("Select PMCd From Param Where PTyp= 'YY' And " + _
                                     "PMCd = '" + NewValue + "'"))
      ErrMsg = "Invalid Bag Year": Exit Sub
    Case Is = UCase("TdBChr")
    
                                     
      Cancel = (Not moCn.RecSeek("Select PMCd From Param Where PTyp= 'BCHR' And " + _
                                     "PMCd = '" + NewValue + "'"))
      If Cancel = False Then
        ws_BagTyp = moCn.GetFldVal("Select PValue From Param Where PTyp= 'BCHR' And " + _
                                     "PMCd = '" + NewValue + "'")
                                        
'        If ws_BagTyp <> ADC("") Then
'          Cancel = True
'          ErrMsg = "Invalid Bag Typ"
'        End If
      End If
    'TdCoCd, TdTc, TdYy, TdChr, TdNo, TdSr, TdSrNo
    
    Case Is = UCase("TdBNo")
      Dim wCnd As String
      If UCase(ADC.MenuCd) <> "BA" Then
        '****** Sachin 3.02 21-12-07 - Current Partition
        wCnd = " and IsNull((Select TdAck from Txnd where TdCoCd= BCoCd and TdSrNo = 0 and " + _
               " TdYy= BLstYy and TdKey= BLstKey and TdPrtKey=BPrtKey), '')<> 'N' "
      Else
        wCnd = ""
      End If

 If UCase(ADC.MenuCd) = "RDT" Or UCase(ADC.MenuCd) = "RBV" Then
        Cancel = (NewValue <> 0) And (Not moCn.RecSeek("Select BNo From Bag Where " + _
                                      "BCoCd='" + ADC("TCoCd") + "' and " + _
                                      "BYy= '" + .Value(RowNum, "TdBYy") + "' And " + _
                                      "BChr = '" + .Value(RowNum, "TdBChr") + "' And " + _
                                      "BNo= " + CStr(NewValue) + " and BRepYN ='Y' and BPrtKey='" + ctCurrPrtn + "' " + wCnd))
        If Cancel = True Then ErrMsg = "Invalid Bag Number Or Bag Closed or Bag is Unacknowledged"
      Else
      '****** Sachin 3.02 21-12-07 - Current Partition
        Cancel = (NewValue <> 0) And (Not moCn.RecSeek("Select BNo From Bag Where " + _
                                      "BCoCd='" + ADC("TCoCd") + "' and " + _
                                      "BYy= '" + .Value(RowNum, "TdBYy") + "' And " + _
                                      "BChr = '" + .Value(RowNum, "TdBChr") + "' And " + _
                                      "BNo= " + CStr(NewValue) + " and BCls='N' and BPrtKey='" + ctCurrPrtn + "' " + wCnd))
        If Cancel = True Then ErrMsg = "Invalid Bag Number Or Bag Closed or Bag is Unacknowledged"
        If Not Cancel And ms_TcTyp = "REJ" Then
          Cancel = (NewValue <> 0) And (Not moCn.RecSeek("Select BNo From Bag Where " + _
                                      "BCoCd='" + ADC("TCoCd") + "' and " + _
                                      "BYy= '" + .Value(RowNum, "TdBYy") + "' And " + _
                                      "BChr = '" + .Value(RowNum, "TdBChr") + "' And " + _
                                      "BNo= " + CStr(NewValue) + " and BCls='N' and (BQty <>0 or BPtQty<>0) and BPrtKey='" + ctCurrPrtn + "' " + wCnd))
          If Cancel = True Then ErrMsg = "Invalid Bag Number or Bag Closed or Bag is Unacknowledged or Bag has no more pieces"
        End If
       End If
       
      If Cancel = True Then Exit Sub
      '**************************************** urmi **************************************
   If moCn.RecSeek("select 'x' from Bag where " + _
          " BCoCd='" + ADC("TCoCd") + "' and " + _
                                      "BYy= '" + .Value(RowNum, "TdBYy") + "' And " + _
                                      "BChr = '" + .Value(RowNum, "TdBChr") + "' And " + _
                                      "BNo= " + CStr(NewValue) + _
                                     " and '" + CStr(ADC("TDt")) + "'<BOpnDt ") Then
    Cancel = True
    ErrMsg = "Transaction Date Cannot Be Less Than Bag Opening Date"
    Exit Sub
   End If
              
      
      Dim wBFrLoc As String, wBToLoc As String
      If UCase(ADC.MenuCd) <> UCase("BA") Then
        If gs_ValidPrdLocSeq <> "" Then
          '****** Sachin 3.02 21-12-07 - Current Partition
          wBFrLoc = moCn.GetFldVal("Select BLoc from Bag where " + _
                                     "BCoCd='" + ADC("TCoCd") + "' and " + _
                                     "BYy='" + .Value(RowNum, "TdBYy") + "' and " + _
                                     "BChr='" + .Value(RowNum, "TdBChr") + "' and " + _
                                     "BNo= " + CStr(NewValue) + " and BPrtKey='" + ctCurrPrtn + "' ")
          Cancel = IIF((InStr(1, "," + UCase(gs_ValidPrdLocSeq) + ",", "," + UCase(wBFrLoc) + ",") <= 0), True, False)
        Else
          Cancel = False
        End If
        If Cancel = True Then ErrMsg = "Invalid Bag Number As Bag Loc. is Not Accessible to the User": Exit Sub
        End If
      ' End If
       '*******************Geeta************Emr206*******06/04/04
      'If (ms_TcTyp = "BV" Or ms_TcTyp = "DT" Or ms_TcTyp = "BS") Then
      If (ms_TcTyp = "BV") Then
        Dim wOdSr As String, wBmOrdHldSeoYN As String
        '****** Sachin 3.02 21-12-07 - Current Partition
        
        wOdSr = moCn.GetFldVal("Select max(OdSr)  From  OrdDsg join Bag " + _
                               "on OdPrtKey=BPrtKey And OdIdNo=BOdIdNo And OdCoCd=BCocd and OdTc=BOdTc and OdYy=BOdYy and OdChr=BOdChr " + _
                               "and OdNo=BOdNo  and OdSr=BOdSr  " + _
                               " where OdHld='Y' and BCoCd= '" + ADC("TCoCd") + "' and BYy= '" + GrdTxndA.Value(RowNum, "TdBYy") + "' " + _
                               "and BChr= '" + GrdTxndA.Value(RowNum, "TdBChr") + "' and BNo= " + CStr(NewValue) + " and BPrtKey='" + ctCurrPrtn + "'")
        '578.13 if hBmOrdHldSeoYN=Y, order on hold and password not given it will not allow bag to move
        wBmOrdHldSeoYN = moCn.GetFldVal("Select hBmOrdHldSeoYN from Head where hCoCd='" + ADC("TCoCd") + "' and HCd='" + ctSelfCmCd + "'")
        If wBmOrdHldSeoYN = "Y" And ms_PwdFlag = False And wOdSr <> 0 Then
          Cancel = True: ErrMsg = "Order Design for this Bag is On Hold, Requires SEO Password": Exit Sub
        Else
          '578.13 mb_CopyBag added to suppress this message on copy
          If wOdSr <> 0 And mb_CopyBag = False Then MsgBox "The Order Design for this Bag is On Hold. Click OK to Continue.", vbInformation + vbOKOnly, "Information"
        End If
        'DispMsg "Sr : " + CStr(.Value(RowNum, "TdSr")) + ", Order Design is on Hold", etWarning
      End If
      '***************Geeta***********
     
' **** Zubin 211 **** '
    Case Is = UCase("wBagNo")
      If NewValue <> "" Then
        'uni.156 calling function to ignore CoCd if entered
        ws_BagNoBarCd = ParseBagNoBarCd(NewValue)
        ws_BagNo = Split(ws_BagNoBarCd, "/")
        If UBound(ws_BagNo) <> 2 Then Cancel = True: ErrMsg = "Invalid Bag No": Exit Sub
        If Not IsNumeric(ws_BagNo(2)) Then Cancel = True: ErrMsg = "Invalid Bag No": Exit Sub
        
        ws_OldYy = .Value(RowNum, "TDBYY")
        ws_OldChr = .Value(RowNum, "TDBCHR")
        ws_OldNo = .Value(RowNum, "TDBNO")
        
        .Value(RowNum, "TDBYY") = UCase(ws_BagNo(0))
        Call GrdTxndA_ColChange(RowNum, "TDBYY", ws_BagNo(0), Cancel, "")
        If Cancel = True Then ErrMsg = "Invalid Bag Year": .Value(RowNum, "TDBYY") = ws_OldYy: Exit Sub
        
        .Value(RowNum, "TDBCHR") = UCase(ws_BagNo(1))
        Call GrdTxndA_ColChange(RowNum, "TDBCHR", ws_BagNo(1), Cancel, "")
        If Cancel = True Then ErrMsg = "Invalid Bag Character": .Value(RowNum, "TDBCHR") = ws_OldChr: Exit Sub
        
        .Value(RowNum, "TDBNO") = ws_BagNo(2)
        Call GrdTxndA_ColChange(RowNum, "TDBNO", ws_BagNo(2), Cancel, "")
        If Cancel = True Then ErrMsg = "Invalid Bag Number": .Value(RowNum, "TDBNO") = ws_OldNo: Exit Sub
    
        ' ***** Manali 3.10.0 - 12/04/12 - Bar Code
        If (ms_TcTyp = "REJ") Then
          Dim wOldRjCd As String, wOldBIr As String
          wOldBIr = .Value(RowNum, "wTdBIr")
          wOldRjCd = .Value(RowNum, "TdRjCd")
          If ADC("wTBIr") <> "" Then
            .Value(RowNum, "wTdBIr") = ADC("wTBIr")
            Call GrdTxndA_ColChange(RowNum, "wTdBIr", ADC("wTBIr"), Cancel, "")
            If Cancel = True Then ErrMsg = "Invalid I/R": .Value(RowNum, "wTdBIr") = wOldBIr: Exit Sub
          End If
          If ADC("wTRjCd") <> "" Then
            .Value(RowNum, "TdRjCd") = ADC("wTRjCd")
            Call GrdTxndA_ColChange(RowNum, "TdRjCd", ADC("wTRjCd"), Cancel, "")
            If Cancel = True Then ErrMsg = "Invalid I/R": .Value(RowNum, "TdRjCd") = wOldRjCd: Exit Sub
          End If
        End If
        
        If (ms_TcTyp = "BV") Then
          Dim wOldToBLoc As String
          wOldToBLoc = .Value(RowNum, "TdToBLoc")
          If ADC("wTToBLoc") <> "" Then
            .Value(RowNum, "TdToBLoc") = ADC("wTToBLoc")
            Call GrdTxndA_ColChange(RowNum, "TdToBLoc", ADC("wTToBLoc"), Cancel, "")
            If Cancel = True Then ErrMsg = "Invalid To Loc": .Value(RowNum, "TdToBLoc") = wOldToBLoc: Exit Sub
          End If
        End If
        
        ' ***** Manali 3.10.0 - 09/05/12 - Bar Code (Ctrl+N)
        If (ms_TcTyp = "BV" Or ms_TcTyp = "DT" Or ms_TcTyp = "REJ") And mb_AddRow Then
           If ms_TcTyp = "DT" Then      '3.11.0
                .Col = 18
            Else
                If .Rows - 1 = RowNum Then
                  .AddItem
                  .Row = RowNum + 1
                  .Col = 3
                End If
            End If
        ' ***** Manali 3.10.0 - 09/05/12 - Bar Code (Ctrl+N)
        ElseIf ms_TcTyp = "MLT" And mb_AddRow Then
          If .Rows - 1 = RowNum Then
            .AddItem
            If .SaveRec(RowNum) Then
              Call EmptyRm
              .Row = RowNum + 1
              .Col = 3
            Else
              .Row = RowNum
              .Col = 3
            End If
          End If
        Else
          mb_AddRow = True    ' ***** Manali 3.10.0 - 09/05/12 - Bar Code (Ctrl+N)
        End If
        ' ***** Manali 3.10.0 - 12/04/12 - Bar Code
    Else
      mb_AddRow = True
    End If
' **** Zubin 211 **** '

    Case Is = UCase("wTdBIr")
      Cancel = (NewValue <> "") And (Not moCn.RecSeek("Select PMCd From Param Where PTyp= 'IR' And " + _
                                     "PMCd = '" + NewValue + "'"))
      ErrMsg = "Enter Issue Or Receipt (I/R)": Exit Sub
    
    Case Is = UCase("TdFrBLoc")
      If NewValue = ctFgLoc Then Cancel = True: ErrMsg = "Bag Cannot Be Moved From FG Location": Exit Sub
      
      'Manoj 2.10.0
      wfhOthLocBag = moCn.RecSeek("Select 'x' from Head where HCoCd= '" + ADC("TCoCd") + "' and HCd='" + ctSelfCmCd + "' And hOthLocBag = 'Y'")
      
      Cancel = (Not wfhOthLocBag And ADC("TFrBLoc") <> "" And ADC("TFrBLoc") <> NewValue) _
                Or Not moCn.RecSeek("Select LocCd From Loc Where " + _
                           "LocCoCd= '" + ADC("TCoCd") + "' and " + _
                           "LocTyp='P' and LocCd = '" + NewValue + "'")
      'Cancel = Not moCn.RecSeek("Select LocCd From Loc Where " + _
                           "LocCoCd= '" + adc("TCoCd") + "' and " + _
                           "LocTyp='P' and LocCd = '" + NewValue + "'")
      'Manoj 2.10.0
      If Cancel = True Then ErrMsg = "Invalid From Bag Location": Exit Sub
    
   Case Is = UCase("TdToBLoc")
      If ms_TcTyp = "BV" Then
        If .Value(RowNum, "TdFrBLoc") = NewValue Then _
          Cancel = True: ErrMsg = "To Bag Loc and From Bag Loc cannot be same": Exit Sub
        '5713 LocValidYN='Y' added to get valid location
        Cancel = Not moCn.RecSeek("Select LocCd From Loc Where " + _
                             "LocCoCd= '" + ADC("TCoCd") + "' and " + _
                             "LocTyp='P' and LocCd = '" + NewValue + "' and LocValidYN='Y'")
        If Cancel = True Then ErrMsg = "Invalid To Bag Location": Exit Sub
              
        If gs_ValidPrdLocSeq <> "" Then
          '****** Sachin 3.02 21-12-07 - Current Partition
'          wBFrLoc = moCn.GetFldVal("Select BLoc from Bag where " + _
'                                 "BCoCd='" + ADC("TCoCd") + "' and " + _
'                                 "BYy='" + .Value(RowNum, "TdBYy") + "' and " + _
'                                 "BChr='" + .Value(RowNum, "TdBChr") + "' and " + _
'                                 "BNo= " + CStr(.Value(RowNum, "TdBNo")) + " and BPrtKey='" + ctCurrPrtn + "'")
'          Cancel = IIF((InStr(1, "," + UCase(gs_ValidPrdLocSeq) + ",", "," + UCase(wBFrLoc) + ",") <= 0), True, False)
'          If Cancel = True Then ErrMsg = "Cannot Change As Bag Loc. is Not Accessible to the User": Exit Sub
'
          wBToLoc = .Value(RowNum, "TdToBLoc")
          Cancel = IIF((InStr(1, "," + UCase(gs_ValidPrdLocSeq) + ",", "," + UCase(wBToLoc) + ",") <= 0), True, False)
          Cancel = (.Value(RowNum, "TdAck") = "Y") And Cancel And .IsNew(RowNum) = False
          If Cancel = True Then ErrMsg = "Cannot change as Bag is Acknowledged ": Exit Sub
        End If
      
      Else
        If NewValue <> "" Then _
           Cancel = True: ErrMsg = "To Bag Location Should Not Be Entered": Exit Sub
      End If
      If NewValue = ctFgLoc Then Cancel = True: ErrMsg = "Bag Cannot Be Moved To FG Location": Exit Sub

    
    Case Is = UCase("TdRjCd")
      Cancel = (NewValue <> "") And (Not moCn.RecSeek("Select PMCd From Param Where " + _
                                     "PTyp= 'REJCD' And PMCd = '" + NewValue + "'"))
      ErrMsg = "Invalid Rejection Code": Exit Sub
    'uma *** new R
    'Case Is = UCase("TdBPrc")
     ' If NewValue = "" Then
     ' Cancel = True: ErrMsg = "Production Process Should Not Be Entered": Exit Sub
     ' Else
     ' Cancel = (NewValue <> "") And (Not moCn.RecSeek("Select LocCd From Loc Where " + _
                             "LocCoCd= '" + ADC("TCoCd") + "' and " + _
                             "LocTyp='R' and LocCd = '" + NewValue + "'"))
     ' ErrMsg = "Invalid Process": Exit Sub
     ' End If

    Case Is = UCase("TdWrk")
      Cancel = Not moCn.RecSeek("Select vPMCd From vParam Where " + _
                           "vPCoCd = '" + ADC("TCoCd") + "' and vPTyp= 'WORK' And " + _
                           "vPMCd = '" + NewValue + "'")
      ErrMsg = "Invalid Worker": Exit Sub
    
    Case Is = UCase("TdAck")
      If ms_TcTyp = "BV" Then
        'Dim wBToLoc As String
        Cancel = Not moCn.RecSeek("Select PMCd From Param Where PTyp= 'YN' And " + _
                           "PMCd = '" + NewValue + "'")
        If Cancel = True Then ErrMsg = "Enter Acknowledgement Yes Or No (Y/N)": Exit Sub
      
      '****************************** urmi *****************************************'
      'Manoj 2.10.0
'        If gs_ValidPrdLocSeq <> "" Then
'          wBFrLoc = .Value(RowNum, "TdFrBLoc")
'          Cancel = IIF((InStr(1, "," + UCase(gs_ValidPrdLocSeq) + ",", "," + UCase(wBFrLoc) + ",") <= 0), True, False)
'          Cancel = (NewValue = "N") And Cancel
'        Else
'          Cancel = False
'        End If
'        If Cancel = True Then ErrMsg = "Since No Access to From Location, cannot change " + _
'                                     "the Ackw to 'N'": Exit Sub
      'Manoj 2.10.0
        
        If gs_ValidPrdLocSeq <> "" Then
          wBToLoc = .Value(RowNum, "TdToBLoc")
          Cancel = IIF((InStr(1, "," + UCase(gs_ValidPrdLocSeq) + ",", "," + UCase(wBToLoc) + ",") <= 0), True, False)
          Cancel = (NewValue = "Y") And Cancel
        Else
          Cancel = False
        End If
        If Cancel = True Then ErrMsg = "Since No Access to To Location, cannot change " + _
                                     "the Ackw to 'Y'/ cannot delete acknowledged bag.": Exit Sub
      '**************************************** urmi **************************************
      Else
        If NewValue <> "" Then Cancel = True: ErrMsg = "Acknowledgement Should Not Be " + _
                                                     "Entered ": Exit Sub
      End If
      
    Case Is = UCase("TdPrdYN")
      Dim wDefSubPrc, wTRChr As String
      If ms_TcTyp = "DT" Then
        Cancel = (NewValue <> "") And (Not moCn.RecSeek("Select PMCd From Param Where " + _
                                     "PTyp= 'YN' And PMCd = '" + NewValue + "'"))
        If Cancel = True Then ErrMsg = "Enter Production Completed (Y/N)": Exit Sub
      Else
        If NewValue <> "" Then _
           Cancel = True: ErrMsg = "Production Completed Should Not Be Entered": Exit Sub
      End If
      'uma *** Repair Tc
'     wTRChr = moCn.GetFldVal("Select vPValue3 from vParam where vPCoCd='" + ADC("TCoCd") + "' and vPTyp='CHR' and vPMCD='" + ADC("TTc") + "' and vPSCd='" + ADC("TChr") + "'")
 '     If wTRChr = "Y" Then Cancel = True: ErrMsg = "No Production for Repair Bags": Exit Sub
      
    'uma PD Qty
    Case Is = UCase("TdPDQty")
      'If (GrdTxndA.Value(RowNum, "TdBChr") <> "PD2" Or GrdTxndA.Value(RowNum, "TDPRDYN") = "N") Then Cancel = True: ErrMsg = "Cannot Edit PD Bag Qty": Exit Sub
      ms_BTyp = moCn.GetFldVal("Select PValue from Param where PTyp ='BCHR' and PMCd='" + GrdTxndA.Value(RowNum, "TdBChr") + "' and PSCd=''")
      If (ms_BTyp <> ctBTyp Or GrdTxndA.Value(RowNum, "TDPRDYN") = "N") Then Cancel = True: ErrMsg = "Cannot Edit PD Bag Qty": Exit Sub
      
      Cancel = moCn.RecSeek("Select 'x' from TxndPrd where TpCoCd= '" + ADC("TCoCd") + "' and TpTc= '" + ADC("TTC") + _
                "' and TpYy= '" + ADC("TYY") + "' and TpChr= '" + ADC("TCHR") + "' and TpNo=" + CStr(ADC("TNO")) + _
                " and TpSr =" + CStr(GrdTxndA.Value(GrdTxndA.Row, "TdSr")))
      If Cancel = True Then ErrMsg = "Cannot Edit PD Bag Qty": Exit Sub
            
    Case Is = UCase("TdByLoc")
      If ms_TcTyp = "DT" And .Value(RowNum, "TdPrdYN") = "Y" And NewValue = "" Then _
        Cancel = True: ErrMsg = "Enter By Location": Exit Sub

      Cancel = (NewValue <> "") And (Not moCn.RecSeek("Select LocCd From Loc Where " + _
                                     "LocCoCd= '" + ADC("TCoCd") + "' and " + _
                                     "LocTyp='P' and LocCd = '" + NewValue + "'"))
      If Cancel = True Then ErrMsg = "Invalid By Location": Exit Sub
    Case Is = UCase("TdByWrk")
      'If ms_TcTyp = "DT" And .Value(RowNum, "TdPrdYN") = "Y" And NewValue = "" Then _
        Cancel = True: ErrMsg = "Enter By Worker": Exit Sub
' **** Zubin 211 (Instead of vParam, Param was looked into for 'WORK') **** '
'''''      'Cancel = (NewValue <> "") And (Not moCn.RecSeek("Select PMCd From Param Where " + _
'''''                                     "PTyp= 'WORK' And PMCd = '" + NewValue + "'"))
'''''      ''''''''If Cancel = True Then ErrMsg = "Invalid By Worker": Exit Sub
      Cancel = (NewValue <> "") And (Not moCn.RecSeek("Select 'x' From vParam Where " + _
                                     "vPTyp= 'WORK' And vPMCd = '" + NewValue + "'"))
      If Cancel = True Then ErrMsg = "Invalid By Worker": Exit Sub
' **** Zubin 211 **** '
    End Select
  End With
End Sub
Private Sub GrdTxndA_RowPostValid(ByVal RowNum As Integer, Cancel As Boolean, ErrMsg As String, ChangedColName As String)
          If Not moCn.RecSeek("Select TYy from Txn where " + _
                            "TCoCd='" + ADC("TCoCd") + "' and " + _
                            "TTc='" + ADC("TTc") + "' and TYy='" + ADC("TYy") + "' and " + _
                            "TChr='" + ADC("TChr") + "' and TNo=" + CStr(ADC("TNo"))) Then
            Cancel = True: ErrMsg = "||Invalid Voucher"
            Exit Sub
          End If
      
      If ADC("TtC") = "REJ" And GrdTxndA.Value(RowNum, "TdBdC") = "D" _
       And Not moCn.RecSeek("Select 'x' From Loc where LocCoCd= '" + ADC("TCocd") + "'" + _
       " and LocTyp= 'P' and LocCd ='" + GrdTxndA.Value(RowNum, "TdFrBLoc") + "' AND LocValidYN = 'Y'") Then
        Cancel = True: ErrMsg = "||Invalid Bag Fr Loc"
        Exit Sub
      End If
      
      Cancel = (Not moCn.RecSeek("Select 'x' From Loc Where " + _
                                     "LocCd = '" + GrdTxndA.Value(RowNum, "TdFrBLoc") + "' And LocCoCd = '" + ADC("TCoCd") + "'"))
      If Cancel = True Then ErrMsg = "||Invalid From Bag Loc": Exit Sub
      
      If ADC("TtC") = "REJ" And GrdTxndA.Value(RowNum, "TdBdC") = "D" _
       And Not moCn.RecSeek("Select 'x' From Loc where LocCoCd= '" + ADC("TCocd") + "'" + _
       " and LocTyp= 'P' and LocCd ='" + GrdTxndA.Value(RowNum, "TdFrBLoc") + "' AND LocValidYN = 'Y'") Then
        Cancel = True: ErrMsg = "||Invalid Bag Fr Loc"
        Exit Sub
      End If
            
      If ms_TcTyp = "BV" Then
        If GrdTxndA.Value(RowNum, "TdFrBLoc") = GrdTxndA.Value(RowNum, "TdToBLoc") Then _
          Cancel = True: ErrMsg = "||To Bag Loc and From Bag Loc cannot be same": Exit Sub
        If GrdTxndA.IsNew(RowNum) Then
          If Not moCn.RecSeek("Select 'x' From Loc where LocCoCd= '" + ADC("TCocd") + "' " + _
            "and LocTyp= 'P' and LocCd ='" + GrdTxndA.Value(RowNum, "TdToBLoc") + "' AND LocValidYN = 'Y'") Then
              Cancel = True: ErrMsg = "Invalid To Bag Location": Exit Sub
          End If
        End If
          
      End If

      If moCn.GetFldVal("Select hOthLocBag from Head where HCoCd= '" + ADC("TCoCd") + "' and HCd= 'zself'") = "N" _
           And ADC("TFrBLoc") <> GrdTxndA.Value(RowNum, "TdFrBLoc") And (ms_TcTyp = "BS" Or ms_TcTyp = "BM") Then
        Cancel = True: ErrMsg = "||From Bag Loc Should Be Equal To The Default Bag Location"
        Exit Sub
      End If
      
      If moCn.GetFldVal("Select hOthLocBag from Head where HCoCd= '" + ADC("TCoCd") + "' and HCd= 'zself'") = "N" _
          And ADC("TFrBLoc") <> IIF(GrdTxndA.Value(RowNum, "TdBdC") = "C", GrdTxndA.Value(RowNum, "TdFrBLoc"), _
          IIF(GrdTxndA.Value(RowNum, "TdBdC") = "D", GrdTxndA.Value(RowNum, "TdToBLoc"), GrdTxndA.Value(RowNum, "TdFrBLoc"))) _
          And (ms_TcTyp = "DT" Or ms_TcTyp = "BV" Or ms_TcTyp = "MLT") Then
        Cancel = True: ErrMsg = "||From Bag Loc Should Be Equal To The Default Bag Location"
        Exit Sub
      End If
      
      '5713.45 I/R cant blank
      If ADC("TTC") = "BM" Or ADC("TTC") = "BS" Or ADC("TTC") = "REJ" Then
        If (Not moCn.RecSeek("Select PMCd From Param Where PTyp= 'IR' And " + _
                                   "PMCd = '" + GrdTxndA.Value(RowNum, "wTdBIr") + "'")) Then _
          Cancel = True: ErrMsg = "||Invalid Issue Or Receipt (I/R) ": Exit Sub
      End If
      
              '5713 checking valid To location for new row


End Sub

Private Sub GrdTxndA_RowSave(ByVal RowNum As Integer, Cancel As Boolean, ErrMsg As String)
   With GrdTxndA
'     If GrdTxndA.IsNew(RowNum) Then
     ' Cancel = moCn.RecSeek("Select * From Bag Where BCoCd = '" & ADC("TCoCd") & "'" & _
                                         " And BYy = '" + GrdTxndA.Value(RowNum, "tDBYy") + "' And BChr = '" & GrdTxndA.Value(RowNum, "tDbcHR") + "' AND bnO=" + CStr(GrdTxndA.Value(RowNum, "tDBNO")) + _
                                        " and BMaxTxnDt>'" + CStr(ADC("TDt")) + "'")
                                         
     ' ErrMsg = "A Transaction of Later Date Exists For this Bag"
     'End If
     Select Case UCase(ms_TcTyp)
     
     Case Is = "MLT"
      .Value(RowNum, "TdRjDc") = "C"
     Case Is = "REJ"
      If .Value(RowNum, "TdBDc") = "C" Then .Value(RowNum, "TdRjDc") = "D"
      If .Value(RowNum, "TdBDc") = "D" Then .Value(RowNum, "TdRjDc") = "C"
     Case Else
      .Value(RowNum, "TdRjDc") = ""
     End Select
     Dim wCorrespondingPrcs As String
     wCorrespondingPrcs = moCn.GetFldVal("Select LocPrcs from Loc where LocCd='" + .Value(RowNum, "TdFrBLoc") + "' and LocCoCd='" + ADC("TCoCd") + "'")
     Dim wErrDesc As String, wValChg As Boolean, wChkLoc As Boolean
     'Bypass check-checking the bag chr is trail chr or not
     Dim wChkOdChr  As Boolean
     wChkOdChr = moCn.RecSeek("select BOdChr from Bag " + _
                "join vParam on vPCoCd ='" + ADC("TCoCd") + "' and vPTyp ='CHR' and vPMCd='SO' and vPSCd =BOdChr " + _
                "Where BCoCd='" + ADC("TCoCd") + "' And BYy='" + .Value(RowNum, "TdBYy") + "' And BChr='" + .Value(RowNum, "TdBChr") + "' And BNo=" + CStr(.Value(RowNum, "TdBNo")) + " and vPValue3 ='N' ")
    
     If UCase$(ADC.MenuCd) = "BV" Then
        Dim wRsOrdRmLst As MwfLib.MDORowSet
        mb_SeoUsr = moCn.RecSeek(" Select 'x' From Param Join UsrAccess on UaUsrCd = PMCd where PTyp= 'USR' and PMCd= '" + gs_UsrCd + "' " + _
                                " and UaMnuCd = '" + ADC("TTc") + "' and UaSeoYn= 'Y' and PValue= '" + Trim(TxtPwd.Text) + "'")
        mb_VldSeoPwd = moCn.RecSeek(" Select 'x' from Param Where PTyp= 'SEOPWD' and PMCD= 'SEOPWD' and " + _
                                    " PValue1='" + Trim(TxtPwd.Text) + "'")
        wValChg = Not moCn.RecSeek(" Select 'x' From Txnd Where TdPrtKey='" + .Value(RowNum, "TdPrtKey") + "' " + _
                                        " and TdCoCd='" + ADC("TCoCd") + "' " + _
                                        " and TdTc='" + ADC("TTc") + "' " + _
                                        " and TdYy='" + ADC("TYy") + "' " + _
                                        " and TdChr='" + ADC("TChr") + "' " + _
                                        " and TdNo=" + CStr(ADC("TNo")) + _
                                        " and TdSr=" + CStr(.Value(RowNum, "TdSr")) + _
                                        " and TdSrNo=" + CStr(.Value(RowNum, "TdSrNo")) + _
                                        " and TdToBLoc='" + .Value(RowNum, "TdToBLoc") + "' ")
                                        
      'Bypass check-bypass the tolerance check for trial chr=Y
      If wChkOdChr Then
        If moCn.RecSeek("Select 'x' From Head Where HCd='" + ctSelfCmCd + "' and HCoCd='" + ADC("TCoCd") + "' and HMinDiaProcess<>'' and  ','+HMinDiaProcess+',' Like '%," + wCorrespondingPrcs + ",%'") Then
          If (mb_ChkTol Or mb_ChkTolQty) And Not (mb_SeoUsr Or mb_VldSeoPwd) And wValChg Then
                wChkLoc = moCn.RecSeek("Select 'x' From Head Where HCd='" + ctSelfCmCd + "' and HCoCd='" + ADC("TCoCd") + "' and HMinDiaProcess<>'' and  ','+HMinDiaProcess+',' Like '%," + wCorrespondingPrcs + ",%'")
                If wChkLoc Then
                  If mb_ChkTol Or mb_ChkTolQty Then
                     Dim wb_DExists As Boolean, wb_CExists As Boolean
                     wb_DExists = False: wb_CExists = False
                     Set wRsOrdRmLst = moCn.OpenRes(" Select Distinct OrRmCtg,(Case when OrRmCtg In ('D','C') Then OrRmSCtg else '' End) as OrRmSCtg from OrdRm " + _
                                                       "    Join Bag On BCoCd=OrCoCd And BOdTc=OrTc And BOdYy=OrYy And BOdChr=OrChr And BOdNo=OrNo And BOdSr=OrSr " + _
                                                       " Where OrRmCtg In ('D','C') and BCoCd='" + gs_CoCd + "' And BYy='" + .Value(RowNum, "TdBYy") + "' And BChr='" + .Value(RowNum, "TdBChr") + "' And BNo=" + CStr(.Value(RowNum, "TdBNo")))
                     Do While Not (wRsOrdRmLst.EOF Or wRsOrdRmLst.BOF)
                        wErrDesc = ChkTolerance(ADC("TCoCd"), .Value(RowNum, "TdBYy"), .Value(RowNum, "TdBChr"), .Value(RowNum, "TdBNo"), wRsOrdRmLst!OrRmCtg, wRsOrdRmLst!OrRmSCtg, True, "", "", "", 0, 0, 0, 0, 0, mb_ChkTol, mb_ChkTolQty)
                        If wErrDesc <> "" Then Cancel = True: ErrMsg = wErrDesc: Exit Sub
                        If (Not wb_DExists And wRsOrdRmLst!OrRmCtg = "D") Or (Not wb_CExists And wRsOrdRmLst!OrRmCtg = "C") Then
                            If wRsOrdRmLst!OrRmCtg = "D" Then wb_DExists = True
                            If wRsOrdRmLst!OrRmCtg = "C" Then wb_CExists = True
                            wErrDesc = ChkTolerance(ADC("TCoCd"), .Value(RowNum, "TdBYy"), .Value(RowNum, "TdBChr"), .Value(RowNum, "TdBNo"), wRsOrdRmLst!OrRmCtg, "", True, "", "", "", 0, 0, 0, 0, 0, mb_ChkTol, mb_ChkTolQty)
                        End If
                        If wErrDesc <> "" Then Cancel = True: ErrMsg = wErrDesc: Exit Sub
                        wRsOrdRmLst.MoveNext
                    Loop
                  End If
                End If
          End If
        End If
      End If
    End If
    If UCase$(ADC.MenuCd) = "BV" Then
      'Bypass check-bypass the tolerance check for trial chr=Y
      If wChkOdChr Then
        If Not (mb_SeoUsr Or mb_VldSeoPwd) And wValChg Then
            wChkLoc = moCn.RecSeek("Select 'x' From Head Where HCd='" + ctSelfCmCd + "' and HCoCd='" + ADC("TCoCd") + "' and hChkCastLocs<>'' and  ','+hChkCastLocs+',' Like '%," + wCorrespondingPrcs + ",%'")
            If wChkLoc Then
                 wErrDesc = ChkTolerance(ADC("TCoCd"), .Value(RowNum, "TdBYy"), .Value(RowNum, "TdBChr"), .Value(RowNum, "TdBNo"), "", "", True, ADC("TTc"), ADC("TYy"), ADC("TChr"), ADC("TNo"), .Value(RowNum, "TdSr"), 0, 0, 0, True, False, "WCT")
                 If wErrDesc <> "" Then Cancel = True: ErrMsg = wErrDesc: Exit Sub
            End If
            wChkLoc = moCn.RecSeek("Select 'x' From Head Where HCd='" + ctSelfCmCd + "' and HCoCd='" + ADC("TCoCd") + "' and hChkWaxPullOutLocs<>'' and  ','+hChkWaxPullOutLocs+',' lIKE '%," + wCorrespondingPrcs + ",%'")
           
            If wChkLoc Then
                wErrDesc = ChkTolerance(ADC("TCoCd"), .Value(RowNum, "TdBYy"), .Value(RowNum, "TdBChr"), .Value(RowNum, "TdBNo"), "", "", True, ADC("TTc"), ADC("TYy"), ADC("TChr"), ADC("TNo"), .Value(RowNum, "TdSr"), 0, .Value(RowNum, "TdWaxWt"), 0, True, False, "WXT")
                If wErrDesc <> "" Then Cancel = True: ErrMsg = wErrDesc: Exit Sub
            End If
            Dim wRsMinTol As MDORowSet
            Dim wchkRmSubCtg As String
            Set wRsMinTol = moCn.OpenRes("Select * from vParam where vPTyp = 'CHKTOL' and vPMCd = '" + wCorrespondingPrcs + "' and charindex(','+'" + GrdTxndA.Value(RowNum, "TdBChr") + "'+',',','+vPdesc+',',0)>0")
            
            Do While Not (wRsMinTol.EOF Or wRsMinTol.BOF)
               wb_DExists = False: wb_CExists = False
               Dim wTolChkRmCtg As String
               Dim wCtrForRmCtg As Integer
               Dim wTolChkRmSCtg As String
               Dim ws_TolChkRmCtgArr() As String, i As Integer
               Dim ws_TolChkRmSCtgArr() As String, j As Integer
                 If wRsMinTol!vPValue1 <> "" Then
                   ws_TolChkRmCtgArr = Split(wRsMinTol!vPValue1, ",")
                   For wCtrForRmCtg = 0 To UBound(ws_TolChkRmCtgArr)
                      wTolChkRmCtg = ws_TolChkRmCtgArr(wCtrForRmCtg)
                      If (wRsMinTol!vPValue4 <> "") Then
                        ws_TolChkRmSCtgArr = Split(wRsMinTol!vPValue4, ",")
                        For j = 0 To UBound(ws_TolChkRmSCtgArr)
                            wTolChkRmSCtg = ws_TolChkRmSCtgArr(j)
                            wErrDesc = ChkTolerance(ADC("TCoCd"), .Value(RowNum, "TdBYy"), .Value(RowNum, "TdBChr"), _
                                                      .Value(RowNum, "TdBNo"), wTolChkRmCtg, wTolChkRmSCtg, True, "", "", "", _
                                                      0, 0, 0, 0, 0, True, True, "", True, wRsMinTol!vPDesc225, wRsMinTol!vPValue3)
                            If wErrDesc <> "" Then Cancel = True: ErrMsg = wErrDesc: Exit Sub
                        Next j
                      Else
                        wTolChkRmSCtg = wRsMinTol!vPValue4
                        wErrDesc = ChkTolerance(ADC("TCoCd"), .Value(RowNum, "TdBYy"), .Value(RowNum, "TdBChr"), _
                                                      .Value(RowNum, "TdBNo"), wTolChkRmCtg, wTolChkRmSCtg, True, "", "", "", _
                                                      0, 0, 0, 0, 0, True, True, "", True, wRsMinTol!vPDesc225, wRsMinTol!vPValue3)
                        If wErrDesc <> "" Then Cancel = True: ErrMsg = wErrDesc: Exit Sub
                      End If
                    Next wCtrForRmCtg
                  End If
                  wRsMinTol.MoveNext
          Loop
       End If
      End If
    End If
  If mi_Count = 1 And GrdTxndA.Value(RowNum, "TdPrdYN") = "Y" And GrdTP.Rows - 1 = 0 Then mi_Count = mi_Count + 1: Call CmdTP_GotFocus   'CmdTP.SetFocus
  End With
 ' If ADC("TTc") = "BS" Then
 '   ChkBSValidations RowNum, Cancel, ErrMsg
 ' End If
End Sub
Private Sub ChkBSValidations(ByVal RowNum As Integer, Cancel As Boolean, ErrMsg As String)
' -- only one parent allowed if MultiIR='I'
'  -- If grdtxnda.Value(Rownum,"TdBagChr")
If GrdTxndA.IsNew(RowNum) Then
  If ADC("TMultiIR") = "R" And GrdTxndA.Value(RowNum, "TdBDC") = "C" Then
    ml_bsnoofcreditbags = ml_bsnoofcreditbags + 1
  End If
  If ml_bsnoofcreditbags > 1 Then
    Cancel = True
    ErrMsg = "Cannot be more than 1 parent"
  End If
  If ADC("TMultiIR") = "I" And GrdTxndA.Value(RowNum, "TdBDC") = "D" Then
    ml_bsnoofdebitbags = ml_bsnoofdebitbags + 1
  End If
  If ml_bsnoofdebitbags > 1 Then
    Cancel = True
    ErrMsg = "Cannot be more than 1 child"
  End If
End If
End Sub
Private Function ChkBagWtLimit(ByVal RowNum As Integer) As String
    If UCase$(ADC.MenuCd) = "DT" Then
        If GrdTxndA.Row < 0 Or GrdTxndA.Row >= GrdTxndA.Rows Then Exit Function
        If RowNum < 0 Or RowNum >= GrdTxndB.Rows Then Exit Function
        If (moCn.GetFldVal("Select PDesc225 from Param join RmMst on PTyp='RMSCTG' and PMCd=RmCtg  and PSCd=RmSCtg " + _
                               " where RmCd='" + GrdTxndB.Value(RowNum, "TdRmCd") + "' And RmPrtKey='" + ctCurrPrtn + "' ") <> "Y") Then Exit Function
        
        Dim wCnd As String
        
        wCnd = " RrFrLn <= " + CStr(GrdTxndB.Value(RowNum, "TdRmSz")) + " And RrToLn >= " + CStr(GrdTxndB.Value(RowNum, "TdRmSz")) + " "

        Dim ws_RSz As MDORowSet
        If Not (mb_SeoUsr Or mb_VldSeoPwd) Then
        
           Set ws_RSz = moCn.OpenRes(" Select RrFixMinTol, RrFixMaxTol from " + _
                    " (Select top 1 RrSalRt, RrFixMinTol, RrFixMaxTol from RmMst join RmRt on RrCtg = RmCtg and RrSCtg = RmSCtg and RrTcTyp='STW' and RrCmCtg= 'C' and " + _
                    "RrCmCd= '" + ctSelfCmCd + "' " + _
                    "and (RrSTWGrp= RmSTWGrp or RrSTWGrp='') " + " and " + wCnd + _
                       " and RmCd='" + GrdTxndB.Value(RowNum, "TdRmCd") + "' order by RrSTWGrp desc) STWChart where  " + _
                       " (abs(" + _
                       CStr(GrdTxndB.Value(RowNum, "TdRmWt")) + ") < (RrFixMinTol * " + IIF(GrdTxndB.Value(RowNum, "TdRmQty") = 0, "1", CStr(GrdTxndB.Value(RowNum, "TdRmQty"))) + _
                       ") or abs(" + CStr(GrdTxndB.Value(RowNum, "TdRmWt")) + ") > (RrFixMaxTol * " + IIF(GrdTxndB.Value(RowNum, "TdRmQty") = 0, "1", CStr(GrdTxndB.Value(RowNum, "TdRmQty"))) + "))")
                       
                       
'            If ws_RSz.RecCount > 0 Then
 '             Set ws_RSz = moCn.OpenRes("Select RrSalRt, RrMinTol, RrMaxTol from RmMst join RmRt where RrTcTyp='STW' and RrCmCtg= 'C' and " + _
                    "RrCmCd= '" + ctSelfCmCd + "' and RrCtg=RmCtg and " + _
                    "RrSCtg=RmSCtg and RrSTWGrp= RmSTWGrp" + " and " + wCnd + _
                       " where RmCd='" + GrdTxndB.Value(RowNum, "TdRmCd"))
                               
                       
                    
                  '  + _
                    ", " + _
                    "IsNull((Select RrSalRt from RmRt where RrTcTyp='STW' and RrCmCtg= 'C' and " + _
                            "RrCmCd= '" + ctSelfCmCd + "' and RrCtg='" + mRmCtg + "' and " + _
                            "RrSCtg='" + mRmSCtg + "' and RrSTWGrp= '' and " + wCnd + "), 0))")

            
           ' Set ws_RSz = moCn.OpenRes("Select RrCstRt, RrSalRt " + _
                    "From RmMst join RmRt on RrTcTyp= 'Rsz' and RrCmCd= '" + ctSelfCmCd + "' and RrCtg= RmCtg and " + _
                    "RrSCtg= RmSCtg and " + _
                    CStr(GrdTxndB.Value(RowNum, "TdRmSz")) + " >= RrFrLn and " + CStr(GrdTxndB.Value(RowNum, "TdRmSz")) + " <= RrToLn " + _
                       " where RmCd='" + GrdTxndB.Value(RowNum, "TdRmCd") + " ' and (abs(" + _
                       CStr(GrdTxndB.Value(RowNum, "TdRmWt")) + ") < (RrCstRt * " + CStr(GrdTxndB.Value(RowNum, "TdRmQty")) + ") or abs(" + CStr(GrdTxndB.Value(RowNum, "TdRmWt")) + ") > (RrSalRt * " + CStr(GrdTxndB.Value(RowNum, "TdRmQty")) + "))")
            Do While Not (ws_RSz.EOF Or ws_RSz.BOF)
              ChkBagWtLimit = ChkBagWtLimit + "Issue of " + GrdTxndB.Value(RowNum, "TdRmCd") + " exceeds allowed Min Limit (" + Format(ws_RSz!RrFixMinTol, "##0.0000") + ") cts and Max Limit (" + Format(ws_RSz!RrFixMaxTol, "##0.0000") + ") cts per stone in Size To Wt (STW) chart."
              ws_RSz.MoveNext
            Loop
         End If
    End If
End Function
Private Sub GrdTxndA_InitKey(ByVal RowNum As Integer)
  Dim ws_TdYyKey As String
  Dim wf_TdKeyLen As Single
  With GrdTxndA
    If ms_TcTyp = "BS" Then
      If ADC("wTrfYN") = "N" Then
        If ADC("TMultiIr") <> "I" And .Value(RowNum, "TdBDc") = "D" Then
          .Store "TdRefYy", moCn.GetFldVal("Select TdYy from Txnd where " + _
                            "TdCoCd='" + ADC("TCoCd") + "' and " + _
                            "TdTc='" + ADC("TTc") + "' and TdYy='" + ADC("TYy") + "' and " + _
                            "TdChr='" + ADC("TChr") + "' and TdNo=" + CStr(ADC("TNo")) + " and " + _
                            "TdSrNo= 0 and TdBDc= 'C' and TdPrtKey='" + ctCurrPrtn + "'")
          
          .Store "TdRefKey", moCn.GetFldVal("Select TdKey from Txnd where " + _
                             "TdCoCd='" + ADC("TCoCd") + "' and " + _
                             "TdTc='" + ADC("TTc") + "' and TdYy='" + ADC("TYy") + "' and " + _
                             "TdChr='" + ADC("TChr") + "' and TdNo=" + CStr(ADC("TNo")) + " and " + _
                             "TdSrNo= 0 and TdBDc= 'C' and TdPrtKey='" + ctCurrPrtn + "'")
        
        ElseIf ADC("TMultiIr") = "I" And .Value(RowNum, "TdBDc") = "C" Then
          If .IsNew(RowNum) Then
            .Store "TdRefYy", ADC("TYy")
            .Store "TdRefKey", mi_TdKey
          End If
        Else
          .Store "TdRefYy", ""
          .Store "TdRefKey", 0
        End If
      Else
        If RowNum Mod 2 = 0 Then
           If RowNum > 1 Then
            .Store "TdRefYy", ADC("TYy")
            .Store "TdRefKey", .Value(RowNum - 1, "TdKey")
           End If
        Else
          .Store "TdRefYy", ""
          .Store "TdRefKey", 0
        End If
      End If
    ElseIf ms_TcTyp = "BM" Then
      .Store "TdRefYy", ""
      If .IsNew(RowNum) Then
        If RowNum = 1 Then
          .Store "TdRefKey", 1
        Else
          Dim wTdRefKey As Long
          If .Value(RowNum, "TdBDc") = "D" Then
            wTdRefKey = moCn.GetFldVal("Select TdRefKey from Txnd where TdCoCd= '" + ADC("TCoCd") + "' and " + _
                               "TdTc= '" + ADC("TTc") + "' and TdYy= '" + ADC("TYy") + "' and " + _
                               "TdChr= '" + ADC("TChr") + "' and TdNo= " + CStr(ADC("TNo")) + " and " + _
                               "TdSr= " + CStr(.Value(RowNum - 1, "TdSr")) + " and TdSrNo= 0 and TdPrtKey='" + ctCurrPrtn + "' ")
            .Store "TdRefKey", wTdRefKey
          Else
            .Store "TdRefKey", moCn.GetFldVal("Select TdRefKey+ " + CStr(IIF(.Value(RowNum - 1, "TdBDc") = "D", 1, 0)) + " " + _
                               "from Txnd where TdCoCd= '" + ADC("TCoCd") + "' and " + _
                               "TdTc= '" + ADC("TTc") + "' and TdYy= '" + ADC("TYy") + "' and " + _
                               "TdChr= '" + ADC("TChr") + "' and TdNo= " + CStr(ADC("TNo")) + " and " + _
                               "TdSr= " + CStr(.Value(RowNum - 1, "TdSr")) + " and TdSrNo= 0 and TdPrtKey='" + ctCurrPrtn + "' ")
          End If
        End If
      End If
    Else
      .Store "TdRefYy", ""
      .Store "TdRefKey", 0
    End If
    
    .Store "TdCoCd", ADC("TCoCd")
    .Store "TdTc", ADC("TTc")
    .Store "TdYy", ADC("TYy")
    .Store "TdDt", ADC("TDt")
    
    wf_TdKeyLen = Len(Trim(Str(.Value(RowNum, "TdKey"))))
    ws_TdYyKey = ADC("TYy") + Space(8 - wf_TdKeyLen) + Trim(Str(.Value(RowNum, "TdKey")))
    .Store "TdYyKey", ws_TdYyKey
    .Store "TdChr", ADC("TChr")
    .Store "TdNo", ADC("TNo")
    .Store "TdFrRmLoc", ""
    .Store "TdFrRmDc", ""
    .Store "TdRmCd", ""
    .Store "TdLotNo", ""
    .Store "TdRmSz", 0
    .Store "TdRmSz2", 0
    .Store "TdRmSz3", 0
    .Store "TdRmStkRt", 0
    .Store "TdRmQty", 0
    .Store "TdRmWt", 0
    .Store "TdToRmLoc", ""
    .Store "TdToRmDc", ""
    .Store "TdDustWt", 0
    .Store "TdPcPerCt", 0
    .Store "TdFbRmQty", ""
    .Store "TdFbRmWt", ""
    .Store "TdRtByQW", ""
    .Store "TdRmCtg", ""
    .Store "TdRmSCtg", ""
    .Store "TdFrRmLocTyp", ""
    .Store "TdToRmLocTyp", ""
    .Store "TdBstOdTc", ""
    .Store "TdBstOdYy", ""
    .Store "TdBstOdChr", ""
    .Store "TdBstOdNo", 0
    .Store "TdBstOdSr", 0
    .Store "TdBstDmCd", ""
    .Store "TdLotMstRecYN", "N"
    
    If .IsNew(RowNum) Then
'      .Store "TdTIdNo", moCn.GetFldVal("Select TIdNo From Txn " + _
                        "Where TCoCd='" + ADC("TCoCd") + "' And TTc='" + ADC("TTc") + "' And TYy='" + ADC("TYy") + "' And Tchr='" + ADC("TChr") + "' And TNo=" + CStr(ADC("TNo")) + " And TPrtKey='" + ctCurrPrtn + "'")
'      .Store "TdDt", ADC("TDt")
      .Store "TdRmCtg", ""
      .Store "TdRmSCtg", ""
      .Store "TdFrRmLocTyp", ""
      .Store "TdToRmLocTyp", ""
      .Store "InsUsr", UCase(Trim(gs_UsrCd))
      .Store "InsDt", moCn.GetFldVal("Select getdate()")
      .Store "TdBstOdTc", ""
      .Store "TdBstOdYy", ""
      .Store "TdBstOdChr", ""
      .Store "TdBstOdNo", 0
      .Store "TdBstOdSr", 0
      .Store "TdBstDmCd", ""
    End If
    .Store "TdTIdNo", moCn.GetFldVal("Select TIdNo From Txn " + _
                        "Where TCoCd='" + ADC("TCoCd") + "' And TTc='" + ADC("TTc") + "' And TYy='" + ADC("TYy") + "' And Tchr='" + ADC("TChr") + "' And TNo=" + CStr(ADC("TNo")) + " And TPrtKey='" + ctCurrPrtn + "'")
    
    .Store "TdBIdNo", moCn.GetFldVal("Select BIdNo From Bag " + _
                      " Where BCoCd='" + ADC("TCoCd") + "' And BYy='" + .Value(RowNum, "TdBYy") + "' And BChr='" + .Value(RowNum, "TdBChr") + "' And BNo=" + CStr(.Value(RowNum, "TdBNo")) + IIF(gs_Partition = ctCurrPrtn, " and BPrtKey='" + ctCurrPrtn + "'", ""))
    .Store "TdBmRsnCd", ""
                      
  End With
End Sub
Private Sub GrdTxndA_Valid()
  '*** Disable all 'A' Grid fields
  
'*** (Bef 16/07/05) (Bef 2.11 color)
'  CmdTP.Enabled = False
'  CmdBD.Enabled = False
'  CmdIR.Enabled = False
'  CmdBS.Enabled = False
'  CmdCR.Enabled = False
'  CmdER.Enabled = False
'  CmdVS.Enabled = False
'  CmdTot.Enabled = False
'  CmdLs.Enabled = False
'  'CmdCB.Enabled = False
'  CmdTR.Enabled = False
'*** (Bef 16/07/05) (Bef 2.11 color)
End Sub
Private Sub GrdTxndA_LostFocus()
  '*** (16/07/05) (2.11 Color)
  'If (mo_CurrentBtn Is Nothing) And Not (Me.ActiveControl Is Nothing) Then
  If ms_CurrentBtn = "" And Not (Me.ActiveControl Is Nothing) Then
    If TypeOf Me.ActiveControl Is MWCTL_BTN1 Then
      ms_CurrentBtn = UCase(Me.ActiveControl.Name)
    End If
  End If
  If ms_CurrentBtn <> UCase("CmdTP") Then CmdTP.Enabled = False
  If ms_CurrentBtn <> UCase("CmdBD") Then CmdBD.Enabled = False
  If ms_CurrentBtn <> UCase("CmdIR") Then CmdIR.Enabled = False
  'If ms_CurrentBtn <> UCase("CmdIE") Then CmdIE.Enabled = False 'Uma KBS
  If ms_CurrentBtn <> UCase("CmdBS") Then CmdBS.Enabled = False
  If ms_CurrentBtn <> UCase("CmdCR") Then CmdCR.Enabled = False
  If ms_CurrentBtn <> UCase("CmdER") Then CmdER.Enabled = False
  If ms_CurrentBtn <> UCase("CmdVS") Then CmdVS.Enabled = False
  If ms_CurrentBtn <> UCase("CmdTot") Then CmdTot.Enabled = False
  If ms_CurrentBtn <> UCase("CmdLs") Then CmdLs.Enabled = False
  If ms_CurrentBtn <> UCase("CmdCert") Then CmdCert.Enabled = False    '******Bhavna
  'CmdCB.Enabled = False
  If ms_CurrentBtn <> UCase("CmdTR") Then CmdTR.Enabled = False
  ' Zubin 212
  If ms_CurrentBtn <> UCase("CmdCRM") Then CmdCRM.Enabled = False
  ' Zubin 212
   If ms_CurrentBtn <> UCase("CmdRwrk") Then CmdRwrk.Enabled = False  '***********Bhavna
  
  'If Not (mo_CurrentBtn Is Nothing) Then
  '  mo_CurrentBtn.Enabled = True
  'End If
  ms_CurrentBtn = ""
  '*** (16/07/05) (2.11 Color)
End Sub
Private Sub GrdTxndB_GotFocus()
  ' **** Zubin 211 **** '
  If GrdTxndA.Row > 0 Then
    If moCn.RecSeek("Select 'x' from Param where PTyp= 'BCHR' and " + _
                    "PMCd= '" + GrdTxndA.Value(GrdTxndA.Row, "TdBChr") + "' and PSCd= '' and PValue= 'F'") Then
      CmdQW.Enabled = True
    Else
      CmdQW.Enabled = False
    End If
    If CmdLN.Visible = True Then CmdLN.Enabled = True     '***********Bhavna
  End If
  
  '****** Sachin 3.02 24-12-07
  If mf_DtLocked = True Or ADC("TLockYN") = "Y" Or ADC("TPrtKey") <> ctCurrPrtn Then
     GrdTxndB.AllowAdd = False: GrdTxndB.AllowDelete = False
  Else
     GrdTxndB.AllowAdd = True: GrdTxndB.AllowDelete = True
  End If
  
  ' **** Zubin 211 **** '
End Sub

Private Sub GrdTxndB_RowSave(ByVal RowNum As Integer, Cancel As Boolean, ErrMsg As String)
  Dim wIssueflg As Boolean
  '****** Sachin 3.5.0 - Tolerance ******
  With GrdTxndB
      Cancel = ValidateStkSize(GrdTxndB.Value(RowNum, "TDRMSZ"), GrdTxndB.Value(RowNum, "TDRMSZ2"), GrdTxndB.Value(RowNum, "TDRMSZ3"), ErrMsg)
      If Cancel = True Then Exit Sub
      'Bypass check -Checking order chr is trial chr or not
      Dim wChkOdChr As Boolean
      wChkOdChr = moCn.RecSeek("select BOdChr from Bag " + _
                  "join vParam on vPCoCd ='" + ADC("TCoCd") + "' and vPTyp ='CHR' and vPMCd='SO' and vPSCd =BOdChr " + _
                  "Where BCoCd='" + ADC("TCoCd") + "' And BYy='" + GrdTxndA.Value(GrdTxndA.Row, "TdBYy") + "' And BChr='" + GrdTxndA.Value(GrdTxndA.Row, "TdBChr") + "' And BNo=" + CStr(GrdTxndA.Value(GrdTxndA.Row, "TdBNo")) + " and vPValue3 ='N' ")
                                
      'If mb_ChkTol And (UCase$(ADC.MenuCd) = "DT" Or UCase$(ADC.MenuCd) = "BM" Or UCase$(ADC.MenuCd) = "BS") Then
      If (mb_ChkTol Or mb_ChkTolQty Or mb_ChkMetTol) And (UCase$(ADC.MenuCd) = "DT" Or UCase$(ADC.MenuCd) = "BM" Or UCase$(ADC.MenuCd) = "BS") Then
         ' **** Manali 3.6.2 - Check only in case of issuing RM (Sangeeta)
        wIssueflg = IIF((.Value(RowNum, "TdFrRmLoc") = "W" And .Value(RowNum, "TdFrRmDC") = "C") Or _
                                                             (.Value(RowNum, "TdToRmLoc") = "W" And .Value(RowNum, "TdFrRmDC") = "D"), True, False)
        mb_SeoUsr = moCn.RecSeek(" Select 'x' From Param Join UsrAccess on UaUsrCd = PMCd where PTyp= 'USR' and PMCd= '" + gs_UsrCd + "' " + _
                               " and UaMnuCd = '" + ADC("TTc") + "' and UaSeoYn= 'Y' and PValue= '" + Trim(TxtPwd.Text) + "'")
        mb_VldSeoPwd = moCn.RecSeek(" Select 'x' from Param Where PTyp= 'SEOPWD' and PMCD= 'SEOPWD' and " + _
                                   " PValue1='" + Trim(TxtPwd.Text) + "'")
         
        If Not (mb_SeoUsr Or mb_VldSeoPwd Or wIssueflg) Then
          Dim wErrDesc As String, wAddRmWt As Single, wRmCtg As String, wRmSCtg As String, wValChg As Boolean, rRmDet As MwfLib.MDORowSet
          Dim wAddRmQty As Single
          wErrDesc = ""
          Set rRmDet = moCn.OpenResultset("Select RmCtg, RmSCtg From RmMst Where RmCd='" + .Value(RowNum, "TdRmCd") + "'")
              
          If Not (rRmDet.EOF Or rRmDet.BOF) Then
            wRmCtg = rRmDet!RmCtg
            wRmSCtg = rRmDet!RmSCtg
          End If
            
          If mb_ChkTol Or mb_ChkTolQty Or mb_ChkMetTol Then
            wValChg = Not moCn.RecSeek(" Select 'x' From Txnd Where TdPrtKey='" + .Value(RowNum, "TdPrtKey") + "' " + _
                                       " and TdCoCd='" + ADC("TCoCd") + "' " + _
                                       " and TdTc='" + ADC("TTc") + "' " + _
                                       " and TdYy='" + ADC("TYy") + "' " + _
                                       " and TdChr='" + ADC("TChr") + "' " + _
                                       " and TdNo=" + CStr(ADC("TNo")) + _
                                       " and TdSr=" + CStr(GrdTxndA.Value(GrdTxndA.Row, "TdSr")) + _
                                       " and TdSrNo=" + CStr(.Value(RowNum, "TdSrNo")) + _
                                       " and TdRmCd='" + .Value(RowNum, "TdRmCd") + "' " + _
                                       " and TdFrRmLoc='" + .Value(RowNum, "TdFrrmLoc") + "' " + _
                                       " and TdToRmLoc='" + .Value(RowNum, "TdToRmLoc") + "' " + _
                                       " and TdFrRmDc='" + .Value(RowNum, "TdFrRmDc") + "' " + _
                                       " and Convert(Decimal(16,3), TdRmWt)=Convert(Decimal(16,3)," + CStr(.Value(RowNum, "TdRmWt")) + ")" + _
                                       " and Convert(Decimal(16,3), TdRmQty)=Convert(Decimal(16,3)," + CStr(.Value(RowNum, "TdRmQty")) + ")")
             
             wAddRmWt = .Value(RowNum, "TdRmWt") * IIF((.Value(RowNum, "TdFrRmLoc") = "W" And .Value(RowNum, "TdFrRmDC") = "C") Or _
                                                       (.Value(RowNum, "TdToRmLoc") = "W" And .Value(RowNum, "TdFrRmDC") = "D"), -1, 1)
            
             wAddRmQty = .Value(RowNum, "TdRmQty") * IIF((.Value(RowNum, "TdFrRmLoc") = "W" And .Value(RowNum, "TdFrRmDC") = "C") Or _
                                                       (.Value(RowNum, "TdToRmLoc") = "W" And .Value(RowNum, "TdFrRmDC") = "D"), -1, 1)
                                                       
                                                                      
               If (.IsNew(RowNum) Or wValChg) Then
                        Dim wChkWt As Boolean, wChkQty As Boolean
                        wChkWt = IIF((wRmCtg = "D" Or wRmCtg = "C"), mb_ChkTol, IIF((wRmCtg = "G" Or wRmCtg = "P" Or wRmCtg = "S" Or wRmCtg = "L"), mb_ChkMetTol, False))
                        wChkQty = IIF((wRmCtg = "D" Or wRmCtg = "C"), mb_ChkTolQty, False)
                  'Bypass check- ignore the tolerance check for trial chr=Y
                  If wChkOdChr Then
                    '3.11.0 - Tolerance Check
                    If wChkWt Or wChkQty Then
                        
                        If Not (wRmCtg = "G" Or wRmCtg = "P" Or wRmCtg = "S" Or wRmCtg = "L") And (wChkWt Or wChkQty) Then
                            wErrDesc = ChkTolerance(ADC("TCoCd"), GrdTxndA.Value(GrdTxndA.Row, "TdBYy"), GrdTxndA.Value(GrdTxndA.Row, "TdBChr"), GrdTxndA.Value(GrdTxndA.Row, "TdBNo"), wRmCtg, "", False, ADC("TTc"), ADC("TYy"), ADC("TChr"), ADC("TNo"), GrdTxndA.Value(GrdTxndA.Row, "TdSr"), .Value(RowNum, "TdSrNo"), wAddRmWt, wAddRmQty, wChkWt, wChkQty)
                            If wErrDesc <> "" Then GoTo TolErr
                            wErrDesc = ChkTolerance(ADC("TCoCd"), GrdTxndA.Value(GrdTxndA.Row, "TdBYy"), GrdTxndA.Value(GrdTxndA.Row, "TdBChr"), GrdTxndA.Value(GrdTxndA.Row, "TdBNo"), wRmCtg, wRmSCtg, False, ADC("TTc"), ADC("TYy"), ADC("TChr"), ADC("TNo"), GrdTxndA.Value(GrdTxndA.Row, "TdSr"), .Value(RowNum, "TdSrNo"), wAddRmWt, wAddRmQty, wChkWt, wChkQty)
                            If wErrDesc <> "" Then GoTo TolErr
                        End If
                    End If
                  End If
TolErr:
                    If wErrDesc <> "" Then
                        gs_ErrMsg = wErrDesc
                        Cancel = True: ErrMsg = wErrDesc: Exit Sub
                    End If
                    'wErrDesc = CheckTolErrMsg(GrdTxndA.Value(GrdTxndA.Row, "TdBYy"), GrdTxndA.Value(GrdTxndA.Row, "TdBChr"), GrdTxndA.Value(GrdTxndA.Row, "TdBNo"), GrdTxndA.Value(GrdTxndA.Row, "TdSr"), .Value(RowNum, "TdSrNo"), wAddRmWt, wRmCtg, wRmSCtg, "W")
              End If
            End If
        End If
    End If
    
    'Sachin 3.11.0 - Major Metal Loss on Bag
    Dim wLocTyp As String, wctg As String
    wLocTyp = moCn.GetFldVal("Select LocTyp From Loc Where LocCd='" + IIF(.Value(RowNum, "TdFrRmDc") = "D", .Value(RowNum, "TdFrRmLoc"), .Value(RowNum, "TdToRmLoc")) + "'")
    wctg = moCn.GetFldVal("Select RmCtg From RmMst Where RmCd='" + .Value(RowNum, "TdRmCd") + "'")
    If (wctg = "G" Or wctg = "P" Or wctg = "S" Or wctg = "L") And wLocTyp = "L" And .Value(RowNum, "TdByLoc") <> "" Then
        wValChg = Not moCn.RecSeek(" Select 'x' From Txnd Where TdPrtKey='" + .Value(RowNum, "TdPrtKey") + "' " + _
                                   " and TdCoCd='" + ADC("TCoCd") + "' " + _
                                   " and TdTc='" + ADC("TTc") + "' " + _
                                   " and TdYy='" + ADC("TYy") + "' " + _
                                   " and TdChr='" + ADC("TChr") + "' " + _
                                   " and TdNo=" + CStr(ADC("TNo")) + _
                                   " and TdSr=" + CStr(GrdTxndA.Value(GrdTxndA.Row, "TdSr")) + _
                                   " and TdSrNo=" + CStr(.Value(RowNum, "TdSrNo")) + _
                                   " and TdRmCd='" + .Value(RowNum, "TdRmCd") + "' " + _
                                   " and TdFrRmLoc='" + .Value(RowNum, "TdFrrmLoc") + "' " + _
                                   " and TdToRmLoc='" + .Value(RowNum, "TdToRmLoc") + "' " + _
                                   " and TdFrRmDc='" + .Value(RowNum, "TdFrRmDc") + "' " + _
                                   " and Convert(Decimal(16,3), TdRmWt)=Convert(Decimal(16,3)," + CStr(.Value(RowNum, "TdRmWt")) + ")" + _
                                   " and Convert(Decimal(16,3), TdRmQty)=Convert(Decimal(16,3)," + CStr(.Value(RowNum, "TdRmQty")) + ")")
       
      If Not (mb_SeoUsr Or mb_VldSeoPwd) Then
        If (.IsNew(RowNum) Or wValChg) Then
        'Bypass check- ignore the tolerance check for trial chr=Y
        If wChkOdChr Then
          wErrDesc = ChkBagLs(GrdTxndA.Value(GrdTxndA.Row, "TdBYy"), GrdTxndA.Value(GrdTxndA.Row, "TdBChr"), GrdTxndA.Value(GrdTxndA.Row, "TdBNo"), .Value(RowNum, "TdRmCd"), .Value(RowNum, "TdByLoc"), .Value(RowNum, "TdRmWt"), GrdTxndA.Value(GrdTxndA.Row, "TdSr"), .Value(RowNum, "TdSrNo"))
          If Trim(wErrDesc) <> "" Then
            gs_ErrMsg = wErrDesc
            Cancel = True: ErrMsg = wErrDesc: Exit Sub
          End If
        End If
        End If
      End If
    End If
    If Not (mb_SeoUsr Or mb_VldSeoPwd) Then
      If (.IsNew(RowNum) Or wValChg) Then
        wErrDesc = ChkBagWtLimit(RowNum)
        If Trim(wErrDesc) <> "" Then
          gs_ErrMsg = wErrDesc
          Cancel = True: ErrMsg = wErrDesc: Exit Sub
        End If
      End If
    End If
  End With
            
'          If mb_ChkTol Then
'             wValChg = Not moCn.RecSeek(" Select 'x' From Txnd Where TdPrtKey='" + .Value(RowNum, "TdPrtKey") + "' " + _
'                                       " and TdCoCd='" + ADC("TCoCd") + "' " + _
'                                       " and TdTc='" + ADC("TTc") + "' " + _
'                                       " and TdYy='" + ADC("TYy") + "' " + _
'                                       " and TdChr='" + ADC("TChr") + "' " + _
'                                       " and TdNo=" + CStr(ADC("TNo")) + _
'                                       " and TdSr=" + CStr(GrdTxndA.Value(GrdTxndA.Row, "TdSr")) + _
'                                       " and TdSrNo=" + CStr(.Value(RowNum, "TdSrNo")) + _
'                                       " and TdRmCd='" + .Value(RowNum, "TdRmCd") + "' " + _
'                                       " and TdFrRmLoc='" + .Value(RowNum, "TdFrrmLoc") + "' " + _
'                                       " and TdToRmLoc='" + .Value(RowNum, "TdToRmLoc") + "' " + _
'                                       " and TdFrRmDc='" + .Value(RowNum, "TdFrRmDc") + "' " + _
'                                       " and Convert(Decimal(16,3), TdRmWt)=Convert(Decimal(16,3)," + CStr(.Value(RowNum, "TdRmWt")) + ")")
'
'             wAddRmWt = .Value(RowNum, "TdRmWt") * IIF((.Value(RowNum, "TdFrRmLoc") = "W" And .Value(RowNum, "TdFrRmDC") = "C") Or _
'                                                       (.Value(RowNum, "TdToRmLoc") = "W" And .Value(RowNum, "TdFrRmDC") = "D"), -1, 1)
'
'             If (.IsNew(RowNum) Or wValChg) Then    'wRmCtg = "D" And
'
'                ' ***** Manali 3.8.0 - QTT at SCtg Level - New Parameter - TolTyp added
'                wErrDesc = ChkTolerance(ADC("TCoCd"), GrdTxndA.Value(GrdTxndA.Row, "TdBYy"), GrdTxndA.Value(GrdTxndA.Row, "TdBChr"), GrdTxndA.Value(GrdTxndA.Row, "TdBNo"), wRmCtg, wRmSCtg, False, ADC("TTc"), ADC("TYy"), ADC("TChr"), ADC("TNo"), GrdTxndA.Value(GrdTxndA.Row, "TdSr"), .Value(RowNum, "TdSrNo"), wAddRmWt, wAddRmQty)
'                'wErrDesc = CheckTolErrMsg(GrdTxndA.Value(GrdTxndA.Row, "TdBYy"), GrdTxndA.Value(GrdTxndA.Row, "TdBChr"), GrdTxndA.Value(GrdTxndA.Row, "TdBNo"), GrdTxndA.Value(GrdTxndA.Row, "TdSr"), .Value(RowNum, "TdSrNo"), wAddRmWt, wRmCtg, wRmSCtg, "W")
'                gs_ErrMsg = wErrDesc
'                If wErrDesc <> "" Then Cancel = True: ErrMsg = wErrDesc: Exit Sub
'
'               ' wErrDesc = CheckTolErrMsg(GrdTxndA.Value(GrdTxndA.Row, "TdBYy"), GrdTxndA.Value(GrdTxndA.Row, "TdBChr"), GrdTxndA.Value(GrdTxndA.Row, "TdBNo"), GrdTxndA.Value(GrdTxndA.Row, "TdSr"), .Value(RowNum, "TdSrNo"), wAddRmWt, wRmCtg, "", "W")
'               ' gs_ErrMsg = wErrDesc
'               ' If wErrDesc <> "" Then Cancel = True: ErrMsg = wErrDesc: Exit Sub
'             End If
'          End If
'
'          ' ***** Manali 3.8.0 - QTT at SCtg Level
'          If mb_ChkTolQty Then
'             wValChg = Not moCn.RecSeek(" Select 'x' From Txnd Where TdPrtKey='" + .Value(RowNum, "TdPrtKey") + "' " + _
'                                       " and TdCoCd='" + ADC("TCoCd") + "' " + _
'                                       " and TdTc='" + ADC("TTc") + "' " + _
'                                       " and TdYy='" + ADC("TYy") + "' " + _
'                                       " and TdChr='" + ADC("TChr") + "' " + _
'                                       " and TdNo=" + CStr(ADC("TNo")) + _
'                                       " and TdSr=" + CStr(GrdTxndA.Value(GrdTxndA.Row, "TdSr")) + _
'                                       " and TdSrNo=" + CStr(.Value(RowNum, "TdSrNo")) + _
'                                       " and TdRmCd='" + .Value(RowNum, "TdRmCd") + "' " + _
'                                       " and TdFrRmLoc='" + .Value(RowNum, "TdFrrmLoc") + "' " + _
'                                       " and TdToRmLoc='" + .Value(RowNum, "TdToRmLoc") + "' " + _
'                                       " and TdFrRmDc='" + .Value(RowNum, "TdFrRmDc") + "' " + _
'                                       " and Convert(Decimal(16,3), TdRmQty)=Convert(Decimal(16,3)," + CStr(.Value(RowNum, "TdRmQty")) + ")")
'
'             wAddRmQty = .Value(RowNum, "TdRmQty") * IIF((.Value(RowNum, "TdFrRmLoc") = "W" And .Value(RowNum, "TdFrRmDC") = "C") Or _
'                                                       (.Value(RowNum, "TdToRmLoc") = "W" And .Value(RowNum, "TdFrRmDC") = "D"), -1, 1)
'
'             If wRmCtg = "D" And (.IsNew(RowNum) Or wValChg) Then
'                wErrDesc = CheckTolErrMsg(GrdTxndA.Value(GrdTxndA.Row, "TdBYy"), GrdTxndA.Value(GrdTxndA.Row, "TdBChr"), GrdTxndA.Value(GrdTxndA.Row, "TdBNo"), GrdTxndA.Value(GrdTxndA.Row, "TdSr"), .Value(RowNum, "TdSrNo"), wAddRmQty, wRmCtg, wRmSCtg, "Q")
'                gs_ErrMsg = wErrDesc
'                If wErrDesc <> "" Then Cancel = True: ErrMsg = wErrDesc: Exit Sub
'             End If
'          End If
'          ' ***** Manali 3.8.0 - QTT at SCtg Level
'
'        End If
'      End If
 
  '****** Sachin 3.5.0 - Tolerance ******
  
  
  '*** (Jen 3.2.0) Temporarily Commented For User Control Error To Be Rectified (EOF or BOF error for 1st rec of grid)
  'With GrdTxndB
  '  If .IsNew(RowNum) Then
  '
  '    If (moCn.GetFldVal("Select hAvgSRtYn from Head where HCoCd= '" + ms_TCoCd + "' and " + _
  '                      "HCd='" + ctSelfCmCd + "'") = "Y") Or _
  '        (moCn.GetFldVal("Select hSizeStkRt from Head where HCoCd= '" + ms_TCoCd + "' and " + _
  '                      "HCd='" + ctSelfCmCd + "'") = "Y") Then
  '
  '      '******* urmi to chk for 'Trial order Chrs' ******************
  '      '****** Sachin 3.02 21-12-07 - Current Partition
  '      If moCn.GetFldVal("Select vPValue3 from vParam where " + _
  '              "vPCoCd= '" + ms_TCoCd + "' and vPTyp= 'CHR' and " + _
  '              "vPmCd= (select BOdTc from Bag where BCoCd= '" + ms_TCoCd + "' and " + _
  '                  "BYy= '" + GrdTxndA.Value(GrdTxndA.Row, "TdBYy") + "' and " + _
  '                  "BChr= '" + GrdTxndA.Value(GrdTxndA.Row, "TdBChr") + "' and " + _
  '                  "BNo= " + CStr(GrdTxndA.Value(GrdTxndA.Row, "TdBNo")) + " and BPrtKey='" + ctCurrPrtn + "') and " + _
  '              "vPSCd= (select BOdChr from Bag where BCoCd= '" + ms_TCoCd + "' and " + _
  '                  "BYy= '" + GrdTxndA.Value(GrdTxndA.Row, "TdBYy") + "' and " + _
  '                  "BChr= '" + GrdTxndA.Value(GrdTxndA.Row, "TdBChr") + "' and " + _
  '                  "BNo= " + CStr(GrdTxndA.Value(GrdTxndA.Row, "TdBNo")) + " and BPrtKey='" + ctCurrPrtn + "') ") = "N" Then
  '
  '        '******* urmi to chk for 'Trial order Chrs' ******************
  '        If moCn.GetFldVal("Select PDesc225 from Param join RmMst on PTyp='RMSCTG' and PMCd=RmCtg  and PSCd=RmSCtg " + _
  '                                 " where RmCd='" + CStr(.Value(RowNum, "TdRmCd")) + "' ") = "Y" Then
  '
  '          Dim wb_OrRmSz As Boolean, wb_RSz As Boolean
  '          '****** Sachin 3.02 21-12-07 - Current Partition
  '          wb_OrRmSz = moCn.RecSeek("Select OrLn1 From OrdRm join Bag on " + _
  '                       " OrOdIdNo=BOdIdNo And OrCoCd= BCoCd and OrTc= BOdTc and OrYy= BOdYy and " + _
  '                       "OrChr= BOdChr and OrNo= BOdNo And OrSr = BOdSr and OrPrtKey=BPrtKey " + _
  '                       "where BCoCd= '" + ADC("TCoCd") + "' and " + _
  '                       "BYy= '" + GrdTxndA.Value(GrdTxndA.Row, "TdBYy") + "' and " + _
  '                       "BChr= '" + GrdTxndA.Value(GrdTxndA.Row, "TdBChr") + "' and " + _
  '                       "BNo= " + CStr(GrdTxndA.Value(GrdTxndA.Row, "TdBNo")) + " and " + _
  '                       "OrRmCd= '" + CStr(.Value(RowNum, "TdRmCd")) + "' and " + _
  '                       "OrLn1= " + CStr(.Value(RowNum, "TdRmSz")) + " and OrPrtKey='" + ctCurrPrtn + "'")
  '          If wb_OrRmSz = True Then Cancel = False: Exit Sub
  '
  '          '****** Sachin 3.02 21-12-07 - Current Partition
  '          wb_RSz = moCn.RecSeek("Select RrRngSz From OrdRm join Bag on " + _
  '                    "OrOdIdNo=BOdIdNo And OrCoCd= BCoCd and OrTc= BOdTc and OrYy= BOdYy and " + _
  '                    "OrChr= BOdChr and OrNo= BOdNo And OrSr = BOdSr and OrPrtKey=BPrtKey " + _
  '                    "and OrRmCd= '" + .Value(RowNum, "TdRMCd") + "' " + _
  '                    "join RmRt on RrTcTyp= 'Rsz' and RrCmCd= '" + ctSelfCmCd + "' and RrCtg= OrRmCtg and " + _
  '                    "RrSCtg= OrRmSCtg and RrPrtKey=OrPrtKey " + _
  '                    "where BCoCd= '" + ADC("TCoCd") + "' and " + _
  '                    "BYy= '" + GrdTxndA.Value(GrdTxndA.Row, "TdBYy") + "' and " + _
  '                    "BChr= '" + GrdTxndA.Value(GrdTxndA.Row, "TdBChr") + "' and " + _
  '                    "BNo= " + CStr(GrdTxndA.Value(GrdTxndA.Row, "TdBNo")) + " and " + _
  '                    "OrLn1 >= RrFrLn and OrLn1 <= RrToLn and OrPrtKey='" + ctCurrPrtn + "' and " + _
  '                    CStr(.Value(RowNum, "TdRmSz")) + " >= RrFrLn and " + CStr(.Value(RowNum, "TdRmSz")) + " <= RrToLn")
  '
  '          If wb_RSz = True Then Cancel = False: Exit Sub
  '          Cancel = True: ErrMsg = "Invalid Raw Material Size"
  '          Exit Sub
  '        Else
  '          '****** Sachin 3.02 21-12-07 - Current Partition
  '          If moCn.GetFldVal("Select PValue3 from Param join RmMst on PTyp='RMSCTG' and PMCd=RmCtg  and PSCd=RmSCtg " + _
  '                                    " where RmCd='" + CStr(.Value(RowNum, "TdRmCd")) + "' and RmPrtKey='" + ctCurrPrtn + "' ") = "Y" Then
  '
  '            '**********Geeta*****Emr207******Ptr Check if Stk By ptr='Y' else Size Check
  '            Dim ws_OrRmPtr As String
  '            Dim wb_OrRmPtr As Boolean, wb_RPtr As Boolean
  '            ws_OrRmPtr = IIF(moCn.GetFldVal("Select PValue2 from Param join RmMst on PTyp='RMSCTG' and PMCd=RmCtg  and PSCd=RmSCtg " + _
  '                                  " where RmCd='" + CStr(.Value(RowNum, "TdRmCd")) + "' ") = "Y", "OrRmPtr", "Case when OrPrdQty=0 then 0 else OrPrdWt/OrPrdQty end")
  '
  '            '****** Sachin 3.02 21-12-07 - Current Partition
  '            wb_OrRmPtr = moCn.RecSeek("Select OrRmPtr From OrdRm join Bag on " + _
  '                       "OrOdIdNo=BOdIdNo And OrCoCd= BCoCd and OrTc= BOdTc and OrYy= BOdYy and " + _
  '                       "OrChr= BOdChr and OrNo= BOdNo And OrSr = BOdSr And OrPrtKey=BPrtKey " + _
  '                       "where BCoCd= '" + ADC("TCoCd") + "' and " + _
  '                       "BYy= '" + GrdTxndA.Value(GrdTxndA.Row, "TdBYy") + "' and " + _
  '                       "BChr= '" + GrdTxndA.Value(GrdTxndA.Row, "TdBChr") + "' and " + _
  '                       "BNo= " + CStr(GrdTxndA.Value(GrdTxndA.Row, "TdBNo")) + " and " + _
  '                       "OrRmCd= '" + CStr(.Value(RowNum, "TdRmCd")) + "' and " + _
  '                       "OrPrtKey= '" + ctCurrPrtn + "' and " + _
  '                       ws_OrRmPtr + "= " + CStr(.Value(RowNum, "TdRmSz")))
  '            If wb_OrRmPtr = True Then Cancel = False: Exit Sub
  '
  '            '****** Sachin 3.02 21-12-07 - Current Partition
  '            wb_RPtr = moCn.RecSeek("Select RrRngSz From OrdRm join Bag on " + _
  '                    "OrOdIdNo=BOdIdNo And OrCoCd= BCoCd and OrTc= BOdTc and OrYy=BOdYy and " + _
  '                    "OrChr= BOdChr and OrNo= BOdNo And OrSr = BOdSr and OrPrtKey=BPrtKey " + _
  '                    "and OrRmCd= '" + .Value(RowNum, "TdRMCd") + "' " + _
  '                    "join RmRt on RrTcTyp= 'Rsz' and RrCmCd= '" + ctSelfCmCd + "' and RrCtg= OrRmCtg and " + _
  '                    "RrSCtg= OrRmSCtg and RrPrtKey=OrPrtKey " + _
  '                    "where BCoCd= '" + ADC("TCoCd") + "' and " + _
  '                    "BYy= '" + GrdTxndA.Value(GrdTxndA.Row, "TdBYy") + "' and " + _
  '                    "BChr= '" + GrdTxndA.Value(GrdTxndA.Row, "TdBChr") + "' and " + _
  '                    "BNo= " + CStr(GrdTxndA.Value(GrdTxndA.Row, "TdBNo")) + " and " + _
  '                    "BPrtKey= '" + ctCurrPrtn + "' and " + _
  '                    ws_OrRmPtr + " >= RrFrLn and " + ws_OrRmPtr + " <= RrToLn and " + _
  '                    CStr(.Value(RowNum, "TdRmSz")) + " >= RrFrLn and " + CStr(.Value(RowNum, "TdRmSz")) + " <= RrToLn")
  '            If wb_RPtr = True Then Cancel = False: Exit Sub
  '
  '            Cancel = True: ErrMsg = "Invalid Raw Material Pointer": Exit Sub
  '          End If
  '        End If
  '      End If
  '    End If
  '
  '  End If
  'End With
  '*** (Jen 3.2.0) Temporarily Commented For User Control Error To Be Rectified (EOF or BOF error for 1st rec of grid)
    
End Sub

Private Sub GrdTxndB_RowWhen(ByVal RowNum As Integer)
  '*** If GrdTxndB.IsNew Then TdSrNo = Max + 1
  '*** put the TdDesc of the current row in the working field wTdDesc
  '*** If a new row is being added then give default value to the TdToRmLoc field

  With GrdTxndB
  
    '****** Sachin 3.02 20-12-07 - Setting the default value in the 'A' grid to Current Partition
    If .IsNew(RowNum) Then .Value(RowNum, "TdPrtKey") = ctCurrPrtn
    'MMAck-Setting default value for acknowledgement as Y
    If .IsNew(RowNum) Then .Value(RowNum, "TdAck") = "Y"
  
    If .Value(RowNum, "TdSrNo") = 0 Then _
       .Value(RowNum, "TdSrNo") = .MaxVal("TdSrNo") + 1
    
    ADC("wTdDesc") = .Value(RowNum, "TdDesc")
  
    If .Value(RowNum, "TdToRmLoc") = "" Then
      If RowNum = 1 Then .Value(RowNum, "TdToRmLoc") = ADC("TToRmLoc")
      If RowNum > 1 Then .Value(RowNum, "TdToRmLoc") = .Value(RowNum - 1, "TdToRmLoc")
    End If
  End With
  ' Manoj 2.9.0 Patch
  mf_AvgRtFldChg = False
  ' Manoj 2.9.0 Patch
  CmdLN.Enabled = True    '***********Bhavna added new lot
End Sub
Private Sub GrdTxndB_ColWhen(ByVal RowNum As Integer, ByVal ColName As String, Cancel As Boolean, ErrMsg As String, ByVal Action As MwfCtl.en_fgAction)
  '*** Check if last transaction of the bag. If then do not allow the user to change
    'the values of the fields in the grid
  '*** TdFrRmLoc
      'If TdFrRmLoc = Blank then give default value to the TdFrRmLoc field
      'If First Row then TdFrRmLoc = TFrRmLoc else TdFrRmLoc = value in previous row
  '*** wTdFrRmIr
      'If wTdFrRmIr = Blank then give default value to the working field wTdFrRmIr
      'If First Row then wTdFrRmIr = 'I' else wTdFrRmIr = value in previous row
  '*** TdRmCd
      'Show only those RM Cds in the help which belongs to the Order corresponding to the Bag
  '*** TdRmSz
      ' Check whether the Rm records have to be entered Size wise or Not
  '*** TdRmStkRt
      ' Check whether the Rm records have to be entered Stock Rate wise or Not
  
  Dim wFrRmLocTyp As String, wToRmLocTyp As String
  'StkRt-Getting stock rate display Y\N & Edit Y\N from head table
  wStkNoDisp = moCn.GetFldVal("Select hStkRtNoDsp from Head where hcocd='" + gs_CoCd + "'")
  wStkNoEdit = moCn.GetFldVal("Select hStkRtNoEdit from Head where hcocd='" + gs_CoCd + "'")
    
  With GrdTxndB
  '****** Sachin 3.02.0 - Lock on Partition Key
  If .Value(RowNum, "TdPrtKey") <> ctCurrPrtn Or mf_DtLocked = True Then Cancel = True: ErrMsg = "Cannot Edit Locked/Previous Partition Entry": Exit Sub
  
' ########################  Manoj #### Ver: 2.0.6 #### Date: 25/12/2003  ########################
  If UCase$(ADC("TLockYN")) = UCase$("Y") Then Cancel = True: ErrMsg = "Cannot Add/Edit/Delete Locked Record": Exit Sub
' ###########################################  Manoj  ###########################################
  If mf_ValidLocSeq = False Then Cancel = True: ErrMsg = "The User Does Not Have Access To One Of The Locs": Exit Sub
  
  Dim wRmCtg As String, wRmSCtg As String, wGrpBySz As String, wGrpByPtr As String, wGrpByRt As String, wHAvgStkRt As String
  If ADC("TTC") = "BS" And ADC("TtrfQty") <> 0# Then
    Cancel = True: Exit Sub
  End If
  
    '****** Sachin 3.02 21-12-07 - Current Partition
    If moCn.RecSeek("Select TdTc From Txnd Where " + _
               "TdCoCd='" + ADC("TCoCd") + "' and TdTc='" + ADC("TTc") + "' and " + _
               "TdYy='" + ADC("TYy") + "' and TdChr='" + ADC("TChr") + "' and " + _
               "TdNo=" + CStr(ADC("TNo")) + " and " + _
               "TdSr=" + CStr(GrdTxndA.Value(GrdTxndA.Row, "TdSr")) + " and " + _
               "TdSrNo=" + CStr(.Value(RowNum, "TdSrNo")) + " and TdPrtKey='" + ctCurrPrtn + "'") And _
       (Not moCn.RecSeek("Select BYy from Bag where BCoCd='" + ADC("TCoCd") + "' and " + _
                    "BYy='" + GrdTxndA.Value(GrdTxndA.Row, "TdBYy") + "' and " + _
                    "BChr='" + GrdTxndA.Value(GrdTxndA.Row, "TdBChr") + "' and " + _
                    "BNo= " + CStr(GrdTxndA.Value(GrdTxndA.Row, "TdBNo")) + " and " + _
                    "BLstYy='" + ADC("TYy") + "' and " + _
                    "BLstKey=" + CStr(GrdTxndA.Value(GrdTxndA.Row, "TdKey")) + " and BPrtKey='" + ctCurrPrtn + "'")) Then
      Cancel = True: ErrMsg = "Cannot Edit Since Not The Last Transaction Of The Bag": Exit Sub
    End If
    
   '**************************************** urmi **************************************
    Dim wBFrLoc As String, wBToLoc As String
   ' If (ms_TcTyp = "DT") = True Then
      If gs_ValidPrdLocSeq <> "" Then
        '****** Sachin 3.02 21-12-07 - Current Partition
        wBFrLoc = moCn.GetFldVal("Select BLoc from Bag where " + _
                                 "BCoCd='" + ADC("TCoCd") + "' and " + _
                                 "BYy='" + GrdTxndA.Value(GrdTxndA.Row, "TdBYy") + "' and " + _
                                 "BChr='" + GrdTxndA.Value(GrdTxndA.Row, "TdBChr") + "' and " + _
                                 "BNo= " + CStr(GrdTxndA.Value(GrdTxndA.Row, "TdBNo")) + " and BPrtKey='" + ctCurrPrtn + "'")
        Cancel = IIF((InStr(1, "," + UCase(gs_ValidPrdLocSeq) + ",", "," + UCase(wBFrLoc) + ",") <= 0), True, False)
      Else
        Cancel = False
      End If
      If Cancel = True Then ErrMsg = "Cannot Edit record As Bag Loc. is Not Accessible to the User": Exit Sub
   ' End If
   '**************************************** urmi **************************************

    Select Case ColName
    Case Is = UCase("TdFrRmLoc")
      ' SRI REMOVED cos its not enough time to correct this bug in 5.4.0 release
'      If moCn.GetFldVal("Select hStkRtNoEdit from Head where hcocd='" + gs_CoCd + "'") = "Y" And .Value(RowNum, "TdRmStkRt") <> 0 Then _
 '     ErrMsg = "Cannot Change From Location If StockRate <> 0": Cancel = True
      
      If .Value(RowNum, "TdFrRmLoc") = "" Then
        If RowNum = 1 Then .Value(RowNum, "TdFrRmLoc") = ADC("TFrRmLoc")
        If RowNum > 1 Then .Value(RowNum, "TdFrRmLoc") = .Value(RowNum - 1, "TdFrRmLoc")
      End If
      Call HlpList.Loc(ADC("TCoCd"), , "'P'")
    Case Is = UCase("wTdFrRmIr")
      ' SRI REMOVED cos its not enough time to correct this bug in 5.4.0 release
      'If moCn.GetFldVal("Select hStkRtNoEdit from Head where hcocd='" + gs_CoCd + "'") = "Y" And .Value(RowNum, "TdRmStkRt") <> 0 Then _
      'ErrMsg = "Cannot Change I/R If StockRate <> 0": Cancel = True
      
      If .Value(RowNum, "wTdFrRmIr") = "" Then
        If RowNum = 1 Then .Value(RowNum, "wTdFrRmIr") = "I"
        If RowNum > 1 Then .Value(RowNum, "wTdFrRmIr") = .Value(RowNum - 1, "wTdFrRmIr")
      End If
      Call HlpList.PMCd("IR")
    Case Is = UCase("wTdBrkMsg")
      Call HlpList.PMCd("BRKMSG")
    Case Is = UCase("TdRmCd")
      Dim wOrdSr As String, wOrd() As String
      '****** Sachin 3.02 21-12-07 - Current Partition
      wOrdSr = moCn.GetFldVal("Select BOdTc+','+BOdYy+','+BOdChr+','+Ltrim(str(BOdNo))+" + _
                              "','+LTrim(str(BOdSr)) from Bag where " + _
                              "BCoCd='" + ADC("TCoCd") + "' and " + _
                              "BYy='" + GrdTxndA.Value(GrdTxndA.Row, "TdBYy") + "' and " + _
                              "BChr='" + GrdTxndA.Value(GrdTxndA.Row, "TdBChr") + "' and " + _
                              "BNo= " + CStr(GrdTxndA.Value(GrdTxndA.Row, "TdBNo")) + " and BPrtKey='" + ctCurrPrtn + "'")
      wOrd() = Split(wOrdSr, ",")
      Call HlpList.OrRmCdForIssueRm(ADC("TCoCd"), wOrd(0), wOrd(1), wOrd(2), wOrd(3), wOrd(4), , , GrdTxndA.Value(GrdTxndA.Row, "TdBChr"))
    Case Is = UCase("TdToRmLoc")
      If moCn.GetFldVal("Select hStkRtNoEdit from Head where hcocd='" + gs_CoCd + "'") = "Y" And .Value(RowNum, "TdRmStkRt") <> 0 Then _
      ErrMsg = "Cannot Change To Location If StockRate <> 0": Cancel = True
      
      Call HlpList.Loc(ADC("TCoCd"), , "'P'")
    Case Is = UCase("TdByLoc")
      Call HlpList.Loc(ADC("TCoCd"), "'P'")
    Case Is = UCase("TdByWrk")
     Call HlpList.vPMCd(ADC("TCoCd"), "WORK", , , , , True) ''' Manali Change Reverted  ---- '''', .Value(RowNum, "TdByLoc"))      ' ***** Manali 3.6.0 - Worker Help Sorted On Process Loc
    Case Is = UCase("TdDesc")
      ADC("wTdDesc").SetFocus
    Case Is = UCase("TdRmSz")
      '****** Sachin 3.02 21-12-07 - Current Partition
      wRmCtg = moCn.GetFldVal("Select RmCtg from RmMst where RmCd= '" + .Value(RowNum, "TdRmCd") + "' and RmPrtKey='" + ctCurrPrtn + "' ")
      wRmSCtg = moCn.GetFldVal("Select RmSCtg from RmMst where RmCd= '" + .Value(RowNum, "TdRmCd") + "' and RmPrtKey='" + ctCurrPrtn + "' ")
      
      wGrpBySz = moCn.GetFldVal("Select PDesc225 from Param where PTyp= 'RMSCTG' and " + _
                 "PMCd= '" + wRmCtg + "' and PSCd= '" + wRmSCtg + "' ")
      wGrpByPtr = moCn.GetFldVal("Select PValue3 from Param where PTyp= 'RMSCTG' and " + _
                  "PMCd= '" + wRmCtg + "' and PSCd= '" + wRmSCtg + "' ")
      ' Zubin 212 (Correction - Exit Sub added)
      If wGrpBySz = "N" And wGrpByPtr = "N" Then Cancel = True: ErrMsg = "Size/ Ptr Should Not Be Entered": Exit Sub
      'StkRt-if NoDisplay =yes AND AvgStkRate=yes then it display stock rate for the users who have Seo rights to this entry,
      'otherwise not show stock rate in help. if NoDisplay =No, NO logic applies, always show stock rate in help.
      If moCn.GetFldVal("Select hAvgSRtYn from Head where HCoCd= '" + gs_CoCd + "'") = "Y" And wStkNoDisp = "Y" Then
        mb_SeoUsr = moCn.RecSeek("Select 'x' From UsrAccess where UaUsrCd = '" + gs_UsrCd + "' " + _
                        " and UaMnuCd = '" + ADC("TTc") + "' and UaSeoYn= 'Y' ")
                        
        Call HlpList.BRmSz(ADC("TCoCd"), CStr(GrdTxndA.Value(GrdTxndA.Row, "TdBYy")), _
                        CStr(GrdTxndA.Value(GrdTxndA.Row, "TdBChr")), _
                        CStr(GrdTxndA.Value(GrdTxndA.Row, "TdBNo")), _
                        .Value(RowNum, "TdRmCd"), mb_SeoUsr)
      Else
        Call HlpList.BRmSz(ADC("TCoCd"), CStr(GrdTxndA.Value(GrdTxndA.Row, "TdBYy")), _
                        CStr(GrdTxndA.Value(GrdTxndA.Row, "TdBChr")), _
                        CStr(GrdTxndA.Value(GrdTxndA.Row, "TdBNo")), _
                        .Value(RowNum, "TdRmCd"), True)
      End If
                                
    Case Is = UCase("TdRmSz2")
      wRmSCtg = moCn.GetFldVal("Select RmSCtg from RmMst where RmCd= '" + .Value(RowNum, "TdRmCd") + "' and RmPrtKey='" + ctCurrPrtn + "' ")
      wRmCtg = moCn.GetFldVal("Select RmCtg from RmMst where RmCd= '" + .Value(RowNum, "TdRmCd") + "' and RmPrtKey='" + ctCurrPrtn + "' ")
      Cancel = ColWhenStkSz(2, wRmCtg, wRmSCtg, ErrMsg)
    Case Is = UCase("TdRmSz3")
      wRmCtg = moCn.GetFldVal("Select RmCtg from RmMst where RmCd= '" + .Value(RowNum, "TdRmCd") + "' and RmPrtKey='" + ctCurrPrtn + "' ")
      wRmSCtg = moCn.GetFldVal("Select RmSCtg from RmMst where RmCd= '" + .Value(RowNum, "TdRmCd") + "' and RmPrtKey='" + ctCurrPrtn + "' ")
      Cancel = ColWhenStkSz(3, wRmCtg, wRmSCtg, ErrMsg)
    Case Is = UCase("TdRmStkRt")
      '****** Sachin 3.02 21-12-07 - Current Partition
      wRmCtg = moCn.GetFldVal("Select RmCtg from RmMst where RmCd= '" + .Value(RowNum, "TdRmCd") + "' and RmPrtKey='" + ctCurrPrtn + "' ")
      wRmSCtg = moCn.GetFldVal("Select RmSCtg from RmMst where RmCd= '" + .Value(RowNum, "TdRmCd") + "' and RmPrtKey='" + ctCurrPrtn + "' ")
      wGrpByRt = moCn.GetFldVal("Select PValue1 from Param where PTyp= 'RMSCTG' and " + _
                 "PMCd= '" + wRmCtg + "' and PSCd= '" + wRmSCtg + "' ")
      If wGrpByRt = "N" Then Cancel = True: ErrMsg = "Stock Rate Should Not Be Entered"
      wHAvgStkRt = moCn.GetFldVal("Select hAvgSRtYn from Head where HCoCd= '" + ms_TCoCd + "' and " + _
                      "HCd='" + ctSelfCmCd + "'")
      wStkNoEdit = moCn.GetFldVal("Select hStkRtNoEdit from Head where hcocd='" + gs_CoCd + "'")
                      
      Dim wRmLoc As String, wRmLocTyp As String, wb_RmLocTyp As Boolean
      
            
      If wHAvgStkRt = "Y" And (.IsNew(RowNum) = True Or mf_AvgRtFldChg = True) Then
          wRmLoc = IIF(.Value(RowNum, "wTdFrRmIr") = "I", _
                             .Value(RowNum, "TdFrRmLoc"), _
                             .Value(RowNum, "TdToRmLoc"))
                                                 
          wb_RmLocTyp = moCn.RecSeek("Select LocTyp from Loc where LocCoCd= '" + ADC("TCoCd") + "' " + _
                                    "and LocCd= '" + wRmLoc + "' " + _
                                    "and LocTyp in " + ctAvgLocTypSeq)
          wRmLocTyp = Trim(moCn.GetFldVal("Select LocTyp from Loc where LocCoCd= '" + ADC("TCoCd") + "' " + _
                                    "and LocCd= '" + wRmLoc + "' "))
          If wb_RmLocTyp = True Then
            If wRmLocTyp <> "W" Then
'              If .Value(RowNum, "TdRmStkRt") = 0# Then
                  .Value(RowNum, "TdRmStkRt") = GetAvgStkRt(ADC("TCoCd"), wRmLoc, _
                                                    .Value(RowNum, "TdRmCd"), .Value(RowNum, "TdLotNo"), .Value(RowNum, "TdRmSz"), .Value(RowNum, "TdRmSz2"), .Value(RowNum, "TdRmSz3"))
            Else
 '             If .Value(RowNum, "TdRmStkRt") = 0# Then
              .Value(RowNum, "TdRmStkRt") = GetAvgStkRt(ADC("TCoCd"), wRmLoc, _
                                                .Value(RowNum, "TdRmCd"), .Value(RowNum, "TdLotNo"), .Value(RowNum, "TdRmSz"), .Value(RowNum, "TdRmSz2"), .Value(RowNum, "TdRmSz3"), _
                                                GrdTxndA.Value(GrdTxndA.Row, "TdBYy"), GrdTxndA.Value(GrdTxndA.Row, "TdBChr"), _
                                                GrdTxndA.Value(GrdTxndA.Row, "TdBNo"))
            End If
          End If
      End If
      
      'StkRt-if Stock rate NoEdit is yes and AvgStkRate=yes then can edit the values for X loctyp if External location is credited else existing behaviour
      If moCn.GetFldVal("Select hAvgSRtYn from Head where HCoCd= '" + gs_CoCd + "'") = "Y" And wStkNoEdit = "Y" Then
        wRmLoc = IIF(.Value(RowNum, "wTdFrRmIr") = "I", .Value(RowNum, "TdFrRmLoc"), .Value(RowNum, "TdToRmLoc"))
                    
        wRmLocTyp = Trim(moCn.GetFldVal("Select LocTyp from Loc where LocCoCd= '" + ADC("TCoCd") + "' " + _
                                    "and LocCd= '" + wRmLoc + "' "))
        If wRmLocTyp Like "X*" Then
          Call HlpList.StkRt(ADC("TCoCd"), IIF(.Value(RowNum, "wTdFrRmIr") = "I", .Value(RowNum, "TdFrRmLoc"), IIF(.Value(RowNum, "wTdFrRmIr") = "", ADC("TFrRmLoc"), .Value(RowNum, ("TdToRmLoc")))), .Value(RowNum, "TdRmCd"), .Value(RowNum, "TdRmSz"), .Value(RowNum, "TdRmSz2"), .Value(RowNum, "TdRmSz3"))
        Else
          Cancel = True
        End If
      Else
        mf_AvgRtFldChg = False
        If moCn.GetFldVal("Select hStkRtPassReq from Head where HCoCd= '" + ms_TCoCd + "' and " + _
                      "HCd='" + ctSelfCmCd + "'") = "Y" And _
                      ((.Value(RowNum, "wTdFrRmIR") = "I" And .Value(RowNum, "TdFrRmLoc") <> "W") Or _
                      (.Value(RowNum, "wTdFrRmIR") = "R" And .Value(RowNum, "TdToRmLoc") <> "W")) Then
          If ms_PwdFlag = False Then
            Cancel = True
            ErrMsg = "SEO Password Is Require To Edit Stock Rate."
          Else
            If wHAvgStkRt = "N" Then
              If (.Value(RowNum, "TdFrRmDc") = "C" And .Value(RowNum, "TdFrRmLoc") = "W") Or (.Value(RowNum, "TdFrRmDc") = "D" And .Value(RowNum, "TdToRmLoc") = "W") Then
                Call HlpList.BagStkRt(ADC("TCoCd"), GrdTxndA.Value(GrdTxndA.Row, "TdBYy"), GrdTxndA.Value(GrdTxndA.Row, "TdBChr"), GrdTxndA.Value(GrdTxndA.Row, "TdBNo"), .Value(RowNum, "TdRmCd"), .Value(RowNum, "TdRmSz"), .Value(RowNum, "TdRmSz2"), .Value(RowNum, "TdRmSz3"))
              Else
                Call HlpList.StkRt(ADC("TCoCd"), IIF(.Value(RowNum, "wTdFrRmIr") = "I", .Value(RowNum, "TdFrRmLoc"), IIF(.Value(RowNum, "wTdFrRmIr") = "", ADC("TFrRmLoc"), .Value(RowNum, ("TdToRmLoc")))), .Value(RowNum, "TdRmCd"), .Value(RowNum, "TdRmSz"), .Value(RowNum, "TdRmSz2"), .Value(RowNum, "TdRmSz3"))
              End If
            Else
              If (.Value(RowNum, "TdFrRmDc") = "C" And .Value(RowNum, "TdFrRmLoc") = "W") Or (.Value(RowNum, "TdFrRmDc") = "D" And .Value(RowNum, "TdToRmLoc") = "W") Then
                Call HlpList.BagStkRt(ADC("TCoCd"), GrdTxndA.Value(GrdTxndA.Row, "TdBYy"), GrdTxndA.Value(GrdTxndA.Row, "TdBChr"), GrdTxndA.Value(GrdTxndA.Row, "TdBNo"), .Value(RowNum, "TdRmCd"), .Value(RowNum, "TdRmSz"), .Value(RowNum, "TdRmSz2"), .Value(RowNum, "TdRmSz3"))
              End If
            End If
          End If
        Else
          If wHAvgStkRt = "N" Then
            If (.Value(RowNum, "TdFrRmDc") = "C" And .Value(RowNum, "TdFrRmLoc") = "W") Or (.Value(RowNum, "TdFrRmDc") = "D" And .Value(RowNum, "TdToRmLoc") = "W") Then
              Call HlpList.BagStkRt(ADC("TCoCd"), GrdTxndA.Value(GrdTxndA.Row, "TdBYy"), GrdTxndA.Value(GrdTxndA.Row, "TdBChr"), GrdTxndA.Value(GrdTxndA.Row, "TdBNo"), .Value(RowNum, "TdRmCd"), .Value(RowNum, "TdRmSz"), .Value(RowNum, "TdRmSz2"), .Value(RowNum, "TdRmSz3"))
            Else
              Call HlpList.StkRt(ADC("TCoCd"), IIF(.Value(RowNum, "wTdFrRmIr") = "I", .Value(RowNum, "TdFrRmLoc"), IIF(.Value(RowNum, "wTdFrRmIr") = "", ADC("TFrRmLoc"), .Value(RowNum, ("TdToRmLoc")))), .Value(RowNum, "TdRmCd"), .Value(RowNum, "TdRmSz"), .Value(RowNum, "TdRmSz2"), .Value(RowNum, "TdRmSz3"))
            End If
          Else
            If (.Value(RowNum, "TdFrRmDc") = "C" And .Value(RowNum, "TdFrRmLoc") = "W") Or (.Value(RowNum, "TdFrRmDc") = "D" And .Value(RowNum, "TdToRmLoc") = "W") Then
                Call HlpList.BagStkRt(ADC("TCoCd"), GrdTxndA.Value(GrdTxndA.Row, "TdBYy"), GrdTxndA.Value(GrdTxndA.Row, "TdBChr"), GrdTxndA.Value(GrdTxndA.Row, "TdBNo"), .Value(RowNum, "TdRmCd"), .Value(RowNum, "TdRmSz"), .Value(RowNum, "TdRmSz2"), .Value(RowNum, "TdRmSz3"))
            End If
          End If
      End If
  End If
    
    
      
  Case Is = UCase("TdRmQty")
    If .Value(RowNum, "TdFbRmQty") <> "" Then
      Cancel = True
      ErrMsg = "Rm Qty For Flute Bags Cannot Be Changed Directly. Click On 'Flute Bag Qty/Wt' To Change Qty."
      Exit Sub
    End If
  Case Is = UCase("TdRmWt")
    If .Value(RowNum, "TdFbRmWt") <> "" Then
      Cancel = True
      ErrMsg = "Rm Wt For Flute Bags Cannot Be Changed Directly. Click On 'Flute Bag Qty/Wt' To Change Wt."
      Exit Sub
    End If
' **** Zubin 211 **** '
' Zubin 212
  Case Is = UCase("TdLotNo")
    '****** Sachin 3.02 21-12-07 - Current Partition
    ' **** Manali 3.03 - 03/07/08 - Lot No Open for 'X'
    If .Value(RowNum, "TdRmCd") = "" Or moCn.RecSeek("select 'x' from RmMst where RmZ= 'N' and RmCd= '" + .Value(RowNum, "TdRmCd") + "' and RmCtg IN " + ctLotNoRmCtg + " and RmPrtKey='" + ctCurrPrtn + "'") Then
      If (.Value(RowNum, "TdFrRmDc") = "C" And .Value(RowNum, "TdFrRmLoc") = "W") Or (.Value(RowNum, "TdFrRmDc") = "D" And .Value(RowNum, "TdToRmLoc") = "W") Then
        'StkRt-if NoDisplay =yes and AvgStkRate=yes then it display stock rate for the users who have Seo rights to this entry,
        'otherwise not show stock rate in help. if NoDisplay =No, NO logic applies, always show stock rate in help.
       If moCn.GetFldVal("Select hAvgSRtYn from Head where HCoCd= '" + gs_CoCd + "'") = "Y" And wStkNoDisp = "Y" Then
          mb_SeoUsr = moCn.RecSeek("Select 'x' From UsrAccess where UaUsrCd = '" + gs_UsrCd + "' " + _
                          " and UaMnuCd = '" + ADC("TTc") + "' and UaSeoYn= 'Y' ")
          Call HlpList.BagLotNo(ADC("TCoCd"), GrdTxndA.Value(GrdTxndA.Row, "TdBYy"), GrdTxndA.Value(GrdTxndA.Row, "TdBChr"), GrdTxndA.Value(GrdTxndA.Row, "TdBNo"), .Value(RowNum, "TdRmCd"), .Value(RowNum, "TdRmSz"), .Value(RowNum, "TdRmSz2"), .Value(RowNum, "TdRmSz3"), mb_SeoUsr)
       Else
'SRITODO
          Call HlpList.BagLotNo(ADC("TCoCd"), GrdTxndA.Value(GrdTxndA.Row, "TdBYy"), GrdTxndA.Value(GrdTxndA.Row, "TdBChr"), GrdTxndA.Value(GrdTxndA.Row, "TdBNo"), .Value(RowNum, "TdRmCd"), .Value(RowNum, "TdRmSz"), .Value(RowNum, "TdRmSz2"), .Value(RowNum, "TdRmSz3"), True)
       End If
      Else
'SRITODO
        Call HlpList.LotNo(ADC("TCoCd"), IIF(.Value(RowNum, "wTdFrRmIr") = "I", .Value(RowNum, "TdFrRmLoc"), IIF(.Value(RowNum, "wTdFrRmIr") = "", ADC("TFrRmLoc"), .Value(RowNum, ("TdToRmLoc")))), .Value(RowNum, "TdRmCd"), .Value(RowNum, "TdRmSz"), .Value(RowNum, "TdRmSz2"), .Value(RowNum, "TdRmSz3"))
      End If
    Else
      Cancel = True: ErrMsg = "Lot No. Allowed For 'D', 'C','X','G','P','S','L','M' Only": Exit Sub
    End If
' Zubin 212
    'IG.44 -When From/ To location is B or M then user can enter reason code
    'Otherwise help not allowed & will be read only
    Case Is = UCase("TdBmRsnCd")
    Dim wFrLocTyp, wToLocTyp As String
        wToLocTyp = moCn.GetFldVal("Select LocTyp from Loc where " + _
                                "LocCoCd= '" + ADC("TCoCd") + "' and " + _
                                "LocCd='" + .Value(RowNum, "TdToRmLoc") + "'")
        wFrLocTyp = moCn.GetFldVal("Select LocTyp from Loc where " + _
                                "LocCoCd= '" + ADC("TCoCd") + "' and " + _
                                "LocCd='" + .Value(RowNum, "TdFrRmLoc") + "'")
                                
        If (wToLocTyp = "B" Or wToLocTyp = "M" Or wFrLocTyp = "B" Or wFrLocTyp = "M") Then
          Call HlpList.PMCd("BMRsnCd")
        Else
          Cancel = True: ErrMsg = "B/M Reason Code Not Allowed"
        End If
    End Select
  End With
End Sub
Private Sub GrdTxndB_ReCalc(ByVal RowNum As Integer, ByVal ColName As String, ByVal OldVal As Variant, ByVal SourceName As String, ByVal SourceColName As String)
  '*** wTdFrRmIr on TdFrRmDc              Init
      'If TdFrRmDc= 'C' Then wTdFrRmIr= 'I'
      'If TdFrRmDc= 'D' Then wTdFrRmIr= 'R'
  '*** TdFrRmDc on wTdFrRmIr              Copy,Normal
      'If wTdFrRmIr= 'I' Then TdFrRmDc= 'C'
      'If wTdFrRmIr= 'R' Then TdFrRmDc= 'D'
  '*** TdRmSz on TdRmCd                   Normal
      'TdRmSz= 0
  '*** TdRmStkRt on TdRmCd                Normal
      'TdRmStkRt= 0
  '*** TdRmQty on TdRmCd                  Normal
      'TdRmQty= 0
  '*** TdRmWt on TdRmCd                   Normal
      'TdRmWt= 0
  '*** wTdBrkMsg on TdToRmLoc             Init
      'If LocTyp of TdToRmLoc= 'B' then 'B'
      'If LocTyp of TdToRmLoc= 'M' then 'M'
      'Else ''
  '*** TdToRmLoc on wTdBrkMsg             Normal
      'If wTdBrkMsg= 'B' then LocBrk from Loc(LocCd= TdFrBLoc of 'A' grid)
      'If wTdBrkMsg= 'M' then LocMsg from Loc(LocCd= TdFrBLoc of 'A' grid)
      'Else TToRmLoc
  '*** TdByLoc on TdToRmLoc               Normal
      'If LocTyp of TdToRmLoc is 'B' or 'M' or 'L' then TdFrBLoc of 'A' grid
      'Else ''
  '*** TdByWrk on TdToRmLoc               Normal
      'If LocTyp of TdToRmLoc is 'B' or 'M' or 'L' then TdWrk of 'A' grid
      'Else ''
  '*** TdDesc on wTdDesc                  Normal
      'If TdDesc <> ADC(wTdDesc) then TdDesc = ADC(wTdDesc)
  
  Dim wLocTyp As String
  With GrdTxndB
    Select Case ColName
    Case Is = UCase("wTdFrRmIr")
      If .Mode = fgminit Then
        If .Value(RowNum, "TdFrRmDc") = "C" Then .Value(RowNum, "wTdFrRmIr") = "I"
        If .Value(RowNum, "TdFrRmDc") = "D" Then .Value(RowNum, "wTdFrRmIr") = "R"
      End If
    Case Is = UCase("TdFrRmDc")
      If .Mode = fgmnorm Or .Mode = fgmCopy Then
        If .Value(RowNum, "wTdFrRmIr") = "I" Then .Value(RowNum, "TdFrRmDc") = "C"
        If .Value(RowNum, "wTdFrRmIr") = "R" Then .Value(RowNum, "TdFrRmDc") = "D"
      End If
    Case Is = UCase("TdRmSz")
      If .Mode = fgmnorm Then .Value(RowNum, "TdRmSz") = 0
    Case Is = UCase("TdRmSz2")
      If .Mode = fgmnorm Then .Value(RowNum, "TdRmSz2") = 0
    Case Is = UCase("TdRmSz3")
      If .Mode = fgmnorm Then .Value(RowNum, "TdRmSz3") = 0
    Case Is = UCase("TdRmStkRt")
      If .Mode = fgmnorm Then .Value(RowNum, "TdRmStkRt") = 0
    Case Is = UCase("TdRmQty")
      If .Mode = fgmnorm Then .Value(RowNum, "TdRmQty") = 0
    Case Is = UCase("TdRmWt")
      If .Mode = fgmnorm Then .Value(RowNum, "TdRmWt") = 0
    Case Is = UCase("wTdBrkMsg")
      If .Mode = fgminit Then
        wLocTyp = moCn.GetFldVal("Select LocTyp from Loc where " + _
                                "LocCoCd= '" + ADC("TCoCd") + "' and " + _
                                "LocCd='" + .Value(RowNum, "TdToRmLoc") + "'")
        .Value(RowNum, "wTdBrkMsg") = IIF(wLocTyp = "B", "B", IIF(wLocTyp = "M", "M", ""))
      End If
    Case Is = UCase("TdToRmLoc")
      If .Mode = fgmnorm Then
        If .Value(RowNum, "wTdBrkMsg") = "B" Then
          .Value(RowNum, "TdToRmLoc") = moCn.GetFldVal("Select LocBrk from Loc where " + _
                                        "LocCoCd= '" + ADC("TCoCd") + "' and " + _
                                        "LocCd='" + GrdTxndA.Value(GrdTxndA.Row, "TdFrBLoc") + "'")
          'IG.44 -If user does not have permission to ToLocation then ToLoc will be blank
          If gs_ValidOthLocSeq <> "" Then
            If (InStr(1, "," + UCase(gs_ValidOthLocSeq) + ",", "," + .Value(RowNum, "TdToRmLoc") + ",")) <= 0 Then .Value(RowNum, "TdToRmLoc") = ""
          End If
        ElseIf .Value(RowNum, "wTdBrkMsg") = "M" Then
          .Value(RowNum, "TdToRmLoc") = moCn.GetFldVal("Select LocMsg from Loc where " + _
                                        "LocCoCd= '" + ADC("TCoCd") + "' and " + _
                                        "LocCd='" + GrdTxndA.Value(GrdTxndA.Row, "TdFrBLoc") + "'")
          'IG.44 -If user does not have permission to ToLocation then ToLoc will be blank
          If gs_ValidOthLocSeq <> "" Then
            If (InStr(1, "," + UCase(gs_ValidOthLocSeq) + ",", "," + .Value(RowNum, "TdToRmLoc") + ",")) <= 0 Then .Value(RowNum, "TdToRmLoc") = ""
            
          End If
        Else
          .Value(RowNum, "TdToRmLoc") = ADC("TToRmLoc")
          '.Value(RowNum, "TdToRmLoc") = ""
        End If
      End If
    Case Is = UCase("TdByLoc")
      If .Mode = fgmnorm Then
        wLocTyp = Trim(moCn.GetFldVal("Select LocTyp from Loc where " + _
                                "LocCoCd= '" + ADC("TCoCd") + "' and " + _
                                "LocCd='" + .Value(RowNum, "TdToRmLoc") + "'"))
        .Value(RowNum, "TdByLoc") = IIF(InStr(1, "'L','B','M'", "'" + wLocTyp + "'") > 0, _
                                        GrdTxndA.Value(GrdTxndA.Row, "TdFrBLoc"), "")
      End If
    Case Is = UCase("TdByWrk")
      If .Mode = fgmnorm Then
        wLocTyp = Trim(moCn.GetFldVal("Select LocTyp from Loc where " + _
                                "LocCoCd= '" + ADC("TCoCd") + "' and " + _
                                "LocCd='" + .Value(RowNum, "TdToRmLoc") + "'"))
        '*** If Loss, Brk Or Msg Then Enter Value For ByWrk ***
        .Value(RowNum, "TdByWrk") = IIF(InStr(1, "'L','B','M'", "'" + wLocTyp + "'") > 0, _
                                    GrdTxndA.Value(GrdTxndA.Row, "TdWrk"), "")
        
'        .Value(RowNum, "TdByWrk") = IIf(InStr(1, "'L','B','M'", "'" + wLocTyp + "'") > 0, _
'                                    moCn.GetFldVal("Select BWrk from Bag where " + _
'                                    "BYy='" + GrdTxndA.Value(GrdTxndA.Row, "TdBYy") + "' and " + _
'                                    "BChr='" + GrdTxndA.Value(GrdTxndA.Row, "TdBChr") + "' and " + _
'                                    "BNo= " + CStr(GrdTxndA.Value(GrdTxndA.Row, "TdBNo"))), "")
      End If
    Case Is = UCase("TdDesc")
      If .Mode = fgmnorm Then
        If .Row > 0 Then
          If GrdTxndB.Value(.Row, "TdDesc") <> ADC("wTdDesc") Then
            GrdTxndB.Value(.Row, "TdDesc") = ADC("wTdDesc")
            GrdTxndB.SaveRec .Row
          End If
        End If
      End If
    ' Zubin 212
    Case Is = UCase("TdLotNo")
      If moCn.RecSeek("Select 'x' from Head where HCd= '" + ctSelfCmCd + "' and " + _
                                      "HCoCd= '" + ADC("TCoCd") + "' and HLotNoYN= 'Y'") Then
        If .Row > 0 And .Mode = fgmnorm Then
          '****** Sachin 3.02 21-12-07 - Current Partition
          ' **** Manali 3.03 - 03/07/08 - Lot No Open for 'X'
          If Not moCn.RecSeek("select 'x' from RmMst where RmZ= 'N' and RmCd= '" + .Value(RowNum, "TdRmCd") + "' " + _
                                      "and RmCtg IN " + ctLotNoRmCtg + IIF(gs_Partition = ctCurrPrtn, " and RmPrtKey='" + ctCurrPrtn + "'", "")) Then
            .Value(RowNum, "TdLotNo") = ""
          'Else
            'If .IsNew(RowNum) And .Row > 1 And .Value(RowNum, "TdLotNo") = "" Then
            '  .Value(RowNum, "TdLotNo") = .Value(RowNum - 1, "TdLotNo")
            'End If
          End If
        End If
      'Else
      '  If RowNum > 0 And .Mode = fgmnorm Then .Value(.Row, "TdLotNo") = ""
      End If
    ' Zubin 212
    'IG.44 -When From/To location type is B/M then gets default reason code as '-' else blank
    Case Is = UCase("TdBmRsnCd")
    Dim wFrLocTyp, wToLocTyp As String
      If .Mode = fgmnorm Then
        wToLocTyp = Trim(moCn.GetFldVal("Select LocTyp from Loc where " + _
                                "LocCoCd= '" + ADC("TCoCd") + "' and " + _
                                "LocCd='" + .Value(RowNum, "TdToRmLoc") + "'"))
        wFrLocTyp = Trim(moCn.GetFldVal("Select LocTyp from Loc where " + _
                                "LocCoCd= '" + ADC("TCoCd") + "' and " + _
                                "LocCd='" + .Value(RowNum, "TdFrRmLoc") + "'"))
                                
        If (wToLocTyp = "B" Or wToLocTyp = "M" Or wFrLocTyp = "B" Or wFrLocTyp = "M") Then
          If .Value(RowNum, "TdBmRsnCd") = "" Then .Value(RowNum, "TdBmRsnCd") = "-"
        Else
          .Value(RowNum, "TdBmRsnCd") = ""
        End If
      End If
      
    '****** Sachin - 3.02 - New Fields Introduced ******
    '****** Sachin - 3.02 - Shifted to InitKey
    '    Case Is = UCase("TdRmCtg")
    '        '****** Sachin 3.02 21-12-07 - Current Partition
    '        If .Mode = fgmnorm Or .Mode = fgmCopy Or .Mode = fgminit Then _
    '            .Value(RowNum, "TdRmCtg") = moCn.GetFldVal("Select RmCtg From RmMst Where RmCd='" + .Value(RowNum, "TdRmCd") + "' and RmPrtKey='" + ctCurrPrtn + "'")
    '
    '    Case Is = UCase("TdRmSCtg")
    '        '****** Sachin 3.02 21-12-07 - Current Partition
    '        If .Mode = fgmnorm Or .Mode = fgmCopy Or .Mode = fgminit Then _
    '            .Value(RowNum, "TdRmSCtg") = moCn.GetFldVal("Select RmSCtg From RmMst Where RmCd='" + .Value(RowNum, "TdRmCd") + "' and RmPrtKey='" + ctCurrPrtn + "'")
    '
    '    Case Is = UCase("TdFrRmLocTyp")
    '        If .Mode = fgmnorm Or .Mode = fgmCopy Or .Mode = fgminit Then _
    '            .Value(RowNum, "TdFrRmLocTyp") = moCn.GetFldVal("Select LocTyp From Loc Where LocCoCd='" + ADC("TCoCd") + "' And LocCd='" + .Value(RowNum, "TdFrRmLoc") + "'")
    '
    '    Case Is = UCase("TdToRmLocTyp")
    '        If .Mode = fgmnorm Or .Mode = fgmCopy Or .Mode = fgminit Then _
    '            .Value(RowNum, "TdToRmLocTyp") = moCn.GetFldVal("Select LocTyp From Loc Where LocCoCd='" + ADC("TCoCd") + "' And LocCd='" + .Value(RowNum, "TdToRmLoc") + "'")
    '****** Sachin - 3.02 - New Fields Introduced ******
    
    End Select
  End With
End Sub
Private Sub GrdTxndB_ColChange(ByVal RowNum As Integer, ByVal ColName As String, ByVal NewValue As Variant, Cancel As Boolean, ErrMsg As String)
  '*** TdFrRmLoc          valid LocCd from Loc(LocTyp<> 'P')
  '*** wTdFrRmIr          valid PMCd('IR')
  '*** TdRmCd             valid OrRmCd From OrdRm for the Order to which the Bag belongs
  '*** wTdBrkMsg          valid PMCd('BRKMSG') or ''
  '*** TdToRmLoc          valid LocCd from Loc(LocTyp<> 'P')
  '*** TdByLoc            valid LocCd from Loc(LocTyp= 'P') or ''
  '*** TdByWrk            valid PMCd('WORK') or ''
  
  With GrdTxndB
    Select Case ColName
    Dim wLocTyp As String
    Case Is = UCase("TdFrRmLoc")
     '****** Sachin 2.13.0 [16-08-2006] - ValidYN added below
      Cancel = (Not moCn.RecSeek("Select LocCd From Loc Where LocTyp<>'P' and LocTyp<>'R' and " + _
                          "LocTyp<>'Y' and LocTyp<>'XT' and " + _
                           "LocCoCd= '" + ADC("TCoCd") + "' and " + _
                           "LocCd = '" + NewValue + "' And LocValidYN='Y'"))
      If Cancel = True Then ErrMsg = "Invalid Raw Material From Location": Exit Sub
      
      
        If ms_TcTyp = "DT" Then
              Cancel = (Not moCn.RecSeek("Select LocCd From Loc Where LocTyp<>'XC' AND  " + _
                           "LocCoCd= '" + ADC("TCoCd") + "' and " + _
                           "LocCd = '" + NewValue + "' And LocValidYN='Y'"))
              If Cancel = True Then ErrMsg = "Invalid Raw Material From Location": Exit Sub


        If NewValue <> "W" And gs_ValidOthLocSeq <> "" Then
           Cancel = IIF((InStr(1, "," + UCase(gs_ValidOthLocSeq) + ",", "," + UCase(NewValue) + ",") <= 0), True, False)
           
        Else
          Cancel = False
        End If
        If Cancel = True Then ErrMsg = "RM From Location is Not Accessible to the User": Exit Sub
      End If

        
      If ms_TcTyp = "BS" Then
        If UCase(NewValue) <> "W" And UCase(NewValue) <> "XC" Then Cancel = True
        If Cancel Then ErrMsg = "Invalid RM From Location ": Exit Sub
      End If
      
    Case Is = UCase("wTdFrRmIr")
      Cancel = Not moCn.RecSeek("Select PMCd From Param Where PTyp= 'IR' And " + _
                           "PMCd = '" + NewValue + "'")
        If Cancel = True Then ErrMsg = "Enter Issue Or Receipt (I/R)": Exit Sub
      
      If ms_TcTyp = "DT" Then
        
        If gs_ValidOthLocSeq <> "" Then
           Cancel = IIF((InStr(1, "," + UCase(gs_ValidOthLocSeq) + ",", "," + _
                        UCase(.Value(RowNum, "TdFrRmLoc")) + ",") <= 0), True, False)
           Cancel = Cancel And (UCase(NewValue) = "I")
        Else
          Cancel = False
        End If
        If Cancel = True Then ErrMsg = "RM From Location is Not Accessible to the User": Exit Sub
      
        If gs_ValidOthLocSeq <> "" Then
           Cancel = IIF((InStr(1, "," + UCase(gs_ValidOthLocSeq) + ",", "," + _
                        UCase(.Value(RowNum, "TdToRmLoc")) + ",") <= 0), True, False)
           Cancel = Cancel And (UCase(NewValue) = "R")
        Else
          Cancel = False
        End If
        If Cancel = True Then ErrMsg = "RM To Location is Not Accessible to the User": Exit Sub
      
      End If
    
    
    Case Is = UCase("TdRmCd")
      '****** Sachin 3.02 21-12-07 - Current Partition
      Cancel = Not moCn.RecSeek("Select OrRmCd From OrdRm join Bag on " + _
                           " OrOdIdNo=BOdIdNo And OrCoCd= BCoCd and OrTc= BOdTc and OrYy= BOdYy and " + _
                           "OrChr= BOdChr and OrNo= BOdNo And OrSr = BOdSr And OrPrtKey=BPrtKey " + _
                           "where BCoCd= '" + ADC("TCoCd") + "' and " + _
                           "BYy= '" + GrdTxndA.Value(GrdTxndA.Row, "TdBYy") + "' and " + _
                           "BChr= '" + GrdTxndA.Value(GrdTxndA.Row, "TdBChr") + "' and " + _
                           "BNo= " + CStr(GrdTxndA.Value(GrdTxndA.Row, "TdBNo")) + " and " + _
                           "OrRmCd= '" + NewValue + "' and OrPrtKey='" + ctCurrPrtn + "'")
                           '(index=Bag)
      If Cancel = True Then ErrMsg = "Invalid Raw Material Code": Exit Sub
    
      Cancel = moCn.RecSeek("select 'a' from " + _
                          " RmMst Join Param ON PTyp= 'BCHR' and PMCd = '" + GrdTxndA.Value(GrdTxndA.Row, "TdBChr") + "' and PSCd= '' and PValue in ('F') " + _
                          "  where RmCd='" + NewValue + "' and " + _
                          " NOT ( " + _
                        "(',' + PVALUE1 + ',') LIKE (CASE WHEN PVALUE1='' THEN '%' ELSE '%,'+RmCtg+',%' END) and " + _
                      "(',' + PVALUE2 + ',') LIKE (CASE WHEN PVALUE2='' THEN '%' ELSE '%,'+RmSCtg+',%' END) and " + _
                    "((Pdesc225='' and PValue3='') " + _
                    "or (exists (Select 'a' from Bag " + _
" join OrdRm on Bag.BOdIdNo = OrdRm.OrOdIdNo " + _
" where Bag.BCoCd='" + ADC("TCoCd") + "' AND Bag.BYy = '" + GrdTxndA.Value(GrdTxndA.Row, "TdByy") + "' and Bag.BChr = '" + GrdTxndA.Value(GrdTxndA.Row, "tDbcHR") + "' AND Bag.BNo =" + CStr(GrdTxndA.Value(GrdTxndA.Row, "TdBNo")) + " AND" + _
" '" + NewValue + "'=OrRmCd and " + _
" (PDesc225='' or PDesc225=OrSubShp) " + _
" and (PValue3='' or (PValue3='WS' and OrWsQty>0) or (PValue3='GS' and OrHsQty>0))))))")

'**** urmi Ech AvgStkRt ******'
    Case Is = UCase("TdRmSz")
    ' Manoj 2.9.0. Patch
    mf_AvgRtFldChg = True
    ' Manoj 2.9.0. Patch
    If (moCn.GetFldVal("Select hAvgSRtYn from Head where HCoCd= '" + ms_TCoCd + "' and " + _
                      "HCd='" + ctSelfCmCd + "'") = "Y") Or _
        (moCn.GetFldVal("Select hSizeStkRt from Head where HCoCd= '" + ms_TCoCd + "' and " + _
                      "HCd='" + ctSelfCmCd + "'") = "Y") Then
        
        '******* urmi to chk for 'Trial order Chrs' ******************
        '****** Sachin 3.02 21-12-07 - Current Partition
        If moCn.GetFldVal("Select vPValue3 from vParam where " + _
                    "vPCoCd= '" + ms_TCoCd + "' and vPTyp= 'CHR' and " + _
                    "vPmCd= (select BOdTc from Bag where BCoCd= '" + ms_TCoCd + "' and " + _
                        "BYy= '" + GrdTxndA.Value(GrdTxndA.Row, "TdBYy") + "' and " + _
                        "BChr= '" + GrdTxndA.Value(GrdTxndA.Row, "TdBChr") + "' and " + _
                        "BNo= " + CStr(GrdTxndA.Value(GrdTxndA.Row, "TdBNo")) + " and BPrtKey='" + ctCurrPrtn + "') and " + _
                    "vPSCd= (select BOdChr from Bag where BCoCd= '" + ms_TCoCd + "' and " + _
                        "BYy= '" + GrdTxndA.Value(GrdTxndA.Row, "TdBYy") + "' and " + _
                        "BChr= '" + GrdTxndA.Value(GrdTxndA.Row, "TdBChr") + "' and " + _
                        "BNo= " + CStr(GrdTxndA.Value(GrdTxndA.Row, "TdBNo")) + " and BPrtKey='" + ctCurrPrtn + "') ") = "N" Then
    '******* urmi to chk for 'Trial order Chrs' ******************
     If moCn.GetFldVal("Select PDesc225 from Param join RmMst on PTyp='RMSCTG' and PMCd=RmCtg and PSCd=RmSCtg and RmPrtKey='" + ctCurrPrtn + "' " + _
                               " where RmCd='" + CStr(GrdTxndB.Value(RowNum, "TdRmCd")) + "' ") = "Y" Then
          
          Dim wb_OrRmSz As Boolean, wb_RSz As Boolean
          '****** Sachin 3.02 21-12-07 - Current Partition
          wb_OrRmSz = moCn.RecSeek("Select OrLn1 From OrdRm join Bag on " + _
                       " OrOdIdNo=BOdIdNo And OrCoCd= BCoCd and OrTc= BOdTc and OrYy= BOdYy and " + _
                       "OrChr= BOdChr and OrNo= BOdNo And OrSr = BOdSr and OrPrtKey=BPrtKey " + _
                       "where BCoCd= '" + ADC("TCoCd") + "' and " + _
                       "BYy= '" + GrdTxndA.Value(GrdTxndA.Row, "TdBYy") + "' and " + _
                       "BChr= '" + GrdTxndA.Value(GrdTxndA.Row, "TdBChr") + "' and " + _
                       "BNo= " + CStr(GrdTxndA.Value(GrdTxndA.Row, "TdBNo")) + " and " + _
                       "OrRmCd= '" + CStr(GrdTxndB.Value(RowNum, "TdRmCd")) + "' and " + _
                       "OrLn1= " + CStr(NewValue) + " and OrPrtKey='" + ctCurrPrtn + "'")
          If wb_OrRmSz = True Then Cancel = False: Exit Sub
          
          '****** Sachin 3.02 21-12-07 - Current Partition
          wb_RSz = moCn.RecSeek("Select RrRngSz From OrdRm join Bag on " + _
                    "OrOdIdNo=BOdIdNo And OrCoCd= BCoCd and OrTc= BOdTc and OrYy= BOdYy and " + _
                    "OrChr= BOdChr and OrNo= BOdNo And OrSr = BOdSr and OrPrtKey=BPrtKey " + _
                    "and OrRmCd= '" + GrdTxndB.Value(RowNum, "TdRMCd") + "' " + _
                    "join RmRt on RrTcTyp= 'Rsz' and RrCmCd= '" + ctSelfCmCd + "' and RrCtg= OrRmCtg and " + _
                    "RrSCtg= OrRmSCtg and RrPrtKey=OrPrtKey " + _
                    "where BCoCd= '" + ADC("TCoCd") + "' and " + _
                    "BYy= '" + GrdTxndA.Value(GrdTxndA.Row, "TdBYy") + "' and " + _
                    "BChr= '" + GrdTxndA.Value(GrdTxndA.Row, "TdBChr") + "' and " + _
                    "BNo= " + CStr(GrdTxndA.Value(GrdTxndA.Row, "TdBNo")) + " and " + _
                    "OrLn1 >= RrFrLn and OrLn1 <= RrToLn and OrPrtKey='" + ctCurrPrtn + "' and " + _
                    CStr(NewValue) + " >= RrFrLn and " + CStr(NewValue) + " <= RrToLn")
          If wb_RSz = True Then Cancel = False: Exit Sub
          mf_AvgRtFldChg = False
          Cancel = True: ErrMsg = "Invalid Raw Material Size": Exit Sub
     Else
     If moCn.GetFldVal("Select PValue3 from Param join RmMst on PTyp='RMSCTG' and PMCd=RmCtg  and PSCd=RmSCtg " + _
                               " where RmCd='" + CStr(GrdTxndB.Value(RowNum, "TdRmCd")) + "' and RmPrtKey='" + ctCurrPrtn + "' ") = "Y" Then
        Dim ws_OrRmPtr As String
        Dim wb_OrRmPtr As Boolean, wb_RPtr As Boolean
          ws_OrRmPtr = IIF(moCn.GetFldVal("Select PValue2 from Param join RmMst on PTyp='RMSCTG' and PMCd=RmCtg  and PSCd=RmSCtg " + _
                               " where RmCd='" + CStr(GrdTxndB.Value(RowNum, "TdRmCd")) + "' And RmPrtKey='" + ctCurrPrtn + "' ") = "Y", "OrRmPtr", "Case when OrPrdQty=0 then 0 else OrPrdWt/OrPrdQty end")
          wb_OrRmPtr = moCn.RecSeek("Select OrRmPtr From OrdRm join Bag on " + _
                       "OrOdIdNo=BOdIdNo And OrCoCd= BCoCd and OrTc= BOdTc and OrYy= BOdYy and " + _
                       "OrChr= BOdChr and OrNo= BOdNo And OrSr = BOdSr And OrPrtKey=BPrtKey " + _
                       "where BCoCd= '" + ADC("TCoCd") + "' and " + _
                       "BYy= '" + GrdTxndA.Value(GrdTxndA.Row, "TdBYy") + "' and " + _
                       "BChr= '" + GrdTxndA.Value(GrdTxndA.Row, "TdBChr") + "' and " + _
                       "BNo= " + CStr(GrdTxndA.Value(GrdTxndA.Row, "TdBNo")) + " and " + _
                       "OrRmCd= '" + CStr(GrdTxndB.Value(RowNum, "TdRmCd")) + "' and " + _
                       "OrPrtKey= '" + ctCurrPrtn + "' and " + _
                       ws_OrRmPtr + "= " + CStr(NewValue))
          If wb_OrRmPtr = True Then Cancel = False: Exit Sub
          wb_RPtr = moCn.RecSeek("Select RrRngSz From OrdRm join Bag on " + _
                    "OrOdIdNo=BOdIdNo And OrCoCd= BCoCd and OrTc= BOdTc and OrYy=BOdYy and " + _
                    "OrChr= BOdChr and OrNo= BOdNo And OrSr = BOdSr and OrPrtKey=BPrtKey " + _
                    "and OrRmCd= '" + GrdTxndB.Value(RowNum, "TdRMCd") + "' " + _
                    "join RmRt on RrTcTyp= 'Rsz' and RrCmCd= '" + ctSelfCmCd + "' and RrCtg= OrRmCtg and " + _
                    "RrSCtg= OrRmSCtg and RrPrtKey=OrPrtKey " + _
                    "where BCoCd= '" + ADC("TCoCd") + "' and " + _
                    "BYy= '" + GrdTxndA.Value(GrdTxndA.Row, "TdBYy") + "' and " + _
                    "BChr= '" + GrdTxndA.Value(GrdTxndA.Row, "TdBChr") + "' and " + _
                    "BNo= " + CStr(GrdTxndA.Value(GrdTxndA.Row, "TdBNo")) + " and " + _
                    "BPrtKey= '" + ctCurrPrtn + "' and " + _
                    ws_OrRmPtr + " >= RrFrLn and " + ws_OrRmPtr + " <= RrToLn and " + _
                    CStr(NewValue) + " >= RrFrLn and " + CStr(NewValue) + " <= RrToLn")
          If wb_RPtr = True Then Cancel = False: Exit Sub
          mf_AvgRtFldChg = True
          Cancel = True: ErrMsg = "Invalid Raw Material Pointer": Exit Sub
        End If
       End If
      End If
    End If
    Case Is = UCase("wTdBrkMsg")
      Cancel = (NewValue <> "") And (Not moCn.RecSeek("Select PMCd From Param Where " + _
                                     "PTyp= 'BRKMSG' And PMCd = '" + NewValue + "'"))
      ErrMsg = "Enter Broken Or Missing (B/M)": Exit Sub
    
   Case Is = UCase("TdToRmLoc")
      Cancel = (Not moCn.RecSeek("Select LocCd From Loc Where LocTyp<>'P' and LocTyp<>'R' and " + _
                          "LocTyp<>'Y' and LocTyp<>'XT' and " + _
                           "LocCoCd= '" + ADC("TCoCd") + "' and " + _
                           "LocCd = '" + NewValue + "' And LocValidYN='Y'"))
      If Cancel = True Then ErrMsg = "Invalid Raw Material To Location": Exit Sub
      
      If ms_TcTyp = "DT" Then
          Cancel = (Not moCn.RecSeek("Select LocCd From Loc Where LocTyp<>'XC'  AND " + _
                           "LocCoCd= '" + ADC("TCoCd") + "' and " + _
                           "LocCd = '" + NewValue + "' And LocValidYN='Y'"))
        If Cancel = True Then ErrMsg = "Invalid Raw Material To Location": Exit Sub
        
        If NewValue <> "W" And gs_ValidOthLocSeq <> "" Then
           Cancel = IIF((InStr(1, "," + UCase(gs_ValidOthLocSeq) + ",", "," + UCase(NewValue) + ",") <= 0), True, False)
        Else
          Cancel = False
        End If
        ErrMsg = "RM To Location is Not Accessible to the User": Exit Sub
      End If
      
      If ms_TcTyp = "BS" Then
        If UCase(NewValue) <> "W" And UCase(NewValue) <> "XC" Then Cancel = True
        If Cancel Then ErrMsg = "Invalid RM To Location ": Exit Sub
      End If
    Case Is = UCase("TdByLoc")
      Cancel = (NewValue <> "") And (Not moCn.RecSeek("Select LocCd From Loc Where " + _
                                     "LocCoCd= '" + ADC("TCoCd") + "' and " + _
                                     "LocTyp= 'P' and LocCd = '" + NewValue + "'"))
      ErrMsg = "Invalid By Location": Exit Sub
    
    Case Is = UCase("TdByWrk")
      Cancel = (NewValue <> "") And (Not moCn.RecSeek("Select vPMCd From vParam Where " + _
                                     "vPCoCd = '" + ADC("TCoCd") + "' and vPTyp= 'WORK' And " + _
                                     "vPMCd = '" + NewValue + "'"))
      ErrMsg = "Invalid By Worker": Exit Sub
    Case Is = UCase("TdLotNo")
      mf_AvgRtFldChg = True
      If (.Value(RowNum, "TdFrRmDc") = "C" And .Value(RowNum, "TdFrRmLoc") = "W") Or (.Value(RowNum, "TdFrRmDc") = "D" And .Value(RowNum, "TdToRmLoc") = "W") Then
        Cancel = NewValue <> "" And Not moCn.RecSeek("Select 'x' From Txnd Where TdRmCd <> '' and " + _
                "TdCoCd= '" + ADC("TCoCd") + "' and TdBYy= '" + CStr(GrdTxndA.Value(GrdTxndA.Row, "TdBYy")) + "' and " + _
                "TdBChr= '" + CStr(GrdTxndA.Value(GrdTxndA.Row, "TdBChr")) + "' and " + _
                "TdBNo= '" + CStr(GrdTxndA.Value(GrdTxndA.Row, "TdBNo")) + "' and TdPrtKey='" + ctCurrPrtn + "' " + _
                IIF(.Value(RowNum, "TdRmCd") <> "", " and TdRmCd= '" + CStr(.Value(RowNum, "TdRmCd")) + "' ", "") + _
                IIF(.Value(RowNum, "TdRmSz") <> 0, " and TdRmSz= '" + CStr(.Value(RowNum, "TdRmSz")) + "' ", "") + _
                IIF(.Value(RowNum, "TdRmSz2") <> 0, " and TdRmSz2= '" + CStr(.Value(RowNum, "TdRmSz2")) + "' ", "") + _
                IIF(.Value(RowNum, "TdRmSz3") <> 0, " and TdRmSz3= '" + CStr(.Value(RowNum, "TdRmSz3")) + "' ", "") + _
                "and TdLotNo = '" + NewValue + "' " + _
                "Group By TdBYy, TdBChr, TdBNo" + _
                IIF(.Value(RowNum, "TdRmCd") <> "", ", TdRmCd", "") + _
                IIF(.Value(RowNum, "TdRmSz") <> 0, ", TdRmSz", "") + IIF(.Value(RowNum, "TdRmSz2") <> 0, ", TdRmSz2", "") + IIF(.Value(RowNum, "TdRmSz3") <> 0, ", TdRmSz3", "") + ", TdLotNo, TdRmStkRt ")
      Else
        Cancel = NewValue <> "" And Not moCn.RecSeek("Select 'x' from SYyMm where SCoCd= '" + ADC("TCoCd") + "' and " + _
                                    "SLotNo= '" + NewValue + "' And SPrtKey='" + ctCurrPrtn + "'")
      End If
      ErrMsg = "Invalid Lot No.": Exit Sub
    Case Is = UCase("TdBmRsnCd")
        Cancel = Not moCn.RecSeek("Select PMCd From Param Where PTyp= 'BMRsnCd' And " + _
                           "PMCd = '" + NewValue + "'")
        If Cancel = True Then ErrMsg = "Enter Valid Breakage/Missing Reason Code": Exit Sub
    End Select
    Select Case ColName
      Case Is = UCase$("TdRmCd"), UCase$("wTdFrRmIr")
        mf_AvgRtFldChg = True
      Case Is = UCase$("TdFrRmLoc")
        If UCase$(GrdTxndB.Value(RowNum, "wTdFrRmIr")) = "I" Then mf_AvgRtFldChg = True
      Case Is = UCase$("TdToRmLoc")
        If UCase$(GrdTxndB.Value(RowNum, "wTdFrRmIr")) = "R" Then mf_AvgRtFldChg = True
    End Select
    ' Manoj 2.9.0 Patch

  
  End With
End Sub

'@@@
'Private Sub GrdTxndB_OnDelete(ByVal RowNum As Integer, Cancel As Boolean, ErrMsg As String)
'  If ms_TcTyp = "BS" And GrdTxndA.Row <> GrdTxndA.Rows - 1 Then
'    Cancel = True: ErrMsg = "Delete The Last Entered Records Of Grid 'A' First": Exit Sub
'  End If
'End Sub
'@@@

Private Sub GrdTxndB_Valid()
  '*** Update the field TdBGrWt of grid 'A' as TdBGrWt = BGrWt
      'from Bag(BYy= TdByy, BChr= TdBChr, BNo= TdBNo of 'A' grid)
      'Update GrsWt Only If Last Transaction Of Bag
  
  With GrdTxndA
    Dim wTxndARow As Single
    'wTxndARow = IIf(.Row > 0, .Row, 1)
    
    wTxndARow = IIF(.Row > 0, .Row, 0)
    If wTxndARow = 0 Then Exit Sub
    
    '****** Sachin 3.02 21-12-07 - Current Partition
    If moCn.RecSeek("Select TdTc From Txnd Where TdCoCd='" + ADC("TCoCd") + "' and " + _
               "TdTc='" + ADC("TTc") + "' and TdYy='" + ADC("TYy") + "' and " + _
               "TdChr='" + ADC("TChr") + "' and TdNo=" + CStr(ADC("TNo")) + " and " + _
               "TdSr=" + CStr(.Value(wTxndARow, "TdSr")) + " and " + _
               "TdSrNo=0 and TdPrtKey='" + ctCurrPrtn + "'") And _
       moCn.RecSeek("Select BYy from Bag where BCoCd='" + ADC("TCoCd") + "' and " + _
               "BYy='" + .Value(wTxndARow, "TdBYy") + "' and " + _
               "BChr='" + .Value(wTxndARow, "TdBChr") + "' and " + _
               "BNo= " + CStr(.Value(wTxndARow, "TdBNo")) + " and BPrtKey='" + ctCurrPrtn + "' and " + _
               "BLstYy='" + ADC("TYy") + "' and " + _
               "BLstKey=" + CStr(.Value(wTxndARow, "TdKey"))) Then
               
        Dim wRBYN As String
        wRBYN = moCn.GetFldVal("Select BRepYN from Bag where BCoCd='" + ADC("TCoCd") + "' and BYy='" + .Value(wTxndARow, "TdBYy") + "' and " + _
                                     "BChr='" + .Value(wTxndARow, "TdBChr") + "' and BNo= " + CStr(.Value(wTxndARow, "TdBNo")) + IIF(gs_Partition = ctCurrPrtn, " And BPrtKey='" + ctCurrPrtn + "'", ""))
        If wRBYN = "N" Then
          .Value(wTxndARow, "TdBGrWt") = moCn.GetFldVal("Select BGrWt from Bag where " + _
                                     "BCoCd='" + ADC("TCoCd") + "' and " + _
                                     "BYy='" + .Value(wTxndARow, "TdBYy") + "' and " + _
                                     "BChr='" + .Value(wTxndARow, "TdBChr") + "' and " + _
                                     "BNo= " + CStr(.Value(wTxndARow, "TdBNo")) + _
                                     "and BPrtKey= '" + ctCurrPrtn + "'")
        Else
          .Value(wTxndARow, "TdBGrWt") = moCn.GetFldVal("Select BRepGrWt from Bag where " + _
                                     "BCoCd='" + ADC("TCoCd") + "' and " + _
                                     "BYy='" + .Value(wTxndARow, "TdBYy") + "' and " + _
                                     "BChr='" + .Value(wTxndARow, "TdBChr") + "' and " + _
                                     "BNo= " + CStr(.Value(wTxndARow, "TdBNo")) + _
                                     "and BPrtKey= '" + ctCurrPrtn + "'")
        End If
      .SaveRec (wTxndARow)
    End If
  End With
  
  ' **** Zubin 211 **** '
  If CmdQW.Visible = True Then CmdQW.Enabled = False
  ' **** Zubin 211 **** '
  If CmdLN.Visible = True Then CmdLN.Enabled = False    '******Bhavna
End Sub
Private Sub GrdTxndB_InitKey(ByVal RowNum As Integer)
  '*** Initialise Key Fields and all the fields which are not included in the grid
  Dim ws_TdYyKey As String
  Dim wf_TdKeyLen As Single
  
  With GrdTxndB
    Dim wTxndARow As Single
    wTxndARow = IIF(GrdTxndA.Row > 0, GrdTxndA.Row, 1)
    '****** Sachin 3.02 21-12-07 - Corresponding Partition
    If Not moCn.RecSeek("Select TdTc From Txnd Where TdCoCd='" + ADC("TCoCd") + "' and " + _
                   "TdTc='" + ADC("TTc") + "' and TdYy='" + ADC("TYy") + "' and " + _
                   "TdChr='" + ADC("TChr") + "' and TdNo=" + CStr(ADC("TNo")) + " and " + _
                   "TdSr=" + CStr(GrdTxndA.Value(GrdTxndA.Row, "TdSr")) + " and " + _
                   "TdSrNo=" + CStr(.Value(RowNum, "TdSrNo")) + IIF(gs_Partition = ctCurrPrtn, " and TdPrtKey='" + ctCurrPrtn + "'", "")) Then
    
      .Store "TdCoCd", ADC("TCoCd")
      .Store "TdTc", ADC("TTc")
      .Store "TdYy", ADC("TYy")
      .Store "TdChr", ADC("TChr")
      .Store "TdNo", ADC("TNo")
      .Store "TdSr", GrdTxndA.Value(wTxndARow, "TdSr")
    
      .Store "TdRefYy", ""
      .Store "TdRefKey", 0
      
      .Store "TdPrevYy", GrdTxndA.Value(wTxndARow, "TdPrevYy")
      .Store "TdPrevKey", GrdTxndA.Value(wTxndARow, "TdPrevKey")
      
      .Store "TdBYy", GrdTxndA.Value(wTxndARow, "TdBYy")
      .Store "TdBChr", GrdTxndA.Value(wTxndARow, "TdBChr")
      .Store "TdBNo", GrdTxndA.Value(wTxndARow, "TdBNo")
      .Store "TdRjDc", ""
      '.Store "TdAck", ""
      .Store "TdPrdYN", ""
      
      .Store "TdBDc", ""
      
      .Store "TdBQty", 0
      .Store "TdPtQty", 0
      .Store "TdBGrWt", 0
      .Store "TdFrBLoc", ""
      .Store "TdToBLoc", ""
      .Store "TdRjQty", 0
      .Store "TdRjPtQty", 0
      .Store "TdRjCd", ""
      .Store "TdWrk", ""
      .Store "TdKey", GrdTxndA.Value(wTxndARow, "TdKey")
      
      wf_TdKeyLen = Len(Trim(Str(GrdTxndA.Value(wTxndARow, "TdKey"))))
      ws_TdYyKey = ADC("TYy") + Space(8 - wf_TdKeyLen) + Trim(Str(GrdTxndA.Value(wTxndARow, "TdKey")))
      .Store "TdYyKey", ws_TdYyKey
      
      .Store "TdPurRt", 0
      .Store "TdPurAmt", 0
      .Store "TdWaxWt", 0#
      '.Store "TdPcPerCt", 0
    End If
    
    If GrdTxndB.Value(RowNum, "TdFrRmDc") = "C" Then
      .Store "TdToRmDc", "D"
    ElseIf GrdTxndB.Value(RowNum, "TdFrRmDc") = "D" Then
      .Store "TdToRmDc", "C"
    Else
      .Store "TdToRmDc", ""
    End If
    
    .Store "TdPcPerCt", 0
    
    '****** Sachin 3.02 21-12-07 - Current Partition
    .Store "TdRmCtg", moCn.GetFldVal("Select RmCtg From RmMst Where RmCd='" + .Value(RowNum, "TdRmCd") + "'" + IIF(gs_Partition = ctCurrPrtn, " and RmPrtKey='" + ctCurrPrtn + "'", ""))
    .Store "TdRmSCtg", moCn.GetFldVal("Select RmSCtg From RmMst Where RmCd='" + .Value(RowNum, "TdRmCd") + "'" + IIF(gs_Partition = ctCurrPrtn, " and RmPrtKey='" + ctCurrPrtn + "'", ""))
    .Store "TdFrRmLocTyp", moCn.GetFldVal("Select LocTyp From Loc Where LocCoCd='" + ADC("TCoCd") + "' And LocCd='" + .Value(RowNum, "TdFrRmLoc") + "'")
    .Store "TdToRmLocTyp", moCn.GetFldVal("Select LocTyp From Loc Where LocCoCd='" + ADC("TCoCd") + "' And LocCd='" + .Value(RowNum, "TdToRmLoc") + "'")
    '****** Sachin 3.02 21-12-07 - Current Partition
    
    .Store "TdRtByQW", ""     ' ***** Manali 3.8.0
    
    '****** Sachin 3.02 (14/09/07) Id Fields
    '****** Sachin 3.02 21-12-07 - Current Partition
      .Store "TdTIdNo", moCn.GetFldVal("Select TIdNo From Txn " + _
                        "Where TCoCd='" + ADC("TCoCd") + "' And TTc='" + ADC("TTc") + "' And TYy='" + ADC("TYy") + "' And Tchr='" + ADC("TChr") + "' And TNo=" + CStr(ADC("TNo")) + " And TPrtKey='" + ctCurrPrtn + "'")
      .Store "TdDt", ADC("TDt")
      .Store "TdBIdNo", moCn.GetFldVal("Select BIdNo From Bag " + _
                          " Where BCoCd='" + ADC("TCoCd") + "' And BYy='" + GrdTxndA.Value(wTxndARow, "TdBYy") + "' And BChr='" + GrdTxndA.Value(wTxndARow, "TdBChr") + "' And BNo=" + CStr(GrdTxndA.Value(wTxndARow, "TdBNo")) + IIF(gs_Partition = ctCurrPrtn, " And BPrtKey='" + ctCurrPrtn + "'", ""))
      .Store "TdBstOdTc", ""
      .Store "TdBstOdYy", ""
      .Store "TdBstOdChr", ""
      .Store "TdBstOdNo", 0
      .Store "TdBstOdSr", 0
      .Store "TdBstDmCd", ""
                          
    If .IsNew(RowNum) Then
'      .Store "TdTIdNo", moCn.GetFldVal("Select TIdNo From Txn " + _
                        "Where TCoCd='" + ADC("TCoCd") + "' And TTc='" + ADC("TTc") + "' And TYy='" + ADC("TYy") + "' And Tchr='" + ADC("TChr") + "' And TNo=" + CStr(ADC("TNo")) + " And TPrtKey='" + ctCurrPrtn + "'")
'      .Store "TdDt", ADC("TDt")
'      .Store "TdBIdNo", moCn.GetFldVal("Select BIdNo From Bag " + _
 '                         " Where BCoCd='" + ADC("TCoCd") + "' And BYy='" + GrdTxndA.Value(wTxndARow, "TdBYy") + "' And BChr='" + GrdTxndA.Value(wTxndARow, "TdBChr") + "' And BNo=" + CStr(GrdTxndA.Value(wTxndARow, "TdBNo")) + IIF(gs_Partition = ctCurrPrtn, " And BPrtKey='" + ctCurrPrtn + "'", ""))
      '******** Manali 3.2.2  - Inserted User and date txnd grids
      .Store "InsUsr", "MIC" 'UCase(Trim(gs_UsrCd))
      .Store "InsDt", moCn.GetFldVal("Select getdate()")
      
      '4.1.2
      .Store "TdBstOdTc", ""
      .Store "TdBstOdYy", ""
      .Store "TdBstOdChr", ""
      .Store "TdBstOdNo", 0
      .Store "TdBstOdSr", 0
      .Store "TdBstDmCd", ""
    End If
    '****** Sachin 3.02 (14/09/07) Id Fields
    .Store "TdAck", "Y"
    .Store "TdLotMstRecYN", ""
    
    .Store "TdPdQty", 0
      .Store "TdCoCd", ADC("TCoCd")
      .Store "TdTc", ADC("TTc")
      .Store "TdYy", ADC("TYy")
      .Store "TdChr", ADC("TChr")
      .Store "TdNo", ADC("TNo")
      .Store "TdSr", GrdTxndA.Value(wTxndARow, "TdSr")
    
      .Store "TdRefYy", ""
      .Store "TdRefKey", 0
      
      .Store "TdPrevYy", GrdTxndA.Value(wTxndARow, "TdPrevYy")
      .Store "TdPrevKey", GrdTxndA.Value(wTxndARow, "TdPrevKey")
      
      .Store "TdBYy", GrdTxndA.Value(wTxndARow, "TdBYy")
      .Store "TdBChr", GrdTxndA.Value(wTxndARow, "TdBChr")
      .Store "TdBNo", GrdTxndA.Value(wTxndARow, "TdBNo")
      .Store "TdRjDc", ""
      '.Store "TdAck", ""
      .Store "TdPrdYN", ""
      
      .Store "TdBDc", ""
      
      .Store "TdBQty", 0
      .Store "TdPtQty", 0
      .Store "TdBGrWt", 0
      .Store "TdFrBLoc", ""
      .Store "TdToBLoc", ""
      .Store "TdRjQty", 0
      .Store "TdRjPtQty", 0
      .Store "TdRjCd", ""
      .Store "TdWrk", ""
      .Store "TdKey", GrdTxndA.Value(wTxndARow, "TdKey")
      
      wf_TdKeyLen = Len(Trim(Str(GrdTxndA.Value(wTxndARow, "TdKey"))))
      ws_TdYyKey = ADC("TYy") + Space(8 - wf_TdKeyLen) + Trim(Str(GrdTxndA.Value(wTxndARow, "TdKey")))
      .Store "TdYyKey", ws_TdYyKey
      
      .Store "TdPurRt", 0
      .Store "TdPurAmt", 0
      .Store "TdWaxWt", 0#
      '.Store "TdPcPerCt", 0
    
  End With
End Sub
Private Sub GrdTP_RowWhen(ByVal RowNum As Integer)
  Dim wDmCd As String, wPrdTyp As String, wPrcLoc As String
  '*** If GrdTP.IsNew Then TpSrNo = Max + 1
  'uma xxxxxxxx Getting R for P - LocPrcs
  wPrcLoc = moCn.GetFldVal("Select LocPrcs from Loc where LocCoCd = '" + ADC("TCoCd") + "' and LocTyp= 'P' and " + _
               "LocCd = '" + GrdTxndA.Value(GrdTxndA.Row, "TdByLoc") + "'")
               
  wPrdTyp = moCn.GetFldVal("Select LocPrdTyp from Loc where " + _
                  "LocCoCd= '" + ADC("TCoCd") + "' and " + _
                  "LocCd= '" + wPrcLoc + "'")
  With GrdTP
    '.Value(RowNum, "TpSrNo") = 0 Then
    If GrdTP.IsNew(RowNum) Then
      .Value(RowNum, "TpSrNo") = .MaxVal("TpSrNo") + 1
      .Value(RowNum, "TpPrtKey") = ctCurrPrtn '****** Sachin 3.02
      ' ****** Manali 3.8.0 - default values from Previous rows
      If RowNum > 1 Then
        .Value(RowNum, "TpByWrk") = .Value(RowNum - 1, "TpByWrk")
        .Value(RowNum, "TpSubPrc") = .Value(RowNum - 1, "TpSubPrc")
        .Value(RowNum, "TpRmCtg") = .Value(RowNum - 1, "TpRmCtg")
        .Value(RowNum, "TpRmSCtg") = .Value(RowNum - 1, "TpRmSCtg")
        If wPrdTyp <> "GN" Then .Value(RowNum, "TpPrdCtg") = .Value(RowNum - 1, "TpPrdCtg")
        .Value(RowNum, "TpSetTyp") = .Value(RowNum - 1, "TpSetTyp")
      End If
      ' ****** Manali 3.8.0 - default values from Previous rows
      If wPrdTyp = "GN" Then  '****** Sachin 2.13.0 - 27/06/2006 ******
        '****** Sachin 3.02 21-12-07 - Current Partition
        Dim wRsBagDsg As MwfLib.MDORowSet, wDmTcTyp As String
        'indigo.164 recordset created and  BOdDmTcTyp added
        Set wRsBagDsg = moCn.OpenRes("Select BOdDmCd,BOdDmTcTyp from Bag where " + _
                               "BCoCd='" + ADC("TCoCd") + "' and " + _
                               "BYy='" + GrdTxndA.Value(GrdTxndA.Row, "TdBYy") + "' and " + _
                               "BChr='" + GrdTxndA.Value(GrdTxndA.Row, "TdBChr") + "' and " + _
                               "BNo= " + CStr(GrdTxndA.Value(GrdTxndA.Row, "TdBNo")) + _
                               " And BPrtKey= '" + ctCurrPrtn + "'")
                               
        'indigo.164 Dsgcd and Tctyp asssigned
        wDmCd = wRsBagDsg!BOdDmCd: wDmTcTyp = wRsBagDsg!BOdDmTcTyp
        '**********Bhavna
        ' Getting Process loc instead of P Loc for Dpploc
        'indigo.164 DppTcTyp= 'DM' changed as DppTcTyp= '" + wDmTcTyp + "'
        .Value(RowNum, "TpPrdCtg") = moCn.GetFldVal("Select DppPrdCtg from DsgPrd where " + _
                                     "DppTcTyp= '" + wDmTcTyp + "' and DppDmCd= '" + wDmCd + "' and DppLoc= '" + _
                                     wPrcLoc + "' And DppPrtKey='" + ctCurrPrtn + "'")
        'indigo.164 DmTcTyp= 'DM' changed as DmTcTyp= '" + wDmTcTyp + "'
        If .Value(RowNum, "TpPrdCtg") = "" Then
           .Value(RowNum, "TpPrdCtg") = moCn.GetFldVal("Select DmPrdCtg from DsgMst where " + _
                                     "DmTcTyp= '" + wDmTcTyp + "' and DmCd= '" + wDmCd + "' and DmSz= '' And DmPrtKey='" + ctCurrPrtn + "'")
        End If
        '***********Bhavna
      End If
    End If
  End With
End Sub
Private Sub GrdTP_ColWhen(ByVal RowNum As Integer, ByVal ColName As String, Cancel As Boolean, ErrMsg As String, ByVal Action As MwfCtl.en_fgAction)
  '*** Set Helps For Each Field
  '****** Sachin 3.02 21-12-07 - Locking the Records of the Previous Partition
  If GrdTP.Value(RowNum, "TpPrtKey") <> ctCurrPrtn Or mf_DtLocked = True Then Cancel = True: ErrMsg = "Cannot Edit Locked/Previous Partition Entry": Exit Sub
  
' ########################  Manoj #### Ver: 2.0.6 #### Date: 25/12/2003  ########################
  If UCase$(ADC("TLockYN")) = UCase$("Y") Then Cancel = True: ErrMsg = "Cannot Add/Edit/Delete Locked Record": Exit Sub
' ###########################################  Manoj  ###########################################
  If mf_ValidLocSeq = False Then Cancel = True: ErrMsg = "The User Does Not Have Access To One Of The Locs": Exit Sub
  Dim wPrcLoc As String
  'uma xxxxxxxx Getting R for P - LocPrcs
  wPrcLoc = moCn.GetFldVal("Select LocPrcs from Loc where LocCoCd = '" + ADC("TCoCd") + "' and LocTyp= 'P' and " + _
               "LocCd = '" + GrdTxndA.Value(GrdTxndA.Row, "TdByLoc") + "'")

  Select Case ColName
  Case Is = UCase("TpByWrk")
    Call HlpList.vPMCd(ADC("TCoCd"), "WORK")    ' ***** Manali Change Reverted, GrdTxndA.Value(GrdTxndA.Row, "TdByLoc"))   ' ***** Manali 3.6.0 - Worker Help Sorted On Process Loc
  'uma P to R
  Case Is = UCase("TpSubPrc")
    Call HlpList.PSCd("SUBPRC", wPrcLoc)
  Case Is = UCase("TpRmCtg")
    Call HlpList.PMCd("RMCTG")
  Case Is = UCase("TpRmSCtg")
    Call HlpList.PSCd("RMSCTG", GrdTP.Value(RowNum, "TpRmCtg"))
  Case Is = UCase("TpPrdCtg")
    Call HlpList.PMCd("PRDCTG")
  Case Is = UCase("TpSetTyp")
    '3.11.0
    If ms_SetTypOnlyfrOrdYN = "N" Then
        Call HlpList.PSCd("LABSCD", "SET")
    Else
        '***** Geeta 2.11
        '***** Sachin - 2.13.0 - 24/6/2006 - Change done below to be Reverted to Original again ******
        Dim wOrdSr As String, wOrd() As String
        wOrdSr = moCn.GetFldVal("Select BOdTc+','+BOdYy+','+BOdChr+','+Ltrim(str(BOdNo))+" + _
                                "','+LTrim(str(BOdSr)) from Bag where " + _
                                "BCoCd='" + ADC("TCoCd") + "' and " + _
                                "BYy='" + GrdTxndA.Value(GrdTxndA.Row, "TdBYy") + "' and " + _
                                "BChr='" + GrdTxndA.Value(GrdTxndA.Row, "TdBChr") + "' and " + _
                                "BNo= " + CStr(GrdTxndA.Value(GrdTxndA.Row, "TdBNo")))
        wOrd() = Split(wOrdSr, ",")
        Call HlpList.OrSetCd(ADC("TCoCd"), wOrd(0), wOrd(1), wOrd(2), wOrd(3), wOrd(4), GrdTP.Value(RowNum, "TpRmCtg"))
    End If
  '****** Sachin - 2.13.0 - 24/6/2006 - Change done above to be Reverted to Original again ******
  '*** Shilpa Emr 206 03/11/2003 ***
  'And GrdTP.IsNew(GrdTP.Row) = False
  Case Is = UCase("TpPts")
    If GrdTP.Mode = fgmnorm Then
      If moCn.GetFldVal("Select HPrdPtsYN from Head where HCoCd = '" + ADC("TCoCd") + "'") = "N" Then
        Cancel = True: ErrMsg = "Cannot Edit Production Points": Exit Sub
      End If
      Dim wPrdTyp As String
      Dim wRsDsg As MDORowSet
      Dim wDmCd As String
      Dim wDmTcTyp As String
      Dim wDefSubPrc As String
      wPrcLoc = moCn.GetFldVal("Select LocPrcs from Loc where LocCoCd = '" + ADC("TCoCd") + "' and LocTyp= 'P' and " + _
               "LocCd = '" + GrdTxndA.Value(GrdTxndA.Row, "TdByLoc") + "'")

      wPrdTyp = moCn.GetFldVal("Select LocPrdTyp from Loc where " + _
            "LocCoCd= '" + ADC("TCoCd") + "' and " + _
            "LocCd= '" + wPrcLoc + "'")
      If wPrdTyp <> "WS" And wPrdTyp <> "GS" Then
        Exit Sub
      End If
        
      wDefSubPrc = moCn.GetFldVal("Select PSCd from Param where PTyp= 'SUBPRC' and " + _
                "PMCd= '" + wPrcLoc + "' and PSCd= '-'")
  
      Set wRsDsg = moCn.OpenRes("Select BQty, BOdDmCd,BOdDmTcTyp from Bag where " + _
                            "BCoCd='" + ADC("TCoCd") + "' and " + _
                            "BYy='" + GrdTxndA.Value(GrdTxndA.Row, "TdBYy") + "' and " + _
                            "BChr='" + GrdTxndA.Value(GrdTxndA.Row, "TdBChr") + "' and " + _
                            "BNo= " + CStr(GrdTxndA.Value(GrdTxndA.Row, "TdBNo")) + " and BPrtKey='" + ctCurrPrtn + "'")
      If wRsDsg.RecCount > 0 Then
         wRsDsg.MoveFirst: wDmCd = wRsDsg!BOdDmCd
         wDmTcTyp = wRsDsg!BOdDmTcTyp
      End If
      HlpList.PRDPTS wPrdTyp, wPrcLoc, wDefSubPrc, wDmTcTyp, wDmCd, GrdTP.Value(RowNum, "tpsettyp"), GrdTP.Value(RowNum, "TpRmQty")

    End If
  '*** Shilpa Emr 206 03/11/2003 ***
  End Select
End Sub
Private Sub GrdTP_ReCalc(ByVal RowNum As Integer, ByVal ColName As String, ByVal OldVal As Variant, ByVal SourceName As String, ByVal SourceColName As String)
  Dim wPrdTyp As String, wTxndARow As Integer
  Dim wDmCd As String, wPrcLoc As String, wDmTcTyp As String

  With GrdTP
    Select Case UCase(ColName)
    Case Is = UCase("TpPts")
      If .Mode = fgmnorm Then
        wTxndARow = GrdTxndA.Row
    'uma xxxxxxxx Getting R for P - LocPrcs
    wPrcLoc = moCn.GetFldVal("Select LocPrcs from Loc where LocCoCd = '" + ADC("TCoCd") + "' and LocTyp= 'P' and " + _
               "LocCd = '" + GrdTxndA.Value(GrdTxndA.Row, "TdByLoc") + "'")

' All GrdTxndA.Value(wTxndARow, "TdByLoc") changed to wPrcLoc if "R" req
        wPrdTyp = moCn.GetFldVal("Select LocPrdTyp from Loc where " + _
                  "LocCoCd= '" + ADC("TCoCd") + "' and " + _
                  "LocCd= '" + wPrcLoc + "'")
                  
                  '"LocCd= '" + GrdTxndA.Value(wTxndARow, "TdByLoc") + "'")
                                    
        '****** Sachin 3.02 21-12-07 - Current Partition
        wDmCd = moCn.GetFldVal("Select BOdDmCd from Bag where " + _
                               "BCoCd='" + ADC("TCoCd") + "' and " + _
                               "BYy='" + GrdTxndA.Value(wTxndARow, "TdBYy") + "' and " + _
                               "BChr='" + GrdTxndA.Value(wTxndARow, "TdBChr") + "' and " + _
                               "BNo= " + CStr(GrdTxndA.Value(wTxndARow, "TdBNo")) + IIF(gs_Partition = ctCurrPrtn, " and BPrtKey='" + ctCurrPrtn + "'", ""))
        wDmTcTyp = moCn.GetFldVal("Select BOdDmTcTyp from Bag where " + _
                               "BCoCd='" + ADC("TCoCd") + "' and " + _
                               "BYy='" + GrdTxndA.Value(wTxndARow, "TdBYy") + "' and " + _
                               "BChr='" + GrdTxndA.Value(wTxndARow, "TdBChr") + "' and " + _
                               "BNo= " + CStr(GrdTxndA.Value(wTxndARow, "TdBNo")) + IIF(gs_Partition = ctCurrPrtn, " and BPrtKey='" + ctCurrPrtn + "'", ""))
        Select Case UCase(wPrdTyp)
        Case Is = "GN"
          .Value(RowNum, "TpPts") = IIF(.Value(RowNum, "TpBQty") > 0, .Value(RowNum, "TpBQty"), 1) _
               * moCn.GetFldVal("Select IsNull" + _
               "((Select PpPts from PrdPts where PpTyp= '" + wPrdTyp + "' and " + _
                 "PpLoc= '" + wPrcLoc + "' and " + _
                 "PpSubPrc= '" + .Value(RowNum, "TpSubPrc") + "' and PpPrdCtg= '" + .Value(RowNum, "TpPrdCtg") + "' and " + _
                 "PpDmTcTyp='" + wDmTcTyp + "' and " + _
                 "PpDmCd= '" + wDmCd + "' and PpSetTyp= '' and PpPts<> 0 and PpPrtKey='" + GrdTxndA.Value(wTxndARow, "TdPrtKey") + "')" + _
                ", " + _
                "IsNull((Select PpPts from PrdPts where PpTyp= '" + wPrdTyp + "' and " + _
                        "PpLoc= '" + wPrcLoc + "' and " + _
                        "PpSubPrc= '" + .Value(RowNum, "TpSubPrc") + "' and PpPrdCtg= '" + .Value(RowNum, "TpPrdCtg") + "' and " + _
                        "PpDmTcTyp='" + wDmTcTyp + "' and PpDmCd= '' and PpSetTyp= '' and PpPts<> 0 and PpPrtKey='" + GrdTxndA.Value(wTxndARow, "TdPrtKey") + "'),1))")
 Case Is = "WS", "GS"
          Dim wValPts As Double
          Dim wValRmQty As Double
          Dim wf_PickFromMaster As Boolean
          wf_PickFromMaster = True
          
          If UCase(SourceColName) = UCase("TpBQty") Then
            Exit Sub
           End If
          If UCase(SourceColName) = UCase("TpRmQty") Then
              wValRmQty = OldVal
              wValPts = GrdTP.Value(RowNum, "TpPts")
              If wValRmQty = 0 Or wValPts = 0 Then
                  wf_PickFromMaster = True
              Else
                  wf_PickFromMaster = False
              End If
          End If
  
          If wf_PickFromMaster Then
            .Value(RowNum, "TpPts") = .Value(RowNum, "TpRmQty") * _
                                    moCn.GetFldVal(" Select IsNull((Select top 1 PpPts from PrdPts " + _
                                    "where PpTyp= '" + wPrdTyp + "' and " + _
                                    "PpLoc= '" + wPrcLoc + "' and " + _
                                    "PpSubPrc= '" + .Value(RowNum, "TpSubPrc") + "' and " + _
                                    "PpSetTyp= '" + .Value(RowNum, "TpSetTyp") + "' and " + _
                                    "PpPrdCtg= '' and PpDmCd= '" + wDmCd + "' " + _
                                    " and PpPts<> 0 and PpPrtKey='" + _
                                    GrdTxndA.Value(wTxndARow, "TdPrtKey") + "' order by PpFrRmWt ), " + _
                                    " IsNull((Select top 1 PpPts from PrdPts " + _
                                    "where PpTyp= '" + wPrdTyp + "' and " + _
                                    "PpLoc= '" + wPrcLoc + "' and " + _
                                    "PpSubPrc= '" + .Value(RowNum, "TpSubPrc") + "' and " + _
                                    "PpSetTyp= '" + .Value(RowNum, "TpSetTyp") + "' and " + _
                                    "PpPrdCtg= '' and PpDmCd= '' " + _
                                    " and PpPts<> 0 and PpPrtKey='" + _
                                    GrdTxndA.Value(wTxndARow, "TdPrtKey") + "' order by PpFrRmWt),1))")
            Else
              .Value(RowNum, "TpPts") = (wValPts / wValRmQty) * GrdTP.Value(RowNum, "TpRmQty")
            End If
    '' ########################  Manoj #### Ver: 2.0.6 #### Date: 26/04/2004  ########################
        Case Is = "WP", "GP"
          .Value(RowNum, "TpPts") = .Value(RowNum, "TpRmQty") * _
                    moCn.GetFldVal("Select ISNull(( Select PpPts from PrdPts " + _
                                    "where PpTyp= '" + wPrdTyp + "' and " + _
                                    "PpPrtKey='" + GrdTxndA.Value(wTxndARow, "TdPrtKey") + "' and PpLoc= '" + wPrcLoc + "' and " + _
                                    "PpSubPrc= '" + .Value(RowNum, "TpSubPrc") + "' and " + _
                                    "PpSetTyp= '' and PpPrdCtg= '' and PpDmCd= ''  And PpPts<> 0 And " + _
                                    "PpRmCtg = '" + .Value(RowNum, "TpRmCtg") + "' And " + _
                                    "PpRmSCtg = '" + .Value(RowNum, "TpRmSCtg") + "'), " + _
                                    "ISNull(( Select PpPts from PrdPts " + _
                                    "where PpPrtKey='" + GrdTxndA.Value(wTxndARow, "TdPrtKey") + "' and  PpTyp= '" + wPrdTyp + "' and " + _
                                    "PpLoc= '" + wPrcLoc + "' and " + _
                                    "PpSubPrc= '" + .Value(RowNum, "TpSubPrc") + "' and " + _
                                    "PpSetTyp= '' and PpPrdCtg= '' and PpDmCd= ''  And PpPts<> 0 And " + _
                                    "PpRmCtg = '" + .Value(RowNum, "TpRmCtg") + "' And PpRmSCtg = ''),1))")
' ###########################################  Manoj  ###########################################
        End Select
      End If
    End Select
  End With
End Sub
Private Sub GrdTP_ColChange(ByVal RowNum As Integer, ByVal ColName As String, ByVal NewValue As Variant, Cancel As Boolean, ErrMsg As String)
  '*** TpByWrk        valid PMCd('WORK')
  '*** TpSubPrc       valid PSCd('SUBPRC', TdByLoc of 'A' grid)
  '*** TpRmCtg        valid PMCd('RMCTG') or ''
  '*** TpRmSCtg       valid PSCd('RMSCTG', TpRmCtg) or ''
  
  Select Case ColName
  Case Is = UCase("TpByWrk")
    Cancel = Not moCn.RecSeek("Select vPMCd From vParam Where vPCoCd = '" + ADC("TCoCd") + "' And " + _
             "vPTyp= 'WORK' and vPMCd = '" + NewValue + "'")
    If Cancel = True Then ErrMsg = "Invalid By Worker": Exit Sub
  'uma xxxxxxxx Getting R for P -LocPrcs
  Case Is = UCase("TpSubPrc")
  Dim wPrcLoc As String
    wPrcLoc = moCn.GetFldVal("Select LocPrcs from Loc where LocCoCd = '" + ADC("TCoCd") + "' and LocTyp= 'P' and " + _
               "LocCd = '" + GrdTxndA.Value(GrdTxndA.Row, "TdByLoc") + "'")
    Cancel = Not moCn.RecSeek("Select PSCd From Param Where PTyp= 'SUBPRC' And " + _
             "PMCd= '" + wPrcLoc + "' and PSCd= '" + NewValue + "'")
    If Cancel = True Then ErrMsg = "Invalid Sub Process": Exit Sub
  Case Is = UCase("TpRmCtg")
    Cancel = (NewValue <> "") And (Not moCn.RecSeek("Select PMCd From Param Where " + _
             "PTyp= 'RMCTG' And PMCd = '" + NewValue + "'"))
    If Cancel = True Then ErrMsg = "Invalid Raw Material Category": Exit Sub
  Case Is = UCase("TpRmSCtg")
    Cancel = (NewValue <> "") And (Not moCn.RecSeek("Select PSCd From Param Where " + _
             "PTyp= 'RMSCTG' And PMCd = '" + GrdTP.Value(RowNum, "TpRmCtg") + "' and " + _
             "PSCd= '" + NewValue + "'"))
    If Cancel = True Then ErrMsg = "Invalid Raw Sub Material Category": Exit Sub
  Case Is = UCase("TpPrdCtg")
    Cancel = (NewValue <> "") And (Not moCn.RecSeek("Select PmCd From Param where PTyp= 'PRDCTG' " + _
                                  "and PmCd = '" + NewValue + "'"))
    If Cancel = True Then ErrMsg = "Invalid Production Category": Exit Sub
  Case Is = UCase("TpSetTyp")
    '3.11.0
    If ms_SetTypOnlyfrOrdYN = "N" Then
        Cancel = (NewValue <> "") And (Not moCn.RecSeek("Select PSCd From Param where PTyp= 'LABSCD' " + _
                                      "and PMCd= 'SET' and PSCd = '" + NewValue + "'"))
        If Cancel = True Then ErrMsg = "Invalid Setting Type": Exit Sub
    Else
        Cancel = (NewValue <> "") And (Not moCn.RecSeek(" Select 'x' from OrdRm " + _
                                                        "    Join Bag On BCoCd=OrCoCd And BOdTc=OrTc And BOdYy=OrYy And BOdChr=OrChr And BOdNo=OrNo And BOdSr=OrSr" + _
                                                        " Where BCoCd='" + ADC("TCoCd") + "' and  BYy='" + GrdTxndA.Value(GrdTxndA.Row, "TdBYy") + "' and " + _
                                                        " BChr='" + GrdTxndA.Value(GrdTxndA.Row, "TdBChr") + "' and BNo= " + CStr(GrdTxndA.Value(GrdTxndA.Row, "TdBNo")) + _
                                                        " And OrRmCtg='" + GrdTP.Value(RowNum, "TpRmCtg") + "' And OrRmSCtg='" + GrdTP.Value(RowNum, "TpRmSCtg") + "' And OrSetSCd='" + NewValue + "'"))
        If Cancel = True Then ErrMsg = "Setting Type Not Defined In Order Entry": Exit Sub
    End If
    
  End Select
End Sub
Private Sub GrdTP_InitKey(ByVal RowNum As Integer)
  '*** Initialise Key Fields
  With GrdTP
    .Store "TpCoCd", ADC("TCoCd")
    .Store "TpTc", ADC("TTc")
    .Store "TpYy", ADC("TYy")
    .Store "TpChr", ADC("TChr")
    .Store "TpNo", ADC("TNo")
    .Store "TpSr", IIF(GrdTxndA.Value(GrdTxndA.Row, "TdSr") > 0, GrdTxndA.Value(GrdTxndA.Row, "TdSr"), GrdTxndA.Value(1, "TdSr"))
    '****** Sachin 3.02.0 - Init Value for TpTdIdNo
    If .IsNew(RowNum) Then
      '****** Sachin 3.02 21-12-07 - Current Partition
      .Store "TpTdIdNo", moCn.GetFldVal("Select TdIdNo From Txnd Where " + _
                                        " TdCoCd='" + ADC("TCoCd") + "' And TdTc='" + ADC("TTc") + "' And TdYy='" + ADC("TYy") + "' " + _
                                        " And TdChr='" + ADC("TChr") + "' And TdNo=" + CStr(ADC("TNo")) + " And TdSr = " + _
                                        IIF(GrdTxndA.Value(GrdTxndA.Row, "TdSr") > 0, CStr(GrdTxndA.Value(GrdTxndA.Row, "TdSr")), CStr(GrdTxndA.Value(1, "TdSr"))) + " And TdSrNo=0 and TdPrtKey='" + ctCurrPrtn + "'")
      '****** Sachin 3.02 - Newly Introduced Fields ******
      .Store "TpByLoc", GrdTxndA.Value(GrdTxndA.Row, "TdByLoc")
      .Store "TpDt", ADC("TDt")
      '****** Sachin 3.02 - Newly Introduced Fields ******
    End If
  End With
End Sub
Private Sub GrdBRS_ColWhen(ByVal RowNum As Integer, ByVal ColName As String, Cancel As Boolean, ErrMsg As String, ByVal Action As MwfCtl.en_fgAction)
'*** No change allowed in the fields of the working grid GrdBRS of the Bag RM Summary frame
  Cancel = True
End Sub

Private Sub GrdTxndB_When()
    CmdLN.Enabled = True     '**********Bhavna added new lot
End Sub

Private Sub GrdVS_ColWhen(ByVal RowNum As Integer, ByVal ColName As String, Cancel As Boolean, ErrMsg As String, ByVal Action As MwfCtl.en_fgAction)
  '*** No change allowed in the fields of the working grid GrdVS of the Voucher Summary frame
  Cancel = True
End Sub
Private Sub ADC_SetRecSource()
  '*** Set the RecSource for ADC (Table Txn)
  '****** Sachin 3.02 21-12-07 - Corresponding Partition
  
  ADC.RecSource = " Select * from Txn " + _
                  " Where TCoCd ='" + ADC("TCoCd") + "'" + _
                  " And TTc ='" + ADC("TTc") + "'" + _
                  " And TYy = '" + ADC("TYy") + "'" + _
                  " And TChr ='" + ADC("TChr") + "' " + _
                  " And TNo =" + CStr(ADC("TNo")) + " And TPrtKey =(Case when '" + gs_Partition + "'='" + ctCurrPrtn + "' then '" + ctCurrPrtn + "' Else TPrtKey End)" + _
                  IIF(gs_Partition = ctCurrPrtn, " And Tdt>'" + ms_PurgeDt + "' ", "")
                  
End Sub
Private Sub GrdTxndA_SetRecSource()
  '*** Set the RecSource for GrdTxndA (Table Txnd where TdSrNo=0)
  '********************** urmi *************************************************************
  If UCase(ADC.MenuCd) <> UCase("BA") Or ADC("wBAckOpt") = "" Then
    GrdTxndA.RecSource = " Select * from Txnd " + _
                       " Where TdCoCd ='" + ADC("TCoCd") + "'" + _
                       " And TdTc ='" + ADC("TTc") + "'" + _
                       " And TdYy = '" + ADC("TYy") + "'" + _
                       " And TdChr ='" + ADC("TChr") + "' " + _
                       " And TdNo =" + CStr(ADC("TNo")) + _
                       " And TdSrNo = 0 " + _
                       IIF(gs_Partition = ctCurrPrtn, " And TdPrtKey='" + ADC("TPrtKey") + "'", "") + _
                       " Order By TdSr"
  Else
    If UCase(ADC.MenuCd) = UCase("BA") And UCase(ADC("wBAckOpt")) = "Y" Then
       GrdTxndA.RecSource = " Select * from Txnd " + _
                         " Where TdCoCd ='" + ADC("TCoCd") + "'" + _
                         " And TdTc ='" + ADC("TTc") + "'" + _
                         " And TdYy = '" + ADC("TYy") + "'" + _
                         " And TdChr ='" + ADC("TChr") + "' " + _
                         " And TdNo =" + CStr(ADC("TNo")) + _
                         IIF(gs_ValidPrdLocSeq <> "", _
                         " And CharIndex(',' + TdToBLoc + ',', '," + UCase(gs_ValidPrdLocSeq) + ",') <> 0 ", "") + _
                         " And TdAck = 'N' " + _
                         " And TdSrNo = 0 " + _
                         IIF(gs_Partition = ctCurrPrtn, " And TdPrtKey='" + ADC("TPrtKey") + "'", "") + _
                         " Order By TdSr"
    
    ElseIf UCase(ADC.MenuCd) = UCase("BA") And UCase(ADC("wBAckOpt")) = "N" Then
       GrdTxndA.RecSource = " Select * from Txnd " + _
                         " Where TdCoCd ='" + ADC("TCoCd") + "'" + _
                         " And TdTc ='" + ADC("TTc") + "'" + _
                         " And TdYy = '" + ADC("TYy") + "'" + _
                         " And TdChr ='" + ADC("TChr") + "' " + _
                         " And TdNo =" + CStr(ADC("TNo")) + _
                         " And TdAck = 'N' " + _
                         " And TdSrNo = 0 " + _
                         IIF(gs_Partition = ctCurrPrtn, " And TdPrtKey='" + ADC("TPrtKey") + "'", "") + _
                         " Order By TdSr"
    End If
  End If
  '********************** urmi ************************************************************
End Sub
Private Sub GrdTxndB_SetRecSource()
  '*** Set the RecSource for GrdTxndB (Table Txnd where TdSrNo > 0)
  '****** Sachin 3.02 21-12-07 - Corresponding Partition
  
  GrdTxndB.RecSource = " Select * from Txnd " + _
                       " Where TdCoCd ='" + ADC("TCoCd") + "'" + _
                       " And TdTc ='" + ADC("TTc") + "'" + _
                       " And TdYy = '" + ADC("TYy") + "'" + _
                       " And TdChr ='" + ADC("TChr") + "' " + _
                       " And TdNo =" + CStr(ADC("TNo")) + _
                       " And TdSr =" + CStr(GrdTxndA.Value(GrdTxndA.Row, "TdSr")) + _
                       " And TdPrtKey ='" + CStr(GrdTxndA.Value(GrdTxndA.Row, "TdPrtKey")) + "'" + _
                       " And TdSrNo<>0 Order By TdSrNo, TdFrRmDc"
End Sub
Private Sub GrdTP_SetRecSource()
  '*** Set the RecSource for GrdTP (Table TxndPrd)
  '****** Sachin 3.02 21-12-07 - Corresponding Partition
  GrdTP.RecSource = " Select * from TxndPrd " + _
                         " Where TpCoCd ='" + ADC("TCoCd") + "'" + _
                         " And TpTc ='" + ADC("TTc") + "'" + _
                         " And TpYy = '" + ADC("TYy") + "'" + _
                         " And TpChr ='" + ADC("TChr") + "' " + _
                         " And TpNo =" + CStr(ADC("TNo")) + _
                         " And TpSr =" + CStr(GrdTxndA.Value(GrdTxndA.Row, "TdSr")) + _
                         " And TpPrtKey ='" + CStr(GrdTxndA.Value(GrdTxndA.Row, "TdPrtKey")) + "'" + _
                         " Order By TpSrNo"
End Sub
Private Sub adc_AutoGen()
  '*** write code for AutoGenerating the Transaction Voucher Number
  If ADC("TNo") = 0 Then
    ADC("TNo") = moCn.AutoGen("Txn", "TNo", ADC("TCoCd") + ADC("TTc") + ADC("TYy") + ADC("TChr"))
  Else
    Call moCn.AutoGen("Txn", "TNo", ADC("TCoCd") + ADC("TTc") + ADC("TYy") + ADC("TChr"), ADC("TNo"))
  End If
End Sub
Private Sub CmdTP_GotFocus()
  '*** Production Points can be entered only whe1n Production = 'Y' in Grid 'A'
  '*** TdByLoc and TdByWrk should be entered in Grid 'A'
  '*** Default SubProcess should already be defined in Param for the By Location
  '*** When the cursor is in the working field wTdPrdPts of the 'A' grid, the focus is
      'sent to the CmdTP CommandButton and when the mi_Count= 2 then call the PrdPoints
      'sub procedure which enters the Production Points record in the grid TxndPrd and
      'send the focus back to the 'A' grid.
      'If there is some invalid value in the 'A' grid when the focus is sent to CmdTP,
      'the grid does not allow the cursor to leave the grid and the column wTdPrdPts tries
      'to send the focus back to the command button CmdTP. Thus the sequence goes into an
      'infinite loop. To avoid this problem the mi_Count is kept.

  Dim wDefSubPrc As String, wPrcLoc As String
  
 ' If GrdTxndA.Value(GrdTxndA.Row, "TdPrdYN") = "N" Then _
 '    DispMsg "Production Points Can Be Entered Only When Production = 'Y'", etInfo: adc.SetFocus: Exit Sub
  
  If GrdTxndA.Value(GrdTxndA.Row, "TdByLoc") = "" Or _
     GrdTxndA.Value(GrdTxndA.Row, "TdByWrk") = "" Then _
     DispMsg "By Location/Worker Should Be Entered", etInfo: GrdTxndA.SetFocus: Exit Sub  'adc.SetFocus (Old Bef 1.10)
  'uma xxxxxxxx Getting R for P - LocPrcs
  wPrcLoc = moCn.GetFldVal("Select LocPrcs from Loc where LocCoCd = '" + ADC("TCoCd") + "' and LocTyp= 'P' and " + _
               "LocCd = '" + GrdTxndA.Value(GrdTxndA.Row, "TdByLoc") + "'")
  'uma xxxxxxxx Getting - for R
  wDefSubPrc = moCn.GetFldVal("Select PSCd from Param where PTyp= 'SUBPRC' and " + _
                "PMCd= '" + wPrcLoc + "' and PSCd= '-'")
               
               '"PMCd= '" + GrdTxndA.Value(GrdTxndA.Row, "TdByLoc") + "' and PSCd= '-'")
  'If wDefSubPrc = "" Then DispMsg "Define Default SubProcess As '-' For Loc = " + _
                     GrdTxndA.Value(GrdTxndA.Row, "TdByLoc") + " In The Param Table " + _
                     "Before Entering Prd Points", etError: adc.SetFocus: Exit Sub '*** OLD (Bef 1.10)
    
  If wDefSubPrc = "" Then DispMsg "Define Default SubProcess As '-' For Loc " + _
                     "in The Param Table " + _
                     "Before Entering Prd Points", etError: GrdTxndA.SetFocus: Exit Sub  '*** NEW (Rel 1.10)

  If mi_Count = 2 Then Call PrdPoints(wDefSubPrc): mi_Count = mi_Count + 1
  
  'If FraNKeyAll.Enabled = True Then GrdTxndA.SetFocus '*** (Bef 2.13) (Bef 28/11/06) Because of this line, if the Colour Scheme is Nothing then clicking on the 'Prd Pts' button does not open up the Production Points frame
  
  '*** (Jen 2.13) (28/11/06)
  If Not (GetRecClrSet Is Nothing) Or mi_Count > 2 Then
    If FraNKeyAll.Enabled = True Then GrdTxndA.SetFocus
  End If
  '*** (Jen 2.13) (28/11/06)
  
  'GrdTxndA.COL = GrdTxndA.COL + 1   '12         '*** Setting focus to the TdByLoc field
End Sub
Private Sub CmdTP_Click()
  '*** Production Points can be entered only when Production = 'Y' in Grid 'A'
  '*** TdByLoc and TdByWrk should be entered in Grid 'A'
  '*** Default SubProcess should already be defined in Param for the By Location

  Dim wDefSubPrc As String, wPrcLoc As String
  
  ' **** Zubin **** '
  If mb_PrdPtsFlg = True Then Exit Sub
  ' **** Zubin **** '
  
  If GrdTxndA.Row <= 0 Then Exit Sub
  
  'If GrdTxndA.Value(GrdTxndA.Row, "TdPrdYN") = "N" Then _
  '   DispMsg "Production Points Can Be Entered Only When Production = 'Y'", etInfo: adc.SetFocus: Exit Sub
  
  If GrdTxndA.Value(GrdTxndA.Row, "TdByLoc") = "" Or _
     GrdTxndA.Value(GrdTxndA.Row, "TdByWrk") = "" Then _
     DispMsg "By Location/Worker Should Be Entered", etInfo: GrdTxndA.SetFocus: Exit Sub  'adc.SetFocus (Old Bef 1.10)
  
  'uma xxxxxxxx Getting R for P - LocPrcs
  wPrcLoc = moCn.GetFldVal("Select LocPrcs from Loc where LocCoCd = '" + ADC("TCoCd") + "' and LocTyp= 'P' and " + _
               "LocCd = '" + GrdTxndA.Value(GrdTxndA.Row, "TdByLoc") + "'")
  'uma xxxxxxxx Getting - for R
  wDefSubPrc = moCn.GetFldVal("Select PSCd from Param where PTyp= 'SUBPRC' and " + _
                "PMCd= '" + wPrcLoc + "' and PSCd= '-'")
                
               '"PMCd= '" + GrdTxndA.Value(GrdTxndA.Row, "TdByLoc") + "' and PSCd= '-'")
  'If wDefSubPrc = "" Then DispMsg "Define Default SubProcess As '-' For Loc = " + _
                     GrdTxndA.Value(GrdTxndA.Row, "TdByLoc") + " In The Param Table " + _
                     "Before Entering Prd Points", etError: adc.SetFocus: Exit Sub '*** OLD (Bef 1.10)
  
  If wDefSubPrc = "" Then DispMsg "Define Default SubProcess As '-' For Loc " + _
                     "in The Param Table " + _
                     "Before Entering Prd Points", etError: GrdTxndA.SetFocus: Exit Sub  '*** NEW (Rel 1.10)
    
  Call DispFra(PRDPTS)
  'If FraTP.Visible = True Then Call PrdPoints(wDefSubPrc)
  If FraTP.Visible = True And GrdTxndA.Value(GrdTxndA.Row, "TdPrdYN") = "Y" Then Call PrdPoints(wDefSubPrc)
End Sub
Private Sub CmdBD_Click()
  '*** Display or hide the Bag Details frame depending on the toggle state
      'and call the Sub Procedure BagDetails in case the frame is visible
      
  '****** Sachin 3.02 24-12-07 - Cannot View Bag Details If Users Logs to Current Prtn and tries to view previous partition bag
  If gs_Partition = ctCurrPrtn And GrdTxndA.Value(GrdTxndA.Row, "TdPrtKey") <> ctCurrPrtn Then DispMsg "Cannot View Details for Previous Partition Bags", etError: Exit Sub
  
  If GrdTxndA.Row <= 0 Then Exit Sub
  Call DispFra(BagDet)
  If FraBD.Visible = True Then Call BagDetails
  
End Sub

Private Sub CmdBS_Click()
  '*** Display or hide the frame FraBS frame depending on the toggle state
      'and enter the current Bag no as the default bag no whose Rm Summary
      'has to be shown.
      'The user has the option of changing the bag no whose Rm Summary he wants to see

  '****** Sachin 3.02 24-12-07 - Cannot View Bag Details If Users Logs to Current Prtn and tries to view previous partition bag
  If gs_Partition = ctCurrPrtn And GrdTxndA.Value(GrdTxndA.Row, "TdPrtKey") <> ctCurrPrtn Then DispMsg "Cannot View Summary for Previous Partition Bags", etError: Exit Sub
  If GrdTxndA.Row <= 0 Then Exit Sub
  
  Call DispFra(BagRmSumm)
  FraBRS.Visible = False
  FraBRS.Enabled = False

  If FraBS.Visible = True Then
    Dim wTxndARow As Single
    wTxndARow = IIF(GrdTxndA.Row > 0, GrdTxndA.Row, 0)
    If wTxndARow = 0 Then Exit Sub
    ADC("wBsBYy") = GrdTxndA.Value(wTxndARow, "TdBYy")
    ADC("wBsBChr") = GrdTxndA.Value(wTxndARow, "TdBChr")
    ADC("wBsBNo") = GrdTxndA.Value(wTxndARow, "TdBNo")
  End If
End Sub
'**************Bhavna added lab certificate details
Private Sub CmdCert_Click()
  If gs_Partition = ctCurrPrtn And GrdTxndA.Value(GrdTxndA.Row, "TdPrtKey") <> ctCurrPrtn Then DispMsg "Cannot Enter Certificate Details for Previous Partition Bags", etError: Exit Sub
  If GrdTxndA.Row <= 0 Then Exit Sub
  
  Call DispFra(Cert)
  Dim wTxndARow As Single, wBagCert As String, wCert() As String
  If FraCert.Visible = True Then

    wTxndARow = IIF(GrdTxndA.Row > 0, GrdTxndA.Row, 0)
    If wTxndARow = 0 Then Exit Sub
    wBagCert = moCn.GetFldVal("Select BLabCert+','+BLabCmnt+','+BLabDocNo+','+BLabRem " + _
                            " from Bag where " + _
                            "BCoCd='" + ADC("TCoCd") + "' and " + _
                            "BYy='" + GrdTxndA.Value(GrdTxndA.Row, "TdBYy") + "' and " + _
                            "BChr='" + GrdTxndA.Value(GrdTxndA.Row, "TdBChr") + "' and " + _
                            "BNo= " + CStr(GrdTxndA.Value(GrdTxndA.Row, "TdBNo")) + " And BPrtKey='" + ctCurrPrtn + "'")
    wCert() = Split(wBagCert, ",")
    ADC("wBLabCert") = wCert(0)
    ADC("wBLabCmnt") = wCert(1)
    ADC("wBLabDocNo") = wCert(2)
    ADC("wBLabRem") = wCert(3)
  Else

    wTxndARow = IIF(GrdTxndA.Row > 0, GrdTxndA.Row, 0)
    If wTxndARow = 0 Then Exit Sub
    moCn.Execute ("update Bag Set BLabCert = '" + ADC("WBLabCert") + "', " + _
                            "BLabCmnt = '" + ADC("WBLabCmnt") + "', " + _
                            "BLabDocNo = '" + ADC("WBLabDocNo") + "', " + _
                            "BLabRem = '" + ADC("WBLabRem") + "' " + _
                            "where " + _
                            "BCoCd='" + ADC("TCoCd") + "' and " + _
                            "BYy='" + GrdTxndA.Value(GrdTxndA.Row, "TdBYy") + "' and " + _
                            "BChr='" + GrdTxndA.Value(GrdTxndA.Row, "TdBChr") + "' and " + _
                            "BNo= " + CStr(GrdTxndA.Value(GrdTxndA.Row, "TdBNo")) + " And BPrtKey='" + ctCurrPrtn + "'")
    
  End If
End Sub
'**************Bhavna added lab certificate details


Private Sub CmdBSGo_Click()

  '*** Show the Bag Rm Summary frame for the Bag No selected by the user
  FraBRS.Visible = True
  FraBRS.Enabled = True
  FraNKeyAll.Enabled = False
  
  If ADC("wBrsGrp") = "" Then
    ADC("wBrsGrp") = "RMCTG"
  Else
    Call BagRmSummary(ADC("wBrsGrp"), ADC("wBrsValTyp"))
  End If
  FraBRS.ZOrder (0)
End Sub
Private Sub CmdVS_Click()
  '*** Display or hide the Voucher Summary frame depending on the toggle state
      'and call the Sub Procedure VouchSumm in case the frame is visible

  If GrdTxndA.Row <= 0 Then Exit Sub

  Call DispFra(VchSumm)
  If FraVS.Visible = True Then
    If ADC("wVSGrp") = "" Then
      ADC("wVSGrp") = "RMCTG"
    Else
      Call VouchSumm(ADC("wVSGrp"), ADC("wVsValTyp"))       '4.1.2
    End If
  End If
End Sub
Private Sub CmdTot_Click()
  '*** Display or hide the Totals frame depending on the toggle state
      'and call the Sub Procedure Totals in case the frame is visible
  
  If GrdTxndA.Row <= 0 Then Exit Sub
  Call DispFra(Tot)
  If FraTot.Visible = True Then Call Totals
End Sub

Private Sub CmdIR_Click()

  '*** Display or hide the Issue RM frame depending on the toggle state
  If GrdTxndA.Row <= 0 Then Exit Sub
  
 ' If FraIR.Visible = False And GrdTxndB.Rows - 1 > 0 Then Exit Sub
  
  Call DispFra(IssRm)
  If FraIR.Visible = True Then
    ADC("wIrRmCd") = ""
    ' Zubin 212
    ADC("wIrLotNo") = ""
    ' Zubin 212
    ADC("wIrRmAddWt") = 0
    ADC("wIrRmAddQty") = 0      ' ***** Manali - 3.6.0 - 04/09/09 - Add Qty in Issue Rm
    ADC("wIrRmGrWt") = 0
    GrdOrLot.Rows = 1
  End If
End Sub
Private Sub CmdIRGo_Click()
  '*** Call the Sub Procedure IssueRm
  
' ########################  Manoj #### Ver: 2.0.6 #### Date: 25/12/2003  ########################
  If mf_DtLocked = True Or UCase$(ADC("TLockYN")) = UCase$("Y") Then
' ###########################################  Manoj  ###########################################
    DispMsg "Cannot Issue/delete Rm For Locked/Previous Partition Entry", etError
  Else
    If mf_ValidLocSeq = False Then
      DispMsg "The User Does Not Have Access To One Of The Locs", etError
    Else
      Call IssueRm
    End If
  End If
  Call CmdIR_Click
End Sub
Private Sub CmdER_Click()
  '****** When the User clicks on the 'Empty Rm' option the 'Valid' event of the grid
         'GrdTxndA disables all the command buttons associated with it. So the cursor
         'remains in the grid GrdTxndA but the Grid 'GotFocus' event does not happen. So the
         'command buttons associated with the grid remain disabled cwhich causes a problem
         'in making the entry. Hence as soon as the user clicks on 'EmptyRm' the focus
         'is sent to the ADC. The 'Empty Rm' routine is called and the focus is sent back
         'to the grid.
         'This is done only for Command Buttons associated with the grid and which do not
         'have frames associated with them
  
  If GrdTxndA.Row <= 0 Then Exit Sub
  
  ADC.SetFocus
' ########################  Manoj #### Ver: 2.0.6 #### Date: 25/12/2003  ########################
  If mf_DtLocked = True Or UCase$(ADC("TLockYN")) = UCase$("Y") Then
' ###########################################  Manoj  ###########################################
    DispMsg "Cannot Empty Rm For Locked Entry", etError
  Else
    If mf_ValidLocSeq = False Then
      DispMsg "The User Does Not Have Access To One Of The Locs", etError
    Else
      Call EmptyRm
    End If
  End If
  GrdTxndA.SetFocus
End Sub
Private Sub CmdCR_Click()
  '****** When the User clicks on the 'Copy Rm' option the 'Valid' event of the grid
         'GrdTxndA disables all the command buttons associated with it. So the cursor
         'remains in the grid GrdTxndA but the Grid 'GotFocus' event does not happen. So the
         'command buttons associated with the grid remain disabled which causes a problem
         'in making the entry. Hence as soon as the user clicks on 'Copy Rm' the focus
         'is sent to the ADC. The 'CopyRm' routine is called and the focus is sent back
         'to the grid.
         'This is done only for Command Buttons associated with the grid and which do not
         'have frames associated with them
  
  If GrdTxndA.Row <= 0 Then Exit Sub
  
  ADC.SetFocus
' ########################  Manoj #### Ver: 2.0.6 #### Date: 25/12/2003  ########################
  If mf_DtLocked = True Or UCase$(ADC("TLockYN")) = UCase$("Y") Then
' ###########################################  Manoj  ###########################################
    DispMsg "Cannot Copy Rm For Locked/Previous Partition Entry", etError
  Else
    If mf_ValidLocSeq = False Then
      DispMsg "The User Does Not Have Access To One Of The Locs", etError
    Else
      Call CopyRm
    
      'If ms_TcTyp = "BM" Then
      '  Call CopyBMRm
      'Else
      '  Call CopyRm
      'End If
    End If
  End If
  GrdTxndA.SetFocus
End Sub

Private Sub CmdLs_Click()

  '*** Display or hide the Loss Entry frame depending on the toggle state
  ' ########################  Manoj #### Ver: 2.0.6 #### Date: 25/12/2003  ########################
  If mf_DtLocked = True Or UCase$(ADC("TLockYN")) = UCase$("Y") Then DispMsg "Cannot Enter Loss Details For Locked/Previous Partition Entry", etError: Exit Sub
  ' ###########################################  Manoj  ###########################################
  
  If GrdTxndA.Row <= 0 Then Exit Sub
  Call DispFra(Loss)
  If FraLs.Visible = True Then
    ' Manoj 2.10.0
    ' ****** Sachin 3.02 (14/09/07) Id fields added
    '****** Sachin 3.02 22-12-07 - Current Partition
    ADC("wLsMainMetCd") = moCn.GetFldVal("Select OrRmCd From OrdRm Join Bag On OrOdIdNo=BOdIdNo And OrCoCd = BCoCd And OrTc = BOdTc And " & _
                          "OrYy = BOdYy And OrChr = BOdChr And OrNo = BOdNo And OrSr = BOdSr And OrPrtKey=BPrtKey " & _
                          "Where OrMainMet = 'Y' And BCoCd = '" & ADC("TCoCd") & "' And BYy = '" & GrdTxndA.Value(GrdTxndA.Row, "TdBYy") & "' And " & _
                          "BChr = '" & GrdTxndA.Value(GrdTxndA.Row, "TdBChr") & "' And BNo = " & GrdTxndA.Value(GrdTxndA.Row, "TdBNo") & " And BPrtKey='" + ctCurrPrtn + "'")
                          
    '****** Sachin 2.12 - 12/12/05 - SJM  - Silver Included Below
    '****** Sachin 3.02 22-12-07 - Current Partition
    ' ***** Manali 3.03 -20/06/08 - Pd Mod - 'L' included
    ADC("wLsSecMetCd").Enabled = moCn.RecSeek("Select OrRmCd From OrdRm Join Bag On OrOdIdNo=BOdIdNo And OrCoCd = BCoCd And OrTc = BOdTc And OrYy = BOdYy And " & _
                          "OrChr = BOdChr And OrNo = BOdNo And OrSr = BOdSr And OrPrtKey=BPrtKey Join RmMst On OrRmCd = RmCd And OrPrtKey=RmPrtKey " & _
                          "Where OrMainMet = 'N' And RmCtg in ('G', 'P', 'S', 'L') And RmCtg <> " & _
                          "(Select RmCtg From RmMst Where RmCd = '" & ADC("wLsMainMetCd") & "' And RmPrtKey='" + ctCurrPrtn + "') And " & _
                          "BCoCd = '" & ADC("TCoCd") & "' And BYy = '" & GrdTxndA.Value(GrdTxndA.Row, "TdBYy") & "' And " & _
                          "BChr = '" & GrdTxndA.Value(GrdTxndA.Row, "TdBChr") & "' And BNo = " & GrdTxndA.Value(GrdTxndA.Row, "TdBNo") & " And BPrtKey='" + ctCurrPrtn + "' ")
    
    ADC("wLsMainMetAllYn").Enabled = moCn.RecSeek("Select OrRmCd From OrdRm Join Bag On OrOdIdNo=BOdIdNo And OrCoCd = BCoCd And OrTc = BOdTc And OrYy = BOdYy And " & _
                          "OrChr = BOdChr And OrNo = BOdNo And OrSr = BOdSr And OrPrtKey=BPrtKey Join RmMst On OrRmCd = RmCd And OrPrtKey=RmPrtKey " & _
                          "Where OrMainMet = 'N' And RmCtg in ('G', 'P', 'S', 'L') And RmCtg = " & _
                          "(Select RmCtg From RmMst Where RmCd = '" & ADC("wLsMainMetCd") & "' And RmPrtKey='" + ctCurrPrtn + "') And " & _
                          "BCoCd = '" & ADC("TCoCd") & "' And BYy = '" & GrdTxndA.Value(GrdTxndA.Row, "TdBYy") & "' And " & _
                          "BChr = '" & GrdTxndA.Value(GrdTxndA.Row, "TdBChr") & "' And BNo = " & GrdTxndA.Value(GrdTxndA.Row, "TdBNo") & " And BPrtKey='" + ctCurrPrtn + "' ")
    
    ADC("wLsSecMetAllYn").Enabled = IIF(moCn.GetFldVal("Select Count(Distinct OrRmcd) From OrdRm Join Bag On OrOdIdNo=BOdIdNo And OrCoCd = BCoCd And OrTc = BOdTc And OrYy = BOdYy And " & _
                          "OrChr = BOdChr And OrNo = BOdNo And OrSr = BOdSr And OrPrtKey=BPrtKey Join RmMst On OrRmCd = RmCd And OrPrtKey=RmPrtKey " & _
                          "Where OrMainMet = 'N' And RmCtg in ('G', 'P', 'S', 'L') And RmCtg <> " & _
                          "(Select RmCtg From RmMst Where RmCd = '" & ADC("wLsMainMetCd") & "' And RmPrtKey='" + ctCurrPrtn + "') And " & _
                          "BCoCd = '" & ADC("TCoCd") & "' And BYy = '" & GrdTxndA.Value(GrdTxndA.Row, "TdBYy") & "' And " & _
                          "BChr = '" & GrdTxndA.Value(GrdTxndA.Row, "TdBChr") & "' And BNo = " & GrdTxndA.Value(GrdTxndA.Row, "TdBNo") & " And BPrtKey='" + ctCurrPrtn + "' ") > 1, True, False)
 
    
    ' Manoj 2.10.0
    ADC("wLsSecMetCd") = ""
    ADC("wLsGrsWt") = 0
    ADC("wLsDustWt") = 0
    
    ADC("wLsMainMetAllYn") = "N"
    ADC("wLsSecMetAllYn") = "N"
    
    If ADC("wLsMainMetAllYn").Enabled Then
        ADC("wLsMainMetAllYn").SetFocus
    ElseIf ADC("wLsSecMetCd").Enabled Then
        ADC("wLsSecMetCd").SetFocus
    ElseIf ADC("wLsSecMetAllYn").Enabled Then
        ADC("wLsSecMetAllYn").SetFocus
    Else
        ADC("WLSGRSWT").SetFocus
    End If
  
  End If
End Sub
Private Sub CmdLsGo_Click()
  '*** Call the Sub Procedure LossEntry
' ########################  Manoj #### Ver: 2.0.6 #### Date: 25/12/2003  ########################
  If mf_DtLocked = True Or UCase$(ADC("TLockYN")) = UCase$("Y") Then
' ###########################################  Manoj  ###########################################
    DispMsg "Cannot Enter Loss Details For Locked/Previous Partition Entry", etError
  Else
    If mf_ValidLocSeq = False Then
      DispMsg "The User Does Not Have Access To One Of The Locs", etError
    Else
      Call LossEntry
    End If
  End If
  Call DispFra(Loss)
End Sub
Private Sub CmdCB_Click()
  '*** The 'X' record (Txn record) has to be saved before the 'A' records can be entered
  '*** Display or hide the Copy Bag frame depending on the toggle state
  
  If Not ADC.SaveRec Then
    DispMsg "X Save Failed", etInfo
    Exit Sub
  End If
  'IG.89-If rejection code or Bag I/R in X is blank then it shows the error message and will not display the copy frame
  'This check will works only for rejection entry
  If ADC.MenuCd = "REJ" Then
    If ADC("WTRJCD") = "" Or ADC("WTBIR") = "" Then DispMsg "Rejection Code & Bag I/R Should Not Be Blank ", etError: Exit Sub
  End If
  
  Call DispFra(CpyBag)
  If FraCB.Visible = True Then
    '*** (Bef 2.13)
    'ADC("wCbToBLoc") = ""
    'ADC("wCbBYyFr") = "": ADC("wCbBChrFr") = "": ADC("wCbBNoFr") = 0
    'ADC("wCbBYyTo") = "": ADC("wCbBChrTo") = "": ADC("wCbBNoTo") = 0
    'ADC("wCbOdTcFr") = "": ADC("wCbOdYyFr") = "": ADC("wCbOdChrFr") = "": ADC("wCbOdNoFr") = 0:: ADC("wCbOdSrFr") = 0
    'ADC("wCbOdTcTo") = "": ADC("wCbOdYyTo") = "": ADC("wCbOdChrTo") = "": ADC("wCbOdNoTo") = 0:: ADC("wCbOdSrTo") = 0
    'ADC("wCbDmCdFr") = "": ADC("wCbDmCdTo") = ""
    '*** (Bef 2.13)
    
    ' **** Manali 3.6.0 - 26/10/09 - Copy Multi Bags into Melting
    If ADC.MenuCd = "BV" Then
      LblCbLoc.Caption = "To Loc"
      ADC("wCbFrBLoc").Enabled = False: ADC("wCbFrBLoc").Visible = False
      FraCbRej.Enabled = False: FraCbRej.Visible = False      ' ****** Manali 3.10.0 - 30/03/12 - Melting Copy From Rej Vch
    ElseIf ADC.MenuCd = "REJ" Then
      ADC("wCbToBLoc").Enabled = False: ADC("wCbToBLoc").Visible = False
      ' ***** Manali 3.8.0 - Wrk and Prd Y/N Scope added
      LblCbWrk.Visible = False: LblCbPrdYn.Visible = False
      ADC("wCbWrk").Enabled = False: ADC("wCbWrk").Visible = False
      ADC("wCbPrdYn").Enabled = False: ADC("wCbPrdYn").Visible = False
      ' ***** Manali 3.8.0 - Wrk and Prd Y/N Scope added
      'IG.89-ToLoc should not visible for rejection entry
      'LblCbLoc.Visible = False
      'ADC("wCbFrBLoc").Enabled = False: ADC("wCbFrBLoc").Visible = False
      'ig.201
      LblCbLoc.Caption = "Bag Loc"
      ADC("wCbFrBLoc").Left = ADC("wCbToBLoc").Left
    ElseIf ADC.MenuCd = "MLT" Then
      LblCbLoc.Caption = "Fr Loc"
      ADC("wCbToBLoc").Enabled = False: ADC("wCbToBLoc").Visible = False
      ADC("wCbFrBLoc").Left = ADC("wCbToBLoc").Left
      
      ' ***** Manali 3.8.0 - Wrk and Prd Y/N Scope added
      LblCbWrk.Visible = False: LblCbPrdYn.Visible = False
      ADC("wCbWrk").Enabled = False: ADC("wCbWrk").Visible = False
      ADC("wCbPrdYn").Enabled = False: ADC("wCbPrdYn").Visible = False
      ' ***** Manali 3.8.0 - Wrk and Prd Y/N Scope added
    End If
    ' **** Manali 3.6.0 - 26/10/09 - Copy Multi Bags into Melting
    
    '*** (Jen 2.13)
    ADC("wCbBYyFr").CmpStr = "BYy>=": ADC("wCbBChrFr").CmpStr = "BChr>=": ADC("wCbBNoFr").CmpStr = "BNo>="
    ADC("wCbBYyTo").CmpStr = "BYy<=": ADC("wCbBChrTo").CmpStr = "BChr<=": ADC("wCbBNoTo").CmpStr = "BNo<="
    ADC("wCbOdTcFr").CmpStr = "BOdTc>=": ADC("wCbOdYyFr").CmpStr = "BOdYy>=": ADC("wCbOdChrFr").CmpStr = "BOdChr>=": ADC("wCbOdNoFr").CmpStr = "BOdNo>=": ADC("wCbOdSrFr").CmpStr = "BOdSr>="
    ADC("wCbOdTcTo").CmpStr = "BOdTc<=": ADC("wCbOdYyTo").CmpStr = "BOdYy<=": ADC("wCbOdChrTo").CmpStr = "BOdChr<=": ADC("wCbOdNoTo").CmpStr = "BOdNo<=": ADC("wCbOdSrTo").CmpStr = "BOdSr<="
    ADC("wCbDmCdFr").CmpStr = "BOdDmCd>=": ADC("wCbDmCdTo").CmpStr = "BOdDmCd<="
    '*** (Jen 2.13)
    ' **** Manali 3.6.0 - 26/10/09 - Copy Multi Bags into Melting
    If ADC.MenuCd = "BV" Then
      ADC("wCbWrk").CmpStr = "BWrk ="
      If ADC("wCbWrk") = "" Then ADC("wCbWrk") = ctTdWrk
      If ADC("wCbPrdYn") = "" Then ADC("wCbPrdYn") = "N"
    End If
    ' **** Manali 3.6.0 - 26/10/09 - Copy Multi Bags into Melting
    ' ****** Manali 3.10.0 - 30/03/12 - Melting Copy From Rej Vch
    If ADC.MenuCd = "MLT" Then
      ADC("wCbFrBLoc").CmpStr = "BLoc ="
      ADC("wCbTdTc").CmpStr = "TdTc =": ADC("wCbTdYy").CmpStr = "TdYY ="
      ADC("wCbTdChr").CmpStr = "TdChr =": ADC("wCbTdNo").CmpStr = "TdNo ="
      ADC("wCbTdSrFr").CmpStr = "TdSr >=": ADC("wCbTdSrTo").CmpStr = "TdYY <="
    End If
    ' ****** Manali 3.10.0 - 30/03/12 - Melting Copy From Rej Vch
    'ig.201
    If ADC.MenuCd = "REJ" Then ADC("wCbFrBLoc").CmpStr = "BLoc ="
  Else
    '*** (Jen 2.13)
    ADC("wCbBYyFr").CmpStr = "": ADC("wCbBChrFr").CmpStr = "": ADC("wCbBNoFr").CmpStr = ""
    ADC("wCbBYyTo").CmpStr = "": ADC("wCbBChrTo").CmpStr = "": ADC("wCbBNoTo").CmpStr = ""
    ADC("wCbOdTcFr").CmpStr = "": ADC("wCbOdYyFr").CmpStr = "": ADC("wCbOdChrFr").CmpStr = "": ADC("wCbOdNoFr").CmpStr = "": ADC("wCbOdSrFr").CmpStr = ""
    ADC("wCbOdTcTo").CmpStr = "": ADC("wCbOdYyTo").CmpStr = "": ADC("wCbOdChrTo").CmpStr = "": ADC("wCbOdNoTo").CmpStr = "": ADC("wCbOdSrTo").CmpStr = ""
    ADC("wCbDmCdFr").CmpStr = "": ADC("wCbDmCdTo").CmpStr = ""
    '*** (Jen 2.13)
    ADC("wCbWrk").CmpStr = ""       ' **** Manali 3.6.0 - 26/10/09 - Copy Multi Bags into Melting
    
    ' ****** Manali 3.10.0 - 30/03/12 - Melting Copy From Rej Vch
    ADC("wCbFrBLoc").CmpStr = ""
    ADC("wCbTdTc").CmpStr = "": ADC("wCbTdYy").CmpStr = ""
    ADC("wCbTdChr").CmpStr = "": ADC("wCbTdNo").CmpStr = ""
    ADC("wCbTdSrFr").CmpStr = "": ADC("wCbTdSrTo").CmpStr = ""
    ' ****** Manali 3.10.0 - 30/03/12 - Melting Copy From Rej Vch
    
  End If
End Sub
Private Sub CmdCBGo_Click()
  Dim wCopySucc As Boolean
  wCopySucc = True
  '*** Call the Sub Procedure CopyBag
' ########################  Manoj #### Ver: 2.0.6 #### Date: 25/12/2003  ########################
  If mf_DtLocked = True Or UCase$(ADC("TLockYN")) = UCase$("Y") Then
' ###########################################  Manoj  ###########################################
    DispMsg "Cannot Copy Bags For Locked /Previous Partition Entry", etError
  Else
    If mf_ValidLocSeq = False Then
      DispMsg "The User Does Not Have Access To One Of The Locs", etError
    Else
        ' **** Manali 3.6.0 - 26/10/09 - Copy Multi Bags into Melting - wCbFrBLoc
        If UCase(ADC.MenuCd) = UCase("MLT") And ADC("wCbFrBLoc") <> "" And gs_ValidPrdLocSeq <> "" And InStr(1, "," + UCase(gs_ValidPrdLocSeq) + ",", "," + UCase(ADC("wCbFrBLoc")) + ",") <= 0 Then
          DispMsg "The User Does Not Have Access To From Loc", etError
        Else
          wCopySucc = True
        ' **** Manali 3.6.0 - 26/10/09 - Copy Multi Bags into Melting - wCbFrBLoc
          wCopySucc = CopyBag
        End If
    End If
  End If
  If wCopySucc Then Call CmdCB_Click
End Sub

Private Sub CmdCpyRet_Click()
  '*** (Jen 2.13)
  '*** The 'X' record (Txn record) has to be saved before the 'A' records can be entered
  '*** Display or hide the 'Return Bags Copy' frame depending on the toggle state
  If mf_DtLocked = True Or UCase$(ADC("TLockYN")) = UCase$("Y") Then DispMsg "Cannot Copy Returns For Locked/Previous Partition Entry", etError: Exit Sub
  
  If Not ADC.SaveRec Then
    DispMsg "X Save Failed", etInfo
    Exit Sub
  End If
  
  Call DispFra(CpyRet)
  If FraCpyRet.Visible = True Then
    'ADC("wCrBLoc") = ""
    'ADC("wCrOdTc") = "": ADC("wCrOdYy") = "": ADC("wCrOdChr") = "": ADC("wCrOdNo") = 0
    'ADC("wCrOdSrFr") = 0: ADC("wCrOdSrTo") = 0
    'ADC("wCrBYyFr") = "": ADC("wCrBChrFr") = "": ADC("wCrBNoFr") = 0
    'ADC("wCrBYyTo") = "": ADC("wCrBChrTo") = "": ADC("wCrBNoTo") = 0

    ADC("wCrBLoc").CmpStr = "BLoc = "
    ADC("wCrBYyFr").CmpStr = "BYy >=": ADC("wCrBChrFr").CmpStr = "BChr >=": ADC("wCrBNoFr").CmpStr = "BNo >="
    ADC("wCrBYyTo").CmpStr = "BYy <=": ADC("wCrBChrTo").CmpStr = "BChr <=": ADC("wCrBNoTo").CmpStr = "BNo <="
    ADC("wCrOdTc").CmpStr = "BOdTc =": ADC("wCrOdYy").CmpStr = "BOdYy =": ADC("wCrOdChr").CmpStr = "BOdChr =": ADC("wCrOdNo").CmpStr = "BOdNo ="
    ADC("wCrOdSrFr").CmpStr = "BOdSr >=": ADC("wCrOdSrTo").CmpStr = "BOdSr <="
    
    If ADC("wCrBLoc") = "" Then ADC("wCrBLoc") = ADC("TFrBLoc")
    ADC("wCrOdTc") = moCn.GetFldVal("Select PValue from Param where PTyp= 'TCTYP' and PMCd= 'SO'")
    If ADC("wCrOdYy") = "" Then ADC("wCrOdYy") = ADC("TYy")
    If ADC("wCpyOrdRmYn") = "" Then ADC("wCpyOrdRmYn") = "N"    ' ******* Manali 3.5.0 - Change in Return Memo Copy
  Else
    ADC("wCrBLoc").CmpStr = ""
    ADC("wCrBYyFr").CmpStr = "": ADC("wCrBChrFr").CmpStr = "": ADC("wCrBNoFr").CmpStr = ""
    ADC("wCrBYyTo").CmpStr = "": ADC("wCrBChrTo").CmpStr = "": ADC("wCrBNoTo").CmpStr = ""
    ADC("wCrOdTc").CmpStr = "": ADC("wCrOdYy").CmpStr = "": ADC("wCrOdChr").CmpStr = "": ADC("wCrOdNo").CmpStr = ""
    ADC("wCrOdSrFr").CmpStr = "": ADC("wCrOdSrTo").CmpStr = ""
  End If
  '*** (Jen 2.13)
End Sub

Private Sub CmdCpyRetGo_Click()
  '*** (Jen 2.13)
  '*** Call the Sub Procedure GenBagForRet
  If mf_DtLocked = True Or UCase$(ADC("TLockYN")) = UCase$("Y") Then
    DispMsg "Cannot Copy Return Bags For Locked/Previous Partition Entry", etError
  Else
    ' ******* Manali 3.5.0 - Change in Return Memo Copy
    If Not moCn.RecSeek("select PMCd from Param where PTyp= 'YN' and PMCd = '" + ADC("wCpyOrdRmYn") + "'") Then _
        DispMsg "Enter 'Y' To Copy Rm Wts from Order and 'N' to Copy Actual Wts", etError: Exit Sub
    
    If mf_ValidLocSeq = False Then
      DispMsg "The User Does Not Have Access To One Of The Locs", etError
    Else
      Call GenBagForRet
    End If
  End If

  Call CmdCpyRet_Click
  '*** (Jen 2.13)
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
Private Sub PropSet(Optional ByVal mType As String)
  '*** Set Visible and Enabled Properties for X Fields ***
  '*** Set Visible and Enabled Properties for A Fields ***
  '*** Set Visible and Enabled Properties for B Fields ***
  '*** Set Visible and Enabled Properties for Label Conrols ***
  '*** Set Visible and Enabled Properties for Command Buttons ***
  
  Dim wRsX As MwfLib.MDORowSet
  If mType <> "" Then
    Set wRsX = moCn.OpenRes("Select * from TcStr where TsTbl='TxnBag' and " + _
                           "TsTyp='" + mType + "' and TsSTyp='X'")
  End If
  
  Do While Not (wRsX.EOF Or wRsX.BOF)
    ADC(wRsX!TsIdName).Enabled = FldEna(wRsX!TsStyle)
    ADC(wRsX!TsIdName).Visible = FldVis(wRsX!TsStyle)
    wRsX.MoveNext
  Loop
  Set wRsX = Nothing
  
  Call GrdPropSet(GrdTxndA, "A")
  
  Call GrdPropSet(GrdTxndB, "B")
  
  Dim wRsLbl As MwfLib.MDORowSet
  Set wRsLbl = moCn.OpenRes("Select * from TcStr where TsTbl='TxnBag' and " + _
                           "TsTyp='" + ms_TcTyp + "' and TsSTyp='Lbl'")
  Dim wLblCtrl As Control
  Do While Not (wRsLbl.EOF Or wRsLbl.BOF)
    For Each wLblCtrl In Controls
      If TypeOf wLblCtrl Is Label Then
        If wLblCtrl.Name = wRsLbl!TsIdName Then
          wLblCtrl.Visible = FldVis(wRsLbl!TsStyle)
        End If
      End If
    Next
    wRsLbl.MoveNext
  Loop
  Set wRsLbl = Nothing
  
  Call CmdPropSet
End Sub
Private Sub GrdPropSet(ByVal mGrd As Object, ByVal mTsSTyp As String)
  '*** Procedure to set the Style of the columns in the grids
  With mGrd
    Dim wGrdVis As String
    '*** Checking If The Grid Is Visible ***
    wGrdVis = moCn.GetFldVal("Select TsStyle from TcStr where TsTbl='TxnBag' and " + _
                            "TsTyp='" + ms_TcTyp + "' and TsSTyp= '" + mTsSTyp + "' and " + _
                            "TsIdName=''")
    .Visible = FldVis(wGrdVis)
    .Enabled = FldEna(wGrdVis)
    
    '*** If The Grid Is Visible Then Set The Vis And Ena Properties For The Columns For
    '*** Which Record Exists in TcStructure Table
    If FldVis(wGrdVis) = True Then
      Dim wRsGrd As MwfLib.MDORowSet
      Set wRsGrd = moCn.OpenRes("Select * from TcStr where TsTbl='TxnBag' and " + _
                                 "TsTyp='" + ms_TcTyp + "' and TsSTyp= '" + mTsSTyp + "' and " + _
                                 "TsIdName<>''")
      Do While Not (wRsGrd.EOF Or wRsGrd.BOF)
        .ColProp(wRsGrd!TsIdName).Style = ColSty(wRsGrd!TsStyle)
        wRsGrd.MoveNext
      Loop
      Set wRsGrd = Nothing
    End If
  End With
End Sub
Private Sub CmdPropSet()
  '*** Setting Vis and Ena Properties for CommandButton Conrols ***
  Dim wRsCmd As MwfLib.MDORowSet
  Set wRsCmd = moCn.OpenRes("Select * from TcStr where TsTbl= 'TxnBag' and " + _
                            "TsTyp='" + ms_TcTyp + "' and TsSTyp= 'Cmd'")
  Dim wCmdCtrl As Control
  Do While Not (wRsCmd.EOF Or wRsCmd.BOF)
    For Each wCmdCtrl In Controls
      If TypeOf wCmdCtrl Is MWCTL_BTN1 Then
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
Private Sub Totals()
  '*** Procedure to find out the Total Bags, Bag Qty, Part Qty, Gross Wt, Rej Qty,
      'Rej Part Qty in the current voucher

  Dim wRsTot As MwfLib.MDORowSet
  '****** Sachin 3.02 22-12-07 - Corresponding Partition
  Set wRsTot = moCn.OpenRes("Select count(TdTc) as qTotBags, " + _
               "IsNull(sum(case when TdBDc='C' then -TdBQty when TdBDc='D' then TdBQty " + _
               "else TdBQty end), 0) as qTotBQty, " + _
               "IsNull(sum(case when TdBDc='C' then -TdPtQty when TdBDc='D' then TdPtQty " + _
               "else TdPtQty end), 0) as qTotPtQty, " + _
               "IsNull(sum(case when TdBDc='C' then -TdBGrWt when TdBDc='D' then TdBGrWt " + _
               "else TdBGrWt end), 0) as qTotBGrWt, " + _
               "IsNull(sum(case when TdRjDc='C' then -TdRjQty when TdRjDc='D' then TdRjQty " + _
               "else TdRjQty end), 0) as qTotRjQty, " + _
               "IsNull(sum(case when TdRjDc='C' then -TdRjPtQty when TdRjDc='D' then TdRjPtQty " + _
               "else TdRjPtQty end), 0) as qTotRjPtQty " + _
               "from Txnd where TdCoCd= '" + ADC("TCoCd") + "' and TdTc= '" + ADC("TTc") + "' " + _
               "and TdYy= '" + ADC("TYy") + "' and TdChr= '" + ADC("TChr") + "' and " + _
               "TdNo= " + CStr(ADC("TNo")) + " and TdSrNo= 0 and TdPrtKey=(Case When '" + gs_Partition + "'='" + ctCurrPrtn + "' Then '" + ctCurrPrtn + "' Else TdPrtKey End)")

  If Not (wRsTot.EOF Or wRsTot.BOF) Then
    ADC("wTotBags") = wRsTot!qTotBags
    ADC("wTotBQty") = wRsTot!qTotBQty
    ADC("wTotBPtQty") = wRsTot!qTotPtQty
    ADC("wTotBGrWt") = wRsTot!qTotBGrWt
    ADC("wTotRjQty") = wRsTot!qTotRjQty
    ADC("wTotRjPtQty") = wRsTot!qTotRjPtQty
  End If
  Set wRsTot = Nothing
End Sub

Private Sub IssueRm()
  '*** Procedure to Issue Raw Materials into the Bag in case of 'DT' entry
  '*** The User has to specify the Rm Code for which he wants to issue the Wt
  '*** The RmCode should be present in the Order to which the bag belongs
  '*** The user can either enter the additional wt which has to be entered into the bag
      'or the total Gross Wt the Bag should contain after the issue has taken place
  '*** The Records in the Grid 'B' are entered depending on the whether the grouping
      'is done Rm Size wise or Stk Rt wise or both or none
  '*** If the Sz Y/N field of param(PTyp ='RMSCTG', PMCd= RmCtg,
      'PSCd= RmSCtg of the specified RmCode)= 'Y' then Size wise records (from OrdRm)
      'should be entered other wise size wise records should not be entered
  '*** If the StkRt Y/N field of param(Ptyp ='RMSCTG', PMCd= RmCtg,
      'PSCd= RmSCtg of the specified RmCode)= 'Y' then StockRate wise records (from OrdRm)
      'should be entered other wise Stock Rate wise records should not be entered
  '*** The Wt in the last record is the balance wt which is obtained by subtracting the
      'Total Wt issued in the previous rows from the Total Wt to be Issued

  Dim wTxndARow  As Single, wBagGrsWt As Single, wBagTotRmWt As Single
  Dim wOrdTotRmWt  As Single, wMulFact As Single, wTotRows As Single, wRow As Single
  Dim wRmWt As Single, wRmCtg As String, wRmSCtg As String, wGrpBySz As String
  Dim wGrpByPtr As String, wGrpByRt As String, wSel As String, wGrpBy As String, wGrpByPrdPtr As String
  Dim wRmSr As Integer  '4.4.1
  
  ' ***** Manali - 3.6.0 - 04/09/09 - Add Qty in Issue Rm
  Dim wBagAddRmQty As Long, wRs_OrdRm As MDORowSet, wOrdTotRmQty As Long, wMulFactQty As Single, wRmQty As Long
  Dim wGrdCnt  As Integer
    
  ' ********** Manali 350Nxt - Follwing check commented as this routine is also called in GenBagForRet to allow multiple RmCd copy
  'If GrdTxndB.Rows - 1 > 0 Then DispMsg "Record Already Exists", etInfo: Exit Sub
  If ADC("wIrRmCd") = "" Then DispMsg "Enter Raw Material Code", etInfo: Exit Sub
  If ADC("TFrRmLoc") = "" Or ADC("TToRmLoc") = "" Then DispMsg "Enter Default From/To Rm Locations", etError: Exit Sub
  
  wTxndARow = IIF(GrdTxndA.Row, GrdTxndA.Row, 1)
  Dim wRsIr As MwfLib.MDORowSet

  '*** Select Total Weight In The Bag ***
  '****** Sachin 3.02 22-12-07 - Current Partition
  wBagGrsWt = moCn.GetFldVal("Select sum((case (case when TdFrRmLoc='W' then TdFrRmDc " + _
                             "when TdToRmLoc='W' then TdToRmDc end) when 'C' then -1 " + _
                             "when 'D' then +1 else 0 end)* TdRmWt/ " + _
                             "(case when RmCtg in('D', 'C') then 5 else 1 end)) " + _
                             "From Txnd Join RmMst on RmCd= TdRmCd And RmPrtKey=TdPrtKey " + _
                             "Where TdCoCd= '" + ADC("TCoCd") + "' and " + _
                             "TdBYy= '" + GrdTxndA.Value(wTxndARow, "TdBYy") + "' and " + _
                             "TdBChr= '" + GrdTxndA.Value(wTxndARow, "TdBChr") + "' and " + _
                             "TdBNo= " + CStr(GrdTxndA.Value(wTxndARow, "TdBNo")) + " and " + _
                             "TdRmCd<>'' And TdPrtKey='" + ctCurrPrtn + "'")

  '*** Select Total Weight To Be Issued For The RmCd ***
  wBagTotRmWt = IIF(ADC("wIrRmAddWt") > 0, ADC("wIrRmAddWt"), ADC("wIrRmGrWt") - wBagGrsWt)
  If wBagTotRmWt <= 0 Then DispMsg "Cannot Issue RM As Weight <= 0", etError: Set wRsIr = Nothing: Exit Sub
  wBagAddRmQty = ADC("wIrRmAddQty")   ' ***** Manali - 3.6.0 - 04/09/09 - Add Qty in Issue Rm
  
  '*** Select Total Weight For The RmCd In OrdRm ***
  '****** Sachin 3.02 22-12-07 - Current Partition
  ' ***** Manali - 3.6.0 - 04/09/09 - Add Qty in Issue Rm
  Set wRs_OrdRm = moCn.OpenRes("Select IsNull(sum(OrPrdWt), 0) as qTotPrdWt, IsNull(Sum(OrPrdQty), 0) as qTotPrdQty from OrdRm join Bag on " + _
                "BOdIdNo=OrOdIdNo And BCoCd= OrCoCd and BOdTc= OrTc and BOdYy= OrYy and BOdChr= OrChr and " + _
                "BOdNo= OrNo and BOdSr= OrSr and BPrtKey=OrPrtKey join Param on Ptyp = 'BCHR' and PMCd=Bag.BChr and PSCd='' where BCoCd= '" + ADC("TCoCd") + "' and " + _
                "BYy= '" + GrdTxndA.Value(wTxndARow, "TdBYy") + "' and " + _
                "BChr= '" + GrdTxndA.Value(wTxndARow, "TdBChr") + "' and " + _
                "BNo= " + CStr(GrdTxndA.Value(wTxndARow, "TdBNo")) + " and " + _
                "OrRmCd= '" + ADC("wIrRmCd") + "' and OrPrdWt> 0 and (PDesc225='' or PDesc225=OrSubShp) and (PValue3='' or (PValue3='WS' and OrWsQty>0) or (PValue3='GS' and OrHsQty>0)) and  OrPrtKey='" + ctCurrPrtn + "'")
  wOrdTotRmWt = wRs_OrdRm!qTotPrdWt
  wOrdTotRmQty = wRs_OrdRm!qTotPrdQty
  
'''' Bef 3.6.0
''''  wOrdTotRmWt = moCn.GetFldVal("Select sum(OrPrdWt) from OrdRm join Bag on " + _
''''                "BOdIdNo=OrOdIdNo And BCoCd= OrCoCd and BOdTc= OrTc and BOdYy= OrYy and BOdChr= OrChr and " + _
''''                "BOdNo= OrNo and BOdSr= OrSr and BPrtKey=OrPrtKey where BCoCd= '" + adc("TCoCd") + "' and " + _
''''                "BYy= '" + GrdTxndA.Value(wTxndARow, "TdBYy") + "' and " + _
''''                "BChr= '" + GrdTxndA.Value(wTxndARow, "TdBChr") + "' and " + _
''''                "BNo= " + CStr(GrdTxndA.Value(wTxndARow, "TdBNo")) + " and " + _
''''                "OrRmCd= '" + adc("wIrRmCd") + "' and OrPrdWt> 0 and OrPrtKey='" + ctCurrPrtn + "'")
  
  '****** Sachin 3.02 22-12-07 - Current Partition
  wRmCtg = moCn.GetFldVal("Select RmCtg from RmMst where RmCd= '" + ADC("wIrRmCd") + "' and RmPrtKey='" + ctCurrPrtn + "' ")
  wRmSCtg = moCn.GetFldVal("Select RmSCtg from RmMst where RmCd= '" + ADC("wIrRmCd") + "' and RmPrtKey='" + ctCurrPrtn + "' ")
  
  '*** Whether the Rm records have to be entered Size wise or Not
  wGrpBySz = moCn.GetFldVal("Select PDesc225 from Param where PTyp= 'RMSCTG' and " + _
             "PMCd= '" + wRmCtg + "' and PSCd= '" + wRmSCtg + "' ")
  '*** Whether the Rm records have to be entered Pointer wise or Not
  wGrpByPtr = moCn.GetFldVal("Select PValue3 from Param where PTyp= 'RMSCTG' and " + _
             "PMCd= '" + wRmCtg + "' and PSCd= '" + wRmSCtg + "' ")
  '*** Whether the Rm records have to be entered Stock Rate wise or Not
  wGrpByRt = moCn.GetFldVal("Select PValue1 from Param where PTyp= 'RMSCTG' and " + _
             "PMCd= '" + wRmCtg + "' and PSCd= '" + wRmSCtg + "' ")
  
' **** Zubin 211 **** '
'''''  'Geeta*************2.10
'''''   wGrpByPrdPtr = moCn.GetFldVal("Select PValue From Param where PTyp='SYSOPT' and PMCd='DTPRDPTR'")
'''''  'Geeta*************2.10
  wGrpByPrdPtr = moCn.GetFldVal("Select HDtPrdPtr From Head where HCoCd= '" + ADC("TCoCd") + "' and HCd='" + ctSelfCmCd + "'")
' **** Zubin 211 **** '
  
  If wGrpBySz = "Y" And wGrpByRt = "Y" Then
    wSel = " max(OrRmCd) as qOrRmCd, OrLn1 as qOrLn1, 0 as qOrLn2, 0 as qOrLn3,0 as qOrCstRt, " + _
           "sum(OrPrdQty) as qOrQty, sum(OrPrdWt) as qOrWt "
    wGrpBy = "Group By OrLn1 "  ', OrCstRt
  ElseIf wGrpBySz = "N" And wGrpByRt = "Y" And wGrpByPtr = "N" Then
    wSel = " max(OrRmCd) as qOrRmCd, 0 as qOrLn1, 0 as qOrLn2, 0 as qOrLn3, 0 as qOrCstRt, " + _
           "sum(OrPrdQty) as qOrQty, sum(OrPrdWt) as qOrWt "
    wGrpBy = "Group By OrLn1 "  ', OrCstRt
  ElseIf wGrpBySz = "Y" And wGrpByRt = "N" Then
    wSel = " max(OrRmCd) as qOrRmCd, OrLn1 as qOrLn1,  OrLn2 as qOrLn2,  OrLn3 as qOrLn3, 0 as qOrCstRt, " + _
           "sum(OrPrdQty) as qOrQty, sum(OrPrdWt) as qOrWt "
    wGrpBy = "Group By OrLn1,OrLn2,OrLn3 "
'  ElseIf wGrpBySz = "N" And wGrpByRt = "Y" Then
'    wSel = " max(OrRmCd) as qOrRmCd, 0 as qOrLn1, 0 as qOrCstRt, " + _
'           "sum(OrQty) as qOrQty, sum(OrWt) as qOrWt "
'    wGrpBy = "Group By OrCstRt "
'Geeta*****2.10***wGrpByPrdPtr="Y"
  ElseIf wGrpByPtr = "Y" And wGrpByPrdPtr = "Y" Then
    wSel = " max(OrRmCd) as qOrRmCd,(Case when OrPrdQty=0 Then 0 Else OrPrdWt/OrPrdQty End ) as qOrLn1, 0 as qOrLn2, 0 as qOrLn3, 0 as qOrCstRt, " + _
           "sum(OrPrdQty) as qOrQty, sum(OrPrdWt) as qOrWt "
    wGrpBy = "Group By (Case when OrPrdQty=0 Then 0 Else OrPrdWt/OrPrdQty End ) "   'OrCstRt
  ElseIf wGrpByPtr = "Y" And wGrpByRt = "Y" Then
    wSel = " max(OrRmCd) as qOrRmCd, OrRmPtr as qOrLn1, 0 as qOrLn2, 0 as qOrLn3, 0 as qOrCstRt, " + _
           "sum(OrPrdQty) as qOrQty, sum(OrPrdWt) as qOrWt "
    wGrpBy = "Group By OrRmPtr "   'OrCstRt
  ElseIf wGrpByPtr = "Y" And wGrpByRt = "N" Then
    wSel = " max(OrRmCd) as qOrRmCd, OrRmPtr as qOrLn1,0 as qOrLn2, 0 as qOrLn3, 0 as qOrCstRt, " + _
           "sum(OrPrdQty) as qOrQty, sum(OrPrdWt) as qOrWt "
    wGrpBy = "Group By OrRmPtr "
  ElseIf wGrpBySz = "N" And wGrpByPtr = "N" And wGrpByRt = "N" Then
    wSel = " OrRmCd as qOrRmCd, 0 as qOrLn1,0 as qOrLn2, 0 as qOrLn3, 0 as qOrCstRt, " + _
           "sum(OrPrdQty) as qOrQty, sum(OrPrdWt) as qOrWt "
    wGrpBy = "Group By OrRmCd "
  End If
  
  '****** Sachin 3.02 22-12-07 - Current Partition
  Set wRsIr = moCn.OpenRes("Select " + wSel + " from OrdRm join Bag on BOdIdNo=OrOdIdNo And BCoCd= OrCoCd and " + _
                "BOdTc= OrTc and BOdYy= OrYy and BOdChr= OrChr and BOdNo= OrNo and " + _
                "BOdSr= OrSr and BPrtKey=OrPrtKey JOIN Param on PTyp = 'BCHR' and PMcd=BChr and PScd='' where BCoCd= '" + ADC("TCoCd") + "' and " + _
                "BYy= '" + GrdTxndA.Value(wTxndARow, "TdBYy") + "' and " + _
                "BChr= '" + GrdTxndA.Value(wTxndARow, "TdBChr") + "' and " + _
                "BNo= " + CStr(GrdTxndA.Value(wTxndARow, "TdBNo")) + " and " + _
                "OrRmCd='" + ADC("wIrRmCd") + "' and OrPrdWt> 0 and (PDesc225='' or PDesc225=OrSubShp) and (PValue3='' or (PValue3='WS' and OrWsQty>0) or (PValue3='GS' and OrHsQty>0)) and OrPrtKey='" + ctCurrPrtn + "' " + wGrpBy)
  wTotRows = wRsIr.RecCount
  
  wRmSr = moCn.GetFldVal("Select Max(TdSrNo) From Txnd Where TdCoCd= '" + ADC("TCoCd") + "' and " + _
                " TdTc= '" + ADC("TTc") + "' and " + _
                " TdYy= '" + ADC("TYy") + "' and " + _
                " TdChr= '" + ADC("TChr") + "' and " + _
                " TdNo= " + CStr(ADC("TNo")) + " and " + _
                " TdSr=" + CStr(GrdTxndA.Value(wTxndARow, "TdSr")))

  wGrdCnt = GrdTxndB.Rows - 1
  GrdTxndB.StartCopy
  MWLib.BeginProcess Me, "Copying Raw Material Details ..."
  With GrdTxndB
    Do While Not (wRsIr.EOF Or wRsIr.BOF)
      wMulFact = MWLib.Div(wBagTotRmWt, wOrdTotRmWt)
      ' ***** Manali - 3.6.0 - 04/09/09 - Add Qty in Issue Rm
      wMulFactQty = IIF(ADC("wIrRmAddQty") <> 0, MWLib.Div(wBagAddRmQty, wOrdTotRmQty), wMulFact)
      ' ***** Manali - 3.6.0 - 04/09/09 - Add Qty in Issue Rm
      .AddItem
      wRow = .Rows - 1
      .Value(wRow, "TdSrNo") = wRow
      .Value(wRow, "TdFrRmLoc") = ADC("TFrRmLoc")
      .Value(wRow, "wTdFrRmIr") = "I"
      .Value(wRow, "TdFrRmDc") = "C"
      .Value(wRow, "TdRmCd") = wRsIr!qOrRmCd

      .Value(wRow, "TdLotNo") = ADC("wIRLotNo")

      .Value(wRow, "TdRmSz") = wRsIr!qOrLn1
      Dim ErrMsg As String
      
      If ColWhenStkSz(2, wRmCtg, wRmSCtg, ErrMsg) Then
        .Value(wRow, "TdRmSz2") = 0
      Else
        .Value(wRow, "TdRmSz2") = wRsIr!qOrLn2
      End If
      
      If ColWhenStkSz(3, wRmCtg, wRmSCtg, "") Then
        .Value(wRow, "TdRmSz3") = 0
      Else
        .Value(wRow, "TdRmSz3") = wRsIr!qOrLn3
      End If
        
  '*********** urmi AvgStkRt *****************
      '.Value(wRow, "TdRmStkRt") = wRsIR!qOrCstRt
      
      If moCn.GetFldVal("Select hAvgSRtYn from Head where HCoCd= '" + ms_TCoCd + "' and " + _
                      "HCd='" + ctSelfCmCd + "'") = "Y" Then
        Dim wRmLocTyp As String, wb_RmLocTyp As Boolean
        wb_RmLocTyp = moCn.RecSeek("Select LocTyp from Loc where LocCoCd= '" + ADC("TCoCd") + "' " + _
                                   "and LocCd= '" + ADC("TFrRmLoc") + "' " + _
                                   "and LocTyp in " + ctAvgLocTypSeq)
        'Calculate Avg Rt
        wRmLocTyp = Trim(moCn.GetFldVal("Select LocTyp from Loc where LocCoCd= '" + ADC("TCoCd") + "' " + _
                                   "and LocCd= '" + ADC("TFrRmLoc") + "' "))

        If wb_RmLocTyp = True And wGrpByRt = "Y" Then
          If wRmLocTyp <> "W" Then
            ' Zubin 212
            .Value(wRow, "TdRmStkRt") = GetAvgStkRt(ADC("TCoCd"), ADC("TFrRmLoc"), _
                                        wRsIr!qOrRmCd, ADC("wIRLotNo"), wRsIr!qOrLn1, wRsIr!qOrLn2, wRsIr!qOrLn3)
          Else
            ' Zubin 212
            .Value(wRow, "TdRmStkRt") = GetAvgStkRt(ADC("TCoCd"), ADC("TFrRmLoc"), _
                                          wRsIr!qOrRmCd, ADC("wIRLotNo"), wRsIr!qOrLn1, wRsIr!qOrLn2, wRsIr!qOrLn3, _
                                          GrdTxndA.Value(GrdTxndA.Row, "TdBYy"), GrdTxndA.Value(GrdTxndA.Row, "TdBChr"), _
                                          GrdTxndA.Value(GrdTxndA.Row, "TdBNo"))
          End If
        Else
          .Value(wRow, "TdRmStkRt") = 0
        End If
      End If
      '*************************************************
      
      If wRow - wGrdCnt < wTotRows Then
        '*** If Not The Last Record ***
        ' ***** Manali - 3.6.0 - 04/09/09 - Add Qty in Issue Rm
        .Value(wRow, "TdRmQty") = wRsIr!qOrQty * wMulFactQty
        wRmQty = Round(wRsIr!qOrQty * wMulFactQty, 0)
        ' ***** Manali - 3.6.0 - 04/09/09 - Add Qty in Issue Rm
        ''''.Value(wRow, "TdRmQty") = wRsIR!qOrQty * wMulFact   ' Bef 3.6.0
        .Value(wRow, "TdRmWt") = wRsIr!qOrWt * wMulFact
        wRmWt = Round(wRsIr!qOrWt * wMulFact, 3)
      Else
        '*** If Last Record ***
        ' ***** Manali - 3.6.0 - 04/09/09 - Add Qty in Issue Rm
        .Value(wRow, "TdRmQty") = wRsIr!qOrQty * wMulFactQty
        wRmQty = wBagAddRmQty
        ' ***** Manali - 3.6.0 - 04/09/09 - Add Qty in Issue Rm
        ''''.Value(wRow, "TdRmQty") = wRsIR!qOrQty * wMulFact ' Bef 3.6.0
        .Value(wRow, "TdRmWt") = wBagTotRmWt
        wRmWt = wBagTotRmWt
      End If
      ' **** Zubin 211 **** ' (Bef 22/06/05)
      'If .IsNew(GrdTxndB.Row) And UCase(ADC.MenuCd) = "DT" And moCn.RecSeek("Select 'x' from Param where PTyp= 'BCHR' and " + _
      '              "PMCd= '" + GrdTxndA.Value(GrdTxndA.Row, "TdBChr") + "' and PSCd= '' and PValue= 'Y'") Then
      '  .Value(wRow, "TdFbRmQty") = CStr(.Value(wRow, "TdRmQty"))
      '  .Value(wRow, "TdFbRmWt") = CStr(.Value(wRow, "TdRmWt"))
      'End If
      ' **** Zubin 211 **** '
      
      '*** (Jen 22/06/05)
      .Value(wRow, "TdFbRmQty") = ""
      .Value(wRow, "TdFbRmWt") = ""
      '*** (Jen 22/06/05)

      .Value(wRow, "TdToRmLoc") = ADC("TToRmLoc")
      .Value(wRow, "TdPrtKey") = ctCurrPrtn     '****** Sachin 302
      'MMAck-Setting default value for acknowledgement as Y
      .Value(wRow, "TdAck") = "Y"
      '*** (Bef 2.13)  .SaveRec (wRow)
      
      '*** (Jen 2.13)
      If .SaveRec(wRow) = False Then
        ''DispMsg "Could Not Copy All Records.", etWarning
        '.DelAllRecs
        '4.4.1
                
        
        
'        For i = 1 To GrdTxndB.Rows - 1

        Dim i As Integer
        i = GrdTxndB.Rows - 1
        Do While i >= 1
            If i > wRmSr Then _
                GrdTxndB.Del_Rec False, i
            i = i - 1
        Loop
        MsgBox gs_ErrMsg & Chr(13) & " Could Not Copy All Records."
        GoTo IssOver
      End If
      '*** (Jen 2.13)

      wOrdTotRmWt = wOrdTotRmWt - wRsIr!qOrWt
      wBagTotRmWt = wBagTotRmWt - wRmWt
      ' ***** Manali - 3.6.0 - 04/09/09 - Add Qty in Issue Rm
      If ADC("wIrRmAddQty") <> 0 Then
        wOrdTotRmQty = wOrdTotRmQty - wRsIr!qOrQty
        wBagAddRmQty = wBagAddRmQty - wRmQty
      Else
        wOrdTotRmQty = wOrdTotRmQty    'wOrdTotRmWt
        wBagAddRmQty = wBagAddRmQty      'wBagTotRmWt
      End If
      ' ***** Manali - 3.6.0 - 04/09/09 - Add Qty in Issue Rm
      wRsIr.MoveNext
    
    Loop
  End With

  '*** Calculation Of TdBGrsWt ***
  '****** Sachin 3.02 22-12-07 - Current Partition
  GrdTxndA.Value(wTxndARow, "TdBGrWt") = moCn.GetFldVal("Select BGrWt from Bag where " + _
                                         "BCoCd='" + ADC("TCoCd") + "' and " + _
                                         "BYy='" + GrdTxndA.Value(wTxndARow, "TdBYy") + "' and " + _
                                         "BChr='" + GrdTxndA.Value(wTxndARow, "TdBChr") + "' and " + _
                                         "BNo= " + CStr(GrdTxndA.Value(wTxndARow, "TdBNo")) + " and BPrtKey='" + ctCurrPrtn + "'")
  GrdTxndA.SaveRec (wTxndARow)

IssOver:
  MWLib.EndProcess Me
  GrdTxndB.EndCopy
  '*** End of the copy process
  Set wRsIr = Nothing
End Sub


Private Sub CopyRm()
  '*** Procedure to Enter Raw Material details into 'B' Grid in 'BS' entry
  
  Dim wTxndARow As Single, wQty As Single, wBagQty As Single, wMulFact As Single
  Dim wRow As Single, i As Single, wTotIssQty As Single, wTotRecQty As Single
  Dim wSr As Single
  Dim wRsCr As MwfLib.MDORowSet
  
  Dim wf_RefKey As Single   '*** (Jen 2.13)
  
  If GrdTxndB.Rows - 1 > 0 Then DispMsg "Record Already Exists", etInfo: Exit Sub
  wTxndARow = IIF(GrdTxndA.Row, GrdTxndA.Row, 1)

'  GrdTxndB.StartCopy
'  MWLib.BeginProcess Me, "Copying Raw Material Details ..."

  '*** Parent Record ***
  If GrdTxndA.Value(wTxndARow, "TdBDc") = "C" Then
    wQty = GrdTxndA.Value(wTxndARow, "TdBQty")

    '*** Select Total Bag Quantity ***
    '****** Sachin 3.02 22-12-07 - Current Partition
    wBagQty = moCn.GetFldVal("Select BQty from Bag where " + _
              "BCoCd='" + ADC("TCoCd") + "' and " + _
              "BYy= '" + GrdTxndA.Value(wTxndARow, "TdBYy") + "' and " + _
              "BChr= '" + GrdTxndA.Value(wTxndARow, "TdBChr") + "' and " + _
              "BNo= " + CStr(GrdTxndA.Value(wTxndARow, "TdBNo")) + " and BPrtKey='" + ctCurrPrtn + "'")
    wBagQty = wBagQty + wQty    '### Not There In Specs
    '*** (Bef 1.09) If wBagQty = 0 Then DispMsg "The Bag Is Empty", etInfo: Exit Sub: Set wRsCR = Nothing
    '*** (Bef 1.09) wMulFact = MWLib.Div(wQty, wBagQty)
    wMulFact = IIF(wBagQty > 0, MWLib.Div(wQty, wBagQty), 1) '*** (After 1.09)
    
    ' Zubin 212
    Set wRsCr = moCn.OpenRes("Select TdRmCd as qRmCd, TdLotNo as qLotNo, TdRmSz as qRmSz,TdRmSz2 as qRmSz2,TdRmSz3 as qRmSz3, " + _
                "TdRmStkRt as qRmStkRt, sum((case (case when TdFrRmLoc='W' then TdFrRmDc " + _
                "when TdToRmLoc='W' then TdToRmDc else '' end) when 'C' then -1 " + _
                "when 'D' then +1 else 0 end)* TdRmQty) as qRmQty, " + _
                "sum((case (case when TdFrRmLoc='W' then TdFrRmDc when TdToRmLoc='W' then " + _
                "TdToRmDc end) when 'C' then -1 when 'D' then +1 else 0 end)* " + _
                "TdRmWt) as qRmWt From Txnd " + _
                "Where TdRmCd <> '' and TdCoCd='" + ADC("TCoCd") + "' and " + _
                "TdBYy= '" + GrdTxndA.Value(wTxndARow, "TdBYy") + "' and " + _
                "TdBChr= '" + GrdTxndA.Value(wTxndARow, "TdBChr") + "' and " + _
                "TdBNo= " + CStr(GrdTxndA.Value(wTxndARow, "TdBNo")) + " and TdPrtKey='" + ctCurrPrtn + "' " + _
                "Group By TdByy,TdBChr,TdBNo,TdRmCd,TdLotNo,TdRmSz,TdRmSz2,TdRmSz3,TdRmStkRt ")

    GrdTxndB.StartCopy
    MWLib.BeginProcess Me, "Copying Raw Material Details ..."
    With GrdTxndB
      Do While Not (wRsCr.EOF Or wRsCr.BOF)
        .AddItem
        wRow = .Rows - 1
        .Value(wRow, "TdSrNo") = wRow
        .Value(wRow, "TdFrRmLoc") = "W"     '### ADC("TFrRmLoc")
        If .Value(wRow, "TdFrRmLoc") = "W" Then
          .Value(wRow, "wTdFrRmIr") = "I"
        Else
          .Value(wRow, "wTdFrRmIr") = "R"
        End If
        .Value(wRow, "TdRmCd") = wRsCr!qRmCd
        ' Zubin 212
        .Value(wRow, "TdLotNo") = wRsCr!qLotNo
        ' Zubin 212
        .Value(wRow, "TdRmSz") = wRsCr!qRmSz
        .Value(wRow, "TdRmSz2") = wRsCr!qRmSz2
        .Value(wRow, "TdRmSz3") = wRsCr!qRmSz3
        .Value(wRow, "TdRmStkRt") = wRsCr!qRmStkRt
        .Value(wRow, "TdRmQty") = wRsCr!qRmQty * wMulFact
        .Value(wRow, "TdRmWt") = wRsCr!qRmWt * wMulFact
        .Value(wRow, "TdToRmLoc") = "XC"    '### ADC("TToRmLoc")
        '*** (Bef 2.13) .SaveRec (wRow)
        
        .Value(wRow, "TdPrtKey") = ctCurrPrtn   '****** Sachin 3.02.0
        
        '*** (Jen 2.13)
        If .SaveRec(wRow) = False Then
          .DelAllRecs
          MsgBox gs_ErrMsg & Chr(13) & " Could Not Copy All Records."
          GoTo CpyOver
        End If
        '*** (Jen 2.13)
        
        wRsCr.MoveNext
      Loop
    End With
    MWLib.EndProcess Me
    GrdTxndB.EndCopy
  End If

  '*** Child Records ***
  If GrdTxndA.Value(wTxndARow, "TdBDc") = "D" Then
  
    wQty = GrdTxndA.Value(wTxndARow, "TdBQty")

    '*** (Jen 2.13)
    If ms_TcTyp = "BM" Then
      '****** Sachin 3.02 22-12-07 - Current Partition
      wf_RefKey = moCn.GetFldVal("Select TdRefKey from Txnd where TdCoCd= '" + ADC("TCoCd") + "' " + _
                  "and TdTc= '" + ADC("TTc") + "' and TdYy= '" + ADC("TYy") + "' and TdChr= '" + ADC("TChr") + "' " + _
                  "and TdNo= " + CStr(ADC("TNo")) + " and TdSr= " + CStr(GrdTxndA.Value(wTxndARow, "TdSr")) + " " + _
                  "and TdSrNo= 0 and TdPrtKey='" + ctCurrPrtn + "' ")
      
      '****** 18/09/2006 (added:  group By TdBChr)  ******
      '*** Getting The Total Iss Qty ***
      '****** Sachin 3.02 22-12-07 - Current Partition
      wTotIssQty = moCn.GetFldVal("Select sum(TdBQty) from Txnd Td where TdCoCd= '" + ADC("TCoCd") + "' " + _
                   "and TdTc= '" + ADC("TTc") + "' and TdYy= '" + ADC("TYy") + "' " + _
                   "and TdChr= '" + ADC("TChr") + "' and TdNo= " + CStr(ADC("TNo")) + " " + _
                   "and TdSrNo= 0 and TdBDc= 'C' and TdRefKey= " + CStr(wf_RefKey) + _
                   "and TdPrtKey='" + ctCurrPrtn + "' group By TdBChr ")
      '*** 18/09/2006
      
      '*** Getting The Total Rec Qty ***
      '****** Sachin 3.02 22-12-07 - Current Partition
      wTotRecQty = moCn.GetFldVal("Select sum(TdBQty) from Txnd Td where TdCoCd= '" + ADC("TCoCd") + "' " + _
                   "and TdTc= '" + ADC("TTc") + "' and TdYy= '" + ADC("TYy") + "' " + _
                   "and TdChr= '" + ADC("TChr") + "' and TdNo= " + CStr(ADC("TNo")) + " " + _
                   "and TdSrNo= 0 and TdPrtKey='" + ctCurrPrtn + "' and TdBDc= 'D' and TdRefKey= " + CStr(wf_RefKey) + " " + _
                   "and exists (Select 'x' from Txnd Ch where Ch.TdCoCd= '" + ADC("TCoCd") + "' " + _
                   "and Ch.TdTc= '" + ADC("TTc") + "' and Ch.TdYy= '" + ADC("TYy") + "' " + _
                   "and Ch.TdChr= '" + ADC("TChr") + "' and Ch.TdNo= " + CStr(ADC("TNo")) + " " + _
                   "and Ch.TdSr= Td.TdSr and Ch.TdPrtKey=Td.TdPrtKey and Ch.TdSrNo> 0 ) ") + wQty
    Else
    '*** (Jen 2.13)
    
      '*** Getting The Total Iss Qty, Rec Qty ***
      For i = 1 To GrdTxndA.Rows - 1
        If GrdTxndA.Value(i, "TdBDc") = "C" Then
            '''''' **** Zubin 211 **** '
            '''''wTotIssQty = GrdTxndA.Value(i, "TdBQty"): wSr = GrdTxndA.Value(i, "TdSr")
            wTotIssQty = wTotIssQty + GrdTxndA.Value(i, "TdBQty")
            'wSr = GrdTxndA.Value(i, "TdSr") ' Zubin 211 -->> var wSr is no more needed
            '''''' **** Zubin 211 **** '
        ElseIf GrdTxndA.Value(i, "TdBDc") = "D" Then
          wTotRecQty = wTotRecQty + GrdTxndA.Value(i, "TdBQty")
        End If
      Next i
      
    '*** (Jen 2.13)
    End If
    '*** (Jen 2.13)
    
    '*** (Bef 1.09) wMulFact = MWLib.Div(wQty, wTotIssQty)
    
    If ms_TcTyp = "BM" Then
    
      wMulFact = IIF(wQty > 0, MWLib.Div(wQty, wTotIssQty - wTotRecQty + wQty), 1) '*** (After 1.09)
    Else
      wMulFact = IIF(wQty > 0, MWLib.Div(wQty, wTotIssQty), 1)  '*** (After 1.09)
    
    End If
    If ms_TcTyp = "BM" Or (wTotIssQty = wTotRecQty) Then
      '****** (Jen 2.13) added TdRefKey in the condition for 'BM' ******
      ' Zubin 212
      '****** Sachin 3.02 22-12-07 - Current Partition
      Set wRsCr = moCn.OpenRes("Select TB.TdRmCd as TdRmCd, TB.TdLotNo as TdLotNo, TB.TdRmSz as TdRmSz,TB.TdRmSz2 as TdRmSz2, TB.TdRmSz3 as TdRmSz3," + _
                  "   TB.TdRmStkRt as TdRmStkRt, sum(case when TA.TdBDc= 'C' then TB.TdRmQty " + _
                  "       when TA.TdBDc= 'D' then -TB.TdRmQty else 0 end) as qRmQty, " + _
                  "   sum(case when TA.TdBDc= 'C' then TB.TdRmWt when TA.TdBDc= 'D' " + _
                  "       then -TB.TdRmWt else 0 end) as qRmWt from Txnd TB " + _
                  "join Txnd TA on TA.TdCoCd= TB.TdCoCd and TA.TdTc= TB.TdTc and TA.TdYy= TB.TdYy " + _
                  "   and TA.TdChr= TB.TdChr and TA.TdNo= TB.TdNo and TA.TdSr= TB.TdSr and TA.TdPrtKey=TB.TdPrtKey " + _
                  "where TA.TdSrNo= 0 and TB.TdSrNo> 0 and TA.TdCoCd= '" + ADC("TCoCd") + "' " + _
                  "   and TA.TdTc= '" + ADC("TTc") + "' and TA.TdYy= '" + ADC("TYy") + "' " + _
                  "   and TA.TdChr= '" + ADC("TChr") + "' and TA.TdNo= " + CStr(ADC("TNo")) + " and TA.TdPrtKey='" + ctCurrPrtn + "' " + _
                  IIF(ms_TcTyp = "BM", " and TA.TdRefKey= " + CStr(wf_RefKey) + " ", " ") + _
                  "Group By TB.TdRmCd, TB.TdLotNo, TB.TdRmSz, TB.TdRmSz2, TB.TdRmSz3,TB.TdRmStkRt")

                  '(index=Txnd)
    Else
       ' **** Zubin 211 **** '
      '''''Set wRsCR = moCn.OpenRes("Select * from Txnd where " + _
                "TdCoCd= '" + ADC("TCoCd") + "' and TdTc= '" + ADC("TTc") + "' and " + _
                "TdYy= '" + ADC("TYy") + "' and TdChr= '" + ADC("TChr") + "' and " + _
                "TdNo= " + CStr(ADC("TNo")) + " and TdSr= " + CStr(wSr) + " and TdSrNo> 0")
      
      '****** (Jen 2.13) added TdRefKey in the condition for 'BM' ******
      ' Zubin 212
      '****** Sachin 3.02 22-12-07 - Current Partition
      Set wRsCr = moCn.OpenRes("Select TB.TdRmCd, TB.TdLotNo, TB.TdRmSz,TB.TdRmSz2,TB.TdRmSz3, TB.TdRmStkRt, " + _
                  "sum(case when TA.TdBDc= 'C' then TB.TdRmQty else 0 end) as TdRmQty, " + _
                  "sum(case when TA.TdBDc= 'C' then TB.TdRmWt else 0 end) as TdRmWt " + _
                  "from Txnd TB " + _
                  "join Txnd TA on TA.TdCoCd= TB.TdCoCd and TA.TdTc= TB.TdTc and TA.TdYy= TB.TdYy and " + _
                  "   TA.TdChr= TB.TdChr and TA.TdNo = TB.TdNo And TA.TdSr = TB.TdSr And TA.TdPrtKey=TB.TdPrtKey " + _
                  "Where TA.TdSrNo= 0 and TB.TdSrNo> 0 and TA.TdCoCd= '" + ADC("TCoCd") + "' " + _
                  "   and TA.TdTc= '" + ADC("TTc") + "' and TA.TdYy= '" + ADC("TYy") + "' " + _
                  "   and TA.TdChr= '" + ADC("TChr") + "' and TA.TdNo= " + CStr(ADC("TNo")) + " and TA.TdPrtKey='" + ctCurrPrtn + "' " + _
                  IIF(ms_TcTyp = "BM", " and TA.TdRefKey= " + CStr(wf_RefKey) + " ", " ") + _
                  "Group By TB.TdRmCd, TB.TdLotNo, TB.TdRmSz,TB.TdRmSz2,TB.TdRmSz3, TB.TdRmStkRt")
      ' **** Zubin 211 **** '
    End If
    
    GrdTxndB.StartCopy
    MWLib.BeginProcess Me, "Copying Raw Material Details ..."
    With GrdTxndB
      Do While Not (wRsCr.EOF Or wRsCr.BOF)
        .AddItem
        wRow = .Rows - 1
        .Value(wRow, "TdSrNo") = wRow
        .Value(wRow, "TdFrRmLoc") = "W"         '### ADC("TFrRmLoc")
        If .Value(wRow, "TdFrRmLoc") = "W" Then
          .Value(wRow, "wTdFrRmIr") = "R"
        Else
          .Value(wRow, "wTdFrRmIr") = "I"
        End If

        .Value(wRow, "TdRmCd") = wRsCr!TdRmCd
        ' Zubin 212
        .Value(wRow, "TdLotNo") = wRsCr!TdLotNo
        ' Zubin 212
        .Value(wRow, "TdRmSz") = wRsCr!TdRmSz
        .Value(wRow, "TdRmSz2") = wRsCr!TdRmSz2
        .Value(wRow, "TdRmSz3") = wRsCr!TdRmSz3
        .Value(wRow, "TdRmStkRt") = wRsCr!TDRMSTKRT
        If ms_TcTyp = "BM" Then
          .Value(wRow, "TdRmQty") = wRsCr!qRmQty * wMulFact
          .Value(wRow, "TdRmWt") = wRsCr!qRmWt * wMulFact
        Else
          If wTotIssQty = wTotRecQty Then
            'DispMsg "aaa", etError
            .Value(wRow, "TdRmQty") = wRsCr!qRmQty
            .Value(wRow, "TdRmWt") = wRsCr!qRmWt
          Else
            'DispMsg "bbb" + "/" + CStr(wTotIssQty) + "/" + CStr(wTotRecQty), etError
            .Value(wRow, "TdRmQty") = wRsCr!TdRmQty * wMulFact
            .Value(wRow, "TdRmWt") = wRsCr!TdRmWt * wMulFact
          End If
        End If
        .Value(wRow, "TdToRmLoc") = "XC"        '### ADC("TToRmLoc")
        .Value(wRow, "TdPrtKey") = ctCurrPrtn       '****** Sachin 302
        '*** (Bef 2.13) .SaveRec (wRow)
        
        '*** (Jen 2.13)
        If .SaveRec(wRow) = False Then
          .DelAllRecs
          MsgBox gs_ErrMsg & Chr(13) & " Could Not Copy All Records."
          GoTo CpyOver
        End If
        '*** (Jen 2.13)
        
        wRsCr.MoveNext
      Loop
    End With
    MWLib.EndProcess Me
    GrdTxndB.EndCopy
  End If

  '*** Calculation Of TdBGrsWt ***
  '****** Sachin 3.02 22-12-07 - Current Partition
  GrdTxndA.Value(wTxndARow, "TdBGrWt") = moCn.GetFldVal("Select BGrWt from Bag where " + _
                                         "BCoCd= '" + ADC("TCoCd") + "' and " + _
                                         "BYy='" + GrdTxndA.Value(wTxndARow, "TdBYy") + "' and " + _
                                         "BChr='" + GrdTxndA.Value(wTxndARow, "TdBChr") + "' and " + _
                                         "BNo= " + CStr(GrdTxndA.Value(wTxndARow, "TdBNo")) + " and BPrtKey='" + ctCurrPrtn + "'")
  GrdTxndA.SaveRec (wTxndARow)

'  MWLib.EndProcess Me
'  GrdTxndB.EndCopy
  Set wRsCr = Nothing
  '*** End of the Copying process

'*** (Jen 2.13)
  Exit Sub
  
CpyOver:
  MWLib.EndProcess Me
  GrdTxndB.EndCopy
'*** (Jen 2.13)

End Sub
Private Sub CopyRm1()
  '*** (Bef 2.13)
'  '*** Procedure to Enter Raw Material details into 'B' Grid in 'BS' entry
'
'  Dim wTxndARow As Single, wQty As Single, wBagQty As Single, wMulFact As Single
'  Dim wRow As Single, i As Single, wTotIssQty As Single, wTotRecQty As Single
'  Dim wSr As Single
'  Dim wRsCR As MwfLib.MDORowSet
'  If GrdTxndB.Rows - 1 > 0 Then DispMsg "Record Already Exists", etInfo: Exit Sub
'  wTxndARow = IIF(GrdTxndA.Row, GrdTxndA.Row, 1)
'
''  GrdTxndB.StartCopy
''  MWLib.BeginProcess Me, "Copying Raw Material Details ..."
'
'  '*** Parent Record ***
'  If GrdTxndA.Value(wTxndARow, "TdBDc") = "C" Then
'    wQty = GrdTxndA.Value(wTxndARow, "TdBQty")
'
'    '*** Select Total Bag Quantity ***
'    wBagQty = moCn.GetFldVal("Select BQty from Bag where " + _
'              "BCoCd='" + adc("TCoCd") + "' and " + _
'              "BYy= '" + GrdTxndA.Value(wTxndARow, "TdBYy") + "' and " + _
'              "BChr= '" + GrdTxndA.Value(wTxndARow, "TdBChr") + "' and " + _
'              "BNo= " + CStr(GrdTxndA.Value(wTxndARow, "TdBNo")))
'    wBagQty = wBagQty + wQty    '### Not There In Specs
'    '*** (Bef 1.09) If wBagQty = 0 Then DispMsg "The Bag Is Empty", etInfo: Exit Sub: Set wRsCR = Nothing
'    '*** (Bef 1.09) wMulFact = MWLib.Div(wQty, wBagQty)
'    wMulFact = IIF(wBagQty > 0, MWLib.Div(wQty, wBagQty), 1) '*** (After 1.09)
'    ' Zubin 212
'    Set wRsCR = moCn.OpenRes("Select TdRmCd as qRmCd, TdLotNo as qLotNo, TdRmSz as qRmSz, " + _
'                "TdRmStkRt as qRmStkRt, sum((case (case when TdFrRmLoc='W' then TdFrRmDc " + _
'                "when TdToRmLoc='W' then TdToRmDc else '' end) when 'C' then -1 " + _
'                "when 'D' then +1 else 0 end)* TdRmQty) as qRmQty, " + _
'                "sum((case (case when TdFrRmLoc='W' then TdFrRmDc when TdToRmLoc='W' then " + _
'                "TdToRmDc end) when 'C' then -1 when 'D' then +1 else 0 end)* " + _
'                "TdRmWt) as qRmWt From Txnd " + _
'                "Where TdRmCd <> '' and TdCoCd='" + adc("TCoCd") + "' and " + _
'                "TdBYy= '" + GrdTxndA.Value(wTxndARow, "TdBYy") + "' and " + _
'                "TdBChr= '" + GrdTxndA.Value(wTxndARow, "TdBChr") + "' and " + _
'                "TdBNo= " + CStr(GrdTxndA.Value(wTxndARow, "TdBNo")) + " " + _
'                "group By TdByy,TdBChr,TdBNo,TdRmCd,TdLotNo,TdRmSz,TdRmStkRt ")
'
'    GrdTxndB.StartCopy
'    MWLib.BeginProcess Me, "Copying Raw Material Details ..."
'    With GrdTxndB
'      Do While Not (wRsCR.EOF Or wRsCR.BOF)
'        .AddItem
'        wRow = .Rows - 1
'        .Value(wRow, "TdSrNo") = wRow
'        .Value(wRow, "TdFrRmLoc") = "W"     '### ADC("TFrRmLoc")
'        If .Value(wRow, "TdFrRmLoc") = "W" Then
'          .Value(wRow, "wTdFrRmIr") = "I"
'        Else
'          .Value(wRow, "wTdFrRmIr") = "R"
'        End If
'        .Value(wRow, "TdRmCd") = wRsCR!qRmCd
'        ' Zubin 212
'        .Value(wRow, "TdLotNo") = wRsCR!qLotNo
'        ' Zubin 212
'        .Value(wRow, "TdRmSz") = wRsCR!qRmSz
'        .Value(wRow, "TdRmStkRt") = wRsCR!qRmStkRt
'        .Value(wRow, "TdRmQty") = wRsCR!qRmQty * wMulFact
'        .Value(wRow, "TdRmWt") = wRsCR!qRmWt * wMulFact
'        .Value(wRow, "TdToRmLoc") = "XC"    '### ADC("TToRmLoc")
'        .SaveRec (wRow)
'        wRsCR.MoveNext
'      Loop
'    End With
'    MWLib.EndProcess Me
'    GrdTxndB.EndCopy
'  End If
'
'  '*** Child Records ***
'  If GrdTxndA.Value(wTxndARow, "TdBDc") = "D" Then
'    wQty = GrdTxndA.Value(wTxndARow, "TdBQty")
'
'    '*** Getting The Total Iss Qty, Rec Qty ***
'    For i = 1 To GrdTxndA.Rows - 1
'      If GrdTxndA.Value(i, "TdBDc") = "C" Then
'          '''''' **** Zubin 211 **** '
'          '''''wTotIssQty = GrdTxndA.Value(i, "TdBQty"): wSr = GrdTxndA.Value(i, "TdSr")
'          wTotIssQty = wTotIssQty + GrdTxndA.Value(i, "TdBQty")
'          'wSr = GrdTxndA.Value(i, "TdSr") ' Zubin 211 -->> var wSr is no more needed
'          '''''' **** Zubin 211 **** '
'      ElseIf GrdTxndA.Value(i, "TdBDc") = "D" Then
'        wTotRecQty = wTotRecQty + GrdTxndA.Value(i, "TdBQty")
'      End If
'    Next i
'
'    '*** (Bef 1.09) wMulFact = MWLib.Div(wQty, wTotIssQty)
'    wMulFact = IIF(wQty > 0, MWLib.Div(wQty, wTotIssQty), 1)  '*** (After 1.09)
'
'    If wTotIssQty = wTotRecQty Then
'      ' Zubin 212
'      Set wRsCR = moCn.OpenRes("Select TB.TdRmCd as TdRmCd, TB.TdLotNo as TdLotNo, TB.TdRmSz as TdRmSz, " + _
'                  "TB.TdRmStkRt as TdRmStkRt, sum(case when TA.TdBDc='C' then TB.TdRmQty " + _
'                  "when TA.TdBDc='D' then -TB.TdRmQty else 0 end) as qRmQty, " + _
'                  "sum(case when TA.TdBDc='C' then TB.TdRmWt when TA.TdBDc='D' " + _
'                  "then -TB.TdRmWt else 0 end) as qRmWt from Txnd TB " + _
'                  "join Txnd TA on TA.TdCoCd= TB.TdCoCd and TA.TdTc= TB.TdTc and " + _
'                  "TA.TdYy= TB.TdYy and TA.TdChr= TB.TdChr and TA.TdNo= TB.TdNo and " + _
'                  "TA.TdSr= TB.TdSr where TA.TdSrNo= 0 and TB.TdSrNo> 0 " + _
'                  " and TA.TdCoCd= '" + adc("TCoCd") + "' and " + _
'                  "TA.TdTc='" + adc("TTc") + "' and TA.TdYy= '" + adc("TYy") + "' and " + _
'                  "TA.TdChr= '" + adc("TChr") + "' and TA.TdNo= " + CStr(adc("TNo")) + _
'                  " Group By TB.TdRmCd, TB.TdLotNo, TB.TdRmSz, TB.TdRmStkRt")
'                  '(index=Txnd)
'    Else
'' **** Zubin 211 **** '
''''''      Set wRsCR = moCn.OpenRes("Select * from Txnd where " + _
'                  "TdCoCd= '" + ADC("TCoCd") + "' and TdTc= '" + ADC("TTc") + "' and " + _
'                  "TdYy= '" + ADC("TYy") + "' and TdChr= '" + ADC("TChr") + "' and " + _
'                  "TdNo= " + CStr(ADC("TNo")) + " and TdSr= " + CStr(wSr) + " and TdSrNo> 0")
'        ' Zubin 212
'        Set wRsCR = moCn.OpenRes("Select TB.TdRmCd, TB.TdLotNo, TB.TdRmSz, TB.TdRmStkRt, " + _
'                    "sum(case when TA.TdBDc='C' then TB.TdRmQty else 0 end) as TdRmQty, " + _
'                    "sum(case when TA.TdBDc='C' then TB.TdRmWt else 0 end) as TdRmWt " + _
'                    "from Txnd TB " + _
'                    "join Txnd TA on TA.TdCoCd= TB.TdCoCd and TA.TdTc= TB.TdTc and TA.TdYy= TB.TdYy and " + _
'                    "   TA.TdChr= TB.TdChr and TA.TdNo = TB.TdNo And TA.TdSr = TB.TdSr " + _
'                    "Where TA.TdSrNo= 0 and TB.TdSrNo> 0 and TA.TdCoCd= '" + adc("TCoCd") + "' and " + _
'                    "   TA.TdTc= '" + adc("TTc") + "' and TA.TdYy= '" + adc("TYy") + "' and " + _
'                    "   TA.TdChr= '" + adc("TChr") + "' and TA.TdNo= " + CStr(adc("TNo")) + " " + _
'                    "Group By TB.TdRmCd, TB.TdLotNo, TB.TdRmSz, TB.TdRmStkRt")
'' **** Zubin 211 **** '
'    End If
'
'    GrdTxndB.StartCopy
'    MWLib.BeginProcess Me, "Copying Raw Material Details ..."
'    With GrdTxndB
'      Do While Not (wRsCR.EOF Or wRsCR.BOF)
'        .AddItem
'        wRow = .Rows - 1
'        .Value(wRow, "TdSrNo") = wRow
'        .Value(wRow, "TdFrRmLoc") = "W"         '### ADC("TFrRmLoc")
'        If .Value(wRow, "TdFrRmLoc") = "W" Then
'          .Value(wRow, "wTdFrRmIr") = "R"
'        Else
'          .Value(wRow, "wTdFrRmIr") = "I"
'        End If
'
'        .Value(wRow, "TdRmCd") = wRsCR!TdRmCd
'        ' Zubin 212
'        .Value(wRow, "TdLotNo") = wRsCR!TdLotNo
'        ' Zubin 212
'        .Value(wRow, "TdRmSz") = wRsCR!TdRmSz
'        .Value(wRow, "TdRmStkRt") = wRsCR!TdRmStkRt
'        If wTotIssQty = wTotRecQty Then
'          .Value(wRow, "TdRmQty") = wRsCR!qRmQty
'          .Value(wRow, "TdRmWt") = wRsCR!qRmWt
'        Else
'          .Value(wRow, "TdRmQty") = wRsCR!TdRmQty * wMulFact
'          .Value(wRow, "TdRmWt") = wRsCR!TdRmWt * wMulFact
'        End If
'        .Value(wRow, "TdToRmLoc") = "XC"        '### ADC("TToRmLoc")
'        .SaveRec (wRow)
'        wRsCR.MoveNext
'      Loop
'    End With
'    MWLib.EndProcess Me
'    GrdTxndB.EndCopy
'  End If
'
'  '*** Calculation Of TdBGrsWt ***
'  GrdTxndA.Value(wTxndARow, "TdBGrWt") = moCn.GetFldVal("Select BGrWt from Bag where " + _
'                                         "BCoCd= '" + adc("TCoCd") + "' and " + _
'                                         "BYy='" + GrdTxndA.Value(wTxndARow, "TdBYy") + "' and " + _
'                                         "BChr='" + GrdTxndA.Value(wTxndARow, "TdBChr") + "' and " + _
'                                         "BNo= " + CStr(GrdTxndA.Value(wTxndARow, "TdBNo")))
'  GrdTxndA.SaveRec (wTxndARow)
'
''  MWLib.EndProcess Me
''  GrdTxndB.EndCopy
'  Set wRsCR = Nothing
'  '*** End of the Copying process
  '*** (Bef 2.13)
End Sub


'/*********************** Shilpa **********************************/
Private Sub TransferRm()
  Dim wTxndAFrRow As Single, wTxndAToRow As Single
  Dim wBagQty As Single, wMulFact As Single, wMulFact1 As Single, wQtyCheck As Single
  Dim wRsTr As MwfLib.MDORowSet, wRow As Single
  Dim i As Single
  '???
  'Shouldn't this be checked from the database?
  '???If GrdTxndB.Rows - 1 > 0 Then DispMsg "Record Already Exists", etInfo: Exit Sub
  On Error GoTo ErrHForTransfer
If GrdTxndA.Rows < 3 Then
  MsgBox "There must be atleast one set of Issue and Receipt to Transfer Rm"
  Exit Sub
End If
For i = 1 To GrdTxndA.Rows - 1 Step 2
'ec.4- checking whether rm details exist or not, if exist donot transfer Rm
 If Not (moCn.RecSeek("Select 'x' from Txnd where TdCoCd= '" + ADC("TCoCd") + "' and " + _
          "TdTc = '" + ADC("TTc") + "' and TdYy = '" + ADC("TYy") + "' and " + _
          "TdChr = '" + ADC("TChr") + "' and TdNo = " + CStr(ADC("TNo")) + " and TdSr=" + CStr(GrdTxndA.Value(i, "TdSr")) + _
          " and TdSrNo<> 0 and TdPrtKey='" + ctCurrPrtn + "' ")) Then
          
  wTxndAFrRow = i: wTxndAToRow = i + 1
    
  '************ Checking Transfer Quantity Available or Not in the First Bag *********************
  '****** Sachin 3.02 22-12-07 - Current Partition
  wQtyCheck = moCn.GetFldVal("Select BQty from Bag where " + _
              "BCoCd= '" + ADC("TCoCd") + "' and " + _
              "BYy= '" + GrdTxndA.Value(wTxndAFrRow, "TdBYy") + "' and " + _
              "BChr= '" + GrdTxndA.Value(wTxndAFrRow, "TdBChr") + "' and " + _
              "BNo= " + CStr(GrdTxndA.Value(wTxndAFrRow, "TdBNo")) + " and BPrtKey='" + ctCurrPrtn + "'")
  
  '??? If wQtyCheck < ADC("TtrfQty") Then DispMsg " Transfer Quantity Less than Bag Quantity", etError: Exit Sub
  wBagQty = wQtyCheck
    
  wMulFact = IIF(wQtyCheck > 0, MWLib.Div(GrdTxndA.Value(wTxndAFrRow, "TdBQty"), wBagQty), 1) '*** (After 1.09)
 ' ************ Checking Transfer Quantity Available or Not in the Second Bag *********************
 '****** Sachin 3.02 22-12-07 - Current Partition
  wQtyCheck = moCn.GetFldVal("Select BQty from Bag where " + _
            "BCoCd= '" + ADC("TCoCd") + "' and " + _
            "BYy= '" + GrdTxndA.Value(wTxndAToRow, "TdBYy") + "' and " + _
            "BChr= '" + GrdTxndA.Value(wTxndAToRow, "TdBChr") + "' and BPrtKey='" + ctCurrPrtn + "' and " + _
            "BNo= " + CStr(GrdTxndA.Value(wTxndAToRow, "TdBNo")))
  '??? If wQtyCheck < ADC("TtrfQty") Then DispMsg " Transfer Quantity Less than Bag Quantity", etError: Exit Sub
  wBagQty = wQtyCheck
  
  wMulFact1 = IIF(wQtyCheck > 0, MWLib.Div(GrdTxndA.Value(wTxndAToRow, "TdBQty"), wBagQty), 1) '*** (After 1.09)
  GrdTxndA.Row = wTxndAFrRow
' ************ Transferring RM From the First Bag in to the Second Bag *********************
'  If GrdTxndA.Value(wTxndAFrRow, "TdBDc") = "C" Then
  Call RmOut(wTxndAFrRow, wTxndAToRow, wMulFact, wMulFact1)
 ' End If
' ************ Transferring RM From the First Bag in to the Second Bag *********************
  GrdTxndA.Row = wTxndAToRow
  'If GrdTxndA.Value(wTxndAToRow, "TdBDc") = "D" Then
  ' Zubin 212
  '****** Sachin 3.02 22-12-07 - Current Partition
  Set wRsTr = moCn.OpenRes("Select TdRmCd as qRmCd, TdLotNo as qLotNo, TdRmSz as qRmSz,TdRmSz2 as qRmSz2,TdRmSz3 as qRmSz3, " + _
              "TdRmStkRt as qRmStkRt, TdRmQty as qRmQty, " + _
              "TdRmWt as qRmWt, TdFrRmDC From Txnd " + _
              "Where TdSrNo<> 0 and TdRmCd <> '' and TdCoCd='" + ADC("TCoCd") + "' and " + _
              "TdTc = '" + ADC("TTc") + "' and TdYy = '" + ADC("TYy") + "' and " + _
              "TdChr = '" + ADC("TChr") + "' and TdNo = " + CStr(ADC("TNo")) + " and " + _
              "TdBYy= '" + GrdTxndA.Value(wTxndAFrRow, "TdBYy") + "' and " + _
              "TdBChr= '" + GrdTxndA.Value(wTxndAFrRow, "TdBChr") + "' and " + _
              "TdBNo= " + CStr(GrdTxndA.Value(wTxndAFrRow, "TdBNo")) + " and " + _
              "TdPrtKey= '" + ctCurrPrtn + "' " + _
              "Order By TdFrRmDc desc, TdBYy, TdBChr, TdBNo, TdRmCd, TdLotNo, TdRmSz,TdRmSz2,TdRmSz3, TdRmStkRt")
     
  GrdTxndB.StartCopy
  MWLib.BeginProcess Me, "Transferying Raw Material Details ..."
  With GrdTxndB
    Do While Not (wRsTr.EOF Or wRsTr.BOF)
      .AddItem
      wRow = .Rows - 1
      .Value(wRow, "TdSrNo") = wRow
      .Value(wRow, "TdFrRmLoc") = "W"     '### ADC("TFrRmLoc")
      If wRsTr("TdFrRmDC") = "D" Then
        .Value(wRow, "wTdFrRmIr") = "I"
      Else
        .Value(wRow, "wTdFrRmIr") = "R"
      End If
      .Value(wRow, "TdRmCd") = wRsTr!qRmCd
      ' Zubin 212
      .Value(wRow, "TdLotNo") = wRsTr!qLotNo
      ' Zubin 212
      .Value(wRow, "TdRmSz") = wRsTr!qRmSz
      .Value(wRow, "TdRmSz2") = wRsTr!qRmSz2
      .Value(wRow, "TdRmSz3") = wRsTr!qRmSz3
      .Value(wRow, "TdRmStkRt") = wRsTr!qRmStkRt
      .Value(wRow, "TdRmQty") = wRsTr!qRmQty
      .Value(wRow, "TdRmWt") = wRsTr!qRmWt
      .Value(wRow, "TdToRmLoc") = "XC"    '### ADC("TToRmLoc")
      .Value(wRow, "TdPrtKey") = ctCurrPrtn   '****** Sachin 3.02.0
      .SaveRec (wRow)
      wRsTr.MoveNext
    Loop
  End With
  MWLib.EndProcess Me
  GrdTxndB.EndCopy
    
 End If
Next
  '*** Calculation Of TdBGrsWt ***
  
  For i = 1 To GrdTxndA.Rows - 1
    '****** Sachin 3.02 22-12-07 - Current Partition
    GrdTxndA.Value(i, "TdBGrWt") = moCn.GetFldVal("Select BGrWt from Bag where " + _
                                         "BCoCd= '" + ADC("TCoCd") + "' and " + _
                                         "BYy='" + GrdTxndA.Value(i, "TdBYy") + "' and " + _
                                         "BChr='" + GrdTxndA.Value(i, "TdBChr") + "' and " + _
                                         "BNo= " + CStr(GrdTxndA.Value(i, "TdBNo")) + " and BPrtKey='" + ctCurrPrtn + "'")
  GrdTxndA.SaveRec (i)
  Next i
  
  MWLib.EndProcess Me
  GrdTxndB.EndCopy
  Exit Sub
'  '*** End of the Copying process
ErrHForTransfer:
  MsgBox "System encountered an error while copying." + Err.Description
End Sub
Private Sub RmOut(FrRowNum As Single, ToRowNum As Single, wMulFact As Single, wMulFact1 As Single)
  Dim wRsIssue As MwfLib.MDORowSet, wRsReciept As MwfLib.MDORowSet
  Dim wRow As Single
  '************ Transferring Rm of First Bag From First Bag To Xc *********************
  ' Zubin 212
  '****** Sachin 3.02 22-12-07 - Current Partition
  Set wRsIssue = moCn.OpenRes("Select TdRmCd as qRmCd, TdLotNo as qLotNo, TdRmSz as qRmSz,TdRmSz2 as qRmSz2, TdRmSz3 as qRmSz3," + _
                 "TdRmStkRt as qRmStkRt, sum((case (case when TdFrRmLoc= 'W' then TdFrRmDc " + _
                 "when TdToRmLoc= 'W' then TdToRmDc else '' end) when 'C' then -1 " + _
                 "when 'D' then +1 else 0 end)* TdRmQty) as qRmQty, " + _
                 "sum((case (case when TdFrRmLoc='W' then TdFrRmDc when TdToRmLoc= 'W' then " + _
                 "TdToRmDc end) when 'C' then -1 when 'D' then +1 else 0 end)* " + _
                 "TdRmWt) as qRmWt From Txnd " + _
                 "Where TdSrNo<> 0 and TdRmCd <> '' and TdCoCd= '" + ADC("TCoCd") + "' and " + _
                 "TdBYy= '" + GrdTxndA.Value(FrRowNum, "TdBYy") + "' and " + _
                 "TdBChr= '" + GrdTxndA.Value(FrRowNum, "TdBChr") + "' and " + _
                 "TdBNo= " + CStr(GrdTxndA.Value(FrRowNum, "TdBNo")) + " and TdPrtKey='" + ctCurrPrtn + "' " + _
                 "Group By TdBYy, TdBChr, TdBNo, TdRmCd, TdLotNo, TdRmSz, TdRmSz2, TdRmSz3, TdRmStkRt ")
      
  GrdTxndB.StartCopy
  MWLib.BeginProcess Me, "Transferring Raw Material Details ..."
  With GrdTxndB
  Do While Not (wRsIssue.EOF Or wRsIssue.BOF)
    .AddItem
    wRow = .Rows - 1
    .Value(wRow, "TdSrNo") = wRow
    .Value(wRow, "TdFrRmLoc") = "W"     '### ADC("TFrRmLoc")
    .Value(wRow, "wTdFrRmIr") = "I"
    .Value(wRow, "TdRmCd") = wRsIssue!qRmCd
    ' Zubin 212
    .Value(wRow, "TdLotNo") = wRsIssue!qLotNo
    ' Zubin 212
    .Value(wRow, "TdRmSz") = wRsIssue!qRmSz
    .Value(wRow, "TdRmSz2") = wRsIssue!qRmSz2
    .Value(wRow, "TdRmSz3") = wRsIssue!qRmSz3
    .Value(wRow, "TdRmStkRt") = wRsIssue!qRmStkRt
    .Value(wRow, "TdRmQty") = wRsIssue!qRmQty * wMulFact
    .Value(wRow, "TdRmWt") = wRsIssue!qRmWt * wMulFact
    .Value(wRow, "TdToRmLoc") = "XC"    '### ADC("TToRmLoc")
    .Value(wRow, "TdPrtKey") = ctCurrPrtn   '****** Sachin 3.02.0
    .SaveRec (wRow)
    wRsIssue.MoveNext
  Loop
  End With
  MWLib.EndProcess Me
  GrdTxndB.EndCopy

  '************ Transferring Rm of Second Bag From Xc to First Bag *********************
  ' Zubin 212
  '****** Sachin 3.02 22-12-07 - Current Partition
  Set wRsReciept = moCn.OpenRes("Select TdRmCd as qRmCd, TdLotNo as qLotNo, TdRmSz as qRmSz,TdRmSz2 as qRmSz2,TdRmSz3 as qRmSz3, " + _
                   "TdRmStkRt as qRmStkRt, sum((case (case when TdFrRmLoc='W' then TdFrRmDc " + _
                   "when TdToRmLoc='W' then TdToRmDc else '' end) when 'C' then -1 " + _
                   "when 'D' then +1 else 0 end)* TdRmQty) as qRmQty, " + _
                   "sum((case (case when TdFrRmLoc='W' then TdFrRmDc when TdToRmLoc='W' then " + _
                   "TdToRmDc end) when 'C' then -1 when 'D' then +1 else 0 end)* " + _
                   "TdRmWt) as qRmWt From Txnd " + _
                   "Where TdSrNo<> 0 and TdRmCd <> '' and TdCoCd='" + ADC("TCoCd") + "' and " + _
                   "TdBYy= '" + GrdTxndA.Value(ToRowNum, "TdBYy") + "' and " + _
                   "TdBChr= '" + GrdTxndA.Value(ToRowNum, "TdBChr") + "' and " + _
                   "TdBNo= " + CStr(GrdTxndA.Value(ToRowNum, "TdBNo")) + " and TdPrtKey='" + ctCurrPrtn + "' " + _
                   "Group By TdBYy, TdBChr, TdBNo, TdRmCd, TdLotNo, TdRmSz,TdRmSz2,TdRmSz3, TdRmStkRt ")

  GrdTxndB.StartCopy
  MWLib.BeginProcess Me, "Transferying Raw Material Details ..."
  With GrdTxndB
    Do While Not (wRsReciept.EOF Or wRsReciept.BOF)
      .AddItem
      wRow = .Rows - 1
      .Value(wRow, "TdSrNo") = wRow
      .Value(wRow, "TdFrRmLoc") = "W"     '### ADC("TFrRmLoc")
      .Value(wRow, "wTdFrRmIr") = "R"
      .Value(wRow, "TdRmCd") = wRsReciept!qRmCd
      ' Zubin 212
      .Value(wRow, "TdLotNo") = wRsReciept!qLotNo
      ' Zubin 212
      .Value(wRow, "TdRmSz") = wRsReciept!qRmSz
      .Value(wRow, "TdRmSz2") = wRsReciept!qRmSz2
      .Value(wRow, "TdRmSz3") = wRsReciept!qRmSz3
      .Value(wRow, "TdRmStkRt") = wRsReciept!qRmStkRt
      .Value(wRow, "TdRmQty") = wRsReciept!qRmQty * wMulFact1
      .Value(wRow, "TdRmWt") = wRsReciept!qRmWt * wMulFact1
      .Value(wRow, "TdToRmLoc") = "XC"    '### ADC("TToRmLoc")
      .Value(wRow, "TdPrtKey") = ctCurrPrtn   '****** Sachin 3.02.0
      .SaveRec (wRow)
      wRsReciept.MoveNext
    Loop
  End With
  MWLib.EndProcess Me
  GrdTxndB.EndCopy
  Set wRsIssue = Nothing
  Set wRsReciept = Nothing
End Sub

'/*********************** Shilpa **********************************/

Private Sub EmptyRm()
  '*** Procedure to enter the RM Details into the 'B' grid in 'MLT' entry
  '*** This procedure moves the contents of the bag from WIP to Melting Location
  
  Dim wTxndARow  As Single, wRow As Single
  If GrdTxndB.Rows - 1 > 0 Then DispMsg "Record Already Exists", etError: Exit Sub
  If ADC("TFrRmLoc") = "" Or ADC("TToRmLoc") = "" Then DispMsg "Enter Default From/To Rm Locations", etError: Exit Sub
  wTxndARow = IIF(GrdTxndA.Row, GrdTxndA.Row, 1)
  Dim wRsER As MwfLib.MDORowSet
  ' Zubin 212
  '****** Sachin 3.02 22-12-07 - Current Partition
  Set wRsER = moCn.OpenRes("Select TdRmCd as qRmCd, TdLotNo as qLotNo, TdRmSz as qRmSz,TdRmSz2 as qRmSz2,TdRmSz3 as qRmSz3, " + _
              "TdRmStkRt as qRmStkRt, sum((case (case when TdFrRmLoc='W' then TdFrRmDc " + _
              "when TdToRmLoc='W' then TdToRmDc else '' end) when 'C' then -1 " + _
              "when 'D' then +1 else 0 end)* TdRmQty) as qRmQty, " + _
              "sum((case (case when TdFrRmLoc='W' then TdFrRmDc when TdToRmLoc='W' then " + _
              "TdToRmDc end) when 'C' then -1 when 'D' then +1 else 0 end)* " + _
              "TdRmWt) as qRmWt From Txnd " + _
              "Where TdRmCd <> '' and TdCoCd= '" + ADC("TCoCd") + "' and " + _
              "TdBYy= '" + GrdTxndA.Value(wTxndARow, "TdBYy") + "' and " + _
              "TdBChr= '" + GrdTxndA.Value(wTxndARow, "TdBChr") + "' and " + _
              "TdBNo= " + CStr(GrdTxndA.Value(wTxndARow, "TdBNo")) + " and " + _
              "TdPrtKey= '" + ctCurrPrtn + "' " + _
              "group By TdByy,TdBChr,TdBNo,TdRmCd,TdLotNo,TdRmSz,TdRmSz2 ,TdRmSz3 ,TdRmStkRt ")

  GrdTxndB.StartCopy
  MWLib.BeginProcess Me, "Emptying The Bag ..."
  With GrdTxndB
    Do While Not (wRsER.EOF Or wRsER.BOF)
      .AddItem
      wRow = .Rows - 1
      .Value(wRow, "TdSrNo") = wRow
      .Value(wRow, "TdFrRmLoc") = ADC("TFrRmLoc")
      .Value(wRow, "wTdFrRmIr") = "I"
      .Value(wRow, "TdFrRmDc") = "C"
      .Value(wRow, "TdRmCd") = wRsER!qRmCd
      ' Zubin 212
      .Value(wRow, "TdLotNo") = wRsER!qLotNo
      ' Zubin 212
      .Value(wRow, "TdRmSz") = wRsER!qRmSz
      .Value(wRow, "TdRmSz2") = wRsER!qRmSz2
      .Value(wRow, "TdRmSz3") = wRsER!qRmSz3
      .Value(wRow, "TdRmStkRt") = wRsER!qRmStkRt
      .Value(wRow, "TdRmQty") = wRsER!qRmQty
      .Value(wRow, "TdRmWt") = wRsER!qRmWt
      .Value(wRow, "TdToRmLoc") = ADC("TToRmLoc")
      .Value(wRow, "TdPrtKey") = ctCurrPrtn   '****** Sachin 3.02.0
      .SaveRec (wRow)
      wRsER.MoveNext
    Loop
  End With
  MWLib.EndProcess Me
  GrdTxndB.EndCopy
  Set wRsER = Nothing
  '*** End of the Emptying process

End Sub
Private Sub BagDetails()

  '*** Procedure to get the Details of the Bag to be displayed in the Bag Details Frame
  If GrdTxndA.Rows - 1 = 0 Then Exit Sub
  Dim wTxndARow  As Single, wCommStrWt As String, wCommStrQty As String
  Dim wRsBagDet As MwfLib.MDORowSet
  
  wTxndARow = IIF(GrdTxndA.Row > 0, GrdTxndA.Row, 1)
  ADC("wBdBag") = GrdTxndA.Value(wTxndARow, "TdBYy") + "/" + GrdTxndA.Value(wTxndARow, "TdBChr") + _
                  "/" + Trim(CStr(GrdTxndA.Value(wTxndARow, "TdBNo")))
  Dim wLocToConsider As String
  wLocToConsider = IIF(ADC.MenuCd = "RDT", "XR", "W")
  wCommStrWt = "sum((case (case when TdFrRmLoc='" + wLocToConsider + "' then TdFrRmDc " + _
             "when TdToRmLoc='" + wLocToConsider + "' then TdToRmDc end) when 'C' then -1 " + _
             "when 'D' then +1 else 0 end)* TdRmWt * "
  wCommStrQty = "sum((case (case when TdFrRmLoc='" + wLocToConsider + "' then TdFrRmDc " + _
             "when TdToRmLoc='" + wLocToConsider + "' then TdToRmDc end) when 'C' then -1 " + _
             "when 'D' then +1 else 0 end)* TdRmQty * "
    
  ' **** Manali 3.03 - 20/08/06 - 'RmDet' split from 'wRsBagDet' due to too many lines
  ' ****** Sachin 2.12 - 02/12/05 - Silver added to Bag Details
  ' ***** Manali 3.03 - 20/06/08 - Pd Mod - 'L' included in Bag Details
  Dim RMDET As String
  RMDET = wCommStrWt + "(case when RmCtg= 'G' then 1 else 0 end)) as qBWtG, " + _
          wCommStrWt + "(case when RmCtg= 'P' then 1 else 0 end)) as qBWtP, " + _
          wCommStrWt + "(case when RmCtg= 'S' then 1 else 0 end)) as qBWtS, " + _
          wCommStrWt + "(case when RmCtg= 'L' then 1 else 0 end)) as qBWtL, " + _
          wCommStrWt + "(case when RmCtg= 'M' then 1 else 0 end)) as qBWtM, " + _
          wCommStrWt + "(case when RmCtg= 'G' and RmSCtg!='-' then 1 else 0 end)) as qBWtFndG, " + _
          wCommStrWt + "(case when RmCtg= 'P' and RmSCtg!='-' then 1 else 0 end)) as qBWtFndP, " + _
          wCommStrWt + "(case when RmCtg= 'S' and RmSCtg!='-' then 1 else 0 end)) as qBWtFndS, " + _
          wCommStrWt + "(case when RmCtg= 'L' and RmSCtg!='-' then 1 else 0 end)) as qBWtFndL, " + _
          wCommStrWt + "(case when RmCtg= 'M' and RmSCtg!='-' then 1 else 0 end)) as qBWtFndM, " + _
          wCommStrWt + "(case when RmCtg= 'D' then 1 else 0 end)) as qBWtD, " + _
          wCommStrWt + "(case when RmCtg= 'C' then 1 else 0 end)) as qBWtC, " + _
          wCommStrWt + "(case when RmCtg= 'X' then 1 else 0 end)) as qBWtX, " + _
          wCommStrQty + "(case when RmCtg= 'D' then 1 else 0 end)) as qBQtyD, " + _
          wCommStrQty + "(case when RmCtg= 'C' then 1 else 0 end)) as qBQtyC, " + _
          wCommStrQty + "(case when RmCtg= 'X' then 1 else 0 end)) as qBQtyX "
  
  Dim wRepairBagExtraCnd As String
  wRepairBagExtraCnd = " and (Txnd.ModDt >=CONVERT(date,BRepINDt) or (Txnd.ModDt =CONVERT(date,BRepINDt)  and Txnd.ModTime >=cast((substring(CONVERT(VARCHAR(10),BRepINDt,108),1,2)+'.'+substring(CONVERT(VARCHAR(10),BRepINDt,108),4,2)) as float))) and TdSrNo <>0  "
          
  ' ****** Sachin 3.02 22-12-07 - Current Partition
  'Ig.200- OdDmCol added
  Set wRsBagDet = moCn.OpenRes("Select max(" + IIF(ADC.MenuCd = "RDT" Or ADC.MenuCd = "RBV", "BRepLoc", "BLoc") + ") as BLoc, max(BWrk) as BWrk, " + _
                  "max(BOdDmTcTyp) as BOdDmTcTyp,max(BOdDmCd) as BOdDmCd, max(BOdSfx) as BOdSfx, max(BOdDmSz) as BOdDmSz, " + _
                  "max(" + IIF(ADC.MenuCd = "RDT" Or ADC.MenuCd = "RBV", "BRepQty", "BQty") + ") as BQty, max(BPtQty) as BPtQty, max(" + IIF(ADC.MenuCd = "RDT" Or ADC.MenuCd = "RBV", "BRepGrWt", "BGrWt") + ") as BGrWt, " + _
                  "max(BRjQty) as BRjQty, max(BRjPtQty) as BRjPtQty, max(" + IIF(ADC.MenuCd = "RDT" Or ADC.MenuCd = "RBV", "BRepOdTc", "OmTc") + ") as OmTc, " + _
                  "max(" + IIF(ADC.MenuCd = "RDT" Or ADC.MenuCd = "RBV", "BRepOdYy", "OmYy") + ") as OmYy, max(" + IIF(ADC.MenuCd = "RDT" Or ADC.MenuCd = "RBV", "BRepOdChr", "OmChr") + ") as OmChr, max(" + IIF(ADC.MenuCd = "RDT" Or ADC.MenuCd = "RBV", "BRepOdNo", "OmNo") + ") as OmNo, " + _
                  "max(" + IIF(ADC.MenuCd = "RDT" Or ADC.MenuCd = "RBV", "BRepOdSr", "BOdSr") + ") as BOdSr, max(OmCmCd) as OmCmCd, max(OmPoNo) as OmPoNo, max(OdKt) as OdKt,max(OdDmCol) as OdDmCol, " + _
                  RMDET + _
                  "From Txnd Left Outer Join RmMst On RmCd= TdRmCd  " + IIF(gs_Partition = ctCurrPrtn, " And RmPrtKey=TdPrtkey ", "") + _
                  "join Bag on BIdNo=TdBIdNo And BCoCd=TdCoCd and BYy=TdBYy and BChr=TdBChr and BNo=TdBNo " + IIF(gs_Partition = ctCurrPrtn, " and BPrtKey=TdPrtKey ", "") + _
                  "join OrdMst on OmCoCd= BCoCd and OmTc= BOdTc and OmYy= BOdYy and " + _
                  "OmChr= BOdChr and OmNo= BOdNo join OrdDsg on OdIdNo=BOdIdNo And OdCoCd= BCoCd and " + _
                  "OdTc= BOdTc and OdYy= BOdYy and OdChr= BOdChr and OdNo= BOdNo and OdSr= BOdSr " + IIF(gs_Partition = ctCurrPrtn, " And OdPrtKey=BPrtKey ", "") + _
                  "Where BCoCd='" + ADC("TCoCd") + "' and " + _
                  "BYy='" + GrdTxndA.Value(wTxndARow, "TdBYy") + "' and " + _
                  "BChr='" + GrdTxndA.Value(wTxndARow, "TdBChr") + "' and " + _
                  "BNo= " + CStr(GrdTxndA.Value(wTxndARow, "TdBNo")) + " and BPrtKey=(Case When '" + gs_Partition + "'='" + ctCurrPrtn + "' Then '" + ctCurrPrtn + "' Else BPrtKey End) " + _
                  IIF(ADC.MenuCd = "RDT" Or ADC.MenuCd = "RBV", wRepairBagExtraCnd, "") + _
                  " group By BYy, BChr, BNo")
                  '(index=RmMst)  '(index=OrdMst)

  If Not (wRsBagDet.EOF Or wRsBagDet.BOF) Then
    ADC("wBdLoc") = wRsBagDet!BLoc
    ADC("wBdWrk") = wRsBagDet!BWrk
    ADC("wBdCmCd") = wRsBagDet!OmCmCd
    ADC("wBdSo") = wRsBagDet!OmTc + "/" + wRsBagDet!OmYy + "/" + wRsBagDet!OmChr + "/" + _
                   Trim(CStr(wRsBagDet!OmNo)) + "/" + Trim(CStr(wRsBagDet!BOdSr))
    ADC("wBdPoNo") = wRsBagDet!OmPoNo
    ADC("wBdDmCd") = wRsBagDet!BOdDmCd
    ADC("wBdDmSfx") = wRsBagDet!BOdSfx
    ADC("wBdDmSz") = wRsBagDet!BOdDmSz
    ADC("wBdKt") = wRsBagDet!OdKt
    ADC("wBdDmCol") = wRsBagDet!OdDmCol
    ADC("wBdBQty") = wRsBagDet!BQty
    ADC("wBdPtQty") = wRsBagDet!BPtQty
    ADC("wBdGrsWt") = wRsBagDet!BGrWt
    ADC("wBdRjQty") = wRsBagDet!BRjQty
    ADC("wBdRjPtQty") = wRsBagDet!BRjPtQty

    ADC("wBdGWt") = wRsBagDet!qBWtG
    ADC("wBdPWt") = wRsBagDet!qBWtP
    ' ****** Sachin 2.12 - 02/12/05 - SJM
    ADC("wBdSWt") = wRsBagDet!qBWtS
    ' ****** Sachin 2.12 - 02/12/05 - SJM
    ' ***** Manali 3.03 -20/06/08 - Pd Mod - 'L' - Included
    ADC("wBdLWt") = wRsBagDet!qBWtL
    ' ***** Manali 3.03 -20/06/08 - Pd Mod - 'L' - Included
    ADC("wBdMWt") = wRsBagDet!qBWtM
    ADC("wBdFndGWt") = wRsBagDet!qBWtFndG
    ADC("wBdFndPWt") = wRsBagDet!qBWtFndP
    ' ****** Sachin 2.12 - 02/12/05 - SJM
    ADC("wBdFndSWt") = wRsBagDet!qBWtFndS
    ' ****** Sachin 2.12 - 02/12/05 - SJM
    ' ***** Manali 3.03 -20/06/08 - Pd Mod - 'L' - Included
    ADC("wBdFndLWt") = wRsBagDet!qBWtFndL
    ' ***** Manali 3.03 -20/06/08 - Pd Mod - 'L' - Included
    ADC("wBdFndMWt") = wRsBagDet!qBWtFndM
    ADC("wBdDWt") = wRsBagDet!qBWtD
    ADC("wBdCWt") = wRsBagDet!qBWtC
    ADC("wBdXWt") = wRsBagDet!qBWtX
    ADC("wBdDQty") = wRsBagDet!qBQtyD
    ADC("wBdCQty") = wRsBagDet!qBQtyC
    ADC("wBdXQty") = wRsBagDet!qBQtyX
    ms_DmTcTyp = wRsBagDet!BOdDmTcTyp     'indigo.164 DmTcTyp assigned
    'indigo.164 "DM" replaced with ms_DmTcTyp
    wHskPth = GetPictPath(ms_DmTcTyp, True, ADC("wBdDmCd"), ADC("wBdDmSz"))
    w3DPth = GetPictPath(ms_DmTcTyp, False, ADC("wBdDmCd"), ADC("wBdDmSz"))
  
    Call ShowPic(wPicHskHt, wPicHskWdt, wHskPth, PicHsk)
    Call ShowPic(wPic3DHt, wPic3DWdt, w3DPth, Pic3D)
  
  End If
  Set wRsBagDet = Nothing
End Sub

Private Sub Pic3D_DblClick()
  '*** If wPic3DDblClk flag option is False then Enlarge the 3-Dimensional picture to fit the FraNKeyAll Frame
  '*** If wPic3DDblClk flag option is True then bring the 3-Dimensional picture to the original pic control size

'  If ADC("DmTcTyp") = "SM" Then Exit Sub
  If wPic3DDblClk = False Then
    '*** Enlarge The Picture ***
    ' **** Zubin 211 **** '
    'indigo.164 "DM" replaced with ms_DmTcTyp
    w3DPth = GetPictPath(ms_DmTcTyp, False, ADC("wBdDmCd"), ADC("wBdDmSz"), , , , , , , , , True)
 
    'Sachin - 4.1.0.0
    'indigo.164 "DM" replaced with ms_DmTcTyp
    If Dir(w3DPth) = "" Then w3DPth = GetPictPath(ms_DmTcTyp, False, ADC("wBdDmCd"), ADC("wBdDmSz"))
    
    Call ShowLargePic(IIF(FraBD.Height - 100 > 0, FraBD.Height - 100, 0), FraBD.Width / 2, w3DPth, Pic3D)
    FraPic3D.Left = 0: FraPic3D.Top = 0
    FraPic3D.Height = Pic3D.Height + 100: FraPic3D.Width = Pic3D.Width + 100
    FraPic3D.ZOrder (0)
  Else
    '*** Make The Picture Small ***
    'indigo.164 "DM" replaced with ms_DmTcTyp
    w3DPth = GetPictPath(ms_DmTcTyp, False, ADC("wBdDmCd"), ADC("wBdDmSz"))
'    w3DPth = GetPictPath(ADC("DmTcTyp"), False, ADC("DmCd"), ADC("DmSz"))
    'w3DPth = GetPictPath(ADC("DmTcTyp"), ADC("DmCd"), ADC("DmSz"), False)
    Call ShowPic(wPic3DHt, wPic3DWdt, w3DPth, Pic3D)
    FraPic3D.Left = wFraPic3DLft: FraPic3D.Top = wFraPic3DTop
    FraPic3D.Height = Pic3D.Height + 100: FraPic3D.Width = Pic3D.Width + 100
  End If
  If wPic3DDblClk = False Then
    wPic3DDblClk = True
  Else
    wPic3DDblClk = False
  End If
End Sub
Private Sub PicHsk_DblClick()
  '*** If wPic3DDblClk flag option is False then Enlarge the Sketch picture to fit the FraNKeyAll Frame
  '*** If wPic3DDblClk flag option is True then bring the Sketch picture to the original pic control size
  
  If wPicHskDblClk = False Then
    '*** Enlarge The Picture ***
    'indigo.164 "DM" replaced with ms_DmTcTyp
    wHskPth = GetPictPath(ms_DmTcTyp, True, ADC("wBdDmCd"), ADC("wBdDmSz"))
    
    Call ShowLargePic(IIF(FraBD.Height - 100 > 0, FraBD.Height - 100, 0), FraBD.Width / 2, wHskPth, PicHsk)
    FraPicHsk.Left = 0: FraPicHsk.Top = 0
    FraPicHsk.Height = PicHsk.Height + 100: FraPicHsk.Width = PicHsk.Width + 100
    FraPicHsk.ZOrder (0)
  Else
    '*** Make The Picture Small ***
    'indigo.164 "DM" replaced with ms_DmTcTyp
    wHskPth = GetPictPath(ms_DmTcTyp, True, ADC("wBdDmCd"), ADC("wBdDmSz"))
    
    Call ShowPic(wPicHskHt, wPicHskWdt, wHskPth, PicHsk)
    FraPicHsk.Left = wFraPicHskLft: FraPicHsk.Top = wFraPicHskTop
    FraPicHsk.Height = PicHsk.Height + 100: FraPicHsk.Width = PicHsk.Width + 100
  End If
  If wPicHskDblClk = False Then
    wPicHskDblClk = True
  Else
    wPicHskDblClk = False
  End If
End Sub

Private Sub VouchSumm(ByVal mGrpByCd As String, ByVal mValTyp As String)
  '*** Procedure to get the total Wt Issued and Received in various locations
      'for the entire voucher
  
  If GrdTxndA.Rows - 1 = 0 Then Exit Sub
  Dim wGrpByCd  As String, wGrpBy As String, wTmpTbl As String, wCommStr1 As String
  Dim wCommStr2 As String, wStr1 As String, wStr2 As String, wSqlStr As String
  Dim wStr As String, wIssTotQty As Single, wIssTotWt As Single, wRecTotQty As Single, wTotVal As Double
  Dim wRecTotWt As Single, wRow As Single
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
  
  Dim wRmQWStr, wRmRtStr, wRmValStr As String
              
  wRmQWStr = ", IsNull((Case when (RmCtg in ('G', 'P', 'S', 'L') And RmQw='W') then (Select RmPr.RmQw From RmMst RmPr " + _
              "   Where " + IIF(gs_Partition = ctCurrPrtn, " RmPr.RmPrtKey='" + ctCurrPrtn + "' and ", "") + " RmPr.RmCd=(Select PValue from param " + _
              "       Where PTyp= 'RMCTG' and PMCd= r.RmCtg)) Else RmQw End),'')  as RmQw," + _
              " IsNull((Case when (RmCtg in ('G', 'P', 'S', 'L') And RmQw='W') then (Select RmPr.RmQwCst From RmMst RmPr " + _
              "   Where " + IIF(gs_Partition = ctCurrPrtn, " RmPr.RmPrtKey='" + ctCurrPrtn + "' and ", "") + " RmPr.RmCd=(Select PValue from param " + _
              "       Where PTyp= 'RMCTG' and PMCd= r.RmCtg)) Else RmQwCst End),'') as RmQwCst "
  
  If UCase(mValTyp) = "STKVAL" Then
    wRmRtStr = "TdRmStkRt"
    wRmValStr = "Sum((Case when qRecRmWt <>0 Then qRecRmWt Else -1*qIssRmWt End)*RmRt) as RmVal "
    
  ElseIf UCase(mValTyp) = "CSTVAL" Then
    wRmRtStr = RmRate("RrCstRt", "TdRmSz", "RmQwCst")
    wRmValStr = "Sum((Case when RmQwCst='Q' Then (Case when qRecRmQty<>0 Then qRecRmQty Else -1*qIssRmQty End ) Else " + _
                                             "   (Case when qRecRmWt <>0 Then qRecRmWt Else -1*qIssRmWt End) End)*RmRt) as RmVal "
  ElseIf UCase(mValTyp) = "SALVAL" Then
    wRmRtStr = RmRate("RrSalRt", "TdRmSz", "RmQw")
    wRmValStr = "Sum((Case when RmQw='Q' Then (Case when qRecRmQty<>0 Then qRecRmQty Else -1*qIssRmQty End ) Else " + _
                                             "   (Case when qRecRmWt <>0 Then qRecRmWt Else -1*qIssRmWt End) End)*RmRt) as RmVal "
  Else
    wRmRtStr = "0.0"
    wRmValStr = " 0 as RmVal"
    
  End If
  
  If Not ms_PwdFlag Then
     wRmRtStr = "0.0"
     wRmValStr = " 0 as RmVal"
  End If
  
  wCommStr1 = "": wCommStr2 = "": wStr1 = "": wStr2 = "": wSqlStr = "": wStr = ""
  wIssTotQty = 0: wIssTotWt = 0: wRecTotQty = 0: wRecTotWt = 0
  wTmpTbl = "#TmpVS"
  '****** Sachin 3.02 22-12-07 - Corresponding Partition
  wCommStr1 = "(case when TdFrRmDc='C' then TdRmQty else 0 end) as qIssRmQty, " + _
              "(case when TdFrRmDc='C' then TdRmWt else 0 end) as qIssRmWt, " + _
              "(case when TdFrRmDc='D' then TdRmQty else 0 end) as qRecRmQty, " + _
              "(case when TdFrRmDc='D' then TdRmWt else 0 end) as qRecRmWt," + _
               wRmRtStr + " as RmRt" + wRmQWStr + " /* into " + wTmpTbl + " */  " + _
              "from Txnd Join RmMst r On RmCd= TdRmCd and RmPrtKey=(Case When '" + gs_Partition + "'='" + ctCurrPrtn + "' Then TdPrtKey Else RmPrtKey End)" + _
              "Where TdCoCd ='" + ADC("TCoCd") + "' And TdTc ='" + ADC("TTc") + "' And " + _
              "TdYy ='" + ADC("TYy") + "' And TdChr ='" + ADC("TChr") + "' And " + _
              "TdNo =" + CStr(ADC("TNo")) + " and TdSrNo>0 and TdPrtKey=(Case When '" + gs_Partition + "'='" + ctCurrPrtn + "' Then '" + ctCurrPrtn + "' Else TdPrtKey End)"
              '(index=RmMst)
  wCommStr2 = "(case when TdToRmDc='C' then TdRmQty else 0 end) as qIssRmQty, " + _
              "(case when TdToRmDc='C' then TdRmWt else 0 end) as qIssRmWt, " + _
              "(case when TdToRmDc='D' then TdRmQty else 0 end) as qRecRmQty, " + _
              "(case when TdToRmDc='D' then TdRmWt else 0 end) as qRecRmWt, " + _
               wRmRtStr + " as RmRt" + wRmQWStr + _
              "from Txnd Join RmMst r On RmCd= TdRmCd and RmPrtKey=(Case When '" + gs_Partition + "'='" + ctCurrPrtn + "' Then TdPrtKey Else RmPrtKey End)" + _
              "Where TdCoCd ='" + ADC("TCoCd") + "' And TdTc ='" + ADC("TTc") + "' And " + _
              "TdYy ='" + ADC("TYy") + "' And TdChr ='" + ADC("TChr") + "' And " + _
              "TdNo =" + CStr(ADC("TNo")) + " and TdSrNo>0 and TdPrtKey=(Case When '" + gs_Partition + "'='" + ctCurrPrtn + "' Then '" + ctCurrPrtn + "' Else TdPrtKey End)"
              '(index=RmMst)
              
 Dim wTblStr  As String
   wTblStr = "Select TdFrRmLoc as qLoc, RmCtg, RmSCtg, RmCd, TdLotNo as qLotNo, Space(8) as RmSz,TdRmQty as qIssRmQty,TdRmWt as qIssRmWt,TdRmQty as qRecRmQty,TdRmWt as qRecRmWt,TdRmStkRt as RmRt,RmQW,RmQWCst Into " + wTmpTbl + " From Txnd ,RmMst  where 1=2"
 moCn.CreateTmpTable wTmpTbl, wTblStr
                           
              
' Zubin 212
  If wGrpBy = "qLoc,RmCtg" Then
    wStr1 = "Select TdFrRmLoc as qLoc, RmCtg, '*' as RmSCtg, '*' as RmCd, '*' as qLotNo, '*' as RmSz, " + wCommStr1
    wStr2 = "Select TdToRmLoc as qLoc, RmCtg, '*' as RmSCtg, '*' as RmCd, '*' as qLotNo, '*' as RmSz, " + wCommStr2
    wSqlStr = " Select qLoc, RmCtg, max(RmSCtg) as RmSCtg, max(RmCd) as RmCd, max(qLotNo) as qLotNo, max(RmSz) as RmSz, " + _
              "sum(qIssRmQty) as qIssRmQty, sum(qIssRmWt) as qIssRmWt, " + _
              "sum(qRecRmQty) as qRecRmQty, sum(qRecRmWt) as qRecRmWt ," + wRmValStr + _
              " from " + wTmpTbl + " Group By " + wGrpBy
  ElseIf wGrpBy = "qLoc,RmCtg,RmSCtg" Then
    wStr1 = "Select TdFrRmLoc as qLoc, RmCtg, RmSCtg, '*' as RmCd, '*' as qLotNo, '*' as RmSz, " + wCommStr1
    wStr2 = "Select TdToRmLoc as qLoc, RmCtg, RmSCtg, '*' as RmCd, '*' as qLotNo, '*' as RmSz, " + wCommStr2
    wSqlStr = " Select qLoc, RmCtg, RmSCtg, max(RmCd) as RmCd, max(qLotNo) as qLotNo, max(RmSz) as RmSz, " + _
              "sum(qIssRmQty) as qIssRmQty, sum(qIssRmWt) as qIssRmWt, " + _
              "sum(qRecRmQty) as qRecRmQty, sum(qRecRmWt) as qRecRmWt  ," + wRmValStr + _
              "from " + wTmpTbl + " Group By " + wGrpBy
  ElseIf wGrpBy = "qLoc,RmCtg,RmSCtg,RmCd" Then
    wStr1 = "Select TdFrRmLoc as qLoc, RmCtg, RmSCtg, RmCd, '*' as qLotNo, '*' as RmSz, " + wCommStr1
    wStr2 = "Select TdToRmLoc as qLoc, RmCtg, RmSCtg, RmCd, '*' as qLotNo, '*' as RmSz, " + wCommStr2
    wSqlStr = " Select qLoc, RmCtg, RmSCtg, RmCd, max(qLotNo) as qLotNo, max(RmSz) as RmSz, " + _
              "sum(qIssRmQty) as qIssRmQty, sum(qIssRmWt) as qIssRmWt, " + _
              "sum(qRecRmQty) as qRecRmQty, sum(qRecRmWt) as qRecRmWt  ," + wRmValStr + _
              " from " + wTmpTbl + " Group By " + wGrpBy
  ElseIf wGrpBy = "qLoc,RmCtg,RmSCtg,RmCd,RmSz" Then
    wStr1 = "Select TdFrRmLoc as qLoc, RmCtg, RmSCtg, RmCd, '*' as qLotNo, TdRmSz as RmSz, " + wCommStr1
    wStr2 = "Select TdToRmLoc as qLoc, RmCtg, RmSCtg, RmCd, '*' as qLotNo, TdRmSz as RmSz, " + wCommStr2
    wSqlStr = " Select qLoc, RmCtg, RmSCtg, RmCd, max(qLotNo) as qLotNo, RmSz, sum(qIssRmQty) as qIssRmQty, " + _
              "sum(qIssRmWt) as qIssRmWt, sum(qRecRmQty) as qRecRmQty, " + _
              "sum(qRecRmWt) as qRecRmWt  ," + wRmValStr + "  from " + wTmpTbl + " Group By " + wGrpBy
  ' Zubin 212
  ElseIf wGrpBy = "qLoc,RmCtg,RmSCtg,RmCd,qLotNo" Then
    wStr1 = "Select TdFrRmLoc as qLoc, RmCtg, RmSCtg, RmCd, TdLotNo as qLotNo, '*' as RmSz, " + wCommStr1
    wStr2 = "Select TdToRmLoc as qLoc, RmCtg, RmSCtg, RmCd, TdLotNo as qLotNo, '*' as RmSz, " + wCommStr2
    wSqlStr = " Select qLoc, RmCtg, RmSCtg, RmCd, qLotNo, max(RmSz) as RmSz, sum(qIssRmQty) as qIssRmQty, " + _
              "sum(qIssRmWt) as qIssRmWt, sum(qRecRmQty) as qRecRmQty, " + _
              "sum(qRecRmWt) as qRecRmWt ," + wRmValStr + " from " + wTmpTbl + " Group By " + wGrpBy
  ElseIf wGrpBy = "qLoc,RmCtg,qLotNo" Then
    wStr1 = "Select TdFrRmLoc as qLoc, RmCtg, '*' as RmSCtg, '*' as RmCd, TdLotNo as qLotNo, '*' as RmSz, " + wCommStr1
    wStr2 = "Select TdToRmLoc as qLoc, RmCtg, '*' as RmSCtg, '*' as RmCd, TdLotNo as qLotNo, '*' as RmSz, " + wCommStr2
    wSqlStr = " Select qLoc, RmCtg, max(RmSCtg) as RmSCtg, max(RmCd) as RmCd, qLotNo, max(RmSz) as RmSz, sum(qIssRmQty) as qIssRmQty, " + _
              "sum(qIssRmWt) as qIssRmWt, sum(qRecRmQty) as qRecRmQty, " + _
              "sum(qRecRmWt) as qRecRmWt  ," + wRmValStr + " from " + wTmpTbl + " Group By " + wGrpBy
  ElseIf wGrpBy = "qLoc,RmCtg,RmSCtg,qLotNo" Then
    wStr1 = "Select TdFrRmLoc as qLoc, RmCtg, RmSCtg, '*' as RmCd, TdLotNo as qLotNo, '*' as RmSz, " + wCommStr1
    wStr2 = "Select TdToRmLoc as qLoc, RmCtg, RmSCtg, '*' as RmCd, TdLotNo as qLotNo, '*' as RmSz, " + wCommStr2
    wSqlStr = " Select qLoc, RmCtg, RmSCtg, max(RmCd) as RmCd, qLotNo, max(RmSz) as RmSz, sum(qIssRmQty) as qIssRmQty, " + _
              "sum(qIssRmWt) as qIssRmWt, sum(qRecRmQty) as qRecRmQty, " + _
              "sum(qRecRmWt) as qRecRmWt ," + wRmValStr + "  from " + wTmpTbl + " Group By " + wGrpBy
  ElseIf wGrpBy = "qLoc,RmCtg,RmSCtg,RmCd,RmSz,qLotNo" Then
    wStr1 = "Select TdFrRmLoc as qLoc, RmCtg, RmSCtg, RmCd, TdLotNo as qLotNo, TdRmSz as RmSz, " + wCommStr1
    wStr2 = "Select TdToRmLoc as qLoc, RmCtg, RmSCtg, RmCd, TdLotNo as qLotNo, TdRmSz as RmSz, " + wCommStr2
    wSqlStr = " Select qLoc, RmCtg, RmSCtg, RmCd, qLotNo, RmSz, sum(qIssRmQty) as qIssRmQty, " + _
              "sum(qIssRmWt) as qIssRmWt, sum(qRecRmQty) as qRecRmQty, " + _
              "sum(qRecRmWt) as qRecRmWt ," + wRmValStr + " from " + wTmpTbl + " Group By " + wGrpBy
  ' Zubin 212
  End If
  
  wStr = wStr1 + " Union All " + wStr2

  '*** Creating Temporary Table #TmpVS ***
  'moCn.CreateTmpTable wTmpTbl, wStr
  moCn.Execute " Insert Into " + wTmpTbl + " " + wStr
  
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
      .Value(wRow, "wVSRmVal") = wRsVS!RmVal
      Select Case Trim(wRsVS!RmCtg)
      ' ***** Manali 3.03 -20/06/08 - Pd Mod - 'L'
      Case Is = "G", "P", "S", "L", "M", "X"
        wIssTotWt = wIssTotWt + wRsVS!qIssRmWt
        wRecTotWt = wRecTotWt + wRsVS!qRecRmWt
      Case Is = "D", "C"
        wIssTotWt = wIssTotWt + MWLib.Div(wRsVS!qIssRmWt, 5)
        wRecTotWt = wRecTotWt + MWLib.Div(wRsVS!qRecRmWt, 5)
      End Select
      wIssTotQty = wIssTotQty + wRsVS!qIssRmQty
      wRecTotQty = wRecTotQty + wRsVS!qRecRmQty
      wTotVal = wTotVal + IIF(wRsVS!qLoc = "W", wRsVS!RmVal, 0)
      wRsVS.MoveNext
    Loop
  End With
  ADC("wVsIssTotWt") = wIssTotWt
  ADC("wVsIssTotQty") = wIssTotQty
  ADC("wVsRecTotWt") = wRecTotWt
  ADC("wVsRecTotQty") = wRecTotQty
  ADC("wVsTotVal") = wTotVal
  GrdVS.AllowAdd = False
  GrdVS.AllowDelete = False
  Set wRsVS = Nothing
End Sub
Private Sub BagRmSummary(ByVal mGrpByCd As String, ByVal mValTyp As String)
  '*** Procedure to get the details for the Bag Raw Material Summary Frame
  Dim wGrpByCd  As String, wGrpBy As String, wCommStr As String
  Dim wRmTotWt As Single, wRmTotQty As Single, wRmTotVal As Double, wStr As String, wRow As Single
  If ADC("wBsBNo") = 0 Then
    DispMsg "Invalid Bag No", etInfo
    GrdBRS.AllowAdd = True
    GrdBRS.AllowDelete = True
    GrdBRS.Rows = 1
    GrdBRS.AllowAdd = False
    GrdBRS.AllowDelete = False
    Exit Sub
  End If
  
  Dim wRepairBagExtraCnd As String
  wRepairBagExtraCnd = " and (Txnd.ModDt >=CONVERT(date,BRepINDt) or (Txnd.ModDt =CONVERT(date,BRepINDt)  and Txnd.ModTime >=cast((substring(CONVERT(VARCHAR(10),BRepINDt,108),1,2)+'.'+substring(CONVERT(VARCHAR(10),BRepINDt,108),4,2)) as float))) and TdSrNo <>0  "

  wGrpByCd = mGrpByCd
  If wGrpByCd = "RMCTG" Then wGrpBy = "RmCtg"
  If wGrpByCd = "RMSCTG" Then wGrpBy = "RmCtg,RmSCtg"
  If wGrpByCd = "RMCD" Then wGrpBy = "RmCtg,RmSCtg,RmCd"
  
  '### If wGrpByCd = "" Then wGrpBy = moCn.GetFldVal("Select hTBagBsGr from Head where HCd='" + ctSelfCmCd + "'")
  'If wGrpByCd = "" Then wGrpBy = "RmCtg"
  wRmTotWt = 0: wRmTotQty = 0: wStr = ""
  '****** Sachin 3.02 22-12-07 - Current Partition
  Dim wRmQWStr, wRmRtStr, wRmValStr As String
  
  wRmQWStr = ", IsNull((Case when (RmCtg in ('G', 'P', 'S', 'L') And RmQw='W') then (Select RmPr.RmQw From RmMst RmPr " + _
              "   Where " + IIF(gs_Partition = ctCurrPrtn, " RmPr.RmPrtKey='" + ctCurrPrtn + "' and ", "") + " RmPr.RmCd=(Select PValue from param " + _
              "       Where PTyp= 'RMCTG' and PMCd= r.RmCtg)) Else RmQw End),'')  as RmQw," + _
              " IsNull((Case when (RmCtg in ('G', 'P', 'S', 'L') And RmQw='W') then (Select RmPr.RmQwCst From RmMst RmPr " + _
              "   Where " + IIF(gs_Partition = ctCurrPrtn, " RmPr.RmPrtKey='" + ctCurrPrtn + "' and ", "") + " RmPr.RmCd=(Select PValue from param " + _
              "       Where PTyp= 'RMCTG' and PMCd= r.RmCtg)) Else RmQwCst End),'') as RmQwCst "
 
  
  If UCase(mValTyp) = "STKVAL" Then
    wRmRtStr = "TdRmStkRt"
      wRmValStr = " RmWt*RmRt "
  ElseIf UCase(mValTyp) = "CSTVAL" Then
    wRmRtStr = RmRate("RrCstRt", "TdRmSz", "RmQwCst")
      wRmValStr = " (Case when RmQWCst='Q' Then RmQty Else RmWt End)*RmRt "
  ElseIf UCase(mValTyp) = "SALVAL" Then
    wRmRtStr = RmRate("RrSalRt", "TdRmSz", "RmQw")
      wRmValStr = " (Case when RmQW='Q' Then RmQty Else RmWt End)*RmRt "
  Else
    wRmRtStr = "0.0"
      wRmValStr = " 0.0 "
  End If
  
 If Not ms_PwdFlag Then
     wRmRtStr = "0.0"
     wRmValStr = " 0.0 "
  End If
  
  wCommStr = " sum(RmWt) as RmWt ,Sum(RmQty) as RmQty,Sum(" + wRmValStr + ") as RmVal  " + _
             "  From  (Select RmCtg,RmSCtg,RmCd,(case (case when TdFrRmLoc='W' then TdFrRmDc " + _
             "when TdToRmLoc='W' then TdToRmDc else '' end) when 'C' then -1 " + _
             "when 'D' then +1 else 0 end)* TdRmQty as RmQty, " + _
             "(case (case when TdFrRmLoc='W' then TdFrRmDc when TdToRmLoc='W' then " + _
             "TdToRmDc end) when 'C' then -1 when 'D' then +1 else 0 end)*TdRmWt as RmWt, " + _
              wRmRtStr + " as RmRt " + wRmQWStr + "   From  Txnd " + _
             "Join RmMst r On RmCd= TdRmCd and RmPrtKey=(Case When '" + gs_Partition + "'='" + ctCurrPrtn + "' Then TdPrtKey Else RmPrtKey End)" + _
            " Join Bag On TdCoCd = BCoCd and TdByy = Byy and TdBChr = BChr and TdBNo = BNo " + _
             "Where TdBYy<> '' and TdBChr<> '' and TdBNo<> 0 and TdRmCd <> '' and " + _
             "TdCoCd ='" + ADC("TCoCd") + "' and " + _
             "TdBYy ='" + ADC("wBsBYy") + "' And " + _
             "TdBChr ='" + ADC("wBsBChr") + "' And " + _
             "TdBNo =" + CStr(ADC("wBsBNo")) + " And TdPrtKey=(Case When '" + gs_Partition + "'='" + ctCurrPrtn + "' Then '" + ctCurrPrtn + "' Else TdPrtKey End) " + wRepairBagExtraCnd + ") T " + _
             " Group By " + wGrpBy
             '(index=RmMst)
                
  If wGrpBy = "RmCtg" Then
    wStr = "Select RmCtg, '*' as RmSCtg, '*' as RmCd, " + wCommStr
  ElseIf wGrpBy = "RmCtg,RmSCtg" Then
    wStr = "Select RmCtg, RmSCtg, '*' as RmCd, " + wCommStr
  ElseIf wGrpBy = "RmCtg,RmSCtg,RmCd" Then
    wStr = "Select RmCtg, RmSCtg, RmCd, " + wCommStr
  End If
  
  Dim wRsBRS As MwfLib.MDORowSet
  Set wRsBRS = moCn.OpenRes(wStr)
  
  GrdBRS.AllowAdd = True
  GrdBRS.AllowDelete = True

  GrdBRS.Rows = 1
  With GrdBRS
    Do While Not (wRsBRS.EOF Or wRsBRS.BOF)
      .AddItem
      wRow = .Rows - 1
      .Value(wRow, "wBRSRmCtg") = wRsBRS!RmCtg
      .Value(wRow, "wBRSRmSCtg") = wRsBRS!RmSCtg
      .Value(wRow, "wBRSRmCd") = wRsBRS!RmCd
      .Value(wRow, "wBRSRmQty") = wRsBRS!RmQty
      .Value(wRow, "wBRSRmWt") = wRsBRS!rmwt
      .Value(wRow, "wBRSRmVal") = wRsBRS!RmVal
                                          
      Select Case Trim(wRsBRS!RmCtg)
      ' ***** Manali 3.03 -20/06/08 - Pd Mod - 'L'
      Case Is = "G", "P", "S", "L", "M", "X"
        wRmTotWt = wRmTotWt + wRsBRS!rmwt
      Case Is = "D", "C"
        wRmTotWt = wRmTotWt + MWLib.Div(wRsBRS!rmwt, 5)
      End Select
      wRmTotQty = wRmTotQty + wRsBRS!RmQty
      wRmTotVal = wRmTotVal + wRsBRS!RmVal
      wRsBRS.MoveNext
    Loop
  End With
  ADC("wBRSRmTotWt") = wRmTotWt
  ADC("wBRSRmTotQty") = wRmTotQty
  ADC("wBRSRmTotVal") = wRmTotVal
  GrdBRS.AllowAdd = False
  GrdBRS.AllowDelete = False
  Set wRsBRS = Nothing
End Sub
Private Sub LossEntry()
  '*** Procedure to enter the Rm Loss record in the 'DT' Entry
  ' ***** Manali 3.8.0 - Procedure Changed To book loss on more than 2 RmCodes
  'SRITODO
  If GrdTxndA.Rows - 1 = 0 Then Exit Sub
  Dim wTxndARow  As Single, wRow As Single, wLsLoc As String, wBGrsWt As Double, wLsTyp As String
  Dim wDmCtg As String, wWorkLsLoc As String
  If GrdTxndB.Rows - 1 > 0 Then DispMsg "Record Already Exists", etError: Exit Sub
  wTxndARow = IIF(GrdTxndA.Row, GrdTxndA.Row, 1)

  Dim wLsGrsWt As Double
  Dim wRsLs As MwfLib.MDORowSet, wRmWt As Double, wLsPer As Double
  
  ' ***** Manali 3.8.0
  Dim wTotRec As Integer, wLsWt As Double, wMainMetRmCtg As String
  Dim wRemLsWt As Double
  If UCase(ADC.MenuCd) = UCase("RDT") Then
  '****** Sachin 3.02 22-12-07 - Corresponding Partition
    wBGrsWt = moCn.GetFldVal("Select BRepGrWt from Bag where " + _
            "BCoCd='" + ADC("TCoCd") + "' and " + _
            "BYy='" + GrdTxndA.Value(wTxndARow, "TdBYy") + "' and " + _
            "BChr='" + GrdTxndA.Value(wTxndARow, "TdBChr") + "' and " + _
            "BNo= " + CStr(GrdTxndA.Value(wTxndARow, "TdBNo")) + " and BPrtKey='" + ctCurrPrtn + "'")
  Else
    wBGrsWt = moCn.GetFldVal("Select BGrWt from Bag where " + _
            "BCoCd='" + ADC("TCoCd") + "' and " + _
            "BYy='" + GrdTxndA.Value(wTxndARow, "TdBYy") + "' and " + _
            "BChr='" + GrdTxndA.Value(wTxndARow, "TdBChr") + "' and " + _
            "BNo= " + CStr(GrdTxndA.Value(wTxndARow, "TdBNo")) + " and BPrtKey='" + ctCurrPrtn + "'")
  End If
  ' Zubin 212
  
  wMainMetRmCtg = moCn.GetFldVal("Select RmCtg From RmMst Where RmCd='" + ADC("wLsMainMetCd") + "' and RmPrtKey='" + ctCurrPrtn + "'")
  wMainMetRmCtg = IIF(wMainMetRmCtg = "", moCn.GetFldVal("Select OrRmCtg From OrdRm Join Bag On OrOdIdNo=BOdIdNo And OrCoCd = BCoCd And OrTc = BOdTc And OrYy = BOdYy And " & _
                        "OrChr = BOdChr And OrNo = BOdNo And OrSr = BOdSr And OrPrtKey=BPrtKey Join RmMst On OrRmCd = RmCd And OrPrtKey=RmPrtKey " & _
                        "Where OrMainMet = 'Y' And OrRmCtg in ('G', 'P', 'S', 'L') and " & _
                        "BCoCd = '" & ADC("TCoCd") & "' And BYy = '" & GrdTxndA.Value(GrdTxndA.Row, "TdBYy") & "' And " & _
                        "BChr = '" & GrdTxndA.Value(GrdTxndA.Row, "TdBChr") & "' And BNo = " & GrdTxndA.Value(GrdTxndA.Row, "TdBNo") & " And BPrtKey='" + ctCurrPrtn + "' "), wMainMetRmCtg)
  
  Dim wCndY As String, wCndN As String, wCnd As String
                                      
  wCndY = "": wCndN = "": wCnd = ""
  If ADC("wLsMainMetAllYn") = "Y" Then
    wCndY = " (OrRmCtg = '" + wMainMetRmCtg + "') "
    'wCndY = " (OrMainMet = 'Y') "
  Else
    wCndY = " (OrRmCd ='" + ADC("wLsMainMetCd") + "') "
  End If
                                        
  If ADC("wLsSecMetAllYn") = "Y" Then
    wCndN = " (OrRmCtg In ('G','P','S','L') and OrRmCtg <> '" + wMainMetRmCtg + "') "
    'wCndN = " (OrMainMet = 'N') "
  ElseIf ADC("wLsSecMetCd") <> "" Then
    wCndN = " (OrRmCd ='" + ADC("wLsSecMetCd") + "') "
  Else
    wCndN = ""
  End If
                                      
  wCnd = " and  (" + wCndY + IIF(wCndN <> "", " Or ", "") + wCndN + ") "
  
   
  '****** Sachin 3.02 22-12-07 - Corresponding Partition
  Set wRsLs = moCn.OpenRes("Select max(BOdDmCd) as qDmCd, Max(OrMainMet) as OrMainMet, Max(OrRmCtg) as OrRmCtg,  " + _
              "TdRmCd as qRmCd, min(TdLotNo) as qLotNo, min(TdRmSz) as qRmSz,min(TdRmSz2) as qRmSz2,min(TdRmSz3) as qRmSz3, " + _
              "min(TdRmStkRt) as qRmStkRt, 0 as qRmQty, " + _
              "Cast(sum(TdRmWt) as Decimal(14,3)) as qRmWt From " + _
              "(Select TdPrtKey, TdBIdNo, TdCoCd, TdBYy, TdBChr, TdBNo, TdRmCd, TdLotNo , TdRmSz ,TdRmSz2,TdRmSz3, " + _
              "TdRmStkRt , 0 as qRmQty, Cast(sum((case (case when TdFrRmLoc='W' then TdFrRmDc when TdToRmLoc='W' then " + _
              " TdToRmDc end) when 'C' then -1 when 'D' then +1 else 0 end)* TdRmWt) as Decimal(14,3)) as TdRmWt " + _
              " From Txnd Group By TdPrtKey, TdBIdNo, TdCoCd, TdBYy, TdBChr, TdBNo, TdRmCd, TdLotNo, TdRmSz, TdRmSz2, TdRmSz3, TdRmStkRt " + _
              " Having Cast(sum((case (case when TdFrRmLoc='W' then TdFrRmDc when TdToRmLoc='W' then " + _
              " TdToRmDc end) when 'C' then -1 when 'D' then +1 else 0 end)* TdRmWt) as Decimal(14,3)) > 0 )Txnd " + _
              "join Bag on BIdNo=TdBIdNo And BCoCd= TdCoCd and " + _
              "BYy= TdBYy and BChr= TdBChr and BNo= TdBNo and BPrtKey=TdPrtKey join OrdRm on OrOdIdNo=BOdIdNo And OrCoCd= BCoCd and " + _
              "OrTc= BOdTc and OrYy= BOdYy and OrChr= BOdChr and OrNo= BOdNo and OrSr= BOdSr and OrPrtKey=BPrtKey " + _
              "Where TdRmCd= OrRmCd and " + _
              "TdCoCd='" + ADC("TCoCd") + "' and " + _
              "TdBYy= '" + GrdTxndA.Value(wTxndARow, "TdBYy") + "' and " + _
              "TdBChr= '" + GrdTxndA.Value(wTxndARow, "TdBChr") + "' and " + _
              "TdBNo= " + CStr(GrdTxndA.Value(wTxndARow, "TdBNo")) + " and TdPrtKey='" + ctCurrPrtn + "' " + wCnd + _
              "group By TdRmCd   " + _
              "Having Cast(sum(TdRmWt) as Decimal(14,3))  > 0 " + _
              " Order By (Case When Max(OrMainMet)='Y' Then 1 Else 2 End), (Case When max(OrRmCtg)= '" + wMainMetRmCtg + "' Then 1 Else 2 End) , TdRmCd ")
  
'  Set wRsLs = moCn.OpenRes("Select max(BOdDmCd) as qDmCd, Max(OrMainMet) as OrMainMet, Max(OrRmCtg) as OrRmCtg,  " + _
'              "TdRmCd as qRmCd, min(TdLotNo) as qLotNo, min(TdRmSz) as qRmSz, " + _
'              "min(TdRmStkRt) as qRmStkRt, 0 as qRmQty, " + _
'              "sum((case (case when TdFrRmLoc='W' then TdFrRmDc when TdToRmLoc='W' then " + _
'              "TdToRmDc end) when 'C' then -1 when 'D' then +1 else 0 end)* " + _
'              "TdRmWt) as qRmWt From Txnd join Bag on BIdNo=TdBIdNo And BCoCd= TdCoCd and " + _
'              "BYy= TdBYy and BChr= TdBChr and BNo= TdBNo and BPrtKey=TdPrtKey join OrdRm on OrOdIdNo=BOdIdNo And OrCoCd= BCoCd and " + _
'              "OrTc= BOdTc and OrYy= BOdYy and OrChr= BOdChr and OrNo= BOdNo and OrSr= BOdSr and OrPrtKey=BPrtKey " + _
'              "Where TdRmCd= OrRmCd and " + _
'              "TdCoCd='" + adc("TCoCd") + "' and " + _
'              "TdBYy= '" + GrdTxndA.Value(wTxndARow, "TdBYy") + "' and " + _
'              "TdBChr= '" + GrdTxndA.Value(wTxndARow, "TdBChr") + "' and " + _
'              "TdBNo= " + CStr(GrdTxndA.Value(wTxndARow, "TdBNo")) + " and TdPrtKey='" + ctCurrPrtn + "' " + wCnd + _
'              "group By TdRmCd  /*, TdLotNo, TdRmSz, TdRmStkRt*/ /*order By TdRmCd , TdLotNo, TdRmSz, TdRmStkRt*/ " + _
'              "Having sum((case (case when TdFrRmLoc='W' then TdFrRmDc when TdToRmLoc='W' then " + _
'              "TdToRmDc end) when 'C' then -1 when 'D' then +1 else 0 end)* " + _
'              "TdRmWt) > 0 " + _
'              " Order By (Case When Max(OrMainMet)='Y' Then 1 Else 2 End), (Case When max(OrRmCtg)= '" + wMainMetRmCtg + "' Then 1 Else 2 End) , TdRmCd ")
              
  
  If wRsLs.RecCount > 0 Then
  
    wTotRec = wRsLs.RecCount
    wLsLoc = moCn.GetFldVal("Select LocLoss from Loc where LocCoCd= '" + ADC("TCoCd") + "' and " + _
                        "LocCd= '" + GrdTxndA.Value(wTxndARow, "TdFrBLoc") + "'")
    wLsTyp = moCn.GetFldVal("Select L2.LocLsTyp from Loc L1 JOIN Loc L2 ON L2.LocCoCd= L1.LocCoCd and L2.LocTyp = 'R' and L2.LocCd=L1.LocPrcs " + _
                        " where L1.LocCoCd= '" + ADC("TCoCd") + "' and " + _
                        "L1.LocCd= '" + GrdTxndA.Value(wTxndARow, "TdFrBLoc") + "'")

    wWorkLsLoc = moCn.GetFldVal("Select  vPValue1 from vParam where vPTyp= 'WORK' and " + _
                                "vPCoCd ='" + ADC("TCoCd") + "' and " + _
                        "vPMCd= '" + GrdTxndA.Value(wTxndARow, "TdWrk") + "' and vPSCd=''")
    
    wRsLs.MoveFirst
    Do While Not (wRsLs.EOF Or wRsLs.BOF)
        wRmWt = wRmWt + Round(CDbl(wRsLs!qRmWt), 3)
    wRsLs.MoveNext
    Loop
      
    wRsLs.MoveFirst
    Dim wDmTcTyp As String
    'indigo.164 getting BOdDmTcTyp
    wDmTcTyp = moCn.GetFldVal("Select BOdDmTcTyp from Bag where " + _
              "BCoCd='" + ADC("TCoCd") + "' and " + _
              "BYy='" + GrdTxndA.Value(wTxndARow, "TdBYy") + "' and " + _
              "BChr='" + GrdTxndA.Value(wTxndARow, "TdBChr") + "' and " + _
              "BNo= " + CStr(GrdTxndA.Value(wTxndARow, "TdBNo")) + " and BPrtKey='" + ctCurrPrtn + "'")
                                                
    'indigo.164 DmTcTyp='DM'  changed as DmTcTyp='" + wDmTcTyp + "'
    wDmCtg = moCn.GetFldVal("Select DmCtg From DsgMst where DmTcTyp='" + wDmTcTyp + "' and DmCd='" + wRsLs!qDmCd + "' and DmSz='' And DmPrtKey='" + ctCurrPrtn + "' ")
    
    wLsGrsWt = Round(wBGrsWt, 3) - ADC("wLsGrsWt")
    
    If Round(wBGrsWt, 3) - Round(ADC("wLsGrsWt"), 3) < 0 Then DispMsg "Gross Wt More Than Issue Wt", etError: Exit Sub
    If (wRmWt - wLsGrsWt) < 0 Then DispMsg "Loss Wt More Than Issue Wt", etError: Exit Sub
    
    'If wRsLs!OrMainMet = "N" Then DispMsg "No Main Metal Issued", etError: Exit Sub
    
    wRemLsWt = Round(wLsGrsWt, 3)
    
    GrdTxndB.StartCopy
    MWLib.BeginProcess Me, "Entering The Loss Records ..."
    With GrdTxndB
      'CHANGE BY URMI TO DO THE LOSS ENTRY IN MIN(STKRT) RECORD 2.05
      'Do While Not (wRsLs.EOF Or wRsLs.BOF)
      'If wRmWt <> 0 And Not (wRsLs.EOF Or wRsLs.BOF) Then
      Do While Not (wRsLs.EOF Or wRsLs.BOF)
        .AddItem
        wLsWt = Round(MWLib.Div(wLsGrsWt * wRsLs!qRmWt, wRmWt), 3)
        wLsPer = GetLsPer(ADC("TCoCd"), GrdTxndA.Value(wTxndARow, "TdFrBLoc"), wRsLs!OrRmCtg, wDmCtg, wRsLs!qDmCd, GrdTxndA.Value(wTxndARow, "TdWrk"), wRsLs!qRmWt - wLsWt)
         wRow = .Rows - 1
        .Value(wRow, "TdSrNo") = wRow
        If (ADC.MenuCd = "RDT") Then
          .Value(wRow, "TdFrRmLoc") = "XR"
        Else
          .Value(wRow, "TdFrRmLoc") = "W"
        End If
        .Value(wRow, "wTdFrRmIr") = "I"
        .Value(wRow, "TdFrRmDc") = "C"
        .Value(wRow, "TdRmCd") = wRsLs!qRmCd
        .Value(wRow, "TdLotNo") = wRsLs!qLotNo
        .Value(wRow, "TdRmSz") = wRsLs!qRmSz
        .Value(wRow, "TdRmSz2") = wRsLs!qRmSz2
        .Value(wRow, "TdRmSz3") = wRsLs!qRmSz3
        .Value(wRow, "TdRmStkRt") = wRsLs!qRmStkRt
        .Value(wRow, "TdRmQty") = wRsLs!qRmQty
        .Value(wRow, "TdDustWt") = ADC("wLsDustWt")
        .Value(wRow, "TdToRmLoc") = wLsLoc
        .Value(wRow, "TdByLoc") = GrdTxndA.Value(wTxndARow, "TdFrBLoc")
        .Value(wRow, "TdByWrk") = GrdTxndA.Value(wTxndARow, "TdWrk")
        .Value(wRow, "TdPrtKey") = ctCurrPrtn
        
        If Not (wLsTyp <> "" And wWorkLsLoc <> "" And wLsPer <> 0) Then
           If wRow < wTotRec Then
             .Value(wRow, "TdRmWt") = wLsWt
           Else
              .Value(wRow, "TdRmWt") = wRemLsWt
           End If
           wRemLsWt = wRemLsWt - wLsWt
          If Not .SaveRec(wRow) Then
            Exit Do
            .DelAllRecs
          End If
        Else
          'To intentionally not change the existing formula Sri is putting IFs around it
          'and still retaining the senseless IIF inside of the IF
          'IG.49 - LossTyp LJ will do percentage calculation based on Rm Wt of that row instead of
          'Total Out RmWt
          If wLsTyp = "LL" Then
            .Value(wRow, "TdRmWt") = IIF(wLsTyp = "LL", wLsWt, _
                               (wRmWt - wLsWt)) * (wLsPer / 100)
          ElseIf wLsTyp = "LW" Then
            .Value(wRow, "TdRmWt") = IIF(wLsTyp = "LL", wLsWt, _
                               (wRmWt - wLsWt)) * (wLsPer / 100)
          ElseIf wLsTyp = "LJ" Then
            .Value(wRow, "TdRmWt") = IIF(wLsTyp = "LL", wLsWt, _
                               (wRsLs!qRmWt - wLsWt)) * (wLsPer / 100)
          End If
          
          
          If Not .SaveRec(wRow) Then
            Exit Do
            .DelAllRecs
          End If
          If (.Value(wRow, "TdRmWt") - wLsWt) <> 0 Then
            .AddItem
             wRow = wRow + 1
            .Value(wRow, "TdSrNo") = wRow
            If (ADC.MenuCd = "RDT") Then
              .Value(wRow, "TdFrRmLoc") = "XR"
            Else
              .Value(wRow, "TdFrRmLoc") = "W"
            End If
            .Value(wRow, "wTdFrRmIr") = "I"
            .Value(wRow, "TdFrRmDc") = "C"
            .Value(wRow, "TdRmCd") = wRsLs!qRmCd
            .Value(wRow, "TdLotNo") = wRsLs!qLotNo
            .Value(wRow, "TdRmSz") = wRsLs!qRmSz
            .Value(wRow, "TdRmSz2") = wRsLs!qRmSz2
            .Value(wRow, "TdRmSz3") = wRsLs!qRmSz3
            .Value(wRow, "TdRmStkRt") = wRsLs!qRmStkRt
            .Value(wRow, "TdRmQty") = wRsLs!qRmQty
            .Value(wRow, "TdRmWt") = wLsWt - .Value(wRow - 1, "TdRmWt")
            .Value(wRow, "TdDustWt") = 0#
            .Value(wRow, "TdToRmLoc") = IIF(.Value(wRow, "TdRmWt") >= 0, wWorkLsLoc, wLsLoc)
            .Value(wRow, "TdByLoc") = GrdTxndA.Value(wTxndARow, "TdFrBLoc")
            .Value(wRow, "TdByWrk") = GrdTxndA.Value(wTxndARow, "TdWrk")
            .Value(wRow, "TdPrtKey") = ctCurrPrtn
            If Not .SaveRec(wRow) Then
              Exit Do
              .DelAllRecs
            End If
          End If
        End If
        wRsLs.MoveNext
      Loop
    End With
  End If

  MWLib.EndProcess Me
  GrdTxndB.EndCopy
  Set wRsLs = Nothing
  '*** End of the Loss Entry

  '*** Calculation Of TdBGrsWt ***
  '****** Sachin 3.02 22-12-07 - Current Partition
  If UCase(ADC.MenuCd) = UCase("RDT") Then
    GrdTxndA.Value(wTxndARow, "TdBGrWt") = moCn.GetFldVal("Select BRepGrWt from Bag where " + _
                                         "BCoCd='" + ADC("TCoCd") + "' and " + _
                                         "BYy='" + GrdTxndA.Value(wTxndARow, "TdBYy") + "' and " + _
                                         "BChr='" + GrdTxndA.Value(wTxndARow, "TdBChr") + "' and " + _
                                         "BNo= " + CStr(GrdTxndA.Value(wTxndARow, "TdBNo")) + " and BPrtKey='" + ctCurrPrtn + "'")
  Else
    GrdTxndA.Value(wTxndARow, "TdBGrWt") = moCn.GetFldVal("Select BGrWt from Bag where " + _
                                         "BCoCd='" + ADC("TCoCd") + "' and " + _
                                         "BYy='" + GrdTxndA.Value(wTxndARow, "TdBYy") + "' and " + _
                                         "BChr='" + GrdTxndA.Value(wTxndARow, "TdBChr") + "' and " + _
                                         "BNo= " + CStr(GrdTxndA.Value(wTxndARow, "TdBNo")) + " and BPrtKey='" + ctCurrPrtn + "'")
  End If
  GrdTxndA.SaveRec (wTxndARow)
End Sub

Private Sub LossEntry_Bef380()
  '*** Procedure to enter the Rm Loss record in the 'DT' Entry

  If GrdTxndA.Rows - 1 = 0 Then Exit Sub
  Dim wTxndARow  As Single, wRow As Single, wLsLoc As String, wBGrsWt As Double, wLsTyp As String
  Dim wDmCtg As String, wWorkLsLoc As String
  If GrdTxndB.Rows - 1 > 0 Then DispMsg "Record Already Exists", etError: Exit Sub
  wTxndARow = IIF(GrdTxndA.Row, GrdTxndA.Row, 1)
  'Manoj
  Dim wLsGrsWt As Double, wLsWtY As Double, wLsWtN As Double
  Dim wRsLs As MwfLib.MDORowSet, wRmWt As Double, wRmCtg As String, wLsPer As Double
  Dim wRsLsN As MwfLib.MDORowSet, wRmWtN As Double, wRmCtgN As String, wLsPerN As Double
  'Manoj
  '****** Sachin 3.02 22-12-07 - Corresponding Partition
  wBGrsWt = moCn.GetFldVal("Select BGrWt from Bag where " + _
            "BCoCd='" + ADC("TCoCd") + "' and " + _
            "BYy='" + GrdTxndA.Value(wTxndARow, "TdBYy") + "' and " + _
            "BChr='" + GrdTxndA.Value(wTxndARow, "TdBChr") + "' and " + _
            "BNo= " + CStr(GrdTxndA.Value(wTxndARow, "TdBNo")) + " and BPrtKey='" + ctCurrPrtn + "'")
  ' Zubin 212
  
  '****** Sachin 3.02 22-12-07 - Corresponding Partition
  Set wRsLs = moCn.OpenRes("Select max(BOdDmCd) as qDmCd,TdRmCd as qRmCd, TdLotNo as qLotNo, TdRmSz as qRmSz,TdRmSz2 as qRmSz2, TdRmSz3 as qRmSz3, " + _
              "TdRmStkRt as qRmStkRt, 0 as qRmQty, " + _
              "sum((case (case when TdFrRmLoc='W' then TdFrRmDc when TdToRmLoc='W' then " + _
              "TdToRmDc end) when 'C' then -1 when 'D' then +1 else 0 end)* " + _
              "TdRmWt) as qRmWt From Txnd join Bag on BIdNo=TdBIdNo And BCoCd= TdCoCd and " + _
              "BYy= TdBYy and BChr= TdBChr and BNo= TdBNo and BPrtKey=TdPrtKey join OrdRm on OrOdIdNo=BOdIdNo And OrCoCd= BCoCd and " + _
              "OrTc= BOdTc and OrYy= BOdYy and OrChr= BOdChr and OrNo= BOdNo and OrSr= BOdSr and OrPrtKey=BPrtKey " + _
              "Where OrMainMet= 'Y' and TdRmCd= OrRmCd and " + _
              "TdCoCd='" + ADC("TCoCd") + "' and " + _
              "TdBYy= '" + GrdTxndA.Value(wTxndARow, "TdBYy") + "' and " + _
              "TdBChr= '" + GrdTxndA.Value(wTxndARow, "TdBChr") + "' and " + _
              "TdBNo= " + CStr(GrdTxndA.Value(wTxndARow, "TdBNo")) + " and TdPrtKey='" + ctCurrPrtn + "' " + _
              "group By TdRmCd, TdLotNo, TdRmSz, TdRmSz2, TdRmSz3, TdRmStkRt order By TdRmCd, TdLotNo, TdRmSz, TdRmSz2, TdRmSz3, TdRmStkRt ")
              
  'Manoj 2.10.0
  wRmWtN = 0
  wRmCtgN = ""
  If ADC("wLsSecMetCd") <> "" Then
    ' Zubin 212
    ' ****** Sachin 3.02 (14/09/07) Id fields added
    Set wRsLsN = moCn.OpenRes("Select max(BOdDmCd) as qDmCd, TdRmCd as qRmCd, TdLotNo as qLotNo, TdRmSz as qRmSz, " + _
              "TdRmStkRt as qRmStkRt, 0 as qRmQty, " + _
              "sum((case (case when TdFrRmLoc='W' then TdFrRmDc when TdToRmLoc='W' then " + _
              "TdToRmDc end) when 'C' then -1 when 'D' then +1 else 0 end)* " + _
              "TdRmWt) as qRmWt From Txnd join Bag on BIdNo=TdBIdNo And BCoCd= TdCoCd and " + _
              "BYy= TdBYy and BChr= TdBChr and BNo= TdBNo and BPrtKey=TdPrtKey join OrdRm on OrOdIdNo=BOdIdNo And OrCoCd= BCoCd and " + _
              "OrTc= BOdTc and OrYy= BOdYy and OrChr= BOdChr and OrNo= BOdNo and OrSr= BOdSr and OrPrtKey=BPrtKey " + _
              "Where OrMainMet= 'N' and OrRmCd= '" & ADC("wLsSecMetCd") & "' and TdRmCd= OrRmCd and " + _
              "TdCoCd='" + ADC("TCoCd") + "' and " + _
              "TdBYy= '" + GrdTxndA.Value(wTxndARow, "TdBYy") + "' and " + _
              "TdBChr= '" + GrdTxndA.Value(wTxndARow, "TdBChr") + "' and " + _
              "TdBNo= " + CStr(GrdTxndA.Value(wTxndARow, "TdBNo")) + " and TdPrtKey='" + ctCurrPrtn + "' " + _
              "group By TdRmCd, TdLotNo, TdRmSz, TdRmStkRt order By TdRmCd, TdLotNo, TdRmSz, TdRmStkRt ")
    If wRsLsN.RecCount > 0 Then
      wRmWtN = Round(CDbl(wRsLsN!qRmWt), 3)
      wRmCtgN = moCn.GetFldVal("Select RmCtg From RmMst Where RmCd = '" & wRsLsN!qRmCd & "' And RmPrtKey='" + ctCurrPrtn + "'")
    End If
  End If
  'Manoj 2.10.0
  
  If wRsLs.RecCount > 0 Or ADC("wLsSecMetCd") <> "" Then
  
    'Manoj 2.10.0
    wLsLoc = moCn.GetFldVal("Select LocLoss from Loc where LocCoCd= '" + ADC("TCoCd") + "' and " + _
                        "LocCd= '" + GrdTxndA.Value(wTxndARow, "TdFrBLoc") + "'")
    wLsTyp = moCn.GetFldVal("Select L2.LocLsTyp from Loc L1 JOIN Loc L2 ON L2.LocCoCd= L1.LocCoCd and L2.LocTyp = 'R' and L2.LocCd=L1.LocPrcs " + _
                        " where L1.LocCoCd= '" + ADC("TCoCd") + "' and " + _
                        "L1.LocCd= '" + GrdTxndA.Value(wTxndARow, "TdFrBLoc") + "'")
             
    wWorkLsLoc = moCn.GetFldVal("Select  vPValue1 from vParam where vPTyp= 'WORK' and " + _
                                "vPCoCd ='" + ADC("TCoCd") + "' and " + _
                        "vPMCd= '" + GrdTxndA.Value(wTxndARow, "TdWrk") + "' and vPSCd=''")
    
    'indigo.164 getting  BOdDmTcTyp
    Dim wDmTcTyp As String
    wDmTcTyp = moCn.GetFldVal("Select BOdDmTcTyp from Bag where " + _
              "BCoCd='" + ADC("TCoCd") + "' and " + _
              "BYy='" + GrdTxndA.Value(wTxndARow, "TdBYy") + "' and " + _
              "BChr='" + GrdTxndA.Value(wTxndARow, "TdBChr") + "' and " + _
              "BNo= " + CStr(GrdTxndA.Value(wTxndARow, "TdBNo")) + " and BPrtKey='" + ctCurrPrtn + "'")
           
               
    If wRsLs.RecCount > 0 Then
      '****** Sachin 3.02 22-12-07 - Current Partition
      wRmCtg = moCn.GetFldVal("Select RmCtg From RmMst Where RmCd = '" & wRsLs!qRmCd & "' and RmPrtKey='" + ctCurrPrtn + "'")
      wRmWt = Round(CDbl(wRsLs!qRmWt), 3)
      '****** Sachin 3.02 22-12-07 - Current Partition
      'indigo.164 DmTcTyp='DM'  changed as DmTcTyp='" + wDmTcTyp + "'
      wDmCtg = moCn.GetFldVal("Select DmCtg From DsgMst where DmTcTyp='" + wDmTcTyp + "' and DmCd='" + wRsLs!qDmCd + "' and DmSz='' And DmPrtKey='" + ctCurrPrtn + "' ")
    End If
    wLsGrsWt = Round(wBGrsWt, 3) - ADC("wLsGrsWt")
    wLsWtY = Round(MWLib.Div(wLsGrsWt * wRmWt, (wRmWt + wRmWtN)), 3)
    wLsWtN = Round(wLsGrsWt - wLsWtY, 3)
    If ADC("wLsSecMetCd") <> "" Then
      '****** Sachin 3.02 22-12-07 - Current Partition
      'indigo.164 DmTcTyp='DM'  changed as DmTcTyp='" + wDmTcTyp + "'
      If wDmCtg = "" And wRsLsN.RecCount > 0 Then wDmCtg = moCn.GetFldVal("Select DmCtg From DsgMst where DmTcTyp='" + wDmTcTyp + "' and DmCd='" + wRsLsN!qDmCd + "' and DmSz='' and DmPrtKey='" + ctCurrPrtn + "'")
    End If
      
    'Manoj 2.10.0

    'If ADC("wLsGrsWt") = 0 Then DispMsg "Enter The New Gross Wt", etError: Exit Sub
    If Round(wBGrsWt, 3) - Round(ADC("wLsGrsWt"), 3) < 0 Then DispMsg "Gross Wt More Than Issue Wt", etError: Exit Sub
  'Manoj 2.10.0
    'If wRsLs!qRmWt - (wBGrsWt - ADC("wLsGrsWt")) < 0 Then DispMsg "Loss Wt More Than Issue Wt", etError: Exit Sub
    If ((wRmWt + wRmWtN) - wLsGrsWt) < 0 Then DispMsg "Loss Wt More Than Issue Wt", etError: Exit Sub
  'End If
  'Manoj 2.10.0
    
    GrdTxndB.StartCopy
    MWLib.BeginProcess Me, "Entering The Loss Records ..."
    With GrdTxndB
      'CHANGE BY URMI TO DO THE LOSS ENTRY IN MIN(STKRT) RECORD 2.05
      'Do While Not (wRsLs.EOF Or wRsLs.BOF)
      If wRmWt <> 0 And Not (wRsLs.EOF Or wRsLs.BOF) Then
        .AddItem
        '*****Geeta****Emr208
        
        'Manoj 2.10.0
        wLsPer = GetLsPer(ADC("TCoCd"), GrdTxndA.Value(wTxndARow, "TdFrBLoc"), wRmCtg, wDmCtg, wRsLs!qDmCd, GrdTxndA.Value(wTxndARow, "TdWrk"), wRmWt - wLsWtY)
'        wLsPer = moCn.GetFldVal("Select " + _
'                                  "IsNull((select  LmLossPer from LossMst " + _
'                                            "where LmCoCd='" + ADC("TCoCd") + "' and  " + _
'                                            "LmLocCd='" + GrdTxndA.Value(wTxndARow, "TdFrBLoc") + "'" + _
'                                            "and LmDmCtg='" + wDmCtg + "' and LmDmCd ='" + wRsLs!qDmCd + "' " + _
'                                            "and LmWorkCd='" + GrdTxndA.Value(wTxndARow, "TdWrk") + "' and LmLossPer<> 0)," + _
'                                  "IsNull((select  LmLossPer from LossMst " + _
'                                           "where LmCoCd='" + ADC("TCoCd") + "' " + _
'                                           "and LmLocCd= '" + GrdTxndA.Value(wTxndARow, "TdFrBLoc") + "' " + _
'                                           "and LmDmCtg='" + wDmCtg + "' and LmDmCd ='" + wRsLs!qDmCd + "' " + _
'                                           "and LmWorkCd=''  and LmLossPer<> 0), " + _
'                                  "IsNull((select  LmLossPer from LossMst " + _
'                                           "where LmCoCd='" + ADC("TCoCd") + "' " + _
'                                           "and LmLocCd='" + GrdTxndA.Value(wTxndARow, "TdFrBLoc") + "' " + _
'                                           "and LmDmCtg='" + wDmCtg + "' and LmDmCd ='' and LmWorkCd='' and LmLossPer<> 0)," + _
'                                  "IsNull((select  LmLossPer from LossMst " + _
'                                           "where LmCoCd='" + ADC("TCoCd") + "' " + _
'                                           "and LmLocCd='" + GrdTxndA.Value(wTxndARow, "TdFrBLoc") + "' " + _
'                                           "and LmDmCtg='' and LmDmCd ='' and LmWorkCd=''  and LmLossPer<> 0)" + _
'                                           " ,0)" + _
'                                           "))) ")

         wRow = .Rows - 1
        .Value(wRow, "TdSrNo") = wRow
        .Value(wRow, "TdFrRmLoc") = "W"
        .Value(wRow, "wTdFrRmIr") = "I"
        .Value(wRow, "TdFrRmDc") = "C"
        .Value(wRow, "TdRmCd") = wRsLs!qRmCd
        ' Zubin 212
        .Value(wRow, "TdLotNo") = wRsLs!qLotNo
        ' Zubin 212
        .Value(wRow, "TdRmSz") = wRsLs!qRmSz
        .Value(wRow, "TdRmStkRt") = wRsLs!qRmStkRt
        .Value(wRow, "TdRmQty") = wRsLs!qRmQty
        .Value(wRow, "TdDustWt") = ADC("wLsDustWt")
        .Value(wRow, "TdToRmLoc") = wLsLoc
        .Value(wRow, "TdByLoc") = GrdTxndA.Value(wTxndARow, "TdFrBLoc")
        .Value(wRow, "TdByWrk") = GrdTxndA.Value(wTxndARow, "TdWrk")
        .Value(wRow, "TdPrtKey") = ctCurrPrtn '****** Sachin 3.02
        'Manoj 2.10.0
'        If Not (wLsTyp <> "" And wWorkLsLoc <> "" And wLsPer <> 0) Then
'          .Value(wRow, "TdRmWt") = (wBGrsWt - ADC("wLsGrsWt")) * wRmMulVal
'          .SaveRec (wRow)
'        Else
'          .Value(wRow, "TdRmWt") = IIF(wLsTyp = "LL", (wBGrsWt - ADC("wLsGrsWt")), _
'                               (wRsLs!qRmWt - (wBGrsWt - ADC("wLsGrsWt")))) * (wLsPer / 100)
'          .SaveRec (wRow)
'          If Round((.Value(wRow, "TdRmWt") - (wBGrsWt - ADC("wLsGrsWt"))), 4) <> 0 Then
        If Not (wLsTyp <> "" And wWorkLsLoc <> "" And wLsPer <> 0) Then
          .Value(wRow, "TdRmWt") = wLsWtY
          .SaveRec (wRow)
        Else
          .Value(wRow, "TdRmWt") = IIF(wLsTyp = "LL", wLsWtY, _
                               (wRmWt - wLsWtY)) * (wLsPer / 100)
          .SaveRec (wRow)
          If (.Value(wRow, "TdRmWt") - Round(wLsWtY, 3)) <> 0 Then
        'Manoj 2.10.0
            .AddItem
             wRow = wRow + 1
            .Value(wRow, "TdSrNo") = wRow
            .Value(wRow, "TdFrRmLoc") = "W"
            .Value(wRow, "wTdFrRmIr") = "I"       'IIF(.Value(wRow - 1, "TdRmWt") < (wBGrsWt - adc("wLsGrsWt")), "I", "R")
            .Value(wRow, "TdFrRmDc") = "C"        'IIF(.Value(wRow - 1, "TdRmWt") < (wBGrsWt - adc("wLsGrsWt")), "C", "D")
            .Value(wRow, "TdRmCd") = wRsLs!qRmCd
            ' Zubin 212
            .Value(wRow, "TdLotNo") = wRsLs!qLotNo
            ' Zubin 212
            .Value(wRow, "TdRmSz") = wRsLs!qRmSz
            .Value(wRow, "TdRmStkRt") = wRsLs!qRmStkRt
            .Value(wRow, "TdRmQty") = wRsLs!qRmQty
            'Manoj 2.10.0
'            .Value(wRow, "TdRmWt") = (wBGrsWt - ADC("wLsGrsWt")) - .Value(wRow - 1, "TdRmWt") 'Abs((wBGrsWt - adc("wLsGrsWt")) - .Value(wRow - 1, "TdRmWt"))
            .Value(wRow, "TdRmWt") = wLsWtY - .Value(wRow - 1, "TdRmWt") 'Abs((wBGrsWt - adc("wLsGrsWt")) - .Value(wRow - 1, "TdRmWt"))
            'Manoj 2.10.0
            .Value(wRow, "TdDustWt") = 0#
            .Value(wRow, "TdToRmLoc") = wWorkLsLoc
            .Value(wRow, "TdByLoc") = GrdTxndA.Value(wTxndARow, "TdFrBLoc")
            .Value(wRow, "TdByWrk") = GrdTxndA.Value(wTxndARow, "TdWrk")
            .Value(wRow, "TdPrtKey") = ctCurrPrtn '****** Sachin 3.02
            .SaveRec (wRow)
          End If
        '*****Geeta****Emr208
        End If
      End If
        
      'Manoj 2.10.0
      If wRmWtN <> 0 And ADC("wLsSecMetCd") <> "" Then
        wLsPerN = GetLsPer(ADC("TCoCd"), GrdTxndA.Value(wTxndARow, "TdFrBLoc"), wRmCtgN, wDmCtg, wRsLsN!qDmCd, GrdTxndA.Value(wTxndARow, "TdWrk"), wRmWtN - wLsWtN)
        .AddItem
        wRow = wRow + 1
        .Value(wRow, "TdSrNo") = wRow
        .Value(wRow, "TdFrRmLoc") = "W"
        .Value(wRow, "wTdFrRmIr") = "I"
        .Value(wRow, "TdFrRmDc") = "C"
        .Value(wRow, "TdRmCd") = wRsLsN!qRmCd
        ' Zubin 212
        .Value(wRow, "TdLotNo") = wRsLsN!qLotNo
        ' Zubin 212
        .Value(wRow, "TdRmSz") = wRsLsN!qRmSz
        .Value(wRow, "TdRmStkRt") = wRsLsN!qRmStkRt
        .Value(wRow, "TdRmQty") = wRsLsN!qRmQty
        .Value(wRow, "TdDustWt") = 0#                                       'adc("wLsDustWt")
        .Value(wRow, "TdToRmLoc") = wLsLoc
        .Value(wRow, "TdByLoc") = GrdTxndA.Value(wTxndARow, "TdFrBLoc")
        .Value(wRow, "TdByWrk") = GrdTxndA.Value(wTxndARow, "TdWrk")
        .Value(wRow, "TdPrtKey") = ctCurrPrtn '****** Sachin 3.02
        If Not (wLsTyp <> "" And wWorkLsLoc <> "" And wLsPerN <> 0) Then
          .Value(wRow, "TdRmWt") = wLsWtN
          .SaveRec (wRow)
        Else
          .Value(wRow, "TdRmWt") = IIF(wLsTyp = "LL", wLsWtN, _
                             (wRmWtN - wLsWtN)) * (wLsPerN / 100)
          .SaveRec (wRow)
          If (.Value(wRow, "TdRmWt") - Round(wLsWtN, 3)) <> 0 Then
            .AddItem
            wRow = wRow + 1
            .Value(wRow, "TdSrNo") = wRow
            .Value(wRow, "TdFrRmLoc") = "W"
            .Value(wRow, "wTdFrRmIr") = "I"       'IIF(.Value(wRow - 1, "TdRmWt") < (wBGrsWt - adc("wLsGrsWt")), "I", "R")
            .Value(wRow, "TdFrRmDc") = "C"        'IIF(.Value(wRow - 1, "TdRmWt") < (wBGrsWt - adc("wLsGrsWt")), "C", "D")
            .Value(wRow, "TdRmCd") = wRsLsN!qRmCd
            ' Zubin 212
            .Value(wRow, "TdLotNo") = wRsLsN!qLotNo
            ' Zubin 212
            .Value(wRow, "TdRmSz") = wRsLsN!qRmSz
            .Value(wRow, "TdRmStkRt") = wRsLsN!qRmStkRt
            .Value(wRow, "TdRmQty") = wRsLsN!qRmQty
            .Value(wRow, "TdRmWt") = wLsWtN - .Value(wRow - 1, "TdRmWt") 'Abs((wBGrsWt - adc("wLsGrsWt")) - .Value(wRow - 1, "TdRmWt"))
            .Value(wRow, "TdDustWt") = 0#
            .Value(wRow, "TdToRmLoc") = wWorkLsLoc
            .Value(wRow, "TdByLoc") = GrdTxndA.Value(wTxndARow, "TdFrBLoc")
            .Value(wRow, "TdByWrk") = GrdTxndA.Value(wTxndARow, "TdWrk")
            .Value(wRow, "TdPrtKey") = ctCurrPrtn '****** Sachin 3.02
            .SaveRec (wRow)
          End If
        End If
      End If
      'Manoj 2.10.0
      
    End With
  End If

  MWLib.EndProcess Me
  GrdTxndB.EndCopy
  Set wRsLs = Nothing
  '*** End of the Loss Entry

  '*** Calculation Of TdBGrsWt ***
  '****** Sachin 3.02 22-12-07 - Current Partition
  GrdTxndA.Value(wTxndARow, "TdBGrWt") = moCn.GetFldVal("Select BGrWt from Bag where " + _
                                         "BCoCd='" + ADC("TCoCd") + "' and " + _
                                         "BYy='" + GrdTxndA.Value(wTxndARow, "TdBYy") + "' and " + _
                                         "BChr='" + GrdTxndA.Value(wTxndARow, "TdBChr") + "' and " + _
                                         "BNo= " + CStr(GrdTxndA.Value(wTxndARow, "TdBNo")) + " and BPrtKey='" + ctCurrPrtn + "'")
  GrdTxndA.SaveRec (wTxndARow)
End Sub
Private Sub PrdPoints(ByVal mDefSubPrc As String)
  '*** Procedure to enter records into the the Production Points table (TxndPrd)

  If mf_DtLocked = True Then DispMsg "Cannot Enter Prd Points For Locked/Previous Partition Entry", etError: Exit Sub
  If mf_ValidLocSeq = False Then DispMsg "The User Does Not Have Access To One Of The Locs", etError: Exit Sub
    
  Dim wTxndARow As Single, wRow As Single, wPrdTyp As String, wBQty As Single
  Dim wRsDsg As MwfLib.MDORowSet, wDmCd As String, wPrdCtg As String, wPts As Single
  Dim wPtsSql As String, wPrcLoc As String
  Dim wRsRm As MwfLib.MDORowSet, wOrRmQty As Single, wTotRows As Integer
  Dim wCtr As Integer, wMulFact As Single, wBagRmQty As Single, wRmQty As Single, wStr As String
  Dim wWHQty As Single
  Dim wDmTcTyp As String
  Dim wParentSave As Boolean
  If UCase(ADC.MenuCd) = UCase("RDT") Then Exit Sub

  If GrdTP.Rows - 1 > 0 Then Exit Sub
  
  wTxndARow = IIF(GrdTxndA.Row, GrdTxndA.Row, 0)
  Dim wRsTp As MwfLib.MDORowSet
  Dim wTpBQty As Single
  
  If wTxndARow > 0 Then wParentSave = GrdTxndA.SaveRec(wTxndARow)
  If wParentSave = False Then Exit Sub
  '*** Check If Last Transaction Of Bag ***
  '****** Sachin 3.02 22-12-07 - Current Partition
  If Not moCn.RecSeek("Select BYy from Bag where " + _
                      "BCoCd='" + ADC("TCoCd") + "' and " + _
                      "BYy='" + GrdTxndA.Value(wTxndARow, "TdBYy") + "' and " + _
                      "BChr='" + GrdTxndA.Value(wTxndARow, "TdBChr") + "' and " + _
                      "BNo= " + CStr(GrdTxndA.Value(wTxndARow, "TdBNo")) + " and " + _
                      "BLstYy='" + ADC("TYy") + "' and " + _
                      "BLstKey=" + CStr(GrdTxndA.Value(wTxndARow, "TdKey")) + " and BPrtKey='" + ctCurrPrtn + "' ") Then
    DispMsg "Cannot Edit Since Not The Last Transaction Of The Bag", etError: Exit Sub
  End If
  'uma xxxxxxxx Getting R for P - LocPrcs
   wPrcLoc = moCn.GetFldVal("Select LocPrcs from Loc where LocCoCd = '" + ADC("TCoCd") + "' and LocTyp= 'P' and " + _
               "LocCd = '" + GrdTxndA.Value(GrdTxndA.Row, "TdByLoc") + "'")

   wPrdTyp = moCn.GetFldVal("Select LocPrdTyp from Loc where " + _
            "LocCoCd= '" + ADC("TCoCd") + "' and " + _
            "LocCd= '" + wPrcLoc + "'")

  '****** Sachin 3.02 22-12-07 - Current Partition
  'indigo.164 BOdDmTcTyp added
  Set wRsDsg = moCn.OpenRes("Select BQty, BOdDmCd,BOdDmTcTyp from Bag where " + _
                            "BCoCd='" + ADC("TCoCd") + "' and " + _
                            "BYy='" + GrdTxndA.Value(wTxndARow, "TdBYy") + "' and " + _
                            "BChr='" + GrdTxndA.Value(wTxndARow, "TdBChr") + "' and " + _
                            "BNo= " + CStr(GrdTxndA.Value(wTxndARow, "TdBNo")) + " and BPrtKey='" + ctCurrPrtn + "'")
  If wRsDsg.RecCount > 0 Then
     wRsDsg.MoveFirst: wBQty = wRsDsg!BQty: wDmCd = wRsDsg!BOdDmCd
     wDmTcTyp = wRsDsg!BOdDmTcTyp   'indigo.164 BOdDmTcTyp assigned
     '****** Sachin 3.02 22-12-07 - Current Partition
     '************Bhavna
     ' Getting Process loc instead of P Loc for Dpploc
     'indigo.164 DppTcTyp= 'DM'  changed as DppTcTyp= '" + wDmTcTyp + "'
     wPrdCtg = moCn.GetFldVal("Select DppPrdCtg from DsgPrd where " + _
                                     "DppTcTyp= '" + wDmTcTyp + "' and DppDmCd= '" + wDmCd + "' and DppLoc= '" + _
                                     wPrcLoc + "' And DppPrtKey='" + ctCurrPrtn + "'")
     'indigo.164 DmTcTyp= 'DM' changed as DmTcTyp= '" + wDmTcTyp + "'
     If wPrdCtg = "" Then
        wPrdCtg = moCn.GetFldVal("Select DmPrdCtg from DsgMst where DmTcTyp= '" + wDmTcTyp + "' and " + _
                              "DmCd= '" + wDmCd + "' and DmSz= '' and DmPrtKey='" + ctCurrPrtn + "'")
     End If
     '***********Bhavna
  End If
  
  If wBQty < 0 Then DispMsg "Current Bag Quantity< 0", etError: Exit Sub
  
  GrdTP.AllowAdd = True: GrdTP.AllowDelete = True
  MWLib.BeginProcess Me, "Entering The Production Points ..."
  With GrdTP
    Select Case UCase(Trim(wPrdTyp))
    Case Is = "GN"
    'xxxxxxx  GrdTxndA.Value(wTxndARow, "TdByLoc")
      wPts = moCn.GetFldVal("Select IsNull" + _
             "((Select PpPts from PrdPts where PpTyp= '" + wPrdTyp + "' and " + _
               "PpLoc= '" + wPrcLoc + "' and " + _
               "PpSubPrc= '" + mDefSubPrc + "' and PpPrdCtg= '" + wPrdCtg + "' and " + _
               "PpDmTcTyp='" + wDmTcTyp + "' and PpDmCd= '" + wDmCd + "' and PpSetTyp= '' and " + _
               "PpFrRmWt <= " + CStr(GrdTxndA.Value(wTxndARow, "TdBGrWt")) + " and " + _
               "PpToRmWt >= " + CStr(GrdTxndA.Value(wTxndARow, "TdBGrWt")) + " and " + _
               " PpPts<> 0 and PpPrtKey='" + GrdTxndA.Value(wTxndARow, "TdPrtKey") + "')" + _
              ", " + _
              "IsNull((Select PpPts from PrdPts where PpTyp= '" + wPrdTyp + "' and " + _
                      "PpLoc= '" + wPrcLoc + "' and " + _
                      "PpSubPrc= '" + mDefSubPrc + "' and PpPrdCtg= '" + wPrdCtg + "' and " + _
                      "PpDmTcTyp='" + wDmTcTyp + "' and PpDmCd= '' and PpSetTyp= '' and " + _
                      "PpFrRmWt <= " + CStr(GrdTxndA.Value(wTxndARow, "TdBGrWt")) + " and " + _
                      "PpToRmWt >= " + CStr(GrdTxndA.Value(wTxndARow, "TdBGrWt")) + " and " + _
                      "PpPts<> 0 and PpPrtKey='" + GrdTxndA.Value(wTxndARow, "TdPrtKey") + "'),0))")
                      
                      
             wTpBQty = moCn.GetFldVal("Select sum(TpBQty) from TxndPrd join Loc on LocCoCd = TpCoCd and LocCd = TpByLoc join Txnd on TdCoCd= TpCoCd and " + _
                  "TdTc= TpTc and TdYy= TpYy and TdChr= TpChr and TdNo= TpNo and TdSr= TpSr and TdPrtKey=TpPrtKey and TdSrNo= 0 " + _
                  " And Loc.LocPrcs = '" + wPrcLoc + "'" + _
                  " where TdCoCd= '" + ADC("TCoCd") + "' and TpSubPrc='" + mDefSubPrc + "' and " + _
                  "TdBYy= '" + GrdTxndA.Value(wTxndARow, "TdBYy") + "' and " + _
                  "TdBChr= '" + GrdTxndA.Value(wTxndARow, "TdBChr") + "' and " + _
                  "TdBNo= " + CStr(GrdTxndA.Value(wTxndARow, "TdBNo")) + " and TdPrtKey='" + ctCurrPrtn + "'")
                  
                  
         Dim wResPrdPtsOnBQty As Boolean
        wResPrdPtsOnBQty = moCn.RecSeek("Select 'x'  From Head where HCd='ZSELF' and HCoCd='" + ADC("TCoCd") + "' and hPrdPtsOnBQtyYN='Y'")
       ' wBQty = IIF(wResPrdPtsOnBQty, wBQty - wTpBQty, wBQty)
       '**** BQty to PDQty
       ms_BTyp = moCn.GetFldVal("Select PValue from Param where PTyp ='BCHR' and PMCd='" + GrdTxndA.Value(wTxndARow, "TdBChr") + "' and PSCd=''")
       If ms_BTyp = ctBTyp Then
          wBQty = GrdTxndA.Value(wTxndARow, "TdPDQty") - wTpBQty
       Else
          wBQty = wBQty - wTpBQty
       End If
            
      If (wResPrdPtsOnBQty And wBQty > 0) Or Not wResPrdPtsOnBQty Then
      .AddItem
      wRow = .Rows - 1
      .Value(wRow, "TpSrNo") = wRow
      .Value(wRow, "TpByWrk") = GrdTxndA.Value(wTxndARow, "TdByWrk")
      .Value(wRow, "TpSubPrc") = mDefSubPrc
      .Value(wRow, "TpBQty") = wBQty
      .Value(wRow, "TpRmQty") = 0
      .Value(wRow, "TpPts") = IIF(wBQty > 0, wBQty, 1) * IIF(wPts > 0, wPts, 1)
      .Value(wRow, "TpRmCtg") = ""
      .Value(wRow, "TpRmSCtg") = ""
      .Value(wRow, "TpSetTyp") = ""
      .Value(wRow, "TpPrdCtg") = wPrdCtg
      .Value(wRow, "TpPrtKey") = ctCurrPrtn '****** Sachin 3.02
            
      .SaveRec (wRow)
      End If
 Case Is = "WS", "GS", "WP", "GP"
 Dim wStrForPts As String
' ###################################  Manoj #### Ver: 2.0.6  ###################################
' Date: 05/12/2003
' Change: In Case Statement Wp & Gp Are Added
' Effects: to Give Production Points to Picking also
' ###########################################  Manoj  ###########################################
      wStr = IIF(UCase(Trim(wPrdTyp)) = "WS" Or UCase(Trim(wPrdTyp)) = "WP", " sum(OrWsQty) as qRmQty ", " sum(OrHsQty) as qRmQty ")
      wStrForPts = IIF(UCase(Trim(wPrdTyp)) = "WS" Or UCase(Trim(wPrdTyp)) = "WP", " (OrWsQty) as qRmQty ", " (OrHsQty) as qRmQty ")
      
      ' Zubin 212 (27/02/06)
      Dim wCndStrForPts As String, wCndStr As String
      wCndStr = IIF(UCase(Trim(wPrdTyp)) = "WS", " having sum(qRmQty)> 0 ", IIF(UCase(Trim(wPrdTyp)) = "GS", " having sum(qRmQty)> 0 ", ""))
      wCndStrForPts = IIF(UCase(Trim(wPrdTyp)) = "WS", " (OrWsQty)> 0 ", IIF(UCase(Trim(wPrdTyp)) = "GS", " OrHsQty> 0 ", " 1 = 1 "))

      ' Zubin 212 (27/02/06)
      '****** Sachin 3.02 22-12-07 - Current Partition
      Set wRsRm = moCn.OpenRes("Select RmCtg, RmSCtg, sum((case (case when TdFrRmLoc='W' " + _
                  "then TdFrRmDc when TdToRmLoc='W' then TdToRmDc end) when 'C' then -1 " + _
                  "when 'D' then +1 else 0 end)* TdRmQty) as qRmQty, max(BOdTc) as BOdTc, max(BOdYy) as BOdYy, max(BOdChr) as BOdChr, " + _
                  "max(BOdNo) as BOdNo, max(BOdSr) as BOdSr From Txnd join Bag on BIdNo=TdBIdNo And BCoCd= TdCoCd and " + _
                  "BYy= TdBYy and BChr= TdBChr and BNo= TdBNo and BPrtKey=TdPrtKey join RmMst on RmCd= TdRmCd and RmPrtKey=TdPrtKey " + _
                  "where TdCoCd='" + ADC("TCoCd") + "' and " + _
                  "TdBYy= '" + GrdTxndA.Value(wTxndARow, "TdBYy") + "' and " + _
                  "TdBChr= '" + GrdTxndA.Value(wTxndARow, "TdBChr") + "' and " + _
                  "TdBNo= " + CStr(GrdTxndA.Value(wTxndARow, "TdBNo")) + " and TdPrtKey='" + ctCurrPrtn + "' " + _
                  "group By RmCtg, RmSCtg order By RmCtg, RmSCtg ")
      Do While Not (wRsRm.EOF Or wRsRm.BOF)
        'wBagRmQty = wRsRm!qRmQty
        If wRsRm!qRmQty > 0 Then
          ' Zubin 212 (Before 27/02/06)
          'wOrRmQty = moCn.GetFldVal("Select sum(OrQty) as qOrQty from OrdRm where " + _
                    "OrCoCd= '" + ADC("TCoCd") + "' and OrTc= '" + wRsRm!BOdTc + "' and " + _
                    "OrYy= '" + wRsRm!BOdYy + "' and OrChr= '" + wRsRm!BOdChr + "' and " + _
                    "OrNo= " + CStr(wRsRm!BOdNo) + " and OrSr= " + CStr(wRsRm!BOdSr) + " and OrRmCtg= '" + wRsRm!RmCtg + "' and " + _
                    "OrRmSCtg= '" + wRsRm!RmSCtg + "' ")
          ' Zubin 212 (27/02/06)
          '****** Sachin 3.02 22-12-07 - Current Partition
          wOrRmQty = moCn.GetFldVal("Select sum(" + IIF(UCase(Trim(wPrdTyp)) = "WS", "OrWsQty", IIF(UCase(Trim(wPrdTyp)) = "GS", "OrHsQty", "OrQty")) + ") as qOrQty from OrdRm where " + _
                    "OrCoCd= '" + ADC("TCoCd") + "' and OrTc= '" + wRsRm!BOdTc + "' and " + _
                    "OrYy= '" + wRsRm!BOdYy + "' and OrChr= '" + wRsRm!BOdChr + "' and " + _
                    "OrNo= " + CStr(wRsRm!BOdNo) + " and OrSr= " + CStr(wRsRm!BOdSr) + " and OrRmCtg= '" + wRsRm!RmCtg + "' and " + _
                    "OrRmSCtg= '" + wRsRm!RmSCtg + "' and OrPrtKey='" + ctCurrPrtn + "' ")
' ###################################  Manoj #### Ver: 2.0.6  ###################################
' Date: 05/12/2003
' Change: To Check Already given Prdpts of Same Type i.e. Setting or Picking
          '     Filter for PrdTyp is Added
          '****** Sachin 3.02 22-12-07 - Current Partition
             '****** Sachin 3.02 22-12-07 - Current Partition
          Dim wPrcLocQuery As String
          wPrcLocQuery = "(Select LocPrcs from Loc where LocCoCd = '" + ADC("TCoCd") + "' and LocTyp= 'P' and " + _
               "LocCd = Txnd.TdByLoc)"
               
               
          wWHQty = moCn.GetFldVal("Select sum(TpRmQty) from TxndPrd join Txnd on TdCoCd= TpCoCd and " + _
                  "TdTc= TpTc and TdYy= TpYy and TdChr= TpChr and TdNo= TpNo and TdSr= TpSr and TdPrtKey=TpPrtKey and TdSrNo= 0 " + _
                  " And (Select LocPrdTyp From Loc Where LocCoCd = TdCoCd And LocCd = " + wPrcLocQuery + ") in (" + _
                  IIF(wPrdTyp = "GS" Or wPrdTyp = "WS", "'GS', 'WS'", "'GP', 'WP'") + _
                  ") where TdCoCd= '" + ADC("TCoCd") + "' and " + _
                  "TdBYy= '" + GrdTxndA.Value(wTxndARow, "TdBYy") + "' and " + _
                  "TdBChr= '" + GrdTxndA.Value(wTxndARow, "TdBChr") + "' and " + _
                  "TdBNo= " + CStr(GrdTxndA.Value(wTxndARow, "TdBNo")) + " and TdPrtKey='" + ctCurrPrtn + "' and " + _
                  "TpRmCtg= '" + wRsRm!RmCtg + "' and TpRmSCtg= '" + wRsRm!RmSCtg + "' ")
                  
        Dim wResPrdPtsOnRmQty As Boolean
        wResPrdPtsOnRmQty = moCn.RecSeek("Select 'x'  From Head where HCd='ZSELF' and HCoCd='" + ADC("TCoCd") + "' and HOrdPrdPts='Y'")
       
          wBagRmQty = wRsRm!qRmQty - wWHQty



          If wPrdTyp = UCase$("GS") Or wPrdTyp = UCase$("WS") Then





            wPtsSql = " IsNull((Select PpPts from PrdPts where PpTyp= '" + wPrdTyp + "' and " + _
                      "PpLoc= '" + wPrcLoc + "' and " + _
                      "PpSubPrc= '" + mDefSubPrc + "' and PpPrdCtg = '' and PpDmTcTyp='" + wDmTcTyp + "' and PpDmCd= '" + wDmCd + "' and " + _
                      "PpSetTyp = OrSetSCd And PpRmCtg = '' And PpRmSCtg = '' and PpFrRmWt<=convert(decimal(9,3),case when OrPrdQty>0 then OrPrdWt/OrPrdQty else 0 end) and PpToRmWt>=convert(decimal(9,3),case when OrPrdQty>0 then OrPrdWt/OrPrdQty else 0 end) " + _
                      "and PpPts<> 0 and PpPrtKey='" + GrdTxndA.Value(wTxndARow, "TdPrtKey") + "'), " + _
                      "IsNull((Select PpPts from PrdPts where PpTyp= '" + wPrdTyp + "' and " + _
                      "PpLoc= '" + wPrcLoc + "' and " + _
                      "PpSubPrc= '" + mDefSubPrc + "' and PpPrdCtg = '' and PpDmCd= '' and " + _
                      "PpSetTyp = OrSetSCd And PpRmCtg = '' And PpRmSCtg = '' and PpFrRmWt<=convert(decimal(9,3),case when OrPrdQty>0 then OrPrdWt/OrPrdQty else 0 end) and PpToRmWt>=convert(decimal(9,3),case when OrPrdQty>0 then OrPrdWt/OrPrdQty else 0 end) " + _
                      "and PpPts<> 0 and PpPrtKey='" + GrdTxndA.Value(wTxndARow, "TdPrtKey") + "'),0)) as qPts "
          ElseIf wPrdTyp = UCase$("GP") Or wPrdTyp = UCase$("WP") Then
            wPtsSql = " IsNull((Select PpPts from PrdPts where PpTyp= '" + wPrdTyp + "' and " + _
                      "PpLoc= '" + wPrcLoc + "' and " + _
                      "PpSubPrc= '" + mDefSubPrc + "' and PpPrdCtg = '' and PpDmTcTyp='" + wDmTcTyp + "' and PpDmCd= '' and " + _
                      "PpSetTyp = '' And PpRmCtg = OrRmCtg And PpRmSCtg = OrRmSCtg " + _
                      "and PpPts<> 0 and PpPrtKey='" + GrdTxndA.Value(wTxndARow, "TdPrtKey") + "'), " + _
                      "   IsNull((Select PpPts from PrdPts where PpTyp= '" + wPrdTyp + "' and " + _
                      "PpLoc= '" + wPrcLoc + "' and " + _
                      "PpSubPrc= '" + mDefSubPrc + "' and PpPrdCtg = '' and PpDmTcTyp='" + wDmTcTyp + "' and PpDmCd= '' and " + _
                      "PpSetTyp = '' And PpRmCtg = OrRmCtg And PpRmSCtg = '' " + _
                      "and PpPts<> 0 and PpPrtKey='" + GrdTxndA.Value(wTxndARow, "TdPrtKey") + "'), 0)) as qPts "
          End If





          Set wRsTp = moCn.OpenRes("Select " + _
                    IIF(wPrdTyp = "GS" Or wPrdTyp = "WS", _
                      "max(OrRmCtg) as RmCtg, max(OrRmSCtg) as RmSCtg, OrSetSCd, ", _
                      "OrRmCtg as RmCtg, OrRmSCtg as RmSCtg, '' As OrSetSCd, ") + _
                      " sum(qRmQty) as qRmQty , (qPts) as qPts From " + _
                      " (Select " + _
                    IIF(wPrdTyp = "GS" Or wPrdTyp = "WS", _
                      "(OrRmCtg) as OrRmCtg, (OrRmSCtg) as OrRmSCtg, OrSetSCd, ", _
                      "OrRmCtg as oRRmCtg, OrRmSCtg as oRRmSCtg, '' As OrSetSCd, ") + _
                      wStrForPts + "," + wPtsSql + _
                      " From  OrdRm join Bag on BOdIdNo=OrOdIdNo And " + _
                      "BCoCd= OrCoCd and BOdTc= OrTc and BOdYy= OrYy and " + _
                      "BOdChr= OrChr and BOdNo= OrNo and BOdSr= OrSr and BPrtKey=OrPrtKey " + _
                      "where OrRmCtg= '" + wRsRm!RmCtg + "' and OrRmSCtg= '" + wRsRm!RmSCtg + "' and OrPrtKey='" + ctCurrPrtn + "' And " + _
                      "OrRmCd in (Select TdRmCd from Txnd where TdCoCd= BCoCd and " + _
                      "TdBYy= BYy and TdBChr= BChr and TdBNo= BNo and TdPrtKey=BPrtKey and TdRmCd<> '') and " + _
                      "BCoCd= '" + ADC("TCoCd") + "' and " + _
                      "BYy= '" + GrdTxndA.Value(wTxndARow, "TdBYy") + "' and " + _
                      "BChr= '" + GrdTxndA.Value(wTxndARow, "TdBChr") + "' and " + _
                      "BNo= " + CStr(GrdTxndA.Value(wTxndARow, "TdBNo")) + " and " + _
                      "BPrtKey= '" + ctCurrPrtn + "' and " + wCndStrForPts + " ) ForPts " + _
                     IIF(wPrdTyp = "GS" Or wPrdTyp = "WS", _
                      "Group By OrSetSCd, qPts" + wCndStr + " Order By OrSetSCd,qPts ", _
                      "Group By OrRmCtg, OrRmSCtg,QpTS Order By OrRmCtg, OrRmSCtg,QPTS "))



























          wTotRows = wRsTp.RecCount
          wCtr = 0
          Do While Not (wRsTp.EOF Or wRsTp.BOF)
            wMulFact = MWLib.Div(wBagRmQty, wOrRmQty)
            wCtr = wCtr + 1
            If (wCtr <> wTotRows And Round(wRsTp!qRmQty * wMulFact) > 0) Or _
               (wCtr = wTotRows And wBagRmQty > 0) Then
              .AddItem
              wRow = .Rows - 1
              .Value(wRow, "TpSrNo") = wRow
              .Value(wRow, "TpByWrk") = GrdTxndA.Value(wTxndARow, "TdByWrk")
              .Value(wRow, "TpSubPrc") = mDefSubPrc
              .Value(wRow, "TpBQty") = wBQty
              .Value(wRow, "TpRmCtg") = wRsTp!RmCtg
              .Value(wRow, "TpRmSCtg") = wRsTp!RmSCtg
              .Value(wRow, "TpSetTyp") = wRsTp!OrSetSCd
              .Value(wRow, "TpPrdCtg") = ""

              If wCtr < wTotRows Then
                '*** If Not The Last Record ***
                .Value(wRow, "TpRmQty") = Round(wRsTp!qRmQty * wMulFact)
                .Value(wRow, "TpPts") = Round(wRsTp!qRmQty * wMulFact) * IIF(wRsTp!qPts > 0, wRsTp!qPts, 1)
                wRmQty = Round(wRsTp!qRmQty * wMulFact)
              Else
                '*** If Last Record ***
                .Value(wRow, "TpRmQty") = Round(wBagRmQty)
                .Value(wRow, "TpPts") = Round(wBagRmQty) * IIF(wRsTp!qPts > 0, wRsTp!qPts, 1)
                wRmQty = Round(wBagRmQty)
              End If
              .Value(wRow, "TpPrtKey") = ctCurrPrtn '****** Sachin 3.02
              .SaveRec (wRow)

              wOrRmQty = wOrRmQty - wRsTp!qRmQty
              wBagRmQty = wBagRmQty - wRmQty
            End If
            wRsTp.MoveNext
          Loop
        End If
        wRsRm.MoveNext
      Loop

'    Case Is = "GS", "GP", "WS"
'      wStr = IIf(UCase(Trim(wPrdTyp)) = "WS", " sum(OrWsQty) ", " sum(OrHsQty) ")
'
'      wPtsSql = " IsNull((Select PpPts from PrdPts where PpTyp= '" + wPrdTyp + "' and " + _
'                "PpLoc= '" + GrdTxndA.Value(wTxndARow, "TdByLoc") + "' and " + _
'                "PpSubPrc= '" + mDefSubPrc + "' and PpSetTyp= OrSetSCd and " + _
'                "PpPrdCtg= '' and PpDmCd= '' and PpPts<> 0), 0) as qPts "
'      Set wRsTp = moCn.OpenRes("Select OrRmCtg as RmCtg, OrRmSCtg as RmSCtg, OrSetSCd, " + _
'                  wStr + " as qRmQty, " + wPtsSql + " From OrdRm join Bag on " + _
'                  "BOdTc= OrTc and BOdYy= OrYy and BOdChr= OrChr and BOdNo= OrNo and " + _
'                  "BOdSr= OrSr " + _
'                  "where OrRmCd in (Select TdRmCd from Txnd where TdBYy= BYy and " + _
'                  "TdBChr= BChr and TdBNo= BNo and TdRmCd<> '') and " + _
'                  "BYy= '" + GrdTxndA.Value(wTxndARow, "TdBYy") + "' and " + _
'                  "BChr= '" + GrdTxndA.Value(wTxndARow, "TdBChr") + "' and " + _
'                  "BNo= " + CStr(GrdTxndA.Value(wTxndARow, "TdBNo")) + " " + _
'                  "group By OrRmCtg, OrRmSCtg, OrSetSCd ")
'      Do While Not (wRsTp.EOF Or wRsTp.BOF)
'        If wRsTp!qRmQty > 0 Then
'          .AddItem
'          wRow = .Rows - 1
'          .Value(wRow, "TpSrNo") = wRow
'          .Value(wRow, "TpByWrk") = GrdTxndA.Value(wTxndARow, "TdByWrk")
'          .Value(wRow, "TpSubPrc") = mDefSubPrc
'          .Value(wRow, "TpBQty") = wBQty
'          .Value(wRow, "TpRmQty") = Round(wRsTp!qRmQty * wBQty)
'          .Value(wRow, "TpPts") = Round(wRsTp!qRmQty * wBQty) * IIf(wRsTp!qPts > 0, wRsTp!qPts, 1)
'          .Value(wRow, "TpRmCtg") = wRsTp!RmCtg
'          .Value(wRow, "TpRmSCtg") = wRsTp!RmSCtg
'          .Value(wRow, "TpSetTyp") = wRsTp!OrSetSCd
'          .Value(wRow, "TpPrdCtg") = ""
'          .SaveRec (wRow)
'        End If
'        wRsTp.MoveNext
'      Loop
    End Select
  End With
  
  MWLib.EndProcess Me
    GrdTP.EndCopy
  
  Set wRsTp = Nothing
  Set wRsDsg = Nothing
  '*** End of the Production Points Entry
End Sub
Private Sub PrdPoints1(ByVal mDefSubPrc As String)
'  '*** Before 11/04/02
'  '*** Procedure to enter records into the the Production Points table (TxndPrd)
'
'
'  Dim wTxndARow As Single, wRow As Single, wPrdTyp As String, wBQty As Single
'  Dim wParentSave As Boolean
'  If GrdTP.Rows - 1 > 0 Then Exit Sub
'  wTxndARow = IIf(GrdTxndA.Row, GrdTxndA.Row, 1)
'  Dim wRsTp As MwfLib.MDORowSet
'
'  If wTxndARow > 0 Then wParentSave = GrdTxndA.SaveRec(wTxndARow)
'  If wParentSave = False Then Exit Sub
'  '*** Check If Last Transaction Of Bag ***
'  If Not moCn.RecSeek("Select BYy from Bag where BYy='" + GrdTxndA.Value(wTxndARow, "TdBYy") + "' and " + _
'                  "BChr='" + GrdTxndA.Value(wTxndARow, "TdBChr") + "' and " + _
'                  "BNo= " + CStr(GrdTxndA.Value(wTxndARow, "TdBNo")) + " and " + _
'                  "BLstYy='" + ADC("TYy") + "' and " + _
'                  "BLstKey=" + CStr(GrdTxndA.Value(wTxndARow, "TdKey"))) Then
'    DispMsg "Cannot Edit Since Not The Last Transaction Of The Bag", etError: Exit Sub
'  End If
'
'  wPrdTyp = moCn.GetFldVal("Select LocPrdTyp from Loc where " + _
'            "LocCd= '" + GrdTxndA.Value(wTxndARow, "TdByLoc") + "'")
'
'  wBQty = moCn.GetFldVal("Select BQty from Bag where " + _
'                         "BYy='" + GrdTxndA.Value(wTxndARow, "TdBYy") + "' and " + _
'                         "BChr='" + GrdTxndA.Value(wTxndARow, "TdBChr") + "' and " + _
'                         "BNo= " + CStr(GrdTxndA.Value(wTxndARow, "TdBNo")))
'
'  If wBQty <= 0 Then DispMsg "Current Bag Quantity= 0", etError: Exit Sub
'
'  GrdTP.AllowAdd = True: GrdTP.AllowDelete = True
'  GrdTP.StartCopy
'  MWLib.BeginProcess Me, "Entering The Production Points ..."
'  With GrdTP
'    Select Case UCase(Trim(wPrdTyp))
'    Case Is = "GN", "FP", "WP"
'      'If wBQty <= 0 Then DispMsg "Current Bag Quantity= 0", etError: Exit Sub
'      .AddItem
'      wRow = .Rows - 1
'      .Value(wRow, "TpSrNo") = wRow
'      .Value(wRow, "TpByWrk") = GrdTxndA.Value(wTxndARow, "TdByWrk")
'      .Value(wRow, "TpSubPrc") = mDefSubPrc
'      .Value(wRow, "TpBQty") = wBQty
'      .Value(wRow, "TpRmQty") = 0
'      .Value(wRow, "TpPts") = .Value(wRow, "TpBQty")
'      .Value(wRow, "TpRmCtg") = ""
'      .Value(wRow, "TpRmSCtg") = ""
'      .SaveRec (wRow)
'    Case Is = "WS"
'      Set wRsTp = moCn.OpenRes("Select RmCtg, RmSCtg, sum((case (case when TdFrRmLoc='W' " + _
'                  "then TdFrRmDc when TdToRmLoc='W' then TdToRmDc end) when 'C' then -1 " + _
'                  "when 'D' then +1 else 0 end)* TdRmQty) as qRmQty " + _
'                  "From Txnd join Bag on " + _
'                  "BYy= TdBYy and BChr= TdBChr and BNo= TdBNo join RmMst on RmCd= TdRmCd " + _
'                  "where TdBYy= '" + GrdTxndA.Value(wTxndARow, "TdBYy") + "' and " + _
'                  "TdBChr= '" + GrdTxndA.Value(wTxndARow, "TdBChr") + "' and " + _
'                  "TdBNo= " + CStr(GrdTxndA.Value(wTxndARow, "TdBNo")) + " " + _
'                  "group By RmCtg, RmSCtg ")
'      Do While Not (wRsTp.EOF Or wRsTp.BOF)
'        If wRsTp!qRmQty > 0 Then
'          .AddItem
'          wRow = .Rows - 1
'          .Value(wRow, "TpSrNo") = wRow
'          .Value(wRow, "TpByWrk") = GrdTxndA.Value(wTxndARow, "TdByWrk")
'          .Value(wRow, "TpSubPrc") = mDefSubPrc
'          .Value(wRow, "TpBQty") = wBQty
'          .Value(wRow, "TpRmQty") = wRsTp!qRmQty
'          .Value(wRow, "TpPts") = wRsTp!qRmQty
'          .Value(wRow, "TpRmCtg") = wRsTp!RmCtg
'          .Value(wRow, "TpRmSCtg") = wRsTp!RmSCtg
'          .SaveRec (wRow)
'        End If
'        wRsTp.MoveNext
'      Loop
'    Case Is = "GS", "GP"
'      Set wRsTp = moCn.OpenRes("Select OrRmCtg as RmCtg, OrRmSCtg as RmSCtg, " + _
'                  "sum(OrHsQty) as qRmQty From OrdRm join Bag on " + _
'                  "BOdTc= OrTc and BOdYy= OrYy and BOdChr= OrChr and BOdNo= OrNo and " + _
'                  "BOdSr= OrSr " + _
'                  "where OrRmCd in (Select TdRmCd from Txnd where TdBYy= BYy and " + _
'                  "TdBChr= BChr and TdBNo= BNo and TdRmCd<> '') and " + _
'                  "BYy= '" + GrdTxndA.Value(wTxndARow, "TdBYy") + "' and " + _
'                  "BChr= '" + GrdTxndA.Value(wTxndARow, "TdBChr") + "' and " + _
'                  "BNo= " + CStr(GrdTxndA.Value(wTxndARow, "TdBNo")) + " " + _
'                  "group By OrRmCtg, OrRmSCtg ")
'      Do While Not (wRsTp.EOF Or wRsTp.BOF)
'        If wRsTp!qRmQty > 0 Then
'          .AddItem
'          wRow = .Rows - 1
'          .Value(wRow, "TpSrNo") = wRow
'          .Value(wRow, "TpByWrk") = GrdTxndA.Value(wTxndARow, "TdByWrk")
'          .Value(wRow, "TpSubPrc") = mDefSubPrc
'          .Value(wRow, "TpBQty") = wBQty
'          .Value(wRow, "TpRmQty") = wRsTp!qRmQty * wBQty
'          .Value(wRow, "TpPts") = wRsTp!qRmQty * wBQty
'          .Value(wRow, "TpRmCtg") = wRsTp!RmCtg
'          .Value(wRow, "TpRmSCtg") = wRsTp!RmSCtg
'          .SaveRec (wRow)
'        End If
'        wRsTp.MoveNext
'      Loop
'    End Select
'  End With
'
'  MWLib.EndProcess Me
'  GrdTxndB.EndCopy
'  Set wRsTp = Nothing
'  '*** End of the Production Points Entry
End Sub
Private Function CopyBag() As Boolean

  '*** Procedure to Copy Bags from one Loc to another in case of 'BV' entry
  Dim wRsCB As MwfLib.MDORowSet, wTotRows As Single, wRow As Single, wCnd As String
  Dim wHldCnt As Integer
  Dim wRfIdStr As String, wRfIdNoArr() As String, wRfIdNo As String, i As Integer, wTotRfid As Integer
  Dim wMinSr As Integer
  
' ######################  Manoj # Ver: 2.0.11 Patch 01 # Date: 17/08/2005  ######################
  Dim wAckYn As Boolean
' ###########################################  Manoj  ###########################################
  Dim ws_TxndJoin As String
  Dim wInvalRfIdStr As String
   '********Bhavna added copy thru bag nos..
   If ADC("WCPYTHRUBAGNO") <> "" Then
        If ADC("wRfIdStr") = "" Then MsgBox "Please provide a valid string of " + IIF(ADC("wCPYTHRUBAGNO") = "Y", " Bag Nos.", "Bag Alias Nos") + " to Copy from": Exit Function
        If ADC("wCPYTHRUBAGNO") = "N" Then
               wInvalRfIdStr = ""
               wRfIdStr = ADC("WRFIDSTR")
               'vk.4 calling function to replace enter with ,
               wRfIdStr = PrepareRfIdString(wRfIdStr)
               
               wRfIdStr = ParseBagNoBarCd(wRfIdStr)

               wRfIdNoArr = Split(wRfIdStr, ",")
            'vk.4 Tag Input changed as Bag Alias No.
            If Left(wRfIdStr, 1) = "," Or Right(wRfIdStr, 1) = "," Then
                DispMsg "Invalid Bag Alias No.", etError: CopyBag = False: Exit Function
            End If
            wRfIdStr = ""
            For i = 0 To UBound(wRfIdNoArr)
              wRfIdNo = wRfIdNoArr(i)
              'vk.4 Tag changed as Bag Alias No.
              If Not IsNumeric(wRfIdNo) Then DispMsg "Invalid Bag Alias No.", etError: Exit Function
              If Not moCn.RecSeek("Select 'x' from RfIdMst Where RimNo=" + wRfIdNo + " ") Then
                wInvalRfIdStr = wInvalRfIdStr + IIF(wInvalRfIdStr = "", "", ",") + wRfIdNo
              End If
              wRfIdStr = wRfIdStr + IIF(wRfIdStr = "", "", ",") + wRfIdNo
            Next
            wTotRfid = UBound(wRfIdNoArr) + 1
            'vk.4 Tag changed as Bag Alias No.
            If wInvalRfIdStr <> "" Then DispMsg "Bag Alias No. Not Defined:" + wInvalRfIdStr, etError: Exit Function
                      wMinSr = GrdTxndA.MaxVal("TdSr") + 1
        Else
               wInvalRfIdStr = ""
               wRfIdStr = ADC("WRFIDSTR")
               'vk.4 calling function to replace enter with ,
               wRfIdStr = PrepareRfIdString(wRfIdStr)
               wRfIdStr = ParseBagNoBarCd(wRfIdStr)
               
               wRfIdNoArr = Split(wRfIdStr, ",")
        
            If Left(wRfIdStr, 1) = "," Or Right(wRfIdStr, 1) = "," Then
                DispMsg "Invalid Bag No Input", etError: CopyBag = False: Exit Function
            End If
            wRfIdStr = ""
            
            For i = 0 To UBound(wRfIdNoArr)
              wRfIdNo = wRfIdNoArr(i)
              If Not moCn.RecSeek("Select 'x' from Bag Where Byy+'/'+BChr+'/'+ltrim(RTrim(str(BNo))) = '" + wRfIdNo + "'") Then
                wInvalRfIdStr = wInvalRfIdStr + IIF(wInvalRfIdStr = "", "", ",") + wRfIdNo
              End If
              wRfIdStr = wRfIdStr + IIF(wRfIdStr = "", "", ",") + wRfIdNo
              
              '*********Chk bag validity as in col chng
              wCnd = " and IsNull((Select TdAck from Txnd where TdCoCd= BCoCd and TdSrNo = 0 and " + _
               " TdYy= BLstYy and TdKey= BLstKey and TdPrtKey=BPrtKey), '')<> 'N' "
'              If (Not moCn.RecSeek("Select BNo From Bag Where Byy+'/'+BChr+'/'+ltrim(RTrim(str(BNo))) = '" + wRfIdNo + "' and " + _
'                                    "BCoCd='" + ADC("TCoCd") + "' and " + _
'                                    "BCls='N' and BPrtKey='" + ctCurrPrtn + "' " + wCnd)) Then
'                DispMsg wRfIdNo + " : Invalid Bag Number Or Bag Closed or Bag is Unacknowledged", etError: Exit Sub
'              End If
            '*********Chk bag validity as in col chng
              
            Next
            wTotRfid = UBound(wRfIdNoArr) + 1
            If wInvalRfIdStr <> "" Then DispMsg "Bag Not Defined:" + wInvalRfIdStr, etError: CopyBag = False: Exit Function
                      wMinSr = GrdTxndA.MaxVal("TdSr") + 1
        End If
  End If
  
  If ADC.MenuCd = "BV" Then
     'vk.4 LTrim(ADC("WRFIDSTR")) replaced with LTrim(wRfIdStr)
     If LTrim(wRfIdStr) <> "" Then
              If ADC("WTTOBLOC") = "" Then DispMsg "To Location Cannot be Blank.", etError: CopyBag = False: Exit Function
                 ADC("wCbToBLoc") = ADC("WTTOBLOC")
          If ADC("WCPYTHRUBAGNO") = "N" Then
          'vk.4 ADC("WRFIDSTR") replaced with wRfIdStr
            wCnd = IIF(LTrim(wRfIdStr) <> "", " Exists(Select 'x' From RfIdMst where RimNo in (" + wRfIdStr + ") and BCoCd=RimBCoCd and BYy=RimBYy and BChr=RimBChr and BNo=RimBNo)", "")
            
          Else
'            wCnd = IIF(LTrim(adc("WRFIDSTR")) <> "", " (BYy+'/'+BChr+'/'+LTrim(RTrim(Str(BNo))) in ('" + adc("WRFIDSTR") + "'))", "")     '*********Bhavna
            wCnd = " charindex(','+BYy + '/' + BChr + '/' + LTrim(RTrim(str(BNo)))+',' , '," + wRfIdStr + ",') <> 0 "      '*************Bhavna
          End If
    Else
      If ADC("TFrBLoc") = "" Then DispMsg "Enter Default From Bag Location", etError: Exit Function
      If ADC("TFrBLoc") = "PFG" Then DispMsg "Default From Bag Location Cannot Not Be 'PFG'", etError: CopyBag = False: Exit Function
      If ADC("wCbToBLoc") = "" Then DispMsg "Enter To Bag Location", etError: Exit Function
      If ADC("wCbPrdYn") = "" Then DispMsg "Enter Prdoduction Y/N", etError: Exit Function
      'If adc("wCbToBLoc") = "PFG" Then DispMsg "To Bag Location Cannot Be 'PFG'", etError: Exit Sub
        wCnd = ADC.RepCond
     End If
      
      '*** (Bef 2.13)
      'adc("wCbBYyFr").CmpStr = "BYy>=": adc("wCbBChrFr").CmpStr = "BChr>=": adc("wCbBNoFr").CmpStr = "BNo>="
      'adc("wCbBYyTo").CmpStr = "BYy<=": adc("wCbBChrTo").CmpStr = "BChr<=": adc("wCbBNoTo").CmpStr = "BNo<="
      'adc("wCbOdTcFr").CmpStr = "BOdTc>=": adc("wCbOdYyFr").CmpStr = "BOdYy>=": adc("wCbOdChrFr").CmpStr = "BOdChr>=": adc("wCbOdNoFr").CmpStr = "BOdNo>=": adc("wCbOdSrFr").CmpStr = "BOdSr>="
      'adc("wCbOdTcTo").CmpStr = "BOdTc<=": adc("wCbOdYyTo").CmpStr = "BOdYy<=": adc("wCbOdChrTo").CmpStr = "BOdChr<=": adc("wCbOdNoTo").CmpStr = "BOdNo<=": adc("wCbOdSrTo").CmpStr = "BOdSr<="
      'adc("wCbDmCdFr").CmpStr = "BOdDmCd>=": adc("wCbDmCdTo").CmpStr = "BOdDmCd<="
      '*** (Bef 2.13)
      
      wCnd = IIF(wCnd <> "", " and " + wCnd, "")
      
      ' ***** Manali 3.8.0 - Wrk and Prd Y/N Scope added
      
      ' ***** Manali 3.8.0 - Wrk and Prd Y/N Scope added
      If ADC("wCbPrdYn") = "Y" Then
        ws_TxndJoin = " Join Txnd On TdCoCd=BCoCd and TdBYy=BYy and TdBChr=BChr and TdBNo=BNo and TdBIdNo=BIdNo and TdPrtKey=BPrtKey " + _
                      "and TdYy= BLstYy and TdKey= BLstKey and TdSrNo=0 and TdPrdYn='Y' "
      End If
      ' ***** Manali 3.8.0 - Wrk and Prd Y/N Scope added
      
      '****** Sachin 3.02 [13-12-07] - Check Below Done after Error at Lily Jewellery
      '****** Sachin 3.02 24-12-07 - Current Partition
      '" + IIF(adc("wCbPrdYn") = "Y", " and TdPrdYn='Y' ", "") + "
      If Not moCn.RecSeek("Select 'x' From Bag Where " + _
                          " BCoCd='" + ADC("TCoCd") + "' and BPrtKey='" + ctCurrPrtn + "' and " + _
                          " IsNull((Select TdAck from Txnd Where TdCoCd= BCoCd and TdSrNo = 0 and TdPrtKey=BPrtKey and " + _
                          " TdYy= BLstYy and TdKey= BLstKey " + IIF(ADC("wCbPrdYn") = "Y", " and TdPrdYn='Y' ", "") + " ), '')<> 'N' " + wCnd) Then _
      DispMsg "One of the Source Bag(s) Is Closed or Unacknowledged, Cannot Copy...", etError: Exit Function
      
      Set wRsCB = moCn.OpenRes("Select * from Bag " + ws_TxndJoin + " where BCoCd= '" + ADC("TCoCd") + "' and BCls= 'N' and BPrtKey='" + ctCurrPrtn + "' and " + _
                                "BLoc= '" + ADC("TFrBLoc") + "' and  not exists (Select 'x' from Txnd where " + _
                               "TdCoCd= '" + ADC("TCoCd") + "' and TdYy= BLstYy and TdKey= BLstKey and TdPrtKey=BPrtKey " + _
                               "and (Select PValue from Param where PTyp= 'TC' and " + _
                               "PMCd= TdTc) in('REJ', 'MLT')) " + wCnd + _
                               " Order By BYy, BChr, BNo")
    
    
      wTotRows = wRsCB.RecCount
      
    ' ######################  Manoj # Ver: 2.0.11 Patch 01 # Date: 17/08/2005  ######################
      wAckYn = IIF(moCn.GetFldVal("Select HTdAck From Head Where HCd = '" + ctSelfCmCd + "' And HCoCd = '" + ADC("TCoCd") + "'") = "N", True, False)
    ' ###########################################  Manoj  ###########################################
      '578.13 getting hBmOrdHldSeoYN value from head
      Dim wBagHldStr As String, wBagNoHld As String, wBmOrdHldSeoYN As String
      wBmOrdHldSeoYN = moCn.GetFldVal("Select hBmOrdHldSeoYN from Head where hCoCd='" + ADC("TCoCd") + "' and HCd='" + ctSelfCmCd + "'")
      
      GrdTxndA.StartCopy
      MWLib.BeginProcess Me, "Copying Bag Details ..."
      With GrdTxndA
        Do While Not (wRsCB.EOF Or wRsCB.BOF)
          wBagNoHld = ""
        '************Geeta*****************Emr206****************
          '****** Sachin 3.02 24-12-07 - Current Partition
          If moCn.GetFldVal("select OdHld from Orddsg where OdCoCd='" + ADC("TCoCd") + "' and  OdTc='" + wRsCB!BOdTc + "'  and " + _
                            "OdYy='" + wRsCB!BOdYy + "' and OdChr='" + wRsCB!BOdChr + "' and OdNo=" + CStr(wRsCB!BOdNo) + " and OdSr=" + CStr(wRsCB!BOdSr) + " and OdPrtKey='" + ctCurrPrtn + "'") = "Y" Then
             wHldCnt = wHldCnt + 1
             '578.13 stores bag nos which is on hold
             wBagNoHld = wRsCB!BYy + "/" + wRsCB!BChr + "/" + CStr(wRsCB!BNo)
             wBagHldStr = wBagHldStr + IIF(wBagHldStr <> "", ",", "") + wBagNoHld
          End If
         '************Geeta*****************Emr206****************
          '578.13 if hBmOrdHldSeoYN=Y, order on hold and password not given it will not copy bag and move to next bag
          If wBmOrdHldSeoYN = "Y" And ms_PwdFlag = False And wBagNoHld <> "" Then GoTo NextRow
                       
    ' ######################  Manoj # Ver: 2.0.11 Patch 01 # Date: 17/08/2005  ######################
          If ((Not wAckYn) And InStr(1, "," + gs_ValidPrdLocSeq + ",", wRsCB!BLoc) > 0 And _
             InStr(1, "," + gs_ValidPrdLocSeq + ",", ADC("wCbToBLoc")) > 0) Or gs_ValidPrdLocSeq = "" Then
             wAckYn = True
          End If
    ' ###########################################  Manoj  ###########################################
          
          If moCn.RecSeek("Select 'x' From Bag Where " + _
                          " BCoCd='" + ADC("TCoCd") + "' and BPrtKey='" + ctCurrPrtn + "' and " + _
                          " BYy='" + wRsCB!BYy + "' and BChr='" + wRsCB!BChr + "' and BNo=" + CStr(wRsCB!BNo) + " and  IsNull((Select TdAck from Txnd Where TdCoCd= BCoCd and TdSrNo = 0 and TdPrtKey=BPrtKey and " + _
                          " TdYy= BLstYy and TdKey= BLstKey " + IIF(ADC("wCbPrdYn") = "Y", " and TdPrdYn='Y' ", "") + " ), '')= 'N' ") Then
            GoTo NextRow
          End If
          
          .AddItem
          
          wRow = .Rows - 1
          .Value(wRow, "TdSr") = .MaxVal("TdSr") + 1
          .Value(wRow, "TdSrNo") = 0
          .Value(wRow, "TdBYy") = wRsCB!BYy
          .Value(wRow, "TdBChr") = wRsCB!BChr
          .Value(wRow, "TdBNo") = wRsCB!BNo
          .Value(wRow, "TdBQty") = wRsCB!BQty
          .Value(wRow, "TdPtQty") = wRsCB!BPtQty
          .Value(wRow, "TdBGrWt") = wRsCB!BGrWt
          .Value(wRow, "TdFrBLoc") = wRsCB!BLoc
          .Value(wRow, "TdToBLoc") = ADC("wCbToBLoc")
          .Value(wRow, "TdWrk") = wRsCB!BWrk
    ' ######################  Manoj # Ver: 2.0.11 Patch 01 # Date: 17/08/2005  ######################
          .Value(wRow, "TdAck") = IIF(wAckYn = True, "Y", "N")
    '      .Value(wRow, "TdAck") = "Y"
    ' ###########################################  Manoj  ###########################################
          .Value(wRow, "TdKey") = moCn.AutoGen("Txnd", "TdKey", ADC("TCoCd") + ADC("TYy"))
          .Value(wRow, "TdPrevYy") = wRsCB!BLstYy
          .Value(wRow, "TdPrevKey") = wRsCB!BLstKey
          .Value(wRow, "TdPrtKey") = ctCurrPrtn   '****** Sachin 3.02.0
          .SaveRec (wRow)
NextRow:
          wRsCB.MoveNext
        Loop
      End With
    
      MWLib.EndProcess Me
      GrdTxndA.EndCopy
      '*** End of the copy process
      '************Geeta*****************Emr206****************
        '578.13 showing error message
        If wHldCnt > 0 Then
          If wBmOrdHldSeoYN = "Y" And ms_PwdFlag = False Then
            DispMsg wBagHldStr + " Bags Not Copied ", etWarning
          ElseIf wBmOrdHldSeoYN = "Y" And ms_PwdFlag = True Then
            DispMsg "No.of Bags on Hold = " + CStr(wHldCnt) + " Have Been Copied", etWarning
          Else
        DispMsg CStr(wHldCnt) + " Order Design(s) on Hold", etWarning
          End If
        End If
      '************Geeta*****************Emr206****************
  ' **** Manali 3.6.0 - 26/10/09 - Copy Multi Bags into Melting
   ElseIf ADC.MenuCd = "REJ" Then
    'vk.4 ADC("WRFIDSTR") replaced with wRfIdStr
    If LTrim(wRfIdStr) <> "" Then
           If ADC("wTBIr") = "" Then DispMsg "Issue/Receipt Cannot be blank.", etError: Exit Function
           If ADC("wTRjCd") = "" Then DispMsg "Rejection Code Cannot be blank.", etError: Exit Function
           If ADC("WCPYTHRUBAGNO") = "N" Then
            'vk.4 ADC("WRFIDSTR") replaced with wRfIdStr
            wCnd = IIF(LTrim(wRfIdStr) <> "", "  Exists(Select 'x' From RfIdMst where RimNo in (" + wRfIdStr + ") and BCoCd=RimBCoCd and BYy=RimBYy and BChr=RimBChr and BNo=RimBNo) and BLoc<>'PFG' and BCls='N' ", "")
           Else
'            wCnd = IIF(LTrim(adc("WRFIDSTR")) <> "", " (BYy+'/'+BChr+'/'+LTrim(RTrim(Str(BNo))) in ('" + adc("WRFIDSTR") + "')) and BLoc<>'PFG' and BCls='N' ", "")     '*********Bhavna
            wCnd = " charindex(','+BYy + '/' + BChr + '/' + LTrim(RTrim(str(BNo)))+',' , '," + wRfIdStr + ",') <> 0  and BLoc<>'PFG' and BCls='N' "      '*************Bhavna
           End If
    Else
     ' If adc("wCbPrdYn") = "" Then DispMsg "Enter Prdoduction Y/N", etError: Exit Sub
      If ADC("wTRjCd") = "" Then DispMsg "Enter Default From Rejection Code", etError: Exit Function
      If ADC("wTBIr") = "" Then DispMsg "Enter Issue Or Receipt (I/R)", etError: Exit Function
      wCnd = ADC.RepCond
    End If
      
      wCnd = IIF(wCnd <> "", " and " + wCnd, "")
      
      If ADC("wCbPrdYn") = "Y" Then
        ws_TxndJoin = " Join Txnd On TdCoCd=BCoCd and TdBYy=BYy and TdBChr=BChr and TdBNo=BNo and TdBIdNo=BIdNo and TdPrtKey=BPrtKey " + _
                      "and TdYy= BLstYy and TdKey= BLstKey and TdSrNo=0 and TdPrdYn='Y' "
      End If
      If Not moCn.RecSeek("Select 'x' From Bag Where " + _
                          " BCoCd='" + ADC("TCoCd") + "' and BPrtKey='" + ctCurrPrtn + "' and " + _
                          " IsNull((Select TdAck from Txnd Where TdCoCd= BCoCd and TdSrNo = 0 and TdPrtKey=BPrtKey and " + _
                          " TdYy= BLstYy and TdKey= BLstKey " + IIF(ADC("wCbPrdYn") = "Y", " and TdPrdYn='Y' ", "") + " ), '')<> 'N' " + wCnd) Then _
      DispMsg "One of the Source Bag(s) Is Closed or Unacknowledged, Cannot Copy...", etError: Exit Function
      '****** Sachin 3.02 24-12-07 - Current Partition
      'vk.4 ADC("WRFIDSTR") replaced with wRfIdStr
      Set wRsCB = moCn.OpenRes("Select * from Bag " + ws_TxndJoin + " where BCoCd= '" + ADC("TCoCd") + "' and BCls= 'N' and BPrtKey='" + ctCurrPrtn + "' and " + _
                               " not exists (Select 'x' from Txnd where " + _
                               "TdCoCd= '" + ADC("TCoCd") + "' and TdYy= BLstYy and TdKey= BLstKey and TdPrtKey=BPrtKey " + _
                               "and (Select PValue from Param where PTyp= 'TC' and " + _
                               "PMCd= TdTc) in('REJ', 'MLT')) " + wCnd + _
                                IIF(wRfIdStr <> "" And ADC("WCPYTHRUBAGNO") = "N", " and  Exists(Select 'x' From RfIdMst where RimNo in (" + wRfIdStr + ") and BCoCd=RimBCoCd and BYy=RimBYy and BChr=RimBChr and BNo=RimBNo)", " /* and BLoc= '" + ADC("wCbFrBLoc") + "'*/ ") + _
                                " Order By BYy, BChr, BNo")
    
    
      wTotRows = wRsCB.RecCount
      
      wAckYn = IIF(moCn.GetFldVal("Select HTdAck From Head Where HCd = '" + ctSelfCmCd + "' And HCoCd = '" + ADC("TCoCd") + "'") = "N", True, False)
      
      GrdTxndA.StartCopy
      MWLib.BeginProcess Me, "Copying Bag Details ..."
      With GrdTxndA
        Do While Not (wRsCB.EOF Or wRsCB.BOF)
          If moCn.GetFldVal("select OdHld from Orddsg where OdCoCd='" + ADC("TCoCd") + "' and  OdTc='" + wRsCB!BOdTc + "'  and " + _
                            "OdYy='" + wRsCB!BOdYy + "' and OdChr='" + wRsCB!BOdChr + "' and OdNo=" + CStr(wRsCB!BOdNo) + " and OdSr=" + CStr(wRsCB!BOdSr) + " and OdPrtKey='" + ctCurrPrtn + "'") = "Y" Then _
             wHldCnt = wHldCnt + 1
          If ((Not wAckYn) And InStr(1, "," + gs_ValidPrdLocSeq + ",", wRsCB!BLoc) > 0 And _
             InStr(1, "," + gs_ValidPrdLocSeq + ",", ADC("wCbToBLoc")) > 0) Or gs_ValidPrdLocSeq = "" Then
             wAckYn = True
          End If
          .AddItem
          wRow = .Rows - 1
          .Value(wRow, "TdSr") = .MaxVal("TdSr") + 1
          .Value(wRow, "TdSrNo") = 0
          .Value(wRow, "TdBYy") = wRsCB!BYy
          .Value(wRow, "TdBChr") = wRsCB!BChr
          .Value(wRow, "TdBNo") = wRsCB!BNo
          .Value(wRow, "TdBQty") = wRsCB!BQty
          .Value(wRow, "TdPtQty") = wRsCB!BPtQty
          .Value(wRow, "TdBGrWt") = wRsCB!BGrWt
          .Value(wRow, "TdFrBLoc") = wRsCB!BLoc
          .Value(wRow, "TdWrk") = wRsCB!BWrk
          .Value(wRow, "wTdBIr") = ADC("wTBIr")
          .Value(wRow, "TdByLoc") = wRsCB!BLoc 'ADC("TFrBLoc")
          .Value(wRow, "TdByWrk") = wRsCB!BWrk
          .Value(wRow, "TdRjQty") = wRsCB!BQty

          .Value(wRow, "TdRjCd") = ADC("wTRjCd")
          .Value(wRow, "TdKey") = moCn.AutoGen("Txnd", "TdKey", ADC("TCoCd") + ADC("TYy"))
          .Value(wRow, "TdPrevYy") = wRsCB!BLstYy
          .Value(wRow, "TdPrevKey") = wRsCB!BLstKey
          .Value(wRow, "TdPrtKey") = ctCurrPrtn   '****** Sachin 3.02.0
          .SaveRec (wRow)
          wRsCB.MoveNext
        Loop
      End With
    
      MWLib.EndProcess Me
      GrdTxndA.EndCopy
        If wHldCnt > 0 Then _
        DispMsg CStr(wHldCnt) + " Order Design(s) on Hold", etWarning
 
  ElseIf ADC.MenuCd = "MLT" Then
      If (ADC("wCbTdTc") <> "" Or ADC("wCbTdYy") <> "" Or ADC("wCbTdChr") <> "" Or ADC("wCbTdNo") <> 0 Or _
        ADC("wCbTdSrFr") <> 0 Or ADC("wCbTdSrTo") <> 0) And _
        Not moCn.RecSeek("Select 'x' From Txn Where TCoCd='" + ADC("TCoCd") + "' and TTc='" + ADC("wCbTdTc") + "' and " + _
                        "TYy='" + ADC("wCbTdYy") + "' and TChr='" + ADC("wCbTdChr") + "' and TNo=" + CStr(ADC("wCbTdNo")) + " and " + _
                        "(Select PValue from Param where PTyp= 'TC' and PMCd= TTc)= 'REJ'") Then
           DispMsg "Invalid Rejection Voucher", etError: Exit Function
      End If
      
      wCnd = ADC.RepCond
      'vk.4 ADC("WRFIDSTR") replaced with wRfIdStr
      wCnd = IIF(LTrim(wRfIdStr) <> "" And ADC("WCPYTHRUBAGNO") = "N", "  Exists(Select 'x' From RfIdMst where RimNo in (" + wRfIdStr + ") and BCoCd=RimBCoCd and BYy=RimBYy and BChr=RimBChr and BNo=RimBNo) and BLoc<>'PFG' and BCls='N' ", wCnd)
'      wCnd = IIF(LTrim(adc("WRFIDSTR")) <> "" And adc("WCPYTHRUBAGNO") = "Y", " (BYy+'/'+BChr+'/'+LTrim(RTrim(Str(BNo))) in ('" + adc("WRFIDSTR") + "')) and BLoc<>'PFG' and BCls='N' ", wCnd)
      'vk.4 ADC("WRFIDSTR") replaced with wRfIdStr
      wCnd = IIF(LTrim(wRfIdStr) <> "" And ADC("WCPYTHRUBAGNO") = "Y", " charindex(','+BYy + '/' + BChr + '/' + LTrim(RTrim(str(BNo)))+',' , '," + wRfIdStr + ",') <> 0  and BLoc<>'PFG' and BCls='N' ", wCnd)      '*************Bhavna
      If wCnd = "" Then DispMsg "Scope Cannot Be Blank", etError: Exit Function
      wCnd = IIF(wCnd <> "", " and " + wCnd, "")
      Set wRsCB = moCn.OpenRes("Select Bag.* from Bag  " + _
                              "Join Txnd On TdCoCd= BCoCd and TdYy= BLstYy and TdKey= BLstKey and TdPrtKey=BPrtKey and TdYyKey=BLstYYKey " + _
                              "and (Select PValue from Param where PTyp= 'TC' and PMCd= TdTc)= 'REJ' " + _
                              "where BCoCd= '" + ADC("TCoCd") + "' and BCls= 'N' and BPrtKey='" + ctCurrPrtn + "' " + _
                              wCnd + " Order By BYy, BChr, BNo")
      wTotRows = wRsCB.RecCount
                             
      If wTotRows = 0 Then DispMsg "Bags not Found", etError: Exit Function
      
      If wTotRows > 999 Then DispMsg "No of Bags is greater than 999, reduce the scope", etError: Exit Function
      
      
      GrdTxndA.StartCopy
      MWLib.BeginProcess Me, "Copying Bag Details ..."
      With GrdTxndA
        Do While Not (wRsCB.EOF Or wRsCB.BOF)
          .AddItem
          wRow = .Rows - 1
          .Value(wRow, "TdSr") = .MaxVal("TdSr") + 1
          .Value(wRow, "TdSrNo") = 0
          .Value(wRow, "TdBYy") = wRsCB!BYy
          .Value(wRow, "TdBChr") = wRsCB!BChr
          .Value(wRow, "TdBNo") = wRsCB!BNo
          .Value(wRow, "TdFrBLoc") = wRsCB!BLoc
          .Value(wRow, "TdRjQty") = wRsCB!BRjQty
          .Value(wRow, "TdRjPtQty") = wRsCB!BRjPtQty
          .Value(wRow, "TdWrk") = wRsCB!BWrk
          .Value(wRow, "TdDesc") = ""
          .Value(wRow, "TdKey") = moCn.AutoGen("Txnd", "TdKey", ADC("TCoCd") + ADC("TYy"))
          .Value(wRow, "TdPrevYy") = wRsCB!BLstYy
          .Value(wRow, "TdPrevKey") = wRsCB!BLstKey
          .Value(wRow, "TdPrtKey") = ctCurrPrtn
          .SaveRec (wRow)
          .Row = wRow
          Call EmptyRm
          wRsCB.MoveNext
        Loop
      End With
    
      MWLib.EndProcess Me
      GrdTxndA.EndCopy
  End If
  
  
    With GrdTxndA
    'vk.4 ADC("WRFIDSTR") replaced with wRfIdStr
    If LTrim(wRfIdStr) <> "" And ADC("WCPYTHRUBAGNO") = "N" Then
        Dim wRsRfId As MwfLib.MDORowSet
        'vk.4 ADC("WRFIDSTR") replaced with wRfIdStr
        Set wRsRfId = moCn.OpenRes("Select RimNo from RfIdMst where Not Exists(Select 'x' From Txnd where " + _
                                        " TdCoCd=RimBCoCd and TdBYy=RimBYY and TdBChr=RimBChr and TdBNo=RimBNo and " + _
                                        " TdCoCd='" + ADC("TCoCd") + "' and TdTc='" + ADC("TTc") + "' and TdYy='" + ADC("TYy") + "' and TdChr='" + ADC("TChr") + "' and TdNo=" + CStr(ADC("TNo")) + _
                                        " and  TdSr>=" + CStr(wMinSr) + " and TdSr<=" + CStr(.MaxVal("TdSr")) + ")" + _
                                        " and RimNo in (" + wRfIdStr + ")")
      
      
          wRfIdStr = ""
          Do While Not (wRsRfId.EOF Or wRsRfId.BOF)
                wRfIdStr = wRfIdStr + IIF(wRfIdStr = "", "", ",") + CStr(wRsRfId!RimNo)
            wRsRfId.MoveNext
          Loop
          'vk.4 RfId changed as Bag Alias No.
            If wRfIdStr <> "" Then DispMsg "Total Bag Alias No. {" + CStr(wTotRfid) + " } :Bag Alias No. Not Copied :  " + wRfIdStr, etError: Exit Function
    End If
    End With
  
  CopyBag = True
  ' **** Manali 3.6.0 - 26/10/09 - Copy Multi Bags into Melting
  Set wRsCB = Nothing
  mb_CopyBag = False
End Function
Private Sub GenBagForRet()
  '*** (Jen 2.13)
  
  '****** Procedure to Issue RM into Multple Bags for a return order in case of 'DT' entry ******
  
  Dim wRsCB As MwfLib.MDORowSet, wTotRows As Single, wi_ARow As Single, ws_Cnd As String
  Dim wHldCnt As Integer

  '****** Sachin 3.02 24-12-07 - Current Partition
  If Not moCn.RecSeek("Select 'x' from OrdMst where OmCoCd= '" + ADC("TCoCd") + "' and OmTc= '" + ADC("wCrOdTc") + "' " + _
     "and OmYy= '" + ADC("wCrOdYy") + "' and OmChr= '" + ADC("wCrOdChr") + "' and OmNo= " + CStr(ADC("wCrOdNo")) + " " + _
     "and OmReNo> 0 and OmPrtKey='" + ctCurrPrtn + "' ") Then _
     DispMsg "Invalid Source Order Number/ Order Not A Return Order", etError: Exit Sub
  
  If ADC("wCrBLoc") = "" Then DispMsg "The Bag Loc Is Compulsory", etError: Exit Sub
  If ADC("wCrBLoc") = ctFgLoc Then DispMsg "The Bag Location Cannot Not Be '" + ctFgLoc + "'", etError: Exit Sub

  If gs_ValidPrdLocSeq <> "" And _
     InStr(1, "," + UCase(gs_ValidPrdLocSeq) + ",", "," + UCase(ADC("wCrBLoc")) + ",") <= 0 Then _
     DispMsg "The User Does Not Have Access To The Bag Loc", etError: Exit Sub
  
     ws_Cnd = ADC.RepCond
     ws_Cnd = IIF(ws_Cnd <> "", " and " + ws_Cnd, "")
     '*******Bhavna added field RedGrsWtDiff to match mainmet wt with edited GrsWt
     '****** Sachin 3.02 24-12-07 - Current Partition
     Set wRsCB = moCn.OpenRes("Select BYy, BChr, BNo, BQty, BPtQty, BGrWt, BLoc, BWrk, OdHld, RedSr, RedIdTc, RedIdYy, " + _
                               "RedIdChr, RedIdNo, RedIdSr, RedGrsWtDiff from Bag " + _
                               "join OrdDsg on OdIdNo=BOdIdNo And OdCoCd= BCoCd and OdTc= BOdTc and OdYy= BOdYy and OdChr= BOdChr " + _
                               "and OdNo=BOdNo and OdSr=BOdSr " + IIF(gs_Partition = ctCurrPrtn, " And OdPrtKey=BPrtKey ", "") + _
                               "join OrdMst on OmIdNo=OdOmIdNo and OmCoCd= OdCoCd and OmTc= OdTc and OmYy= OdYy and OmChr= OdChr and OmNo= OdNo " + IIF(gs_Partition = ctCurrPrtn, " And OmPrtKey=OdPrtKey ", "") + _
                               "join RetMemDet on RedCoCd= OmCoCd and RedYy= OmReYy and RedChr= OmReChr and RedNo= OmReNo " + _
                               "and RedSr=BOdSr " + _
                               "Where BCoCd= '" + ADC("TCoCd") + "' " + ws_Cnd + " and BCls= 'N' " + _
                               "and convert(decimal(16, 3), BGrWt)= 0 and BPrtKey='" + ctCurrPrtn + "' " + _
                               "Order By BYy, BChr, BNo")     'and BLstYy= '' and BLstKey= 0

  wTotRows = wRsCB.RecCount
  If wTotRows = 0 Then DispMsg "No Bags are Found for This Order.", etError: GoTo CopyEnd
  
  GrdTxndA.StartCopy
  MWLib.BeginProcess Me, "Copying Bag Details ..."
  
  Do While Not (wRsCB.EOF Or wRsCB.BOF)
    If wRsCB!OdHld = "Y" Then wHldCnt = wHldCnt + 1
    
    With GrdTxndA
      .AddItem
      wi_ARow = .Rows - 1
      .Row = wi_ARow
      '.Value(wi_ARow, "TdSr") = .MaxVal("TdSr") + 1
      .Value(wi_ARow, "TdSrNo") = 0
      .Value(wi_ARow, "TdBYy") = wRsCB!BYy
      .Value(wi_ARow, "TdBChr") = wRsCB!BChr
      .Value(wi_ARow, "TdBNo") = wRsCB!BNo
      .Value(wi_ARow, "TdBQty") = wRsCB!BQty
      .Value(wi_ARow, "TdPtQty") = wRsCB!BPtQty
      .Value(wi_ARow, "TdBGrWt") = wRsCB!BGrWt
      .Value(wi_ARow, "TdFrBLoc") = wRsCB!BLoc
      .Value(wi_ARow, "TdToBLoc") = ""
      .Value(wi_ARow, "TdWrk") = wRsCB!BWrk
      .Value(wi_ARow, "TdAck") = ""
      .Value(wi_ARow, "TdPrdYN") = "N"
      .Value(wi_ARow, "TdKey") = moCn.AutoGen("Txnd", "TdKey", ADC("TCoCd") + ADC("TYy"))
      .Value(wi_ARow, "TdPrevYy") = ""
      .Value(wi_ARow, "TdPrevKey") = 0
      .Value(wi_ARow, "TdPrtKey") = ctCurrPrtn    '****** Sachin 3.02.0
    End With
    If GrdTxndA.SaveRec(wi_ARow) = False Then GoTo NxtBag
  
    '*** Routine To Copy Rm into The Bags
    
    ' ******* Manali 3.5.0 - Change in Return Memo Copy
    If ADC("wCpyOrdRmYn") = "Y" Then
      Dim wRs_ORm As MwfLib.MDORowSet
      Set wRs_ORm = moCn.OpenRes("Select max(OrRmCtg) as OrRmCtg, max(OrRmSCtg) as OrRmsCtg, OrRmCd, Sum(OrPrdWt) as OrPrdWt  " + _
                                "From OrdRm " + _
                              "Join Bag On BCoCd=OrCoCd and BOdTc=OrTc and BOdYy=OrYy and BOdChr=OrChr and BOdNo=OrNo and BOdSr=OrSr " + _
                              "Where BCoCd='" + ADC("TCoCd") + "' and BYy='" + GrdTxndA.Value(GrdTxndA.Row, "TdBYy") + "' " + _
                              "and BChr='" + GrdTxndA.Value(GrdTxndA.Row, "TdBChr") + "' and BNo=" + CStr(GrdTxndA.Value(GrdTxndA.Row, "TdBNo")) + " " + _
                              "Group By OrRmCd " + _
                              "Order By (Case Max(OrRmCtg) When 'G' Then 1 When 'P' Then 2 When 'S' Then 3 When 'L' Then 4 When 'D' Then 5 When 'C' Then 6 " + _
                              "When 'X' Then 7 When 'M' Then 8 Else 9 End), OrRmSCtg, OrRmCd ")
    ' ******* Manali 3.5.0 - Change in Return Memo Copy
    Else
      Dim wf_MulFact As Single
      Dim wi_BRow As Single, wf_TotQty As Single, wf_TotRetQty As Single
      Dim wRsCr As MwfLib.MDORowSet
      
      '****** Getting The Total Iss Qty, Rec Qty ******
      '****** Sachin 3.02 24-12-07 - Current Partition
      wf_TotQty = moCn.GetFldVal("Select sum(IfQty) from InvFgd where IfCoCd= '" + ADC("TCoCd") + "' and " + _
                  "IfTc= '" + wRsCB!RedIdTc + "' and IfYy= '" + wRsCB!RedIdYy + "' and " + _
                  "IfChr= '" + wRsCB!RedIdChr + "' and IfNo= " + CStr(wRsCB!RedIdNo) + " and " + _
                  "IfSr= " + CStr(wRsCB!RedIdSr) + " and IfPrtKey='" + ctCurrPrtn + "' ")
      wf_TotRetQty = GrdTxndA.Value(GrdTxndA.Row, "TdBQty")
      
      wf_MulFact = IIF(wf_TotQty > wf_TotRetQty, MWLib.Div(wf_TotRetQty, wf_TotQty), 1)
      
      '*********Bhavna changed wt for mainmet = FrRmWt +/- edited GrsWt in RetMemo
      '****** Sachin 3.02 24-12-07 - Current Partition
      Set wRsCr = moCn.OpenRes("Select FrRmCd, FrRmSz, FrRmSz2,FrRmSz3,  FrRmStkRt, /* FrLotNo, */ Sum(FrRmQty) as FrRmQty, " + _
                  "Sum(FrRmWt) as FrRmWt, RmCtg " + _
                  "from InvFgd join Fgd on FdIdNo=IfFdIdNo And FdCoCd= IfCoCd and FdTc= IfFdTc and FdYy= IfFdYy and FdChr= IfFdChr " + _
                  "   and FdNo= IfFdNo and FdSr= IfFdSr and FdPrtKey=IfPrtKey " + _
                  "Join FgRm On FdIdNo=FrFdIdNo And FdCoCd= FrCoCd and FdTc= FrTc and FdYy= FrYy and FdChr= FrChr and FdNo= FrNo and FdSr= FrSr and FdPrtKey=FrPrtKey " + _
                  "Join RmMst on RmCd = FrRmCd " + _
                  "Join Loc LFr on LFr.LocCoCd= FrCoCd and LFr.LocCd= FrFrRmLoc " + _
                  "Join Loc LTo on LTo.LocCoCd= FrCoCd and LTo.LocCd= FrToRmLoc " + _
                  "Where IfCoCd= '" + ADC("TCoCd") + "' and IfTc= '" + wRsCB!RedIdTc + "' " + _
                  "   and IfYy= '" + wRsCB!RedIdYy + "' and IfChr= '" + wRsCB!RedIdChr + "' " + _
                  "   and IfNo= " + CStr(wRsCB!RedIdNo) + " and IfSr= " + CStr(wRsCB!RedIdSr) + " " + _
                  "   and (Case when LFr.LocTyp= 'F' and FrFrRmDc= 'D' then 1 when LTo.LocTyp= 'F' and FrToRmDc= 'D' then 1 " + _
                  "Else 0 end)= 1 Group By RmCtg, FrRmCd, /* FrLotNo, */ FrRmSz, FrRmSz2, FrRmSz3, FrRmStkRt")
    End If
    GrdTxndB.StartCopy
    MWLib.BeginProcess Me, "Copying Raw Material Details ..."
    
    With GrdTxndB
      If ADC("wCpyOrdRmYn") = "N" Then
        Do While Not (wRsCr.EOF Or wRsCr.BOF)
          .AddItem
          wi_BRow = .Rows - 1
          .Value(wi_BRow, "TdSrNo") = wi_BRow
          .Value(wi_BRow, "TdFrRmLoc") = ADC("TFrRmLoc")
          .Value(wi_BRow, "wTdFrRmIr") = "I"
          '.Value(wi_BRow, "TdFrRmDc") = "C"
          .Value(wi_BRow, "TdRmCd") = wRsCr!FrRmCd
          .Value(wi_BRow, "TdLotNo") = "" 'wRsCr!FrLotNo
          .Value(wi_BRow, "TdRmSz") = wRsCr!FrRmSz
          .Value(wi_BRow, "TdRmSz2") = wRsCr!FrRmSz2
          .Value(wi_BRow, "TdRmSz3") = wRsCr!FrRmSz3
          .Value(wi_BRow, "TdRmStkRt") = wRsCr!FrRmStkRt
          .Value(wi_BRow, "TdRmQty") = wRsCr!FrRmQty * wf_MulFact
          '************Bhavna
          If wRsCr!RmCtg = "G" Then
               .Value(wi_BRow, "TdRmWt") = (wRsCr!FrRmWt + wRsCB!RedGrsWtDiff) * wf_MulFact
          Else
            .Value(wi_BRow, "TdRmWt") = wRsCr!FrRmWt * wf_MulFact
          End If
          '**************Bhavna
          .Value(wi_BRow, "TdToRmLoc") = ADC("TToRmLoc")
          .Value(wi_BRow, "TdPrtKey") = ctCurrPrtn    '****** Sachin 3.02.0
          If .SaveRec(wi_BRow) = False Then
            .DelAllRecs
            DispMsg "Cannot Copy RM For The Bag: " + wRsCB!BYy + "/" + wRsCB!BChr + "/" + CStr(wRsCB!BNo) + " ", etError
            Exit Do
          End If
          wRsCr.MoveNext
        Loop
      Else
        ' ******* Manali 3.5.0  - Change in Return Memo Copy (IssueRm routine called)
        Do While Not (wRs_ORm.EOF Or wRs_ORm.BOF)
          ADC("wIrRmCd") = wRs_ORm!OrRmCd
          ADC("wIrRmAddWt") = wRs_ORm!OrPrdWt * wRsCB!BQty
          Call IssueRm
          If .Rows > 1 Then
            wRs_ORm.MoveNext
          Else
            DispMsg "Cannot Copy RM For The Bag: " + wRsCB!BYy + "/" + wRsCB!BChr + "/" + CStr(wRsCB!BNo) + " ", etError
            Exit Do
          End If
        Loop
      End If
      ' ******* Manali 3.5.0 - Change in Return Memo Copy (IssueRm routine called)

    End With
    
    GrdTxndB.EndCopy
    
    If GrdTxndB.Rows > 1 Then
      '****** Calculation Of TdBGrsWt ******
      '****** Sachin 3.02 24-12-07 - Current Partition
      GrdTxndA.Value(GrdTxndA.Row, "TdBGrWt") = moCn.GetFldVal("Select BGrWt from Bag where " + _
                                                "BCoCd= '" + ADC("TCoCd") + "' and " + _
                                                "BYy= '" + GrdTxndA.Value(GrdTxndA.Row, "TdBYy") + "' and " + _
                                                "BChr= '" + GrdTxndA.Value(GrdTxndA.Row, "TdBChr") + "' and " + _
                                                "BNo= " + CStr(GrdTxndA.Value(GrdTxndA.Row, "TdBNo")) + " and BPrtKey='" + ctCurrPrtn + "' ")
      GrdTxndA.SaveRec (GrdTxndA.Row)
    End If
    
    Set wRsCr = Nothing
    '*** Routine To Copy Rm into The Bags
  
NxtBag:
    wRsCB.MoveNext
  Loop

  
CopyEnd:
  
  MWLib.EndProcess Me
  GrdTxndA.EndCopy
  '*** End of the copy process
 
  Set wRsCB = Nothing
  
  If wHldCnt > 0 Then _
  DispMsg CStr(wHldCnt) + " Order Design(s) on Hold", etWarning
  
  '*** (Jen 2.13)
End Sub
Private Sub DispFra(ByVal pv_TxnBagFra As en_TxnBagFra)
  '*** This is the common procedure which is used to toggle between hiding and displaying
      'the frames (i.e. if the frame is visible then make it invisible and viseversa)
      'associated with a command button
  
  Select Case pv_TxnBagFra
  Case Is = PRDPTS
    If FraTP.Visible = True Then
      FraTP.Visible = False
      FraTP.Enabled = False
      '*** required to pass the array position PrdPts-0, to resotre the 'Causes Validation' State
      Call EnaDisaCmds(False, CmdTP, PRDPTS)
      'CmdTP.SetFocus
      GrdTxndA.SetFocus
    Else
      FraTP.Visible = True
      FraTP.Enabled = True
      GrdTP.SetFocus
      FraTP.ZOrder
      Call EnaDisaCmds(True, CmdTP)
    End If
'*****************Bhavna
 Case Is = Rwrk
    If FraRwrk.Visible = True Then
      FraRwrk.Visible = False
      FraRwrk.Enabled = False
      Call EnaDisaCmds(False, CmdRwrk)
      GrdTxndA.SetFocus
    ElseIf ms_RwrkModYN Then
      FraRwrk.Visible = True
      FraRwrk.Enabled = True
      GrdRwrk.SetFocus
      FraRwrk.ZOrder
      Call EnaDisaCmds(True, CmdRwrk)
    End If
 '*******************Bhavna
   Case Is = BagDet
    If FraBD.Visible = True Then
      FraBD.Visible = False
      FraBD.Enabled = False
      Call EnaDisaCmds(False, CmdBD, BagDet)
      'CmdBD.SetFocus
      GrdTxndA.SetFocus
    Else
      FraBD.Visible = True
      FraBD.Enabled = True
      FraBD.ZOrder
      Call EnaDisaCmds(True, CmdBD)
      CmdBD.SetFocus
    End If
  Case Is = IssRm
    If FraIR.Visible = True Then
      FraIR.Visible = False
      FraIR.Enabled = False
      Call EnaDisaCmds(False, CmdIR, IssRm)
      'CmdIR.SetFocus
      GrdTxndA.SetFocus
    Else
      FraIR.Visible = True
      FraIR.Enabled = True
      ADC("wIrRmCd").SetFocus
      FraIR.ZOrder
      Call EnaDisaCmds(True, CmdIR)
    End If
  Case Is = BagRmSumm
    If FraBS.Visible = True Then
      FraBS.Visible = False
      FraBS.Enabled = False
      Call EnaDisaCmds(False, CmdBS, BagRmSumm)
      'CmdBS.SetFocus
      GrdTxndA.SetFocus
    Else
      FraBS.Visible = True
      FraBS.Enabled = True
      CmdBSGo.SetFocus
      FraBS.ZOrder
      Call EnaDisaCmds(True, CmdBS)
    End If
  Case Is = VchSumm
    If FraVS.Visible = True Then
      FraVS.Visible = False
      FraVS.Enabled = False
      Call EnaDisaCmds(False, CmdVS, VchSumm)
      'CmdVS.SetFocus
      GrdTxndA.SetFocus
    Else
      FraVS.Visible = True
      FraVS.Enabled = True
      FraVS.ZOrder
      Call EnaDisaCmds(True, CmdVS)
      CmdVS.SetFocus
    End If
  Case Is = Tot
    If FraTot.Visible = True Then
      FraTot.Visible = False
      FraTot.Enabled = False
      Call EnaDisaCmds(False, CmdTot, Tot)
      'CmdTot.SetFocus
      GrdTxndA.SetFocus
    Else
      FraTot.Visible = True
      FraTot.Enabled = True
      FraTot.ZOrder
      Call EnaDisaCmds(True, CmdTot)
      CmdTot.SetFocus
    End If
  Case Is = Loss
    If FraLs.Visible = True Then
      FraLs.Visible = False
      FraLs.Enabled = False
      Call EnaDisaCmds(False, CmdLs, Loss)
      'CmdLs.SetFocus
      GrdTxndA.SetFocus
    Else
      FraLs.Visible = True
      FraLs.Enabled = True
      FraLs.ZOrder
 '     ADC("WLSMAINMETALLYN").SetFocus
      Call EnaDisaCmds(True, CmdLs)
      CmdLs.SetFocus
    End If
  Case Is = CpyBag
    If FraCB.Visible = True Then
      FraCB.Visible = False
      FraCB.Enabled = False
      Call EnaDisaCmds(False, CmdCB, IssRm)
      CmdCB.SetFocus
      'GrdTxndA.SetFocus
    Else
      FraCB.Visible = True
      FraCB.Enabled = True
      ' **** Manali 3.6.0 - 26/10/09 - Copy Multi Bags into Melting - wCbFrBLoc
      If UCase(ADC.MenuCd) = UCase("BV") Then
        'vk.4 cursor setfocus will change as per scope/bag frame enable/disable status
        'ADC("wCbToBLoc").SetFocus
        If FraCBScope.Enabled = True Then ADC("wCbToBLoc").SetFocus Else ADC("wRfIdStr").SetFocus
      ElseIf UCase(ADC.MenuCd) = UCase("REJ") Then
        'ADC("WCBBYYFR").SetFocus
        'vk.4 cursor setfocus will change as per scope/bag frame enable/disable status
        'ADC("wCbFrBLoc").SetFocus 'vk.4
        If FraCBScope.Enabled = True Then ADC("wCbFrBLoc").SetFocus Else ADC("wRfIdStr").SetFocus
      ElseIf UCase(ADC.MenuCd) = UCase("MLT") Then
        'vk.4 cursor setfocus will change as per scope/bag frame enable/disable status
        'ADC("wCbFrBLoc").SetFocus 'vk.4
        If FraCBScope.Enabled = True Then ADC("wCbFrBLoc").SetFocus Else ADC("wRfIdStr").SetFocus
      End If
      ' **** Manali 3.6.0 - 26/10/09 - Copy Multi Bags into Melting - wCbFrBLoc
      FraCB.ZOrder
      Call EnaDisaCmds(True, CmdCB)
    End If
  Case Is = GenBagDet
    If FraGBD.Visible = True Then
      FraGBD.Visible = False
      FraGBD.Enabled = False
      Call EnaDisaCmds(False, CmdGBD, GenBagDet)
      CmdGBD.SetFocus
      GrdTxndA.SetFocus
    Else
      FraGBD.Visible = True
      FraGBD.Enabled = True
      CmdGBDGo.SetFocus
      FraGBD.ZOrder
      Call EnaDisaCmds(True, CmdGBD)
    End If
' **** Zubin 211 **** '
  Case Is = FBQtyWt
    If FraQW.Visible = True Then
      FraQW.Visible = False
      FraQW.Enabled = False
      Call EnaDisaCmds(False, CmdQW, FBQtyWt)
      CmdQW.SetFocus
      GrdTxndB.SetFocus
    Else
      FraQW.Visible = True
      FraQW.Enabled = True
      ADC("wTdFbRmQty").SetFocus
      FraQW.ZOrder
      Call EnaDisaCmds(True, CmdQW)
    End If
' **** Zubin 211 **** '
  ' Zubin 212
  Case Is = CpyFgBagRm
    If FraCRM.Visible = True Then
      FraCRM.Visible = False
      FraCRM.Enabled = False
      Call EnaDisaCmds(False, CmdCRM, CpyFgBagRm)
      'CmdCRM.SetFocus
      GrdTxndA.SetFocus
    Else
      FraCRM.Visible = True
      FraCRM.Enabled = True
      ADC("wFgBYy").SetFocus
      FraCRM.ZOrder
      Call EnaDisaCmds(True, CmdCRM)
    End If
    ' Zubin 212
  
  '*** (Jen 2.13)
  Case Is = CpyRet
    If FraCpyRet.Visible = True Then
      FraCpyRet.Visible = False
      FraCpyRet.Enabled = False
      Call EnaDisaCmds(False, CmdCpyRet, CpyRet)
      'CmdCRM.SetFocus
      GrdTxndA.SetFocus
    Else
      FraCpyRet.Visible = True
      FraCpyRet.Enabled = True
      'ADC("wCrOdTc").SetFocus
      ADC("wCrBLoc").SetFocus
      FraCpyRet.ZOrder
      Call EnaDisaCmds(True, CmdCpyRet)
    End If
  '*** (Jen 2.13)
  
  '*****Bhavna added new lot
  Case Is = LotNo
    If FraLotNo.Visible = True Then
      FraLotNo.Visible = False
      Call EnaDisaCmds(False, CmdLN, LotNo)
      GrdTxndB.SetFocus
    Else
      FraLotNo.Visible = True
      ADC("wLotNo").SetFocus
      FraLotNo.ZOrder
      Call EnaDisaCmds(True, CmdLN)
    End If
  Case Is = Cert
    If FraCert.Visible = True Then
      FraCert.Visible = False
      FraCert.Enabled = False
      Call EnaDisaCmds(False, CmdCert, Cert)
      GrdTxndA.SetFocus
    Else
      FraCert.Visible = True
      FraCert.Enabled = True
      CmdCert.SetFocus
      FraCert.ZOrder
      Call EnaDisaCmds(True, CmdCert)
    End If
 '******* Bhavna
 'Uma KBS
 Case Is = ImpExcel
    If FraExImp.Visible = True Then
      FraExImp.Visible = False
      FraExImp.Enabled = False
      Call EnaDisaCmds(False, CmdIE, ImpExcel)
      GrdTxndA.SetFocus
    Else
      FraExImp.Visible = True
      FraExImp.Enabled = True
      'ADC("WXLSHTNO").SetFocus
      CmdXlImport.SetFocus
      FraExImp.ZOrder
      Call EnaDisaCmds(True, CmdIE)
    End If
  'ec.4
  Case Is = HlpSrt
    If FraHlpSrt.Visible = True Then
      FraHlpSrt.Visible = False
      FraHlpSrt.Enabled = False
      Call EnaDisaCmds(False, CmdHlpSrt, HlpSrt)
      CmdHlpSrt.SetFocus
    Else
      FraHlpSrt.Visible = True
      FraHlpSrt.Enabled = True
      ChkHSMtchSfx.SetFocus
      FraHlpSrt.ZOrder
      Call EnaDisaCmds(True, CmdHlpSrt)
    End If
      
  End Select
  
End Sub
Private Sub EnaDisaCmds(ByVal pv_ShowFra As Boolean, Optional ByVal pv_CmdBut As MWCTL_BTN1, Optional ByVal pv_TxnBagFra As en_TxnBagFra)
  '*** This procedure is used to give the toggle option (enable and disable) to
      'the command buttons
  '*** storing the state of cmd buttons, when the frame is going to be vis.
  
  Dim X As Integer
  If pv_ShowFra = True Then
    Call StoreState
  End If
  
  '*** (Jenny Color)
  If Not (pv_CmdBut Is Nothing) And pv_ShowFra = True Then
    'Set mo_CurrentBtn = pv_CmdBut
    ms_CurrentBtn = UCase(pv_CmdBut.Name)
  Else
    'Set mo_CurrentBtn = Nothing
    ms_CurrentBtn = ""
  End If
  '*** (Jenny Color)
  
  If Not pv_CmdBut Is Nothing Then
    If pv_ShowFra = True Then
      '*** (Bef Color)
      'pv_CmdBut.BackColor = vbWhite
      'pv_CmdBut.FontBold = True
      '*** (Bef Color)
      If pv_CmdBut.Pic = 0 Then
        pv_CmdBut.BackColor = vbWhite
        'pv_CmdBut.FontBold = True
      End If
      pv_CmdBut.CausesValidation = True
    Else
      '*** (bef Color)
      'pv_CmdBut.BackColor = &H8000000F
      'pv_CmdBut.FontBold = False
      '*** (Bef Color)
      If pv_CmdBut.Pic = 0 Then
        pv_CmdBut.BackColor = &H8000000F
        'pv_CmdBut.FontBold = False
      End If
      '### pv_CmdBut.CausesValidation = False
      '*** restoring the 'Causes Validation' state of the current pv_cmdbutton
      If Not IsMissing(pv_TxnBagFra) Then
        pv_CmdBut.CausesValidation = mArr_Cv(pv_TxnBagFra)
      End If
    End If
  End If
  FraNKeyAll.Enabled = Not pv_ShowFra
  ADC.AllowSave = Not pv_ShowFra
  ADC.AllowDelete = Not pv_ShowFra
  
  If pv_ShowFra = True Then
    '** All command buttons are disabled
    CmdTP.Enabled = Not pv_ShowFra
    CmdBD.Enabled = Not pv_ShowFra
    CmdIR.Enabled = Not pv_ShowFra
    CmdBS.Enabled = Not pv_ShowFra
    CmdVS.Enabled = Not pv_ShowFra
    CmdER.Enabled = Not pv_ShowFra
    CmdTot.Enabled = Not pv_ShowFra
    CmdLs.Enabled = Not pv_ShowFra
    CmdCB.Enabled = Not pv_ShowFra
    CmdCR.Enabled = Not pv_ShowFra
    CmdTR.Enabled = Not pv_ShowFra
    CmdGBD.Enabled = Not pv_ShowFra
    CmdQW.Enabled = Not pv_ShowFra   ' **** Zubin 211 **** '
    CmdCRM.Enabled = Not pv_ShowFra   ' Zubin 212
    CmdCpyRet.Enabled = Not pv_ShowFra    '*** (Jen 2.13)
    CmdRwrk.Enabled = Not pv_ShowFra    '*** Bhavna
    CmdLN.Enabled = Not pv_ShowFra      '*** Bhavna added new lot
    CmdCert.Enabled = Not pv_ShowFra      '*** Bhavna added cert details frame
    CmdIE.Enabled = Not pv_ShowFra  ' Uma KBS
    CmdHlpSrt.Enabled = Not pv_ShowFra 'ec.4
    CmdInvHd.Enabled = Not pv_ShowFra   'Sv.22
  Else
    '*** Restoring Ena/Disa state of all Command Buttons
    CmdTP.Enabled = mArr_Ed(PRDPTS)
    CmdBD.Enabled = mArr_Ed(BagDet)
    CmdIR.Enabled = mArr_Ed(IssRm)
    CmdBS.Enabled = mArr_Ed(BagRmSumm)
    CmdVS.Enabled = mArr_Ed(VchSumm)
    CmdCR.Enabled = mArr_Ed(CpyRm)
    CmdER.Enabled = mArr_Ed(EmpRm)
    CmdTot.Enabled = mArr_Ed(Tot)
    CmdLs.Enabled = mArr_Ed(Loss)
    CmdCB.Enabled = mArr_Ed(CpyBag)
    CmdTR.Enabled = mArr_Ed(TrfRm)
    CmdGBD.Enabled = mArr_Ed(GenBagDet)
    CmdQW.Enabled = mArr_Ed(FBQtyWt)  ' **** Zubin 211 **** '
    CmdCRM.Enabled = mArr_Ed(CpyFgBagRm)  ' Zubin 212
    CmdCpyRet.Enabled = mArr_Ed(CpyRet)    '*** (Jen 2.13)
    CmdRwrk.Enabled = mArr_Ed(CpyRet)    '*** Bhavna
    CmdLN.Enabled = mArr_Ed(LotNo)       '*** BHavna
    CmdCert.Enabled = mArr_Ed(Cert)   '*** Bhavna
    CmdIE.Enabled = mArr_Ed(ImpExcel)   ' Uma KBS
    CmdHlpSrt.Enabled = mArr_Ed(HlpSrt) 'ec.4
    CmdInvHd.Enabled = mArr_Ed(InvHd)   'Sv.22
  End If
  If Not pv_CmdBut Is Nothing Then pv_CmdBut.Enabled = True
 End Sub
Private Sub StoreState()
  '*** storing the Enabled and CausesValidation state of cmd buttons
  mArr_Cv(PRDPTS) = CmdTP.CausesValidation
  mArr_Cv(BagDet) = CmdBD.CausesValidation
  mArr_Cv(IssRm) = CmdIR.CausesValidation
  mArr_Cv(BagRmSumm) = CmdBS.CausesValidation
  mArr_Cv(VchSumm) = CmdVS.CausesValidation
  mArr_Cv(CpyRm) = CmdCR.CausesValidation
  mArr_Cv(EmpRm) = CmdER.CausesValidation
  mArr_Cv(Tot) = CmdTot.CausesValidation
  mArr_Cv(Loss) = CmdLs.CausesValidation
  mArr_Cv(CpyBag) = CmdCB.CausesValidation
  mArr_Cv(TrfRm) = CmdTR.CausesValidation
  mArr_Cv(GenBagDet) = CmdGBD.CausesValidation
  mArr_Cv(FBQtyWt) = CmdQW.CausesValidation   ' **** Zubin 211 **** '
  mArr_Cv(CpyFgBagRm) = CmdCRM.CausesValidation   ' Zubin 212
  mArr_Cv(CpyRet) = CmdCpyRet.CausesValidation    '*** (Jen 2.13)
  mArr_Cv(Rwrk) = CmdRwrk.CausesValidation    '*** Bhavna
  mArr_Cv(LotNo) = CmdLN.CausesValidation     '*** Bhavna added new lot
  mArr_Cv(Cert) = CmdCert.CausesValidation    '*** Bhavna added cert details
  mArr_Cv(ImpExcel) = CmdIE.CausesValidation   'Uma KBS
  mArr_Cv(HlpSrt) = CmdHlpSrt.CausesValidation   'ec.4
  mArr_Cv(InvHd) = CmdInvHd.CausesValidation    'Sv.22
  
  mArr_Ed(PRDPTS) = CmdTP.Enabled
  mArr_Ed(BagDet) = CmdBD.Enabled
  mArr_Ed(IssRm) = CmdIR.Enabled
  mArr_Ed(BagRmSumm) = CmdBS.Enabled
  mArr_Ed(VchSumm) = CmdVS.Enabled
  mArr_Ed(CpyRm) = CmdCR.Enabled
  mArr_Ed(EmpRm) = CmdER.Enabled
  mArr_Ed(Tot) = CmdTot.Enabled
  mArr_Ed(Loss) = CmdLs.Enabled
  mArr_Ed(CpyBag) = CmdCB.Enabled
  mArr_Ed(TrfRm) = CmdTR.Enabled
  mArr_Ed(GenBagDet) = CmdGBD.Enabled
  mArr_Ed(FBQtyWt) = CmdQW.Enabled    ' **** Zubin 211 **** '
  mArr_Ed(CpyFgBagRm) = CmdCRM.Enabled    ' Zubin 212
  mArr_Ed(CpyRet) = CmdCpyRet.Enabled   '*** (Jen 2.13)
  mArr_Ed(Rwrk) = CmdRwrk.Enabled   '*** Bhavna
  mArr_Ed(LotNo) = CmdLN.Enabled    '*** Bhavna
  mArr_Ed(Cert) = CmdCert.Enabled
  mArr_Ed(ImpExcel) = CmdIE.Enabled  'Uma KBS
  mArr_Ed(HlpSrt) = CmdHlpSrt.Enabled   'ec.4
  mArr_Ed(InvHd) = CmdInvHd.Enabled     'Sv.22
End Sub
Private Sub HideAllFras()
  '*** This procedure hides all the frames other than FraNKeyAll
  FraNKeyAll.Visible = True
  FraTP.Visible = False
  FraTP.Enabled = False
  FraBD.Visible = False
  FraBD.Enabled = False
  FraIR.Visible = False
  FraIR.Enabled = False
  FraBS.Visible = False
  FraBS.Enabled = False
  FraBRS.Visible = False
  FraBRS.Enabled = False
  FraVS.Visible = False
  FraVS.Enabled = False
  FraTot.Visible = False
  FraTot.Enabled = False
  FraLs.Visible = False
  FraLs.Enabled = False
  FraCB.Visible = False
  FraCB.Enabled = False
  FraGBD.Visible = False
  FraGBD.Enabled = False
  FraHlpSrt.Visible = False 'ec.4
  FraHlpSrt.Enabled = False 'ec.4
  ' **** Zubin 211 **** '
  FraQW.Visible = False
  FraQW.Enabled = False
  ' **** Zubin 211 **** '
  ' Zubin 212
  FraCRM.Visible = False
  FraCRM.Enabled = False
  ' Zubin 212
  
  '*** (Jen 2.13)
  FraCpyRet.Visible = False
  FraCpyRet.Enabled = False
  '*** (Jen 2.13)

  '*** Bhavna
  FraRwrk.Visible = False
  FraRwrk.Enabled = False
  '*** Bhavna
  FraExImp.Visible = False 'Uma KBS
  FraExImp.Enabled = False 'Uma KBS

End Sub

'''  ***** Std Code not to be changed **********************************
Private Sub Form_GotFocus()
  '*** Jenny Color (For Setting System Colors)
  SetMDISysColorsAfterLogin
  '*** Jenny Color (For Setting System Colors)
End Sub
Private Sub Form_Load()
  '*** Jenny Color
  '(Bef 13/08/05)
  'SetMDISysColorsAfterLogin   '*** Jenny Color (For Setting System Colors)
  '****** (Jen 2.13) added field CmdCpyRetGo ******
  Dim ws_BtnStr As String
  '****** (Jen 2.13) added CmdCpyRetGo to the list ******
  ' Zubin 212 (CmdCRMGo added)
  ws_BtnStr = "CmdBSGo,CmdCBGo,CmdLsGo,CmdIRGo,CmdQWOk,CmdCRMGo,CmdCpyRetGo,CmdLNGo"
  Call SetObjColors(Me, ws_BtnStr)
  '*** Jenny Color
  
  Me.KeyPreview = True      ' ***** Manali 3.10.0 - 12/04/12 - Bar Code
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
Private Sub GrdTxndA_Validate(Cancel As Boolean)
  Cancel = GrdTxndA.Validate
End Sub
Private Sub GrdTxndB_Validate(Cancel As Boolean)
  Cancel = GrdTxndB.Validate
End Sub
Private Sub GrdTP_Validate(Cancel As Boolean)
  Cancel = GrdTP.Validate
End Sub

' ########################  Manoj #### Ver: 2.0.6 #### Date: 25/12/2003  ########################
Private Sub TxtPwd_GotFocus()
  DispMsg "Enter SEO Password", etInfo
End Sub

Private Sub TxtPwd_LostFocus()
Call TxtPwd_Validate(False)
End Sub

Private Sub TxtPwd_Validate(Cancel As Boolean)
'mf_DtLocked = IIF((ADC("TDt") <> "" And ADC("TDt") < CDate(ms_LockDt)) _
'                                Or _
'                   (ADC("TDt") <> "" And ADC("TDt") > CDate(ms_FutureDt)) _
'                                Or _
'                  (ADC("TDt") <= CDate(ms_PurgeDt)), _
'                True, False)

ADC("TSeoPwd") = ""
'*** Jay 2.14 *** (SEO Change)*
'  If ADC.AddFndMode = xFndMode Then
    If Trim(TxtPwd.Text) <> "" Then
      If Trim(TxtPwd.Text) = moCn.GetFldVal("select IsNull ((select PValue from Param Join UsrAccess on UaUsrCd = PMCd " + _
                   "where PTyp= 'USR' and PMCd= '" + gs_UsrCd + "' and UaMnuCd = '" + ADC.MenuCd + "' and " + _
                   "UaSeoYn= 'Y' and PValue= '" + Trim(TxtPwd.Text) + "'), " + _
                   "(Select PValue1 from Param where Ptyp= 'SEOPWD' and PmCd= 'SEOPWD' )) ") Then
          ms_PwdFlag = True: ADC("TSeoPwd") = TxtPwd.Text
      Else
          ms_PwdFlag = False: Cancel = True: DispMsg "Invalid Edit Password / User Does Not Have SEO Rights ", etError
      End If
    Else
      ms_PwdFlag = moCn.RecSeek("Select PValue1 from Param where PTyp= 'SEOPWD' and PmCd= 'SEOPWD' and PValue1= '' ")
      If ms_PwdFlag = True Then ADC("TSeoPwd") = TxtPwd.Text
    End If
'  End If
  
'  If Trim$(txtPwd.text) = Trim$(ms_Pwd) Then
'    ADC("TSeoPwd") = txtPwd.text
'  Else
'    If Trim(txtPwd.text) <> "" Then
'      Cancel = True
'      DispMsg "Invalid SEO Password", etError
'    End If
'  End If
 ADC.SaveRec
End Sub

Private Sub GenBagDetails()

On Error GoTo ErrHndlr
  '****** Sachin 3.4.0 ******
  Dim mMsg As Integer
  mMsg = MsgBox("Please Confirm the Following before Cross Company Invoice Bags Copy " + vbCrLf + _
         "(1) Source and Target Companies should have the Same Order Tc-Yy-Chr-No-Sr." + vbCrLf + _
         "(2) Source and Target Design Code/Sfx/Sz should match " + vbCrLf + _
         "(3) Check Whether Stock exists in the Specified RM Location " + vbCrLf + _
         "(4) Source Raw Materials Should Exists In Target Bag Order ", vbOKCancel, "Invoice Bags Copy")
  If mMsg = vbOK Then
  '****** Sachin 3.4.0 ******
  
    If ADC("TFrBLoc") = "" Then DispMsg "Enter Bag Generation Location.", etError: Exit Sub
    If ADC("TFrBLoc") = "PFG" Then DispMsg "Bag Generation Location Cannot Not Be 'PFG'", etError: Exit Sub
    
    If ADC("WGBDCOCD") = "" Or ADC("WGBDINTC") = "" Or ADC("WGBDINYY") = "" Or ADC("WGBDINCHR") = "" Or ADC("WGBDINNO") = "" Then _
      DispMsg "Enter Company Code & Voucher No For Copying Bag", etError: Exit Sub
    
    If ADC("WGBDCOCD") = ADC("TCoCd") Then DispMsg "Current Company Code & From Company Code Can Not be Same", etError: Exit Sub
    If GrdTxndA.Row > 1 Then DispMsg "Record Already Exists for Bags", etError: Exit Sub
    
    Dim wCnd As String, wOpnLoc As String, wSrOnHld As String, wSrExistingBag  As String
    Dim wsBag As String, wsql As String
    Dim wOdBQty As Double, wRowA As Integer, wRowB As Integer, wRow As Integer
    Dim wRsBag As MwfLib.MDORowSet, wRsOdQty As MDORowSet
    Dim wRsBagGen As MDORowSet, wRsRm As MDORowSet
    Dim wRmCtg As String, wRmSCtg As String
    Dim wLocSeq, wPosComma, wBYy, wBChr
    Dim wTotLabVal As Double    '****** Sachin 3.4.1
  
    If ADC("WGBDDCFRLOC") = ADC("WGBDDCTOLOC") Then DispMsg "From and To Location Cannot be Same", etError: Exit Sub
  
  ' ###########################################  Manoj  ###########################################
  '  Generate Bag
    GrdTxndA.StartCopy
    MWLib.BeginProcess Me, "Copying Bag Details ..."
    
    ' ****** Check Whether Bags Exist In Source ******
    Set wRsBag = moCn.OpenRes("Select IfSr, IfSrNo, BYy, BChr, BNo, BQty, BPtQty, BGrwt, FdGrWt, FdExpOdTc As OdTc, FdExpOdYy As OdYy, FdExpOdChr As OdChr," & _
                  " FdExpOdNo As OdNo, FdExpOdSr As OdSr, FdPrdOdTc, FdPrdOdYy, FdPrdOdChr, FdPrdOdNo, FdPrdOdSr, FdDmCd, FdSfx, FdDmSz" & _
                  " From InvFgd InF " + _
                  "   Join Fgd Fd On IfFdIdNo=FdIdNo And IfCoCd = FdCoCD And IfFdTc = FdTc And IfFdYy =  FdYy And IfFdChr = FdChr" & _
                  "                 And IfFdNo=FdNo And IfFdSr = FdSr And IfPrtKey=FdPrtKey And IfFdBYy = FdBYy And IfFdBChr = FdBChr And IfFdBNo = FdBNo" & _
                  "   Join Bag B On BIdNo=IfBIdNo And BCoCd = IfCoCd And BYy = IfFdBYy And BChr = IfFdBChr And BNo = IfFdBNo And BPrtKey=IfPrtKey " & _
                  " Where IfCoCd = '" & ADC("WGBDCoCd") & "' And IfTc = '" & ADC("WGBDInTc") & "' And IfYy = '" & ADC("WGBDInYy") & _
                  "'  And IfChr = '" & ADC("WGBDInChr") & "' And IfNo = " & CStr(ADC("WGBDInNo")) & " And IfPrtKey='" + ctCurrPrtn + "'")
    If wRsBag.RecCount = 0 Then DispMsg "No Bags are Found for This Invoice.", etError: GoTo CopyEnd
    
    Do While Not (wRsBag.EOF Or wRsBag.BOF)
      If Not (wRsBag!OdTc = wRsBag!FdPrdOdTc And wRsBag!OdYy = wRsBag!FdPrdOdYy And wRsBag!OdChr = wRsBag!FdPrdOdChr _
              And wRsBag!OdNo = wRsBag!FdPrdOdNo And wRsBag!OdSr = wRsBag!FdPrdOdSr) Then
        wsBag = wsBag & vbCrLf & wRsBag!BYy & "/" & wRsBag!BChr & "/" & CStr(wRsBag!BNo) & " [Prd-Exp Order Mismatch]"
        GoTo NextBag
      End If
            
      '****** Sachin 3.4.1 ******
      '****** DsgCd-Sfx-Sz Should Match in Source and Target Orders
      If Not moCn.RecSeek("Select 'x' From OrdDsg Where OdCoCd = '" & ADC("TCoCd") & "' And OdTc = '" & wRsBag!OdTc & "' And OdYy = '" & wRsBag!OdYy & "'" & _
                        " And OdChr = '" & wRsBag!OdChr & "' And OdNo = " & CStr(wRsBag!OdNo) & " And OdSr = " & CStr(wRsBag!OdSr) & " And OdPrtKey='" + ctCurrPrtn + "' " & _
                        " And OdDmCd='" & wRsBag!FdDmCd & "' And OdSfx='" & wRsBag!FdSfx & "' And OdDmSz='" + wRsBag!FdDmSz + "'") Then
        wsBag = wsBag & vbCrLf & wRsBag!BYy & "/" & wRsBag!BChr & "/" & CStr(wRsBag!BNo & " [DsgCd-Sfx-Sz Mismatch]")
        GoTo NextBag
      End If
      '****** Sachin 3.4.1 ******
            
      Set wRsOdQty = moCn.OpenRes("Select IsNull(Sum(BQty), 0) As BQty, Max(OdPrdQty) As OdPrdQty From OrdDsg LEFT OUTER JOIN Bag On" & _
                        " BPrtKey=OdPrtKey And BOdIdNo=OdIdNo And BCoCd = OdCoCd And BOdTc = OdTc And BOdYy = OdYy And BOdChr = OdChr And BOdNo = OdNo And BOdSr = OdSr" & _
                        " Where OdCoCd = '" & ADC("TCoCd") & "' And OdTc = '" & wRsBag!OdTc & "' And OdYy = '" & wRsBag!OdYy & "'" & _
                        " And OdChr = '" & wRsBag!OdChr & "' And OdNo = " & CStr(wRsBag!OdNo) & " And OdSr = " & CStr(wRsBag!OdSr) & " And OdPrtKey='" + ctCurrPrtn + "'" + _
                        " Group By OdCoCd, OdTc, OdYy, OdChr, OdNo, OdSr")
      If wRsOdQty.RecCount > 0 Then
        wOdBQty = wRsBag!BQty + wRsOdQty!BQty
        If wOdBQty > wRsOdQty!OdPrdQty Then
          wsBag = wsBag & vbCrLf & wRsBag!BYy & "/" & wRsBag!BChr & "/" & CStr(wRsBag!BNo) & " [Bg Qty Exceeds Prd Qty]"
          GoTo NextBag
        End If
      Else
          wsBag = wsBag & vbCrLf & wRsBag!BYy & "/" & wRsBag!BChr & "/" & CStr(wRsBag!BNo) & " [Ord Does Not Exist In Tgt]"
          GoTo NextBag
      End If
      
      If moCn.RecSeek("Select 1 From Bag Where BPrtKey='" & ctCurrPrtn & "' And BCoCd = '" & ADC("TCoCd") & "'" & _
            " And BYy = '" & wRsBag!BYy & "' And BChr = '" & wRsBag!BChr & "' And BNo = " & CStr(wRsBag!BNo)) = True Then
        wsBag = wsBag & vbCrLf & wRsBag!BYy & "/" & wRsBag!BChr & "/" & CStr(wRsBag!BNo) & " [Bag Already Exists]"
        GoTo NextBag
      End If
  
      'MWLib.RunningProcess "Copying Details For Bag '" & wRsBag!BYy & "/" & wRsBag!BChr & "/" & CStr(wRsBag!BNo) + "'"
  
      Set wRsBagGen = moCn.OpenResultset("Select * From Bag Where BPrtKey='" & ctCurrPrtn & "' And BCoCd = '" & ADC("TCoCd") & "'" & _
                                         " And BYy = '" & wRsBag!BYy & "' And BChr = '" & wRsBag!BChr & "' And BNo = " & CStr(wRsBag!BNo))
      If wRsBagGen.RecCount > 0 Then GoTo NextBag
  
      '****** Sachin 3.4.1 - AddRecSet below to get values for new fields in Bag Table ******
      Dim wRsBagOdDet As MDORowSet
      Set wRsBagOdDet = moCn.OpenResultset("Select * From OrdDsg Where OdPrtKey='" & ctCurrPrtn & "' And OdCoCd = '" & ADC("TCoCd") & "' And OdTc='" & wRsBag!OdTc & "' " & _
                                         " And OdYy = '" & wRsBag!OdYy & "' And OdChr = '" & wRsBag!OdChr & "' And OdNo = " & CStr(wRsBag!OdNo) & " And OdSr = " & CStr(wRsBag!OdSr))
      If wRsBagOdDet.RecCount = 0 Then GoTo NextBag
  
      wRsBagGen.Addnew
        wRsBagGen.FldValue("BCoCd") = ADC("TCoCd")
        wRsBagGen.FldValue("BYy") = wRsBag!BYy
        wRsBagGen.FldValue("BChr") = wRsBag!BChr
        wRsBagGen.FldValue("BNo") = wRsBag!BNo
        wRsBagGen.FldValue("BOdTc") = wRsBag!OdTc
        wRsBagGen.FldValue("BOdYy") = wRsBag!OdYy
        wRsBagGen.FldValue("BOdChr") = wRsBag!OdChr
        wRsBagGen.FldValue("BOdNo") = wRsBag!OdNo
        wRsBagGen.FldValue("BOdSr") = wRsBag!OdSr
        wRsBagGen.FldValue("BOdDmCd") = wRsBag!FdDmCd
        wRsBagGen.FldValue("BOdSfx") = wRsBag!FdSfx
        wRsBagGen.FldValue("BOdDmSz") = wRsBag!FdDmSz
        wRsBagGen.FldValue("BOpnQty") = wRsBag!BQty
        wRsBagGen.FldValue("BOpnPtQty") = wRsBag!BPtQty
        wRsBagGen.FldValue("BOpnDt") = moCn.SrvrDate     'Date
        wRsBagGen.FldValue("BOpnLoc") = ADC("TFRBLOC")
        wRsBagGen.FldValue("BOpnDesc") = ""
        wRsBagGen.FldValue("BLoc") = ADC("TFRBLOC")
        wRsBagGen.FldValue("BCls") = "N"
        wRsBagGen.FldValue("BGrWt") = 0
        wRsBagGen.FldValue("BQty") = wRsBag!BQty
        wRsBagGen.FldValue("BPtQty") = wRsBag!BPtQty
        wRsBagGen.FldValue("BWrk") = ctTdWrk
        wRsBagGen.FldValue("BRjQty") = 0
        wRsBagGen.FldValue("BRjPtQty") = 0
        wRsBagGen.FldValue("BLstYy") = ""
        wRsBagGen.FldValue("BLstKey") = 0
        wRsBagGen.FldValue("BLstYyKey") = ""
        wRsBagGen.FldValue("BPrdStg") = ""
        wRsBagGen.FldValue("BFgSubLoc") = ""
        wRsBagGen.FldValue("BRecvDt") = moCn.SrvrDate
        wRsBagGen.FldValue("BOdCtCd") = ""
        wRsBagGen.FldValue("BClsDt") = MWLib.EmptyDate
        wRsBagGen.FldValue("ModUsr") = ADC.UsrCd
        wRsBagGen.FldValue("ModDt") = moCn.SrvrDate
        wRsBagGen.FldValue("ModTime") = moCn.SrvrTime
        wRsBagGen.FldValue("BOmIdNo") = wRsBagOdDet!OdOmIdNo
        wRsBagGen.FldValue("BOdIdNo") = wRsBagOdDet!OdIdNo
        wRsBagGen.FldValue("BDmIdNo") = wRsBagOdDet!OdDmIdNo
        wRsBagGen.FldValue("BPrtKey") = ctCurrPrtn
        wRsBagGen.FldValue("BMaxTxnDt") = moCn.SrvrDate
        wRsBagGen.FldValue("BCnvJtYn") = "N"      ' Manali - Trading Module - new field added
      wRsBagGen.Update
  
      If wRsBagGen.RecCount > 0 Then
  
  '      wsql = "Update AutoGen WITH (ROWLOCK) Set ANum = IsNull(qNum,0) From AutoGen JOIN" & _
  '            " (Select 'Bag' qTable, 'BNo' qFld, BCoCd+BYy+BChr qPreKey, max(BNo) qNum from Bag" & _
  '            " Where BCoCd = '" & ADC("TCoCd") & "' AND BYy = '" & wRsBag!BYy & "' AND BChr = '" & wRsBag!BChr & "'" & _
  '            " group by BCoCd, BYy, BChr) B ON ATable = qTable And AFld = qFld And APreKey = qPreKey"
  
  '****** Sachin 3.4.1 *** Autogen commented below as BagNo is not based on Autogen now
  '      wsql = "If Exists (Select 1 From AutoGen Where ATable = 'Bag' And AFld = 'BNo' And" & _
  '            " APreKey = '" & ADC("TCoCd") & wRsBag!BYy & wRsBag!BChr & "')" & _
  '            " Update AutoGen WITH (ROWLOCK) Set ANum = IsNull(qNum,0) From AutoGen JOIN" & _
  '            " (Select 'Bag' qTable, 'BNo' qFld, BCoCd+BYy+BChr qPreKey, max(BNo) qNum from Bag" & _
  '            " Where BPrtKey='" & ctCurrPrtn & "' And BCoCd = '" & ADC("TCoCd") & "' AND BYy = '" & wRsBag!BYy & "' AND BChr = '" & wRsBag!BChr & "'" & _
  '            " group by BCoCd, BYy, BChr) B ON ATable = qTable And AFld = qFld And APreKey = qPreKey" & _
  '            " Else" & _
  '            " Insert Into AutoGen WITH (ROWLOCK) (ATABLE,AFLD,APREKEY,ANUM)" & _
  '            " Select 'Bag' qTable, 'BNo' qFld, BCoCd+BYy+BChr qPreKey, max(BNo) qNum from Bag" & _
  '            " Where BPrtKey='" & ctCurrPrtn & "' And BCoCd = '" & ADC("TCoCd") & "' AND BYy = '" & wRsBag!BYy & "' AND BChr = '" & wRsBag!BChr & "'" & _
  '            " group by BCoCd, BYy, BChr"
  '      moCn.Execute (wsql)
        
   '****** Sachin 3.4.1 - Fetching Labour Value ******
   wTotLabVal = moCn.GetFldVal("Select Convert(Decimal(16,2),Sum(IsNull(qLabVal,0))*" + CStr(wRsBag!BQty) + ") From " + _
                              " (Select Sum(IlIVal/IdQty) as qLabVal from InvLab " + _
                              " Join InvDsg On IdCoCd=IlCoCd and IdTc=IlTc and IdYy=IlYy and IdChr=IlChr and IdNo=IlNo and IdSr=IlSr" + _
                              "     Where IlCoCd='" + ADC("WGBDCoCd") + "' and IlTc='" + ADC("WGBDInTc") + "' and IlYy='" + ADC("WGBDInYy") + "' and IlChr='" + ADC("WGBDInChr") + "' and IlNo=" + CStr(ADC("WGBDInNo")) + " and IlSr=" + CStr(wRsBag!IfSr) + _
                              " Union " + _
                              " Select Sum(IrSetIVal/IdQty) from InvRm " + _
                              " Join InvDsg On IdCoCd=IrCoCd and IdTc=IrTc and IdYy=IrYy and IdChr=IrChr and IdNo=IrNo and IdSr=IrSr" + _
                              "  Where IrCoCd='" + ADC("WGBDCoCd") + "' and IrTc='" + ADC("WGBDInTc") + "' and IrYy='" + ADC("WGBDInYy") + "' and IrChr='" + ADC("WGBDInChr") + "' and IrNo=" + CStr(ADC("WGBDInNo")) + " and IrSr=" + CStr(wRsBag!IfSr) + " and IrSetIVal<>0) LVal")
  '****** Sachin 3.4.1 ******
        
        GrdTxndA.AddItem
        wRowA = GrdTxndA.Rows - 1
        GrdTxndA.Row = wRowA
        GrdTxndA.Value(wRowA, "TdSr") = wRowA
        GrdTxndA.Value(wRowA, "TdSrNo") = 0
        GrdTxndA.Value(wRowA, "TdBYy") = wRsBag!BYy
        GrdTxndA.Value(wRowA, "TdBChr") = wRsBag!BChr
        GrdTxndA.Value(wRowA, "TdBNo") = wRsBag!BNo
        GrdTxndA.Value(wRowA, "TdBQty") = wRsBag!BQty
        GrdTxndA.Value(wRowA, "TdPtQty") = wRsBag!BPtQty
        GrdTxndA.Value(wRowA, "TdBGrWt") = wRsBag!BGrWt
        GrdTxndA.Value(wRowA, "TdFrBLoc") = ADC("TFrBLoc")
        GrdTxndA.Value(wRowA, "TdWrk") = ctTdWrk
        GrdTxndA.Value(wRowA, "TdAck") = ""
        GrdTxndA.Value(wRowA, "TdPrdYN") = "N"
        GrdTxndA.Value(wRowA, "TdKey") = moCn.AutoGen("Txnd", "TdKey", ADC("TCoCd") + ADC("TYy"))
        GrdTxndA.Value(wRowA, "TdPrevYy") = ""
        GrdTxndA.Value(wRowA, "TdPrevKey") = ""
        GrdTxndA.Value(wRowA, "TdDesc") = "Lab Val : " + CStr(wTotLabVal)
        GrdTxndA.Value(wRowA, "TdPrtKey") = ctCurrPrtn    '****** Sachin 3.02.0
              
        If GrdTxndA.SaveRec(wRowA) Then
          '****** Sachin 3.02 24-12-07 - Current Partition
          Set wRsRm = moCn.OpenRes("Select InvFgd.*, FrRmCd, FrLotNo, FrRmSz, FrRmSz2, FrRmSz3, FrRmStkRt, FrRmQty, FrRmWt from InvFgd Join" & _
                        " (Fgd Join FgRm On FdIdNo=FrFdIdNo And FdCoCd = FrCocd And FdTc = FrTc And FdYy = FrYy And FdChr = FrChr And FdNo = FrNo And FdSr = FrSr and FdPrtKey=FrPrtKey) On" & _
                        " IfCoCd = FdCoCd And IfFdTc = FdTc And IfFdYy = FdYy And IfFdChr = FdChr And IfFdNo = FdNo And IfFdSr = FdSr and IfPrtKey=FdPrtKey Where" & _
                        " IfCoCd = '" & ADC("WGBDCoCd") & "' And IfTc = '" & ADC("WGBDInTc") & "' And IfYy = '" & ADC("WGBDInYy") & "'" & _
                        " And IfChr = '" & ADC("WGBDInChr") & "' And IfNo = " & CStr(ADC("WGBDInNo")) & " And IfSr = " & CStr(wRsBag!IfSr) & " And IfSrNo = " & CStr(wRsBag!IfSrNo) & " And IfPrtKey='" + ctCurrPrtn + "'")
          GrdTxndB.StartCopy
          Do While Not (wRsRm.EOF Or wRsRm.BOF)
            GrdTxndB.AddItem
            wRowB = GrdTxndB.Rows - 1
            GrdTxndB.Row = wRowB
            GrdTxndB.Value(wRowB, "TdSrNo") = wRowB
            'Manoj 2.10.0
            wRmCtg = moCn.GetFldVal("Select RmCtg from RmMst where RmCd= '" + wRsRm!FrRmCd + "' And RmPrtKey='" + ctCurrPrtn + "' ")
            wRmSCtg = moCn.GetFldVal("Select RmSCtg from RmMst where RmCd= '" + wRsRm!FrRmCd + "' And RmPrtKey='" + ctCurrPrtn + "' ")
            'GrdTxndB.Value(wRowB, "TdFrRmLoc") = ADC("TFrRmLoc")
            'GrdTxndB.Value(wRowB, "TdToRmLoc") = ADC("TToRmLoc")
            If UCase$(wRmCtg) = "D" Or UCase$(wRmCtg) = "C" Then
              GrdTxndB.Value(wRowB, "TdFrRmLoc") = IIF(ADC("WGBDDCFRLoc") <> "", ADC("WGBDDCFRLoc"), ADC("TFrRmLoc"))
              GrdTxndB.Value(wRowB, "TdToRmLoc") = IIF(ADC("WGBDDCTOLoc") <> "", ADC("WGBDDCTOLoc"), ADC("TToRmLoc"))
            Else
              GrdTxndB.Value(wRowB, "TdFrRmLoc") = ADC("TFrRmLoc")
              GrdTxndB.Value(wRowB, "TdToRmLoc") = ADC("TToRmLoc")
            End If
            'Manoj 2.10.0
            GrdTxndB.Value(wRowB, "wTdFrRmIr") = "I"
            GrdTxndB.Value(wRowB, "TdFrRmDc") = "C"
            GrdTxndB.Value(wRowB, "TdRmCd") = wRsRm!FrRmCd
            ' Zubin 212
            If moCn.GetFldVal("Select HLotNoYN from Head where HCoCd='" + ms_TCoCd + "' and HCd='" + ctSelfCmCd + "' ") = "Y" Then
              GrdTxndB.Value(wRowB, "TdLotNo") = wRsRm!FrLotNo
            Else
              GrdTxndB.Value(wRowB, "TdLotNo") = ""
            End If
            ' Zubin 212
            GrdTxndB.Value(wRowB, "TdRmSz") = wRsRm!FrRmSz
            GrdTxndB.Value(wRowB, "TdRmSz2") = wRsRm!FrRmSz2
            GrdTxndB.Value(wRowB, "TdRmSz3") = wRsRm!FrRmSz3

''            If moCn.GetFldVal("Select hAvgSRtYn from Head where HCoCd= '" + ms_TCoCd + "' and " + _
''                            "HCd='" + ctSelfCmCd + "'") = "Y" Then
''              Dim wRmLocTyp As String, wb_RmLocTyp As Boolean, wGrpByRt As String
''              wb_RmLocTyp = moCn.RecSeek("Select LocTyp from Loc where LocCoCd= '" + ADC("TCoCd") + "' " + _
''                                         "and LocCd= '" + ADC("TFrRmLoc") + "' " + _
''                                         "and LocTyp in " + ctAvgLocTypSeq)
''              'Calculate Avg Rt
''              wRmLocTyp = Trim(moCn.GetFldVal("Select LocTyp from Loc where LocCoCd= '" + ADC("TCoCd") + "' " + _
''                                         "and LocCd= '" + ADC("TFrRmLoc") + "' "))
''
''              wGrpByRt = moCn.GetFldVal("Select PValue1 from Param where PTyp= 'RMSCTG' and " + _
''                         "PMCd= '" + wRmCtg + "' and PSCd= '" + wRmSCtg + "' ")
''
''              If wb_RmLocTyp = True And wGrpByRt = "Y" Then
''                If wRmLocTyp <> "W" Then
''                  ' Zubin 212
''                  GrdTxndB.Value(wRowB, "TdRmStkRt") = GetAvgStkRt(ADC("TCoCd"), ADC("TFrRmLoc"), _
''                                              wRsRm!FrRmCd, wRsRm!FrLotNo, wRsRm!FrRmSz)
''                Else
''                  ' Zubin 212
''                  GrdTxndB.Value(wRowB, "TdRmStkRt") = GetAvgStkRt(ADC("TCoCd"), ADC("TFrRmLoc"), _
''                                                wRsRm!FrRmCd, wRsRm!FrLotNo, wRsRm!FrRmSz, _
''                                                wRsBag!BYy, wRsBag!BChr, wRsBag!BNo)
''                End If
''              Else
''                GrdTxndB.Value(wRowB, "TdRmStkRt") = 0
''              End If
''            End If

            ' **** Manali 27/03/09 - StkRt from Source
            If moCn.GetFldVal("Select hAvgSRtYn from Head where HCoCd= '" + ms_TCoCd + "' and " + _
                            "HCd='" + ctSelfCmCd + "'") = "Y" Then
                 GrdTxndB.Value(wRowB, "TdRmStkRt") = wRsRm!FrRmStkRt
            Else
                GrdTxndB.Value(wRowB, "TdRmStkRt") = 0
            End If
            ' **** Manali 27/03/09 - StkRt from Source
            
            GrdTxndB.Value(wRowB, "TdRmQty") = wRsRm!FrRmQty
            GrdTxndB.Value(wRowB, "TdRmWt") = wRsRm!FrRmWt
            GrdTxndB.Value(wRowB, "TdPrtKey") = ctCurrPrtn    '****** Sachin 3.02.0

            If Not GrdTxndB.SaveRec(wRowB) Then
              wsBag = wsBag & vbCrLf & wRsBag!BYy & "/" & wRsBag!BChr & "/" & CStr(wRsBag!BNo) & " [Raw-Material Not Copied]"
              GrdTxndB.DelAllRecs
              Exit Do
            End If
            wRsRm.MoveNext
          Loop
          GrdTxndB.EndCopy
'
'  '        If moCn.RecSeek("Select 1 From (Select BCoCd, BOdTc, BOdYy, BOdChr, BOdNo, BOdSr, TdRmCd From Txnd" & _
'                          " JOIN Bag On TdSrNo <> 0 And TdCoCd = BCoCd And TdBYy = BYy And TdBChr = BChr And TdBNo = BNo" & _
'                          " Where BCoCd = '" & ADC("TCoCd") & "' And BOdTc = '" & wRsBag!OdTc & "' And BOdYy = '" & wRsBag!OdYy & "'" & _
'                          " And BOdChr = '" & wRsBag!OdChr & "' And BOdNo = " & CStr(wRsBag!OdNo) & " And BOdSr = " & CStr(wRsBag!OdSr) & "" & _
'                          " And TdTc = '" & adc("TTc") & "' And TdYy = '" & adc("TYy") & "' And TdChr = '" & adc("TChr") & "'" & _
'                          " And TdNo = " & CStr(adc("TNo")) & " And TdSr = " & CStr(GrdTxndA.Value(wRowA, "TdSr")) & ") As T" & _
'                          " FULL OUTER JOIN (Select OrCoCd, OrTc, OrYy, OrChr, OrNo, OrSr, OrRmCd from OrdRm Where" & _
'                          " OrCoCd = '" & ADC("TCoCd") & "' And OrTc = '" & wRsBag!OdTc & "' And OrYy = '" & wRsBag!OdYy & "'" & _
'                          " And OrChr = '" & wRsBag!OdChr & "' And OrNo = " & CStr(wRsBag!OdNo) & " And OrSr = " & CStr(wRsBag!OdSr) & ") As O" & _
'                          " On BCoCd = OrCoCd And BOdTc = OrTc And BOdYy = OrYy And BOdChr = OrChr And BOdNo = OrNo And BOdSr = OrSr" & _
'                          " And TdRmCd = OrRmCd Where (TdRmCd Is Null Or OrRmCd Is Null)") = True Then GrdTxndB.DelAllRecs
                          
                          
          '****** Sachin 3.02 24-12-07 - Current Partition
          If moCn.RecSeek("Select 1 From (Select BCoCd, BOdTc, BOdYy, BOdChr, BOdNo, BOdSr, TdRmCd From Txnd" & _
                          " JOIN Bag On TdSrNo <> 0 And TdBIdNo=BIdNo And TdCoCd = BCoCd And TdBYy = BYy And TdBChr = BChr And TdBNo = BNo And TdPrtKey=BPrtKey" & _
                          " Where BCoCd = '" & ADC("TCoCd") & "' And BOdTc = '" & wRsBag!OdTc & "' And BOdYy = '" & wRsBag!OdYy & "'" & _
                          " And BOdChr = '" & wRsBag!OdChr & "' And BOdNo = " & CStr(wRsBag!OdNo) & " And BOdSr = " & CStr(wRsBag!OdSr) & " And BPrtKey='" & ctCurrPrtn & "'" & _
                          " And TdTc = '" & ADC("TTc") & "' And TdYy = '" & ADC("TYy") & "' And TdChr = '" & ADC("TChr") & "'" & _
                          " And TdNo = " & CStr(ADC("TNo")) & " And TdSr = " & CStr(GrdTxndA.Value(wRowA, "TdSr")) & ") As T" & _
                          " LEFT OUTER JOIN (Select OrCoCd, OrTc, OrYy, OrChr, OrNo, OrSr, OrRmCd from OrdRm Where" & _
                          " OrCoCd = '" & ADC("TCoCd") & "' And OrTc = '" & wRsBag!OdTc & "' And OrYy = '" & wRsBag!OdYy & "'" & _
                          " And OrChr = '" & wRsBag!OdChr & "' And OrNo = " & CStr(wRsBag!OdNo) & " And OrSr = " & CStr(wRsBag!OdSr) & " And OrPrtKey='" & ctCurrPrtn & "') As O" & _
                          " On BCoCd = OrCoCd And BOdTc = OrTc And BOdYy = OrYy And BOdChr = OrChr And BOdNo = OrNo And BOdSr = OrSr" & _
                          " And TdRmCd = OrRmCd Where OrRmCd Is Null") = True Then
              wsBag = wsBag & vbCrLf & wRsBag!BYy & "/" & wRsBag!BChr & "/" & CStr(wRsBag!BNo) & " [Raw Material Not Copied]"
              GrdTxndB.DelAllRecs
           End If
          
          GrdTxndA.Value(wRowA, "TdBGrWt") = moCn.GetFldVal("Select isnull(BGrWt,0) from Bag where " + _
                              " BPrtKey='" + ctCurrPrtn + "' And BCoCd='" + ADC("TCoCd") + "' and " + _
                              "BYy='" + GrdTxndA.Value(wRowA, "TdBYy") + "' and " + _
                              "BChr='" + GrdTxndA.Value(wRowA, "TdBChr") + "' and " + _
                              "BNo= " + CStr(GrdTxndA.Value(wRowA, "TdBNo")))
          GrdTxndA.SaveRec (wRowA)
        End If
      End If
NextBag:
      wRsBag.MoveNext
    Loop
    GrdTxndA.EndCopy
    If wsBag <> "" Then DispMsg "Following Bags are Not Copied/Have Errors : " & wsBag, etWarning
CopyEnd:
    MWLib.EndProcess Me
End If
  
Exit Sub

ErrHndlr:
  DispMsg Err.Description, etError
End Sub


Private Function GetLsPer(ByVal pCoCd As String, ByVal pLocCd As String, ByVal pRmCtg As String, ByVal pDmCtg As String, _
                          ByVal pDmCd As String, ByVal pWrk As String, ByVal pRmWt As Double) As Double
  Dim wPrcLoc As String
  wPrcLoc = moCn.GetFldVal("Select LocPrcs from Loc where LocTyp = 'P' and LocCoCd = '" + ADC("TCoCd") + "' and LocCd = '" + pLocCd + "'")
                            
  GetLsPer = moCn.GetFldVal("Select " + _
              "IsNull((select  LmLossPer from LossMst where LmValidYn='Y' and LmCoCd = '" + pCoCd + "' and  LmRmCtg = '" + pRmCtg + "' " + _
                "and LmLocCd='" + wPrcLoc + "' and LmDmCtg='" + pDmCtg + "' and LmDmCd ='" + pDmCd + "' and LmWorkCd='" + pWrk + "' " + _
                "and Cast(LmFrRmWt As Decimal(10,3)) <= " + CStr(Round(pRmWt, 3)) + " and Cast(LmToRmWt As Decimal(10,3)) >= " + CStr(Round(pRmWt, 3)) + " " + _
                "and LmLossPer<> 0)," + _
              "IsNull((select  LmLossPer from LossMst where LmValidYn='Y' and LmCoCd='" + pCoCd + "' and  LmRmCtg = '" + pRmCtg + "' " + _
                "and LmLocCd='" + wPrcLoc + "' and LmDmCtg='" + pDmCtg + "' and LmDmCd ='" + pDmCd + "' and LmWorkCd='' " + _
                "and Cast(LmFrRmWt As Decimal(10,3)) <= " + CStr(Round(pRmWt, 3)) + " and Cast(LmToRmWt As Decimal(10,3)) >= " + CStr(Round(pRmWt, 3)) + " " + _
                "and LmLossPer<> 0)," + _
              "IsNull((select  LmLossPer from LossMst where LmValidYn='Y' and LmCoCd='" + pCoCd + "' and  LmRmCtg = '" + pRmCtg + "' " + _
                "and LmLocCd='" + wPrcLoc + "' and LmDmCtg='" + pDmCtg + "' and LmDmCd ='' and LmWorkCd='" + pWrk + "' " + _
                "and Cast(LmFrRmWt As Decimal(10,3)) <= " + CStr(Round(pRmWt, 3)) + " and Cast(LmToRmWt As Decimal(10,3)) >= " + CStr(Round(pRmWt, 3)) + " " + _
                "and LmLossPer<> 0)," + _
              "IsNull((select  LmLossPer from LossMst where LmValidYn='Y' and LmCoCd='" + pCoCd + "' and  LmRmCtg = '" + pRmCtg + "' " + _
                "and LmLocCd='" + wPrcLoc + "' and LmDmCtg='" + pDmCtg + "' and LmDmCd ='' and LmWorkCd='' " + _
                "and Cast(LmFrRmWt As Decimal(10,3)) <= " + CStr(Round(pRmWt, 3)) + " and Cast(LmToRmWt As Decimal(10,3)) >= " + CStr(Round(pRmWt, 3)) + " " + _
                "and LmLossPer<> 0)," + _
              "IsNull((select  LmLossPer from LossMst where LmValidYn='Y' and LmCoCd='" + pCoCd + "' and  LmRmCtg = '" + pRmCtg + "' " + _
                "and LmLocCd='" + wPrcLoc + "' and LmDmCtg='' and LmDmCd ='' and LmWorkCd='' " + _
                "and Cast(LmFrRmWt As Decimal(10,3)) <= " + CStr(Round(pRmWt, 3)) + " and Cast(LmToRmWt As Decimal(10,3)) >= " + CStr(Round(pRmWt, 3)) + " " + _
                "and LmLossPer<> 0)" + _
              " ,0))))) ")
End Function


' Zubin 212
Private Sub CopyFgRm()
  '*** Procedure to Copy RM details from FG bags in 'DT' entry
  Dim wMulFact As Single
  Dim wRow As Single, wTotIssQty As Single, wTotRecQty As Single
  Dim wRsCr As MwfLib.MDORowSet

  '*** (Jen 2.14 Next) ***
  Dim ws_ArrFgVch() As String
  
  '****** Sachin 3.02 24-12-07 - Current Partition
  ws_ArrFgVch = Split(moCn.GetFldVal("Select FdCoCd+ '/'+ FdTc+ '/'+ FdYy+ '/'+ FdChr+ '/'+ RTrim(LTrim(Str(FdNo)))+ '/'+ RTrim(LTrim(Str(FdSr))) " + _
                      "from Fgd B where FdCoCd= '" + ADC("TCoCd") + "' " + _
                      "and exists(Select 'x' from Param where PTyp= 'TC' and PMCd= B.FdTc and PValue= 'FB') " + _
                      "and FdBYy= '" + ADC("wFgBYy") + "' and FdBChr= '" + ADC("wFgBChr") + "' " + _
                      "and FdBNo= " + CStr(ADC("wFgBNo")) + " and FdPrtKey='" + ctCurrPrtn + "' " + _
                      "and not exists (select 'x' from Fgd M where M.FdCoCd= B.FdCoCd and M.FdRefYy= B.FdYy and " + _
                      "    M.FdRefKey= B.FdKey and M.FdPrtKey=B.FdPrtKey and exists(Select 'x' from Param where PTyp= 'TC' and PMCd= M.FdTc and PValue= 'FWB')) ") _
                , "/")
  
  If UBound(ws_ArrFgVch) <> 5 Then
    DispMsg "Cannot Copy From FG Rm", etError: Exit Sub
  End If
  
  '*** Getting The Total Iss Qty, Rec Qty ***
  '****** Sachin 3.02 24-12-07 - Current Partition
  wTotIssQty = moCn.GetFldVal("Select FdQty from Fgd where FdCoCd= '" + ws_ArrFgVch(0) + "' " + _
               "and FdTc= '" + ws_ArrFgVch(1) + "' and FdYy= '" + ws_ArrFgVch(2) + "' " + _
               "and FdChr= '" + ws_ArrFgVch(3) + "' and FdNo= " + ws_ArrFgVch(4) + " " + _
               "and FdSr= " + ws_ArrFgVch(5) + " and FdPrtKey='" + ctCurrPrtn + "' ")
  
  wTotRecQty = GrdTxndA.Value(GrdTxndA.Row, "TdBQty")
  wMulFact = IIF(wTotIssQty > 0, MWLib.Div(wTotRecQty, wTotIssQty), 1)
  
  '*** LotNo is not copied but is kept in Group By as Sir asked not to remove the same. Multiple lines for
  '    the same FrRmCd, FrRmSz and FrRmStkRt are ok (Sachin).
  '****** Sachin 3.02 24-12-07 - Current Partition
  Set wRsCr = moCn.OpenRes("Select FrRmCd, FrRmSz,FrRmSz2, FrRmSz3,FrRmStkRt, Sum(FrRmQty) as FrRmQty, Sum(FrRmWt) as FrRmWt " + _
              "from FgRm " + _
              "join Loc LFr on LFr.LocCoCd= FrCoCd and LFr.LocCd= FrFrRmLoc " + _
              "join Loc LTo on LTo.LocCoCd= FrCoCd and LTo.LocCd= FrToRmLoc " + _
              "where FrCoCd= '" + ws_ArrFgVch(0) + "' and FrTc= '" + ws_ArrFgVch(1) + "' " + _
              "   and FrYy= '" + ws_ArrFgVch(2) + "' and FrChr= '" + ws_ArrFgVch(3) + "' " + _
              "   and FrNo= " + ws_ArrFgVch(4) + " and FrSr= " + ws_ArrFgVch(5) + " and FrPrtKey='" + ctCurrPrtn + "' " + _
              "   and (Case when LFr.LocTyp= 'F' and FrFrRmDc= 'D' then 1 " + _
              "             when LTo.LocTyp= 'F' and FrToRmDc= 'D' then 1 " + _
              "       Else 0 end)= 1 " + _
              "group By FrRmCd, FrLotNo, FrRmSz,FrRmSz2, FrRmSz3, FrRmStkRt")
  '*** (Jen 2.14 Next) ***
  
  '*** (Bef 2.14 Next)
  'wTotIssQty = moCn.GetFldVal("Select FdQty from Fgd where FdCoCd= '" + ADC("TCoCd") + "' and " + _
  '             "FdBYy= '" + ADC("wFgBYy") + "' and FdBChr= '" + ADC("wFgBChr") + "' and " + _
  '             "FdBNo= " + CStr(ADC("wFgBNo")) + " ")
  '
  'wTotRecQty = GrdTxndA.Value(GrdTxndA.Row, "TdBQty")
  'wMulFact = IIF(wTotIssQty > wTotRecQty, MWLib.Div(wTotRecQty, wTotIssQty), 1) '*** (After 1.09)
  '
  '' LotNo is not copied but is kept in Group By as Sir asked not to remove the same. Multiple lines for
  '' the same FrRmCd, FrRmSz and FrRmStkRt are ok (Sachin).
  '
  'Set wRsCr = moCn.OpenRes("Select FrRmCd, FrRmSz, FrRmStkRt, Sum(FrRmQty) as FrRmQty, Sum(FrRmWt) as FrRmWt " + _
  '            "from Fgd " + _
  '            "Join FgRm On FdCoCd= FrCoCd and FdTc= FrTc and FdYy= FrYy and FdChr= FrChr and " + _
  '            "   FdNo= FrNo and FdSr= FrSr " + _
  '            " Join Loc LFr on LFr.LocCoCd= FrCoCd and LFr.LocCd= FrFrRmLoc " + _
  '            " Join Loc LTo on LTo.LocCoCd= FrCoCd and LTo.LocCd= FrToRmLoc " + _
  '            "Where FdCoCd= '" + ADC("TCoCd") + "' and FdBYy= '" + ADC("wFgBYy") + "' and " + _
  '            " FdBChr= '" + ADC("wFgBChr") + "' and FdBNo= " + CStr(ADC("wFgBNo")) + " and " + _
  '            "(Case when LFr.LocTyp= 'F' and FrFrRmDc='D' then 1 " + _
  '            " when LTo.LocTyp= 'F' and FrToRmDc='D' then 1 " + _
  '            " Else 0 end)= 1 " + _
  '            "Group By FrRmCd, FrLotNo, FrRmSz, FrRmStkRt")
  '*** (Bef 2.14 Next)
              
  GrdTxndB.StartCopy
  MWLib.BeginProcess Me, "Copying Raw Material Details ..."
  With GrdTxndB
    Do While Not (wRsCr.EOF Or wRsCr.BOF)
      .AddItem
      wRow = .Rows - 1
      .Value(wRow, "TdSrNo") = wRow
      .Value(wRow, "TdFrRmLoc") = ADC("TFrRmLoc")
      .Value(wRow, "wTdFrRmIr") = "I"
      '.Value(wRow, "TdFrRmDc") = "C"
      .Value(wRow, "TdRmCd") = wRsCr!FrRmCd
      .Value(wRow, "TdLotNo") = ""
      .Value(wRow, "TdRmSz") = wRsCr!FrRmSz
      .Value(wRow, "TdRmSz2") = wRsCr!FrRmSz2
      .Value(wRow, "TdRmSz3") = wRsCr!FrRmSz3
      .Value(wRow, "TdRmStkRt") = wRsCr!FrRmStkRt
      .Value(wRow, "TdRmQty") = wRsCr!FrRmQty * wMulFact
      .Value(wRow, "TdRmWt") = wRsCr!FrRmWt * wMulFact
      .Value(wRow, "TdToRmLoc") = ADC("TToRmLoc")
      .Value(wRow, "TdPrtKey") = ctCurrPrtn   '****** Sachin 3.02.0
      If .SaveRec(wRow) = False Then
        .DelAllRecs
        DispMsg "Cannot Copy RM From " + ADC("wFgBYy") + "/" + ADC("wFgBChr") + "/" + CStr(ADC("wFgBNo")) + " FG Bag", etError
        Exit Do
      End If
      wRsCr.MoveNext
    Loop
  End With
  MWLib.EndProcess Me
  GrdTxndB.EndCopy
  '*** End of the Copying process
  
  If GrdTxndB.Rows > 1 Then
    '*** Calculation Of TdBGrsWt ***
    '****** Sachin 3.02 24-12-07 - Current Partition
    GrdTxndA.Value(GrdTxndA.Row, "TdBGrWt") = moCn.GetFldVal("Select BGrWt from Bag Where " + _
                                              "BCoCd= '" + ADC("TCoCd") + "' and " + _
                                              "BYy='" + GrdTxndA.Value(GrdTxndA.Row, "TdBYy") + "' and " + _
                                              "BChr='" + GrdTxndA.Value(GrdTxndA.Row, "TdBChr") + "' and " + _
                                              "BNo= " + CStr(GrdTxndA.Value(GrdTxndA.Row, "TdBNo")) + " and BPrtKey='" + ctCurrPrtn + "'")
    GrdTxndA.SaveRec (GrdTxndA.Row)
  End If
  
  Set wRsCr = Nothing
End Sub
'*** Jay 3.1.3
Private Sub CopyInvRm()

  '*** Procedure to Copy RM details from FG bags in 'DT' entry
  Dim wMulFact As Single
  Dim wRow As Single, wTotIssQty As Single, wTotRecQty As Single
  Dim wRsCr As MwfLib.MDORowSet
  
  '*** Getting The Total Iss Qty, Rec Qty ***
  wTotIssQty = moCn.GetFldVal("Select IdQty from InvDsg where IdPrtKey='" & ctCurrPrtn & "' And IdCoCd= '" + ADC("wInCoCd") + "' " + _
               "and IdTc= '" + ADC("wInTc") + "' and IdYy= '" + ADC("wInYy") + "' " + _
               "and IdChr= '" + ADC("wInChr") + "' and IdNo= " + CStr(ADC("wInNo")) + " " + _
               "and IdSr= " + CStr(ADC("wInSr")) + " ")
  
  wTotRecQty = GrdTxndA.Value(GrdTxndA.Row, "TdBQty")
  wMulFact = IIF(wTotIssQty > 0, MWLib.Div(wTotRecQty, wTotIssQty), 1)
  
  '*** LotNo is not copied but is kept in Group By as Sir asked not to remove the same. Multiple lines for
  '    the same FrRmCd, FrRmSz and FrRmStkRt are ok (Sachin).
  Set wRsCr = moCn.OpenRes("Select IrRmCd, IrRmSz,IrRmSz2,IrRmSz3, IrStkRt, Sum(IrRmQty) as IrRmQty, Sum(IrRmAWt) as IrRmAWt " + _
              "from InvRm " + _
              "where IrPrtKey='" + ctCurrPrtn + "' And IrCoCd= '" + ADC("wInCoCd") + "' and IrTc= '" + ADC("wInTc") + "' and IrYy= '" + ADC("wInYy") + "' " + _
              "and IrChr= '" + ADC("wInChr") + "' and IrNo= " + CStr(ADC("wInNo")) + " and IrSr= " + CStr(ADC("wInSr")) + " " + _
              "group By IrRmCd, IrLotNo, IrRmSz, IrRmSz2, IrRmSz3, IrStkRt")
  
  GrdTxndB.StartCopy
  MWLib.BeginProcess Me, "Copying Raw Material Details ..."
  With GrdTxndB
    Do While Not (wRsCr.EOF Or wRsCr.BOF)
      .AddItem
      wRow = .Rows - 1
      .Value(wRow, "TdSrNo") = wRow
      .Value(wRow, "TdFrRmLoc") = ADC("TFrRmLoc")
      .Value(wRow, "wTdFrRmIr") = "I"
      .Value(wRow, "TdRmCd") = wRsCr!IrRmCd
      .Value(wRow, "TdLotNo") = ""
      .Value(wRow, "TdRmSz") = wRsCr!IrRmSz
      .Value(wRow, "TdRmSz2") = wRsCr!IrRmSz2
      .Value(wRow, "TdRmSz3") = wRsCr!IrRmSz3
      .Value(wRow, "TdRmStkRt") = wRsCr!IrStkRt
      .Value(wRow, "TdRmQty") = wRsCr!IrRmQty * wMulFact
      .Value(wRow, "TdRmWt") = wRsCr!IrRmAWt * wMulFact
      .Value(wRow, "TdToRmLoc") = ADC("TToRmLoc")
      .Value(wRow, "TdPrtKey") = ctCurrPrtn       '****** Sachin 3.02.0
      
      If .SaveRec(wRow) = False Then
        .DelAllRecs
        DispMsg "Cannot Copy RM From " + ADC("wInCoCd") + "/" + ADC("wInTc") + "/" + ADC("wInYy") + "/" + _
              ADC("wInChr") + "/" + CStr(ADC("wInNo")) + "/" + CStr(ADC("wInSr")) + " Invoice Serial", etError
        Exit Do
      End If
      wRsCr.MoveNext
    Loop
  End With
  MWLib.EndProcess Me
  GrdTxndB.EndCopy
  '*** End of the Copying process
  
  If GrdTxndB.Rows > 1 Then
    '*** Calculation Of TdBGrsWt ***
    GrdTxndA.Value(GrdTxndA.Row, "TdBGrWt") = moCn.GetFldVal("Select BGrWt from Bag Where " + _
                                              " BPrtKey='" + ctCurrPrtn + "' And BCoCd= '" + ADC("TCoCd") + "' and " + _
                                              "BYy='" + GrdTxndA.Value(GrdTxndA.Row, "TdBYy") + "' and " + _
                                              "BChr='" + GrdTxndA.Value(GrdTxndA.Row, "TdBChr") + "' and " + _
                                              "BNo= " + CStr(GrdTxndA.Value(GrdTxndA.Row, "TdBNo")))
    GrdTxndA.SaveRec (GrdTxndA.Row)
  End If
  Set wRsCr = Nothing
End Sub
Function ChkBagLs(ByVal pBYy As String, ByVal pBChr As String, ByVal pBNo As Single, pRmCd As String, pTdLoc As String, pTdRmWt As Single, ByVal pTdSr As Single, ByVal pTdSrNo As Single) As String

'Sachin 3.11.0 - Major Loss On Bag
Dim wMetOutWt As Single, wPpPct As Single, wStr As String, wArr() As String

ChkBagLs = ""
'indigo.164 DmTcTyp='DM'  changed as DmTcTyp=BOdDmTcTyp in Dsgmst join
wStr = moCn.GetFldVal(" Select (DmCtg + '/'+ DmCd + '/'+Cast(BQty as varChar(10))) as qDmDet From Bag " + _
                      "  Join DsgMst On DmTcTyp=BOdDmTcTyp And DmCd=BOdDmCd And DmSz='' " + _
                      " Where BCoCd='" + ADC("TCoCd") + "' And BYy='" + pBYy + "' And BChr='" + pBChr + "' And BNo=" + CStr(pBNo))
wArr = Split(wStr, "/")
wMetOutWt = moCn.GetFldVal(" Select sum(TdRmWt* (case when TdFrRmLoc= 'W' and TdFrRmDC= 'C' then -1 " + _
                           "  when TdFrRmLoc = 'W' and TdFrRmDC= 'D' then 1 " + _
                           "  when TdToRmLoc = 'W' and TdToRmDC= 'C' then -1 " + _
                           "  when TdToRmLoc = 'W' and TdToRmDC= 'D' then 1 else 0 end)) as qTdRmOutWt " + _
                           " From Txnd " + _
                           " Where TdPrtKey='" + ctCurrPrtn + "' And TdCoCd='" + ADC("TCoCd") + "' And TdBYy='" + pBYy + "' And TdBChr='" + pBChr + "' And TdBNo=" + CStr(pBNo) + _
                           "    And TdRmCd ='" + pRmCd + "'" + _
                           "    And Not (TdCoCd='" + ADC("TCoCd") + "' And TdTc='" + ADC("TTc") + "' And TdYy='" + ADC("TYy") + "' And TdChr='" + ADC("TChr") + "' And TdNo=" + CStr(ADC("TNo")) + " And TdSr=" + CStr(pTdSr) + " and TdSrNo=" + CStr(pTdSrNo) + ") ")
                           
Dim wPrcLoc As String
wPrcLoc = moCn.GetFldVal("Select LocPrcs from Loc where LocTyp = 'P' and LocCoCd = '" + ADC("TCoCd") + "' and LocCd = '" + pTdLoc + "'")

wMetOutWt = (wMetOutWt - pTdRmWt)
If Val(wArr(2)) = 0 Then Exit Function
wMetOutWt = wMetOutWt / Val(wArr(2))
wPpPct = moCn.GetFldVal(" Select IsNull((Select Top 1 PpPts From PrdPts " + _
                        " Where PpTyp='BL' And PpLoc='" + wPrcLoc + "' And PpPrdCtg In ('" + wArr(0) + "','') And PpDmTcTyp='DM' and PpDmCd In ('" + wArr(1) + "','') " + _
                        "   And PpFrRmWt<=" + CStr(wMetOutWt) + " And PpToRmWt >=" + CStr(wMetOutWt) + " And PpPts<>0" + _
                        " Order by PpPrdCtg Desc, PpDmCd Desc),0)")
'***********Bhavna added MsgBox
If wPpPct <> 0 Then
    If (pTdRmWt / Val(wArr(2))) > (wMetOutWt * wPpPct / 100) Then
        ChkBagLs = "Bag Loss More Than Permissible Limits of " + CStr(wPpPct) + "%  for " + pRmCd         '4.1.2 - Percent added to msg
        MsgBox "Bag Loss More Than Permissible Limits of " + CStr(wPpPct) + "%  for " + pRmCd + ". Click OK to Continue.", vbInformation + vbOKOnly, "Information"
    End If
End If
End Function
Private Function RmRate(ByVal mRtFld As String, ByVal mSzFld As String, ByVal mQw As String) As String
  Dim wPureGldCd As String, wPurePlCd As String, wPureSlCd As String
  Dim wGRt As String, wPRt As String, wSRt As String, wDCXARt As String
  Dim wPurePdCd As String, wLRt As String
  
  wPureGldCd = moCn.GetFldVal("Select PValue from param where PTyp= 'RMCTG' and PMCd= 'G'")
  wPurePlCd = moCn.GetFldVal("Select PValue from param where PTyp= 'RMCTG' and PMCd= 'P'")
  wPureSlCd = moCn.GetFldVal("Select PValue from param where PTyp= 'RMCTG' and PMCd= 'S'")
  wPurePdCd = moCn.GetFldVal("Select PValue from param where PTyp= 'RMCTG' and PMCd= 'L'")
  
  Dim ws_BaseCurCd As String
  ws_BaseCurCd = moCn.GetFldVal("Select HBaseCurCd from Head where HCd= '" + ctSelfCmCd + "' and " + _
                  "HCoCd= '" + gs_CoCd + "'")
  
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
  wLRt = " isnull((select " + mRtFld + " from RmRt where " + IIF(gs_Partition = ctCurrPrtn, " RrPrtKey='" + ctCurrPrtn + "' and ", "") + " RrTcTyp= 'RM' and RrCmCtg= 'C' and " + _
         "RrCd= '" + wPurePdCd + "' and " + _
         "RrFrLn <= " + mSzFld + " And RrToLn >= " + mSzFld + " and RrCmCd = '" + ctValueCmCd + "' and " + _
         "RrCmCurCd= '" + ws_BaseCurCd + "' and RrDmCtg= '' and RrLabMCd= '') , 0) "
  
  wDCXARt = " isnull((Select " + mRtFld + " from RmRt where " + IIF(gs_Partition = ctCurrPrtn, " RrPrtKey='" + ctCurrPrtn + "' and ", "") + " RrTcTyp='RM' and RrCmCtg= 'C' and " + _
         "RrCtg= RmCtg and RrSCtg= RmSCtg and RrCd= RmCd and RrCmCd = '" + ctValueCmCd + "' and " + _
         "RrCmCurCd= '" + ws_BaseCurCd + "' and " + _
         "RrFrLn <= " + mSzFld + " And RrToLn >= " + mSzFld + " and RrDmCtg= '' and RrLabMCd= '' and " + _
         mRtFld + " <> 0), 0) "
         
  RmRate = " (Case When (RmCtg= 'G' And " + mQw + "='W') Then " + wGRt + " When (RmCtg= 'P' And " + mQw + "='W') then " + wPRt + " " + _
           " when (RmCtg= 'S' And " + mQw + "='W') then " + wSRt + " when (RmCtg= 'L' And " + mQw + "='W') then " + wLRt + " " + _
           " when (RmCtg In ('G','P','S', 'L') And " + mQw + "='Q') Or  (RmCtg in ('D', 'C', 'X', 'A')) then " + wDCXARt + " else 0 end) "
           
End Function

'vk.4 Copy from scope-"" ,Copy from Bag-Y , copy from Bag alias no. -N
Private Sub OptCpyBagFr_Click(Index As Integer)
  Select Case Index
  Case Is = 0
    ADC("wCpyThruBagNo") = ""
    FraCBScope.Enabled = True
    FraCBBag.Enabled = False
 '   FraCBScope.BackColor = RGB(0, 0, 255)
    ADC("wRfIdStr") = ""
  Case Is = 1
    ADC("wCpyThruBagNo") = "Y"
    FraCBScope.Enabled = False
    FraCBBag.Enabled = True
  Case Is = 2
    ADC("wCpyThruBagNo") = "N"
    FraCBScope.Enabled = False
    FraCBBag.Enabled = True
  End Select
End Sub
'vk.4 tooltip for copy bag options
Private Sub OptCpyBagFr_GotFocus(Index As Integer)
  DispMsg OptCpyBagFr(Index).ToolTipText, etInfo
End Sub







