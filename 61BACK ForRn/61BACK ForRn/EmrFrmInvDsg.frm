VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "tabctl32.ocx"
Object = "{AB2CC32A-E652-42A2-B341-26ECE02F6386}#2.1#0"; "Mwfctlforspeed9.ocx"
Begin VB.Form EmrFrmInvDsg 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Invoice Design Entry"
   ClientHeight    =   9810
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   15210
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   9810
   ScaleWidth      =   15210
   WindowState     =   2  'Maximized
   Begin VB.Frame FraSC 
      Height          =   1530
      Left            =   960
      TabIndex        =   261
      Top             =   7680
      Visible         =   0   'False
      Width           =   5385
      Begin MwfCtl.MWCTL_MED atxt 
         Height          =   285
         Index           =   88
         Left            =   1785
         TabIndex        =   265
         ToolTipText     =   "Enter To Design Code"
         Top             =   975
         Width           =   2415
         _ExtentX        =   4260
         _ExtentY        =   503
         MaxLength       =   15
         DataType        =   4
         IdName          =   "WSCDMCDTO"
      End
      Begin MwfCtl.MWCTL_MED atxt 
         Height          =   285
         Index           =   89
         Left            =   885
         TabIndex        =   262
         ToolTipText     =   "Enter From Design Category"
         Top             =   690
         Width           =   855
         _ExtentX        =   1508
         _ExtentY        =   503
         MaxLength       =   5
         DataType        =   4
         IdName          =   "WSCDMCTGFR"
      End
      Begin MwfCtl.MWCTL_MED atxt 
         Height          =   285
         Index           =   90
         Left            =   880
         TabIndex        =   263
         ToolTipText     =   "Enter To Design Category"
         Top             =   975
         Width           =   855
         _ExtentX        =   1508
         _ExtentY        =   503
         MaxLength       =   5
         DataType        =   4
         IdName          =   "WSCDMCTGTO"
      End
      Begin MwfCtl.MWCTL_MED atxt 
         Height          =   285
         Index           =   91
         Left            =   1785
         TabIndex        =   264
         ToolTipText     =   "Enter From Design Code"
         Top             =   690
         Width           =   2415
         _ExtentX        =   4260
         _ExtentY        =   503
         MaxLength       =   15
         DataType        =   4
         IdName          =   "WSCDMCDFR"
      End
      Begin MwfCtl.MWCTL_MED atxt 
         Height          =   285
         Index           =   298
         Left            =   4245
         TabIndex        =   266
         ToolTipText     =   "Enter From Invoice Sr."
         Top             =   690
         Width           =   855
         _ExtentX        =   1508
         _ExtentY        =   503
         Alignment       =   2
         Mask            =   "###0"
         MaxLength       =   4
         DataType        =   1
         IdName          =   "WSCIDSRFR"
      End
      Begin MwfCtl.MWCTL_MED atxt 
         Height          =   285
         Index           =   299
         Left            =   4245
         TabIndex        =   267
         ToolTipText     =   "Enter To Invoice Sr."
         Top             =   975
         Width           =   855
         _ExtentX        =   1508
         _ExtentY        =   503
         Alignment       =   2
         Mask            =   "###0"
         MaxLength       =   4
         DataType        =   1
         IdName          =   "WSCIDSRTO"
      End
      Begin VB.Label ALBL 
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
         Index           =   65
         Left            =   1830
         TabIndex        =   273
         Top             =   420
         Width           =   735
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
         Height          =   285
         Index           =   63
         Left            =   880
         TabIndex        =   272
         Top             =   420
         Width           =   825
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
         Index           =   66
         Left            =   60
         TabIndex        =   271
         Top             =   975
         Width           =   585
      End
      Begin VB.Label LblMain 
         BackStyle       =   0  'Transparent
         BorderStyle     =   1  'Fixed Single
         Caption         =   "  Scope Option For Selecting Records Into Grid"
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
         TabIndex        =   270
         Top             =   90
         Width           =   5130
      End
      Begin VB.Label ALBL 
         BackStyle       =   0  'Transparent
         Caption         =   "Inv Sr"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Index           =   228
         Left            =   4245
         TabIndex        =   269
         Top             =   420
         Width           =   705
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
         Index           =   67
         Left            =   60
         TabIndex        =   268
         Top             =   690
         Width           =   585
      End
   End
   Begin MwfCtl.MWCTL_BTN1 CmdSc 
      Height          =   435
      Left            =   1920
      TabIndex        =   260
      ToolTipText     =   "Option To Select The Scope Options For The Entry"
      Top             =   9300
      Width           =   945
      _ExtentX        =   1667
      _ExtentY        =   767
      BackColor       =   -2147483633
      MaskColor       =   8454143
      Caption         =   "Scope"
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
   Begin MwfCtl.MWCTL_BTN1 cmdGetFOB 
      Height          =   285
      Left            =   9600
      TabIndex        =   259
      ToolTipText     =   "Get FOB"
      Top             =   1380
      Width           =   1275
      _ExtentX        =   2249
      _ExtentY        =   503
      BackColor       =   -2147483633
      MaskColor       =   8454143
      Caption         =   "Get FO&B"
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
   Begin MwfCtl.MWCTL_BTN1 cmdGetLotRt 
      Height          =   285
      Left            =   8340
      TabIndex        =   245
      ToolTipText     =   "Get Lot rate"
      Top             =   1380
      Width           =   1275
      _ExtentX        =   2249
      _ExtentY        =   503
      BackColor       =   -2147483633
      MaskColor       =   8454143
      Caption         =   "&Get Lot Rt"
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
   Begin MwfCtl.MWCTL_BTN1 CmdInvDel 
      Height          =   435
      Left            =   12390
      TabIndex        =   135
      ToolTipText     =   "Refresh Met Rt"
      Top             =   9300
      Width           =   945
      _ExtentX        =   1667
      _ExtentY        =   767
      BackColor       =   -2147483633
      MaskColor       =   8454143
      Caption         =   "Dele&te Invoice"
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
   Begin MwfCtl.MWCTL_BTN1 CmdRefRt 
      Height          =   435
      Left            =   11450
      TabIndex        =   134
      ToolTipText     =   "Refresh Met Rt"
      Top             =   9300
      Width           =   945
      _ExtentX        =   1667
      _ExtentY        =   767
      BackColor       =   -2147483633
      MaskColor       =   8454143
      Caption         =   "&Refresh Met Rt"
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
   Begin MwfCtl.MWCTL_BTN1 CmdDRL 
      Height          =   435
      Left            =   10500
      TabIndex        =   133
      ToolTipText     =   "Delete RM and Labour For this Sr"
      Top             =   9300
      Width           =   945
      _ExtentX        =   1667
      _ExtentY        =   767
      BackColor       =   -2147483633
      MaskColor       =   8454143
      Caption         =   "&Delete Rm/Lab"
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
   Begin MwfCtl.MWCTL_BTN1 CmdSH 
      CausesValidation=   0   'False
      Height          =   435
      Left            =   9555
      TabIndex        =   132
      ToolTipText     =   "Summary Option"
      Top             =   9300
      Width           =   945
      _ExtentX        =   1667
      _ExtentY        =   767
      BackColor       =   -2147483633
      MaskColor       =   8454143
      Caption         =   "&History"
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
   Begin MwfCtl.MWCTL_BTN1 CmdCRL 
      Height          =   435
      Left            =   5760
      TabIndex        =   128
      ToolTipText     =   "Design Copy Option"
      Top             =   9300
      Width           =   945
      _ExtentX        =   1667
      _ExtentY        =   767
      BackColor       =   -2147483633
      MaskColor       =   8454143
      Caption         =   "C&opy Rm/Lab"
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
   Begin MwfCtl.MWCTL_BTN1 CmdDD 
      CausesValidation=   0   'False
      Height          =   435
      Left            =   4810
      TabIndex        =   127
      ToolTipText     =   "Option To See Design Details"
      Top             =   9300
      Width           =   945
      _ExtentX        =   1667
      _ExtentY        =   767
      BackColor       =   -2147483633
      MaskColor       =   8454143
      Caption         =   "Dsg D&et"
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
   Begin MwfCtl.MWCTL_BTN1 CmdSumm 
      Height          =   435
      Left            =   8610
      TabIndex        =   131
      ToolTipText     =   "Summary Option"
      Top             =   9300
      Width           =   945
      _ExtentX        =   1667
      _ExtentY        =   767
      BackColor       =   -2147483633
      MaskColor       =   8454143
      Caption         =   "&Summary"
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
   Begin MwfCtl.MWCTL_BTN1 CmdDC 
      CausesValidation=   0   'False
      Height          =   435
      Left            =   3845
      TabIndex        =   126
      ToolTipText     =   "Design Copy Option"
      Top             =   9300
      Width           =   945
      _ExtentX        =   1667
      _ExtentY        =   767
      BackColor       =   -2147483633
      MaskColor       =   8454143
      Caption         =   "&Copy Dsg"
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
   Begin MwfCtl.MWCTL_BTN1 CmdMI 
      Height          =   435
      Left            =   2855
      TabIndex        =   125
      ToolTipText     =   "Option To Make Invoice"
      Top             =   9300
      Width           =   975
      _ExtentX        =   1720
      _ExtentY        =   767
      BackColor       =   -2147483633
      MaskColor       =   8454143
      Caption         =   "&Make Inv"
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
   Begin MwfCtl.MWCTL_BTN1 CmdIL 
      Height          =   435
      Left            =   6720
      TabIndex        =   129
      ToolTipText     =   "Option To See The Invoice Labour Details"
      Top             =   9300
      Width           =   945
      _ExtentX        =   1667
      _ExtentY        =   767
      BackColor       =   -2147483633
      MaskColor       =   8454143
      Caption         =   "Inv &Lab"
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
   Begin MwfCtl.MWCTL_BTN1 CmdIB 
      Height          =   435
      Left            =   7645
      TabIndex        =   130
      ToolTipText     =   "Option To See The Invoice Bag Detail"
      Top             =   9300
      Width           =   945
      _ExtentX        =   1667
      _ExtentY        =   767
      BackColor       =   -2147483633
      MaskColor       =   8454143
      Caption         =   "I&nv Bag"
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
      Left            =   12060
      TabIndex        =   12
      TabStop         =   0   'False
      Top             =   9360
      Visible         =   0   'False
      Width           =   555
      _ExtentX        =   979
      _ExtentY        =   714
   End
   Begin MwfCtl.MWCTL_ADC ADC 
      CausesValidation=   0   'False
      Height          =   555
      Left            =   60
      TabIndex        =   13
      Top             =   9300
      Width           =   15090
      _ExtentX        =   26617
      _ExtentY        =   979
      AllowAdd        =   0   'False
      AllowDelete     =   0   'False
      BtnWidth        =   885
      BtnHeight       =   435
      BtnMaskColor    =   8454143
      FontSz          =   8.25
   End
   Begin VB.Frame FRA_KEYS 
      BorderStyle     =   0  'None
      Height          =   290
      Left            =   -30
      TabIndex        =   90
      Top             =   0
      Width           =   4170
      Begin MwfCtl.MWCTL_MED atxt 
         Height          =   285
         Index           =   76
         Left            =   3840
         TabIndex        =   191
         ToolTipText     =   "Partition Key"
         Top             =   0
         Visible         =   0   'False
         Width           =   345
         _ExtentX        =   609
         _ExtentY        =   503
         Enabled         =   0   'False
         IdName          =   "WIDPRTKEY"
      End
      Begin MwfCtl.MWCTL_MED atxt 
         Height          =   285
         Index           =   1
         Left            =   1305
         TabIndex        =   0
         ToolTipText     =   "Enter Voucher Type"
         Top             =   0
         Width           =   570
         _ExtentX        =   1005
         _ExtentY        =   503
         MaxLength       =   3
         DataType        =   4
         ReCalcParent    =   "WIDCHR"
         IdName          =   "WIDTC"
      End
      Begin MwfCtl.MWCTL_MED atxt 
         Height          =   285
         Index           =   40
         Left            =   1980
         TabIndex        =   1
         ToolTipText     =   "Enter Voucher Year"
         Top             =   0
         Width           =   330
         _ExtentX        =   582
         _ExtentY        =   503
         MaxLength       =   2
         DataType        =   4
         IdName          =   "WIDYY"
      End
      Begin MwfCtl.MWCTL_MED atxt 
         Height          =   285
         Index           =   41
         Left            =   2400
         TabIndex        =   2
         ToolTipText     =   "Enter Voucher Character"
         Top             =   0
         Width           =   570
         _ExtentX        =   1005
         _ExtentY        =   503
         MaxLength       =   3
         DataType        =   4
         ReCalcOn        =   "WIDTC"
         IdName          =   "WIDCHR"
      End
      Begin MwfCtl.MWCTL_MED atxt 
         Height          =   285
         Index           =   42
         Left            =   3060
         TabIndex        =   3
         ToolTipText     =   "Enter Voucher No."
         Top             =   0
         Width           =   750
         _ExtentX        =   1323
         _ExtentY        =   503
         Alignment       =   2
         Mask            =   "#####0"
         MaxLength       =   6
         DataType        =   1
         IdName          =   "WIDNO"
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
         Left            =   2985
         TabIndex        =   96
         Top             =   0
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
         Left            =   2325
         TabIndex        =   95
         Top             =   0
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
         Index           =   38
         Left            =   1890
         TabIndex        =   94
         Top             =   0
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
         Index           =   0
         Left            =   90
         TabIndex        =   91
         Top             =   0
         Width           =   1065
      End
   End
   Begin VB.Frame FRA_NKEYS 
      BorderStyle     =   0  'None
      Enabled         =   0   'False
      Height          =   9630
      Left            =   -120
      TabIndex        =   52
      Top             =   0
      Width           =   15315
      Begin VB.Frame FraDD 
         Height          =   2610
         Left            =   1020
         TabIndex        =   18
         Top             =   6600
         Visible         =   0   'False
         Width           =   13680
         Begin VB.Frame Fra3 
            Height          =   2245
            Left            =   9930
            TabIndex        =   140
            Top             =   360
            Width           =   3745
            Begin MwfCtl.MWCTL_MED atxt 
               Height          =   285
               Index           =   35
               Left            =   2145
               TabIndex        =   141
               ToolTipText     =   "Invoice FOB Value Per Piece"
               Top             =   90
               Width           =   1365
               _ExtentX        =   2408
               _ExtentY        =   503
               Alignment       =   2
               Enabled         =   0   'False
               Mask            =   "-########0.00"
               MaxLength       =   13
               DataType        =   2
               IdName          =   "WDDIDIFOBPERPC"
            End
            Begin MwfCtl.MWCTL_MED atxt 
               Height          =   285
               Index           =   36
               Left            =   2145
               TabIndex        =   142
               ToolTipText     =   "Customs FOB Value Per Piece"
               Top             =   375
               Width           =   1365
               _ExtentX        =   2408
               _ExtentY        =   503
               Alignment       =   2
               Enabled         =   0   'False
               Mask            =   "-########0.00"
               MaxLength       =   13
               DataType        =   2
               IdName          =   "WDDIDZFOBPERPC"
            End
            Begin MwfCtl.MWCTL_MED atxt 
               Height          =   285
               Index           =   49
               Left            =   2145
               TabIndex        =   152
               ToolTipText     =   "Invoice FOB Value Per Piece without  FOB %age "
               Top             =   660
               Width           =   1365
               _ExtentX        =   2408
               _ExtentY        =   503
               Alignment       =   2
               Enabled         =   0   'False
               Mask            =   "-########0.00"
               MaxLength       =   13
               DataType        =   2
               IdName          =   "WCALCIDIFOBPERPC"
            End
            Begin MwfCtl.MWCTL_MED atxt 
               Height          =   285
               Index           =   50
               Left            =   2145
               TabIndex        =   153
               ToolTipText     =   "Customs FOB Value Per Piece without  FOB %age"
               Top             =   945
               Width           =   1365
               _ExtentX        =   2408
               _ExtentY        =   503
               Alignment       =   2
               Enabled         =   0   'False
               Mask            =   "-########0.00"
               MaxLength       =   13
               DataType        =   2
               IdName          =   "WCALCIDZFOBPERPC"
            End
            Begin MwfCtl.MWCTL_MED atxt 
               Height          =   285
               Index           =   83
               Left            =   2145
               TabIndex        =   220
               ToolTipText     =   "LME Rate From Order"
               Top             =   1230
               Width           =   315
               _ExtentX        =   556
               _ExtentY        =   503
               Enabled         =   0   'False
               DataType        =   4
               IdName          =   "WDDIDLMERTFRORD"
            End
            Begin VB.Label Label2 
               BackStyle       =   0  'Transparent
               Caption         =   "Metal Rt. Fr Ord"
               BeginProperty Font 
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
               TabIndex        =   233
               Top             =   1230
               Width           =   1545
            End
            Begin VB.Label ALBL 
               BackStyle       =   0  'Transparent
               Caption         =   "Calc Custms FOB/ Pc"
               BeginProperty Font 
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
               Left            =   90
               TabIndex        =   155
               Top             =   945
               Width           =   2175
            End
            Begin VB.Label ALBL 
               BackStyle       =   0  'Transparent
               Caption         =   "Calc Inv FOB/ Pc"
               BeginProperty Font 
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
               TabIndex        =   154
               Top             =   660
               Width           =   1935
            End
            Begin VB.Label ALBL 
               BackStyle       =   0  'Transparent
               Caption         =   "Inv FOB/ Pc"
               BeginProperty Font 
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
               TabIndex        =   144
               Top             =   90
               Width           =   1215
            End
            Begin VB.Label ALBL 
               BackStyle       =   0  'Transparent
               Caption         =   "Custms FOB/ Pc"
               BeginProperty Font 
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
               Left            =   90
               TabIndex        =   143
               Top             =   375
               Width           =   1695
            End
         End
         Begin VB.Frame Fra2 
            Height          =   2245
            Left            =   6675
            TabIndex        =   89
            Top             =   360
            Width           =   3285
            Begin MwfCtl.MWCTL_MED atxt 
               Height          =   285
               Index           =   11
               Left            =   1665
               TabIndex        =   34
               ToolTipText     =   "Invoice Value"
               Top             =   945
               Width           =   1365
               _ExtentX        =   2408
               _ExtentY        =   503
               Alignment       =   2
               Enabled         =   0   'False
               Mask            =   "########0.00"
               MaxLength       =   12
               DataType        =   2
               IdName          =   "WDDIDIVAL"
            End
            Begin MwfCtl.MWCTL_MED atxt 
               Height          =   285
               Index           =   12
               Left            =   1665
               TabIndex        =   35
               ToolTipText     =   "Customs Value"
               Top             =   1230
               Width           =   1365
               _ExtentX        =   2408
               _ExtentY        =   503
               Alignment       =   2
               Enabled         =   0   'False
               Mask            =   "########0.00"
               MaxLength       =   12
               DataType        =   2
               IdName          =   "WDDIDZVAL"
            End
            Begin MwfCtl.MWCTL_MED atxt 
               Height          =   285
               Index           =   16
               Left            =   1665
               TabIndex        =   36
               ToolTipText     =   "Enter Invoice Additional Value"
               Top             =   1515
               Width           =   1365
               _ExtentX        =   2408
               _ExtentY        =   503
               Alignment       =   2
               Mask            =   "-#######0.00"
               MaxLength       =   12
               DataType        =   2
               IdName          =   "WDDIDIADD"
            End
            Begin MwfCtl.MWCTL_MED atxt 
               Height          =   285
               Index           =   19
               Left            =   1665
               TabIndex        =   37
               ToolTipText     =   "Enter Customs Additional Value"
               Top             =   1800
               Width           =   1365
               _ExtentX        =   2408
               _ExtentY        =   503
               Alignment       =   2
               Mask            =   "-#######0.00"
               MaxLength       =   12
               DataType        =   2
               IdName          =   "WDDIDZADD"
            End
            Begin MwfCtl.MWCTL_MED atxt 
               Height          =   285
               Index           =   15
               Left            =   1665
               TabIndex        =   31
               ToolTipText     =   "Enter Fixed Price (Y / N)"
               Top             =   90
               Width           =   315
               _ExtentX        =   556
               _ExtentY        =   503
               DataType        =   4
               IdName          =   "WDDIDFIXPRC"
            End
            Begin MwfCtl.MWCTL_MED atxt 
               Height          =   285
               Index           =   17
               Left            =   1665
               TabIndex        =   33
               ToolTipText     =   "Specify whether Customs Value Equal To Invoice Value"
               Top             =   660
               Width           =   315
               _ExtentX        =   556
               _ExtentY        =   503
               Enabled         =   0   'False
               DataType        =   4
               IdName          =   "WDDIDHZVALEQINV"
            End
            Begin MwfCtl.MWCTL_MED atxt 
               Height          =   285
               Index           =   77
               Left            =   1665
               TabIndex        =   32
               ToolTipText     =   "Enter Labor Weight From Order (Y / N)"
               Top             =   375
               Width           =   315
               _ExtentX        =   556
               _ExtentY        =   503
               Enabled         =   0   'False
               DataType        =   4
               IdName          =   "WDDIDLABWTFRORD"
            End
            Begin VB.Label Label1 
               BackStyle       =   0  'Transparent
               Caption         =   "Lab Wt. Fr Ord"
               BeginProperty Font 
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
               TabIndex        =   192
               Top             =   375
               Width           =   1545
            End
            Begin VB.Label LblDdIdHZValEqInv 
               BackStyle       =   0  'Transparent
               Caption         =   "Custms Val= Inv"
               BeginProperty Font 
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
               TabIndex        =   136
               Top             =   675
               Width           =   1785
            End
            Begin VB.Label LblDdIdFixPrc 
               BackStyle       =   0  'Transparent
               Caption         =   "Fixed Prc"
               BeginProperty Font 
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
               Top             =   90
               Width           =   1275
            End
            Begin VB.Label ALBL 
               BackStyle       =   0  'Transparent
               Caption         =   "Custms Add Val"
               BeginProperty Font 
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
               Left            =   90
               TabIndex        =   113
               Top             =   1815
               Width           =   1815
            End
            Begin VB.Label ALBL 
               BackStyle       =   0  'Transparent
               Caption         =   "Inv Add Val"
               BeginProperty Font 
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
               Left            =   90
               TabIndex        =   112
               Top             =   1530
               Width           =   1245
            End
            Begin VB.Label ALBL 
               BackStyle       =   0  'Transparent
               Caption         =   "Custms Val"
               BeginProperty Font 
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
               Left            =   90
               TabIndex        =   111
               Top             =   1245
               Width           =   1335
            End
            Begin VB.Label ALBL 
               BackStyle       =   0  'Transparent
               Caption         =   "Inv Val"
               BeginProperty Font 
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
               Left            =   90
               TabIndex        =   110
               Top             =   960
               Width           =   1245
            End
         End
         Begin VB.Frame Fra1 
            Height          =   2245
            Left            =   3285
            TabIndex        =   88
            Top             =   360
            Width           =   3420
            Begin MwfCtl.MWCTL_MED atxt 
               Height          =   285
               Index           =   52
               Left            =   1755
               TabIndex        =   25
               ToolTipText     =   "Enter Value Addition Category"
               Top             =   375
               Width           =   885
               _ExtentX        =   1561
               _ExtentY        =   503
               MaxLength       =   5
               DataType        =   4
               ReCalcParent    =   "WDDIDGLDLS,WDDIDVALADDN"
               IdName          =   "WDDIDVACTG"
            End
            Begin MwfCtl.MWCTL_MED atxt 
               Height          =   285
               Index           =   45
               Left            =   1755
               TabIndex        =   29
               ToolTipText     =   "Invoice Gold As Weight"
               Top             =   1515
               Width           =   1140
               _ExtentX        =   2011
               _ExtentY        =   503
               Alignment       =   2
               Enabled         =   0   'False
               Mask            =   "#####0.000"
               MaxLength       =   10
               DataType        =   2
               IdName          =   "WDDIDIGLDASWT"
            End
            Begin MwfCtl.MWCTL_MED atxt 
               Height          =   285
               Index           =   13
               Left            =   1755
               TabIndex        =   26
               ToolTipText     =   "Enter Gold Loss"
               Top             =   660
               Width           =   885
               _ExtentX        =   1561
               _ExtentY        =   503
               Alignment       =   2
               Mask            =   "####0.000"
               MaxLength       =   9
               DataType        =   2
               ReCalcParent    =   "WDDIDCUSTMMTCHWT"
               ReCalcOn        =   "WDDIDVACTG"
               IdName          =   "WDDIDGLDLS"
            End
            Begin MwfCtl.MWCTL_MED atxt 
               Height          =   285
               Index           =   21
               Left            =   1755
               TabIndex        =   28
               ToolTipText     =   "Enter Karat"
               Top             =   1230
               Width           =   885
               _ExtentX        =   1561
               _ExtentY        =   503
               MaxLength       =   5
               DataType        =   4
               IdName          =   "WDDIDGLDAS"
            End
            Begin MwfCtl.MWCTL_MED atxt 
               Height          =   285
               Index           =   25
               Left            =   1755
               TabIndex        =   30
               ToolTipText     =   "Customs Gold As Weight"
               Top             =   1800
               Width           =   1140
               _ExtentX        =   2011
               _ExtentY        =   503
               Alignment       =   2
               Enabled         =   0   'False
               Mask            =   "#####0.000"
               MaxLength       =   10
               DataType        =   2
               IdName          =   "WDDIDZGLDASWT"
            End
            Begin MwfCtl.MWCTL_MED atxt 
               Height          =   285
               Index           =   43
               Left            =   1755
               TabIndex        =   24
               ToolTipText     =   "Enter Karat"
               Top             =   90
               Width           =   885
               _ExtentX        =   1561
               _ExtentY        =   503
               MaxLength       =   5
               DataType        =   4
               IdName          =   "WDDIDKT"
            End
            Begin MwfCtl.MWCTL_MED atxt 
               Height          =   285
               Index           =   72
               Left            =   1755
               TabIndex        =   27
               ToolTipText     =   "Enter Custom Match Weight (+ Gold Loss Wt) For Main Metal Record"
               Top             =   945
               Width           =   1395
               _ExtentX        =   2461
               _ExtentY        =   503
               Alignment       =   2
               Mask            =   "#####0.00000"
               MaxLength       =   12
               DataType        =   2
               ReCalcOn        =   "WDDIDGLDLS"
               IdName          =   "WDDIDCUSTMMTCHWT"
            End
            Begin VB.Label ALBL 
               BackStyle       =   0  'Transparent
               Caption         =   "Custom Match Wt"
               BeginProperty Font 
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
               Left            =   60
               TabIndex        =   184
               Top             =   945
               Width           =   1725
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
               Index           =   43
               Left            =   60
               TabIndex        =   124
               Top             =   90
               Width           =   1275
            End
            Begin VB.Label ALBL 
               BackStyle       =   0  'Transparent
               Caption         =   "Custms Gld As Wt"
               BeginProperty Font 
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
               TabIndex        =   123
               Top             =   1800
               Width           =   1905
            End
            Begin VB.Label ALBL 
               BackStyle       =   0  'Transparent
               Caption         =   "Gold As"
               BeginProperty Font 
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
               TabIndex        =   122
               Top             =   1230
               Width           =   1275
            End
            Begin VB.Label LblDdIdGldLs 
               BackStyle       =   0  'Transparent
               Caption         =   "Gold Loss"
               BeginProperty Font 
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
               TabIndex        =   109
               Top             =   660
               Width           =   1095
            End
            Begin VB.Label ALBL 
               BackStyle       =   0  'Transparent
               Caption         =   "Inv Gold As Wt"
               BeginProperty Font 
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
               Left            =   60
               TabIndex        =   108
               Top             =   1515
               Width           =   1575
            End
            Begin VB.Label ALBL 
               BackStyle       =   0  'Transparent
               Caption         =   "Val Addn Ctg"
               BeginProperty Font 
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
               Left            =   60
               TabIndex        =   99
               Top             =   375
               Width           =   1395
            End
         End
         Begin MwfCtl.MWCTL_MED atxt 
            Height          =   285
            Index           =   2
            Left            =   1785
            TabIndex        =   19
            ToolTipText     =   "Export Customer Code"
            Top             =   420
            Width           =   1350
            _ExtentX        =   2381
            _ExtentY        =   503
            Enabled         =   0   'False
            MaxLength       =   8
            DataType        =   4
            IdName          =   "WDDIDEXPCMCD"
         End
         Begin MwfCtl.MWCTL_MED atxt 
            Height          =   285
            Index           =   22
            Left            =   1785
            TabIndex        =   149
            ToolTipText     =   "Design Category"
            Top             =   705
            Width           =   885
            _ExtentX        =   1561
            _ExtentY        =   503
            Enabled         =   0   'False
            MaxLength       =   5
            DataType        =   4
            IdName          =   "WDDIDDMCTG"
         End
         Begin MwfCtl.MWCTL_MED atxt 
            Height          =   285
            Index           =   51
            Left            =   1785
            TabIndex        =   22
            ToolTipText     =   "Enter Value Addition "
            Top             =   1275
            Width           =   885
            _ExtentX        =   1561
            _ExtentY        =   503
            Alignment       =   2
            Mask            =   "####0.000"
            MaxLength       =   9
            DataType        =   2
            ReCalcOn        =   "WDDIDVACTG"
            IdName          =   "WDDIDVALADDN"
         End
         Begin MwfCtl.MWCTL_MED atxt 
            Height          =   285
            Index           =   64
            Left            =   1785
            TabIndex        =   177
            ToolTipText     =   "Invoice Labour As Weight"
            Top             =   1845
            Width           =   1140
            _ExtentX        =   2011
            _ExtentY        =   503
            Alignment       =   2
            Enabled         =   0   'False
            Mask            =   "#####0.000"
            MaxLength       =   10
            DataType        =   2
            IdName          =   "WDDIDILABASWT"
         End
         Begin MwfCtl.MWCTL_MED atxt 
            Height          =   285
            Index           =   70
            Left            =   1785
            TabIndex        =   23
            ToolTipText     =   "Enter Labour As"
            Top             =   1560
            Width           =   885
            _ExtentX        =   1561
            _ExtentY        =   503
            MaxLength       =   5
            DataType        =   4
            IdName          =   "WDDIDLABAS"
         End
         Begin MwfCtl.MWCTL_MED atxt 
            Height          =   285
            Index           =   71
            Left            =   1785
            TabIndex        =   178
            ToolTipText     =   "Customs Labour As Weight"
            Top             =   2130
            Width           =   1140
            _ExtentX        =   2011
            _ExtentY        =   503
            Alignment       =   2
            Enabled         =   0   'False
            Mask            =   "#####0.000"
            MaxLength       =   10
            DataType        =   2
            IdName          =   "WDDIDZLABASWT"
         End
         Begin MwfCtl.MWCTL_MED atxt 
            Height          =   285
            Index           =   23
            Left            =   1785
            TabIndex        =   182
            ToolTipText     =   "Export Order Design Colour"
            Top             =   990
            Width           =   885
            _ExtentX        =   1561
            _ExtentY        =   503
            Enabled         =   0   'False
            MaxLength       =   5
            DataType        =   4
            IdName          =   "WDDIDDMCOL"
         End
         Begin VB.Label LblMain 
            BackStyle       =   0  'Transparent
            BorderStyle     =   1  'Fixed Single
            Caption         =   "  Design Details"
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
            TabIndex        =   173
            Top             =   90
            Width           =   13710
         End
         Begin VB.Label ALBL 
            BackStyle       =   0  'Transparent
            Caption         =   "Colour"
            BeginProperty Font 
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
            TabIndex        =   183
            Top             =   990
            Width           =   1095
         End
         Begin VB.Label ALBL 
            BackStyle       =   0  'Transparent
            Caption         =   "Inv Lab As Wt"
            BeginProperty Font 
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
            TabIndex        =   181
            Top             =   1845
            Width           =   1575
         End
         Begin VB.Label ALBL 
            BackStyle       =   0  'Transparent
            Caption         =   "Lab As"
            BeginProperty Font 
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
            Left            =   60
            TabIndex        =   180
            Top             =   1560
            Width           =   1275
         End
         Begin VB.Label ALBL 
            BackStyle       =   0  'Transparent
            Caption         =   "Custms Lab As Wt"
            BeginProperty Font 
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
            TabIndex        =   179
            Top             =   2130
            Width           =   1905
         End
         Begin VB.Label LblDdIdValAddn 
            BackStyle       =   0  'Transparent
            Caption         =   "Val Addn"
            BeginProperty Font 
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
            TabIndex        =   156
            Top             =   1275
            Width           =   975
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
            Height          =   285
            Index           =   8
            Left            =   60
            TabIndex        =   137
            Top             =   705
            Width           =   1095
         End
         Begin VB.Label ALBL 
            BackStyle       =   0  'Transparent
            Caption         =   "Prd Cust"
            BeginProperty Font 
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
            Left            =   60
            TabIndex        =   107
            Top             =   420
            Width           =   1095
         End
      End
      Begin VB.Frame FraSumm 
         BorderStyle     =   0  'None
         Height          =   3465
         Left            =   150
         TabIndex        =   53
         Top             =   5850
         Width           =   15180
         Begin TabDlg.SSTab TabSumm 
            Height          =   3225
            Left            =   0
            TabIndex        =   54
            Top             =   0
            Width           =   15135
            _ExtentX        =   26696
            _ExtentY        =   5689
            _Version        =   393216
            Tabs            =   5
            Tab             =   3
            TabsPerRow      =   5
            TabHeight       =   520
            TabCaption(0)   =   "<&1> Design RM Summary"
            TabPicture(0)   =   "EmrFrmInvDsg.frx":0000
            Tab(0).ControlEnabled=   0   'False
            Tab(0).Control(0)=   "FraTabSumm(0)"
            Tab(0).ControlCount=   1
            TabCaption(1)   =   "<&2> Design Lab Summary"
            TabPicture(1)   =   "EmrFrmInvDsg.frx":001C
            Tab(1).ControlEnabled=   0   'False
            Tab(1).Control(0)=   "FraTabSumm(1)"
            Tab(1).ControlCount=   1
            TabCaption(2)   =   "<&3> Invoice Summary"
            TabPicture(2)   =   "EmrFrmInvDsg.frx":0038
            Tab(2).ControlEnabled=   0   'False
            Tab(2).Control(0)=   "FraTabSumm(2)"
            Tab(2).ControlCount=   1
            TabCaption(3)   =   "<&4> Invoice RM Summary"
            TabPicture(3)   =   "EmrFrmInvDsg.frx":0054
            Tab(3).ControlEnabled=   -1  'True
            Tab(3).Control(0)=   "FraTabSumm(3)"
            Tab(3).Control(0).Enabled=   0   'False
            Tab(3).ControlCount=   1
            TabCaption(4)   =   "<&5> Invoice Lab Summary"
            TabPicture(4)   =   "EmrFrmInvDsg.frx":0070
            Tab(4).ControlEnabled=   0   'False
            Tab(4).Control(0)=   "FraTabSumm(4)"
            Tab(4).ControlCount=   1
            Begin VB.Frame FraTabSumm 
               BorderStyle     =   0  'None
               Enabled         =   0   'False
               Height          =   2840
               Index           =   4
               Left            =   -74940
               TabIndex        =   82
               Top             =   360
               Width           =   15050
               Begin MwfCtl.MWCTL_FLX GrdSummIL 
                  Height          =   2310
                  Left            =   90
                  TabIndex        =   83
                  Top             =   150
                  Width           =   6675
                  _ExtentX        =   11774
                  _ExtentY        =   4075
                  Cols            =   7
                  colname1        =   "WILMCD"
                  heading1        =   "Main Cd"
                  datatype1       =   4
                  maxlength1      =   5
                  tooltiptext1    =   "Labour Main Code"
                  colname2        =   "WILSCD"
                  heading2        =   "Sub Cd       "
                  datatype2       =   4
                  maxlength2      =   8
                  tooltiptext2    =   "Labour Sub Code"
                  colname3        =   "WILQW"
                  heading3        =   "By Q/W"
                  datatype3       =   4
                  maxlength3      =   1
                  style3          =   2
                  tooltiptext3    =   "Labour Rate By Qty/Wt"
                  colname4        =   "WILQTY"
                  heading4        =   "Qty        "
                  datatype4       =   1
                  mask4           =   "#####0"
                  maxlength4      =   6
                  tooltiptext4    =   "Labour Quantity"
                  colname5        =   "WILIVAL"
                  heading5        =   "Inv Val             "
                  datatype5       =   2
                  mask5           =   "-########0.00"
                  maxlength5      =   13
                  tooltiptext5    =   "Invoice Labour Value"
                  colname6        =   "WILZVAL"
                  heading6        =   "Customs Val    "
                  datatype6       =   2
                  mask6           =   "-#######0.00"
                  maxlength6      =   12
                  tooltiptext6    =   "Customs Labour Value"
               End
               Begin MwfCtl.MWCTL_MED atxt 
                  Height          =   285
                  Index           =   33
                  Left            =   3600
                  TabIndex        =   85
                  ToolTipText     =   "Total Invoice Labour Value"
                  Top             =   2460
                  Width           =   1395
                  _ExtentX        =   2461
                  _ExtentY        =   503
                  Alignment       =   2
                  Enabled         =   0   'False
                  Mask            =   "########0.00"
                  MaxLength       =   12
                  DataType        =   2
                  IdName          =   "WILTOTIVAL"
               End
               Begin MwfCtl.MWCTL_MED atxt 
                  Height          =   285
                  Index           =   65
                  Left            =   1095
                  TabIndex        =   84
                  ToolTipText     =   "Enter Group Option For Order Labour Summary"
                  Top             =   2460
                  Width           =   1365
                  _ExtentX        =   2408
                  _ExtentY        =   503
                  MaxLength       =   8
                  DataType        =   4
                  IdName          =   "WILGRP"
               End
               Begin MwfCtl.MWCTL_MED atxt 
                  Height          =   285
                  Index           =   20
                  Left            =   4980
                  TabIndex        =   86
                  ToolTipText     =   "Total Customs Labour Value"
                  Top             =   2460
                  Width           =   1485
                  _ExtentX        =   2619
                  _ExtentY        =   503
                  Alignment       =   2
                  Enabled         =   0   'False
                  Mask            =   "########0.00"
                  MaxLength       =   12
                  DataType        =   2
                  IdName          =   "WILTOTZVAL"
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
                  Index           =   32
                  Left            =   90
                  TabIndex        =   120
                  Top             =   2460
                  Width           =   1005
               End
            End
            Begin VB.Frame FraTabSumm 
               BorderStyle     =   0  'None
               Enabled         =   0   'False
               Height          =   2840
               Index           =   3
               Left            =   60
               TabIndex        =   74
               Top             =   360
               Width           =   15050
               Begin MwfCtl.MWCTL_MED atxt 
                  Height          =   285
                  Index           =   27
                  Left            =   2790
                  TabIndex        =   77
                  ToolTipText     =   "Total Raw Material Quantity"
                  Top             =   2460
                  Width           =   915
                  _ExtentX        =   1614
                  _ExtentY        =   503
                  Alignment       =   2
                  Enabled         =   0   'False
                  Mask            =   "######0"
                  MaxLength       =   7
                  DataType        =   2
                  IdName          =   "WIRTOTQTY"
               End
               Begin MwfCtl.MWCTL_FLX GrdSummIR 
                  Height          =   2310
                  Left            =   90
                  TabIndex        =   75
                  Top             =   150
                  Width           =   10125
                  _ExtentX        =   17859
                  _ExtentY        =   4075
                  Cols            =   9
                  colname1        =   "WIRCTG"
                  heading1        =   "Ctg"
                  datatype1       =   4
                  maxlength1      =   2
                  tooltiptext1    =   "Raw Material Category"
                  colname2        =   "WIRSCTG"
                  heading2        =   "SubCtg"
                  datatype2       =   4
                  maxlength2      =   5
                  tooltiptext2    =   "Raw Material Sub Category"
                  colname3        =   "WIRCD"
                  heading3        =   "Rm Code                                 "
                  datatype3       =   4
                  maxlength3      =   16
                  tooltiptext3    =   "Raw Material Code"
                  colname4        =   "WIRQTY"
                  heading4        =   "Qty         "
                  datatype4       =   1
                  mask4           =   "#####0"
                  maxlength4      =   6
                  tooltiptext4    =   "Raw Material Quantity"
                  colname5        =   "WIRIWT"
                  heading5        =   "Inv Wt.           "
                  datatype5       =   2
                  mask5           =   "#####0.000"
                  maxlength5      =   10
                  tooltiptext5    =   "Invoice Raw Material Weight"
                  colname6        =   "WIRIVAL"
                  heading6        =   "Inv Val              "
                  datatype6       =   2
                  mask6           =   "########0.00"
                  maxlength6      =   12
                  tooltiptext6    =   "Invoice Raw Material Value"
                  colname7        =   "WIRZWT"
                  heading7        =   "Customs Wt "
                  datatype7       =   2
                  mask7           =   "#####0.000"
                  maxlength7      =   10
                  tooltiptext7    =   "Invoice Raw Material Weight"
                  colname8        =   "WIRZVAL"
                  heading8        =   "Customs Val      "
                  datatype8       =   2
                  mask8           =   "########0.00"
                  maxlength8      =   12
                  tooltiptext8    =   "Customs Raw Material Value"
               End
               Begin MwfCtl.MWCTL_MED atxt 
                  Height          =   285
                  Index           =   28
                  Left            =   3690
                  TabIndex        =   78
                  ToolTipText     =   "Total Raw Material Invoice Weight"
                  Top             =   2460
                  Width           =   1215
                  _ExtentX        =   2143
                  _ExtentY        =   503
                  Alignment       =   2
                  Enabled         =   0   'False
                  Mask            =   "######0.000"
                  MaxLength       =   11
                  DataType        =   2
                  IdName          =   "WIRTOTIWT"
               End
               Begin MwfCtl.MWCTL_MED atxt 
                  Height          =   285
                  Index           =   32
                  Left            =   4890
                  TabIndex        =   79
                  ToolTipText     =   "Total Raw Material Invoice Value"
                  Top             =   2460
                  Width           =   1365
                  _ExtentX        =   2408
                  _ExtentY        =   503
                  Alignment       =   2
                  Enabled         =   0   'False
                  Mask            =   "########0.000"
                  MaxLength       =   13
                  DataType        =   2
                  IdName          =   "WIRTOTIVAL"
               End
               Begin MwfCtl.MWCTL_MED atxt 
                  Height          =   285
                  Index           =   63
                  Left            =   1095
                  TabIndex        =   76
                  ToolTipText     =   "Enter Group Option For Order Raw Material Summary"
                  Top             =   2460
                  Width           =   1365
                  _ExtentX        =   2408
                  _ExtentY        =   503
                  MaxLength       =   8
                  DataType        =   4
                  IdName          =   "WIRGRP"
               End
               Begin MwfCtl.MWCTL_MED atxt 
                  Height          =   285
                  Index           =   14
                  Left            =   6240
                  TabIndex        =   80
                  ToolTipText     =   "Total Raw Material Customs Weight"
                  Top             =   2460
                  Width           =   1185
                  _ExtentX        =   2090
                  _ExtentY        =   503
                  Alignment       =   2
                  Enabled         =   0   'False
                  Mask            =   "######0.000"
                  MaxLength       =   11
                  DataType        =   2
                  IdName          =   "WIRTOTZWT"
               End
               Begin MwfCtl.MWCTL_MED atxt 
                  Height          =   285
                  Index           =   18
                  Left            =   7410
                  TabIndex        =   81
                  ToolTipText     =   "Total Raw Material Customs Value"
                  Top             =   2460
                  Width           =   1545
                  _ExtentX        =   2725
                  _ExtentY        =   503
                  Alignment       =   2
                  Enabled         =   0   'False
                  Mask            =   "########0.000"
                  MaxLength       =   13
                  DataType        =   2
                  IdName          =   "WIRTOTZVAL"
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
                  Index           =   31
                  Left            =   90
                  TabIndex        =   119
                  Top             =   2460
                  Width           =   1005
               End
            End
            Begin VB.Frame FraTabSumm 
               BorderStyle     =   0  'None
               Enabled         =   0   'False
               Height          =   2840
               Index           =   1
               Left            =   -74940
               TabIndex        =   63
               Top             =   360
               Width           =   15050
               Begin MwfCtl.MWCTL_FLX GrdSummDL 
                  Height          =   2310
                  Left            =   90
                  TabIndex        =   64
                  ToolTipText     =   "Design Labour Summary"
                  Top             =   150
                  Width           =   7275
                  _ExtentX        =   12832
                  _ExtentY        =   4075
                  Cols            =   7
                  colname1        =   "WDLMCD"
                  heading1        =   "Main Cd"
                  datatype1       =   4
                  maxlength1      =   5
                  tooltiptext1    =   "Labour Main Code"
                  colname2        =   "WDLSCD"
                  heading2        =   "Sub Cd       "
                  datatype2       =   4
                  maxlength2      =   8
                  tooltiptext2    =   "Labour Sub Code"
                  colname3        =   "WDLQW"
                  heading3        =   "By Q/W"
                  datatype3       =   4
                  maxlength3      =   1
                  style3          =   2
                  tooltiptext3    =   "Labour Rt By Qty/Wt"
                  colname4        =   "WDLQTY"
                  heading4        =   "Qty       "
                  datatype4       =   1
                  mask4           =   "#####0"
                  maxlength4      =   6
                  tooltiptext4    =   "Labour Quantity"
                  colname5        =   "WDLIVAL"
                  heading5        =   "Inv Val              "
                  datatype5       =   2
                  mask5           =   "-########0.00"
                  maxlength5      =   13
                  tooltiptext5    =   "Invoice Labour Value"
                  colname6        =   "WDLZVAL"
                  heading6        =   "Customs Val     "
                  datatype6       =   2
                  mask6           =   "-########0.00"
                  maxlength6      =   13
                  tooltiptext6    =   "Customs Labour Value"
               End
               Begin MwfCtl.MWCTL_MED atxt 
                  Height          =   285
                  Index           =   66
                  Left            =   3660
                  TabIndex        =   66
                  ToolTipText     =   "Total Invoice Labour Value"
                  Top             =   2460
                  Width           =   1335
                  _ExtentX        =   2355
                  _ExtentY        =   503
                  Alignment       =   2
                  Enabled         =   0   'False
                  Mask            =   "#########0.00"
                  MaxLength       =   13
                  DataType        =   2
                  IdName          =   "WDLTOTIVAL"
               End
               Begin MwfCtl.MWCTL_MED atxt 
                  Height          =   285
                  Index           =   67
                  Left            =   1095
                  TabIndex        =   65
                  ToolTipText     =   "Enter Group Option For Design Labour Summary"
                  Top             =   2490
                  Width           =   1365
                  _ExtentX        =   2408
                  _ExtentY        =   503
                  MaxLength       =   8
                  DataType        =   4
                  IdName          =   "WDLGRP"
               End
               Begin MwfCtl.MWCTL_MED atxt 
                  Height          =   285
                  Index           =   9
                  Left            =   4980
                  TabIndex        =   67
                  ToolTipText     =   "Total Customs Labour Value"
                  Top             =   2460
                  Width           =   1455
                  _ExtentX        =   2566
                  _ExtentY        =   503
                  Alignment       =   2
                  Enabled         =   0   'False
                  Mask            =   "-########0.00"
                  MaxLength       =   13
                  DataType        =   2
                  IdName          =   "WDLTOTZVAL"
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
                  Index           =   33
                  Left            =   90
                  TabIndex        =   117
                  Top             =   2460
                  Width           =   1005
               End
            End
            Begin VB.Frame FraTabSumm 
               BorderStyle     =   0  'None
               Enabled         =   0   'False
               Height          =   2840
               Index           =   0
               Left            =   -74940
               TabIndex        =   55
               Top             =   360
               Width           =   15050
               Begin MwfCtl.MWCTL_MED atxt 
                  Height          =   285
                  Index           =   26
                  Left            =   3060
                  TabIndex        =   58
                  ToolTipText     =   "Total Raw Material Quantity"
                  Top             =   2460
                  Width           =   855
                  _ExtentX        =   1508
                  _ExtentY        =   503
                  Alignment       =   2
                  Enabled         =   0   'False
                  Mask            =   "#####0"
                  MaxLength       =   6
                  DataType        =   2
                  IdName          =   "WDRTOTQTY"
               End
               Begin MwfCtl.MWCTL_FLX GrdSummDR 
                  Height          =   2310
                  Left            =   90
                  TabIndex        =   56
                  ToolTipText     =   "Design Raw Material Summary"
                  Top             =   150
                  Width           =   10185
                  _ExtentX        =   17965
                  _ExtentY        =   4075
                  Cols            =   9
                  colname1        =   "WDRCTG"
                  heading1        =   "Ctg"
                  datatype1       =   4
                  maxlength1      =   2
                  tooltiptext1    =   "Raw Material Category"
                  colname2        =   "WDRSCTG"
                  heading2        =   "Sub Ctg"
                  datatype2       =   4
                  maxlength2      =   5
                  tooltiptext2    =   "Raw Material Sub Category"
                  colname3        =   "WDRCD"
                  heading3        =   "Rm Code                                 "
                  datatype3       =   4
                  maxlength3      =   16
                  tooltiptext3    =   "Raw Material Code"
                  colname4        =   "WDRQTY"
                  heading4        =   "Qty        "
                  datatype4       =   1
                  mask4           =   "#####0"
                  maxlength4      =   6
                  tooltiptext4    =   "Raw Material Quantity"
                  colname5        =   "WDRIWT"
                  heading5        =   "Inv Wt.          "
                  datatype5       =   2
                  mask5           =   "#####0.000"
                  maxlength5      =   10
                  tooltiptext5    =   "Invoice Raw Material Weight"
                  colname6        =   "WDRIVAL"
                  heading6        =   "Inv Val              "
                  datatype6       =   2
                  mask6           =   "########0.00"
                  maxlength6      =   12
                  tooltiptext6    =   "Invoice Raw Material Value"
                  colname7        =   "WDRZWT"
                  heading7        =   "Customs Wt    "
                  datatype7       =   2
                  mask7           =   "#####0.000"
                  maxlength7      =   10
                  tooltiptext7    =   "Customs Raw Material Weight"
                  colname8        =   "WDRZVAL"
                  heading8        =   "Customs Val     "
                  datatype8       =   2
                  mask8           =   "########0.00"
                  maxlength8      =   12
                  tooltiptext8    =   "Customs Raw Material Value"
               End
               Begin MwfCtl.MWCTL_MED atxt 
                  Height          =   285
                  Index           =   34
                  Left            =   3900
                  TabIndex        =   59
                  ToolTipText     =   "Total Invoice Raw Material Weight"
                  Top             =   2460
                  Width           =   1125
                  _ExtentX        =   1984
                  _ExtentY        =   503
                  Alignment       =   2
                  Enabled         =   0   'False
                  Mask            =   "######0.000"
                  MaxLength       =   11
                  DataType        =   2
                  IdName          =   "WDRTOTIWT"
               End
               Begin MwfCtl.MWCTL_MED atxt 
                  Height          =   285
                  Index           =   29
                  Left            =   5010
                  TabIndex        =   60
                  ToolTipText     =   "Total Invoice Raw Material Value"
                  Top             =   2460
                  Width           =   1335
                  _ExtentX        =   2355
                  _ExtentY        =   503
                  Alignment       =   2
                  Enabled         =   0   'False
                  Mask            =   "########0.000"
                  MaxLength       =   13
                  DataType        =   2
                  IdName          =   "WDRTOTIVAL"
               End
               Begin MwfCtl.MWCTL_MED atxt 
                  Height          =   285
                  Index           =   56
                  Left            =   1095
                  TabIndex        =   57
                  ToolTipText     =   "Enter Group Option For Design Raw Material Summary"
                  Top             =   2460
                  Width           =   1365
                  _ExtentX        =   2408
                  _ExtentY        =   503
                  MaxLength       =   8
                  DataType        =   4
                  IdName          =   "WDRGRP"
               End
               Begin MwfCtl.MWCTL_MED atxt 
                  Height          =   285
                  Index           =   7
                  Left            =   6330
                  TabIndex        =   61
                  ToolTipText     =   "Total Customs Raw Material Weight"
                  Top             =   2460
                  Width           =   1335
                  _ExtentX        =   2355
                  _ExtentY        =   503
                  Alignment       =   2
                  Enabled         =   0   'False
                  Mask            =   "######0.000"
                  MaxLength       =   11
                  DataType        =   2
                  IdName          =   "WDRTOTZWT"
               End
               Begin MwfCtl.MWCTL_MED atxt 
                  Height          =   285
                  Index           =   8
                  Left            =   7650
                  TabIndex        =   62
                  ToolTipText     =   "Total Customs Raw Material Value"
                  Top             =   2460
                  Width           =   1455
                  _ExtentX        =   2566
                  _ExtentY        =   503
                  Alignment       =   2
                  Enabled         =   0   'False
                  Mask            =   "########0.000"
                  MaxLength       =   13
                  DataType        =   2
                  IdName          =   "WDRTOTZVAL"
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
                  Index           =   1
                  Left            =   90
                  TabIndex        =   87
                  Top             =   2460
                  Width           =   1005
               End
            End
            Begin VB.Frame FraTabSumm 
               BorderStyle     =   0  'None
               Enabled         =   0   'False
               Height          =   2840
               Index           =   2
               Left            =   -74940
               TabIndex        =   68
               Top             =   360
               Width           =   15050
               Begin MwfCtl.MWCTL_MED atxt 
                  Height          =   285
                  Index           =   54
                  Left            =   3735
                  TabIndex        =   71
                  ToolTipText     =   "Total Order Quantity"
                  Top             =   2460
                  Width           =   1005
                  _ExtentX        =   1773
                  _ExtentY        =   503
                  Alignment       =   2
                  Enabled         =   0   'False
                  Mask            =   "######0.0"
                  MaxLength       =   9
                  DataType        =   2
                  IdName          =   "WIDTOTQTY"
               End
               Begin MwfCtl.MWCTL_FLX GrdSummID 
                  Height          =   2310
                  Left            =   90
                  TabIndex        =   69
                  ToolTipText     =   "Order Design Summary"
                  Top             =   150
                  Width           =   10875
                  _ExtentX        =   19182
                  _ExtentY        =   4075
                  Cols            =   9
                  colname1        =   "WIDSR"
                  heading1        =   "Sr.    "
                  mask1           =   "###0"
                  maxlength1      =   4
                  tooltiptext1    =   "Design Serial No"
                  colname2        =   "WIDDMCTG"
                  heading2        =   "Dsg Ctg "
                  datatype2       =   4
                  maxlength2      =   5
                  tooltiptext2    =   "Design Category"
                  colname3        =   "WIDDMCD"
                  heading3        =   "Design Cd                        "
                  datatype3       =   4
                  maxlength3      =   15
                  tooltiptext3    =   "Design Code"
                  colname4        =   "WIDQTY"
                  heading4        =   "Qty             "
                  datatype4       =   2
                  mask4           =   "#####0.0"
                  maxlength4      =   8
                  tooltiptext4    =   "Order Quantity"
                  colname5        =   "WIDIPRC"
                  heading5        =   "Inv Price          "
                  datatype5       =   2
                  mask5           =   "########0.00"
                  maxlength5      =   12
                  tooltiptext5    =   "Invoice Price"
                  colname6        =   "WIDIFOB"
                  heading6        =   "Inv FOB               "
                  datatype6       =   2
                  mask6           =   "#########0.00"
                  maxlength6      =   13
                  tooltiptext6    =   "Invoice FOB"
                  colname7        =   "WIDZPRC"
                  heading7        =   "Customs Price     "
                  datatype7       =   2
                  mask7           =   "#########0.00"
                  maxlength7      =   13
                  tooltiptext7    =   "Customs Price"
                  colname8        =   "WIDZFOB"
                  heading8        =   "Customs FOB   "
                  datatype8       =   2
                  mask8           =   "########0.00"
                  maxlength8      =   12
                  tooltiptext8    =   "Customs FOB"
               End
               Begin MwfCtl.MWCTL_MED atxt 
                  Height          =   285
                  Index           =   31
                  Left            =   6030
                  TabIndex        =   72
                  ToolTipText     =   "Total Invoice FOB"
                  Top             =   2460
                  Width           =   1575
                  _ExtentX        =   2778
                  _ExtentY        =   503
                  Alignment       =   2
                  Enabled         =   0   'False
                  Mask            =   "#######0.00"
                  MaxLength       =   11
                  DataType        =   2
                  IdName          =   "WIDTOTIFOB"
               End
               Begin MwfCtl.MWCTL_MED atxt 
                  Height          =   285
                  Index           =   62
                  Left            =   1095
                  TabIndex        =   70
                  ToolTipText     =   "Enter Group Option For Order Design Summary"
                  Top             =   2460
                  Width           =   1365
                  _ExtentX        =   2408
                  _ExtentY        =   503
                  MaxLength       =   8
                  DataType        =   4
                  IdName          =   "WIDGRP"
               End
               Begin MwfCtl.MWCTL_MED atxt 
                  Height          =   285
                  Index           =   10
                  Left            =   9120
                  TabIndex        =   73
                  ToolTipText     =   "Total Customs FOB"
                  Top             =   2460
                  Width           =   1455
                  _ExtentX        =   2566
                  _ExtentY        =   503
                  Alignment       =   2
                  Enabled         =   0   'False
                  Mask            =   "#######0.00"
                  MaxLength       =   11
                  DataType        =   2
                  IdName          =   "WIDTOTZFOB"
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
                  Index           =   23
                  Left            =   90
                  TabIndex        =   118
                  Top             =   2460
                  Width           =   1005
               End
            End
         End
      End
      Begin VB.Frame FraNKeyAll 
         BorderStyle     =   0  'None
         Height          =   9315
         Left            =   120
         TabIndex        =   92
         Top             =   0
         Width           =   15150
         Begin MwfCtl.MWCTL_FLX GrdID 
            Height          =   4312
            Left            =   30
            TabIndex        =   10
            Top             =   1710
            Width           =   15105
            _ExtentX        =   26644
            _ExtentY        =   7594
            Cols            =   43
            colname1        =   "IDSR"
            heading1        =   "Sr.    "
            datafld1        =   "IdSr"
            datatype1       =   1
            mask1           =   "###0"
            maxlength1      =   4
            style1          =   2
            tooltiptext1    =   "Enter Invoice Design Sr No."
            colname2        =   "IDEXPODTC"
            heading2        =   "OdTc"
            datafld2        =   "IdExpOdTc"
            datatype2       =   4
            recalcparent2   =   "GrdID(IDEXPCMCD),GrdID(IDDMCD),GrdID(IDSFX),GrdID(IDDMSZ),GrdID(IDIVAL),GrdID(IDZVAL)"
            maxlength2      =   3
            tooltiptext2    =   "Enter Export Order Tc"
            colname3        =   "IDEXPODYY"
            heading3        =   "OdYy"
            datafld3        =   "IdExpOdYy"
            datatype3       =   4
            recalcparent3   =   "GrdID(IDEXPCMCD),GrdID(IDDMCD),GrdID(IDSFX),GrdID(IDDMSZ),GrdID(IDIVAL),GrdID(IDZVAL)"
            maxlength3      =   2
            tooltiptext3    =   "Enter Export Order Year"
            colname4        =   "IDEXPODCHR"
            heading4        =   "OdChr"
            datafld4        =   "IdExpOdChr"
            datatype4       =   4
            recalcparent4   =   "GrdID(IDEXPCMCD),GrdID(IDDMCD),GrdID(IDSFX),GrdID(IDDMSZ),GrdID(IDIVAL),GrdID(IDZVAL)"
            maxlength4      =   3
            tooltiptext4    =   "Enter Export Order Character"
            colname5        =   "IDEXPODNO"
            heading5        =   "OdNo   "
            datafld5        =   "IdExpOdNo"
            datatype5       =   1
            mask5           =   "#####0"
            recalcparent5   =   "GrdID(IDEXPCMCD),GrdID(IDDMCD),GrdID(IDSFX),GrdID(IDDMSZ),GrdID(IDIVAL),GrdID(IDZVAL)"
            maxlength5      =   6
            tooltiptext5    =   "Enter Export Order Number"
            colname6        =   "IDEXPODSR"
            heading6        =   "OdSr"
            datafld6        =   "IdExpOdSr"
            datatype6       =   1
            mask6           =   "###0"
            recalcparent6   =   "GrdID(IDEXPCMCD),GrdID(IDDMCD),GrdID(IDSFX),GrdID(IDDMSZ),GrdID(IDIVAL),GrdID(IDZVAL)"
            maxlength6      =   4
            tooltiptext6    =   "Enter Export Order Serial Number"
            colname7        =   "IDDMCD"
            heading7        =   "Design Cd                       "
            datafld7        =   "IdDmCd"
            datatype7       =   4
            recalcon7       =   "GrdID(IDEXPODTC),GrdID(IDEXPODYY),GrdID(IDEXPODCHR),GrdID(IDEXPODNO),GrdID(IDEXPODSR)"
            maxlength7      =   15
            tooltiptext7    =   "Enter Design Code"
            colname8        =   "IDSFX"
            heading8        =   "Suffix                 "
            datafld8        =   "IdSfx"
            datatype8       =   4
            recalcon8       =   "GrdID(IDEXPODTC),GrdID(IDEXPODYY),GrdID(IDEXPODCHR),GrdID(IDEXPODNO),GrdID(IDEXPODSR)"
            maxlength8      =   15
            style8          =   2
            tooltiptext8    =   "Enter Suffix"
            colname9        =   "IDDMSZ"
            heading9        =   "Size      "
            datafld9        =   "IdDmSz"
            datatype9       =   4
            recalcon9       =   "GrdID(IDEXPODTC),GrdID(IDEXPODYY),GrdID(IDEXPODCHR),GrdID(IDEXPODNO),GrdID(IDEXPODSR)"
            maxlength9      =   5
            style9          =   2
            tooltiptext9    =   "Enter Design Size"
            colname10       =   "IDQTY"
            heading10       =   "Qty         "
            datafld10       =   "IdQty"
            datatype10      =   2
            mask10          =   "####0.0"
            recalcon10      =   "GrdIB(IFQTY)"
            recalcparent10  =   "GrdID(IDIVAL),GrdID(IDZVAL)"
            maxlength10     =   7
            style10         =   2
            tooltiptext10   =   "Design Quantity"
            colname11       =   "IDAVAL"
            datafld11       =   "IdAVal"
            datatype11      =   2
            mask11          =   "########0.00"
            recalcon11      =   "GrdIR(IRRMAVAL),GrdIR(IRSETAVAL),GrdIL(ILAVAL)"
            maxlength11     =   12
            style11         =   0
            tooltiptext11   =   "Actual Value"
            colname12       =   "IDIVAL"
            datafld12       =   "IdIVal"
            datatype12      =   2
            mask12          =   "########0.00"
            recalcon12      =   $"EmrFrmInvDsg.frx":008C
            recalcparent12  =   "GrdID(IDIFOB),GrdID(IDZVAL),GrdID(IDIADD)"
            maxlength12     =   12
            style12         =   0
            tooltiptext12   =   "Invoice Value"
            colname13       =   "IDZVAL"
            datafld13       =   "IdZVal"
            datatype13      =   2
            mask13          =   "########0.00"
            recalcon13      =   $"EmrFrmInvDsg.frx":0131
            recalcparent13  =   "GrdID(IDZFOB),GrdID(IDZADD)"
            maxlength13     =   12
            style13         =   0
            tooltiptext13   =   "Customs Value"
            colname14       =   "IDIADD"
            datafld14       =   "IdIAdd"
            datatype14      =   2
            mask14          =   "-#######0.00"
            recalcon14      =   "GrdID(IDIVAL)"
            recalcparent14  =   "GrdID(IDIFOB)"
            maxlength14     =   12
            style14         =   0
            tooltiptext14   =   "Additional Invoice Value"
            colname15       =   "IDZADD"
            datafld15       =   "IdZAdd"
            datatype15      =   2
            mask15          =   "-#######0.00"
            recalcon15      =   "GrdID(IDZVAL)"
            recalcparent15  =   "GrdID(IDZFOB)"
            maxlength15     =   12
            style15         =   0
            tooltiptext15   =   "Additional Customs Value"
            colname16       =   "IDIFOB"
            heading16       =   "Inv Fob              "
            datafld16       =   "IdIFob"
            datatype16      =   2
            mask16          =   "-########0.00"
            recalcon16      =   "GrdID(IDIVAL),GrdID(IDIADD)"
            recalcparent16  =   "WTOTIDIFOB,GrdIL(ILZRT)"
            maxlength16     =   13
            style16         =   2
            tooltiptext16   =   "Invoice FOB"
            colname17       =   "IDZFOB"
            heading17       =   "Customs Fob   "
            datafld17       =   "IdZFob"
            datatype17      =   2
            mask17          =   "-########0.00"
            recalcon17      =   "GrdID(IDZADD),GrdID(IDZVAL)"
            recalcparent17  =   "WTOTIDZFOB,GrdIL(ILZRT)"
            maxlength17     =   13
            style17         =   2
            tooltiptext17   =   "Customs FOB"
            colname18       =   "IDGLDLS"
            datafld18       =   "IdGldLs"
            datatype18      =   2
            mask18          =   "###0.000"
            recalcon18      =   "GrdID(IDVACTG)"
            maxlength18     =   8
            style18         =   0
            tooltiptext18   =   "Gold Loss"
            colname19       =   "IDVALADDN"
            datafld19       =   "IdValAddn"
            datatype19      =   2
            mask19          =   "###0.000"
            recalcon19      =   "GrdID(IDVACTG)"
            maxlength19     =   8
            tooltiptext19   =   "Value Addition "
            colname20       =   "IDKT"
            datafld20       =   "IdKt"
            datatype20      =   4
            maxlength20     =   5
            style20         =   0
            tooltiptext20   =   "Enter Karat"
            colname21       =   "IDVACTG"
            datafld21       =   "IdVaCtg"
            datatype21      =   4
            recalcparent21  =   "GrdID(IDGLDLS),GrdID(IDVALADDN)"
            maxlength21     =   5
            style21         =   0
            tooltiptext21   =   "Enter Value Addition Category"
            colname22       =   "IDFIXPRC"
            datafld22       =   "IdFixPrc"
            datatype22      =   4
            recalcparent22  =   "GrdID(IDIVAL),GrdID(IDZVAL)"
            maxlength22     =   1
            style22         =   0
            tooltiptext22   =   "Enter Fixed Prc (Y/N)"
            colname23       =   "IDHZVALEQINV"
            datafld23       =   "IdHZValEqInv"
            datatype23      =   4
            recalcparent23  =   "GrdID(IDZVAL)"
            maxlength23     =   1
            style23         =   0
            tooltiptext23   =   "Specify whether Customs Value Equal To Invoice Value"
            colname24       =   "IDGLDAS"
            datafld24       =   "IdGldAs"
            datatype24      =   4
            recalcparent24  =   "GrdID(IDAGLDASWT),GrdID(IDIGLDASWT),GrdID(IDZGLDASWT)"
            maxlength24     =   5
            style24         =   0
            tooltiptext24   =   "Enter Gold As "
            colname25       =   "IDAGLDASWT"
            datafld25       =   "IdAGldAsWt"
            datatype25      =   2
            mask25          =   "####0.000"
            recalcon25      =   "GrdID(IDGLDAS),GrdIR(IRRMCD),GrdIR(IRRMAWT),GrdIR(IRMAINMET)"
            maxlength25     =   9
            style25         =   0
            tooltiptext25   =   "Enter Actual Gold As Weight"
            colname26       =   "IDIGLDASWT"
            datafld26       =   "IdiGldAsWt"
            datatype26      =   2
            mask26          =   "####0.000"
            recalcon26      =   "GrdID(IDGLDAS),GrdIR(IRRMCD),GrdIR(IRRMIWT),GrdIR(IRMAINMET)"
            maxlength26     =   9
            style26         =   0
            tooltiptext26   =   "Enter Invoice Gold As Weight"
            colname27       =   "IDZGLDASWT"
            datafld27       =   "IdZGldAsWt"
            datatype27      =   2
            mask27          =   "####0.000"
            recalcon27      =   "GrdID(IDGLDAS),GrdIR(IRRMCD),GrdIR(IRRMZWT),GrdIR(IRMAINMET)"
            maxlength27     =   9
            style27         =   0
            tooltiptext27   =   "Enter Customs Gold As Weight"
            colname28       =   "IDEXPCMCD"
            datafld28       =   "IdExpCmCd"
            datatype28      =   4
            recalcon28      =   "GrdID(IDEXPODTC),GrdID(IDEXPODYY),GrdID(IDEXPODCHR),GrdID(IDEXPODNO),GrdID(IDEXPODSR)"
            maxlength28     =   8
            style28         =   0
            tooltiptext28   =   "Export Customer Code"
            colname29       =   "IDLABAS"
            datafld29       =   "IdLabAs"
            datatype29      =   4
            recalcparent29  =   "GrdID(IDALABASWT),GrdID(IDILABASWT),GrdID(IDZLABASWT)"
            maxlength29     =   5
            style29         =   0
            colname30       =   "IDALABASWT"
            datafld30       =   "IdALabAsWt"
            datatype30      =   2
            mask30          =   "####0.000"
            recalcon30      =   "GrdID(IDLABAS),GrdIR(IRRMCD),GrdIR(IRRMAWT)"
            maxlength30     =   9
            style30         =   0
            tooltiptext30   =   "Enter Actual Lab As Weight"
            colname31       =   "IDILABASWT"
            datafld31       =   "IdILabAsWt"
            datatype31      =   2
            mask31          =   "####0.000"
            recalcon31      =   "GrdID(IDLABAS),GrdIR(IRRMCD),GrdIR(IRRMIWT)"
            maxlength31     =   9
            style31         =   0
            tooltiptext31   =   "Enter Invoice Labour As Weight"
            colname32       =   "IDZLABASWT"
            datafld32       =   "IdZLabAsWt"
            datatype32      =   2
            mask32          =   "####0.000"
            recalcon32      =   "GrdID(IDLABAS),GrdIR(IRRMCD),GrdIR(IRRMZWT)"
            maxlength32     =   9
            style32         =   0
            tooltiptext32   =   "Enter Custom Labour As Weight"
            colname33       =   "IDCUSTMMTCHWT"
            datafld33       =   "IdCustmMtchWt"
            datatype33      =   2
            mask33          =   "#####0.00000"
            recalcon33      =   "GrdIR(IRRMZWT),GrdIR(IRMAINMET)"
            maxlength33     =   12
            style33         =   0
            tooltiptext33   =   "Enter Custom Match Weight For Main Metal Record"
            colname34       =   "IDTRAYNO"
            heading34       =   "Tray No"
            datafld34       =   "IdTrayNo"
            datatype34      =   1
            mask34          =   "###0"
            maxlength34     =   4
            colname35       =   "IDPRTKEY"
            heading35       =   "Prtn "
            datafld35       =   "IdPrtKey"
            maxlength35     =   1
            style35         =   0
            colname36       =   "IDLABWTFRORD"
            datafld36       =   "IdLabWtFrOrd"
            maxlength36     =   1
            style36         =   0
            colname37       =   "WIDADIAWT"
            datatype37      =   2
            mask37          =   "###0.000"
            recalcon37      =   "GrdIR(IRRMAWT)"
            maxlength37     =   8
            style37         =   0
            colname38       =   "WIDIDIAWT"
            datatype38      =   2
            mask38          =   "###0.000"
            recalcon38      =   "GrdIR(IRRMIWT)"
            maxlength38     =   8
            style38         =   0
            colname39       =   "WIDZDIAWT"
            datatype39      =   2
            mask39          =   "###0.000"
            recalcon39      =   "GrdIR(IRRMZWT)"
            maxlength39     =   8
            style39         =   0
            colname40       =   "WIDACSWT"
            datatype40      =   2
            mask40          =   "###0.000"
            recalcon40      =   "GrdIR(IRRMAWT)"
            maxlength40     =   8
            style40         =   0
            colname41       =   "WIDICSWT"
            datatype41      =   2
            mask41          =   "###0.000"
            recalcon41      =   "GrdIR(IRRMIWT)"
            maxlength41     =   8
            style41         =   0
            colname42       =   "WIDZCSWT"
            datatype42      =   2
            mask42          =   "###0.000"
            recalcon42      =   "GrdIR(IRRMZWT)"
            maxlength42     =   8
            style42         =   0
         End
         Begin VB.Frame FraPic3D 
            BorderStyle     =   0  'None
            Height          =   1635
            Left            =   13395
            TabIndex        =   105
            Top             =   -20
            Width           =   1605
            Begin VB.Image Pic3D 
               Height          =   1515
               Left            =   15
               ToolTipText     =   "3D - Image"
               Top             =   30
               Width           =   1515
            End
         End
         Begin VB.Frame FraPicLD 
            BorderStyle     =   0  'None
            Height          =   1635
            Left            =   11610
            TabIndex        =   104
            Top             =   -15
            Width           =   1605
            Begin VB.Image PicLD 
               Height          =   1515
               Left            =   15
               ToolTipText     =   "Hand Sketch Image"
               Top             =   30
               Width           =   1515
            End
         End
         Begin MwfCtl.MWCTL_MED atxt 
            Height          =   285
            Index           =   30
            Left            =   1245
            TabIndex        =   6
            ToolTipText     =   "Invoice Customer Code"
            Top             =   285
            Width           =   1365
            _ExtentX        =   2408
            _ExtentY        =   503
            Enabled         =   0   'False
            MaxLength       =   8
            DataType        =   4
            ReCalcParent    =   "WCHNCUSTYN"
            IdName          =   "WIDCMCD"
         End
         Begin MwfCtl.MWCTL_MED atxt 
            Height          =   285
            Index           =   0
            Left            =   4785
            TabIndex        =   4
            ToolTipText     =   "Invoice Date"
            Top             =   0
            Width           =   885
            _ExtentX        =   1561
            _ExtentY        =   503
            Enabled         =   0   'False
            MaxLength       =   8
            DataType        =   3
            IdName          =   "WIDDT"
         End
         Begin MwfCtl.MWCTL_MED atxt 
            Height          =   285
            Index           =   44
            Left            =   1245
            TabIndex        =   5
            ToolTipText     =   "LMG Sales Rate"
            Top             =   1140
            Width           =   885
            _ExtentX        =   1561
            _ExtentY        =   503
            Alignment       =   2
            Enabled         =   0   'False
            Mask            =   "###0.00"
            MaxLength       =   7
            DataType        =   2
            IdName          =   "WIDLMGSAL"
         End
         Begin MwfCtl.MWCTL_MED atxt 
            Height          =   285
            Index           =   3
            Left            =   4785
            TabIndex        =   7
            ToolTipText     =   "Invoice Export Number"
            Top             =   285
            Width           =   3405
            _ExtentX        =   6006
            _ExtentY        =   503
            Enabled         =   0   'False
            MaxLength       =   20
            DataType        =   4
            IdName          =   "WIDEXPNO"
         End
         Begin MwfCtl.MWCTL_MED atxt 
            Height          =   285
            Index           =   5
            Left            =   4785
            TabIndex        =   9
            ToolTipText     =   "Customs FOB Value"
            Top             =   570
            Width           =   1425
            _ExtentX        =   2514
            _ExtentY        =   503
            Alignment       =   2
            Enabled         =   0   'False
            Mask            =   "#########0.00"
            MaxLength       =   13
            DataType        =   2
            ReCalcOn        =   "GrdID(IDZFOB)"
            IdName          =   "WTOTIDZFOB"
         End
         Begin MwfCtl.MWCTL_MED atxt 
            Height          =   285
            Index           =   6
            Left            =   1245
            TabIndex        =   8
            ToolTipText     =   "Invoice FOB Value"
            Top             =   570
            Width           =   1455
            _ExtentX        =   2566
            _ExtentY        =   503
            Alignment       =   2
            Enabled         =   0   'False
            Mask            =   "#########0.00"
            MaxLength       =   13
            DataType        =   2
            ReCalcOn        =   "GrdID(IDIFOB)"
            IdName          =   "WTOTIDIFOB"
         End
         Begin MwfCtl.MWCTL_MED atxt 
            Height          =   285
            Index           =   24
            Left            =   3225
            TabIndex        =   138
            ToolTipText     =   "LMP Sales Rate"
            Top             =   1140
            Width           =   885
            _ExtentX        =   1561
            _ExtentY        =   503
            Alignment       =   2
            Enabled         =   0   'False
            Mask            =   "###0.00"
            MaxLength       =   7
            DataType        =   2
            IdName          =   "WIDLMPSAL"
         End
         Begin MwfCtl.MWCTL_MED atxt 
            Height          =   285
            Index           =   127
            Left            =   7425
            TabIndex        =   151
            ToolTipText     =   "Invoice/Custom  FOB Percentage"
            Top             =   570
            Width           =   765
            _ExtentX        =   1349
            _ExtentY        =   503
            Alignment       =   2
            Enabled         =   0   'False
            Mask            =   "-#0.00"
            MaxLength       =   6
            DataType        =   2
            IdName          =   "WIDFOBPER"
         End
         Begin MwfCtl.MWCTL_MED atxt 
            Height          =   285
            Index           =   53
            Left            =   1245
            TabIndex        =   157
            ToolTipText     =   "LMG Cost Rate"
            Top             =   855
            Width           =   885
            _ExtentX        =   1561
            _ExtentY        =   503
            Alignment       =   2
            Enabled         =   0   'False
            Mask            =   "###0.00"
            MaxLength       =   7
            DataType        =   2
            IdName          =   "WIDLMGCST"
         End
         Begin MwfCtl.MWCTL_MED atxt 
            Height          =   285
            Index           =   55
            Left            =   3225
            TabIndex        =   158
            ToolTipText     =   "LMP Cost Rate"
            Top             =   855
            Width           =   885
            _ExtentX        =   1561
            _ExtentY        =   503
            Alignment       =   2
            Enabled         =   0   'False
            Mask            =   "###0.00"
            MaxLength       =   7
            DataType        =   2
            IdName          =   "WIDLMPCST"
         End
         Begin MwfCtl.MWCTL_MED atxt 
            Height          =   285
            Index           =   138
            Left            =   1245
            TabIndex        =   161
            ToolTipText     =   "Enter Gold Fineness"
            Top             =   1425
            Width           =   885
            _ExtentX        =   1561
            _ExtentY        =   503
            Alignment       =   2
            Enabled         =   0   'False
            Mask            =   "#0.000"
            MaxLength       =   6
            DataType        =   2
            IdName          =   "WIDGLDFINESS"
         End
         Begin MwfCtl.MWCTL_MED atxt 
            Height          =   285
            Index           =   139
            Left            =   3225
            TabIndex        =   162
            ToolTipText     =   "Enter Platinum Fineness"
            Top             =   1425
            Width           =   885
            _ExtentX        =   1561
            _ExtentY        =   503
            Alignment       =   2
            Enabled         =   0   'False
            Mask            =   "#0.000"
            MaxLength       =   6
            DataType        =   2
            IdName          =   "WIDPLFINESS"
         End
         Begin MwfCtl.MWCTL_MED atxt 
            Height          =   285
            Index           =   61
            Left            =   5325
            TabIndex        =   174
            ToolTipText     =   "LMS Cost Rate"
            Top             =   855
            Width           =   885
            _ExtentX        =   1561
            _ExtentY        =   503
            Alignment       =   2
            Enabled         =   0   'False
            Mask            =   "###0.00"
            MaxLength       =   7
            DataType        =   2
            IdName          =   "WIDLMSCST"
         End
         Begin MwfCtl.MWCTL_MED atxt 
            Height          =   285
            Index           =   73
            Left            =   9555
            TabIndex        =   185
            ToolTipText     =   "Invoice Customer Currency Code"
            Top             =   0
            Width           =   1095
            _ExtentX        =   1931
            _ExtentY        =   503
            Enabled         =   0   'False
            MaxLength       =   5
            DataType        =   4
            IdName          =   "WIDCMCURCD"
         End
         Begin MwfCtl.MWCTL_MED atxt 
            Height          =   285
            Index           =   74
            Left            =   9555
            TabIndex        =   187
            ToolTipText     =   "Conversion Factor"
            Top             =   285
            Width           =   1095
            _ExtentX        =   1931
            _ExtentY        =   503
            Enabled         =   0   'False
            Mask            =   "###0.0000"
            MaxLength       =   9
            DataType        =   2
            IdName          =   "WIDCNVFCT"
         End
         Begin MwfCtl.MWCTL_MED atxt 
            Height          =   285
            Index           =   75
            Left            =   9555
            TabIndex        =   189
            ToolTipText     =   "Multiplication/ Division"
            Top             =   570
            Width           =   345
            _ExtentX        =   609
            _ExtentY        =   503
            Enabled         =   0   'False
            DataType        =   4
            IdName          =   "WIDMULDIV"
         End
         Begin MwfCtl.MWCTL_MED atxt 
            Height          =   285
            Index           =   78
            Left            =   7305
            TabIndex        =   193
            ToolTipText     =   "LMS Cost Rate"
            Top             =   855
            Width           =   885
            _ExtentX        =   1561
            _ExtentY        =   503
            Alignment       =   2
            Enabled         =   0   'False
            Mask            =   "###0.00"
            MaxLength       =   7
            DataType        =   2
            IdName          =   "WIDLMLCST"
         End
         Begin MwfCtl.MWCTL_MED atxt 
            Height          =   285
            Index           =   79
            Left            =   7305
            TabIndex        =   194
            ToolTipText     =   "LMP Sales Rate"
            Top             =   1140
            Width           =   885
            _ExtentX        =   1561
            _ExtentY        =   503
            Alignment       =   2
            Enabled         =   0   'False
            Mask            =   "###0.00"
            MaxLength       =   7
            DataType        =   2
            IdName          =   "WIDLMLSAL"
         End
         Begin MwfCtl.MWCTL_MED atxt 
            Height          =   285
            Index           =   80
            Left            =   7305
            TabIndex        =   195
            ToolTipText     =   "Enter Palladium Fineness"
            Top             =   1425
            Width           =   885
            _ExtentX        =   1561
            _ExtentY        =   503
            Alignment       =   2
            Enabled         =   0   'False
            Mask            =   "#0.000"
            MaxLength       =   6
            DataType        =   2
            IdName          =   "WIDPDFINESS"
         End
         Begin MwfCtl.MWCTL_MED atxt 
            Height          =   285
            Index           =   81
            Left            =   7845
            TabIndex        =   199
            ToolTipText     =   "Invoice By Bag"
            Top             =   0
            Width           =   345
            _ExtentX        =   609
            _ExtentY        =   503
            Enabled         =   0   'False
            DataType        =   4
            IdName          =   "WIDINBYBAG"
         End
         Begin MwfCtl.MWCTL_MED atxt 
            Height          =   285
            Index           =   68
            Left            =   5325
            TabIndex        =   234
            ToolTipText     =   "LMP Sales Rate"
            Top             =   1140
            Width           =   885
            _ExtentX        =   1561
            _ExtentY        =   503
            Alignment       =   2
            Enabled         =   0   'False
            Mask            =   "###0.00"
            MaxLength       =   7
            DataType        =   2
            IdName          =   "WIDLMSSAL"
         End
         Begin MwfCtl.MWCTL_MED atxt 
            Height          =   285
            Index           =   69
            Left            =   5325
            TabIndex        =   235
            ToolTipText     =   "Enter Platinum Fineness"
            Top             =   1425
            Width           =   885
            _ExtentX        =   1561
            _ExtentY        =   503
            Alignment       =   2
            Enabled         =   0   'False
            Mask            =   "#0.000"
            MaxLength       =   6
            DataType        =   2
            IdName          =   "WIDSLFINESS"
         End
         Begin MwfCtl.MWCTL_MED atxt 
            Height          =   285
            Index           =   82
            Left            =   9555
            TabIndex        =   237
            ToolTipText     =   "Invoice Customer Currency Code"
            Top             =   855
            Width           =   1095
            _ExtentX        =   1931
            _ExtentY        =   503
            MaxLength       =   5
            DataType        =   4
            IdName          =   "WINSRT"
         End
         Begin MwfCtl.MWCTL_MED atxt 
            Height          =   285
            Index           =   86
            Left            =   1200
            TabIndex        =   243
            ToolTipText     =   "Invoice Customer Currency Code"
            Top             =   8920
            Width           =   3405
            _ExtentX        =   6006
            _ExtentY        =   503
            Enabled         =   0   'False
            MaxLength       =   20
            DataType        =   4
            IdName          =   "WIRCUSTRMCD"
         End
         Begin MwfCtl.MWCTL_MED atxt 
            Height          =   285
            Index           =   300
            Left            =   2760
            TabIndex        =   244
            ToolTipText     =   "Tag for chain logic for current customer"
            Top             =   360
            Visible         =   0   'False
            Width           =   345
            _ExtentX        =   609
            _ExtentY        =   503
            Enabled         =   0   'False
            DataType        =   4
            ReCalcOn        =   "WIDCMCD"
            IdName          =   "WCHNCUSTYN"
         End
         Begin MwfCtl.MWCTL_FLX GrdIR 
            Height          =   2880
            Left            =   120
            TabIndex        =   11
            Top             =   6000
            Width           =   15105
            _ExtentX        =   26644
            _ExtentY        =   5080
            Cols            =   32
            colname1        =   "IRSRNO"
            heading1        =   "Sr.  "
            datafld1        =   "IrSrNo"
            datatype1       =   1
            mask1           =   "##0"
            maxlength1      =   3
            style1          =   2
            tooltiptext1    =   "Enter Raw Material Sr No."
            colname2        =   "IRRMCD"
            heading2        =   "Rm Code                                   "
            datafld2        =   "IrRmCd"
            datatype2       =   4
            recalcparent2   =   $"EmrFrmInvDsg.frx":01F8
            maxlength2      =   16
            tooltiptext2    =   "Enter Raw Material Code"
            colname3        =   "IRLOTNO"
            heading3        =   "Lot No.                              "
            datafld3        =   "IrLotNo"
            datatype3       =   4
            maxlength3      =   16
            tooltiptext3    =   "Enter Lot No."
            colname4        =   "IRRMSZ"
            heading4        =   "Size        "
            datafld4        =   "IrRmSz"
            datatype4       =   2
            mask4           =   "##0.0000"
            recalcparent4   =   "GrdIR(IRCUSTRMCD)"
            maxlength4      =   8
            tooltiptext4    =   "Enter Raw Material Size"
            colname5        =   "IRRMSZ2"
            heading5        =   "Bdth    "
            datafld5        =   "IrRmSz2"
            datatype5       =   2
            mask5           =   "#0.000"
            maxlength5      =   6
            colname6        =   "IRRMSZ3"
            heading6        =   "Dpth    "
            datafld6        =   "IrRmSz3"
            datatype6       =   2
            mask6           =   "#0.000"
            maxlength6      =   6
            colname7        =   "IRSTKRT"
            heading7        =   "StockRate    "
            datafld7        =   "IrStkRt"
            datatype7       =   2
            mask7           =   "#######0.000"
            maxlength7      =   12
            tooltiptext7    =   "Enter Raw Material StockRate"
            colname8        =   "WQWSAL"
            heading8        =   "Q/W"
            datatype8       =   4
            recalcon8       =   "GrdIR(IRRMCD)"
            maxlength8      =   1
            style8          =   2
            colname9        =   "IRRMQTY"
            heading9        =   "Qty        "
            datafld9        =   "IrRmQty"
            datatype9       =   1
            mask9           =   "#####0"
            recalcparent9   =   "GrdIR(IRSETAVAL),GrdIR(IRSETIVAL),GrdIR(IRSETZVAL),GrdIR(IRRMIVAL),GrdIR(IRRMZVAL),GrdIR(IRRMAVAL)"
            maxlength9      =   6
            tooltiptext9    =   "Enter Raw Material Quantity"
            colname10       =   "IRRMAWT"
            heading10       =   "Actual Wt    "
            datafld10       =   "IrRmAWt"
            datatype10      =   2
            mask10          =   "#####0.000"
            recalcparent10  =   "GrdIR(IRRMAVAL),GrdID(IDAGLDASWT),GrdID(IDALABASWT),GrdID(WIDADIAWT),GrdID(WIDACSWT)"
            maxlength10     =   10
            style10         =   2
            tooltiptext10   =   "Actual Raw Material Weight"
            colname11       =   "IRRMIWT"
            heading11       =   "Inv Wt        "
            datafld11       =   "IrRmIWt"
            datatype11      =   2
            mask11          =   "#####0.000"
            recalcparent11  =   "GrdIR(IRRMIVAL),GrdID(IDIGLDASWT),GrdID(IDILABASWT),GrdID(WIDIDIAWT),GrdID(WIDICSWT)"
            maxlength11     =   10
            tooltiptext11   =   "Enter Invoice Raw Material Weight"
            colname12       =   "IRRMZWT"
            heading12       =   "Custm Wt    "
            datafld12       =   "IrRmZWt"
            datatype12      =   2
            mask12          =   "#####0.000"
            recalcparent12  =   "GrdID(IDCUSTMMTCHWT),GrdIR(IRRMZVAL),GrdIR(IRRMZPURERT),GrdID(IDZGLDASWT),GrdID(IDZLABASWT),GrdID(WIDZDIAWT),GrdID(WIDZCSWT)"
            maxlength12     =   10
            tooltiptext12   =   "Enter Customs Raw Material Weight"
            colname13       =   "IRRMART"
            datafld13       =   "IrRmARt"
            datatype13      =   2
            mask13          =   "#######0.000"
            recalcparent13  =   "GrdIR(IRRMAVAL)"
            maxlength13     =   12
            style13         =   0
            tooltiptext13   =   "Actual Raw Material Rate"
            colname14       =   "IRRMIRT"
            heading14       =   "Inv Rt                "
            datafld14       =   "IrRmIRt"
            datatype14      =   2
            mask14          =   "#######0.000"
            recalcon14      =   "GrdIR(IRRMCD)"
            recalcparent14  =   "GrdIR(IRRMZRT),GrdIR(IRRMIVAL)"
            maxlength14     =   12
            tooltiptext14   =   "Enter Invoice Raw Material Rate"
            colname15       =   "IRRMZRT"
            heading15       =   "Custm Rt         "
            datafld15       =   "IrRmZRt"
            datatype15      =   2
            mask15          =   "#######0.000"
            recalcon15      =   "GrdIR(IRRMIRT),GrdIR(IRRMCD)"
            recalcparent15  =   "GrdIR(IRRMZVAL),GrdIR(IRRMZPURERT)"
            maxlength15     =   12
            tooltiptext15   =   "Enter Customs Raw Material Rate"
            colname16       =   "IRRMZPURERT"
            heading16       =   "Custm PureRt"
            datafld16       =   "IrRmZPureRt"
            datatype16      =   2
            mask16          =   "#######0.000"
            recalcon16      =   "GrdIR(IRRMCD),GrdIR(IRRMZRT),GrdIR(IRRMZVAL),GrdIR(IRRMZWT)"
            maxlength16     =   12
            style16         =   2
            tooltiptext16   =   "Enter Customs Raw Material Pure Rate"
            colname17       =   "IRRMAVAL"
            datafld17       =   "IrRmAVal"
            datatype17      =   2
            mask17          =   "########0.00"
            recalcon17      =   "GrdIR(IRRMAWT),GrdIR(IRRMART),GrdIR(IRRMCD),GrdIR(IRRMQTY)"
            recalcparent17  =   "GrdID(IDAVAL)"
            maxlength17     =   12
            style17         =   0
            tooltiptext17   =   "Actual Raw Material Value"
            colname18       =   "IRRMIVAL"
            heading18       =   "Inv Val             "
            datafld18       =   "IrRmIVal"
            datatype18      =   2
            mask18          =   "########0.00"
            recalcon18      =   "GrdIR(IRRMIWT),GrdIR(IRRMIRT),GrdIR(IRMAINMET),GrdIR(IRRMCD),GrdIR(IRRMQTY)"
            recalcparent18  =   "GrdID(IDIVAL)"
            maxlength18     =   12
            style18         =   2
            tooltiptext18   =   "Enter Invoice Raw Material Value"
            colname19       =   "IRRMZVAL"
            heading19       =   "Custm Val       "
            datafld19       =   "IrRmZVal"
            datatype19      =   2
            mask19          =   "########0.00"
            recalcon19      =   "GrdIR(IRRMZWT),GrdIR(IRRMZRT),GrdIR(IRMAINMET),GrdIR(IRRMCD),GrdIR(IRRMQTY)"
            recalcparent19  =   "GrdID(IDZVAL),GrdIR(IRRMZPURERT)"
            maxlength19     =   12
            style19         =   2
            tooltiptext19   =   "Enter Customs Raw Material Value"
            colname20       =   "IRSETSCD"
            heading20       =   "Setting        "
            datafld20       =   "IrSetSCd"
            datatype20      =   4
            recalcparent20  =   "GrdIR(IRSETAVAL),GrdIR(IRSETIVAL),GrdIR(IRSETZVAL),GrdIR(IRSETIRT)"
            maxlength20     =   8
            tooltiptext20   =   "Enter Setting Code"
            colname21       =   "IRSETART"
            datafld21       =   "IrSetARt"
            datatype21      =   2
            mask21          =   "#####0.00"
            recalcparent21  =   "GrdIR(IRSETAVAL)"
            maxlength21     =   9
            style21         =   0
            tooltiptext21   =   "Actual Setting Rate"
            colname22       =   "IRSETIRT"
            heading22       =   "Inv SetRt  "
            datafld22       =   "IrSetIRt"
            datatype22      =   2
            mask22          =   "#####0.00"
            recalcon22      =   "GrdIR(IRSETSCD)"
            recalcparent22  =   "GrdIR(IRSETIVAL),GrdIR(IRSETZRT)"
            maxlength22     =   9
            tooltiptext22   =   "Enter Invoice Setting Rate"
            colname23       =   "IRSETZRT"
            heading23       =   "Custm SetRt"
            datafld23       =   "IrSetZRt"
            datatype23      =   2
            mask23          =   "#####0.00"
            recalcon23      =   "GrdIR(IRSETIRT)"
            recalcparent23  =   "GrdIR(IRSETZVAL)"
            maxlength23     =   9
            tooltiptext23   =   "Enter Customs Setting Rate"
            colname24       =   "IRSETAVAL"
            datafld24       =   "IrSetAVal"
            datatype24      =   2
            mask24          =   "######0.00"
            recalcon24      =   "GrdIR(IRSETSCD),GrdIR(IRRMQTY),GrdIR(IRSETART)"
            recalcparent24  =   "GrdID(IDAVAL)"
            maxlength24     =   10
            style24         =   0
            tooltiptext24   =   "Actual Setting Value"
            colname25       =   "IRSETIVAL"
            heading25       =   "Inv SetVal"
            datafld25       =   "IrSetIVal"
            datatype25      =   2
            mask25          =   "######0.00"
            recalcon25      =   "GrdIR(IRSETSCD),GrdIR(IRRMQTY),GrdIR(IRSETIRT)"
            recalcparent25  =   "GrdID(IDIVAL)"
            maxlength25     =   10
            style25         =   2
            tooltiptext25   =   "Enter Invoice Setting Value"
            colname26       =   "IRSETZVAL"
            heading26       =   "Custm SetVal"
            datafld26       =   "IrSetZVal"
            datatype26      =   2
            mask26          =   "######0.00"
            recalcon26      =   "GrdIR(IRSETSCD),GrdIR(IRRMQTY),GrdIR(IRSETZRT)"
            recalcparent26  =   "GrdID(IDZVAL)"
            maxlength26     =   10
            style26         =   2
            tooltiptext26   =   "Enter Customs Setting Value"
            colname27       =   "IRALYSALRT"
            heading27       =   "Aly Sal Rt"
            datafld27       =   "IrAlySalRt"
            datatype27      =   2
            mask27          =   "######0.00"
            maxlength27     =   10
            style27         =   2
            tooltiptext27   =   "Alloy Sale Rate"
            colname28       =   "IRMAINMET"
            heading28       =   "MMet"
            datafld28       =   "IrMainMet"
            datatype28      =   4
            recalcon28      =   "GrdIR(IRRMCD)"
            recalcparent28  =   "GrdID(IDCUSTMMTCHWT),GrdIR(IRRMIVAL),GrdIR(IRRMZVAL),GrdID(IDAGLDASWT),GrdID(IDIGLDASWT),GrdID(IDZGLDASWT)"
            maxlength28     =   1
            tooltiptext28   =   "Enter Main Metal (Y/N)"
            colname29       =   "IRRMZCD"
            heading29       =   "Custm Code                               "
            datafld29       =   "IrRmZCd"
            datatype29      =   4
            recalcon29      =   "GrdIR(IRRMCD)"
            maxlength29     =   16
            style29         =   2
            tooltiptext29   =   "Enter Customs Code"
            colname30       =   "IRPRTKEY"
            heading30       =   "Prtn "
            datafld30       =   "IrPrtKey"
            maxlength30     =   1
            style30         =   0
            colname31       =   "IRCUSTRMCD"
            datafld31       =   "IrCustRmCd"
            datatype31      =   4
            recalcon31      =   "GrdIR(IRRMCD),GrdIR(IRRMSZ)"
            maxlength31     =   20
            style31         =   0
         End
         Begin VB.Label Label3 
            BackStyle       =   0  'Transparent
            Caption         =   "Cust RmCd"
            BeginProperty Font 
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
            TabIndex        =   242
            Top             =   8920
            Width           =   1185
         End
         Begin VB.Label ALBL 
            BackStyle       =   0  'Transparent
            Caption         =   "Sort By"
            BeginProperty Font 
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
            Left            =   8370
            TabIndex        =   238
            Top             =   855
            Width           =   1005
         End
         Begin VB.Label LblLmsSal 
            BackStyle       =   0  'Transparent
            Caption         =   "LMS Sales"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Left            =   4220
            TabIndex        =   236
            Top             =   1140
            Width           =   1095
         End
         Begin VB.Label ALBL 
            BackStyle       =   0  'Transparent
            Caption         =   "Inv By Bag"
            BeginProperty Font 
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
            Left            =   6780
            TabIndex        =   200
            Top             =   0
            Width           =   1125
         End
         Begin VB.Label LblLmlCst 
            BackStyle       =   0  'Transparent
            Caption         =   "LML Cost"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Left            =   6250
            TabIndex        =   198
            Top             =   855
            Width           =   1065
         End
         Begin VB.Label LblLmlSal 
            BackStyle       =   0  'Transparent
            Caption         =   "LML Sales"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Left            =   6250
            TabIndex        =   197
            Top             =   1140
            Width           =   1095
         End
         Begin VB.Label LblPdFiness 
            BackStyle       =   0  'Transparent
            Caption         =   "Pd Finess"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Left            =   6250
            TabIndex        =   196
            Top             =   1440
            Width           =   1125
         End
         Begin VB.Label ALBL 
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
            Index           =   47
            Left            =   8370
            TabIndex        =   190
            Top             =   570
            Width           =   1005
         End
         Begin VB.Label ALBL 
            BackStyle       =   0  'Transparent
            Caption         =   "Conv Fact"
            BeginProperty Font 
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
            Left            =   8370
            TabIndex        =   188
            Top             =   285
            Width           =   1005
         End
         Begin VB.Label ALBL 
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
            Index           =   37
            Left            =   8370
            TabIndex        =   186
            Top             =   0
            Width           =   1005
         End
         Begin VB.Label LblSlFiness 
            BackStyle       =   0  'Transparent
            Caption         =   "Slv Finess"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Left            =   4220
            TabIndex        =   176
            Top             =   1425
            Width           =   1125
         End
         Begin VB.Label LblLmsCst 
            BackStyle       =   0  'Transparent
            Caption         =   "LMS Cost"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Left            =   4220
            TabIndex        =   175
            Top             =   855
            Width           =   1065
         End
         Begin VB.Label LblGldFiness 
            BackStyle       =   0  'Transparent
            Caption         =   "Gld Finess"
            BeginProperty Font 
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
            TabIndex        =   164
            Top             =   1425
            Width           =   1185
         End
         Begin VB.Label LblPlFiness 
            BackStyle       =   0  'Transparent
            Caption         =   "Plt Finess"
            BeginProperty Font 
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
            TabIndex        =   163
            Top             =   1425
            Width           =   1125
         End
         Begin VB.Label ALBL 
            BackStyle       =   0  'Transparent
            Caption         =   "LMG Cost"
            BeginProperty Font 
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
            Left            =   30
            TabIndex        =   160
            Top             =   855
            Width           =   1185
         End
         Begin VB.Label ALBL 
            BackStyle       =   0  'Transparent
            Caption         =   "LMP Cost"
            BeginProperty Font 
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
            Left            =   2130
            TabIndex        =   159
            Top             =   855
            Width           =   1065
         End
         Begin VB.Label ALBL 
            BackStyle       =   0  'Transparent
            Caption         =   "FOB %age "
            BeginProperty Font 
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
            Left            =   6300
            TabIndex        =   150
            Top             =   570
            Width           =   1185
         End
         Begin VB.Label ALBL 
            BackStyle       =   0  'Transparent
            Caption         =   "LMP Sales"
            BeginProperty Font 
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
            Left            =   2130
            TabIndex        =   139
            Top             =   1155
            Width           =   1095
         End
         Begin VB.Label ALBL 
            BackStyle       =   0  'Transparent
            Caption         =   "Tot Inv FOB"
            BeginProperty Font 
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
            Left            =   30
            TabIndex        =   116
            Top             =   570
            Width           =   1305
         End
         Begin VB.Label ALBL 
            BackStyle       =   0  'Transparent
            Caption         =   "Tot Customs FOB"
            BeginProperty Font 
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
            Left            =   3090
            TabIndex        =   115
            Top             =   570
            Width           =   1665
         End
         Begin VB.Label ALBL 
            BackStyle       =   0  'Transparent
            Caption         =   "Inv ExpNo"
            BeginProperty Font 
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
            Left            =   3780
            TabIndex        =   106
            Top             =   285
            Width           =   1125
         End
         Begin VB.Label ALBL 
            BackStyle       =   0  'Transparent
            Caption         =   "LMG Sales"
            BeginProperty Font 
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
            Left            =   30
            TabIndex        =   98
            Top             =   1125
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
            Index           =   41
            Left            =   4320
            TabIndex        =   97
            Top             =   0
            Width           =   585
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
            Index           =   2
            Left            =   30
            TabIndex        =   93
            Top             =   285
            Width           =   1005
         End
      End
      Begin VB.Frame FraIL 
         Height          =   3045
         Left            =   150
         TabIndex        =   16
         Top             =   6165
         Width           =   15120
         Begin MwfCtl.MWCTL_FLX GrdIL 
            Height          =   2550
            Left            =   90
            TabIndex        =   17
            ToolTipText     =   "Order Amendment"
            Top             =   360
            Width           =   9105
            _ExtentX        =   16060
            _ExtentY        =   4498
            Cols            =   14
            colname1        =   "ILSRNO"
            heading1        =   "Sr."
            datafld1        =   "iLSrNo"
            datatype1       =   1
            mask1           =   "#0"
            maxlength1      =   2
            style1          =   2
            tooltiptext1    =   "Enter Sr No."
            colname2        =   "ILMCD"
            heading2        =   "Lab Cd "
            datafld2        =   "iLMCd"
            datatype2       =   4
            recalcparent2   =   "GrdIL(ILQW),GrdIL(ILAVAL),GrdIL(ILIVAL),GrdIL(ILZVAL),GrdIL(ILIRT)"
            maxlength2      =   5
            tooltiptext2    =   "Enter Invoice Labour Main Code"
            colname3        =   "ILSCD"
            heading3        =   "Lab SCd"
            datafld3        =   "iLSCd"
            datatype3       =   4
            recalcparent3   =   "GrdIL(ILQW),GrdIL(ILAVAL),GrdIL(ILIVAL),GrdIL(ILZVAL),GrdIL(ILIRT)"
            maxlength3      =   8
            tooltiptext3    =   "Enter Invoice Labour Sub Code"
            colname4        =   "ILQW"
            heading4        =   "By Q/W"
            datafld4        =   "iLQw"
            datatype4       =   4
            recalcon4       =   "GrdIL(ILMCD),GrdIL(ILSCD)"
            recalcparent4   =   "GrdIL(ILQTY),GrdIL(ILIVAL),GrdIL(ILZVAL)"
            maxlength4      =   1
            style4          =   2
            colname5        =   "ILCSTQW"
            datafld5        =   "iLCstQw"
            datatype5       =   4
            recalcparent5   =   "GrdIL(ILAVAL)"
            maxlength5      =   1
            style5          =   0
            colname6        =   "ILQTY"
            heading6        =   "Qty"
            datafld6        =   "iLQty"
            datatype6       =   1
            mask6           =   "##0"
            recalcon6       =   "GrdIL(ILQW)"
            recalcparent6   =   "GrdIL(ILAVAL),GrdIL(ILIVAL),GrdIL(ILZVAL)"
            maxlength6      =   3
            tooltiptext6    =   "Enter Invoice Labour Quantity"
            colname7        =   "ILART"
            datafld7        =   "iLARt"
            datatype7       =   2
            mask7           =   "#######0.00"
            recalcparent7   =   "GrdIL(ILAVAL)"
            maxlength7      =   11
            style7          =   0
            tooltiptext7    =   "Actual Labour Sales Rate"
            colname8        =   "ILIRT"
            heading8        =   "Inv Rt             "
            datafld8        =   "iLiRt"
            datatype8       =   2
            mask8           =   "-#######0.00"
            recalcon8       =   "GrdIL(ILMCD),GrdIL(ILSCD)"
            recalcparent8   =   "GrdIL(ILIVAL),GrdIL(ILZRT)"
            maxlength8      =   12
            tooltiptext8    =   "Enter Invoice Labour Sales Rate"
            colname9        =   "ILZRT"
            heading9        =   "Customs Rt     "
            datafld9        =   "iLZRt"
            datatype9       =   2
            mask9           =   "-#######0.00"
            recalcon9       =   "GrdIL(ILIRT),GrdID(IDIFOB),GrdID(IDZFOB)"
            recalcparent9   =   "GrdIL(ILZVAL)"
            maxlength9      =   12
            tooltiptext9    =   "Enter Customs Labour Sales Rate"
            colname10       =   "ILAVAL"
            datafld10       =   "iLAVal"
            datatype10      =   2
            mask10          =   "-#######0.00"
            recalcon10      =   "GrdIL(ILMCD),GrdIL(ILSCD),GrdIL(ILCSTQW),GrdIL(ILQTY),GrdIL(ILART)"
            recalcparent10  =   "GrdID(IDAVAL)"
            maxlength10     =   12
            style10         =   0
            tooltiptext10   =   "Actual Labour Sales Value"
            colname11       =   "ILIVAL"
            heading11       =   "Inv Val           "
            datafld11       =   "iLiVal"
            datatype11      =   2
            mask11          =   "-#######0.00"
            recalcon11      =   "GrdIL(ILMCD),GrdIL(ILSCD),GrdIL(ILQW),GrdIL(ILQTY),GrdIL(ILIRT)"
            recalcparent11  =   "GrdID(IDIVAL)"
            maxlength11     =   12
            style11         =   2
            tooltiptext11   =   "Enter Invoice Labour Sales Value"
            colname12       =   "ILZVAL"
            heading12       =   "Customs Val   "
            datafld12       =   "iLZVal"
            datatype12      =   2
            mask12          =   "-#######0.00"
            recalcon12      =   "GrdIL(ILMCD),GrdIL(ILSCD),GrdIL(ILQW),GrdIL(ILQTY),GrdIL(ILZRT)"
            recalcparent12  =   "GrdID(IDZVAL)"
            maxlength12     =   12
            style12         =   2
            tooltiptext12   =   "Enter Customs Labour Sales Value"
            colname13       =   "ILPRTKEY"
            heading13       =   "Prtn "
            datafld13       =   "IlPrtKey"
            maxlength13     =   1
            style13         =   0
         End
         Begin VB.Label LblMain 
            BackStyle       =   0  'Transparent
            BorderStyle     =   1  'Fixed Single
            Caption         =   "  Invoice Design Labour"
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
            TabIndex        =   171
            Top             =   90
            Width           =   15120
         End
      End
      Begin VB.Frame FraIB 
         Height          =   3045
         Left            =   150
         TabIndex        =   14
         Top             =   6165
         Width           =   15120
         Begin MwfCtl.MWCTL_FLX GrdIB 
            Height          =   2580
            Left            =   90
            TabIndex        =   15
            Top             =   360
            Width           =   8655
            _ExtentX        =   15266
            _ExtentY        =   4551
            Cols            =   13
            colname1        =   "IFSRNO"
            heading1        =   "SrNo  "
            datafld1        =   "IfSrNo"
            datatype1       =   1
            mask1           =   "####0"
            maxlength1      =   5
            style1          =   2
            tooltiptext1    =   "Inv Bag Serial Number"
            colname2        =   "IFFDBYY"
            heading2        =   "BYy"
            datafld2        =   "IfFdBYy"
            datatype2       =   4
            recalcon2       =   "GrdIB(IFFDTC),GrdIB(IFFDYY),GrdIB(IFFDCHR),GrdIB(IFFDNO),GrdIB(IFFDSR)"
            recalcparent2   =   "GrdIB(IFFDBNO),GrdIB(IFFDTC),GrdIB(IFFDYY),GrdIB(IFFDCHR),GrdIB(IFFDNO),GrdIB(IFFDSR)"
            maxlength2      =   2
            tooltiptext2    =   "Enter Bag Year"
            colname3        =   "IFFDBCHR"
            heading3        =   "BChr   "
            datafld3        =   "IfFdBChr"
            datatype3       =   4
            recalcon3       =   "GrdIB(IFFDTC),GrdIB(IFFDYY),GrdIB(IFFDCHR),GrdIB(IFFDNO),GrdIB(IFFDSR)"
            recalcparent3   =   "GrdIB(IFFDBNO),GrdIB(IFFDTC),GrdIB(IFFDYY),GrdIB(IFFDCHR),GrdIB(IFFDNO),GrdIB(IFFDSR)"
            maxlength3      =   3
            tooltiptext3    =   "Enter Bag Character"
            colname4        =   "IFFDBNO"
            heading4        =   "BNo          "
            datafld4        =   "IfFdBNo"
            datatype4       =   1
            mask4           =   "######0"
            recalcon4       =   "GrdIB(IFFDBYY),GrdIB(IFFDBCHR),GrdIB(IFFDTC),GrdIB(IFFDYY),GrdIB(IFFDCHR),GrdIB(IFFDNO),GrdIB(IFFDSR)"
            recalcparent4   =   "GrdIB(IFFDTC),GrdIB(IFFDYY),GrdIB(IFFDCHR),GrdIB(IFFDNO),GrdIB(IFFDSR)"
            maxlength4      =   7
            tooltiptext4    =   "Enter Bag Number"
            colname5        =   "IFFDTC"
            heading5        =   "FgTc "
            datafld5        =   "IfFdTc"
            datatype5       =   4
            recalcon5       =   "GrdIB(IFFDBYY),GrdIB(IFFDBCHR),GrdIB(IFFDBNO)"
            recalcparent5   =   "GrdIB(IFFDBYY),GrdIB(IFFDBCHR),GrdIB(IFFDBNO),GrdIB(IFFDNO),GrdIB(IFFDSR),GrdIB(IFQTY),GrdIB(IFGRSWT)"
            maxlength5      =   3
            tooltiptext5    =   "Enter Finished Goods Tc"
            colname6        =   "IFFDYY"
            heading6        =   "FgYy"
            datafld6        =   "IfFdYy"
            datatype6       =   4
            recalcon6       =   "GrdIB(IFFDBYY),GrdIB(IFFDBCHR),GrdIB(IFFDBNO)"
            recalcparent6   =   "GrdIB(IFFDBYY),GrdIB(IFFDBCHR),GrdIB(IFFDBNO),GrdIB(IFFDNO),GrdIB(IFFDSR),GrdIB(IFQTY),GrdIB(IFGRSWT)"
            maxlength6      =   2
            tooltiptext6    =   "Enter Finished Goods Year"
            colname7        =   "IFFDCHR"
            heading7        =   "FgChr"
            datafld7        =   "IfFdChr"
            datatype7       =   4
            recalcon7       =   "GrdIB(IFFDBYY),GrdIB(IFFDBCHR),GrdIB(IFFDBNO)"
            recalcparent7   =   "GrdIB(IFFDBYY),GrdIB(IFFDBCHR),GrdIB(IFFDBNO),GrdIB(IFFDNO),GrdIB(IFFDSR),GrdIB(IFQTY),GrdIB(IFGRSWT)"
            maxlength7      =   3
            tooltiptext7    =   "Enter Finished Goods Character"
            colname8        =   "IFFDNO"
            heading8        =   "FgNo      "
            datafld8        =   "IfFdNo"
            datatype8       =   1
            mask8           =   "#####0"
            recalcon8       =   "GrdIB(IFFDBYY),GrdIB(IFFDBCHR),GrdIB(IFFDBNO),GrdIB(IFFDTC),GrdIB(IFFDYY),GrdIB(IFFDCHR)"
            recalcparent8   =   "GrdIB(IFFDBYY),GrdIB(IFFDBCHR),GrdIB(IFFDBNO),GrdIB(IFFDSR),GrdIB(IFQTY),GrdIB(IFGRSWT)"
            maxlength8      =   6
            tooltiptext8    =   "Enter Finished Goods Number"
            colname9        =   "IFFDSR"
            heading9        =   "FgSr"
            datafld9        =   "IfFdSr"
            datatype9       =   1
            mask9           =   "###0"
            recalcon9       =   "GrdIB(IFFDBYY),GrdIB(IFFDBCHR),GrdIB(IFFDBNO),GrdIB(IFFDTC),GrdIB(IFFDYY),GrdIB(IFFDCHR),GrdIB(IFFDNO)"
            recalcparent9   =   "GrdIB(IFFDBYY),GrdIB(IFFDBCHR),GrdIB(IFFDBNO),GrdIB(IFQTY),GrdIB(IFGRSWT)"
            maxlength9      =   4
            tooltiptext9    =   "Enter Finished Goods Serial"
            colname10       =   "IFQTY"
            heading10       =   "Bag Qty"
            datafld10       =   "IfQty"
            datatype10      =   2
            mask10          =   "##0.0"
            recalcon10      =   "GrdIB(IFFDTC),GrdIB(IFFDYY),GrdIB(IFFDCHR),GrdIB(IFFDNO),GrdIB(IFFDSR)"
            recalcparent10  =   "GrdID(IDQTY)"
            maxlength10     =   5
            tooltiptext10   =   "Bag Quantity"
            colname11       =   "IFGRSWT"
            heading11       =   "Gross Wt     "
            datafld11       =   "IfGrsWt"
            datatype11      =   2
            mask11          =   "#####0.000"
            recalcon11      =   "GrdIB(IFFDTC),GrdIB(IFFDYY),GrdIB(IFFDCHR),GrdIB(IFFDNO),GrdIB(IFFDSR)"
            maxlength11     =   10
            style11         =   2
            tooltiptext11   =   "Gross Weight"
            colname12       =   "IFPRTKEY"
            heading12       =   "Prtn "
            datafld12       =   "IfPrtKey"
            maxlength12     =   1
            style12         =   0
         End
         Begin VB.Label LblMain 
            BackStyle       =   0  'Transparent
            BorderStyle     =   1  'Fixed Single
            Caption         =   "  Invoice Design Bag"
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
            TabIndex        =   169
            Top             =   90
            Width           =   15120
         End
      End
      Begin VB.Frame FraRefRt 
         Height          =   1650
         Left            =   1680
         TabIndex        =   201
         Top             =   7560
         Visible         =   0   'False
         Width           =   12360
         Begin VB.Frame FraRt 
            Height          =   1215
            Left            =   150
            TabIndex        =   202
            Top             =   360
            Width           =   11475
            Begin VB.TextBox txtPwd 
               Height          =   285
               IMEMode         =   3  'DISABLE
               Left            =   1560
               MaxLength       =   30
               PasswordChar    =   "*"
               TabIndex        =   204
               ToolTipText     =   "Enter SEO Password"
               Top             =   555
               Width           =   1095
            End
            Begin MwfCtl.MWCTL_MED atxt 
               Height          =   285
               Index           =   100
               Left            =   1560
               TabIndex        =   203
               ToolTipText     =   "Enter Date for Refreshing Rates"
               Top             =   270
               Width           =   855
               _ExtentX        =   1508
               _ExtentY        =   503
               MaxLength       =   8
               DataType        =   3
               ReCalcParent    =   "WINLMGDT,WINLMPDT,WINLMSDT,WINLMLDT,WINLMLCST,WINLMSCST,WINLMPCST,WINLMGCST,WINCNVRT"
               IdName          =   "WINREFRESHDT"
            End
            Begin MwfCtl.MWCTL_MED atxt 
               Height          =   285
               Index           =   102
               Left            =   6210
               TabIndex        =   212
               ToolTipText     =   "Enter LME Platinum Sales Value"
               Top             =   840
               Width           =   885
               _ExtentX        =   1561
               _ExtentY        =   503
               Alignment       =   2
               Mask            =   "####0.00"
               MaxLength       =   8
               DataType        =   2
               ReCalcOn        =   "WINLMPCST"
               IdName          =   "WINLMPSAL"
            End
            Begin MwfCtl.MWCTL_MED atxt 
               Height          =   285
               Index           =   104
               Left            =   4110
               TabIndex        =   209
               ToolTipText     =   "Enter LME Gold Sales Value"
               Top             =   840
               Width           =   855
               _ExtentX        =   1508
               _ExtentY        =   503
               Alignment       =   2
               Mask            =   "####0.00"
               MaxLength       =   8
               DataType        =   2
               ReCalcOn        =   "WINLMGCST"
               IdName          =   "WINLMGSAL"
            End
            Begin MwfCtl.MWCTL_MED atxt 
               Height          =   285
               Index           =   98
               Left            =   4110
               TabIndex        =   208
               ToolTipText     =   "Enter LME Gold Cost Value"
               Top             =   555
               Width           =   855
               _ExtentX        =   1508
               _ExtentY        =   503
               Alignment       =   2
               Mask            =   "####0.00"
               MaxLength       =   8
               DataType        =   2
               ReCalcParent    =   "WINLMGSAL"
               ReCalcOn        =   "WINREFRESHDT"
               IdName          =   "WINLMGCST"
            End
            Begin MwfCtl.MWCTL_MED atxt 
               Height          =   285
               Index           =   105
               Left            =   6210
               TabIndex        =   210
               ToolTipText     =   "Enter Date Of LME Platinum Rate"
               Top             =   270
               Width           =   885
               _ExtentX        =   1561
               _ExtentY        =   503
               MaxLength       =   8
               DataType        =   3
               ReCalcOn        =   "WINREFRESHDT"
               IdName          =   "WINLMPDT"
            End
            Begin MwfCtl.MWCTL_MED atxt 
               Height          =   285
               Index           =   99
               Left            =   4110
               TabIndex        =   207
               ToolTipText     =   "Enter  Date Of LME Gold Rate"
               Top             =   270
               Width           =   885
               _ExtentX        =   1561
               _ExtentY        =   503
               MaxLength       =   8
               DataType        =   3
               ReCalcOn        =   "WINREFRESHDT"
               IdName          =   "WINLMGDT"
            End
            Begin MwfCtl.MWCTL_MED atxt 
               Height          =   285
               Index           =   166
               Left            =   8340
               TabIndex        =   215
               ToolTipText     =   "Enter LME Silver Sales Value"
               Top             =   840
               Width           =   885
               _ExtentX        =   1561
               _ExtentY        =   503
               Alignment       =   2
               Mask            =   "####0.00"
               MaxLength       =   8
               DataType        =   2
               ReCalcOn        =   "WINLMSCST"
               IdName          =   "WINLMSSAL"
            End
            Begin MwfCtl.MWCTL_MED atxt 
               Height          =   285
               Index           =   168
               Left            =   8340
               TabIndex        =   213
               ToolTipText     =   "Enter Date Of LME Silver Rate"
               Top             =   270
               Width           =   885
               _ExtentX        =   1561
               _ExtentY        =   503
               MaxLength       =   8
               DataType        =   3
               ReCalcOn        =   "WINREFRESHDT"
               IdName          =   "WINLMSDT"
            End
            Begin MwfCtl.MWCTL_MED atxt 
               Height          =   285
               Index           =   167
               Left            =   8340
               TabIndex        =   214
               ToolTipText     =   "Enter LME Platinum Cost Value"
               Top             =   555
               Width           =   885
               _ExtentX        =   1561
               _ExtentY        =   503
               Alignment       =   2
               Mask            =   "####0.00"
               MaxLength       =   8
               DataType        =   2
               ReCalcParent    =   "WINLMSSAL"
               ReCalcOn        =   "WINREFRESHDT"
               IdName          =   "WINLMSCST"
            End
            Begin MwfCtl.MWCTL_MED atxt 
               Height          =   285
               Index           =   245
               Left            =   10440
               TabIndex        =   216
               ToolTipText     =   "Enter Date Of LME Palladium Rate"
               Top             =   270
               Width           =   885
               _ExtentX        =   1561
               _ExtentY        =   503
               MaxLength       =   8
               DataType        =   3
               ReCalcOn        =   "WINREFRESHDT"
               IdName          =   "WINLMLDT"
            End
            Begin MwfCtl.MWCTL_MED atxt 
               Height          =   285
               Index           =   246
               Left            =   10440
               TabIndex        =   217
               ToolTipText     =   "Enter LME Palladium Cost Value"
               Top             =   555
               Width           =   885
               _ExtentX        =   1561
               _ExtentY        =   503
               Alignment       =   2
               Mask            =   "####0.00"
               MaxLength       =   8
               DataType        =   2
               ReCalcParent    =   "WINLMLSAL"
               ReCalcOn        =   "WINREFRESHDT"
               IdName          =   "WINLMLCST"
            End
            Begin MwfCtl.MWCTL_MED atxt 
               Height          =   285
               Index           =   247
               Left            =   10440
               TabIndex        =   218
               ToolTipText     =   "Enter LME Palladium Sales Value"
               Top             =   840
               Width           =   885
               _ExtentX        =   1561
               _ExtentY        =   503
               Alignment       =   2
               Mask            =   "####0.00"
               MaxLength       =   8
               DataType        =   2
               ReCalcOn        =   "WINLMLCST"
               IdName          =   "WINLMLSAL"
            End
            Begin MwfCtl.MWCTL_MED atxt 
               Height          =   285
               Index           =   101
               Left            =   6210
               TabIndex        =   211
               ToolTipText     =   "Enter LME Platinum Cost Value"
               Top             =   555
               Width           =   885
               _ExtentX        =   1561
               _ExtentY        =   503
               Alignment       =   2
               Mask            =   "####0.00"
               MaxLength       =   8
               DataType        =   2
               ReCalcParent    =   "WINLMPSAL"
               ReCalcOn        =   "WINREFRESHDT"
               IdName          =   "WINLMPCST"
            End
            Begin MwfCtl.MWCTL_MED atxt 
               Height          =   285
               Index           =   84
               Left            =   1995
               TabIndex        =   206
               ToolTipText     =   "Enter Date for Refreshing Rates"
               Top             =   840
               Width           =   945
               _ExtentX        =   1667
               _ExtentY        =   503
               Mask            =   "###0.0000"
               MaxLength       =   9
               DataType        =   2
               ReCalcParent    =   "WINLMGDT,WINLMPDT,WINLMSDT,WINLMLDT,WINLMLCST,WINLMSCST,WINLMPCST,WINLMGCST"
               ReCalcOn        =   "WINREFRESHDT"
               IdName          =   "WINCNVRT"
            End
            Begin MwfCtl.MWCTL_MED atxt 
               Height          =   285
               Index           =   85
               Left            =   1560
               TabIndex        =   205
               ToolTipText     =   "Enter Whether to Refresh the Conversion Rate"
               Top             =   840
               Width           =   345
               _ExtentX        =   609
               _ExtentY        =   503
               DataType        =   4
               ReCalcParent    =   "WINLMGDT,WINLMPDT,WINLMSDT,WINLMLDT,WINLMLCST,WINLMSCST,WINLMPCST,WINLMGCST"
               IdName          =   "WREFCNVRTYN"
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
               Height          =   255
               Index           =   124
               Left            =   60
               TabIndex        =   241
               Top             =   570
               Width           =   945
            End
            Begin VB.Label ALBL 
               BackStyle       =   0  'Transparent
               Caption         =   "Ref YN/Cnv Rt."
               BeginProperty Font 
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
               Left            =   90
               TabIndex        =   240
               Top             =   840
               Width           =   1425
            End
            Begin VB.Label ALBL 
               BackStyle       =   0  'Transparent
               Caption         =   "LMG Cost"
               BeginProperty Font 
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
               Left            =   3000
               TabIndex        =   232
               Top             =   555
               Width           =   1125
            End
            Begin VB.Label ALBL 
               BackStyle       =   0  'Transparent
               Caption         =   "LMG Sales"
               BeginProperty Font 
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
               Left            =   3000
               TabIndex        =   231
               Top             =   840
               Width           =   1125
            End
            Begin VB.Label ALBL 
               BackStyle       =   0  'Transparent
               Caption         =   "LMP Cost"
               BeginProperty Font 
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
               Left            =   5100
               TabIndex        =   230
               Top             =   555
               Width           =   1125
            End
            Begin VB.Label ALBL 
               BackStyle       =   0  'Transparent
               Caption         =   "LMP Sales"
               BeginProperty Font 
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
               Left            =   5100
               TabIndex        =   229
               Top             =   840
               Width           =   1125
            End
            Begin VB.Label ALBL 
               BackStyle       =   0  'Transparent
               Caption         =   "Refresh Date"
               BeginProperty Font 
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
               Left            =   60
               TabIndex        =   228
               Top             =   270
               Width           =   1305
            End
            Begin VB.Label ALBL 
               BackStyle       =   0  'Transparent
               Caption         =   "LMP Date"
               BeginProperty Font 
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
               Left            =   5100
               TabIndex        =   227
               Top             =   270
               Width           =   1125
            End
            Begin VB.Label ALBL 
               BackStyle       =   0  'Transparent
               Caption         =   "LMG Date"
               BeginProperty Font 
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
               Left            =   3000
               TabIndex        =   226
               Top             =   270
               Width           =   1125
            End
            Begin VB.Label LblRefRtLmsDt 
               BackStyle       =   0  'Transparent
               Caption         =   "LMS Date"
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   285
               Left            =   7200
               TabIndex        =   225
               Top             =   300
               Width           =   1125
            End
            Begin VB.Label LblRefRtLmsSal 
               BackStyle       =   0  'Transparent
               Caption         =   "LMS Sales"
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   285
               Left            =   7200
               TabIndex        =   224
               Top             =   840
               Width           =   1125
            End
            Begin VB.Label LblRefRtLmsCst 
               BackStyle       =   0  'Transparent
               Caption         =   "LMS Cost"
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   285
               Left            =   7200
               TabIndex        =   223
               Top             =   555
               Width           =   1125
            End
            Begin VB.Label LblRefRtLmlCst 
               BackStyle       =   0  'Transparent
               Caption         =   "LML Cost"
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   285
               Left            =   9330
               TabIndex        =   222
               Top             =   555
               Width           =   1125
            End
            Begin VB.Label LblRefRtLmlSal 
               BackStyle       =   0  'Transparent
               Caption         =   "LML Sales"
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   285
               Left            =   9330
               TabIndex        =   221
               Top             =   840
               Width           =   1125
            End
            Begin VB.Label LblRefRtLmlDt 
               BackStyle       =   0  'Transparent
               Caption         =   "LML Date"
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   285
               Left            =   9330
               TabIndex        =   219
               Top             =   270
               Width           =   1125
            End
         End
         Begin MwfCtl.MWCTL_BTN1 CmdRefRtGo 
            Height          =   495
            Left            =   11640
            TabIndex        =   239
            ToolTipText     =   "Click 'Go' to Proceed"
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
         Begin VB.Label LblMain 
            BackStyle       =   0  'Transparent
            BorderStyle     =   1  'Fixed Single
            Caption         =   "  Rate Refresh"
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
            TabIndex        =   51
            Top             =   90
            Width           =   13560
         End
      End
      Begin VB.Frame FraDC 
         Height          =   1425
         Left            =   5010
         TabIndex        =   38
         Top             =   7845
         Width           =   5970
         Begin MwfCtl.MWCTL_MED atxt 
            Height          =   285
            Index           =   57
            Left            =   1935
            TabIndex        =   39
            ToolTipText     =   "Enter Sales Order Tc"
            Top             =   420
            Width           =   570
            _ExtentX        =   1005
            _ExtentY        =   503
            MaxLength       =   3
            DataType        =   4
            ReCalcParent    =   "WDCODCHRFR,WDCODTCTO"
            IdName          =   "WDCODTCFR"
            CmpStr          =   "OdTc>= "
         End
         Begin MwfCtl.MWCTL_MED atxt 
            Height          =   285
            Index           =   58
            Left            =   2610
            TabIndex        =   40
            ToolTipText     =   "Enter Sales Order Year"
            Top             =   420
            Width           =   330
            _ExtentX        =   582
            _ExtentY        =   503
            MaxLength       =   2
            DataType        =   4
            ReCalcParent    =   "WDCODYYTO"
            IdName          =   "WDCODYYFR"
            CmpStr          =   "OdYy>= "
         End
         Begin MwfCtl.MWCTL_MED atxt 
            Height          =   285
            Index           =   59
            Left            =   3030
            TabIndex        =   41
            ToolTipText     =   "Enter Sales Order Character"
            Top             =   420
            Width           =   570
            _ExtentX        =   1005
            _ExtentY        =   503
            MaxLength       =   3
            DataType        =   4
            ReCalcParent    =   "WDCODCHRTO"
            ReCalcOn        =   "WDCODTCFR"
            IdName          =   "WDCODCHRFR"
            CmpStr          =   "OdChr>= "
         End
         Begin MwfCtl.MWCTL_MED atxt 
            Height          =   285
            Index           =   60
            Left            =   3690
            TabIndex        =   42
            ToolTipText     =   "Enter Sales Order No."
            Top             =   420
            Width           =   750
            _ExtentX        =   1323
            _ExtentY        =   503
            Alignment       =   2
            Mask            =   "#####0"
            MaxLength       =   6
            DataType        =   1
            ReCalcParent    =   "WDCODNOTO"
            IdName          =   "WDCODNOFR"
            CmpStr          =   "OdNo>= "
         End
         Begin MwfCtl.MWCTL_MED atxt 
            Height          =   285
            Index           =   4
            Left            =   4530
            TabIndex        =   43
            ToolTipText     =   "Enter Sales Order Sr."
            Top             =   420
            Width           =   555
            _ExtentX        =   979
            _ExtentY        =   503
            Alignment       =   2
            Mask            =   "###0"
            MaxLength       =   4
            DataType        =   1
            ReCalcParent    =   "WDCODSRTO"
            IdName          =   "WDCODSRFR"
            CmpStr          =   "OdSr>= "
         End
         Begin MwfCtl.MWCTL_MED atxt 
            Height          =   285
            Index           =   37
            Left            =   1935
            TabIndex        =   44
            ToolTipText     =   "Enter Sales Order Tc"
            Top             =   705
            Width           =   570
            _ExtentX        =   1005
            _ExtentY        =   503
            MaxLength       =   3
            DataType        =   4
            ReCalcOn        =   "WDCODTCFR"
            IdName          =   "WDCODTCTO"
            CmpStr          =   "OdTc <= "
         End
         Begin MwfCtl.MWCTL_MED atxt 
            Height          =   285
            Index           =   38
            Left            =   2610
            TabIndex        =   45
            ToolTipText     =   "Enter Sales Order Year"
            Top             =   705
            Width           =   330
            _ExtentX        =   582
            _ExtentY        =   503
            MaxLength       =   2
            DataType        =   4
            ReCalcOn        =   "WDCODYYFR"
            IdName          =   "WDCODYYTO"
            CmpStr          =   "OdYy <= "
         End
         Begin MwfCtl.MWCTL_MED atxt 
            Height          =   285
            Index           =   39
            Left            =   3030
            TabIndex        =   46
            ToolTipText     =   "Enter Sales Order Character"
            Top             =   705
            Width           =   570
            _ExtentX        =   1005
            _ExtentY        =   503
            MaxLength       =   3
            DataType        =   4
            ReCalcOn        =   "WDCODCHRFR"
            IdName          =   "WDCODCHRTO"
            CmpStr          =   "OdChr <= "
         End
         Begin MwfCtl.MWCTL_MED atxt 
            Height          =   285
            Index           =   46
            Left            =   3690
            TabIndex        =   47
            ToolTipText     =   "Enter Sales Order No."
            Top             =   705
            Width           =   750
            _ExtentX        =   1323
            _ExtentY        =   503
            Alignment       =   2
            Mask            =   "#####0"
            MaxLength       =   6
            DataType        =   1
            ReCalcOn        =   "WDCODNOFR"
            IdName          =   "WDCODNOTO"
            CmpStr          =   "OdNo <= "
         End
         Begin MwfCtl.MWCTL_MED atxt 
            Height          =   285
            Index           =   47
            Left            =   4530
            TabIndex        =   48
            ToolTipText     =   "Enter Sales Order Sr."
            Top             =   705
            Width           =   555
            _ExtentX        =   979
            _ExtentY        =   503
            Alignment       =   2
            Mask            =   "###0"
            MaxLength       =   4
            DataType        =   1
            ReCalcOn        =   "WDCODSRFR"
            IdName          =   "WDCODSRTO"
            CmpStr          =   "OdSr <= "
         End
         Begin MwfCtl.MWCTL_MED atxt 
            Height          =   285
            Index           =   48
            Left            =   1935
            TabIndex        =   50
            ToolTipText     =   "Specify Whether To Wait For All The Bags Before Making An Invoice"
            Top             =   990
            Visible         =   0   'False
            Width           =   330
            _ExtentX        =   582
            _ExtentY        =   503
            DataType        =   4
            IdName          =   "WWAITFORALLBAGS"
         End
         Begin MwfCtl.MWCTL_BTN1 CmdDCGo 
            Height          =   465
            Left            =   5280
            TabIndex        =   49
            ToolTipText     =   "Click Here To Copy"
            Top             =   840
            Width           =   555
            _ExtentX        =   979
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
         Begin VB.Label LblMain 
            BackStyle       =   0  'Transparent
            BorderStyle     =   1  'Fixed Single
            Caption         =   "  Design Copy From"
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
            TabIndex        =   170
            Top             =   90
            Width           =   6030
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
            Index           =   27
            Left            =   2520
            TabIndex        =   20
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
            Index           =   26
            Left            =   2955
            TabIndex        =   21
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
            Index           =   25
            Left            =   3615
            TabIndex        =   148
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
            Index           =   22
            Left            =   4440
            TabIndex        =   147
            Top             =   705
            Width           =   105
         End
         Begin VB.Label ALBL 
            BackStyle       =   0  'Transparent
            Caption         =   "Wait For All Bags"
            BeginProperty Font 
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
            Left            =   90
            TabIndex        =   146
            Top             =   990
            Visible         =   0   'False
            Width           =   1875
         End
         Begin VB.Label ALBL 
            BackStyle       =   0  'Transparent
            Caption         =   "                     To"
            BeginProperty Font 
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
            Left            =   90
            TabIndex        =   145
            Top             =   705
            Width           =   1575
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
            Index           =   4
            Left            =   4440
            TabIndex        =   114
            Top             =   420
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
            Index           =   54
            Left            =   3615
            TabIndex        =   103
            Top             =   420
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
            Left            =   2955
            TabIndex        =   102
            Top             =   420
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
            Left            =   2520
            TabIndex        =   101
            Top             =   420
            Width           =   105
         End
         Begin VB.Label ALBL 
            BackStyle       =   0  'Transparent
            Caption         =   "SO Voucher Fr"
            BeginProperty Font 
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
            TabIndex        =   100
            Top             =   420
            Width           =   1545
         End
      End
      Begin VB.Frame FraSalHst 
         Height          =   4665
         Left            =   150
         TabIndex        =   165
         Top             =   4545
         Width           =   15120
         Begin MwfCtl.MWCTL_FLX GrdHR 
            Height          =   2250
            Left            =   120
            TabIndex        =   167
            Top             =   2370
            Width           =   9285
            _ExtentX        =   16378
            _ExtentY        =   3969
            Cols            =   14
            AllowAdd        =   0   'False
            AllowDelete     =   0   'False
            colname1        =   "WHIRRMCD"
            heading1        =   "Rm Cd                                     "
            datatype1       =   4
            maxlength1      =   16
            colname2        =   "WHIRLOTNO"
            heading2        =   "Lot No.                   "
            datatype2       =   4
            maxlength2      =   12
            colname3        =   "WHIRRMSZ"
            heading3        =   "Size       "
            datatype3       =   2
            mask3           =   "#0.0000"
            maxlength3      =   7
            colname4        =   "WHIRRMSZ2"
            heading4        =   "Size2      "
            datatype4       =   2
            mask4           =   "#0.0000"
            maxlength4      =   7
            colname5        =   "WHIRRMSZ3"
            heading5        =   "Size3      "
            datatype5       =   2
            mask5           =   "#0.0000"
            maxlength5      =   7
            colname6        =   "WHIRQTY"
            heading6        =   "Qty   "
            datatype6       =   1
            mask6           =   "####0"
            maxlength6      =   5
            colname7        =   "WHIRRMWT"
            heading7        =   "Weight    "
            datatype7       =   2
            mask7           =   "###0.000"
            maxlength7      =   8
            colname8        =   "WHIRRMRT"
            heading8        =   "Rate                "
            datatype8       =   2
            mask8           =   "#######0.000"
            maxlength8      =   12
            colname9        =   "WHIRRMVAL"
            heading9        =   "Value               "
            datatype9       =   2
            mask9           =   "########0.00"
            maxlength9      =   12
            colname10       =   "WHIRSETSCD"
            heading10       =   "Set Cd          "
            datatype10      =   4
            maxlength10     =   8
            colname11       =   "WHIRSETRT"
            heading11       =   "Set Rt        "
            datatype11      =   2
            mask11          =   "####0.000"
            maxlength11     =   9
            colname12       =   "WHIRSETVAL"
            heading12       =   "Set Val            "
            datatype12      =   2
            mask12          =   "########0.00"
            maxlength12     =   12
            colname13       =   "WHIRMAINMET"
            heading13       =   "Main Met"
            datatype13      =   4
            maxlength13     =   1
            style13         =   0
         End
         Begin MwfCtl.MWCTL_FLX GrdHL 
            Height          =   2250
            Left            =   9390
            TabIndex        =   168
            Top             =   2370
            Width           =   5700
            _ExtentX        =   10054
            _ExtentY        =   3969
            Cols            =   7
            AllowAdd        =   0   'False
            AllowDelete     =   0   'False
            colname1        =   "WHILLMCD"
            heading1        =   "Main Cd      "
            datatype1       =   4
            maxlength1      =   8
            colname2        =   "WHILLSCD"
            heading2        =   "Sub Code  "
            datatype2       =   4
            maxlength2      =   8
            colname3        =   "WHILQW"
            heading3        =   "Q/W"
            datatype3       =   4
            maxlength3      =   1
            colname4        =   "WHILQTY"
            heading4        =   "Qty   "
            datatype4       =   1
            mask4           =   "###0"
            maxlength4      =   4
            colname5        =   "WHILRATE"
            heading5        =   "Rate         "
            datatype5       =   2
            mask5           =   "#####0.00"
            maxlength5      =   9
            colname6        =   "WHILIVAL"
            heading6        =   "Value                "
            datatype6       =   2
            mask6           =   "########0.00"
            maxlength6      =   12
         End
         Begin MwfCtl.MWCTL_FLX GrdHD 
            Height          =   2040
            Left            =   120
            TabIndex        =   166
            Top             =   360
            Width           =   14445
            _ExtentX        =   25479
            _ExtentY        =   3598
            Cols            =   16
            AllowAdd        =   0   'False
            AllowDelete     =   0   'False
            colname1        =   "WHDSR"
            heading1        =   "Sr"
            datatype1       =   1
            mask1           =   "0"
            maxlength1      =   1
            colname2        =   "WHIDTC"
            heading2        =   "IdTc"
            datatype2       =   4
            maxlength2      =   2
            colname3        =   "WHIDYY"
            heading3        =   "IdYy"
            datatype3       =   4
            maxlength3      =   2
            colname4        =   "WHIDCHR"
            heading4        =   "IdChr  "
            maxlength4      =   3
            colname5        =   "WHIDNO"
            heading5        =   "IdNo       "
            datatype5       =   1
            mask5           =   "#####0"
            maxlength5      =   6
            colname6        =   "WHIDSR"
            heading6        =   "IdSr    "
            datatype6       =   1
            mask6           =   "#####0"
            maxlength6      =   6
            colname7        =   "WHIDDT"
            heading7        =   "IdDt          "
            datatype7       =   3
            maxlength7      =   8
            colname8        =   "WHDMCD"
            heading8        =   "Design Cd                     "
            datatype8       =   4
            maxlength8      =   15
            colname9        =   "WHDMSFX"
            heading9        =   "Suffix               "
            datatype9       =   4
            maxlength9      =   10
            colname10       =   "WHDMSZ"
            heading10       =   "Size     "
            datatype10      =   4
            maxlength10     =   5
            colname11       =   "WHIDKT"
            heading11       =   "Karat  "
            datatype11      =   4
            maxlength11     =   5
            style11         =   0
            colname12       =   "WHIDCOL"
            heading12       =   "Color     "
            datatype12      =   4
            maxlength12     =   5
            colname13       =   "WHIDQTY"
            heading13       =   "Qty         "
            datatype13      =   2
            mask13          =   "####0.0"
            maxlength13     =   7
            colname14       =   "WHIDRT"
            heading14       =   "Rate                 "
            datatype14      =   2
            mask14          =   "########0.00"
            maxlength14     =   12
            colname15       =   "WHIDVAL"
            heading15       =   "Value                "
            datatype15      =   2
            mask15          =   "########0.00"
            maxlength15     =   12
            style15         =   0
         End
         Begin VB.Label LblMain 
            BackStyle       =   0  'Transparent
            BorderStyle     =   1  'Fixed Single
            Caption         =   "  Design Sale History"
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
            TabIndex        =   172
            Top             =   90
            Width           =   15165
         End
      End
   End
   Begin VB.Frame FraGetFOB 
      Height          =   1530
      Left            =   3240
      TabIndex        =   256
      Top             =   7680
      Visible         =   0   'False
      Width           =   2430
      Begin MwfCtl.MWCTL_BTN1 CmdGetFOBGo 
         Height          =   495
         Left            =   780
         TabIndex        =   257
         ToolTipText     =   "Click 'Go' to Proceed"
         Top             =   720
         Width           =   615
         _ExtentX        =   1085
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
      Begin VB.Label LblMain 
         BackStyle       =   0  'Transparent
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Get FOB"
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
         Left            =   60
         TabIndex        =   258
         Top             =   90
         Width           =   2250
      End
   End
   Begin VB.Frame FraGetLotRt 
      Height          =   1650
      Left            =   4650
      TabIndex        =   246
      Top             =   7590
      Visible         =   0   'False
      Width           =   6870
      Begin VB.Frame Frame2 
         Height          =   1215
         Left            =   180
         TabIndex        =   247
         Top             =   360
         Width           =   5865
         Begin VB.TextBox txtGetLotRtSEOPwd 
            Height          =   285
            IMEMode         =   3  'DISABLE
            Left            =   2850
            MaxLength       =   30
            PasswordChar    =   "*"
            TabIndex        =   250
            ToolTipText     =   "Enter SEO Password"
            Top             =   540
            Width           =   1095
         End
         Begin MwfCtl.MWCTL_MED atxt 
            Height          =   285
            Index           =   108
            Left            =   2850
            TabIndex        =   248
            ToolTipText     =   "Enter ""Y"" for Only Customer “N” for ALL"
            Top             =   210
            Width           =   345
            _ExtentX        =   609
            _ExtentY        =   503
            DataType        =   4
            ReCalcParent    =   "WINLMGDT,WINLMPDT,WINLMSDT,WINLMLDT,WINLMLCST,WINLMSCST,WINLMPCST,WINLMGCST"
            IdName          =   "WGETLOTRTONLYCUST"
         End
         Begin MwfCtl.MWCTL_MED atxt 
            Height          =   285
            Index           =   87
            Left            =   1110
            TabIndex        =   249
            ToolTipText     =   "Enter Y to replace all DCX"
            Top             =   780
            Visible         =   0   'False
            Width           =   345
            _ExtentX        =   609
            _ExtentY        =   503
            DataType        =   4
            ReCalcParent    =   "WINLMGDT,WINLMPDT,WINLMSDT,WINLMLDT,WINLMLCST,WINLMSCST,WINLMPCST,WINLMGCST"
            IdName          =   "WGETLOTRTALLDCX"
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
            Height          =   255
            Index           =   51
            Left            =   1890
            TabIndex        =   254
            Top             =   570
            Width           =   945
         End
         Begin VB.Label ALBL 
            BackStyle       =   0  'Transparent
            Caption         =   "Only Customer [Y] / ALL [N]"
            BeginProperty Font 
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
            Left            =   90
            TabIndex        =   253
            ToolTipText     =   "Enter ""Y"" for Only Customer “N” for ALL"
            Top             =   240
            Width           =   2715
         End
         Begin VB.Label ALBL 
            BackStyle       =   0  'Transparent
            Caption         =   "All DCX"
            BeginProperty Font 
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
            Left            =   90
            TabIndex        =   251
            Top             =   690
            Visible         =   0   'False
            Width           =   705
         End
      End
      Begin MwfCtl.MWCTL_BTN1 cmdGetLotRtGo 
         Height          =   495
         Left            =   6060
         TabIndex        =   252
         ToolTipText     =   "Click 'Go' to Proceed"
         Top             =   1080
         Width           =   615
         _ExtentX        =   1085
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
      Begin VB.Label LblMain 
         BackStyle       =   0  'Transparent
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Get Lot Rate"
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
         Left            =   180
         TabIndex        =   255
         Top             =   90
         Width           =   5850
      End
   End
End
Attribute VB_Name = "EmrFrmInvDsg"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'----- General comments on the Entry  ----
' X-'', A-InvDsg(GrdID), B-InvRm(GrdIR), C-InvLab(GrdIL), 3rd Child-InvDsg(GrdIB)
'
' CmdMI - Make Invoice Option (This option generates all the design, Rm & Lab level records
'         for the entire invoice at one shot if all the invoice bags have been entered through
'         the 'Invoice All Bags' Entry form.
'
' CmdCRL - Copy Rm & Lab Details for only the Bags in the Inv Sr
'
' Design Copy option is used to copy the Design, Bag, Rm & Lab details of a design
' FraDC, CmdDC, CmdDCGo, wDcOdTc, wDcOdYy, wDcOdChr, wDcOdNo, wDcOdSr
'        all part of Design Copy routine.
'
' Labour Details are shown on a frame as it could not be fitted on the main frame because of space constraints
' FraIL, GrdIL
'
' Bag Details are shown on a frame as it could not be fitted on the main frame because of space constraints
' FraIB, GrdIB
'
' Design Level fields which could not be fitted in the grid GrdID are shown on a seperate
' frame which could be seen on the click of a button(CmdDD)
' FraDD, CmdDD
'
' Summary option summarises the Design, Rm and Lab  details of the entire Invoice (5 Tabs)
' FraSumm, FraTabSumm, CmdSumm,
' GrdSummDR, wDrGrp, wDrTotQty, wDrTotiWt, wDrTotiVal, wDrTotZWt, wDrTotZVal
' GrdSummDL, wDlGrp, wDlTotiVal, wDlTotZVal
' GrdSummID, wIdGrp, wIdTotQty, wIdTotiFOB, wIdTotZFOB
' GrdSummIR, wIrGrp, wIrTotQty, wIrTotiWt, wIrTotiVal, wIrTotZWt, wIrTotZVal
' GrdSummIL, wIlGrp, wIlTotiVal, wIlTotZVal
' all part of Summary routine.
'
'-----
Option Explicit
Dim moCn As MwfLib.MDOConnection   '*** (Bef speed) bef 24/11/06
Dim ms_IdTcTyp As String, ms_IdYy As String, ms_IdTc As String, ms_IdChr As String
Dim ms_hSizeStkRt As String
  
Dim ms_CmCtg As String
Dim ms_RtbyStkYN As String '3.11.0
Dim mf_inspecialCopy As Boolean
Dim mf_StopRecalcDuringCopyIR As Boolean
Dim mf_StopRecalcDuringCopyIL As Boolean
Dim ms_CustRm As Boolean, mb_CHCDHCOnCustRm As Boolean
Enum en_InvDsgFra
  MakeInv = 0
  DsgDet = 1
  DsgCpy = 2
  InvLab = 3
  InvBag = 4
  Summ = 5
  CpyRmLab = 6
  'Manoj 2.10.0
  SalHst = 7
  'Manoj 2.10.0
  DelRmLab = 8    '****** Sachin 2.13.0 - 27-05-2006 ******
  RefRt = 9       ' ***** Manali 3.6.0 - 29/10/09 - Refresh Metal Rate
  DelInv = 10       ' 4.1.2.0
GetLotRt = 11       ' ***** Manali 3.6.0 - 29/10/09 - Refresh Metal Rate
  GetFOB = 12   '18 Get FOB added
  Scope = 13  '6.1_37
End Enum

'*** Arr_Cv stores the causes validation state, Arr_Ed the Ena/Disa state of Cmd Buttons
Dim mArr_Cv As Variant, mArr_Ed As Variant

Dim wPic3DHt As Double
Dim wPic3DWdt As Double
Dim wPicHskHt As Double
Dim wPicHskWdt As Double
Dim wFraPic3DLft As Double
Dim wFraPic3DTop As Double
Dim wFraPicHskLft As Double
Dim wFraPicHskTop As Double
Dim wPic3DDblClk As Boolean
Dim wPicHskDblClk As Boolean

Dim wFraPic3DHt As Single
Dim wFraPic3DWdt As Single
Dim wFraPicHskHt As Single
Dim wFraPicHskWdt As Single

'*** Declare the variables ms_LockDt, mf_DtLocked for the Lock Date Check
Dim ms_LockDt As String
Dim ms_FutureDt As String
Dim mf_DtLocked As Boolean

Dim ms_PurgeDt As String        '****** Sachin 3.02

'*** Declare the variables ms_RndOffFOB, ms_ShowActWt to store the fields HRndOffFOB and HShowActWt of Head file
Dim ms_RndOffFOB As String
Dim ms_ShowActWt As String

Dim mInCls As String
Dim mIdHZValEqInv As String
Dim mLocalExp As String, mSlvMod As Boolean, mf_QWModYn As Boolean, mf_AddBrkMsgYN As Boolean
 
Const ctDsgSumm = "D": Const ctInvSumm = "I"
Const ctAddLMCd = "ZADD": Const ctAddLSCd = "-"

''Const ctBrkLMCd = "ZBRK": Const ctBrkLSCd = "-"
'*** To decide if the picture for the design should be shown or not
Dim mf_ShowDsg As Boolean

Dim ms_OrdTyp As String, ms_DespTyp As String
Dim md_InCnvRt As Double

Dim ms_CurrentBtn As String   '*** Jenny Color

' **** Manali 3.03 - 21/06/08 - Pd Mod
Dim mb_PdMod As Boolean
' **** Manali 3.03 - 21/06/08 - Pd Mod

'*** (Jen 3.01.0_1)
Dim mb_InInvGrsWtEqActual As Boolean
Dim mb_DsgOnHld As Boolean
Dim ms_RemakeOpt As String
'*** (Jen 3.01.0_1)

'****** Sachin 3.3.0 (23-06-08) - [42. Make Invoice By Bag] ******
Dim mb_InByBag As Boolean
'****** Sachin 3.3.0 (23-06-08) - [42. Make Invoice By Bag] ******

Dim ms_PwdFlag As Boolean   ' **** Manali 3.6.0 - 19/11/09 - Pwd for Refresh Rate

Dim ms_RmRtfrRtChrt As String, ms_LabRtfrRtChrt As String     'Sangeeta 3.11.0
Dim ms_RtFrInvCust As String
Dim ms_InOrdMulbyYN As String
Dim mf_TS_MakeInvoice As Boolean
Dim wStkNoDisp As String
Dim mb_SeoUsr As Boolean, ms_hInvPtr As String

Private Sub TrackStatus(MessageToAdd As String)
 If mf_TS_MakeInvoice Then _
      mFrmTrackStatus.txtStatus.text = mFrmTrackStatus.txtStatus.text + vbCrLf + MessageToAdd + "::" + CStr(Now)
End Sub

Private Sub ADC_Load()
  '*** Redim the arrays mArr_Cv and mArr_Ed to the number of command buttons
  '*** Set the First Key Control property to the field wIdChr
  '*** Set the First Non Key Control property to adc
  '*** Set the Child Property Of ADC to GrdID
  '*** Set the Child Property Of GrdID to GrdIR, GrdIL and GrdIB
  '*** Set the Hot Keys for the Grids GrdID, GrdIR, GrdIL, GrdIB as 'A', 'B', 'C', 'D' respectively
  '*** Set the Previous and Next Control for the Grids GrdID, GrdIR, GrdIL, GrdIB,
      'GrdSummDR, GrdSummDL, GrdSummID, GrdSummIR, GrdSummIL
  '*** Store the defa values for Inv Tc,Yy in variables
  '*** Bring Frame FraNKeyAll to the top
  '*** Store the original Size and Position of the Picture Controls ***
  '*** This is necessary to bring back the picture size to the original control size
      'from the enlarged state when the user toggels the double click option of the picture control.
  '*** Store the initial Double Click Option of the Picture Controls as False

' Manoj 2.10.0
'  ReDim mArr_Cv(7), mArr_Ed(7)
  'ReDim mArr_Cv(9), mArr_Ed(9)
  ReDim mArr_Cv(13), mArr_Ed(13)      ' ***** Manali 3.6.0 - 29/10/09 - Refresh Metal Rate
' Manoj 2.10.0
  
  Set moCn = ADC.Connection    '*** (Bef speed) bef 24/11/06
  Set ADC.FirKeyCtl = ADC("wIdChr")
  Set ADC.FirNKeyCtl = ADC
  ADC.Child = "GrdID"
  GrdID.Child = "GrdIR,GrdIL,GrdIB"

  GrdID.HotKey = "A"
  GrdIR.HotKey = "B"
  GrdIL.HotKey = "C"
  GrdIB.HotKey = "D"
  
  Set GrdID.PrevCtl = ADC
  Set GrdID.NextCtl = ADC
  Set GrdIR.PrevCtl = GrdID
  Set GrdIR.NextCtl = GrdID
  Set GrdIL.PrevCtl = CmdIL
  Set GrdIL.NextCtl = CmdIL
  Set GrdIB.PrevCtl = CmdIB
  Set GrdIB.NextCtl = CmdIB
  
  Set GrdSummDR.NextCtl = ADC("wDRGrp")
  Set GrdSummDL.NextCtl = ADC("wDLGrp")
  Set GrdSummID.NextCtl = ADC("wIDGrp")
  Set GrdSummIR.NextCtl = ADC("wIRGrp")
  Set GrdSummIL.NextCtl = ADC("wILGrp")
  
  ' Manoj 2.10.0
  Set GrdHD.PrevCtl = CmdSH
  Set GrdHD.NextCtl = GrdHR
  Set GrdHR.PrevCtl = GrdHD
  Set GrdHR.NextCtl = GrdHL
  Set GrdHL.PrevCtl = GrdHR
  Set GrdHL.NextCtl = CmdSH
  ' Manoj 2.10.0
  
  FraNKeyAll.ZOrder
  'ms_IdTcTyp = ctInTcTyp
'      Set m_ObjCmdGetRmRt = ADC.Connection.OpenQuery("ForCpy_GetRmRt")
 '   Set m_ObjCmdGetLabRt = ADC.Connection.OpenQuery("ForCpy_GetLabRt")

  Select Case UCase(ADC.MenuCd)
  Case Is = UCase("InDsg")
    ms_IdTcTyp = ctInTcTyp: ms_CmCtg = "C"
    ADC("wDdIdHZValEqInv").Enabled = True: ADC("wDdIdHZValEqInv").Visible = True
    ADC("wDdIdFixPrc").Enabled = True: ADC("wDdIdFixPrc").Visible = True
    '****** Zubin 3.02.01 - 27/03/08
    ADC("wDdIdLabWtFrOrd").Visible = True
    '****** Zubin 3.02.01 - 27/03/08
    ADC("wDdIdGldLs").Enabled = True: ADC("wDdIdGldLs").Visible = True
    '*********Geeta***Emr208**
    ADC("wDdIdValAddn").Enabled = True: ADC("wDdIdValAddn").Visible = True
    LblDdIdValAddn.Visible = True
    '**********
    LblDdIdHZValEqInv.Visible = True
    LblDdIdFixPrc.Visible = True
    LblDdIdGldLs.Visible = True
    ms_OrdTyp = "Exp": ms_DespTyp = "FdDespQty"
    
  Case Is = UCase("CTBInDsg")
    ms_IdTcTyp = ctCTBTcTyp: ms_CmCtg = "T"
    ADC("wDdIdHZValEqInv").Enabled = False: ADC("wDdIdHZValEqInv").Visible = False
    ADC("wDdIdFixPrc").Enabled = False: ADC("wDdIdFixPrc").Visible = False
    '****** Zubin 3.02.01 - 27/03/08
    ADC("wDdIdLabWtFrOrd").Visible = False
    '****** Zubin 3.02.01 - 27/03/08
    ADC("wDdIdGldLs").Enabled = False: ADC("wDdIdGldLs").Visible = False
    '*********Geeta***Emr208**
    ADC("wDdIdValAddn").Enabled = False: ADC("wDdIdValAddn").Visible = False
    LblDdIdValAddn.Visible = False
    '**********
    LblDdIdHZValEqInv.Visible = False
    LblDdIdFixPrc.Visible = False
    LblDdIdGldLs.Visible = False
    ms_OrdTyp = "Prd": ms_DespTyp = "FdCTBDespQty"
    ' Manoj 2.10.0
    CmdSH.Visible = False
    ' Manoj 2.10.0
  End Select
  
  ms_IdTc = moCn.GetFldVal("Select PValue from Param where PTyp= 'TCTYP' and PMCd= '" + ms_IdTcTyp + "' and PSCd= ''")
  mf_QWModYn = IIF(moCn.GetFldVal("Select hQWModYn From Head Where HCd='" + ctSelfCmCd + "' And HCoCd='" + ctSelfCoCd + "'") = "Y", True, False)
  mf_AddBrkMsgYN = IIF(moCn.GetFldVal("Select  hAddBrkMsgYN From Head Where HCd='" + ctSelfCmCd + "' And HCoCd='" + gs_CoCd + "'") = "Y", True, False)
  mf_TS_MakeInvoice = IIF(moCn.GetFldVal("Select 'Y' from Param where PTyp= 'PTYP' and PMCd= 'TS' and PSCd='' and PDesc225 like '%InDsg%'") = "Y", True, False)
  If moCn.GetFldVal("Select hCHCDHCOnCustRm from Head where HCoCd= '" + gs_CoCd + "' and HCd='" + ctSelfCmCd + "'") = "Y" Then
    mb_CHCDHCOnCustRm = True
  Else
    mb_CHCDHCOnCustRm = False
  End If
  
  '*** (Bef 3.01.0_1)
  'ms_IdYy = moCn.GetFldVal("Select hYy from Head where HCoCd= '" + gs_CoCd + "' and HCd= '" + ctSelfCmCd + "'")
  'mIdHZValEqInv = moCn.GetFldVal("Select hZValEqInv from Head where HCoCd= '" + gs_CoCd + "' and HCd= '" + ctSelfCmCd + "'")
  'mLocalExp = moCn.GetFldVal("Select hExpLocal from Head where HCoCd= '" + gs_CoCd + "' and HCd= '" + ctSelfCmCd + "' ")
  'mSlvMod = IIF(moCn.GetFldVal("Select HSlvModYN From Head Where HCoCd='" + gs_CoCd + "' And HCd='" + ctSelfCmCd + "'") = "Y", True, False)
  '*** (Bef 3.01.0_1)
  
  '*** (Jen 3.01.0_1)
  Dim wRs_Hd As MwfLib.MDORowSet
  ' **** Manali 3.03 - 21/06/08 - Pd Mod
  Set wRs_Hd = moCn.OpenRes("Select hYy, hZValEqInv, hExpLocal, HSlvModYN, HPdModYN from Head " + _
               "where HCoCd= '" + gs_CoCd + "' and HCd= '" + ctSelfCmCd + "'")
  If Not (wRs_Hd.EOF Or wRs_Hd.BOF) Then
    ms_IdYy = wRs_Hd!hYy
    mIdHZValEqInv = wRs_Hd!hZValEqInv
    mLocalExp = wRs_Hd!hExpLocal
    mSlvMod = IIF(UCase(wRs_Hd!HSlvModYN) = "Y", True, False)
    mb_PdMod = IIF(UCase(wRs_Hd!HPdModYN) = "Y", True, False)   ' **** Manali 3.03 - 21/06/08 - Pd Mod
  End If
  '*** (Jen 3.01.0_1)
      
   '****Geeta***Emr210
  If UCase$(mLocalExp) = "N" Then
       LblGldFiness.Visible = False: ADC("wIdGldFiness").Visible = False: ADC("wIdGldFiness").Enabled = False
       LblPlFiness.Visible = False: ADC("wIdPlFiness").Visible = False: ADC("wIdPlFiness").Enabled = False
       ' ****** Sachin 2.12 - 06/12/05 - SJM
       LblSlFiness.Visible = False: ADC("wIdSlFiness").Visible = False: ADC("wIdSlFiness").Enabled = False
       ' ****** Sachin 2.12 - 06/12/05 - SJM
       ' **** Manali 3.03 - 21/06/08 - Pd Mod
       LblPdFiness.Visible = False: ADC("wIdPdFiness").Visible = False: ADC("wIdPdFiness").Enabled = False
       ' **** Manali 3.03 - 21/06/08 - Pd Mod
       LblDdIdValAddn.Visible = False: ADC("wDdIdValAddn").Visible = False
  End If

  If mSlvMod = False Then
    LblLmsCst.Visible = False: LblLmsSal.Visible = False: LblSlFiness.Visible = False
    ADC("wIdLMSCst").Visible = False: ADC("wIdLmsSal").Visible = False: ADC("wIdSlFiness").Visible = False
    ' ***** Manali 3.6.0 - 29/11/09 - Refresh Metal Rate
    ADC("wInLMSCst").Visible = False: ADC("wInLmsDt").Visible = False: ADC("wInLmsSal").Visible = False
    ADC("wInLMSCst").Enabled = False: ADC("wInLmsDt").Enabled = False: ADC("wInLmsSal").Enabled = False
    LblRefRtLmsDt.Visible = False: LblRefRtLmsCst.Visible = False: LblRefRtLmsSal.Visible = False
    ' ***** Manali 3.6.0 - 29/11/09 - Refresh Metal Rate
  End If
  
  ' **** Manali 3.03 - 21/06/08 - Pd Mod
  If mb_PdMod = False Then
    LblLmlCst.Visible = False: LblLmlSal.Visible = False: LblPdFiness.Visible = False
    ADC("wIdLmlCst").Visible = False: ADC("wIdLmlSal").Visible = False: ADC("wIdPdFiness").Visible = False
    ' ***** Manali 3.6.0 - 29/11/09 - Refresh Metal Rate
    ADC("wInLMLCst").Visible = False: ADC("wInLmlDt").Visible = False: ADC("wInLmlSal").Visible = False
    ADC("wInLMLCst").Enabled = False: ADC("wInLmlDt").Enabled = False: ADC("wInLmlSal").Enabled = False
    LblRefRtLmlDt.Visible = False: LblRefRtLmlCst.Visible = False: LblRefRtLmlSal.Visible = False
    ' ***** Manali 3.6.0 - 29/11/09 - Refresh Metal Rate
  End If
  ' **** Manali 3.03 - 21/06/08 - Pd Mod

  ' Zubin 212
  If moCn.RecSeek("select 'x' from Head where HCd= '" + ctSelfCmCd + "' and HCoCd= '" + gs_CoCd + "' and HLotNoYN= 'N'") Then
    GrdIR.ColProp("IrLotNo").Style = fgcinvisible
    GrdHR.ColProp("whIrLotNo").Style = fgcinvisible
  End If
  ' Zubin 212

  wPic3DHt = Pic3D.Height
  wPic3DWdt = Pic3D.Width
  wPicHskHt = PicLD.Height
  wPicHskWdt = PicLD.Width
  
  wFraPic3DLft = FraPic3D.Left
  wFraPic3DTop = FraPic3D.Top
  wFraPicHskLft = FraPicLD.Left
  wFraPicHskTop = FraPicLD.Top
  
  wPic3DDblClk = False
  wPicHskDblClk = False

  wFraPic3DHt = FraPic3D.Height
  wFraPic3DWdt = FraPic3D.Width
  wFraPicHskHt = FraPicLD.Height
  wFraPicHskWdt = FraPicLD.Width
  
  '****** Sachin 3.02.0 - Hiding the Partition Key Column based on Partition Selected
  ADC("WIDPRTKEY").Visible = IIF(gs_Partition = ctCurrPrtn, False, True)
  GrdID.ColProp("IdPrtKey").Style = IIF(gs_Partition = ctCurrPrtn, fgcinvisible, fgcReadOnly)
  GrdIR.ColProp("IrPrtKey").Style = IIF(gs_Partition = ctCurrPrtn, fgcinvisible, fgcReadOnly)
  GrdIL.ColProp("IlPrtKey").Style = IIF(gs_Partition = ctCurrPrtn, fgcinvisible, fgcReadOnly)
  GrdIB.ColProp("IfPrtKey").Style = IIF(gs_Partition = ctCurrPrtn, fgcinvisible, fgcReadOnly)
  
  Call SetStockSizeVisible(2, GrdIR.ColProp("IRRMSZ2"), ADC.Connection)
  Call SetStockSizeVisible(3, GrdIR.ColProp("IRRMSZ3"), ADC.Connection)
  If moCn.GetFldVal("Select hAvgSRtYn from Head where HCoCd= '" + gs_CoCd + "'") = "Y" Then
  'StkRt-Getting stock rate display Y\N from head table
  wStkNoDisp = moCn.GetFldVal("Select hStkRtNoDsp from Head where hcocd='" + gs_CoCd + "'")
  'StkRt-when NoDisplay =NO then stock rate is visible,if NoDisplay =yes and AvgStkrate =yes & user donot have
  'SEO rights for the form then stock rate column is Invisible else visible as per current logic.
  If wStkNoDisp = "N" Then
    GrdIR.ColProp("IRSTKRT").Style = fgcNormal
  ElseIf wStkNoDisp = "Y" Then
    mb_SeoUsr = moCn.RecSeek("Select 'x' From UsrAccess where UaUsrCd = '" + gs_UsrCd + "' " + _
                        " and UaMnuCd = '" + ADC.MenuCd + "' and UaSeoYn= 'Y' ")
    If mb_SeoUsr = True Then GrdIR.ColProp("IRSTKRT").Style = fgcNormal
    If mb_SeoUsr = False Then GrdIR.ColProp("IRSTKRT").Style = fgcinvisible
  End If
  End If
  'sv.9 getting purge date from vparam
  ms_PurgeDt = GetPurgeDt(gs_CoCd, "")
  
End Sub
Private Sub ADC_KeyWhen()
  '*** setting the default values for ms_IdTc, ms_IdYy and ms_IdChr from the previous entry
  '*** Disable All command buttons
  '*** Hide all frames except Frame FraNKeyAll
  '*** Make sure that the Picture controls do not show any picture (of some prevoius entry)

  mf_DtLocked = False
  mf_ShowDsg = True
  
  If ms_IdYy <> "" Then ADC("wIdYy") = ms_IdYy
  If ms_IdTc <> "" Then ADC("wIdTc") = ms_IdTc
  '****** do not change the order of initialising
  '****** InChr must occur after InTc as recalc of InChr 'd happen on moving value into Tc
  If ms_IdChr <> "" Then ADC("wIdChr") = ms_IdChr
  Call EnaDisaCmds(True)
  Call HideAllFras
  Call ShowPic(0, 0, "", PicLD)
  Call ShowPic(0, 0, "", Pic3D)

  Me.Caption = GetFrmCaption(gs_CoCd, "Invoice Design Entry")

  FraPic3D.Height = wFraPic3DHt
  FraPic3D.Width = wFraPic3DWdt
  FraPicLD.Height = wFraPicHskHt
  FraPicLD.Width = wFraPicHskWdt
  
  FraPic3D.Left = wFraPic3DLft
  FraPic3D.Top = wFraPic3DTop
  FraPicLD.Left = wFraPicHskLft
  FraPicLD.Top = wFraPicHskTop
  
  If moCn.GetFldVal("Select hSizeStkRt from Head where HCoCd= '" + _
     gs_CoCd + "' and " + "HCd= '" + ctSelfCmCd + "'") = "N" Then
     GrdIR.ColProp("IrRmSz").Style = fgcinvisible
     GrdIR.ColProp("IrRmSz2").Style = fgcinvisible
     GrdIR.ColProp("IrRmSz3").Style = fgcinvisible
     GrdIR.ColProp("IrStkRt").Style = fgcinvisible
     GrdIR.ColProp("IrLotNo").Style = fgcinvisible
     GrdHR.ColProp("whIrLotNo").Style = fgcinvisible
     GrdHR.ColProp("whIrRmSz").Style = fgcinvisible
  End If
  CmdSC.Enabled = True    '6.1_37
End Sub
Private Sub ADC_KeyValid(Cancel As Boolean, ErrMsg As String)
  '*** Checking Key values in Key Valid as they are copied in Init
  '*** Check if the fields wIdYy, wIdChr, wIdNo are blank

  If ADC("wIdYy") = "" Then Cancel = True: ErrMsg = "Enter Invoice Voucher Year": Exit Sub
  If ADC("wIdChr") = "" Then Cancel = True: ErrMsg = "Enter Invoice Voucher Character": Exit Sub
  If ADC("wIdNo") = 0 Then Cancel = True: ErrMsg = "Enter Invoice Voucher Number": Exit Sub
  
  '****** Sachin 3.02 - Corresponding Partition
  Cancel = Not moCn.RecSeek("Select InNo From InvHd Where InCoCd= '" + gs_CoCd + "' and " + _
           "InTc= '" + ADC("wIdTc") + "' And InYy= '" + ADC("wIdYy") + "' and " + _
           "InChr= '" + ADC("wIdChr") + "' and InNo= " + CStr(ADC("wIdNo")) + " and InPrtKey=(Case When '" + gs_Partition + "'='" + ctCurrPrtn + "' Then '" + ctCurrPrtn + "' Else InPrtKey End) " + _
           "")
  If Cancel = True Then ErrMsg = "Invalid Invoice Voucher": Exit Sub
    
End Sub
Private Sub ADC_NonKeyWhen(Cancel As Boolean, ErrMsg As String)
  '*** Store the wIdTc,wIdYy and wIdChr values for the next entry
  '*** Enable All Command Buttons Except Buttons associated with a row in GrdID
  '*** If there are recs in the grid GrdID then disable the Make Invoice Option
  '*** give default values to the working fields wIdDt, wIdCmCd, wIdExpNo, wIdLmgSal,wIdLmgCst,wIdLmpCst,wIdGldFiness,wIdPltFiness from
  '    the InvHd table
  '*** Do not allow the user to edit any field if InCls = 'Y' in InvHd Table
      If ADC.IsNew And (Not moCn.RecSeek("select vPSCd from vParam where " + _
                         "vPCoCd= '" + gs_CoCd + "' and vPTyp= 'CHR' and " + _
                         "vPMCd = '" + ADC("wIdTc") + "' and vPSCd= '" + ADC("wIdChr") + "' and vPValidYn<>'N'")) Then
       Cancel = True: ErrMsg = "Voucher Character No More Valid": ADC.AllowSave = False: Exit Sub
      Else
        ADC.AllowSave = True
      End If
  
  ms_LockDt = moCn.GetFldVal("Select PDesc225 from Param where PTyp= 'LOCKDT' and PMCd= 'LOCKDT' and PSCd= ''")
  ms_FutureDt = moCn.GetFldVal("Select PDesc225 from Param where PTyp= 'LOCKDT' and PMCd= 'FUTUREDT' and PSCd= ''")
  
  '****** Sachin 3.02 31-12-07 - Setting the purge date
  'ms_PurgeDt = moCn.GetFldVal("Select PDesc225 from Param where PTyp= 'PURGEDT' and PMCd= 'PURGEDT' and PSCd= ''")
  
  '*** For the time being do not put Lock Date Check in the Invoice Entry
  '### If ms_LockDt = "" Then ms_LockDt = "01/01/80"
  'If ms_FutureDt = "" Then ms_FutureDt = "01/01/80"
 ' ms_LockDt = "01/01/80"
  'ms_FutureDt = "31/12/28"
'  If ms_PurgeDt = "" Then ms_PurgeDt = "01/01/80"   '****** Sachin 3.02 31-12-07
  
  Call GetLockDt(gs_CoCd, "TC", ms_IdTcTyp, ms_LockDt, ms_FutureDt)
  
  Dim wRsInvDet As MwfLib.MDORowSet
  ms_IdYy = ADC("wIdYy")
  ms_IdTc = ADC("wIdTc")
  ms_IdChr = ADC("wIdChr")
  
  '6.1_37 Display or hide the Scope frame depending on the toggle satate
  If FraSC.Visible = True Then Call DispFra(Scope)
  
  Call EnaDisaCmds(False)
  CmdDD.Enabled = False
  '%%% CmdDC.Enabled = False
  CmdCRL.Enabled = False
  CmdIL.Enabled = False
  CmdIB.Enabled = False
  CmdSumm.Enabled = False
  CmdSH.Enabled = False
  CmdDRL.Enabled = False  '****** Sachin 2.13.0 - 27-05-2006 ******
  CmdInvDel.Enabled = False  '****** Sachin 3.12.0 ******
  CmdSC.Enabled = False   '6.1_37 scope frame will not enable after click of find
  
    
  '****** (Jen 3.01.0_1 added InRemakingCtgs, InInvGrsWtEqActual) ******
  '****** Sachin 3.02 28-12-07 - Partition Key added below
  ' **** Manali 3.03 - 21/06/08 - Pd Mod - LML Cst, Sal, Fineness added
  Set wRsInvDet = moCn.OpenRes(" Select InPrtKey, InDt, InCmCd, InExpNo, InLmgSal, InLmpSal, InLmsSal, InLmgCst, InLmpCst, InLmsCst, " + _
                               " InLmlSal, InLmlCst, " + _
                               " InCnvRt, InHGldFiness, InHPlFiness, InHSlFiness, InHPdFiness, InFOBPer, InCls, InHRndOffFOB, " + _
                               " InCnvFct, InMulDiv, CmCurCd, InDsgOnHld, InRemakingCtgs, InInvGrsWtEqActual, InByBagYN, InRmRtfrRtChrt, InLabRtfrRtChrt,InRtFrInvCust,InOrdMulByYn, InIRtwrtStkYn, InXIRtwrtStkYn, InMetIRtwrtStkYN " + _
                               " from InvHd join CustMst on CmCtg= '" + ms_CmCtg + "' and CmCd= InCmCd " + _
                               " where InCoCd= '" + gs_CoCd + "' and " + _
                               " InTc= '" + ADC("wIdTc") + "' and " + _
                               " InYy= '" + ADC("wIdYy") + "' and " + _
                               " InChr= '" + ADC("wIdChr") + "' and " + _
                               " InNo= " + CStr(ADC("wIdNo")) + " and InPrtKey=(Case When '" + gs_Partition + "'='" + ctCurrPrtn + "' Then '" + ctCurrPrtn + "' Else InPrtKey End) ")
                               
  If Not (wRsInvDet.EOF Or wRsInvDet.BOF) Then
    
    ADC("wIdPrtKey") = wRsInvDet!InPrtKey       '****** Sachin 3.02 28-12-07
  
    ADC("wIdDt") = wRsInvDet!InDt
    ADC("wIdCmCd") = wRsInvDet!InCmCd
    
    '*** (Jen 2.13)
    ADC("wIdCmCurCd") = wRsInvDet!CmCurCd
    ADC("wIdCnvFct") = wRsInvDet!InCnvFct
    ADC("wIdMulDiv") = wRsInvDet!InMulDiv
    '*** (Jen 2.13)
    
    ADC("wIdExpNo") = wRsInvDet!InExpNo
    ADC("wIdLmgSal") = wRsInvDet!InLmgSal
    ADC("wIdLmpSal") = wRsInvDet!InLmpSal
    ' ****** Sachin 2.12 - 06/12/05 - SJM
    ADC("wIdLmsSal") = wRsInvDet!InLmsSal
    ' ****** Sachin 2.12 - 06/12/05 - SJM
    ' **** Manali 3.03 - 21/06/08 - Pd Mod - LML
    ADC("wIdLmlSal") = wRsInvDet!InLmlSal
    ' **** Manali 3.03 - 21/06/08 - Pd Mod - LML
    md_InCnvRt = wRsInvDet!InCnvRt
    ADC("wIdFOBPer") = wRsInvDet!InFOBPer
    '****Geeta****Emr208**
    ADC("wIdLmgCst") = wRsInvDet!InLmgCst
    ADC("wIdLmpCst") = wRsInvDet!InLmpCst
    ' ****** Sachin 2.12 - 06/12/05 - SJM
    ADC("wIdLmsCst") = wRsInvDet!InLmsCst
    ' ****** Sachin 2.12 - 06/12/05 - SJM
    ' **** Manali 3.03 - 21/06/08 - Pd Mod - LML
    ADC("wIdLmlCst") = wRsInvDet!InLmlCst
    ' **** Manali 3.03 - 21/06/08 - Pd Mod - LML
    ADC("wIdGldFiness") = wRsInvDet!InHGldFiness
    ADC("wIdPlFiness") = wRsInvDet!InHPlFiness
    ADC("wIdSlFiness") = wRsInvDet!InHSlFiness
    ' **** Manali 3.03 - 21/06/08 - Pd Mod - LML
    ADC("wIdPdFiness") = wRsInvDet!InHPdFiness
    ' **** Manali 3.03 - 21/06/08 - Pd Mod - LML
    mInCls = wRsInvDet!InCls
    ms_RndOffFOB = wRsInvDet!InHRndOffFOB
    
    '****** Sachin 3.3.0 (23-06-08) - [42. Make Invoice By Bag] ******
    ADC("wIdInByBag") = wRsInvDet!InByBagYN
    mb_InByBag = IIF(wRsInvDet!InByBagYN = "Y", True, False)
    '****** Sachin 3.3.0 (23-06-08) - [42. Make Invoice By Bag] ******
    
    '*** (Jen 3.01.0_1)
    mb_DsgOnHld = IIF(UCase(wRsInvDet!InDsgOnHld) = "Y", True, False)
    ms_RemakeOpt = moCn.GetFldVal("Select PValue From Param where PTyp= 'REMAKE' and PmCd= '" + wRsInvDet!InRemakingCtgs + "'")
    mb_InInvGrsWtEqActual = IIF(UCase(wRsInvDet!InInvGrsWtEqActual) = "Y", True, False)
    '*** (Jen 3.01.0_1)
    
    'Sangeeta 3.11.0 - Invoice Rate from Rate Chart
    ms_RmRtfrRtChrt = UCase(wRsInvDet!InRmRtfrRtChrt)
    ms_LabRtfrRtChrt = UCase(wRsInvDet!InLabRtfrRtChrt)
    ms_RtFrInvCust = UCase(wRsInvDet!InRtFrInvCust)
    
    ms_InOrdMulbyYN = wRsInvDet!InOrdMulbyYN
    
    '3.11.0
    ms_RtbyStkYN = IIF(wRsInvDet!InMetIRtWrtStkYN = "Y" Or wRsInvDet!InIRtWrtStkYN = "Y" Or wRsInvDet!InXIRtWrtStkYN = "Y", "Y", "N")
 
  End If
  Set wRsInvDet = Nothing
  
  '******Bhavna added chain cust logic y/n
    If moCn.RecSeek("select RrTcTyp from RmRt where RrTcTyp= 'CHN' and " + _
                         "RrCmCtg = '" + ms_CmCtg + "' and " + _
                         "RrCmCd = '" + ADC("wIdCmCd") + "' ") Then
        ADC("wChnCustYn") = "Y"
    Else
        ADC("wChnCustYn") = "N"
    End If
  '******Bhavna added chain cust logic y/n
  mf_DtLocked = IIF((ADC("wIdDt") < CDate(ms_LockDt)) _
                               Or _
                   (ADC("wIdDt") > CDate(ms_FutureDt)), _
                True, False)
  
  'mInCls = moCn.GetFldVal("Select InCls from InvHd where InCoCd= '" + gs_CoCd + "' and " + _
                          "InTc= '" + ADC("wIdTc") + "' and InYy= '" + ADC("wIdYy") + "' and " + _
                          "InChr= '" + ADC("wIdChr") + "' and InNo= " + CStr(ADC("wIdNo")))
  If GrdID.Rows <= 1 And mInCls = "N" And Not mf_DtLocked Then
    CmdMI.Enabled = True
  Else
    CmdMI.Enabled = False
  End If
  
  
  '****** Sachin 3.02 - Locking records for previous partition
  If ADC("WIDPRTKEY") <> ctCurrPrtn Or mInCls = "Y" Or mf_DtLocked Then
    GrdID.AllowAdd = False: GrdID.AllowDelete = False
    GrdIR.AllowAdd = False: GrdIR.AllowDelete = False
    GrdIL.AllowAdd = False: GrdIL.AllowDelete = False
    GrdIB.AllowAdd = False: GrdIB.AllowDelete = False
  Else
    If ADC("WIDPRTKEY") = ctCurrPrtn And mf_DtLocked = False Then
      GrdID.AllowAdd = True: GrdID.AllowDelete = True
      GrdIR.AllowAdd = True: 'GrdIR.AllowDelete = True
      GrdIL.AllowAdd = True: GrdIL.AllowDelete = True
      GrdIB.AllowAdd = True: GrdIB.AllowDelete = True
    End If
  End If

  '*** (Jen 3.01.0_1)
  If mb_DsgOnHld = True Then
    GrdID.AllowAdd = False
  Else
    If mf_DtLocked = False And mInCls = "N" Then GrdID.AllowAdd = True
  End If
  '*** (Jen 3.01.0_1)

  GrdIR.AllowDelete = False  '****** Sachin 2.13.0 - 27-05-2006 ******

  'ms_RndOffFOB = moCn.GetFldVal("Select InHRndOffFOB From InvHd where InCoCd='" + gs_CoCd + "' and " + _
                 "InTc='" + ADC("wIdTc") + "' and InYy= '" + ADC("wIdYy") + "' and " + _
                 "InChr='" + ADC("wIdChr") + "' and InNo= " + CStr(ADC("wIdNo")) + " ")
  ms_ShowActWt = moCn.GetFldVal("Select HShowActWt From Head where HCoCd= '" + gs_CoCd + "' and HCd='" + ctSelfCmCd + "'")
  If ms_ShowActWt = "Y" Then
    GrdIR.ColProp("IrRmAWt").Style = en_fgColStyle.fgcReadOnly
  Else
    GrdIR.ColProp("IrRmAWt").Style = en_fgColStyle.fgcinvisible
  End If
  
  '****** Sachin 2.13.0 - 05-06-2006 ******
  If mf_QWModYn = True Then
    GrdIR.ColProp("WQWSAL").Style = fgcReadOnly
  Else
    GrdIR.ColProp("WQWSAL").Style = fgcinvisible
  End If
  '****** Sachin 2.13.0 - 05-06-2006 ******
  
  ADC("wInRefreshDt") = moCn.SrvrDate       ' ***** Manali 3.6.0 - 29/10/09 - Refresh Metal Reate
  
  ' ***** Manali 3.8.0 - Default Invoice Sort By Option
  ADC("wInSrt") = moCn.GetFldVal("Select HInSrt From Head Where HCoCd= '" + gs_CoCd + "' and HCd='" + ctSelfCmCd + "'")
  
  ms_PwdFlag = False    ' **** Manali 3.6.0 - 19/11/09 - Pwd for Refresh Rate
'MWERROR IG.132 may'17
  Dim wHCoCd As String
  wHCoCd = moCn.GetFldVal("Select InCoCd From InvHd Where InCoCd= '" + gs_CoCd + "' and " + _
           "InTc= '" + ADC("wIdTc") + "' And InYy= '" + ADC("wIdYy") + "' and " + _
           "InChr= '" + ADC("wIdChr") + "' and InNo= " + CStr(ADC("wIdNo")) + " and InPrtKey=(Case When '" + gs_Partition + "'='" + ctCurrPrtn + "' Then '" + ctCurrPrtn + "' Else InPrtKey End) " + _
           "")
  cmdGetLotRt.Visible = False
  cmdGetFOB.Visible = False   '18 GetFOb will be visible only for  DJ & MW1
  
  If wHCoCd = "IG" Or wHCoCd = "MW" Then
    cmdGetLotRt.Visible = True
  End If
  '18 GetFOb will be visible only for  DJ & MW1
  If wHCoCd = "MW" Or wHCoCd = "DJ" Or wHCoCd = "MW1" Then
    cmdGetFOB.Visible = True: cmdGetFOB.Enabled = True
    If mInCls = "Y" Or mf_DtLocked = True Then cmdGetFOB.Enabled = False  '18.b if invoice locked then getfob button will disable
  End If
  'Initially IG.132 had asked for a field to be accepted ALLDCX was later removed from specs
  'So instead of removing it from code all over have left the field in just initialized internally always to Y
  ADC("WGETLOTRTALLDCX") = "Y"
  '6.1-13 first value not assinged to ms_hSizeStkRt, so MakeRm not calling
  'now value assigned to ms_hSizeStkRt
  ms_hSizeStkRt = moCn.GetFldVal("Select hSizeStkRt from Head where HCoCd= '" + gs_CoCd + "' and " + _
                      "HCd= '" + ctSelfCmCd + "' And hSizeStkRt = 'N'")
  '6.1b-8 stores ptr to be used for rate calculation
  ms_hInvPtr = moCn.GetFldVal("Select HInvPtr from Head where HCoCd= 'ZZZ' ")
  
End Sub
Private Sub ATXT_FldWhen(Index As Integer, ByVal IdName As String, _
                          Cancel As Boolean, ErrMsg As String)
  '*** Give help for wIdTc, wIdYy, wIdChr, wIdNo
  '*** Give help for all fields in the Design Details frame
  '*** Give help for the fields wDRGrp, wDLGrp, wIRGrp, wILGrp, wIDGrp in the Summary frame
  '*** Give help for all the fields in the Design Copy frame
  
  Dim i As Integer
  
  '****** Sachin 3.02 - Locking Previous Partition Key Records
  '6.1_37 dmctg dmcd, sr fr/to added to get input
  If (ADC("wIdPrtKey") <> ctCurrPrtn Or mf_DtLocked = True) _
       And (IdName <> UCase("WIDTC") And IdName <> UCase("WIDYY") And IdName <> UCase("WIDCHR") And IdName <> UCase("WIDNO") And _
            IdName <> UCase("wDRGrp") And IdName <> UCase("wIRGrp") And _
            IdName <> UCase("wDLGrp") And IdName <> UCase("wILGrp") And IdName <> UCase("wIDGrp") _
            And IdName <> UCase("wScDmCtgFr") And IdName <> UCase("wScDmCtgTo") _
            And IdName <> UCase("wScDmCdFr") And IdName <> UCase("wScDmCdTo") _
            And IdName <> UCase("wScIdSrFr") And IdName <> UCase("wScIdSrTo")) Then
    Cancel = True: ErrMsg = "Cannot Edit Previous Partition/Locked Entry": Exit Sub
  End If
  
  '*** (Jen 2.13)
  'wDdIdKt, wDdIdIAdd, wDdIdZAdd, wDdIdCustmMtchWt
  '****** Sachin 2.13.0 - wDdIdLabAs - 26-05-2006 ******
  '****** Sachin 3.02.0 28-12-07 - prtkey added below
  If moCn.RecSeek("Select 'x' from InvZRm where IzCoCd= '" + gs_CoCd + "' and IzTc= '" + ADC("wIdTc") + "' and " + _
                  "IzYy=  '" + ADC("wIdYy") + "' and IzChr=  '" + ADC("wIdChr") + "' and " + _
                  "IzNo=  " + CStr(ADC("wIdNo")) + " and IzPrtKey='" + ctCurrPrtn + "' ") And _
                  (IdName = UCase("wDdIdValAddn") Or IdName = UCase("wDdIdVaCtg") Or _
                   IdName = UCase("wDdIdGldLs") Or IdName = UCase("wDdIdGldAs") Or _
                   IdName = UCase("wDdIdLabAs") Or IdName = UCase("wDdIdHZValEqInv") Or _
                   IdName = UCase("wDdIdCustmMtchWt") Or IdName = UCase("wDdIdFixPrc")) Then
    '
    Cancel = True: ErrMsg = "Cannot Edit As Custom Matching Has Been Done": Exit Sub
  End If
  '*** (Jen 2.13)
  
  Select Case IdName
  Case Is = UCase("wIdTc")
    Call HlpList.PMCd("TC", "'" + ms_IdTcTyp + "'")
  Case Is = UCase("wIdYy")
    Call HlpList.PMCd("YY")
  Case Is = UCase("wIdChr")
' **** Zubin 211 **** '
    'Call HlpList.vPSCd(gs_CoCd, "CHR", ADC("wIdTc"))
    Call HlpList.vPSCd(gs_CoCd, "CHR", ADC("wIdTc"), "Y")
' **** Zubin 211 **** '
  Case Is = UCase("wIdNo")
    Call HlpList.InNo(gs_CoCd, ADC("wIdTc"), ADC("wIdYy"), ADC("wIdChr"))
  Case Is = UCase("wDdIdKt")
    If mInCls = "Y" Then Cancel = True: ErrMsg = "Cannot Edit As Invoice Is Closed": Exit Sub
    Call HlpList.PMCd("KT")
  Case Is = UCase("wDdIdFixPrc")
    If mInCls = "Y" Then Cancel = True: ErrMsg = "Cannot Edit As Invoice Is Closed": Exit Sub
    Call HlpList.PMCd("YN")
 Case Is = UCase("WGETLOTRTONLYCUST")
    Call HlpList.PMCd("YN")
  Case Is = UCase("wDdIdHZValEqInv")
    If mInCls = "Y" Then Cancel = True: ErrMsg = "Cannot Edit As Invoice Is Closed": Exit Sub
    Call HlpList.PMCd("YN")
  Case Is = UCase("wDdIdVaCtg")
    If mInCls = "Y" Then Cancel = True: ErrMsg = "Cannot Edit As Invoice Is Closed": Exit Sub
    Call HlpList.PMCd("VACTG")
  Case Is = UCase("wDdIdGldAs"), UCase("wDdIdLabAs")
    If mInCls = "Y" Then Cancel = True: ErrMsg = "Cannot Edit As Invoice Is Closed": Exit Sub
    Call HlpList.PMCd("GLDAS")
  Case Is = UCase("wDRGrp"), UCase("wIRGrp")
   Call HlpList.PMCd("ORDRMGR")
  Case Is = UCase("wDLGrp"), UCase("wILGrp")
    Call HlpList.PMCd("ORDLABGR")
  Case Is = UCase("wIDGrp")
    Call HlpList.PMCd("ORDDSGGR")
  Case Is = UCase("wDcOdTcFr"), UCase("wDcOdTcTo")
    Call HlpList.PMCd("TC", "'SO'")
  Case Is = UCase("wDcOdYyFr"), UCase("wDcOdYyTo")
    Call HlpList.PMCd("YY")
  Case Is = UCase("wDcOdChrFr")
    Call HlpList.vPSCd(gs_CoCd, "CHR", ADC("wDcOdTcFr"))
  Case Is = UCase("wDcOdNoFr")
    Call HlpList.OmNo(gs_CoCd, ADC("wDcOdTcFr"), ADC("wDcOdYyFr"), ADC("wDcOdChrFr"))
  Case Is = UCase("wDcOdSrFr")
    Call HlpList.OdSr(gs_CoCd, ADC("wDcOdTcFr"), ADC("wDcOdYyFr"), ADC("wDcOdChrFr"), ADC("wDcOdNoFr"))
  
  Case Is = UCase("wDcOdChrTo")
    Call HlpList.vPSCd(gs_CoCd, "CHR", ADC("wDcOdTcTo"))
  Case Is = UCase("wDcOdNoTo")
    Call HlpList.OmNo(gs_CoCd, ADC("wDcOdTcTo"), ADC("wDcOdYyTo"), ADC("wDcOdChrTo"))
  Case Is = UCase("wDcOdSrTo")
    Call HlpList.OdSr(gs_CoCd, ADC("wDcOdTcTo"), ADC("wDcOdYyTo"), ADC("wDcOdChrTo"), ADC("wDcOdNoTo"))
  Case Is = UCase("wWaitForAllBags")
    Call HlpList.PMCd("YN")
  
  '*** (Jen 2.13)
  Case Is = UCase("wDdIdCustmMtchWt")
    Dim wb_Valid As Boolean, ws_RmCtg As String
    If ADC.Mode = xNorm Then
      wb_Valid = False
      For i = 1 To GrdIR.Rows - 1
        If GrdIR.Value(i, "IrMainMet") = "Y" And (Not GrdIR.IsDel(i)) Then
          ws_RmCtg = UCase(moCn.GetFldVal("Select RmCtg from RmMst where RmCd= '" + GrdIR.Value(i, "IrRmCd") + "'" + IIF(gs_Partition = ctCurrPrtn, " And RmPrtKey='" + GrdIR.Value(i, "IrPrtKey") + "' ", "")))
          ' **** Manali 3.03 - 21/06/08 - Pd Mod - "L" included
          If ws_RmCtg = "G" Or ws_RmCtg = "P" Or ws_RmCtg = "S" Or ws_RmCtg = "L" Then wb_Valid = True
        End If
      Next i
      If wb_Valid = False Then Cancel = True: ErrMsg = "Cannot Edit As Main Metal Category Is Not 'G', 'P', 'S' or 'L'": Exit Sub
        ' **** Manali 3.03 - 21/06/08 - Pd Mod - "L" included
    End If
  '*** (Jen 2.13)
  
  ' ***** Manali 3.8.0 - Invoice Sort By
  Case Is = UCase("wInSrt")
    Call HlpList.PMCd("INSRT")
    
  '3.11.0
  Case Is = UCase("wRefCnvRtYN")
    Call HlpList.PMCd("YN")
  '6.1_37 calling help
  Case Is = UCase("wScDmCtgFr"), UCase("wScDmCtgTo")
    Call HlpList.PMCd("DMCTG")
  Case Is = UCase("wScDmCdFr"), UCase("wScDmCdTo")
    Call HlpList.DmCd("DM")
  Case Is = UCase("wScIdSrFr"), UCase("wScIdSrTo")
    Call HlpList.InSr(gs_CoCd, ADC("wIdTc"), ADC("wIdYy"), ADC("wIdChr"), ADC("wIdNo"))
    
  End Select
End Sub
Private Sub atxt_ReCalculate(Index As Integer, ByVal IdName As String, ByVal OldVal As Variant, ByVal RowNum As Integer, ByVal SourceName As String, ByVal SourceColName As String)
  '*** wIdChr on wIdTc                                    Normal
      'PDesc225('TC', adc(wIdTc))
  '*** wDcOdChr on wDcOdTc                                Normal
      'PDesc225('TC', adc(wDcOdTc))
  '*** wDdIdGldLs on wDdIdVaCtg                           Normal
      'PNum('VACTG', adc(wDdIdVaCtg))
  '*** wTotIdIFob on IdIFob                               ADC= Normal, GrdID= Normal,Delete,InitOver,CopyOver
      'sum(IdIFob) over the entire invoice
  '*** wTotIdZFob on IdZFob                               ADC= Normal, GrdID= Normal,Delete,InitOver,CopyOver
      'sum(IdZFob) over the entire invoice
  
  Dim i As Integer
  Dim wDtVal As String, wDtDt As String
  
  Select Case UCase(IdName)
  Case Is = UCase("wIdChr")           '*** recalc on wIdTc
    If ADC.Mode = xNorm Then
      ADC("wIdChr") = moCn.GetFldVal("Select vPValue from vParam where " + _
                                "vPCoCd= '" + gs_CoCd + "' and vPTyp='DEFCHR' " + _
                                "and vPMCd= '" + ADC("wIdTc") + "' and vPSCd= ''")
      '("Select PDesc225 from Param where PTyp='TC' " + _
                                     "and PMCd='" + ADC("wIdTc") + "'")
    End If
  Case Is = UCase("wDcOdChrFr")         '*** recalc on wDcOdTcFr
    If ADC.Mode = xNorm Then
      ADC("wDcOdChrFr") = moCn.GetFldVal("Select vPValue from vParam where " + _
                               "vPCoCd= '" + gs_CoCd + "' and vPTyp= 'DEFCHR' " + _
                               "and vPMCd= '" + ADC("wDcOdTcFr") + "' and vPSCd= ''")
      '("Select PDesc225 from Param where PTyp='TC' " + _
                                       "and PMCd='" + ADC("wDcOdTc") + "'")
    End If
  Case Is = UCase("wDcOdTcTo")         '*** recalc on wDcOdTcFr
    If ADC.Mode = xNorm Then ADC("wDcOdTcTo") = ADC("wDcOdTcFr")
  Case Is = UCase("wDcOdYyTo")         '*** recalc on wDcOdYyFr
    If ADC.Mode = xNorm Then ADC("wDcOdYyTo") = ADC("wDcOdYyFr")
  Case Is = UCase("wDcOdChrTo")         '*** recalc on wDcOdChrFr
    If ADC.Mode = xNorm Then
      ADC("wDcOdChrTo") = ADC("wDcOdChrFr")
      'ADC("wDcOdChrTo") = moCn.GetFldVal("Select vPValue from vParam where " + _
                               "vPCoCd= '" + gs_CoCd + "' and vPTyp= 'DEFCHR' " + _
                               "and vPMCd='" + ADC("wDcOdTcTo") + "'")
    End If
  Case Is = UCase("wDcOdNoTo")         '*** recalc on wDcOdNoFr
    If ADC.Mode = xNorm Then ADC("wDcOdNoTo") = ADC("wDcOdNoFr")
  Case Is = UCase("wDcOdSrTo")         '*** recalc on wDcOdSrFr
    If ADC.Mode = xNorm Then ADC("wDcOdSrTo") = ADC("wDcOdSrFr")
  Case Is = UCase("wDdIdGldLs")       '*** recalc on wDdIdVaCtg
    If ADC.Mode = xNorm And UCase(ms_IdTcTyp) = UCase(ctInTcTyp) Then
      ADC("wDdIdGldLs") = moCn.GetFldVal("Select PNum from Param where PTyp='VACTG' " + _
                                       "and PMCd= '" + ADC("wDdIdVaCtg") + "' and PSCd= ''")
    End If
  Case Is = UCase("wTotIdIFob")       '*** recalc on IdIFob
    'Dim wIdiFOB As Single ' ******** Manali 3.03 19/09/08 - To remove decimal rounding off problem changed to double
    Dim wIdiFOB As Double
    If ADC.Mode = xNorm Then
      If GrdID.Mode = fgmnorm Or GrdID.Mode = fgmdel Or GrdID.Mode = fgmInitOver Or GrdID.Mode = fgmCopyOver Then
        For i = 1 To GrdID.Rows - 1
          If Not GrdID.IsDel(i) Then
            wIdiFOB = wIdiFOB + GrdID.Value(i, "IdIFob")
          End If
        Next i
        ADC("wTotIdIFob") = wIdiFOB
      End If
    End If
  Case Is = UCase("wTotIdZFob")       '*** recalc on IdZFob
    'Dim wIdZFOB As Single  ' ******** Manali 3.03 19/09/08 - To remove decimal rounding off problem changed to double
    Dim wIdZFOB As Double
    If ADC.Mode = xNorm Then
      If GrdID.Mode = fgmnorm Or GrdID.Mode = fgmdel Or GrdID.Mode = fgmInitOver Or GrdID.Mode = fgmCopyOver Then
        For i = 1 To GrdID.Rows - 1
          If Not GrdID.IsDel(i) Then
            wIdZFOB = wIdZFOB + GrdID.Value(i, "IdZFob")
          End If
        Next i
        ADC("wTotIdZFob") = wIdZFOB
      End If
    End If

'  Case Is = UCase("wDdIdGldAsWt")     '*** recalc on wDdIdGldAs
'    If ADC.Mode = fgmnorm Then
'      ADC("wDdIdGldAsWt") =
'    End If
    '*********Geeta***Emr208**
    Case Is = UCase("wDdIdValAddn")       '*** recalc on wDdIdVaCtg
    If ADC.Mode = xNorm And UCase(ms_IdTcTyp) = UCase(ctInTcTyp) Then
      ADC("wDdIdValAddn") = moCn.GetFldVal("Select PNum1 from Param where PTyp= 'VACTG' " + _
                                       "and PMCd= '" + ADC("wDdIdVaCtg") + "' and PSCd= ''")
    End If
    '**********
    
    '*** (Jen 2.13)
    Case Is = UCase("wDdIdCustmMtchWt")
      '*** recalc on wDdIdGldLs
      Dim wd_IdCustmMtchWt As Double, ws_RmCtg As String, wd_IdGldLs As Double
      If ADC.Mode = xNorm Then
        wd_IdCustmMtchWt = 0: wd_IdGldLs = 0
        For i = 1 To GrdIR.Rows - 1
          If GrdIR.Value(i, "IrMainMet") = "Y" And (Not GrdIR.IsDel(i)) Then
            ws_RmCtg = UCase(moCn.GetFldVal("Select RmCtg from RmMst where RmCd= '" + GrdIR.Value(i, "IrRmCd") + "'" + IIF(gs_Partition = ctCurrPrtn, " And RmPrtKey='" + GrdIR.Value(i, "IrPrtKey") + "' ", "")))
            ' **** Manali 3.03 - 21/06/08 - Pd Mod - "L" included
            If ws_RmCtg = "G" Or ws_RmCtg = "P" Or ws_RmCtg = "S" Or ws_RmCtg = "L" Then
              wd_IdGldLs = ADC("wDdIdGldLs")
            End If
            wd_IdCustmMtchWt = GrdIR.Value(i, "IrRmZWt") * (1 + wd_IdGldLs / 100)
            Exit For
          End If
        Next i
        ADC("wDdIdCustmMtchWt") = wd_IdCustmMtchWt
      End If
    '*** (Jen 2.13)
    
    ' ***** Manali 3.6.0 - 29/10/09 - Refresh Metal Rate
    Case Is = UCase("wInLmgCst")
      If ADC.Mode = xNorm Then
        Call GetDtVal("LMG", ADC("wInRefreshDt"), wDtVal, , ADC("wIdCmCurCd"))
  
        ADC("wInLmgCst") = wDtVal
      End If
    
    Case Is = UCase("wInLmgDt")
      If ADC.Mode = xNorm Then
        Call GetDtVal("LMG", ADC("wInRefreshDt"), , wDtDt, ADC("wIdCmCurCd"))
        
        ADC("wInLmgDt") = wDtDt
      End If
   
    Case Is = UCase("wInLmgSal")
      If ADC.Mode = xNorm Then
        ADC("wInLmgSal") = GetLmgSal("LMG", ADC("wIdCmCd"), ADC("wInLmgCst"), ms_CmCtg, ADC("wIdCmCurCd"))
        
        If ADC("wInLmgSal") = 0 Then ADC("wInLmgSal") = ADC("wInLmgCst")
      End If
   
    Case Is = UCase("wInLmpCst")
      If ADC.Mode = xNorm Then
        Call GetDtVal("LMP", ADC("wInRefreshDt"), wDtVal, , ADC("wIdCmCurCd"))
  
        ADC("wInLmpCst") = wDtVal
      End If
    
    Case Is = UCase("wInLmpDt")
      If ADC.Mode = xNorm Then
        Call GetDtVal("LMP", ADC("wInRefreshDt"), , wDtDt, ADC("wIdCmCurCd"))
        
        ADC("wInLmpDt") = wDtDt
      End If
   
    Case Is = UCase("wInLmpSal")
      If ADC.Mode = xNorm Then
        ADC("wInLmpSal") = GetLmgSal("LMP", ADC("wIdCmCd"), ADC("wInLmpCst"), ms_CmCtg, ADC("wIdCmCurCd"))

        If ADC("wInLmpSal") = 0 Then ADC("wInLmpSal") = ADC("wInLmpCst")
      End If
   
   Case Is = UCase("wInLmsCst")
      If mSlvMod = True Then
        If ADC.Mode = xNorm Then
          Call GetDtVal("LMS", ADC("wInRefreshDt"), wDtVal, , ADC("wIdCmCurCd"))
    
          ADC("wInLmsCst") = wDtVal
        End If
      Else
          ADC("wInLmsCst") = 1
      End If
      
    Case Is = UCase("wInLmsDt")
      If mSlvMod = True Then
        If ADC.Mode = xNorm Then
          Call GetDtVal("LMS", ADC("wInRefreshDt"), , wDtDt, ADC("wIdCmCurCd"))
          
          ADC("wInLmsDt") = wDtDt
        End If
      Else
        ADC("wInLmsDt") = ADC("wIdDt")
      End If
      
    Case Is = UCase("wInLmsSal")
      If mSlvMod = True Then
        If ADC.Mode = xNorm Then
          
          ADC("wInLmsSal") = GetLmgSal("LMS", ADC("wIdCmCd"), ADC("wInLmsCst"), ms_CmCtg, ADC("wIdCmCurCd"))
          
          If ADC("wInLmsSal") = 0 Then ADC("wInLmsSal") = ADC("wInLmsCst")
        End If
      Else
        ADC("wInLmsSal") = 1
      End If
      
    Case Is = UCase("wInLmlCst")
      If mb_PdMod = True Then
        If ADC.Mode = xNorm Then
          Call GetDtVal("LML", ADC("wInRefreshDt"), wDtVal, , ADC("wIdCmCurCd"))
    
          ADC("wInLmlCst") = wDtVal
        End If
      Else
        ADC("wInLmlCst") = 1
      End If
      
    Case Is = UCase("wInLmlDt")
      If mb_PdMod = True Then
        If ADC.Mode = xNorm Then
          Call GetDtVal("LML", ADC("wInRefreshDt"), , wDtDt, ADC("wIdCmCurCd"))
          
          ADC("wInLmlDt") = wDtDt
        End If
      Else
        ADC("wInLmlDt") = ADC("wIdDt")
      End If
      
    Case Is = UCase("wInLmlSal")
      If mb_PdMod = True Then
        If ADC.Mode = xNorm Then
          ADC("wInLmlSal") = GetLmgSal("LML", ADC("wIdCmCd"), ADC("wInLmlCst"), ms_CmCtg, ADC("wIdCmCurCd"))
          
          If ADC("wInLmlSal") = 0 Then ADC("wInLmlSal") = ADC("wInLmlCst")
        End If
      Else
        ADC("wInLmlSal") = 1
      End If
    ' ***** Manali 3.6.0 - 29/10/09 - Refresh Metal Rate
    
    '3.11.0
    Case Is = UCase("wInCnvRt")
        If ADC.Mode = xNorm Then _
          ADC("wInCnvRt") = CnvRt(moCn.GetFldVal("Select HBaseCurCd from Head where HCoCd= '" + gs_CoCd + "' and HCd= '" + ctSelfCmCd + "' "), _
                          ADC("wIdCmCurCd"), ADC("wInRefreshDt"))
          'CnvRt(adc("wIdCmCurCd"), ctIndCurncy, adc("wInRefreshDt"))
          
 '******bhavna added chain logic tag
  Case Is = UCase("wChnCustYn")
    If ADC.Mode = xNorm Then
    If moCn.RecSeek("select RrTcTyp from RmRt where RrTcTyp= 'CHN' and " + _
                         "RrCmCtg = '" + ms_CmCtg + "' and " + _
                         "RrCmCd = '" + ADC("wIdCmCd") + "' ") Then
        ADC("wChnCustYn") = "Y"
    Else
        ADC("wChnCustYn") = "N"
    End If
    End If
  
  End Select
End Sub
Private Sub ATXT_FldChange(Index As Integer, ByVal IdName As String, Cancel As Boolean, ErrMsg As String, ByVal pv_NewValue As Variant)
  '*** wIdTc        valid PMCd('TC', PValue= ctInTcTyp) and the type of entry should be 'IN'
  '*** wIdYy        valid PMCd('YY')
  '*** wIdChr       valid PSCd('CHR', adc(wIdTc)) or ''
  '*** wIdNo        valid InNo from InvHd(InTc= adc(wIdTc), InYy= adc(wIdYy),
      '             InChr= adc(wIdChr)) or Zero
  '*** Give validation for all fields in the Design Details frame
  '*** Give validation for all the fields in the Design Copy frame
  '*** Give validation for group fields of the Summary frame
  
  Dim i As Integer  '*** (Jen 2.13)
  
  Select Case IdName
  Case Is = UCase("wIdTc")
    Cancel = Not moCn.RecSeek("select PMCd from Param where PTyp= 'TC' and " + _
                         "PMCd = '" + pv_NewValue + "' and PSCd= '' and " + _
                         "PValue= '" + ms_IdTcTyp + "' ")
    If Cancel = True Then ErrMsg = "Invalid Invoice Voucher Tc": Exit Sub
  Case Is = UCase("wIdYy")
    Cancel = Not moCn.RecSeek("select PMCd from Param where PTyp= 'YY' and " + _
                         "PMCd = '" + pv_NewValue + "' and PSCd= ''")
    If Cancel = True Then ErrMsg = "Invalid Invoice Voucher Year": Exit Sub
  Case Is = UCase("wIdChr")
    Cancel = pv_NewValue <> "" And (Not moCn.RecSeek("select vPSCd from vParam where " + _
                         "vPCoCd= '" + gs_CoCd + "' and vPTyp= 'CHR' and " + _
                         "vPMCd = '" + ADC("wIdTc") + "' and vPSCd= '" + pv_NewValue + "'"))
    If Cancel = True Then ErrMsg = "Invalid Invoice Voucher Character": Exit Sub
' **** Zubin 211 **** '
    If Cancel = True Then ErrMsg = "Invalid Voucher Character": Exit Sub
    Dim ws_UsrArr() As String, wi_Index As Integer
    ws_UsrArr = Split(moCn.GetFldVal("Select vPValue4 from vParam where vPCoCd= '" + gs_CoCd + "' and " + _
                                        "vPTyp='CHR' and vPMCd='" + ADC("wIdTc") + "' and vPSCd= '" + pv_NewValue + "'"), ",")
    If IsArray(ws_UsrArr) Then
      For wi_Index = 0 To UBound(ws_UsrArr)
        Cancel = IIF(UCase(gs_UsrCd) = UCase(ws_UsrArr(wi_Index)), False, True)
        If Cancel = False Then: Exit For
      Next wi_Index
    End If
   If Cancel = True Then ErrMsg = "Access To This Invoice Voucher Character Is Denied To User " + gs_UsrCd: Exit Sub
' **** Zubin 211 **** '

  Case Is = UCase("wIdNo")
    '****** Sachin 3.02 28-12-07
    Cancel = (pv_NewValue <> 0) And (Not moCn.RecSeek("Select InNo From InvHd Where " + _
                                    "InCoCd= '" + gs_CoCd + "' and " + _
                                    "InTc= '" + ADC("wIdTc") + "' And " + _
                                    "InYy= '" + ADC("wIdYy") + "' and " + _
                                    "InChr= '" + ADC("wIdChr") + "' and " + _
                                    "InNo= " + CStr(pv_NewValue) + " and InPrtKey=(Case When '" + gs_Partition + "'='" + ctCurrPrtn + "' Then '" + ctCurrPrtn + "' Else InPrtKey End) " + _
                                    IIF(gs_Partition = ctCurrPrtn, " And InDt>'" + CStr(ms_PurgeDt) + "'", "")))
    If Cancel = True Then ErrMsg = "Invalid Voucher Invoice Number": Exit Sub
  Case Is = UCase("wDdIdKt")
    Cancel = Not moCn.RecSeek("Select PMCd From Param Where " + _
                              "PTyp= 'KT' And PMCd = '" + pv_NewValue + "'")
    ErrMsg = "Invalid Karatage": Exit Sub
  Case Is = UCase("wDdIdVaCtg")
    Cancel = Not moCn.RecSeek("select PMCd from Param where " + _
                         "PTyp= 'VACTG' and PMCd = '" + pv_NewValue + "'")
    ErrMsg = "Invalid Value Addition Category": Exit Sub
  Case Is = UCase("wDdIdGldLs")
    If pv_NewValue < 0 Then Cancel = True: ErrMsg = "Invalid Gold Loss": Exit Sub
  Case Is = UCase("wDdIdFixPrc")
    Cancel = Not moCn.RecSeek("select PMCd from Param where PTyp= 'YN' and " + _
                         "PMCd = '" + pv_NewValue + "'")
    ErrMsg = "Enter Fixed Price (Y/N)": Exit Sub
  Case Is = UCase("WGETLOTRTONLYCUST")
    Cancel = Not moCn.RecSeek("select PMCd from Param where PTyp= 'YN' and " + _
                         "PMCd = '" + pv_NewValue + "'")
    ErrMsg = "Specify Y for Only Customer N for All": Exit Sub
  Case Is = UCase("wDdIdHZValEqInv")
    Cancel = Not moCn.RecSeek("select PMCd from Param where PTyp= 'YN' and " + _
                         "PMCd = '" + pv_NewValue + "'")
    ErrMsg = "Specify whether Customs Value Equal To Invoice Value. Enter (Y/N)": Exit Sub
  Case Is = UCase("wDdIdGldAs")
    '*** (Bef 2.14)
    'Cancel = Not moCn.RecSeek("select PMCd from Param where " + _
    '                     "PTyp= 'GLDAS' and PMCd = '" + pv_NewValue + "'")
    '*** (Bef 2.14)
    
    '*** (Jen 2.14)
    'Cancel = pv_NewValue <> "" And Not moCn.RecSeek("select PMCd from Param where PTyp= 'GLDAS' " + _
    '         "and PMCd = '" + pv_NewValue + "' and not (',' + PDesc225 + ',' like '%,G,%' Or ',' + PDesc225 + ',' like '%,P,%' Or ',' + PDesc225 + ',' like '%,S,%') ")
    
    Cancel = pv_NewValue <> "" And Not moCn.RecSeek("select PMCd from Param where PTyp= 'GLDAS' and PMCd = '" + pv_NewValue + "' ")
    '*** (Jen 2.14)
    
    ErrMsg = "Invalid Gold As": Exit Sub
    
  '****** Sachin 2.13.0 - 26-05-2006 ******
  Case Is = UCase("wDdIdLabAs")
    Cancel = Not moCn.RecSeek("select PMCd from Param where " + _
                         "PTyp= 'GLDAS' and PMCd = '" + pv_NewValue + "'")
    ErrMsg = "Invalid Labour As": Exit Sub
  '****** Sachin 2.13.0 - 26-05-2006 ******
    
    '*********Geeta***Emr208**
    Case Is = UCase("wDdIdValAddn")
    If pv_NewValue < 0 Then Cancel = True: ErrMsg = "Invalid Value Addition": Exit Sub
    '**********
  
'  Case Is = UCase("wDcOdTcFr"), UCase("wDcOdTcTo")
'    Cancel = pv_NewValue <> "" And ( _
'                    moCn.GetFldVal("Select PValue From Param Where PTyp= 'TC' " + _
'                    "And PMCd = '" + pv_NewValue + "'") <> "SO" _
'                    Or _
'                    moCn.GetFldVal("select PValue1 from Param where PTyp= 'TcTyp' and " + _
'                    "PMCd = (select PValue from Param where PTyp='TC' and " + _
'                    "PMCd='" + pv_NewValue + "')") <> "OM")
'    ErrMsg = "Invalid Order Tc": Exit Sub
'  Case Is = UCase("wDcOdYyFr"), UCase("wDcOdYyTo")
'    Cancel = (pv_NewValue <> "") And (Not moCn.RecSeek("Select PMCd From Param Where " + _
'                                  "PTyp= 'Yy' And PMCd = '" + pv_NewValue + "'"))
'    ErrMsg = "Invalid Order Yy": Exit Sub
'  Case Is = UCase("wDcOdChrFr")
'    Cancel = (pv_NewValue <> "") And (Not moCn.RecSeek("Select vPSCd From vParam Where " + _
'                                  "vPCoCd= '" + gs_CoCd + "' and vPTyp= 'CHR' And " + _
'                                  "vPMCd = '" + ADC("wDcOdTcFr") + "' and vPSCd='" + pv_NewValue + "'"))
'    ErrMsg = "Invalid From Order Chr": Exit Sub
'  Case Is = UCase("wDcOdChrTo")
'    Cancel = (pv_NewValue <> "") And (Not moCn.RecSeek("Select vPSCd From vParam Where " + _
'                                  "vPCoCd= '" + gs_CoCd + "' and vPTyp= 'CHR' And " + _
'                                  "vPMCd = '" + ADC("wDcOdTcTo") + "' and vPSCd='" + pv_NewValue + "'"))
'    ErrMsg = "Invalid To Order Chr": Exit Sub
'  Case Is = UCase("wDcOdNoFr")
'    Cancel = (pv_NewValue <> 0) And (Not moCn.RecSeek("Select OdNo From OrdDsg Where " + _
'                                  "OdCoCd= '" + gs_CoCd + "' and " + _
'                                  "OdTc= '" + ADC("wDcOdTcFr") + "' And " + _
'                                  "OdYy= '" + ADC("wDcOdYyFr") + "' and " + _
'                                  "OdChr= '" + ADC("wDcOdChrFr") + "' and " + _
'                                  "OdNo= " + CStr(pv_NewValue)))
'    ErrMsg = "Invalid From Order No": Exit Sub
'  Case Is = UCase("wDcOdNoTo")
'    Cancel = (pv_NewValue <> 0) And (Not moCn.RecSeek("Select OdNo From OrdDsg Where " + _
'                                  "OdCoCd= '" + gs_CoCd + "' and " + _
'                                  "OdTc= '" + ADC("wDcOdTcTo") + "' And " + _
'                                  "OdYy= '" + ADC("wDcOdYyTo") + "' and " + _
'                                  "OdChr= '" + ADC("wDcOdChrTo") + "' and " + _
'                                  "OdNo= " + CStr(pv_NewValue)))
'    ErrMsg = "Invalid To Order No": Exit Sub
'  Case Is = UCase("wDcOdSrFr")
'    Cancel = (pv_NewValue <> 0) And (Not moCn.RecSeek("Select OdSr From OrdDsg Where " + _
'                                  "OdCoCd= '" + gs_CoCd + "' and " + _
'                                  "OdTc= '" + ADC("wDcOdTcFr") + "' And " + _
'                                  "OdYy= '" + ADC("wDcOdYyFr") + "' and " + _
'                                  "OdChr= '" + ADC("wDcOdChrFr") + "' and " + _
'                                  "OdNo= " + CStr(ADC("wDcOdNoFr")) + " and " + _
'                                  "OdSr= " + CStr(pv_NewValue)))
'    ErrMsg = "Invalid From Order Sr": Exit Sub
'  Case Is = UCase("wDcOdSrTo")
'    Cancel = (pv_NewValue <> 0) And (Not moCn.RecSeek("Select OdSr From OrdDsg Where " + _
'                                  "OdCoCd= '" + gs_CoCd + "' and " + _
'                                  "OdTc= '" + ADC("wDcOdTcTo") + "' And " + _
'                                  "OdYy= '" + ADC("wDcOdYyTo") + "' and " + _
'                                  "OdChr= '" + ADC("wDcOdChrTo") + "' and " + _
'                                  "OdNo= " + CStr(ADC("wDcOdNoTo")) + " and " + _
'                                  "OdSr= " + CStr(pv_NewValue)))
'    ErrMsg = "Invalid To Order Sr": Exit Sub
  Case Is = UCase("wWaitForAllBags")
    Call HlpList.PMCd("YN")
    Cancel = Not moCn.RecSeek("select PMCd from Param where PTyp= 'YN' and PMCd = '" + pv_NewValue + "'")
    ErrMsg = "Invalid 'Wait For All Bags' Option": Exit Sub
  Case Is = UCase("wDRGrp")
    Cancel = Not moCn.RecSeek("select PMCd from Param " + _
             "where PTyp= 'ORDRMGR' and PMCd = '" + pv_NewValue + "'")
    ErrMsg = "Invalid Group By Value": Exit Sub
  Case Is = UCase("wIRGrp")
    Cancel = Not moCn.RecSeek("select PMCd from Param " + _
             "where PTyp= 'ORDRMGR' and PMCd = '" + pv_NewValue + "'")
    ErrMsg = "Invalid Group By Value": Exit Sub
  Case Is = UCase("wILGrp")
    Cancel = Not moCn.RecSeek("select PMCd from Param " + _
             "where PTyp= 'ORDLABGR' and PMCd = '" + pv_NewValue + "'")
    ErrMsg = "Invalid Group By Value": Exit Sub
  Case Is = UCase("wDLGrp")
    Cancel = Not moCn.RecSeek("select PMCd from Param " + _
             "where PTyp= 'ORDLABGR' and PMCd = '" + pv_NewValue + "'")
    ErrMsg = "Invalid Group By Value": Exit Sub
  Case Is = UCase("wIDGrp")
    Cancel = Not moCn.RecSeek("select PMCd from Param " + _
             "where PTyp= 'ORDDSGGR' and PMCd = '" + pv_NewValue + "'")
    ErrMsg = "Invalid Group By Value": Exit Sub
  
  '*** (Jen 2.13)
  Case Is = UCase("wDdIdCustmMtchWt")
    Dim wb_Valid As Boolean, ws_RmCtg As String
    wb_Valid = True
    For i = 1 To GrdIR.Rows - 1
      If GrdIR.Value(i, "IrMainMet") = "Y" And (Not GrdIR.IsDel(i)) Then
        If Val(pv_NewValue) < GrdIR.Value(i, "IrRmZWt") Then wb_Valid = False: Exit For
      End If
    Next i
    If wb_Valid = False Then Cancel = True: ErrMsg = "Custom Match Weight Cannot Be Less Than Custom Weight Of Main Metal": Exit Sub
  '*** (Jen 2.13)
  
  ' ***** Manali 3.8.0 - Invoice Sort By
  Case Is = UCase("wInSrt")
    Cancel = pv_NewValue <> "" And Not moCn.RecSeek("Select PMCd from Param where PTyp= 'INSRT' and PMCd = '" + pv_NewValue + "'")
    ErrMsg = "Invalid Design Sort By Option": Exit Sub
    
  '3.11.0
  Case Is = UCase("wRefCnvRtYn")
    Cancel = Not moCn.RecSeek("select PMCd from Param where PTyp= 'YN' and " + _
                         "PMCd = '" + pv_NewValue + "'")
    ErrMsg = "Enter Refresh Conversion Rate (Y/N)": Exit Sub
    
  End Select
End Sub
Private Sub atxt_FldValid(Index As Integer, ByVal IdName As String, ByVal OldVal As Variant)
  '*** call the respective summary routines when the user tabs out of the Summary group fields
  
  If ADC(IdName) = OldVal Then Exit Sub
  Select Case UCase(IdName)
  Case Is = UCase("wDRGrp")
    Call DsgRmSumm
  Case Is = UCase("wIRGrp")
    Call InvRmSumm
  Case Is = UCase("wILGrp")
    Call InvLabSumm
  Case Is = UCase("wDLGrp")
    Call DsgLabSumm
  Case Is = UCase("wIDGrp")
    Call InvDsgSumm
  End Select
End Sub

Private Sub cmdGetFOB_Click()
  Call DispFra(GetFOB)  '18 showing Get FOB frame

End Sub

'18 updating new Fob value to all rows
Private Sub CmdGetFOBGo_Click()
  Dim wRsLab As MDORowSet, wRsSet As MDORowSet, wRsRmWithoutD As MDORowSet, wRsAllRm As MDORowSet
  Dim wTotWithoutD As Double, wDisVal As Double, wFOBVal As Double
  Dim wRow As Integer
 ' MWLib.BeginProcess Me, "Updating FOB"
  With GrdID
    For wRow = 1 To .Rows - 1
      .Row = wRow
      Set wRsRmWithoutD = moCn.OpenRes("Select IsNull(sum(IrRmIVal), 0) as qRmIVal from InvRm " + _
                    "where IrCoCd = '" + gs_CoCd + "' And IrTc = '" + ADC("wIdTc") + "' And IrYy = '" + ADC("wIdYy") + "' and " + _
                    "IrChr = '" + ADC("wIdChr") + "' And IrNo= " + CStr(ADC("wIdNo")) + " and IrSr= " + CStr(.Value(wRow, "IdSr")) + _
                    " and IrRmCtg <>'D' and IrPrtKey=(Case When '" + gs_Partition + "'='" + ctCurrPrtn + "' Then '" + ctCurrPrtn + "' Else IrPrtKey End) " + "")
                       
      Set wRsSet = moCn.OpenRes("Select IsNull(sum(IrSetIVal), 0) as qSetVal from InvRm " + _
                    "where IrCoCd = '" + gs_CoCd + "' And IrTc = '" + ADC("wIdTc") + "' And IrYy = '" + ADC("wIdYy") + "' and " + _
                    "IrChr = '" + ADC("wIdChr") + "' And IrNo= " + CStr(ADC("wIdNo")) + " and IrSr= " + CStr(.Value(wRow, "IdSr")) + _
                    " and IrSetSCd<> '' and IrPrtKey=(Case When '" + gs_Partition + "'='" + ctCurrPrtn + "' Then '" + ctCurrPrtn + "' Else IrPrtKey End) " + "")
         
      Set wRsLab = moCn.OpenRes("Select IsNull(sum(ILIVal), 0) as qLabVal from InvLab " + _
                    "Where ILCoCd= '" + gs_CoCd + "' and ILTc= '" + ADC("wIdTc") + "' And ILYy= '" + ADC("wIdYy") + "' and " + _
                    "ILChr= '" + ADC("wIdChr") + "' and ILNo= " + CStr(ADC("wIdNo")) + " and IlSr = " + CStr(.Value(wRow, "IdSr")) + _
                    " and IlPrtKey=(Case When '" + gs_Partition + "'='" + ctCurrPrtn + "' Then '" + ctCurrPrtn + "' Else IlPrtKey End) " + "")
                    
      Set wRsAllRm = moCn.OpenRes("Select IsNull(sum(IrRmIVal), 0) as qRmIVal from InvRm " + _
                    "where IrCoCd = '" + gs_CoCd + "' And IrTc = '" + ADC("wIdTc") + "' And IrYy = '" + ADC("wIdYy") + "' and " + _
                    "IrChr = '" + ADC("wIdChr") + "' And IrNo= " + CStr(ADC("wIdNo")) + " and IrSr= " + CStr(.Value(wRow, "IdSr")) + _
                    " and IrPrtKey=(Case When '" + gs_Partition + "'='" + ctCurrPrtn + "' Then '" + ctCurrPrtn + "' Else IrPrtKey End) " + "")
        
      wTotWithoutD = wRsRmWithoutD!qRmIVal + wRsSet!qSetVal + wRsLab!qLabVal
      wDisVal = wTotWithoutD * 3 / 100
      wFOBVal = (wRsAllRm!qRmIVal + wRsSet!qSetVal + wRsLab!qLabVal) - wDisVal
                                   
      .Value(wRow, "IDIFOB") = wFOBVal
      .Value(wRow, "IDZFOB") = wFOBVal
    Next wRow
  End With
  
'  MWLib.EndProcess Me
  Call DispFra(GetFOB)

End Sub

Private Sub cmdGetLotRtGo_Click()
  Dim Cancel As Boolean, ErrMsg As String
  If ms_PwdFlag = False Then DispMsg "Enter SEO Password to get Lot Rates", etError: Exit Sub

  If moCn.RecSeek("Select 'x' from InvZRm where IzCoCd= '" + gs_CoCd + "' and IzTc= '" + ADC("wIdTc") + "' and " + _
                  "IzYy=  '" + ADC("wIdYy") + "' and IzChr=  '" + ADC("wIdChr") + "' and " + _
                  "IzNo=  " + CStr(ADC("wIdNo")) + " and IzPrtKey='" + ctCurrPrtn + "' ") Then
   DispMsg "Cannot Refresh Rate As Custom Matching Has Been Done", etError: Exit Sub
  End If
  Cancel = Not moCn.RecSeek("select PMCd from Param where PTyp= 'YN' and " + _
                         "PMCd = '" + ADC("WGETLOTRTALLDCX") + "'")
  ErrMsg = "Enter All DCX (Y/N)"
  If Cancel = True Then DispMsg ErrMsg, etError: Exit Sub
  Cancel = Not moCn.RecSeek("select PMCd from Param where PTyp= 'YN' and " + _
                         "PMCd = '" + ADC("WGETLOTRTONLYCUST") + "'")
  ErrMsg = "Enter Y\N for Only Customer field"
  If Cancel = True Then DispMsg ErrMsg, etError: Exit Sub
  cmdGetLotRt.Enabled = False
  Call UpdateIrRtFromLotRt
  Call DispFra(GetLotRt)
End Sub

Private Sub CmdDRL_Click()

 '****** sachin 3.02 28-12-07 - Deletion Not allowed if InvDsg belongs to the previous partition
 If ADC("wIdPrtKey") <> ctCurrPrtn Then DispMsg "Cannot Delete Rm/Lab From Invoice of Previous Partition ", etError: Exit Sub

'****** Sachin 2.13.0 - Delete RM/Lab Button - 27-05-2006 ******
  If moCn.RecSeek("Select 'x' from InvZRm where IzCoCd= '" + gs_CoCd + "' and IzTc= '" + ADC("wIdTc") + "' and " + _
                  "IzYy=  '" + ADC("wIdYy") + "' and IzChr=  '" + ADC("wIdChr") + "' and " + _
                  "IzNo=  " + CStr(ADC("wIdNo")) + " and IzPrtKey='" + ctCurrPrtn + "' ") Then
    DispMsg "Cannot Delete RM/Labour As Custom Matching Has Been Done", etError: GrdID.SetFocus: Exit Sub
  End If

  GrdIR.DelAllRecs
  GrdIL.DelAllRecs
  GrdID.SetFocus
'****** Sachin 2.13.0 - Delete RM/Lab Button - 27-05-2006 ******

End Sub

Private Sub CmdRefRt_Click()
  If ADC("WIDPRTKEY") <> ctCurrPrtn Then DispMsg "Cannot Refresh Rate in Previous Partition Records", etError: Exit Sub
  
  If mf_DtLocked = True Then DispMsg "Cannot Refresh Rate in Locked Entry", etError: Exit Sub
  
  If moCn.RecSeek("Select 'x' from InvZRm where IzCoCd= '" + gs_CoCd + "' and IzTc= '" + ADC("wIdTc") + "' and " + _
                  "IzYy=  '" + ADC("wIdYy") + "' and IzChr=  '" + ADC("wIdChr") + "' and " + _
                  "IzNo=  " + CStr(ADC("wIdNo")) + " and IzPrtKey='" + ctCurrPrtn + "' ") Then
    DispMsg "Cannot Refresh Rate As Custom Matching Has Been Done", etError: Exit Sub
  End If
  
  If mInCls = "Y" Then DispMsg "Cannot Refresh Rate As Invoice Is Closed", etError: Exit Sub
  
  CmdRefRtGo.Enabled = True
  ADC("wRefCnvRtYn") = "N"  '3.11.0
  Call DispFra(RefRt)

End Sub

Private Sub CmdRefRtGo_Click()

  If ms_PwdFlag = False Then DispMsg "Enter SEO Password for Refresh Rates", etError: Exit Sub
  
  If ADC("wInLmgCst") = 0 Or ADC("wInLmgSal") = 0 Or ADC("wInLmpCst") = 0 Or ADC("wInLmpSal") = 0 Or _
     ADC("wInLmsCst") = 0 Or ADC("wInLmsSal") = 0 Or ADC("wInLmlCst") = 0 Or ADC("wInLmlSal") = 0 Or (ADC("wRefCnvRtYn") = "Y" And ADC("wInCnvRt") = 0) Then _
     DispMsg "Lme Value(s) Or Conversion Rate Cannot be Zero", etError: Exit Sub

  If moCn.RecSeek("Select 'x' from InvZRm where IzCoCd= '" + gs_CoCd + "' and IzTc= '" + ADC("wIdTc") + "' and " + _
                  "IzYy=  '" + ADC("wIdYy") + "' and IzChr=  '" + ADC("wIdChr") + "' and " + _
                  "IzNo=  " + CStr(ADC("wIdNo")) + " and IzPrtKey='" + ctCurrPrtn + "' ") Then
    DispMsg "Cannot Refresh Rate As Custom Matching Has Been Done", etError: Exit Sub
  End If

  CmdRefRtGo.Enabled = False
  Call RefreshRt
  Call DispFra(RefRt)
End Sub

Private Sub CmdGetLotRt_Click()
  txtGetLotRtSEOPwd = ""
  ms_PwdFlag = False
  Call DispFra(GetLotRt)
    
End Sub
Private Sub UpdateIrRtFromLotRt()
  Dim wRsRmMstRec As MDORowSet
  Dim wIrRmCtg As String
  Dim wi_D As Integer
  Dim wi_R As Integer
  Dim WLotRt As Double
 ' MWLib.BeginProcess Me, "Updating LotRt"        '4.1.3.0
  For wi_D = 1 To GrdID.Rows - 1
    GrdID.Row = wi_D
    If GrdID.Value(wi_D, "IdPrtKey") = ctCurrPrtn Then
      With GrdIR
        For wi_R = 1 To .Rows - 1
          If .Value(wi_R, "IrPrtKey") = ctCurrPrtn And .Value(wi_R, "IrLotNo") <> "" Then
            Set wRsRmMstRec = moCn.OpenRes("Select RmCtg,RmCmCd,RmQw,RmQwCst From RmMst Where RmCd='" + .Value(wi_R, "IrRmCd") + "' and RmPrtKey='" + ctCurrPrtn + "'")
            wIrRmCtg = wRsRmMstRec!RmCtg
            If (wRsRmMstRec.RecCount <= 0) Then DispMsg "Unexpected error failing to find RmCd in RmMst", etError: Exit Sub
            If ADC("WGETLOTRTALLDCX") = "Y" And (wIrRmCtg = "D" Or wIrRmCtg = "C" Or wIrRmCtg = "X") Then
              If (ADC("WGETLOTRTONLYCUST") = "N" Or wRsRmMstRec!RmCmCd <> "") Then
                WLotRt = GetRmLotRt(.Value(wi_R, "IrRmCd"), .Value(wi_R, "IrLotNo"), wRsRmMstRec!RmQw)
                If (WLotRt <> 0) Then
                .Value(wi_R, "IrRmIRt") = WLotRt
                .Value(wi_R, "IrRmZRt") = WLotRt
                End If
              End If
            End If
          End If
          GrdIR.SaveRec (wi_R)
        Next wi_R
      End With
    End If
  Next wi_D
'  MWLib.EndProcess Me
End Sub
Private Function GetRmLotRt(pRmCd As String, pLotNo As String, pRmQw As String) As Double
  Dim wSqlString As String
  Dim wInCmCd As String
  Dim wRtVal As Double
  Dim wResOfSeekIntoTxnd As MDORowSet
  wRtVal = 0
  wInCmCd = moCn.GetFldVal("Select InCmCd From InvHd Where InCoCd= '" + gs_CoCd + "' and " + _
           "InTc= '" + ADC("wIdTc") + "' And InYy= '" + ADC("wIdYy") + "' and " + _
           "InChr= '" + ADC("wIdChr") + "' and InNo= " + CStr(ADC("wIdNo")) + " and InPrtKey=(Case When '" + gs_Partition + "'='" + ctCurrPrtn + "' Then '" + ctCurrPrtn + "' Else InPrtKey End) " + _
           "")
  
  wSqlString = "Select TdRmQty, TdRmWt,TdPurRt,HBaseCurCd, CmCurCd,TCurrCd, TdRtByQw from Txnd " + _
              "              Join Txn on TIdNo = TdTIdNo " + _
                "              Join Head on HCoCd = '" + gs_CoCd + "'" + _
                "              Join CustMst on CmCtg  = '" + ms_CmCtg + "' and CmCd = '" + wInCmCd + "'" + _
                  " where TdIdNo = (Select min(TdIdNo) From Txnd " + _
                  " where ((TdFrRmLocTyp = 'S' and TdFrRmDc = 'D') or (TdToRmLocTyp ='S' and TdToRmDc = 'D')) and TdLotNo='" + pLotNo + "')"
  Set wResOfSeekIntoTxnd = moCn.OpenRes(wSqlString)
  
  If wResOfSeekIntoTxnd.RecCount > 0 Then
    If (wResOfSeekIntoTxnd!HBaseCurCd = wResOfSeekIntoTxnd!CmCurCd And wResOfSeekIntoTxnd!TCurrCd = wResOfSeekIntoTxnd!CmCurCd) Then
      If wResOfSeekIntoTxnd!TdRtByQW = pRmQw Then
        wRtVal = wResOfSeekIntoTxnd!TdPurRt
      Else
        If wResOfSeekIntoTxnd!TdRtByQW = "W" And wResOfSeekIntoTxnd!TdRmQty <> 0 Then wRtVal = wResOfSeekIntoTxnd!TdPurRt * wResOfSeekIntoTxnd!TdRmWt / wResOfSeekIntoTxnd!TdRmQty
        If wResOfSeekIntoTxnd!TdRtByQW = "Q" And wResOfSeekIntoTxnd!TdRmWt <> 0 Then wRtVal = wResOfSeekIntoTxnd!TdPurRt * wResOfSeekIntoTxnd!TdRmQty / wResOfSeekIntoTxnd!TdRmWt
      End If
    End If
  End If
  GetRmLotRt = wRtVal
End Function
Private Function GetRmLotRtOld(pRmCd As String, pLotNo As String, pRmQw As String) As Double
  Dim wSqlString As String
  Dim wInCmCd As String
  Dim wRtVal As Double
  Dim wResOfSeekForLotMstRt  As MDORowSet
  Dim wResOfInvHd As MDORowSet
  wRtVal = 0
  Set wResOfInvHd = moCn.OpenRes("Select INMULDIV, INCNVFCT From InvHd Where InCoCd= '" + gs_CoCd + "' and " + _
           "InTc= '" + ADC("wIdTc") + "' And InYy= '" + ADC("wIdYy") + "' and " + _
           "InChr= '" + ADC("wIdChr") + "' and InNo= " + CStr(ADC("wIdNo")) + " and InPrtKey=(Case When '" + gs_Partition + "'='" + ctCurrPrtn + "' Then '" + ctCurrPrtn + "' Else InPrtKey End) " + _
           "")
           
  If wResOfInvHd.RecCount > 0 Then
    If wResOfInvHd!InCnvFct <> 0 Then
      wSqlString = "Select " + IIF(pRmQw = "Q", "LotBaseRtByQ", "LotBaseRtByW") + _
                        IIF(wResOfInvHd!InMulDiv = "M", "*", "/") + CStr(wResOfInvHd!InCnvFct) + " as CalcLotRt from LotMst " + _
                 " where LotNo='" + pLotNo + "'"
      Set wResOfSeekForLotMstRt = moCn.OpenRes(wSqlString)
    
      If wResOfSeekForLotMstRt.RecCount > 0 Then
          wRtVal = wResOfSeekForLotMstRt!CalcLotRt
      End If
    End If
  End If
  GetRmLotRtOld = wRtVal
End Function
'6.1_37
Private Sub CmdSc_Click()
  Call DispFra(Scope)
End Sub

' Manoj 2.10.0
Private Sub CmdSH_Click()
  '*** Display or hide the Summary frame depending on the toggle state
  '*** Call the Design Sale History sub procedure
  If FraSalHst.Visible = False Then
    If GrdID.HalfValidate Then
      GrdID.SetFocus
      Exit Sub
    End If
    Call DispFra(SalHst)
    Call DsgSalHst(GrdID.Value(GrdID.Row, "IdDmCd"))
  Else
    Call DispFra(SalHst)
  End If
End Sub


Private Sub GrdHD_ColWhen(ByVal RowNum As Integer, ByVal ColName As String, Cancel As Boolean, ErrMsg As String, ByVal Action As MwfCtl.en_fgAction)
  Cancel = True
End Sub
Private Sub GrdHD_RowWhen(ByVal RowNum As Integer)
  Dim wSqlStr As String, wRsRm As MDORowSet, wRow As Integer, wRsLab As MDORowSet
  If RowNum > 0 Then
     GrdHD.StartCopy
     GrdHR.StartCopy
     GrdHL.StartCopy
    ' Zubin 212
    '****** Sachin 3.02 - Summary from Corresponding Partition Only
    wSqlStr = "Select IrSrNo, IrRmCd, IrLotNo, IrRmSz, IrRmSz2, IrRmSz3, IrMainMet, " & _
              " (IrRmQty/IdQty) As IrRmUQty, (IrRmIWt/IdQty) As IrRmUWt, IrRmIRt, (IrRmIVal/IdQty) As IrRmIUVal," & _
              " IrSetSCd, IrSetIRt, (IrSetIVal/IdQty) As IrSetIUVal from InvDsg Join InvRm On" & _
              " IdIdNo=IrIdIdNo and IdCoCd = IrCoCd And IdTc = IrTc And IdYy = IrYy And IdChr = IrChr And IdNo = IrNo And IdSr = IrSr And IdPrtKey=IrPrtKey " & _
              " Where IdCoCd = '" & gs_CoCd & "' And IdTc = '" & GrdHD.Value(RowNum, "whIdTc") & "' And IdYy = '" & GrdHD.Value(RowNum, "whIdYy") & "'" & _
              " And IdChr = '" & GrdHD.Value(RowNum, "whIdChr") & "' And IdNo = " & GrdHD.Value(RowNum, "whIdNo") & " And IdSr = " & GrdHD.Value(RowNum, "whIdSr") & " And IdPrtKey=(Case When '" + ADC("wIdPrtKey") + "'='" + ctCurrPrtn + "' Then '" + ctCurrPrtn + "' Else IdPrtKey End)"
    Set wRsRm = moCn.OpenResultset(wSqlStr)
    If wRsRm.RecCount > 0 Then
      
      wRsRm.MoveFirst
      With GrdHR
      .Rows = 1
        Do While Not (wRsRm.BOF Or wRsRm.EOF)
          .AddItem
          wRow = .Rows - 1
'          .Value(wRow, "whIrSrNo") = wRsRm.FldValue("IrSrNo")
          .Value(wRow, "whIrRmCd") = wRsRm.FldValue("IrRmCd")
          ' Zubin 212
          .Value(wRow, "whIrLotNo") = wRsRm.FldValue("IrLotNo")
          ' Zubin 212
          .Value(wRow, "whIrRmSz") = wRsRm.FldValue("IrRmSz")
          .Value(wRow, "whIrRmSz2") = wRsRm.FldValue("IrRmSz2")
          .Value(wRow, "whIrRmSz3") = wRsRm.FldValue("IrRmSz3")
          .Value(wRow, "whIrQty") = wRsRm.FldValue("IrRmUQty")
          .Value(wRow, "whIrRmWt") = wRsRm.FldValue("IrRmUWt")
          .Value(wRow, "whIrRmRt") = wRsRm.FldValue("IrRmIRt")
          .Value(wRow, "whIrRmVal") = wRsRm.FldValue("IrRmIUVal")
          .Value(wRow, "whIrSetSCd") = wRsRm.FldValue("IrSetSCd")
          .Value(wRow, "whIrSetRt") = wRsRm.FldValue("IrSetIRt")
          .Value(wRow, "whIrSetVal") = wRsRm.FldValue("IrSetIUVal")
          .Value(wRow, "whIrMainMet") = wRsRm.FldValue("IrMainMet")
          wRsRm.MoveNext
        Loop
      End With
    End If
    Set wRsRm = Nothing
    '****** Sachin 3.02 - Summary from the Corresponding partition
    wSqlStr = "Select IlSrNo, IlMCd, IlSCd, ILQw, " & _
              " (IlQty/IdQty) As IlUQty, IlIRt, (IlIVal/IdQty) As IlIUVal" & _
              " From InvDsg Join InvLab On" & _
              " IdIdNo=IlIdIdNo and IdCoCd = IlCoCd And IdTc = IlTc And IdYy = IlYy And IdChr = IlChr And IdNo = IlNo And IdSr = IlSr And IdPrtKey=IlPrtKey" & _
              " Where IdCoCd = '" & gs_CoCd & "' And IdTc = '" & GrdHD.Value(RowNum, "whIdTc") & "' And IdYy = '" & GrdHD.Value(RowNum, "whIdYy") & "'" & _
              " And IdChr = '" & GrdHD.Value(RowNum, "whIdChr") & "' And IdNo = " & GrdHD.Value(RowNum, "whIdNo") & " And IdSr = " & GrdHD.Value(RowNum, "whIdSr") & " And IdPrtKey=(Case When '" + ADC("wIdPrtKey") + "'='" + ctCurrPrtn + "' Then '" + ctCurrPrtn + "' Else IdPrtKey End)"
    Set wRsLab = moCn.OpenResultset(wSqlStr)
    If wRsLab.RecCount > 0 Then
      
      wRsLab.MoveFirst
      With GrdHL
      .Rows = 1
        Do While Not (wRsLab.BOF Or wRsLab.EOF)
          .AddItem
          wRow = .Rows - 1
          .Value(wRow, "WHILLMCD") = wRsLab.FldValue("IlMCd")
          .Value(wRow, "WHILLSCD") = wRsLab.FldValue("IlSCd")
          .Value(wRow, "WHILQW") = wRsLab.FldValue("IlQw")
          .Value(wRow, "WHILQTY") = wRsLab.FldValue("IlUQty")
          .Value(wRow, "WHILRATE") = wRsLab.FldValue("IlIRt")
          .Value(wRow, "WHILIVAL") = wRsLab.FldValue("IlIUVal")
          wRsLab.MoveNext
        Loop
      End With
    End If
    Set wRsLab = Nothing
    GrdHD.EndCopy
    GrdHR.EndCopy
    GrdHL.EndCopy
  End If
End Sub
Private Sub GrdHR_ColWhen(ByVal RowNum As Integer, ByVal ColName As String, Cancel As Boolean, ErrMsg As String, ByVal Action As MwfCtl.en_fgAction)
  Cancel = True
End Sub
Private Sub GrdHL_ColWhen(ByVal RowNum As Integer, ByVal ColName As String, Cancel As Boolean, ErrMsg As String, ByVal Action As MwfCtl.en_fgAction)
  Cancel = True
End Sub
' Manoj 2.10.0

Private Sub GrdIB_OnDelete(ByVal RowNum As Integer, Cancel As Boolean, ErrMsg As String)
  '*** (Jen 3.2.0)
  '****** Sachin 3.02.0 28-12-07 - prtkey added below
  If GrdIB.Mode = fgmnorm Then
    If moCn.RecSeek("Select 'x' from InvZRm where IzCoCd= '" + gs_CoCd + "' and IzTc= '" + ADC("wIdTc") + "' and " + _
                    "IzYy=  '" + ADC("wIdYy") + "' and IzChr=  '" + ADC("wIdChr") + "' and " + _
                    "IzNo=  " + CStr(ADC("wIdNo")) + " and IzPrtKey= '" + ctCurrPrtn + "' ") Then
      Cancel = True: ErrMsg = "Cannot Edit As Custom Matching Has Been Done": Exit Sub
    End If
  End If
  '*** (Jen 3.2.0)
End Sub

Private Sub GrdID_GotFocus()
'New Code
  '*** Enable all the Command Buttons associated with the Inv Design level when the
      'focus is in the Design Grid (Enable only if the Invoice is not closed)
  '*** Disable the Make Invoice option when the cursor is in the Design grid

 
  If mInCls = "Y" Or mf_DtLocked Then
    GrdID.AllowAdd = False: GrdID.AllowDelete = False
    CmdDC.Enabled = False
    CmdCRL.Enabled = False
    CmdDRL.Enabled = False  '****** Sachin 2.13.0 - 27-05-2006 ******
    CmdInvDel.Enabled = False   '4.1.2 - Complete Invoice Delete
  ElseIf mInCls = "N" Then
 
    GrdID.AllowAdd = True: GrdID.AllowDelete = True
    CmdDC.Enabled = True
    CmdCRL.Enabled = True
    CmdDRL.Enabled = True   '****** Sachin 2.13.0 - 27-05-2006 ******
    CmdInvDel.Enabled = True   '4.1.2 - Complete Invoice Delete
  End If
  
  If GrdID.Rows > 1 Then
     CmdDD.Enabled = True
      CmdIL.Enabled = True
      CmdIB.Enabled = True
      CmdSumm.Enabled = True
      CmdSH.Enabled = True
    
      CmdMI.Enabled = False
  End If
End Sub
Private Sub GrdID_When()
'Old Code
  '*** Enable all the Command Buttons associated with the Inv Design level when the
      'focus is in the Design Grid (Enable only if the Invoice is not closed)
  '*** Disable the Make Invoice option when the cursor is in the Design grid
  
'  CmdDD.Enabled = True
'  If mInCls = "Y" Then
'    CmdDC.Enabled = False
'    CmdCRL.Enabled = False
'  ElseIf mInCls = "N" Then
'    CmdDC.Enabled = True
'    CmdCRL.Enabled = True
'  End If
'  CmdIL.Enabled = True
'  CmdIB.Enabled = True
'  CmdSumm.Enabled = True
'
'  CmdMI.Enabled = False
                                

End Sub
Private Sub GrdID_RowWhen(ByVal RowNum As Integer)
  '*** If New Row then IdSr= max(IdSr) of the grid + 1 and
       'mIdHZValEqInv = default HZValEqInv from Head file
  
  Dim w3DPth As String, wHskPth  As String
  With GrdID
    '%%% If Not GrdID.IsNew(RowNum) Then
    If Not GrdID.IsNew(RowNum) And GrdID.Mode = fgmnorm Then
      FraPic3D.Left = wFraPic3DLft: FraPic3D.Top = wFraPic3DTop
      FraPic3D.Height = wFraPic3DHt: FraPic3D.Width = wFraPic3DWdt
      FraPicLD.Left = wFraPicHskLft: FraPicLD.Top = wFraPicHskTop
      FraPicLD.Height = wFraPicHskHt: FraPicLD.Width = wFraPicHskWdt
      '*** Jay 3.2.0 [OdPicNm]
      w3DPth = GetPictPath("DM", False, .Value(RowNum, "IdDmCd"), .Value(RowNum, "IdDmSz"), gs_CoCd, .Value(RowNum, "IdExpOdTc"), .Value(RowNum, "IdExpOdYy"), .Value(RowNum, "IdExpOdChr"), .Value(RowNum, "IdExpOdNo"), .Value(RowNum, "IdExpOdSr"))
      wHskPth = GetPictPath("DM", True, .Value(RowNum, "IdDmCd"), .Value(RowNum, "IdDmSz"), gs_CoCd, .Value(RowNum, "IdExpOdTc"), .Value(RowNum, "IdExpOdYy"), .Value(RowNum, "IdExpOdChr"), .Value(RowNum, "IdExpOdNo"), .Value(RowNum, "IdExpOdSr"))
      
      'w3DPth = GetPictPath("DM", False, .Value(RowNum, "IdDmCd"))
      'wHskPth = GetPictPath("DM", True, .Value(RowNum, "IdDmCd"))
      '*** Jay 3.2.0 [OdPicNm]
      ' **** Zubin 211 **** '
      'w3DPth = GetPictPath("DM", .Value(RowNum, "IdDmCd"), , False)
      'wHskPth = GetPictPath("DM", .Value(RowNum, "IdDmCd"), , True)
      ' **** Zubin 211 **** '
      
      Call ShowPic(wPic3DHt, wPic3DWdt, w3DPth, Pic3D)
      Call ShowPic(wPicHskHt, wPicHskWdt, wHskPth, PicLD)
    End If
    
    If GrdID.IsNew(RowNum) Then
      If .Value(RowNum, "IdSr") = 0 Then .Value(RowNum, "IdSr") = .MaxVal("IdSr") + 1
      .Value(RowNum, "IdPrtKey") = ctCurrPrtn   '****** Sachin 3.02 - 28-12-07
      If .Value(RowNum, "IdHZValEqInv") = "" Then .Value(RowNum, "IdHZValEqInv") = mIdHZValEqInv
      .Value(RowNum, "IdTrayNo") = 0
    End If
    
  End With
End Sub
Private Sub GrdID_ColWhen(ByVal RowNum As Integer, ByVal ColName As String, Cancel As Boolean, ErrMsg As String, ByVal Action As MwfCtl.en_fgAction)
  '*** Do not allow the user to edit any field if InCls = 'Y' in InvHd Table
  '*** Give help for IdExpOdTc, IdExpOdYy, IdExpOdChr, IdExpOdNo, IdExpOdSr
  '*** Do not allow the user to edit the Design Code
  
  '****** Sachin 3.02 - Locking InvDsg Grid for Previous Partition records.
  If ADC("WIDPRTKEY") <> ctCurrPrtn Then Cancel = True: ErrMsg = "Cannot Edit Previous Partition Invoices": Exit Sub
  
  If mf_DtLocked = True Then Cancel = True: ErrMsg = "Cannot Edit Locked Entry": Exit Sub
  If mInCls = "Y" Then Cancel = True: ErrMsg = "Cannot Edit As Invoice Is Closed": Exit Sub
  With GrdID
    Select Case ColName
    Case Is = UCase("IdExpOdTc")
      Call HlpList.PMCd("TC", "'SO'")
    Case Is = UCase("IdExpOdYy")
      Call HlpList.PMCd("YY")
    Case Is = UCase("IdExpOdChr")
      Call HlpList.vPSCd(gs_CoCd, "CHR", .Value(RowNum, "IdExpOdTc"))
    Case Is = UCase("IdExpOdNo")
      Call HlpList.OmNo(gs_CoCd, .Value(RowNum, "IdExpOdTc"), .Value(RowNum, "IdExpOdYy"), .Value(RowNum, "IdExpOdChr"))
    Case Is = UCase("IdExpOdSr")
      Call HlpList.OdSr(gs_CoCd, .Value(RowNum, "IdExpOdTc"), .Value(RowNum, "IdExpOdYy"), .Value(RowNum, "IdExpOdChr"), .Value(RowNum, "IdExpOdNo"))
    Case Is = UCase("IdDmCd")
      '****** This field could have been kept read only; but it is kept as 'Normal' because
             'if the user is entering the Source Order (Exp Order) manually, he will have to
             'press 'shift+tab' instead of 'tab' if the field is readonly
      Cancel = True: ErrMsg = "Cannot Edit Design Code": Exit Sub
    End Select
  End With
End Sub
Private Sub GrdIdRecalcHelper(ByVal ColumnsList As String, ByVal RowNum As Integer, ByVal SourceRowNum As Double, ByVal SourceColName As String)
    Dim wf_IdAVal As Boolean, wf_IdIVal As Boolean, wf_IDZVal As Boolean
    Dim wf_IdAGldAsWt As Boolean, wf_IdIGldAsWt As Boolean, wf_IdZGldAsWt As Boolean
    Dim wf_IdZLabAsWt  As Boolean
    Dim wf_IdALabAsWt As Boolean
    Dim wf_IdILabAsWt As Boolean
    Dim wf_IdCustmMtchWt As Boolean
    
    Dim wf_wIdADiaWt As Boolean, wf_wIdIDiaWt As Boolean, wf_wIdZDiaWt As Boolean
    Dim wf_wIdACsWt As Boolean, wf_wIdICsWt As Boolean, wf_wIdZCsWt As Boolean
    
    Dim wRmCtg As String
    Dim wARmVal As Double, wIRmVal As Double, wZRmVal As Double, wALabVal As Double, wILabVal As Double, wZLabVal As Double

    Dim wf_IdADiaWt, wf_IDIDiaWt, wf_IdZDiaWt, wf_IdACSWt, wf_IdICsWt, wf_IdZCsWt, wf_IdGldAs As Boolean
    Dim wf_wAGldAsWt  As Boolean
    Dim wf_wIGldAsWt As Boolean
    Dim wf_wZGldAsWt As Boolean
    
    Dim wAGldAsWt As Double
    Dim wIGldAsWt As Double
    Dim wZGldAsWt As Double
    Dim wALabAsWt As Double
    Dim wILabAsWt As Double
    Dim wZLabAsWt As Double
    Dim wIdADiaWt As Double, wIDIDiaWt As Double, wIdZDiaWt As Double, wIdACSWt     As Double, wIdICsWt As Double, wIdZCsWt As Double, wd_IdGldAs As Double, wd_IdCustmMtchWt As Double, wIdCustmMtchWt As Double, wd_IdGldLs As Double
   
    
    ColumnsList = "," + UCase(ColumnsList) + ","
    
    wf_IdAVal = IIF(InStr(1, "," + ColumnsList + ",", UCase(",IdAVal,")) > 0, True, False)
    wf_IdIVal = IIF(InStr(1, "," + ColumnsList + ",", UCase(",IdIVal,")) > 0, True, False)
    wf_IDZVal = IIF(InStr(1, "," + ColumnsList + ",", UCase(",IdZVal,")) > 0, True, False)
    wf_IdAGldAsWt = IIF(InStr(1, "," + ColumnsList + ",", UCase(",IdAGldAsWt,")) > 0, True, False)
    wf_IdIGldAsWt = IIF(InStr(1, "," + ColumnsList + ",", UCase(",IdIGldAsWt,")) > 0, True, False)
    wf_IdZGldAsWt = IIF(InStr(1, "," + ColumnsList + ",", UCase(",IdZGldAsWt,")) > 0, True, False)
    wf_IdALabAsWt = IIF(InStr(1, "," + ColumnsList + ",", UCase(",IdALabAsWt,")) > 0, True, False)
    wf_IdILabAsWt = IIF(InStr(1, "," + ColumnsList + ",", UCase(",IdILabAsWt,")) > 0, True, False)
    wf_IdZLabAsWt = IIF(InStr(1, "," + ColumnsList + ",", UCase(",IdZLabAsWt,")) > 0, True, False)
    wf_IdCustmMtchWt = IIF(InStr(1, "," + ColumnsList + ",", UCase(",IdCustmMtchWt,")) > 0, True, False)
    wf_wIdADiaWt = IIF(InStr(1, "," + ColumnsList + ",", UCase(",wIdADiaWt,")) > 0, True, False)
    wf_wIdIDiaWt = IIF(InStr(1, "," + ColumnsList + ",", UCase(",wIdIDiaWt,")) > 0, True, False)
    wf_wIdZDiaWt = IIF(InStr(1, "," + ColumnsList + ",", UCase(",wIdZDiaWt,")) > 0, True, False)
    wf_wIdACsWt = IIF(InStr(1, "," + ColumnsList + ",", UCase(",wIdACSWt,")) > 0, True, False)
    wf_wIdICsWt = IIF(InStr(1, "," + ColumnsList + ",", UCase(",wIdICSWt,")) > 0, True, False)
    wf_wIdZCsWt = IIF(InStr(1, "," + ColumnsList + ",", UCase(",wIdZCsWt,")) > 0, True, False)
    
    
    wIdZDiaWt = 0#
    wIDIDiaWt = 0#
    wIdADiaWt = 0#
    wIdZCsWt = 0#
    wIdICsWt = 0#
    wIdACSWt = 0#
    wd_IdGldAs = 0
    wd_IdCustmMtchWt = 0
    If (wf_IdIVal) Then
          GrdID.Value(RowNum, "IdIVal") = 0
          If GrdID.Value(RowNum, "IdFixPrc") = "Y" Then
            GrdID.Value(RowNum, "IdIVal") = GrdID.Value(RowNum, "IdQty") * _
                                          moCn.GetFldVal("Select OdSalPrc from OrdDsg where " + _
                                          "OdCoCd= '" + gs_CoCd + "' and " + _
                                          "OdTc= '" + GrdID.Value(RowNum, "IdExpOdTc") + "' and " + _
                                          "OdYy= '" + GrdID.Value(RowNum, "IdExpOdYy") + "' and " + _
                                          "OdChr= '" + GrdID.Value(RowNum, "IdExpOdChr") + "' and " + _
                                          "OdNo= " + CStr(GrdID.Value(RowNum, "IdExpOdNo")) + " and " + _
                                          "OdSr= " + CStr(GrdID.Value(RowNum, "IdExpOdSr")) + IIF(ADC("wIdPrtKey") = ctCurrPrtn, " and OdPrtKey='" + ctCurrPrtn + "' ", ""))
          ElseIf GrdID.Value(RowNum, "IdFixPrc") = "N" Then
            ' Calculation continues below where it will loop thru Rm and Lab and compute values

          End If
    End If
      
    Dim wf_IDZValCalc As Boolean
    If wf_IDZVal Then
        If GrdID.Value(RowNum, "IdHZValEqInv") = "Y" Then
            GrdID.Value(RowNum, "IdZVal") = 0
            GrdID.Value(RowNum, "IdZVal") = GrdID.Value(RowNum, "IdiVal")
            wf_IDZValCalc = False
          ElseIf GrdID.Value(RowNum, "IdHZValEqInv") = "N" And UCase(SourceColName) <> UCase("IdiVal") Then
            If GrdID.Value(RowNum, "IdFixPrc") = "Y" Then
              GrdID.Value(RowNum, "IdZVal") = GrdID.Value(RowNum, "IdQty") * _
                                            moCn.GetFldVal("Select OdSalPrc from OrdDsg where " + _
                                            "OdCoCd= '" + gs_CoCd + "' and " + _
                                            "OdTc= '" + GrdID.Value(RowNum, "IdExpOdTc") + "' and " + _
                                            "OdYy= '" + GrdID.Value(RowNum, "IdExpOdYy") + "' and " + _
                                            "OdChr= '" + GrdID.Value(RowNum, "IdExpOdChr") + "' and " + _
                                            "OdNo= " + CStr(GrdID.Value(RowNum, "IdExpOdNo")) + " and " + _
                                            "OdSr= " + CStr(GrdID.Value(RowNum, "IdExpOdSr")) + IIF(ADC("wIdPrtKey") = ctCurrPrtn, " and OdPrtKey='" + ctCurrPrtn + "' ", ""))
              wf_IDZValCalc = False
            ElseIf GrdID.Value(RowNum, "IdFixPrc") = "N" Then
              wf_IDZValCalc = True
              ' Calculation continues below where it will loop thru Rm and Lab and calc the value
            End If
          End If
    End If
    Dim i As Integer
        For i = 1 To GrdIR.Rows - 1
          If Not GrdIR.IsDel(i) Then
              If wf_IdAVal Then _
              wARmVal = wARmVal + GrdIR.Value(i, "IrRmAVal") + GrdIR.Value(i, "IrSetAVal")
            
              If wf_IdIVal Then _
              wIRmVal = wIRmVal + GrdIR.Value(i, "IrRmIVal") + GrdIR.Value(i, "IrSetIVal")
              
              If wf_IDZVal Then _
              wZRmVal = wZRmVal + GrdIR.Value(i, "IrRmZVal") + GrdIR.Value(i, "IrSetZVal")
            
              If wf_IdCustmMtchWt Or wf_IdADiaWt Or wf_IDIDiaWt Or wf_IdZDiaWt Then
              wRmCtg = moCn.GetFldVal("Select RmCtg From RmMst where RmCd= '" + GrdIR.Value(i, "IrRmCd") + "' and RmPrtKey=(Case When '" + ADC("wIdPrtKey") + "'='" + ctCurrPrtn + "' Then '" + ctCurrPrtn + "' Else RmPrtKey End) ")
              
              If wf_wIdADiaWt Then _
                wIdADiaWt = wIdADiaWt + IIF(wRmCtg = "D", GrdIR.Value(i, "IrRmAWt"), 0)
              
              If wf_wIdIDiaWt Then _
                wIDIDiaWt = wIDIDiaWt + IIF(wRmCtg = "D", GrdIR.Value(i, "IrRmIWt"), 0)
                
              If wf_wIdZDiaWt Then _
                wIdZDiaWt = wIdZDiaWt + IIF(wRmCtg = "D", GrdIR.Value(i, "IrRmZWt"), 0)
                
              If wf_wIdACsWt Then _
                wIdACSWt = wIdACSWt + IIF(wRmCtg = "C", GrdIR.Value(i, "IrRmAWt"), 0)
              
              If wf_wIdICsWt Then _
                wIdICsWt = wIdICsWt + IIF(wRmCtg = "D", GrdIR.Value(i, "IrRmIWt"), 0)
                
              If wf_wIdZCsWt Then _
                wIdZCsWt = wIdZCsWt + IIF(wRmCtg = "C", GrdIR.Value(i, "IrRmZWt"), 0)
            
              If wf_IdCustmMtchWt Then
                If GrdIR.Value(i, "IrMainMet") = "Y" And (Not GrdIR.IsDel(i)) Then
                  If wRmCtg = "G" Or wRmCtg = "P" Or wRmCtg = "S" Or wRmCtg = "L" Then
                    wd_IdGldLs = GrdID.Value(RowNum, "IdGldLs")
                  End If
                  wd_IdCustmMtchWt = GrdIR.Value(i, "IrRmZWt") * (1 + wd_IdGldLs / 100)
                End If
              End If
             End If
          End If
     Next i
  
     For i = 1 To GrdIL.Rows - 1
        If Not GrdIL.IsDel(i) Then
          If wf_IdAVal Then _
            wALabVal = wALabVal + GrdIL.Value(i, "ilAVal")
          
          If wf_IdIVal Then _
            wILabVal = wILabVal + GrdIL.Value(i, "ilIVal")
            
          If wf_IDZVal Then _
            wZLabVal = wZLabVal + GrdIL.Value(i, "ilZVal")
        End If
        
      Next i
    
    If (wf_IdAVal) Then
        GrdID.Value(RowNum, "IdAVal") = (wARmVal + wALabVal)
    End If

    If (wf_IdIVal) Then
          If GrdID.Value(RowNum, "IdFixPrc") = "N" Then _
            GrdID.Value(RowNum, "IdIVal") = (wIRmVal + wILabVal)
    End If
    
    If wf_IDZVal Then
      If wf_IDZValCalc Then _
          GrdID.Value(RowNum, "IdZVal") = (wZRmVal + wZLabVal)
    End If
    'sri new
    If wf_IDZVal Then
        If GrdID.Value(RowNum, "IdHZValEqInv") = "Y" Then
            GrdID.Value(RowNum, "IdZVal") = 0
            GrdID.Value(RowNum, "IdZVal") = GrdID.Value(RowNum, "IdiVal")
            wf_IDZValCalc = False
        End If
    End If
    Dim AsColumnsList As String
    
    AsColumnsList = IIF(wf_IdAGldAsWt, "A", "")
    AsColumnsList = AsColumnsList + IIF(AsColumnsList = "", "", ",") + IIF(wf_IdIGldAsWt = True, "I", "")
    AsColumnsList = AsColumnsList + IIF(AsColumnsList = "", "", ",") + IIF(wf_IdZGldAsWt = True, "Z", "")
    
    Call GetGldAsWt(RowNum, AsColumnsList, wAGldAsWt, wIGldAsWt, wZGldAsWt, SourceRowNum)
    
    If wf_IdAGldAsWt Then _
      GrdID.Value(RowNum, "IdAGldAsWt") = wAGldAsWt
      
    If wf_IdIGldAsWt Then _
      GrdID.Value(RowNum, "IdIGldAsWt") = wIGldAsWt
      
    If wf_IdZGldAsWt Then _
      GrdID.Value(RowNum, "IdZGldAsWt") = wZGldAsWt
      
    
    AsColumnsList = IIF(wf_IdALabAsWt = True, "A", "")
    AsColumnsList = AsColumnsList + IIF(AsColumnsList = "", "", ",") + IIF(wf_IdALabAsWt = True, "I", "")
    AsColumnsList = AsColumnsList + IIF(AsColumnsList = "", "", ",") + IIF(wf_IdZLabAsWt = True, "Z", "")
    
    Call GetLabAsWt(RowNum, ColumnsList, wALabAsWt, wILabAsWt, wZLabAsWt)
    
    If wf_IdALabAsWt Then _
      GrdID.Value(RowNum, "IdALabAsWt") = wALabAsWt
      
    If wf_IdILabAsWt And GrdID.Value(RowNum, "IdLabWtFrOrd") = "N" Then _
      GrdID.Value(RowNum, "IdILabAsWt") = wILabAsWt
      
    If wf_IdZLabAsWt And GrdID.Value(RowNum, "IdLabWtFrOrd") = "N" Then _
      GrdID.Value(RowNum, "IdZLabAsWt") = wZLabAsWt
      
    If wf_IdCustmMtchWt Then _
      GrdID.Value(GrdID.Row, "IdCustmMtchWt") = wd_IdCustmMtchWt
    
    If wf_wIdADiaWt Then _
        GrdID.Value(GrdID.Row, "wIdADiaWt") = wIdADiaWt
        
    If wf_wIdIDiaWt Then _
        GrdID.Value(GrdID.Row, "wIdIDiaWt") = wIDIDiaWt

    If wf_wIdZDiaWt Then _
        GrdID.Value(GrdID.Row, "wIdZDiaWt") = wIdZDiaWt

    If wf_wIdACsWt Then _
        GrdID.Value(GrdID.Row, "wIdACsWt") = wIdACSWt
        
    If wf_wIdICsWt Then _
        GrdID.Value(GrdID.Row, "wIdICsWt") = wIdICsWt

    If wf_wIdZCsWt Then _
        GrdID.Value(GrdID.Row, "wIdZCsWt") = wIdZCsWt
End Sub
Private Sub GrdID_ReCalc(ByVal RowNum As Integer, ByVal ColName As String, ByVal OldVal As Variant, ByVal SourceName As String, ByVal SourceColName As String)
  Dim i As Integer, wSourceGrdMode As en_fgMode, wGrdIdRow As Single
  Dim wf_IdFOB As Double
  Dim wDiaWt As Single, wCSWt As Single, wRmCtg As String      ' **** Manali 3.5.0 - DHC & CHC
  Dim wOdMulBy As Double    '3.11.0
  
  With GrdID
    Select Case ColName
    Case Is = UCase("IdExpCmCd")    '*** recalc on IdExpOdTc, IdExpOdYy, IdExpOdChr, IdExpOdNo
      If .Mode = fgmnorm Then
        .Value(RowNum, "IdExpCmCd") = moCn.GetFldVal("Select OmCmCd from OrdMst where " + _
                                      "OmCoCd= '" + gs_CoCd + "' and " + _
                                      "OmTc= '" + .Value(RowNum, "IdExpOdTc") + "' and " + _
                                      "OmYy= '" + .Value(RowNum, "IdExpOdYy") + "' and " + _
                                      "OmChr= '" + .Value(RowNum, "IdExpOdChr") + "' and " + _
                                      "OmNo= " + CStr(.Value(RowNum, "IdExpOdNo")) + IIF(ADC("wIdPrtKey") = ctCurrPrtn, " and OmPrtKey='" + ctCurrPrtn + "' ", ""))
      End If
    Case Is = UCase("IdDmCd")       '*** recalc on IdExpOdTc, IdExpOdYy, IdExpOdChr, IdExpOdNo, IdExpOdSr
      If .Mode = fgmnorm Then
        'Or .Mode = fgmCopy
        '****** Sachin 3.02 28-12-07 - Current Partition
        .Value(RowNum, "IdDmCd") = moCn.GetFldVal("Select OdDmCd from OrdDsg where " + _
                                   "OdCoCd= '" + gs_CoCd + "' and " + _
                                   "OdTc= '" + .Value(RowNum, "IdExpOdTc") + "' and " + _
                                   "OdYy= '" + .Value(RowNum, "IdExpOdYy") + "' and " + _
                                   "OdChr= '" + .Value(RowNum, "IdExpOdChr") + "' and " + _
                                   "OdNo= " + CStr(.Value(RowNum, "IdExpOdNo")) + " and " + _
                                   "OdSr= " + CStr(.Value(RowNum, "IdExpOdSr")) + IIF(ADC("wIdPrtKey") = ctCurrPrtn, " and OdPrtKey='" + ctCurrPrtn + "' ", ""))
      End If
    Case Is = UCase("IdSfx")       '*** recalc on IdExpOdTc, IdExpOdYy, IdExpOdChr, IdExpOdNo, IdExpOdSr
      If .Mode = fgmnorm Then
        'Or .Mode = fgmCopy
        '****** Sachin 3.02 28-12-07 - Current Partition
        .Value(RowNum, "IdSfx") = moCn.GetFldVal("Select OdSfx from OrdDsg where " + _
                                  "OdCoCd= '" + gs_CoCd + "' and " + _
                                  "OdTc= '" + .Value(RowNum, "IdExpOdTc") + "' and " + _
                                  "OdYy= '" + .Value(RowNum, "IdExpOdYy") + "' and " + _
                                  "OdChr= '" + .Value(RowNum, "IdExpOdChr") + "' and " + _
                                  "OdNo= " + CStr(.Value(RowNum, "IdExpOdNo")) + " and " + _
                                  "OdSr= " + CStr(.Value(RowNum, "IdExpOdSr")) + IIF(ADC("wIdPrtKey") = ctCurrPrtn, " and OdPrtKey='" + ctCurrPrtn + "' ", ""))
      End If
    Case Is = UCase("IdDmSz")       '*** recalc on IdExpOdTc, IdExpOdYy, IdExpOdChr, IdExpOdNo, IdExpOdSr
      If .Mode = fgmnorm Then
        'Or .Mode = fgmCopy
        '****** Sachin 3.02 28-12-07 - Current Partition
        .Value(RowNum, "IdDmSz") = moCn.GetFldVal("Select OdDmSz from OrdDsg where " + _
                                   "OdCoCd= '" + gs_CoCd + "' and " + _
                                   "OdTc= '" + .Value(RowNum, "IdExpOdTc") + "' and " + _
                                   "OdYy= '" + .Value(RowNum, "IdExpOdYy") + "' and " + _
                                   "OdChr= '" + .Value(RowNum, "IdExpOdChr") + "' and " + _
                                   "OdNo= " + CStr(.Value(RowNum, "IdExpOdNo")) + " and " + _
                                   "OdSr= " + CStr(.Value(RowNum, "IdExpOdSr")) + IIF(ADC("wIdPrtKey") = ctCurrPrtn, " and OdPrtKey='" + ctCurrPrtn + "' ", ""))
      End If
    Case Is = UCase("IdGldLs")       '*** recalc on IdVaCtg
      If (.Mode = fgmnorm Or .Mode = fgmCopy) And UCase(ms_IdTcTyp) = UCase(ctInTcTyp) Then
        .Value(RowNum, "IdGldLs") = moCn.GetFldVal("Select PNum from Param where " + _
                                   "PTyp = 'VACTG' and PMCd= '" + .Value(RowNum, "IdVaCtg") + "'")
      End If
    Case Is = UCase("IdQty")         '*** recalc on IfQty
      Dim wIfQty As Single
      If .Mode = fgmnorm Then
        wSourceGrdMode = Me.Controls(SourceName).Mode
        If (wSourceGrdMode = fgmnorm Or wSourceGrdMode = fgmCopyOver Or wSourceGrdMode = fgmdel) Then
          For i = 1 To GrdIB.Rows - 1
            If Not GrdIB.IsDel(i) Then
              wIfQty = wIfQty + GrdIB.Value(i, "IfQty")
            End If
          Next i
          .Value(GrdID.Row, "IdQty") = wIfQty
          If UCase(SourceName) = UCase("GrdIB") Then
            Controls(SourceName).AddSave "GrdID", asCurrent
          End If
        End If
      End If
    Case Is = UCase("IdAVal")        '*** recalc on IrRmAVal, IrSetAVal, iLAVal
      Dim wARmVal As Double, wALabVal As Double
      If .Mode = fgmnorm Or .Mode = fgmCopy Then
        wSourceGrdMode = Me.Controls(SourceName).Mode
        If (UCase(SourceName) = UCase("GrdIR") Or UCase(SourceName) = UCase("GrdIL")) And (wSourceGrdMode = fgmCopy Or wSourceGrdMode = fgmCopyOver) Then
          Exit Sub
        End If
        If (wSourceGrdMode = fgmnorm Or wSourceGrdMode = fgmCopyOver Or wSourceGrdMode = fgmdel) Then
          Call GrdIdRecalcHelper("IdAVal", GrdID.Row, 0, SourceColName)
          If (UCase(SourceName) = UCase("GrdIR") Or UCase(SourceName) = UCase("GrdIL")) And (wSourceGrdMode <> fgmCopy Or wSourceGrdMode <> fgmCopyOver) Then
            Controls(SourceName).AddSave "GrdID", asCurrent
          End If
        End If
      End If
    Case Is = UCase("IdIVal")        '*** recalc on IrRmIVal, IrSetIVal, iLiVal, IdQty, IdFixPrc, IdExpOdTc, IdExpOdYy, IdExpOdChr, IdExpOdNo, IdExpOdSr
      Dim wIRmVal As Double, wILabVal As Double
      wGrdIdRow = IIF(UCase(SourceName) = UCase("GrdID"), RowNum, GrdID.Row)
      If .Mode = fgmnorm Or .Mode = fgmCopy Then
        wSourceGrdMode = Me.Controls(SourceName).Mode
        If (UCase(SourceName) = UCase("GrdIR") Or UCase(SourceName) = UCase("GrdIL")) And (wSourceGrdMode = fgmCopy Or wSourceGrdMode = fgmCopyOver) Then
          Exit Sub
        End If
        
        If (wSourceGrdMode = fgmnorm Or wSourceGrdMode = fgmCopy Or wSourceGrdMode = fgmCopyOver Or wSourceGrdMode = fgmdel) Then
          Call GrdIdRecalcHelper("IdIVal", GrdID.Row, 0, SourceColName)
          If UCase(SourceName) = UCase("GrdIR") Or UCase(SourceName) = UCase("GrdIL") Then
            Controls(SourceName).AddSave "GrdID", asCurrent
          End If
        End If
      End If
    Case Is = UCase("IdZVal")        '*** recalc on IrRmZVal, IrSetZVal, iLZVal, IdQty, IdFixPrc, IdExpOdTc, IdExpOdYy, IdExpOdChr, IdExpOdNo, IdExpOdSr, IdHZValEqInv, IdiVal
      Dim wZRmVal As Double, wZLabVal As Double
      wGrdIdRow = IIF(UCase(SourceName) = UCase("GrdID"), RowNum, GrdID.Row)
      If .Mode = fgmnorm Or .Mode = fgmCopy Then
        wSourceGrdMode = Me.Controls(SourceName).Mode
        If (UCase(SourceName) = UCase("GrdIR") Or UCase(SourceName) = UCase("GrdIL")) And (wSourceGrdMode = fgmCopy Or wSourceGrdMode = fgmCopyOver) Then
          Exit Sub
        End If
        If (wSourceGrdMode = fgmnorm Or wSourceGrdMode = fgmCopy Or wSourceGrdMode = fgmCopyOver Or wSourceGrdMode = fgmdel) Then
          Call GrdIdRecalcHelper("IdZVal", GrdID.Row, 0, SourceColName)
          If GrdID.Value(wGrdIdRow, "IdHZValEqInv") = "N" And UCase(SourceColName) <> UCase("IdiVal") Then
            If GrdID.Value(wGrdIdRow, "IdFixPrc") = "N" Then
              If UCase(SourceName) = UCase("GrdIR") Or UCase(SourceName) = UCase("GrdIL") Then
                Controls(SourceName).AddSave "GrdID", asCurrent
              End If
            End If
          End If
        End If
    End If
    
    Case Is = UCase("IdAGldAsWt")
      If (.Mode = fgmnorm) Or (.Mode = fgmCopy) Then
        wGrdIdRow = IIF(UCase(SourceName) = UCase("GrdID"), RowNum, GrdID.Row)
        If (GrdIR.Mode = fgmnorm Or GrdIR.Mode = fgmdel) Then
            Call GrdIdRecalcHelper("IdAGldAsWt", wGrdIdRow, IIF(UCase(SourceColName) = UCase("IrMainMet") And RowNum > 0, RowNum, 0), SourceColName)
        End If
        If SourceName = UCase("GrdIR") And GrdIR.Mode <> fgmCopy Then
          GrdIR.AddSave "GrdID", asCurrent
        End If
      End If
    Case Is = UCase("IdiGldAsWt")
      If (.Mode = fgmnorm) Or (.Mode = fgmCopy) Then
        wGrdIdRow = IIF(UCase(SourceName) = UCase("GrdID"), RowNum, GrdID.Row)
        If (GrdIR.Mode = fgmnorm Or GrdIR.Mode = fgmdel) Then
            Call GrdIdRecalcHelper("IdIGldAsWt", wGrdIdRow, IIF(UCase(SourceColName) = UCase("IrMainMet") And RowNum > 0, RowNum, 0), SourceColName)
        End If
        If SourceName = UCase("GrdIR") And GrdIR.Mode <> fgmCopy Then
          GrdIR.AddSave "GrdID", asCurrent
        End If
      End If
    Case Is = UCase("IdZGldAsWt")
      If (.Mode = fgmnorm) Or (.Mode = fgmCopy) Then
        wGrdIdRow = IIF(UCase(SourceName) = UCase("GrdID"), RowNum, GrdID.Row)
        If (GrdIR.Mode = fgmnorm Or GrdIR.Mode = fgmdel) Then
          Call GrdIdRecalcHelper("IdZGldAsWt", wGrdIdRow, IIF(UCase(SourceColName) = UCase("IrMainMet") And RowNum > 0, RowNum, 0), SourceColName)
        End If
        If SourceName = UCase("GrdIR") And GrdIR.Mode <> fgmCopy Then
          GrdIR.AddSave "GrdID", asCurrent
        End If
      End If
    
    Case Is = UCase("IdALabAsWt")
    If (.Mode = fgmnorm) Or (.Mode = fgmCopy) Then
        wGrdIdRow = IIF(UCase(SourceName) = UCase("GrdID"), RowNum, GrdID.Row)
        If (GrdIR.Mode = fgmnorm Or GrdIR.Mode = fgmdel) Then
          Call GrdIdRecalcHelper("IdALabAsWt", wGrdIdRow, 0, SourceColName)
        End If
        If SourceName = UCase("GrdIR") And GrdIR.Mode <> fgmCopy Then
          GrdIR.AddSave "GrdID", asCurrent
        End If
      End If
    
    Case Is = UCase("IdILabAsWt")
     If (.Mode = fgmnorm Or .Mode = fgmCopy) Then
        If (UCase(SourceName) = UCase("GrdIR") Or UCase(SourceName) = UCase("GrdIL")) And (wSourceGrdMode = fgmCopy Or wSourceGrdMode = fgmCopyOver) Then
          Exit Sub
        End If
        
        wGrdIdRow = IIF(UCase(SourceName) = UCase("GrdID"), RowNum, GrdID.Row)
        If ((GrdIR.Mode = fgmnorm Or GrdIR.Mode = fgmdel Or GrdIR.Mode = fgmCopyOver) _
                          And .Value(wGrdIdRow, "IdLabWtFrOrd") = "N") Then
                          
            Call GrdIdRecalcHelper("IdILabAsWt", wGrdIdRow, 0, SourceColName)
          If SourceName = UCase("GrdIR") And GrdIR.Mode <> fgmCopy Then
            GrdIR.AddSave "GrdID", asCurrent
          End If
        
        End If
      End If
    
    Case Is = UCase("IdZLabAsWt")
    If (.Mode = fgmnorm) Or (.Mode = fgmCopy) Then
        
        wGrdIdRow = IIF(UCase(SourceName) = UCase("GrdID"), RowNum, GrdID.Row)
        If ((GrdIR.Mode = fgmnorm Or GrdIR.Mode = fgmdel Or GrdIR.Mode = fgmCopyOver) And .Value(wGrdIdRow, "IdLabWtFrOrd") = "N") Then
            Call GrdIdRecalcHelper("IdZLabAsWt", wGrdIdRow, 0, SourceColName)
        If SourceName = UCase("GrdIR") And GrdIR.Mode <> fgmCopy Then
          GrdIR.AddSave "GrdID", asCurrent
        End If
        
        End If
      End If
    Case Is = UCase("IdIAdd")        '*** recalc on IdIVal
      If (.Mode = fgmnorm) Or (.Mode = fgmCopy) Then
        wf_IdFOB = (.Value(RowNum, "IdIVal") + .Value(RowNum, "IdIAdd")) * (1 + ADC("wIdFOBPer") / 100)
        If ms_RndOffFOB = "Y" And UCase(ms_IdTcTyp) = UCase(ctCTBTcTyp) And wf_IdFOB <> Round(wf_IdFOB) Then
          .Value(RowNum, "IdIAdd") = .Value(RowNum, "IdIAdd") + Round(wf_IdFOB) - wf_IdFOB
        End If
      End If
      
    Case Is = UCase("IdZAdd")        '*** recalc on IdZVal
      If (.Mode = fgmnorm) Or (.Mode = fgmCopy) Then
        wf_IdFOB = (.Value(RowNum, "IdZVal") + .Value(RowNum, "IdZAdd")) * (1 + ADC("wIdFOBPer") / 100)
        If ms_RndOffFOB = "Y" And UCase(ms_IdTcTyp) = UCase(ctCTBTcTyp) And wf_IdFOB <> Round(wf_IdFOB) Then
          .Value(RowNum, "IdZAdd") = .Value(RowNum, "IdZAdd") + Round(wf_IdFOB) - wf_IdFOB
        End If
      End If

    Case Is = UCase("IdIFob")        '*** recalc on IdIVal, IdIAdd
      If (.Mode = fgmnorm) Or (.Mode = fgmCopy) Then
        .Value(RowNum, "IdIFob") = 0 '*** (Jen 2.12) (04/02/60)
         wOdMulBy = 1
          If UCase(ms_InOrdMulbyYN) = "Y" Then _
            wOdMulBy = moCn.GetFldVal(" Select (Case When OdFixPrc='N' Then OdMulby Else 1 End) From OrdDsg Where " + _
                                      " OdCoCd='" + gs_CoCd + "' And OdTc='" + .Value(RowNum, "IdExpOdTc") + "' and OdYy='" + .Value(RowNum, "IdExpOdYy") + "' and OdChr='" + .Value(RowNum, "IdExpOdChr") + "'" + _
                                      " and OdNo=" + CStr(.Value(RowNum, "IdExpOdNo")) + " and OdSr=" + CStr(.Value(RowNum, "IdExpOdSr")))
            wf_IdFOB = (.Value(RowNum, "IdIVal") * wOdMulBy + .Value(RowNum, "IdIAdd")) * (1 + ADC("wIdFOBPer") / 100)
         If ms_RndOffFOB = "Y" Then
          .Value(RowNum, "IdIFob") = Round(wf_IdFOB)
         Else
          .Value(RowNum, "IdIFob") = wf_IdFOB
         End If
      End If
      
    Case Is = UCase("IdZFob")        '*** recalc on IdZVal, IdZAdd
      If (.Mode = fgmnorm) Or (.Mode = fgmCopy) Then
        .Value(RowNum, "IdZFob") = 0 '*** (Jen 2.12) (04/02/60)
         
         wOdMulBy = 1
          If UCase(ms_InOrdMulbyYN) = "Y" Then _
            wOdMulBy = moCn.GetFldVal(" Select (Case When OdFixPrc='N' Then OdMulby Else 1 End) From OrdDsg Where " + _
                                      " OdCoCd='" + gs_CoCd + "' And OdTc='" + .Value(RowNum, "IdExpOdTc") + "' and OdYy='" + .Value(RowNum, "IdExpOdYy") + "' and OdChr='" + .Value(RowNum, "IdExpOdChr") + "'" + _
                                      " and OdNo=" + CStr(.Value(RowNum, "IdExpOdNo")) + " and OdSr=" + CStr(.Value(RowNum, "IdExpOdSr")))
   
        wf_IdFOB = (.Value(RowNum, "IdZVal") * wOdMulBy + .Value(RowNum, "IdZAdd")) * (1 + ADC("wIdFOBPer") / 100)
        If ms_RndOffFOB = "Y" Then
          .Value(RowNum, "IdZFob") = Round(wf_IdFOB)
        Else
          .Value(RowNum, "IdZFob") = wf_IdFOB
        End If
      End If
    Case Is = UCase("IdValAddn")       '*** recalc on IdVaCtg
      If (.Mode = fgmnorm Or .Mode = fgmCopy) And UCase(ms_IdTcTyp) = UCase(ctInTcTyp) Then
        .Value(RowNum, "IdValAddn") = moCn.GetFldVal("Select PNum1 from Param where " + _
                                   "PTyp = 'VACTG' and PMCd= '" + .Value(RowNum, "IdVaCtg") + "'")
      End If
    Case Is = UCase("IdCustmMtchWt")
      If .Mode = fgmnorm Or .Mode = fgmCopy Then
        wSourceGrdMode = Me.Controls(SourceName).Mode
        If (wSourceGrdMode = fgmnorm Or wSourceGrdMode = fgmCopyOver Or wSourceGrdMode = fgmdel) Then
          Call GrdIdRecalcHelper("IdCustmMtchWt", .Row, 0, SourceColName)
          Controls(SourceName).AddSave "GrdID", asCurrent
        End If
      End If
    Case Is = UCase("wIdADiaWt")
      If (GrdIR.Mode = fgmnorm Or GrdIR.Mode = fgmdel Or GrdIR.Mode = fgmInitOver) Then
          Call GrdIdRecalcHelper("wIdADiaWt", .Row, 0, SourceColName)
      End If
    Case Is = UCase("wIdIDiaWt")
      If (GrdIR.Mode = fgmnorm Or GrdIR.Mode = fgmdel Or GrdIR.Mode = fgmInitOver) Then
          Call GrdIdRecalcHelper("wIdIDiaWt", .Row, 0, SourceColName)
      End If
      
    Case Is = UCase("wIdZDiaWt")
      If (GrdIR.Mode = fgmnorm Or GrdIR.Mode = fgmdel Or GrdIR.Mode = fgmInitOver) Then
         
          Call GrdIdRecalcHelper("wIdZDiaWt", .Row, 0, SourceColName)
      End If
    
    Case Is = UCase("wIdACSWt")

      If (GrdIR.Mode = fgmnorm Or GrdIR.Mode = fgmdel Or GrdIR.Mode = fgmInitOver) Then
        Call GrdIdRecalcHelper("wIdACsWt", .Row, 0, SourceColName)
      End If
      
    
    Case Is = UCase("wIdICSWt")

      If (GrdIR.Mode = fgmnorm Or GrdIR.Mode = fgmdel Or GrdIR.Mode = fgmInitOver) Then
        Call GrdIdRecalcHelper("wIdICsWt", .Row, 0, SourceColName)
      End If
      
    
    Case Is = UCase("wIdZCSWt")

      If (GrdIR.Mode = fgmnorm Or GrdIR.Mode = fgmdel Or GrdIR.Mode = fgmInitOver) Then
        Call GrdIdRecalcHelper("wIdZCsWt", .Row, 0, SourceColName)
      End If
    End Select
  End With
End Sub
Private Sub GrdID_ColChange(ByVal RowNum As Integer, ByVal ColName As String, ByVal NewValue As Variant, Cancel As Boolean, ErrMsg As String)
  '*** IdExpOdTc        PMCd('TC', PValue= 'SO' and the type of entry should be 'OM') or ''
  '*** IdExpOdYy        PMCd('YY') or ''
  '*** IdExpOdChr       PSCd('CHR', adc(IdExpOdTc)) or ''
  '*** IdExpOdNo        OdNo from OrdDsg(OdTc= adc(IdExpOdTc), OdYy= adc(IdExpOdYy)
       '                OdChr= adc(IdExpOdChr)) or 0
  '*** IdExpOdSr        OdSr from OrdDsg(OdTc= adc(IdExpOdTc), OdYy= adc(IdExpOdYy)
       '                OdChr= adc(IdExpOdChr),OdNo= adc(IdExpOdNo)) or 0
  '*** IdDmCd           Show the 3D Picture & Line Drawing
  '*** IdALabAsWt       For all the lab records of the DsgSr if ilQw = 'W' then
      '                     ilAVal = GetLabVal(IdExpCmCd, ilMCd, ilSCd, iLARt, ilQty, IdALabAsWt, ilQw, 'C')
  '*** IdiGldAsWt       For all the Rm records of the DsgSr if IrMainMet = 'Y' then
      '                     IrRmIVal = NewValue * IrRmIRt
      '                 For all the lab records of the DsgSr if ilQw = 'W' then
      '                     iliVal = GetLabVal(IdExpCmCd, ilMCd, ilSCd, iLiRt, ilQty, IdiGldAsWt, ilQw, 'C')
  '*** IdILabAsWt       For all the Rm records of the DsgSr if IrMainMet = 'Y' then
      '                     IrRmIVal = NewValue * IrRmIRt
      '                 For all the lab records of the DsgSr if ilQw = 'W' then
      '                     iliVal = GetLabVal(IdExpCmCd, ilMCd, ilSCd, iLiRt, ilQty, IdILabAsWt, ilQw, 'C')
  '*** IdZGldAsWt       For all the Rm records of the DsgSr if IrMainMet = 'Y' then
      '                     IrRmZVal = NewValue * IrRmZRt
      '                 For all the lab records of the DsgSr if ilQw = 'W' then
      '                     ilZVal = GetLabVal(IdExpCmCd, ilMCd, ilSCd, iLZRt, ilQty, IdZGldAsWt, ilQw, 'C')
  '*** IdZLabAsWt       For all the Rm records of the DsgSr if IrMainMet = 'Y' then
      '                     IrRmZVal = NewValue * IrRmZRt
      '                 For all the lab records of the DsgSr if ilQw = 'W' then
      '                     ilZVal = GetLabVal(IdExpCmCd, ilMCd, ilSCd, iLZRt, ilQty, IdZLabAsWt, ilQw, 'C')
  '*** IdQty            For all the lab records of the DsgSr if ilQw = 'W' then
      '                     ilAVal = GetLabVal(IdExpCmCd, ilMCd, ilSCd, iLARt, ilQty, IdALabAsWt, ilQw, 'C') * (If iLQw= 'Q' then IdQty else 1)
      '                     iliVal = GetLabVal(IdExpCmCd, ilMCd, ilSCd, iLiRt, ilQty, IdILabAsWt, ilQw, 'C') * (If iLQw= 'Q' then IdQty else 1)
      '                     ilZVal = GetLabVal(IdExpCmCd, ilMCd, ilSCd, iLZRt, ilQty, IdZLabAsWt, ilQw, 'C') * (If iLQw= 'Q' then IdQty else 1)
  
  Dim i As Integer, wGldLsPrc As Single, wCmCd As String
  'Dim ws_AQw As String    '*** (19/03/07)
  Dim wd_LabAsWt As Double    '*** (Jen 2.14 Next 19/03/07)
  Dim ws_OmLkUpLabRt As String  ' **** Manali 3.03 - 05/07/08 - OmLkUp
  Dim wiLQw As String
  
  Dim w3DPth As String, wHskPth As String     '*** (Jen 3.2.0)
 
  If UCase(ms_IdTcTyp) = UCase(ctInTcTyp) And RowNum > 0 Then
    wCmCd = IIF(UCase(ColName) = UCase("IdExpCmCd"), NewValue, GrdID.Value(RowNum, "IdExpCmCd"))
    ' ****** Manali 3.03 - 05/07/08 - OmLkUp
    ws_OmLkUpLabRt = ""
   If GrdID.Value(RowNum, "IdExpOdNo") <> 0 Then
    ws_OmLkUpLabRt = moCn.GetFldVal("Select OmLkUpLabRt from OrdMst where " + _
                                     "OmCoCd= '" + gs_CoCd + "' and OmTc= '" + GrdID.Value(RowNum, "IdExpOdTc") + "' and " + _
                                     "OmYy= '" + GrdID.Value(RowNum, "IdExpOdYy") + "' " + _
                                     "and OmChr= '" + GrdID.Value(RowNum, "IdExpOdChr") + "' and " + _
                                     "OmNo= " + CStr(GrdID.Value(RowNum, "IdExpOdNo")) + " and OmPrtKey='" + ctCurrPrtn + "' ")
   End If
  ElseIf UCase(ms_IdTcTyp) = UCase(ctCTBTcTyp) And RowNum > 0 Then
    wCmCd = ADC("wIdCmCd")
  End If
  If ms_RtFrInvCust = "Y" Then
        wCmCd = ADC("wIdCmCd") '6.1-55
        ws_OmLkUpLabRt = ""
  End If
  
  With GrdID
    Select Case ColName
    Case Is = UCase("IdExpOdTc")
      '*** (Jen 3.01)
      If .Mode = fgmnorm Then
      '*** (Jen 3.01)
        Cancel = NewValue <> "" And ( _
                        moCn.GetFldVal("Select PValue From Param Where PTyp= 'TC' " + _
                        "And PMCd = '" + NewValue + "'") <> "SO" _
                        Or _
                        moCn.GetFldVal("select PValue1 from Param where PTyp= 'TcTyp' and " + _
                        "PMCd = (select PValue from Param where PTyp='TC' and " + _
                        "PMCd='" + NewValue + "')") <> "OM")
        ErrMsg = "Invalid Order Tc": Exit Sub
      End If
    Case Is = UCase("IdExpOdYy")
      '*** (Jen 3.01)
      If .Mode = fgmnorm Then
      '*** (Jen 3.01)
        Cancel = (NewValue <> "") And (Not moCn.RecSeek("Select PMCd From Param Where " + _
                                      "PTyp= 'Yy' And PMCd = '" + NewValue + "'"))
        ErrMsg = "Invalid Order Yy": Exit Sub
      End If
    Case Is = UCase("IdExpOdChr")
      '*** (Jen 3.01)
      If .Mode = fgmnorm Then
      '*** (Jen 3.01)
        Cancel = (NewValue <> "") And (Not moCn.RecSeek("Select vPSCd From vParam Where " + _
                                      "vPCoCd= '" + gs_CoCd + "' and vPTyp= 'CHR' And " + _
                                      "vPMCd = '" + .Value(RowNum, "IdExpOdTc") + "' " + _
                                      "and vPSCd='" + NewValue + "'"))
        ErrMsg = "Invalid Order Chr": Exit Sub
      End If
    Case Is = UCase("IdExpOdNo")
      '*** (Jen 3.01)
      If .Mode = fgmnorm Then
      '*** (Jen 3.01)
        '****** Sachin 3.02 28-12-07 - Current Partition
        Cancel = (NewValue <> 0) And (Not moCn.RecSeek("Select OmNo From OrdMst Where " + _
                                      "OmCoCd= '" + gs_CoCd + "' and " + _
                                      "OmTc= '" + .Value(RowNum, "IdExpOdTc") + "' And " + _
                                      "OmYy= '" + .Value(RowNum, "IdExpOdYy") + "' and " + _
                                      "OmChr= '" + .Value(RowNum, "IdExpOdChr") + "' and " + _
                                      "OmNo= " + CStr(NewValue) + " and OmPrtKey='" + ctCurrPrtn + "' "))
        ErrMsg = "Invalid Order No": Exit Sub
      End If
    Case Is = UCase("IdExpOdSr")
      '*** (Jen 3.01)
      If .Mode = fgmnorm Then
      '*** (Jen 3.01)
        '****** Sachin 3.02 28-12-07 - Current Partition
        Cancel = (NewValue <> 0) And (Not moCn.RecSeek("Select OdSr From OrdDsg Where " + _
                                      "OdCoCd= '" + gs_CoCd + "' and " + _
                                      "OdTc= '" + .Value(RowNum, "IdExpOdTc") + "' And " + _
                                      "OdYy= '" + .Value(RowNum, "IdExpOdYy") + "' and " + _
                                      "OdChr= '" + .Value(RowNum, "IdExpOdChr") + "' and " + _
                                      "OdNo= " + CStr(.Value(RowNum, "IdExpOdNo")) + " and " + _
                                      "OdSr= " + CStr(NewValue) + " and OdPrtKey='" + ctCurrPrtn + "' "))
        ErrMsg = "Invalid Order Sr": Exit Sub
      End If
    Case Is = UCase("IdDmCd")
      'cannot do in fldchange , do in fldvalid (but fldvalid is not triggered)
      If mf_ShowDsg = True Then
        'Dim w3DPth As String, wHskPth As String      '*** (Bef 3.2.0)
        '*** Jay 3.2.0 [OdPicNm]
        w3DPth = GetPictPath("DM", False, NewValue, .Value(RowNum, "IdDmSz"), gs_CoCd, .Value(RowNum, "IdExpOdTc"), .Value(RowNum, "IdExpOdYy"), .Value(RowNum, "IdExpOdChr"), .Value(RowNum, "IdExpOdNo"), .Value(RowNum, "IdExpOdSr"))
        wHskPth = GetPictPath("DM", True, NewValue, .Value(RowNum, "IdDmSz"), gs_CoCd, .Value(RowNum, "IdExpOdTc"), .Value(RowNum, "IdExpOdYy"), .Value(RowNum, "IdExpOdChr"), .Value(RowNum, "IdExpOdNo"), .Value(RowNum, "IdExpOdSr"))
        'w3DPth = GetPictPath("DM", False, NewValue)
        'wHskPth = GetPictPath("DM", True, NewValue)
        ' **** Zubin 211 **** '
        'w3DPth = GetPictPath("DM", NewValue, , False)
        'wHskPth = GetPictPath("DM", NewValue, , True)
        ' **** Zubin 211 **** '
        '*** Jay 3.2.0 [OdPicNm]
        Call ShowPic(wPic3DHt, wPic3DWdt, w3DPth, Pic3D)
        Call ShowPic(wPicHskHt, wPicHskWdt, wHskPth, PicLD)
      End If
        
    '*** (Jen 3.2.0)
    Case Is = UCase("IdDmSz")
      If mf_ShowDsg = True Then
        w3DPth = GetPictPath("DM", False, .Value(RowNum, "IdDmCd"), NewValue, gs_CoCd, .Value(RowNum, "IdExpOdTc"), .Value(RowNum, "IdExpOdYy"), .Value(RowNum, "IdExpOdChr"), .Value(RowNum, "IdExpOdNo"), .Value(RowNum, "IdExpOdSr"))
        wHskPth = GetPictPath("DM", True, .Value(RowNum, "IdDmCd"), NewValue, gs_CoCd, .Value(RowNum, "IdExpOdTc"), .Value(RowNum, "IdExpOdYy"), .Value(RowNum, "IdExpOdChr"), .Value(RowNum, "IdExpOdNo"), .Value(RowNum, "IdExpOdSr"))
        Call ShowPic(wPic3DHt, wPic3DWdt, w3DPth, Pic3D)
        Call ShowPic(wPicHskHt, wPicHskWdt, wHskPth, PicLD)
      End If
    '*** (Jen 3.2.0)
      
'''    Case Is = UCase("IdAGldAsWt")
'''      '*** Calculating the Lab Val
'''      '*** Irrespective of INV or CTB the Customer considered will be the ExpSOCust
'''      For i = 1 To GrdIL.Rows - 1
'''        If Not GrdIL.IsDel(i) And GrdIL.Value(i, "ilQw") = "W" Then
'''           GrdIL.Value(i, "ilAVal") = GetLabVal(GrdID.Value(RowNum, "IdExpCmCd"), _
'''                     GrdIL.Value(i, "ilMCd"), GrdIL.Value(i, "ilSCd"), _
'''                     GrdIL.Value(i, "iLARt"), GrdIL.Value(i, "ilQty"), _
'''                     MWLib.Div(NewValue, GrdID.Value(RowNum, "IdQty")), GrdIL.Value(i, "ilQw"), "C", "C") * _
'''                     GrdID.Value(RowNum, "IdQty")
'''                     'IIf(Trim(GrdIL.Value(i, "iLQw")) = "Q", GrdID.Value(RowNum, "IdQty"), NewValue)
'''           If Not GrdIL.IsNew(i) Then
'''             GrdIL.SaveRec (i)
'''           End If
'''        End If
'''      Next i

    '****** Sachin 2.13.0 - 26-05-2006 ******
    Case Is = UCase("IdALabAsWt")
      '*** Calculating the Lab Val
      '*** Irrespective of INV or CTB the Customer considered will be the ExpSOCust
      
      wd_LabAsWt = MWLib.Div(NewValue, GrdID.Value(RowNum, "IdQty"))    '*** (Jen 2.14 Next 19/03/07)
      
      For i = 1 To GrdIL.Rows - 1
        '*** (Bef 2.14 Next)
        'If Not GrdIL.IsDel(i) And GrdIL.Value(i, "ilQw") = "W" Then
        '*** (Bef 2.14 Next)
           
        '*** (Jen 2.14 Next 20/03/07)
        ' ******** Manali - 18/05/09 - "DHC" and "CHC" excluded
        
        If Not (GrdIL.Value(i, "iLMCd") = ctDHCLMCd Or GrdIL.Value(i, "iLMCd") = ctGHCLMCd Or GrdIL.Value(i, "iLMCd") = ctCHCLMCd) And Not GrdIL.IsDel(i) And GrdIL.Value(i, "iLCstQw") = "W" Then
          '6.1-55 GrdID.Value(RowNum, "IdExpCmCd") changed to wCmCd
          GrdIL.Value(i, "ilAVal") = GetLabVal(wCmCd, _
                     GrdIL.Value(i, "ilMCd"), GrdIL.Value(i, "ilSCd"), _
                     GrdIL.Value(i, "iLARt"), GrdIL.Value(i, "ilQty"), _
                     wd_LabAsWt, GrdIL.Value(i, "ilCstQw"), "C", "C", ADC("wIdCmCurCd"), _
                     IIF(ws_OmLkUpLabRt <> "", ws_OmLkUpLabRt, "")) * _
                     GrdID.Value(RowNum, "IdQty")
          If Not GrdIL.IsNew(i) Then
            GrdIL.SaveRec (i)
          End If
        ElseIf GrdIL.Value(i, "iLMCd") = ctCDWLMCd And Not GrdIL.IsDel(i) And GrdIL.Value(i, "iLCstQw") = "W" Then
          '6.1-55 GrdID.Value(RowNum, "IdExpCmCd") changed to wCmCd
          GrdIL.Value(i, "ilAVal") = GetLabVal(wCmCd, _
                     GrdIL.Value(i, "ilMCd"), GrdIL.Value(i, "ilSCd"), _
                     GrdIL.Value(i, "iLARt"), GrdIL.Value(i, "ilQty"), _
                             MWLib.Div(GetDCRmVal("IrRmAWt", RowNum, ctCDWLMCd, GetLaborSubCdLabLogic(GrdIL.Value(i, "ilMCd"), GrdIL.Value(i, "ilSCd"))), GrdID.Value(GrdID.Row, "IdQty")), _
                             GrdIL.Value(i, "ilCstQw"), "C", "C", ADC("wIdCmCurCd"), _
                     IIF(ws_OmLkUpLabRt <> "", ws_OmLkUpLabRt, "")) * _
                     GrdID.Value(RowNum, "IdQty")
          If Not GrdIL.IsNew(i) Then
            GrdIL.SaveRec (i)
          End If

        End If
      Next i
    '****** Sachin 2.13.0 - 26-05-2006 ******
    
    Case Is = UCase("IdIGldAsWt")
      ' exceptional recalc - calling recalc of IrRmIVal,cannot call saverec of invdsg as value
      ' of column IdiGldAsWt isnt into the grid yet
      
      '*** Calculating Rm Val
      For i = 1 To GrdIR.Rows - 1
        If Not GrdIR.IsDel(i) And GrdIR.Value(i, "IrMainMet") = "Y" Then
          '****** Sachin 2.13.0 - 31-05-2006 ******
          '****** Sachin 3.02 28-12-07 - Current Partition
          GrdIR.Value(i, "IrRmIVal") = IIF(moCn.GetFldVal("Select RmQw from RmMst where " + _
                                       "RmCd= '" + GrdIR.Value(i, "IrRmCd") + "' and RmPrtKey='" + ctCurrPrtn + "' ") = "Q", _
                                       GrdIR.Value(i, "IrRmQty"), NewValue) * GrdIR.Value(i, "IrRmIRt")
          '****** Sachin 2.13.0 - 31-05-2006 ******
           
''          If UCase(ms_IdTcTyp) = UCase(ctCTBTcTyp) Then
''            GrdIR.Value(i, "IrSetIRt") = GetLabRt(wCmCd, "SET", GrdIR.Value(i, "IrSetSCd"), "S", NewValue, ms_CmCtg)
''          End If
          '###2.06 If Not GrdIR.IsNew(i) Then
             GrdIR.SaveRec (i)
          '###2.06 End If
          Exit For
        End If
      Next i
            
      For i = 1 To GrdIL.Rows - 1
      ' ******** Manali - 18/05/09 - "DHC" and "CHC" excluded
        If (GrdIL.Value(i, "iLMCd") = ctGHCLMCd And Not GrdIL.IsDel(i) And GrdIL.Value(i, "ilQw") = "W") Then
          If UCase(ms_IdTcTyp) = UCase(ctInTcTyp) Then
            ' **** Manali 3.03 - 05/07/08 - Om Lk Up fields passed if not blank
          Dim wLaborSubLogic As String
          wLaborSubLogic = GetLaborSubCdLabLogic(GrdIL.Value(i, "ilMCd"), GrdIL.Value(i, "ilSCd"))
          
          GrdIL.Value(i, "iLiVal") = GrdIL.Value(i, "iLiRt") / 100 * GetDCRmVal(IIF(wLaborSubLogic = "P", "IrRmiWt", "IrRmiVal"), RowNum, ctGHCLMCd, wLaborSubLogic)
          GrdIL.SaveRec (i)
            
          End If
                     
           '###2.06 If Not GrdIL.IsNew(i) Then
              GrdIL.SaveRec (i)
           '###2.06 End If
        End If
      Next i
    Case Is = UCase("IdILabAsWt") '****** Sachin 2.13.0 - 26-05-2006 ******
      '*** Calculating the Lab Val
      For i = 1 To GrdIL.Rows - 1
      ' ******** Manali - 18/05/09 - "DHC" and "CHC" excluded
        If Not (GrdIL.Value(i, "iLMCd") = ctDHCLMCd Or GrdIL.Value(i, "iLMCd") = ctGHCLMCd Or GrdIL.Value(i, "iLMCd") = ctCHCLMCd Or GrdIL.Value(i, "iLMCd") = ctCDWLMCd Or GrdIL.Value(i, "iLMCd") = ctCBRKLMCd) And Not GrdIL.IsDel(i) Then
          wiLQw = ""
          If UCase(ms_IdTcTyp) = UCase(ctCTBTcTyp) Then
          
            '*** (Bef 2.13)
            'GrdIL.Value(I, "iLiRt") = GetLabRt(wCmCd, GrdIL.Value(I, "iLMCd"), GrdIL.Value(I, "iLSCd"), _
            '                          "S", MWLib.Div(NewValue, GrdID.Value(RowNum, "IdQty")), ms_CmCtg, wiLQw)
            '*** (Bef 2.13)
          
            '*** (Jen 2.13)
            GrdIL.Value(i, "iLiRt") = GetLabRt(wCmCd, GrdIL.Value(i, "iLMCd"), GrdIL.Value(i, "iLSCd"), _
                                      "S", MWLib.Div(NewValue, GrdID.Value(RowNum, "IdQty")), ms_CmCtg, wiLQw, , , ADC("wIdCmCurCd"))
            '*** (Jen 2.13)
            If wiLQw <> "" Then GrdIL.Value(i, "iLQw") = wiLQw
          End If
          
          GrdIL.Value(i, "iliVal") = GetLabVal(wCmCd, _
                      GrdIL.Value(i, "ilMCd"), GrdIL.Value(i, "ilSCd"), _
                      GrdIL.Value(i, "iLiRt"), GrdIL.Value(i, "ilQty"), _
                      MWLib.Div(NewValue, GrdID.Value(RowNum, "IdQty")), GrdIL.Value(i, "ilQw"), "S", ms_CmCtg, ADC("wIdCmCurCd"), _
                      IIF(ws_OmLkUpLabRt <> "", ws_OmLkUpLabRt, "")) * _
                      GrdID.Value(RowNum, "IdQty")
                      'IIf(Trim(GrdIL.Value(i, "iLQw")) = "Q", GrdID.Value(RowNum, "IdQty"), NewValue)
         ' End If
          '###2.06 If Not GrdIL.IsNew(i) Then
            GrdIL.SaveRec (i)
          '###2.06 End If
        End If
      Next i
    Case Is = UCase("IdZGldAsWt")
      ' exceptional recalc - calling recalc of IrRmZVal,cannot call saverec of invdsg as value
      ' of column IdZGldAsWt isnt into the grid yet
      
      '*** Calculating Rm Val
      For i = 1 To GrdIR.Rows - 1
      
        '*** (Bef 2.13) Remove 'Local Inv' Check
        'If Not GrdIR.IsDel(i) And GrdIR.Value(i, "IrMainMet") = "Y" And mLocalExp = "N" And ms_IdTcTyp = ctInTcTyp Then
        '*** (Bef 2.13) Remove 'Local Inv' Check
      
        '*** (Jen 2.13) Remove 'Local Inv' Check
        If Not GrdIR.IsDel(i) And GrdIR.Value(i, "IrMainMet") = "Y" And ms_IdTcTyp = ctInTcTyp Then
        '*** (Jen 2.13) Remove 'Local Inv' Check
        
          'New
          'wGldLsPrc = IIf(mLocalExp = "Y" And ms_IdTcTyp = ctInTcTyp And moCn.RecSeek("Select 'x' from RmMst where " + _
                      "RmCd= '" + GrdIR.Value(i, "IrRmCd") + "' and RmCtg in ('G', 'P')"), .Value(RowNum, "IdGldLs") / 100, 0)
          '****** Sachin 2.13.0 - 31-05-2006 ******
          '****** Sachin 3.02 28-12-07 - Current Partition
          GrdIR.Value(i, "IrRmZVal") = IIF(moCn.GetFldVal("Select RmQw from RmMst where " + _
                                       "RmCd= '" + GrdIR.Value(i, "IrRmCd") + "' and RmPrtKey='" + ctCurrPrtn + "' ") = "Q", _
                                       GrdIR.Value(i, "IrRmQty"), NewValue) * GrdIR.Value(i, "IrRmZRt") '* (1 + wGldLsPrc)
          '****** Sachin 2.13.0 - 31-05-2006 ******
          'New
          
          'GrdIR.Value(i, "IrRmZVal") = NewValue * GrdIR.Value(i, "IrRmZRt")
          '###2.06 If Not GrdIR.IsNew(i) Then
             GrdIR.SaveRec (i)
          '###2.06 End If
          Exit For
        End If
      Next i
      
      For i = 1 To GrdIL.Rows - 1
      ' ******** Manali - 18/05/09 - "DHC" and "CHC" excluded
        If (GrdIL.Value(i, "iLMCd") = ctGHCLMCd And Not GrdIL.IsDel(i) And GrdIL.Value(i, "ilQw") = "W") Then
          If UCase(ms_IdTcTyp) = UCase(ctInTcTyp) Then
            ' **** Manali 3.03 - 05/07/08 - Om Lk Up fields passed if not blank
'          Dim wLaborSubLogic As String
          wLaborSubLogic = GetLaborSubCdLabLogic(GrdIL.Value(i, "ilMCd"), GrdIL.Value(i, "ilSCd"))
          
          GrdIL.Value(i, "iLZVal") = GrdIL.Value(i, "iLZRt") / 100 * GetDCRmVal(IIF(wLaborSubLogic = "P", "IrRmZWt", "IrRmZVal"), RowNum, ctGHCLMCd, wLaborSubLogic)
          GrdIL.SaveRec (i)
            
          End If
                     
           '###2.06 If Not GrdIL.IsNew(i) Then
              GrdIL.SaveRec (i)
           '###2.06 End If
        End If
      Next i
      
    Case Is = UCase("IdZLabAsWt") '****** Sachin 2.13.0 - 26-05-2006 ******
      '*** Calculating the Lab Val
      For i = 1 To GrdIL.Rows - 1
      ' ******** Manali - 18/05/09 - "DHC" and "CHC" excluded
        If Not (GrdIL.Value(i, "iLMCd") = ctDHCLMCd Or GrdIL.Value(i, "iLMCd") = ctGHCLMCd Or GrdIL.Value(i, "iLMCd") = ctCBRKLMCd Or GrdIL.Value(i, "iLMCd") = ctCHCLMCd Or GrdIL.Value(i, "iLMCd") = ctCDWLMCd) And Not GrdIL.IsDel(i) And GrdIL.Value(i, "ilQw") = "W" Then
          If UCase(ms_IdTcTyp) = UCase(ctInTcTyp) Then
            ' **** Manali 3.03 - 05/07/08 - Om Lk Up fields passed if not blank
            GrdIL.Value(i, "ilZVal") = GetLabVal(wCmCd, _
                      GrdIL.Value(i, "ilMCd"), GrdIL.Value(i, "ilSCd"), _
                      GrdIL.Value(i, "iLZRt"), GrdIL.Value(i, "ilQty"), _
                      MWLib.Div(NewValue, GrdID.Value(RowNum, "IdQty")), GrdIL.Value(i, "ilQw"), "S", ms_CmCtg, ADC("wIdCmCurCd"), _
                      IIF(ws_OmLkUpLabRt <> "", ws_OmLkUpLabRt, "")) * _
                      GrdID.Value(RowNum, "IdQty")
                      'IIf(Trim(GrdIL.Value(i, "iLQw")) = "Q", GrdID.Value(RowNum, "IdQty"), NewValue)
          ElseIf UCase(ms_IdTcTyp) = UCase(ctCTBTcTyp) Then
            GrdIL.Value(i, "ilZVal") = GrdIL.Value(i, "iLZRt") * IIF(Trim(GrdIL.Value(i, "ilQw")) = "Q", GrdIL.Value(i, "ilQty") * GrdID.Value(RowNum, "IdQty"), NewValue)
          End If
                     
           '###2.06 If Not GrdIL.IsNew(i) Then
              GrdIL.SaveRec (i)
           '###2.06 End If
        End If
      Next i
    Case Is = UCase("IdQty")
      '*** Calculating the Lab Val
      ' Zubin 3.02.01
      Dim wOrLabWt As Double
      With GrdID
        If .Value(RowNum, "IdLabWtFrOrd") = "Y" Then
          wOrLabWt = GetLabAsWtFrOrd * NewValue
          .Value(RowNum, "IdILabAsWt") = wOrLabWt
          .Value(RowNum, "IdZLabAsWt") = wOrLabWt
        End If
      End With
      ' Zubin 3.02.01
      
      wd_LabAsWt = MWLib.Div(GrdID.Value(RowNum, "IdALabAsWt"), NewValue)   '*** (Jen 2.14 Next 19/03/07)
      
      For i = 1 To GrdIL.Rows - 1
        If Not GrdIL.IsDel(i) Then
           '*** Before 09/01/02 GrdIL.Value(i, "ilAVal") = GetLabVal(adc("wIdCmCd"), _
                     GrdIL.Value(i, "ilMCd"), GrdIL.Value(i, "ilSCd"), _
                     GrdIL.Value(i, "iLARt"), GrdIL.Value(i, "ilQty"), _
                     GrdID.Value(RowNum, "IdALabAsWt"), GrdIL.Value(i, "ilQw"), "C") * NewValue
          
          If UCase(ms_IdTcTyp) = UCase(ctCTBTcTyp) Then
            '*** (Bef 2.13)
            'GrdIL.Value(I, "iLiRt") = GetLabRt(wCmCd, GrdIL.Value(I, "iLMCd"), _
            '                          GrdIL.Value(I, "iLSCd"), "S", MWLib.Div(GrdID.Value(RowNum, "IdILabAsWt"), NewValue), ms_CmCtg, wiLQw)
            '*** (Bef 2.13)
          
            '*** (Jen 2.13)
            GrdIL.Value(i, "iLiRt") = GetLabRt(wCmCd, GrdIL.Value(i, "iLMCd"), _
                                      GrdIL.Value(i, "iLSCd"), "S", MWLib.Div(GrdID.Value(RowNum, "IdILabAsWt"), NewValue), ms_CmCtg, wiLQw, , , ADC("wIdCmCurCd"))
            '*** (Jen 2.13)
            If wiLQw <> "" Then GrdIL.Value(i, "iLQw") = wiLQw
          End If
          
          '*** (Jen 2.14 Next 20/03/07)
          ' **** Manali 3.03 - 05/07/08 - Om Lk Up fields passed if not blank '6.1-55 GrdID.Value(RowNum, "IdExpCmCd") changed to wCmCd
          If Not (GrdIL.Value(i, "iLMCd") = ctDHCLMCd Or GrdIL.Value(i, "iLMCd") = ctGHCLMCd Or GrdIL.Value(i, "iLMCd") = ctCHCLMCd Or GrdIL.Value(i, "iLMCd") = ctCBRKLMCd Or GrdIL.Value(i, "iLMCd") = ctCDWLMCd) And Not GrdIL.IsDel(i) Then
             GrdIL.Value(i, "ilAVal") = GetLabVal(wCmCd, _
                    GrdIL.Value(i, "ilMCd"), GrdIL.Value(i, "ilSCd"), _
                    GrdIL.Value(i, "iLARt"), GrdIL.Value(i, "ilQty"), _
                    wd_LabAsWt, GrdIL.Value(i, "ilCstQw"), "C", "C", ADC("wIdCmCurCd"), _
                    IIF(ws_OmLkUpLabRt <> "", ws_OmLkUpLabRt, "")) * _
                    NewValue
          '*** (Jen 2.14 Next 20/03/07)
          
          '*** (19/03/07)
          'ws_AQw = moCn.GetFldVal("Select LrQw from LabRt where LrMCd= '" + GrdIL.Value(i, "ilMCd") + "' " + _
          '                        "and LrSCd= '" + GrdIL.Value(i, "ilSCd") + "' and LrCmCurCd= '" + adc("wIdCmCurCd") + "' " + _
          '                        "and LrCmCtg= 'C' and LrCmCd= 'ZSELF' " + _
          '                        "and LrFrWt<= " + CStr(wd_LabAsWt) + "  " + _
          '                        "and LrToWt>= " + CStr(wd_LabAsWt) + " ")
          'ws_AQw = IIF(ws_AQw = "", ws_AQw, GrdIL.Value(i, "ilQw"))
          '
          'GrdIL.Value(i, "ilAVal") = GetLabVal(GrdID.Value(RowNum, "IdExpCmCd"), _
          '          GrdIL.Value(i, "ilMCd"), GrdIL.Value(i, "ilSCd"), _
          '          GrdIL.Value(i, "iLARt"), GrdIL.Value(i, "ilQty"), _
          '          wd_LabAsWt, ws_AQw, "C", "C", adc("wIdCmCurCd")) * _
          '          NewValue
          '*** (19/03/07)
          
          '*** (Bef 2.14 Next)
          'GrdIL.Value(i, "ilAVal") = GetLabVal(GrdID.Value(RowNum, "IdExpCmCd"), _
                    GrdIL.Value(i, "ilMCd"), GrdIL.Value(i, "ilSCd"), _
                    GrdIL.Value(i, "iLARt"), GrdIL.Value(i, "ilQty"), _
                    MWLib.Div(GrdID.Value(RowNum, "IdALabAsWt"), NewValue), GrdIL.Value(i, "ilQw"), "C", "C", ADC("wIdCmCurCd")) * _
                    NewValue
          '*** (Bef 2.14 Next)
                    'IIf(Trim(GrdIL.Value(i, "iLQw")) = "Q", NewValue, 1)
          
          ' **** Manali 3.03 - 05/07/08 - Om Lk Up fields passed if not blank
            GrdIL.Value(i, "iliVal") = GetLabVal(wCmCd, _
                    GrdIL.Value(i, "ilMCd"), GrdIL.Value(i, "ilSCd"), _
                    GrdIL.Value(i, "iLiRt"), GrdIL.Value(i, "ilQty"), _
                    MWLib.Div(GrdID.Value(RowNum, "IdILabAsWt"), NewValue), GrdIL.Value(i, "ilQw"), "S", ms_CmCtg, ADC("wIdCmCurCd"), _
                    IIF(ws_OmLkUpLabRt <> "", ws_OmLkUpLabRt, "")) * _
                    NewValue
                    'IIf(Trim(GrdIL.Value(i, "iLQw")) = "Q", NewValue, 1)
            If UCase(ms_IdTcTyp) = UCase(ctInTcTyp) Then
            ' **** Manali 3.03 - 05/07/08 - Om Lk Up fields passed if not blank
              GrdIL.Value(i, "ilZVal") = GetLabVal(wCmCd, _
                      GrdIL.Value(i, "ilMCd"), GrdIL.Value(i, "ilSCd"), _
                      GrdIL.Value(i, "iLZRt"), GrdIL.Value(i, "ilQty"), _
                      MWLib.Div(GrdID.Value(RowNum, "IdZLabAsWt"), NewValue), GrdIL.Value(i, "ilQw"), "S", ms_CmCtg, ADC("wIdCmCurCd"), _
                      IIF(ws_OmLkUpLabRt <> "", ws_OmLkUpLabRt, "")) * _
                      NewValue
                      'IIf(Trim(GrdIL.Value(i, "iLQw")) = "Q", NewValue, 1)
            ElseIf UCase(ms_IdTcTyp) = UCase(ctCTBTcTyp) Then
              GrdIL.Value(i, "ilZVal") = GrdIL.Value(i, "iLZRt") * IIF(Trim(GrdIL.Value(i, "iLQw")) = "Q", GrdIL.Value(i, "ilQty") * NewValue, GrdID.Value(RowNum, "IdZLabAsWt"))
            End If
          End If
          '###2.06 If Not GrdIL.IsNew(i) Then
            GrdIL.SaveRec (i)
          '###2.06 End If
        End If
      Next i
    
    '*** (Bef 2.13) Remove 'Local Inv' Check
    'Case Is = UCase("IdGldLs")
    '  '*** Calculating the Customs Val
    '  For i = 1 To GrdIR.Rows - 1
    '    If Not GrdIR.IsDel(i) Then
    '
    '      wGldLsPrc = IIF(mLocalExp = "Y" And ms_IdTcTyp = ctInTcTyp And moCn.RecSeek("Select 'x' from RmMst where " + _
    '                  "RmCd= '" + GrdIR.Value(i, "IrRmCd") + "' and RmCtg in ('G', 'P', 'S')"), .Value(RowNum, "IdGldLs") / 100, 0)
    '      If GrdIR.Value(i, "IrMainMet") = "Y" And mLocalExp = "N" And ms_IdTcTyp = ctInTcTyp Then
    '        GrdIR.Value(i, "IrRmZVal") = .Value(RowNum, "IdZGldAsWt") * (1 + wGldLsPrc) * GrdIR.Value(i, "IrRmZRt")
    '      ElseIf GrdIR.Value(i, "IrMainMet") = "N" Or mLocalExp = "Y" Or UCase(ms_IdTcTyp) = UCase(ctCTBTcTyp) Then
    '        GrdIR.Value(i, "IrRmZVal") = IIF(moCn.GetFldVal("Select RmQw from RmMst where " + _
    '                                     "RmCd= '" + GrdIR.Value(i, "IrRmCd") + "' ") = "Q", _
    '                                     GrdIR.Value(i, "IrRmQty"), (GrdIR.Value(i, "IrRmZWt") * _
    '                                     (1 + wGldLsPrc))) * GrdIR.Value(i, "IrRmZRt")
    '      End If
    '
    '      '###2.06 If Not GrdIR.IsNew(i) Then
    '      GrdIR.SaveRec (i)
    '      '###2.06 End If
    '    End If
    '  Next i
    '*** (Bef 2.13) Remove 'Local Inv' Check
    
    ' ****** Manali 3.5.0 - 18/05/09 - "DHC" and "CHC"
    Case Is = UCase("wIdADiaWt")
      '*** Calculating the Lab Val - DHC
      '*** Irrespective of INV or CTB the Customer considered will be the ExpSOCust
      Dim wd_DiaWt As Double
      If GrdIR.Mode = fgmCopyOver Or GrdIR.Mode = fgmnorm Or GrdIR.Mode = fgmdel Then
        wd_DiaWt = MWLib.Div(NewValue, GrdID.Value(RowNum, "IdQty"))
        
        For i = 1 To GrdIL.Rows - 1
            
          If (GrdIL.Value(i, "iLMCd") = ctDHCLMCd) And Not GrdIL.IsDel(i) Then
            '6.1-55 GrdID.Value(RowNum, "IdExpCmCd") changed to wCmCd
            GrdIL.Value(i, "ilAVal") = GetLabVal(wCmCd, _
                       GrdIL.Value(i, "ilMCd"), GrdIL.Value(i, "ilSCd"), _
                       GrdIL.Value(i, "iLARt"), GrdIL.Value(i, "ilQty"), _
                       wd_DiaWt, GrdIL.Value(i, "ilCstQw"), "C", "C", ADC("wIdCmCurCd"), _
                       IIF(ws_OmLkUpLabRt <> "", ws_OmLkUpLabRt, "")) * _
                       GrdID.Value(RowNum, "IdQty")
  
            
            If Not GrdIL.IsNew(i) Then
              GrdIL.SaveRec (i)
            End If
          
          ElseIf (GrdIL.Value(i, "ILMCD") = ctCDWLMCd) And Not GrdIL.IsDel(i) Then
          
            '6.1-55 GrdID.Value(RowNum, "IdExpCmCd") changed to wCmCd
            GrdIL.Value(i, "ilAVal") = GetLabVal(wCmCd, _
                       GrdIL.Value(i, "ilMCd"), _
                       GrdIL.Value(i, "ilSCd"), GrdIL.Value(i, "iLARt"), _
                                      GrdIL.Value(i, "ilQty"), MWLib.Div(GetDCRmVal("IrRmAWt", i, ctCDWLMCd, GetLaborSubCdLabLogic(GrdIL.Value(i, "ilMcd"), GrdIL.Value(i, "ilSCd"))), GrdID.Value(GrdID.Row, "idQty")), _
                                      GrdIL.Value(i, "ilCstQw"), "C", "C", ADC("wIdCmCurCd"), _
                       IIF(ws_OmLkUpLabRt <> "", ws_OmLkUpLabRt, "")) * _
                       GrdID.Value(RowNum, "IdQty")
  
            
            If Not GrdIL.IsNew(i) Then
              GrdIL.SaveRec (i)
            End If
          
          End If
        Next i
      End If
    ' ****** Manali 3.5.0 - 18/05/09 - "DHC" and "CHC"
    
    ' ****** Manali 3.5.0 - 18/05/09 - "DHC" and "CHC"
    Case Is = UCase("wIdIDiaWt")
      '*** Calculating the Lab Val - "DHC"
      If GrdIR.Mode = fgmCopyOver Or GrdIR.Mode = fgmnorm Or GrdIR.Mode = fgmdel Then
        For i = 1 To GrdIL.Rows - 1
          If (GrdIL.Value(i, "iLMCd") = ctDHCLMCd) And Not GrdIL.IsDel(i) Then
            wiLQw = ""
            If UCase(ms_IdTcTyp) = UCase(ctCTBTcTyp) Then
            
              GrdIL.Value(i, "iLiRt") = GetLabRt(wCmCd, GrdIL.Value(i, "iLMCd"), GrdIL.Value(i, "iLSCd"), _
                                        "S", MWLib.Div(GetDCRmVal("IrRmIWt", RowNum, ctCDWLMCd, GetLaborSubCdLabLogic(GrdIL.Value(i, "ilmcd"), GrdIL.Value(i, "ilscd"))), GrdID.Value(RowNum, "IdQty")), ms_CmCtg, wiLQw, , , ADC("wIdCmCurCd"))
              
              If wiLQw <> "" Then GrdIL.Value(i, "iLQw") = wiLQw
            End If
            Call GetLab_izVal(i)
            
          ElseIf (GrdIL.Value(i, "iLMCd") = ctCDWLMCd) And Not GrdIL.IsDel(i) Then
            wiLQw = ""
            If UCase(ms_IdTcTyp) = UCase(ctCTBTcTyp) Then
            
              GrdIL.Value(i, "iLiRt") = GetLabRt(wCmCd, GrdIL.Value(i, "iLMCd"), GrdIL.Value(i, "iLSCd"), _
                                        "S", MWLib.Div(GetDCRmVal("IrRmIWt", RowNum, ctCDWLMCd, GetLaborSubCdLabLogic(GrdIL.Value(i, "ilmcd"), GrdIL.Value(i, "ilscd"))), GrdID.Value(RowNum, "IdQty")), _
                                        ms_CmCtg, wiLQw, , , ADC("wIdCmCurCd"))
              
              If wiLQw <> "" Then GrdIL.Value(i, "iLQw") = wiLQw
              
            End If
             Call GetLab_izVal(i)
           ElseIf (GrdIL.Value(i, "iLMCd") = ctGHCLMCd) And Not GrdIL.IsDel(i) Then
            wiLQw = ""
            If UCase(ms_IdTcTyp) = UCase(ctCTBTcTyp) Then
            

              GrdIL.Value(i, "iLiRt") = GetLabRt(wCmCd, GrdIL.Value(i, "iLMCd"), GrdIL.Value(i, "iLSCd"), _
                                        "S", MWLib.Div(GetDCRmVal("IrRmIWt", RowNum, ctGHCLMCd, GetLaborSubCdLabLogic(GrdIL.Value(i, "ilmcd"), GrdIL.Value(i, "ilscd"))), GrdID.Value(RowNum, "IdQty")), _
                                        ms_CmCtg, wiLQw, , , ADC("wIdCmCurCd"))
              
              If wiLQw <> "" Then GrdIL.Value(i, "iLQw") = wiLQw
              
            End If
            ' ramram
             ' GrdIL.Value(i, "iliVal") = GetLabVal(wCmCd, _
                        GrdIL.Value(i, "ilMCd"), GrdIL.Value(i, "ilSCd"), _
                        GrdIL.Value(i, "iLiRt"), GrdIL.Value(i, "ilQty"), _
                        MWLib.Div(NewValue, GrdID.Value(RowNum, "IdQty")), GrdIL.Value(i, "ilQw"), "S", ms_CmCtg, ADC("wIdCmCurCd"), _
                        IIF(ws_OmLkUpLabRt <> "", ws_OmLkUpLabRt, "")) * _
                        GrdID.Value(RowNum, "IdQty")
             ' GrdIL.SaveRec (i)
             Call GetLab_izVal(i)
          End If
        Next i
      End If
    Case Is = UCase("wIdZDiaWt")
      If GrdIR.Mode = fgmCopyOver Or GrdIR.Mode = fgmnorm Or GrdIR.Mode = fgmdel Then
        For i = 1 To GrdIL.Rows - 1
          If (GrdIL.Value(i, "iLMCd") = ctDHCLMCd Or GrdIL.Value(i, "iLMCd") = ctGHCLMCd Or GrdIL.Value(i, "iLMCd") = ctCDWLMCd) And Not GrdIL.IsDel(i) Then
            If UCase(ms_IdTcTyp) = UCase(ctInTcTyp) Then
             Call GetLab_izVal(i)
            ElseIf UCase(ms_IdTcTyp) = UCase(ctCTBTcTyp) Then
              GrdIL.Value(i, "ilZVal") = GrdIL.Value(i, "iLZRt") * IIF(Trim(GrdIL.Value(i, "ilQw")) = "Q", GrdIL.Value(i, "ilQty") * GrdID.Value(RowNum, "IdQty"), NewValue)
            End If
                       
            GrdIL.SaveRec (i)
          End If
      Next i
    End If
    ' ****** Manali 3.5.0 - 18/05/09 - "DHC" and "CHC"
    
    ' ****** Manali 3.5.0 - 18/05/09 - "DHC" and "CHC"
    Case Is = UCase("wIdACSWt")
      '*** Calculating the Lab Val - "CHC"
      '*** Irrespective of INV or CTB the Customer considered will be the ExpSOCust
      Dim wd_CSWt As Double
      If GrdIR.Mode = fgmCopyOver Or GrdIR.Mode = fgmnorm Or GrdIR.Mode = fgmdel Then
        wd_CSWt = MWLib.Div(NewValue, GrdID.Value(RowNum, "IdQty"))
        
        For i = 1 To GrdIL.Rows - 1
            
          If GrdIL.Value(i, "iLMCd") = ctCHCLMCd And Not GrdIL.IsDel(i) Then
            '6.1-55 GrdID.Value(RowNum, "IdExpCmCd") changed to wCmCd
            GrdIL.Value(i, "ilAVal") = GetLabVal(wCmCd, _
                       GrdIL.Value(i, "ilMCd"), GrdIL.Value(i, "ilSCd"), _
                       GrdIL.Value(i, "iLARt"), GrdIL.Value(i, "ilQty"), _
                       wd_CSWt, GrdIL.Value(i, "ilCstQw"), "C", "C", ADC("wIdCmCurCd"), _
                       IIF(ws_OmLkUpLabRt <> "", ws_OmLkUpLabRt, "")) * _
                       GrdID.Value(RowNum, "IdQty")
            If Not GrdIL.IsNew(i) Then
              GrdIL.SaveRec (i)
            End If
          ElseIf (GrdIL.Value(i, "ILMCD") = ctCBRKLMCd) And Not GrdIL.IsDel(i) Then '6.1-55 GrdID.Value(RowNum, "IdExpCmCd") changed to wCmCd
            GrdIL.Value(i, "ilAVal") = GetLabVal(wCmCd, _
                       GrdIL.Value(i, "ilMCd"), _
                       GrdIL.Value(i, "ilSCd"), GrdIL.Value(i, "iLARt"), _
                                      GrdIL.Value(i, "ilQty"), wd_CSWt, _
                                      GrdIL.Value(i, "ilCstQw"), "C", "C", ADC("wIdCmCurCd"), _
                       IIF(ws_OmLkUpLabRt <> "", ws_OmLkUpLabRt, "")) * _
                       GrdID.Value(RowNum, "IdQty")
            If Not GrdIL.IsNew(i) Then
              GrdIL.SaveRec (i)
            End If
          End If
        Next i
      End If
    Case Is = UCase("wIdICSWt")
      '*** Calculating the Lab Val - "CHC"
      If GrdIR.Mode = fgmCopyOver Or GrdIR.Mode = fgmnorm Or GrdIR.Mode = fgmdel Then
        For i = 1 To GrdIL.Rows - 1
          If (GrdIL.Value(i, "iLMCd") = ctCHCLMCd Or GrdIL.Value(i, "ilMCd") = ctCBRKLMCd) And Not GrdIL.IsDel(i) Then
            wiLQw = ""
            If UCase(ms_IdTcTyp) = UCase(ctCTBTcTyp) Then
              GrdIL.Value(i, "iLiRt") = GetLabRt(wCmCd, GrdIL.Value(i, "iLMCd"), GrdIL.Value(i, "iLSCd"), _
                                        "S", MWLib.Div(NewValue, GrdID.Value(RowNum, "IdQty")), ms_CmCtg, wiLQw, , , ADC("wIdCmCurCd"))
              
              If wiLQw <> "" Then GrdIL.Value(i, "iLQw") = wiLQw
            End If
              ' ramram
            'GrdIL.Value(i, "iliVal") = GetLabVal(wCmCd, _
                        GrdIL.Value(i, "ilMCd"), GrdIL.Value(i, "ilSCd"), _
                        GrdIL.Value(i, "iLiRt"), GrdIL.Value(i, "ilQty"), _
                        MWLib.Div(NewValue, GrdID.Value(RowNum, "IdQty")), GrdIL.Value(i, "ilQw"), "S", ms_CmCtg, ADC("wIdCmCurCd"), _
                        IIF(ws_OmLkUpLabRt <> "", ws_OmLkUpLabRt, "")) * _
                        GrdID.Value(RowNum, "IdQty")
            Call GetLab_izVal(i)
            GrdIL.SaveRec (i)
           
          End If
        Next i
      End If
    ' ****** Manali 3.5.0 - 18/05/09 - "DHC" and "CHC"
    
    ' ****** Manali 3.5.0 - 18/05/09 - "DHC" and "CHC"
    Case Is = UCase("wIdZCSWt")
      '*** Calculating the Lab Val - "CHC"
      If GrdIR.Mode = fgmCopyOver Or GrdIR.Mode = fgmnorm Or GrdIR.Mode = fgmdel Then
        For i = 1 To GrdIL.Rows - 1
          If ((GrdIL.Value(i, "iLMCd") = ctCHCLMCd Or GrdIL.Value(i, "iLMCd") = ctCBRKLMCd) And Not GrdIL.IsDel(i)) Then
            If UCase(ms_IdTcTyp) = UCase(ctInTcTyp) Then
              
              'GrdIL.Value(i, "ilZVal") = GetLabVal(wCmCd, _
                        GrdIL.Value(i, "ilMCd"), GrdIL.Value(i, "ilSCd"), _
                        GrdIL.Value(i, "iLZRt"), GrdIL.Value(i, "ilQty"), _
                        MWLib.Div(NewValue, GrdID.Value(RowNum, "IdQty")), GrdIL.Value(i, "ilQw"), "S", ms_CmCtg, ADC("wIdCmCurCd"), _
                        IIF(ws_OmLkUpLabRt <> "", ws_OmLkUpLabRt, "")) * _
                        GrdID.Value(RowNum, "IdQty")
               Call GetLab_izVal(i)
            ElseIf UCase(ms_IdTcTyp) = UCase(ctCTBTcTyp) Then
              GrdIL.Value(i, "ilZVal") = GrdIL.Value(i, "iLZRt") * IIF(Trim(GrdIL.Value(i, "ilQw")) = "Q", GrdIL.Value(i, "ilQty") * GrdID.Value(RowNum, "IdQty"), NewValue)
            End If
                       
            GrdIL.SaveRec (i)
          End If
        Next i
     End If
      ' ****** Manali 3.5.0 - 18/05/09 - "DHC" and "CHC"
    
    End Select
  End With
End Sub
Private Sub GrdID_InitKey(ByVal RowNum As Integer)
  '*** Initialise Key Fields (i.e. IdTc, IdYy, IdChr, IdNo)
  With GrdID
    If mf_inspecialCopy Then
      .Store "Cpy", "Y"
    Else
      .Store "Cpy", ""
    End If

    .Store "IdCoCd", gs_CoCd
    .Store "IdTc", ADC("wIdTc")
    .Store "IdYy", ADC("wIdYy")
    .Store "IdChr", ADC("wIdChr")
    .Store "IdNo", ADC("wIdNo")
    
    '****** Sachin 3.02 - Id fields Initialization
    .Store "IdDmIdNo", moCn.GetFldVal("Select DmIdNo From DsgMst " + _
                                                       " Where DmTcTyp='DM' and DmCd='" + .Value(RowNum, "IdDmCd") + "' And DmSz='' " + IIF(ADC("wIdPrtKey") = ctCurrPrtn, " And DmPrtKey='" + ctCurrPrtn + "'", ""))

    .Store "IdOdIdNo", moCn.GetFldVal("Select OdIdNo From OrdDsg " + _
                                                       " Where OdCoCd='" + gs_CoCd + "' and OdTc='" + .Value(RowNum, "IdExpOdTc") + "' And OdYy='" + .Value(RowNum, "IdExpOdYy") + "' " + _
                                                       " And  OdChr ='" + .Value(RowNum, "IdExpOdChr") + "' And OdNo=" + CStr(.Value(RowNum, "IdExpOdNo")) + " And OdSr=" + CStr(.Value(RowNum, "IdExpOdSr")) + IIF(ADC("wIdPrtKey") = ctCurrPrtn, " and OdPrtKey='" + ctCurrPrtn + "' ", ""))
    
    If .IsNew(RowNum) Then
        .Store "IdInIdNo", moCn.GetFldVal("Select InIdNo From InvHd " + _
                                          " Where InCoCd='" + gs_CoCd + "' And InTc='" + ADC("wIdTc") + "' and InYy='" + ADC("wIdYy") + "' and InChr='" + ADC("wIdChr") + "' And InNo=" + CStr(ADC("wIdNo")) + " and InPrtKey='" + ctCurrPrtn + "' ")
                                                                                    
    '****** Sachin 3.02 - Id fields Initialization
    End If
  End With
  
End Sub
Private Sub GrdID_RowSave(ByVal RowNum As Integer, Cancel As Boolean, ErrMsg As String)
''  Dim wZAddVal As Single, wLabRow As Integer
''  If GrdIB.Rows - 1 = 0 Then Exit Sub
''  If GrdID.Value(RowNum, "IdIVal") <> GrdID.Value(RowNum, "IdZVal") Then
''    wZAddVal = (GrdID.Value(RowNum, "IdIVal") - GrdID.Value(RowNum, "IdZVal")) + _
''               moCn.GetFldVal("Select iLZVal from InvLab where iLTc= '" + ADC("wIdTc") + "' and " + _
''               "iLYy= '" + ADC("wIdYy") + "' and iLChr= '" + ADC("wIdChr") + "' and " + _
''               "iLNo= " + CStr(ADC("wIdNo")) + " and iLSr= " + CStr(GrdID.Value(RowNum, "IdSr")) + " and " + _
''               "iLMCd= '"+ctAddLMCd+"' and iLSCd= '-' ")
''
''    moCn.Execute ("Delete From InvLab where iLTc= '" + ADC("wIdTc") + "' and " + _
''                 "iLYy= '" + ADC("wIdYy") + "' and iLChr= '" + ADC("wIdChr") + "' and " + _
''                 "iLNo= " + CStr(ADC("wIdNo")) + " and iLSr= " + CStr(GrdID.Value(RowNum, "IdSr")) + " and " + _
''                 "iLMCd= '"+ctAddLMCd+"' and iLSCd= '-' ")
''
''    If Not moCn.RecSeek("Select 'x' from Param where PTyp= 'LABSCD' and PMCd= '"+ctAddLMCd+"' and " + _
''       "PSCd= '-'") Then
''      Cancel = True: ErrMsg = "The Lab SubCd '-' Should Be Compulsorily Present For The MainCd '"+ctAddLMCd+"'": Exit Sub
''    End If
''
''    With GrdIL
''      .StartCopy
''      .AddItem
''      wLabRow = .Rows - 1
''      .Value(wLabRow, "iLSrNo") = wLabRow
''      .Value(wLabRow, "iLMCd") = ctAddLMCd
''      .Value(wLabRow, "iLSCd") = "-"
''      .Value(wLabRow, "iLQty") = 1
''      .Value(wLabRow, "iLARt") = 0
''      .Value(wLabRow, "iLiRt") = 0
''      .Value(wLabRow, "iLZRt") = wZAddVal
''      .Value(wLabRow, "ilQw") = "Q"
''      .Value(wLabRow, "iLZVal") = wZAddVal
''      .SaveRec (wLabRow)
''      .EndCopy
''    End With
''  End If

End Sub
Private Sub GrdID_Valid()
  '*** (Bef 16/07/05) (Bef 2.11 color)
'  '*** Disable all the Command Buttons associated with the grid GrdID
'  '*** If there are no records in the GrdID row and the Invoice is not closed,
'      'then enable the Make Invoice option else disable it
'
'  CmdDD.Enabled = False
'  '%%% CmdDC.Enabled = False
'  CmdCRL.Enabled = False
'  If GrdID.Rows = 1 Then
'    CmdIL.Enabled = False
'    CmdIB.Enabled = False
'    CmdSumm.Enabled = False
'  End If
'
'  If GrdID.Rows = 1 And mInCls = "N" Then
'    CmdMI.Enabled = True
'  Else
'    CmdMI.Enabled = False
'  End If
'  CmdSH.Enabled = False
  '*** (Bef 16/07/05) (Bef 2.11 color)
  
End Sub
Private Sub GrdID_LostFocus()
  '*** (16/07/05) (2.11 Color)
  '*** Disable all the Command Buttons associated with the grid GrdID
  '*** If there are no records in the GrdID row and the Invoice is not closed,
      'then enable the Make Invoice option else disable it

  If ms_CurrentBtn = "" And Not (Me.ActiveControl Is Nothing) Then
    If TypeOf Me.ActiveControl Is MWCTL_BTN1 Then
      ms_CurrentBtn = UCase(Me.ActiveControl.Name)
    End If
  End If

  If ms_CurrentBtn <> UCase("CmdDD") Then CmdDD.Enabled = False
  '%%% CmdDC.Enabled = False
  If ms_CurrentBtn <> UCase("CmdCRL") Then CmdCRL.Enabled = False
  
  If ms_CurrentBtn <> UCase("CmdDRL") Then CmdDRL.Enabled = False '****** Sachin 2.13.0 - 27-05-2006 ******
  If ms_CurrentBtn <> UCase("CmdInvDel") Then CmdInvDel.Enabled = False '****** Sachin 4.1.2.0 ******
  
  If GrdID.Rows = 1 Then
    If ms_CurrentBtn <> UCase("CmdIL") Then CmdIL.Enabled = False
    If ms_CurrentBtn <> UCase("CmdIB") Then CmdIB.Enabled = False
    If ms_CurrentBtn <> UCase("CmdSumm") Then CmdSumm.Enabled = False
  End If
  
  If GrdID.Rows = 1 And mInCls = "N" Then
    CmdMI.Enabled = True
  Else
    If ms_CurrentBtn <> UCase("CmdMI") Then CmdMI.Enabled = False
  End If
  If ms_CurrentBtn <> UCase("CmdSH") Then CmdSH.Enabled = False
  ms_CurrentBtn = ""
  '*** (16/07/05) (2.11 Color)
End Sub
Private Sub GrdIR_RowWhen(ByVal RowNum As Integer)
  '*** If New Row IrSrNo= max(IrSrNo) of the grid GrdIR + 1
  
  '****** Sachin 3.02 - Default value 'C'
  If GrdIR.IsNew(RowNum) Then GrdIR.Value(RowNum, "IrPrtKey") = ctCurrPrtn
  
  '****** Sachin 2.13.0 - 27-05-2006 ******
  '****** Sachin 3.02 - Locking the Previous Partition Records
  If ADC("wIDDT") < CDate(ms_PurgeDt) Or moCn.RecSeek("Select 'x' from InvZRm where IzCoCd= '" + gs_CoCd + "' and IzTc= '" + ADC("wIdTc") + "' and " + _
                  "IzYy=  '" + ADC("wIdYy") + "' and IzChr=  '" + ADC("wIdChr") + "' and " + _
                  "IzNo=  " + CStr(ADC("wIdNo")) + " and IzPrtKey='" + ctCurrPrtn + "' ") Then
    GrdIR.AllowAdd = False
  Else
    If mf_DtLocked = False And mInCls = "N" Then GrdIR.AllowAdd = True
  End If
  '****** Sachin 2.13.0 - 27-05-2006 ******
    
  With GrdIR
    '****** Sachin 3.02 - Locking the Previous Partition Records
    If GrdIR.IsNew(RowNum) Then
      If .Value(RowNum, "IrSrNo") = 0 Then .Value(RowNum, "IrSrNo") = .MaxVal("IrSrNo") + 1
      GrdIR.AllowDelete = True  '****** Sachin 2.13.0 - 27-05-2006 ******
    Else
      GrdIR.AllowDelete = False  '****** Sachin 2.13.0 - 27-05-2006 ******
    End If
    
    ADC("WIRCUSTRMCD") = .Value(RowNum, "IrCustRmCd")   '4.1.0.0
  End With
End Sub
Private Sub GrdIR_ColWhen(ByVal RowNum As Integer, ByVal ColName As String, Cancel As Boolean, ErrMsg As String, ByVal Action As MwfCtl.en_fgAction)
  '*** Do not allow the user to edit any field if InCls = 'Y' in InvHd Table
  '*** Give help for IrRmCd, IrSetSCd, IrMainMet
  
  '****** Sachin 3.02 - Locking the Previous Partition Records
  If ADC("WIDPRTKEY") <> ctCurrPrtn Then Cancel = True: ErrMsg = "Cannot Edit Previous Partition Records": Exit Sub
  
  If mf_DtLocked = True Then Cancel = True: ErrMsg = "Cannot Edit Locked Entry": Exit Sub
  
  '*** (Jen 2.13)
  If moCn.RecSeek("Select 'x' from InvZRm where IzCoCd= '" + gs_CoCd + "' and IzTc= '" + ADC("wIdTc") + "' and " + _
                  "IzYy=  '" + ADC("wIdYy") + "' and IzChr=  '" + ADC("wIdChr") + "' and " + _
                  "IzNo=  " + CStr(ADC("wIdNo")) + " and IzPrtKey='" + ctCurrPrtn + "' ") Then
    '*** (Bef 3.1.0 this code is not required)
    'GrdIR.AllowAdd = False    '****** Sachin 2.13.0 - 27-05-2006 ******
    '*** (Bef 3.1.0 this code is not required)
    Cancel = True: ErrMsg = "Cannot Edit As Custom Matching Has Been Done": Exit Sub
  '*** (Bef 3.1.0 this code is not required)
  'Else
  '  GrdIR.AllowAdd = True     '****** Sachin 2.13.0 - 27-05-2006 ******
  '*** (Bef 3.1.0 this code is not required)
  End If
  '*** (Jen 2.13)
  
  If mInCls = "Y" Then Cancel = True: ErrMsg = "Cannot Edit As Invoice Is Closed": Exit Sub
  Select Case UCase(ColName)
  Case Is = UCase("IrRmCd")
    Call HlpList.RmCd("N")
  Case Is = UCase("IrSetSCd")
    Call HlpList.PSCd("LABSCD", "SET")
  Case Is = UCase("IrMainMet")
    Call HlpList.PMCd("YN")
' Zubin 212
  ' ***** Manali 3.03 - 16/07/08 - LotNo Open for 'X'
  Case Is = UCase("IrLotNo")
    Cancel = Not moCn.RecSeek("Select 'x' from RmMst where RmCd= '" + GrdIR.Value(RowNum, "IrRmCd") + "' and " + _
                                "RmCtg In " + ctLotNoRmCtg + IIF(gs_Partition = ctCurrPrtn, " And RmPrtKey='" + GrdIR.Value(RowNum, "IrPrtKey") + "' ", ""))
                                
    If Cancel = True Then ErrMsg = "Cannot Enter Lot No. For This RM Category": Exit Sub
    Call HlpList.LotNo(gs_CoCd, "FG", GrdIR.Value(RowNum, "IrRmCd"), GrdIR.Value(RowNum, "IrRmSz"))
' Zubin 212
  End Select
End Sub
Private Sub GrdIR_ReCalcRoutine(ByVal RowNum As Integer, ByVal ColName As String, ByVal OldVal As Variant, ByVal SourceName As String, ByVal SourceColName As String)
  '*** IrRmAVal on IrRmAWt,IrRmARt                                 Normal,Copy
      '   IrRmAVal = IrRmAWt * IrRmARt
  '*** IrRmIVal on IrRmIWt,IrRmIRt,IrMainMet                       Normal,Copy
      'If IrMainMet = 'Y'
      '   IrRmIVal = IdIGldAsWt * IrRmIRt
      'If IrMainMet = 'N'
      '   IrRmIVal = IrRmiWt * IrRmIRt
  '*** IrRmZVal on IrRmZWt,IrRmZRt,IrMainMet                       Normal,Copy
      'If IrMainMet = 'Y'
      '   IrRmZVal = IdZGldAsWt * IrRmZRt
      'If IrMainMet = 'N'
      '   IrRmZVal = IrRmZWt * IrRmZRt
  '*** IrSetAVal on IrSetSCd,IrSetARt,IrRmQty                      Normal,Copy
      '   IrSetAVal = GetLabVal(IdExpCmCd, 'SET', IrSetSCd, IrSetARt, IrRmQty, 0, 'Q', 'C')
  '*** IrSetIVal on IrSetSCd,IrSetIRt,IrRmQty                      Normal,Copy
      '   IrSetIVal = GetLabVal(IdExpCmCd, 'SET', IrSetSCd, IrSetIRt, IrRmQty, 0, 'Q', 'C')
  '*** IrSetZVal on IrSetSCd,IrSetZRt,IrRmQty                      Normal,Copy
      '   IrSetZVal = GetLabVal(IdExpCmCd, 'SET', IrSetSCd, IrSetZRt, IrRmQty, 0, 'Q', 'C')
  '*** IrRmZPureRt on IrSetSCd,IrSetZRt,IrRmQty                    Normal,Copy
      '   If RmPurityZ of IrRmCd > 0 then IrRmZPureRt = IrRmZRt/wRmPurityZ
      '   If RmPurityZ of IrRmCd <= 0 then IrRmZPureRt = IrRmZRt
  '*** IrRmZCd on IrRmCd                                           Normal,Copy
      '   IrRmZCd = RmZCd from RmMst(RmCd= IrRmCd)
  '*** IrMainMet on IrRmCd                                         Normal,Copy
      '   IrMainMet = 'N'
  
  Dim wCmCd As String
  
  '*** (Bef 2.13) Dim wRmPurityZ As Single

  Dim wRmPurityZ As Double, ws_RmCtg As String
  Dim wf_CustmCdPurityZ As Double   '*** (Jen 2.13)
  Dim ws_OmLkUpLabRt As String    ' **** Manali 3.03 - 05/07/08 - Om Lk Up
  
 ' If mf_StopRecalcDuringCopyIR = True Then Exit Sub
  
  If UCase(ms_IdTcTyp) = UCase(ctInTcTyp) And GrdID.Rows - 1 > 0 Then
    wCmCd = GrdID.Value(GrdID.Row, "IdExpCmCd")
    ' ***** Manali 3.03 - 05/07/08 - OmLkUp
    ws_OmLkUpLabRt = ""
   If GrdID.Value(GrdID.Row, "IdExpOdNo") <> 0 Then
    ws_OmLkUpLabRt = moCn.GetFldVal("Select OmLkUpLabRt from OrdMst where " + _
                                     "OmCoCd= '" + gs_CoCd + "' and OmTc= '" + GrdID.Value(GrdID.Row, "IdExpOdTc") + "' and " + _
                                     "OmYy= '" + GrdID.Value(GrdID.Row, "IdExpOdYy") + "' " + _
                                     "and OmChr= '" + GrdID.Value(GrdID.Row, "IdExpOdChr") + "' and " + _
                                     "OmNo= " + CStr(GrdID.Value(GrdID.Row, "IdExpOdNo")) + " and OmPrtKey='" + ctCurrPrtn + "' ")
   End If
  ElseIf UCase(ms_IdTcTyp) = UCase(ctCTBTcTyp) Then
    wCmCd = ADC("wIdCmCd")
  End If
  If ms_RtFrInvCust = "Y" Then
        wCmCd = ADC("wIdCmCd") '6.1-55
        ws_OmLkUpLabRt = ""
  End If
  
  With GrdIR
    Select Case ColName
    Case Is = UCase("IrRmAVal")   '*** recalc on IrRmAWt, IrRmARt
      If .Mode = fgmnorm Or .Mode = fgmCopy Then
        .Value(RowNum, "IrRmAVal") = IIF(moCn.GetFldVal("Select RmQwCst from RmMst where " + _
                                     "RmCd= '" + .Value(RowNum, "IrRmCd") + "'" + IIF(ADC("wIdPrtKey") = ctCurrPrtn, " and RmPrtKEy='" + ctCurrPrtn + "' ", "")) = "Q", _
                                     .Value(RowNum, "IrRmQty"), .Value(RowNum, "IrRmAWt")) _
                                     * .Value(RowNum, "IrRmARt")
      End If
    Case Is = UCase("IrRmIRt")
      If (.Mode = fgmnorm Or .Mode = fgmCopy) And (ms_IdTcTyp = ctCTBTcTyp) Then
        ws_RmCtg = UCase(Trim(moCn.GetFldVal("Select RmCtg from RmMst where RmCd= '" + .Value(RowNum, "IrRmCd") + "' " + IIF(ADC("wIdPrtKey") = ctCurrPrtn, " And RmPrtKey='" + ctCurrPrtn + "' ", ""))))

        If ADC("wIdLmgSal") > 0 And ws_RmCtg = "G" Then '6.1-55 ADC("wIdCmCd") changed to wCmCd
         .Value(RowNum, "IrRmIRt") = GetGldRt(wCmCd, ADC("wIdLmgSal"), _
                                      .Value(RowNum, "IrRmCd"), .Value(RowNum, "IrAlySalRt"), "DM", _
                                      GrdID.Value(GrdID.Row, "IdDmCd"), "S",  "C", "T", md_InCnvRt, , , , , .Value(RowNum, "IRRMSZ"), .Value(RowNum, "IRRMSZ2"), 0)
        ElseIf ADC("wIdLmpSal") > 0 And ws_RmCtg = "P" Then '6.1-55 ADC("wIdCmCd") changed to wCmCd
          .Value(RowNum, "IrRmIRt") = GetGldRt(wCmCd, ADC("wIdLmpSal"), _
                                      .Value(RowNum, "IrRmCd"), .Value(RowNum, "IrAlySalRt"), "DM", _
                                      GrdID.Value(GrdID.Row, "IdDmCd"), "S",  "T", md_InCnvRt, , , , , .Value(RowNum, "IRRMSZ"), .Value(RowNum, "IRRMSZ2"), 0)
        ElseIf ADC("wIdLmsSal") > 0 And ws_RmCtg = "S" Then '6.1-55 ADC("wIdCmCd") changed to wCmCd
          If mSlvMod = True Then
            .Value(RowNum, "IrRmIRt") = GetGldRt(wCmCd, ADC("wIdLmsSal"), _
                                        .Value(RowNum, "IrRmCd"), .Value(RowNum, "IrAlySalRt"), "DM", _
                                        GrdID.Value(GrdID.Row, "IdDmCd"), "S", "T", md_InCnvRt, , , , , .Value(RowNum, "IRRMSZ"), .Value(RowNum, "IRRMSZ2"), 0)
          Else
            .Value(RowNum, "IrRmIRt") = 0
          End If
        ElseIf ADC("wIdLmlSal") > 0 And ws_RmCtg = "L" Then '6.1-55 ADC("wIdCmCd") changed to wCmCd
          If mb_PdMod = True Then
            .Value(RowNum, "IrRmIRt") = GetGldRt(wCmCd, ADC("wIdLmlSal"), _
                                        .Value(RowNum, "IrRmCd"), .Value(RowNum, "IrAlySalRt"), "DM", _
                                        GrdID.Value(GrdID.Row, "IdDmCd"), "S","T", md_InCnvRt, , , , , .Value(RowNum, "IRRMSZ"), .Value(RowNum, "IRRMSZ2"), 0)
          Else
            .Value(RowNum, "IrRmIRt") = 0
          End If
        ElseIf ws_RmCtg = "X" Then
          If moCn.RecSeek("select 'x' from Param where PTyp= 'RMSCTG' and PMCd= '" + ws_RmCtg + "' " + _
                          "and PSCd= '" + moCn.GetFldVal("Select RmSCtg from RmMst where " + _
                          "RmCd= '" + .Value(RowNum, "IrRmCd") + "' " + IIF(ADC("wIdPrtKey") = ctCurrPrtn, " And RmPrtKey='" + ctCurrPrtn + "' ", "")) + "' and PValue= 'S' ") Then
          
            Dim wf_BaseLme As Single
            wf_BaseLme = moCn.GetFldVal("Select (case B.RmCtg when 'G' then " + CStr(ADC("wIdLmgSal")) + " " + _
                         "when 'P' then " + CStr(ADC("wIdLmpSal")) + " when 'S' then " + CStr(ADC("wIdLmsSal")) + " " + _
                         "when 'L' then " + CStr(ADC("wIdLmlSal")) + " else 0 end) from RmMst B where " + _
                         "RmCd= (Select Rm.RmBaseCd from RmMst Rm where Rm.RmCd= '" + .Value(RowNum, "IrRmCd") + "'" + IIF(ADC("wIdPrtKey") = ctCurrPrtn, " and Rm.RmPrtKey='" + ctCurrPrtn + "'", "") + ")")
            
            If wf_BaseLme > 0 Then '6.1-55 ADC("wIdCmCd") changed to wCmCd
              .Value(RowNum, "IrRmIRt") = GetRmRt(wCmCd, .Value(RowNum, "IrRmCd"), 0, 0, 0, "S", wf_BaseLme, "T", , , , ADC("wIdCmCurCd"))
            End If
          End If
        End If
      End If
    
    Case Is = UCase("IrRmIVal")   '*** recalc on IrRmIWt, IrRmIRt, IrMainMet
      If .Mode = fgmnorm Or .Mode = fgmCopy Then
        If .Value(RowNum, "IrMainMet") = "Y" And UCase(ms_IdTcTyp) = UCase(ctInTcTyp) Then
          '****** Sachin 2.13.0 - 31-05-2006 - Updated for Q/W Change
          '****** Sachin 3.02 - Current Partition
          .Value(RowNum, "IrRmIVal") = IIF(moCn.GetFldVal("Select RmQw from RmMst where " + _
                                       "RmCd= '" + .Value(RowNum, "IrRmCd") + "'" + IIF(ADC("wIdPrtKey") = ctCurrPrtn, " and RmPrtKey='" + ctCurrPrtn + "' ", "")) = "Q", _
                                       .Value(RowNum, "IrRmQty"), GrdID.Value(GrdID.Row, "IdIGldAsWt")) * .Value(RowNum, "IrRmIRt")
          '****** Sachin 2.13.0 - 31-05-2006 ******
        ElseIf .Value(RowNum, "IrMainMet") = "N" Or UCase(ms_IdTcTyp) = UCase(ctCTBTcTyp) Then
          '.Value(RowNum, "IrRmIVal") = .Value(RowNum, "IrRmIWt") * .Value(RowNum, "IrRmIRt")
          .Value(RowNum, "IrRmIVal") = IIF(moCn.GetFldVal("Select RmQw from RmMst where " + _
                                       "RmCd= '" + .Value(RowNum, "IrRmCd") + "'" + IIF(ADC("wIdPrtKey") = ctCurrPrtn, " and RmPrtKey='" + ctCurrPrtn + "' ", "")) = "Q", _
                                       .Value(RowNum, "IrRmQty"), .Value(RowNum, "IrRmIWt")) _
                                       * .Value(RowNum, "IrRmIRt")
        End If
      End If
    Case Is = UCase("IrRmZRt")
      If (.Mode = fgmnorm Or .Mode = fgmCopy) Then
        If (UCase(ms_IdTcTyp) = UCase(ctCTBTcTyp)) Then
          .Value(RowNum, "IrRmZRt") = MWLib.Div(.Value(RowNum, "IrRmIRt"), md_InCnvRt)
        
        End If
      End If
    Case Is = UCase("IrRmZVal")   '*** recalc on IrRmZWt, IrRmZRt, IrMainMet
      Dim wGldLsPrc As Single
      If .Mode = fgmnorm Or .Mode = fgmCopy Then
        If UCase(ms_IdTcTyp) = UCase(ctInTcTyp) Then
          If .Value(RowNum, "IrMainMet") = "Y" Then
            .Value(RowNum, "IrRmZVal") = IIF(moCn.GetFldVal("Select RmQw from RmMst where " + _
                                         "RmCd= '" + .Value(RowNum, "IrRmCd") + "'" + IIF(ADC("wIdPrtKey") = ctCurrPrtn, " and RmPrtKey='" + ctCurrPrtn + "' ", "")) = "Q", _
                                         .Value(RowNum, "IrRmQty"), GrdID.Value(GrdID.Row, "IdZGldAsWt")) * .Value(RowNum, "IrRmZRt")
          ElseIf .Value(RowNum, "IrMainMet") = "N" Then
            .Value(RowNum, "IrRmZVal") = IIF(moCn.GetFldVal("Select RmQw from RmMst where " + _
                                         "RmCd= '" + .Value(RowNum, "IrRmCd") + "'" + IIF(ADC("wIdPrtKey") = ctCurrPrtn, " and RmPrtKey='" + ctCurrPrtn + "' ", "")) = "Q", _
                                         .Value(RowNum, "IrRmQty"), .Value(RowNum, "IrRmZWt")) _
                                         * .Value(RowNum, "IrRmZRt")
          End If
        ElseIf UCase(ms_IdTcTyp) = UCase(ctCTBTcTyp) Then
          .Value(RowNum, "IrRmZVal") = IIF(moCn.GetFldVal("Select RmQw from RmMst where " + _
                                         "RmCd= '" + .Value(RowNum, "IrRmCd") + "'" + IIF(ADC("wIdPrtKey") = ctCurrPrtn, " and RmPrtKey='" + ctCurrPrtn + "' ", "")) = "Q", _
                                         .Value(RowNum, "IrRmQty"), (.Value(RowNum, "IrRmZWt"))) _
                                         * .Value(RowNum, "IrRmZRt")
        End If
      End If
    Case Is = UCase("IrSetAVal")    '*** recalc on IrSetSCd, IrSetARt, IrRmQty  '6.1-55 GrdID.Value(GrdID.Row, "IdExpCmCd") changed to wCmCd
      If .Mode = fgmnorm Or .Mode = fgmCopy Then _
         .Value(RowNum, "IrSetAVal") = GetLabVal(wCmCd, "SET", _
                                       .Value(RowNum, "IrSetSCd"), _
                                       .Value(RowNum, "IrSetARt"), _
                                       .Value(RowNum, "IrRmQty"), 0, "Q", "C", "C", ADC("wIdCmCurCd"), _
                                       IIF(ws_OmLkUpLabRt <> "", ws_OmLkUpLabRt, ""))   '###
    Case Is = UCase("IrSetIRt")
      If (.Mode = fgmnorm Or .Mode = fgmCopy) And (UCase(ms_IdTcTyp) = UCase(ctCTBTcTyp)) Then
        '*** (Bef 2.13) .Value(RowNum, "IrSetIRt") = GetLabRt(wCmCd, "SET", .Value(RowNum, "IrSetSCd"), "S", 0, ms_CmCtg)
        '*** (Jen 2.13)
        .Value(RowNum, "IrSetIRt") = GetLabRt(wCmCd, "SET", .Value(RowNum, "IrSetSCd"), "S", 0, ms_CmCtg, , , , ADC("wIdCmCurCd"))
        '*** (Jen 2.13)
      End If
    Case Is = UCase("IrSetIVal")    '*** recalc on IrSetSCd, IrSetIRt, IrRmQty
      ' **** Manali 3.03 - 05/07/08 - Om Lk Up fields passed if not blank
      If .Mode = fgmnorm Or .Mode = fgmCopy Then _
         .Value(RowNum, "IrSetIVal") = GetLabVal(wCmCd, "SET", _
                                         .Value(RowNum, "IrSetSCd"), _
                                         .Value(RowNum, "IrSetIRt"), _
                                         .Value(RowNum, "IrRmQty"), 0, "Q", "S", ms_CmCtg, ADC("wIdCmCurCd"), _
                                         IIF(ws_OmLkUpLabRt <> "", ws_OmLkUpLabRt, ""))
    Case Is = UCase("IrSetZRt")
      If (.Mode = fgmnorm Or .Mode = fgmCopy) And (UCase(ms_IdTcTyp) = UCase(ctCTBTcTyp)) Then
        .Value(RowNum, "IrSetZRt") = MWLib.Div(.Value(RowNum, "IrSetIRt"), md_InCnvRt)
      End If
    Case Is = UCase("IrSetZVal")    '*** recalc on IrSetSCd, IrSetZRt, IrRmQty
      If .Mode = fgmnorm Or .Mode = fgmCopy Then
        If UCase(ms_IdTcTyp) = UCase(ctInTcTyp) Then
          ' **** Manali 3.03 - 05/07/08 - Om Lk Up fields passed if not blank
          .Value(RowNum, "IrSetZVal") = GetLabVal(wCmCd, "SET", _
                                        .Value(RowNum, "IrSetSCd"), _
                                        .Value(RowNum, "IrSetZRt"), _
                                        .Value(RowNum, "IrRmQty"), 0, "Q", "S", ms_CmCtg, ADC("wIdCmCurCd"), _
                                        IIF(ws_OmLkUpLabRt <> "", ws_OmLkUpLabRt, ""))
        ElseIf UCase(ms_IdTcTyp) = UCase(ctCTBTcTyp) Then
          .Value(RowNum, "IrSetZVal") = .Value(RowNum, "IrSetZRt") * .Value(RowNum, "IrRmQty")
        End If
      End If
    Case Is = UCase("IrRmZPureRt")    '*** recalc on IrRmCd,IrRmZRt,IrRmZVal
      'If (.Mode = fgmnorm Or .Mode = fgmCopy) And UCase(ms_IdTcTyp) = UCase(ctInTcTyp) Then
      If (.Mode = fgmnorm Or .Mode = fgmCopy) And UCase(ms_IdTcTyp) = UCase(ctInTcTyp) Then
        '****** Sachin 3.02 - Current Partition
        wRmPurityZ = moCn.GetFldVal("Select RmPurityZ from RmMst where RmCd= '" + .Value(RowNum, "IrRmCd") + "'" + IIF(gs_Partition = ctCurrPrtn, " and RmPrtKey='" + ctCurrPrtn + "'", ""))
        
        '*** (Jen 2.13)
        'wf_CustmCdPurityZ = moCn.GetFldVal("Select R.RmPurityZ from RmMst R where R.RmCd= (Select R1.RmZCd from RmMst R1 where R1.RmCd= '" + .Value(RowNum, "IrRmCd") + "') ")
        '****** Sachin 3.02 - Current Partition
        ' **** Manali 3.03 - 21/06/08 - Pd Mod - 'L' included
        wf_CustmCdPurityZ = moCn.GetFldVal("Select R.RmPurityZ* " + _
                            "(case when (R.RmCtg= 'G' and " + CStr(ADC("wIdGldFiness")) + "> 0) then " + CStr(ADC("wIdGldFiness")) + " " + _
                            "      when (R.RmCtg= 'P' and " + CStr(ADC("wIdPlFiness")) + "> 0) then " + CStr(ADC("wIdPlFiness")) + " " + _
                            "      when (R.RmCtg= 'S' and " + CStr(ADC("wIdSlFiness")) + "> 0) then " + CStr(ADC("wIdSlFiness")) + " " + _
                            "      when (R.RmCtg= 'L' and " + CStr(ADC("wIdPdFiness")) + "> 0) then " + CStr(ADC("wIdPdFiness")) + " " + _
                            "else 1 end) from RmMst R where R.RmCd= (Select R1.RmZCd from RmMst R1 where R1.RmCd= '" + .Value(RowNum, "IrRmCd") + "'" + IIF(ADC("wIdPrtKey") = ctCurrPrtn, " and R1.RmPrtKey='" + ctCurrPrtn + "'", "") + ")" + IIF(ADC("wIdPrtKey") = ctCurrPrtn, " and R.RmPrtKey='" + ctCurrPrtn + "' ", ""))
        
        '.Value(RowNum, "IrRmZPureRt") = IIF(wRmPurityZ > 0, MWLib.Div(.Value(RowNum, "IrRmZRt"), wRmPurityZ) * wf_CustmCdPurityZ, .Value(RowNum, "IrRmZRt"))
        '****** Sachin 2.13.0 - 31-05-2006 - Updated for Q/W Change
        '.Value(RowNum, "IrRmZPureRt") = IIF(wRmPurityZ > 0, MWLib.Div(IIF(moCn.GetFldVal("Select RmQw from RmMst where " + _
                                         "RmCd= '" + .Value(RowNum, "IrRmCd") + "' ") = "Q", _
                                         MWLib.Div(.Value(RowNum, "IrRmZVal"), .Value(RowNum, "IrRmZWt")), .Value(RowNum, "IrRmZRt")), wRmPurityZ) * wf_CustmCdPurityZ, .Value(RowNum, "IrRmZRt"))
                                         
        '****** Sachin 3.02 - Current Partition
        .Value(RowNum, "IrRmZPureRt") = IIF(moCn.GetFldVal("Select RmQw from RmMst where " + _
                                         "RmCd= '" + .Value(RowNum, "IrRmCd") + "'" + IIF(ADC("wIdPrtKey") = ctCurrPrtn, " and RmPrtKey='" + ctCurrPrtn + "' ", "")) = "Q", _
                                         MWLib.Div(.Value(RowNum, "IrRmZVal"), .Value(RowNum, "IrRmZWt")), _
                                         .Value(RowNum, "IrRmZRt")) / IIF(wRmPurityZ > 0, wRmPurityZ, 1) * _
                                         IIF(wf_CustmCdPurityZ > 0, wf_CustmCdPurityZ, 1)
        '****** Sachin 2.13.0 - 31-05-2006 ******
        
        
        '*** (Jen 2.13)
        
        '*** (Bef 2.13) .Value(RowNum, "IrRmZPureRt") = IIF(wRmPurityZ > 0, MWLib.Div(.Value(RowNum, "IrRmZRt"), wRmPurityZ), .Value(RowNum, "IrRmZRt"))
      End If
    Case Is = UCase("IrRmZCd")    '*** recalc on IrRmCd
      '****** Sachin 3.02 - Current Partition
      If .Mode = fgmnorm Or .Mode = fgmCopy Then _
         .Value(RowNum, "IrRmZCd") = moCn.GetFldVal("Select RmZCd from RmMst " + _
                                "where RmCd= '" + .Value(RowNum, "IrRmCd") + "'" + IIF(ADC("wIdPrtKey") = ctCurrPrtn, " and RmPrtKey='" + ctCurrPrtn + "'", ""))
    Case Is = UCase("IrMainMet")  '*** recalc on IrRmCd   '###
      If .Mode = fgmnorm Then _
         .Value(RowNum, "IrMainMet") = "N"
    
    '****** Sachin 2.13.0 - 29-05-2006 ******
    Case Is = UCase("wQwSal")
      'For a new row (manually entered) in some cases when RMCD is entered this field is not getting recalculated
      'For Normal Mode this  update has been shifted to the ColChange event due to above problem
      '****** Sachin 3.02 - Current Partition
      If .Mode = fgminit Or .Mode = fgmCopy Then _
         .Value(RowNum, "WQWSAL") = moCn.GetFldVal("Select RmQw From RmMst Where RmCd='" + .Value(RowNum, "IrRmCd") + "'" + IIF(ADC("wIdPrtKey") = ctCurrPrtn, " and RmPrtKey='" + ctCurrPrtn + "'", ""))
    '****** Sachin 2.13.0 - 29-05-2006 ******
    
    
    ' Sachin 4.1.0.0 ******
    Case Is = UCase("IrCustRmCd")
      If .Mode = fgmnorm Then
         .Value(RowNum, "IrCustRmCd") = moCn.GetFldVal(" Select IsNull((Select CrCustRmCd from CustRm " + _
                                                            "  Join RmMst On RmCd='" + .Value(RowNum, "IrRmCd") + "' " + IIF(ADC("wIdPrtKey") = ctCurrPrtn, " And RmPrtKey='" + ctCurrPrtn + "' ", "") + _
                                                            " Where " + IIF(gs_Partition = ctCurrPrtn, " CrPrtKey='" + ctCurrPrtn + "' And ", "") + _
                                                            " CrCtg = '" + ms_CmCtg + "' And CrCd = '" + ADC("wIdCmCd") + "' and CrRmCtg=RmCtg And " + _
                                                            " CrOurRmCd ='" + .Value(RowNum, "IrRmCd") + "'  and " + CStr(.Value(RowNum, "IrRmSz")) + " Between CrFrLn and CrToLn),'')")
        ADC("WIRCUSTRMCD") = .Value(RowNum, "IrCustRmCd")
      End If
      
    '****** Sachin 3.02.0 - 16/10/07 ******
    '****** Sachin 3.02 - 14-01-08 - Shifted to InitKey
    '    Case Is = UCase("IrRmCtg")
    '      If .Mode = fgmnorm Or .Mode = fgminit Or .Mode = fgmCopy Then _
    '         .Value(RowNum, "IrRmCtg") = moCn.GetFldVal("Select RmCtg From RmMst Where RmCd='" + .Value(RowNum, "IrRmCd") + "'" + IIF(ADC("wIdPrtKey") = ctCurrPrtn, " and RmPrtKey='" + ctCurrPrtn + "'", ""))
    '
    '    Case Is = UCase("IrRmSCtg")
    '      If .Mode = fgmnorm Or .Mode = fgminit Or .Mode = fgmCopy Then _
    '         .Value(RowNum, "IrRmSCtg") = moCn.GetFldVal("Select RmSCtg From RmMst Where RmCd='" + .Value(RowNum, "IrRmCd") + "'" + IIF(ADC("wIdPrtKey") = ctCurrPrtn, " and RmPrtKey='" + ctCurrPrtn + "'", ""))
    '****** Sachin 3.02.0 - 16/10/07 ******
    
    End Select
  End With
End Sub
Private Sub GrdIR_ColChange(ByVal RowNum As Integer, ByVal ColName As String, ByVal NewValue As Variant, Cancel As Boolean, ErrMsg As String)
  '*** IrRmCd         valid RmCd from RmMst(RmZ= 'N') or Blank
  '*** IrSetSCd       valid PSCd('LABSCD', 'SET')  or Blank
  '*** IrMainMet      valid PMCd('YN')
  
  With GrdIR
    Select Case ColName
    Case Is = UCase("IrRmCd")
      '****** Sachin 3.02 - Current Partition
      Cancel = (NewValue <> "") And (Not moCn.RecSeek("Select RmCd From RmMst Where RmZ='N' " + _
                                                      "And RmCd = '" + NewValue + "' and RmPrtKey='" + ctCurrPrtn + "'"))
      If Cancel = True Then ErrMsg = "Invalid Raw Material Code": Exit Sub
      
      '****** Sachin 2.13.0 [09-09-06] Shifted due to problem in Recalc in Normal Mode for some invoices
      '****** Sachin 3.02 - Current Partition
      .Value(RowNum, "WQWSAL") = moCn.GetFldVal("Select RmQw From RmMst Where RmCd='" + NewValue + "' and RmPrtKey='" + ctCurrPrtn + "'")
      '****** Sachin 2.13.0 [09-09-06] Shifted due to problem in Recalc in Normal Mode for some invoices
      
    Case Is = UCase("IrSetSCd")
      '*** (Jen 3.01)
      If .Mode = fgmnorm Then
      '*** (Jen 3.01)
        Cancel = (NewValue <> "") And (Not moCn.RecSeek("Select PSCd From Param Where " + _
                 "PTyp= 'LABSCD' And PMCd='SET' And PSCd = '" + NewValue + "'"))
        ErrMsg = "Invalid Setting Code": Exit Sub
      End If
    Case Is = UCase("IrMainMet")
      '*** (Jen 3.01)
      If .Mode = fgmnorm Then
      '*** (Jen 3.01)
    
        Cancel = Not moCn.RecSeek("Select PMCd From Param Where " + _
                                  "PTyp= 'YN' And PMCd = '" + NewValue + "'")
        If Cancel = True Then ErrMsg = "Invalid Main Metal Y/N Option": Exit Sub
        
        '*** (Jen 2.13)
        '****** Sachin 3.02 - Current Partition
        Cancel = NewValue = "Y" And moCn.RecSeek("select 'x' from InvRm where IrCoCd= '" + gs_CoCd + "' and " + _
                        "IrTc= '" + ADC("wIdTc") + "' and IrYy= '" + ADC("wIdYy") + "' and " + _
                        "IrChr= '" + ADC("wIdChr") + "' and IrNo= " + CStr(ADC("wIdNo")) + " and " + _
                        "IrSr= " + CStr(GrdID.Value(GrdID.Row, "IdSr")) + " and " + _
                        "IrSrNo<> " + CStr(.Value(RowNum, "IrSrNo")) + " and IrPrtKey='" + ctCurrPrtn + "' " + _
                        "and IrMainMet = 'Y' ")
        If Cancel = True Then ErrMsg = "Cannot Have More Than One Main Metal": Exit Sub
        '*** (Jen 2.13)
      
      End If
      
' Zubin 212
    Case Is = UCase("IrLotNo")
      '*** (Jen 3.01)
      If .Mode = fgmnorm Then
      '*** (Jen 3.01)
        '****** Sachin 3.02 - Current Partition
        Cancel = NewValue <> "" And Not moCn.RecSeek("Select 'x' from SYyMm where SCoCd= '" + gs_CoCd + "' and SLotNo= '" + NewValue + "' and SLoc= 'FG' and SPrtKey='" + ctCurrPrtn + "'")
        If Cancel = True Then ErrMsg = "Invalid Lot No.": Exit Sub
      End If
' Zubin 212
    End Select
  End With
End Sub
Private Function GetLaborSubCdLabLogic(ByVal ws_LabMCd As String, ByVal ws_LabSCd As String)

  GetLaborSubCdLabLogic = moCn.GetFldVal("Select ForLabLogic.PValue1 from Param join Param ForLabLogic on Param.PValue3=ForLabLogic.PSCd and  ForLabLogic.PTyp='LABLOGIC'  and ForLabLogic.PMCd=Param.PMCd where Param.PTyp='LABSCD' and Param.PMCd='" + ws_LabMCd + "' and Param.PSCd='" + ws_LabSCd + "' ")
End Function
Private Sub GrdIR_OnDelete(ByVal RmRow As Integer, Cancel As Boolean, ErrMsg As String)
  If GrdID.Mode = fgmnorm And GrdIL.Mode = fgmnorm Then
    Dim LabRow As Integer
    
    With GrdIL
      For LabRow = 1 To .Rows - 1
        Call GetLab_izVal(LabRow, RmRow)
      Next LabRow
      
    End With
  End If
    
End Sub
Private Function GetLab_izVal(ByVal wi_i As Integer, Optional ByVal RowNum As Integer)
 If IsMissing(RowNum) Then
    RowNum = 0
  End If
    Dim wCmCd As String, ws_OmLkUpLabRt As String
    If UCase(ms_IdTcTyp) = UCase(ctInTcTyp) And GrdID.Rows - 1 > 0 Then
    wCmCd = GrdID.Value(GrdID.Row, "IdExpCmCd")
    ' ***** Manali 3.03 - 05/07/08 - OmLkUp
    ws_OmLkUpLabRt = ""
   If GrdID.Value(GrdID.Row, "IdExpOdNo") <> 0 Then
    ws_OmLkUpLabRt = moCn.GetFldVal("Select OmLkUpLabRt from OrdMst where " + _
                                     "OmCoCd= '" + gs_CoCd + "' and OmTc= '" + GrdID.Value(GrdID.Row, "IdExpOdTc") + "' and " + _
                                     "OmYy= '" + GrdID.Value(GrdID.Row, "IdExpOdYy") + "' " + _
                                     "and OmChr= '" + GrdID.Value(GrdID.Row, "IdExpOdChr") + "' and " + _
                                     "OmNo= " + CStr(GrdID.Value(GrdID.Row, "IdExpOdNo")) + " and OmPrtKey='" + ctCurrPrtn + "' ")
   End If
  ElseIf UCase(ms_IdTcTyp) = UCase(ctCTBTcTyp) Then
    wCmCd = ADC("wIdCmCd")
  End If

    With GrdIL
        If .Value(wi_i, "ilMCd") = ctBrkLMCd And .Value(wi_i, "iLSCd") = ctBrkLSCd And Not .IsDel(wi_i) Then
          .Value(wi_i, "iLiVal") = .Value(wi_i, "iLiRt") / 100 * GetDCRmVal("IrRmiVal", RowNum)
          .Value(wi_i, "iLZVal") = .Value(wi_i, "iLZRt") / 100 * GetDCRmVal("IrRmZVal", RowNum)
          .SaveRec (wi_i)
        ' ****** Manali 3.5.0 - 18/05/09 - "DHC" and "CHC"
        ' ***** Manali 3.9.0 - ctDCHCLSCd='-' removed
        ElseIf .Value(wi_i, "ilMCd") = ctDHCLMCd And Not .IsDel(wi_i) Then      ''''And .Value(wi_i, "iLSCd") = ctDCHCLSCd
          .Value(wi_i, "iLiVal") = .Value(wi_i, "iLiRt") / 100 * GetDCRmVal("IrRmiVal", RowNum, "D")
          .Value(wi_i, "iLZVal") = .Value(wi_i, "iLZRt") / 100 * GetDCRmVal("IrRmZVal", RowNum, "D")
          .SaveRec (wi_i)
        ElseIf .Value(wi_i, "ilMCd") = ctGHCLMCd And Not .IsDel(wi_i) Then      ''''And .Value(wi_i, "iLSCd") = ctDCHCLSCd
          Dim wLaborSubLogic As String
          wLaborSubLogic = GetLaborSubCdLabLogic(GrdIL.Value(wi_i, "ilMCd"), GrdIL.Value(wi_i, "ilSCd"))
          
          .Value(wi_i, "iLiVal") = .Value(wi_i, "iLiRt") / 100 * GetDCRmVal(IIF(wLaborSubLogic = "P", "IrRmIWt", "IrRmiVal"), RowNum, ctGHCLMCd, wLaborSubLogic)
          .Value(wi_i, "iLZVal") = .Value(wi_i, "iLZRt") / 100 * GetDCRmVal(IIF(wLaborSubLogic = "P", "IrRmZWt", "IrRmZVal"), RowNum, ctGHCLMCd, wLaborSubLogic)
          .SaveRec (wi_i)
          
        ElseIf .Value(wi_i, "ilMCd") = ctCHCLMCd And Not .IsDel(wi_i) Then      '''' And .Value(wi_i, "iLSCd") = ctDCHCLSCd
          .Value(wi_i, "iLiVal") = .Value(wi_i, "iLiRt") / 100 * GetDCRmVal("IrRmiVal", RowNum, "C")
          .Value(wi_i, "iLZVal") = .Value(wi_i, "iLZRt") / 100 * GetDCRmVal("IrRmZVal", RowNum, "C")
          .SaveRec (wi_i)
        ' ****** Manali 3.5.0 - 18/05/09 - "DHC" and "CHC"
         ElseIf .Value(wi_i, "ilMCd") = "CHN" And ADC("wChnCustYn") = "Y" And Not .IsDel(wi_i) Then      '''' And .Value(wi_i, "iLSCd") = ctDCHCLSCd
          .Value(wi_i, "iLiVal") = .Value(wi_i, "iLiRt") * GetDCRmVal("IrRmIWt", RowNum, "CHN")
          .Value(wi_i, "iLZVal") = .Value(wi_i, "iLZRt") * GetDCRmVal("IrRmZWt", RowNum, "CHN")
          .SaveRec (wi_i)
        ElseIf .Value(wi_i, "ilMCd") = ctCDWLMCd And .Value(wi_i, "ilQw") = "W" And Not .IsDel(wi_i) Then
          .Value(wi_i, "iLiVal") = GetLabVal(wCmCd, .Value(wi_i, "ilMCd"), _
                                     .Value(wi_i, "ilSCd"), .Value(wi_i, "iLiRt"), _
                                     .Value(wi_i, "iLQty"), MWLib.Div(GetDCRmVal("IrRmIWt", RowNum, ctCDWLMCd, GetLaborSubCdLabLogic(GrdIL.Value(wi_i, "ilMCd"), GrdIL.Value(wi_i, "ilSCd"))), GrdID.Value(GrdID.Row, "IdQty")), _
                                     .Value(wi_i, "iLQw"), "S", ms_CmCtg, ADC("wIdCmCurCd"), _
                                     IIF(ws_OmLkUpLabRt <> "", ws_OmLkUpLabRt, "")) * _
                                     GrdID.Value(GrdID.Row, "IdQty")
          .Value(wi_i, "iLzVal") = GetLabVal(wCmCd, .Value(wi_i, "ilMCd"), _
                                     .Value(wi_i, "ilSCd"), .Value(wi_i, "iLzRt"), _
                                     .Value(wi_i, "iLQty"), MWLib.Div(GetDCRmVal("IrRmZWt", RowNum, ctCDWLMCd, GetLaborSubCdLabLogic(GrdIL.Value(wi_i, "ilMCd"), GrdIL.Value(wi_i, "ilSCd"))), GrdID.Value(GrdID.Row, "IdQty")), _
                                     .Value(wi_i, "iLQw"), "S", ms_CmCtg, ADC("wIdCmCurCd"), _
                                     IIF(ws_OmLkUpLabRt <> "", ws_OmLkUpLabRt, "")) * _
                                     GrdID.Value(GrdID.Row, "IdQty")
          .SaveRec (wi_i)
       ElseIf .Value(wi_i, "ilMCd") = ctCBRKLMCd And .Value(wi_i, "ilQw") = "W" And Not .IsDel(wi_i) Then
          .Value(wi_i, "iLiVal") = GetLabVal(wCmCd, .Value(wi_i, "ilMCd"), _
                                     .Value(wi_i, "ilSCd"), .Value(wi_i, "iLiRt"), _
                                     .Value(wi_i, "iLQty"), MWLib.Div(GetDCRmVal("iRrMiwT", , "C", ""), GrdID.Value(GrdID.Row, "IdQty")), _
                                     .Value(wi_i, "iLQw"), "S", ms_CmCtg, ADC("wIdCmCurCd"), _
                                     IIF(ws_OmLkUpLabRt <> "", ws_OmLkUpLabRt, "")) * _
                                     GrdID.Value(GrdID.Row, "IdQty")
          .Value(wi_i, "iLzVal") = GetLabVal(wCmCd, .Value(wi_i, "ilMCd"), _
                                     .Value(wi_i, "ilSCd"), .Value(wi_i, "iLzRt"), _
                                     .Value(wi_i, "iLQty"), MWLib.Div(GetDCRmVal("iRrMzwT", , "C", ""), GrdID.Value(GrdID.Row, "IdQty")), _
                                     .Value(wi_i, "iLQw"), "S", ms_CmCtg, ADC("wIdCmCurCd"), _
                                     IIF(ws_OmLkUpLabRt <> "", ws_OmLkUpLabRt, "")) * _
                                     GrdID.Value(GrdID.Row, "IdQty")

'          .Value(wi_i, "iLiVal") = .Value(wi_i, "iLiRt") * GetDCRmVal("IrRmIWt", RowNum, ctCDWLMCd)
'          .Value(wi_i, "iLZVal") = .Value(wi_i, "iLZRt") * GetDCRmVal("IrRmZWt", RowNum, ctCDWLMCd)
          .SaveRec (wi_i)
        End If
     End With
        ''If .Value(wi_i, "iLMCd") = ctAddLMCd And .Value(wi_i, "iLSCd") = ctAddLSCd And Not .IsDel(wi_i) Then
        ''  wi_ZAdd = wi_i
        ''End If
  '*** (Jen 2.12) 08/02/06
End Function
Private Sub GrdIR_ReCalc(ByVal RowNum As Integer, ByVal ColName As String, ByVal OldVal As Variant, ByVal SourceName As String, ByVal SourceColName As String)
  Call GrdIR_ReCalcRoutine(RowNum, ColName, OldVal, SourceName, SourceColName)
End Sub
Private Sub GrdIR_InitKey(ByVal RowNum As Integer)
  '*** Initialise Key Fields (i.e. IrTc, IrYy, IrChr, IrNo, IrSr)
  With GrdIR
    .Store "IrCoCd", gs_CoCd
    .Store "IrTc", ADC("wIdTc")
    .Store "IrYy", ADC("wIdYy")
    .Store "IrChr", ADC("wIdChr")
    .Store "IrNo", ADC("wIdNo")
    .Store "IrSr", GrdID.Value(GrdID.Row, "IdSr")
    If mf_inspecialCopy Then
    .Store "Cpy", "Y"
    Else
    .Store "Cpy", ""
    End If

    '****** Sachin 3.02 Id fields in Joins
    If .IsNew(RowNum) Then
        .Store "IrIdIdNo", moCn.GetFldVal("Select IdIdNo From InvDsg " + _
                                          " Where IdCoCd='" + gs_CoCd + "' and IdTc='" + ADC("wIdTc") + "' and IdYy='" + ADC("wIdYy") + "' and IdChr='" + ADC("wIdChr") + "' and IdNo=" + CStr(ADC("wIdNo")) + " and IdSr=" + CStr(GrdID.Value(GrdID.Row, "IdSr")) + " and IdPrtKey='" + ctCurrPrtn + "' ")
        '.Store "IrDt", ADC("wIdDt")
        .Store "IrDt", moCn.GetFldVal("Select InDt From InvHd " + _
                                          " Where InCoCd='" + gs_CoCd + "' and InTc='" + ADC("wIdTc") + "' and InYy='" + ADC("wIdYy") + "' and InChr='" + ADC("wIdChr") + "' and InNo=" + CStr(ADC("wIdNo")) + " and InPrtKey='" + ctCurrPrtn + "' ")
    End If
       .Store "IrRmCtg", moCn.GetFldVal("Select RmCtg From RmMst Where RmCd='" + .Value(RowNum, "IrRmCd") + "'" + IIF(ADC("wIdPrtKey") = ctCurrPrtn, " and RmPrtKey='" + ctCurrPrtn + "'", ""))
       .Store "IrRmSCtg", moCn.GetFldVal("Select RmSCtg From RmMst Where RmCd='" + .Value(RowNum, "IrRmCd") + "'" + IIF(ADC("wIdPrtKey") = ctCurrPrtn, " and RmPrtKey='" + ctCurrPrtn + "'", ""))
    '****** Sachin 3.02 Id fields in Joins
    
    '****************** urmi 5/7/03 ***************************************'
    If UCase(ms_IdTcTyp) = UCase(ctInTcTyp) Then
      Dim wRs_Inv As MwfLib.MDORowSet
      '****** Sachin 3.02 - Current Partition
      Set wRs_Inv = moCn.OpenRes("Select InFrRmLoc, InFrRmDc, InToRmLoc, InToRmDc from InvHd where " + _
                         "InCoCd= '" + gs_CoCd + "' and InTc= '" + ADC("wIdTc") + "' and " + _
                         "InYy= '" + ADC("wIdYy") + "' and InChr= '" + ADC("wIdChr") + "' and " + _
                         "InNo= " + CStr(ADC("wIdNo")) + IIF(ADC("wIdPrtKey") = ctCurrPrtn, " and InPrtKey='" + ctCurrPrtn + "' ", ""))
      If Not (wRs_Inv.EOF Or wRs_Inv.BOF) Then
        .Store "IrFrRmLoc", wRs_Inv!InFrRmLoc
        .Store "IrFrRmLocTyp", moCn.GetFldVal("Select LocTyp From Loc Where LocCoCd='" + gs_CoCd + "'" + _
                                                " And LocCd='" + wRs_Inv!InFrRmLoc + "'")
        .Store "IrFrRmDc", wRs_Inv!InFrRmDc
        .Store "IrToRmLoc", wRs_Inv!InToRmLoc
        .Store "IrToRmLocTyp", moCn.GetFldVal("Select LocTyp From Loc Where LocCoCd='" + gs_CoCd + "'" + _
                                                " And LocCd='" + wRs_Inv!InToRmLoc + "'")
        .Store "IrToRmDc", wRs_Inv!InToRmDc
      End If
      Set wRs_Inv = Nothing
    
      '*** (Bef 2.12)
      '.Store "IrFrRmLoc", moCn.GetFldVal("Select InFrRmLoc from InvHd where " + _
      '                                   "InCoCd= '" + gs_CoCd + "' and InTc= '" + ADC("wIdTc") + "' and " + _
      '                                   "InYy= '" + ADC("wIdYy") + "' and InChr= '" + ADC("wIdChr") + "' and " + _
      '                                   "InNo= " + CStr(ADC("wIdNo")))
      '.Store "IrFrRmDc", moCn.GetFldVal("Select InFrRmDc from InvHd where " + _
      '                                   "InCoCd= '" + gs_CoCd + "' and InTc= '" + ADC("wIdTc") + "' and " + _
      '                                   "InYy= '" + ADC("wIdYy") + "' and InChr= '" + ADC("wIdChr") + "' and " + _
      '                                   "InNo= " + CStr(ADC("wIdNo")))
      '.Store "IrToRmLoc", moCn.GetFldVal("Select InToRmLoc from InvHd where " + _
      '                                   "InCoCd= '" + gs_CoCd + "' and InTc= '" + ADC("wIdTc") + "' and " + _
      '                                   "InYy= '" + ADC("wIdYy") + "' and InChr= '" + ADC("wIdChr") + "' and " + _
      '                                   "InNo= " + CStr(ADC("wIdNo")))
      '.Store "IrToRmDc", moCn.GetFldVal("Select InToRmDc from InvHd where " + _
      '                                   "InCoCd= '" + gs_CoCd + "' and InTc= '" + ADC("wIdTc") + "' and " + _
      '                                   "InYy= '" + ADC("wIdYy") + "' and InChr= '" + ADC("wIdChr") + "' and " + _
      '                                   "InNo= " + CStr(ADC("wIdNo")))
      '*** (Bef 2.12)
    Else
      .Store "IrFrRmLoc", ""
      .Store "IrFrRmDc", ""
      .Store "IrToRmLoc", ""
      .Store "IrToRmDc", ""
      .Store "IrFrRmLocTyp", ""
      .Store "IrToRmLocTyp", ""
    End If
    '****************** urmi 5/7/03 ***************************************'
  End With
End Sub
Private Sub GrdIR_RowSave(ByVal RmRow As Integer, Cancel As Boolean, ErrMsg As String)
  If GrdID.Mode = fgmnorm And GrdIL.Mode = fgmnorm Then
    Dim LabRow As Integer
    
    With GrdIL
      For LabRow = 1 To .Rows - 1
        Call GetLab_izVal(LabRow)
      Next LabRow
      
    End With
  End If
   
End Sub
Private Sub GrdIR_Valid()

ADC("WIRCUSTRMCD") = ""     '4.1.0.0

'  '*** (Jen 2.12)
'  Dim wf_AddLab As Double, wi_i As Integer
'  Dim wf_RmZVal As Double, wf_LZVal As Double
'  wf_RmZVal = moCn.GetFldVal("Select sum(IrRmZVal + IrSetZVal) from InvRm " + _
'              "where IrCoCd= '" + gs_CoCd + "' and " + _
'              "IrTc= '" + ADC("wIdTc") + "' and IrYy= '" + ADC("wIdYy") + "' and " + _
'              "IrChr= '" + ADC("wIdChr") + "' and IrNo= " + CStr(ADC("wIdNo")) + " and " + _
'              "IrSr= " + CStr(GrdID.Value(GrdID.Row, "IdSr")))
'  wf_LZVal = moCn.GetFldVal("Select sum(IlZVal) from InvLab " + _
'              "where IlCoCd= '" + gs_CoCd + "' and " + _
'              "IlTc= '" + ADC("wIdTc") + "' and IlYy= '" + ADC("wIdYy") + "' and " + _
'              "IlChr= '" + ADC("wIdChr") + "' and IlNo= " + CStr(ADC("wIdNo")) + " and " + _
'              "IlSr= " + CStr(GrdID.Value(GrdID.Row, "IdSr")) + " and IlMCd<> '" + ctAddLMCd + "'")
'
'  wf_AddLab = GrdID.Value(GrdID.Row, "IdIFOB") - (wf_RmZVal + wf_LZVal + GrdID.Value(GrdID.Row, "IdZAdd"))
'
'  With GrdIL
'    For wi_i = 1 To (.Rows - 1)
'      If .Value(wi_i, "iLMCd") = ctAddLMCd And .Value(wi_i, "iLSCd") = ctAddLSCd Then
'        .Value(wi_i, "iLZRt") = wf_AddLab
'        .Value(wi_i, "ilQw") = "Q"
'        .Value(wi_i, "iLZVal") = wf_AddLab
'        .SaveRec (wi_i)
'        Exit For
'      End If
'    Next wi_i
'  End With
'  '*** (Jen 2.12)
End Sub
Private Sub GrdIL_RowWhen(ByVal RowNum As Integer)
  '*** If New Row then IlSrNo= max(IlSrNo) of grid GrdIL + 1

  With GrdIL
    If .Value(RowNum, "IlSrNo") = 0 Then .Value(RowNum, "IlSrNo") = .MaxVal("IlSrNo") + 1
    If .Value(RowNum, "ilCstQw") = "" Then .Value(RowNum, "ilCstQw") = "Q"    '*** (Jen 2.14 Next) (20/03/07)
    If .IsNew(RowNum) Then .Value(RowNum, "IlPrtKey") = ctCurrPrtn      '****** Sachin 3.02 - 28-12-07
  End With
End Sub
Private Sub GrdIL_ColWhen(ByVal RowNum As Integer, ByVal ColName As String, Cancel As Boolean, ErrMsg As String, ByVal Action As MwfCtl.en_fgAction)
  '*** Do not allow the user to edit any field if InCls = 'Y' in InvHd Table
  '*** Give help for iLMCd, iLSCd
  '*** Do not allow the user to edit iLQty when iLQw= 'W' (i.e. the Labour value is calculated on Wt and not on Qty)
  
  '****** Sachin 3.02 - Locking the Previous Partition Records
  If ADC("WIDPRTKEY") <> ctCurrPrtn Then Cancel = True: ErrMsg = "Cannot Edit Previous Partition Records": Exit Sub
  
  If mf_DtLocked = True Then Cancel = True: ErrMsg = "Cannot Edit Locked Entry": Exit Sub
  If mInCls = "Y" Then Cancel = True: ErrMsg = "Cannot Edit As Invoice Is Closed": Exit Sub
  
  '*** (Jen 2.13)
  '****** Sachin 3.02 - Current Partition
  If moCn.RecSeek("Select 'x' from InvZRm where IzCoCd= '" + gs_CoCd + "' and IzTc= '" + ADC("wIdTc") + "' and " + _
                  "IzYy=  '" + ADC("wIdYy") + "' and IzChr=  '" + ADC("wIdChr") + "' and " + _
                  "IzNo=  " + CStr(ADC("wIdNo")) + " and IzPrtKey='" + ctCurrPrtn + "' ") Then
    Cancel = True: ErrMsg = "Cannot Edit As Custom Matching Has Been Done": Exit Sub
  End If
  '*** (Jen 2.13)
    
    
    
  Select Case UCase(ColName)
  Case Is = UCase("iLMCd")
    Call HlpList.PMCd("LABMCD")
  Case Is = UCase("iLSCd")
    Call HlpList.PSCd("LABSCD", GrdIL.Value(RowNum, "iLMCd"))
  Case Is = UCase("iLQty")
    '*** (Bef 2.14 Next 20/03/07)
    'If GrdIL.Value(RowNum, "iLQw") = "W" Then _
    '   Cancel = True: ErrMsg = "Cannot Enter Labour Quantity": Exit Sub
    '*** (Bef 2.14 Next 20/03/07)
       
    '*** (Jen 2.12) 06/01/06
    ' ****** Manali 3.5.0 - 18/05/09 - "DHC" and "CHC"
    ' ***** Manali 3.9.0 - ctDCHCLSCd='-' removed  ( And GrdIL.Value(RowNum, "iLSCd") = ctDCHCLSCd)
    If (GrdIL.Value(RowNum, "ilMCd") = ctBrkLMCd And GrdIL.Value(RowNum, "iLSCd") = ctBrkLSCd) Or _
       (GrdIL.Value(RowNum, "ilMCd") = ctAddLMCd And GrdIL.Value(RowNum, "iLSCd") = ctAddLSCd) Or _
       (GrdIL.Value(RowNum, "ilMCd") = ctDHCLMCd) Or _
       (GrdIL.Value(RowNum, "ilMCd") = ctGHCLMCd) Or _
       (GrdIL.Value(RowNum, "ilMCd") = ctCHCLMCd) Or (GrdIL.Value(RowNum, "ilMCd") = ctCBRKLMCd) Then
      Cancel = True: ErrMsg = "Cannot Edit Labour Quantity": Exit Sub
    End If
    '*** (Jen 2.12) 06/01/06
  
  '*** (Jen 2.12) 06/01/06
  Case Is = UCase("iLQw")
    ' ***** Manali 3.9.0 - ctDCHCLSCd='-' removed   -   And GrdIL.Value(RowNum, "iLSCd") = ctDCHCLSCd
    If (GrdIL.Value(RowNum, "ilMCd") = ctBrkLMCd And GrdIL.Value(RowNum, "iLSCd") = ctBrkLSCd) Or _
       (GrdIL.Value(RowNum, "ilMCd") = ctAddLMCd And GrdIL.Value(RowNum, "iLSCd") = ctAddLSCd) Or _
       (GrdIL.Value(RowNum, "ilMCd") = ctDHCLMCd) Or _
       (GrdIL.Value(RowNum, "ilMCd") = ctGHCLMCd) Or _
       (GrdIL.Value(RowNum, "ilMCd") = ctCHCLMCd) Or _
       (GrdIL.Value(RowNum, "ilMCd") = ctCBRKLMCd) Or _
       (GrdIL.Value(RowNum, "ilMCd") = ctCDWLMCd) Then
      Cancel = True: ErrMsg = "Cannot Edit Labour Q/W": Exit Sub
    End If
  Case Is = UCase("iLiRt")
    If GrdIL.Value(RowNum, "ilMCd") = ctAddLMCd And GrdIL.Value(RowNum, "iLSCd") = ctAddLSCd Then
      Cancel = True: ErrMsg = "Cannot Edit Invoice Labour Rate For " & ctAddLMCd: Exit Sub
    End If
  '*** (Jen 2.12) 06/01/06
  End Select
End Sub
Private Sub GrdIL_ReCalc(ByVal RowNum As Integer, ByVal ColName As String, ByVal OldVal As Variant, ByVal SourceName As String, ByVal SourceColName As String)

    Call GrdIL_RecalcRoutine(RowNum, ColName, OldVal, SourceName, SourceColName)
End Sub
Private Sub GrdIL_RecalcRoutine(ByVal RowNum As Integer, ByVal ColName As String, ByVal OldVal As Variant, ByVal SourceName As String, ByVal SourceColName As String)
  '*** iLQw on iLMCd,iLSCd                        Normal,Copy
      '   iLQw = PValue('LABSCD', iLMCd)
  '*** iLQty on iLQW                              Normal,Copy
      '   If iLQw = 'W' Then iLQty = 0
  '*** iLAVal on iLMCd,iLSCd,iLARt,iLQw,iLQty     Normal,Copy
      '   iLAVal = GetLabVal(IdExpCmCd, ilMCd, ilSCd, iLARt, iLQty, IdALabAsWt, iLQw, 'S') *
      '   (If iLQw= 'Q' then IdQty else 1)
  '*** iLiVal on iLMCd,iLSCd,iLiRt,iLQw,iLQty     Normal,Copy
      '   iLiVal = GetLabVal(IdExpCmCd, ilMCd, ilSCd, iLiRt, iLQty, IdILabAsWt, iLQw, 'S') *
      '   (If iLQw= 'Q' then IdQty else 1)
  '*** iLZVal on iLMCd,iLSCd,iLZRt,iLQw,iLQty     Normal,Copy
      '   iLZVal = GetLabVal(IdExpCmCd, ilMCd, ilSCd, iLZRt, iLQty, IdZLabAsWt, iLQw, 'S') *
      '   (If iLQw= 'Q' then IdQty else 1)
  Dim wGrdIdRow As Single, wCmCd As String
  Dim wi_i As Integer, wi_j As Integer, wd_LabRt As Double, wf_RmVal As Double, ws_RmCtg As String   '*** (Jen 2.12)
  Dim ws_OmLkUpLabRt As String   ' ***** Manali 3.03 - 05/07/08 - Om LkUp fields passed
  
  If GrdID.Row = GrdID.Rows Then Exit Sub   '*** Jen 2.12 Patch 2 (04/05/06)
  
  If (mf_StopRecalcDuringCopyIL = True And UCase(SourceName) <> "ILQW") Then Exit Sub
  If UCase(ms_IdTcTyp) = UCase(ctInTcTyp) And GrdID.Rows - 1 > 0 Then
    If UCase(SourceName) = UCase("GrdID") Then
      wCmCd = GrdID.Value(RowNum, "IdExpCmCd")
     ws_OmLkUpLabRt = ""
    If GrdID.Value(RowNum, "IdExpOdNo") <> 0 Then
      ws_OmLkUpLabRt = moCn.GetFldVal("Select OmLkUpLabRt from OrdMst where " + _
                                     "OmCoCd= '" + gs_CoCd + "' and OmTc= '" + GrdID.Value(RowNum, "IdExpOdTc") + "' and " + _
                                     "OmYy= '" + GrdID.Value(RowNum, "IdExpOdYy") + "' " + _
                                     "and OmChr= '" + GrdID.Value(RowNum, "IdExpOdChr") + "' and " + _
                                     "OmNo= " + CStr(GrdID.Value(RowNum, "IdExpOdNo")) + " and OmPrtKey='" + ctCurrPrtn + "' ")
     End If
    Else
      wCmCd = GrdID.Value(GrdID.Row, "IdExpCmCd")
      ws_OmLkUpLabRt = ""
     If GrdID.Value(GrdID.Row, "IdExpOdNo") <> 0 Then
      ws_OmLkUpLabRt = moCn.GetFldVal("Select OmLkUpLabRt from OrdMst where " + _
                                     "OmCoCd= '" + gs_CoCd + "' and OmTc= '" + GrdID.Value(GrdID.Row, "IdExpOdTc") + "' and " + _
                                     "OmYy= '" + GrdID.Value(GrdID.Row, "IdExpOdYy") + "' " + _
                                     "and OmChr= '" + GrdID.Value(GrdID.Row, "IdExpOdChr") + "' and " + _
                                      "OmNo= " + CStr(GrdID.Value(GrdID.Row, "IdExpOdNo")) + " and OmPrtKey='" + ctCurrPrtn + "' ")
        End If
    End If
  ElseIf UCase(ms_IdTcTyp) = UCase(ctCTBTcTyp) Then
    wCmCd = ADC("wIdCmCd")
  End If
  If ms_RtFrInvCust = "Y" Then
        wCmCd = ADC("wIdCmCd") '6.1-55
        ws_OmLkUpLabRt = ""
  End If
      
  With GrdIL
    Select Case ColName
    Case Is = UCase("iLQw")       '*** recalc on iLMCd, iLSCd
      If .Mode = fgmnorm Or .Mode = fgmCopy Then
        .Value(RowNum, "iLQw") = moCn.GetFldVal("Select PValue from Param where " + _
                                 "PTyp= 'LABSCD' and PMCd= '" + .Value(RowNum, "iLMCd") + "' " + _
                                 "and PSCd= '" + .Value(RowNum, "iLSCd") + "'")
      End If
    Case Is = UCase("iLQty")      '*** recalc on iLQW
      If .Mode = fgmnorm Or .Mode = fgmCopy Then
        If .Value(RowNum, "iLQty") = 0 Then .Value(RowNum, "iLQty") = 1
        
      End If
    Case Is = UCase("iLAVal")     '*** recalc of iLAVal on iLMCd,iLSCd,iLARt,iLQw,iLQty
                                  '*** In the Change Event Of IdALabAsWt & IdQty, recalculation of iLAVal takes place
                                  '*** In the change event of IdExpCmCd, recalc of iLAVal is not reqiured as IdExpCmCd is not allowed to be changed
      If .Mode = fgmnorm Or .Mode = fgmCopy Then
        If .Value(RowNum, "ilMCd") = ctBrkLMCd Then
          .Value(RowNum, "iLAVal") = .Value(RowNum, "iLARt") / 100 * GetDCRmVal("IrRmAVal")
        ElseIf .Value(RowNum, "ilMCd") = ctDHCLMCd Then
          .Value(RowNum, "iLAVal") = .Value(RowNum, "iLARt") / 100 * GetDCRmVal("IrRmAVal", , "D")
        ElseIf .Value(RowNum, "ilMCd") = ctCHCLMCd Then
          .Value(RowNum, "iLAVal") = .Value(RowNum, "iLARt") / 100 * GetDCRmVal("IrRmAVal", , "C")
        ElseIf .Value(RowNum, "ilMCd") = ctGHCLMCd Then
          Dim wLaborSubLogic As String
          wLaborSubLogic = GetLaborSubCdLabLogic(GrdIL.Value(RowNum, "ilMCd"), GrdIL.Value(RowNum, "ilSCd"))
          .Value(RowNum, "iLAVal") = .Value(RowNum, "iLARt") / 100 * GetDCRmVal(IIF(wLaborSubLogic = "P", "IrRmAWt", "IrRmAVal"), RowNum, ctGHCLMCd, wLaborSubLogic)
        ElseIf .Value(RowNum, "ilMCd") = "CHN" And ADC("wChnCustYn") = "Y" Then
           .Value(RowNum, "iLAVal") = .Value(RowNum, "iLARt") * GetDCRmVal("IrRmAWt", , "CHN")
        ElseIf .Value(RowNum, "ilMCd") = ctCDWLMCd And .Value(RowNum, "ilQw") = "W" Then                     ' Diamond certification charges
          .Value(RowNum, "iLAVal") = GetLabVal(wCmCd, .Value(RowNum, "ilMCd"), _
                                     .Value(RowNum, "ilSCd"), .Value(RowNum, "iLARt"), _
                                     .Value(RowNum, "iLQty"), _
                                     MWLib.Div(GetDCRmVal("IrRmAWt", RowNum, ctCDWLMCd, GetLaborSubCdLabLogic(GrdIL.Value(RowNum, "ilmcd"), GrdIL.Value(RowNum, "ilscd"))), GrdID.Value(GrdID.Row, "IdQty")), _
                                     .Value(RowNum, "iLcstQw"), "S", ms_CmCtg, ADC("wIdCmCurCd"), _
                                     IIF(ws_OmLkUpLabRt <> "", ws_OmLkUpLabRt, "")) * _
                                     GrdID.Value(GrdID.Row, "IdQty")
        ElseIf .Value(RowNum, "ilMCd") = ctCBRKLMCd And .Value(RowNum, "ilQw") = "W" Then                     ' Diamond certification charges
          .Value(RowNum, "iLAVal") = GetLabVal(wCmCd, .Value(RowNum, "ilMCd"), _
                                     .Value(RowNum, "ilSCd"), .Value(RowNum, "iLARt"), _
                                     .Value(RowNum, "iLQty"), _
                                     MWLib.Div(GetDCRmVal("iRrMAwT", , "C", ""), GrdID.Value(GrdID.Row, "IdQty")), _
                                     .Value(RowNum, "iLcstQw"), "S", ms_CmCtg, ADC("wIdCmCurCd"), _
                                     IIF(ws_OmLkUpLabRt <> "", ws_OmLkUpLabRt, "")) * _
                                     GrdID.Value(GrdID.Row, "IdQty")
                                     
        ElseIf .Value(RowNum, "ilMCd") = ctAddLMCd Then
        Else  '6.1-55 GrdID.Value(GrdID.Row, "IdExpCmCd") changed to wCmCd
          .Value(RowNum, "iLAVal") = GetLabVal(wCmCd, .Value(RowNum, "ilMCd"), _
                                     .Value(RowNum, "ilSCd"), .Value(RowNum, "iLARt"), _
                                     .Value(RowNum, "iLQty"), MWLib.Div(GrdID.Value(GrdID.Row, "IdALabAsWt"), GrdID.Value(GrdID.Row, "IdQty")), _
                                     .Value(RowNum, "ilCstQw"), "S", "C", ADC("wIdCmCurCd"), _
                                     IIF(ws_OmLkUpLabRt <> "", ws_OmLkUpLabRt, "")) * _
                                     GrdID.Value(GrdID.Row, "IdQty")
        
        End If
      End If
    Case Is = UCase("iLiRt")
      
      '*** (Jen 2.12) (06/01/06)
      Dim wiLQw As String
      wiLQw = ""
'      If ms_RtFrInvCust = "Y" Then
'        wCmCd = ADC("wIdCmCd") '6.1-55
'        ws_OmLkUpLabRt = ""
'      End If
      
      If (UCase(ms_IdTcTyp) = UCase(ctCTBTcTyp)) Then
        If (.Mode = fgmnorm Or .Mode = fgmCopy) And UCase(SourceName) = UCase("GrdIL") Then
          .Value(RowNum, "iLiRt") = GetLabRt(wCmCd, .Value(RowNum, "iLMCd"), _
                                    .Value(RowNum, "iLSCd"), "S", _
                                    MWLib.Div(GrdID.Value(GrdID.Row, "IdILabAsWt"), _
                                    GrdID.Value(GrdID.Row, "IdQty")), ms_CmCtg, wiLQw, , , ADC("wIdCmCurCd"))
          If wiLQw <> "" Then .Value(RowNum, "iLQw") = wiLQw
          If .Value(RowNum, "iLQty") = 0 Then .Value(RowNum, "iLQty") = 1
        End If
      Else
        If .Mode = fgmnorm Or .Mode = fgmCopy Then
          If .Value(RowNum, "ilMCd") = ctBrkLMCd And .Value(RowNum, "iLSCd") = ctBrkLSCd Then
            If .Mode = fgmnorm Then
              .Value(RowNum, "iLiRt") = GetLabRt(wCmCd, .Value(RowNum, "iLMCd"), _
                                        .Value(RowNum, "iLSCd"), "S", _
                                        MWLib.Div(GrdID.Value(GrdID.Row, "IdILabAsWt"), _
                                        GrdID.Value(GrdID.Row, "IdQty")), ms_CmCtg, wiLQw, , , ADC("wIdCmCurCd"), _
                                        IIF(ws_OmLkUpLabRt <> "", ws_OmLkUpLabRt, ""))
            End If
            .Value(RowNum, "iLQw") = "Q"
            .Value(RowNum, "iLQty") = 1
            .Value(RowNum, "iLZRt") = .Value(RowNum, "iLiRt") '*** (Jen 2.12) 04/02/06
          ElseIf .Mode = fgmnorm And .Value(RowNum, "iLMCd") = ctAddLMCd And .Value(RowNum, "iLSCd") = ctAddLSCd Then
            .Value(RowNum, "iLiRt") = 0
            .Value(RowNum, "iLQw") = "Q"
            .Value(RowNum, "iLQty") = 1
          ElseIf .Value(RowNum, "ilMCd") = ctDHCLMCd Then       ''''And .Value(RowNum, "iLSCd") = ctDCHCLSCd
            If .Mode = fgmnorm Then
              
              .Value(RowNum, "iLiRt") = GetLabRt(wCmCd, .Value(RowNum, "iLMCd"), _
                                        .Value(RowNum, "iLSCd"), "S", _
                                        MWLib.Div(GrdID.Value(GrdID.Row, "wIdIDiaWt"), _
                                        GrdID.Value(GrdID.Row, "IdQty")), ms_CmCtg, wiLQw, , , ADC("wIdCmCurCd"), _
                                        IIF(ws_OmLkUpLabRt <> "", ws_OmLkUpLabRt, ""))
            End If
            .Value(RowNum, "iLQw") = "Q"
            .Value(RowNum, "iLQty") = 1
            .Value(RowNum, "iLZRt") = .Value(RowNum, "iLiRt")
          ElseIf .Value(RowNum, "Ilmcd") = ctGHCLMCd Then
            If .Mode = fgmnorm Then
              .Value(RowNum, "iLiRt") = GetLabRt(wCmCd, .Value(RowNum, "iLMCd"), _
                                          .Value(RowNum, "iLSCd"), "S", _
                                          MWLib.Div(GetDCRmVal("IrRmiWt", RowNum, ctGHCLMCd, GetLaborSubCdLabLogic(GrdIL.Value(RowNum, "ilmcd"), GrdIL.Value(RowNum, "ilscd"))), GrdID.Value(GrdID.Row, "IdQty")), _
                                           ms_CmCtg, wiLQw, , , ADC("wIdCmCurCd"), _
                                          IIF(ws_OmLkUpLabRt <> "", ws_OmLkUpLabRt, ""))

            End If
            If wiLQw <> "" And .Value(RowNum, "iLQw") = "" Then .Value(RowNum, "iLQw") = wiLQw
            If .Value(RowNum, "iLQty") = 0 Then .Value(RowNum, "iLQty") = 1
        .Value(RowNum, "iLZRt") = .Value(RowNum, "iLiRt")
          
          ElseIf .Value(RowNum, "ilMCd") = ctCHCLMCd Then       '''' And .Value(RowNum, "iLSCd") = ctDCHCLSCd
            If .Mode = fgmnorm Then
              
              .Value(RowNum, "iLiRt") = GetLabRt(wCmCd, .Value(RowNum, "iLMCd"), _
                                        .Value(RowNum, "iLSCd"), "S", _
                                        MWLib.Div(GrdID.Value(GrdID.Row, "wIdICSWt"), _
                                        GrdID.Value(GrdID.Row, "IdQty")), ms_CmCtg, wiLQw, , , ADC("wIdCmCurCd"), _
                                        IIF(ws_OmLkUpLabRt <> "", ws_OmLkUpLabRt, ""))
            End If
            .Value(RowNum, "iLQw") = "Q"
            .Value(RowNum, "iLQty") = 1
            .Value(RowNum, "iLZRt") = .Value(RowNum, "iLiRt")
          ElseIf .Value(RowNum, "ilMCd") = ctCDWLMCd Then
            If .Mode = fgmnorm Then
              
              .Value(RowNum, "iLiRt") = GetLabRt(wCmCd, .Value(RowNum, "iLMCd"), _
                                        .Value(RowNum, "iLSCd"), "S", _
                                        MWLib.Div(GetDCRmVal("IrRmIWt", RowNum, ctCDWLMCd, GetLaborSubCdLabLogic(GrdIL.Value(RowNum, "ilMCd"), GrdIL.Value(RowNum, "ilSCd"))), GrdID.Value(GrdID.Row, "IdQty")), _
                                        ms_CmCtg, wiLQw, , , ADC("wIdCmCurCd"), _
                                        IIF(ws_OmLkUpLabRt <> "", ws_OmLkUpLabRt, ""))
            End If
            If wiLQw <> "" And .Value(RowNum, "iLQw") = "" Then .Value(RowNum, "iLQw") = wiLQw
            If .Value(RowNum, "iLQty") = 0 Then .Value(RowNum, "iLQty") = 1
            .Value(RowNum, "iLZRt") = .Value(RowNum, "iLiRt")
          ElseIf .Value(RowNum, "ilMCd") = ctCBRKLMCd Then
            If .Mode = fgmnorm Then
              
              .Value(RowNum, "iLiRt") = GetLabRt(wCmCd, .Value(RowNum, "iLMCd"), _
                                        .Value(RowNum, "iLSCd"), "S", _
                                        MWLib.Div(GetDCRmVal("iRrMiwT", , "C", ""), GrdID.Value(GrdID.Row, "IdQty")), _
                                        ms_CmCtg, wiLQw, , , ADC("wIdCmCurCd"), _
                                        IIF(ws_OmLkUpLabRt <> "", ws_OmLkUpLabRt, ""))
            End If
            If wiLQw <> "" And .Value(RowNum, "iLQw") = "" Then .Value(RowNum, "iLQw") = wiLQw
            If .Value(RowNum, "iLQty") = 0 Then .Value(RowNum, "iLQty") = 1
            .Value(RowNum, "iLZRt") = .Value(RowNum, "iLiRt")
          
            
          
          End If
        End If
      End If
    Case Is = UCase("iLiVal")     '*** recalc of iLiVal on iLMCd,iLSCd,iLiRt,iLQw,iLQty
                                  '*** In the Change Event Of IdILabAsWt & IdQty, recalculation of iLiVal takes place
                                  '*** In the change event of IdExpCmCd, recalc of iLiVal is not reqiured as IdExpCmCd is not allowed to be changed
      If .Mode = fgmnorm Or .Mode = fgmCopy Then
        If .Value(RowNum, "ilMCd") = ctAddLMCd Then
          .Value(RowNum, "iLiVal") = 0
        ElseIf .Value(RowNum, "ilMCd") = ctBrkLMCd Then
          .Value(RowNum, "iLiVal") = .Value(RowNum, "iLiRt") / 100 * GetDCRmVal("IrRmiVal")
        ElseIf .Value(RowNum, "ilMCd") = ctDHCLMCd Then
          .Value(RowNum, "iLiVal") = .Value(RowNum, "iLiRt") / 100 * GetDCRmVal("IrRmiVal", , "D")
        ElseIf .Value(RowNum, "ilMCd") = ctCHCLMCd Then
          .Value(RowNum, "iLiVal") = .Value(RowNum, "iLiRt") / 100 * GetDCRmVal("IrRmiVal", , "C")
        ElseIf .Value(RowNum, "ilMCd") = ctGHCLMCd Then
       '   Dim wLaborSubLogic As String
          wLaborSubLogic = GetLaborSubCdLabLogic(GrdIL.Value(RowNum, "ilMCd"), GrdIL.Value(RowNum, "ilSCd"))
          .Value(RowNum, "iLiVal") = .Value(RowNum, "iLiRt") / 100 * GetDCRmVal(IIF(wLaborSubLogic = "P", "IrRmiWt", "IrRmiVal"), RowNum, ctGHCLMCd, wLaborSubLogic)
          
        '**********Bhavna added chain labour value
        ElseIf .Value(RowNum, "ilMCd") = "CHN" And ADC("wChnCustYn") = "Y" Then
           .Value(RowNum, "iLiVal") = .Value(RowNum, "iLiRt") * GetDCRmVal("IrRmIWt", , "CHN")
        '**********Bhavna added chain labour value
        ElseIf .Value(RowNum, "ilMCd") = ctCDWLMCd And .Value(RowNum, "ilQw") = "W" Then
            .Value(RowNum, "iLiVal") = GetLabVal(wCmCd, .Value(RowNum, "ilMCd"), _
                                     .Value(RowNum, "ilSCd"), .Value(RowNum, "iLiRt"), _
                                     .Value(RowNum, "iLQty"), MWLib.Div(GetDCRmVal("IrRmIWt", RowNum, ctCDWLMCd, GetLaborSubCdLabLogic(GrdIL.Value(RowNum, "ilMCd"), .Value(RowNum, "ilSCd"))), GrdID.Value(GrdID.Row, "IdQty")), _
                                     .Value(RowNum, "iLQw"), "S", ms_CmCtg, ADC("wIdCmCurCd"), _
                                     IIF(ws_OmLkUpLabRt <> "", ws_OmLkUpLabRt, "")) * _
                                     GrdID.Value(GrdID.Row, "IdQty")
        ElseIf .Value(RowNum, "ilMCd") = ctCBRKLMCd And .Value(RowNum, "ilQw") = "W" Then
            .Value(RowNum, "iLiVal") = GetLabVal(wCmCd, .Value(RowNum, "ilMCd"), _
                                     .Value(RowNum, "ilSCd"), .Value(RowNum, "iLiRt"), _
                                     .Value(RowNum, "iLQty"), MWLib.Div(GetDCRmVal("iRrMiwT", , "C", ""), GrdID.Value(GrdID.Row, "IdQty")), _
                                     .Value(RowNum, "iLQw"), "S", ms_CmCtg, ADC("wIdCmCurCd"), _
                                     IIF(ws_OmLkUpLabRt <> "", ws_OmLkUpLabRt, "")) * _
                                     GrdID.Value(GrdID.Row, "IdQty")
          
        Else
          .Value(RowNum, "iLiVal") = GetLabVal(wCmCd, .Value(RowNum, "ilMCd"), _
                                     .Value(RowNum, "ilSCd"), .Value(RowNum, "iLiRt"), _
                                     .Value(RowNum, "iLQty"), MWLib.Div(GrdID.Value(GrdID.Row, "IdILabAsWt"), GrdID.Value(GrdID.Row, "IdQty")), _
                                     .Value(RowNum, "iLQw"), "S", ms_CmCtg, ADC("wIdCmCurCd"), _
                                     IIF(ws_OmLkUpLabRt <> "", ws_OmLkUpLabRt, "")) * _
                                     GrdID.Value(GrdID.Row, "IdQty")
                                     'IIf(Trim(.Value(RowNum, "iLQw")) = "Q", GrdID.Value(GrdID.Row, "IdQty"), GrdID.Value(GrdID.Row, "IdILabAsWt"))
        End If
      End If
    Case Is = UCase("iLZRt")
      If UCase(ms_IdTcTyp) = UCase(ctCTBTcTyp) Then
        If (.Mode = fgmnorm Or .Mode = fgmCopy) And UCase(SourceName) = UCase("GrdIL") Then
          .Value(RowNum, "iLZRt") = MWLib.Div(.Value(RowNum, "iLiRt"), md_InCnvRt)
        End If
      Else
        If UCase(SourceName) = UCase("GrdIL") Then
          If .Mode = fgmnorm Or .Mode = fgmCopy Then
            If .Mode = fgmnorm And .Value(RowNum, "iLMCd") = ctAddLMCd And .Value(RowNum, "iLSCd") = ctAddLSCd Then
              .Value(RowNum, "iLZRt") = GetZAddVal(GrdID.Row)
            ElseIf .Value(RowNum, "iLMCd") = ctBrkLMCd And .Value(RowNum, "iLSCd") = ctBrkLSCd Then
              If .Mode = fgmnorm Then
                .Value(RowNum, "iLZRt") = GetLabRt(wCmCd, .Value(RowNum, "iLMCd"), _
                                          .Value(RowNum, "iLSCd"), "S", _
                                          MWLib.Div(GrdID.Value(GrdID.Row, "IdZLabAsWt"), _
                                          GrdID.Value(GrdID.Row, "IdQty")), ms_CmCtg, wiLQw, , , ADC("wIdCmCurCd"), _
                                          IIF(ws_OmLkUpLabRt <> "", ws_OmLkUpLabRt, ""))
              End If
            ElseIf .Value(RowNum, "iLMCd") = ctDHCLMCd Then       ''''' And .Value(RowNum, "iLSCd") = ctDCHCLSCd
              If .Mode = fgmnorm Then
                
                .Value(RowNum, "iLZRt") = GetLabRt(wCmCd, .Value(RowNum, "iLMCd"), _
                                          .Value(RowNum, "iLSCd"), "S", _
                                          MWLib.Div(GrdID.Value(GrdID.Row, "wIdZDiaWt"), _
                                          GrdID.Value(GrdID.Row, "IdQty")), ms_CmCtg, wiLQw, , , ADC("wIdCmCurCd"), _
                                          IIF(ws_OmLkUpLabRt <> "", ws_OmLkUpLabRt, ""))
              End If
            ElseIf .Value(RowNum, "iLMCd") = ctGHCLMCd Then
'            Then       ''''' And .Value(RowNum, "iLSCd") = ctDCHCLSCd
              If .Mode = fgmnorm Then
                .Value(RowNum, "iLZRt") = GetLabRt(wCmCd, .Value(RowNum, "iLMCd"), _
                                          .Value(RowNum, "iLSCd"), "S", _
                                          MWLib.Div(GetDCRmVal("IrRmZWt", RowNum, ctGHCLMCd, GetLaborSubCdLabLogic(GrdIL.Value(RowNum, "ilmcd"), GrdIL.Value(RowNum, "ilscd"))), GrdID.Value(GrdID.Row, "IdQty")), _
                                           ms_CmCtg, wiLQw, , , ADC("wIdCmCurCd"), _
                                          IIF(ws_OmLkUpLabRt <> "", ws_OmLkUpLabRt, ""))
              End If
              If wiLQw <> "" And .Value(RowNum, "iLQw") = "" Then .Value(RowNum, "iLQw") = wiLQw
              If .Value(RowNum, "iLQty") = 0 Then .Value(RowNum, "iLQty") = 1
            ElseIf .Value(RowNum, "iLMCd") = ctCHCLMCd Then       '''' And .Value(RowNum, "iLSCd") = ctDCHCLSCd
              If .Mode = fgmnorm Then
                .Value(RowNum, "iLZRt") = GetLabRt(wCmCd, .Value(RowNum, "iLMCd"), _
                                          .Value(RowNum, "iLSCd"), "S", _
                                          MWLib.Div(GrdID.Value(GrdID.Row, "wIdZCSWt"), _
                                          GrdID.Value(GrdID.Row, "IdQty")), ms_CmCtg, wiLQw, , , ADC("wIdCmCurCd"), _
                                          IIF(ws_OmLkUpLabRt <> "", ws_OmLkUpLabRt, ""))
              End If
            ElseIf .Value(RowNum, "iLMCd") = ctCBRKLMCd Then
              If .Mode = fgmnorm Then
                .Value(RowNum, "iLZRt") = GetLabRt(wCmCd, .Value(RowNum, "iLMCd"), _
                                          .Value(RowNum, "iLSCd"), "S", _
                                          MWLib.Div(GetDCRmVal("iRrMzwT", , "C", ""), GrdID.Value(GrdID.Row, "IdQty")), _
                                          ms_CmCtg, wiLQw, , , ADC("wIdCmCurCd"), _
                                          IIF(ws_OmLkUpLabRt <> "", ws_OmLkUpLabRt, ""))
              End If
            ElseIf .Value(RowNum, "iLMCd") = ctCDWLMCd Then
              If .Mode = fgmnorm Then
                .Value(RowNum, "iLZRt") = GetLabRt(wCmCd, .Value(RowNum, "iLMCd"), _
                                          .Value(RowNum, "iLSCd"), "S", _
                                          MWLib.Div(GetDCRmVal("IrRmZWt", RowNum, ctCDWLMCd, GetLaborSubCdLabLogic(GrdIL.Value(RowNum, "ilMCd"), .Value(RowNum, "ilSCd"))), GrdID.Value(GrdID.Row, "IdQty")), _
                                          ms_CmCtg, wiLQw, , , ADC("wIdCmCurCd"), _
                                          IIF(ws_OmLkUpLabRt <> "", ws_OmLkUpLabRt, ""))
              End If
            End If
            
            End If
        ElseIf UCase(SourceName) = UCase("GrdID") Then
          If (.Mode = fgmnorm Or .Mode = fgmdel Or .Mode = fgmCopy) Then
            For wi_i = 1 To (.Rows - 1)
              If GrdID.Mode = fgmnorm And (.Mode = fgmnorm Or .Mode = fgmdel) And _
                .Value(wi_i, "iLMCd") = ctAddLMCd And .Value(wi_i, "iLSCd") = ctAddLSCd Then
                If Not .IsDel(wi_i) Then
                  .Value(wi_i, "iLZRt") = GetZAddVal(RowNum)
                  .SaveRec (wi_i)
                End If
              End If
            Next wi_i
          End If
        End If
      End If
    Case Is = UCase("iLZVal")     '*** recalc of iLiVal on iLMCd,iLSCd,iLZRt,iLQw,iLQty
                                  '*** In the Change Event Of IdZLabAsWt & IdQty, recalculation of iLiVal takes place
                                  '*** In the change event of IdExpCmCd, recalc of iLZVal is not reqiured as IdExpCmCd is not allowed to be changed
      
      If .Mode = fgmnorm Or .Mode = fgmCopy Or .Mode = fgmdel Then
        If .Value(RowNum, "ilMCd") = ctAddLMCd Then
          If (.Mode = fgmnorm Or .Mode = fgmdel) And Not .IsDel(RowNum) Then
            Dim wd_LabVal As Double
            wd_LabVal = .Value(RowNum, "iLQty") * .Value(RowNum, "iLZRt")
            If wd_LabVal <> .Value(RowNum, "iLZVal") Then .Value(RowNum, "iLZVal") = wd_LabVal
          End If
        ElseIf .Value(RowNum, "ilMCd") = ctBrkLMCd Then
          If .Mode = fgmnorm Or .Mode = fgmCopy Then
            .Value(RowNum, "iLZVal") = .Value(RowNum, "iLZRt") / 100 * GetDCRmVal("IrRmZVal")
          End If
        ElseIf .Value(RowNum, "ilMCd") = ctDHCLMCd Then
          If .Mode = fgmnorm Or .Mode = fgmCopy Then
            .Value(RowNum, "iLZVal") = .Value(RowNum, "iLZRt") / 100 * GetDCRmVal("IrRmZVal", , "D")
          End If
        ElseIf .Value(RowNum, "ilMCd") = ctCHCLMCd Then
          If .Mode = fgmnorm Or .Mode = fgmCopy Then
            .Value(RowNum, "iLZVal") = .Value(RowNum, "iLZRt") / 100 * GetDCRmVal("IrRmZVal", , "C")
          End If
        '**************Bhavna added chain labour value
        ElseIf .Value(RowNum, "ilMCd") = "CHN" And ADC("wChnCustYn") = "Y" Then
          If .Mode = fgmnorm Or .Mode = fgmCopy Then
            .Value(RowNum, "iLZVal") = .Value(RowNum, "iLZRt") * GetDCRmVal("IrRmZWt", , "CHN")
          End If
          '**************Bhavna added chain labour value
        ElseIf .Value(RowNum, "ilMCd") = ctCDWLMCd And .Value(RowNum, "ilQw") = "W" Then
          If .Mode = fgmnorm Or .Mode = fgmCopy Then
         .Value(RowNum, "iLzVal") = GetLabVal(wCmCd, .Value(RowNum, "ilMCd"), _
                                     .Value(RowNum, "ilSCd"), .Value(RowNum, "iLzRt"), _
                                     .Value(RowNum, "iLQty"), MWLib.Div(GetDCRmVal("IrRmZWt", RowNum, ctCDWLMCd, GetLaborSubCdLabLogic(GrdIL.Value(RowNum, "ilMCd"), GrdIL.Value(RowNum, "ilSCd"))), GrdID.Value(GrdID.Row, "IdQty")), _
                                     .Value(RowNum, "iLQw"), "S", ms_CmCtg, ADC("wIdCmCurCd"), _
                                     IIF(ws_OmLkUpLabRt <> "", ws_OmLkUpLabRt, "")) * _
                                     GrdID.Value(GrdID.Row, "IdQty")
          End If
        ElseIf .Value(RowNum, "ilMCd") = ctCBRKLMCd And .Value(RowNum, "ilQw") = "W" Then
          If .Mode = fgmnorm Or .Mode = fgmCopy Then
         .Value(RowNum, "iLzVal") = GetLabVal(wCmCd, .Value(RowNum, "ilMCd"), _
                                     .Value(RowNum, "ilSCd"), .Value(RowNum, "iLzRt"), _
                                     .Value(RowNum, "iLQty"), MWLib.Div(GetDCRmVal("iRrMzwT", , "C", ""), GrdID.Value(GrdID.Row, "IdQty")), _
                                     .Value(RowNum, "iLQw"), "S", ms_CmCtg, ADC("wIdCmCurCd"), _
                                     IIF(ws_OmLkUpLabRt <> "", ws_OmLkUpLabRt, "")) * _
                                     GrdID.Value(GrdID.Row, "IdQty")
          End If
          
        Else
          If .Mode = fgmnorm Or .Mode = fgmCopy Then
            If UCase(ms_IdTcTyp) = UCase(ctInTcTyp) Then
              .Value(RowNum, "iLZVal") = GetLabVal(wCmCd, .Value(RowNum, "ilMCd"), _
                                         .Value(RowNum, "ilSCd"), .Value(RowNum, "iLZRt"), _
                                         .Value(RowNum, "iLQty"), MWLib.Div(GrdID.Value(GrdID.Row, "IdZLabAsWt"), GrdID.Value(GrdID.Row, "IdQty")), _
                                         .Value(RowNum, "iLQw"), "S", ms_CmCtg, ADC("wIdCmCurCd"), _
                                         IIF(ws_OmLkUpLabRt <> "", ws_OmLkUpLabRt, "")) * _
                                         GrdID.Value(GrdID.Row, "IdQty")
                                         'IIf(Trim(.Value(RowNum, "iLQw")) = "Q", GrdID.Value(GrdID.Row, "IdQty"), GrdID.Value(GrdID.Row, "IdZLabAsWt"))
            ElseIf UCase(ms_IdTcTyp) = UCase(ctCTBTcTyp) Then
              .Value(RowNum, "iLZVal") = .Value(RowNum, "iLZRt") * IIF(Trim(.Value(RowNum, "iLQw")) = "Q", .Value(RowNum, "iLQty") * GrdID.Value(GrdID.Row, "IdQty"), GrdID.Value(GrdID.Row, "IdZLabAsWt"))
            End If
          End If
        End If
      End If
    End Select
  End With
End Sub
Private Sub GrdIL_ColChange(ByVal RowNum As Integer, ByVal ColName As String, ByVal NewValue As Variant, Cancel As Boolean, ErrMsg As String)
  '*** iLMCd          valid PMCd('LABMCD') or Blank
  '*** ilSCd          valid PSCd('LABSCD', ilMCd)  or Blank
  
  With GrdIL
    Select Case ColName
    Case Is = UCase("iLMCd")
      '*** (Jen 3.01)
      If .Mode = fgmnorm Then
      '*** (Jen 3.01)
      
        Cancel = Not moCn.RecSeek("Select PMCd From Param Where " + _
                             "PTyp= 'LABMCD' And PMCd = '" + NewValue + "'")
        ErrMsg = "Invalid Labour Main Code": Exit Sub
      End If
    Case Is = UCase("iLSCd")
      '*** (Jen 3.01)
      If .Mode = fgmnorm Then
      '*** (Jen 3.01)
    
        Cancel = (NewValue <> "") And (Not moCn.RecSeek("Select PSCd From Param Where " + _
                  "PTyp= 'LABSCD' And PMCd='" + .Value(RowNum, "ilMCd") + "' " + _
                  "And PSCd = '" + NewValue + "'"))
        ErrMsg = "Invalid Labour Sub Code": Exit Sub
      End If
    End Select
  End With
End Sub
Private Sub GrdIL_InitKey(ByVal RowNum As Integer)
  '*** Initialise Key Fields (i.e. IlTc, IlYy, IlChr, IlNo, IlSr)
  
  With GrdIL
    .Store "IlCoCd", gs_CoCd
    .Store "IlTc", ADC("wIdTc")
    .Store "IlYy", ADC("wIdYy")
    .Store "IlChr", ADC("wIdChr")
    .Store "IlNo", ADC("wIdNo")
    .Store "IlSr", GrdID.Value(GrdID.Row, "IdSr")
    If mf_inspecialCopy Then
    .Store "Cpy", "Y"
    Else
    .Store "Cpy", ""
    End If
    
    '****** Sachin 3.02 Id fields in Joins
    If .IsNew(RowNum) Then
        .Store "IlIdIdNo", moCn.GetFldVal("Select IdIdNo From InvDsg" + _
                                          " Where " + IIF(gs_Partition = ctCurrPrtn, " IdPrtKey='" + ctCurrPrtn + "' And ", "") + " IdCoCd='" + gs_CoCd + "' and IdTc='" + ADC("wIdTc") + "' and IdYy='" + ADC("wIdYy") + "' and IdChr='" + ADC("wIdChr") + "' and IdNo=" + CStr(ADC("wIdNo")) + " and IdSr=" + CStr(GrdID.Value(GrdID.Row, "IdSr")))
    End If
    '****** Sachin 3.02 Id fields in Joins
    
  End With
End Sub

Private Sub GrdIB_When()
    '*** Jen 2.12 Path 2 (04/05/06)
    If mf_DtLocked = False And mInCls = "N" Then
      GrdIB.AllowDelete = True
      GrdIB.AllowAdd = True
    End If
    
    If GrdIR.Rows - 1 > 0 Then GrdIB.AllowDelete = False: GrdIB.AllowAdd = False
    If GrdIL.Rows - 1 > 0 Then GrdIB.AllowDelete = False: GrdIB.AllowAdd = False
    '*** Jen 2.12 Path 2 (04/05/06)
    
    '****** Sachin 3.3.0 (23-06-08) - [42. Make Invoice By Bag] ******
    If mb_InByBag = True And GrdIB.Rows - 1 = 1 Then
        GrdIB.AllowAdd = False
    Else
        GrdIB.AllowAdd = True
    End If
    '****** Sachin 3.3.0 (23-06-08) - [42. Make Invoice By Bag] ******
End Sub

Private Sub GrdIB_RowWhen(ByVal RowNum As Integer)
  '*** If New Row then IfSrNo= max(IfSrNo) for the Invoice + 1
  '*** Note that IfSrNo is unique for the whole invoice
  With GrdIB
    If .Value(RowNum, "IfSrNo") = 0 Then _
       .Value(RowNum, "IfSrNo") = moCn.GetFldVal("Select max(IfSrNo) from InvFgd " + _
                "where IfCoCd= '" + gs_CoCd + "' and  IfTc= '" + ADC("wIdTc") + "' and  " + _
                "IfYy= '" + ADC("wIdYy") + "' and IfChr= '" + ADC("wIdChr") + "' and " + _
                "IfNo= " + CStr(ADC("wIdNo")) + IIF(ADC("wIdPrtKey") = ctCurrPrtn, " and IfPrtKey='" + ctCurrPrtn + "' ", "")) + 1
     If .IsNew(RowNum) Then .Value(RowNum, "IfPrtKey") = ctCurrPrtn '****** Sachin 3.02.0
    '##### .Value(RowNum, "IfSrNo") = moCn.AutoGen("InvFgd", "IfSrNo", adc("wIdTc") + adc("wIdYy") + adc("wIdChr") + CStr(adc("wIdNo")))
  End With
End Sub
Private Sub GrdIB_ColWhen(ByVal RowNum As Integer, ByVal ColName As String, Cancel As Boolean, ErrMsg As String, ByVal Action As MwfCtl.en_fgAction)
  '*** Do not allow the user to edit any field if InCls = 'Y' in InvHd Table
  '*** Do not allow the user to edit any column if record already exists
  '*** Give help for IfFdTc, IfFdYy, IfFdChr, IfFdNo, IfFdSr, IfFdBYy, IfFdBChr, IfFdBNo
  '*** If IfFdBYy<>'' or IfFdBChr<> '' or IfFdBNo<> 0 then do not allow user to edit
      'IfFdTc, IfFdYy, IfFdChr, IfFdNo, IfFdSr
  '*** Do not allow the user to edit IfQty
  
    '****** Sachin 3.02 - Locking the Previous Partition Records
    If ADC("WIDPRTKEY") <> ctCurrPrtn Then Cancel = True: ErrMsg = "Cannot Edit Previous Partition Records": Exit Sub
  
    If mf_DtLocked = True Then Cancel = True: ErrMsg = "Cannot Edit Locked Entry": Exit Sub
    If mInCls = "Y" Then Cancel = True: ErrMsg = "Cannot Edit As Invoice Is Closed": Exit Sub
  
  
  '*** Jen 2.12 Path 2 (04/05/06)
  Dim ws_DelMsg As String
  If GrdIR.Rows - 1 > 0 Or GrdIL.Rows - 1 > 0 Then
    ws_DelMsg = "/ Cannot Delete As Rm Or Lab Rec Exists "
  Else
    ws_DelMsg = " "
  End If
  '*** Jen 2.12 Path 2 (04/05/06)
  '****** Sachin 3.02 - Current Partition
  If moCn.RecSeek("Select IfTc From InvFgd where IfCoCd= '" + gs_CoCd + "' and " + _
          "IfTc= '" + ADC("wIdTc") + "' and " + _
          "IfYy= '" + ADC("wIdYy") + "' and " + _
          "IfChr= '" + ADC("wIdChr") + "' and " + _
          "IfNo= " + CStr(ADC("wIdNo")) + " and " + _
          "IfSr= " + CStr(GrdID.Value(GrdID.Row, "IdSr")) + " and " + _
          "IfSrNo= " + CStr(GrdIB.Value(RowNum, "IfSrNo")) + IIF(ADC("wIdPrtKey") = ctCurrPrtn, " And  IfPrtKey='" + ctCurrPrtn + "' ", "")) Then _
          Cancel = True: ErrMsg = "Cannot Edit This Value" + ws_DelMsg: Exit Sub
  
  Select Case ColName
  Case Is = UCase("IfFdBYy")
    Call HlpList.PMCd("YY")
  Case Is = UCase("IfFdBChr")
' **** Zubin 211 **** '
'''''    Call HlpList.PMCd("BCHR")
    Call HlpList.PMCd("BCHR", "'P'")
' **** Zubin 211 **** '
  Case Is = UCase("IfFdBNo")
    Call HlpList.BNo(gs_CoCd, GrdIB.Value(RowNum, "IfFdBYy"), GrdIB.Value(RowNum, "IfFdBChr"), "Y", ctFgLoc)
  Case Is = UCase("IfFdTc")
    If GrdIB.Value(RowNum, "IfFdBYy") <> "" Or GrdIB.Value(RowNum, "IfFdBChr") <> "" Or _
       GrdIB.Value(RowNum, "IfFdBNo") <> 0 Then _
       Cancel = True: ErrMsg = "Cannot Edit Finished Goods Tc": Exit Sub
    Call HlpList.PMCd("TC", "'FB','FR'")
  Case Is = UCase("IfFdYy")
    If GrdIB.Value(RowNum, "IfFdBYy") <> "" Or GrdIB.Value(RowNum, "IfFdBChr") <> "" Or _
       GrdIB.Value(RowNum, "IfFdBNo") <> 0 Then _
       Cancel = True: ErrMsg = "Cannot Edit Finished Goods Year": Exit Sub
    Call HlpList.PMCd("YY")
  Case Is = UCase("IfFdChr")
    If GrdIB.Value(RowNum, "IfFdBYy") <> "" Or GrdIB.Value(RowNum, "IfFdBChr") <> "" Or _
       GrdIB.Value(RowNum, "IfFdBNo") <> 0 Then _
       Cancel = True: ErrMsg = "Cannot Edit Finished Goods Character": Exit Sub
    Call HlpList.vPSCd(gs_CoCd, "CHR", GrdIB.Value(RowNum, "IfFdTc"))
  Case Is = UCase("IfFdNo")
    If GrdIB.Value(RowNum, "IfFdBYy") <> "" Or GrdIB.Value(RowNum, "IfFdBChr") <> "" Or _
       GrdIB.Value(RowNum, "IfFdBNo") <> 0 Then _
       Cancel = True: ErrMsg = "Cannot Edit Finished Goods Number": Exit Sub
    Call HlpList.FgNo(gs_CoCd, GrdIB.Value(RowNum, "IfFdTc"), GrdIB.Value(RowNum, "IfFdYy"), GrdIB.Value(RowNum, "IfFdChr"))
  Case Is = UCase("IfFdSr")
    If GrdIB.Value(RowNum, "IfFdBYy") <> "" Or GrdIB.Value(RowNum, "IfFdBChr") <> "" Or _
       GrdIB.Value(RowNum, "IfFdBNo") <> 0 Then _
       Cancel = True: ErrMsg = "Cannot Edit Finished Goods Serial": Exit Sub
    Call HlpList.FdSr(gs_CoCd, GrdIB.Value(RowNum, "IfFdTc"), GrdIB.Value(RowNum, "IfFdYy"), GrdIB.Value(RowNum, "IfFdChr"), GrdIB.Value(RowNum, "IfFdNo"))
  Case Is = UCase("IfQty")
    Cancel = True: ErrMsg = "Cannot Edit This Value": Exit Sub
  End Select
End Sub
Private Sub GrdIB_ReCalc(ByVal RowNum As Integer, ByVal ColName As String, ByVal OldVal As Variant, ByVal SourceName As String, ByVal SourceColName As String)
  '*** IfFdBYy on IfFdTc,IfFdYy,IfFdChr,IfFdNo,IfFdSr                   Normal
      '   if if IfFdSr<> 0 and exists Fgd(FdTc= IfFdTc, FdYy= IfFdYy, FdChr= IfFdChr, FdNo= IfFdNo, FdSr= IfFdSr)
      '   then IfFdBYy= FdBYy from Fgd(FdTc= IfFdTc, FdYy= IfFdYy, FdChr= IfFdChr, FdNo= IfFdNo, FdSr= IfFdSr)
  '*** IfFdBChr on IfFdTc,IfFdYy,IfFdChr,IfFdNo,IfFdSr                  Normal
      '   if if IfFdSr<> 0 and exists Fgd(FdTc= IfFdTc, FdYy= IfFdYy, FdChr= IfFdChr, FdNo= IfFdNo, FdSr= IfFdSr)
      '   then IfFdBChr= FdBChr from Fgd(FdTc= IfFdTc, FdYy= IfFdYy, FdChr= IfFdChr, FdNo= IfFdNo, FdSr= IfFdSr)
  '*** IfFdBNo on IfFdBYy,IfFdBChr,IfFdTc,IfFdYy,IfFdChr,IfFdNo,IfFdSr  Normal
      'If SourceColName= IfFdBYy or IfFdBChr then
      '   IfFdBNo = 0
      'If SourceColName= IfFdTc or IfFdYy or IfFdChr or IfFdNo or IfFdSr then
      '   if IfFdSr<> 0 and exists Fgd(FdTc= IfFdTc, FdYy= IfFdYy, FdChr= IfFdChr, FdNo= IfFdNo, FdSr= IfFdSr)
      '   then IfFdBNo= FdBNo from Fgd(FdTc= IfFdTc, FdYy= IfFdYy, FdChr= IfFdChr, FdNo= IfFdNo, FdSr= IfFdSr)
  '*** IfFdTc on IfFdBYy,IfFdBChr,IfFdBNo                               Normal
      'If IfFdBNo<> 0 then FdTc from Fgd(FdBYy= IfFdBYy, FdBChr= IfFdBChr, FdBNo= IfFdBNo and TcType= 'FB')
  '*** IfFdYy on IfFdBYy,IfFdBChr,IfFdBNo                               Normal
      'If IfFdBNo<> 0 then FdYy from Fgd(FdBYy= IfFdBYy, FdBChr= IfFdBChr, FdBNo= IfFdBNo and TcType= 'FB')
  '*** IfFdChr on IfFdBYy,IfFdBChr,IfFdBNo                              Normal
      'If IfFdBNo<> 0 then FdChr from Fgd(FdBYy= IfFdBYy, FdBChr= IfFdBChr, FdBNo= IfFdBNo and TcType= 'FB')
  '*** IfFdNo on IfFdBYy,IfFdBChr,IfFdBNo,IfFdTc,IfFdYy,IfFdChr         Normal
      'If SourceColName= IfFdBYy or IfFdBChr or IfFdBNo then
      '   if IfFdBNo<> 0 then IfFdNo= FdNo from Fgd(FdBYy= IfFdBYy, FdBChr= IfFdBChr, FdBNo= IfFdBNo and TcType= 'FB')
      'If SourceColName= IfFdTc or IfFdYy or IfFdChr then IfFdNo= 0
  '*** IfFdSr on IfFdBYy,IfFdBChr,IfFdBNo,IfFdTc,IfFdYy,IfFdChr,IfFdNo  Normal
      'If SourceColName= IfFdBYy or IfFdBChr or IfFdBNo then
      '   if IfFdBNo<> 0 then IfFdSr= FdSr from Fgd(FdBYy= IfFdBYy, FdBChr= IfFdBChr, FdBNo= IfFdBNo and TcType= 'FB')
      'If SourceColName= IfFdTc or IfFdYy or IfFdChr or IfFdNo then IfFdSr= 0
  '*** IfQty on IfFdTc,IfFdYy,IfFdChr,IfFdNo,IfFdSr                     Normal
      'IfQty= FdQty-FdDespQty from Fgd(FdTc= IfFdTc, FdYy= IfFdYy, FdChr= IfFdChr, FdNo= IfFdNo, FdSr= IfFdSr)
  '*** IfGrsWt on IfFdTc,IfFdYy,IfFdChr,IfFdNo,IfFdSr                   Normal
      'IfGrsWt= FdGrWt*(FdQty-FdDespQty)/FdQty from Fgd(FdTc= IfFdTc, FdYy= IfFdYy, FdChr= IfFdChr, FdNo= IfFdNo, FdSr= IfFdSr)
  
  With GrdIB
    Select Case ColName
    Case Is = UCase("IfFdBYy")
      If .Mode = fgmnorm Then
        '****** Sachin 3.02 - Corresponding Partition
        If .Value(RowNum, "IfFdSr") <> 0 And moCn.RecSeek("Select FdBYy From Fgd where " + _
                        "FdCoCd= '" + gs_CoCd + "' and " + _
                        "FdTc= '" + .Value(RowNum, "IfFdTc") + "' and " + _
                        "FdYy= '" + .Value(RowNum, "IfFdYy") + "' and " + _
                        "FdChr= '" + .Value(RowNum, "IfFdChr") + "' and " + _
                        "FdNo= " + CStr(.Value(RowNum, "IfFdNo")) + " and " + _
                        "FdSr= " + CStr(.Value(RowNum, "IfFdSr")) + IIF(ADC("wIdPrtKey") = ctCurrPrtn, " and FdPrtKey='" + ctCurrPrtn + "' ", "")) Then _
          .Value(RowNum, "IfFdBYy") = moCn.GetFldVal("Select FdBYy From Fgd where " + _
                                      "FdCoCd= '" + gs_CoCd + "' and " + _
                                      "FdTc= '" + .Value(RowNum, "IfFdTc") + "' and " + _
                                      "FdYy= '" + .Value(RowNum, "IfFdYy") + "' and " + _
                                      "FdChr= '" + .Value(RowNum, "IfFdChr") + "' and " + _
                                      "FdNo= " + CStr(.Value(RowNum, "IfFdNo")) + " and " + _
                                      "FdSr= " + CStr(.Value(RowNum, "IfFdSr")) + IIF(ADC("wIdPrtKey") = ctCurrPrtn, " and FdPrtKey='" + ctCurrPrtn + "' ", ""))
      End If
    Case Is = UCase("IfFdBChr")
      If .Mode = fgmnorm Then
        If .Value(RowNum, "IfFdSr") <> 0 And moCn.RecSeek("Select FdBChr From Fgd where " + _
                        "FdCoCd= '" + gs_CoCd + "' and " + _
                        "FdTc= '" + .Value(RowNum, "IfFdTc") + "' and " + _
                        "FdYy= '" + .Value(RowNum, "IfFdYy") + "' and " + _
                        "FdChr= '" + .Value(RowNum, "IfFdChr") + "' and " + _
                        "FdNo= " + CStr(.Value(RowNum, "IfFdNo")) + " and " + _
                        "FdSr= " + CStr(.Value(RowNum, "IfFdSr")) + IIF(ADC("wIdPrtKey") = ctCurrPrtn, " and FdPrtKey='" + ctCurrPrtn + "' ", "")) Then _
          .Value(RowNum, "IfFdBChr") = moCn.GetFldVal("Select FdBChr From Fgd where " + _
                                      "FdCoCd= '" + gs_CoCd + "' and " + _
                                      "FdTc= '" + .Value(RowNum, "IfFdTc") + "' and " + _
                                      "FdYy= '" + .Value(RowNum, "IfFdYy") + "' and " + _
                                      "FdChr= '" + .Value(RowNum, "IfFdChr") + "' and " + _
                                      "FdNo= " + CStr(.Value(RowNum, "IfFdNo")) + " and " + _
                                      "FdSr= " + CStr(.Value(RowNum, "IfFdSr")) + IIF(ADC("wIdPrtKey") = ctCurrPrtn, " and FdPrtKey='" + ctCurrPrtn + "' ", ""))
      End If
    Case Is = UCase("IfFdBNo")
      If .Mode = fgmnorm Then
        Select Case UCase(SourceColName)
        Case Is = UCase("IfFdBYy"), UCase("IfFdBChr")
          .Value(RowNum, "IfFdBNo") = 0
        Case Is = UCase("IfFdTc"), UCase("IfFdYy"), UCase("IfFdChr"), UCase("IfFdNo"), UCase("IfFdSr")
          '****** Sachin 3.02 -
          If .Value(RowNum, "IfFdSr") <> 0 And moCn.RecSeek("Select FdBNo From Fgd where " + _
                          "FdCoCd= '" + gs_CoCd + "' and " + _
                          "FdTc= '" + .Value(RowNum, "IfFdTc") + "' and " + _
                          "FdYy= '" + .Value(RowNum, "IfFdYy") + "' and " + _
                          "FdChr= '" + .Value(RowNum, "IfFdChr") + "' and " + _
                          "FdNo= " + CStr(.Value(RowNum, "IfFdNo")) + " and " + _
                          "FdSr= " + CStr(.Value(RowNum, "IfFdSr")) + IIF(ADC("wIdPrtKey") = ctCurrPrtn, " and FdPrtKey='" + ctCurrPrtn + "' ", "")) Then _
            .Value(RowNum, "IfFdBNo") = moCn.GetFldVal("Select FdBNo From Fgd where " + _
                                        "FdCoCd= '" + gs_CoCd + "' and " + _
                                        "FdTc= '" + .Value(RowNum, "IfFdTc") + "' and " + _
                                        "FdYy= '" + .Value(RowNum, "IfFdYy") + "' and " + _
                                        "FdChr= '" + .Value(RowNum, "IfFdChr") + "' and " + _
                                        "FdNo= " + CStr(.Value(RowNum, "IfFdNo")) + " and " + _
                                        "FdSr= " + CStr(.Value(RowNum, "IfFdSr")) + IIF(ADC("wIdPrtKey") = ctCurrPrtn, " and FdPrtKey='" + ctCurrPrtn + "' ", ""))
        End Select
      End If
    Case Is = UCase("IfFdTc")
      '### Should a new key be made on FdBYy, FdBChr, FdBNo ?     '???
      If .Mode = fgmnorm Then
        If .Value(RowNum, "IfFdBNo") <> 0 Then
          If UCase(ms_IdTcTyp) = UCase(ctInTcTyp) Then
            .Value(RowNum, "IfFdTc") = moCn.GetFldVal("Select FdTc From Fgd where " + _
                                       "FdCoCd= '" + gs_CoCd + "' and " + _
                                       "FdBYy= '" + .Value(RowNum, "IfFdBYy") + "' and " + _
                                       "FdBChr= '" + .Value(RowNum, "IfFdBChr") + "' and " + _
                                       "FdBNo= " + CStr(.Value(RowNum, "IfFdBNo")) + IIF(ADC("wIdPrtKey") = ctCurrPrtn, " and FdPrtKey='" + ctCurrPrtn + "' ", "") + " and " + _
                                       "(Select PValue from Param where PTyp= 'TC' and " + _
                                       "PMCd= FdTc)= 'FB' and Round(FdQty- FdDespQty, 1)> 0 ")
          ElseIf UCase(ms_IdTcTyp) = UCase(ctCTBTcTyp) Then
            .Value(RowNum, "IfFdTc") = moCn.GetFldVal("Select FdTc From Fgd Fb where " + _
                                       "FdCoCd= '" + gs_CoCd + "' and " + _
                                       "FdBYy= '" + .Value(RowNum, "IfFdBYy") + "' and " + _
                                       "FdBChr= '" + .Value(RowNum, "IfFdBChr") + "' and " + _
                                       "FdBNo= " + CStr(.Value(RowNum, "IfFdBNo")) + IIF(ADC("wIdPrtKey") = ctCurrPrtn, " and FdPrtKey='" + ctCurrPrtn + "'", "") + " and " + _
                                       "(Select PValue from Param where PTyp= 'TC' and " + _
                                       "PMCd= FdTc)= 'FB' and Round(FdQty, 1)> Round(FdCTBDespQty, 1) and " + _
                                       "not exists(Select 'x' from Fgd Fwb where " + _
                                       "(Select PValue from Param where PTyp= 'TC' and " + _
                                       "PMCd= Fwb.FdTc)= 'FWB' and Fwb.FdPrtKey=Fb.FdPrtKey and Fwb.FdCoCd= Fb.FdCoCd and Fwb.FdRefYy= Fb.FdYy and Fwb.FdRefKey= Fb.FdKey) ")
          End If
        End If
      End If
    Case Is = UCase("IfFdYy")
      If .Mode = fgmnorm Then
        If .Value(RowNum, "IfFdBNo") <> 0 Then
          If UCase(ms_IdTcTyp) = UCase(ctInTcTyp) Then
            .Value(RowNum, "IfFdYy") = moCn.GetFldVal("Select FdYy From Fgd where " + _
                                       "FdCoCd= '" + gs_CoCd + "' and " + _
                                       "FdBYy= '" + .Value(RowNum, "IfFdBYy") + "' and " + _
                                       "FdBChr= '" + .Value(RowNum, "IfFdBChr") + "' and " + _
                                       "FdBNo= " + CStr(.Value(RowNum, "IfFdBNo")) + IIF(ADC("wIdPrtKey") = ctCurrPrtn, " and FdPrtKey='" + ctCurrPrtn + "'", "") + " and " + _
                                       "(Select PValue from Param where PTyp= 'TC' and " + _
                                       "PMCd= FdTc)= 'FB' and FdQty- FdDespQty> 0 ")
          ElseIf UCase(ms_IdTcTyp) = UCase(ctCTBTcTyp) Then
            .Value(RowNum, "IfFdYy") = moCn.GetFldVal("Select FdYy From Fgd Fb where " + _
                                       "FdCoCd= '" + gs_CoCd + "' and " + _
                                       "FdBYy= '" + .Value(RowNum, "IfFdBYy") + "' and " + _
                                       "FdBChr= '" + .Value(RowNum, "IfFdBChr") + "' and " + _
                                       "FdBNo= " + CStr(.Value(RowNum, "IfFdBNo")) + IIF(ADC("wIdPrtKey") = ctCurrPrtn, " and FdPrtKey='" + ctCurrPrtn + "'", "") + " and " + _
                                       "(Select PValue from Param where PTyp= 'TC' and " + _
                                       "PMCd= FdTc)= 'FB' and Round(FdQty, 1)> Round(FdCTBDespQty, 1) and " + _
                                       "not exists(Select 'x' from Fgd Fwb where " + _
                                       "(Select PValue from Param where PTyp= 'TC' and " + _
                                       "PMCd= Fwb.FdTc)= 'FWB' and Fwb.FdPrtKey=Fb.FdPrtKey and Fwb.FdCoCd= Fb.FdCoCd and Fwb.FdRefYy= Fb.FdYy and Fwb.FdRefKey= Fb.FdKey) ")
          End If
        End If
      End If
    Case Is = UCase("IfFdChr")
      If .Mode = fgmnorm Then
        If .Value(RowNum, "IfFdBNo") <> 0 Then
          If UCase(ms_IdTcTyp) = UCase(ctInTcTyp) Then
            '****** Sachin 3.02 28-12-07
            .Value(RowNum, "IfFdChr") = moCn.GetFldVal("Select FdChr From Fgd where " + _
                                        "FdCoCd= '" + gs_CoCd + "' and " + _
                                        "FdBYy= '" + .Value(RowNum, "IfFdBYy") + "' and " + _
                                        "FdBChr= '" + .Value(RowNum, "IfFdBChr") + "' and " + _
                                        "FdBNo= " + CStr(.Value(RowNum, "IfFdBNo")) + IIF(ADC("wIdPrtKey") = ctCurrPrtn, " And FdPrtKey='" + ctCurrPrtn + "' ", "") + " and " + _
                                        "(Select PValue from Param where PTyp= 'TC' and " + _
                                        "PMCd= FdTc)= 'FB' and FdQty- FdDespQty> 0 ")
          ElseIf UCase(ms_IdTcTyp) = UCase(ctCTBTcTyp) Then
            .Value(RowNum, "IfFdChr") = moCn.GetFldVal("Select FdChr From Fgd Fb where " + _
                                        "FdCoCd= '" + gs_CoCd + "' and " + _
                                        "FdBYy= '" + .Value(RowNum, "IfFdBYy") + "' and " + _
                                        "FdBChr= '" + .Value(RowNum, "IfFdBChr") + "' and " + _
                                        "FdBNo= " + CStr(.Value(RowNum, "IfFdBNo")) + IIF(ADC("wIdPrtKey") = ctCurrPrtn, " FdPrtKey='" + ctCurrPrtn + "' ", "") + " and " + _
                                        "(Select PValue from Param where PTyp= 'TC' and " + _
                                        "PMCd= FdTc)= 'FB' and Round(FdQty, 1)> Round(FdCTBDespQty, 1) and " + _
                                        "not exists(Select 'x' from Fgd Fwb where " + _
                                        "(Select PValue from Param where PTyp= 'TC' and " + _
                                        "PMCd= Fwb.FdTc)= 'FWB' and Fwb.FdPrtKey=Fb.FdPrtKey and Fwb.FdCoCd= Fb.FdCoCd and Fwb.FdRefYy= Fb.FdYy and Fwb.FdRefKey= Fb.FdKey) ")
          End If
        End If
      End If
    Case Is = UCase("IfFdNo")
      If .Mode = fgmnorm Then
        Select Case UCase(SourceColName)
        Case Is = UCase("IfFdBYy"), UCase("IfFdBChr"), UCase("IfFdBNo")
          If .Value(RowNum, "IfFdBNo") <> 0 Then
            If UCase(ms_IdTcTyp) = UCase(ctInTcTyp) Then
              '****** Sachin 3.02 28-12-07
              .Value(RowNum, "IfFdNo") = moCn.GetFldVal("Select FdNo From Fgd where " + _
                                         "FdCoCd= '" + gs_CoCd + "' and " + _
                                         "FdBYy= '" + .Value(RowNum, "IfFdBYy") + "' and " + _
                                         "FdBChr= '" + .Value(RowNum, "IfFdBChr") + "' and " + _
                                         "FdBNo= " + CStr(.Value(RowNum, "IfFdBNo")) + IIF(ADC("wIdPrtKey") = ctCurrPrtn, " And FdPrtKey='" + ctCurrPrtn + "' ", "") + " and " + _
                                         "(Select PValue from Param where PTyp= 'TC' and " + _
                                         "PMCd= FdTc)= 'FB' and FdQty- FdDespQty> 0 ")
            ElseIf UCase(ms_IdTcTyp) = UCase(ctCTBTcTyp) Then
              .Value(RowNum, "IfFdNo") = moCn.GetFldVal("Select FdNo From Fgd Fb where " + _
                                         "FdCoCd= '" + gs_CoCd + "' and " + _
                                         "FdBYy= '" + .Value(RowNum, "IfFdBYy") + "' and " + _
                                         "FdBChr= '" + .Value(RowNum, "IfFdBChr") + "' and " + _
                                         "FdBNo= " + CStr(.Value(RowNum, "IfFdBNo")) + IIF(ADC("wIdPrtKey") = ctCurrPrtn, " And FdPRtKey='" + ctCurrPrtn + "' ", "") + " and " + _
                                         "(Select PValue from Param where PTyp= 'TC' and " + _
                                         "PMCd= FdTc)= 'FB' and Round(FdQty, 1)> Round(FdCTBDespQty, 1) and " + _
                                         "not exists(Select 'x' from Fgd Fwb where " + _
                                         "(Select PValue from Param where PTyp= 'TC' and " + _
                                         "PMCd= Fwb.FdTc)= 'FWB' and Fwb.FdPrtKey=Fb.FdPrtKey and Fwb.FdCoCd= Fb.FdCoCd and Fwb.FdRefYy= Fb.FdYy and Fwb.FdRefKey= Fb.FdKey) ")
            End If
          End If
        Case Is = UCase("IfFdTc"), UCase("IfFdYy"), UCase("IfFdChr")
          .Value(RowNum, "IfFdNo") = 0
        End Select
      End If
    Case Is = UCase("IfFdSr")
      If .Mode = fgmnorm Then
        Select Case UCase(SourceColName)
        Case Is = UCase("IfFdBYy"), UCase("IfFdBChr"), UCase("IfFdBNo")
          If .Value(RowNum, "IfFdBNo") <> 0 Then
            If UCase(ms_IdTcTyp) = UCase(ctInTcTyp) Then
              '****** Sachin 3.02 28-12-07
              .Value(RowNum, "IfFdSr") = moCn.GetFldVal("Select FdSr From Fgd where " + _
                                       "FdCoCd= '" + gs_CoCd + "' and " + _
                                       "FdBYy= '" + .Value(RowNum, "IfFdBYy") + "' and " + _
                                       "FdBChr= '" + .Value(RowNum, "IfFdBChr") + "' and " + _
                                       "FdBNo= " + CStr(.Value(RowNum, "IfFdBNo")) + IIF(ADC("wIdPrtKey") = ctCurrPrtn, " And FdPrtKey='" + ctCurrPrtn + "' ", "") + " and " + _
                                       "(Select PValue from Param where PTyp= 'TC' and " + _
                                       "PMCd= FdTc)= 'FB' and FdQty- FdDespQty> 0 ")
            ElseIf UCase(ms_IdTcTyp) = UCase(ctCTBTcTyp) Then
              .Value(RowNum, "IfFdSr") = moCn.GetFldVal("Select FdSr From Fgd Fb where " + _
                                         "FdCoCd= '" + gs_CoCd + "' and " + _
                                         "FdBYy= '" + .Value(RowNum, "IfFdBYy") + "' and " + _
                                         "FdBChr= '" + .Value(RowNum, "IfFdBChr") + "' and " + _
                                         "FdBNo= " + CStr(.Value(RowNum, "IfFdBNo")) + IIF(ADC("wIdPrtKey") = ctCurrPrtn, " And FdPRtKey='" + ctCurrPrtn + "' ", "") + " and " + _
                                         "(Select PValue from Param where PTyp= 'TC' and " + _
                                         "PMCd= FdTc)= 'FB' and Round(FdQty, 1)> Round(FdCTBDespQty, 1) and " + _
                                         "not exists(Select 'x' from Fgd Fwb where " + _
                                         "(Select PValue from Param where PTyp= 'TC' and " + _
                                         "PMCd= Fwb.FdTc)= 'FWB' and Fwb.FdPrtKey=Fb.FdPrtKey and Fwb.FdCoCd= Fb.FdCoCd and Fwb.FdRefYy= Fb.FdYy and Fwb.FdRefKey= Fb.FdKey) ")
            End If
          End If
        Case Is = UCase("IfFdTc"), UCase("IfFdYy"), UCase("IfFdChr"), UCase("IfFdNo")
          .Value(RowNum, "IfFdSr") = 0
        End Select
      End If
    Case Is = UCase("IfQty")
      If .Mode = fgmnorm Then
        .Value(RowNum, "IfQty") = moCn.GetFldVal("Select FdQty- " + _
                                  IIF(UCase(ms_IdTcTyp) = UCase(ctInTcTyp), "FdDespQty", "FdCTBDespQty") + " " + _
                                  "From Fgd where FdCoCd= '" + gs_CoCd + "' and " + _
                                  "FdTc= '" + .Value(RowNum, "IfFdTc") + "' and " + _
                                  "FdYy= '" + .Value(RowNum, "IfFdYy") + "' and " + _
                                  "FdChr= '" + .Value(RowNum, "IfFdChr") + "' and " + _
                                  "FdNo= " + CStr(.Value(RowNum, "IfFdNo")) + " and " + _
                                  "FdSr= " + CStr(.Value(RowNum, "IfFdSr")) + IIF(ADC("wIdPrtKey") = ctCurrPrtn, " and FdPrtKey='" + ctCurrPrtn + "' ", ""))
      End If
    Case Is = UCase("IfGrsWt")
      If .Mode = fgmnorm Then
        .Value(RowNum, "IfGrsWt") = moCn.GetFldVal("Select " + _
              "(case when FdQty> 0 then FdGrWt*(FdQty- " + _
              IIF(UCase(ms_IdTcTyp) = UCase(ctInTcTyp), "FdDespQty", "FdCTBDespQty") + ")" + _
              "/FdQty else 0 end) From Fgd where FdCoCd= '" + gs_CoCd + "' and " + _
              "FdTc= '" + .Value(RowNum, "IfFdTc") + "' and " + _
              "FdYy= '" + .Value(RowNum, "IfFdYy") + "' and " + _
              "FdChr= '" + .Value(RowNum, "IfFdChr") + "' and " + _
              "FdNo= " + CStr(.Value(RowNum, "IfFdNo")) + " and " + _
              "FdSr= " + CStr(.Value(RowNum, "IfFdSr")) + IIF(ADC("wIdPrtKey") = ctCurrPrtn, " and FdPrtKey='" + ctCurrPrtn + "' ", ""))
      End If
    End Select
  End With
End Sub
Private Sub GrdIB_ColChange(ByVal RowNum As Integer, ByVal ColName As String, ByVal NewValue As Variant, Cancel As Boolean, ErrMsg As String)
  '*** IfFdBYy        valid PMCd('YY') or Blank
  '*** IfFdBChr       valid PMCd('BCHR')  or Blank
  '*** IfFdBNo        valid FdBNo from Fgd(FdBYy= IfFdBYy, FdBChr= IfFdBChr, FdQty-FdDespQty> 0) or Zero
  '*** IfFdTc         (valid PMCd('TC', PValue= 'FB' or 'FR') and the type of entry should be 'FgBag' or 'FgRm') or Blank
  '*** IfFdYy         valid PMCd('YY') or Blank
  '*** IfFdChr        valid PSCd('CHR', IfFdTc) or ''
  '*** IfFdNo         valid FgNo from Fg(FgTc= IfFdTc, FgYy= IfFdYy, FgChr= IfFdChr) or Zero
  '*** IfFdSr         (valid FdSr from Fgd(FdTc= IfFdTc, FdYy= IfFdYy, FdChr= IfFdChr, FdNo= IfFdNo, FdQty-FdDespQty> 0)) or Zero
  
  '*** (Jen 3.2.0)
  '****** Sachin 3.02.0 28-12-07 - prtkey added below
  If GrdIB.Mode = fgmnorm Then
    If moCn.RecSeek("Select 'x' from InvZRm where IzCoCd= '" + gs_CoCd + "' and IzTc= '" + ADC("wIdTc") + "' and " + _
                    "IzYy=  '" + ADC("wIdYy") + "' and IzChr=  '" + ADC("wIdChr") + "' and " + _
                    "IzNo=  " + CStr(ADC("wIdNo")) + " and IzPrtKey= '" + ctCurrPrtn + "' ") Then
      Cancel = True: ErrMsg = "Cannot Edit As Custom Matching Has Been Done": Exit Sub
    End If
  End If
  '*** (Jen 3.2.0)
  
  With GrdIB
    Select Case UCase(ColName)
    Case Is = UCase("IfFdBYy")
      Cancel = (NewValue <> "") And (Not moCn.RecSeek("Select PMCd From Param where " + _
                                    "PTyp= 'YY' and PMCd = '" + NewValue + "'"))
      ErrMsg = "Invalid Bag Year": Exit Sub
    Case Is = UCase("IfFdBChr")
' **** Zubin 211 **** '
'''''      Cancel = (NewValue <> "") And (Not moCn.RecSeek("Select PMCd From Param where " + _
                                    "PTyp= 'BCHR' and PMCd = '" + NewValue + "'"))
'''''      ErrMsg = "Invalid Bag Character": Exit Sub
      Cancel = (NewValue <> "") And (Not moCn.RecSeek("Select PMCd From Param where " + _
                                    "PTyp= 'BCHR' and PMCd = '" + NewValue + "' and PValue= 'P'"))
      ErrMsg = "Invalid Bag Character": Exit Sub
' **** Zubin 211 **** '
    Case Is = UCase("IfFdBNo")
      '*** (Bef 3.01)
      'Cancel = (NewValue <> 0) And (Not moCn.RecSeek("Select BNo From Bag where " + _
      '                              "BCoCd= '" + gs_CoCd + "' and " + _
      '                              "BYy= '" + .Value(RowNum, "IfFdBYy") + "' and " + _
      '                              "BChr= '" + .Value(RowNum, "IfFdBChr") + "' and " + _
      '                              "BNo= " + CStr(NewValue) + " and " + _
      '                              "BLoc= '" + ctFgLoc + "' "))
      'If Cancel = True Then ErrMsg = "Invalid Bag Number/ Bag Not In FG Loc": Exit Sub
      '*** (Bef 3.01)
    
      '*** (Jen 3.01)
      If NewValue <> 0 Then
        '****** Sachin 3.02 28-12-07
        Cancel = Not moCn.RecSeek("Select BNo From Bag where " + _
                                  "BCoCd= '" + gs_CoCd + "' and " + _
                                  "BYy= '" + .Value(RowNum, "IfFdBYy") + "' and " + _
                                  "BChr= '" + .Value(RowNum, "IfFdBChr") + "' and " + _
                                  "BNo= " + CStr(NewValue) + " and " + _
                                  "BLoc= '" + ctFgLoc + "' and BPrtKey='" + ctCurrPrtn + "' ")
        If Cancel = True Then ErrMsg = "Invalid Bag Number/ Bag Not In FG Loc": Exit Sub
      
        If .Mode = fgmnorm Then
          '****** Sachin 3.02 28-12-07
          Cancel = Not moCn.RecSeek("Select FdSr From Fgd Fb where " + _
                                    "FdCoCd= '" + gs_CoCd + "' and " + _
                                    "FdBYy= '" + .Value(RowNum, "IfFdBYy") + "' and " + _
                                    "FdBChr= '" + .Value(RowNum, "IfFdBChr") + "' and " + _
                                    "FdBNo= " + CStr(NewValue) + " and FdPrtKey='" + ctCurrPrtn + "' and " + _
                                    "FdExpOdTc= '" + GrdID.Value(GrdID.Row, "IdExpOdTc") + "' and " + _
                                    "FdExpOdYy= '" + GrdID.Value(GrdID.Row, "IdExpOdYy") + "' and " + _
                                    "FdExpOdChr= '" + GrdID.Value(GrdID.Row, "IdExpOdChr") + "' and " + _
                                    "FdExpOdNo= " + CStr(GrdID.Value(GrdID.Row, "IdExpOdNo")) + " and " + _
                                    "FdExpOdSr= " + CStr(GrdID.Value(GrdID.Row, "IdExpOdSr")) + " and " + _
                                    "(Select PValue from Param where PTyp= 'TC' and " + _
                                    "PMCd= FdTc)= 'FB' and Round(FdQty, 1)> Round(FdCTBDespQty, 1) and " + _
                                    "not exists(Select 'x' from Fgd Fwb where " + _
                                    "(Select PValue from Param where PTyp= 'TC' and " + _
                                    "PMCd= Fwb.FdTc)= 'FWB' and Fwb.FdPrtKey=Fb.FdPrtKey and Fwb.FdCoCd= Fb.FdCoCd and Fwb.FdRefYy= Fb.FdYy and Fwb.FdRefKey= Fb.FdKey) ")
          If Cancel = True Then ErrMsg = "Fg Export Order Not Matching With Invoice Design Export Order": Exit Sub
        End If
      End If
      '*** (Jen 3.01)
      
    Case Is = UCase("IfFdTc")
      '*** (Jen 3.01)
      If .Mode = fgmnorm Then
      '*** (Jen 3.01)

        Cancel = (NewValue <> "") And (Not moCn.RecSeek("select PMCd from Param where " + _
                             "PTyp= 'TC' and PMCd = '" + NewValue + "' and " + _
                             "PValue in ('FB', 'FR')"))
        If Cancel = True Then ErrMsg = "Invalid Finished Goods Voucher Tc": Exit Sub
      End If
    Case Is = UCase("IfFdYy")
      '*** (Jen 3.01)
      If .Mode = fgmnorm Then
      '*** (Jen 3.01)
    
        Cancel = (NewValue <> "") And (Not moCn.RecSeek("select PMCd from Param where " + _
                             "PTyp= 'YY' and PMCd = '" + NewValue + "'"))
        If Cancel = True Then ErrMsg = "Invalid Finished Goods Voucher Year": Exit Sub
      End If
    Case Is = UCase("IfFdChr")
      '*** (Jen 3.01)
      If .Mode = fgmnorm Then
      '*** (Jen 3.01)
    
        Cancel = (NewValue <> "") And (Not moCn.RecSeek("select vPSCd from vParam where " + _
                 "vPCoCd= '" + gs_CoCd + "' and vPTyp= 'CHR' and " + _
                 "vPMCd = '" + .Value(RowNum, "IfFdTc") + "' " + _
                 "and vPSCd='" + NewValue + "'"))
        ErrMsg = "Invalid Finished Goods Voucher Character": Exit Sub
      End If
    Case Is = UCase("IfFdNo")
      '*** (Jen 3.01)
      If .Mode = fgmnorm Then
      '*** (Jen 3.01)
        '****** Sachin 3.02 28-12-07 - Current Partition
        Cancel = (NewValue <> 0) And (Not moCn.RecSeek("Select FgNo From Fg where " + _
                                      "FgCoCd= '" + gs_CoCd + "' and " + _
                                      "FgTc= '" + .Value(RowNum, "IfFdTc") + "' and " + _
                                      "FgYy= '" + .Value(RowNum, "IfFdYy") + "' and " + _
                                      "FgChr= '" + .Value(RowNum, "IfFdChr") + "' and " + _
                                      "FgNo= " + CStr(NewValue) + " and FgPrtKey='" + ctCurrPrtn + "' "))
        ErrMsg = "Invalid Finished Goods Voucher Number": Exit Sub
      End If
    Case Is = UCase("IfFdSr")
      '*** (Bef 3.01)
      'Cancel = (NewValue <> 0) And (Not moCn.RecSeek("Select FdSr From Fgd where " + _
      '                              "FdCoCd= '" + gs_CoCd + "' and " + _
      '                              "FdTc= '" + .Value(RowNum, "IfFdTc") + "' and " + _
      '                              "FdYy= '" + .Value(RowNum, "IfFdYy") + "' and " + _
      '                              "FdChr= '" + .Value(RowNum, "IfFdChr") + "' and " + _
      '                              "FdNo= " + CStr(.Value(RowNum, "IfFdNo")) + " and " + _
      '                              "FdSr= " + CStr(NewValue)))
      'If Cancel = True Then ErrMsg = "Invalid Finished Goods Voucher Serial": Exit Sub
      '*** (Bef 3.01)
    
      '*** (Jen 3.01)
      If .Mode = fgmnorm Then
        If NewValue <> 0 Then
          '****** Sachin 3.02 28-12-07 - Current Partition
          Cancel = Not moCn.RecSeek("Select FdSr From Fgd where " + _
                                    "FdCoCd= '" + gs_CoCd + "' and " + _
                                    "FdTc= '" + .Value(RowNum, "IfFdTc") + "' and " + _
                                    "FdYy= '" + .Value(RowNum, "IfFdYy") + "' and " + _
                                    "FdChr= '" + .Value(RowNum, "IfFdChr") + "' and " + _
                                    "FdNo= " + CStr(.Value(RowNum, "IfFdNo")) + " and " + _
                                    "FdSr= " + CStr(NewValue) + " and FdPrtKey='" + ctCurrPrtn + "' ")
          
          If Cancel = True Then ErrMsg = "Invalid Finished Goods Voucher Serial": Exit Sub
          
          '****** Sachin 3.02 28-12-07 - Current Partition
          Cancel = Not moCn.RecSeek("Select FdSr From Fgd Fb where " + _
                                    "FdCoCd= '" + gs_CoCd + "' and " + _
                                    "FdTc= '" + .Value(RowNum, "IfFdTc") + "' and " + _
                                    "FdYy= '" + .Value(RowNum, "IfFdYy") + "' and " + _
                                    "FdChr= '" + .Value(RowNum, "IfFdChr") + "' and " + _
                                    "FdNo= " + CStr(.Value(RowNum, "IfFdNo")) + " and " + _
                                    "FdSr= " + CStr(NewValue) + " and FdPrtKey='" + ctCurrPrtn + "' and " + _
                                    "FdExpOdTc= '" + GrdID.Value(GrdID.Row, "IdExpOdTc") + "' and " + _
                                    "FdExpOdYy= '" + GrdID.Value(GrdID.Row, "IdExpOdYy") + "' and " + _
                                    "FdExpOdChr= '" + GrdID.Value(GrdID.Row, "IdExpOdChr") + "' and " + _
                                    "FdExpOdNo= " + CStr(GrdID.Value(GrdID.Row, "IdExpOdNo")) + " and " + _
                                    "FdExpOdSr= " + CStr(GrdID.Value(GrdID.Row, "IdExpOdSr")) + " and " + _
                                    "(Select PValue from Param where PTyp= 'TC' and " + _
                                    "PMCd= FdTc)= 'FB' and Round(FdQty, 1)> Round(FdCTBDespQty, 1) and " + _
                                    "not exists(Select 'x' from Fgd Fwb where " + _
                                    "(Select PValue from Param where PTyp= 'TC' and " + _
                                    "PMCd= Fwb.FdTc)= 'FWB' and Fwb.FdPrtKey=Fb.FdPrtKey and Fwb.FdCoCd= Fb.FdCoCd and Fwb.FdRefYy= Fb.FdYy and Fwb.FdRefKey= Fb.FdKey) ")
          If Cancel = True Then ErrMsg = "Fg Export Order Not Matching With Invoice Design Export Order": Exit Sub
        End If
      End If
      '*** (Jen 3.01)
    End Select
  End With
End Sub
Private Sub GrdIB_InitKey(ByVal RowNum As Integer)
  '*** Initialise Key Fields (i.e. IfTc, IfYy, IfChr, IfNo, IfSr)
  With GrdIB
    If mf_inspecialCopy Then
    .Store "Cpy", "Y"
    Else
    .Store "Cpy", ""
    End If
    
    .Store "IfCoCd", gs_CoCd
    .Store "IfTc", ADC("wIdTc")
    .Store "IfYy", ADC("wIdYy")
    .Store "IfChr", ADC("wIdChr")
    .Store "IfNo", ADC("wIdNo")
    .Store "IfSr", GrdID.Value(GrdID.Row, "IdSr")
  
   '****** Sachin 3.02 - Initializing Id fields
    If .IsNew(RowNum) Then
        .Store "IfPrtKey", ctCurrPrtn       '****** Sachin 3.02 - 28-12-07
        .Store "IfInIdNo", moCn.GetFldVal("Select InIdNo From InvHd " + _
                                          " Where InPrtKey='" + ctCurrPrtn + "' And InCoCd='" + gs_CoCd + "' And InTc='" + ADC("wIdTc") + "' And  InYy='" + ADC("wIdYy") + "' And InChr='" + ADC("wIdChr") + "' And InNo=" + CStr(ADC("wIdNo")))
                                          
                                                 
                                                         
        Dim ws_str() As String
        ws_str = Split(moCn.GetFldVal("Select LTrim(str(FdIdNo))+ '|'+ LTrim(Str(FdBIdNo)) From Fgd " + _
                                          " Where FdPrtKey='" + ctCurrPrtn + "' And FdCoCd= '" + gs_CoCd + "' And FdTc= '" + GrdIB.Value(RowNum, "IfFdTc") + "' And  FdYy= '" + GrdIB.Value(RowNum, "IfFdYy") + "' And FdChr= '" + GrdIB.Value(RowNum, "IfFdChr") + "' And FdNo= " + CStr(GrdIB.Value(RowNum, "IfFdNo")) + " And FdSr= " + CStr(GrdIB.Value(RowNum, "IfFdSr"))), "|")
        If UBound(ws_str) > 0 Then
          .Store "IfFdIdNo", ws_str(0)
          .Store "IfBIdNo", ws_str(1)
        Else
          '****** Sachin 3.02.0 - Can go in else only if erroneous BagNo is entered and Fg Vch No Is Not Recalculated - Trigger level error is displayed
          .Store "IfFdIdNo", 0
          .Store "IfBIdNo", 0
        End If
    End If
  End With
  
End Sub

Private Sub GrdSummDR_ColWhen(ByVal RowNum As Integer, ByVal ColName As String, Cancel As Boolean, ErrMsg As String, ByVal Action As MwfCtl.en_fgAction)
  '*** Cannot edit Design Rm Summary grid values
  Cancel = True
End Sub
Private Sub GrdSummDL_ColWhen(ByVal RowNum As Integer, ByVal ColName As String, Cancel As Boolean, ErrMsg As String, ByVal Action As MwfCtl.en_fgAction)
  '*** Cannot edit Design Labour Summary grid values
  Cancel = True
End Sub
Private Sub GrdSummIR_ColWhen(ByVal RowNum As Integer, ByVal ColName As String, Cancel As Boolean, ErrMsg As String, ByVal Action As MwfCtl.en_fgAction)
  '*** Cannot edit Inv Rm Summary grid values
  Cancel = True
End Sub
Private Sub GrdSummIL_ColWhen(ByVal RowNum As Integer, ByVal ColName As String, Cancel As Boolean, ErrMsg As String, ByVal Action As MwfCtl.en_fgAction)
  '*** Cannot edit Inv Labour Summary grid values
  Cancel = True
End Sub
Private Sub GrdSummID_ColWhen(ByVal RowNum As Integer, ByVal ColName As String, Cancel As Boolean, ErrMsg As String, ByVal Action As MwfCtl.en_fgAction)
  '*** Cannot edit Inv Design Summary grid values
  Cancel = True
End Sub
Private Sub ADC_NonKeyValid(pr_Cancel As Boolean, pr_ErrMsg As String)
  '*** (Jen 2.12) 06/01/06
	ms_hSizeStkRt = moCn.GetFldVal("Select hSizeStkRt from Head where HCoCd= '" + gs_CoCd + "' and " + _
                          "HCd= '" + ctSelfCmCd + "'") = "N"
  
  If GrdIR.Rows - 1 > 0 Then
    If GrdIR.Row > 0 And Not GrdIR.IsDel(GrdIR.Row) Then GrdIR.SaveRec (GrdIR.Row)
  End If
  If GrdIL.Rows - 1 > 0 Then
    If GrdIL.Row > 0 And Not GrdIL.IsDel(GrdIL.Row) Then GrdIL.SaveRec (GrdIL.Row)
  End If
  If GrdID.Rows - 1 > 0 Then
    If GrdID.Row > 0 And Not GrdID.IsDel(GrdID.Row) Then GrdID.SaveRec (GrdID.Row)
  End If
  '*** (Jen 2.12) 06/01/06
  
  If UCase(ms_IdTcTyp) = UCase(ctInTcTyp) And ADC("wIdPrtKey") = ctCurrPrtn Then
    'Geeta 2.8.0
    Dim wi_IdSr As Integer
  
    '*** (Bef 23/11/05)
    'If Not moCn.RecSeek("Select 'x' from InvZRm where IzCoCd= '" + gs_CoCd + "'  and IzTc='" + adc("wIdTc") + "'" + _
    '       "and IzYy= '" + adc("wIdYy") + "' and IZChr='" + adc("wIdChr") + "' and IzNo= " + CStr(adc("wIdNo"))) And mInCls = "N" Then
    '*** (Bef 23/11/05)
  
    '*** (Jen 23/11/05) (2.11 Patch3)
    '****** Sachin 3.02 28-12-07 - Current Partition
    If gs_CoCd <> "DJ" And gs_CoCd <> "MW" And gs_CoCd <> "MW1" And mInCls = "N" Then
    '*** (Jen 23/11/05) (2.11 Patch3)
      '****** (Jen 2.13) introduced 'set dateformat dmy' in the following check ******
      '****** Sachin 3.02 - Id fields In Joins
      '***** Sachin 3.11.0 - Order Mulby added to Check
      '6.1
      wi_IdSr = moCn.GetFldVal("Set dateformat " + DtFmtStr() + " " + _
                     "Select IdSr From InvHd Join InvDsg " + _
                     " On InIdNo=IdInIdNo and InCoCd = IdCoCd And InTc = IdTc And InYy = IdYy And InChr = IdChr And InNo = IdNo and InPrtKey=IdPRtKEy  " + _
                     " Join OrdDsg On OdCoCd=IdCoCd And OdTc=IdExpOdTc and OdYy=IdExpOdYy and OdChr=IdExpOdChr and OdNo=IdExpOdNo and OdSr=IdExpOdSr " + _
                     " Left Outer Join  (Select IrPrtKey, IrCoCd, IrTc, IrYy, IrChr, IrNo, IrSr, " + _
                     "Sum(Cast(IrRmIVal as Decimal(12,2)) + Cast(IrSetIVal as Decimal(12,2))) as RmVal " + _
                     " From InvRm Group By IrPrtKey, IrCoCd, IrTc, IrYy, IrChr, IrNo, IrSr) InvR " + _
                     "On IdCoCd = IrCoCd And IdTc = IrTc And IdYy = IrYy And IdChr = IrChr And IdNo = IrNo And IdSr = IrSr and IdPrtKey=IrPrtKey " + _
                     "Left Outer Join (Select IlPrtKey, IlCoCd, IlTc, IlYy, IlChr, IlNo, IlSr, " + _
                     "Sum(Cast(isNull(IlIVal, 0) as Decimal(12,2)))  as LabVal " + _
                     "From InvLab Group By IlPrtKey, IlCoCd, IlTc, IlYy, IlChr, IlNo, IlSr) InvL  " + _
                     "On IdCoCd = IlCoCd And IdTc = IlTc And IdYy = IlYy And IdChr = IlChr And IdNo = IlNo And IdSr = IlSr And IdPrtKey=IlPrtKey " + _
                     "Where InCoCd = '" + gs_CoCd + "' And InTc = '" + ADC("wIdTc") + "' " + _
                     "And InYy = '" + ADC("wIdYy") + "' And InChr = '" + ADC("wIdChr") + "' " + _
                     "And InNo = " + CStr(ADC("wIdNo")) + " and InPrtKey='" + ctCurrPrtn + "' and InCls= 'N' and InDt>= '01/12/05' and " + _
                     " Abs(Cast((IdIFOB - (((isNull(RmVal, 0) + isNull(LabVal, 0))*(Case When InOrdMulbyYN='Y' Then OdMulby Else 1 End)+ IdIAdd))* (1+ InFobPer/100)) as Decimal(12, 2))) >1  and IdFixPrc= 'N'  ")
      
      'and InDt>='01/01/05'   '*** (Bef 23/11/05)
      If wi_IdSr > 0 Then _
               pr_Cancel = True: pr_ErrMsg = "Fob is not calculatd Properly for Sr " + CStr(wi_IdSr): Exit Sub
      
      '****** Sachin 3.02 28-12-07 - Current Partition
      wi_IdSr = moCn.GetFldVal("Select IdSr From  InvDsg where IdCoCd= '" + gs_CoCd + "'  and IdTc='" + ADC("wIdTc") + "'" + _
                    "and IdYy= '" + ADC("wIdYy") + "' and IdChr='" + ADC("wIdChr") + "' and IdNo= " + CStr(ADC("wIdNo")) + " and IdPrtKey='" + ctCurrPrtn + "' and " + _
                    " Not Exists (Select 'x' From InvRm where IrCoCd= IdCoCd and IrTc=IdTc and IrYy= IdYy and IrChr=IdChr  " + _
                    " and IrNo= IdNo and IrSr= IdSr and IrPrtKey=IdPrtKey)")
                    
      If wi_IdSr > 0 Then _
              pr_Cancel = True: pr_ErrMsg = "Raw Material Details Not Copied for Sr " + CStr(wi_IdSr): Exit Sub
    
    End If
    'Geeta 2.8.0
    '****** Sachin 3.02 28-12-07 - Current Partition
    If moCn.GetFldVal("Select Str(sum(IdIFob), 16, 2) from InvDsg where " + _
                       "IdCoCd= '" + gs_CoCd + "' and IdTc= '" + ADC("wIdTc") + "' and " + _
                       "IdYy= '" + ADC("wIdYy") + "' and IdChr= '" + ADC("wIdChr") + "' and " + _
                       "IdNo= " + CStr(ADC("wIdNo")) + " and IdPrtKey='" + ctCurrPrtn + "' ") _
        <> moCn.GetFldVal("Select Str(sum(IdZFob), 16, 2) from InvDsg where " + _
                       "IdCoCd= '" + gs_CoCd + "' and IdTc= '" + ADC("wIdTc") + "' and " + _
                       "IdYy= '" + ADC("wIdYy") + "' and IdChr= '" + ADC("wIdChr") + "' and " + _
                       "IdNo= " + CStr(ADC("wIdNo")) + " and IdPrtkey='" + ctCurrPrtn + "' ") Then _
        pr_Cancel = True: pr_ErrMsg = "Total Invoice FOB and Customs FOB Not Matching": Exit Sub
        
   '******* Manali 3.2.2 - Check for InvRmWt= bagGrsWt and IdQty=InvFgd bag qty
   Dim wIdRec As MDORowSet, wIdSrStr As String
   Set wIdRec = moCn.OpenRes("Select IdSr From InvDsg " + _
                "left outer join " + _
                "(Select IrPrtKey, IrIdIdNo, IrCoCd, IrTc, IrYy, IrChr, IrNo, IrSr, " + _
                "IsNull(sum(IrRmAWt/(case when IrRmCtg in ('D','C') then 5 else 1 End)), 0) as qIrGrsWt from InvRm " + _
                "Group By IrPrtKey, IrIdIdNo, IrCoCd, IrTc, IrYy, IrChr, IrNo, IrSr) Ir " + _
                "on IdCoCd=IrCoCd and IdTc=IrTc and IdYy=IrYy and IdChr=IrChr and IdNo=IrNo and IdSr=IrSr and IdPrtKey=IrPrtKey and IdIdNo=IrIdIdNo" + _
                " left outer join " + _
                "(Select IfPrtKey, IfCoCd, IfTc, IfYy, IfChr, IfNo, IfSr, IsNull(sum(IfGrsWt), 0) as qIfGrsWt, IsNull(sum(IfQty), 0) as qIfQty from InvFgd " + _
                "Group By IfPrtKey, IfCoCd, IfTc, IfYy, IfChr, IfNo, IfSr) Inf " + _
                "on IrCoCd=IfCoCd and IrTc=IfTc and IrYy=IfYy and IrChr=IfChr and IrNo=IfNo and IrSr=IfSr and IrPrtKey=IfPrtKey " + _
                "where IdCoCd= '" + gs_CoCd + "' and IdTc= '" + ADC("wIdTc") + "' and " + _
                "IdYy= '" + ADC("wIdYy") + "' and IdChr= '" + ADC("wIdChr") + "' and " + _
                "IdNo= " + CStr(ADC("wIdNo")) + " and IdPrtkey='" + ctCurrPrtn + "' and " + _
                "(abs(cast(cast(IsNull(qIrGrsWt, 0) as decimal(14,3)) as decimal(14,2)) - cast(cast(IsNull(qIfGrsWt, 0) as decimal (14,3)) as decimal (14,2))) > 0.9 " + _
                " or cast(IsNull(qIfQty ,0) as decimal(14,0)) <> cast(IsNull(IdQty ,0) as decimal (14,0))) " + _
                "Order By IdCoCd, IdTc, IdYy, IdChr, IdNo, IdSr ")
                
'********* Manali 19/09/08 - instead of checking inequality of wt, diff upto 0.9 [As told by RN] (Old - 0.05) is allowed in the check, as many clients are facing problem frequently
'                "(cast(cast(IsNull(qIrGrsWt, 0) as decimal(14,3)) as decimal(14,2)) <> cast(cast(IsNull(qIfGrsWt, 0) as decimal (14,3)) as decimal (14,2)) " + _
'                " or cast(IsNull(qIfQty ,0) as decimal(14,0)) <> cast(IsNull(IdQty ,0) as decimal (14,0))) " + _
'                "Order By IdCoCd, IdTc, IdYy, IdChr, IdNo, IdSr ")
     
     If wIdRec.RecCount > 0 Then
        With wIdRec
          wIdSrStr = ""
          .MoveFirst
          Do While Not (.BOF Or .EOF)
            wIdSrStr = IIF(wIdSrStr = "", "", wIdSrStr + ",") + CStr(wIdRec.FldValue("IdSr"))
            .MoveNext
          Loop
        End With
      End If
    
    If wIdSrStr <> "" Then _
              pr_Cancel = True: pr_ErrMsg = "Invoice Serial(s) " + wIdSrStr + " Contain(s) Discrepancy, Call MICROWAY Immediately": Exit Sub
              
              
   '******* Sachin 3.11.0 - Check for Rm with Zero Rates
   Dim wRtRec As MDORowSet, wIrSrStr As String
   Set wRtRec = moCn.OpenRes("Select Distinct IrSr From InvRm   " + _
                " Where IrCoCd= '" + gs_CoCd + "' and IrTc= '" + ADC("wIdTc") + "' and " + _
                "  IrYy= '" + ADC("wIdYy") + "' and IrChr= '" + ADC("wIdChr") + "' and " + _
                "  IrNo= " + CStr(ADC("wIdNo")) + " and IrPrtkey='" + ctCurrPrtn + "' And IrRmIRt<=0")
     
    If wRtRec.RecCount > 0 Then
        With wRtRec
          wIrSrStr = ""
          .MoveFirst
          Do While Not (.BOF Or .EOF)
            wIrSrStr = IIF(wIrSrStr = "", "", wIrSrStr + ",") + CStr(wRtRec.FldValue("IrSr"))
            .MoveNext
          Loop
        End With
        
       If MsgBox("Invoice Serial(s) (" + wIrSrStr + ") have Raw Materials with Zero Rates. Do you want to Continue? ", vbYesNo) = vbNo Then _
          pr_Cancel = True: pr_ErrMsg = "Cannot Continue.": Exit Sub
    End If
  '******* Sachin 3.11.0 - Check for Rm with Zero Rates
    '6.1 checks Rm with 0 value and raise warning message
    Dim wIrValRec As MDORowSet, wIrValSr As String
    Set wIrValRec = moCn.OpenRes("Select Distinct IrSr From InvRm   " + _
                  " Where IrCoCd= '" + gs_CoCd + "' and IrTc= '" + ADC("wIdTc") + "' and " + _
                  "  IrYy= '" + ADC("wIdYy") + "' and IrChr= '" + ADC("wIdChr") + "' and " + _
                  "  IrNo= " + CStr(ADC("wIdNo")) + " and IrPrtkey='" + ctCurrPrtn + "' And IrRmIVal<=0")
       
    If wIrValRec.RecCount > 0 Then
      With wIrValRec
        wIrValSr = ""
        .MoveFirst
        Do While Not (.BOF Or .EOF)
          wIrValSr = IIF(wIrValSr = "", "", wIrValSr + ",") + CStr(wIrValRec.FldValue("IrSr"))
          .MoveNext
        Loop
      End With
      If MsgBox("Invoice Serial(s) (" + wIrValSr + ") have Raw Materials with Zero Values. Do you want to Continue? ", vbYesNo) = vbNo Then _
          pr_Cancel = True: pr_ErrMsg = "Cannot Continue.": Exit Sub
    End If
  
  End If
End Sub
Private Sub ADC_SetRecSource()
  '*** Since this is an 'ABC' Entry, there is no RecSource associated with ADC. So set the
      'DataMode property of ADC control as xdatUnbound
  ADC.DataMode = xdatUnbound
End Sub
Private Sub GrdID_SetRecSource()
  '*** Set the Record Source of the Grid GrdID (Table = InvDsg)
  
  '****** Sachin 3.02 28-12-07 - Corresponding Partition
  Dim wInCnd As String
  
  '****** Sachin 3.02.0 ******
  '6.1_37 where conditions for scope fields
  ADC("wScDmCdFr").CmpStr = "DmCd>= "
  ADC("wScDmCdTo").CmpStr = "DmCd<= "
  ADC("wScDmCtgFr").CmpStr = "DmCtg>= "
  ADC("wScDmCtgTo").CmpStr = "DmCtg<= "
  ADC("wScIdSrFr").CmpStr = "IdSr>= "
  ADC("wScIdSrTo").CmpStr = "IdSr<= "
  wInCnd = ADC.RepCond
  wInCnd = IIF(wInCnd <> "", " and " + wInCnd, "")
  wInCnd = IIF(wInCnd = "", "", " and exists(Select 'x' from DsgMst where DmCd= IdDmCd " + wInCnd + ")")
        
  wInCnd = wInCnd + IIF(gs_Partition = ctCurrPrtn, " And Exists (Select 'x' From InvHd Where InPrtKey=IdPrtKey And InCoCd=IdCoCd And InTc=IdTc And InYy=IdYy And InChr=IdChr And InNo=IdNo And InDt>'" + ms_PurgeDt + "' ) ", "")
  
  GrdID.RecSource = " Select * from InvDsg " + _
                    " Where IdCoCd= '" + gs_CoCd + "'" + _
                    " And IdTc = '" + ADC("wIdTc") + "'" + _
                    " And IdYy = '" + ADC("wIdYy") + "'" + _
                    " And IdChr = '" + ADC("wIdChr") + "' " + _
                    " And IdNo = " + CStr(ADC("wIdNo")) + " And IdPrtKey=(Case When '" + gs_Partition + "'='" + ctCurrPrtn + "' Then '" + ctCurrPrtn + "' Else IdPrtKey End) " + _
                    wInCnd + " Order By IdCoCd, IdTc, IdYy, IdChr, IdNo, IdSr"
                    '"Order By IdSr"
       
  '6.1_37 where conditions for scope fields
  ADC("wScDmCdFr").CmpStr = ""
  ADC("wScDmCdTo").CmpStr = ""
  ADC("wScDmCtgFr").CmpStr = ""
  ADC("wScDmCtgTo").CmpStr = ""
  ADC("wScIdSrFr").CmpStr = ""
  ADC("wScIdSrTo").CmpStr = ""
                    
End Sub
Private Sub GrdIR_SetRecSource()
  '*** Set the Record Source of the Grid GrdIR (Table = InvRm)
  '****** Sachin 3.02 28-12-07 - Corresponding Partition
  GrdIR.RecSource = " Select * from InvRm " + _
                    " Where IrCoCd= '" + gs_CoCd + "'" + _
                    " And IrTc ='" + ADC("wIdTc") + "'" + _
                    " And IrYy = '" + ADC("wIdYy") + "'" + _
                    " And IrChr ='" + ADC("wIdChr") + "' " + _
                    " And IrNo =" + CStr(ADC("wIdNo")) + _
                    " And IrSr =" + CStr(GrdID.Value(GrdID.Row, "IdSr")) + _
                    " And IrPrtKey ='" + CStr(GrdID.Value(GrdID.Row, "IdPrtKey")) + "' " + _
                    " Order By IrCoCd, IrTc, IrYy, IrChr, IrNo, IrSr, IrSrNo"
                    '" Order By IrSrNo"
End Sub
Private Sub GrdIL_SetRecSource()
  '*** Set the Record Source of the Grid GrdIL (Table = InvLab)
  '****** Sachin 3.02 28-12-07 - Corresponding Partition
  GrdIL.RecSource = " Select * from InvLab " + _
                    " Where IlCoCd= '" + gs_CoCd + "'" + _
                    " And IlTc ='" + ADC("wIdTc") + "'" + _
                    " And IlYy = '" + ADC("wIdYy") + "'" + _
                    " And IlChr ='" + ADC("wIdChr") + "' " + _
                    " And IlNo =" + CStr(ADC("wIdNo")) + _
                    " And IlSr =" + CStr(GrdID.Value(GrdID.Row, "IdSr")) + _
                    " And IlPrtKey ='" + CStr(GrdID.Value(GrdID.Row, "IdPrtKey")) + "' " + _
                    " Order By IlCoCd, IlTc, IlYy, IlChr, IlNo, IlSr, IlSrNo"
                    
                    '" Order By IlSrNo"
End Sub
Private Sub GrdIB_SetRecSource()
  '*** Set the Record Source of the Grid GrdIB (Table = InvFgd)
  '****** Sachin 3.02 28-12-07 - Corresponding Partition
  GrdIB.RecSource = " Select * from InvFgd " + _
                    " Where IfCoCd= '" + gs_CoCd + "'" + _
                    " And IfTc ='" + ADC("wIdTc") + "'" + _
                    " And IfYy = '" + ADC("wIdYy") + "'" + _
                    " And IfChr ='" + ADC("wIdChr") + "' " + _
                    " And IfNo =" + CStr(ADC("wIdNo")) + _
                    " And IfSr =" + CStr(GrdID.Value(GrdID.Row, "IdSr")) + _
                    " And IfPrtKey ='" + CStr(GrdID.Value(GrdID.Row, "IdPrtKey")) + "' " + _
                    " And IfSr> 0 " + _
                    " Order By IfCoCd, IfTc, IfYy, IfChr, IfNo, IfSr, IfSrNo"
                    
                    '" Order By IfSrNo"
End Sub
Private Sub TabSumm_Click(PreviousTab As Integer)
  '*** Whenever a tab is selected enable the frame on the clicked tab and disable the previous tab frame
  '*** Initialise the group option fields (working fields) in the current tab from the Head table
  '*** Call the Summary sub procedures associated with the current tab
  FraTabSumm(TabSumm.Tab).Enabled = True
  FraTabSumm(PreviousTab).Enabled = False
  Select Case TabSumm.Tab
  Case Is = 0
    Call InitSummFlds("wDRGrp", "hOrdRmGr")
    Call DsgRmSumm
  Case Is = 1
    Call InitSummFlds("wDLGrp", "hOrdLabGr")
    Call DsgLabSumm
  Case Is = 2
    Call InitSummFlds("wIDGrp", "hOrdDsgGr")
    Call InvDsgSumm
  Case Is = 3
    Call InitSummFlds("wIRGrp", "hOrdRmGr")
    Call InvRmSumm
  Case Is = 4
    Call InitSummFlds("wILGrp", "hOrdLabGr")
    Call InvLabSumm
  End Select
End Sub
Private Sub Pic3D_DblClick()
  '*** If wPic3DDblClk flag option is False then Enlarge the 3-Dimensional picture to fit the FraNKeyAll Frame
  '*** If wPic3DDblClk flag option is True then bring the 3-Dimensional picture to the original pic control size
  If GrdID.Rows - 1 = 0 Then Exit Sub
  Dim w3DPth As String
  If wPic3DDblClk = False Then
    '*** Enlarge The Picture ***
    ' **** Zubin 211 **** '
    '*** Jay 3.2.0 [OdPicNm]
    
    w3DPth = GetPictPath("DM", False, GrdID.Value(GrdID.Row, "IdDmCd"), GrdID.Value(GrdID.Row, "IdDmSz"), gs_CoCd, GrdID.Value(GrdID.Row, "IdExpOdTc"), GrdID.Value(GrdID.Row, "IdExpOdYy"), GrdID.Value(GrdID.Row, "IdExpOdChr"), GrdID.Value(GrdID.Row, "IdExpOdNo"), GrdID.Value(GrdID.Row, "IdExpOdSr"))
    'w3DPth = GetPictPath("DM", False, GrdID.Value(GrdID.Row, "IdDmCd"))
    'w3DPth = GetPictPath("DM", GrdID.Value(GrdID.Row, "IdDmCd"), , False)
    ' **** Zubin 211 **** '
    '*** Jay 3.2.0 [OdPicNm]
    Call ShowLargePic(IIF(FraNKeyAll.Height - 195 > 0, FraNKeyAll.Height - 195, 0), FraNKeyAll.Width, w3DPth, Pic3D)
   
    FraPic3D.Left = 0: FraPic3D.Top = 195
    FraPic3D.Height = Pic3D.Height + 100: FraPic3D.Width = Pic3D.Width + 100
    FraNKeyAll.ZOrder
    FraPic3D.ZOrder (0)
  Else
    '*** Make The Picture Small ***
    '*** Jay 3.2.0 [OdPicNm]
    w3DPth = GetPictPath("DM", False, GrdID.Value(GrdID.Row, "IdDmCd"), GrdID.Value(GrdID.Row, "IdDmSz"), gs_CoCd, GrdID.Value(GrdID.Row, "IdExpOdTc"), GrdID.Value(GrdID.Row, "IdExpOdYy"), GrdID.Value(GrdID.Row, "IdExpOdChr"), GrdID.Value(GrdID.Row, "IdExpOdNo"), GrdID.Value(GrdID.Row, "IdExpOdSr"))
    ' **** Zubin 211 **** '
    'w3DPth = GetPictPath("DM", False, GrdID.Value(GrdID.Row, "IdDmCd"))
    'w3DPth = GetPictPath("DM", GrdID.Value(GrdID.Row, "IdDmCd"), , False)
    ' **** Zubin 211 **** '
    '*** Jay 3.2.0 [OdPicNm]
    Call ShowPic(wPic3DHt, wPic3DWdt, w3DPth, Pic3D)
    FraPic3D.Left = wFraPic3DLft: FraPic3D.Top = wFraPic3DTop
    FraPic3D.Height = Pic3D.Height + 100: FraPic3D.Width = Pic3D.Width + 100
    FraNKeyAll.ZOrder
  End If
  If wPic3DDblClk = False Then
    wPic3DDblClk = True
  Else
    wPic3DDblClk = False
  End If
End Sub
Private Sub PicLD_DblClick()
  '*** If wPicHskDblClk flag option is False then Enlarge the Line Drawing to fit the FraNKeyAll Frame
  '*** If wPicHskDblClk flag option is True then bring the Line Drawing to the original pic control size
  If GrdID.Rows - 1 = 0 Then Exit Sub
  Dim wHskPth As String
  If wPicHskDblClk = False Then
    '*** Enlarge The Picture ***
    '*** Jay 3.2.0 [OdPicNm]
    wHskPth = GetPictPath("DM", True, GrdID.Value(GrdID.Row, "IdDmCd"), GrdID.Value(GrdID.Row, "IdDmSz"), gs_CoCd, GrdID.Value(GrdID.Row, "IdExpOdTc"), GrdID.Value(GrdID.Row, "IdExpOdYy"), GrdID.Value(GrdID.Row, "IdExpOdChr"), GrdID.Value(GrdID.Row, "IdExpOdNo"), GrdID.Value(GrdID.Row, "IdExpOdSr"))
    ' **** Zubin 211 **** '
    'wHskPth = GetPictPath("DM", True, GrdID.Value(GrdID.Row, "IdDmCd"))
    'wHskPth = GetPictPath("DM", GrdID.Value(GrdID.Row, "IdDmCd"), , True)
    ' **** Zubin 211 **** '
    '*** Jay 3.2.0 [OdPicNm]
    Call ShowLargePic(IIF(FraNKeyAll.Height - 195 > 0, FraNKeyAll.Height - 195, 0), FraNKeyAll.Width, wHskPth, PicLD)
    FraPicLD.Left = 0: FraPicLD.Top = 195
    FraPicLD.Height = PicLD.Height + 100: FraPicLD.Width = PicLD.Width + 100
    FraNKeyAll.ZOrder
    FraPicLD.ZOrder (0)
  Else
    '*** Make The Picture Small ***
    '*** Jay 3.2.0 [OdPicNm]
    wHskPth = GetPictPath("DM", True, GrdID.Value(GrdID.Row, "IdDmCd"), GrdID.Value(GrdID.Row, "IdDmSz"), gs_CoCd, GrdID.Value(GrdID.Row, "IdExpOdTc"), GrdID.Value(GrdID.Row, "IdExpOdYy"), GrdID.Value(GrdID.Row, "IdExpOdChr"), GrdID.Value(GrdID.Row, "IdExpOdNo"), GrdID.Value(GrdID.Row, "IdExpOdSr"))
    ' **** Zubin 211 **** '
    'wHskPth = GetPictPath("DM", True, GrdID.Value(GrdID.Row, "IdDmCd"))
    'wHskPth = GetPictPath("DM", GrdID.Value(GrdID.Row, "IdDmCd"), , True)
    ' **** Zubin 211 **** '
    '*** Jay 3.2.0 [OdPicNm]
    Call ShowPic(wPicHskHt, wPicHskWdt, wHskPth, PicLD)
    FraPicLD.Left = wFraPicHskLft: FraPicLD.Top = wFraPicHskTop
    FraPicLD.Height = PicLD.Height + 100: FraPicLD.Width = PicLD.Width + 100
    FraNKeyAll.ZOrder
  End If
  If wPicHskDblClk = False Then
    wPicHskDblClk = True
  Else
    wPicHskDblClk = False
  End If
End Sub
Private Sub CmdSumm_Click()
  '*** Display or hide the Summary frame depending on the toggle state
  '*** Call the Design Raw Material Summary sub procedure (as it is the first summary tab)
  Call DispFra(Summ)
  If FraSumm.Visible = True Then
    Call InitSummFlds("wDRGrp", "hOrdRmGr")
    Call DsgRmSumm
  End If
End Sub
Private Sub CmdDD_Click()
  '*** Display or hide the Design Details frame depending on the toggle state
  '*** The CausesValidation property of this command button is false. The purpose of this is
  '*** when the Design Details option is selected, the grid GrdID should not be fully
      'validated i.e. only the validation of the column on which the cursor was previously
      'standing should take place and the row itself should not be saved. When the user exits
      'out of this option, the cursor returns to the same row.
  '*** All the fields which could not fit into the grid GrdID due to space constraint
      'is displayed in the frame associated with this command button
    '*** InitialiseFld is used to bypass the validation for a field while assigning values to it
    '*** For example if a field does not accept blank value and the default value being
        'assigned through the code is blank, then InitialiseFld stops error message
        'from being generated
  Dim wRow As Integer
  If FraDD.Visible = False Then
    If GrdID.HalfValidate Then
      GrdID.SetFocus
      Exit Sub
    End If
    Call DispFra(DsgDet)
    
    wRow = GrdID.Row
    ADC("wDdIdExpCmCd").InitialiseFld GrdID.Value(wRow, "IdExpCmCd")
    ADC("wDdIdIVal").InitialiseFld GrdID.Value(wRow, "IdIVal")
    ADC("wDdIdZVal").InitialiseFld GrdID.Value(wRow, "IdZVal")
    ADC("wDdIdIAdd").InitialiseFld GrdID.Value(wRow, "IdIAdd")
    ADC("wDdIdZAdd").InitialiseFld GrdID.Value(wRow, "IdZAdd")
    ADC("wDdIdVaCtg").InitialiseFld GrdID.Value(wRow, "IdVaCtg")
    ADC("wDdIdKt").InitialiseFld GrdID.Value(wRow, "IdKt")
    ADC("wDdIdGldLs").InitialiseFld GrdID.Value(wRow, "IdGldLs")
    
    ADC("wDdIdFixPrc").InitialiseFld GrdID.Value(wRow, "IdFixPrc")
    
    '****** Zubin 3.02.01 - 27/03/08
    ADC("wDdIdLabWtFrOrd").InitialiseFld GrdID.Value(wRow, "IdLabWtFrOrd")
    '****** Zubin 3.02.01 - 27/03/08

    ADC("wDdIdHZValEqInv").InitialiseFld GrdID.Value(wRow, "IdHZValEqInv")
    ADC("wDdIdGldAs").InitialiseFld GrdID.Value(wRow, "IdGldAs")
    ADC("wDdIdCustmMtchWt").InitialiseFld GrdID.Value(wRow, "IdCustmMtchWt")    '*** (Jen 2.13)
    
    ADC("wDdIdiGldAsWt").InitialiseFld GrdID.Value(wRow, "IdiGldAsWt")
    ADC("wDdIdZGldAsWt").InitialiseFld GrdID.Value(wRow, "IdZGldAsWt")
    
    '****** Sachin 2.13.0 - 26-05-2006 ******
    ADC("wDdIdLabAs").InitialiseFld GrdID.Value(wRow, "IdLabAs")
    ADC("wDdIdILabAsWt").InitialiseFld GrdID.Value(wRow, "IdILabAsWt")
    ADC("wDdIdZLabAsWt").InitialiseFld GrdID.Value(wRow, "IdZLabAsWt")
    '****** Sachin 2.13.0 - 26-05-2006 ******
    
    ADC("wDDIdiFOBPerPc").InitialiseFld MWLib.Div(GrdID.Value(wRow, "IdIFob"), GrdID.Value(wRow, "IdQty"))
    ADC("wDDIdZFOBPerPc").InitialiseFld MWLib.Div(GrdID.Value(wRow, "IdZFob"), GrdID.Value(wRow, "IdQty"))
    '****************Geeta**********Emr207
    ADC("wCalcIdiFOBPerPc").InitialiseFld MWLib.Div(MWLib.Div(GrdID.Value(wRow, "IdIFob"), (1 + ADC("wIdFobPer") / 100)), GrdID.Value(wRow, "IdQty"))
    ADC("wCalcIdZFOBPerPc").InitialiseFld MWLib.Div(MWLib.Div(GrdID.Value(wRow, "IdZFob"), (1 + ADC("wIdFobPer") / 100)), GrdID.Value(wRow, "IdQty"))
    '****************Geeta**********Emr207
    '****************Geeta**********Emr208
    ADC("wDdIdValAddn").InitialiseFld GrdID.Value(wRow, "IdValAddn")
     '***************
    '****** Sachin 3.02 28-12-07 - Corresponding Partition
    ADC("wDdIdDmCtg").InitialiseFld moCn.GetFldVal("Select DmCtg from DsgMst where " + _
                                    "DmTcTyp= 'DM' and DmCd= '" + GrdID.Value(wRow, "IdDmCd") + "' " + _
                                    "and DmSz= '' and DmPrtKEy=(Case When '" + ADC("wIdPrtKey") + "'='" + ctCurrPrtn + "' Then '" + ctCurrPrtn + "' Else DmPrtKey End)")
    ADC("wDdIdDmCol").InitialiseFld moCn.GetFldVal("Select OdDmCol from OrdDsg where " + _
                                    "OdCoCd= '" + gs_CoCd + "' and " + _
                                    "OdTc= '" + GrdID.Value(wRow, "IdExpOdTc") + "' and " + _
                                    "OdYy= '" + GrdID.Value(wRow, "IdExpOdYy") + "' and " + _
                                    "OdChr= '" + GrdID.Value(wRow, "IdExpOdChr") + "' and " + _
                                    "OdNo= " + CStr(GrdID.Value(wRow, "IdExpOdNo")) + " and " + _
                                    "OdSr= " + CStr(GrdID.Value(wRow, "IdExpOdSr")) + " and OdPrtKEy=(Case When '" + ADC("wIdPrtKey") + "'='" + ctCurrPrtn + "' Then '" + ctCurrPrtn + "' Else OdPrtKey End)")
                                    
        
    '3.11.0
    If ms_RmRtfrRtChrt = "N" Then
        ADC("wDDIdLMERtFrOrd").InitialiseFld moCn.GetFldVal("Select OdIGldRtEqOrd from OrdDsg where " + _
                                        "OdCoCd= '" + gs_CoCd + "' and " + _
                                        "OdTc= '" + GrdID.Value(wRow, "IdExpOdTc") + "' and " + _
                                        "OdYy= '" + GrdID.Value(wRow, "IdExpOdYy") + "' and " + _
                                        "OdChr= '" + GrdID.Value(wRow, "IdExpOdChr") + "' and " + _
                                        "OdNo= " + CStr(GrdID.Value(wRow, "IdExpOdNo")) + " and " + _
                                        "OdSr= " + CStr(GrdID.Value(wRow, "IdExpOdSr")) + " and OdPrtKEy=(Case When '" + ADC("wIdPrtKey") + "'='" + ctCurrPrtn + "' Then '" + ctCurrPrtn + "' Else OdPrtKey End)")
    Else
        ADC("wDDIdLMERtFrOrd").InitialiseFld ms_RmRtfrRtChrt
    End If
    
  Else
    Call DispFra(DsgDet)
    wRow = GrdID.Row
    GrdID.Value(wRow, "IdExpCmCd") = ADC("wDdIdExpCmCd")
    'GrdID.Value(wRow, "IdIVal") = ADC("wDdIdIVal")
    'GrdID.Value(wRow, "IdZVal") = ADC("wDdIdZVal")
    GrdID.Value(wRow, "IdIAdd") = ADC("wDdIdIAdd")
    GrdID.Value(wRow, "IdZAdd") = ADC("wDdIdZAdd")
    GrdID.Value(wRow, "IdVaCtg") = ADC("wDdIdVaCtg")
    GrdID.Value(wRow, "IdKt") = ADC("wDdIdKt")
    GrdID.Value(wRow, "IdGldLs") = ADC("wDdIdGldLs")
    GrdID.Value(wRow, "IdCustmMtchWt") = ADC("wDdIdCustmMtchWt")    '*** (Jen 2.13)
    GrdID.Value(wRow, "IdFixPrc") = ADC("wDdIdFixPrc")
    
    GrdID.Value(wRow, "IdHZValEqInv") = ADC("wDdIdHZValEqInv")
    GrdID.Value(wRow, "IdGldAs") = ADC("wDdIdGldAs")
    GrdID.Value(wRow, "IdLabAs") = ADC("wDdIdLabAs")
    'GrdID.Value(wRow, "IdiGldAsWt") = ADC("wDdIdiGldAsWt")
    'GrdID.Value(wRow, "IdZGldAsWt") = ADC("wDdIdZGldAsWt")
    '********Geeta*****Emr208
    GrdID.Value(wRow, "IdValAddn") = ADC("wDdIdValAddn")
    '********Geeta***
  End If
End Sub
Private Sub CmdIB_Click()
  '*** Display or hide the Bag frame depending on the toggle state
  Call DispFra(InvBag)
End Sub
Private Sub CmdIL_Click()
  '*** Display or hide the Labour frame depending on the toggle state
  Call DispFra(InvLab)
End Sub
Private Sub CmdMI_Click()
  '*** Call the Make Invoice procedure
  '****** Sachin 3.02 - Locking the Previous Partition Records
  If ADC("WIDPRTKEY") <> ctCurrPrtn Then DispMsg "Cannot Make Invoice In Previous Partition", etError: Exit Sub
  
  If mf_DtLocked = True Then Exit Sub
  
  '*** (Jen 3.01.0_1)
  If mb_DsgOnHld = True Then
    DispMsg "Cannot Make Invoice As Designs On Hold For The Invoice", etError: ADC.SetFocus: Exit Sub
  End If
  '*** (Jen 3.01.0_1)
  Call MakeInvoice
  GrdIR.Visible = True
  GrdIB.Visible = True
  GrdIL.Visible = True
End Sub
Private Sub CmdCRL_Click()
  '*** On the click of this button the write Rm & Lab records (if there are no records already)
  '*** Also this copy is possible only if bag records exist
  '*** Call the MakeRm and MakeLab procedure
  
  
  '****** When the User clicks on the 'Copy Rm/Lab' option the 'Valid' event of the grid
         'GrdID disables all the command buttons associated with it. So the cursor
         'remains in the grid GrdID but the Grid 'GotFocus' event does not happen. So the
         'command buttons associated with the grid remain disabled cwhich causes a problem
         'in making the entry. Hence as soon as the user clicks on 'Copy Rm/Lab' the focus
         'is sent to the ADC. After the copying, the focus is sent back to the grid.
         'This is done only for Command Buttons associated with the grid and which do not
         'have frames associated with them
  ADC.SetFocus
  
  '****** Sachin 3.02 28-12-07
  If ADC("wIdPrtKEy") <> ctCurrPrtn Then DispMsg "Cannot Copy Rm/Lab in Previous Partition Invoice", etError: Exit Sub
  
  If mf_DtLocked = True Then Exit Sub
  
  Dim wOdIWtEqOrd As String, wOdIWtFrOrd As String, wOdIGldRtEqOrd As String
  Dim wArrStr As String, wArr() As String
    
  If GrdIR.Rows > 1 Then DispMsg "Cannot Copy As Raw Material Details Already Exist", etError: GrdID.SetFocus: Exit Sub
  If GrdIL.Rows > 1 Then DispMsg "Cannot Copy As Labour Details Already Exist", etError: GrdID.SetFocus: Exit Sub
  If GrdIB.Rows <= 1 Then DispMsg "Cannot Copy As Bag Does Not Exist For The Design", etError: GrdID.SetFocus: Exit Sub
  
  '*** (Jen 3.2.0)
  '****** Sachin 3.02.0 28-12-07 - prtkey added below
  If moCn.RecSeek("Select 'x' from InvZRm where IzCoCd= '" + gs_CoCd + "' and IzTc= '" + ADC("wIdTc") + "' and " + _
                  "IzYy=  '" + ADC("wIdYy") + "' and IzChr=  '" + ADC("wIdChr") + "' and " + _
                  "IzNo=  " + CStr(ADC("wIdNo")) + " and IzPrtKey= '" + ctCurrPrtn + "' ") Then
    DispMsg "Cannot Copy Rm/ Labour As Custom Matching Has Been Done", etError: Exit Sub
  End If
  
  With GrdID
    wArrStr = moCn.GetFldVal("Select OdIWtEqOrd +'|'+ OdIWtFrOrd +'|'+ OdIGldRtEqOrd  from OrdDsg where " + _
                                 "OdCoCd= '" + gs_CoCd + "' and " + _
                                 "OdTc= '" + .Value(GrdID.Row, "IdExpOdTc") + "' and " + _
                                 "OdYy= '" + .Value(GrdID.Row, "IdExpOdYy") + "' and " + _
                                 "OdChr= '" + .Value(GrdID.Row, "IdExpOdChr") + "' and " + _
                                 "OdNo= " + CStr(.Value(GrdID.Row, "IdExpOdNo")) + " and " + _
                                 "OdSr= " + CStr(.Value(GrdID.Row, "IdExpOdSr")) + " and " + _
                                 "OdPrtKey= '" + ctCurrPrtn + "' ")
    wArr = Split(wArrStr, "|")
    If UBound(wArr) = 2 Then
      wOdIWtEqOrd = wArr(0): wOdIWtFrOrd = wArr(1): wOdIGldRtEqOrd = wArr(2)
    End If
    
 '   GrdIR.StartCopy
    '*** Copy Raw Material Dets
    If moCn.GetFldVal("Select hSizeStkRt from Head where HCoCd= '" + gs_CoCd + "' and " + _
                      "HCd= '" + ctSelfCmCd + "'") = "N" Then
      '6.1-55 IIF(ms_RtFrInvCust = "Y", ADC("wIdCmCd") added
      Call MakeRm(.Value(GrdID.Row, "IdSr"), .Value(GrdID.Row, "IdExpOdTc"), _
                  .Value(GrdID.Row, "IdExpOdYy"), .Value(GrdID.Row, "IdExpOdChr"), _
                  .Value(GrdID.Row, "IdExpOdNo"), .Value(GrdID.Row, "IdExpOdSr"), _
                  IIF(ms_RtFrInvCust = "Y", ADC("wIdCmCd"), .Value(GrdID.Row, "IdExpCmCd")), .Value(GrdID.Row, "IdDmCd"), _
                  .Value(GrdID.Row, "IdQty"), wOdIWtEqOrd, wOdIWtFrOrd, wOdIGldRtEqOrd)
    Else
      Call MakeRmFromFg(wOdIWtEqOrd, wOdIWtFrOrd, wOdIGldRtEqOrd, .Value(GrdID.Row, "IdQty"))
    End If
 '   GrdIR.EndCopy
'    GrdIL.StartCopy
    '*** Copy Lab Dets
     Call MakeLab(.Value(GrdID.Row, "IdExpOdTc"), .Value(GrdID.Row, "IdExpOdYy"), _
               .Value(GrdID.Row, "IdExpOdChr"), .Value(GrdID.Row, "IdExpOdNo"), _
               .Value(GrdID.Row, "IdExpOdSr"))
 '    GrdIL.EndCopy
     
     Call CalcGrdIdVals(GrdID.Row)

      GrdID.SaveRec GrdID.Row
  End With
  mf_StopRecalcDuringCopyIR = False
  mf_StopRecalcDuringCopyIL = False
  GrdID.SetFocus
End Sub
Private Sub CmdDC_Click()
  '*** Display or hide the Design Copy frame depending on the toggle state
  '*** The CausesValidation property of this command button is false. The purpose of this is
      'when the Design Copy option is selected, the grid GrdID should not be fully
      'validated i.e. only the validation of the column on which the cursor was previously
      'standing should take place and the row itself should not be saved. When the user exits
      'out of this option, the cursor returns to the same row. The reason is some of the
      'key fields can be filled up only during copy. Hence the row should not be saved
      'before copy as it might generate a trigger error

    '****** Sachin 3.02 28-12-07
    If ADC("wIdPrtKey") <> ctCurrPrtn Then DispMsg "Cannot Copy Design In Previous Partition Invoice ", etError: Exit Sub

  '*** (Jen 3.01.0_1)
  If mb_DsgOnHld = True Then
    DispMsg "Cannot Copy Design As Designs On Hold For The Invoice", etError: ADC.SetFocus: Exit Sub
  End If
  '*** (Jen 3.01.0_1)

  '****** Sachin 2.13.0 - Delete RM/Lab Button - 27-05-2006 ******
  '****** Sachin 3.02 28-12-07 - Current Partition
  If moCn.RecSeek("Select 'x' from InvZRm where IzCoCd= '" + gs_CoCd + "' and IzTc= '" + ADC("wIdTc") + "' and " + _
                  "IzYy=  '" + ADC("wIdYy") + "' and IzChr=  '" + ADC("wIdChr") + "' and " + _
                  "IzNo=  " + CStr(ADC("wIdNo")) + " and IzPrtKey='" + ctCurrPrtn + "' ") Then _
    DispMsg "Cannot Copy Design As Custom Matching Has Been Done", etError: GrdID.SetFocus: Exit Sub
  '****** Sachin 2.13.0 - Delete RM/Lab Button - 27-05-2006 ******

  If mf_DtLocked = True Then Exit Sub
  If FraDC.Visible = False Then
    If GrdID.HalfValidate Then
      GrdID.SetFocus
      Exit Sub
    End If
    Call DispFra(DsgCpy)
'    If ADC("wWaitForAllBags") = "" Then ADC("wWaitForAllBags") = "Y"
  Else
    Call DispFra(DsgCpy)
  End If
End Sub
Private Sub CmdDCGo_Click()
  '*** Make sure that the Export Order No mentioned in the Copy frame is a valid Order
  '*** Copy is not allowed if the Order is closed
  '*** Cannot copy if Bag, Rm or Lab record exists for the Design Sr
  '*** First the Inv Design record is copied
  '*** All the bags in finished goods belonging to the Export Order mentioned in the
      'Design Copy frame are copied in the InvFgd table
  '*** All the Rm records for the Inv Design record are copied
  '*** All the Lab records for the Inv Design record are copied
  '*** All this is done via the procedure MakeDsgWithBags

  Dim wRsSrcDsg As MDORowSet
  ''''If GrdIR.Rows > 1 Then DispMsg "Cannot Copy As Raw Material Details Already Exist", etError: Exit Sub
  ''''If GrdIL.Rows > 1 Then DispMsg "Cannot Copy As Labour Details Already Exist", etError: Exit Sub
  ''''If GrdIB.Rows > 1 Then DispMsg "Cannot Copy As Bag Details Already Exist", etError: Exit Sub
  
'  If ADC("wWaitForAllBags") = "" Then DispMsg "Cannot Copy As 'Wait For All Bags' Option Is Blank", etError: Exit Sub
  'New
  Dim wCnd As String, wAllBagsCnd As String
  
  Dim ws_SqlStr As String   '*** (Jen 3.01.0_1)
  
  ' **** Manali 3.8.0 - Invoice Sort By
  Dim ws_OrdBy As String
  ws_OrdBy = moCn.GetFldVal("Select PDesc225 From Param Where PTyp='INSRT' and PMCd='" + ADC("wInSrt") + "' and PSCd=''")
  ws_OrdBy = ws_OrdBy + IIF(ws_OrdBy <> "", ", ", "") + " OdCoCd, OdTc, OdYy, OdChr, OdNo, OdSr"
  ws_OrdBy = " Order By " + ws_OrdBy
  ' **** Manali 3.8.0 - Invoice Sort By
  
  
  '****** Sachin 3.3.0 (23-06-08) - [42. Make Invoice By Bag] ******
  ADC("wIdInByBag") = moCn.GetFldVal(" Select InByBagYN From InvHd " + _
                                  " Where InCoCd= '" + gs_CoCd + "' and " + _
                                  " InTc= '" + ADC("wIdTc") + "' and " + _
                                  " InYy= '" + ADC("wIdYy") + "' and " + _
                                  " InChr= '" + ADC("wIdChr") + "' and " + _
                                  " InNo= " + CStr(ADC("wIdNo")) + " and InPrtKey=(Case When '" + gs_Partition + "'='" + ctCurrPrtn + "' Then '" + ctCurrPrtn + "' Else InPrtKey End) ")
  mb_InByBag = IIF(ADC("wIdInByBag") = "Y", True, False)
  '****** Sachin 3.3.0 (23-06-08) - [42. Make Invoice By Bag] ******
  
  '*** Checking if copy is possible
  If ADC("wDcOdTcFr") = "" And ADC("wDcOdYyFr") = "" And ADC("wDcOdChrFr") = "" And _
     ADC("wDcOdNoFr") = 0 And ADC("wDcOdSrFr") = 0 And ADC("wDcOdTcTo") = "" And _
     ADC("wDcOdYyTo") = "" And ADC("wDcOdChrTo") = "" And ADC("wDcOdNoTo") = 0 And _
     ADC("wDcOdSrTo") = 0 Then _
     DispMsg "Enter Atleast One Selection Parameter For Copying", etError: Exit Sub
  
  ADC("wDcOdTcFr").CmpStr = "OdTc >= "
  ADC("wDcOdYyFr").CmpStr = "OdYy >= "
  ADC("wDcOdChrFr").CmpStr = "OdChr >= "
  ADC("wDcOdNoFr").CmpStr = "OdNo >= "
  ADC("wDcOdSrFr").CmpStr = "OdSr >= "
  ADC("wDcOdTcTo").CmpStr = "OdTc <= "
  ADC("wDcOdYyTo").CmpStr = "OdYy <= "
  ADC("wDcOdChrTo").CmpStr = "OdChr <= "
  ADC("wDcOdNoTo").CmpStr = "OdNo <= "
  ADC("wDcOdSrTo").CmpStr = "OdSr <= "
  
  If UCase(ms_IdTcTyp) = UCase(ctInTcTyp) Then
 '   If ADC("wWaitForAllBags") = "Y" Then
 '     wAllBagsCnd = " having Str(max(OdOrdQty), 16, 1)= Str(max(OdExpQty)+ sum(IsNull(IfQty, FdQty- FdDespQty)), 16, 1) "
 '   ElseIf ADC("wWaitForAllBags") = "N" Then
 '     wAllBagsCnd = " having Round(max(OdOrdQty), 1)>= Round(max(OdExpQty)+ sum(IsNull(IfQty, FdQty- FdDespQty)), 1) "
 '   End If
        wAllBagsCnd = " having ((max(OdInvAllBags) = 'Y' and Str(max(OdOrdQty), 16, 1)= Str(max(OdExpQty)+ sum(IsNull(IfQty, FdQty- FdDespQty)), 16, 1)) Or (max(OdInvAllBags) = 'N' and Round(max(OdOrdQty), 1)>= Round(max(OdExpQty)+ sum(IsNull(IfQty, FdQty- FdDespQty)), 1)) ) "
  ElseIf UCase(ms_IdTcTyp) = UCase(ctCTBTcTyp) Then
    wAllBagsCnd = " "
  End If
  
  wCnd = ADC.RepCond
  wCnd = IIF(wCnd <> "", " and " + wCnd, "")
  
  'New
  '********Geeta*************Emr206***************09/01/04******included max(OdHld)*****
  '****** (Jen 3.01.0_1 added OrdRm join for MainMet) ******
  '****** Sachin 3.02 - Id fields In Joins
  '**Geeta**3.02 Patch****OdILabWtFrOrd,OdLabAsWt
  ' **** Manali 3.8.0 - ws_OrdBy added, DsgMst Join added
  ws_SqlStr = "Select Max(OdDmCd) as qDmCd, Max(OdSfx) as qSfx,Max(OdDmSz) as qDmSz, sum(IsNull(IfQty, FdQty- " + ms_DespTyp + ")) as qQty, " + _
               "OdTc as qOdTc, OdYy as qOdYy, OdChr as qOdChr, OdNo as qOdNo, " + _
               "OdSr as qOdSr, " + IIF(mb_InByBag = True, " B.FdBYy as qBYy, B.FdBChr as qBChr, B.FdBNo as qBNo,", "") + _
               " Max(OmCmCd) as qCmCd, Max(OdVaCtg) as qVaCtg, Max(OdKt) as qKt,max(OdILabWtFrOrd) as OdILabWtFrOrd,max(OdInvAllBags) as OdInvAllBags,max(OdLabAsWt) as OdLabAsWt, " + _
               "Max(OdFixPrc) as qFixPrc, Max(OdGldAs) as qGldAs, Max(OdLabAs) as qLabAs, Max(OdCls) as qOdCls, " + _
               "max(OdIWtEqOrd) as OdIWtEqOrd, max(OdIWtFrOrd) as OdIWtFrOrd, max(OdIGldRtEqOrd) as OdIGldRtEqOrd, " + _
               "max(OdHld) as qOdHld, max(IsNull(OrRmCtg, '')) as qMainMetRmCtg " + _
               "from OrdMst join OrdDsg on OdOmIdNo=OmIdNo and OdCoCd= OmCoCd and OdTc= OmTc and OdYy= OmYy and OdChr= OmChr and OdNo= OmNo and OdPrtKey=OmPrtKey " + _
               "Join DsgMst On DmTcTyp='DM' and DmCd=OdDmCd and DmSz='' " + _
               "join Fgd B on /* Fd" + ms_OrdTyp + "OdIdNo=OdIdNo and */ FdCoCd= OdCoCd and " + _
               "Fd" + ms_OrdTyp + "OdTc= OdTc and Fd" + ms_OrdTyp + "OdYy= OdYy and Fd" + ms_OrdTyp + "OdChr= OdChr and " + _
               "Fd" + ms_OrdTyp + "OdNo= OdNo and Fd" + ms_OrdTyp + "OdSr= OdSr and FdPrtKey=OdPrtKey " + _
               "left outer join InvFgd on IfFdIdNo=FdIdNo and IfCoCd= FdCoCd and IfFdTc= FdTc and IfFdYy= FdYy and IfFdChr= FdChr " + _
               "and IfFdNo= FdNo and IfFdSr= FdSr and IfPrtKey=FdPrtKey and " + _
               "IsNull(IfCoCd, '" + gs_CoCd + "')= '" + gs_CoCd + "' and " + _
               "IsNull(IfTc, '" + ADC("wIdTc") + "')= '" + ADC("wIdTc") + "' and " + _
               "IsNull(IfYy, '" + ADC("wIdYy") + "')= '" + ADC("wIdYy") + "' and " + _
               "IsNull(IfChr, '" + ADC("wIdChr") + "')= '" + ADC("wIdChr") + "' and " + _
               "IsNull(IfNo, " + CStr(ADC("wIdNo")) + ")= " + CStr(ADC("wIdNo")) + " and " + _
               "IsNull(IfSr, 0)= 0  " + _
               "left outer join OrdRm on OrOdIdNo=OdIdNo and OrCoCd= OdCoCd and OrTc= OdTc and OrYy= OdYy and OrChr= OdChr " + _
               "and OrNo= OdNo and OrSr= OdSr and OrPrtKey=OdPrtKey and OrMainMet= 'Y' "
               
  ws_SqlStr = ws_SqlStr + " where OdDmTcTyp='DM' and OdCoCd= '" + gs_CoCd + "' " + _
                IIF(UCase(ms_IdTcTyp) = UCase(ctInTcTyp), " and OdCls= 'N' ", " ") + " and OdPrtKey='" + ctCurrPrtn + "' and " + _
                "FdTc in (Select PMCd from Param where PTyp= 'TC' and PSCd= '' and PValue in ('FB', 'FR')) and " + _
                "Round(IsNull(IfQty, FdQty- " + ms_DespTyp + "), 1)> 0 and " + _
                "not exists(Select 'x' from Fgd M where M.FdCoCd= B.FdCoCd and " + _
                "M.FdRefYy= B.FdYy and M.FdRefKey= B.FdKey and M.FdPrtKey=B.FdPrtKey and " + _
                "(Select PValue from Param where PTyp= 'TC' and PMCd= M.FdTc)= 'FWB') " + wCnd + _
                "and OmCmCurCd= '" + ADC("wIdCmCurCd") + "' " + _
                "group by OdCoCd, OdTc, OdYy, OdChr, OdNo, OdSr " + IIF(mb_InByBag = True, ",B.FdBYy, B.FdBChr, B.FdBNo ", " ") + wAllBagsCnd + ws_OrdBy
    
  Set wRsSrcDsg = moCn.OpenRes(ws_SqlStr)
  

  If wRsSrcDsg.RecCount = 0 Then DispMsg "Bags Do Not Exist For The Specified Order Sr. Range/ Order Closed/ All Designs Are Of Different Currency", etError: Exit Sub: Set wRsSrcDsg = Nothing
  '*** (Jen 2.13)
  'wRsSrcDsg.MoveFirst
  'If wRsSrcDsg!qOdCls = "Y" Then DispMsg "Order Is Closed", etError: Exit Sub: Set wRsSrcDsg = Nothing
   GrdID.StartSpecialCopy
   GrdIB.StartSpecialCopy
  GrdIL.StartSpecialCopy
  GrdIR.StartSpecialCopy
  mf_inspecialCopy = True

  MWLib.BeginProcess Me, "Start copy.."
  Call MakeDsgWithBags(False, wRsSrcDsg)
  
  If GrdID.Rows > 1 Then CmdMI.Enabled = False
  GrdID.EndSpecialCopy
   GrdIB.EndSpecialCopy
  GrdIL.EndSpecialCopy
  GrdIR.EndSpecialCopy
  mf_inspecialCopy = False
  Call CmdDC_Click
  MWLib.EndProcess Me
  
  GrdIR.Visible = True
'  GrdIB.Visible = True
'  GrdIL.Visible = True
  Set wRsSrcDsg = Nothing
End Sub
Private Sub CmdDCGo1_Click()
'  '*** Make sure that the Export Order No mentioned in the Copy frame is a valid Order
'  '*** Copy is not allowed if the Order is closed
'  '*** Cannot copy if Bag, Rm or Lab record exists for the Design Sr
'  '*** First the Inv Design record is copied
'  '*** All the bags in finished goods belonging to the Export Order mentioned in the
'      'Design Copy frame are copied in the InvFgd table
'  '*** All the Rm records for the Inv Design record are copied
'  '*** All the Lab records for the Inv Design record are copied
'  '*** All this is done via the procedure MakeDsgWithBags
'
'  Dim wRsSrcDsg As MDORowSet
'  If GrdIR.Rows > 1 Then DispMsg "Cannot Copy As Raw Material Details Already Exist", etError: Exit Sub
'  If GrdIL.Rows > 1 Then DispMsg "Cannot Copy As Labour Details Already Exist", etError: Exit Sub
'  If GrdIB.Rows > 1 Then DispMsg "Cannot Copy As Bag Details Already Exist", etError: Exit Sub
'
'  Set wRsSrcDsg = moCn.OpenRes("Select Max(OdDmCd) as qDmCd, Max(OdSfx) as qSfx, " + _
'                  "Max(OdDmSz) as qDmSz, sum(IfQty) as qQty, OdTc as qExpOdTc, " + _
'                  "OdYy as qExpOdYy, OdChr as qExpOdChr, OdNo as qExpOdNo, " + _
'                  "OdSr as qExpOdSr, Max(OmCmCd) as qExpCmCd, " + _
'                  "Max(OdVaCtg) as qVaCtg, Max(OdKt) as qKt, Max(OdFixPrc) as qFixPrc, " + _
'                  "Max(OdGldAs) as qGldAs, Max(OdCls) as qOdCls, " + _
'                  "max(OdIWtEqOrd) as OdIWtEqOrd, max(OdIWtFrOrd) as OdIWtFrOrd, max(OdIGldRtEqOrd) as OdIGldRtEqOrd " + _
'                  "from OrdMst join OrdDsg on OdCoCd= OmCoCd and OdTc= OmTc and OdYy= OmYy " + _
'                  "and OdChr= OmChr and OdNo= OmNo join Fgd on FdCoCd= OdCoCd and " + _
'                  "FdExpOdTc= OdTc and FdExpOdYy= OdYy and FdExpOdChr= OdChr and " + _
'                  "FdExpOdNo= OdNo and FdExpOdSr= OdSr join InvFgd on " + _
'                  "IfCoCd= FdCoCd and IfFdTc= FdTc and IfFdYy= FdYy and " + _
'                  "IfFdChr= FdChr and IfFdNo= FdNo and IfFdSr= FdSr and IfCoCd= '" + gs_CoCd + "' and " + _
'                  "IfTc= '" + ADC("wIdTc") + "' and IfYy= '" + ADC("wIdYy") + "' and " + _
'                  "IfChr= '" + ADC("wIdChr") + "' and IfNo= " + CStr(ADC("wIdNo")) + " and IfSr= 0 " + _
'                  "where OdCoCd= '" + gs_CoCd + "' and OdTc='" + ADC("wDcOdTc") + "' and " + _
'                  "OdYy= '" + ADC("wDcOdYy") + "' and OdChr='" + ADC("wDcOdChr") + "' and " + _
'                  "OdNo=" + CStr(ADC("wDcOdNo")) + " and OdSr= " + CStr(ADC("wDcOdSr")) + " " + _
'                  "group by  OdTc, OdYy, OdChr, OdNo, OdSr")
'  If wRsSrcDsg.RecCount = 0 Then
'    Set wRsSrcDsg = moCn.OpenRes("Select Max(OdDmCd) as qDmCd, Max(OdSfx) as qSfx, " + _
'                    "Max(OdDmSz) as qDmSz, 0 as qQty, OdTc as qExpOdTc, " + _
'                    "OdYy as qExpOdYy, OdChr as qExpOdChr, OdNo as qExpOdNo, " + _
'                    "OdSr as qExpOdSr, Max(OmCmCd) as qExpCmCd, " + _
'                    "Max(OdVaCtg) as qVaCtg, Max(OdKt) as qKt, Max(OdFixPrc) as qFixPrc, " + _
'                    "Max(OdGldAs) as qGldAs, Max(OdCls) as qOdCls, " + _
'                    "max(OdIWtEqOrd) as OdIWtEqOrd, max(OdIWtFrOrd) as OdIWtFrOrd, max(OdIGldRtEqOrd) as OdIGldRtEqOrd " + _
'                    "from OrdMst join OrdDsg on OdCoCd= OmCoCd and OdTc= OmTc and OdYy= OmYy " + _
'                    "and OdChr= OmChr and OdNo= OmNo left outer join Fgd on FdCoCd= OdCoCd and " + _
'                    "FdExpOdTc= OdTc and FdExpOdYy= OdYy and FdExpOdChr= OdChr and " + _
'                    "FdExpOdNo= OdNo and FdExpOdSr= OdSr and Round(FdQty- FdDespQty, 1)> 0 " + _
'                    "where OdCoCd= '" + gs_CoCd + "' and OdTc='" + ADC("wDcOdTc") + "' and " + _
'                    "OdYy='" + ADC("wDcOdYy") + "' and OdChr='" + ADC("wDcOdChr") + "' and " + _
'                    "OdNo=" + CStr(ADC("wDcOdNo")) + " and OdSr= " + CStr(ADC("wDcOdSr")) + " " + _
'                    "group by OdTc, OdYy, OdChr, OdNo, OdSr")
'  End If
'
''  Set wRsSrcDsg = moCn.OpenRes("Select Max(OdDmCd) as qDmCd, Max(OdSfx) as qSfx, " + _
''                "Max(OdDmSz) as qDmSz, sum(IsNull(IfQty, 0)) as qQty, OdTc as qExpOdTc, " + _
''                "OdYy as qExpOdYy, OdChr as qExpOdChr, OdNo as qExpOdNo, " + _
''                "OdSr as qExpOdSr, Max(OmCmCd) as qExpCmCd, " + _
''                "Max(OdVaCtg) as qVaCtg, Max(OdKt) as qKt, Max(OdFixPrc) as qFixPrc, " + _
''                "Max(OdGldAs) as qGldAs, Max(OdCls) as qOdCls, " + _
''                "max(OdIWtEqOrd) as OdIWtEqOrd, max(OdIWtFrOrd) as OdIWtFrOrd " + _
''                "from OrdMst join OrdDsg on OdTc= OmTc and OdYy= OmYy and OdChr= OmChr and " + _
''                "OdNo= OmNo left outer join Fgd on " + _
''                "FdExpOdTc= OdTc and FdExpOdYy= OdYy and FdExpOdChr= OdChr and " + _
''                "FdExpOdNo= OdNo and FdExpOdSr= OdSr left outer join InvFgd on IfFdTc= FdTc and " + _
''                "IfFdYy= FdYy and IfFdChr= FdChr and IfFdNo= FdNo and IfFdSr= FdSr and " + _
''                "IfTc= '" + ADC("wIdTc") + "' and IfYy= '" + ADC("wIdYy") + "' and " + _
''                "IfChr= '" + ADC("wIdChr") + "' and IfNo= " + CStr(ADC("wIdNo")) + " and IfSr= 0 " + _
''                "where OdTc='" + ADC("wDcOdTc") + "' and OdYy='" + ADC("wDcOdYy") + "' and " + _
''                "OdChr='" + ADC("wDcOdChr") + "' and OdNo=" + CStr(ADC("wDcOdNo")) + " and " + _
''                "OdSr= " + CStr(ADC("wDcOdSr")) + " " + _
''                "group by OdTc, OdYy, OdChr, OdNo, OdSr")
'
'
'  If wRsSrcDsg.RecCount = 0 Then DispMsg "Invalid Order Sr.", etError: Exit Sub: Set wRsSrcDsg = Nothing
'  wRsSrcDsg.MoveFirst
'  If wRsSrcDsg!qOdCls = "Y" Then DispMsg "Order Is Closed", etError: Exit Sub: Set wRsSrcDsg = Nothing
'  Call MakeDsgWithBags(False, wRsSrcDsg)
'
'  If GrdID.Rows > 1 Then CmdMI.Enabled = False
'  Call CmdDC_Click
'  Set wRsSrcDsg = Nothing
End Sub
Private Sub InitSummFlds(ByVal wReqIdName As String, ByVal wReqFldName As String)
  '*** Procedure to initialise the Group Option Fields in the summary frame
  Dim wOrgVal As String
  If ADC(wReqIdName) = "" Then
    wOrgVal = moCn.GetFldVal("Select " + wReqFldName + " from Head where HCoCd= '" + gs_CoCd + "' and HCd='" + ctSelfCmCd + "'")
  Else
    wOrgVal = ADC(wReqIdName)
  End If
  ADC(wReqIdName).InitialiseFld ""
  ADC(wReqIdName) = wOrgVal
End Sub
Private Sub DsgRmSumm()
  '*** Procedure to fill up the Design Rm Summary grid in the Summary Frame
  '*** The grid is filled on the basis of the Design Rm Group option selected by the user
  
  Dim wDsgRmTotIWt As Double, wDsgRmTotZWt As Double, wDsgRmTotQty As Double
  Dim wDsgRmTotIVal As Double, wDsgRmTotZVal As Double, wRsDsgRmSumm As MwfLib.MDORowSet
  Dim wRow As Integer, wsql As String
  wDsgRmTotIWt = 0: wDsgRmTotZWt = 0: wDsgRmTotQty = 0: wDsgRmTotIVal = 0: wDsgRmTotZVal = 0
  wsql = RmSummSql(ctDsgSumm)
  GrdSummDR.Rows = 1
  If wsql = "" Then DispMsg "Invalid GroupBy option", etError: Exit Sub
  Set wRsDsgRmSumm = moCn.OpenRes(wsql)
  GrdSummDR.AllowAdd = True
  With GrdSummDR
    Do While Not (wRsDsgRmSumm.EOF Or wRsDsgRmSumm.BOF)
      .AddItem
      wRow = .Rows - 1
      .Value(wRow, "wDrCtg") = wRsDsgRmSumm!qRmCtg
      .Value(wRow, "wDrSCtg") = wRsDsgRmSumm!qRmSCtg
      .Value(wRow, "wDrCd") = wRsDsgRmSumm!IrRmCd
      .Value(wRow, "wDrQty") = wRsDsgRmSumm!IrRmQty
      .Value(wRow, "wDrIWt") = wRsDsgRmSumm!IrRmIWt
      .Value(wRow, "wDrIVal") = wRsDsgRmSumm!IrRmIVal
      .Value(wRow, "wDrZWt") = wRsDsgRmSumm!irrmzwt
      .Value(wRow, "wDrZVal") = wRsDsgRmSumm!IrRmZVal
      Select Case Trim(wRsDsgRmSumm!qRmCtg)
      ' **** Manali 3.03 - 21/06/08 - Pd Mod - "L" included
      Case Is = "G", "S", "P", "L", "M", "X"
        wDsgRmTotIWt = wDsgRmTotIWt + wRsDsgRmSumm!IrRmIWt
        wDsgRmTotZWt = wDsgRmTotZWt + wRsDsgRmSumm!irrmzwt
      Case Is = "D", "C"
        wDsgRmTotIWt = wDsgRmTotIWt + MWLib.Div(wRsDsgRmSumm!IrRmIWt, 5)
        wDsgRmTotZWt = wDsgRmTotZWt + MWLib.Div(wRsDsgRmSumm!irrmzwt, 5)
      End Select
      wDsgRmTotQty = wDsgRmTotQty + wRsDsgRmSumm!IrRmQty
      wDsgRmTotIVal = wDsgRmTotIVal + wRsDsgRmSumm!IrRmIVal
      wDsgRmTotZVal = wDsgRmTotZVal + wRsDsgRmSumm!IrRmZVal
      wRsDsgRmSumm.MoveNext
    Loop
  End With
  
  ADC("wDRTotIWt") = wDsgRmTotIWt
  ADC("wDRTotZWt") = wDsgRmTotZWt
  ADC("wDRTotQty") = wDsgRmTotQty
  ADC("wDRTotIVal") = wDsgRmTotIVal
  ADC("wDRTotZVal") = wDsgRmTotZVal
  GrdSummDR.AllowAdd = False
End Sub
Private Sub InvRmSumm()
  '*** Procedure to fill up the Invoice Rm Summary grid in the Summary Frame
  '*** The grid is filled on the basis of the Invoice Rm Group option selected by the user
  
  Dim wInvRmTotIWt As Double, wInvRmTotZWt As Double, wInvRmTotQty As Double
  Dim wInvRmTotIVal As Double, wInvRmTotZVal As Double, wRsInvRmSumm As MwfLib.MDORowSet
  Dim wRow As Integer, wsql As String
  wInvRmTotIWt = 0: wInvRmTotZWt = 0: wInvRmTotQty = 0: wInvRmTotIVal = 0: wInvRmTotZVal = 0
  wsql = RmSummSql(ctInvSumm)
  GrdSummIR.Rows = 1
  If wsql = "" Then DispMsg "Invalid GroupBy option", etInfo: Exit Sub
  Set wRsInvRmSumm = moCn.OpenRes(wsql)
  GrdSummIR.AllowAdd = True
  With GrdSummIR
    Do While Not (wRsInvRmSumm.EOF Or wRsInvRmSumm.BOF)
      .AddItem
      wRow = .Rows - 1
      .Value(wRow, "wIrCtg") = wRsInvRmSumm!qRmCtg
      .Value(wRow, "wIrSCtg") = wRsInvRmSumm!qRmSCtg
      .Value(wRow, "wIrCd") = wRsInvRmSumm!IrRmCd
      .Value(wRow, "wIrQty") = wRsInvRmSumm!IrRmQty
      .Value(wRow, "wIrIWt") = wRsInvRmSumm!IrRmIWt
      .Value(wRow, "wIrIVal") = wRsInvRmSumm!IrRmIVal
      .Value(wRow, "wIrZWt") = wRsInvRmSumm!irrmzwt
      .Value(wRow, "wIrZVal") = wRsInvRmSumm!IrRmZVal
      Select Case Trim(wRsInvRmSumm!qRmCtg)
      ' **** Manali 3.03 - 21/06/08 - Pd Mod
      Case Is = "G", "S", "P", "L", "M", "X"
        wInvRmTotIWt = wInvRmTotIWt + wRsInvRmSumm!IrRmIWt
        wInvRmTotZWt = wInvRmTotZWt + wRsInvRmSumm!irrmzwt
      Case Is = "D", "C"
        wInvRmTotIWt = wInvRmTotIWt + MWLib.Div(wRsInvRmSumm!IrRmIWt, 5)
        wInvRmTotZWt = wInvRmTotZWt + MWLib.Div(wRsInvRmSumm!irrmzwt, 5)
      End Select
      wInvRmTotQty = wInvRmTotQty + wRsInvRmSumm!IrRmQty
      wInvRmTotIVal = wInvRmTotIVal + wRsInvRmSumm!IrRmIVal
      wInvRmTotZVal = wInvRmTotZVal + wRsInvRmSumm!IrRmZVal
      wRsInvRmSumm.MoveNext
    Loop
  End With
  
  ADC("wIrTotIWt") = wInvRmTotIWt
  ADC("wIrTotZWt") = wInvRmTotZWt
  ADC("wIrTotQty") = wInvRmTotQty
  ADC("wIrTotIVal") = wInvRmTotIVal
  ADC("wIrTotZVal") = wInvRmTotZVal
  GrdSummIR.AllowAdd = False
End Sub
Private Sub DsgLabSumm()
  '*** Procedure to fill up the Design Labour Summary grid in the Summary Frame
  '*** The grid is filled on the basis of the Design Lab Group option selected by the user
  
  Dim wRsDsgLabSumm As MDORowSet, wDsgLabTotIVal As Double, wDsgLabTotZVal As Double, wRow As Integer, wsql As String
  wDsgLabTotIVal = 0: wDsgLabTotZVal = 0
  wsql = LabSummSql(ctDsgSumm)
  GrdSummDL.Rows = 1
  If wsql = "" Then DispMsg "Invalid GroupBy option", etInfo: Exit Sub
  Set wRsDsgLabSumm = moCn.OpenRes(wsql)
  GrdSummDL.AllowAdd = True
  With GrdSummDL
    Do While Not (wRsDsgLabSumm.EOF Or wRsDsgLabSumm.BOF)
      .AddItem
      wRow = .Rows - 1
      .Value(wRow, "wDlMCd") = wRsDsgLabSumm!IlMCd
      .Value(wRow, "wDlSCd") = wRsDsgLabSumm!IlSCd
      .Value(wRow, "wDlQw") = wRsDsgLabSumm!IlQw
      .Value(wRow, "wDlQty") = wRsDsgLabSumm!IlQty
      .Value(wRow, "wDliVal") = wRsDsgLabSumm!IlIVal
      .Value(wRow, "wDlZVal") = wRsDsgLabSumm!iLZVal
      wDsgLabTotIVal = wDsgLabTotIVal + wRsDsgLabSumm!IlIVal
      wDsgLabTotZVal = wDsgLabTotZVal + wRsDsgLabSumm!iLZVal
      wRsDsgLabSumm.MoveNext
    Loop
  End With
  ADC("wDlTotiVal") = wDsgLabTotIVal
  ADC("wDlTotZVal") = wDsgLabTotZVal
  GrdSummDL.AllowAdd = False
End Sub
Private Sub InvLabSumm()
  '*** Procedure to fill up the Invoice Labour Summary grid in the Summary Frame
  '*** The grid is filled on the basis of the Invoice Lab Group option selected by the user
  
  Dim wRsInvLabSumm As MDORowSet, wInvLabTotIVal As Double, wInvLabTotZVal As Double, wRow As Integer, wsql As String
  wInvLabTotIVal = 0: wInvLabTotZVal = 0
  GrdSummIL.Rows = 1
  wsql = LabSummSql(ctInvSumm)
  If wsql = "" Then DispMsg "Invalid GroupBy option", etInfo: Exit Sub
  Set wRsInvLabSumm = moCn.OpenRes(wsql)
  GrdSummIL.AllowAdd = True
  With GrdSummIL
    Do While Not (wRsInvLabSumm.EOF Or wRsInvLabSumm.BOF)
      .AddItem
      wRow = .Rows - 1
      .Value(wRow, "wiLMCd") = wRsInvLabSumm!IlMCd
      .Value(wRow, "wiLSCd") = wRsInvLabSumm!IlSCd
      .Value(wRow, "wiLQw") = wRsInvLabSumm!IlQw
      .Value(wRow, "wiLQty") = wRsInvLabSumm!IlQty
      .Value(wRow, "wiLiVal") = wRsInvLabSumm!IlIVal
      .Value(wRow, "wiLZVal") = wRsInvLabSumm!iLZVal
      wInvLabTotIVal = wInvLabTotIVal + wRsInvLabSumm!IlIVal
      wInvLabTotZVal = wInvLabTotZVal + wRsInvLabSumm!iLZVal
      wRsInvLabSumm.MoveNext
    Loop
  End With
  ADC("wiLTotiVal") = wInvLabTotIVal
  ADC("wiLTotZVal") = wInvLabTotZVal
  GrdSummIL.AllowAdd = False
End Sub
Private Sub InvDsgSumm()
  '*** Procedure to fill up the Invoice Design Summary grid in the Summary Frame
  '*** The grid is filled on the basis of the Invoice Design Group option selected by the user
  
  Dim wInvDsgTotQty As Double, wInvDsgTotiFOB As Double, wInvDsgTotZFOB As Double, wReqGrpBy As String, wStr As String, wCommStr As String, wRow As Integer
  Dim wRsInvDsgSumm As MDORowSet
  Select Case ADC("wIDGrp")
  Case Is = "DMCD"
    wStr = "Select '*' as IdSr, DmCtg as qDmCtg, IdDmCd, 0 as qIPrc, 0 as qZPrc, "
    wReqGrpBy = "DmCtg, IdDmCd"
  Case Is = "DMCTG"
    wStr = "Select '*' as IdSr, DmCtg as qDmCtg, '*' as IdDmCd, 0 as qIPrc, 0 as qZPrc, "
    wReqGrpBy = "DmCtg"
  Case Is = "DSGSR"
    wStr = "Select IdSr, max(DmCtg) as qDmCtg, max(IdDmCd) as IdDmCd, max(case when IdQty> 0 then IdIFOB/IdQty else 0 end) as qIPrc, max(case when IdQty> 0 then IdZFOB/IdQty else 0 end) as qZPrc, "
    wReqGrpBy = "IdSr"
  Case Else
    DispMsg "Enter Invoice Design Group Option", etError: Exit Sub
  End Select
  
  wInvDsgTotQty = 0: wInvDsgTotiFOB = 0: wInvDsgTotZFOB = 0
  
  '****** Sachin 3.02 28-12-07 - Corresponding Partition
  wCommStr = "sum(IdQty) as IdQty, sum(IdIFOB) as IdIFOB, sum(IdZFOB) as IdZFOB " + _
             "from InvDsg join DsgMst on DmTcTyp= 'DM' and " + _
             "DmCd= IdDmCd and DmSz= '' " + _
             "Where IdCoCd= '" + gs_CoCd + "' and IdTc= '" + ADC("wIdTc") + "' And " + _
             "IdYy= '" + ADC("wIdYy") + "' And IdChr= '" + ADC("wIdChr") + "' And " + _
             "IdNo= " + CStr(ADC("wIdNo")) + " and IdPrtKey=(Case When '" + ADC("wIdPrtKey") + "'='" + ctCurrPrtn + "' Then '" + ctCurrPrtn + "' Else IdPrtKey End) group by " + wReqGrpBy
             '(index= DsgMst)
  wStr = wStr + wCommStr
  
  Set wRsInvDsgSumm = moCn.OpenRes(wStr)
  
  GrdSummID.AllowAdd = True
  GrdSummID.Rows = 1
  
  With GrdSummID
    Do While Not (wRsInvDsgSumm.EOF Or wRsInvDsgSumm.BOF)
      .AddItem
      wRow = .Rows - 1
      .Value(wRow, "wIdSr") = wRsInvDsgSumm!IdSr
      .Value(wRow, "wIdDmCtg") = wRsInvDsgSumm!qDmCtg
      .Value(wRow, "wIdDmCd") = wRsInvDsgSumm!IdDmCd
      .Value(wRow, "wIdQty") = wRsInvDsgSumm!IdQty
      .Value(wRow, "wIdIPrc") = wRsInvDsgSumm!qIPrc
      .Value(wRow, "wIdIFOB") = wRsInvDsgSumm!IdIFob
      .Value(wRow, "wIdZPrc") = wRsInvDsgSumm!qZPrc
      .Value(wRow, "wIdZFOB") = wRsInvDsgSumm!IdZFob
      wInvDsgTotQty = wInvDsgTotQty + wRsInvDsgSumm!IdQty
      wInvDsgTotiFOB = wInvDsgTotiFOB + wRsInvDsgSumm!IdIFob
      wInvDsgTotZFOB = wInvDsgTotZFOB + wRsInvDsgSumm!IdZFob
      wRsInvDsgSumm.MoveNext
    Loop
  End With
  ADC("wIdTotQty") = wInvDsgTotQty
  ADC("wIdTotiFOB") = wInvDsgTotiFOB
  ADC("wIdTotZFob") = wInvDsgTotZFOB
  GrdSummID.AllowAdd = False
End Sub
Private Function RmSummSql(ByVal mInvDsgSumm As String) As String
  '*** Common Procedure used in the Invoice and Design Rm Summary option
  
  Dim wCommStr As String, wCancel As Boolean, wErrMsg As String, wReqCnd As String, wReqIdName As String, wReqGrpBy As String
  If mInvDsgSumm = ctDsgSumm Then
    wReqCnd = " And IrSr =" + CStr(GrdID.Value(GrdID.Row, "IdSr"))
    wReqIdName = "wDRGrp"
  ElseIf mInvDsgSumm = ctInvSumm Then
    wReqIdName = "wIRGrp"
  End If
  Call ATXT_FldChange(99, wReqIdName, wCancel, wErrMsg, ADC(wReqIdName))
  If wCancel = True Then
    DispMsg wErrMsg, etError
    Exit Function
  End If
  Select Case ADC(wReqIdName)
  Case Is = "RMCTG"
    wReqGrpBy = "RmCtg"
    RmSummSql = "Select RmCtg as qRmCtg, '*' as qRmSCtg, '*' as IrRmCd, "
  Case Is = "RMSCTG"
    wReqGrpBy = "RmCtg,RmSCtg"
    RmSummSql = "Select RmCtg as qRmCtg, RmSCtg as qRmSCtg, '*' as IrRmCd, "
  Case Is = "RMCD"
    wReqGrpBy = "RmCtg,RmSCtg,IrRmCd"
    RmSummSql = "Select RmCtg as qRmCtg, RmSCtg as qRmSCtg, IrRmCd, "
  Case Else
    DispMsg "No Grouping Flds defined in system for selected GroupBy value", etError
    Exit Function
  End Select

  wCommStr = "sum(IsNull(IrRmQty, 0)) as IrRmQty, sum(IsNull(IrRmIWt, 0)) as IrRmIWt, " + _
             "sum(IsNull(IrRmIVal, 0)) as IrRmIVal, sum(IsNull(IrRmZWt, 0)) as IrRmZWt, " + _
             "sum(IsNull(IrRmZVal, 0)) as IrRmZVal " + _
             "from InvRm Join RmMst On RmCd= IrRmCd " + IIF(gs_Partition = ctCurrPrtn, " And RmPrtKey=IrPrtKey ", "") + _
             "Where " + IIF(gs_Partition = ctCurrPrtn, " IrPrtKey='" + ctCurrPrtn + "' And ", "") + " IrCoCd= '" + gs_CoCd + "' and IrTc ='" + ADC("wIdTc") + "' And " + _
             "IrYy = '" + ADC("wIdYy") + "' And IrChr ='" + ADC("wIdChr") + "' And " + _
             "IrNo =" + CStr(ADC("wIdNo")) + wReqCnd + " Group By " + wReqGrpBy
             '(index=RmMst)
  RmSummSql = RmSummSql + wCommStr
End Function
Private Function LabSummSql(ByVal pv_DsgOrdSumm As String) As String
  '*** Common Procedure used in the Invoice and Design Labour Summary option
  
  Dim wCommStr1 As String, wCommStr2 As String, wDsgStr1 As String, wDsgStr2 As String, wReqIdName As String, wCancel As Boolean, wErrMsg As String
  If pv_DsgOrdSumm = ctDsgSumm Then
    wDsgStr1 = " And IlSr =" + CStr(GrdID.Value(GrdID.Row, "IdSr"))
    wDsgStr2 = " And IrSr =" + CStr(GrdID.Value(GrdID.Row, "IdSr"))
    wReqIdName = "wDLGrp"
  ElseIf pv_DsgOrdSumm = ctInvSumm Then
    wReqIdName = "wILGrp"
  End If
  Call ATXT_FldChange(99, wReqIdName, wCancel, wErrMsg, ADC(wReqIdName))
  If wCancel = True Then
    DispMsg wErrMsg, etError
    Exit Function
  End If
  '****** Sachin 3.02 28-12-07 - Corresponding Partition
  wCommStr1 = "sum(IsNull(iLQty, 0)) as iLQty, sum(IsNull(iLiVal, 0)) as iLiVal, " + _
              "sum(IsNull(iLZVal, 0)) as iLZVal " + _
              "from InvLab Where iLCoCd= '" + gs_CoCd + "' and iLTc ='" + ADC("wIdTc") + "' " + _
              "And iLYy = '" + ADC("wIdYy") + "' And iLChr ='" + ADC("wIdChr") + "' And " + _
              "iLNo =" + CStr(ADC("wIdNo")) + " and IlPrtKey=(Case when '" + ADC("wIdPrtKey") + "'='" + ctCurrPrtn + "' Then '" + ctCurrPrtn + "' Else IlPrtKey End) " + wDsgStr1

  '****** Sachin 3.02 28-12-07 - Corresponding Partition
  wCommStr2 = "sum(IsNull(IrRmQty, 0)) as iLQty, sum(IsNull(IrSetIVal, 0)) as iLiVal, " + _
              "sum(IsNull(IrSetZVal, 0)) as iLZVal " + _
              "from InvRm where IrCoCd = '" + gs_CoCd + "' And " + _
              "IrTc = '" + ADC("wIdTc") + "' And IrYy = '" + ADC("wIdYy") + "' And " + _
              "IrChr = '" + ADC("wIdChr") + "' And IrNo = " + CStr(ADC("wIdNo")) + " and " + _
              "IrSetSCd<> '' and IrPrtKey=(Case when '" + ADC("wIdPrtKey") + "'='" + ctCurrPrtn + "' Then '" + ctCurrPrtn + "' Else IrPrtKey End) " + wDsgStr2
  Select Case ADC(wReqIdName)
  Case Is = "LMCD"
    LabSummSql = "Select iLMCd, '*' as iLSCd, '*' as iLQw, " + wCommStr1 + _
           " group by iLMCd " + _
           " Union All " + _
           " Select 'SET' as iLMCd, '*' as iLSCd, '*' as iLQw, " + wCommStr2 + _
           " group by IrTc, IrYy, IrChr, IrNo "
  Case Is = "LSCD"
    LabSummSql = "Select iLMCd, iLSCd, max(IsNull(iLQw, '')) as iLQw, " + wCommStr1 + _
           " group by iLMCd,iLSCd " + _
           " Union All " + _
           "Select 'SET' as iLMCd, IrSetSCd as iLSCd, " + _
           "(Select max(PValue) from Param where PTyp='LABSCD' and " + _
           "PMCd= 'SET' and PSCd= IrSetSCd) as iLQw, " + wCommStr2 + _
           " Group By IrSetSCd"
  Case Else
    DispMsg "No Grouping Flds defined in system for selected GroupBy value", etError
    Exit Function
  End Select
End Function
Private Sub MakeInvoice()
'*** The Make Invoice Option generates all the design, Rm & Lab level records
    'for the entire invoice at one shot if all the invoice bags have been entered through
    'the 'Invoice All Bags' Entry form.
'*** This option cannot be used if there are records existing in the GrdID grid
  
  Dim wRsSrcDsg As MDORowSet
  
  ' **** Manali 3.8.0 - Invoice Sort By
  Dim ws_OrdBy As String, wAllBagsCnd As String
  
  ws_OrdBy = moCn.GetFldVal("Select PDesc225 From Param Where PTyp='INSRT' and PMCd='" + ADC("wInSrt") + "' and PSCd=''")
  ws_OrdBy = ws_OrdBy + IIF(ws_OrdBy <> "", ", ", "") + " OdCoCd, OdTc, OdYy, OdChr, OdNo, OdSr"
  ws_OrdBy = " Order By " + ws_OrdBy
  ' **** Manali 3.8.0 - Invoice Sort By
  
  '****** Sachin 3.3.0 (23-06-08) - [42. Make Invoice By Bag] ******
  ADC("wIdInByBag") = moCn.GetFldVal(" Select InByBagYN From InvHd " + _
                                  " Where InCoCd= '" + gs_CoCd + "' and " + _
                                  " InTc= '" + ADC("wIdTc") + "' and " + _
                                  " InYy= '" + ADC("wIdYy") + "' and " + _
                                  " InChr= '" + ADC("wIdChr") + "' and " + _
                                  " InNo= " + CStr(ADC("wIdNo")) + " and InPrtKey=(Case When '" + gs_Partition + "'='" + ctCurrPrtn + "' Then '" + ctCurrPrtn + "' Else InPrtKey End) ")
  mb_InByBag = IIF(ADC("wIdInByBag") = "Y", True, False)
  '****** Sachin 3.3.0 (23-06-08) - [42. Make Invoice By Bag] ******
    
  If GrdID.Rows > 1 Then DispMsg "Cannot Make Invoice As Design Record Already Exists", etError: Exit Sub
  '****** (Jen 3.01.0_1 added OrdRm join for MainMet) ******
  '****** Sachin 3.02 - Current Partition
  '***Geeta****3.02 Patch****Added OdILabWtFrOrd
  '****** Sachin 3.3.0 (23-06-08) - [42. Make Invoice By Bag] - mb_InByBag added in the Group By Clause ******
  ' **** Manali 3.8.0 - ws_OrdBy added, DsgMst Join added
  
  If UCase(ms_IdTcTyp) = UCase(ctInTcTyp) Then
        wAllBagsCnd = "having ((max(OdInvAllBags) = 'Y' and Str(max(OdOrdQty), 16, 1)= Str(max(OdExpQty)+ sum(IsNull(IfQty, FdQty- FdDespQty)), 16, 1)) Or (max(OdInvAllBags) = 'N' and Round(max(OdOrdQty), 1)>= Round(max(OdExpQty)+ sum(IsNull(IfQty, FdQty- FdDespQty)), 1)) ) "
  ElseIf UCase(ms_IdTcTyp) = UCase(ctCTBTcTyp) Then
    wAllBagsCnd = " "
  End If

  Set wRsSrcDsg = moCn.OpenRes("Select Max(OdDmCd) as qDmCd, Max(OdSfx) as qSfx, " + _
                  "Max(OdDmSz) as qDmSz, sum(IfQty) as qQty, OdTc as qOdTc, " + _
                  "OdYy as qOdYy, OdChr as qOdChr, OdNo as qOdNo, " + _
                  "OdSr as qOdSr, " + IIF(mb_InByBag = True, " IfFdBYy as qBYy, IfFdBChr as qBChr, IfFdBNo as qBNo, ", "") + " Max(Fd" + ms_OrdTyp + "CmCd) as qCmCd, " + _
                  "Max(OdVaCtg) as qVaCtg, Max(OdKt) as qKt, Max(OdFixPrc) as qFixPrc, " + _
                  "Max(OdGldAs) as qGldAs, Max(OdLabAs) as qLabAs, max(OdIWtEqOrd) as OdIWtEqOrd,max(OdILabWtFrOrd) as OdILabWtFrOrd, " + _
                  "max(OdInvAllBags) as OdInvAllBags,max(OdLabAsWt) as OdLabAsWt, " + _
                  "max(OdIWtFrOrd) as OdIWtFrOrd, max(OdIGldRtEqOrd) as OdIGldRtEqOrd, max(OdHld) as qOdHld, " + _
                  "max(IsNull(OrRmCtg, '')) as qMainMetRmCtg " + _
                  "from InvFgd join Fgd on FdIdNo=IfFdIdNo and FdCoCd= IfCoCd and FdTc= IfFdTc and " + _
                  "FdYy= IfFdYy and FdChr= IfFdChr and FdNo= IfFdNo and FdSr= IfFdSr and FdPrtKey=IfPrtKey " + _
                  "join OrdDsg on /* OdIdNo=Fd" + ms_OrdTyp + "OdIdNo and */ OdCoCd= FdCoCd and OdTc= Fd" + ms_OrdTyp + "OdTc " + _
                  "and OdYy= Fd" + ms_OrdTyp + "OdYy and OdChr= Fd" + ms_OrdTyp + "OdChr " + _
                  "and OdNo= Fd" + ms_OrdTyp + "OdNo and OdSr= Fd" + ms_OrdTyp + "OdSr and OdPrtKey=FdPrtKey " + _
                  "Join DsgMst On DmTcTyp='DM' and DmCd=OdDmCd and DmSz='' " + _
                  "left outer join OrdRm on OrOdIdNo=OdIdNo and OrCoCd= OdCoCd and OrTc= OdTc and OrYy= OdYy and OrChr= OdChr " + _
                  "and OrNo= OdNo and OrSr= OdSr and OrPrtKey=OdPrtKey and OrMainMet= 'Y' " + _
                  "where IfCoCd= '" + gs_CoCd + "' and IfTc= '" + ADC("wIdTc") + "' and " + _
                  "IfYy= '" + ADC("wIdYy") + "' and IfChr= '" + ADC("wIdChr") + "' and " + _
                  "IfNo= " + CStr(ADC("wIdNo")) + " and IfPrtKey='" + ctCurrPrtn + "' group by OdCoCd, OdTc, OdYy, OdChr, OdNo, OdSr" + _
                  IIF(mb_InByBag = True, ",IfFdBYy, IfFdBChr, IfFdBNo  ", " ") + wAllBagsCnd + ws_OrdBy)
  If wRsSrcDsg.RecCount = 0 Then DispMsg "No Records In Invoice Bags / All Bags Not Yet ready for Export.", etError: Exit Sub: Set wRsSrcDsg = Nothing
  TrackStatus ("Retrieved source")
  
  Call MakeDsgWithBags(True, wRsSrcDsg)
  If GrdID.Rows > 1 Then CmdMI.Enabled = False
  Set wRsSrcDsg = Nothing
End Sub
Private Sub MakeDsgWithBags(ByVal mMkInv As Boolean, ByVal mRsSrcDsg As MDORowSet)
  '*** If this procedure was called by the MakeInv procedure, all the Design, Rm & Lab
      'records for the entire invoice are generated
      '(the bags are already existing in the InvFgd table)
  '*** If this procedure is called by the Design Copy option then only Design, Bag, Rm &
      'Lab details of the current Inv Design row are copied
  Dim wRsSrcBag As MDORowSet, wDsgRow As Integer, wBagRow As Integer, wRmRow As Integer
  Dim wLabRow As Integer
  'wSelCnd As String
  '%%%, wInvBagExists As Boolean, wTotDsgQty As Single
  Dim i As Single, wStr As String, wRsSrcLab As MDORowSet, wIfSrNo As Single, wCnt As Integer
   Dim wDmHldCnt As Integer
   
  Dim wi_Cnt As Single, wi_TotCnt As Single         '4.1.3.0
  Dim wByBagStr As String                           '****** Sachin 3.3.0 (23-06-08) - [42. Make Invoice By Bag] ******
   
  With GrdID
    mf_ShowDsg = False
    wCnt = 0
    wDmHldCnt = 0
    GrdID.StartCopy
   ' GrdIR.StartCopy
    'GrdIL.StartCopy
    GrdID.StopRecursiveRecalc
    GrdIR.StopRecursiveRecalc
    GrdIL.StopRecursiveRecalc
    
    mf_StopRecalcDuringCopyIR = False
    mf_StopRecalcDuringCopyIL = False
    wi_TotCnt = mRsSrcDsg.RecCount
    wi_Cnt = 0
    GrdID.Col = 0
    BeginEmrProcess Me, "Start"
    Do While Not (mRsSrcDsg.EOF Or mRsSrcDsg.BOF)
    If (mRsSrcDsg!qOdHld = "N") Then
      wCnt = wCnt + 1
      If mMkInv = True Then
        .AddItem
        wDsgRow = .Rows - 1
        .Row = wDsgRow
      Else
        If Not (wCnt = 1 And GrdIR.Rows = 1 And GrdIL.Rows = 1 And GrdIB.Rows = 1) Or GrdID.Rows = 1 Then
          .AddItem
          wDsgRow = .Rows - 1
          .Row = wDsgRow
        End If
      End If
      wDsgRow = .Row
      
      UpdateEmrProcess "Copying........InvDesign" + CStr(wDsgRow)
      
      GrdID.StopRecursiveRecalc
      GrdIR.StopRecursiveRecalc
      GrdIL.StopRecursiveRecalc
      GrdIB.StopRecursiveRecalc
      
      If .Value(wDsgRow, "IdSr") = 0 Then .Value(wDsgRow, "IdSr") = .MaxVal("IdSr") + 1
      .Value(wDsgRow, "IdPrtKey") = ctCurrPrtn
      .Value(wDsgRow, "IdDmCd") = mRsSrcDsg!qDmCd
      .Value(wDsgRow, "IdSfx") = mRsSrcDsg!qSfx
      .Value(wDsgRow, "IdDmSz") = mRsSrcDsg!qDmSz
      .Value(wDsgRow, "IdQty") = mRsSrcDsg!qQty
      .Value(wDsgRow, "IdExpOdTc") = mRsSrcDsg!qOdTc
      .Value(wDsgRow, "IdExpOdYy") = mRsSrcDsg!qOdYy
      .Value(wDsgRow, "IdExpOdChr") = mRsSrcDsg!qOdChr
      .Value(wDsgRow, "IdExpOdNo") = mRsSrcDsg!qOdNo
      .Value(wDsgRow, "IdExpOdSr") = mRsSrcDsg!qOdSr
      .Value(wDsgRow, "IdExpCmCd") = mRsSrcDsg!qCmCd
      .Value(wDsgRow, "IdVaCtg") = mRsSrcDsg!qVaCtg
      .Value(wDsgRow, "IdKt") = mRsSrcDsg!qKt
      .Value(wDsgRow, "IdCustmMtchWt") = 0
      If ms_RemakeOpt <> "" Then
        If moCn.RecSeek("Select 'x' where ','+ '" + ms_RemakeOpt + "'+ ',' like '%,'+ '" + mRsSrcDsg!qMainMetRmCtg + "'+ ',%' ") _
          Then .Value(wDsgRow, "IdGldLs") = 0
      End If
      .Value(wDsgRow, "IdLabWtFrOrd") = mRsSrcDsg!OdILabWtFrOrd
      If UCase(ms_IdTcTyp) = UCase(ctInTcTyp) Then
        .Value(wDsgRow, "IdFixPrc") = mRsSrcDsg!qFixPrc
        .Value(wDsgRow, "IdHZValEqInv") = mIdHZValEqInv
        If mRsSrcDsg!OdILabWtFrOrd = "Y" Then
          .Value(wDsgRow, "IdILabAsWt") = mRsSrcDsg!OdLabAsWt * mRsSrcDsg!qQty
          .Value(wDsgRow, "IdZLabAsWt") = mRsSrcDsg!OdLabAsWt * mRsSrcDsg!qQty
        End If
      ElseIf UCase(ms_IdTcTyp) = UCase(ctCTBTcTyp) Then
        .Value(wDsgRow, "IdFixPrc") = "N"
        .Value(wDsgRow, "IdHZValEqInv") = "N"
      End If
      .Value(wDsgRow, "IdGldAs") = mRsSrcDsg!qGldAs
      .Value(wDsgRow, "IdLabAs") = mRsSrcDsg!qLabAs     '****** Sachin 2.13.0 - 26-05-2006 ******
      If Not .SaveRec(wDsgRow) Then GoTo NextOd
      wByBagStr = ""
      If mb_InByBag = True Then
        wByBagStr = " and IfFdBYy='" + mRsSrcDsg!qBYy + "' and IfFdBChr='" + mRsSrcDsg!qBChr + "' and IfFdBNo=" + CStr(mRsSrcDsg!qBNo)
      End If
        wStr = " Update InvFgd Set IfSr= " + CStr(GrdID.Value(wDsgRow, "IdSr")) + _
               " From InvFgd If1 Where IfCoCd = '" + gs_CoCd + "' " + _
               " And IfTc ='" + ADC("wIdTc") + "' And IfYy = '" + ADC("wIdYy") + "'" + _
               " And IfChr ='" + ADC("wIdChr") + "' And IfNo =" + CStr(ADC("wIdNo")) + _
               " And IfSr= 0 and IfPrtKey='" + ctCurrPrtn + "' and exists(Select 'x' from Fgd " + _
               " Where FdCoCd= IfCoCd and FdTc= IfFdTc " + _
               " And FdYy= IfFdYy and FdChr= IfFdChr and FdNo= IfFdNo and FdSr= IfFdSr And FdPrtKey=IfPrtKey " + _
               " And Fd" + ms_OrdTyp + "OdTc= '" + GrdID.Value(wDsgRow, "IdExpOdTc") + "' " + _
               " And Fd" + ms_OrdTyp + "OdYy= '" + GrdID.Value(wDsgRow, "IdExpOdYy") + "' " + _
               " And Fd" + ms_OrdTyp + "OdChr= '" + GrdID.Value(wDsgRow, "IdExpOdChr") + "' " + _
               " And Fd" + ms_OrdTyp + "OdNo= " + CStr(GrdID.Value(wDsgRow, "IdExpOdNo")) + _
               " And Fd" + ms_OrdTyp + "OdSr= " + CStr(GrdID.Value(wDsgRow, "IdExpOdSr")) + ") " + _
               wByBagStr
         moCn.Execute (wStr)
      If mMkInv = False Then
          wByBagStr = ""
          If mb_InByBag = True Then
            wByBagStr = " and FdBYy='" + mRsSrcDsg!qBYy + "' and FdBChr='" + mRsSrcDsg!qBChr + "' and FdBNo=" + CStr(mRsSrcDsg!qBNo)
          End If
          Set wRsSrcBag = moCn.OpenRes("Select * from Fgd B join Fg on FgIdNo=FdFgIdNo and FgCoCd= FdCoCd and " + _
                          "FgTc= FdTc and FgYy= FdYy and FgChr= FdChr and FgNo= FdNo and FgPrtKey=FdPrtKey " + _
                          "join Param on PTyp= 'TC' and PmCd= FgTc and PSCd= '' and " + _
                          "PValue in ('FB', 'FR') where FdCoCd= '" + gs_CoCd + "' And " + _
                          "Fd" + ms_OrdTyp + "OdTc= '" + GrdID.Value(wDsgRow, "IdExpOdTc") + "' " + _
                          "And Fd" + ms_OrdTyp + "OdYy= '" + GrdID.Value(wDsgRow, "IdExpOdYy") + "' " + _
                          "And Fd" + ms_OrdTyp + "OdChr= '" + GrdID.Value(wDsgRow, "IdExpOdChr") + "' " + _
                          "And Fd" + ms_OrdTyp + "OdNo= " + CStr(GrdID.Value(wDsgRow, "IdExpOdNo")) + _
                          "And Fd" + ms_OrdTyp + "OdSr= " + CStr(GrdID.Value(wDsgRow, "IdExpOdSr")) + " and FdPrtKey='" + ctCurrPrtn + "' " + _
                          "And FdQty> " + ms_DespTyp + wByBagStr + _
                          IIF(UCase(ms_IdTcTyp) = UCase(ctCTBTcTyp), _
                          " and not exists(Select 'x' from Fgd M where M.FdCoCd= B.FdCoCd " + _
                          "and M.FdRefYy= B.FdYy and M.FdRefKey= B.FdKey and M.FdPrtKey=B.FdPrtKey " + _
                          "and (Select PValue from Param where PTyp= 'TC' and PMCd= M.FdTc)= 'FWB') ", "") + _
                          " Order By FdBYy, FdBChr, FdBNo")
        With GrdIB
          GrdIB.StartCopy
          wIfSrNo = 0
            wIfSrNo = moCn.GetFldVal("Select max(IfSrNo) from InvFgd " + _
                "where IfCoCd= '" + gs_CoCd + "' and " + _
                "IfTc= '" + ADC("wIdTc") + "' and IfYy= '" + ADC("wIdYy") + "' and " + _
                "IfChr= '" + ADC("wIdChr") + "' and IfNo= " + CStr(ADC("wIdNo")) + " and IfPrtKey='" + ctCurrPrtn + "'")
          Do While Not (wRsSrcBag.EOF Or wRsSrcBag.BOF)
            .AddItem
            wBagRow = .Rows - 1
            wIfSrNo = wIfSrNo + 1
            .Value(wBagRow, "IfSrNo") = wIfSrNo
            .Value(wBagRow, "IfFdTc") = wRsSrcBag!FdTc
            .Value(wBagRow, "IfFdYy") = wRsSrcBag!FdYy
            .Value(wBagRow, "IfFdChr") = wRsSrcBag!FdChr
            .Value(wBagRow, "IfFdNo") = wRsSrcBag!FdNo
            .Value(wBagRow, "IfFdSr") = wRsSrcBag!FdSr
            .Value(wBagRow, "IfFdBYy") = wRsSrcBag!FdBYy
            .Value(wBagRow, "IfFdBChr") = wRsSrcBag!FdBChr
            .Value(wBagRow, "IfFdBNo") = wRsSrcBag!FdBNo
            .Value(wBagRow, "IfQty") = wRsSrcBag!Fdqty - wRsSrcBag(ms_DespTyp)
            .Value(wBagRow, "IfGrsWt") = wRsSrcBag!FdGrWt * MWLib.Div((wRsSrcBag!Fdqty - wRsSrcBag(ms_DespTyp)), wRsSrcBag!Fdqty)
            .Value(wBagRow, "IfPrtKey") = ctCurrPrtn        '****** Sachin 3.02 28-12-07
            .SaveRec (wBagRow)
            wRsSrcBag.MoveNext
          Loop
          GrdIB.EndCopy
        End With
      End If
      Dim wb_BagExists As Boolean
      wb_BagExists = moCn.RecSeek(" Select 'x' From InvFgd Where IfCoCd = '" + gs_CoCd + "' " + _
                                  " And IfTc ='" + ADC("wIdTc") + "' And IfYy = '" + ADC("wIdYy") + "'" + _
                                  " And IfChr ='" + ADC("wIdChr") + "' And IfNo= " + CStr(ADC("wIdNo")) + _
                                  " And IfSr= " + CStr(GrdID.Value(wDsgRow, "IdSr")) + " and IfPrtKey='" + ctCurrPrtn + "' ")
      If wb_BagExists = True Then
        If ms_hSizeStkRt = "N" Then '6.1-55 IIF(ms_RtFrInvCust = "Y", ADC("wIdCmCd") added
          Call MakeRm(.Value(wDsgRow, "IdSr"), .Value(wDsgRow, "IdExpOdTc"), _
                      .Value(wDsgRow, "IdExpOdYy"), .Value(wDsgRow, "IdExpOdChr"), _
                      .Value(wDsgRow, "IdExpOdNo"), .Value(wDsgRow, "IdExpOdSr"), _
                      IIF(ms_RtFrInvCust = "Y", ADC("wIdCmCd"), .Value(wDsgRow, "IdExpCmCd")), .Value(wDsgRow, "IdDmCd"), _
                      .Value(wDsgRow, "IdQty"), mRsSrcDsg!OdIwtEqOrd, _
                      mRsSrcDsg!OdIwtFrOrd, mRsSrcDsg!OdIGldRtEqOrd)
        Else
          Call MakeRmFromFg(mRsSrcDsg!OdIwtEqOrd, mRsSrcDsg!OdIwtFrOrd, mRsSrcDsg!OdIGldRtEqOrd, .Value(wDsgRow, "IdQty"))
        End If
      End If
      CalcGrdIdValsForWts (wDsgRow)
      ForCpy_UpdIRMainMetVal
      If wb_BagExists = True Then
        Call MakeLab(.Value(wDsgRow, "IdExpOdTc"), .Value(wDsgRow, "IdExpOdYy"), _
                     .Value(wDsgRow, "IdExpOdChr"), .Value(wDsgRow, "IdExpOdNo"), _
                     .Value(wDsgRow, "IdExpOdSr"))
      End If
  
      CalcGrdIdValsForVal (wDsgRow)
      'GoTo NextOd
      GrdID.SaveRec wDsgRow
'      If mMkInv = False Then GrdID.Row = wDsgRow       '%%% (New)
    ElseIf mRsSrcDsg!qOdHld = "Y" Then
      wDmHldCnt = wDmHldCnt + 1
    End If

NextOd:
      mf_StopRecalcDuringCopyIR = False
      mf_StopRecalcDuringCopyIL = False
      mRsSrcDsg.MoveNext
    Loop
EndCopy:
    mf_StopRecalcDuringCopyIR = False
    mf_StopRecalcDuringCopyIL = False
    wi_Cnt = 0: wi_TotCnt = 0   '4.1.3.0
 '   MWLib.EndProcess Me         '4.1.3.0
    EndEmrProcess Me
    GrdIR.EndCopy
    GrdIL.EndCopy
    GrdID.EndCopy
    GrdID.StartRecursiveRecalc
    GrdIR.StartRecursiveRecalc
    GrdIL.StartRecursiveRecalc
    If wDmHldCnt > 0 Then _
    DispMsg "Cannot Copy " + CStr(wDmHldCnt) + " Design(s) on Hold", etWarning
    mf_ShowDsg = True
    Call atxt_ReCalculate(0, "wTotIdIFob", 0, 0, "", "")
    Call atxt_ReCalculate(0, "wTotIdZFob", 0, 0, "", "")
    
    Set mRsSrcDsg = Nothing
    Set wRsSrcBag = Nothing
  End With
End Sub
Private Sub ForCpy_UpdIRMainMetVal()
            Dim wQw As String
      Dim i As Long
      For i = 1 To GrdIR.Rows - 1
      
        If Not GrdIR.IsDel(i) And GrdIR.Value(i, "IrMainMet") = "Y" Then
        
          wQw = moCn.GetFldVal("Select RmQw from RmMst where " + _
                                       "RmCd= '" + GrdIR.Value(i, "IrRmCd") + "' and RmPrtKey='" + ctCurrPrtn + "' ")
          GrdIR.Value(i, "IrRmIVal") = IIF(wQw = "Q", _
                                       GrdIR.Value(i, "IrRmQty"), GrdID.Value(GrdID.Row, "IdIGldAsWt")) * GrdIR.Value(i, "IrRmIRt")
          GrdIR.Value(i, "IrRmZVal") = IIF(wQw = "Q", _
                                       GrdIR.Value(i, "IrRmQty"), GrdID.Value(GrdID.Row, "IdZGldAsWt")) * GrdIR.Value(i, "IrRmIRt")
                                          
          GrdIR.SaveRec (i)
          Exit For
        End If
      Next i

End Sub

Private Sub CalcGrdIdValsForWts(ByVal RowNum As Integer)
  Call GrdIdRecalcHelper("IdCustmMtchWt,wIdADiaWt,wIdIDiaWt,wIdZDiaWt,wIdACsWt,wIdICsWt,wIdZCsWt,IdAGldAsWt,IdIGldAsWt,IdZGldAsWt,IdALabAsWt,IdILabAsWt,IdZLabAsWt", RowNum, 0, "")
  
End Sub
Private Sub CalcGrdIdValsForVal(ByVal RowNum As Integer)
  GrdID_ReCalc RowNum, UCase("IdValAddn"), 0, "", ""
  GrdID_ReCalc RowNum, UCase("IdGldLs"), 0, "", ""
  GrdID_ReCalc RowNum, UCase("iDiaDD"), 0, "", ""
  GrdID_ReCalc RowNum, UCase("iDzaDD"), 0, "", ""
        
  GrdIdRecalcHelper UCase("IdCustmMtchWt"), RowNum, 0, ""
  
  Call GrdIdRecalcHelper("IdAVal,IdIVal,IdZVal", RowNum, 0, "")
  GrdID_ReCalc RowNum, UCase("IdIFob"), 0, "", ""
  GrdID_ReCalc RowNum, UCase("IdZFob"), 0, "", ""
End Sub
Private Sub CalcGrdIdVals(ByVal RowNum As Integer)
  CalcGrdIdValsForWts (RowNum)
  CalcGrdIdValsForVal (RowNum)
'  Call CalcGrdIdValsForWts("IdCustmMtchWt,wIdADiaWt,wIdIDiaWt,wIdZDiaWt,wIdACsWt,wIdICsWt,wIdZCsWt,IdAGldAsWt,IdIGldAsWt,IdZGldAsWt,IdALabAsWt,IdILabAsWt,IdZLabAsWt", RowNum, 0, "")
End Sub
Private Sub MakeRm1(ByVal mIdSr As Single, ByVal mExpOdTc As String, ByVal mExpOdYy As String, _
                   ByVal mExpOdChr As String, ByVal mExpOdNo As Single, _
                   ByVal mExpOdSr As Single, ByVal mExpCmCd As String, _
                   ByVal mDmCd As String, ByVal mIdQty As Single, _
                   ByVal mIWtEqOrd As String, mIWtFrOrd As String)
'*** Before 22/11/01
  
''  Dim whZWtEq As String, wDRrTcTyp As String, wCRrTcTyp As String, wCmDiaRtFA As String
''  Dim wCmCSRtFA As String, wDmCtg As String, wTmpTbl As String, wStr As String
''  Dim wRsUniqRm As MDORowSet, wTotRmOWt As Single, wTotRmAWt As Single, wFrOrdYN As String
''  Dim wCtgStr As String, wRsRmDet As MDORowSet, wTotRecs As Single, wRow As Single
''  Dim wOldVal As Single, wRmWt As Single, wRmQty As Single, wMulFact As Single
''  Dim wDFrOrdYN As String, wCFrOrdYN As String, wFinalResultSet As MDORowSet
''  Dim wTotIXtraWt As Single, wTotZXtraWt As Single, wDivFactor As Single
''  Dim wRmRow As Single
''
''  whZWtEq = moCn.GetFldVal("Select hZWtEq from Head where HCd='" + ctSelfCmCd + "'")
''  wCtgStr = "," + moCn.GetFldVal("Select PDesc225 from Param where " + _
''              "PTyp='GLDAS' and PMCd = '" + mIWtFrOrd + "'") + ","
''  wCmDiaRtFA = moCn.GetFldVal("Select CmDiaRtFA from CustMst where CmCtg= 'C' and " + _
''               "CmCd= '" + mExpCmCd + "' ")
''  wCmCSRtFA = moCn.GetFldVal("Select CmCSRtFA from CustMst where CmCtg= 'C' and " + _
''               "CmCd= '" + mExpCmCd + "' ")
''  wDmCtg = moCn.GetFldVal("Select DmCtg from DsgMst where DmTcTyp= 'DM' and DmSz= '' and " + _
''               "DmCd= '" + mDmCd + "' ")
''  wTmpTbl = "#TmpInvRm"
''  wStr = "Select OrSr as qSr, OrSrNo as qSrNo, Rtrim(OrRmCtg) as qRmCtg, Rtrim(OrRmSCtg) as qRmSCtg, " + _
''         "OrRmCd as qRmCd, OrQty*" + CStr(mIdQty) + " as OrQty, " + _
''         "OrWt*" + CStr(mIdQty) + " as qRmAWt, " + _
''         "OrWt*" + CStr(mIdQty) + " as qRmIWt, 9999.999 as qXtraIWt, " + _
''         "OrWt*" + CStr(mIdQty) + " as qRmZWt, 9999.999 as qXtraZWt, OrCstRt as qRmARt, " + _
''         "OrSalRt as qRmIRt, OrSalRt as qRmZRt, OrSetSCd, OrSetCstRt as qSetARt, " + _
''         "OrSetSalRt as qSetIRt, OrSetSalRt as qSetZRt, OrMainMet into " + wTmpTbl + " " + _
''         "from OrdRm " + _
''         "where OrTc= '" + mExpOdTc + "' and OrYy= '" + mExpOdYy + "' and " + _
''         "OrChr= '" + mExpOdChr + "' and OrNo= " + CStr(mExpOdNo) + " and " + _
''         "OrSr= " + CStr(mExpOdSr) + " Order By OrSr"
''
''  '*** Creating Temporary Table #TmpInvRm ***
''  moCn.CreateTmpTable wTmpTbl, wStr
''
''  Set wRsUniqRm = moCn.OpenRes("Select Rtrim(max(OrRmCtg)) as OrRmCtg, OrRmCd, sum(OrWt) as qTotRmOWt " + _
''                  "from OrdRm where OrTc= '" + mExpOdTc + "' and OrYy= '" + mExpOdYy + "' and " + _
''                  "OrChr= '" + mExpOdChr + "' and OrNo= " + CStr(mExpOdNo) + " and " + _
''                  "OrSr= " + CStr(mExpOdSr) + " Group By OrRmCd")
''  Do While Not (wRsUniqRm.EOF Or wRsUniqRm.BOF)
''    wTotRmOWt = wRsUniqRm!qTotRmOWt * mIdQty
''    wTotRmAWt = moCn.GetFldVal("Select sum(FrRmWt* " + _
''                "(Case when F.LocTyp= 'F' and FrFrRmDc= 'D' then 1 " + _
''                "when F.LocTyp='F' and FrFrRmDc= 'C' then -1 " + _
''                "when T.LocTyp='F' and FrToRmDc= 'D' then 1 " + _
''                "when T.LocTyp='F' and FrToRmDc= 'C' then -1 " + _
''                "Else 0 end)) " + _
''                "From InvFgd Join Fgd on FdTc= IfFdTc and FdYy= IfFdYy and FdChr= IfFdChr and " + _
''                "FdNo= IfFdNo and FdSr= IfFdSr join FgRm on FrTc= FdTc and FrYy= FdYy and FrChr= FdChr and " + _
''                "FrNo= FdNo and FrSr= FdSr join Loc F on F.LocCd= FrFrRmLoc " + _
''                "join Loc T on T.LocCd= FrToRmLoc " + _
''                "where IfTc= '" + ADC("wIdTc") + "' and IfYy= '" + ADC("wIdYy") + "' and " + _
''                "IfChr= '" + ADC("wIdChr") + "' and IfNo= " + CStr(ADC("wIdNo")) + " and " + _
''                "IfSr= " + CStr(mIdSr) + " and FrRmCd= '" + wRsUniqRm!OrRmCd + "' " + _
''                "group by FrRmCd ")
''
''    wFrOrdYN = IIf(Trim(mIWtEqOrd) = "Y", "Y", "N")
''    If wFrOrdYN = "N" Then wFrOrdYN = IIf(InStr(1, wCtgStr, "," + wRsUniqRm!OrRmCtg + ",") <> 0, "Y", "N")
''
''    Set wRsRmDet = moCn.OpenResultset("Select * from " + wTmpTbl + " where " + _
''                                   "qRmCd='" + wRsUniqRm!OrRmCd + "' ")
''    wTotRecs = wRsRmDet.RecCount
''    wRsRmDet.MoveFirst
''    wRow = 0
''    Do While Not (wRsRmDet.EOF Or wRsRmDet.BOF)
''      wMulFact = MWLib.Div(wTotRmAWt, wTotRmOWt)
''      wRsRmDet.edit
''      wRow = wRow + 1
''      wOldVal = wRsRmDet!qRmAWt
''      If wRow < wTotRecs Then
''        '*** If Not The Last Record ***
''        wRmWt = wRsRmDet!qRmAWt * wMulFact
''      Else
''        '*** If Last Record ***
''        wRmWt = wTotRmAWt
''      End If
''      wRmQty = wRsRmDet!OrQty * wMulFact
''      wRsRmDet!qRmAWt = wRmWt
''      wRsRmDet!qRmIWt = IIf(wFrOrdYN = "Y", wRsRmDet!qRmIWt, wRmWt)
''      wRsRmDet!qRmZWt = IIf(whZWtEq = "A", wRmWt, IIf(whZWtEq = "I", wRsRmDet!qRmIWt, 0))
''      wRsRmDet!OrQty = IIf(wFrOrdYN = "Y", wRsRmDet!OrQty, wRmQty)
''      wRsRmDet.Update
''
''      wTotRmOWt = wTotRmOWt - wOldVal
''      wTotRmAWt = wTotRmAWt - wRmWt
''      wRsRmDet.MoveNext
''    Loop
''    wRsUniqRm.MoveNext
''  Loop
''
''  moCn.Execute ("Update " + wTmpTbl + " Set qXtraIWt= .000, qXtraZWt= .000 ")
''
''  wDFrOrdYN = IIf(Trim(mIWtEqOrd) = "Y", "Y", "N")
''  If wDFrOrdYN = "N" Then wDFrOrdYN = IIf(InStr(1, wCtgStr, "," + "D" + ",") <> 0, "Y", "N")
''  wCFrOrdYN = IIf(Trim(mIWtEqOrd) = "Y", "Y", "N")
''  If wCFrOrdYN = "N" Then wCFrOrdYN = IIf(InStr(1, wCtgStr, "," + "D" + ",") <> 0, "Y", "N")
''
''  wDRrTcTyp = IIf(wDFrOrdYN = "Y", "XIO", "XIA")
''  wCRrTcTyp = IIf(wCFrOrdYN = "Y", "XIO", "XIA")
''  Call GetXtraWt(wDRrTcTyp, wCRrTcTyp, wTmpTbl, wCmDiaRtFA, wCmCSRtFA, mExpCmCd, wDmCtg, _
''                 "qRmIWt", "qXtraIWt", "qRmIRt", mIdQty)
''
''  If whZWtEq = "A" Then
''    wDRrTcTyp = "XZA"
''    wCRrTcTyp = "XZA"
''    Call GetXtraWt(wDRrTcTyp, wCRrTcTyp, wTmpTbl, wCmDiaRtFA, wCmCSRtFA, mExpCmCd, wDmCtg, _
''                   "qRmZWt", "qXtraZWt", "qRmZRt", mIdQty)
''  End If
''
''  '##### Old Code
''  'Set wFinalResultSet = moCn.OpenRes("Select * from " + wTmpTbl + " Order By OrMainMet")
''  '##### Old Code
''
''  '##### New Code
''  Set wFinalResultSet = moCn.OpenRes("Select max(qSr) as qSr, max(qSrNo) as qSrNo, " + _
''                        "max(qRmCtg) as qRmCtg, max(qRmSCtg) as qRmSCtg, qRmCd, " + _
''                        "sum(OrQty) as OrQty, sum(qRmAWt) as qRmAWt, sum(qRmIWt) as qRmIWt, " + _
''                        "sum(qXtraIWt) as qXtraIWt, sum(qRmZWt) as qRmZWt, " + _
''                        "sum(qXtraZWt) as qXtraZWt, qRmARt, qRmIRt, qRmZRt, OrSetSCd, " + _
''                        "qSetARt, qSetIRt, qSetZRt, OrMainMet from " + wTmpTbl + _
''                        " Group By qRmCd, OrSetSCd, qRmARt, qRmIRt, qRmZRt, qSetARt, " + _
''                        "qSetIRt, qSetZRt, OrMainMet Order By OrMainMet")
''  '#####
''
''  '##### temporary
''  'SELECT IrSr, IrSrNo, IrRmCd, IrSetSCd, str(IrRmARt, 9, 2) as IrRmARt,
''  '  str(IrRmIRt, 9, 2) as IrRmIRt, str(IrRmZRt, 9, 2) as IrRmZRt,
''  '  str(IrSetARt, 9, 2) as IrSetARt,
''  '  str(IrSetIRt, 9, 2) as IrSetIRt, str(IrSetZRt, 9, 2) as IrSetZRt
''  'From INVRM
''  '#####
''
''
''  wTotIXtraWt = 0
''  wTotZXtraWt = 0
''  With GrdIR
''    GrdIR.StartCopy
''    Do While Not (wFinalResultSet.EOF Or wFinalResultSet.BOF)
''      .AddItem
''      wRmRow = .Rows - 1
''      wDivFactor = IIf(wFinalResultSet!qRmCtg = "C" Or wFinalResultSet!qRmCtg = "D", 5, 1)
''      wTotIXtraWt = wTotIXtraWt + MWLib.Div(wFinalResultSet!qXtraIWt, wDivFactor)
''      wTotZXtraWt = wTotZXtraWt + MWLib.Div(wFinalResultSet!qXtraZWt, wDivFactor)
''      .Value(wRmRow, "IrSrNo") = wRmRow
''      .Value(wRmRow, "IrRmCd") = wFinalResultSet!qRmCd
''      .Value(wRmRow, "IrRmQty") = wFinalResultSet!OrQty
''      .Value(wRmRow, "IrRmAWt") = wFinalResultSet!qRmAWt
''      If wFinalResultSet!OrMainMet = "Y" Then
''        .Value(wRmRow, "IrRmIWt") = wFinalResultSet!qRmIWt - wTotIXtraWt
''        If whZWtEq = "A" Then
''          .Value(wRmRow, "IrRmZWt") = wFinalResultSet!qRmZWt - wTotZXtraWt
''        ElseIf whZWtEq = "I" Then
''          .Value(wRmRow, "IrRmZWt") = wFinalResultSet!qRmIWt - wTotIXtraWt
''        End If
''      Else
''        .Value(wRmRow, "IrRmIWt") = wFinalResultSet!qRmIWt + wFinalResultSet!qXtraIWt
''        If whZWtEq = "A" Then
''          .Value(wRmRow, "IrRmZWt") = wFinalResultSet!qRmZWt + wFinalResultSet!qXtraZWt
''        ElseIf whZWtEq = "I" Then
''          .Value(wRmRow, "IrRmZWt") = wFinalResultSet!qRmIWt + wFinalResultSet!qXtraIWt
''        End If
''      End If
''      .Value(wRmRow, "IrRmARt") = wFinalResultSet!qRmARt
''      .Value(wRmRow, "IrRmIRt") = wFinalResultSet!qRmIRt
''      .Value(wRmRow, "IrRmZRt") = wFinalResultSet!qRmZRt
''      .Value(wRmRow, "IrSetSCd") = wFinalResultSet!OrSetSCd
''      .Value(wRmRow, "IrSetARt") = wFinalResultSet!qSetARt
''      .Value(wRmRow, "IrSetIRt") = wFinalResultSet!qSetIRt
''      .Value(wRmRow, "IrSetZRt") = wFinalResultSet!qSetZRt
''      .Value(wRmRow, "IrMainMet") = wFinalResultSet!OrMainMet
''      .SaveRec (wRmRow)
''      wFinalResultSet.MoveNext
''    Loop
''    GrdIR.EndCopy
''    Set wRsUniqRm = Nothing
''    Set wRsRmDet = Nothing
''    Set wFinalResultSet = Nothing
''  End With
End Sub
Private Sub MakeRm2(ByVal mIdSr As Single, ByVal mExpOdTc As String, ByVal mExpOdYy As String, _
                   ByVal mExpOdChr As String, ByVal mExpOdNo As Single, _
                   ByVal mExpOdSr As Single, ByVal mExpCmCd As String, _
                   ByVal mDmCd As String, ByVal mIdQty As Single, _
                   ByVal mIWtEqOrd As String, mIWtFrOrd As String)

'*** Before 27/11/01

''  Dim whZWtEq As String, wDRrTcTyp As String, wCRrTcTyp As String, wCmDiaRtFA As String
''  Dim wCmCSRtFA As String, wDmCtg As String, wTmpTbl As String, wStr As String
''  Dim wRsUniqRm As MDORowSet, wRsTotActWtQty As MDORowSet, wTotRmOWt As Single
''  Dim wTotRmAWt As Single, wTotRmOQty As Single, wTotRmAQty As Single, wFrOrdYN As String
''  Dim wCtgStr As String, wRsRmDet As MDORowSet, wTotRecs As Single, wRow As Single
''  Dim wOldVal As Single, wRmWt As Single, wRmQty As Single, wMulFactWt As Single
''  Dim wMulFactQty As Single, wDFrOrdYN As String, wCFrOrdYN As String
''  Dim wFinalResultSet As MDORowSet, wTotIXtraWt As Single, wTotZXtraWt As Single
''  Dim wDivFactor As Single, wRmRow As Single
''
''  whZWtEq = moCn.GetFldVal("Select hZWtEq from Head where HCd='" + ctSelfCmCd + "'")
''  wCtgStr = "," + moCn.GetFldVal("Select PDesc225 from Param where " + _
''              "PTyp='GLDAS' and PMCd = '" + mIWtFrOrd + "'") + ","
''  wCmDiaRtFA = moCn.GetFldVal("Select CmDiaRtFA from CustMst where CmCtg= 'C' and " + _
''               "CmCd= '" + mExpCmCd + "' ")
''  wCmCSRtFA = moCn.GetFldVal("Select CmCSRtFA from CustMst where CmCtg= 'C' and " + _
''               "CmCd= '" + mExpCmCd + "' ")
''  wDmCtg = moCn.GetFldVal("Select DmCtg from DsgMst where DmTcTyp= 'DM' and DmSz= '' and " + _
''               "DmCd= '" + mDmCd + "' ")
''  wTmpTbl = "#TmpInvRm"
''  wStr = "Select OrSr as qSr, OrSrNo as qSrNo, Rtrim(OrRmCtg) as qRmCtg, Rtrim(OrRmSCtg) as qRmSCtg, " + _
''         "OrRmCd as qRmCd, OrQty*" + CStr(mIdQty) + " as OrQty, " + _
''         "OrWt*" + CStr(mIdQty) + " as qRmAWt, " + _
''         "OrWt*" + CStr(mIdQty) + " as qRmIWt, 9999.999 as qXtraIWt, " + _
''         "OrWt*" + CStr(mIdQty) + " as qRmZWt, 9999.999 as qXtraZWt, OrCstRt as qRmARt, " + _
''         "OrSalRt as qRmIRt, OrSalRt as qRmZRt, OrSetSCd, OrSetCstRt as qSetARt, " + _
''         "OrSetSalRt as qSetIRt, OrSetSalRt as qSetZRt, OrMainMet into " + wTmpTbl + " " + _
''         "from OrdRm " + _
''         "where OrTc= '" + mExpOdTc + "' and OrYy= '" + mExpOdYy + "' and " + _
''         "OrChr= '" + mExpOdChr + "' and OrNo= " + CStr(mExpOdNo) + " and " + _
''         "OrSr= " + CStr(mExpOdSr) + " Order By OrSr"
''
''  '*** Creating Temporary Table #TmpInvRm ***
''  moCn.CreateTmpTable wTmpTbl, wStr
''
''  Set wRsUniqRm = moCn.OpenRes("Select Rtrim(max(OrRmCtg)) as OrRmCtg, OrRmCd, sum(OrWt) as qTotRmOWt, sum(OrQty) as qTotRmOQty " + _
''                  "from OrdRm where OrTc= '" + mExpOdTc + "' and OrYy= '" + mExpOdYy + "' and " + _
''                  "OrChr= '" + mExpOdChr + "' and OrNo= " + CStr(mExpOdNo) + " and " + _
''                  "OrSr= " + CStr(mExpOdSr) + " Group By OrRmCd")
''  Do While Not (wRsUniqRm.EOF Or wRsUniqRm.BOF)
''    wTotRmOWt = wRsUniqRm!qTotRmOWt * mIdQty
''    wTotRmOQty = wRsUniqRm!qTotRmOQty * mIdQty
''    Set wRsTotActWtQty = moCn.OpenRes("Select sum(FrRmWt* " + _
''                "(Case when F.LocTyp= 'F' and FrFrRmDc= 'D' then 1 " + _
''                "when F.LocTyp='F' and FrFrRmDc= 'C' then -1 " + _
''                "when T.LocTyp='F' and FrToRmDc= 'D' then 1 " + _
''                "when T.LocTyp='F' and FrToRmDc= 'C' then -1 " + _
''                "Else 0 end)) as qTotAWt, " + _
''                "sum(FrRmQty* (Case when F.LocTyp= 'F' and FrFrRmDc= 'D' then 1 " + _
''                "when F.LocTyp='F' and FrFrRmDc= 'C' then -1 " + _
''                "when T.LocTyp='F' and FrToRmDc= 'D' then 1 " + _
''                "when T.LocTyp='F' and FrToRmDc= 'C' then -1 " + _
''                "Else 0 end)) as qTotAQty " + _
''                "From InvFgd Join Fgd on FdTc= IfFdTc and FdYy= IfFdYy and FdChr= IfFdChr and " + _
''                "FdNo= IfFdNo and FdSr= IfFdSr join FgRm on FrTc= FdTc and FrYy= FdYy and FrChr= FdChr and " + _
''                "FrNo= FdNo and FrSr= FdSr join Loc F on F.LocCd= FrFrRmLoc " + _
''                "join Loc T on T.LocCd= FrToRmLoc " + _
''                "where IfTc= '" + ADC("wIdTc") + "' and IfYy= '" + ADC("wIdYy") + "' and " + _
''                "IfChr= '" + ADC("wIdChr") + "' and IfNo= " + CStr(ADC("wIdNo")) + " and " + _
''                "IfSr= " + CStr(mIdSr) + " and FrRmCd= '" + wRsUniqRm!OrRmCd + "' " + _
''                "group by FrRmCd ")
''    If wRsTotActWtQty.RecCount > 0 Then
''      wTotRmAWt = wRsTotActWtQty!qTotAWt
''      wTotRmAQty = wRsTotActWtQty!qTotAQty
''    End If
''    wFrOrdYN = IIf(Trim(mIWtEqOrd) = "Y", "Y", "N")
''    If wFrOrdYN = "N" Then wFrOrdYN = IIf(InStr(1, wCtgStr, "," + wRsUniqRm!OrRmCtg + ",") <> 0, "Y", "N")
''
''    '#####
''    If (wCmDiaRtFA = "F" And UCase(Trim(wRsUniqRm!OrRmCtg)) = "D") Or _
''       (wCmCSRtFA = "F" And UCase(Trim(wRsUniqRm!OrRmCtg)) = "C") Then
''      If wTotRmAWt >= wTotRmOWt Then
''        Set wRsRmDet = moCn.OpenResultset("Select * from " + wTmpTbl + " where " + _
''                                   "qRmCd='" + wRsUniqRm!OrRmCd + "' Order By qRmIRt asc, qSetIRt asc")
''      Else
''        Set wRsRmDet = moCn.OpenResultset("Select * from " + wTmpTbl + " where " + _
''                                   "qRmCd='" + wRsUniqRm!OrRmCd + "' Order By qRmIRt desc, qSetIRt desc")
''      End If
''
''      wTotRecs = wRsRmDet.RecCount
''      wRsRmDet.MoveFirst
''      wRow = 0
''      Do While Not (wRsRmDet.EOF Or wRsRmDet.BOF)
''        wRsRmDet.edit
''        wRow = wRow + 1
''        If wRow < wTotRecs Then
''          '*** If Not The Last Record ***
''          wRmWt = IIf(wRsRmDet!qRmAWt < wTotRmAWt, wRsRmDet!qRmAWt, wTotRmAWt)
''          wRmQty = IIf(wRsRmDet!OrQty < wTotRmAQty, wRsRmDet!OrQty, wTotRmAQty)
''        Else
''          '*** If Last Record ***
''          wRmWt = wTotRmAWt
''          wRmQty = wTotRmAQty
''        End If
''        wRsRmDet!qRmAWt = wRmWt
''        wRsRmDet!qRmIWt = IIf(wFrOrdYN = "Y", wRsRmDet!qRmIWt, wRmWt)
''        wRsRmDet!qRmZWt = IIf(whZWtEq = "A", wRmWt, IIf(whZWtEq = "I", wRsRmDet!qRmIWt, 0))
''        wRsRmDet!OrQty = wRmQty   '''IIf(wFrOrdYN = "Y", wRsRmDet!OrQty, wRmQty)
''        wRsRmDet.Update
''        wTotRmAWt = wTotRmAWt - wRmWt
''        wTotRmAQty = wTotRmAQty - wRmQty
''        wRsRmDet.MoveNext
''      Loop
''    Else
''      Set wRsRmDet = moCn.OpenResultset("Select * from " + wTmpTbl + " where " + _
''                                     "qRmCd='" + wRsUniqRm!OrRmCd + "' ")
''      wTotRecs = wRsRmDet.RecCount
''      wRsRmDet.MoveFirst
''      wRow = 0
''      Do While Not (wRsRmDet.EOF Or wRsRmDet.BOF)
''        wMulFactWt = MWLib.Div(wTotRmAWt, wTotRmOWt)
''        wMulFactQty = MWLib.Div(wTotRmAQty, wTotRmOQty)
''        wRsRmDet.edit
''        wRow = wRow + 1
''        wOldVal = wRsRmDet!qRmAWt
''        If wRow < wTotRecs Then
''          '*** If Not The Last Record ***
''          wRmWt = wRsRmDet!qRmAWt * wMulFactWt
''          wRmQty = wRsRmDet!OrQty * wMulFactQty
''        Else
''          '*** If Last Record ***
''          wRmWt = wTotRmAWt
''          wRmQty = wTotRmAQty
''        End If
''        wRsRmDet!qRmAWt = wRmWt
''        wRsRmDet!qRmIWt = IIf(wFrOrdYN = "Y", wRsRmDet!qRmIWt, wRmWt)
''        wRsRmDet!qRmZWt = IIf(whZWtEq = "A", wRmWt, IIf(whZWtEq = "I", wRsRmDet!qRmIWt, 0))
''        wRsRmDet!OrQty = IIf(wFrOrdYN = "Y", wRsRmDet!OrQty, wRmQty)
''        wRsRmDet.Update
''
''        wTotRmOWt = wTotRmOWt - wOldVal
''        wTotRmAWt = wTotRmAWt - wRmWt
''        wTotRmOQty = wTotRmOWt - wOldVal
''        wTotRmAQty = wTotRmAQty - wRmQty
''        wRsRmDet.MoveNext
''      Loop
''    End If
''    '#####
''
''    'Old##### Set wRsRmDet = moCn.OpenResultset("Select * from " + wTmpTbl + " where " + _
''                                   "qRmCd='" + wRsUniqRm!OrRmCd + "' ")
'''    wTotRecs = wRsRmDet.RecCount
'''    wRsRmDet.MoveFirst
'''    wRow = 0
'''    Do While Not (wRsRmDet.EOF Or wRsRmDet.BOF)
'''      wMulFact = MWLib.Div(wTotRmAWt, wTotRmOWt)
'''      wRsRmDet.edit
'''      wRow = wRow + 1
'''      wOldVal = wRsRmDet!qRmAWt
'''      If wRow < wTotRecs Then
'''        '*** If Not The Last Record ***
'''        wRmWt = wRsRmDet!qRmAWt * wMulFact
'''      Else
'''        '*** If Last Record ***
'''        wRmWt = wTotRmAWt
'''      End If
'''      wRmQty = wRsRmDet!OrQty * wMulFact
'''      wRsRmDet!qRmAWt = wRmWt
'''      wRsRmDet!qRmIWt = IIf(wFrOrdYN = "Y", wRsRmDet!qRmIWt, wRmWt)
'''      wRsRmDet!qRmZWt = IIf(whZWtEq = "A", wRmWt, IIf(whZWtEq = "I", wRsRmDet!qRmIWt, 0))
'''      wRsRmDet!OrQty = IIf(wFrOrdYN = "Y", wRsRmDet!OrQty, wRmQty)
'''      wRsRmDet.Update
'''
'''      wTotRmOWt = wTotRmOWt - wOldVal
'''      wTotRmAWt = wTotRmAWt - wRmWt
'''      wRsRmDet.MoveNext
'''    Loop
''    wRsUniqRm.MoveNext
''  Loop
''
''  moCn.Execute ("Update " + wTmpTbl + " Set qXtraIWt= .000, qXtraZWt= .000 ")
''
''  wDFrOrdYN = IIf(Trim(mIWtEqOrd) = "Y", "Y", "N")
''  If wDFrOrdYN = "N" Then wDFrOrdYN = IIf(InStr(1, wCtgStr, "," + "D" + ",") <> 0, "Y", "N")
''  wCFrOrdYN = IIf(Trim(mIWtEqOrd) = "Y", "Y", "N")
''  If wCFrOrdYN = "N" Then wCFrOrdYN = IIf(InStr(1, wCtgStr, "," + "D" + ",") <> 0, "Y", "N")
''
''  wDRrTcTyp = IIf(wDFrOrdYN = "Y", "XIO", "XIA")
''  wCRrTcTyp = IIf(wCFrOrdYN = "Y", "XIO", "XIA")
''  Call GetXtraWt(wDRrTcTyp, wCRrTcTyp, wTmpTbl, wCmDiaRtFA, wCmCSRtFA, mExpCmCd, wDmCtg, _
''                 "qRmIWt", "qXtraIWt", "qRmIRt", mIdQty)
''
''  If whZWtEq = "A" Then
''    wDRrTcTyp = "XZA"
''    wCRrTcTyp = "XZA"
''    Call GetXtraWt(wDRrTcTyp, wCRrTcTyp, wTmpTbl, wCmDiaRtFA, wCmCSRtFA, mExpCmCd, wDmCtg, _
''                   "qRmZWt", "qXtraZWt", "qRmZRt", mIdQty)
''  End If
''
''  '##### Old Code
''  'Set wFinalResultSet = moCn.OpenRes("Select * from " + wTmpTbl + " Order By OrMainMet")
''  '##### Old Code
''
''  '##### New Code
''  Set wFinalResultSet = moCn.OpenRes("Select max(qSr) as qSr, max(qSrNo) as qSrNo, " + _
''                        "max(qRmCtg) as qRmCtg, max(qRmSCtg) as qRmSCtg, qRmCd, " + _
''                        "sum(OrQty) as OrQty, sum(qRmAWt) as qRmAWt, sum(qRmIWt) as qRmIWt, " + _
''                        "sum(qXtraIWt) as qXtraIWt, sum(qRmZWt) as qRmZWt, " + _
''                        "sum(qXtraZWt) as qXtraZWt, qRmARt, qRmIRt, qRmZRt, OrSetSCd, " + _
''                        "qSetARt, qSetIRt, qSetZRt, OrMainMet from " + wTmpTbl + _
''                        " Group By qRmCd, OrSetSCd, qRmARt, qRmIRt, qRmZRt, qSetARt, " + _
''                        "qSetIRt, qSetZRt, OrMainMet Order By OrMainMet")
''  '#####
''
''  '##### temporary
''  'SELECT IrSr, IrSrNo, IrRmCd, IrSetSCd, str(IrRmARt, 9, 2) as IrRmARt,
''  '  str(IrRmIRt, 9, 2) as IrRmIRt, str(IrRmZRt, 9, 2) as IrRmZRt,
''  '  str(IrSetARt, 9, 2) as IrSetARt,
''  '  str(IrSetIRt, 9, 2) as IrSetIRt, str(IrSetZRt, 9, 2) as IrSetZRt
''  'From INVRM
''  '#####
''
''
''  wTotIXtraWt = 0
''  wTotZXtraWt = 0
''  With GrdIR
''    GrdIR.StartCopy
''    Do While Not (wFinalResultSet.EOF Or wFinalResultSet.BOF)
''      .AddItem
''      wRmRow = .Rows - 1
''      wDivFactor = IIf(wFinalResultSet!qRmCtg = "C" Or wFinalResultSet!qRmCtg = "D", 5, 1)
''      wTotIXtraWt = wTotIXtraWt + MWLib.Div(wFinalResultSet!qXtraIWt, wDivFactor)
''      wTotZXtraWt = wTotZXtraWt + MWLib.Div(wFinalResultSet!qXtraZWt, wDivFactor)
''      .Value(wRmRow, "IrSrNo") = wRmRow
''      .Value(wRmRow, "IrRmCd") = wFinalResultSet!qRmCd
''      .Value(wRmRow, "IrRmQty") = wFinalResultSet!OrQty
''      .Value(wRmRow, "IrRmAWt") = wFinalResultSet!qRmAWt
''      If wFinalResultSet!OrMainMet = "Y" Then
''        .Value(wRmRow, "IrRmIWt") = wFinalResultSet!qRmIWt - wTotIXtraWt
''        If whZWtEq = "A" Then
''          .Value(wRmRow, "IrRmZWt") = wFinalResultSet!qRmZWt - wTotZXtraWt
''        ElseIf whZWtEq = "I" Then
''          .Value(wRmRow, "IrRmZWt") = wFinalResultSet!qRmIWt - wTotIXtraWt
''        End If
''      Else
''        .Value(wRmRow, "IrRmIWt") = wFinalResultSet!qRmIWt + wFinalResultSet!qXtraIWt
''        If whZWtEq = "A" Then
''          .Value(wRmRow, "IrRmZWt") = wFinalResultSet!qRmZWt + wFinalResultSet!qXtraZWt
''        ElseIf whZWtEq = "I" Then
''          .Value(wRmRow, "IrRmZWt") = wFinalResultSet!qRmIWt + wFinalResultSet!qXtraIWt
''        End If
''      End If
''      .Value(wRmRow, "IrRmARt") = wFinalResultSet!qRmARt
''      .Value(wRmRow, "IrRmIRt") = wFinalResultSet!qRmIRt
''      .Value(wRmRow, "IrRmZRt") = wFinalResultSet!qRmZRt
''      .Value(wRmRow, "IrSetSCd") = wFinalResultSet!OrSetSCd
''      .Value(wRmRow, "IrSetARt") = wFinalResultSet!qSetARt
''      .Value(wRmRow, "IrSetIRt") = wFinalResultSet!qSetIRt
''      .Value(wRmRow, "IrSetZRt") = wFinalResultSet!qSetZRt
''      .Value(wRmRow, "IrMainMet") = wFinalResultSet!OrMainMet
''      .SaveRec (wRmRow)
''      wFinalResultSet.MoveNext
''    Loop
''    GrdIR.EndCopy
''    Set wRsUniqRm = Nothing
''    Set wRsTotActWtQty = Nothing
''    Set wRsRmDet = Nothing
''    Set wFinalResultSet = Nothing
''  End With
End Sub

Private Sub MakeRm(ByVal mIdSr As Single, ByVal mExpOdTc As String, ByVal mExpOdYy As String, _
                   ByVal mExpOdChr As String, ByVal mExpOdNo As Single, _
                   ByVal mExpOdSr As Single, ByVal mExpCmCd As String, _
                   ByVal mDmCd As String, ByVal mIdQty As Double, _
                   ByVal mIWtEqOrd As String, mIWtFrOrd As String, mIGldRtEqOrd As String)
  '*** This procedure is used to write records into the Raw Material Grid
      'The records are initially written into a temporary table
      'This is for the ease in distributing the weights among the various Rm Codes
      'based on the different factors like:
      '1.   Diamond Rate is favourable or average (in case of favourable max Wt will go to
      '     the RmCode record with the max Rate)
      '     i.e. If the Actual Wt is more than the order Wt for a particular RmCode then
      '     the extra Wt will be assigned to the Rec with the max rate
      '     If the Actual Wt is less than the order Wt for a particular RmCode then
      '     the Rec with the max rate will be assigned the Wt mentioned in the order and
      '     the record with the minimum rate will have to suffer the loss
      '2.   Colour Stone Rate is favourable or average
      '3.   Customs Wt is Equal to Invoice Wt or not
      '4.   All Wts are to be taken from order or only Weights of some Rm categories has to
      '     be taken from Order (only for Invoice & Customs)
      '     When some weights are taken from the order, the Gross Wt is adjusted by adding
      '     or subtracting the difference weights to the Main Metal Weight.
      '     i.e. the Inv& Customs Gross Wt is made equal to the Actual Gross Wt
      '     But if the Main Metal Wt is taken from the Order then no such adjustment
      '     is possible
      '5.   Application of various factors like 'XIA', 'XIO', 'XZA'
      '     XIA - Application of Extra Wt if Invoice Weights are taken from Actuals
      '     XIO - Application of Extra Wt if Invoice Weights are taken from Order
      '     XZA - Application of Extra Wt if Customs Weights are taken from Actuals
  '*** The records are written as one line per RmCode, Setting Sub Code, Rm Actual Rate,
      'Rm Inv Rate, Rm Customs Rate, Setting Actual Rate, Setting Inv Rate,
      'Setting Customs Rate, Main Metal
  
  Dim whZWtEq As String, wDRrTcTyp As String, wCRrTcTyp As String, wCmDiaRtFA As String
  Dim wCmCsRtFA As String, wDmCtg As String, wTmpTbl As String, wStr As String
  Dim wRsUniqRm As MDORowSet, wRsTotActWtQty As MDORowSet, wTotRmOWt As Double
  Dim wTotRmAWt As Double, wTotRmOQty As Double, wTotRmAQty As Double, wFrOrdYN As String
  Dim wCtgStr As String, wRsRmDet As MDORowSet, wTotRecs As Single, wRow As Single
  Dim wOldVal As Double, wOldRmQty As Double, wRmWt As Double, wRmQty As Double, wMulFactWt As Single
  Dim wMulFactQty As Single, wDFrOrdYN As String, wCFrOrdYN As String
  Dim wFinalResultSet As MDORowSet, wTotIXtraWt As Single, wTotZXtraWt As Single
  Dim wDivFactor As Single, wRmRow As Single
  
  Dim wTotAWt As Single, wTotIWt As Single, wTotZWt As Single, wWtFrOrdYN As String
  Dim wf_AddLab As Single, wLabRow As Integer, wf_CnvRt As Single
  
  Dim wf_CustmCdPurityZ As Double   '*** (Jen 2.13)


  whZWtEq = moCn.GetFldVal("Select hZWtEq from Head where HCoCd= '" + gs_CoCd + "' and HCd='" + ctSelfCmCd + "'")
  'wCtgStr = "," + moCn.GetFldVal("Select PDesc225 from Param where " + _
              "PTyp='GLDAS' and PMCd = '" + mIWtFrOrd + "'") + ","
  wCtgStr = "," + moCn.GetFldVal("Select PDesc225 from Param where " + _
              "PTyp='IWTFRORD' and PMCd = '" + mIWtFrOrd + "'") + ","
  wCmDiaRtFA = moCn.GetFldVal("Select CmDiaRtFA from CustMst where CmCtg= 'C' and " + _
               "CmCd= '" + mExpCmCd + "' ")
  wCmCsRtFA = moCn.GetFldVal("Select CmCSRtFA from CustMst where CmCtg= 'C' and " + _
               "CmCd= '" + mExpCmCd + "' ")
  '****** Sachin 3.02 31-12-07 - prtkey added
  wDmCtg = moCn.GetFldVal("Select DmCtg from DsgMst where DmTcTyp= 'DM' and " + _
               "DmCd= '" + mDmCd + "' and DmSz= '' and DmPrtKey='" + ctCurrPrtn + "' ")
  wTmpTbl = "#TmpInvRm"
  If UCase(ms_IdTcTyp) = UCase(ctInTcTyp) Then
    '****** Sachin 3.02 31-12-07 - prtkey added
     wStr = "Select OrSr as qSr, OrSrNo as qSrNo, Rtrim(OrRmCtg) as qRmCtg, Rtrim(OrRmSCtg) as qRmSCtg, " + _
           "OrRmCd as qRmCd, OrQty*" + CStr(mIdQty) + " as OrQty, " + _
           "OrWt*" + CStr(mIdQty) + " as qRmAWt, " + _
           "OrWt*" + CStr(mIdQty) + " as qRmIWt, 9999.999 as qXtraIWt, " + _
           "OrWt*" + CStr(mIdQty) + " as qRmZWt, 9999.999 as qXtraZWt, OrCstRt as qRmARt, " + _
           "OrSalRt as qRmIRt, OrSalRt as qRmZRt, OrSetSCd, OrSetCstRt as qSetARt, " + _
           "OrSetSalRt as qSetIRt, OrSetSalRt as qSetZRt,OrMainMet,OrCustRmCd into " + wTmpTbl + " " + _
           "from OrdRm " + _
           "where OrCoCd= '" + gs_CoCd + "' and OrTc= '" + mExpOdTc + "' and OrYy= '" + mExpOdYy + "'" + _
           " and OrChr= '" + mExpOdChr + "' and OrNo= " + CStr(mExpOdNo) + " and " + _
           "OrSr= " + CStr(mExpOdSr) + " and OrPrtKey='" + ctCurrPrtn + "' Order By OrSr"
           
  ElseIf UCase(ms_IdTcTyp) = UCase(ctCTBTcTyp) Then
    '****** Sachin 3.02 31-12-07 - prtkey added
    wStr = "Select OrSr as qSr, OrSrNo as qSrNo, Rtrim(OrRmCtg) as qRmCtg, Rtrim(OrRmSCtg) as qRmSCtg, " + _
           "OrRmCd as qRmCd, OrQty* " + CStr(mIdQty) + " as OrQty, " + _
           "OrWt* " + CStr(mIdQty) + " as qRmAWt, " + _
           "OrWt* " + CStr(mIdQty) + " as qRmIWt, 0 as qXtraIWt, " + _
           "OrWt* " + CStr(mIdQty) + " as qRmZWt, 0 as qXtraZWt, 0 as qRmARt, " + _
           "0 as qRmIRt, 0 as qRmZRt, OrSetSCd, OrSetCstRt as qSetARt, " + _
           "0 as qSetIRt, 0 as qSetZRt,OrMainMet,OrCustRmCd into " + wTmpTbl + " " + _
           "from OrdRm " + _
           "where OrCoCd= '" + gs_CoCd + "' and OrTc= '" + mExpOdTc + "' and OrYy= '" + mExpOdYy + "'" + _
           " and OrChr= '" + mExpOdChr + "' and OrNo= " + CStr(mExpOdNo) + " and " + _
           "OrSr= " + CStr(mExpOdSr) + " and OrPrtKey='" + ctCurrPrtn + "' Order By OrSr"
           
           'urmi 28/02/04
           '"OrSr= " + CStr(mExpOdSr) + " and OrWt> 0 Order By OrSr"
  End If
  '*** Creating Temporary Table #TmpInvRm ***
  
'  TrackStatus ("MakeRm - Creating TmpTable ")
  moCn.CreateTmpTable wTmpTbl, wStr
'  TrackStatus ("MakeRm - Done Creating TmpTable ")
  '****** Sachin 3.02 31-12-07 - prtkey added
  Set wRsUniqRm = moCn.OpenRes("Select Rtrim(max(OrRmCtg)) as OrRmCtg, OrRmCd, sum(OrWt) as qTotRmOWt, sum(OrQty) as qTotRmOQty " + _
                  "from OrdRm where OrCoCd= '" + gs_CoCd + "' and OrTc= '" + mExpOdTc + "' and " + _
                  "OrYy= '" + mExpOdYy + "' and OrChr= '" + mExpOdChr + "' and " + _
                  "OrNo= " + CStr(mExpOdNo) + " and OrSr= " + CStr(mExpOdSr) + _
                  " and OrPrtKey='" + ctCurrPrtn + "' Group By OrRmCd")

' " and OrWt> 0 Group By OrRmCd")
  Do While Not (wRsUniqRm.EOF Or wRsUniqRm.BOF)
    wTotRmOWt = wRsUniqRm!qTotRmOWt * mIdQty
    wTotRmOQty = wRsUniqRm!qTotRmOQty * mIdQty
    
    '###### Join Of Fgd Not required
    '****** Sachin 3.02 - Id fields In Joins
    Set wRsTotActWtQty = moCn.OpenRes("Select sum(FrRmWt* " + _
                "(Case when F.LocTyp= 'F' and FrFrRmDc= 'D' then 1 " + _
                "when F.LocTyp= 'F' and FrFrRmDc= 'C' then -1 " + _
                "when T.LocTyp= 'F' and FrToRmDc= 'D' then 1 " + _
                "when T.LocTyp= 'F' and FrToRmDc= 'C' then -1 " + _
                "Else 0 end)) as qTotAWt, " + _
                "sum(FrRmQty* (Case when F.LocTyp= 'F' and FrFrRmDc= 'D' then 1 " + _
                "when F.LocTyp= 'F' and FrFrRmDc= 'C' then -1 " + _
                "when T.LocTyp= 'F' and FrToRmDc= 'D' then 1 " + _
                "when T.LocTyp= 'F' and FrToRmDc= 'C' then -1 " + _
                "Else 0 end)) as qTotAQty " + _
                "From InvFgd Join Fgd on FdIdNo=IfFdIdNo and FdCoCd= IfCoCd and  FdTc= IfFdTc and FdYy= IfFdYy " + _
                "and FdChr= IfFdChr and FdNo= IfFdNo and FdSr= IfFdSr and FdPrtKey=IfPrtKey " + _
                "join FgRm on FrFdIdNo=FdIdNo and FrCoCd= FdCoCd and FrTc= FdTc and FrYy= FdYy and " + _
                "FrChr= FdChr and FrNo= FdNo and FrSr= FdSr and FrPrtKey=FdPrtKey " + _
                "join Loc F on F.LocCoCd= FrCoCd and F.LocCd= FrFrRmLoc " + _
                "join Loc T on T.LocCoCd= FrCoCd and T.LocCd= FrToRmLoc " + _
                "where IfCoCd= '" + gs_CoCd + "' and " + _
                "IfTc= '" + ADC("wIdTc") + "' and IfYy= '" + ADC("wIdYy") + "' and " + _
                "IfChr= '" + ADC("wIdChr") + "' and IfNo= " + CStr(ADC("wIdNo")) + " and " + _
                "IfSr= " + CStr(mIdSr) + " and IfPrtKey='" + ctCurrPrtn + "' and FrRmCd= '" + wRsUniqRm!OrRmCd + "' " + _
                "group by FrRmCd ")
    If wRsTotActWtQty.RecCount > 0 Then
      wTotRmAWt = wRsTotActWtQty!qTotAWt
      wTotRmAQty = wRsTotActWtQty!qTotAQty
    End If
    wFrOrdYN = IIF(Trim(mIWtEqOrd) = "Y", "Y", "N")
    If wFrOrdYN = "N" Then wFrOrdYN = IIF(InStr(1, wCtgStr, "," + wRsUniqRm!OrRmCtg + ",") <> 0, "Y", "N")

    If (UCase(ms_IdTcTyp) = UCase(ctInTcTyp)) And ((wCmDiaRtFA = "F" And UCase(Trim(wRsUniqRm!OrRmCtg)) = "D") Or _
       (wCmCsRtFA = "F" And UCase(Trim(wRsUniqRm!OrRmCtg)) = "C")) Then
      If wTotRmAWt >= wTotRmOWt Then
        Set wRsRmDet = moCn.OpenResultset("Select * from " + wTmpTbl + " where " + _
                                   "qRmCd= '" + wRsUniqRm!OrRmCd + "' Order By qRmIRt asc, qSetIRt asc")
      Else
        Set wRsRmDet = moCn.OpenResultset("Select * from " + wTmpTbl + " where " + _
                                   "qRmCd= '" + wRsUniqRm!OrRmCd + "' Order By qRmIRt desc, qSetIRt desc")
      End If
    
      wTotRecs = wRsRmDet.RecCount
      If wTotRecs > 0 Then wRsRmDet.MoveFirst
      wRow = 0
      Do While Not (wRsRmDet.EOF Or wRsRmDet.BOF)
        wRsRmDet.edit
        wRow = wRow + 1
        If wRow < wTotRecs Then
          '*** If Not The Last Record ***
          wRmWt = IIF(wRsRmDet!qRmAWt < wTotRmAWt, wRsRmDet!qRmAWt, wTotRmAWt)
          'wRmQty = IIf(wRsRmDet!OrQty < wTotRmAQty, wRsRmDet!OrQty, wTotRmAQty)   '*** Bef 20/04/02
          wRmQty = IIF(wRsRmDet!OrQty < wTotRmAQty And Round(wTotRmAWt - wRmWt, 3) > 0, wRsRmDet!OrQty, wTotRmAQty) '*** After 20/04/02
        Else
          '*** If Last Record ***
          wRmWt = wTotRmAWt
          wRmQty = wTotRmAQty
        End If
        wRsRmDet!qRmAWt = wRmWt
        wRsRmDet!qRmIWt = IIF(wFrOrdYN = "Y", wRsRmDet!qRmIWt, wRmWt)
        wRsRmDet!qRmZWt = IIF(whZWtEq = "A", wRmWt, IIF(whZWtEq = "I", wRsRmDet!qRmIWt, 0))
        wRsRmDet!OrQty = wRmQty   '''IIf(wFrOrdYN = "Y", wRsRmDet!OrQty, wRmQty)
        wRsRmDet.Update
        wTotRmAWt = wTotRmAWt - wRmWt
        'Geeta****2.11
        'wTotRmAQty = wTotRmAQty - wRmQty
        wTotRmAQty = wTotRmAQty - wRsRmDet!OrQty
        wRsRmDet.MoveNext
      Loop
    Else
      Set wRsRmDet = moCn.OpenResultset("Select * from " + wTmpTbl + " where " + _
                                     "qRmCd='" + wRsUniqRm!OrRmCd + "' ")
      wTotRecs = wRsRmDet.RecCount
      If wTotRecs > 0 Then wRsRmDet.MoveFirst
      wRow = 0
      Do While Not (wRsRmDet.EOF Or wRsRmDet.BOF)
        wMulFactWt = MWLib.Div(wTotRmAWt, wTotRmOWt)
        wMulFactQty = MWLib.Div(wTotRmAQty, wTotRmOQty)
        wRsRmDet.edit
        wRow = wRow + 1
        wOldVal = wRsRmDet!qRmAWt
        wOldRmQty = wRsRmDet!OrQty
        If wRow < wTotRecs Then
          '*** If Not The Last Record ***
          wRmWt = wRsRmDet!qRmAWt * wMulFactWt
          wRmQty = wRsRmDet!OrQty * wMulFactQty
        Else
          '*** If Last Record ***
          wRmWt = wTotRmAWt
          wRmQty = wTotRmAQty
        End If
        wRsRmDet!qRmAWt = wRmWt
        
        If UCase(ms_IdTcTyp) = UCase(ctInTcTyp) Then
          wRsRmDet!qRmIWt = IIF(wFrOrdYN = "Y", wRsRmDet!qRmIWt, wRmWt)
          wRsRmDet!qRmZWt = IIF(whZWtEq = "A", wRmWt, IIF(whZWtEq = "I", wRsRmDet!qRmIWt, 0))
        ElseIf UCase(ms_IdTcTyp) = UCase(ctCTBTcTyp) Then
          wRsRmDet!qRmIWt = wRmWt
          wRsRmDet!qRmZWt = wRmWt
        End If
        
        'wRsRmDet!OrQty = IIF(wFrOrdYN = "Y", wRsRmDet!OrQty, IIF(wRmQty > 0, wRmQty, 0))   '*** (Bef 2.14 Next Patch 1)
        '*** (Jen 2.14 Next Patch 1)
        wRsRmDet!OrQty = IIF(wRmQty > 0, wRmQty, 0)
        '*** (Jen 2.14 Next Patch 1)

        wRsRmDet.Update
  
        wTotRmOWt = wTotRmOWt - wOldVal
        wTotRmAWt = wTotRmAWt - wRmWt
        wTotRmOQty = wTotRmOQty - wOldRmQty
        'Geeta***2.11
       ' wTotRmAQty = wTotRmAQty - wRmQty
        wTotRmAQty = wTotRmAQty - wRsRmDet!OrQty
        wRsRmDet.MoveNext
      Loop
    End If
    
    wRsUniqRm.MoveNext
  Loop

  If UCase(ms_IdTcTyp) = UCase(ctInTcTyp) Then
    moCn.Execute ("Update " + wTmpTbl + " Set qXtraIWt= .000, qXtraZWt= .000 ")
    
    wDFrOrdYN = IIF(Trim(mIWtEqOrd) = "Y", "Y", "N")
    If wDFrOrdYN = "N" Then wDFrOrdYN = IIF(InStr(1, wCtgStr, "," + "D" + ",") <> 0, "Y", "N")
    wCFrOrdYN = IIF(Trim(mIWtEqOrd) = "Y", "Y", "N")
    If wCFrOrdYN = "N" Then wCFrOrdYN = IIF(InStr(1, wCtgStr, "," + "C" + ",") <> 0, "Y", "N")
    
    wDRrTcTyp = IIF(wDFrOrdYN = "Y", "XIO", "XIA")
    wCRrTcTyp = IIF(wCFrOrdYN = "Y", "XIO", "XIA")
    Call GetXtraWt(wDRrTcTyp, wCRrTcTyp, wTmpTbl, wCmDiaRtFA, wCmCsRtFA, mExpCmCd, wDmCtg, _
                   "qRmIWt", "qXtraIWt", "qRmIRt", mIdQty)
  
    If whZWtEq = "A" Then
      wDRrTcTyp = "XZA"
      wCRrTcTyp = "XZA"
      Call GetXtraWt(wDRrTcTyp, wCRrTcTyp, wTmpTbl, wCmDiaRtFA, wCmCsRtFA, mExpCmCd, wDmCtg, _
                     "qRmZWt", "qXtraZWt", "qRmZRt", mIdQty)
    End If
  End If
  
  '****** Sachin 3.02 31-12-07 - prtkey added
  Set wFinalResultSet = moCn.OpenRes("Select max(qSr) as qSr, max(qSrNo) as qSrNo, " + _
                        "max(qRmCtg) as qRmCtg, max(qRmSCtg) as qRmSCtg, qRmCd, " + _
                        "sum(OrQty) as OrQty, sum(qRmAWt) as qRmAWt, sum(qRmIWt) as qRmIWt, " + _
                        "sum(qXtraIWt) as qXtraIWt, sum(qRmZWt) as qRmZWt, " + _
                        "sum(qXtraZWt) as qXtraZWt, qRmARt, qRmIRt, qRmZRt, OrSetSCd, " + _
                        "qSetARt, qSetIRt, qSetZRt,/*sum(qSetIVal) as qSetIVal,sum(qSetZVal) as qSetZVal,*/OrMainMet, max(RmPurityZ) as RmPurityZ,Max(OrCustRmCd) as OrCustRmCd " + _
                        "from " + wTmpTbl + _
                        " join RmMst on RmCd= qRmCd and RmPrtKey='" + ctCurrPrtn + "'  " + _
                        " Group By qRmCd, OrSetSCd, qRmARt, qRmIRt, qRmZRt, qSetARt, " + _
                        "qSetIRt, qSetZRt, OrMainMet " + _
                        " having Round(sum(qRmAWt), 3)> 0 or Round(sum(qRmIWt), 3)> 0 or " + _
                        "Round(sum(qRmZWt), 3)> 0 Order By OrMainMet")
  
  wTotIXtraWt = 0: wTotZXtraWt = 0
  wTotAWt = 0: wTotIWt = 0: wTotZWt = 0: wWtFrOrdYN = ""
  
  With GrdIR
    GrdIR.StartCopy

    Do While Not (wFinalResultSet.EOF Or wFinalResultSet.BOF)
      .AddItem
      wRmRow = .Rows - 1
      mf_StopRecalcDuringCopyIR = False
      If UCase(ms_IdTcTyp) = UCase(ctInTcTyp) Then
        wDivFactor = IIF(wFinalResultSet!qRmCtg = "C" Or wFinalResultSet!qRmCtg = "D", 5, 1)
        wTotIXtraWt = wTotIXtraWt + MWLib.Div(wFinalResultSet!qXtraIWt, wDivFactor)
        wTotZXtraWt = wTotZXtraWt + MWLib.Div(wFinalResultSet!qXtraZWt, wDivFactor)
        .Value(wRmRow, "IrSrNo") = wRmRow
        .Value(wRmRow, "IrPrtKey") = ctCurrPrtn         '****** Sachin 3.02 02-01-08
        .Value(wRmRow, "IrRmCd") = wFinalResultSet!qRmCd
        .Value(wRmRow, "IrLotNo") = ""
        .Value(wRmRow, "IrRmSz") = 0
        .Value(wRmRow, "IrRmSz2") = 0
        .Value(wRmRow, "IrRmSz3") = 0
        .Value(wRmRow, "IrStkRt") = 0
        .Value(wRmRow, "IrRmQty") = wFinalResultSet!OrQty
        .Value(wRmRow, "IrRmAWt") = wFinalResultSet!qRmAWt
        
        If wFinalResultSet!OrMainMet = "Y" Then
          wWtFrOrdYN = IIF(Trim(mIWtEqOrd) = "Y", "Y", "N")
          If wWtFrOrdYN = "N" Then wWtFrOrdYN = IIF(InStr(1, wCtgStr, "," + wFinalResultSet!qRmCtg + ",") <> 0, "Y", "N")
          Dim wd_Wt As Double
          If mb_InInvGrsWtEqActual = True Then
            wd_Wt = IIF(wWtFrOrdYN = "N", wFinalResultSet!qRmIWt - (wTotIWt - wTotAWt), wFinalResultSet!qRmIWt - wTotIXtraWt)
          Else
            wd_Wt = wFinalResultSet!qRmIWt
          End If
          .Value(wRmRow, "IrRmIWt") = wd_Wt
          If whZWtEq = "A" Then
            .Value(wRmRow, "IrRmZWt") = wFinalResultSet!qRmZWt - (wTotZWt - wTotAWt)
          ElseIf whZWtEq = "I" Then
            '*** (Bef 3.01.0_1)
            '.Value(wRmRow, "IrRmZWt") = IIF(wWtFrOrdYN = "N", wFinalResultSet!qRmIWt - (wTotIWt - wTotAWt), wFinalResultSet!qRmIWt - wTotIXtraWt)
            '*** (Bef 3.01.0_1)
            .Value(wRmRow, "IrRmZWt") = wd_Wt       '*** (Jen 3.01.0_1)
          End If
        Else
          ' Manoj 2.10.0
          '.Value(wRmRow, "IrRmIWt") = wFinalResultSet!qRmIWt + wFinalResultSet!qXtraIWt
          .Value(wRmRow, "IrRmIWt") = wFinalResultSet!qRmIWt + _
                IIF(UCase(wFinalResultSet!qRmCtg) = "C" Or UCase(wFinalResultSet!qRmCtg) = "D", _
                  Round(wFinalResultSet!qXtraIWt, 3), _
                  wFinalResultSet!qXtraIWt)
                  ' Bef 3.8.0 - Round off Changed to 3 instead of 2
'                  Round(wFinalResultSet!qXtraIWt, 2), _
'                  wFinalResultSet!qXtraIWt)
                  
                 
          If whZWtEq = "A" Then
            '.Value(wRmRow, "IrRmZWt") = wFinalResultSet!qRmzWt + wFinalResultSet!qXtraZWt
            .Value(wRmRow, "IrRmZWt") = wFinalResultSet!qRmZWt + _
                IIF(UCase(wFinalResultSet!qRmCtg) = "C" Or UCase(wFinalResultSet!qRmCtg) = "D", _
                  Round(wFinalResultSet!qXtraZWt, 3), _
                  wFinalResultSet!qXtraZWt)
                  ' Bef 3.8.0 - Round off Changed to 3 instead of 2
                  'Round(wFinalResultSet!qXtraZWt, 2), _
                  'wFinalResultSet!qXtraZWt)
          
          ElseIf whZWtEq = "I" Then
            '.Value(wRmRow, "IrRmZWt") = wFinalResultSet!qRmIWt + wFinalResultSet!qXtraIWt
            .Value(wRmRow, "IrRmZWt") = wFinalResultSet!qRmIWt + _
                IIF(UCase(wFinalResultSet!qRmCtg) = "C" Or UCase(wFinalResultSet!qRmCtg) = "D", _
                  Round(wFinalResultSet!qXtraIWt, 3), _
                  wFinalResultSet!qXtraIWt)
                  
                  
            ' Bef 3.8.0 - Round off Changed to 3 instead of 2
            ' Round(wFinalResultSet!qXtraIWt, 2), _
            ' wFinalResultSet!qXtraIWt)
                  
          End If
          ' Manoj 2.10.0
        End If
        
        wTotAWt = wTotAWt + MWLib.Div(.Value(wRmRow, "IrRmAWt"), wDivFactor)
        wTotIWt = wTotIWt + MWLib.Div(.Value(wRmRow, "IrRmIWt"), wDivFactor)
        wTotZWt = wTotZWt + MWLib.Div(.Value(wRmRow, "IrRmZWt"), wDivFactor)
        
        .Value(wRmRow, "IrRmARt") = wFinalResultSet!qRmARt
        
        ' ***** Manali 3.03 - 05/07/08
        Dim wOmLkUpRmRt As String, wOmLkUpMetLs As String, wOmLkUp As String
        Dim wOmLkUpArr() As String
        
        wOmLkUp = moCn.GetFldVal("Select OmLkUpRmRt+','+OmLkUpMetLs from OrdMst where " + _
                                 "OmCoCd= '" + gs_CoCd + "' and OmTc= '" + mExpOdTc + "' and " + _
                                 "OmYy= '" + mExpOdYy + "'  and OmChr= '" + mExpOdChr + "' and " + _
                                 "OmNo= " + CStr(mExpOdNo) + " and OmPrtKey='" + ctCurrPrtn + "' ")
        If wOmLkUp <> "" Then
          wOmLkUpArr() = Split(wOmLkUp, ",")
          wOmLkUpRmRt = wOmLkUpArr(0)
          wOmLkUpMetLs = wOmLkUpArr(1)
        End If
        ' ***** Manali 3.03 - 05/07/08
        If ms_RtFrInvCust = "Y" Then  '6.1-55
          wOmLkUpRmRt = ""
          wOmLkUpMetLs = ""
        End If
        'Geeta***07/06/05
        '****** Sachin 3.02 31-12-07 - prtkey added
        If wFinalResultSet!OrMainMet = "Y" Then _
             .Value(wRmRow, "IrAlySalRt") = moCn.GetFldVal("Select  OrAlySalRt from OrdRm where OrCoCd='" + gs_CoCd + "' and OrTc= '" + mExpOdTc + "' and " + _
                                                         "OrYy= '" + mExpOdYy + "' and OrChr= '" + mExpOdChr + "' and " + _
                                                         "OrNo= " + CStr(mExpOdNo) + " and OrSr= " + CStr(mExpOdSr) + " and OrPrtKey='" + ctCurrPrtn + "' and  OrMainMet='Y' ")
        'Geeta***07/06/05
        
        If UCase(wFinalResultSet!qRmCtg) = "G" And mIGldRtEqOrd <> "Y" Then
' Geeta 07/06/05 Before 2.11
'          .Value(wRmRow, "IrRmIRt") = GetGldRt(mExpCmCd, ADC("wIdLmgSal"), _
                                      wFinalResultSet!qRmCd, 0, "DM", mDmCd, "S", "C")
                                                
          '*** (Bef 2.13)
          '.Value(wRmRow, "IrRmIRt") = GetGldRt(mExpCmCd, ADC("wIdLmgSal"), _
                                      wFinalResultSet!qRmCd, .Value(wRmRow, "IrAlySalRt"), "DM", mDmCd, "S", "C")
          '*** (Bef 2.13)
          
          '*** (Jen 2.13)
          ' **** Manali 3.03 - 05/07/08 - Om Lk Up fields passed if not blank
          .Value(wRmRow, "IrRmIRt") = GetGldRt(mExpCmCd, ADC("wIdLmgSal"), _
                                      wFinalResultSet!qRmCd, .Value(wRmRow, "IrAlySalRt"), "DM", mDmCd, "S", "C", , ADC("wIdCmCurCd"), _
                                      IIF(wOmLkUpRmRt <> "", wOmLkUpRmRt, ""), IIF(wOmLkUpMetLs <> "", wOmLkUpMetLs, ""), , , .Value(wRmRow, "IRRMSZ"), .Value(wRmRow, "IRRMSZ2"), 0)
          '*** (Jen 2.13)
                                      
          '*** (Bef 2.13) Remove 'Local Inv' Check
          'If mLocalExp = "N" Then .Value(wRmRow, "IrRmZRt") = .Value(wRmRow, "IrRmIRt")
          '*** (Bef 2.13) Remove 'Local Inv' Check
        
          '*** (Jen 2.13) Remove 'Local Inv' Check
          .Value(wRmRow, "IrRmZRt") = .Value(wRmRow, "IrRmIRt")
          '*** (Jen 2.13) Remove 'Local Inv' Check
        ElseIf UCase(wFinalResultSet!qRmCtg) = "P" And mIGldRtEqOrd <> "Y" Then
' Geeta 07/06/05 Before 2.11
'          .Value(wRmRow, "IrRmIRt") = GetGldRt(mExpCmCd, ADC("wIdLmpSal"), _
                                      wFinalResultSet!qRmCd, 0, "DM", mDmCd, "S", "C")
          
          '*** (Bef 2.13)
          '.Value(wRmRow, "IrRmIRt") = GetGldRt(mExpCmCd, ADC("wIdLmpSal"), _
                                      wFinalResultSet!qRmCd, .Value(wRmRow, "IrAlySalRt"), "DM", mDmCd, "S", "C")
          '*** (Bef 2.13)
          
          '*** (Jen 2.13)
          ' **** Manali 3.03 - 05/07/08 - Om Lk Up fields passed if not blank
          .Value(wRmRow, "IrRmIRt") = GetGldRt(mExpCmCd, ADC("wIdLmpSal"), _
                                      wFinalResultSet!qRmCd, .Value(wRmRow, "IrAlySalRt"), "DM", mDmCd, "S", "C", , ADC("wIdCmCurCd"), _
                                      IIF(wOmLkUpRmRt <> "", wOmLkUpRmRt, ""), IIF(wOmLkUpMetLs <> "", wOmLkUpMetLs, ""))
          '*** (Jen 2.13)
          
          '*** (Bef 2.13) Remove 'Local Inv' Check
          'If mLocalExp = "N" Then .Value(wRmRow, "IrRmZRt") = .Value(wRmRow, "IrRmIRt")
          '*** (Bef 2.13) Remove 'Local Inv' Check
          
          '*** (Jen 2.13) Remove 'Local Inv' Check
          .Value(wRmRow, "IrRmZRt") = .Value(wRmRow, "IrRmIRt")
          '*** (Jen 2.13) Remove 'Local Inv' Check

        ' ****** Sachin 2.12 - 06/12/05 - SJM
        ElseIf UCase(wFinalResultSet!qRmCtg) = "S" And mIGldRtEqOrd <> "Y" Then
          If mSlvMod = True Then
            '*** (Bef 2.13)
            '.Value(wRmRow, "IrRmIRt") = GetGldRt(mExpCmCd, ADC("wIdLmsSal"), _
                                      wFinalResultSet!qRmCd, .Value(wRmRow, "IrAlySalRt"), "DM", mDmCd, "S", "C")
            '*** (Bef 2.13)
            
            '*** (Jen 2.13)
            ' **** Manali 3.03 - 05/07/08 - Om Lk Up fields passed if not blank
            .Value(wRmRow, "IrRmIRt") = GetGldRt(mExpCmCd, ADC("wIdLmsSal"), _
                                        wFinalResultSet!qRmCd, .Value(wRmRow, "IrAlySalRt"), "DM", mDmCd, "S", "C", , ADC("wIdCmCurCd"), _
                                        IIF(wOmLkUpRmRt <> "", wOmLkUpRmRt, ""), IIF(wOmLkUpMetLs <> "", wOmLkUpMetLs, ""), , , .Value(wRmRow, "IRRMSZ"), .Value(wRmRow, "IRRMSZ2"), 0)
            '*** (Jen 2.13)
          Else
            .Value(wRmRow, "IrRmIRt") = 0
          End If
          
          '*** (Bef 2.13) Remove 'Local Inv' Check
          'If mLocalExp = "N" Then .Value(wRmRow, "IrRmZRt") = .Value(wRmRow, "IrRmIRt")
          '*** (Bef 2.13) Remove 'Local Inv' Check
          
          '*** (Jen 2.13) Remove 'Local Inv' Check
          .Value(wRmRow, "IrRmZRt") = .Value(wRmRow, "IrRmIRt")
          '*** (Jen 2.13) Remove 'Local Inv' Check
          
        ' ****** Sachin 2.12 - 06/12/05 - SJM
        
        '*** (Jen 2.13)
        ' **** Manali 3.03 - 21/06/08 - Pd Mod
        ElseIf UCase(wFinalResultSet!qRmCtg) = "L" And mIGldRtEqOrd <> "Y" Then
          If mb_PdMod = True Then
            ' **** Manali 3.03 - 05/07/08 - Om Lk Up fields passed if not blank
            .Value(wRmRow, "IrRmIRt") = GetGldRt(mExpCmCd, ADC("wIdLmlSal"), _
                                        wFinalResultSet!qRmCd, .Value(wRmRow, "IrAlySalRt"), "DM", mDmCd, "S", "C", , ADC("wIdCmCurCd"), _
                                        IIF(wOmLkUpRmRt <> "", wOmLkUpRmRt, ""), IIF(wOmLkUpMetLs <> "", wOmLkUpMetLs, ""), , , .Value(wRmRow, "IRRMSZ"), .Value(wRmRow, "IRRMSZ2"), 0)
          Else
            .Value(wRmRow, "IrRmIRt") = 0
          End If
          .Value(wRmRow, "IrRmZRt") = .Value(wRmRow, "IrRmIRt")
        ' **** Manali 3.03 - 21/06/08 - Pd Mod
        ' **** Manali 3.6.0 - 27/10/09 - Cost Markup And Limit In Invoicing Checked For Accessories
        ElseIf UCase(wFinalResultSet!qRmCtg) = "X" And mIGldRtEqOrd <> "Y" And _
               moCn.RecSeek("select 'x' from Param where PTyp= 'RMSCTG' and PMCd= '" + wFinalResultSet!qRmCtg + "' " + _
                            "and PSCd= '" + wFinalResultSet!qRmSCtg + "' and PValue= 'S' ") And _
               moCn.RecSeek("Select 'x' From InvHd Where InCoCd='" + gs_CoCd + "' And InTc='" + ADC("wIdTc") + "' And InYy='" + ADC("wIdYy") + "' " + _
                            " And InChr='" + ADC("wIdChr") + "' And InNo=" + CStr(ADC("wIdNo")) + " And InPrtKey='" + ADC("wIdPrtKey") + "' And InXIRtWrtStkYN='N' ") Then
          
          '****** Sachin 3.02 31-12-07 - prtkey added
          Dim ws_BaseRmCtg As String, wf_BaseLmeOrd As Single, wf_baselmeinv As Single
          ws_BaseRmCtg = moCn.GetFldVal("Select B.RmCtg from RmMst B where " + _
                         "RmCd= (Select Rm.RmBaseCd from RmMst Rm where Rm.RmCd= '" + wFinalResultSet!qRmCd + "' and Rm.RmPrtKey='" + ctCurrPrtn + "') and RmPrtKey='" + ctCurrPrtn + "' ")
          
          '****** Sachin 3.02 31-12-07 - prtkey added
          ' **** Manali 3.03 - 21/06/08 - Pd Mod - 'L' included
          wf_BaseLmeOrd = moCn.GetFldVal("Select (case '" + ws_BaseRmCtg + "' when 'G' then OmLmgSal " + _
                          "when 'P' then OmLmpSal when 'S' then OmLmsSal when 'L' then OmLmlSal else 0 end) from OrdMst where " + _
                          "OmCoCd= '" + gs_CoCd + "' and OmTc= '" + mExpOdTc + "' and OmYy= '" + mExpOdYy + "' " + _
                          "and OmChr= '" + mExpOdChr + "' and OmNo= " + CStr(mExpOdNo) + " and OmPrtKey='" + ctCurrPrtn + "' ")
                                 
          ' **** Manali 3.03 - 21/06/08 - Pd Mod - 'L' included
          wf_baselmeinv = moCn.GetFldVal("Select (case '" + ws_BaseRmCtg + "' when 'G' then " + CStr(ADC("wIdLmgSal")) + " " + _
                          "when 'P' then " + CStr(ADC("wIdLmpSal")) + " when 'S' then " + CStr(ADC("wIdLmsSal")) + " " + _
                          "when 'L' then " + CStr(ADC("wIdLmlSal")) + " else 0 end) ")

          '*** (Bef 2.13)
          'If GetRmRt(mExpCmCd, wFinalResultSet!qRmCd, 0, 0, 0, "S", wf_BaseLmeOrd) = _
             GetRmRt(mExpCmCd, wFinalResultSet!qRmCd, 0, 0, 0, "S", wf_BaseLmeInv) Then
          '*** (Bef 2.13)

          '*** (Jen 2.13)
          ' **** Manali 3.03 - 05/07/08 - Om Lk Up fields passed if not blank
          If GetRmRt(mExpCmCd, wFinalResultSet!qRmCd, 0, 0, 0, "S", wf_BaseLmeOrd, , , , , ADC("wIdCmCurCd"), IIF(wOmLkUpRmRt <> "", wOmLkUpRmRt, "")) = _
             GetRmRt(mExpCmCd, wFinalResultSet!qRmCd, 0, 0, 0, "S", wf_baselmeinv, , , , , ADC("wIdCmCurCd"), IIF(wOmLkUpRmRt <> "", wOmLkUpRmRt, "")) Then
          '*** (Jen 2.13)
            .Value(wRmRow, "IrRmIRt") = wFinalResultSet!qRmIRt
          Else
            '*** (Bef 2.13)
            '.Value(wRmRow, "IrRmIRt") = GetRmRt(mExpCmCd, wFinalResultSet!qRmCd, 0, 0, 0, "S", wf_BaseLmeInv)
            '*** (Bef 2.13)
            
            '*** (Jen 2.13)
            ' **** Manali 3.03 - 05/07/08 - Om Lk Up fields passed if not blank
            .Value(wRmRow, "IrRmIRt") = GetRmRt(mExpCmCd, wFinalResultSet!qRmCd, 0, 0, 0, "S", wf_baselmeinv, , , , , ADC("wIdCmCurCd"), IIF(wOmLkUpRmRt <> "", wOmLkUpRmRt, ""))
            '*** (Jen 2.13)
          End If
          '.Value(wRmRow, "IrRmZRt") = wFinalResultSet!qRmZRt
          .Value(wRmRow, "IrRmZRt") = .Value(wRmRow, "IrRmIRt")
        '*** (Jen 2.13)
        
        Else
        
          '****** Sachin 3.3.0 (25-06-08) - [14. Cost Markup And Limit In Invoicing] ******
          ' ***** Manali 3.6.0 - Cost Markup And Limit In Invoicing For Accessories
          If ((wFinalResultSet!qRmCtg = "D" Or wFinalResultSet!qRmCtg = "C") And _
              moCn.RecSeek("Select 'x' From InvHd Where InCoCd='" + gs_CoCd + "' And InTc='" + ADC("wIdTc") + "' And InYy='" + ADC("wIdYy") + "' And InChr='" + ADC("wIdChr") + "' And InNo=" + CStr(ADC("wIdNo")) + " And InPrtKey='" + ADC("wIdPrtKey") + "' And InIRtWrtStkYN='Y' ")) Or _
              (wFinalResultSet!qRmCtg = "X" And _
              moCn.RecSeek("Select 'x' From InvHd Where InCoCd='" + gs_CoCd + "' And InTc='" + ADC("wIdTc") + "' And InYy='" + ADC("wIdYy") + "' And InChr='" + ADC("wIdChr") + "' And InNo=" + CStr(ADC("wIdNo")) + " And InPrtKey='" + ADC("wIdPrtKey") + "' And InXIRtWrtStkYN='Y' ")) Then
              '****** Sachin 3.3.0 - Since Stock Rate is Straightaway set to Zero IRt will be set to 0
              .Value(wRmRow, "IrRmIRt") = 0
              .Value(wRmRow, "IrRmZRt") = 0
            
          '****** Sachin 3.3.0 (25-06-08) - [14. Cost Markup And Limit In Invoicing] ******
          Else

            .Value(wRmRow, "IrRmIRt") = wFinalResultSet!qRmIRt
            '*** (Bef 2.13) Remove 'Local Inv' Check
            'If mLocalExp = "N" Or (mLocalExp = "Y" And UCase(wFinalResultSet!qRmCtg) <> "G" And UCase(wFinalResultSet!qRmCtg) <> "P" And UCase(wFinalResultSet!qRmCtg) <> "S") Then
            '  .Value(wRmRow, "IrRmZRt") = wFinalResultSet!qRmZRt
            'End If
            '*** (Bef 2.13) Remove 'Local Inv' Check
            
            '*** (Jen 2.13) Remove 'Local Inv' Check
            .Value(wRmRow, "IrRmZRt") = wFinalResultSet!qRmZRt
            '*** (Jen 2.13) Remove 'Local Inv' Check
          End If
        End If
        
        .Value(wRmRow, "IrCustRmCd") = wFinalResultSet!OrCustRmCd
        .Value(wRmRow, "IrSetSCd") = wFinalResultSet!OrSetSCd
        .Value(wRmRow, "IrSetARt") = wFinalResultSet!qSetARt
        .Value(wRmRow, "IrSetIRt") = wFinalResultSet!qSetIRt
        .Value(wRmRow, "IrSetZRt") = wFinalResultSet!qSetZRt
        
        .Value(wRmRow, "IrMainMet") = wFinalResultSet!OrMainMet
        
        '*** (Jen 2.13)
        'wf_CustmCdPurityZ = moCn.GetFldVal("Select R.RmPurityZ from RmMst R where R.RmCd= (Select R1.RmZCd from RmMst R1 where R1.RmCd= '" + wFinalResultSet!qRmCd + "') ")
        '****** Sachin 3.02 31-12-07 - prtkey added
        ' **** Manali 3.03 - 21/06/08 - Pd Mod - 'L' included
        wf_CustmCdPurityZ = moCn.GetFldVal("Select R.RmPurityZ* " + _
                            "(case when (R.RmCtg= 'G' and " + CStr(ADC("wIdGldFiness")) + "> 0) then " + CStr(ADC("wIdGldFiness")) + " " + _
                            "      when (R.RmCtg= 'P' and " + CStr(ADC("wIdPlFiness")) + "> 0) then " + CStr(ADC("wIdPlFiness")) + " " + _
                            "      when (R.RmCtg= 'S' and " + CStr(ADC("wIdSlFiness")) + "> 0) then " + CStr(ADC("wIdSlFiness")) + " " + _
                            "      when (R.RmCtg= 'L' and " + CStr(ADC("wIdPdFiness")) + "> 0) then " + CStr(ADC("wIdPdFiness")) + " " + _
                            "else 1 end) from RmMst R where R.RmCd= (Select R1.RmZCd from RmMst R1 where R1.RmCd= '" + wFinalResultSet!qRmCd + "' and R1.RmPrtKey='" + ctCurrPrtn + "') and R.RmPrtKey='" + ctCurrPrtn + "' ")
        '.Value(wRmRow, "IrRmZPureRt") = IIF(wFinalResultSet!RmPurityZ > 0, MWLib.Div(.Value(wRmRow, "IrRmZRt"), wFinalResultSet!RmPurityZ) * wf_CustmCdPurityZ, .Value(wRmRow, "IrRmZRt"))
        '.Value(wRmRow, "IrRmZPureRt") = IIF(wFinalResultSet!RmPurityZ > 0, MWLib.Div(IIF(moCn.GetFldVal("Select RmQw from RmMst where " + _
                                         "RmCd= '" + .Value(wRmRow, "IrRmCd") + "' ") = "Q", _
                                         MWLib.Div(.Value(wRmRow, "IrRmZVal"), .Value(wRmRow, "IrRmZWt")), .Value(wRmRow, "IrRmZRt")), wFinalResultSet!RmPurityZ) * wf_CustmCdPurityZ, .Value(wRmRow, "IrRmZRt"))
          
        '****** Sachin 3.02 31-12-07 - prtkey added
        .Value(wRmRow, "IrRmZPureRt") = IIF(moCn.GetFldVal("Select RmQw from RmMst where " + _
                                         "RmCd= '" + .Value(wRmRow, "IrRmCd") + "' and RmPrtKey='" + ctCurrPrtn + "' ") = "Q", _
                                         MWLib.Div(.Value(wRmRow, "IrRmZVal"), .Value(wRmRow, "IrRmZWt")), _
                                         .Value(wRmRow, "IrRmZRt")) / IIF(wFinalResultSet!RmPurityZ > 0, wFinalResultSet!RmPurityZ, 1) * _
                                         IIF(wf_CustmCdPurityZ > 0, wf_CustmCdPurityZ, 1)
        '*** (Jen 2.13)
        
        '*** (Bef 2.13)
        '.Value(wRmRow, "IrRmZPureRt") = IIF(wFinalResultSet!RmPurityZ > 0, MWLib.Div(.Value(wRmRow, "IrRmZRt"), wFinalResultSet!RmPurityZ), .Value(wRmRow, "IrRmZRt"))
        '*** (Bef 2.13)
        
      ElseIf UCase(ms_IdTcTyp) = UCase(ctCTBTcTyp) Then
        .Value(wRmRow, "IrSrNo") = wRmRow
        .Value(wRmRow, "IrPrtKey") = ctCurrPrtn     '****** Sachin 3.02 02-01-08
        .Value(wRmRow, "IrRmCd") = wFinalResultSet!qRmCd
        ' Zubin 212
        .Value(wRmRow, "IrLotNo") = ""
        ' Zubin 212
        .Value(wRmRow, "IrRmSz") = 0
        .Value(wRmRow, "IrRmSz2") = 0
        .Value(wRmRow, "IrRmSz3") = 0
        .Value(wRmRow, "IrStkRt") = 0
        .Value(wRmRow, "IrRmQty") = wFinalResultSet!OrQty
        .Value(wRmRow, "IrRmAWt") = wFinalResultSet!qRmAWt
        .Value(wRmRow, "IrRmIWt") = wFinalResultSet!qRmIWt
        .Value(wRmRow, "IrRmZWt") = wFinalResultSet!qRmZWt
        .Value(wRmRow, "IrCustRmCd") = wFinalResultSet!OrCustRmCd
        .Value(wRmRow, "IrSetSCd") = wFinalResultSet!OrSetSCd
        .Value(wRmRow, "IrSetARt") = wFinalResultSet!qSetARt
        .Value(wRmRow, "IrMainMet") = wFinalResultSet!OrMainMet
        .Value(wRmRow, "IrRmZPureRt") = 0
        .Value(wRmRow, "IrPrtKey") = ctCurrPrtn     '****** Sachin 3.02 31-12-07
      End If
      
 'sri      mf_StopRecalcDuringCopyIR = True
      
      mf_StopRecalcDuringCopyIR = False
      
      'Call GrdIR_ReCalcRoutine(wRmRow, UCase("IrRmAVal"), 0, "", "")
 '
     ' Call GrdIR_ReCalcRoutine(wRmRow, UCase("IrRmIRt"), 0, "", "")
      
     ' Call GrdIR_ReCalcRoutine(wRmRow, UCase("IrRmIVal"), 0, "", "")
      
     ' Call GrdIR_ReCalcRoutine(wRmRow, UCase("IrRmZRt"), 0, "", "")
      
      'Call GrdIR_ReCalcRoutine(wRmRow, UCase("IrRmZVal"), 0, "", "")
      
'      Call GrdIR_ReCalcRoutine(wRmRow, UCase("IrSetAVal"), 0, "", "")
'
 '     Call GrdIR_ReCalcRoutine(wRmRow, UCase("IrSetIRt"), 0, "", "")
      
  '    Call GrdIR_ReCalcRoutine(wRmRow, UCase("IrSetIVal"), 0, "", "")
      
   '   Call GrdIR_ReCalcRoutine(wRmRow, UCase("IrSetZRt"), 0, "", "")
   
    '  Call GrdIR_ReCalcRoutine(wRmRow, UCase("IrSetZVal"), 0, "", "")
      
   '  Call GrdIR_ReCalcRoutine(wRmRow, UCase("IrRmZPureRt"), 0, "", "")
      
      'Call GrdIR_ReCalcRoutine(wRmRow, UCase("IrRmZCd"), 0, "", "")
      
      'Call GrdIR_ReCalcRoutine(wRmRow, UCase("wQwSal"), 0, "", "")

      
'      TrackStatus ("Update of IrRm record Started ") + CStr(wRmRow)
      .SaveRec (wRmRow)
 '     TrackStatus ("Update of IrRm record Completed ") + CStr(wRmRow)
      wFinalResultSet.MoveNext
    Loop
    GrdIR.EndCopy
    
    
    Set wRsUniqRm = Nothing
    Set wRsTotActWtQty = Nothing
    Set wRsRmDet = Nothing
    Set wFinalResultSet = Nothing
  End With
End Sub

Public Sub MakeRmFromFg1(ByVal mIWtEqOrd As String, ByVal mIWtFrOrd As String, _
                        ByVal mIGldRtEqOrd As String, ByVal mIdQty As Single)
'  Dim whZWtEq As String, wCtgStr As String, wDmCtg As String, wTmpTbl As String
'  Dim wOrStr As String, wStr As String, wMainMet As Boolean
'  Dim wRsRmDet As MDORowSet, wOrRecSet As MDORowSet
'  Dim wDivFactor As Single, wRmRow As Single, wLabRow As Single, wf_AddLab As Single
'  Dim wTotIXtraWt As Double, wTotZXtraWt As Double, wTotAWt As Double
'  Dim wTotIWt As Double, wTotZWt As Double
'  Dim wCmDiaRtFA As String, wCmCsRtFA As String, wf_CnvRt As Single
'  Dim wDRrTcTyp As String, wCRrTcTyp As String, wDFrOrdYN As String
'  Dim wCFrOrdYN As String
'  Dim wSetRt As Boolean
'  Dim wTotWtDiff As Double, wWtDiff As Double, wTotIrRmWt As Double, wAddWt As Double
'  Dim wIrRmWt As MDORowSet, wRmCtg As String, wFrOrdYN As String, wMainMetWt As Double
'  Dim wFrRmWt As Double, wFrRmQty As Double, wOrTotWt As Double, wOrTotQty As Double
'  Dim wFRmDet As MDORowSet
'  whZWtEq = moCn.GetFldVal("Select hZWtEq from Head where HCoCd= '" + gs_CoCd + "' and HCd='" + ctSelfCmCd + "'")
'  Dim wSetTyp As Boolean
'  'mIWtFrOrd = "DC"
'  'wCtgStr= ",D,C"
'  wCtgStr = "," + moCn.GetFldVal("Select PDesc225 from Param where " + _
'              "PTyp='IWTFRORD' and PMCd = '" + mIWtFrOrd + "'") + ","
'
'  wCmDiaRtFA = moCn.GetFldVal("Select CmDiaRtFA from CustMst where CmCtg= 'C' and " + _
'               "CmCd= '" + GrdID.Value(GrdID.Row, "IdExpCmCd") + "' ")
'
'  wCmCsRtFA = moCn.GetFldVal("Select CmCSRtFA from CustMst where CmCtg= 'C' and " + _
'               "CmCd= '" + GrdID.Value(GrdID.Row, "IdExpCmCd") + "' ")
'
'
'  wDmCtg = moCn.GetFldVal("Select DmCtg from DsgMst where DmTcTyp= 'DM' and DmSz= '' and " + _
'               "DmCd= '" + GrdID.Value(GrdID.Row, "IdDmCd") + "' ")
'  wTmpTbl = "#TmpInvRm"
'
'  '*******Geeta*******Emr207
'
'  If UCase(ms_IdTcTyp) = UCase(ctInTcTyp) Then
'
'     wStr = "Select FrRmCd as qIrRmCd, FrRmSz as qIrRmSz,sum(FrRmStkRt*FrRmWt)/(case when sum(FrRmWt)> 0  " + _
'           "then sum(FrRmWt) else 1 end) as qIrStkRt, " + _
'           "sum(FrRmQty) as qIrRmQty, sum(FrRmWt) As qIrRmAWt, sum(FrRmWt) as qIrRmIWt, " + _
'           "sum(FrRmWt) as qIrRmZWt,max(Isnull(qOrWt,0)) as qOrWt,max(IsNull(qOrQty,0)) as qOrQty, 9999.999 as qXtraIWt, 9999.999 as qXtraZWt, " + _
'           "999999.999 as qIrRmARt, 999999.999 as qIrRmIRt, 999999.999 as qIrRmZRt, " + _
'           "space(8) as qIrSetSCd,isnull(qOrSetSalRt,0.00) as qIrSetARt,isnull(qOrSetSalRt,0.00) as qIrSetIRt,isnull(qOrSetSalRt,0.00) as qIrSetZRt, " + _
'           "Space(1) as qIrMainMet, max(RmPurityZ) as qRmPurityZ, max(RmCtg) as qRmCtg, max(RmSCtg) as qRmSCtg " + _
'           "into " + wTmpTbl + " from InvFgd join FgRm on FrCoCd= IfCoCd and FrTc= IfFdTc and FrYy= IfFdyy and " + _
'           "FrChr= IfFdChr and FrNo = IfFdNo And FrSr = IfFdSr " + _
'           "Left Outer Join (select OrRmcd as qOrRmcd ,Orln1 as qOrLn1,OrSetSalRt as qOrSetSalRt," + _
'           "sum(OrWt) as qOrWt,sum(OrQty) as qOrQty  from OrdRm  " + _
'           "where OrCoCd= '" + gs_CoCd + "'  and OrTc = '" + GrdID.Value(GrdID.Row, "IdExpOdTc") + "' " + _
'           "and OrYy= '" + GrdID.Value(GrdID.Row, "IdExpOdYy") + "' " + _
'           " and OrChr= '" + GrdID.Value(GrdID.Row, "IdExpOdChr") + "' and OrNo = " + CStr(GrdID.Value(GrdID.Row, "IdExpOdNo")) + _
'           " and OrSr= " + CStr(GrdID.Value(GrdID.Row, "IdExpOdSr")) + " group by OrCocd ,OrTc,OrYy,OrChr,OrNo,OrSr, " + _
'           "OrRmcd,Orln1,OrSetSalRt) Or2  On qOrRmcd=FrRmcd and qOrln1=FrRmSz " + _
'           "Join RmMst on FrRmCd = RmCd " + _
'           "where IFcocd = '" + gs_CoCd + "' and IFTc= '" + ADC("wIdTc") + "' and IFYy= '" + _
'           ADC("wIdYy") + "' and IFChr= '" + ADC("wIdChr") + "' and " + _
'           "IFNo = " + CStr(ADC("wIdNo")) + " and IfSr = " + CStr(GrdID.Value(GrdID.Row, "IdSr")) + " " + _
'           "Group by FrRmCd, FrRmSz,qOrSetSalRt "
'
'
'        '**********Prev***********
''    wStr = "Select FrRmCd as qIrRmCd, FrRmSz as qIrRmSz, sum(FrRmStkRt*FrRmWt)/(case when sum(FrRmWt)> 0  " + _
''           "then sum(FrRmWt) else 1 end) as qIrStkRt, " + _
''           "sum(FrRmQty) as qIrRmQty, sum(FrRmWt) As qIrRmAWt, sum(FrRmWt) as qIrRmIWt, " + _
''           "sum(FrRmWt) as qIrRmZWt, 9999.999 as qXtraIWt, 9999.999 as qXtraZWt, " + _
''           "999999.999 as qIrRmARt, 999999.999 as qIrRmIRt, 999999.999 as qIrRmZRt, " + _
''           "space(8) as qIrSetSCd, 9999.99 as qIrSetARt, 9999.99 as qIrSetIRt, 9999.99 as qIrSetZRt, " + _
''           "Space(1) as qIrMainMet, max(RmPurityZ) as qRmPurityZ, max(RmCtg) as qRmCtg, max(RmSCtg) as qRmSCtg " + _
''           "into " + wTmpTbl + " " + _
''           "from InvFgd join FgRm on FrCoCd= IfCoCd and FrTc= IfFdTc and FrYy= IfFdyy and " + _
''           "FrChr= IfFdChr and FrNo = IfFdNo And FrSr = IfFdSr " + _
''           "Join RmMst on FrRmCd = RmCd " + _
''           "where IFcocd = '" + gs_CoCd + "' and IFTc= '" + adc("wIdTc") + "' and IFYy= '" + _
''           adc("wIdYy") + "' and IFChr= '" + adc("wIdChr") + "' and " + _
''           "IFNo = " + CStr(adc("wIdNo")) + " and IfSr = " + CStr(GrdID.Value(GrdID.Row, "IdSr")) + " " + _
''           "Group by FrRmCd, FrRmSz "
'
''   End If
'    '**********Geeta******emr207
'  ElseIf UCase(ms_IdTcTyp) = UCase(ctCTBTcTyp) Then
'   '**********Prev***********
'  '  wStr = "Select FrRmCd as qIrRmCd, FrRmSz as qIrRmSz, sum(FrRmStkRt*FrRmWt)/(case when sum(FrRmWt)> 0  " + _
'           "then sum(FrRmWt) else 1 end) as qIrStkRt, " + _
'           "sum(FrRmQty) as qIrRmQty, sum(FrRmWt) As qIrRmAWt, sum(FrRmWt) as qIrRmIWt, " + _
'           "sum(FrRmWt) as qIrRmZWt, 0 as qXtraIWt, 0 as qXtraZWt, " + _
'           "0 as qIrRmARt, 0 as qIrRmIRt, 0 as qIrRmZRt, " + _
'           "space(8) as qIrSetSCd, 9999.99 as qIrSetARt, 0 as qIrSetIRt, 0 as qIrSetZRt, " + _
'           "Space(1) as qIrMainMet, max(RmPurityZ) as qRmPurityZ, max(RmCtg) as qRmCtg, max(RmSCtg) as qRmSCtg " + _
'           "into " + wTmpTbl + " " + _
'           "from InvFgd join FgRm on FrCoCd= IfCoCd and FrTc= IfFdTc and FrYy= IfFdyy and " + _
'           "FrChr= IfFdChr and FrNo = IfFdNo And FrSr = IfFdSr " + _
'           "Join RmMst on FrRmCd = RmCd " + _
'           "where IfCoCd = '" + gs_CoCd + "' and IfTc= '" + ADC("wIdTc") + "' " + _
'           "and IfYy= '" + ADC("wIdYy") + "' and IfChr= '" + ADC("wIdChr") + "' and " + _
'           "IfNo = " + CStr(ADC("wIdNo")) + " and IfSr = " + CStr(GrdID.Value(GrdID.Row, "IdSr")) + " " + _
'           "Group by FrRmCd, FrRmSz "
'
'    wStr = "Select FrRmCd as qIrRmCd, FrRmSz as qIrRmSz,sum(FrRmStkRt*FrRmWt)/(case when sum(FrRmWt)> 0  " + _
'           "then sum(FrRmWt) else 1 end) as qIrStkRt, " + _
'           "sum(FrRmQty) as qIrRmQty, sum(FrRmWt) As qIrRmAWt, sum(FrRmWt) as qIrRmIWt, " + _
'           "sum(FrRmWt) as qIrRmZWt,max(Isnull(qOrWt,0)) as qOrWt,max(IsNull(qOrQty,0)) as qOrQty," + _
'           "0 as qXtraIWt, 0 as qXtraZWt, " + _
'           "0 as qIrRmARt, 0 as qIrRmIRt, 0 as qIrRmZRt, " + _
'           "space(8) as qIrSetSCd, isnull(qOrSetSalRt,0.00) as qIrSetARt, isnull(qOrSetSalRt,0.00) as qIrSetIRt, 0 as qIrSetZRt, " + _
'           "Space(1) as qIrMainMet, max(RmPurityZ) as qRmPurityZ, max(RmCtg) as qRmCtg, max(RmSCtg) as qRmSCtg " + _
'           "into " + wTmpTbl + " " + _
'           "from InvFgd join FgRm on FrCoCd= IfCoCd and FrTc= IfFdTc and FrYy= IfFdyy and " + _
'           "FrChr= IfFdChr and FrNo = IfFdNo And FrSr = IfFdSr " + _
'           "Left Outer Join (select OrRmcd as qOrRmcd ,Orln1 as qOrLn1,OrSetSalRt as qOrSetSalRt," + _
'           "sum(OrWt) as qOrWt,sum(OrQty) as qOrQty  from OrdRm  " + _
'           "where OrCoCd= '" + gs_CoCd + "'  and OrTc = '" + GrdID.Value(GrdID.Row, "IdExpOdTc") + "' " + _
'           "and OrYy= '" + GrdID.Value(GrdID.Row, "IdExpOdYy") + "' " + _
'           " and OrChr= '" + GrdID.Value(GrdID.Row, "IdExpOdChr") + "' and OrNo = " + CStr(GrdID.Value(GrdID.Row, "IdExpOdNo")) + _
'           " and OrSr= " + CStr(GrdID.Value(GrdID.Row, "IdExpOdSr")) + " group by OrCocd ,OrTc,OrYy,OrChr,OrNo,OrSr, " + _
'           "OrRmcd,Orln1,OrSetSalRt) Or2  On qOrRmcd=FrRmcd and qOrln1=FrRmSz " + _
'           "Join RmMst on FrRmCd = RmCd " + _
'           "where IfCoCd = '" + gs_CoCd + "' and IfTc= '" + ADC("wIdTc") + "' " + _
'           "and IfYy= '" + ADC("wIdYy") + "' and IfChr= '" + ADC("wIdChr") + "' and " + _
'           "IfNo = " + CStr(ADC("wIdNo")) + " and IfSr = " + CStr(GrdID.Value(GrdID.Row, "IdSr")) + " " + _
'           "Group by FrRmCd, FrRmSz,qOrSetSalRt "
'     End If
'  '**********Geeta******Emr207
'
'  '*** Creating Temporary Table #TmpInvRm ***
'  moCn.CreateTmpTable wTmpTbl, wStr
'  'moCn.Execute (wStr)
'
'  Set wRsRmDet = moCn.OpenResultset("Select * from " + wTmpTbl)
'
'  '**********Geeta******Emr207
'    If moCn.RecSeek("select count(distinct qIrSetIRt) From " + wTmpTbl + " " + _
'                      "group by qIrRmCd,qIrRmSz Having count(distinct qIrSetIRt)>1") Then
'    '  To get Proportionate weights
'     Set wFRmDet = moCn.OpenRes("Select FrRmCd as qFrRmCd,FrRmsz as qFrRmsz, " + _
'                            " Sum(FrRmWt) as qFrRmWt,Sum(FrRmQty) as qFrRmQty from   FgRm  " + _
'                            "join InvFgd on FrCoCd= IfCoCd and FrTc= IfFdTc and FrYy= IfFdyy and  " + _
'                            " FrChr= IfFdChr and FrNo = IfFdNo And FrSr = IfFdSr " + _
'                            "where IFcocd = '" + gs_CoCd + "' and IFTc= '" + ADC("wIdTc") + "' and IFYy= '" + _
'                            ADC("wIdYy") + "' and IFChr= '" + ADC("wIdChr") + "' and " + _
'                            "IFNo = " + CStr(ADC("wIdNo")) + " and IfSr = " + CStr(GrdID.Value(GrdID.Row, "IdSr")) + " Group By FrRmCd,FrRmsz Order By FrRmCd,FrRmsz")
'
'
'   If (wFRmDet.RecCount > 0) Then wFRmDet.MoveFirst
'     Do While Not (wFRmDet.EOF Or wFRmDet.BOF)
'        wFrRmWt = wFRmDet!qFrRmWt
'        wFrRmQty = wFRmDet!qFrRmQty
'        wOrTotWt = moCn.GetFldVal("Select Sum(qOrWt) From " + wTmpTbl + " Where qIrRmCd='" + wFRmDet!qFrRmCd + "' and qIrRmSz='" + CStr(wFRmDet!qFrRmSz) + "' Group By qIrRmcd,qIrRmsz")
'        wOrTotQty = moCn.GetFldVal("Select Sum(qOrQty) From " + wTmpTbl + "  Where qIrRmCd='" + wFRmDet!qFrRmCd + "' and qIrRmSz='" + CStr(wFRmDet!qFrRmSz) + "' Group By qIrRmCd,qIrRmsz")
'
'       Do While Not (wRsRmDet.EOF Or wRsRmDet.BOF)
'           wRsRmDet.edit
'           If wFRmDet!qFrRmCd = wRsRmDet!qIrRmCd Then
'               If moCn.RecSeek("select count(distinct qIrSetIRt) From " + wTmpTbl + " " + _
'                           "Where qIrRmCd= '" + wRsRmDet!qIrRmCd + "' " + _
'                           " and qIrRmSz = " + CStr(wRsRmDet!qIrRmSz) + _
'                           "group by qIrRmCd,qIrRmSz Having count(distinct qIrSetIRt)>1") Then
'
'                   wRsRmDet!qIrRmIWt = Round(MWLib.Div(wRsRmDet!qOrWt, wOrTotWt) * wFrRmWt, 3)
'                   wRsRmDet!qIrRmAWt = wRsRmDet!qIrRmIWt
'                   wRsRmDet!qIrRmZWt = wRsRmDet!qIrRmIWt
'                   wOrTotWt = wOrTotWt - wRsRmDet!qOrWt
'                   wFrRmWt = wFrRmWt - wRsRmDet!qIrRmIWt
'
'                   wRsRmDet!qIrRmQty = Round(MWLib.Div(wRsRmDet!qOrQty, wOrTotQty) * wFrRmQty)
'                   wOrTotQty = wOrTotQty - wRsRmDet!qOrQty
'                   wFrRmQty = wFrRmQty - wRsRmDet!qIrRmQty
'
'           End If
'         End If
'       wRsRmDet.Update
'       wRsRmDet.MoveNext
'     Loop
'      wFRmDet.MoveNext
'     Loop
'  Set wFRmDet = Nothing
'  End If
'  '**********Geeta******Emr207
'  If (wRsRmDet.RecCount > 0) Then wRsRmDet.MoveFirst
'  Do While Not (wRsRmDet.EOF Or wRsRmDet.BOF)
'    wRsRmDet.edit
'
' ' If moCn.GetFldVal("select count(distinct OrSetSalRt) " + _
'             "From OrdRm where OrCoCd= '" + gs_CoCd + "' " + _
'             " and OrTc = '" + GrdID.Value(GrdID.Row, "IdExpOdTc") + "' " + _
'             " and OrYy= '" + GrdID.Value(GrdID.Row, "IdExpOdYy") + "' " + _
'             " and OrChr= '" + GrdID.Value(GrdID.Row, "IdExpOdChr") + "' " + _
'             " and OrNo = " + CStr(GrdID.Value(GrdID.Row, "IdExpOdNo")) + _
'             " and OrSr= " + CStr(GrdID.Value(GrdID.Row, "IdExpOdSr")) + _
'             " and OrRmCd= '" + wRsRmDet!qIrRmCd + "' " + _
'             " and OrLn1 = " + CStr(wRsRmDet!qIrRmSz) + _
'             " group by OrRmCd,OrLn1") > 1 Then
'
'    If moCn.RecSeek("select count(distinct qIrSetIRt) From " + wTmpTbl + " " + _
'                    "Where qIrRmCd= '" + wRsRmDet!qIrRmCd + "' " + _
'                    " and qIrRmSz = " + CStr(wRsRmDet!qIrRmSz) + _
'                    "group by qIrRmCd,qIrRmSz Having count(distinct qIrSetIRt)>1") Then
'
'    '*******Geeta*******Emr207
'
'     wOrStr = "Select OrCstRt, OrSalRt, OrSetSCd, OrSetCstRt,OrSetSalRt, OrMainMet, OrRmCtg " + _
'             "From OrdRm where OrCoCd= '" + gs_CoCd + "' " + _
'             " and OrTc = '" + GrdID.Value(GrdID.Row, "IdExpOdTc") + "' " + _
'             " and OrYy= '" + GrdID.Value(GrdID.Row, "IdExpOdYy") + "' " + _
'             " and OrChr= '" + GrdID.Value(GrdID.Row, "IdExpOdChr") + "' " + _
'             " and OrNo = " + CStr(GrdID.Value(GrdID.Row, "IdExpOdNo")) + _
'             " and OrSr= " + CStr(GrdID.Value(GrdID.Row, "IdExpOdSr")) + _
'             " and OrRmCd= '" + wRsRmDet!qIrRmCd + "' " + _
'             " and OrLn1 = " + CStr(wRsRmDet!qIrRmSz) + _
'             " and OrSetSalRt=" + CStr(wRsRmDet!qIrSetIRt) + _
'             " Order by OrMainMet Desc, OrSalRt Desc"
'   ' Set wOrRecSet = moCn.OpenRes(wOrStr)
'
' Else
'   wOrStr = "Select OrCstRt, OrSalRt, OrSetSCd, OrSetCstRt, OrSetSalRt, OrMainMet, OrRmCtg " + _
'             "From OrdRm where OrCoCd= '" + gs_CoCd + "' " + _
'             " and OrTc = '" + GrdID.Value(GrdID.Row, "IdExpOdTc") + "' " + _
'             " and OrYy= '" + GrdID.Value(GrdID.Row, "IdExpOdYy") + "' " + _
'             " and OrChr= '" + GrdID.Value(GrdID.Row, "IdExpOdChr") + "' " + _
'             " and OrNo = " + CStr(GrdID.Value(GrdID.Row, "IdExpOdNo")) + _
'             " and OrSr= " + CStr(GrdID.Value(GrdID.Row, "IdExpOdSr")) + _
'             " and OrRmCd= '" + wRsRmDet!qIrRmCd + "' " + _
'             " and OrLn1 = " + CStr(wRsRmDet!qIrRmSz) + _
'             "Order by OrMainMet Desc, OrSalRt Desc"
'End If
'    Set wOrRecSet = moCn.OpenRes(wOrStr)
'
'
''    Query to get weighted Average rate
''    wOrStr = "Select (Case when sum(OrWt)<>0 then sum(OrCstRt*OrWt)/sum(OrWt) else 0 end) as OrCstRt, " + _
'              " (Case when sum(OrWt)<>0 then sum(OrSalRt*OrWt)/sum(OrWt) else 0 end) as OrSalRt, " + _
'              " (select max(OrSetSCd) from OrdRm B where  B.OrCoCd= '" + gs_CoCd + "' " + _
'              " and B.OrTc = '" + GrdID.Value(GrdID.Row, "IdExpOdTc") + "' " + _
'              " and B.OrYy= '" + GrdID.Value(GrdID.Row, "IdExpOdYy") + "' " + _
'              " and B.OrChr= '" + GrdID.Value(GrdID.Row, "IdExpOdChr") + "' " + _
'              " and B.OrNo = " + CStr(GrdID.Value(GrdID.Row, "IdExpOdNo")) + _
'              " and B.OrSr= " + CStr(GrdID.Value(GrdID.Row, "IdExpOdSr")) + _
'              " and B.OrRmCd= '" + wRsRmDet!qIrRmCd + "' " + _
'              " and B.OrLn1 = " + CStr(wRsRmDet!qIrRmSz) + _
'              " and B.OrSetSalRt = max(A.OrSetSalRt)) as OrSetSCd, " + _
'              " (Case when sum(OrQty)<>0 then sum(OrSetCstRt*OrQty)/sum(OrQty) else 0 end) as OrSetCstRt, " + _
'              " (case when sum(OrQty)<>0 then sum(OrSetSalRt*OrQty)/sum(OrQty) else 0 end) as OrSetSalRt, " + _
'              " max(OrMainMet), max(OrRmCtg) " + _
'              " From OrdRm A where OrCoCd= '" + gs_CoCd + "' " + _
'              " and OrTc = '" + GrdID.Value(GrdID.Row, "IdExpOdTc") + "' " + _
'              " and OrYy= '" + GrdID.Value(GrdID.Row, "IdExpOdYy") + "' " + _
'              " and OrChr= '" + GrdID.Value(GrdID.Row, "IdExpOdChr") + "' " + _
'              " and OrNo = " + CStr(GrdID.Value(GrdID.Row, "IdExpOdNo")) + _
'              " and OrSr= " + CStr(GrdID.Value(GrdID.Row, "IdExpOdSr")) + _
'              " and OrRmCd= '" + wRsRmDet!qIrRmCd + "' " + _
'              " and OrLn1 = " + CStr(wRsRmDet!qIrRmSz) + _
'              " Group By OrRmCd, OrLn1 "
'
'
'    If (wOrRecSet.EOF Or wOrRecSet.BOF) Then
'     wOrStr = "Select OrCstRt, OrSalRt, OrSetSCd, OrSetCstRt, OrSetSalRt, OrMainMet, OrRmCtg " + _
'               "From OrdRm join RmRt on RrTcTyp= 'Rsz' and RrCmCd= '" + ctSelfCmCd + "' and " + _
'               "RrCtg= OrRmCtg and RrSCtg= OrRmSCtg " + _
'               "where OrCoCd= '" + gs_CoCd + "' " + _
'               " and OrTc = '" + GrdID.Value(GrdID.Row, "IdExpOdTc") + "' " + _
'               " and OrYy= '" + GrdID.Value(GrdID.Row, "IdExpOdYy") + "' " + _
'               " and OrChr = '" + GrdID.Value(GrdID.Row, "IdExpOdChr") + "' " + _
'               " and OrNo = " + CStr(GrdID.Value(GrdID.Row, "IdExpOdNo")) + _
'               " and OrSr= " + CStr(GrdID.Value(GrdID.Row, "IdExpOdSr")) + _
'               " and OrRmCd =  '" + wRsRmDet!qIrRmCd + "' " + _
'               " and OrLn1 >= RrFrLn and OrLn1 <= RrToLn and " + _
'               CStr(wRsRmDet!qIrRmSz) + " >= RrFrLn and " + _
'               CStr(wRsRmDet!qIrRmSz) + " <= RrToLn " + _
'               "Order by OrMainMet Desc, OrSalRt Desc "
'
'      Set wOrRecSet = moCn.OpenRes(wOrStr)
'    End If
'    If (wOrRecSet.EOF Or wOrRecSet.BOF) Then
'      wOrStr = "Select OrCstRt, OrSalRt, OrSetSCd, OrSetCstRt, OrSetSalRt, OrMainMet, OrRmCtg " + _
'               "From OrdRm where OrCoCd= '" + gs_CoCd + "' " + _
'               "  and OrTc = '" + GrdID.Value(GrdID.Row, "IdExpOdTc") + "' " + _
'               " and OrYy= '" + GrdID.Value(GrdID.Row, "IdExpOdYy") + "' " + _
'               " and OrChr= '" + GrdID.Value(GrdID.Row, "IdExpOdChr") + "' " + _
'               " and OrNo = " + CStr(GrdID.Value(GrdID.Row, "IdExpOdNo")) + _
'               " and OrSr= " + CStr(GrdID.Value(GrdID.Row, "IdExpOdSr")) + _
'               " and OrRmCd= '" + wRsRmDet!qIrRmCd + "' " + _
'               "Order by OrMainMet Desc, OrSalRt Desc"
'      Set wOrRecSet = moCn.OpenRes(wOrStr)
'    End If
'
'    If (wOrRecSet.EOF Or wOrRecSet.BOF) Then
'      wRsRmDet!qIrRmARt = 0#
'      wRsRmDet!qIrRmIRt = 0#
'      wRsRmDet!qIrSetSCd = ""
'      wRsRmDet!qIrSetARt = 0#
'      If UCase(ms_IdTcTyp) = UCase(ctInTcTyp) Then
'        wRsRmDet!qIrRmZRt = 0#
'        wRsRmDet!qIrSetIRt = 0#
'        wRsRmDet!qIrSetZRt = 0#
'      End If
'      wRsRmDet!qIrMainMet = "N"
'    Else
'      wOrRecSet.MoveFirst
'      wRsRmDet!qIrSetSCd = wOrRecSet!OrSetSCd
'      wRsRmDet!qIrSetARt = wOrRecSet!OrSetCstRt
'      If UCase(ms_IdTcTyp) = UCase(ctInTcTyp) Then
'        wRsRmDet!qIrRmARt = wOrRecSet!OrCstRt
'        wRsRmDet!qIrRmIRt = wOrRecSet!OrSalRt
'        wRsRmDet!qIrRmZRt = wOrRecSet!OrSalRt
'        wRsRmDet!qIrSetIRt = wOrRecSet!OrSetSalRt
'        wRsRmDet!qIrSetZRt = wOrRecSet!OrSetSalRt
'      End If
'
'      If wOrRecSet!OrMainMet = "Y" And wMainMet = False Then
'        wRsRmDet!qIrMainMet = "Y": wMainMet = True
'      ElseIf wOrRecSet!OrMainMet = "Y" And wMainMet = True Then
'        wRsRmDet!qIrMainMet = "N"
'      ElseIf wOrRecSet!OrMainMet = "N" Then
'        wRsRmDet!qIrMainMet = "N"
'      End If
'    End If
'
'    wRsRmDet.Update
'    wRsRmDet.MoveNext
'  Loop
'
'  If UCase(ms_IdTcTyp) = UCase(ctInTcTyp) Then
'
'    '****** check if mainMet wt > other Rm wt Diff. ***/
'    Set wRsRmDet = moCn.OpenRes("Select qIrRmCd, sum(qIrRmIWt) as qIrRmWt, " + _
'                                " max(qIrMainMet) as qMainMet from " + wTmpTbl + _
'                                " Group By qIrRmCd Order By qIrRmCd")
'    If (wRsRmDet.RecCount > 0) Then wRsRmDet.MoveFirst
'    wTotWtDiff = 0: wMainMetWt = 0
'    Do While Not (wRsRmDet.EOF Or wRsRmDet.BOF)
'        wFrOrdYN = Trim(mIWtEqOrd)
'        wTotIrRmWt = wRsRmDet!qIrRmWt
'        wRmCtg = moCn.GetFldVal("Select RmCtg from RmMst where RmCd= '" + wRsRmDet!qIrRmCd + "' ")
'        wDivFactor = IIF(wRmCtg = "C" Or wRmCtg = "D", 5, 1)
'
'        If wFrOrdYN = "N" Then wFrOrdYN = IIF(InStr(1, wCtgStr, "," + wRmCtg + ",") <> 0, "Y", "N")
'
'        If wFrOrdYN = "Y" Then
'            wOrStr = "Select sum(OrWt*" + CStr(mIdQty) + ") as qOrRmWt from OrdRm " + _
'                    "where OrCoCd= '" + gs_CoCd + "' " + _
'                    " and OrTc = '" + GrdID.Value(GrdID.Row, "IdExpOdTc") + "' " + _
'                    " and OrYy= '" + GrdID.Value(GrdID.Row, "IdExpOdYy") + "' " + _
'                    " and OrChr= '" + GrdID.Value(GrdID.Row, "IdExpOdChr") + "' " + _
'                    " and OrNo = " + CStr(GrdID.Value(GrdID.Row, "IdExpOdNo")) + _
'                    " and OrSr= " + CStr(GrdID.Value(GrdID.Row, "IdExpOdSr")) + _
'                    " and OrRmCd= '" + wRsRmDet!qIrRmCd + "' "
'
'            Set wOrRecSet = moCn.OpenRes(wOrStr)
'            If Not (wOrRecSet.EOF Or wOrRecSet.BOF) Then
'                wOrRecSet.MoveFirst
'                wWtDiff = 0
'                'wTotIrRmWt = wRsRmDet!qIrRmWt
'                If (wTotIrRmWt <> wOrRecSet!qOrRmWt) Then
'                    wWtDiff = wOrRecSet!qOrRmWt - wTotIrRmWt
'                    If UCase(wRsRmDet!qMainMet) <> "Y" Then _
'                        wTotWtDiff = wTotWtDiff + MWLib.Div(wWtDiff, wDivFactor)
'                End If
'            End If
'        End If
'
'        'urmi <>, = 23/02/04
'        If UCase(wRsRmDet!qMainMet) = "Y" Then _
'            wMainMetWt = MWLib.Div(wTotIrRmWt, wDivFactor)
'      wRsRmDet.MoveNext
'    Loop
'
'
'    'urmi <>, < 23/02/04
'    If wMainMetWt < wTotWtDiff Then
'
'        DispMsg "MainMetal Wt is less than the other Wt Diff.", etError
'        Exit Sub
'    End If
'
'    Set wRsRmDet = Nothing
'    Set wOrRecSet = Nothing
'    '********************************************************
'
'    Set wRsRmDet = moCn.OpenRes("Select qIrRmCd, sum(qIrRmIWt) as qIrRmWt, max(qIrMainMet) as qMainMet from " + wTmpTbl + _
'                                " Group By qIrRmCd Order By qIrRmCd")
'    If Not (wRsRmDet.EOF Or wRsRmDet.BOF) Then wRsRmDet.MoveFirst
'    wTotWtDiff = 0
'    Do While Not (wRsRmDet.EOF Or wRsRmDet.BOF)
'        wFrOrdYN = Trim(mIWtEqOrd)
'        wTotIrRmWt = wRsRmDet!qIrRmWt
'        wRmCtg = moCn.GetFldVal("Select RmCtg from RmMst where RmCd= '" + wRsRmDet!qIrRmCd + "' ")
'        If wFrOrdYN = "N" Then wFrOrdYN = IIF(InStr(1, wCtgStr, "," + wRmCtg + ",") <> 0, "Y", "N")
'
'        If wFrOrdYN = "Y" Then
'            wOrStr = "Select sum(OrWt*" + CStr(mIdQty) + ") as qOrRmWt from OrdRm " + _
'                    "where OrCoCd= '" + gs_CoCd + "' " + _
'                    " and OrTc = '" + GrdID.Value(GrdID.Row, "IdExpOdTc") + "' " + _
'                    " and OrYy= '" + GrdID.Value(GrdID.Row, "IdExpOdYy") + "' " + _
'                    " and OrChr= '" + GrdID.Value(GrdID.Row, "IdExpOdChr") + "' " + _
'                    " and OrNo = " + CStr(GrdID.Value(GrdID.Row, "IdExpOdNo")) + _
'                    " and OrSr= " + CStr(GrdID.Value(GrdID.Row, "IdExpOdSr")) + _
'                    " and OrRmCd= '" + wRsRmDet!qIrRmCd + "' "
'
'            Set wOrRecSet = moCn.OpenRes(wOrStr)
'            If Not (wOrRecSet.EOF Or wOrRecSet.BOF) Then
'                wOrRecSet.MoveFirst
'                wWtDiff = 0
'                'wTotIrRmWt = wRsRmDet!qIrRmWt
'
'                If (wTotIrRmWt <> wOrRecSet!qOrRmWt) Then
'                    wWtDiff = wOrRecSet!qOrRmWt - wTotIrRmWt
'                    'Debug.Print "wWtDiff = " + CStr(wWtDiff) + "'" + wRsRmDet!qIrRmCd + "' "
'                    wDivFactor = IIF(wRmCtg = "C" Or wRmCtg = "D", 5, 1)
'                    If UCase(wRsRmDet!qMainMet) = "Y" Then
'                        wTotWtDiff = wTotWtDiff + MWLib.Div(wWtDiff, wDivFactor)
'                    End If
'
'                    '###
'                    Set wIrRmWt = moCn.OpenResultset("Select * from " + wTmpTbl + _
'                                 " Where qIrRmCd = '" + wRsRmDet!qIrRmCd + "' ")
'
'                    If Not (wIrRmWt.EOF Or wIrRmWt.BOF) Then wIrRmWt.MoveFirst
'                    Do While Not (wIrRmWt.EOF Or wIrRmWt.BOF)
'                        wAddWt = (MWLib.Div(wWtDiff, wTotIrRmWt)) * wIrRmWt!qIrRmIWt
'                        wTotIrRmWt = wTotIrRmWt - wIrRmWt!qIrRmIWt
'                        wIrRmWt!qIrRmIWt = wIrRmWt!qIrRmIWt + wAddWt
'                        wIrRmWt.Update
'                        wWtDiff = wWtDiff - wAddWt
'                        wIrRmWt.MoveNext
'                    Loop
'                End If
'            End If
'
'        End If
'        wRsRmDet.MoveNext
'    Loop
'
'
'    Set wRsRmDet = moCn.OpenResultset("Select * from " + wTmpTbl + _
'                                      " where qIrMainMet= 'Y'")
'    If Not (wRsRmDet.EOF Or wRsRmDet.BOF) Then
'        wRsRmDet.MoveFirst
'        wRsRmDet!qIrRmIWt = wRsRmDet!qIrRmIWt - wTotWtDiff
'        wRsRmDet.Update
'    End If
'
'  End If
'
'
'  If UCase(ms_IdTcTyp) = UCase(ctInTcTyp) Then
'    moCn.Execute ("Update " + wTmpTbl + " Set qXtraIWt= .000, qXtraZWt= .000 ")
'
'
''********* URMI ****** 23/02/04 (commented code) *********
'    wDFrOrdYN = IIF(Trim(mIWtEqOrd) = "Y", "Y", "N")
'    If wDFrOrdYN = "N" Then wDFrOrdYN = IIF(InStr(1, wCtgStr, "," + "D" + ",") <> 0, "Y", "N")
'
'    wCFrOrdYN = IIF(Trim(mIWtEqOrd) = "Y", "Y", "N")
'    If wCFrOrdYN = "N" Then wCFrOrdYN = IIF(InStr(1, wCtgStr, "," + "C" + ",") <> 0, "Y", "N")
'
'    wDRrTcTyp = IIF(wDFrOrdYN = "Y", "XIO", "XIA")
'    wCRrTcTyp = IIF(wCFrOrdYN = "Y", "XIO", "XIA")
'
'    Call GetXtraWt(wDRrTcTyp, wCRrTcTyp, wTmpTbl, wCmDiaRtFA, wCmCsRtFA, GrdID.Value(GrdID.Row, "IdExpCmCd"), wDmCtg, _
'                   "qIrRmIWt", "qXtraIWt", "qIrRmIRt", mIdQty)
'    If whZWtEq = "A" Then
'      Call GetXtraWt("XZA", "XZA", wTmpTbl, wCmDiaRtFA, wCmCsRtFA, GrdID.Value(GrdID.Row, "IdExpCmCd"), wDmCtg, _
'                     "qIrRmZWt", "qXtraZWt", "qIrRmZRt", mIdQty)
'    End If
'
'
'    '*********************************************************
'
'  End If
'
'  Set wRsRmDet = moCn.OpenRes("select * from " + wTmpTbl + " Order By qIrMainMet")
'
'  wTotIXtraWt = 0: wTotZXtraWt = 0
'  wTotAWt = 0: wTotIWt = 0: wTotZWt = 0
'
'  With GrdIR
'    .StartCopy
'    Do While Not (wRsRmDet.EOF Or wRsRmDet.BOF)
'      .AddItem
'      wRmRow = .Rows - 1
'      If UCase(ms_IdTcTyp) = UCase(ctInTcTyp) Then
'        wDivFactor = IIF(wRsRmDet!qRmCtg = "C" Or wRsRmDet!qRmCtg = "D", 5, 1)
'        wTotIXtraWt = wTotIXtraWt + MWLib.Div(wRsRmDet!qXtraIWt, wDivFactor)
'        wTotZXtraWt = wTotZXtraWt + MWLib.Div(wRsRmDet!qXtraZWt, wDivFactor)
'
'        .Value(wRmRow, "IrSrNo") = wRmRow
'        .Value(wRmRow, "IrRmCd") = wRsRmDet!qIrRmCd
'        .Value(wRmRow, "IrRmSz") = wRsRmDet!qIrRmSz
'        .Value(wRmRow, "IrStkRt") = wRsRmDet!qIrStkRt
'        .Value(wRmRow, "IrRmQty") = wRsRmDet!qIrRmQty
'        .Value(wRmRow, "IrRmAWt") = wRsRmDet!qIrRmAWt
'        .Value(wRmRow, "IrRmARt") = wRsRmDet!qIrRmARt
'        .Value(wRmRow, "IrSetSCd") = wRsRmDet!qIrSetSCd
'        .Value(wRmRow, "IrSetARt") = wRsRmDet!qIrSetARt
'        .Value(wRmRow, "IrSetIRt") = wRsRmDet!qIrSetIRt
'        .Value(wRmRow, "IrMainMet") = wRsRmDet!qIrMainMet
'        .Value(wRmRow, "IrSetIRt") = wRsRmDet!qIrSetIRt
'        .Value(wRmRow, "IrSetZRt") = wRsRmDet!qIrSetZRt
'
'        If wRsRmDet!qIrMainMet = "Y" Then
'            'Debug.Print "IWt = " + CStr(wRsRmDet!qIrRmIWt)
'           .Value(wRmRow, "IrRmIWt") = wRsRmDet!qIrRmIWt - (wTotIWt - wTotAWt)
'          If whZWtEq = "A" Then
'            .Value(wRmRow, "IrRmZWt") = wRsRmDet!qIrRmZWt - (wTotZWt - wTotAWt)
'          ElseIf whZWtEq = "I" Then
'            .Value(wRmRow, "IrRmZWt") = wRsRmDet!qIrRmIWt - (wTotIWt - wTotAWt)
'          End If
'        Else
'
'          .Value(wRmRow, "IrRmIWt") = wRsRmDet!qIrRmIWt + wRsRmDet!qXtraIWt
'
'          If whZWtEq = "A" Then
'            .Value(wRmRow, "IrRmZWt") = wRsRmDet!qIrRmZWt + wRsRmDet!qXtraZWt
'          ElseIf whZWtEq = "I" Then
'            .Value(wRmRow, "IrRmZWt") = wRsRmDet!qIrRmIWt + wRsRmDet!qXtraIWt
'          End If
'        End If
'
'        wTotAWt = wTotAWt + MWLib.Div(.Value(wRmRow, "IrRmAWt"), wDivFactor)
'        wTotIWt = wTotIWt + MWLib.Div(.Value(wRmRow, "IrRmIWt"), wDivFactor)
'        wTotZWt = wTotZWt + MWLib.Div(.Value(wRmRow, "IrRmZWt"), wDivFactor)
'
'        If UCase(wRsRmDet!qRmCtg) = "G" And mIGldRtEqOrd <> "Y" Then
'          .Value(wRmRow, "IrRmIRt") = GetGldRt(GrdID.Value(GrdID.Row, "IdExpCmCd"), ADC("wIdLmgSal"), _
'                                      wRsRmDet!qIrRmCd, 0, "DM", GrdID.Value(GrdID.Row, "IdDmCd"), "S", "C")
'          If mLocalExp = "N" Then .Value(wRmRow, "IrRmZRt") = .Value(wRmRow, "IrRmIRt")
'        ElseIf UCase(wRsRmDet!qRmCtg) = "P" And mIGldRtEqOrd <> "Y" Then
'          .Value(wRmRow, "IrRmIRt") = GetGldRt(GrdID.Value(GrdID.Row, "IdExpCmCd"), ADC("wIdLmpSal"), _
'                                      wRsRmDet!qIrRmCd, 0, "DM", GrdID.Value(GrdID.Row, "IdDmCd"), "S", "C")
'          If mLocalExp = "N" Then .Value(wRmRow, "IrRmZRt") = .Value(wRmRow, "IrRmIRt")
'        Else
'          .Value(wRmRow, "IrRmIRt") = wRsRmDet!qIrRmIRt
'          If mLocalExp = "N" Or (mLocalExp = "Y" And UCase(wRsRmDet!qRmCtg) <> "G" And UCase(wRsRmDet!qRmCtg) <> "P") Then
'            .Value(wRmRow, "IrRmZRt") = wRsRmDet!qIrRmZRt
'          End If
'        End If
'        .Value(wRmRow, "IrRmZPureRt") = IIF(wRsRmDet!qRmPurityZ > 0, MWLib.Div(wRsRmDet!qIrRmZRt, wRsRmDet!qRmPurityZ), wRsRmDet!qIrRmZRt)
'      ElseIf UCase(ms_IdTcTyp) = UCase(ctCTBTcTyp) Then
'        .Value(wRmRow, "IrSrNo") = wRmRow
'        .Value(wRmRow, "IrRmCd") = wRsRmDet!qIrRmCd
'        .Value(wRmRow, "IrRmSz") = wRsRmDet!qIrRmSz
'        .Value(wRmRow, "IrStkRt") = wRsRmDet!qIrStkRt
'        .Value(wRmRow, "IrRmQty") = wRsRmDet!qIrRmQty
'        .Value(wRmRow, "IrRmAWt") = wRsRmDet!qIrRmAWt
'
'        .Value(wRmRow, "IrRmIWt") = wRsRmDet!qIrRmIWt
'
'        .Value(wRmRow, "IrRmZWt") = wRsRmDet!qIrRmZWt
'        '.Value(wRmRow, "IrRmARt") = wRsRmDet!qIrRmARt
'        '.Value(wRmRow, "IrRmIRt") = wRsRmDet!qIrRmIRt
'        '.Value(wRmRow, "IrRmZRt") = wRsRmDet!qIrRmZRt
'        .Value(wRmRow, "IrSetSCd") = wRsRmDet!qIrSetSCd
'        .Value(wRmRow, "IrSetARt") = wRsRmDet!qIrSetARt
'        '.Value(wRmRow, "IrSetIRt") = wRsRmDet!qIrSetIRt
'        '.Value(wRmRow, "IrSetZRt") = wRsRmDet!qIrSetZRt
'        .Value(wRmRow, "IrRmZPureRt") = 0
'        .Value(wRmRow, "IrMainMet") = wRsRmDet!qIrMainMet
'      End If
'      .SaveRec (wRmRow)
'      wRsRmDet.MoveNext
'    Loop
'    .EndCopy
'  End With
'
'  If UCase(ms_IdTcTyp) = UCase(ctInTcTyp) Then
'    wf_CnvRt = 1
'  ElseIf UCase(ms_IdTcTyp) = UCase(ctCTBTcTyp) Then
'    wf_CnvRt = moCn.GetFldVal("Select InCnvRt from InvHd where " + _
'                              "InCoCd= '" + gs_CoCd + "' and InTc= '" + ADC("wIdTc") + "' " + _
'                              "and InYy= '" + ADC("wIdYy") + "' and InChr= '" + ADC("wIdChr") + "' " + _
'                              "and InNo= " + CStr(ADC("wIdNo")))
'  End If
'
'  wf_AddLab = moCn.GetFldVal("Select sum(IrRmIVal- IrRmZVal* " + CStr(wf_CnvRt) + ") from InvRm " + _
'              "join RmMst on RmCd= IrRmCd where IrCoCd= '" + gs_CoCd + "' and " + _
'              "IrTc= '" + ADC("wIdTc") + "' and IrYy= '" + ADC("wIdYy") + "' and " + _
'              "IrChr= '" + ADC("wIdChr") + "' and IrNo= " + CStr(ADC("wIdNo")) + " and " + _
'              "IrSr= " + CStr(GrdID.Value(GrdID.Row, "IdSr")) + " and RmCtg in ('G', 'P')")
'
'  If mLocalExp = "Y" And Round(MWLib.Div(wf_AddLab, wf_CnvRt), 3) <> 0 Then
'    'And ms_IdTcTyp = ctInTcTyp
'    With GrdIL
'      .StartCopy
'      .AddItem
'      wLabRow = .Rows - 1
'      .Value(wLabRow, "iLSrNo") = wLabRow
'      .Value(wLabRow, "iLMCd") = ctAddLMCd
'      .Value(wLabRow, "iLSCd") = ctAddLSCd
'      .Value(wLabRow, "iLQty") = 1
'      .Value(wLabRow, "iLARt") = 0
'      .Value(wLabRow, "iLiRt") = 0
'      .Value(wLabRow, "iLZRt") = MWLib.Div(wf_AddLab, wf_CnvRt)
'      .Value(wLabRow, "ilQw") = "Q"
'      .Value(wLabRow, "iLZVal") = MWLib.Div(wf_AddLab, wf_CnvRt)
'      .SaveRec (wLabRow)
'      .EndCopy
'    End With
'  End If
'
'  Set wRsRmDet = Nothing
'  Set wOrRecSet = Nothing
End Sub
Public Sub MakeRmFromFg2(ByVal mIWtEqOrd As String, ByVal mIWtFrOrd As String, _
                        ByVal mIGldRtEqOrd As String, ByVal mIdQty As Double)
  '*** (Bef 2.11)
'  Dim whZWtEq As String, wCtgStr As String, wDmCtg As String, wTmpTbl As String
'  Dim wOrStr As String, wStr As String, wMainMet As Boolean
'  Dim wRsRmDet As MDORowSet, wOrRecSet As MDORowSet
'  Dim wDivFactor As Single, wRmRow As Single, wLabRow As Single, wf_AddLab As Single
'  Dim wTotIXtraWt As Double, wTotZXtraWt As Double, wTotAWt As Double
'  Dim wTotIWt As Double, wTotZWt As Double
'  Dim wCmDiaRtFA As String, wCmCsRtFA As String, wf_CnvRt As Single
'  Dim wDRrTcTyp As String, wCRrTcTyp As String, wDFrOrdYN As String
'  Dim wCFrOrdYN As String
'  Dim wSetRt As Boolean
'  Dim wTotWtDiff As Double, wWtDiff As Double, wTotIrRmWt As Double, wAddWt As Double
'  Dim wIrRmWt As MDORowSet, wRmCtg As String, wFrOrdYN As String, wMainMetWt As Double
'  Dim wFrRmWt As Double, wFrRmQty As Double, wOrTotWt As Double, wOrTotQty As Double
'  Dim wFRmDet As MDORowSet
'  whZWtEq = moCn.GetFldVal("Select hZWtEq from Head where HCoCd= '" + gs_CoCd + "' and HCd='" + ctSelfCmCd + "'")
'  Dim wSetTyp As Boolean
'  Dim wOrLnStr As String
'  'mIWtFrOrd = "DC"
'  'wCtgStr= ",D,C"
'  wCtgStr = "," + moCn.GetFldVal("Select PDesc225 from Param where " + _
'              "PTyp='IWTFRORD' and PMCd = '" + mIWtFrOrd + "'") + ","
'
'  wCmDiaRtFA = moCn.GetFldVal("Select CmDiaRtFA from CustMst where CmCtg= 'C' and " + _
'               "CmCd= '" + GrdID.Value(GrdID.Row, "IdExpCmCd") + "' ")
'
'  wCmCsRtFA = moCn.GetFldVal("Select CmCSRtFA from CustMst where CmCtg= 'C' and " + _
'               "CmCd= '" + GrdID.Value(GrdID.Row, "IdExpCmCd") + "' ")
'
'
'  wDmCtg = moCn.GetFldVal("Select DmCtg from DsgMst where DmTcTyp= 'DM' and DmSz= '' and " + _
'               "DmCd= '" + GrdID.Value(GrdID.Row, "IdDmCd") + "' ")
'  wTmpTbl = "#TmpInvRm"
'
'  '*******Geeta*******Emr207
'  Dim ws_OrLn As String
'    ws_OrLn = "(case when Pvalue3='Y' then (Case when PValue2='Y' then  " + _
'            "cast(OrRmPtr as Decimal(12,4)) else (case when OrPrdQty=0 then 0 else  " + _
'            "cast((OrPrdWt/OrPrdQty) as Decimal(12,4)) end ) end )  " + _
'            "else cast(Orln1  as Decimal(12,4)) end ) "
'  If UCase(ms_IdTcTyp) = UCase(ctInTcTyp) Then
'
'     wStr = "Select FrRmCd as qIrRmCd, FrRmSz as qIrRmSz,sum(FrRmStkRt*FrRmWt)/(case when sum(FrRmWt)> 0  " + _
'           "then sum(FrRmWt) else 1 end) as qIrStkRt, " + _
'           "sum(FrRmQty) as qIrRmQty, sum(FrRmWt) As qIrRmAWt, sum(FrRmWt) as qIrRmIWt, " + _
'           "sum(FrRmWt) as qIrRmZWt,max(Isnull(qOrWt,0)) as qOrWt,max(IsNull(qOrQty,0)) as qOrQty, 9999.999 as qXtraIWt, 9999.999 as qXtraZWt, " + _
'           "999999.999 as qIrRmARt, 999999.999 as qIrRmIRt, 999999.999 as qIrRmZRt, " + _
'           "space(8) as qIrSetSCd,isnull(qOrSetSalRt,0.00) as qIrSetARt,isnull(qOrSetSalRt,0.00) as qIrSetIRt,isnull(qOrSetSalRt,0.00) as qIrSetZRt, " + _
'           "Space(1) as qIrMainMet, max(RmPurityZ) as qRmPurityZ, max(RmCtg) as qRmCtg, max(RmSCtg) as qRmSCtg " + _
'           "into " + wTmpTbl + " from InvFgd join FgRm on FrCoCd= IfCoCd and FrTc= IfFdTc and FrYy= IfFdyy and " + _
'           "FrChr= IfFdChr and FrNo = IfFdNo And FrSr = IfFdSr " + _
'           "Left Outer Join (select OrRmcd as qOrRmcd ," + ws_OrLn + "as qOrLn1, " + _
'           "OrSetSalRt as qOrSetSalRt,sum(OrPrdWt) as qOrWt,sum(OrPrdQty) as qOrQty  from OrdRm  " + _
'           "Join RmMst on RmCd=OrRmCd " + _
'           "Join Param on PTyp='RMSCTG' and PMCd=RmCtg and PSCd=RmSCtg " + _
'           "where OrCoCd= '" + gs_CoCd + "'  and OrTc = '" + GrdID.Value(GrdID.Row, "IdExpOdTc") + "' " + _
'           "and OrYy= '" + GrdID.Value(GrdID.Row, "IdExpOdYy") + "' " + _
'           " and OrChr= '" + GrdID.Value(GrdID.Row, "IdExpOdChr") + "' and OrNo = " + CStr(GrdID.Value(GrdID.Row, "IdExpOdNo")) + _
'           " and OrSr= " + CStr(GrdID.Value(GrdID.Row, "IdExpOdSr")) + " group by OrCocd ,OrTc,OrYy,OrChr,OrNo,OrSr, " + _
'           "OrRmcd," + ws_OrLn + ",OrSetSalRt) Or2  On qOrRmcd=FrRmcd and qOrln1=FrRmSz " + _
'           "Join RmMst on FrRmCd = RmCd " + _
'           "where IFcocd = '" + gs_CoCd + "' and IFTc= '" + ADC("wIdTc") + "' and IFYy= '" + _
'           ADC("wIdYy") + "' and IFChr= '" + ADC("wIdChr") + "' and " + _
'           "IFNo = " + CStr(ADC("wIdNo")) + " and IfSr = " + CStr(GrdID.Value(GrdID.Row, "IdSr")) + " " + _
'           "Group by FrRmCd, FrRmSz,qOrSetSalRt "
'
' ElseIf UCase(ms_IdTcTyp) = UCase(ctCTBTcTyp) Then
'
'    wStr = "Select FrRmCd as qIrRmCd, FrRmSz as qIrRmSz,sum(FrRmStkRt*FrRmWt)/(case when sum(FrRmWt)> 0  " + _
'           "then sum(FrRmWt) else 1 end) as qIrStkRt, " + _
'           "sum(FrRmQty) as qIrRmQty, sum(FrRmWt) As qIrRmAWt, sum(FrRmWt) as qIrRmIWt, " + _
'           "sum(FrRmWt) as qIrRmZWt,max(Isnull(qOrWt,0)) as qOrWt,max(IsNull(qOrQty,0)) as qOrQty," + _
'           "0 as qXtraIWt, 0 as qXtraZWt, " + _
'           "0 as qIrRmARt, 0 as qIrRmIRt, 0 as qIrRmZRt, " + _
'           "space(8) as qIrSetSCd, isnull(qOrSetSalRt,0.00) as qIrSetARt, isnull(qOrSetSalRt,0.00) as qIrSetIRt, 0 as qIrSetZRt, " + _
'           "Space(1) as qIrMainMet, max(RmPurityZ) as qRmPurityZ, max(RmCtg) as qRmCtg, max(RmSCtg) as qRmSCtg " + _
'           "into " + wTmpTbl + " " + _
'           "from InvFgd join FgRm on FrCoCd= IfCoCd and FrTc= IfFdTc and FrYy= IfFdyy and " + _
'           "FrChr= IfFdChr and FrNo = IfFdNo And FrSr = IfFdSr " + _
'           "Left Outer Join (select OrRmcd as qOrRmcd ," + ws_OrLn + " as qOrLn1,OrSetSalRt as qOrSetSalRt," + _
'           "sum(OrPrdWt) as qOrWt,sum(OrPrdQty) as qOrQty  from OrdRm  " + _
'           "Join RmMst on RmCd=OrRmCd " + _
'           "Join Param on PTyp='RMSCTG' and PMCd=RmCtg and PSCd=RmSCtg " + _
'           "where OrCoCd= '" + gs_CoCd + "'  and OrTc = '" + GrdID.Value(GrdID.Row, "IdExpOdTc") + "' " + _
'           "and OrYy= '" + GrdID.Value(GrdID.Row, "IdExpOdYy") + "' " + _
'           " and OrChr= '" + GrdID.Value(GrdID.Row, "IdExpOdChr") + "' and OrNo = " + CStr(GrdID.Value(GrdID.Row, "IdExpOdNo")) + _
'           " and OrSr= " + CStr(GrdID.Value(GrdID.Row, "IdExpOdSr")) + " group by OrCocd ,OrTc,OrYy,OrChr,OrNo,OrSr, " + _
'           "OrRmcd," + ws_OrLn + ",OrSetSalRt) Or2  On qOrRmcd=FrRmcd and qOrln1=FrRmSz " + _
'           "Join RmMst on FrRmCd = RmCd " + _
'           "where IfCoCd = '" + gs_CoCd + "' and IfTc= '" + ADC("wIdTc") + "' " + _
'           "and IfYy= '" + ADC("wIdYy") + "' and IfChr= '" + ADC("wIdChr") + "' and " + _
'           "IfNo = " + CStr(ADC("wIdNo")) + " and IfSr = " + CStr(GrdID.Value(GrdID.Row, "IdSr")) + " " + _
'           "Group by FrRmCd, FrRmSz,qOrSetSalRt "
'     End If
'  '**********Geeta******Emr207
'
'  '*** Creating Temporary Table #TmpInvRm ***
'  moCn.CreateTmpTable wTmpTbl, wStr
'  'moCn.Execute (wStr)
'
'  Set wRsRmDet = moCn.OpenResultset("Select * from " + wTmpTbl)
'
'  '**********Geeta******Emr207
'    If moCn.RecSeek("select count(distinct qIrSetIRt) From " + wTmpTbl + " " + _
'                      "group by qIrRmCd,qIrRmSz Having count(distinct qIrSetIRt)>1") Then
'    '  To get Proportionate weights
'     Set wFRmDet = moCn.OpenRes("Select FrRmCd as qFrRmCd,FrRmsz as qFrRmsz, " + _
'                            " Sum(FrRmWt) as qFrRmWt,Sum(FrRmQty) as qFrRmQty from   FgRm  " + _
'                            "join InvFgd on FrCoCd= IfCoCd and FrTc= IfFdTc and FrYy= IfFdyy and  " + _
'                            " FrChr= IfFdChr and FrNo = IfFdNo And FrSr = IfFdSr " + _
'                            "where IFcocd = '" + gs_CoCd + "' and IFTc= '" + ADC("wIdTc") + "' and IFYy= '" + _
'                            ADC("wIdYy") + "' and IFChr= '" + ADC("wIdChr") + "' and " + _
'                            "IFNo = " + CStr(ADC("wIdNo")) + " and IfSr = " + CStr(GrdID.Value(GrdID.Row, "IdSr")) + " Group By FrRmCd,FrRmsz Order By FrRmCd,FrRmsz")
'
'
'
'     Do While Not (wFRmDet.EOF Or wFRmDet.BOF)
'        wFrRmWt = wFRmDet!qFrRmWt
'        wFrRmQty = wFRmDet!qFrRmQty
'        wOrTotWt = moCn.GetFldVal("Select Sum(qOrWt) From " + wTmpTbl + " Where qIrRmCd='" + wFRmDet!qFrRmCd + "' and qIrRmSz='" + CStr(wFRmDet!qFrRmSz) + "' Group By qIrRmcd,qIrRmsz")
'        wOrTotQty = moCn.GetFldVal("Select Sum(qOrQty) From " + wTmpTbl + "  Where qIrRmCd='" + wFRmDet!qFrRmCd + "' and qIrRmSz='" + CStr(wFRmDet!qFrRmSz) + "' Group By qIrRmCd,qIrRmsz")
'        If (wRsRmDet.RecCount > 0) Then wRsRmDet.MoveFirst
'     Do While Not (wRsRmDet.EOF Or wRsRmDet.BOF)
'         If wFRmDet!qFrRmCd = wRsRmDet!qIrRmCd And wFRmDet!qFrRmSz = wRsRmDet!qIrRmSz Then
'           If moCn.RecSeek("select count(distinct qIrSetIRt) From " + wTmpTbl + " " + _
'                           "Where qIrRmCd= '" + wRsRmDet!qIrRmCd + "' " + _
'                           " and qIrRmSz = " + CStr(wRsRmDet!qIrRmSz) + _
'                           "group by qIrRmCd,qIrRmSz Having count(distinct qIrSetIRt)>1") Then
'                   wRsRmDet.edit
'                   wRsRmDet!qIrRmIWt = Round(MWLib.Div(wRsRmDet!qOrWt, wOrTotWt) * wFrRmWt, 3)
'                   wRsRmDet!qIrRmAWt = wRsRmDet!qIrRmIWt
'                   wRsRmDet!qIrRmZWt = wRsRmDet!qIrRmIWt
'                   wOrTotWt = wOrTotWt - wRsRmDet!qOrWt
'                   wFrRmWt = wFrRmWt - wRsRmDet!qIrRmIWt
'                   wRsRmDet!qIrRmQty = Round(MWLib.Div(wRsRmDet!qOrQty, wOrTotQty) * wFrRmQty)
'                   wOrTotQty = wOrTotQty - wRsRmDet!qOrQty
'                   wFrRmQty = wFrRmQty - wRsRmDet!qIrRmQty
'                   wRsRmDet.Update
'
'           End If
'         End If
'       wRsRmDet.MoveNext
'     Loop
'       wFRmDet.MoveNext
'     Loop
'  Set wFRmDet = Nothing
'  End If
'   If (wRsRmDet.RecCount > 0) Then wRsRmDet.MoveFirst
'   Do While Not (wRsRmDet.EOF Or wRsRmDet.BOF)
'   'Debug.Print CStr(wRsRmDet!qIrRmCd) + "," + CStr(wRsRmDet!qIrRmSz) + "," + CStr(wRsRmDet!qIrRmQty) + "," + CStr(wRsRmDet!qIrRmAWt) + "," + CStr(wRsRmDet!qOrWt) + "," + CStr(wRsRmDet!qOrQty)
'        wRsRmDet.MoveNext
'    Loop
'  '**********Geeta******Emr207
'  If (wRsRmDet.RecCount > 0) Then wRsRmDet.MoveFirst
'  Do While Not (wRsRmDet.EOF Or wRsRmDet.BOF)
'    wRsRmDet.edit
'
'    wOrLnStr = IIF(moCn.GetFldVal("select PValue3 From Param Join  RmMst " + _
'                            " on PTyp='RMSCTG' and PMCd=RmCtg and PSCd=RmSCtg " + _
'                            "where RmCd='" + wRsRmDet!qIrRmCd + "' ") = "Y", _
'                   IIF(moCn.GetFldVal("select PValue2 From Param Join  RmMst " + _
'                            "on PTyp='RMSCTG' and PMCd=RmCtg and PSCd=RmSCtg " + _
'                            "where RmCd='" + wRsRmDet!qIrRmCd + "' ") = "Y", "OrRmPtr", _
'                            " case when OrPrdQty =0 then 0 else OrPrdWt/OrPrdQty end "), "OrLn1")
'
'    If moCn.RecSeek("select count(distinct qIrSetIRt) From " + wTmpTbl + " " + _
'                    "Where qIrRmCd= '" + wRsRmDet!qIrRmCd + "' " + _
'                    " and qIrRmSz = " + CStr(wRsRmDet!qIrRmSz) + _
'                    "group by qIrRmCd,qIrRmSz Having count(distinct qIrSetIRt)>1") Then
'
'    '*******Geeta*******Emr207
'
'     wOrStr = "Select OrCstRt, OrSalRt, OrSetSCd, OrSetCstRt,OrSetSalRt, OrMainMet, OrRmCtg " + _
'             "From OrdRm where OrCoCd= '" + gs_CoCd + "' " + _
'             " and OrTc = '" + GrdID.Value(GrdID.Row, "IdExpOdTc") + "' " + _
'             " and OrYy= '" + GrdID.Value(GrdID.Row, "IdExpOdYy") + "' " + _
'             " and OrChr= '" + GrdID.Value(GrdID.Row, "IdExpOdChr") + "' " + _
'             " and OrNo = " + CStr(GrdID.Value(GrdID.Row, "IdExpOdNo")) + _
'             " and OrSr= " + CStr(GrdID.Value(GrdID.Row, "IdExpOdSr")) + _
'             " and OrRmCd= '" + wRsRmDet!qIrRmCd + "' " + _
'             " and " + wOrLnStr + " = " + CStr(wRsRmDet!qIrRmSz) + _
'             " and OrSetSalRt=" + CStr(wRsRmDet!qIrSetIRt) + _
'             " Order by OrMainMet Desc, OrSalRt Desc"
'   ' Set wOrRecSet = moCn.OpenRes(wOrStr)
'
'    Else
'      wOrStr = "Select OrCstRt, OrSalRt, OrSetSCd, OrSetCstRt, OrSetSalRt, OrMainMet, OrRmCtg " + _
'                "From OrdRm where OrCoCd= '" + gs_CoCd + "' " + _
'                " and OrTc = '" + GrdID.Value(GrdID.Row, "IdExpOdTc") + "' " + _
'                " and OrYy= '" + GrdID.Value(GrdID.Row, "IdExpOdYy") + "' " + _
'                " and OrChr= '" + GrdID.Value(GrdID.Row, "IdExpOdChr") + "' " + _
'                " and OrNo = " + CStr(GrdID.Value(GrdID.Row, "IdExpOdNo")) + _
'                " and OrSr= " + CStr(GrdID.Value(GrdID.Row, "IdExpOdSr")) + _
'                " and OrRmCd= '" + wRsRmDet!qIrRmCd + "' " + _
'                " and " + wOrLnStr + " = " + CStr(wRsRmDet!qIrRmSz) + _
'                "Order by OrMainMet Desc, OrSalRt Desc"
'    End If
'    Set wOrRecSet = moCn.OpenRes(wOrStr)
'
'
''    Query to get weighted Average rate
''    wOrStr = "Select (Case when sum(OrWt)<>0 then sum(OrCstRt*OrWt)/sum(OrWt) else 0 end) as OrCstRt, " + _
'              " (Case when sum(OrWt)<>0 then sum(OrSalRt*OrWt)/sum(OrWt) else 0 end) as OrSalRt, " + _
'              " (select max(OrSetSCd) from OrdRm B where  B.OrCoCd= '" + gs_CoCd + "' " + _
'              " and B.OrTc = '" + GrdID.Value(GrdID.Row, "IdExpOdTc") + "' " + _
'              " and B.OrYy= '" + GrdID.Value(GrdID.Row, "IdExpOdYy") + "' " + _
'              " and B.OrChr= '" + GrdID.Value(GrdID.Row, "IdExpOdChr") + "' " + _
'              " and B.OrNo = " + CStr(GrdID.Value(GrdID.Row, "IdExpOdNo")) + _
'              " and B.OrSr= " + CStr(GrdID.Value(GrdID.Row, "IdExpOdSr")) + _
'              " and B.OrRmCd= '" + wRsRmDet!qIrRmCd + "' " + _
'              " and B.OrLn1 = " + CStr(wRsRmDet!qIrRmSz) + _
'              " and B.OrSetSalRt = max(A.OrSetSalRt)) as OrSetSCd, " + _
'              " (Case when sum(OrQty)<>0 then sum(OrSetCstRt*OrQty)/sum(OrQty) else 0 end) as OrSetCstRt, " + _
'              " (case when sum(OrQty)<>0 then sum(OrSetSalRt*OrQty)/sum(OrQty) else 0 end) as OrSetSalRt, " + _
'              " max(OrMainMet), max(OrRmCtg) " + _
'              " From OrdRm A where OrCoCd= '" + gs_CoCd + "' " + _
'              " and OrTc = '" + GrdID.Value(GrdID.Row, "IdExpOdTc") + "' " + _
'              " and OrYy= '" + GrdID.Value(GrdID.Row, "IdExpOdYy") + "' " + _
'              " and OrChr= '" + GrdID.Value(GrdID.Row, "IdExpOdChr") + "' " + _
'              " and OrNo = " + CStr(GrdID.Value(GrdID.Row, "IdExpOdNo")) + _
'              " and OrSr= " + CStr(GrdID.Value(GrdID.Row, "IdExpOdSr")) + _
'              " and OrRmCd= '" + wRsRmDet!qIrRmCd + "' " + _
'              " and OrLn1 = " + CStr(wRsRmDet!qIrRmSz) + _
'              " Group By OrRmCd, OrLn1 "
'
'
'    If (wOrRecSet.EOF Or wOrRecSet.BOF) Then
'     wOrStr = "Select OrCstRt, OrSalRt, OrSetSCd, OrSetCstRt, OrSetSalRt, OrMainMet, OrRmCtg " + _
'               "From OrdRm join RmRt on RrTcTyp= 'Rsz' and RrCmCd= '" + ctSelfCmCd + "' and " + _
'               "RrCtg= OrRmCtg and RrSCtg= OrRmSCtg " + _
'               "where OrCoCd= '" + gs_CoCd + "' " + _
'               " and OrTc = '" + GrdID.Value(GrdID.Row, "IdExpOdTc") + "' " + _
'               " and OrYy= '" + GrdID.Value(GrdID.Row, "IdExpOdYy") + "' " + _
'               " and OrChr = '" + GrdID.Value(GrdID.Row, "IdExpOdChr") + "' " + _
'               " and OrNo = " + CStr(GrdID.Value(GrdID.Row, "IdExpOdNo")) + _
'               " and OrSr= " + CStr(GrdID.Value(GrdID.Row, "IdExpOdSr")) + _
'               " and OrRmCd =  '" + wRsRmDet!qIrRmCd + "' " + _
'               " and " + wOrLnStr + " >= RrFrLn and " + wOrLnStr + " <= RrToLn and " + _
'               CStr(wRsRmDet!qIrRmSz) + " >= RrFrLn and " + _
'               CStr(wRsRmDet!qIrRmSz) + " <= RrToLn " + _
'               "Order by OrMainMet Desc, OrSalRt Desc "
'
'      Set wOrRecSet = moCn.OpenRes(wOrStr)
'    End If
'    If (wOrRecSet.EOF Or wOrRecSet.BOF) Then
'      wOrStr = "Select OrCstRt, OrSalRt, OrSetSCd, OrSetCstRt, OrSetSalRt, OrMainMet, OrRmCtg " + _
'               "From OrdRm where OrCoCd= '" + gs_CoCd + "' " + _
'               "  and OrTc = '" + GrdID.Value(GrdID.Row, "IdExpOdTc") + "' " + _
'               " and OrYy= '" + GrdID.Value(GrdID.Row, "IdExpOdYy") + "' " + _
'               " and OrChr= '" + GrdID.Value(GrdID.Row, "IdExpOdChr") + "' " + _
'               " and OrNo = " + CStr(GrdID.Value(GrdID.Row, "IdExpOdNo")) + _
'               " and OrSr= " + CStr(GrdID.Value(GrdID.Row, "IdExpOdSr")) + _
'               " and OrRmCd= '" + wRsRmDet!qIrRmCd + "' " + _
'               "Order by OrMainMet Desc, OrSalRt Desc"
'      Set wOrRecSet = moCn.OpenRes(wOrStr)
'    End If
'
'    If (wOrRecSet.EOF Or wOrRecSet.BOF) Then
'      wRsRmDet!qIrRmARt = 0#
'      wRsRmDet!qIrRmIRt = 0#
'      wRsRmDet!qIrSetSCd = ""
'      wRsRmDet!qIrSetARt = 0#
'      If UCase(ms_IdTcTyp) = UCase(ctInTcTyp) Then
'        wRsRmDet!qIrRmZRt = 0#
'        wRsRmDet!qIrSetIRt = 0#
'        wRsRmDet!qIrSetZRt = 0#
'      End If
'      wRsRmDet!qIrMainMet = "N"
'    Else
'      wOrRecSet.MoveFirst
'      wRsRmDet!qIrSetSCd = wOrRecSet!OrSetSCd
'      wRsRmDet!qIrSetARt = wOrRecSet!OrSetCstRt
'      If UCase(ms_IdTcTyp) = UCase(ctInTcTyp) Then
'        wRsRmDet!qIrRmARt = wOrRecSet!OrCstRt
'        wRsRmDet!qIrRmIRt = wOrRecSet!OrSalRt
'        wRsRmDet!qIrRmZRt = wOrRecSet!OrSalRt
'        wRsRmDet!qIrSetIRt = wOrRecSet!OrSetSalRt
'        wRsRmDet!qIrSetZRt = wOrRecSet!OrSetSalRt
'      End If
'
'      If wOrRecSet!OrMainMet = "Y" And wMainMet = False Then
'        wRsRmDet!qIrMainMet = "Y": wMainMet = True
'      ElseIf wOrRecSet!OrMainMet = "Y" And wMainMet = True Then
'        wRsRmDet!qIrMainMet = "N"
'      ElseIf wOrRecSet!OrMainMet = "N" Then
'        wRsRmDet!qIrMainMet = "N"
'      End If
'    End If
'
'    wRsRmDet.Update
'    wRsRmDet.MoveNext
'  Loop
'
'  If UCase(ms_IdTcTyp) = UCase(ctInTcTyp) Then
'
'    '****** check if mainMet wt > other Rm wt Diff. ***/
'    Set wRsRmDet = moCn.OpenRes("Select qIrRmCd, sum(qIrRmIWt) as qIrRmWt, " + _
'                                " max(qIrMainMet) as qMainMet from " + wTmpTbl + _
'                                " Group By qIrRmCd Order By qIrRmCd")
'    If (wRsRmDet.RecCount > 0) Then wRsRmDet.MoveFirst
'    wTotWtDiff = 0: wMainMetWt = 0
'    Do While Not (wRsRmDet.EOF Or wRsRmDet.BOF)
'        wFrOrdYN = Trim(mIWtEqOrd)
'        wTotIrRmWt = wRsRmDet!qIrRmWt
'        wRmCtg = moCn.GetFldVal("Select RmCtg from RmMst where RmCd= '" + wRsRmDet!qIrRmCd + "' ")
'        wDivFactor = IIF(wRmCtg = "C" Or wRmCtg = "D", 5, 1)
'
'        If wFrOrdYN = "N" Then wFrOrdYN = IIF(InStr(1, wCtgStr, "," + wRmCtg + ",") <> 0, "Y", "N")
'
'        If wFrOrdYN = "Y" Then
'            wOrStr = "Select sum(OrWt*" + CStr(mIdQty) + ") as qOrRmWt from OrdRm " + _
'                    "where OrCoCd= '" + gs_CoCd + "' " + _
'                    " and OrTc = '" + GrdID.Value(GrdID.Row, "IdExpOdTc") + "' " + _
'                    " and OrYy= '" + GrdID.Value(GrdID.Row, "IdExpOdYy") + "' " + _
'                    " and OrChr= '" + GrdID.Value(GrdID.Row, "IdExpOdChr") + "' " + _
'                    " and OrNo = " + CStr(GrdID.Value(GrdID.Row, "IdExpOdNo")) + _
'                    " and OrSr= " + CStr(GrdID.Value(GrdID.Row, "IdExpOdSr")) + _
'                    " and OrRmCd= '" + wRsRmDet!qIrRmCd + "' "
'
'            Set wOrRecSet = moCn.OpenRes(wOrStr)
'            If Not (wOrRecSet.EOF Or wOrRecSet.BOF) Then
'                wOrRecSet.MoveFirst
'                wWtDiff = 0
'                'wTotIrRmWt = wRsRmDet!qIrRmWt
'                If (wTotIrRmWt <> wOrRecSet!qOrRmWt) Then
'                    wWtDiff = wOrRecSet!qOrRmWt - wTotIrRmWt
'                    If UCase(wRsRmDet!qMainMet) <> "Y" Then _
'                        wTotWtDiff = wTotWtDiff + MWLib.Div(wWtDiff, wDivFactor)
'                End If
'            End If
'        End If
'
'        'urmi <>, = 23/02/04
'        If UCase(wRsRmDet!qMainMet) = "Y" Then _
'            wMainMetWt = MWLib.Div(wTotIrRmWt, wDivFactor)
'      wRsRmDet.MoveNext
'    Loop
'
'
'    'urmi <>, < 23/02/04
'    If wMainMetWt < wTotWtDiff Then
'
'        DispMsg "MainMetal Wt is less than the other Wt Diff.", etError
'        Exit Sub
'    End If
'
'    Set wRsRmDet = Nothing
'    Set wOrRecSet = Nothing
'    '********************************************************
'
'    Set wRsRmDet = moCn.OpenRes("Select qIrRmCd, sum(qIrRmIWt) as qIrRmWt, max(qIrMainMet) as qMainMet from " + wTmpTbl + _
'                                " Group By qIrRmCd Order By qIrRmCd")
'    If Not (wRsRmDet.EOF Or wRsRmDet.BOF) Then wRsRmDet.MoveFirst
'    wTotWtDiff = 0
'    Do While Not (wRsRmDet.EOF Or wRsRmDet.BOF)
'        wFrOrdYN = Trim(mIWtEqOrd)
'        wTotIrRmWt = wRsRmDet!qIrRmWt
'        wRmCtg = moCn.GetFldVal("Select RmCtg from RmMst where RmCd= '" + wRsRmDet!qIrRmCd + "' ")
'        If wFrOrdYN = "N" Then wFrOrdYN = IIF(InStr(1, wCtgStr, "," + wRmCtg + ",") <> 0, "Y", "N")
'
'        If wFrOrdYN = "Y" Then
'            wOrStr = "Select sum(OrWt*" + CStr(mIdQty) + ") as qOrRmWt from OrdRm " + _
'                    "where OrCoCd= '" + gs_CoCd + "' " + _
'                    " and OrTc = '" + GrdID.Value(GrdID.Row, "IdExpOdTc") + "' " + _
'                    " and OrYy= '" + GrdID.Value(GrdID.Row, "IdExpOdYy") + "' " + _
'                    " and OrChr= '" + GrdID.Value(GrdID.Row, "IdExpOdChr") + "' " + _
'                    " and OrNo = " + CStr(GrdID.Value(GrdID.Row, "IdExpOdNo")) + _
'                    " and OrSr= " + CStr(GrdID.Value(GrdID.Row, "IdExpOdSr")) + _
'                    " and OrRmCd= '" + wRsRmDet!qIrRmCd + "' "
'
'            Set wOrRecSet = moCn.OpenRes(wOrStr)
'            If Not (wOrRecSet.EOF Or wOrRecSet.BOF) Then
'                wOrRecSet.MoveFirst
'                wWtDiff = 0
'                'wTotIrRmWt = wRsRmDet!qIrRmWt
'
'                If (wTotIrRmWt <> wOrRecSet!qOrRmWt) Then
'                    wWtDiff = wOrRecSet!qOrRmWt - wTotIrRmWt
'                    wDivFactor = IIF(wRmCtg = "C" Or wRmCtg = "D", 5, 1)
'                    If UCase(wRsRmDet!qMainMet) = "Y" Then
'                        wTotWtDiff = wTotWtDiff + MWLib.Div(wWtDiff, wDivFactor)
'                    End If
'
'                    '###
'                    Set wIrRmWt = moCn.OpenResultset("Select * from " + wTmpTbl + _
'                                 " Where qIrRmCd = '" + wRsRmDet!qIrRmCd + "' ")
'
'                    If Not (wIrRmWt.EOF Or wIrRmWt.BOF) Then wIrRmWt.MoveFirst
'                    Do While Not (wIrRmWt.EOF Or wIrRmWt.BOF)
'                        wAddWt = (MWLib.Div(wWtDiff, wTotIrRmWt)) * wIrRmWt!qIrRmIWt
'                        wTotIrRmWt = wTotIrRmWt - wIrRmWt!qIrRmIWt
'                        wIrRmWt!qIrRmIWt = wIrRmWt!qIrRmIWt + wAddWt
'                        wIrRmWt.Update
'                        wWtDiff = wWtDiff - wAddWt
'                        wIrRmWt.MoveNext
'                    Loop
'                End If
'            End If
'
'        End If
'        wRsRmDet.MoveNext
'    Loop
'
'
'    Set wRsRmDet = moCn.OpenResultset("Select * from " + wTmpTbl + _
'                                      " where qIrMainMet= 'Y'")
'    If Not (wRsRmDet.EOF Or wRsRmDet.BOF) Then
'        wRsRmDet.MoveFirst
'        wRsRmDet!qIrRmIWt = wRsRmDet!qIrRmIWt - wTotWtDiff
'        wRsRmDet.Update
'    End If
'
'  End If
'
'
'  If UCase(ms_IdTcTyp) = UCase(ctInTcTyp) Then
'    moCn.Execute ("Update " + wTmpTbl + " Set qXtraIWt= .000, qXtraZWt= .000 ")
'
'
''********* URMI ****** 23/02/04 (commented code) *********
'    wDFrOrdYN = IIF(Trim(mIWtEqOrd) = "Y", "Y", "N")
'    If wDFrOrdYN = "N" Then wDFrOrdYN = IIF(InStr(1, wCtgStr, "," + "D" + ",") <> 0, "Y", "N")
'
'    wCFrOrdYN = IIF(Trim(mIWtEqOrd) = "Y", "Y", "N")
'    If wCFrOrdYN = "N" Then wCFrOrdYN = IIF(InStr(1, wCtgStr, "," + "C" + ",") <> 0, "Y", "N")
'
'    wDRrTcTyp = IIF(wDFrOrdYN = "Y", "XIO", "XIA")
'    wCRrTcTyp = IIF(wCFrOrdYN = "Y", "XIO", "XIA")
'
'    Call GetXtraWt(wDRrTcTyp, wCRrTcTyp, wTmpTbl, wCmDiaRtFA, wCmCsRtFA, GrdID.Value(GrdID.Row, "IdExpCmCd"), wDmCtg, _
'                   "qIrRmIWt", "qXtraIWt", "qIrRmIRt", mIdQty)
'    If whZWtEq = "A" Then
'      Call GetXtraWt("XZA", "XZA", wTmpTbl, wCmDiaRtFA, wCmCsRtFA, GrdID.Value(GrdID.Row, "IdExpCmCd"), wDmCtg, _
'                     "qIrRmZWt", "qXtraZWt", "qIrRmZRt", mIdQty)
'    End If
'
'
'    '*********************************************************
'
'  End If
'
'  Set wRsRmDet = moCn.OpenRes("select * from " + wTmpTbl + " Order By qIrMainMet")
'
'  wTotIXtraWt = 0: wTotZXtraWt = 0
'  wTotAWt = 0: wTotIWt = 0: wTotZWt = 0
'
'  With GrdIR
'    .StartCopy
'    Do While Not (wRsRmDet.EOF Or wRsRmDet.BOF)
'      .AddItem
'      wRmRow = .Rows - 1
'      If UCase(ms_IdTcTyp) = UCase(ctInTcTyp) Then
'        wDivFactor = IIF(wRsRmDet!qRmCtg = "C" Or wRsRmDet!qRmCtg = "D", 5, 1)
'        wTotIXtraWt = wTotIXtraWt + MWLib.Div(wRsRmDet!qXtraIWt, wDivFactor)
'        wTotZXtraWt = wTotZXtraWt + MWLib.Div(wRsRmDet!qXtraZWt, wDivFactor)
'
'        .Value(wRmRow, "IrSrNo") = wRmRow
'        .Value(wRmRow, "IrRmCd") = wRsRmDet!qIrRmCd
'        .Value(wRmRow, "IrRmSz") = wRsRmDet!qIrRmSz
'        .Value(wRmRow, "IrStkRt") = wRsRmDet!qIrStkRt
'        .Value(wRmRow, "IrRmQty") = wRsRmDet!qIrRmQty
'        .Value(wRmRow, "IrRmAWt") = wRsRmDet!qIrRmAWt
'        .Value(wRmRow, "IrRmARt") = wRsRmDet!qIrRmARt
'        .Value(wRmRow, "IrSetSCd") = wRsRmDet!qIrSetSCd
'        .Value(wRmRow, "IrSetARt") = wRsRmDet!qIrSetARt
'        .Value(wRmRow, "IrSetIRt") = wRsRmDet!qIrSetIRt
'        .Value(wRmRow, "IrMainMet") = wRsRmDet!qIrMainMet
'        .Value(wRmRow, "IrSetIRt") = wRsRmDet!qIrSetIRt
'        .Value(wRmRow, "IrSetZRt") = wRsRmDet!qIrSetZRt
'
'        If wRsRmDet!qIrMainMet = "Y" Then
'          'Geeta****07/06/05
'          .Value(wRmRow, "IrAlySalRt") = moCn.GetFldVal("Select  OrAlySalRt from OrdRm where OrCoCd='" + gs_CoCd + "' and OrTc= '" + GrdID.Value(GrdID.Row, "IdExpOdTc") + "' and " + _
'                                                         "OrYy= '" + GrdID.Value(GrdID.Row, "IdExpOdYy") + "' and OrChr= '" + GrdID.Value(GrdID.Row, "IdExpOdChr") + "' and " + _
'                                                         "OrNo= " + CStr(GrdID.Value(GrdID.Row, "IdExpOdNo")) + " and OrSr= " + CStr(GrdID.Value(GrdID.Row, "IdExpOdSr")) + " and  OrMainMet='Y' ")
'          'Geeta***07/06/05
'             .Value(wRmRow, "IrRmIWt") = wRsRmDet!qIrRmIWt - (wTotIWt - wTotAWt)
'          If whZWtEq = "A" Then
'            .Value(wRmRow, "IrRmZWt") = wRsRmDet!qIrRmZWt - (wTotZWt - wTotAWt)
'          ElseIf whZWtEq = "I" Then
'            .Value(wRmRow, "IrRmZWt") = wRsRmDet!qIrRmIWt - (wTotIWt - wTotAWt)
'          End If
'        Else
'          ' Manoj 2.10.0
'          '.Value(wRmRow, "IrRmZWt") = wRsRmDet!qIrRmiWt + wRsRmDet!qXtraZWt
'          .Value(wRmRow, "IrRmIWt") = wRsRmDet!qIrRmIWt + _
'                IIF(UCase(wRsRmDet!qRmCtg) = "C" Or UCase(wRsRmDet!qRmCtg) = "D", Round(wRsRmDet!qXtraIWt, 2), wRsRmDet!qXtraIWt)
'
'          If whZWtEq = "A" Then
'            '.Value(wRmRow, "IrRmZWt") = wRsRmDet!qIrRmZWt + wRsRmDet!qXtraZWt
'            .Value(wRmRow, "IrRmZWt") = wRsRmDet!qIrRmZWt + _
'                  IIF(UCase(wRsRmDet!qRmCtg) = "C" Or UCase(wRsRmDet!qRmCtg) = "D", Round(wRsRmDet!qXtraZWt, 2), wRsRmDet!qXtraZWt)
'          ElseIf whZWtEq = "I" Then
'            '.Value(wRmRow, "IrRmZWt") = wRsRmDet!qIrRmIWt + wRsRmDet!qXtraIWt
'            .Value(wRmRow, "IrRmZWt") = wRsRmDet!qIrRmIWt + _
'                  IIF(UCase(wRsRmDet!qRmCtg) = "C" Or UCase(wRsRmDet!qRmCtg) = "D", Round(wRsRmDet!qXtraIWt, 2), wRsRmDet!qXtraIWt)
'          End If
'          ' Manoj 2.10.0
'        End If
'
'        wTotAWt = wTotAWt + MWLib.Div(.Value(wRmRow, "IrRmAWt"), wDivFactor)
'        wTotIWt = wTotIWt + MWLib.Div(.Value(wRmRow, "IrRmIWt"), wDivFactor)
'        wTotZWt = wTotZWt + MWLib.Div(.Value(wRmRow, "IrRmZWt"), wDivFactor)
'
'        If UCase(wRsRmDet!qRmCtg) = "G" And mIGldRtEqOrd <> "Y" Then
''Geeta***07/06/05 (Before 2.11)
''          .Value(wRmRow, "IrRmIRt") = GetGldRt(GrdID.Value(GrdID.Row, "IdExpCmCd"), ADC("wIdLmgSal"), _
'                                      wRsRmDet!qIrRmCd, 0, "DM", GrdID.Value(GrdID.Row, "IdDmCd"), "S", "C")
'         .Value(wRmRow, "IrRmIRt") = GetGldRt(GrdID.Value(GrdID.Row, "IdExpCmCd"), ADC("wIdLmgSal"), _
'                                      wRsRmDet!qIrRmCd, .Value(wRmRow, "IrAlySalRt"), "DM", GrdID.Value(GrdID.Row, "IdDmCd"), "S", "C")
''Geeta***07/06/05 (Before 2.11)
'          If mLocalExp = "N" Then .Value(wRmRow, "IrRmZRt") = .Value(wRmRow, "IrRmIRt")
'        ElseIf UCase(wRsRmDet!qRmCtg) = "P" And mIGldRtEqOrd <> "Y" Then
''Geeta***07/06/05 (Before 2.11)
''          .Value(wRmRow, "IrRmIRt") = GetGldRt(GrdID.Value(GrdID.Row, "IdExpCmCd"), ADC("wIdLmpSal"), _
'                                      wRsRmDet!qIrRmCd, 0, "DM", GrdID.Value(GrdID.Row, "IdDmCd"), "S", "C")
'          .Value(wRmRow, "IrRmIRt") = GetGldRt(GrdID.Value(GrdID.Row, "IdExpCmCd"), ADC("wIdLmpSal"), _
'                                      wRsRmDet!qIrRmCd, .Value(wRmRow, "IrAlySalRt"), "DM", GrdID.Value(GrdID.Row, "IdDmCd"), "S", "C")
''Geeta***07/06/05 (Before 2.11)
'          If mLocalExp = "N" Then .Value(wRmRow, "IrRmZRt") = .Value(wRmRow, "IrRmIRt")
'        Else
'          .Value(wRmRow, "IrRmIRt") = wRsRmDet!qIrRmIRt
'          If mLocalExp = "N" Or (mLocalExp = "Y" And UCase(wRsRmDet!qRmCtg) <> "G" And UCase(wRsRmDet!qRmCtg) <> "P") Then
'            .Value(wRmRow, "IrRmZRt") = wRsRmDet!qIrRmZRt
'          End If
'        End If
'        .Value(wRmRow, "IrRmZPureRt") = IIF(wRsRmDet!qRmPurityZ > 0, MWLib.Div(wRsRmDet!qIrRmZRt, wRsRmDet!qRmPurityZ), wRsRmDet!qIrRmZRt)
'      ElseIf UCase(ms_IdTcTyp) = UCase(ctCTBTcTyp) Then
'        .Value(wRmRow, "IrSrNo") = wRmRow
'        .Value(wRmRow, "IrRmCd") = wRsRmDet!qIrRmCd
'        .Value(wRmRow, "IrRmSz") = wRsRmDet!qIrRmSz
'        .Value(wRmRow, "IrStkRt") = wRsRmDet!qIrStkRt
'        .Value(wRmRow, "IrRmQty") = wRsRmDet!qIrRmQty
'        .Value(wRmRow, "IrRmAWt") = wRsRmDet!qIrRmAWt
'
'        .Value(wRmRow, "IrRmIWt") = wRsRmDet!qIrRmIWt
'
'        .Value(wRmRow, "IrRmZWt") = wRsRmDet!qIrRmZWt
'        '.Value(wRmRow, "IrRmARt") = wRsRmDet!qIrRmARt
'        '.Value(wRmRow, "IrRmIRt") = wRsRmDet!qIrRmIRt
'        '.Value(wRmRow, "IrRmZRt") = wRsRmDet!qIrRmZRt
'        .Value(wRmRow, "IrSetSCd") = wRsRmDet!qIrSetSCd
'        .Value(wRmRow, "IrSetARt") = wRsRmDet!qIrSetARt
'        '.Value(wRmRow, "IrSetIRt") = wRsRmDet!qIrSetIRt
'        '.Value(wRmRow, "IrSetZRt") = wRsRmDet!qIrSetZRt
'        .Value(wRmRow, "IrRmZPureRt") = 0
'        .Value(wRmRow, "IrMainMet") = wRsRmDet!qIrMainMet
'      End If
'
'      ' Manoj 2.10.0
''      If UCase(wRsRmDet!qRmCtg) = "C" Or UCase(wRsRmDet!qRmCtg) = "D" Then
''        .Value(wRmRow, "IrRmAWt") = Round(.Value(wRmRow, "IrRmAWt"), 2)
''        .Value(wRmRow, "IrRmIWt") = Round(.Value(wRmRow, "IrRmIWt"), 2)
''        .Value(wRmRow, "IrRmZWt") = Round(.Value(wRmRow, "IrRmZWt"), 2)
''      End If
'      ' Manoj 2.10.0
'
'      .SaveRec (wRmRow)
'      wRsRmDet.MoveNext
'    Loop
'    .EndCopy
'  End With
'  '******Geeta***Emr208*************
'  '*********Zadd is Calculated in Make Lab ****
''  If UCase(ms_IdTcTyp) = UCase(ctInTcTyp) Then
''    wf_CnvRt = 1
''  ElseIf UCase(ms_IdTcTyp) = UCase(ctCTBTcTyp) Then
''    wf_CnvRt = moCn.GetFldVal("Select InCnvRt from InvHd where " + _
''                              "InCoCd= '" + gs_CoCd + "' and InTc= '" + adc("wIdTc") + "' " + _
''                              "and InYy= '" + adc("wIdYy") + "' and InChr= '" + adc("wIdChr") + "' " + _
''                              "and InNo= " + CStr(adc("wIdNo")))
''  End If
''
''  wf_AddLab = moCn.GetFldVal("Select sum(IrRmIVal- IrRmZVal* " + CStr(wf_CnvRt) + ") from InvRm " + _
'              "join RmMst on RmCd= IrRmCd where IrCoCd= '" + gs_CoCd + "' and " + _
'              "IrTc= '" + adc("wIdTc") + "' and IrYy= '" + adc("wIdYy") + "' and " + _
'              "IrChr= '" + adc("wIdChr") + "' and IrNo= " + CStr(adc("wIdNo")) + " and " + _
'              "IrSr= " + CStr(GrdID.Value(GrdID.Row, "IdSr")) + " and RmCtg in ('G', 'P')")
''
''  If mLocalExp = "Y" And Round(MWLib.Div(wf_AddLab, wf_CnvRt), 3) <> 0 Then
''    'And ms_IdTcTyp = ctInTcTyp
''    With GrdIL
''      .StartCopy
''      .AddItem
''      wLabRow = .Rows - 1
''      .Value(wLabRow, "iLSrNo") = wLabRow
''      .Value(wLabRow, "iLMCd") = ctAddLMCd
''      .Value(wLabRow, "iLSCd") = ctAddLSCd
''      .Value(wLabRow, "iLQty") = 1
''      .Value(wLabRow, "iLARt") = 0
''      .Value(wLabRow, "iLiRt") = 0
''      .Value(wLabRow, "iLZRt") = MWLib.Div(wf_AddLab, wf_CnvRt)
''      .Value(wLabRow, "ilQw") = "Q"
''      .Value(wLabRow, "iLZVal") = MWLib.Div(wf_AddLab, wf_CnvRt)
''      .SaveRec (wLabRow)
''      .EndCopy
''    End With
''  End If
''
'  Set wRsRmDet = Nothing
'  Set wOrRecSet = Nothing
  '*** (Bef 2.11)
End Sub
Public Sub MakeRmFromFg3(ByVal mIWtEqOrd As String, ByVal mIWtFrOrd As String, _
                        ByVal mIGldRtEqOrd As String, ByVal mIdQty As Double)
'  '*** (Bef 18/01/2006 Ver 2.12 without the change for WtFrOrd)
'  Dim whZWtEq As String, wCtgStr As String, wDmCtg As String, wTmpTbl As String
'  Dim wOrStr As String, wStr As String, wMainMet As Boolean
'  Dim wRsRmDet As MDORowSet, wOrRecSet As MDORowSet
'  Dim wDivFactor As Single, wRmRow As Single, wLabRow As Single, wf_AddLab As Single
'  Dim wTotIXtraWt As Double, wTotZXtraWt As Double, wTotAWt As Double
'  Dim wTotIWt As Double, wTotZWt As Double
'  Dim wCmDiaRtFA As String, wCmCsRtFA As String, wf_CnvRt As Single
'  Dim wDRrTcTyp As String, wCRrTcTyp As String, wDFrOrdYN As String
'  Dim wCFrOrdYN As String
'  Dim wSetRt As Boolean
'  Dim wTotWtDiff As Double, wWtDiff As Double, wTotIrRmWt As Double, wAddWt As Double
'  Dim wIrRmWt As MDORowSet, wRmCtg As String, wFrOrdYN As String, wMainMetWt As Double
'  Dim wFrRmWt As Double, wFrRmQty As Double, wOrTotWt As Double, wOrTotQty As Double
'  Dim wFRmDet As MDORowSet
'  whZWtEq = moCn.GetFldVal("Select hZWtEq from Head where HCoCd= '" + gs_CoCd + "' and HCd='" + ctSelfCmCd + "'")
'  Dim wSetTyp As Boolean
'  Dim wOrLnStr As String
'
'  '*** (Jen 2.12)
'  Dim wRs_RmCtg As MDORowSet
'  '*** (Jen 2.12)
'
'  'mIWtFrOrd = "DC"
'  'wCtgStr= ",D,C"
'  wCtgStr = "," + moCn.GetFldVal("Select PDesc225 from Param where " + _
'              "PTyp='IWTFRORD' and PMCd = '" + mIWtFrOrd + "'") + ","
'
'  wCmDiaRtFA = moCn.GetFldVal("Select CmDiaRtFA from CustMst where CmCtg= 'C' and " + _
'               "CmCd= '" + GrdID.Value(GrdID.Row, "IdExpCmCd") + "' ")
'
'  wCmCsRtFA = moCn.GetFldVal("Select CmCSRtFA from CustMst where CmCtg= 'C' and " + _
'               "CmCd= '" + GrdID.Value(GrdID.Row, "IdExpCmCd") + "' ")
'
'
'  wDmCtg = moCn.GetFldVal("Select DmCtg from DsgMst where DmTcTyp= 'DM' and DmSz= '' and " + _
'               "DmCd= '" + GrdID.Value(GrdID.Row, "IdDmCd") + "' ")
'  wTmpTbl = "#TmpInvRm"
'
'  '*******Geeta*******Emr207
'  Dim ws_OrLn As String
'  ws_OrLn = "(case when Pvalue3='Y' then (Case when PValue2='Y' then  " + _
'          "cast(OrRmPtr as Decimal(12,4)) else (case when OrPrdQty=0 then 0 else  " + _
'          "cast((OrPrdWt/OrPrdQty) as Decimal(12,4)) end ) end )  " + _
'          "else cast(Orln1  as Decimal(12,4)) end ) "
'  If UCase(ms_IdTcTyp) = UCase(ctInTcTyp) Then
'    ' Zubin 212 (Introduced Lot No)
'    wStr = "Select FrRmCd as qIrRmCd, FrLotNo as qIrLotNo, FrRmSz as qIrRmSz, sum(FrRmStkRt*FrRmWt)/(case when sum(FrRmWt)> 0  " + _
'           "then sum(FrRmWt) else 1 end) as qIrStkRt, " + _
'           "sum(FrRmQty) as qIrRmQty, sum(FrRmWt) As qIrRmAWt, sum(FrRmWt) as qIrRmIWt, " + _
'           "sum(FrRmWt) as qIrRmZWt,max(Isnull(qOrWt,0)) as qOrWt,max(IsNull(qOrQty,0)) as qOrQty, 9999.999 as qXtraIWt, 9999.999 as qXtraZWt, " + _
'           "999999.999 as qIrRmARt, 999999.999 as qIrRmIRt, 999999.999 as qIrRmZRt, IsNull(qOrSetSCd, '') as qIrSetSCd, " + _
'           "isnull(qOrSetSalRt, 0.00) as qIrSetARt, isnull(qOrSetSalRt, 0.00) as qIrSetIRt, isnull(qOrSetSalRt, 0.00) as qIrSetZRt, " + _
'           "Space(1) as qIrMainMet, max(RmPurityZ) as qRmPurityZ, max(RmCtg) as qRmCtg, max(RmSCtg) as qRmSCtg " + _
'           "into " + wTmpTbl + " from InvFgd join FgRm on FrCoCd= IfCoCd and FrTc= IfFdTc and FrYy= IfFdyy and " + _
'           "FrChr= IfFdChr and FrNo = IfFdNo And FrSr = IfFdSr " + _
'           "Left Outer Join (select OrRmCd as qOrRmCd, " + ws_OrLn + " as qOrLn1, OrSetSCd as qOrSetSCd, " + _
'           "OrSetSalRt as qOrSetSalRt, sum(OrPrdWt) as qOrWt, sum(OrPrdQty) as qOrQty  from OrdRm  " + _
'           "Join RmMst on RmCd= OrRmCd " + _
'           "Join Param on PTyp= 'RMSCTG' and PMCd= RmCtg and PSCd= RmSCtg " + _
'           "where OrCoCd= '" + gs_CoCd + "'  and OrTc = '" + GrdID.Value(GrdID.Row, "IdExpOdTc") + "' " + _
'           "and OrYy= '" + GrdID.Value(GrdID.Row, "IdExpOdYy") + "' " + _
'           " and OrChr= '" + GrdID.Value(GrdID.Row, "IdExpOdChr") + "' and OrNo = " + CStr(GrdID.Value(GrdID.Row, "IdExpOdNo")) + _
'           " and OrSr= " + CStr(GrdID.Value(GrdID.Row, "IdExpOdSr")) + " Group By OrCoCd, OrTc, OrYy, OrChr, OrNo, OrSr, " + _
'           "OrRmCd, " + ws_OrLn + ", OrSetSCd, OrSetSalRt) Or2 On qOrRmCd= FrRmCd and qOrLn1= FrRmSz " + _
'           "Join RmMst on FrRmCd = RmCd " + _
'           "where IfCoCd = '" + gs_CoCd + "' and IfTc= '" + ADC("wIdTc") + "' and IfYy= '" + _
'           ADC("wIdYy") + "' and IfChr= '" + ADC("wIdChr") + "' and " + _
'           "IfNo = " + CStr(ADC("wIdNo")) + " and IfSr = " + CStr(GrdID.Value(GrdID.Row, "IdSr")) + " " + _
'           "Group by FrRmCd, FrLotNo, FrRmSz, qOrSetSCd, qOrSetSalRt "
'  ElseIf UCase(ms_IdTcTyp) = UCase(ctCTBTcTyp) Then
'    ' Zubin 212 (Introduced Lot No)
'    wStr = "Select (case when isnull(qOrRmCd, '')= '' then '' else 'Y' end) as qOrRmExists, " + _
'           "FrRmCd as qIrRmCd, FrLotNo as qIrLotNo, FrRmSz as qIrRmSz,sum(FrRmStkRt*FrRmWt)/(case when sum(FrRmWt)> 0  " + _
'           "then sum(FrRmWt) else 1 end) as qIrStkRt, sum(FrRmQty) as qIrRmQty, " + _
'           "sum(FrRmWt) As qIrRmAWt, sum(FrRmWt) as qIrRmIWt, sum(FrRmWt) as qIrRmZWt, " + _
'           "max(Isnull(qOrWt,0)) as qOrWt,max(IsNull(qOrQty,0)) as qOrQty, 0 as qXtraIWt, 0 as qXtraZWt, " + _
'           "0 as qIrRmARt, 0 as qIrRmIRt, 0 as qIrRmZRt, IsNull(qOrSetSCd, '') as qIrSetSCd, " + _
'           "isnull(qOrSetSalRt, 0.00) as qIrSetARt, isnull(qOrSetSalRt, 0.00) as qIrSetIRt, 0 as qIrSetZRt, " + _
'           "Space(1) as qIrMainMet, max(RmPurityZ) as qRmPurityZ, max(RmCtg) as qRmCtg, max(RmSCtg) as qRmSCtg " + _
'           "into " + wTmpTbl + " from InvFgd " + _
'           "join FgRm on FrCoCd= IfCoCd and FrTc= IfFdTc and FrYy= IfFdYy and FrChr= IfFdChr and FrNo = IfFdNo And FrSr = IfFdSr " + _
'           "Left Outer Join (select OrRmcd as qOrRmcd ," + ws_OrLn + " as qOrLn1, OrSetSCd as qOrSetSCd, " + _
'           "OrSetSalRt as qOrSetSalRt, sum(OrPrdWt) as qOrWt, sum(OrPrdQty) as qOrQty from OrdRm " + _
'           "Join RmMst on RmCd= OrRmCd " + _
'           "Join Param on PTyp= 'RMSCTG' and PMCd= RmCtg and PSCd= RmSCtg " + _
'           "where OrCoCd= '" + gs_CoCd + "' and OrTc= '" + GrdID.Value(GrdID.Row, "IdExpOdTc") + "' " + _
'           "and OrYy= '" + GrdID.Value(GrdID.Row, "IdExpOdYy") + "' " + _
'           " and OrChr= '" + GrdID.Value(GrdID.Row, "IdExpOdChr") + "' and OrNo = " + CStr(GrdID.Value(GrdID.Row, "IdExpOdNo")) + _
'           " and OrSr= " + CStr(GrdID.Value(GrdID.Row, "IdExpOdSr")) + " group by OrCoCd, OrTc, OrYy, OrChr, OrNo, OrSr, " + _
'           "OrRmCd, " + ws_OrLn + ", OrSetSCd, OrSetSalRt) Or2 On qOrRmcd= FrRmCd and qOrln1= FrRmSz " + _
'           "Join RmMst on FrRmCd= RmCd " + _
'           "where IfCoCd = '" + gs_CoCd + "' and IfTc= '" + ADC("wIdTc") + "' " + _
'           "and IfYy= '" + ADC("wIdYy") + "' and IfChr= '" + ADC("wIdChr") + "' and " + _
'           "IfNo = " + CStr(ADC("wIdNo")) + " and IfSr = " + CStr(GrdID.Value(GrdID.Row, "IdSr")) + " " + _
'           "Group by FrRmCd, FrLotNo, FrRmSz, qOrSetSCd, qOrSetSalRt "
'  End If
'  '**********Geeta******Emr207
'
'  '*** Creating Temporary Table #TmpInvRm ***
'  moCn.CreateTmpTable wTmpTbl, wStr
'  'moCn.Execute (wStr)
'
'  Set wRsRmDet = moCn.OpenResultset("Select * from " + wTmpTbl)
'
'  '**********Geeta******Emr207
'  '*** (Bef 2.11)
'  'If moCn.RecSeek("select count(distinct qIrSetIRt) From " + wTmpTbl + " " + _
'  '                  "group by qIrRmCd,qIrRmSz Having count(distinct qIrSetIRt)>1") Then
'  '*** (Bef 2.11)
'
'  '*** 2.11 (Jen)
'  ' Zubin 212
'  If moCn.RecSeek("select count(distinct Str(qIrSetIRt, 16, 2) + qIrSetSCd) From " + wTmpTbl + " " + _
'                  "group by qIrRmCd, qIrLotNo, qIrRmSz Having count(distinct Str(qIrSetIRt, 16, 2) + qIrSetSCd)> 1") Then
'  '*** 2.11 (Jen)
'
'    '  To get Proportionate weights
'    ' Zubin 212
'    Set wFRmDet = moCn.OpenRes("Select FrRmCd as qFrRmCd, FrLotNo as qFrLotNo, FrRmsz as qFrRmsz, " + _
'                       " Sum(FrRmWt) as qFrRmWt, Sum(FrRmQty) as qFrRmQty from FgRm  " + _
'                       "join InvFgd on FrCoCd= IfCoCd and FrTc= IfFdTc and FrYy= IfFdyy and  " + _
'                       " FrChr= IfFdChr and FrNo = IfFdNo And FrSr = IfFdSr " + _
'                       "where IfCoCd = '" + gs_CoCd + "' and IfTc= '" + ADC("wIdTc") + "' and IfYy= '" + _
'                       ADC("wIdYy") + "' and IfChr= '" + ADC("wIdChr") + "' and " + _
'                       "IfNo = " + CStr(ADC("wIdNo")) + " and IfSr = " + CStr(GrdID.Value(GrdID.Row, "IdSr")) + " " + _
'                       "Group By FrRmCd, FrLotNo, FrRmSz Order By FrRmCd, FrLotNo, FrRmSz")
'
'    Do While Not (wFRmDet.EOF Or wFRmDet.BOF)
'      wFrRmWt = wFRmDet!qFrRmWt
'      wFrRmQty = wFRmDet!qFrRmQty
'      ' Zubin 212
'      wOrTotWt = moCn.GetFldVal("Select Sum(qOrWt) From " + wTmpTbl + " Where qIrRmCd= '" + wFRmDet!qFrRmCd + "' and qIrRmSz= '" + CStr(wFRmDet!qFrRmSz) + "' and qIrLotNo= '" + CStr(wFRmDet!qFrLotNo) + "' Group By qIrRmcd, qIrLotNo, qIrRmSz")
'      wOrTotQty = moCn.GetFldVal("Select Sum(qOrQty) From " + wTmpTbl + "  Where qIrRmCd= '" + wFRmDet!qFrRmCd + "' and qIrRmSz= '" + CStr(wFRmDet!qFrRmSz) + "' and qIrLotNo= '" + CStr(wFRmDet!qFrLotNo) + "' Group By qIrRmCd, qIrLotNo, qIrRmSz")
'      ' Zubin 212
'      If (wRsRmDet.RecCount > 0) Then wRsRmDet.MoveFirst
'      Do While Not (wRsRmDet.EOF Or wRsRmDet.BOF)
'        ' Zubin 212
'        If wFRmDet!qFrRmCd = wRsRmDet!qIrRmCd And wFRmDet!qFrLotNo = wRsRmDet!qIrLotNo And wFRmDet!qFrRmSz = wRsRmDet!qIrRmSz Then
'          '*** (Bef 2.11)
'          'If moCn.RecSeek("select count(distinct qIrSetIRt) From " + wTmpTbl + " " + _
'          '                "Where qIrRmCd= '" + wRsRmDet!qIrRmCd + "' " + _
'          '                " and qIrRmSz = " + CStr(wRsRmDet!qIrRmSz) + _
'          '                "group by qIrRmCd,qIrRmSz Having count(distinct qIrSetIRt)>1") Then
'          '*** (Bef 2.11)
'          ' Zubin 212
'          If moCn.RecSeek("select count(distinct Str(qIrSetIRt, 16, 2) + qIrSetSCd) From " + wTmpTbl + " " + _
'                          "Where qIrRmCd= '" + wRsRmDet!qIrRmCd + "' " + _
'                          " and qIrLotNo= '" + CStr(wRsRmDet!qIrLotNo) + "' " + _
'                          " and qIrRmSz= " + CStr(wRsRmDet!qIrRmSz) + _
'                          "group by qIrRmCd, qIrLotNo, qIrRmSz Having count(distinct Str(qIrSetIRt, 16, 2) + qIrSetSCd)> 1") Then
'            wRsRmDet.edit
'            wRsRmDet!qIrRmIWt = Round(MWLib.Div(wRsRmDet!qOrWt, wOrTotWt) * wFrRmWt, 3)
'            wRsRmDet!qIrRmAWt = wRsRmDet!qIrRmIWt
'            wRsRmDet!qIrRmZWt = wRsRmDet!qIrRmIWt
'            wOrTotWt = wOrTotWt - wRsRmDet!qOrWt
'            wFrRmWt = wFrRmWt - wRsRmDet!qIrRmIWt
'            wRsRmDet!qIrRmQty = Round(MWLib.Div(wRsRmDet!qOrQty, wOrTotQty) * wFrRmQty)
'            wOrTotQty = wOrTotQty - wRsRmDet!qOrQty
'            wFrRmQty = wFrRmQty - wRsRmDet!qIrRmQty
'            wRsRmDet.Update
'          End If
'        End If
'        wRsRmDet.MoveNext
'      Loop
'      wFRmDet.MoveNext
'    Loop
'    Set wFRmDet = Nothing
'  End If
'
'  'If (wRsRmDet.RecCount > 0) Then wRsRmDet.MoveFirst
'  'Do While Not (wRsRmDet.EOF Or wRsRmDet.BOF)
'  '  Debug.Print CStr(wRsRmDet!qIrRmCd) + "," + CStr(wRsRmDet!qIrRmSz) + "," + CStr(wRsRmDet!qIrRmQty) + "," + CStr(wRsRmDet!qIrRmAWt) + "," + CStr(wRsRmDet!qOrWt) + "," + CStr(wRsRmDet!qOrQty)
'  '  wRsRmDet.MoveNext
'  'Loop
'
'  '**********Geeta******Emr207
'  If (wRsRmDet.RecCount > 0) Then wRsRmDet.MoveFirst
'  Do While Not (wRsRmDet.EOF Or wRsRmDet.BOF)
'    wRsRmDet.edit
'
'    wOrLnStr = IIF(moCn.GetFldVal("select PValue3 From Param Join  RmMst " + _
'                            " on PTyp='RMSCTG' and PMCd=RmCtg and PSCd=RmSCtg " + _
'                            "where RmCd='" + wRsRmDet!qIrRmCd + "' ") = "Y", _
'                   IIF(moCn.GetFldVal("select PValue2 From Param Join  RmMst " + _
'                            "on PTyp='RMSCTG' and PMCd=RmCtg and PSCd=RmSCtg " + _
'                            "where RmCd='" + wRsRmDet!qIrRmCd + "' ") = "Y", "OrRmPtr", _
'                            " case when OrPrdQty =0 then 0 else OrPrdWt/OrPrdQty end "), "OrLn1")
'
'    '*** (Bef 2.11)
'    'If moCn.RecSeek("select count(distinct qIrSetIRt) From " + wTmpTbl + " " + _
'                    "Where qIrRmCd= '" + wRsRmDet!qIrRmCd + "' " + _
'                    " and qIrRmSz= " + CStr(wRsRmDet!qIrRmSz) + _
'                    "group by qIrRmCd, qIrRmSz Having count(distinct qIrSetIRt)> 1") Then
'    '*** (Bef 2.11)
'
'    '*** 2.11 (Jen)
'    ' Zubin 212
'    If moCn.RecSeek("select count(distinct Str(qIrSetIRt, 16, 2) + qIrSetSCd) From " + wTmpTbl + " " + _
'                    "Where qIrRmCd= '" + wRsRmDet!qIrRmCd + "' " + _
'                    " and qIrLotNo= '" + CStr(wRsRmDet!qIrLotNo) + "' " + _
'                    " and qIrRmSz= " + CStr(wRsRmDet!qIrRmSz) + _
'                    "group by qIrRmCd, qIrLotNo, qIrRmSz Having count(distinct Str(qIrSetIRt, 16, 2) + qIrSetSCd)> 1") Then
'    '*** 2.11 (Jen)
'      '*******Geeta*******Emr207
'      wOrStr = "Select OrCstRt, OrSalRt, OrSetSCd, OrSetCstRt, OrSetSalRt, OrMainMet, OrRmCtg " + _
'              "From OrdRm where OrCoCd= '" + gs_CoCd + "' " + _
'              " and OrTc = '" + GrdID.Value(GrdID.Row, "IdExpOdTc") + "' " + _
'              " and OrYy= '" + GrdID.Value(GrdID.Row, "IdExpOdYy") + "' " + _
'              " and OrChr= '" + GrdID.Value(GrdID.Row, "IdExpOdChr") + "' " + _
'              " and OrNo = " + CStr(GrdID.Value(GrdID.Row, "IdExpOdNo")) + _
'              " and OrSr= " + CStr(GrdID.Value(GrdID.Row, "IdExpOdSr")) + _
'              " and OrRmCd= '" + wRsRmDet!qIrRmCd + "' " + _
'              " and " + wOrLnStr + " = " + CStr(wRsRmDet!qIrRmSz) + _
'              " and OrSetSCd= '" + wRsRmDet!qIrSetSCd + "' " + _
'              " and OrSetSalRt= " + CStr(wRsRmDet!qIrSetIRt) + _
'              " Order by OrMainMet Desc, OrSalRt Desc"
'
'      'Set wOrRecSet = moCn.OpenRes(wOrStr)
'    Else
'      wOrStr = "Select OrCstRt, OrSalRt, OrSetSCd, OrSetCstRt, OrSetSalRt, OrMainMet, OrRmCtg " + _
'                "From OrdRm where OrCoCd= '" + gs_CoCd + "' " + _
'                " and OrTc = '" + GrdID.Value(GrdID.Row, "IdExpOdTc") + "' " + _
'                " and OrYy= '" + GrdID.Value(GrdID.Row, "IdExpOdYy") + "' " + _
'                " and OrChr= '" + GrdID.Value(GrdID.Row, "IdExpOdChr") + "' " + _
'                " and OrNo = " + CStr(GrdID.Value(GrdID.Row, "IdExpOdNo")) + _
'                " and OrSr= " + CStr(GrdID.Value(GrdID.Row, "IdExpOdSr")) + _
'                " and OrRmCd= '" + wRsRmDet!qIrRmCd + "' " + _
'                " and " + wOrLnStr + " = " + CStr(wRsRmDet!qIrRmSz) + _
'                "Order by OrMainMet Desc, OrSalRt Desc"
'    End If
'    Set wOrRecSet = moCn.OpenRes(wOrStr)
'
'    '    Query to get weighted Average rate
'    '    wOrStr = "Select (Case when sum(OrWt)<>0 then sum(OrCstRt*OrWt)/sum(OrWt) else 0 end) as OrCstRt, " + _
'                  " (Case when sum(OrWt)<>0 then sum(OrSalRt*OrWt)/sum(OrWt) else 0 end) as OrSalRt, " + _
'                  " (select max(OrSetSCd) from OrdRm B where  B.OrCoCd= '" + gs_CoCd + "' " + _
'                  " and B.OrTc = '" + GrdID.Value(GrdID.Row, "IdExpOdTc") + "' " + _
'                  " and B.OrYy= '" + GrdID.Value(GrdID.Row, "IdExpOdYy") + "' " + _
'                  " and B.OrChr= '" + GrdID.Value(GrdID.Row, "IdExpOdChr") + "' " + _
'                  " and B.OrNo = " + CStr(GrdID.Value(GrdID.Row, "IdExpOdNo")) + _
'                  " and B.OrSr= " + CStr(GrdID.Value(GrdID.Row, "IdExpOdSr")) + _
'                  " and B.OrRmCd= '" + wRsRmDet!qIrRmCd + "' " + _
'                  " and B.OrLn1 = " + CStr(wRsRmDet!qIrRmSz) + _
'                  " and B.OrSetSalRt = max(A.OrSetSalRt)) as OrSetSCd, " + _
'                  " (Case when sum(OrQty)<>0 then sum(OrSetCstRt*OrQty)/sum(OrQty) else 0 end) as OrSetCstRt, " + _
'                  " (case when sum(OrQty)<>0 then sum(OrSetSalRt*OrQty)/sum(OrQty) else 0 end) as OrSetSalRt, " + _
'                  " max(OrMainMet), max(OrRmCtg) " + _
'                  " From OrdRm A where OrCoCd= '" + gs_CoCd + "' " + _
'                  " and OrTc = '" + GrdID.Value(GrdID.Row, "IdExpOdTc") + "' " + _
'                  " and OrYy= '" + GrdID.Value(GrdID.Row, "IdExpOdYy") + "' " + _
'                  " and OrChr= '" + GrdID.Value(GrdID.Row, "IdExpOdChr") + "' " + _
'                  " and OrNo = " + CStr(GrdID.Value(GrdID.Row, "IdExpOdNo")) + _
'                  " and OrSr= " + CStr(GrdID.Value(GrdID.Row, "IdExpOdSr")) + _
'                  " and OrRmCd= '" + wRsRmDet!qIrRmCd + "' " + _
'                  " and OrLn1 = " + CStr(wRsRmDet!qIrRmSz) + _
'                  " Group By OrRmCd, OrLn1 "
'
'    If (wOrRecSet.EOF Or wOrRecSet.BOF) Then
'      wOrStr = "Select OrCstRt, OrSalRt, OrSetSCd, OrSetCstRt, OrSetSalRt, OrMainMet, OrRmCtg " + _
'               "From OrdRm join RmRt on RrTcTyp= 'Rsz' and RrCmCd= '" + ctSelfCmCd + "' and " + _
'               "RrCtg= OrRmCtg and RrSCtg= OrRmSCtg " + _
'               "where OrCoCd= '" + gs_CoCd + "' " + _
'               " and OrTc = '" + GrdID.Value(GrdID.Row, "IdExpOdTc") + "' " + _
'               " and OrYy= '" + GrdID.Value(GrdID.Row, "IdExpOdYy") + "' " + _
'               " and OrChr = '" + GrdID.Value(GrdID.Row, "IdExpOdChr") + "' " + _
'               " and OrNo = " + CStr(GrdID.Value(GrdID.Row, "IdExpOdNo")) + _
'               " and OrSr= " + CStr(GrdID.Value(GrdID.Row, "IdExpOdSr")) + _
'               " and OrRmCd =  '" + wRsRmDet!qIrRmCd + "' " + _
'               " and " + wOrLnStr + " >= RrFrLn and " + wOrLnStr + " <= RrToLn and " + _
'               CStr(wRsRmDet!qIrRmSz) + " >= RrFrLn and " + _
'               CStr(wRsRmDet!qIrRmSz) + " <= RrToLn " + _
'               "Order by OrMainMet Desc, OrSalRt Desc "
'
'      Set wOrRecSet = moCn.OpenRes(wOrStr)
'    End If
'    If (wOrRecSet.EOF Or wOrRecSet.BOF) Then
'      wOrStr = "Select OrCstRt, OrSalRt, OrSetSCd, OrSetCstRt, OrSetSalRt, OrMainMet, OrRmCtg " + _
'               "From OrdRm where OrCoCd= '" + gs_CoCd + "' " + _
'               "  and OrTc = '" + GrdID.Value(GrdID.Row, "IdExpOdTc") + "' " + _
'               " and OrYy= '" + GrdID.Value(GrdID.Row, "IdExpOdYy") + "' " + _
'               " and OrChr= '" + GrdID.Value(GrdID.Row, "IdExpOdChr") + "' " + _
'               " and OrNo = " + CStr(GrdID.Value(GrdID.Row, "IdExpOdNo")) + _
'               " and OrSr= " + CStr(GrdID.Value(GrdID.Row, "IdExpOdSr")) + _
'               " and OrRmCd= '" + wRsRmDet!qIrRmCd + "' " + _
'               "Order by OrMainMet Desc, OrSalRt Desc"
'      Set wOrRecSet = moCn.OpenRes(wOrStr)
'    End If
'
'    If (wOrRecSet.EOF Or wOrRecSet.BOF) Then
'      wRsRmDet!qIrRmARt = 0#
'      wRsRmDet!qIrRmIRt = 0#
'      wRsRmDet!qIrSetSCd = ""
'      wRsRmDet!qIrSetARt = 0#
'      If UCase(ms_IdTcTyp) = UCase(ctInTcTyp) Then
'        wRsRmDet!qIrRmZRt = 0#
'        wRsRmDet!qIrSetIRt = 0#
'        wRsRmDet!qIrSetZRt = 0#
'      End If
'      wRsRmDet!qIrMainMet = "N"
'    Else
'      wOrRecSet.MoveFirst
'      wRsRmDet!qIrSetSCd = wOrRecSet!OrSetSCd
'      wRsRmDet!qIrSetARt = wOrRecSet!OrSetCstRt
'      If UCase(ms_IdTcTyp) = UCase(ctInTcTyp) Then
'        wRsRmDet!qIrRmARt = wOrRecSet!OrCstRt
'        wRsRmDet!qIrRmIRt = wOrRecSet!OrSalRt
'        wRsRmDet!qIrRmZRt = wOrRecSet!OrSalRt
'        wRsRmDet!qIrSetIRt = wOrRecSet!OrSetSalRt
'        wRsRmDet!qIrSetZRt = wOrRecSet!OrSetSalRt
'      End If
'
'      If wOrRecSet!OrMainMet = "Y" And wMainMet = False Then
'        wRsRmDet!qIrMainMet = "Y": wMainMet = True
'      ElseIf wOrRecSet!OrMainMet = "Y" And wMainMet = True Then
'        wRsRmDet!qIrMainMet = "N"
'      ElseIf wOrRecSet!OrMainMet = "N" Then
'        wRsRmDet!qIrMainMet = "N"
'      End If
'    End If
'
'    wRsRmDet.Update
'    wRsRmDet.MoveNext
'  Loop
'
'  If UCase(ms_IdTcTyp) = UCase(ctInTcTyp) Then
'
'    '****** check if mainMet wt > other Rm wt Diff. ***/
'    Set wRsRmDet = moCn.OpenRes("Select qIrRmCd, sum(qIrRmIWt) as qIrRmWt, " + _
'                                " max(qIrMainMet) as qMainMet from " + wTmpTbl + _
'                                " Group By qIrRmCd Order By qIrRmCd")
'    If (wRsRmDet.RecCount > 0) Then wRsRmDet.MoveFirst
'    wTotWtDiff = 0: wMainMetWt = 0
'    Do While Not (wRsRmDet.EOF Or wRsRmDet.BOF)
'      wFrOrdYN = Trim(mIWtEqOrd)
'      wTotIrRmWt = wRsRmDet!qIrRmWt
'      wRmCtg = moCn.GetFldVal("Select RmCtg from RmMst where RmCd= '" + wRsRmDet!qIrRmCd + "' ")
'      wDivFactor = IIF(wRmCtg = "C" Or wRmCtg = "D", 5, 1)
'
'      If wFrOrdYN = "N" Then wFrOrdYN = IIF(InStr(1, wCtgStr, "," + wRmCtg + ",") <> 0, "Y", "N")
'
'      If wFrOrdYN = "Y" Then
'        wOrStr = "Select sum(OrWt*" + CStr(mIdQty) + ") as qOrRmWt from OrdRm " + _
'                "where OrCoCd= '" + gs_CoCd + "' " + _
'                " and OrTc = '" + GrdID.Value(GrdID.Row, "IdExpOdTc") + "' " + _
'                " and OrYy= '" + GrdID.Value(GrdID.Row, "IdExpOdYy") + "' " + _
'                " and OrChr= '" + GrdID.Value(GrdID.Row, "IdExpOdChr") + "' " + _
'                " and OrNo = " + CStr(GrdID.Value(GrdID.Row, "IdExpOdNo")) + _
'                " and OrSr= " + CStr(GrdID.Value(GrdID.Row, "IdExpOdSr")) + _
'                " and OrRmCd= '" + wRsRmDet!qIrRmCd + "' "
'
'        Set wOrRecSet = moCn.OpenRes(wOrStr)
'        If Not (wOrRecSet.EOF Or wOrRecSet.BOF) Then
'          wOrRecSet.MoveFirst
'          wWtDiff = 0
'          'wTotIrRmWt = wRsRmDet!qIrRmWt
'          If (wTotIrRmWt <> wOrRecSet!qOrRmWt) Then
'            wWtDiff = wOrRecSet!qOrRmWt - wTotIrRmWt
'            If UCase(wRsRmDet!qMainMet) <> "Y" Then _
'               wTotWtDiff = wTotWtDiff + MWLib.Div(wWtDiff, wDivFactor)
'          End If
'        End If
'      End If
'
'      'urmi <>, = 23/02/04
'      If UCase(wRsRmDet!qMainMet) = "Y" Then _
'         wMainMetWt = MWLib.Div(wTotIrRmWt, wDivFactor)
'      wRsRmDet.MoveNext
'    Loop
'
'
'    'urmi <>, < 23/02/04
'    If wMainMetWt < wTotWtDiff Then
'      DispMsg "MainMetal Wt is less than the other Wt Diff.", etError
'      Exit Sub
'    End If
'
'    Set wRsRmDet = Nothing
'    Set wOrRecSet = Nothing
'    '********************************************************
'
'    Set wRsRmDet = moCn.OpenRes("Select qIrRmCd, sum(qIrRmIWt) as qIrRmWt, max(qIrMainMet) as qMainMet from " + wTmpTbl + _
'                                " Group By qIrRmCd Order By qIrRmCd")
'    If Not (wRsRmDet.EOF Or wRsRmDet.BOF) Then wRsRmDet.MoveFirst
'    wTotWtDiff = 0
'    Do While Not (wRsRmDet.EOF Or wRsRmDet.BOF)
'      wFrOrdYN = Trim(mIWtEqOrd)
'      wTotIrRmWt = wRsRmDet!qIrRmWt
'      wRmCtg = moCn.GetFldVal("Select RmCtg from RmMst where RmCd= '" + wRsRmDet!qIrRmCd + "' ")
'      If wFrOrdYN = "N" Then wFrOrdYN = IIF(InStr(1, wCtgStr, "," + wRmCtg + ",") <> 0, "Y", "N")
'
'      If wFrOrdYN = "Y" Then
'        wOrStr = "Select sum(OrWt*" + CStr(mIdQty) + ") as qOrRmWt from OrdRm " + _
'                "where OrCoCd= '" + gs_CoCd + "' " + _
'                " and OrTc = '" + GrdID.Value(GrdID.Row, "IdExpOdTc") + "' " + _
'                " and OrYy= '" + GrdID.Value(GrdID.Row, "IdExpOdYy") + "' " + _
'                " and OrChr= '" + GrdID.Value(GrdID.Row, "IdExpOdChr") + "' " + _
'                " and OrNo = " + CStr(GrdID.Value(GrdID.Row, "IdExpOdNo")) + _
'                " and OrSr= " + CStr(GrdID.Value(GrdID.Row, "IdExpOdSr")) + _
'                " and OrRmCd= '" + wRsRmDet!qIrRmCd + "' "
'
'        Set wOrRecSet = moCn.OpenRes(wOrStr)
'        If Not (wOrRecSet.EOF Or wOrRecSet.BOF) Then
'          wOrRecSet.MoveFirst
'          wWtDiff = 0
'          'wTotIrRmWt = wRsRmDet!qIrRmWt
'
'          If (wTotIrRmWt <> wOrRecSet!qOrRmWt) Then
'            wWtDiff = wOrRecSet!qOrRmWt - wTotIrRmWt
'            wDivFactor = IIF(wRmCtg = "C" Or wRmCtg = "D", 5, 1)
'            If UCase(wRsRmDet!qMainMet) = "Y" Then
'                wTotWtDiff = wTotWtDiff + MWLib.Div(wWtDiff, wDivFactor)
'            End If
'
'            '###
'            Set wIrRmWt = moCn.OpenResultset("Select * from " + wTmpTbl + _
'                         " Where qIrRmCd = '" + wRsRmDet!qIrRmCd + "' ")
'
'            If Not (wIrRmWt.EOF Or wIrRmWt.BOF) Then wIrRmWt.MoveFirst
'            Do While Not (wIrRmWt.EOF Or wIrRmWt.BOF)
'              wAddWt = (MWLib.Div(wWtDiff, wTotIrRmWt)) * wIrRmWt!qIrRmIWt
'              wTotIrRmWt = wTotIrRmWt - wIrRmWt!qIrRmIWt
'              wIrRmWt!qIrRmIWt = wIrRmWt!qIrRmIWt + wAddWt
'              wIrRmWt.Update
'              wWtDiff = wWtDiff - wAddWt
'              wIrRmWt.MoveNext
'            Loop
'          End If
'        End If
'      End If
'      wRsRmDet.MoveNext
'    Loop
'
'    Set wRsRmDet = moCn.OpenResultset("Select * from " + wTmpTbl + _
'                                      " where qIrMainMet= 'Y'")
'    If Not (wRsRmDet.EOF Or wRsRmDet.BOF) Then
'      wRsRmDet.MoveFirst
'      wRsRmDet!qIrRmIWt = wRsRmDet!qIrRmIWt - wTotWtDiff
'      wRsRmDet.Update
'    End If
'  End If
'
'  If UCase(ms_IdTcTyp) = UCase(ctInTcTyp) Then
'    moCn.Execute ("Update " + wTmpTbl + " Set qXtraIWt= .000, qXtraZWt= .000 ")
'
'    '********* URMI ****** 23/02/04 (commented code) *********
'    wDFrOrdYN = IIF(Trim(mIWtEqOrd) = "Y", "Y", "N")
'    If wDFrOrdYN = "N" Then wDFrOrdYN = IIF(InStr(1, wCtgStr, "," + "D" + ",") <> 0, "Y", "N")
'
'    wCFrOrdYN = IIF(Trim(mIWtEqOrd) = "Y", "Y", "N")
'    If wCFrOrdYN = "N" Then wCFrOrdYN = IIF(InStr(1, wCtgStr, "," + "C" + ",") <> 0, "Y", "N")
'
'    wDRrTcTyp = IIF(wDFrOrdYN = "Y", "XIO", "XIA")
'    wCRrTcTyp = IIF(wCFrOrdYN = "Y", "XIO", "XIA")
'
'    Call GetXtraWt(wDRrTcTyp, wCRrTcTyp, wTmpTbl, wCmDiaRtFA, wCmCsRtFA, GrdID.Value(GrdID.Row, "IdExpCmCd"), wDmCtg, _
'                   "qIrRmIWt", "qXtraIWt", "qIrRmIRt", mIdQty)
'    If whZWtEq = "A" Then
'      Call GetXtraWt("XZA", "XZA", wTmpTbl, wCmDiaRtFA, wCmCsRtFA, GrdID.Value(GrdID.Row, "IdExpCmCd"), wDmCtg, _
'                     "qIrRmZWt", "qXtraZWt", "qIrRmZRt", mIdQty)
'    End If
'    '*********************************************************
'  End If
'
'  Set wRsRmDet = moCn.OpenRes("select * from " + wTmpTbl + " Order By qIrMainMet")
'
'  wTotIXtraWt = 0: wTotZXtraWt = 0
'  wTotAWt = 0: wTotIWt = 0: wTotZWt = 0
'
'  With GrdIR
'    .StartCopy
'    Do While Not (wRsRmDet.EOF Or wRsRmDet.BOF)
'      .AddItem
'      wRmRow = .Rows - 1
'      If UCase(ms_IdTcTyp) = UCase(ctInTcTyp) Then
'        wDivFactor = IIF(wRsRmDet!qRmCtg = "C" Or wRsRmDet!qRmCtg = "D", 5, 1)
'        wTotIXtraWt = wTotIXtraWt + MWLib.Div(wRsRmDet!qXtraIWt, wDivFactor)
'        wTotZXtraWt = wTotZXtraWt + MWLib.Div(wRsRmDet!qXtraZWt, wDivFactor)
'
'        .Value(wRmRow, "IrSrNo") = wRmRow
'        .Value(wRmRow, "IrRmCd") = wRsRmDet!qIrRmCd
'        ' Zubin 212
'        .Value(wRmRow, "IrLotNo") = wRsRmDet!qIrLotNo
'        ' Zubin 212
'        .Value(wRmRow, "IrRmSz") = wRsRmDet!qIrRmSz
'        .Value(wRmRow, "IrStkRt") = wRsRmDet!qIrStkRt
'        .Value(wRmRow, "IrRmQty") = wRsRmDet!qIrRmQty
'        .Value(wRmRow, "IrRmAWt") = wRsRmDet!qIrRmAWt
'        .Value(wRmRow, "IrRmARt") = wRsRmDet!qIrRmARt
'        .Value(wRmRow, "IrSetSCd") = wRsRmDet!qIrSetSCd
'        .Value(wRmRow, "IrSetARt") = wRsRmDet!qIrSetARt
'        .Value(wRmRow, "IrSetIRt") = wRsRmDet!qIrSetIRt
'        .Value(wRmRow, "IrMainMet") = wRsRmDet!qIrMainMet
'        .Value(wRmRow, "IrSetIRt") = wRsRmDet!qIrSetIRt
'        .Value(wRmRow, "IrSetZRt") = wRsRmDet!qIrSetZRt
'
'        If wRsRmDet!qIrMainMet = "Y" Then
'          'Geeta****07/06/05
'          .Value(wRmRow, "IrAlySalRt") = moCn.GetFldVal("Select  OrAlySalRt from OrdRm where OrCoCd='" + gs_CoCd + "' and OrTc= '" + GrdID.Value(GrdID.Row, "IdExpOdTc") + "' and " + _
'                                                         "OrYy= '" + GrdID.Value(GrdID.Row, "IdExpOdYy") + "' and OrChr= '" + GrdID.Value(GrdID.Row, "IdExpOdChr") + "' and " + _
'                                                         "OrNo= " + CStr(GrdID.Value(GrdID.Row, "IdExpOdNo")) + " and OrSr= " + CStr(GrdID.Value(GrdID.Row, "IdExpOdSr")) + " and  OrMainMet='Y' ")
'          'Geeta***07/06/05
'             .Value(wRmRow, "IrRmIWt") = wRsRmDet!qIrRmIWt - (wTotIWt - wTotAWt)
'          If whZWtEq = "A" Then
'            .Value(wRmRow, "IrRmZWt") = wRsRmDet!qIrRmZWt - (wTotZWt - wTotAWt)
'          ElseIf whZWtEq = "I" Then
'            .Value(wRmRow, "IrRmZWt") = wRsRmDet!qIrRmIWt - (wTotIWt - wTotAWt)
'          End If
'        Else
'          ' Manoj 2.10.0
'          '.Value(wRmRow, "IrRmZWt") = wRsRmDet!qIrRmiWt + wRsRmDet!qXtraZWt
'          .Value(wRmRow, "IrRmIWt") = wRsRmDet!qIrRmIWt + _
'                IIF(UCase(wRsRmDet!qRmCtg) = "C" Or UCase(wRsRmDet!qRmCtg) = "D", Round(wRsRmDet!qXtraIWt, 2), wRsRmDet!qXtraIWt)
'
'          If whZWtEq = "A" Then
'            '.Value(wRmRow, "IrRmZWt") = wRsRmDet!qIrRmZWt + wRsRmDet!qXtraZWt
'            .Value(wRmRow, "IrRmZWt") = wRsRmDet!qIrRmZWt + _
'                  IIF(UCase(wRsRmDet!qRmCtg) = "C" Or UCase(wRsRmDet!qRmCtg) = "D", Round(wRsRmDet!qXtraZWt, 2), wRsRmDet!qXtraZWt)
'          ElseIf whZWtEq = "I" Then
'            '.Value(wRmRow, "IrRmZWt") = wRsRmDet!qIrRmIWt + wRsRmDet!qXtraIWt
'            .Value(wRmRow, "IrRmZWt") = wRsRmDet!qIrRmIWt + _
'                  IIF(UCase(wRsRmDet!qRmCtg) = "C" Or UCase(wRsRmDet!qRmCtg) = "D", Round(wRsRmDet!qXtraIWt, 2), wRsRmDet!qXtraIWt)
'          End If
'          ' Manoj 2.10.0
'        End If
'
'        wTotAWt = wTotAWt + MWLib.Div(.Value(wRmRow, "IrRmAWt"), wDivFactor)
'        wTotIWt = wTotIWt + MWLib.Div(.Value(wRmRow, "IrRmIWt"), wDivFactor)
'        wTotZWt = wTotZWt + MWLib.Div(.Value(wRmRow, "IrRmZWt"), wDivFactor)
'
'        If UCase(wRsRmDet!qRmCtg) = "G" And mIGldRtEqOrd <> "Y" Then
'          'Geeta***07/06/05 (Before 2.11)
'          '.Value(wRmRow, "IrRmIRt") = GetGldRt(GrdID.Value(GrdID.Row, "IdExpCmCd"), ADC("wIdLmgSal"), _
'                                      wRsRmDet!qIrRmCd, 0, "DM", GrdID.Value(GrdID.Row, "IdDmCd"), "S", "C")
'          .Value(wRmRow, "IrRmIRt") = GetGldRt(GrdID.Value(GrdID.Row, "IdExpCmCd"), ADC("wIdLmgSal"), _
'                                      wRsRmDet!qIrRmCd, .Value(wRmRow, "IrAlySalRt"), "DM", GrdID.Value(GrdID.Row, "IdDmCd"), "S", "C")
'          'Geeta***07/06/05 (Before 2.11)
'          If mLocalExp = "N" Then .Value(wRmRow, "IrRmZRt") = .Value(wRmRow, "IrRmIRt")
'        ElseIf UCase(wRsRmDet!qRmCtg) = "P" And mIGldRtEqOrd <> "Y" Then
'          'Geeta***07/06/05 (Before 2.11)
'          '.Value(wRmRow, "IrRmIRt") = GetGldRt(GrdID.Value(GrdID.Row, "IdExpCmCd"), ADC("wIdLmpSal"), _
'                                      wRsRmDet!qIrRmCd, 0, "DM", GrdID.Value(GrdID.Row, "IdDmCd"), "S", "C")
'          .Value(wRmRow, "IrRmIRt") = GetGldRt(GrdID.Value(GrdID.Row, "IdExpCmCd"), ADC("wIdLmpSal"), _
'                                      wRsRmDet!qIrRmCd, .Value(wRmRow, "IrAlySalRt"), "DM", GrdID.Value(GrdID.Row, "IdDmCd"), "S", "C")
'          'Geeta***07/06/05 (Before 2.11)
'          If mLocalExp = "N" Then .Value(wRmRow, "IrRmZRt") = .Value(wRmRow, "IrRmIRt")
'        '****** Sachin 2.12 - 06/12/05 - SJM
'        ElseIf UCase(wRsRmDet!qRmCtg) = "S" And mIGldRtEqOrd <> "Y" Then
'          .Value(wRmRow, "IrRmIRt") = GetGldRt(GrdID.Value(GrdID.Row, "IdExpCmCd"), ADC("wIdLmsSal"), _
'                                      wRsRmDet!qIrRmCd, .Value(wRmRow, "IrAlySalRt"), "DM", GrdID.Value(GrdID.Row, "IdDmCd"), "S", "C")
'          If mLocalExp = "N" Then .Value(wRmRow, "IrRmZRt") = .Value(wRmRow, "IrRmIRt")
'        '****** Sachin 2.12 - 06/12/05 - SJM
'        Else
'          .Value(wRmRow, "IrRmIRt") = wRsRmDet!qIrRmIRt
'          If mLocalExp = "N" Or (mLocalExp = "Y" And UCase(wRsRmDet!qRmCtg) <> "G" And UCase(wRsRmDet!qRmCtg) <> "P" And UCase(wRsRmDet!qRmCtg) <> "S") Then
'            .Value(wRmRow, "IrRmZRt") = wRsRmDet!qIrRmZRt
'          End If
'        End If
'        .Value(wRmRow, "IrRmZPureRt") = IIF(wRsRmDet!qRmPurityZ > 0, MWLib.Div(.Value(wRmRow, "IrRmZRt"), wRsRmDet!qRmPurityZ), .Value(wRmRow, "IrRmZRt"))
'      ElseIf UCase(ms_IdTcTyp) = UCase(ctCTBTcTyp) Then
'        .Value(wRmRow, "IrSrNo") = wRmRow
'        .Value(wRmRow, "IrRmCd") = wRsRmDet!qIrRmCd
'        ' Zubin 212
'        .Value(wRmRow, "IrLotNo") = wRsRmDet!qIrLotNo
'        ' Zubin 212
'        .Value(wRmRow, "IrRmSz") = wRsRmDet!qIrRmSz
'        .Value(wRmRow, "IrStkRt") = wRsRmDet!qIrStkRt
'        .Value(wRmRow, "IrRmQty") = wRsRmDet!qIrRmQty
'        .Value(wRmRow, "IrRmAWt") = wRsRmDet!qIrRmAWt
'
'        .Value(wRmRow, "IrRmIWt") = wRsRmDet!qIrRmIWt
'
'        .Value(wRmRow, "IrRmZWt") = wRsRmDet!qIrRmZWt
'        '.Value(wRmRow, "IrRmARt") = wRsRmDet!qIrRmARt
'        '.Value(wRmRow, "IrRmIRt") = wRsRmDet!qIrRmIRt
'        '.Value(wRmRow, "IrRmZRt") = wRsRmDet!qIrRmZRt
'        .Value(wRmRow, "IrSetSCd") = wRsRmDet!qIrSetSCd
'        .Value(wRmRow, "IrSetARt") = wRsRmDet!qIrSetARt
'        '.Value(wRmRow, "IrSetIRt") = wRsRmDet!qIrSetIRt
'        '.Value(wRmRow, "IrSetZRt") = wRsRmDet!qIrSetZRt
'        .Value(wRmRow, "IrRmZPureRt") = 0
'        .Value(wRmRow, "IrMainMet") = wRsRmDet!qIrMainMet
'      End If
'
'      ' Manoj 2.10.0
'      ' If UCase(wRsRmDet!qRmCtg) = "C" Or UCase(wRsRmDet!qRmCtg) = "D" Then
'      '   .Value(wRmRow, "IrRmAWt") = Round(.Value(wRmRow, "IrRmAWt"), 2)
'      '   .Value(wRmRow, "IrRmIWt") = Round(.Value(wRmRow, "IrRmIWt"), 2)
'      '   .Value(wRmRow, "IrRmZWt") = Round(.Value(wRmRow, "IrRmZWt"), 2)
'      ' End If
'      ' Manoj 2.10.0
'
'      .SaveRec (wRmRow)
'      wRsRmDet.MoveNext
'    Loop
'    .EndCopy
'  End With
'  '******Geeta***Emr208*************
'  '*********Zadd is Calculated in Make Lab ****
''  If UCase(ms_IdTcTyp) = UCase(ctInTcTyp) Then
''    wf_CnvRt = 1
''  ElseIf UCase(ms_IdTcTyp) = UCase(ctCTBTcTyp) Then
''    wf_CnvRt = moCn.GetFldVal("Select InCnvRt from InvHd where " + _
''                              "InCoCd= '" + gs_CoCd + "' and InTc= '" + adc("wIdTc") + "' " + _
''                              "and InYy= '" + adc("wIdYy") + "' and InChr= '" + adc("wIdChr") + "' " + _
''                              "and InNo= " + CStr(adc("wIdNo")))
''  End If
''
''  wf_AddLab = moCn.GetFldVal("Select sum(IrRmIVal- IrRmZVal* " + CStr(wf_CnvRt) + ") from InvRm " + _
'              "join RmMst on RmCd= IrRmCd where IrCoCd= '" + gs_CoCd + "' and " + _
'              "IrTc= '" + adc("wIdTc") + "' and IrYy= '" + adc("wIdYy") + "' and " + _
'              "IrChr= '" + adc("wIdChr") + "' and IrNo= " + CStr(adc("wIdNo")) + " and " + _
'              "IrSr= " + CStr(GrdID.Value(GrdID.Row, "IdSr")) + " and RmCtg in ('G', 'P')")
''
''  If mLocalExp = "Y" And Round(MWLib.Div(wf_AddLab, wf_CnvRt), 3) <> 0 Then
''    'And ms_IdTcTyp = ctInTcTyp
''    With GrdIL
''      .StartCopy
''      .AddItem
''      wLabRow = .Rows - 1
''      .Value(wLabRow, "iLSrNo") = wLabRow
''      .Value(wLabRow, "iLMCd") = ctAddLMCd
''      .Value(wLabRow, "iLSCd") = ctAddLSCd
''      .Value(wLabRow, "iLQty") = 1
''      .Value(wLabRow, "iLARt") = 0
''      .Value(wLabRow, "iLiRt") = 0
''      .Value(wLabRow, "iLZRt") = MWLib.Div(wf_AddLab, wf_CnvRt)
''      .Value(wLabRow, "ilQw") = "Q"
''      .Value(wLabRow, "iLZVal") = MWLib.Div(wf_AddLab, wf_CnvRt)
''      .SaveRec (wLabRow)
''      .EndCopy
''    End With
''  End If
''
'  Set wRsRmDet = Nothing
'  Set wOrRecSet = Nothing
'
'  '*** (Bef 18/01/2006 Ver 2.12 without the change for WtFrOrd)
End Sub
Public Sub MakeRmFromFg_(ByVal mIWtEqOrd As String, ByVal mIWtFrOrd As String, _
                        ByVal mIGldRtEqOrd As String, ByVal mIdQty As Double)
'  '*** (Bef 31/01/06)
'  Dim whZWtEq As String, wCtgStr As String, wDmCtg As String, wTmpTbl As String
'  Dim wOrStr As String, wStr As String, wMainMet As Boolean
'  Dim wRsRmDet As MDORowSet, wOrRecSet As MDORowSet
'  Dim wDivFactor As Single, wRmRow As Single, wLabRow As Single, wf_AddLab As Single
'  Dim wTotIXtraWt As Double, wTotZXtraWt As Double, wTotAWt As Double
'  Dim wTotIWt As Double, wTotZWt As Double
'  Dim wCmDiaRtFA As String, wCmCsRtFA As String, wf_CnvRt As Single
'  Dim wDRrTcTyp As String, wCRrTcTyp As String, wDFrOrdYN As String
'  Dim wCFrOrdYN As String
'  Dim wSetRt As Boolean
'  Dim wTotWtDiff As Double, wWtDiff As Double, wTotIrRmWt As Double, wAddWt As Double
'  Dim wIrRmWt As MDORowSet, wRmCtg As String, wFrOrdYN As String, wMainMetWt As Double
'  Dim wFrRmWt As Double, wFrRmQty As Double, wOrTotWt As Double, wOrTotQty As Double
'  Dim wFRmDet As MDORowSet
'  whZWtEq = moCn.GetFldVal("Select hZWtEq from Head where HCoCd= '" + gs_CoCd + "' and HCd='" + ctSelfCmCd + "'")
'  Dim wSetTyp As Boolean
'  Dim wOrLnStr As String
'
'  '*** (Jen 2.12)
'  Dim wRs_RmCtg As MDORowSet, wb_WtFrOrd As Boolean
'  '*** (Jen 2.12)
'
'  'mIWtFrOrd = "DC"
'  'wCtgStr= ",D,C"
'  wCtgStr = "," + moCn.GetFldVal("Select PDesc225 from Param where " + _
'              "PTyp='IWTFRORD' and PMCd = '" + mIWtFrOrd + "'") + ","
'
'  wCmDiaRtFA = moCn.GetFldVal("Select CmDiaRtFA from CustMst where CmCtg= 'C' and " + _
'               "CmCd= '" + GrdID.Value(GrdID.Row, "IdExpCmCd") + "' ")
'
'  wCmCsRtFA = moCn.GetFldVal("Select CmCSRtFA from CustMst where CmCtg= 'C' and " + _
'               "CmCd= '" + GrdID.Value(GrdID.Row, "IdExpCmCd") + "' ")
'
'
'  wDmCtg = moCn.GetFldVal("Select DmCtg from DsgMst where DmTcTyp= 'DM' and DmSz= '' and " + _
'               "DmCd= '" + GrdID.Value(GrdID.Row, "IdDmCd") + "' ")
'  wTmpTbl = "#TmpInvRm"
'
'  '*******Geeta*******Emr207
'  Dim ws_OrLn As String
'  ws_OrLn = "(case when Pvalue3='Y' then (Case when PValue2='Y' then  " + _
'          "cast(OrRmPtr as Decimal(12,4)) else (case when OrPrdQty=0 then 0 else  " + _
'          "cast((OrPrdWt/OrPrdQty) as Decimal(12,4)) end ) end )  " + _
'          "else cast(Orln1  as Decimal(12,4)) end ) "
'  If UCase(ms_IdTcTyp) = UCase(ctInTcTyp) Then
'    ' Zubin 212 (Introduced Lot No)
'    '****** (Jen 2.12) (Included qOrRmExists in the select list) ******
'    wStr = "Select (case when max(isnull(qOrRmCd, ''))= '' then '' else 'Y' end) as qOrRmExists, " + _
'           "FrRmCd as qIrRmCd, FrLotNo as qIrLotNo, FrRmSz as qIrRmSz, sum(FrRmStkRt*FrRmWt)/(case when sum(FrRmWt)> 0  " + _
'           "then sum(FrRmWt) else 1 end) as qIrStkRt, " + _
'           "sum(FrRmQty) as qIrRmQty, sum(FrRmWt) As qIrRmAWt, sum(FrRmWt) as qIrRmIWt, " + _
'           "sum(FrRmWt) as qIrRmZWt,max(Isnull(qOrWt,0)) as qOrWt,max(IsNull(qOrQty,0)) as qOrQty, 9999.999 as qXtraIWt, 9999.999 as qXtraZWt, " + _
'           "999999.999 as qIrRmARt, 999999.999 as qIrRmIRt, 999999.999 as qIrRmZRt, IsNull(qOrSetSCd, '') as qIrSetSCd, " + _
'           "isnull(qOrSetSalRt, 0.00) as qIrSetARt, isnull(qOrSetSalRt, 0.00) as qIrSetIRt, isnull(qOrSetSalRt, 0.00) as qIrSetZRt, " + _
'           "Space(1) as qIrMainMet, max(RmPurityZ) as qRmPurityZ, max(RmCtg) as qRmCtg, max(RmSCtg) as qRmSCtg " + _
'           "into " + wTmpTbl + " from InvFgd join FgRm on FrCoCd= IfCoCd and FrTc= IfFdTc and FrYy= IfFdyy and " + _
'           "FrChr= IfFdChr and FrNo = IfFdNo And FrSr = IfFdSr " + _
'           "Left Outer Join (select OrRmCd as qOrRmCd, " + ws_OrLn + " as qOrLn1, OrSetSCd as qOrSetSCd, " + _
'           "OrSetSalRt as qOrSetSalRt, sum(OrWt) as qOrWt, sum(OrQty) as qOrQty  from OrdRm  " + _
'           "Join RmMst on RmCd= OrRmCd " + _
'           "Join Param on PTyp= 'RMSCTG' and PMCd= RmCtg and PSCd= RmSCtg " + _
'           "where OrCoCd= '" + gs_CoCd + "'  and OrTc = '" + GrdID.Value(GrdID.Row, "IdExpOdTc") + "' " + _
'           "and OrYy= '" + GrdID.Value(GrdID.Row, "IdExpOdYy") + "' " + _
'           " and OrChr= '" + GrdID.Value(GrdID.Row, "IdExpOdChr") + "' and OrNo = " + CStr(GrdID.Value(GrdID.Row, "IdExpOdNo")) + _
'           " and OrSr= " + CStr(GrdID.Value(GrdID.Row, "IdExpOdSr")) + " Group By " + _
'           "OrRmCd, " + ws_OrLn + ", OrSetSCd, OrSetSalRt) Or2 On qOrRmCd= FrRmCd and qOrLn1= FrRmSz " + _
'           "Join RmMst on FrRmCd = RmCd " + _
'           "where IfCoCd = '" + gs_CoCd + "' and IfTc= '" + ADC("wIdTc") + "' and IfYy= '" + _
'           ADC("wIdYy") + "' and IfChr= '" + ADC("wIdChr") + "' and " + _
'           "IfNo = " + CStr(ADC("wIdNo")) + " and IfSr = " + CStr(GrdID.Value(GrdID.Row, "IdSr")) + " " + _
'           "Group by FrRmCd, FrLotNo, FrRmSz, qOrSetSCd, qOrSetSalRt "
'
'           '(Bef 2.12) OrCoCd, OrTc, OrYy, OrChr, OrNo, OrSr,
'
'  ElseIf UCase(ms_IdTcTyp) = UCase(ctCTBTcTyp) Then
'    ' Zubin 212 (Introduced Lot No)
'    wStr = "Select (case when max(isnull(qOrRmCd, ''))= '' then '' else 'Y' end) as qOrRmExists, " + _
'           "FrRmCd as qIrRmCd, FrLotNo as qIrLotNo, FrRmSz as qIrRmSz,sum(FrRmStkRt*FrRmWt)/(case when sum(FrRmWt)> 0  " + _
'           "then sum(FrRmWt) else 1 end) as qIrStkRt, sum(FrRmQty) as qIrRmQty, " + _
'           "sum(FrRmWt) As qIrRmAWt, sum(FrRmWt) as qIrRmIWt, sum(FrRmWt) as qIrRmZWt, " + _
'           "max(Isnull(qOrWt,0)) as qOrWt,max(IsNull(qOrQty,0)) as qOrQty, 0 as qXtraIWt, 0 as qXtraZWt, " + _
'           "0 as qIrRmARt, 0 as qIrRmIRt, 0 as qIrRmZRt, IsNull(qOrSetSCd, '') as qIrSetSCd, " + _
'           "isnull(qOrSetSalRt, 0.00) as qIrSetARt, isnull(qOrSetSalRt, 0.00) as qIrSetIRt, 0 as qIrSetZRt, " + _
'           "Space(1) as qIrMainMet, max(RmPurityZ) as qRmPurityZ, max(RmCtg) as qRmCtg, max(RmSCtg) as qRmSCtg " + _
'           "into " + wTmpTbl + " from InvFgd " + _
'           "join FgRm on FrCoCd= IfCoCd and FrTc= IfFdTc and FrYy= IfFdYy and FrChr= IfFdChr and FrNo = IfFdNo And FrSr = IfFdSr " + _
'           "Left Outer Join (select OrRmcd as qOrRmcd ," + ws_OrLn + " as qOrLn1, OrSetSCd as qOrSetSCd, " + _
'           "OrSetSalRt as qOrSetSalRt, sum(OrWt) as qOrWt, sum(OrQty) as qOrQty from OrdRm " + _
'           "Join RmMst on RmCd= OrRmCd " + _
'           "Join Param on PTyp= 'RMSCTG' and PMCd= RmCtg and PSCd= RmSCtg " + _
'           "where OrCoCd= '" + gs_CoCd + "' and OrTc= '" + GrdID.Value(GrdID.Row, "IdExpOdTc") + "' " + _
'           "and OrYy= '" + GrdID.Value(GrdID.Row, "IdExpOdYy") + "' " + _
'           " and OrChr= '" + GrdID.Value(GrdID.Row, "IdExpOdChr") + "' and OrNo = " + CStr(GrdID.Value(GrdID.Row, "IdExpOdNo")) + _
'           " and OrSr= " + CStr(GrdID.Value(GrdID.Row, "IdExpOdSr")) + " group by " + _
'           "OrRmCd, " + ws_OrLn + ", OrSetSCd, OrSetSalRt) Or2 On qOrRmcd= FrRmCd and qOrln1= FrRmSz " + _
'           "Join RmMst on FrRmCd= RmCd " + _
'           "where IfCoCd = '" + gs_CoCd + "' and IfTc= '" + ADC("wIdTc") + "' " + _
'           "and IfYy= '" + ADC("wIdYy") + "' and IfChr= '" + ADC("wIdChr") + "' and " + _
'           "IfNo = " + CStr(ADC("wIdNo")) + " and IfSr = " + CStr(GrdID.Value(GrdID.Row, "IdSr")) + " " + _
'           "Group by FrRmCd, FrLotNo, FrRmSz, qOrSetSCd, qOrSetSalRt "
'
'           '(Bef 2.12) OrCoCd, OrTc, OrYy, OrChr, OrNo, OrSr,
'
'  End If
'  '**********Geeta******Emr207
'
'  '*** Creating Temporary Table #TmpInvRm ***
'  moCn.CreateTmpTable wTmpTbl, wStr
'  'moCn.Execute (wStr)
'
'  Set wRsRmDet = moCn.OpenResultset("Select * from " + wTmpTbl)
'
'  '**********Geeta******Emr207
'  '*** (Bef 2.11)
'  'If moCn.RecSeek("select count(distinct qIrSetIRt) From " + wTmpTbl + " " + _
'  '                  "group by qIrRmCd,qIrRmSz Having count(distinct qIrSetIRt)>1") Then
'  '*** (Bef 2.11)
'
'  '*** 2.11 (Jen)
'  ' Zubin 212
'  'If moCn.RecSeek("select count(distinct Str(qIrSetIRt, 16, 2) + qIrSetSCd) From " + wTmpTbl + " " + _
'                  "group by qIrRmCd, qIrLotNo, qIrRmSz Having count(distinct Str(qIrSetIRt, 16, 2) + qIrSetSCd)> 1") Then
'  '*** 2.11 (Jen)
'
'  '*** (Jen 2.12)
'  If mIWtEqOrd = "Y" Or mIWtFrOrd <> "" Or moCn.RecSeek("select count(distinct Str(qIrSetIRt, 16, 2) + qIrSetSCd) From " + wTmpTbl + " " + _
'                "group by qIrRmCd, qIrLotNo, qIrRmSz Having count(distinct Str(qIrSetIRt, 16, 2) + qIrSetSCd)> 1") Then
'  '*** (Jen 2.12)
'
'    '  To get Proportionate weights
'    ' Zubin 212
'
'    '*** (Jen 2.12)
'    Set wRs_RmCtg = moCn.OpenRes("Select IsNull((Select min(qOrRmExists) From " + wTmpTbl + " where qIrRmCd= qOrRmCd), 'Y') as qOrdExists, " + _
'                    "min(qFgExists) as qFgExists, sum(qOrRmWt)* " + CStr(mIdQty) + " as qOrRmWt, " + _
'                    "sum(qOrRmQty)* " + CStr(mIdQty) + " as qOrRmQty, (Select sum(FrRmWt) from FgRm  " + _
'                    "   join InvFgd on FrCoCd= IfCoCd and FrTc= IfFdTc and FrYy= IfFdYy and  " + _
'                    "   FrChr= IfFdChr and FrNo = IfFdNo And FrSr = IfFdSr " + _
'                    "   where IfCoCd = '" + gs_CoCd + "' and IfTc= '" + ADC("wIdTc") + "' and " + _
'                    "   IfYy= '" + ADC("wIdYy") + "' and IfChr= '" + ADC("wIdChr") + "' and " + _
'                    "   IfNo = " + CStr(ADC("wIdNo")) + " and IfSr = " + CStr(GrdID.Value(GrdID.Row, "IdSr")) + " " + _
'                    "   and FrRmCd= qOrRmCd) as qFgRmWt, qOrRmCtg, qOrRmCd " + _
'                    "From (Select OrRmCtg as qOrRmCtg, OrRmCd as qOrRmCd, " + _
'                    " " + ws_OrLn + " as qOrRmSz, OrWt as qOrRmWt, OrQty as qOrRmQty, " + _
'                    "(case when not exists(Select 'x' from " + wTmpTbl + " where qIrRmCd= OrRmCd and qIrRmSz= " + ws_OrLn + ") then '' else 'Y' end) as qFgExists " + _
'                    "from OrdRm Join RmMst on RmCd= OrRmCd " + _
'                    "Join Param on PTyp= 'RMSCTG' and PMCd= RmCtg and PSCd= RmSCtg " + _
'                    "where OrCoCd= '" + gs_CoCd + "' and OrTc= '" + GrdID.Value(GrdID.Row, "IdExpOdTc") + "' " + _
'                    "and OrYy= '" + GrdID.Value(GrdID.Row, "IdExpOdYy") + "' " + _
'                    " and OrChr= '" + GrdID.Value(GrdID.Row, "IdExpOdChr") + "' and OrNo = " + CStr(GrdID.Value(GrdID.Row, "IdExpOdNo")) + _
'                    " and OrSr= " + CStr(GrdID.Value(GrdID.Row, "IdExpOdSr")) + " and OrSetSCd<> '') Ord " + _
'                    " Group By qOrRmCtg, qOrRmCd ")
'    Do While Not (wRs_RmCtg.EOF Or wRs_RmCtg.BOF)
'      wFrOrdYN = Trim(mIWtEqOrd)
'      If wFrOrdYN = "N" Then wFrOrdYN = IIF(InStr(1, wCtgStr, "," + wRs_RmCtg!qOrRmCtg + ",") <> 0, "Y", "N")
'      '### Debug.Print "OrdExists: " & wRs_RmCtg!qOrdExists & "   FgExists: " & wRs_RmCtg!qFgExists & "   WtFrOrd: " & wFrOrdYN & "   OrWt: " & wRs_RmCtg!qOrRmWt & "   FgWt: " & wRs_RmCtg!qFgRmWt
'      If wRs_RmCtg!qOrdExists = "Y" And wRs_RmCtg!qFgExists = "Y" And wFrOrdYN = "Y" And wRs_RmCtg!qOrRmWt = wRs_RmCtg!qFgRmWt Then
'        wb_WtFrOrd = True
'      Else
'        wb_WtFrOrd = False
'      End If
'
'      Set wFRmDet = moCn.OpenRes("Select FrRmCd as qFrRmCd, FrLotNo as qFrLotNo, max(FrRmSz) as qFrRmsz, " + _
'                         " Sum(FrRmWt) as qFrRmWt, Sum(FrRmQty) as qFrRmQty from FgRm  " + _
'                         "join InvFgd on FrCoCd= IfCoCd and FrTc= IfFdTc and FrYy= IfFdYy and  " + _
'                         " FrChr= IfFdChr and FrNo = IfFdNo And FrSr = IfFdSr " + _
'                         "where IfCoCd = '" + gs_CoCd + "' and IfTc= '" + ADC("wIdTc") + "' and IfYy= '" + _
'                         ADC("wIdYy") + "' and IfChr= '" + ADC("wIdChr") + "' and " + _
'                         "IfNo = " + CStr(ADC("wIdNo")) + " and IfSr = " + CStr(GrdID.Value(GrdID.Row, "IdSr")) + " " + _
'                         "and FrRmCd= '" + wRs_RmCtg!qOrRmCd + "' " + _
'                         "Group By FrRmCd, FrLotNo " + IIF(wb_WtFrOrd = True, " ", ", FrRmSz ") + _
'                         "Order By FrRmCd, FrLotNo" + IIF(wb_WtFrOrd = True, " ", ", FrRmSz "))
'    '*** (Jen 2.12)
'
'      '*** (Bef 2.12)
'      'Set wFRmDet = moCn.OpenRes("Select FrRmCd as qFrRmCd, FrLotNo as qFrLotNo, max(FrRmSz) as qFrRmsz, " + _
'                         " Sum(FrRmWt) as qFrRmWt, Sum(FrRmQty) as qFrRmQty from FgRm  " + _
'                         "join InvFgd on FrCoCd= IfCoCd and FrTc= IfFdTc and FrYy= IfFdYy and  " + _
'                         " FrChr= IfFdChr and FrNo = IfFdNo And FrSr = IfFdSr " + _
'                         "where IfCoCd = '" + gs_CoCd + "' and IfTc= '" + ADC("wIdTc") + "' and IfYy= '" + _
'                         ADC("wIdYy") + "' and IfChr= '" + ADC("wIdChr") + "' and " + _
'                         "IfNo = " + CStr(ADC("wIdNo")) + " and IfSr = " + CStr(GrdID.Value(GrdID.Row, "IdSr")) + " " + _
'                         "Group By FrRmCd, FrLotNo, FrRmSz Order By FrRmCd, FrLotNo, FrRmSz")
'      '*** (Bef 2.12)
'
'      Do While Not (wFRmDet.EOF Or wFRmDet.BOF)
'        wFrRmWt = wFRmDet!qFrRmWt
'        wFrRmQty = wFRmDet!qFrRmQty
'        ' Zubin 212
'        '*** (Bef 2.12)
'        'wOrTotWt = moCn.GetFldVal("Select Sum(qOrWt) From " + wTmpTbl + " Where qIrRmCd= '" + wFRmDet!qFrRmCd + "' and qIrRmSz= '" + CStr(wFRmDet!qFrRmSz) + "' and qIrLotNo= '" + CStr(wFRmDet!qFrLotNo) + "' Group By qIrRmcd, qIrLotNo, qIrRmSz")
'        'wOrTotQty = moCn.GetFldVal("Select Sum(qOrQty) From " + wTmpTbl + "  Where qIrRmCd= '" + wFRmDet!qFrRmCd + "' and qIrRmSz= '" + CStr(wFRmDet!qFrRmSz) + "' and qIrLotNo= '" + CStr(wFRmDet!qFrLotNo) + "' Group By qIrRmCd, qIrLotNo, qIrRmSz")
'        '*** (Bef 2.12)
'
'        '*** (Jen 2.12)
'        If wb_WtFrOrd = True Then
'          wOrTotWt = wRs_RmCtg!qOrRmWt
'          wOrTotQty = wRs_RmCtg!qOrRmQty
'          'wOrTotWt = moCn.GetFldVal("Select Sum(qOrWt) From " + wTmpTbl + " Where qIrRmCd= '" + wFRmDet!qFrRmCd + "' " + IIF(wb_WtFrOrd = True, " ", " and qIrRmSz= '" + CStr(wFRmDet!qFrRmSz) + "' ") + " and qIrLotNo= '" + CStr(wFRmDet!qFrLotNo) + "' Group By qIrRmcd, qIrLotNo " + IIF(wb_WtFrOrd = True, " ", ", qIrRmSz "))
'          'wOrTotQty = moCn.GetFldVal("Select Sum(qOrQty) From " + wTmpTbl + "  Where qIrRmCd= '" + wFRmDet!qFrRmCd + "' " + IIF(wb_WtFrOrd = True, " ", " and qIrRmSz= '" + CStr(wFRmDet!qFrRmSz) + "' ") + " and qIrLotNo= '" + CStr(wFRmDet!qFrLotNo) + "' Group By qIrRmCd, qIrLotNo " + IIF(wb_WtFrOrd = True, " ", ", qIrRmSz "))
'        Else
'          wOrTotWt = moCn.GetFldVal("Select Sum(qOrWt) From " + wTmpTbl + " Where qIrRmCd= '" + wFRmDet!qFrRmCd + "' and qIrRmSz= '" + CStr(wFRmDet!qFrRmSz) + "' and qIrLotNo= '" + CStr(wFRmDet!qFrLotNo) + "' Group By qIrRmcd, qIrLotNo, qIrRmSz")
'          wOrTotQty = moCn.GetFldVal("Select Sum(qOrQty) From " + wTmpTbl + "  Where qIrRmCd= '" + wFRmDet!qFrRmCd + "' and qIrRmSz= '" + CStr(wFRmDet!qFrRmSz) + "' and qIrLotNo= '" + CStr(wFRmDet!qFrLotNo) + "' Group By qIrRmCd, qIrLotNo, qIrRmSz")
'        End If
'        '*** (Jen 2.12)
'        ' Zubin 212
'
'        If (wRsRmDet.RecCount > 0) Then wRsRmDet.MoveFirst
'        Do While Not (wRsRmDet.EOF Or wRsRmDet.BOF)
'          ' Zubin 212
'          '*** (Jen 2.12)
'          If wFRmDet!qFrRmCd = wRsRmDet!qIrRmCd And wFRmDet!qFrLotNo = wRsRmDet!qIrLotNo And (wb_WtFrOrd = True Or (wb_WtFrOrd = False And wFRmDet!qFrRmSz = wRsRmDet!qIrRmSz)) Then
'          '*** (Jen 2.12)
'          '*** (Bef 2.12)
'          'If wFRmDet!qFrRmCd = wRsRmDet!qIrRmCd And wFRmDet!qFrLotNo = wRsRmDet!qIrLotNo And wFRmDet!qFrRmSz = wRsRmDet!qIrRmSz Then
'          '*** (Bef 2.12)
'            '*** (Bef 2.11)
'            'If moCn.RecSeek("select count(distinct qIrSetIRt) From " + wTmpTbl + " " + _
'            '                "Where qIrRmCd= '" + wRsRmDet!qIrRmCd + "' " + _
'            '                " and qIrRmSz = " + CStr(wRsRmDet!qIrRmSz) + _
'            '                "group by qIrRmCd,qIrRmSz Having count(distinct qIrSetIRt)>1") Then
'            '*** (Bef 2.11)
'            ' Zubin 212
'            '*** (Bef 2.12)
'            'If moCn.RecSeek("select count(distinct Str(qIrSetIRt, 16, 2) + qIrSetSCd) From " + wTmpTbl + " " + _
'                            "Where qIrRmCd= '" + wRsRmDet!qIrRmCd + "' " + _
'                            " and qIrLotNo= '" + CStr(wRsRmDet!qIrLotNo) + "' " + _
'                            " and qIrRmSz= " + CStr(wRsRmDet!qIrRmSz) + _
'                            "group by qIrRmCd, qIrLotNo, qIrRmSz Having count(distinct Str(qIrSetIRt, 16, 2) + qIrSetSCd)> 1") Then
'            '*** (Bef 2.12)
'            '*** (Jen 2.12)
'            If moCn.RecSeek("select count(distinct Str(qIrSetIRt, 16, 2) + qIrSetSCd) From " + wTmpTbl + " " + _
'                            "Where qIrRmCd= '" + wRsRmDet!qIrRmCd + "' " + _
'                            " and qIrLotNo= '" + CStr(wRsRmDet!qIrLotNo) + "' " + _
'                            IIF(wb_WtFrOrd = True, " ", " and qIrRmSz= " + CStr(wRsRmDet!qIrRmSz)) + _
'                            "group by qIrRmCd, qIrLotNo" + IIF(wb_WtFrOrd = True, " ", ", qIrRmSz ") + _
'                            IIF(wb_WtFrOrd = True, " ", "Having count(distinct Str(qIrSetIRt, 16, 2) + qIrSetSCd)> 1")) Then
'            '*** (Jen 2.12)
'
'              wRsRmDet.edit
'              '### Debug.Print "FrWt: " & wRsRmDet!qIrRmIWt & "   OrWt: " & wRsRmDet!qOrWt * IIF(wb_WtFrOrd = True, mIdQty, 1) & "   TotOrd: " & wOrTotWt & "   TotFg: " & wFrRmWt
'              '*** (Bef 2.12) wRsRmDet!qIrRmIWt = Round(MWLib.Div(wRsRmDet!qOrWt, wOrTotWt) * wFrRmWt, 3)
'              '*** (Jen 2.12)
'              If wb_WtFrOrd = True Then
'                wRsRmDet!qIrRmIWt = wRsRmDet!qOrWt * mIdQty
'                'wRsRmDet!qIrRmIWt = wRsRmDet!qOrWt * IIF(wb_WtFrOrd = True, mIdQty, 1)
'              Else
'                wRsRmDet!qIrRmIWt = Round(MWLib.Div(wRsRmDet!qOrWt, wOrTotWt) * wFrRmWt, 3)
'              End If
'              '*** (Jen 2.12)
'              wRsRmDet!qIrRmAWt = wRsRmDet!qIrRmIWt
'              wRsRmDet!qIrRmZWt = wRsRmDet!qIrRmIWt
'              wOrTotWt = wOrTotWt - wRsRmDet!qOrWt
'              wFrRmWt = wFrRmWt - wRsRmDet!qIrRmIWt
'              '*** (Bef 2.12)
'              'wRsRmDet!qIrRmQty = Round(MWLib.Div(wRsRmDet!qOrQty, wOrTotQty) * wFrRmQty)
'              'wOrTotQty = wOrTotQty - wRsRmDet!qOrQty
'              '*** (Bef 2.12)
'              '*** (Jen 2.12)
'              If wb_WtFrOrd = True Then
'                wRsRmDet!qIrRmQty = Round(MWLib.Div(wRsRmDet!qOrQty * mIdQty, wOrTotQty) * wFrRmQty)
'                wOrTotQty = wOrTotQty - wRsRmDet!qOrQty * mIdQty
'              Else
'                wRsRmDet!qIrRmQty = Round(MWLib.Div(wRsRmDet!qOrQty, wOrTotQty) * wFrRmQty)
'                wOrTotQty = wOrTotQty - wRsRmDet!qOrQty
'              End If
'              '*** (Jen 2.12)
'              wFrRmQty = wFrRmQty - wRsRmDet!qIrRmQty
'              wRsRmDet.Update
'            End If
'          End If
'          wRsRmDet.MoveNext
'        Loop
'        wFRmDet.MoveNext
'      Loop
'      Set wFRmDet = Nothing
'
'    '*** (Jen 2.12)
'      wRs_RmCtg.MoveNext
'    Loop
'    Set wRs_RmCtg = Nothing
'    '*** (Jen 2.12)
'  End If
'
'  'If (wRsRmDet.RecCount > 0) Then wRsRmDet.MoveFirst
'  'Do While Not (wRsRmDet.EOF Or wRsRmDet.BOF)
'  '  Debug.Print CStr(wRsRmDet!qIrRmCd) + "," + CStr(wRsRmDet!qIrRmSz) + "," + CStr(wRsRmDet!qIrRmQty) + "," + CStr(wRsRmDet!qIrRmAWt) + "," + CStr(wRsRmDet!qOrWt) + "," + CStr(wRsRmDet!qOrQty)
'  '  wRsRmDet.MoveNext
'  'Loop
'
'  '**********Geeta******Emr207
'  If (wRsRmDet.RecCount > 0) Then wRsRmDet.MoveFirst
'  Do While Not (wRsRmDet.EOF Or wRsRmDet.BOF)
'    wRsRmDet.edit
'
'    wOrLnStr = IIF(moCn.GetFldVal("select PValue3 From Param Join  RmMst " + _
'                            " on PTyp='RMSCTG' and PMCd=RmCtg and PSCd=RmSCtg " + _
'                            "where RmCd='" + wRsRmDet!qIrRmCd + "' ") = "Y", _
'                   IIF(moCn.GetFldVal("select PValue2 From Param Join  RmMst " + _
'                            "on PTyp='RMSCTG' and PMCd=RmCtg and PSCd=RmSCtg " + _
'                            "where RmCd='" + wRsRmDet!qIrRmCd + "' ") = "Y", "OrRmPtr", _
'                            " case when OrPrdQty =0 then 0 else OrPrdWt/OrPrdQty end "), "OrLn1")
'
'    '*** (Bef 2.11)
'    'If moCn.RecSeek("select count(distinct qIrSetIRt) From " + wTmpTbl + " " + _
'                    "Where qIrRmCd= '" + wRsRmDet!qIrRmCd + "' " + _
'                    " and qIrRmSz= " + CStr(wRsRmDet!qIrRmSz) + _
'                    "group by qIrRmCd, qIrRmSz Having count(distinct qIrSetIRt)> 1") Then
'    '*** (Bef 2.11)
'
'    '*** 2.11 (Jen)
'    ' Zubin 212
'    If moCn.RecSeek("select count(distinct Str(qIrSetIRt, 16, 2) + qIrSetSCd) From " + wTmpTbl + " " + _
'                    "Where qIrRmCd= '" + wRsRmDet!qIrRmCd + "' " + _
'                    " and qIrLotNo= '" + CStr(wRsRmDet!qIrLotNo) + "' " + _
'                    " and qIrRmSz= " + CStr(wRsRmDet!qIrRmSz) + _
'                    "group by qIrRmCd, qIrLotNo, qIrRmSz Having count(distinct Str(qIrSetIRt, 16, 2) + qIrSetSCd)> 1") Then
'    '*** 2.11 (Jen)
'      '*******Geeta*******Emr207
'      wOrStr = "Select OrCstRt, OrSalRt, OrSetSCd, OrSetCstRt, OrSetSalRt, OrMainMet, OrRmCtg " + _
'              "From OrdRm where OrCoCd= '" + gs_CoCd + "' " + _
'              " and OrTc = '" + GrdID.Value(GrdID.Row, "IdExpOdTc") + "' " + _
'              " and OrYy= '" + GrdID.Value(GrdID.Row, "IdExpOdYy") + "' " + _
'              " and OrChr= '" + GrdID.Value(GrdID.Row, "IdExpOdChr") + "' " + _
'              " and OrNo = " + CStr(GrdID.Value(GrdID.Row, "IdExpOdNo")) + _
'              " and OrSr= " + CStr(GrdID.Value(GrdID.Row, "IdExpOdSr")) + _
'              " and OrRmCd= '" + wRsRmDet!qIrRmCd + "' " + _
'              " and " + wOrLnStr + " = " + CStr(wRsRmDet!qIrRmSz) + _
'              " and OrSetSCd= '" + wRsRmDet!qIrSetSCd + "' " + _
'              " and OrSetSalRt= " + CStr(wRsRmDet!qIrSetIRt) + _
'              " Order by OrMainMet Desc, OrSalRt Desc"
'
'      'Set wOrRecSet = moCn.OpenRes(wOrStr)
'    Else
'      wOrStr = "Select OrCstRt, OrSalRt, OrSetSCd, OrSetCstRt, OrSetSalRt, OrMainMet, OrRmCtg " + _
'                "From OrdRm where OrCoCd= '" + gs_CoCd + "' " + _
'                " and OrTc = '" + GrdID.Value(GrdID.Row, "IdExpOdTc") + "' " + _
'                " and OrYy= '" + GrdID.Value(GrdID.Row, "IdExpOdYy") + "' " + _
'                " and OrChr= '" + GrdID.Value(GrdID.Row, "IdExpOdChr") + "' " + _
'                " and OrNo = " + CStr(GrdID.Value(GrdID.Row, "IdExpOdNo")) + _
'                " and OrSr= " + CStr(GrdID.Value(GrdID.Row, "IdExpOdSr")) + _
'                " and OrRmCd= '" + wRsRmDet!qIrRmCd + "' " + _
'                " and " + wOrLnStr + " = " + CStr(wRsRmDet!qIrRmSz) + _
'                "Order by OrMainMet Desc, OrSalRt Desc"
'    End If
'    Set wOrRecSet = moCn.OpenRes(wOrStr)
'
'    '    Query to get weighted Average rate
'    '    wOrStr = "Select (Case when sum(OrWt)<>0 then sum(OrCstRt*OrWt)/sum(OrWt) else 0 end) as OrCstRt, " + _
'                  " (Case when sum(OrWt)<>0 then sum(OrSalRt*OrWt)/sum(OrWt) else 0 end) as OrSalRt, " + _
'                  " (select max(OrSetSCd) from OrdRm B where  B.OrCoCd= '" + gs_CoCd + "' " + _
'                  " and B.OrTc = '" + GrdID.Value(GrdID.Row, "IdExpOdTc") + "' " + _
'                  " and B.OrYy= '" + GrdID.Value(GrdID.Row, "IdExpOdYy") + "' " + _
'                  " and B.OrChr= '" + GrdID.Value(GrdID.Row, "IdExpOdChr") + "' " + _
'                  " and B.OrNo = " + CStr(GrdID.Value(GrdID.Row, "IdExpOdNo")) + _
'                  " and B.OrSr= " + CStr(GrdID.Value(GrdID.Row, "IdExpOdSr")) + _
'                  " and B.OrRmCd= '" + wRsRmDet!qIrRmCd + "' " + _
'                  " and B.OrLn1 = " + CStr(wRsRmDet!qIrRmSz) + _
'                  " and B.OrSetSalRt = max(A.OrSetSalRt)) as OrSetSCd, " + _
'                  " (Case when sum(OrQty)<>0 then sum(OrSetCstRt*OrQty)/sum(OrQty) else 0 end) as OrSetCstRt, " + _
'                  " (case when sum(OrQty)<>0 then sum(OrSetSalRt*OrQty)/sum(OrQty) else 0 end) as OrSetSalRt, " + _
'                  " max(OrMainMet), max(OrRmCtg) " + _
'                  " From OrdRm A where OrCoCd= '" + gs_CoCd + "' " + _
'                  " and OrTc = '" + GrdID.Value(GrdID.Row, "IdExpOdTc") + "' " + _
'                  " and OrYy= '" + GrdID.Value(GrdID.Row, "IdExpOdYy") + "' " + _
'                  " and OrChr= '" + GrdID.Value(GrdID.Row, "IdExpOdChr") + "' " + _
'                  " and OrNo = " + CStr(GrdID.Value(GrdID.Row, "IdExpOdNo")) + _
'                  " and OrSr= " + CStr(GrdID.Value(GrdID.Row, "IdExpOdSr")) + _
'                  " and OrRmCd= '" + wRsRmDet!qIrRmCd + "' " + _
'                  " and OrLn1 = " + CStr(wRsRmDet!qIrRmSz) + _
'                  " Group By OrRmCd, OrLn1 "
'
'    If (wOrRecSet.EOF Or wOrRecSet.BOF) Then
'      wOrStr = "Select OrCstRt, OrSalRt, OrSetSCd, OrSetCstRt, OrSetSalRt, OrMainMet, OrRmCtg " + _
'               "From OrdRm join RmRt on RrTcTyp= 'Rsz' and RrCmCd= '" + ctSelfCmCd + "' and " + _
'               "RrCtg= OrRmCtg and RrSCtg= OrRmSCtg " + _
'               "where OrCoCd= '" + gs_CoCd + "' " + _
'               " and OrTc = '" + GrdID.Value(GrdID.Row, "IdExpOdTc") + "' " + _
'               " and OrYy= '" + GrdID.Value(GrdID.Row, "IdExpOdYy") + "' " + _
'               " and OrChr = '" + GrdID.Value(GrdID.Row, "IdExpOdChr") + "' " + _
'               " and OrNo = " + CStr(GrdID.Value(GrdID.Row, "IdExpOdNo")) + _
'               " and OrSr= " + CStr(GrdID.Value(GrdID.Row, "IdExpOdSr")) + _
'               " and OrRmCd =  '" + wRsRmDet!qIrRmCd + "' " + _
'               " and " + wOrLnStr + " >= RrFrLn and " + wOrLnStr + " <= RrToLn and " + _
'               CStr(wRsRmDet!qIrRmSz) + " >= RrFrLn and " + _
'               CStr(wRsRmDet!qIrRmSz) + " <= RrToLn " + _
'               "Order by OrMainMet Desc, OrSalRt Desc "
'
'      Set wOrRecSet = moCn.OpenRes(wOrStr)
'    End If
'    If (wOrRecSet.EOF Or wOrRecSet.BOF) Then
'      wOrStr = "Select OrCstRt, OrSalRt, OrSetSCd, OrSetCstRt, OrSetSalRt, OrMainMet, OrRmCtg " + _
'               "From OrdRm where OrCoCd= '" + gs_CoCd + "' " + _
'               "  and OrTc = '" + GrdID.Value(GrdID.Row, "IdExpOdTc") + "' " + _
'               " and OrYy= '" + GrdID.Value(GrdID.Row, "IdExpOdYy") + "' " + _
'               " and OrChr= '" + GrdID.Value(GrdID.Row, "IdExpOdChr") + "' " + _
'               " and OrNo = " + CStr(GrdID.Value(GrdID.Row, "IdExpOdNo")) + _
'               " and OrSr= " + CStr(GrdID.Value(GrdID.Row, "IdExpOdSr")) + _
'               " and OrRmCd= '" + wRsRmDet!qIrRmCd + "' " + _
'               "Order by OrMainMet Desc, OrSalRt Desc"
'      Set wOrRecSet = moCn.OpenRes(wOrStr)
'    End If
'
'    If (wOrRecSet.EOF Or wOrRecSet.BOF) Then
'      wRsRmDet!qIrRmARt = 0#
'      wRsRmDet!qIrRmIRt = 0#
'      wRsRmDet!qIrSetSCd = ""
'      wRsRmDet!qIrSetARt = 0#
'      If UCase(ms_IdTcTyp) = UCase(ctInTcTyp) Then
'        wRsRmDet!qIrRmZRt = 0#
'        wRsRmDet!qIrSetIRt = 0#
'        wRsRmDet!qIrSetZRt = 0#
'      End If
'      wRsRmDet!qIrMainMet = "N"
'    Else
'      wOrRecSet.MoveFirst
'      wRsRmDet!qIrSetSCd = wOrRecSet!OrSetSCd
'      wRsRmDet!qIrSetARt = wOrRecSet!OrSetCstRt
'      If UCase(ms_IdTcTyp) = UCase(ctInTcTyp) Then
'        wRsRmDet!qIrRmARt = wOrRecSet!OrCstRt
'        wRsRmDet!qIrRmIRt = wOrRecSet!OrSalRt
'        wRsRmDet!qIrRmZRt = wOrRecSet!OrSalRt
'        wRsRmDet!qIrSetIRt = wOrRecSet!OrSetSalRt
'        wRsRmDet!qIrSetZRt = wOrRecSet!OrSetSalRt
'      End If
'
'      If wOrRecSet!OrMainMet = "Y" And wMainMet = False Then
'        wRsRmDet!qIrMainMet = "Y": wMainMet = True
'      ElseIf wOrRecSet!OrMainMet = "Y" And wMainMet = True Then
'        wRsRmDet!qIrMainMet = "N"
'      ElseIf wOrRecSet!OrMainMet = "N" Then
'        wRsRmDet!qIrMainMet = "N"
'      End If
'    End If
'
'    wRsRmDet.Update
'    wRsRmDet.MoveNext
'  Loop
'
'  If UCase(ms_IdTcTyp) = UCase(ctInTcTyp) Then
'
'    '****** check if mainMet wt > other Rm wt Diff. ***/
'    Set wRsRmDet = moCn.OpenRes("Select qIrRmCd, sum(qIrRmIWt) as qIrRmWt, " + _
'                                " max(qIrMainMet) as qMainMet from " + wTmpTbl + _
'                                " Group By qIrRmCd Order By qIrRmCd")
'    If (wRsRmDet.RecCount > 0) Then wRsRmDet.MoveFirst
'    wTotWtDiff = 0: wMainMetWt = 0
'    Do While Not (wRsRmDet.EOF Or wRsRmDet.BOF)
'      wFrOrdYN = Trim(mIWtEqOrd)
'      wTotIrRmWt = wRsRmDet!qIrRmWt
'      wRmCtg = moCn.GetFldVal("Select RmCtg from RmMst where RmCd= '" + wRsRmDet!qIrRmCd + "' ")
'      wDivFactor = IIF(wRmCtg = "C" Or wRmCtg = "D", 5, 1)
'
'      If wFrOrdYN = "N" Then wFrOrdYN = IIF(InStr(1, wCtgStr, "," + wRmCtg + ",") <> 0, "Y", "N")
'
'      If wFrOrdYN = "Y" Then
'        wOrStr = "Select sum(OrWt*" + CStr(mIdQty) + ") as qOrRmWt from OrdRm " + _
'                "where OrCoCd= '" + gs_CoCd + "' " + _
'                " and OrTc = '" + GrdID.Value(GrdID.Row, "IdExpOdTc") + "' " + _
'                " and OrYy= '" + GrdID.Value(GrdID.Row, "IdExpOdYy") + "' " + _
'                " and OrChr= '" + GrdID.Value(GrdID.Row, "IdExpOdChr") + "' " + _
'                " and OrNo = " + CStr(GrdID.Value(GrdID.Row, "IdExpOdNo")) + _
'                " and OrSr= " + CStr(GrdID.Value(GrdID.Row, "IdExpOdSr")) + _
'                " and OrRmCd= '" + wRsRmDet!qIrRmCd + "' "
'
'        Set wOrRecSet = moCn.OpenRes(wOrStr)
'        If Not (wOrRecSet.EOF Or wOrRecSet.BOF) Then
'          wOrRecSet.MoveFirst
'          wWtDiff = 0
'          'wTotIrRmWt = wRsRmDet!qIrRmWt
'          If (wTotIrRmWt <> wOrRecSet!qOrRmWt) Then
'            wWtDiff = wOrRecSet!qOrRmWt - wTotIrRmWt
'            If UCase(wRsRmDet!qMainMet) <> "Y" Then _
'               wTotWtDiff = wTotWtDiff + MWLib.Div(wWtDiff, wDivFactor)
'          End If
'        End If
'      End If
'
'      'urmi <>, = 23/02/04
'      If UCase(wRsRmDet!qMainMet) = "Y" Then _
'         wMainMetWt = MWLib.Div(wTotIrRmWt, wDivFactor)
'      wRsRmDet.MoveNext
'    Loop
'
'
'    'urmi <>, < 23/02/04
'    If wMainMetWt < wTotWtDiff Then
'      DispMsg "MainMetal Wt is less than the other Wt Diff.", etError
'      Exit Sub
'    End If
'
'    Set wRsRmDet = Nothing
'    Set wOrRecSet = Nothing
'    '********************************************************
'
'    Set wRsRmDet = moCn.OpenRes("Select qIrRmCd, sum(qIrRmIWt) as qIrRmWt, max(qIrMainMet) as qMainMet from " + wTmpTbl + _
'                                " Group By qIrRmCd Order By qIrRmCd")
'    If Not (wRsRmDet.EOF Or wRsRmDet.BOF) Then wRsRmDet.MoveFirst
'    wTotWtDiff = 0
'    Do While Not (wRsRmDet.EOF Or wRsRmDet.BOF)
'      wFrOrdYN = Trim(mIWtEqOrd)
'      wTotIrRmWt = wRsRmDet!qIrRmWt
'      wRmCtg = moCn.GetFldVal("Select RmCtg from RmMst where RmCd= '" + wRsRmDet!qIrRmCd + "' ")
'      If wFrOrdYN = "N" Then wFrOrdYN = IIF(InStr(1, wCtgStr, "," + wRmCtg + ",") <> 0, "Y", "N")
'
'      If wFrOrdYN = "Y" Then
'        wOrStr = "Select sum(OrWt*" + CStr(mIdQty) + ") as qOrRmWt from OrdRm " + _
'                "where OrCoCd= '" + gs_CoCd + "' " + _
'                " and OrTc = '" + GrdID.Value(GrdID.Row, "IdExpOdTc") + "' " + _
'                " and OrYy= '" + GrdID.Value(GrdID.Row, "IdExpOdYy") + "' " + _
'                " and OrChr= '" + GrdID.Value(GrdID.Row, "IdExpOdChr") + "' " + _
'                " and OrNo = " + CStr(GrdID.Value(GrdID.Row, "IdExpOdNo")) + _
'                " and OrSr= " + CStr(GrdID.Value(GrdID.Row, "IdExpOdSr")) + _
'                " and OrRmCd= '" + wRsRmDet!qIrRmCd + "' "
'
'        Set wOrRecSet = moCn.OpenRes(wOrStr)
'        If Not (wOrRecSet.EOF Or wOrRecSet.BOF) Then
'          wOrRecSet.MoveFirst
'          wWtDiff = 0
'          'wTotIrRmWt = wRsRmDet!qIrRmWt
'
'          If (wTotIrRmWt <> wOrRecSet!qOrRmWt) Then
'            wWtDiff = wOrRecSet!qOrRmWt - wTotIrRmWt
'            wDivFactor = IIF(wRmCtg = "C" Or wRmCtg = "D", 5, 1)
'            If UCase(wRsRmDet!qMainMet) = "Y" Then
'                wTotWtDiff = wTotWtDiff + MWLib.Div(wWtDiff, wDivFactor)
'            End If
'
'            '###
'            Set wIrRmWt = moCn.OpenResultset("Select * from " + wTmpTbl + _
'                         " Where qIrRmCd = '" + wRsRmDet!qIrRmCd + "' ")
'
'            If Not (wIrRmWt.EOF Or wIrRmWt.BOF) Then wIrRmWt.MoveFirst
'            Do While Not (wIrRmWt.EOF Or wIrRmWt.BOF)
'              '*** (Bef 2.12) wAddWt = (MWLib.Div(wWtDiff, wTotIrRmWt)) * wIrRmWt!qIrRmIWt
'              '*** (Jen 2.12)
'              wAddWt = Round((MWLib.Div(wWtDiff, wTotIrRmWt)) * wIrRmWt!qIrRmIWt, 3)
'              '*** (Jen 2.12)
'              '### Debug.Print "IrWt: " & (wIrRmWt!qIrRmIWt + wAddWt) & "   WtDiff: " & wWtDiff & "   TotIrWt: " & wTotIrRmWt
'              wTotIrRmWt = wTotIrRmWt - wIrRmWt!qIrRmIWt
'              wIrRmWt!qIrRmIWt = wIrRmWt!qIrRmIWt + wAddWt
'              wIrRmWt.Update
'              wWtDiff = wWtDiff - wAddWt
'              wIrRmWt.MoveNext
'            Loop
'          End If
'        End If
'      End If
'      wRsRmDet.MoveNext
'    Loop
'
'    Set wRsRmDet = moCn.OpenResultset("Select * from " + wTmpTbl + _
'                                      " where qIrMainMet= 'Y'")
'    If Not (wRsRmDet.EOF Or wRsRmDet.BOF) Then
'      wRsRmDet.MoveFirst
'      wRsRmDet!qIrRmIWt = wRsRmDet!qIrRmIWt - wTotWtDiff
'      wRsRmDet.Update
'    End If
'  End If
'
'  If UCase(ms_IdTcTyp) = UCase(ctInTcTyp) Then
'    moCn.Execute ("Update " + wTmpTbl + " Set qXtraIWt= .000, qXtraZWt= .000 ")
'
'    '********* URMI ****** 23/02/04 (commented code) *********
'    wDFrOrdYN = IIF(Trim(mIWtEqOrd) = "Y", "Y", "N")
'    If wDFrOrdYN = "N" Then wDFrOrdYN = IIF(InStr(1, wCtgStr, "," + "D" + ",") <> 0, "Y", "N")
'
'    wCFrOrdYN = IIF(Trim(mIWtEqOrd) = "Y", "Y", "N")
'    If wCFrOrdYN = "N" Then wCFrOrdYN = IIF(InStr(1, wCtgStr, "," + "C" + ",") <> 0, "Y", "N")
'
'    wDRrTcTyp = IIF(wDFrOrdYN = "Y", "XIO", "XIA")
'    wCRrTcTyp = IIF(wCFrOrdYN = "Y", "XIO", "XIA")
'
'    Call GetXtraWt(wDRrTcTyp, wCRrTcTyp, wTmpTbl, wCmDiaRtFA, wCmCsRtFA, GrdID.Value(GrdID.Row, "IdExpCmCd"), wDmCtg, _
'                   "qIrRmIWt", "qXtraIWt", "qIrRmIRt", mIdQty)
'    If whZWtEq = "A" Then
'      Call GetXtraWt("XZA", "XZA", wTmpTbl, wCmDiaRtFA, wCmCsRtFA, GrdID.Value(GrdID.Row, "IdExpCmCd"), wDmCtg, _
'                     "qIrRmZWt", "qXtraZWt", "qIrRmZRt", mIdQty)
'    End If
'    '*********************************************************
'  End If
'
'  Set wRsRmDet = moCn.OpenRes("select * from " + wTmpTbl + " Order By qIrMainMet")
'
'  wTotIXtraWt = 0: wTotZXtraWt = 0
'  wTotAWt = 0: wTotIWt = 0: wTotZWt = 0
'
'  With GrdIR
'    .StartCopy
'    Do While Not (wRsRmDet.EOF Or wRsRmDet.BOF)
'      .AddItem
'      wRmRow = .Rows - 1
'      If UCase(ms_IdTcTyp) = UCase(ctInTcTyp) Then
'        wDivFactor = IIF(wRsRmDet!qRmCtg = "C" Or wRsRmDet!qRmCtg = "D", 5, 1)
'        wTotIXtraWt = wTotIXtraWt + MWLib.Div(wRsRmDet!qXtraIWt, wDivFactor)
'        wTotZXtraWt = wTotZXtraWt + MWLib.Div(wRsRmDet!qXtraZWt, wDivFactor)
'
'        .Value(wRmRow, "IrSrNo") = wRmRow
'        .Value(wRmRow, "IrRmCd") = wRsRmDet!qIrRmCd
'        ' Zubin 212
'        .Value(wRmRow, "IrLotNo") = wRsRmDet!qIrLotNo
'        ' Zubin 212
'        .Value(wRmRow, "IrRmSz") = wRsRmDet!qIrRmSz
'        .Value(wRmRow, "IrStkRt") = wRsRmDet!qIrStkRt
'        .Value(wRmRow, "IrRmQty") = wRsRmDet!qIrRmQty
'        .Value(wRmRow, "IrRmAWt") = wRsRmDet!qIrRmAWt
'        .Value(wRmRow, "IrRmARt") = wRsRmDet!qIrRmARt
'        .Value(wRmRow, "IrSetSCd") = wRsRmDet!qIrSetSCd
'        .Value(wRmRow, "IrSetARt") = wRsRmDet!qIrSetARt
'        .Value(wRmRow, "IrSetIRt") = wRsRmDet!qIrSetIRt
'        .Value(wRmRow, "IrMainMet") = wRsRmDet!qIrMainMet
'        .Value(wRmRow, "IrSetIRt") = wRsRmDet!qIrSetIRt
'        .Value(wRmRow, "IrSetZRt") = wRsRmDet!qIrSetZRt
'
'        If wRsRmDet!qIrMainMet = "Y" Then
'          'Geeta****07/06/05
'          .Value(wRmRow, "IrAlySalRt") = moCn.GetFldVal("Select  OrAlySalRt from OrdRm where OrCoCd='" + gs_CoCd + "' and OrTc= '" + GrdID.Value(GrdID.Row, "IdExpOdTc") + "' and " + _
'                                                         "OrYy= '" + GrdID.Value(GrdID.Row, "IdExpOdYy") + "' and OrChr= '" + GrdID.Value(GrdID.Row, "IdExpOdChr") + "' and " + _
'                                                         "OrNo= " + CStr(GrdID.Value(GrdID.Row, "IdExpOdNo")) + " and OrSr= " + CStr(GrdID.Value(GrdID.Row, "IdExpOdSr")) + " and  OrMainMet='Y' ")
'          'Geeta***07/06/05
'             .Value(wRmRow, "IrRmIWt") = wRsRmDet!qIrRmIWt - (wTotIWt - wTotAWt)
'          If whZWtEq = "A" Then
'            .Value(wRmRow, "IrRmZWt") = wRsRmDet!qIrRmZWt - (wTotZWt - wTotAWt)
'          ElseIf whZWtEq = "I" Then
'            .Value(wRmRow, "IrRmZWt") = wRsRmDet!qIrRmIWt - (wTotIWt - wTotAWt)
'          End If
'        Else
'          ' Manoj 2.10.0
'          '.Value(wRmRow, "IrRmZWt") = wRsRmDet!qIrRmiWt + wRsRmDet!qXtraZWt
'          .Value(wRmRow, "IrRmIWt") = wRsRmDet!qIrRmIWt + _
'                IIF(UCase(wRsRmDet!qRmCtg) = "C" Or UCase(wRsRmDet!qRmCtg) = "D", Round(wRsRmDet!qXtraIWt, 2), wRsRmDet!qXtraIWt)
'
'          If whZWtEq = "A" Then
'            '.Value(wRmRow, "IrRmZWt") = wRsRmDet!qIrRmZWt + wRsRmDet!qXtraZWt
'            .Value(wRmRow, "IrRmZWt") = wRsRmDet!qIrRmZWt + _
'                  IIF(UCase(wRsRmDet!qRmCtg) = "C" Or UCase(wRsRmDet!qRmCtg) = "D", Round(wRsRmDet!qXtraZWt, 2), wRsRmDet!qXtraZWt)
'          ElseIf whZWtEq = "I" Then
'            '.Value(wRmRow, "IrRmZWt") = wRsRmDet!qIrRmIWt + wRsRmDet!qXtraIWt
'            .Value(wRmRow, "IrRmZWt") = wRsRmDet!qIrRmIWt + _
'                  IIF(UCase(wRsRmDet!qRmCtg) = "C" Or UCase(wRsRmDet!qRmCtg) = "D", Round(wRsRmDet!qXtraIWt, 2), wRsRmDet!qXtraIWt)
'          End If
'          ' Manoj 2.10.0
'        End If
'
'        wTotAWt = wTotAWt + MWLib.Div(.Value(wRmRow, "IrRmAWt"), wDivFactor)
'        wTotIWt = wTotIWt + MWLib.Div(.Value(wRmRow, "IrRmIWt"), wDivFactor)
'        wTotZWt = wTotZWt + MWLib.Div(.Value(wRmRow, "IrRmZWt"), wDivFactor)
'
'        If UCase(wRsRmDet!qRmCtg) = "G" And mIGldRtEqOrd <> "Y" Then
'          'Geeta***07/06/05 (Before 2.11)
'          '.Value(wRmRow, "IrRmIRt") = GetGldRt(GrdID.Value(GrdID.Row, "IdExpCmCd"), ADC("wIdLmgSal"), _
'                                      wRsRmDet!qIrRmCd, 0, "DM", GrdID.Value(GrdID.Row, "IdDmCd"), "S", "C")
'          .Value(wRmRow, "IrRmIRt") = GetGldRt(GrdID.Value(GrdID.Row, "IdExpCmCd"), ADC("wIdLmgSal"), _
'                                      wRsRmDet!qIrRmCd, .Value(wRmRow, "IrAlySalRt"), "DM", GrdID.Value(GrdID.Row, "IdDmCd"), "S", "C")
'          'Geeta***07/06/05 (Before 2.11)
'          If mLocalExp = "N" Then .Value(wRmRow, "IrRmZRt") = .Value(wRmRow, "IrRmIRt")
'        ElseIf UCase(wRsRmDet!qRmCtg) = "P" And mIGldRtEqOrd <> "Y" Then
'          'Geeta***07/06/05 (Before 2.11)
'          '.Value(wRmRow, "IrRmIRt") = GetGldRt(GrdID.Value(GrdID.Row, "IdExpCmCd"), ADC("wIdLmpSal"), _
'                                      wRsRmDet!qIrRmCd, 0, "DM", GrdID.Value(GrdID.Row, "IdDmCd"), "S", "C")
'          .Value(wRmRow, "IrRmIRt") = GetGldRt(GrdID.Value(GrdID.Row, "IdExpCmCd"), ADC("wIdLmpSal"), _
'                                      wRsRmDet!qIrRmCd, .Value(wRmRow, "IrAlySalRt"), "DM", GrdID.Value(GrdID.Row, "IdDmCd"), "S", "C")
'          'Geeta***07/06/05 (Before 2.11)
'          If mLocalExp = "N" Then .Value(wRmRow, "IrRmZRt") = .Value(wRmRow, "IrRmIRt")
'        '****** Sachin 2.12 - 06/12/05 - SJM
'        ElseIf UCase(wRsRmDet!qRmCtg) = "S" And mIGldRtEqOrd <> "Y" Then
'          .Value(wRmRow, "IrRmIRt") = GetGldRt(GrdID.Value(GrdID.Row, "IdExpCmCd"), ADC("wIdLmsSal"), _
'                                      wRsRmDet!qIrRmCd, .Value(wRmRow, "IrAlySalRt"), "DM", GrdID.Value(GrdID.Row, "IdDmCd"), "S", "C")
'          If mLocalExp = "N" Then .Value(wRmRow, "IrRmZRt") = .Value(wRmRow, "IrRmIRt")
'        '****** Sachin 2.12 - 06/12/05 - SJM
'        Else
'          .Value(wRmRow, "IrRmIRt") = wRsRmDet!qIrRmIRt
'          If mLocalExp = "N" Or (mLocalExp = "Y" And UCase(wRsRmDet!qRmCtg) <> "G" And UCase(wRsRmDet!qRmCtg) <> "P" And UCase(wRsRmDet!qRmCtg) <> "S") Then
'            .Value(wRmRow, "IrRmZRt") = wRsRmDet!qIrRmZRt
'          End If
'        End If
'        .Value(wRmRow, "IrRmZPureRt") = IIF(wRsRmDet!qRmPurityZ > 0, MWLib.Div(.Value(wRmRow, "IrRmZRt"), wRsRmDet!qRmPurityZ), .Value(wRmRow, "IrRmZRt"))
'      ElseIf UCase(ms_IdTcTyp) = UCase(ctCTBTcTyp) Then
'        .Value(wRmRow, "IrSrNo") = wRmRow
'        .Value(wRmRow, "IrRmCd") = wRsRmDet!qIrRmCd
'        ' Zubin 212
'        .Value(wRmRow, "IrLotNo") = wRsRmDet!qIrLotNo
'        ' Zubin 212
'        .Value(wRmRow, "IrRmSz") = wRsRmDet!qIrRmSz
'        .Value(wRmRow, "IrStkRt") = wRsRmDet!qIrStkRt
'        .Value(wRmRow, "IrRmQty") = wRsRmDet!qIrRmQty
'        .Value(wRmRow, "IrRmAWt") = wRsRmDet!qIrRmAWt
'
'        .Value(wRmRow, "IrRmIWt") = wRsRmDet!qIrRmIWt
'
'        .Value(wRmRow, "IrRmZWt") = wRsRmDet!qIrRmZWt
'        '.Value(wRmRow, "IrRmARt") = wRsRmDet!qIrRmARt
'        '.Value(wRmRow, "IrRmIRt") = wRsRmDet!qIrRmIRt
'        '.Value(wRmRow, "IrRmZRt") = wRsRmDet!qIrRmZRt
'        .Value(wRmRow, "IrSetSCd") = wRsRmDet!qIrSetSCd
'        .Value(wRmRow, "IrSetARt") = wRsRmDet!qIrSetARt
'        '.Value(wRmRow, "IrSetIRt") = wRsRmDet!qIrSetIRt
'        '.Value(wRmRow, "IrSetZRt") = wRsRmDet!qIrSetZRt
'        .Value(wRmRow, "IrRmZPureRt") = 0
'        .Value(wRmRow, "IrMainMet") = wRsRmDet!qIrMainMet
'      End If
'
'      ' Manoj 2.10.0
'      ' If UCase(wRsRmDet!qRmCtg) = "C" Or UCase(wRsRmDet!qRmCtg) = "D" Then
'      '   .Value(wRmRow, "IrRmAWt") = Round(.Value(wRmRow, "IrRmAWt"), 2)
'      '   .Value(wRmRow, "IrRmIWt") = Round(.Value(wRmRow, "IrRmIWt"), 2)
'      '   .Value(wRmRow, "IrRmZWt") = Round(.Value(wRmRow, "IrRmZWt"), 2)
'      ' End If
'      ' Manoj 2.10.0
'
'      .SaveRec (wRmRow)
'      wRsRmDet.MoveNext
'    Loop
'    .EndCopy
'  End With
'  '******Geeta***Emr208*************
'  '*********Zadd is Calculated in Make Lab ****
''  If UCase(ms_IdTcTyp) = UCase(ctInTcTyp) Then
''    wf_CnvRt = 1
''  ElseIf UCase(ms_IdTcTyp) = UCase(ctCTBTcTyp) Then
''    wf_CnvRt = moCn.GetFldVal("Select InCnvRt from InvHd where " + _
''                              "InCoCd= '" + gs_CoCd + "' and InTc= '" + adc("wIdTc") + "' " + _
''                              "and InYy= '" + adc("wIdYy") + "' and InChr= '" + adc("wIdChr") + "' " + _
''                              "and InNo= " + CStr(adc("wIdNo")))
''  End If
''
''  wf_AddLab = moCn.GetFldVal("Select sum(IrRmIVal- IrRmZVal* " + CStr(wf_CnvRt) + ") from InvRm " + _
'              "join RmMst on RmCd= IrRmCd where IrCoCd= '" + gs_CoCd + "' and " + _
'              "IrTc= '" + adc("wIdTc") + "' and IrYy= '" + adc("wIdYy") + "' and " + _
'              "IrChr= '" + adc("wIdChr") + "' and IrNo= " + CStr(adc("wIdNo")) + " and " + _
'              "IrSr= " + CStr(GrdID.Value(GrdID.Row, "IdSr")) + " and RmCtg in ('G', 'P')")
''
''  If mLocalExp = "Y" And Round(MWLib.Div(wf_AddLab, wf_CnvRt), 3) <> 0 Then
''    'And ms_IdTcTyp = ctInTcTyp
''    With GrdIL
''      .StartCopy
''      .AddItem
''      wLabRow = .Rows - 1
''      .Value(wLabRow, "iLSrNo") = wLabRow
''      .Value(wLabRow, "iLMCd") = ctAddLMCd
''      .Value(wLabRow, "iLSCd") = ctAddLSCd
''      .Value(wLabRow, "iLQty") = 1
''      .Value(wLabRow, "iLARt") = 0
''      .Value(wLabRow, "iLiRt") = 0
''      .Value(wLabRow, "iLZRt") = MWLib.Div(wf_AddLab, wf_CnvRt)
''      .Value(wLabRow, "ilQw") = "Q"
''      .Value(wLabRow, "iLZVal") = MWLib.Div(wf_AddLab, wf_CnvRt)
''      .SaveRec (wLabRow)
''      .EndCopy
''    End With
''  End If
''
'  Set wRsRmDet = Nothing
'  Set wOrRecSet = Nothing
'  '*** (Bef 31/01/06)
End Sub
Public Sub MakeRmFromFg(ByVal mIWtEqOrd As String, ByVal mIWtFrOrd As String, _
                        ByVal mIGldRtEqOrd As String, ByVal mIdQty As Double)
'SRITODO
  Dim whZWtEq As String, wCtgStr As String, wDmCtg As String, wTmpTbl As String
  Dim wOrStr As String, wStr As String, wMainMet As Boolean
  Dim wRsRmDet As MDORowSet, wOrRecSet As MDORowSet
  Dim wDivFactor As Single, wRmRow As Single, wLabRow As Single, wf_AddLab As Single
  Dim wTotIXtraWt As Double, wTotZXtraWt As Double, wTotAWt As Double
  Dim wTotIWt As Double, wTotZWt As Double
  Dim wCmDiaRtFA As String, wCmCsRtFA As String, wf_CnvRt As Single
  Dim wDRrTcTyp As String, wCRrTcTyp As String, wDFrOrdYN As String
  Dim wCFrOrdYN As String
  Dim wSetRt As Boolean
  Dim wTotWtDiff As Double, wWtDiff As Double, wTotIrRmWt As Double, wAddWt As Double
  Dim wIrRmWt As MDORowSet, wRmCtg As String, wFrOrdYN As String, wMainMetWt As Double
  Dim wFrRmWt As Double, wFrRmQty As Double, wOrTotWt As Double, wOrTotQty As Double
  Dim wFRmDet As MDORowSet
  whZWtEq = moCn.GetFldVal("Select hZWtEq from Head where HCoCd= '" + gs_CoCd + "' and HCd='" + ctSelfCmCd + "'")
  Dim wSetTyp As Boolean
  Dim wOrLnStr As String
  Dim wOrLnStr2 As String
  Dim wOrLnStr3 As String
  
  
  
  '****** Sachin 3.3.0 (24-06-08) - [14. Cost Markup In Invoicing] ******
  Dim wTotMrkUpVal As Single, wCurrMrkUpRt As Single, wMrkUpLmt As Single, wFnlIncrVal As Single
  Dim wMrkUpVal As Single, wRmQtyWt As Single
  '****** Sachin 3.3.0 (24-06-08) - [14. Cost Markup In Invoicing] ******
  
  '*** (Jen 2.12)
  Dim wRs_RmCd As MDORowSet, wb_WtFrOrd As Boolean
  '*** (Jen 2.12)
  
  Dim wf_CustmCdPurityZ As Double     '*** (Jen 2.13)
  
  'mIWtFrOrd = "DC"
  'wCtgStr= ",D,C"
  wCtgStr = "," + moCn.GetFldVal("Select PDesc225 from Param where " + _
              "PTyp='IWTFRORD' and PMCd = '" + mIWtFrOrd + "'") + ","
  
  wCmDiaRtFA = moCn.GetFldVal("Select CmDiaRtFA from CustMst where CmCtg= 'C' and " + _
               "CmCd= '" + GrdID.Value(GrdID.Row, "IdExpCmCd") + "' ")
  
  wCmCsRtFA = moCn.GetFldVal("Select CmCSRtFA from CustMst where CmCtg= 'C' and " + _
               "CmCd= '" + GrdID.Value(GrdID.Row, "IdExpCmCd") + "' ")
  
  wDmCtg = moCn.GetFldVal("Select DmCtg from DsgMst where DmTcTyp= 'DM' and DmSz= '' and " + _
               "DmCd= '" + GrdID.Value(GrdID.Row, "IdDmCd") + "' and DmPrtKey='" + ctCurrPrtn + "' ")
  wTmpTbl = "#TmpInvRm"

  wMrkUpLmt = moCn.GetFldVal(" Select Top 1 RrSalRt From RmRt Where RrTcTyp='IML' And RrCmCtg='C' And RrCmCd In ('" + GrdID.Value(GrdID.Row, "IdExpCmCd") + "','" + ctSelfCmCd + "') Order By (Case When RrCmCd='" + GrdID.Value(GrdID.Row, "IdExpCmCd") + "' Then 1 Else 2 End)")
        Dim wOmLkUpRmRt As String, wOmLkUpMetLs As String, wOmLkUpStr As String, wOmLkUpLabRt As String
        Dim wOmLkUpArr() As String
        
        wOmLkUpStr = moCn.GetFldVal("Select OmLkUpRmRt+','+OmLkUpMetLs+','+OmLkUpLabRt from OrdMst where " + _
                                    "OmCoCd= '" + gs_CoCd + "' and OmTc= '" + GrdID.Value(GrdID.Row, "IdExpOdTc") + "' " + _
                                    "and OmYy= '" + GrdID.Value(GrdID.Row, "IdExpOdYy") + "' " + _
                                    "and OmChr= '" + GrdID.Value(GrdID.Row, "IdExpOdChr") + "' " + _
                                    "and OmNo= " + CStr(GrdID.Value(GrdID.Row, "IdExpOdNo")) + " and OmPrtKey='" + ctCurrPrtn + "' ")
        If wOmLkUpStr <> "" Then
          wOmLkUpArr() = Split(wOmLkUpStr, ",")
          wOmLkUpRmRt = wOmLkUpArr(0)
          wOmLkUpMetLs = wOmLkUpArr(1)
          wOmLkUpLabRt = wOmLkUpArr(2)
        End If
        ' ***** Manali 3.03 - 05/07/08
  
  '*******Geeta*******Emr207
  Dim ws_OrLn As String
  ws_OrLn = "(case when Pvalue3='Y' then (Case when PValue2='Y' then  " + _
          "cast(OrRmPtr as Decimal(12,4)) else (case when OrPrdQty=0 then 0 else  " + _
          "cast((OrPrdWt/OrPrdQty) as Decimal(12,4)) end ) end )  " + _
          "else cast(Orln1  as Decimal(12,4)) end ) "
  Dim ws_OrLn2 As String
  ws_OrLn2 = "(case when Pvalue3='Y' then (Case when PValue2='Y' then  " + _
          "cast(OrRmPtr as Decimal(12,4)) else (case when OrPrdQty=0 then 0 else  " + _
          "cast((OrPrdWt/OrPrdQty) as Decimal(12,4)) end ) end )  " + _
          "else cast(Orln2  as Decimal(12,4)) end ) "
  Dim ws_OrLn3 As String
  ws_OrLn3 = "(case when Pvalue3='Y' then (Case when PValue2='Y' then  " + _
          "cast(OrRmPtr as Decimal(12,4)) else (case when OrPrdQty=0 then 0 else  " + _
          "cast((OrPrdWt/OrPrdQty) as Decimal(12,4)) end ) end )  " + _
          "else cast(Orln3  as Decimal(12,4)) end ) "

  If UCase(ms_IdTcTyp) = UCase(ctInTcTyp) Then
    wStr = "Select max(RmMst.RmQw) as RmQw, max(RmMst.RmCmCd) as RmCmCd, max(RmMst.RmZCd) as RmZCd, max(RmMst.RmCtg) as RmCtg, (case when max(isnull(qOrRmCd, ''))= '' then '' else 'Y' end) as qOrRmExists, " + _
           "FrRmCd as qIrRmCd, FrLotNo as qIrLotNo, FrRmSz as qIrRmSz,FrRmSz2 as qIrRmSz2, FrRmSz3 as qIrRmSz3,sum(FrRmStkRt*FrRmWt)/(case when sum(FrRmWt)> 0  " + _
           "then sum(FrRmWt) else 1 end) as qIrStkRt, " + _
           "sum(FrRmQty) as qIrRmQty, sum(FrRmWt) As qIrRmAWt, sum(FrRmWt) as qIrRmIWt, " + _
           "sum(FrRmWt) as qIrRmZWt,max(Isnull(qOrWt,0)) as qOrWt,max(IsNull(qOrQty,0)) as qOrQty, 9999.999 as qXtraIWt, 9999.999 as qXtraZWt, " + _
           "999999.999 as qIrRmARt, 999999.999 as qIrRmIRt, 999999.999 as qIrRmZRt, IsNull(qOrSetSCd, '') as qIrSetSCd, " + _
           "isnull(qOrSetSalRt, 0.00) as qIrSetARt, isnull(qOrSetSalRt, 0.00) as qIrSetIRt, isnull(qOrSetSalRt, 0.00) as qIrSetZRt, " + _
           "Space(1) as qIrMainMet, max(RmPurityZ) as qRmPurityZ, max(RmCtg) as qRmCtg, max(RmSCtg) as qRmSCtg,Max(OrCustRmCd) as OrCustRmCd  " + _
           "into " + wTmpTbl + " from InvFgd join FgRm on FrFdIdNo=IfFdIdNo and FrCoCd= IfCoCd and FrTc= IfFdTc and FrYy= IfFdyy and " + _
           "FrChr= IfFdChr and FrNo = IfFdNo And FrSr = IfFdSr And FrPrtKey=IfPrtKey " + _
           "Left Outer Join (select OrRmCd as qOrRmCd, " + ws_OrLn + " as qOrLn1, " + "0" + " as qOrLn2," + "0" + " as qOrLn3,OrSetSCd as qOrSetSCd, " + _
           "OrSetSalRt as qOrSetSalRt, sum(OrWt) as qOrWt, sum(OrQty) as qOrQty,Max(OrCustRmCd) as OrCustRmCd  from OrdRm  " + _
           "Join RmMst on RmCd= OrRmCd and RmPrtKey=OrPrtKey " + _
           "Join Param on PTyp= 'RMSCTG' and PMCd= RmCtg and PSCd= RmSCtg " + _
              "where OrCoCd= '" + gs_CoCd + "'  and OrTc = '" + GrdID.Value(GrdID.Row, "IdExpOdTc") + "' " + _
           "and OrYy= '" + GrdID.Value(GrdID.Row, "IdExpOdYy") + "' " + _
           " and OrChr= '" + GrdID.Value(GrdID.Row, "IdExpOdChr") + "' and OrNo = " + CStr(GrdID.Value(GrdID.Row, "IdExpOdNo")) + _
           " and OrSr= " + CStr(GrdID.Value(GrdID.Row, "IdExpOdSr")) + " and OrPrtKey='" + ctCurrPrtn + "' Group By " + _
           "OrRmCd, " + ws_OrLn + ", OrSetSCd, OrSetSalRt) Or2 On qOrRmCd= FrRmCd and qOrLn1= FrRmSz /*and qOrLn2 = FrRmSz2 and qOrLn3 = qOrLn3*/ " + _
           "Join RmMst on FrRmCd = RmCd and FrPrtKey=RmPrtKey " + _
           "where IfCoCd = '" + gs_CoCd + "' and IfTc= '" + ADC("wIdTc") + "' and IfYy= '" + _
           ADC("wIdYy") + "' and IfChr= '" + ADC("wIdChr") + "' and " + _
           "IfNo = " + CStr(ADC("wIdNo")) + " and IfSr = " + CStr(GrdID.Value(GrdID.Row, "IdSr")) + " and IfPrtKey='" + ctCurrPrtn + "' " + _
           "Group by FrRmCd, FrLotNo, FrRmSz, FrRmSz2, FrRmSz3, qOrSetSCd, qOrSetSalRt "
  End If
  Dim wRsInvHd As MDORowSet
  Set wRsInvHd = moCn.OpenRes("Select * From InvHd Where InCoCd='" + gs_CoCd + "' And InTc='" + ADC("wIdTc") + "' And InYy='" + ADC("wIdYy") + "' " + _
                          " And InChr='" + ADC("wIdChr") + "' And InNo=" + CStr(ADC("wIdNo")) + " And InPrtKey='" + ADC("wIdPrtKey") + "'")
  Dim wRsOm As MDORowSet
  Set wRsOm = moCn.OpenRes("Select * from OrdMst where " + _
                          "OmCoCd= '" + gs_CoCd + "' and OmTc= '" + GrdID.Value(GrdID.Row, "IdExpOdTc") + "' " + _
                          "and OmYy= '" + GrdID.Value(GrdID.Row, "IdExpOdYy") + "' " + _
                          "and OmChr= '" + GrdID.Value(GrdID.Row, "IdExpOdChr") + "' " + _
                          "and OmNo= " + CStr(GrdID.Value(GrdID.Row, "IdExpOdNo")) + " and OmPrtKey='" + ctCurrPrtn + "' ")
  Dim wBaseCurCd As String
  wBaseCurCd = moCn.GetFldVal("Select HBaseCurCd from Head Where HCoCd='" + ctSelfCoCd + "' and HCd='" + ctSelfCmCd + "'")

  moCn.CreateTmpTable wTmpTbl, wStr
          wDmCtg = moCn.GetFldVal("Select DmCtg from DsgMst where DmTcTyp= 'DM' and DmSz= '' and " + _
                           "DmCd= '" + GrdID.Value(GrdID.Row, "IdDmCd") + "' and DmPrtKey='" + ctCurrPrtn + "' ")

  Set wRsRmDet = moCn.OpenResultset("Select * from " + wTmpTbl)
  '**********Geeta******Emr207
  '*** (Bef 2.11)
  'If moCn.RecSeek("select count(distinct qIrSetIRt) From " + wTmpTbl + " " + _
  '                "group by qIrRmCd,qIrRmSz Having count(distinct qIrSetIRt)>1") Then
  '*** (Bef 2.11)
  
  '*** 2.11 (Jen)
  ' Zubin 212 (introduced Lot No)
  If moCn.RecSeek("select count(distinct Str(qIrSetIRt, 16, 2) + qIrSetSCd) From " + wTmpTbl + " " + _
                  "group by qIrRmCd, qIrLotNo, qIrRmSz,qIrRmSz2,qIrRmSz3 Having count(distinct Str(qIrSetIRt, 16, 2) + qIrSetSCd)> 1") Then
  '*** 2.11 (Jen)
    '*** To get Proportionate weights
    '****** Sachin 3.02 31-12-07 - PrtKey added
    Set wFRmDet = moCn.OpenRes("Select FrRmCd as qFrRmCd, FrLotNo as qFrLotNo, max(FrRmSz) as qFrRmsz,max(FrRmSz2) as qFrRmsz2,max(FrRmSz3) as qFrRmsz3, " + _
                       " Sum(FrRmWt) as qFrRmWt, Sum(FrRmQty) as qFrRmQty from FgRm  " + _
                       "join InvFgd on FrFdIdNo=IfFdIdNo and FrCoCd= IfCoCd and FrTc= IfFdTc and FrYy= IfFdYy and  " + _
                       " FrChr= IfFdChr and FrNo = IfFdNo And FrSr = IfFdSr And FrPrtKey=IfPrtKey " + _
                       "where IfCoCd = '" + gs_CoCd + "' and IfTc= '" + ADC("wIdTc") + "' and IfYy= '" + _
                       ADC("wIdYy") + "' and IfChr= '" + ADC("wIdChr") + "' and " + _
                       "IfNo = " + CStr(ADC("wIdNo")) + " and IfSr = " + CStr(GrdID.Value(GrdID.Row, "IdSr")) + " and IfPrtKey='" + ctCurrPrtn + "' " + _
                       "Group By FrRmCd, FrLotNo, FrRmSz,FrRmSz2, FrRmSz3 Order By FrRmCd, FrLotNo, FrRmSz, FrRmSz2, FrRmSz3")
    
    Do While Not (wFRmDet.EOF Or wFRmDet.BOF)
      wFrRmWt = wFRmDet!qFrRmWt
      wFrRmQty = wFRmDet!qFrRmQty
      ' Zubin 212 (introduced Lot No)
      wOrTotWt = moCn.GetFldVal("Select Sum(qOrWt) From " + wTmpTbl + " Where qIrRmCd= '" + wFRmDet!qFrRmCd + "' and qIrRmSz= '" + CStr(wFRmDet!qFrRmSz) + "' and qIrRmSz2= '" + CStr(wFRmDet!qFrRmSz2) + "' and qIrRmSz3= '" + CStr(wFRmDet!qFrRmSz3) + "' and qIrLotNo= '" + CStr(wFRmDet!qFrLotNo) + "' Group By qIrRmcd, qIrLotNo, qIrRmSz,qIrRmSz2, qIrRmSz3")
      wOrTotQty = moCn.GetFldVal("Select Sum(qOrQty) From " + wTmpTbl + "  Where qIrRmCd= '" + wFRmDet!qFrRmCd + "' and qIrRmSz= '" + CStr(wFRmDet!qFrRmSz) + "' and qIrRmSz2= '" + CStr(wFRmDet!qFrRmSz2) + "' and qIrRmSz3= '" + CStr(wFRmDet!qFrRmSz3) + "' and qIrLotNo= '" + CStr(wFRmDet!qFrLotNo) + "' Group By qIrRmCd, qIrLotNo, qIrRmSz,qIrRmSz2, qIrRmSz3")
      ' Zubin 212
      
      If (wRsRmDet.RecCount > 0) Then wRsRmDet.MoveFirst
      Do While Not (wRsRmDet.EOF Or wRsRmDet.BOF)
        ' Zubin 212 (introduced Lot No)
        If wFRmDet!qFrRmCd = wRsRmDet!qirrmcd And wFRmDet!qFrLotNo = wRsRmDet!qIrLotNo And wFRmDet!qFrRmSz = wRsRmDet!qIrRmSz And wFRmDet!qFrRmSz2 = wRsRmDet!qIrRmSz2 And wFRmDet!qFrRmSz3 = wRsRmDet!qIrRmSz3 Then
          ' Zubin 212 (introduced Lot No)
          If moCn.RecSeek("select count(distinct Str(qIrSetIRt, 16, 2) + qIrSetSCd) From " + wTmpTbl + " " + _
                          "Where qIrRmCd= '" + wRsRmDet!qirrmcd + "' " + _
                          " and qIrLotNo= '" + CStr(wRsRmDet!qIrLotNo) + "' " + _
                          " and qIrRmSz= " + CStr(wRsRmDet!qIrRmSz) + _
                          " and qIrRmSz2= " + CStr(wRsRmDet!qIrRmSz2) + _
                          " and qIrRmSz3= " + CStr(wRsRmDet!qIrRmSz3) + _
                          "group by qIrRmCd, qIrLotNo, qIrRmSz, qIrRmSz2, qIrRmSz3 Having count(distinct Str(qIrSetIRt, 16, 2) + qIrSetSCd)> 1") Then
                          
            wRsRmDet.edit
            wRsRmDet!qIrRmAWt = Round(MWLib.Div(wRsRmDet!qOrWt, wOrTotWt) * wFrRmWt, 3)
            wRsRmDet!qIrRmIWt = wRsRmDet!qIrRmAWt
            wRsRmDet!qIrRmZWt = wRsRmDet!qIrRmAWt
            wOrTotWt = wOrTotWt - wRsRmDet!qOrWt
            wFrRmWt = wFrRmWt - wRsRmDet!qIrRmAWt
            wRsRmDet!qIrRmQty = Round(MWLib.Div(wRsRmDet!qOrQty, wOrTotQty) * wFrRmQty)
            wOrTotQty = wOrTotQty - wRsRmDet!qOrQty
            wFrRmQty = wFrRmQty - wRsRmDet!qIrRmQty
            wRsRmDet.Update
          End If
        End If
        wRsRmDet.MoveNext
      Loop
      wFRmDet.MoveNext
    Loop
    Set wFRmDet = Nothing
  End If
    
  '*** (Jen 2.12) (31/01/06)
  If mIWtEqOrd = "Y" Or mIWtFrOrd <> "" Then
    '*** To get Inv Weights as per Order
    '****** Sachin 3.02 31-12-07 - PrtKey added
    Set wRs_RmCd = moCn.OpenRes("Select IsNull((Select min(qOrRmExists) From " + wTmpTbl + " where qIrRmCd= qOrRmCd), 'Y') as qOrdExists, " + _
                    "min(qFgExists) as qFgExists, sum(qOrRmWt)* " + CStr(mIdQty) + " as qOrRmWt, " + _
                    "sum(qOrRmQty)* " + CStr(mIdQty) + " as qOrRmQty, (Select sum(FrRmWt) from FgRm  " + _
                    "   join InvFgd on FrFdIdNo=IfFdIdNo and FrCoCd= IfCoCd and FrTc= IfFdTc and FrYy= IfFdYy and  " + _
                    "   FrChr= IfFdChr and FrNo = IfFdNo And FrSr = IfFdSr And FrPrtKey=IfPrtKey " + _
                    "   where IfCoCd = '" + gs_CoCd + "' and IfTc= '" + ADC("wIdTc") + "' and " + _
                    "   IfYy= '" + ADC("wIdYy") + "' and IfChr= '" + ADC("wIdChr") + "' and " + _
                    "   IfNo = " + CStr(ADC("wIdNo")) + " and IfSr = " + CStr(GrdID.Value(GrdID.Row, "IdSr")) + " and IfPrtKey='" + ctCurrPrtn + "' " + _
                    "   and FrRmCd= qOrRmCd) as qFgRmWt, qOrRmCtg, qOrRmCd " + _
                    "From (Select OrRmCtg as qOrRmCtg, OrRmCd as qOrRmCd, " + _
                    " " + ws_OrLn + " as qOrRmSz," + ws_OrLn2 + " as qOrRmSz2," + ws_OrLn3 + " as qOrRmSz3, OrWt as qOrRmWt, OrQty as qOrRmQty, " + _
                    "(case when not exists(Select 'x' from " + wTmpTbl + " where qIrRmCd= OrRmCd and qIrRmSz= " + ws_OrLn + ") then '' else 'Y' end) as qFgExists " + _
                    "from OrdRm Join RmMst on RmCd= OrRmCd and RmPrtKey=OrPrtKey " + _
                    "Join Param on PTyp= 'RMSCTG' and PMCd= RmCtg and PSCd= RmSCtg " + _
                    "where OrCoCd= '" + gs_CoCd + "' and OrTc= '" + GrdID.Value(GrdID.Row, "IdExpOdTc") + "' " + _
                    "and OrYy= '" + GrdID.Value(GrdID.Row, "IdExpOdYy") + "' " + _
                    " and OrChr= '" + GrdID.Value(GrdID.Row, "IdExpOdChr") + "' and OrNo = " + CStr(GrdID.Value(GrdID.Row, "IdExpOdNo")) + _
                    " and OrSr= " + CStr(GrdID.Value(GrdID.Row, "IdExpOdSr")) + " and OrPrtKey='" + ctCurrPrtn + "' and OrSetSCd<> '') Ord " + _
                    " Group By qOrRmCtg, qOrRmCd ")

    
    Do While Not (wRs_RmCd.EOF Or wRs_RmCd.BOF)
      wFrOrdYN = Trim(mIWtEqOrd)
      If wFrOrdYN = "N" Then wFrOrdYN = IIF(InStr(1, wCtgStr, "," + wRs_RmCd!qOrRmCtg + ",") <> 0, "Y", "N")
      ' Debug.Print "OrdExists: " & wRs_RmCd!qOrdExists & "   FgExists: " & wRs_RmCd!qFgExists & "   WtFrOrd: " & wFrOrdYN & "   OrWt: " & wRs_RmCd!qOrRmWt & "   FgWt: " & wRs_RmCd!qFgRmWt
      '*** (Bef 02/02/06)
      'If wRs_RmCd!qOrdExists = "Y" And wRs_RmCd!qFgExists = "Y" And wFrOrdYN = "Y" And wRs_RmCd!qOrRmWt = wRs_RmCd!qFgRmWt Then
      '*** (Bef 02/02/06)
      If wRs_RmCd!qOrdExists = "Y" And wRs_RmCd!qFgExists = "Y" And wFrOrdYN = "Y" Then
        wb_WtFrOrd = True
      Else
        wb_WtFrOrd = False
      End If
      
      If (wRsRmDet.RecCount > 0) Then wRsRmDet.MoveFirst
      Do While Not (wRsRmDet.EOF Or wRsRmDet.BOF)
        If wRs_RmCd!qOrRmCd = wRsRmDet!qirrmcd And wb_WtFrOrd = True Then
            wRsRmDet.edit
            wRsRmDet!qIrRmIWt = wRsRmDet!qOrWt * mIdQty
            If Trim(whZWtEq) = "I" Then
              wRsRmDet!qIrRmZWt = wRsRmDet!qIrRmIWt
            End If
            wRsRmDet.Update
        End If
        wRsRmDet.MoveNext
      Loop
      wRs_RmCd.MoveNext
    Loop
    Set wRs_RmCd = Nothing
  End If
  '*** (Jen 2.12) (31/01/06)
                    

  'If (wRsRmDet.RecCount > 0) Then wRsRmDet.MoveFirst
  'Do While Not (wRsRmDet.EOF Or wRsRmDet.BOF)
  '  Debug.Print CStr(wRsRmDet!qIrRmCd) + "," + CStr(wRsRmDet!qIrRmSz) + "," + CStr(wRsRmDet!qIrRmQty) + "," + CStr(wRsRmDet!qIrRmAWt) + "," + CStr(wRsRmDet!qOrWt) + "," + CStr(wRsRmDet!qOrQty)
  '  wRsRmDet.MoveNext
  'Loop
  
  '**********Geeta******Emr207
  If (wRsRmDet.RecCount > 0) Then wRsRmDet.MoveFirst
  Do While Not (wRsRmDet.EOF Or wRsRmDet.BOF)
    wRsRmDet.edit
             
    '****** Sachin 3.02 31-12-07 - PrtKey added
    wOrLnStr = IIF(moCn.GetFldVal("select PValue3 From Param Join  RmMst " + _
                            " on PTyp= 'RMSCTG' and PMCd= RmCtg and PSCd= RmSCtg " + _
                            "where RmCd= '" + wRsRmDet!qirrmcd + "' and RmPrtKey='" + ctCurrPrtn + "' ") = "Y", _
                   IIF(moCn.GetFldVal("select PValue2 From Param Join RmMst " + _
                            "on PTyp= 'RMSCTG' and PMCd= RmCtg and PSCd= RmSCtg " + _
                            "where RmCd= '" + wRsRmDet!qirrmcd + "' and RmPrtKey='" + ctCurrPrtn + "' ") = "Y", "OrRmPtr", _
                            " case when OrPrdQty =0 then 0 else OrPrdWt/OrPrdQty end "), "OrLn1")
    wOrLnStr2 = IIF(moCn.GetFldVal("select PValue3 From Param Join  RmMst " + _
                            " on PTyp= 'RMSCTG' and PMCd= RmCtg and PSCd= RmSCtg " + _
                            "where RmCd= '" + wRsRmDet!qirrmcd + "' and RmPrtKey='" + ctCurrPrtn + "' ") = "Y", _
                   IIF(moCn.GetFldVal("select PValue2 From Param Join RmMst " + _
                            "on PTyp= 'RMSCTG' and PMCd= RmCtg and PSCd= RmSCtg " + _
                            "where RmCd= '" + wRsRmDet!qirrmcd + "' and RmPrtKey='" + ctCurrPrtn + "' ") = "Y", "OrRmPtr", _
                            " case when OrPrdQty =0 then 0 else OrPrdWt/OrPrdQty end "), "OrLn2")
    wOrLnStr3 = IIF(moCn.GetFldVal("select PValue3 From Param Join  RmMst " + _
                            " on PTyp= 'RMSCTG' and PMCd= RmCtg and PSCd= RmSCtg " + _
                            "where RmCd= '" + wRsRmDet!qirrmcd + "' and RmPrtKey='" + ctCurrPrtn + "' ") = "Y", _
                   IIF(moCn.GetFldVal("select PValue2 From Param Join RmMst " + _
                            "on PTyp= 'RMSCTG' and PMCd= RmCtg and PSCd= RmSCtg " + _
                            "where RmCd= '" + wRsRmDet!qirrmcd + "' and RmPrtKey='" + ctCurrPrtn + "' ") = "Y", "OrRmPtr", _
                            " case when OrPrdQty =0 then 0 else OrPrdWt/OrPrdQty end "), "OrLn3")
                            
                            
    '*** (Bef 2.11)
    'If moCn.RecSeek("select count(distinct qIrSetIRt) From " + wTmpTbl + " " + _
                    "Where qIrRmCd= '" + wRsRmDet!qIrRmCd + "' " + _
                    " and qIrRmSz= " + CStr(wRsRmDet!qIrRmSz) + _
                    "group by qIrRmCd, qIrRmSz Having count(distinct qIrSetIRt)> 1") Then
    '*** (Bef 2.11)
    
    '*** 2.11 (Jen)
    ' Zubin 212
    If moCn.RecSeek("select count(distinct Str(qIrSetIRt, 16, 2) + qIrSetSCd) From " + wTmpTbl + " " + _
                    "Where qIrRmCd= '" + wRsRmDet!qirrmcd + "' " + _
                    " and qIrLotNo= '" + CStr(wRsRmDet!qIrLotNo) + "' " + _
                    " and qIrRmSz= " + CStr(wRsRmDet!qIrRmSz) + _
                    " and qIrRmSz2= " + CStr(wRsRmDet!qIrRmSz2) + _
                    " and qIrRmSz3= " + CStr(wRsRmDet!qIrRmSz3) + _
                    "group by qIrRmCd, qIrLotNo, qIrRmSz, qIrRmSz2, qIrRmSz3 Having count(distinct Str(qIrSetIRt, 16, 2) + qIrSetSCd)> 1") Then
    '*** 2.11 (Jen)
      '*******Geeta*******Emr207
      '****** Sachin 3.02 31-12-07 - PrtKey Added
      wOrStr = "Select OrLn1,OrLn2,(Case When OrQty>0 Then OrWt/OrQty Else 0 End) as OrRmPtr, OrCstRt, OrSalRt, OrSetSCd, OrSetCstRt, OrSetSalRt, OrMainMet, OrRmCtg,OrCustRmCd as OrCustRmCd  " + _
              "From OrdRm where OrCoCd= '" + gs_CoCd + "' " + _
              " and OrTc = '" + GrdID.Value(GrdID.Row, "IdExpOdTc") + "' " + _
              " and OrYy= '" + GrdID.Value(GrdID.Row, "IdExpOdYy") + "' " + _
              " and OrChr= '" + GrdID.Value(GrdID.Row, "IdExpOdChr") + "' " + _
              " and OrNo = " + CStr(GrdID.Value(GrdID.Row, "IdExpOdNo")) + _
              " and OrSr= " + CStr(GrdID.Value(GrdID.Row, "IdExpOdSr")) + _
              " and OrPrtKey= '" + ctCurrPrtn + "' " + _
              " and OrRmCd= '" + wRsRmDet!qirrmcd + "' " + _
              " and " + wOrLnStr + " = " + CStr(wRsRmDet!qIrRmSz) + _
              " and OrSetSCd= '" + wRsRmDet!qIrSetSCd + "' " + _
              " and OrSetSalRt= " + CStr(wRsRmDet!qIrSetIRt) + _
              " Order by OrMainMet Desc, OrSalRt Desc"
               
'sri " and " + wOrLnStr2 + " = " + CStr(wRsRmDet!qIrRmSz2) + _
              " and " + wOrLnStr3 + " = " + CStr(wRsRmDet!qIrRmSz3) + _


      'Set wOrRecSet = moCn.OpenRes(wOrStr)
    Else
      '****** Sachin 3.02 31-12-07 - PrtKey Added
      wOrStr = "Select OrLn1,OrLn2,(Case When OrQty>0 Then OrWt/OrQty Else 0 End) as OrRmPtr, OrCstRt, OrSalRt, OrSetSCd, OrSetCstRt, OrSetSalRt, OrMainMet, OrRmCtg,OrCustRmCd as OrCustRmCd " + _
                "From OrdRm where OrCoCd= '" + gs_CoCd + "' " + _
                " and OrTc = '" + GrdID.Value(GrdID.Row, "IdExpOdTc") + "' " + _
                " and OrYy= '" + GrdID.Value(GrdID.Row, "IdExpOdYy") + "' " + _
                " and OrChr= '" + GrdID.Value(GrdID.Row, "IdExpOdChr") + "' " + _
                " and OrNo = " + CStr(GrdID.Value(GrdID.Row, "IdExpOdNo")) + _
                " and OrSr= " + CStr(GrdID.Value(GrdID.Row, "IdExpOdSr")) + _
                " and OrPrtKey='" + ctCurrPrtn + "' " + _
                " and OrRmCd= '" + wRsRmDet!qirrmcd + "' " + _
                " and " + wOrLnStr + " = " + CStr(wRsRmDet!qIrRmSz) + _
                " Order by OrMainMet Desc, OrSalRt Desc"
'sri                + _
                " and " + wOrLnStr2 + " = " + CStr(wRsRmDet!qIrRmSz2) + _
                " and " + wOrLnStr3 + " = " + CStr(wRsRmDet!qIrRmSz3) + _

                
    End If
    Set wOrRecSet = moCn.OpenRes(wOrStr)
    
    If (wOrRecSet.EOF Or wOrRecSet.BOF) Then
      '****** Sachin 3.02 31-12-07 - PrtKey Added
      wOrStr = "Select OrLn1,OrLn2,(Case When OrQty>0 Then OrWt/OrQty Else 0 End) as OrRmPtr, OrCstRt, OrSalRt, OrSetSCd, OrSetCstRt, OrSetSalRt, OrMainMet, OrRmCtg,OrCustRmCd as OrCustRmCd " + _
               "From OrdRm join RmRt on RrTcTyp= 'Rsz' and RrCmCd= '" + ctSelfCmCd + "' and " + _
               "RrCtg= OrRmCtg and RrSCtg= OrRmSCtg " + _
               "where OrCoCd= '" + gs_CoCd + "' " + _
               " and OrTc = '" + GrdID.Value(GrdID.Row, "IdExpOdTc") + "' " + _
               " and OrYy= '" + GrdID.Value(GrdID.Row, "IdExpOdYy") + "' " + _
               " and OrChr = '" + GrdID.Value(GrdID.Row, "IdExpOdChr") + "' " + _
               " and OrNo = " + CStr(GrdID.Value(GrdID.Row, "IdExpOdNo")) + _
               " and OrSr= " + CStr(GrdID.Value(GrdID.Row, "IdExpOdSr")) + _
               " and OrPrtkey='" + ctCurrPrtn + "' " + _
               " and OrRmCd =  '" + wRsRmDet!qirrmcd + "' " + _
               " and " + wOrLnStr + " >= RrFrLn and " + wOrLnStr + " <= RrToLn and " + _
               CStr(wRsRmDet!qIrRmSz) + " >= RrFrLn and " + _
               CStr(wRsRmDet!qIrRmSz) + " <= RrToLn " + _
               "Order by OrMainMet Desc, OrSalRt Desc "

'SRI
'                              " and " + wOrLnStr2 + " >= RrFrLn and " + wOrLnStr2 + " <= RrToLn and " + _
               CStr(wRsRmDet!qIrRmSz2) + " >= RrFrLn and " + _
               CStr(wRsRmDet!qIrRmSz2) + " <= RrToLn " + _
               " and " + wOrLnStr3 + " >= RrFrLn and " + wOrLnStr3 + " <= RrToLn and " + _
               CStr(wRsRmDet!qIrRmSz3) + " >= RrFrLn and " + _
               CStr(wRsRmDet!qIrRmSz3) + " <= RrToLn " + _

               
      Set wOrRecSet = moCn.OpenRes(wOrStr)
    End If
    If (wOrRecSet.EOF Or wOrRecSet.BOF) Then
      '****** Sachin 3.02 31-12-07 - PrtKey Added
      wOrStr = "Select OrLn1,OrLn2,(Case When OrQty>0 Then OrWt/OrQty Else 0 End) as OrRmPtr, OrCstRt, OrSalRt, OrSetSCd, OrSetCstRt, OrSetSalRt, OrMainMet, OrRmCtg,OrCustRmCd as OrCustRmCd " + _
               "From OrdRm where OrCoCd= '" + gs_CoCd + "' " + _
               "  and OrTc = '" + GrdID.Value(GrdID.Row, "IdExpOdTc") + "' " + _
               " and OrYy= '" + GrdID.Value(GrdID.Row, "IdExpOdYy") + "' " + _
               " and OrChr= '" + GrdID.Value(GrdID.Row, "IdExpOdChr") + "' " + _
               " and OrNo = " + CStr(GrdID.Value(GrdID.Row, "IdExpOdNo")) + _
               " and OrSr= " + CStr(GrdID.Value(GrdID.Row, "IdExpOdSr")) + _
               " and OrPrtKey='" + ctCurrPrtn + "' " + _
               " and OrRmCd= '" + wRsRmDet!qirrmcd + "' " + _
               "Order by OrMainMet Desc, OrSalRt Desc"
      Set wOrRecSet = moCn.OpenRes(wOrStr)
    End If
    
    If (wOrRecSet.EOF Or wOrRecSet.BOF) Then
      wRsRmDet!qIrRmARt = 0#
      wRsRmDet!qIrRmIRt = 0#
      wRsRmDet!qIrSetSCd = ""
      wRsRmDet!qIrSetARt = 0#
      If UCase(ms_IdTcTyp) = UCase(ctInTcTyp) Then
        wRsRmDet!qIrRmZRt = 0#
        wRsRmDet!qIrSetIRt = 0#
        wRsRmDet!qIrSetZRt = 0#
      End If
      wRsRmDet!qIrMainMet = "N"
    Else
      
      wOrRecSet.MoveFirst
      wRsRmDet!OrCustRmCd = wOrRecSet!OrCustRmCd
      wRsRmDet!qIrSetSCd = wOrRecSet!OrSetSCd
      wRsRmDet!qIrSetARt = wOrRecSet!OrSetCstRt
      If UCase(ms_IdTcTyp) = UCase(ctInTcTyp) Then
        wRsRmDet!qIrRmARt = wOrRecSet!OrCstRt
        
        'Sachin 3.11.0 - Invoice Rate from Rate Chart
        '6.1b-08 if Use Ptr setting in head is yes as Ord Ptr then it takes order ptr else it takes wt/qty from invoice
        If wRsRmDet!RmCtg = "D" Or wRsRmDet!RmCtg = "C" Then
          wRsRmDet!qIrRmIRt = IIF(ms_RmRtfrRtChrt = "Y", _
                                 IIF(ms_RtFrInvCust = "Y", _
                                     GetRmRt(ADC("wIdCmCd"), wRsRmDet!qirrmcd, wOrRecSet!OrLn1, wOrRecSet!OrLn2, IIF(ms_hInvPtr = "Y", wOrRecSet!OrRmPtr, MWLib.Div(wRsRmDet!qIrRmAWt, wRsRmDet!qIrRmQty)), "S", 0, ms_CmCtg, , , , ADC("wIdCmCurCd"), ""), _
                                     GetRmRt(GrdID.Value(GrdID.Row, "IdExpCmCd"), wRsRmDet!qirrmcd, wOrRecSet!OrLn1, wOrRecSet!OrLn2, IIF(ms_hInvPtr = "Y", wOrRecSet!OrRmPtr, MWLib.Div(wRsRmDet!qIrRmAWt, wRsRmDet!qIrRmQty)), "S", 0, ms_CmCtg, , , , ADC("wIdCmCurCd"), IIF(wOmLkUpRmRt <> "", wOmLkUpRmRt, ""))), _
                                wOrRecSet!OrSalRt)
        Else
          wRsRmDet!qIrRmIRt = IIF(ms_RmRtfrRtChrt = "Y", _
                                 IIF(ms_RtFrInvCust = "Y", _
                                     GetRmRt(ADC("wIdCmCd"), wRsRmDet!qirrmcd, wOrRecSet!OrLn1, wOrRecSet!OrLn2, wOrRecSet!OrRmPtr, "S", 0, ms_CmCtg, , , , ADC("wIdCmCurCd"), ""), _
                                     GetRmRt(GrdID.Value(GrdID.Row, "IdExpCmCd"), wRsRmDet!qirrmcd, wOrRecSet!OrLn1, wOrRecSet!OrLn2, wOrRecSet!OrRmPtr, "S", 0, ms_CmCtg, , , , ADC("wIdCmCurCd"), IIF(wOmLkUpRmRt <> "", wOmLkUpRmRt, ""))), _
                                wOrRecSet!OrSalRt)
        End If
        wRsRmDet!qIrRmZRt = wRsRmDet!qIrRmIRt
        '6.1b-08 if Use Ptr setting in head is yes as Ord Ptr then it takes order ptr else it takes wt/qty from invoice
        If wRsRmDet!RmCtg = "D" Or wRsRmDet!RmCtg = "C" Then
          wRsRmDet!qIrSetIRt = IIF(ms_LabRtfrRtChrt = "Y", _
                                 IIF(ms_RtFrInvCust = "Y", _
                                GetLabRt(ADC("wIdCmCd"), "SET", wOrRecSet!OrSetSCd, "S", IIF(ms_hInvPtr = "Y", wOrRecSet!OrRmPtr, MWLib.Div(wRsRmDet!qIrRmAWt, wRsRmDet!qIrRmQty)), ms_CmCtg, "Q", , , ADC("wIdCmCurCd")), _
                                GetLabRt(GrdID.Value(GrdID.Row, "IdExpCmCd"), "SET", wOrRecSet!OrSetSCd, "S", IIF(ms_hInvPtr = "Y", wOrRecSet!OrRmPtr, MWLib.Div(wRsRmDet!qIrRmAWt, wRsRmDet!qIrRmQty)), ms_CmCtg, "Q", , , ADC("wIdCmCurCd"), IIF(wOmLkUpLabRt <> "", wOmLkUpLabRt, ""))), _
                                wOrRecSet!OrSetSalRt)
                                
        Else
          wRsRmDet!qIrSetIRt = IIF(ms_LabRtfrRtChrt = "Y", _
                                 IIF(ms_RtFrInvCust = "Y", _
                                GetLabRt(ADC("wIdCmCd"), "SET", wOrRecSet!OrSetSCd, "S", wOrRecSet!OrRmPtr, ms_CmCtg, "Q", , , ADC("wIdCmCurCd")), _
                                GetLabRt(GrdID.Value(GrdID.Row, "IdExpCmCd"), "SET", wOrRecSet!OrSetSCd, "S", wOrRecSet!OrRmPtr, ms_CmCtg, "Q", , , ADC("wIdCmCurCd"), IIF(wOmLkUpLabRt <> "", wOmLkUpLabRt, ""))), _
                                wOrRecSet!OrSetSalRt)
        End If
        wRsRmDet!qIrSetZRt = wRsRmDet!qIrSetIRt
      End If
         
      If wOrRecSet!OrMainMet = "Y" And wMainMet = False Then
        wRsRmDet!qIrMainMet = "Y": wMainMet = True
      ElseIf wOrRecSet!OrMainMet = "Y" And wMainMet = True Then
        wRsRmDet!qIrMainMet = "N"
      ElseIf wOrRecSet!OrMainMet = "N" Then
        wRsRmDet!qIrMainMet = "N"
      End If
    End If
    
    wRsRmDet.Update
    wRsRmDet.MoveNext
  Loop
  
  
  
  If UCase(ms_IdTcTyp) = UCase(ctInTcTyp) Then
    Set wRsRmDet = moCn.OpenRes("Select qIrRmCd, sum(qIrRmIWt) as qIrRmWt, " + _
                                " max(qIrMainMet) as qMainMet from " + wTmpTbl + _
                                " Group By qIrRmCd Order By qIrRmCd")
    If (wRsRmDet.RecCount > 0) Then wRsRmDet.MoveFirst
    wTotWtDiff = 0: wMainMetWt = 0
    Do While Not (wRsRmDet.EOF Or wRsRmDet.BOF)
      wFrOrdYN = Trim(mIWtEqOrd)
      wTotIrRmWt = wRsRmDet!qIrRmWt
      wRmCtg = moCn.GetFldVal("Select RmCtg from RmMst Where RmPrtKey='" + ctCurrPrtn + "' and RmCd= '" + wRsRmDet!qirrmcd + "' ")
      wDivFactor = IIF(wRmCtg = "C" Or wRmCtg = "D", 5, 1)
        
      If wFrOrdYN = "N" Then wFrOrdYN = IIF(InStr(1, wCtgStr, "," + wRmCtg + ",") <> 0, "Y", "N")
                
      If wFrOrdYN = "Y" Then
        '****** Sachin 3.02 31-12-07 - PrtKey Added
        wOrStr = "Select sum(OrWt*" + CStr(mIdQty) + ") as qOrRmWt from OrdRm " + _
                "where OrCoCd= '" + gs_CoCd + "' " + _
                " and OrTc = '" + GrdID.Value(GrdID.Row, "IdExpOdTc") + "' " + _
                " and OrYy= '" + GrdID.Value(GrdID.Row, "IdExpOdYy") + "' " + _
                " and OrChr= '" + GrdID.Value(GrdID.Row, "IdExpOdChr") + "' " + _
                " and OrNo = " + CStr(GrdID.Value(GrdID.Row, "IdExpOdNo")) + _
                " and OrSr= " + CStr(GrdID.Value(GrdID.Row, "IdExpOdSr")) + _
                " and OrPrtKey='" + ctCurrPrtn + "' " + _
                " and OrRmCd= '" + wRsRmDet!qirrmcd + "' "
          
        Set wOrRecSet = moCn.OpenRes(wOrStr)
        If Not (wOrRecSet.EOF Or wOrRecSet.BOF) Then
          wOrRecSet.MoveFirst
          wWtDiff = 0
          'wTotIrRmWt = wRsRmDet!qIrRmWt
          If (wTotIrRmWt <> wOrRecSet!qOrRmWt) Then
            wWtDiff = wOrRecSet!qOrRmWt - wTotIrRmWt
            If UCase(wRsRmDet!qMainMet) <> "Y" Then _
               wTotWtDiff = wTotWtDiff + MWLib.Div(wWtDiff, wDivFactor)
          End If
        End If
      End If
        
      'urmi <>, = 23/02/04
      If UCase(wRsRmDet!qMainMet) = "Y" Then _
         wMainMetWt = MWLib.Div(wTotIrRmWt, wDivFactor)
      wRsRmDet.MoveNext
    Loop
    
    
    'urmi <>, < 23/02/04
    '*** (Jen 3.1.3)
    If mb_InInvGrsWtEqActual And wMainMetWt < wTotWtDiff Then
    '*** (Jen 3.1.3)
    'If wMainMetWt < wTotWtDiff Then  '*** (Bef 3.1.3)
      DispMsg "MainMetal Wt is less than the other Wt Diff.", etError
      Exit Sub
    End If
    
    Set wRsRmDet = Nothing
    Set wOrRecSet = Nothing
    '********************************************************
    
    Set wRsRmDet = moCn.OpenRes("Select qIrRmCd, sum(qIrRmIWt) as qIrRmWt, max(qIrMainMet) as qMainMet from " + wTmpTbl + _
                                " Group By qIrRmCd Order By qIrRmCd")
    If Not (wRsRmDet.EOF Or wRsRmDet.BOF) Then wRsRmDet.MoveFirst
    wTotWtDiff = 0
    Do While Not (wRsRmDet.EOF Or wRsRmDet.BOF)
      wFrOrdYN = Trim(mIWtEqOrd)
      wTotIrRmWt = wRsRmDet!qIrRmWt
      wRmCtg = moCn.GetFldVal("Select RmCtg from RmMst where RmPrtKey='" + ctCurrPrtn + "' And RmCd= '" + wRsRmDet!qirrmcd + "' ")
      If wFrOrdYN = "N" Then wFrOrdYN = IIF(InStr(1, wCtgStr, "," + wRmCtg + ",") <> 0, "Y", "N")
      
      If wFrOrdYN = "Y" Then
        '****** Sachin 3.02 31-12-07 - PrtKey Added
        wOrStr = "Select sum(OrWt*" + CStr(mIdQty) + ") as qOrRmWt from OrdRm " + _
                "where OrCoCd= '" + gs_CoCd + "' " + _
                " and OrTc = '" + GrdID.Value(GrdID.Row, "IdExpOdTc") + "' " + _
                " and OrYy= '" + GrdID.Value(GrdID.Row, "IdExpOdYy") + "' " + _
                " and OrChr= '" + GrdID.Value(GrdID.Row, "IdExpOdChr") + "' " + _
                " and OrNo = " + CStr(GrdID.Value(GrdID.Row, "IdExpOdNo")) + _
                " and OrSr= " + CStr(GrdID.Value(GrdID.Row, "IdExpOdSr")) + _
                " and OrPrtKey='" + ctCurrPrtn + "' " + _
                " and OrRmCd= '" + wRsRmDet!qirrmcd + "' "
        
        Set wOrRecSet = moCn.OpenRes(wOrStr)
        If Not (wOrRecSet.EOF Or wOrRecSet.BOF) Then
          wOrRecSet.MoveFirst
          wWtDiff = 0
          'wTotIrRmWt = wRsRmDet!qIrRmWt
          
          If (wTotIrRmWt <> wOrRecSet!qOrRmWt) Then
            wWtDiff = wOrRecSet!qOrRmWt - wTotIrRmWt
            wDivFactor = IIF(wRmCtg = "C" Or wRmCtg = "D", 5, 1)
            If UCase(wRsRmDet!qMainMet) = "Y" Then
                wTotWtDiff = wTotWtDiff + MWLib.Div(wWtDiff, wDivFactor)
            End If
                
            '###
            Set wIrRmWt = moCn.OpenResultset("Select * from " + wTmpTbl + _
                         " Where qIrRmCd = '" + wRsRmDet!qirrmcd + "' ")
    
            If Not (wIrRmWt.EOF Or wIrRmWt.BOF) Then wIrRmWt.MoveFirst
            Do While Not (wIrRmWt.EOF Or wIrRmWt.BOF)
              '*** (Bef 2.12) wAddWt = (MWLib.Div(wWtDiff, wTotIrRmWt)) * wIrRmWt!qIrRmIWt
              '*** (Jen 2.12)
              wAddWt = Round((MWLib.Div(wWtDiff, wTotIrRmWt)) * wIrRmWt!qIrRmIWt, 3)
              '*** (Jen 2.12)
              '### Debug.Print "IrWt: " & (wIrRmWt!qIrRmIWt + wAddWt) & "   WtDiff: " & wWtDiff & "   TotIrWt: " & wTotIrRmWt
              wTotIrRmWt = wTotIrRmWt - wIrRmWt!qIrRmIWt
              wIrRmWt!qIrRmIWt = wIrRmWt!qIrRmIWt + wAddWt
              wIrRmWt.Update
              wWtDiff = wWtDiff - wAddWt
              wIrRmWt.MoveNext
            Loop
          End If
        End If
      End If
      wRsRmDet.MoveNext
    Loop
    
    Set wRsRmDet = moCn.OpenResultset("Select * from " + wTmpTbl + _
                                      " where qIrMainMet= 'Y'")
    If Not (wRsRmDet.EOF Or wRsRmDet.BOF) Then
      wRsRmDet.MoveFirst
      
      '*** (Jen 3.1.1)
      If mb_InInvGrsWtEqActual = False Then
        wFrOrdYN = Trim(mIWtEqOrd)
        wRmCtg = moCn.GetFldVal("Select RmCtg from RmMst where RmPrtKey='" + ctCurrPrtn + "' and RmCd= '" + wRsRmDet!qirrmcd + "' ")
        If wFrOrdYN = "N" Then wFrOrdYN = IIF(InStr(1, wCtgStr, "," + wRmCtg + ",") <> 0, "Y", "N")
        'If wRs_RmCd!qOrdExists = "Y" And wRs_RmCd!qFgExists = "Y" And wFrOrdYN = "Y" Then
        If wFrOrdYN = "Y" Then
          wb_WtFrOrd = True
        Else
          wb_WtFrOrd = False
        End If
        
        If wb_WtFrOrd = True Then
          '5717 copy Gold wt from order is giving wrong wt. after average wt calculation again it is doing OrWt*OdQty only for main metal. so this line commented
          'wRsRmDet!qIrRmIWt = wRsRmDet!qOrWt * mIdQty
          If Trim(whZWtEq) = "I" Then
            wRsRmDet!qIrRmZWt = wRsRmDet!qIrRmIWt
          End If
        Else
          wRsRmDet!qIrRmIWt = wRsRmDet!qIrRmIWt - wTotWtDiff
        End If
      Else
        wRsRmDet!qIrRmIWt = wRsRmDet!qIrRmIWt - wTotWtDiff
      End If
      '*** (Jen 3.1.1)
      
      '*** (Bef 3.1.1)
      'wRsRmDet!qIrRmIWt = wRsRmDet!qIrRmIWt - wTotWtDiff
      '*** (Bef 3.1.1)
      wRsRmDet.Update
    End If
  End If
  
  If UCase(ms_IdTcTyp) = UCase(ctInTcTyp) Then
    moCn.Execute ("Update " + wTmpTbl + " Set qXtraIWt= .000, qXtraZWt= .000 ")
    
    '********* URMI ****** 23/02/04 (commented code) *********
    wDFrOrdYN = IIF(Trim(mIWtEqOrd) = "Y", "Y", "N")
    If wDFrOrdYN = "N" Then wDFrOrdYN = IIF(InStr(1, wCtgStr, "," + "D" + ",") <> 0, "Y", "N")
    
    wCFrOrdYN = IIF(Trim(mIWtEqOrd) = "Y", "Y", "N")
    If wCFrOrdYN = "N" Then wCFrOrdYN = IIF(InStr(1, wCtgStr, "," + "C" + ",") <> 0, "Y", "N")
    
    wDRrTcTyp = IIF(wDFrOrdYN = "Y", "XIO", "XIA")
    wCRrTcTyp = IIF(wCFrOrdYN = "Y", "XIO", "XIA")
    
    Call GetXtraWt(wDRrTcTyp, wCRrTcTyp, wTmpTbl, wCmDiaRtFA, wCmCsRtFA, GrdID.Value(GrdID.Row, "IdExpCmCd"), wDmCtg, _
                   "qIrRmIWt", "qXtraIWt", "qIrRmIRt", mIdQty)
    If whZWtEq = "A" Then
      Call GetXtraWt("XZA", "XZA", wTmpTbl, wCmDiaRtFA, wCmCsRtFA, GrdID.Value(GrdID.Row, "IdExpCmCd"), wDmCtg, _
                     "qIrRmZWt", "qXtraZWt", "qIrRmZRt", mIdQty)
    End If
    '*********************************************************
  End If

  Set wRsRmDet = moCn.OpenRes("select * from " + wTmpTbl + " Order By qIrMainMet")
    
  wTotIXtraWt = 0: wTotZXtraWt = 0
  wTotAWt = 0: wTotIWt = 0: wTotZWt = 0
  Dim wAddBMWtToZWtTotBMWt As Double
  wAddBMWtToZWtTotBMWt = 0

  With GrdIR
    .StartCopy
    Do While Not (wRsRmDet.EOF Or wRsRmDet.BOF)
'sri
    mf_StopRecalcDuringCopyIR = True
      .AddItem
      wRmRow = .Rows - 1
'      UpdateEmrProcess "Copying........InvDesign" + CStr(wDsgRow)
      UpdateEmrProcess "Copying........InvDesign" + CStr(GrdID.Row) + "::OrdRm: " + CStr(wRmRow)
      
      If UCase(ms_IdTcTyp) = UCase(ctInTcTyp) Then
        wDivFactor = IIF(wRsRmDet!qRmCtg = "C" Or wRsRmDet!qRmCtg = "D", 5, 1)
        wTotIXtraWt = wTotIXtraWt + MWLib.Div(wRsRmDet!qXtraIWt, wDivFactor)
        wTotZXtraWt = wTotZXtraWt + MWLib.Div(wRsRmDet!qXtraZWt, wDivFactor)
        
        .Value(wRmRow, "IrSrNo") = wRmRow
        .Value(wRmRow, "IrPrtKey") = ctCurrPrtn
        .Value(wRmRow, "IrRmCd") = wRsRmDet!qirrmcd
        .Value(wRmRow, "IrLotNo") = wRsRmDet!qIrLotNo
        .Value(wRmRow, "IrRmSz") = wRsRmDet!qIrRmSz
        .Value(wRmRow, "IrRmSz2") = wRsRmDet!qIrRmSz2
        .Value(wRmRow, "IrRmSz3") = wRsRmDet!qIrRmSz3
        .Value(wRmRow, "IrStkRt") = wRsRmDet!qIrStkRt
        .Value(wRmRow, "IrRmQty") = wRsRmDet!qIrRmQty
        .Value(wRmRow, "IrRmAWt") = wRsRmDet!qIrRmAWt
        .Value(wRmRow, "IrRmARt") = wRsRmDet!qIrRmARt
        .Value(wRmRow, "IrSetSCd") = wRsRmDet!qIrSetSCd
        .Value(wRmRow, "IrCustRmCd") = wRsRmDet!OrCustRmCd
        .Value(wRmRow, "IrSetARt") = wRsRmDet!qIrSetARt
        .Value(wRmRow, "IrSetIRt") = wRsRmDet!qIrSetIRt
        .Value(wRmRow, "IrMainMet") = wRsRmDet!qIrMainMet
        .Value(wRmRow, "IrSetIRt") = wRsRmDet!qIrSetIRt
        .Value(wRmRow, "IrSetZRt") = wRsRmDet!qIrSetZRt
        Dim wf_BaseLmeOrd As Double
        Dim wf_baselmeinv As Double

        If wRsRmDet!qIrMainMet = "Y" Then
          .Value(wRmRow, "IrAlySalRt") = moCn.GetFldVal("Select  OrAlySalRt from OrdRm where OrPrtKey='" + ctCurrPrtn + "' and OrCoCd='" + gs_CoCd + "' and OrTc= '" + GrdID.Value(GrdID.Row, "IdExpOdTc") + "' and " + _
                                                         "OrYy= '" + GrdID.Value(GrdID.Row, "IdExpOdYy") + "' and OrChr= '" + GrdID.Value(GrdID.Row, "IdExpOdChr") + "' and " + _
                                                         "OrNo= " + CStr(GrdID.Value(GrdID.Row, "IdExpOdNo")) + " and OrSr= " + CStr(GrdID.Value(GrdID.Row, "IdExpOdSr")) + " and  OrMainMet='Y' ")
          Dim wd_Wt As Double
          If mb_InInvGrsWtEqActual = True Then
            wd_Wt = wRsRmDet!qIrRmIWt - (wTotIWt - wTotAWt)
          Else
            wd_Wt = wRsRmDet!qIrRmIWt
          End If
          .Value(wRmRow, "IrRmIWt") = wd_Wt
          If whZWtEq = "A" Then
            .Value(wRmRow, "IrRmZWt") = wRsRmDet!qIrRmZWt - (wTotZWt - wTotAWt)
          ElseIf whZWtEq = "I" Then
            .Value(wRmRow, "IrRmZWt") = wd_Wt     '*** (Jen 3.01.0_1)
          End If
        Else
          .Value(wRmRow, "IrRmIWt") = wRsRmDet!qIrRmIWt + _
                IIF(UCase(wRsRmDet!qRmCtg) = "C" Or UCase(wRsRmDet!qRmCtg) = "D", Round(wRsRmDet!qXtraIWt, 3), wRsRmDet!qXtraIWt)
          If whZWtEq = "A" Then
            .Value(wRmRow, "IrRmZWt") = wRsRmDet!qIrRmZWt + _
                  IIF(UCase(wRsRmDet!qRmCtg) = "C" Or UCase(wRsRmDet!qRmCtg) = "D", Round(wRsRmDet!qXtraZWt, 3), wRsRmDet!qXtraZWt)
          ElseIf whZWtEq = "I" Then
            .Value(wRmRow, "IrRmZWt") = wRsRmDet!qIrRmIWt + _
                  IIF(UCase(wRsRmDet!qRmCtg) = "C" Or UCase(wRsRmDet!qRmCtg) = "D", Round(wRsRmDet!qXtraIWt, 3), wRsRmDet!qXtraIWt)
          End If
        End If
         
        wTotAWt = wTotAWt + MWLib.Div(.Value(wRmRow, "IrRmAWt"), wDivFactor)
        wTotIWt = wTotIWt + MWLib.Div(.Value(wRmRow, "IrRmIWt"), wDivFactor)
        wTotZWt = wTotZWt + MWLib.Div(.Value(wRmRow, "IrRmZWt"), wDivFactor)
        Dim wMetStkRt As Boolean
        wMetStkRt = IIF(wRsInvHd!InMetIRtWrtStkYN = "Y", True, False)
        If UCase(wRsRmDet!qRmCtg) = "G" And (mIGldRtEqOrd <> "Y" Or ms_RmRtfrRtChrt = "Y") And Not wMetStkRt Then
          .Value(wRmRow, "IrRmIRt") = IIF(ms_RtFrInvCust = "Y", _
                                        GetGldRt(ADC("wIdCmCd"), ADC("wIdLmgSal"), _
                                            wRsRmDet!qirrmcd, .Value(wRmRow, "IrAlySalRt"), "DM", GrdID.Value(GrdID.Row, "IdDmCd"), "S", "C", , ADC("wIdCmCurCd"), "", ""), _
                                          GetGldRt(GrdID.Value(GrdID.Row, "IdExpCmCd"), ADC("wIdLmgSal"), _
                                            wRsRmDet!qirrmcd, .Value(wRmRow, "IrAlySalRt"), "DM", GrdID.Value(GrdID.Row, "IdDmCd"), "S", "C", , ADC("wIdCmCurCd"), _
                                            IIF(wOmLkUpRmRt <> "", wOmLkUpRmRt, ""), IIF(wOmLkUpMetLs <> "", wOmLkUpMetLs, ""), , , .Value(wRmRow, "IRRMSZ"), .Value(wRmRow, "IRRMSZ2"), 0))
          .Value(wRmRow, "IrRmZRt") = .Value(wRmRow, "IrRmIRt")
        ElseIf UCase(wRsRmDet!qRmCtg) = "P" And (mIGldRtEqOrd <> "Y" Or ms_RmRtfrRtChrt = "Y") And Not wMetStkRt Then
         .Value(wRmRow, "IrRmIRt") = IIF(ms_RtFrInvCust = "Y", _
                                        GetGldRt(ADC("wIdCmCd"), ADC("wIdLmpSal"), _
                                            wRsRmDet!qirrmcd, .Value(wRmRow, "IrAlySalRt"), "DM", GrdID.Value(GrdID.Row, "IdDmCd"), "S", "C", , ADC("wIdCmCurCd"), "", ""), _
                                      GetGldRt(GrdID.Value(GrdID.Row, "IdExpCmCd"), ADC("wIdLmpSal"), _
                                      wRsRmDet!qirrmcd, .Value(wRmRow, "IrAlySalRt"), "DM", GrdID.Value(GrdID.Row, "IdDmCd"), "S", "C", , ADC("wIdCmCurCd"), _
                                      IIF(wOmLkUpRmRt <> "", wOmLkUpRmRt, ""), IIF(wOmLkUpMetLs <> "", wOmLkUpMetLs, ""), , , .Value(wRmRow, "IRRMSZ"), .Value(wRmRow, "IRRMSZ2"), 0))
          .Value(wRmRow, "IrRmZRt") = .Value(wRmRow, "IrRmIRt")
        ElseIf UCase(wRsRmDet!qRmCtg) = "S" And (mIGldRtEqOrd <> "Y" Or ms_RmRtfrRtChrt = "Y") And Not wMetStkRt Then
          If mSlvMod = True Then
            .Value(wRmRow, "IrRmIRt") = IIF(ms_RtFrInvCust = "Y", _
                                        GetGldRt(ADC("wIdCmCd"), ADC("wIdLmsSal"), _
                                            wRsRmDet!qirrmcd, .Value(wRmRow, "IrAlySalRt"), "DM", GrdID.Value(GrdID.Row, "IdDmCd"), "S", "C", , ADC("wIdCmCurCd"), "", ""), _
                                      GetGldRt(GrdID.Value(GrdID.Row, "IdExpCmCd"), ADC("wIdLmsSal"), _
                                      wRsRmDet!qirrmcd, .Value(wRmRow, "IrAlySalRt"), "DM", GrdID.Value(GrdID.Row, "IdDmCd"), "S", "C", , ADC("wIdCmCurCd"), _
                                      IIF(wOmLkUpRmRt <> "", wOmLkUpRmRt, ""), IIF(wOmLkUpMetLs <> "", wOmLkUpMetLs, "")))
          Else
            .Value(wRmRow, "IrRmIRt") = 0
          End If
          .Value(wRmRow, "IrRmZRt") = .Value(wRmRow, "IrRmIRt")
        ElseIf UCase(wRsRmDet!qRmCtg) = "L" And (mIGldRtEqOrd <> "Y" Or ms_RmRtfrRtChrt = "Y") And Not wMetStkRt Then
          If mb_PdMod = True Then
          .Value(wRmRow, "IrRmIRt") = IIF(ms_RtFrInvCust = "Y", _
                                        GetGldRt(ADC("wIdCmCd"), ADC("wIdLmlSal"), _
                                            wRsRmDet!qirrmcd, .Value(wRmRow, "IrAlySalRt"), "DM", GrdID.Value(GrdID.Row, "IdDmCd"), "S", "C", , ADC("wIdCmCurCd"), "", ""), _
                                      GetGldRt(GrdID.Value(GrdID.Row, "IdExpCmCd"), ADC("wIdLmlSal"), _
                                      wRsRmDet!qirrmcd, .Value(wRmRow, "IrAlySalRt"), "DM", GrdID.Value(GrdID.Row, "IdDmCd"), "S", "C", , ADC("wIdCmCurCd"), _
                                      IIF(wOmLkUpRmRt <> "", wOmLkUpRmRt, ""), IIF(wOmLkUpMetLs <> "", wOmLkUpMetLs, ""), , , .Value(wRmRow, "IRRMSZ"), .Value(wRmRow, "IRRMSZ2"), 0))
          Else
            .Value(wRmRow, "IrRmIRt") = 0
          End If
          .Value(wRmRow, "IrRmZRt") = .Value(wRmRow, "IrRmIRt")
          
        ElseIf UCase(wRsRmDet!qRmCtg) = "X" And (mIGldRtEqOrd <> "Y" Or ms_RmRtfrRtChrt = "Y") And _
               moCn.RecSeek("select 'x' from Param where PTyp= 'RMSCTG' and PMCd= '" + wRsRmDet!qRmCtg + "' " + _
                            "and PSCd= '" + wRsRmDet!qRmSCtg + "' and PValue= 'S' ") And _
                wRsInvHd!InXIRtWrtStkYN = "N" Then
                            
          
          Dim ws_BaseRmCtg As String
          ws_BaseRmCtg = moCn.GetFldVal("Select B.RmCtg from RmMst B where " + _
                         "RmCd= (Select Rm.RmBaseCd from RmMst Rm where Rm.RmCd= '" + wRsRmDet!qirrmcd + "' and Rm.RmPrtKey='" + ctCurrPrtn + "' ) and RmPrtKey='" + ctCurrPrtn + "' ")
          wf_BaseLmeOrd = IIF(ws_BaseRmCtg = "G" Or ws_BaseRmCtg = "P" Or ws_BaseRmCtg = "S" Or ws_BaseRmCtg = "L", wRsOm.FldValue("OmLm" + ws_BaseRmCtg + "Sal"), 0)
          
          If ws_BaseRmCtg = "G" Or ws_BaseRmCtg = "P" Or ws_BaseRmCtg = "L" Or ws_BaseRmCtg = "S" Then
            wf_baselmeinv = ADC("wIdLm" + ws_BaseRmCtg + "Sal")
          Else
            wf_baselmeinv = 0
          End If
          '6.1-55 cmcd & LkUprate are changed. ms_RtFrInvCust = "Y" is added
          If GetRmRt(IIF(ms_RtFrInvCust = "Y", ADC("wIdCmCd"), GrdID.Value(GrdID.Row, "IdExpCmCd")), wRsRmDet!qirrmcd, 0, 0, 0, "S", wf_BaseLmeOrd, , , , , ADC("wIdCmCurCd"), IIF(ms_RtFrInvCust = "Y", "", IIF(wOmLkUpRmRt <> "", wOmLkUpRmRt, ""))) = _
             GetRmRt(IIF(ms_RtFrInvCust = "Y", ADC("wIdCmCd"), GrdID.Value(GrdID.Row, "IdExpCmCd")), wRsRmDet!qirrmcd, 0, 0, 0, "S", wf_baselmeinv, , , , , ADC("wIdCmCurCd"), IIF(ms_RtFrInvCust = "Y", "", IIF(wOmLkUpRmRt <> "", wOmLkUpRmRt, ""))) Then
          '*** (Jen 2.13)
            .Value(wRmRow, "IrRmIRt") = wRsRmDet!qIrRmIRt
          Else  '6.1-55 cmcd & LkUprate are changed. ms_RtFrInvCust = "Y" is added
            .Value(wRmRow, "IrRmIRt") = GetRmRt(IIF(ms_RtFrInvCust = "Y", ADC("wIdCmCd"), GrdID.Value(GrdID.Row, "IdExpCmCd")), wRsRmDet!qirrmcd, 0, 0, 0, "S", wf_baselmeinv, , , , , ADC("wIdCmCurCd"), IIF(ms_RtFrInvCust = "Y", "", IIF(wOmLkUpRmRt <> "", wOmLkUpRmRt, "")))
          End If
          .Value(wRmRow, "IrRmZRt") = .Value(wRmRow, "IrRmIRt")
        Else
          If ((wRsRmDet!qRmCtg = "D" Or wRsRmDet!qRmCtg = "C") And _
                moCn.RecSeek("Select 'x' From InvHd Where InCoCd='" + gs_CoCd + "' And InTc='" + ADC("wIdTc") + "' And InYy='" + ADC("wIdYy") + "' And InChr='" + ADC("wIdChr") + "' And InNo=" + CStr(ADC("wIdNo")) + " And InPrtKey='" + ADC("wIdPrtKey") + "' And InIRtWrtStkYN='Y' ")) Or _
              (wRsRmDet!qRmCtg = "X" And _
                moCn.RecSeek("Select 'x' From InvHd Where InCoCd='" + gs_CoCd + "' And InTc='" + ADC("wIdTc") + "' And InYy='" + ADC("wIdYy") + "' And InChr='" + ADC("wIdChr") + "' And InNo=" + CStr(ADC("wIdNo")) + " And InPrtKey='" + ADC("wIdPrtKey") + "' And InXIRtWrtStkYN='Y' ")) Or _
              ((wRsRmDet!qRmCtg = "G" Or wRsRmDet!qRmCtg = "P" Or wRsRmDet!qRmCtg = "S" Or wRsRmDet!qRmCtg = "L") And wMetStkRt) Then
             Dim wStkRtQW As Single, wRmQW As String
            wRmQW = wRsRmDet!RmQw
             Dim wStkRtInCmCur As Single
             wStkRtInCmCur = IIF(wBaseCurCd <> ADC("wIdCmCurCd"), IIF(ADC("wIdMulDiv") = "M", wRsRmDet!qIrStkRt * ADC("wIdCnvFct"), MWLib.Div(wRsRmDet!qIrStkRt, ADC("wIdCnvFct"))), wRsRmDet!qIrStkRt)
             
             wStkRtQW = IIF(wRmQW = "Q" And (wRsRmDet!qRmCtg = "X" Or wRsRmDet!qRmCtg = "D" Or wRsRmDet!qRmCtg = "C"), wStkRtInCmCur * wRsRmDet!qIrRmIWt / IIF(wRsRmDet!qIrRmQty = 0, 1, wRsRmDet!qIrRmQty), wStkRtInCmCur)
             
             Dim wb_IsCustRmCd As Boolean
              wb_IsCustRmCd = IIF(wRsRmDet!RmCmCd <> "", True, False)
             
             wCurrMrkUpRt = GetMarkUpRt(GrdID.Value(GrdID.Row, "IdExpCmCd"), wRsRmDet!qRmCtg, wRsRmDet!qRmSCtg, wStkRtQW, ADC("wIdCmCurCd"), wDmCtg, GrdID.Value(GrdID.Row, "IdDmCd"), wb_IsCustRmCd)
             
             
             wRmQtyWt = IIF(wRmQW = "Q", (MWLib.Div(wRsRmDet!qIrRmQty, mIdQty)), (MWLib.Div(.Value(wRmRow, "IrRmIWt"), mIdQty)))
             wMrkUpVal = wRmQtyWt * wCurrMrkUpRt
             wFnlIncrVal = IIF(wMrkUpLmt > 0, IIF(wMrkUpVal > (wMrkUpLmt - wTotMrkUpVal), (wMrkUpLmt - wTotMrkUpVal), wMrkUpVal), wMrkUpVal)
             .Value(wRmRow, "IrRmIRt") = wStkRtQW + MWLib.Div(wFnlIncrVal, wRmQtyWt)
             wTotMrkUpVal = wTotMrkUpVal + wFnlIncrVal
             
             .Value(wRmRow, "IrRmZRt") = .Value(wRmRow, "IrRmIRt")
          Else
           .Value(wRmRow, "IrRmIRt") = wRsRmDet!qIrRmIRt
           .Value(wRmRow, "IrRmZRt") = wRsRmDet!qIrRmZRt
          End If
        End If
        wf_CustmCdPurityZ = moCn.GetFldVal("Select R.RmPurityZ * " + _
                            "(case when (R.RmCtg= 'G' and " + CStr(ADC("wIdGldFiness")) + "> 0) then " + CStr(ADC("wIdGldFiness")) + " " + _
                            "      when (R.RmCtg= 'P' and " + CStr(ADC("wIdPlFiness")) + "> 0) then " + CStr(ADC("wIdPlFiness")) + " " + _
                            "      when (R.RmCtg= 'S' and " + CStr(ADC("wIdSlFiness")) + "> 0) then " + CStr(ADC("wIdSlFiness")) + " " + _
                            "      when (R.RmCtg= 'L' and " + CStr(ADC("wIdPdFiness")) + "> 0) then " + CStr(ADC("wIdPdFiness")) + " " + _
                            "else 1 end) from RmMst R where R.RmPrtKey='" + ctCurrPrtn + "' and R.RmCd= (Select R1.RmZCd from RmMst R1 where R1.RmPrtKey='" + ctCurrPrtn + "' And R1.RmCd= '" + wRsRmDet!qirrmcd + "') ")
        
        .Value(wRmRow, "IrRmZPureRt") = IIF(wRsRmDet!RmQw = "Q", _
                                         MWLib.Div(.Value(wRmRow, "IrRmZVal"), .Value(wRmRow, "IrRmZWt")), _
                                         .Value(wRmRow, "IrRmZRt")) / IIF(wRsRmDet!qRmPurityZ > 0, wRsRmDet!qRmPurityZ, 1) * _
                                         IIF(wf_CustmCdPurityZ > 0, wf_CustmCdPurityZ, 1)
      End If
'      sri
    mf_StopRecalcDuringCopyIR = False
      
    'SRI
    Call GrdIR_ReCalcRoutine(wRmRow, UCase("IrRmAVal"), 0, "", "")
      
    'SRI
    Call GrdIR_ReCalcRoutine(wRmRow, UCase("IrRmIRt"), 0, "", "")
      
    'sri
    'SRI
    Call GrdIR_ReCalcRoutine(wRmRow, UCase("IrRmIVal"), 0, "", "")
      
     'sri
     'SRI
     Call GrdIR_ReCalcRoutine(wRmRow, UCase("IrRmZRt"), 0, "", "")
      
     'sri
     Call GrdIR_ReCalcRoutine(wRmRow, UCase("IrRmZVal"), 0, "", "")
      
     'sri
     Call GrdIR_ReCalcRoutine(wRmRow, UCase("IrSetAVal"), 0, "", "")
      
    'srI
    Call GrdIR_ReCalcRoutine(wRmRow, UCase("IrSetIRt"), 0, "", "")
      
     'srI
     Call GrdIR_ReCalcRoutine(wRmRow, UCase("IrSetIVal"), 0, "", "")
      
     'sri
     Call GrdIR_ReCalcRoutine(wRmRow, UCase("IrSetZRt"), 0, "", "")
      
     'sri
     Call GrdIR_ReCalcRoutine(wRmRow, UCase("IrSetZVal"), 0, "", "")
      
     'sri
     Call GrdIR_ReCalcRoutine(wRmRow, UCase("IrRmZPureRt"), 0, "", "")
      
     'sri
    .Value(wRmRow, "IrRmZCd") = wRsRmDet!RmZCd
'     Call GrdIR_ReCalcRoutine(wRmRow, UCase("IrRmZCd"), 0, "", "")
      
     'srI
     'Call GrdIR_ReCalcRoutine(wRmRow, UCase("wQwSal"), 0, "", "")
    .Value(wRmRow, "WQWSAL") = wRsRmDet!RmQw

'      TrackStatus ("Update of IrRm record Started ") + CStr(wRmRow)
      
      .SaveRec (wRmRow)
'      TrackStatus ("Update of IrRm record Completed ") + CStr(wRmRow)

      wRsRmDet.MoveNext
    Loop
    .EndCopy
  End With
  
  Set wRsRmDet = Nothing
  Set wOrRecSet = Nothing
End Sub


Private Sub MakeLab(ByVal mExpOdTc As String, ByVal mExpOdYy As String, _
                    ByVal mExpOdChr As String, ByVal mExpOdNo As Single, _
                    ByVal mExpOdSr As Single)
 '****Geeta***3.02 Patch******New Parameter added IdQty
  
  '*** This procedure is used to write records into the Labour grid
      'The copying is done from the OrdLab table for the Export Order
  
  Dim wRsSrcLab As MDORowSet, wLabRow As Integer
  
        Dim wOmLkUpRmRt As String, wOmLkUpMetLs As String, wOmLkUpStr As String, wOmLkUpLabRt As String
        Dim wOmLkUpArr() As String
'        ADC.UsrCd = "MIC"

        wOmLkUpLabRt = moCn.GetFldVal("Select OmLkUpLabRt from OrdMst where " + _
                                    "OmCoCd= '" + gs_CoCd + "' and OmTc= '" + GrdID.Value(GrdID.Row, "IdExpOdTc") + "' " + _
                                    "and OmYy= '" + GrdID.Value(GrdID.Row, "IdExpOdYy") + "' " + _
                                    "and OmChr= '" + GrdID.Value(GrdID.Row, "IdExpOdChr") + "' " + _
                                    "and OmNo= " + CStr(GrdID.Value(GrdID.Row, "IdExpOdNo")) + " and OmPrtKey='" + ctCurrPrtn + "' ")
                                    
  '****** (Jen 2.14 Next) (20/03/07) added field OlCstQw ******
  '****** Sachin 3.02 31-12-07 - PrtKey added
  Set wRsSrcLab = moCn.OpenRes("Select OlMCd, OlSCd, OlQty, " + _
      "OlCstRt as qLabARt, OlSalRt as qLabIRt, OlSalRt as qLabZRt, OlQw, OlCstQw " + _
      "from OrdLab where OlCoCd= '" + gs_CoCd + "' and " + _
      "OlTc= '" + mExpOdTc + "' and OlYy= '" + mExpOdYy + "' and " + _
      "OlChr= '" + mExpOdChr + "' and OlNo= " + CStr(mExpOdNo) + " and " + _
      "OlSr= " + CStr(mExpOdSr) + " and OlPrtKey='" + ctCurrPrtn + "' Order By OlSr")
  
  With GrdIL
    GrdIL.StartCopy
   'SRI mf_StopRecalcDuringCopyIL = True
    Do While Not (wRsSrcLab.EOF Or wRsSrcLab.BOF)
      .AddItem
      
      wLabRow = .Rows - 1
      UpdateEmrProcess "Copying........InvDesign" + CStr(GrdID.Row) + "::OrdLab: " + CStr(wLabRow)
      
      .Value(wLabRow, "iLSrNo") = wLabRow
      .Value(wLabRow, "iLMCd") = wRsSrcLab!OlMCd
      .Value(wLabRow, "iLSCd") = wRsSrcLab!OlSCd
      .Value(wLabRow, "iLQty") = wRsSrcLab!OlQty
      .Value(wLabRow, "ilQw") = wRsSrcLab!OlQw
      .Value(wLabRow, "ilCstQw") = wRsSrcLab!OlCstQw    '*** (Jen 2.14 Next) (20/03/07)
      .Value(wLabRow, "iLARt") = wRsSrcLab!qLabARt
      
      If UCase(ms_IdTcTyp) = UCase(ctInTcTyp) Then
        'Sachin 3.11.0 - Labor Rt Chart from Rate Chart
        .Value(wLabRow, "iLiRt") = IIF(ms_LabRtfrRtChrt = "Y", _
                                        IIF(ms_RtFrInvCust = "Y", _
                                            GetLabRt(ADC("wIdCmCd"), wRsSrcLab!OlMCd, wRsSrcLab!OlSCd, "S", MWLib.Div(GrdID.Value(GrdID.Row, "IdILabAsWt"), GrdID.Value(GrdID.Row, "IdQty")), ms_CmCtg, wRsSrcLab!OlQw, , , ADC("wIdCmCurCd"), ""), _
                                            GetLabRt(GrdID.Value(GrdID.Row, "IdExpCmCd"), wRsSrcLab!OlMCd, wRsSrcLab!OlSCd, "S", MWLib.Div(GrdID.Value(GrdID.Row, "IdILabAsWt"), GrdID.Value(GrdID.Row, "IdQty")), ms_CmCtg, wRsSrcLab!OlQw, , , ADC("wIdCmCurCd"), IIF(wOmLkUpLabRt <> "", wOmLkUpLabRt, ""))), _
                                    wRsSrcLab!qLabiRt)
        .Value(wLabRow, "iLZRt") = .Value(wLabRow, "iLiRt")
      End If
      .Value(wLabRow, "IlPrtKey") = ctCurrPrtn
      
' SRI     mf_StopRecalcDuringCopyIL = True
      
      
      Call GrdIL_RecalcRoutine(wLabRow, UCase("IlQty"), 0, "", "")
      
      Call GrdIL_RecalcRoutine(wLabRow, UCase("IlAVal"), 0, "", "")
      
      Call GrdIL_RecalcRoutine(wLabRow, UCase("ILiRt"), 0, "", "")
      
      Call GrdIL_RecalcRoutine(wLabRow, UCase("ILiVal"), 0, "", "")
      
      Call GrdIL_RecalcRoutine(wLabRow, UCase("ILZRt"), 0, "", "")
      
      Call GrdIL_RecalcRoutine(wLabRow, UCase("IlZVal"), 0, "", "")
      
 '   SRI  mf_StopRecalcDuringCopyIL = True
'      TrackStatus ("Saving of IL Started ") + CStr(wLabRow)
      .SaveRec (wLabRow)
 '     TrackStatus ("Saving of IL done ") + CStr(wLabRow)
      wRsSrcLab.MoveNext
    Loop
    GrdIL.EndCopy
  End With
  Set wRsSrcLab = Nothing

End Sub
Private Sub GetGldAsWt(ByVal mIdRow As Double, ByVal mTypeList As String, ByRef wAGldAsWt As Double, ByRef wIGldAsWt As Double, ByRef wZGldAsWt As Double, _
              Optional ByVal mIrRow As Double = 0)
  
  Dim wChnGldAs As String, ws_RmSCtg As String                                                           '************Bhavna for chain gold wt calc
  Dim wBigStr As String, i As Integer, wRmCtg As String, wWt As Double
  Dim wMainMetRmCtg As String
  Dim wfIrRmAWt As Boolean, wfIrRmIWt As Boolean, wfIrRmZWt As Boolean
  
  wfIrRmAWt = IIF(InStr(1, mTypeList, "A") > 0, True, False)
  wfIrRmIWt = IIF(InStr(1, mTypeList, "I") > 0, True, False)
  wfIrRmZWt = IIF(InStr(1, mTypeList, "Z") > 0, True, False)
    
  wBigStr = "," + moCn.GetFldVal("Select PDesc225 from Param where " + _
                  "PTyp='GLDAS' and PMCd = '" + GrdID.Value(mIdRow, "IdGldAs") + "'") + ","
  
  If ADC("wChnCustYn") = "Y" Then
    wChnGldAs = "," + moCn.GetFldVal("Select hChnSubCtgForGld from Head where hCoCd = 'ZZZ'") + ","        '************Bhavna for chain gold wt calc
  Else
    wChnGldAs = ""
  End If
  
 ' wChnGldAs = "," + moCn.GetFldVal("Select hChnSubCtgForGld from Head where hCoCd = 'ZZZ'") + ","        '************Bhavna for chain gold wt calc
  
  If mIrRow > 0 And Not GrdIR.IsDel(mIrRow) Then
    If GrdIR.Value(mIrRow, "IrMainMet") = "Y" Then wMainMetRmCtg = moCn.GetFldVal("Select RmCtg from RmMst where RmCd= '" + GrdIR.Value(mIrRow, "IrRmCd") + "' " + IIF(gs_Partition = ctCurrPrtn, " And RmPrtKey='" + GrdIR.Value(mIrRow, "IrPrtKey") + "'", ""))
  Else
    wMainMetRmCtg = moCn.GetFldVal("Select RmCtg from InvRm join RmMst on RmCd= IrRmCd " + IIF(gs_Partition = ctCurrPrtn, " And RmPrtKey=IrPrtKey ", "") + _
                    "Where IrCoCd= '" + gs_CoCd + "' and IrTc= '" + ADC("wIdTc") + "' and " + _
                    "IrYy= '" + ADC("wIdYy") + "' and IrChr= '" + ADC("wIdChr") + "' and " + _
                    "IrNo= " + CStr(ADC("wIdNo")) + " and " + _
                    "IrSr= " + CStr(GrdID.Value(GrdID.Row, "IdSr")) + " and IrPrtKey='" + ctCurrPrtn + "' and IrMainMet= 'Y'")
  End If
  
  For i = 1 To GrdIR.Rows - 1
    If Not GrdIR.IsDel(i) Then
      wRmCtg = Trim(moCn.GetFldVal("Select RmCtg from RmMst where RmCd = '" + GrdIR.Value(i, "IrRmCd") + "' and RmPrtKey='" + ctCurrPrtn + "'"))
      '************Bhavna for chain gold wt calc
      ws_RmSCtg = Trim(moCn.GetFldVal("Select RmSCtg from RmMst where RmCd = '" + GrdIR.Value(i, "IrRmCd") + "' and RmPrtKey='" + ctCurrPrtn + "'"))
      '************Bhavna for chain gold wt calc
      
      ' *****bhavna included the check that rmsctg not in Chain gold as ctg
      wWt = IIF(((InStr(1, wBigStr, "," + wRmCtg + ",") <> 0) And _
            wRmCtg <> wMainMetRmCtg And _
            (InStr(1, wChnGldAs, "," + ws_RmSCtg + ",") = 0) And _
            wRmCtg <> "G" And wRmCtg <> "P" And wRmCtg <> "S" And wRmCtg <> "L") _
            Or GrdIR.Value(i, "IrMainMet") = "Y", _
            GrdIR.Value(i, "IrRmAWt"), 0)
      wAGldAsWt = wAGldAsWt + IIF(wRmCtg = "D" Or wRmCtg = "C", MWLib.Div(wWt, 5), wWt)
      
      wWt = IIF(((InStr(1, wBigStr, "," + wRmCtg + ",") <> 0) And _
            wRmCtg <> wMainMetRmCtg And _
            (InStr(1, wChnGldAs, "," + ws_RmSCtg + ",") = 0) And _
            wRmCtg <> "G" And wRmCtg <> "P" And wRmCtg <> "S" And wRmCtg <> "L") _
            Or GrdIR.Value(i, "IrMainMet") = "Y", _
            GrdIR.Value(i, "IrRmIWt"), 0)
            
      wIGldAsWt = wIGldAsWt + IIF(wRmCtg = "D" Or wRmCtg = "C", MWLib.Div(wWt, 5), wWt)
      
      
      wWt = IIF(((InStr(1, wBigStr, "," + wRmCtg + ",") <> 0) And _
            wRmCtg <> wMainMetRmCtg And _
            (InStr(1, wChnGldAs, "," + ws_RmSCtg + ",") = 0) And _
            wRmCtg <> "G" And wRmCtg <> "P" And wRmCtg <> "S" And wRmCtg <> "L") _
            Or GrdIR.Value(i, "IrMainMet") = "Y", _
            GrdIR.Value(i, "IrRmZWt"), 0)
            
      wZGldAsWt = wZGldAsWt + IIF(wRmCtg = "D" Or wRmCtg = "C", MWLib.Div(wWt, 5), wWt)
      
    End If
  Next i
End Sub

Private Sub GetLabAsWt(ByVal mIdRow As Double, ByVal mTypeList As String, _
            ByRef wIdLabAAsWt As Double, ByRef wIdLabIAsWt As Double, ByRef wIdLabZAsWt As Double)
            
  Dim wChnLabAs As String, ws_RmSCtg As String                                                   ' **************Bhavna for Chain Labour wt calc
  Dim wLabAsWt As Double, wBigStr As String, i As Integer, wRmCtg As String, wWt As Double
  Dim wWtColName As String
  Dim wIrRmAWt As Double, wIrRmIWt As Double, wIrRmZWt As Double
  
  wBigStr = "," + moCn.GetFldVal("Select PDesc225 from Param where " + _
                  "PTyp='GLDAS' and PMCd = '" + GrdID.Value(mIdRow, "IdLabAs") + "'") + ","
                  
  If ADC("wChnCustYn") = "Y" Then
    wChnLabAs = "," + moCn.GetFldVal("Select hChnSubCtgForLab from Head where hCoCd = 'ZZZ'") + ","        '************Bhavna for chain gold wt calc
  Else
    wChnLabAs = ""
  End If
  
  ' *****bhavna included the check that rmsctg not in Chain gold as ctg
  For i = 1 To GrdIR.Rows - 1
    If Not GrdIR.IsDel(i) Then
      wRmCtg = Trim(moCn.GetFldVal("Select RmCtg from RmMst where RmCd = '" + GrdIR.Value(i, "IrRmCd") + "' and RmPrtKey='" + ctCurrPrtn + "'"))
      '************Bhavna for chain labour gold wt calc
      ws_RmSCtg = Trim(moCn.GetFldVal("Select RmSCtg from RmMst where RmCd = '" + GrdIR.Value(i, "IrRmCd") + "' and RmPrtKey='" + ctCurrPrtn + "'"))
      '************Bhavna for chain labour gold wt calc
      
      wWt = IIF((InStr(1, wBigStr, "," + wRmCtg + ",") <> 0) And (InStr(1, wChnLabAs, "," + ws_RmSCtg + ",") = 0), _
            GrdIR.Value(i, "IrRmAWt"), 0)
      wIdLabAAsWt = wIdLabAAsWt + IIF(wRmCtg = "D" Or wRmCtg = "C", MWLib.Div(wWt, 5), wWt)
      
      wWt = IIF((InStr(1, wBigStr, "," + wRmCtg + ",") <> 0) And (InStr(1, wChnLabAs, "," + ws_RmSCtg + ",") = 0), _
            GrdIR.Value(i, "IrRmIWt"), 0)
      wIdLabIAsWt = wIdLabIAsWt + IIF(wRmCtg = "D" Or wRmCtg = "C", MWLib.Div(wWt, 5), wWt)

      wWt = IIF((InStr(1, wBigStr, "," + wRmCtg + ",") <> 0) And (InStr(1, wChnLabAs, "," + ws_RmSCtg + ",") = 0), _
            GrdIR.Value(i, "IrRmZWt"), 0)
      wIdLabZAsWt = wIdLabZAsWt + IIF(wRmCtg = "D" Or wRmCtg = "C", MWLib.Div(wWt, 5), wWt)

    End If
  Next i
End Sub


' Zubin 3.02.01
Private Function GetLabAsWtFrOrd() As Double
  Dim wLabAsWt As Double, wBigStr As String, wWt As Double

  Dim wRs As MwfLib.MDORowSet
  Dim wOdLabAs As String
  wOdLabAs = moCn.GetFldVal("Select OdLabAs from OrdDsg " + _
                 "where OdCoCd= '" + gs_CoCd + "' and OdTc= '" + GrdID.Value(GrdID.Row, "IdExpOdTc") + "' and " + _
                "OdYy= '" + GrdID.Value(GrdID.Row, "IdExpOdYy") + "' and " + _
                "OdChr= '" + GrdID.Value(GrdID.Row, "IdExpOdChr") + "' and " + _
                "OdNo= " + CStr(GrdID.Value(GrdID.Row, "IdExpOdNo")) + " and " + _
                "OdSr= " + CStr(GrdID.Value(GrdID.Row, "IdExpOdSr")) + " and " + _
                "OdPrtKey='" + ctCurrPrtn + "' ")

  Set wRs = moCn.OpenRes("Select OrRmCtg, OrRmCd, OrWt from OrdRm " + _
               "where OrCoCd= '" + gs_CoCd + "' and OrTc= '" + GrdID.Value(GrdID.Row, "IdExpOdTc") + "' and " + _
               "OrYy= '" + GrdID.Value(GrdID.Row, "IdExpOdYy") + "' and " + _
               "OrChr= '" + GrdID.Value(GrdID.Row, "IdExpOdChr") + "' and " + _
               "OrNo= " + CStr(GrdID.Value(GrdID.Row, "IdExpOdNo")) + " and " + _
               "OrSr= " + CStr(GrdID.Value(GrdID.Row, "IdExpOdSr")) + " and " + _
               "OrPrtKey='" + ctCurrPrtn + "' ")
  
  If Not (wRs.EOF Or wRs.BOF) Then
    wBigStr = "," + moCn.GetFldVal("Select PDesc225 from Param where " + _
                  "PTyp='GLDAS' and PMCd = '" + wOdLabAs + "'") + ","
    Do While Not (wRs.EOF Or wRs.BOF)
      wWt = IIF((InStr(1, wBigStr, "," + wRs!OrRmCtg + ",") <> 0), wRs!OrWt, 0)
      wLabAsWt = wLabAsWt + IIF(wRs!OrRmCtg = "D" Or wRs!OrRmCtg = "C", MWLib.Div(wWt, 5), wWt)
      wRs.MoveNext
    Loop
    GetLabAsWtFrOrd = wLabAsWt
  Else
    GetLabAsWtFrOrd = 0
  End If

End Function

Private Function GetGldAsWt1(ByVal mIdRow As Double, ByVal mType As String) As Double
'  '*** Bef 12/05/03
'  '*** Procedure used for calculating value of the IdAGldAsWt or IdiGldAsWt or IdZGldAsWt fields
'
'  Dim wGldAsWt As Double, wBigStr As String, i As Integer, wRmCtg As String, wWt As Double
'  Dim wWtColName As String
'  If mType = "A" Then wWtColName = "IrRmAWt"
'  If mType = "I" Then wWtColName = "IrRmIWt"
'  If mType = "Z" Then wWtColName = "IrRmZWt"
'  wBigStr = "," + moCn.GetFldVal("Select PDesc225 from Param where " + _
'                  "PTyp='GLDAS' and PMCd = '" + GrdID.Value(mIdRow, "IdGldAs") + "'") + ","
'  For i = 1 To GrdIR.Rows - 1
'    If Not GrdIR.IsDel(i) Then
'      wRmCtg = Trim(moCn.GetFldVal("Select RmCtg from RmMst where RmCd = '" + GrdIR.Value(i, "IrRmCd") + "'"))
'      wWt = IIf(InStr(1, wBigStr, "," + wRmCtg + ",") <> 0, GrdIR.Value(i, wWtColName), 0)
'      wGldAsWt = wGldAsWt + IIf(wRmCtg = "D" Or wRmCtg = "C", MWLib.Div(wWt, 5), wWt)
'    End If
'  Next i
'  GetGldAsWt = wGldAsWt
End Function

' Manoj 2.10.0
Private Sub DsgSalHst(ByVal ps_DmCd As String)
  Dim wSqlStr As String, wRsDsg As MDORowSet, wRow As Integer
  '****** Sachin 3.02 31-12-07 - Corresponding Partition
  '6.1
  wSqlStr = "Set DateFormat " + DtFmtStr() + " " & vbCrLf & "Select Top 5 IdCoCd, IdTc, IdYy, IdChr, IdNo, IdSr, InDt As IdDt, IdExpCmCd, IdKt, OdDmCol," & _
            " IdDmCd, IdSfx, IdDmSz, IdQty, (IdIFob/IdQty) As IdIUFOB, IdIFob" & _
            " from InvHd Join InvDsg  On InIdNo=IdInIdNo and InCoCd = IdCoCd And InTc = IdTc And InYy = IdYy And InChr = IdChr And InNo = IdNo And InPrtKey=IdPrtKey " & _
            " Join OrdDsg On IdPrtKey=OdPrtKey and IdOdIdNo=OdIdNo and InCoCd = OdCoCd And IdExpOdTc = OdTc And IdExpOdYy = OdYy And IdExpOdChr = OdChr And IdExpOdNo = OdNo And IdExpOdSr = OdSr" & _
            " Where InCoCd = '" & gs_CoCd & "' And InCmCd = '" & ADC("wIdCmCd") & "' And IdDmCd = '" & ps_DmCd & "' And InDt <= '" & CStr(ADC("wIdDt")) & "'" & _
            " And Not (InCoCd = '" & gs_CoCd & "' And InTc = '" & ADC("wIdTc") & "' And InYy = '" & ADC("wIdYy") & "' And InChr = '" & ADC("wIdChr") & "' And InNo = " & ADC("wIdNo") & ") And InPrtKey=(Case When '" + ADC("wIdPrtKey") + "'='" + ctCurrPrtn + "' Then '" + ctCurrPrtn + "' Else InPrtKey End) " & _
            " Order By InDt Desc, InTc, InYy, InChr, InNo Desc, IdSr Desc"
  Set wRsDsg = moCn.OpenResultset(wSqlStr)
  If wRsDsg.RecCount > 0 Then
    
    wRsDsg.MoveFirst
    With GrdHD
    .Rows = 1
      Do While Not (wRsDsg.BOF Or wRsDsg.EOF)
        .AddItem
        wRow = .Rows - 1
        .Value(wRow, "whDSr") = .MaxVal("WHDSr") + 1
        .Value(wRow, "whIdTc") = wRsDsg.FldValue("IdTc")
        .Value(wRow, "whIdYy") = wRsDsg.FldValue("IdYy")
        .Value(wRow, "whIdChr") = wRsDsg.FldValue("IdChr")
        .Value(wRow, "whIdNo") = wRsDsg.FldValue("IdNo")
        .Value(wRow, "whIdSr") = wRsDsg.FldValue("IdSr")
        .Value(wRow, "whIdDt") = wRsDsg.FldValue("IdDt")
        .Value(wRow, "whDmCd") = wRsDsg.FldValue("IdDmCD")
        .Value(wRow, "whDmSfx") = wRsDsg.FldValue("IdSfx")
        .Value(wRow, "whDmSz") = wRsDsg.FldValue("IdDmSz")
        .Value(wRow, "whIdKT") = wRsDsg.FldValue("IdKt")
        .Value(wRow, "whIdCol") = wRsDsg.FldValue("OdDmCol")
        .Value(wRow, "whIdQty") = wRsDsg.FldValue("IdQty")
        .Value(wRow, "whIdRt") = wRsDsg.FldValue("IdIUFob")
        .Value(wRow, "whIdVal") = wRsDsg.FldValue("IdIFob")
        wRsDsg.MoveNext
      Loop
      .Row = 1
    End With
  Else
    DispMsg "No History Found For This Design and Customer.", etError
    Call DispFra(SalHst)
  End If
End Sub
' Manoj 2.10.0

Private Sub DispFra(ByVal pv_InvDsgFra As en_InvDsgFra)
  '*** This is the common procedure which is used to toggle between hiding and displaying
      'the frames (i.e. if the frame is visible then make it invisible and viseversa)
      'associated with a command button
  
  Select Case pv_InvDsgFra
  Case Is = DsgDet
    If FraDD.Visible = True Then
      FraDD.Visible = False
      '*** required to pass the array position DsgDet-1, to resotre the 'Causes Validation' State
      Call EnaDisaCmds(False, CmdDD, DsgDet)
      GrdID.SetFocus
    Else
      FraDD.Visible = True
      ADC("wDdIdVaCtg").SetFocus
      FraDD.ZOrder
      Call EnaDisaCmds(True, CmdDD)
    End If
  Case Is = DsgCpy
    If FraDC.Visible = True Then
      FraDC.Visible = False
      '*** required to pass the array position DsgCpy-2, to resotre the 'Causes Validation' State
      Call EnaDisaCmds(False, CmdDC, DsgCpy)
      GrdID.SetFocus
    Else
      FraDC.Visible = True
      ADC("wDcOdTcFr").SetFocus
      FraDC.ZOrder
      Call EnaDisaCmds(True, CmdDC)
    End If
  Case Is = InvLab
    If FraIL.Visible = True Then
      FraIL.Visible = False
      '*** required to pass the array position InvLab-3, to resotre the 'Causes Validation' State
      Call EnaDisaCmds(False, CmdIL, InvLab)
      CmdIL.SetFocus
    Else
      FraIL.Visible = True
      GrdIL.SetFocus
      FraIL.ZOrder
      Call EnaDisaCmds(True, CmdIL)
    End If
  Case Is = InvBag
    If FraIB.Visible = True Then
      FraIB.Visible = False
      '*** required to pass the array position InvBag-4, to resotre the 'Causes Validation' State
      Call EnaDisaCmds(False, CmdIB, InvBag)
      CmdIB.SetFocus
    Else
      FraIB.Visible = True
      GrdIB.SetFocus
      FraIB.ZOrder
      Call EnaDisaCmds(True, CmdIB)
    End If
  Case Is = Summ
    If FraSumm.Visible = True Then
      FraSumm.Visible = False
      '*** required to pass the array position Summ-5, to resotre the 'Causes Validation' State
      Call EnaDisaCmds(False, CmdSumm, Summ)
      CmdSumm.SetFocus
    Else
      FraSumm.Visible = True
      TabSumm.SetFocus
      TabSumm.Tab = 0
      FraTabSumm(TabSumm.Tab).Enabled = True
      FraSumm.ZOrder
      Call EnaDisaCmds(True, CmdSumm)
    End If
' Manoj 2.10.0
  Case Is = SalHst
    If FraSalHst.Visible = True Then
      FraSalHst.Visible = False
      '*** required to pass the array position Summ-5, to resotre the 'Causes Validation' State
      Call EnaDisaCmds(False, CmdSH, SalHst)
      GrdID.SetFocus
    Else
      FraSalHst.Visible = True
      GrdHD.SetFocus
      FraSalHst.ZOrder
      Call EnaDisaCmds(True, CmdSH)
    End If
' Manoj 2.10.0
  ' ***** Manali 3.6.0 - 29/10/09 - Refresh Metal Rate
  Case Is = RefRt
    If FraRefRt.Visible = True Then
      FraRefRt.Visible = False
      '*** required to pass the array position RefRt, to resotre the 'Causes Validation' State
      Call EnaDisaCmds(False, CmdRefRt, RefRt)
      GrdID.SetFocus
    Else
      FraRefRt.Visible = True
      ADC("wInRefreshDt").SetFocus
      FraRefRt.ZOrder
      Call EnaDisaCmds(True, CmdRefRt)
    End If
  ' ***** Manali 3.6.0 - 29/10/09 - Refresh Metal Rate
  Case Is = GetLotRt
    If FraGetLotRt.Visible = True Then
      FraGetLotRt.Visible = False
      '*** required to pass the array position RefRt, to resotre the 'Causes Validation' State
      Call EnaDisaCmds(False, cmdGetLotRt, GetLotRt)
      GrdID.SetFocus
    Else
      FraGetLotRt.Visible = True
      ADC("WGETLOTRTONLYCUST").SetFocus
      FraGetLotRt.ZOrder
      Call EnaDisaCmds(True, cmdGetLotRt)
    End If
    '18
    Case Is = GetFOB
    If FraGetFOB.Visible = True Then
      FraGetFOB.Visible = False
      Call EnaDisaCmds(False, cmdGetFOB, GetFOB)
      cmdGetFOB.SetFocus
    Else
      FraGetFOB.Visible = True
      CmdGetFOBGo.SetFocus
      FraGetFOB.ZOrder
      Call EnaDisaCmds(True, cmdGetFOB)
    End If
    '6.1_37 visible/invisible of scope frame
    Case Is = Scope
    If FraSC.Visible = True Then
      FraSC.Visible = False
      Call EnaDisaCmds(False, CmdSC, Scope)
      CmdSC.SetFocus
    Else
      FraSC.Visible = True
      ADC("wScDmCtgFr").SetFocus
      FraSC.ZOrder
      Call EnaDisaCmds(True, CmdSC)
    End If
      
  End Select
End Sub
Private Sub EnaDisaCmds(ByVal pv_ShowFra As Boolean, Optional ByVal pv_CmdBut As MWCTL_BTN1, Optional ByVal pv_InvDsgFra As en_InvDsgFra)
  '*** This procedure is used to give the toggle option (enable and disable) to
      'the command buttons

  '*** storing the state of cmd buttons, when the frame is going to be vis.
  Dim X As Integer
  If pv_ShowFra = True Then
    Call StoreState
  End If

  '*** (Jenny Color)
  If Not (pv_CmdBut Is Nothing) And pv_ShowFra = True Then
    ms_CurrentBtn = UCase(pv_CmdBut.Name)
  Else
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
        ' pv_CmdBut.CausesValidation = False
        '*** restoring the 'Causes Validation' state of the current pv_cmdbutton
        If Not IsMissing(pv_InvDsgFra) Then
          pv_CmdBut.CausesValidation = mArr_Cv(pv_InvDsgFra)
        End If
    End If
  End If
  FraNKeyAll.Enabled = Not pv_ShowFra
  ADC.AllowSave = Not pv_ShowFra
'  ADC.AllowDelete = Not pv_ShowFra

  If pv_ShowFra = True Then
    '** All command buttons are disabled
    CmdMI.Enabled = Not pv_ShowFra
    CmdDD.Enabled = Not pv_ShowFra
    CmdDC.Enabled = Not pv_ShowFra
    CmdCRL.Enabled = Not pv_ShowFra
    CmdDRL.Enabled = Not pv_ShowFra   '****** Sachin 2.13.0 - 27-05-2006 ******
    CmdIL.Enabled = Not pv_ShowFra
    CmdIB.Enabled = Not pv_ShowFra
    CmdSumm.Enabled = Not pv_ShowFra
    CmdSH.Enabled = Not pv_ShowFra
    CmdRefRt.Enabled = Not pv_ShowFra    ' ***** Manali 3.6.0 - 29/10/09 - Refresh Metal Rate
    cmdGetLotRt.Enabled = Not pv_ShowFra    ' ***** Manali 3.6.0 - 29/10/09 - Refresh Metal Rate
    CmdInvDel.Enabled = Not pv_ShowFra   '4.1.2.0
    cmdGetFOB.Enabled = Not pv_ShowFra    '18
    CmdSC.Enabled = Not pv_ShowFra    '6.1_37
  Else
    '*** Restoring Ena/Disa state of all Command Buttons
    CmdMI.Enabled = mArr_Ed(MakeInv)
    CmdDD.Enabled = mArr_Ed(DsgDet)
    CmdDC.Enabled = mArr_Ed(DsgCpy)
    CmdCRL.Enabled = mArr_Ed(CpyRmLab)
    CmdDRL.Enabled = mArr_Ed(DelRmLab)    '****** Sachin 2.13.0 - 27-05-2006 ******
    CmdInvDel.Enabled = mArr_Ed(DelInv)     '4.1.2
    CmdIL.Enabled = mArr_Ed(InvLab)
    CmdIB.Enabled = mArr_Ed(InvBag)
    CmdSumm.Enabled = mArr_Ed(Summ)
' Manoj 2.10.0
    CmdSH.Enabled = mArr_Ed(SalHst)
' Manoj 2.10.0
    CmdRefRt.Enabled = mArr_Ed(RefRt)     ' ***** Manali 3.6.0 - 29/10/09 - Refresh Metal Rate
    cmdGetLotRt.Enabled = mArr_Ed(GetLotRt)     ' ***** Manali 3.6.0 - 29/10/09 - Refresh Metal Rate
    cmdGetFOB.Enabled = mArr_Ed(GetFOB)   '18
    CmdSC.Enabled = mArr_Ed(Scope)   '6.1_37
  End If
  If Not pv_CmdBut Is Nothing Then pv_CmdBut.Enabled = True
End Sub
Private Sub HideAllFras()
  '*** This procedure hides all the frames other than FraNKeyAll
  FraNKeyAll.Visible = True
  FraDD.Visible = False
  FraDC.Visible = False
  FraIL.Visible = False
  FraIB.Visible = False
  FraSumm.Visible = False
' Manoj 2.10.0
  FraSalHst.Visible = False
' Manoj 2.10.0
  FraRefRt.Visible = False    ' ***** Manali 3.6.0 - 29/10/09 - Refresh Metal Rate
  FraGetLotRt.Visible = False
  FraGetFOB.Visible = False   '18 GetFOb frame added
End Sub
Private Sub StoreState()
  '*** This procedure is used to store the CausesValidation State and the Enabled/Disabled
      'states of all the Command Buttons
  '*** This is used for restoring the original states in the future

  mArr_Cv(DsgDet) = CmdDD.CausesValidation
  mArr_Cv(DsgCpy) = CmdDC.CausesValidation
  mArr_Cv(CpyRmLab) = CmdCRL.CausesValidation
  mArr_Cv(DelRmLab) = CmdDRL.CausesValidation   '****** Sachin 2.13.0 - 27-05-2006 ******
  mArr_Cv(DelInv) = CmdInvDel.CausesValidation '4.1.2
  mArr_Cv(InvLab) = CmdIL.CausesValidation
  mArr_Cv(InvBag) = CmdIB.CausesValidation
  mArr_Cv(Summ) = CmdSumm.CausesValidation
' Manoj 2.10.0
  mArr_Cv(SalHst) = CmdSH.CausesValidation
' Manoj 2.10.0
  mArr_Cv(RefRt) = CmdSH.CausesValidation    ' ***** Manali 3.6.0 - 29/10/09 - Refresh Metal Rate
  mArr_Cv(GetLotRt) = CmdSH.CausesValidation    ' ***** Manali 3.6.0 - 29/10/09 - Refresh Metal Rate
  mArr_Ed(DsgDet) = CmdDD.Enabled
  mArr_Ed(DsgCpy) = CmdDC.Enabled
  mArr_Ed(CpyRmLab) = CmdCRL.Enabled
  mArr_Ed(DelRmLab) = CmdDRL.Enabled    '****** Sachin 2.13.0 - 27-05-2006 ******
  mArr_Ed(DelInv) = CmdInvDel.Enabled   '4.1.2
  mArr_Ed(InvLab) = CmdIL.Enabled
  mArr_Ed(InvBag) = CmdIB.Enabled
  mArr_Ed(Summ) = CmdSumm.Enabled
' Manoj 2.10.0
  mArr_Ed(SalHst) = CmdSH.Enabled
' Manoj 2.10.0
  mArr_Ed(RefRt) = CmdRefRt.Enabled   ' ***** Manali 3.6.0 - 29/10/09 - Refresh Metal Rate
  mArr_Ed(GetLotRt) = cmdGetLotRt.Enabled   ' ***** Manali 3.6.0 - 29/10/09 - Refresh Metal Rate
  mArr_Ed(GetFOB) = cmdGetFOB.Enabled   '18
End Sub
Private Function GetZAddVal(ByVal pf_i As Single) As Double
  '*** (Jen 2.12)
  Dim wf_AddLab As Double, wf_RmZVal As Double, wf_LZVal As Double
  Dim wi_i As Integer ', wf_RmZVal1 As Double, wf_LZVal1 As Double
  'wf_RmZVal = moCn.GetFldVal("Select sum(IrRmZVal + IrSetZVal) from InvRm " + _
  '            "where IrCoCd= '" + gs_CoCd + "' and " + _
  '            "IrTc= '" + adc("wIdTc") + "' and IrYy= '" + adc("wIdYy") + "' and " + _
  '            "IrChr= '" + adc("wIdChr") + "' and IrNo= " + CStr(adc("wIdNo")) + " and " + _
  '            "IrSr= " + CStr(GrdID.Value(pf_i, "IdSr")))
  'wf_LZVal = moCn.GetFldVal("Select sum(IlZVal) from InvLab " + _
  '            "where IlCoCd= '" + gs_CoCd + "' and " + _
  '            "IlTc= '" + adc("wIdTc") + "' and IlYy= '" + adc("wIdYy") + "' and " + _
  '            "IlChr= '" + adc("wIdChr") + "' and IlNo= " + CStr(adc("wIdNo")) + " and " + _
  '            "IlSr= " + CStr(GrdID.Value(pf_i, "IdSr")) + " and IlMCd<> '" + ctAddLMCd + "'")

  wf_RmZVal = 0
  For wi_i = 1 To GrdIR.Rows - 1
    wf_RmZVal = wf_RmZVal + IIF(GrdIR.IsDel(wi_i), 0, GrdIR.Value(wi_i, "IrRmZVal") + GrdIR.Value(wi_i, "IrSetZVal"))
  Next wi_i
  wf_LZVal = 0
  For wi_i = 1 To GrdIL.Rows - 1
    wf_LZVal = wf_LZVal + IIF(GrdIL.IsDel(wi_i) Or GrdIL.Value(wi_i, "IlMCd") = ctAddLMCd, 0, GrdIL.Value(wi_i, "IlZVal"))
  Next wi_i
  
  'Debug.Print "Inv FOB: " & GrdID.Value(pf_i, "IdIFOB") & "  Rm ZVal: " & wf_RmZVal & "  Lab ZVal: " & wf_LZVal
  wf_AddLab = GrdID.Value(pf_i, "IdIFOB") - (wf_RmZVal + wf_LZVal + GrdID.Value(pf_i, "IdZAdd"))
  'DispMsg "FOB: " & GrdID.Value(pf_i, "IdIFOB") & " RmVal: " & wf_RmZVal & " LabVal: " & wf_LZVal & " AddVal: " & GrdID.Value(pf_i, "IdZAdd"), etError
  'DispMsg "FOB: " & GrdID.Value(pf_i, "IdIFOB") & " LabVal: " & wf_LZVal & " LabVal1: " & wf_LZVal1 & " AddVal: " & GrdID.Value(pf_i, "IdZAdd"), etError
  GetZAddVal = wf_AddLab
  '*** (Jen 2.12)
End Function
Private Sub SetZAddLab(ByVal pf_IdRow As Single)
  '*** (Jen 2.12) 02/02/06
  If pf_IdRow = 0 Or Not (GrdIL.Mode = fgmnorm Or GrdIL.Mode = fgmdel) Then Exit Sub
  Dim wi_i As Integer
  '*** (02/02/06)
  Dim wd_AddLabVal As Double
  wd_AddLabVal = 0
  '*** (02/02/06)
  
  'If (GrdIL.Mode = fgmnorm) Then
    For wi_i = 1 To (GrdIL.Rows - 1)
      If GrdIL.Value(wi_i, "iLMCd") = ctAddLMCd And GrdIL.Value(wi_i, "iLSCd") = ctAddLSCd Then
        '*** (02/02/06)
        If Not GrdIL.IsDel(wi_i) Then
          wd_AddLabVal = GetZAddVal(pf_IdRow)
          If wd_AddLabVal <> GrdIL.Value(wi_i, "iLZRt") Then GrdIL.Value(wi_i, "iLZRt") = wd_AddLabVal
          GrdIL.SaveRec (wi_i)
        End If
        Exit For
        '*** (02/02/06)
        '*** (Bef 02/02/06)
        'GrdIL.Value(wi_i, "iLZRt") = GetZAddVal(pf_IdRow)
        'If Not GrdIL.IsDel(wi_i) Then GrdIL.SaveRec (wi_i)
        '*** (Bef 02/02/06)
      End If
    Next wi_i
  'End If
  '*** (Jen 2.12) 02/02/06
End Sub
Private Function GetDCRmVal(ByVal fs_IdName As String, Optional ByVal fi_DelRow As Variant, Optional ByVal fs_RmCtg As String, Optional ByVal fs_SubLogic As String) As Double
  ' ****** Manali 3.5.0 - new optional parameter fs_RmCtg added for "DHC" and "CHC"
  '*** (Jen 2.12) 06/01/06
  Dim wIrMainMetRmCtg As String
  wIrMainMetRmCtg = GetOrRmCtgMainMet
  
  Dim ws_RmSCtg As String, wChnLabAs As String                                   '************Bhavna added for CHN lab val calcn
  Dim wf_RmVal As Double, wi_i As Integer, ws_RmCtg As String, wi_DelRow As Integer
  Dim wb_CustRm As Boolean, ws_RmCd As String
  
  wf_RmVal = 0
  If ADC("wChnCustYn") = "Y" Then
    wChnLabAs = "," + moCn.GetFldVal("Select hChnSubCtgForLab from Head where hCoCd = 'ZZZ'") + ","        '************Bhavna for chain labour gold wt calc
  Else
    wChnLabAs = ""
  End If
  
  '*** (08/02/06)
  If IsMissing(fi_DelRow) Then
    wi_DelRow = 0
  Else
    wi_DelRow = fi_DelRow
  End If
  '*** (08/02/06)
  For wi_i = 1 To GrdIR.Rows - 1
    '****** Sachin 3.02 31-12-07 - PrtKey added
    ws_RmCtg = moCn.GetFldVal("Select RmCtg From RmMst where RmCd= '" + GrdIR.Value(wi_i, "IrRmCd") + "' and RmPrtKey=(Case When '" + ADC("wIdPrtKey") + "'='" + ctCurrPrtn + "' Then '" + ctCurrPrtn + "' Else RmPrtKey End) ")
    ws_RmSCtg = moCn.GetFldVal("Select RmSCtg From RmMst where RmCd= '" + GrdIR.Value(wi_i, "IrRmCd") + "' and RmPrtKey=(Case When '" + ADC("wIdPrtKey") + "'='" + ctCurrPrtn + "' Then '" + ctCurrPrtn + "' Else RmPrtKey End) ")
    
    ws_RmCd = GrdIR.Value(wi_i, "IrRmCd")
    If moCn.RecSeek("Select 'x' from CustMst where CmCtg='C' and CmCd = (select RmCmCd from RmMst where RmCd = '" + ws_RmCd + "')") Then
        wb_CustRm = True
    Else
        wb_CustRm = False
    End If
    
    ' ****** Manali 3.5.0 - 18/05/09 - Change for "DHC" and "CHC"
    'If Not IsMissing(fs_RmCtg) Then
    If fs_RmCtg <> "" Then
      '*****************Bhavna added chain labour value
      If fs_RmCtg = "CHN" And ADC("wChnCustYn") = "Y" And (InStr(1, wChnLabAs, "," + ws_RmSCtg + ",") <> 0) And Not (GrdIR.IsDel(wi_i)) And wi_i <> wi_DelRow Then
      '  And ws_RmSCtg = "CHN"
        wf_RmVal = wf_RmVal + GrdIR.Value(wi_i, fs_IdName)    '******Bhavna - tot of chain lab as subctg wts
      '*****************Bhavna added chain labour value
      'ElseIf fs_RmCtg = ctCDWLMCd And ws_RmCtg = "D" And Not (GrdIR.IsDel(wi_i)) Then
       ElseIf fs_RmCtg = ctCDWLMCd And ws_RmCtg = "D" And Not (GrdIR.IsDel(wi_i)) Then
        If fs_SubLogic = "C" And GetSubShpFrOrd(wi_i) = "C" Then
          wf_RmVal = wf_RmVal + GrdIR.Value(wi_i, fs_IdName)
        ElseIf fs_SubLogic = "-" Then
          wf_RmVal = wf_RmVal + GrdIR.Value(wi_i, fs_IdName)
        ElseIf fs_SubLogic = "NC" And GetSubShpFrOrd(wi_i) <> "C" Then
           wf_RmVal = wf_RmVal + GrdIR.Value(wi_i, fs_IdName)
        End If
       ' wf_RmVal = wf_RmVal + GrdOrdRm.Value(wi_i, ws_IdName)
     
        
        ' *** For DHC and CHC
      ElseIf fs_RmCtg = ctGHCLMCd And fs_SubLogic = "K" And (moCn.GetFldVal("Select RmCtg From RmMst Where RmCd='" + GrdIR.Value(wi_i, "IrRmCd") + "'") = wIrMainMetRmCtg And InStr(1, wChnLabAs, "," + ws_RmSCtg + ",") = 0) Then
        wf_RmVal = wf_RmVal + GrdIR.Value(wi_i, fs_IdName)
      ElseIf fs_RmCtg = ctGHCLMCd And fs_SubLogic = "P" And (moCn.GetFldVal("Select RmCtg From RmMst Where RmCd='" + GrdIR.Value(wi_i, "IrRmCd") + "'") = wIrMainMetRmCtg And InStr(1, wChnLabAs, "," + ws_RmSCtg + ",") = 0) Then
        If GrdIR.Value(wi_i, "IrMainMet") = "Y" Then
          wf_RmVal = wf_RmVal + GrdID.Value(GrdID.Row, IIF(fs_IdName = "IrRmIWt", "IdIGldAsWt", IIF(fs_IdName = "IrRmAWt", "IdAGldAsWt", "IdZGldAsWt")))
        Else
          wf_RmVal = wf_RmVal + GrdIR.Value(wi_i, fs_IdName)
        End If
      'End If
      
  
      ElseIf ws_RmCtg = fs_RmCtg And Not (GrdIR.IsDel(wi_i)) And wi_i <> wi_DelRow And (Not mb_CHCDHCOnCustRm Or wb_CustRm) Then
        wf_RmVal = wf_RmVal + GrdIR.Value(wi_i, fs_IdName)
        ' *** For DHC and CHC
      End If
'      ElseIf ws_RmCtg = fs_RmCtg And Not (GrdIR.IsDel(wi_i)) And wi_i <> wi_DelRow And (Not mb_CHCDHCOnCustRm Or wb_CustRm) Then _
'        wf_RmVal = wf_RmVal + GrdIR.Value(wi_i, fs_IdName)
'      End If
    Else
    ' ****** Manali 3.5.0 - 18/05/09 - Change for "DHC" and "CHC"
      If (ws_RmCtg = "C" Or ws_RmCtg = "D") And Not (GrdIR.IsDel(wi_i)) And wi_i <> wi_DelRow Then _
        wf_RmVal = wf_RmVal + GrdIR.Value(wi_i, fs_IdName)
    End If
  Next wi_i
  If fs_RmCtg = ctGHCLMCd And fs_SubLogic = "P" Then
    wf_RmVal = wf_RmVal * GetOrRmCtgMainMetVal()
  End If
  
  GetDCRmVal = wf_RmVal
  '*** (Jen 2.12) 06/01/06
End Function
Private Function GetOrRmCtgMainMetVal(Optional ps_CstSalRt As String) As Double
  Dim i As Integer
  If IsMissing(ps_CstSalRt) Or ps_CstSalRt = "" Then ps_CstSalRt = "S"
  Dim wMainRmCtg As String
  Dim wLme As Double
  For i = 1 To GrdIR.Rows - 1
    If Not GrdIR.IsDel(i) Then
      If GrdIR.Value(i, "IrMainMet") = "Y" Then
        wMainRmCtg = moCn.GetFldVal("Select RmCtg From RmMst Where RmCd='" + GrdIR.Value(i, "IrRmCd") + "'")
            Dim wf_BaseLmeOrd As Double
            Dim wf_baselmeinv As Double
            Dim ws_GldRtFromOrd As String
                     wf_BaseLmeOrd = moCn.GetFldVal("Select (case '" + wMainRmCtg + "' when 'G' then OmLmgSal " + _
                          "when 'P' then OmLmpSal when 'S' then OmLmsSal when 'L' then OmLmlSal else 0 end) from OrdMst where " + _
                          "OmCoCd= '" + gs_CoCd + "' and OmTc= '" + GrdID.Value(GrdID.Row, "IdExpOdTc") + "' " + _
                          "and OmYy= '" + GrdID.Value(GrdID.Row, "IdExpOdYy") + "' " + _
                          "and OmChr= '" + GrdID.Value(GrdID.Row, "IdExpOdChr") + "' " + _
                          "and OmNo= " + CStr(GrdID.Value(GrdID.Row, "IdExpOdNo")) + " and OmPrtKey='" + ctCurrPrtn + "' ")
                                            
                    wf_baselmeinv = moCn.GetFldVal("Select (case '" + wMainRmCtg + "' when 'G' then " + CStr(ADC("wIdLmgSal")) + " " + _
                          "when 'P' then " + CStr(ADC("wIdLmpSal")) + " when 'S' then " + CStr(ADC("wIdLmsSal")) + " " + _
                          "when 'L' then " + CStr(ADC("wIdLmlSal")) + " else 0 end) ")
        
                     ws_GldRtFromOrd = moCn.GetFldVal("Select OdIGldRtEqOrd from OrdDsg where " + _
                          "OdCoCd= '" + gs_CoCd + "' and OdTc= '" + GrdID.Value(GrdID.Row, "IdExpOdTc") + "' " + _
                          "and OdYy= '" + GrdID.Value(GrdID.Row, "IdExpOdYy") + "' " + _
                          "and OdChr= '" + GrdID.Value(GrdID.Row, "IdExpOdChr") + "' " + _
                          "and OdNo= " + CStr(GrdID.Value(GrdID.Row, "IdExpOdNo")) + " and OdSr=" + CStr(GrdID.Value(GrdID.Row, "IdExpOdSr")) + "  and OdPrtKey='" + ctCurrPrtn + "' ")
        
       If ws_GldRtFromOrd = "Y" Then wLme = wf_BaseLmeOrd
        If ws_GldRtFromOrd = "N" Then wLme = wf_baselmeinv
        If ws_GldRtFromOrd = "M" Then
      wLme = IIF(wf_BaseLmeOrd > wf_baselmeinv, wf_BaseLmeOrd, wf_baselmeinv)
       End If
        GetOrRmCtgMainMetVal = GetGldRtForLabCdPure(ADC("wIdCmCd"), wLme, GetOrRmCd, 0, "", "", ps_CstSalRt, "", , ADC("wIdCmCurCd"), "", "", "")
        
        Exit For
      End If
    End If
  Next i

End Function

Private Function GetOrRmCd() As String
  Dim i As Integer
  For i = 1 To GrdIR.Rows - 1
    If Not GrdIR.IsDel(i) Then
      If GrdIR.Value(i, "IrMainMet") = "Y" Then
        GetOrRmCd = GrdIR.Value(i, "IrRmCd")
        Exit For
      End If
    End If
  Next i

End Function
Private Function GetOrRmCtgMainMet() As String
  Dim i As Integer
  For i = 1 To GrdIR.Rows - 1
    If Not GrdIR.IsDel(i) Then
      If GrdIR.Value(i, "IrMainMet") = "Y" Then
        GetOrRmCtgMainMet = moCn.GetFldVal("Select RmCtg From RmMst Where RmCd='" + GrdIR.Value(i, "IrRmCd") + "'")
        Exit For
      End If
    End If
  Next i

End Function
Private Function GetSubShpFrOrd(IrRownum As Integer) As String
  Dim wOrLnStr As String
      wOrLnStr = IIF(moCn.GetFldVal("select PValue3 From Param Join  RmMst " + _
                            " on PTyp= 'RMSCTG' and PMCd= RmCtg and PSCd= RmSCtg " + _
                            "where RmCd= '" + GrdIR.Value(IrRownum, "IrRmCd") + "' and RmPrtKey='" + ctCurrPrtn + "' ") = "Y", _
                   IIF(moCn.GetFldVal("select PValue2 From Param Join RmMst " + _
                            "on PTyp= 'RMSCTG' and PMCd= RmCtg and PSCd= RmSCtg " + _
                            "where RmCd= '" + GrdIR.Value(IrRownum, "IrRmCd") + "' and RmPrtKey='" + ctCurrPrtn + "' ") = "Y", "OrRmPtr", _
                            " case when OrPrdQty =0 then 0 else OrPrdWt/OrPrdQty end "), "OrLn1")

  GetSubShpFrOrd = moCn.GetFldVal("Select max(OrSubShp) from OrdRm where " + _
              " OrCoCd ='" + gs_CoCd + "'" + _
              " and OrTc = '" + GrdID.Value(GrdID.Row, "IdExpOdTc") + "' " + _
              " and OrYy= '" + GrdID.Value(GrdID.Row, "IdExpOdYy") + "' " + _
              " and OrChr= '" + GrdID.Value(GrdID.Row, "IdExpOdChr") + "' " + _
              " and OrNo = " + CStr(GrdID.Value(GrdID.Row, "IdExpOdNo")) + _
              " and OrSr= " + CStr(GrdID.Value(GrdID.Row, "IdExpOdSr")) + _
              " and OrPrtKey= '" + ctCurrPrtn + "' " + _
              " and OrRmCd= '" + GrdIR.Value(IrRownum, "IrRmCd") + "' " + _
              " and " + wOrLnStr + " = " + CStr(GrdIR.Value(IrRownum, "IrRmSz")))
              
End Function
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
  
  Dim ws_BtnStr As String
  ws_BtnStr = "CmdDCGo"
  Call SetObjColors(Me, ws_BtnStr)
  '*** Jenny Color
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
    If Not m_ObjCmdGetRmRt_IPD Is Nothing Then m_ObjCmdGetRmRt_IPD.QryClose
    If Not m_objCmdGetRmRt_Base Is Nothing Then m_objCmdGetRmRt_Base.QryClose

    Set moCn = Nothing    '*** (09/08/05)    '*** (Bef speed) bef 24/11/06
End Sub
Private Sub GrdID_Validate(Cancel As Boolean)
  Cancel = GrdID.Validate
End Sub
Private Sub GrdIR_Validate(Cancel As Boolean)
  Cancel = GrdIR.Validate
End Sub
Private Sub GrdIL_Validate(Cancel As Boolean)
  Cancel = GrdIL.Validate
End Sub
Private Sub GrdIB_Validate(Cancel As Boolean)
  Cancel = GrdIB.Validate
End Sub
Private Sub GrdHD_Validate(Cancel As Boolean)
  Cancel = GrdHD.Validate
End Sub
Private Sub GrdHR_Validate(Cancel As Boolean)
  Cancel = GrdHR.Validate
End Sub
Private Sub GrdHL_Validate(Cancel As Boolean)
  Cancel = GrdHL.Validate
End Sub

Private Sub RefreshRt()
  Dim wi_D As Integer, wi_R As Integer
  Dim wIGldRtEqOrd As String, ws_RmCtg As String
  Dim wOmLkUpRmRt As String, wOmLkUpMetLs As String, wOmLkUpStr As String
  Dim wOmLkUpArr() As String
  Dim wStr As String
  Dim wBseCurCd As String, wCnvUpdStr As String
    
  '3.11.0
  wCnvUpdStr = ""
  If ADC("wRefCnvRtYn") = "Y" Then
    wBseCurCd = moCn.GetFldVal("Select HBaseCurCd from Head Where HCoCd='" + ctSelfCoCd + "' and HCd='" + ctSelfCmCd + "'")
    If ms_RtbyStkYN = "N" And ADC("wIdCmCurCd") = wBseCurCd Then _
        wCnvUpdStr = ", InCnvFct=" + CStr(ADC("wInCnvRt"))
  End If
    
  '3.11.0 - Conversion Rate added below
  wStr = " Update InvHd " + _
    " Set InRefreshDt='" + CStr(ADC("wInRefreshDt")) + "', " + _
    " InLmgDt='" + CStr(ADC("wInLmgDt")) + "', InLmgCst=" + CStr(ADC("wInLmgCst")) + ", InLmgSal=" + CStr(ADC("wInLmgSal")) + ", " + _
    " InLmpDt='" + CStr(ADC("wInLmpDt")) + "', InLmpCst=" + CStr(ADC("wInLmpCst")) + ", InLmpSal=" + CStr(ADC("wInLmpSal")) + ", " + _
    " InLmsDt='" + CStr(ADC("wInLmsDt")) + "', InLmsCst=" + CStr(ADC("wInLmsCst")) + ", InLmsSal=" + CStr(ADC("wInLmsSal")) + ", " + _
    " InLmlDt='" + CStr(ADC("wInLmlDt")) + "', InLmlCst=" + CStr(ADC("wInLmlCst")) + ", InLmlSal=" + CStr(ADC("wInLmlSal")) + " " + wCnvUpdStr + _
    " Where InCoCd= '" + gs_CoCd + "'" + _
    " And InTc= '" + ADC("wIdTc") + "'" + _
    " And InYy= '" + ADC("wIdYy") + "'" + _
    " And InChr= '" + ADC("wIdChr") + "'" + _
    " And InNo= " + CStr(ADC("wIdNo")) + " And InPrtKey='" + ctCurrPrtn + "' "
  moCn.Execute (wStr)
        
  ADC("wIdLmgCst") = ADC("wInLmgCst")
  ADC("wIdLmgSal") = ADC("wInLmgSal")

  ADC("wIdLmpCst") = ADC("wInLmpCst")
  ADC("wIdLmpSal") = ADC("wInLmpSal")

  ADC("wIdLmsCst") = ADC("wInLmsCst")
  ADC("wIdLmsSal") = ADC("wInLmsSal")

  ADC("wIdLmlCst") = ADC("wInLmlCst")
  ADC("wIdLmlSal") = ADC("wInLmlSal")
  
  '3.11.0
  If wCnvUpdStr <> "" Then ADC("wIdCnvFct") = ADC("wInCnvRt")

  For wi_D = 1 To GrdID.Rows - 1
    GrdID.Row = wi_D
    wIGldRtEqOrd = moCn.GetFldVal("Select OdIGldRtEqOrd from OrdDsg where " + _
                            "OdCoCd= '" + gs_CoCd + "' and OdTc= '" + GrdID.Value(wi_D, "IdExpOdTc") + "' " + _
                            "and OdYy= '" + GrdID.Value(wi_D, "IdExpOdYy") + "' " + _
                            "and OdChr= '" + GrdID.Value(wi_D, "IdExpOdChr") + "' " + _
                            "and OdNo= " + CStr(GrdID.Value(wi_D, "IdExpOdNo")) + " " + _
                            "and OdSr= " + CStr(GrdID.Value(wi_D, "IdExpOdSr")) + " and OdPrtKey='" + ctCurrPrtn + "' ")

    
    If GrdID.Value(wi_D, "IdPrtKey") = ctCurrPrtn And wIGldRtEqOrd <> "Y" Then
        
      wOmLkUpStr = moCn.GetFldVal("Select OmLkUpRmRt+','+OmLkUpMetLs from OrdMst where " + _
                                  "OmCoCd= '" + gs_CoCd + "' and OmTc= '" + GrdID.Value(wi_D, "IdExpOdTc") + "' " + _
                                  "and OmYy= '" + GrdID.Value(wi_D, "IdExpOdYy") + "' " + _
                                  "and OmChr= '" + GrdID.Value(wi_D, "IdExpOdChr") + "' " + _
                                  "and OmNo= " + CStr(GrdID.Value(wi_D, "IdExpOdNo")) + " and OmPrtKey='" + ctCurrPrtn + "' ")
      If wOmLkUpStr <> "" Then
        wOmLkUpArr() = Split(wOmLkUpStr, ",")
        wOmLkUpRmRt = wOmLkUpArr(0)
        wOmLkUpMetLs = wOmLkUpArr(1)
      End If
      Dim wCmCd As String
      If ms_RtFrInvCust = "Y" Then
        wCmCd = ADC("wIdCmCd") '6.1-55
        wOmLkUpRmRt = ""
        wOmLkUpMetLs = ""
      Else
        wCmCd = GrdID.Value(wi_D, "IdExpCmCd")
      End If

      With GrdIR
        For wi_R = 1 To .Rows - 1
          If .Value(wi_R, "IrPrtKey") = ctCurrPrtn Then
            ws_RmCtg = moCn.GetFldVal("Select RmCtg From RmMst Where RmCd='" + .Value(wi_R, "IrRmCd") + "' and RmPrtKey='" + ctCurrPrtn + "'")
  
            If ws_RmCtg = "G" Then
            '6.1-55 GrdID.Value(wi_D, "IdExpCmCd") changed as wCmCd
              .Value(wi_R, "IrRmIRt") = GetGldRt(wCmCd, ADC("wIdLmgSal"), _
                                      .Value(wi_R, "IrRmCd"), .Value(wi_R, "IrAlySalRt"), "DM", GrdID.Value(wi_D, "IdDmCd"), "S", "C", , ADC("wIdCmCurCd"), _
                                      IIF(wOmLkUpRmRt <> "", wOmLkUpRmRt, ""), IIF(wOmLkUpMetLs <> "", wOmLkUpMetLs, ""), , , .Value(wi_R, "IRRMSZ"), .Value(wi_R, "IRRMSZ2"), 0)
              .Value(wi_R, "IrRmZRt") = .Value(wi_R, "IrRmIRt")
              
            ElseIf ws_RmCtg = "P" Then
            '6.1-55 GrdID.Value(wi_D, "IdExpCmCd") changed as wCmCd
              .Value(wi_R, "IrRmIRt") = GetGldRt(wCmCd, ADC("wIdLmpSal"), _
                                      .Value(wi_R, "IrRmCd"), .Value(wi_R, "IrAlySalRt"), "DM", GrdID.Value(wi_D, "IdDmCd"), "S", "C", , ADC("wIdCmCurCd"), _
                                      IIF(wOmLkUpRmRt <> "", wOmLkUpRmRt, ""), IIF(wOmLkUpMetLs <> "", wOmLkUpMetLs, ""), , , .Value(wi_R, "IRRMSZ"), .Value(wi_R, "IRRMSZ2"), 0)
              .Value(wi_R, "IrRmZRt") = .Value(wi_R, "IrRmIRt")
            
            ElseIf ws_RmCtg = "S" Then '6.1-55 GrdID.Value(wi_D, "IdExpCmCd") changed as wCmCd
              If mSlvMod = True Then
                .Value(wi_R, "IrRmIRt") = GetGldRt(wCmCd, ADC("wIdLmsSal"), _
                                        .Value(wi_R, "IrRmCd"), .Value(wi_R, "IrAlySalRt"), "DM", GrdID.Value(wi_D, "IdDmCd"), "S", "C", , ADC("wIdCmCurCd"), _
                                        IIF(wOmLkUpRmRt <> "", wOmLkUpRmRt, ""), IIF(wOmLkUpMetLs <> "", wOmLkUpMetLs, ""), , , .Value(wi_R, "IRRMSZ"), .Value(wi_R, "IRRMSZ2"), 0)
              Else
                .Value(wi_R, "IrRmIRt") = 0
              End If
              .Value(wi_R, "IrRmZRt") = .Value(wi_R, "IrRmIRt")
              
            ElseIf ws_RmCtg = "L" Then
              If mb_PdMod = True Then '6.1-55 GrdID.Value(wi_D, "IdExpCmCd") changed as wCmCd
                .Value(wi_R, "IrRmIRt") = GetGldRt(wCmCd, ADC("wIdLmlSal"), _
                                        .Value(wi_R, "IrRmCd"), .Value(wi_R, "IrAlySalRt"), "DM", GrdID.Value(wi_D, "IdDmCd"), "S", "C", , ADC("wIdCmCurCd"), _
                                        IIF(wOmLkUpRmRt <> "", wOmLkUpRmRt, ""), IIF(wOmLkUpMetLs <> "", wOmLkUpMetLs, ""), , , .Value(wi_R, "IRRMSZ"), .Value(wi_R, "IRRMSZ2"), 0)
              Else
                .Value(wi_R, "IrRmIRt") = 0
              End If
              .Value(wi_R, "IrRmZRt") = .Value(wi_R, "IrRmIRt")
              
            ElseIf ws_RmCtg = "X" Then
              Dim ws_BaseRmCtg As String, wf_BaseLmeOrd As Single, wf_baselmeinv As Single

              ws_BaseRmCtg = moCn.GetFldVal("Select B.RmCtg from RmMst B where " + _
                             "RmCd= (Select Rm.RmBaseCd from RmMst Rm where Rm.RmCd= '" + .Value(wi_R, "IrRmCd") + "' and Rm.RmPrtKey='" + ctCurrPrtn + "' ) and RmPrtKey='" + ctCurrPrtn + "' ")
              
              wf_BaseLmeOrd = moCn.GetFldVal("Select (case '" + ws_BaseRmCtg + "' when 'G' then OmLmgSal " + _
                              "when 'P' then OmLmpSal when 'S' then OmLmsSal when 'L' then OmLmlSal else 0 end) from OrdMst where " + _
                              "OmCoCd= '" + gs_CoCd + "' and OmTc= '" + GrdID.Value(wi_D, "IdExpOdTc") + "' " + _
                              "and OmYy= '" + GrdID.Value(wi_D, "IdExpOdYy") + "' " + _
                              "and OmChr= '" + GrdID.Value(wi_D, "IdExpOdChr") + "' " + _
                              "and OmNo= " + CStr(GrdID.Value(wi_D, "IdExpOdNo")) + " and OmPrtKey='" + ctCurrPrtn + "' ")
                                                
              wf_baselmeinv = moCn.GetFldVal("Select (case '" + ws_BaseRmCtg + "' when 'G' then " + CStr(ADC("wIdLmgSal")) + " " + _
                              "when 'P' then " + CStr(ADC("wIdLmpSal")) + " when 'S' then " + CStr(ADC("wIdLmsSal")) + " " + _
                              "when 'L' then " + CStr(ADC("wIdLmlSal")) + " else 0 end) ")
              '6.1-55 GrdID.Value(wi_D, "IdExpCmCd") changed as wCmCd- 3places
              If GetRmRt(wCmCd, .Value(wi_R, "IrRmCd"), 0, 0, 0, "S", wf_BaseLmeOrd, , , , , ADC("wIdCmCurCd"), IIF(wOmLkUpRmRt <> "", wOmLkUpRmRt, "")) <> _
                 GetRmRt(wCmCd, .Value(wi_R, "IrRmCd"), 0, 0, 0, "S", wf_baselmeinv, , , , , ADC("wIdCmCurCd"), IIF(wOmLkUpRmRt <> "", wOmLkUpRmRt, "")) Then
                  .Value(wi_R, "IrRmIRt") = GetRmRt(wCmCd, .Value(wi_R, "IrRmCd"), 0, 0, 0, "S", wf_baselmeinv, , , , , ADC("wIdCmCurCd"), IIF(wOmLkUpRmRt <> "", wOmLkUpRmRt, ""))
                  .Value(wi_R, "IrRmZRt") = .Value(wi_R, "IrRmIRt")
              End If
              
            End If
          End If
          
          GrdIR.SaveRec (wi_R)
        Next wi_R
      End With
    End If
  Next wi_D

End Sub

Private Sub TxtPwd_GotFocus()
  DispMsg "Enter SEO Password", etInfo
End Sub
Private Sub TxtGetLotRtSEOPwd_GotFocus()
  DispMsg "Enter SEO Password", etInfo
End Sub
Private Sub TxtGetLotRtSEOPwd_Validate(Cancel As Boolean)

'  mf_DtLocked = IIF((ADC("wIdDt") < CDate(ms_LockDt)) _
'                               Or _
'                   (ADC("wIdDt") > CDate(ms_FutureDt)), _
'                True, False)
  'ms_PwdFlag = False
  If ADC.AddFndMode = xFndMode Then
    If Trim(txtGetLotRtSEOPwd.text) <> "" Then
      If Trim(txtGetLotRtSEOPwd.text) = UCase(moCn.GetFldVal("select IsNull((select PValue from Param Join UsrAccess on UaUsrCd = PMCd " + _
                   "where PTyp= 'USR' and PMCd= '" + gs_UsrCd + "' and UaMnuCd = '" + ADC.MenuCd + "' and " + _
                   "UaSeoYn= 'Y' and PValue= '" + Trim(txtGetLotRtSEOPwd.text) + "'), " + _
                   "(Select PValue1 from Param where Ptyp= 'SEOPWD' and PmCd= 'SEOPWD' )) ")) Then
          ms_PwdFlag = True: Exit Sub
      Else
          ms_PwdFlag = False: Cancel = True: DispMsg "Invalid Edit Password / User Does Not Have SEO Rights ", etError
      End If
    Else
        If moCn.GetFldVal("Select PValue1 from Param where Ptyp = 'SEOPWD' and PmCd = 'SEOPWD'") = "" Then _
           ms_PwdFlag = True Else ms_PwdFlag = False
    End If
  End If
End Sub
Private Sub TxtPwd_Validate(Cancel As Boolean)

'  mf_DtLocked = IIF((ADC("wIdDt") < CDate(ms_LockDt)) _
'                               Or _
'                   (ADC("wIdDt") > CDate(ms_FutureDt)), _
'                True, False)
  'ms_PwdFlag = False
  If ADC.AddFndMode = xFndMode Then
    If Trim(TxtPwd.text) <> "" Then
      If Trim(TxtPwd.text) = UCase(moCn.GetFldVal("select IsNull((select PValue from Param Join UsrAccess on UaUsrCd = PMCd " + _
                   "where PTyp= 'USR' and PMCd= '" + gs_UsrCd + "' and UaMnuCd = '" + ADC.MenuCd + "' and " + _
                   "UaSeoYn= 'Y' and PValue= '" + Trim(TxtPwd.text) + "'), " + _
                   "(Select PValue1 from Param where Ptyp= 'SEOPWD' and PmCd= 'SEOPWD' )) ")) Then
          ms_PwdFlag = True: Exit Sub
      Else
          ms_PwdFlag = False: Cancel = True: DispMsg "Invalid Edit Password / User Does Not Have SEO Rights ", etError
      End If
    Else
        If moCn.GetFldVal("Select PValue1 from Param where Ptyp = 'SEOPWD' and PmCd = 'SEOPWD'") = "" Then _
           ms_PwdFlag = True Else ms_PwdFlag = False
    End If
  End If
  
End Sub

Private Sub CmdInvDel_Click()
 '****** Sachin 3.12 - Deletion Not allowed if InvDsg belongs to the previous partition
 If ADC("wIdPrtKey") <> ctCurrPrtn Then DispMsg "Cannot Delete Rm/Lab From Invoice of Previous Partition ", etError: Exit Sub

 '****** Sachin 3.12.0 - Delete Invoice Button  ******
 If moCn.RecSeek("Select 'x' from InvZRm where IzCoCd= '" + gs_CoCd + "' and IzTc= '" + ADC("wIdTc") + "' and " + _
                  "IzYy=  '" + ADC("wIdYy") + "' and IzChr=  '" + ADC("wIdChr") + "' and " + _
                  "IzNo=  " + CStr(ADC("wIdNo")) + " and IzPrtKey='" + ctCurrPrtn + "' ") Then
    DispMsg "Cannot Delete RM/Labour As Custom Matching Has Been Done", etError: GrdID.SetFocus: Exit Sub
 End If

 If GrdID.Rows <= 0 Then DispMsg "Nothing To Delete.", etInfo: GrdID.SetFocus: Exit Sub
 If MsgBox("This will Completely DELETE the Invoice. Are U Sure U Want To Continue?", vbYesNo) = vbNo Then Exit Sub
 On Error GoTo InvDelErr
 Dim wInIdNo As Long
 wInIdNo = moCn.GetFldVal("Select InIdNo from InvHd where InCoCd= '" + gs_CoCd + "' and InTc= '" + ADC("wIdTc") + "' and " + _
                  "InYy=  '" + ADC("wIdYy") + "' and InChr=  '" + ADC("wIdChr") + "' and " + _
                  "InNo=  " + CStr(ADC("wIdNo")))
  BeginEmrProcess Me, "Deleting Invoice"
'  MWLib.BeginProcess Me, "Deleting Invoice"
 moCn.Execute ("Delete From InvDsg where IdInIdNo = '" + CStr(wInIdNo) + "'")
 
 GrdIR.Rows = 1
 GrdIL.Rows = 1
 GrdID.Rows = 1
 EndEmrProcess Me
 MWLib.EndProcess Me
 CmdInvDel.SetFocus
'****** Sachin 2.13.0 - Delete RM/Lab Button - 27-05-2006 ******
  Exit Sub
InvDelErr:
 EndEmrProcess Me
 MWLib.EndProcess Me
  MsgBox "Unable to delete Invoice, Error encountered." + Err.Description
  DispMsg "Unable to delete Invoice, Error encountered." + Err.Description, etError
  CmdInvDel.SetFocus
End Sub

'IdTc           IdYy           IdChr           IdNo
'IdSr           IdDmCd         IdSfx           IdDmSz
'IdQty          IdExpOdTc      IdExpOdYy       IdExpOdChr
'IdExpOdNo      IdExpOdSr      IdExpCmCd
'IdAVal         IdIVal         IdZVal
'IdIAdd         IdZAdd         IdIFob          IdZFob
'IdVaCtg        IdKt           IdGldLs         IdFixPrc
'IdHZValEqInv   IdGldAs        IdAGldAsWt     IdiGldAsWt
'IdZGldAsWt


'InvFgd
'
'IfTc     IfYy      IfChr     IfNo      IfSr      IfSrNo      IfFdTc
'IfFdYy   IfFdChr   IfFdNo    IfFdSr    IfFdBYy   IfFdBChr    IfFdBNo
'IfQty    IfGrsWt



'invrm
'
'IrTc           IrYy           IrChr           IrNo
'IrSr           IrSrNo         IrRmCd          IrRmQty
'IrRmAWt        IrRmIWt        IrRmZWt         IrRmARt
'IrRmIRt        IrRmZRt        IrRmAVal        IrRmIVal
'IrRmZVal       IrSetSCd       IrSetARt        IrSetIRt
'IrSetZRt       IrSetAVal      IrSetIVal       IrSetZVal
'IrMainMet      IrRmZCd



'invlab
'
'IlTc           IlYy           IlChr           IlNo
'IlSr           IlSrNo         IlMCd           IlSCd
'IlQty          IlARt          IlIRt           IlZRt
'IlAVal         IlIVal         IlZVal          IlQw





'WDRCTG
'WDRSCTG
'WDRCD
'WDRQTY
'WDRIWT
'WDRIVAL
'WDRZWT
'WDRZVAL
'
'WDRTOTQTY
'WDRTOTIWT
'WDRTOTIVAL
'WDRTOTZWT
'WDRTOTZVAL


'WDLMCD
'WDLSCD
'WDLQW
'WDLQTY
'WDLIVAL
'WDLZVAL
'
'WDLTOTIVAL
'WDLTOTZVAL


'WIDSR
'WIDDMCTG
'WIDDMCD
'WIDQTY
'WIDIPRC
'WIDIFOB
'WIDZPRC
'WIDZFOB
'
'WIDTOTQTY
'WIDTOTIFOB
'WIDTOTZFOB







