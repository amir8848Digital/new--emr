VERSION 5.00
Object = "{AB2CC32A-E652-42A2-B341-26ECE02F6386}#2.1#0"; "mwfctlforspeed9.ocx"
Begin VB.Form EmrFrmBag 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Bag Master"
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
   Begin MwfCtl.MWCTL_BTN1 CmdBagGen 
      Height          =   465
      Left            =   6870
      TabIndex        =   40
      ToolTipText     =   "Option To Generate Bags For An Order Sr Range"
      Top             =   9300
      Width           =   1065
      _ExtentX        =   1879
      _ExtentY        =   820
      BackColor       =   -2147483633
      MaskColor       =   8454143
      Caption         =   "&Bag Generation"
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
      Height          =   435
      Left            =   9990
      TabIndex        =   13
      TabStop         =   0   'False
      Top             =   9150
      Visible         =   0   'False
      Width           =   435
      _ExtentX        =   767
      _ExtentY        =   767
   End
   Begin MwfCtl.MWCTL_ADC ADC 
      CausesValidation=   0   'False
      Height          =   555
      Left            =   30
      TabIndex        =   14
      Top             =   9360
      Width           =   15135
      _ExtentX        =   26696
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
      Height          =   570
      Left            =   0
      TabIndex        =   16
      Top             =   24
      Width           =   12540
      Begin MwfCtl.MWCTL_MED ATXT 
         Height          =   285
         Index           =   7
         Left            =   5145
         TabIndex        =   5
         ToolTipText     =   "Enter From Design"
         Top             =   0
         Width           =   2370
         _ExtentX        =   4180
         _ExtentY        =   503
         MaxLength       =   15
         DataType        =   4
         IdName          =   "WBODFRDMCD"
         CmpStr          =   "BOdDmCd >="
      End
      Begin MwfCtl.MWCTL_MED ATXT 
         Height          =   285
         Index           =   5
         Left            =   9000
         TabIndex        =   7
         ToolTipText     =   "Sort Records On"
         Top             =   0
         Width           =   1410
         _ExtentX        =   2487
         _ExtentY        =   503
         MaxLength       =   8
         DataType        =   4
         IdName          =   "WBSRT"
      End
      Begin MwfCtl.MWCTL_MED ATXT 
         Height          =   285
         Index           =   1
         Left            =   1395
         TabIndex        =   1
         ToolTipText     =   "Enter Order Type"
         Top             =   0
         Width           =   570
         _ExtentX        =   1005
         _ExtentY        =   503
         MaxLength       =   3
         DataType        =   4
         ReCalcParent    =   "WBODCHR"
         IdName          =   "WBODTC"
         CmpStr          =   "BOdTc="
      End
      Begin MwfCtl.MWCTL_MED ATXT 
         Height          =   285
         Index           =   40
         Left            =   2070
         TabIndex        =   2
         ToolTipText     =   "Enter Order Year"
         Top             =   0
         Width           =   330
         _ExtentX        =   582
         _ExtentY        =   503
         MaxLength       =   2
         DataType        =   4
         IdName          =   "WBODYY"
         CmpStr          =   "BOdYy="
      End
      Begin MwfCtl.MWCTL_MED ATXT 
         Height          =   285
         Index           =   41
         Left            =   2520
         TabIndex        =   3
         ToolTipText     =   "Enter Order Character"
         Top             =   0
         Width           =   570
         _ExtentX        =   1005
         _ExtentY        =   503
         MaxLength       =   3
         DataType        =   4
         ReCalcOn        =   "WBODTC"
         IdName          =   "WBODCHR"
         CmpStr          =   "BOdchr="
      End
      Begin MwfCtl.MWCTL_MED ATXT 
         Height          =   285
         Index           =   42
         Left            =   3210
         TabIndex        =   4
         ToolTipText     =   "Enter Order No."
         Top             =   0
         Width           =   750
         _ExtentX        =   1323
         _ExtentY        =   503
         Alignment       =   2
         Mask            =   "#####0"
         MaxLength       =   6
         DataType        =   1
         IdName          =   "WBODNO"
         CmpStr          =   "BOdNo="
      End
      Begin MwfCtl.MWCTL_MED ATXT 
         Height          =   285
         Index           =   0
         Left            =   5145
         TabIndex        =   6
         ToolTipText     =   "Enter To Design"
         Top             =   285
         Width           =   2370
         _ExtentX        =   4180
         _ExtentY        =   503
         MaxLength       =   15
         DataType        =   4
         IdName          =   "WBODTODMCD"
         CmpStr          =   "BOdDmCd <="
      End
      Begin MwfCtl.MWCTL_MED ATXT 
         Height          =   285
         Index           =   14
         Left            =   9000
         TabIndex        =   8
         ToolTipText     =   "Enter 'Y' for Flute Bags, 'N' for Primary Bags and Blank for Both"
         Top             =   285
         Width           =   330
         _ExtentX        =   582
         _ExtentY        =   503
         DataType        =   4
         IdName          =   "WBAGTYP"
      End
      Begin MwfCtl.MWCTL_MED ATXT 
         Height          =   285
         Index           =   15
         Left            =   11655
         TabIndex        =   9
         ToolTipText     =   "Enter From Bag Character"
         Top             =   0
         Width           =   630
         _ExtentX        =   1111
         _ExtentY        =   503
         MaxLength       =   3
         DataType        =   4
         IdName          =   "WBCHRFR"
         CmpStr          =   "BChr>= "
      End
      Begin MwfCtl.MWCTL_MED ATXT 
         Height          =   285
         Index           =   16
         Left            =   11655
         TabIndex        =   10
         ToolTipText     =   "Enter To Bag Character"
         Top             =   285
         Width           =   630
         _ExtentX        =   1111
         _ExtentY        =   503
         MaxLength       =   3
         DataType        =   4
         IdName          =   "WBCHRTO"
         CmpStr          =   "BChr <="
      End
      Begin MwfCtl.MWCTL_MED ATXT 
         Height          =   285
         Index           =   22
         Left            =   750
         TabIndex        =   0
         ToolTipText     =   "Enter Company Code"
         Top             =   0
         Width           =   570
         _ExtentX        =   1005
         _ExtentY        =   503
         MaxLength       =   3
         DataType        =   4
         ReCalcParent    =   "WBODCHR"
         IdName          =   "WBCOCD"
         CmpStr          =   "BCoCd="
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
         Left            =   1320
         TabIndex        =   73
         Top             =   0
         Width           =   105
      End
      Begin VB.Label ALBL 
         Alignment       =   1  'Right Justify
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
         Index           =   17
         Left            =   10740
         TabIndex        =   57
         Top             =   285
         Width           =   855
      End
      Begin VB.Label ALBL 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "BChr Fr"
         BeginProperty Font 
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
         Left            =   10740
         TabIndex        =   56
         Top             =   0
         Width           =   855
      End
      Begin VB.Label ALBL 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Bag Type"
         BeginProperty Font 
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
         Left            =   7620
         TabIndex        =   55
         Top             =   285
         Width           =   1335
      End
      Begin VB.Label ALBL 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "       To"
         BeginProperty Font 
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
         Left            =   4290
         TabIndex        =   24
         Top             =   285
         Width           =   795
      End
      Begin VB.Label ALBL 
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
         Index           =   0
         Left            =   90
         TabIndex        =   23
         Top             =   0
         Width           =   735
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
         Left            =   1980
         TabIndex        =   22
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
         Left            =   2415
         TabIndex        =   21
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
         Index           =   40
         Left            =   3105
         TabIndex        =   20
         Top             =   0
         Width           =   105
      End
      Begin VB.Label ALBL 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Sort On"
         BeginProperty Font 
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
         Left            =   7620
         TabIndex        =   19
         Top             =   0
         Width           =   1335
      End
      Begin VB.Label ALBL 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Dsg Fr"
         BeginProperty Font 
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
         Left            =   4290
         TabIndex        =   18
         Top             =   0
         Width           =   795
      End
   End
   Begin VB.Frame FRA_NKEYS 
      BorderStyle     =   0  'None
      Enabled         =   0   'False
      Height          =   9810
      Left            =   -120
      TabIndex        =   15
      Top             =   0
      Width           =   15315
      Begin VB.Frame FraNKeyAll 
         BorderStyle     =   0  'None
         Height          =   9645
         Left            =   150
         TabIndex        =   17
         Top             =   -180
         Width           =   15150
         Begin VB.TextBox txtSEOPwd 
            Height          =   285
            IMEMode         =   3  'DISABLE
            Left            =   13950
            MaxLength       =   30
            PasswordChar    =   "*"
            TabIndex        =   80
            ToolTipText     =   "Enter SEO Password"
            Top             =   480
            Width           =   1095
         End
         Begin VB.Frame FraOrdDet 
            Height          =   1155
            Left            =   0
            TabIndex        =   27
            Top             =   8340
            Width           =   15105
            Begin MwfCtl.MWCTL_MED ATXT 
               Height          =   285
               Index           =   2
               Left            =   1485
               TabIndex        =   34
               ToolTipText     =   "Order Sr"
               Top             =   180
               Width           =   570
               _ExtentX        =   1005
               _ExtentY        =   503
               Enabled         =   0   'False
               Mask            =   "###0"
               MaxLength       =   4
               DataType        =   1
               ReCalcOn        =   "GrdBag(BODSR)"
               IdName          =   "WODBODSR"
            End
            Begin MwfCtl.MWCTL_MED ATXT 
               Height          =   285
               Index           =   4
               Left            =   6885
               TabIndex        =   37
               ToolTipText     =   "Current Component Bag Part Quantity"
               Top             =   750
               Width           =   660
               _ExtentX        =   1164
               _ExtentY        =   503
               Enabled         =   0   'False
               Mask            =   "-####0"
               MaxLength       =   6
               DataType        =   1
               IdName          =   "WODCTPTQTY"
            End
            Begin MwfCtl.MWCTL_MED ATXT 
               Height          =   285
               Index           =   6
               Left            =   8580
               TabIndex        =   53
               ToolTipText     =   "Balance Quantity For THe Ord Sr"
               Top             =   180
               Width           =   825
               _ExtentX        =   1455
               _ExtentY        =   503
               Alignment       =   2
               Enabled         =   0   'False
               Mask            =   "####0.0"
               MaxLength       =   7
               DataType        =   2
               IdName          =   "WODBALQTY"
            End
            Begin MwfCtl.MWCTL_MED ATXT 
               Height          =   285
               Index           =   3
               Left            =   4965
               TabIndex        =   54
               ToolTipText     =   "Current Bag Quantity For The Ord Sr"
               Top             =   180
               Width           =   825
               _ExtentX        =   1455
               _ExtentY        =   503
               Alignment       =   2
               Enabled         =   0   'False
               Mask            =   "####0.0"
               MaxLength       =   7
               DataType        =   2
               IdName          =   "WODBQTY"
            End
            Begin MwfCtl.MWCTL_MED ATXT 
               Height          =   285
               Index           =   18
               Left            =   4965
               TabIndex        =   59
               ToolTipText     =   "Current Flute Bag Quantity For The Ord Sr"
               Top             =   465
               Width           =   825
               _ExtentX        =   1455
               _ExtentY        =   503
               Alignment       =   2
               Enabled         =   0   'False
               Mask            =   "####0.0"
               MaxLength       =   7
               DataType        =   2
               IdName          =   "WODFBBQTY"
            End
            Begin MwfCtl.MWCTL_MED ATXT 
               Height          =   285
               Index           =   12
               Left            =   8580
               TabIndex        =   58
               ToolTipText     =   "Balance Flute Bag Quantity For The  Bag Character"
               Top             =   465
               Width           =   825
               _ExtentX        =   1455
               _ExtentY        =   503
               Alignment       =   2
               Enabled         =   0   'False
               Mask            =   "####0.0"
               MaxLength       =   7
               DataType        =   2
               IdName          =   "WODFBBALQTY"
            End
            Begin MwfCtl.MWCTL_MED ATXT 
               Height          =   285
               Index           =   13
               Left            =   1485
               TabIndex        =   62
               ToolTipText     =   "Flute Bag Character"
               Top             =   465
               Width           =   630
               _ExtentX        =   1111
               _ExtentY        =   503
               Enabled         =   0   'False
               MaxLength       =   3
               DataType        =   4
               ReCalcOn        =   "GrdBag(BCHR)"
               IdName          =   "WFBBCHR"
            End
            Begin MwfCtl.MWCTL_MED ATXT 
               Height          =   285
               Index           =   17
               Left            =   1485
               TabIndex        =   65
               ToolTipText     =   "Enter From Bag Character"
               Top             =   750
               Width           =   630
               _ExtentX        =   1111
               _ExtentY        =   503
               Enabled         =   0   'False
               MaxLength       =   3
               DataType        =   4
               IdName          =   "WCTBCHR"
            End
            Begin MwfCtl.MWCTL_MED ATXT 
               Height          =   285
               Index           =   19
               Left            =   4965
               TabIndex        =   67
               ToolTipText     =   "Current Flute Bag Quantity For The Ord Sr"
               Top             =   750
               Width           =   825
               _ExtentX        =   1455
               _ExtentY        =   503
               Alignment       =   2
               Enabled         =   0   'False
               Mask            =   "####0.0"
               MaxLength       =   7
               DataType        =   2
               IdName          =   "WODCTBQTY"
            End
            Begin MwfCtl.MWCTL_MED ATXT 
               Height          =   285
               Index           =   20
               Left            =   8580
               TabIndex        =   68
               ToolTipText     =   "Balance Flute Bag Quantity For The  Bag Character"
               Top             =   750
               Width           =   825
               _ExtentX        =   1455
               _ExtentY        =   503
               Alignment       =   2
               Enabled         =   0   'False
               Mask            =   "####0.0"
               MaxLength       =   7
               DataType        =   2
               IdName          =   "WODCTBALQTY"
            End
            Begin MwfCtl.MWCTL_MED ATXT 
               Height          =   285
               Index           =   21
               Left            =   10860
               TabIndex        =   71
               ToolTipText     =   "Balance Flute Bag Quantity For The  Bag Character"
               Top             =   750
               Width           =   825
               _ExtentX        =   1455
               _ExtentY        =   503
               Alignment       =   2
               Enabled         =   0   'False
               Mask            =   "####0.0"
               MaxLength       =   7
               DataType        =   2
               IdName          =   "WODCTBALPTQTY"
            End
            Begin MwfCtl.MWCTL_MED ATXT 
               Height          =   285
               Index           =   23
               Left            =   6885
               TabIndex        =   74
               ToolTipText     =   "Current Part Quantity"
               Top             =   180
               Width           =   660
               _ExtentX        =   1164
               _ExtentY        =   503
               Enabled         =   0   'False
               Mask            =   "-####0"
               MaxLength       =   6
               DataType        =   1
               IdName          =   "WODTPTQTY"
            End
            Begin MwfCtl.MWCTL_MED ATXT 
               Height          =   285
               Index           =   24
               Left            =   6885
               TabIndex        =   75
               ToolTipText     =   "Current Flute Bag Part Quantity"
               Top             =   465
               Width           =   660
               _ExtentX        =   1164
               _ExtentY        =   503
               Enabled         =   0   'False
               Mask            =   "-####0"
               MaxLength       =   6
               DataType        =   1
               IdName          =   "WODFBPTQTY"
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
               Index           =   26
               Left            =   6180
               TabIndex        =   77
               Top             =   465
               Width           =   705
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
               Index           =   25
               Left            =   6180
               TabIndex        =   76
               Top             =   180
               Width           =   705
            End
            Begin VB.Label ALBL 
               Alignment       =   1  'Right Justify
               BackStyle       =   0  'Transparent
               Caption         =   "Bal Prt Qty"
               BeginProperty Font 
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
               Left            =   9630
               TabIndex        =   72
               Top             =   750
               Width           =   1185
            End
            Begin VB.Label ALBL 
               BackStyle       =   0  'Transparent
               Caption         =   "Comp Bag Curr Qty"
               BeginProperty Font 
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
               Left            =   3030
               TabIndex        =   70
               Top             =   750
               Width           =   1905
            End
            Begin VB.Label ALBL 
               BackStyle       =   0  'Transparent
               Caption         =   "Bal Qty"
               BeginProperty Font 
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
               Left            =   7830
               TabIndex        =   69
               Top             =   750
               Width           =   855
            End
            Begin VB.Label ALBL 
               BackStyle       =   0  'Transparent
               Caption         =   "Comp Bag Chr"
               BeginProperty Font 
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
               Left            =   60
               TabIndex        =   66
               Top             =   750
               Width           =   1395
            End
            Begin VB.Label ALBL 
               BackStyle       =   0  'Transparent
               Caption         =   "Flute Bag Chr"
               BeginProperty Font 
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
               Left            =   90
               TabIndex        =   63
               Top             =   465
               Width           =   1305
            End
            Begin VB.Label ALBL 
               BackStyle       =   0  'Transparent
               Caption         =   "Bal Qty"
               BeginProperty Font 
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
               Left            =   7830
               TabIndex        =   61
               Top             =   465
               Width           =   855
            End
            Begin VB.Label ALBL 
               BackStyle       =   0  'Transparent
               Caption         =   "Flute Bag Curr Qty"
               BeginProperty Font 
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
               Left            =   3030
               TabIndex        =   60
               Top             =   465
               Width           =   1905
            End
            Begin VB.Label ALBL 
               BackStyle       =   0  'Transparent
               Caption         =   "Bal Qty"
               BeginProperty Font 
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
               Left            =   7830
               TabIndex        =   39
               Top             =   180
               Width           =   855
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
               Index           =   9
               Left            =   6180
               TabIndex        =   38
               Top             =   750
               Width           =   705
            End
            Begin VB.Label ALBL 
               BackStyle       =   0  'Transparent
               Caption         =   "Prim. Bag Curr Qty"
               BeginProperty Font 
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
               Left            =   3030
               TabIndex        =   36
               Top             =   180
               Width           =   1905
            End
            Begin VB.Label ALBL 
               BackStyle       =   0  'Transparent
               Caption         =   "Ord Sr."
               BeginProperty Font 
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
               Left            =   90
               TabIndex        =   35
               Top             =   180
               Width           =   795
            End
         End
         Begin VB.Frame FraBagDet 
            Height          =   495
            Left            =   0
            TabIndex        =   26
            Top             =   7890
            Width           =   8025
            Begin MwfCtl.MWCTL_MED ATXT 
               Height          =   285
               Index           =   33
               Left            =   975
               TabIndex        =   28
               ToolTipText     =   "Bag Number"
               Top             =   150
               Width           =   1680
               _ExtentX        =   2963
               _ExtentY        =   503
               Enabled         =   0   'False
               MaxLength       =   14
               DataType        =   4
               ReCalcOn        =   "GrdBag(BYY),GrdBag(BCHR),GrdBag(BNO)"
               IdName          =   "WBDBAG"
            End
            Begin MwfCtl.MWCTL_MED ATXT 
               Height          =   285
               Index           =   34
               Left            =   4965
               TabIndex        =   30
               ToolTipText     =   "Current Bag Quantity"
               Top             =   150
               Width           =   600
               _ExtentX        =   1058
               _ExtentY        =   503
               Alignment       =   2
               Enabled         =   0   'False
               Mask            =   "##0.0"
               MaxLength       =   5
               DataType        =   2
               ReCalcOn        =   "GrdBag(BYY),GrdBag(BCHR),GrdBag(BNO)"
               IdName          =   "WBDBQTY"
            End
            Begin MwfCtl.MWCTL_MED ATXT 
               Height          =   285
               Index           =   35
               Left            =   6885
               TabIndex        =   32
               ToolTipText     =   "Bag Part Quantity"
               Top             =   150
               Width           =   660
               _ExtentX        =   1164
               _ExtentY        =   503
               Enabled         =   0   'False
               Mask            =   "-####0"
               MaxLength       =   6
               DataType        =   1
               ReCalcOn        =   "GrdBag(BYY),GrdBag(BCHR),GrdBag(BNO)"
               IdName          =   "WBDPTQTY"
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
               Index           =   3
               Left            =   6180
               TabIndex        =   33
               Top             =   150
               Width           =   705
            End
            Begin VB.Label ALBL 
               BackStyle       =   0  'Transparent
               Caption         =   "Curr Qty"
               BeginProperty Font 
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
               Left            =   3120
               TabIndex        =   31
               Top             =   150
               Width           =   885
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
               Index           =   2
               Left            =   120
               TabIndex        =   29
               Top             =   150
               Width           =   945
            End
         End
         Begin MwfCtl.MWCTL_FLX GrdBag 
            Height          =   6780
            Left            =   0
            TabIndex        =   11
            Top             =   810
            Width           =   15120
            _ExtentX        =   26670
            _ExtentY        =   11959
            Cols            =   19
            colname1        =   "BYY"
            heading1        =   "BYy"
            datafld1        =   "BYy"
            datatype1       =   4
            recalcparent1   =   "GrdBag(BNO),WBDBAG,WBDBQTY,WBDPTQTY,GrdBag(BOPNPTQTY)"
            maxlength1      =   2
            tooltiptext1    =   "Enter Bag Year"
            colname2        =   "BCHR"
            heading2        =   "BChr  "
            datafld2        =   "BChr"
            datatype2       =   4
            recalcparent2   =   "GrdBag(BNO),WBDBAG,WBDBQTY,WBDPTQTY,WFBBCHR,GrdBag(BOPNQTY),GrdBag(BOPNPTQTY),GrdBag(BOPNLOC)"
            maxlength2      =   3
            tooltiptext2    =   "Enter Bag Character"
            colname3        =   "BNO"
            heading3        =   "BNo         "
            datafld3        =   "BNo"
            datatype3       =   1
            mask3           =   "######0"
            recalcon3       =   "GrdBag(BYY),GrdBag(BCHR)"
            recalcparent3   =   "WBDBAG,WBDBQTY,WBDPTQTY,GrdBag(BOPNPTQTY)"
            maxlength3      =   7
            tooltiptext3    =   "Enter Bag Number"
            colname4        =   "BODSR"
            heading4        =   "OrdSr"
            datafld4        =   "BOdSr"
            datatype4       =   1
            mask4           =   "###0"
            recalcparent4   =   "GrdBag(BODDMCD),GrdBag(BODSFX),GrdBag(BODDMSZ),WODBODSR,GrdBag(BODCTCD),GrdBag(BODIDNO),GrdBag(BDMIDNO),GrdBag(BOMIDNO)"
            maxlength4      =   4
            tooltiptext4    =   "Enter Order Sr."
            colname5        =   "BODDMCD"
            heading5        =   "Design Cd                   "
            datafld5        =   "BOdDmCd"
            datatype5       =   4
            recalcon5       =   "GrdBag(BODSR)"
            maxlength5      =   15
            style5          =   2
            tooltiptext5    =   "Design Code"
            colname6        =   "BODSFX"
            heading6        =   "Suffix                 "
            datafld6        =   "BOdSfx"
            datatype6       =   4
            recalcon6       =   "GrdBag(BODSR)"
            maxlength6      =   15
            style6          =   2
            tooltiptext6    =   "Design Suffix"
            colname7        =   "BODDMSZ"
            heading7        =   "Size     "
            datafld7        =   "BOdDmSz"
            datatype7       =   4
            recalcon7       =   "GrdBag(BODSR)"
            maxlength7      =   5
            style7          =   2
            tooltiptext7    =   "Design Size"
            colname8        =   "BOPNQTY"
            heading8        =   "OpnQty"
            datafld8        =   "BOpnQty"
            datatype8       =   2
            mask8           =   "##0.0"
            recalcon8       =   "GrdBag(BCHR)"
            recalcparent8   =   "GrdBag(BOPNPTQTY)"
            maxlength8      =   5
            tooltiptext8    =   "Enter Opening Quantity"
            colname9        =   "BOPNPTQTY"
            heading9        =   "OpnPt"
            datafld9        =   "BOpnPtQty"
            datatype9       =   1
            mask9           =   "####0"
            recalcon9       =   "GrdBag(BCHR),GrdBag(BNO),GrdBag(BOPNQTY),GrdBag(BYY)"
            maxlength9      =   5
            tooltiptext9    =   "Enter Opening Part Quantity"
            colname10       =   "BOPNDT"
            heading10       =   "OpnDate"
            datafld10       =   "BOpnDt"
            datatype10      =   3
            maxlength10     =   8
            tooltiptext10   =   "Enter Opening Date"
            colname11       =   "BOPNLOC"
            heading11       =   "OpnLoc        "
            datafld11       =   "BOpnLoc"
            datatype11      =   4
            recalcon11      =   "GrdBag(BCHR)"
            recalcparent11  =   "GrdBag(BLOC)"
            maxlength11     =   8
            tooltiptext11   =   "Enter Opening Location"
            colname12       =   "BLOC"
            heading12       =   "BLoc            "
            datafld12       =   "BLoc"
            datatype12      =   4
            recalcon12      =   "GrdBag(BOPNLOC)"
            maxlength12     =   8
            style12         =   2
            tooltiptext12   =   "Enter Bag Location"
            colname13       =   "BODLSRNO"
            heading13       =   "OrdLn"
            datafld13       =   "BOdlSrNo"
            datatype13      =   1
            mask13          =   "######0"
            recalcparent13  =   "GrdBag(BODLNIDKEY)"
            maxlength13     =   7
            tooltiptext13   =   "Enter BOdlSrNo"
            colname14       =   "BODLNIDKEY"
            heading14       =   "OdlID"
            datafld14       =   "BOdLnIdKey"
            datatype14      =   1
            mask14          =   "######0"
            recalcon14      =   "GrdBag(BODLSRNO)"
            maxlength14     =   7
            style14         =   0
            colname15       =   "BCLS"
            datafld15       =   "BCls"
            datatype15      =   4
            maxlength15     =   1
            style15         =   0
            tooltiptext15   =   "Enter Bag Closed (Y/N)"
            colname16       =   "BOPNDESC"
            datafld16       =   "BOpnDesc"
            datatype16      =   4
            recalcon16      =   "WBOPNDESC"
            maxlength16     =   45
            style16         =   0
            tooltiptext16   =   "Enter Description"
            colname17       =   "BODCTCD"
            heading17       =   "Contractor "
            datafld17       =   "BOdCtCd"
            datatype17      =   4
            recalcon17      =   "GrdBag(BODSR)"
            maxlength17     =   8
            tooltiptext17   =   "Enter Contractor Code"
            colname18       =   "BPRTKEY"
            heading18       =   "Prtn "
            datafld18       =   "BPrtKey"
            maxlength18     =   1
            style18         =   0
         End
         Begin MwfCtl.MWCTL_MED ATXT 
            Height          =   285
            Index           =   25
            Left            =   615
            TabIndex        =   12
            ToolTipText     =   "Enter Description"
            Top             =   7590
            Width           =   6930
            _ExtentX        =   12224
            _ExtentY        =   503
            MaxLength       =   45
            DataType        =   4
            ReCalcParent    =   "GrdBag(BOPNDESC)"
            IdName          =   "WBOPNDESC"
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
            Left            =   12900
            TabIndex        =   81
            Top             =   480
            Width           =   945
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
            TabIndex        =   25
            Top             =   7590
            Width           =   660
         End
      End
      Begin VB.Frame FraBagGen 
         Height          =   1600
         Left            =   5070
         TabIndex        =   41
         Top             =   7600
         Width           =   5040
         Begin MwfCtl.MWCTL_BTN1 CmdBagGenGo 
            Height          =   495
            Left            =   4260
            TabIndex        =   51
            ToolTipText     =   "Click Here To Generate Bags (Fr Sr=0 And To Sr=0 Generates Bags For Entire Order)"
            Top             =   495
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
            Index           =   9
            Left            =   1260
            TabIndex        =   44
            ToolTipText     =   "Enter From Order Sr."
            Top             =   420
            Width           =   555
            _ExtentX        =   979
            _ExtentY        =   503
            Alignment       =   2
            Mask            =   "###0"
            MaxLength       =   4
            DataType        =   1
            ReCalcParent    =   "WBGDEFBPCS"
            IdName          =   "WBGODSRFR"
         End
         Begin MwfCtl.MWCTL_MED ATXT 
            Height          =   285
            Index           =   8
            Left            =   1260
            TabIndex        =   45
            ToolTipText     =   "Enter To Order Sr."
            Top             =   705
            Width           =   555
            _ExtentX        =   979
            _ExtentY        =   503
            Alignment       =   2
            Mask            =   "###0"
            MaxLength       =   4
            DataType        =   1
            ReCalcParent    =   "WBGDEFBPCS"
            IdName          =   "WBGODSRTO"
         End
         Begin MwfCtl.MWCTL_MED ATXT 
            Height          =   285
            Index           =   10
            Left            =   3240
            TabIndex        =   47
            ToolTipText     =   "Enter Default Bag Character"
            Top             =   420
            Width           =   555
            _ExtentX        =   979
            _ExtentY        =   503
            MaxLength       =   3
            DataType        =   4
            IdName          =   "WBGDEFBCHR"
         End
         Begin MwfCtl.MWCTL_MED ATXT 
            Height          =   285
            Index           =   11
            Left            =   3240
            TabIndex        =   48
            ToolTipText     =   "Enter No. Of Default Bag Pieces."
            Top             =   705
            Width           =   555
            _ExtentX        =   979
            _ExtentY        =   503
            Alignment       =   2
            Mask            =   "##0.0"
            MaxLength       =   5
            DataType        =   2
            ReCalcOn        =   "WBGODSRFR,WBGODSRTO"
            IdName          =   "WBGDEFBPCS"
         End
         Begin MwfCtl.MWCTL_MED ATXT 
            Height          =   285
            Index           =   26
            Left            =   3240
            TabIndex        =   49
            ToolTipText     =   "Enter No. Of Default Bag Pieces."
            Top             =   990
            Width           =   1250
            _ExtentX        =   2196
            _ExtentY        =   503
            Alignment       =   2
            MaxLength       =   8
            DataType        =   4
            ReCalcOn        =   "WBGODSRFR,WBGODSRTO"
            IdName          =   "WBGDEFBOPNLOC"
         End
         Begin MwfCtl.MWCTL_MED ATXT 
            Height          =   285
            Index           =   27
            Left            =   1260
            TabIndex        =   46
            ToolTipText     =   "Enter Order Line"
            Top             =   990
            Width           =   555
            _ExtentX        =   979
            _ExtentY        =   503
            Alignment       =   2
            Mask            =   "#0"
            MaxLength       =   2
            DataType        =   1
            ReCalcParent    =   "WBGDEFBPCS"
            IdName          =   "WBGODLN"
         End
         Begin VB.Label ALBL 
            BackStyle       =   0  'Transparent
            Caption         =   "OrdLn"
            BeginProperty Font 
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
            Left            =   90
            TabIndex        =   79
            Top             =   990
            Width           =   1035
         End
         Begin VB.Label ALBL 
            BackStyle       =   0  'Transparent
            Caption         =   "Def Opn Loc"
            BeginProperty Font 
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
            Left            =   1980
            TabIndex        =   78
            Top             =   990
            Width           =   1305
         End
         Begin VB.Label LblMain 
            BackStyle       =   0  'Transparent
            BorderStyle     =   1  'Fixed Single
            Caption         =   "  Bag Generation"
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
            TabIndex        =   64
            Top             =   90
            Width           =   5130
         End
         Begin VB.Label ALBL 
            BackStyle       =   0  'Transparent
            Caption         =   "Def Bag Pcs"
            BeginProperty Font 
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
            Left            =   1980
            TabIndex        =   52
            Top             =   705
            Width           =   1305
         End
         Begin VB.Label ALBL 
            BackStyle       =   0  'Transparent
            Caption         =   "Def Bag Chr"
            BeginProperty Font 
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
            Left            =   1980
            TabIndex        =   50
            Top             =   420
            Width           =   1305
         End
         Begin VB.Label ALBL 
            BackStyle       =   0  'Transparent
            Caption         =   "To Sr No"
            BeginProperty Font 
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
            TabIndex        =   43
            Top             =   705
            Width           =   1065
         End
         Begin VB.Label ALBL 
            BackStyle       =   0  'Transparent
            Caption         =   "From Sr No"
            BeginProperty Font 
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
            Left            =   90
            TabIndex        =   42
            Top             =   420
            Width           =   1215
         End
      End
   End
End
Attribute VB_Name = "EmrFrmBag"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


'---- General Comments ----
' Entry for generating Bags
' 2 frames Bd-BagDetail Od-Order Details only for display purpose.
' A- Entry -
'----
Option Explicit
Dim moCn As MwfLib.MDOConnection   '*** (Bef speed) bef 24/11/06
'****** Sachin 2.14.0 - 03-11-2006 [Multi-Company] ******
Dim ms_wBCoCd As String, ms_MultiComp As Boolean
'****** Sachin 2.14.0 - 03-11-2006 [Multi-Company] ******
Dim ms_wBOdTc As String, ms_wBOdYy As String, ms_wBOdChr As String, ms_wBOdNo As Long
Dim ms_LockDt As String, ms_PurgeDt As String
'*** For the Toggle Option of Command Buttons
Enum en_BagFra
  BagGen = 0
End Enum
Dim ms_PwdFlag As Boolean
Dim ms_OrdLck As Boolean    '****** Sachin 3.02 20-12-07 - For Checking if the Order belongs to previous partition
Dim ms_wBOdDmTcTyp As String
Private Sub ADC_Load()
  '*** Set the First Key Control property to the working field wBOdTc
  '*** Set the First Non Key Control property to the grid GrdBag
  '*** Set the Child Property Of ADC to GrdBag
  '*** Set the Previous and Next Control Properties for the Grid GrdBag
  '*** Set the Hot Keys for the Grid GrdBag as 'A'
  '*** initialising the values for wBOdTc, wBOdYy, wBOdChr and wBOdNo
  Set moCn = ADC.Connection    '*** (Bef speed) bef 24/11/06
  
  '****** Sachin 2.14.0 - 03-11-2006 [Multi-Company] ******
  ms_MultiComp = IIF(moCn.GetFldVal("Select hImpComLoc From Head Where hCd= '" + ctSelfCmCd + "' And hCoCd='" + ctSelfCoCd + "'") = "Y", True, False)
  Set ADC.FirKeyCtl = IIF(ms_MultiComp = True, ADC("wBCoCd"), ADC("WBODTC"))
  '****** Sachin 2.14.0 - 03-11-2006 [Multi-Company] ******
  
  Set ADC.FirNKeyCtl = ADC
  ADC.Child = "GrdBag"
  Set GrdBag.PrevCtl = ADC
  Set GrdBag.NextCtl = ADC
  GrdBag.HotKey = "A"
    
  '****** Sachin 2.14.0 - 03-11-2006 [Multi-Company] ******
  ms_wBCoCd = ""
  '****** Sachin 2.14.0 - 03-11-2006 [Multi-Company] ******
  ms_wBOdTc = ""
  ms_wBOdYy = ""
  ms_wBOdChr = ""
  ms_wBOdNo = 0
  
    
  Me.Caption = GetFrmCaption(ADC("wBCoCd"), "Bag Master")
  

    
End Sub

Private Sub ADC_KeyWhen()
  '*** setting the default values for wBOdTc, wBOdYy, wBOdChr and wBOdNo from the
      'previous entry
  '*** Disable All command buttons
  '*** Bring the frame FraNKeyAll to the top and hide all other frames
  If ms_wBCoCd <> "" Then
    ADC("wBCoCd") = ms_wBCoCd
  Else
    ADC("wBCoCd") = gs_CoCd
  End If
  If ms_MultiComp = False Then
    ADC("wBCoCd").Enabled = False
  Else
    ADC("wBCoCd").Enabled = True
  End If

  'FraNKeyAll.ZOrder (0)
  
  '****** Sachin 2.14.0 - 03-11-2006 [Multi-Company] ******
  '****** Sachin 2.14.0 - 03-11-2006 [Multi-Company] ******
  
  ' ***** Manali 3.6.0 - 01/09/09 - Default Values Set
  If ms_wBOdTc = "" Then
    ms_wBOdTc = moCn.GetFldVal("Select PValue from Param where PTyp='TCTYP' and PMCd='SO'")
  End If
  If ms_wBOdYy = "" Then
    ms_wBOdYy = moCn.GetFldVal("Select HYy from Head where HCoCd='" + ADC("wBCoCd") + "' and HCd='" + ctSelfCmCd + "'")
  End If
  ' ***** Manali 3.6.0 - 01/09/09 - Default Values Set
  
  If ms_wBOdTc <> "" Then ADC("wBOdTc") = ms_wBOdTc
  If ms_wBOdYy <> "" Then ADC("wBOdYy") = ms_wBOdYy
  If ms_wBOdChr <> "" Then ADC("wBOdChr") = ms_wBOdChr
  If ms_wBOdNo <> 0 Then ADC("wBOdNo") = ms_wBOdNo
  
  Call EnaDisaCmds(True)
  Call HideAllFras
  txtseopwd.text = ""
  
End Sub

Private Sub ADC_KeyValid(pr_Cancel As Boolean, pr_ErrMsg As String)
  '*** Force the user to enter valid values in wBOdTc, wBOdYy, wBOdChr, wBOdNo
  '****** Sachin 2.14.0 - 03-11-2006 [Multi-Company] ******
  If ADC("wBCoCd") = "" Then pr_Cancel = True: pr_ErrMsg = "Enter Order Company Code": Exit Sub
  '****** Sachin 2.14.0 - 03-11-2006 [Multi-Company] ******
  
  If ADC("wBOdTc") = "" Then pr_Cancel = True: pr_ErrMsg = "Enter Order Tc": Exit Sub
  If ADC("wBOdYy") = "" Then pr_Cancel = True: pr_ErrMsg = "Enter Order Year": Exit Sub
  If ADC("wBOdChr") = "" Then pr_Cancel = True: pr_ErrMsg = "Enter Order Character": Exit Sub
  If ADC("wBOdNo") = 0 Then pr_Cancel = True: pr_ErrMsg = "Enter Order Number": Exit Sub
  
  Call ATXT_FldChange(0, UCase("wBOdNo"), pr_Cancel, pr_ErrMsg, ADC("wBOdNo"))
  If pr_Cancel = True Then Exit Sub
   
  '****** Sachin 2.14.0 - 04-11-2006 [Multi-Company] - Shifted here From ADC_Load ******
  If moCn.GetFldVal("Select HCtbMod from Head where HCoCd= '" + ADC("wBCoCd") + "' and HCd= '" + ctSelfCmCd + "'") = "N" Then
    GrdBag.ColProp("BOdCtCd").Style = fgcinvisible
  Else
    GrdBag.ColProp("BOdCtCd").Style = fgcNormal
  End If
  '****** Sachin 2.14.0 - 04-11-2006 [Multi-Company] - Shifted here From ADC_Load ******
  
End Sub

Private Sub ADC_NonKeyValid(pr_Cancel As Boolean, pr_ErrMsg As String)
    Dim wRsBagChrExceedingPrdQty   As MDORowSet
    Dim wErrStr As String
    wErrStr = ""
    Set wRsBagChrExceedingPrdQty = moCn.OpenRes("Select BChr, BOdSr, sum(BOpnQty) as BOpnQty, max(OdPrdQty) as OdPrdQty From Bag join Param on PTyp = 'BCHR' and PMCd = BChr and PScd = '' join OrdDsg on OdIdNo = BOdIdNo where " + _
                      " BCoCD = '" + ADC("wBCoCd") + "' And BOdTc = '" + ADC("wBOdTc") + "' And " + _
                      " bOdYy = '" + ADC("wBOdYy") + "' And BOdChr = '" + ADC("wBOdChr") + "' And BOdNo = " + CStr(ADC("wBOdNo")) + " And bPrtKey = '" + IIF(gs_Partition = ctCurrPrtn, ctCurrPrtn, "OmPrtKey") + "' and Pvalue = 'F' " + _
                      " group by BChr, BOdSr having sum(BOpnQty) > max(OdPrdQty)")

    If wRsBagChrExceedingPrdQty.RecCount > 0 Then
     With wRsBagChrExceedingPrdQty
      wErrStr = ""
      .MoveFirst
      Do While Not (.BOF Or .EOF)
        wErrStr = IIF(wErrStr = "", "", wErrStr + ";") + CStr(wRsBagChrExceedingPrdQty.FldValue("BOdSr")) + "/" + CStr(wRsBagChrExceedingPrdQty.FldValue("BChr") + "/" + CStr(wRsBagChrExceedingPrdQty.FldValue("BOpnQty")) + "/" + CStr(wRsBagChrExceedingPrdQty.FldValue("OdPrdQty")))
        .MoveNext
      Loop
      End With
    End If
    If wErrStr <> "" Then
    ' We made a change for IG where we wanted this to be based on SEO PWD and then reverted
    ' it so I am removing the Pwd flag from here and changing it to just if True
      If True Then
        If (MsgBox("Warning: Some of the bags opened exceed the Production Quantity - OrdSr/BagChr/BagQty/PrdQty:" + wErrStr + ". Press Ok to continue, or Cancel to cancel", vbOKCancel, "Bag Qty exceeds Prdn Qty") = vbCancel) Then
          pr_Cancel = True
          pr_ErrMsg = "Warning: Some of the bags opened exceed the Production Quantity - OrdSr/BagChr:" + wErrStr + ""
          Exit Sub
        End If
      Else
        MsgBox ("Cannot proceed, some of the bags opened exceed the Production Quantity - OrdSr/BagChr/BagQty/PrdQty:" + wErrStr)
        pr_Cancel = True
        pr_ErrMsg = "Cannot proceed, some of the bags opened exceed the Production Quantity - OrdSr/BagChr:" + wErrStr + ""
      End If
    End If
    
End Sub

Private Sub ADC_NonKeyWhen(pr_Cancel As Boolean, pr_ErrMsg As String)
  '*** Store the wBOdTc, wBOdYy, wBOdChr, wBOdNo values for the next entry
  ms_wBOdDmTcTyp = IIF(ADC("wBOdTc") = "OS", "SM", "DM")
  ms_LockDt = moCn.GetFldVal("Select PDesc225 from Param where PTyp= 'LOCKDT' and PMCd= 'LOCKDT'")
  'sv.9 getting purge date from vparam
  ms_PurgeDt = GetPurgeDt(ADC("wBCoCd"), "")
  '*** For the time being do not put Lock Date Check in the Order Entry
  If ms_LockDt = "" Then ms_LockDt = "01/01/80"
    
  Call EnaDisaCmds(False)
  '****** Sachin 2.14.0 - 03-11-2006 [Multi-Company] ******
  ms_wBCoCd = ADC("wBCoCd")
  '****** Sachin 2.14.0 - 03-11-2006 [Multi-Company] ******
  ms_wBOdTc = ADC("wBOdTc")
  ms_wBOdYy = ADC("wBOdYy")
  ms_wBOdChr = ADC("wBOdChr")
  ms_wBOdNo = ADC("wBOdNo")
  
  '****** Sachin 3.02 20-12-07 - Checking whether the Order Is Locked
  ms_OrdLck = moCn.RecSeek("Select 'x' From OrdMst Where " + _
                               "OmCoCd='" + ADC("wBCoCd") + "' And OmTc='" + ADC("wBOdTc") + "' And OmYy='" + ADC("wBOdYy") + "' " + _
                               " And OmChr='" + ADC("wBOdChr") + "' And OmNo=" + CStr(ADC("wBOdNo")) + " And OmPrtKey<>'" + ctCurrPrtn + "' ")
  If ms_OrdLck = True Then
     GrdBag.AllowAdd = False: GrdBag.AllowDelete = False
  Else
     GrdBag.AllowAdd = True: GrdBag.AllowDelete = True
  End If
  '****** Sachin 3.02 20-12-07 - Checking whether the Order Is Locked
  Call SetSEOPwdFlag
End Sub

Private Sub ATXT_FldWhen(Index As Integer, ByVal IdName As String, Cancel As Boolean, ErrMsg As String)
  
  '*** Set Helps For Each Field
  Select Case IdName
  '****** Sachin 2.14.0 - 03-11-2006 [Multi-Company] ******
  Case Is = UCase("wBCoCd")
    Call HlpList.hCoCd
  '****** Sachin 2.14.0 - 03-11-2006 [Multi-Company] ******
  'indigo.164 SS Tctype added
  Case Is = UCase("wBOdTc")
    Call HlpList.PMCd("TC", "'SO','OS'")
  Case Is = UCase("wBOdYy")
    Call HlpList.PMCd("YY")
  Case Is = UCase("wBOdChr")
  ' **** Zubin 211 **** '
    'Call HlpList.vPSCd(ADC("wBCoCd"), "CHR", ADC("wBOdTc"))
    Call HlpList.vPSCd(ADC("wBCoCd"), "CHR", ADC("wBOdTc"), "Y")
  ' **** Zubin 211 **** '
  Case Is = UCase("wBOdNo")
    Call HlpList.OmNo(ADC("wBCoCd"), ADC("wBOdTc"), ADC("wBOdYy"), ADC("wBOdChr"), , IIF(gs_Partition = ctCurrPrtn, ctCurrPrtn, ""))
  Case Is = UCase("wBOdFrDmCd")
    Call HlpList.OdDmCd(ADC("wBCoCd"), ADC("wBOdTc"), ADC("wBOdYy"), ADC("wBOdChr"), ADC("wBOdNo"))
  Case Is = UCase("wBOdToDmCd")
    Call HlpList.OdDmCd(ADC("wBCoCd"), ADC("wBOdTc"), ADC("wBOdYy"), ADC("wBOdChr"), ADC("wBOdNo"))
  Case Is = UCase("wBSrt")
    Call HlpList.PMCd("SRTBAG")
  Case Is = UCase("wBGOdSrFr"), UCase("wBGOdSrTo")
    Call HlpList.OdSr(ADC("wBCoCd"), ADC("wBOdTc"), ADC("wBOdYy"), ADC("wBOdChr"), ADC("wBOdNo"), True)
  
  '************************************************************************************
  Case Is = UCase("wBgDefBChr")
    Call HlpList.PMCd("BCHR")
  '************************************************************************************
  ' **** Zubin 211 **** '
  Case Is = UCase("wBagTyp")
    Call HlpList.PMCd("BTYP") '*** Jay 2.13(CT) ***
  Case Is = UCase("wBChrFr")
    Call HlpList.PMCd("BCHR")
  Case Is = UCase("wBChrTo")
    Call HlpList.PMCd("BCHR")
  ' **** Zubin 211 **** '
  
  'Sachin 3.11.0
  Case Is = UCase("WBGDEFBOPNLOC")
    Call HlpList.Loc(ADC("wBCoCd"), "'P'")
  'Uma EmrSv2
  Case Is = UCase("WBGODLN")
  If Not moCn.RecSeek("Select 'x' From SysDb where SysId='UNI'") Then
    Cancel = True
    ErrMsg = "To be entered only when Ppc System is installed"
  Else
    Call HlpList.OdlSrNo(ADC("wBOdTc"), ADC("wBOdYy"), ADC("WBODCHR"), ADC("wBOdNo"), 0, ADC("WBGODSRFR"), ADC("WBGODSRTO"))
  End If
  End Select
End Sub
Private Sub atxt_ReCalculate(Index As Integer, ByVal IdName As String, ByVal OldVal As Variant, ByVal RowNum As Integer, ByVal SourceName As String, ByVal SourceColName As String)
  '*** wBOdChr on wBOdTc         adc-Normal
      'PDesc225('TC', wBOdTc)
  '*** wBdBag on BYy,BChr,BNo    adc-Normal, GrdBag-Normal
      'wBdBag= BYy+BChr+BNo
  '*** wBdBQty on BYy,BChr,BNo   adc-Normal, GrdBag-Normal
      'wBdBQty= BQty from Bag(BYy,BChr,BNo)
  '*** wBdPtQty on BYy,BChr,BNo  adc-Normal, GrdBag-Normal
      'wBdPtQty= BPtQty from Bag(BYy,BChr,BNo)
  '*** wOdBOdSr on BOdSr         adc-Normal, GrdBag-Normal
      'wOdBOdSr= BOdSr
      'wOdBQty= Total Bag Quantity for the specified OrderSr
      'WODCTPTQTY= Total Bag Part Quantity for the specified OrderSr
      'wOdBalQty= Balance quantity for the specified OrderSr
      'If new row then BOpnQty = IIf(OdBagPcs from OrdDsg <= wOdBalQty, OdBagPcs, wOdBalQty)
  
  Dim wBagRow As Single, wBagOpnQty As Single
  ' **** Zubin 211 **** '
  Dim ws_BagType As String
  ' **** Zubin 211 **** '
  
  wBagRow = GrdBag.Row
  
  Select Case IdName
  Case Is = UCase("wBOdChr")
    If ADC.Mode = xNorm Then _
       ADC("wBOdChr") = moCn.GetFldVal("Select vPValue from vParam where " + _
                             "vPCoCd= '" + ADC("wBCoCd") + "' and vPTyp= 'DEFCHR' " + _
                             "and vPMCd= '" + ADC("wBOdTc") + "' and vPSCd= ''")
       '("Select PDesc225 from Param where PTyp='TC' " + _
                                      "and PMCd='" + ADC("wBOdTc") + "'")
                                      
                                      
  Case Is = UCase("wBdBag")
    If ADC.Mode = xNorm And GrdBag.Mode = fgmnorm Then
      If wBagRow > 0 And wBagRow <= GrdBag.Rows - 1 Then
        
        ADC("wBdBag") = GrdBag.Value(wBagRow, "BYy") + "/" + GrdBag.Value(wBagRow, "BChr") + _
                      "/" + Trim(CStr(GrdBag.Value(wBagRow, "BNo")))
      End If
    End If
    
  Case Is = UCase("wBdPtQty")
    '*** Getting Current Part Qty From Bag ***
    '****** Recalc below based on Partition selected during login
    If wBagRow > 0 And wBagRow <= GrdBag.Rows - 1 Then
      If ADC.Mode = xNorm And GrdBag.Mode = fgmnorm Then
      If wBagRow > 0 And wBagRow <= GrdBag.Rows - 1 Then
      ADC("wBdPtQty") = moCn.GetFldVal("Select BPtQty from Bag where " + _
                                      "BCoCd='" + ADC("wBCoCd") + "' and " + _
                                      "BYy='" + GrdBag.Value(wBagRow, "BYy") + "' and " + _
                                      "BChr='" + GrdBag.Value(wBagRow, "BChr") + "' and " + _
                                      "BNo= " + CStr(GrdBag.Value(wBagRow, "BNo")) + IIF(GrdBag.Value(wBagRow, "BPrtKey") = ctCurrPrtn, " And BPrtKey='" + ctCurrPrtn + "'", ""))
      End If
    End If
  End If
  Case Is = UCase("wOdBOdSr")
  '****Bef 2.11
'      '*** Getting Current Bag Qty, Part Qty From Bag and Bal Qty For The Order Sr.***
'    If ADC.Mode = xNorm And GrdBag.Mode = fgmnorm Then
'      Dim wRsOrdDet As MwfLib.MDORowSet
'      ADC("wOdBOdSr") = GrdBag.Value(wBagRow, "BOdSr")
'      Set wRsOrdDet = moCn.OpenRes("Select IsNull(sum(BQty), 0) as qTotQty, IsNull(sum(BPtQty),0) as qTotPtQty " + _
'                                    "from Bag where " + _
'                                    "BCoCd='" + ADC("wBCoCd") + "' and " + _
'                                    "BOdTc='" + ADC("wBOdTc") + "' and " + _
'                                    "BOdYy='" + ADC("wBOdYy") + "' and " + _
'                                    "BOdChr='" + ADC("wBOdChr") + "' and " + _
'                                    "BOdNo= " + CStr(ADC("wBOdNo")) + " and " + _
'                                    "BOdSr= " + CStr(GrdBag.Value(wBagRow, "BOdSr")))
'      If Not (wRsOrdDet.EOF Or wRsOrdDet.BOF) Then
'        ADC("wOdBQty") = wRsOrdDet!qTotQty
'        ADC("WODCTPTQTY") = wRsOrdDet!qTotPtQty
'      End If
'      ADC("wOdBalQty") = moCn.GetFldVal("Select OdPrdQty from OrdDsg where " + _
'                                       "OdCoCd='" + ADC("wBCoCd") + "' and " + _
'                                       "OdTc='" + ADC("wBOdTc") + "' and " + _
'                                       "OdYy='" + ADC("wBOdYy") + "' and " + _
'                                       "OdChr='" + ADC("wBOdChr") + "' and " + _
'                                       "OdNo= " + CStr(ADC("wBOdNo")) + " and " + _
'                                       "OdSr= " + CStr(GrdBag.Value(wBagRow, "BOdSr"))) - ADC("wOdBQty")
'
'      If GrdBag.IsNew(wBagRow) Then
'        wBagOpnQty = moCn.GetFldVal("Select OdBagPcs from OrdDsg where " + _
'              "OdCoCd='" + ADC("wBCoCd") + "' and " + _
'              "OdTc= '" + ADC("wBOdTc") + "' and OdYy= '" + ADC("wBOdYy") + "' and " + _
'              "OdChr= '" + ADC("wBOdChr") + "' and OdNo= " + CStr(ADC("wBOdNo")) + " and " + _
'              "OdSr= " + CStr(GrdBag.Value(wBagRow, "BOdSr")))
'
'        wBagOpnQty = IIF(wBagOpnQty <= ADC("wOdBalQty"), wBagOpnQty, ADC("wOdBalQty"))
'        GrdBag.Value(wBagRow, "BOpnQty") = wBagOpnQty
'      End If
'
'      Set wRsOrdDet = Nothing
'    End If
'
  
    '*** Getting Current Bag Qty, Part Qty From Bag and Bal Qty For The Order Sr.***
    If ADC.Mode = xNorm And GrdBag.Mode = fgmnorm Then
     '****Geeta***2.11
     If wBagRow > 0 And wBagRow <= GrdBag.Rows - 1 Then
      ADC("wOdBOdSr") = GrdBag.Value(wBagRow, "BOdSr")
       Call SetBagDet
      '*** This is just a warning. The User can generate bags with 0 Qty.
     If GrdBag.Value(wBagRow, "BOpnQty") = 0 Then DispMsg "Bag With Qty= 0 Has Been Generated", etWarning
     End If
    End If
    
  Case Is = UCase("wBgDefBPcs")
    If ADC.Mode = xNorm Then
      If wBagRow > 0 And wBagRow <= GrdBag.Rows - 1 Then
      '****** Sachin 3.02 10-12-07 - Recalc below based on Partition selected during login
      If ADC("wBgOdSrFr") = ADC("wBgOdSrTo") Then
        ADC("wBgDefBPcs") = moCn.GetFldVal("Select OdBagPcs from OrdDsg where " + _
        "OdCoCd = '" + ADC("wBCoCd") + "' and OdTc= '" + ADC("wBOdTc") + "' and OdYy='" + ADC("wBOdYy") + "' and " + _
        "OdChr= '" + ADC("wBOdChr") + "' and OdNo= " + CStr(ADC("wBOdNo")) + " and " + _
        "OdSr= " + CStr(ADC("wBgOdSrFr")) + " And OdPrtKey='" + ctCurrPrtn + "'")
      Else
        ADC("wBgDefBPcs") = 0
      End If
    End If
    End If
  End Select
End Sub
Private Sub txtSEOPwd_GotFocus()
  DispMsg "Enter SEO Password", etInfo
End Sub
Private Sub txtSEOPwd_Validate(Cancel As Boolean)


    If Trim(txtseopwd.text) <> "" Then
      If Trim(txtseopwd.text) = UCase(moCn.GetFldVal("select IsNull((select PValue from Param Join UsrAccess on UaUsrCd = PMCd " + _
                   "where PTyp= 'USR' and PMCd= '" + gs_UsrCd + "' and UaMnuCd = '" + ADC.MenuCd + "' and " + _
                   "UaSeoYn= 'Y' and PValue= '" + Trim(txtseopwd.text) + "'), " + _
                   "(Select PValue1 from Param where Ptyp= 'SEOPWD' and PmCd= 'SEOPWD' )) ")) Then
          ms_PwdFlag = True: Exit Sub
      Else
          ms_PwdFlag = False: Cancel = True: DispMsg "Invalid Edit Password / User Does Not Have SEO Rights ", etError
      End If
    Else
        If moCn.GetFldVal("Select PValue1 from Param where Ptyp = 'SEOPWD' and PmCd = 'SEOPWD'") = "" Then _
           ms_PwdFlag = True Else ms_PwdFlag = False
    End If

End Sub
Private Sub SetSEOPwdFlag()
    
End Sub
Private Sub ATXT_FldChange(Index As Integer, ByVal IdName As String, Cancel As Boolean, ErrMsg As String, ByVal pv_NewValue As Variant)
  '*** wBOdTc         valid PMCd('TC', PValue= 'SO')
  '*** wBOdYy         valid PMCd('YY')
  '*** wBOdChr        valid PMCd('CHR', adc(wBOdTc)) or ''
  '*** wBOdNo         zero or valid OmNo from OrdMst(OmTc= adc(wBOdTc), OmYy= adc(wBOdYy), OmChr= adc(wBOdChr))
  '*** wBSrt          blank or valid PMCd('SRTBAG')
  
  Select Case IdName
  '****** Sachin 2.14.0 - 04-11-2006 [Multi-Company] ******
  Case Is = UCase("wBCoCd")
    Cancel = pv_NewValue = ctSelfCoCd
    If Cancel = True Then ErrMsg = "Company Code Cannot be 'ZZZ'": Exit Sub
  
    Cancel = Not moCn.RecSeek("Select 'x' From Head Where HCd='" + ctSelfCmCd + "' And HCoCd='" + pv_NewValue + "'")
    If Cancel = True Then ErrMsg = "Invalid Company Code": Exit Sub
  
    Cancel = Not moCn.RecSeek("Select 'x' From UsrAccess Where UaMnuCd= '" + ADC.MenuCd + "' And " + _
                         "UaUsrCd = '" + gs_UsrCd + "' And UaMnuForm='MDI' And (UaValidCoCd='' Or ''','+UaValidCoCd+',''' Like '%,'+'" + pv_NewValue + "'+',%')")
      If Cancel = True Then ErrMsg = "User Doesn't Have Access To This Company Code": Exit Sub
    'last changed on 29-11
    
    Me.Caption = GetFrmCaption(pv_NewValue, "Bag Master")
  '****** Sachin 2.14.0 - 04-11-2006 [Multi-Company] ******
  
  Case Is = UCase("wBOdTc")
  'indigo.164 PValue= 'SO' changed
    Cancel = Not moCn.RecSeek("select PMCd from Param where PTyp= 'TC' and " + _
                         "PMCd = '" + pv_NewValue + "' and " + _
                         "PValue In('SO','OS') ")
    If Cancel = True Then ErrMsg = "Invalid Voucher Tc": Exit Sub
  Case Is = UCase("wBOdYy")
    Cancel = Not moCn.RecSeek("select PMCd from Param where PTyp= 'YY' and " + _
                         "PMCd = '" + pv_NewValue + "'")
    If Cancel = True Then ErrMsg = "Invalid Voucher Year": Exit Sub
  Case Is = UCase("wBOdChr")
    Cancel = pv_NewValue <> "" And (Not moCn.RecSeek("select vPSCd from vParam where " + _
                         "vPCoCd= '" + ADC("wBCoCd") + "' and vPTyp= 'CHR' and " + _
                         "vPMCd = '" + ADC("wBOdTc") + "' and vPSCd= '" + pv_NewValue + "'"))
' **** Zubin 211 **** '
'''''    ErrMsg = "Invalid Voucher Character": Exit Sub
    If Cancel = True Then ErrMsg = "Invalid Voucher Character": Exit Sub
    Dim ws_UsrArr() As String, wi_Index As Integer
    ws_UsrArr = Split(moCn.GetFldVal("Select vPValue4 from vParam where vPCoCd= '" + ADC("wBCoCd") + "' and " + _
                                        "vPTyp='CHR' and vPMCd='" + ADC("wBOdTc") + "' and vPSCd= '" + pv_NewValue + "'"), ",")
    If IsArray(ws_UsrArr) Then
      For wi_Index = 0 To UBound(ws_UsrArr)
        Cancel = IIF(UCase(gs_UsrCd) = UCase(ws_UsrArr(wi_Index)), False, True)
        If Cancel = False Then: Exit For
      Next wi_Index
    End If
   If Cancel = True Then ErrMsg = "Access To This Voucher Character Is Denied To User " + gs_UsrCd: Exit Sub
' **** Zubin 211 **** '

  Case Is = UCase("wBOdNo")
    '****** Sachin 3.02 10-12-07 - Recalc below based on Current Partition
    Cancel = (pv_NewValue <> 0) And (Not moCn.RecSeek("select OmNo from OrdMst where " + _
                                                 "OmCoCd= '" + ADC("wBCoCd") + "' and " + _
                                                 "OmTc= '" + ADC("wBOdTc") + "' and " + _
                                                 "OmYy= '" + ADC("wBOdYy") + "' and " + _
                                                 "OmChr= '" + ADC("wBOdChr") + "' and " + _
                                                 "OmNo= " + CStr(pv_NewValue) + " And OmPrtKey=" + IIF("'" + gs_Partition + "'" = "'" + ctCurrPrtn + "'", "'" + ctCurrPrtn + "'", "OmPrtKey")))
    ErrMsg = "Invalid Voucher Number": Exit Sub
    
  Case Is = UCase("wBSrt")
    Cancel = (pv_NewValue <> "") And (Not moCn.RecSeek("Select PMCd From Param Where " + _
                                      "PTyp= 'SRTBAG' And PMCd= '" + pv_NewValue + "'"))
    If Cancel = True Then ErrMsg = "Invalid Sort Option": Exit Sub

'  Case Is = UCase("wBOpnDesc")
'    Dim wBagRow As Integer
'    If GrdBag.Rows - 1 > 0 Then
'      wBagRow = IIf(GrdBag.Row > 0, GrdBag.Row, 1)
'      GrdBag.Value(wBagRow, "BOpnDesc") = pv_NewValue
'      GrdBag.SaveRec wBagRow
'      GrdBag.Row = wBagRow
'      GrdBag.COL = 1: GrdBag.SetFocus
'    End If
  
  
  '************************************************************************************
  Case Is = UCase("wBgDefBChr")
      Cancel = (Not moCn.RecSeek("Select PMCd From Param Where PTyp= 'BCHR' And " + _
                                     "PMCd = '" + pv_NewValue + "' "))
      If Cancel = True Then ErrMsg = "Invalid Bag Character": Exit Sub
      
      '*** 2.11 (Jen)
      'and PValue2= (case when PValue2= '' then '' else OrRmSCtg end)
      '****** Sachin 3.02 10-12-07 - Recalc below based on Partition selected during login
      'IG BCHR- OrRmCtg changed as Multiselect, removed OrRmCtg= PValue1*/
      'IG BCHR- if PrdTyp in BCHR it will check in the OrdRm the corresponding RmCtg, RmSctg is WSet or HSet **/
      Cancel = (Not moCn.RecSeek("Select PMCd From Param Where PTyp= 'BCHR' " + _
                         "and PMCd = '" + pv_NewValue + "' " + _
                         "and ((PValue1= '' and PValue2= '') or ((PValue1<> '' Or PValue2<> '' Or PValue3<> '') and " + _
                         "     not exists(Select 'x' from OrdDsg where OdCoCd= '" + ADC("wBCoCd") + "' and " + _
                         "            OdTc= '" + ADC("wBOdTc") + "' and OdYy= '" + ADC("wBOdYy") + "' and " + _
                         "            OdChr= '" + ADC("wBOdChr") + "' and OdNo= " + CStr(ADC("wBOdNo")) + " And OdPrtKey='" + ctCurrPrtn + "' " + _
                         IIF(ADC("wBgOdSrFr") = 0, " ", " and OdSr>= " + CStr(ADC("wBgOdSrFr")) + " ") + _
                         IIF(ADC("wBgOdSrTo") = 0, " ", " and OdSr<= " + CStr(ADC("wBgOdSrTo")) + " ") + _
                         "            and not exists(Select 'x' From OrdRm where OrPrtKey=OdPrtKey and OrOdIdNo=OdIdNo and OrCoCd= OdCoCd and OrTc= OdTc " + _
                         "            and OrYy= OdYy and OrChr= OdChr and OrNo= OdNo and OrSr= OdSr " + _
                         "            and (PValue1= '' Or ','+ PValue1+ ',' like '%,'+ OrRmCtg + ',%' ) " + _
                         "            and (PValue2= '' Or ','+ PValue2+ ',' like '%,'+ OrRmSCtg + ',%' ) " + _
                         "            and (PDesc225 ='' or PDesc225=OrSubShp) " + _
                         "            and (PValue3= '' Or (case when PValue3='WS' then OrWsQty else (case when PValue3='GS' then OrHsQty else 1 end) end)>0 )) " + _
                         "           ) ) ) "))
      If Cancel = True Then ErrMsg = "Design(s) In Range Does Not Contain Rm Ctg/ SubCtg/ PrdTyp/SubShp Of Given Flute Bag Character": Exit Sub
      '*** 2.11 (Jen) OrRmCtg= PValue1
  Case Is = UCase("wBgDefBPcs")
      Cancel = (pv_NewValue < 0)
      If Cancel = True Then ErrMsg = "Invalid Bag Pieces": Exit Sub
  '************************************************************************************
  
  ' **** Zubin 211 **** '
  '*** Jay 2.13(CT) ***
  Case Is = UCase("wBagTyp")
    Cancel = (pv_NewValue <> "") And (Not moCn.RecSeek("Select 'x' From Param Where PTyp= 'BTYP' And " + _
                                     "PMCd = '" + pv_NewValue + "' and PSCd= ''"))
    ErrMsg = "Invalid Primary/Flute Bag Option. Enter 'P' For Primary Bags, 'F' For Flute Bags, 'C' For Component Bags and Blank For All Bags": Exit Sub
  '*** Jay 2.13(CT) ***
  ' **** Zubin 211 **** '
  
  'Sachin 3.11.0
  Case Is = UCase("wBgDefBOpnLoc")
      Cancel = (pv_NewValue <> "") And (Not moCn.RecSeek("Select LocCd From Loc Where " + _
                                    " LocCoCd= '" + ADC("wBCoCd") + "' and LocTyp= 'P' " + _
                                    " and LocCd= '" + pv_NewValue + "' and LocValidYN='Y'"))
      If Cancel = True Then ErrMsg = "Invalid Opening Location": Exit Sub
  
  
  End Select
End Sub

Private Sub GrdBag_RowWhen(ByVal RowNum As Integer)
  '*** If New Row then
      '   If Row = 1 then
      '       Get default values for BYy & BChr from Head table. BOdSr = min OdSr from
      '       OrdDsg(OdTc= adc(wBOdTc), OdYy= adc(wBOdYy), OdChr= adc(wBOdChr), OdNo= adc(wBOdNo), OdHld= 'N')
      '       BOpnLoc= First location in the production sequence for the Design Sr
      '   Else
      '       Get default values for BYy & BChr, BOdSr and BOpnLoc from the previous record
      'BOpnDt = Today's Date
      'BCls = 'N'
  '*** adc(wBOpnDesc) = BOpnDesc
  '*** Displaying the Bag and Order details at the foot of the frame
      'BagNo   -  display Current Bag Qty, Part Qty
      'OrderSr -  display sum of Current Bag Qty, sum of current Part Qty & Bal Qty
      'If new row then BOpnQty = IIf(OdBagPcs from OrdDsg <= adc(wOdBalQty), OdBagPcs, adc(wOdBalQty))
  
  Dim wBagRow As Integer, wBagOpnQty As Integer, ws_BTyp As String, wOdLnId As Integer
  
  With GrdBag
    If .IsNew(RowNum) Then
      .Value(RowNum, "BPrtKey") = ctCurrPrtn '****** Sachin 3.02 19-12-07 - Default value 'C'
      If RowNum = 1 Then
        .Value(RowNum, "BYy") = moCn.GetFldVal("Select hYy from Head where HCoCd= '" + ADC("wBCoCd") + "' and HCd= '" + ctSelfCmCd + "'")
        .Value(RowNum, "BChr") = moCn.GetFldVal("Select hBChr from Head where HCoCd= '" + ADC("wBCoCd") + "' and HCd= '" + ctSelfCmCd + "'")
        '****** Sachin 3.02 10-12-07 - Rowwhen below based on Partition selected during login
        .Value(RowNum, "BOdSr") = moCn.GetFldVal("Select min(OdSr) from OrdDsg where " + _
              "OdCoCd= '" + ADC("wBCoCd") + "' and " + _
              "OdTc= '" + ADC("wBOdTc") + "' and OdYy= '" + ADC("wBOdYy") + "' and " + _
              "OdChr= '" + ADC("wBOdChr") + "' and OdNo= " + CStr(ADC("wBOdNo")) + " and OdPrtKey='" + ctCurrPrtn + "' and " + _
              "OdHld= 'N' ")
            
        '*** Jay 2.13(CT) ***
        ws_BTyp = moCn.GetFldVal("Select PValue from Param where PTyp= 'BCHR' and " + _
                      "PMCd= '" + .Value(RowNum, "BChr") + "' and PSCd= '' ")
        Call SetBagOpnLoc(RowNum, ws_BTyp)
        '.Value(RowNum, "BOdCtCd") = moCn.GetFldVal("Select OdOmCtCd from OrdDsg where " + _
                                "OdTc= '" + adc("wBOdTc") + "' and " + _
                                "OdYy= '" + adc("wBOdYy") + "' and " + _
                                "OdChr= '" + adc("wBOdChr") + "' and " + _
                                "OdNo= " + CStr(adc("wBOdNo")) + " and " + _
                                "OdSr= " + CStr(.Value(RowNum, "BOdSr")))
    'EmrSv
       ' .Value(RowNum, "BOdLnIdKey") = moCn.GetFldVal("Select IdKey from OrdLn where OdlCoCd ='" + adc("wBCoCd") + "' and " + _
                " OdlTc ='" + adc("wBOdTc") + "' and OdlYy= '" + adc("wBOdYy") + "' and " + _
                " OdlChr='" + adc("wBOdChr") + "' and OdlNo= " + CStr(adc("wBOdNo")) + " and " + _
                " OdlSr= " + CStr(.Value(RowNum, "BOdSr")) + " and OdlSrNo=1")
      Else
        .Value(RowNum, "BYy") = .Value(RowNum - 1, "BYy")
        ' Swap By Jay ('*** Jay 2.13(CT) ***)
        .Value(RowNum, "BOdSr") = .Value(RowNum - 1, "BOdSr")
        .Value(RowNum, "BChr") = .Value(RowNum - 1, "BChr")
        ' Swap By Jay ('*** Jay 2.13(CT) ***)
        '*** Jay 2.13(CT) ***
        ws_BTyp = moCn.GetFldVal("Select PValue from Param where PTyp= 'BCHR' and " + _
                      "PMCd= '" + .Value(RowNum, "BChr") + "' and PSCd= '' ")
        If ws_BTyp = "C" Then
          Call SetBagOpnLoc(RowNum, ws_BTyp)
        Else
          .Value(RowNum, "BOpnLoc") = .Value(RowNum - 1, "BOpnLoc")
        End If
        
        '*** Jay 2.13(CT) ***
        .Value(RowNum, "BOdCtCd") = .Value(RowNum - 1, "BOdCtCd")
        .Value(RowNum, "BOdlSrNo") = .Value(RowNum - 1, "BOdlSrNo")
      End If
      '53 company code parameter added
      .Value(RowNum, "BOpnDt") = GetNextDate(ADC("wBCoCd")) 'moCn.SrvrDate
      '''(Old Code) .Value(RowNum, "BNo") = moCn.AutoGen("Bag", "BNo", .Value(RowNum, "BYy") + .Value(RowNum, "BChr"))
      .Value(RowNum, "BCls") = "N"
      
      'EmrSv
      '.Value(RowNum, "BOdLnIdKey") = moCn.GetFldVal("Select IdKey from OrdLn where OdlCoCd ='" + adc("wBCoCd") + "' and " + _
                " OdlTc ='" + adc("wBOdTc") + "' and OdlYy= '" + adc("wBOdYy") + "' and " + _
                " OdlChr='" + adc("wBOdChr") + "' and OdlNo= " + CStr(adc("wBOdNo")) + " and " + _
                " OdlSr= " + CStr(.Value(RowNum, "BOdSr")) + " and OdlSrNo=1")
    Else
        ws_BTyp = moCn.GetFldVal("Select PValue from Param where PTyp= 'BCHR' and " + _
                                  "PMCd= '" + .Value(RowNum, "BChr") + "' and PSCd= '' ")
    End If
    ADC("wBOpnDesc") = .Value(RowNum, "BOpnDesc")
    
    wBagRow = GrdBag.Row
    ADC("wBdBag") = .Value(wBagRow, "BYy") + "/" + .Value(wBagRow, "BChr") + _
                    "/" + Trim(CStr(.Value(wBagRow, "BNo")))

    '*** Getting Current Bag Qty, Part Qty From Bag ***
    '****** Sachin 3.02 10-12-07 - Selected partition
    ADC("wBdBQty") = moCn.GetFldVal("Select BQty from Bag where " + _
                                   "BCoCd= '" + ADC("wBCoCd") + "' and " + _
                                   "BYy='" + .Value(wBagRow, "BYy") + "' and " + _
                                   "BChr='" + .Value(wBagRow, "BChr") + "' and " + _
                                   "BNo= " + CStr(.Value(wBagRow, "BNo")) + IIF(GrdBag.Value(wBagRow, "BPrtKey") = ctCurrPrtn, " and BPrtKey='" + ctCurrPrtn + "'", ""))
                                   
    ADC("wBdPtQty") = moCn.GetFldVal("Select BPtQty from Bag where " + _
                                    "BCoCd= '" + ADC("wBCoCd") + "' and " + _
                                    "BYy='" + .Value(wBagRow, "BYy") + "' and " + _
                                    "BChr='" + .Value(wBagRow, "BChr") + "' and " + _
                                    "BNo= " + CStr(.Value(wBagRow, "BNo")) + IIF(GrdBag.Value(wBagRow, "BPrtKey") = ctCurrPrtn, " and BPrtKey='" + ctCurrPrtn + "'", ""))
        
    '*** Jay 2.13(CT) ***
    '.ColProp("BOpnPtQty").Style = IIF(ws_BTyp <> "C", fgcReadOnly, fgcNormal)     '*** (Bef 27/11/06)
    .ColProp("BOpnPtQty").Style = fgcNormal   '*** (Jen 27/11/06)
    '*** Jay 2.13(CT) ***
    
    'Bef 2.11
'      '*** Getting Current Bag Qty, Part Qty From Bag For The Order Sr.***
'    Dim wRsOrdDet As MwfLib.MDORowSet
'    ADC("wOdBOdSr") = .Value(wBagRow, "BOdSr")
'    Set wRsOrdDet = moCn.OpenRes("Select IsNull(sum(BQty), 0) as qTotQty, IsNull(sum(BPtQty),0) as qTotPtQty " + _
'                                  "from Bag where BCoCd='" + ADC("wBCoCd") + "' and " + _
'                                  "BOdTc= '" + ADC("wBOdTc") + "' and " + _
'                                  "BOdYy= '" + ADC("wBOdYy") + "' and " + _
'                                  "BOdChr= '" + ADC("wBOdChr") + "' and " + _
'                                  "BOdNo= " + CStr(ADC("wBOdNo")) + " and " + _
'                                  "BOdSr= " + CStr(.Value(wBagRow, "BOdSr")))
'    If Not (wRsOrdDet.EOF Or wRsOrdDet.BOF) Then
'      ADC("wOdBQty") = wRsOrdDet!qTotQty
'      ADC("WODCTPTQTY") = wRsOrdDet!qTotPtQty
'    End If
'    ADC("wOdBalQty") = moCn.GetFldVal("Select OdPrdQty from OrdDsg where " + _
'                                     "OdCoCd= '" + ADC("wBCoCd") + "' and " + _
'                                     "OdTc= '" + ADC("wBOdTc") + "' and " + _
'                                     "OdYy= '" + ADC("wBOdYy") + "' and " + _
'                                     "OdChr= '" + ADC("wBOdChr") + "' and " + _
'                                     "OdNo= " + CStr(ADC("wBOdNo")) + " and " + _
'                                     "OdSr= " + CStr(.Value(wBagRow, "BOdSr"))) - ADC("wOdBQty")
'
'    If .IsNew(RowNum) Then
'      wBagOpnQty = moCn.GetFldVal("Select OdBagPcs from OrdDsg where " + _
'                   "OdCoCd='" + ADC("wBCoCd") + "' and OdTc= '" + ADC("wBOdTc") + "' and " + _
'                   "OdYy= '" + ADC("wBOdYy") + "' and OdChr= '" + ADC("wBOdChr") + "' and " + _
'                   "OdNo= " + CStr(ADC("wBOdNo")) + " and OdSr= " + CStr(.Value(RowNum, "BOdSr")))
'
'      wBagOpnQty = IIF(wBagOpnQty <= ADC("wOdBalQty"), wBagOpnQty, ADC("wOdBalQty"))
'      .Value(RowNum, "BOpnQty") = wBagOpnQty
'
'      '*** This is just a warning. The User can generate bags with 0 Qty.
'      If wBagOpnQty = 0 Then DispMsg "Bag With Qty= 0 Has Been Generated", etWarning
'    End If
'    Set wRsOrdDet = Nothing
    
    ADC("wOdBOdSr") = .Value(wBagRow, "BOdSr")
    '****Geeta***2.11
    '*** Getting Current Bag Qty, Part Qty From Bag For The Order Sr.***
    Call SetBagDet
  End With
End Sub
Private Sub GrdBag_ColWhen(ByVal RowNum As Integer, ByVal ColName As String, Cancel As Boolean, ErrMsg As String, ByVal Action As MwfCtl.en_fgAction)
  '*** Cannot edit Bag details if transaction exists for the bag
  '*** Cannot edit BYy, BChr, BNo, BOdSr, BOpnDt when the bag already exists
  '*** Set Helps For Each Field
  
  With GrdBag
    '****** Sachin 3.02 10-12-07 - Locking Bags of Previous Partition
    If ms_OrdLck = True Or .Value(RowNum, "BPrtKey") <> ctCurrPrtn Then
        Cancel = True: ErrMsg = "Cannot Edit Bags In Previous Partition"
        CmdBagGen.Enabled = False
        Exit Sub
    Else
        Cancel = False: CmdBagGen.Enabled = True
    End If
      
    '****** Current Partition as previous transactions are locked
    If UCase(ColName) <> UCase("BOdCtCd") And moCn.GetFldVal("Select BLstYyKey From Bag Where " + _
            "BCoCd='" + ADC("wBCoCd") + "' and BYy= '" + .Value(RowNum, "BYy") + "' and " + _
            "BChr= '" + .Value(RowNum, "BChr") + "' and BNo= " + CStr(.Value(RowNum, "BNo")) + " and BPrtKey='" + ctCurrPrtn + "'") <> "" Then _
       Cancel = True: ErrMsg = "Cannot Edit This Value as Transaction Exists": Exit Sub

    Select Case ColName
    Case Is = UCase("BYy")
      If Not .IsNew(RowNum) Then Cancel = True: ErrMsg = "Cannot Edit This Value": Exit Sub
      Call HlpList.PMCd("YY")
    Case Is = UCase("BChr")
      If Not .IsNew(RowNum) Then Cancel = True: ErrMsg = "Cannot Edit This Value": Exit Sub
      'umaa
      'Call HlpList.PMCd("BChr")
      If ADC("WBODTC") <> "PD" Then Call HlpList.PMCd("BChr", , , "PValue <> 'D'")
      If ADC("WBODTC") = "PD" Then Call HlpList.PMCd("BChr", "'D'")
            
    Case Is = UCase("BNo")
      If Not .IsNew(RowNum) Then Cancel = True: ErrMsg = "Cannot Edit This Value": Exit Sub
      If .IsNew(RowNum) Then
        If moCn.GetFldVal("Select hBagNoEditYN from Head where HCoCd = '" + ADC("wBCoCd") + "'") = "N" Then
            Cancel = True: ErrMsg = "Cannot Edit This Value": Exit Sub
        End If
      End If
    Case Is = UCase("BOdSr")
      If Not .IsNew(RowNum) Then Cancel = True: ErrMsg = "Cannot Edit This Value": Exit Sub
      Call HlpList.OdSr(ADC("wBCoCd"), ADC("wBOdTc"), ADC("wBOdYy"), ADC("wBOdChr"), ADC("wBOdNo"), , ctCurrPrtn)
    Case Is = UCase("BOpnDt")
      If Not .IsNew(RowNum) Then Cancel = True: ErrMsg = "Cannot Edit This Value": Exit Sub
    Case Is = UCase("BOpnLoc")
      Call HlpList.Loc(ADC("wBCoCd"), "'P'")
'    Case Is = UCase("BOpnDesc")
'      ADC("wBOpnDesc").SetFocus
    Case Is = UCase("BOdCtCd")
      Call HlpList.CustCd("T")
    'Uma EmrSv2
    Case Is = UCase("BOdLSrNo")
    If moCn.RecSeek("Select 'x' From SysDb where SysId='UNI' and (SdCoCd ='' or charindex(','+'" + ADC("wBOdChr") + "'+',',','+SdCoCd+',')>0)") Then
      Call HlpList.OdlSrNo(ADC("wBOdTc"), ADC("wBOdYy"), ADC("wBOdChr"), ADC("wBOdNo"), CStr(.Value(RowNum, "BODSR")), 0, 0)
    Else
      Cancel = True
      ErrMsg = "No need to enter OrdLn"
    End If
    End Select
  End With
End Sub
Private Sub GrdBag_ReCalc(ByVal RowNum As Integer, ByVal ColName As String, ByVal OldVal As Variant, ByVal SourceName As String, ByVal SourceColName As String)
  '*** BNo on BYy,BChr                           Normal, Copy
      'If BYy and BChr<> Blank then BNo = Autogenerated Number (BYy, BChr are the PreKey)
  '*** BOdDmCd on BOdSr                          Normal
      'BOdDmCd = OdDmCd from OrdDsg(OdTc= adc(wBOdTc),OdYy= adc(wBOdYy), OdChr= adc(wBOdChr),
      '          OdNo= adc(wBOdNo), OdSr= BOdSr)
  '*** BOdSfx on BOdSr                           Normal
      'BOdSfx = OdSfx from OrdDsg(OdTc= adc(wBOdTc),OdYy= adc(wBOdYy), OdChr= adc(wBOdChr),
      '         OdNo= adc(wBOdNo), OdSr= BOdSr)
  '*** BOdDmSz on BOdSr                          Normal
      'BOdDmSz = OdDmSz from OrdDsg(OdTc= adc(wBOdTc),OdYy= adc(wBOdYy), OdChr= adc(wBOdChr),
      '          OdNo= adc(wBOdNo), OdSr= BOdSr)
  '*** BLoc on BOpnLoc                           Normal
      'BLoc = BOpnLoc
  '*** BOpnDesc on wBOpnDesc                     Normal
      'if current Row > 0 then if BOpnDesc<> adc(wBOpnDesc) then BOpnDesc= adc(wBOpnDesc)
      'Save The Current Row
  
  ' **** Zubin 211 **** '
  Dim wBagOpnQty As Single, wAGBNo As Single, wMaxBNo As Single
  ' **** Zubin 211 **** '
  With GrdBag
    Select Case UCase(ColName)
    Case Is = UCase("BNo")
      If .Mode = fgmnorm Or .Mode = fgmCopy Then
      ' **** Zubin 211 **** '
      '''''        If .Value(RowNum, "BYy") <> "" And .Value(RowNum, "BChr") <> "" Then _
      '''''          .Value(RowNum, "BNo") = moCn.AutoGen("Bag", "BNo", ADC("wBCoCd") + .Value(RowNum, "BYy") + .Value(RowNum, "BChr"))

      '****** Sachin 3.02 10-12-07 - Current Db (Number
        If .Value(RowNum, "BYy") <> "" And .Value(RowNum, "BChr") <> "" Then
           wAGBNo = moCn.GetFldVal("Select IsNull(aNum,0)+ 1 From Autogen Where aTable='Bag' And aFld='BNo' " + _
                    " And aPreKey='" + ADC("wBCoCd") + .Value(RowNum, "BYy") + .Value(RowNum, "BChr") + "'")
           wMaxBNo = moCn.GetFldVal("select IsNull(max(BNo), 0)+ 1 from Bag where " + _
                                       "BCoCd= '" + ADC("wBCoCd") + "' and BYy= '" + .Value(RowNum, "BYy") + "' and " + _
                                       "BChr= '" + .Value(RowNum, "BChr") + "' " + IIF(gs_Partition = ctCurrPrtn, " And BPrtKey='" + ctCurrPrtn + "'", ""))
           .Value(RowNum, "BNo") = IIF(wAGBNo > wMaxBNo, wAGBNo, wMaxBNo)
        End If
      ' **** Zubin 211 **** '
      End If
         
    Case Is = UCase("BOdDmCd")
      '****** Sachin 3.02 10-12-07 - recalc below on current partition
      If .Mode = fgmnorm Then _
         .Value(RowNum, "BOdDmCd") = moCn.GetFldVal("Select OdDmCd from OrdDsg where " + _
                                          "OdCoCd= '" + ADC("wBCoCd") + "' and " + _
                                          "OdTc= '" + ADC("wBOdTc") + "' and " + _
                                          "OdYy= '" + ADC("wBOdYy") + "' and " + _
                                          "OdChr= '" + ADC("wBOdChr") + "' and " + _
                                          "OdNo= " + CStr(ADC("wBOdNo")) + " and " + _
                                          "OdSr= " + CStr(.Value(RowNum, "BOdSr")) + IIF(gs_Partition = ctCurrPrtn, " and OdPrtKey='" + ctCurrPrtn + "'", ""))
    Case Is = UCase("BOdSfx")
      '****** Sachin 3.02 10-12-07 - recalc below on current partition
      If .Mode = fgmnorm Then _
         .Value(RowNum, "BOdSfx") = moCn.GetFldVal("Select OdSfx from OrdDsg where " + _
                                         "OdCoCd= '" + ADC("wBCoCd") + "' and " + _
                                         "OdTc= '" + ADC("wBOdTc") + "' and " + _
                                         "OdYy= '" + ADC("wBOdYy") + "' and " + _
                                         "OdChr= '" + ADC("wBOdChr") + "' and " + _
                                         "OdNo= " + CStr(ADC("wBOdNo")) + " and " + _
                                         "OdSr= " + CStr(.Value(RowNum, "BOdSr")) + IIF(gs_Partition = ctCurrPrtn, " and OdPrtKey='" + ctCurrPrtn + "'", ""))

    Case Is = UCase("BOdDmSz")
      '****** Sachin 3.02 10-12-07 - recalc below on current partition
      If .Mode = fgmnorm Then _
         .Value(RowNum, "BOdDmSz") = moCn.GetFldVal("Select OdDmSz from OrdDsg where " + _
                                                   "OdCoCd= '" + ADC("wBCoCd") + "' and " + _
                                                   "OdTc= '" + ADC("wBOdTc") + "' and " + _
                                                   "OdYy= '" + ADC("wBOdYy") + "' and " + _
                                                   "OdChr= '" + ADC("wBOdChr") + "' and " + _
                                                   "OdNo= " + CStr(ADC("wBOdNo")) + " and " + _
                                                   "OdSr= " + CStr(.Value(RowNum, "BOdSr")) + IIF(gs_Partition = ctCurrPrtn, " and OdPrtKey='" + ctCurrPrtn + "'", ""))


'    ****** Sachin 3.02 (10/09/07) ******
'    Case Is = UCase("BOdIdNo")
'      If .Mode = fgmnorm Then
'          .Value(RowNum, "BOdIdNo") = moCn.GetFldVal("Select OdIdNo From OrdDsg " + _
'                                          "Where OdCoCd='" + ADC("wBCoCd") + "' And OdTc='" + ADC("wBOdTc") + "' " + _
'                                          "  And OdYy='" + ADC("wBOdYy") + "' And OdChr='" + ADC("wBOdChr") + "' And OdNo=" + CStr(ADC("wBOdNo")) + " And OdSr= " + CStr(.Value(RowNum, "BOdSr")) + " and OdPrtKey='" + ctCurrPrtn + "'")
'      End If
'
'    Case Is = UCase("BOmIdNo")
'      If .Mode = fgmnorm Then
'          .Value(RowNum, "BOmIdNo") = moCn.GetFldVal("Select OmIdNo From OrdMst " + _
'                                          "Where OmCoCd='" + ADC("wBCoCd") + "' And OmTc='" + ADC("wBOdTc") + "' " + _
'                                          "  And OmYy='" + ADC("wBOdYy") + "' And OmChr='" + ADC("wBOdChr") + "' And OmNo=" + CStr(ADC("wBOdNo")) + " and OmPrtKey='" + ctCurrPrtn + "'")
'      End If
'
'    Case Is = UCase("BDmIdNo")
'        If .Mode = fgmnorm Then
'          .Value(RowNum, "BDmIdNo") = moCn.GetFldVal("Select DmIdNo From DsgMst " + _
'                                          "Where DmTcTyp='DM' And DmCd='" + .Value(RowNum, "BOdDmCd") + "' " + _
'                                          "  And DmSz='' And DmPrtKey='" + ctCurrPrtn + "'")
'        End If
    '****** Sachin 3.02 (10/09/07) ******

    Case Is = UCase("BLoc")
      If .Mode = fgmnorm Then .Value(RowNum, "BLoc") = .Value(RowNum, "BOpnLoc")
    Case Is = UCase("BOdCtCd")
      If .Mode = fgmnorm Then
        If OldVal <> .Value(RowNum, "BOdSr") Then
          '****** Sachin 3.02 10-12-07 - Recalc below on current partition
          .Value(RowNum, "BOdCtCd") = moCn.GetFldVal("Select OdOmCtCd from OrdDsg where " + _
                                                     "OdCoCd= '" + ADC("wBCoCd") + "' and " + _
                                                     "OdTc= '" + ADC("wBOdTc") + "' and " + _
                                                     "OdYy= '" + ADC("wBOdYy") + "' and " + _
                                                     "OdChr= '" + ADC("wBOdChr") + "' and " + _
                                                     "OdNo= " + CStr(ADC("wBOdNo")) + " and " + _
                                                     "OdSr= " + CStr(.Value(RowNum, "BOdSr")) + IIF(gs_Partition = ctCurrPrtn, " and OdPrtKey='" + ctCurrPrtn + "'", ""))
        End If
      End If
    Case Is = UCase("BOpnDesc")
      If .Mode = fgmnorm Then
        If .Row > 0 Then
          If .Value(.Row, "BOpnDesc") <> ADC("wBOpnDesc") Then
            .Value(.Row, "BOpnDesc") = ADC("wBOpnDesc")
            .SaveRec .Row
          End If
        End If
      End If

    ' **** Zubin 211 **** '
    Case Is = UCase("BOpnQty")
    '****GEETA***2.11
      If ADC.Mode = xNorm And .Mode = fgmnorm Then Call SetBagDet
    '*** jay 2.13(CT) ***
    Case Is = UCase("BOpnPtQty")
      If .Mode = fgmnorm Then
        Dim Wi_CTPrts As Integer
        '*** Jay 2.14Next *** [Multiple Comp in One BCHR]
        '****** Sachin 3.02 10-12-07 - Recalc below on corresponding partition
        Wi_CTPrts = moCn.GetFldVal("Select Sum(OctParts) as OctParts from OrdCT where " + _
                        "OctCoCd= '" + ADC("wBCoCd") + "' and OctTc= '" + ADC("wBOdTc") + "' and " + _
                        "OctYy= '" + ADC("wBOdYy") + "' and OctChr= '" + ADC("wBOdChr") + "' and " + _
                        "OctNo= " + CStr(ADC("wBOdNo")) + " and OctSr= " + CStr(.Value(RowNum, "BOdSr")) + IIF(gs_Partition = ctCurrPrtn, " and OctPrtKey='" + ctCurrPrtn + "'", "") + " and " + _
                        "OctCtChr ='" + .Value(RowNum, "BCHR") + "' " + _
                        "Group By OctCoCd, OctTc, OctYy, OctChr, OctNo, OctSr, OctCtChr")
        '*** Jay 2.14Next *** [Multiple Comp in One BCHR]
        If .Value(RowNum, "BOpnPtQty") <> Wi_CTPrts * .Value(RowNum, "BOpnQty") Then _
                                      .Value(RowNum, "BOpnPtQty") = Wi_CTPrts * .Value(RowNum, "BOpnQty")
      End If
    Case Is = UCase("BOpnLoc")
      If .Mode = fgmnorm Then
        Call SetBagOpnLoc(RowNum, moCn.GetFldVal("Select PValue from Param where PTyp= 'BCHR' and " + _
                      "PMCd= '" + .Value(RowNum, "BChr") + "' and PSCd= '' "))
      End If
    '*** jay 2.13(CT) ***
  'EmrSv
      Case Is = UCase("BODLNIDKEY")
        If (.Mode = fgmnorm Or .Mode = fgmCopy Or .Mode = fgmCopyOver) Then
          If moCn.RecSeek("Select 'x' From SysDb where SysId='UNI'") Then
            If RowNum > 0 And RowNum <= GrdBag.Rows - 1 Then
              .Value(RowNum, "BOdLnIdKey") = moCn.GetFldVal("Select IdKey from OrdLn where OdlCoCd ='" + ADC("wBCoCd") + "' and " + _
                  " OdlTc ='" + ADC("wBOdTc") + "' and OdlYy= '" + ADC("wBOdYy") + "' and " + _
                  " OdlChr='" + ADC("wBOdChr") + "' and OdlNo= " + CStr(ADC("wBOdNo")) + " and " + _
                  " OdlSr= " + CStr(.Value(RowNum, "BOdSr")) + " and OdlSrNo=" + CStr(.Value(RowNum, "BOdlSrNo")))
            End If
          End If
        End If
      End Select
  End With
End Sub
Private Sub GrdBag_ColChange(ByVal RowNum As Integer, ByVal ColName As String, ByVal NewValue As Variant, Cancel As Boolean, ErrMsg As String)
  '*** BYy              Blank or PMCd('YY')
  '*** BChr             Blank or PMCd('BCHR')
  '*** BNo              Write the value of the BNo field in the autogen table if it is
  '                     greater than the one already existing in the Autogen table for the
  '                     same Yy,Chr combination
  '                     (Ignore if it is less than the one already existing in the autogen
  '                     table)
  '*** BOdSr            valid OdSr from OrdDsg(OdTc= adc(wBOdTc), OdYy= adc(wBOdYy),
  '                     OdChr= adc(wBOdChr), OdNo= adc(wBOdNo))
  '*** BOpnDt           Not Blank
  '*** BOpnLoc          Blank or LocCd from Loc(LocTyp= 'P')
  
  With GrdBag
    Select Case ColName
    Case Is = UCase("BYy")
      Cancel = (NewValue <> "") And (Not moCn.RecSeek("Select PMCd From Param Where PTyp= 'YY' And " + _
                                     "PMCd = '" + NewValue + "'"))
      If Cancel = True Then ErrMsg = "Invalid Bag Year": Exit Sub
    '*** Jay 2.13(CT) ***(Replaced 'Y' with 'F')
    Case Is = UCase("BChr")
    Dim ws_BTyp As String
      Cancel = (NewValue <> "") And (Not moCn.RecSeek("Select PMCd From Param Where PTyp= 'BCHR' And " + _
                                     "PMCd = '" + NewValue + "'"))
                                     
      If Cancel = True Then ErrMsg = "Invalid Bag Character": Exit Sub
      'uma
      If ADC("WBODTC") <> "PD" Then
      Cancel = (Not moCn.RecSeek("Select PMCd From Param Where PTyp= 'BCHR' And PValue <> 'D' and " + _
                                     "PMCd = '" + NewValue + "'"))
      If Cancel = True Then ErrMsg = "Invalid Bag Character": Exit Sub
      End If
      If ADC("WBODTC") = "PD" Then
      Cancel = (Not moCn.RecSeek("Select PMCd From Param Where PTyp= 'BCHR' And PValue = 'D' and " + _
                                   "PMCd = '" + NewValue + "'"))
      If Cancel = True Then ErrMsg = "Invalid Bag Character for PD Bag": Exit Sub
      End If
           
      ' **** Zubin 211 **** '
      ' Bag Chr cannot be a flute bag chr when hFluteBagYN= 'N' in Head table
      '***Geeta 2.11 BTyp
      '*** Jay 2.13(CT) ***(Replaced 'Y' with 'F')
      'Cannot open Component Bags if hCompBagsYn= 'N' already checked @ parm level
      ws_BTyp = moCn.GetFldVal("Select PValue from Param where PTyp= 'BCHR' and " + _
                      "PMCd= '" + NewValue + "' and PSCd= '' ")
      If ws_BTyp = "F" Then
        Cancel = moCn.RecSeek("Select 'x' from Head where HCoCd= '" + ctSelfCoCd + "' and " + _
                      "HCd='" + ctSelfCmCd + "' and hFluteBagYN= 'N' ")
        If Cancel = True Then ErrMsg = "Flute Bag Cannot Be Generated As The 'Allow Flute Bags' Option In Head Table Is 'N'": Exit Sub
      ElseIf ws_BTyp = "C" Then
        If .Mode = fgmnorm Then
          '****** Sachin 3.02 10-12-07 - Check below on Corresponding Partition
          Cancel = Not moCn.RecSeek("Select OctCtChr from OrdCT where " + _
                        "OctCoCd= '" + ADC("wBCoCd") + "' and OctTc= '" + ADC("wBOdTc") + "' and " + _
                        "OctYy= '" + ADC("wBOdYy") + "' and OctChr= '" + ADC("wBOdChr") + "' and " + _
                        "OctNo= " + CStr(ADC("wBOdNo")) + " and OctSr= " + CStr(.Value(RowNum, "BOdSr")) + " and OctPrtKey='" + ctCurrPrtn + "' and " + _
                        "OctCtChr ='" + NewValue + "'")
          If Cancel = True Then ErrMsg = "Cannot Open This Component Bag, Component Is Not Present In Order"
        End If
      End If
      'Geeta***2.11
      '.ColProp("BOpnPtQty").Style = IIF(ws_BTyp <> "C", fgcReadOnly, fgcNormal)    '*** (Bef 27/11/06)
      .ColProp("BOpnPtQty").Style = fgcNormal     '*** (Jen 27/11/06)
      '*** Jay 2.13(CT) ***
    Case Is = UCase("BNo")
' **** Zubin 211 **** '
'''''      Call moCn.AutoGen("Bag", "BNo", ADC("wBCoCd") + .Value(RowNum, "BYy") + .Value(RowNum, "BChr"), NewValue)
' **** Zubin 211 **** '
    Case Is = UCase("BOdSr")
      '****** Sachin 3.02 10-12-07 - Corresponding Partition
      Cancel = Not moCn.RecSeek("Select OdSr From OrdDsg Where " + _
                           "OdCoCd= '" + ADC("wBCoCd") + "' and " + _
                           "OdTc= '" + ADC("wBOdTc") + "' and " + _
                           "OdYy= '" + ADC("wBOdYy") + "' and " + _
                           "OdChr= '" + ADC("wBOdChr") + "' and " + _
                           "OdNo= " + CStr(ADC("wBOdNo")) + " and " + _
                           "OdSr= " + CStr(NewValue) + " and OdPrtKey='" + ctCurrPrtn + "'")
      ErrMsg = "Invalid Order Sr./Order Sr. Not Present In Current Partition": Exit Sub
                                                       
    Case Is = UCase("BOpnDt")
      If NewValue = "" Then Cancel = True: ErrMsg = "Date Cannot Be Blank": Exit Sub
      If ms_LockDt <> "01/01/80" And CDate(ms_LockDt) > CDate(NewValue) Then _
        Cancel = True: ErrMsg = "Opening Date Cannot Be Less Than Lock Date.": Exit Sub
        
      '****** Sachin 3.02 19-12-07 - Bag Opening Date should Greater than the Purge Date
      If ms_PurgeDt <> "01/01/80" And CDate(ms_PurgeDt) >= CDate(NewValue) Then _
        Cancel = True: ErrMsg = "Opening Date Cannot Be Less Than Or Equal To Purge Date.": Exit Sub
        
    Case Is = UCase("BOpnLoc")
      Cancel = (NewValue <> "") And (Not moCn.RecSeek("Select LocCd From Loc where " + _
                                    "LocCoCd= '" + ADC("wBCoCd") + "' and LocTyp= 'P' " + _
                                    "and LocCd= '" + NewValue + "' and LocValidYN='Y'"))
      If Cancel = True Then ErrMsg = "Invalid Opening Location": Exit Sub
      
    Case Is = UCase("BOdCtCd")
      Cancel = (NewValue <> "") And (Not moCn.RecSeek("select CmCd from CustMst where " + _
                                         "CmCtg= 'T' and CmCd = '" + NewValue + "'"))
      ErrMsg = "Invalid Contractor Code": Exit Sub
    
    Case Is = UCase("BOdLSrNo")
      If moCn.RecSeek("Select 'x' From SysDb where SysId='UNI' and (SdCoCd ='' or charindex(','+'" + ADC("wBOdChr") + "'+',',','+SdCoCd+',')>0) ") Then
        Cancel = (NewValue = 0)
        ErrMsg = "OrdLn must be specified"
        
        Cancel = Not moCn.RecSeek("Select 'a' From OrdLn Where " + _
                             "OdlCoCd= '" + ADC("wBCoCd") + "' and " + _
                             "OdlTc= '" + ADC("wBOdTc") + "' and " + _
                             "OdlYy= '" + ADC("wBOdYy") + "' and " + _
                             "OdlChr= '" + ADC("wBOdChr") + "' and " + _
                             "OdlNo= " + CStr(ADC("wBOdNo")) + " and " + _
                             "OdlSr= " + CStr(GrdBag.Value(RowNum, "BOdSr")) + _
                             " and OdlSrNo=" + CStr(NewValue))
  
        ErrMsg = "Invalid OrdLn": Exit Sub
      End If

    '*** (Bef 27/11/06)
    ''*** jay 2.13(CT) ***
    'Case Is = UCase("BOpnPtQty")
    '  If .Mode = fgmnorm Then
    '    Cancel = moCn.GetFldVal("Select PValue from Param where PTyp= 'BCHR' and  PMCd = '" + GrdBag.Value(RowNum, "BCHR") + "' and PSCd = ''") <> "C"
    '    If Cancel = True Then DispMsg "Part Quantity Is Only Editable For Component Bags", etError: Exit Sub
    '  End If
    ' '*** jay 2.13(CT) ***
     '*** (Bef 27/11/06)
    End Select
  End With
End Sub
Private Sub GrdBag_RowSave(ByVal RowNum As Integer, Cancel As Boolean, ErrMsg As String)
  '*** If the record with PTyp= 'WORK' and PMCd= 'ZSELF' does not exist in param then
      'give a message to the user to enter the record in Param
  If Not moCn.RecSeek("Select vPMCd From vParam Where vPCoCd = '" + ADC("wBCoCd") + "' and " + _
                            "vPTyp= 'WORK' And vPMCd = '" + ctTdWrk + "'") Then _
     Cancel = True: ErrMsg = "Enter Worker 'ZSELF' in Variable Param": Exit Sub
  With GrdBag
    If .Value(RowNum, "BOpnDt") = "" Then Cancel = True: ErrMsg = "Date Cannot Be Blank": Exit Sub
    If ms_LockDt <> "01/01/80" And CDate(ms_LockDt) > CDate(.Value(RowNum, "BOpnDt")) And .IsNew(RowNum) = True Then _
      Cancel = True: ErrMsg = "Opening Date Cannot Be Less Than Lock Date.": Exit Sub
    If ms_PurgeDt <> "01/01/80" And CDate(ms_PurgeDt) >= CDate(.Value(RowNum, "BOpnDt")) And .IsNew(RowNum) = True Then _
      Cancel = True: ErrMsg = "Opening Date Cannot Be Less Than Or Equal To Purge Date.": Exit Sub
  End With
End Sub
Private Sub GrdBag_InitKey(ByVal RowNum As Integer)
  '*** Initialise Key Fields
  '*** If New row (i.e. BLstYyKey= '') then BQty & BPtQty are equal to BOpnQty & BOpnPtQty
      'respectively
  
  With GrdBag
    .Store "BCoCd", ADC("wBCoCd")
    .Store "BOdTc", ADC("wBOdTc")
    .Store "BOdYy", ADC("wBOdYy")
    .Store "BOdChr", ADC("wBOdChr")
    .Store "BOdNo", ADC("wBOdNo")
    
    If GrdBag.IsNew(RowNum) Then
      .Store "BGrWt", 0
      .Store "BQty", .Value(RowNum, "BOpnQty")
      .Store "BWrk", ctTdWrk
      .Store "BPtQty", 0
      .Store "BRjQty", 0
      .Store "BRjPtQty", 0
      .Store "BLstYy", ""
      .Store "BLstKey", 0
      .Store "BLstYyKey", ""
      .Store "BPrdStg", ""
      .Store "BFgSubLoc", ""
      .Store "BRecvDt", .Value(RowNum, "BOpnDt")
      .Store "BClsDt", MWLib.EmptyDate
      .Store "BCnvJtYn", "N"      ' ***** Manali Trading Module
      
      '****** Sachin 3.02 20-12-07 - Setting the values for Id No's and New fields
      .Store "BOdIdNo", moCn.GetFldVal("Select OdIdNo From OrdDsg " + _
                                       " Where OdCoCd='" + ADC("wBCoCd") + "' And OdTc='" + ADC("wBOdTc") + "' " + _
                                       "  And OdYy='" + ADC("wBOdYy") + "' And OdChr='" + ADC("wBOdChr") + "' And OdNo=" + CStr(ADC("wBOdNo")) + " And OdSr= " + CStr(.Value(RowNum, "BOdSr")) + " and OdPrtKey='" + ctCurrPrtn + "'")
      
      .Store "BOmIdNo", moCn.GetFldVal("Select OmIdNo From OrdMst " + _
                                          "Where OmCoCd='" + ADC("wBCoCd") + "' And OmTc='" + ADC("wBOdTc") + "' " + _
                                          "  And OmYy='" + ADC("wBOdYy") + "' And OmChr='" + ADC("wBOdChr") + "' And OmNo=" + CStr(ADC("wBOdNo")) + " and OmPrtKey='" + ctCurrPrtn + "'")
      
                                          
      .Store "BMaxTxnDt", .Value(RowNum, "BOpnDt")
      '****** Sachin 3.02 20-12-07 - Setting the values for Id No's and New fields
      
    End If
      'indigo.164 BOdDmTcTyp added
      .Store "BOdDmTcTyp", ms_wBOdDmTcTyp
      
      'indigo.164 DmTcTyp='DM'  removed
      .Store "BDmIdNo", moCn.GetFldVal("Select DmIdNo From DsgMst " + _
                                          " Where DmTcTyp = '" + ms_wBOdDmTcTyp + "'" + " and DmCd='" + .Value(RowNum, "BOdDmCd") + "' " + _
                                          "  And DmSz='' " + IIF(.Value(RowNum, "BPrtKey") = ctCurrPrtn, " And DmPrtKey='" + ctCurrPrtn + "'", ""))
   
    '****** Sachin 3.02 - Current Partition
    If moCn.GetFldVal("Select BLstYyKey From Bag Where BCoCd= '" + ADC("wBCoCd") + "' " + _
                      "and BYy= '" + .Value(RowNum, "BYy") + "' " + _
                      "and BChr= '" + .Value(RowNum, "BChr") + "' and " + _
                      "BNo= " + CStr(.Value(RowNum, "BNo")) + IIF(.Value(RowNum, "BPrtKey") = ctCurrPrtn, " and BPrtKey='" + ctCurrPrtn + "'", "")) = "" Then
      .Store "BQty", .Value(RowNum, "BOpnQty")
      .Store "BPtQty", .Value(RowNum, "BOpnPtQty")
    End If
  
  End With
End Sub
Private Sub ADC_SetRecSource()
  '*** Since this is an 'ABC' Entry, there is no RecSource associated with ADC. So set the
      'DataMode property of ADC control as xdatUnbound

  ADC.DataMode = xdatUnbound
End Sub
Private Sub GrdBag_SetRecSource()
  '*** Set the RecSource for GrdBag (Table = Bag) based on the scope entered by the user
  
  Dim wCnd As String, wOrdBy As String
  wCnd = ADC.RepCond
  wCnd = IIF(wCnd <> "", " Where " + wCnd, "")
  
  ' **** Zubin 211 **** '
  If ADC("wBagTyp") <> "" Then _
    wCnd = wCnd + IIF(wCnd <> "", " and Exists (Select 'x' from Param where PTyp= 'BCHR' and PMCd= BChr and " + _
                                              "PSCd= '' and PValue= '" + ADC("wBagTyp") + "') ", " ")
  ' **** Zubin 211 **** '
  
  wOrdBy = moCn.GetFldVal("Select PDesc225 from Param " + _
           "where PTyp= 'SRTBAG' and PMCd= '" + ADC("wBSrt") + "'")
  wOrdBy = IIF(wOrdBy = "", " Order By BYy, BChr, BNo ", "Order By " + wOrdBy)
  
  '****** Sachin 3.02 10-12-07 - Corresponding Partitions
  GrdBag.RecSource = " Select * from Bag " + wCnd + wOrdBy
  
End Sub

Private Sub CmdBagGen_Click()
  
  If ms_OrdLck = True Then DispMsg "Cannot Generate Bags In Previous Partition", etError: Exit Sub
  
  '*** Display or hide the Bag Generation frame depending on the toggle state
  Call DispFra(BagGen)
  If FraBagGen.Visible = True Then
    '*** (Bef 2.11) If ADC("wBgDefBChr") = "" Then ADC("wBgDefBChr") = moCn.GetFldVal("Select hBChr from Head where HCoCd= '" + ADC("wBCoCd") + "' and HCd='" + ctSelfCmCd + "'")
    
    '*** 2.11 (Jen)
    ADC("wBgDefBChr").InitialiseFld ""
    ADC("wBgDefBChr") = moCn.GetFldVal("Select hBChr from Head where HCoCd= '" + ADC("wBCoCd") + "' and HCd='" + ctSelfCmCd + "'")
    '*** 2.11 (Jen)
  End If
End Sub
Private Sub CmdBagGenGo_Click()
  '*** Call the BagGeneration procedure to automatically generate the bags for the
      'Order Sr range specified by the user
  '*** Close the Bag Generation Frame
  
  Call BagGeneration
  Call CmdBagGen_Click
End Sub
Private Sub BagGeneration()
  '*** Bag Generation Option is used to automatically generate the bags for the
      'Order Sr range specified by the user
      'If no range is specified by the user then Bags are generated for the entire order
      'Also bags can be generated for an Order Sr No if there are no bags already existing
      'for the Order Sr (The bags will be generated for the other Order Sr Numbers)
  '*** Check if 'ZSELF' worker exists in Param. If not then copy is not possible
  '*** Bags are generated only if the Order is not on Hold
  
  Dim wBYy As String, wBChr As String, wCnd As String, wLocSeq As String, wOpnLoc As String
  Dim wPosComma As Integer, wOdBagQty As Single, wSrOnHld As String, wSrExistingBag As String
  Dim wRow As Single, wBagExists As Boolean, wDispStr As String
  '*** jay 2.13(CT) ***
  Dim wOctParts As Single     ', wOdBagPtQty As Single, wPrdParts As Single
  '*** jay 2.13(CT) ***
  
  If ADC("wBgOdSrFr") > ADC("wBgOdSrTo") Then DispMsg "From OrdSr Should Be <= To OrdSr", etError: Exit Sub
  If moCn.RecSeek("Select 'x' From SysDb where SysId='UNI' and (SdCoCd='' or charindex(','+'" + ADC("wBOdChr") + "'+',',','+SdCoCd+',')>0)") Then
      If (ADC("WBGODLN") = 0 Or ADC("WBGODLN") = "") Then
        DispMsg "Order Line cannot be blank when Ppc system is installed", etError
        Exit Sub
      End If
  End If

  If Not moCn.RecSeek("Select vPMCd From vParam Where vPCoCd= '" + ADC("wBCoCd") + "' and " + _
     "vPTyp= 'WORK' And vPMCd= '" + ctTdWrk + "'") Then _
     DispMsg "Enter Worker 'ZSELF' in Variable Param", etError: Exit Sub
  
  'Geeta ***Crm***2.12*********
  '****** Sachin 3.02 10-12-07 - current Partition
  If moCn.RecSeek("Select 'x' From OrdMst where OmCoCd= '" + ADC("wBCoCd") + "' and " + _
                            "OmTc= '" + ADC("wBOdTc") + "' and OmYy= '" + ADC("wBOdYy") + "' and " + _
                            "OmChr= '" + ADC("wBOdChr") + "' and " + _
                            "OmNo= " + CStr(ADC("wBOdNo")) + " And OmAckYn = 'N' And OmPrtKey='" + ctCurrPrtn + "'") Then _
    DispMsg "Cannot Generate Bags As Order Is Not Acknowledged", etError: Exit Sub
  
  wBYy = moCn.GetFldVal("Select hYy from Head where HCoCd= '" + ADC("wBCoCd") + "' and HCd= '" + ctSelfCmCd + "'")
    
  '****************************************************************************************
  'wBChr = moCn.GetFldVal("Select hBChr from Head where HCoCd= '" + ADC("wBCoCd") + "' and HCd= '" + ctSelfCmCd + "'")
  wBChr = ADC("wBgDefBChr")
  '******************************************************************************************
  
  wCnd = IIF(ADC("wBgOdSrFr") = 0, "", " and OdSr>= " + CStr(ADC("wBgOdSrFr")) + " ")
  wCnd = wCnd + IIF(ADC("wBgOdSrTo") = 0, "", " and OdSr<= " + CStr(ADC("wBgOdSrTo")) + " ")
  wBagExists = False
  '*** Beginning the copy process
  GrdBag.StartCopy
  MWLib.BeginProcess Me, "Copying Bag Details ..."
  Dim wRsBag As MwfLib.MDORowSet
  '****** Sachin 3.02 10-12-07 - current Partition
  Set wRsBag = moCn.OpenRes("Select OdIdNo, OdOmIdNo, OdDmIdNo, OdSr, OdDmCd, OdSfx, OdDmSz, " + _
                            "OdPrdSeq, OdPrdQty, OdHld, " + _
                            "OdBagPcs, OdOmCtCd, OdPrtExpQty from OrdDsg where OdCoCd= '" + ADC("wBCoCd") + "' and " + _
                            "OdTc= '" + ADC("wBOdTc") + "' and OdYy= '" + ADC("wBOdYy") + "' and " + _
                            "OdChr= '" + ADC("wBOdChr") + "' and " + _
                            "OdNo= " + CStr(ADC("wBOdNo")) + " and OdPrtKey='" + ctCurrPrtn + "' " + wCnd + " Order By OdSr")
                            
  
  With wRsBag
    wLocSeq = "": wOpnLoc = "": wPosComma = 0: wOdBagQty = 0
    Do While Not (.EOF Or .BOF)
       If !OdHld = "Y" Then _
          wSrOnHld = wSrOnHld + IIF(wSrOnHld = "", "", ", ") + CStr(!OdSr)
      ' **** Zubin 211 **** '
      Dim ws_CommStr As String, ws_BagExistsCnd As String, ws_BagType As String
      '''''      wBagExists = moCn.RecSeek("Select 'x' from Bag where " + _
                                "BCoCd= '" + ADC("wBCoCd") + "' and " + _
                                "BOdTc= '" + ADC("wBOdTc") + "' and " + _
                                "BOdYy= '" + ADC("wBOdYy") + "' and " + _
                                "BOdChr= '" + ADC("wBOdChr") + "' and " + _
                                "BOdNo= " + CStr(ADC("wBOdNo")) + " and " + _
                                "BOdSr= " + CStr(!OdSr))
                                
      '****** Sachin 3.02 10-12-07 - current Partitions
      ws_CommStr = "Select 'x' from Bag where " + _
                                "BCoCd= '" + ADC("wBCoCd") + "' and " + _
                                "BOdTc= '" + ADC("wBOdTc") + "' and " + _
                                "BOdYy= '" + ADC("wBOdYy") + "' and " + _
                                "BOdChr= '" + ADC("wBOdChr") + "' and " + _
                                "BOdNo= " + CStr(ADC("wBOdNo")) + " and " + _
                                "BOdSr= " + CStr(!OdSr) + " and BPrtKey='" + ctCurrPrtn + "'"
      
      ws_BagType = moCn.GetFldVal("Select PValue from Param where PTyp= 'BCHR' and PMCd='" + ADC("wBgDefBChr") + "' and PSCd= '' ")
      
      If UCase(ws_BagType) = "P" Then
        ws_BagExistsCnd = " and Exists (Select 'x' from Param where PTyp= 'BCHR' and PMCd= BChr and PSCd= '' and " + _
                                "PValue= 'P') " + _
                                "and (Select IsNull(Sum(BQty), 0) from Bag " + _
                                "Join Param On  PTyp= 'BCHR' and PMCd= BChr and PSCd= '' and PValue= 'P'  where " + _
                                "BCoCd= '" + ADC("wBCoCd") + "' and BOdTc= '" + ADC("wBOdTc") + "' and " + _
                                "BOdYy= '" + ADC("wBOdYy") + "' and BOdChr= '" + ADC("wBOdChr") + "' and " + _
                                "BOdNo= " + CStr(ADC("wBOdNo")) + " and BOdSr= " + CStr(!OdSr) + " and BPrtKey='" + ctCurrPrtn + "') >= " + _
                                " (select OdPrdQty from OrdDsg where OdCoCd= '" + ADC("wBCoCd") + "' and " + _
                                "       OdTc = '" + ADC("wBOdTc") + "' and OdYy= '" + ADC("wBOdYy") + "' and " + _
                                "       OdChr= '" + ADC("wBOdChr") + "' and OdNo= " + CStr(ADC("wBOdNo")) + " and " + _
                                "       OdSr =  " + CStr(!OdSr) + " and OdPrtKey='" + ctCurrPrtn + "' ) "
                                
                                
      ElseIf UCase(ws_BagType) = "F" Then
        '****** Sachin 3.02 10-12-07 - current partitions
        ws_BagExistsCnd = " and Exists (Select 'x' from Param where PTyp= 'BCHR' and PMCd= BChr and PSCd= '' and " + _
                                "PValue= 'F') and BChr= '" + ADC("wBgDefBChr") + "' and " + _
                                "(Select IsNull(Sum(BQty), 0) from Bag where " + _
                                "BCoCd= '" + ADC("wBCoCd") + "' and BOdTc= '" + ADC("wBOdTc") + "' and " + _
                                "BOdYy= '" + ADC("wBOdYy") + "' and BOdChr= '" + ADC("wBOdChr") + "' and " + _
                                "BOdNo= " + CStr(ADC("wBOdNo")) + " and BOdSr= " + CStr(!OdSr) + " and BPrtKey='" + ctCurrPrtn + "'  and  " + _
                                "BChr= '" + ADC("wBgDefBChr") + "') >= " + _
                                "     (select OdPrdQty from OrdDsg where OdCoCd= '" + ADC("wBCoCd") + "' and " + _
                                "       OdTc = '" + ADC("wBOdTc") + "' and OdYy= '" + ADC("wBOdYy") + "' and " + _
                                "       OdChr= '" + ADC("wBOdChr") + "' and OdNo= " + CStr(ADC("wBOdNo")) + " and " + _
                                "       OdSr =  " + CStr(!OdSr) + " and OdPrtKey='" + ctCurrPrtn + "' ) "
                                
      ElseIf UCase(ws_BagType) = "C" Then
        '****** Sachin 3.02 10-12-07 - current Partitions
        If Not moCn.RecSeek("Select OctCtChr from OrdCT where " + _
                        "OctCoCd= '" + ADC("wBCoCd") + "' and OctTc= '" + ADC("wBOdTc") + "' and " + _
                        "OctYy= '" + ADC("wBOdYy") + "' and OctChr= '" + ADC("wBOdChr") + "' and " + _
                        "OctNo= " + CStr(ADC("wBOdNo")) + " and OctSr= " + CStr(!OdSr) + " and " + _
                        "OctCtChr ='" + ADC("wBgDefBChr") + "' and OctPrtKey='" + ctCurrPrtn + "'") Then _
            DispMsg "Component Bag Does Not Exists In OrdCT Table for " + CStr(!OdSr), etWarning: GoTo NextOdSr
            
        '****** Sachin 3.02 10-12-07 - current Partitions
        ws_BagExistsCnd = " and Exists (Select 'x' from Param where PTyp= 'BCHR' and PMCd= BChr and PSCd= '' and " + _
                                "PValue= 'C') and BChr= '" + ADC("wBgDefBChr") + "' and " + _
                                "(Select IsNull(Sum(BQty), 0) from Bag where " + _
                                "BCoCd= '" + ADC("wBCoCd") + "' and BOdTc= '" + ADC("wBOdTc") + "' and " + _
                                "BOdYy= '" + ADC("wBOdYy") + "' and BOdChr= '" + ADC("wBOdChr") + "' and " + _
                                "BOdNo= " + CStr(ADC("wBOdNo")) + " and BOdSr= " + CStr(!OdSr) + " and BPrtKey='" + ctCurrPrtn + "' and " + _
                                "BChr= '" + ADC("wBgDefBChr") + "') >= " + _
                                "     (select OdPrdQty from OrdDsg where OdCoCd= '" + ADC("wBCoCd") + "' and " + _
                                "        OdTc= '" + ADC("wBOdTc") + "' and OdYy= '" + ADC("wBOdYy") + "' and " + _
                                "        OdChr= '" + ADC("wBOdChr") + "' and OdNo= " + CStr(ADC("wBOdNo")) + " and " + _
                                "        OdSr = " + CStr(!OdSr) + " and OdPrtKey='" + ctCurrPrtn + "' ) "
                                      
      End If
      wBagExists = moCn.RecSeek(ws_CommStr + ws_BagExistsCnd)
      ' **** Zubin 211 **** '
      
      If wBagExists Then _
         wSrExistingBag = wSrExistingBag + IIF(wSrExistingBag = "", "", ", ") + CStr(!OdSr)
      
      If wBagExists = False Then    ' !OdHld = "N" And     'geeta***
        '*** Jay 2.13(CT) ***
        '*** Get The Location Seqence ***
        If ws_BagType <> "C" Then
          wLocSeq = moCn.GetFldVal("Select vPDesc225 from vParam where vPCoCd= '" + ADC("wBCoCd") + "' " + _
                                "and vPTyp= 'PRDSEQ' and vPMCd= '" + !OdPrdSeq + "'")
        ElseIf ws_BagType = "C" Then
          '*** Jay 2.14Next *** [Multiple Comp in One BCHR]
          '****** Sachin 3.02 10-12-07 - Corresponding Partitions
          wLocSeq = moCn.GetFldVal("Select vPDesc225 from vParam where vPCoCd= '" + ADC("wBCoCd") + "' " + _
                            "and vPTyp= 'PRDSEQ' and vPMCd= (Select Max(OctPrdSeq) as OctPrdSeq from OrdCT where " + _
                            "OctCoCd= '" + ADC("wBCoCd") + "' and " + _
                            "OctTc= '" + ADC("wBOdTc") + "' and OctYy= '" + ADC("wBOdYy") + "' and " + _
                            "OctChr= '" + ADC("wBOdChr") + "' and OctNo= " + CStr(ADC("wBOdNo")) + " and " + _
                            "OctSr= " + CStr(!OdSr) + " and OctCtChr= '" + ADC("wBgDefBChr") + "' and OctPrtKey='" + ctCurrPrtn + "' )")
          '*** Jay 2.14Next *** [Multiple Comp in One BCHR]
        End If
        '*** Get The First Location In The Seqence ***
        wPosComma = InStr(1, wLocSeq, ",")
        If wPosComma > 0 Then
          wOpnLoc = Mid(wLocSeq, 1, wPosComma - 1)
        Else
          wOpnLoc = wLocSeq
        End If
        
        'Sachin 3.11.0
        If ADC("WBGDEFBOPNLOC") <> "" Then wOpnLoc = ADC("WBGDEFBOPNLOC")
                
        
        ' **** Zubin 211 **** '
        '''''        wOdBagQty = 0
        If UCase(ws_BagType) = "P" Then
         ' wOdBagQty = 0
          wOdBagQty = moCn.GetFldVal("Select IsNull(Sum(BQty), 0) from Bag " + _
                                    "Join Param On  PTyp= 'BCHR' and PMCd= BChr and PSCd= '' and PValue= 'P'  where " + _
                                    "BCoCd= '" + ADC("wBCoCd") + "' and BOdTc= '" + ADC("wBOdTc") + "' and " + _
                                    "BOdYy= '" + ADC("wBOdYy") + "' and BOdChr= '" + ADC("wBOdChr") + "' and " + _
                                    "BOdNo= " + CStr(ADC("wBOdNo")) + " and BOdSr= " + CStr(!OdSr) + " and BPrtKey='" + ctCurrPrtn + "'")
        ElseIf UCase(ws_BagType) = "F" Then
          '****** Sachin 3.02 10-12-07 - current Partition
          wOdBagQty = moCn.GetFldVal("Select IsNull(Sum(BQty), 0) from Bag where " + _
                                    "BCoCd= '" + ADC("wBCoCd") + "' and BOdTc= '" + ADC("wBOdTc") + "' and " + _
                                    "BOdYy= '" + ADC("wBOdYy") + "' and BOdChr= '" + ADC("wBOdChr") + "' and " + _
                                    "BOdNo= " + CStr(ADC("wBOdNo")) + " and BOdSr= " + CStr(!OdSr) + " and BPrtKey='" + ctCurrPrtn + "' and " + _
                                    "Exists (Select 'x' from Param where PTyp= 'BCHR' and PMCd= BChr and PSCd= '' and " + _
                                    "   PValue= 'F') and BChr= '" + ADC("wBgDefBChr") + "'")
                                    
        ElseIf UCase(ws_BagType) = "C" Then
          '****** Sachin 3.02 10-12-07 - current Partition
          wOdBagQty = moCn.GetFldVal("Select IsNull(Sum(BQty), 0) from Bag where " + _
                                    "BCoCd= '" + ADC("wBCoCd") + "' and BOdTc= '" + ADC("wBOdTc") + "' and " + _
                                    "BOdYy= '" + ADC("wBOdYy") + "' and BOdChr= '" + ADC("wBOdChr") + "' and " + _
                                    "BOdNo= " + CStr(ADC("wBOdNo")) + " and BOdSr= " + CStr(!OdSr) + " and BPrtKey='" + ctCurrPrtn + "' and " + _
                                    "Exists (Select 'x' from Param where PTyp= 'BCHR' and PMCd= BChr and PSCd= '' and " + _
                                    "   PValue= 'C') and BChr= '" + ADC("wBgDefBChr") + "'")
                                    
          'wOdBagPtQty = moCn.GetFldVal("Select IsNull(Sum(BPtQty), 0) from Bag where " + _
                                    "BCoCd= '" + ADC("wBCoCd") + "' and BOdTc= '" + ADC("wBOdTc") + "' and " + _
                                    "BOdYy= '" + ADC("wBOdYy") + "' and BOdChr= '" + ADC("wBOdChr") + "' and " + _
                                    "BOdNo= " + CStr(ADC("wBOdNo")) + " and BOdSr= " + CStr(!OdSr) + " and " + _
                                    "Exists (Select 'x' from Param where PTyp= 'BCHR' and PMCd= BChr and PSCd= '' and " + _
                                    "   PValue= 'C') and BChr= '" + ADC("wBgDefBChr") + "'")
          'Chk pending for OctCtChr in ordmst
          '*** Jay 2.14Next *** [Multiple Comp in One BCHR]
          '****** Sachin 3.02 10-12-07 - current Partitions
          wOctParts = moCn.GetFldVal("Select Sum(OctParts) as OctParts from OrdCT where " + _
                        "OctCoCd= '" + ADC("wBCoCd") + "' and OctTc= '" + ADC("wBOdTc") + "' and " + _
                        "OctYy= '" + ADC("wBOdYy") + "' and OctChr= '" + ADC("wBOdChr") + "' and " + _
                        "OctNo= " + CStr(ADC("wBOdNo")) + " and OctSr= " + CStr(!OdSr) + " and OctPrtKey='" + ctCurrPrtn + "' and " + _
                        "OctCtChr ='" + ADC("wBgDefBChr") + "'")
          '*** Jay 2.14Next *** [Multiple Comp in One BCHR]
          'wPrdParts = wOctParts * (!OdPrdQty)
          'ADC("wOdCTBalPtQty") = wi_OrdCTPrt * (ADC("wOdBQty") + ADC("wOdBalQty")) - wRsOrdDet!qCtBTotPtQty
        End If
' **** Zubin 211 **** '
        Do While (wOdBagQty < (!OdPrdQty - !OdPrtExpQty))
          MWLib.RunningProcess "Copying Details For Order Sr '" + CStr(!OdSr) + "'"
          GrdBag.AddItem
          wRow = GrdBag.Rows - 1
          GrdBag.Value(wRow, "BPrtKey") = ctCurrPrtn      '****** Sachin 3.02 20-12-07 - Partition Key
          GrdBag.Value(wRow, "BYy") = wBYy
          GrdBag.Value(wRow, "BChr") = wBChr
          GrdBag.Value(wRow, "BOdSr") = !OdSr
          GrdBag.Value(wRow, "BOdDmCd") = !OdDmCd
          GrdBag.Value(wRow, "BOdSfx") = !OdSfx
          GrdBag.Value(wRow, "BOdDmSz") = !OdDmSz
          
          '****** Sachin 3.02 - Id fields ******
          'GrdBag.Value(wRow, "BOdIdNo") = !OdIdNo
          'GrdBag.Value(wRow, "BDmIdNo") = !OdDmIdNo
          'GrdBag.Value(wRow, "BOmIdNo") = !OdOmIdNo
          '****** Sachin 3.02 - Id fields ******
          
          GrdBag.Value(wRow, "BOpnLoc") = wOpnLoc
          GrdBag.Value(wRow, "BLoc") = wOpnLoc
          GrdBag.Value(wRow, "BOpnDt") = moCn.SrvrDate     'Date
          '''(Old COde) GrdBag.Value(wRow, "BNo") = moCn.AutoGen("Bag", "BNo", wBYy + wBChr)
          GrdBag.Value(wRow, "BCls") = "N"
          GrdBag.Value(wRow, "BOdCtCd") = !OdOmCtCd
      'EmrSv
          'GrdBag.Value(wRow, "BOdLnIdKey") = moCn.GetFldVal("Select IdKey from OrdLn where OdlCoCd ='" + adc("wBCoCd") + "' and " + _
                " OdlTc ='" + adc("wBOdTc") + "' and OdlYy= '" + adc("wBOdYy") + "' and " + _
                " OdlChr='" + adc("wBOdChr") + "' and OdlNo= " + CStr(adc("wBOdNo")) + " and " + _
                " OdlSr= " + CStr(GrdBag.Value(wRow, "BOdSr")) + " and OdlSrNo=1")
          GrdBag.Value(wRow, "BOdLSrNo") = CStr(ADC("WBGODLN"))
                                
    '******************************************************************************************
          If ADC("wBgDefBPcs") > 0 Then
            '****** Sachin 3.02 20-12-07 - OdPrtExpQty included below
            If wOdBagQty + ADC("wBgDefBPcs") > (!OdPrdQty - !OdPrtExpQty) Then
              GrdBag.Value(wRow, "BOpnQty") = (!OdPrdQty - !OdPrtExpQty) - wOdBagQty
              '*** Jay 2.13(CT)***
              GrdBag.Value(wRow, "BOpnPtQty") = ((!OdPrdQty - !OdPrtExpQty) - wOdBagQty) * wOctParts
              '*** Jay 2.13(CT)***
              wOdBagQty = (!OdPrdQty - !OdPrtExpQty)
            Else
              GrdBag.Value(wRow, "BOpnQty") = ADC("wBgDefBPcs")
              '*** Jay 2.13(CT)***
              GrdBag.Value(wRow, "BOpnPtQty") = ADC("wBgDefBPcs") * wOctParts
              '*** Jay 2.13(CT)***
              wOdBagQty = wOdBagQty + ADC("wBgDefBPcs")
            End If
          Else
            If wOdBagQty + !OdBagPcs > (!OdPrdQty - !OdPrtExpQty) Then
              GrdBag.Value(wRow, "BOpnQty") = (!OdPrdQty - !OdPrtExpQty) - wOdBagQty
              GrdBag.Value(wRow, "BOpnPtQty") = ((!OdPrdQty - !OdPrtExpQty) - wOdBagQty) * wOctParts
              '*** Jay 2.13(CT)***
'Sample Template to Check Part Qty with PrtPrdQty
'              If ws_BagType = "C" Then
'                If (!OdPrdQty - wOdBagQty) * wOctParts > wPrdParts Then
'                  GrdBag.Value(wRow, "BOpnPtQty") = wPrdParts - wOdBagPtQty
'                  wOdBagPtQty = wPrdParts
'                Else
'                  GrdBag.Value(wRow, "BOpnPtQty") = (!OdPrdQty - wOdBagQty) * wOctParts
'                  wOdBagPtQty = wOdBagPtQty + ((!OdPrdQty - wOdBagQty) * wOctParts)
'                End If
'              End If
              '*** Jay 2.13(CT)***
              wOdBagQty = (!OdPrdQty - !OdPrtExpQty)
            Else
              GrdBag.Value(wRow, "BOpnQty") = !OdBagPcs
              '*** Jay 2.13(CT)***
              If ws_BagType = "C" Then
                  GrdBag.Value(wRow, "BOpnPtQty") = !OdBagPcs * wOctParts
'Sample Template to Check Part Qty with PrtPrdQty
'                If !OdBagPcs * wOctParts > wPrdParts Then
'                  GrdBag.Value(wRow, "BOpnPtQty") = wPrdParts - wOdBagPtQty
'                  wOdBagPtQty = wPrdParts
'                Else
'                  GrdBag.Value(wRow, "BOpnPtQty") = !OdBagPcs * wOctParts
'                  wOdBagPtQty = wOdBagPtQty + (!OdBagPcs * wOctParts)
'                End If
              End If
              '*** Jay 2.13(CT)***
              wOdBagQty = wOdBagQty + !OdBagPcs
            End If
          End If
'******************************************************************************************
          
          GrdBag.SaveRec (wRow)
        Loop
      End If
NextOdSr:
      .MoveNext
    Loop
  End With

  MWLib.EndProcess Me
  GrdBag.EndCopy

  '*** Bag Generation is not possible for the following Sr Nos. as bag already exists for the Sr Nos.
  If wSrOnHld <> "" Or wSrExistingBag <> "" Then
     wDispStr = IIF(wSrOnHld = "", "", "The following Sr Nos. " + wSrOnHld + " Are On Hold. ")
     wDispStr = wDispStr + IIF(wSrExistingBag = "", "", "Bag Exists For The Sr Nos. " + wSrExistingBag)
  End If
  
  If wDispStr <> "" Then DispMsg wDispStr, etError: Exit Sub
  Set wRsBag = Nothing
  '*** End of the copy process
  
End Sub

Private Sub SetBagDet()
Dim wi_RowNum As Integer, wBagOpnQty As Single, wBTyp As String, wi_OrdCTPrt As String
Dim wPrtnFGQty As Single, wPrtnExpQty As Single, wOdSrChk As Boolean

With GrdBag
wi_RowNum = .Row
If wi_RowNum <= 0 Or wi_RowNum >= GrdBag.Rows Then Exit Sub
wOdSrChk = moCn.RecSeek("Select 'x' From OrdDsg " + _
                                     " Join Bag On BOdIdNo=OdIdNo and BCoCd=OdCoCd and BOdTc=OdTc and BOdYy=OdYy and BOdChr=OdChr and BOdNo=OdNo and BOdSr=OdSr and BPrtKey<>OdPrtKey " + _
                                     " Where " + _
                                     "OdCoCd= '" + ADC("wBCoCd") + "' and " + _
                                     "OdTc= '" + ADC("wBOdTc") + "' and " + _
                                     "OdYy= '" + ADC("wBOdYy") + "' and " + _
                                     "OdChr= '" + ADC("wBOdChr") + "' and " + _
                                     "OdNo= " + CStr(ADC("wBOdNo")) + " and " + _
                                     "OdSr= " + CStr(.Value(wi_RowNum, "BOdSr")))

'*** Getting Current Bag Qty, Part Qty From Bag For The Order Sr.***
    Dim wRsOrdDet As MwfLib.MDORowSet
     ' **** Zubin 211 (Change in Query. select recs of only Primary bags) ***** '
'''''    Set wRsOrdDet = moCn.OpenRes("Select IsNull(sum(BQty), 0) as qTotQty, IsNull(sum(BPtQty),0) as qTotPtQty " + _
                                  "from Bag where BCoCd='" + ADC("wBCoCd") + "' and " + _
                                  "BOdTc= '" + ADC("wBOdTc") + "' and " + _
                                  "BOdYy= '" + ADC("wBOdYy") + "' and " + _
                                  "BOdChr= '" + ADC("wBOdChr") + "' and " + _
                                  "BOdNo= " + CStr(ADC("wBOdNo")) + " and " + _
                                  "BOdSr= " + CStr(.Value(wBagRow, "BOdSr")))
'*** jay 2.13(CT) ***
    ADC("wFbBChr") = "": ADC("wOdFbBalQty") = "": ADC("wOdFbPtQty") = ""
    ADC("wCtBChr") = "": ADC("wOdCTBalQty") = "": ADC("wOdCTPtQty") = "": ADC("wOdCTBalPtQty") = ""
    
    wPrtnExpQty = moCn.GetFldVal("Select IsNull(OdPrtExpQty,0) From OrdDsg Where " + _
                                     "(Select PValue From Param where PTyp='BCHR' and PMCd='" + GrdBag.Value(wi_RowNum, "BChr") + "'" + _
                                     " and PSCd='')='P' and " + _
                                     "OdCoCd= '" + ADC("wBCoCd") + "' and " + _
                                     "OdTc= '" + ADC("wBOdTc") + "' and " + _
                                     "OdYy= '" + ADC("wBOdYy") + "' and " + _
                                     "OdChr= '" + ADC("wBOdChr") + "' and " + _
                                     "OdNo= " + CStr(ADC("wBOdNo")) + " and " + _
                                     "OdSr= " + CStr(.Value(wi_RowNum, "BOdSr")) + IIF(wOdSrChk = False And .Value(wi_RowNum, "BPrtKey") = ctCurrPrtn, " and OdPrtKey='" + ctCurrPrtn + "'", ""))
    
    '****** Sachin 3.02 10-12-07 - Current Partitions
    Set wRsOrdDet = moCn.OpenRes("Select IsNull(sum(case PValue when 'P' then BQty else 0 end), 0) as qPBTotQty, " + _
                    "IsNull(sum(case PValue when 'P' then BPtQty else 0 end), 0) as qTotPtQty, " + _
                    "IsNull(sum(case when PValue ='C' and BChr= '" + GrdBag.Value(wi_RowNum, "BChr") + "' then BPtQty else 0 end), 0) as qCtBTotPtQty, " + _
                    "IsNull(sum(case when PValue ='F' and BChr= '" + GrdBag.Value(wi_RowNum, "BChr") + "' then BPtQty else 0 end), 0) as qFbBTotPtQty, " + _
                    "IsNull(sum(case when PValue= 'F' and BChr= '" + GrdBag.Value(wi_RowNum, "BChr") + "' then BQty else 0 end), 0) as qFBTotQty, " + _
                    "IsNull(sum(case when PValue= 'C' and BChr= '" + GrdBag.Value(wi_RowNum, "BChr") + "' then BQty else 0 end), 0) as qCTBTotQty " + _
                    "From Bag " + _
                    "Join Param On PTyp= 'BCHR' and PMCd= BChr and PSCd= '' " + _
                    "where BCoCd='" + ADC("wBCoCd") + "' and BOdTc= '" + ADC("wBOdTc") + "' and " + _
                    "BOdYy= '" + ADC("wBOdYy") + "' and BOdChr= '" + ADC("wBOdChr") + "' and " + _
                    "BOdNo= " + CStr(ADC("wBOdNo")) + " and " + _
                    "BOdSr= " + CStr(.Value(wi_RowNum, "BOdSr")) + IIF(wOdSrChk = False And .Value(wi_RowNum, "BPrtKey") = ctCurrPrtn, " and BPrtKey='" + ctCurrPrtn + "'", ""))
                    
    ' **** Zubin 211 **** '
    If Not (wRsOrdDet.EOF Or wRsOrdDet.BOF) Then
' **** Zubin 211 **** '
'''''      ADC("wOdBQty") = wRsOrdDet!qTotQty
      ADC("wOdBQty") = wRsOrdDet!qPBTotQty
      ADC("wOdTPtQty") = wRsOrdDet!qTotPtQty
      ADC("wOdFBBQty") = wRsOrdDet!qFbTotQty
      ADC("wOdCTBQty") = wRsOrdDet!qCTBTotQty
      ADC("wOdFbPtQty") = wRsOrdDet!qFbBTotPtQty
      'Geeta***2.11
      wBTyp = moCn.GetFldVal("Select Pvalue From Param where PTyp='BCHR' and PMCd='" + GrdBag.Value(wi_RowNum, "BChr") + "' " + _
                                       " and PSCd=''")
      If wBTyp = "F" Then
        '****** Sachin 3.02 10-12-07 - current Partitions
        ADC("wOdFbBalQty") = moCn.GetFldVal("Select OdPrdQty from OrdDsg where " + _
                                     "(Select PValue From Param where PTyp='BCHR' and PMCd='" + GrdBag.Value(wi_RowNum, "BChr") + "'" + _
                                     " and PSCd='')='F' and " + _
                                     "OdCoCd= '" + ADC("wBCoCd") + "' and " + _
                                     "OdTc= '" + ADC("wBOdTc") + "' and " + _
                                     "OdYy= '" + ADC("wBOdYy") + "' and " + _
                                     "OdChr= '" + ADC("wBOdChr") + "' and " + _
                                     "OdNo= " + CStr(ADC("wBOdNo")) + " and " + _
                                     "OdSr= " + CStr(.Value(wi_RowNum, "BOdSr")) + IIF(wOdSrChk = False And .Value(wi_RowNum, "BPrtKey") = ctCurrPrtn, " and OdPrtKey='" + ctCurrPrtn + "'", "")) - ADC("wOdFBBQty")
        ADC("wFbBChr") = GrdBag.Value(wi_RowNum, "BChr"): ADC("wCTBChr") = ""
        
      ElseIf wBTyp = "C" Then
        '****** Sachin 3.02 10-12-07 - current Partitions
        ADC("wOdCTBalQty") = moCn.GetFldVal("Select OdPrdQty from OrdDsg where " + _
                                     "(Select PValue From Param where PTyp='BCHR' and PMCd='" + GrdBag.Value(wi_RowNum, "BChr") + "'" + _
                                     " and PSCd='')= 'C' and " + _
                                     "OdCoCd= '" + ADC("wBCoCd") + "' and OdTc= '" + ADC("wBOdTc") + "' and " + _
                                     "OdYy= '" + ADC("wBOdYy") + "' and OdChr= '" + ADC("wBOdChr") + "' and " + _
                                     "OdNo= " + CStr(ADC("wBOdNo")) + " and " + _
                                     "OdSr= " + CStr(.Value(wi_RowNum, "BOdSr")) + IIF(wOdSrChk = False And .Value(wi_RowNum, "BPrtKey") = ctCurrPrtn, " and OdPrtKey='" + ctCurrPrtn + "'", "")) - ADC("wOdCTBQty")
        ADC("WODCTPTQTY") = wRsOrdDet!qCtBTotPtQty
        ADC("wCtBChr") = GrdBag.Value(wi_RowNum, "BChr")
        
      End If
    ' **** Zubin 211 **** '
    End If
    
        ADC("wOdBalQty") = moCn.GetFldVal("Select OdPrdQty from OrdDsg where " + _
                                         "OdCoCd= '" + ADC("wBCoCd") + "' and OdTc= '" + ADC("wBOdTc") + "' and " + _
                                         "OdYy= '" + ADC("wBOdYy") + "' and OdChr= '" + ADC("wBOdChr") + "' and " + _
                                         "OdNo= " + CStr(ADC("wBOdNo")) + " and " + _
                                         "OdSr= " + CStr(.Value(wi_RowNum, "BOdSr")) + IIF(wOdSrChk = False And .Value(wi_RowNum, "BPrtKey") = ctCurrPrtn, " and OdPrtKey='" + ctCurrPrtn + "'", "")) - (ADC("wOdBQty") + IIF(wOdSrChk = False And .Value(wi_RowNum, "BPrtKey") = ctCurrPrtn, wPrtnExpQty, 0))
      
    If wBTyp = "C" Then
        '****** Sachin 3.02 10-12-07 - current Partitions
        wi_OrdCTPrt = moCn.GetFldVal("Select Sum(OctParts) as OctParts from OrdCT where " + _
                        "OctCoCd= '" + ADC("wBCoCd") + "' and OctTc= '" + ADC("wBOdTc") + "' and " + _
                        "OctYy= '" + ADC("wBOdYy") + "' and OctChr= '" + ADC("wBOdChr") + "' and " + _
                        "OctNo= " + CStr(ADC("wBOdNo")) + " and OctSr= " + CStr(.Value(wi_RowNum, "BOdSr")) + IIF(wOdSrChk = False And .Value(wi_RowNum, "BPrtKey") = ctCurrPrtn, " and OctPrtKey='" + ctCurrPrtn + "'", "") + " and " + _
                        "OctCtChr ='" + .Value(wi_RowNum, "BCHR") + "' " + _
                        "Group By OctCoCd, OctTc, OctYy, OctChr, OctNo, OctSr, OctCtChr ")
        ADC("wOdCTBalPtQty") = wi_OrdCTPrt * (ADC("wOdBQty") + ADC("wOdBalQty")) - wRsOrdDet!qCtBTotPtQty
    End If
    
    If .IsNew(wi_RowNum) Then
      wBTyp = moCn.GetFldVal("Select PValue from Param where PTyp= 'BCHR' and " + _
                            "PMCd= '" + .Value(wi_RowNum, "BChr") + "' and PSCd= ''")
      If wBTyp = "P" Or wi_RowNum <= 1 Then
         '****** Sachin 3.02 10-12-07 - current Partitions
         wBagOpnQty = moCn.GetFldVal("Select OdBagPcs from OrdDsg where " + _
                       "OdCoCd='" + ADC("wBCoCd") + "' and OdTc= '" + ADC("wBOdTc") + "' and " + _
                       "OdYy= '" + ADC("wBOdYy") + "' and OdChr= '" + ADC("wBOdChr") + "' and " + _
                       "OdNo= " + CStr(ADC("wBOdNo")) + " and OdSr= " + CStr(.Value(wi_RowNum, "BOdSr")) + IIF(wOdSrChk = False And .Value(wi_RowNum, "BPrtKey") = ctCurrPrtn, " and OdPrtKey='" + ctCurrPrtn + "'", ""))
         wBagOpnQty = IIF(wBagOpnQty <= ADC("wOdBalQty"), wBagOpnQty, ADC("wOdBalQty"))
      Else
         wBagOpnQty = .Value(wi_RowNum - 1, "BOpnQty")
         If wBTyp = "F" Then
           wBagOpnQty = IIF(wBagOpnQty <= ADC("wOdFbBalQty"), wBagOpnQty, ADC("wOdFbBalQty"))
         ElseIf wBTyp = "C" Then
           wBagOpnQty = IIF(wBagOpnQty <= ADC("wOdCTBalQty"), wBagOpnQty, ADC("wOdCTBalQty"))
         End If
      End If
      
      'wBagOpnQty = IIF(wBagOpnQty <= IIF(wBTyp = "P", ADC("wOdBalQty"), ADC("wOdFbBalQty")), wBagOpnQty, IIF(wBTyp = "P", ADC("wOdBalQty"), ADC("wOdFbBalQty")))
      .Value(wi_RowNum, "BOpnQty") = wBagOpnQty
      '*** This is just a warning. The User can generate bags with 0 Qty.
      ' If wBagOpnQty = 0 Then DispMsg "Bag With Qty= 0 Has Been Generated", etWarning
     End If
'*** Jay 2.13(CT) ***
    Set wRsOrdDet = Nothing
   End With
End Sub

'*** Jay 2.13(CT) ***
Private Sub SetBagOpnLoc(ByVal ps_RowNum As Integer, ByVal ps_BTyp As String)
'Bag Open Loction is Same for Primary & flute Bags, So it is not complsory to pass ps_BTyp value for flute bags.
'By default it is considered as Prim.bag's open loc
Dim wPrdSeq As String, wLocSeq As String, wPosComma As Integer
  If ps_BTyp = UCase("C") Then
    '*** Get The Production Seqence ***
    '****** Sachin 3.02 10-12-07 - Corresponding Partitions
    wPrdSeq = moCn.GetFldVal("Select OctPrdSeq from OrdCT where " + _
                            "OctCoCd= '" + ADC("wBCoCd") + "' and " + _
                            "OctTc= '" + ADC("wBOdTc") + "' and OctYy= '" + ADC("wBOdYy") + "' and " + _
                            "OctChr= '" + ADC("wBOdChr") + "' and OctNo= " + CStr(ADC("wBOdNo")) + " and " + _
                            "OctSr= " + CStr(GrdBag.Value(ps_RowNum, "BOdSr")) + " And OctPrtKey='" + ctCurrPrtn + "' And " + _
                            "OctCtChr= '" + GrdBag.Value(ps_RowNum, "BChr") + "' ")
  Else
    '*** Get The Production Seqence ***
    '****** Sachin 3.02 10-12-07 - Corresponding Partitions
    wPrdSeq = moCn.GetFldVal("Select OdPrdSeq from OrdDsg where " + _
                            "OdCoCd= '" + ADC("wBCoCd") + "' and " + _
                            "OdTc= '" + ADC("wBOdTc") + "' and OdYy= '" + ADC("wBOdYy") + "' and " + _
                            "OdChr= '" + ADC("wBOdChr") + "' and OdNo= " + CStr(ADC("wBOdNo")) + " and " + _
                            "OdSr= " + CStr(GrdBag.Value(ps_RowNum, "BOdSr")) + " and OdPrtKey='" + ctCurrPrtn + "'")

  End If
  '*** Get The Location Seqence ***
  wLocSeq = moCn.GetFldVal("Select vPDesc225 from vParam where " + _
                            "vPCoCd = '" + ADC("wBCoCd") + "' and vPTyp= 'PRDSEQ' and " + _
                           "vPMCd= '" + wPrdSeq + "'")
  
  '*** Get The First Location In The Seqence ***
  wPosComma = InStr(1, wLocSeq, ",")
  Dim wBOpnLoc As String
  If wPosComma > 0 Then
    If GrdBag.Value(ps_RowNum, "BOpnLoc") <> Mid(wLocSeq, 1, wPosComma - 1) Then _
              wBOpnLoc = Mid(wLocSeq, 1, wPosComma - 1)
  Else
    If GrdBag.Value(ps_RowNum, "BOpnLoc") = wLocSeq Then _
               wBOpnLoc = wLocSeq
  End If
  If moCn.RecSeek("Select 'a' from Loc where LocCd = '" + wBOpnLoc + "' and LocCoCd = '" + ADC("wBCoCd") + "' and LocValidYN='Y'") Then
    GrdBag.Value(ps_RowNum, "BOpnLoc") = wBOpnLoc
  End If
End Sub
Private Sub DispFra(ByVal pv_BagFra As en_BagFra)
  '*** This is the common procedure which is used to toggle between hiding and displaying
      'the frames (i.e. if the frame is visible then make it invisible and viseversa)
      'associated with a command button

  Select Case pv_BagFra
  Case Is = BagGen
    '*** If Frame FraBagGen is visible then make it invisible and enable all the command buttons
    '*** If Frame FraBagGen is invisible then make it visible and disable all the
        'command buttons except the one which is clicked. Change the background colour of
        'the clicked command button to white and make the font bold
    If FraBagGen.Visible = True Then
      FraBagGen.Visible = False
      FraBagGen.Enabled = False
      Call EnaDisaCmds(False, CmdBagGen)
      CmdBagGen.SetFocus
    Else
      FraBagGen.Visible = True
      FraBagGen.Enabled = True
      ADC("wBgOdSrFr").SetFocus
      FraBagGen.ZOrder
      Call EnaDisaCmds(True, CmdBagGen)
    End If
  End Select
End Sub
Private Sub EnaDisaCmds(ByVal pv_ShowFra As Boolean, Optional ByVal pv_CmdBut As MWCTL_BTN1)
  '*** If second parameter is not given and pv_ShowFra = True then the frame FraNKeyAll
      'on Frame Fra_nkeys is disabled and allowadd and allowdelete options of adc are
      'set to false

  '*** If second parameter is not given and pv_ShowFra = False then the frame FraNKeyAll
      'on Frame Fra_nkeys is enabled and allowadd and allowdelete options of adc are
      'set to true

  '*** If second parameter is given and pv_ShowFra = True then the frame FraNKeyAll
      'on Frame Fra_nkeys is disabled and allowadd and allowdelete options of adc are
      'set to false and for the command button specified in the second parameter,
      'the BackColor is changed to white and the font is made Bold

  '*** If second parameter is given and pv_ShowFra = False then the frame FraNKeyAll
      'on Frame Fra_nkeys is enabled and allowadd and allowdelete options of adc are
      'set to true and for the command button specified in the second parameter,
      'the BackColor is changed to normal and the font is made Regular


  If Not pv_CmdBut Is Nothing Then
    If pv_ShowFra = True Then
      '*** Jenny (Old Code Bef Color)
'      pv_CmdBut.BackColor = vbWhite
'      pv_CmdBut.FontBold = True
      '*** Jenny (Old Code Bef Color)
      pv_CmdBut.CausesValidation = True
    Else
      '*** Jenny (Old Code Bef Color)
'      pv_CmdBut.BackColor = &H8000000F
'      pv_CmdBut.FontBold = False
      '*** Jenny (Old Code Bef Color)
      'pv_CmdBut.CausesValidation = False
    End If
  End If
  FraNKeyAll.Enabled = Not pv_ShowFra
  ADC.AllowSave = Not pv_ShowFra
  CmdBagGen.Enabled = Not pv_ShowFra
  If Not pv_CmdBut Is Nothing Then pv_CmdBut.Enabled = True
End Sub
Private Sub HideAllFras()
  '*** This procedure hides all the frames other than FraNKeyAll
  FraNKeyAll.Visible = True
  FraBagGen.Visible = False
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
  ws_BtnStr = "CmdBagGenGo"
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
Private Sub GrdBag_Validate(Cancel As Boolean)
  Cancel = GrdBag.Validate
End Sub





