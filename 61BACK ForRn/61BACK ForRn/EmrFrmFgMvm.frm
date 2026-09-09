VERSION 5.00
Object = "{AB2CC32A-E652-42A2-B341-26ECE02F6386}#2.1#0"; "mwfctlforspeed9.ocx"
Begin VB.Form EmrFrmFgMvm 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "FG Movement Entry"
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
   Begin VB.Frame FraCpyRfId 
      Height          =   3315
      Left            =   3000
      TabIndex        =   68
      Top             =   5760
      Visible         =   0   'False
      Width           =   9555
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
         Left            =   7140
         TabIndex        =   76
         ToolTipText     =   "Specify Copy From Bag Nos / Bag Alias No."
         Top             =   140
         Width           =   1695
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
         Index           =   0
         Left            =   5730
         TabIndex        =   75
         ToolTipText     =   "Specify Copy From Bag Nos / Bag Alias No."
         Top             =   140
         Width           =   1335
      End
      Begin VB.Frame Frame2 
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
         Height          =   2895
         Left            =   60
         TabIndex        =   73
         Top             =   405
         Width           =   8790
         Begin MwfCtl.MWCTL_MED ATXT 
            Height          =   2565
            Index           =   40
            Left            =   60
            TabIndex        =   74
            ToolTipText     =   "Copy Bag Nos / Bag Alias No."
            Top             =   220
            Width           =   8655
            _ExtentX        =   15266
            _ExtentY        =   4524
            Alignment       =   3
            MaxLength       =   0
            IdName          =   "WRFIDSEL"
         End
      End
      Begin MwfCtl.MWCTL_BTN1 CmdCpyRfIdGo 
         Height          =   525
         Left            =   8880
         TabIndex        =   70
         Top             =   2640
         Width           =   585
         _ExtentX        =   1032
         _ExtentY        =   926
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
         Left            =   8850
         TabIndex        =   69
         ToolTipText     =   "Enter 'Y' to Copy Thru Bag Nos and  'N' to Copy thru RfId"
         Top             =   120
         Visible         =   0   'False
         Width           =   315
         _ExtentX        =   556
         _ExtentY        =   503
         DataType        =   4
         IdName          =   "WCPYTHRUBAGNO"
      End
      Begin VB.Label Label8 
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
         Left            =   4440
         TabIndex        =   77
         Top             =   120
         Width           =   1005
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
         Left            =   5520
         TabIndex        =   72
         Top             =   360
         Width           =   2325
      End
      Begin VB.Label LblMain 
         BackStyle       =   0  'Transparent
         BorderStyle     =   1  'Fixed Single
         Caption         =   "  Copy Bag Nos / Bag Alias No."
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
         Left            =   120
         TabIndex        =   71
         Top             =   90
         Width           =   4170
      End
   End
   Begin MwfCtl.MWCTL_BTN1 CmdCpyRfId 
      Height          =   435
      Left            =   5580
      TabIndex        =   67
      ToolTipText     =   "Copy Option"
      Top             =   9225
      Width           =   1530
      _ExtentX        =   2699
      _ExtentY        =   767
      BackColor       =   -2147483633
      MaskColor       =   8454143
      Caption         =   "Copy &Bag Nos / Bag Alias No."
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
      Height          =   435
      Left            =   7140
      TabIndex        =   64
      ToolTipText     =   "Copy Bags"
      Top             =   9225
      Width           =   855
      _ExtentX        =   1508
      _ExtentY        =   767
      BackColor       =   -2147483633
      MaskColor       =   8454143
      Caption         =   "&Copy Bags"
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
      Left            =   8580
      TabIndex        =   8
      TabStop         =   0   'False
      Top             =   9270
      Visible         =   0   'False
      Width           =   555
      _ExtentX        =   979
      _ExtentY        =   714
   End
   Begin MwfCtl.MWCTL_ADC ADC 
      CausesValidation=   0   'False
      Height          =   555
      Left            =   30
      TabIndex        =   9
      Top             =   9270
      Width           =   15105
      _ExtentX        =   26644
      _ExtentY        =   979
      BtnWidth        =   885
      BtnHeight       =   435
      BtnMaskColor    =   8454143
      FontSz          =   8.25
   End
   Begin VB.Frame FRA_KEYS 
      BorderStyle     =   0  'None
      Height          =   280
      Left            =   -36
      TabIndex        =   11
      Top             =   30
      Width           =   5040
      Begin MwfCtl.MWCTL_MED ATXT 
         Height          =   285
         Index           =   4
         Left            =   2835
         TabIndex        =   3
         ToolTipText     =   "Enter Voucher Number"
         Top             =   0
         Width           =   930
         _ExtentX        =   1640
         _ExtentY        =   503
         Alignment       =   2
         Mask            =   "#####0"
         MaxLength       =   6
         DataType        =   1
         DataField       =   "FmNo"
         IdName          =   "FMNO"
      End
      Begin MwfCtl.MWCTL_MED ATXT 
         Height          =   285
         Index           =   6
         Left            =   1620
         TabIndex        =   1
         ToolTipText     =   "Enter Voucher Year"
         Top             =   0
         Width           =   435
         _ExtentX        =   767
         _ExtentY        =   503
         MaxLength       =   2
         DataType        =   4
         DataField       =   "FmYy"
         IdName          =   "FMYY"
      End
      Begin MwfCtl.MWCTL_MED ATXT 
         Height          =   285
         Index           =   31
         Left            =   4500
         TabIndex        =   19
         Top             =   0
         Visible         =   0   'False
         Width           =   135
         _ExtentX        =   238
         _ExtentY        =   503
         Enabled         =   0   'False
         MaxLength       =   3
         DataType        =   4
         DataField       =   "FmCoCd"
         IdName          =   "FMCOCD"
      End
      Begin MwfCtl.MWCTL_MED ATXT 
         Height          =   285
         Index           =   0
         Left            =   960
         TabIndex        =   0
         ToolTipText     =   "Enter Voucher TC"
         Top             =   0
         Width           =   570
         _ExtentX        =   1005
         _ExtentY        =   503
         Enabled         =   0   'False
         MaxLength       =   3
         DataType        =   4
         DataField       =   "FmTc"
         ReCalcParent    =   "FMCHR"
         IdName          =   "FMTC"
      End
      Begin MwfCtl.MWCTL_MED ATXT 
         Height          =   285
         Index           =   1
         Left            =   2160
         TabIndex        =   2
         ToolTipText     =   "Enter Voucher CHR"
         Top             =   0
         Width           =   570
         _ExtentX        =   1005
         _ExtentY        =   503
         MaxLength       =   3
         DataType        =   4
         DataField       =   "FmChr"
         ReCalcOn        =   "FMTC"
         IdName          =   "FMCHR"
      End
      Begin MwfCtl.MWCTL_MED ATXT 
         Height          =   285
         Index           =   2
         Left            =   3840
         TabIndex        =   20
         Top             =   0
         Visible         =   0   'False
         Width           =   375
         _ExtentX        =   661
         _ExtentY        =   503
         Enabled         =   0   'False
         DataField       =   "FmPrtKey"
         IdName          =   "FMPRTKEY"
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
         Left            =   2070
         TabIndex        =   17
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
         Left            =   2745
         TabIndex        =   16
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
         Left            =   1530
         TabIndex        =   14
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
         TabIndex        =   12
         Top             =   -5
         Width           =   1065
      End
   End
   Begin VB.Frame FRA_NKEYS 
      BorderStyle     =   0  'None
      Enabled         =   0   'False
      Height          =   9810
      Left            =   -120
      TabIndex        =   10
      Top             =   0
      Width           =   15315
      Begin VB.Frame FraCB 
         Height          =   2865
         Left            =   5340
         TabIndex        =   21
         Top             =   6285
         Width           =   4560
         Begin MwfCtl.MWCTL_BTN1 CmdCBGo 
            Height          =   285
            Left            =   3870
            TabIndex        =   22
            Top             =   2475
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
         Begin MwfCtl.MWCTL_MED ATXT 
            Height          =   285
            Index           =   61
            Left            =   1185
            TabIndex        =   25
            ToolTipText     =   "Enter From Bag Year"
            Top             =   1020
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
            Left            =   2340
            TabIndex        =   27
            ToolTipText     =   "Enter From Bag Number"
            Top             =   1020
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
            Left            =   1635
            TabIndex        =   26
            ToolTipText     =   "Enter From Bag Character"
            Top             =   1020
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
            Left            =   1185
            TabIndex        =   28
            ToolTipText     =   "Enter To Bag Year"
            Top             =   1305
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
            Left            =   2340
            TabIndex        =   30
            ToolTipText     =   "Enter To Bag Number"
            Top             =   1305
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
            Left            =   1635
            TabIndex        =   29
            ToolTipText     =   "Enter To Bag Character"
            Top             =   1305
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
            Left            =   1890
            TabIndex        =   32
            ToolTipText     =   "Enter From Bag Order Year "
            Top             =   1590
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
            Left            =   3030
            TabIndex        =   34
            ToolTipText     =   "Enter From Bag Order Number"
            Top             =   1590
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
            Left            =   1185
            TabIndex        =   31
            ToolTipText     =   "Enter From Bag Order Tc "
            Top             =   1590
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
            Left            =   2340
            TabIndex        =   33
            ToolTipText     =   "Enter From Bag Order Character"
            Top             =   1590
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
            Left            =   3885
            TabIndex        =   35
            ToolTipText     =   "Enter From Bag Order Serial Number"
            Top             =   1590
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
            Left            =   1890
            TabIndex        =   37
            ToolTipText     =   "Enter To Bag Order Year"
            Top             =   1875
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
            Left            =   3030
            TabIndex        =   39
            ToolTipText     =   "Enter To Bag Order Number"
            Top             =   1875
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
            Left            =   1185
            TabIndex        =   36
            ToolTipText     =   "Enter To Bag Order Tc"
            Top             =   1875
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
            Left            =   2340
            TabIndex        =   38
            ToolTipText     =   "Enter To Bag Order Character"
            Top             =   1875
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
            Left            =   3885
            TabIndex        =   40
            ToolTipText     =   "Enter To Bag Order Serial Number"
            Top             =   1875
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
            Left            =   1185
            TabIndex        =   41
            ToolTipText     =   "Enter From Design Code"
            Top             =   2160
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
            Left            =   1185
            TabIndex        =   42
            ToolTipText     =   "Enter To Design Code"
            Top             =   2445
            Width           =   2415
            _ExtentX        =   4260
            _ExtentY        =   503
            MaxLength       =   15
            DataType        =   4
            IdName          =   "WCBDMCDTO"
         End
         Begin MwfCtl.MWCTL_MED ATXT 
            Height          =   285
            Index           =   33
            Left            =   1185
            TabIndex        =   23
            ToolTipText     =   "Enter Finished Goods Sub Location"
            Top             =   450
            Width           =   1650
            _ExtentX        =   2910
            _ExtentY        =   503
            MaxLength       =   8
            DataType        =   4
            IdName          =   "WCBFGSUBLOCFR"
         End
         Begin MwfCtl.MWCTL_MED ATXT 
            Height          =   285
            Index           =   5
            Left            =   1185
            TabIndex        =   24
            ToolTipText     =   "Enter Finished Goods Sub Location"
            Top             =   750
            Width           =   1650
            _ExtentX        =   2910
            _ExtentY        =   503
            MaxLength       =   8
            DataType        =   4
            IdName          =   "WCBFGSUBLOCTO"
         End
         Begin VB.Label ALBL 
            BackStyle       =   0  'Transparent
            Caption         =   "SubLoc To"
            BeginProperty Font 
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
            TabIndex        =   63
            Top             =   750
            Width           =   1095
         End
         Begin VB.Label ALBL 
            BackStyle       =   0  'Transparent
            Caption         =   "SubLoc Fr"
            BeginProperty Font 
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
            TabIndex        =   62
            Top             =   450
            Width           =   1095
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
            Left            =   90
            TabIndex        =   61
            Top             =   2445
            Width           =   1095
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
            Left            =   90
            TabIndex        =   60
            Top             =   2160
            Width           =   1095
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
            Left            =   2955
            TabIndex        =   59
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
            Index           =   48
            Left            =   2265
            TabIndex        =   58
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
            Index           =   47
            Left            =   1800
            TabIndex        =   57
            Top             =   1875
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
            Left            =   90
            TabIndex        =   56
            Top             =   1875
            Width           =   1035
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
            Index           =   7
            Left            =   3810
            TabIndex        =   55
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
            Index           =   44
            Left            =   2955
            TabIndex        =   54
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
            Index           =   9
            Left            =   2265
            TabIndex        =   53
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
            Index           =   26
            Left            =   1800
            TabIndex        =   52
            Top             =   1590
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
            Left            =   90
            TabIndex        =   51
            Top             =   1590
            Width           =   1035
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
            Index           =   10
            Left            =   3810
            TabIndex        =   50
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
            Index           =   42
            Left            =   2265
            TabIndex        =   49
            Top             =   1305
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
            Left            =   1560
            TabIndex        =   48
            Top             =   1305
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
            Left            =   90
            TabIndex        =   47
            Top             =   1305
            Width           =   855
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
            Left            =   2265
            TabIndex        =   46
            Top             =   1020
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
            Index           =   23
            Left            =   1560
            TabIndex        =   45
            Top             =   1020
            Width           =   105
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
            Left            =   90
            TabIndex        =   44
            Top             =   1020
            Width           =   855
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
            Index           =   0
            Left            =   -30
            TabIndex        =   43
            Top             =   90
            Width           =   5100
         End
      End
      Begin VB.Frame FraNKeyAll 
         BorderStyle     =   0  'None
         Height          =   9465
         Left            =   150
         TabIndex        =   13
         Top             =   30
         Width           =   15135
         Begin MwfCtl.MWCTL_MED ATXT 
            Height          =   285
            Index           =   30
            Left            =   895
            TabIndex        =   4
            ToolTipText     =   "Enter Voucher Date"
            Top             =   285
            Width           =   855
            _ExtentX        =   1508
            _ExtentY        =   503
            MaxLength       =   8
            DataType        =   3
            DataField       =   "FmDt"
            IdName          =   "FMDT"
         End
         Begin MwfCtl.MWCTL_MED ATXT 
            Height          =   615
            Index           =   3
            Left            =   6135
            TabIndex        =   6
            ToolTipText     =   "Enter Vch description"
            Top             =   0
            Width           =   4755
            _ExtentX        =   8387
            _ExtentY        =   1085
            Alignment       =   3
            MaxLength       =   60
            DataType        =   4
            DataField       =   "FMDESC"
            IdName          =   "FMDESC"
         End
         Begin MwfCtl.MWCTL_FLX GrdFmd 
            Height          =   8505
            Left            =   15
            TabIndex        =   7
            ToolTipText     =   "Fg Movement Details"
            Top             =   615
            Width           =   15105
            _ExtentX        =   26644
            _ExtentY        =   15002
            Cols            =   14
            colname1        =   "FMDSR"
            heading1        =   "Sr.  "
            datafld1        =   "FmdSr"
            datatype1       =   1
            mask1           =   "##0"
            maxlength1      =   3
            style1          =   2
            tooltiptext1    =   "Enter Sr. No"
            colname2        =   "WFGBNO"
            datatype2       =   4
            maxlength2      =   18
            tooltiptext2    =   "Enter Bag Number Using Barcode Reader"
            colname3        =   "FMDBYY"
            heading3        =   "BagYy"
            datafld3        =   "FmdBYy"
            datatype3       =   4
            recalcparent3   =   "GrdFmd(FMDFRFGSUBLOC),GrdFmd(FMDBNO),GrdFmd(WFMDDMCD)"
            maxlength3      =   2
            tooltiptext3    =   "Enter Bag Year"
            colname4        =   "FMDBCHR"
            heading4        =   "BagChr"
            datafld4        =   "FmdBChr"
            datatype4       =   4
            recalcparent4   =   "GrdFmd(FMDFRFGSUBLOC),GrdFmd(FMDBNO),GrdFmd(WFMDDMCD)"
            maxlength4      =   3
            tooltiptext4    =   "Enter Bag Chr"
            colname5        =   "FMDBNO"
            heading5        =   "Bag No "
            datafld5        =   "FmdBNo"
            datatype5       =   1
            mask5           =   "######0"
            recalcon5       =   "GrdFmd(FMDBYY),GrdFmd(FMDBCHR)"
            recalcparent5   =   "GrdFmd(FMDFRFGSUBLOC),GrdFmd(WFMDDMCD)"
            maxlength5      =   7
            tooltiptext5    =   "Enter Bag Number"
            colname6        =   "WFMDDMCD"
            heading6        =   "Design Code      "
            datatype6       =   4
            recalcon6       =   "GrdFmd(FMDBYY),GrdFmd(FMDBCHR),GrdFmd(FMDBNO)"
            maxlength6      =   15
            style6          =   2
            colname7        =   "WFMDSFX"
            heading7        =   "Suffix          "
            datatype7       =   4
            maxlength7      =   15
            style7          =   2
            colname8        =   "WFMDDMSZ"
            heading8        =   "Size        "
            datatype8       =   4
            maxlength8      =   5
            style8          =   2
            colname9        =   "FMDFRFGSUBLOC"
            heading9        =   "Fr FgSubLoc      "
            datafld9        =   "FmdFrFgSubLoc"
            datatype9       =   4
            recalcon9       =   "GrdFmd(FMDBYY),GrdFmd(FMDBCHR),GrdFmd(FMDBNO)"
            maxlength9      =   8
            style9          =   2
            tooltiptext9    =   "Enter From FG Sub Loc"
            colname10       =   "FMDTOFGSUBLOC"
            heading10       =   "To FgSubLoc     "
            datafld10       =   "FmdToFgSubLoc"
            datatype10      =   4
            maxlength10     =   8
            tooltiptext10   =   "Enter To FG Sub Loc"
            colname11       =   "WFMDBQTY"
            heading11       =   "Bag Qty"
            datatype11      =   2
            mask11          =   "##0.0"
            maxlength11     =   5
            style11         =   2
            colname12       =   "FMDKEY"
            datafld12       =   "FmdKey"
            datatype12      =   1
            mask12          =   "#######0"
            maxlength12     =   8
            style12         =   0
            colname13       =   "FMDPRTKEY"
            heading13       =   "Prtn "
            datafld13       =   "FmdPrtKey"
            maxlength13     =   1
            style13         =   0
         End
         Begin MwfCtl.MWCTL_MED ATXT 
            Height          =   285
            Index           =   165
            Left            =   3540
            TabIndex        =   5
            ToolTipText     =   "Enter To FG Sub Location"
            Top             =   285
            Width           =   1380
            _ExtentX        =   2434
            _ExtentY        =   503
            MaxLength       =   8
            DataType        =   4
            IdName          =   "WFMTOFGSUBLOC"
         End
         Begin VB.Label LblTToBLoc 
            BackStyle       =   0  'Transparent
            Caption         =   "To FgSubLoc"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Left            =   2100
            TabIndex        =   65
            Top             =   285
            Width           =   1470
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
            Index           =   3
            Left            =   5580
            TabIndex        =   18
            Top             =   0
            Width           =   570
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
            Left            =   0
            TabIndex        =   15
            Top             =   285
            Width           =   855
         End
      End
   End
   Begin MwfCtl.MWCTL_BTN1 MWCTL_BTN11 
      Height          =   435
      Left            =   -1200
      TabIndex        =   66
      ToolTipText     =   "Copy Bags"
      Top             =   0
      Width           =   855
      _ExtentX        =   1508
      _ExtentY        =   767
      BackColor       =   -2147483633
      MaskColor       =   8454143
      Caption         =   "&Copy Bags"
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
Attribute VB_Name = "EmrFrmFgMvm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'---- General Comments on the Entry --------'
' 1 Menu
'   1. 'FM' - Finished Goods Movement(Bag)
' X-Fm, A-Fmd
'--------------------

Option Explicit
Dim moCn As MwfLib.MDOConnection '*** (Bef speed) bef 24/11/06

' ****** Manali 3.5.0 - 08/11/08 - Copy Bags Frame Added
Enum en_FmFra
  CpyBag = 0
  CpyRfId = 1
End Enum
' ****** Manali 3.5.0 - 08/11/08 - Copy Bags Frame Added

'*** Declare the variables ms_LockDt, mf_DtLocked for the Lock Date Check
Dim ms_LockDt As String
Dim ms_FutureDt As String
Dim mf_DtLocked As Boolean
Dim ms_PurgeDt As String    '****** Sachin 3.02 26-12-07
Dim pb_CpyRfId As Boolean
Dim ms_TcTyp As String, ms_EntTyp As String, ms_FmCoCd As String
Dim ms_FmTc As String, ms_FmYy As String, ms_FmChr As String

Private Sub ADC_Load()
  '*** Set the First Key Control property to the field FmChr
  '*** Set the First Non Key Control property to FmDt
  '*** Set the Child Property Of ADC to GrdFmd
  '*** Set the Previous and Next Control Properties for the Grids GrdFmd
  '*** Set the Hot Keys for the Grids GrdFmd as 'A'
  '*** Bring the FraNKeyAll frame to the top
  '*** Get the default values for Tc and Yy from param and head files for the first time the
  '    form is loaded
   
  Set moCn = ADC.Connection  '*** (Bef speed) bef 24/11/06
  Set ADC.FirKeyCtl = ADC("FmChr")
  Set ADC.FirNKeyCtl = ADC("FmDt")
  ADC.Child = "GrdFmd"
  
  Set GrdFmd.PrevCtl = ADC("FmDesc")
  Set GrdFmd.NextCtl = ADC
  
  GrdFmd.HotKey = "A"
  FraNKeyAll.ZOrder (0)

  ms_FmCoCd = gs_CoCd
  ms_TcTyp = ADC.MenuCd
  ms_EntTyp = moCn.GetFldVal("Select PValue1 from Param where PTyp= 'TCTYP' and " + _
            "PMCd= '" + ms_TcTyp + "'")
  ms_FmTc = moCn.GetFldVal("Select PValue from Param where PTyp= 'TCTYP' and " + _
               "PMCd='" + ms_TcTyp + "' ")
  ms_FmYy = moCn.GetFldVal("Select hYy from Head where HCoCd= '" + ms_FmCoCd + "' and HCd='" + ctSelfCmCd + "'")
  
    '****** Sachin 3.02.0 - Hiding the Partition Key Column based on Partition Selected
  ADC("FMPRTKEY").Visible = IIF(gs_Partition = ctCurrPrtn, False, True)
  GrdFmd.ColProp("FmdPrtKey").Style = IIF(gs_Partition = ctCurrPrtn, fgcinvisible, fgcReadOnly)
  
  '4.1.0.0 - RFID Module Check  '6.1b-22 Rfid module check is removed
'  If Not moCn.RecSeek("Select 'x' from Head Where HCoCd='" + gs_CoCd + "' and HCd='" + ctSelfCmCd + "' and HRfIdMod='Y' ") Then
'    CmdCpyRfId.Enabled = False: CmdCpyRfId.Visible = False
'  End If
  'sv.9 getting purge date from vparam
  ms_PurgeDt = GetPurgeDt(gs_CoCd, "")
          
End Sub
Private Sub ADC_KeyWhen()
  '*** Set the dafault value for the form level variable mf_DtLocked as false.
  '*** Hide all frames except Frame FraNKeyAll
  '*** setting the default values for FmTc, FmYy and FmChr from the previous entry
  
  mf_DtLocked = False
   
  Call EnaDisaCmds(True)    ' ****** Manali 3.5.0
  Call HideAllFras
  
  If ms_FmCoCd <> "" Then ADC("FmCoCd") = ms_FmCoCd
  If ms_FmTc <> "" Then ADC("FmTc") = ms_FmTc
  If ms_FmYy <> "" Then ADC("FmYy") = ms_FmYy
  If ms_FmChr <> "" Then ADC("FmChr") = ms_FmChr
  
  '****** Sachin 2.14.0 - [24-11-2006] - [Multi-Company] ******
  Me.Caption = GetFrmCaption(gs_CoCd, "FG Movement Entry")
  '****** Sachin 2.14.0 - [24-11-2006] - [Multi-Company] ******
  
End Sub
Private Sub ADC_KeyValid(Cancel As Boolean, ErrMsg As String)
  '*** Check whether the FmYy, FmChr are valid
  If ADC("FmYy") = "" Then Cancel = True: ErrMsg = "Enter Voucher Year": Exit Sub
  If ADC("FmChr") = "" Then Cancel = True: ErrMsg = "Enter Voucher Character": Exit Sub
End Sub
Private Sub ADC_NonKeyWhen(Cancel As Boolean, ErrMsg As String)
  '*** Get the Lock Date from Param('LOCKDT', 'LOCKDT') and store in the form level variable ms_LockDt
  '    If ms_LockDt is blank then store '01/01/80' in ms_LockDt
  '*** Store the FmTc, FmYy and FmChr values for the next entry
  '*** Give default values for the fields FmDt
  '*** If the entry Date is < ms_LockDt then mf_DtLocked= true (i.e the entry
  '    If mf_DtLocked = true then the entry is locked. i.e. the user cannot add, edit or
  '    delete records in the grids GrdFmd. Similarly if the user does not have rights to
  '    any of locations, then the user cannot add, edit or delete records in the grids
  '    GrdFmd.

      If ADC.IsNew And (Not moCn.RecSeek("select vPSCd from vParam where " + _
             "vPCoCd = '" + ADC("FmCoCd") + "' and vPTyp= 'CHR' and " + _
             "vPMCd = '" + ADC("FmTc") + "' and vPSCd='" + ADC("FmChr") + "' and vPValidYn<>'N'")) Then
       Cancel = True: ErrMsg = "Voucher Character No More Valid": ADC.AllowSave = False: Exit Sub
      Else
        ADC.AllowSave = True
      End If

  '****** Sachin 3.02 26-12-07 - Setting partition key to current
  If ADC.IsNew And Trim(ADC("FmPrtKey")) = "" Then ADC("FmPrtKey") = ctCurrPrtn

 ' ms_LockDt = moCn.GetFldVal("Select PDesc225 from Param where PTyp= 'LOCKDT' and PMCd= 'LOCKDT'")
 ' ms_PurgeDt = moCn.GetFldVal("Select PDesc225 from Param where PTyp= 'PURGEDT' and PMCd= 'PURGEDT'")
  
 ' If ms_LockDt = "" Then ms_LockDt = "01/01/80"
 ' If ms_PurgeDt = "" Then ms_PurgeDt = "01/01/80"   '****** Sachin 3.02 26-12-07
   
 '6.1_51 wVchrNo added to GetLockDt to check that given voucher number is exist in unlock ptyp,
  'if given voucher is in unlock ptyp it will not check lock date
  Dim wVchrNo As String
  wVchrNo = ADC("FmTc") + "/" + ADC("FmYy") + "/" + ADC("FmChr") + "/" + CStr(ADC("FmNo"))
  
   Call GetLockDt(ms_FmCoCd, "TC", ms_FmTc, ms_LockDt, ms_FutureDt, wVchrNo)
   
  
  FraNKeyAll.Enabled = True
  ms_FmTc = ADC("FmTc")
  ms_FmYy = ADC("FmYy")
  ms_FmChr = ADC("FmChr")

  Call EnaDisaCmds(False)
  
  If ADC.AddFndMode = xaddmode Then
    '53 company code parameter added
    ADC("FmDt") = GetNextDate(ADC("FmCoCd")) 'moCn.SrvrDate
  End If
    
  
  mf_DtLocked = IIF(ADC("FmDt") < CDate(ms_LockDt) Or ADC("FmDt") > CDate(ms_FutureDt), True, False)    
  
  If mf_DtLocked = True Then
    GrdFmd.AllowAdd = False: GrdFmd.AllowDelete = False
  Else
    GrdFmd.AllowAdd = True: GrdFmd.AllowDelete = True
    '  If adc("WFmToFgSubLoc") <> "" Then
    '    If Not moCn.RecSeek("select 'x' From vParam where vPTyp='FgSubLoc' and vPCoCd= '" + gs_CoCd + "' and vPMCd='" + adc("WFmToFgSubLoc") + "'" + _
    '                    " and (vPValue4='' Or ','+vPValue4+',' Like '%," + gs_UsrCd + ",%')") Then
    '          mf_ValidLocSeq = False
    '          GrdFmd.AllowAdd = False: GrdFmd.AllowDelete = False
    '     Else
    '          If mf_ValidLocSeq = True Then
    '            GrdFmd.AllowAdd = True: GrdFmd.AllowDelete = True
    '          End If
    '    End If
    '  End If
  End If
  'vk.4 setting default values
  OptCpyBagFr(0).Value = True
  ADC("wCpyThruBagNo") = "Y"
  
End Sub
Private Sub ATXT_FldWhen(Index As Integer, ByVal IdName As String, Cancel As Boolean, ErrMsg As String)
  '*** Set Helps For Each Field
  '*** Do not allow the user to edit the voucher date if record exists in the 'A' grid
  
  If mf_DtLocked = True Then Cancel = True: ErrMsg = "Cannot Edit Locked Entry": Exit Sub
  
  Select Case IdName
  Case Is = UCase("FmYy")
    Call HlpList.PMCd("YY")
  Case Is = UCase("FmChr")
' **** Zubin 211 **** '
    'Call HlpList.vPSCd(ADC("FmCoCd"), "CHR", ADC("FmTc"))
    Call HlpList.vPSCd(ADC("FmCoCd"), "CHR", ADC("FmTc"), "Y")
' **** Zubin 211 **** '
  Case Is = UCase("FmNo")
    Call HlpList.FmNo(ADC("FmCoCd"), ADC("FmTc"), ADC("FmYy"), ADC("FmChr"))
  Case Is = UCase("FmDt")
    If GrdFmd.Rows - 1 > 0 Then Cancel = True: ErrMsg = "Cannot Edit Date": Exit Sub
    
  ' ***** Manali 3.5.0 - 08/11/08 - Copy Bags Frame Added
  Case Is = UCase("wCbFgSubLocFr"), UCase("wCbFgSubLocTo")
    Call HlpList.vPMCd(ADC("FmCoCd"), "FgSubLoc")
  Case Is = UCase("wCbBYyFr"), UCase("wCbBYyTo")
    Call HlpList.PMCd("YY")
  Case Is = UCase("wCbBChrFr"), UCase("wCbBChrTo")
    Call HlpList.PMCd("BCHR", "'P'")
  Case Is = UCase("wCbBNoFr")
   ' wOthLocBag = moCn.GetFldVal("Select hOthLocBag from Head where HCoCd= '" + ADC("FgCoCd") + "' and HCd='" + ctSelfCmCd + "'")
  '  Call HlpList.BNo(ADC("FgCoCd"), ADC("wCbBYyFr"), ADC("wCbBChrFr"), wOthLocBag, ADC("FgFrBLoc"))
    Call HlpList.BNo(ADC("FmCoCd"), ADC("wCbBYyFr"), ADC("wCbBChrFr"), "N", ctFgLoc)
  Case Is = UCase("wCbBNoTo")
    'wOthLocBag = moCn.GetFldVal("Select hOthLocBag from Head where HCoCd= '" + ADC("FgCoCd") + "' and HCd='" + ctSelfCmCd + "'")
    'Call HlpList.BNo(ADC("FgCoCd"), ADC("wCbBYyTo"), ADC("wCbBChrTo"), wOthLocBag, ADC("FgFrBLoc"))
    Call HlpList.BNo(ADC("FmCoCd"), ADC("wCbBYyTo"), ADC("wCbBChrTo"), "N", ctFgLoc)
  Case Is = UCase("wCbOdTcFr"), UCase("wCbOdTcTo") 'indigo.164 OS tc added
    Call HlpList.PMCd("TC", "'SO','OS'")
  Case Is = UCase("wCbOdYyFr"), UCase("wCbOdYyTo")
    Call HlpList.PMCd("YY")
  Case Is = UCase("wCbOdChrFr")
    Call HlpList.vPSCd(ADC("FmCoCd"), "CHR", ADC("wCbOdTcFr"))
  Case Is = UCase("wCbOdChrTo")
    Call HlpList.vPSCd(ADC("FmCoCd"), "CHR", ADC("wCbOdTcTo"))
  Case Is = UCase("wCbOdNoFr")
    Call HlpList.OmNo(ADC("FmCoCd"), ADC("wCbOdTcFr"), ADC("wCbOdYyFr"), ADC("wCbOdChrFr"))
  Case Is = UCase("wCbOdNoTo")
    Call HlpList.OmNo(ADC("FmCoCd"), ADC("wCbOdTcFr"), ADC("wCbOdYyFr"), ADC("wCbOdChrFr"))
  Case Is = UCase("wCbOdSrFr")
    Call HlpList.OdSr(ADC("FmCoCd"), ADC("wCbOdTcFr"), ADC("wCbOdYyFr"), ADC("wCbOdChrFr"), ADC("wCbOdNoFr"))
  Case Is = UCase("wCbOdSrTo")
    Call HlpList.OdSr(ADC("FmCoCd"), ADC("wCbOdTcTo"), ADC("wCbOdYyTo"), ADC("wCbOdChrTo"), ADC("wCbOdNoTo"))
  Case Is = UCase("wCbDmCdFr"), UCase("wCbDmCdTo")  'indigo.164 SM design type added
    Call HlpList.DmCd("DM','SM")
  ' ***** Manali 3.5.0 - 08/11/08 - Copy Bags Frame Added
  
  Case Is = UCase("wFmToFgSubLoc")
    Call HlpList.vPMCd(ADC("FmCoCd"), "FgSubLoc")
  '*********Bhavna cpy thru bag no
  Case Is = UCase("wCPYTHRUBAGNO")
        If ADC("WCPYTHRUBAGNO") = "" Then
            ADC("WCPYTHRUBAGNO") = "N"
        End If

  End Select
End Sub

Private Sub atxt_ReCalculate(Index As Integer, ByVal IdName As String, ByVal OldVal As Variant, ByVal RowNum As Integer, ByVal SourceName As String, ByVal SourceColName As String)
  '*** FmChr on FmTc                Normal
      'PDesc225('TC', adc(FmTc))
  
  Select Case UCase(IdName)
  Case Is = UCase("FmChr")
    If ADC.Mode = xNorm Then _
       ADC("FmChr") = moCn.GetFldVal("Select vPValue from vParam where " + _
                                     "vPCoCd= '" + ADC("FmCoCd") + "' and vPTyp= 'DEFCHR' " + _
                                     "and vPMCd= '" + ADC("FmTc") + "'")
  End Select
End Sub
Private Sub ATXT_FldChange(Index As Integer, ByVal IdName As String, Cancel As Boolean, ErrMsg As String, ByVal pv_NewValue As Variant)
  '*** FmTc             PMCd('TC', PValue= ms_TcTyp) and the type of entry should be 'FgMvm'
  '*** FmYy             PMCd('YY')
  '*** FmChr            PSCd('CHR', adc(FmTc)) or ''
  '*** FmDt             Not Blank
    
  Select Case UCase(IdName)
  Case Is = UCase("FmTc")
    Dim wPValue1 As String
    Cancel = Not moCn.RecSeek("select PMCd from Param where PTyp= 'TC' and " + _
                         "PMCd = '" + pv_NewValue + "' and " + _
                         "PValue= '" + ms_TcTyp + "' ")
    If Cancel = True Then ErrMsg = "Invalid Voucher Tc": Exit Sub
  Case Is = UCase("FmYy")
    Cancel = Not moCn.RecSeek("select PMCd from Param where PTyp= 'YY' and " + _
                         "PMCd = '" + pv_NewValue + "'")
    If Cancel = True Then ErrMsg = "Invalid Voucher Year": Exit Sub
  Case Is = UCase("FmChr")
    Cancel = (pv_NewValue <> "") And (Not moCn.RecSeek("select vPSCd from vParam where " + _
             "vPCoCd = '" + ADC("FmCoCd") + "' and vPTyp= 'CHR' and " + _
             "vPMCd = '" + ADC("FmTc") + "' and vPSCd='" + pv_NewValue + "'"))
' **** Zubin 211 **** '
'''''    ErrMsg = "Invalid Voucher Character": Exit Sub
    If Cancel = True Then ErrMsg = "Invalid Voucher Character": Exit Sub
    Dim ws_UsrArr() As String, wi_Index As Integer
    ws_UsrArr = Split(moCn.GetFldVal("Select vPValue4 from vParam where vPCoCd= '" + ADC("FmCoCd") + "' and " + _
                                        "vPTyp='CHR' and vPMCd='" + ADC("FmTc") + "' and vPSCd= '" + pv_NewValue + "'"), ",")
    If IsArray(ws_UsrArr) Then
      For wi_Index = 0 To UBound(ws_UsrArr)
        Cancel = IIF(UCase(gs_UsrCd) = UCase(ws_UsrArr(wi_Index)), False, True)
        If Cancel = False Then: Exit For
      Next wi_Index
    End If
    If Cancel = True Then ErrMsg = "Access To This Voucher Character Is Denied To User " + gs_UsrCd: Exit Sub
' **** Zubin 211 **** '
  Case Is = UCase("FmDt")
    If pv_NewValue = "" Then Cancel = True: ErrMsg = "Date Cannot Be Blank": Exit Sub
    If pv_NewValue > CDate(ms_FutureDt) Then Cancel = True: ErrMsg = "Finished Goods Movement Voucher Date Cannot Be Greater Than Future Date": Exit Sub
    If pv_NewValue < CDate(ms_LockDt) Then Cancel = True: ErrMsg = "Finished Goods Movement Voucher Date Cannot Be Less Than Lock Date": Exit Sub
    'CC.999.13 check for purge date added
    If pv_NewValue <= CDate(ms_PurgeDt) Then Cancel = True: ErrMsg = "Finished Goods Movement Voucher Date Cannot Be Less Than Or Equal To Purge Date": Exit Sub
        
    If moCn.GetFldVal("Select hHolChkYN From Head where HCd='ZSELF' and HCoCd='" + gs_CoCd + "'") = "Y" Then
        If IsHOLYN(ADC("FmCoCd"), pv_NewValue) Then Cancel = True: ErrMsg = "Voucher Cannot be Added on Holidays": Exit Sub
    End If
  
  ' ***** Manali 3.10.0 - 12/04/12 - Bar Code
  Case Is = UCase("wFmToFgSubLoc")
     Cancel = (pv_NewValue <> "") And (Not moCn.RecSeek("Select vPMCd From vParam Where " + _
              "vPCoCd= '" + ADC("FmCoCd") + "' and vPTyp= 'FGSUBLOC' And " + _
              "vPMCd = '" + pv_NewValue + "'"))
    ErrMsg = "Invalid To FG Sub Location": Exit Sub
  ' ***** Manali 3.10.0 - 12/04/12 - Bar Code
'**** Bhavna added cpy thru bag no
  Case Is = UCase("wCpyThruBagNo")
    If ADC.Mode = xNorm Then
      Cancel = Not moCn.RecSeek("select PMCd from Param where PTyp= 'YN' and PMCd = '" + pv_NewValue + "'")
      ErrMsg = "Enter 'Y' to Copy thru Bag No Or 'N' to Copy thru RfId No ": Exit Sub
    End If
  End Select
End Sub

Private Sub CmdCB_Click()
  Call DispFra(CpyBag)
End Sub

Private Sub CmdCBGo_Click()
  Call CopyBag
  Call DispFra(CpyBag)
End Sub

Private Sub CmdCpyRfId_Click()
pb_CpyRfId = True
Call DispFra(CpyRfId)
End Sub

Private Sub CmdCpyRfIdGo_Click()
    If Not ADC.SaveRec Then
       DispMsg "X Save Failed", etInfo
       Exit Sub
    End If
  Call CopyBag
  Call DispFra(CpyRfId)
  pb_CpyRfId = False
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
' ***** Manali 3.10.0 - 12/04/12 - Bar Code
If KeyAscii <> vbKeyReturn Then Exit Sub
If UCase(Me.ActiveControl.Name) = UCase("GrdFmd") Then
  With Me.ActiveControl
    If .Col = 2 Then
          If .Rows - 1 = .Row Then
            .Col = 3
          End If
    End If
  End With
End If
' ***** Manali 3.10.0 - 12/04/12 - Bar Code
End Sub



Private Sub GrdFmd_RowWhen(ByVal RowNum As Integer)
  '*** If GrdFmd.IsNew Then FdSr = Max + 1 and Autogenerate FdKey and
  With GrdFmd
    If GrdFmd.IsNew(RowNum) Then
      If .Value(RowNum, "FmdSr") = 0 Then .Value(RowNum, "FmdSr") = .MaxVal("FmdSr") + 1
      .Value(RowNum, "FmdKey") = moCn.AutoGen("Fmd", "FmdKey", ADC("FmCoCd") + ADC("FmYy"))
      .Value(RowNum, "FmdPrtKey") = ctCurrPrtn  '****** Sachin 3.02.0 Default Value for Partition Key
    End If
    
  If Not mf_DtLocked Then
     If Not .IsNew(RowNum) Then
         Dim wBSubLoc As String
          wBSubLoc = moCn.GetFldVal("Select BFgSubLoc from Bag Where " + _
                                      "BCoCd= '" + ADC("FmCoCd") + "' and " + _
                                      "BYy= '" + .Value(RowNum, "FmdBYy") + "' And " + _
                                      "BChr = '" + .Value(RowNum, "FmdBChr") + "' And " + _
                                      "BNo= " + CStr(.Value(RowNum, "FmdBNo")) + " and BPrtKey='" + ctCurrPrtn + "'")
         .AllowDelete = moCn.RecSeek("select 'x' From vParam where vPTyp='FgSubLoc' and vPCoCd= '" + gs_CoCd + "' and vPMCd='" + wBSubLoc + "'" + _
                              " and (vPValue4='' Or ','+vPValue4+',' Like '%," + gs_UsrCd + ",%')")
         If .AllowDelete = False Then
            DispMsg "Cannot Delete As Bag Sub Loc. is Not Accessible to the User", etError
         Else
            .AllowDelete = True
         End If
    End If
   End If
    
  End With
End Sub
Private Sub GrdFmd_ColWhen(ByVal RowNum As Integer, ByVal ColName As String, Cancel As Boolean, ErrMsg As String, ByVal Action As MwfCtl.en_fgAction)
  '*** Set Helps For Each Field
  '*** If 1st row, then get the default values for FdBYy and FdBChr from the Head file
      'If Row> 1 then get the defa values for FdBYy and FdBChr from the previous row of the grid
      'If not GrdFmd.IsNew the cannot edit FdBYy, FdBChr and FdBNo
      'If the value of the field hOthLocBag of the Head table is 'N' then only bags belonging
      'to FgFrBLoc is shown in the help of FdBNo
      'If hOthLocBag = 'Y' then bags belonging to all the Process locations are shown in the
      'help of FdBNo. The bags belonging to FgFrBLoc Location are shown first.
  
  '****** Sachin 3.02.0 - Same Lock Date Check for Previous Partition Entries also
  If mf_DtLocked = True Then Cancel = True: ErrMsg = "Cannot Edit Locked/Previous Partition Entry": Exit Sub
  
  With GrdFmd
    'If .IsNew(RowNum) <> True And ColName <> UCase("FmdToFgSubLoc") Then Cancel = True
    
     If Not .IsNew(RowNum) Then
         Dim wBSubLoc As String
          wBSubLoc = moCn.GetFldVal("Select BFgSubLoc from Bag Where " + _
                                      "BCoCd= '" + ADC("FmCoCd") + "' and " + _
                                      "BYy= '" + .Value(RowNum, "FmdBYy") + "' And " + _
                                      "BChr = '" + .Value(RowNum, "FmdBChr") + "' And " + _
                                      "BNo= " + CStr(.Value(RowNum, "FmdBNo")) + " and BPrtKey='" + ctCurrPrtn + "'")
          Cancel = Not moCn.RecSeek("Select 'x' From vParam where vPTyp='FgSubLoc' and vPCoCd= '" + gs_CoCd + "' and vPMCd='" + wBSubLoc + "'" + _
                              " and (vPValue4='' Or ','+vPValue4+',' Like '%," + gs_UsrCd + ",%')")
        If Cancel = True Then ErrMsg = "Invalid Bag Number As Bag Sub Loc. is Not Accessible to the User": Exit Sub
    End If
           
    Select Case ColName
' **** Zubin 211 **** '
    Case Is = UCase("wFgBNo")
      If Not .IsNew(RowNum) Then Cancel = True: ErrMsg = "Cannot Edit FG Bag Number": Exit Sub
      If Not .IsNew(RowNum) And .Value(RowNum, "FmdBNo") <> 0 Then Cancel = True: ErrMsg = "Cannot Edit FG Bag Number": Exit Sub
' **** Zubin 211 **** '
    Case Is = UCase("FmdBYy")
' **** Zubin 211 **** '
'''''      If .Value(RowNum, "wFGBNo") <> "" Then Cancel = True: ErrMsg = "Cannot Edit Bag Year When Bag Number Is Read Using Barcode Reader": Exit Sub
' **** Zubin 211 **** '
      If .IsNew(RowNum) And .Value(RowNum, "FmdBYy") = "" Then
        If RowNum = 1 Then .Value(RowNum, "FmdBYy") = _
           moCn.GetFldVal("Select hYy from Head where HCoCd= '" + ADC("FmCoCd") + "' and HCd='" + ctSelfCmCd + "'")
        If RowNum > 1 Then .Value(RowNum, "FmdBYy") = .Value(RowNum - 1, "FmdBYy")
      Else
           Cancel = True: ErrMsg = "Cannot Edit Bag Year": Exit Sub
      End If
      Call HlpList.PMCd("YY")
      
    Case Is = UCase("FmdBChr")
' **** Zubin 211 **** '
'''''      If .Value(RowNum, "wFGBNo") <> "" Then Cancel = True: ErrMsg = "Cannot Edit Bag Character When Bag Number Is Read Using Barcode Reader": Exit Sub
' **** Zubin 211 **** '
      If .IsNew(RowNum) And .Value(RowNum, "FmdBChr") = "" Then
        If RowNum = 1 Then .Value(RowNum, "FmdBChr") = _
           moCn.GetFldVal("Select hBChr from Head where HCoCd= '" + ADC("FmCoCd") + "' and HCd='" + ctSelfCmCd + "'")
        If RowNum > 1 Then .Value(RowNum, "FmdBChr") = .Value(RowNum - 1, "FmdBChr")
      Else
        Cancel = True: ErrMsg = "Cannot Edit Bag Character": Exit Sub
      End If
      Call HlpList.PMCd("BCHR")
    Case Is = UCase("FmdBNo")
' **** Zubin 211 **** '
'''''      If .Value(RowNum, "wFGBNo") <> "" Then Cancel = True: ErrMsg = "Cannot Edit Bag Number When Bag Number Is Read Using Barcode Reader": Exit Sub
' **** Zubin 211 **** '
      If Not .IsNew(RowNum) Then Cancel = True: ErrMsg = "Cannot Edit Bag Number": Exit Sub
      '6.1_20 once bagno. entered we cant change it again
      If .IsNew(RowNum) And .Value(RowNum, "FmdBNo") <> 0 Then Cancel = True: ErrMsg = "Cannot Edit Bag Number": Exit Sub
         Call HlpList.BNo(ADC("FmCoCd"), .Value(RowNum, "FmdBYy"), .Value(RowNum, "FmdBChr"), "N", ctFgLoc)
    'Case Is = UCase("FmdFrFgSubLoc")
    '  If Not .IsNew(RowNum) Then Cancel = True: ErrMsg = "Cannot Edit Bag From Location": Exit Sub
    '  Call HlpList.vPMCd(ADC("FmCoCd"), "FgSubLoc")
    Case Is = UCase("FmdToFgSubLoc")
      Call HlpList.vPMCd(ADC("FmCoCd"), "FgSubLoc")
    End Select
  End With
End Sub
Private Sub GrdFmd_ReCalc(ByVal RowNum As Integer, ByVal ColName As String, ByVal OldVal As Variant, ByVal SourceName As String, ByVal SourceColName As String)
  '*** FdBNo on FdBYy,FdBChr                                    Normal
      'FdBNo = 0
' **** Zubin 211 **** '
  Dim ws_FGBNo() As String
' **** Zubin 211 **** '

  With GrdFmd
    Select Case ColName
    Case Is = UCase("FmdBNo")
' **** Zubin 211 **** '
'''''      If .Mode = fgmnorm Then .Value(RowNum, "FmdBNo") = 0
      If RowNum > 0 Then If .Mode = fgmnorm Then If .Value(RowNum, "wFgBNo") = "" Then .Value(RowNum, "FmdBNo") = 0
' **** Zubin 211 **** '
    '****** Sachin 3.02 26-12-07 - Corresponding Partition
    Case Is = UCase("FmdFrFgSubLoc")
      If .Mode = fgmnorm Then .Value(RowNum, "FmdFrFgSubLoc") = moCn.GetFldVal("Select BFgSubLoc from Bag where " + _
                              "BCoCd= '" + ADC("FmCoCd") + "' and BYy = '" + .Value(RowNum, "FmdBYy") + "' and " + _
                              "BChr = '" + .Value(RowNum, "FmdBChr") + "' and " + _
                              "BNo = " + CStr(.Value(RowNum, "FmdBNo")) + IIF(gs_Partition = ctCurrPrtn, " and BPrtKey='" + ctCurrPrtn + "' ", ""))
                              
    Case Is = UCase("WFMDDMCD")
        If .Mode = fgmnorm Or .Mode = fgminit Then
          Dim wVal As String, wArr() As String
          'M.suresh.1- Bag qty added
          wVal = moCn.GetFldVal("Select BOdDmCd+'|'+BOdDmSz+'|'+BOdSfx+'|'+Ltrim(str(BQty)) from Bag where " + _
                              "BCoCd= '" + ADC("FmCoCd") + "' and BYy = '" + .Value(RowNum, "FmdBYy") + "' and " + _
                              "BChr = '" + .Value(RowNum, "FmdBChr") + "' and " + _
                              "BNo = " + CStr(.Value(RowNum, "FmdBNo")) + IIF(gs_Partition = ctCurrPrtn, " and BPrtKey='" + ctCurrPrtn + "' ", ""))
          wArr = Split(wVal, "|")
          If UBound(wArr) = 3 Then
            .Value(RowNum, "wFmdDmCd") = wArr(0)
            .Value(RowNum, "wFmdDmSz") = wArr(1)
            .Value(RowNum, "wFmdSfx") = wArr(2)
            .Value(RowNum, "wFmdBQty") = wArr(3)
                        
          End If
        End If
    End Select
  End With
End Sub
Private Sub GrdFmd_ColChange(ByVal RowNum As Integer, ByVal ColName As String, ByVal NewValue As Variant, Cancel As Boolean, ErrMsg As String)
  '*** FmdBYy            valid PMCd('YY') or ''
  '*** FmdBChr           valid PMCd('BCHR') or ''
  '*** FmdBNo            FmdBNo= valid BNo from Bag(BYy= FdBYy, BChr= FdBChr, BCls= 'N') or 0
  
' **** Zubin 211 **** '
  Dim ws_FGBNo() As String, ws_BagNoBarCd As String
  Dim ws_OldYy As String, ws_OldChr As String, ws_OldNo As Single
' **** Zubin 211 **** '

  With GrdFmd
    Select Case ColName
    Case Is = UCase("FmdBYy")
      Cancel = (NewValue <> "") And (Not moCn.RecSeek("Select PMCd From Param Where PTyp= 'YY' And " + _
                                     "PMCd = '" + NewValue + "'"))
      ErrMsg = "Invalid Bag Year": Exit Sub
    Case Is = UCase("FmdBChr")
      Cancel = (NewValue <> "") And (Not moCn.RecSeek("Select PMCd From Param Where PTyp= 'BCHR' And " + _
                                     "PMCd = '" + NewValue + "'"))
      ErrMsg = "Invalid Bag Character": Exit Sub
' **** Zubin 211 **** '
    Case Is = UCase("wFGBNo")
      If NewValue <> "" Then
        'uni.156 calling function to ignore CoCd if entered
        ws_BagNoBarCd = ParseBagNoBarCd(NewValue)
        ws_FGBNo = Split(ws_BagNoBarCd, "/")
        If UBound(ws_FGBNo) <> 2 Then Cancel = True: ErrMsg = "Invalid FG Bag No": Exit Sub
        If Not IsNumeric(ws_FGBNo(2)) Then Cancel = True: ErrMsg = "Invalid FG Bag No": Exit Sub
'        Cancel = (NewValue <> 0) And (Not moCn.RecSeek("Select BNo From Bag Where " + _
'                                      "BCoCd= '" + ADC("FmCoCd") + "' and " + _
'                                      "BYy= '" + ws_FGBNo(0) + "' And " + _
'                                      "BChr = '" + ws_FGBNo(1) + "' And " + _
'                                      "BNo= " + CStr(ws_FGBNo(2)) + " and BCls='N' and BLoc = '" + ctFgLoc + "'"))
'        If Cancel = True Then ErrMsg = "Invalid Bag Number Or Bag Closed Or Bag not in PFG ": Exit Sub

        ' **** Focus is not set on ToFg-SubLoc because Help of ToFg-SubLoc is not seen
'        .COL = 6
        
        ws_OldYy = .Value(RowNum, "FMDBYY")
        ws_OldChr = .Value(RowNum, "FMDBCHR")
        ws_OldNo = .Value(RowNum, "FMDBNO")
        
        .Value(RowNum, "FMDBYY") = ws_FGBNo(0)
        Call GrdFmd_ColChange(RowNum, "FMDBYY", ws_FGBNo(0), Cancel, "")
        If Cancel = True Then ErrMsg = "Invalid Bag Year": .Value(RowNum, "FMDBYY") = ws_OldYy: Exit Sub
        
        .Value(RowNum, "FMDBCHR") = ws_FGBNo(1)
        Call GrdFmd_ColChange(RowNum, "FMDBCHR", ws_FGBNo(1), Cancel, "")
        If Cancel = True Then ErrMsg = "Invalid Bag Character / Flute Bags Cannot Be Taken To FG": .Value(RowNum, "FMDBCHR") = ws_OldChr: Exit Sub
        
        .Value(RowNum, "FMDBNO") = ws_FGBNo(2)
        Call GrdFmd_ColChange(RowNum, "FMDBNO", ws_FGBNo(2), Cancel, "")
        If Cancel = True Then ErrMsg = "Invalid Bag Number": .Value(RowNum, "FMDBNO") = ws_OldNo: Exit Sub

        ' ***** Manali 3.10.0 - 12/04/12 - Bar Code
        Dim wOldToSubLoc As String
        wOldToSubLoc = .Value(RowNum, "FmdToFgSubLoc")
        If ADC("wFmToFgSubLoc") <> "" Then
          .Value(RowNum, "FmdToFgSubLoc") = ADC("wFmToFgSubLoc")
          Call GrdFmd_ColChange(RowNum, "FmdToFgSubLoc", ADC("wFmToFgSubLoc"), Cancel, "")
          If Cancel = True Then ErrMsg = "Invalid To Fg Sub Loc": .Value(RowNum, "FmdToFgSubLoc") = wOldToSubLoc: Exit Sub
        End If
        
        If .Rows - 1 = RowNum Then
          .AddItem
          .Row = RowNum + 1
          .Col = 2
        End If
        ' ***** Manali 3.10.0 - 12/04/12 - Bar Code
      End If
' **** Zubin 211 **** '
    Case Is = UCase("FmdBNo")
      '****** Sachin 3.02 26-12-07 - Current Partition
      Cancel = (NewValue <> 0) And (Not moCn.RecSeek("Select BNo From Bag Where " + _
                                      "BCoCd= '" + ADC("FmCoCd") + "' and " + _
                                      "BYy= '" + .Value(RowNum, "FmdBYy") + "' And " + _
                                      "BChr = '" + .Value(RowNum, "FmdBChr") + "' And " + _
                                      "BNo= " + CStr(NewValue) + " and BPrtKey='" + ctCurrPrtn + "' and BCls='N' and BLoc = '" + ctFgLoc + "'"))
       ErrMsg = "Invalid Bag Number Or Bag Closed Or Bag not in PFG ": Exit Sub
       
       
       If NewValue <> 0 Then
        Dim wBSubLoc As String
          wBSubLoc = moCn.GetFldVal("Select BFgSubLoc from Bag Where " + _
                                      "BCoCd= '" + ADC("FmCoCd") + "' and " + _
                                      "BYy= '" + .Value(RowNum, "FmdBYy") + "' And " + _
                                      "BChr = '" + .Value(RowNum, "FmdBChr") + "' And " + _
                                      "BNo= " + CStr(NewValue) + " and BPrtKey='" + ctCurrPrtn + "'")
          Cancel = Not moCn.RecSeek("select 'x' From vParam where vPTyp='FgSubLoc' and vPCoCd= '" + gs_CoCd + "' and vPMCd='" + wBSubLoc + "'" + _
                              " and (vPValue4='' Or ','+vPValue4+',' Like '%," + gs_UsrCd + ",%')")
        If Cancel = True Then ErrMsg = "Invalid Bag Number As Bag Sub Loc. is Not Accessible to the User": Exit Sub
       End If
       
    Case Is = UCase("FmdToFgSubLoc")
       Cancel = (NewValue <> "") And (Not moCn.RecSeek("Select vPMCd From vParam Where " + _
                "vPCoCd= '" + ADC("FmCoCd") + "' and vPTyp= 'FGSUBLOC' And " + _
                "vPMCd = '" + NewValue + "'"))
      ErrMsg = "Invalid To FG Sub Location": Exit Sub
    End Select
  End With
End Sub
Private Sub GrdFmd_InitKey(ByVal RowNum As Integer)
  '*** Initialise Key Fields and all the fields which are not included in the grid
  
  Dim wFdKey As Single, wFdYy As String
  With GrdFmd
    If .IsNew(RowNum) Then
      .Store "FmdCoCd", ADC("FmCoCd")
      .Store "FmdTc", ADC("FmTc")
      .Store "FmdYy", ADC("FmYy")
      .Store "FmdChr", ADC("FmChr")
      .Store "FmdNo", ADC("FmNo")
      
      '****** Sachin 3.02 26-12-07 - Current partition only
      wFdYy = moCn.GetFldVal("Select max(FdYy) from Fgd where FdCoCd = '" + ADC("FmCoCd") + "' and " + _
                             "FdBYy= '" + .Value(RowNum, "FmdBYy") + "' and " + _
                             "FdBChr= '" + .Value(RowNum, "FmdBchr") + "' and " + _
                             "FdBNo= " + CStr(.Value(RowNum, "FmdBNo")) + " and FdPrtKey='" + ctCurrPrtn + "' ")
      wFdKey = moCn.GetFldVal("Select max(FdKey) from Fgd where FdCoCd = '" + ADC("FmCoCd") + "' and " + _
                             "FdBYy= '" + .Value(RowNum, "FmdBYy") + "' and " + _
                             "FdBChr= '" + .Value(RowNum, "FmdBchr") + "' and " + _
                             "FdBNo= " + CStr(.Value(RowNum, "FmdBNo")) + "and FdPrtKey='" + ctCurrPrtn + "' and " + _
                             "FdYy= '" + wFdYy + "'")
      .Store "FmdFdKey", wFdKey
      .Store "FmdFdYy", wFdYy
      
      '****** Sachin 3.02 - Default values for new fields ******
      .Store "FmdFmIdNo", moCn.GetFldVal("Select FmIdNo From Fm Where " + _
                                " FmCoCd='" + ADC("FmCoCd") + "' And FmTc='" + ADC("FmTc") + "' And FmYy='" + ADC("FmYy") + "' " + _
                                " And FmChr='" + ADC("FmChr") + "' And FmNo=" + CStr(ADC("FmNo")) + " And FmPrtKey='" + ctCurrPrtn + "' ")
      
    End If
  End With
End Sub
Private Sub GrdFmd_Valid()
  'CmdCP.Enabled = False: CmdEX.Enabled = False
End Sub
Private Sub ADC_SetRecSource()
  '*** Set the RecSource for ADC (Table Fg)
  '****** Sachin 3.02 26-12-07 - Corresponding Partition
  ADC.RecSource = " Select * from Fm " + _
                  " Where FmCoCd= '" + ADC("FmCoCd") + "' " + _
                  " And FmTc= '" + ADC("FmTc") + "' " + _
                  " And FmYy= '" + ADC("FmYy") + "' " + _
                  " And FmChr= '" + ADC("FmChr") + "' " + _
                  " And FmNo= " + CStr(ADC("FmNo")) + _
                  " And FmPrtKey=(Case When '" + gs_Partition + "'='" + ctCurrPrtn + "' Then '" + ctCurrPrtn + "' Else FmPrtKey End)" + _
                  IIF(gs_Partition = ctCurrPrtn, " And Fmdt>'" + ms_PurgeDt + "' ", "")
End Sub

Private Sub GrdFmd_SetRecSource()
  '*** Set the RecSource for GrdFmd (Table Fmd)
  GrdFmd.RecSource = " Select * from Fmd " + _
                     " Where FmdCoCd= '" + ADC("FmCoCd") + "' " + _
                     " And FmdTc= '" + ADC("FmTc") + "' " + _
                     " And FmdYy= '" + ADC("FmYy") + "' " + _
                     " And FmdChr= '" + ADC("FmChr") + "' " + _
                     " And FmdNo= " + CStr(ADC("FmNo")) + _
                     IIF(gs_Partition = ctCurrPrtn, " And FmdPrtKey ='" + ctCurrPrtn + "' ", "") + _
                     " Order By FmdSr"
End Sub
Private Sub adc_AutoGen()
  '*** write code for AutoGenerating the Finished Goods Voucher Number
  If ADC("FmNo") = 0 Then
    ADC("FmNo") = moCn.AutoGen("FM", "FmNo", ADC("FmCoCd") + ADC("FmTc") + ADC("FmYy") + ADC("FmChr"))
  Else
    Call moCn.AutoGen("FM", "FmNo", ADC("FmCoCd") + ADC("FmTc") + ADC("FmYy") + ADC("FmChr"), ADC("FmNo"))
  End If
End Sub
Private Sub HideAllFras()
  '*** This procedure hides all the frames other than FraNKeyAll
  FraNKeyAll.Visible = True
  ' ****** Manali 3.5.0 - 08/11/08 - Copy Bags Frame Added
  FraCB.Visible = False
  FraCB.Enabled = False
  ' ****** Manali 3.5.0 - 08/11/08 - Copy Bags Frame Added
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
  ws_BtnStr = ""
  Call SetObjColors(Me, ws_BtnStr)
  '*** Jenny Colour
  
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
Private Sub GrdFmd_Validate(Cancel As Boolean)
  Cancel = GrdFmd.Validate
End Sub

Private Sub EnaDisaCmds(ByVal pv_ShowFra As Boolean, Optional ByVal pv_CmdBut As MWCTL_BTN1, Optional ByVal pv_FmFra As en_FmFra)
  '*** This procedure is used to give the toggle option (enable and disable) to
      'the command buttons
  ' ****** Manali 3.5.0 - 08/11/08 - Copy Bags Frame Added
  If Not pv_CmdBut Is Nothing Then
    If pv_ShowFra = True Then
      pv_CmdBut.CausesValidation = True
    Else
    End If
  End If
  FraNKeyAll.Enabled = Not pv_ShowFra
  ADC.AllowSave = Not pv_ShowFra
  CmdCB.Enabled = Not pv_ShowFra
  CmdCpyRfId.Enabled = Not pv_ShowFra
  
  If Not pv_CmdBut Is Nothing Then pv_CmdBut.Enabled = True
  ' ****** Manali 3.5.0 - 08/11/08 - Copy Bags Frame Added
End Sub

Private Sub DispFra(ByVal pv_FmFra As en_FmFra)
  ' ****** Manali 3.5.0 - 08/11/08 - Copy Bags Frame Added
  Select Case UCase(pv_FmFra)
  Case Is = CpyBag
    If FraCB.Visible = True Then
      FraCB.Visible = False
      FraCB.Enabled = False
      Call EnaDisaCmds(False, CmdCB, CpyBag)
      CmdCB.SetFocus
    Else
      FraCB.Visible = True
      FraCB.Enabled = True
      ADC("WCBFGSUBLOCFR").SetFocus
      FraCB.ZOrder
      Call EnaDisaCmds(True, CmdCB)
    End If
    Case Is = CpyRfId
    If FraCpyRfId.Visible = True Then
      FraCpyRfId.Visible = False
      FraCpyRfId.Enabled = False
      Call EnaDisaCmds(False, CmdCpyRfId, CpyRfId)
      CmdCpyRfId.SetFocus
    Else
      FraCpyRfId.Visible = True
      FraCpyRfId.Enabled = True
      'vk.4 setfocus changed to Bag no./bag alias no. entry text box
      'ADC("WCPYTHRUBAGNO").SetFocus
      ADC("wRfIdSel").SetFocus
      FraCpyRfId.ZOrder
      Call EnaDisaCmds(True, CmdCpyRfId)
    End If
  End Select
  ' ****** Manali 3.5.0 - 08/11/08 - Copy Bags Frame Added
End Sub
  
'****** Manali 3.5.0 - 08/11/08 - Copy Bags Frame Added
Private Sub CopyBag()
  Dim wRsCB As MwfLib.MDORowSet, wTotRows As Single, wRow As Single, wCnd As String
  Dim wRfIdStr As String, wRfIdNoArr() As String, wRfIdNo As String, i As Integer, wTotRfid As Integer
  Dim wMinSr As Integer
  If Not pb_CpyRfId Then
    
    If ADC("wCbFgSubLocFr") = "" Or ADC("wCbFgSubLocTo") = "" Then DispMsg "Enter From / To Fg Sub Location", etError: Exit Sub
    
    If ADC("wCbFgSubLocFr") = ADC("wCbFgSubLocTo") Then DispMsg "From Sub Location and To Sub Location Cannot be Same", etError: Exit Sub
    
    ADC("wCbBYyFr").CmpStr = "BYy>=": ADC("wCbBChrFr").CmpStr = "BChr>=": ADC("wCbBNoFr").CmpStr = "BNo>="
    ADC("wCbBYyTo").CmpStr = "BYy<=": ADC("wCbBChrTo").CmpStr = "BChr<=": ADC("wCbBNoTo").CmpStr = "BNo<="
    ADC("wCbOdTcFr").CmpStr = "BOdTc>=": ADC("wCbOdYyFr").CmpStr = "BOdYy>=": ADC("wCbOdChrFr").CmpStr = "BOdChr>=": ADC("wCbOdNoFr").CmpStr = "BOdNo>=": ADC("wCbOdSrFr").CmpStr = "BOdSr>="
    ADC("wCbOdTcTo").CmpStr = "BOdTc<=": ADC("wCbOdYyTo").CmpStr = "BOdYy<=": ADC("wCbOdChrTo").CmpStr = "BOdChr<=": ADC("wCbOdNoTo").CmpStr = "BOdNo<=": ADC("wCbOdSrTo").CmpStr = "BOdSr<="
    ADC("wCbDmCdFr").CmpStr = "BOdDmCd>=": ADC("wCbDmCdTo").CmpStr = "BOdDmCd<="
    wCnd = ADC.RepCond
   Else
        If ADC("WFMTOFGSUBLOC") = "" Then DispMsg "Enter To Fg Sub Location", etError: Exit Sub
        'vk.4 RfId changed as Bag Alias No.
        If ADC("wRfIdSel") = "" Then DispMsg "Enter Bag Alias No. / Bag No List", etError: Exit Sub
        
         'Dim wRfIdStr As String, wRfIdNoArr() As String, wRfIdNo As String, i As Integer
         Dim wInvalRfIdStr As String
         wInvalRfIdStr = ""
         'vk.4 calling function to replace enter with ,
         wRfIdStr = PrepareRfIdString(ADC("wRfIdSel"))
         
    If ADC("wRfIdSel") <> "" Then
        wRfIdStr = ParseBagNoBarCd(wRfIdStr)

        wRfIdNoArr = Split(wRfIdStr, ",")
  
      If Left(wRfIdStr, 1) = "," Or Right(wRfIdStr, 1) = "," Then
          'vk.4 Tag changed as Bag Alias No.
          DispMsg "Invalid Bag Alias No. / Bag Input", etError: Exit Sub
      End If
      wRfIdStr = ""
      For i = 0 To UBound(wRfIdNoArr)
        wRfIdNo = wRfIdNoArr(i)
          '****Bhavna added bag no seln
        If ADC("WCPYTHRUBAGNO") = "N" Then
              'vk.4 Tag changed as Bag Alias No.
              If Not IsNumeric(wRfIdNo) Then DispMsg "Bag Alias No. Not Numeric", etError: Exit Sub
              If Not moCn.RecSeek("Select 'x' from RfIdMst Where RimNo=" + wRfIdNo + " ") Then
                    wInvalRfIdStr = wInvalRfIdStr + IIF(wInvalRfIdStr = "", "", ",") + wRfIdNo
              End If
        Else
            If Not moCn.RecSeek("Select 'x' from Bag Where Byy+'/'+BChr+'/'+ltrim(RTrim(str(BNo))) = '" + wRfIdNo + "'") Then
                wInvalRfIdStr = wInvalRfIdStr + IIF(wInvalRfIdStr = "", "", ",") + wRfIdNo
            End If
         End If
        
        wRfIdStr = wRfIdStr + IIF(wRfIdStr = "", "", ",") + wRfIdNo
      Next
      wTotRfid = UBound(wRfIdNoArr) + 1
      'vk.4 Tag changed as Bag Alias No.
      If wInvalRfIdStr <> "" Then DispMsg "Bag Alias No. / Bag No Not Defined:" + wInvalRfIdStr, etError: Exit Sub
      'vk.4 ADC("wRfIdSel") replaced with wRfIdStr
      wCnd = IIF(LTrim(wRfIdStr) <> "" And ADC("WCPYTHRUBAGNO") = "N", " Exists(Select 'x' From RfIdMst where RimNo in (" + wRfIdStr + ") and BCoCd=RimBCoCd and BYy=RimBYy and BChr=RimBChr and BNo=RimBNo)", wCnd)
      'vk.4 ADC("wRfIdSel") replaced with wRfIdStr
      wCnd = IIF(LTrim(wRfIdStr) <> "" And ADC("WCPYTHRUBAGNO") = "Y", " charindex(','+BYy + '/' + BChr + '/' + LTrim(RTrim(str(BNo)))+',' , '," + wRfIdStr + ",') <> 0 ", wCnd)
'      wCnd = " Exists(Select 'x' From RfIdMst where RimNo in (" + adc("wRfIdSel") + ") and BCoCd=RimBCoCd and BYy=RimBYy and BChr=RimBChr and BNo=RimBNo)"
      End If
   End If
  wCnd = IIF(wCnd <> "", " and " + wCnd, "")
  
  Set wRsCB = moCn.OpenRes("Select * from Bag where BCoCd= '" + ADC("FmCoCd") + "' and BCls= 'N' and " + _
                           "BLoc= '" + ctFgLoc + "' and BPrtKey='" + ctCurrPrtn + "' " + _
                           "and BFgSubLoc=" + IIF(pb_CpyRfId, "BFgSubLoc", "'" + ADC("wCbFgSubLocFr") + "' ") + _
                           wCnd + " Order By BCoCd, BYy, BChr, BNo")
                           
                           
  wTotRows = wRsCB.RecCount
  If wTotRows <= 0 Then DispMsg "No Bags Found for Copy in given Scope.", etError: GoTo CopyEnd
  GrdFmd.StartCopy
  MWLib.BeginProcess Me, "Copying Bags  ..."
  With GrdFmd
    wMinSr = .MaxVal("FmdSr") + 1
    Do While Not (wRsCB.EOF Or wRsCB.BOF)
      
      .AddItem
      wRow = .Rows - 1
      .Row = wRow
'      .Value(wRow, "FmdSr") = .MaxVal("FmdSr") + 1
      .Value(wRow, "FmdBYy") = wRsCB!BYy
      .Value(wRow, "FmdBChr") = wRsCB!BChr
      .Value(wRow, "FmdBNo") = wRsCB!BNo
      .Value(wRow, "FmdFrFgSubLoc") = IIF(pb_CpyRfId, wRsCB!BFgSubLoc, ADC("wCbFgSubLocFr"))
      .Value(wRow, "FmdToFgSubLoc") = IIF(pb_CpyRfId, ADC("WFMTOFGSUBLOC"), ADC("wCbFgSubLocTo"))
      .SaveRec (wRow)
      wRsCB.MoveNext
    Loop
    
  MWLib.EndProcess Me
  GrdFmd.EndCopy
    
    If pb_CpyRfId And ADC("wCpyThruBagNo") = "N" Then
    'vk.4 RfId changed as Bag Alias No.
        If ADC("wRfIdSel") = "" Then DispMsg "Enter Bag Alias No. List", etError: Exit Sub
        Dim wRsRfId As MwfLib.MDORowSet
        'vk.4 ADC("wRfIdSel") replaced with wRfIdStr
        Set wRsRfId = moCn.OpenRes("Select RimNo from RfIdMst where Not Exists(Select 'x' From Fmd where " + _
                                        " FmdCoCd=RimBCoCd and FmdBYy=RimBYY and FmdBChr=RimBChr and FmdBNo=RimBNo and " + _
                                        " FmdCoCd='" + ADC("FmCoCd") + "' and FmdTc='" + ADC("FmTc") + "' and FmdYy='" + ADC("FmYy") + "' and FmdChr='" + ADC("FmChr") + "' and FmdNo=" + CStr(ADC("FmNo")) + _
                                        " and  FmdSr>=" + CStr(wMinSr) + " and FmdSr<=" + CStr(.MaxVal("FmdSr")) + ")" + _
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
 
CopyEnd:
  Set wRsCB = Nothing

End Sub
' ****** Manali 3.5.0 - 08/11/08 - Copy Bags Frame Added
Private Sub GrdFmd_When()
       If Not mf_DtLocked Then GrdFmd.AllowAdd = True: GrdFmd.AllowDelete = True
End Sub
'vk.4 Copy from Bag-Y , copy from Bag alias no. -N
Private Sub OptCpyBagFr_Click(Index As Integer)
  Select Case Index
  Case Is = 0
    ADC("wCpyThruBagNo") = "Y"
  Case Is = 1
    ADC("wCpyThruBagNo") = "N"
  End Select
End Sub
'vk.4 tooltip for copy bag options
Private Sub OptCpyBagFr_GotFocus(Index As Integer)
  DispMsg OptCpyBagFr(Index).ToolTipText, etInfo
End Sub

