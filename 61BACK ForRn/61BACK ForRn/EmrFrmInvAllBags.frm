VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "tabctl32.ocx"
Object = "{AB2CC32A-E652-42A2-B341-26ECE02F6386}#2.1#0"; "mwfctlforspeed9.ocx"
Begin VB.Form EmrFrmInvAllBags 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Invoice Bag Entry"
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
   Begin MwfCtl.MWCTL_BTN1 CmdDelAllBags 
      Height          =   435
      Left            =   3720
      TabIndex        =   174
      ToolTipText     =   "Delete all Bags"
      Top             =   9330
      Width           =   1005
      _ExtentX        =   1773
      _ExtentY        =   767
      BackColor       =   -2147483633
      MaskColor       =   8454143
      Caption         =   "Dele&te Bags"
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
   Begin MwfCtl.MWCTL_BTN1 CmdCpyRfId 
      Height          =   465
      Left            =   7395
      TabIndex        =   147
      ToolTipText     =   "Copy Option"
      Top             =   9300
      Width           =   1560
      _ExtentX        =   2752
      _ExtentY        =   820
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
   Begin MwfCtl.MWCTL_BTN1 CmdSC 
      Height          =   465
      Left            =   4800
      TabIndex        =   10
      ToolTipText     =   "Option To Select The Scope Options For The Entry"
      Top             =   9300
      Width           =   825
      _ExtentX        =   1455
      _ExtentY        =   820
      BackColor       =   -2147483633
      MaskColor       =   8454143
      Caption         =   "Sco&pe"
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
      Height          =   465
      Left            =   6450
      TabIndex        =   12
      ToolTipText     =   "Option To See The Summary Details"
      Top             =   9300
      Width           =   945
      _ExtentX        =   1667
      _ExtentY        =   820
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
   Begin MwfCtl.MWCTL_BTN1 CmdCP 
      Height          =   465
      Left            =   5625
      TabIndex        =   11
      ToolTipText     =   "Option To Copy The Rates From Another Customer"
      Top             =   9300
      Width           =   825
      _ExtentX        =   1455
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
   Begin MwfCtl.MWCTL_HLP MWCTL_HLP1 
      Height          =   405
      Left            =   8580
      TabIndex        =   48
      TabStop         =   0   'False
      Top             =   9330
      Visible         =   0   'False
      Width           =   525
      _ExtentX        =   926
      _ExtentY        =   714
   End
   Begin MwfCtl.MWCTL_ADC ADC 
      CausesValidation=   0   'False
      Height          =   555
      Left            =   30
      TabIndex        =   13
      Top             =   9300
      Width           =   12705
      _ExtentX        =   22410
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
      Height          =   300
      Left            =   -36
      TabIndex        =   50
      Top             =   24
      Width           =   6720
      Begin MwfCtl.MWCTL_MED ATXT 
         Height          =   285
         Index           =   5
         Left            =   5115
         TabIndex        =   4
         ToolTipText     =   "Enter Sort Option"
         Top             =   0
         Width           =   1380
         _ExtentX        =   2434
         _ExtentY        =   503
         MaxLength       =   8
         DataType        =   4
         IdName          =   "WIFSRT"
      End
      Begin MwfCtl.MWCTL_MED ATXT 
         Height          =   285
         Index           =   40
         Left            =   1980
         TabIndex        =   1
         ToolTipText     =   "Enter Invoice Voucher Year"
         Top             =   0
         Width           =   360
         _ExtentX        =   635
         _ExtentY        =   503
         MaxLength       =   2
         DataType        =   4
         IdName          =   "WIFYY"
      End
      Begin MwfCtl.MWCTL_MED ATXT 
         Height          =   285
         Index           =   42
         Left            =   3240
         TabIndex        =   3
         ToolTipText     =   "Enter Invoice Voucher No."
         Top             =   0
         Width           =   780
         _ExtentX        =   1376
         _ExtentY        =   503
         Alignment       =   2
         Mask            =   "#####0"
         MaxLength       =   6
         DataType        =   1
         IdName          =   "WIFNO"
      End
      Begin MwfCtl.MWCTL_MED ATXT 
         Height          =   285
         Index           =   1
         Left            =   1245
         TabIndex        =   0
         ToolTipText     =   "Enter Invoice Voucher Tc"
         Top             =   0
         Width           =   600
         _ExtentX        =   1058
         _ExtentY        =   503
         MaxLength       =   3
         DataType        =   4
         ReCalcParent    =   "WIFCHR"
         IdName          =   "WIFTC"
      End
      Begin MwfCtl.MWCTL_MED ATXT 
         Height          =   285
         Index           =   0
         Left            =   2490
         TabIndex        =   2
         ToolTipText     =   "Enter Invoice Voucher Character"
         Top             =   0
         Width           =   600
         _ExtentX        =   1058
         _ExtentY        =   503
         MaxLength       =   3
         DataType        =   4
         ReCalcOn        =   "WIFTC"
         IdName          =   "WIFCHR"
      End
      Begin MwfCtl.MWCTL_MED ATXT 
         Height          =   285
         Index           =   61
         Left            =   3960
         TabIndex        =   141
         ToolTipText     =   "Invoice Date"
         Top             =   0
         Visible         =   0   'False
         Width           =   340
         _ExtentX        =   609
         _ExtentY        =   503
         Enabled         =   0   'False
         IdName          =   "WIFPRTKEY"
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
         Left            =   3135
         TabIndex        =   59
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
         Left            =   2385
         TabIndex        =   58
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
         Left            =   1870
         TabIndex        =   57
         Top             =   0
         Width           =   105
      End
      Begin VB.Label ALBL 
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
         Height          =   285
         Index           =   0
         Left            =   60
         TabIndex        =   56
         Top             =   0
         Width           =   1215
      End
      Begin VB.Label ALBL 
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
         Left            =   4335
         TabIndex        =   52
         Top             =   0
         Width           =   825
      End
   End
   Begin VB.Frame FRA_NKEYS 
      BorderStyle     =   0  'None
      Enabled         =   0   'False
      Height          =   9810
      Left            =   -120
      TabIndex        =   49
      Top             =   0
      Width           =   15315
      Begin VB.Frame FraCP 
         Height          =   4335
         Left            =   3630
         TabIndex        =   15
         Top             =   4830
         Visible         =   0   'False
         Width           =   5100
         Begin MwfCtl.MWCTL_BTN1 CmdCpGo 
            Height          =   285
            Left            =   4320
            TabIndex        =   47
            ToolTipText     =   "Press This Button To Copy"
            Top             =   3600
            Width           =   585
            _ExtentX        =   1032
            _ExtentY        =   503
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
            Index           =   12
            Left            =   1605
            TabIndex        =   16
            ToolTipText     =   "Enter Production Customer"
            Top             =   450
            Width           =   1320
            _ExtentX        =   2328
            _ExtentY        =   503
            MaxLength       =   8
            DataType        =   4
            IdName          =   "WCPPRDCMCD"
         End
         Begin MwfCtl.MWCTL_MED ATXT 
            Height          =   285
            Index           =   13
            Left            =   1605
            TabIndex        =   40
            ToolTipText     =   "Enter From Finished Goods Date (Can Be Blank)"
            Top             =   3315
            Width           =   855
            _ExtentX        =   1508
            _ExtentY        =   503
            MaxLength       =   8
            DataType        =   3
            IdName          =   "WCPFGDTFR"
         End
         Begin MwfCtl.MWCTL_MED ATXT 
            Height          =   285
            Index           =   15
            Left            =   3060
            TabIndex        =   41
            ToolTipText     =   "Enter To Finished Goods Date (Can Be Blank)"
            Top             =   3315
            Width           =   885
            _ExtentX        =   1561
            _ExtentY        =   503
            MaxLength       =   8
            DataType        =   3
            IdName          =   "WCPFGDTTO"
         End
         Begin MwfCtl.MWCTL_MED ATXT 
            Height          =   285
            Index           =   2
            Left            =   2310
            TabIndex        =   20
            ToolTipText     =   "Enter Production Sales Order Year"
            Top             =   1305
            Width           =   360
            _ExtentX        =   635
            _ExtentY        =   503
            MaxLength       =   2
            DataType        =   4
            IdName          =   "WCPPRDODYY"
         End
         Begin MwfCtl.MWCTL_MED ATXT 
            Height          =   285
            Index           =   3
            Left            =   3480
            TabIndex        =   22
            ToolTipText     =   "Enter Production Sales Order Number"
            Top             =   1305
            Width           =   780
            _ExtentX        =   1376
            _ExtentY        =   503
            Alignment       =   2
            Mask            =   "#####0"
            MaxLength       =   6
            DataType        =   1
            IdName          =   "WCPPRDODNO"
         End
         Begin MwfCtl.MWCTL_MED ATXT 
            Height          =   285
            Index           =   4
            Left            =   1605
            TabIndex        =   19
            ToolTipText     =   "Enter Production Sales Order Tc"
            Top             =   1305
            Width           =   600
            _ExtentX        =   1058
            _ExtentY        =   503
            MaxLength       =   3
            DataType        =   4
            IdName          =   "WCPPRDODTC"
         End
         Begin MwfCtl.MWCTL_MED ATXT 
            Height          =   285
            Index           =   6
            Left            =   2790
            TabIndex        =   21
            ToolTipText     =   "Enter Production Sales Order Character"
            Top             =   1305
            Width           =   600
            _ExtentX        =   1058
            _ExtentY        =   503
            MaxLength       =   3
            DataType        =   4
            IdName          =   "WCPPRDODCHR"
         End
         Begin MwfCtl.MWCTL_MED ATXT 
            Height          =   285
            Index           =   7
            Left            =   2790
            TabIndex        =   33
            ToolTipText     =   "Enter Packing List Number"
            Top             =   2460
            Width           =   930
            _ExtentX        =   1640
            _ExtentY        =   503
            Alignment       =   2
            Mask            =   "######0"
            MaxLength       =   7
            DataType        =   1
            IdName          =   "WCPPLNO"
         End
         Begin MwfCtl.MWCTL_MED ATXT 
            Height          =   285
            Index           =   8
            Left            =   2310
            TabIndex        =   32
            ToolTipText     =   "Enter Packing List Year"
            Top             =   2460
            Width           =   360
            _ExtentX        =   635
            _ExtentY        =   503
            MaxLength       =   2
            DataType        =   4
            IdName          =   "WCPPLYY"
         End
         Begin MwfCtl.MWCTL_MED ATXT 
            Height          =   285
            Index           =   14
            Left            =   1605
            TabIndex        =   31
            ToolTipText     =   "Enter Packing List Character"
            Top             =   2460
            Width           =   600
            _ExtentX        =   1058
            _ExtentY        =   503
            MaxLength       =   3
            DataType        =   4
            IdName          =   "WCPPLCHR"
         End
         Begin MwfCtl.MWCTL_MED ATXT 
            Height          =   285
            Index           =   9
            Left            =   4335
            TabIndex        =   23
            ToolTipText     =   "Enter Production Sales Order Serial Number"
            Top             =   1305
            Width           =   570
            _ExtentX        =   1005
            _ExtentY        =   503
            Alignment       =   2
            Mask            =   "###0"
            MaxLength       =   4
            DataType        =   1
            IdName          =   "WCPPRDODSR"
         End
         Begin MwfCtl.MWCTL_MED ATXT 
            Height          =   285
            Index           =   17
            Left            =   1605
            TabIndex        =   17
            ToolTipText     =   "Enter Export Customer"
            Top             =   735
            Width           =   1320
            _ExtentX        =   2328
            _ExtentY        =   503
            MaxLength       =   8
            DataType        =   4
            IdName          =   "WCPEXPCMCD"
         End
         Begin MwfCtl.MWCTL_MED ATXT 
            Height          =   285
            Index           =   18
            Left            =   2310
            TabIndex        =   25
            ToolTipText     =   "Enter Export Sales Order Year"
            Top             =   1590
            Width           =   360
            _ExtentX        =   635
            _ExtentY        =   503
            MaxLength       =   2
            DataType        =   4
            IdName          =   "WCPEXPODYY"
         End
         Begin MwfCtl.MWCTL_MED ATXT 
            Height          =   285
            Index           =   20
            Left            =   3480
            TabIndex        =   27
            ToolTipText     =   "Enter Export Sales Order Number"
            Top             =   1590
            Width           =   780
            _ExtentX        =   1376
            _ExtentY        =   503
            Alignment       =   2
            Mask            =   "#####0"
            MaxLength       =   6
            DataType        =   1
            IdName          =   "WCPEXPODNO"
         End
         Begin MwfCtl.MWCTL_MED ATXT 
            Height          =   285
            Index           =   21
            Left            =   1605
            TabIndex        =   24
            ToolTipText     =   "Enter Export Sales Order Tc"
            Top             =   1590
            Width           =   600
            _ExtentX        =   1058
            _ExtentY        =   503
            MaxLength       =   3
            DataType        =   4
            IdName          =   "WCPEXPODTC"
         End
         Begin MwfCtl.MWCTL_MED ATXT 
            Height          =   285
            Index           =   22
            Left            =   2790
            TabIndex        =   26
            ToolTipText     =   "Enter Export Sales Order Character"
            Top             =   1590
            Width           =   600
            _ExtentX        =   1058
            _ExtentY        =   503
            MaxLength       =   3
            DataType        =   4
            IdName          =   "WCPEXPODCHR"
         End
         Begin MwfCtl.MWCTL_MED ATXT 
            Height          =   285
            Index           =   24
            Left            =   4335
            TabIndex        =   28
            ToolTipText     =   "Enter Export Sales Order Serial Number"
            Top             =   1590
            Width           =   570
            _ExtentX        =   1005
            _ExtentY        =   503
            Alignment       =   2
            Mask            =   "###0"
            MaxLength       =   4
            DataType        =   1
            IdName          =   "WCPEXPODSR"
         End
         Begin MwfCtl.MWCTL_MED ATXT 
            Height          =   285
            Index           =   25
            Left            =   1605
            TabIndex        =   30
            ToolTipText     =   "Enter Design Code"
            Top             =   2175
            Width           =   2415
            _ExtentX        =   4260
            _ExtentY        =   503
            MaxLength       =   15
            DataType        =   4
            IdName          =   "WCPDMCD"
         End
         Begin MwfCtl.MWCTL_MED ATXT 
            Height          =   285
            Index           =   26
            Left            =   2310
            TabIndex        =   35
            ToolTipText     =   "Enter Invoice Year"
            Top             =   2745
            Width           =   360
            _ExtentX        =   635
            _ExtentY        =   503
            MaxLength       =   2
            DataType        =   4
            IdName          =   "WCPINYY"
         End
         Begin MwfCtl.MWCTL_MED ATXT 
            Height          =   285
            Index           =   29
            Left            =   3480
            TabIndex        =   37
            ToolTipText     =   "Enter Invoice Number"
            Top             =   2745
            Width           =   780
            _ExtentX        =   1376
            _ExtentY        =   503
            Alignment       =   2
            Mask            =   "#####0"
            MaxLength       =   6
            DataType        =   1
            IdName          =   "WCPINNO"
         End
         Begin MwfCtl.MWCTL_MED ATXT 
            Height          =   285
            Index           =   30
            Left            =   1605
            TabIndex        =   34
            ToolTipText     =   "Enter Invoice Tc"
            Top             =   2745
            Width           =   600
            _ExtentX        =   1058
            _ExtentY        =   503
            MaxLength       =   3
            DataType        =   4
            IdName          =   "WCPINTC"
         End
         Begin MwfCtl.MWCTL_MED ATXT 
            Height          =   285
            Index           =   31
            Left            =   2790
            TabIndex        =   36
            ToolTipText     =   "Enter Invoice Character"
            Top             =   2745
            Width           =   600
            _ExtentX        =   1058
            _ExtentY        =   503
            MaxLength       =   3
            DataType        =   4
            IdName          =   "WCPINCHR"
         End
         Begin MwfCtl.MWCTL_MED ATXT 
            Height          =   285
            Index           =   32
            Left            =   4335
            TabIndex        =   38
            ToolTipText     =   "Enter Invoice Serial Number"
            Top             =   2745
            Width           =   570
            _ExtentX        =   1005
            _ExtentY        =   503
            Alignment       =   2
            Mask            =   "###0"
            MaxLength       =   4
            DataType        =   1
            IdName          =   "WCPIDSR"
         End
         Begin MwfCtl.MWCTL_MED ATXT 
            Height          =   285
            Index           =   33
            Left            =   1605
            TabIndex        =   39
            ToolTipText     =   "Enter Invoice Export Number"
            Top             =   3030
            Width           =   3015
            _ExtentX        =   5318
            _ExtentY        =   503
            MaxLength       =   20
            DataType        =   4
            IdName          =   "WCPINEXPNO"
         End
         Begin MwfCtl.MWCTL_MED ATXT 
            Height          =   285
            Index           =   59
            Left            =   1605
            TabIndex        =   42
            ToolTipText     =   "Enter The Fg Sub Location Of The Bags"
            Top             =   3600
            Width           =   1320
            _ExtentX        =   2328
            _ExtentY        =   503
            MaxLength       =   8
            DataType        =   4
            IdName          =   "WCPFGSUBLOC"
         End
         Begin MwfCtl.MWCTL_MED ATXT 
            Height          =   285
            Index           =   60
            Left            =   1605
            TabIndex        =   18
            ToolTipText     =   "Enter Bag Contractor Code"
            Top             =   1020
            Width           =   1320
            _ExtentX        =   2328
            _ExtentY        =   503
            MaxLength       =   8
            DataType        =   4
            IdName          =   "WCPCTCD"
         End
         Begin MwfCtl.MWCTL_MED ATXT 
            Height          =   285
            Index           =   67
            Left            =   1605
            TabIndex        =   29
            ToolTipText     =   "Enter Design Category"
            Top             =   1890
            Width           =   2415
            _ExtentX        =   4260
            _ExtentY        =   503
            MaxLength       =   15
            DataType        =   4
            IdName          =   "WCPDMCTG"
         End
         Begin MwfCtl.MWCTL_MED ATXT 
            Height          =   285
            Index           =   72
            Left            =   2310
            TabIndex        =   44
            ToolTipText     =   "Enter Shipment Note Year"
            Top             =   3885
            Width           =   360
            _ExtentX        =   635
            _ExtentY        =   503
            MaxLength       =   2
            DataType        =   4
            IdName          =   "WCPSNYY"
         End
         Begin MwfCtl.MWCTL_MED ATXT 
            Height          =   285
            Index           =   73
            Left            =   3480
            TabIndex        =   46
            ToolTipText     =   "Enter Shipment Note Number"
            Top             =   3885
            Width           =   780
            _ExtentX        =   1376
            _ExtentY        =   503
            Alignment       =   2
            Mask            =   "#####0"
            MaxLength       =   6
            DataType        =   1
            IdName          =   "WCPSNNO"
         End
         Begin MwfCtl.MWCTL_MED ATXT 
            Height          =   285
            Index           =   74
            Left            =   1605
            TabIndex        =   43
            ToolTipText     =   "Enter Shipment Note Tc"
            Top             =   3885
            Width           =   600
            _ExtentX        =   1058
            _ExtentY        =   503
            MaxLength       =   3
            DataType        =   4
            IdName          =   "WCPSNTC"
         End
         Begin MwfCtl.MWCTL_MED ATXT 
            Height          =   285
            Index           =   77
            Left            =   2790
            TabIndex        =   45
            ToolTipText     =   "Enter Shipment Note Character"
            Top             =   3885
            Width           =   600
            _ExtentX        =   1058
            _ExtentY        =   503
            MaxLength       =   3
            DataType        =   4
            IdName          =   "WCPSNCHR"
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
            Index           =   67
            Left            =   3405
            TabIndex        =   184
            Top             =   3885
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
            Index           =   66
            Left            =   2685
            TabIndex        =   183
            Top             =   3885
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
            Index           =   65
            Left            =   2220
            TabIndex        =   182
            Top             =   3885
            Width           =   105
         End
         Begin VB.Label ALBL 
            BackStyle       =   0  'Transparent
            Caption         =   "Shipment Note"
            BeginProperty Font 
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
            Left            =   90
            TabIndex        =   181
            Top             =   3885
            Width           =   1455
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
            Index           =   36
            Left            =   90
            TabIndex        =   146
            Top             =   1890
            Width           =   1095
         End
         Begin VB.Label LblMain 
            BackStyle       =   0  'Transparent
            BorderStyle     =   1  'Fixed Single
            Caption         =   "  Copy From Finished Goods"
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
            TabIndex        =   140
            Top             =   90
            Width           =   4770
         End
         Begin VB.Label LblCpCtCd 
            BackStyle       =   0  'Transparent
            Caption         =   "Contractor"
            BeginProperty Font 
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
            TabIndex        =   138
            Top             =   1020
            Width           =   1095
         End
         Begin VB.Label ALBL 
            BackStyle       =   0  'Transparent
            Caption         =   "FgSubLoc"
            BeginProperty Font 
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
            TabIndex        =   137
            Top             =   3600
            Width           =   1095
         End
         Begin VB.Label ALBL 
            BackStyle       =   0  'Transparent
            Caption         =   "Inv Exp No"
            BeginProperty Font 
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
            TabIndex        =   91
            Top             =   3030
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
            Index           =   28
            Left            =   3405
            TabIndex        =   90
            Top             =   2745
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
            Index           =   27
            Left            =   2685
            TabIndex        =   89
            Top             =   2745
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
            Left            =   2220
            TabIndex        =   88
            Top             =   2745
            Width           =   105
         End
         Begin VB.Label ALBL 
            BackStyle       =   0  'Transparent
            Caption         =   "Invoice"
            BeginProperty Font 
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
            TabIndex        =   87
            Top             =   2745
            Width           =   1215
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
            Left            =   4260
            TabIndex        =   86
            Top             =   2745
            Width           =   105
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
            Index           =   23
            Left            =   90
            TabIndex        =   85
            Top             =   2175
            Width           =   1095
         End
         Begin VB.Label ALBL 
            BackStyle       =   0  'Transparent
            Caption         =   "Exp Cust"
            BeginProperty Font 
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
            TabIndex        =   84
            Top             =   735
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
            Index           =   20
            Left            =   3405
            TabIndex        =   83
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
            Index           =   19
            Left            =   2685
            TabIndex        =   82
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
            Index           =   18
            Left            =   2220
            TabIndex        =   81
            Top             =   1590
            Width           =   105
         End
         Begin VB.Label ALBL 
            BackStyle       =   0  'Transparent
            Caption         =   "Exp SO"
            BeginProperty Font 
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
            TabIndex        =   80
            Top             =   1590
            Width           =   1215
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
            Index           =   13
            Left            =   4260
            TabIndex        =   79
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
            Index           =   6
            Left            =   4260
            TabIndex        =   67
            Top             =   1305
            Width           =   105
         End
         Begin VB.Label ALBL 
            BackStyle       =   0  'Transparent
            Caption         =   "Pck List"
            BeginProperty Font 
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
            TabIndex        =   66
            Top             =   2460
            Width           =   1035
         End
         Begin VB.Label Label4 
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
            Left            =   2685
            TabIndex        =   65
            Top             =   2460
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
            Index           =   16
            Left            =   2220
            TabIndex        =   64
            Top             =   2460
            Width           =   135
         End
         Begin VB.Label ALBL 
            BackStyle       =   0  'Transparent
            Caption         =   "Prd SO"
            BeginProperty Font 
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
            Top             =   1305
            Width           =   1215
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
            Index           =   3
            Left            =   2220
            TabIndex        =   62
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
            Index           =   2
            Left            =   2685
            TabIndex        =   61
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
            Index           =   1
            Left            =   3405
            TabIndex        =   60
            Top             =   1305
            Width           =   105
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
            Index           =   12
            Left            =   2760
            TabIndex        =   55
            Top             =   3315
            Width           =   345
         End
         Begin VB.Label ALBL 
            BackStyle       =   0  'Transparent
            Caption         =   "Fg Dt  Fr"
            BeginProperty Font 
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
            TabIndex        =   54
            Top             =   3315
            Width           =   975
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
            Index           =   11
            Left            =   90
            TabIndex        =   53
            Top             =   450
            Width           =   1095
         End
      End
      Begin VB.Frame FraNKeyAll 
         BorderStyle     =   0  'None
         Height          =   9435
         Left            =   150
         TabIndex        =   51
         Top             =   30
         Width           =   15120
         Begin VB.Frame FraFt 
            Height          =   555
            Left            =   -15
            TabIndex        =   68
            Top             =   8580
            Width           =   12705
            Begin MwfCtl.MWCTL_MED ATXT 
               Height          =   285
               Index           =   10
               Left            =   4155
               TabIndex        =   69
               ToolTipText     =   "Production Customer"
               Top             =   180
               Width           =   1320
               _ExtentX        =   2328
               _ExtentY        =   503
               Enabled         =   0   'False
               MaxLength       =   8
               DataType        =   4
               IdName          =   "WIFPRDCMCD"
            End
            Begin MwfCtl.MWCTL_MED ATXT 
               Height          =   285
               Index           =   19
               Left            =   1365
               TabIndex        =   70
               ToolTipText     =   "Production Sales Order Voucher "
               Top             =   180
               Width           =   2700
               _ExtentX        =   4763
               _ExtentY        =   503
               Enabled         =   0   'False
               MaxLength       =   23
               DataType        =   4
               IdName          =   "WIFPRDORD"
            End
            Begin MwfCtl.MWCTL_MED ATXT 
               Height          =   285
               Index           =   23
               Left            =   7005
               TabIndex        =   71
               ToolTipText     =   "Total Number Of Bags"
               Top             =   180
               Width           =   780
               _ExtentX        =   1376
               _ExtentY        =   503
               Enabled         =   0   'False
               Mask            =   "#####0"
               MaxLength       =   6
               DataType        =   2
               ReCalcOn        =   "GrdInvFgd(IFFDBNO),GrdInvFgd(IFSRNO)"
               IdName          =   "WIFTOTBAGS"
            End
            Begin MwfCtl.MWCTL_MED ATXT 
               Height          =   285
               Index           =   11
               Left            =   9420
               TabIndex        =   75
               ToolTipText     =   "Total Bag Quantity"
               Top             =   180
               Width           =   1035
               _ExtentX        =   1826
               _ExtentY        =   503
               Alignment       =   2
               Enabled         =   0   'False
               Mask            =   "######0.0"
               MaxLength       =   9
               DataType        =   2
               ReCalcOn        =   "GrdInvFgd(IFQTY)"
               IdName          =   "WIFTOTQTY"
            End
            Begin MwfCtl.MWCTL_MED ATXT 
               Height          =   285
               Index           =   16
               Left            =   10290
               TabIndex        =   76
               ToolTipText     =   "Total Bag Weight"
               Top             =   180
               Width           =   1125
               _ExtentX        =   1984
               _ExtentY        =   503
               Alignment       =   2
               Enabled         =   0   'False
               Mask            =   "#####0.000"
               MaxLength       =   10
               DataType        =   2
               ReCalcOn        =   "GrdInvFgd(IFGRSWT)"
               IdName          =   "WIFTOTWT"
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
               Left            =   10215
               TabIndex        =   78
               Top             =   180
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
               Index           =   7
               Left            =   4080
               TabIndex        =   77
               Top             =   180
               Width           =   105
            End
            Begin VB.Label ALBL 
               BackStyle       =   0  'Transparent
               Caption         =   "Tot Qty/Wt"
               BeginProperty Font 
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
               Left            =   8280
               TabIndex        =   74
               Top             =   180
               Width           =   1185
            End
            Begin VB.Label ALBL 
               BackStyle       =   0  'Transparent
               Caption         =   "Prd SO/Cust"
               BeginProperty Font 
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
               TabIndex        =   73
               Top             =   180
               Width           =   1395
            End
            Begin VB.Label ALBL 
               BackStyle       =   0  'Transparent
               Caption         =   "Tot Bags"
               BeginProperty Font 
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
               Left            =   6030
               TabIndex        =   72
               Top             =   180
               Width           =   1005
            End
         End
         Begin MwfCtl.MWCTL_MED ATXT 
            Height          =   285
            Index           =   34
            Left            =   1185
            TabIndex        =   7
            ToolTipText     =   "Invoice Customer Code"
            Top             =   285
            Width           =   1365
            _ExtentX        =   2408
            _ExtentY        =   503
            Enabled         =   0   'False
            MaxLength       =   8
            DataType        =   4
            IdName          =   "WIFCMCD"
         End
         Begin MwfCtl.MWCTL_MED ATXT 
            Height          =   285
            Index           =   35
            Left            =   7275
            TabIndex        =   5
            ToolTipText     =   "Invoice Date"
            Top             =   0
            Width           =   885
            _ExtentX        =   1561
            _ExtentY        =   503
            Enabled         =   0   'False
            MaxLength       =   8
            DataType        =   3
            IdName          =   "WIFDT"
         End
         Begin MwfCtl.MWCTL_MED ATXT 
            Height          =   285
            Index           =   44
            Left            =   9525
            TabIndex        =   6
            ToolTipText     =   "LMG Sales Rate"
            Top             =   0
            Width           =   885
            _ExtentX        =   1561
            _ExtentY        =   503
            Alignment       =   2
            Enabled         =   0   'False
            Mask            =   "###0.00"
            MaxLength       =   7
            DataType        =   2
            IdName          =   "WIFLMGSAL"
         End
         Begin MwfCtl.MWCTL_MED ATXT 
            Height          =   285
            Index           =   36
            Left            =   5055
            TabIndex        =   8
            ToolTipText     =   "Invoice Export Number"
            Top             =   285
            Width           =   3105
            _ExtentX        =   5477
            _ExtentY        =   503
            Enabled         =   0   'False
            MaxLength       =   20
            DataType        =   4
            IdName          =   "WIFEXPNO"
         End
         Begin MwfCtl.MWCTL_MED ATXT 
            Height          =   285
            Index           =   58
            Left            =   9525
            TabIndex        =   135
            ToolTipText     =   "LMP Sales Rate"
            Top             =   285
            Width           =   885
            _ExtentX        =   1561
            _ExtentY        =   503
            Alignment       =   2
            Enabled         =   0   'False
            Mask            =   "###0.00"
            MaxLength       =   7
            DataType        =   2
            IdName          =   "WIFLMPSAL"
         End
         Begin MwfCtl.MWCTL_MED ATXT 
            Height          =   285
            Index           =   64
            Left            =   11790
            TabIndex        =   144
            ToolTipText     =   "LMS Sales Rate"
            Top             =   0
            Width           =   885
            _ExtentX        =   1561
            _ExtentY        =   503
            Alignment       =   2
            Enabled         =   0   'False
            Mask            =   "###0.00"
            MaxLength       =   7
            DataType        =   2
            IdName          =   "WIFLMSSAL"
         End
         Begin MwfCtl.MWCTL_MED ATXT 
            Height          =   285
            Index           =   66
            Left            =   11790
            TabIndex        =   145
            ToolTipText     =   "LML Sales Rate"
            Top             =   285
            Width           =   885
            _ExtentX        =   1561
            _ExtentY        =   503
            Alignment       =   2
            Enabled         =   0   'False
            Mask            =   "###0.00"
            MaxLength       =   7
            DataType        =   2
            IdName          =   "WIFLMLSAL"
         End
         Begin MwfCtl.MWCTL_FLX GrdInvFgd 
            Height          =   8040
            Left            =   0
            TabIndex        =   9
            Top             =   630
            Width           =   12675
            _ExtentX        =   22357
            _ExtentY        =   14182
            Cols            =   18
            colname1        =   "IFSRNO"
            heading1        =   "SrNo "
            datafld1        =   "IfSrNo"
            datatype1       =   1
            mask1           =   "####0"
            recalcparent1   =   "WIFTOTBAGS"
            maxlength1      =   5
            style1          =   2
            tooltiptext1    =   "Serial Number"
            colname2        =   "IFSR"
            heading2        =   "Inv Sr. "
            datafld2        =   "IFSR"
            datatype2       =   1
            mask2           =   "###0"
            maxlength2      =   4
            tooltiptext2    =   "Invoice Serial"
            colname3        =   "WIFFDBAGNO"
            datatype3       =   4
            maxlength3      =   18
            tooltiptext3    =   "Enter Bag Number Using Barcode Reader"
            colname4        =   "IFFDBYY"
            heading4        =   "BYy"
            datafld4        =   "IfFdBYy"
            datatype4       =   4
            recalcon4       =   "GrdInvFgd(IFFDTC),GrdInvFgd(IFFDYY),GrdInvFgd(IFFDCHR),GrdInvFgd(IFFDNO),GrdInvFgd(IFFDSR)"
            recalcparent4   =   "GrdInvFgd(IFFDTC),GrdInvFgd(IFFDYY),GrdInvFgd(IFFDCHR),GrdInvFgd(IFFDNO),GrdInvFgd(IFFDSR),GrdInvFgd(IFFDBNO),GrdInvFgd(IFBIDNO)"
            maxlength4      =   2
            tooltiptext4    =   "Enter Bag Year"
            colname5        =   "IFFDBCHR"
            heading5        =   "BChr"
            datafld5        =   "IfFdBChr"
            datatype5       =   4
            recalcon5       =   "GrdInvFgd(IFFDTC),GrdInvFgd(IFFDYY),GrdInvFgd(IFFDCHR),GrdInvFgd(IFFDNO),GrdInvFgd(IFFDSR)"
            recalcparent5   =   "GrdInvFgd(IFFDTC),GrdInvFgd(IFFDYY),GrdInvFgd(IFFDCHR),GrdInvFgd(IFFDNO),GrdInvFgd(IFFDSR),GrdInvFgd(IFFDBNO),GrdInvFgd(IFBIDNO)"
            maxlength5      =   3
            tooltiptext5    =   "Enter Bag Character"
            colname6        =   "IFFDBNO"
            heading6        =   "BNo       "
            datafld6        =   "IfFdBNo"
            datatype6       =   1
            mask6           =   "######0"
            recalcon6       =   "GrdInvFgd(IFFDBYY),GrdInvFgd(IFFDBCHR),GrdInvFgd(IFFDTC),GrdInvFgd(IFFDYY),GrdInvFgd(IFFDCHR),GrdInvFgd(IFFDNO),GrdInvFgd(IFFDSR)"
            recalcparent6   =   "GrdInvFgd(IFFDTC),GrdInvFgd(IFFDYY),GrdInvFgd(IFFDCHR),GrdInvFgd(IFFDNO),GrdInvFgd(IFFDSR),WIFTOTBAGS,GrdInvFgd(IFBIDNO)"
            maxlength6      =   7
            tooltiptext6    =   "Enter Bag Number"
            colname7        =   "IFFDTC"
            heading7        =   "FgTc"
            datafld7        =   "IfFdTc"
            datatype7       =   4
            recalcon7       =   "GrdInvFgd(IFFDBYY),GrdInvFgd(IFFDBCHR),GrdInvFgd(IFFDBNO)"
            recalcparent7   =   $"EmrFrmInvAllBags.frx":0000
            maxlength7      =   3
            tooltiptext7    =   "Enter Finished Goods Tc"
            colname8        =   "IFFDYY"
            heading8        =   "FgYy"
            datafld8        =   "IfFdYy"
            datatype8       =   4
            recalcon8       =   "GrdInvFgd(IFFDBYY),GrdInvFgd(IFFDBCHR),GrdInvFgd(IFFDBNO)"
            recalcparent8   =   $"EmrFrmInvAllBags.frx":00D1
            maxlength8      =   2
            tooltiptext8    =   "Enter Finished Goods Year"
            colname9        =   "IFFDCHR"
            heading9        =   "FgChr"
            datafld9        =   "IfFdChr"
            datatype9       =   4
            recalcon9       =   "GrdInvFgd(IFFDBYY),GrdInvFgd(IFFDBCHR),GrdInvFgd(IFFDBNO)"
            recalcparent9   =   $"EmrFrmInvAllBags.frx":01A2
            maxlength9      =   3
            tooltiptext9    =   "Enter Finished Goods Character"
            colname10       =   "IFFDNO"
            heading10       =   "FgNo   "
            datafld10       =   "IfFdNo"
            datatype10      =   1
            mask10          =   "#####0"
            recalcon10      =   "GrdInvFgd(IFFDBYY),GrdInvFgd(IFFDBCHR),GrdInvFgd(IFFDBNO),GrdInvFgd(IFFDTC),GrdInvFgd(IFFDYY),GrdInvFgd(IFFDCHR)"
            recalcparent10  =   $"EmrFrmInvAllBags.frx":0273
            maxlength10     =   6
            tooltiptext10   =   "Enter Finished Goods Number"
            colname11       =   "IFFDSR"
            heading11       =   "FgSr"
            datafld11       =   "IfFdSr"
            datatype11      =   1
            mask11          =   "###0"
            recalcon11      =   "GrdInvFgd(IFFDBYY),GrdInvFgd(IFFDBCHR),GrdInvFgd(IFFDBNO),GrdInvFgd(IFFDTC),GrdInvFgd(IFFDYY),GrdInvFgd(IFFDCHR),GrdInvFgd(IFFDNO)"
            recalcparent11  =   $"EmrFrmInvAllBags.frx":0332
            maxlength11     =   4
            tooltiptext11   =   "Enter Finished Goods Serial"
            colname12       =   "WIFDMCD"
            heading12       =   "Design Cd                     "
            datatype12      =   4
            recalcon12      =   "GrdInvFgd(IFFDTC),GrdInvFgd(IFFDYY),GrdInvFgd(IFFDCHR),GrdInvFgd(IFFDNO),GrdInvFgd(IFFDSR)"
            maxlength12     =   15
            tooltiptext12   =   "Design Code"
            colname13       =   "WIFSFX"
            heading13       =   "Suffix                 "
            datatype13      =   4
            recalcon13      =   "GrdInvFgd(IFFDTC),GrdInvFgd(IFFDYY),GrdInvFgd(IFFDCHR),GrdInvFgd(IFFDNO),GrdInvFgd(IFFDSR)"
            maxlength13     =   15
            tooltiptext13   =   "Design Suffix"
            colname14       =   "WIFDMSZ"
            heading14       =   "Size     "
            datatype14      =   4
            recalcon14      =   "GrdInvFgd(IFFDTC),GrdInvFgd(IFFDYY),GrdInvFgd(IFFDCHR),GrdInvFgd(IFFDNO),GrdInvFgd(IFFDSR)"
            maxlength14     =   5
            tooltiptext14   =   "Design Size"
            colname15       =   "IFQTY"
            heading15       =   "BQty "
            datafld15       =   "IfQty"
            datatype15      =   2
            mask15          =   "##0.0"
            recalcon15      =   "GrdInvFgd(IFFDTC),GrdInvFgd(IFFDYY),GrdInvFgd(IFFDCHR),GrdInvFgd(IFFDNO),GrdInvFgd(IFFDSR)"
            recalcparent15  =   "WIFTOTQTY"
            maxlength15     =   5
            style15         =   2
            tooltiptext15   =   "Bag Quantity"
            colname16       =   "IFGRSWT"
            heading16       =   "Gross Wt  "
            datafld16       =   "IfGrsWt"
            datatype16      =   2
            mask16          =   "####0.000"
            recalcon16      =   "GrdInvFgd(IFFDTC),GrdInvFgd(IFFDYY),GrdInvFgd(IFFDCHR),GrdInvFgd(IFFDNO),GrdInvFgd(IFFDSR)"
            recalcparent16  =   "WIFTOTWT"
            maxlength16     =   9
            style16         =   2
            tooltiptext16   =   "Gross Weight"
            colname17       =   "IFPRTKEY"
            heading17       =   "Prtn "
            datafld17       =   "IfPrtKey"
            maxlength17     =   1
            style17         =   0
         End
         Begin VB.Label LblIfLmsSal 
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
            Left            =   10650
            TabIndex        =   143
            Top             =   0
            Width           =   1185
         End
         Begin VB.Label LblIfLmlSal 
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
            Left            =   10650
            TabIndex        =   142
            Top             =   285
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
            Index           =   34
            Left            =   8400
            TabIndex        =   136
            Top             =   285
            Width           =   1185
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
            Index           =   32
            Left            =   0
            TabIndex        =   95
            Top             =   285
            Width           =   1005
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
            Left            =   6750
            TabIndex        =   94
            Top             =   0
            Width           =   585
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
            Left            =   8400
            TabIndex        =   93
            Top             =   0
            Width           =   1185
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
            Index           =   30
            Left            =   3990
            TabIndex        =   92
            Top             =   285
            Width           =   1125
         End
      End
      Begin VB.Frame FraSumm 
         Height          =   4335
         Left            =   150
         TabIndex        =   14
         Top             =   4920
         Width           =   13575
         Begin TabDlg.SSTab TabSumm 
            Height          =   4095
            Left            =   60
            TabIndex        =   152
            Top             =   180
            Width           =   13485
            _ExtentX        =   23786
            _ExtentY        =   7223
            _Version        =   393216
            Tabs            =   2
            Tab             =   1
            TabsPerRow      =   2
            TabHeight       =   520
            TabCaption(0)   =   "General Summary"
            TabPicture(0)   =   "EmrFrmInvAllBags.frx":03DF
            Tab(0).ControlEnabled=   0   'False
            Tab(0).Control(0)=   "GrdSumm"
            Tab(0).Control(1)=   "ATXT(27)"
            Tab(0).Control(2)=   "ATXT(28)"
            Tab(0).Control(3)=   "ATXT(63)"
            Tab(0).Control(4)=   "ATXT(65)"
            Tab(0).Control(5)=   "ALBL(62)"
            Tab(0).Control(6)=   "ALBL(61)"
            Tab(0).Control(7)=   "ALBL(33)"
            Tab(0).Control(8)=   "ALBL(31)"
            Tab(0).ControlCount=   9
            TabCaption(1)   =   "Raw Material Summary"
            TabPicture(1)   =   "EmrFrmInvAllBags.frx":03FB
            Tab(1).ControlEnabled=   -1  'True
            Tab(1).Control(0)=   "LblMain(3)"
            Tab(1).Control(0).Enabled=   0   'False
            Tab(1).Control(1)=   "LblMain(2)"
            Tab(1).Control(1).Enabled=   0   'False
            Tab(1).Control(2)=   "ALBL(64)"
            Tab(1).Control(2).Enabled=   0   'False
            Tab(1).Control(3)=   "ALBL(63)"
            Tab(1).Control(3).Enabled=   0   'False
            Tab(1).Control(4)=   "ATXT(71)"
            Tab(1).Control(4).Enabled=   0   'False
            Tab(1).Control(5)=   "ATXT(70)"
            Tab(1).Control(5).Enabled=   0   'False
            Tab(1).Control(6)=   "ATXT(69)"
            Tab(1).Control(6).Enabled=   0   'False
            Tab(1).Control(7)=   "GrdVRS"
            Tab(1).Control(7).Enabled=   0   'False
            Tab(1).Control(8)=   "GrdBRS"
            Tab(1).Control(8).Enabled=   0   'False
            Tab(1).Control(9)=   "ATXT(78)"
            Tab(1).Control(9).Enabled=   0   'False
            Tab(1).Control(10)=   "ATXT(76)"
            Tab(1).Control(10).Enabled=   0   'False
            Tab(1).Control(11)=   "ATXT(75)"
            Tab(1).Control(11).Enabled=   0   'False
            Tab(1).ControlCount=   12
            Begin MwfCtl.MWCTL_FLX GrdSumm 
               Height          =   3225
               Left            =   -74910
               TabIndex        =   153
               Top             =   420
               Width           =   12285
               _ExtentX        =   21669
               _ExtentY        =   5689
               Cols            =   13
               colname1        =   "WSUMMPRDCMCD"
               heading1        =   "Prd Cust       "
               datatype1       =   4
               maxlength1      =   8
               tooltiptext1    =   "Production Customer Code"
               colname2        =   "WSUMMPRDORD"
               heading2        =   "Prd SO                                       "
               datatype2       =   4
               maxlength2      =   22
               tooltiptext2    =   "Production Sales Order Voucher Number"
               colname3        =   "WSUMMEXPCMCD"
               heading3        =   "Exp Cust       "
               datatype3       =   4
               maxlength3      =   8
               style3          =   0
               tooltiptext3    =   "Export Customer Code"
               colname4        =   "WSUMMEXPORD"
               heading4        =   "Exp SO                                       "
               datatype4       =   4
               maxlength4      =   22
               tooltiptext4    =   "Export Sales Order Voucher Number"
               colname5        =   "WSUMMDMCD"
               heading5        =   "Design Cd                       "
               datatype5       =   4
               maxlength5      =   15
               tooltiptext5    =   "Design Code"
               colname6        =   "WSUMMPL"
               heading6        =   "Pack List                 "
               datatype6       =   4
               maxlength6      =   14
               tooltiptext6    =   "Packing List"
               colname7        =   "WSUMMINV"
               heading7        =   "Inv Sr "
               datatype7       =   1
               mask7           =   "###0"
               maxlength7      =   4
               tooltiptext7    =   "Invoice Serial"
               colname8        =   "WSUMMINVEXPNO"
               heading8        =   "Inv Exp No                                    "
               datatype8       =   4
               maxlength8      =   20
               tooltiptext8    =   "Invoice Export Number"
               colname9        =   "WSUMMFGDT"
               heading9        =   "Fg Date  "
               datatype9       =   3
               maxlength9      =   8
               tooltiptext9    =   "Finished Goods Voucher Date"
               colname10       =   "WSUMMBAGCOUNT"
               heading10       =   "No Of Bags"
               datatype10      =   1
               mask10          =   "######0"
               maxlength10     =   7
               tooltiptext10   =   "Number Of Bags"
               colname11       =   "WSUMMQTY"
               heading11       =   "Bag Qty   "
               datatype11      =   2
               mask11          =   "####0.0"
               maxlength11     =   7
               tooltiptext11   =   "Bag Quantity"
               colname12       =   "WSUMMWT"
               heading12       =   "Wt.                 "
               datatype12      =   2
               mask12          =   "####0.000"
               maxlength12     =   9
               tooltiptext12   =   "Weight"
            End
            Begin MwfCtl.MWCTL_MED ATXT 
               Height          =   285
               Index           =   27
               Left            =   -68655
               TabIndex        =   154
               ToolTipText     =   "Total Bag Quantity"
               Top             =   3720
               Width           =   975
               _ExtentX        =   1720
               _ExtentY        =   503
               Alignment       =   2
               Enabled         =   0   'False
               Mask            =   "####0.0"
               MaxLength       =   7
               DataType        =   2
               IdName          =   "WSUMMTOTQTY"
            End
            Begin MwfCtl.MWCTL_MED ATXT 
               Height          =   285
               Index           =   28
               Left            =   -66300
               TabIndex        =   155
               ToolTipText     =   "Total Bag Weight"
               Top             =   3720
               Width           =   1155
               _ExtentX        =   2037
               _ExtentY        =   503
               Alignment       =   2
               Enabled         =   0   'False
               Mask            =   "#####0.000"
               MaxLength       =   10
               DataType        =   2
               IdName          =   "WSUMMTOTWT"
            End
            Begin MwfCtl.MWCTL_MED ATXT 
               Height          =   285
               Index           =   63
               Left            =   -73845
               TabIndex        =   156
               ToolTipText     =   "Enter Group Option For Bag Summary"
               Top             =   3720
               Width           =   1365
               _ExtentX        =   2408
               _ExtentY        =   503
               MaxLength       =   8
               DataType        =   4
               IdName          =   "WSUMMBAGGRP"
            End
            Begin MwfCtl.MWCTL_MED ATXT 
               Height          =   285
               Index           =   65
               Left            =   -71160
               TabIndex        =   157
               ToolTipText     =   "Total No Of Bags"
               Top             =   3720
               Width           =   915
               _ExtentX        =   1614
               _ExtentY        =   503
               Alignment       =   2
               Enabled         =   0   'False
               Mask            =   "######0"
               MaxLength       =   7
               DataType        =   2
               IdName          =   "WSUMMTOTBAGS"
            End
            Begin MwfCtl.MWCTL_MED ATXT 
               Height          =   285
               Index           =   75
               Left            =   3210
               TabIndex        =   158
               Top             =   3690
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
               Left            =   3840
               TabIndex        =   159
               Top             =   3690
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
               Left            =   1155
               TabIndex        =   160
               ToolTipText     =   "Enter Group By Option"
               Top             =   3690
               Width           =   1365
               _ExtentX        =   2408
               _ExtentY        =   503
               MaxLength       =   8
               DataType        =   4
               IdName          =   "WBRSGRP"
            End
            Begin MwfCtl.MWCTL_FLX GrdBRS 
               Height          =   2850
               Left            =   150
               TabIndex        =   162
               Top             =   780
               Width           =   5820
               _ExtentX        =   10266
               _ExtentY        =   5027
               Cols            =   6
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
               heading3        =   "Rm Code                                   "
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
            End
            Begin MwfCtl.MWCTL_FLX GrdVRS 
               Height          =   2850
               Left            =   6045
               TabIndex        =   163
               Top             =   780
               Width           =   7320
               _ExtentX        =   12912
               _ExtentY        =   5027
               Cols            =   6
               AllowAdd        =   0   'False
               AllowDelete     =   0   'False
               colname1        =   "WVRSRMCTG"
               heading1        =   "Ctg"
               datatype1       =   4
               maxlength1      =   2
               colname2        =   "WVRSRMSCTG"
               heading2        =   "Sub Ctg"
               datatype2       =   4
               maxlength2      =   5
               colname3        =   "WVRSRMCD"
               heading3        =   "Rm Code                                   "
               datatype3       =   4
               maxlength3      =   16
               colname4        =   "WVRSRMQTY"
               heading4        =   "Quantity            "
               datatype4       =   1
               mask4           =   "####0"
               maxlength4      =   5
               colname5        =   "WVRSRMWT"
               heading5        =   "Weight                      "
               datatype5       =   2
               mask5           =   "#####0.000"
               maxlength5      =   10
            End
            Begin MwfCtl.MWCTL_MED ATXT 
               Height          =   285
               Index           =   69
               Left            =   9060
               TabIndex        =   171
               Top             =   3690
               Width           =   1215
               _ExtentX        =   2143
               _ExtentY        =   503
               Alignment       =   2
               Enabled         =   0   'False
               Mask            =   "#####0"
               MaxLength       =   6
               DataType        =   2
               IdName          =   "WVRSRMTOTQTY"
            End
            Begin MwfCtl.MWCTL_MED ATXT 
               Height          =   285
               Index           =   70
               Left            =   10290
               TabIndex        =   172
               Top             =   3690
               Width           =   1725
               _ExtentX        =   3043
               _ExtentY        =   503
               Alignment       =   2
               Enabled         =   0   'False
               Mask            =   "#####0.000"
               MaxLength       =   10
               DataType        =   2
               IdName          =   "WVRSRMTOTWT"
            End
            Begin MwfCtl.MWCTL_MED ATXT 
               Height          =   285
               Index           =   71
               Left            =   6885
               TabIndex        =   161
               ToolTipText     =   "Enter Group By Option"
               Top             =   3690
               Width           =   1365
               _ExtentX        =   2408
               _ExtentY        =   503
               MaxLength       =   8
               DataType        =   4
               IdName          =   "WVRSGRP"
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
               Index           =   63
               Left            =   5880
               TabIndex        =   173
               Top             =   3690
               Width           =   1005
            End
            Begin VB.Label ALBL 
               BackStyle       =   0  'Transparent
               Caption         =   "Tot Bags"
               BeginProperty Font 
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
               Left            =   -72120
               TabIndex        =   170
               Top             =   3720
               Width           =   1005
            End
            Begin VB.Label ALBL 
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
               Index           =   61
               Left            =   -69870
               TabIndex        =   169
               Top             =   3720
               Width           =   1245
            End
            Begin VB.Label ALBL 
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
               Index           =   33
               Left            =   -67440
               TabIndex        =   168
               Top             =   3720
               Width           =   1215
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
               Left            =   -74850
               TabIndex        =   167
               Top             =   3720
               Width           =   1005
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
               Index           =   64
               Left            =   150
               TabIndex        =   166
               Top             =   3690
               Width           =   1005
            End
            Begin VB.Label LblMain 
               BackStyle       =   0  'Transparent
               BorderStyle     =   1  'Fixed Single
               Caption         =   "  Bag Level Summary"
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
               Index           =   2
               Left            =   150
               TabIndex        =   165
               Top             =   510
               Width           =   5820
            End
            Begin VB.Label LblMain 
               BackStyle       =   0  'Transparent
               BorderStyle     =   1  'Fixed Single
               Caption         =   "  Voucher Level Summary"
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
               Index           =   3
               Left            =   6050
               TabIndex        =   164
               Top             =   510
               Width           =   7320
            End
         End
      End
      Begin VB.Frame FraCpyRfId 
         Height          =   3075
         Left            =   405
         TabIndex        =   148
         Top             =   6060
         Visible         =   0   'False
         Width           =   14355
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
            Left            =   8460
            TabIndex        =   179
            ToolTipText     =   "Specify Copy From Bag Nos / Bag Alias No."
            Top             =   180
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
            Left            =   7050
            TabIndex        =   178
            ToolTipText     =   "Specify Copy From Bag Nos / Bag Alias No."
            Top             =   180
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
            Height          =   2655
            Left            =   60
            TabIndex        =   176
            Top             =   405
            Width           =   13590
            Begin MwfCtl.MWCTL_MED ATXT 
               Height          =   2295
               Index           =   68
               Left            =   60
               TabIndex        =   177
               ToolTipText     =   "Copy Bag Alias No. / Bag Nos (Yy/Chr/No,Yy/Chr....,Yy/Chr/No)  String"
               Top             =   240
               Width           =   13455
               _ExtentX        =   23733
               _ExtentY        =   4048
               Alignment       =   3
               MaxLength       =   0
               IdName          =   "WRFIDSEL"
            End
         End
         Begin MwfCtl.MWCTL_BTN1 CmdCpyRfIdGo 
            Height          =   525
            Left            =   13680
            TabIndex        =   150
            Top             =   2400
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
            Left            =   13050
            TabIndex        =   149
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
            Left            =   5760
            TabIndex        =   180
            Top             =   170
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
            Left            =   10560
            TabIndex        =   175
            Top             =   480
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
            Left            =   0
            TabIndex        =   151
            Top             =   90
            Width           =   5610
         End
      End
   End
   Begin VB.Frame FraSC 
      Height          =   2835
      Left            =   3630
      TabIndex        =   96
      Top             =   6330
      Visible         =   0   'False
      Width           =   4740
      Begin MwfCtl.MWCTL_MED ATXT 
         Height          =   285
         Index           =   37
         Left            =   1245
         TabIndex        =   97
         ToolTipText     =   "Enter Production Customer"
         Top             =   420
         Width           =   1320
         _ExtentX        =   2328
         _ExtentY        =   503
         MaxLength       =   8
         DataType        =   4
         IdName          =   "WSCPRDCMCD"
      End
      Begin MwfCtl.MWCTL_MED ATXT 
         Height          =   285
         Index           =   38
         Left            =   1245
         TabIndex        =   114
         ToolTipText     =   "Enter From Finished Goods Date (Can Be Blank)"
         Top             =   2415
         Width           =   855
         _ExtentX        =   1508
         _ExtentY        =   503
         MaxLength       =   8
         DataType        =   3
         IdName          =   "WSCFGDTFR"
      End
      Begin MwfCtl.MWCTL_MED ATXT 
         Height          =   285
         Index           =   39
         Left            =   2700
         TabIndex        =   115
         ToolTipText     =   "Enter To Finished Goods Date (Can Be Blank)"
         Top             =   2415
         Width           =   885
         _ExtentX        =   1561
         _ExtentY        =   503
         MaxLength       =   8
         DataType        =   3
         IdName          =   "WSCFGDTTO"
      End
      Begin MwfCtl.MWCTL_MED ATXT 
         Height          =   285
         Index           =   41
         Left            =   1950
         TabIndex        =   100
         ToolTipText     =   "Enter Production Sales Order Year"
         Top             =   990
         Width           =   360
         _ExtentX        =   635
         _ExtentY        =   503
         MaxLength       =   2
         DataType        =   4
         IdName          =   "WSCPRDODYY"
      End
      Begin MwfCtl.MWCTL_MED ATXT 
         Height          =   285
         Index           =   43
         Left            =   3120
         TabIndex        =   102
         ToolTipText     =   "Enter Production Sales Order Number"
         Top             =   990
         Width           =   780
         _ExtentX        =   1376
         _ExtentY        =   503
         Alignment       =   2
         Mask            =   "#####0"
         MaxLength       =   6
         DataType        =   1
         IdName          =   "WSCPRDODNO"
      End
      Begin MwfCtl.MWCTL_MED ATXT 
         Height          =   285
         Index           =   45
         Left            =   1245
         TabIndex        =   99
         ToolTipText     =   "Enter Production Sales Order Tc"
         Top             =   990
         Width           =   600
         _ExtentX        =   1058
         _ExtentY        =   503
         MaxLength       =   3
         DataType        =   4
         IdName          =   "WSCPRDODTC"
      End
      Begin MwfCtl.MWCTL_MED ATXT 
         Height          =   285
         Index           =   46
         Left            =   2430
         TabIndex        =   101
         ToolTipText     =   "Enter Production Sales Order Character"
         Top             =   990
         Width           =   600
         _ExtentX        =   1058
         _ExtentY        =   503
         MaxLength       =   3
         DataType        =   4
         IdName          =   "WSCPRDODCHR"
      End
      Begin MwfCtl.MWCTL_MED ATXT 
         Height          =   285
         Index           =   47
         Left            =   2430
         TabIndex        =   112
         ToolTipText     =   "Enter Packing List Number"
         Top             =   1845
         Width           =   930
         _ExtentX        =   1640
         _ExtentY        =   503
         Alignment       =   2
         Mask            =   "######0"
         MaxLength       =   7
         DataType        =   1
         IdName          =   "WSCPLNO"
      End
      Begin MwfCtl.MWCTL_MED ATXT 
         Height          =   285
         Index           =   48
         Left            =   1950
         TabIndex        =   111
         ToolTipText     =   "Enter Packing List Year"
         Top             =   1845
         Width           =   360
         _ExtentX        =   635
         _ExtentY        =   503
         MaxLength       =   2
         DataType        =   4
         IdName          =   "WSCPLYY"
      End
      Begin MwfCtl.MWCTL_MED ATXT 
         Height          =   285
         Index           =   49
         Left            =   1245
         TabIndex        =   110
         ToolTipText     =   "Enter Packing List Character"
         Top             =   1845
         Width           =   600
         _ExtentX        =   1058
         _ExtentY        =   503
         MaxLength       =   3
         DataType        =   4
         IdName          =   "WSCPLCHR"
      End
      Begin MwfCtl.MWCTL_MED ATXT 
         Height          =   285
         Index           =   50
         Left            =   3975
         TabIndex        =   103
         ToolTipText     =   "Enter Production Sales Order Serial Number"
         Top             =   990
         Width           =   570
         _ExtentX        =   1005
         _ExtentY        =   503
         Alignment       =   2
         Mask            =   "###0"
         MaxLength       =   4
         DataType        =   1
         IdName          =   "WSCPRDODSR"
      End
      Begin MwfCtl.MWCTL_MED ATXT 
         Height          =   285
         Index           =   51
         Left            =   1245
         TabIndex        =   98
         ToolTipText     =   "Enter Export Customer"
         Top             =   705
         Width           =   1320
         _ExtentX        =   2328
         _ExtentY        =   503
         MaxLength       =   8
         DataType        =   4
         IdName          =   "WSCEXPCMCD"
      End
      Begin MwfCtl.MWCTL_MED ATXT 
         Height          =   285
         Index           =   52
         Left            =   1950
         TabIndex        =   105
         ToolTipText     =   "Enter Export Sales Order Year"
         Top             =   1275
         Width           =   360
         _ExtentX        =   635
         _ExtentY        =   503
         MaxLength       =   2
         DataType        =   4
         IdName          =   "WSCEXPODYY"
      End
      Begin MwfCtl.MWCTL_MED ATXT 
         Height          =   285
         Index           =   53
         Left            =   3120
         TabIndex        =   107
         ToolTipText     =   "Enter Export Sales Order Number"
         Top             =   1275
         Width           =   780
         _ExtentX        =   1376
         _ExtentY        =   503
         Alignment       =   2
         Mask            =   "#####0"
         MaxLength       =   6
         DataType        =   1
         IdName          =   "WSCEXPODNO"
      End
      Begin MwfCtl.MWCTL_MED ATXT 
         Height          =   285
         Index           =   54
         Left            =   1245
         TabIndex        =   104
         ToolTipText     =   "Enter Export Sales Order Tc"
         Top             =   1275
         Width           =   600
         _ExtentX        =   1058
         _ExtentY        =   503
         MaxLength       =   3
         DataType        =   4
         IdName          =   "WSCEXPODTC"
      End
      Begin MwfCtl.MWCTL_MED ATXT 
         Height          =   285
         Index           =   55
         Left            =   2430
         TabIndex        =   106
         ToolTipText     =   "Enter Export Sales Order Character"
         Top             =   1275
         Width           =   600
         _ExtentX        =   1058
         _ExtentY        =   503
         MaxLength       =   3
         DataType        =   4
         IdName          =   "WSCEXPODCHR"
      End
      Begin MwfCtl.MWCTL_MED ATXT 
         Height          =   285
         Index           =   56
         Left            =   3975
         TabIndex        =   108
         ToolTipText     =   "Enter Export Sales Order Serial Number"
         Top             =   1275
         Width           =   570
         _ExtentX        =   1005
         _ExtentY        =   503
         Alignment       =   2
         Mask            =   "###0"
         MaxLength       =   4
         DataType        =   1
         IdName          =   "WSCEXPODSR"
      End
      Begin MwfCtl.MWCTL_MED ATXT 
         Height          =   285
         Index           =   57
         Left            =   1245
         TabIndex        =   109
         ToolTipText     =   "Enter Design Code"
         Top             =   1560
         Width           =   2415
         _ExtentX        =   4260
         _ExtentY        =   503
         MaxLength       =   15
         DataType        =   4
         IdName          =   "WSCDMCD"
      End
      Begin MwfCtl.MWCTL_MED ATXT 
         Height          =   285
         Index           =   62
         Left            =   1245
         TabIndex        =   113
         ToolTipText     =   "Enter Invoice Serial Number"
         Top             =   2130
         Width           =   570
         _ExtentX        =   1005
         _ExtentY        =   503
         Alignment       =   2
         Mask            =   "###0"
         MaxLength       =   4
         DataType        =   1
         IdName          =   "WSCIDSR"
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
         Index           =   0
         Left            =   -30
         TabIndex        =   139
         Top             =   90
         Width           =   4770
      End
      Begin VB.Label ALBL 
         BackStyle       =   0  'Transparent
         Caption         =   "Invoice Sr"
         BeginProperty Font 
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
         TabIndex        =   134
         Top             =   2130
         Width           =   1215
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
         Index           =   43
         Left            =   90
         TabIndex        =   133
         Top             =   1560
         Width           =   1095
      End
      Begin VB.Label ALBL 
         BackStyle       =   0  'Transparent
         Caption         =   "Exp Cust"
         BeginProperty Font 
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
         Left            =   90
         TabIndex        =   132
         Top             =   705
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
         Index           =   46
         Left            =   3045
         TabIndex        =   131
         Top             =   1275
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
         Left            =   2325
         TabIndex        =   130
         Top             =   1275
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
         Left            =   1860
         TabIndex        =   129
         Top             =   1275
         Width           =   105
      End
      Begin VB.Label ALBL 
         BackStyle       =   0  'Transparent
         Caption         =   "Exp SO"
         BeginProperty Font 
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
         Left            =   90
         TabIndex        =   128
         Top             =   1275
         Width           =   1215
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
         Index           =   50
         Left            =   3900
         TabIndex        =   127
         Top             =   1275
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
         Index           =   51
         Left            =   3900
         TabIndex        =   126
         Top             =   990
         Width           =   105
      End
      Begin VB.Label ALBL 
         BackStyle       =   0  'Transparent
         Caption         =   "Pck List"
         BeginProperty Font 
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
         Left            =   90
         TabIndex        =   125
         Top             =   1845
         Width           =   1035
      End
      Begin VB.Label Label1 
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
         Left            =   2325
         TabIndex        =   124
         Top             =   1845
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
         Index           =   53
         Left            =   1860
         TabIndex        =   123
         Top             =   1845
         Width           =   135
      End
      Begin VB.Label ALBL 
         BackStyle       =   0  'Transparent
         Caption         =   "Prd SO"
         BeginProperty Font 
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
         Left            =   90
         TabIndex        =   122
         Top             =   990
         Width           =   1215
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
         Index           =   55
         Left            =   1860
         TabIndex        =   121
         Top             =   990
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
         Index           =   56
         Left            =   2325
         TabIndex        =   120
         Top             =   990
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
         Left            =   3045
         TabIndex        =   119
         Top             =   990
         Width           =   105
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
         Index           =   58
         Left            =   2400
         TabIndex        =   118
         Top             =   2415
         Width           =   345
      End
      Begin VB.Label ALBL 
         BackStyle       =   0  'Transparent
         Caption         =   "Fg Dt  Fr"
         BeginProperty Font 
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
         TabIndex        =   117
         Top             =   2415
         Width           =   975
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
         Index           =   60
         Left            =   90
         TabIndex        =   116
         Top             =   420
         Width           =   1095
      End
   End
End
Attribute VB_Name = "EmrFrmInvAllBags"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'----- General comments on the Entry  ----
' X-'', A-InvFgd(GrdInvFgd)
'
' Copy option is used to copy bags into Invoice from Finished Goods
' FraCP, CmdCP, CmdCPGo, wCpPrdCmCd, wCpExpCmCd, wCpPrdOdTc, wCpPrdOdYy, wCpPrdOdChr,
'        wCpPrdOdNo, wCpPrdOdSr,wCpExpOdTc, wCpExpOdYy, wCpExpOdChr, wCpExpOdNo,
'        wCpExpOdSr, wCpDmCd, wCpPlChr, wCpPlYy, wCpPlNo, wCpInTc, wCpInYy, wCpInChr,
'        wCpInNo, wCpIdSr, wCpInExpNo, wCpFgDtFr, wCpFgDtTo
'        all part of Copy routine.
'
' Scope option allows the user to see only those bags which fall within the scope specified
' by the user
' FraSC, CmdSC, wScPrdCmCd, wScExpCmCd, wScPrdOdTc, wScPrdOdYy, wScPrdOdChr,
'        wScPrdOdNo, wScPrdOdSr, wScExpOdTc, wScExpOdYy, wScExpOdChr, wScExpOdNo,
'        wScExpOdSr, wScDmCd, wScPlChr, wScPlYy, wScPlNo, wScIdSr, wScFgDtFr, wScFgDtTo
'        all part of Selection Scope routine.
'
' Summary option summarises the Bag details of the entire Invoice
' FraSumm, CmdSumm, GrdSumm, wSummBagGrp, wSummTotBags, wSummTotQty, wSummTotWt
'        all part of Summary routine.
'
' private Sub CmpString to set the values for Compare String property,
'        passing 'K'ey Fields, 'C'opy
'
'-----

Option Explicit
Dim moCn As MwfLib.MDOConnection   '*** (Bef speed) bef 24/11/06
Dim ms_IfTc As String, ms_IfYy As String, ms_IfChr As String
Dim ms_TcTyp As String, ms_CmCtg As String

'*** Declare the variables ms_LockDt, mf_DtLocked for the Lock Date Check
Dim ms_LockDt As String
Dim ms_FutureDt As String
Dim mf_DtLocked As Boolean

Dim ms_PurgeDt As String        '****** Sachin 3.02 28-12-07

'*** For the Toggle Options of Command Buttons
Enum en_InvBagFra
  Cpy = 0
  Scop = 1
  Summ = 2
  CpyRfId = 3     ' ***** Manali 3.10.0 - 19/04/12 - Cpy Bags from RfId
  DelAllBag = 4
End Enum

' **** Manali 3.03 - 21/06/08 - Pd Mod and Slv Mod
Dim mb_SlvMod As Boolean, mb_PdMod As Boolean
' **** Manali 3.03 - 21/06/08 - Pd Mod

Dim wDsgExists As Boolean, wInCls As String, wSOTcTyp As String
'*** Arr_Cv stores the causes validation state, Arr_Ed the Ena/Disa state of Cmd Buttons
Dim mArr_Cv As Variant, mArr_Ed As Variant

Private Sub ADC_Load()
  '*** Set the First Key Control property to the field wIfChr
  '*** Set the First Non Key Control property to GrdInvFgd
  '*** Set the Child Property Of ADC to GrdInvFgd
  '*** Set the Previous and Next Control for the Grid GrdInvFgd as ADC
  '*** Set the Hot Keys for the Grid GrdInvFgd as 'A'
  '*** Store the defa values for Inv Tc,Yy in variables
  
  'ReDim mArr_Cv(3), mArr_Ed(3)
  ReDim mArr_Cv(4), mArr_Ed(4)      ' ***** Manali 3.10.0 - 19/04/12 - Cpy Bags from RfId
  
  Set moCn = ADC.Connection    '*** (Bef speed) bef 24/11/06

  Set ADC.FirKeyCtl = ADC("wIfChr")
  Set ADC.FirNKeyCtl = CmdCP     '****** (The reason why GrdInvFgd is not set as the
                                        'FirstKeyControl is that it unnecessarily creates
                                        'a row when there are no rows in the grid)
  ADC.Child = "GrdInvFgd"

  Set GrdInvFgd.PrevCtl = ADC
  Set GrdInvFgd.NextCtl = ADC
  Set GrdSumm.NextCtl = ADC("wSummBagGrp")
  GrdInvFgd.HotKey = "A"

  Select Case UCase(ADC.MenuCd)
  Case Is = UCase("InAllBags")
    ms_TcTyp = ctInTcTyp
    ms_CmCtg = "C"
    LblCpCtCd.Visible = False
    ADC("wCpCtCd").Enabled = False
    ADC("wCpCtCd").Visible = False
  Case Is = UCase("CTBInAllBags")
    ms_TcTyp = ctCTBTcTyp
    ms_CmCtg = "T"
    LblCpCtCd.Visible = True
    ADC("wCpCtCd").Enabled = True
    ADC("wCpCtCd").Visible = True
  End Select

  ms_IfTc = moCn.GetFldVal("Select PValue from Param where PTyp= 'TCTYP' and " + _
            "PMCd='" + ms_TcTyp + "' ")
  ms_IfYy = moCn.GetFldVal("Select hYy from Head where HCoCd= '" + gs_CoCd + "' and HCd='" + ctSelfCmCd + "'")
  wSOTcTyp = "SO"
  
  ' **** Manali 3.03 - 21/06/08 - Pd Mod and Silver
  mb_SlvMod = moCn.RecSeek("Select 'x' from Head where HCoCd= '" + gs_CoCd + "' and HCd='" + ctSelfCmCd + "' and HSlvModYn='Y'")
  mb_PdMod = moCn.RecSeek("Select 'x' from Head where HCoCd= '" + gs_CoCd + "' and HCd='" + ctSelfCmCd + "' and HPdModYn='Y'")
  
  If mb_SlvMod = False Then
    LblIfLmsSal.Visible = False: ADC("wIfLmsSal").Visible = False: ADC("wIfLmsSal").Enabled = False
  End If
  
  If mb_PdMod = False Then
    LblIfLmlSal.Visible = False: ADC("wIfLmlSal").Visible = False: ADC("wIfLmlSal").Enabled = False
  End If
  ' **** Manali 3.03 - 21/06/08 - Pd Mod and Silver
  
  '****** Sachin 3.02.0 - Hiding the Partition Key Column based on Partition Selected
  ADC("WIFPRTKEY").Visible = IIF(gs_Partition = ctCurrPrtn, False, True)
  GrdInvFgd.ColProp("IfPrtKey").Style = IIF(gs_Partition = ctCurrPrtn, fgcinvisible, fgcReadOnly)
  
  ' ***** Manali 3.10.0 - 19/04/12 - Cpy Bags from RfId
  '6.1b-22 Rfid module check is removed
'  If Not (UCase(ADC.MenuCd) = UCase("InAllBags") And moCn.RecSeek("Select 'x' from Head Where HCoCd='" + gs_CoCd + "' and HCd='" + ctSelfCmCd + "' and HRfIdMod='Y' ")) Then
'    CmdCpyRfId.Enabled = False: CmdCpyRfId.Visible = False
'  End If
  ' ***** Manali 3.10.0 - 19/04/12 - Cpy Bags from RfId
  'sv.9 getting purge date from vparam
  ms_PurgeDt = GetPurgeDt(gs_CoCd, "")
End Sub
Private Sub ADC_KeyWhen()
  '*** set the Compare String Property of the ADC controls when the Focus is in the Key Frame
      '(The CmpStr Property of only the fields in the Scope frame are set and those of
      'the Copy frame are made blank)
  '*** setting the default values for wIfTc,wIfYy and wIfChr from the previous entry
  '*** Disable All command buttons
  '*** Hide all frames except Frame FraNKeyAll
  
  mf_DtLocked = False
  Call CmpString("K")
  If ms_IfTc <> "" Then ADC("wIfTc") = ms_IfTc
  If ms_IfYy <> "" Then ADC("wIfYy") = ms_IfYy
  If ms_IfChr <> "" Then ADC("wIfChr") = ms_IfChr
  
  Call EnaDisaCmds(True)
  CmdSC.Enabled = True
  CmdDelAllBags.Enabled = False
  Call HideAllFras
  
  '****** Sachin 2.14.0 - [24-11-2006] - [Multi-Company] ******
  Me.Caption = GetFrmCaption(gs_CoCd, "Invoice Bag Entry")
  '****** Sachin 2.14.0 - [24-11-2006] - [Multi-Company] ******
  
End Sub
Private Sub ADC_KeyValid(Cancel As Boolean, ErrMsg As String)
  '*** Checking Key values in Key Valid as they are copied in Init
  '*** Check whether the wIfYy, wIfChr, wIfNo are blank

  If ADC("wIfYy") = "" Then Cancel = True: ErrMsg = "Enter Invoice Voucher Year": Exit Sub
  If ADC("wIfChr") = "" Then Cancel = True: ErrMsg = "Enter Invoice Voucher Character": Exit Sub
  If ADC("wIfNo") = 0 Then Cancel = True: ErrMsg = "Enter Invoice Voucher Number": Exit Sub
  Cancel = Not moCn.RecSeek("Select InNo From InvHd Where " + IIF(gs_Partition = ctCurrPrtn, " InPrtKey='" + ctCurrPrtn + "' And ", "") + " InCoCd= '" + gs_CoCd + "' " + _
           "And InTc= '" + ADC("wIfTc") + "' And InYy= '" + ADC("wIfYy") + "' " + _
           "And InChr= '" + ADC("wIfChr") + "' And InNo= " + CStr(ADC("wIfNo")))
  If Cancel = True Then ErrMsg = "Invalid Invoice Voucher": Exit Sub
End Sub
Private Sub ADC_NonKeyWhen(pr_Cancel As Boolean, pr_ErrMsg As String)
  '*** Store the wIfTc,wIfYy and wIfChr values for the next entry
  '*** Enable All Command Buttons Except Scope Button
  '*** give default values to the working fields wIfDt, wIfCmCd, wIfExpNo, wIfLmgSal from
  '    the InvHd table
  '*** 'A' grid will become non-editable (only a front end check),
      'once any records exist in InvDsg, (i.e. record exist in InvFgd with Sr<>0) or
      'if the Invoice is closed
  
      If ADC.IsNew And (Not moCn.RecSeek("select vPSCd from vParam where " + _
                         "vPCoCd = '" + gs_CoCd + "' and vPTyp= 'CHR' and " + _
                         "vPMCd = '" + ADC("wIfTc") + "' and vPSCd='" + ADC("wIfChr") + "' and vPValidYn<>'N'")) Then
       pr_Cancel = True: pr_ErrMsg = "Voucher Character No More Valid": ADC.AllowSave = False: Exit Sub
      Else
        ADC.AllowSave = True
      End If

  'ms_LockDt = moCn.GetFldVal("Select PDesc225 from Param where PTyp= 'LOCKDT' and PMCd= 'LOCKDT'")
  'ms_FutureDt = moCn.GetFldVal("Select PDesc225 from Param where PTyp= 'LOCKDT' and PMCd= 'FUTUREDT'")
  
  '****** Sachin 3.02 28-12-07
  'ms_PurgeDt = moCn.GetFldVal("Select PDesc225 from Param where PTyp= 'PURGEDT' and PMCd= 'PURGEDT'")
  
  Call GetLockDt(gs_CoCd, "TC", ms_IfTc, ms_LockDt, ms_FutureDt)
  
  '*** For the time being do not put Lock Date Check in the Invoice Entry
  '### If ms_LockDt = "" Then ms_LockDt = "01/01/80"
  'If ms_FutureDt = "" Then ms_FutureDt = "01/01/80"
  ms_LockDt = "01/01/80"
  ms_FutureDt = IIF(UCase(MWLib.DtSetting) = "MM/DD/YY", "12/31/28", "31/12/28")   '6.1
  
  'If ms_PurgeDt = "" Then ms_PurgeDt = "01/01/80"   '****** Sachin 3.02 28-12-07
  
  Dim wRsInvDet As MwfLib.MDORowSet
  ms_IfTc = ADC("wIfTc")
  ms_IfYy = ADC("wIfYy")
  ms_IfChr = ADC("wIfChr")
    
  '*** This is required if the user fails to close the Scope frame
  If FraSC.Visible = True Then
    '*** Display or hide the Scope frame depending on the toggle satate
    Call DispFra(Scop)
  End If
    
  CmdCP.Enabled = True
  CmdSumm.Enabled = True
  FraNKeyAll.Enabled = True
  ADC.AllowSave = True
  'Call EnaDisaCmds(False)
  CmdSC.Enabled = False
  CmdCpyRfId.Enabled = True       ' ***** Manali 3.10.0 - 19/04/12 - Cpy Bags from RfId
  CmdDelAllBags.Enabled = True
  
  '****** Sachin 3.02 28-12-07
  
  ' **** Manali 3.03 - 21/06/08 - LMS, LML added
  Set wRsInvDet = moCn.OpenRes("Select InDt, InCmCd, InExpNo, InLmgSal, InLmpSal, InLmsSal, InLmlSal, InPrtKey from InvHd where " + _
                               "InCoCd='" + gs_CoCd + "' and InTc='" + ADC("wIfTc") + "' and " + _
                               "InYy='" + ADC("wIfYy") + "' and InChr='" + ADC("wIfChr") + "' and " + _
                               "InNo= " + CStr(ADC("wIfNo")) + " and InPrtKey=(Case When '" + gs_Partition + "'='" + ctCurrPrtn + "' Then '" + ctCurrPrtn + "' Else InPrtKey End) ")
  If Not (wRsInvDet.EOF Or wRsInvDet.BOF) Then
    ADC("wIfDt") = wRsInvDet!InDt
    ADC("wIfCmCd") = wRsInvDet!InCmCd
    ADC("wIfExpNo") = wRsInvDet!InExpNo
    ADC("wIfLmgSal") = wRsInvDet!InLmgSal
    ADC("wIfLmpSal") = wRsInvDet!InLmpSal
    ADC("wIfLmpSal") = wRsInvDet!InLmpSal
    ' **** Manali 3.03 - 21/06/08 - Pd Mod and silver
    ADC("wIfLmsSal") = wRsInvDet!InLmsSal
    ADC("wIfLmlSal") = wRsInvDet!InLmlSal
    ' **** Manali 3.03 - 21/06/08 - Pd Mod and silver
    
    ADC("wIfPrtKey") = wRsInvDet!InPrtKey       '****** Sachin 3.02 28-12-07
    
  End If
  
  mf_DtLocked = IIF((ADC("wIfDt") < CDate(ms_LockDt)) _
                               Or _
                   (ADC("wIfDt") > CDate(ms_FutureDt)), _
                True, False)

  '****** Sachin 3.02 - Locking Previous Partition Records
  If mf_DtLocked = True Or ADC("wIfPrtKey") <> ctCurrPrtn Then
    GrdInvFgd.AllowAdd = False: GrdInvFgd.AllowDelete = False
  Else
    GrdInvFgd.AllowAdd = True: GrdInvFgd.AllowDelete = True
  End If
  
  '****** Sachin 3.02 28-12-07
  wInCls = moCn.GetFldVal("Select InCls from InvHd where InCoCd= '" + gs_CoCd + "' and " + _
                          "InTc= '" + ADC("wIfTc") + "' and InYy= '" + ADC("wIfYy") + "' and " + _
                          "InChr= '" + ADC("wIfChr") + "' and InNo= " + CStr(ADC("wIfNo")) + " and InPrtKey=(Case When '" + ADC("wIfPrtKey") + "'='" + ctCurrPrtn + "' Then '" + ctCurrPrtn + "' Else InPrtKey End)")
  
  '****** Sachin 3.02 28-12-07
  wDsgExists = moCn.RecSeek("Select IdTc from InvDsg where IdCoCd= '" + gs_CoCd + "' and " + _
                          "IdTc= '" + ADC("wIfTc") + "' and IdYy= '" + ADC("wIfYy") + "' and " + _
                          "IdChr= '" + ADC("wIfChr") + "' and IdNo= " + CStr(ADC("wIfNo")) + " and IdPrtKey=(Case When '" + ADC("wIfPrtKey") + "'='" + ctCurrPrtn + "' Then '" + ctCurrPrtn + "' Else IdPrtKey End) ")

  If wInCls = "Y" Or wDsgExists = True Then
    CmdCP.Enabled = False
    GrdInvFgd.AllowAdd = False: GrdInvFgd.AllowDelete = False
  Else
    CmdCP.Enabled = True
    If mf_DtLocked = False Then GrdInvFgd.AllowAdd = True: GrdInvFgd.AllowDelete = True
  End If

  '****** Sachin 2.14.0 [Next] [07-02-2007] ******
  '****** Sachin 3.02 28-12-07
  If moCn.RecSeek("Select 'x' From InvDsg Where IdCoCd='" + gs_CoCd + "' And IdTc='" + ADC("WIFTC") + "' And IdYy='" + ADC("WIFYY") + "' And IdChr='" + ADC("WIFCHR") + "' And IdNo=" + CStr(ADC("WIFNO")) + " and IdPrtKey=(Case When '" + ADC("wIfPrtKey") + "'='" + ctCurrPrtn + "' Then '" + ctCurrPrtn + "' Else IdPrtKey End) ") Then
      GrdInvFgd.ColProp("IFSR").Style = fgcReadOnly
  Else
      GrdInvFgd.ColProp("IFSR").Style = fgcinvisible
  End If
  '****** Sachin 2.14.0 [Next] [07-02-2007] ******
  'vk.4 setting default values
  OptCpyBagFr(0).Value = True
  ADC("wCpyThruBagNo") = "Y"

End Sub
Private Sub ATXT_FldWhen(Index As Integer, ByVal IdName As String, Cancel As Boolean, ErrMsg As String)
  '*** Give help for wIfTc, wIfYy, wIfChr, wIfNo, wIfSrt
  '*** Give help for all fields in the Scope frame
  '*** Give help for all the fields in the Copy frame
  '*** Give help for the field wSummBagGrp in the Summary frame   (PTyp= 'INBAGGR')
  
  If mf_DtLocked = True And (IdName <> UCase("wSummBagGrp")) Then
    Cancel = True: ErrMsg = "Cannot Edit Locked Entry": Exit Sub
  End If
  
  Select Case IdName
  Case Is = UCase("wIfTc")
    Call HlpList.PMCd("TC", "'" + ms_TcTyp + "'")
  Case Is = UCase("wIfYy")
    Call HlpList.PMCd("YY")
  Case Is = UCase("wIfChr")
' **** Zubin 211 **** '
    'Call HlpList.vPSCd(gs_CoCd, "CHR", ADC("wIfTc"))
    Call HlpList.vPSCd(gs_CoCd, "CHR", ADC("wIfTc"), "Y")
' **** Zubin 211 **** '
  Case Is = UCase("wIfNo")
    Call HlpList.InNo(gs_CoCd, ADC("wIfTc"), ADC("wIfYy"), ADC("wIfChr"))
  Case Is = UCase("wIfSrt")
    Call HlpList.PMCd("SRTINBAG")
  Case Is = UCase("wSummBagGrp")
    Call HlpList.PMCd("INBAGGR")
    
  'Sachin 3.11.0
  Case Is = UCase("wBrsGrp"), UCase("wVrsGrp")
    Call HlpList.PMCd("TBAGBSGR")

' fields of Copy frame & Selection Scope frame.
  Case Is = UCase("wCpPrdCmCd"), UCase("wCpExpCmCd"), UCase("wScPrdCmCd"), UCase("wScExpCmCd")
    Call HlpList.CustCd("C")
  Case Is = UCase("wCpCtCd")
    Call HlpList.CustCd("T")
  Case Is = UCase("wCpPrdOdTc"), UCase("wCpExpOdTc"), UCase("wScPrdOdTc"), UCase("wScExpOdTc")
    Call HlpList.PMCd("TC", "'SO'")
  Case Is = UCase("wCpPrdOdYy"), UCase("wCpExpOdYy"), UCase("wScPrdOdYy"), UCase("wScExpOdYy")
    Call HlpList.PMCd("YY")
  Case Is = UCase("wCpPrdOdChr")
    Call HlpList.vPSCd(gs_CoCd, "CHR", ADC("wCpPrdOdTc"))
  Case Is = UCase("wCpExpOdChr")
    Call HlpList.vPSCd(gs_CoCd, "CHR", ADC("wCpExpOdTc"))
  Case Is = UCase("wScPrdOdChr")
    Call HlpList.vPSCd(gs_CoCd, "CHR", ADC("wScPrdOdTc"))
  Case Is = UCase("wScExpOdChr")
    Call HlpList.vPSCd(gs_CoCd, "CHR", ADC("wScExpOdTc"))
  Case Is = UCase("wCpPrdOdNo")
    Call HlpList.OmNo(gs_CoCd, ADC("wCpPrdOdTc"), ADC("wCpPrdOdYy"), ADC("wCpPrdOdChr"))
  Case Is = UCase("wCpExpOdNo")
    Call HlpList.OmNo(gs_CoCd, ADC("wCpExpOdTc"), ADC("wCpExpOdYy"), ADC("wCpExpOdChr"))
  Case Is = UCase("wScPrdOdNo")
    Call HlpList.OmNo(gs_CoCd, ADC("wScPrdOdTc"), ADC("wScPrdOdYy"), ADC("wScPrdOdChr"))
  Case Is = UCase("wScExpOdNo")
    Call HlpList.OmNo(gs_CoCd, ADC("wScExpOdTc"), ADC("wScExpOdYy"), ADC("wScExpOdChr"))
  Case Is = UCase("wCpPrdOdSr")
    Call HlpList.OdSr(gs_CoCd, ADC("wCpPrdOdTc"), ADC("wCpPrdOdYy"), ADC("wCpPrdOdChr"), ADC("wCpPrdOdNo"))
  Case Is = UCase("wCpExpOdSr")
    Call HlpList.OdSr(gs_CoCd, ADC("wCpExpOdTc"), ADC("wCpExpOdYy"), ADC("wCpExpOdChr"), ADC("wCpExpOdNo"))
  Case Is = UCase("wScPrdOdSr")
    Call HlpList.OdSr(gs_CoCd, ADC("wScPrdOdTc"), ADC("wScPrdOdYy"), ADC("wScPrdOdChr"), ADC("wCpExpOdNo"))
  Case Is = UCase("wScExpOdSr")
    Call HlpList.OdSr(gs_CoCd, ADC("wScExpOdTc"), ADC("wScExpOdYy"), ADC("wScExpOdChr"), ADC("wCpExpOdNo"))
  Case Is = UCase("wCpDmCtg")
    Call HlpList.PMCd("DMCTG")     ' ***** Manali 3.5.0 - 11/11/08 - DmCtg added in scope for copy
  Case Is = UCase("wCpDmCd")
    Call HlpList.DmCd("DM")
  Case Is = UCase("wScDmCd")
    Call HlpList.DmCd("DM")
  Case Is = UCase("wCpPlChr"), UCase("wScPlChr")
    Call HlpList.PMCd("PLCHR")
  Case Is = UCase("wCpPlYy"), UCase("wScPlYy")
    Call HlpList.PMCd("YY")
  Case Is = UCase("wCpPlNo")
    Call HlpList.PlNo(gs_CoCd, ADC("wCpPlYy"), ADC("wCpPlChr"))
  Case Is = UCase("wScPlNo")
    Call HlpList.PlNo(gs_CoCd, ADC("wScPlYy"), ADC("wScPlChr"))
  Case Is = UCase("wCpInExpNo")
    Call HlpList.InExpNo(gs_CoCd)
  Case Is = UCase("wCpInTc")
    Call HlpList.PMCd("TC", "'" + ctInTcTyp + "'")
  Case Is = UCase("wCpInYy")
    Call HlpList.PMCd("YY")
  Case Is = UCase("wCpInChr")
    Call HlpList.vPSCd(gs_CoCd, "CHR", ADC("wCpInTc"))
  Case Is = UCase("wCpInNo")
    Call HlpList.InNo(gs_CoCd, ADC("wCpInTc"), ADC("wCpInYy"), ADC("wCpInChr"))
  Case Is = UCase("wCpIdSr")
    Call HlpList.InSr(gs_CoCd, ADC("wCpInTc"), ADC("wCpInYy"), ADC("wCpInChr"), ADC("wCpInNo"))
  Case Is = UCase("wScIdSr")
    Call HlpList.InSr(gs_CoCd, ADC("wIfTc"), ADC("wIfYy"), ADC("wIfChr"), ADC("wIfNo"))
  Case Is = UCase("wCpFgSubLoc")
    Call HlpList.vPMCd(gs_CoCd, "FGSUBLOC")
     '*********Bhavna cpy thru bag no
    Case Is = UCase("wCPYTHRUBAGNO")
        If ADC("WCPYTHRUBAGNO") = "" Then
            ADC("WCPYTHRUBAGNO") = "N"
        End If
  '6.1_11 calling help for shipment note scope
  Case Is = UCase("wCpSnTc")
    Call HlpList.PMCd("TC", "'SN'")
  Case Is = UCase("wCpSnYy")
    Call HlpList.PMCd("YY")
  Case Is = UCase("wCpSnChr")
    Call HlpList.vPSCd(gs_CoCd, "CHR", ADC("wCpSnTc"))
  Case Is = UCase("wCpSnNo")
    Call HlpList.ShpNoteNo(gs_CoCd, ADC("wCpSnTc"), ADC("wCpSnYy"), ADC("wCpSnChr"))
    
  End Select
End Sub
Private Sub atxt_ReCalculate(Index As Integer, ByVal IdName As String, ByVal OldVal As Variant, ByVal RowNum As Integer, ByVal SourceName As String, ByVal SourceColName As String)
  '*** wIfChr on wIfTc                                    Normal
      'PDesc225('TC', adc(wIfTc))
  
  ''*** wIfPrdOrd on IfFdTc,IfFdYy,IfFdChr,IfFdNo,IfFdSr   ADC=Normal and GrdInvFgd=Normal
  '    'FdPrdOdTc+ FdPrdOdYy+ FdPrdOdChr+ FdPrdOdNo+ FdPrdOdSr
  '    'from Fgd(FdTc= IfFdTc,FdYy= IfFdYy,FdChr= IfFdChr,FdNo= IfFdNo,FdSr= IfFdSr)
  ''*** wIfPrdCmCd on IfFdTc,IfFdYy,IfFdChr,IfFdNo,IfFdSr  ADC=Normal and GrdInvFgd=Normal
  '    'FdPrdCmCd from Fgd(FdTc= IfFdTc,FdYy= IfFdYy,FdChr= IfFdChr,FdNo= IfFdNo,FdSr= IfFdSr)
  
  '*** wIfTotBags on IfSrNo,IfFdBNo                       ADC=Normal and GrdInvFgd=Normal,Delete,InitOver,CopyOver
      'count(IfFdBNo<> 0)
  '*** wIfTotQty on IfQty                                 ADC=Normal and GrdInvFgd=Normal,Delete,InitOver,CopyOver
      'sum(IfQty)
  '*** wIfTotWt on IfGrsWt                                ADC=Normal and GrdInvFgd=Normal,Delete,InitOver,CopyOver
      'sum(IfGrsWt)
    
  ''wIfTotQty on IfFdTc,IfFdYy,IfFdChr,IfFdNo,IfFdSr       Normal
  '    'sum(IfQty)
  ''wIfTotWt on IfFdTc,IfFdYy,IfFdChr,IfFdNo,IfFdSr        Normal
  '    'sum(IfGrsWt)


  Dim i As Integer
  Select Case UCase(IdName)
  Case Is = UCase("wIfChr")
    If ADC.Mode = xNorm Then
      ADC("wIfChr") = moCn.GetFldVal("Select vPValue from vParam where " + _
                                "vPCoCd = '" + gs_CoCd + "' and vPTyp= 'DEFCHR' " + _
                                "and vPMCd='" + ADC("wIfTc") + "'")
      '("Select PDesc225 From Param where PTyp= 'TC' and PmCd= '" + ADC("wIfTc") + "'")
    End If
'  Case Is = UCase("wIfPrdOrd")
'    If ADC.Mode = xNorm Then
'      If GrdInvFgd.Mode = fgmnorm Then _
'      ADC("wIfPrdOrd") = moCn.GetFldVal("Select FdPrdOdTc+ '/'+ FdPrdOdYy+ '/'+ " + _
'                         "FdPrdOdChr+ '/'+ LTrim(str(FdPrdOdNo))+ '/'+ LTrim(str(FdPrdOdSr)) from Fgd " + _
'                         "where FdTc= '" + GrdInvFgd.Value(RowNum, "IfFdTc") + "' and " + _
'                         "FdYy= '" + GrdInvFgd.Value(RowNum, "IfFdYy") + "' and " + _
'                         "FdChr= '" + GrdInvFgd.Value(RowNum, "IfFdChr") + "' and " + _
'                         "FdNo= " + CStr(GrdInvFgd.Value(RowNum, "IfFdNo")) + " and " + _
'                         "FdSr= " + CStr(GrdInvFgd.Value(RowNum, "IfFdSr")))
'    End If
'  Case Is = UCase("wIfPrdCmCd")
'    If ADC.Mode = xNorm Then
'      If GrdInvFgd.Mode = fgmnorm Then _
'      ADC("wIfPrdCmCd") = moCn.GetFldVal("Select FdPrdCmCd from Fgd where " + _
'                         "FdTc= '" + GrdInvFgd.Value(RowNum, "IfFdTc") + "' and " + _
'                         "FdYy= '" + GrdInvFgd.Value(RowNum, "IfFdYy") + "' and " + _
'                         "FdChr= '" + GrdInvFgd.Value(RowNum, "IfFdChr") + "' and " + _
'                         "FdNo= " + CStr(GrdInvFgd.Value(RowNum, "IfFdNo")) + " and " + _
'                         "FdSr= " + CStr(GrdInvFgd.Value(RowNum, "IfFdSr")))
'    End If
  Case Is = UCase("wIfTotBags")
    Dim wTotBags As Single
    If ADC.Mode = xNorm Then
      If GrdInvFgd.Mode = fgmnorm Or GrdInvFgd.Mode = fgmdel Or GrdInvFgd.Mode = fgmInitOver Or GrdInvFgd.Mode = fgmCopyOver Then
        For i = 1 To GrdInvFgd.Rows - 1
          If Not GrdInvFgd.IsDel(i) Then
            If GrdInvFgd.Value(i, "IfFdBNo") <> 0 Then wTotBags = wTotBags + 1
          End If
        Next i
        ADC("wIfTotBags") = wTotBags
      End If
    End If
  Case Is = UCase("wIfTotQty")
    Dim wTotQty As Single
    If ADC.Mode = xNorm Then
      If GrdInvFgd.Mode = fgmnorm Or GrdInvFgd.Mode = fgmdel Or GrdInvFgd.Mode = fgmInitOver Or GrdInvFgd.Mode = fgmCopyOver Then
        For i = 1 To GrdInvFgd.Rows - 1
          If Not GrdInvFgd.IsDel(i) Then
            wTotQty = wTotQty + GrdInvFgd.Value(i, "IfQty")
          End If
        Next i
        ADC("wIfTotQty") = wTotQty
      End If
    End If
  Case Is = UCase("wIfTotWt")
    Dim wTotWt As Single
    If ADC.Mode = xNorm Then
      If GrdInvFgd.Mode = fgmnorm Or GrdInvFgd.Mode = fgmdel Or GrdInvFgd.Mode = fgmInitOver Or GrdInvFgd.Mode = fgmCopyOver Then
        For i = 1 To GrdInvFgd.Rows - 1
          If Not GrdInvFgd.IsDel(i) Then
            wTotWt = wTotWt + GrdInvFgd.Value(i, "IfGrsWt")
          End If
        Next i
        ADC("wIfTotWt") = wTotWt
      End If
    End If
  End Select
End Sub
Private Sub ATXT_FldChange(Index As Integer, ByVal IdName As String, Cancel As Boolean, ErrMsg As String, ByVal pv_NewValue As Variant)
  '*** wIfTc        valid PMCd('TC', PValue= ctInTcTyp) and the type of entry should be 'IN'
  '*** wIfYy        valid PMCd('YY')
  '*** wIfChr       valid PSCd('CHR', adc(wIfTc)) or ''
  '*** wIfNo        valid InNo from InvHd(InTc= adc(wIfTc), InYy= adc(wIfYy),
      '             InChr= adc(wIfChr)) or Zero
  '*** wIfSrt       valid PMCd('SRTINBAG') or ''
  '*** wSummBagGrp  valid PMCd('INBAGGR') or ''
      '             call the summary routine
  '*** Give validation for all fields in the Scope frame
  '*** Give validation for all the fields in the Copy frame
  
  Select Case UCase(IdName)
  Case Is = UCase("wIfTc")
    Cancel = Not moCn.RecSeek("select PMCd from Param where PTyp= 'TC' and " + _
                         "PMCd = '" + pv_NewValue + "' and " + _
                         "PValue= '" + ms_TcTyp + "' ")
    If Cancel = True Then ErrMsg = "Invalid Invoice Voucher Tc": Exit Sub
  Case Is = UCase("wIfYy")
    Cancel = Not moCn.RecSeek("select PMCd from Param where PTyp= 'YY' and " + _
                         "PMCd = '" + pv_NewValue + "'")
    If Cancel = True Then ErrMsg = "Invalid Invoice Voucher Year": Exit Sub
  Case Is = UCase("wIfChr")
    Cancel = pv_NewValue <> "" And (Not moCn.RecSeek("select vPSCd from vParam where " + _
                         "vPCoCd = '" + gs_CoCd + "' and vPTyp= 'CHR' and " + _
                         "vPMCd = '" + ADC("wIfTc") + "' and vPSCd='" + pv_NewValue + "'"))
' **** Zubin 211 **** '
'''''    ErrMsg = "Invalid Invoice Voucher Character": Exit Sub
    If Cancel = True Then ErrMsg = "Invalid Voucher Character": Exit Sub
    Dim ws_UsrArr() As String, wi_Index As Integer
    ws_UsrArr = Split(moCn.GetFldVal("Select vPValue4 from vParam where vPCoCd= '" + gs_CoCd + "' and " + _
                                        "vPTyp='CHR' and vPMCd='" + ADC("wIfTc") + "' and vPSCd= '" + pv_NewValue + "'"), ",")
    If IsArray(ws_UsrArr) Then
      For wi_Index = 0 To UBound(ws_UsrArr)
        Cancel = IIF(UCase(gs_UsrCd) = UCase(ws_UsrArr(wi_Index)), False, True)
        If Cancel = False Then: Exit For
      Next wi_Index
    End If
   If Cancel = True Then ErrMsg = "Access To This Invoice Voucher Character Is Denied To User " + gs_UsrCd: Exit Sub
' **** Zubin 211 **** '

  Case Is = UCase("wIfNo")
    '****** Sachin 3.02 28-12-07 - Current Partition
    Cancel = pv_NewValue <> 0 And (Not moCn.RecSeek("select InNo from InvHd where " + _
                                                  "InCoCd= '" + gs_CoCd + "' and " + _
                                                  "InTc= '" + ADC("wIfTc") + "' and " + _
                                                  "InYy= '" + ADC("wIfYy") + "' and " + _
                                                  "InChr= '" + ADC("wIfChr") + "' and " + _
                                                  "InNo= " + CStr(pv_NewValue) + " and InPrtKey=(Case When '" + gs_Partition + "'='" + ctCurrPrtn + "' Then '" + ctCurrPrtn + "' Else InPrtKey End) "))
                                                  
    ErrMsg = "Invalid Invoice Voucher Number": Exit Sub
  Case Is = UCase("wIfSrt")
    If pv_NewValue <> "" Then
      Cancel = Not moCn.RecSeek("Select PMCd From Param where PTyp= 'SRTINBAG' " + _
                           "and PMCd = '" + pv_NewValue + "'")
      ErrMsg = "Invalid Sort Option": Exit Sub
    End If
  Case Is = UCase("wSummBagGrp")
    If pv_NewValue <> "" Then
      Cancel = Not moCn.RecSeek("Select PMCd From Param where PTyp= 'INBAGGR' " + _
                                "and PMCd = '" + pv_NewValue + "'")
      If Cancel = True Then ErrMsg = "Invalid Summary Group Option": Exit Sub
      Call Summary(pv_NewValue)
    End If
  Case Is = UCase("wCpPrdCmCd"), UCase("wScPrdCmCd")
    Cancel = pv_NewValue <> "" And Not moCn.RecSeek("Select CmCd From CustMst where " + _
                                            "CmCtg= 'C' and CmCd = '" + pv_NewValue + "'")
     ErrMsg = "Invalid Production Customer Code": Exit Sub
  Case Is = UCase("wCpExpCmCd"), UCase("wScExpCmCd")
    Cancel = pv_NewValue <> "" And Not moCn.RecSeek("Select CmCd From CustMst where " + _
                                            "CmCtg= 'C' and CmCd = '" + pv_NewValue + "'")
     ErrMsg = "Invalid Export Customer Code": Exit Sub
  Case Is = UCase("wCpCtCd")
    Cancel = pv_NewValue <> "" And Not moCn.RecSeek("Select CmCd From CustMst where " + _
                                            "CmCtg= 'T' and CmCd = '" + pv_NewValue + "'")
    ErrMsg = "Invalid Bag Contractor Code": Exit Sub
  Case Is = UCase("wCpPrdOdTc"), UCase("wScPrdOdTc")
    Cancel = pv_NewValue <> "" And Not moCn.RecSeek("select PMCd from Param where PTyp= 'TC' and " + _
                                            "PMCd = '" + pv_NewValue + "' and " + _
                                            "PValue= '" + wSOTcTyp + "' ")
    If Cancel = True Then ErrMsg = "Invalid Production Order Voucher Tc": Exit Sub
  Case Is = UCase("wCpExpOdTc"), UCase("wScExpOdTc")
    Cancel = pv_NewValue <> "" And Not moCn.RecSeek("select PMCd from Param where PTyp= 'TC' and " + _
                                            "PMCd = '" + pv_NewValue + "' and " + _
                                            "PValue= '" + wSOTcTyp + "' ")
    If Cancel = True Then ErrMsg = "Invalid Export Order Voucher Tc": Exit Sub
  Case Is = UCase("wCpPrdOdYy"), UCase("wScPrdOdYy")
    Cancel = pv_NewValue <> "" And Not moCn.RecSeek("select PMCd from Param where PTyp= 'YY' and " + _
                                            "PMCd = '" + pv_NewValue + "'")
    If Cancel = True Then ErrMsg = "Invalid Production Order Voucher Year": Exit Sub
  Case Is = UCase("wCpExpOdYy"), UCase("wScExpOdYy")
    Cancel = pv_NewValue <> "" And Not moCn.RecSeek("select PMCd from Param where PTyp= 'YY' and " + _
                                            "PMCd = '" + pv_NewValue + "'")
    If Cancel = True Then ErrMsg = "Invalid Export Order Voucher Year": Exit Sub
  Case Is = UCase("wCpPrdOdChr")
    Cancel = pv_NewValue <> "" And (Not moCn.RecSeek("select vPSCd from vParam where " + _
                        "vPCoCd= '" + gs_CoCd + "' and vPTyp= 'CHR' and " + _
                        "vPMCd = '" + ADC("wCpPrdOdTc") + "' and vPSCd='" + pv_NewValue + "'"))
    ErrMsg = "Invalid Production Order Voucher Character": Exit Sub
  Case Is = UCase("wScPrdOdChr")
    Cancel = pv_NewValue <> "" And (Not moCn.RecSeek("select vPSCd from vParam where " + _
                         "vPCoCd= '" + gs_CoCd + "' and vPTyp= 'CHR' and " + _
                         "vPMCd = '" + ADC("wScPrdOdTc") + "' and vPSCd='" + pv_NewValue + "'"))
    ErrMsg = "Invalid Production Order Voucher Character": Exit Sub
  Case Is = UCase("wCpExpOdChr")
    Cancel = pv_NewValue <> "" And (Not moCn.RecSeek("select vPSCd from vParam where " + _
                         "vPCoCd= '" + gs_CoCd + "' and vPTyp= 'CHR' and " + _
                         "vPMCd = '" + ADC("wCpExpOdTc") + "' and vPSCd='" + pv_NewValue + "'"))
    ErrMsg = "Invalid Export Order Voucher Character": Exit Sub
  Case Is = UCase("wScExpOdChr")
    Cancel = pv_NewValue <> "" And (Not moCn.RecSeek("select vPSCd from vParam where " + _
                         "vPCoCd= '" + gs_CoCd + "' and vPTyp= 'CHR' and " + _
                         "vPMCd = '" + ADC("wScExpOdTc") + "' and vPSCd='" + pv_NewValue + "'"))
    ErrMsg = "Invalid Export Order Voucher Character": Exit Sub
    
  ' ***** Manali 3.5.0 - 11/11/08 - DmCtg added in scope for copy
  Case Is = UCase("wCpDmCtg")
    Cancel = pv_NewValue <> "" And (Not moCn.RecSeek("Select 'x' from Param where PTyp='DMCTG' and PMCd='" + pv_NewValue + "' "))
    ErrMsg = "Invalid Design Category": Exit Sub
  ' ***** Manali 3.5.0 - 11/11/08 - DmCtg added in scope for copy
    
  Case Is = UCase("wCpDmCd"), UCase("wScDmCd")
    '****** Sachin 3.02 28-12-07
    Cancel = pv_NewValue <> "" And (Not moCn.RecSeek("select DmCd from DsgMst where " + _
                                    "DmTcTyp= 'DM' and DmCd = '" + pv_NewValue + "' and " + _
                                    "DmSz= '' and DmPrtKey='" + ctCurrPrtn + "'"))
    ErrMsg = "Invalid Design Code": Exit Sub
  Case Is = UCase("wCpPlChr"), UCase("wScPlChr")
    Cancel = (pv_NewValue <> "") And (Not moCn.RecSeek("Select PMCd From Param Where PTyp='PLCHR' " + _
                                  "And PMCd = '" + pv_NewValue + "'"))
    ErrMsg = "Invalid Packing List Chr": Exit Sub
  Case Is = UCase("wCpPlYy"), UCase("wScPlYy")
    Cancel = (pv_NewValue <> "") And (Not moCn.RecSeek("Select PMCd From Param Where PTyp= 'Yy' " + _
                                  "And PMCd = '" + pv_NewValue + "'"))
    ErrMsg = "Invalid Packing List Yy": Exit Sub
  Case Is = UCase("wCpInExpNo")
    Cancel = (pv_NewValue <> "") And (Not moCn.RecSeek("Select InExpNo From InvHd Where InCoCd='" + gs_CoCd + "' "))
    ErrMsg = "Invalid Invoice Export Number": Exit Sub
  Case Is = UCase("wCpInExpNo")
    Cancel = (pv_NewValue <> "") And (Not moCn.RecSeek("Select InExpNo From InvHd Where InCoCd='" + gs_CoCd + "'"))
    ErrMsg = "Invalid Invoice Export Number": Exit Sub
  Case Is = UCase("wCpInTc")
    Cancel = pv_NewValue <> "" And Not moCn.RecSeek("select PMCd from Param where PTyp= 'TC' and " + _
                                            "PMCd = '" + pv_NewValue + "' and " + _
                                            "PValue= '" + ctInTcTyp + "' ")
    If Cancel = True Then ErrMsg = "Invalid Invoice Voucher Tc": Exit Sub
  Case Is = UCase("wCpInYy")
    Cancel = pv_NewValue <> "" And Not moCn.RecSeek("select PMCd from Param where PTyp= 'YY' and " + _
                                            "PMCd = '" + pv_NewValue + "'")
    If Cancel = True Then ErrMsg = "Invalid Invoice Voucher Year": Exit Sub
  Case Is = UCase("wCpInChr")
    Cancel = pv_NewValue <> "" And (Not moCn.RecSeek("select vPSCd from vParam where " + _
                            "vPCoCd= '" + gs_CoCd + "' and vPTyp= 'CHR' and " + _
                            "vPMCd = '" + ADC("wCpInTc") + "' and vPSCd='" + pv_NewValue + "'"))
    ErrMsg = "Invalid Invoice Voucher Character": Exit Sub
  Case Is = UCase("wCpFgSubLoc")
    Cancel = (pv_NewValue <> "") And (Not moCn.RecSeek("select vPMCd from vParam where " + _
                                      "vPCoCd= '" + gs_CoCd + "' and vPTyp= 'FGSUBLOC' and " + _
                                      "vPMCd = '" + pv_NewValue + "'"))
    If Cancel = True Then ErrMsg = "Invalid FG Sub Location": Exit Sub
    
  'Sachin 3.11.0
  Case Is = UCase("wBrsGrp")
    Cancel = Not moCn.RecSeek("select PMCd from Param where PTyp= 'TBAGBSGR' and " + _
                         "PMCd = '" + pv_NewValue + "'")
    If Cancel = True Then ErrMsg = "Invalid Group By Option": Exit Sub
    Call RMSummary(pv_NewValue, "N")
    
  'Sachin 3.11.0
  Case Is = UCase("wVrsGrp")
    Cancel = Not moCn.RecSeek("select PMCd from Param where PTyp= 'TBAGBSGR' and " + _
                         "PMCd = '" + pv_NewValue + "'")
    If Cancel = True Then ErrMsg = "Invalid Group By Option": Exit Sub
    Call RMSummary(pv_NewValue, "Y")
    '**** Bhavna added cpy thru bag no
  Case Is = UCase("wCpyThruBagNo")
    If ADC.Mode = xNorm Then
      Cancel = Not moCn.RecSeek("select PMCd from Param where PTyp= 'YN' and PMCd = '" + pv_NewValue + "'")
      ErrMsg = "Enter 'Y' to Copy thru Bag No Or 'N' to Copy thru RfId No ": Exit Sub
    End If
  '6.1_11 validating shipment note scope
  Case Is = UCase("wCpSnTc")
    Cancel = pv_NewValue <> "" And Not moCn.RecSeek("select PMCd from Param where PTyp= 'TC' and " + _
                                            "PMCd = '" + pv_NewValue + "' and " + _
                                            "PValue= 'SN' ")
    If Cancel = True Then ErrMsg = "Invalid Shipment Note Tc": Exit Sub
  Case Is = UCase("wCpSnYy")
    Cancel = pv_NewValue <> "" And Not moCn.RecSeek("select PMCd from Param where PTyp= 'YY' and " + _
                                            "PMCd = '" + pv_NewValue + "'")
    If Cancel = True Then ErrMsg = "Invalid Shipment Note Year": Exit Sub
  Case Is = UCase("wCpSnChr")
    Cancel = pv_NewValue <> "" And (Not moCn.RecSeek("select vPSCd from vParam where " + _
                        "vPCoCd= '" + gs_CoCd + "' and vPTyp= 'CHR' and " + _
                        "vPMCd = '" + ADC("wCpSnTc") + "' and vPSCd='" + pv_NewValue + "'"))
    ErrMsg = "Invalid Shipment Note Character": Exit Sub
            
  End Select
End Sub

Private Sub CmdCpyRfId_Click()
' ***** Manali 3.10.0 - 19/04/12 - Cpy Bags from RfId
Call DispFra(CpyRfId)
End Sub

Private Sub CmdCpyRfIdGo_Click()
' ***** Manali 3.10.0 - 19/04/12 - Cpy Bags from RfId
'vk.4 RfId changed as Bag Alias No.
If ADC("wRfIdSel") = "" Then DispMsg "Enter Bag Alias No. / Bag Nos List", etError: Exit Sub
Call CpyFgBags(True)
Call CmdCpyRfId_Click
End Sub

Private Sub CmdDelAllBags_Click()
 If ADC("WIFPRTKEY") <> ctCurrPrtn Then DispMsg "Cannot Delete Bags of Previous Partition ", etError: Exit Sub

 If moCn.RecSeek("Select 'x' from InvFgd  where IfCoCd= '" + gs_CoCd + "' and IfTc= '" + ADC("wIfTc") + "' and " + _
                  "IfYy=  '" + ADC("wIfYy") + "' and IfChr=  '" + ADC("wIFChr") + "' and " + _
                  "IFNo=  " + CStr(ADC("wIFNo")) + " and IfSr<>0  and IFPrtKey='" + ctCurrPrtn + "' ") Then
    DispMsg "Cannot Delete Bags As Invoice Has Been Done", etError: GrdInvFgd.SetFocus: Exit Sub
 End If

 If GrdInvFgd.Rows <= 0 Then DispMsg "Nothing To Delete.", etInfo: GrdInvFgd.SetFocus: Exit Sub
 If MsgBox("This will Completely DELETE the Invoice. Are U Sure U Want To Continue?", vbYesNo) = vbNo Then Exit Sub
 
 GrdInvFgd.DelAllRecs
 CmdDelAllBags.SetFocus
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
' ***** Manali 3.10.0 - 18/04/12 - Bar Code
If KeyAscii <> vbKeyReturn Then Exit Sub
If UCase(Me.ActiveControl.Name) = UCase("GrdInvFgd") Then
  With Me.ActiveControl
      If .Col = 3 Then
          If .Rows - 1 = .Row Then .Col = 4
      End If
  End With
End If
' ***** Manali 3.10.0 - 18/04/12 - Bar Code

End Sub

Private Sub GrdInvFgd_RowPostValid(ByVal RowNum As Integer, Cancel As Boolean, ErrMsg As String, ChangedColName As String)
Dim wSrcCurCd As String, wTarCurCd As String
With GrdInvFgd
wTarCurCd = moCn.GetFldVal("select OmCmCurCd from bag " + _
                          " join OrdMst on OmCoCd=BCoCd and BOdTc=OmTc and BOdYy=OmYy and BOdChr=OmChr and BOdNo=OmNo " + _
                          " where BCoCd='" + gs_CoCd + "' and BYy='" + .Value(RowNum, "IfFdBYy") + _
                          "' and BChr='" + .Value(RowNum, "IfFdBChr") + "' and BNo=" + CStr(.Value(RowNum, "IfFdBNo")))
                          
wSrcCurCd = moCn.GetFldVal("Select CmCurCd from InvHd join CustMst on CmCtg= '" + ms_CmCtg + "' and CmCd= InCmCd " + _
                          "where InCoCd= '" + gs_CoCd + "' and InTc= '" + ADC("wIfTc") + "' " + _
                          "and InYy= '" + ADC("wIfYy") + "' and InChr= '" + ADC("wIfChr") + "' " + _
                          "and InNo= " + CStr(ADC("wIfNo")) + " and InPrtKey='" + ctCurrPrtn + "' ")
                              
If wSrcCurCd <> wTarCurCd Then
  Cancel = True: ErrMsg = "||Different Currency Code of Bag should not copy"
  Exit Sub
End If
End With

End Sub

Private Sub GrdInvFgd_RowWhen(ByVal RowNum As Integer)
  '*** If New Row then IfSr= 0 and IfSrNo= max(IfSrNo) for the Invoice + 1
  '*** Note that IfSrNo is unique for the whole invoice
  '*** Give values to the display fields wIfPrdOrd and wIfPrdCmCd
  
  With GrdInvFgd
    If .IsNew(RowNum) Then
      .Value(RowNum, "IfSr") = 0
      '##### .Value(RowNum, "IfSrNo") = moCn.AutoGen("InvFgd", "IfSrNo", adc("wIfTc") + adc("wIfYy") + adc("wIfChr") + CStr(adc("wIfNo")))
      '****** Autogen could also have been used for genarating IfSrNo. But it creates as many
             'records in the autogen table as there are records in the InvHd table. This
             'makes the size of the autogen table very large which is not desirable.
      '****** Sachin 3.02 28-12-07
      
      ' ***** Manali 3.5.0 - 11/11/08 - Default values for IfFdByy, IfFdChr
      If RowNum = 1 Then
        .Value(RowNum, "IfFdBYy") = moCn.GetFldVal("Select hYy from Head where HCoCd= '" + gs_CoCd + "' and HCd= '" + ctSelfCmCd + "'")
        .Value(RowNum, "IfFdBChr") = moCn.GetFldVal("Select hBChr from Head where HCoCd= '" + gs_CoCd + "' and HCd= '" + ctSelfCmCd + "'")
      Else
        .Value(RowNum, "IfFdBYy") = .Value(RowNum - 1, "IfFdBYy")
        .Value(RowNum, "IfFdBChr") = .Value(RowNum - 1, "IfFdBChr")
      End If
      ' ***** Manali 3.5.0 - 11/11/08 - Default values for IfFdByy, IfFdChr
      
      .Value(RowNum, "IfPrtKey") = ctCurrPrtn
      .Value(RowNum, "IfSrNo") = moCn.GetFldVal("Select max(IfSrNo) from InvFgd where " + _
                                                "IfCoCd= '" + gs_CoCd + "' and " + _
                                                "IfTc= '" + ADC("wIfTc") + "' and  " + _
                                                "IfYy= '" + ADC("wIfYy") + "' and " + _
                                                "IfChr= '" + ADC("wIfChr") + "' and " + _
                                                "IfNo= " + CStr(ADC("wIfNo")) + " and IfPrtKey='" + ctCurrPrtn + "' ") + 1
         
    End If
    '****** Sachin 3.02 28-12-07 - Corresponding Partition
    ADC("wIfPrdOrd") = moCn.GetFldVal("Select FdPrdOdTc+ '/'+ FdPrdOdYy+ '/'+ " + _
                       "FdPrdOdChr+ '/'+ LTrim(str(FdPrdOdNo))+ '/'+ LTrim(str(FdPrdOdSr)) " + _
                       "from Fgd where FdCoCd= '" + gs_CoCd + "' and " + _
                       "FdTc= '" + GrdInvFgd.Value(RowNum, "IfFdTc") + "' and " + _
                       "FdYy= '" + GrdInvFgd.Value(RowNum, "IfFdYy") + "' and " + _
                       "FdChr= '" + GrdInvFgd.Value(RowNum, "IfFdChr") + "' and " + _
                       "FdNo= " + CStr(GrdInvFgd.Value(RowNum, "IfFdNo")) + " and " + _
                       "FdSr= " + CStr(GrdInvFgd.Value(RowNum, "IfFdSr")) + " and FdPrtKey=(Case When '" + ADC("wIfPrtKey") + "'='" + ctCurrPrtn + "' Then '" + ctCurrPrtn + "' Else FdPrtKey End) ")
    '****** Sachin 3.02 28-12-07 - Corresponding Partition
    ADC("wIfPrdCmCd") = moCn.GetFldVal("Select FdPrdCmCd from Fgd where " + _
                       "FdCoCd= '" + gs_CoCd + "' and " + _
                       "FdTc= '" + GrdInvFgd.Value(RowNum, "IfFdTc") + "' and " + _
                       "FdYy= '" + GrdInvFgd.Value(RowNum, "IfFdYy") + "' and " + _
                       "FdChr= '" + GrdInvFgd.Value(RowNum, "IfFdChr") + "' and " + _
                       "FdNo= " + CStr(GrdInvFgd.Value(RowNum, "IfFdNo")) + " and " + _
                       "FdSr= " + CStr(GrdInvFgd.Value(RowNum, "IfFdSr")) + " and FdPrtKey=(Case When '" + ADC("wIfPrtKey") + "'='" + ctCurrPrtn + "' Then '" + ctCurrPrtn + "' Else FdPrtKey End) ")
  End With
End Sub
Private Sub GrdInvFgd_ColWhen(ByVal RowNum As Integer, ByVal ColName As String, Cancel As Boolean, ErrMsg As String, ByVal Action As MwfCtl.en_fgAction)
  '*** Do not allow the user to edit any field if InCls = 'Y' in InvHd Table
  '*** Do not allow the user to edit any column if record already exists
  '*** Give help for IfFdTc, IfFdYy, IfFdChr, IfFdNo, IfFdSr, IfFdBYy, IfFdBChr, IfFdBNo
  '*** If IfFdBYy<>'' or IfFdBChr<> '' or IfFdBNo<> 0 then do not allow user to edit
      'IfFdTc, IfFdYy, IfFdChr, IfFdNo, IfFdSr
  '*** Do not allow the user to edit the Design, Suffix & Size fields (wIfDmCd, wIfSfx, wIfDmSz)
      'They are working fields
          
  '****** Sachin 3.02 28-12-07
  If ADC("wIfPrtKey") <> "C" Then Cancel = True: ErrMsg = "Cannot Edit Previous Partition Entry": Exit Sub
          
  If mf_DtLocked = True Then Cancel = True: ErrMsg = "Cannot Edit Locked Entry": Exit Sub
  If wInCls = "Y" Then Cancel = True: ErrMsg = "Cannot Edit As Invoice Is Closed": Exit Sub
  
  '****** Sachin 2.14.0 [Next] [07-02-2007] ******
  If GrdInvFgd.Value(RowNum, "IfSr") = 0 Then
    GrdInvFgd.AllowDelete = True
  Else
    GrdInvFgd.AllowDelete = False
  End If
  '****** Sachin 2.14.0 [Next] [07-02-2007] ******
  
  If moCn.RecSeek("Select IfTc From InvFgd where IfCoCd= '" + gs_CoCd + "' and " + _
        "IfTc= '" + ADC("wIfTc") + "' and IfYy= '" + ADC("wIfYy") + "' and " + _
        "IfChr= '" + ADC("wIfChr") + "' and IfNo= " + CStr(ADC("wIfNo")) + " and " + _
        "IfSrNo= " + CStr(GrdInvFgd.Value(RowNum, "IfSrNo")) + " and IfPrtKey='" + ctCurrPrtn + "' ") Then _
        Cancel = True: ErrMsg = "Cannot Edit This Value": Exit Sub
        
  Select Case ColName
  Case Is = UCase("IfFdBYy")
    Call HlpList.PMCd("YY")
  Case Is = UCase("IfFdBChr")
' **** Zubin 211 **** '
'''''    Call HlpList.PMCd("BCHR")
    Call HlpList.PMCd("BCHR", "'P'")
' **** Zubin 211 **** '
  Case Is = UCase("IfFdBNo")
    Call HlpList.BNo(gs_CoCd, GrdInvFgd.Value(RowNum, "IfFdBYy"), GrdInvFgd.Value(RowNum, "IfFdBChr"), "N", ctFgLoc)
  Case Is = UCase("IfFdTc")
    If GrdInvFgd.Value(RowNum, "IfFdBYy") <> "" Or GrdInvFgd.Value(RowNum, "IfFdBChr") <> "" Or _
       GrdInvFgd.Value(RowNum, "IfFdBNo") <> 0 Then _
       Cancel = True: ErrMsg = "Cannot Edit Finished Goods Tc": Exit Sub
    Call HlpList.PMCd("TC", "'FB','FR'")
  Case Is = UCase("IfFdYy")
    If GrdInvFgd.Value(RowNum, "IfFdBYy") <> "" Or GrdInvFgd.Value(RowNum, "IfFdBChr") <> "" Or _
       GrdInvFgd.Value(RowNum, "IfFdBNo") <> 0 Then _
       Cancel = True: ErrMsg = "Cannot Edit Finished Goods Year": Exit Sub
    Call HlpList.PMCd("YY")
  Case Is = UCase("IfFdChr")
    If GrdInvFgd.Value(RowNum, "IfFdBYy") <> "" Or GrdInvFgd.Value(RowNum, "IfFdBChr") <> "" Or _
       GrdInvFgd.Value(RowNum, "IfFdBNo") <> 0 Then _
       Cancel = True: ErrMsg = "Cannot Edit Finished Goods Character": Exit Sub
    Call HlpList.vPSCd(gs_CoCd, "CHR", GrdInvFgd.Value(RowNum, "IfFdTc"))
  Case Is = UCase("IfFdNo")
    If GrdInvFgd.Value(RowNum, "IfFdBYy") <> "" Or GrdInvFgd.Value(RowNum, "IfFdBChr") <> "" Or _
       GrdInvFgd.Value(RowNum, "IfFdBNo") <> 0 Then _
       Cancel = True: ErrMsg = "Cannot Edit Finished Goods Number": Exit Sub
    Call HlpList.FgNo(gs_CoCd, GrdInvFgd.Value(RowNum, "IfFdTc"), GrdInvFgd.Value(RowNum, "IfFdYy"), GrdInvFgd.Value(RowNum, "IfFdChr"))
  Case Is = UCase("IfFdSr")
    If GrdInvFgd.Value(RowNum, "IfFdBYy") <> "" Or GrdInvFgd.Value(RowNum, "IfFdBChr") <> "" Or _
       GrdInvFgd.Value(RowNum, "IfFdBNo") <> 0 Then _
       Cancel = True: ErrMsg = "Cannot Edit Finished Goods Serial": Exit Sub
    Call HlpList.FdSr(gs_CoCd, GrdInvFgd.Value(RowNum, "IfFdTc"), GrdInvFgd.Value(RowNum, "IfFdYy"), GrdInvFgd.Value(RowNum, "IfFdChr"), GrdInvFgd.Value(RowNum, "IfFdNo"))
  Case Is = UCase("wIfDmCd")
    Cancel = True: ErrMsg = "Cannot Edit This Value": Exit Sub
  Case Is = UCase("wIfSfx")
    Cancel = True: ErrMsg = "Cannot Edit This Value": Exit Sub
  Case Is = UCase("wIfDmSz")
    Cancel = True: ErrMsg = "Cannot Edit This Value": Exit Sub
  End Select
End Sub
Private Sub GrdInvFgd_ReCalc(ByVal RowNum As Integer, ByVal ColName As String, ByVal OldVal As Variant, ByVal SourceName As String, ByVal SourceColName As String)
  '### why check for the existence of the record in Fgd first??? (for recalc of IfFdBYy, Chr, No)
  '*** IfFdBYy on IfFdTc,IfFdYy,IfFdChr,IfFdNo,IfFdSr                   Normal
      '   if IfFdSr<> 0 and exists Fgd(FdTc= IfFdTc, FdYy= IfFdYy, FdChr= IfFdChr, FdNo= IfFdNo, FdSr= IfFdSr)
      '   then IfFdBYy= FdBYy from Fgd(FdTc= IfFdTc, FdYy= IfFdYy, FdChr= IfFdChr, FdNo= IfFdNo, FdSr= IfFdSr)
  '*** IfFdBChr on IfFdTc,IfFdYy,IfFdChr,IfFdNo,IfFdSr                  Normal
      '   if IfFdSr<> 0 and exists Fgd(FdTc= IfFdTc, FdYy= IfFdYy, FdChr= IfFdChr, FdNo= IfFdNo, FdSr= IfFdSr)
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
  '*** wIfDmCd on IfFdTc,IfFdYy,IfFdChr,IfFdNo,IfFdSr                   Normal, Init
      'If the mode is Normal then
      '   wIfDmCd= FdDmCd from Fgd(FdTc= IfFdTc, FdYy= IfFdYy, FdChr= IfFdChr, FdNo= IfFdNo, FdSr= IfFdSr)
      'If the mode is Init then
      '   wIfDmCd= FdDmCd from Fgd(FdTc= IfFdTc, FdYy= IfFdYy, FdChr= IfFdChr, FdNo= IfFdNo, FdSr= IfFdSr)
      '   wIfSfx= FdSfx from Fgd(FdTc= IfFdTc, FdYy= IfFdYy, FdChr= IfFdChr, FdNo= IfFdNo, FdSr= IfFdSr)
      '   wIfDmSz= FdDmSz from Fgd(FdTc= IfFdTc, FdYy= IfFdYy, FdChr= IfFdChr, FdNo= IfFdNo, FdSr= IfFdSr)
  '*** wIfSfx on IfFdTc,IfFdYy,IfFdChr,IfFdNo,IfFdSr                    Normal
      'wIfSfx= FdSfx from Fgd(FdTc= IfFdTc, FdYy= IfFdYy, FdChr= IfFdChr, FdNo= IfFdNo, FdSr= IfFdSr)
  '*** wIfDmSz on IfFdTc,IfFdYy,IfFdChr,IfFdNo,IfFdSr                   Normal
      'wIfDmSz= FdDmSz from Fgd(FdTc= IfFdTc, FdYy= IfFdYy, FdChr= IfFdChr, FdNo= IfFdNo, FdSr= IfFdSr)
  '*** IfQty on IfFdTc,IfFdYy,IfFdChr,IfFdNo,IfFdSr                     Normal
      'IfQty= FdQty-FdDespQty from Fgd(FdTc= IfFdTc, FdYy= IfFdYy, FdChr= IfFdChr, FdNo= IfFdNo, FdSr= IfFdSr)
  '*** IfGrsWt on IfFdTc,IfFdYy,IfFdChr,IfFdNo,IfFdSr                   Normal
      'IfGrsWt= FdGrsWt*(FdQty-FdDespQty)/FdQty from Fgd(FdTc= IfFdTc, FdYy= IfFdYy, FdChr= IfFdChr, FdNo= IfFdNo, FdSr= IfFdSr)
  
  
  If ColName <> UCase("wIfDmCd") Then
    If GrdInvFgd.Mode <> fgmnorm Then Exit Sub
  End If
  
  With GrdInvFgd
    Select Case ColName
      Case Is = UCase("IfFdBYy")
        If .Mode = fgmnorm Then
          '****** Sachin 3.02 28-12-07
          If .Value(RowNum, "IfFdSr") <> 0 And moCn.RecSeek("Select FdBYy From Fgd where " + _
                          "FdCoCd= '" + gs_CoCd + "' and " + _
                          "FdTc= '" + .Value(RowNum, "IfFdTc") + "' and " + _
                          "FdYy= '" + .Value(RowNum, "IfFdYy") + "' and " + _
                          "FdChr= '" + .Value(RowNum, "IfFdChr") + "' and " + _
                          "FdNo= " + CStr(.Value(RowNum, "IfFdNo")) + " and " + _
                          "FdSr= " + CStr(.Value(RowNum, "IfFdSr")) + IIF(ADC("wIfPrtKey") = ctCurrPrtn, " and FdPrtKey='" + ctCurrPrtn + "' ", "")) Then _
            .Value(RowNum, "IfFdBYy") = moCn.GetFldVal("Select FdBYy From Fgd where " + _
                                        "FdCoCd= '" + gs_CoCd + "' and " + _
                                        "FdTc= '" + .Value(RowNum, "IfFdTc") + "' and " + _
                                        "FdYy= '" + .Value(RowNum, "IfFdYy") + "' and " + _
                                        "FdChr= '" + .Value(RowNum, "IfFdChr") + "' and " + _
                                        "FdNo= " + CStr(.Value(RowNum, "IfFdNo")) + " and " + _
                                        "FdSr= " + CStr(.Value(RowNum, "IfFdSr")) + IIF(ADC("wIfPrtKey") = ctCurrPrtn, " and FdPrtKey='" + ctCurrPrtn + "' ", ""))
        End If
      Case Is = UCase("IfFdBChr")
        If .Mode = fgmnorm Then
          '****** Sachin 3.02 28-12-07
          If .Value(RowNum, "IfFdSr") <> 0 And moCn.RecSeek("Select FdBChr From Fgd where " + _
                          "FdCoCd= '" + gs_CoCd + "' and " + _
                          "FdTc= '" + .Value(RowNum, "IfFdTc") + "' and " + _
                          "FdYy= '" + .Value(RowNum, "IfFdYy") + "' and " + _
                          "FdChr= '" + .Value(RowNum, "IfFdChr") + "' and " + _
                          "FdNo= " + CStr(.Value(RowNum, "IfFdNo")) + " and " + _
                          "FdSr= " + CStr(.Value(RowNum, "IfFdSr")) + IIF(ADC("wIfPrtKey") = ctCurrPrtn, " and FdPrtKey='" + ctCurrPrtn + "' ", "")) Then _
            .Value(RowNum, "IfFdBChr") = moCn.GetFldVal("Select FdBChr From Fgd where " + _
                                        "FdCoCd= '" + gs_CoCd + "' and " + _
                                        "FdTc= '" + .Value(RowNum, "IfFdTc") + "' and " + _
                                        "FdYy= '" + .Value(RowNum, "IfFdYy") + "' and " + _
                                        "FdChr= '" + .Value(RowNum, "IfFdChr") + "' and " + _
                                        "FdNo= " + CStr(.Value(RowNum, "IfFdNo")) + " and " + _
                                        "FdSr= " + CStr(.Value(RowNum, "IfFdSr")) + IIF(ADC("wIfPrtKey") = ctCurrPrtn, " and FdPrtKey='" + ctCurrPrtn + "' ", ""))
        End If
      Case Is = UCase("IfFdBNo")
        If .Mode = fgmnorm Then
          Select Case UCase(SourceColName)
          Case Is = UCase("IfFdBYy"), UCase("IfFdBChr")
            .Value(RowNum, "IfFdBNo") = 0
          Case Is = UCase("IfFdTc"), UCase("IfFdYy"), UCase("IfFdChr"), UCase("IfFdNo"), UCase("IfFdSr")
            '****** Sachin 3.02 28-12-07
            If .Value(RowNum, "IfFdSr") <> 0 And moCn.RecSeek("Select FdBNo From Fgd where " + _
                            "FdCoCd= '" + gs_CoCd + "' and " + _
                            "FdTc= '" + .Value(RowNum, "IfFdTc") + "' and " + _
                            "FdYy= '" + .Value(RowNum, "IfFdYy") + "' and " + _
                            "FdChr= '" + .Value(RowNum, "IfFdChr") + "' and " + _
                            "FdNo= " + CStr(.Value(RowNum, "IfFdNo")) + " and " + _
                            "FdSr= " + CStr(.Value(RowNum, "IfFdSr")) + IIF(ADC("wIfPrtKey") = ctCurrPrtn, " and FdPrtKey='" + ctCurrPrtn + "' ", "")) Then _
              .Value(RowNum, "IfFdBNo") = moCn.GetFldVal("Select FdBNo From Fgd where " + _
                                          "FdCoCd= '" + gs_CoCd + "' and " + _
                                          "FdTc= '" + .Value(RowNum, "IfFdTc") + "' and " + _
                                          "FdYy= '" + .Value(RowNum, "IfFdYy") + "' and " + _
                                          "FdChr= '" + .Value(RowNum, "IfFdChr") + "' and " + _
                                          "FdNo= " + CStr(.Value(RowNum, "IfFdNo")) + " and " + _
                                          "FdSr= " + CStr(.Value(RowNum, "IfFdSr")) + IIF(ADC("wIfPrtKey") = ctCurrPrtn, " and FdPrtKey='" + ctCurrPrtn + "' ", ""))
          End Select
        End If
      Case Is = UCase("IfFdTc")
        '****** Sachin 3.02 28-12-07
        If .Mode = fgmnorm Then
          If .Value(RowNum, "IfFdBNo") <> 0 Then
            If UCase(ms_TcTyp) = UCase(ctInTcTyp) Then
              .Value(RowNum, "IfFdTc") = moCn.GetFldVal("Select FdTc From Fgd where " + _
                                         "FdCoCd= '" + gs_CoCd + "' and " + _
                                         "FdBYy= '" + .Value(RowNum, "IfFdBYy") + "' and " + _
                                         "FdBChr= '" + .Value(RowNum, "IfFdBChr") + "' and " + _
                                         "FdBNo= " + CStr(.Value(RowNum, "IfFdBNo")) + IIF(ADC("wIfPrtKey") = ctCurrPrtn, " and FdPrtKey='" + ctCurrPrtn + "'", "") + " and " + _
                                         "(Select PValue from Param where PTyp= 'TC' and " + _
                                         "PMCd= FdTc)= 'FB' and Round(FdQty, 1)> Round(FdDespQty, 1) ")
            ElseIf UCase(ms_TcTyp) = UCase(ctCTBTcTyp) Then
              .Value(RowNum, "IfFdTc") = moCn.GetFldVal("Select FdTc From Fgd Fb where " + _
                                         "FdCoCd= '" + gs_CoCd + "' and " + _
                                         "FdBYy= '" + .Value(RowNum, "IfFdBYy") + "' and " + _
                                         "FdBChr= '" + .Value(RowNum, "IfFdBChr") + "' and " + _
                                         "FdBNo= " + CStr(.Value(RowNum, "IfFdBNo")) + IIF(ADC("wIfPrtKey") = ctCurrPrtn, " and FdPrtKey='" + ctCurrPrtn + "'", "") + " and " + _
                                         "(Select PValue from Param where PTyp= 'TC' and " + _
                                         "PMCd= FdTc)= 'FB' and Round(FdQty, 1)> Round(FdCTBDespQty, 1) and " + _
                                         "not exists(Select 'x' from Fgd Fwb where " + _
                                         "(Select PValue from Param where PTyp= 'TC' and " + _
                                         "PMCd= Fwb.FdTc)= 'FWB' and Fwb.FdCoCd= Fb.FdCoCd " + _
                                         "and Fwb.FdRefYy= Fb.FdYy and Fwb.FdRefKey= Fb.FdKey And Fwb.FdPrtKey=Fb.FdPrtKey) ")
            End If
          End If
        End If
      Case Is = UCase("IfFdYy")
        If .Mode = fgmnorm Then
          If .Value(RowNum, "IfFdBNo") <> 0 Then
            If UCase(ms_TcTyp) = UCase(ctInTcTyp) Then
              '****** Sachin 3.02 28-12-07
              .Value(RowNum, "IfFdYy") = moCn.GetFldVal("Select FdYy From Fgd where " + _
                                         "FdCoCd= '" + gs_CoCd + "' and " + _
                                         "FdBYy= '" + .Value(RowNum, "IfFdBYy") + "' and " + _
                                         "FdBChr= '" + .Value(RowNum, "IfFdBChr") + "' and " + _
                                         "FdBNo= " + CStr(.Value(RowNum, "IfFdBNo")) + IIF(ADC("wIfPrtKey") = ctCurrPrtn, " and FdPrtKey='" + ctCurrPrtn + "'", "") + " and " + _
                                         "(Select PValue from Param where PTyp= 'TC' and " + _
                                         "PMCd= FdTc)= 'FB' and Round(FdQty, 1)> Round(FdDespQty, 1) ")
            ElseIf UCase(ms_TcTyp) = UCase(ctCTBTcTyp) Then
              .Value(RowNum, "IfFdYy") = moCn.GetFldVal("Select FdYy From Fgd Fb where " + _
                                         "FdCoCd= '" + gs_CoCd + "' and " + _
                                         "FdBYy= '" + .Value(RowNum, "IfFdBYy") + "' and " + _
                                         "FdBChr= '" + .Value(RowNum, "IfFdBChr") + "' and " + _
                                         "FdBNo= " + CStr(.Value(RowNum, "IfFdBNo")) + IIF(ADC("wIfPrtKey") = ctCurrPrtn, " and FdPrtKey='" + ctCurrPrtn + "'", "") + " and " + _
                                         "(Select PValue from Param where PTyp= 'TC' and " + _
                                         "PMCd= FdTc)= 'FB' and Round(FdQty, 1)> Round(FdCTBDespQty, 1) and " + _
                                         "not exists(Select 'x' from Fgd Fwb where " + _
                                         "(Select PValue from Param where PTyp= 'TC' and " + _
                                         "PMCd= Fwb.FdTc)= 'FWB' and Fwb.FdCoCd= Fb.FdCoCd " + _
                                         "and Fwb.FdRefYy= Fb.FdYy and Fwb.FdRefKey= Fb.FdKey And Fwb.FdPrtKey=Fb.FdPrtKey) ")
            End If
          End If
        End If
      Case Is = UCase("IfFdChr")
        If .Mode = fgmnorm Then
          If .Value(RowNum, "IfFdBNo") <> 0 Then
            If UCase(ms_TcTyp) = UCase(ctInTcTyp) Then
              '****** Sachin 3.02 28-12-07
              .Value(RowNum, "IfFdChr") = moCn.GetFldVal("Select FdChr From Fgd where " + _
                                         "FdCoCd= '" + gs_CoCd + "' and " + _
                                         "FdBYy= '" + .Value(RowNum, "IfFdBYy") + "' and " + _
                                         "FdBChr= '" + .Value(RowNum, "IfFdBChr") + "' and " + _
                                         "FdBNo= " + CStr(.Value(RowNum, "IfFdBNo")) + IIF(ADC("wIfPrtKey") = ctCurrPrtn, " and FdPrtKey='" + ctCurrPrtn + "'", "") + " and " + _
                                         "(Select PValue from Param where PTyp= 'TC' and " + _
                                         "PMCd= FdTc)= 'FB' and Round(FdQty, 1)> Round(FdDespQty, 1) ")
            ElseIf UCase(ms_TcTyp) = UCase(ctCTBTcTyp) Then
              .Value(RowNum, "IfFdChr") = moCn.GetFldVal("Select FdChr From Fgd Fb where " + _
                                         "FdCoCd= '" + gs_CoCd + "' and " + _
                                         "FdBYy= '" + .Value(RowNum, "IfFdBYy") + "' and " + _
                                         "FdBChr= '" + .Value(RowNum, "IfFdBChr") + "' and " + _
                                         "FdBNo= " + CStr(.Value(RowNum, "IfFdBNo")) + IIF(ADC("wIfPrtKey") = ctCurrPrtn, " and FdPrtKey='" + ctCurrPrtn + "'", "") + " and " + _
                                         "(Select PValue from Param where PTyp= 'TC' and " + _
                                         "PMCd= FdTc)= 'FB' and Round(FdQty, 1)> Round(FdCTBDespQty, 1) and " + _
                                         "not exists(Select 'x' from Fgd Fwb where " + _
                                         "(Select PValue from Param where PTyp= 'TC' and " + _
                                         "PMCd= Fwb.FdTc)= 'FWB' and Fwb.FdCoCd= Fb.FdCoCd " + _
                                         "and Fwb.FdRefYy= Fb.FdYy and Fwb.FdRefKey= Fb.FdKey and Fwb.FdPrtKey=Fb.FdPrtKey) ")
            End If
          End If
        End If
      Case Is = UCase("IfFdNo")
        If .Mode = fgmnorm Then
          Select Case UCase(SourceColName)
          Case Is = UCase("IfFdBYy"), UCase("IfFdBChr"), UCase("IfFdBNo")
            If .Value(RowNum, "IfFdBNo") <> 0 Then
              If UCase(ms_TcTyp) = UCase(ctInTcTyp) Then
                .Value(RowNum, "IfFdNo") = moCn.GetFldVal("Select FdNo From Fgd where " + _
                                         "FdCoCd= '" + gs_CoCd + "' and " + _
                                         "FdBYy= '" + .Value(RowNum, "IfFdBYy") + "' and " + _
                                         "FdBChr= '" + .Value(RowNum, "IfFdBChr") + "' and " + _
                                         "FdBNo= " + CStr(.Value(RowNum, "IfFdBNo")) + IIF(ADC("wIfPrtKey") = ctCurrPrtn, " and FdPRtKey='" + ctCurrPrtn + "'", "") + " and " + _
                                         "(Select PValue from Param where PTyp= 'TC' and " + _
                                         "PMCd= FdTc)= 'FB' and Round(FdQty, 1)> Round(FdDespQty, 1) ")
              ElseIf UCase(ms_TcTyp) = UCase(ctCTBTcTyp) Then
                .Value(RowNum, "IfFdNo") = moCn.GetFldVal("Select FdNo From Fgd Fb where " + _
                                           "FdCoCd= '" + gs_CoCd + "' and " + _
                                           "FdBYy= '" + .Value(RowNum, "IfFdBYy") + "' and " + _
                                           "FdBChr= '" + .Value(RowNum, "IfFdBChr") + "' and " + _
                                           "FdBNo= " + CStr(.Value(RowNum, "IfFdBNo")) + IIF(ADC("wIfPrtKey") = ctCurrPrtn, " and FdPrtKey='" + ctCurrPrtn + "'", "") + " and " + _
                                           "(Select PValue from Param where PTyp= 'TC' and " + _
                                           "PMCd= FdTc)= 'FB' and Round(FdQty, 1)> Round(FdCTBDespQty, 1) and " + _
                                           "not exists(Select 'x' from Fgd Fwb where " + _
                                           "(Select PValue from Param where PTyp= 'TC' and " + _
                                           "PMCd= Fwb.FdTc)= 'FWB' and Fwb.FdCoCd= Fb.FdCoCd " + _
                                           "and Fwb.FdRefYy= Fb.FdYy and Fwb.FdRefKey= Fb.FdKey and Fwb.FdPrtKey=Fb.FdPrtKey) ")
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
              If UCase(ms_TcTyp) = UCase(ctInTcTyp) Then
                '****** Sachin 3.02 28-12-07
                .Value(RowNum, "IfFdSr") = moCn.GetFldVal("Select FdSr From Fgd where " + _
                                         "FdCoCd= '" + gs_CoCd + "' and " + _
                                         "FdBYy= '" + .Value(RowNum, "IfFdBYy") + "' and " + _
                                         "FdBChr= '" + .Value(RowNum, "IfFdBChr") + "' and " + _
                                         "FdBNo= " + CStr(.Value(RowNum, "IfFdBNo")) + IIF(ADC("wIfPrtKey") = ctCurrPrtn, " and FdPrtKey='" + ctCurrPrtn + "'", "") + " and " + _
                                         "(Select PValue from Param where PTyp= 'TC' and " + _
                                         "PMCd= FdTc)= 'FB' and Round(FdQty, 1)> Round(FdDespQty, 1) ")
              ElseIf UCase(ms_TcTyp) = UCase(ctCTBTcTyp) Then
                .Value(RowNum, "IfFdSr") = moCn.GetFldVal("Select FdSr From Fgd Fb where " + _
                                           "FdCoCd= '" + gs_CoCd + "' and " + _
                                           "FdBYy= '" + .Value(RowNum, "IfFdBYy") + "' and " + _
                                           "FdBChr= '" + .Value(RowNum, "IfFdBChr") + "' and " + _
                                           "FdBNo= " + CStr(.Value(RowNum, "IfFdBNo")) + IIF(ADC("wIfPrtKey") = ctCurrPrtn, " and FdPrtKey='" + ctCurrPrtn + "'", "") + " and " + _
                                           "(Select PValue from Param where PTyp= 'TC' and " + _
                                           "PMCd= FdTc)= 'FB' and Round(FdQty, 1)> Round(FdCTBDespQty, 1) and " + _
                                           "not exists(Select 'x' from Fgd Fwb where " + _
                                           "(Select PValue from Param where PTyp= 'TC' and " + _
                                           "PMCd= Fwb.FdTc)= 'FWB' and Fwb.FdCoCd= Fb.FdCoCd " + _
                                           "and Fwb.FdRefYy= Fb.FdYy and Fwb.FdRefKey= Fb.FdKey and Fwb.FdPrtKey=Fb.FdPrtKey) ")
              End If
            End If
          Case Is = UCase("IfFdTc"), UCase("IfFdYy"), UCase("IfFdChr"), UCase("IfFdNo")
            .Value(RowNum, "IfFdSr") = 0
          End Select
        End If
      Case Is = UCase("wIfDmCd")
        '*** This recalc is not done during copy(i.e. when Mode = fgmCopy) as copy becomes slow
        If .Mode = fgmnorm Then
          '****** Sachin 3.02 28-12-07
          .Value(RowNum, "wIfDmCd") = moCn.GetFldVal("Select FdDmCd From Fgd where " + _
                                      "FdCoCd= '" + gs_CoCd + "' and " + _
                                      "FdTc= '" + .Value(RowNum, "IfFdTc") + "' and " + _
                                      "FdYy= '" + .Value(RowNum, "IfFdYy") + "' and " + _
                                      "FdChr= '" + .Value(RowNum, "IfFdChr") + "' and " + _
                                      "FdNo= " + CStr(.Value(RowNum, "IfFdNo")) + " and " + _
                                      "FdSr= " + CStr(.Value(RowNum, "IfFdSr")) + IIF(ADC("wIfPrtKey") = ctCurrPrtn, " and FdPrtKey='" + ctCurrPrtn + "' ", ""))
        End If
      
        If .Mode = fgminit Then
          Dim wVal As String, wArr() As String
          '****** Sachin 3.02 28-12-07
          wVal = moCn.GetFldVal("Select FdDmCd+'|'+FdDmSz+'|'+FdSfx From Fgd where " + _
                                "FdCoCd= '" + gs_CoCd + "' and " + _
                                "FdTc= '" + .Value(RowNum, "IfFdTc") + "' and " + _
                                "FdYy= '" + .Value(RowNum, "IfFdYy") + "' and " + _
                                "FdChr= '" + .Value(RowNum, "IfFdChr") + "' and " + _
                                "FdNo= " + CStr(.Value(RowNum, "IfFdNo")) + " and " + _
                                "FdSr= " + CStr(.Value(RowNum, "IfFdSr")) + IIF(ADC("wIfPrtKey") = ctCurrPrtn, " and FdPrtKey='" + ctCurrPrtn + "' ", ""))
          wArr = Split(wVal, "|")
          If UBound(wArr) = 2 Then
            .Value(RowNum, "wIfDmCd") = wArr(0)
            .Value(RowNum, "wIfDmSz") = wArr(1)
            .Value(RowNum, "wIfSfx") = wArr(2)
          End If
        End If
      Case Is = UCase("wIfSfx")
        '*** This recalc is not done during copy or Init (i.e. when Mode = fgmCopy or Mode = fgminit) as copy becomes slow and the find option also becomes slow
        If .Mode = fgmnorm Then
           '****** Sachin 3.02 28-12-07
           .Value(RowNum, "wIfSfx") = moCn.GetFldVal("Select FdSfx From Fgd where " + _
                                      "FdCoCd= '" + gs_CoCd + "' and " + _
                                      "FdTc= '" + .Value(RowNum, "IfFdTc") + "' and " + _
                                      "FdYy= '" + .Value(RowNum, "IfFdYy") + "' and " + _
                                      "FdChr= '" + .Value(RowNum, "IfFdChr") + "' and " + _
                                      "FdNo= " + CStr(.Value(RowNum, "IfFdNo")) + " and " + _
                                      "FdSr= " + CStr(.Value(RowNum, "IfFdSr")) + IIF(ADC("wIfPrtKey") = ctCurrPrtn, " and FdPrtKey='" + ctCurrPrtn + "' ", ""))
        End If
      Case Is = UCase("wIfDmSz")
        '*** This recalc is not done during copy or Init (i.e. when Mode = fgmCopy or Mode = fgminit) as copy becomes slow
        If .Mode = fgmnorm Then
          '****** Sachin 3.02 28-12-07
          .Value(RowNum, "wIfDmSz") = moCn.GetFldVal("Select FdDmSz From Fgd where " + _
                                      "FdCoCd= '" + gs_CoCd + "' and " + _
                                      "FdTc= '" + .Value(RowNum, "IfFdTc") + "' and " + _
                                      "FdYy= '" + .Value(RowNum, "IfFdYy") + "' and " + _
                                      "FdChr= '" + .Value(RowNum, "IfFdChr") + "' and " + _
                                      "FdNo= " + CStr(.Value(RowNum, "IfFdNo")) + " and " + _
                                      "FdSr= " + CStr(.Value(RowNum, "IfFdSr")) + IIF(ADC("wIfPrtKey") = ctCurrPrtn, " and FdPrtKey='" + ctCurrPrtn + "' ", ""))
        End If
      Case Is = UCase("IfQty")
        If .Mode = fgmnorm Then
          '****** Sachin 3.02 28-12-07
          .Value(RowNum, "IfQty") = moCn.GetFldVal("Select FdQty- " + _
                                    IIF(UCase(ms_TcTyp) = UCase(ctInTcTyp), "FdDespQty", "FdCTBDespQty") + " " + _
                                    "From Fgd where FdCoCd= '" + gs_CoCd + "' and " + _
                                    "FdTc= '" + .Value(RowNum, "IfFdTc") + "' and " + _
                                    "FdYy= '" + .Value(RowNum, "IfFdYy") + "' and " + _
                                    "FdChr= '" + .Value(RowNum, "IfFdChr") + "' and " + _
                                    "FdNo= " + CStr(.Value(RowNum, "IfFdNo")) + " and " + _
                                    "FdSr= " + CStr(.Value(RowNum, "IfFdSr")) + IIF(ADC("wIfPrtKey") = ctCurrPrtn, " and FdPrtKey='" + ctCurrPrtn + "' ", ""))
        End If
      Case Is = UCase("IfGrsWt")
        If .Mode = fgmnorm Then
          '****** Sachin 3.02 28-12-07
          .Value(RowNum, "IfGrsWt") = moCn.GetFldVal("Select " + _
                "(case when FdQty> 0 then FdGrWt*(FdQty- " + _
                IIF(UCase(ms_TcTyp) = UCase(ctInTcTyp), "FdDespQty", "FdCTBDespQty") + " " + _
                ")/FdQty else 0 end) From Fgd where FdCoCd= '" + gs_CoCd + "' and " + _
                "FdTc= '" + .Value(RowNum, "IfFdTc") + "' and " + _
                "FdYy= '" + .Value(RowNum, "IfFdYy") + "' and " + _
                "FdChr= '" + .Value(RowNum, "IfFdChr") + "' and " + _
                "FdNo= " + CStr(.Value(RowNum, "IfFdNo")) + " and " + _
                "FdSr= " + CStr(.Value(RowNum, "IfFdSr")) + IIF(ADC("wIfPrtKey") = ctCurrPrtn, " and FdPrtKey='" + ctCurrPrtn + "' ", ""))
        End If
        
      '****** Sachin 3.02 - Id fields In Joins
      '****** Sachin 3.02 - Shifted to InitKey
      '     Case Is = UCase("IfBIdNo")
      '        If .Mode = fgmnorm Then _
      '          .Value(RowNum, "IfBIdNo") = moCn.GetFldVal("Select BIdNo From Bag " + _
      '                                                     " Where BCoCd='" + gs_CoCd + "' And BYy='" + .Value(RowNum, "IfFdBYy") + "' And BChr='" + .Value(RowNum, "IfFdBChr") + "' And BNo=" + CStr(.Value(RowNum, "IfFdBNo")))
      '
      '     Case Is = UCase("IfFdIdNo")
      '        If .Mode = fgmnorm Then _
      '          .Value(RowNum, "IfFdIdNo") = moCn.GetFldVal("Select FdIdNo From Fgd " + _
      '                                                     " Where FdCoCd='" + gs_CoCd + "' And FdTc='" + .Value(RowNum, "IfFdTc") + "' And FdYy='" + .Value(RowNum, "IfFdYy") + "' And FdChr='" + .Value(RowNum, "IfFdChr") + "' And FdNo=" + CStr(.Value(RowNum, "IfFdNo")))
      '****** Sachin 3.02 - Id fields In Joins

    End Select
  End With
End Sub
Private Sub GrdInvFgd_ColChange(ByVal RowNum As Integer, ByVal ColName As String, ByVal NewValue As Variant, Cancel As Boolean, ErrMsg As String)
  '*** IfFdBYy        valid PMCd('YY') or Blank
  '*** IfFdBChr       valid PMCd('BCHR')  or Blank
  '*** IfFdBNo        valid FdBNo from Fgd(FdBYy= IfFdBYy, FdBChr= IfFdBChr, FdQty-FdDespQty> 0) or Zero
  '*** IfFdTc         (valid PMCd('TC', PValue= 'FB' or 'FR') and the type of entry should be 'FgBag' or 'FgRm') or Blank
  '*** IfFdYy         valid PMCd('YY') or Blank
  '*** IfFdChr        valid PSCd('CHR', IfFdTc) or ''
  '*** IfFdNo         valid FgNo from Fg(FgTc= IfFdTc, FgYy= IfFdYy, FgChr= IfFdChr) or Zero
  '*** IfFdSr         (valid FdSr from Fgd(FdTc= IfFdTc, FdYy= IfFdYy, FdChr= IfFdChr, FdNo= IfFdNo, FdQty-FdDespQty> 0)) or Zero
  
  Dim ws_FGBNo() As String, ws_BagNoBarCd As String
  Dim ws_OldYy As String, ws_OldChr As String, ws_OldNo As Single
  
  With GrdInvFgd
    Select Case UCase(ColName)
    Case Is = UCase("IfFdBYy")
      Cancel = (NewValue <> "") And (Not moCn.RecSeek("Select PMCd From Param where " + _
                                    "PTyp= 'YY' and PMCd = '" + NewValue + "'"))
      ErrMsg = "Invalid Bag Year": Exit Sub
    Case Is = UCase("IfFdBChr")
' **** Zubin 211 **** '
'''''      Cancel = (NewValue <> "") And (Not moCn.RecSeek("Select PMCd From Param where " + _
                                    "PTyp= 'BCHR' and PMCd = '" + NewValue + "'"))
      Cancel = (NewValue <> "") And (Not moCn.RecSeek("Select PMCd From Param where " + _
                                    "PTyp= 'BCHR' and PMCd = '" + NewValue + "' and PValue= 'P'"))
' **** Zubin 211 **** '
      ErrMsg = "Invalid Bag Character": Exit Sub
    Case Is = UCase("IfFdBNo")
      '****** Sachin 3.02 28-12-07
      Cancel = (NewValue <> 0) And (Not moCn.RecSeek("Select BNo From Bag where BOdTc='SO' and " + _
                                    "BCoCd= '" + gs_CoCd + "' and " + _
                                    "BYy= '" + .Value(RowNum, "IfFdBYy") + "' and " + _
                                    "BChr= '" + .Value(RowNum, "IfFdBChr") + "' and " + _
                                    "BNo= " + CStr(NewValue) + " and BPrtKey='" + ctCurrPrtn + "' and BLoc= '" + ctFgLoc + "' "))
      ErrMsg = "Invalid Bag Number/ Bag Not In FG Loc/ Bag not belong to SO": Exit Sub
    Case Is = UCase("wIfFdBagNo")
      If NewValue <> "" Then
        'uni.156 calling function to ignore CoCd if entered
        ws_BagNoBarCd = ParseBagNoBarCd(NewValue)
        ws_FGBNo = Split(ws_BagNoBarCd, "/")
        If UBound(ws_FGBNo) <> 2 Then Cancel = True: ErrMsg = "Invalid Bag No": Exit Sub
        If Not IsNumeric(ws_FGBNo(2)) Then Cancel = True: ErrMsg = "Invalid Bag No": Exit Sub
        
        ws_OldYy = .Value(RowNum, "IfFDBYY")
        ws_OldChr = .Value(RowNum, "IfFDBCHR")
        ws_OldNo = .Value(RowNum, "IfFDBNO")
        
        .Value(RowNum, "IfFDBYY") = UCase(ws_FGBNo(0))
        Call GrdInvFgd_ColChange(RowNum, "IFFDBYY", ws_FGBNo(0), Cancel, "")
        If Cancel = True Then ErrMsg = "Invalid Bag Year": .Value(RowNum, "IFFDBYY") = ws_OldYy: Exit Sub
        
        .Value(RowNum, "IFFDBCHR") = UCase(ws_FGBNo(1))
        Call GrdInvFgd_ColChange(RowNum, "IFFDBCHR", ws_FGBNo(1), Cancel, "")
        If Cancel = True Then ErrMsg = "Invalid Bag Character / Flute or Component Bags Cannot Be Taken To FG": .Value(RowNum, "IFFDBCHR") = ws_OldChr: Exit Sub
        
        .Value(RowNum, "IFFDBNO") = ws_FGBNo(2)
        Call GrdInvFgd_ColChange(RowNum, "IFFDBNO", ws_FGBNo(2), Cancel, "")
        If Cancel = True Then ErrMsg = "Invalid Bag Number": .Value(RowNum, "IFFDBNO") = ws_OldNo: Exit Sub

        If .Rows - 1 = RowNum Then
          .AddItem
          .Row = RowNum + 1
          .Col = 3
        End If
      End If
    
    Case Is = UCase("IfFdTc")
      Cancel = (NewValue <> "") And (Not moCn.RecSeek("select PMCd from Param where " + _
                           "PTyp= 'TC' and PMCd = '" + NewValue + "' and " + _
                           "PValue in ('FB', 'FR')"))
      If Cancel = True Then ErrMsg = "Invalid Finished Goods Voucher Tc": Exit Sub
    Case Is = UCase("IfFdYy")
      Cancel = (NewValue <> "") And (Not moCn.RecSeek("select PMCd from Param where " + _
                           "PTyp= 'YY' and PMCd = '" + NewValue + "'"))
      If Cancel = True Then ErrMsg = "Invalid Finished Goods Voucher Year": Exit Sub
    Case Is = UCase("IfFdChr")
      Cancel = (NewValue <> "") And (Not moCn.RecSeek("select vPSCd from vParam where " + _
               "vPCoCd= '" + gs_CoCd + "' and vPTyp= 'CHR' and " + _
               "vPMCd = '" + .Value(RowNum, "IfFdTc") + "' and vPSCd='" + NewValue + "'"))
      If Cancel = True Then ErrMsg = "Invalid Finished Goods Voucher Character": Exit Sub
    Case Is = UCase("IfFdNo")
      '****** Sachin 3.02 28-12-07
      Cancel = (NewValue <> 0) And (Not moCn.RecSeek("Select FgNo From Fg where " + _
                                    "FgCoCd= '" + gs_CoCd + "' and " + _
                                    "FgTc= '" + .Value(RowNum, "IfFdTc") + "' and " + _
                                    "FgYy= '" + .Value(RowNum, "IfFdYy") + "' and " + _
                                    "FgChr= '" + .Value(RowNum, "IfFdChr") + "' and " + _
                                    "FgNo= " + CStr(NewValue) + " and FgPrtKey='" + ctCurrPrtn + "' "))
      If Cancel = True Then ErrMsg = "Invalid Finished Goods Voucher Number": Exit Sub
    Case Is = UCase("IfFdSr")
      '****** Sachin 3.02 28-12-07
      Cancel = (NewValue <> 0) And (Not moCn.RecSeek("Select FdSr From Fgd where " + _
                                    "FdCoCd= '" + gs_CoCd + "' and " + _
                                    "FdTc= '" + .Value(RowNum, "IfFdTc") + "' and " + _
                                    "FdYy= '" + .Value(RowNum, "IfFdYy") + "' and " + _
                                    "FdChr= '" + .Value(RowNum, "IfFdChr") + "' and " + _
                                    "FdNo= " + CStr(.Value(RowNum, "IfFdNo")) + " and " + _
                                    "FdSr= " + CStr(NewValue) + " and FdPrtKey='" + ctCurrPrtn + "' "))
      If Cancel = True Then ErrMsg = "Invalid Finished Goods Voucher Serial": Exit Sub
    End Select
  End With
End Sub
Private Sub GrdInvFgd_InitKey(ByVal RowNum As Integer)
  '*** Initialise Key Fields (i.e. IfTc, IfYy, IfChr, IfNo)
  With GrdInvFgd
    .Store "IfCoCd", gs_CoCd
    .Store "IfTc", ADC("wIfTc")
    .Store "IfYy", ADC("wIfYy")
    .Store "IfChr", ADC("wIfChr")
    .Store "IfNo", ADC("wIfNo")
    
    '****** Sachin 3.02 - Initializing Id fields
        .Store "IfBIdNo", moCn.GetFldVal("Select BIdNo From Bag " + _
                                                     " Where BCoCd='" + gs_CoCd + "' And BYy='" + .Value(RowNum, "IfFdBYy") + "' And BChr='" + .Value(RowNum, "IfFdBChr") + "' And BNo=" + CStr(.Value(RowNum, "IfFdBNo")) + IIF(ADC("wIfPrtKey") = ctCurrPrtn, " and BPrtKey='" + ctCurrPrtn + "' ", ""))

        .Store "IfFdIdNo", moCn.GetFldVal("Select FdIdNo From Fgd " + _
                                                     " Where FdCoCd='" + gs_CoCd + "' And FdTc='" + .Value(RowNum, "IfFdTc") + "' And FdYy='" + .Value(RowNum, "IfFdYy") + "' And FdChr='" + .Value(RowNum, "IfFdChr") + "' And FdNo=" + CStr(.Value(RowNum, "IfFdNo")) + " And FdSr=" + CStr(.Value(RowNum, "IfFdSr")) + IIF(ADC("wIfPrtKey") = ctCurrPrtn, " and FdPrtKey='" + ctCurrPrtn + "' ", ""))
   
    If .IsNew(RowNum) Then _
        .Store "IfInIdNo", moCn.GetFldVal("Select InIdNo From InvHd " + _
                                          " Where InCoCd='" + gs_CoCd + "' And InTc='" + ADC("wIfTc") + "' And  InYy='" + ADC("wIfYy") + "' And InChr='" + ADC("wIfChr") + "' And InNo=" + CStr(ADC("wIfNo")) + " and InPrtKey='" + ctCurrPrtn + "' ")
    '****** Sachin 3.02 - Initializing Id fields
  End With
End Sub
Private Sub ADC_SetRecSource()
  '*** Since this is an 'A' Entry, there is no RecSource associated with ADC. So set the
      'DataMode property of ADC control as xdatUnbound
  ADC.DataMode = xdatUnbound
End Sub
Private Sub GrdInvFgd_SetRecSource()
  '*** Set the Record Source of the Grid GrdInvFgd
      'Default sort option is IfSr- 'IfTc, IfYy, IfChr, IfNo, IfSrNo' (normal full key)
      'Other Sort Options :
      '   IfSr    - IfSr, IfSrNo (normal full key)
      '   IfSrno  - IfSrNo (Key w/o IfSr)
      '   BagNo   - BagNo, IfSrNo (though Bag No is Unique)
      '   DsgCd   - DsgCd, IfSrNo
      '   PrdCmCd - PrdCmCd, IfSrNo
      '   ExpCmCd - ExpCmCd, IfSrNo
      '   PrdOrdNo- PrdOrdNo, IfSrNo
      '   ExpOrdNo- ExpOrdNo, IfSrNo
      '   FdVchNo - FdVchNo, IfSrNo

  Dim wCnd As String, wScCnd As String, ws_OrdBy As String, wInCnd As String
  
  wCnd = IIF(ADC("wScIdSr") = 0, "", " and IfSr= " + CStr(ADC("wScIdSr")) + " ")
  wScCnd = ADC.RepCond
  wScCnd = IIF(wScCnd <> "", " and " + wScCnd, "")
  
  '****** Sachin 3.02 - Id fields in Joins
  wScCnd = IIF(wScCnd = "", "", " and exists(Select 'x' From Fgd Join Fg on " + IIF(gs_Partition = ctCurrPrtn, " FgPrtKey=FdPrtKey and ", "") + " FgIdNo=FdFgIdNo and FgCoCd= FdCoCd and " + _
                        "FgTc= FdTc and FgYy= FdYy and FgChr= FdChr and FgNo= FdNo Where " + _
                        "FdCoCd= IfCoCd and FdTc= IfFdTc and FdYy= IfFdYy and FdChr= IfFdChr " + _
                        "and FdNo= IfFdNo and FdSr= IfFdSr " + wScCnd + " ) ")
  '****** Sachin 3.02.0 ******
  wInCnd = IIF(gs_Partition = ctCurrPrtn, " And Exists (Select 'x' From InvHd Where InPrtKey=IfPrtKey And InCoCd=IfCoCd And InTc=IfTc And InYy=IfYy And InChr=IfChr And InNo=IfNo And InDt>'" + ms_PurgeDt + "' ) ", "")
                        
  If ADC("wIfSrt") = "" Then
    '*** (Bef 2.06) ws_OrdBy = " Order By IfTc, IfYy, IfChr, IfNo, IfSr, IfSrNo "
    ws_OrdBy = " Order By IfTc, IfYy, IfChr, IfNo, IfSrNo "
  ElseIf ADC("wIfSrt") = "DSGODSR" Then
    ws_OrdBy = "Order By (Select (BOdDmCd+BOdTc+BOdYy+BOdChr+str(BOdNo,10)+str(BOdSr,10)) " + _
                "from Bag Where " + IIF(gs_Partition = ctCurrPrtn, " BPrtKey=IfPrtKey and ", "") + " BIdNo=IfBIdNo and BCoCd= IfCoCd and BYy= IfFdBYy and BChr= IfFdBChr and BNo= IfFdBNo)"
  ElseIf ADC("wIfSrt") = "DSGBAG" Then
    ws_OrdBy = "Order By (Select (BOdDmCd+BYy+BChr+str(BNo,10)) from Bag " + _
                "Where " + IIF(gs_Partition = ctCurrPrtn, " BPrtKey=IfPrtKey and ", "") + " BIdNo=IfBIdNo and BCoCd= IfCoCd and BYy= IfFdBYy and BChr= IfFdBChr and BNo= IfFdBNo)"
  ElseIf ADC("wIfSrt") = "DSGFGSR" Then
    ws_OrdBy = "Order By (Select (FdDmCd+FdTc+FdYy+FdChr+ str(FdNo,10)+ str(FdSr,10)) " + _
                "from Fgd Where " + IIF(gs_Partition = ctCurrPrtn, " FdPrtKey=IfPrtKey and ", "") + " FdIdNo=IfFdIdNo and FdCoCd= IfCoCd and FdTc= IfFdTc and FdYy= IfFdYy " + _
                "and FdChr= IfFdChr and FdNo= IfFdNo and FdSr= IfFdSr)"
  ElseIf ADC("wIfSrt") = "ODSRBAG" Then
    ws_OrdBy = "Order By (Select (BOdTc+ BOdYy+ BOdChr+ str(BOdNo,10)+ str(BOdSr,10)) " + _
                "from Bag Where " + IIF(gs_Partition = ctCurrPrtn, " BPrtKey=IfPrtKey and ", "") + " BIdNo=IfBIdNo and BCoCd= IfCoCd and BYy= IfFdBYy and BChr= IfFdBChr and BNo= IfFdBNo), " + _
                "(Select (BYy+BChr+str(BNo,10)) from Bag " + _
                "Where " + IIF(gs_Partition = ctCurrPrtn, " BPrtKey=IfPrtKey and ", "") + " BIdNo=IfBIdNo and BCoCd= IfCoCd and BYy= IfFdBYy and BChr= IfFdBChr and BNo= IfFdBNo)"
  Else
    ws_OrdBy = "Order By " + moCn.GetFldVal("Select PDesc225 from Param " + _
             "where PTyp='SRTINBAG' and PMCd='" + ADC("wIfSrt") + "'")
  End If

  '*** tables Fgd & Fg have been joined only to select the recs in the scope given by user
  '(index= xInvFgdSrNo)
  '****** Sachin 3.02 28-12-07 - PrtKey added
  GrdInvFgd.RecSource = " Select * From InvFgd Where IfCoCd= '" + gs_CoCd + "' and " + _
                        "IfTc= '" + ADC("wIfTc") + "' and IfYy= '" + ADC("wIfYy") + "' and " + _
                        "IfChr= '" + ADC("wIfChr") + "' and IfNo= " + CStr(ADC("wIfNo")) + " and IfPrtKey=(Case When '" + gs_Partition + "'='" + ctCurrPrtn + "' Then '" + ctCurrPrtn + "' Else IfPrtKey End) " + _
                        wCnd + wScCnd + wInCnd + ws_OrdBy
End Sub
Private Sub GrdSumm_ColWhen(ByVal RowNum As Integer, ByVal ColName As String, Cancel As Boolean, ErrMsg As String, ByVal Action As MwfCtl.en_fgAction)
  '*** Cannot edit Summary grid values
  Cancel = True
End Sub
Private Sub CmdCP_Click()
  '*** Display or hide the Copy frame depending on the toggle satate
  
  If mf_DtLocked = True Then Exit Sub
  Call DispFra(Cpy)
  If FraCP.Visible = True Then
    If UCase(ms_TcTyp) = UCase(ctInTcTyp) Then ADC("wCpPrdCmCd") = ADC("wIfCmCd")
    If ADC("wCpFgSubLoc") = "" Then ADC("wCpFgSubLoc") = moCn.GetFldVal("Select hFgSubLoc from Head where HCoCd= '" + gs_CoCd + "' and HCd= '" + ctSelfCmCd + "'")
  End If
End Sub
Private Sub CmdCpGo_Click()
  '*** Check if the values filled in the copy frame fields are valid
  '*** Call the Copy Routine and Close the Copy Frame
  If FraCP.Visible = True Then
    If ADC("wCpPrdOdNo") <> 0 Then
      '****** Sachin 3.02 28-12-07
      If Not moCn.RecSeek("select OmNo from OrdMst where OmCoCd= '" + gs_CoCd + "' and " + _
           "OmTc= '" + ADC("wCpPrdOdTc") + "' and OmYy= '" + ADC("wCpPrdOdYy") + "' and " + _
           "OmChr= '" + ADC("wCpPrdOdChr") + "' and OmNo= " + CStr(ADC("wCpPrdOdNo")) + " and OmPrtKey='" + ctCurrPrtn + "' ") _
           Then DispMsg "Invalid Production Order Voucher Number", etError: Exit Sub
     End If

    If ADC("wCpExpOdNo") <> 0 Then
      '****** Sachin 3.02 28-12-07 - PrtKey added
      If Not moCn.RecSeek("select OmNo from OrdMst where OmCoCd= '" + gs_CoCd + "' and " + _
           "OmTc= '" + ADC("wCpExpOdTc") + "' and OmYy= '" + ADC("wCpExpOdYy") + "' and " + _
           "OmChr= '" + ADC("wCpExpOdChr") + "' and OmNo= " + CStr(ADC("wCpExpOdNo")) + " and OmPrtKey='" + ctCurrPrtn + "' ") _
           Then DispMsg "Invalid Export Order Voucher Number", etError: Exit Sub
     End If
    
    If ADC("wCpPrdOdSr") <> 0 Then
      '****** Sachin 3.02 28-12-07 - PrtKey added
      If Not moCn.RecSeek("select OdSr from OrdDsg where OdCoCd= '" + gs_CoCd + "' and " + _
           "OdTc= '" + ADC("wCpPrdOdTc") + "' and OdYy= '" + ADC("wCpPrdOdYy") + "' and " + _
           "OdChr= '" + ADC("wCpPrdOdChr") + "' and OdNo= " + CStr(ADC("wCpPrdOdNo")) + " " + _
           "and OdSr= " + CStr(ADC("wCpPrdOdSr")) + " and OdPrtKey='" + ctCurrPrtn + "' ") _
           Then DispMsg "Invalid Production Order Voucher Serial", etError: Exit Sub
    End If
    
    If ADC("wCpExpOdSr") <> 0 Then
      '****** Sachin 3.02 28-12-07 - PrtKey added
      If Not moCn.RecSeek("select OdSr from OrdDsg where OdCoCd= '" + gs_CoCd + "' and " + _
           "OdTc= '" + ADC("wCpExpOdTc") + "' and OdYy= '" + ADC("wCpExpOdYy") + "' and " + _
           "OdChr= '" + ADC("wCpExpOdChr") + "' and OdNo= " + CStr(ADC("wCpExpOdNo")) + " " + _
           "and OdSr= " + CStr(ADC("wCpExpOdSr")) + " and OdPrtKey='" + ctCurrPrtn + "' ") _
           Then DispMsg "Invalid Export Order Voucher Serial", etError: Exit Sub
    End If
    
    If ADC("wCpInNo") <> 0 Then
      '****** Sachin 3.02 28-12-07 - PrtKey added
      If Not moCn.RecSeek("select InNo from InvHd where InCoCd= '" + gs_CoCd + "' and " + _
           "InTc= '" + ADC("wCpInTc") + "' and InYy= '" + ADC("wCpInYy") + "' and " + _
           "InChr= '" + ADC("wCpInChr") + "' and InNo= " + CStr(ADC("wCpInNo")) + " and InPrtKey='" + ctCurrPrtn + "' ") _
           Then DispMsg "Invalid Invoice Voucher Number", etError: Exit Sub
    End If
    
    If ADC("wCpIdSr") <> 0 Then
      '****** Sachin 3.02 28-12-07 - PrtKey added
      If Not moCn.RecSeek("select IdSr from InvDsg where IdCoCd= '" + gs_CoCd + "' and " + _
           "IdTc= '" + ADC("wCpInTc") + "' and IdYy= '" + ADC("wCpInYy") + "' and " + _
           "IdChr= '" + ADC("wCpInChr") + "' and IdNo= " + CStr(ADC("wCpInNo")) + " " + _
           "and IdSr= " + CStr(ADC("wCpIdSr")) + " and IdPrtKey='" + ctCurrPrtn + "' ") _
           Then DispMsg "Invalid Invoice Voucher Serial", etError: Exit Sub
    End If
    
    If ADC("wCpPlNo") <> 0 Then
      '****** Sachin 3.02 28-12-07 - PrtKey added
      If Not moCn.RecSeek("Select FdPlNo From Fgd Where FdCoCd= '" + gs_CoCd + "' and " + _
         "FdPlYy= '" + ADC("wCpPlYy") + "' And " + _
         "FdPlChr= '" + ADC("wCpPlChr") + "' and FdPlNo= " + CStr(ADC("wCpPlNo")) + " And FdPrtKey='" + ctCurrPrtn + "' ") Then _
         DispMsg "Invalid Packing List Number", etError: Exit Sub
    End If
    '6.1_11 checking validate shipment note number
    If ADC("wCpSnNo") <> 0 Then
      If Not moCn.RecSeek("select SnNo from ShpNote where SnCoCd= '" + gs_CoCd + "' and " + _
             "SnTc= '" + ADC("wCpSnTc") + "' and SnYy= '" + ADC("wCpSnYy") + "' and " + _
             "SnChr= '" + ADC("wCpSnChr") + "' and SnNo= " + CStr(ADC("wCpSnNo"))) _
             Then DispMsg "Invalid Shipment Note Number", etError: Exit Sub
    End If
  End If
  
  Call CpyFgBags
  Call CmdCP_Click
End Sub
Private Sub CmdSc_Click()
  '*** Display or hide the Scope frame depending on the toggle satate
      'Before closing the visible frame validate all the fields of the frame whose values have been entered
      
  Call DispFra(Scop)
  
  '*** Close the Scope Frame
  If FraSC.Visible = True Then
    If ADC("wScPrdOdNo") <> 0 Then
      '****** Sachin 3.02 28-12-07 - PrtKey added
      If Not moCn.RecSeek("select OmNo from OrdMst where OmCoCd= '" + gs_CoCd + "' and " + _
           "OmTc= '" + ADC("wScPrdOdTc") + "' and OmYy= '" + ADC("wScPrdOdYy") + "' and " + _
           "OmChr= '" + ADC("wScPrdOdChr") + "' and OmNo= " + CStr(ADC("wScPrdOdNo")) + IIF(ADC("wIfPrtKey") = ctCurrPrtn, " and OmPrtKey='" + ctCurrPrtn + "' ", "")) _
           Then DispMsg "Invalid Production Order Voucher Number", etError: Exit Sub
     End If

    If ADC("wScExpOdNo") <> 0 Then
      '****** Sachin 3.02 28-12-07 - PrtKey added
      If Not moCn.RecSeek("select OmNo from OrdMst where OmCoCd= '" + gs_CoCd + "' and " + _
           "OmTc= '" + ADC("wScExpOdTc") + "' and OmYy= '" + ADC("wScExpOdYy") + "' and " + _
           "OmChr= '" + ADC("wScExpOdChr") + "' and OmNo= " + CStr(ADC("wScExpOdNo")) + IIF(ADC("wIfPrtKey") = ctCurrPrtn, " and OmPrtKey='" + ctCurrPrtn + "' ", "")) _
           Then DispMsg "Invalid Export Order Voucher Number", etError: Exit Sub
     End If

    If ADC("wScPrdOdSr") <> 0 Then
      '****** Sachin 3.02 28-12-07 - PrtKey added
      If Not moCn.RecSeek("select OdSr from OrdDsg where OdCoCd= '" + gs_CoCd + "' and " + _
           "OdTc= '" + ADC("wScPrdOdTc") + "' and OdYy= '" + ADC("wScPrdOdYy") + "' and " + _
           "OdChr= '" + ADC("wScPrdOdChr") + "' and OdNo= " + CStr(ADC("wScPrdOdNo")) + "  " + _
           "and OdSr= " + CStr(ADC("wScPrdOdSr")) + IIF(ADC("wIfPrtKey") = ctCurrPrtn, " and OdPrtKey='" + ctCurrPrtn + "' ", "")) _
           Then DispMsg "Invalid Production Order Voucher Serial", etError: Exit Sub
    End If

    If ADC("wScExpOdSr") <> 0 Then
      '****** Sachin 3.02 28-12-07 - PrtKey added
      If Not moCn.RecSeek("select OdSr from OrdDsg where OdCoCd= '" + gs_CoCd + "' and " + _
           "OdTc= '" + ADC("wScExpOdTc") + "' and OdYy= '" + ADC("wScExpOdYy") + "' and " + _
           "OdChr= '" + ADC("wScExpOdChr") + "' and OdNo= " + CStr(ADC("wScExpOdNo")) + " " + _
           "and OdSr= " + CStr(ADC("wScExpOdSr")) + IIF(ADC("wIfPrtKey") = ctCurrPrtn, " and OdPrtKey='" + ctCurrPrtn + "' ", "")) _
           Then DispMsg "Invalid Export Order Voucher Serial", etError: Exit Sub
    End If

    If ADC("wScIdSr") <> 0 Then
      '****** Sachin 3.02 28-12-07 - PrtKey added
      If Not moCn.RecSeek("select IdSr from InvDsg where IdCoCd= '" + gs_CoCd + "' and " + _
           "IdTc= '" + ADC("wIfTc") + "' and IdYy= '" + ADC("wIfYy") + "' and " + _
           "IdChr= '" + ADC("wIfChr") + "' and IdNo= " + CStr(ADC("wIfNo")) + " " + _
           "and IdSr= " + CStr(ADC("wScIdSr")) + IIF(ADC("wIfPrtKey") = ctCurrPrtn, " and IdPrtKey='" + ctCurrPrtn + "' ", "")) _
           Then DispMsg "Invalid Invoice Voucher Serial", etError: Exit Sub
    End If

    If ADC("wScPlNo") <> 0 Then
      '****** Sachin 3.02 28-12-07 - PrtKey added
      If Not moCn.RecSeek("Select FdPlNo From Fgd Where FdCoCd= '" + gs_CoCd + "' and " + _
         "FdPlYy= '" + ADC("wScPlYy") + "' And " + _
         "FdPlChr= '" + ADC("wScPlChr") + "' and FdPlNo= " + CStr(ADC("wScPlNo")) + IIF(ADC("wIfPrtKey") = ctCurrPrtn, " and FdPRtKey='" + ctCurrPrtn + "' ", "")) Then _
         DispMsg "Invalid Packing List Number", etError: Exit Sub
    End If
  End If
End Sub
Private Sub CmdSumm_Click()
  '*** Display or hide the Summmary frame depending on the toggle satate
      'If Summary option has been selected call the Summary routine
      
  Call DispFra(Summ)
  If FraSumm.Visible = True Then
    If ADC("wSummBagGrp") = "" Then
      ADC("wSummBagGrp") = "PRDORDSR"
    Else
      Call Summary(ADC("wSummBagGrp"))
    End If
  
    '***** Sachin 3.11.0 - RM Summary in Invoice All Bags
    If ADC("wBrsGrp") = "" Then ADC("wBrsGrp") = "RMCTG": ADC("wVrsGrp") = "RMCTG"    '***** Sachin 3.11.0
    Call RMSummary(ADC("wBrsGrp"), "N")
    Call RMSummary(ADC("wVrsGrp"), "Y")
  End If
End Sub
Private Sub CpyFgBags(Optional ByVal pb_CpyRfId As Boolean = False)
  '*** Copy is allowed even if there are records existing in InvFgd
  '*** Check if rec exists in InvDsg (if true then copy not possible)
  '*** atleast one of the values in the copy frame should be filled up (else do not allow copy as all records in fgd will be copied which is not desirable)
  '*** If entered (in selection option), check if the PrdOrd, ExpOrd, Inv, PckLst are valid
  '*** Select only those recs from Fgd which are not already present in InvFgd for the given
      'scope and where FdQty- FdDespQty> 0
      'IfGrsWt is FdGrsWt*(FdQty-FdDespQty)/FdQty
  
  Dim wSelCnd As String, wIfSrNo As Single
  Dim ws_InCurCd As String    ' ***** (Jen 2.13)
  Dim ws_RfIdJoin As String   ' ***** Manali 3.10.0 - 19/04/12 - Cpy Bags from RfId
  Dim ws_ShpNoteJoin As String   '6.1
  ' ***** Manali 3.10.0 - 19/04/12 - Cpy Bags from RfId
  If Not pb_CpyRfId Then
    Call CmpString("C")     ' to set the Compare String Property of the ADC controls when the Copy Option Is clicked
    
    '*** Checking if copy is possible
    '* ***** Manali 3.5.0 - 11/11/08 - DmCtg added
    '6.1_11 ShpNote number added
    If ADC("wCpPrdCmCd") = "" And ADC("wCpExpCmCd") = "" And ADC("wCpCtCd") = "" And _
       ADC("wCpPrdOdNo") = 0 And ADC("wCpExpOdNo") = 0 And ADC("wCpDmCd") = "" And ADC("wCpDmCtg") = "" And _
       ADC("wCpPlNo") = 0 And ADC("wCpInNo") = 0 And ADC("wCpInExpNo") = "" And ADC("wCpSnNo") = 0 And _
       ADC("wCpFgSubLoc") = "" And (ADC("wCpFgDtFr") = "" Or ADC("wCpFgDtTo") = "") Then _
       DispMsg "Enter Atleast One Selection Parameter For Copying", etError: Exit Sub
    
    wSelCnd = ADC.RepCond
    wSelCnd = IIF(wSelCnd <> "", " and " + wSelCnd, "")
    ws_RfIdJoin = ""
    '6.1_11 if shipment note filters are given then shipment note tables will join
    If ADC("wCpSnTc") <> "" And ADC("wCpSnYy") <> "" And ADC("wCpSnChr") <> "" And ADC("wCpSnNo") > 0 Then
      ws_ShpNoteJoin = " join ShpNoteOrdDsg on SodOdCoCd =FdCoCd and SodOdTc =FdExpOdTc and SodOdYy =FdExpOdYy and SodOdYy =FdExpOdYy " + _
                      " and SodOdChr =FdExpOdChr and SodOdNo =FdExpOdNo and SodSrNo =FdExpOdSr " + _
                      " join ShpNoteOrdMst on SomCoCd =SodCoCd and SomTc =SodTc and SomYy =SodYy and SomChr =SodChr and SomNo =SodNo and SomSr =SodSr " + _
                      " join ShpNote on SnCoCd =SomCoCd and SnTc =SomTc and SnYy =SomYy and SomChr =SnChr and SnNo =SomNo "
    End If
    
  Else
    ' ***** Manali 3.10.0 - 19/04/12 - Cpy Bags from RfId
    'vk.4 RfId changed as Bag Alias No.
    If ADC("wRfIdSel") = "" Then DispMsg "Enter Bag Alias No. List/ Bag No List", etError: Exit Sub
    Dim wRfIdStr As String, wRfIdNoArr() As String, wRfIdNo As String, i As Integer
    'vk.4 calling function to replace enter with ,
    wRfIdStr = PrepareRfIdString(ADC("wRfIdSel"))
    ws_ShpNoteJoin = ""     '6.1_11 no join for rfid
    If ADC("wRfIdSel") <> "" Then
      wRfIdStr = ParseBagNoBarCd(wRfIdStr)
      wRfIdNoArr = Split(wRfIdStr, ",")
  
      If Left(wRfIdStr, 1) = "," Or Right(wRfIdStr, 1) = "," Then
      'vk.4 Tag Input changed as Bag Alias No. / Bag No.
          DispMsg "Invalid Bag Alias No. / Bag No.", etError: Exit Sub
      End If
      wRfIdStr = ""
      For i = 0 To UBound(wRfIdNoArr)
      '****Bhavna added bag no seln
        wRfIdNo = wRfIdNoArr(i)
        If ADC("WCPYTHRUBAGNO") = "N" Then
            'vk.4 Tag changed as Bag Alias No.
            If Not IsNumeric(wRfIdNo) Then DispMsg "Bag Alias No. Not Numeric", etError: Exit Sub
            If Not moCn.RecSeek("Select 'x' from RfIdMst Where RimNo=" + wRfIdNo + " ") Then
                'vk.4 Tag changed as Bag Alias No.
                DispMsg wRfIdNo + " : Bag Alias No. Not Defined", etError: Exit Sub
            End If
         Else
            If Not moCn.RecSeek("Select 'x' from Bag Where Byy+'/'+BChr+'/'+ltrim(RTrim(str(BNo))) = '" + wRfIdNo + "'") Then
                DispMsg wRfIdNo + " : Bag Not Defined", etError: Exit Sub
            End If
         End If
        wRfIdStr = wRfIdStr + IIF(wRfIdStr = "", "", ",") + wRfIdNo
      Next
      
      If ADC("WCPYTHRUBAGNO") = "N" Then
        ws_RfIdJoin = " Join RfIdMst On RimBCoCd=BCoCd and RimBYy=BYy and RimBChr=BChr and RimBNo=BNo "
        wSelCnd = " and RimNo In (" + wRfIdStr + ") "
      Else
          ws_RfIdJoin = " "
          wSelCnd = " and  charindex(','+BYy + '/' + BChr + '/' + LTrim(RTrim(str(BNo)))+',' , '," + wRfIdStr + ",') <> 0 "
      End If
    End If
    ' ***** Manali 3.10.0 - 19/04/12 - Cpy Bags from RfId
  End If
  
  '*** Beginning the copy process
  GrdInvFgd.StartCopy
  MWLib.BeginProcess Me, "Copying Finished Goods Details ..."
  Dim wRsAllBags As MwfLib.MDORowSet, wRow As Integer

  '*** (Jen 2.13)
  '****** Sachin 3.02 28-12-07 - PrtKey added
  ws_InCurCd = moCn.GetFldVal("Select CmCurCd from InvHd join CustMst on CmCtg= '" + ms_CmCtg + "' and CmCd= InCmCd " + _
                              "where InCoCd= '" + gs_CoCd + "' and InTc= '" + ADC("wIfTc") + "' " + _
                              "and InYy= '" + ADC("wIfYy") + "' and InChr= '" + ADC("wIfChr") + "' " + _
                              "and InNo= " + CStr(ADC("wIfNo")) + " and InPrtKey='" + ctCurrPrtn + "' ")
  '*** (Jen 2.13)

  If UCase(ms_TcTyp) = UCase(ctInTcTyp) Then
  
    '*** (Bef 2.13) Set wRsAllBags = moCn.OpenRes("Select * from Fgd join Fg on FgCoCd= FdCoCd and FgTc= FdTc "
  
    '****** (Jen 2.13) added OrdMst join for OmCmCurCd check ******
    '****** Sachin 3.02 - prtkey added below
     ' ***** Manali 3.5.0 - 11/11/08 - DmCtg added in scope for copy [Join with DsgMst]
     ' ***** Manali 3.10.0 - 19/04/12 - Cpy Bags from RfId
     '6.1_11 ws_ShpNoteJoin added
        Set wRsAllBags = moCn.OpenRes("Select FdIdNo, FdBIdNo, FdTc, FdYy, FdChr, FdNo, FdSr, FdBYy, FdBChr, FdBNo, FdDmCd, FdDmSz, FdSfx, " + _
                     "FdQty, FdDespQty, FdCTBDespQty, FdGrWt, OdCoCd, OdTc, OdYy, OdChr, OdNo, OdSr, OdHld, OdOrdQty, " + _
                     "OdExpQty from Fgd join Fg on FgIdNo=FdFgIdNo and FgPrtKey=FdPrtKey and FgCoCd= FdCoCd and FgTc= FdTc " + _
                     " and FgYy= FdYy and FgChr= FdChr and FgNo= FdNo " + _
                     "join Bag on BIdNo=FdBIdNo and BCoCd= FdCoCd and BYy= FdBYy and BChr= FdBChr and BNo= FdBNo and BPrtKey=FdPrtKey " + _
                     "join OrdDsg on OdCoCd= FdCoCd and OdTc= FdExpOdTc and OdYy= FdExpOdYy " + _
                     "and OdChr= FdExpOdChr and OdNo= FdExpOdNo and OdSr= FdExpOdSr and OdPrtKey=FdPrtKey " + _
                     "join OrdMst on OmIdNo=OdOmIdNo and OmCoCd= OdCoCd and OmTc= OdTc and OmYy= OdYy and OmChr= OdChr and OmNo=OdNo and OmPrtKey=OdPrtKey " + _
                     "join Param on PTyp= 'TC' and PMCd= FgTc " + _
                     "Join DsgMst on DmTcTyp='DM' and DmCd=FdDmCd and DmPrtKey=FdPrtKey and DmIdNo=FdDmIdNo " + ws_RfIdJoin + ws_ShpNoteJoin + _
                     "where OdDmTcTyp='DM' and FdCoCd= '" + gs_CoCd + "' and FdPrtKey='" + ctCurrPrtn + "' and " + _
                     "Round(FdQty, 1)> Round(FdDespQty, 1) and PValue in ('FB', 'FR') " + _
                     wSelCnd + " and OmCmCurCd= '" + ws_InCurCd + "' Order By FdBYy, FdBChr, FdBNo")
      
   ElseIf UCase(ms_TcTyp) = UCase(ctCTBTcTyp) Then
    '****** (Jen 2.13) added OrdMst join for OmCmCurCd check ******
    '****** Sachin 3.02 - Id fields In Joins
    Set wRsAllBags = moCn.OpenRes("Select FdIdNo, FdBIdNo, FdTc, FdYy, FdChr, FdNo, FdSr, FdBYy, FdBChr, FdBNo, FdDmCd, FdDmSz, FdSfx, " + _
                     "FdQty, FdDespQty, FdCTBDespQty, FdGrWt, OdCoCd, OdTc, OdYy, OdChr, OdNo, OdSr, OdHld, OdOrdQty, " + _
                     "OdExpQty, OmCmCurCd from Fgd B join Fg on FgIdNo=FdFgIdNo and FgPrtKey=FdPrtKey and FgCoCd= FdCoCd " + _
                     "and FgTc= FdTc and FgYy= FdYy and FgChr= FdChr and FgNo= FdNo " + _
                     "join OrdDsg on OdCoCd= FdCoCd and OdTc= FdPrdOdTc and OdYy= FdPrdOdYy " + _
                     "and OdChr= FdPrdOdChr and OdNo= FdPrdOdNo and OdSr= FdPrdOdSr and OdPrtKey=FdPrtKey " + _
                     "join OrdMst on OmIdNo=OdOmIdNo and OmCoCd= OdCoCd and OmTc= OdTc and OmYy= OdYy and OmChr= OdChr and OmNo= OdNo " + _
                     "join Bag on BIdNo=FdBIdNo and BYy= FdBYy and BChr= FdBChr and BNo= FdBNo and BPrtKey=FdPrtKey " + _
                     "join Param on PTyp= 'TC' and PMCd= FgTc " + _
                     "where FdCoCd= '" + gs_CoCd + "' and FdPrtKey='" + ctCurrPrtn + "' and PValue in ('FB') " + _
                     "and Round(FdQty, 1)> Round(FdCTBDespQty, 1) and  " + _
                     "not exists(Select 'x' from Fgd M where M.FdCoCd= B.FdCoCd " + _
                     "and M.FdRefYy= B.FdYy and M.FdRefKey= B.FdKey and M.FdPrtKey=B.FdPrtKey " + _
                     "and (Select PValue from Param where PTyp= 'TC' and PMCd= M.FdTc)= 'FWB') " + _
                     wSelCnd + " and OmCmCurCd= '" + ws_InCurCd + "' Order By FdBYy, FdBChr, FdBNo")
  End If
 Dim wDmHldCnt As Integer
  
  '****** Sachin 3.02 28-12-07 - PrtKey added
  wIfSrNo = moCn.GetFldVal("Select max(IfSrNo) from InvFgd where IfCoCd= '" + gs_CoCd + "' " + _
            "and IfTc= '" + ADC("wIfTc") + "' and  IfYy= '" + ADC("wIfYy") + "' and " + _
            "IfChr= '" + ADC("wIfChr") + "' and IfNo= " + CStr(ADC("wIfNo")) + " and IfPrtKey='" + ctCurrPrtn + "' ")
   wDmHldCnt = 0
  With wRsAllBags
    Do While Not (.EOF Or .BOF)
      MWLib.RunningProcess "Copying Details For Voucher " + !FdTc + "/" + !FdYy + "/" + !FdChr + "/" + CStr(!FdNo) + "/" + CStr(!FdSr)
      
      'If !OdOrdQty - !OdExpQty - (!FdQty - !FdDespQty) - _
         moCn.GetFldVal("Select sum(IfQty) from InvFgd join Fgd on FdCoCd= IfCoCd and " + _
         "FdTc= IfFdTc and FdYy= IfFdYy and FdChr= IfFdChr and FdNo= IfFdNo and FdSr= IfFdSr where IfSr= 0 and " + _
         "FdCoCd = '" + !OdCoCd + "' and FdExpOdTc = '" + !OdTc + "' and FdExpOdYy= '" + !OdYy + "' and " + _
         "FdExpOdChr= '" + !OdChr + "' and FdExpOdNo= " + CStr(!OdNo) + " and " + _
         "FdExpOdSr= " + CStr(!OdSr) + " ") >= 0 Then
          '******************Geeta*******************08/04/04
      '****** Sachin 3.02 - PrtKey added below
      ' ***** Manali 3.9.0- 07/02/12 - TcTyp=IntTcTyp added
      If (UCase(ms_TcTyp) = UCase(ctInTcTyp) And (!OdOrdQty - !OdExpQty - (!Fdqty - !FdDespQty) - _
         moCn.GetFldVal("Select sum(IfQty) from InvFgd join Fgd on FdIdNo=IfFdIdNo and FdCoCd= IfCoCd and " + _
         "FdTc= IfFdTc and FdYy= IfFdYy and FdChr= IfFdChr and FdNo= IfFdNo and FdSr= IfFdSr and FdPrtKey=IfPrtKey where IfSr= 0 and " + _
         "(Select PValue from Param where PTyp= 'TC' and PMCd= IfTc)= '" + ctInTcTyp + "' and " + _
         "FdCoCd = '" + !OdCoCd + "' and FdExpOdTc = '" + !OdTc + "' and FdExpOdYy= '" + !OdYy + "' and " + _
         "FdExpOdChr= '" + !OdChr + "' and FdExpOdNo= " + CStr(!OdNo) + " and " + _
         "FdExpOdSr= " + CStr(!OdSr) + " and FdPrtKey='" + ctCurrPrtn + "' ") >= 0) And (!OdHld = "N")) Or _
         (UCase(ms_TcTyp) = UCase(ctCTBTcTyp) And (!OdOrdQty - (!Fdqty - !FdCTBDespQty) - _
         moCn.GetFldVal("Select sum(IfQty) from InvFgd join Fgd on FdIdNo=IfFdIdNo and FdCoCd= IfCoCd and " + _
         "FdTc= IfFdTc and FdYy= IfFdYy and FdChr= IfFdChr and FdNo= IfFdNo and FdSr= IfFdSr and FdPrtKey=IfPrtKey where " + _
         "(Select PValue from Param where PTyp= 'TC' and PMCd= IfTc)= '" + ctCTBTcTyp + "' and " + _
         "FdCoCd = '" + !OdCoCd + "' and FdPrdOdTc = '" + !OdTc + "' " + _
         "and FdPrdOdYy= '" + !OdYy + "' and FdPrdOdChr= '" + !OdChr + "' and " + _
         "FdPrdOdNo= " + CStr(!OdNo) + " and FdPrdOdSr= " + CStr(!OdSr) + " and FdPrtKey='" + ctCurrPrtn + "' ") >= 0)) Then
        
        GrdInvFgd.AddItem
        wRow = GrdInvFgd.Rows - 1
        GrdInvFgd.Value(wRow, "IfSr") = 0
        '##### GrdInvFgd.Value(wRow, "IfSrNo") = moCn.AutoGen("InvFgd", "IfSrNo", adc("wIfTc") + adc("wIfYy") + adc("wIfChr") + CStr(adc("wIfNo")))
        wIfSrNo = wIfSrNo + 1
        GrdInvFgd.Value(wRow, "IfSrNo") = wIfSrNo
        
        GrdInvFgd.Value(wRow, "IfFdTc") = !FdTc
        GrdInvFgd.Value(wRow, "IfFdYy") = !FdYy
        GrdInvFgd.Value(wRow, "IfFdChr") = !FdChr
        GrdInvFgd.Value(wRow, "IfFdNo") = !FdNo
        GrdInvFgd.Value(wRow, "IfFdSr") = !FdSr
        
        GrdInvFgd.Value(wRow, "IfFdBYy") = !FdBYy
        GrdInvFgd.Value(wRow, "IfFdBChr") = !FdBChr
        GrdInvFgd.Value(wRow, "IfFdBNo") = !FdBNo
        
        GrdInvFgd.Value(wRow, "wIfDmCd") = !FdDmCd
        GrdInvFgd.Value(wRow, "wIfDmSz") = !FdDmSz
        GrdInvFgd.Value(wRow, "wIfSfx") = !FdSfx
        
        If UCase(ms_TcTyp) = UCase(ctInTcTyp) Then
          GrdInvFgd.Value(wRow, "IfQty") = !Fdqty - !FdDespQty
          GrdInvFgd.Value(wRow, "IfGrsWt") = !FdGrWt * MWLib.Div((!Fdqty - !FdDespQty), !Fdqty)
        ElseIf UCase(ms_TcTyp) = UCase(ctCTBTcTyp) Then
          GrdInvFgd.Value(wRow, "IfQty") = !Fdqty - !FdCTBDespQty
          GrdInvFgd.Value(wRow, "IfGrsWt") = !FdGrWt * MWLib.Div((!Fdqty - !FdCTBDespQty), !Fdqty)
        End If
        
        '****** Sachin 3.02 - Setting Values To Id fields
        '****** Sachin 3.02 - 14-01-08 - Shifted to InitKey
        '        GrdInvFgd.Value(wRow, "IfFdIdNo") = !FdIdNo
        '        GrdInvFgd.Value(wRow, "IfBIdNo") = !FdBIdNo
        '****** Sachin 3.02 - Setting Values To Id fields
     
        GrdInvFgd.Value(wRow, "IfPrtKey") = ctCurrPrtn      '****** Sachin 3.02 28-12-07
     
        GrdInvFgd.SaveRec (wRow)
        
      ElseIf !OdHld = "Y" Then
        wDmHldCnt = wDmHldCnt + 1
      End If
         
     .MoveNext
    Loop
  End With
  MWLib.EndProcess Me
  GrdInvFgd.EndCopy
  If wDmHldCnt > 0 And UCase(ms_TcTyp) = UCase(ctInTcTyp) Then _
    DispMsg "Cannot Copy " + CStr(wDmHldCnt) + " Design(s) on Hold", etWarning
  Set wRsAllBags = Nothing
  '*** End of the copy process
End Sub
Private Sub CpyFgBags1()
'  '(Bef Rel 1.11)
'  '*** Copy is allowed even if there are records existing in InvFgd
'  '*** Check if rec exists in InvDsg (if true then copy not possible)
'  '*** atleast one of the values in the copy frame should be filled up (else do not allow copy as all records in fgd will be copied which is not desirable)
'  '*** If entered (in selection option), check if the PrdOrd, ExpOrd, Inv, PckLst are valid
'  '*** Select only those recs from Fgd which are not already present in InvFgd for the given
'      'scope and where FdQty- FdDespQty> 0
'      'IfGrsWt is FdGrsWt*(FdQty-FdDespQty)/FdQty
'
'  Dim wSelCnd As String, wIfSrNo As Single
'  Call CmpString("C")     ' to set the Compare String Property of the ADC controls when the Copy Option Is clicked
'
'  '*** Checking if copy is possible
'  If ADC("wCpPrdCmCd") = "" And ADC("wCpExpCmCd") = "" And ADC("wCpPrdOdNo") = 0 And _
'     ADC("wCpExpOdNo") = 0 And ADC("wCpDmCd") = "" And ADC("wCpPlNo") = 0 And _
'     ADC("wCpInNo") = 0 And ADC("wCpInExpNo") = "" And _
'     (ADC("wCpFgDtFr") = "" Or ADC("wCpFgDtTo") = "") Then _
'     DispMsg "Enter Atleast One Selection Parameter For Copying", etError: Exit Sub
'
'  wSelCnd = ADC.RepCond
'  wSelCnd = IIf(wSelCnd <> "", " and " + wSelCnd, "")
'
'  '*** Beginning the copy process
'  GrdInvFgd.StartCopy
'  MWLib.BeginProcess Me, "Copying Finished Goods Details ..."
'  Dim wRsAllBags As MwfLib.MDORowSet, wRow As Integer
'  Set wRsAllBags = moCn.OpenRes("Select * from Fgd join Fg on FgTc= FdTc and FgYy= FdYy " + _
'                   "and FgChr= FdChr and FgNo= FdNo " + _
'                   "join Param on PTyp= 'TC' and PMCd= FgTc " + _
'                   "where Round(FdQty, 1)> Round(FdDespQty, 1) and PValue in ('FB', 'FR') " + _
'                   wSelCnd + _
'                   " Order By FdBYy, FdBChr, FdBNo")
'
'  wIfSrNo = moCn.GetFldVal("Select max(IfSrNo) from InvFgd " + _
'            "where IfTc= '" + ADC("wIfTc") + "' and  IfYy= '" + ADC("wIfYy") + "' and " + _
'            "IfChr= '" + ADC("wIfChr") + "' and IfNo= " + CStr(ADC("wIfNo")))
'  With wRsAllBags
'    Do While Not (.EOF Or .BOF)
'      MWLib.RunningProcess "Copying Details For Voucher " + !FdTc + "/" + !FdYy + "/" + !FdChr + "/" + CStr(!FdNo) + "/" + CStr(!FdSr)
'      GrdInvFgd.AddItem
'      wRow = GrdInvFgd.Rows - 1
'      GrdInvFgd.Value(wRow, "IfSr") = 0
'      '##### GrdInvFgd.Value(wRow, "IfSrNo") = moCn.AutoGen("InvFgd", "IfSrNo", adc("wIfTc") + adc("wIfYy") + adc("wIfChr") + CStr(adc("wIfNo")))
'      wIfSrNo = wIfSrNo + 1
'      GrdInvFgd.Value(wRow, "IfSrNo") = wIfSrNo
'
'      GrdInvFgd.Value(wRow, "IfFdTc") = !FdTc
'      GrdInvFgd.Value(wRow, "IfFdYy") = !FdYy
'      GrdInvFgd.Value(wRow, "IfFdChr") = !FdChr
'      GrdInvFgd.Value(wRow, "IfFdNo") = !FdNo
'      GrdInvFgd.Value(wRow, "IfFdSr") = !FdSr
'
'      GrdInvFgd.Value(wRow, "IfFdBYy") = !FdBYy
'      GrdInvFgd.Value(wRow, "IfFdBChr") = !FdBChr
'      GrdInvFgd.Value(wRow, "IfFdBNo") = !FdBNo
'
'      GrdInvFgd.Value(wRow, "wIfDmCd") = !FdDmCd
'      GrdInvFgd.Value(wRow, "wIfDmSz") = !FdDmSz
'      GrdInvFgd.Value(wRow, "wIfSfx") = !FdSfx
'
'      GrdInvFgd.Value(wRow, "IfQty") = !FdQty - !FdDespQty
'      GrdInvFgd.Value(wRow, "IfGrsWt") = !FdGrWt * MWLib.Div((!FdQty - !FdDespQty), !FdQty)
'      GrdInvFgd.SaveRec (wRow)
'      .MoveNext
'    Loop
'  End With
'  MWLib.EndProcess Me
'  GrdInvFgd.EndCopy
'  Set wRsAllBags = Nothing
'  '*** End of the copy process
End Sub
Private Sub Summary(ByVal mGrpByCd As String)
  '*** Summary Group Options :
      'Group By FdPrdCmCd
      'Group By FdExpCmCd
      'Group By FdPrdOdTc,FdPrdOdYy,FdPrdOdChr,FdPrdOdNo
      'Group By FdPrdOdTc,FdPrdOdYy,FdPrdOdChr,FdPrdOdNo,FdPrdOdSr
      'Group By FdExpOdTc,FdExpOdYy,FdExpOdChr,FdExpOdNo
      'Group By FdExpOdTc,FdExpOdYy,FdExpOdChr,FdExpOdNo,FdExpOdSr
      'Group By FdDmCd
      'Group By FdPlChr,FdPlYy,FdPlNo
      'Group By IfTc,IfYy,IfChr,IfNo,IfSr
      'Group By InExpNo
      'Group By FgDt
  '*** Whenever a particular group option is selected, some of the summary grid fields
      'become irrevelent. So these fields should be made invisible.
  
  '### Should a key be created on IfFdTc, IfFdYy, IfFdChr, IfFdNo, IfFdSr    (Ans = No)

  Dim wGrpByCd  As String, wGrpBy As String, wCommStr As String
  Dim wTotRmWt As Single, wTotBagQty As Single, wTotBags As Single, wStr As String, wRow As Single
  
  wGrpByCd = mGrpByCd
  If wGrpByCd = "PRDCMCD" Then wGrpBy = "FdPrdCmCd"
  If wGrpByCd = "EXPCMCD" Then wGrpBy = "FdExpCmCd"
  If wGrpByCd = "PRDORDNO" Then wGrpBy = "FdPrdOdTc,FdPrdOdYy,FdPrdOdChr,FdPrdOdNo"
  If wGrpByCd = "PRDORDSR" Then wGrpBy = "FdPrdOdTc,FdPrdOdYy,FdPrdOdChr,FdPrdOdNo,FdPrdOdSr"
  If wGrpByCd = "EXPORDNO" Then wGrpBy = "FdPrdOdTc,FdPrdOdYy,FdPrdOdChr,FdPrdOdNo"
  If wGrpByCd = "EXPORDSR" Then wGrpBy = "FdPrdOdTc,FdPrdOdYy,FdPrdOdChr,FdPrdOdNo,FdPrdOdSr"
  If wGrpByCd = "DSGCD" Then wGrpBy = "FdDmCd"
  If wGrpByCd = "PCKLST" Then wGrpBy = "FdPlChr,FdPlYy,FdPlNo"
  If wGrpByCd = "INVSR" Then wGrpBy = "IfTc,IfYy,IfChr,IfNo,IfSr"
  If wGrpByCd = "INVEXPNO" Then wGrpBy = "InExpNo"
  If wGrpByCd = "FGDATE" Then wGrpBy = "FgDt"
  
  wTotRmWt = 0: wTotBagQty = 0: wTotBags = 0: wStr = ""
    
  '****** Sachin 3.02 - Corresponding Partition
  wCommStr = " sum(case when IfFdBNo<>0 then 1 else 0 end) as qBags, sum(IfQty) as qBagQty, " + _
             "sum(IfGrsWt) as qRmWt From InvFgd " + _
             "Join Fgd On FdIdNo=IfFdIdNo and " + IIF(gs_Partition = ctCurrPrtn, " FdPrtKey=IfPrtKey And ", "") + " FdCoCd= IfCoCd and FdTc= IfFdTc and FdYy= IfFdYy and FdChr= IfFdChr " + _
             "and FdNo= IfFdNo and FdSr= IfFdSr join Fg on FgIdNo=FdFgIdNo and FgCoCd=FdCoCd and FgTc= FdTc " + _
             "and FgYy= FdYy and FgChr= FdChr and FgNo= FdNo " + IIF(gs_Partition = ctCurrPrtn, " And FgPrtKey=FdPrtKey ", "") + " join InvHd on InIdNo=IfInIdNo and InCoCd= IfCoCd and " + _
             "InTc= IfTc and InYy= IfYy and InChr= IfChr and InNo= IfNo " + IIF(gs_Partition = ctCurrPrtn, " And InPrtKey=IfPrtKey ", "") + _
             "Where IfCoCd= '" + gs_CoCd + "' and IfTc= '" + ADC("wIfTc") + "' and " + _
             "IfYy= '" + ADC("wIfYy") + "' And IfChr ='" + ADC("wIfChr") + "' And " + _
             "IfNo= " + CStr(ADC("wIfNo")) + " and IfPrtKey=(Case When '" + ADC("wIfPrtKey") + "'='" + ctCurrPrtn + "' Then '" + ctCurrPrtn + "' Else IfPrtKey End) " + " Group By " + wGrpBy
  
  If wGrpBy = "FdPrdCmCd" Then
    wStr = "Select FdPrdCmCd as qPrdCmCd, '*' as qPrdOrd, '*' as qExpCmCd, " + _
           "'*' as qExpOrd, '*' as qDmCd, '*' as qPckLst, 0 as qInv, '*' as qExpNo, " + _
           "'' as qFgDt, " + wCommStr
    GrdSumm.ColProp("wSummPrdCmCd").Style = fgcNormal
    GrdSumm.ColProp("wSummPrdOrd").Style = fgcinvisible
    GrdSumm.ColProp("wSummExpCmCd").Style = fgcinvisible
    GrdSumm.ColProp("wSummExpOrd").Style = fgcinvisible
    GrdSumm.ColProp("wSummDmCd").Style = fgcinvisible
    GrdSumm.ColProp("wSummPl").Style = fgcinvisible
    GrdSumm.ColProp("wSummInv").Style = fgcinvisible
    GrdSumm.ColProp("wSummInvExpNo").Style = fgcinvisible
    GrdSumm.ColProp("wSummFgDt").Style = fgcinvisible
  ElseIf wGrpBy = "FdExpCmCd" Then
    wStr = "Select '*' as qPrdCmCd, '*' as qPrdOrd, FdExpCmCd as qExpCmCd, " + _
           "'*' as qExpOrd, '*' as qDmCd, '*' as qPckLst, '*' as qInv, '*' as qExpNo, " + _
           "'' as qFgDt, " + wCommStr
    GrdSumm.ColProp("wSummPrdCmCd").Style = fgcinvisible
    GrdSumm.ColProp("wSummPrdOrd").Style = fgcinvisible
    GrdSumm.ColProp("wSummExpCmCd").Style = fgcNormal
    GrdSumm.ColProp("wSummExpOrd").Style = fgcinvisible
    GrdSumm.ColProp("wSummDmCd").Style = fgcinvisible
    GrdSumm.ColProp("wSummPl").Style = fgcinvisible
    GrdSumm.ColProp("wSummInv").Style = fgcinvisible
    GrdSumm.ColProp("wSummInvExpNo").Style = fgcinvisible
    GrdSumm.ColProp("wSummFgDt").Style = fgcinvisible
  ElseIf wGrpBy = "FdPrdOdTc,FdPrdOdYy,FdPrdOdChr,FdPrdOdNo" Then
    wStr = "Select max(FdPrdCmCd) as qPrdCmCd, FdPrdOdTc+'/'+FdPrdOdYy+'/'+FdPrdOdChr+'/'+" + _
           "LTrim(Str(FdPrdOdNo)) as qPrdOrd, '*' as qExpCmCd, " + _
           "'*' as qExpOrd, '*' as qDmCd, '*' as qPckLst, '*' as qInv, '*' as qExpNo, " + _
           "'' as qFgDt, " + wCommStr
    GrdSumm.ColProp("wSummPrdCmCd").Style = fgcNormal
    GrdSumm.ColProp("wSummPrdOrd").Style = fgcNormal
    GrdSumm.ColProp("wSummExpCmCd").Style = fgcinvisible
    GrdSumm.ColProp("wSummExpOrd").Style = fgcinvisible
    GrdSumm.ColProp("wSummDmCd").Style = fgcinvisible
    GrdSumm.ColProp("wSummPl").Style = fgcinvisible
    GrdSumm.ColProp("wSummInv").Style = fgcinvisible
    GrdSumm.ColProp("wSummInvExpNo").Style = fgcinvisible
    GrdSumm.ColProp("wSummFgDt").Style = fgcinvisible
  ElseIf wGrpBy = "FdPrdOdTc,FdPrdOdYy,FdPrdOdChr,FdPrdOdNo,FdPrdOdSr" Then
    wStr = "Select max(FdPrdCmCd) as qPrdCmCd, FdPrdOdTc+'/'+FdPrdOdYy+'/'+FdPrdOdChr+'/'+" + _
           "LTrim(Str(FdPrdOdNo))+'/'+LTrim(Str(FdPrdOdSr)) as qPrdOrd, '*' as qExpCmCd, " + _
           "'*' as qExpOrd, max(FdDmCd) as qDmCd, '*' as qPckLst, '*' as qInv, '*' as qExpNo, " + _
           "'' as qFgDt, " + wCommStr
    GrdSumm.ColProp("wSummPrdCmCd").Style = fgcNormal
    GrdSumm.ColProp("wSummPrdOrd").Style = fgcNormal
    GrdSumm.ColProp("wSummExpCmCd").Style = fgcinvisible
    GrdSumm.ColProp("wSummExpOrd").Style = fgcinvisible
    GrdSumm.ColProp("wSummDmCd").Style = fgcNormal
    GrdSumm.ColProp("wSummPl").Style = fgcinvisible
    GrdSumm.ColProp("wSummInv").Style = fgcinvisible
    GrdSumm.ColProp("wSummInvExpNo").Style = fgcinvisible
    GrdSumm.ColProp("wSummFgDt").Style = fgcinvisible
  ElseIf wGrpBy = "FdExpOdTc,FdExpOdYy,FdExpOdChr,FdExpOdNo" Then
    wStr = "Select '*' as qPrdCmCd, '*' as qPrdOrd, max(FdExpCmCd) as qExpCmCd, " + _
           "FdExpOdTc+'/'+FdExpOdYy+'/'+FdExpOdChr+'/'+LTrim(Str(FdExpOdNo)) " + _
           "as qExpOrd, '*' as qDmCd, '*' as qPckLst, " + _
           "'*' as qInv, '*' as qExpNo, '' as qFgDt, " + wCommStr
    GrdSumm.ColProp("wSummPrdCmCd").Style = fgcinvisible
    GrdSumm.ColProp("wSummPrdOrd").Style = fgcinvisible
    GrdSumm.ColProp("wSummExpCmCd").Style = fgcNormal
    GrdSumm.ColProp("wSummExpOrd").Style = fgcNormal
    GrdSumm.ColProp("wSummDmCd").Style = fgcinvisible
    GrdSumm.ColProp("wSummPl").Style = fgcinvisible
    GrdSumm.ColProp("wSummInv").Style = fgcinvisible
    GrdSumm.ColProp("wSummInvExpNo").Style = fgcinvisible
    GrdSumm.ColProp("wSummFgDt").Style = fgcinvisible
  ElseIf wGrpBy = "FdExpOdTc,FdExpOdYy,FdExpOdChr,FdExpOdNo,FdExpOdSr" Then
    wStr = "Select '*' as qPrdCmCd, '*' as qPrdOrd, max(FdExpCmCd) as qExpCmCd, " + _
           "FdExpOdTc+'/'+FdExpOdYy+'/'+FdExpOdChr+'/'+LTrim(Str(FdExpOdNo))+'/'+" + _
           "LTrim(Str(FdExpOdSr)) as qExpOrd, max(FdDmCd) as qDmCd, '*' as qPckLst, " + _
           "'*' as qInv, '*' as qExpNo, '' as qFgDt, " + wCommStr
    GrdSumm.ColProp("wSummPrdCmCd").Style = fgcinvisible
    GrdSumm.ColProp("wSummPrdOrd").Style = fgcinvisible
    GrdSumm.ColProp("wSummExpCmCd").Style = fgcNormal
    GrdSumm.ColProp("wSummExpOrd").Style = fgcNormal
    GrdSumm.ColProp("wSummDmCd").Style = fgcNormal
    GrdSumm.ColProp("wSummPl").Style = fgcinvisible
    GrdSumm.ColProp("wSummInv").Style = fgcinvisible
    GrdSumm.ColProp("wSummInvExpNo").Style = fgcinvisible
    GrdSumm.ColProp("wSummFgDt").Style = fgcinvisible
  ElseIf wGrpBy = "FdDmCd" Then
    wStr = "Select '*' as qPrdCmCd, '*' as qPrdOrd, '*' as qExpCmCd, '*' as qExpOrd, " + _
           "FdDmCd as qDmCd, '*' as qPckLst, '*' as qInv, '*' as qExpNo, " + _
           "'' as qFgDt, " + wCommStr
    GrdSumm.ColProp("wSummPrdCmCd").Style = fgcinvisible
    GrdSumm.ColProp("wSummPrdOrd").Style = fgcinvisible
    GrdSumm.ColProp("wSummExpCmCd").Style = fgcinvisible
    GrdSumm.ColProp("wSummExpOrd").Style = fgcinvisible
    GrdSumm.ColProp("wSummDmCd").Style = fgcNormal
    GrdSumm.ColProp("wSummPl").Style = fgcinvisible
    GrdSumm.ColProp("wSummInv").Style = fgcinvisible
    GrdSumm.ColProp("wSummInvExpNo").Style = fgcinvisible
    GrdSumm.ColProp("wSummFgDt").Style = fgcinvisible
  ElseIf wGrpBy = "FdPlChr,FdPlYy,FdPlNo" Then
    wStr = "Select '*' as qPrdCmCd, '*' as qPrdOrd, '*' as qExpCmCd, '*' as qExpOrd, " + _
           "'*' as qDmCd, FdPlChr+'/'+FdPlYy+'/'+LTrim(Str(FdPlNo)) as qPckLst, " + _
           "'*' as qInv, '*' as qExpNo, '' as qFgDt, " + wCommStr
    GrdSumm.ColProp("wSummPrdCmCd").Style = fgcinvisible
    GrdSumm.ColProp("wSummPrdOrd").Style = fgcinvisible
    GrdSumm.ColProp("wSummExpCmCd").Style = fgcinvisible
    GrdSumm.ColProp("wSummExpOrd").Style = fgcinvisible
    GrdSumm.ColProp("wSummDmCd").Style = fgcinvisible
    GrdSumm.ColProp("wSummPl").Style = fgcNormal
    GrdSumm.ColProp("wSummInv").Style = fgcinvisible
    GrdSumm.ColProp("wSummInvExpNo").Style = fgcinvisible
    GrdSumm.ColProp("wSummFgDt").Style = fgcinvisible
  ElseIf wGrpBy = "IfTc,IfYy,IfChr,IfNo,IfSr" Then
    wStr = "Select '*' as qPrdCmCd, '*' as qPrdOrd, '*' as qExpCmCd, '*' as qExpOrd, " + _
           "(case when IfSr=0 then '*' else max(FdDmCd) end) as qDmCd, '*' as qPckLst, IfSr as qInv, '*' as qExpNo, " + _
           "'' as qFgDt, " + wCommStr
    GrdSumm.ColProp("wSummPrdCmCd").Style = fgcinvisible
    GrdSumm.ColProp("wSummPrdOrd").Style = fgcinvisible
    GrdSumm.ColProp("wSummExpCmCd").Style = fgcinvisible
    GrdSumm.ColProp("wSummExpOrd").Style = fgcinvisible
    GrdSumm.ColProp("wSummDmCd").Style = fgcNormal
    GrdSumm.ColProp("wSummPl").Style = fgcinvisible
    GrdSumm.ColProp("wSummInv").Style = fgcNormal
    GrdSumm.ColProp("wSummInvExpNo").Style = fgcinvisible
    GrdSumm.ColProp("wSummFgDt").Style = fgcinvisible
  ElseIf wGrpBy = "InExpNo" Then
    wStr = "Select '*' as qPrdCmCd, '*' as qPrdOrd, '*' as qExpCmCd, '*' as qExpOrd, " + _
           "'*' as qDmCd, '*' as qPckLst, '*' as qInv, InExpNo as qExpNo, " + _
           "'' as qFgDt, " + wCommStr
    GrdSumm.ColProp("wSummPrdCmCd").Style = fgcinvisible
    GrdSumm.ColProp("wSummPrdOrd").Style = fgcinvisible
    GrdSumm.ColProp("wSummExpCmCd").Style = fgcinvisible
    GrdSumm.ColProp("wSummExpOrd").Style = fgcinvisible
    GrdSumm.ColProp("wSummDmCd").Style = fgcinvisible
    GrdSumm.ColProp("wSummPl").Style = fgcinvisible
    GrdSumm.ColProp("wSummInv").Style = fgcinvisible
    GrdSumm.ColProp("wSummInvExpNo").Style = fgcNormal
    GrdSumm.ColProp("wSummFgDt").Style = fgcinvisible
  ElseIf wGrpBy = "FgDt" Then
    wStr = "Select '*' as qPrdCmCd, '*' as qPrdOrd, '*' as qExpCmCd, '*' as qExpOrd, " + _
           "'*' as qDmCd, '*' as qPckLst, '*' as qInv, '*' as qExpNo, " + _
           "FgDt as qFgDt, " + wCommStr
    GrdSumm.ColProp("wSummPrdCmCd").Style = fgcinvisible
    GrdSumm.ColProp("wSummPrdOrd").Style = fgcinvisible
    GrdSumm.ColProp("wSummExpCmCd").Style = fgcinvisible
    GrdSumm.ColProp("wSummExpOrd").Style = fgcinvisible
    GrdSumm.ColProp("wSummDmCd").Style = fgcinvisible
    GrdSumm.ColProp("wSummPl").Style = fgcinvisible
    GrdSumm.ColProp("wSummInv").Style = fgcinvisible
    GrdSumm.ColProp("wSummInvExpNo").Style = fgcinvisible
    GrdSumm.ColProp("wSummFgDt").Style = fgcNormal
  End If
  
  Dim wRsSumm As MwfLib.MDORowSet
  Set wRsSumm = moCn.OpenRes(wStr)
  
  GrdSumm.AllowAdd = True
  GrdSumm.AllowDelete = True

  GrdSumm.Rows = 1
  With GrdSumm
    Do While Not (wRsSumm.EOF Or wRsSumm.BOF)
      .AddItem
      wRow = .Rows - 1
      .Value(wRow, "wSummPrdCmCd") = wRsSumm!qPrdCmCd
      .Value(wRow, "wSummPrdOrd") = wRsSumm!qPrdOrd
      .Value(wRow, "wSummExpCmCd") = wRsSumm!qExpCmCd
      .Value(wRow, "wSummExpOrd") = wRsSumm!qExpOrd
      .Value(wRow, "wSummDmCd") = wRsSumm!qDmCd
      .Value(wRow, "wSummPl") = wRsSumm!qPckLst
      .Value(wRow, "wSummInv") = wRsSumm!qInv
      .Value(wRow, "wSummInvExpNo") = wRsSumm!qExpNo
      .Value(wRow, "wSummFgDt") = wRsSumm!qFgDt
      .Value(wRow, "wSummBagCount") = wRsSumm!qBags
      .Value(wRow, "wSummQty") = wRsSumm!qBagQty
      .Value(wRow, "wSummWt") = wRsSumm!qRmWt
      
      wTotRmWt = wTotRmWt + wRsSumm!qRmWt
      wTotBagQty = wTotBagQty + wRsSumm!qBagQty
      wTotBags = wTotBags + wRsSumm!qBags
      wRsSumm.MoveNext
    Loop
  End With
  ADC("wSummTotWt") = wTotRmWt
  ADC("wSummTotQty") = wTotBagQty
  ADC("wSummTotBags") = wTotBags
  
  GrdSumm.AllowAdd = False
  GrdSumm.AllowDelete = False
  Set wRsSumm = Nothing
End Sub
Private Sub CmpString(ByVal mTyp As String)
' Seting values for CmpStr of 'K'ey fields, for Adc.RepCnd (Making Copy fields scope = '')
  
  If mTyp = "K" Then
    ADC("wScPrdCmCd").CmpStr = "FdPrdCmCd= "
    ADC("wScExpCmCd").CmpStr = "FdExpCmCd= "
    ADC("wScPrdOdTc").CmpStr = "FdPrdOdTc= "
    ADC("wScPrdOdYy").CmpStr = "FdPrdOdYy= "
    ADC("wScPrdOdChr").CmpStr = "FdPrdOdChr= "
    ADC("wScPrdOdNo").CmpStr = "FdPrdOdNo= "
    ADC("wScPrdOdSr").CmpStr = "FdPrdOdSr= "
    ADC("wScExpOdTc").CmpStr = "FdExpOdTc= "
    ADC("wScExpOdYy").CmpStr = "FdExpOdYy= "
    ADC("wScExpOdChr").CmpStr = "FdExpOdChr= "
    ADC("wScExpOdNo").CmpStr = "FdExpOdNo= "
    ADC("wScExpOdSr").CmpStr = "FdExpOdSr= "
    ADC("wScDmCd").CmpStr = "FdDmCd= "
    ADC("wScPlChr").CmpStr = "FdPlChr= "
    ADC("wScPlYy").CmpStr = "FdPlYy= "
    ADC("wScPlNo").CmpStr = "FdPlNo= "
    'ADC("wScIdSr").CmpStr = "IfSr= "       '*** Compare Str not to be written for this field
    ADC("wScFgDtFr").CmpStr = "FgDt>= "
    ADC("wScFgDtTo").CmpStr = "FgDt<= "
    
    ADC("wCpPrdCmCd").CmpStr = ""
    ADC("wCpExpCmCd").CmpStr = ""
    ADC("wCpCtCd").CmpStr = ""
    ADC("wCpPrdOdTc").CmpStr = ""
    ADC("wCpPrdOdYy").CmpStr = ""
    ADC("wCpPrdOdChr").CmpStr = ""
    ADC("wCpPrdOdNo").CmpStr = ""
    ADC("wCpPrdOdSr").CmpStr = ""
    ADC("wCpExpOdTc").CmpStr = ""
    ADC("wCpExpOdYy").CmpStr = ""
    ADC("wCpExpOdChr").CmpStr = ""
    ADC("wCpExpOdNo").CmpStr = ""
    ADC("wCpExpOdSr").CmpStr = ""
    ADC("wCpDmCtg").CmpStr = ""       ' ***** Manali 3.5.0 - 11/11/08 - DmCtg added in scope for copy
    ADC("wCpDmCd").CmpStr = ""
    ADC("wCpPlChr").CmpStr = ""
    ADC("wCpPlYy").CmpStr = ""
    ADC("wCpPlNo").CmpStr = ""
    ADC("wCpInTc").CmpStr = ""
    ADC("wCpInYy").CmpStr = ""
    ADC("wCpInChr").CmpStr = ""
    ADC("wCpInNo").CmpStr = ""
    ADC("wCpIdSr").CmpStr = ""
    ADC("wCpInExpNo").CmpStr = ""
    ADC("wCpFgDtFr").CmpStr = ""
    ADC("wCpFgDtTo").CmpStr = ""
    ADC("wCpFgSubLoc").CmpStr = ""
    '6.1_11 setting default value for other than copy
    ADC("wCpSnTc").CmpStr = ""
    ADC("wCpSnYy").CmpStr = ""
    ADC("wCpSnChr").CmpStr = ""
    ADC("wCpSnNo").CmpStr = ""
    
' Seting values for CmpStr of 'C'opy fields, for Adc.RepCnd (Making Entry fields scope = '')
  ElseIf mTyp = "C" Then
    ADC("wCpPrdCmCd").CmpStr = "FdPrdCmCd= "
    ADC("wCpExpCmCd").CmpStr = "FdExpCmCd= "
    ADC("wCpCtCd").CmpStr = IIF(UCase(ms_TcTyp) = UCase(ctCTBTcTyp), "BOdCtCd= ", " ")
    ADC("wCpPrdOdTc").CmpStr = "FdPrdOdTc= "
    ADC("wCpPrdOdYy").CmpStr = "FdPrdOdYy= "
    ADC("wCpPrdOdChr").CmpStr = "FdPrdOdChr= "
    ADC("wCpPrdOdNo").CmpStr = "FdPrdOdNo= "
    ADC("wCpPrdOdSr").CmpStr = "FdPrdOdSr= "
    ADC("wCpExpOdTc").CmpStr = "FdExpOdTc= "
    ADC("wCpExpOdYy").CmpStr = "FdExpOdYy= "
    ADC("wCpExpOdChr").CmpStr = "FdExpOdChr= "
    ADC("wCpExpOdNo").CmpStr = "FdExpOdNo= "
    ADC("wCpExpOdSr").CmpStr = "FdExpOdSr= "
    ADC("wCpDmCtg").CmpStr = "DmCtg= "      ' ***** Manali 3.5.0 - 11/11/08 - DmCtg added in scope for copy
    ADC("wCpDmCd").CmpStr = "FdDmCd= "
    ADC("wCpPlChr").CmpStr = "FdPlChr= "
    ADC("wCpPlYy").CmpStr = "FdPlYy= "
    ADC("wCpPlNo").CmpStr = "FdPlNo= "
    ADC("wCpInTc").CmpStr = "FdInTc= "
    ADC("wCpInYy").CmpStr = "FdInYy= "
    ADC("wCpInChr").CmpStr = "FdInChr= "
    ADC("wCpInNo").CmpStr = "FdInNo= "
    ADC("wCpIdSr").CmpStr = "FdInSr= "
    ADC("wCpInExpNo").CmpStr = "FdInExpNo= "
    ADC("wCpFgDtFr").CmpStr = "FgDt>= "
    ADC("wCpFgDtTo").CmpStr = "FgDt<= "
    ADC("wCpFgSubLoc").CmpStr = "BFgSubLoc= "
    '6.1_11 cmpstr for shpment note filters
    ADC("wCpSnTc").CmpStr = "SnTc="
    ADC("wCpSnYy").CmpStr = "SnYy="
    ADC("wCpSnChr").CmpStr = "SnChr="
    ADC("wCpSnNo").CmpStr = "SnNo="
    
    ADC("wScPrdCmCd").CmpStr = ""
    ADC("wScExpCmCd").CmpStr = ""
    ADC("wScPrdOdTc").CmpStr = ""
    ADC("wScPrdOdYy").CmpStr = ""
    ADC("wScPrdOdChr").CmpStr = ""
    ADC("wScPrdOdNo").CmpStr = ""
    ADC("wScPrdOdSr").CmpStr = ""
    ADC("wScExpOdTc").CmpStr = ""
    ADC("wScExpOdYy").CmpStr = ""
    ADC("wScExpOdChr").CmpStr = ""
    ADC("wScExpOdNo").CmpStr = ""
    ADC("wScExpOdSr").CmpStr = ""
    ADC("wScDmCd").CmpStr = ""
    ADC("wScPlChr").CmpStr = ""
    ADC("wScPlYy").CmpStr = ""
    ADC("wScPlNo").CmpStr = ""
    'ADC("wScIdSr").CmpStr = ""       '*** Compare Str not to be written for this field
    ADC("wScFgDtFr").CmpStr = ""
    ADC("wScFgDtTo").CmpStr = ""
  End If
End Sub
Private Sub DispFra(ByVal pv_InvBagFra As en_InvBagFra)
  '*** This is the common procedure which is used to toggle between hiding and displaying
      'the frames (i.e. if the frame is visible then make it invisible and viseversa)
      'associated with a command button
  
  Select Case pv_InvBagFra
  Case Is = Cpy
    If FraCP.Visible = True Then
      FraCP.Visible = False
      FraCP.Enabled = False
      '*** required to pass the array position Cpy-0, to resotre the 'Causes Validation' State
      Call EnaDisaCmds(False, CmdCP, Cpy)
      CmdCP.SetFocus
    Else
      FraCP.Visible = True
      FraCP.Enabled = True
      ADC("wCpPrdCmCd").SetFocus
      FraCP.ZOrder
      Call EnaDisaCmds(True, CmdCP)
    End If
  Case Is = Summ
    If FraSumm.Visible = True Then
      FraSumm.Visible = False
      FraSumm.Enabled = False
      Call EnaDisaCmds(False, CmdSumm, Summ)
      CmdSumm.SetFocus
    Else
      FraSumm.Visible = True
      FraSumm.Enabled = True
      ADC("wSummBagGrp").SetFocus
      FraSumm.ZOrder
      Call EnaDisaCmds(True, CmdSumm)
    End If
  Case Is = Scop
    If FraSC.Visible = True Then
      FraSC.Visible = False
      FraSC.Enabled = False
      Call EnaDisaCmds(False, CmdSC, Scop)
      CmdSC.SetFocus
    Else
      FraSC.Visible = True
      FraSC.Enabled = True
      ADC("wScPrdCmCd").SetFocus
      FraSC.ZOrder
      Call EnaDisaCmds(True, CmdSC)
    End If
  ' ***** Manali 3.10.0 - 19/04/12 - Cpy Bags from RfId
  Case Is = CpyRfId
    If FraCpyRfId.Visible = True Then
      FraCpyRfId.Visible = False
      FraCpyRfId.Enabled = False
      Call EnaDisaCmds(False, CmdCpyRfId, CpyRfId)
      CmdCpyRfId.SetFocus
    Else
      FraCpyRfId.Visible = True
      FraCpyRfId.Enabled = True
      ADC("wRfIdSel").SetFocus
      FraCpyRfId.ZOrder
      Call EnaDisaCmds(True, CmdCpyRfId)
    End If
  ' ***** Manali 3.10.0 - 19/04/12 - Cpy Bags from RfId
  End Select
End Sub
Private Sub EnaDisaCmds(ByVal pv_ShowFra As Boolean, Optional ByVal pv_CmdBut As MWCTL_BTN1, Optional ByVal pv_InvBagFra As en_InvBagFra)
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
      'pv_CmdBut.CausesValidation = False
      '*** restoring the 'Causes Validation' state of the current pv_cmdbutton
      If Not IsMissing(pv_InvBagFra) Then
        pv_CmdBut.CausesValidation = mArr_Cv(pv_InvBagFra)
      End If
    End If
  End If
  FraNKeyAll.Enabled = Not pv_ShowFra
  ADC.AllowSave = Not pv_ShowFra
  
  If pv_ShowFra = True Then
    '** All command buttons are disabled
    CmdCP.Enabled = Not pv_ShowFra
    CmdSumm.Enabled = Not pv_ShowFra
    CmdSC.Enabled = Not pv_ShowFra
    CmdCpyRfId.Enabled = Not pv_ShowFra     ' ***** Manali 3.10.0 - 19/04/12 - Cpy Bags from RfId
    CmdDelAllBags.Enabled = Not pv_ShowFra
  Else
    '*** Restoring Ena/Disa state of all Command Buttons
    CmdCP.Enabled = mArr_Ed(Cpy)
    CmdSumm.Enabled = mArr_Ed(Summ)
    CmdSC.Enabled = mArr_Ed(Scop)
    CmdCpyRfId.Enabled = mArr_Ed(CpyRfId)       ' ***** Manali 3.10.0 - 19/04/12 - Cpy Bags from RfId
    CmdDelAllBags.Enabled = mArr_Ed(DelAllBag)
  End If
    
  If Not pv_CmdBut Is Nothing Then pv_CmdBut.Enabled = True
End Sub
Private Sub HideAllFras()
  '*** This procedure hides all the frames other than FraNKeyAll

  FraNKeyAll.Visible = True
  FraCP.Visible = False
  FraCP.Enabled = False
  FraSumm.Visible = False
  FraSumm.Enabled = False
  FraSC.Visible = False
  FraSC.Enabled = False
  ' ***** Manali 3.10.0 - 19/04/12 - Cpy Bags from RfId
  FraCpyRfId.Visible = False
  FraCpyRfId.Enabled = False
  ' ***** Manali 3.10.0 - 19/04/12 - Cpy Bags from RfId
End Sub
Private Sub StoreState()
  '*** Procedure to store the Causesvalidation and Ena/Disa state of command buttons
  mArr_Cv(Cpy) = CmdCP.CausesValidation
  mArr_Cv(Summ) = CmdSumm.CausesValidation
  mArr_Cv(Scop) = CmdSC.CausesValidation
  mArr_Cv(CpyRfId) = CmdCpyRfId.CausesValidation      ' ***** Manali 3.10.0 - 19/04/12 - Cpy Bags from RfId
  mArr_Cv(DelAllBag) = CmdDelAllBags.CausesValidation
  
  mArr_Ed(Cpy) = CmdCP.Enabled
  mArr_Ed(Summ) = CmdSumm.Enabled
  mArr_Ed(Scop) = CmdSC.Enabled
  mArr_Ed(CpyRfId) = CmdCpyRfId.Enabled         ' ***** Manali 3.10.0 - 19/04/12 - Cpy Bags from RfId
   mArr_Ed(DelAllBag) = CmdDelAllBags.Enabled
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
  ws_BtnStr = "CmdCpGo"
  Call SetObjColors(Me, ws_BtnStr)
  '*** Jenny Colour
  
  Me.KeyPreview = True      ' ***** Manali 3.10.0 - 18/04/12 - Bar Code
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
Private Sub GrdInvFgd_Validate(Cancel As Boolean)
  Cancel = GrdInvFgd.Validate
End Sub

Private Sub RMSummary(ByVal mGrpByCd As String, ByVal mVchYN As String)
  
  '***** Sachin 3.11.0 - RM Summary in Invoice Bags entry
  Dim wGrpByCd  As String, wGrpBy As String, wCommStr As String
  Dim wRmTotWt As Single, wRmTotQty As Single, wStr As String, wRow As Single
  Dim wVchSummJoin As String, wSummCnd As String, wCurrRow As Integer
  
  If GrdInvFgd.Rows <= 1 Then
    DispMsg "Invalid Bag No", etInfo
    GrdBRS.AllowAdd = True: GrdVRS.AllowAdd = True
    GrdBRS.AllowDelete = True: GrdVRS.AllowDelete = True
    GrdBRS.Rows = 1: GrdVRS.Rows = 1
    GrdBRS.AllowAdd = False: GrdVRS.AllowAdd = False
    GrdBRS.AllowDelete = False: GrdVRS.AllowDelete = False
    Exit Sub
  End If
  
  wGrpByCd = mGrpByCd
  If wGrpByCd = "RMCTG" Then wGrpBy = "RmCtg"
  If wGrpByCd = "RMSCTG" Then wGrpBy = "RmCtg,RmSCtg"
  If wGrpByCd = "RMCD" Then wGrpBy = "RmCtg,RmSCtg,RmCd"
  wCurrRow = IIF(GrdInvFgd.Row = 0, 1, GrdInvFgd.Row)
  
  If mVchYN = "N" Then
    wVchSummJoin = ""
    wSummCnd = " FdCoCd ='" + gs_CoCd + "' and  FdBYy ='" + GrdInvFgd.Value(wCurrRow, "IfFdBYy") + "' " + _
             " And  FdBChr ='" + GrdInvFgd.Value(wCurrRow, "IfFdBChr") + "' And  " + _
             " FdBNo =" + CStr(GrdInvFgd.Value(wCurrRow, "IfFdBNo")) + " And FdPrtKey=(Case When '" + gs_Partition + "'='" + ctCurrPrtn + "' Then '" + ctCurrPrtn + "' Else FdPrtKey End)"
    LblMain(2).Caption = "Bag Level Summary for " + GrdInvFgd.Value(wCurrRow, "IfFdBYy") + "/" + GrdInvFgd.Value(wCurrRow, "IfFdBChr") + "/" + CStr(GrdInvFgd.Value(wCurrRow, "IfFdBNo"))
  Else
    wVchSummJoin = " Join InvFgd On IfCoCd=FdCoCd And IfFdBYy=FdBYy And IfFdBChr=FdBChr And IfFdBNo=FdBNo "
    wSummCnd = " IfCoCd='" + gs_CoCd + "' and IfTc='" + ADC("wIfTc") + "' and " + _
               " IfYy='" + ADC("wIfYy") + "' and IfChr='" + ADC("wIfChr") + "' and " + _
               " IfNo= " + CStr(ADC("wIfNo")) + " and IfPrtKey=(Case When '" + gs_Partition + "'='" + ctCurrPrtn + "' Then '" + ctCurrPrtn + "' Else IfPrtKey End)"
  End If
  
  wRmTotWt = 0: wRmTotQty = 0: wStr = ""
  wCommStr = " Sum(FrRmQty) as RmQty, Sum(FrRmWt*(Case When FdTc='FB' Then 1 Else -1 End)) as RmWt from Fgrm " + _
             "    Join RmMst On RmCd= FrRmCd " + _
             "    Join Fgd On FdCoCd=FrCoCd And FdTc=FrTc And FdYy=FrYy And FdChr=FrChr And FdNo=FrNo And FdSr=FrSr " + wVchSummJoin + _
             " Where " + wSummCnd + _
             " Group By " + wGrpBy + " " + _
             " Order by RmCtg, RmSCtg, RmCd "
  
  If wGrpBy = "RmCtg" Then
    wStr = "Select RmCtg, '*' as RmSCtg, '*' as RmCd, " + wCommStr
  ElseIf wGrpBy = "RmCtg,RmSCtg" Then
    wStr = "Select RmCtg, RmSCtg, '*' as RmCd, " + wCommStr
  ElseIf wGrpBy = "RmCtg,RmSCtg,RmCd" Then
    wStr = "Select RmCtg, RmSCtg, RmCd, " + wCommStr
  End If
  
  Dim wRsRS As MwfLib.MDORowSet
  Set wRsRS = moCn.OpenRes(wStr)
  If mVchYN = "N" Then
    GrdBRS.AllowAdd = True
    GrdBRS.AllowDelete = True
    GrdBRS.Rows = 1
    With GrdBRS
      Do While Not (wRsRS.EOF Or wRsRS.BOF)
        .AddItem
        wRow = .Rows - 1
        .Value(wRow, "wBRSRmCtg") = wRsRS!RmCtg
        .Value(wRow, "wBRSRmSCtg") = wRsRS!RmSCtg
        .Value(wRow, "wBRSRmCd") = wRsRS!RmCd
        .Value(wRow, "wBRSRmQty") = wRsRS!RmQty
        .Value(wRow, "wBRSRmWt") = wRsRS!rmwt
            Select Case Trim(wRsRS!RmCtg)
            Case Is = "G", "P", "S", "L", "M", "X"
              wRmTotWt = wRmTotWt + wRsRS!rmwt
            Case Is = "D", "C"
              wRmTotWt = wRmTotWt + MWLib.Div(wRsRS!rmwt, 5)
            End Select
        wRmTotQty = wRmTotQty + wRsRS!RmQty
        wRsRS.MoveNext
      Loop
    End With
    ADC("wBRSRmTotWt") = wRmTotWt
    ADC("wBRSRmTotQty") = wRmTotQty
    GrdBRS.AllowAdd = False
    GrdBRS.AllowDelete = False
    Set wRsRS = Nothing
  Else
    GrdVRS.AllowAdd = True
    GrdVRS.AllowDelete = True
    GrdVRS.Rows = 1
    With GrdVRS
      Do While Not (wRsRS.EOF Or wRsRS.BOF)
        .AddItem
        wRow = .Rows - 1
        .Value(wRow, "wVRSRmCtg") = wRsRS!RmCtg
        .Value(wRow, "wVRSRmSCtg") = wRsRS!RmSCtg
        .Value(wRow, "wVRSRmCd") = wRsRS!RmCd
        .Value(wRow, "wVRSRmQty") = wRsRS!RmQty
        .Value(wRow, "wVRSRmWt") = wRsRS!rmwt
            Select Case Trim(wRsRS!RmCtg)
            Case Is = "G", "P", "S", "L", "M", "X"
              wRmTotWt = wRmTotWt + wRsRS!rmwt
            Case Is = "D", "C"
              wRmTotWt = wRmTotWt + MWLib.Div(wRsRS!rmwt, 5)
            End Select
        wRmTotQty = wRmTotQty + wRsRS!RmQty
        wRsRS.MoveNext
      Loop
    End With
    ADC("wVRSRmTotWt") = wRmTotWt
    ADC("wVRSRmTotQty") = wRmTotQty
    GrdVRS.AllowAdd = False
    GrdVRS.AllowDelete = False
    Set wRsRS = Nothing
  End If
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

