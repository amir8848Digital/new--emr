VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Object = "{AB2CC32A-E652-42A2-B341-26ECE02F6386}#2.1#0"; "Mwfctlforspeed9.ocx"
Begin VB.Form EmrFrmRmRt 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Raw Material Rate Master"
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
   Begin MSComDlg.CommonDialog CdImpExc 
      Left            =   3840
      Top             =   9360
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin MwfCtl.MWCTL_BTN1 CmdImpExc 
      Height          =   465
      Left            =   4440
      TabIndex        =   87
      ToolTipText     =   "Copy Option"
      Top             =   9360
      Width           =   960
      _ExtentX        =   1693
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
   Begin MwfCtl.MWCTL_BTN1 CmdCpy 
      Height          =   465
      Left            =   7320
      TabIndex        =   20
      ToolTipText     =   "Option To Copy The Rates From Another Customer"
      Top             =   9360
      Width           =   945
      _ExtentX        =   1667
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
      Left            =   10320
      TabIndex        =   38
      TabStop         =   0   'False
      Top             =   9360
      Visible         =   0   'False
      Width           =   525
      _ExtentX        =   926
      _ExtentY        =   714
   End
   Begin VB.Frame FRA_KEYS 
      BorderStyle     =   0  'None
      Height          =   570
      Left            =   0
      TabIndex        =   41
      Top             =   24
      Width           =   15090
      Begin MwfCtl.MWCTL_MED ATXT 
         Height          =   285
         Index           =   0
         Left            =   3930
         TabIndex        =   4
         ToolTipText     =   "Enter From Raw Material Sub Category"
         Top             =   0
         Width           =   885
         _ExtentX        =   1561
         _ExtentY        =   503
         MaxLength       =   5
         DataType        =   4
         IdName          =   "WRRSCTGFR"
         CmpStr          =   "RrSCtg>="
      End
      Begin MwfCtl.MWCTL_MED ATXT 
         Height          =   285
         Index           =   3
         Left            =   3930
         TabIndex        =   5
         ToolTipText     =   "Enter To Raw Material Sub Category"
         Top             =   285
         Width           =   885
         _ExtentX        =   1561
         _ExtentY        =   503
         MaxLength       =   5
         DataType        =   4
         IdName          =   "WRRSCTGTO"
         CmpStr          =   "RrSCtg<="
      End
      Begin MwfCtl.MWCTL_MED ATXT 
         Height          =   285
         Index           =   4
         Left            =   5415
         TabIndex        =   6
         ToolTipText     =   "Enter From Raw Material Code"
         Top             =   0
         Width           =   2370
         _ExtentX        =   4180
         _ExtentY        =   503
         MaxLength       =   16
         DataType        =   4
         IdName          =   "WRRCDFR"
         CmpStr          =   "RrCd>="
      End
      Begin MwfCtl.MWCTL_MED ATXT 
         Height          =   285
         Index           =   1
         Left            =   2700
         TabIndex        =   3
         ToolTipText     =   "Enter Raw Material Category"
         Top             =   285
         Width           =   480
         _ExtentX        =   847
         _ExtentY        =   503
         MaxLength       =   2
         DataType        =   4
         IdName          =   "WRRCTG"
         CmpStr          =   "RrCtg="
      End
      Begin MwfCtl.MWCTL_MED ATXT 
         Height          =   285
         Index           =   2
         Left            =   5415
         TabIndex        =   7
         ToolTipText     =   "Enter To Raw Material Code"
         Top             =   285
         Width           =   2370
         _ExtentX        =   4180
         _ExtentY        =   503
         MaxLength       =   16
         DataType        =   4
         IdName          =   "WRRCDTO"
         CmpStr          =   "RrCd<="
      End
      Begin MwfCtl.MWCTL_MED ATXT 
         Height          =   285
         Index           =   5
         Left            =   850
         TabIndex        =   1
         ToolTipText     =   "Enter Sort Option"
         Top             =   285
         Width           =   1140
         _ExtentX        =   2011
         _ExtentY        =   503
         MaxLength       =   8
         DataType        =   4
         IdName          =   "WRRSRT"
      End
      Begin MwfCtl.MWCTL_MED ATXT 
         Height          =   285
         Index           =   7
         Left            =   850
         TabIndex        =   0
         ToolTipText     =   "Enter Rate Type"
         Top             =   0
         Width           =   630
         _ExtentX        =   1111
         _ExtentY        =   503
         MaxLength       =   3
         DataType        =   4
         IdName          =   "WRRTCTYP"
         CmpStr          =   "RrTcTyp="
      End
      Begin MwfCtl.MWCTL_MED ATXT 
         Height          =   285
         Index           =   8
         Left            =   10245
         TabIndex        =   10
         ToolTipText     =   "Enter From Customer"
         Top             =   0
         Width           =   1200
         _ExtentX        =   2117
         _ExtentY        =   503
         MaxLength       =   8
         DataType        =   4
         IdName          =   "WRRCMCDFR"
         CmpStr          =   "RrCmCd>="
      End
      Begin MwfCtl.MWCTL_MED ATXT 
         Height          =   285
         Index           =   9
         Left            =   10245
         TabIndex        =   11
         ToolTipText     =   "Enter To Customer"
         Top             =   285
         Width           =   1200
         _ExtentX        =   2117
         _ExtentY        =   503
         MaxLength       =   8
         DataType        =   4
         IdName          =   "WRRCMCDTO"
         CmpStr          =   "RrCmCd<="
      End
      Begin MwfCtl.MWCTL_MED ATXT 
         Height          =   285
         Index           =   10
         Left            =   2700
         TabIndex        =   2
         ToolTipText     =   "Enter Customer Category"
         Top             =   0
         Width           =   480
         _ExtentX        =   847
         _ExtentY        =   503
         MaxLength       =   2
         DataType        =   4
         IdName          =   "WRRCMCTG"
         CmpStr          =   "RrCmCtg="
      End
      Begin MwfCtl.MWCTL_MED ATXT 
         Height          =   285
         Index           =   19
         Left            =   850
         TabIndex        =   62
         ToolTipText     =   "Enter YyyyMm"
         Top             =   0
         Width           =   770
         _ExtentX        =   1349
         _ExtentY        =   503
         Mask            =   "#####0"
         MaxLength       =   6
         DataType        =   1
         IdName          =   "WRHYYYYMM"
      End
      Begin MwfCtl.MWCTL_MED ATXT 
         Height          =   285
         Index           =   28
         Left            =   11895
         TabIndex        =   12
         ToolTipText     =   "Enter From Currency Code"
         Top             =   0
         Width           =   780
         _ExtentX        =   1376
         _ExtentY        =   503
         MaxLength       =   5
         DataType        =   4
         IdName          =   "WRRCMCURCDFR"
         CmpStr          =   "RrCmCurCd>="
      End
      Begin MwfCtl.MWCTL_MED ATXT 
         Height          =   285
         Index           =   29
         Left            =   11895
         TabIndex        =   13
         ToolTipText     =   "Enter To Currency Code"
         Top             =   285
         Width           =   780
         _ExtentX        =   1376
         _ExtentY        =   503
         MaxLength       =   5
         DataType        =   4
         IdName          =   "WRRCMCURCDTO"
         CmpStr          =   "RrCmCurCd<="
      End
      Begin MwfCtl.MWCTL_MED ATXT 
         Height          =   285
         Index           =   27
         Left            =   12960
         TabIndex        =   14
         ToolTipText     =   "Enter From Size"
         Top             =   0
         Width           =   900
         _ExtentX        =   1588
         _ExtentY        =   503
         Alignment       =   2
         Mask            =   "####0.0000"
         MaxLength       =   10
         DataType        =   2
         IdName          =   "WRRSZFR"
         CmpStr          =   "RrFrSz>="
      End
      Begin MwfCtl.MWCTL_MED ATXT 
         Height          =   285
         Index           =   30
         Left            =   12960
         TabIndex        =   15
         ToolTipText     =   "Enter To Size"
         Top             =   285
         Width           =   900
         _ExtentX        =   1588
         _ExtentY        =   503
         Alignment       =   2
         Mask            =   "####0.0000"
         MaxLength       =   10
         DataType        =   2
         IdName          =   "WRRSZTO"
         CmpStr          =   "RrFrSz<="
      End
      Begin MwfCtl.MWCTL_MED ATXT 
         Height          =   285
         Index           =   33
         Left            =   8410
         TabIndex        =   8
         ToolTipText     =   "Enter From Design Code"
         Top             =   0
         Width           =   1380
         _ExtentX        =   2434
         _ExtentY        =   503
         MaxLength       =   12
         DataType        =   4
         IdName          =   "WRRDMCDFR"
         CmpStr          =   "RrDmCd>="
      End
      Begin MwfCtl.MWCTL_MED ATXT 
         Height          =   285
         Index           =   34
         Left            =   8410
         TabIndex        =   9
         ToolTipText     =   "Enter To Design  Code"
         Top             =   285
         Width           =   1380
         _ExtentX        =   2434
         _ExtentY        =   503
         MaxLength       =   12
         DataType        =   4
         IdName          =   "WRRDMCDTO"
         CmpStr          =   "RrDmCd<="
      End
      Begin MwfCtl.MWCTL_MED ATXT 
         Height          =   285
         Index           =   35
         Left            =   14235
         TabIndex        =   16
         ToolTipText     =   "Enter From Size To Weight "
         Top             =   0
         Width           =   900
         _ExtentX        =   1588
         _ExtentY        =   503
         MaxLength       =   5
         DataType        =   4
         IdName          =   "WRRSTWGRPFR"
         CmpStr          =   "RrSTWGrp>="
      End
      Begin MwfCtl.MWCTL_MED ATXT 
         Height          =   285
         Index           =   36
         Left            =   14235
         TabIndex        =   17
         ToolTipText     =   "Enter To Size To Weight "
         Top             =   285
         Width           =   900
         _ExtentX        =   1588
         _ExtentY        =   503
         MaxLength       =   5
         DataType        =   4
         IdName          =   "WRRSTWGRPTO"
         CmpStr          =   "RrSTWGrp<="
      End
      Begin MwfCtl.MWCTL_MED ATXT 
         Height          =   285
         Index           =   25
         Left            =   15015
         TabIndex        =   18
         ToolTipText     =   "Enter Partition Key (C-Current,P-Previous,Blank For All)"
         Top             =   270
         Width           =   330
         _ExtentX        =   582
         _ExtentY        =   503
         DataType        =   4
         IdName          =   "WRRPRTKEY"
         CmpStr          =   "RrPrtKey = "
      End
      Begin VB.Label ALBL 
         BackStyle       =   0  'Transparent
         Caption         =   "Grp"
         BeginProperty Font 
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
         Left            =   13875
         TabIndex        =   100
         Top             =   0
         Width           =   375
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
         Index           =   13
         Left            =   12705
         TabIndex        =   84
         Top             =   0
         Width           =   255
      End
      Begin VB.Label LblPrtnKey 
         BackStyle       =   0  'Transparent
         Caption         =   "Prtn.(C/P)"
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   14780
         TabIndex        =   83
         Top             =   0
         Width           =   675
      End
      Begin VB.Label ALBL 
         BackStyle       =   0  'Transparent
         Caption         =   "DsgCd"
         BeginProperty Font 
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
         Left            =   7795
         TabIndex        =   73
         Top             =   0
         Width           =   675
      End
      Begin VB.Label ALBL 
         BackStyle       =   0  'Transparent
         Caption         =   "Curr"
         BeginProperty Font 
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
         Left            =   11475
         TabIndex        =   72
         Top             =   0
         Width           =   435
      End
      Begin VB.Label ALBL 
         BackStyle       =   0  'Transparent
         Caption         =   "CustCtg"
         BeginProperty Font 
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
         Left            =   1950
         TabIndex        =   57
         Top             =   0
         Width           =   735
      End
      Begin VB.Label ALBL 
         BackStyle       =   0  'Transparent
         Caption         =   "Cust"
         BeginProperty Font 
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
         Left            =   9810
         TabIndex        =   50
         Top             =   0
         Width           =   495
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
         Left            =   60
         TabIndex        =   48
         Top             =   285
         Width           =   735
      End
      Begin VB.Label ALBL 
         BackStyle       =   0  'Transparent
         Caption         =   "Fr:"
         BeginProperty Font 
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
         Left            =   3210
         TabIndex        =   47
         Top             =   0
         Width           =   435
      End
      Begin VB.Label ALBL 
         BackStyle       =   0  'Transparent
         Caption         =   "RmCd"
         BeginProperty Font 
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
         Left            =   4830
         TabIndex        =   46
         Top             =   0
         Width           =   585
      End
      Begin VB.Label ALBL 
         BackStyle       =   0  'Transparent
         Caption         =   "To:"
         BeginProperty Font 
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
         Left            =   3210
         TabIndex        =   45
         Top             =   285
         Width           =   345
      End
      Begin VB.Label ALBL 
         BackStyle       =   0  'Transparent
         Caption         =   "SCtg"
         BeginProperty Font 
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
         Left            =   3480
         TabIndex        =   44
         Top             =   0
         Width           =   705
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
         Index           =   0
         Left            =   2030
         TabIndex        =   42
         Top             =   285
         Width           =   765
      End
      Begin VB.Label LblYyyyMm 
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
         Height          =   285
         Left            =   60
         TabIndex        =   61
         Top             =   0
         Width           =   885
      End
      Begin VB.Label LblTcTyp 
         BackStyle       =   0  'Transparent
         Caption         =   "TcType"
         BeginProperty Font 
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
         TabIndex        =   49
         Top             =   0
         Width           =   1065
      End
   End
   Begin MwfCtl.MWCTL_ADC ADC 
      CausesValidation=   0   'False
      Height          =   555
      Left            =   0
      TabIndex        =   39
      Top             =   9360
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
   Begin MwfCtl.MWCTL_BTN1 CmdPwd 
      Height          =   465
      Left            =   5400
      TabIndex        =   79
      ToolTipText     =   "Enter Password to Edit Design"
      Top             =   9360
      Width           =   960
      _ExtentX        =   1693
      _ExtentY        =   820
      BackColor       =   -2147483633
      MaskColor       =   8454143
      Caption         =   "&Password"
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
   Begin MwfCtl.MWCTL_BTN1 CmdCpyHist 
      Height          =   465
      Left            =   8250
      TabIndex        =   63
      ToolTipText     =   "Option To Copy The Rates From Another Customer"
      Top             =   9360
      Width           =   945
      _ExtentX        =   1667
      _ExtentY        =   820
      BackColor       =   -2147483633
      MaskColor       =   8454143
      Caption         =   "Cpy To &Hist"
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
   Begin MwfCtl.MWCTL_BTN1 CmdILC 
      Height          =   465
      Left            =   6360
      TabIndex        =   74
      ToolTipText     =   "Refresh Rates Thourgh ILC"
      Top             =   9360
      Width           =   945
      _ExtentX        =   1667
      _ExtentY        =   820
      BackColor       =   -2147483633
      MaskColor       =   8454143
      Caption         =   "&Refresh Rates"
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
      Height          =   8730
      Left            =   0
      TabIndex        =   40
      Top             =   600
      Width           =   15195
      Begin VB.Frame FraCpy 
         Height          =   2025
         Left            =   0
         TabIndex        =   21
         Top             =   6600
         Visible         =   0   'False
         Width           =   15105
         Begin MwfCtl.MWCTL_BTN1 CmdCpyGo 
            Height          =   285
            Left            =   14400
            TabIndex        =   37
            ToolTipText     =   "Press This Button To Copy"
            Top             =   1590
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
            Index           =   11
            Left            =   1680
            TabIndex        =   23
            ToolTipText     =   "Enter Source Customer"
            Top             =   735
            Width           =   1320
            _ExtentX        =   2328
            _ExtentY        =   503
            MaxLength       =   8
            DataType        =   4
            ReCalcParent    =   "WCPYCUSTNAMEFR,WCPYCURCDFR"
            ReCalcOn        =   "WCPYCUSTNAMEFR"
            IdName          =   "WCPYCUSTFR"
         End
         Begin MwfCtl.MWCTL_MED ATXT 
            Height          =   285
            Index           =   12
            Left            =   1680
            TabIndex        =   26
            ToolTipText     =   "Enter Target Customer"
            Top             =   1020
            Width           =   1320
            _ExtentX        =   2328
            _ExtentY        =   503
            MaxLength       =   8
            DataType        =   4
            ReCalcParent    =   "WCPYCUSTNAMETO,WCPYCURCDTO"
            ReCalcOn        =   "WCPYCUSTNAMETO"
            IdName          =   "WCPYCUSTTO"
         End
         Begin MwfCtl.MWCTL_MED ATXT 
            Height          =   285
            Index           =   13
            Left            =   1680
            TabIndex        =   31
            ToolTipText     =   "Enter From Raw Material Sub Category"
            Top             =   1305
            Width           =   885
            _ExtentX        =   1561
            _ExtentY        =   503
            MaxLength       =   5
            DataType        =   4
            IdName          =   "wCpySCtgFr"
         End
         Begin MwfCtl.MWCTL_MED ATXT 
            Height          =   285
            Index           =   14
            Left            =   4515
            TabIndex        =   32
            ToolTipText     =   "Enter From Raw Material Code"
            Top             =   1305
            Width           =   2460
            _ExtentX        =   4339
            _ExtentY        =   503
            MaxLength       =   16
            DataType        =   4
            IdName          =   "wCpyRmCdFr"
         End
         Begin MwfCtl.MWCTL_MED ATXT 
            Height          =   285
            Index           =   15
            Left            =   1680
            TabIndex        =   34
            ToolTipText     =   "Enter To Raw Material Sub Category"
            Top             =   1590
            Width           =   885
            _ExtentX        =   1561
            _ExtentY        =   503
            MaxLength       =   5
            DataType        =   4
            IdName          =   "wCpySCtgTo"
         End
         Begin MwfCtl.MWCTL_MED ATXT 
            Height          =   285
            Index           =   16
            Left            =   4515
            TabIndex        =   35
            ToolTipText     =   "Enter To Raw Material Code"
            Top             =   1590
            Width           =   2460
            _ExtentX        =   4339
            _ExtentY        =   503
            MaxLength       =   16
            DataType        =   4
            IdName          =   "wCpyRmCdTo"
         End
         Begin MwfCtl.MWCTL_MED ATXT 
            Height          =   285
            Index           =   6
            Left            =   1680
            TabIndex        =   22
            ToolTipText     =   "Enter Source Customer Category"
            Top             =   450
            Width           =   480
            _ExtentX        =   847
            _ExtentY        =   503
            MaxLength       =   2
            DataType        =   4
            IdName          =   "WCPYCUSTCTGFR"
         End
         Begin MwfCtl.MWCTL_MED ATXT 
            Height          =   285
            Index           =   17
            Left            =   4515
            TabIndex        =   24
            ToolTipText     =   "Enter Source Customer Name"
            Top             =   735
            Width           =   5190
            _ExtentX        =   9155
            _ExtentY        =   503
            MaxLength       =   45
            ReCalcParent    =   "WCPYCUSTFR"
            ReCalcOn        =   "WCPYCUSTFR"
            IdName          =   "WCPYCUSTNAMEFR"
         End
         Begin MwfCtl.MWCTL_MED ATXT 
            Height          =   285
            Index           =   18
            Left            =   4515
            TabIndex        =   27
            ToolTipText     =   "Enter Target Customer Name"
            Top             =   1020
            Width           =   5190
            _ExtentX        =   9155
            _ExtentY        =   503
            MaxLength       =   45
            ReCalcParent    =   "WCPYCUSTTO"
            ReCalcOn        =   "WCPYCUSTTO"
            IdName          =   "WCPYCUSTNAMETO"
         End
         Begin MwfCtl.MWCTL_MED ATXT 
            Height          =   285
            Index           =   20
            Left            =   10650
            TabIndex        =   25
            ToolTipText     =   "Enter From Currency  Code"
            Top             =   735
            Width           =   885
            _ExtentX        =   1561
            _ExtentY        =   503
            MaxLength       =   5
            DataType        =   4
            ReCalcParent    =   "WMULDIVFR,WCNVRTFR"
            ReCalcOn        =   "WCPYCUSTFR"
            IdName          =   "WCPYCURCDFR"
         End
         Begin MwfCtl.MWCTL_MED ATXT 
            Height          =   285
            Index           =   23
            Left            =   10650
            TabIndex        =   28
            ToolTipText     =   "Enter To Currency  Code"
            Top             =   1020
            Width           =   885
            _ExtentX        =   1561
            _ExtentY        =   503
            MaxLength       =   5
            DataType        =   4
            ReCalcParent    =   "WMULDIVFR,WCNVRTFR"
            ReCalcOn        =   "WCPYCUSTTO"
            IdName          =   "WCPYCURCDTO"
         End
         Begin MwfCtl.MWCTL_MED ATXT 
            Height          =   285
            Index           =   24
            Left            =   12600
            TabIndex        =   29
            ToolTipText     =   "Enter From Conversion Rate"
            Top             =   1020
            Width           =   1005
            _ExtentX        =   1773
            _ExtentY        =   503
            Mask            =   "###0.0000"
            MaxLength       =   9
            DataType        =   2
            ReCalcOn        =   "WCPYCURCDFR,WCPYCURCDTO"
            IdName          =   "WCNVRTFR"
         End
         Begin MwfCtl.MWCTL_MED ATXT 
            Height          =   285
            Index           =   26
            Left            =   14550
            TabIndex        =   30
            Top             =   1020
            Width           =   375
            _ExtentX        =   661
            _ExtentY        =   503
            Enabled         =   0   'False
            DataType        =   4
            ReCalcOn        =   "WCPYCURCDFR,WCPYCURCDTO"
            IdName          =   "WMULDIVFR"
         End
         Begin MwfCtl.MWCTL_MED ATXT 
            Height          =   285
            Index           =   31
            Left            =   8595
            TabIndex        =   33
            ToolTipText     =   "Enter From Size"
            Top             =   1305
            Width           =   1110
            _ExtentX        =   1958
            _ExtentY        =   503
            Alignment       =   2
            Mask            =   "####0.0000"
            MaxLength       =   10
            DataType        =   2
            IdName          =   "WCPYRMSZFR"
         End
         Begin MwfCtl.MWCTL_MED ATXT 
            Height          =   285
            Index           =   32
            Left            =   8595
            TabIndex        =   36
            ToolTipText     =   "Enter To Size"
            Top             =   1590
            Width           =   1110
            _ExtentX        =   1958
            _ExtentY        =   503
            Alignment       =   2
            Mask            =   "####0.0000"
            MaxLength       =   10
            DataType        =   2
            IdName          =   "WCPYRMSZTO"
         End
         Begin VB.Label ALBL 
            BackStyle       =   0  'Transparent
            Caption         =   "    To"
            BeginProperty Font 
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
            Left            =   7920
            TabIndex        =   86
            Top             =   1590
            Width           =   615
         End
         Begin VB.Label ALBL 
            BackStyle       =   0  'Transparent
            Caption         =   "Sz Fr"
            BeginProperty Font 
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
            Left            =   7920
            TabIndex        =   85
            Top             =   1305
            Width           =   615
         End
         Begin VB.Label LblMulDivFr 
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
            Left            =   13770
            TabIndex        =   71
            Top             =   1020
            Width           =   765
         End
         Begin VB.Label LblCnvRtFr 
            BackStyle       =   0  'Transparent
            Caption         =   "Conv Fct"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Left            =   11700
            TabIndex        =   70
            Top             =   1020
            Width           =   885
         End
         Begin VB.Label LblCurCdTo 
            BackStyle       =   0  'Transparent
            Caption         =   "Curr Cd"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Left            =   9870
            TabIndex        =   69
            Top             =   1020
            Width           =   915
         End
         Begin VB.Label LblCurCdFr 
            BackStyle       =   0  'Transparent
            Caption         =   "Curr Cd "
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Left            =   9870
            TabIndex        =   68
            Top             =   735
            Width           =   915
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
            Index           =   1
            Left            =   -30
            TabIndex        =   67
            Top             =   90
            Width           =   15105
         End
         Begin VB.Label ALBL 
            BackStyle       =   0  'Transparent
            Caption         =   "Target Name"
            BeginProperty Font 
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
            Left            =   3180
            TabIndex        =   60
            Top             =   1020
            Width           =   1425
         End
         Begin VB.Label ALBL 
            BackStyle       =   0  'Transparent
            Caption         =   "Source Name"
            BeginProperty Font 
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
            Left            =   3180
            TabIndex        =   59
            Top             =   735
            Width           =   1425
         End
         Begin VB.Label ALBL 
            BackStyle       =   0  'Transparent
            Caption         =   "Source Cust Ctg"
            BeginProperty Font 
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
            Left            =   90
            TabIndex        =   58
            Top             =   450
            Width           =   1725
         End
         Begin VB.Label LblSubCtgTo 
            BackStyle       =   0  'Transparent
            Caption         =   "              To"
            BeginProperty Font 
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
            TabIndex        =   56
            Top             =   1590
            Width           =   1215
         End
         Begin VB.Label LblRmCdFr 
            BackStyle       =   0  'Transparent
            Caption         =   "Rm Code  Fr"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Left            =   3180
            TabIndex        =   55
            Top             =   1305
            Width           =   1425
         End
         Begin VB.Label LblSubCtgFr 
            BackStyle       =   0  'Transparent
            Caption         =   "Sub Ctg  Fr"
            BeginProperty Font 
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
            TabIndex        =   54
            Top             =   1305
            Width           =   1185
         End
         Begin VB.Label LblRmCdTo 
            BackStyle       =   0  'Transparent
            Caption         =   "                To"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Left            =   3180
            TabIndex        =   53
            Top             =   1590
            Width           =   1425
         End
         Begin VB.Label ALBL 
            BackStyle       =   0  'Transparent
            Caption         =   "Target Cust"
            BeginProperty Font 
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
            TabIndex        =   52
            Top             =   1020
            Width           =   1425
         End
         Begin VB.Label ALBL 
            BackStyle       =   0  'Transparent
            Caption         =   "Source Cust"
            BeginProperty Font 
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
            Left            =   90
            TabIndex        =   51
            Top             =   735
            Width           =   1425
         End
      End
      Begin VB.Frame FraCpyHist 
         Height          =   1215
         Left            =   6600
         TabIndex        =   64
         Top             =   7440
         Visible         =   0   'False
         Width           =   4125
         Begin MwfCtl.MWCTL_BTN1 CmdCpyHistGo 
            Height          =   405
            Left            =   3360
            TabIndex        =   65
            ToolTipText     =   "Press This Button To Copy"
            Top             =   660
            Width           =   585
            _ExtentX        =   1032
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
         Begin VB.Label LblMain 
            BackStyle       =   0  'Transparent
            BorderStyle     =   1  'Fixed Single
            Caption         =   "This will copy records to History for the given range in the header, Are You Sure to Proceed? "
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   1110
            Index           =   0
            Left            =   90
            TabIndex        =   66
            Top             =   90
            Width           =   3975
         End
      End
      Begin VB.Frame FraPwd 
         Height          =   870
         Left            =   4440
         TabIndex        =   75
         Top             =   7800
         Width           =   4530
         Begin VB.TextBox TxtPwd 
            Height          =   285
            IMEMode         =   3  'DISABLE
            Left            =   2040
            MaxLength       =   30
            PasswordChar    =   "*"
            TabIndex        =   76
            Top             =   405
            Width           =   2340
         End
         Begin VB.Label ALBL 
            BackStyle       =   0  'Transparent
            Caption         =   "Password To Edit"
            BeginProperty Font 
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
            TabIndex        =   78
            Top             =   405
            Width           =   1935
         End
         Begin VB.Label LblMain 
            BackStyle       =   0  'Transparent
            BorderStyle     =   1  'Fixed Single
            Caption         =   "  Password"
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
            Index           =   1000
            Left            =   -30
            TabIndex        =   77
            Top             =   90
            Width           =   4530
         End
      End
      Begin VB.Frame FraDelRt 
         Height          =   1140
         Left            =   7440
         TabIndex        =   102
         Top             =   7440
         Width           =   4860
         Begin MwfCtl.MWCTL_BTN1 CmdDelRtGo 
            Height          =   465
            Left            =   4050
            TabIndex        =   103
            ToolTipText     =   "Option to Delete Rates"
            Top             =   600
            Width           =   690
            _ExtentX        =   1217
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
            Caption         =   $"EmrFrmRmRt.frx":0000
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   1050
            Index           =   3
            Left            =   30
            TabIndex        =   104
            Top             =   90
            Width           =   4830
         End
      End
      Begin VB.Frame FraNKeyAll 
         BorderStyle     =   0  'None
         Height          =   8715
         Left            =   0
         TabIndex        =   43
         Top             =   0
         Width           =   15120
         Begin MwfCtl.MWCTL_FLX GrdRm 
            Height          =   8580
            Left            =   0
            TabIndex        =   19
            Top             =   120
            Width           =   15105
            _ExtentX        =   26644
            _ExtentY        =   15134
            Cols            =   26
            colname1        =   "RRCMCD"
            heading1        =   "Cust Cd                "
            datafld1        =   "RrCmCd"
            datatype1       =   4
            recalcparent1   =   "GrdRm(RRCSTRT),GrdRm(RRCMCURCD)"
            maxlength1      =   8
            tooltiptext1    =   "Enter Customer Code"
            colname2        =   "RRCMCURCD"
            heading2        =   "Cur Cd     "
            datafld2        =   "RrCmCurCd"
            datatype2       =   4
            recalcon2       =   "GrdRm(RRCMCD)"
            maxlength2      =   5
            tooltiptext2    =   "Enter Currency Code"
            colname3        =   "RRSCTG"
            heading3        =   "Sub Ctg"
            datafld3        =   "RrSCtg"
            datatype3       =   4
            recalcparent3   =   "GrdRm(RRFRLN),GrdRm(RRRNGSZ),GrdRm(RRTOLN)"
            maxlength3      =   5
            tooltiptext3    =   "Enter Raw Material Sub Category"
            colname4        =   "RRSTWGRP"
            heading4        =   "STW Grp"
            datafld4        =   "RrSTWGrp"
            datatype4       =   4
            maxlength4      =   5
            tooltiptext4    =   "Enter Size To Weight Group"
            colname5        =   "RRCD"
            heading5        =   "Rm Code                                  "
            datafld5        =   "RrCd"
            datatype5       =   4
            recalcparent5   =   "GrdRm(WRRQW),GrdRm(WRRQWCST),GrdRm(WRRQW),GrdRm(WRRQWCST),GrdRm(WRRQW)"
            maxlength5      =   16
            tooltiptext5    =   "Enter Raw Material Code"
            colname6        =   "RRDMCTG"
            heading6        =   "Dsg Ctg "
            datafld6        =   "RrDmCtg"
            datatype6       =   4
            recalcparent6   =   "GrdRm(RRDMCD)"
            maxlength6      =   5
            tooltiptext6    =   "Enter Design Category"
            colname7        =   "RRDMCD"
            heading7        =   "Dm Cd                                    "
            datafld7        =   "RrDmCd"
            datatype7       =   4
            recalcon7       =   "GrdRm(RRDMCTG)"
            maxlength7      =   15
            style7          =   0
            tooltiptext7    =   "Enter Design Code"
            colname8        =   "RRDMSFX"
            heading8        =   "Suffix                     "
            datafld8        =   "RrDmSfx"
            datatype8       =   4
            maxlength8      =   15
            style8          =   0
            tooltiptext8    =   "Enter Design Suffix"
            colname9        =   "RRDMSZ"
            heading9        =   "Size         "
            datafld9        =   "RrDmSz"
            datatype9       =   4
            maxlength9      =   5
            style9          =   0
            tooltiptext9    =   "Enter Design Size"
            colname10       =   "RRLSCTG"
            heading10       =   "Loss Ctg"
            datafld10       =   "RrLsCtg"
            datatype10      =   4
            maxlength10     =   5
            tooltiptext10   =   "Enter Loss Category"
            colname11       =   "RRRMKT"
            heading11       =   "Rm KT  "
            datafld11       =   "RrRmKT"
            datatype11      =   4
            maxlength11     =   5
            style11         =   0
            tooltiptext11   =   "Enter Rm KT"
            colname12       =   "RRLABMCD"
            heading12       =   "Main Cd"
            datafld12       =   "RrLabMCd"
            datatype12      =   4
            maxlength12     =   5
            tooltiptext12   =   "Enter Labour Main Code"
            colname13       =   "RRFRLN"
            heading13       =   "Fr Size              "
            datafld13       =   "RrFrLn"
            datatype13      =   2
            mask13          =   "####0.0000"
            recalcon13      =   "GrdRm(RRSCTG)"
            maxlength13     =   10
            tooltiptext13   =   "Enter From Value"
            colname14       =   "RRTOLN"
            heading14       =   "To Size             "
            datafld14       =   "RrToLn"
            datatype14      =   2
            mask14          =   "####0.0000"
            recalcon14      =   "GrdRm(RRSCTG)"
            maxlength14     =   10
            tooltiptext14   =   "Enter To Value"
            colname15       =   "RRRNGSZ"
            heading15       =   "Range Size    "
            datafld15       =   "RrRngSz"
            datatype15      =   4
            recalcon15      =   "GrdRm(RRSCTG)"
            maxlength15     =   10
            tooltiptext15   =   "Enter Range Size"
            colname16       =   "WRRQWCST"
            heading16       =   "Q/W Cst"
            datatype16      =   4
            recalcon16      =   "GrdRm(RRCD)"
            maxlength16     =   1
            style16         =   2
            tooltiptext16   =   "Specifies Whether Rm Cost Rate Is By Qty or Wt"
            colname17       =   "RRCSTRT"
            heading17       =   "Cst Rate       "
            datafld17       =   "RrCstRt"
            datatype17      =   2
            mask17          =   "######0.000"
            recalcon17      =   "GrdRm(RRCMCD)"
            recalcparent17  =   "GrdRm(RRSALRT)"
            maxlength17     =   11
            style17         =   2
            tooltiptext17   =   "Enter Cost Rate"
            colname18       =   "WRRQW"
            heading18       =   "Q/W Sal"
            datatype18      =   4
            recalcon18      =   "GrdRm(RRCD)"
            maxlength18     =   1
            style18         =   2
            tooltiptext18   =   "Specifies Whether Rm Sale Rate Is By Qty or Wt"
            colname19       =   "RRSALRT"
            heading19       =   "Sales Rate     "
            datafld19       =   "RrSalRt"
            datatype19      =   2
            mask19          =   "######0.000"
            recalcon19      =   "GrdRm(RRCSTRT)"
            maxlength19     =   11
            tooltiptext19   =   "Enter Sales Rate"
            colname20       =   "RRXWTPER"
            heading20       =   "% Increment"
            datafld20       =   "RrXWtPer"
            datatype20      =   2
            mask20          =   "##0.00"
            maxlength20     =   6
            tooltiptext20   =   "Enter the Increment % on Wt"
            colname21       =   "RRFIXMINTOL"
            heading21       =   "Fix -ve Tol     "
            datafld21       =   "RrFixMinTol"
            datatype21      =   2
            mask21          =   "#####0.000"
            maxlength21     =   10
            tooltiptext21   =   "Enter Fixed Minimum Tolerance"
            colname22       =   "RRFIXMAXTOL"
            heading22       =   "Fix +ve Tol    "
            datafld22       =   "RrFixMaxTol"
            datatype22      =   2
            mask22          =   "#####0.000"
            maxlength22     =   10
            tooltiptext22   =   "Enter Fixed Maximum Tolerance"
            colname23       =   "RROPNMINTOL"
            heading23       =   "Opn -ve Tol "
            datafld23       =   "RrOpnMinTol"
            datatype23      =   2
            mask23          =   "#####0.000"
            maxlength23     =   10
            tooltiptext23   =   "Enter Minimum Open Tolerance"
            colname24       =   "RROPNMAXTOL"
            heading24       =   "Opn +ve Tol "
            datafld24       =   "RrOpnMaxTol"
            datatype24      =   2
            mask24          =   "#####0.000"
            maxlength24     =   10
            tooltiptext24   =   "Enter Maximum Open Tolerance"
            colname25       =   "RRPRTKEY"
            heading25       =   "Prtn "
            datafld25       =   "RrPrtKey"
            maxlength25     =   1
            style25         =   0
         End
      End
      Begin VB.Frame FraImpExc 
         Height          =   1515
         Left            =   2760
         TabIndex        =   88
         Top             =   7200
         Width           =   5355
         Begin MwfCtl.MWCTL_MED ATXT 
            Height          =   285
            Index           =   248
            Left            =   660
            TabIndex        =   89
            ToolTipText     =   "Excel File Name and Path"
            Top             =   555
            Width           =   3915
            _ExtentX        =   6906
            _ExtentY        =   503
            Enabled         =   0   'False
            MaxLength       =   250
            IdName          =   "WXLFILE"
         End
         Begin MwfCtl.MWCTL_MED ATXT 
            Height          =   285
            Index           =   249
            Left            =   750
            TabIndex        =   90
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
            TabIndex        =   91
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
            TabIndex        =   92
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
            Left            =   4680
            TabIndex        =   93
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
            Left            =   4680
            TabIndex        =   94
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
            TabIndex        =   99
            Top             =   120
            Width           =   5640
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
            TabIndex        =   98
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
            TabIndex        =   97
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
            TabIndex        =   96
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
            TabIndex        =   95
            Top             =   1065
            Width           =   645
         End
      End
      Begin VB.Frame FrmRefRates 
         Height          =   1140
         Left            =   4890
         TabIndex        =   80
         Top             =   7440
         Width           =   4860
         Begin MwfCtl.MWCTL_BTN1 CmdILCGo 
            Height          =   465
            Left            =   4050
            TabIndex        =   82
            Top             =   600
            Width           =   690
            _ExtentX        =   1217
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
            Caption         =   "Refresh Rates will update rates for the selected scope, are you sure? Also ensure the rates are copied to History as required."
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   1050
            Index           =   2
            Left            =   30
            TabIndex        =   81
            Top             =   90
            Width           =   4830
         End
      End
   End
   Begin MwfCtl.MWCTL_BTN1 CmdDelRt 
      Height          =   465
      Left            =   9180
      TabIndex        =   101
      ToolTipText     =   "Option to Delete Rates"
      Top             =   9360
      Width           =   945
      _ExtentX        =   1667
      _ExtentY        =   820
      BackColor       =   -2147483633
      MaskColor       =   8454143
      Caption         =   "&Delete Rates"
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
Attribute VB_Name = "EmrFrmRmRt"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'
'----- General comments on the Entry  ----
' X-'', A-RmRt, GrdRmRt
' The grid properties are set by the TcStructure Table.
'
' FraCpy, CmdCpy, CmdCpyGo, CpyRmRt, wCpyCustCtgFr, wCpyCustFr, wCpyCustTo, wCpySCtgFr,...wCpyRmCdTo
'          all part of Copy routine.
'
' private Sub GrdPropset, called from AdcRecsource, sets all properties for each grid col.
'     reads the TcStru table, assumes records for all fields will be avlbl for Tc/Ctg Level
'
' private Sub LnFrToProp set, called from Rowwhen & recalc on change of SCtg
'   sets the Style, heading of LrFr, LnTo basrm ed on the SubCtg.
'
' private Sub CmpString is used to set the values for Compare String property,
'            passing 'K'ey Fields, 'C'opy
'
' private fn ColSty, DataTyp, to set the Col Style, & Datatype property.
'
'-----
Option Explicit
Dim moCn As MwfLib.MDOConnection   '*** (Bef speed) bef 24/11/06
Dim ms_RrTcTyp As String, ms_RrCtg As String, ms_RrCmCtg As String, ms_RrSrt As String
'****** Sachin 2.13.0 - 05-06-2006 ******
Dim ms_QWModYN As Boolean
'****** Sachin 2.13.0 - 05-06-2006 ******
Dim ms_CstRtEdtYN As Boolean    '*********Bhavna added cst rt edt mod
Dim ms_TolWtPctDC As String, ms_TolQtyPctDC As String         ' ***** Sachin 3.11.0 - Tolerance Pct for 'D' and 'C'

Dim ms_RhYyyyMm As String, ms_MenuCd As String  '******** ZUBIN - 01st Dec 2003, EMR206 ********'
'*** For the Toggle Options of Command Buttons
Dim ms_PwdFlag As Boolean        '*** Jay 3.01 *** [ILC]
Enum en_RmRtFra
  Cpy = 0
  CpyHist = 1  '*************************** ZUBIN **************************
  Pwd = 2      '*** Jay 3.01 *** [ILC]
  ILC = 3      '*** Jay 3.01 *** [ILC]
  ImpExc = 4    ' ***** Bhavna - Import from Excel
  DelRt = 5 'Uni.24- Delete rates
End Enum
'*** Arr_Cv stores the causes validation state, Arr_Ed the Ena/Disa state of Cmd Buttons
Dim mArr_Ed As Variant
Private Sub CpyRmRtHist()
  '*************************** ZUBIN **************************
  ' 02nd Dec 2003, EMR206
  '*** Copy To History Option will copy the Rm rates from a range of customers
      'for the given TcTyp = 'RM', RmCtg and CmCtg
  '*** pass 'C' as the parameter for the CmpString procedure. This will set the values for the
      'CmpStr option of the key frame fields as blank and set proper values for the CmpStr
      'option of the fields in the copy frame
  '*** Cannot copy if Record exists in RmRtHist for the given TcTyp = 'RM', RmCtg, CmCtg and YyyyMm
  
  Dim ws_RmCnd As String
  Dim wRs_CmRmRtHist As MwfLib.MDORowSet, wi_Row As Integer
  Dim ws_SQL As String
  'sv.23 C changes as K to get compare string of header for RmRtHist
  Call CmpString("K")     ' to set the Compare String Property of the ADC controls when the Copy Option Is clicked
  
  '*** Checking if copy is possible
  '''''  If Not moCn.RecSeek("Select CmCd From CustMst where CmCtg= '" + ADC("wRrCmCtg") + "' " + _
  '''''     "and CmCd = '" + ADC("wCpyHistCmCdFr") + "'") Then _
  '''''     DispMsg "Enter Proper Source Customer Code Fr", etError: ADC("wCpyHistCmCdFr").SetFocus: Exit Sub
  '''''  If Not moCn.RecSeek("Select CmCd From CustMst where CmCtg= '" + ADC("wRrCmCtg") + "' " + _
  '''''     "and CmCd = '" + ADC("wCpyHistCmCdTo") + "'") Then _
  '''''     DispMsg "Enter Proper Source Customer Code To", etError: ADC("wCpyHistCmCdTo").SetFocus: Exit Sub

  ws_RmCnd = ADC.RepCond
  ws_RmCnd = IIF(ws_RmCnd <> vbNullString, " where " + ws_RmCnd, "")

  If moCn.RecSeek("Select 'x' from RmRtHist " + ws_RmCnd) Then
    DispMsg "Cannot Copy as RM Rate History Exists For Given Customer(s) For Given YyyyMm ", etError
    Exit Sub
  End If
  'If UCase(ms_MenuCd) = UCase("RmRt") Then
   ' Dim wErrMsg As String
    'If Not ChkCmpUsrAccess(ADC("WRRCMCTG"), ADC.MenuCd, "", "", ADC("WCPYHISTCMCDTO"), wErrMsg) Then _
     ' DispMsg "You don't have permissions to add this Customer", etError: Exit Sub
  'End If

  '*** Beginning the copy process
  GrdRm.StartCopy
  MWLib.BeginProcess Me, "Copying Raw Material Rate History ..."

  'P.N.: var ws_RmCnd is not used below because CmpString() Procedure considers
    'Parameter and MenuCd only to set 'CmpStr' property.
    'Incase of MenuCd = 'RmRtHist' we need to check records in RmRt also which is not
    'handled in CmpString()
  ' Zubin 212 (Before 18/12/05)
  'ws_SQL = "Select * from RmRt where RrTcTyp='" + ADC("wRrTcTyp") + "' and " + _
             "RrCtg='" + ADC("wRrCtg") + "' and RrCmCtg= '" + ADC("wRrCmCtg") + "' and " + _
             "RrCmCd>='" + ADC("wCpyHistCmCdFr") + "' and RrCmCd<='" + _
             ADC("wCpyHistCmCdTo") + "' order by RrCmCtg, RrCmCd"
  ' Zubin 212 (18/12/05)
  'ws_SQL = "Select * from RmRt where RrPrtKey='" + ctCurrPrtn + "' and RrTcTyp='" + ADC("wRrTcTyp") + "' and " + _
             "RrCtg='" + ADC("wRrCtg") + "' and RrCmCtg= '" + ADC("wRrCmCtg") + "' and " + _
             "RrCmCd>='" + ADC("wCpyHistCmCdFr") + "' and RrCmCd<='" + _
             ADC("wCpyHistCmCdTo") + "' and " + _
             "((Select PNum from Param where PTyp= 'USR' and PMCd= '" + gs_UsrCd + "' and PSCd= '')= 1 OR " + _
             "RrCmCd NOT IN ('" + ctSelfCmCd + "', '" + ctValueCmCd + "') ) order by RrCmCtg, RrCmCd"
  'sv.23 to set compare string for headers and ws_SQL query changed
  Call CmpString("C")
  ws_RmCnd = ADC.RepCond
  
  ws_SQL = "Select * from RmRt where RrPrtKey='" + ctCurrPrtn + "' and " + ws_RmCnd + _
             " and ((Select PNum from Param where PTyp= 'USR' and PMCd= '" + gs_UsrCd + "' and PSCd= '')= 1 OR " + _
             "RrCmCd NOT IN ('" + ctSelfCmCd + "', '" + ctValueCmCd + "') ) order by RrCmCtg, RrCmCd"

  Set wRs_CmRmRtHist = moCn.OpenRes(ws_SQL)

  With wRs_CmRmRtHist
    Do While Not (.EOF Or .BOF)
      MWLib.RunningProcess "Copying RM Rate To History For Customer " + !RrCmCd
      GrdRm.AddItem
      wi_Row = GrdRm.Rows - 1

      GrdRm.Value(wi_Row, "RrPrtKey") = ctCurrPrtn    '****** Sachin 3.02.0
      GrdRm.Value(wi_Row, "RrCmCd") = !RrCmCd
      GrdRm.Value(wi_Row, "RrSCtg") = !RrSCtg
      GrdRm.Value(wi_Row, "RrSTWGrp") = !RrSTWGrp
      GrdRm.Value(wi_Row, "RrCd") = !RrCd
      GrdRm.Value(wi_Row, "RrDmCtg") = !RrDmCtg
      GrdRm.Value(wi_Row, "RrLsCtg") = !RrLsCtg
      GrdRm.Value(wi_Row, "RrLabMCd") = !RrLabMCd
      GrdRm.Value(wi_Row, "RrRngSz") = !RrRngSz
      GrdRm.Value(wi_Row, "RrFrLn") = !RrFrLn
      GrdRm.Value(wi_Row, "RrToLn") = !RrToLn
      GrdRm.Value(wi_Row, "RrCmCurCd") = !RrCmCurCd
      GrdRm.Value(wi_Row, "RrCstRt") = !RrCstRt
      GrdRm.Value(wi_Row, "RrSalRt") = !RrSalRt
      '******************Geeta*****************
      '18th Jan 2004, Emr206
      GrdRm.Value(wi_Row, "RrXWtPer") = !RrXWtPer
      '****************************************
      GrdRm.SaveRec (wi_Row)
      .MoveNext
    Loop
  End With
  MWLib.EndProcess Me
  GrdRm.EndCopy
  Set wRs_CmRmRtHist = Nothing
  '*** End of the copy process

End Sub

Private Sub GrdTblField(ByVal pv_TblName As String)
  '*************************** ZUBIN **************************
  ' 27th Nov 2003, EMR206
  ' Procedure to assign Field Names Dynamically to GrdRm
  '*************************** ZUBIN **************************
  
  With GrdRm
    If UCase(pv_TblName) = UCase("RmRt") Then
      .ColProp("RrCmCd").DataFld = "RrCmCd"
      .ColProp("RrCmCurCd").DataFld = "RrCmCurCd"
      .ColProp("RrSCtg").DataFld = "RrSCtg"
      .ColProp("RrSTWGrp").DataFld = "RrSTWGrp"
      .ColProp("RrCd").DataFld = "RrCd"
      .ColProp("RrDmCtg").DataFld = "RrDmCtg"
      .ColProp("RrLabMCd").DataFld = "RrLabMCd"
      .ColProp("RrFrLn").DataFld = "RrFrLn"
      .ColProp("RrToLn").DataFld = "RrToLn"
      .ColProp("RrRngSz").DataFld = "RrRngSz"
      .ColProp("RrCstRt").DataFld = "RrCstRt"
      .ColProp("RrSalRt").DataFld = "RrSalRt"
      .ColProp("RrLsCtg").DataFld = "RrLsCtg"
      '*****************Geeta**************************
      '18th Jan 2004,Emr206
      .ColProp("RRXWTPER").DataFld = "RRXWtPer"
      '*************************************************
      ' **** Zubin 211 **** '
      .ColProp("RrFixMinTol").DataFld = "RrFixMinTol"
      .ColProp("RrFixMaxTol").DataFld = "RrFixMaxTol"
      .ColProp("RrOpnMinTol").DataFld = "RrOpnMinTol"
      .ColProp("RrOpnMaxTol").DataFld = "RrOpnMaxTol"
      .ColProp("RrPrtKey").DataFld = "RrPrtKey"     '****** Sachin 3.02.0
      ' **** Zubin 211 **** '
      ' Zubin 212 (Tell Sachin)
      .ColProp("RrDmCd").DataFld = "RrDmCd"
      .ColProp("RrRmKt").DataFld = "RrRmKt"
      ' Zubin 212 (Tell Sachin)
      .ColProp("RrDmSfx").DataFld = "RrDmSfx"
      .ColProp("RrDmSz").DataFld = "RrDmSz"
    ElseIf UCase(pv_TblName) = UCase("RmRtHist") Then
      .ColProp("RrCmCd").DataFld = "RhCmCd"
      .ColProp("RrCmCurCd").DataFld = "RhCmCurCd"
      .ColProp("RrSCtg").DataFld = "RhSCtg"
      .ColProp("RrSTWGrp").DataFld = "RhSTWGrp"
      .ColProp("RrCd").DataFld = "RhCd"
      .ColProp("RrDmCtg").DataFld = "RhDmCtg"
      .ColProp("RrLabMCd").DataFld = "RhLabMCd"
      .ColProp("RrFrLn").DataFld = "RhFrLn"
      .ColProp("RrToLn").DataFld = "RhToLn"
      .ColProp("RrRngSz").DataFld = "RhRngSz"
      .ColProp("RrCstRt").DataFld = "RhCstRt"
      .ColProp("RrSalRt").DataFld = "RhSalRt"
      .ColProp("RrLsCtg").DataFld = "RhLsCtg"
      '*****************Geeta**************************
      '18th Jan 2004,Emr206
      .ColProp("RRXWTPER").DataFld = "RhXWtPer"
      '*************************************************
      ' **** Zubin 211 **** '
      .ColProp("RrFixMinTol").DataFld = "RhFixMinTol"
      .ColProp("RrFixMaxTol").DataFld = "RhFixMaxTol"
      .ColProp("RrOpnMinTol").DataFld = "RhOpnMinTol"
      .ColProp("RrOpnMaxTol").DataFld = "RhOpnMaxTol"
      ' **** Zubin 211 **** '
      ' Zubin 212 (Tell Sachin)
      .ColProp("RrDmCd").DataFld = ""
      .ColProp("RrRmKt").DataFld = ""
      .ColProp("RrPrtKey").DataFld = "RhPrtKey"     '****** Sachin 3.02.0
      ' Zubin 212 (Tell Sachin)
      'Uni.186 - Item cannot be found error in RmRt hist. RrDmSfx & RrDmSz initialized
      .ColProp("RrDmSfx").DataFld = ""
      .ColProp("RrDmSz").DataFld = ""
    End If
  End With
   
End Sub

Private Sub ADC_Load()
  '*** Set the First Key Control property to the working field wRrTcTyp
  '*** Set the First Non Key Control property to the grid GrdRm
  '*** Set the Child Property Of ADC to GrdRm
  '*** Set the Previous and Next Control Properties for the Grid GrdRm as ADC
  '*** Set the Hot Key for the Grid GrdRm as 'A'
  '*** initialise the default values for the working fields wRrTcTyp, wRrCtg and wRrSrt

  Set moCn = ADC.Connection    '*** (Bef speed) bef 24/11/06
 
  ReDim mArr_Ed(5)
 
  '*************************** ZUBIN **************************
  ' 27th Nov 2003, EMR206
  If UCase(ADC.MenuCd) = UCase("RmRt") Then
    ms_MenuCd = "RmRt"
    Set ADC.FirKeyCtl = ADC("wRrTcTyp")
  ElseIf UCase(ADC.MenuCd) = UCase("RmRtHist") Then
    ms_MenuCd = "RmRtHist"
    Set ADC.FirKeyCtl = ADC("wRhYyyyMm")
  '*** Jay 3.1.2 *** [ILC]
  ElseIf UCase(ADC.MenuCd) = UCase("ILC") Then
    ms_MenuCd = "RmRt"
    Set ADC.FirKeyCtl = ADC("wRrCtg")
  '*** Jay 3.1.2 *** [ILC]
      
  End If
  '*************************** ZUBIN **************************
    
  Set ADC.FirNKeyCtl = GrdRm
  ADC.Child = "GrdRm"

  Set GrdRm.PrevCtl = ADC
  Set GrdRm.NextCtl = ADC
  GrdRm.HotKey = "A"

  ms_RrTcTyp = ""
  ms_RrCtg = ""
  ms_RrCmCtg = ""
  ms_RrSrt = ""
  '*************************** ZUBIN **************************
  ms_RhYyyyMm = ""
  '*************************** ZUBIN **************************
  
  '****** Sachin 2.13.0 - 05-06-2006 ******
  ms_QWModYN = IIF(moCn.GetFldVal("Select hQWModYn From Head Where HCd='" + ctSelfCmCd + "' And HCoCd='" + ctSelfCoCd + "'") = "Y", True, False)
  '****** Sachin 2.13.0 - 05-06-2006 ******
  '***Bhavna added sct rt edt yn
  ms_CstRtEdtYN = IIF(moCn.GetFldVal("Select hCstRtEdtYn From Head Where HCd='" + ctSelfCmCd + "' And HCoCd='" + ctSelfCoCd + "'") = "Y", True, False)
  
  '***** Sachin 3.11.0 - Tolerance (WTT & QTT) Percent for 'D' and 'C'
  Dim wRsTol As MDORowSet
  Set wRsTol = moCn.OpenRes("Select hTolWtPctDC, hTolQtyPctDC From Head Where HCd='" + ctSelfCmCd + "' And HCoCd='" + gs_CoCd + "'")
  ms_TolWtPctDC = IIF(Not (wRsTol.BOF Or wRsTol.EOF), wRsTol!hTolWtPctDC, "")
  ms_TolQtyPctDC = IIF(Not (wRsTol.BOF Or wRsTol.EOF), wRsTol!hTolQtyPctDC, "")
  Set wRsTol = Nothing
  
  ms_PwdFlag = moCn.RecSeek("Select PValue1 from Param where PTyp= 'SEOPWD' and PmCd= 'SEOPWD' and PValue1= ''")  '*** Jay 3.01 *** [ILC]
  
  '****** Sachin 3.02.0
  LblPrtnKey.Visible = IIF(gs_Partition = ctCurrPrtn, False, True)
  ADC("WRRPRTKEY").Visible = IIF(gs_Partition = ctCurrPrtn, False, True)
  GrdRm.ColProp("RrPrtKey").Style = IIF(gs_Partition = ctCurrPrtn, fgcinvisible, fgcReadOnly)
  
End Sub
Private Sub ADC_KeyWhen()
  '*** pass 'K' as the parameter for the CmpString procedure. This sets the CmpStr value of
      'the fields in the copy frame as blank and sets the proper cmpstr values for fileds in
      'the key frame (i.e. the range selection fields)
  '*** set the default values for RrTcTyp, RRCtg and Sort Option from the previous entry
  '*** Disable All command buttons
  '*** Hide all frames except Frame FraNKeyAll
  '*************************** ZUBIN **************************
  ' 27th Nov 2003, EMR206
  ' Set Form Caption and enable/disable or hide/show appropriate controls
  If UCase(ms_MenuCd) = UCase("RmRt") Then
    LblYyyyMm.Visible = False
    '*** Jay 3.1.2 ***
    If UCase(ADC.MenuCd) = UCase("ILC") Then
      ADC("wRrTcTyp") = "ILC"
      ADC("wRrTcTyp").Enabled = False
      Me.Caption = "Increment Logic Chart"
    Else
      Me.Caption = "Raw Material Rate Master"
    End If
    '*** Jay 3.1.2 ***
        '===== HELP FOR TcTyp fails if below statement is executed ====='
    ADC("wRhYyyyMm").Enabled = False
    ADC("wRhYyyyMm").Visible = False
    ADC("wRhYyyyMm").TabStop = False
    CmdCpyHist.Visible = False
  ElseIf UCase(ms_MenuCd) = UCase("RmRtHist") Then
    LblTcTyp.Visible = False
    Me.Caption = "Raw Material Rate Master History (TcTyp 'RM')"
    ADC("wRrTcTyp") = "RM"
    ADC("wRrTcTyp").Enabled = False
    ADC("wRrTcTyp").Visible = False
    CmdCpy.Visible = False
    CmdILC.Visible = False    '*** Jay 3.01 *** [ILC]
    CmdPwd.Visible = False      '*** Jay 3.01 *** [ILC]
    CmdDelRt.Visible = False  'sv.23 DelRt should not visible for RmRate history
  End If
  '*************************** ZUBIN **************************
  
  Call CmpString("K")     ' to set the Compare String Property of the ADC controls when the Focus is in the Key Frame
  If ms_RrTcTyp <> "" Then ADC("wRrTcTyp") = ms_RrTcTyp
  If ms_RrCtg <> "" Then ADC("wRrCtg") = ms_RrCtg
  If ms_RrCmCtg <> "" Then ADC("wRrCmCtg") = ms_RrCmCtg
  If ms_RrSrt <> "" Then ADC("wRrSrt") = ms_RrSrt

  
  '*************************** ZUBIN **************************
  ' 01st Dec 2003, EMR206
  If UCase(ms_MenuCd) = UCase("RmRtHist") And ms_RhYyyyMm <> "" Then
    ADC("wRhYyyyMm") = ms_RhYyyyMm
  End If
  '*************************** ZUBIN **************************

  'PREVIOUS - EMR204 <<<<<<====================================
  'Call EnaDisaCmds(True)
  'PREVIOUS - EMR204 <<<<<<====================================
  
  '*************************** ZUBIN **************************
  ' 27th Nov 2003, EMR206
  Call EnaDisaCmds(True)    '*** (Jen 3.01)
  
  'If UCase(ms_MenuCd) = UCase("RmRt") Then
  '  Call zEnaDisaCmds(True)
  'ElseIf UCase(ms_MenuCd) = UCase("RmRtHist") Then
  '  Call zEnaDisaCmds(True)
  'End If
  '*************************** ZUBIN **************************
  TxtPwd.text = ""      '*** Jay 3.01 *** [ILC]
  Call HideAllFras
End Sub
Private Sub ADC_KeyValid(Cancel As Boolean, ErrMsg As String)
  '*** Check whether the RrTcTyp, RrCtg and the Sort Option is valid
      '(Sort option can be blank in which case the default sorting is on RrCmCtg, RrCmCd, RrSCtg, RrCd,
      ', RrToLn)
 
'*** Checking Key values in Key Valid as they are copied in Init
  Call ATXT_FldChange(0, UCase("wRrTcTyp"), Cancel, ErrMsg, ADC("wRrTcTyp"))
  If Cancel = True Then Exit Sub
  
  Call ATXT_FldChange(0, UCase("wRrCtg"), Cancel, ErrMsg, ADC("wRrCtg"))
  If Cancel = True Then Exit Sub
  
  Call ATXT_FldChange(0, UCase("wRrCmCtg"), Cancel, ErrMsg, ADC("wRrCmCtg"))
  If Cancel = True Then Exit Sub
  '*************************** ZUBIN **************************
  ' 01st Dec 2003, EMR206
  If UCase(ms_MenuCd) = UCase("RmRtHist") Then
    Call ATXT_FldChange(0, UCase("wRhYyyyMm"), Cancel, ErrMsg, ADC("wRhYyyyMm"))
    If Cancel = True Then Exit Sub
  End If
  '*************************** ZUBIN **************************
  
'*** for 'INL' (Increment on Labour) RrCtg should be blank.
' Zubin 212 ('MUP' added)
  '****** Sachin 3.3.0 (24-06-08) - [14. Cost Markup In Invoicing] ******
  '******Bhavna added chain CHN
  'pg.5- RmCtg should be blank for MRP,MRD
  If (ADC("wRrTcTyp") = "INL" Or ADC("wRrTcTyp") = "MUP" Or ADC("wRrTcTyp") = "FRT" Or ADC("wRrTcTyp") = "IML" Or ADC("wRrTcTyp") = "CHN" Or ADC("wRrTcTyp") = "MRP" Or ADC("wRrTcTyp") = "MRD") And ADC("wRrCtg") <> "" Then _
     Cancel = True: ErrMsg = "Raw Material Category Should Be Blank": Exit Sub
    
     Cancel = (ADC("wRrSrt") <> "") And _
          (Not moCn.RecSeek("Select PMCd From Param Where " + _
                       "PTyp = 'SRTRR' And PMCd = '" + ADC("wRrSrt") + "'"))
  If Cancel = True Then ErrMsg = "Invalid Sort Option": Exit Sub

'Indigo.1- Checking user access permission for the given customers
  If UCase(ms_MenuCd) = UCase("RmRt") Then
    Dim wErrMsg As String
    If Not ChkCmpUsrAccess(ADC("WRRCMCTG"), ADC.MenuCd, ADC("WRRCMCDFR"), ADC("WRRCMCDTO"), "", wErrMsg) Then
      Cancel = True: ErrMsg = "You Don't Have Permissions To View The Details of some Customers in given Range, example:" + wErrMsg: Exit Sub
    End If
  End If
    
End Sub

Private Sub ADC_NonKeyWhen(pr_Cancel As Boolean, pr_ErrMsg As String)
  '*** Store the RrTcTyp, RrCtg, Sort Option for the next entry
  '*** Enable all the command buttons
  
  ms_RrTcTyp = ADC("wRrTcTyp")
  ms_RrCtg = ADC("wRrCtg")
  ms_RrCmCtg = ADC("wRrCmCtg")
  ms_RrSrt = ADC("wRrSrt")
  
  '*************************** ZUBIN **************************
  ' 01st Dec 2003, EMR206
   ms_RhYyyyMm = ADC("wRhYyyyMm")
  '*************************** ZUBIN **************************
  
  'PREVIOUS - EMR204 <<<<<<====================================
  'Call EnaDisaCmds(False)
  'PREVIOUS - EMR204 <<<<<<====================================
    
  '*************************** ZUBIN **************************
  ' 28th Nov 2003, EMR206 - Enable/Disable cmd buttons as per MenuCd
  Call EnaDisaCmds(False)   '*** (Jen 3.01)
  
  'If UCase(ms_MenuCd) = UCase("RmRt") Then
  '  Call zEnaDisaCmds(False)
  'ElseIf UCase(ms_MenuCd) = UCase("RmRtHist") Then
  '  Call zEnaDisaCmds(False)
  'End If
  
  '****** Sachin 2.13.0 - 05-06-2006 ******
    If ms_RrTcTyp = "RM" Then
     If ms_QWModYN = False Then
        If ADC("wRrCtg") <> "X" Then
          GrdRm.ColProp("WRRQW").Style = fgcinvisible: GrdRm.ColProp("WRRQWCST").Style = fgcinvisible
        Else
          GrdRm.ColProp("WRRQW").Style = fgcReadOnly: GrdRm.ColProp("WRRQWCST").Style = fgcinvisible
        End If
      Else
          GrdRm.ColProp("WRRQW").Style = fgcReadOnly: GrdRm.ColProp("WRRQWCST").Style = fgcReadOnly
      End If
      CmdILC.Enabled = True     '*** Jay 3.01 *** [ILC]
      CmdPwd.Enabled = True       '*** Jay 3.01 *** [ILC]
      If ms_RrCtg = "D" Or ms_RrCtg = "C" Then CmdImpExc.Enabled = True
      CmdDelRt.Enabled = True 'Uni.24-Delete rates button will display only for RM Tctyp
    Else
      GrdRm.ColProp("WRRQW").Style = fgcinvisible: GrdRm.ColProp("WRRQWCST").Style = fgcinvisible
      CmdILC.Enabled = False    '*** Jay 3.01 *** [ILC]
      CmdPwd.Enabled = False      '*** Jay 3.01 *** [ILC]
      CmdImpExc.Enabled = False
      CmdDelRt.Enabled = False 'Uni.24-Delete rates button will display only for RM Tctyp
    End If
    '****** Sachin 2.13.0 - 05-06-2006 ******

  '*************************** ZUBIN **************************
  'mw.136 showing tooltip for weight fields
  If ADC("wRrTcTyp") = "STW" Then
    GrdRm.ColProp("RrFixMinTol").ToolTipText = "Enter Minimum Weight"
    GrdRm.ColProp("RrFixMaxTol").ToolTipText = "Enter Maximum Weight"
    GrdRm.ColProp("RrSalRt").ToolTipText = "Enter Weight"
  Else
    GrdRm.ColProp("RrFixMinTol").ToolTipText = "Enter Fixed Minimum Tolerance"
    GrdRm.ColProp("RrFixMaxTol").ToolTipText = "Enter Fixed Maximum Tolerance"
    GrdRm.ColProp("RrSalRt").ToolTipText = "Enter Sales Rate"
  End If
End Sub


Private Sub ATXT_FldWhen(Index As Integer, ByVal IdName As String, Cancel As Boolean, ErrMsg As String)
  '*** Set Helps For Each Field
  '*** For the help of working field wRrCdFr ('RmCode From' in the selection range option)
      'give the first parameter as 'N' (i.e. select RmCodes which are not custom codes) and
      'the second parameter as the RmCtg
  '*** Similarly for the help of working field wRrCdTo ('RmCode To' in the selection range option)
  '*** Similarly for the helps of the working fields in the copy option frame wCpyRmCdFr and
      'wCpyRmCdTo
  Select Case IdName
  Case Is = UCase("wRrTcTyp")
    Call HlpList.PMCd("RRTCTYP")
  Case Is = UCase("wRrCtg")
    Call HlpList.PMCd("RMCTG")
  Case Is = UCase("wRrCmCtg")
    Call HlpList.PMCd("CMCTG")
  '*** Jay 2.14 ***
  Case Is = UCase("wRrSCtgFr")
    Call HlpList.PSCd("RMSCTG", ADC("wRrCtg"))
  Case Is = UCase("wRrSCtgTo")
    Call HlpList.PSCd("RMSCTG", ADC("wRrCtg"))
  Case Is = UCase("wRrCdFr")
    Call HlpList.RmCd("N", "'" + ADC("wRrCtg") + "'", , False)
  Case Is = UCase("wRrCdTo")
    Call HlpList.RmCd("N", "'" + ADC("wRrCtg") + "'", , False)
  Case Is = UCase("wRrCmCdFr")
    Call HlpList.CustCd(ADC("wRrCmCtg"), False)
  Case Is = UCase("wRrCmCdTo")
    Call HlpList.CustCd(ADC("wRrCmCtg"), False)
  '*** Jay 2.14 ***
  '**********Bhavna 5.1.0
  Case Is = UCase("wRrDmCdFr")
    Call HlpList.DmCd("DM")
  Case Is = UCase("wRrDmCdTo")
    Call HlpList.DmCd("DM")
  '**********Bhavna 5.1.0
  'PREVIOUS - EMR204 <<<<<<====================================
  'Case Is = UCase("wRrSrt")
  '  Call HlpList.PMCd("SRTRR")
  '*************************** ZUBIN **************************
  Case Is = UCase("wRrSrt")
    If UCase(ms_MenuCd) = UCase("RmRt") Then
      Call HlpList.PMCd("SRTRR")
    ElseIf UCase(ms_MenuCd) = UCase("RmRtHist") Then
      Call HlpList.PMCd("SRTRH")
    End If
  '*************************** ZUBIN **************************
  
' fields of Copy frame.
  Case Is = UCase("wCpyCustCtgFr")
    Call HlpList.PMCd("CMCTG")
'*** jay 2.14 ***
  Case Is = UCase("wCpyCustFr")
    Call HlpList.CustCd(ADC("wCpyCustCtgFr"), False)
  Case Is = UCase("wCpyCustTo")
    Call HlpList.CustCd(ADC("wRrCmCtg"), False)
  Case Is = UCase("wCpySCtgFr")
    Call HlpList.PSCd("RMSCTG", ADC("wRrCtg"), False)
  Case Is = UCase("wCpySCtgTo")
    Call HlpList.PSCd("RMSCTG", ADC("wRrCtg"), False)
  Case Is = UCase("wCpyRmCdFr")
    Call HlpList.RmCd("N", "'" + ADC("wRrCtg") + "'", , False)
  Case Is = UCase("wCpyRmCdTo")
    Call HlpList.RmCd("N", "'" + ADC("wRrCtg") + "'", , False)
  Case Is = UCase("wCpyCustNameFr")
    Call HlpList.CmName(ADC("wCpyCustCtgFr"), False)
  Case Is = UCase("wCpyCustNameTo")
    Call HlpList.CmName(ADC("wRrCmCtg"), False)
 '*************************** ZUBIN **************************
  ' 27th Nov 2003, EMR206
  ' fields of Copy History frame.
  Case Is = UCase("wRhYyyyMm")
    gb_ShowYyMmHlp = True
    Call HlpList.YyyyMm
    gb_ShowYyMmHlp = False
  
  '*************************** ZUBIN **************************
  '*** Jay 2.14 ***
  ' Zubin 213
  Case Is = UCase("wCpyCurCdFr")
    If UCase(ADC("wCpyCustFr")) = ctSelfCmCd Or UCase(ADC("wCpyCustFr")) = ctValueCmCd Then
      Call HlpList.PMCd("CURNCY")
    Else
      Cancel = True
    End If
  Case Is = UCase("wCpyCurCdTo")
    If UCase(ADC("wCpyCustTo")) = UCase(ctSelfCmCd) Or UCase(ADC("wCpyCustTo")) = UCase(ctValueCmCd) Then
      Call HlpList.PMCd("CURNCY")
    Else
      Cancel = True
    End If
  Case Is = UCase("wRrCmCurCdFr"), UCase("wRrCmCurCdTo")
      Call HlpList.PMCd("CURNCY")
  ' Zubin 213
  'IG.72-Calling Help for scope of From & To STW/BSW
  Case Is = UCase("wRrSTWGrpFr"), UCase("wRrSTWGrpTo")
      Call HlpList.PMCd("STWGRP")
 End Select
End Sub

Private Sub atxt_ReCalculate(Index As Integer, ByVal IdName As String, ByVal OldVal As Variant, ByVal RowNum As Integer, ByVal SourceName As String, ByVal SourceColName As String)
  Dim wSrcCmCd As String, wTgtCmCd As String, wSrcCustName As String, wTgtCustName As String
  Select Case UCase(IdName)
  Case Is = UCase("wCpyCustFr")
    If ADC.Mode = xNorm Then
      wSrcCmCd = moCn.GetFldVal("Select CmCd from CustMst where CmCtg= '" + ADC("wCpyCustCtgFr") + "' and " + _
                                "CmName= '" + ADC("wCpyCustNameFr") + "'")
      If wSrcCmCd <> ADC("wCpyCustFr") Then ADC("wCpyCustFr") = wSrcCmCd
    End If
  Case Is = UCase("wCpyCustNameFr")
    If ADC.Mode = xNorm Then
       wSrcCustName = moCn.GetFldVal("Select CmName from CustMst where CmCtg='" + ADC("wCpyCustCtgFr") + "' and " + _
                                     "CmCd='" + ADC("wCpyCustFr") + "'")
       If wSrcCustName <> ADC("wCpyCustNameFr") Then ADC("wCpyCustNameFr") = wSrcCustName
    ElseIf ADC.Mode = xInit Then
       ADC("wCpyCustNameFr") = moCn.GetFldVal("Select CmName from CustMst where CmCtg= '" + ADC("wCpyCustCtgFr") + "' and " + _
                                  "CmCd='" + ADC("wCpyCustFr") + "'")
    End If
  Case Is = UCase("wCpyCustTo")
    If ADC.Mode = xNorm Then
      wTgtCmCd = moCn.GetFldVal("Select CmCd from CustMst where CmCtg= '" + ADC("wRrCmCtg") + "' and " + _
                                "CmName= '" + ADC("wCpyCustNameTo") + "'")
      If wTgtCmCd <> ADC("wCpyCustTo") Then ADC("wCpyCustTo") = wTgtCmCd
    End If
  Case Is = UCase("wCpyCustNameTo")
    If ADC.Mode = xNorm Then
       wTgtCustName = moCn.GetFldVal("Select CmName from CustMst where CmCtg='" + ADC("wRrCmCtg") + "' and " + _
                                     "CmCd='" + ADC("wCpyCustTo") + "'")
       If wTgtCustName <> ADC("wCpyCustNameTo") Then ADC("wCpyCustNameTo") = wTgtCustName
    ElseIf ADC.Mode = xInit Then
       ADC("wCpyCustNameTo") = moCn.GetFldVal("Select CmName from CustMst where CmCtg='" + ADC("wRrCmCtg") + "' and " + _
                                  "CmCd='" + ADC("wCpyCustTo") + "'")
    End If
  ' Zubin 213
  Case Is = UCase("wCpyCurCdFr")
    If ADC.Mode = xNorm Then
      ADC("wCpyCurCdFr") = moCn.GetFldVal("Select CmCurCd from CustMst where CmCtg='" + ADC("wCpyCustCtgFr") + "' and " + _
                                   "CmCd='" + ADC("wCpyCustFr") + "'")
    End If
  Case Is = UCase("wCpyCurCdTo")
    If ADC.Mode = xNorm Then
      ADC("wCpyCurCdTo") = moCn.GetFldVal("Select CmCurCd from CustMst where CmCtg='" + ADC("wRrCmCtg") + "' and " + _
                                   "CmCd='" + ADC("wCpyCustTo") + "'")
    End If
  Case Is = UCase("wCnvRtFr")
    If ADC.Mode = xNorm Then
      ADC("wCnvRtFr") = CnvRt(ADC("wCpyCurCdFr"), ADC("wCpyCurCdTo"))
    End If
  Case Is = UCase("wMulDivFr")
    If ADC.Mode = xNorm Then ADC("wMulDivFr") = MulDiv(ADC("wCpyCurCdFr"), ADC("wCpyCurCdTo"))
  ' Zubin 213
  End Select
End Sub
Private Sub ATXT_FldChange(Index As Integer, ByVal IdName As String, Cancel As Boolean, ErrMsg As String, ByVal pv_NewValue As Variant)
  '*** wRrTcTyp     valid PMCd('RRTCTYP')
  '*** wRrCtg       If ADC(wRrTcTyp)<> 'INL' then valid PMCd('RMCTG')
  '*** wRrSrt       Blank or valid PMCd('SRTRR')
  '*** wRrCmCtg     valid Customer Category ('C' or 'P')
  '*** wCpyCustFr   Blank or valid CmCd from CustMst(CmCtg= adc(wCpyCustCtgFr))
  '*** wCpyCustTo   Blank or valid CmCd from CustMst(CmCtg= adc(wRrCmCtg))
  
  Select Case UCase(IdName)
  Case Is = UCase("wRrTcTyp")
    Cancel = Not moCn.RecSeek("Select PMCd From Param where PTyp= 'RRTCTYP' " + _
                         "and PMCd = '" + pv_NewValue + "'")
                         
    '*** Jay 3.1.2 *** [ILC]
    If Cancel = True Then ErrMsg = "Invalid Tc Type": Exit Sub
    If UCase(ADC.MenuCd) = UCase("RmRt") And pv_NewValue = "ILC" Then
      Cancel = True: ErrMsg = "ILC Can Not Be Selected In This Entry": Exit Sub
    End If
    '*** Jay 3.1.2 *** [ILC]
  Case Is = UCase("wRrCtg")
  
    'KC.4 - RM Category validation. INP applicable for D & C category only
    '6.1 new GPSLXM added
'    If UCase(ADC("wRrTcTyp")) = "INP" And UCase(pv_NewValue) <> "D" And UCase(pv_NewValue) <> "C" Then
'      Cancel = True: ErrMsg = "Invalid RM Category": Exit Sub
'    End If
  
    ' Zubin 212 ('MUP' added) ('FRT' '****Geeta***Crm**212)
    
    '****** Sachin 3.3.0 (24-06-08) - [14. Cost Markup In Invoicing] ******
    '*********bhavna added chain CHN
    'pg.5- RMCTg should be blank for MRP,MRD
    Cancel = (ADC("wRrTcTyp") <> "INL" And ADC("wRrTcTyp") <> "MRP" And ADC("wRrTcTyp") <> "MRD" And ADC("wRrTcTyp") <> "MUP" And ADC("wRrTcTyp") <> "FRT" And ADC("wRrTcTyp") <> "CHN") And ADC("wRrTcTyp") <> "IML" And ADC("wRrTcTyp") <> "WCT" And ADC("wRrTcTyp") <> "WXT" And _
                (Not moCn.RecSeek("Select PMCd From Param where PTyp= 'RMCTG' and PMCd = '" + pv_NewValue + "'"))
    If Cancel = True Then ErrMsg = "Invalid Raw Material Category": Exit Sub
    '****** Sachin 3.3.0 (24-06-08) - [14. Cost Markup In Invoicing] ******
    
    '*************************** Geeta**************************
    ' 29th Jan 2004, EMR206
'    If ADC("wRrTcTyp") = "QTT" Then
'      Cancel = UCase(pv_NewValue) <> "D" And UCase(pv_NewValue) <> "C"
'      ErrMsg = "Raw Material Categories 'D' or 'C' are only allowed.": Exit Sub
'    ElseIf ADC("wRrTcTyp") = "WTT" Then
'      Cancel = UCase(pv_NewValue) = "X"
'      ErrMsg = "Raw Material Categories 'X' is not allowed.": Exit Sub
'    End If
'    '****************************Geeta*************************
     
    '*************************** ZUBIN *************************
    ' 01st Dec 2003, EMR206
    '*** Jay 3.01 *** [ILC] [NAS]
    '****** Sachin 3.3.0 (24-06-08) - [14. Cost Markup In Invoicing] ******
    ' ***** Manali 3.6.0 - 07/09/09 - ADC("wRrTcTyp") = "IMK" (Cost Markup In Invoicing) Removed from next Check
    If UCase(ms_MenuCd) = UCase("RmRtHist") Or ADC("wRrTcTyp") = "ILC" Or ADC("wRrTcTyp") = "NAS" Then
      Cancel = UCase(pv_NewValue) <> "D" And UCase(pv_NewValue) <> "C"
      ErrMsg = "Raw Material Categories 'D' or 'C' are only allowed.": Exit Sub
    End If
    
    ' ***** Manali 3.6.0 - 07/09/09 - ADC("wRrTcTyp") = "IMK" (Cost Markup In Invoicing) , RrCtg='X' added below
    If ADC("wRrTcTyp") = "IMK" Or ADC("wRrTcTyp") = "IMC" Then
      If InStr(1, ",G,P,S,L,D,C,X,", "," + UCase(pv_NewValue) + ",", vbTextCompare) <= 0 Then Cancel = True   ' Sachin 3.11.0
      'Cancel = UCase(pv_NewValue) <> "D" And UCase(pv_NewValue) <> "C" And UCase(pv_NewValue) <> "X"
      ErrMsg = "Raw Material Categories 'G', 'P', 'S', 'L', 'D', 'C' or 'X' are only allowed.": Exit Sub
    End If
    '*** Jay 3.01 *** [ILC] [NAS]
    '******bhavna added chain CHN -blank rmctg
     If ADC("wRrTcTyp") = "CHN" Then
      Cancel = UCase(pv_NewValue) <> ""
      ErrMsg = "Raw Material Category should be blank": Exit Sub
    End If
   
    '*************************** ZUBIN **************************
  Case Is = UCase("wRrCmCtg")
  
    'KC.4 - Customer category validation. INP is applicable for 'P' customer only
    If UCase(ADC("wRrTcTyp")) = "INP" And UCase(pv_NewValue) <> "P" Then
      Cancel = True: ErrMsg = "Invalid Customer Category": Exit Sub
    End If
    
    'Cancel = Not moCn.RecSeek("Select PMCd From Param where PTyp= 'CMCTG' and PMCd= " + _
             "'" + pv_NewValue + "' ")
    '*************************** Geeta**************************
    ' 17th Jan 2004, EMR206
    If UCase(ms_MenuCd) = UCase("RmRtHist") Then
      Cancel = UCase(pv_NewValue) <> "C" And UCase(pv_NewValue) <> "P"
      ErrMsg = "Invalid Customer Category": Exit Sub
    Else
      '*** JAY 3.01 [NAS] ***
      '*** Jay 3.1.3 [CRP]
      '****** Sachin 3.3.0 (24-06-08) - [14. Cost Markup In Invoicing] ******
      'pg.5- MRP accept only "C" customer category
      If ADC("wRrTcTyp") = "QTT" Or ADC("wRrTcTyp") = "WTT" Or ADC("wRrTcTyp") = "MRP" Or ADC("wRrTcTyp") = "NAS" Or ADC("wRrTcTyp") = "CRP" Or ADC("wRrTcTyp") = "IMC" Or ADC("wRrTcTyp") = "IMK" Or ADC("wRrTcTyp") = "IML" Then
        Cancel = UCase(pv_NewValue) <> "C"
        ErrMsg = "Invalid Customer Category": Exit Sub
    '********************************Geeta*******************************
    '********bhavna added chk for CHN
      ElseIf ADC("wRrTcTyp") = "CHN" Then
        Cancel = UCase(pv_NewValue) <> "C" And UCase(pv_NewValue) <> "S"
        ErrMsg = "Invalid Customer Category (Enter C/S)": Exit Sub
    '********bhavna added chk for CHN
      'pg.5- MRD accept only "P" customer category
      ElseIf ADC("wRrTcTyp") = "MRD" Then
        Cancel = UCase(pv_NewValue) <> "P"
        ErrMsg = "Invalid Customer Category (Enter P)": Exit Sub
      Else
        Cancel = UCase(pv_NewValue) <> "C" And UCase(pv_NewValue) <> "P" And UCase(pv_NewValue) <> "S" And UCase(pv_NewValue) <> "T"
        ErrMsg = "Invalid Customer Category": Exit Sub
      End If
    End If
    
  Case Is = UCase("wRrSrt")
    If pv_NewValue <> "" Then
      Cancel = Not moCn.RecSeek("Select PMCd From Param where PTyp= 'SRTRR' " + _
                           "and PMCd = '" + pv_NewValue + "'")
      ErrMsg = "Invalid Sort Option": Exit Sub
    End If
  Case Is = UCase("wCpyCustCtgFr")
    Cancel = pv_NewValue <> "" And Not moCn.RecSeek("Select PmCd From Param where " + _
             "PTyp= 'CMCTG' and PmCd = '" + pv_NewValue + "'")
     ErrMsg = "Invalid Source Customer Category": Exit Sub
  Case Is = UCase("wCpyCustFr")
    Cancel = pv_NewValue <> "" And Not moCn.RecSeek("Select CmCd From CustMst where " + _
             "CmCtg= '" + ADC("wCpyCustCtgFr") + "' and CmCd = '" + pv_NewValue + "'")
     ErrMsg = "Invalid Source Customer Code": Exit Sub
  Case Is = UCase("wCpyCustTo")
    Cancel = pv_NewValue <> "" And Not moCn.RecSeek("Select CmCd From CustMst where " + _
             "CmCtg= '" + ADC("wRrCmCtg") + "' and CmCd = '" + pv_NewValue + "'")
     ErrMsg = "Invalid Target Customer Code": Exit Sub
  Case Is = UCase("wCpyCustNameFr")
    Cancel = pv_NewValue <> "" And (Not moCn.RecSeek("select CmName from CustMst where CmCtg= '" + ADC("wCpyCustCtgFr") + "' and CmName = '" + pv_NewValue + "'"))
    If Cancel = True Then ErrMsg = "Invalid Source Customer Name": Exit Sub
  Case Is = UCase("wCpyCustNameTo")
    Cancel = pv_NewValue <> "" And (Not moCn.RecSeek("select CmName from CustMst where CmCtg= '" + ADC("wRrCmCtg") + "' and CmName = '" + pv_NewValue + "'"))
    If Cancel = True Then ErrMsg = "Invalid Target Customer Name": Exit Sub
  '*************************** ZUBIN **************************
  ' 27th Nov 2003, EMR206
  Case Is = UCase("wRhYyyyMm")
  If UCase(ms_MenuCd) = UCase("RmRtHist") Then
      Cancel = Len(CStr(pv_NewValue)) <> 6 Or _
          (Mid(CStr(pv_NewValue), 1, 4) < 1900 Or Mid(CStr(pv_NewValue), 1, 4) > 2099) Or _
          (Mid(pv_NewValue, 5, 2) < 1 Or Mid(pv_NewValue, 5, 2) > 12)
      ErrMsg = "Invalid YyyyMm. Yyyy should be between 1900 and 2099 and Mm should be between 01 and 12 "
      Exit Sub
  End If
  '*************************** ZUBIN **************************
  ' Zubin 213
  Case Is = UCase("wCpyCurCdFr")
    If UCase(ADC("wCpyCustFr")) = ctSelfCmCd Or UCase(ADC("wCpyCustFr")) = ctValueCmCd Then
      Cancel = pv_NewValue <> "" And Not moCn.RecSeek("Select 'x' from Param where PTyp= 'CURNCY' and PMCd= '" + pv_NewValue + "' and " + _
              "PSCd= ''")
      If Cancel = True Then ErrMsg = "Invalid Currency Code": Exit Sub
    Else
      Cancel = pv_NewValue <> "" And Not moCn.RecSeek("Select 'x' from CustMst where CmCtg= '" + ADC("wCpyCustCtgFr") + "' and " + _
              "CmCd= '" + ADC("wCpyCustFr") + "' and CmCurCd= '" + pv_NewValue + "' ")
      If Cancel = True Then ErrMsg = "Invalid Currency Code": Exit Sub
    End If
  Case Is = UCase("wCpyCurCdTo")
    If UCase(ADC("wCpyCustTo")) = ctSelfCmCd Or UCase(ADC("wCpyCustTo")) = ctValueCmCd Then
      Cancel = pv_NewValue <> "" And Not moCn.RecSeek("Select 'x' from Param where PTyp= 'CURNCY' and PMCd= '" + pv_NewValue + "' and " + _
              "PSCd= ''")
      If Cancel = True Then ErrMsg = "Invalid Currency Code": Exit Sub
    Else
      Cancel = pv_NewValue <> "" And Not moCn.RecSeek("Select 'x' from CustMst where CmCtg= '" + ADC("wRrCmCtg") + "' and " + _
              "CmCd= '" + ADC("wCpyCustTo") + "' and CmCurCd= '" + pv_NewValue + "' ")
      If Cancel = True Then ErrMsg = "Invalid Currency Code": Exit Sub
    End If
  ' Zubin 213
  
 End Select
End Sub
Private Sub CmdCpyHist_Click()
  'PREVIOUS - EMR204 <<<<<<====================================
  'Call DispFra(CpyHist)
  '*************************** ZUBIN **************************
  ' EMR206
  Call DispFra(CpyHist)
  '*************************** ZUBIN **************************
End Sub
Private Sub CmdCpyHistGo_Click()
  '*************************** ZUBIN **************************
  ' 01st Dec 2003, EMR206
  Call CpyRmRtHist
  Call CmdCpyHist_Click
End Sub

Private Sub CmdILC_Click()
  If Not ms_PwdFlag Then _
     DispMsg "SEO Password Required", etError: Exit Sub
  Call DispFra(ILC)
  
End Sub
'***********Bhavna added import from excel
Private Sub CmdImpExc_Click()
  If Not ADC.SaveRec Then
    DispMsg "X Save Failed", etInfo
    Exit Sub
  End If

  Call DispFra(ImpExc)

End Sub

Private Sub CmdImpExcGo_Click()
If Trim(ADC("WXLFILE")) = "" Then DispMsg "Please Specify a Proper Excel File Name !!", etError: Exit Sub
If ADC("WXLSHTNO") = 0 Or ADC("WXLFRROW") = 0 Or ADC("WXLTOROW") = 0 Then DispMsg "Please Specify a Proper Sheet No./From Row/To Row !!", etError: Exit Sub

Call ImportExcel
Call DispFra(ImpExc)

End Sub

Private Sub CmdImpExcPath_Click()
  CdImpExc.FileName = ""
  CdImpExc.DialogTitle = "Select Rm Rate Excel File "
  CdImpExc.Filter = "Excel Files(*.xls;*.xlsx)|*.xls;*.xlsx"
  CdImpExc.FilterIndex = 1
  CdImpExc.ShowOpen
  ADC("WXLFILE") = CdImpExc.FileName

End Sub
'***********Bhavna added import from excel

Private Sub GrdRm_RowWhen(ByVal RowNum As Integer)
  '*** Set the col properties Style and Heading for the fields RrFrLn and RrToLn
      'based on the RmCtg and the Rate By option of the RmSCtg
    If GrdRm.IsNew(RowNum) Then GrdRm.Value(RowNum, "RrPrtKey") = ctCurrPrtn  '****** Sachin 3.02.0
    Call LnFrToPropSet(UCase("RrFrLn"))
    Call LnFrToPropSet(UCase("RrToLn"))
    '*****geeta******Emr207
    Call LnFrToPropSet(UCase("RrRngSz"))
    '*****geeta******Emr207
  
End Sub
Private Sub GrdRm_ColWhen(ByVal RowNum As Integer, ByVal ColName As String, Cancel As Boolean, ErrMsg As String, ByVal Action As MwfCtl.en_fgAction)
  '*** Set Helps For Each Field
  '*** Cost Rate can be entered only if the Customer is ZSELF or when the Rate TcType is 'XDM' or 'XIA' or 'XIO' or 'XZA'

  '****** Sachin 3.02.0 - Locking Previous Partition Records
  If GrdRm.Value(RowNum, "RrPrtKey") <> ctCurrPrtn Then Cancel = True: ErrMsg = "Cannot Edit Rates in Previous Partition"

  ' Zubin 212 (18/12/05)
  If ((UCase(ADC.MenuCd) = UCase("RmRt") And UCase(ADC("wRrTcTyp")) = "RM") Or _
                          UCase(ADC.MenuCd) = UCase("RmRtHist")) And (Not GrdRm.IsNew(RowNum)) And _
    (UCase(GrdRm.Value(RowNum, "RrCmCd")) = ctValueCmCd Or UCase(GrdRm.Value(RowNum, "RrCmCd")) = ctSelfCmCd) And _
    Not moCn.RecSeek("Select 'x' from Param where PTyp= 'USR' and PMCd= '" + gs_UsrCd + "' and PSCd= '' And " + _
                          "PNum = 1") Then
    ErrMsg = "Cannot Edit 'ZVALUE' and 'ZSELF' Records As You Are Not Authorised To Do So"
    Cancel = True: GrdRm.AllowDelete = False: Exit Sub
  Else
    GrdRm.AllowDelete = True
  End If
  ' Zubin 212 (18/12/05)
  
  Select Case ColName
  Case Is = UCase("RrCmCd")
    Call HlpList.CustCd(ADC("wRrCmCtg"))
    
  Case Is = UCase("RrCmCurCd")
    ' Zubin 213 (ctValueCmCd, "LMG", "LMP", "LMS" added)
    ' ******  Manali 3.03 - 18/06/08 - "LML" added
    ' ***** Manali 3.6.0 - 15/02/10 - "IMK" added
  '  If ADC("wRrTcTyp") = "RM" Then
   '   If UCase(GrdRm.Value(RowNum, "RrCmCd")) = "ZSELF" Then
    '      Call HlpList.PMCd("CURNCY")
   '   Else
    '      Cancel = True
'      End If
 '   Else
 'indigo-162- Tctyp-LS included
      If UCase(GrdRm.Value(RowNum, "RrCmCd")) = "ZSELF" Or UCase(GrdRm.Value(RowNum, "RrCmCd")) = ctValueCmCd _
        And (ADC("wRrTcTyp") = "RM" Or ADC("wRrTcTyp") = "LMG" Or ADC("wRrTcTyp") = "LMP" Or _
        ADC("wRrTcTyp") = "LMS" Or ADC("wRrTcTyp") = "LML" Or ADC("wRrTcTyp") = "IMK" Or ADC("wRrTcTyp") = "IMC" Or ADC("wRrTcTyp") = "LS" Or ADC("wRrTcTyp") = "IMC") Then
        Call HlpList.PMCd("CURNCY")
      Else
        Cancel = True
        ErrMsg = "Cannot enter Currency for this Customer Code"
      End If
  '  End If
  Case Is = UCase("RrSCtg")
    Call HlpList.PSCd("RMSCTG", ADC("wRrCtg"))
  Case Is = UCase("RrCd")
    Call HlpList.RmCd("N", "'" + ADC("wRrCtg") + "'", GrdRm.Value(RowNum, "RrSCtg"))
  Case Is = UCase("RrDmCtg")
    '********************Bhavna for Loss % for specific Design 5.1.0
    If ADC("wRrTcTyp") = "LS" And GrdRm.Value(RowNum, "RrDmCd") <> "" Then
        GrdRm.Value(RowNum, "RrDmCtg") = ""
        Cancel = True: ErrMsg = "Cannot Enter Design Ctg if Dsg Cd Non-Blank": Exit Sub
    Else
        Call HlpList.PMCd("DMCTG")
    End If
    '********************Bhavna for Loss % for specific Design 5.1.0
    
  Case Is = UCase("RrLabMCd")
    Call HlpList.PMCd("LABMCD")
  '*****************************Geeta **************************
    '' 07th Jan 2004, EMR206***********Check not required for RrTcTyp<>'QTT' ,'WTT'
   Case Is = UCase("RrCstRt")
    If ADC("wRrTcTyp") <> "XDM" And ADC("wRrTcTyp") <> "XIA" And ADC("wRrTcTyp") <> "XIO" And _
       ADC("wRrTcTyp") <> "XZA" And ADC("wRrTcTyp") <> "QTT" And ADC("wRrTcTyp") <> "WTT" Then

        If ADC("wRrTcTyp") <> "RM" And GrdRm.Value(RowNum, "RrCmCd") <> ctSelfCmCd And GrdRm.Value(RowNum, "RrCmCd") <> ctValueCmCd Then
         Cancel = True: ErrMsg = "Cannot Edit Value": Exit Sub
        ElseIf ADC("wRrTcTyp") = "RM" And Not ms_CstRtEdtYN And GrdRm.Value(RowNum, "RrCmCd") <> ctSelfCmCd And GrdRm.Value(RowNum, "RrCmCd") <> ctValueCmCd Then
         Cancel = True: ErrMsg = "Cannot Edit Value": Exit Sub
        End If
    End If
  Case Is = UCase("RrRngSz")
    Call HlpList.RangeSz(ADC("wRrCtg"), GrdRm.Value(RowNum, "RrSCtg"))
  Case Is = UCase("RrLsCtg")
   '********************Bhavna for Loss % for specific Design 5.1.0
    If ADC("wRrTcTyp") = "LS" And GrdRm.Value(RowNum, "RrDmCd") <> "" Then
        GrdRm.Value(RowNum, "p") = ""
        Cancel = True: ErrMsg = "Cannot Enter Loss Ctg if Dsg Cd Non-Blank": Exit Sub
    Else
        Call HlpList.PMCd("LSCTG")
    End If
    '********************Bhavna for Loss % for specific Design 5.1.0
        
  Case Is = UCase("RrSTWGrp")
    Call HlpList.PMCd("STWGRP")
  ' ****** Sachin 2.12 - 07/12/05 - DmCd In Tolerance
  ' **** Manali 3.03 - 18/06/08 - RmCtg "L" included
  Case Is = UCase("RrDmCd")
    ' ***** Manali 3.8.0 - IMK added below
    '********************Bhavna for Loss % for specific Design 5.1.0
    If ADC("wRrTcTyp") <> "LS" Then
      If ((ADC("wRrCtg") = "G" Or ADC("wRrCtg") = "P" Or ADC("wRrCtg") = "S" Or ADC("wRrCtg") = "L") And GrdRm.Value(RowNum, "RrDmCtg") <> "") Or _
       (ADC("wRrCtg") <> "G" And ADC("wRrCtg") <> "P" And ADC("wRrCtg") <> "S" And ADC("wRrCtg") <> "L" And GrdRm.Value(RowNum, "RrDmCtg") = "") Or _
        ADC("wRrTcTyp") = "IMK" Or ADC("wRrTcTyp") = "IMC" Then
        Call HlpList.DmCd("DM", GrdRm.Value(RowNum, "RrDmCtg"))
      End If
    Else
     If GrdRm.Value(RowNum, "RrDmCtg") <> "" Or GrdRm.Value(RowNum, "RrLsCtg") <> "" Then
        GrdRm.Value(RowNum, "RrDmCd") = ""
        Cancel = True: ErrMsg = "Cannot Enter Design Code if Dsg Ctg or Loss Ctg Non-Blank": Exit Sub
      Else
        Call HlpList.DmCd("DM")
      End If
    End If
    '********************Bhavna for Loss % for specific Design 5.1.0
  Case Is = UCase("RrRmKt")
    Call HlpList.PMCd("KT")
  ' ****** Sachin 2.12 - 07/12/05 - DmCd In Tolerance
  'pg.5- calling hlp for Suffix & size
  Case Is = UCase("RrDmSfx")
    Call HlpList.OdSfx(ctSelfCoCd, GrdRm.Value(RowNum, "RrDmCd"), ADC("wRrCmCtg"), GrdRm.Value(RowNum, "RrCmCd"))
  Case Is = UCase("RrDmSz")
    Call HlpList.PMCd("DMSZ")
  'ec.2- CstRtInc cannot edit if cmcd<>zself
  Case Is = UCase("RrFixMinTol")
    If ADC("wRrTcTyp") = "RM" And ADC("wRrCtg") = "X" Then
      If Not ms_CstRtEdtYN And GrdRm.Value(RowNum, "RrCmCd") <> ctSelfCmCd And GrdRm.Value(RowNum, "RrCmCd") <> ctValueCmCd Then _
        Cancel = True: ErrMsg = "Cannot Edit Value": Exit Sub
    End If

  End Select
End Sub
Private Sub GrdRm_ReCalc(ByVal RowNum As Integer, ByVal ColName As String, ByVal OldVal As Variant, ByVal SourceName As String, ByVal SourceColName As String)
  '*** Setting Heading, Vis, Ena Properties for RrLnFr, RrLnTo
  '*** RrFrLn on RrSCtg         Normal
      'RrFrLn = 0
      'Set the col properties Style and Heading
      'based on the RmCtg and the Rate By option of the RmSCtg
  '*** RrToLn on RrSCtg         Normal
      'RrToLn = 0
      'Set the col properties Style and Heading
      'based on the RmCtg and the Rate By option of the RmSCtg
  '*** RrCstRt on RrCmCd        Normal
      'RrCstRt = 0
Dim wSalRt As Double
  
  With GrdRm
  Select Case ColName
  Case Is = UCase("RrCmCurCd")
    If .Mode = fgmnorm Then
      ' Zubin 213 (ctValueCmCd, "LMG", "LMP", "LMS" added)
      ' ****** Manali 3.03 - 18/06/08 - "LML" added
      ' ***** Manali 3.6.0 - 15/02/10 - "IMK" added
      'pg.5
      'indigo-162- Tctyp-LS included
      If ADC("wRrTcTyp") <> "WTT" And ADC("wRrTcTyp") <> "QTT" Then
        If (.Value(RowNum, "RrCmCd") <> "ZSELF" And .Value(RowNum, "RrCmCd") <> ctValueCmCd And _
            (ADC("wRrTcTyp") = "RM" Or ADC("wRrTcTyp") = "LMG" Or ADC("wRrTcTyp") = "LMP" Or ADC("wRrTcTyp") = "IPD" Or _
            ADC("wRrTcTyp") = "LMS" Or ADC("wRrTcTyp") = "LML" Or ADC("wRrTcTyp") = "IMK" Or ADC("wRrTcTyp") = "IMC")) Then _
         GrdRm.Value(RowNum, "RrCmCurCd") = moCn.GetFldVal("select CmCurCd from CustMst where " + _
                                            "CmCtg= '" + ADC("wRrCmCtg") + "' " + _
                                            "and CmCd= '" + .Value(RowNum, "RrCmCd") + "' ")
      End If
     
    End If
  Case Is = UCase("RrFrLn")
    If .Mode = fgmnorm Then
      Call LnFrToPropSet(ColName)
      .Value(RowNum, "RrFrLn") = 0
    End If
  Case Is = UCase("RrToLn")
    If .Mode = fgmnorm Then
      Call LnFrToPropSet(ColName)
      .Value(RowNum, "RrToLn") = 0
    End If
  Case Is = UCase("RrCstRt")
    If .Mode = fgmnorm Then
      .Value(RowNum, "RrCstRt") = 0
    End If
  Case Is = UCase("wRrQw")
    If (.Mode = fgmnorm Or .Mode = fgminit Or .Mode = fgmCopy) Then
      .Value(RowNum, "wRrQw") = moCn.GetFldVal("Select RmQw from RmMst " + _
                                "where RmCd= '" + .Value(RowNum, "RrCd") + "'" + IIF(gs_Partition = ctCurrPrtn, " And RmPrtKey='" + ctCurrPrtn + "' ", ""))
    End If
    
  '****** Sachin 2.13.0 - 29-05-2006 ******
  Case Is = UCase("wRrQwCst")
    If (.Mode = fgmnorm Or .Mode = fgminit Or .Mode = fgmCopy) Then
      .Value(RowNum, "wRrQwCst") = moCn.GetFldVal("Select RmQwCst from RmMst " + _
                                "where RmCd= '" + .Value(RowNum, "RrCd") + "'" + IIF(gs_Partition = ctCurrPrtn, " And RmPrtKey='" + ctCurrPrtn + "' ", ""))
    End If
  '****** Sachin 2.13.0 - 29-05-2006 ******
  
  '******** Shilpa Emr 204 ******
  Case Is = UCase("RrSalRt")
    If (.Mode = fgmnorm Or .Mode = fgmCopy) Then
        If UCase(ADC("WRrTcTyp")) = "RM" And (UCase(ADC("wRrCtg")) = "C" Or UCase(ADC("wRrCtg")) = "D") And _
          UCase(ADC("wRrCmCtg")) = "C" And (UCase(.Value(RowNum, "RrCmCd")) = "ZSELF" Or UCase(.Value(RowNum, "RrCmCd")) = "ZVALUE") _
          And UCase(.Value(RowNum, "RrSCtg")) <> "" Then
            wSalRt = .Value(RowNum, "RrCstRt") + (.Value(RowNum, "RrCstRt") * (moCn.GetFldVal("Select RrSalRt from RmRt where " + _
                              "RrTcTyp = 'INC' and RrCmCd = '" + (.Value(RowNum, "RrCmCd")) + "' AND RrCtg = '" + UCase(ADC("wRrCtg")) + "' AND RrSCtg = '" + UCase(.Value(RowNum, "RrSCtg")) + "'") / 100))
            '*** Jay 3.01 *** [Stop ZSelf Round off]
            'If (wSalRt Mod 5) = 0 Then
            '  .Value(RowNum, "RrSalRt") = Round(wSalRt, 0)
            'ElseIf (wSalRt Mod 5) > 0 And (wSalRt Mod 5) < 2.5 Then
            '  .Value(RowNum, "RrSalRt") = Round(FnRndOff(wSalRt, 1, "RNDDN5"), 0)
            'ElseIf (wSalRt Mod 5) >= 2.5 And (wSalRt Mod 5) < 5 Then
            '  .Value(RowNum, "RrSalRt") = Round(FnRndOff(wSalRt, 1, "RNDUP5"), 0)
            'End If
            .Value(RowNum, "RrSalRt") = Round(wSalRt, 0)
            '*** Jay 3.01 *** [Stop ZSelf Round off]
        End If
    End If
  '******** Shilpa Emr 204 ******
  '******** geeta Emr 207 ******
  Case Is = UCase("RrRngSz")
    If .Mode = fgmnorm And ADC("wRrTcTyp") = "RSZ" Then
      Call LnFrToPropSet(ColName)
      .Value(RowNum, "RrRngSz") = ""
    End If
  '******** geeta Emr 207 ******
  
  ' ****** Sachin 2.12 - 09/12/05 - Design Code Depth in Tolerance
  Case Is = UCase("RrDmCd")
    If .Mode = fgmnorm And ADC("wRrTcTyp") = "WTT" Then .Value(RowNum, "RrDmCd") = ""
  ' ****** Sachin 2.12 - 09/12/05 - Design Code Depth in Tolerance
  End Select
  End With
End Sub
Private Sub GrdRm_ColChange(ByVal RowNum As Integer, ByVal ColName As String, ByVal NewValue As Variant, Cancel As Boolean, ErrMsg As String)
  '*** as set in TcStructure table, check if blank is permitted in the field
  '*** RrSCtg     Blank or valid PSCd('RMSCTG', adc(wRrCtg))
  '*** RrCd       Blank or valid RmCd from RmMst(RmCtg= adc(wRrCtg), RmSCtg= RrCtg and RmZ= 'N')
  '*** RrCmCd
      'If ws_BlnkYN= 'N' then valid CmCd from CustMst(CmCtg= adc(wRrCmCtg))
      'If ws_BlnkYN= 'Y' then '' or valid CmCd from CustMst(CmCtg= adc(wRrCmCtg))
  '*** RrDmCtg
      'If ws_BlnkYN= 'N' then valid PmCd('DMCTG')
      'If ws_BlnkYN= 'Y' then '' or valid PmCd('DMCTG')
  '*** RrLabMCd
      'If ws_BlnkYN= 'N' then valid PmCd('LABMCD')
      'If ws_BlnkYN= 'Y' then '' or valid PmCd('LABMCD')
  '*** RrCstRt
      'If ws_BlnkYN= 'N' then > 0
      'If ws_BlnkYN= 'Y' then 0 or > 0
  '*** RrSalRt
      'If ws_BlnkYN= 'N' then > 0
      'If ws_BlnkYN= 'Y' then 0 or > 0
  
  
  Dim ws_BlnkYN As String
  With GrdRm
    
    ws_BlnkYN = moCn.GetFldVal("Select TsBlank from TcStr where TsTbl='RmRt' and " + _
                            "TsTyp='" + ADC("wRrTcTyp") + "' and " + _
                            "TsSTyp='" + ADC("wRrCtg") + "' and " + _
                            "TsIdName='" + ColName + "' ")
                            
    Select Case UCase(ColName)
    Case Is = UCase("RrSCtg")
      '****** Sachin 3.3.0 (24-06-08) - [14. Cost Markup on Stock Rate] ******
      If (ADC("wRrTcTyp") = "IMK" Or ADC("wRrTcTyp") = "IMC") Then
        Cancel = Not moCn.RecSeek("Select PSCd From Param where PTyp= 'RMSCTG' " + _
                             "and PMCd = '" + ADC("wRrCtg") + "' and " + _
                             "PSCd = '" + NewValue + "'")
      Else
      '****** Sachin 3.3.0 (24-06-08) - [14. Cost Markup on Stock Rate] ******
        Cancel = (NewValue <> "") And (Not moCn.RecSeek("Select PSCd From Param where PTyp= 'RMSCTG' " + _
                             "and PMCd = '" + ADC("wRrCtg") + "' and " + _
                             "PSCd = '" + NewValue + "'"))
      End If
      ErrMsg = "Invalid Raw Material Sub Category": Exit Sub
      
    Case Is = UCase("RrCd")
      '****** Sachin 3.02.0 - Current Partition
      Cancel = (NewValue <> "") And (Not moCn.RecSeek("Select RmCd From RmMst where RmCtg= '" + ADC("wRrCtg") + "' " + _
                           "and RmSCtg = '" + .Value(RowNum, "RrSCtg") + "' and " + _
                           "RmCd = '" + NewValue + "' and RmZ= 'N' and RmPrtKey='" + ctCurrPrtn + "'"))
      ErrMsg = "Invalid Raw Material Code": Exit Sub
    Case Is = UCase("RrCmCd")
      If ws_BlnkYN = "N" And NewValue = "" Then Cancel = True: ErrMsg = "Enter Customer Code": Exit Sub
      '*** Jay 3.1.3 [CRP]
      If ADC("wRrTcTyp") = "CRP" And (NewValue = ctSelfCmCd Or NewValue = ctValueCmCd) Then Cancel = True: ErrMsg = "Customer Code Cannot Be " + ctSelfCmCd + " or " + ctValueCmCd: Exit Sub
      '*** Jay 3.1.3 [CRP]
      'pg.5- TcTyp-MRP accepts only ZSelf customer
      If ADC("wRrTcTyp") = "MRP" And NewValue <> ctSelfCmCd Then Cancel = True: ErrMsg = "Customer Code Must Be " + ctSelfCmCd: Exit Sub
      
      Cancel = (NewValue <> "") And (Not moCn.RecSeek("Select CmCd From CustMst where CmCtg= '" + ADC("wRrCmCtg") + "' " + _
                             "and CmCd = '" + NewValue + "'"))
      If Cancel = True Then ErrMsg = "Invalid Customer Code": Exit Sub
      'Indigo.1- Checking user access permission for the given customers
      If UCase(ms_MenuCd) = UCase("RmRt") Then
        Dim wErrMsg As String
        If Not ChkCmpUsrAccess(ADC("WRRCMCTG"), ADC.MenuCd, "", "", NewValue, wErrMsg) Then _
          Cancel = True: ErrMsg = "You don't have permission to add this Customer": Exit Sub
      End If
    Case Is = UCase("RrCmCurCd")
      Cancel = (NewValue <> "") And (Not moCn.RecSeek("Select PMCd From Param where  " + _
                           "PTyp= 'CURNCY' and PMCd = '" + NewValue + "' and " + _
                           "PSCd = '' "))
      ErrMsg = "Invalid Customer Currency Code": Exit Sub
      
    Case Is = UCase("RrDmCtg")
      If ws_BlnkYN = "N" And NewValue = "" Then _
         Cancel = True: ErrMsg = "Enter Design Category": Exit Sub
      Cancel = (NewValue <> "") And (Not moCn.RecSeek("Select PMCd From Param where PTyp= 'DMCTG' " + _
                           "and PMCd = '" + NewValue + "'"))
      ErrMsg = "Invalid Design Category": Exit Sub
    Case Is = UCase("RrLabMCd")
      If ws_BlnkYN = "N" And NewValue = "" Then _
         Cancel = True: ErrMsg = "Enter Labour Main Code": Exit Sub
      Cancel = (NewValue <> "") And (Not moCn.RecSeek("Select PMCd From Param where PTyp= 'LABMCD' " + _
                           "and PMCd = '" + NewValue + "'"))
      ErrMsg = "Invalid Labour Main Code": Exit Sub
    Case Is = UCase("RrCstRt")
     '******************Geeta*********************
     '08th jan 2004 ********No check required for RrTcTyp='QTt' and "WTT'
      If ADC("wRrTcTyp") <> "QTT" And ADC("wRrTcTyp") <> "WTT" And (ws_BlnkYN = "N") And (NewValue = 0) Then
        Cancel = True: ErrMsg = "Cost Rate Cannot Be Zero": Exit Sub
      End If
    Case Is = UCase("RrSalRt")
      '******************Geeta*********************
      '08th jan 2004 ********No check required for RrTcTyp='QTt' and "WTT'
      '*** Jay 3.01 *** [ILC]
      '****** Sachin 3.3.0 (24-06-08) - [14. Cost Markup In Invoicing] ******
      '**********Bhavna added 10/50/100
      If ADC("wRrTcTyp") = "ILC" Or (ADC("wRrTcTyp") = "IMK" And (ADC("wRrCtg") = "D" Or ADC("wRrCtg") = "C")) Or _
        (ADC("wRrTcTyp") = "IMC" And (ADC("wRrCtg") = "D" Or ADC("wRrCtg") = "C")) Then
        If NewValue <> 0 And NewValue <> 1 And NewValue <> 5 And NewValue <> 10 And NewValue <> 50 And NewValue <> 100 Then
          Cancel = True: ErrMsg = "Valid Round Of Values Can Be (0/ 1/ 5/ 10/ 50/ 100)": Exit Sub
        End If
      
      ' ***** Manali 3.6.0 - 28/10/09 - Cost Markup In Invoicing for Accessories
      ElseIf ADC("wRrTcTyp") = "IMK" And ADC("wRrCtg") = "X" And ADC("wRrCtg") = "G" And ADC("wRrCtg") = "P" And ADC("wRrCtg") = "S" And ADC("wRrCtg") = "L" Then
        If NewValue <> 0# And NewValue <> 1# And NewValue <> 5# And NewValue <> 0.5 And NewValue <> 0.05 Then
          Cancel = True: ErrMsg = "Valid Round Of Values Can Be (0/ 0.05/ 0.5 / 1/ 5)": Exit Sub
        End If
      '*** Jay 3.1.3 [CRP]
      ElseIf ADC("wRrTcTyp") = "IMC" And ADC("wRrCtg") = "X" And ADC("wRrCtg") = "G" And ADC("wRrCtg") = "P" And ADC("wRrCtg") = "S" And ADC("wRrCtg") = "L" Then
        If NewValue <> 0# And NewValue <> 1# And NewValue <> 5# And NewValue <> 0.5 And NewValue <> 0.05 Then
          Cancel = True: ErrMsg = "Valid Round Of Values Can Be (0/ 0.05/ 0.5 / 1/ 5)": Exit Sub
        End If
     
      ElseIf ADC("wRrTcTyp") = "CRP" And NewValue <= 0 Then
        Cancel = True: ErrMsg = "Customer Rate Purity Should Be Greater Than Zero": Exit Sub
      '*** Jay 3.1.3 [CRP]
      'mw.136 Weight <> 0
      ElseIf ADC("wRrTcTyp") = "STW" And (ws_BlnkYN = "N") And NewValue = 0 Then
        Cancel = True: ErrMsg = "Weight Cannot Be Zero": Exit Sub
      ElseIf ADC("wRrTcTyp") <> "QTT" And ADC("wRrTcTyp") <> "WTT" And (ws_BlnkYN = "N") And (NewValue = 0) Then
        Cancel = True: ErrMsg = "Sales Rate Cannot Be Zero": Exit Sub
      End If
      '*** Jay 3.01 *** [ILC]
    Case Is = UCase("RrLsCtg")
      If ws_BlnkYN = "N" And NewValue = "" Then _
         Cancel = True: ErrMsg = "Enter Loss Category": Exit Sub
      Cancel = (NewValue <> "") And (Not moCn.RecSeek("Select PMCd From Param where PTyp= 'LSCTG' " + _
                           "and PMCd = '" + NewValue + "'"))
      ErrMsg = "Invalid Loss Category": Exit Sub
    Case Is = UCase("RrSTWGrp")
      If ws_BlnkYN = "N" And NewValue = "" Then _
         Cancel = True: ErrMsg = "Enter Size To Weight Group": Exit Sub
      Cancel = (NewValue <> "") And (Not moCn.RecSeek("Select PMCd From Param where PTyp= 'STWGRP' " + _
                           "and PMCd = '" + NewValue + "'"))
      ErrMsg = "Invalid Size To Weight Group": Exit Sub
    Case Is = UCase("RrSalRt")
    If (.Mode = fgmnorm Or .Mode = fgmCopy) And UCase(ADC("WRRTCTYP")) = "RM" And _
        (UCase(ADC("wRrCtg")) = "C" Or UCase(ADC("wRrCtg")) = "D") And _
        UCase(ADC("wRrCmCtg")) = "C" And (UCase(.Value(RowNum, "RrCmCd")) = "ZSELF" Or UCase(.Value(RowNum, "RrCmCd")) = "ZVALUE") _
        And UCase(.Value(RowNum, "RrSCtg")) <> "" Then
      .Value(RowNum, "RrSalRt") = .Value(RowNum, "RrCstRt") + (.Value(RowNum, "RrCstRt") * (moCn.GetFldVal("Select RrSalRt from RmRt where RrPrtKey='" + ctCurrPrtn + "' and " + _
                              "RrTcTyp = 'INC' and RrCmCd = '" + (.Value(RowNum, "RrCmCd")) + "' AND RrCtg = '" + UCase(ADC("wRrCtg")) + "' AND RrSCtg = '" + UCase(.Value(RowNum, "RrSCtg")) + "'") / 100))
    End If
    ' ****** Sachin 2.12 - 07/12/05 - DmCd in Tolerance
    Case Is = UCase("RrDmCd")
      '****** Sachin 3.02.0 - Current partition added
      ' ***** Manali 3.03 - 18/06/08 - RmCtg "L" included
      Cancel = (NewValue <> "") And (Not moCn.RecSeek("Select 'x' From DsgMst Where DmPrtKey='" + ctCurrPrtn + "' and  DmTcTyp='DM' And DmCd='" + NewValue + "'" + _
                     IIF(UCase(ADC("WRRTCTYP")) = "WTT" And (ADC("WRRCTG") = "G" Or ADC("WRRCTG") = "P" Or ADC("WRRCTG") = "S" Or ADC("WRRCTG") = "L"), " And DmCtg = '" + .Value(RowNum, "RrDmCtg") + "'", "") + _
                     " And DmSz=''"))
      ErrMsg = "Not a Valid Design Code ": Exit Sub
      
      If ADC("WRRTCTYP") = "MRP" Then
        Cancel = NewValue = ""
        If Cancel Then
          ErrMsg = "Design Cd cannot be blank"
        End If
      End If
    ' ****** Sachin 2.12 - 07/12/05 - DmCd in Tolerance
    Case Is = UCase("RrRmKt")
      If Trim(NewValue) <> "" Then
        Cancel = Not moCn.RecSeek("Select PMCd From Param where PTyp= 'KT' " + _
                             "and PMCd = '" + NewValue + "'")
        ErrMsg = "Invalid Raw Material Karat": Exit Sub
      End If
    'pg.5- checking valid DmSz
    Case Is = UCase("RrDmSz")
      Cancel = (NewValue <> "") And (Not moCn.RecSeek("Select PMCd From Param where PTyp= 'DMSZ' " + _
                             "and PMCd = '" + NewValue + "'"))
      ErrMsg = "Invalid Design Size": Exit Sub
    Case Is = UCase("RrFixMaxTol")
      'mw.136 max wt <> 0
      If ADC("wRrTcTyp") = "STW" And (ws_BlnkYN = "N") And NewValue = 0 Then _
        Cancel = True: ErrMsg = "Max Weight Cannot Be Zero": Exit Sub
        
  End Select
  End With
End Sub
 
  Private Sub GrdRm_InitKey(ByVal RowNum As Integer)
  '*** Initialise Key Fields
  With GrdRm
  '*************************** ZUBIN **************************
  ' 01st Dec 2003, EMR206 - Set Field names as per MenuCd
    If UCase(ms_MenuCd) = UCase("RmRt") Then
      .Store "RrTcTyp", ADC("wRrTcTyp")
      .Store "RrCtg", ADC("wRrCtg")
      .Store "RrCmCtg", ADC("wRrCmCtg")
    ElseIf UCase(ms_MenuCd) = UCase("RmRtHist") Then
      .Store "RhTcTyp", ADC("wRrTcTyp")
      .Store "RhCtg", ADC("wRrCtg")
      .Store "RhCmCtg", ADC("wRrCmCtg")
      .Store "RhYyyyMm", ADC("wRhYyyyMm")
    End If
  End With
End Sub

Private Sub ADC_SetRecSource()
  '*** Since this is an 'A' Entry, there is no RecSource associated with ADC. So set the
      'DataMode property of ADC control as xdatUnbound
  ADC.DataMode = xdatUnbound
  Call GrdPropSet
 '*************************** ZUBIN **************************
  ' 29th Nov 2003, EMR206 - Set Field names as per MenuCd
  Call GrdTblField(ms_MenuCd)
  '*************************** ZUBIN **************************
End Sub
Private Sub GrdRm_SetRecSource()
  '*** Set the Record Source of the Grid GrdRm
      'Default sort option is RrCmCd, RrSCtg, RrCd, RrFrLn, RrToLn
  Dim wsql As String
  Dim ws_Cnd As String, ws_OrdBy As String
  Dim ws_Tbl As String  'ZUBIN - EMR206, 01st Dec 2003
  Dim ws_UsrLvl As Integer
  
  Dim ws_Prtn As String
  
  ws_Cnd = ADC.RepCond
  ws_Cnd = IIF(ws_Cnd <> "", " where ", " and ") + ws_Cnd
  ws_Tbl = IIF(UCase(ms_MenuCd) = UCase("RmRt"), "RmRt", "RmRtHist")
  ws_UsrLvl = moCn.GetFldVal("Select PNum from Param where PTyp= 'USR' and PMCd= '" + gs_UsrCd + "' and PSCd= ''")
  
  If ADC("wRrSrt") = "" Then
  'PREVIOUS - EMR204 <<<<<<====================================
  '' defa order by to include CustCd, RrSCtg, RrCd
  '  ws_OrdBy = " Order By RrCmCd, RrCmCurCd, RrSCtg, RrCd, RrFrLn, RrToLn "
  'PREVIOUS - EMR204 <<<<<<====================================
    
  '*************************** ZUBIN **************************
  ' 27th Nov 2003, EMR206
    'IG.72-If Menucode is RmRt and TcType is STW or BSW then sorting should be CustCd, SubCtg, STW Group,
    'otherwise sorting option is common for all TcTyp
    If UCase(ms_MenuCd) = UCase("RmRt") Then
      If ADC("WRRTCTYP") = "STW" Or ADC("WRRTCTYP") = "BSW" Then
        ws_OrdBy = " Order By RrCmCd, RrSCtg, RrSTWGrp, RrFrLn, RrToLn "
      Else
        ws_OrdBy = " Order By RrCmCd, RrCmCurCd, RrSCtg, RrCd, RrFrLn, RrToLn "
      End If
    ElseIf UCase(ms_MenuCd) = UCase("RmRtHist") Then
      ws_OrdBy = " Order By RhCmCd, RhCmCurCd, RhSCtg, RhCd, RhFrLn, RhToLn "
    End If
  Else
  'PREVIOUS - EMR204 <<<<<<====================================
  '  ws_OrdBy = "Order By " + moCn.GetFldVal("Select PDesc225 from Param " + _
             "where PTyp='SRTRR' and PMCd='" + ADC("wRrSrt") + "'")
  'PREVIOUS - EMR204 <<<<<<====================================
    
    If UCase(ms_MenuCd) = UCase("RmRt") Then
      ws_OrdBy = "Order By " + moCn.GetFldVal("Select PDesc225 from Param " + _
                  "where PTyp='SRTRR' and PMCd='" + ADC("wRrSrt") + "'")
    ElseIf UCase(ms_MenuCd) = UCase("RmRtHist") Then
      ws_OrdBy = "Order By " + moCn.GetFldVal("Select PDesc225 from Param " + _
                  "where PTyp='SRTRH' and PMCd='" + ADC("wRrSrt") + "'")
    End If
  '*************************** ZUBIN **************************
  End If
  'PREVIOUS - EMR204 <<<<<<====================================
  ' wsql = " Select * from RmRt " + ws_Cnd + ws_OrdBy
  'PREVIOUS - EMR204 <<<<<<====================================
  
  '*************************** ZUBIN **************************
   wsql = " Select * from " + ws_Tbl + ws_Cnd + ws_OrdBy
  '*************************** ZUBIN **************************
  
  GrdRm.RecSource = wsql
  
End Sub

Private Sub CmdCpy_Click()
  If ms_RrTcTyp = "STW" Or ms_RrTcTyp = "RSZ" Then
    DispMsg "Copy Option Not Valid For This TcType", etError: Exit Sub
  End If
  If ms_RrTcTyp <> "RM" And ms_RrTcTyp <> "LMG" And ms_RrTcTyp <> "LMP" And ms_RrTcTyp <> "LMS" And ms_RrTcTyp <> "LML" And ms_RrTcTyp <> "IMK" And ms_RrTcTyp <> "IMC" Then
    LblCurCdFr.Visible = False: LblCurCdTo.Visible = False
    LblMulDivFr.Visible = False: ADC("wMulDivFr").Visible = False
    LblCnvRtFr.Visible = False: ADC("wCnvRtFr").Visible = False
    ADC("wCpyCurCdFr").Visible = False: ADC("wCpyCurCdTo").Visible = False
  Else
    LblCurCdFr.Visible = True: LblCurCdTo.Visible = True
    LblMulDivFr.Visible = True: ADC("wMulDivFr").Visible = True
    LblCnvRtFr.Visible = True: ADC("wCnvRtFr").Visible = True
    ADC("wCpyCurCdFr").Visible = True: ADC("wCpyCurCdTo").Visible = True
  End If
  Call DispFra(Cpy)
End Sub
Private Sub CmdCpyGo_Click()
  Call CpyRmRt
  Call CmdCpy_Click
End Sub
Private Sub CpyRmRt()
  Dim ws_RmCnd As String, ws_CurncyCnd As String
  Dim ws_TgtCurncyCnd As String
  Dim ws_SrcMulDiv As String

  Call CmpString("C")     ' to set the Compare String Property of the ADC controls when the Copy Option Is clicked
  
  '*** Checking if copy is possible
  If Not moCn.RecSeek("Select PmCd From Param where PTyp= 'CMCTG' " + _
     "and PmCd = '" + ADC("wCpyCustCtgFr") + "'") Then _
     DispMsg "Enter Proper Source Customer Category", etError: Exit Sub
 
  If Not moCn.RecSeek("Select CmCd From CustMst where CmCtg= '" + ADC("wCpyCustCtgFr") + "' " + _
     "and CmCd = '" + ADC("wCpyCustFr") + "'") Then _
     DispMsg "Enter Proper Source Customer Code", etError: Exit Sub
  
  If Not moCn.RecSeek("Select CmCd From CustMst where CmCtg= '" + ADC("wRrCmCtg") + "' " + _
     "and CmCd = '" + ADC("wCpyCustTo") + "'") Then _
     DispMsg "Enter Proper Target Customer Code", etError: Exit Sub
     
  'Zubin 212
  If UCase(ms_MenuCd) = UCase("RmRt") Then
    Dim wErrMsg As String
    If Not ChkCmpUsrAccess(ADC("wCpyCustCtgFr"), ADC.MenuCd, "", "", ADC("wCpyCustTo"), wErrMsg) Then _
      DispMsg "You don't have permissions to add this Customer", etError: Exit Sub
  End If
  
  If UCase(ADC("wRrTcTyp")) = "RM" And (UCase(ADC("WCPYCUSTTO")) = "ZSELF" Or UCase(ADC("WCPYCUSTTO")) = "ZVALUE") And _
    Not (moCn.RecSeek("Select 'x' from Param where PTyp= 'USR' and PMCd= '" + gs_UsrCd + "' and PSCd= '' and PNum= 1")) Then
    DispMsg "Cannot Add Records For 'ZSELF' and 'ZVALUE' As User Is Not Authorised To Do So", etError: Exit Sub
  End If
  ' Zubin 212
  
  ' Zubin 213
  ' ***** Manali 3.03 - 18/06/08 - "LML" added
  ' ***** Manali 3.6.1 - 15/02/10 - "IMK" added
  If (ADC("wRrTcTyp") = "RM" Or ms_RrTcTyp = "LMG" Or ms_RrTcTyp = "LMP" Or ms_RrTcTyp = "LMS" Or ms_RrTcTyp = "LML" Or ms_RrTcTyp = "IMK" Or ms_RrTcTyp = "IMC") _
    And (ADC("wCpyCurCdFr") = "" Or ADC("wCpyCurCdTo") = "") Then _
      DispMsg "Source-Target Currency Currency Code Cannot Be Blank", etWarning: Exit Sub

  If (ADC("wRrTcTyp") = "RM" Or ms_RrTcTyp = "LMG" Or ms_RrTcTyp = "LMP" Or ms_RrTcTyp = "LMS" Or ms_RrTcTyp = "LML" Or ms_RrTcTyp = "IMK" Or ms_RrTcTyp = "IMC") _
    And ADC("wMulDivFr") = "" Then _
      DispMsg "Source-Target Currency Conversion Is Not Entered In The System", etWarning: Exit Sub

  If (ADC("wRrTcTyp") = "RM" Or ms_RrTcTyp = "LMG" Or ms_RrTcTyp = "LMP" Or ms_RrTcTyp = "LMS" Or ms_RrTcTyp = "LML" Or ms_RrTcTyp = "IMK" Or ms_RrTcTyp = "IMC") _
    And ADC("wCnvRtFr") = "0" Then _
      DispMsg "Source-Target Currency Conversion Factor Cannot Be Zero", etWarning: Exit Sub
  ' ***** Manali 3.03 - 18/06/08 - "LML" added
  ' Zubin 213

  ws_RmCnd = ADC.RepCond
  ws_RmCnd = IIF(ws_RmCnd <> "", " and " + ws_RmCnd, "")
  
  ' Zubin 213 (Before 213)
  'ws_TgtCurCd = moCn.GetFldVal("Select CmCurCd From CustMst where CmCtg= '" + ADC("wRrCmCtg") + "' " + _
                "and CmCd = '" + ADC("wCpyCustTo") + "' ")
  ' Zubin 213 (Before 213)
  'ws_CurncyCnd = IIF(ADC("wRrTcTyp") = "RM", " and RrCmCurCd= '" + ws_TgtCurCd + "' ", " ")
  
  ' Zubin 213
  ' ****** Manali 3.03 - 18/06/08 - "LML" added
  ' ****** Manali 3.6.1 - 15/02/10 - "IMK" added
  ws_CurncyCnd = IIF(ADC("wRrTcTyp") = "RM" Or ms_RrTcTyp = "LMG" Or ms_RrTcTyp = "LMP" Or ms_RrTcTyp = "LMS" Or ms_RrTcTyp = "LML" Or ms_RrTcTyp = "IMK" Or ms_RrTcTyp = "IMC", _
                    " and RrCmCurCd= '" + ADC("wCpyCurCdFr") + "' ", " ")
  ws_TgtCurncyCnd = IIF(ADC("wRrTcTyp") = "RM" Or ms_RrTcTyp = "LMG" Or ms_RrTcTyp = "LMP" Or ms_RrTcTyp = "LMS" Or ms_RrTcTyp = "LML" Or ms_RrTcTyp = "IMK" Or ms_RrTcTyp = "IMC", _
                    " and RrCmCurCd= '" + ADC("wCpyCurCdTo") + "' ", " ")
  
  ' Zubin 213 (ws_TgtCurncyCnd instead of ws_CurncyCnd)
  If moCn.RecSeek("Select 'x' from RmRt where RrPrtKey='" + ctCurrPrtn + "' and RrTcTyp='" + ADC("wRrTcTyp") + "' and " + _
          "RrCtg= '" + ADC("wRrCtg") + "' and RrCmCtg= '" + ADC("wRrCmCtg") + "' and " + _
          "RrCmCd= '" + ADC("wCpyCustTo") + "' " + ws_TgtCurncyCnd + ws_RmCnd) Then
    DispMsg "Cannot Copy as RM Rates Exist For Given Cust And Given Ranges", etError
    Exit Sub
  End If
  
' ###################################  Manoj #### Ver: 2.0.6  ###################################
' Date:   21/01/2004
  If moCn.RecSeek("Select 'x' from RmRt where RrPrtKey='" + ctCurrPrtn + "' and RrTcTyp='" + ADC("wRrTcTyp") + "' and " + _
             "RrCtg='" + ADC("wRrCtg") + "' and RrCmCtg= '" + ADC("wCpyCustCtgFr") + "' and " + _
             "RrCmCd='" + ADC("wCpyCustFr") + "' " + ws_CurncyCnd + ws_RmCnd + _
             " And Exists (Select RmCd from RmMst Where RmPrtKey=RrPrtKey and RmCd= RrCd And RmValidYN = 'N')") Then
    DispMsg "Cannot Copy as Some Invalid RM are Present", etError
    Exit Sub
  End If
' ###########################################  Manoj  ###########################################
'*************************Geeta*************************Emr206******22/04/04*************
'****** Sachin 3.02.0 - Partition Key added
If moCn.RecSeek("select 'BS' from RmRt Rr where Rr.RrPrtKey='" + ctCurrPrtn + "' and " + _
                "Rr.RrTcTyp='BS' and Rr.RrCmCtg='" + ADC("wCpyCustCtgFr") + "' and  Rr.RrCmCd='" + ADC("wCpyCustFr") + "'  and  " + _
                "Rr.RrCmCurCd='' and  Rr.RrCtg='" + ADC("wRrCtg") + "' and " + _
                "Rr.RrSctg in ('" + ADC("wCpySCtgFr") + "','" + ADC("wCpySCtgTo") + "') and  Rr.RrSTWGrp='' and   Rr.RrCd='' and  Rr.RrDmCtg='' and  " + _
                "Rr.RrLsCtg=''  and  Rr.RrLabMCd='' and  Rr.RrFrLn='' and Rr.RrToLn='' " + _
                "and not exists (select * from RmRt Rt where Rt.RrPrtKey='" + ctCurrPrtn + "' and " + _
                "Rt.RrTcTyp='BS' and Rt.RrCmCtg='C' and Rt.RrCmCd='" + ADC("wCpyCustTo") + "'  and  " + _
                "Rt.RrCmCurCd='' and Rt.RrCtg='" + ADC("wRrCtg") + "' and " + _
                "Rt.RrSctg=Rr.RrSctg and Rt.RrSTWGrp='' and  Rt.RrCd='' and Rt.RrDmCtg='' and  " + _
                "Rt.RrLsCtg=''  and  Rt.RrLabMCd='' and  Rt.RrFrLn='' and Rt.RrToLn='') ") Then
      DispMsg "Cannot Copy as From Customer Record Exists in 'BS' for some RmSCtg ", etError
    Exit Sub
 End If
'*************************Geeta**************************************
  
  ' Zubin 213
  ws_SrcMulDiv = MulDivOp(ADC("wMulDivFr"))
  
  '*** Beginning the copy process
  GrdRm.StartCopy
  MWLib.BeginProcess Me, "Copying Raw Material Rates ..."
  Dim wRsCustRmRt As MwfLib.MDORowSet, wi_Row As Integer
  Set wRsCustRmRt = moCn.OpenRes("Select * from RmRt where RrPrtKey='" + ctCurrPrtn + "' and RrTcTyp= '" + ADC("wRrTcTyp") + "' and " + _
             "RrCtg= '" + ADC("wRrCtg") + "' and RrCmCtg= '" + ADC("wCpyCustCtgFr") + "' and " + _
             "RrCmCd= '" + ADC("wCpyCustFr") + "' " + ws_CurncyCnd + ws_RmCnd + _
             " Order By RrSCtg, RrCd")
  With wRsCustRmRt
    Do While Not (.EOF Or .BOF)
      MWLib.RunningProcess "Copying RM Rates For RM SubCtg " + !RrSCtg
      GrdRm.AddItem
      wi_Row = GrdRm.Rows - 1
' **** Zubin 211 (below line makes the cell value= '') **** '
'''      GrdRm.Row = wi_Row
' **** Zubin 211 **** '
      GrdRm.Value(wi_Row, "RrPrtKey") = ctCurrPrtn   '****** Sachin 3.02.0
      GrdRm.Value(wi_Row, "RrCmCd") = ADC("wCpyCustTo")
      GrdRm.Value(wi_Row, "RrSCtg") = !RrSCtg
      GrdRm.Value(wi_Row, "RrSTWGrp") = !RrSTWGrp
      GrdRm.Value(wi_Row, "RrCd") = !RrCd
      GrdRm.Value(wi_Row, "RrDmCtg") = !RrDmCtg
      ' ****** Sachin 2.12 - 07/12/05 - Dsg Cd in Tolerance
      GrdRm.Value(wi_Row, "RrDmCd") = !RrDmCd
      ' ****** Sachin 2.12 - 07/12/05 - Dsg Cd in Tolerance
      GrdRm.Value(wi_Row, "RrLsCtg") = !RrLsCtg
      GrdRm.Value(wi_Row, "RrLabMCd") = !RrLabMCd
      GrdRm.Value(wi_Row, "RrRngSz") = !RrRngSz
      ' ***** Manali 3.03 - 18/06/08 - RmCtg "L" and "LML" included
      If (ms_RrTcTyp = "RM" And (ADC("wRrCtg") = "G" Or ADC("wRrCtg") = "P" Or ADC("wRrCtg") = "S" Or ADC("wRrCtg") = "L")) Or ms_RrTcTyp = "LMG" Or ms_RrTcTyp = "LMP" Or ms_RrTcTyp = "LMS" Or ms_RrTcTyp = "LML" Then
        GrdRm.Value(wi_Row, "RrFrLn") = IIF(moCn.GetFldVal("Select " + CStr(!RrFrLn) + " " + ws_SrcMulDiv + " " + CStr(ADC("wCnvRtFr"))) > ctLmeLimit, ctLmeLimit, moCn.GetFldVal("Select " + CStr(!RrFrLn) + " " + ws_SrcMulDiv + " " + CStr(ADC("wCnvRtFr"))))
        GrdRm.Value(wi_Row, "RrToLn") = IIF(moCn.GetFldVal("Select " + CStr(!RrToLn) + " " + ws_SrcMulDiv + " " + CStr(ADC("wCnvRtFr"))) > ctLmeLimit, ctLmeLimit, moCn.GetFldVal("Select " + CStr(!RrToLn) + " " + ws_SrcMulDiv + " " + CStr(ADC("wCnvRtFr"))))
      Else
        GrdRm.Value(wi_Row, "RrFrLn") = !RrFrLn
        GrdRm.Value(wi_Row, "RrToLn") = !RrToLn
      End If
      ' Zubin 213
      ' ***** Manali 3.03 -18/06/08 - "LML" included
      ' ***** Manali - 3.6.1 - 15/02/10 - "IMK" added
      If ADC("wRrTcTyp") = "RM" Or ms_RrTcTyp = "LMG" Or ms_RrTcTyp = "LMP" Or ms_RrTcTyp = "LMS" Or ms_RrTcTyp = "LML" Or ms_RrTcTyp = "IMK" Or ms_RrTcTyp = "IMC" Then
        GrdRm.Value(wi_Row, "RrCmCurCd") = ADC("wCpyCurCdTo")
       Else
        GrdRm.Value(wi_Row, "RrCmCurCd") = !RrCmCurCd
      End If
      
      If ADC("wRrTcTyp") = "XDM" Or ADC("wRrTcTyp") = "XIA" Or ADC("wRrTcTyp") = "XIO" Or _
         ADC("wRrTcTyp") = "XZA" Then
        GrdRm.Value(wi_Row, "RrCstRt") = !RrCstRt
        'If adc("wRrTcTyp") <> "XZA" Then
        '  GrdRm.Value(wi_Row, "RrXwtPer") = !RrXWtPer
        'End If
      Else
        If (ADC("wCpyCustTo") = ctSelfCmCd) Or (ADC("wCpyCustTo") = ctValueCmCd) Then
          ' Zubin 213 (Before 213)
          'GrdRm.Value(wi_Row, "RrCstRt") = !RrCstRt
          ' Zubin 213
          ' ***** Manali 3.03 - 18/06/08 - "LML" added
          If ADC("wRrTcTyp") = "RM" Or ms_RrTcTyp = "LMG" Or ms_RrTcTyp = "LMP" Or ms_RrTcTyp = "LMS" Or ms_RrTcTyp = "LML" Then
            GrdRm.Value(wi_Row, "RrCstRt") = IIF(moCn.GetFldVal("Select " + CStr(!RrCstRt) + " " + ws_SrcMulDiv + " " + CStr(ADC("wCnvRtFr"))) > ctLmeSalRtLimit, ctLmeSalRtLimit, moCn.GetFldVal("Select " + CStr(!RrCstRt) + " " + ws_SrcMulDiv + " " + CStr(ADC("wCnvRtFr"))))
          Else
            GrdRm.Value(wi_Row, "RrCstRt") = !RrCstRt
          End If
        Else
          If ADC("wRrTcTyp") = "RM" And ms_CstRtEdtYN Then
            GrdRm.Value(wi_Row, "RrCstRt") = IIF(moCn.GetFldVal("Select " + CStr(!RrCstRt) + " " + ws_SrcMulDiv + " " + CStr(ADC("wCnvRtFr"))) > ctLmeSalRtLimit, ctLmeSalRtLimit, moCn.GetFldVal("Select " + CStr(!RrCstRt) + " " + ws_SrcMulDiv + " " + CStr(ADC("wCnvRtFr"))))
          End If
        End If
      End If
      
      ' Zubin 213 (Before 213)
      'GrdRm.Value(wi_Row, "RrSalRt") = !RrSalRt
      ' Zubin 213
      ' ***** Manali 3.03 - 18/06/08 - "LML" added
      If ADC("wRrTcTyp") = "RM" Or ms_RrTcTyp = "LMG" Or ms_RrTcTyp = "LMP" Or ms_RrTcTyp = "LMS" Or ms_RrTcTyp = "LML" Then
        GrdRm.Value(wi_Row, "RrSalRt") = IIF(moCn.GetFldVal("Select " + CStr(!RrSalRt) + " " + ws_SrcMulDiv + " " + CStr(ADC("wCnvRtFr"))) > ctLmeSalRtLimit, ctLmeSalRtLimit, moCn.GetFldVal("Select " + CStr(!RrSalRt) + " " + ws_SrcMulDiv + " " + CStr(ADC("wCnvRtFr"))))
      Else
        GrdRm.Value(wi_Row, "RrSalRt") = !RrSalRt
      End If
      
      '*** Jay 3.01 *** [ILC]
      If ADC("wRrTcTyp") = "ILC" Then GrdRm.Value(wi_Row, "RrXwtPer") = !RrXWtPer
      'Ec.2 Copying CstRtInc, SalRtInc only for Rmctg=X
      If ADC("wRrTcTyp") = "ILC" Or (ADC("wRrTcTyp") = "RM" And ADC("wRrCtg") = "X") Then
        GrdRm.Value(wi_Row, "RrFixMinTol") = !RrFixMinTol
        GrdRm.Value(wi_Row, "RrFixMaxTol") = !RrFixMaxTol
      End If
      GrdRm.Value(wi_Row, "RrRmKT") = !RrRmKT

      '*** Jay 3.01 *** [ILC]
      GrdRm.SaveRec (wi_Row)
      .MoveNext
    Loop
  End With
  MWLib.EndProcess Me
  GrdRm.EndCopy
  Set wRsCustRmRt = Nothing
  '*** End of the copy process
End Sub
Private Function ColSty(ByVal mColSty As String) As en_fgColStyle
  '*** Sets the Col Style property for the fields in the Grid
  Select Case mColSty
    Case Is = "N"
      ColSty = en_fgColStyle.fgcNormal
    Case Is = "I"
      ColSty = en_fgColStyle.fgcinvisible
    Case Is = "R"
      ColSty = en_fgColStyle.fgcReadOnly
  End Select
End Function
Private Function DataTyp(ByVal mDataTyp As String) As en_FldDataType
  '*** Sets the datatype property for the fields in the Grid
  Select Case mDataTyp
    Case Is = "C"
      DataTyp = fdtCharType
    Case Is = "N"
      DataTyp = fdtNumber
    Case Is = "F"
      DataTyp = fdtFloat
    Case Is = "D"
      DataTyp = fdtDateType
    Case Is = "U"
      DataTyp = fdtConvertToUpper
    Case Is = "T"
      DataTyp = fdtTimeType
  End Select
End Function
Private Sub GrdPropSet()
  '*** Sets all col Properties for the grid based on the TcStr table
  Dim wRsTc As MwfLib.MDORowSet
  Set wRsTc = moCn.OpenRes(" Select * from TcStr where TsTbl='RmRt' and " + _
                            " TsTyp='" + ADC("wRrTcTyp") + "' and " + _
                            " TsSTyp='" + ADC("wRrCtg") + "' ")
  With GrdRm
    Do While Not (wRsTc.EOF Or wRsTc.BOF)
      .ColProp(wRsTc!TsIdName).Style = ColSty(wRsTc!TsStyle)
      .ColProp(wRsTc!TsIdName).Heading = wRsTc!TsLbl
      
      'Sachin 3.11.0 - Tolerance by Pct for 'D' and 'C'
      If (ADC("wRrTcTyp") = "WTT" Or ADC("wRrTcTyp") = "QTT") And _
            InStr(1, "," + IIF(ADC("wRrTcTyp") = "WTT", ms_TolWtPctDC, ms_TolQtyPctDC) + ",", "," + UCase(ADC("wRrCtg")) + ",", vbTextCompare) > 0 Then
        If UCase(wRsTc!TsIdName) = UCase("RrFixMaxTol") Or UCase(wRsTc!TsIdName) = UCase("RrFixMinTol") Or _
            UCase(wRsTc!TsIdName) = UCase("RrOpnMaxTol") Or UCase(wRsTc!TsIdName) = UCase("RrOpnMinTol") Then _
          .ColProp(wRsTc!TsIdName).Heading = Replace(wRsTc!TsLbl, "Tol", "%")
      End If
      
      .ColProp(wRsTc!TsIdName).DataType = DataTyp(wRsTc!TsDataTyp)
      .ColProp(wRsTc!TsIdName).Mask = wRsTc!TsMask
      .ColProp(wRsTc!TsIdName).MaxLength = wRsTc!TsMaxLen
      wRsTc.MoveNext
    Loop
    '.SetColWidthToHd
  End With
  Set wRsTc = Nothing
End Sub
Private Sub CmpString(ByVal mTyp As String)
'*** Sets the CmpStr properties for the fields in the Key Frame and the Copy Frame

  '****** Seting values for CmpStr of 'K'ey fields, for Adc.RepCnd (Making Copy fields CmpStr = '')
  If mTyp = "K" Then
  If UCase(ms_MenuCd) = UCase("RmRt") Then
    ADC("wRrTcTyp").CmpStr = "RrTcTyp= "
    ADC("wRrCtg").CmpStr = "RrCtg= "
    ADC("wRrCmCtg").CmpStr = "RrCmCtg= "
    ADC("wRrSCtgFr").CmpStr = "RrSCtg>= "
    ADC("wRrSCtgTo").CmpStr = "RrSCtg<= "
    ADC("wRrCdFr").CmpStr = "RrCd>= "
    ADC("wRrCdTo").CmpStr = "RrCd<= "
    ADC("wRrCmCdFr").CmpStr = "RrCmCd>= "
    ADC("wRrCmCdTo").CmpStr = "RrCmCd<= "
    ADC("wCpySCtgFr").CmpStr = ""
    ADC("wCpySCtgTo").CmpStr = ""
    ADC("wCpyRmCdFr").CmpStr = ""
    ADC("wCpyRmCdTo").CmpStr = ""
    ' Zubin 213
    ADC("wRrCmCurCdFr").CmpStr = "RrCmCurCd>= "
    ADC("wRrCmCurCdTo").CmpStr = "RrCmCurCd<= "
    ' Zubin 213
    ' *** Manali 3.03 - Fr Sz - To Sz
    ADC("wRrSzFr").CmpStr = "RrFrLn>= "
    ADC("wRrSzTo").CmpStr = "RrFrLn<= "
    ' *** Manali 3.03 - Fr Sz - To Sz
    ' Sachin 3.02.0
    ADC("wRrPrtKey").CmpStr = "RrPrtKey= "
    ADC("wRrSTWGrpFr").CmpStr = "RrSTWGrp>= "
    ADC("wRrSTWGrpTo").CmpStr = "RrSTWGrp<= "
    ADC("wRrDmCdFr").CmpStr = "RrDmCd>= "  'sv.23 Fr Dsgcd added
    ADC("wRrDmCdTo").CmpStr = "RrDmCd<= "  'sv.23 To Dsgcd added
    
    ElseIf UCase(ms_MenuCd) = UCase("RmRtHist") Then
    '*************************** ZUBIN **************************
    ' 01st Dec 2003, EMR206
      ADC("wRhYyyyMm").CmpStr = "RhYyyyMm= "
      ADC("wRrTcTyp").CmpStr = "RhTcTyp= "
      ADC("wRrCtg").CmpStr = "RhCtg= "
      ADC("wRrCmCtg").CmpStr = "RhCmCtg= "
      ADC("wRrSCtgFr").CmpStr = "RhSCtg>= "
      ADC("wRrSCtgTo").CmpStr = "RhSCtg<= "
      ADC("wRrCdFr").CmpStr = "RhCd>= "
      ADC("wRrCdTo").CmpStr = "RhCd<= "
      ADC("wRrCmCdFr").CmpStr = "RhCmCd>= "
      ADC("wRrCmCdTo").CmpStr = "RhCmCd<= "

    '*************************** ZUBIN **************************
    ' Zubin 213
    ADC("wRrCmCurCdFr").CmpStr = "RhCmCurCd>= "
    ADC("wRrCmCurCdTo").CmpStr = "RhCmCurCd<= "
    ' Zubin 213
    ' *** Manali 3.03 - Fr Sz - To Sz
    ADC("wRrSzFr").CmpStr = "RhFrLn>= "
    ADC("wRrSzTo").CmpStr = "RhFrLn<= "
    ADC("wRrSTWGrpFr").CmpStr = "RhSTWGrp>= "
    ADC("wRrSTWGrpTo").CmpStr = "RhSTWGrp<= "
    ' *** Manali 3.03 - Fr Sz - To Sz
    ' Sachin 3.02.0
    ADC("wRrPrtKey").CmpStr = "RrPrtKey= "
    ADC("wRrDmCdFr").CmpStr = ""  'sv.23 Fr Dsgcd added
    ADC("wRrDmCdTo").CmpStr = ""  'sv.23 TO Dsgcd added
    
   End If
  '****** Seting values for CmpStr of 'C'opy fields, for Adc.RepCnd (Making Key fields CmpStr = '')
  ElseIf mTyp = "C" Then
    If UCase(ms_MenuCd) = UCase("RmRt") Then
    ADC("wRrTcTyp").CmpStr = ""
    ADC("wRrCtg").CmpStr = ""
    ADC("wRrCmCtg").CmpStr = ""
    ADC("wRrSCtgFr").CmpStr = ""
    ADC("wRrSCtgTo").CmpStr = ""
    ADC("wRrCdFr").CmpStr = ""
    ADC("wRrCdTo").CmpStr = ""
    ADC("wRrCmCdFr").CmpStr = ""
    ADC("wRrCmCdTo").CmpStr = ""
    ADC("wCpySCtgFr").CmpStr = "RrSCtg>= "
    ADC("wCpySCtgTo").CmpStr = "RrSCtg<= "
    ADC("wCpyRmCdFr").CmpStr = "RrCd>= "
    ADC("wCpyRmCdTo").CmpStr = "RrCd<= "
    ' Sachin 3.02.0
    ADC("wRrPrtKey").CmpStr = "RrPrtKey= "

    ADC("wCpyRmSzFr").CmpStr = "RrFrLn>= "
    ADC("wCpyRmSzTo").CmpStr = "RrFrLn<= "
    ADC("wRrSTWGrpFr").CmpStr = ""
    ADC("wRrSTWGrpTo").CmpStr = ""
    
    'sv.23 wRrCmCurCdFr,wRrCmCurCdTo,wRrSzFr,wRrSzTo copied from below (after endif)
    'wRrSTWGrpFr,wRrSTWGrpTo deleted. it is exist above with same condition
    'these 3 Fr & To was common for RmRt,RmRtHist. now cmpstring applied to RmRtHist, so that place changed
    ' Zubin 213
    ADC("wRrCmCurCdFr").CmpStr = ""
    ADC("wRrCmCurCdTo").CmpStr = ""
    ' Zubin 213
    ' *** Manali 3.03 - Fr Sz - To Sz
    ADC("wRrSzFr").CmpStr = ""
    ADC("wRrSzTo").CmpStr = ""
    ' *** Manali 3.03 - Fr Sz - To Sz
    
   '*************************** ZUBIN **************************
    ' 01st Dec 2003, EMR206
    ElseIf UCase(ms_MenuCd) = UCase("RmRtHist") Then
      ADC("wRhYyyyMm").CmpStr = ""          'sv.23 RhYyyyMm= removed
      ADC("wRrTcTyp").CmpStr = "RrTcTyp= "  'sv.23 RhTcTyp changed as RrTcTyp
      ADC("wRrCtg").CmpStr = "RrCtg= "      'sv.23 RhCtg changed as RrCtg
      ADC("wRrCmCtg").CmpStr = "RrCmCtg= "  'sv.23 RhCmCtg changed as RrCmCtg
      ADC("wRrSCtgFr").CmpStr = "RrSCtg>= "  'sv.23 "" removed & value applied
      ADC("wRrSCtgTo").CmpStr = "RrSCtg<= "  'sv.23 "" removed & value applied
      ADC("wRrCdFr").CmpStr = "RrCd>= "  'sv.23 "" removed & value applied
      ADC("wRrCdTo").CmpStr = "RrCd<= "  'sv.23 "" removed & value applied
      ADC("wRrCmCdFr").CmpStr = "RrCmCd>= "  'sv.23 "" removed & value applied
      ADC("wRrCmCdTo").CmpStr = "RrCmCd<= "  'sv.23 "" removed & value applied
      
    '*************************** ZUBIN **************************
      ' Sachin 3.02.0
      ADC("wRrPrtKey").CmpStr = "RrPrtKey= " 'sv.23 RhPrtKey changed as RrPrtKey
      
      'sv.23 wRrCmCurCdFr,wRrCmCurCdTo,wRrSzFr,wRrSzTo,wRrSTWGrpFr,wRrSTWGrpTo copied from below (after endif)
      'these 3 Fr & To was common for RmRt,RmRtHist. now cmpstring applied to RmRtHist, so that place changed
      ADC("wRrCmCurCdFr").CmpStr = "RrCmCurCd>= "
      ADC("wRrCmCurCdTo").CmpStr = "RrCmCurCd<= "
      ADC("wRrSzFr").CmpStr = "RrFrLn>= "
      ADC("wRrSzTo").CmpStr = "RrFrLn<= "
      ADC("wRrSTWGrpFr").CmpStr = ""
      ADC("wRrSTWGrpTo").CmpStr = ""
                
    End If
    'sv.23 Fr & To Dsgcd added
    ADC("wRrDmCdFr").CmpStr = ""
    ADC("wRrDmCdTo").CmpStr = ""
    
  End If
End Sub
Private Sub LnFrToPropSet(ByVal mColName As String)
  Dim ws_Style As String, ws_RtBy As String
  '*** The column properties for the columns RrFrLn, RrToLn are changed based on the RmCtg &
      'RmSCtg values
      'the settings change only for 'RM', 'INC' Type,
      'for 'LMG' FrLn is Lmg From,
      '     (as SCtg is blank, record wRtBy will not be found in param, setings will not change)
      'for 'STW' it is always Size1 irrrespective of RtBy.
    '*********Geeta*********Emr207*******For 'RSZ'
  If GrdRm.Row = 0 Then Exit Sub    '*** (zubin 2.13)
  If ADC("wRrTcTyp") = "RM" Or ADC("wRrTcTyp") = "INC" Or ADC("wRrTcTyp") = "RSZ" Then
    With GrdRm
    '*********Geeta*********Emr207*******
     Select Case ADC("wRrTcTyp")
      Case Is = UCase("RM"), UCase("INC")
        '*********************Geeta****************emr206*********************
         ws_RtBy = moCn.GetFldVal("Select (case when RrTcTyp='PTR' then 'P' else RrTcTyp end) from RmRt where RrTcTyp in ('PTR' ,'BS') " + _
                                  "and  RrCmCtg= '" + ADC("wRRCmCtg") + "' and  RrCmCd ='" + .Value(.Row, "RrCmCd") + "' " + _
                                  " and RrCmCurCd='' and  RrCtg='" + ADC("wRrCtg") + "' and  " + _
                                  "RrSctg= '" + .Value(.Row, "RrSCtg") + "' and " + _
                                  "RrSTWGrp='' and  RrCd='' and RrDmCtg ='' and  RrLsCtg ='' and  " + _
                                  "RrLabMCd='' and RrFrLn=0 and RrToLn=0 ")
        '*********************Geeta****************emr206*********************
        If ws_RtBy = "" Then _
         ws_RtBy = moCn.GetFldVal("Select PValue from Param where PTyp='RMSCTG' and " + _
                             "PMCd='" + ADC("wRrCtg") + "' and PSCd='" + .Value(.Row, "RrSCtg") + "' ")
      Case Is = UCase("RSZ")
        If (moCn.GetFldVal("Select hAvgSRtYn from Head where HCoCd= '" + gs_CoCd + "' and " + _
                       "HCd='" + ctSelfCmCd + "'") = "Y") Or _
         (moCn.GetFldVal("Select hSizeStkRt from Head where HCoCd= '" + gs_CoCd + "' and " + _
                       "HCd='" + ctSelfCmCd + "'") = "Y") Then _
           ws_RtBy = IIF(moCn.GetFldVal("Select PValue3 from Param where PTyp='RMSCTG'  " + _
                                    "and PMCd='" + ADC("wRrCtg") + "' and PSCd='" + .Value(.Row, "RrSCtg") + "'") = "N", "S", "P")
      End Select
     '*********Geeta*********Emr207*******
      Select Case mColName
      Case Is = UCase("RrFrLn")
        ws_Style = moCn.GetFldVal("Select TsStyle from TcStr where TsTbl='RmRt' and " + _
                               "TsTyp='" + ADC("wRrTcTyp") + "' and " + _
                               "TsSTyp='" + ADC("wRrCtg") + "' and " + _
                               "TsIdName='RrFrLn'")
        If ws_Style = "N" Then
          Select Case ws_RtBy
          Case Is = "S"
              .ColProp("RrFrLn").Style = ColSty("N")
              '****** (Jen 2.13) 'X' added ******
              ' ****** Manali 3.03 - 18/06/08 - RmCtg 'L' included
              If (ADC("wRrCtg") = "G") Or (ADC("wRrCtg") = "P") Or (ADC("wRrCtg") = "S") Or (ADC("wRrCtg") = "L") Or _
                 (ADC("wRrCtg") = "A") Or (ADC("wRrCtg") = "X") Then .ColProp("RrFrLn").Heading = "Fr Size "  ' sachin 2.12 - Silver added
                            
              '*** (Bef 2.13)
              'If (ADC("wRrCtg") = "D") Or (ADC("wRrCtg") = "C") Or (ADC("wRrCtg") = "X") Then _
              '.ColProp("RrFrLn").Heading = "Fr Size"
              '*** (Bef 2.13)
              
              '*** (Jen 2.13)
              If (ADC("wRrCtg") = "D") Or (ADC("wRrCtg") = "C") Then .ColProp("RrFrLn").Heading = "Fr Size"
              '*** (Jen 2.13)
              
          Case Is = "-"
              .ColProp("RrFrLn").Style = ColSty("R")
          Case Is = "BS"
              .ColProp("RrFrLn").Style = ColSty("N")
              'urmi 2.04-3 ColHeading Changed From FrLn to Ln1
              '.ColProp("RrFrLn").Heading = "Fr Ln  "
              .ColProp("RrFrLn").Heading = "Ln1    "
          Case Is = "P"
              .ColProp("RrFrLn").Style = ColSty("N")
              .ColProp("RrFrLn").Heading = "Fr Ptr "
          End Select
        End If
      Case Is = UCase("RrXWtPer")
              If (ADC("wRrCtg") = "G" Or ADC("wRrCtg") = "P" Or ADC("wRrCtg") = "S" Or ADC("wRrCtg") = "L") Then .ColProp("RrFrLn").Heading = "LM" + ADC("wRrCtg")
              
              If (ADC("wRrCtg") = "X") Then
              
              .ColProp("RrFrLn").Heading = "LM" + moCn.GetFldVal("Select RmCtg from RmMst join RmMst ForParent on ForParent.RmCd = RmMst.RmBaseCd" + _
                                "where RmCd= '" + .Value(.Row, "RrCd") + "'" + IIF(gs_Partition = ctCurrPrtn, " And RmPrtKey='" + ctCurrPrtn + "' ", ""))
              
              End If
      Case Is = UCase("RrToLn")
        ws_Style = moCn.GetFldVal("Select TsStyle from TcStr where TsTbl='RmRt' and " + _
                               "TsTyp='" + ADC("wRrTcTyp") + "' and " + _
                               "TsSTyp='" + ADC("wRrCtg") + "' and " + _
                               "TsIdName='RrToLn'")
        If ws_Style = "N" Then
          Select Case ws_RtBy
          Case Is = "S"
              .ColProp("RrToLn").Style = ColSty("N")
              .ColProp("RrToLn").Heading = "To Size"
              '*** (Jen 2.13)
              
          Case Is = "-"
              .ColProp("RrToLn").Style = ColSty("R")
          
          Case Is = "BS"
              .ColProp("RrToLn").Style = ColSty("N")
              'urmi 2.04-3 ColHeading Changed From ToLn to Ln2
              '.ColProp("RrToLn").Heading = "To Ln  "
              .ColProp("RrToLn").Heading = "Ln2    "
          Case Is = "P"
              .ColProp("RrToLn").Style = ColSty("N")
              .ColProp("RrToLn").Heading = "To Ptr "
          End Select
        End If
      '*****Geeta********Emr207
      Case Is = UCase("RrRngSz")
        ws_Style = moCn.GetFldVal("Select TsStyle from TcStr where TsTbl='RmRt' and " + _
                               "TsTyp='" + ADC("wRrTcTyp") + "' and " + _
                               "TsSTyp='" + ADC("wRrCtg") + "' and " + _
                               "TsIdName='RrRngSz'")
        If ws_Style = "N" Then
          Select Case ws_RtBy
          Case Is = "S"
              .ColProp("RrRngSz").Style = ColSty("N")
              .ColProp("RrRngSz").Heading = "Range Size"
          Case Is = "P"
              .ColProp("RrRngSz").Style = ColSty("N")
              .ColProp("RrRngSz").Heading = "Range Ptr "
          End Select
        End If
        '*****Geeta********Emr207
        
        
      End Select
    End With
  End If
End Sub
Private Sub DispFra(ByVal pv_RmRtFra As en_RmRtFra)
  '*** This is the common procedure which is used to toggle between hiding and displaying
      'the frames (i.e. if the frame is visible then make it invisible and viseversa)
      'associated with a command button
  
  Select Case pv_RmRtFra
  Case Is = Cpy
    If FraCpy.Visible = True Then
      FraCpy.Visible = False
      FraCpy.Enabled = False
      Call EnaDisaCmds(False, CmdCpy)
      CmdCpy.SetFocus
    Else
      FraCpy.Visible = True
      FraCpy.Enabled = True
      ADC("wCpyCustCtgFr").SetFocus
      FraCpy.ZOrder
      Call EnaDisaCmds(True, CmdCpy)
    End If
  Case Is = Pwd
    If FraPwd.Visible = True Then
      FraPwd.Visible = False
      FraPwd.Enabled = False
      Call EnaDisaCmds(False, CmdPwd)
      CmdPwd.SetFocus
    Else
      FraPwd.Visible = True
      FraPwd.Enabled = True
      TxtPwd.SetFocus
      FraPwd.ZOrder
      Call EnaDisaCmds(True, CmdPwd)
    End If
  Case Is = ILC
    If FrmRefRates.Visible = True Then
      FrmRefRates.Visible = False
      FrmRefRates.Enabled = False
      Call EnaDisaCmds(False, CmdILC)
      CmdILC.SetFocus
    Else
      FrmRefRates.Visible = True
      FrmRefRates.Enabled = True
      CmdILCGo.SetFocus
      FrmRefRates.ZOrder
      Call EnaDisaCmds(True, CmdILC)
    End If
  Case Is = CpyHist
    If FraCpyHist.Visible = True Then
      FraCpyHist.Visible = False: FraCpyHist.Enabled = False
      Call EnaDisaCmds(False, CmdCpyHist): CmdCpyHist.SetFocus
    Else
      FraCpyHist.Visible = True: FraCpyHist.Enabled = True
      CmdCpyHistGo.SetFocus
      FraCpyHist.ZOrder
      Call EnaDisaCmds(True, CmdCpyHist)
    End If
    '*********Bhavna - import from excel
    Case Is = ImpExc
    If FraImpExc.Visible = True Then
      FraImpExc.Visible = False
      FraImpExc.Enabled = False
      Call EnaDisaCmds(False, CmdImpExc)
      CmdImpExc.SetFocus
    Else
      FraImpExc.Visible = True
      FraImpExc.Enabled = True
      CmdImpExcPath.SetFocus
      FraImpExc.ZOrder
      Call EnaDisaCmds(True, CmdImpExc)
    End If
    '*********Bhavna - import from excel
    'Uni.24- if Delete rates frame is visible then will be invisible and viseversa
    Case Is = DelRt
    If FraDelRt.Visible = True Then
      FraDelRt.Visible = False
      FraDelRt.Enabled = False
      Call EnaDisaCmds(False, CmdDelRt)
      CmdDelRt.SetFocus
    Else
      FraDelRt.Visible = True
      FraDelRt.Enabled = True
      CmdDelRtGo.SetFocus
      FraDelRt.ZOrder
      Call EnaDisaCmds(True, CmdDelRt)
    End If
  End Select
End Sub

Private Sub EnaDisaCmds(ByVal pv_ShowFra As Boolean, Optional ByVal pv_CmdBut As MWCTL_BTN1)
  '*** This procedure is used to give the toggle option (enable and disable) to
      'the command buttons
  
  '*** (Jen 3.01)
  If pv_ShowFra = True Then
    Call StoreState
  End If
  '*** (Jen 3.01)
  
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
    End If
  End If
  FraNKeyAll.Enabled = Not pv_ShowFra
  ADC.AllowSave = Not pv_ShowFra
  
  If UCase(ms_MenuCd) = UCase("RmRt") Then
    If pv_ShowFra = True Then
      '** All command buttons are disabled
      CmdCpy.Enabled = Not pv_ShowFra
      CmdILC.Enabled = Not pv_ShowFra   '*** Jay 3.01 *** [ILC]
      CmdPwd.Enabled = Not pv_ShowFra   '*** Jay 3.01 *** [ILC]
      CmdImpExc.Enabled = Not pv_ShowFra  '********Bhavna - Imprt from Excel
      CmdDelRt.Enabled = Not pv_ShowFra 'Uni.24- Delete Rates
    Else
      '*** Restoring Ena/Disa state of all Command Buttons
      CmdCpy.Enabled = mArr_Ed(Cpy)
      CmdILC.Enabled = mArr_Ed(ILC)   '*** Jay 3.01 *** [ILC]
      CmdPwd.Enabled = mArr_Ed(Pwd)  '*** Jay 3.01 *** [ILC]
      CmdImpExc.Enabled = mArr_Ed(ImpExc)    ' ***** Bhavna - Import from Excel
      CmdDelRt.Enabled = mArr_Ed(DelRt) 'Uni.24- Delete rates
    End If
  ElseIf UCase(ms_MenuCd) = UCase("RmRtHist") Then
    If pv_ShowFra = True Then
      CmdCpyHist.Enabled = Not pv_ShowFra
      'uni.186 CmdImpExc,CmdDelRt buttons available before user does Find,
      'it will be enable only after find
      CmdImpExc.Enabled = Not pv_ShowFra
      CmdDelRt.Enabled = Not pv_ShowFra
    Else
      CmdCpyHist.Enabled = mArr_Ed(CpyHist)
      'uni.186 CmdImpExc,CmdDelRt buttons available before user does Find,
      'it will be enable only after find
      CmdImpExc.Enabled = mArr_Ed(ImpExc)
      CmdDelRt.Enabled = mArr_Ed(DelRt)
    End If
  End If
  
  If Not pv_CmdBut Is Nothing Then pv_CmdBut.Enabled = True
End Sub
Private Sub StoreState()
  '*** This procedure is used to store the Enabled/Disabled
      'states of all the Command Buttons
  '*** This is used for restoring the original states in the future
  If UCase(ms_MenuCd) = UCase("RmRt") Then
    mArr_Ed(Cpy) = CmdCpy.Enabled
    mArr_Ed(ILC) = CmdILC.Enabled
    mArr_Ed(Pwd) = CmdPwd.Enabled
    mArr_Ed(ImpExc) = CmdImpExc.Enabled     ' ***** Bhavna - Import From Excel
    mArr_Ed(DelRt) = CmdDelRt.Enabled     'Uni.24- Delete Rates
  ElseIf UCase(ms_MenuCd) = UCase("RmRtHist") Then
    mArr_Ed(CpyHist) = CmdCpyHist.Enabled
    'uni.186 CmdImpExc,CmdDelRt buttons available before user does Find,
    'it will be enable only after find
    mArr_Ed(ImpExc) = CmdImpExc.Enabled
    mArr_Ed(DelRt) = CmdDelRt.Enabled


  End If
End Sub

Private Sub HideAllFras()
  '*** This procedure hides all the frames other than FraNKeyAll
  FraNKeyAll.Visible = True
  FraCpy.Visible = False
  FraCpy.Enabled = False
  '*** Jay 3.01 *** [ILC]
  FraPwd.Visible = False
  FraPwd.Enabled = False
  '*** Jay 3.01 *** [ILC]
  
  ' ***** Manali 3.5.0 - 02/06/09 - [Earlier CopyGo Not Working]
  FraCpyHist.Visible = False
  FraCpyHist.Enabled = False
    
  FrmRefRates.Enabled = False
  FrmRefRates.Visible = False
  ' ***** Manali 3.5.0 - 02/06/09
  
   ' ***** Bhavna - Import from excel
  FraImpExc.Visible = False
  FraImpExc.Enabled = False
  ' *****  bhavna
  FraDelRt.Visible = False  'Uni.24- Delete rates
  FraDelRt.Enabled = False  'Uni.24- Delete rates
 
End Sub
'**************Bhavna for import RmRt from excel
Private Sub ImportExcel()
On Error GoTo ErrHdlr
    
    Dim wo_rsExcel As New MDORowSet, i As Integer
    
    Dim wRrCmCd As String, wRrCmCurCd As String, wRrSCtg As String, wQWCst As String, wQWSale As String
    Dim wRrCd As String, wRmCtg As String, wRrFrLn As Double, wRrToLn As Double
    Dim wRrCstRt As Double, wRrXWtPer As Double, wRrSalRt As Double, ws_BlnkYN As String, wRrCstRtInc As Double, wRrSalRtInc As Double

    '****** Checking For Valid Inputs From the User ******
    ' ***** Manali 3.9.0 - modified to accept Excel 2007 (.xlsx) files
    If Not OpenExcelFile(ADC("WXLFILE")) Then DispMsg "Improper File Format", etError: Exit Sub
    
    'Call OpenExcelFile(adc("WXLFILE"))
    If ChkSheet(ADC("WXLSHTNO")) = False Then DispMsg "This WorkSheet Number Does Not Exist in the specified Excel File", etError: CloseExcel: Exit Sub
    '****** Checking For Valid Inputs From the User ******

    ' ***** Manali 3.9.1 - 21/02/12 - qSr added for Order By according to data entered in Excel
    'ec.2-CstRtInc, SalRtInc added
    moCn.CreateTmpTable "#TmpXls", "Select 0 as qSr, RrCmCd, RrCmCurCd, RrSCtg, RrCd, RrFrLn, RrToLn, RrCstRt, RrSalRt, RrFixMinTol, RrFixMaxTol, RrXWtPer Into #TmpXls From RmRt Where 1=2"
        
    ADC("WXLFRROW") = IIF(ADC("WXLFRROW") <= 1, 2, ADC("WXLFRROW"))
    For i = ADC("WXLFRROW") To ADC("WXLTOROW")
        '************validate cust code
        wRrCmCd = Trim(ReadCell(ADC("WXLSHTNO"), i, 1))
        ws_BlnkYN = moCn.GetFldVal("Select TsBlank from TcStr where TsTbl='RmRt' and " + _
                            "TsTyp='" + ADC("wRrTcTyp") + "' and " + _
                            "TsSTyp='" + ADC("wRrCtg") + "' and " + _
                            "TsIdName='RrCmCd' ")
        If ws_BlnkYN = "N" And wRrCmCd = "" Then _
                DispMsg "Cannot Copy as Customer Code cannot be blank [Row :" + CStr(i) + "]", etError: CloseExcel: Exit Sub
                
        If (wRrCmCd <> "") And (Not moCn.RecSeek("Select CmCd From CustMst where CmCtg= '" + ADC("wRrCmCtg") + "' " + _
                             "and CmCd = '" + wRrCmCd + "'")) Then _
                DispMsg "Cannot Copy as Invalid Customer Code [Row :" + CStr(i) + "]", etError: CloseExcel: Exit Sub
        
        'Indigo.1- Checking user access permission for the given customer
        If UCase(ms_MenuCd) = UCase("RmRt") Then
          Dim wErrMsg As String
          If Not ChkCmpUsrAccess(ADC("WRRCMCTG"), ADC.MenuCd, "", "", wRrCmCd, wErrMsg) Then _
             DispMsg "You don't have permissions to add this Customer [Row :" + CStr(i) + "]", etError: CloseExcel: Exit Sub
        End If
      
      '********Validate cur cd
        wRrCmCurCd = Trim(ReadCell(ADC("WXLSHTNO"), i, 2))
        ws_BlnkYN = moCn.GetFldVal("Select TsBlank from TcStr where TsTbl='RmRt' and " + _
                            "TsTyp='" + ADC("wRrTcTyp") + "' and " + _
                            "TsSTyp='" + ADC("wRrCtg") + "' and " + _
                            "TsIdName='RrCmCurCd' ")
        If ws_BlnkYN = "N" And wRrCmCurCd = "" Then _
                DispMsg "Cannot Copy as Customer Currency Code cannot be blank [Row :" + CStr(i) + "]", etError: CloseExcel: Exit Sub
                
        If (wRrCmCurCd <> "") And (Not moCn.RecSeek("Select PMCd From Param where  " + _
                           "PTyp= 'CURNCY' and PMCd = '" + wRrCmCurCd + "' and " + _
                           "PSCd = '' ")) Then _
                DispMsg "Cannot Copy as Invalid Customer Currency Code [Row :" + CStr(i) + "]", etError: CloseExcel: Exit Sub
      
        '**********Valid sub ctg
        wRrSCtg = Trim(ReadCell(ADC("WXLSHTNO"), i, 3))
        If (wRrSCtg <> "") And (Not moCn.RecSeek("Select PSCd From Param where PTyp= 'RMSCTG' " + _
                             "and PMCd = '" + ADC("wRrCtg") + "' and " + _
                             "PSCd = '" + wRrSCtg + "'")) Then _
                DispMsg "Cannot Copy as Invalid Rm Sub Ctg Code [Row :" + CStr(i) + "]", etError: CloseExcel: Exit Sub
      
        '****** Valid RmCode
        wRrCd = Trim(ReadCell(ADC("WXLSHTNO"), i, 4))
        If Not moCn.RecSeek("Select 'x' from RmMst Where RmCd='" + wRrCd + "' And RmPrtKey='" + ctCurrPrtn + "' and RmValidYN='Y' and RmZ='N' ") Then _
          DispMsg "Cannot Import as RmCode is Not Valid [Row :" + CStr(i) + "]", etError: CloseExcel: Exit Sub
    
        wRmCtg = moCn.GetFldVal("Select RmCtg from RmMst Where RmCd='" + wRrCd + "' ")
        If wRmCtg <> ADC("wRrCtg") Then _
            DispMsg "Cannot Import as RmCtg is Not same as selected in X (D/C) [Row :" + CStr(i) + "]", etError: CloseExcel: Exit Sub
        
        '****** Checking for valid RmCd Size in the Excel File
        'Debug.Print Val(ReadCell(ADC("WXLSHTNO"), i, 3))
        If ReadCell(ADC("WXLSHTNO"), i, 5) <> "" And Not IsNumeric(ReadCell(ADC("WXLSHTNO"), i, 5)) Then
            DispMsg "Cannot Copy as From Rm Size Value is not Number [Row :" + CStr(i) + "]", etError: CloseExcel: Exit Sub
        End If
        If ReadCell(ADC("WXLSHTNO"), i, 6) <> "" And Not IsNumeric(ReadCell(ADC("WXLSHTNO"), i, 6)) Then
            DispMsg "Cannot Copy as To RmSize Value is not Number [Row :" + CStr(i) + "]", etError: CloseExcel: Exit Sub
        End If
        If ReadCell(ADC("WXLSHTNO"), i, 8) <> "" And Not IsNumeric(ReadCell(ADC("WXLSHTNO"), i, 8)) Then
            DispMsg "Cannot Copy as Cost Rate Value is not Number [Row :" + CStr(i) + "]", etError: CloseExcel: Exit Sub
        End If
        If ReadCell(ADC("WXLSHTNO"), i, 10) <> "" And Not IsNumeric(ReadCell(ADC("WXLSHTNO"), i, 10)) Then
            DispMsg "Cannot Copy as Sale Rate Value is not Number [Row :" + CStr(i) + "]", etError: CloseExcel: Exit Sub
        End If
        'ec.2- CstRtInc, SalRtInc validation is added
        If ReadCell(ADC("WXLSHTNO"), i, 11) <> "" And Not IsNumeric(ReadCell(ADC("WXLSHTNO"), i, 11)) Then
            DispMsg "Cannot Copy as LMG is not Number [Row :" + CStr(i) + "]", etError: CloseExcel: Exit Sub
        End If
        
        
        wRrFrLn = Val(ReadCell(ADC("WXLSHTNO"), i, 5))
        wRrToLn = Val(ReadCell(ADC("WXLSHTNO"), i, 6))
        wRrCstRt = Val(ReadCell(ADC("WXLSHTNO"), i, 8))
        wRrSalRt = Val(ReadCell(ADC("WXLSHTNO"), i, 10))
        wRrCstRtInc = 0
        wRrSalRtInc = 0
        wRrXWtPer = Val(ReadCell(ADC("WXLSHTNO"), i, 11))
'        wRrSalRtInc = Val(ReadCell(ADC("WXLSHTNO"), i, 12))
                
        'IG.189- Calling function and Validating Rm From & To size
        If RmSzLen(GrdRm.ColProp("RRFRLN").Mask, wRrFrLn) Then _
            DispMsg "Cannot Copy as Invalid From Rm Size Entered, [Row :" + CStr(i) + "]", etError: CloseExcel: Exit Sub
        
        If RmSzLen(GrdRm.ColProp("RRTOLN").Mask, wRrToLn) Then _
            DispMsg "Cannot Copy as Invalid To Rm Size Entered, [Row :" + CStr(i) + "]", etError: CloseExcel: Exit Sub
             
       
        If wRrCstRt < 0 Or wRrCstRt > 9999999.999 Then _
            DispMsg "Cannot Copy as Invalid Cost Rate Entered, Should be Between 0 and 9999999.999 [Row :" + CStr(i) + "]", etError: CloseExcel: Exit Sub
        If wRrSalRt < 0 Or wRrSalRt > 9999999.999 Then _
            DispMsg "Cannot Copy as Invalid Sale Rate Entered, Should be Between 0 and 9999999.999 [Row :" + CStr(i) + "]", etError: CloseExcel: Exit Sub
        'ec.2- CstRtInc, SalRtInc mask is added
        If wRrXWtPer < 0 Then _
            DispMsg "Cannot Copy as Invalid LMG rate, Should be between 0 and 999.9999 [Row :" + CStr(i) + "]", etError: CloseExcel: Exit Sub
        
        If wRrCstRtInc < 0 Or wRrCstRtInc > 999.9999 Then _
            DispMsg "Cannot Copy as Invalid Cost Rate Inc/$ Entered, Should be Between 0 and 999.9999 [Row :" + CStr(i) + "]", etError: CloseExcel: Exit Sub
        If wRrSalRtInc < 0 Or wRrSalRtInc > 999.9999 Then _
            DispMsg "Cannot Copy as Invalid Sale Rate Inc/$ Entered, Should be Between 0 and 999.9999 [Row :" + CStr(i) + "]", etError: CloseExcel: Exit Sub
        
        If ADC("wRrTcTyp") = "RM" And ADC("wRrCtg") = "X" And wRrCstRtInc > 0 Then
          If Not ms_CstRtEdtYN And wRrCmCd <> ctSelfCmCd And wRrCmCd <> ctValueCmCd Then _
            DispMsg "Cannot Copy as Invalid Cost Rate Inc/$ Entered, Should be 0  [Row :" + CStr(i) + "]", etError: CloseExcel: Exit Sub
        End If


              '************add record in tmp file
'       moCn.Execute ("Insert Into #TmpXls Values (" + CStr(i) + ", '" + wRrCmCd + "', '" + wRrCmCurCd + "', '" + wRrSCtg + "'," + _
                        "'" + wRrCd + "', " + Str(wRrFrLn) + ", " + Str(wRrToLn) + ", " + _
                        "'" + ReadCell(ADC("WXLSHTNO"), i, 7) + "', " + Str(wRrCstRt) + " , '" + Trim(ReadCell(ADC("WXLSHTNO"), i, 9)) + "', " + _
                        Str(wRrSalRt) + ")")
       moCn.Execute ("Insert Into #TmpXls Values (" + CStr(i) + ", '" + wRrCmCd + "', '" + wRrCmCurCd + "', '" + wRrSCtg + "'," + _
                        "'" + wRrCd + "', " + Str(wRrFrLn) + ", " + Str(wRrToLn) + ", " + _
                        Str(wRrCstRt) + " ,  " + Str(wRrSalRt) + ", " + Str(wRrCstRtInc) + "," + Str(wRrSalRtInc) + "," + CStr(wRrXWtPer) + ")")
                                
    Next

    CloseExcel
    ' ***** Manali 3.9.1 - 21/02/12 - qSr added for Order By according to data entered in Excel
    Set wo_rsExcel = moCn.OpenRes("Select * From #TmpXls Order By qSr")
    wo_rsExcel.MoveFirst
    
    Dim wRow As Double
    
    If wo_rsExcel.RecCount > 0 Then
      With GrdRm
        .StartCopy
        MWLib.BeginProcess Me, "Copying RM Details ..."
        Do While Not wo_rsExcel.EOF
          .AddItem
          wRow = .Rows - 1
          .Row = wRow
          
          .Value(wRow, "RrPrtKey") = ctCurrPrtn
          .Value(wRow, "RrCmCd") = wo_rsExcel!RrCmCd
          .Value(wRow, "RrCmCurCd") = wo_rsExcel!RrCmCurCd
          .Value(wRow, "RrSCtg") = wo_rsExcel!RrSCtg
          .Value(wRow, "RrCd") = wo_rsExcel!RrCd
          .Value(wRow, "RrFrLn") = wo_rsExcel!RrFrLn
          .Value(wRow, "RrToLn") = wo_rsExcel!RrToLn
          .Value(wRow, "RrCstRt") = wo_rsExcel!RrCstRt
          .Value(wRow, "RrSalRt") = wo_rsExcel!RrSalRt
          If ADC("wRrTcTyp") = "RM" And ADC("wRrCtg") = "X" Then
            .Value(wRow, "RrFixMinTol") = wo_rsExcel!RrFixMinTol
            .Value(wRow, "RrFixMaxTol") = wo_rsExcel!RrFixMaxTol
          End If
            .Value(wRow, "RrXWtPer") = wo_rsExcel!RrXWtPer
          
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

'**************Bhavna for import RmRt from excel
Private Sub ImportExcelIPD()
On Error GoTo ErrHdlr
    
    Dim wo_rsExcel As New MDORowSet, i As Integer
    
    Dim wRrCmCd As String, wRrCmCurCd As String, wRrSCtg As String, wQWCst As String, wQWSale As String
    Dim wRrCd As String, wRmCtg As String, wRrFrLn As Double, wRrToLn As Double
    Dim wRrCstRt As Double, wRrXWtPer As Double, wRrSalRt As Double, ws_BlnkYN As String, wRrCstRtInc As Double, wRrSalRtInc As Double

    '****** Checking For Valid Inputs From the User ******
    ' ***** Manali 3.9.0 - modified to accept Excel 2007 (.xlsx) files
    If Not OpenExcelFile(ADC("WXLFILE")) Then DispMsg "Improper File Format", etError: Exit Sub
    
    'Call OpenExcelFile(adc("WXLFILE"))
    If ChkSheet(ADC("WXLSHTNO")) = False Then DispMsg "This WorkSheet Number Does Not Exist in the specified Excel File", etError: CloseExcel: Exit Sub
    '****** Checking For Valid Inputs From the User ******

    ' ***** Manali 3.9.1 - 21/02/12 - qSr added for Order By according to data entered in Excel
    'ec.2-CstRtInc, SalRtInc added
    moCn.CreateTmpTable "#TmpXls", "Select 0 as qSr, RrCmCd, RrCmCurCd, RrSCtg, RrCd, RrFrLn, RrToLn, RrCstRt, RrSalRt, RrFixMinTol, RrFixMaxTol, RrXWtPer Into #TmpXls From RmRt Where 1=2"
        
    ADC("WXLFRROW") = IIF(ADC("WXLFRROW") <= 1, 2, ADC("WXLFRROW"))
    For i = ADC("WXLFRROW") To ADC("WXLTOROW")
        '************validate cust code
        wRrCmCd = Trim(ReadCell(ADC("WXLSHTNO"), i, 1))
        ws_BlnkYN = moCn.GetFldVal("Select TsBlank from TcStr where TsTbl='RmRt' and " + _
                            "TsTyp='" + ADC("wRrTcTyp") + "' and " + _
                            "TsSTyp='" + ADC("wRrCtg") + "' and " + _
                            "TsIdName='RrCmCd' ")
        If ws_BlnkYN = "N" And wRrCmCd = "" Then _
                DispMsg "Cannot Copy as Customer Code cannot be blank [Row :" + CStr(i) + "]", etError: CloseExcel: Exit Sub
                
        If (wRrCmCd <> "") And (Not moCn.RecSeek("Select CmCd From CustMst where CmCtg= '" + ADC("wRrCmCtg") + "' " + _
                             "and CmCd = '" + wRrCmCd + "'")) Then _
                DispMsg "Cannot Copy as Invalid Customer Code [Row :" + CStr(i) + "]", etError: CloseExcel: Exit Sub
        
        'Indigo.1- Checking user access permission for the given customer
        If UCase(ms_MenuCd) = UCase("RmRt") Then
          Dim wErrMsg As String
          If Not ChkCmpUsrAccess(ADC("WRRCMCTG"), ADC.MenuCd, "", "", wRrCmCd, wErrMsg) Then _
             DispMsg "You don't have permissions to add this Customer [Row :" + CStr(i) + "]", etError: CloseExcel: Exit Sub
        End If
      
      '********Validate cur cd
        wRrCmCurCd = Trim(ReadCell(ADC("WXLSHTNO"), i, 2))
        ws_BlnkYN = moCn.GetFldVal("Select TsBlank from TcStr where TsTbl='RmRt' and " + _
                            "TsTyp='" + ADC("wRrTcTyp") + "' and " + _
                            "TsSTyp='" + ADC("wRrCtg") + "' and " + _
                            "TsIdName='RrCmCurCd' ")
        If ws_BlnkYN = "N" And wRrCmCurCd = "" Then _
                DispMsg "Cannot Copy as Customer Currency Code cannot be blank [Row :" + CStr(i) + "]", etError: CloseExcel: Exit Sub
                
        If (wRrCmCurCd <> "") And (Not moCn.RecSeek("Select PMCd From Param where  " + _
                           "PTyp= 'CURNCY' and PMCd = '" + wRrCmCurCd + "' and " + _
                           "PSCd = '' ")) Then _
                DispMsg "Cannot Copy as Invalid Customer Currency Code [Row :" + CStr(i) + "]", etError: CloseExcel: Exit Sub
      
        '**********Valid sub ctg
        wRrSCtg = Trim(ReadCell(ADC("WXLSHTNO"), i, 3))
        If (wRrSCtg <> "") And (Not moCn.RecSeek("Select PSCd From Param where PTyp= 'RMSCTG' " + _
                             "and PMCd = '" + ADC("wRrCtg") + "' and " + _
                             "PSCd = '" + wRrSCtg + "'")) Then _
                DispMsg "Cannot Copy as Invalid Rm Sub Ctg Code [Row :" + CStr(i) + "]", etError: CloseExcel: Exit Sub
      
        '****** Valid RmCode
        wRrCd = Trim(ReadCell(ADC("WXLSHTNO"), i, 4))
        If Not moCn.RecSeek("Select 'x' from RmMst Where RmCd='" + wRrCd + "' And RmPrtKey='" + ctCurrPrtn + "' and RmValidYN='Y' and RmZ='N' ") Then _
          DispMsg "Cannot Import as RmCode is Not Valid [Row :" + CStr(i) + "]", etError: CloseExcel: Exit Sub
    
        wRmCtg = moCn.GetFldVal("Select RmCtg from RmMst Where RmCd='" + wRrCd + "' ")
        If wRmCtg <> ADC("wRrCtg") Then _
            DispMsg "Cannot Import as RmCtg is Not same as selected in header [Row :" + CStr(i) + "]", etError: CloseExcel: Exit Sub
        
        'ec.2- CstRtInc, SalRtInc validation is added
        If ReadCell(ADC("WXLSHTNO"), i, 5) <> "" And Not IsNumeric(ReadCell(ADC("WXLSHTNO"), i, 5)) Then
            DispMsg "Cannot Copy as CstRtInc is not Number [Row :" + CStr(i) + "]", etError: CloseExcel: Exit Sub
        End If
        If ReadCell(ADC("WXLSHTNO"), i, 6) <> "" And Not IsNumeric(ReadCell(ADC("WXLSHTNO"), i, 6)) Then
            DispMsg "Cannot Copy as SalRtInc is not Number [Row :" + CStr(i) + "]", etError: CloseExcel: Exit Sub
        End If
        
        
        wRrFrLn = 0
        wRrToLn = 999.9999
        wRrCstRt = 0
        wRrSalRt = 0
        wRrCstRtInc = Val(ReadCell(ADC("WXLSHTNO"), i, 5))
        wRrSalRtInc = Val(ReadCell(ADC("WXLSHTNO"), i, 6))
        wRrXWtPer = 0
        'IG.189- Calling function and Validating Rm From & To size
        If RmSzLen(GrdRm.ColProp("RRFRLN").Mask, wRrFrLn) Then _
            DispMsg "Cannot Copy as Invalid From Rm Size Entered, [Row :" + CStr(i) + "]", etError: CloseExcel: Exit Sub
        
        If RmSzLen(GrdRm.ColProp("RRTOLN").Mask, wRrToLn) Then _
            DispMsg "Cannot Copy as Invalid To Rm Size Entered, [Row :" + CStr(i) + "]", etError: CloseExcel: Exit Sub
             
       
        
        If wRrCstRtInc < 0 Or wRrCstRtInc > 999.9999 Then _
            DispMsg "Cannot Copy as Invalid Cost Rate Inc/$ Entered, Should be Between 0 and 999.9999 [Row :" + CStr(i) + "]", etError: CloseExcel: Exit Sub
        If wRrSalRtInc < 0 Or wRrSalRtInc > 999.9999 Then _
            DispMsg "Cannot Copy as Invalid Sale Rate Inc/$ Entered, Should be Between 0 and 999.9999 [Row :" + CStr(i) + "]", etError: CloseExcel: Exit Sub
        
        If ADC("wRrTcTyp") = "RM" And ADC("wRrCtg") = "X" And wRrCstRtInc > 0 Then
          If Not ms_CstRtEdtYN And wRrCmCd <> ctSelfCmCd And wRrCmCd <> ctValueCmCd Then _
            DispMsg "Cannot Copy as Invalid Cost Rate Inc/$ Entered, Should be 0  [Row :" + CStr(i) + "]", etError: CloseExcel: Exit Sub
        End If


              '************add record in tmp file
'       moCn.Execute ("Insert Into #TmpXls Values (" + CStr(i) + ", '" + wRrCmCd + "', '" + wRrCmCurCd + "', '" + wRrSCtg + "'," + _
                        "'" + wRrCd + "', " + Str(wRrFrLn) + ", " + Str(wRrToLn) + ", " + _
                        "'" + ReadCell(ADC("WXLSHTNO"), i, 7) + "', " + Str(wRrCstRt) + " , '" + Trim(ReadCell(ADC("WXLSHTNO"), i, 9)) + "', " + _
                        Str(wRrSalRt) + ")")
       moCn.Execute ("Insert Into #TmpXls Values (" + CStr(i) + ", '" + wRrCmCd + "', '" + wRrCmCurCd + "', '" + wRrSCtg + "'," + _
                        "'" + wRrCd + "', " + Str(wRrFrLn) + ", " + Str(wRrToLn) + ", " + _
                        Str(wRrCstRt) + " ,  " + Str(wRrSalRt) + ", " + Str(wRrCstRtInc) + "," + Str(wRrSalRtInc) + "," + CStr(wRrXWtPer) + ")")
                                
    Next

    CloseExcel
    ' ***** Manali 3.9.1 - 21/02/12 - qSr added for Order By according to data entered in Excel
    Set wo_rsExcel = moCn.OpenRes("Select * From #TmpXls Order By qSr")
    wo_rsExcel.MoveFirst
    
    Dim wRow As Double
    
    If wo_rsExcel.RecCount > 0 Then
      With GrdRm
        .StartCopy
        MWLib.BeginProcess Me, "Copying IPD Details ..."
        Do While Not wo_rsExcel.EOF
          .AddItem
          wRow = .Rows - 1
          .Row = wRow
          
          .Value(wRow, "RrPrtKey") = ctCurrPrtn
          .Value(wRow, "RrCmCd") = wo_rsExcel!RrCmCd
          .Value(wRow, "RrCmCurCd") = wo_rsExcel!RrCmCurCd
          .Value(wRow, "RrSCtg") = wo_rsExcel!RrSCtg
          .Value(wRow, "RrCd") = wo_rsExcel!RrCd
          .Value(wRow, "RrFrLn") = wo_rsExcel!RrFrLn
          .Value(wRow, "RrToLn") = wo_rsExcel!RrToLn
          .Value(wRow, "RrCstRt") = wo_rsExcel!RrCstRt
          .Value(wRow, "RrSalRt") = wo_rsExcel!RrSalRt
            .Value(wRow, "RrFixMinTol") = wo_rsExcel!RrFixMinTol
            .Value(wRow, "RrFixMaxTol") = wo_rsExcel!RrFixMaxTol
            .Value(wRow, "RrXWtPer") = wo_rsExcel!RrXWtPer
          
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
  ws_BtnStr = "CmdCpyHistGo,CmdCpyGo,CmdILCGo"     '*** Jay 3.01 *** [ILC]
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
Private Sub GrdRm_Validate(Cancel As Boolean)
  Cancel = GrdRm.Validate
End Sub
'*** Jay 3.01 *** [ILC]
Private Sub GetILC_Rts()
Dim wILC_Rts As MwfLib.MDORowSet
Dim wi_D As Integer
Dim w_NewRt As Single
Dim wSqlStrg As String, ws_RoundOff As String

For wi_D = 1 To GrdRm.Cols - 1
  If GrdRm.ColProp(wi_D).Style = fgcNormal Then GrdRm.Col = wi_D: Exit For
Next wi_D

For wi_D = 1 To GrdRm.Rows - 1
  GrdRm.Row = wi_D
  w_NewRt = 0
  '****** Sachin 3.02.0 - Partition Key added
  wSqlStrg = "Select Rt.RrCmCtg as qRrCmCtg, Rt.RrCmCd as qRrCmCd, Rt.RrCtg as qRrCtg, Rt.RrSCtg as qRrSCtg, " + _
    "Rt.RrCd as qRrCd, Rt.RrFrLn as qRrFrLn, Rt.RrToLn as qRrToLn, Z.RrCstRt as qZCstRt, " + _
    "IsNull(D1.RrXWtPer, IsNull(D2.RrXWtPer, IsNull(D3.RrXWtPer, IsNull(D4.RrXWtPer, 0)))) as qPer, " + _
    "IsNull(D1.RrSalRt, IsNull(D2.RrSalRt, IsNull(D3.RrSalRt, IsNull(D4.RrSalRt, 0)))) as qRndOff, " + _
    "IsNull(D1.RrFixMinTol, IsNull(D2.RrFixMinTol, IsNull(D3.RrFixMinTol, IsNull(D4.RrFixMinTol, 0)))) as qMinInc, " + _
    "IsNull(D1.RrFixMaxTol, IsNull(D2.RrFixMaxTol, IsNull(D3.RrFixMaxTol, IsNull(D4.RrFixMaxTol, 0)))) As qMaxInc " + _
    "From RmRt Rt"
  wSqlStrg = wSqlStrg + " join RmRt Z on Z.RrPrtKey='" + ctCurrPrtn + "' and Z.RrTcTyp= 'RM' and Z.RrCmCd= 'ZSELF' and Z.RrCtg= Rt.RrCtg and Z.RrSCtg= Rt.RrSCtg " + _
    "  and Z.RrCd= Rt.RrCd and convert( decimal (14, 4), Z.RrFrLn)= convert( decimal (14, 4), Rt.RrFrLn) " + _
    "  and convert( decimal (14, 4), Z.RrToLn)= convert( decimal (14, 4), Rt.RrToLn) and Z.RrCmCurCd= Rt.RrCmCurCd " + _
    "left outer join RmRt D1 on Rt.RrPrtKey=D1.RrPrtKey and D1.RrTcTyp= 'ILC' and D1.RrCmCtg= Rt.RrCmCtg and D1.RrCmCd= Rt.RrCmCd " + _
    "  and D1.RrCtg= Rt.RrCtg and D1.RrSCtg= Rt.RrSCtg and D1.RrCd= Rt.RrCd " + _
    "  and convert( decimal (14, 4), D1.RrFrLn)= convert( decimal (14, 4), Rt.RrFrLn) " + _
    "  and convert( decimal (14, 4), D1.RrToLn)= convert( decimal (14, 4), Rt.RrToLn) " + _
    "left outer join RmRt D2 on D2.RrPrtKey=Rt.RrPrtKey and D2.RrTcTyp= 'ILC' and D2.RrCmCtg= Rt.RrCmCtg and D2.RrCmCd= Rt.RrCmCd " + _
    "  and D2.RrCtg= Rt.RrCtg and D2.RrSCtg= Rt.RrSCtg and D2.RrCd= Rt.RrCd " + _
    "  and convert( decimal (14, 4), D2.RrFrLn)= convert( decimal (14, 4), 0) " + _
    "  and convert( decimal (14, 4), D2.RrToLn)= convert( decimal (14, 4), 0) " + _
    "left outer join RmRt D3 on D3.RrPrtKey=Rt.RrPrtKey and D3.RrTcTyp= 'ILC' and D3.RrCmCtg= Rt.RrCmCtg and D3.RrCmCd= Rt.RrCmCd " + _
    "  and D3.RrCtg= Rt.RrCtg and D3.RrSCtg= Rt.RrSCtg and D3.RrCd= '' " + _
    "  and convert( decimal (14, 4), D3.RrFrLn)= convert( decimal (14, 4), 0) " + _
    "  and convert( decimal (14, 4), D3.RrToLn)= convert( decimal (14, 4), 0) " + _
    "left outer join RmRt D4 on D4.RrPrtKey=Rt.RrPrtKey and D4.RrTcTyp= 'ILC' and D4.RrCmCtg= Rt.RrCmCtg and D4.RrCmCd= Rt.RrCmCd " + _
    "  and D4.RrCtg= Rt.RrCtg and D4.RrSCtg= '' and D4.RrCd= '' " + _
    "  and convert( decimal (14, 4), D4.RrFrLn)= convert( decimal (14, 4), 0) " + _
    "  and convert( decimal (14, 4), D4.RrToLn)= convert( decimal (14, 4), 0) " + _
    "where Rt.RrPrtKey='" + ctCurrPrtn + "' and Rt.RrTcTyp= '" + ADC("wRrTcTyp") + "' and Rt.RrCmCtg= '" + ADC("wRrCmCtg") + "' and Rt.RrCmCd= '" + GrdRm.Value(wi_D, "RrCmCd") + "' " + _
    "  and Rt.RrCtg= '" + ADC("wRrCtg") + "' and Rt.RrSCtg= '" + GrdRm.Value(wi_D, "RrSCtg") + "' " + _
    "  and Rt.RrCd= '" + GrdRm.Value(wi_D, "RrCd") + "' and Rt.RrFrLn= convert( decimal (14, 4), " + CStr(GrdRm.Value(wi_D, "RrFrLn")) + ") " + _
    "  and Rt.RrToLn= convert( decimal (14, 4), " + CStr(GrdRm.Value(wi_D, "RrToLn")) + ") " + _
    "  and not exists(Select 'x' from RmRt BS where BS.RrPrtKey=Rt.RrPrtKey and BS.RrTcTyp in('BS', 'PTR') and BS.RrCmCtg= Rt.RrCmCtg " + _
    "      and BS.RrCmCd= Rt.RrCmCd and BS.RrCtg= Rt.RrCtg and BS.RrSCtg= Rt.RrSCtg) "
  Set wILC_Rts = moCn.OpenRes(wSqlStrg)
  If Not (wILC_Rts.EOF Or wILC_Rts.BOF) Then
    With wILC_Rts
      w_NewRt = (!qZCstRt * !qPer / 100)
      w_NewRt = IIF(w_NewRt < !qMinInc, !qMinInc, IIF(w_NewRt > !qMaxInc, !qMaxInc, w_NewRt))
      If !qRndOff = 1 Then
         ws_RoundOff = "RNDUP1"
      ElseIf !qRndOff = 5 Then
         ws_RoundOff = "RNDUP5"
      ElseIf !qRndOff = 10 Then
         ws_RoundOff = "RNDUP10"
      ElseIf !qRndOff = 50 Then
         ws_RoundOff = "RNDUP50"
      ElseIf !qRndOff = 100 Then
         ws_RoundOff = "RNDUP100"
      Else
          ws_RoundOff = "NORND"
      End If
      w_NewRt = !qZCstRt + w_NewRt
      w_NewRt = FnRndOff(w_NewRt, 3, ws_RoundOff)
      GrdRm.Value(wi_D, "RrSalRt") = w_NewRt
      GrdRm.SaveRec (wi_D)
    End With
  End If
Next wi_D
Set wILC_Rts = Nothing
End Sub
'*** Jay 3.01 *** [ILC]

'*** Jay 3.01 *** [ILC]
Private Sub CmdILCGo_Click()
  Call CmdILC_Click
  Call GetILC_Rts
End Sub

'*** Jay 3.01 *** [ILC]
Private Sub CmdPwd_Click()
  Call DispFra(Pwd)
End Sub
'*** Jay 3.01 *** [ILC]
Private Sub TxtPwd_Validate(Cancel As Boolean)
  If TxtPwd.text = "" Then
      ms_PwdFlag = moCn.RecSeek("Select PValue1 from Param where PTyp= 'SEOPWD' and PmCd= 'SEOPWD' and PValue1= ''")
  Else
      If Trim(TxtPwd.text) = moCn.GetFldVal("select IsNull ((select PValue from Param Join UsrAccess on UaUsrCd = PMCd " + _
                       "where PTyp= 'USR' and PMCd= '" + gs_UsrCd + "' and UaMnuCd = '" + ADC.MenuCd + "' and " + _
                       "UaSeoYn= 'Y' and PValue= '" + Trim(TxtPwd.text) + "'), " + _
                       "(Select PValue1 from Param where Ptyp= 'SEOPWD' and PmCd= 'SEOPWD' )) ") Then
        ms_PwdFlag = True: Exit Sub
      Else
        ms_PwdFlag = False
        Cancel = True: DispMsg "Invalid SEO Password / User Does Not Have SEO Rights", etError
      End If
  End If
End Sub
'IG.189- Validating Rm From, To Size
Private Function RmSzLen(ByVal wTsMask As String, ByVal wRrLn As Double) As Boolean

Dim wMask() As String, wRmLn() As String, wMaskSz, wRmSz As String
'Getting Rm Mask size
wMask = Split(wTsMask, ".")
wMaskSz = wMask(0)

'Getting Rm size entered in excel
wRmLn = Split(wRrLn, ".")
wRmSz = wRmLn(0)
'Comparing both sizes, if entered size is greater then original mask size then will gives error message
If Len(CStr(wRrLn)) < 0 Or Len(wRmSz) > Len(wMaskSz) Then
  RmSzLen = True
Else
  RmSzLen = False
End If
End Function

'Uni.24-When user click Delete rates button system will check the Seo password entered, if not then gives error message
Private Sub CmdDelRt_Click()
  If Not ms_PwdFlag Then DispMsg "SEO Password Required", etError: Exit Sub
  Call DispFra(DelRt)
End Sub
'Uni.24-Calling DelRts function
Private Sub CmdDelRtGo_Click()
  Call CmdDelRt_Click
  Call DelRts
End Sub
'Uni.24- Deletes all the rows in the Grid
Private Sub DelRts()
  Dim ws_Cnd As String, wsql As String
  
  ws_Cnd = ADC.RepCond
  ws_Cnd = IIF(ws_Cnd <> "", " where ", " and ") + ws_Cnd

  wsql = "Delete from RmRt " + ws_Cnd
  moCn.Execute (wsql)
  GrdRm.Rows = 1

End Sub


