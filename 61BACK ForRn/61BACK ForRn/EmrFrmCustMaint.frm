VERSION 5.00
Object = "{AB2CC32A-E652-42A2-B341-26ECE02F6386}#2.1#0"; "mwfctlforspeed9.ocx"
Begin VB.Form EmrFrmCustMaint 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Invoice Custom Matching"
   ClientHeight    =   9810
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   15210
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MDIChild        =   -1  'True
   ScaleHeight     =   9810
   ScaleWidth      =   15210
   WindowState     =   2  'Maximized
   Begin MwfCtl.MWCTL_BTN1 CmdDelRm 
      Height          =   465
      Left            =   9660
      TabIndex        =   9
      ToolTipText     =   "Delete Raw Material"
      Top             =   9300
      Width           =   1035
      _ExtentX        =   1826
      _ExtentY        =   820
      BackColor       =   -2147483633
      MaskColor       =   8454143
      Caption         =   "&Delete Rm"
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
   Begin MwfCtl.MWCTL_BTN1 CmdAuto 
      Height          =   465
      Left            =   8610
      TabIndex        =   8
      ToolTipText     =   "Auto Custom Matching"
      Top             =   9300
      Width           =   1035
      _ExtentX        =   1826
      _ExtentY        =   820
      BackColor       =   -2147483633
      MaskColor       =   8454143
      Caption         =   "Auto &Match"
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
      Left            =   5505
      TabIndex        =   86
      ToolTipText     =   "Raw Material Issue"
      Top             =   9300
      Width           =   1035
      _ExtentX        =   1826
      _ExtentY        =   820
      BackColor       =   -2147483633
      MaskColor       =   8454143
      Caption         =   "&Help Sort Option "
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
   Begin MwfCtl.MWCTL_BTN1 CmdIZR 
      Height          =   465
      Left            =   7575
      TabIndex        =   7
      ToolTipText     =   "Bag Details "
      Top             =   9300
      Width           =   1035
      _ExtentX        =   1826
      _ExtentY        =   820
      BackColor       =   -2147483633
      MaskColor       =   8454143
      Caption         =   "&Rm Details"
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
      Left            =   6540
      TabIndex        =   6
      ToolTipText     =   "Copy Raw Material"
      Top             =   9300
      Width           =   1035
      _ExtentX        =   1826
      _ExtentY        =   820
      BackColor       =   -2147483633
      MaskColor       =   8454143
      Caption         =   "C&opy"
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
      Left            =   9750
      TabIndex        =   60
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
      Left            =   30
      TabIndex        =   61
      Top             =   9300
      Width           =   15105
      _ExtentX        =   26644
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
      Height          =   280
      Left            =   -36
      TabIndex        =   63
      Top             =   30
      Width           =   4320
      Begin MwfCtl.MWCTL_MED ATXT 
         Height          =   285
         Index           =   1
         Left            =   1095
         TabIndex        =   0
         ToolTipText     =   "Enter Invoice Voucher Tc"
         Top             =   0
         Width           =   570
         _ExtentX        =   1005
         _ExtentY        =   503
         Enabled         =   0   'False
         MaxLength       =   3
         DataType        =   4
         ReCalcParent    =   "WIZCHR"
         IdName          =   "WIZTC"
      End
      Begin MwfCtl.MWCTL_MED ATXT 
         Height          =   285
         Index           =   4
         Left            =   2985
         TabIndex        =   3
         ToolTipText     =   "Enter Invoice Voucher Number"
         Top             =   0
         Width           =   930
         _ExtentX        =   1640
         _ExtentY        =   503
         Alignment       =   2
         Mask            =   "######0"
         MaxLength       =   7
         DataType        =   1
         IdName          =   "WIZNO"
      End
      Begin MwfCtl.MWCTL_MED ATXT 
         Height          =   285
         Index           =   5
         Left            =   2310
         TabIndex        =   2
         ToolTipText     =   "Enter Invoice Voucher Tc"
         Top             =   0
         Width           =   570
         _ExtentX        =   1005
         _ExtentY        =   503
         MaxLength       =   3
         DataType        =   4
         ReCalcOn        =   "WIZTC"
         IdName          =   "WIZCHR"
      End
      Begin MwfCtl.MWCTL_MED ATXT 
         Height          =   285
         Index           =   6
         Left            =   1770
         TabIndex        =   1
         ToolTipText     =   "Enter Invoice Voucher Year"
         Top             =   0
         Width           =   435
         _ExtentX        =   767
         _ExtentY        =   503
         MaxLength       =   2
         DataType        =   4
         IdName          =   "WIZYY"
      End
      Begin MwfCtl.MWCTL_MED ATXT 
         Height          =   285
         Index           =   14
         Left            =   3960
         TabIndex        =   107
         ToolTipText     =   "Partition Key"
         Top             =   0
         Visible         =   0   'False
         Width           =   345
         _ExtentX        =   609
         _ExtentY        =   503
         Enabled         =   0   'False
         IdName          =   "WIZPRTKEY"
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
         Left            =   2220
         TabIndex        =   68
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
         Left            =   2895
         TabIndex        =   67
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
         Left            =   1680
         TabIndex        =   66
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
         TabIndex        =   64
         Top             =   -5
         Width           =   1065
      End
   End
   Begin VB.Frame FRA_NKEYS 
      BorderStyle     =   0  'None
      Enabled         =   0   'False
      Height          =   9420
      Left            =   -120
      TabIndex        =   62
      Top             =   0
      Width           =   15315
      Begin VB.Frame FraNKeyAll 
         BorderStyle     =   0  'None
         Height          =   9315
         Left            =   150
         TabIndex        =   65
         Top             =   30
         Width           =   15135
         Begin MwfCtl.MWCTL_FLX GrdIZ 
            Height          =   6075
            Left            =   15
            TabIndex        =   4
            Top             =   615
            Width           =   15120
            _ExtentX        =   26670
            _ExtentY        =   10716
            Cols            =   11
            AllowAdd        =   0   'False
            colname1        =   "IZSR"
            heading1        =   "Sr.    "
            datafld1        =   "IzSr"
            datatype1       =   1
            mask1           =   "###0"
            recalcparent1   =   "GrdIZ(WIZPURAMT)"
            maxlength1      =   4
            style1          =   2
            tooltiptext1    =   "Enter Sr. No"
            colname2        =   "IZRMZCD"
            heading2        =   "Customs Code          "
            datafld2        =   "IzRmZCd"
            datatype2       =   4
            maxlength2      =   12
            tooltiptext2    =   "Raw Material Customs Code"
            colname3        =   "IZRMQTY"
            heading3        =   "Rm Qty "
            datafld3        =   "IzRmQty"
            datatype3       =   1
            mask3           =   "#####0"
            recalcparent3   =   "GrdIZ(WIZPCPERCT)"
            maxlength3      =   6
            tooltiptext3    =   "Raw Material Quantity"
            colname4        =   "IZRMZWT"
            heading4        =   "Rm Wt       "
            datafld4        =   "IzRmZWt"
            datatype4       =   2
            mask4           =   "####0.000"
            recalcparent4   =   "GrdIZ(WIZRMZAMT),GrdIZ(WIZRMBALWT),GrdIZ(WIZPCPERCT)"
            maxlength4      =   9
            tooltiptext4    =   "Raw Material Weight"
            colname5        =   "WIZPCPERCT"
            heading5        =   "Pc/Ct"
            datatype5       =   1
            mask5           =   "##0"
            recalcon5       =   "GrdIZ(IZRMQTY),GrdIZ(IZRMZWT)"
            maxlength5      =   3
            tooltiptext5    =   "Piece Per Carat"
            colname6        =   "IZRMZRT"
            heading6        =   "Rm Rt               "
            datafld6        =   "IzRmZRt"
            datatype6       =   2
            mask6           =   "#######0.000"
            recalcparent6   =   "GrdIZ(WIZRMZAMT)"
            maxlength6      =   12
            tooltiptext6    =   "Raw Material Rate"
            colname7        =   "WIZRMZAMT"
            heading7        =   "Rm Amt            "
            datatype7       =   2
            mask7           =   "########0.00"
            recalcon7       =   "GrdIZ(IZRMZWT),GrdIZ(IZRMZRT)"
            maxlength7      =   12
            tooltiptext7    =   "Raw Material Amount"
            colname8        =   "WIZRMBALWT"
            heading8        =   "Balance Wt"
            datatype8       =   2
            mask8           =   "####0.000"
            recalcon8       =   "GrdIZ(IZRMZWT),GrdTxndZ(TZRMWT)"
            maxlength8      =   9
            tooltiptext8    =   "Balance Raw Material Weight"
            colname9        =   "WIZPURAMT"
            heading9        =   "Pur Amount      "
            datatype9       =   2
            mask9           =   "########0.00"
            recalcon9       =   "GrdTxndZ(TZPURAMT),GrdIZ(IZSR)"
            maxlength9      =   12
            tooltiptext9    =   "Purchase Amount"
            colname10       =   "IZPRTKEY"
            heading10       =   "Prtn "
            datafld10       =   "IzPrtKey"
            maxlength10     =   1
            style10         =   0
         End
         Begin MwfCtl.MWCTL_FLX GrdTxndZ 
            Height          =   2520
            Left            =   15
            TabIndex        =   5
            Top             =   6690
            Width           =   9795
            _ExtentX        =   17277
            _ExtentY        =   4445
            Cols            =   12
            colname1        =   "TZSRNO"
            heading1        =   "Sr. "
            datafld1        =   "TzSrNo"
            datatype1       =   1
            mask1           =   "##0"
            maxlength1      =   3
            style1          =   2
            tooltiptext1    =   "Enter Sr No."
            colname2        =   "WTZBASEVCHNO"
            heading2        =   "Base Voucher Number        "
            datatype2       =   4
            recalcon2       =   "GrdTxndZ(TZREFYY),GrdTxndZ(TZREFKEY)"
            recalcparent2   =   $"EmrFrmCustMaint.frx":0000
            maxlength2      =   22
            tooltiptext2    =   "Enter Base Voucher Number"
            colname3        =   "WTZBASEBOENO"
            heading3        =   "Base BOE No.                        "
            datatype3       =   4
            recalcon3       =   "GrdTxndZ(WTZBASEVCHNO)"
            maxlength3      =   20
            tooltiptext3    =   "Base Bill Of Entry Number"
            colname4        =   "TZRMWT"
            heading4        =   "Rm Wt      "
            datafld4        =   "TzRmWt"
            datatype4       =   2
            mask4           =   "####0.000"
            recalcon4       =   "GrdTxndZ(WTZBASEVCHNO)"
            recalcparent4   =   "GrdIZ(WIZRMBALWT),GrdTxndZ(TZPURAMT)"
            maxlength4      =   9
            tooltiptext4    =   "Enter Raw Material Weight"
            colname5        =   "TZPURRT"
            heading5        =   "Pur Rate         "
            datafld5        =   "TzPurRt"
            datatype5       =   2
            mask5           =   "#######0.000"
            recalcon5       =   "GrdTxndZ(WTZBASEVCHNO)"
            recalcparent5   =   "GrdTxndZ(TZPURAMT)"
            maxlength5      =   12
            tooltiptext5    =   "Purchase Rate"
            colname6        =   "WPURRTINRS"
            heading6        =   "Pur Rt In Rs     "
            datatype6       =   2
            mask6           =   "########0.00"
            recalcon6       =   "GrdTxndZ(WTZBASEVCHNO)"
            maxlength6      =   12
            tooltiptext6    =   "Purchase Rate In Rs."
            colname7        =   "TZPURAMT"
            heading7        =   "Pur Amount      "
            datafld7        =   "TzPurAmt"
            datatype7       =   2
            mask7           =   "########0.00"
            recalcon7       =   "GrdTxndZ(TZRMWT),GrdTxndZ(TZPURRT)"
            recalcparent7   =   "GrdIZ(WIZPURAMT)"
            maxlength7      =   12
            tooltiptext7    =   "Purchase Amount"
            colname8        =   "TZKEY"
            datafld8        =   "TzKey"
            datatype8       =   1
            mask8           =   "######0"
            maxlength8      =   7
            style8          =   0
            tooltiptext8    =   "Key"
            colname9        =   "TZREFYY"
            heading9        =   "RefYY "
            datafld9        =   "TzRefYy"
            datatype9       =   4
            recalcon9       =   "GrdTxndZ(WTZBASEVCHNO)"
            recalcparent9   =   "GrdTxndZ(WTZBASEVCHNO)"
            maxlength9      =   2
            style9          =   0
            tooltiptext9    =   "Reference Year"
            colname10       =   "TZREFKEY"
            heading10       =   "RefKey"
            datafld10       =   "TzRefKey"
            datatype10      =   1
            mask10          =   "######0"
            recalcon10      =   "GrdTxndZ(WTZBASEVCHNO)"
            recalcparent10  =   "GrdTxndZ(WTZBASEVCHNO)"
            maxlength10     =   7
            style10         =   0
            tooltiptext10   =   "Reference Key"
            colname11       =   "TZPRTKEY"
            heading11       =   "Prtn "
            datafld11       =   "TzPrtKey"
            maxlength11     =   1
            style11         =   0
         End
         Begin MwfCtl.MWCTL_MED ATXT 
            Height          =   285
            Index           =   30
            Left            =   1035
            TabIndex        =   73
            ToolTipText     =   "Invoice Customer Code"
            Top             =   285
            Width           =   1365
            _ExtentX        =   2408
            _ExtentY        =   503
            Enabled         =   0   'False
            MaxLength       =   8
            DataType        =   4
            IdName          =   "WIZCMCD"
         End
         Begin MwfCtl.MWCTL_MED ATXT 
            Height          =   285
            Index           =   0
            Left            =   4815
            TabIndex        =   74
            ToolTipText     =   "Invoice Date"
            Top             =   0
            Width           =   885
            _ExtentX        =   1561
            _ExtentY        =   503
            Enabled         =   0   'False
            MaxLength       =   8
            DataType        =   3
            IdName          =   "WIZDT"
         End
         Begin MwfCtl.MWCTL_MED ATXT 
            Height          =   285
            Index           =   2
            Left            =   7005
            TabIndex        =   75
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
            IdName          =   "WIZLMGSAL"
         End
         Begin MwfCtl.MWCTL_MED ATXT 
            Height          =   285
            Index           =   3
            Left            =   4815
            TabIndex        =   76
            ToolTipText     =   "Invoice Export Number"
            Top             =   285
            Width           =   3075
            _ExtentX        =   5424
            _ExtentY        =   503
            Enabled         =   0   'False
            MaxLength       =   20
            DataType        =   4
            IdName          =   "WIZEXPNO"
         End
         Begin MwfCtl.MWCTL_MED ATXT 
            Height          =   285
            Index           =   8
            Left            =   9105
            TabIndex        =   84
            ToolTipText     =   "LMG Sales Rate"
            Top             =   15
            Width           =   885
            _ExtentX        =   1561
            _ExtentY        =   503
            Alignment       =   2
            Enabled         =   0   'False
            Mask            =   "###0.00"
            MaxLength       =   7
            DataType        =   2
            IdName          =   "WIZLMPSAL"
         End
         Begin MwfCtl.MWCTL_FLX GrdIZD 
            Height          =   2520
            Left            =   9630
            TabIndex        =   83
            Top             =   6690
            Width           =   5505
            _ExtentX        =   9710
            _ExtentY        =   4445
            Cols            =   5
            AllowAdd        =   0   'False
            AllowDelete     =   0   'False
            colname1        =   "WIZDIRRMCD"
            heading1        =   "Rm Code               "
            datatype1       =   4
            maxlength1      =   12
            tooltiptext1    =   "Raw Material Code"
            colname2        =   "WIZDRMRSZ"
            heading2        =   "Range Sz             "
            datatype2       =   4
            maxlength2      =   12
            tooltiptext2    =   "Range Size"
            colname3        =   "WIZDIRRMSZ"
            heading3        =   "Size      "
            maxlength3      =   7
            tooltiptext3    =   "Rm Size"
            colname4        =   "WIZDIRRMZWT"
            heading4        =   "Wt.               "
            datatype4       =   2
            mask4           =   "####0.000"
            maxlength4      =   9
            tooltiptext4    =   "Raw Material Weight"
         End
         Begin MwfCtl.MWCTL_MED ATXT 
            Height          =   285
            Index           =   10
            Left            =   11535
            TabIndex        =   93
            ToolTipText     =   "LMG Sales Rate"
            Top             =   15
            Width           =   885
            _ExtentX        =   1561
            _ExtentY        =   503
            Alignment       =   2
            Enabled         =   0   'False
            Mask            =   "###0.00"
            MaxLength       =   7
            DataType        =   2
            IdName          =   "WIZLMSSAL"
         End
         Begin MwfCtl.MWCTL_MED ATXT 
            Height          =   285
            Index           =   73
            Left            =   9105
            TabIndex        =   95
            ToolTipText     =   "Invoice Customer Currency Code"
            Top             =   285
            Width           =   1035
            _ExtentX        =   1826
            _ExtentY        =   503
            Enabled         =   0   'False
            MaxLength       =   5
            DataType        =   4
            IdName          =   "WIZCMCURCD"
         End
         Begin MwfCtl.MWCTL_MED ATXT 
            Height          =   285
            Index           =   74
            Left            =   11535
            TabIndex        =   96
            ToolTipText     =   "Conversion Factor"
            Top             =   285
            Width           =   1095
            _ExtentX        =   1931
            _ExtentY        =   503
            Enabled         =   0   'False
            Mask            =   "###0.0000"
            MaxLength       =   9
            DataType        =   2
            IdName          =   "WIZCNVFCT"
         End
         Begin MwfCtl.MWCTL_MED ATXT 
            Height          =   285
            Index           =   11
            Left            =   14205
            TabIndex        =   97
            ToolTipText     =   "Multiplication/Division"
            Top             =   285
            Width           =   345
            _ExtentX        =   609
            _ExtentY        =   503
            Enabled         =   0   'False
            DataType        =   4
            IdName          =   "WIZMULDIV"
         End
         Begin MwfCtl.MWCTL_MED ATXT 
            Height          =   285
            Index           =   32
            Left            =   14205
            TabIndex        =   137
            ToolTipText     =   "LML Sales Rate"
            Top             =   15
            Width           =   885
            _ExtentX        =   1561
            _ExtentY        =   503
            Alignment       =   2
            Enabled         =   0   'False
            Mask            =   "###0.00"
            MaxLength       =   7
            DataType        =   2
            IdName          =   "WIZLMLSAL"
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
            Left            =   13080
            TabIndex        =   138
            Top             =   15
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
            Left            =   4290
            TabIndex        =   79
            Top             =   0
            Width           =   585
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
            Left            =   7980
            TabIndex        =   100
            Top             =   285
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
            Left            =   10410
            TabIndex        =   99
            Top             =   285
            Width           =   1005
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
            Left            =   13080
            TabIndex        =   98
            Top             =   285
            Width           =   855
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
            Left            =   10410
            TabIndex        =   94
            Top             =   15
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
            Index           =   3
            Left            =   7980
            TabIndex        =   85
            Top             =   15
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
            Index           =   4
            Left            =   0
            TabIndex        =   80
            Top             =   285
            Width           =   1005
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
            Left            =   5880
            TabIndex        =   78
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
            Index           =   24
            Left            =   3750
            TabIndex        =   77
            Top             =   285
            Width           =   1125
         End
      End
      Begin VB.Frame FraIZR 
         Enabled         =   0   'False
         Height          =   3555
         Left            =   155
         TabIndex        =   69
         Top             =   5685
         Width           =   15140
         Begin MwfCtl.MWCTL_MED ATXT 
            Height          =   285
            Index           =   75
            Left            =   2970
            TabIndex        =   70
            Top             =   3180
            Width           =   825
            _ExtentX        =   1455
            _ExtentY        =   503
            Alignment       =   2
            Enabled         =   0   'False
            Mask            =   "######0"
            MaxLength       =   7
            DataType        =   2
            IdName          =   "WIZRRMTOTQTY"
         End
         Begin MwfCtl.MWCTL_MED ATXT 
            Height          =   285
            Index           =   76
            Left            =   3780
            TabIndex        =   72
            Top             =   3180
            Width           =   1005
            _ExtentX        =   1773
            _ExtentY        =   503
            Alignment       =   2
            Enabled         =   0   'False
            Mask            =   "#####0.000"
            MaxLength       =   10
            DataType        =   2
            IdName          =   "WIZRRMTOTWT"
         End
         Begin MwfCtl.MWCTL_FLX GrdIZR 
            CausesValidation=   0   'False
            Height          =   2670
            Left            =   60
            TabIndex        =   71
            Top             =   510
            Width           =   10575
            _ExtentX        =   18653
            _ExtentY        =   4710
            Cols            =   8
            AllowDelete     =   0   'False
            colname1        =   "WIZRIDSR"
            heading1        =   "InvSr."
            datatype1       =   1
            mask1           =   "###0"
            maxlength1      =   4
            tooltiptext1    =   "Invoice Design Serial"
            colname2        =   "WIZRSRNO"
            heading2        =   "SrNo"
            datatype2       =   1
            mask2           =   "##0"
            maxlength2      =   3
            tooltiptext2    =   "Serial Number"
            colname3        =   "WIZRIRRMCD"
            heading3        =   "Rm Code                "
            datatype3       =   4
            maxlength3      =   12
            tooltiptext3    =   "Raw Material Code"
            colname4        =   "WIZRIRRMQTY"
            heading4        =   "Qty       "
            datatype4       =   1
            mask4           =   "#####0"
            maxlength4      =   6
            tooltiptext4    =   "Raw Material Quantity"
            colname5        =   "WIZRIRRMZWT"
            heading5        =   "Wt.             "
            datatype5       =   2
            mask5           =   "####0.000"
            maxlength5      =   9
            tooltiptext5    =   "Raw Material Weight"
            colname6        =   "WIZRIDORDNO"
            heading6        =   "Order Number                            "
            datatype6       =   4
            maxlength6      =   21
            tooltiptext6    =   "Order Number"
            colname7        =   "WIZRIDDMCD"
            heading7        =   "Design Cd                     "
            datatype7       =   4
            maxlength7      =   15
            tooltiptext7    =   "Design Code"
         End
         Begin VB.Label LblMain 
            BackStyle       =   0  'Transparent
            BorderStyle     =   1  'Fixed Single
            Caption         =   "  Raw Material Details"
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
            TabIndex        =   92
            Top             =   90
            Width           =   15135
         End
      End
      Begin VB.Frame FraAuto 
         Height          =   3735
         Left            =   150
         TabIndex        =   108
         Top             =   5520
         Width           =   15180
         Begin MwfCtl.MWCTL_BTN1 CmdAutoGo 
            Height          =   495
            Left            =   14490
            TabIndex        =   55
            ToolTipText     =   "Start Auto Custom Matching For The Given Range"
            Top             =   3075
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
            Index           =   82
            Left            =   7860
            TabIndex        =   12
            ToolTipText     =   "Enter The Rm Categories For Which Auto Matching Have To Be Done"
            Top             =   480
            Visible         =   0   'False
            Width           =   1365
            _ExtentX        =   2408
            _ExtentY        =   503
            Enabled         =   0   'False
            MaxLength       =   8
            DataType        =   4
            IdName          =   "WAUTOCPYOPT"
         End
         Begin MwfCtl.MWCTL_MED ATXT 
            Height          =   285
            Index           =   158
            Left            =   540
            TabIndex        =   13
            ToolTipText     =   "Enter Rate Variation % From Range For Diamonds"
            Top             =   1590
            Width           =   765
            _ExtentX        =   1349
            _ExtentY        =   503
            Alignment       =   2
            Mask            =   "-##0.00"
            MaxLength       =   7
            DataType        =   2
            IdName          =   "WAUTOHRTMINUSD"
         End
         Begin MwfCtl.MWCTL_MED ATXT 
            Height          =   285
            Index           =   159
            Left            =   1380
            TabIndex        =   14
            ToolTipText     =   "Enter Rate Variation % To Range For Diamonds"
            Top             =   1590
            Width           =   765
            _ExtentX        =   1349
            _ExtentY        =   503
            Alignment       =   2
            Mask            =   "-##0.00"
            MaxLength       =   7
            DataType        =   2
            IdName          =   "WAUTOHRTPLUSD"
         End
         Begin MwfCtl.MWCTL_MED ATXT 
            Height          =   285
            Index           =   160
            Left            =   2340
            TabIndex        =   15
            ToolTipText     =   "Enter Pc/ Ct Variation % From Range For Diamonds"
            Top             =   1590
            Width           =   765
            _ExtentX        =   1349
            _ExtentY        =   503
            Alignment       =   2
            Mask            =   "-##0.00"
            MaxLength       =   7
            DataType        =   2
            IdName          =   "WAUTOHPTRMINUSD"
         End
         Begin MwfCtl.MWCTL_MED ATXT 
            Height          =   285
            Index           =   161
            Left            =   3180
            TabIndex        =   16
            ToolTipText     =   "Enter Pc/ Ct Variation % To Range For Diamonds"
            Top             =   1590
            Width           =   765
            _ExtentX        =   1349
            _ExtentY        =   503
            Alignment       =   2
            Mask            =   "-##0.00"
            MaxLength       =   7
            DataType        =   2
            IdName          =   "WAUTOHPTRPLUSD"
         End
         Begin MwfCtl.MWCTL_MED ATXT 
            Height          =   285
            Index           =   162
            Left            =   540
            TabIndex        =   20
            ToolTipText     =   "Enter Rate Variation % From Range For Color Stones"
            Top             =   1875
            Width           =   765
            _ExtentX        =   1349
            _ExtentY        =   503
            Alignment       =   2
            Mask            =   "-##0.00"
            MaxLength       =   7
            DataType        =   2
            IdName          =   "WAUTOHRTMINUSC"
         End
         Begin MwfCtl.MWCTL_MED ATXT 
            Height          =   285
            Index           =   163
            Left            =   1380
            TabIndex        =   21
            ToolTipText     =   "Enter Rate Variation % To Range For Color Stones"
            Top             =   1875
            Width           =   765
            _ExtentX        =   1349
            _ExtentY        =   503
            Alignment       =   2
            Mask            =   "-##0.00"
            MaxLength       =   7
            DataType        =   2
            IdName          =   "WAUTOHRTPLUSC"
         End
         Begin MwfCtl.MWCTL_MED ATXT 
            Height          =   285
            Index           =   164
            Left            =   2340
            TabIndex        =   22
            ToolTipText     =   "Enter Pc/ Ct Variation % From Range For Color Stones"
            Top             =   1875
            Width           =   765
            _ExtentX        =   1349
            _ExtentY        =   503
            Alignment       =   2
            Mask            =   "-##0.00"
            MaxLength       =   7
            DataType        =   2
            IdName          =   "WAUTOHPTRMINUSC"
         End
         Begin MwfCtl.MWCTL_MED ATXT 
            Height          =   285
            Index           =   165
            Left            =   3180
            TabIndex        =   23
            ToolTipText     =   "Enter Pc/ Ct Variation % To Range For Color Stones"
            Top             =   1875
            Width           =   765
            _ExtentX        =   1349
            _ExtentY        =   503
            Alignment       =   2
            Mask            =   "-##0.00"
            MaxLength       =   7
            DataType        =   2
            IdName          =   "WAUTOHPTRPLUSC"
         End
         Begin MwfCtl.MWCTL_MED ATXT 
            Height          =   285
            Index           =   166
            Left            =   540
            TabIndex        =   27
            ToolTipText     =   "Enter Rate Variation % From Range For Gold"
            Top             =   2160
            Width           =   765
            _ExtentX        =   1349
            _ExtentY        =   503
            Alignment       =   2
            Mask            =   "-##0.00"
            MaxLength       =   7
            DataType        =   2
            IdName          =   "WAUTOHRTMINUSG"
         End
         Begin MwfCtl.MWCTL_MED ATXT 
            Height          =   285
            Index           =   167
            Left            =   1380
            TabIndex        =   28
            ToolTipText     =   "Enter Rate Variation % To Range For Gold"
            Top             =   2160
            Width           =   765
            _ExtentX        =   1349
            _ExtentY        =   503
            Alignment       =   2
            Mask            =   "-##0.00"
            MaxLength       =   7
            DataType        =   2
            IdName          =   "WAUTOHRTPLUSG"
         End
         Begin MwfCtl.MWCTL_MED ATXT 
            Height          =   285
            Index           =   168
            Left            =   2340
            TabIndex        =   29
            ToolTipText     =   "Enter Pc/ Ct Variation % From Range For Gold"
            Top             =   2160
            Width           =   765
            _ExtentX        =   1349
            _ExtentY        =   503
            Alignment       =   2
            Mask            =   "-##0.00"
            MaxLength       =   7
            DataType        =   2
            IdName          =   "WAUTOHPTRMINUSG"
         End
         Begin MwfCtl.MWCTL_MED ATXT 
            Height          =   285
            Index           =   169
            Left            =   3180
            TabIndex        =   30
            ToolTipText     =   "Enter Pc/ Ct Variation % To Range For Gold"
            Top             =   2160
            Width           =   765
            _ExtentX        =   1349
            _ExtentY        =   503
            Alignment       =   2
            Mask            =   "-##0.00"
            MaxLength       =   7
            DataType        =   2
            IdName          =   "WAUTOHPTRPLUSG"
         End
         Begin MwfCtl.MWCTL_MED ATXT 
            Height          =   285
            Index           =   170
            Left            =   540
            TabIndex        =   34
            ToolTipText     =   "Enter Rate Variation % From Range For Platinum"
            Top             =   2445
            Width           =   765
            _ExtentX        =   1349
            _ExtentY        =   503
            Alignment       =   2
            Mask            =   "-##0.00"
            MaxLength       =   7
            DataType        =   2
            IdName          =   "WAUTOHRTMINUSP"
         End
         Begin MwfCtl.MWCTL_MED ATXT 
            Height          =   285
            Index           =   171
            Left            =   1380
            TabIndex        =   35
            ToolTipText     =   "Enter Rate Variation % To Range For Platinum"
            Top             =   2445
            Width           =   765
            _ExtentX        =   1349
            _ExtentY        =   503
            Alignment       =   2
            Mask            =   "-##0.00"
            MaxLength       =   7
            DataType        =   2
            IdName          =   "WAUTOHRTPLUSP"
         End
         Begin MwfCtl.MWCTL_MED ATXT 
            Height          =   285
            Index           =   172
            Left            =   2340
            TabIndex        =   36
            ToolTipText     =   "Enter Pc/ Ct Variation % From Range For Platinum"
            Top             =   2445
            Width           =   765
            _ExtentX        =   1349
            _ExtentY        =   503
            Alignment       =   2
            Mask            =   "-##0.00"
            MaxLength       =   7
            DataType        =   2
            IdName          =   "WAUTOHPTRMINUSP"
         End
         Begin MwfCtl.MWCTL_MED ATXT 
            Height          =   285
            Index           =   173
            Left            =   3180
            TabIndex        =   37
            ToolTipText     =   "Enter Pc/ Ct Variation % To Range For Platinum"
            Top             =   2445
            Width           =   765
            _ExtentX        =   1349
            _ExtentY        =   503
            Alignment       =   2
            Mask            =   "-##0.00"
            MaxLength       =   7
            DataType        =   2
            IdName          =   "WAUTOHPTRPLUSP"
         End
         Begin MwfCtl.MWCTL_MED ATXT 
            Height          =   285
            Index           =   174
            Left            =   540
            TabIndex        =   41
            ToolTipText     =   "Enter Rate Variation % From Range For Silver"
            Top             =   2730
            Width           =   765
            _ExtentX        =   1349
            _ExtentY        =   503
            Alignment       =   2
            Mask            =   "-##0.00"
            MaxLength       =   7
            DataType        =   2
            IdName          =   "WAUTOHRTMINUSS"
         End
         Begin MwfCtl.MWCTL_MED ATXT 
            Height          =   285
            Index           =   175
            Left            =   1380
            TabIndex        =   42
            ToolTipText     =   "Enter Rate Variation % To Range For Silver"
            Top             =   2730
            Width           =   765
            _ExtentX        =   1349
            _ExtentY        =   503
            Alignment       =   2
            Mask            =   "-##0.00"
            MaxLength       =   7
            DataType        =   2
            IdName          =   "WAUTOHRTPLUSS"
         End
         Begin MwfCtl.MWCTL_MED ATXT 
            Height          =   285
            Index           =   176
            Left            =   2340
            TabIndex        =   43
            ToolTipText     =   "Enter Pc/ Ct Variation % From Range For Silver"
            Top             =   2730
            Width           =   765
            _ExtentX        =   1349
            _ExtentY        =   503
            Alignment       =   2
            Mask            =   "-##0.00"
            MaxLength       =   7
            DataType        =   2
            IdName          =   "WAUTOHPTRMINUSS"
         End
         Begin MwfCtl.MWCTL_MED ATXT 
            Height          =   285
            Index           =   177
            Left            =   3180
            TabIndex        =   44
            ToolTipText     =   "Enter Pc/ Ct Variation % To Range For Silver"
            Top             =   2730
            Width           =   765
            _ExtentX        =   1349
            _ExtentY        =   503
            Alignment       =   2
            Mask            =   "-##0.00"
            MaxLength       =   7
            DataType        =   2
            IdName          =   "WAUTOHPTRPLUSS"
         End
         Begin MwfCtl.MWCTL_MED ATXT 
            Height          =   285
            Index           =   178
            Left            =   540
            TabIndex        =   48
            ToolTipText     =   "Enter Rate Variation % From Range For Accessories"
            Top             =   3300
            Width           =   765
            _ExtentX        =   1349
            _ExtentY        =   503
            Alignment       =   2
            Mask            =   "-##0.00"
            MaxLength       =   7
            DataType        =   2
            IdName          =   "WAUTOHRTMINUSX"
         End
         Begin MwfCtl.MWCTL_MED ATXT 
            Height          =   285
            Index           =   179
            Left            =   1380
            TabIndex        =   49
            ToolTipText     =   "Enter Rate Variation % To Range For Accessories"
            Top             =   3300
            Width           =   765
            _ExtentX        =   1349
            _ExtentY        =   503
            Alignment       =   2
            Mask            =   "-##0.00"
            MaxLength       =   7
            DataType        =   2
            IdName          =   "WAUTOHRTPLUSX"
         End
         Begin MwfCtl.MWCTL_MED ATXT 
            Height          =   285
            Index           =   180
            Left            =   2340
            TabIndex        =   50
            ToolTipText     =   "Enter Pc/ Ct Variation % From Range For Accessories"
            Top             =   3300
            Width           =   765
            _ExtentX        =   1349
            _ExtentY        =   503
            Alignment       =   2
            Mask            =   "-##0.00"
            MaxLength       =   7
            DataType        =   2
            IdName          =   "WAUTOHPTRMINUSX"
         End
         Begin MwfCtl.MWCTL_MED ATXT 
            Height          =   285
            Index           =   181
            Left            =   3180
            TabIndex        =   51
            ToolTipText     =   "Enter Pc/ Ct Variation % To Range For Accessories"
            Top             =   3300
            Width           =   765
            _ExtentX        =   1349
            _ExtentY        =   503
            Alignment       =   2
            Mask            =   "-##0.00"
            MaxLength       =   7
            DataType        =   2
            IdName          =   "WAUTOHPTRPLUSX"
         End
         Begin MwfCtl.MWCTL_MED ATXT 
            Height          =   285
            Index           =   182
            Left            =   4230
            TabIndex        =   17
            ToolTipText     =   "Enter Import Types Not To Be Used For Diamond RM Category"
            Top             =   1590
            Width           =   7755
            _ExtentX        =   13679
            _ExtentY        =   503
            MaxLength       =   60
            DataType        =   4
            IdName          =   "WAUTOHIMPTYPNOTINSELD"
         End
         Begin MwfCtl.MWCTL_MED ATXT 
            Height          =   285
            Index           =   183
            Left            =   4230
            TabIndex        =   24
            ToolTipText     =   "Enter Import Types Not To Be Used For Color Stone RM Category"
            Top             =   1875
            Width           =   7755
            _ExtentX        =   13679
            _ExtentY        =   503
            MaxLength       =   60
            DataType        =   4
            IdName          =   "WAUTOHIMPTYPNOTINSELC"
         End
         Begin MwfCtl.MWCTL_MED ATXT 
            Height          =   285
            Index           =   184
            Left            =   4230
            TabIndex        =   31
            ToolTipText     =   "Enter Import Types Not To Be Used For Gold RM Category"
            Top             =   2160
            Width           =   7755
            _ExtentX        =   13679
            _ExtentY        =   503
            MaxLength       =   60
            DataType        =   4
            IdName          =   "WAUTOHIMPTYPNOTINSELG"
         End
         Begin MwfCtl.MWCTL_MED ATXT 
            Height          =   285
            Index           =   185
            Left            =   4230
            TabIndex        =   38
            ToolTipText     =   "Enter Import Types Not To Be Used For Platinum RM Category"
            Top             =   2445
            Width           =   7755
            _ExtentX        =   13679
            _ExtentY        =   503
            MaxLength       =   60
            DataType        =   4
            IdName          =   "WAUTOHIMPTYPNOTINSELP"
         End
         Begin MwfCtl.MWCTL_MED ATXT 
            Height          =   285
            Index           =   186
            Left            =   4230
            TabIndex        =   45
            ToolTipText     =   "Enter Import Types Not To Be Used For Silver RM Category"
            Top             =   2730
            Width           =   7755
            _ExtentX        =   13679
            _ExtentY        =   503
            MaxLength       =   60
            DataType        =   4
            IdName          =   "WAUTOHIMPTYPNOTINSELS"
         End
         Begin MwfCtl.MWCTL_MED ATXT 
            Height          =   285
            Index           =   187
            Left            =   4230
            TabIndex        =   52
            ToolTipText     =   "Enter Import Types Not To Be Used For Accessories RM Category"
            Top             =   3300
            Width           =   7755
            _ExtentX        =   13679
            _ExtentY        =   503
            MaxLength       =   60
            DataType        =   4
            IdName          =   "WAUTOHIMPTYPNOTINSELX"
         End
         Begin MwfCtl.MWCTL_MED ATXT 
            Height          =   285
            Index           =   188
            Left            =   12180
            TabIndex        =   18
            ToolTipText     =   "Enter Date From Which Purchase Vouchers Have To Be Considered For Diamonds"
            Top             =   1590
            Width           =   885
            _ExtentX        =   1561
            _ExtentY        =   503
            MaxLength       =   8
            DataType        =   3
            IdName          =   "WAUTOHIMPAFTD"
         End
         Begin MwfCtl.MWCTL_MED ATXT 
            Height          =   285
            Index           =   189
            Left            =   12180
            TabIndex        =   25
            ToolTipText     =   "Enter Date From Which Purchase Vouchers Have To Be Considered For Color Stones"
            Top             =   1875
            Width           =   885
            _ExtentX        =   1561
            _ExtentY        =   503
            MaxLength       =   8
            DataType        =   3
            IdName          =   "WAUTOHIMPAFTC"
         End
         Begin MwfCtl.MWCTL_MED ATXT 
            Height          =   285
            Index           =   190
            Left            =   12180
            TabIndex        =   32
            ToolTipText     =   "Enter Date From Which Purchase Vouchers Have To Be Considered For Gold"
            Top             =   2160
            Width           =   885
            _ExtentX        =   1561
            _ExtentY        =   503
            MaxLength       =   8
            DataType        =   3
            IdName          =   "WAUTOHIMPAFTG"
         End
         Begin MwfCtl.MWCTL_MED ATXT 
            Height          =   285
            Index           =   191
            Left            =   12180
            TabIndex        =   39
            ToolTipText     =   "Enter Date From Which Purchase Vouchers Have To Be Considered For Platinum"
            Top             =   2445
            Width           =   885
            _ExtentX        =   1561
            _ExtentY        =   503
            MaxLength       =   8
            DataType        =   3
            IdName          =   "WAUTOHIMPAFTP"
         End
         Begin MwfCtl.MWCTL_MED ATXT 
            Height          =   285
            Index           =   192
            Left            =   12180
            TabIndex        =   46
            ToolTipText     =   "Enter Date From Which Purchase Vouchers Have To Be Considered For Silver"
            Top             =   2730
            Width           =   885
            _ExtentX        =   1561
            _ExtentY        =   503
            MaxLength       =   8
            DataType        =   3
            IdName          =   "WAUTOHIMPAFTS"
         End
         Begin MwfCtl.MWCTL_MED ATXT 
            Height          =   285
            Index           =   193
            Left            =   12180
            TabIndex        =   53
            ToolTipText     =   "Enter Date From Which Purchase Vouchers Have To Be Considered For Accessories"
            Top             =   3300
            Width           =   885
            _ExtentX        =   1561
            _ExtentY        =   503
            MaxLength       =   8
            DataType        =   3
            IdName          =   "WAUTOHIMPAFTX"
         End
         Begin MwfCtl.MWCTL_MED ATXT 
            Height          =   285
            Index           =   15
            Left            =   5610
            TabIndex        =   11
            ToolTipText     =   "Specify Whether To Consider Customer Goods Only Or All"
            Top             =   480
            Width           =   345
            _ExtentX        =   609
            _ExtentY        =   503
            DataType        =   4
            IdName          =   "WAUTOUSEONLYCUSTGOODS"
         End
         Begin MwfCtl.MWCTL_MED ATXT 
            Height          =   285
            Index           =   16
            Left            =   2130
            TabIndex        =   10
            ToolTipText     =   "Specify Whether To Match Only Those Records For Which Best Use Purchase Records Are Available"
            Top             =   480
            Width           =   345
            _ExtentX        =   609
            _ExtentY        =   503
            DataType        =   4
            IdName          =   "WAUTOBESTUSEONLY"
         End
         Begin MwfCtl.MWCTL_MED ATXT 
            Height          =   285
            Index           =   17
            Left            =   13440
            TabIndex        =   19
            ToolTipText     =   "Enter Date Upto Which Purchase Vouchers Have To Be Considered For Diamonds"
            Top             =   1590
            Width           =   885
            _ExtentX        =   1561
            _ExtentY        =   503
            MaxLength       =   8
            DataType        =   3
            IdName          =   "WAUTOHIMPBEFD"
         End
         Begin MwfCtl.MWCTL_MED ATXT 
            Height          =   285
            Index           =   18
            Left            =   13440
            TabIndex        =   26
            ToolTipText     =   "Enter Date Upto Which Purchase Vouchers Have To Be Considered For Color Stones"
            Top             =   1875
            Width           =   885
            _ExtentX        =   1561
            _ExtentY        =   503
            MaxLength       =   8
            DataType        =   3
            IdName          =   "WAUTOHIMPBEFC"
         End
         Begin MwfCtl.MWCTL_MED ATXT 
            Height          =   285
            Index           =   19
            Left            =   13440
            TabIndex        =   33
            ToolTipText     =   "Enter Date Upto Which Purchase Vouchers Have To Be Considered For Gold"
            Top             =   2160
            Width           =   885
            _ExtentX        =   1561
            _ExtentY        =   503
            MaxLength       =   8
            DataType        =   3
            IdName          =   "WAUTOHIMPBEFG"
         End
         Begin MwfCtl.MWCTL_MED ATXT 
            Height          =   285
            Index           =   20
            Left            =   13440
            TabIndex        =   40
            ToolTipText     =   "Enter Date Upto Which Purchase Vouchers Have To Be Considered For Platinum"
            Top             =   2445
            Width           =   885
            _ExtentX        =   1561
            _ExtentY        =   503
            MaxLength       =   8
            DataType        =   3
            IdName          =   "WAUTOHIMPBEFP"
         End
         Begin MwfCtl.MWCTL_MED ATXT 
            Height          =   285
            Index           =   21
            Left            =   13440
            TabIndex        =   47
            ToolTipText     =   "Enter Date Upto Which Purchase Vouchers Have To Be Considered For Silver"
            Top             =   2730
            Width           =   885
            _ExtentX        =   1561
            _ExtentY        =   503
            MaxLength       =   8
            DataType        =   3
            IdName          =   "WAUTOHIMPBEFS"
         End
         Begin MwfCtl.MWCTL_MED ATXT 
            Height          =   285
            Index           =   22
            Left            =   13440
            TabIndex        =   54
            ToolTipText     =   "Enter Date Upto Which Purchase Vouchers Have To Be Considered For Accessories"
            Top             =   3300
            Width           =   885
            _ExtentX        =   1561
            _ExtentY        =   503
            MaxLength       =   8
            DataType        =   3
            IdName          =   "WAUTOHIMPBEFX"
         End
         Begin MwfCtl.MWCTL_MED ATXT 
            Height          =   285
            Index           =   23
            Left            =   540
            TabIndex        =   129
            ToolTipText     =   "Enter Rate Variation % From Range For Silver"
            Top             =   3015
            Width           =   765
            _ExtentX        =   1349
            _ExtentY        =   503
            Alignment       =   2
            Mask            =   "-##0.00"
            MaxLength       =   7
            DataType        =   2
            IdName          =   "WAUTOHRTMINUSL"
         End
         Begin MwfCtl.MWCTL_MED ATXT 
            Height          =   285
            Index           =   24
            Left            =   1380
            TabIndex        =   130
            ToolTipText     =   "Enter Rate Variation % To Range For Silver"
            Top             =   3015
            Width           =   765
            _ExtentX        =   1349
            _ExtentY        =   503
            Alignment       =   2
            Mask            =   "-##0.00"
            MaxLength       =   7
            DataType        =   2
            IdName          =   "WAUTOHRTPLUSL"
         End
         Begin MwfCtl.MWCTL_MED ATXT 
            Height          =   285
            Index           =   25
            Left            =   2340
            TabIndex        =   131
            ToolTipText     =   "Enter Pc/ Ct Variation % From Range For Silver"
            Top             =   3015
            Width           =   765
            _ExtentX        =   1349
            _ExtentY        =   503
            Alignment       =   2
            Mask            =   "-##0.00"
            MaxLength       =   7
            DataType        =   2
            IdName          =   "WAUTOHPTRMINUSL"
         End
         Begin MwfCtl.MWCTL_MED ATXT 
            Height          =   285
            Index           =   26
            Left            =   3180
            TabIndex        =   132
            ToolTipText     =   "Enter Pc/ Ct Variation % To Range For Silver"
            Top             =   3015
            Width           =   765
            _ExtentX        =   1349
            _ExtentY        =   503
            Alignment       =   2
            Mask            =   "-##0.00"
            MaxLength       =   7
            DataType        =   2
            IdName          =   "WAUTOHPTRPLUSL"
         End
         Begin MwfCtl.MWCTL_MED ATXT 
            Height          =   285
            Index           =   28
            Left            =   4230
            TabIndex        =   133
            ToolTipText     =   "Enter Import Types Not To Be Used For Silver RM Category"
            Top             =   3015
            Width           =   7755
            _ExtentX        =   13679
            _ExtentY        =   503
            MaxLength       =   60
            DataType        =   4
            IdName          =   "WAUTOHIMPTYPNOTINSELL"
         End
         Begin MwfCtl.MWCTL_MED ATXT 
            Height          =   285
            Index           =   29
            Left            =   12180
            TabIndex        =   134
            ToolTipText     =   "Enter Date From Which Purchase Vouchers Have To Be Considered For Silver"
            Top             =   3015
            Width           =   885
            _ExtentX        =   1561
            _ExtentY        =   503
            MaxLength       =   8
            DataType        =   3
            IdName          =   "WAUTOHIMPAFTL"
         End
         Begin MwfCtl.MWCTL_MED ATXT 
            Height          =   285
            Index           =   31
            Left            =   13440
            TabIndex        =   135
            ToolTipText     =   "Enter Date Upto Which Purchase Vouchers Have To Be Considered For Silver"
            Top             =   3015
            Width           =   885
            _ExtentX        =   1561
            _ExtentY        =   503
            MaxLength       =   8
            DataType        =   3
            IdName          =   "WAUTOHIMPBEFL"
         End
         Begin VB.Label ALBL 
            BackStyle       =   0  'Transparent
            Caption         =   "L"
            BeginProperty Font 
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
            Left            =   150
            TabIndex        =   136
            Top             =   3015
            Width           =   315
         End
         Begin VB.Label ALBL 
            BackStyle       =   0  'Transparent
            Caption         =   "Imports Before"
            BeginProperty Font 
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
            Left            =   13440
            TabIndex        =   128
            Top             =   930
            Width           =   1395
         End
         Begin VB.Label ALBL 
            BackStyle       =   0  'Transparent
            Caption         =   "Match Best Use Only"
            BeginProperty Font 
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
            Left            =   150
            TabIndex        =   127
            Top             =   480
            Width           =   1995
         End
         Begin VB.Label LblMain 
            BackStyle       =   0  'Transparent
            BorderStyle     =   1  'Fixed Single
            Caption         =   "  Auto Custom Matching"
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
            TabIndex        =   126
            Top             =   90
            Width           =   15180
         End
         Begin VB.Label ALBL 
            BackStyle       =   0  'Transparent
            Caption         =   "Rm Ctg List"
            BeginProperty Font 
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
            Left            =   6690
            TabIndex        =   125
            Top             =   480
            Visible         =   0   'False
            Width           =   1095
         End
         Begin VB.Label ALBL 
            BackStyle       =   0  'Transparent
            Caption         =   "X"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Index           =   147
            Left            =   150
            TabIndex        =   124
            Top             =   3300
            Width           =   315
         End
         Begin VB.Label ALBL 
            BackStyle       =   0  'Transparent
            Caption         =   "S"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Index           =   146
            Left            =   150
            TabIndex        =   123
            Top             =   2730
            Width           =   315
         End
         Begin VB.Label ALBL 
            BackStyle       =   0  'Transparent
            Caption         =   "P"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Index           =   145
            Left            =   150
            TabIndex        =   122
            Top             =   2445
            Width           =   315
         End
         Begin VB.Label ALBL 
            BackStyle       =   0  'Transparent
            Caption         =   "G"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Index           =   144
            Left            =   150
            TabIndex        =   121
            Top             =   2160
            Width           =   315
         End
         Begin VB.Label ALBL 
            BackStyle       =   0  'Transparent
            Caption         =   "C"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Index           =   143
            Left            =   150
            TabIndex        =   120
            Top             =   1875
            Width           =   315
         End
         Begin VB.Label ALBL 
            BackStyle       =   0  'Transparent
            Caption         =   "D"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Index           =   142
            Left            =   150
            TabIndex        =   119
            Top             =   1590
            Width           =   315
         End
         Begin VB.Label ALBL 
            BackStyle       =   0  'Transparent
            Caption         =   "Upper"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Index           =   141
            Left            =   3180
            TabIndex        =   118
            Top             =   1215
            Width           =   765
         End
         Begin VB.Label ALBL 
            BackStyle       =   0  'Transparent
            Caption         =   "Lower"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Index           =   140
            Left            =   2340
            TabIndex        =   117
            Top             =   1215
            Width           =   765
         End
         Begin VB.Label ALBL 
            BackStyle       =   0  'Transparent
            Caption         =   "Upper"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Index           =   139
            Left            =   1380
            TabIndex        =   116
            Top             =   1215
            Width           =   765
         End
         Begin VB.Label ALBL 
            BackStyle       =   0  'Transparent
            Caption         =   "Lower"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Index           =   137
            Left            =   540
            TabIndex        =   115
            Top             =   1215
            Width           =   765
         End
         Begin VB.Label ALBL 
            BackStyle       =   0  'Transparent
            Caption         =   "Do Not Use Import Type(s)"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Index           =   136
            Left            =   4230
            TabIndex        =   114
            Top             =   930
            Width           =   2745
         End
         Begin VB.Label ALBL 
            BackStyle       =   0  'Transparent
            Caption         =   "Pc/ Ct Variation (%)"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Index           =   135
            Left            =   2340
            TabIndex        =   113
            Top             =   930
            Width           =   1815
         End
         Begin VB.Label ALBL 
            BackStyle       =   0  'Transparent
            Caption         =   "Rate Variation (%)"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Index           =   134
            Left            =   540
            TabIndex        =   112
            Top             =   930
            Width           =   1695
         End
         Begin VB.Label ALBL 
            BackStyle       =   0  'Transparent
            Caption         =   "Ctg"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Index           =   133
            Left            =   150
            TabIndex        =   111
            Top             =   930
            Width           =   435
         End
         Begin VB.Label ALBL 
            BackStyle       =   0  'Transparent
            Caption         =   "Imports After"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Index           =   138
            Left            =   12060
            TabIndex        =   110
            Top             =   930
            Width           =   1215
         End
         Begin VB.Label ALBL 
            BackStyle       =   0  'Transparent
            Caption         =   "Use Customer Goods Only"
            BeginProperty Font 
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
            Left            =   3090
            TabIndex        =   109
            Top             =   480
            Width           =   2565
         End
      End
      Begin VB.Frame FraHlpSrt 
         Height          =   1785
         Left            =   5790
         TabIndex        =   101
         Top             =   7470
         Width           =   3840
         Begin MwfCtl.MWCTL_MED ATXT 
            Height          =   285
            Index           =   27
            Left            =   2370
            TabIndex        =   88
            ToolTipText     =   "Enter The Sort Option for Custom Matching  Help "
            Top             =   480
            Width           =   1275
            _ExtentX        =   2249
            _ExtentY        =   503
            MaxLength       =   8
            DataType        =   4
            IdName          =   "WSORTBY"
         End
         Begin MwfCtl.MWCTL_MED ATXT 
            Height          =   285
            Index           =   9
            Left            =   2370
            TabIndex        =   90
            ToolTipText     =   "Enter Purchase Rate >=  for Custom Matching  Help "
            Top             =   1050
            Width           =   1275
            _ExtentX        =   2249
            _ExtentY        =   503
            Alignment       =   2
            Mask            =   "#####0.00"
            MaxLength       =   9
            DataType        =   2
            ReCalcOn        =   "WPURRTPRC"
            IdName          =   "WPURRTFR"
         End
         Begin MwfCtl.MWCTL_MED ATXT 
            Height          =   285
            Index           =   12
            Left            =   2370
            TabIndex        =   91
            ToolTipText     =   "Enter Purchase Rate <=  for Custom Matching  Help"
            Top             =   1335
            Width           =   1275
            _ExtentX        =   2249
            _ExtentY        =   503
            Alignment       =   2
            Mask            =   "#####0.00"
            MaxLength       =   9
            DataType        =   2
            ReCalcOn        =   "WPURRTPRC"
            IdName          =   "WPURRTTO"
         End
         Begin MwfCtl.MWCTL_MED ATXT 
            Height          =   285
            Index           =   13
            Left            =   2370
            TabIndex        =   89
            ToolTipText     =   "Enter Percentage For Purchase Rate Range"
            Top             =   765
            Width           =   735
            _ExtentX        =   1296
            _ExtentY        =   503
            Alignment       =   2
            Mask            =   "##0.00"
            MaxLength       =   6
            DataType        =   2
            ReCalcParent    =   "WPURRTFR,WPURRTTO"
            IdName          =   "WPURRTPRC"
         End
         Begin VB.Label ALBL 
            BackStyle       =   0  'Transparent
            Caption         =   "Enter Help Sort  Option"
            BeginProperty Font 
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
            Left            =   120
            TabIndex        =   106
            Top             =   495
            Width           =   2235
         End
         Begin VB.Label ALBL 
            BackStyle       =   0  'Transparent
            Caption         =   "Pur Rt  >="
            BeginProperty Font 
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
            TabIndex        =   105
            Top             =   1065
            Width           =   2085
         End
         Begin VB.Label LblMain 
            BackStyle       =   0  'Transparent
            BorderStyle     =   1  'Fixed Single
            Caption         =   "  Help Sort Option"
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
            Left            =   0
            TabIndex        =   104
            Top             =   90
            Width           =   3840
         End
         Begin VB.Label ALBL 
            BackStyle       =   0  'Transparent
            Caption         =   "Pur Rt  <="
            BeginProperty Font 
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
            Left            =   120
            TabIndex        =   103
            Top             =   1350
            Width           =   2085
         End
         Begin VB.Label ALBL 
            BackStyle       =   0  'Transparent
            Caption         =   "% For Rate Range"
            BeginProperty Font 
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
            Left            =   120
            TabIndex        =   102
            Top             =   780
            Width           =   2085
         End
      End
      Begin VB.Frame FraCP 
         Height          =   1185
         Left            =   6090
         TabIndex        =   56
         Top             =   8055
         Width           =   2790
         Begin MwfCtl.MWCTL_BTN1 CmdCPGo 
            Height          =   495
            Left            =   1950
            TabIndex        =   59
            Top             =   525
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
            Index           =   72
            Left            =   1215
            TabIndex        =   57
            ToolTipText     =   "Enter From Raw Material Category"
            Top             =   435
            Width           =   540
            _ExtentX        =   953
            _ExtentY        =   503
            MaxLength       =   2
            DataType        =   4
            IdName          =   "WCPRMCTGFR"
         End
         Begin MwfCtl.MWCTL_MED ATXT 
            Height          =   285
            Index           =   7
            Left            =   1215
            TabIndex        =   58
            ToolTipText     =   "Enter To Raw Material Category"
            Top             =   720
            Width           =   540
            _ExtentX        =   953
            _ExtentY        =   503
            MaxLength       =   2
            DataType        =   4
            IdName          =   "WCPRMCTGTO"
         End
         Begin VB.Label LblMain 
            BackStyle       =   0  'Transparent
            BorderStyle     =   1  'Fixed Single
            Caption         =   "  Copy Customs Rm Code Details"
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
            TabIndex        =   87
            Top             =   90
            Width           =   2820
         End
         Begin VB.Label ALBL 
            BackStyle       =   0  'Transparent
            Caption         =   "             To"
            BeginProperty Font 
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
            TabIndex        =   82
            Top             =   720
            Width           =   1185
         End
         Begin VB.Label ALBL 
            BackStyle       =   0  'Transparent
            Caption         =   "Rm Ctg Fr"
            BeginProperty Font 
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
            Left            =   120
            TabIndex        =   81
            Top             =   435
            Width           =   1185
         End
      End
      Begin VB.Frame FraDelRm 
         Height          =   1185
         Left            =   7440
         TabIndex        =   139
         Top             =   8040
         Width           =   3330
         Begin MwfCtl.MWCTL_BTN1 CmdDelRmGo 
            Height          =   495
            Left            =   2580
            TabIndex        =   141
            Top             =   495
            Width           =   585
            _ExtentX        =   1032
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
            Index           =   33
            Left            =   990
            TabIndex        =   140
            ToolTipText     =   "Enter RmCtg Combination to be deleted"
            Top             =   600
            Width           =   1335
            _ExtentX        =   2355
            _ExtentY        =   503
            MaxLength       =   8
            DataType        =   4
            IdName          =   "WDELRMCTG"
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
            Height          =   285
            Index           =   16
            Left            =   120
            TabIndex        =   143
            Top             =   615
            Width           =   735
         End
         Begin VB.Label LblMain 
            BackStyle       =   0  'Transparent
            BorderStyle     =   1  'Fixed Single
            Caption         =   "  Delete Customs Rm Code Details"
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
            TabIndex        =   142
            Top             =   90
            Width           =   3660
         End
      End
   End
End
Attribute VB_Name = "EmrFrmCustMaint"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'---- General Comments on the Entry --------'
' 1 Menu 'InCustmMtch'
'   The goods which are going out of the factory due to invoicing have to be matched against
'   the Custom Purchase entries which has to be shown to the Customs authorities
'   This Entry is used for Matching the Invoice Rm records with the Custom Purchase records
'   for an Invoice.
'
' X-'', A-InvZRm(GrdIZ), B-TzndZ(GrdTxndZ)
'
' CmdCp to copy records into the GrdIz grid from the InvRm table
'   The records of InvRm are grouped on CustomCode And CustomPureRate.
'   FraCp, CmdCpGo, wCpRmCtgFr, wCpRmCtgTo
'   Private sub CopyInvZRm
'
' CmdIZR, FraIZR, GrdIZR, wIZRRmTotQty, wIZRRmTotWt
'   For one record in GrdIZ, the Rm Details Option shows all the the Inv Design Sr records
'   which uses this Custom Code. The records are shown InvSr and Rm Code wise.
'   The corresponding Rm Qty, Rm Wt, Export Order Number & Design Code are also shown
'   Private Subs : RmDetails
'
' If the Date of the current invoice is less than that of the lock date then the entry is
'   not allowed to be edited or deleted. The user can only read the entry. (This option
'   has been disabled for the time being)
'
'--------------------

Option Explicit
Dim moCn As MwfLib.MDOConnection   '*** (Bef speed) bef 24/11/06
Dim mIzTc As String, mIzYy As String, mIzChr As String
Dim mTcTyp As String

'*** Declare the variables ms_LockDt, mf_DtLocked for the Lock Date Check
Dim ms_LockDt As String
Dim ms_FutureDt As String
Dim mf_DtLocked As Boolean

Dim ms_PurgeDt As String        '****** Sachin 3.02 31-12-07

'*** For the Toggle Options of Command Buttons
Enum en_InvZFra2
  InvRmDet = 0
  Cpy = 1
  HlpSrt = 2
  AutoCustmMtch = 3
  DelRm = 4     ' ****** Manali 3.5.0 - 10/11/08 - Delete Rm Details on Button Click
End Enum

Dim ms_InCls As String
'*** Arr_Cv stores the causes validation state, Arr_Ed the Ena/Disa state of Cmd Buttons
Dim mArr_Cv As Variant, mArr_Ed As Variant
Dim wFlag As Boolean
Dim ms_LocalExp As String   '*** Jen (2.12 patch)

'*** (Jen 2.13)
Dim mb_ShowSizeStkRt As Boolean
Dim md_CnvFct As Double, ms_MulDiv As String
Dim md_RsCnvFct As Double, ms_RsMulDiv As String
Dim mb_ShowSize As Boolean    '*** (07/11/06)
'*** (Jen 2.13)
' ****** Sachin 2.14.0 - [24-11-06] ******
Dim mb_MultiComp As Boolean
' ****** Sachin 2.14.0 - [24-11-06] ******
' ***** Manali 3.6.0 - 31/12/09 - Notional Instead od Purchase Rate For DTA
Dim md_GldFiness As Double, md_PlFiness As Double, md_SlFiness As Double, md_PdFiness As Double, ms_NotionRt As String
Dim md_LmgCst As Double, md_LmpCst As Double, md_LmsCst As Double, md_LmlCst As Double
' ***** Manali 3.6.0 - 31/12/09 - Notional Instead od Purchase Rate For DTA

Dim mb_InCompInv As Boolean  '***** Manali 3.9.0 - Incomplete invoice Design


Private Sub ADC_Load()
  '*** Redim the arrays mArr_Cv and mArr_Ed to the number of command buttons
  '*** Set the First Key Control property to wIzChr
  '*** Set the First Non Key Control property to CmdCp
  '*** Set the Child Property Of ADC to GrdIZ
  '*** Set the Child Property Of GrdIZ to GrdTxndZ
  '*** Set the Hot Keys for the Grids GrdIZ, GrdTxndZ and GrdIZD as 'A', 'B' and 'C'
  '*** Set the Previous and Next Control Properties for the Grids
      'GrdIZ,GrdTxndZ and GrdIZD
  '*** Bring the frame FraNKeyAll to the top.
  '*** Get the default values for Tc and Yy from param and head files when the form is loaded
  
  'ReDim mArr_Cv(4), mArr_Ed(4)
  ReDim mArr_Cv(5), mArr_Ed(5)      ' ****** Manali 3.5.0 - 10/11/08 - Delete Rm Details on Button Click
  
  Set moCn = adc.Connection    '*** (Bef speed) bef 24/11/06
  Set adc.FirKeyCtl = adc("wIzChr")
  Set adc.FirNKeyCtl = CmdCP
  adc.Child = "GrdIZ"
  GrdIZ.Child = "GrdTxndZ"
  Set GrdIZ.PrevCtl = adc
  Set GrdIZ.NextCtl = adc
  Set GrdTxndZ.PrevCtl = GrdIZ
  Set GrdTxndZ.NextCtl = GrdIZ
  Set GrdIZD.NextCtl = GrdIZ
  Set GrdIZR.NextCtl = CmdIZR
  GrdIZ.HotKey = "A"
  GrdTxndZ.HotKey = "B"
  GrdIZD.HotKey = "C"
  
  FraNKeyAll.ZOrder (0)
  mTcTyp = ctInTcTyp
 
  
  mIzTc = moCn.GetFldVal("Select PValue from Param where PTyp= 'TCTYP' and " + _
               "PMCd='" + mTcTyp + "' ")
  mIzYy = moCn.GetFldVal("Select hYy from Head where HCoCd= '" + gs_CoCd + "' and HCd= '" + ctSelfCmCd + "'")
  
  '*** (Jen 2.13)
  mb_ShowSize = moCn.RecSeek("Select 'x' from Param where PTyp= 'SYSOPT' and PmCd= 'ZRMDET' and PValue= 'Y'")
  mb_ShowSizeStkRt = moCn.RecSeek("Select 'x' from Head where HCoCd= '" + gs_CoCd + "' and " + "HCd= '" + ctSelfCmCd + "' and hSizeStkRt= 'Y'")
  
  '*** (07/11/06)
  If mb_ShowSize = True Then
  '*** (07/11/06)
  
    If mb_ShowSizeStkRt = False Then
      GrdIZD.ColProp("wIzdIrRmSz").Style = fgcinvisible
    Else
      GrdIZD.ColProp("wIzdIrRmSz").Style = fgcNormal
    End If
    
  '*** (07/11/06)
    GrdIZD.ColProp("wIzdRmRSz").Style = fgcNormal
  Else
    GrdIZD.ColProp("wIzdIrRmSz").Style = fgcinvisible
    GrdIZD.ColProp("wIzdRmRSz").Style = fgcinvisible
  End If
  '*** (07/11/06)
  '*** (Jen 2.13)
  
  '*** (Jen 3.2.0)
  If moCn.GetFldVal("Select HSlvModYN From Head Where HCoCd='" + gs_CoCd + "' And HCd='" + ctSelfCmCd + "'") = "N" Then
    adc("wIzLmsSal").Visible = False: LblLmsSal.Visible = False
  End If
  '*** (Jen 3.2.0)
  
  ' ***** Manali 3.03 - 23/06/08 - Pd Mod - 'L' added
  If moCn.GetFldVal("Select HPdModYN From Head Where HCoCd='" + gs_CoCd + "' And HCd='" + ctSelfCmCd + "'") = "N" Then
    adc("wIzLmlSal").Visible = False: LblLmlSal.Visible = False
  End If
  ' ***** Manali 3.03 - 23/06/08 - Pd Mod - 'L' added
  
  '*** (Jen 3.01)
  If moCn.RecSeek("Select 'x' from Head where HCoCd= '" + gs_CoCd + "' and " + "HCd= '" + ctSelfCmCd + "' and HAutoCustmMtch= 'Y'") Then
    CmdAuto.Visible = True
  Else
    CmdAuto.Visible = False
  End If
  '*** (Jen 3.01)
  
  
  '****** Sachin 3.02.0 - Hiding the Partition Key Column based on Partition Selected
  adc("WIZPRTKEY").Visible = IIF(gs_Partition = ctCurrPrtn, False, True)
  GrdIZ.ColProp("IzPrtKey").Style = IIF(gs_Partition = ctCurrPrtn, fgcinvisible, fgcReadOnly)
  GrdTxndZ.ColProp("TzPrtKey").Style = IIF(gs_Partition = ctCurrPrtn, fgcinvisible, fgcReadOnly)
  
End Sub
Private Sub ADC_KeyWhen()
  '*** Disable All command buttons
  '*** Hide all frames except Frame FraNKeyAll
  '*** Delete all rows from the working grid GrdIZD
  '*** Set the default values for wIzTc, wIzYy, wIzChr from the previous entry
  '*** Reset the Flag option to flase.
      'When the focus leaves GrdTxndZ, this option checks that if atleast one match
      'record exists for a record of GrdIZ then the matching for that GrdIZ row has
      'to be completed for the entire Weight (Half matching is not allowed)
      'The alternative option is no matching should be done for that row
  
  mf_DtLocked = False
  Call EnaDisaCmds(True)
  Call HideAllFras
  GrdIZD.Rows = 1

  If mIzTc <> "" Then adc("wIzTc") = mIzTc
  If mIzYy <> "" Then adc("wIzYy") = mIzYy
  If mIzChr <> "" Then adc("wIzChr") = mIzChr
  wFlag = False

  '****** Sachin 2.14.0 - [24-11-2006] - [Multi-Company] ******
  Me.Caption = GetFrmCaption(gs_CoCd, "Invoice Custom Matching")
  '****** Sachin 2.14.0 - [24-11-2006] - [Multi-Company] ******

End Sub

Private Sub ADC_KeyValid(Cancel As Boolean, ErrMsg As String)
  '*** Force the user to enter a valid Invoice Number
  
  If adc("wIzYy") = "" Then Cancel = True: ErrMsg = "Enter Invoice Voucher Year": Exit Sub
  If adc("wIzChr") = "" Then Cancel = True: ErrMsg = "Enter Invoice Voucher Character": Exit Sub
  If adc("wIzNo") = 0 Then Cancel = True: ErrMsg = "Enter Invoice Voucher Number": Exit Sub
  
  '****** Sachin 3.02 31-12-07 - Coresponding Partition (PrtKey Added)
  Cancel = Not moCn.RecSeek("Select InNo From InvHd Where " + _
           "InCoCd= '" + gs_CoCd + "' and InTc= '" + adc("wIzTc") + "' " + _
           "And InYy= '" + adc("wIzYy") + "' and InChr= '" + adc("wIzChr") + "' and " + _
           "InNo= " + CStr(adc("wIzNo")) + " and InPrtKey=(Case When '" + adc("WIZPRTKEY") + "'='" + ctCurrPrtn + "' Then '" + ctCurrPrtn + "' Else InPrtKey End ) ")
  If Cancel = True Then ErrMsg = "Invalid Invoice Voucher": Exit Sub
  
  '*** (Jen 2.13) These variables have to be provided values before NonKeyWhen event
  Dim wRsInvDet As MwfLib.MDORowSet
  
  '****** Sachin 3.02 31-12-07 - Coresponding Partition (PrtKey Added)
  Set wRsInvDet = moCn.OpenRes("Select InCnvFct, InMulDiv, inCnvRt, InRsMulDiv " + _
                  "from InvHd join CustMst on CmCtg= 'C' and CmCd= InCmCd " + _
                  "where InCoCd= '" + gs_CoCd + "' and " + _
                  "InTc= '" + adc("wIzTc") + "' and InYy= '" + adc("wIzYy") + "' and " + _
                  "InChr= '" + adc("wIzChr") + "' and InNo= " + CStr(adc("wIzNo")) + " and InPrtKey=(Case When '" + adc("WIZPRTKEY") + "'='" + ctCurrPrtn + "' Then '" + ctCurrPrtn + "' Else InPrtKey End) ")
  With wRsInvDet
    If Not (.EOF Or .BOF) Then
      md_CnvFct = !InCnvFct
      ms_MulDiv = UCase(Trim(!InMulDiv))
      md_RsCnvFct = !InCnvRt
      ms_RsMulDiv = UCase(Trim(!InRsMulDiv))
    End If
  End With
  '*** (Jen 2.13)
  
  ' ****** Manali 3.9.0 - Following check removed
  '******** Manali 3.2.2 - IrAWt= Bag Grs Wt and IdQty= InvFGd bag Qty check
'''  Dim wi_IdSr As Integer
'''  wi_IdSr = moCn.GetFldVal("Select IdSr From InvDsg " + _
'''                "left outer join " + _
'''                "(Select IrPrtKey, IrIdIdNo, IrCoCd, IrTc, IrYy, IrChr, IrNo, IrSr, " + _
'''                "IsNull(sum(IrRmAWt/(case when IrRmCtg in ('D','C') then 5 else 1 End)), 0) as qIrGrsWt from InvRm " + _
'''                "Group By IrPrtKey, IrIdIdNo, IrCoCd, IrTc, IrYy, IrChr, IrNo, IrSr) Ir " + _
'''                "on IdCoCd=IrCoCd and IdTc=IrTc and IdYy=IrYy and IdChr=IrChr and IdNo=IrNo and IdSr=IrSr and IdPrtKey=IrPrtKey and IdIdNo=IrIdIdNo" + _
'''                " left outer join " + _
'''                "(Select IfPrtKey, IfCoCd, IfTc, IfYy, IfChr, IfNo, IfSr, IsNull(sum(IfGrsWt), 0) as qIfGrsWt, IsNull(sum(IfQty), 0) as qIfQty from InvFgd " + _
'''                "Group By IfPrtKey, IfCoCd, IfTc, IfYy, IfChr, IfNo, IfSr) Inf " + _
'''                "on IrCoCd=IfCoCd and IrTc=IfTc and IrYy=IfYy and IrChr=IfChr and IrNo=IfNo and IrSr=IfSr and IrPrtKey=IfPrtKey " + _
'''                "where IdCoCd= '" + gs_CoCd + "' and " + _
'''                "IdTc= '" + adc("wIzTc") + "' and IdYy= '" + adc("wIzYy") + "' and IdChr= '" + adc("wIzChr") + "' and " + _
'''                "IdNo= " + CStr(adc("wIzNo")) + " and IdPrtKey=(Case When '" + adc("WIZPRTKEY") + "'='" + ctCurrPrtn + "' Then '" + ctCurrPrtn + "' Else IdPrtKey End) " + _
'''                "and (abs(cast(cast(IsNull(qIrGrsWt, 0) as decimal(14,3)) as decimal(14,2)) - cast(cast(IsNull(qIfGrsWt, 0) as decimal (14,3)) as decimal(14,2))) > 0.9 " + _
'''                " or cast(IsNull(qIfQty ,0) as decimal(14,0)) <> cast(IsNull(IdQty ,0) as decimal (14,0))) " + _
'''                "Order By IdCoCd, IdTc, IdYy, IdChr, IdNo, IdSr ")

  '***** Manali 3.9.0 - Incomplete invoice Design
  mb_InCompInv = moCn.RecSeek("Select IdSr From InvDsg " + _
                "left outer join " + _
                "(Select IrPrtKey, IrIdIdNo, IrCoCd, IrTc, IrYy, IrChr, IrNo, IrSr, " + _
                "IsNull(sum(IrRmAWt/(case when IrRmCtg in ('D','C') then 5 else 1 End)), 0) as qIrGrsWt from InvRm " + _
                "Group By IrPrtKey, IrIdIdNo, IrCoCd, IrTc, IrYy, IrChr, IrNo, IrSr) Ir " + _
                "on IdCoCd=IrCoCd and IdTc=IrTc and IdYy=IrYy and IdChr=IrChr and IdNo=IrNo and IdSr=IrSr and IdPrtKey=IrPrtKey and IdIdNo=IrIdIdNo" + _
                " left outer join " + _
                "(Select IfPrtKey, IfCoCd, IfTc, IfYy, IfChr, IfNo, IfSr, IsNull(sum(IfGrsWt), 0) as qIfGrsWt, IsNull(sum(IfQty), 0) as qIfQty from InvFgd " + _
                "Group By IfPrtKey, IfCoCd, IfTc, IfYy, IfChr, IfNo, IfSr) Inf " + _
                "on IrCoCd=IfCoCd and IrTc=IfTc and IrYy=IfYy and IrChr=IfChr and IrNo=IfNo and IrSr=IfSr and IrPrtKey=IfPrtKey " + _
                "where IdCoCd= '" + gs_CoCd + "' and " + _
                "IdTc= '" + adc("wIzTc") + "' and IdYy= '" + adc("wIzYy") + "' and IdChr= '" + adc("wIzChr") + "' and " + _
                "IdNo= " + CStr(adc("wIzNo")) + " and IdPrtKey=(Case When '" + adc("WIZPRTKEY") + "'='" + ctCurrPrtn + "' Then '" + ctCurrPrtn + "' Else IdPrtKey End) " + _
                "and (abs(cast(cast(IsNull(qIrGrsWt, 0) as decimal(14,3)) as decimal(14,2)) - cast(cast(IsNull(qIfGrsWt, 0) as decimal (14,3)) as decimal(14,2))) > 0.9 " + _
                " or cast(IsNull(qIfQty ,0) as decimal(14,0)) <> cast(IsNull(IdQty ,0) as decimal (14,0))) " + _
                "/*Order By IdCoCd, IdTc, IdYy, IdChr, IdNo, IdSr*/ ")


' ****** Manali - 19/09/08                 [check EmrFrmInvDsg - adc_nonkeyvalid]
'                "and (cast(cast(IsNull(qIrGrsWt, 0) as decimal(14,3)) as decimal(14,2))  <> cast(cast(IsNull(qIfGrsWt, 0) as decimal (14,3)) as decimal(14,2)) " + _
'                " or cast(IsNull(qIfQty ,0) as decimal(14,0)) <> cast(IsNull(IdQty ,0) as decimal (14,0))) " + _
'                "Order By IdCoCd, IdTc, IdYy, IdChr, IdNo, IdSr ")

    'If wi_IdSr > 0 Then
    '***** Manali 3.9.0 - Incomplete invoice Design
    If mb_InCompInv Then
      ''Cancel = True: ErrMsg = "Invoice Contains Discrepancy, Call MICROWAY Immediately": Exit Sub
      ' ***** Manali 3.9.0
      MsgBox "This Invoice Contains Discrepancy", vbOKOnly, "Error"
    End If
  
End Sub
Private Sub ADC_NonKeyWhen(Cancel As Boolean, ErrMsg As String)
  '*** Get the Lock Date from Param('LOCKDT', 'LOCKDT') and store in the form level variable ms_LockDt
      'If ms_LockDt is blank then store '01/01/80' in ms_LockDt (the lock date option is currently disabled in the invoice entry but can be enabled any time)
  '*** Store the Tc, Yy, Chr values for the next entry
  '*** Enable all the command buttons
  '*** Give values to the working fields ADC(wIzDt), ADC(wIzCmCd), ADC(wIzExpNo) and
      'ADC(wIzLmgSal) from the InvHd table
  '*** If the Invoice Date is < ms_LockDt then mf_DtLocked= true (i.e the entry
      'If mf_DtLocked = true then the entry is locked. i.e. the user cannot add, edit or
      'delete records in the grids GrdIZ and GrdTxndZ
  '*** If Invoice is closed then do not allow the user to enter any values in the grids
      'GrdIZ and GrdTxndZ
  
'  ms_LockDt = moCn.GetFldVal("Select PDesc225 from Param where PTyp= 'LOCKDT' and PMCd= 'LOCKDT'")
'  ms_FutureDt = moCn.GetFldVal("Select PDesc225 from Param where PTyp= 'LOCKDT' and PMCd= 'FUTUREDT'")
        If adc.IsNew And Not moCn.RecSeek("select vPSCd from vParam where " _
                         + "vPCoCd= '" + gs_CoCd + "' and vPTyp= 'CHR' and " _
                         + "vPMCd = '" + adc("wIzTc") + "' and vPSCd= '" + adc("wIzChr") + "' and vPValidYn<>'N'") Then
       Cancel = True: ErrMsg = "Voucher Character No More Valid": adc.AllowSave = False: Exit Sub
      Else
        adc.AllowSave = True
      End If

  Call GetLockDt(gs_CoCd, "TC", mIzTc, ms_LockDt, ms_FutureDt)
  
  '****** Sachin 3.02 31-12-07
  ms_PurgeDt = moCn.GetFldVal("Select PDesc225 from Param where PTyp= 'PURGEDT' and PMCd= 'PURGEDT'")
  
  '*** For the time being do not put Lock Date Check in the Invoice Entry
  '### If ms_LockDt = "" Then ms_LockDt = "01/01/80"
  'If ms_FutureDt = "" Then ms_FutureDt = "01/01/80"
  ms_LockDt = "01/01/80": ms_FutureDt = IIF(UCase(MWLib.DtSetting) = "MM/DD/YY", "12/31/28", "31/12/28")  '6.1
  
  If ms_PurgeDt = "" Then ms_PurgeDt = "01/01/80"       '****** Sachin 3.02 31-12-07
  Dim wRsInvDet As MwfLib.MDORowSet
  
  Call EnaDisaCmds(False)
      
  mIzTc = adc("wIzTc")
  mIzYy = adc("wIzYy")
  mIzChr = adc("wIzChr")
  
  '*** Set The Visible, Enable Properties For All The Fields ***
  'CmdIZR.Enabled = False
' **** Zubin 211 **** '
  CmdIZR.Enabled = False
' **** Zubin 211 **** '
  
  '*** (Jen 2.12)
  ms_LocalExp = moCn.GetFldVal("Select hExpLocal from Head where HCoCd= '" + gs_CoCd + "' " + _
                               "and HCd= '" + ctSelfCmCd + "' ")
  '*** (Jen 2.12)
  
  '****** Sachin 3.02 31-12-07 - Coresponding Partition (PrtKey Added)
  ' ***** Manali 3.03 - 23/06/08 - Pd Mod - 'LML' included
  ' ***** Manali 3.6.0 - 31/12/09 - Notional Instead of Purchase Rate For DTA
  Set wRsInvDet = moCn.OpenRes("Select InPrtKey, InDt, InCmCd, InExpNo, InLmgSal, InLmpSal, InLmsSal, InLmlSal, InCls, " + _
                  "InCnvFct, InMulDiv, CmCurCd, inCnvRt, InRsMulDiv, InHGldFiness, InHPlFiness, InHSlFiness, InHPdFiness, " + _
                  "InNotionRtYn,InLmgCst, InLmpCst, InLmsCst, InLmlCst " + _
                  "from InvHd join CustMst on CmCtg= 'C' and CmCd= InCmCd " + _
                  "where InCoCd= '" + gs_CoCd + "' and " + _
                  "InTc= '" + adc("wIzTc") + "' and InYy= '" + adc("wIzYy") + "' and " + _
                  "InChr= '" + adc("wIzChr") + "' and InNo= " + CStr(adc("wIzNo")) + " and InPrtKey=(Case When '" + gs_Partition + "'='" + ctCurrPrtn + "' Then '" + ctCurrPrtn + "' Else InPrtKey End) ")
  With wRsInvDet
    If Not (.EOF Or .BOF) Then
        
      adc("wIzPrtKey") = !InPrtKey      '****** Sachin 3.02 31-12-07
      adc("wIzDt") = !InDt
      adc("wIzCmCd") = !InCmCd
      adc("wIzExpNo") = !InExpNo
      adc("wIzLmgSal") = !InLmgSal
      adc("wIzLmpSal") = !InLmpSal
      ' ****** Sachin 2.12 - 06/12/05 - SJM
      adc("wIzLmsSal") = !InLmsSal
      ' ****** Sachin 2.12 - 06/12/05 - SJM
      ' ***** Manali 3.03 - 23/06/08 - Pd Mod - 'LML'
      adc("wIzLmlSal") = !InLmlSal
      ' ***** Manali 3.03 - 23/06/08 - Pd Mod - 'LML'
      '*** (Jen 2.13)
      ms_InCls = !InCls
      adc("wIzCmCurCd") = !CmCurCd
      adc("wIzCnvFct") = !InCnvFct
      adc("wIzMulDiv") = !InMulDiv
      'md_RsCnvFct = !inCnvRt
      'ms_RsMulDiv = UCase(Trim(!InRsMulDiv))
      '*** (Jen 2.13)
      ' ***** Manali 3.6.0 - 31/12/09 - Notional Rate Instead of Purchase Rate For DTA
      md_LmgCst = !InLmgCst
      md_LmpCst = !InLmpCst
      md_LmsCst = !InLmsCst
      md_LmlCst = !InLmlCst
      md_GldFiness = !InHGldFiness
      md_PlFiness = !InHPlFiness
      md_SlFiness = !InHSlFiness
      md_PdFiness = !InHPdFiness
      ms_NotionRt = !InNotionRtYn
      ' ***** Manali 3.6.0 - 31/12/09 - Notional Rate Instead of Purchase Rate For DTA
    End If
  End With
  
  mf_DtLocked = IIF((adc("wIzDt") < CDate(ms_LockDt)) _
                                Or _
                   (adc("wIzDt") > CDate(ms_FutureDt)) _
                                Or _
                  (adc("wIzDt") <= CDate(ms_PurgeDt)), _
                True, False)
  
  If mf_DtLocked = True Then
    GrdIZ.AllowAdd = False: GrdIZ.AllowDelete = False
    GrdTxndZ.AllowAdd = False: GrdTxndZ.AllowDelete = False
  End If
  
  '*** (Bef 2.13)
  'ms_InCls = moCn.GetFldVal("Select InCls from InvHd where " + _
                        "InCoCd= '" + gs_CoCd + "' and " + _
                        "InTc= '" + ADC("wIzTc") + "' and " + _
                        "InYy= '" + ADC("wIzYy") + "' and " + _
                        "InChr= '" + ADC("wIzChr") + "' and " + _
                        "InNo= " + CStr(ADC("wIzNo")))
  '*** (Bef 2.13)

  '****** Sachin 3.02 31-12-07 - purge date condition added below
  If ms_InCls = "Y" Or adc("wIzPrtKey") <> ctCurrPrtn Or mf_DtLocked Then
    GrdIZ.AllowAdd = False: GrdIZ.AllowDelete = False
    GrdTxndZ.AllowAdd = False: GrdTxndZ.AllowDelete = False
  Else
    If mf_DtLocked = False And adc("wIzPrtKey") = ctCurrPrtn Then
      GrdIZ.AllowAdd = True: GrdIZ.AllowDelete = True
      GrdTxndZ.AllowAdd = True: GrdTxndZ.AllowDelete = True
    End If
  End If
  '********Geeta*********Emr207*********15/07/04
  adc("wSortBy") = moCn.GetFldVal("Select HSrtBasePurVch from Head where HCoCd='" + gs_CoCd + "' and HCd='" + ctSelfCmCd + "'")

  '*** (Jen 2.13 speed)
  adc("wPurRtPrc") = 10
  '*** (Jen 2.13 speed)

  '***** Manali 3.9.0 - Incomplete invoice Design
  If mb_InCompInv Then
    GrdIZ.AllowAdd = False
    GrdTxndZ.AllowAdd = False
  End If
End Sub
Private Sub ADC_NonKeyValid(pr_Cancel As Boolean, pr_ErrMsg As String)
  '*** Do not allow the user to save the entry if only partial matching has been done for
      'any of the records of InvZRm
      
  '3.11.2 - Commented as this check is not reqd
  If wFlag = True Then pr_Cancel = True: pr_ErrMsg = "Weight Matching Has Not Been Done Properly"
  
  '***** Manali 3.9.0 - Incomplete invoice Design
  If moCn.RecSeek("Select IdSr From InvDsg " + _
                "left outer join " + _
                "(Select IrPrtKey, IrIdIdNo, IrCoCd, IrTc, IrYy, IrChr, IrNo, IrSr, " + _
                "IsNull(sum(IrRmAWt/(case when IrRmCtg in ('D','C') then 5 else 1 End)), 0) as qIrGrsWt from InvRm " + _
                "Group By IrPrtKey, IrIdIdNo, IrCoCd, IrTc, IrYy, IrChr, IrNo, IrSr) Ir " + _
                "on IdCoCd=IrCoCd and IdTc=IrTc and IdYy=IrYy and IdChr=IrChr and IdNo=IrNo and IdSr=IrSr and IdPrtKey=IrPrtKey and IdIdNo=IrIdIdNo" + _
                " left outer join " + _
                "(Select IfPrtKey, IfCoCd, IfTc, IfYy, IfChr, IfNo, IfSr, IsNull(sum(IfGrsWt), 0) as qIfGrsWt, IsNull(sum(IfQty), 0) as qIfQty from InvFgd " + _
                "Group By IfPrtKey, IfCoCd, IfTc, IfYy, IfChr, IfNo, IfSr) Inf " + _
                "on IrCoCd=IfCoCd and IrTc=IfTc and IrYy=IfYy and IrChr=IfChr and IrNo=IfNo and IrSr=IfSr and IrPrtKey=IfPrtKey " + _
                "where IdCoCd= '" + gs_CoCd + "' and " + _
                "IdTc= '" + adc("wIzTc") + "' and IdYy= '" + adc("wIzYy") + "' and IdChr= '" + adc("wIzChr") + "' and " + _
                "IdNo= " + CStr(adc("wIzNo")) + " and IdPrtKey=(Case When '" + adc("WIZPRTKEY") + "'='" + ctCurrPrtn + "' Then '" + ctCurrPrtn + "' Else IdPrtKey End) " + _
                "and (abs(cast(cast(IsNull(qIrGrsWt, 0) as decimal(14,3)) as decimal(14,2)) - cast(cast(IsNull(qIfGrsWt, 0) as decimal (14,3)) as decimal(14,2))) > 0.9 " + _
                " or cast(IsNull(qIfQty ,0) as decimal(14,0)) <> cast(IsNull(IdQty ,0) as decimal (14,0))) " + _
                "/*Order By IdCoCd, IdTc, IdYy, IdChr, IdNo, IdSr*/ ") Then
    If GrdIZ.Rows > 1 Then
      pr_Cancel = True: pr_ErrMsg = "This Invoice Contains Discrepancy. Delete Custom Matching."
    Else
    
    End If
  End If
  
  Dim wIzSr As Integer
  If moCn.RecSeek(" Select 'x' From Param where PTyp = 'MTCHQTY' and PMCd = 'MTCHQTY' and PSCd ='' and PValue = 'N'") Then
    Exit Sub
  End If
  If moCn.RecSeek(" Select 'x' Where Cast('" + CStr(adc("WIZDT")) + "'  as smalldatetime) > '20/02/15'") Then
       wIzSr = moCn.GetFldVal(" Select IzSr from InvZRm Join RmMst On IzRmZCd=RmCd and RmZ='Y' " + _
                              " where RmCtg in ('D','C') " + _
                              " and IzCoCd= '" + gs_CoCd + "' and " + _
                              " IzTc= '" + adc("wIzTc") + "' and IzYy= '" + adc("wIzYy") + "' and IzChr= '" + adc("wIzChr") + "' and " + _
                              " IzNo= " + CStr(adc("wIzNo")) + " and IzPrtKey=(Case When '" + adc("WIZPRTKEY") + "'='" + ctCurrPrtn + "' Then '" + ctCurrPrtn + "' Else IzPrtKey End) " + _
                              " and (Select Count('x') From TxndZ M where IzCoCd=TzCoCd and IzTc=TzTc and IzYy=TzYy and IzChr=TzChr and IzNo=TzNo and IzSr=TzSr)>1 " + _
                              " and (Select Count('x') From TxndZ M where IzCoCd=TzCoCd and IzTc=TzTc and IzYy=TzYy and IzChr=TzChr and IzNo=TzNo and IzSr=TzSr)>IzRmQty ")
    
    If wIzSr > 0 Then pr_Cancel = True: pr_ErrMsg = " Matching for SrNo : " + Str(wIzSr) + " more than Stone Qty."
  End If
End Sub



Private Sub ATXT_FldWhen(Index As Integer, ByVal IdName As String, Cancel As Boolean, ErrMsg As String)
  '*** Cannot edit any field if the entry has been locked
  '*** Set Helps For Each Field
  
  If mf_DtLocked = True Then
    Cancel = True: ErrMsg = "Cannot Edit Locked Entry": Exit Sub
  End If
  
  Select Case UCase(IdName)
  Case Is = UCase("wIzTc")
    Call HlpList.PMCd("TC", "'" + mTcTyp + "'")
  Case Is = UCase("wIzYy")
    Call HlpList.PMCd("YY")
  Case Is = UCase("wIzChr")
' **** Zubin 211 **** '
    'Call HlpList.vPSCd(gs_CoCd, "CHR", ADC("wIzTc"))
    Call HlpList.vPSCd(gs_CoCd, "CHR", adc("wIzTc"), "Y")
' **** Zubin 211 **** '
  Case Is = UCase("wIzNo")
    Call HlpList.InNo(gs_CoCd, adc("wIzTc"), adc("wIzYy"), adc("wIzChr"), , True)
  Case Is = UCase("wCpRmCtgFr"), UCase("wCpRmCtgTo")
    Call HlpList.PMCd("RMCTG")
  '************Geeta**********Emr207********
  Case Is = UCase("wSortBy")
    Call HlpList.PMCd("SRTPRHLP")
  
  '*** (Bef 3.1.1)
  ''*** (Jen 3.01)
  'Case Is = UCase("wAutoCpyOpt")
  '  Call HlpList.PMCd("CPYRT")
  ''*** (Jen 3.01)
  '*** (Bef 3.1.1)
  
  '****** (Jen 3.2.0) added wAutoBestUseOnly
  '*** (Jen 3.1.1)
  Case Is = UCase("wAutoUseOnlyCustGoods"), UCase("wAutoBestUseOnly")
    Call HlpList.PMCd("YN")
  Case Is = UCase("wAutoHImpTypNotInSelD"), UCase("wAutoHImpTypNotInSelC"), UCase("wAutoHImpTypNotInSelG"), _
            UCase("wAutoHImpTypNotInSelP"), UCase("wAutoHImpTypNotInSelS"), UCase("wAutoHImpTypNotInSelX")
    HlpList.MultiSelect = True
    Call HlpList.PMCd("IMPTYP")
  '*** (Jen 3.1.1)
  
  ' ***** Manali 3.5.0 - Delete Records on button Click
  Case Is = UCase("wDelRmCtg")
    Call HlpList.PMCd("CPYRT")
  ' ***** Manali 3.5.0 - Delete Records on button Click
  End Select
End Sub
Private Sub atxt_ReCalculate(Index As Integer, ByVal IdName As String, ByVal OldVal As Variant, ByVal RowNum As Integer, ByVal SourceName As String, ByVal SourceColName As String)
  '*** wIzChr on wIzTc                            Normal
      'valid PDesc225('TC', wIzTc)
  
  Select Case UCase(IdName)
  Case Is = UCase("wIzChr")
' **** Zubin 211 **** '
    If adc.Mode = xNorm Then _
       adc("wIzChr") = moCn.GetFldVal("Select vPValue from vParam where " + _
                            "vPCoCd = '" + gs_CoCd + "' and vPTyp= 'DEFCHR' " + _
                            "and vPMCd= '" + adc("wIzTc") + "' and vPSCd= ''")
       '("Select PDesc225 from Param where PTyp='TC' " + _
                                      "and PMCd='" + ADC("wIzTc") + "'")
'''''    If ADC.Mode = xNorm Then
'''''      If IsValidUsr(gs_CoCd, ADC("wIzTc"), moCn.GetFldVal("Select vPValue from vParam where " + _
'''''                          "vPCoCd= '" + gs_CoCd + "' and vPTyp= 'DEFCHR' and vPMCd= '" + ADC("wIzTc") + "' and " + _
'''''                          "vPSCd= ''"), gs_UsrCd) = True Then
'''''        ADC("wIzChr") = moCn.GetFldVal("Select vPValue from vParam where " + _
'''''                             "vPCoCd= '" + gs_CoCd + "' and vPTyp= 'DEFCHR' " + _
'''''                             "and vPMCd= '" + ADC("wIzTc") + "' and vPSCd= ''")
'''''      End If
'''''    End If
' **** Zubin 211 **** '
  
  '*** (Jen 2.13 speed)
  Case Is = UCase("wPurRtFr")
    If GrdIZ.Row > 0 And adc("wPurRtPrc") > 0 Then
      If GrdIZ.Value(GrdIZ.Row, "IzRmZRt") > 0 Then
        adc("wPurRtFr") = GrdIZ.Value(GrdIZ.Row, "IzRmZRt") - (GrdIZ.Value(GrdIZ.Row, "IzRmZRt") * adc("wPurRtPrc") / 100)
      Else
        adc("wPurRtFr") = 0
      End If
    Else
      adc("wPurRtFr") = 0
    End If
  Case Is = UCase("wPurRtTo")
    If GrdIZ.Row > 0 And adc("wPurRtPrc") > 0 Then
      If GrdIZ.Value(GrdIZ.Row, "IzRmZRt") > 0 Then
        adc("wPurRtTo") = GrdIZ.Value(GrdIZ.Row, "IzRmZRt") + (GrdIZ.Value(GrdIZ.Row, "IzRmZRt") * adc("wPurRtPrc") / 100)
      Else
        adc("wPurRtTo") = 0
      End If
    Else
      adc("wPurRtTo") = 0
    End If
  '*** (Jen 2.13 speed)
  End Select
End Sub
Private Sub ATXT_FldChange(Index As Integer, ByVal IdName As String, Cancel As Boolean, ErrMsg As String, ByVal pv_NewValue As Variant)
  '*** wIzTc            valid PMCd('TC' and PValue= TcTyp obtained from the MenuCode)
  '*** wIzYy            valid PMCd('YY')
  '*** wIzChr           valid PSCd('CHR', wIzTc) or Blank
  '*** wIzNo            valid InNo from InvHd(InTc= adc(wIzTc), InYy= adc(wIzYy),
  '                     InChr= adc(wIzChr)) or Zero
  
  Select Case UCase(IdName)
  Case Is = UCase("wIzTc")
    Cancel = Not moCn.RecSeek("select PMCd from Param where PTyp= 'TC' and " + _
                         "PMCd = '" + pv_NewValue + "' and " + _
                         "PValue= '" + mTcTyp + "' ")
    If Cancel = True Then ErrMsg = "Invalid Invoice Voucher Tc": Exit Sub
  Case Is = UCase("wIzYy")
    Cancel = Not moCn.RecSeek("select PMCd from Param where PTyp= 'YY' and " + _
                         "PMCd = '" + pv_NewValue + "'")
    If Cancel = True Then ErrMsg = "Invalid Invoice Voucher Year": Exit Sub
  Case Is = UCase("wIzChr")
    Cancel = pv_NewValue <> "" And (Not moCn.RecSeek("select vPSCd from vParam where " + _
                         "vPCoCd= '" + gs_CoCd + "' and vPTyp= 'CHR' and " + _
                         "vPMCd = '" + adc("wIzTc") + "' and vPSCd= '" + pv_NewValue + "'"))
    If Cancel = True Then ErrMsg = "Invalid Invoice Voucher Character": Exit Sub
    ' **** Zubin 211 **** '
    '''''    ErrMsg = "Invalid Voucher Character": Exit Sub
    Dim ws_UsrArr() As String, wi_Index As Integer
    ws_UsrArr = Split(moCn.GetFldVal("Select vPValue4 from vParam where vPCoCd= '" + gs_CoCd + "' and " + _
                                        "vPTyp='CHR' and vPMCd='" + adc("wIzTc") + "' and vPSCd= '" + pv_NewValue + "'"), ",")
    If IsArray(ws_UsrArr) Then
      For wi_Index = 0 To UBound(ws_UsrArr)
        Cancel = IIF(UCase(gs_UsrCd) = UCase(ws_UsrArr(wi_Index)), False, True)
        If Cancel = False Then: Exit For
      Next wi_Index
    End If
    If Cancel = True Then ErrMsg = "Access To This Invoice Voucher Character Is Denied To User " + gs_UsrCd: Exit Sub
    ' **** Zubin 211 **** '
  Case Is = UCase("wIzNo")
    '****** Sachin 3.02 31-12-07 - Partition Key added
    Cancel = (pv_NewValue <> 0) And (Not moCn.RecSeek("Select InNo From InvHd Where " + _
                                    "InCoCd= '" + gs_CoCd + "' and " + _
                                    "InTc= '" + adc("wIzTc") + "' And " + _
                                    "InYy= '" + adc("wIzYy") + "' and " + _
                                    "InChr= '" + adc("wIzChr") + "' and " + _
                                    "InNo= " + CStr(pv_NewValue) + " and InPrtKey=(Case When '" + gs_Partition + "'='" + ctCurrPrtn + "' Then '" + ctCurrPrtn + "' Else InPrtKey End ) "))
    If Cancel = True Then ErrMsg = "Invalid Voucher Invoice Number": Exit Sub
  
  '*** (Bef 3.1.1)
  ''*** (Jen 3.01)
  'Case Is = UCase("wAutoCpyOpt")
  '  Cancel = (pv_NewValue <> "") And (Not moCn.RecSeek("select PMCd from Param where PTyp= 'CPYRT' and " + _
  '                                                    "PMCd = '" + pv_NewValue + "' "))
  '  If Cancel = True Then ErrMsg = "Invalid List Of RM Categories": Exit Sub
  ''*** (Jen 3.01)
  '*** (Bef 3.1.1)
  
  '*** (Jen 3.1.1)
  Case Is = UCase("wAutoUseOnlyCustGoods")
    Cancel = Not moCn.RecSeek("select PMCd from Param where PTyp= 'YN' and PMCd = '" + pv_NewValue + "' ")
    If Cancel = True Then ErrMsg = "Invalid 'Use Customer Goods Only Option'": Exit Sub
  Case Is = UCase("wAutoHImpTypNotInSelD"), UCase("wAutoHImpTypNotInSelC"), UCase("wAutoHImpTypNotInSelG"), _
            UCase("wAutoHImpTypNotInSelP"), UCase("wAutoHImpTypNotInSelS"), UCase("wAutoHImpTypNotInSelX")
    Dim ws_ImpTypArr() As String, wi_i As Integer
    If pv_NewValue <> "" Then
      ws_ImpTypArr = Split(pv_NewValue, ",")
      If IsArray(ws_ImpTypArr) Then
        For wi_i = 0 To UBound(ws_ImpTypArr)
          Cancel = Not moCn.RecSeek("Select PMCd From Param where PTyp= 'IMPTYP' " + _
                   "and PMCd = '" + ws_ImpTypArr(wi_i) + "'")
          If Cancel = True Then ErrMsg = "Value(" + CStr(wi_i + 1) + ") Is An Invalid Import Type": Exit Sub
        Next wi_i
      End If
    End If
  '*** (Jen 3.1.1)
  
  '*** (Jen 3.2.0)
  Case Is = UCase("wAutoBestUseOnly")
    Cancel = Not moCn.RecSeek("select PMCd from Param where PTyp= 'YN' and PMCd = '" + pv_NewValue + "' ")
    If Cancel = True Then ErrMsg = "Invalid 'Match Best Use Only' Option": Exit Sub
  '*** (Jen 3.2.0)
  
  ' ****** Manalli 3.4.1 - 12/12/08
  Case Is = UCase("wSortBy")
    Cancel = Not moCn.RecSeek("Select PMCd from Param where PTyp='SRTPRHLP' and PMCd = '" + pv_NewValue + "' ")
    If Cancel = True Then ErrMsg = "Invalid Help Sort Option": Exit Sub
  ' ****** Manalli 3.4.1 - 12/12/08

  ' **** Manali 3.5.0 - Delete Records
  Case Is = UCase("wDelRmCtg")
    Cancel = pv_NewValue <> "" And Not moCn.RecSeek("Select 'x' from Param where PTyp='CPYRT' and PMCd='" + pv_NewValue + "' ")
    If Cancel = True Then ErrMsg = "Invalid RmCtg Combination ": Exit Sub
  End Select
End Sub

Private Sub CmdDelRm_Click()
  If wFlag = True Then Exit Sub
  
  If ms_InCls = "Y" Then DispMsg "Cannot Delete Records As Invoice Is Closed", etError: Exit Sub
  
  If GrdIZ.Row > 0 Then
    If GrdIZ.Value(GrdIZ.Row, "wIzRmBalWt") > 0 And GrdIZ.Value(GrdIZ.Row, "IzRmZWt") - GrdIZ.Value(GrdIZ.Row, "wIzRmBalWt") > 0 Then Exit Sub
  End If

  Call DispFra(DelRm)
End Sub

Private Sub CmdDelRmGo_Click()
  Call DeleteRm
End Sub

Private Sub CmdHlpSrt_Click()

  '****** Sachin 3.02 31-12-07 - Restrictions
  If adc("wIzDt") < CDate(ms_PurgeDt) = True Then DispMsg "Cannot Copy In Previous Partition Invoice", etError: Exit Sub

  'Zubin 211 (Patch 03)
 ' If wFlag = True Then Exit Sub
  
  '3.11.2
  'If GrdIZ.Row > 0 Then
  '  If GrdIZ.Value(GrdIZ.Row, "wIzRmBalWt") > 0 And GrdIZ.Value(GrdIZ.Row, "IzRmZWt") - GrdIZ.Value(GrdIZ.Row, "wIzRmBalWt") > 0 Then Exit Sub
  'End If
  'Zubin 211 (Patch 03)
  
  Call DispFra(HlpSrt)
End Sub

Private Sub GrdIZ_GotFocus()
  If mb_InCompInv Then DispMsg "Invoice contains Disrepancy", etError: Exit Sub

End Sub

Private Sub GrdIZ_When()
  
  'If CmdIZR.Visible = True Then CmdIZR.Enabled = True
  ' **** Zubin 211 **** '
  CmdIZR.Enabled = True
  ' **** Zubin 211 **** '
  
  CmdAuto.Enabled = True    '*** (Jen 3.01)
  
End Sub
Private Sub GrdIZ_RowWhen(ByVal RowNum As Integer)
  '*** If GrdIZ.IsNew Then
      'IzSr= Max(IzSr) of the grid GrdIZ + 1
  '*** Call procedure RmDetails to fill in records in the grid GrdIZD
      'This grid shows all the RmCodes and their corresponding Weights in the Invoice
      'for the current custom Code record of GrdIZ
  With GrdIZ
    If GrdIZ.IsNew(RowNum) Then
      If .Value(RowNum, "IzSr") = 0 Then .Value(RowNum, "IzSr") = .MaxVal("IzSr") + 1
    
    '*** (Jen 2.13 speed)
    Else
      If .Value(RowNum, "IzRmZRt") > 0 And adc("wPurRtPrc") > 0 Then
        Dim wd_Incr As Double
        wd_Incr = .Value(RowNum, "IzRmZRt") * adc("wPurRtPrc") / 100
        adc("wPurRtFr") = .Value(RowNum, "IzRmZRt") - wd_Incr
        adc("wPurRtTo") = .Value(RowNum, "IzRmZRt") + wd_Incr
      Else
        adc("wPurRtFr") = 0: adc("wPurRtTo") = 0
      End If
    '*** (Jen 2.13 speed)
   
    End If
    Call RmDetails("GrdIZD")
  End With
End Sub

Private Sub GrdIZ_ColWhen(ByVal RowNum As Integer, ByVal ColName As String, Cancel As Boolean, ErrMsg As String, ByVal Action As MwfCtl.en_fgAction)
  
  '*** Cannot edit any field in the grid GrdIZ
  '****** Sachin 3.02 - Locking Previous Partition Invoices
  If adc("WIZPRTKEY") <> ctCurrPrtn Then Cancel = True: ErrMsg = "Cannot Edit Previous Partition Invoices": Exit Sub
  
  Cancel = True: ErrMsg = "Cannot Edit This Value": Exit Sub
End Sub
Private Sub GrdIZ_ReCalc(ByVal RowNum As Integer, ByVal ColName As String, ByVal OldVal As Variant, ByVal SourceName As String, ByVal SourceColName As String)
  '*** wIzPcPerCt on IzRmZWt,IzRmQty                          Normal,Copy,Init
      'wIzPcPerCt= Round(IzRmZWt * IzRmQty)
  '*** wIzRmZAmt on IzRmZWt,IzRmZRt                           Normal,Copy,Init
      'wIzRmZAmt= IzRmZWt * IzRmZRt
  '*** wIzRmBalWt on IzRmZWt,TzRmWt                           GrdIZ= Normal,Init and GrdTxndZ= Normal,CopyOver,Delete
      'wIzRmBalWt= IzRmZWt - sum(TzRmWt) for the GrdIZ Row
  '*** wIzPurAmt on IzSr,TzPurAmt                             GrdIZ= Normal,Init and GrdTxndZ= Normal,CopyOver,Delete
      'wIzPurAmt= sum(TzPurAmt) for the Row
  
  Dim wSourceGrdMode As en_fgMode, i As Integer, wIzRow As Integer
  With GrdIZ
    Select Case UCase(ColName)
    Case Is = UCase("wIzPcPerCt")         '*** recalc on IzRmZWt, IzRmQty
      If .Mode = fgmnorm Or .Mode = fgmCopy Or .Mode = fgminit Then
        .Value(RowNum, "wIzPcPerCt") = Round(MWLib.Div(.Value(RowNum, "IzRmQty"), .Value(RowNum, "IzRmZWt")), 0)
      End If
    Case Is = UCase("wIzRmZAmt")          '*** recalc on IzRmZWt, IzRmZRt
      If .Mode = fgmnorm Or .Mode = fgmCopy Or .Mode = fgminit Then
        .Value(RowNum, "wIzRmZAmt") = .Value(RowNum, "IzRmZWt") * .Value(RowNum, "IzRmZRt")
      End If
    Case Is = UCase("wIzRmBalWt")         '*** recalc on IzRmZWt, TzRmWt
      Dim wTzRmWt As Double
      wIzRow = IIF(UCase(SourceName) = UCase("GrdIZ"), RowNum, .Row)
      wSourceGrdMode = Me.Controls(SourceName).Mode
      
      If .Mode = fgmnorm Then
      'Or .Mode = fgmCopy
      
        If UCase(SourceName) = UCase("GrdIZ") Or (UCase(SourceName) = UCase("GrdTxndZ") And _
           (wSourceGrdMode = fgmnorm Or wSourceGrdMode = fgmCopyOver Or wSourceGrdMode = fgmdel)) Then
          For i = 1 To GrdTxndZ.Rows - 1
            If Not GrdTxndZ.IsDel(i) Then wTzRmWt = wTzRmWt + GrdTxndZ.Value(i, "TzRmWt")
          Next i
          .Value(wIzRow, "wIzRmBalWt") = .Value(wIzRow, "IzRmZWt") - wTzRmWt
          'If UCase(SourceName) = UCase("GrdTxndZ") Then
          '  Controls(SourceName).AddSave "GrdIZ", asCurrent
          'End If
        End If
      End If
      If .Mode = fgminit Then
        If UCase(SourceName) = UCase("GrdIZ") Or (UCase(SourceName) = UCase("GrdTxndZ") And _
           (wSourceGrdMode = fgmnorm Or wSourceGrdMode = fgmCopyOver Or wSourceGrdMode = fgmdel)) Then
          '****** Sachin 3.02 31-12-07 - Partition Key added
          wTzRmWt = moCn.GetFldVal("Select Sum(TzRmWt) from TxndZ where " + _
                    "TzCoCd= '" + gs_CoCd + "' and " + _
                    "TzTc= '" + adc("wIzTc") + "' and " + _
                    "TzYy= '" + adc("wIzYy") + "' and " + _
                    "TzChr= '" + adc("wIzChr") + "' and " + _
                    "TzNo= " + CStr(adc("wIzNo")) + " and " + _
                    "TzSr= " + CStr(.Value(wIzRow, "IzSr")) + " And TzPrtKey=(Case When '" + adc("WIZPRTKEY") + "'='" + ctCurrPrtn + "' Then '" + ctCurrPrtn + "' Else TzPrtKey End) ")
          .Value(wIzRow, "wIzRmBalWt") = .Value(wIzRow, "IzRmZWt") - wTzRmWt
        End If
      End If
    Case Is = UCase("wIzPurAmt")         '*** recalc on IzSr, TzPurAmt
      Dim wTzPurAmt As Single
      wSourceGrdMode = Me.Controls(SourceName).Mode
      wIzRow = IIF(UCase(SourceName) = UCase("GrdIZ"), RowNum, .Row)
      If .Mode = fgmnorm Then
        If UCase(SourceName) = UCase("GrdIZ") Or (UCase(SourceName) = UCase("GrdTxndZ") And _
           (wSourceGrdMode = fgmnorm Or wSourceGrdMode = fgmCopyOver Or wSourceGrdMode = fgmdel)) Then
          For i = 1 To GrdTxndZ.Rows - 1
            If Not GrdTxndZ.IsDel(i) Then wTzPurAmt = wTzPurAmt + GrdTxndZ.Value(i, "TzPurAmt")
          Next i
          .Value(wIzRow, "wIzPurAmt") = wTzPurAmt
        End If
      End If
      If .Mode = fgminit Then
        If UCase(SourceName) = UCase("GrdIZ") Or (UCase(SourceName) = UCase("GrdTxndZ") And _
           (wSourceGrdMode = fgmnorm Or wSourceGrdMode = fgmCopyOver Or wSourceGrdMode = fgmdel)) Then
          '****** Sachin 3.02 31-12-07 - Partition Key added
          wTzPurAmt = moCn.GetFldVal("Select Sum(TzPurAmt) from TxndZ where " + _
                      "TzCoCd= '" + gs_CoCd + "' and " + _
                      "TzTc= '" + adc("wIzTc") + "' and " + _
                      "TzYy= '" + adc("wIzYy") + "' and " + _
                      "TzChr= '" + adc("wIzChr") + "' and " + _
                      "TzNo= " + CStr(adc("wIzNo")) + " and " + _
                      "TzSr= " + CStr(.Value(wIzRow, "IzSr")) + " And TzPrtKey=(Case When '" + adc("WIZPRTKEY") + "'='" + ctCurrPrtn + "' Then '" + ctCurrPrtn + "' Else TzPrtKey End) ")
          .Value(wIzRow, "wIzPurAmt") = wTzPurAmt
        End If
      End If
    End Select
  End With
End Sub
Private Sub GrdIZ_ColChange(ByVal RowNum As Integer, ByVal ColName As String, ByVal NewValue As Variant, Cancel As Boolean, ErrMsg As String)
  With GrdIZ
    Select Case ColName
    End Select
  End With
End Sub
Private Sub GrdIZ_InitKey(ByVal RowNum As Integer)
  '*** Initialise Key Fields
  With GrdIZ
    .Store "IzCoCd", gs_CoCd
    .Store "IzTc", adc("wIzTc")
    .Store "IzYy", adc("wIzYy")
    .Store "IzChr", adc("wIzChr")
    .Store "IzNo", adc("wIzNo")
    
    '****** Sachin 3.02 31-12-07 - Partition Key added
    If .IsNew(RowNum) Then _
    .Store "IzInIdNo", moCn.GetFldVal("Select InIdNo From InvHd " + _
                                      " Where InCoCd='" + gs_CoCd + "' And InTc='" + adc("wIzTc") + "' and InYy='" + adc("wIzYy") + "'" + _
                                      " and InChr='" + adc("wIzChr") + "' and InNo=" + CStr(adc("wIzNo")) + " and InPrtKey='" + ctCurrPrtn + "' ")
    '****** Sachin 3.02.0 - Id Fields Change
    
  End With
End Sub
Private Sub GrdIZ_Valid()
  'CmdIZR.Enabled = False
End Sub
Private Sub GrdTxndZ_GotFocus()
  If mb_InCompInv Then DispMsg "Invoice contains Disrepancy", etError: Exit Sub
  '*** If there are no rows in GrdIZ then the user cannot add a row in GrdTxndZ
  If GrdIZ.Row = 0 Then
    GrdTxndZ.AllowAdd = False
    GrdTxndZ.AllowDelete = False
   '******Geeta******Emr207*And ms_InCls = "N"
   '****** Sachin 3.02 31-12-07 - purge date condition added below
  ElseIf GrdIZ.Row > 0 And ms_InCls = "N" Then
    GrdTxndZ.AllowAdd = True
    GrdTxndZ.AllowDelete = True
  End If
  
  
End Sub
Private Sub GrdTxndZ_When()
  '*** Disable the command buttons CmdCp & CmdIZR
  CmdCP.Enabled = False
  CmdIZR.Enabled = False
  CmdAuto.Enabled = False    '*** (Jen 3.01)
  If mb_InCompInv Then GrdTxndZ.AllowAdd = False

End Sub
Private Sub GrdTxndZ_RowWhen(ByVal RowNum As Integer)
  '*** If GrdTxndZ.IsNew Then TzSrNo= Max + 1
      'TzKey= autogenerated value (prekey= adc(wIzYy))
  With GrdTxndZ
    If .IsNew(RowNum) Then
      If .Value(RowNum, "TzSrNo") = 0 Then .Value(RowNum, "TzSrNo") = .MaxVal("TzSrNo") + 1
      .Value(RowNum, "TzPrtKey") = ctCurrPrtn   '****** Sachin 3.02 31-12-07
      .Value(RowNum, "TzKey") = moCn.AutoGen("TxndZ", "TzKey", gs_CoCd + adc("wIzYy"))
    End If
  End With
End Sub
Private Sub GrdTxndZ_ColWhen(ByVal RowNum As Integer, ByVal ColName As String, Cancel As Boolean, ErrMsg As String, ByVal Action As MwfCtl.en_fgAction)
  '*** Set Helps For Each Field
  '*** Cannot edit any field when Invoice is closed
  '*** Cannot edit fields wTzBaseBoeNo, TzPurRt, TzPurAmt
  
  '****** Sachin 3.02 - Locking Previous Partition Invoices
  If adc("WIZPRTKEY") <> ctCurrPrtn Then Cancel = True: ErrMsg = "Cannot Edit Previous Partition Invoices": Exit Sub
  
  If mf_DtLocked = True Then Cancel = True: ErrMsg = "Cannot Edit Locked Entry": Exit Sub
  Dim wRmCtg As String, wRmSCtg As String, wGrpBySz As String, wGrpByRt As String
  
  '****** Sachin 3.02 31-12-07 - Purge Date condition added below
  If ms_InCls = "Y" Or adc("wIzDt") < CDate(ms_PurgeDt) Then Cancel = True: Exit Sub
  With GrdTxndZ
    Select Case ColName
    Case Is = UCase("wTzBaseVchNo")
     '*******************Geeta***************Emr207*********15/07/04
     ' Call HlpList.BaseCustmPurNo(gs_CoCd, GrdIZ.Value(GrdIZ.Row, "IzRmZCd"), _
           moCn.GetFldVal("Select PDesc225 from Param where PTyp= 'SRTPRHLP' and " + _
                          "PMCd= (Select HSrtBasePurVch From Head where HCoCd= '" + gs_CoCd + "' and HCd= '" + ctSelfCmCd + "') and PSCd= ''"))
      Call HlpList.BaseCustmPurNo(gs_CoCd, GrdIZ.Value(GrdIZ.Row, "IzRmZCd"), moCn.GetFldVal("Select PDesc225 from Param where PTyp= 'SRTPRHLP' and " + _
                          "PMCd= '" + adc("wSortBy") + "'"), md_CnvFct, ms_MulDiv, adc("wPurRtFr"), adc("wPurRtTo"))
    Case Is = UCase("wTzBaseBoeNo")
      Cancel = True: ErrMsg = "Cannot Edit This Value": Exit Sub
    Case Is = UCase("TzPurRt")
      Cancel = True: ErrMsg = "Cannot Edit This Value": Exit Sub
    Case Is = UCase("TzPurAmt")
      Cancel = True: ErrMsg = "Cannot Edit This Value": Exit Sub
    End Select
  End With
End Sub
Private Sub GrdTxndZ_ReCalc(ByVal RowNum As Integer, ByVal ColName As String, ByVal OldVal As Variant, ByVal SourceName As String, ByVal SourceColName As String)
  '*** wTzBaseVchNo on TzRefYy,TzRefKey                     init
      'wTzBaseVchNo = TzTc+TzYy+TzChr+TzNo+TzSr from TxndZ(TzBM= 'B', TzYy= TzRefYy, TzKey= TzRefKey)
  '*** wTzBaseBoeNo on wTzBaseVchNo                         Normal,init
      'wTzBaseBoeNo = TBOENo from the Txn record where Voucher Number equal to wFdBaseVchNo
  '*** TzRmWt on wTzBaseVchNo                               Normal
      'TzRmWt = TzRmWt- TzRmDespWt from the TxndZ record where Voucher Number equal to wFdBaseVchNo
  '*** TzPurRt on wTzBaseVchNo                              Normal
      'TzPurRt = TzPurRt from the TxndZ record where Voucher Number equal to wFdBaseVchNo
  '*** TzRefYy on wTzBaseVchNo                              Normal
      'TzRefYy = TzYy from the TxndZ record where Voucher Number equal to wFdBaseVchNo
  '*** TzRefKey on wTzBaseVchNo                             Normal
      'TzRefKey = TzKey from the TxndZ record where Voucher Number equal to wFdBaseVchNo
  '*** TzPurAmt on TzRmWt,TzPurRt                           Normal
      'TzPurAmt = TzRmWt * TzPurRt
  
  Dim wBaseVchSr() As String
  Dim ws_RmCtg As String
  
  With GrdTxndZ
    Select Case ColName
    Case Is = UCase("wTzBaseVchNo")         '*** Recalc on TzRefYy, TzRefKey
      If .Mode = fgminit Then
        '****** Sachin 3.02 31-12-07 - Partition Key added
        .Value(RowNum, "wTzBaseVchNo") = _
               moCn.GetFldVal("Select TzTc+'/'+TzYy+'/'+TzChr+'/'+LTrim(Str(TzNo))+'/'+" + _
                              "LTrim(Str(TzSr)) from TxndZ where TzBM= 'B' and " + _
                              "TzCoCd= '" + gs_CoCd + "' and " + _
                              "TzYy='" + .Value(RowNum, "TzRefYy") + "' and " + _
                              "TzKey= " + CStr(.Value(RowNum, "TzRefKey")) + " and TzPrtKey=(Case When '" + adc("WIZPRTKEY") + "'='" + ctCurrPrtn + "' Then '" + ctCurrPrtn + "' Else TzPrtKey End) ")
      End If
    Case Is = UCase("wTzBaseBoeNo")         '*** recalc on wTzBaseVchNo
      If .Mode = fgmnorm Or .Mode = fgminit Then
        wBaseVchSr() = Split(.Value(RowNum, "wTzBaseVchNo"), "/")
        '****** Sachin 3.02 31-12-07 - Partition Key added
        .Value(RowNum, "wTzBaseBoeNo") = moCn.GetFldVal("Select TBOENo from Txn where " + _
               "TCoCd= '" + gs_CoCd + "' and " + _
               "TTc= '" + wBaseVchSr(0) + "' and " + _
               "TYy= '" + wBaseVchSr(1) + "' and " + _
               "TChr= '" + wBaseVchSr(2) + "' and " + _
               "TNo= " + wBaseVchSr(3) + " and TPrtKey=(Case When '" + adc("WIZPRTKEY") + "'='" + ctCurrPrtn + "' Then '" + ctCurrPrtn + "' Else TPrtKey End) ")
      End If
    Case Is = UCase("TzRmWt")               '*** recalc on wTzBaseVchNo
      Dim wReqWt As Single
      If .Mode = fgmnorm Then
        wBaseVchSr() = Split(.Value(RowNum, "wTzBaseVchNo"), "/")
        '****** Sachin 3.02 31-12-07 - Partition Key added
        wReqWt = moCn.GetFldVal("Select TzRmWt- TzRmDespWt from TxndZ where " + _
                 "TzCoCd= '" + gs_CoCd + "' and " + _
                 "TzTc= '" + wBaseVchSr(0) + "' and " + _
                 "TzYy= '" + wBaseVchSr(1) + "' and " + _
                 "TzChr= '" + wBaseVchSr(2) + "' and " + _
                 "TzNo= " + wBaseVchSr(3) + " and " + _
                 "TzSr= " + wBaseVchSr(4) + " and TzPrtKey=(Case When '" + adc("WIZPRTKEY") + "'='" + ctCurrPrtn + "' Then '" + ctCurrPrtn + "' Else TzPrtKey End) and TzBM= 'B' ")
        .Value(RowNum, "TzRmWt") = IIF(wReqWt > GrdIZ.Value(GrdIZ.Row, "wIzRmBalWt"), GrdIZ.Value(GrdIZ.Row, "wIzRmBalWt"), wReqWt)
      End If
    Case Is = UCase("TzPurRt")              '*** recalc on wTzBaseVchNo
      If .Mode = fgmnorm Then
        ' ***** Manali 3.6.0 - 31/12/09 - Notional Instead od Purchase Rate For DTA
        ws_RmCtg = moCn.GetFldVal("Select RmCtg from RmMst Where RmCd='" + GrdIZ.Value(GrdIZ.Row, "IzRmZCd") + "'  " + _
                                  IIF(adc("wIzPrtKey") = ctCurrPrtn, " and RmPrtKey= '" + ctCurrPrtn + "' ", ""))
        
        'Sachin - 4.1.0.0
        If ms_NotionRt = "Y" And (ws_RmCtg = "G" Or ws_RmCtg = "P" Or ws_RmCtg = "S" Or ws_RmCtg = "L") Then
          Dim wd_CustmCdPurityZ As Double
          
        
          wd_CustmCdPurityZ = moCn.GetFldVal("Select RmPurityZ * " + _
                                IIF(ws_RmCtg = "G" And md_GldFiness > 0, CStr(md_GldFiness), IIF(ws_RmCtg = "P" And md_PlFiness > 0, CStr(md_PlFiness), _
                                IIF(ws_RmCtg = "S" And md_SlFiness > 0, CStr(md_SlFiness), IIF(ws_RmCtg = "L" And md_PdFiness > 0, CStr(md_PdFiness), 1)))) + _
                              " from RmMst Where RmCd='" + GrdIZ.Value(GrdIZ.Row, "IzRmZCd") + "'  " + _
                              IIF(adc("wIzPrtKey") = ctCurrPrtn, " and RmPrtKey= '" + ctCurrPrtn + "' ", ""))
        
          .Value(RowNum, "TzPurRt") = MWLib.Div(IIF(ws_RmCtg = "G", md_LmgCst, IIF(ws_RmCtg = "P", md_LmpCst, IIF(ws_RmCtg = "S", md_LmsCst, IIF(ws_RmCtg = "L", md_LmlCst, 1)))), _
                                      moCn.GetFldVal("Select PNum from Param Where PTyp='GMFACTOR' and PMCd='" + ws_RmCtg + "' and PSCd='" + adc("wIzCmCurCd") + "'")) _
                                      * IIF(wd_CustmCdPurityZ > 0, wd_CustmCdPurityZ, 1)
        ' ***** Manali 3.6.0 - 31/12/09 - Notional Instead od Purchase Rate For DTA
        Else
          wBaseVchSr() = Split(.Value(RowNum, "wTzBaseVchNo"), "/")
          ' .Value(RowNum, "TzPurRt") = moCn.GetFldVal("Select TzPurRt from TxndZ where " + _
                   "TzCoCd= '" + gs_CoCd + "' and " + _
                   "TzTc= '" + wBaseVchSr(0) + "' and " + _
                   "TzYy= '" + wBaseVchSr(1) + "' and " + _
                   "TzChr= '" + wBaseVchSr(2) + "' and " + _
                   "TzNo= " + wBaseVchSr(3) + " and " + _
                   "TzSr= " + wBaseVchSr(4) + " and TzBM= 'B' ")
          '**********Geeta***********Emr207************
                     
          '*** (Bef 2.13)
          '.Value(RowNum, "TzPurRt") = moCn.GetFldVal("Select TzPurRt " + _
                  "from TxndZ Join Txn On TzCoCd= TCoCd and TzTc= TTc and TzYy= TYy " + _
                  "and TzChr= TChr and TzNo= TNo where " + _
                  "TzCoCd= '" + gs_CoCd + "' and " + _
                  "TzTc= '" + wBaseVchSr(0) + "' and " + _
                  "TzYy= '" + wBaseVchSr(1) + "' and " + _
                  "TzChr= '" + wBaseVchSr(2) + "' and " + _
                  "TzNo= " + wBaseVchSr(3) + " and " + _
                  "TzSr= " + wBaseVchSr(4) + " and TzBM= 'B' ")
          '*** (Bef 2.13)
            
          '*** (Bef 3.1.1) (Bef 05/11/2007)
          ''*** (Jen 2.13)
          ''****** Sachin 3.02 31-12-07 - Partition Key added
          '.Value(RowNum, "TzPurRt") = moCn.GetFldVal("Select TzPurRt " + _
          '      "* (case when TMulDiv= 'D' then TCnvRt else 1 end)/ (case when TMulDiv= 'M' then TCnvRt else 1 end) " + _
          '      "* " + IIF(ms_MulDiv = "M", CStr(md_CnvFct), CStr(1)) + _
          '      "/ " + IIF(ms_MulDiv = "D", CStr(md_CnvFct), CStr(1)) + _
          '      " from TxndZ Join Txn On /* TzTIdNo=TIdNo and */ TzCoCd= TCoCd and TzTc= TTc and TzYy= TYy " + _
          '      "and TzChr= TChr and TzNo= TNo where " + _
          '      "TzCoCd= '" + gs_CoCd + "' and " + _
          '      "TzTc= '" + wBaseVchSr(0) + "' and " + _
          '      "TzYy= '" + wBaseVchSr(1) + "' and " + _
          '      "TzChr= '" + wBaseVchSr(2) + "' and " + _
          '      "TzNo= " + wBaseVchSr(3) + " and " + _
          '      "TzSr= " + wBaseVchSr(4) + " and TzPrtKey=(Case When '" + ADC("WIZPRTKEY") + "'='" + ctCurrPrtn + "' Then '" + ctCurrPrtn + "' Else TzPrtKey End) and TzBM= 'B' ")
          ''*** (Jen 2.13)
          '*** (Bef 3.1.1) (Bef 05/11/2007)
          
          '****** Sachin 3.02 31-12-07 - Current Partition
          '*** (Jen 3.1.1) (05/11/2007)
          .Value(RowNum, "TzPurRt") = moCn.GetFldVal("Select convert (decimal (14, 3), " + _
                "convert (decimal (14, 3), TzPurRt) " + _
                "* (case when TMulDiv= 'D' then convert (decimal (14, 4), TCnvRt) else 1.0000 end)" + _
                "/ (case when TMulDiv= 'M' then convert (decimal (14, 4), TCnvRt) else 1.0000 end) " + _
                "* " + Format(IIF(ms_MulDiv = "M", CStr(md_CnvFct), "1"), "0.0000") + _
                "/ " + Format(IIF(ms_MulDiv = "D", CStr(md_CnvFct), "1"), "0.0000") + ") " + _
                " from TxndZ Join Txn On " + IIF(gs_Partition = ctCurrPrtn, " TzPrtKey=TPrtKey and ", "") + " TzTIdNo=TIdNo and TzCoCd= TCoCd and TzTc= TTc and TzYy= TYy " + _
                "and TzChr= TChr and TzNo= TNo where " + _
                "TzCoCd= '" + gs_CoCd + "' and " + _
                "TzTc= '" + wBaseVchSr(0) + "' and " + _
                "TzYy= '" + wBaseVchSr(1) + "' and " + _
                "TzChr= '" + wBaseVchSr(2) + "' and " + _
                "TzNo= " + wBaseVchSr(3) + " and " + _
                "TzSr= " + wBaseVchSr(4) + " and TzPrtKey=(Case When '" + adc("WIZPRTKEY") + "'='" + ctCurrPrtn + "' Then '" + ctCurrPrtn + "' Else TzPrtKey End) and TzBM= 'B' ")
          '*** (Jen 3.1.1) (05/11/2007)
        End If
      End If
    '************geeta************Emr207********
    Case Is = UCase("wPurRtInRs")              '*** recalc on wTzBaseVchNo
      If adc.Mode = xNorm And (.Mode = fgminit Or .Mode = fgmnorm) Then
        ws_RmCtg = moCn.GetFldVal("Select RmCtg from RmMst Where RmCd='" + GrdIZ.Value(GrdIZ.Row, "IzRmZCd") + "'  " + _
                                  IIF(adc("wIzPrtKey") = ctCurrPrtn, " and RmPrtKey= '" + ctCurrPrtn + "' ", ""))
        
        ' ***** Manali 3.6.0 - 31/12/09 - Notional Instead od Purchase Rate For DTA
        'Sachin - 4.1.0.0
        If ms_NotionRt = "Y" And (ws_RmCtg = "G" Or ws_RmCtg = "P" Or ws_RmCtg = "S" Or ws_RmCtg = "L") Then
          .Value(RowNum, "wPurRtInRs") = .Value(RowNum, "TzPurRt") * _
                                        MWLib.Div(Format(IIF(ms_RsMulDiv = "M", CStr(md_RsCnvFct), CStr(1)), "0.0000"), _
                                                  Format(IIF(ms_RsMulDiv = "D", CStr(md_RsCnvFct), CStr(1)), "0.0000"))
        ' ***** Manali 3.6.0 - 31/12/09 - Notional Instead od Purchase Rate For DTA
        Else
          wBaseVchSr() = Split(.Value(RowNum, "wTzBaseVchNo"), "/")
          '     .Value(RowNum, "wPurRtInRs") = .Value(RowNum, "TzPurRt") _
                   * moCn.GetFldVal("Select TCnvRt  from Txn where " + _
                    "TCoCd= '" + gs_CoCd + "' and " + _
                    "TTc= '" + wBaseVchSr(0) + "' and " + _
                    "TYy= '" + wBaseVchSr(1) + "' and " + _
                    "TChr= '" + wBaseVchSr(2) + "' and " + _
                    "TNo= " + wBaseVchSr(3))
                    
          '*** (Bef 2.13)
          '.Value(RowNum, "wPurRtInRs") = moCn.GetFldVal("Select TzPurRt " + _
                 "from TxndZ Join Txn On TzCoCd=TCoCd and TzTc=TTc and TzYy=TYy " + _
                 "and TzChr=TChr and TzNo=TNo where " + _
                 "TzCoCd= '" + gs_CoCd + "' and " + _
                 "TzTc= '" + wBaseVchSr(0) + "' and " + _
                 "TzYy= '" + wBaseVchSr(1) + "' and " + _
                 "TzChr= '" + wBaseVchSr(2) + "' and " + _
                 "TzNo= " + wBaseVchSr(3) + " and " + _
                 "TzSr= " + wBaseVchSr(4) + " and TzBM= 'B' ")
          '*** (Bef 2.13)
                    
          '*** (Jen 2.13)
          '*** (Bef 23/08/06)
          '.Value(RowNum, "wPurRtInRs") = moCn.GetFldVal("Select TzPurRt " + _
          '       "* (case when TMulDiv= 'D' then TCnvRt else 1 end)/ (case when TMulDiv= 'M' then TCnvRt else 1 end) " + _
          '       "* " + IIF(ms_RsMulDiv = "M", CStr(md_RsCnvFct), CStr(1)) + _
          '       "/ " + IIF(ms_RsMulDiv = "D", CStr(md_RsCnvFct), CStr(1)) + _
          '       " from TxndZ Join Txn On TzCoCd= TCoCd and TzTc= TTc and TzYy= TYy " + _
          '       "and TzChr=TChr and TzNo=TNo where " + _
          '       "TzCoCd= '" + gs_CoCd + "' and " + _
          '       "TzTc= '" + wBaseVchSr(0) + "' and " + _
          '       "TzYy= '" + wBaseVchSr(1) + "' and " + _
          '       "TzChr= '" + wBaseVchSr(2) + "' and " + _
          '       "TzNo= " + wBaseVchSr(3) + " and " + _
          '       "TzSr= " + wBaseVchSr(4) + " and TzBM= 'B' ")
          '*** (Bef 23/08/06)
          '****** Sachin 3.02 31-12-07 - Partition Key added
          .Value(RowNum, "wPurRtInRs") = moCn.GetFldVal("Select TzPurRt " + _
                 "* (case when TMulDiv= 'D' then TCnvRt else 1 end)/ (case when TMulDiv= 'M' then TCnvRt else 1 end) " + _
                 "* " + IIF(ms_MulDiv = "M", CStr(md_CnvFct), CStr(1)) + _
                 "/ " + IIF(ms_MulDiv = "D", CStr(md_CnvFct), CStr(1)) + _
                 "* " + IIF(ms_RsMulDiv = "M", CStr(md_RsCnvFct), CStr(1)) + _
                 "/ " + IIF(ms_RsMulDiv = "D", CStr(md_RsCnvFct), CStr(1)) + _
                 " from TxndZ Join Txn On " + IIF(gs_Partition = ctCurrPrtn, " TzPrtKey=TPrtKey and ", "") + " TzTIdNo=TIdNo and TzCoCd= TCoCd and TzTc= TTc and TzYy= TYy " + _
                 "and TzChr=TChr and TzNo=TNo where " + _
                 "TzCoCd= '" + gs_CoCd + "' and " + _
                 "TzTc= '" + wBaseVchSr(0) + "' and " + _
                 "TzYy= '" + wBaseVchSr(1) + "' and " + _
                 "TzChr= '" + wBaseVchSr(2) + "' and " + _
                 "TzNo= " + wBaseVchSr(3) + " and " + _
                 "TzSr= " + wBaseVchSr(4) + " and TzPrtKey=(Case When '" + adc("WIZPRTKEY") + "'='" + ctCurrPrtn + "' Then '" + ctCurrPrtn + "' Else TzPrtKey End) and TzBM= 'B' ")
          '*** (Jen 2.13)
        End If
      End If
    '************geeta************Emr207********
    Case Is = UCase("TzRefYy")              '*** recalc on wTzBaseVchNo
      If .Mode = fgmnorm Then
        wBaseVchSr() = Split(.Value(RowNum, "wTzBaseVchNo"), "/")
        '****** Sachin 3.02 31-12-07 - Partition Key added
        .Value(RowNum, "TzRefYy") = moCn.GetFldVal("Select TzYy from TxndZ where " + _
                 "TzCoCd= '" + gs_CoCd + "' and " + _
                 "TzTc= '" + wBaseVchSr(0) + "' and " + _
                 "TzYy= '" + wBaseVchSr(1) + "' and " + _
                 "TzChr= '" + wBaseVchSr(2) + "' and " + _
                 "TzNo= " + wBaseVchSr(3) + " and " + _
                 "TzSr= " + wBaseVchSr(4) + " and TzPrtKey=(Case When '" + adc("WIZPRTKEY") + "'='" + ctCurrPrtn + "' Then '" + ctCurrPrtn + "' Else TzPrtKey End) and TzBM= 'B' ")
      End If
    Case Is = UCase("TzRefKey")             '*** recalc on wTzBaseVchNo
      If .Mode = fgmnorm Then
        wBaseVchSr() = Split(.Value(RowNum, "wTzBaseVchNo"), "/")
        '****** Sachin 3.02 31-12-07 - Partition Key added
        .Value(RowNum, "TzRefKey") = moCn.GetFldVal("Select TzKey from TxndZ where " + _
                 "TzCoCd= '" + gs_CoCd + "' and " + _
                 "TzTc= '" + wBaseVchSr(0) + "' and " + _
                 "TzYy= '" + wBaseVchSr(1) + "' and " + _
                 "TzChr= '" + wBaseVchSr(2) + "' and " + _
                 "TzNo= " + wBaseVchSr(3) + " and " + _
                 "TzSr= " + wBaseVchSr(4) + " and TzPrtKey=(Case When '" + adc("WIZPRTKEY") + "'='" + ctCurrPrtn + "' Then '" + ctCurrPrtn + "' Else TzPrtKey End) and TzBM= 'B' ")
      End If
    Case Is = UCase("TzPurAmt")             '*** recalc on TzRmWt,TzPurRt
      Dim wPurAmt As Single
      If .Mode = fgmnorm Then
        ' ***** Manali 3.10.0 - 23/05/12 - Change done for Uni, when amount upto 3 decimals, get amount as 0.01 to skip amt>0 check
        wPurAmt = .Value(RowNum, "TzRmWt") * .Value(RowNum, "TzPurRt")
        wPurAmt = IIF(Round(wPurAmt, 3) > 0 And Round(wPurAmt, 2) = 0, 0.01, wPurAmt)
        .Value(RowNum, "TzPurAmt") = wPurAmt
        'bef 3.10.0
        '.Value(RowNum, "TzPurAmt") = .Value(RowNum, "TzRmWt") * .Value(RowNum, "TzPurRt")
      End If
    End Select
  End With
End Sub
Private Sub GrdTxndZ_ColChange(ByVal RowNum As Integer, ByVal ColName As String, ByVal NewValue As Variant, Cancel As Boolean, ErrMsg As String)
  '*** wTzBaseVchNo     valid Voucher Sr from TxndZ(TzBm= 'B')
  
  With GrdTxndZ
    Select Case ColName
    Case Is = UCase("wTzBaseVchNo")
      Dim wBaseVchSr() As String
      wBaseVchSr() = Split(NewValue, "/")
      If UBound(wBaseVchSr) <> 4 Then Cancel = True: ErrMsg = "Invalid Base Voucher Number": Exit Sub
      If Val(wBaseVchSr(3)) = 0 Or Len(Trim(wBaseVchSr(3))) <> Len(CStr(Val(wBaseVchSr(3)))) Then Cancel = True: ErrMsg = "Invalid Base Voucher Number": Exit Sub
      If Val(wBaseVchSr(4)) = 0 Or Len(Trim(wBaseVchSr(4))) <> Len(CStr(Val(wBaseVchSr(4)))) Then Cancel = True: ErrMsg = "Invalid Base Voucher Number": Exit Sub
      
      '****** Sachin 3.02 31-12-07 - Partition Key added
      ' ***** Manali 3.5.0 - 07/11/08 - added [TzCls = 'N']
      ' **** Manali 3.5.0 - 08/06/09 - GrdTxndZ.Mode checked for TzCls
      Cancel = Not moCn.RecSeek("Select TzTc from TxndZ where " + _
                 "TzCoCd= '" + gs_CoCd + "' and " + _
                 "TzTc= '" + wBaseVchSr(0) + "' and " + _
                 "TzYy= '" + wBaseVchSr(1) + "' and " + _
                 "TzChr= '" + wBaseVchSr(2) + "' and " + _
                 "TzNo= " + wBaseVchSr(3) + " and " + _
                 "TzSr= " + wBaseVchSr(4) + " and TzPrtKey=(Case When '" + adc("WIZPRTKEY") + "'='" + ctCurrPrtn + "' Then '" + ctCurrPrtn + "' Else TzPrtKey End) and TzBM= 'B' " + _
                 IIF(GrdTxndZ.Mode = fgmnorm Or GrdTxndZ.Mode = fgmCopy, " and TzCls='N' ", ""))
      
      ' ***** Manali 3.5.0 - 07/11/08 - added [TzCls = 'N']
      If Cancel = True Then ErrMsg = "Invalid Base Voucher Number Or Voucher is Closed": Exit Sub
      
    End Select
  End With
End Sub
Private Sub GrdTxndZ_InitKey(ByVal RowNum As Integer)
  '*** Initialise Key Fields
  Dim wBaseVchSr() As String, wTzKey As Single, wTzRefYy As String, wTzRefKey As Single
  With GrdTxndZ
    Dim wIzRow As Single
    wIzRow = IIF(GrdIZ.Row > 0, GrdIZ.Row, 1)
    If .IsNew(RowNum) Then
      .Store "TzCoCd", gs_CoCd
      .Store "TzTc", adc("wIzTc")
      .Store "TzYy", adc("wIzYy")
      .Store "TzChr", adc("wIzChr")
      .Store "TzNo", adc("wIzNo")
      .Store "TzSr", GrdIZ.Value(GrdIZ.Row, "IzSr")
      .Store "TzRmZCd", GrdIZ.Value(GrdIZ.Row, "IzRmZCd")
      .Store "TzRmCdDesc", ""
      .Store "TzRmQty", 0
      .Store "TzPcPerCt", 0
      .Store "TzBm", "M"
      .Store "TzRmDespWt", 0
      .Store "TzCls", ""
      .Store "TzClsDt", MWLib.EmptyDate   '*** Manali 350nxt
      'Geeta***Emr210
      .Store "TzRmDespPurgeWt", 0
      .Store "TzDesc", ""     '****** Sachin 2.12 - 09-12-2006 - Desc in Custom Purchase
      
      '*** (Jen 3.1.2)
      .Store "TzOdTc", ""
      .Store "TzOdYy", ""
      .Store "TzOdChr", ""
      .Store "TzOdNo", 0
      .Store "TzDmCd", ""
      .Store "TzRmCd", ""
      .Store "TzRmSzFr", 0
      .Store "TzRmSzTo", 0
      '*** (Jen 3.1.2)
      
      '****** Sachin 3.02 - Id fields Change
      .Store "TzTIdNo", 0
      .Store "TzInIdNo", moCn.GetFldVal("Select InIdNo From InvHd " + _
                                        " Where InCoCd='" + gs_CoCd + "' and InTc='" + adc("wIzTc") + "' and InYy='" + adc("wIzYy") + "'" + _
                                        " and InChr='" + adc("wIzChr") + "' and InNo=" + CStr(adc("wIzNo")) + " and InPrtKey='" + ctCurrPrtn + "' ")
                                       
      .Store "TzPrtDespQty", 0
      '****** Sachin 3.02 - Id fields Change
    End If
    
    '****** Sachin 3.02 - Values to New Fields
      .Store "TzRmCtg", moCn.GetFldVal("Select RmCtg From RmMst " + _
                                       " Where RmCd='" + GrdIZ.Value(GrdIZ.Row, "IzRmZCd") + "'" + IIF(adc("WIZPRTKEY") = ctCurrPrtn, " and RmPrtKey='" + ctCurrPrtn + "'", ""))
      .Store "TzRmSCtg", moCn.GetFldVal("Select RmSCtg From RmMst " + _
                                       " Where RmCd='" + GrdIZ.Value(GrdIZ.Row, "IzRmZCd") + "'" + IIF(adc("WIZPRTKEY") = ctCurrPrtn, " and RmPrtKey='" + ctCurrPrtn + "'", ""))
    '****** Sachin 3.02 - Values to New Fields
    
  End With
End Sub
Private Sub GrdTxndZ_LostFocus()
  '*** Do not allow the user to leave the grid GrdTxndZ if only partial matching has been
      'done for any of the records of InvZRm
  
  wFlag = False
  
    ' Commented in 3.11.2 as this check was not required
      With GrdIZ
        Dim wIzRow As Single, wTotRmWt As Single
        '****** Sachin 3.02 31-12-07 - Current Partition
        If GrdTxndZ.Rows - 1 > 0 And adc("wIzDt") > CDate(ms_PurgeDt) Then
          wIzRow = IIF(.Row > 0, .Row, 1)
          '****** Sachin 3.02 31-12-07 - Current Partition
          wTotRmWt = moCn.GetFldVal("Select Round(sum(TzRmWt), 3) from TxndZ where " + _
                                    "TzCoCd= '" + gs_CoCd + "' and " + _
                                    "TzTc='" + adc("wIzTc") + "' and " + _
                                    "TzYy='" + adc("wIzYy") + "' and " + _
                                    "TzChr='" + adc("wIzChr") + "' and " + _
                                    "TzNo=" + CStr(adc("wIzNo")) + " and " + _
                                    "TzSr= " + CStr(GrdIZ.Value(wIzRow, "IzSr")) + " and " + _
                                    "TzPrtKey='" + ctCurrPrtn + "' and " + _
                                    "TzBM= 'M'")
          If wTotRmWt <> GrdIZ.Value(wIzRow, "IzRmZWt") Then
            wFlag = True
'            DispMsg "Weight Matching Has Not Been Done Properly", etError
'            GrdIZ.Row = wIzRow
'            GrdTxndZ.SetFocus
            Exit Sub
          End If
        End If
      End With
    
  CmdCP.Enabled = True
  CmdIZR.Enabled = True
  CmdAuto.Enabled = True    '*** (Jen 3.01)
End Sub
Private Sub GrdIZR_ColWhen(ByVal RowNum As Integer, ByVal ColName As String, Cancel As Boolean, ErrMsg As String, ByVal Action As MwfCtl.en_fgAction)
  '*** Cannot edit any value of the grid GrdIZR as it is a working grid and is used only
      'for showing temporary information
      
  Cancel = True
End Sub
Private Sub GrdIZD_ColWhen(ByVal RowNum As Integer, ByVal ColName As String, Cancel As Boolean, ErrMsg As String, ByVal Action As MwfCtl.en_fgAction)
  '*** Cannot edit any value of the grid GrdIZR as it is a working grid and is used only
      'for showing temporary information
  
  Cancel = True
End Sub
Private Sub ADC_SetRecSource()
  adc.DataMode = 1
End Sub
Private Sub GrdIZ_SetRecSource()
  '*** Set the Record Source of the Grid GrdIZ (Table = InvZRm)
  '****** Sachin 3.02 31-12-07 - Corresponding Partition
  GrdIZ.RecSource = " Select * from InvZRm " + _
                    " Where IzCoCd= '" + gs_CoCd + "'" + _
                    " And IzTc ='" + adc("wIzTc") + "'" + _
                    " And IzYy = '" + adc("wIzYy") + "'" + _
                    " And IzChr ='" + adc("wIzChr") + "' " + _
                    " And IzNo =" + CStr(adc("wIzNo")) + " And IzPrtKey=(Case When '" + gs_Partition + "'='" + ctCurrPrtn + "' Then '" + ctCurrPrtn + "' Else IzPrtKey End) Order By IzSr"
End Sub
Private Sub GrdTxndZ_SetRecSource()
  '*** Set the Record Source of the Grid GrdTxndZ (Table = TxndZ)
  '****** Sachin 3.02 31-12-07 - Corresponding Partition
  GrdTxndZ.RecSource = " Select * from TxndZ " + _
                       " Where TzCoCd = '" + gs_CoCd + "'" + _
                       " And TzTc ='" + adc("wIzTc") + "'" + _
                       " And TzYy = '" + adc("wIzYy") + "'" + _
                       " And TzChr ='" + adc("wIzChr") + "' " + _
                       " And TzNo =" + CStr(adc("wIzNo")) + _
                       " And TzSr =" + CStr(GrdIZ.Value(GrdIZ.Row, "IzSr")) + " And TzPrtKey='" + adc("wIzPrtKey") + "' " + _
                       " And TzSrNo<> 0 Order By TzSrNo"
End Sub
Private Sub CmdIZR_Click()
  '*** Display or hide the Rm Details frame depending on the toggle state
      'If visible then call the RmDetails procedure
  Call DispFra(InvRmDet)
  If FraIZR.Visible = True Then
    Call RmDetails("GrdIZR")
  End If
End Sub
Private Sub CmdCP_Click()
  '*** Display or hide the Copy Frame frame depending on the toggle state
  '****** Sachin 3.02 31-12-07 - Locking Previous Partition Invoices
  If adc("wIzPrtKey") <> ctCurrPrtn = True Then DispMsg "Cannot Copy In Previous Partition Invoice", etError: Exit Sub
  If mf_DtLocked = True Then Exit Sub
  If mb_InCompInv Then Exit Sub     '''    '***** Manali 3.9.0 - Incomplete invoice Design
  
  '4.1.0.0
  If moCn.RecSeek("Select 'x' From InvFgd where " + _
                    " IfCoCd ='" + gs_CoCd + "' and IfTc= '" + adc("wIzTc") + "' and IfYy= '" + adc("wIzYy") + "' and IfChr= '" + adc("wIzChr") + "' and " + _
                    " IfNo= " + CStr(adc("wIzNo")) + " and IfSr=0 and IfPrtKey=(Case When '" + adc("WIZPRTKEY") + "'='" + ctCurrPrtn + "' Then '" + ctCurrPrtn + "' Else IfPrtKey End) ") Then _
                     DispMsg " IfSr=0 Record Exist cannot copy", etError: Exit Sub
            
  Call DispFra(Cpy)
End Sub
Private Sub CmdCpGo_Click()
  '*** Call the Copy procedure
  '*** Hide the Copy Frame
  
  Call CopyInvZRm
  Call CmdCP_Click
End Sub

'*** (Jen 3.01)
Private Sub CmdAuto_Click()
  '*** Display or hide the Auto Custom Matching Frame frame depending on the toggle state
  '****** Sachin 3.02 31-12-07 - Restrictions
  If adc("wIzPrtKey") <> ctCurrPrtn Then DispMsg "Auto Custom Matching Not Allowed In Previous Partition Invoice", etError: Exit Sub
  
  If mf_DtLocked = True Then Exit Sub
  
  If mb_InCompInv Then Exit Sub     '''    '***** Manali 3.9.0 - Incomplete invoice Design
  
  Call DispFra(AutoCustmMtch)
  
  If FraAuto.Visible = True And adc("wAutoUseOnlyCustGoods") = "" Then
    Dim wRs_Hd As MwfLib.MDORowSet
    
    adc("wAutoUseOnlyCustGoods") = "N"
    adc("wAutoBestUseOnly") = "N"     '*** (Jen 3.2.0)
    
    '********* Manali 3.2.2 Before date fields added
    Set wRs_Hd = moCn.OpenRes("Select HRtMinusD, HRtPlusD, HPtrMinusD, HPtrPlusD, HImpTypNotInSelD, HImpAftD, HImpBefD, " + _
                 "HRtMinusC, HRtPlusC, HPtrMinusC, HPtrPlusC, HImpTypNotInSelC, HImpAftC, HImpBefC, " + _
                 "HRtMinusG, HRtPlusG, HPtrMinusG, HPtrPlusG, HImpTypNotInSelG, HImpAftG, HImpBefG, " + _
                 "HRtMinusP, HRtPlusP, HPtrMinusP, HPtrPlusP, HImpTypNotInSelP, HImpAftP, HImpBefP, " + _
                 "HRtMinusS, HRtPlusS, HPtrMinusS, HPtrPlusS, HImpTypNotInSelS, HImpAftS, HImpBefS, " + _
                 "HRtMinusL, HRtPlusL, HPtrMinusL, HPtrPlusL, HImpTypNotInSelL, HImpAftL, HImpBefL, " + _
                 "HRtMinusX, HRtPlusX, HPtrMinusX, HPtrPlusX, HImpTypNotInSelX, HImpAftX, HImpBefX " + _
                 "From Head where HCoCd= '" + gs_CoCd + "'")
    
    If wRs_Hd.RecCount > 0 Then
      adc("wAutoHRtMinusD") = wRs_Hd!HRtMinusD
      adc("wAutoHRtPlusD") = wRs_Hd!HRtPlusD
      adc("wAutoHPtrMinusD") = wRs_Hd!HPtrMinusD
      adc("wAutoHPtrPlusD") = wRs_Hd!HPtrPlusD
      adc("wAutoHImpTypNotInSelD") = wRs_Hd!HImpTypNotInSelD
      adc("wAutoHImpAftD") = wRs_Hd!HImpAftD
      adc("wAutoHImpBefD") = wRs_Hd!HImpBefD    '****** Manali 3.2.2 - Before Date field added
      adc("wAutoHRtMinusC") = wRs_Hd!HRtMinusC
      adc("wAutoHRtPlusC") = wRs_Hd!HRtPlusC
      adc("wAutoHPtrMinusC") = wRs_Hd!HPtrMinusC
      adc("wAutoHPtrPlusC") = wRs_Hd!HPtrPlusC
      adc("wAutoHImpTypNotInSelC") = wRs_Hd!HImpTypNotInSelC
      adc("wAutoHImpAftC") = wRs_Hd!HImpAftC
      adc("wAutoHImpBefC") = wRs_Hd!HImpBefC   '****** Manali 3.2.2 - Before Date field added
      adc("wAutoHRtMinusG") = wRs_Hd!HRtMinusG
      adc("wAutoHRtPlusG") = wRs_Hd!HRtPlusG
      adc("wAutoHPtrMinusG") = wRs_Hd!HPtrMinusG
      adc("wAutoHPtrPlusG") = wRs_Hd!HPtrPlusG
      adc("wAutoHImpTypNotInSelG") = wRs_Hd!HImpTypNotInSelG
      adc("wAutoHImpAftG") = wRs_Hd!HImpAftG
      adc("wAutoHImpBefG") = wRs_Hd!HImpBefG    '****** Manali 3.2.2 - Before Date field added
      adc("wAutoHRtMinusP") = wRs_Hd!HRtMinusP
      adc("wAutoHRtPlusP") = wRs_Hd!HRtPlusP
      adc("wAutoHPtrMinusP") = wRs_Hd!HPtrMinusP
      adc("wAutoHPtrPlusP") = wRs_Hd!HPtrPlusP
      adc("wAutoHImpTypNotInSelP") = wRs_Hd!HImpTypNotInSelP
      adc("wAutoHImpAftP") = wRs_Hd!HImpAftP
      adc("wAutoHImpBefP") = wRs_Hd!HImpBefP    '****** Manali 3.2.2 - Before Date field added
      adc("wAutoHRtMinusS") = wRs_Hd!HRtMinusS
      adc("wAutoHRtPlusS") = wRs_Hd!HRtPlusS
      adc("wAutoHPtrMinusS") = wRs_Hd!HPtrMinusS
      adc("wAutoHPtrPlusS") = wRs_Hd!HPtrPlusS
      adc("wAutoHImpTypNotInSelS") = wRs_Hd!HImpTypNotInSelS
      adc("wAutoHImpAftS") = wRs_Hd!HImpAftS
      adc("wAutoHImpBefS") = wRs_Hd!HImpBefS    '****** Manali 3.2.2 - Before Date field added
      ' ***** Manali 3.03 - 23/06/08 - Pd Mod - 'L' fields added
      adc("wAutoHRtMinusL") = wRs_Hd!HRtMinusL
      adc("wAutoHRtPlusL") = wRs_Hd!HRtPlusL
      adc("wAutoHPtrMinusL") = wRs_Hd!HPtrMinusL
      adc("wAutoHPtrPlusL") = wRs_Hd!HPtrPlusL
      adc("wAutoHImpTypNotInSelL") = wRs_Hd!HImpTypNotInSelL
      adc("wAutoHImpAftL") = wRs_Hd!HImpAftL
      adc("wAutoHImpBefL") = wRs_Hd!HImpBefL
      ' ***** Manali 3.03 - 23/06/08 - Pd Mod - 'L' fields added
      adc("wAutoHRtMinusX") = wRs_Hd!HRtMinusX
      adc("wAutoHRtPlusX") = wRs_Hd!HRtPlusX
      adc("wAutoHPtrMinusX") = wRs_Hd!HPtrMinusX
      adc("wAutoHPtrPlusX") = wRs_Hd!HPtrPlusX
      adc("wAutoHImpTypNotInSelX") = wRs_Hd!HImpTypNotInSelX
      adc("wAutoHImpAftX") = wRs_Hd!HImpAftX
      adc("wAutoHImpBefX") = wRs_Hd!HImpBefX    '****** Manali 3.2.2 - Before Date field added
    End If
  End If
  
End Sub

Private Sub CmdAutoGo_Click()
  '*** Call the Auto Custom Matching procedure
  '*** Hide the Auto Custom Matching Frame
  
  '*** (Jen 3.2.0)
  If adc("wAutoBestUseOnly") = "" Then
    DispMsg "Specify 'Match Best Use Only' Option", etError: Exit Sub
  End If
  
  If adc("wAutoUseOnlyCustGoods") = "" Then
    DispMsg "Specify 'Use Only Customer Goods' Option", etError: Exit Sub
  End If
  '*** (Jen 3.2.0)
  
  '******* Manali 3.2.2
  If (adc("wAutoHImpAftD") <> MWLib.EmptyDate And adc("wAutoHImpBefD") <> MWLib.EmptyDate And adc("wAutoHImpAftD") > adc("wAutoHImpBefD")) _
    Or (adc("wAutoHImpAftC") <> MWLib.EmptyDate And adc("wAutoHImpBefC") <> MWLib.EmptyDate And adc("wAutoHImpAftC") > adc("wAutoHImpBefC")) _
    Or (adc("wAutoHImpAftG") <> MWLib.EmptyDate And adc("wAutoHImpBefG") <> MWLib.EmptyDate And adc("wAutoHImpAftG") > adc("wAutoHImpBefG")) _
    Or (adc("wAutoHImpAftP") <> MWLib.EmptyDate And adc("wAutoHImpBefP") <> MWLib.EmptyDate And adc("wAutoHImpAftP") > adc("wAutoHImpBefP")) _
    Or (adc("wAutoHImpAftS") <> MWLib.EmptyDate And adc("wAutoHImpBefS") <> MWLib.EmptyDate And adc("wAutoHImpAftS") > adc("wAutoHImpBefS")) _
    Or (adc("wAutoHImpAftL") <> MWLib.EmptyDate And adc("wAutoHImpBefL") <> MWLib.EmptyDate And adc("wAutoHImpAftL") > adc("wAutoHImpBefL")) _
    Or (adc("wAutoHImpAftX") <> MWLib.EmptyDate And adc("wAutoHImpBefX") <> MWLib.EmptyDate And adc("wAutoHImpAftX") > adc("wAutoHImpBefX")) Then
        DispMsg "Import Before Date(s) Cannot Be Less Than Import After Date(s)", etError: Exit Sub
  End If
  '******* Manali 3.2.2
  
  Call AutoCustomMatching                '*** (Jen 3.1.1)
  'Call AutoCustomMatchingAllRecsTogether  '*** (Bef 3.1.1)
  Call CmdAuto_Click
End Sub
'*** (Jen 3.01)


'*** (Bef 3.01)
'Private Sub CopyInvZRm()
''*** This option is used to Copy records into the grid GrdIZ
'    '(the recs in the grid are unique for Custom code & Pure Rate combination)
''*** The copy is possible even after records exists in the grid GrdIZ. But it is possible
'    'only for the custom code, Pure Rate combination which does not already exist in the grid.
''*** The records are copied one line per Custom Code, Custom Pure Rate
''*** The Weights are converted into pure Weight
'
'  Dim wRsIZ As MwfLib.MDORowSet, wRow As Single, wSelCnd As String
'  'If GrdIZ.Rows - 1 > 0 Then DispMsg "Record Already Exists", etError: Exit Sub
'  '*** (Bef 2.13)
'  'ADC("wCpRmCtgFr").CmpStr = "RmCtg>= "
'  'ADC("wCpRmCtgTo").CmpStr = "RmCtg<= "
'  '*** (Bef 2.13)
'
'  '*** (Jen 2.13)
'  ADC("wCpRmCtgFr").CmpStr = "R1.RmCtg>= "
'  ADC("wCpRmCtgTo").CmpStr = "R1.RmCtg<= "
'  '*** (Jen 2.13)
'
'  wSelCnd = ADC.RepCond
'  wSelCnd = IIF(wSelCnd <> "", " and " + wSelCnd, "")
'
''  Set wRsIZ = moCn.OpenRes("Select IrRmZCd, sum(IrRmQty) as qRmQty, " + _
''              "sum(IrRmZWt* (case when RmCtg in ('G', 'P', 'S') then RmPurityWt* (1 + IdGldLs/100)  else 1 end)/ " + _
''              "(case when PNum1> 0 then PNum1 else 1 end)) as qRmZWt, " + _
''              "IrRmZRt from InvRm join InvDsg on IdTc= IrTc and IdYy= IrYy and " + _
''              "IdChr= IrChr and IdNo= IrNo and IdSr= IrSr join RmMst on RmCd= IrRmCd " + _
''              "join Param on PTyp= 'RMCTG' and PMCd= RmCtg " + _
''              "where IrTc= '" + ADC("wIzTc") + "' and IrYy= '" + ADC("wIzYy") + "' and " + _
''              "IrChr= '" + ADC("wIzChr") + "' and IrNo= " + CStr(ADC("wIzNo")) + " and " + _
''              "IrRmZRt> 0 and IrRmZWt> 0 and " + _
''              "not exists(Select 'x' from InvZRm where IzTc='" + ADC("wIzTc") + "' and " + _
''              "IzYy= '" + ADC("wIzYy") + "' and IzChr= '" + ADC("wIzChr") + "' and " + _
''              "IzNo= " + CStr(ADC("wIzNo")) + " and IzRmZCd= IrRmZCd) " + wSelCnd + _
''              " Group by IrRmZCd, IrRmZRt Order By IrRmZCd, IrRmZRt")
'
'  '*** (Bef 2.13)
'  ''*** Jen (2.12 patch)
'  'ms_LocalExp = moCn.GetFldVal("Select hExpLocal from Head where HCoCd= '" + gs_CoCd + "' " + _
'  '                             "and HCd= '" + ctSelfCmCd + "' ")
'  'If ms_LocalExp = "Y" Then
'  '  Set wRsIZ = moCn.OpenRes("Select IrRmZCd, sum(IrRmQty) as qRmQty, " + _
'  '              "sum(IrRmZWt* (case when RmCtg in ('G', 'P', 'S') then RmPurityZ* (1 + IdGldLs/100)  else 1 end) " + _
'  '              "/ (case when (RmCtg= 'G' and InHGldFiness> 0) then InHGldFiness " + _
'  '              "        when (RmCtg= 'P' and InHPlFiness> 0) then InHPlFiness " + _
'  '              "        when (RmCtg= 'S' and InHSlFiness> 0) then InHSlFiness else 1 end) ) as qRmZWt, " + _
'  '              "IrRmZPureRt from InvRm " + _
'  '              "join InvDsg on IdCoCd= IrCoCd and IdTc= IrTc and IdYy= IrYy and " + _
'  '              "IdChr= IrChr and IdNo= IrNo and IdSr= IrSr join RmMst on RmCd= IrRmCd " + _
'  '              "join InvHd on InCoCd= IrCoCd and InTc= IrTc and InYy= IrYy and " + _
'  '              "InChr= IrChr and InNo= IrNo " + _
'  '              "where IrCoCd= '" + gs_CoCd + "' and " + _
'  '              "IrTc= '" + ADC("wIzTc") + "' and IrYy= '" + ADC("wIzYy") + "' and " + _
'  '              "IrChr= '" + ADC("wIzChr") + "' and IrNo= " + CStr(ADC("wIzNo")) + " and " + _
'  '              "IrRmZWt> 0 and " + _
'  '              "not exists(Select 'x' from InvZRm where " + _
'  '              "IzCoCd= '" + gs_CoCd + "' and IzTc='" + ADC("wIzTc") + "' and " + _
'  '              "IzYy= '" + ADC("wIzYy") + "' and IzChr= '" + ADC("wIzChr") + "' and " + _
'  '              "IzNo= " + CStr(ADC("wIzNo")) + " and IzRmZCd= IrRmZCd and Round(IzRmZRt, 3)= Round(IrRmZPureRt, 3)) " + wSelCnd + _
'  '              " Group by IrRmZCd, IrRmZPureRt Order By IrRmZCd, IrRmZPureRt")
'  'Else
'  ''*** Jen (2.12 patch)
'  '
'  '  Set wRsIZ = moCn.OpenRes("Select IrRmZCd, sum(IrRmQty) as qRmQty, " + _
'  '              "sum(IrRmZWt* (case when RmCtg in ('G', 'P', 'S') then RmPurityZ* (1 + IdGldLs/100)  else 1 end) " + _
'  '              ") as qRmZWt, IrRmZPureRt from InvRm " + _
'  '              "join InvDsg on IdCoCd= IrCoCd and IdTc= IrTc and IdYy= IrYy and " + _
'  '              "IdChr= IrChr and IdNo= IrNo and IdSr= IrSr join RmMst on RmCd= IrRmCd " + _
'  '              "where IrCoCd= '" + gs_CoCd + "' and " + _
'  '              "IrTc= '" + ADC("wIzTc") + "' and IrYy= '" + ADC("wIzYy") + "' and " + _
'  '              "IrChr= '" + ADC("wIzChr") + "' and IrNo= " + CStr(ADC("wIzNo")) + " and " + _
'  '              "IrRmZWt> 0 and " + _
'  '              "not exists(Select 'x' from InvZRm where " + _
'  '              "IzCoCd= '" + gs_CoCd + "' and IzTc='" + ADC("wIzTc") + "' and " + _
'  '              "IzYy= '" + ADC("wIzYy") + "' and IzChr= '" + ADC("wIzChr") + "' and " + _
'  '              "IzNo= " + CStr(ADC("wIzNo")) + " and IzRmZCd= IrRmZCd and Round(IzRmZRt, 3)= Round(IrRmZPureRt, 3)) " + wSelCnd + _
'  '              " Group by IrRmZCd, IrRmZPureRt Order By IrRmZCd, IrRmZPureRt")
'  '
'  ''*** Jen (2.12 patch)
'  'End If
'  ''*** Jen (2.12 patch)
'  '*** (Bef 2.13)
'
'  '*** (Jen 2.13)
'  If ms_LocalExp = "Y" Then
'    '           "sum(IrRmZWt* (case when R1.RmCtg in ('G', 'P', 'S') then R1.RmPurityZ* (1 + IdGldLs/100) " + _
'    '            "                         / (case when R2.RmPurityZ> 0 then R2.RmPurityZ else 1 end) else 1 end) " + _
'    '            "/ (case when (R1.RmCtg= 'G' and InHGldFiness> 0) then InHGldFiness " + _
'    '            "        when (R1.RmCtg= 'P' and InHPlFiness> 0) then InHPlFiness " + _
'    '            "        when (R1.RmCtg= 'S' and InHSlFiness> 0) then InHSlFiness else 1 end) ) as qRmZWt, "
'
'    Set wRsIZ = moCn.OpenRes("Select IrRmZCd, sum(IrRmQty) as qRmQty, " + _
'                "sum((case when IrMainMet= 'Y' then IdCustmMtchWt else IrRmZWt end)* " + _
'                "(case when R1.RmCtg in ('G', 'P', 'S') then R1.RmPurityZ " + _
'                "                         * (1 + (case when IrMainMet= 'Y' then 0 else IdGldLs end)/100) " + _
'                "                         / (case when R2.RmPurityZ> 0 then R2.RmPurityZ else 1 end) else 1 end) " + _
'                "/ (case when (R1.RmCtg= 'G' and InHGldFiness> 0) then InHGldFiness " + _
'                "        when (R1.RmCtg= 'P' and InHPlFiness> 0) then InHPlFiness " + _
'                "        when (R1.RmCtg= 'S' and InHSlFiness> 0) then InHSlFiness else 1 end) ) as qRmZWt, " + _
'                "IrRmZPureRt from InvRm " + _
'                "join InvDsg on IdCoCd= IrCoCd and IdTc= IrTc and IdYy= IrYy and " + _
'                "IdChr= IrChr and IdNo= IrNo and IdSr= IrSr " + _
'                "join RmMst R1 on R1.RmCd= IrRmCd " + _
'                "join RmMst R2 on R2.RmCd= R1.RmZCd " + _
'                "join InvHd on InCoCd= IrCoCd and InTc= IrTc and InYy= IrYy and " + _
'                "InChr= IrChr and InNo= IrNo " + _
'                "where IrCoCd= '" + gs_CoCd + "' and " + _
'                "IrTc= '" + ADC("wIzTc") + "' and IrYy= '" + ADC("wIzYy") + "' and " + _
'                "IrChr= '" + ADC("wIzChr") + "' and IrNo= " + CStr(ADC("wIzNo")) + " and " + _
'                "IrRmZWt> 0 and " + _
'                "not exists(Select 'x' from InvZRm where " + _
'                "IzCoCd= '" + gs_CoCd + "' and IzTc='" + ADC("wIzTc") + "' and " + _
'                "IzYy= '" + ADC("wIzYy") + "' and IzChr= '" + ADC("wIzChr") + "' and " + _
'                "IzNo= " + CStr(ADC("wIzNo")) + " and IzRmZCd= IrRmZCd and Round(IzRmZRt, 3)= Round(IrRmZPureRt, 3)) " + wSelCnd + _
'                " Group by IrRmZCd, IrRmZPureRt Order By IrRmZCd, IrRmZPureRt")
'  Else
'
'    '"sum(IrRmZWt* (case when R1.RmCtg in ('G', 'P', 'S') then R1.RmPurityZ* (1 + IdGldLs/100) " + _
'    '            "/ (case when R2.RmPurityZ> 0 then R2.RmPurityZ else 1 end) else 1 end) " + _
'    '            ") as qRmZWt "
'
'    Set wRsIZ = moCn.OpenRes("Select IrRmZCd, sum(IrRmQty) as qRmQty, " + _
'                "sum((case when IrMainMet= 'Y' then IdCustmMtchWt else IrRmZWt end) " + _
'                "   * (case when R1.RmCtg in ('G', 'P', 'S') then R1.RmPurityZ " + _
'                "       * (1 + (case when IrMainMet= 'Y' then 0 else IdGldLs end)/100) " + _
'                "       / (case when R2.RmPurityZ> 0 then R2.RmPurityZ else 1 end) else 1 end) " + _
'                ") as qRmZWt, IrRmZPureRt from InvRm " + _
'                "join InvDsg on IdCoCd= IrCoCd and IdTc= IrTc and IdYy= IrYy and " + _
'                "IdChr= IrChr and IdNo= IrNo and IdSr= IrSr " + _
'                "join RmMst R1 on R1.RmCd= IrRmCd " + _
'                "join RmMst R2 on R2.RmCd= R1.RmZCd " + _
'                "where IrCoCd= '" + gs_CoCd + "' and " + _
'                "IrTc= '" + ADC("wIzTc") + "' and IrYy= '" + ADC("wIzYy") + "' and " + _
'                "IrChr= '" + ADC("wIzChr") + "' and IrNo= " + CStr(ADC("wIzNo")) + " and " + _
'                "IrRmZWt> 0 and " + _
'                "not exists(Select 'x' from InvZRm where " + _
'                "IzCoCd= '" + gs_CoCd + "' and IzTc='" + ADC("wIzTc") + "' and " + _
'                "IzYy= '" + ADC("wIzYy") + "' and IzChr= '" + ADC("wIzChr") + "' and " + _
'                "IzNo= " + CStr(ADC("wIzNo")) + " and IzRmZCd= IrRmZCd and Round(IzRmZRt, 3)= Round(IrRmZPureRt, 3)) " + wSelCnd + _
'                " Group by IrRmZCd, IrRmZPureRt Order By IrRmZCd, IrRmZPureRt")
'  End If
'  '*** (Jen 2.13)
'
'  '*** Begin Copying InvZRm Records
'  GrdIZ.StartCopy
'  MWLib.BeginProcess Me, "Copying Custom Codes ..."
'  With GrdIZ
'    Do While Not (wRsIZ.EOF Or wRsIZ.BOF)
'      .AddItem
'      wRow = .Rows - 1
'      .Value(wRow, "IzSr") = .MaxVal("IzSr") + 1
'      .Value(wRow, "IzRmZCd") = wRsIZ!IrRmZCd
'      .Value(wRow, "IzRmQty") = wRsIZ!qRmQty
'      .Value(wRow, "IzRmZWt") = wRsIZ!qRmZWt
'      .Value(wRow, "IzRmZRt") = wRsIZ!IrRmZPureRt
'      .Value(wRow, "wIzRmBalWt") = .Value(wRow, "IzRmZWt")
'      .SaveRec (wRow)
'      wRsIZ.MoveNext
'    Loop
'  End With
'  MWLib.EndProcess Me
'  GrdIZ.EndCopy
'  Set wRsIZ = Nothing
'  '*** End of Copying InvZRm Records
'End Sub
'*** (Bef 3.01)

Private Sub CopyInvZRm()
  '*** (Jen 3.01)
  '*** This option is used to Copy records into the grid GrdIZ
      '(the recs in the grid are unique for Custom code & Pure Rate combination)
  '*** The copy is possible even after records exists in the grid GrdIZ. But it is possible
      'only for the custom code, Pure Rate combination which does not already exist in the grid.
  '*** The records are copied one line per Custom Code, Custom Pure Rate
  '*** The Weights are converted into pure Weight
  
  Dim wSelCnd As String

  adc("wCpRmCtgFr").CmpStr = "R1.RmCtg>= "
  adc("wCpRmCtgTo").CmpStr = "R1.RmCtg<= "
  
  wSelCnd = adc.RepCond
  wSelCnd = IIF(wSelCnd <> "", " and " + wSelCnd, "")
  Call GetInvZRmRecs(wSelCnd)
  '*** (Jen 3.01)
End Sub

Private Sub AutoCustomMatching()
  '*** (Jen 3.1.2)
  '*** This option is used to Copy records into the grid GrdIZ & GrdTxndZ
      '(the recs in the grid are unique for Custom code & Pure Rate combination)
  '*** The copy is possible even after records exists in the grid GrdIZ. But it is possible
      'only for the custom code, Pure Rate combination which does not already exist in the grid.
  '*** The records are copied one line per Custom Code, Custom Pure Rate
  '*** The Weights are converted into pure Weight
  
  Dim wf_StartRow As Single, wf_EndRow As Single, wf_IZRow As Single
  Dim wb_ChildExists As Boolean
  
  Dim ws_str As String, wRs_Txndz As MwfLib.MDORowSet, wf_Row As Single
  Dim ws_RmCtg As String, ws_ZRmCd As String, wd_Rt As Double, wd_Wt As Double, wf_Qty As Single, wd_BalWt As Double
  Dim wb_BestUse As Boolean
  
  Dim wb_BestUseOnly As Boolean   '*** (Jen 3.2.0)
  
  wb_BestUse = moCn.RecSeek("Select 'x' from Head where HCoCd= '" + gs_CoCd + "' and " + "HCd= '" + ctSelfCmCd + "' and hBestUse= 'Y'")
  wb_BestUseOnly = IIF(adc("wAutoBestUseOnly") = "Y", True, False)    '*** (Jen 3.2.0)
  
  wf_StartRow = 1
  wf_EndRow = GrdIZ.Rows - 1
  If wf_EndRow = 0 Then Exit Sub
  
  For wf_IZRow = wf_StartRow To wf_EndRow
    GrdIZ.Row = wf_IZRow
    
    wb_ChildExists = IIF(GrdTxndZ.Rows - 1 > 0, True, False)
    
    '*** (Jen 3.2.0)
    If wb_BestUse = True And wb_ChildExists = False Then
      Call AutoCustomMatchingWithBestUse(wf_IZRow)
    Else
      If Not (wb_BestUseOnly) Then
        Call AutoCustomMatchingNormal(wf_IZRow)
      End If
    End If
    '*** (Jen 3.2.0)
    
    '*** (Bef 3.2.0)
    'If wb_BestUse = True And wb_ChildExists = False Then
    '  Call AutoCustomMatchingWithBestUse(wf_IZRow)
    'Else
    '  Call AutoCustomMatchingNormal(wf_IZRow)
    'End If
    '*** (Bef 3.2.0)
    
  Next wf_IZRow
  '*** (Jen 3.1.2)
End Sub
Private Sub AutoCustomMatchingWithBestUse(ByVal pi_Row As Integer)
  '*** (Jen 3.1.2)
  Dim ws_str As String, wRs_InvDets As MwfLib.MDORowSet, wRs_Txndz As MwfLib.MDORowSet, wf_Row As Single
  Dim ws_RmCtg As String, ws_ZRmCd As String, wd_Rt As Double, wd_Wt As Double, wf_Qty As Single, wd_BalWt As Double
  Dim wd_AssignWt As Double
  
  Dim ws_BestUseOnlyCnd As String     '*** (Jen 3.2.0)
  
  ws_ZRmCd = GrdIZ.Value(pi_Row, "IzRmZCd")
  ws_RmCtg = moCn.GetFldVal("Select RmCtg from RmMst where RmPrtKey='" + ctCurrPrtn + "' and RmCd= '" + ws_ZRmCd + "'")
  wd_Rt = GrdIZ.Value(pi_Row, "IzRmZRt")
  wd_Wt = GrdIZ.Value(pi_Row, "wIzRmBalWt")
  wf_Qty = GrdIZ.Value(pi_Row, "IzRmQty")
  '### wd_BalWt = wd_Wt
  
  '****** (Jen 3.1.3) added qSr ******
  '****** Sachin 3.02 31-12-07 - Current Partition
  '***** Manali 3.03 - 23/06/08 - Pd Mod - 'L' added
  ws_str = "Select ROW_NUMBER() OVER(ORDER BY IdExpOdTc, IdExpOdYy, IdExpOdChr, IdExpOdNo, IdDmCd, IrRmCd, IrRmSz) as qSr, " + _
           "  IdExpOdTc, IdExpOdYy, IdExpOdChr, IdExpOdNo, IdDmCd, IrRmCd, IrRmSz, " + _
           "  convert(decimal (14, 3), sum((case when IrMainMet= 'Y' then IdCustmMtchWt else IrRmZWt end) " + _
           "    * (case when R1.RmCtg in ('G', 'P', 'S', 'L') then R1.RmPurityZ " + _
           "    * (1 + (case when IrMainMet= 'Y' then 0 else IdGldLs end)/100) " + _
           IIF(ms_LocalExp = "N", "", "    / (case when R2.RmPurityZ> 0 then R2.RmPurityZ else 1 end) else 1 end) " + _
           "    / (case when (R1.RmCtg= 'G' and InHGldFiness> 0) then InHGldFiness " + _
           "        when (R1.RmCtg= 'P' and InHPlFiness> 0) then InHPlFiness " + _
           "        when (R1.RmCtg= 'S' and InHSlFiness> 0) then InHSlFiness " + _
           "        when (R1.RmCtg= 'L' and InHPdFiness> 0) then InHPdFiness ") + " else 1 end) )) as qRmZWt " + _
           "into #TmpInv From InvRm " + _
           "join InvDsg on IdIdNo= IrIdIdNo and IdCoCd= IrCoCd and IdTc= IrTc and IdYy= IrYy and IdChr= IrChr " + _
           "  and IdNo= IrNo and IdSr= IrSr and IdPrtKey= IrPrtKey " + _
           "join InvHd on InIdNo= IdInIdNo and InCoCd= IdCoCd and InTc= IdTc and InYy= IdYy and InChr= IdChr " + _
           "  and InNo= IdNo and InPrtKey= IdPrtKey " + _
           "join RmMst R1 on R1.RmPrtKey=IrPrtKey and R1.RmCd= IrRmCd " + _
           "join RmMst R2 on R1.RmPrtKey=R2.RmPrtKey and R2.RmCd= R1.RmZCd " + _
           "where IrCoCd= '" + gs_CoCd + "' and IrTc= '" + adc("wIzTc") + "' and IrYy= '" + adc("wIzYy") + "' " + _
           "  and IrChr= '" + adc("wIzChr") + "' and IrNo= " + CStr(adc("wIzNo")) + " " + _
           "  and IrRmZCd= '" + ws_ZRmCd + "' " + _
           "  and convert(decimal (14, 3), IrRmZPureRt)= convert(decimal (14, 3), " + CStr(wd_Rt) + ") " + _
           "  and IrPrtKey= '" + ctCurrPrtn + "' " + _
           "group By IdExpOdTc, IdExpOdYy, IdExpOdChr, IdExpOdNo, IdDmCd, IrRmCd, IrRmSz "
  
  'Set wRs_InvDets = moCn.OpenRes(ws_str)
  moCn.CreateTmpTable "#TmpInv", ws_str
  
  '*** (Jen 3.1.3)
  Dim wd_Diff As Double
  wd_Diff = Round(wd_Wt - moCn.GetFldVal("Select sum(qRmZWt) from #TmpInv"), 3)
  If wd_Diff <> 0 Then
    moCn.Execute ("update #TmpInv set qRmZWt= qRmZWt+ convert(decimal (14, 3), " + CStr(wd_Diff) + ") where qSr= 1")
  End If
  Set wRs_InvDets = moCn.OpenResultset("Select * from #TmpInv Order By qSr")
  '*** (Jen 3.1.3)
  
  '*** (Bef 3.1.3)
  'Set wRs_InvDets = moCn.OpenResultset("Select * from #TmpInv Order By IdExpOdTc, IdExpOdYy, IdExpOdChr, IdExpOdNo, IdDmCd, IrRmCd, IrRmSz")
  '*** (Bef 3.1.3)
    
  '*** (Jen 3.2.0)
  ws_BestUseOnlyCnd = IIF(adc("wAutoBestUseOnly") = "N", " ", _
                      " and not (TzOdTc= '' and TzDmCd= '' and TzRmCd= '' and TzRmSzFr= 0 and TzRmSzTo= 0) ")
  '*** (Jen 3.2.0)
    
  '****** Sachin 3.02 31-12-07 - Current Partition
  ws_str = "/* set dateformat dmy */ " + _
           "Select (case when TzOdNo= 0 then '1' else '0' end)+ (case when TzDmCd= '' then '1' else '0' end) " + _
           "       + (case when TzRmCd= '' then '1' else '0' end) " + _
           "       + (case when TzRmSzFr= 0 and TzRmSzTo= 0 then '1' else '0' end) " + _
           "       + (case when TzRmSzFr= 0 then '1' else '0' end)+ (case when TzRmSzTo= 0 then '1' else '0' end) as qTag, " + _
           "    T.TDt as TDt, TzOdTc, TzOdYy, TzOdChr, TzOdNo, TzDmCd, TzRmCd, TzRmSzFr, TzRmSzTo, TzTc, TzYy, TzChr, " + _
           "    TzNo, TzSr, TzKey, TzTc+ '/'+ TzYy+ '/'+ TzChr+ '/'+ LTrim(Str(TzNo))+ '/'+ LTrim(Str(TzSr)) as qBaseVch, " + _
           "    TBOENo, convert (decimal (16, 3), (TzRmWt- TzRmDespWt)) as qBalWt, " + _
           "    convert (decimal (14, 3), convert (decimal (14, 3), TzPurRt) " + _
           "          * (case when TMulDiv= 'D' then convert (decimal (14, 4), TCnvRt) else 1.0000 end) " + _
           "          / (case when TMulDiv= 'M' then convert (decimal (14, 4), TCnvRt) else 1.0000 end) " + _
           "          * " + Format(IIF(ms_MulDiv = "M", CStr(md_CnvFct), "1"), "0.0000") + " " + _
           "          / " + Format(IIF(ms_MulDiv = "D", CStr(md_CnvFct), "1"), "0.0000") + ") As qPurRtInInvCur, " + _
           "    convert (decimal (14, 3), convert (decimal (14, 3), TzPurRt) " + _
           "          * (case when TMulDiv= 'D' then convert (decimal (14, 4), TCnvRt) else 1.0000 end) " + _
           "          / (case when TMulDiv= 'M' then convert (decimal (14, 4), TCnvRt) else 1.0000 end) " + _
           "          * " + Format(IIF(ms_MulDiv = "M", md_CnvFct, 1), "0.0000") + " " + _
           "          / " + Format(IIF(ms_MulDiv = "D", md_CnvFct, 1), "0.0000") + " " + _
           "          * " + Format(IIF(ms_RsMulDiv = "M", md_RsCnvFct, 1), "0.0000") + " " + _
           "          / " + Format(IIF(ms_RsMulDiv = "D", md_RsCnvFct, 1), "0.0000") + ") As qPurRtInRs, " + _
           "    convert (decimal (14, 3), 0) as qAssignedWt "
           
  '****** (Jen 3.2.0) added ws_BestUseOnlyCnd
  ws_str = ws_str + " into #TmpBestUse from Txndz Tz " + _
           "join Txn T on " + IIF(gs_Partition = ctCurrPrtn, " TPrtKey=TzPrtKey and ", "") + " TIdNo= TzTIdNo and TCoCd= TzCoCd and TTc= TzTc and TYy= TzYy and TChr= TzChr and TNo= TzNo " + _
           "where TzCls= 'N' and TzRmZCd= '" + ws_ZRmCd + "' and TzCoCd= '" + gs_CoCd + "' and TzPrtKey= '" + ctCurrPrtn + "' " + _
           IIF(adc("wAutoUseOnlyCustGoods") = "Y", " and TCmCd= '" + adc("wIzCmCd") + "' ", " and TCmCd in ('', '" + adc("wIzCmCd") + "') ") + _
           "    " + GetAutoCustomMatchCnd(ws_RmCtg, "", wd_Rt, wd_Wt, wf_Qty) + " " + _
           "  and exists(Select 'x' from #TmpInv " + _
           "             where ( (IdExpOdTc= TzOdTc and IdExpOdYy= TzOdYy and IdExpOdChr= TzOdChr and IdExpOdNo= TzOdNo) " + _
           "                    or  (TzOdTc= '' and TzOdYy= '' and TzOdChr= '' and TzOdNo= 0) " + _
           "                    ) " + _
           "                and (IdDmCd= TzDmCd or TzDmCd= '') and (IrRmCd= TzRmCd or TzRmCd= '') " + _
           "                and (IrRmSz>= TzRmSzFr or TzRmSzFr= 0 or IrRmSz= 0) " + _
           "                and (IrRmSz<= TzRmSzTo or TzRmSzTo= 0 or IrRmSz= 0) " + ws_BestUseOnlyCnd + " " + _
           "             ) " + _
           "/* Order By (case when TzOdNo= 0 then '1' else '0' end)+ (case when TzDmCd= '' then '1' else '0' end) " + _
           "       + (case when TzRmCd= '' then '1' else '0' end) " + _
           "       + (case when TzRmSzFr= 0 and TzRmSzTo= 0 then '1' else '0' end) " + _
           "       + (case when TzRmSzFr= 0 then '1' else '0' end)+ (case when TzRmSzTo= 0 then '1' else '0' end), " + _
           "  TzOdTc, TzOdYy, TzOdChr, TzOdNo, TzDmCd, TzRmCd, TzRmSzFr, TzRmSzTo */ "
           
           
  'ws_str = ws_str + " into #TmpBestUse from Txndz Tz " + _
  '         "join Txn T on TCoCd= TzCoCd and TTc= TzTc and TYy= TzYy and TChr= TzChr and TNo= TzNo " + _
  '         "where TzCls= 'N' and TzRmZCd= '" + ws_ZRmCd + "' and TzCoCd= '" + gs_CoCd + "' " + _
  '         IIF(ADC("wAutoUseOnlyCustGoods") = "Y", " and TCmCd= '" + ADC("wIzCmCd") + "' ", " and TCmCd in ('', '" + ADC("wIzCmCd") + "') ") + _
  '         "    " + GetAutoCustomMatchCnd(ws_RmCtg, "", wd_Rt, wd_Wt, wf_Qty) + " " + _
  '         "  and exists(Select 'x' from InvDsg " + _
  '         "             join InvRm on IrCoCd= IdCoCd and IrTc= IdTc and IrYy= IdYy and IrChr= IdChr and IrNo= IdNo " + _
  '         "                and IrSr= IdSr " + _
  '         "             join RmMst on RmCd= IrRmCd " + _
  '         "             where IdCoCd= '" + gs_CoCd + "' and IdTc= '" + ADC("wIzTc") + "' and IdYy= '" + ADC("wIzYy") + "' " + _
  '         "                and IdChr= '" + ADC("wIzChr") + "' and IdNo= " + CStr(ADC("wIzNo")) + " " + _
  '         "                and IrRmZCd= '" + ws_ZRmCd + "' " + _
  '         "                and convert(decimal (14, 3), IrRmZPureRt)= convert(decimal (14, 3), " + CStr(wd_Rt) + ") " + _
  '         "                and ( (IdExpOdTc= TzOdTc and IdExpOdYy= TzOdYy and IdExpOdChr= TzOdChr and IdExpOdNo= TzOdNo) " + _
  '         "                    or  (TzOdTc= '' and TzOdYy= '' and TzOdChr= '' and TzOdNo= 0) " + _
  '         "                    ) " + _
  '         "                and (IdDmCd= TzDmCd or TzDmCd= '') and (IrRmCd= TzRmCd or TzRmCd= '') " + _
  '         "                and (IrRmSz>= TzRmSzFr or TzRmSzFr= 0 or IrRmSz= 0) " + _
  '         "                and (IrRmSz<= TzRmSzTo or TzRmSzTo= 0 or IrRmSz= 0) " + _
  '         "             ) " + _
  '         "Order By (case when TzOdNo= 0 then '1' else '0' end)+ (case when TzDmCd= '' then '1' else '0' end) " + _
  '         "       + (case when TzRmCd= '' then '1' else '0' end) " + _
  '         "       + (case when TzRmSzFr= 0 and TzRmSzTo= 0 then '1' else '0' end) " + _
  '         "       + (case when TzRmSzFr= 0 then '1' else '0' end)+ (case when TzRmSzTo= 0 then '1' else '0' end), " + _
  '         "  TzOdTc, TzOdYy, TzOdChr, TzOdNo, TzDmCd, TzRmCd, TzRmSzFr, TzRmSzTo "
  
  'ws_str = ws_str + " into #TmpBestUse from Txndz Tz " + _
  '         "join Txn T on TCoCd= TzCoCd and TTc= TzTc and TYy= TzYy and TChr= TzChr and TNo= TzNo " + _
  '         "where TzCls= 'N' and TzRmZCd= '" + ws_ZRmCd + "' and TzCoCd= '" + gs_CoCd + "' " + _
  '         IIF(ADC("wAutoUseOnlyCustGoods") = "Y", " and TCmCd= '" + ADC("wIzCmCd") + "' ", " and TCmCd in ('', '" + ADC("wIzCmCd") + "') ") + _
  '         "    " + GetAutoCustomMatchCnd(ws_RmCtg, "", wd_Rt, wd_Wt, wf_Qty) + " " + _
  '         "  and ( (TzOdTc= '" + wRs_InvDets!IdExpOdTc + "' and TzOdYy= '" + wRs_InvDets!IdExpOdYy + "' " + _
  '         "        and TzOdChr= '" + wRs_InvDets!IdExpOdChr + "' " + _
  '         "        and TzOdNo= " + CStr(wRs_InvDets!IdExpOdNo) + ") " + _
  '         "      or  (TzOdTc= '' and TzOdYy= '' and TzOdChr= '' and TzOdNo= 0) " + _
  '         "      ) " + _
  '         "  and (TzDmCd= '" + wRs_InvDets!IdDmCd + "' or TzDmCd= '') " + _
  '         "  and (TzRmCd= '" + wRs_InvDets!IrRmCd + "' or TzRmCd= '') " + _
  '         "  and (convert(decimal (7, 4), TzRmSzFr)<= convert(decimal (7, 4), " + CStr(wRs_InvDets!IrRmSz) + ") " + _
  '         "      or convert(decimal (7, 4), TzRmSzFr)= 0 or convert(decimal (7, 4), " + CStr(wRs_InvDets!IrRmSz) + ")= 0) " + _
  '         "  and (convert(decimal (7, 4), TzRmSzTo)>= convert(decimal (7, 4), " + CStr(wRs_InvDets!IrRmSz) + ") " + _
  '         "      or convert(decimal (7, 4), TzRmSzTo)= 0 or convert(decimal (7, 4), " + CStr(wRs_InvDets!IrRmSz) + ")= 0) " + _
  '         "Order By (case when TzOdNo= 0 then '1' else '0' end)+ (case when TzDmCd= '' then '1' else '0' end) " + _
  '         "       + (case when TzRmCd= '' then '1' else '0' end) " + _
  '         "       + (case when TzRmSzFr= 0 and TzRmSzTo= 0 then '1' else '0' end) " + _
  '         "       + (case when TzRmSzFr= 0 then '1' else '0' end)+ (case when TzRmSzTo= 0 then '1' else '0' end), " + _
  '         "  TzOdTc, TzOdYy, TzOdChr, TzOdNo, TzDmCd, TzRmCd, TzRmSzFr, TzRmSzTo "
  
  moCn.CreateTmpTable "#TmpBestUse", ws_str
  Set wRs_Txndz = moCn.OpenResultset("Select * from #TmpBestUse " + _
                  "Order By qTag, TzOdTc, TzOdYy, TzOdChr, TzOdNo, TzDmCd, TzRmCd, TzRmSzFr, TzRmSzTo, TDt, TzYy, TzKey")
  
  With wRs_InvDets
    Do While Not (.EOF Or .BOF)
      wd_BalWt = wRs_InvDets!qRmZWt
      If wRs_Txndz.RecCount > 0 Then wRs_Txndz.MoveFirst
      Do While Round(wd_BalWt, 3) > 0 And Not (wRs_Txndz.EOF Or wRs_Txndz.BOF)
        'If wRs_Txndz!qBalWt - wRs_Txndz!qAssignedWt > 0 Then
        If wRs_Txndz!qBalWt - wRs_Txndz!qAssignedWt > 0 _
            And ((wRs_Txndz!TzOdTc = wRs_InvDets!IdExpOdTc And wRs_Txndz!TzOdYy = wRs_InvDets!IdExpOdYy _
                     And wRs_Txndz!TzOdChr = wRs_InvDets!IdExpOdChr _
                     And wRs_Txndz!TzOdNo = wRs_InvDets!IdExpOdNo) _
                   Or (wRs_Txndz!TzOdTc = "" And wRs_Txndz!TzOdYy = "" And wRs_Txndz!TzOdChr = "" And wRs_Txndz!TzOdNo = 0) _
                 ) _
            And (wRs_Txndz!TzDmCd = wRs_InvDets!IdDmCd Or wRs_Txndz!TzDmCd = "") _
            And (wRs_Txndz!TzRmCd = wRs_InvDets!IrRmCd Or wRs_Txndz!TzRmCd = "") _
            And (wRs_Txndz!TzRmSzFr <= wRs_InvDets!IrRmSz Or wRs_Txndz!TzRmSzFr = 0 Or wRs_InvDets!IrRmSz = 0) _
            And (wRs_Txndz!TzRmSzTo >= wRs_InvDets!IrRmSz Or wRs_Txndz!TzRmSzTo = 0 Or wRs_InvDets!IrRmSz = 0) Then
        
          wd_AssignWt = IIF(Round(wRs_Txndz!qBalWt - wRs_Txndz!qAssignedWt - wd_BalWt, 3) > 0, wd_BalWt, wRs_Txndz!qBalWt - wRs_Txndz!qAssignedWt)
          ''*** (Jen 3.1.3) *** For Rouding Off Problem
          'wd_BalWt = wd_BalWt - wd_AssignWt
          'wd_Wt = wd_Wt - wd_AssignWt
          'If wd_Wt < 0 Then wd_AssignWt = wd_AssignWt + wd_Wt
          ''*** (Jen 3.1.3) *** For Rouding Off Problem
          wRs_Txndz.edit
          wRs_Txndz!qAssignedWt = wRs_Txndz!qAssignedWt + wd_AssignWt
          wd_BalWt = wd_BalWt - wd_AssignWt
          wRs_Txndz.Update
        End If
        wRs_Txndz.MoveNext
      Loop
      wRs_InvDets.MoveNext
    Loop
  End With
  
  Set wRs_Txndz = Nothing
  Set wRs_Txndz = moCn.OpenResultset("Select * from #TmpBestUse where convert (decimal (14, 3), qAssignedWt)> 0 " + _
                  "Order By TzTc, TzYy, TzChr, TzNo, TzSr")
  
  '*** Begin Copying GrdTxndZ Records
  GrdTxndZ.StartCopy
  MWLib.BeginProcess Me, "Copying Custom Match Records ..."
  With GrdTxndZ
    Do While Not (wRs_Txndz.EOF Or wRs_Txndz.BOF)
      .AddItem
      wf_Row = .Rows - 1
      .Value(wf_Row, "TzSrNo") = wf_Row
      .Value(wf_Row, "wTzBaseVchNo") = wRs_Txndz!qBaseVch
      .Value(wf_Row, "wTzBaseBoeNo") = wRs_Txndz!TBOENo
      .Value(wf_Row, "TzRmWt") = wRs_Txndz!qAssignedWt
      ' ***** Manali 3.6.0 - 31/12/09 - Notional Rate Instead od Purchase Rate For DTA
      'Sachin - 4.1.0.0
      If ms_NotionRt = "Y" And (ws_RmCtg = "G" Or ws_RmCtg = "P" Or ws_RmCtg = "S" Or ws_RmCtg = "L") Then
        Dim wd_CustmCdPurityZ As Double
         wd_CustmCdPurityZ = moCn.GetFldVal("Select RmPurityZ * " + _
                                IIF(ws_RmCtg = "G" And md_GldFiness > 0, CStr(md_GldFiness), IIF(ws_RmCtg = "P" And md_PlFiness > 0, CStr(md_PlFiness), _
                                IIF(ws_RmCtg = "S" And md_SlFiness > 0, CStr(md_SlFiness), IIF(ws_RmCtg = "L" And md_PdFiness > 0, CStr(md_PdFiness), 1)))) + _
                                " from RmMst Where RmCd='" + ws_ZRmCd + "'  " + _
                                IIF(adc("wIzPrtKey") = ctCurrPrtn, " and RmPrtKey= '" + ctCurrPrtn + "' ", ""))
    
        .Value(wf_Row, "TzPurRt") = MWLib.Div(IIF(ws_RmCtg = "G", md_LmgCst, IIF(ws_RmCtg = "P", md_LmpCst, IIF(ws_RmCtg = "S", md_LmsCst, IIF(ws_RmCtg = "L", md_LmlCst, 1)))), _
                                  moCn.GetFldVal("Select PNum from Param Where PTyp='GMFACTOR' and PMCd='" + ws_RmCtg + "' and PSCd='" + adc("wIzCmCurCd") + "'")) _
                                  * IIF(wd_CustmCdPurityZ > 0, wd_CustmCdPurityZ, 1)
        .Value(wf_Row, "wPurRtInRs") = .Value(wf_Row, "TzPurRt") * _
                                        MWLib.Div(Format(IIF(ms_RsMulDiv = "M", CStr(md_RsCnvFct), CStr(1)), "0.0000"), _
                                                  Format(IIF(ms_RsMulDiv = "D", CStr(md_RsCnvFct), CStr(1)), "0.0000"))
      ' ***** Manali 3.6.0 - 31/12/09 - Notional Rate Instead od Purchase Rate For DTA
      Else
        .Value(wf_Row, "TzPurRt") = wRs_Txndz!qPurRtInInvCur
        .Value(wf_Row, "wPurRtInRs") = wRs_Txndz!qPurRtInRs
      End If
      ' ***** Manali 3.10.0 - 23/05/12 - Change done for Uni, when amount upto 3 decimals, get amount as 0.01 to skip amt>0 check
      Dim wPurAmt As Single
      wPurAmt = .Value(wf_Row, "TzRmWt") * .Value(wf_Row, "TzPurRt")
      wPurAmt = IIF(Round(wPurAmt, 3) > 0 And Round(wPurAmt, 2) = 0, 0.01, wPurAmt)
      .Value(wf_Row, "TzPurAmt") = wPurAmt
      'bef 3.10.0
      '.Value(wf_Row, "TzPurAmt") = .Value(wf_Row, "TzRmWt") * .Value(wf_Row, "TzPurRt")   '   wRs_Txndz!qPurRtInInvCur
      .Value(wf_Row, "TzKey") = moCn.AutoGen("TxndZ", "TzKey", gs_CoCd + adc("wIzYy"))
      .Value(wf_Row, "TzRefYy") = wRs_Txndz!TzYy
      .Value(wf_Row, "TzRefKey") = wRs_Txndz!TzKey
      .Value(wf_Row, "TzPrtKey") = ctCurrPrtn       '****** Sachin 3.02 31-12-07 - Current Partition
      .SaveRec (wf_Row)
      wRs_Txndz.MoveNext
    Loop
  End With
  MWLib.EndProcess Me
  GrdTxndZ.EndCopy
  '*** End of Copying GrdTxndZ Records
  
  Set wRs_Txndz = Nothing
  moCn.Execute ("drop table #TmpBestUse")
  '*** (Jen 3.1.2)
End Sub
Private Sub AutoCustomMatchingNormal(ByVal pi_Row As Integer)
  '*** (Jen 3.1.2)
  Dim ws_str As String, wRs_Txndz As MwfLib.MDORowSet, wf_Row As Single
  Dim ws_RmCtg As String, ws_ZRmCd As String, wd_Rt As Double, wd_Wt As Double, wf_Qty As Single, wd_BalWt As Double

  wd_Wt = GrdIZ.Value(pi_Row, "wIzRmBalWt")
  If Round(wd_Wt, 3) <= 0 Then Exit Sub
  ws_ZRmCd = GrdIZ.Value(pi_Row, "IzRmZCd")
  ws_RmCtg = moCn.GetFldVal("Select RmCtg from RmMst where RmPrtKey='" + ctCurrPrtn + "' and RmCd= '" + ws_ZRmCd + "'")
  wd_Rt = GrdIZ.Value(pi_Row, "IzRmZRt")
  wf_Qty = GrdIZ.Value(pi_Row, "IzRmQty")
  wd_BalWt = wd_Wt

  '****** Sachin 3.02 31-12-07 - Current Partition
  '6.1
  ws_str = "set dateformat " + DtFmtStr() + " " + _
           "Select top 1 TDt, TzYy, TzKey, TzTc+ '/'+ TzYy+ '/'+ TzChr+ '/'+ LTrim(Str(TzNo))+ '/'+ " + _
           "       LTrim(Str(TzSr)) as qBaseVch, TBOENo, convert (decimal (16, 3), " + _
           "       (TzRmWt- TzRmDespWt)) as qBalWt, convert (decimal (14, 3), convert (decimal (14, 3), TzPurRt) " + _
           "          * (case when TMulDiv= 'D' then convert (decimal (14, 4), TCnvRt) else 1.0000 end) " + _
           "          / (case when TMulDiv= 'M' then convert (decimal (14, 4), TCnvRt) else 1.0000 end) " + _
           "          * " + Format(IIF(ms_MulDiv = "M", CStr(md_CnvFct), "1"), "0.0000") + " " + _
           "          / " + Format(IIF(ms_MulDiv = "D", CStr(md_CnvFct), "1"), "0.0000") + ") As qPurRtInInvCur, " + _
           "       convert (decimal (14, 3), convert (decimal (14, 3), TzPurRt) " + _
           "          * (case when TMulDiv= 'D' then convert (decimal (14, 4), TCnvRt) else 1.0000 end) " + _
           "          / (case when TMulDiv= 'M' then convert (decimal (14, 4), TCnvRt) else 1.0000 end) " + _
           "          * " + Format(IIF(ms_MulDiv = "M", md_CnvFct, 1), "0.0000") + " " + _
           "          / " + Format(IIF(ms_MulDiv = "D", md_CnvFct, 1), "0.0000") + " " + _
           "          * " + Format(IIF(ms_RsMulDiv = "M", md_RsCnvFct, 1), "0.0000") + " " + _
           "          / " + Format(IIF(ms_RsMulDiv = "D", md_RsCnvFct, 1), "0.0000") + ") As qPurRtInRs "

  '*** (Bef 3.1.3)
  'ws_str = ws_str + " from Txndz Tz " + _
  '         "join Txn T on TCoCd= TzCoCd and TTc= TzTc and TYy= TzYy and TChr= TzChr and TNo= TzNo " + _
  '         "where TzCls= 'N' and TzRmZCd= '" + ws_ZRmCd + "' and TzCoCd= '" + gs_CoCd + "' " + _
  '         IIF(ADC("wAutoUseOnlyCustGoods") = "Y", " and TCmCd= '" + ADC("wIzCmCd") + "' ", " and TCmCd in ('', '" + ADC("wIzCmCd") + "') ") + _
  '         "    " + GetAutoCustomMatchCnd(ws_RmCtg, "", wd_Rt, wd_Wt, wf_Qty) + " " + _
  '         "Order By TDt, TzYy, TzKey "
  '*** (Bef 3.1.3)
  
  'ws_str = ws_str + " from Txndz Tz " + _
  '         "join Txn T on TIdNo=TzTIdNo and TCoCd= TzCoCd and TTc= TzTc and TYy= TzYy and TChr= TzChr and TNo= TzNo " + _
  '         "and TPrtKey= TzPrtKey " + _
  '         "where TzCls= 'N' and TzRmZCd= '" + ws_ZRmCd + "' and TzCoCd= '" + gs_CoCd + "' and TzPrtKey='" + ctCurrPrtn + "' " + _
  '         IIF(ADC("wAutoUseOnlyCustGoods") = "Y", " and TCmCd= '" + ADC("wIzCmCd") + "' ", " and TCmCd in ('', '" + ADC("wIzCmCd") + "') ") + _
  '         "    " + GetAutoCustomMatchCnd(ws_RmCtg, "", wd_Rt, wd_Wt, wf_Qty) + " " + _
  '         "Order By TDt, TzYy, TzKey "

  '*** (Jen 3.1.3)
  ws_str = ws_str + " from Txndz Tz " + _
           "join Txn T on TIdNo=TzTIdNo and TCoCd= TzCoCd and TTc= TzTc and TYy= TzYy and TChr= TzChr and TNo= TzNo " + _
           "  and TPrtKey= TzPrtKey " + _
           "where TzCls= 'N' and TzRmZCd= '" + ws_ZRmCd + "' and TzCoCd= '" + gs_CoCd + "' and TzPrtKey='" + ctCurrPrtn + "' " + _
           IIF(adc("wAutoUseOnlyCustGoods") = "Y", " and TCmCd= '" + adc("wIzCmCd") + "' ", " and TCmCd in ('', '" + adc("wIzCmCd") + "') ") + _
           "    " + GetAutoCustomMatchCnd(ws_RmCtg, "", wd_Rt, GrdIZ.Value(pi_Row, "IzRmZWt"), wf_Qty) + " " + _
           "Order By TDt, TzYy, TzKey "
  '*** (Jen 3.1.3)


  '*** Begin Copying GrdTxndZ Records
  GrdTxndZ.StartCopy
  MWLib.BeginProcess Me, "Copying Custom Match Records ..."
  With GrdTxndZ
    Do While Round(wd_BalWt, 3) > 0
      Set wRs_Txndz = moCn.OpenRes(ws_str)
      If wRs_Txndz.RecCount > 0 Then
        wRs_Txndz.MoveFirst
        .AddItem
        wf_Row = .Rows - 1
        .Value(wf_Row, "TzSrNo") = wf_Row
        .Value(wf_Row, "wTzBaseVchNo") = wRs_Txndz!qBaseVch
        .Value(wf_Row, "wTzBaseBoeNo") = wRs_Txndz!TBOENo
        .Value(wf_Row, "TzRmWt") = IIF(Round(wRs_Txndz!qBalWt - wd_BalWt, 3) > 0, wd_BalWt, wRs_Txndz!qBalWt)
        wd_BalWt = wd_BalWt - .Value(wf_Row, "TzRmWt")
        ' ***** Manali 3.6.0 - 31/12/09 - Notional Rate Instead od Purchase Rate For DTA
        'Sachin - 4.1.0.0
        If ms_NotionRt = "Y" And (ws_RmCtg = "G" Or ws_RmCtg = "P" Or ws_RmCtg = "S" Or ws_RmCtg = "L") Then
          Dim wd_CustmCdPurityZ As Double
          wd_CustmCdPurityZ = moCn.GetFldVal("Select RmPurityZ * " + _
                                IIF(ws_RmCtg = "G" And md_GldFiness > 0, CStr(md_GldFiness), IIF(ws_RmCtg = "P" And md_PlFiness > 0, CStr(md_PlFiness), _
                                IIF(ws_RmCtg = "S" And md_SlFiness > 0, CStr(md_SlFiness), IIF(ws_RmCtg = "L" And md_PdFiness > 0, CStr(md_PdFiness), 1)))) + _
                                " from RmMst Where RmCd='" + ws_ZRmCd + "'  " + _
                                IIF(adc("wIzPrtKey") = ctCurrPrtn, " and RmPrtKey= '" + ctCurrPrtn + "' ", ""))
      
          .Value(wf_Row, "TzPurRt") = MWLib.Div(IIF(ws_RmCtg = "G", md_LmgCst, IIF(ws_RmCtg = "P", md_LmpCst, IIF(ws_RmCtg = "S", md_LmsCst, IIF(ws_RmCtg = "L", md_LmlCst, 1)))), _
                                    moCn.GetFldVal("Select PNum from Param Where PTyp='GMFACTOR' and PMCd='" + ws_RmCtg + "' and PSCd='" + adc("wIzCmCurCd") + "'")) _
                                    * IIF(wd_CustmCdPurityZ > 0, wd_CustmCdPurityZ, 1)
          .Value(wf_Row, "wPurRtInRs") = .Value(wf_Row, "TzPurRt") * _
                                          MWLib.Div(Format(IIF(ms_RsMulDiv = "M", CStr(md_RsCnvFct), CStr(1)), "0.0000"), _
                                                    Format(IIF(ms_RsMulDiv = "D", CStr(md_RsCnvFct), CStr(1)), "0.0000"))
        ' ***** Manali 3.6.0 - 31/12/09 - Notional Rate Instead od Purchase Rate For DTA
        Else
          .Value(wf_Row, "TzPurRt") = wRs_Txndz!qPurRtInInvCur
          .Value(wf_Row, "wPurRtInRs") = wRs_Txndz!qPurRtInRs
        End If
        ' ***** Manali 3.10.0 - 23/05/12 - Change done for Uni, when amount upto 3 decimals, get amount as 0.01 to skip amt>0 check
        Dim wPurAmt As Single
        wPurAmt = .Value(wf_Row, "TzRmWt") * .Value(wf_Row, "TzPurRt")
        wPurAmt = IIF(Round(wPurAmt, 3) > 0 And Round(wPurAmt, 2) = 0, 0.01, wPurAmt)
        .Value(wf_Row, "TzPurAmt") = wPurAmt
        'bef 3.10.0
        '.Value(wf_Row, "TzPurAmt") = .Value(wf_Row, "TzRmWt") * .Value(wf_Row, "TzPurRt")     'wRs_Txndz!qPurRtInInvCur
        
        .Value(wf_Row, "TzKey") = moCn.AutoGen("TxndZ", "TzKey", gs_CoCd + adc("wIzYy"))
        .Value(wf_Row, "TzRefYy") = wRs_Txndz!TzYy
        .Value(wf_Row, "TzRefKey") = wRs_Txndz!TzKey
        .Value(wf_Row, "TzPrtKey") = ctCurrPrtn       '****** Sachin 3.02 31-12-07 - Current Partition
        .SaveRec (wf_Row)
      Else
        Exit Do
      End If
    Loop
  End With
  MWLib.EndProcess Me
  GrdTxndZ.EndCopy
  '*** End of Copying GrdTxndZ Records

  Set wRs_Txndz = Nothing
  '*** (Jen 3.1.2)
End Sub

Private Sub AutoCustomMatching1()
  '*** (Bef 3.1.2)
  ''*** (Jen 3.01)
  ''*** This option is used to Copy records into the grid GrdIZ & GrdTxndZ
  '    '(the recs in the grid are unique for Custom code & Pure Rate combination)
  ''*** The copy is possible even after records exists in the grid GrdIZ. But it is possible
  '    'only for the custom code, Pure Rate combination which does not already exist in the grid.
  ''*** The records are copied one line per Custom Code, Custom Pure Rate
  ''*** The Weights are converted into pure Weight
  '
  ''Dim ws_SelCnd As String    '*** (Bef 3.1.1)
  'Dim wf_StartRow As Single, wf_EndRow As Single, wf_IZRow As Single
  'Dim ws_str As String, wRs_Txndz As MwfLib.MDORowSet, wf_Row As Single
  'Dim ws_RmCtg As String, ws_ZRmCd As String, wd_Rt As Double, wd_Wt As Double, wf_Qty As Single, wd_BalWt As Double
  '
  ''*** (Bef 3.1.1)
  ''If adc("wAutoCpyOpt") = "" Then DispMsg "RM Category List Cannot Be Blank", etError: Exit Sub
  ''ws_SelCnd = " and R1.RmCtg in ('" + Replace(moCn.GetFldVal("Select PDesc225 from Param where PTyp= 'CPYRT' and " + _
  ''            "PMCd= '" + adc("wAutoCpyOpt") + "' "), ",", "','") + "') "
  ''wf_StartRow = GetInvZRmRecs(ws_SelCnd)
  ''wf_EndRow = GrdIZ.Rows - 1
  ''If wf_StartRow = 0 Then Exit Sub
  ''*** (Bef 3.1.1)
  '
  ''*** (Jen 3.1.1)
  'wf_StartRow = 1
  'wf_EndRow = GrdIZ.Rows - 1
  'If wf_EndRow = 0 Then Exit Sub
  ''*** (Jen 3.1.1)
  '
  'For wf_IZRow = wf_StartRow To wf_EndRow
  '  GrdIZ.Row = wf_IZRow
  '  ws_ZRmCd = GrdIZ.Value(wf_IZRow, "IzRmZCd")
  '  ws_RmCtg = moCn.GetFldVal("Select RmCtg from RmMst where RmCd= '" + ws_ZRmCd + "'")
  '  wd_Rt = GrdIZ.Value(wf_IZRow, "IzRmZRt")
  '  'wd_Wt = GrdIZ.Value(wf_IZRow, "IzRmZWt")   '*** (Bef 3.1.1)
  '  wd_Wt = GrdIZ.Value(wf_IZRow, "wIzRmBalWt") '*** (Jen 3.1.1)
  '  wf_Qty = GrdIZ.Value(wf_IZRow, "IzRmQty")
  '  wd_BalWt = wd_Wt
  '
  '  '*** (changed Jen 3.1.1)
  '  '*** (Bef 05/11/2007)
  '  'ws_str = "set dateformat dmy " + _
  '           "Select top 1 TDt, TzYy, TzKey, TzTc+ '/'+ TzYy+ '/'+ TzChr+ '/'+ LTrim(Str(TzNo))+ '/'+ " + _
  '           "       LTrim(Str(TzSr)) as qBaseVch, TBOENo, convert (decimal (16, 3), " + _
  '           "       (TzRmWt- TzRmDespWt)) as qBalWt, convert (decimal (16, 3), TzPurRt " + _
  '           "          * (case when TMulDiv= 'D' then TCnvRt else 1.0000 end)/ (case when TMulDiv= 'M' then TCnvRt else 1.0000 end) " + _
  '           "          * " + Format(IIF(ms_MulDiv = "M", CStr(md_CnvFct), "1"), "0.0000") + " " + _
  '           "          / " + Format(IIF(ms_MulDiv = "D", CStr(md_CnvFct), "1"), "0.0000") + ") As qPurRtInInvCur, " + _
  '           "       convert (decimal (16, 3), TzPurRt " + _
  '           "          * (case when TMulDiv= 'D' then TCnvRt else 1.0000 end)/ (case when TMulDiv= 'M' then TCnvRt else 1.0000 end) " + _
  '           "          * " + Format(IIF(ms_MulDiv = "M", md_CnvFct, 1), "0.0000") + " " + _
  '           "          / " + Format(IIF(ms_MulDiv = "D", md_CnvFct, 1), "0.0000") + " " + _
  '           "          * " + Format(IIF(ms_RsMulDiv = "M", md_RsCnvFct, 1), "0.0000") + " " + _
  '           "          / " + Format(IIF(ms_RsMulDiv = "D", md_RsCnvFct, 1), "0.0000") + ") As qPurRtInRs "
  '  '*** (Bef 05/11/2007)
  '
  '  '*** (Jen 05/11/2007)
  '  ws_str = "set dateformat dmy " + _
  '           "Select top 1 TDt, TzYy, TzKey, TzTc+ '/'+ TzYy+ '/'+ TzChr+ '/'+ LTrim(Str(TzNo))+ '/'+ " + _
  '           "       LTrim(Str(TzSr)) as qBaseVch, TBOENo, convert (decimal (16, 3), " + _
  '           "       (TzRmWt- TzRmDespWt)) as qBalWt, convert (decimal (14, 3), convert (decimal (14, 3), TzPurRt) " + _
  '           "          * (case when TMulDiv= 'D' then convert (decimal (14, 4), TCnvRt) else 1.0000 end) " + _
  '           "          / (case when TMulDiv= 'M' then convert (decimal (14, 4), TCnvRt) else 1.0000 end) " + _
  '           "          * " + Format(IIF(ms_MulDiv = "M", CStr(md_CnvFct), "1"), "0.0000") + " " + _
  '           "          / " + Format(IIF(ms_MulDiv = "D", CStr(md_CnvFct), "1"), "0.0000") + ") As qPurRtInInvCur, " + _
  '           "       convert (decimal (14, 3), convert (decimal (14, 3), TzPurRt) " + _
  '           "          * (case when TMulDiv= 'D' then convert (decimal (14, 4), TCnvRt) else 1.0000 end) " + _
  '           "          / (case when TMulDiv= 'M' then convert (decimal (14, 4), TCnvRt) else 1.0000 end) " + _
  '           "          * " + Format(IIF(ms_MulDiv = "M", md_CnvFct, 1), "0.0000") + " " + _
  '           "          / " + Format(IIF(ms_MulDiv = "D", md_CnvFct, 1), "0.0000") + " " + _
  '           "          * " + Format(IIF(ms_RsMulDiv = "M", md_RsCnvFct, 1), "0.0000") + " " + _
  '           "          / " + Format(IIF(ms_RsMulDiv = "D", md_RsCnvFct, 1), "0.0000") + ") As qPurRtInRs "
  '  '*** (Jen 05/11/2007)
  '
  '  ws_str = ws_str + " from Txndz Tz " + _
  '           "join Txn T on TCoCd= TzCoCd and TTc= TzTc and TYy= TzYy and TChr= TzChr and TNo= TzNo " + _
  '           "where TzCls= 'N' and TzRmZCd= '" + ws_ZRmCd + "' and TzCoCd= '" + gs_CoCd + "' " + _
  '           IIF(ADC("wAutoUseOnlyCustGoods") = "Y", " and TCmCd= '" + ADC("wIzCmCd") + "' ", " and TCmCd in ('', '" + ADC("wIzCmCd") + "') ") + _
  '           "    " + GetAutoCustomMatchCnd(ws_RmCtg, "", wd_Rt, wd_Wt, wf_Qty) + " " + _
  '           "Order By TDt, TzYy, TzKey "
  '  '*** (changed Jen 3.1.1)
  '
  '  '*** Begin Copying GrdTxndZ Records
  '  GrdTxndZ.StartCopy
  '  MWLib.BeginProcess Me, "Copying Custom Match Records ..."
  '  With GrdTxndZ
  '    Do While Round(wd_BalWt, 3) > 0
  '      Set wRs_Txndz = moCn.OpenRes(ws_str)
  '      If wRs_Txndz.RecCount > 0 Then
  '        wRs_Txndz.MoveFirst
  '        .AddItem
  '        wf_Row = .Rows - 1
  '        .Value(wf_Row, "TzSrNo") = wf_Row
  '        .Value(wf_Row, "wTzBaseVchNo") = wRs_Txndz!qBaseVch
  '        .Value(wf_Row, "wTzBaseBoeNo") = wRs_Txndz!TBOENo
  '        .Value(wf_Row, "TzRmWt") = IIF(Round(wRs_Txndz!qBalWt - wd_BalWt, 3) > 0, wd_BalWt, wRs_Txndz!qBalWt)
  '        wd_BalWt = wd_BalWt - .Value(wf_Row, "TzRmWt")
  '        .Value(wf_Row, "TzPurRt") = wRs_Txndz!qPurRtInInvCur
  '        .Value(wf_Row, "wPurRtInRs") = wRs_Txndz!qPurRtInRs
  '        .Value(wf_Row, "TzPurAmt") = .Value(wf_Row, "TzRmWt") * wRs_Txndz!qPurRtInInvCur
  '        .Value(wf_Row, "TzKey") = moCn.AutoGen("TxndZ", "TzKey", gs_CoCd + ADC("wIzYy"))
  '        .Value(wf_Row, "TzRefYy") = wRs_Txndz!TzYy
  '        .Value(wf_Row, "TzRefKey") = wRs_Txndz!TzKey
  '        .SaveRec (wf_Row)
  '      Else
  '        Exit Do
  '      End If
  '    Loop
  '  End With
  '  MWLib.EndProcess Me
  '  GrdTxndZ.EndCopy
  '  '*** End of Copying GrdTxndZ Records
  '
  'Next wf_IZRow
  '
  'Set wRs_Txndz = Nothing
  ''*** (Jen 3.01)
  '*** (Bef 3.1.2)
End Sub
Private Sub AutoCustomMatching2()
  ''*** (Jen 3.01)
  ''*** This option is used to Copy records into the grid GrdIZ & GrdTxndZ
  '    '(the recs in the grid are unique for Custom code & Pure Rate combination)
  ''*** The copy is possible even after records exists in the grid GrdIZ. But it is possible
  '    'only for the custom code, Pure Rate combination which does not already exist in the grid.
  ''*** The records are copied one line per Custom Code, Custom Pure Rate
  ''*** The Weights are converted into pure Weight
  '
  'Dim ws_SelCnd As String, wf_StartRow As Single, wf_EndRow As Single, wf_IZRow As Single
  'Dim ws_str As String, wRs_Txndz As MwfLib.MDORowSet, wf_Row As Single
  'Dim ws_RmCtg As String, ws_ZRmCd As String, wd_Rt As Double, wd_Wt As Double, wf_Qty As Single, wd_BalWt As Double
  '
  'If ADC("wAutoCpyOpt") = "" Then DispMsg "RM Category List Cannot Be Blank", etError: Exit Sub
  'ws_SelCnd = " and R1.RmCtg in ('" + Replace(moCn.GetFldVal("Select PDesc225 from Param where PTyp= 'CPYRT' and " + _
  '            "PMCd= '" + ADC("wAutoCpyOpt") + "' "), ",", "','") + "') "
  'wf_StartRow = GetInvZRmRecs(ws_SelCnd)
  '
  'wf_EndRow = GrdIZ.Rows - 1
  '
  'If wf_StartRow = 0 Then Exit Sub
  '
  'For wf_IZRow = wf_StartRow To wf_EndRow
  '  GrdIZ.Row = wf_IZRow
  '  ws_ZRmCd = GrdIZ.Value(wf_IZRow, "IzRmZCd")
  '  ws_RmCtg = moCn.GetFldVal("Select RmCtg from RmMst where RmCd= '" + ws_ZRmCd + "'")
  '  wd_Rt = GrdIZ.Value(wf_IZRow, "IzRmZRt")
  '  wd_Wt = GrdIZ.Value(wf_IZRow, "IzRmZWt")
  '  wf_Qty = GrdIZ.Value(wf_IZRow, "IzRmQty")
  '  wd_BalWt = wd_Wt
  '
  '  ws_str = "set dateformat dmy " + _
  '           "Select top 1 TDt, TzYy, TzKey, TzTc+ '/'+ TzYy+ '/'+ TzChr+ '/'+ LTrim(Str(TzNo))+ '/'+ " + _
  '           "       LTrim(Str(TzSr)) as qBaseVch, TBOENo, convert (decimal (16, 3), " + _
  '           "       (TzRmWt- TzRmDespWt)) as qBalWt, convert (decimal (16, 3), TzPurRt " + _
  '           "          * (case when TMulDiv= 'D' then TCnvRt else 1.0000 end)/ (case when TMulDiv= 'M' then TCnvRt else 1.0000 end) " + _
  '           "          * " + Format(IIF(ms_MulDiv = "M", CStr(md_CnvFct), "1"), "0.0000") + " " + _
  '           "          / " + Format(IIF(ms_MulDiv = "D", CStr(md_CnvFct), "1"), "0.0000") + ") As qPurRtInInvCur, " + _
  '           "       convert (decimal (16, 3), TzPurRt " + _
  '           "          * (case when TMulDiv= 'D' then TCnvRt else 1.0000 end)/ (case when TMulDiv= 'M' then TCnvRt else 1.0000 end) " + _
  '           "          * " + Format(IIF(ms_MulDiv = "M", md_CnvFct, 1), "0.0000") + " " + _
  '           "          / " + Format(IIF(ms_MulDiv = "D", md_CnvFct, 1), "0.0000") + " " + _
  '           "          * " + Format(IIF(ms_RsMulDiv = "M", md_RsCnvFct, 1), "0.0000") + " " + _
  '           "          / " + Format(IIF(ms_RsMulDiv = "D", md_RsCnvFct, 1), "0.0000") + ") As qPurRtInRs "
  '
  '  ws_str = ws_str + "from Txndz Tz " + _
  '           "join Txn T on TCoCd= TzCoCd and TTc= TzTc and TYy= TzYy and TChr= TzChr and TNo= TzNo " + _
  '           "join Head H on HCoCd= TCoCd " + _
  '           "where TzCls= 'N' and TzRmZCd= '" + ws_ZRmCd + "' and TzCoCd= '" + gs_CoCd + "' " + _
  '           "    " + GetAutoCustomMatchCnd(ws_RmCtg, "", wd_Rt, wd_Wt, wf_Qty) + " " + _
  '           "Order By TDt, TzYy, TzKey "
  '
  '  '*** Begin Copying GrdTxndZ Records
  '  GrdTxndZ.StartCopy
  '  MWLib.BeginProcess Me, "Copying Custom Match Records ..."
  '  With GrdTxndZ
  '    Do While Round(wd_BalWt, 3) > 0
  '      Set wRs_Txndz = moCn.OpenRes(ws_str)
  '      If wRs_Txndz.RecCount > 0 Then
  '        wRs_Txndz.MoveFirst
  '        .AddItem
  '        wf_Row = .Rows - 1
  '        .Value(wf_Row, "TzSrNo") = wf_Row
  '        .Value(wf_Row, "wTzBaseVchNo") = wRs_Txndz!qBaseVch
  '        .Value(wf_Row, "wTzBaseBoeNo") = wRs_Txndz!TBOENo
  '        .Value(wf_Row, "TzRmWt") = IIF(Round(wRs_Txndz!qBalWt - wd_BalWt, 3) > 0, wd_BalWt, wRs_Txndz!qBalWt)
  '        wd_BalWt = wd_BalWt - .Value(wf_Row, "TzRmWt")
  '        .Value(wf_Row, "TzPurRt") = wRs_Txndz!qPurRtInInvCur
  '        .Value(wf_Row, "wPurRtInRs") = wRs_Txndz!qPurRtInRs
  '        .Value(wf_Row, "TzPurAmt") = .Value(wf_Row, "TzRmWt") * wRs_Txndz!qPurRtInInvCur
  '        .Value(wf_Row, "TzKey") = moCn.AutoGen("TxndZ", "TzKey", gs_CoCd + ADC("wIzYy"))
  '        .Value(wf_Row, "TzRefYy") = wRs_Txndz!TzYy
  '        .Value(wf_Row, "TzRefKey") = wRs_Txndz!TzKey
  '        .SaveRec (wf_Row)
  '      Else
  '        Exit Do
  '      End If
  '    Loop
  '  End With
  '  MWLib.EndProcess Me
  '  GrdTxndZ.EndCopy
  '  '*** End of Copying GrdTxndZ Records
  '
  'Next wf_IZRow
  '
  'Set wRs_Txndz = Nothing
  ''*** (Jen 3.01)
End Sub
Private Sub AutoCustomMatchingAllRecsTogether()
'*** (Bef 3.1.1)
'  '*** (Jen 3.01)
'  '*** This works & can be used instead of AutoCustomMatching
'  '*** Currently not being used
'  '*** In this procedure, all Child Records Of One GrdIZ Record are selected at a time
'
'  '*** This option is used to Copy records into the grid GrdIZ & GrdTxndZ
'      '(the recs in the grid are unique for Custom code & Pure Rate combination)
'  '*** The copy is possible even after records exists in the grid GrdIZ. But it is possible
'      'only for the custom code, Pure Rate combination which does not already exist in the grid.
'  '*** The records are copied one line per Custom Code, Custom Pure Rate
'  '*** The Weights are converted into pure Weight
'
'  Dim ws_SelCnd As String, wf_StartRow As Single, wf_EndRow As Single, wf_IZRow As Single
'  Dim ws_str As String, wRs_Txndz As MwfLib.MDORowSet, wf_Row As Single
'  Dim ws_RmCtg As String, ws_ZRmCd As String, wd_Rt As Double, wd_Wt As Double, wf_Qty As Single, wd_BalWt As Double
'
'  If adc("wAutoCpyOpt") = "" Then DispMsg "RM Category List Cannot Be Blank", etError: Exit Sub
'
'  ws_SelCnd = " and R1.RmCtg in ('" + Replace(moCn.GetFldVal("Select PDesc225 from Param where PTyp= 'CPYRT' and " + _
'              "PMCd= '" + adc("wAutoCpyOpt") + "' "), ",", "','") + "') "
'
'  wf_StartRow = GetInvZRmRecs(ws_SelCnd)
'  wf_EndRow = GrdIZ.Rows - 1
'
'  If wf_StartRow = 0 Then Exit Sub
'
'  For wf_IZRow = wf_StartRow To wf_EndRow
'    GrdIZ.Row = wf_IZRow
'    ws_ZRmCd = GrdIZ.Value(wf_IZRow, "IzRmZCd")
'    ws_RmCtg = moCn.GetFldVal("Select RmCtg from RmMst where RmCd= '" + ws_ZRmCd + "'")
'    wd_Rt = GrdIZ.Value(wf_IZRow, "IzRmZRt")
'    wd_Wt = GrdIZ.Value(wf_IZRow, "IzRmZWt")
'    wf_Qty = GrdIZ.Value(wf_IZRow, "IzRmQty")
'    wd_BalWt = wd_Wt
'
'    ws_str = "set dateformat dmy " + _
'             "Select * From " + _
'             "(  select TDt, TzYy, TzKey, TzTc+ '/'+ TzYy+ '/'+ TzChr+ '/'+ LTrim(Str(TzNo))+ '/'+ " + _
'             "       LTrim(Str(TzSr)) as qBaseVch, TBOENo, convert (decimal (16, 3), " + _
'             "       (TzRmWt- TzRmDespWt)) as qBalWt, convert (decimal (16, 3), TzPurRt " + _
'             "          * (case when TMulDiv= 'D' then TCnvRt else 1.0000 end)/ (case when TMulDiv= 'M' then TCnvRt else 1.0000 end) " + _
'             "          * " + Format(IIF(ms_MulDiv = "M", CStr(md_CnvFct), "1"), "0.0000") + " " + _
'             "          / " + Format(IIF(ms_MulDiv = "D", CStr(md_CnvFct), "1"), "0.0000") + ") As qPurRtInInvCur, " + _
'             "       convert (decimal (16, 3), TzPurRt " + _
'             "          * (case when TMulDiv= 'D' then TCnvRt else 1.0000 end)/ (case when TMulDiv= 'M' then TCnvRt else 1.0000 end) " + _
'             "          * " + Format(IIF(ms_MulDiv = "M", md_CnvFct, 1), "0.0000") + " " + _
'             "          / " + Format(IIF(ms_MulDiv = "D", md_CnvFct, 1), "0.0000") + " " + _
'             "          * " + Format(IIF(ms_RsMulDiv = "M", md_RsCnvFct, 1), "0.0000") + " " + _
'             "          / " + Format(IIF(ms_RsMulDiv = "D", md_RsCnvFct, 1), "0.0000") + ") As qPurRtInRs, " + _
'             "       IsNull((select sum(TzRmWt- TzRmDespWt) From Txndz Tz1 " + _
'             "               join Txn T1 on T1.TCoCd= Tz1.TzCoCd and T1.TTc= Tz1.TzTc and T1.TYy= Tz1.TzYy " + _
'             "               and T1.TChr= Tz1.TzChr and T1.TNo= Tz1.TzNo " + _
'             "               join Head H1 on H1.HCoCd= T1.TCoCd " + _
'             "               where Tz1.TzCls= 'N' and Tz1.TzRmZCd= '" + ws_ZRmCd + "' and Tz1.TzCoCd= '" + gs_CoCd + "' " + _
'             "               " + GetAutoCustomMatchCnd(ws_RmCtg, "1", wd_Rt, wd_Wt, wf_Qty) + " " + _
'             "               and ( T1.TDt< T.TDt or (T1.TDt= T.TDt and Tz1.TzYy< Tz.TzYy) " + _
'             "                     or (T1.TDt= T.TDt and Tz1.TzYy= Tz.TzYy and Tz1.TzKey< Tz.TzKey) " + _
'             "                )), 0) qWt "
'
'
'    'ws_Str = "set dateformat dmy " + _
'             "Select * From " + _
'             "(  select TDt, TzYy, TzKey, TzTc+ '/'+ TzYy+ '/'+ TzChr+ '/'+ LTrim(Str(TzNo))+ '/'+ " + _
'             "       LTrim(Str(TzSr)) as qBaseVch, TBOENo, convert (decimal (16, 3), " + _
'             "       (TzRmWt- TzRmDespWt)) as qBalWt, convert (decimal (16, 3), TzPurRt " + _
'             "          * (case when TMulDiv= 'D' then TCnvRt else 1 end)/ (case when TMulDiv= 'M' then TCnvRt else 1 end) " + _
'             "          * " + IIF(ms_MulDiv = "M", CStr(md_CnvFct), CStr(1)) + " " + _
'             "          / " + IIF(ms_MulDiv = "D", CStr(md_CnvFct), CStr(1)) + ") As qPurRtInInvCur, " + _
'             "       convert (decimal (16, 2), TzPurRt " + _
'             "          * (case when TMulDiv= 'D' then TCnvRt else 1 end)/ (case when TMulDiv= 'M' then TCnvRt else 1 end) " + _
'             "          * " + IIF(ms_MulDiv = "M", CStr(md_CnvFct), CStr(1)) + _
'             "          / " + IIF(ms_MulDiv = "D", CStr(md_CnvFct), CStr(1)) + _
'             "          * " + IIF(ms_RsMulDiv = "M", CStr(md_RsCnvFct), CStr(1)) + _
'             "          / " + IIF(ms_RsMulDiv = "D", CStr(md_RsCnvFct), CStr(1)) + ") As qPurRtInRs, " + _
'             "       IsNull((select sum(TzRmWt- TzRmDespWt) From Txndz Tz1 " + _
'             "               join Txn T1 on T1.TCoCd= Tz1.TzCoCd and T1.TTc= Tz1.TzTc and T1.TYy= Tz1.TzYy " + _
'             "               and T1.TChr= Tz1.TzChr and T1.TNo= Tz1.TzNo " + _
'             "               join Head H1 on H1.HCoCd= T1.TCoCd " + _
'             "               where Tz1.TzCls= 'N' and Tz1.TzRmZCd= '" + ws_ZRmCd + "' and Tz1.TzCoCd= '" + gs_CoCd + "' " + _
'             "               " + GetAutoCustomMatchCnd(ws_RmCtg, "1", wd_Rt, wd_Wt, wf_Qty) + " " + _
'             "               and ( T1.TDt< T.TDt or (T1.TDt= T.TDt and Tz1.TzYy< Tz.TzYy) " + _
'             "                     or (T1.TDt= T.TDt and Tz1.TzYy= Tz.TzYy and Tz1.TzKey< Tz.TzKey) " + _
'             "                )), 0) qWt "
'
'    ws_str = ws_str + "from Txndz Tz " + _
'             "join Txn T on TCoCd= TzCoCd and TTc= TzTc and TYy= TzYy and TChr= TzChr and TNo= TzNo " + _
'             "join Head H on HCoCd= TCoCd " + _
'             "where TzCls= 'N' and TzRmZCd= '" + ws_ZRmCd + "' and TzCoCd= '" + gs_CoCd + "' " + _
'             "    " + GetAutoCustomMatchCnd(ws_RmCtg, "", wd_Rt, wd_Wt, wf_Qty) + " " + _
'             ") PTbl " + _
'             "where convert (decimal (14, 3), qWt)<= convert (decimal (14, 3), " + CStr(wd_Wt) + ") " + _
'             "Order By TDt, TzYy, TzKey "
'
'    Set wRs_Txndz = moCn.OpenRes(ws_str)
'
'    '*** Begin Copying GrdTxndZ Records
'    GrdTxndZ.StartCopy
'    MWLib.BeginProcess Me, "Copying Custom Match Records ..."
'    With GrdTxndZ
'      Do While Not (wRs_Txndz.EOF Or wRs_Txndz.BOF)
'        .AddItem
'        wf_Row = .Rows - 1
'        .Value(wf_Row, "TzSrNo") = wf_Row  '.MaxVal("TzSrNo") + 1
'        .Value(wf_Row, "wTzBaseVchNo") = wRs_Txndz!qBaseVch
'        .Value(wf_Row, "wTzBaseBoeNo") = wRs_Txndz!TBOENo
'        .Value(wf_Row, "TzRmWt") = IIF(Round(wRs_Txndz!qBalWt - wd_BalWt, 3) > 0, wd_BalWt, wRs_Txndz!qBalWt)
'        wd_BalWt = wd_BalWt - .Value(wf_Row, "TzRmWt")
'        .Value(wf_Row, "TzPurRt") = wRs_Txndz!qPurRtInInvCur
'        .Value(wf_Row, "wPurRtInRs") = wRs_Txndz!qPurRtInRs
'        .Value(wf_Row, "TzPurAmt") = .Value(wf_Row, "TzRmWt") * wRs_Txndz!qPurRtInInvCur
'        .Value(wf_Row, "TzKey") = moCn.AutoGen("TxndZ", "TzKey", gs_CoCd + adc("wIzYy"))
'        .Value(wf_Row, "TzRefYy") = wRs_Txndz!TzYy
'        .Value(wf_Row, "TzRefKey") = wRs_Txndz!TzKey
'        .SaveRec (wf_Row)
'        wRs_Txndz.MoveNext
'      Loop
'    End With
'    MWLib.EndProcess Me
'    GrdTxndZ.EndCopy
'    '*** End of Copying GrdTxndZ Records
'
'  Next wf_IZRow
'
'  Set wRs_Txndz = Nothing
'  '*** (Jen 3.01)
'*** (Bef 3.1.1)
End Sub

Private Function GetAutoCustomMatchCnd(ByVal ps_RmCtg As String, ByVal ps_Alias As String, ByVal pd_Rt As Double, _
        ByVal pd_Wt As Double, ByVal pf_Qty As Single) As String
  
  '*** (Jen 3.1.1)
  Dim ws_Cnd As String, ws_Rt As String
  ps_RmCtg = Trim(UCase(ps_RmCtg))
  ps_Alias = Trim(ps_Alias)
  '*** (Bef 05/11/2007)
  'ws_Rt = " Tz" + ps_Alias + ".TzPurRt " + _
          "* (case when T" + ps_Alias + ".TMulDiv= 'D' then T" + ps_Alias + ".TCnvRt else 1.0000 end)/ (case when T" + ps_Alias + ".TMulDiv= 'M' then T" + ps_Alias + ".TCnvRt else 1.0000 end) " + _
          "* " + Format(IIF(ms_MulDiv = "M", md_CnvFct, 1), "0.0000") + " " + _
          "/ " + Format(IIF(ms_MulDiv = "D", md_CnvFct, 1), "0.0000") + " "
  '*** (Bef 05/11/2007)
  
  '*** (Jen 05/11/2007)
  ws_Rt = " convert (decimal (14, 3), Tz" + ps_Alias + ".TzPurRt) " + _
          "* (case when T" + ps_Alias + ".TMulDiv= 'D' then convert (decimal (14, 4), T" + ps_Alias + ".TCnvRt) else 1.0000 end)" + _
          "/ (case when T" + ps_Alias + ".TMulDiv= 'M' then convert (decimal (14, 4), T" + ps_Alias + ".TCnvRt) else 1.0000 end) " + _
          "* " + Format(IIF(ms_MulDiv = "M", md_CnvFct, 1), "0.0000") + " " + _
          "/ " + Format(IIF(ms_MulDiv = "D", md_CnvFct, 1), "0.0000") + " "
  '*** (Jen 05/11/2007)
  
  ' ***** Manali 3.03 - 23/06/08 - Pd Mod - "L" added
  If ps_RmCtg = "D" Or ps_RmCtg = "C" Or ps_RmCtg = "G" Or ps_RmCtg = "P" Or ps_RmCtg = "S" Or ps_RmCtg = "L" Or ps_RmCtg = "X" Then
    '*** Using Pc/ Ct
    'Dim wi_PcPerCt As Integer    '*** (Bef 3.1.3)
    Dim wi_PcPerCt As Single      '*** (Jen 3.1.3)
    
    wi_PcPerCt = Round(MWLib.Div(pf_Qty, pd_Wt), 0)
    
    '******* Manali 3.2.2 - Before date fields added in condition
    ws_Cnd = " and (convert (decimal (14, 2), " + CStr(adc("wAutoHRtMinus" + ps_RmCtg)) + ")= convert (decimal (14, 2), 0) or convert (decimal (14, 3), " + ws_Rt + ") >= convert (decimal (14, 3), " + Format(pd_Rt, "0.000") + "+ (" + Format(pd_Rt, "0.000") + "* " + CStr(adc("wAutoHRtMinus" + ps_RmCtg)) + "/ 100.00))) " + _
             " and (convert (decimal (14, 2), " + CStr(adc("wAutoHRtPlus" + ps_RmCtg)) + ")= convert (decimal (14, 2), 0) or convert (decimal (14, 3), " + ws_Rt + ") <= convert (decimal (14, 3), " + Format(pd_Rt, "0.000") + "+ (" + Format(pd_Rt, "0.000") + "* " + CStr(adc("wAutoHRtPlus" + ps_RmCtg)) + "/ 100.00))) " + _
             " and (convert (decimal (14, 2), " + CStr(adc("wAutoHPtrMinus" + ps_RmCtg)) + ")= convert (decimal (14, 2), 0) " + _
             "      or Tz" + ps_Alias + ".TzPcPerCt = 0 or Tz" + ps_Alias + ".TzPcPerCt >= " + _
             "      Round(" + Format(wi_PcPerCt, "0.000") + "+ " + _
             "      (" + Format(wi_PcPerCt, "0.000") + " * " + CStr(adc("wAutoHPtrMinus" + ps_RmCtg)) + "/ 100.00), 0)) " + _
             " and (convert (decimal (14, 2), " + CStr(adc("wAutoHPtrPlus" + ps_RmCtg)) + ")= convert (decimal (14, 2), 0) " + _
             "      or Tz" + ps_Alias + ".TzPcPerCt = 0 or Tz" + ps_Alias + ".TzPcPerCt <= " + _
             "      Round(" + Format(wi_PcPerCt, "0.000") + " + " + _
             "      (" + Format(wi_PcPerCt, "0.000") + " * " + CStr(adc("wAutoHPtrPlus" + ps_RmCtg)) + "/ 100.00), 0)) " + _
             " and ('" + adc("wAutoHImpTypNotInSel" + ps_RmCtg) + "'= '' or ',' + '" + adc("wAutoHImpTypNotInSel" + ps_RmCtg) + "'+ ','  not like '%,'+ T" + ps_Alias + ".TImpTyp +',%') " + _
             " and ('" + CStr(adc("wAutoHImpAft" + ps_RmCtg)) + "'= '01/01/80' Or T" + ps_Alias + ".TDt> '" + CStr(adc("wAutoHImpAft" + ps_RmCtg)) + "') " + _
             " and ('" + CStr(adc("wAutoHImpBef" + ps_RmCtg)) + "'= '01/01/80' Or T" + ps_Alias + ".TDt< '" + CStr(adc("wAutoHImpBef" + ps_RmCtg)) + "') "
    
  Else
    ws_Cnd = " "
  End If

  GetAutoCustomMatchCnd = ws_Cnd
  '*** (Jen 3.1.1)
End Function
Private Function GetAutoCustomMatchCnd1(ByVal ps_RmCtg As String, ByVal ps_Alias As String, ByVal pd_Rt As Double, _
        ByVal pd_Wt As Double, ByVal pf_Qty As Single) As String
  
  '*** (Bef 3.1.1)
  'Dim ws_Cnd As String, ws_Rt As String
  'ps_RmCtg = Trim(UCase(ps_RmCtg))
  'ps_Alias = Trim(ps_Alias)
  'ws_Rt = " Tz" + ps_Alias + ".TzPurRt " + _
  '        "* (case when T" + ps_Alias + ".TMulDiv= 'D' then T" + ps_Alias + ".TCnvRt else 1.0000 end)/ (case when T" + ps_Alias + ".TMulDiv= 'M' then T" + ps_Alias + ".TCnvRt else 1.0000 end) " + _
  '        "* " + Format(IIF(ms_MulDiv = "M", md_CnvFct, 1), "0.0000") + " " + _
  '        "/ " + Format(IIF(ms_MulDiv = "D", md_CnvFct, 1), "0.0000") + " "
  '
  'If ps_RmCtg = "D" Or ps_RmCtg = "C" Or ps_RmCtg = "G" Or ps_RmCtg = "P" Or ps_RmCtg = "S" Or ps_RmCtg = "X" Then
  '  '*** Using Pointer
  '  'ws_Cnd = " and (H" + ps_Alias + ".HRtMinus" + ps_RmCtg + "= 0 or " + ws_Rt + " >= " + CStr(pd_Rt) + "+ (" + CStr(pd_Rt) + "* H" + ps_Alias + ".HRtMinus" + ps_RmCtg + "/ 100)) " + _
  '           " and (H" + ps_Alias + ".HRtPlus" + ps_RmCtg + "= 0 or " + ws_Rt + " <= " + CStr(pd_Rt) + "+ (" + CStr(pd_Rt) + "* H" + ps_Alias + ".HRtPlus" + ps_RmCtg + "/ 100)) " + _
  '           " and (H" + ps_Alias + ".HPtrMinus" + ps_RmCtg + "= 0 or Tz" + ps_Alias + ".TzRmWt/ (case when Tz" + ps_Alias + ".TzRmQty> 0 then Tz" + ps_Alias + ".TzRmQty else 1 end)>= " + _
  '           "      " + CStr(pd_Wt) + "/ (case when " + CStr(pf_Qty) + "> 0 then " + CStr(pf_Qty) + " else 1 end)+ " + _
  '           "      (" + CStr(pd_Wt) + "/ (case when " + CStr(pf_Qty) + "> 0 then " + CStr(pf_Qty) + " else 1 end)* H" + ps_Alias + ".HPtrMinus" + ps_RmCtg + "/ 100)) " + _
  '           " and (H" + ps_Alias + ".HPtrPlus" + ps_RmCtg + "= 0 or Tz" + ps_Alias + ".TzRmWt/ (case when Tz" + ps_Alias + ".TzRmQty> 0 then Tz" + ps_Alias + ".TzRmQty else 1 end)<= " + _
  '           "      " + CStr(pd_Wt) + "/ (case when " + CStr(pf_Qty) + "> 0 then " + CStr(pf_Qty) + " else 1 end)+ " + _
  '           "      (" + CStr(pd_Wt) + "/ (case when " + CStr(pf_Qty) + "> 0 then " + CStr(pf_Qty) + " else 1 end)* H" + ps_Alias + ".HPtrPlus" + ps_RmCtg + "/ 100)) " + _
  '           " and (H" + ps_Alias + ".HImpTypNotInSel" + ps_RmCtg + "= '' or ',' + H" + ps_Alias + ".HImpTypNotInSel" + ps_RmCtg + "+ ','  not like '%,'+ T" + ps_Alias + ".TImpTyp +',%') " + _
  '           " and (H" + ps_Alias + ".HImpAft" + ps_RmCtg + "= '01/01/80' Or T" + ps_Alias + ".TDt> H" + ps_Alias + ".HImpAft" + ps_RmCtg + ") "
  '
  '  '*** Using Pc/ Ct
  '  Dim wi_PcPerCt As Integer
  '  wi_PcPerCt = Round(MWLib.Div(pf_Qty, pd_Wt), 0)
  '  ws_Cnd = " and (convert (decimal (14, 2), H" + ps_Alias + ".HRtMinus" + ps_RmCtg + ")= convert (decimal (14, 2), 0) or convert (decimal (14, 3), " + ws_Rt + ") >= convert (decimal (14, 3), " + Format(pd_Rt, "0.000") + "+ (" + Format(pd_Rt, "0.000") + "* H" + ps_Alias + ".HRtMinus" + ps_RmCtg + "/ 100.00))) " + _
  '           " and (convert (decimal (14, 2), H" + ps_Alias + ".HRtPlus" + ps_RmCtg + ")= convert (decimal (14, 2), 0) or convert (decimal (14, 3), " + ws_Rt + ") <= convert (decimal (14, 3), " + Format(pd_Rt, "0.000") + "+ (" + Format(pd_Rt, "0.000") + "* H" + ps_Alias + ".HRtPlus" + ps_RmCtg + "/ 100.00))) " + _
  '           " and (convert (decimal (14, 2), H" + ps_Alias + ".HPtrMinus" + ps_RmCtg + ")= convert (decimal (14, 2), 0) " + _
  '           "      or Tz" + ps_Alias + ".TzPcPerCt = 0 or Tz" + ps_Alias + ".TzPcPerCt >= " + _
  '           "      Round(" + Format(wi_PcPerCt, "0.000") + "+ " + _
  '           "      (" + Format(wi_PcPerCt, "0.000") + " * H" + ps_Alias + ".HPtrMinus" + ps_RmCtg + "/ 100.00), 0)) " + _
  '           " and (convert (decimal (14, 2), H" + ps_Alias + ".HPtrPlus" + ps_RmCtg + ")= convert (decimal (14, 2), 0) " + _
  '           "      or Tz" + ps_Alias + ".TzPcPerCt = 0 or Tz" + ps_Alias + ".TzPcPerCt <= " + _
  '           "      Round(" + Format(wi_PcPerCt, "0.000") + " + " + _
  '           "      (" + Format(wi_PcPerCt, "0.000") + " * H" + ps_Alias + ".HPtrPlus" + ps_RmCtg + "/ 100.00), 0)) " + _
  '           " and (H" + ps_Alias + ".HImpTypNotInSel" + ps_RmCtg + "= '' or ',' + H" + ps_Alias + ".HImpTypNotInSel" + ps_RmCtg + "+ ','  not like '%,'+ T" + ps_Alias + ".TImpTyp +',%') " + _
  '           " and (H" + ps_Alias + ".HImpAft" + ps_RmCtg + "= '01/01/80' Or T" + ps_Alias + ".TDt> H" + ps_Alias + ".HImpAft" + ps_RmCtg + ") "
  '
  'Else
  '  ws_Cnd = " "
  'End If
  '
  'GetAutoCustomMatchCnd = ws_Cnd
  '*** (Jen 3.1.1)
End Function


Private Function GetInvZRmRecs(ByVal ps_Cnd As String) As Single
  '*** (Jen 3.01)
  '*** This option is used to Copy records into the grid GrdIZ
      '(the recs in the grid are unique for Custom code & Pure Rate combination)
  '*** The copy is possible even after records exists in the grid GrdIZ. But it is possible
      'only for the custom code, Pure Rate combination which does not already exist in the grid.
  '*** The records are copied one line per Custom Code, Custom Pure Rate
  '*** The Weights are converted into pure Weight
  
  Dim wRsIZ As MwfLib.MDORowSet, wRow As Single, wf_StartRow As Single
  
  '****** Sachin 3.02 31-12-07 - Current Partition]
  ' ***** Manali 3.03 - 23/06/08 - Pd Mod - 'L' added
  If ms_LocalExp = "Y" Then
    Set wRsIZ = moCn.OpenRes("Select IrRmZCd, sum(IrRmQty) as qRmQty, " + _
                "sum((case when IrMainMet= 'Y' then IdCustmMtchWt else IrRmZWt end)* " + _
                "(case when R1.RmCtg in ('G', 'P', 'S', 'L') then R1.RmPurityZ " + _
                "                         * (1 + (case when IrMainMet= 'Y' then 0 else IdGldLs end)/100) " + _
                "                         / (case when R2.RmPurityZ> 0 then R2.RmPurityZ else 1 end) else 1 end) " + _
                "/ (case when (R1.RmCtg= 'G' and InHGldFiness> 0) then InHGldFiness " + _
                "        when (R1.RmCtg= 'P' and InHPlFiness> 0) then InHPlFiness " + _
                "        when (R1.RmCtg= 'S' and InHSlFiness> 0) then InHSlFiness " + _
                "        when (R1.RmCtg= 'L' and InHPdFiness> 0) then InHPdFiness else 1 end) ) as qRmZWt, " + _
                "IrRmZPureRt from InvRm " + _
                "join InvDsg on IdIdNo=IrIdIdNo and IdCoCd= IrCoCd and IdTc= IrTc and IdYy= IrYy and " + _
                "IdChr= IrChr and IdNo= IrNo and IdSr= IrSr and IdPrtKey=IrPrtKey " + _
                "join RmMst R1 on R1.RmCd= IrRmCd and R1.RmPrtKey=IrPrtKey " + _
                "join RmMst R2 on R2.RmCd= R1.RmZCd and R2.RmPrtKey=R1.RmPrtKey " + _
                "join InvHd on InCoCd= IrCoCd and InTc= IrTc and InYy= IrYy and " + _
                "InChr= IrChr and InNo= IrNo and InPrtKey=IrPrtKey " + _
                "where IrCoCd= '" + gs_CoCd + "' and IrTc= '" + adc("wIzTc") + "' and IrYy= '" + adc("wIzYy") + "' and " + _
                "IrChr= '" + adc("wIzChr") + "' and IrNo= " + CStr(adc("wIzNo")) + " and IrPrtKey='" + ctCurrPrtn + "' " + _
                "   and IrRmZWt> 0 and " + _
                "not exists(Select 'x' from InvZRm where " + _
                "IzCoCd= '" + gs_CoCd + "' and IzTc='" + adc("wIzTc") + "' " + _
                "   and IzYy= '" + adc("wIzYy") + "' and IzChr= '" + adc("wIzChr") + "' " + _
                "   and IzNo= " + CStr(adc("wIzNo")) + " and IzPrtKey='" + ctCurrPrtn + "' and IzRmZCd= IrRmZCd " + _
                "   and Round(IzRmZRt, 3)= Round(IrRmZPureRt, 3)) " + ps_Cnd + _
                " Group by IrRmZCd, IrRmZPureRt Order By IrRmZCd, IrRmZPureRt")
  Else
    '****** Sachin 3.02 31-12-07 - Current Partition
    ' ***** Manali 3.03 - 23/06/08 - Pd Mod - 'L' added
    Set wRsIZ = moCn.OpenRes("Select IrRmZCd, sum(IrRmQty) as qRmQty, " + _
                "sum((case when IrMainMet= 'Y' then IdCustmMtchWt else IrRmZWt end) " + _
                "   * (case when R1.RmCtg in ('G', 'P', 'S', 'L') then R1.RmPurityZ " + _
                "       * (1 + (case when IrMainMet= 'Y' then 0 else IdGldLs end)/100) " + _
                "       / (case when R2.RmPurityZ> 0 then R2.RmPurityZ else 1 end) else 1 end) " + _
                ") as qRmZWt, IrRmZPureRt from InvRm " + _
                "join InvDsg on IdIdNo=IrIdIdNo and IdCoCd= IrCoCd and IdTc= IrTc and IdYy= IrYy and " + _
                "IdChr= IrChr and IdNo= IrNo and IdSr= IrSr and IdPrtKey=IrPrtKey " + _
                "join RmMst R1 on R1.RmCd= IrRmCd and R1.RmPrtKey=IrPrtKey " + _
                "join RmMst R2 on R2.RmCd= R1.RmZCd and R2.RmPrtKey=R1.RmPrtKey " + _
                "where IrCoCd= '" + gs_CoCd + "' and " + _
                "IrTc= '" + adc("wIzTc") + "' and IrYy= '" + adc("wIzYy") + "' and " + _
                "IrChr= '" + adc("wIzChr") + "' and IrNo= " + CStr(adc("wIzNo")) + " and IrPrtKey='" + ctCurrPrtn + "' " + _
                "and IrRmZWt> 0 and " + _
                "not exists(Select 'x' from InvZRm where " + _
                "IzCoCd= '" + gs_CoCd + "' and IzTc='" + adc("wIzTc") + "' and " + _
                "IzYy= '" + adc("wIzYy") + "' and IzChr= '" + adc("wIzChr") + "' and " + _
                "IzNo= " + CStr(adc("wIzNo")) + " and IzPrtKey='" + ctCurrPrtn + "' and IzRmZCd= IrRmZCd " + _
                "and Round(IzRmZRt, 3)= Round(IrRmZPureRt, 3)) " + ps_Cnd + _
                " Group by IrRmZCd, IrRmZPureRt Order By IrRmZCd, IrRmZPureRt")
  End If

  '*** Begin Copying InvZRm Records
  GrdIZ.StartCopy
  MWLib.BeginProcess Me, "Copying Custom Codes ..."
  wf_StartRow = 0
  With GrdIZ
    Do While Not (wRsIZ.EOF Or wRsIZ.BOF)
      .AddItem
      wRow = .Rows - 1
      If wf_StartRow = 0 Then wf_StartRow = wRow
      .Value(wRow, "IzSr") = .MaxVal("IzSr") + 1
      .Value(wRow, "IzRmZCd") = wRsIZ!IrRmZCd
      .Value(wRow, "IzRmQty") = wRsIZ!qRmQty
      .Value(wRow, "IzRmZWt") = wRsIZ!qRmZWt
      .Value(wRow, "IzRmZRt") = wRsIZ!IrRmZPureRt
      .Value(wRow, "wIzRmBalWt") = .Value(wRow, "IzRmZWt")
      .Value(wRow, "IzPrtKey") = ctCurrPrtn         '****** Sachin 31-12-07
      .SaveRec (wRow)
      wRsIZ.MoveNext
    Loop
  End With
  MWLib.EndProcess Me
  GrdIZ.EndCopy
  Set wRsIZ = Nothing
  '*** End of Copying InvZRm Records
  
  GetInvZRmRecs = wf_StartRow
  '*** (Jen 3.01)
End Function
Private Sub RmDetails(ByVal mGrd As String)
  '*** For the current record of GrdIZ, the Rm Details Option shows all the the
      'Inv Design Sr records which have the same Custom Code. The records are shown InvSr
      'and Rm Code wise. The corresponding Rm Qty, Rm Wt, Export Order Number & Design
      'Code are also shown
  '*** In the GrdIZD grid only one line per RmCode & the corresponding Wt is shown for
      'the current row of GrdIZ. This makes it easier for the user to make the Match entries.

  If GrdIZ.Rows - 1 = 0 Then GrdIZR.Rows = 1: GrdIZD.Rows = 1: Exit Sub
  Dim wRsIZRm As MwfLib.MDORowSet, wRmTotWt As Single, wRmTotQty As Single, wStr As String
  Dim wRow As Single, wGrpBy As String, wIrSr As String
  
  wRmTotWt = 0: wRmTotQty = 0: wStr = ""
  If UCase(mGrd) = UCase("GrdIZR") Then
    wGrpBy = " group by IrSr, IrRmCd "
    wIrSr = " IrSr, "
  End If
  If UCase(mGrd) = UCase("GrdIZD") Then
    '*** (Bef 2.13) wGrpBy = " group by IrRmCd "
    
    '*** (Jen 2.13)
    '*** (07/11/06)
    If mb_ShowSize = True Then
      If mb_ShowSizeStkRt = False Then
        wGrpBy = " group by IrRmCd, IsNull(RrRngSz, '') "
      Else
        wGrpBy = " group by IrRmCd, IsNull(RrRngSz, IrRmSz) "
      End If
    Else
      wGrpBy = " group by IrRmCd "
    End If
    '*** (07/11/06)
    
    '*** (Bef 07/11/06)
    'If mb_ShowSizeStkRt = False Then
    '  wGrpBy = " group by IrRmCd, IsNull(RrRngSz, '') "
    'Else
    '  wGrpBy = " group by IrRmCd, IsNull(RrRngSz, IrRmSz) "
    'End If
    '*** (Bef 07/11/06)
    '*** (Jen 2.13)
    
    wIrSr = " 0 as IrSr, "
  End If
  
'  wStr = "Select " + wIrSr + " max(IdExpOdTc+'/'+IdExpOdYy+'/'+IdExpOdChr+'/'+LTrim(Str(IdExpOdNo))+" + _
'         "'/'+LTrim(Str(IdExpOdSr))) as qOrdSr, max(IdDmCd) as IdDmCd, IrRmCd, " + _
'         "sum(IrRmQty) as qRmQty, max(RmCtg) as RmCtg, " + _
'         "sum(IrRmZWt* (case when RmCtg in ('G', 'P', 'S') then RmPurityWt* (1 + IdGldLs/100) else 1 end)/ (case when PNum1> 0 then PNum1 else 1 end)) as qRmZWt " + _
'         "from InvRm join InvDsg on IdTc= IrTc and IdYy= IrYy and IdChr= IrChr and " + _
'         "IdNo= IrNo and IdSr= IrSr join RmMst on RmCd= IrRmCd join Param on PTyp= 'RMCTG' " + _
'         "and PMCd= RmCtg where IrRmZCd= '" + GrdIZ.Value(GrdIZ.Row, "IzRmZCd") + "' and " + _
'         "IrRmZRt= " + CStr(GrdIZ.Value(GrdIZ.Row, "IzRmZRt")) + " and " + _
'         "IrTc='" + ADC("wIzTc") + "' and IrYy= '" + ADC("wIzYy") + "' and " + _
'         "IrChr= '" + ADC("wIzChr") + "' and IrNo= " + CStr(ADC("wIzNo")) + wGrpBy
  
  
  
  '*** (Jen 2.13)
  Dim ws_RmZWt As String
  If ms_LocalExp = "Y" Then
  ' ***** Manali 3.03 - 23/06/08 - Pd Mod - 'L' added
    ws_RmZWt = " sum((case when IrMainMet= 'Y' then IdCustmMtchWt else IrRmZWt end) " + _
               "* (case when R1.RmCtg in ('G', 'P', 'S', 'L') then R1.RmPurityZ" + _
               "            * (1 + (case when IrMainMet= 'Y' then 0 else IdGldLs end)/100) " + _
               "            / (case when R2.RmPurityZ> 0 then R2.RmPurityZ else 1 end) else 1 end) " + _
               "/ (case when (R1.RmCtg= 'G' and InHGldFiness> 0) then InHGldFiness " + _
               "        when (R1.RmCtg= 'P' and InHPlFiness> 0) then InHPlFiness " + _
               "        when (R1.RmCtg= 'S' and InHSlFiness> 0) then InHSlFiness " + _
               "        when (R1.RmCtg= 'L' and InHPdFiness> 0) then InHPdFiness else 1 end) ) as qRmZWt "
  Else
  ' ***** Manali 3.03 - 23/06/08 - Pd Mod - 'L' added
    ws_RmZWt = "sum((case when IrMainMet= 'Y' then IdCustmMtchWt else IrRmZWt end) " + _
               "   * (case when R1.RmCtg in ('G', 'P', 'S', 'L') then R1.RmPurityZ " + _
               "       * (1 + (case when IrMainMet= 'Y' then 0 else IdGldLs end)/100) " + _
               "       / (case when R2.RmPurityZ> 0 then R2.RmPurityZ else 1 end) else 1 end) " + _
               ") as qRmZWt "
  End If
  '*** (Jen 2.13)
  
  '"sum(IrRmZWt* (case when RmCtg in ('G', 'P', 'S') then RmPurityZ* (1 + IdGldLs/100) else 1 end)) as qRmZWt "
    
  '****** (Jen 2.13) Added join for Range Size, InvHd, Custom Code Purity ******'
  '****** Sachin 3.02 31-12-07 - Current Partition
  wStr = "Select " + wIrSr + " max(IdExpOdTc+'/'+IdExpOdYy+'/'+IdExpOdChr+'/'+LTrim(Str(IdExpOdNo))+" + _
         "'/'+LTrim(Str(IdExpOdSr))) as qOrdSr, max(IdDmCd) as IdDmCd, IrRmCd, max(IsNull(RrRngSz, '')) as qRngSz, " + _
         "max(case when IsNull(RrRngSz, '###ZZZ###ZZZ')= '###ZZZ###ZZZ' then IrRmSz else 0 end) as qIrRmSz, " + _
         "sum(IrRmQty) as qRmQty, max(R1.RmCtg) as RmCtg, " + ws_RmZWt + _
         " from InvRm join InvDsg on IdIdNo=IrIdIdNo and IdCoCd= IrCoCd and IdTc= IrTc and IdYy= IrYy and " + _
         "IdChr= IrChr and IdNo= IrNo and IdSr= IrSr and IdPrtKey=IrPrtKey " + _
         "join RmMst R1 on R1.RmPrtKey=IrPrtKey and R1.RmCd= IrRmCd  " + _
         "join RmMst R2 on R2.RmCd= R1.RmZCd and R2.RmPrtKey=R1.RmPrtKey " + _
         "join InvHd on InCoCd= IrCoCd and InTc= IrTc and InYy= IrYy and InChr= IrChr and InNo= IrNo and InPrtKey=IrPrtKey " + _
         "Left Outer Join RmRt RSZ on RrTcTyp= 'RSZ' and RrCmCtg= 'C' and RrCmCd= '" + ctSelfCmCd + "' and " + _
         "RrCtg= R1.RmCtg and RrSCtg= R1.RmSCtg and RrSTWGrp= '' and RrCd= '' and RrDmCtg= '' and " + _
         "RrLsCtg= '' and RrLabMCd= '' and RrFrLn<= IrRmSz and RrToLn>= IrRmSz " + _
         "where IrRmZCd= '" + GrdIZ.Value(GrdIZ.Row, "IzRmZCd") + "' and " + _
         "LTrim(Str(IrRmZPureRt, 16, 3))= LTrim(Str(" + CStr(GrdIZ.Value(GrdIZ.Row, "IzRmZRt")) + ", 16, 3)) and " + _
         "IrCoCd= '" + gs_CoCd + "' and " + _
         "IrTc= '" + adc("wIzTc") + "' and IrYy= '" + adc("wIzYy") + "' and " + _
         "IrChr= '" + adc("wIzChr") + "' and IrNo= " + CStr(adc("wIzNo")) + " and IrPrtKey=(Case When '" + adc("WIZPRTKEY") + "'='" + ctCurrPrtn + "' Then '" + ctCurrPrtn + "' Else IrPrtKey End) " + wGrpBy
  
  Set wRsIZRm = moCn.OpenRes(wStr)

  If UCase(mGrd) = UCase("GrdIZR") Then
    GrdIZR.AllowAdd = True
    GrdIZR.AllowDelete = True
    GrdIZR.Rows = 1
    With GrdIZR
      Do While Not (wRsIZRm.EOF Or wRsIZRm.BOF)
        .AddItem
        wRow = .Rows - 1
        .Value(wRow, "wIzrIdSr") = wRsIZRm!IrSr
        .Value(wRow, "wIzrSrNo") = wRow
        .Value(wRow, "wIzrIdOrdNo") = wRsIZRm!qOrdSr
        .Value(wRow, "wIzrIdDmCd") = wRsIZRm!IdDmCd
        .Value(wRow, "wIzrIrRmCd") = wRsIZRm!IrRmCd
        .Value(wRow, "wIzrIrRmQty") = wRsIZRm!qRmQty
        .Value(wRow, "wIzrIrRmZWt") = wRsIZRm!qRmZWt
        
        wRmTotWt = wRmTotWt + wRsIZRm!qRmZWt
        wRmTotQty = wRmTotQty + wRsIZRm!qRmQty
        wRsIZRm.MoveNext
      Loop
    End With
    adc("wIZRRmTotWt") = wRmTotWt
    adc("wIZRRmTotQty") = wRmTotQty
    GrdIZR.AllowAdd = False
    GrdIZR.AllowDelete = False
  End If
  
  If UCase(mGrd) = UCase("GrdIZD") Then
    GrdIZD.AllowAdd = True
    GrdIZD.AllowDelete = True
    GrdIZD.Rows = 1
    With GrdIZD
      Do While Not (wRsIZRm.EOF Or wRsIZRm.BOF)
        .AddItem
        wRow = .Rows - 1
        .Value(wRow, "wIzdIrRmCd") = wRsIZRm!IrRmCd
        '*** (Jen 2.13)
        .Value(wRow, "wIzdRmRSz") = wRsIZRm!qRngSz
        .Value(wRow, "wIzdIrRmSz") = IIF(wRsIZRm!qIrRmSz <> 0 And wRsIZRm!qRngSz = "", Format(wRsIZRm!qIrRmSz, "#0.0000"), "")
        '*** (Jen 2.13)
        .Value(wRow, "wIzdIrRmZWt") = wRsIZRm!qRmZWt
        wRsIZRm.MoveNext
      Loop
    End With
    GrdIZD.AllowAdd = False
    GrdIZD.AllowDelete = False
  End If
  
  Set wRsIZRm = Nothing
  
End Sub
Private Sub DispFra(ByVal pv_InvZFra As en_InvZFra2)
  '*** This is the common procedure which is used to toggle between hiding and displaying
      'the frames (i.e. if the frame is visible then make it invisible and viseversa)
      'associated with a command button
  
  Select Case pv_InvZFra
  Case Is = InvRmDet
    If FraIZR.Visible = True Then
      FraIZR.Visible = False
      FraIZR.Enabled = False
      '*** required to pass the array position InvRmDet-0, to resotre the 'Causes Validation' State
      Call EnaDisaCmds(False, CmdIZR, InvRmDet)
      CmdIZR.SetFocus
    Else
      FraIZR.Visible = True
      FraIZR.Enabled = True
      FraIZR.ZOrder
      Call EnaDisaCmds(True, CmdIZR)
      CmdIZR.SetFocus
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
      adc("wCpRmCtgFr").SetFocus
      FraCP.ZOrder
      Call EnaDisaCmds(True, CmdCP)
    End If
   '************Geeta********Emr207*****
  Case Is = HlpSrt
    If FraHlpSrt.Visible = True Then
      FraHlpSrt.Visible = False
      FraHlpSrt.Enabled = False
      Call EnaDisaCmds(False, CmdHlpSrt, HlpSrt)
      CmdHlpSrt.SetFocus
    Else
      FraHlpSrt.Visible = True
      FraHlpSrt.Enabled = True
      adc("wSortBy").SetFocus
      FraHlpSrt.ZOrder
      Call EnaDisaCmds(True, CmdHlpSrt)
    End If
   '************Geeta********Emr207*****
   
   '*** (Jen 3.01)
    Case Is = AutoCustmMtch
      If FraAuto.Visible = True Then
        FraAuto.Visible = False
        FraAuto.Enabled = False
        Call EnaDisaCmds(False, CmdAuto, AutoCustmMtch)
        CmdAuto.SetFocus
      Else
        FraAuto.Visible = True
        FraAuto.Enabled = True
        'ADC("wAutoCpyOpt").SetFocus            '*** (Bef 3.1.1)
        'ADC("wAutoUseOnlyCustGoods").SetFocus  '*** (Jen 3.1.1)   '*** (Bef 3.2.0)
        adc("wAutoBestUseOnly").SetFocus        '*** (Jen 3.2.0)
        FraAuto.ZOrder
        Call EnaDisaCmds(True, CmdAuto)
      End If
   '*** (Jen 3.01)
   
   ' ****** Manali 3.5.0 - 10/11/08 - Delete Rm Details on Button Click
    Case Is = DelRm
      If FraDelRm.Visible = True Then
        FraDelRm.Visible = False
        FraDelRm.Enabled = False
        Call EnaDisaCmds(False, CmdDelRm, DelRm)
        CmdDelRm.SetFocus
      Else
        FraDelRm.Visible = True
        FraDelRm.Enabled = True
        adc("wDelRmCtg").SetFocus
        FraDelRm.ZOrder
        Call EnaDisaCmds(True, CmdDelRm)
      End If
   ' ****** Manali 3.5.0 - 10/11/08 - Delete Rm Details on Button Click
   
  End Select
End Sub
Private Sub EnaDisaCmds(ByVal pv_ShowFra As Boolean, Optional ByVal pv_CmdBut As MWCTL_BTN1, Optional ByVal pv_InvZFra As en_InvZFra2)
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
      If Not IsMissing(pv_InvZFra) Then
        pv_CmdBut.CausesValidation = mArr_Cv(pv_InvZFra)
      End If
    End If
  End If
  FraNKeyAll.Enabled = Not pv_ShowFra
  adc.AllowSave = Not pv_ShowFra
  'ADC.AllowDelete = Not pv_ShowFra
  
  If pv_ShowFra = True Then
    '** All command buttons are disabled
    CmdIZR.Enabled = Not pv_ShowFra
    CmdCP.Enabled = Not pv_ShowFra
    CmdHlpSrt.Enabled = Not pv_ShowFra
    CmdAuto.Enabled = Not pv_ShowFra    '*** (Jen 3.01)
    CmdDelRm.Enabled = Not pv_ShowFra    ' ****** Manali 3.5.0 - 10/11/08 - Delete Rm Details on Button Click
  Else
    '*** Restoring Ena/Disa state of all Command Buttons
    CmdIZR.Enabled = mArr_Ed(InvRmDet)
    CmdCP.Enabled = mArr_Ed(Cpy)
    CmdHlpSrt.Enabled = mArr_Ed(HlpSrt)
    CmdAuto.Enabled = mArr_Ed(AutoCustmMtch)    '*** (Jen 3.01)
    CmdDelRm.Enabled = mArr_Ed(DelRm)    ' ****** Manali 3.5.0 - 10/11/08 - Delete Rm Details on Button Click
  End If
  If Not pv_CmdBut Is Nothing Then pv_CmdBut.Enabled = True
End Sub
Private Sub HideAllFras()
  '*** This procedure hides all the frames other than FraNKeyAll
  FraNKeyAll.Visible = True
  FraIZR.Visible = False
  FraIZR.Enabled = False
  FraCP.Visible = False
  FraCP.Enabled = False
  '********Geeta*********Emr207****
  FraHlpSrt.Visible = False
  FraHlpSrt.Enabled = False
  
  '*** (Jen 3.01)
  FraAuto.Visible = False
  FraAuto.Enabled = False
  '*** (Jen 3.01)
End Sub
Private Sub StoreState()
  '*** This procedure is used to store the CausesValidation State and the Enabled/Disabled
      'states of all the Command Buttons
  '*** This is used for restoring the original states in the future
  
  mArr_Cv(InvRmDet) = CmdIZR.CausesValidation
  mArr_Cv(Cpy) = CmdCP.CausesValidation
  mArr_Cv(HlpSrt) = CmdHlpSrt.CausesValidation
  mArr_Cv(AutoCustmMtch) = CmdAuto.CausesValidation   '*** (Jen 3.01)
  mArr_Cv(DelRm) = CmdDelRm.CausesValidation     ' ****** Manali 3.5.0 - 10/11/08 - Delete Rm Details on Button Click
  
  mArr_Ed(InvRmDet) = CmdIZR.Enabled
  mArr_Ed(Cpy) = CmdCP.Enabled
  mArr_Ed(HlpSrt) = CmdHlpSrt.Enabled
  mArr_Ed(AutoCustmMtch) = CmdAuto.Enabled    '*** (Jen 3.01)
  mArr_Ed(DelRm) = CmdDelRm.Enabled            ' ****** Manali 3.5.0 - 10/11/08 - Delete Rm Details on Button Click
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
  ws_BtnStr = "CmdCpGo,CmdAutoGo"
  Call SetObjColors(Me, ws_BtnStr)
  '*** Jenny Colour
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
    
    Set moCn = Nothing    '*** (09/08/05)    '*** (Bef speed) bef 24/11/06
End Sub
Private Sub GrdIZ_Validate(Cancel As Boolean)
  Cancel = GrdIZ.Validate
End Sub
Private Sub GrdTxndZ_Validate(Cancel As Boolean)
  Cancel = GrdTxndZ.Validate
End Sub

Private Sub DeleteRm()
  Dim ws_CtgStr As String, wDelStrTz As String
  
  If adc("wDelRmCtg") = "" Then DispMsg "Please Enter RmCtg", etError: Exit Sub
  
  ws_CtgStr = moCn.GetFldVal("Select PDesc225 from Param where PTyp='CPYRT' and PMCd='" + adc("wDelRmCtg") + "' and PSCd=''")
  
    Dim i As Integer, wRows As Integer
    i = 1
    wRows = GrdIZ.Rows
    Do While i <= wRows - 1
      If InStr(1, "," + ws_CtgStr + ",", "," + moCn.GetFldVal("Select RmCtg From RmMst Where RmCd='" + GrdIZ.Value(i, "IzRmZCd") + "'") + ",") >= 1 Then
          wDelStrTz = " Delete from TxndZ " + _
                      " Where TzCoCd='" + gs_CoCd + "' and TzTc='" + adc("WIZTC") + "' and TzYy='" + adc("WIZYY") + "' and TzChr='" + adc("WIZCHR") + "' and TzNo=" + CStr(adc("WIZNO")) + " and TzSr=" + CStr(GrdIZ.Value(i, "IzSr"))
          moCn.Execute wDelStrTz
        GrdIZ.Del_Rec True, i
        wRows = GrdIZ.Rows
        i = 0
      End If
      i = i + 1
    Loop
    Call DispFra(DelRm)
End Sub


'InvZRm

'IzTc             IzYy             IzChr             IzNo
'IzSr             IzRmZCd          IzRmQty           IzRmZWt
'IzRmZRt



'TxndZ
'
'TzTc          TzYy          TzChr          TzNo
'TzSr          TzSrNo        TzRmZCd        TzRmCdDesc
'TzPcPerCt     TzRmQty       TzRmWt         TzKey
'TzRefYy       TzRefKey      TzPurRt        TzPurAmt
'TzBm          TzRmDespQty   TzRmDespWt     TzCls



'wIzrIdSr
'wIzrSrNo
'wIzrIdOrdNo
'wIzrIdDmCd
'wIzrIrRmCd
'wIzrIrRmQty
'wIzrIrRmZWt



