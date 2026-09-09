VERSION 5.00
Object = "{AB2CC32A-E652-42A2-B341-26ECE02F6386}#2.1#0"; "MwfCtlForSpeed610.ocx"
Begin VB.Form EmrFrmShpNote 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Shipment Note"
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
   Begin MwfCtl.MWCTL_BTN1 CmdCpyDsg 
      Height          =   465
      Left            =   6870
      TabIndex        =   24
      ToolTipText     =   "Option To Copy Design"
      Top             =   9300
      Width           =   1905
      _ExtentX        =   3360
      _ExtentY        =   820
      BackColor       =   -2147483633
      MaskColor       =   8454143
      Caption         =   "&Copy Order Design"
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
      TabIndex        =   15
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
      TabIndex        =   16
      Top             =   9360
      Width           =   15135
      _ExtentX        =   26696
      _ExtentY        =   979
      BtnWidth        =   885
      BtnHeight       =   435
      BtnMaskColor    =   8454143
      FontSz          =   8.25
   End
   Begin VB.Frame FRA_KEYS 
      BorderStyle     =   0  'None
      Height          =   330
      Left            =   0
      TabIndex        =   18
      Top             =   50
      Width           =   4380
      Begin MwfCtl.MWCTL_MED ATXT 
         Height          =   285
         Index           =   1
         Left            =   1755
         TabIndex        =   1
         ToolTipText     =   "Enter Shipment Note Tc"
         Top             =   0
         Width           =   570
         _ExtentX        =   1005
         _ExtentY        =   503
         MaxLength       =   3
         DataType        =   4
         DataField       =   "SnTc"
         IdName          =   "SNTC"
      End
      Begin MwfCtl.MWCTL_MED ATXT 
         Height          =   285
         Index           =   40
         Left            =   2430
         TabIndex        =   2
         ToolTipText     =   "Enter Shipment Note Year"
         Top             =   0
         Width           =   330
         _ExtentX        =   582
         _ExtentY        =   503
         MaxLength       =   2
         DataType        =   4
         DataField       =   "SnYy"
         IdName          =   "SNYY"
      End
      Begin MwfCtl.MWCTL_MED ATXT 
         Height          =   285
         Index           =   41
         Left            =   2860
         TabIndex        =   3
         ToolTipText     =   "Enter Shipment Note Character"
         Top             =   0
         Width           =   570
         _ExtentX        =   1005
         _ExtentY        =   503
         MaxLength       =   3
         DataType        =   4
         DataField       =   "SnChr"
         IdName          =   "SNCHR"
      End
      Begin MwfCtl.MWCTL_MED ATXT 
         Height          =   285
         Index           =   42
         Left            =   3540
         TabIndex        =   4
         ToolTipText     =   "Enter Shipment Note No."
         Top             =   0
         Width           =   750
         _ExtentX        =   1323
         _ExtentY        =   503
         Alignment       =   2
         Mask            =   "#####0"
         MaxLength       =   6
         DataType        =   1
         DataField       =   "SnNo"
         IdName          =   "SNNO"
      End
      Begin MwfCtl.MWCTL_MED ATXT 
         Height          =   285
         Index           =   22
         Left            =   1080
         TabIndex        =   0
         ToolTipText     =   "Enter Company Code"
         Top             =   0
         Width           =   570
         _ExtentX        =   1005
         _ExtentY        =   503
         MaxLength       =   3
         DataType        =   4
         DataField       =   "SnCoCd"
         IdName          =   "SNCOCD"
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
         Left            =   1680
         TabIndex        =   32
         Top             =   0
         Width           =   105
      End
      Begin VB.Label ALBL 
         BackStyle       =   0  'Transparent
         Caption         =   "Shipment"
         BeginProperty Font 
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
         Width           =   975
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
         Left            =   2340
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
         Left            =   2775
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
         Left            =   3465
         TabIndex        =   20
         Top             =   0
         Width           =   105
      End
   End
   Begin VB.Frame FRA_NKEYS 
      BorderStyle     =   0  'None
      Enabled         =   0   'False
      Height          =   9810
      Left            =   -120
      TabIndex        =   17
      Top             =   0
      Width           =   15315
      Begin VB.Frame FraCpyDsg 
         Height          =   1245
         Left            =   5790
         TabIndex        =   25
         Top             =   7965
         Visible         =   0   'False
         Width           =   3465
         Begin MwfCtl.MWCTL_BTN1 CmdCpyDsgGo 
            Height          =   495
            Left            =   2580
            TabIndex        =   30
            ToolTipText     =   "Click Here To Copy Designs "
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
            TabIndex        =   28
            ToolTipText     =   "Enter From Order Sr"
            Top             =   420
            Width           =   795
            _ExtentX        =   1402
            _ExtentY        =   503
            Alignment       =   2
            Mask            =   "#####0"
            MaxLength       =   6
            DataType        =   1
            IdName          =   "WODSRFR"
         End
         Begin MwfCtl.MWCTL_MED ATXT 
            Height          =   285
            Index           =   8
            Left            =   1260
            TabIndex        =   29
            ToolTipText     =   "Enter To Order Sr"
            Top             =   705
            Width           =   795
            _ExtentX        =   1402
            _ExtentY        =   503
            Alignment       =   2
            Mask            =   "#####0"
            MaxLength       =   6
            DataType        =   1
            IdName          =   "WODSRTO"
         End
         Begin VB.Label LblMain 
            BackStyle       =   0  'Transparent
            BorderStyle     =   1  'Fixed Single
            Caption         =   "  Copy Order Design"
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
            TabIndex        =   31
            Top             =   90
            Width           =   3450
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
            TabIndex        =   27
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
            TabIndex        =   26
            Top             =   420
            Width           =   1215
         End
      End
      Begin VB.Frame FraNKeyAll 
         BorderStyle     =   0  'None
         Height          =   9645
         Left            =   150
         TabIndex        =   19
         Top             =   -180
         Width           =   15150
         Begin MwfCtl.MWCTL_MED ATXT 
            Height          =   285
            Index           =   4
            Left            =   3120
            TabIndex        =   50
            ToolTipText     =   "Enter Order Pcs"
            Top             =   9120
            Visible         =   0   'False
            Width           =   870
            _ExtentX        =   1535
            _ExtentY        =   503
            Alignment       =   2
            Mask            =   "####0.0"
            MaxLength       =   7
            DataType        =   2
            ReCalcParent    =   "GrdShpOrdMst(WSOMORDPCS),GrdShpOrdMst(WSOMSHPBAL)"
            ReCalcOn        =   "GrdShpOrdDsg(WSODSHPBAL)"
            IdName          =   "WODSHPBAL"
         End
         Begin MwfCtl.MWCTL_MED ATXT 
            Height          =   285
            Index           =   3
            Left            =   1680
            TabIndex        =   49
            ToolTipText     =   "Enter Order Pcs"
            Top             =   9120
            Visible         =   0   'False
            Width           =   870
            _ExtentX        =   1535
            _ExtentY        =   503
            Alignment       =   2
            Mask            =   "####0.0"
            MaxLength       =   7
            DataType        =   2
            ReCalcParent    =   "GrdShpOrdMst(WSOMORDPCS)"
            ReCalcOn        =   "GrdShpOrdDsg(WSODORDPCS)"
            IdName          =   "WODORDPCS"
         End
         Begin MwfCtl.MWCTL_FLX GrdShpOrdMst 
            Height          =   4125
            Left            =   60
            TabIndex        =   14
            Top             =   930
            Width           =   9840
            _ExtentX        =   17357
            _ExtentY        =   7276
            Cols            =   12
            colname1        =   "SOMSR"
            heading1        =   "Sr     "
            datafld1        =   "SomSr"
            datatype1       =   1
            mask1           =   "#####0"
            maxlength1      =   6
            style1          =   2
            colname2        =   "SOMOMCOCD"
            heading2        =   "CoCd"
            datafld2        =   "SomOmCoCd"
            datatype2       =   4
            recalcparent2   =   "GrdShpOrdMst(WSOMCMCD),GrdShpOrdMst(WSOMEXPDELDT),GrdShpOrdMst(WSOMSHOWINV)"
            maxlength2      =   3
            style2          =   0
            tooltiptext2    =   "Enter Order Company Code"
            colname3        =   "SOMOMTC"
            heading3        =   "Tc    "
            datafld3        =   "SomOmTc"
            datatype3       =   4
            recalcparent3   =   "GrdShpOrdMst(WSOMCMCD),GrdShpOrdMst(WSOMEXPDELDT),GrdShpOrdMst(WSOMSHOWINV)"
            maxlength3      =   3
            tooltiptext3    =   "Enter Order Tc"
            colname4        =   "SOMOMYY"
            heading4        =   "Yy  "
            datafld4        =   "SomOmYy"
            datatype4       =   4
            recalcparent4   =   "GrdShpOrdMst(WSOMCMCD),GrdShpOrdMst(WSOMEXPDELDT),GrdShpOrdMst(WSOMSHOWINV)"
            maxlength4      =   2
            tooltiptext4    =   "Enter Order Year"
            colname5        =   "SOMOMCHR"
            heading5        =   "Chr     "
            datafld5        =   "SomOmChr"
            datatype5       =   4
            recalcparent5   =   "GrdShpOrdMst(WSOMCMCD),GrdShpOrdMst(WSOMEXPDELDT),GrdShpOrdMst(WSOMSHOWINV),GrdShpOrdMst(WSOMORDPCS)"
            maxlength5      =   3
            tooltiptext5    =   "Enter Order Character"
            colname6        =   "SOMOMNO"
            heading6        =   "No          "
            datafld6        =   "SomOmNo"
            datatype6       =   1
            mask6           =   "#####0"
            recalcparent6   =   "GrdShpOrdMst(WSOMCMCD),GrdShpOrdMst(WSOMEXPDELDT),GrdShpOrdMst(WSOMSHOWINV)"
            maxlength6      =   6
            tooltiptext6    =   "Enter Order Number (Press Shift + Tab To Come Out Of This Field)"
            colname7        =   "WSOMCMCD"
            heading7        =   "Customer            "
            datatype7       =   4
            recalcon7       =   "GrdShpOrdMst(SOMOMCOCD),GrdShpOrdMst(SOMOMTC),GrdShpOrdMst(SOMOMYY),GrdShpOrdMst(SOMOMCHR),GrdShpOrdMst(SOMOMNO)"
            maxlength7      =   8
            colname8        =   "WSOMEXPDELDT"
            heading8        =   "Exp Del Dt  "
            datatype8       =   3
            recalcon8       =   "GrdShpOrdMst(SOMOMCOCD),GrdShpOrdMst(SOMOMTC),GrdShpOrdMst(SOMOMYY),GrdShpOrdMst(SOMOMCHR),GrdShpOrdMst(SOMOMNO)"
            maxlength8      =   8
            colname9        =   "WSOMORDPCS"
            heading9        =   "Ord Pcs"
            datatype9       =   2
            mask9           =   "####0.0"
            recalcon9       =   "WODORDPCS,GrdShpOrdDsg(WSODORDPCS),GrdShpOrdMst(SOMOMCHR)"
            recalcparent9   =   "WORDPCS"
            maxlength9      =   7
            colname10       =   "WSOMSHPBAL"
            heading10       =   "Shp Bal"
            datatype10      =   2
            mask10          =   "####0.0"
            recalcon10      =   "WODSHPBAL"
            recalcparent10  =   "WSHPBAL"
            maxlength10     =   7
            colname11       =   "WSOMSHOWINV"
            recalcon11      =   "GrdShpOrdMst(SOMOMCOCD),GrdShpOrdMst(SOMOMTC),GrdShpOrdMst(SOMOMYY),GrdShpOrdMst(SOMOMCHR),GrdShpOrdMst(SOMOMNO)"
            maxlength11     =   10
            style11         =   0
         End
         Begin MwfCtl.MWCTL_MED ATXT 
            Height          =   285
            Index           =   0
            Left            =   4815
            TabIndex        =   13
            ToolTipText     =   "Enter Invoice Date"
            Top             =   560
            Width           =   930
            _ExtentX        =   1640
            _ExtentY        =   503
            MaxLength       =   8
            DataType        =   3
            DataField       =   "SnInDt"
            IdName          =   "SNINDT"
         End
         Begin MwfCtl.MWCTL_MED ATXT 
            Height          =   285
            Index           =   16
            Left            =   6645
            TabIndex        =   33
            ToolTipText     =   "Enter Order Pcs"
            Top             =   555
            Width           =   870
            _ExtentX        =   1535
            _ExtentY        =   503
            Alignment       =   2
            Enabled         =   0   'False
            Mask            =   "####0.0"
            MaxLength       =   7
            DataType        =   2
            ReCalcOn        =   "GrdShpOrdMst(WSOMORDPCS)"
            IdName          =   "WORDPCS"
         End
         Begin MwfCtl.MWCTL_MED ATXT 
            Height          =   285
            Index           =   28
            Left            =   1050
            TabIndex        =   9
            ToolTipText     =   "Enter Invoice Tc"
            Top             =   560
            Width           =   570
            _ExtentX        =   1005
            _ExtentY        =   503
            MaxLength       =   3
            DataType        =   4
            DataField       =   "SnInTc"
            IdName          =   "SNINTC"
         End
         Begin MwfCtl.MWCTL_MED ATXT 
            Height          =   285
            Index           =   29
            Left            =   1725
            TabIndex        =   10
            ToolTipText     =   "Enter Invoice Year"
            Top             =   555
            Width           =   330
            _ExtentX        =   582
            _ExtentY        =   503
            MaxLength       =   2
            DataType        =   4
            DataField       =   "SnInYy"
            IdName          =   "SNINYY"
         End
         Begin MwfCtl.MWCTL_MED ATXT 
            Height          =   285
            Index           =   30
            Left            =   2160
            TabIndex        =   11
            ToolTipText     =   "Enter Invoice Character"
            Top             =   555
            Width           =   570
            _ExtentX        =   1005
            _ExtentY        =   503
            MaxLength       =   3
            DataType        =   4
            DataField       =   "SnInChr"
            IdName          =   "SNINCHR"
         End
         Begin MwfCtl.MWCTL_MED ATXT 
            Height          =   285
            Index           =   31
            Left            =   2840
            TabIndex        =   12
            ToolTipText     =   "Enter Invoice No."
            Top             =   555
            Width           =   750
            _ExtentX        =   1323
            _ExtentY        =   503
            Alignment       =   2
            Mask            =   "#####0"
            MaxLength       =   6
            DataType        =   1
            DataField       =   "SnInNo"
            IdName          =   "SNINNO"
         End
         Begin MwfCtl.MWCTL_MED ATXT 
            Height          =   285
            Index           =   7
            Left            =   4815
            TabIndex        =   5
            ToolTipText     =   "Enter Shipment Note Date"
            Top             =   240
            Width           =   930
            _ExtentX        =   1640
            _ExtentY        =   503
            MaxLength       =   8
            DataType        =   3
            DataField       =   "SnDt"
            IdName          =   "SNDT"
         End
         Begin MwfCtl.MWCTL_MED ATXT 
            Height          =   285
            Index           =   5
            Left            =   6645
            TabIndex        =   6
            ToolTipText     =   "Enter Customer Code"
            Top             =   240
            Width           =   1410
            _ExtentX        =   2487
            _ExtentY        =   503
            MaxLength       =   8
            DataType        =   4
            DataField       =   "SnCmCd"
            ReCalcParent    =   "WCMNAME,SNCURCD"
            ReCalcOn        =   "WCMNAME"
            IdName          =   "SNCMCD"
         End
         Begin MwfCtl.MWCTL_MED ATXT 
            Height          =   285
            Index           =   14
            Left            =   8760
            TabIndex        =   7
            ToolTipText     =   "Enter Customer Name"
            Top             =   240
            Width           =   4890
            _ExtentX        =   8625
            _ExtentY        =   503
            Enabled         =   0   'False
            MaxLength       =   45
            DataType        =   4
            ReCalcParent    =   "SNCMCD"
            ReCalcOn        =   "SNCMCD"
            IdName          =   "WCMNAME"
         End
         Begin MwfCtl.MWCTL_MED ATXT 
            Height          =   285
            Index           =   15
            Left            =   14205
            TabIndex        =   8
            ToolTipText     =   "Enter Currency Code"
            Top             =   240
            Width           =   870
            _ExtentX        =   1535
            _ExtentY        =   503
            Enabled         =   0   'False
            MaxLength       =   5
            DataType        =   4
            DataField       =   "SnCurCd"
            ReCalcOn        =   "SNCMCD"
            IdName          =   "SNCURCD"
         End
         Begin MwfCtl.MWCTL_FLX GrdShpOrdDsg 
            Height          =   4245
            Left            =   60
            TabIndex        =   45
            Top             =   5160
            Width           =   9840
            _ExtentX        =   17357
            _ExtentY        =   7488
            Cols            =   16
            colname1        =   "SODSRNO"
            heading1        =   "Sr      "
            datafld1        =   "SodSrNo"
            datatype1       =   1
            mask1           =   "#####0"
            maxlength1      =   6
            style1          =   2
            colname2        =   "SODODSR"
            heading2        =   "OdSr    "
            datafld2        =   "SodOdSr"
            datatype2       =   1
            mask2           =   "#####0"
            recalcparent2   =   $"EmrFrmShpNote.frx":0000
            maxlength2      =   6
            colname3        =   "WSODDMCD"
            heading3        =   "Design Code                                  "
            datatype3       =   4
            recalcon3       =   $"EmrFrmShpNote.frx":00A2
            maxlength3      =   15
            colname4        =   "WSODSFX"
            heading4        =   "Suffix                 "
            datatype4       =   4
            recalcon4       =   $"EmrFrmShpNote.frx":012C
            maxlength4      =   15
            colname5        =   "WSODSZ"
            heading5        =   "Size       "
            datatype5       =   4
            recalcon5       =   $"EmrFrmShpNote.frx":01B6
            maxlength5      =   5
            colname6        =   "WSODCOL"
            heading6        =   "Color     "
            datatype6       =   4
            recalcon6       =   $"EmrFrmShpNote.frx":0240
            maxlength6      =   5
            colname7        =   "WSODKT"
            heading7        =   "KT        "
            datatype7       =   4
            recalcon7       =   $"EmrFrmShpNote.frx":02CA
            maxlength7      =   5
            colname8        =   "WSODORDPCS"
            heading8        =   "Ord Pcs"
            datatype8       =   2
            mask8           =   "####0.0"
            recalcon8       =   $"EmrFrmShpNote.frx":0354
            recalcparent8   =   "WODORDPCS,GrdShpOrdMst(WSOMORDPCS)"
            maxlength8      =   7
            colname9        =   "WSODSHPBAL"
            heading9        =   "Shp Bal"
            datatype9       =   2
            mask9           =   "-####0.0"
            recalcon9       =   $"EmrFrmShpNote.frx":03DE
            recalcparent9   =   "WODSHPBAL"
            maxlength9      =   8
            colname10       =   "SODODCOCD"
            heading10       =   "           "
            datafld10       =   "SodOdCoCd"
            datatype10      =   4
            recalcparent10  =   $"EmrFrmShpNote.frx":0468
            maxlength10     =   3
            style10         =   0
            colname11       =   "SODODTC"
            heading11       =   "        "
            datafld11       =   "SodOdTc"
            datatype11      =   4
            recalcparent11  =   $"EmrFrmShpNote.frx":050A
            maxlength11     =   3
            style11         =   0
            colname12       =   "SODODYY"
            heading12       =   "       "
            datafld12       =   "SodOdYy"
            datatype12      =   4
            recalcparent12  =   $"EmrFrmShpNote.frx":05AC
            maxlength12     =   2
            style12         =   0
            colname13       =   "SODODCHR"
            heading13       =   "             "
            datafld13       =   "SodOdChr"
            datatype13      =   4
            recalcparent13  =   $"EmrFrmShpNote.frx":064E
            maxlength13     =   3
            style13         =   0
            colname14       =   "SODODNO"
            heading14       =   "              "
            datafld14       =   "SodOdNo"
            datatype14      =   1
            mask14          =   "#####0"
            recalcparent14  =   $"EmrFrmShpNote.frx":06F0
            maxlength14     =   6
            style14         =   0
            colname15       =   "SODSR"
            heading15       =   "          "
            datafld15       =   "SodSr"
            datatype15      =   1
            mask15          =   "#####0"
            maxlength15     =   6
            style15         =   0
         End
         Begin MwfCtl.MWCTL_FLX GrdShpInv 
            Height          =   4125
            Left            =   9840
            TabIndex        =   46
            Top             =   930
            Width           =   5280
            _ExtentX        =   9313
            _ExtentY        =   7276
            Cols            =   8
            colname1        =   "WINVSR"
            heading1        =   "Sr     "
            datatype1       =   1
            mask1           =   "#####0"
            maxlength1      =   6
            style1          =   0
            colname2        =   "WINTC"
            heading2        =   "InTc"
            datatype2       =   4
            maxlength2      =   3
            style2          =   2
            tooltiptext2    =   "Enter Invoice Tc"
            colname3        =   "WINYY"
            heading3        =   "InYy"
            datatype3       =   4
            maxlength3      =   2
            style3          =   2
            tooltiptext3    =   "Enter Invoice Year"
            colname4        =   "WINCHR"
            heading4        =   "InChr "
            datatype4       =   4
            maxlength4      =   3
            style4          =   2
            tooltiptext4    =   "Enter Invoice Character"
            colname5        =   "WINNO"
            heading5        =   "InNo         "
            datatype5       =   1
            mask5           =   "#####0"
            maxlength5      =   6
            style5          =   2
            tooltiptext5    =   "Enter Invoice Number"
            colname6        =   "WINDT"
            heading6        =   "Date        "
            datatype6       =   3
            maxlength6      =   8
            style6          =   2
            colname7        =   "WINPCS"
            heading7        =   "Pcs     "
            datatype7       =   2
            mask7           =   "####0.0"
            maxlength7      =   7
            style7          =   2
         End
         Begin MwfCtl.MWCTL_FLX GrdShpInvDsg 
            Height          =   4245
            Left            =   9840
            TabIndex        =   47
            Top             =   5160
            Width           =   5280
            _ExtentX        =   9313
            _ExtentY        =   7488
            Cols            =   8
            colname1        =   "WIFTC"
            heading1        =   "IfTc "
            datatype1       =   4
            maxlength1      =   3
            style1          =   2
            tooltiptext1    =   "Enter Invoice Tc"
            colname2        =   "WIFYY"
            heading2        =   "IfYy "
            datatype2       =   4
            maxlength2      =   2
            style2          =   2
            tooltiptext2    =   "Enter Invoice Year"
            colname3        =   "WIFCHR"
            heading3        =   "IfChr   "
            datatype3       =   4
            maxlength3      =   3
            style3          =   2
            tooltiptext3    =   "Enter Invoice Character"
            colname4        =   "WIFNO"
            heading4        =   "IfNo          "
            datatype4       =   1
            mask4           =   "#####0"
            maxlength4      =   6
            style4          =   2
            tooltiptext4    =   "Enter Invoice Number"
            colname5        =   "WIFSRNO"
            heading5        =   "IfSrNo  "
            datatype5       =   1
            mask5           =   "#####0"
            maxlength5      =   6
            style5          =   2
            colname6        =   "WIFDT"
            heading6        =   "Date        "
            datatype6       =   3
            maxlength6      =   8
            style6          =   2
            colname7        =   "WIFPCS"
            heading7        =   "Pcs     "
            datatype7       =   2
            mask7           =   "##0.0"
            maxlength7      =   5
            style7          =   2
         End
         Begin MwfCtl.MWCTL_MED ATXT 
            Height          =   285
            Index           =   2
            Left            =   8760
            TabIndex        =   48
            ToolTipText     =   "Enter Order Pcs"
            Top             =   555
            Width           =   870
            _ExtentX        =   1535
            _ExtentY        =   503
            Alignment       =   2
            Enabled         =   0   'False
            Mask            =   "####0.0"
            MaxLength       =   7
            DataType        =   2
            ReCalcOn        =   "GrdShpOrdMst(WSOMSHPBAL)"
            IdName          =   "WSHPBAL"
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
            Index           =   7
            Left            =   4350
            TabIndex        =   44
            Top             =   240
            Width           =   435
         End
         Begin VB.Label ALBL 
            BackStyle       =   0  'Transparent
            Caption         =   "Inv Cust"
            BeginProperty Font 
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
            Left            =   5850
            TabIndex        =   43
            Top             =   240
            Width           =   855
         End
         Begin VB.Label ALBL 
            BackStyle       =   0  'Transparent
            Caption         =   "Name"
            BeginProperty Font 
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
            Left            =   8160
            TabIndex        =   42
            Top             =   240
            Width           =   615
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
            Index           =   18
            Left            =   13740
            TabIndex        =   41
            Top             =   240
            Width           =   495
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
            Index           =   1
            Left            =   60
            TabIndex        =   40
            Top             =   555
            Width           =   795
         End
         Begin VB.Label ALBL 
            BackStyle       =   0  'Transparent
            Caption         =   "Ord Pcs"
            BeginProperty Font 
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
            Left            =   5850
            TabIndex        =   39
            Top             =   560
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
            Index           =   30
            Left            =   1645
            TabIndex        =   38
            Top             =   555
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
            Left            =   2080
            TabIndex        =   37
            Top             =   555
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
            Left            =   2760
            TabIndex        =   36
            Top             =   555
            Width           =   105
         End
         Begin VB.Label ALBL 
            BackStyle       =   0  'Transparent
            Caption         =   "Inv Date"
            BeginProperty Font 
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
            Left            =   3990
            TabIndex        =   35
            Top             =   555
            Width           =   795
         End
         Begin VB.Label ALBL 
            BackStyle       =   0  'Transparent
            Caption         =   "Shp Bal"
            BeginProperty Font 
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
            Left            =   7920
            TabIndex        =   34
            Top             =   555
            Width           =   855
         End
      End
   End
End
Attribute VB_Name = "EmrFrmShpNote"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim moCn As MwfLib.MDOConnection
Dim ms_CoCd As String, ms_Tc As String, ms_Yy As String

Enum en_ShpNoteFra
  CpyDsg = 0
End Enum

Dim mArr_Cv As Variant, mArr_Ed As Variant

Private Sub adc_AutoGen()
  If ADC("SnNo") = 0 Then
    ADC("SnNo") = moCn.AutoGen("ShpNote", "SnNo", ADC("SnCoCd") + ADC("SnTc") + ADC("SnYy") + ADC("SnChr"))
  Else
    Call moCn.AutoGen("ShpNote", "SnNo", ADC("SnCoCd") + ADC("SnTc") + ADC("SnYy") + ADC("SnChr"), ADC("SnNo"))
  End If
End Sub

Private Sub ADC_Load()
  ReDim mArr_Cv(1), mArr_Ed(1)
  Set moCn = ADC.Connection
  Set ADC.FirKeyCtl = ADC("SnCoCd")
  Set ADC.FirNKeyCtl = ADC("SnDt")
    
  ADC.Child = "GrdShpOrdMst"
  GrdShpOrdMst.Child = "GrdShpOrdDsg,GrdShpInv"
  GrdShpOrdDsg.Child = "GrdShpInvDsg"
  Set GrdShpOrdMst.PrevCtl = ADC("SnInDt")
  
  Set GrdShpOrdDsg.PrevCtl = GrdShpOrdMst
  Set GrdShpOrdDsg.NextCtl = GrdShpOrdMst
  GrdShpOrdMst.HotKey = "A"
  GrdShpOrdDsg.HotKey = "B"
  FraNKeyAll.ZOrder (0)
  
  ms_CoCd = gs_CoCd
  ms_Tc = moCn.GetFldVal("Select PValue from Param where PTyp= 'TCTYP' and PMCd= 'SN' ")
  ms_Yy = moCn.GetFldVal("Select hYy from Head where HCoCd= '" + gs_CoCd + "' and HCd= '" + ctSelfCmCd + "'")
    
End Sub

Private Sub ADC_NonKeyWhen(pr_Cancel As Boolean, pr_ErrMsg As String)
  ms_CoCd = ADC("SnCoCd")
  ms_Tc = ADC("SnTc")
  ms_Yy = ADC("SnYy")
      
  If ADC.AddFndMode = xaddmode Then
    Dim ws_Dt As String
    ws_Dt = GetNextDate(ADC("SnCoCd"))
    If ws_Dt <> "" Then ADC("SnDt") = ws_Dt
    ADC("SnInDt") = MWLib.EmptyDate
  End If
  If GrdShpOrdMst.Rows > 1 Then GrdShpOrdMst.Row = 1
  If GrdShpOrdDsg.Rows > 1 Then GrdShpOrdDsg.Row = 1
  Call EnaDisaCmds(False)

End Sub
Private Sub ADC_KeyWhen()
  Call EnaDisaCmds(True)
  Call HideAllFras
      
  If ms_CoCd <> "" Then ADC("SnCoCd") = ms_CoCd
  If ms_Tc <> "" Then ADC("SnTc") = ms_Tc
  If ms_Yy <> "" Then ADC("SnYy") = ms_Yy
     
End Sub

Private Sub Form_Activate()
  Call FrmActivate(Me)
End Sub
Private Sub Form_Deactivate()
  Call FrmDeActivate(Me)
End Sub
Private Sub Form_GotFocus()
  SetMDISysColorsAfterLogin
  
End Sub
Private Sub Form_Load()
  Dim ws_BtnStr As String
  ws_BtnStr = "CmdCpyDsg"
  Call SetObjColors(Me, ws_BtnStr)
  
End Sub

Private Sub Form_KeyUp(KeyCode As Integer, Shift As Integer)
  If ADC.HotKeys(KeyCode, Shift) Then
    KeyCode = 0
    Shift = 0
  End If
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    If UnloadMode <> vbFormCode Then Cancel = True: Exit Sub
    Set moCn = Nothing
End Sub


Private Sub ADC_SetRecSource()
  '*** Set the RecSource for ADC (Table ShpNote)
  ADC.RecSource = " Select * from ShpNote Where SnCoCd ='" + ADC("SnCoCd") + "' And SnTc ='" + ADC("SnTc") + "'" + _
                  " And SnYy = '" + ADC("SnYy") + "' And SnChr ='" + ADC("SnChr") + "' And SnNo =" + CStr(ADC("SnNo"))

End Sub


Private Sub ATXT_Validate(Index As Integer, Cancel As Boolean)
  Cancel = ATXT(Index).Validate
End Sub
Private Sub ATXT_FldChange(Index As Integer, ByVal IdName As String, Cancel As Boolean, ErrMsg As String, ByVal pv_NewValue As Variant)
  Select Case IdName
  Case Is = UCase("SnCoCd")
    Cancel = Not moCn.RecSeek("Select 'x' From Head Where HCd='" + ctSelfCmCd + "' And HCoCd='" + pv_NewValue + "'")
    If Cancel = True Then ErrMsg = "Invalid Company Code": Exit Sub
  Case Is = UCase("SnTc")
    Cancel = Not moCn.RecSeek("select PMCd from Param where PTyp= 'TC' and " + _
                         "PMCd = '" + pv_NewValue + "' and PValue= 'SN' ")
    If Cancel = True Then ErrMsg = "Invalid Shipment Note Tc": Exit Sub
  Case Is = UCase("SnYy")
    Cancel = Not moCn.RecSeek("select PMCd from Param where PTyp= 'YY' and PMCd = '" + pv_NewValue + "'")
    If Cancel = True Then ErrMsg = "Invalid Year": Exit Sub
  Case Is = UCase("SnChr")
    Cancel = pv_NewValue <> "" And Not moCn.RecSeek("select vPSCd from vParam where " + _
                         "vPCoCd= '" + ADC("SnCoCd") + "' and vPTyp= 'CHR' and " + _
                         "vPMCd = '" + ADC("SnTc") + "' and vPSCd='" + pv_NewValue + "'")
    If Cancel = True Then ErrMsg = "Invalid Shipment Voucher Character": Exit Sub

  Case Is = UCase("SnCmCd")
    Cancel = Not moCn.RecSeek("select CmCd from CustMst where CmCtg= 'C' and CmCd= '" + pv_NewValue + "'")
    If Cancel = True Then ErrMsg = "Invalid Customer Code": Exit Sub
  Case Is = UCase("wCmName")
    Cancel = Not moCn.RecSeek("select CmName from CustMst where CmCtg= 'C' and CmName= '" + pv_NewValue + "'")
    If Cancel = True Then ErrMsg = "Invalid Customer Name": Exit Sub
  Case Is = UCase("SnCurCd")
    Cancel = (Not moCn.RecSeek("select 'x' from Param where PTyp= 'CURNCY' and PMCd = '" + pv_NewValue + "' and PSCd= ''"))
    If Cancel = True Then ErrMsg = "Invalid Currency Code": Exit Sub
  Case Is = UCase("SnInTc")
    Cancel = pv_NewValue <> "" And Not moCn.RecSeek("select PMCd from Param where PTyp= 'TC' and " + _
                         "PMCd = '" + pv_NewValue + "' and PValue= 'IN' ")
    If Cancel = True Then ErrMsg = "Invalid Invoice Tc": Exit Sub
  Case Is = UCase("SnInYy")
    Cancel = pv_NewValue <> "" And Not moCn.RecSeek("select PMCd from Param where PTyp= 'YY' and PMCd = '" + pv_NewValue + "'")
    If Cancel = True Then ErrMsg = "Invalid Year": Exit Sub
  Case Is = UCase("SnInChr")
    Cancel = pv_NewValue <> "" And (Not moCn.RecSeek("select vPSCd from vParam where " + _
                         "vPCoCd= '" + ADC("SnCoCd") + "' and vPTyp= 'CHR' and " + _
                         "vPMCd = '" + ADC("SnInTc") + "' and vPSCd='" + pv_NewValue + "'"))
    If Cancel = True Then ErrMsg = "Invalid Invoice Character": Exit Sub
  Case Is = UCase("SnInNo")
    Cancel = pv_NewValue <> 0 And (Not moCn.RecSeek("Select InNo From InvHd Where InCoCd= '" + ADC("SnCoCd") + "' and " + _
                          "InTc= '" + ADC("SnInTc") + "' And InYy= '" + ADC("SnInYy") + "' and " + _
                          "InChr= '" + ADC("SnInChr") + "' and InNo= " + CStr(pv_NewValue)))
    If Cancel = True Then ErrMsg = "Invalid Invoice Number": Exit Sub
    ADC("SnInDt") = moCn.GetFldVal("Select InDt From InvHd Where InCoCd= '" + ADC("SnCoCd") + "' and " + _
                          "InTc= '" + ADC("SnInTc") + "' And InYy= '" + ADC("SnInYy") + "' and " + _
                          "InChr= '" + ADC("SnInChr") + "' and InNo= " + CStr(pv_NewValue))
    
    Cancel = pv_NewValue <> 0 And (Not moCn.RecSeek("Select InCmCd From InvHd Where InCoCd= '" + ADC("SnCoCd") + "' and " + _
                          "InTc= '" + ADC("SnInTc") + "' And InYy= '" + ADC("SnInYy") + "' and " + _
                          "InChr= '" + ADC("SnInChr") + "' and InNo= " + CStr(pv_NewValue) + " and InCmCd='" + ADC("SnCmCd") + "'"))
    If Cancel = True Then ErrMsg = "Customer For This Invoice Number And Invoice Customer Are Not Matching ": Exit Sub

    
  End Select
End Sub

Private Sub ATXT_FldWhen(Index As Integer, ByVal IdName As String, Cancel As Boolean, ErrMsg As String)
  If GrdShpOrdMst.Rows > 1 Then
    Select Case IdName
    Case Is = UCase("SnDt"), UCase("SnCmCd"), UCase("wCmName"), UCase("SnCurCd"), _
              UCase("SnInTc"), UCase("SnInYy"), UCase("SnInChr"), UCase("SnInNo"), _
              UCase("SnInDt"), UCase("wOrdPcs"), UCase("wShpBal")
      Cancel = True
      ErrMsg = "Cannot Edit Value when Records exist in Order Grid"
      Exit Sub
    End Select
  End If
  
  '*** Set Helps For Each Field
  Select Case IdName
  Case Is = UCase("SnCoCd")
    Call HlpList.hCoCd
  Case Is = UCase("SnTc")
    Call HlpList.PMCd("TC", "'SN'")
  Case Is = UCase("SnYy")
    Call HlpList.PMCd("YY")
  Case Is = UCase("SnChr")
    Call HlpList.vPSCd(ADC("SnCoCd"), "CHR", ADC("SnTc"), "Y")
  Case Is = UCase("SnNo")
    Call HlpList.ShpNoteNo(ADC("SnCoCd"), ADC("SnTc"), ADC("SnYy"), ADC("SnChr"))

  Case Is = UCase("SnCmCd")
    Call HlpList.CustCd("C")
  
  Case Is = UCase("SnInTc")
    Call HlpList.PMCd("TC", "'IN'")
  Case Is = UCase("SnInYy")
    Call HlpList.PMCd("YY")
  Case Is = UCase("SnInChr")
    Call HlpList.vPSCd(ADC("SnCoCd"), "CHR", ADC("SnInTc"), "Y")
  Case Is = UCase("SnInNo")
    Call HlpList.InNo(ADC("SnCoCd"), ADC("SnInTc"), ADC("SnInYy"), ADC("SnInChr"))

  Case Is = UCase("wOdSrFr")
  If GrdShpOrdMst.Row > 1 And GrdShpOrdMst.Row <= GrdShpOrdMst.Rows - 1 Then Call HlpList.OdSr(GrdShpOrdMst.Value(GrdShpOrdMst.Row, "SomOmCoCd"), GrdShpOrdMst.Value(GrdShpOrdMst.Row, "SomOmTc"), GrdShpOrdMst.Value(GrdShpOrdMst.Row, "SomOmYy"), GrdShpOrdMst.Value(GrdShpOrdMst.Row, "SomOmChr"), GrdShpOrdMst.Value(GrdShpOrdMst.Row, "SomOmNo"))
  Case Is = UCase("wOdSrTo")
    If GrdShpOrdMst.Row > 1 And GrdShpOrdMst.Row <= GrdShpOrdMst.Rows - 1 Then Call HlpList.OdSr(GrdShpOrdMst.Value(GrdShpOrdMst.Row, "SomOmCoCd"), GrdShpOrdMst.Value(GrdShpOrdMst.Row, "SomOmTc"), GrdShpOrdMst.Value(GrdShpOrdMst.Row, "SomOmYy"), GrdShpOrdMst.Value(GrdShpOrdMst.Row, "SomOmChr"), GrdShpOrdMst.Value(GrdShpOrdMst.Row, "SomOmNo"))

  End Select
End Sub

Private Sub atxt_ReCalculate(Index As Integer, ByVal IdName As String, ByVal OldVal As Variant, ByVal RowNum As Integer, ByVal SourceName As String, ByVal SourceColName As String)
Dim i As Integer, wTotOrdPcs As Double, wTotShpBal As Double
  Select Case IdName
  Case Is = UCase("wCmName")
    Dim wCustName As String
    If ADC.Mode = xNorm Or ADC.Mode = xInit Then
      ADC("wCmName") = moCn.GetFldVal("Select CmName from CustMst where CmCtg='C' and CmCd= '" + ADC("SnCmCd") + "'")
    End If
  Case Is = UCase("SnCurCd")
    If ADC.Mode = xNorm Or ADC.Mode = xInit Then
      ADC("SnCurCd") = moCn.GetFldVal("Select CmCurCd from CustMst where CmCtg= 'C' and CmCd= '" + ADC("SnCmCd") + "'")
    End If
    
  Case Is = UCase("wOdOrdPcs")
    wTotOrdPcs = 0
    If GrdShpOrdDsg.Rows > 0 Then
      For i = 1 To GrdShpOrdDsg.Rows - 1
        If Not GrdShpOrdDsg.IsDel(i) Then
          wTotOrdPcs = wTotOrdPcs + GrdShpOrdDsg.Value(i, "wSodOrdPcs")
        End If
      Next i
      ADC("wOdOrdPcs") = wTotOrdPcs
    End If
  Case Is = UCase("wOdShpBal")
    wTotShpBal = 0
    If GrdShpOrdDsg.Rows > 0 Then
      For i = 1 To GrdShpOrdDsg.Rows - 1
        If Not GrdShpOrdDsg.IsDel(i) Then
          wTotShpBal = wTotShpBal + GrdShpOrdDsg.Value(i, "wSodShpBal")
        End If
      Next i
      ADC("wOdShpBal") = wTotShpBal
    End If
  Case Is = UCase("wOrdPcs")
    wTotOrdPcs = 0
    If GrdShpOrdMst.Rows > 0 Then
      For i = 1 To GrdShpOrdMst.Rows - 1
        If Not GrdShpOrdMst.IsDel(i) Then
          wTotOrdPcs = wTotOrdPcs + GrdShpOrdMst.Value(i, "wSomOrdPcs")
        End If
      Next i
      ADC("wOrdPcs") = wTotOrdPcs
    End If
  Case Is = UCase("wShpBal")
    wTotShpBal = 0
    If GrdShpOrdMst.Rows > 0 Then
      For i = 1 To GrdShpOrdMst.Rows - 1
        If Not GrdShpOrdMst.IsDel(i) Then
          wTotShpBal = wTotShpBal + GrdShpOrdMst.Value(i, "wSomShpBal")
        End If
      Next i
      ADC("wShpBal") = wTotShpBal
    End If
     
  End Select
End Sub

Private Sub GrdShpInv_ReCalc(ByVal RowNum As Integer, ByVal ColName As String, ByVal OldVal As Variant, ByVal SourceName As String, ByVal SourceColName As String)
'If GrdShpInv.Mode = fgminit Then
'  GrdShpInv.Value(RowNum, wSomShpBal) = 15
'End If
End Sub

Private Sub GrdShpOrdMst_ColChange(ByVal RowNum As Integer, ByVal ColName As String, ByVal NewValue As Variant, Cancel As Boolean, ErrMsg As String)
With GrdShpOrdMst
  Select Case ColName
  Case Is = UCase("SomOmTc")
    Cancel = Not moCn.RecSeek("select PMCd from Param where PTyp= 'TC' and " + _
                         "PMCd = '" + NewValue + "' and PValue= 'SO' ")
    If Cancel = True Then ErrMsg = "Invalid Order Tc": Exit Sub
  Case Is = UCase("SomOmYy")
    Cancel = Not moCn.RecSeek("select PMCd from Param where PTyp= 'YY' and PMCd = '" + NewValue + "'")
    If Cancel = True Then ErrMsg = "Invalid Order Year": Exit Sub
  Case Is = UCase("SomOmChr")
    Cancel = Not moCn.RecSeek("select vPSCd from vParam where " + _
                         "vPCoCd= '" + .Value(RowNum, "SomOmCoCd") + "' and vPTyp= 'CHR' and " + _
                         "vPMCd = '" + .Value(RowNum, "SomOmTc") + "' and vPSCd='" + NewValue + "'")
    If Cancel = True Then ErrMsg = "Invalid Order Character": Exit Sub
  Case Is = UCase("SomOmNo")
    Cancel = Not moCn.RecSeek("select OmNo from OrdMst where OmCoCd= '" + .Value(RowNum, "SomOmCoCd") + "' and " + _
                        "OmTc= '" + .Value(RowNum, "SomOmTc") + "' and OmYy= '" + .Value(RowNum, "SomOmYy") + "' and " + _
                        "OmChr= '" + .Value(RowNum, "SomOmChr") + "' and OmNo= " + CStr(NewValue))
    ErrMsg = "Invalid Order Number": Exit Sub
  End Select
End With

End Sub

Private Sub GrdShpOrdMst_ColWhen(ByVal RowNum As Integer, ByVal ColName As String, Cancel As Boolean, ErrMsg As String, ByVal Action As MwfCtl.en_fgAction)
With GrdShpOrdMst
  If GrdShpOrdDsg.Rows > 1 Then
    Select Case ColName
    Case Is = UCase("SomOmTc"), UCase("SomOmYy"), UCase("SomOmChr"), UCase("SomOmNo")
      Cancel = True
      ErrMsg = "Cannot Edit Value when Records exist in Order Design"
      Exit Sub
    End Select
  End If
  
  Select Case ColName
  Case Is = UCase("SomOmTc")
    Call HlpList.PMCd("TC", "'SO'")
  Case Is = UCase("SomOmYy")
    Call HlpList.PMCd("YY")
  Case Is = UCase("SomOmChr")
    Call HlpList.vPSCd(.Value(RowNum, "SomOmCoCd"), "CHR", .Value(RowNum, "SomOmTc"), "Y")
  Case Is = UCase("SomOmNo")
    Call HlpList.OmNo(.Value(RowNum, "SomOmCoCd"), .Value(RowNum, "SomOmTc"), .Value(RowNum, "SomOmYy"), .Value(RowNum, "SomOmChr"), , IIF(gs_Partition = ctCurrPrtn, ctCurrPrtn, ""))
  Case Is = UCase("WSOMCMCD"), UCase("WSOMEXPDELDT"), UCase("WSOMORDPCS"), UCase("WSOMSHPBAL")
    Cancel = True
  End Select
End With
End Sub

Private Sub GrdShpOrdMst_InitKey(ByVal RowNum As Integer)
  With GrdShpOrdMst
    If .IsNew(RowNum) Then
      .Store "SomCoCd", ADC("SnCoCd")
      .Store "SomTc", ADC("SnTc")
      .Store "SomYy", ADC("SnYy")
      .Store "SomChr", ADC("SnChr")
      .Store "SomNo", ADC("SnNo")
      .Store "SomSnIdNo", moCn.GetFldVal("Select SnIdNo From ShpNote Where SnCoCd='" + ADC("SnCoCd") + "' And SnTc='" + ADC("SnTc") + _
                  "' And SnYy='" + ADC("SnYy") + "' And SnChr='" + ADC("SnChr") + "' And SnNo=" + CStr(ADC("SnNo")))

    End If
  End With
End Sub
Private Sub GrdShpOrdDsg_RowPostValid(ByVal RowNum As Integer, Cancel As Boolean, ErrMsg As String, ChangedColName As String)
  If Not moCn.RecSeek("Select 'A' from OrdDsg where OdCoCd='" + GrdShpOrdDsg.Value(RowNum, "SodOdCoCd") + _
                            "' and OdTc='" + GrdShpOrdDsg.Value(RowNum, "SodOdTc") + "' and OdYy='" + GrdShpOrdDsg.Value(RowNum, "SodOdYy") + _
                            "' and OdChr='" + GrdShpOrdDsg.Value(RowNum, "SodOdChr") + "' and OdNo=" + CStr(GrdShpOrdDsg.Value(RowNum, "SodOdNo")) + " AND OdSr= " + CStr(GrdShpOrdDsg.Value(RowNum, "SOdOdSr"))) Then
  
    Cancel = True: ErrMsg = "||Order Design is invalid"
    Exit Sub
  End If
  If Not moCn.RecSeek("Select 'A' from OrdDsg where OdCoCd='" + GrdShpOrdDsg.Value(RowNum, "SodOdCoCd") + _
                            "' and OdTc='" + GrdShpOrdDsg.Value(RowNum, "SodOdTc") + "' and OdYy='" + GrdShpOrdDsg.Value(RowNum, "SodOdYy") + _
                            "' and OdChr='" + GrdShpOrdDsg.Value(RowNum, "SodOdChr") + "' and OdNo=" + CStr(GrdShpOrdDsg.Value(RowNum, "SodOdNo")) + " AND OdSr= " + CStr(GrdShpOrdDsg.Value(RowNum, "SOdOdSr")) + " and (OdOrdQty>OdExpQty and OdCls='N') ") Then
  
    Cancel = True: ErrMsg = "||Order Design is fully exported "
    Exit Sub
  End If
End Sub
Private Sub GrdShpOrdMst_RowPostValid(ByVal RowNum As Integer, Cancel As Boolean, ErrMsg As String, ChangedColName As String)
Dim wSrcCurCd As String, wTarCurCd As String
With GrdShpOrdMst
  wTarCurCd = moCn.GetFldVal("Select OmCmCurCd from OrdMst where OmCoCd='" + .Value(RowNum, "SomOmCoCd") + _
                            "' and OmTc='" + .Value(RowNum, "SomOmTc") + "' and OmYy='" + .Value(RowNum, "SomOmYy") + _
                            "' and OmChr='" + .Value(RowNum, "SomOmChr") + "' and OmNo=" + CStr(.Value(RowNum, "SomOmNo")))
  
  If ADC("SnCurCd") <> wTarCurCd Then
    Cancel = True: ErrMsg = "||Order Currency Is Not Matched With Currency for given Customer"
    Exit Sub
  End If
  If Not moCn.RecSeek("Select 'A' from OrdMst where OmCoCd='" + .Value(RowNum, "SomOmCoCd") + _
                            "' and OmTc='" + .Value(RowNum, "SomOmTc") + "' and OmYy='" + .Value(RowNum, "SomOmYy") + _
                            "' and OmChr='" + .Value(RowNum, "SomOmChr") + "' and OmNo=" + CStr(.Value(RowNum, "SomOmNo"))) Then
  
    Cancel = True: ErrMsg = "||Order is invalid "
    Exit Sub
  End If
  If Not moCn.RecSeek("Select 'A' from OrdMst where OmCoCd='" + .Value(RowNum, "SomOmCoCd") + _
                            "' and OmTc='" + .Value(RowNum, "SomOmTc") + "' and OmYy='" + .Value(RowNum, "SomOmYy") + _
                            "' and OmChr='" + .Value(RowNum, "SomOmChr") + "' and OmNo=" + CStr(.Value(RowNum, "SomOmNo")) + " AND exists  (Select 'a' from OrdDsg where OdOmIdNo = OmIdNo and (OdOrdQty>OdExpQty and OdCls='N')) ") Then
  
    Cancel = True: ErrMsg = "||Order is fully exported "
    Exit Sub
  End If
End With

End Sub

Private Sub GrdShpOrdMst_RowWhen(ByVal RowNum As Integer)
With GrdShpOrdMst
  If .IsNew(RowNum) Then
    If .Value(RowNum, "SomSr") = 0 Then .Value(RowNum, "SomSr") = .MaxVal("SomSr") + 1
    .Value(RowNum, "SomOmCoCd") = ADC("SnCoCd")
  Else
    If RowNum > 0 And .Value(RowNum, "SomOmCoCd") <> "" And .Value(RowNum, "SomOmTc") <> "" + _
      .Value(RowNum, "SomOmYy") <> "" And .Value(RowNum, "SomOmChr") <> "" And .Value(RowNum, "SomOmNo") <> 0 Then
        Call FillGrdShpInv(RowNum)
    End If
'    If ADC("wOdOrdPcs") <> 0 Then .Value(RowNum, "wSomOrdPcs") = ADC("wOdOrdPcs")
  End If
End With
  If GrdShpOrdDsg.Rows > 1 Then
    GrdShpOrdDsg.Row = 1
    Call FillGrdShpInvDsg(1)
  End If
End Sub
Private Sub GrdShpOrdMst_SetRecSource()
  '*** Set the RecSource for GridA (Table GrdShpOrdMst)
  GrdShpOrdMst.RecSource = " Select * from ShpNoteOrdMst Where SomCoCd ='" + ADC("SnCoCd") + "' And SomTc ='" + ADC("SnTc") + "'" + _
                  " And SomYy = '" + ADC("SnYy") + "' And SomChr ='" + ADC("SnChr") + "' And SomNo =" + CStr(ADC("SnNo"))

End Sub

Private Sub GrdShpOrdMst_ReCalc(ByVal RowNum As Integer, ByVal ColName As String, ByVal OldVal As Variant, ByVal SourceName As String, ByVal SourceColName As String)
With GrdShpOrdMst
  Select Case ColName
  Case Is = UCase("wSomCmCd")
    If .Mode = fgmnorm Or .Mode = fgminit Then
      .Value(RowNum, "wSomCmCd") = moCn.GetFldVal("Select OmCmCd From OrdMst where " + _
                                   "OmCoCd= '" + .Value(RowNum, "SomOmCoCd") + "' and " + _
                                   "OmTc= '" + .Value(RowNum, "SomOmTc") + "' and " + _
                                   "OmYy= '" + .Value(RowNum, "SomOmYy") + "' and " + _
                                   "OmChr= '" + .Value(RowNum, "SomOmChr") + "' and " + _
                                   "OmNo= " + CStr(.Value(RowNum, "SomOmNo")))
    End If
  Case Is = UCase("wSomExpDelDt")
    If .Mode = fgmnorm Or .Mode = fgminit Then
      .Value(RowNum, "wSomExpDelDt") = moCn.GetFldVal("Select OmExpDelDt From OrdMst where " + _
                                   "OmCoCd= '" + .Value(RowNum, "SomOmCoCd") + "' and " + _
                                   "OmTc= '" + .Value(RowNum, "SomOmTc") + "' and " + _
                                   "OmYy= '" + .Value(RowNum, "SomOmYy") + "' and " + _
                                   "OmChr= '" + .Value(RowNum, "SomOmChr") + "' and " + _
                                   "OmNo= " + CStr(.Value(RowNum, "SomOmNo")))
    End If
  Case Is = UCase("wSomShowInv")
    If .Mode = fgmnorm Then
      If .IsNew(RowNum) Then
        If RowNum > 0 And .Value(RowNum, "SomOmCoCd") <> "" And .Value(RowNum, "SomOmTc") <> "" + _
          .Value(RowNum, "SomOmYy") <> "" And .Value(RowNum, "SomOmChr") <> "" And .Value(RowNum, "SomOmNo") <> 0 Then
          Call FillGrdShpInv(RowNum)
        End If
      End If
    End If
  
  Case Is = UCase("wSomOrdPcs")
      Dim e As Integer
      If .Mode = fgminit Then
      With GrdShpOrdMst
          If RowNum <= 0 Or RowNum > GrdShpOrdMst.Rows - 1 Then Exit Sub
          GrdShpOrdMst.Value(RowNum, "wSomOrdPcs") = moCn.GetFldVal("Select sum(OdOrdQty) From OrdDsg where " + _
                                   "OdCoCd= '" + .Value(RowNum, "SomOmCoCd") + "' and " + _
                                   "OdTc= '" + .Value(RowNum, "SomOmTc") + "' and " + _
                                   "OdYy= '" + .Value(RowNum, "SomOmYy") + "' and " + _
                                   "OdChr= '" + .Value(RowNum, "SomOmChr") + "' and " + _
                                   "OdNo= " + CStr(.Value(RowNum, "SomOmNo")))

          Dim wsOrdShpBal As Double
      
          wsOrdShpBal = moCn.GetFldVal("Select sum(isnull(IfQty,0)) from OrdDsg " + _
                                    "left outer join Bag  on OdCoCd =BCoCd and OdTc =BOdTc and OdYy =BOdYy and OdChr =BOdChr and OdNo=BOdNo and OdSr =BOdSr " + _
                                    "left outer join InvFgd on BCoCd =IfCoCd and BYy =IfFdBYy and BChr =IfFdBChr and BNo =IfFdBNo " + _
                                    "left outer join InvHd on InCoCd=IfCoCd and InTc=IfTc and InYy=IfYy and InChr=IfChr and InNo=IfNo " + _
                                    "where OdCoCd ='" + .Value(RowNum, "SomOmCoCd") + _
                                    "' and OdTc ='" + .Value(RowNum, "SomOmTc") + _
                                    "' and OdYy ='" + .Value(RowNum, "SomOmYy") + _
                                    "' and OdChr ='" + .Value(RowNum, "SomOmChr") + _
                                    "' and OdNo =" + CStr(.Value(RowNum, "SomOmNo")))
         .Value(RowNum, "wSomShpBal") = GrdShpOrdMst.Value(RowNum, "wSomOrdPcs") - wsOrdShpBal
      End With
      End If
      If .Mode = fgmnorm Or .Mode = fgmCopyOver Then
        e = GrdShpOrdMst.Row
        If e <= 0 Or e > GrdShpOrdMst.Rows - 1 Then Exit Sub
        Dim i As Integer
        Dim c As Integer
        For i = 1 To GrdShpOrdDsg.Rows - 1
          If Not GrdShpOrdDsg.IsDel(i) Then
            c = c + GrdShpOrdDsg.Value(i, "WsODORDPCS")
          End If
        Next i
        .Value(e, "wSomOrdPcs") = c
      End If
  Case Is = UCase("wSomShpBal")
      If .Mode = fgmnorm Or .Mode = fgmCopyOver Then
        e = GrdShpOrdMst.Row
        If e <= 0 Or e > GrdShpOrdMst.Rows - 1 Then Exit Sub
        For i = 1 To GrdShpOrdDsg.Rows - 1
          If Not GrdShpOrdDsg.IsDel(i) Then
            c = c + GrdShpOrdDsg.Value(i, "WsODshpbal")
          End If
        Next i
        .Value(e, "wSomshpbal") = c
      End If
  End Select
End With
End Sub

Private Sub GrdShpOrdMst_Validate(Cancel As Boolean)
  Cancel = GrdShpOrdMst.Validate
End Sub


Private Sub GrdShpOrdDsg_ColChange(ByVal RowNum As Integer, ByVal ColName As String, ByVal NewValue As Variant, Cancel As Boolean, ErrMsg As String)
With GrdShpOrdDsg
  Select Case ColName
  Case Is = UCase("SodOdSr")
    Cancel = Not moCn.RecSeek("select OdSr from OrdDsg where OdCoCd= '" + .Value(RowNum, "SodOdCoCd") + "' and " + _
                        "OdTc= '" + .Value(RowNum, "SodOdTc") + "' and OdYy= '" + .Value(RowNum, "SodOdYy") + "' and " + _
                        "OdChr= '" + .Value(RowNum, "SodOdChr") + "' and OdNo= " + CStr(.Value(RowNum, "SodOdNo")) + " and OdSr=" + CStr(NewValue))
    ErrMsg = "Invalid Order Serial": Exit Sub
  End Select
End With

End Sub

Private Sub GrdShpOrdDsg_ColWhen(ByVal RowNum As Integer, ByVal ColName As String, Cancel As Boolean, ErrMsg As String, ByVal Action As MwfCtl.en_fgAction)
With GrdShpOrdDsg
  Select Case ColName
  Case Is = UCase("WSODDMCD"), UCase("WSODSFX"), UCase("WSODSZ"), UCase("WSODCOL"), UCase("WSODKT"), UCase("wsOdOrdPcs"), UCase("WSODSHPBAL")
    Cancel = True
  Case Is = UCase("SodOdSr")
    Call HlpList.OdSr(.Value(RowNum, "SodOdCoCd"), .Value(RowNum, "SodOdTc"), .Value(RowNum, "SodOdYy"), .Value(RowNum, "SodOdChr"), .Value(RowNum, "SodOdNo"), True)
  End Select
End With
End Sub

Private Sub GrdShpOrdDsg_InitKey(ByVal RowNum As Integer)
  With GrdShpOrdDsg
    If .IsNew(RowNum) Then
      .Store "SodCoCd", ADC("SnCoCd")
      .Store "SodTc", ADC("SnTc")
      .Store "SodYy", ADC("SnYy")
      .Store "SodChr", ADC("SnChr")
      .Store "SodNo", ADC("SnNo")
      .Store "SodSomIdNo", moCn.GetFldVal("Select SomIdNo From ShpNoteOrdMst Where SomCoCd='" + ADC("SnCoCd") + "' And SomTc='" + ADC("SnTc") + _
                  "' And SomYy='" + ADC("SnYy") + "' And SomChr='" + ADC("SnChr") + "' And SomNo=" + CStr(ADC("SnNo")) + _
                  " and SomSr=" + CStr(GrdShpOrdMst.Value(GrdShpOrdMst.Row, "SomSr")))

      .Store "SodSnIdNo", moCn.GetFldVal("Select SnIdNo From ShpNote Where SnCoCd='" + ADC("SnCoCd") + "' And SnTc='" + ADC("SnTc") + _
                  "' And SnYy='" + ADC("SnYy") + "' And SnChr='" + ADC("SnChr") + "' And SnNo=" + CStr(ADC("SnNo")))

    End If
  End With
End Sub

Private Sub GrdShpOrdDsg_RowWhen(ByVal RowNum As Integer)
With GrdShpOrdDsg
  If .IsNew(RowNum) Then
    If .Value(RowNum, "SodSrNo") = 0 Then .Value(RowNum, "SodSrNo") = .MaxVal("SodSrNo") + 1
    .Value(RowNum, "SodOdCoCd") = GrdShpOrdMst.Value(GrdShpOrdMst.Row, "SomOmCoCd")
    .Value(RowNum, "SodOdTc") = GrdShpOrdMst.Value(GrdShpOrdMst.Row, "SomOmTc")
    .Value(RowNum, "SodOdYy") = GrdShpOrdMst.Value(GrdShpOrdMst.Row, "SomOmYy")
    .Value(RowNum, "SodOdChr") = GrdShpOrdMst.Value(GrdShpOrdMst.Row, "SomOmChr")
    .Value(RowNum, "SodOdNo") = GrdShpOrdMst.Value(GrdShpOrdMst.Row, "SomOmNo")
    .Value(RowNum, "SodSr") = GrdShpOrdMst.Value(GrdShpOrdMst.Row, "SomSr")
  Else
    If RowNum > 0 And .Value(RowNum, "SodOdCoCd") <> "" And .Value(RowNum, "SodOdTc") <> "" + _
      .Value(RowNum, "SodOdYy") <> "" And .Value(RowNum, "SodOdChr") <> "" And .Value(RowNum, "SodOdNo") <> 0 Then
        Call FillGrdShpInvDsg(RowNum)
    End If
  End If
End With
End Sub

Private Sub GrdShpOrdDsg_SetRecSource()
  '*** Set the RecSource for GridA (Table GrdShpOrdDsg)
  GrdShpOrdDsg.RecSource = " Select * from ShpNoteOrdDsg Where SodCoCd ='" + ADC("SnCoCd") + "' And SodTc ='" + ADC("SnTc") + "'" + _
                  " And SodYy = '" + ADC("SnYy") + "' And SodChr ='" + ADC("SnChr") + _
                  "' And SodNo =" + CStr(ADC("SnNo")) + " and SodSr=" + CStr(GrdShpOrdMst.Value(GrdShpOrdMst.Row, "SomSr"))

End Sub

Private Sub GrdShpOrdDsg_ReCalc(ByVal RowNum As Integer, ByVal ColName As String, ByVal OldVal As Variant, ByVal SourceName As String, ByVal SourceColName As String)
With GrdShpOrdDsg
  Select Case ColName
  Case Is = UCase("wSodDmCd")
    If .Mode = fgmnorm Or .Mode = fgminit Then
      .Value(RowNum, "wSodDmCd") = moCn.GetFldVal("Select OdDmCd From OrdDsg where " + _
                                   "OdCoCd= '" + .Value(RowNum, "SodOdCoCd") + "' and " + _
                                   "OdTc= '" + .Value(RowNum, "SodOdTc") + "' and " + _
                                   "OdYy= '" + .Value(RowNum, "SodOdYy") + "' and " + _
                                   "OdChr= '" + .Value(RowNum, "SodOdChr") + "' and " + _
                                   "OdNo= " + CStr(.Value(RowNum, "SodOdNo")) + _
                                   " and OdSr=" + CStr(.Value(RowNum, "SodOdSr")))
    End If
  Case Is = UCase("wSodSfx")
    If .Mode = fgmnorm Or .Mode = fgminit Then
      .Value(RowNum, "wSodSfx") = moCn.GetFldVal("Select OdSfx From OrdDsg where " + _
                                   "OdCoCd= '" + .Value(RowNum, "SodOdCoCd") + "' and " + _
                                   "OdTc= '" + .Value(RowNum, "SodOdTc") + "' and " + _
                                   "OdYy= '" + .Value(RowNum, "SodOdYy") + "' and " + _
                                   "OdChr= '" + .Value(RowNum, "SodOdChr") + "' and " + _
                                   "OdNo= " + CStr(.Value(RowNum, "SodOdNo")) + _
                                   " and OdSr=" + CStr(.Value(RowNum, "SodOdSr")))
    End If
  Case Is = UCase("wSodSz")
    If .Mode = fgmnorm Or .Mode = fgminit Then
      .Value(RowNum, "wSodSz") = moCn.GetFldVal("Select OdDmSz From OrdDsg where " + _
                                   "OdCoCd= '" + .Value(RowNum, "SodOdCoCd") + "' and " + _
                                   "OdTc= '" + .Value(RowNum, "SodOdTc") + "' and " + _
                                   "OdYy= '" + .Value(RowNum, "SodOdYy") + "' and " + _
                                   "OdChr= '" + .Value(RowNum, "SodOdChr") + "' and " + _
                                   "OdNo= " + CStr(.Value(RowNum, "SodOdNo")) + _
                                   " and OdSr=" + CStr(.Value(RowNum, "SodOdSr")))
    End If
  Case Is = UCase("wSodCol")
    If .Mode = fgmnorm Or .Mode = fgminit Then
      .Value(RowNum, "wSodCol") = moCn.GetFldVal("Select OdDmCol From OrdDsg where " + _
                                   "OdCoCd= '" + .Value(RowNum, "SodOdCoCd") + "' and " + _
                                   "OdTc= '" + .Value(RowNum, "SodOdTc") + "' and " + _
                                   "OdYy= '" + .Value(RowNum, "SodOdYy") + "' and " + _
                                   "OdChr= '" + .Value(RowNum, "SodOdChr") + "' and " + _
                                   "OdNo= " + CStr(.Value(RowNum, "SodOdNo")) + _
                                   " and OdSr=" + CStr(.Value(RowNum, "SodOdSr")))
    End If
  Case Is = UCase("wSodKt")
    If .Mode = fgmnorm Or .Mode = fgminit Then
      .Value(RowNum, "wSodKt") = moCn.GetFldVal("Select OdKt From OrdDsg where " + _
                                   "OdCoCd= '" + .Value(RowNum, "SodOdCoCd") + "' and " + _
                                   "OdTc= '" + .Value(RowNum, "SodOdTc") + "' and " + _
                                   "OdYy= '" + .Value(RowNum, "SodOdYy") + "' and " + _
                                   "OdChr= '" + .Value(RowNum, "SodOdChr") + "' and " + _
                                   "OdNo= " + CStr(.Value(RowNum, "SodOdNo")) + _
                                   " and OdSr=" + CStr(.Value(RowNum, "SodOdSr")))
    End If
  Case Is = UCase("wSodOrdPcs")
    If .Mode = fgmnorm Or .Mode = fgminit Then
      .Value(RowNum, "wSodOrdPcs") = moCn.GetFldVal("Select OdOrdQty From OrdDsg where " + _
                                   "OdCoCd= '" + .Value(RowNum, "SodOdCoCd") + "' and " + _
                                   "OdTc= '" + .Value(RowNum, "SodOdTc") + "' and " + _
                                   "OdYy= '" + .Value(RowNum, "SodOdYy") + "' and " + _
                                   "OdChr= '" + .Value(RowNum, "SodOdChr") + "' and " + _
                                   "OdNo= " + CStr(.Value(RowNum, "SodOdNo")) + _
                                   " and OdSr=" + CStr(.Value(RowNum, "SodOdSr")))

    End If
  Case Is = UCase("wSodShpBal")
    If .Mode = fgmnorm Or .Mode = fgminit Then
      Dim wsOdOrdPcs As Double
      wsOdOrdPcs = moCn.GetFldVal("Select OdOrdQty From OrdDsg where " + _
                                   "OdCoCd= '" + .Value(RowNum, "SodOdCoCd") + "' and " + _
                                   "OdTc= '" + .Value(RowNum, "SodOdTc") + "' and " + _
                                   "OdYy= '" + .Value(RowNum, "SodOdYy") + "' and " + _
                                   "OdChr= '" + .Value(RowNum, "SodOdChr") + "' and " + _
                                   "OdNo= " + CStr(.Value(RowNum, "SodOdNo")) + _
                                   " and OdSr=" + CStr(.Value(RowNum, "SodOdSr")))
      Dim wsOrdShpBal As Double
      
      wsOrdShpBal = moCn.GetFldVal("Select sum(isnull(IfQty,0)) from OrdDsg " + _
                                    "left outer join Bag  on OdCoCd =BCoCd and OdTc =BOdTc and OdYy =BOdYy and OdChr =BOdChr and OdNo=BOdNo and OdSr =BOdSr " + _
                                    "left outer join InvFgd on BCoCd =IfCoCd and BYy =IfFdBYy and BChr =IfFdBChr and BNo =IfFdBNo " + _
                                    "left outer join InvHd on InCoCd=IfCoCd and InTc=IfTc and InYy=IfYy and InChr=IfChr and InNo=IfNo " + _
                                    "where OdCoCd ='" + .Value(RowNum, "SodOdCoCd") + _
                                    "' and OdTc ='" + .Value(RowNum, "SodOdTc") + _
                                    "' and OdYy ='" + .Value(RowNum, "SodOdYy") + _
                                    "' and OdChr ='" + .Value(RowNum, "SodOdChr") + _
                                    "' and OdNo =" + CStr(.Value(RowNum, "SodOdNo")) + _
                                 " and OdSr =" + CStr(.Value(RowNum, "SodOdSr")))
         .Value(RowNum, "wSodShpBal") = wsOdOrdPcs - wsOrdShpBal
                                 
    End If
  End Select
End With
End Sub

Private Sub GrdShpOrdDsg_Validate(Cancel As Boolean)
  Cancel = GrdShpOrdDsg.Validate
End Sub


Public Sub FillGrdShpInv(ByVal RowNum As Integer)
  Dim wsql As String, wo_rsInv As MwfLib.MDORowSet, wi_Row As Integer

  wsql = "Select InCoCd ,InTc ,InYy ,InChr ,InNo ,max(InDt) as InDt ,sum(IfQty) as IfQty from InvHd " + _
        "join InvFgd on InCoCd=IfCoCd and InTc=IfTc and InYy=IfYy and InChr=IfChr and InNo=IfNo " + _
        "join Bag on BCoCd =IfCoCd and BYy =IfFdBYy and BChr =IfFdBChr and BNo =IfFdBNo " + _
        "where  exists (Select 'a' from ShpNoteOrdDsg where SodOdCoCd='" + GrdShpOrdMst.Value(RowNum, "SomOmCoCd") + _
        "' and SodOdTc ='" + GrdShpOrdMst.Value(RowNum, "SomOmTc") + _
        "' and SodOdYy ='" + GrdShpOrdMst.Value(RowNum, "SomOmYy") + _
        "' and SodOdChr ='" + GrdShpOrdMst.Value(RowNum, "SomOmChr") + _
        "' and SodOdNo =" + CStr(GrdShpOrdMst.Value(RowNum, "SomOmNo")) + _
        " and SodSr =" + CStr(GrdShpOrdMst.Value(RowNum, "SomSr")) + _
        " and BOdSr =SodOdSr ) " + _
        " and BCoCd ='" + GrdShpOrdMst.Value(RowNum, "SomOmCoCd") + _
        "' and BOdTc ='" + GrdShpOrdMst.Value(RowNum, "SomOmTc") + _
        "' and BOdYy ='" + GrdShpOrdMst.Value(RowNum, "SomOmYy") + _
        "' and BOdChr ='" + GrdShpOrdMst.Value(RowNum, "SomOmChr") + _
        "' and BOdNo =" + CStr(GrdShpOrdMst.Value(RowNum, "SomOmNo")) + _
        " group by InCoCd ,InTc ,InYy ,InChr ,InNo "

  GrdShpInv.Rows = 1
  Set wo_rsInv = moCn.OpenRes(wsql)
  GrdShpInv.AllowAdd = True
  With GrdShpInv
    Do While Not (wo_rsInv.EOF Or wo_rsInv.BOF)
      .AddItem
      wi_Row = .Rows - 1
      .Value(wi_Row, "wInvSr") = wi_Row
      .Value(wi_Row, "wInTc") = Trim(wo_rsInv!InTc)
      .Value(wi_Row, "wInYy") = Trim(wo_rsInv!InYy)
      .Value(wi_Row, "wInChr") = Trim(wo_rsInv!InChr)
      .Value(wi_Row, "wInNo") = Trim(wo_rsInv!InNo)
      .Value(wi_Row, "wInDt") = Trim(wo_rsInv!InDt)
      .Value(wi_Row, "wInPcs") = Trim(wo_rsInv!IfQty)

      wo_rsInv.MoveNext
    Loop
  End With
  GrdShpInv.AllowAdd = False
  Set wo_rsInv = Nothing

End Sub

Private Sub CmdCpyDsg_Click()
  ADC("wOdSrFr") = 0
  ADC("wOdSrTo") = 0
  Call DispFra(CpyDsg)
End Sub

Private Sub DispFra(ByVal pv_ShpNoteFra As en_InvDsgFra)
  '*** This is the common procedure which is used to toggle between hiding and displaying
      'the frames (i.e. if the frame is visible then make it invisible and viseversa) associated with a command button
  Select Case pv_ShpNoteFra
  Case Is = CpyDsg
    If FraCpyDsg.Visible = True Then
      FraCpyDsg.Visible = False
      Call EnaDisaCmds(False, CmdCpyDsg, CpyDsg)
      CmdCpyDsg.SetFocus
    Else
      FraCpyDsg.Visible = True
      ADC("wOdSrFr").SetFocus
      FraCpyDsg.ZOrder
      Call EnaDisaCmds(True, CmdCpyDsg)
    End If
  End Select
End Sub

Private Sub EnaDisaCmds(ByVal pv_ShowFra As Boolean, Optional ByVal pv_CmdBut As MWCTL_BTN1, Optional ByVal pv_ShpNoteFra As en_InvDsgFra)
  '*** This procedure is used to give the toggle option (enable and disable) to the command buttons
  '*** storing the state of cmd buttons, when the frame is going to be vis.
  Dim X As Integer
  If pv_ShowFra = True Then
    Call StoreState
  End If

  If Not pv_CmdBut Is Nothing Then
    If pv_ShowFra = True Then
'        If pv_CmdBut.Pic = 0 Then
'          pv_CmdBut.BackColor = vbWhite
'        End If
        pv_CmdBut.CausesValidation = True
    Else
'        If pv_CmdBut.Pic = 0 Then
'          pv_CmdBut.BackColor = &H8000000F
'        End If
'        '*** restoring the 'Causes Validation' state of the current pv_cmdbutton
'        If Not IsMissing(pv_ShpNoteFra) Then
'          pv_CmdBut.CausesValidation = mArr_Cv(pv_ShpNoteFra)
'        End If
    End If
  End If
  FraNKeyAll.Enabled = Not pv_ShowFra
  ADC.AllowSave = Not pv_ShowFra

'  If pv_ShowFra = True Then
'    '** All command buttons are disabled
'    CmdCpyDsg.Enabled = Not pv_ShowFra    '18
'  Else
'    '*** Restoring Ena/Disa state of all Command Buttons
'    CmdCpyDsg.Enabled = mArr_Ed(CpyDsg)
'  End If
  If Not pv_CmdBut Is Nothing Then pv_CmdBut.Enabled = True
End Sub

Private Sub StoreState()
  '*** This procedure is used to store the CausesValidation State and the Enabled/Disabled
      'states of all the Command Buttons
  '*** This is used for restoring the original states in the future
  mArr_Cv(CpyDsg) = CmdCpyDsg.CausesValidation
  mArr_Ed(CpyDsg) = CmdCpyDsg.Enabled
End Sub

Private Sub CmdCpyDsgGo_Click()
  If GrdShpOrdMst.Row <= 0 Then Exit Sub
  If GrdShpOrdMst.Row > GrdShpOrdMst.Rows - 1 Then Exit Sub
  Dim wsql As String, wo_rsOrdDsg As MwfLib.MDORowSet, wi_Row As Integer
  wsql = "select OdSr ,OdDmCd ,OdSfx ,OdDmSz ,OdDmCol ,OdKt ,OdOrdQty, " + _
        "OdCoCd ,OdTc ,OdYy ,OdChr ,OdNo ,OdSr from OrdDsg " + _
        "join OrdMst on OmCoCd =OdCoCd and OmTc =OdTc and OmYy =OdYy and OmChr =OdChr and OmNo =OdNo " + _
        "where OdCoCd ='" + GrdShpOrdMst.Value(GrdShpOrdMst.Row, "SomOmCoCd") + _
        "' and OdTc ='" + GrdShpOrdMst.Value(GrdShpOrdMst.Row, "SomOmTc") + _
        "' and OdYy ='" + GrdShpOrdMst.Value(GrdShpOrdMst.Row, "SomOmYy") + _
        "' and OdChr ='" + GrdShpOrdMst.Value(GrdShpOrdMst.Row, "SomOmChr") + _
        "' and OdNo =" + CStr(GrdShpOrdMst.Value(GrdShpOrdMst.Row, "SomOmNo")) + _
        IIF(ADC("wOdSrFr") = 0 And ADC("wOdSrTo") = 0, "", " and OdSr>=" + CStr(ADC("wOdSrFr")) + _
        " and OdSr<=" + CStr(ADC("wOdSrTo"))) + " and (OdOrdQty-OdExpQty) >0 and OdCls='N' " + _
        " and not exists (select 'x' from ShpNoteOrdDsg where SodCoCd ='" + ADC("SnCoCd") + _
        "' and SodTc ='" + ADC("SnTc") + "' and SodYy ='" + ADC("SnYy") + "' and SodChr ='" + ADC("SnChr") + _
        "' and SodNo =" + CStr(ADC("SnNo")) + " and SodOdCoCd =OdCoCd and SodOdTc =OdTc and SodOdYy =OdYy " + _
        " and SodOdChr =OdChr and SodOdNo =OdNo and SodOdSr =OdSr )"
  
      
  Set wo_rsOrdDsg = moCn.OpenRes(wsql)
  GrdShpOrdDsg.AllowAdd = True
  With GrdShpOrdDsg
    Do While Not (wo_rsOrdDsg.EOF Or wo_rsOrdDsg.BOF)
      .AddItem
      wi_Row = .Rows - 1
      .Value(wi_Row, "SodSr") = GrdShpOrdMst.Value(GrdShpOrdMst.Row, "SomSr")
      .Value(wi_Row, "SodSrNo") = .MaxVal("SodSrNo") + 1
      .Value(wi_Row, "SodOdCoCd") = Trim(wo_rsOrdDsg!OdCoCd)
      .Value(wi_Row, "SodOdTc") = Trim(wo_rsOrdDsg!OdTc)
      .Value(wi_Row, "SodOdYy") = Trim(wo_rsOrdDsg!OdYy)
      .Value(wi_Row, "SodOdChr") = Trim(wo_rsOrdDsg!OdChr)
      .Value(wi_Row, "SodOdNo") = Trim(wo_rsOrdDsg!OdNo)
      .Value(wi_Row, "SodOdSr") = Trim(wo_rsOrdDsg!OdSr)
      
      .Value(wi_Row, "wSodDmCd") = Trim(wo_rsOrdDsg!OdDmCd)
      .Value(wi_Row, "wSodSfx") = Trim(wo_rsOrdDsg!OdSfx)
      .Value(wi_Row, "wSodSz") = Trim(wo_rsOrdDsg!OdDmSz)
      .Value(wi_Row, "wSodCol") = Trim(wo_rsOrdDsg!OdDmCol)
      .Value(wi_Row, "wSodKt") = Trim(wo_rsOrdDsg!OdKt)
      .Value(wi_Row, "wSodOrdPcs") = Trim(wo_rsOrdDsg!OdOrdQty)
     
      .SaveRec (wi_Row)
      wo_rsOrdDsg.MoveNext
    Loop
  End With
  Set wo_rsOrdDsg = Nothing
  Call DispFra(CpyDsg)

End Sub

Public Sub FillGrdShpInvDsg(ByVal RowNum As Integer)
  Dim wsql As String, wo_rsInvFgd As MwfLib.MDORowSet, wi_Row As Integer

  wsql = "Select IfCoCd ,IfTc ,IfYy ,IfChr ,IfNo ,IfSrNo ,InDt ,IfQty  from InvFgd " + _
        "join InvHd on InCoCd=IfCoCd and InTc=IfTc and InYy=IfYy and InChr=IfChr and InNo=IfNo " + _
        "join Bag on BCoCd =IfCoCd and BYy =IfFdBYy and BChr =IfFdBChr and BNo =IfFdBNo " + _
        "where BCoCd ='" + GrdShpOrdDsg.Value(RowNum, "SodOdCoCd") + _
        "' and BOdTc ='" + GrdShpOrdDsg.Value(RowNum, "SodOdTc") + _
        "' and BOdYy ='" + GrdShpOrdDsg.Value(RowNum, "SodOdYy") + _
        "' and BOdChr ='" + GrdShpOrdDsg.Value(RowNum, "SodOdChr") + _
        "' and BOdNo =" + CStr(GrdShpOrdDsg.Value(RowNum, "SodOdNo")) + _
        " and BOdSr =" + CStr(GrdShpOrdDsg.Value(RowNum, "SodOdSr"))

  GrdShpInvDsg.Rows = 1
  Set wo_rsInvFgd = moCn.OpenRes(wsql)
  GrdShpInvDsg.AllowAdd = True
  With GrdShpInvDsg
    Do While Not (wo_rsInvFgd.EOF Or wo_rsInvFgd.BOF)
      .AddItem
      wi_Row = .Rows - 1
      .Value(wi_Row, "wIfTc") = Trim(wo_rsInvFgd!IfTc)
      .Value(wi_Row, "wIfYy") = Trim(wo_rsInvFgd!IfYy)
      .Value(wi_Row, "wIfChr") = Trim(wo_rsInvFgd!IfChr)
      .Value(wi_Row, "wIfNo") = Trim(wo_rsInvFgd!IfNo)
      .Value(wi_Row, "wIfDt") = Trim(wo_rsInvFgd!InDt)
      .Value(wi_Row, "wIfSrNo") = Trim(wo_rsInvFgd!IfSrNo)
      .Value(wi_Row, "wIfPcs") = Trim(wo_rsInvFgd!IfQty)

      wo_rsInvFgd.MoveNext
    Loop
  End With
  GrdShpInvDsg.AllowAdd = False
  Set wo_rsInvFgd = Nothing

End Sub

Private Sub HideAllFras()
  FraNKeyAll.Visible = True
End Sub




