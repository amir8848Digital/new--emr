VERSION 5.00
Object = "{AB2CC32A-E652-42A2-B341-26ECE02F6386}#2.1#0"; "mwfctlforspeed9.ocx"
Begin VB.Form EmrFrmBagReco 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Bag Reconciliation"
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
      Height          =   3195
      Left            =   240
      TabIndex        =   40
      Top             =   6120
      Visible         =   0   'False
      Width           =   14595
      Begin MwfCtl.MWCTL_MED ATXT 
         Height          =   2205
         Index           =   40
         Left            =   120
         TabIndex        =   42
         ToolTipText     =   "Copy Bag Alias No. / Enter Bag Nos  (Yy/Chr/No,YY....,Yy/Chr/No)"
         Top             =   840
         Width           =   13695
         _ExtentX        =   24156
         _ExtentY        =   3889
         Alignment       =   3
         MaxLength       =   0
         IdName          =   "WRFIDSEL"
      End
      Begin MwfCtl.MWCTL_BTN1 CmdCpyRfIdGo 
         Height          =   525
         Left            =   13920
         TabIndex        =   43
         Top             =   2520
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
         Left            =   2610
         TabIndex        =   41
         ToolTipText     =   "Enter 'Y' to Copy Thru Bag Nos and  'N' to Copy thru Bag Alias No."
         Top             =   480
         Width           =   315
         _ExtentX        =   556
         _ExtentY        =   503
         DataType        =   4
         IdName          =   "WCPYTHRUBAGNO"
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
         Left            =   120
         TabIndex        =   48
         Top             =   480
         Width           =   2325
      End
      Begin VB.Label LblMain 
         BackStyle       =   0  'Transparent
         BorderStyle     =   1  'Fixed Single
         Caption         =   "  Copy Bag Alias No. / Bag Nos"
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
         TabIndex        =   44
         Top             =   90
         Width           =   5850
      End
   End
   Begin MwfCtl.MWCTL_BTN1 CmdCpyRfId 
      Height          =   405
      Left            =   5715
      TabIndex        =   39
      ToolTipText     =   "Copy Option"
      Top             =   9360
      Width           =   1155
      _ExtentX        =   2037
      _ExtentY        =   714
      BackColor       =   -2147483633
      MaskColor       =   8454143
      Caption         =   "&Copy Bag "
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
   Begin MwfCtl.MWCTL_BTN1 CmdReco 
      Height          =   405
      Left            =   6900
      TabIndex        =   21
      ToolTipText     =   "Click to Reconcile this Voucher"
      Top             =   9360
      Width           =   1065
      _ExtentX        =   1879
      _ExtentY        =   714
      BackColor       =   -2147483633
      MaskColor       =   8454143
      Caption         =   "&Reconcile"
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
      Left            =   3120
      TabIndex        =   23
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
      Left            =   60
      TabIndex        =   22
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
      Height          =   465
      Left            =   -120
      TabIndex        =   25
      Top             =   0
      Width           =   4260
      Begin MwfCtl.MWCTL_MED ATXT 
         Height          =   285
         Index           =   1
         Left            =   1065
         TabIndex        =   0
         ToolTipText     =   "Enter Voucher Tc"
         Top             =   120
         Width           =   570
         _ExtentX        =   1005
         _ExtentY        =   503
         Enabled         =   0   'False
         MaxLength       =   3
         DataType        =   4
         DataField       =   "BMTC"
         ReCalcParent    =   "BMCHR"
         IdName          =   "BMTC"
      End
      Begin MwfCtl.MWCTL_MED ATXT 
         Height          =   285
         Index           =   4
         Left            =   2955
         TabIndex        =   3
         ToolTipText     =   "Enter Voucher Number"
         Top             =   120
         Width           =   930
         _ExtentX        =   1640
         _ExtentY        =   503
         Alignment       =   2
         Mask            =   "######0"
         MaxLength       =   7
         DataType        =   1
         DataField       =   "BMNO"
         ReCalcParent    =   "BMDT"
         IdName          =   "BMNO"
      End
      Begin MwfCtl.MWCTL_MED ATXT 
         Height          =   285
         Index           =   6
         Left            =   1740
         TabIndex        =   1
         ToolTipText     =   "Enter Voucher Year"
         Top             =   120
         Width           =   435
         _ExtentX        =   767
         _ExtentY        =   503
         MaxLength       =   2
         DataType        =   4
         DataField       =   "BMYY"
         IdName          =   "BMYY"
      End
      Begin MwfCtl.MWCTL_MED ATXT 
         Height          =   285
         Index           =   41
         Left            =   2280
         TabIndex        =   2
         ToolTipText     =   "Enter Voucher Character"
         Top             =   120
         Width           =   570
         _ExtentX        =   1005
         _ExtentY        =   503
         MaxLength       =   3
         DataType        =   4
         DataField       =   "BMCHR"
         ReCalcOn        =   "BMTC"
         IdName          =   "BMCHR"
      End
      Begin MwfCtl.MWCTL_MED ATXT 
         Height          =   315
         Index           =   28
         Left            =   3900
         TabIndex        =   4
         Top             =   120
         Visible         =   0   'False
         Width           =   135
         _ExtentX        =   238
         _ExtentY        =   556
         Enabled         =   0   'False
         MaxLength       =   3
         DataType        =   4
         DataField       =   "BMCOCD"
         IdName          =   "BMCOCD"
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
         Left            =   120
         TabIndex        =   30
         Top             =   120
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
         Left            =   2190
         TabIndex        =   29
         Top             =   120
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
         Left            =   2865
         TabIndex        =   28
         Top             =   120
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
         Left            =   1650
         TabIndex        =   27
         Top             =   120
         Width           =   135
      End
   End
   Begin VB.Frame FRA_NKEYS 
      BorderStyle     =   0  'None
      Enabled         =   0   'False
      Height          =   9300
      Left            =   0
      TabIndex        =   24
      Top             =   0
      Width           =   15225
      Begin VB.Frame FraMsg 
         Height          =   915
         Left            =   5430
         TabIndex        =   35
         Top             =   8325
         Visible         =   0   'False
         Width           =   3825
         Begin MwfCtl.MWCTL_BTN1 CmdGo 
            Height          =   495
            Left            =   3120
            TabIndex        =   20
            ToolTipText     =   "Click Here To Copy"
            Top             =   210
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
            Caption         =   "This is an irrevokable process. "
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
            Index           =   24
            Left            =   60
            TabIndex        =   37
            Top             =   180
            Width           =   3105
         End
         Begin VB.Label ALBL 
            BackStyle       =   0  'Transparent
            Caption         =   "Do you want to continue..."
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
            Index           =   3
            Left            =   90
            TabIndex        =   36
            Top             =   450
            Width           =   3105
         End
      End
      Begin VB.Frame FraNKeyAll 
         BorderStyle     =   0  'None
         Height          =   7785
         Left            =   30
         TabIndex        =   26
         Top             =   1440
         Width           =   15165
         Begin MwfCtl.MWCTL_FLX GrdBgReco 
            Height          =   7905
            Left            =   0
            TabIndex        =   17
            Top             =   0
            Width           =   15135
            _ExtentX        =   26696
            _ExtentY        =   13944
            Cols            =   15
            colname1        =   "BDSR"
            heading1        =   "Sr         "
            datafld1        =   "BDSr"
            datatype1       =   1
            mask1           =   "###0"
            maxlength1      =   4
            style1          =   2
            tooltiptext1    =   "Voucher Serial Number"
            colname2        =   "BDBAGCOCD"
            heading2        =   "Comp Cd"
            datafld2        =   "BdBagCoCd"
            datatype2       =   4
            recalcparent2   =   "GrdBgReco(BDDMCD),GrdBgReco(BDGRWT),GrdBgReco(BDOTHLOCBG)"
            maxlength2      =   3
            tooltiptext2    =   "Enter Bag Company Code"
            colname3        =   "WBAGNO"
            datatype3       =   4
            maxlength3      =   18
            tooltiptext3    =   "Read Bag Number Using Barcode Reader"
            colname4        =   "BDBYY"
            heading4        =   "BYy  "
            datafld4        =   "BDBYy"
            recalcparent4   =   "GrdBgReco(BDACTUALQTY),GrdBgReco(BDDMCD),GrdBgReco(BDGRWT),GrdBgReco(BDOTHLOCBG)"
            maxlength4      =   2
            style4          =   2
            tooltiptext4    =   "Bag Year (Yy)"
            colname5        =   "BDBCHR"
            heading5        =   "BChr     "
            datafld5        =   "BDBChr"
            datatype5       =   4
            recalcparent5   =   "GrdBgReco(BDACTUALQTY),GrdBgReco(BDDMCD),GrdBgReco(BDGRWT),GrdBgReco(BDOTHLOCBG)"
            maxlength5      =   3
            style5          =   2
            tooltiptext5    =   "Bag Chr"
            colname6        =   "BDBNO"
            heading6        =   "BNo       "
            datafld6        =   "BDBNo"
            datatype6       =   1
            mask6           =   "######0"
            recalcparent6   =   "GrdBgReco(BDACTUALQTY),GrdBgReco(BDDMCD),GrdBgReco(BDGRWT),GrdBgReco(BDOTHLOCBG)"
            maxlength6      =   7
            style6          =   2
            tooltiptext6    =   "Bag Number"
            colname7        =   "BDACTUALQTY"
            heading7        =   "Rept Qty  "
            datafld7        =   "BDActualQty"
            datatype7       =   2
            mask7           =   "##0.0"
            recalcparent7   =   "WTOTBAGS,WTOTACTQTY"
            maxlength7      =   5
            tooltiptext7    =   "Enter Reported (Physical) Quantity"
            colname8        =   "BDBQTY"
            heading8        =   "Bag Qty"
            datafld8        =   "BDBQty"
            datatype8       =   2
            mask8           =   "##0.0"
            maxlength8      =   5
            style8          =   0
            tooltiptext8    =   "Bag Quantity in WIP"
            colname9        =   "BDDMCD"
            heading9        =   "Design Cd                   "
            datafld9        =   "BDDmCd"
            datatype9       =   4
            recalcon9       =   "GrdBgReco(BDBYY),GrdBgReco(BDBCHR),GrdBgReco(BDBNO),GrdBgReco(BDBAGCOCD)"
            style9          =   2
            tooltiptext9    =   "Bag Design Code"
            colname10       =   "BDGRWT"
            heading10       =   "Grs Wt         "
            datafld10       =   "BDGrWt"
            datatype10      =   2
            mask10          =   "-##0.000"
            recalcon10      =   "GrdBgReco(BDBYY),GrdBgReco(BDBCHR),GrdBgReco(BDBNO),GrdBgReco(BDBAGCOCD)"
            maxlength10     =   8
            style10         =   0
            tooltiptext10   =   "Bag Gross Wt"
            colname11       =   "BDEXTQTY"
            heading11       =   "Extra        "
            datafld11       =   "BDExtQty"
            datatype11      =   2
            mask11          =   "###0.0"
            maxlength11     =   6
            style11         =   0
            tooltiptext11   =   "Bag Extra Quantity"
            colname12       =   "BDUNACQTY"
            heading12       =   "Un A/c     "
            datafld12       =   "BDUnAcQty"
            datatype12      =   2
            mask12          =   "###0.0"
            maxlength12     =   6
            style12         =   0
            tooltiptext12   =   "Bag Un-Accounted Quantity"
            colname13       =   "BDOTHLOCBG"
            heading13       =   "Oth Loc  "
            datafld13       =   "BDOthLocBg"
            datatype13      =   4
            recalcon13      =   "GrdBgReco(BDBYY),GrdBgReco(BDBCHR),GrdBgReco(BDBNO),GrdBgReco(BDBAGCOCD)"
            maxlength13     =   1
            style13         =   0
            tooltiptext13   =   "Specify whether Bag is Misplaced/Belongs to Other Location"
            colname14       =   "BDOTHLOC"
            heading14       =   "Loc             "
            datafld14       =   "BdOthLoc"
            datatype14      =   4
            maxlength14     =   5
            style14         =   0
         End
         Begin MwfCtl.MWCTL_MED ATXT 
            Height          =   285
            Index           =   7
            Left            =   9975
            TabIndex        =   19
            ToolTipText     =   "Enter Bag Location"
            Top             =   8340
            Width           =   900
            _ExtentX        =   1588
            _ExtentY        =   503
            Alignment       =   2
            Enabled         =   0   'False
            Mask            =   "###0.0"
            MaxLength       =   6
            DataType        =   1
            ReCalcOn        =   "GrdBgReco(BDACTUALQTY)"
            IdName          =   "WTOTACTQTY"
         End
         Begin MwfCtl.MWCTL_MED ATXT 
            Height          =   285
            Index           =   8
            Left            =   1260
            TabIndex        =   18
            Top             =   8340
            Width           =   900
            _ExtentX        =   1588
            _ExtentY        =   503
            Alignment       =   2
            Enabled         =   0   'False
            Mask            =   "###0"
            MaxLength       =   4
            DataType        =   1
            ReCalcOn        =   "GrdBgReco(BDACTUALQTY)"
            IdName          =   "WTOTBAGS"
         End
         Begin VB.Label lbltotactqty 
            BackStyle       =   0  'Transparent
            Caption         =   "Reported Qty"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Left            =   8520
            TabIndex        =   33
            Top             =   8340
            Width           =   1440
         End
         Begin VB.Label LblTotBags 
            BackStyle       =   0  'Transparent
            Caption         =   "Total Bags "
            BeginProperty Font 
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
            TabIndex        =   32
            Top             =   8340
            Width           =   1200
         End
      End
      Begin MwfCtl.MWCTL_MED ATXT 
         Height          =   285
         Index           =   30
         Left            =   6600
         TabIndex        =   5
         ToolTipText     =   "Enter Voucher Date"
         Top             =   150
         Width           =   1020
         _ExtentX        =   1799
         _ExtentY        =   503
         MaxLength       =   8
         DataType        =   3
         DataField       =   "BMDT"
         ReCalcOn        =   "BMNO"
         IdName          =   "BMDT"
      End
      Begin MwfCtl.MWCTL_MED ATXT 
         Height          =   285
         Index           =   10
         Left            =   12555
         TabIndex        =   16
         Top             =   150
         Width           =   420
         _ExtentX        =   741
         _ExtentY        =   503
         Enabled         =   0   'False
         DataType        =   4
         DataField       =   "BMRECOYN"
         ReCalcOn        =   "BMNO"
         IdName          =   "BMRECOYN"
      End
      Begin MwfCtl.MWCTL_MED ATXT 
         Height          =   285
         Index           =   3
         Left            =   14700
         TabIndex        =   6
         ToolTipText     =   "Enter 'Y' to Calculate Reported (Physical) Quantity"
         Top             =   150
         Width           =   345
         _ExtentX        =   609
         _ExtentY        =   503
         DataType        =   4
         DataField       =   "BmCalcActQtyYN"
         IdName          =   "BMCALCACTQTYYN"
      End
      Begin MwfCtl.MWCTL_MED ATXT 
         Height          =   285
         Index           =   0
         Left            =   1605
         TabIndex        =   13
         ToolTipText     =   "Enter FG Sub Location From"
         Top             =   1050
         Width           =   1410
         _ExtentX        =   2487
         _ExtentY        =   503
         MaxLength       =   8
         DataType        =   4
         DataField       =   "BmFgSubLocFr"
         IdName          =   "BMFGSUBLOCFR"
      End
      Begin MwfCtl.MWCTL_MED ATXT 
         Height          =   285
         Index           =   11
         Left            =   1605
         TabIndex        =   10
         ToolTipText     =   "Enter Bag Location From"
         Top             =   765
         Width           =   1410
         _ExtentX        =   2487
         _ExtentY        =   503
         MaxLength       =   8
         DataType        =   4
         DataField       =   "BmBagLocFr"
         IdName          =   "BMBAGLOCFR"
      End
      Begin MwfCtl.MWCTL_MED ATXT 
         Height          =   285
         Index           =   12
         Left            =   3840
         TabIndex        =   11
         ToolTipText     =   "Enter Bag Location To"
         Top             =   765
         Width           =   1410
         _ExtentX        =   2487
         _ExtentY        =   503
         MaxLength       =   8
         DataType        =   4
         DataField       =   "BmBagLocTo"
         IdName          =   "BMBAGLOCTO"
      End
      Begin MwfCtl.MWCTL_MED ATXT 
         Height          =   285
         Index           =   13
         Left            =   6600
         TabIndex        =   12
         ToolTipText     =   "Enter Bag Location"
         Top             =   765
         Width           =   8610
         _ExtentX        =   15187
         _ExtentY        =   503
         MaxLength       =   255
         DataType        =   4
         DataField       =   "BmBagLocSel"
         IdName          =   "BMBAGLOCSEL"
      End
      Begin MwfCtl.MWCTL_MED ATXT 
         Height          =   285
         Index           =   14
         Left            =   1605
         TabIndex        =   7
         ToolTipText     =   "Enter Bag Company Code From"
         Top             =   480
         Width           =   1410
         _ExtentX        =   2487
         _ExtentY        =   503
         MaxLength       =   3
         DataType        =   4
         DataField       =   "BmBagCoCdFr"
         IdName          =   "BMBAGCOCDFR"
      End
      Begin MwfCtl.MWCTL_MED ATXT 
         Height          =   285
         Index           =   15
         Left            =   3840
         TabIndex        =   8
         ToolTipText     =   "Enter Bag Company Code To"
         Top             =   480
         Width           =   1410
         _ExtentX        =   2487
         _ExtentY        =   503
         MaxLength       =   3
         DataType        =   4
         DataField       =   "BmBagCoCdTo"
         IdName          =   "BMBAGCOCDTO"
      End
      Begin MwfCtl.MWCTL_MED ATXT 
         Height          =   285
         Index           =   16
         Left            =   6600
         TabIndex        =   9
         ToolTipText     =   "Enter Bag Company Codes"
         Top             =   480
         Width           =   8610
         _ExtentX        =   15187
         _ExtentY        =   503
         MaxLength       =   255
         DataType        =   4
         DataField       =   "BmBagCoCdSel"
         IdName          =   "BMBAGCOCDSEL"
      End
      Begin MwfCtl.MWCTL_MED ATXT 
         Height          =   285
         Index           =   5
         Left            =   3840
         TabIndex        =   14
         ToolTipText     =   "Enter FG Sub Location To"
         Top             =   1050
         Width           =   1410
         _ExtentX        =   2487
         _ExtentY        =   503
         MaxLength       =   8
         DataType        =   4
         DataField       =   "BmFgSubLocTo"
         IdName          =   "BMFGSUBLOCTO"
      End
      Begin MwfCtl.MWCTL_MED ATXT 
         Height          =   285
         Index           =   9
         Left            =   6600
         TabIndex        =   15
         ToolTipText     =   "Enter FG Sub Locations"
         Top             =   1050
         Width           =   8610
         _ExtentX        =   15187
         _ExtentY        =   503
         MaxLength       =   255
         DataType        =   4
         DataField       =   "BmFgSubLocSel"
         IdName          =   "BMFGSUBLOCSEL"
      End
      Begin VB.Label Label11 
         BackStyle       =   0  'Transparent
         Caption         =   "Selection"
         BeginProperty Font 
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
         TabIndex        =   54
         Top             =   480
         Width           =   1095
      End
      Begin VB.Label Label10 
         BackStyle       =   0  'Transparent
         Caption         =   " To"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   3360
         TabIndex        =   53
         Top             =   480
         Width           =   495
      End
      Begin VB.Label Label9 
         BackStyle       =   0  'Transparent
         Caption         =   "Comp Code Fr"
         BeginProperty Font 
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
         TabIndex        =   52
         Top             =   480
         Width           =   1470
      End
      Begin VB.Label Label8 
         BackStyle       =   0  'Transparent
         Caption         =   "Selection"
         BeginProperty Font 
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
         TabIndex        =   51
         Top             =   765
         Width           =   1095
      End
      Begin VB.Label Label7 
         BackStyle       =   0  'Transparent
         Caption         =   " To"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   3360
         TabIndex        =   50
         Top             =   765
         Width           =   495
      End
      Begin VB.Label Label6 
         BackStyle       =   0  'Transparent
         Caption         =   "Bag Loc Cd Fr"
         BeginProperty Font 
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
         TabIndex        =   49
         Top             =   765
         Width           =   1470
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Fg Sub Loc Fr"
         BeginProperty Font 
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
         TabIndex        =   45
         Top             =   1050
         Width           =   1470
      End
      Begin VB.Label Label4 
         BackStyle       =   0  'Transparent
         Caption         =   " To"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   3360
         TabIndex        =   46
         Top             =   1050
         Width           =   495
      End
      Begin VB.Label Label5 
         BackStyle       =   0  'Transparent
         Caption         =   "Selection"
         BeginProperty Font 
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
         TabIndex        =   47
         Top             =   1050
         Width           =   1095
      End
      Begin VB.Label ALBL 
         BackStyle       =   0  'Transparent
         Caption         =   "Calc Rept Qty"
         BeginProperty Font 
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
         Left            =   13290
         TabIndex        =   38
         Top             =   150
         Width           =   1365
      End
      Begin VB.Label ALBL 
         BackStyle       =   0  'Transparent
         Caption         =   "Reco"
         BeginProperty Font 
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
         Left            =   11940
         TabIndex        =   34
         Top             =   150
         Width           =   615
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
         Left            =   5520
         TabIndex        =   31
         Top             =   150
         Width           =   855
      End
   End
End
Attribute VB_Name = "EmrFrmBagReco"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

'---- General Comments -------
' Entry for Bag Reconciliation
'******Sachin 2.11.0********
'-----------------------------

Option Explicit
Dim moCn As MwfLib.MDOConnection   '*** (Bef speed) bef 24/11/06
Dim ms_TcTyp As String, ms_wBrCoCd As String
Dim ms_wBrTc As String, ms_wBrYy As String, ms_wBrChr As String
Dim pb_CpyRfId As Boolean
Dim mb_AddRow As Boolean      ' ***** Manali 3.10.0 - 09/05/12 - Bar Code (Ctrl+N)

Private Sub adc_AutoGen()

  'Autogen Proc
  If ADC("BmNo") = 0 Then
    ADC("BmNo") = moCn.AutoGen("BgRecoMst", "BmNo", ADC("BmCoCd") + ADC("BmTc") + ADC("BmYy") + ADC("BmChr"))
  Else
    Call moCn.AutoGen("BgRecoMst", "BmNo", ADC("BmCoCd") + ADC("BmTc") + ADC("BmYy") + ADC("BmChr"), ADC("BmNo"))
  End If
  
End Sub

Private Sub ADC_Load()
    '*** Set the First Key Control property to the field wBrYy
    '*** Set the First Non Key Control property to wBrLoc
    '*** Set the Child Property Of ADC to GrdBgReco
    '*** Set the Previous and Next Control for the Grid GrdBgReco as wBrLoc and ADC respectively
    '*** Set the Hot Keys for the Grid GrdInvFgd as 'A'
  
    Set moCn = ADC.Connection    '*** (Bef speed) bef 24/11/06
    Set ADC.FirKeyCtl = ADC("BmYy")
    Set ADC.FirNKeyCtl = ADC("BmDt")
    
    ADC.Child = "GrdBgReco"
    
    Set GrdBgReco.PrevCtl = ADC("BmDt")
    Set GrdBgReco.NextCtl = ADC
    GrdBgReco.HotKey = "A"
    FraNKeyAll.ZOrder (0)
     
    ms_wBrCoCd = gs_CoCd
    ms_TcTyp = Replace(ADC.MenuCd, "SUPP", "")
    ms_wBrTc = moCn.GetFldVal("Select PValue from Param where PTyp= 'TCTYP' and " + _
                 "PMCd= '" + ms_TcTyp + "' ")
    ms_wBrYy = moCn.GetFldVal("Select hYy from Head where HCoCd= '" + ms_wBrCoCd + "' and HCd= '" + ctSelfCmCd + "'")
    
    '****** Sachin 2.14.0 - [24-11-2006] - [Multi-Company] ******
    Me.Caption = GetFrmCaption(gs_CoCd, "Bag Reconciliation")
    '****** Sachin 2.14.0 - [24-11-2006] - [Multi-Company] ******
  
    mb_AddRow = True    ' ***** Manali 3.10.0 - 09/05/12 - Bar Code (Ctrl+N)
End Sub

Private Sub ADC_KeyWhen()
   '*** Setting the default values for BrCoCd, BrTc, BrYy and BrChr
   '*** Hide all frames except Frame FraNKeyAll
  
   Call EnaDisaCmds(True) '###
   Call HideAllFras
      
   If ms_wBrCoCd <> "" Then ADC("BmCoCd") = ms_wBrCoCd
   If ms_wBrTc <> "" Then ADC("BmTc") = ms_wBrTc
   If ms_wBrYy <> "" Then ADC("BmYy") = ms_wBrYy
   If ms_wBrChr <> "" Then ADC("BmChr") = ms_wBrChr
  
End Sub

Private Sub ADC_KeyValid(Cancel As Boolean, ErrMsg As String)
    
  If ADC("BmYy") = "" Then Cancel = True: ErrMsg = "Enter Voucher Year": Exit Sub
  If ADC("BmChr") = "" Then Cancel = True: ErrMsg = "Enter Voucher Character": Exit Sub
  
  '  If Not moCn.RecSeek(" Select 'x' from BgReco " + _
  '                         " Where BrCoCd = '" + ADC("wBrCoCd") + "'" + _
  '                         " And BrTc ='" + ADC("wBrTc") + "'" + _
  '                         " And BrYy = '" + ADC("wBrYy") + "'" + _
  '                         " And BrChr ='" + ADC("wBrChr") + "' " + _
  '                         " And BrNo =" + CStr(ADC("wBrNo"))) Then
  '       ws_addchk = True
  '  End If
  'If ADC("wBrNo") = 0 Then Cancel = True: Exit Sub
  
  
End Sub

Private Sub ADC_NonKeyValid(pr_Cancel As Boolean, pr_ErrMsg As String)

  'If moCn.RecSeek("Select 'x' From BgReco Where " + _
  '                " BrCoCd = '" + ADC("wBrCoCd") + "' And BrTc= '" + ADC("wBrTc") + "' And" + _
  '                                " BrYy = '" + ADC("wBrYy") + "' And BrChr= '" + ADC("wBrChr") + "' And BrNo= " + CStr(ADC("wBrNo")) + " And BrActualQty<0") Then
  '    pr_Cancel = True: pr_ErrMsg = "Re-Conciliation not completed, Cannot save": Exit Sub
  '
  'End If

End Sub

Private Sub ADC_NonKeyWhen(Cancel As Boolean, ErrMsg As String)
       If ADC.IsNew And Not moCn.RecSeek("select vPSCd from vParam where " + _
             " vPCoCd='" + ADC("BmCoCd") + "' and vPTyp= 'CHR' and " + _
             " vPMCd = '" + ADC("BmTc") + "' and vPSCd='" + ADC("BmChr") + "' and vPValidYn<>'N'") Then
       Cancel = True: ErrMsg = "Voucher Character No More Valid": ADC.AllowSave = False: Exit Sub
      Else
        ADC.AllowSave = True
      End If
  
  Dim ws_Loc As String
    ms_wBrTc = ADC("BmTc")
    ms_wBrYy = ADC("BmYy")
    ms_wBrChr = ADC("BmChr")
  
    'Manoj 2.11.0 Corrections
    If ADC.AddFndMode = xaddmode Then
      ws_Loc = moCn.GetFldVal("Select vPValue1 from vParam where " + _
                                     "vPCoCd= '" + ADC("BmCoCd") + "' and vPTyp='CHR' " + _
                                     "and vPMCd= '" + ADC("BmTc") + "' and vPSCd='" + ADC("BmChr") + "'")
      If ws_Loc <> "" Then
        ADC("BmBagLocFr") = ws_Loc
        ADC("BmBagLocTo") = ws_Loc
      End If
      ADC("BmBagCoCdFr") = gs_CoCd
      ADC("BmBagCoCdTo") = gs_CoCd
    End If
    'Manoj 2.11.0 Corrections
    
    If UCase(ADC("BmRecoYN")) = "Y" Then
      GrdBgReco.ColProp("BdGrWt").Style = fgcReadOnly
      GrdBgReco.ColProp("BdBQty").Style = fgcReadOnly
      GrdBgReco.ColProp("BdExtQty").Style = fgcReadOnly
      GrdBgReco.ColProp("BdUnAcQty").Style = fgcReadOnly
      GrdBgReco.ColProp("BdOthLoc").Style = fgcReadOnly
      GrdBgReco.AllowAdd = False
      'Mw.9999 Err.44- A Grid records can be delete after reconciliation
      GrdBgReco.AllowDelete = True
    Else
      GrdBgReco.ColProp("BdGrWt").Style = fgcinvisible
      GrdBgReco.ColProp("BdBQty").Style = fgcinvisible
      GrdBgReco.ColProp("BdExtQty").Style = fgcinvisible
      GrdBgReco.ColProp("BdUnAcQty").Style = fgcinvisible
      GrdBgReco.ColProp("BdOthLoc").Style = fgcinvisible
      GrdBgReco.AllowAdd = True
      GrdBgReco.AllowDelete = True
    End If
  
  ' Set The Visible, Enable Properties For All The command buttons ***
    Call EnaDisaCmds(False) '###
    
    If ADC.AddFndMode = xaddmode Then

         ADC("BmDt") = moCn.SrvrDate
        ADC("BmRecoYN") = "N"
        ADC("BmCalcActQtyYN") = "N"
        
    End If

    mb_AddRow = IIF(ADC("BmCalcActQtyYN") = "Y", True, False)     ' ***** Manali 3.10.0 - 09/05/12 - Bar Code (Ctrl+N)

End Sub

Private Sub ATXT_FldWhen(Index As Integer, ByVal IdName As String, Cancel As Boolean, ErrMsg As String)
  
  '*** Set Helps For Each Field
  '*** Do not allow the user to edit the voucher date if record exists in the grid
  Select Case IdName
  Case Is = UCase("BmTc")
    Call HlpList.PMCd("TC", "'RC'")
    
  Case Is = UCase("BmYy")
    Call HlpList.PMCd("YY")
    
  Case Is = UCase("BmChr")
    Call HlpList.vPSCd(ADC("BmCoCd"), "CHR", ADC("BmTc"), "Y")
    
  Case Is = UCase("BmNo")
    Call HlpList.BrNo(ADC("BmCoCD"), ADC("BmTc"), ADC("BmYy"), ADC("BmChr"))
    
  '**********Bhavna added co and bag loc scopes
  Case Is = UCase("BmBagCoCdFr"), UCase("BmBagCoCdTo"), UCase("BmCoCdTo"), UCase("BmBagCoCdSel")
    If moCn.RecSeek("Select 'x' From BgRecoDet Where BdCoCd='" + ADC("BmCoCd") + "' And BdTc='" + ADC("BmTc") + "' " + _
          "  And BdYy= '" + ADC("BmYy") + "' And BdChr='" + ADC("BmChr") + "' And BdNo= " + CStr(ADC("BmNo"))) Then
        Cancel = True: ErrMsg = "Cannot Edit Company Codes, Child Record Exists": Exit Sub
    Else
        HlpList.MultiSelect = IIF(UCase(IdName) = UCase("BmBagCoCdSel"), True, False)
        Call HlpList.hCoCd
    End If
  
  Case Is = UCase("BmBagLocFr"), UCase("BmBagLocTo"), UCase("BmBagLocSel")
    If moCn.RecSeek("Select 'x' From BgRecoDet Where BdCoCd='" + ADC("BmCoCd") + "' And BdTc='" + ADC("BmTc") + "' " + _
          "  And BdYy= '" + ADC("BmYy") + "' And BdChr='" + ADC("BmChr") + "' And BdNo= " + CStr(ADC("BmNo"))) Then
        Cancel = True: ErrMsg = "Cannot Edit Location, Child Record Exists": Exit Sub
    Else
        HlpList.MultiSelect = IIF(UCase(IdName) = UCase("BmBagLocSel"), True, False)
        Call HlpList.Loc("", "'P'")
    End If
  '**********Bhavna added co and bag loc scopes
'        Call HlpList.Loc(ADC("BmCoCd"), "'P'")
  
  Case Is = UCase("BmDt")
    If moCn.RecSeek("Select 'x' From BgRecoDet Where BdCoCd='" + ADC("BmCoCd") + "' And BdTc='" + ADC("BmTc") + "' " + _
          "  And BdYy= '" + ADC("BmYy") + "' And BdChr='" + ADC("BmChr") + "' And BdNo= " + CStr(ADC("BmNo"))) Then
        Cancel = True: ErrMsg = "Cannot Edit Date, Child Record Exists": Exit Sub
    End If
  
  Case Is = UCase("BmCalcActQtyYN")
    If GrdBgReco.Rows > 1 Then Cancel = True: ErrMsg = "Cannot Edit Calc Act Qty, Child Record Exists": Exit Sub
    Call HlpList.PMCd("YN")
    
  Case Is = UCase("BmFgSubLocFr"), UCase("BmFgSubLocTo"), UCase("BmFgSubLocSel")
    If (ADC("BmBagLocFr") = "PFG" And ADC("BmBagLocTo") = "PFG" And ADC("BmBagLocSel") = "") Or _
      (ADC("BmBagLocFr") = "" And ADC("BmBagLocTo") = "" And Trim(ADC("BmBagLocSel")) = "PFG") Then
        Cancel = moCn.RecSeek("Select 'x' from BgRecoDet Where  BdCoCd= '" + ADC("BmCoCd") + "' And " + _
                "BdTc= '" + ADC("BmTc") + "' and BdYy= '" + ADC("BmYy") + "' and BdChr= '" + ADC("BmChr") + "' and BdNo= " + CStr(ADC("BmNo")))
        If Cancel = True Then ErrMsg = "Cannot change when Child records exist": Exit Sub
        
            HlpList.MultiSelect = IIF(UCase(IdName) = UCase("BmFgSubLocSel"), True, False)
            Call HlpList.vPMCd(ADC("BmCoCd"), "FgSubLoc")
    
    Else
      Cancel = True: ErrMsg = "Cannot Enter Fg Sub Location": Exit Sub
    End If
    
       '*********Bhavna cpy thru bag no
    Case Is = UCase("wCPYTHRUBAGNO")
        If ADC("WCPYTHRUBAGNO") = "" Then
            ADC("WCPYTHRUBAGNO") = "N"
        End If
  
  End Select
  
End Sub

Private Sub atxt_ReCalculate(Index As Integer, ByVal IdName As String, ByVal OldVal As Variant, ByVal RowNum As Integer, ByVal SourceName As String, ByVal SourceColName As String)
  
  Dim i As Integer
  Dim wTotBQty  As Double, wTotActQty  As Double
  
  Select Case UCase(IdName)
  Case Is = UCase("BmChr")
    If ADC.Mode = xNorm Then
       ADC("BmChr") = moCn.GetFldVal("Select vPValue from vParam Where " + _
                                "vPCoCd = '" + ADC("BmCoCd") + "' And vPTyp= 'DEFCHR' " + _
                                "and vPMCd= '" + ADC("BmTc") + "'")
    End If
      
  'Calculating the Total Actual Quantity
  Case Is = UCase("wTotActQty")
    If ADC.Mode = xNorm Or ADC.Mode = xInit Then
      If GrdBgReco.Mode = fgmnorm Or GrdBgReco.Mode = fgmdel Then
        wTotActQty = 0
        For i = 1 To (GrdBgReco.Rows - 1)
          If Not GrdBgReco.IsDel(i) Then wTotActQty = wTotActQty + GrdBgReco.Value(i, "BdActualQty")
        Next i
        ADC("wTotActQty") = wTotActQty
      ElseIf GrdBgReco.Mode = fgminit Then
        wTotActQty = moCn.GetFldVal("Select Sum(BdActualQty) from BgRecoDet " + _
                       " Where BdCoCd = '" + ADC("BmCoCd") + "'" + _
                       " And BdTc ='" + ADC("BmTc") + "'" + _
                       " And BdYy = '" + ADC("BmYy") + "'" + _
                       " And BdChr ='" + ADC("BmChr") + "' " + _
                       " And BdNo =" + CStr(ADC("BmNo")))
        ADC("wTotActQty") = wTotActQty
      End If
    End If
    
  'Calculating the Total Bags in the Voucher
  Case Is = UCase("wTotBags")
    Dim wTotBags As Integer
    If GrdBgReco.Mode = fgmnorm Or GrdBgReco.Mode = fgmdel Then
        wTotBags = 0
        For i = 1 To (GrdBgReco.Rows - 1)
          If Not GrdBgReco.IsDel(i) Then
            wTotBags = wTotBags + 1
          End If
        Next i
       ADC("wTotBags") = wTotBags
    ElseIf GrdBgReco.Mode = fgminit Then
        wTotBags = moCn.GetFldVal("Select Count(*) From BgRecoDet Where BdCoCd = '" + ADC("BmCoCd") + "' And " + _
                           " BdTc = '" + ADC("BmTc") + "' And + BdYy = '" + ADC("BmYy") + "' And BdChr = '" + ADC("BmChr") + "' And BdNo = " + CStr(ADC("BmNo")))
        ADC("wTotBags") = wTotBags
    End If
'   Case Is = UCase("BmLoc")
'         If GrdBgReco.Mode = fgmnorm Then
'         If adc("BmFgSubLoc") = "" Then
'            adc("BmFgSubLoc") = moCn.GetFldVal("Select hFgSubLoc from Head where HCoCd= '" + _
'                                    adc("BmCoCd") + "' and HCd='" + ctSelfCmCd + "'")
'         Else
'            adc("BmFgSubLoc") = ""
'         End If
'         End If
  End Select
  
End Sub

Private Sub ATXT_FldChange(Index As Integer, ByVal IdName As String, Cancel As Boolean, ErrMsg As String, ByVal pv_NewValue As Variant)
  
  Select Case UCase(IdName)
  Case Is = UCase("BmTc")
    'Check for Invalid Voucher Tc
    Cancel = Not moCn.RecSeek("select PMCd from Param where PTyp= 'TC' and " + _
                         " PMCd = '" + pv_NewValue + "' and " + _
                         " PValue= '" + ms_TcTyp + "' and " + _
                         " (Select Pvalue1 from Param where PTyp='TCTYP' and " + _
                         " PMCd= '" + ms_TcTyp + "') ='RC'")
    If Cancel = True Then ErrMsg = "Invalid Voucher Tc": Exit Sub
    
  Case Is = UCase("BmYy")
    'Check for Invalid Voucher Yy
    Cancel = Not moCn.RecSeek(" select PMCd from Param where PTyp= 'YY' and " + _
                         " PMCd = '" + pv_NewValue + "'")
    If Cancel = True Then ErrMsg = "Invalid Voucher Year": Exit Sub
    
  Case Is = UCase("BmChr")
    'Check for Invalid Voucher Chr
    Cancel = (pv_NewValue <> "") And ((Not moCn.RecSeek("select vPSCd from vParam where " + _
             " vPCoCd='" + ADC("BmCoCd") + "' and vPTyp= 'CHR' and " + _
             " vPMCd = '" + ADC("BmTc") + "' and vPSCd='" + pv_NewValue + "'")))
    
    If Cancel = True Then ErrMsg = "Invalid Voucher Character": Exit Sub
    Dim ws_UsrArr() As String, wi_Index As Integer
    ws_UsrArr = Split(moCn.GetFldVal("Select vPValue4 from vParam where vPCoCd= '" + ADC("BmCoCd") + "' and " + _
                                        "vPTyp='CHR' and vPMCd='" + ADC("BmTc") + "' and vPSCd= '" + pv_NewValue + "'"), ",")
    If IsArray(ws_UsrArr) Then
      For wi_Index = 0 To UBound(ws_UsrArr)
        Cancel = IIF(UCase(gs_UsrCd) = UCase(ws_UsrArr(wi_Index)), False, True)
        If Cancel = False Then: Exit For
      Next wi_Index
    End If
   If Cancel = True Then ErrMsg = "Access To This Voucher Character Is Denied To User " + gs_UsrCd: Exit Sub
      
'  Case Is = UCase("BmLoc")
'        'Check for Blank Process Location
'        If pv_NewValue = "" Then
'            Cancel = True: ErrMsg = "Voucher Location Cannot be Blank": Exit Sub
'
'        ElseIf pv_NewValue = moCn.GetFldVal("Select LocCd From Loc Where LocCoCd='" + adc("BmCoCd") + "' And LocValidYN='N'") Then
'            Cancel = True: ErrMsg = "Invalid Voucher Location, Check Location Master": Exit Sub
'    Cancel = Not moCn.RecSeek("Select 'x' From Loc Where LocCoCd='" + ADC("BmCoCd") + "' and LocTyp= 'P' and LocCd='" + pv_NewValue + "' And LocValidYN='Y' ")
'    If Cancel = True Then ErrMsg = "Invalid Voucher Location": Exit Sub
'        'Location and Voucher Date combination cannot be repeated
'        ElseIf moCn.RecSeek("Select 'x' From BgRecoMst Where BmCoCd='" + ADC("BmCoCd") + "' " + _
'                            " And BmTc= '" + ADC("BmTc") + "' " + _
'                            " And BmYy='" + ADC("BmYy") + "' " + _
'                            " And BmLoc='" + pv_NewValue + "' And BmDt = '" + Format((ADC("BmDt")), "DD/MM/YY") + "'") = True Then
'            Cancel = True: ErrMsg = "Voucher for the Date and Location already exists ": Exit Sub
'        End If
 Case Is = UCase("BmFgSubLoc")
    If (ADC("BmBagLocFr") = ADC("BmBagLocTo") = "PFG" And ADC("BmBagLocSel") = "") Or (ADC("BmBagLocFr") = ADC("BmBagLocTo") = "" And Trim(ADC("BmBagLocSel")) = "PFG") Then
      Cancel = Not moCn.RecSeek("Select vPmCd from vParam where " + _
              "vPCoCd = '" + ADC("BmCoCd") + "' and " + _
              "vPTyp = 'FGSUBLOC' and vPMcd = '" + pv_NewValue + "'")
      If Cancel = True Then ErrMsg = "Invalid Fg Sub Location": Exit Sub
    Else
      If pv_NewValue <> "" Then Cancel = True: ErrMsg = "Fg SubLoc Should be Blank": Exit Sub
    End If
  Case Is = UCase("BmDt")
    'Check for Invalid Voucher Date
    If ADC.Mode = xNorm Then
        Cancel = pv_NewValue = ""
        If Cancel = False Then Cancel = CDate(pv_NewValue) > moCn.SrvrDate
        If Cancel = True Then ErrMsg = "Invalid Voucher Date": Exit Sub
    End If
    
 Case Is = UCase("BmRecoYN")
    'Check for Invalid Reco Option
    If UCase(pv_NewValue) = "Y" Then
        GrdBgReco.AllowAdd = False
        'Mw.9999 Err.44- A Grid records can be delete after reconciliation
        GrdBgReco.AllowDelete = True
    Else
        GrdBgReco.AllowAdd = True
        GrdBgReco.AllowDelete = True
    End If
    
  Case Is = UCase("BmCalcActQtyYN")
    Cancel = Not moCn.RecSeek("Select 'x' From Param Where PTyp='YN' and PMCd='" + pv_NewValue + "' ")
    If Cancel = True Then ErrMsg = "Enter 'Y'/'N'": Exit Sub
    mb_AddRow = IIF(pv_NewValue = "Y", True, False)
    
    '**** Bhavna added cpy thru bag no
  Case Is = UCase("wCpyThruBagNo")  '6.1b-22 rfid changed as Bag Alias No
    If ADC.Mode = xNorm Then
      Cancel = Not moCn.RecSeek("select PMCd from Param where PTyp= 'YN' and PMCd = '" + pv_NewValue + "'")
      ErrMsg = "Enter 'Y' to Copy thru Bag No Or 'N' to Copy thru Bag Alias No. ": Exit Sub
    End If
  
  End Select
  
End Sub

Private Sub CmdCpyRfId_Click()
  
  If Not ADC.SaveRec Then
    DispMsg "X Save Failed", etInfo
    Exit Sub
  End If
  
     If FraCpyRfId.Visible = True Then
      FraCpyRfId.Visible = False
      FraCpyRfId.Enabled = False
      Call EnaDisaCmds(False, CmdCpyRfId)
      CmdCpyRfId.SetFocus
    Else
    '*******Bhavna added co scope  ..rfid only if one comp
      If (ADC("BmBagCoCdFr") = ADC("BmBagCoCdTo") And Trim(ADC("BmBagCoCdSel") = "") Or (ADC("BmBagCoCdFr") = ADC("BmBagCoCdTo") = "" And InStr(1, Trim(ADC("BmBagCoCdSel")), ",") = 0)) Then
        FraCpyRfId.Visible = True
        FraCpyRfId.Enabled = True
        FraCpyRfId.ZOrder
        Call EnaDisaCmds(True, CmdCpyRfId)
        CmdCpyRfId.SetFocus
      Else  '6.1b-22 rfid changed as Bag Alias No
        DispMsg "Cannot Enter Bag Alias No. List as multiple companies selected", etError: Exit Sub
      End If
    End If
   '*******Bhavna added co scope  ..rfid only if one comp
   
End Sub
'6.1b-22 rfid changed as Bag Alias No
Private Sub CmdCpyRfIdGo_Click()
If LTrim(ADC("WRFIDSEL")) = "" Then DispMsg "Enter Bag Alias No. List / Bag Nos List", etError: Exit Sub
          pb_CpyRfId = True
         Call CpyWIPBags
         pb_CpyRfId = False
         Call CmdCpyRfId_Click
End Sub

Private Sub CmdGo_Click()
     If Not ADC.SaveRec Then
    DispMsg "X Save Failed", etInfo
    Exit Sub
  End If
  
    Call VchReco
    FraMsg.Enabled = False
    FraMsg.Visible = False
    Call EnaDisaCmds(False, CmdReco)
    CmdReco.SetFocus
    
End Sub

Private Sub CmdReco_Click()
    
    If FraMsg.Visible = True Then
      FraMsg.Visible = False
      FraMsg.Enabled = False
      Call EnaDisaCmds(False, CmdReco)
      CmdReco.SetFocus
    Else
      'Manoj 2.11.0 Correction
      'Check whether the Current Voucher is already Re-Conciled
      If moCn.RecSeek("Select 'x' From BgRecoMst Where " + _
                    " BmCoCd = '" + ADC("BmCoCd") + "' And BmTc= '" + ADC("BmTc") + "' And" + _
                                    " BmYy = '" + ADC("BmYy") + "' And BmChr= '" + ADC("BmChr") + "' And BmNo= " + CStr(ADC("BmNo")) + " And BmRecoYN='Y'") Then
        DispMsg "Voucher already Reconciled, Cannot Reconcile again", etError: Exit Sub
      End If
      'Manoj 2.11.0 Correction
      If ADC("BmBagCoCdFr") = "" And ADC("BmBagCoCdTo") = "" And ADC("BmBagCoCdSel") = "" Then
          DispMsg "Enter Bag's Company scope Fields", etError: Exit Sub
      End If
      If ADC("BmBagLocFr") = "" And ADC("BmBagLocTo") = "" And ADC("BmBagLocSel") = "" Then
          DispMsg "Enter Bag's Location scope Fields", etError: Exit Sub
      End If
      
      FraMsg.Visible = True
      FraMsg.Enabled = True
      FraMsg.ZOrder
      Call EnaDisaCmds(True, CmdReco)
      CmdReco.SetFocus
    End If
    
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
' ***** Manali 3.10.0 - 09/05/12 - Bar Code (Ctrl+N)
  If Not (KeyCode = vbKeyN And Shift = 2) Then Exit Sub
  If UCase(Me.ActiveControl.Name) = UCase("GrdBgReco") Then
    With Me.ActiveControl
      If .Col = 2 Then
        mb_AddRow = False
      End If
    End With
  End If
' ***** Manali 3.10.0 - 09/05/12 - Bar Code (Ctrl+N)
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
' ***** Manali 3.10.0 - 12/04/12 - Bar Code
If KeyAscii <> vbKeyReturn Then Exit Sub
If UCase(Me.ActiveControl.Name) = UCase("GrdBgReco") Then
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



Private Sub GrdBgReco_ColChange(ByVal RowNum As Integer, ByVal ColName As String, ByVal NewValue As Variant, Cancel As Boolean, ErrMsg As String)
  
  Dim ws_OldYy As String, ws_OldChr As String, ws_OldNo As Single, ws_BagNo() As String, ws_BagNoBarCd As String
  
  'All Checks for Valid Bag Numbers
  With GrdBgReco
    Select Case ColName
    '**********Bhavna added co cd
    'uni.172-Bag selection is added
    Case Is = UCase("BdBagCoCd")
        Cancel = (NewValue = "ZZZ") Or (Not moCn.RecSeek("Select 'x' from Head where HCoCd= '" + NewValue + "'")) Or _
                IIF(ADC("BmBagCoCdFr") = "" And ADC("BmBagCoCdTo") = "", "False", (Not (NewValue >= ADC("BmBagCoCdFr") And NewValue <= ADC("BmBagCoCdTo")))) Or _
                IIF(ADC("BmBagCoCdSel") = "", "False", (Not ("," + ADC("BmBagCoCdSel") + ",") Like ("*," + NewValue + ",*")))
        If Cancel = True Then ErrMsg = "Invalid Bag Company Code/not in Company Cd range Selected above": Exit Sub
    Case Is = UCase("wbagno")
        'uni.173-if company code in bag no, then it will overwrite the existing company code in grid.
        ws_BagNo = Split(NewValue, "/")
        If UBound(ws_BagNo) = 3 Then .Value(RowNum, "BdBagCoCd") = ws_BagNo(0)
                  
        'uni.156 calling function to ignore CoCd if entered
        ws_BagNoBarCd = ParseBagNoBarCd(NewValue)
        ws_BagNo = Split(ws_BagNoBarCd, "/")
        
        If UBound(ws_BagNo) <> 2 Then Cancel = True: ErrMsg = "Invalid Bag No": Exit Sub
        If Not IsNumeric(ws_BagNo(2)) Then Cancel = True: ErrMsg = "Invalid Bag No": Exit Sub
        
        ws_OldYy = .Value(RowNum, "BDBYY")
        ws_OldChr = .Value(RowNum, "BDBCHR")
        ws_OldNo = .Value(RowNum, "BDBNO")
        
        .Value(RowNum, "BDBYY") = ws_BagNo(0)
        Call GrdBgReco_ColChange(RowNum, "BDBYY", ws_BagNo(0), Cancel, "")
        If Cancel = True Then ErrMsg = "Invalid Bag Year": .Value(RowNum, "BDBYY") = ws_OldYy: Exit Sub

        .Value(RowNum, "BDBCHR") = UCase(ws_BagNo(1))
        Call GrdBgReco_ColChange(RowNum, "BDBCHR", ws_BagNo(1), Cancel, "")
        If Cancel = True Then ErrMsg = "Either the Bag Character is Invalid": .Value(RowNum, "BDBCHR") = ws_OldChr: Exit Sub

        .Value(RowNum, "BDBNO") = ws_BagNo(2)
        Call GrdBgReco_ColChange(RowNum, "BDBNO", ws_BagNo(2), Cancel, "")
        If Cancel = True Then ErrMsg = "Invalid Bag Number Or Bag Closed": .Value(RowNum, "BDBNO") = ws_OldNo: Exit Sub
        
        ' ***** Manali 3.10.0 - 12/04/12 - Bar Code
        ' ***** Manali 3.10.0 - 09/05/12 - Bar Code (Ctrl+N)
        If mb_AddRow Then
          If .Rows - 1 = RowNum Then
            .AddItem
            .Row = RowNum + 1
            .Col = 2
          End If
        Else
          mb_AddRow = IIF(ADC("BmCalcActQtyYN") = "Y", True, False)     ' ***** Manali 3.10.0 - 09/05/12 - Bar Code (Ctrl+N)
        End If
        ' ***** Manali 3.10.0 - 12/04/12 - Bar Code
        
    Case Is = UCase("BdBYy")
      Cancel = (NewValue <> "") And (Not moCn.RecSeek("Select PMCd From Param Where PTyp= 'YY' And " + _
                                     "PMCd = '" + NewValue + "'"))
      'ErrMsg = "Invalid Bag Year": Exit Sub
      
    Case Is = UCase("BdBChr")
      Cancel = (NewValue <> "") And (Not moCn.RecSeek("Select PMCd From Param Where PTyp= 'BCHR' " + _
                                     "And PMCd = '" + NewValue + "' "))
      'ErrMsg = "Either the Bag Character is Invalid": Exit Sub
      '*********Bhavna changed BmCocd to BdCoCd
    Case Is = UCase("BdBNo")
         Cancel = (NewValue <> 0) And (Not moCn.RecSeek("Select BNo From Bag Where " + _
                                      IIF(gs_Partition = ctCurrPrtn, " BPrtKey='" + ctCurrPrtn + "' And ", "") + " BCoCd= '" + .Value(RowNum, "BdBagCoCd") + "' and " + _
                                      "BYy= '" + .Value(RowNum, "BdBYy") + "' And " + _
                                      "BChr = '" + .Value(RowNum, "BdBChr") + "' And " + _
                                      "BNo= " + CStr(NewValue) + " and BCls='N'"))
      'If Cancel = True Then: ErrMsg = "Invalid Bag Number Or Bag Closed": Exit Sub
    
    End Select
  End With
  
End Sub
Private Sub GrdBgReco_RowWhen(ByVal RowNum As Integer)
    
    'Generating the Serial Number in the Grid
    With GrdBgReco
      'If UCase(ADC("BmRecoYN")) = "Y" Then ADC.AllowDelete = False
      If .IsNew(RowNum) And .Value(RowNum, "BdSr") = 0 Then
        .Value(RowNum, "BdSr") = .MaxVal("BdSr") + 1
      End If
    End With
    
    mb_AddRow = IIF(ADC("BmCalcActQtyYN") = "Y", True, False)    ' ***** Manali 3.10.0 - 09/05/12 - Bar Code (Ctrl+N)
End Sub

Private Sub GrdBgReco_ColWhen(ByVal RowNum As Integer, ByVal ColName As String, Cancel As Boolean, ErrMsg As String, ByVal Action As MwfCtl.en_fgAction)
        
        'Lock Grid if Voucher already Reconciled
        If UCase(ADC("BmRecoYN")) = "Y" Then
           Cancel = True: ErrMsg = "Voucher already Reconciled, Cannot Edit": Exit Sub
        End If
    
      If ADC("BmBagCoCdFr") = "" And ADC("BmBagCoCdTo") = "" And ADC("BmBagCoCdSel") = "" Then
          Cancel = True: ErrMsg = "Enter Bag's Company scope Fields": Exit Sub
      End If
      If ADC("BmBagLocFr") = "" And ADC("BmBagLocTo") = "" And ADC("BmBagLocSel") = "" Then
          Cancel = True: ErrMsg = "Enter Bag's Location scope Fields": Exit Sub
      End If
    
    With GrdBgReco
        '*** Set Helps For Each Field
         Select Case ColName
            '****Bhavna added bag co cd
            Case UCase("BdBagCoCd")
            If .Value(RowNum, "BdBagCoCd") = "" Then
                If RowNum = 1 Then .Value(RowNum, "BdBagCoCd") = ADC("BmCoCd")
                If RowNum > 1 Then .Value(RowNum, "BdBagCoCd") = .Value(RowNum - 1, "BdBagCoCd")
            End If
            
            'Cannot edit Bag co cd once entered, User will have to delete and then add it again
            If moCn.RecSeek("Select 'x' From BgRecoDet Where BdCoCd= '" + gs_CoCd + "' and " + _
                "BdTc= '" + ADC("BmTc") + "' and BdYy= '" + ADC("BmYy") + "' and " + _
                "BdChr= '" + ADC("BmChr") + "' and BdNo= " + CStr(ADC("BmNo")) + " and " + _
                "BdSr= " + CStr(GrdBgReco.Value(RowNum, "BdSr"))) Then _
            Cancel = True: ErrMsg = "Cannot Edit This Value once Bag details have been entered": Exit Sub
            Call HlpList.hCoCd
            
            Case UCase("wBagNo")
            
            'Cannot edit Bag No. once entered, User will have to delete and then add it again
            If moCn.RecSeek("Select 'x' From BgRecoDet Where BdCoCd= '" + gs_CoCd + "' and " + _
                "BdTc= '" + ADC("BmTc") + "' and BdYy= '" + ADC("BmYy") + "' and " + _
                "BdChr= '" + ADC("BmChr") + "' and BdNo= " + CStr(ADC("BmNo")) + " and " + _
                "BdSr= " + CStr(GrdBgReco.Value(RowNum, "BdSr"))) Then _
            Cancel = True: ErrMsg = "Cannot Edit This Value once Bag details have been entered": Exit Sub
             
            'wBagNo Readonly if Voucher is reconciled
            Case UCase("BdActualQty")
                 'All Bags details to be generated on wBagNo entry
'                    If .IsNew(RowNum) And .Mode = fgmnorm Then
'                        .Value(RowNum, "BdBQty") = 0
'                        .Value(RowNum, "BdOthLocBg") = "N"
'
'
'                        .Value(RowNum, "BdDmCd") = moCn.GetFldVal("Select BOdDmCD From Bag Where " + _
'                                                        " BCoCd= '" + gs_CoCd + "' and " + _
'                                                        " BYy= '" + ws_BagNo(0) + "' And " + _
'                                                        " BChr = '" + ws_BagNo(1) + "' And " + _
'                                                        " BNo= " + CStr(ws_BagNo(2)))
'
'                        .Value(RowNum, "BdGrWt") = moCn.GetFldVal("Select BGrWt From Bag Where " + _
'                                                        " BCoCd= '" + gs_CoCd + "' and " + _
'                                                        " BYy= '" + ws_BagNo(0) + "' And " + _
'                                                        " BChr = '" + ws_BagNo(1) + "' And " + _
'                                                        " BNo= " + CStr(ws_BagNo(2)))
'
'                    End If
                                                        
         End Select
     End With
End Sub
 
Private Sub GrdBgReco_ReCalc(ByVal RowNum As Integer, ByVal ColName As String, ByVal OldVal As Variant, ByVal SourceName As String, ByVal SourceColName As String)
  Dim wi_Row As Integer
  With GrdBgReco
    wi_Row = .Row
    'wi_Row = RowNum
    Select Case ColName
    Case Is = UCase("BdDmCd")
      '*** Bhavna changed gs_co cd to bag co cd
      If .IsNew(wi_Row) And .Mode = fgmnorm Then
        .Value(wi_Row, "BdDmCd") = moCn.GetFldVal("Select BOdDmCD From Bag Where " + IIF(gs_Partition = ctCurrPrtn, " BPrtKey='" + ctCurrPrtn + "' And ", "") + _
                                        " BCoCd= '" + .Value(wi_Row, "BdBagCoCd") + "' and " + _
                                        " BYy= '" + .Value(wi_Row, "BdBYy") + "' And " + _
                                        " BChr = '" + .Value(wi_Row, "BdBChr") + "' And " + _
                                        " BNo= " + CStr(.Value(wi_Row, "BdBNo")))
      End If
    Case Is = UCase("BdGrWt")
      If .IsNew(wi_Row) And .Mode = fgmnorm Then
        .Value(wi_Row, "BdGrWt") = moCn.GetFldVal("Select BGrWt From Bag Where " + IIF(gs_Partition = ctCurrPrtn, " BPrtKey='" + ctCurrPrtn + "' And ", "") + _
                                        " BCoCd= '" + .Value(wi_Row, "BdBagCoCd") + "' and " + _
                                        " BYy= '" + .Value(wi_Row, "BdBYy") + "' And " + _
                                        " BChr = '" + .Value(wi_Row, "BdBChr") + "' And " + _
                                        " BNo= " + CStr(.Value(wi_Row, "BdBNo")))
      End If
    Case Is = UCase("BdOthLocBg")
      If .IsNew(wi_Row) And .Mode = fgmnorm Then
        .Value(wi_Row, "BdOthLocBg") = "N"
      End If
    
    Case Is = UCase("BDActualQty")
      If .IsNew(wi_Row) And .Mode = fgmnorm Then
        If ADC("BmCalcActQtyYN") = "Y" Then
          
          .Value(wi_Row, "BDActualQty") = moCn.GetFldVal("Select BQty From Bag Where " + IIF(gs_Partition = ctCurrPrtn, " BPrtKey='" + ctCurrPrtn + "' And ", "") + _
                                          " BCoCd= '" + .Value(wi_Row, "BdBagCoCd") + "' and " + _
                                          " BYy= '" + .Value(wi_Row, "BdBYy") + "' And " + _
                                          " BChr = '" + .Value(wi_Row, "BdBChr") + "' And " + _
                                          " BNo= " + CStr(.Value(wi_Row, "BdBNo")))
        End If
      End If
    
    End Select
  End With
  
End Sub
Private Sub GrdBgReco_InitKey(ByVal RowNum As Integer)
  
  '*** Initialise Key Fields and all the fields which are not included in the grid
  With GrdBgReco
    If .IsNew(RowNum) Then
      .Store "BdCoCd", ADC("BmCoCd")
      .Store "BdTc", ADC("BmTc")
      .Store "BdYy", ADC("BmYy")
      .Store "BdChr", ADC("BmChr")
      .Store "BdNo", ADC("BmNo")
      .Store "BdOthLocBg", "N"
      .Store "BdOthLoc", ""
    End If
  End With
  
End Sub

Private Sub ADC_SetRecSource()
  '*** Set the RecSource for GrdBgReco (Table BgReco)
  ADC.RecSource = " Select * from BgRecoMst " + _
                       " Where BmCoCd = '" + ADC("BmCoCd") + "'" + _
                       " And BmTc ='" + ADC("BmTc") + "'" + _
                       " And BmYy = '" + ADC("BmYy") + "'" + _
                       " And BmChr ='" + ADC("BmChr") + "' " + _
                       " And BmNo =" + CStr(ADC("BmNo"))

End Sub

Private Sub GrdBgReco_SetRecSource()
  '*** Set the RecSource for GrdBgReco (Table BgReco)
  GrdBgReco.RecSource = " Select * from BgRecoDet " + _
                       " Where BdCoCd = '" + ADC("BmCoCd") + "'" + _
                       " And BdTc ='" + ADC("BmTc") + "'" + _
                       " And BdYy = '" + ADC("BmYy") + "'" + _
                       " And BdChr ='" + ADC("BmChr") + "' " + _
                       " And BdNo =" + CStr(ADC("BmNo")) + _
                       " Order By BdSr"
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
  If Not pv_CmdBut Is Nothing Then pv_CmdBut.Enabled = True
  
End Sub

Private Sub CpyWIPBags()
    Dim wCnd As String
    Dim wRfIdStr As String, wRfIdNoArr() As String, wRfIdNo As String, i As Integer, wTotRfid As Integer
    Dim wMinSr As Integer, wBagNoCoCd() As String, wRfIdStrCoCd As String
    '****Bhavna added bag no list selection
    wCnd = ""
    If pb_CpyRfId Then
           Dim wInvalRfIdStr As String
           '6.1b-22 rfid changed as Bag Alias No
           If ADC("wRfIdSel") = "" Then DispMsg "Enter Bag Alias No. List / Bag Nos List", etError: Exit Sub
           wInvalRfIdStr = ""
                      
           'vk.4 calling function to replace enter with ,
            wRfIdStr = PrepareRfIdString(ADC("wRfIdSel"))
            Dim wRfIdStrNew As String
            wRfIdStrNew = wRfIdStr
      
           If ADC("wRfIdSel") <> "" Then
            wRfIdNoArr = Split(wRfIdStr, ",")
        
            If Left(wRfIdStr, 1) = "," Or Right(wRfIdStr, 1) = "," Then
                DispMsg "Invalid Tag Input / Bag No List", etError: Exit Sub
            End If
            wRfIdStr = ""
            For i = 0 To UBound(wRfIdNoArr)
              wRfIdNo = wRfIdNoArr(i)
              
              wBagNoCoCd = Split(wRfIdNo, "/")
              wRfIdNo = ParseBagNoBarCd(wRfIdNo)
              
              If ADC("WCPYTHRUBAGNO") = "N" Then
                If Not IsNumeric(wRfIdNo) Then DispMsg "Tag Not Numeric", etError: Exit Sub
                If Not moCn.RecSeek("Select 'x' from RfIdMst Where RimNo=" + wRfIdNo + " ") Then
                    wInvalRfIdStr = wInvalRfIdStr + IIF(wInvalRfIdStr = "", "", ",") + wRfIdNo
                 End If
              Else
                'uni.173- if company code in bagno , that is used to check bagno
                If Not moCn.RecSeek("Select 'x' from Bag Where BCoCd = '" + IIF(UBound(wBagNoCoCd) = 3, wBagNoCoCd(0), ADC("BmCoCd")) + "' and Byy+'/'+BChr+'/'+ltrim(RTrim(str(BNo))) = '" + wRfIdNo + "'") Then
                    wInvalRfIdStr = wInvalRfIdStr + IIF(wInvalRfIdStr = "", "", ",") + wRfIdNo
                End If
                'uni.173- company code in bagno is added
                wRfIdStrCoCd = wRfIdStrCoCd + IIF(wRfIdStrCoCd = "", "", ",") + IIF(UBound(wBagNoCoCd) = 3, wBagNoCoCd(0), ADC("BmCoCd")) + "/" + wRfIdNo
              End If
              wRfIdStr = wRfIdStr + IIF(wRfIdStr = "", "", ",") + wRfIdNo
            Next
            wTotRfid = UBound(wRfIdNoArr) + 1
            If wInvalRfIdStr <> "" Then DispMsg "Tag / Bag No Not Defined:" + wInvalRfIdStr, etError: Exit Sub
            'vk.4 ADC("wRfIdSel") replaced with wRfIdStr
            If LTrim(wRfIdStr) <> "" And ADC("WCPYTHRUBAGNO") = "N" Then _
              wCnd = " Exists(Select 'x' From RfIdMst where RimNo in (" + wRfIdStr + ") and BCoCd=RimBCoCd and BYy=RimBYy and BChr=RimBChr and BNo=RimBNo)"
            'uni.173- company code in bagno is added
            'vk.4 ADC("wRfIdSel") replaced with wRfIdStr
            If LTrim(wRfIdStr) <> "" And ADC("WCPYTHRUBAGNO") = "Y" Then _
              wCnd = " charindex(',' + BCoCd + '/' + BYy + '/' + BChr + '/' + LTrim(RTrim(str(BNo)))+',' , '," + wRfIdStrCoCd + ",') <> 0 "
          End If
       Else
       '*******Bhavna added co cd and loc scope
            If ADC("BmBagCoCdFr") <> "" Then wCnd = "BCoCd >='" + ADC("BmBagCoCdFr") + "'"
            If ADC("BmBagCoCdTo") <> "" Then wCnd = wCnd + IIF(wCnd <> "", " and ", "") + "BCoCd <='" + ADC("BmBagCoCdTo") + "'"
            If ADC("BmBagCoCdSel") <> "" Then wCnd = wCnd + IIF(wCnd <> "", " and ", "") + "'," + ADC("BmBagCoCdSel") + ",' Like '%,'+BCoCd+',%' "
            If ADC("BmBagLocFr") <> "" Then wCnd = wCnd + IIF(wCnd <> "", " and ", "") + "BLoc >='" + ADC("BmBagLocFr") + "'"
            If ADC("BmBagLocTo") <> "" Then wCnd = wCnd + IIF(wCnd <> "", " and ", "") + "BLoc <='" + ADC("BmBagLocTo") + "'"
            If ADC("BmBagLocSel") <> "" Then wCnd = wCnd + IIF(wCnd <> "", " and ", "") + "'," + ADC("BmBagLocSel") + ",' Like '%,'+BLoc+',%' "
            If ADC("BmFgSubLocFr") <> "" Then wCnd = wCnd + IIF(wCnd <> "", " and ", "") + "BFgSubLoc >='" + ADC("BmFgSubLocFr") + "'"
            If ADC("BmFgSubLocTo") <> "" Then wCnd = wCnd + IIF(wCnd <> "", " and ", "") + "BFgSubLoc <='" + ADC("BmFgSubLocTo") + "'"
            If ADC("BmFgSubLocSel") <> "" Then wCnd = wCnd + IIF(wCnd <> "", " and ", "") + "'," + ADC("BmFgSubLocSel") + ",' Like '%,'+BFgSubLoc+',%' "
            
             wCnd = wCnd + IIF(wCnd <> "", " and ", "") + " BCls='N' And " + _
                    " (BQty<> 0 or BPtQty<> 0 or BRjQty <> 0 or " + _
                    " (BQty= 0 and BPtQty= 0 and BRjQty= 0 and BRjPtQty= 0 and Round(BGrWt, 3)<> 0)) And " + _
                    " Not Exists (Select 'x' From BgRecoDet " + _
                    " Join BgRecoMst On BdCoCd=BmCocd And BdTc=BmTc And BdYy=BmYy And BdChr=BmChr And BdNo=BmNo " + _
                    " Where BdCoCd='" + ADC("BmCoCd") + "' And BdTc='" + ADC("BmTc") + "' And BdYy='" + ADC("Bmyy") + "' And BdChr='" + ADC("BmChr") + "' And BdNo=" + CStr(ADC("BmNo")) + _
                    "           And BdBagCoCd=BCoCd And BdBYy=BYy And BdBChr=BChr And BdBNo=BNo) "
                    
                    '/* And BmLoc=BLoc And  " + "' and BLoc>='" + adc("BmFgSubLocFr") + "' and BLoc<='" + adc("BmFgSubLocTo") + "' " + _
                    " and ','+BLoc+',' Like '%," + adc("BmFgSubLocSel") + ",%' and BmDt='" + CStr(adc("BmDt")) + "'*/ )"
      End If
 

        

        If pb_CpyRfId Then MWLib.BeginProcess Me, "Copying Bags ..."
        Dim wRsBg As MwfLib.MDORowSet, wRow As Single, wsActQty As Double, wLstSr As Single
                
        Set wRsBg = moCn.OpenRes("Select * from Bag where " + IIF(gs_Partition = ctCurrPrtn, " BPrtKey='" + ctCurrPrtn + "' And ", "") + wCnd)
                   
        wLstSr = GrdBgReco.MaxVal("BdSr") + 1
        'If wLstSr = 0 Then wLstSr = 1
        
        With wRsBg
          GrdBgReco.StartCopy
          wMinSr = GrdBgReco.MaxVal("BdSr") + 1
          Do While Not (.EOF Or .BOF)
              GrdBgReco.AddItem
              MWLib.RunningProcess "Adding WIP Bag " + !BYy + "/" + !BChr + "/" + CStr(!BNo)
              wRow = GrdBgReco.Rows - 1
                            
              GrdBgReco.Value(wRow, "BdSr") = wLstSr
              GrdBgReco.Value(wRow, "BdBYy") = !BYy
              GrdBgReco.Value(wRow, "BdBagCoCd") = !BCoCd
              GrdBgReco.Value(wRow, "BdBChr") = !BChr
              GrdBgReco.Value(wRow, "BdBNo") = !BNo
              GrdBgReco.Value(wRow, "BdActualQty") = IIF(pb_CpyRfId And ADC("BmCalcActQtyYN") = "Y", IIF(!BQty = 0, !BRjQty, !BQty), 0)
              GrdBgReco.Value(wRow, "BdBQty") = IIF(!BQty = 0, !BRjQty, !BQty)
              GrdBgReco.Value(wRow, "BdDmCd") = !BOdDmCd
              GrdBgReco.Value(wRow, "BdGrWt") = !BGrWt
              GrdBgReco.Value(wRow, "BdExtQty") = 0
              GrdBgReco.Value(wRow, "BdUnAcQty") = IIF(!BQty = 0, !BRjQty, !BQty)
              GrdBgReco.Value(wRow, "BdOthLocBg") = "N"
              wLstSr = wLstSr + 1
                                
            GrdBgReco.SaveRec (wRow)
            .MoveNext
          Loop
          
          MWLib.EndProcess Me
          GrdBgReco.EndCopy
          Call atxt_ReCalculate(ADC("WTOTBAGS").Index, "WTOTBAGS", ADC("WTOTBAGS"), GrdBgReco.Rows - 1, "GrdBgReco", "BDACTUALQTY")
          Call atxt_ReCalculate(ADC("WTOTACTQTY").Index, "WTOTACTQTY", ADC("WTOTACTQTY"), GrdBgReco.Rows - 1, "GrdBgReco", "BDACTUALQTY")
            
          
     If pb_CpyRfId And ADC("wCpyThruBagNo") = "N" Then
        Dim wRsRfId As MwfLib.MDORowSet
        'vk.4  adc("wRfIdSel") replaced with wRfIdStr
        Set wRsRfId = moCn.OpenRes("Select RimNo from RfIdMst where Not Exists(Select 'x' From BgRecoDet where " + _
                                        " BdCoCd=RimBCoCd and BdBYy=RimBYY and BdBChr=RimBChr and BdBNo=RimBNo and " + _
                                        " BdCoCd='" + ADC("BmCoCd") + "' and BdTc='" + ADC("BmTc") + "' and BdYy='" + ADC("BmYy") + "' and BdChr='" + ADC("BmChr") + "' and BdNo=" + CStr(ADC("BmNo")) + _
                                        " and  BdSr>=" + CStr(wMinSr) + " and BdSr<=" + CStr(GrdBgReco.MaxVal("BdSr")) + ")" + _
                                        " and RimNo in (" + wRfIdStr + ")")
      
      
          wRfIdStr = ""
          Do While Not (wRsRfId.EOF Or wRsRfId.BOF)
                wRfIdStr = wRfIdStr + IIF(wRfIdStr = "", "", ",") + CStr(wRsRfId!RimNo)
            wRsRfId.MoveNext
          Loop
            '6.1b-22 rfid changed as Bag Alias No
            If wRfIdStr <> "" Then DispMsg "Total Bag Alias No. {" + CStr(wTotRfid) + " } :Bag Alias No. Not Copied :  " + wRfIdStr, etError: Exit Sub
        End If
            
        End With
        Set wRsBg = Nothing
        '*** End of the copy process

End Sub

Private Sub VchReco()
  
  '*** Voucher Reconciliation Routine ***
  Dim wReco As Single
  Dim moRs As MDORowSet
  'Check whether the Voucher is of current Date (Below Check not required)
  'If moCn.SrvrDate <> ADC("BmDt") Then DispMsg "Cannot Reconcile Previous Date Vouchers", etError: Exit Sub
    
'  'Check whether the Current Voucher is already Re-Conciled
'  If moCn.RecSeek("Select 'x' From BgRecoMst Where " + _
'                " BmCoCd = '" + ADC("BmCoCd") + "' And BmTc= '" + ADC("BmTc") + "' And" + _
'                                " BmYy = '" + ADC("BmYy") + "' And BmChr= '" + ADC("BmChr") + "' And BmNo= " + CStr(ADC("BmNo")) + " And BmRecoYN='Y'") Then
'    DispMsg "Voucher already Reconciled, Cannot Reconcile again", etError: Exit Sub
'  End If
     '********Bhavna added loc,co scope
    Dim wCnd, wCoCnd As String
    wCnd = ""
    wCoCnd = ""
            If ADC("BmBagLocFr") <> "" Then wCnd = wCnd + IIF(wCnd <> "", " and ", "") + "BLoc >='" + ADC("BmBagLocFr") + "'"
            If ADC("BmBagLocTo") <> "" Then wCnd = wCnd + IIF(wCnd <> "", " and ", "") + "BLoc <='" + ADC("BmBagLocTo") + "'"
            If ADC("BmBagLocSel") <> "" Then wCnd = wCnd + IIF(wCnd <> "", " and ", "") + "'," + ADC("BmBagLocSel") + ",' Like '%,'+BLoc+',%' "
            If ADC("BmBagCoCdFr") <> "" Then wCoCnd = wCoCnd + IIF(wCoCnd <> "", " and ", "") + "BCoCd >='" + ADC("BmBagCoCdFr") + "'"
            If ADC("BmBagCoCdTo") <> "" Then wCoCnd = wCoCnd + IIF(wCoCnd <> "", " and ", "") + "BCoCd <='" + ADC("BmBagCocdTo") + "'"
            If ADC("BmBagCocdSel") <> "" Then wCoCnd = wCoCnd + IIF(wCoCnd <> "", " and ", "") + "'," + ADC("BmBagCocdSel") + ",' Like '%,'+BCoCd+',%' "
    '*******
    '********Bhavna added loc,co scope
   
  '*** Beginning the copy process
  MWLib.BeginProcess Me, "Reconciling Voucher..."
      wReco = 0
      With GrdBgReco
        For wReco = 1 To (GrdBgReco.Rows - 1)
'           Set moRs = moCn.OpenRes("Select (Case When BQty<>0 Then BQty When BRjQty<> 0 Then BRjQty Else 0 End) as qBQty, " + _
                                  "(Case BLoc When '" + ADC("BmLoc") + "' Then 'N' Else 'Y' END) qOthLocBg, " + _
                                  "(Case BLoc When '" + ADC("BmLoc") + "' Then '' Else BLoc END) qOthLoc From Bag Where " + IIF(gs_Partition = ctCurrPrtn, " BPrtKey='" + ctCurrPrtn + "' And ", "") + _
                                  "BCoCd= '" + gs_CoCd + "' and " + _
                                  "BYy= '" + .Value(wReco, "BdBYy") + "' And " + _
                                  "BChr = '" + .Value(wReco, "BdBChr") + "' And " + _
                                  "BNo= " + CStr(.Value(wReco, "BdBNo")))
            Set moRs = moCn.OpenRes("Select (Case When BQty<>0 Then BQty When BRjQty<> 0 Then BRjQty Else 0 End) as qBQty, " + _
                                  "(case when " + wCnd + " Then 'N' Else 'Y' end) qOthLocBg, " + _
                                  "(case when " + wCnd + " Then '' Else BLoc END) qOthLoc From Bag Where " + IIF(gs_Partition = ctCurrPrtn, " BPrtKey='" + ctCurrPrtn + "' And ", "") + _
                                  wCoCnd + IIF(wCoCnd <> "", " and ", "") + _
                                  "BYy= '" + .Value(wReco, "BdBYy") + "' And " + _
                                  "BChr = '" + .Value(wReco, "BdBChr") + "' And " + _
                                  "BNo= " + CStr(.Value(wReco, "BdBNo")))
'           "BCoCd= '" + .Value(wReco, "BdBagCoCd") + "' and " + _
            'Entering the Bag Quantity (Normal or Rejected)
            .Value(wReco, "BdBQty") = moRs.FldValue("qBQty")
            'Entering whether the Bag belongs to Other Location
            .Value(wReco, "BdOthLocbg") = moRs.FldValue("qOthLocbg")
            'Entering the Bag Location
            .Value(wReco, "BdOthLoc") = moRs.FldValue("qOthLoc")
            'Calculation of Extra Qty and Un-Accounted Quantity
            If .Value(wReco, "BdOthLocBg") = "N" Then
              If .Value(wReco, "BdActualQty") > .Value(wReco, "BdBQty") Then
                 .Value(wReco, "BdExtQty") = .Value(wReco, "BdActualQty") - .Value(wReco, "BdBQty")
                 .Value(wReco, "BdUnAcQty") = 0
      
              ElseIf .Value(wReco, "BdActualQty") < .Value(wReco, "BdBQty") Then
                 .Value(wReco, "BdUnAcQty") = .Value(wReco, "BdBQty") - GrdBgReco.Value(wReco, "BdActualQty")
                 .Value(wReco, "BdExtQty") = 0
      
              Else
                 .Value(wReco, "BdExtQty") = 0
                 .Value(wReco, "BdUnAcQty") = 0
      
              End If
            Else
                 .Value(wReco, "BdExtQty") = GrdBgReco.Value(wReco, "BdActualQty")
                 .Value(wReco, "BdUnAcQty") = 0
            End If
                                   
           'Entering the Bag Reconciled status
           GrdBgReco.SaveRec (wReco)
           
        Next wReco
        Call CpyWIPBags
        ADC("BmRECOYN") = "Y"
        ADC.SaveRec
        
        'Displaying the Invisible columns with Read Only Property
          .ColProp("BdGrWt").Style = fgcReadOnly
          .ColProp("BdBQty").Style = fgcReadOnly
          .ColProp("BdExtQty").Style = fgcReadOnly
          .ColProp("BdUnAcQty").Style = fgcReadOnly
          .ColProp("BdOthLoc").Style = fgcReadOnly
              
      End With
      MWLib.EndProcess Me

  DispMsg "Voucher Reconciliation Completed", etInfo
  
End Sub

Private Sub HideAllFras()
  FraNKeyAll.Visible = True
  
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
  ''(Bef 13/08/05)
  'SetMDISysColorsAfterLogin   '*** Jenny Color (For Setting System Colors)
  
  Dim ws_BtnStr As String
  ws_BtnStr = "CmdGo"
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
Private Sub GrdBgReco_Validate(Cancel As Boolean)
  Cancel = GrdBgReco.Validate
End Sub


