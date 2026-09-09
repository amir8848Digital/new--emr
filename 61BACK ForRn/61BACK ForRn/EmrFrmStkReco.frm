VERSION 5.00
Object = "{AB2CC32A-E652-42A2-B341-26ECE02F6386}#2.1#0"; "MwfCtl.ocx"
Begin VB.Form EmrFrmStkReco 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Stock Reconciliation"
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
      TabIndex        =   37
      Top             =   6120
      Visible         =   0   'False
      Width           =   14595
      Begin VB.OptionButton OptCpyBagFr 
         Caption         =   "Stock ID Nos"
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
         Left            =   7680
         TabIndex        =   57
         ToolTipText     =   "Specify Copy From Stock ID Nos / Bag Alias No."
         Top             =   150
         Width           =   1695
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
         Index           =   1
         Left            =   9600
         TabIndex        =   56
         ToolTipText     =   "Specify Copy From Stock ID Nos / Bag Alias No."
         Top             =   150
         Width           =   1815
      End
      Begin VB.Frame Frame5 
         Caption         =   "Stock ID Nos / Bag Alias No."
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   2775
         Left            =   60
         TabIndex        =   54
         Top             =   360
         Width           =   13850
         Begin MwfCtl.MWCTL_MED ATXT 
            Height          =   2445
            Index           =   40
            Left            =   60
            TabIndex        =   55
            ToolTipText     =   "Copy Bag Alias No. / Enter Bag Nos  (Yy/Chr/No,YY....,Yy/Chr/No)"
            Top             =   240
            Width           =   13695
            _ExtentX        =   24156
            _ExtentY        =   4313
            Alignment       =   3
            MaxLength       =   0
            IdName          =   "WRFIDSEL"
         End
      End
      Begin MwfCtl.MWCTL_BTN1 CmdCpyRfIdGo 
         Height          =   525
         Left            =   13920
         TabIndex        =   39
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
         Left            =   13530
         TabIndex        =   38
         ToolTipText     =   "Enter 'Y' to Copy Thru Bag Nos and  'N' to Copy thru RfId"
         Top             =   120
         Visible         =   0   'False
         Width           =   315
         _ExtentX        =   556
         _ExtentY        =   503
         DataType        =   4
         IdName          =   "WCPYTHRUSTKNO"
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
         Left            =   6240
         TabIndex        =   58
         Top             =   120
         Width           =   1365
      End
      Begin VB.Label LblCpyThruBagNo 
         BackStyle       =   0  'Transparent
         Caption         =   "Copy thru Stock Id No Y/N"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   10920
         TabIndex        =   41
         Top             =   360
         Width           =   2565
      End
      Begin VB.Label LblMain 
         BackStyle       =   0  'Transparent
         BorderStyle     =   1  'Fixed Single
         Caption         =   "  Copy Stock ID Nos / Bag Alias No."
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
         TabIndex        =   40
         Top             =   90
         Width           =   5850
      End
   End
   Begin MwfCtl.MWCTL_BTN1 CmdCpyRfId 
      Height          =   405
      Left            =   5310
      TabIndex        =   36
      ToolTipText     =   "Copy Option"
      Top             =   9360
      Width           =   1560
      _ExtentX        =   2752
      _ExtentY        =   714
      BackColor       =   -2147483633
      MaskColor       =   8454143
      Caption         =   "C&opy Stk ID Nos / Bag Alias No."
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
      TabIndex        =   18
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
      TabIndex        =   20
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
      TabIndex        =   19
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
      TabIndex        =   22
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
         TabIndex        =   27
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
         TabIndex        =   26
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
         TabIndex        =   25
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
         TabIndex        =   24
         Top             =   120
         Width           =   135
      End
   End
   Begin VB.Frame FRA_NKEYS 
      BorderStyle     =   0  'None
      Enabled         =   0   'False
      Height          =   9300
      Left            =   0
      TabIndex        =   21
      Top             =   0
      Width           =   15225
      Begin VB.Frame FraMsg 
         Height          =   915
         Left            =   5430
         TabIndex        =   32
         Top             =   8325
         Visible         =   0   'False
         Width           =   3825
         Begin MwfCtl.MWCTL_BTN1 CmdGo 
            Height          =   495
            Left            =   3120
            TabIndex        =   17
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
            TabIndex        =   34
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
            TabIndex        =   33
            Top             =   450
            Width           =   3105
         End
      End
      Begin VB.Frame FraNKeyAll 
         BorderStyle     =   0  'None
         Height          =   8025
         Left            =   30
         TabIndex        =   23
         Top             =   1200
         Width           =   15165
         Begin MwfCtl.MWCTL_FLX GrdStkReco 
            Height          =   8025
            Left            =   0
            TabIndex        =   14
            Top             =   0
            Width           =   15135
            _ExtentX        =   26696
            _ExtentY        =   14155
            Cols            =   13
            colname1        =   "BDSR"
            heading1        =   "Sr         "
            datafld1        =   "BdSr"
            datatype1       =   1
            mask1           =   "###0"
            maxlength1      =   4
            style1          =   2
            tooltiptext1    =   "Voucher Serial Number"
            colname2        =   "BDBAGCOCD"
            heading2        =   "Comp Cd"
            datafld2        =   "BdBagCoCd"
            datatype2       =   4
            recalcparent2   =   "GrdStkReco(BDACTUALQTY),GrdStkReco(BDDMCD)"
            maxlength2      =   3
            tooltiptext2    =   "Enter Stock Company Code"
            colname3        =   "WSTKNO"
            datatype3       =   4
            maxlength3      =   18
            tooltiptext3    =   "Read Stock ID Number Using Barcode Reader"
            colname4        =   "BDBYY"
            heading4        =   "StkYy  "
            datafld4        =   "BdBYy"
            recalcparent4   =   "GrdStkReco(BDACTUALQTY),GrdStkReco(BDDMCD)"
            maxlength4      =   2
            style4          =   2
            tooltiptext4    =   "Stock Year (Yy)"
            colname5        =   "BDBCHR"
            heading5        =   "StkChr     "
            datafld5        =   "BdBChr"
            datatype5       =   4
            recalcparent5   =   "GrdStkReco(BDACTUALQTY),GrdStkReco(BDDMCD)"
            maxlength5      =   3
            style5          =   2
            tooltiptext5    =   "Stock Chr"
            colname6        =   "BDBNO"
            heading6        =   "StkNo       "
            datafld6        =   "BdBNo"
            datatype6       =   1
            mask6           =   "######0"
            recalcparent6   =   "GrdStkReco(BDACTUALQTY),GrdStkReco(BDDMCD)"
            maxlength6      =   7
            style6          =   2
            tooltiptext6    =   "Stock Number"
            colname7        =   "BDACTUALQTY"
            heading7        =   "Rept Qty  "
            datafld7        =   "BdActualQty"
            datatype7       =   2
            mask7           =   "##0.0"
            recalcon7       =   "GrdStkReco(BDBAGCOCD),GrdStkReco(BDBYY),GrdStkReco(BDBCHR),GrdStkReco(BDBNO)"
            maxlength7      =   5
            tooltiptext7    =   "Enter Reported (Physical) Quantity"
            colname8        =   "BDBQTY"
            heading8        =   "Stk Qty  "
            datafld8        =   "BdBQty"
            datatype8       =   2
            mask8           =   "##0.0"
            maxlength8      =   5
            style8          =   0
            tooltiptext8    =   "Bag Quantity in WIP"
            colname9        =   "BDDMCD"
            heading9        =   "Design Cd                   "
            datafld9        =   "BdDmCd"
            datatype9       =   4
            recalcon9       =   "GrdStkReco(BDBAGCOCD),GrdStkReco(BDBYY),GrdStkReco(BDBCHR),GrdStkReco(BDBNO)"
            style9          =   2
            tooltiptext9    =   "Bag Design Code"
            colname10       =   "BDEXTQTY"
            heading10       =   "Extra        "
            datafld10       =   "BdExtQty"
            datatype10      =   2
            mask10          =   "###0.0"
            maxlength10     =   6
            style10         =   0
            tooltiptext10   =   "Bag Extra Quantity"
            colname11       =   "BDUNACQTY"
            heading11       =   "Un A/c     "
            datafld11       =   "BdUnAcQty"
            datatype11      =   2
            mask11          =   "###0.0"
            maxlength11     =   6
            style11         =   0
            tooltiptext11   =   "Bag Un-Accounted Quantity"
            colname12       =   "BDOTHLOCBG"
            heading12       =   "Oth Loc  "
            datafld12       =   "BdOthLocBg"
            datatype12      =   4
            maxlength12     =   1
            style12         =   0
            tooltiptext12   =   "Specify whether Bag is Misplaced/Belongs to Other Location"
         End
         Begin MwfCtl.MWCTL_MED ATXT 
            Height          =   285
            Index           =   7
            Left            =   9975
            TabIndex        =   16
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
            ReCalcOn        =   "GrdStkReco(BDACTUALQTY)"
            IdName          =   "WTOTACTQTY"
         End
         Begin MwfCtl.MWCTL_MED ATXT 
            Height          =   285
            Index           =   8
            Left            =   1260
            TabIndex        =   15
            Top             =   8340
            Width           =   900
            _ExtentX        =   1588
            _ExtentY        =   503
            Alignment       =   2
            Enabled         =   0   'False
            Mask            =   "###0"
            MaxLength       =   4
            DataType        =   1
            ReCalcOn        =   "GrdStkReco(BDACTUALQTY)"
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
            TabIndex        =   30
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
            TabIndex        =   29
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
         TabIndex        =   13
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
         Enabled         =   0   'False
         DataType        =   4
         DataField       =   "BmCalcActQtyYN"
         IdName          =   "BMCALCACTQTYYN"
      End
      Begin MwfCtl.MWCTL_MED ATXT 
         Height          =   285
         Index           =   11
         Left            =   1605
         TabIndex        =   10
         ToolTipText     =   "Enter Stock Location From"
         Top             =   765
         Width           =   1410
         _ExtentX        =   2487
         _ExtentY        =   503
         MaxLength       =   5
         DataType        =   4
         DataField       =   "BmBagLocFr"
         IdName          =   "BMBAGLOCFR"
      End
      Begin MwfCtl.MWCTL_MED ATXT 
         Height          =   285
         Index           =   12
         Left            =   3840
         TabIndex        =   11
         ToolTipText     =   "Enter Stock Location To"
         Top             =   765
         Width           =   1410
         _ExtentX        =   2487
         _ExtentY        =   503
         MaxLength       =   5
         DataType        =   4
         DataField       =   "BmBagLocTo"
         IdName          =   "BMBAGLOCTO"
      End
      Begin MwfCtl.MWCTL_MED ATXT 
         Height          =   285
         Index           =   13
         Left            =   6600
         TabIndex        =   12
         ToolTipText     =   "Enter Stock Location Selection"
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
         ToolTipText     =   "Enter Stock Company Code From"
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
         ToolTipText     =   "Enter Stock Company Code To"
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
         ToolTipText     =   "Enter Stock Company Code Selection"
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
         Index           =   0
         Left            =   1605
         TabIndex        =   48
         ToolTipText     =   "Enter FG Sub Location From"
         Top             =   1200
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
         Index           =   5
         Left            =   3840
         TabIndex        =   49
         ToolTipText     =   "Enter FG Sub Location To"
         Top             =   1200
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
         TabIndex        =   50
         ToolTipText     =   "Enter FG Sub Locations"
         Top             =   1200
         Width           =   8610
         _ExtentX        =   15187
         _ExtentY        =   503
         MaxLength       =   255
         DataType        =   4
         DataField       =   "BmFgSubLocSel"
         IdName          =   "BMFGSUBLOCSEL"
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
         TabIndex        =   53
         Top             =   1200
         Width           =   1095
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
         TabIndex        =   52
         Top             =   1200
         Width           =   495
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
         TabIndex        =   51
         Top             =   1200
         Width           =   1470
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
         TabIndex        =   47
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
         TabIndex        =   46
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
         TabIndex        =   45
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
         TabIndex        =   44
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
         TabIndex        =   43
         Top             =   765
         Width           =   495
      End
      Begin VB.Label Label6 
         BackStyle       =   0  'Transparent
         Caption         =   "Stock Loc Cd Fr"
         BeginProperty Font 
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
         TabIndex        =   42
         Top             =   765
         Width           =   1590
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
         TabIndex        =   35
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
         TabIndex        =   31
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
         TabIndex        =   28
         Top             =   150
         Width           =   855
      End
   End
End
Attribute VB_Name = "EmrFrmStkReco"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim moCn As MwfLib.MDOConnection
Dim ms_TcTyp As String, ms_wSrCoCd As String
Dim ms_wSrTc As String, ms_wSrYy As String, ms_wSrChr As String
Dim pb_CpyRfId As Boolean

Private Sub adc_AutoGen()
  If ADC("BmNo") = 0 Then
    ADC("BmNo") = moCn.AutoGen("BgRecoMst", "BmNo", ADC("BmCoCd") + ADC("BmTc") + ADC("BmYy") + ADC("BmChr"))
  Else
    Call moCn.AutoGen("BgRecoMst", "BmNo", ADC("BmCoCd") + ADC("BmTc") + ADC("BmYy") + ADC("BmChr"), ADC("BmNo"))
  End If
End Sub

Private Sub ADC_Load()
  '*** Set the First Key Control property to the field Year
  '*** Set the First Non Key Control property to Date
  '*** Set the Child Property Of ADC to GrdStkReco
  '*** Set the Previous and Next Control for the Grid GrdStkReco
  '*** Set the Hot Keys for the Grid GrdStkReco as 'A'
  Set moCn = ADC.Connection
  Set ADC.FirKeyCtl = ADC("BmYy")
  Set ADC.FirNKeyCtl = ADC("BmDt")
  
  ADC.Child = "GrdStkReco"
  Set GrdStkReco.PrevCtl = ADC("BmDt")
  Set GrdStkReco.NextCtl = ADC
  GrdStkReco.HotKey = "A"
  FraNKeyAll.ZOrder (0)

  ms_wSrCoCd = gs_CoCd
  ms_TcTyp = "JRC"
  ms_wSrTc = moCn.GetFldVal("Select PValue from Param where PTyp= 'TCTYP' and " + _
                "PMCd= '" + ms_TcTyp + "' ")
  ms_wSrYy = moCn.GetFldVal("Select hYy from Head where HCoCd= '" + ms_wSrCoCd + "' and HCd= '" + ctSelfCmCd + "'")

End Sub

Private Sub ADC_KeyWhen()
  '*** Setting the default values for BrCoCd, BrTc, BrYy and BrChr
  '*** Hide all frames except Frame FraNKeyAll
  Call EnaDisaCmds(True)
  Call HideAllFras

  If ms_wSrCoCd <> "" Then ADC("BmCoCd") = ms_wSrCoCd
  If ms_wSrTc <> "" Then ADC("BmTc") = ms_wSrTc
  If ms_wSrYy <> "" Then ADC("BmYy") = ms_wSrYy
  If ms_wSrChr <> "" Then ADC("BmChr") = ms_wSrChr
End Sub

Private Sub ADC_KeyValid(Cancel As Boolean, ErrMsg As String)
  If ADC("BmYy") = "" Then Cancel = True: ErrMsg = "Enter Voucher Year": Exit Sub
  If ADC("BmChr") = "" Then Cancel = True: ErrMsg = "Enter Voucher Character": Exit Sub
  
End Sub

Private Sub ADC_NonKeyWhen(Cancel As Boolean, ErrMsg As String)
  If ADC.IsNew And Not moCn.RecSeek("select vPSCd from vParam where " + _
             " vPCoCd='" + ADC("BmCoCd") + "' and vPTyp= 'CHR' and " + _
             " vPMCd = '" + ADC("BmTc") + "' and vPSCd='" + ADC("BmChr") + "' and vPValidYn<>'N'") Then
    Cancel = True: ErrMsg = "Voucher Character No More Valid": ADC.AllowSave = False: Exit Sub
  Else
    ADC.AllowSave = True
  End If
  
  ms_wSrTc = ADC("BmTc")
  ms_wSrYy = ADC("BmYy")
  ms_wSrChr = ADC("BmChr")
  
  If ADC.AddFndMode = xaddmode Then
    ADC("BmBagCoCdFr") = gs_CoCd
    ADC("BmBagCoCdTo") = gs_CoCd
  End If
  
  'Set The Visible, Enable Properties for the grid fields
  If UCase(ADC("BmRecoYN")) = "Y" Then
    GrdStkReco.ColProp("BdBQty").Style = fgcReadOnly
    GrdStkReco.ColProp("BdExtQty").Style = fgcReadOnly
    GrdStkReco.ColProp("BdUnAcQty").Style = fgcReadOnly
    GrdStkReco.AllowAdd = False
    'Mw.9999 Err.44- A Grid records can be delete after reconciliation
    GrdStkReco.AllowDelete = True
  Else
    GrdStkReco.ColProp("BdBQty").Style = fgcinvisible
    GrdStkReco.ColProp("BdExtQty").Style = fgcinvisible
    GrdStkReco.ColProp("BdUnAcQty").Style = fgcinvisible
    GrdStkReco.AllowAdd = True
    GrdStkReco.AllowDelete = True
  End If
  
  Call EnaDisaCmds(False)

  If ADC.AddFndMode = xaddmode Then
    ADC("BmDt") = moCn.SrvrDate
    ADC("BmRecoYN") = "N"
    ADC("BmCalcActQtyYN") = "N"
  End If
  'vk.4 setting default values
  OptCpyBagFr(0).Value = True
  ADC("wCpyThruStkNo") = "Y"
  
End Sub

Private Sub ATXT_FldWhen(Index As Integer, ByVal IdName As String, Cancel As Boolean, ErrMsg As String)
  '*** Set Help For Each Field
  Select Case IdName
  Case Is = UCase("BmYy")
    Call HlpList.PMCd("YY")
    
  Case Is = UCase("BmChr")
    Call HlpList.vPSCd(ADC("BmCoCd"), "CHR", ADC("BmTc"), "Y")

  Case Is = UCase("BmNo")
    Call HlpList.BrNo(ADC("BmCoCD"), ADC("BmTc"), ADC("BmYy"), ADC("BmChr"))
  
  Case Is = UCase("BmBagCoCdFr"), UCase("BmBagCoCdTo"), UCase("BmBagCoCdSel")
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
        Call HlpList.Loc("", "'S'")
    End If
  
  Case Is = UCase("BmDt")
    If moCn.RecSeek("Select 'x' From BgRecoDet Where BdCoCd='" + ADC("BmCoCd") + "' And BdTc='" + ADC("BmTc") + "' " + _
          "  And BdYy= '" + ADC("BmYy") + "' And BdChr='" + ADC("BmChr") + "' And BdNo= " + CStr(ADC("BmNo"))) Then
        Cancel = True: ErrMsg = "Cannot Edit Date, Child Record Exists": Exit Sub
    End If

  Case Is = UCase("wCpyThruStkNo")
    If ADC("wCpyThruStkNo") = "" Then ADC("wCpyThruStkNo") = "Y"
        
  End Select

End Sub

Private Sub atxt_ReCalculate(Index As Integer, ByVal IdName As String, ByVal OldVal As Variant, ByVal RowNum As Integer, ByVal SourceName As String, ByVal SourceColName As String)
  Select Case UCase(IdName)
  'Getting defaukt Chr for the given TC
  Case Is = UCase("BmChr")
    If ADC.Mode = xNorm Then
      ADC("BmChr") = moCn.GetFldVal("Select vPValue from vParam Where " + _
                                "vPCoCd = '" + ADC("BmCoCd") + "' And vPTyp= 'DEFCHR' " + _
                                "and vPMCd= '" + ADC("BmTc") + "'")
    End If
  End Select
End Sub

Private Sub atxt_FldChange(Index As Integer, ByVal IdName As String, Cancel As Boolean, ErrMsg As String, ByVal pv_NewValue As Variant)
  Select Case UCase(IdName)
  Case Is = UCase("BmYy")
  'Check for valid Year
    Cancel = Not moCn.RecSeek(" select PMCd from Param where PTyp= 'YY' and " + _
                         " PMCd = '" + pv_NewValue + "'")
    If Cancel = True Then ErrMsg = "Invalid Voucher Year": Exit Sub

  Case Is = UCase("BmChr")
  'Check for valid Chr
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
        If Cancel = False Then Exit For
      Next wi_Index
    End If
    If Cancel = True Then ErrMsg = "Access To This Voucher Character Is Denied To User " + gs_UsrCd: Exit Sub

  Case Is = UCase("BmDt")
  'Check for valid Voucher Date
    If ADC.Mode = xNorm Then
      Cancel = pv_NewValue = ""
      If Cancel = False Then Cancel = CDate(pv_NewValue) > moCn.SrvrDate
      If Cancel = True Then ErrMsg = "Invalid Voucher Date": Exit Sub
    End If

  Case Is = UCase("wCpyThruStkNo")
  'Check for valid YN for Copy thru stock/Rfid nos
    If ADC.Mode = xNorm Then
      Cancel = Not moCn.RecSeek("select PMCd from Param where PTyp= 'YN' and PMCd = '" + pv_NewValue + "'")
      ErrMsg = "Enter 'Y' to Copy thru Stock ID No Or 'N' to Copy thru RfId No ": Exit Sub
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
    'Check whether the Current Voucher is already Re-Conciled
    If moCn.RecSeek("Select 'x' From BgRecoMst Where " + _
                    " BmCoCd = '" + ADC("BmCoCd") + "' And BmTc= '" + ADC("BmTc") + "' And" + _
                    " BmYy = '" + ADC("BmYy") + "' And BmChr= '" + ADC("BmChr") + "' And BmNo= " + CStr(ADC("BmNo")) + " And BmRecoYN='Y'") Then
      DispMsg "Voucher already Reconciled, Cannot Edit", etError: Exit Sub
    End If
    'vk.4 ADC("wCpyThruStkNo") = "Y"
'    If (ADC("BmBagCoCdFr") = ADC("BmBagCoCdTo") And Trim(ADC("BmBagCoCdSel") = "") Or (ADC("BmBagCoCdFr") = ADC("BmBagCoCdTo") = "" And InStr(1, Trim(ADC("BmBagCoCdSel")), ",") = 0)) Then
    FraCpyRfId.Visible = True
    FraCpyRfId.Enabled = True
    FraCpyRfId.ZOrder
    Call EnaDisaCmds(True, CmdCpyRfId)
    ADC("WRFIDSEL").SetFocus
'    Else
'      DispMsg "Cannot Enter RfId List as Multiple Companies Selected", etError: Exit Sub
'    End If
  End If
     
End Sub

Private Sub CmdCpyRfIdGo_Click()
'if Copy Thru Stk/Rfid YN is blank OR Stk/Rfid list is blank then gives error message
'else call the procedure to copy records
  If ADC("wCpyThruStkNo") = "" Then DispMsg "Copy thru Stock Id No Y/N Cannot Be Blank", etError: Exit Sub
  'vk.4 RfId changed as Bag Alias No.
  If LTrim(ADC("wRfIdSel")) = "" Then DispMsg "Enter Bag Alias No. List / Stock ID Nos List", etError: Exit Sub
  pb_CpyRfId = True
  Call CpyStkID
  pb_CpyRfId = False
  Call CmdCpyRfId_Click
  
End Sub

Private Sub CmdGo_Click()
'calling the procedure for stock reconciliation
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
    'Check whether the Current Voucher is already Re-Conciled
    If moCn.RecSeek("Select 'x' From BgRecoMst Where " + _
                    " BmCoCd = '" + ADC("BmCoCd") + "' And BmTc= '" + ADC("BmTc") + "' And" + _
                    " BmYy = '" + ADC("BmYy") + "' And BmChr= '" + ADC("BmChr") + "' And BmNo= " + CStr(ADC("BmNo")) + " And BmRecoYN='Y'") Then
      DispMsg "Voucher already Reconciled, Cannot Reconcile again", etError: Exit Sub
    End If
    'Check whether Company code & Loc code is blank
    If ADC("BmBagCoCdFr") = "" And ADC("BmBagCoCdTo") = "" And ADC("BmBagCoCdSel") = "" Then
      DispMsg "Enter Stock Company Scope Fields", etError: Exit Sub
    End If
    If ADC("BmBagLocFr") = "" And ADC("BmBagLocTo") = "" And ADC("BmBagLocSel") = "" Then
      DispMsg "Enter Stock Location Scope Fields", etError: Exit Sub
    End If

    FraMsg.Visible = True
    FraMsg.Enabled = True
    FraMsg.ZOrder
    Call EnaDisaCmds(True, CmdReco)
    CmdReco.SetFocus
  End If
    
End Sub

Private Sub GrdStkReco_ColChange(ByVal RowNum As Integer, ByVal ColName As String, ByVal NewValue As Variant, Cancel As Boolean, ErrMsg As String)
  Dim ws_OldYy As String, ws_OldChr As String, ws_OldNo As Single, ws_StkNo() As String, ws_StkNoBarCd As String
  With GrdStkReco
    Select Case ColName
    Case Is = UCase("BdBagCoCd")
      'check for valid company code
'SRICHANGE
      'Cancel = (NewValue = "ZZZ") Or (Not moCn.RecSeek("Select 'x' from Head where HCoCd= '" + NewValue + "'")) Or _
                IIF(ADC("BmBagCoCdFr") = "" And ADC("BmBagCoCdTo") = "", "False", (Not (NewValue >= ADC("BmBagCoCdFr") And NewValue <= ADC("BmBagCoCdTo")))) Or _
                IIF(ADC("BmBagCoCdSel") = "", "False", (Not ("," + ADC("BmBagCoCdSel") + ",") Like ("*," + NewValue + ",*")))
      Cancel = (NewValue = "ZZZ") Or (Not moCn.RecSeek("Select 'x' from Head where HCoCd= '" + NewValue + "'"))
      If Cancel = True Then ErrMsg = "Invalid Stock Company Code/not in Company Code range Selected above": Exit Sub
    
    Case Is = UCase("wStkNo")
      'if company code in Stock no., then it will overwrite the existing company code in grid.
      ws_StkNo = Split(NewValue, "/")
      If UBound(ws_StkNo) = 3 Then .Value(RowNum, "BdBagCoCd") = ws_StkNo(0)
      
      'calling function to ignore CoCd if entered
      ws_StkNoBarCd = ParseBagNoBarCd(NewValue)
      ws_StkNo = Split(ws_StkNoBarCd, "/")

      If UBound(ws_StkNo) <> 2 Then Cancel = True: ErrMsg = "Invalid Stock No": Exit Sub
      If Not IsNumeric(ws_StkNo(2)) Then Cancel = True: ErrMsg = "Invalid Stock No": Exit Sub
      ws_OldYy = .Value(RowNum, "BdBYy")
      ws_OldChr = .Value(RowNum, "BdBChr")
      ws_OldNo = .Value(RowNum, "BdBNo")

      .Value(RowNum, "BdBYy") = ws_StkNo(0)
      'SRICHANGE
'      Call GrdStkReco_ColChange(RowNum, "BDBYY", ws_StkNo(0), Cancel, "")
'      If Cancel = True Then ErrMsg = "Invalid Stock Year": .Value(RowNum, "BdBYy") = ws_OldYy: Exit Sub

      .Value(RowNum, "BdBChr") = UCase(ws_StkNo(1))
      'SRICHANGE
 '     Call GrdStkReco_ColChange(RowNum, "BDBCHR", ws_StkNo(1), Cancel, "")
 '     If Cancel = True Then ErrMsg = "Invalid Stock Character": .Value(RowNum, "BdBChr") = ws_OldChr: Exit Sub

      .Value(RowNum, "BdBNo") = ws_StkNo(2)
'      SRICHANGE
'      Call GrdStkReco_ColChange(RowNum, "BDBNO", ws_StkNo(2), Cancel, "")
'      If Cancel = True Then ErrMsg = "Invalid Stock Number": .Value(RowNum, "BdBNo") = ws_OldNo: Exit Sub
    
    'check for valid year, Chr, and stk no
    Case Is = UCase("BdBYy")
      Cancel = (NewValue <> "") And (Not moCn.RecSeek("Select PMCd From Param Where PTyp= 'YY' And " + _
                                     "PMCd = '" + NewValue + "'"))
    Case Is = UCase("BdBChr")
      Cancel = (NewValue <> "") And (Not moCn.RecSeek("Select PMCd From Param Where PTyp= 'BCHR' " + _
                                     "And PMCd = '" + NewValue + "' "))
    Case Is = UCase("BdBNo")
      'SRICHANGE
      Cancel = (NewValue <> 0) And (Not moCn.RecSeek("Select 'a' where exists (Select 'a' from DmYyMm Where " + _
                                      "DdmStkCoCd= '" + .Value(RowNum, "BdBagCoCd") + "' And " + _
                                      "DDmStkYy= '" + .Value(RowNum, "BdBYy") + "' And " + _
                                      "DDmStkChr = '" + .Value(RowNum, "BdBChr") + "' And " + _
                                      "DDmStkNo= " + CStr(NewValue) + ")"))

    End Select
  End With
  
End Sub
Private Sub GrdStkReco_RowWhen(ByVal RowNum As Integer)
  'Generating the Serial Number in the Grid
  With GrdStkReco
    If .IsNew(RowNum) And .Value(RowNum, "BdSr") = 0 Then
      .Value(RowNum, "BdSr") = .MaxVal("BdSr") + 1
    End If
  End With

End Sub

Private Sub GrdStkReco_ColWhen(ByVal RowNum As Integer, ByVal ColName As String, Cancel As Boolean, ErrMsg As String, ByVal Action As MwfCtl.en_fgAction)
  'if Voucher already Reconciled cannot change any record
  If UCase(ADC("BmRecoYN")) = "Y" Then
    Cancel = True: ErrMsg = "Voucher already Reconciled, Cannot Edit": Exit Sub
  End If

  'company and loc fields cannot be blank
  If ADC("BmBagCoCdFr") = "" And ADC("BmBagCoCdTo") = "" And ADC("BmBagCoCdSel") = "" Then
      Cancel = True: ErrMsg = "Enter Stock Company scope Fields": Exit Sub
      'srichange
      If Cancel Then MsgBox ErrMsg
  End If
  If ADC("BmBagLocFr") = "" And ADC("BmBagLocTo") = "" And ADC("BmBagLocSel") = "" Then
      Cancel = True: ErrMsg = "Enter Stock Location scope Fields": Exit Sub
      If Cancel Then MsgBox ErrMsg
  End If

  With GrdStkReco
  '*** Set Helps For Each Field
    Select Case ColName
    Case UCase("BdBagCoCd")
      If .Value(RowNum, "BdBagCoCd") = "" Then
        If RowNum = 1 Then .Value(RowNum, "BdBagCoCd") = ADC("BmCoCd")
        If RowNum > 1 Then .Value(RowNum, "BdBagCoCd") = .Value(RowNum - 1, "BdBagCoCd")
      End If

      'Cannot edit company code once entered, User will have to delete and then add it again
      If moCn.RecSeek("Select 'x' From BgRecoDet Where BdCoCd= '" + gs_CoCd + "' and " + _
         "BdTc= '" + ADC("BmTc") + "' and BdYy= '" + ADC("BmYy") + "' and " + _
         "BdChr= '" + ADC("BmChr") + "' and BdNo= " + CStr(ADC("BmNo")) + " and " + _
         "BdSr= " + CStr(GrdStkReco.Value(RowNum, "BdSr"))) Then _
      Cancel = True: ErrMsg = "Cannot Edit This Value once Stock details have been entered": Exit Sub
      Call HlpList.hCoCd

    Case UCase("wStkNo")
      'Cannot edit stock No. once entered, User will have to delete and then add it again
      If moCn.RecSeek("Select 'x' From BgRecoDet Where BdCoCd= '" + gs_CoCd + "' and " + _
         "BdTc= '" + ADC("BmTc") + "' and BdYy= '" + ADC("BmYy") + "' and " + _
         "BdChr= '" + ADC("BmChr") + "' and BdNo= " + CStr(ADC("BmNo")) + " and " + _
         "BdSr= " + CStr(GrdStkReco.Value(RowNum, "BdSr"))) Then _
      Cancel = True: ErrMsg = "Cannot Edit This Value once Stock details have been entered": Exit Sub

    End Select
  End With
End Sub
 
Private Sub GrdStkReco_ReCalc(ByVal RowNum As Integer, ByVal ColName As String, ByVal OldVal As Variant, ByVal SourceName As String, ByVal SourceColName As String)
  Dim wi_Row As Integer
  With GrdStkReco
    wi_Row = .Row
    Select Case ColName
    Case Is = UCase("BdDmCd")
      If .Mode = fgmnorm Then
      If .IsNew(RowNum) Then
        'SRICHANGE -> it is in NOTES comments as to why
        .Value(RowNum, "BdDmCd") = moCn.GetFldVal("Select top 1 DDmCd From DmYyMm Where " + IIF(gs_Partition = ctCurrPrtn, " DPrtKey='" + ctCurrPrtn + "' And ", "") + _
                                        " DdmStkCoCd= '" + .Value(RowNum, "BdBagCoCd") + "' and " + _
                                        " DdmStkYy= '" + .Value(RowNum, "BdBYy") + "' And " + _
                                        " DdmStkChr = '" + .Value(RowNum, "BdBChr") + "' And " + _
                                        " DdmStkNo= " + CStr(.Value(RowNum, "BdBNo")))
      End If
      End If
    Case Is = UCase("BdActualQty")
      If .Mode = fgmnorm Then
      If .IsNew(RowNum) Then
          .Value(RowNum, "BdActualQty") = 1
      End If
      End If

    End Select
  End With
  
End Sub
Private Sub GrdStkReco_InitKey(ByVal RowNum As Integer)
  '*** Initialise Key Fields and all the fields which are not included in the grid
  With GrdStkReco
    If .IsNew(RowNum) Then
      .Store "BdCoCd", ADC("BmCoCd")
      .Store "BdTc", ADC("BmTc")
      .Store "BdYy", ADC("BmYy")
      .Store "BdChr", ADC("BmChr")
      .Store "BdNo", ADC("BmNo")
      .Store "BdGrWt", 0
  '    .Store "BdOthLocBg", ""
      .Store "BdOthLoc", ""
    End If
  End With

End Sub

Private Sub ADC_SetRecSource()
  '*** Set the RecSource for GrdStkReco
  ADC.RecSource = " Select * from BgRecoMst " + _
                       " Where BmCoCd = '" + ADC("BmCoCd") + "'" + _
                       " And BmTc ='" + ADC("BmTc") + "'" + _
                       " And BmYy = '" + ADC("BmYy") + "'" + _
                       " And BmChr ='" + ADC("BmChr") + "' " + _
                       " And BmNo =" + CStr(ADC("BmNo"))

End Sub

Private Sub GrdStkReco_SetRecSource()
  '*** Set the RecSource for GrdStkReco (Table BgReco)
  GrdStkReco.RecSource = " Select * from BgRecoDet " + _
                       " Where BdCoCd = '" + ADC("BmCoCd") + "'" + _
                       " And BdTc ='" + ADC("BmTc") + "'" + _
                       " And BdYy = '" + ADC("BmYy") + "'" + _
                       " And BdChr ='" + ADC("BmChr") + "' " + _
                       " And BdNo =" + CStr(ADC("BmNo")) + _
                       " Order By BdSr"
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
      pv_CmdBut.CausesValidation = True
    End If
  End If
  FraNKeyAll.Enabled = Not pv_ShowFra
  ADC.AllowSave = Not pv_ShowFra
  If Not pv_CmdBut Is Nothing Then pv_CmdBut.Enabled = True

End Sub

Private Sub CpyStkID()
  Dim wCnd As String, wRfIdStr As String, wRfIdNoArr() As String, wRfIdNo As String
  Dim wMinSr As Integer, wRow As Integer, wStkNoCoCd() As String, wRfIdStrCoCd As String, i As Integer, wTotRfid As Integer
  wCnd = ""
  'copy thru RfId/ stk Id from
  If pb_CpyRfId Then
    Dim wInvalRfIdStr As String
    wInvalRfIdStr = ""
    'vk.4 calling function to replace enter with ,
    wRfIdStr = PrepareRfIdString(ADC("wRfIdSel"))
        
    Dim wLstSr As Integer
    'checking valid Rfid/ stock id
    If ADC("wRfIdSel") <> "" Then
      wRfIdNoArr = Split(wRfIdStr, ",")
      If Left(wRfIdStr, 1) = "," Or Right(wRfIdStr, 1) = "," Then
        'vk.4 Tag Input changed as Bag Alias No.
        DispMsg "Invalid Bag Alias No. / Stock ID Nos List", etError: Exit Sub
      End If
      wRfIdStr = ""
      For i = 0 To UBound(wRfIdNoArr)
        wRfIdNo = wRfIdNoArr(i)
        wStkNoCoCd = Split(wRfIdNo, "/")
        wRfIdNo = ParseBagNoBarCd(wRfIdNo)
              
        If ADC("wCpyThruStkNo") = "N" Then
          If Not IsNumeric(wRfIdNo) Then
              'vk.4 Tag changed as Bag Alias No.
              DispMsg "Bag Alias No. Not Numeric", etError
              MsgBox "Bag Alias No. Not Numeric"
              Exit Sub
          End If
          If Not moCn.RecSeek("Select 'x' from RfIdMst Where RimNo=" + wRfIdNo + " ") Then
            wInvalRfIdStr = wInvalRfIdStr + IIF(wInvalRfIdStr = "", "", ",") + wRfIdNo
          End If
        Else
          Dim wStkCoCd As String
          Dim wStkNoCoCdLen As Integer
          wStkNoCoCdLen = UBound(wStkNoCoCd)
          wStkCoCd = IIF(wStkNoCoCdLen = 3, wStkNoCoCd(0), ADC("BmCoCd"))
          If Not moCn.RecSeek("Select 'x' from DmYyMm Where DdmStkCoCd = '" + wStkCoCd + "' and DDmStkYy+'/'+DDmStkChr+'/'+ltrim(RTrim(str(DDmStkNo))) = '" + wRfIdNo + "'") Then
            wInvalRfIdStr = wInvalRfIdStr + IIF(wInvalRfIdStr = "", "", ",") + wRfIdNo
          Else
              wMinSr = GrdStkReco.MaxVal("BdSr") + 1
              GrdStkReco.AddItem
              wRow = GrdStkReco.Rows - 1
              GrdStkReco.Value(wRow, "BdSr") = wMinSr
              GrdStkReco.Value(wRow, "BdBYy") = wStkNoCoCd(wStkNoCoCdLen - 2)
              GrdStkReco.Value(wRow, "BdBagCoCd") = wStkCoCd
              GrdStkReco.Value(wRow, "BdBChr") = wStkNoCoCd(wStkNoCoCdLen - 1)
              GrdStkReco.Value(wRow, "BdBNo") = wStkNoCoCd(wStkNoCoCdLen)
              wLstSr = wLstSr + 1
              GrdStkReco.SaveRec (wRow)
          End If
          wRfIdStrCoCd = wRfIdStrCoCd + IIF(wRfIdStrCoCd = "", "", ",") + IIF(UBound(wStkNoCoCd) = 3, wStkNoCoCd(0), ADC("BmCoCd")) + "/" + wRfIdNo
        End If
        wRfIdStr = wRfIdStr + IIF(wRfIdStr = "", "", ",") + wRfIdNo
              
      Next
      
      wTotRfid = UBound(wRfIdNoArr) + 1
      If wInvalRfIdStr <> "" Then
            DispMsg "Processing Done. These DmStkNum were not copied:" + wInvalRfIdStr, etWarning
            MsgBox "Processing Done. These DmStkNum were not copied:" + wInvalRfIdStr
            Exit Sub
      End If
      
'      Dim wCoCnd As String
'      SRICHANGE
'      If ADC("BmBagCoCdFr") <> "" Then wCoCnd = wCoCnd + IIF(wCoCnd <> "", " and ", "") + "DCoCd >='" + ADC("BmBagCoCdFr") + "'"
'      If ADC("BmBagCoCdTo") <> "" Then wCoCnd = wCoCnd + IIF(wCoCnd <> "", " and ", "") + "DCoCd <='" + ADC("BmBagCocdTo") + "'"
'      If ADC("BmBagCocdSel") <> "" Then wCoCnd = wCoCnd + IIF(wCoCnd <> "", " and ", "") + "'," + ADC("BmBagCocdSel") + ",' Like '%,'+DCoCd+',%' "
    
'      If LTrim(ADC("wRfIdSel")) <> "" And ADC("wCpyThruStkNo") = "N" Then _
'          wCnd = wCoCnd + " and Exists(Select 'x' From RfIdMst where RimNo in (" + ADC("wRfIdSel") + ") and DdmStkCoCd=RimBCoCd and DDmStkYy=RimBYy and DDmStkChr=RimBChr and DDmStkNo=RimBNo) " + _
'                  " group by DdmStkCoCd, DDmStkYy, DDmStkChr, DDmStkNo"
    
'      If LTrim(ADC("wRfIdSel")) <> "" And ADC("wCpyThruStkNo") = "Y" Then _
'          wCnd = wCoCnd + " and charindex(',' + DdmStkCoCd + '/' + DDmStkYy + '/' + DDmStkChr + '/' + LTrim(RTrim(str(DDmStkNo)))+',' , '," + wRfIdStrCoCd + ",') <> 0 " + _
'                  " group by DdmStkCoCd, DDmStkYy, DDmStkChr, DDmStkNo"
    End If
  Else
'   Copy Thru Stock Reco
    If ADC("BmBagCoCdFr") <> "" Then wCnd = "DdmStkCoCd >='" + ADC("BmBagCoCdFr") + "'"
    If ADC("BmBagCoCdTo") <> "" Then wCnd = wCnd + IIF(wCnd <> "", " and ", "") + "DdmStkCoCd <='" + ADC("BmBagCoCdTo") + "'"
    If ADC("BmBagCoCdSel") <> "" Then wCnd = wCnd + IIF(wCnd <> "", " and ", "") + "'," + ADC("BmBagCoCdSel") + ",' Like '%,'+DdmStkCoCd+',%' "
    If ADC("BmBagLocFr") <> "" Then wCnd = wCnd + IIF(wCnd <> "", " and ", "") + "DLoc >='" + ADC("BmBagLocFr") + "'"
    If ADC("BmBagLocTo") <> "" Then wCnd = wCnd + IIF(wCnd <> "", " and ", "") + "DLoc <='" + ADC("BmBagLocTo") + "'"
    If ADC("BmBagLocSel") <> "" Then wCnd = wCnd + IIF(wCnd <> "", " and ", "") + "'," + ADC("BmBagLocSel") + ",' Like '%,'+DLoc+',%' "

    wCnd = wCnd + IIF(wCnd <> "", " and ", "") + " DCoCd='" + gs_CoCd + "' and DdmStkCoCd<>'' and DDmStkYy<>'' and DDmStkChr<>'' and DDmStkNo<>0 and " + _
                    " Not Exists (Select 'x' From BgRecoDet " + _
                    " Join BgRecoMst On BdCoCd=BmCocd And BdTc=BmTc And BdYy=BmYy And BdChr=BmChr And BdNo=BmNo " + _
                    " Where BdCoCd='" + ADC("BmCoCd") + "' And BdTc='" + ADC("BmTc") + "' And BdYy='" + ADC("Bmyy") + "' And BdChr='" + ADC("BmChr") + "' And BdNo=" + CStr(ADC("BmNo")) + _
                    " And BdBagCoCd=DdmStkCoCd And BdBYy=DDmStkYy And BdBChr=DDmStkChr And BdBNo=DDmStkNo) " + _
                    " group by DdmStkCoCd, DDmStkYy, DDmStkChr, DDmStkNo having sum(DDmDrQty - DDmCrQty) <> 0 "


    If pb_CpyRfId Then MWLib.BeginProcess Me, "Copying Stock Ids ..."
    Dim wRsStk As MwfLib.MDORowSet, wsActQty As Double
    Set wRsStk = moCn.OpenRes("Select DdmStkCoCd, DDmStkYy, DDmStkChr, DDmStkNo, max(DDmCd) as DDmCd,sum(DDmDrQty - DDmCrQty) as qStkQty from DmYyMm " + _
                            "where " + IIF(gs_Partition = ctCurrPrtn, " DPrtKey='" + ctCurrPrtn + "' And ", "") + wCnd)
        
    wLstSr = GrdStkReco.MaxVal("BdSr") + 1
       
    With wRsStk
      GrdStkReco.StartCopy
      wMinSr = GrdStkReco.MaxVal("BdSr") + 1
      Do While Not (.EOF Or .BOF)
          GrdStkReco.AddItem
          MWLib.RunningProcess "Adding Stock " + !DDmStkYy + "/" + !DDmStkChr + "/" + CStr(!DDmStkNo)
          wRow = GrdStkReco.Rows - 1

          GrdStkReco.Value(wRow, "BdSr") = wLstSr
          GrdStkReco.Value(wRow, "BdBYy") = !DDmStkYy
          GrdStkReco.Value(wRow, "BdBagCoCd") = !DDmStkCoCd
          GrdStkReco.Value(wRow, "BdBChr") = !DDmStkChr
          GrdStkReco.Value(wRow, "BdBNo") = !DDmStkNo
          GrdStkReco.Value(wRow, "BdActualQty") = 0
          GrdStkReco.Value(wRow, "BdBQty") = !qStkQty
          GrdStkReco.Value(wRow, "BdDmCd") = !DDmCd
          GrdStkReco.Value(wRow, "BdExtQty") = 0
          GrdStkReco.Value(wRow, "BdUnAcQty") = !qStkQty
          GrdStkReco.Value(wRow, "BdOthLocBg") = ""
          wLstSr = wLstSr + 1

          GrdStkReco.SaveRec (wRow)
          .MoveNext
      Loop

      MWLib.EndProcess Me
      GrdStkReco.EndCopy
    End With
    Set wRsStk = Nothing
  
  End If
    '*** End of the copy process
End Sub

Private Sub VchReco()
  '*** Voucher Reconciliation Routine ***
  Dim wReco As Single
  Dim moRs As MDORowSet
  Dim wCnd, wCoCnd As String
  wCnd = ""
  wCoCnd = ""
  
  If ADC("BmBagLocFr") <> "" Then wCnd = wCnd + IIF(wCnd <> "", " and ", "") + "DLoc >='" + ADC("BmBagLocFr") + "'"
  If ADC("BmBagLocTo") <> "" Then wCnd = wCnd + IIF(wCnd <> "", " and ", "") + "DLoc <='" + ADC("BmBagLocTo") + "'"
  If ADC("BmBagLocSel") <> "" Then wCnd = wCnd + IIF(wCnd <> "", " and ", "") + "'," + ADC("BmBagLocSel") + ",' Like '%,'+DLoc+',%' "

  If ADC("BmBagCoCdFr") <> "" Then wCoCnd = wCoCnd + IIF(wCoCnd <> "", " and ", "") + "DCoCd >='" + ADC("BmBagCoCdFr") + "'"
  If ADC("BmBagCoCdTo") <> "" Then wCoCnd = wCoCnd + IIF(wCoCnd <> "", " and ", "") + "DCoCd <='" + ADC("BmBagCocdTo") + "'"
  If ADC("BmBagCocdSel") <> "" Then wCoCnd = wCoCnd + IIF(wCoCnd <> "", " and ", "") + "'," + ADC("BmBagCocdSel") + ",' Like '%,'+DCoCd+',%' "
  
  '*** Beginning the copy process
  MWLib.BeginProcess Me, "Reconciling Voucher..."
  wReco = 0
  With GrdStkReco
    For wReco = 1 To (GrdStkReco.Rows - 1)
      Set moRs = moCn.OpenRes("Select sum(DDmDrQty - DDmCrQty) as qStkQty From DmYyMm JOIN Loc on LocCoCd = DCoCd and LocCd = DLoc  " + _
                          "Where " + IIF(gs_Partition = ctCurrPrtn, " DPrtKey='" + ctCurrPrtn + "' And ", "") + _
                          IIF(wCoCnd <> "", wCoCnd + " and ", "") + IIF(wCnd <> "", wCnd + " and ", "") + _
                          "DdmStkCoCd= '" + .Value(wReco, "BdBagCoCd") + "' And LocTyp='S' And " + _
                          "DDmStkYy= '" + .Value(wReco, "BdBYy") + "' And " + _
                          "DDmStkChr = '" + .Value(wReco, "BdBChr") + "' And " + _
                          "DDmStkNo= " + CStr(.Value(wReco, "BdBNo")))
            
  
      .Value(wReco, "BdBQty") = moRs.FldValue("qStkQty")
      'Calculation of Extra Qty and Un-Accounted Quantity
      If .Value(wReco, "BdActualQty") > .Value(wReco, "BdBQty") Then
        .Value(wReco, "BdExtQty") = .Value(wReco, "BdActualQty") - .Value(wReco, "BdBQty")
        .Value(wReco, "BdUnAcQty") = 0
      ElseIf .Value(wReco, "BdActualQty") < .Value(wReco, "BdBQty") Then
        .Value(wReco, "BdUnAcQty") = .Value(wReco, "BdBQty") - GrdStkReco.Value(wReco, "BdActualQty")
        .Value(wReco, "BdExtQty") = 0
      Else
        .Value(wReco, "BdExtQty") = 0
        .Value(wReco, "BdUnAcQty") = 0
      End If
      GrdStkReco.SaveRec (wReco)

    Next wReco
    Call CpyStkID
    ADC("BmRECOYN") = "Y"
    ADC.SaveRec

    'Displaying the Invisible columns with Read Only Property
    .ColProp("BdBQty").Style = fgcReadOnly
    .ColProp("BdExtQty").Style = fgcReadOnly
    .ColProp("BdUnAcQty").Style = fgcReadOnly

  End With
  MWLib.EndProcess Me

  DispMsg "Voucher Reconciliation Completed", etInfo
  
End Sub

Private Sub HideAllFras()
  FraNKeyAll.Visible = True
End Sub

Private Sub Form_GotFocus()
  '*** Jenny Color (For Setting System Colors)
  SetMDISysColorsAfterLogin
  '*** Jenny Color (For Setting System Colors)
End Sub
Private Sub Form_Load()
  Dim ws_BtnStr As String
  ws_BtnStr = "CmdGo"
  Call SetObjColors(Me, ws_BtnStr)
  Me.KeyPreview = True
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
    Set moCn = Nothing
End Sub
Private Sub GrdStkReco_Validate(Cancel As Boolean)
  Cancel = GrdStkReco.Validate
End Sub

'vk.4 Copy from Bag-Y , copy from Bag alias no. -N
Private Sub OptCpyBagFr_Click(Index As Integer)
  Select Case Index
  Case Is = 0
    ADC("wCpyThruStkNo") = "Y"
  Case Is = 1
    ADC("wCpyThruStkNo") = "N"
  End Select
End Sub
'vk.4 tooltip for copy bag options
Private Sub OptCpyBagFr_GotFocus(Index As Integer)
  DispMsg OptCpyBagFr(Index).ToolTipText, etInfo
End Sub

