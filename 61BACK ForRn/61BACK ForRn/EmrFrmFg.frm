VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "tabctl32.ocx"
Object = "{AB2CC32A-E652-42A2-B341-26ECE02F6386}#2.1#0"; "Mwfctlforspeed9.ocx"
Begin VB.Form EmrFrmFg 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Finished Goods Entry"
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
   Begin MwfCtl.MWCTL_BTN1 CmdCB 
      Height          =   465
      Left            =   5595
      TabIndex        =   85
      ToolTipText     =   "Copy Bags From One Loc To Another "
      Top             =   9300
      Width           =   855
      _ExtentX        =   1508
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
   Begin MwfCtl.MWCTL_BTN1 CmdCpyRfId 
      Height          =   465
      Left            =   8835
      TabIndex        =   39
      ToolTipText     =   "Copy thru Bag No / Bag Alias No."
      Top             =   9300
      Width           =   1500
      _ExtentX        =   2646
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
   Begin MwfCtl.MWCTL_BTN1 CmdInvHd 
      Height          =   435
      Left            =   4890
      TabIndex        =   178
      ToolTipText     =   "Enter invoice header details"
      Top             =   9330
      Width           =   705
      _ExtentX        =   1244
      _ExtentY        =   767
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
   Begin MwfCtl.MWCTL_BTN1 CmdSumm 
      Height          =   465
      Left            =   8010
      TabIndex        =   38
      ToolTipText     =   "Option To See The Summary Details"
      Top             =   9300
      Width           =   825
      _ExtentX        =   1455
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
   Begin MwfCtl.MWCTL_BTN1 CmdEX 
      CausesValidation=   0   'False
      Height          =   465
      Left            =   6420
      TabIndex        =   36
      ToolTipText     =   "Export Details"
      Top             =   9300
      Width           =   795
      _ExtentX        =   1402
      _ExtentY        =   820
      BackColor       =   -2147483633
      MaskColor       =   8454143
      Caption         =   "&Export Details"
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
      Left            =   7215
      TabIndex        =   37
      ToolTipText     =   "Copy Option"
      Top             =   9300
      Width           =   795
      _ExtentX        =   1402
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
      Left            =   9750
      TabIndex        =   40
      TabStop         =   0   'False
      Top             =   9330
      Visible         =   0   'False
      Width           =   555
      _ExtentX        =   979
      _ExtentY        =   714
   End
   Begin MwfCtl.MWCTL_ADC ADC 
      CausesValidation=   0   'False
      Height          =   555
      Left            =   30
      TabIndex        =   41
      Top             =   9300
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
      TabIndex        =   43
      Top             =   30
      Width           =   4080
      Begin MwfCtl.MWCTL_MED ATXT 
         Height          =   285
         Index           =   1
         Left            =   960
         TabIndex        =   0
         ToolTipText     =   "Enter Voucher Tc"
         Top             =   0
         Width           =   570
         _ExtentX        =   1005
         _ExtentY        =   503
         MaxLength       =   3
         DataType        =   4
         DataField       =   "FgTc"
         ReCalcParent    =   "FGCHR"
         IdName          =   "FGTC"
      End
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
         DataField       =   "FgNo"
         IdName          =   "FGNO"
      End
      Begin MwfCtl.MWCTL_MED ATXT 
         Height          =   285
         Index           =   5
         Left            =   2160
         TabIndex        =   2
         ToolTipText     =   "Enter Voucher Character"
         Top             =   0
         Width           =   570
         _ExtentX        =   1005
         _ExtentY        =   503
         MaxLength       =   3
         DataType        =   4
         DataField       =   "FgChr"
         ReCalcParent    =   "FGFRBLOC,FGFRRMLOC,FGTORMLOC"
         ReCalcOn        =   "FGTC"
         IdName          =   "FGCHR"
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
         DataField       =   "FgYy"
         IdName          =   "FGYY"
      End
      Begin MwfCtl.MWCTL_MED ATXT 
         Height          =   285
         Index           =   31
         Left            =   3780
         TabIndex        =   78
         Top             =   0
         Visible         =   0   'False
         Width           =   135
         _ExtentX        =   238
         _ExtentY        =   503
         Enabled         =   0   'False
         MaxLength       =   3
         DataType        =   4
         DataField       =   "FgCoCd"
         IdName          =   "FGCOCD"
      End
      Begin MwfCtl.MWCTL_MED ATXT 
         Height          =   285
         Index           =   37
         Left            =   3810
         TabIndex        =   90
         Top             =   0
         Visible         =   0   'False
         Width           =   255
         _ExtentX        =   450
         _ExtentY        =   503
         Enabled         =   0   'False
         DataField       =   "FgPrtKey"
         IdName          =   "FGPRTKEY"
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
         TabIndex        =   51
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
         TabIndex        =   50
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
         TabIndex        =   46
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
         TabIndex        =   44
         Top             =   -5
         Width           =   1065
      End
   End
   Begin VB.Frame FRA_NKEYS 
      BorderStyle     =   0  'None
      Enabled         =   0   'False
      Height          =   9810
      Left            =   -120
      TabIndex        =   42
      Top             =   0
      Width           =   15315
      Begin VB.Frame FraNKeyAll 
         BorderStyle     =   0  'None
         Caption         =   "- "
         Height          =   9585
         Left            =   150
         TabIndex        =   45
         Top             =   30
         Width           =   15165
         Begin VB.Frame FraBagScaleWt 
            Height          =   600
            Left            =   12960
            TabIndex        =   179
            Top             =   6960
            Width           =   2175
            Begin MwfCtl.MWCTL_MED ATXT 
               Height          =   285
               Index           =   53
               Left            =   930
               TabIndex        =   180
               ToolTipText     =   "Scale Weight"
               Top             =   165
               Width           =   1140
               _ExtentX        =   2011
               _ExtentY        =   503
               Alignment       =   2
               Enabled         =   0   'False
               Mask            =   "-#####0.000"
               MaxLength       =   11
               DataType        =   2
               ReCalcOn        =   "GrdFgd(WSCALEWT)"
               IdName          =   "WBAGSCALEWT"
            End
            Begin VB.Label LblBagScaleWt 
               BackStyle       =   0  'Transparent
               Caption         =   "Scale Wt"
               BeginProperty Font 
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
               TabIndex        =   181
               Top             =   165
               Width           =   975
            End
         End
         Begin MwfCtl.MWCTL_FLX GrdFgRm 
            Height          =   2220
            Left            =   15
            TabIndex        =   16
            ToolTipText     =   "Raw Material Details"
            Top             =   6990
            Width           =   12825
            _ExtentX        =   22622
            _ExtentY        =   3916
            Cols            =   16
            colname1        =   "FRSRNO"
            heading1        =   "Sr.  "
            datafld1        =   "FrSrNo"
            datatype1       =   1
            mask1           =   "##0"
            maxlength1      =   3
            style1          =   2
            tooltiptext1    =   "Enter Sr No."
            colname2        =   "FRFRRMLOC"
            heading2        =   "Fr Loc                   "
            datafld2        =   "FrFrRmLoc"
            datatype2       =   4
            recalcparent2   =   "GrdFgd(FDGRWT)"
            maxlength2      =   8
            tooltiptext2    =   "Enter Raw Material From Location"
            colname3        =   "WFRFRRMIR"
            heading3        =   "I/R "
            datatype3       =   4
            recalcon3       =   "GrdFgRm(FRFRRMDC)"
            recalcparent3   =   "GrdFgRm(FRFRRMDC)"
            maxlength3      =   1
            tooltiptext3    =   "Enter Issue Or Receipt"
            colname4        =   "FRFRRMDC"
            datafld4        =   "FrFrRmDc"
            datatype4       =   4
            recalcon4       =   "GrdFgRm(WFRFRRMIR)"
            recalcparent4   =   "GrdFgRm(WFRFRRMIR),GrdFgd(FDGRWT)"
            maxlength4      =   1
            style4          =   0
            tooltiptext4    =   "Enter From Debit Or Credit (D/C)"
            colname5        =   "FRRMCD"
            heading5        =   "Rm Code                                 "
            datafld5        =   "FrRmCd"
            datatype5       =   4
            recalcparent5   =   "GrdFgRm(FRRMSZ),GrdFgRm(FRRMSTKRT),GrdFgRm(FRRMQTY),GrdFgRm(FRRMWT),GrdFgd(FDGRWT),GrdFgRm(FRLOTNO)"
            maxlength5      =   16
            tooltiptext5    =   "Enter Raw Material Code"
            colname6        =   "FRLOTNO"
            heading6        =   "Lot No.                                   "
            datafld6        =   "FrLotNo"
            datatype6       =   4
            recalcon6       =   "GrdFgRm(FRRMCD)"
            maxlength6      =   16
            tooltiptext6    =   "Enter Lot. No."
            colname7        =   "FRRMSZ"
            heading7        =   "Size          "
            datafld7        =   "FrRmSz"
            datatype7       =   2
            mask7           =   "##0.0000"
            recalcon7       =   "GrdFgRm(FRRMCD)"
            maxlength7      =   8
            tooltiptext7    =   "Enter Raw Material Size"
            colname8        =   "FRRMSZ2"
            heading8        =   "Brdth   "
            datafld8        =   "FrRmSz2"
            datatype8       =   2
            mask8           =   "#0.000"
            maxlength8      =   6
            colname9        =   "FRRMSZ3"
            heading9        =   "Dpth   "
            datafld9        =   "FrRmSz3"
            datatype9       =   2
            mask9           =   "#0.000"
            maxlength9      =   6
            colname10       =   "FRRMSTKRT"
            heading10       =   "Stock Rate       "
            datafld10       =   "FrRmStkRt"
            datatype10      =   2
            mask10          =   "-######0.000"
            recalcon10      =   "GrdFgRm(FRRMCD)"
            maxlength10     =   12
            tooltiptext10   =   "Enter Raw Material Stock Rate"
            colname11       =   "FRRMQTY"
            heading11       =   "Qty      "
            datafld11       =   "FrRmQty"
            datatype11      =   1
            mask11          =   "####0"
            recalcon11      =   "GrdFgRm(FRRMCD)"
            maxlength11     =   5
            tooltiptext11   =   "Enter Raw Material Quantity"
            colname12       =   "FRRMWT"
            heading12       =   "Rm Wt         "
            datafld12       =   "FrRmWt"
            datatype12      =   2
            mask12          =   "####0.000"
            recalcon12      =   "GrdFgRm(FRRMCD)"
            recalcparent12  =   "GrdFgd(FDGRWT)"
            maxlength12     =   9
            tooltiptext12   =   "Enter Raw Material Weight"
            colname13       =   "FRTORMLOC"
            heading13       =   "To Loc                    "
            datafld13       =   "FrToRmLoc"
            datatype13      =   4
            recalcparent13  =   "GrdFgd(FDGRWT)"
            maxlength13     =   8
            tooltiptext13   =   "Enter To Location"
            colname14       =   "FRTORMDC"
            datafld14       =   "FrToRmDc"
            datatype14      =   4
            maxlength14     =   1
            style14         =   0
            tooltiptext14   =   "Enter To Debit Or Credit (D/C)"
            colname15       =   "FRPRTKEY"
            datafld15       =   "FrPrtKey"
            maxlength15     =   1
            style15         =   0
         End
         Begin VB.TextBox TxtPwd 
            Height          =   285
            IMEMode         =   3  'DISABLE
            Left            =   10920
            MaxLength       =   20
            PasswordChar    =   "*"
            TabIndex        =   11
            Top             =   285
            Width           =   1290
         End
         Begin MwfCtl.MWCTL_FLX GrdFgd 
            Height          =   6105
            Left            =   0
            TabIndex        =   15
            ToolTipText     =   "Design Details"
            Top             =   600
            Width           =   15135
            _ExtentX        =   26696
            _ExtentY        =   10769
            Cols            =   42
            colname1        =   "FDSR"
            heading1        =   "Sr.  "
            datafld1        =   "FdSr"
            datatype1       =   1
            mask1           =   "##0"
            maxlength1      =   3
            style1          =   2
            tooltiptext1    =   "Enter Sr. No"
            colname2        =   "WSCALEWT"
            heading2        =   " "
            datatype2       =   2
            mask2           =   "-#####0.000"
            recalcparent2   =   "WBAGSCALEWT"
            maxlength2      =   11
            style2          =   0
            tooltiptext2    =   "Enter Scale Wt"
            colname3        =   "WFGBNO"
            datatype3       =   4
            maxlength3      =   18
            tooltiptext3    =   "Enter Bag Number Using Barcode Reader"
            colname4        =   "FDBYY"
            heading4        =   "BYy"
            datafld4        =   "FdBYy"
            datatype4       =   4
            recalcparent4   =   $"EmrFrmFg.frx":0000
            maxlength4      =   2
            tooltiptext4    =   "Enter Bag Year"
            colname5        =   "FDBCHR"
            heading5        =   "BChr"
            datafld5        =   "FdBChr"
            datatype5       =   4
            recalcparent5   =   $"EmrFrmFg.frx":010E
            maxlength5      =   3
            tooltiptext5    =   "Enter Bag Character"
            colname6        =   "FDBNO"
            heading6        =   "BNo.        "
            datafld6        =   "FdBNo"
            datatype6       =   1
            mask6           =   "######0"
            recalcon6       =   "GrdFgd(FDBYY),GrdFgd(FDBCHR)"
            recalcparent6   =   $"EmrFrmFg.frx":021C
            maxlength6      =   7
            tooltiptext6    =   "Enter Bag No. (Press Shift + Tab To Come Out Of This Field)"
            colname7        =   "WFDBASEVCHNO"
            heading7        =   "Base Voucher                         "
            datatype7       =   4
            recalcon7       =   "GrdFgd(FDBYY),GrdFgd(FDBCHR),GrdFgd(FDBNO),GrdFgd(FDREFYY),GrdFgd(FDREFKEY)"
            recalcparent7   =   $"EmrFrmFg.frx":031C
            maxlength7      =   22
            tooltiptext7    =   "Enter Base Voucher Number"
            colname8        =   "FDDMCD"
            heading8        =   "Design Cd                     "
            datafld8        =   "FdDmCd"
            datatype8       =   4
            recalcon8       =   "GrdFgd(FDBYY),GrdFgd(FDBCHR),GrdFgd(FDBNO),GrdFgd(WFDBASEVCHNO)"
            maxlength8      =   15
            tooltiptext8    =   "Enter Design Code"
            colname9        =   "FDSFX"
            heading9        =   "Suffix                 "
            datafld9        =   "FdSfx"
            datatype9       =   4
            recalcon9       =   "GrdFgd(FDBYY),GrdFgd(FDBCHR),GrdFgd(FDBNO),GrdFgd(WFDBASEVCHNO)"
            maxlength9      =   15
            tooltiptext9    =   "Enter Design Suffix"
            colname10       =   "FDDMSZ"
            heading10       =   "Size      "
            datafld10       =   "FdDmSz"
            datatype10      =   4
            recalcon10      =   "GrdFgd(FDBYY),GrdFgd(FDBCHR),GrdFgd(FDBNO),GrdFgd(WFDBASEVCHNO)"
            maxlength10     =   5
            tooltiptext10   =   "Enter Design Size"
            colname11       =   "FDQTY"
            heading11       =   "Qty    "
            datafld11       =   "FdQty"
            datatype11      =   2
            mask11          =   "##0.0"
            recalcon11      =   "GrdFgd(FDBYY),GrdFgd(FDBCHR),GrdFgd(FDBNO),GrdFgd(WFDBASEVCHNO)"
            recalcparent11  =   "WTOTQTY"
            maxlength11     =   5
            style11         =   0
            tooltiptext11   =   "Enter Bag Quantity"
            colname12       =   "FDGRWT"
            heading12       =   "GrossWt       "
            datafld12       =   "FdGrWt"
            datatype12      =   2
            mask12          =   "-#####0.000"
            recalcon12      =   $"EmrFrmFg.frx":04E0
            recalcparent12  =   "WTOTGRWT"
            maxlength12     =   11
            tooltiptext12   =   "Enter Bag Gross Weight"
            colname13       =   "FDPRDODTC"
            heading13       =   "OTc "
            datafld13       =   "FdPrdOdTc"
            datatype13      =   4
            recalcon13      =   "GrdFgd(FDBYY),GrdFgd(FDBCHR),GrdFgd(FDBNO),GrdFgd(WFDBASEVCHNO)"
            recalcparent13  =   "GrdFgd(FDEXPODTC)"
            maxlength13     =   3
            tooltiptext13   =   "Enter Order Tc"
            colname14       =   "FDPRDODYY"
            heading14       =   "OYy"
            datafld14       =   "FdPrdOdYy"
            datatype14      =   4
            recalcon14      =   "GrdFgd(FDBYY),GrdFgd(FDBCHR),GrdFgd(FDBNO),GrdFgd(WFDBASEVCHNO)"
            recalcparent14  =   "GrdFgd(FDEXPODYY)"
            maxlength14     =   2
            tooltiptext14   =   "Enter Order Year"
            colname15       =   "FDPRDODCHR"
            heading15       =   "OChr"
            datafld15       =   "FdPrdOdChr"
            datatype15      =   4
            recalcon15      =   "GrdFgd(FDBYY),GrdFgd(FDBCHR),GrdFgd(FDBNO),GrdFgd(WFDBASEVCHNO)"
            recalcparent15  =   "GrdFgd(FDEXPODCHR)"
            maxlength15     =   3
            tooltiptext15   =   "Enter Order Character"
            colname16       =   "FDPRDODNO"
            heading16       =   "ONo       "
            datafld16       =   "FdPrdOdNo"
            datatype16      =   1
            mask16          =   "#####0"
            recalcon16      =   "GrdFgd(FDBYY),GrdFgd(FDBCHR),GrdFgd(FDBNO),GrdFgd(WFDBASEVCHNO)"
            recalcparent16  =   "GrdFgd(FDEXPODNO)"
            maxlength16     =   6
            tooltiptext16   =   "Enter Order Number"
            colname17       =   "FDPRDODSR"
            heading17       =   "OSr    "
            datafld17       =   "FdPrdOdSr"
            datatype17      =   1
            mask17          =   "###0"
            recalcon17      =   "GrdFgd(FDBYY),GrdFgd(FDBCHR),GrdFgd(FDBNO),GrdFgd(WFDBASEVCHNO)"
            recalcparent17  =   "GrdFgd(FDEXPODSR)"
            maxlength17     =   4
            tooltiptext17   =   "Enter Order Sr No."
            colname18       =   "FDPRDCMCD"
            datafld18       =   "FdPrdCmCd"
            datatype18      =   4
            recalcon18      =   "GrdFgd(FDBYY),GrdFgd(FDBCHR),GrdFgd(FDBNO),GrdFgd(WFDBASEVCHNO)"
            recalcparent18  =   "GrdFgd(FDEXPCMCD)"
            maxlength18     =   8
            style18         =   0
            tooltiptext18   =   "Customer Code"
            colname19       =   "FDEXPCMCD"
            datafld19       =   "FdExpCmCd"
            datatype19      =   4
            recalcon19      =   "GrdFgd(FDPRDCMCD),GrdFgd(WFDBASEVCHNO)"
            maxlength19     =   8
            style19         =   0
            tooltiptext19   =   "Export Customer Code"
            colname20       =   "FDEXPODTC"
            datafld20       =   "FdExpOdTc"
            datatype20      =   4
            recalcon20      =   "GrdFgd(FDPRDODTC),GrdFgd(WFDBASEVCHNO)"
            maxlength20     =   3
            style20         =   0
            tooltiptext20   =   "Export Order Tc"
            colname21       =   "FDEXPODYY"
            datafld21       =   "FdExpOdYy"
            datatype21      =   4
            recalcon21      =   "GrdFgd(FDPRDODYY),GrdFgd(WFDBASEVCHNO)"
            maxlength21     =   2
            style21         =   0
            tooltiptext21   =   "Export Order Year"
            colname22       =   "FDEXPODCHR"
            datafld22       =   "FdExpOdChr"
            datatype22      =   4
            recalcon22      =   "GrdFgd(FDPRDODCHR),GrdFgd(WFDBASEVCHNO)"
            maxlength22     =   3
            style22         =   0
            tooltiptext22   =   "Export Order Character"
            colname23       =   "FDEXPODNO"
            datafld23       =   "FdExpOdNo"
            datatype23      =   1
            mask23          =   "#####0"
            recalcon23      =   "GrdFgd(FDPRDODNO),GrdFgd(WFDBASEVCHNO)"
            maxlength23     =   6
            style23         =   0
            tooltiptext23   =   "Export Order Number"
            colname24       =   "FDEXPODSR"
            datafld24       =   "FdExpOdSr"
            datatype24      =   1
            mask24          =   "###0"
            recalcon24      =   "GrdFgd(FDPRDODSR),GrdFgd(WFDBASEVCHNO)"
            maxlength24     =   4
            style24         =   0
            tooltiptext24   =   "Export Order Sr No."
            colname25       =   "FDINTC"
            datafld25       =   "FdInTc"
            datatype25      =   4
            recalcon25      =   "GrdFgd(WFDBASEVCHNO)"
            maxlength25     =   3
            style25         =   0
            tooltiptext25   =   "Invoice Tc"
            colname26       =   "FDINYY"
            datafld26       =   "FdInYy"
            datatype26      =   4
            recalcon26      =   "GrdFgd(WFDBASEVCHNO)"
            maxlength26     =   2
            style26         =   0
            tooltiptext26   =   "Invoice Year"
            colname27       =   "FDINCHR"
            datafld27       =   "FdInChr"
            datatype27      =   4
            recalcon27      =   "GrdFgd(WFDBASEVCHNO)"
            maxlength27     =   3
            style27         =   0
            tooltiptext27   =   "Invoice Character"
            colname28       =   "FDINNO"
            datafld28       =   "FdInNo"
            datatype28      =   1
            mask28          =   "######0"
            recalcon28      =   "GrdFgd(WFDBASEVCHNO)"
            maxlength28     =   7
            style28         =   0
            tooltiptext28   =   "Invoice Number"
            colname29       =   "FDINEXPNO"
            datafld29       =   "FdInExpNo"
            datatype29      =   4
            recalcon29      =   "GrdFgd(WFDBASEVCHNO)"
            maxlength29     =   20
            style29         =   0
            tooltiptext29   =   "Export Number"
            colname30       =   "FDDESPQTY"
            datafld30       =   "FdDespQty"
            datatype30      =   2
            mask30          =   "##0.0"
            maxlength30     =   5
            style30         =   0
            tooltiptext30   =   "Despatch Quantity"
            colname31       =   "FDKEY"
            datafld31       =   "FdKey"
            datatype31      =   1
            mask31          =   "#######0"
            maxlength31     =   8
            style31         =   0
            tooltiptext31   =   "Key"
            colname32       =   "FDREFYY"
            datafld32       =   "FdRefYy"
            datatype32      =   4
            recalcon32      =   "GrdFgd(WFDBASEVCHNO)"
            recalcparent32  =   "GrdFgd(WFDBASEVCHNO)"
            maxlength32     =   2
            style32         =   0
            tooltiptext32   =   "Reference Year"
            colname33       =   "FDREFKEY"
            datafld33       =   "FdRefKey"
            datatype33      =   1
            mask33          =   "#######0"
            recalcon33      =   "GrdFgd(WFDBASEVCHNO)"
            recalcparent33  =   "GrdFgd(WFDBASEVCHNO)"
            maxlength33     =   8
            style33         =   0
            tooltiptext33   =   "Reference Key"
            colname34       =   "FDBYYKEY"
            datafld34       =   "FdBYyKey"
            datatype34      =   4
            recalcon34      =   "GrdFgd(FDBYY),GrdFgd(FDBCHR),GrdFgd(FDBNO)"
            maxlength34     =   10
            style34         =   0
            tooltiptext34   =   "Enter Bag Yy & Key"
            colname35       =   "FGDESC"
            datafld35       =   "FdDesc"
            recalcon35      =   "WFGDESC"
            recalcparent35  =   "WFGDESC"
            maxlength35     =   60
            style35         =   0
            colname36       =   "FDSUBLOC"
            heading36       =   "SubLoc              "
            datafld36       =   "FdSubLoc"
            datatype36      =   4
            maxlength36     =   8
            tooltiptext36   =   "Enter FG Sub Loc"
            colname37       =   "FDCTBDESPQTY"
            datafld37       =   "FdCTBDespQty"
            datatype37      =   2
            mask37          =   "##0.0"
            maxlength37     =   5
            style37         =   0
            tooltiptext37   =   "CTB Despatch Qty"
            colname38       =   "FDPRTKEY"
            datafld38       =   "FdPrtKey"
            maxlength38     =   1
            style38         =   0
            colname39       =   "FDPLYY"
            heading39       =   "PlYy"
            datafld39       =   "FdPlYy"
            datatype39      =   4
            maxlength39     =   2
            tooltiptext39   =   "Packing List Year"
            colname40       =   "FDPLCHR"
            heading40       =   "PlChr"
            datafld40       =   "FdPlChr"
            datatype40      =   4
            maxlength40     =   3
            tooltiptext40   =   "Packing List Character"
            colname41       =   "FDPLNO"
            heading41       =   "PlNo          "
            datafld41       =   "FdPlNo"
            datatype41      =   1
            mask41          =   "######0"
            maxlength41     =   7
            tooltiptext41   =   "Packing List Number"
         End
         Begin MwfCtl.MWCTL_MED ATXT 
            Height          =   285
            Index           =   30
            Left            =   885
            TabIndex        =   4
            ToolTipText     =   "Enter Voucher Date"
            Top             =   285
            Width           =   855
            _ExtentX        =   1508
            _ExtentY        =   503
            MaxLength       =   8
            DataType        =   3
            DataField       =   "FgDt"
            IdName          =   "FGDT"
         End
         Begin MwfCtl.MWCTL_MED ATXT 
            Height          =   285
            Index           =   22
            Left            =   13245
            TabIndex        =   13
            ToolTipText     =   "Enter Raw Material Default To Location"
            Top             =   285
            Width           =   1410
            _ExtentX        =   2487
            _ExtentY        =   503
            MaxLength       =   5
            DataType        =   4
            DataField       =   "FgToRmLoc"
            ReCalcOn        =   "FGCHR"
            IdName          =   "FGTORMLOC"
         End
         Begin MwfCtl.MWCTL_MED ATXT 
            Height          =   285
            Index           =   2
            Left            =   5160
            TabIndex        =   5
            ToolTipText     =   "Enter Bag Location"
            Top             =   0
            Width           =   4530
            _ExtentX        =   7990
            _ExtentY        =   503
            MaxLength       =   60
            DataType        =   4
            DataField       =   "FgFrBLoc"
            ReCalcOn        =   "FGCHR"
            IdName          =   "FGFRBLOC"
         End
         Begin MwfCtl.MWCTL_MED ATXT 
            Height          =   285
            Index           =   3
            Left            =   14925
            TabIndex        =   14
            ToolTipText     =   "Enter From RM DC"
            Top             =   0
            Width           =   285
            _ExtentX        =   503
            _ExtentY        =   503
            DataType        =   4
            IdName          =   "WFRRMDC"
         End
         Begin MwfCtl.MWCTL_MED ATXT 
            Height          =   285
            Index           =   19
            Left            =   5160
            TabIndex        =   6
            ToolTipText     =   "Enter Finished Goods Location"
            Top             =   285
            Width           =   1365
            _ExtentX        =   2408
            _ExtentY        =   503
            MaxLength       =   8
            DataType        =   4
            DataField       =   "FgToBLoc"
            IdName          =   "FGTOBLOC"
         End
         Begin VB.Frame FraTot 
            Height          =   1200
            Left            =   12960
            TabIndex        =   70
            Top             =   7605
            Width           =   2175
            Begin MwfCtl.MWCTL_MED ATXT 
               Height          =   285
               Index           =   26
               Left            =   930
               TabIndex        =   71
               ToolTipText     =   "Total Quantiyt In The Bags"
               Top             =   690
               Width           =   780
               _ExtentX        =   1376
               _ExtentY        =   503
               Alignment       =   2
               Enabled         =   0   'False
               Mask            =   "####0.0"
               MaxLength       =   7
               DataType        =   2
               ReCalcOn        =   "GrdFgd(FDQTY)"
               IdName          =   "WTOTQTY"
            End
            Begin MwfCtl.MWCTL_MED ATXT 
               Height          =   285
               Index           =   29
               Left            =   930
               TabIndex        =   72
               ToolTipText     =   "Total Gross Weight"
               Top             =   405
               Width           =   1140
               _ExtentX        =   2011
               _ExtentY        =   503
               Alignment       =   2
               Enabled         =   0   'False
               Mask            =   "######0.000"
               MaxLength       =   11
               DataType        =   2
               ReCalcOn        =   "GrdFgd(FDGRWT)"
               IdName          =   "WTOTGRWT"
            End
            Begin VB.Label LblMain 
               BackStyle       =   0  'Transparent
               BorderStyle     =   1  'Fixed Single
               Caption         =   "  Total"
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
               TabIndex        =   86
               Top             =   90
               Width           =   2205
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
               Index           =   45
               Left            =   30
               TabIndex        =   74
               Top             =   405
               Width           =   975
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
               Index           =   43
               Left            =   30
               TabIndex        =   73
               Top             =   690
               Width           =   975
            End
         End
         Begin MwfCtl.MWCTL_MED ATXT 
            Height          =   285
            Index           =   32
            Left            =   660
            TabIndex        =   80
            ToolTipText     =   "Enter Product Description"
            Top             =   6720
            Width           =   11835
            _ExtentX        =   20876
            _ExtentY        =   503
            MaxLength       =   60
            DataType        =   4
            ReCalcParent    =   "GrdFgd(FGDESC)"
            ReCalcOn        =   "GrdFgd(FGDESC)"
            IdName          =   "WFGDESC"
         End
         Begin MwfCtl.MWCTL_MED ATXT 
            Height          =   285
            Index           =   33
            Left            =   10920
            TabIndex        =   10
            ToolTipText     =   "Enter Finished Goods Location"
            Top             =   0
            Width           =   1290
            _ExtentX        =   2275
            _ExtentY        =   503
            MaxLength       =   8
            DataType        =   4
            DataField       =   "FgSubLoc"
            IdName          =   "FGSUBLOC"
         End
         Begin MwfCtl.MWCTL_MED ATXT 
            Height          =   285
            Index           =   0
            Left            =   13250
            TabIndex        =   12
            ToolTipText     =   "Enter Raw Material Default From Location"
            Top             =   0
            Width           =   1410
            _ExtentX        =   2487
            _ExtentY        =   503
            MaxLength       =   5
            DataType        =   4
            DataField       =   "FgFrRmLoc"
            ReCalcOn        =   "FGCHR"
            IdName          =   "FGFRRMLOC"
         End
         Begin MwfCtl.MWCTL_MED ATXT 
            Height          =   285
            Index           =   36
            Left            =   10800
            TabIndex        =   88
            ToolTipText     =   "Enter Raw Material Default From Location"
            Top             =   240
            Visible         =   0   'False
            Width           =   150
            _ExtentX        =   265
            _ExtentY        =   503
            Enabled         =   0   'False
            MaxLength       =   5
            DataType        =   4
            DataField       =   "FGseopwd"
            IdName          =   "FGSEOPWD"
         End
         Begin MwfCtl.MWCTL_MED ATXT 
            Height          =   285
            Index           =   45
            Left            =   7800
            TabIndex        =   7
            ToolTipText     =   "Enter Packing List Year"
            Top             =   285
            Width           =   435
            _ExtentX        =   767
            _ExtentY        =   503
            MaxLength       =   2
            DataType        =   4
            IdName          =   "WFGPLYY"
         End
         Begin MwfCtl.MWCTL_MED ATXT 
            Height          =   285
            Index           =   47
            Left            =   8280
            TabIndex        =   8
            ToolTipText     =   "Enter Packing List Character"
            Top             =   285
            Width           =   570
            _ExtentX        =   1005
            _ExtentY        =   503
            MaxLength       =   3
            DataType        =   4
            ReCalcParent    =   "FGFRBLOC,FGFRRMLOC,FGTORMLOC"
            ReCalcOn        =   "FGTC"
            IdName          =   "WFGPLCHR"
         End
         Begin MwfCtl.MWCTL_MED ATXT 
            Height          =   285
            Index           =   48
            Left            =   8880
            TabIndex        =   9
            ToolTipText     =   "Enter Packing List Number"
            Top             =   285
            Width           =   930
            _ExtentX        =   1640
            _ExtentY        =   503
            Alignment       =   2
            Mask            =   "######0"
            MaxLength       =   7
            DataType        =   1
            IdName          =   "WFGPLNO"
         End
         Begin VB.Label LblPlNo 
            BackStyle       =   0  'Transparent
            Caption         =   "Packing List"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Left            =   6600
            TabIndex        =   125
            Top             =   285
            Width           =   1215
         End
         Begin VB.Label Label6 
            BackStyle       =   0  'Transparent
            Caption         =   "SeoPwd"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Left            =   9840
            TabIndex        =   89
            Top             =   240
            Width           =   870
         End
         Begin VB.Label LblFgSubLoc 
            BackStyle       =   0  'Transparent
            Caption         =   "BFgSubLoc"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Left            =   9720
            TabIndex        =   81
            Top             =   0
            Width           =   1200
         End
         Begin VB.Label LblFgDesc 
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
            Left            =   0
            TabIndex        =   79
            Top             =   6720
            Width           =   660
         End
         Begin VB.Label LblFgToBLoc 
            BackStyle       =   0  'Transparent
            Caption         =   "Bag To Loc"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Left            =   4020
            TabIndex        =   75
            Top             =   285
            Width           =   1230
         End
         Begin VB.Label ALBL 
            BackStyle       =   0  'Transparent
            Caption         =   "Dc"
            BeginProperty Font 
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
            Left            =   14640
            TabIndex        =   55
            Top             =   0
            Width           =   285
         End
         Begin VB.Label LblFgFrBLoc 
            BackStyle       =   0  'Transparent
            Caption         =   "Bag Fr Loc"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Left            =   4020
            TabIndex        =   52
            Top             =   0
            Width           =   1200
         End
         Begin VB.Label LblTToRmLoc 
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
            Left            =   12540
            TabIndex        =   49
            Top             =   285
            Width           =   675
         End
         Begin VB.Label LblTFrRmLoc 
            BackStyle       =   0  'Transparent
            Caption         =   "Rm Fr Loc"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Left            =   12240
            TabIndex        =   48
            Top             =   0
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
            Index           =   2
            Left            =   0
            TabIndex        =   47
            Top             =   240
            Width           =   855
         End
      End
      Begin VB.Frame FraSumm 
         Height          =   4815
         Left            =   1320
         TabIndex        =   97
         Top             =   4320
         Width           =   13455
         Begin TabDlg.SSTab TabSumm 
            Height          =   4455
            Left            =   60
            TabIndex        =   98
            Top             =   180
            Width           =   13365
            _ExtentX        =   23574
            _ExtentY        =   7858
            _Version        =   393216
            Tabs            =   2
            Tab             =   1
            TabsPerRow      =   2
            TabHeight       =   520
            TabCaption(0)   =   "General Summary"
            TabPicture(0)   =   "EmrFrmFg.frx":057B
            Tab(0).ControlEnabled=   0   'False
            Tab(0).Control(0)=   "ALBL(62)"
            Tab(0).Control(1)=   "ALBL(61)"
            Tab(0).Control(2)=   "ALBL(33)"
            Tab(0).Control(3)=   "ALBL(31)"
            Tab(0).Control(4)=   "ALBL(6)"
            Tab(0).Control(5)=   "ALBL(5)"
            Tab(0).Control(6)=   "ATXT(42)"
            Tab(0).Control(7)=   "ATXT(25)"
            Tab(0).Control(8)=   "ATXT(65)"
            Tab(0).Control(9)=   "ATXT(63)"
            Tab(0).Control(10)=   "ATXT(28)"
            Tab(0).Control(11)=   "ATXT(27)"
            Tab(0).Control(12)=   "GrdSumm"
            Tab(0).ControlCount=   13
            TabCaption(1)   =   "Raw Material Summary"
            TabPicture(1)   =   "EmrFrmFg.frx":0597
            Tab(1).ControlEnabled=   -1  'True
            Tab(1).Control(0)=   "ALBL(63)"
            Tab(1).Control(0).Enabled=   0   'False
            Tab(1).Control(1)=   "ALBL(64)"
            Tab(1).Control(1).Enabled=   0   'False
            Tab(1).Control(2)=   "LblMain(1)"
            Tab(1).Control(2).Enabled=   0   'False
            Tab(1).Control(3)=   "LblMain(5)"
            Tab(1).Control(3).Enabled=   0   'False
            Tab(1).Control(4)=   "ATXT(41)"
            Tab(1).Control(4).Enabled=   0   'False
            Tab(1).Control(5)=   "ATXT(24)"
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
               TabIndex        =   99
               Top             =   420
               Width           =   12285
               _ExtentX        =   21669
               _ExtentY        =   5689
               Cols            =   14
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
               heading7        =   "Inv No "
               maxlength7      =   22
               tooltiptext7    =   "Invoice Voucher No"
               colname8        =   "WSUMMINVEXPNO"
               heading8        =   "Inv Exp No                                    "
               datatype8       =   4
               maxlength8      =   20
               tooltiptext8    =   "Invoice Export Number"
               colname9        =   "WSUMMBAGCOUNT"
               heading9        =   "No Of Bags"
               datatype9       =   1
               mask9           =   "######0"
               maxlength9      =   7
               tooltiptext9    =   "Number Of Bags"
               colname10       =   "WSUMMQTY"
               heading10       =   "Bag Qty"
               datatype10      =   1
               mask10          =   "##0.0"
               maxlength10     =   5
               tooltiptext10   =   "Bag Quantity"
               colname11       =   "WSUMMWT"
               heading11       =   "Wt.            "
               datatype11      =   2
               mask11          =   "#####0.000"
               maxlength11     =   10
               tooltiptext11   =   "Weight"
               colname12       =   "WSUMMBALQTY"
               heading12       =   "Bal Qty                "
               datatype12      =   2
               mask12          =   "##0.0"
               maxlength12     =   5
               tooltiptext12   =   "Weight"
               colname13       =   "WSUMMBALWT"
               heading13       =   "Bal Wt"
               datatype13      =   2
               mask13          =   "####0.000"
               maxlength13     =   9
            End
            Begin MwfCtl.MWCTL_MED ATXT 
               Height          =   285
               Index           =   27
               Left            =   -68655
               TabIndex        =   100
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
               TabIndex        =   101
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
               TabIndex        =   102
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
               TabIndex        =   103
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
               TabIndex        =   104
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
               TabIndex        =   105
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
               TabIndex        =   106
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
               TabIndex        =   107
               Top             =   780
               Width           =   5745
               _ExtentX        =   10134
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
            End
            Begin MwfCtl.MWCTL_FLX GrdVRS 
               Height          =   2850
               Left            =   5970
               TabIndex        =   108
               Top             =   780
               Width           =   7275
               _ExtentX        =   12832
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
               heading3        =   "Rm Code                                 "
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
               TabIndex        =   109
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
               Index           =   24
               Left            =   10275
               TabIndex        =   110
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
               Index           =   41
               Left            =   6960
               TabIndex        =   111
               ToolTipText     =   "Enter Group By Option"
               Top             =   3720
               Width           =   1365
               _ExtentX        =   2408
               _ExtentY        =   503
               MaxLength       =   8
               DataType        =   4
               IdName          =   "WVRSGRP"
            End
            Begin MwfCtl.MWCTL_MED ATXT 
               Height          =   285
               Index           =   25
               Left            =   -68655
               TabIndex        =   120
               ToolTipText     =   "Balance Bag Quantity"
               Top             =   4005
               Width           =   975
               _ExtentX        =   1720
               _ExtentY        =   503
               Alignment       =   2
               Enabled         =   0   'False
               Mask            =   "####0.0"
               MaxLength       =   7
               DataType        =   2
               IdName          =   "WSUMMTOTBALQTY"
            End
            Begin MwfCtl.MWCTL_MED ATXT 
               Height          =   285
               Index           =   42
               Left            =   -66300
               TabIndex        =   122
               ToolTipText     =   "Balance Bag Weight"
               Top             =   4005
               Width           =   1155
               _ExtentX        =   2037
               _ExtentY        =   503
               Alignment       =   2
               Enabled         =   0   'False
               Mask            =   "######0.000"
               MaxLength       =   11
               DataType        =   2
               IdName          =   "WSUMMTOTBALWT"
            End
            Begin VB.Label ALBL 
               BackStyle       =   0  'Transparent
               Caption         =   "Bal Wt"
               BeginProperty Font 
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
               Left            =   -67440
               TabIndex        =   123
               Top             =   4005
               Width           =   1215
            End
            Begin VB.Label ALBL 
               BackStyle       =   0  'Transparent
               Caption         =   "Bal Bag Qty"
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
               Index           =   6
               Left            =   -69870
               TabIndex        =   121
               Top             =   4005
               Width           =   1335
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
               Index           =   5
               Left            =   5970
               TabIndex        =   119
               Top             =   510
               Width           =   7275
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
               Index           =   1
               Left            =   150
               TabIndex        =   118
               Top             =   510
               Width           =   5745
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
               TabIndex        =   117
               Top             =   3690
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
               Index           =   31
               Left            =   -74850
               TabIndex        =   116
               Top             =   3720
               Width           =   1005
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
               TabIndex        =   115
               Top             =   3720
               Width           =   1215
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
               TabIndex        =   114
               Top             =   3720
               Width           =   1245
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
               TabIndex        =   113
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
               Index           =   63
               Left            =   6000
               TabIndex        =   112
               Top             =   3690
               Width           =   1005
            End
         End
      End
      Begin VB.Frame FraCpyRfId 
         Height          =   3435
         Left            =   960
         TabIndex        =   91
         Top             =   5880
         Visible         =   0   'False
         Width           =   13395
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
            Left            =   7380
            TabIndex        =   186
            ToolTipText     =   "Specify Copy From Bag Nos / Bag Alias No."
            Top             =   200
            Width           =   1575
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
            Left            =   5970
            TabIndex        =   185
            ToolTipText     =   "Specify Copy From Bag Nos / Bag Alias No."
            Top             =   200
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
            TabIndex        =   183
            Top             =   480
            Width           =   12615
            Begin MwfCtl.MWCTL_MED ATXT 
               Height          =   2565
               Index           =   40
               Left            =   60
               TabIndex        =   184
               ToolTipText     =   "Copy Bag Alias No. / Bag Nos  (Yy/Chr/No,YY....,Yy/Chr/No)"
               Top             =   240
               Width           =   12495
               _ExtentX        =   22040
               _ExtentY        =   4524
               Alignment       =   3
               MaxLength       =   0
               IdName          =   "WRFIDSEL"
            End
         End
         Begin VB.Frame FraUpdPlNo 
            Height          =   495
            Left            =   9240
            TabIndex        =   175
            ToolTipText     =   "Select 'CopyBag' to Copy Bags and 'Update PLNo.' to Update PLNo to given Bags"
            Top             =   0
            Width           =   3375
            Begin VB.OptionButton OptCpyBag 
               Caption         =   "Copy Bag"
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   9
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   345
               Left            =   120
               TabIndex        =   93
               ToolTipText     =   "Select 'CopyBag' to Copy Bags and 'Update PLNo.' to Update PLNo to given Bags"
               Top             =   120
               Value           =   -1  'True
               Width           =   1215
            End
            Begin VB.OptionButton OptUpdPlNo 
               Caption         =   "Update PL No."
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   9
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   345
               Left            =   1680
               TabIndex        =   94
               ToolTipText     =   "Select 'CopyBag' to Copy Bags and 'Update PLNo.' to Update PLNo to given Bags"
               Top             =   120
               Width           =   1575
            End
         End
         Begin MwfCtl.MWCTL_BTN1 CmdCpyRfIdGo 
            Height          =   525
            Left            =   12720
            TabIndex        =   95
            Top             =   2760
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
            Left            =   8880
            TabIndex        =   92
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
            Left            =   4680
            TabIndex        =   187
            Top             =   200
            Width           =   1005
         End
         Begin VB.Label lblCbRfIdCpy 
            BackStyle       =   0  'Transparent
            Caption         =   "Copy / Update"
            BeginProperty Font 
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
            TabIndex        =   182
            Top             =   540
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
            Left            =   6600
            TabIndex        =   124
            Top             =   600
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
            TabIndex        =   96
            Top             =   120
            Width           =   4410
         End
      End
      Begin VB.Frame FraEX 
         Height          =   1935
         Left            =   3000
         TabIndex        =   17
         Top             =   7305
         Width           =   9165
         Begin MwfCtl.MWCTL_MED ATXT 
            Height          =   285
            Index           =   39
            Left            =   1065
            TabIndex        =   18
            ToolTipText     =   "Customer Code"
            Top             =   375
            Width           =   1350
            _ExtentX        =   2381
            _ExtentY        =   503
            Enabled         =   0   'False
            MaxLength       =   8
            DataType        =   4
            IdName          =   "WEXFDPRDCMCD"
         End
         Begin MwfCtl.MWCTL_MED ATXT 
            Height          =   285
            Index           =   7
            Left            =   1065
            TabIndex        =   24
            ToolTipText     =   "Enter Export Customer Code"
            Top             =   660
            Width           =   1350
            _ExtentX        =   2381
            _ExtentY        =   503
            MaxLength       =   8
            DataType        =   4
            ReCalcParent    =   "WEXCMNAME"
            ReCalcOn        =   "WEXFDEXPODTC,WEXFDEXPODYY,WEXFDEXPODCHR,WEXFDEXPODNO"
            IdName          =   "WEXFDEXPCMCD"
         End
         Begin MwfCtl.MWCTL_MED ATXT 
            Height          =   285
            Index           =   8
            Left            =   7425
            TabIndex        =   22
            ToolTipText     =   "Enter Sales Order Number"
            Top             =   375
            Width           =   930
            _ExtentX        =   1640
            _ExtentY        =   503
            Alignment       =   2
            Mask            =   "#####0"
            MaxLength       =   6
            DataType        =   1
            ReCalcParent    =   "WEXFDEXPCMCD"
            IdName          =   "WEXFDEXPODNO"
         End
         Begin MwfCtl.MWCTL_MED ATXT 
            Height          =   285
            Index           =   9
            Left            =   6210
            TabIndex        =   20
            ToolTipText     =   "Enter Sales Order Year"
            Top             =   375
            Width           =   435
            _ExtentX        =   767
            _ExtentY        =   503
            MaxLength       =   2
            DataType        =   4
            ReCalcParent    =   "WEXFDEXPCMCD"
            IdName          =   "WEXFDEXPODYY"
         End
         Begin MwfCtl.MWCTL_MED ATXT 
            Height          =   285
            Index           =   10
            Left            =   5535
            TabIndex        =   19
            ToolTipText     =   "Enter Sales Order Tc"
            Top             =   375
            Width           =   570
            _ExtentX        =   1005
            _ExtentY        =   503
            MaxLength       =   3
            DataType        =   4
            ReCalcParent    =   "WEXFDEXPCMCD"
            IdName          =   "WEXFDEXPODTC"
         End
         Begin MwfCtl.MWCTL_MED ATXT 
            Height          =   285
            Index           =   11
            Left            =   6750
            TabIndex        =   21
            ToolTipText     =   "Enter Sales Order Character"
            Top             =   375
            Width           =   570
            _ExtentX        =   1005
            _ExtentY        =   503
            MaxLength       =   3
            DataType        =   4
            ReCalcParent    =   "WEXFDEXPCMCD"
            IdName          =   "WEXFDEXPODCHR"
         End
         Begin MwfCtl.MWCTL_MED ATXT 
            Height          =   285
            Index           =   12
            Left            =   6750
            TabIndex        =   34
            ToolTipText     =   "Enter Packing List Number"
            Top             =   1230
            Width           =   930
            _ExtentX        =   1640
            _ExtentY        =   503
            Alignment       =   2
            Mask            =   "######0"
            MaxLength       =   7
            DataType        =   1
            IdName          =   "WEXFDPLNO"
         End
         Begin MwfCtl.MWCTL_MED ATXT 
            Height          =   285
            Index           =   13
            Left            =   6210
            TabIndex        =   33
            ToolTipText     =   "Enter Packing List Year"
            Top             =   1230
            Width           =   435
            _ExtentX        =   767
            _ExtentY        =   503
            MaxLength       =   2
            DataType        =   4
            IdName          =   "WEXFDPLYY"
         End
         Begin MwfCtl.MWCTL_MED ATXT 
            Height          =   285
            Index           =   14
            Left            =   5535
            TabIndex        =   32
            ToolTipText     =   "Enter Packing List Character"
            Top             =   1230
            Width           =   570
            _ExtentX        =   1005
            _ExtentY        =   503
            MaxLength       =   3
            DataType        =   4
            IdName          =   "WEXFDPLCHR"
         End
         Begin MwfCtl.MWCTL_MED ATXT 
            Height          =   285
            Index           =   15
            Left            =   7425
            TabIndex        =   30
            ToolTipText     =   "Enter Invoice Voucher Number"
            Top             =   945
            Width           =   930
            _ExtentX        =   1640
            _ExtentY        =   503
            Alignment       =   2
            Mask            =   "######0"
            MaxLength       =   7
            DataType        =   1
            IdName          =   "WEXFDINNO"
         End
         Begin MwfCtl.MWCTL_MED ATXT 
            Height          =   285
            Index           =   16
            Left            =   6210
            TabIndex        =   28
            ToolTipText     =   "Enter Invoice Voucher Year"
            Top             =   945
            Width           =   435
            _ExtentX        =   767
            _ExtentY        =   503
            MaxLength       =   2
            DataType        =   4
            IdName          =   "WEXFDINYY"
         End
         Begin MwfCtl.MWCTL_MED ATXT 
            Height          =   285
            Index           =   17
            Left            =   5535
            TabIndex        =   27
            ToolTipText     =   "Enter Invoice Voucher Tc"
            Top             =   945
            Width           =   570
            _ExtentX        =   1005
            _ExtentY        =   503
            MaxLength       =   3
            DataType        =   4
            IdName          =   "WEXFDINTC"
         End
         Begin MwfCtl.MWCTL_MED ATXT 
            Height          =   285
            Index           =   18
            Left            =   6750
            TabIndex        =   29
            ToolTipText     =   "Enter Invoice Voucher Character"
            Top             =   945
            Width           =   570
            _ExtentX        =   1005
            _ExtentY        =   503
            MaxLength       =   3
            DataType        =   4
            IdName          =   "WEXFDINCHR"
         End
         Begin MwfCtl.MWCTL_MED ATXT 
            Height          =   285
            Index           =   46
            Left            =   1065
            TabIndex        =   26
            ToolTipText     =   "Enter Export Number"
            Top             =   945
            Width           =   2940
            _ExtentX        =   5186
            _ExtentY        =   503
            MaxLength       =   20
            DataType        =   4
            IdName          =   "WEXFDINEXPNO"
         End
         Begin MwfCtl.MWCTL_MED ATXT 
            Height          =   285
            Index           =   34
            Left            =   1065
            TabIndex        =   31
            ToolTipText     =   "Despatch Quantity"
            Top             =   1230
            Width           =   600
            _ExtentX        =   1058
            _ExtentY        =   503
            Alignment       =   2
            Enabled         =   0   'False
            Mask            =   "##0.0"
            MaxLength       =   5
            DataType        =   2
            IdName          =   "WEXFDDESPQTY"
         End
         Begin MwfCtl.MWCTL_MED ATXT 
            Height          =   285
            Index           =   20
            Left            =   8460
            TabIndex        =   23
            ToolTipText     =   "Enter Sales Order Sr Number"
            Top             =   375
            Width           =   570
            _ExtentX        =   1005
            _ExtentY        =   503
            Alignment       =   2
            Mask            =   "###0"
            MaxLength       =   4
            DataType        =   1
            IdName          =   "WEXFDEXPODSR"
         End
         Begin MwfCtl.MWCTL_MED ATXT 
            Height          =   285
            Index           =   21
            Left            =   1065
            TabIndex        =   35
            ToolTipText     =   "Base Voucher Number"
            Top             =   1515
            Width           =   2940
            _ExtentX        =   5186
            _ExtentY        =   503
            Enabled         =   0   'False
            MaxLength       =   22
            DataType        =   4
            IdName          =   "WEXFDBASEVCHNO"
         End
         Begin MwfCtl.MWCTL_MED ATXT 
            Height          =   285
            Index           =   23
            Left            =   5535
            TabIndex        =   25
            ToolTipText     =   "Enter Export Customer Name"
            Top             =   660
            Width           =   2940
            _ExtentX        =   5186
            _ExtentY        =   503
            MaxLength       =   45
            ReCalcOn        =   "WEXFDEXPCMCD"
            IdName          =   "WEXCMNAME"
         End
         Begin MwfCtl.MWCTL_MED ATXT 
            Height          =   285
            Index           =   35
            Left            =   3405
            TabIndex        =   82
            ToolTipText     =   "CTB Despatch Quantity"
            Top             =   1230
            Width           =   600
            _ExtentX        =   1058
            _ExtentY        =   503
            Alignment       =   2
            Enabled         =   0   'False
            Mask            =   "##0.0"
            MaxLength       =   5
            DataType        =   2
            IdName          =   "WEXFDCTBDESPQTY"
         End
         Begin VB.Label LblMain 
            BackStyle       =   0  'Transparent
            BorderStyle     =   1  'Fixed Single
            Caption         =   "  Export Details"
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
            TabIndex        =   87
            Top             =   90
            Width           =   9195
         End
         Begin VB.Label LblExFdCTBDespQty 
            BackStyle       =   0  'Transparent
            Caption         =   "CTB Desp Qty"
            BeginProperty Font 
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
            TabIndex        =   83
            Top             =   1230
            Width           =   1515
         End
         Begin VB.Label ALBL 
            BackStyle       =   0  'Transparent
            Caption         =   "Exp Cust Nm"
            BeginProperty Font 
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
            Left            =   4200
            TabIndex        =   77
            Top             =   660
            Width           =   1455
         End
         Begin VB.Label LblExFdBaseVchNo 
            BackStyle       =   0  'Transparent
            Caption         =   "Base Vch"
            BeginProperty Font 
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
            TabIndex        =   76
            Top             =   1515
            Width           =   1095
         End
         Begin VB.Label Label7 
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
            Left            =   8370
            TabIndex        =   69
            Top             =   375
            Width           =   135
         End
         Begin VB.Label LblExFdDespQty 
            BackStyle       =   0  'Transparent
            Caption         =   "Desp Qty"
            BeginProperty Font 
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
            TabIndex        =   68
            Top             =   1230
            Width           =   1095
         End
         Begin VB.Label ALBL 
            BackStyle       =   0  'Transparent
            Caption         =   "Exp No."
            BeginProperty Font 
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
            TabIndex        =   67
            Top             =   945
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
            Index           =   21
            Left            =   6660
            TabIndex        =   66
            Top             =   945
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
            Index           =   20
            Left            =   6120
            TabIndex        =   65
            Top             =   945
            Width           =   135
         End
         Begin VB.Label Label5 
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
            Left            =   7335
            TabIndex        =   64
            Top             =   945
            Width           =   135
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
            Index           =   17
            Left            =   4200
            TabIndex        =   63
            Top             =   945
            Width           =   945
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
            Left            =   6120
            TabIndex        =   62
            Top             =   1230
            Width           =   135
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
            Left            =   6660
            TabIndex        =   61
            Top             =   1230
            Width           =   135
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
            Left            =   4200
            TabIndex        =   60
            Top             =   1230
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
            Index           =   12
            Left            =   6660
            TabIndex        =   59
            Top             =   375
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
            Index           =   8
            Left            =   6120
            TabIndex        =   58
            Top             =   375
            Width           =   135
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
            Left            =   7335
            TabIndex        =   57
            Top             =   375
            Width           =   135
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
            Index           =   4
            Left            =   60
            TabIndex        =   56
            Top             =   660
            Width           =   1095
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
            Index           =   13
            Left            =   4200
            TabIndex        =   54
            Top             =   375
            Width           =   945
         End
         Begin VB.Label LblExFdPrdCmCd 
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
            Left            =   60
            TabIndex        =   53
            Top             =   375
            Width           =   1095
         End
      End
   End
   Begin VB.Frame FraCB 
      Height          =   3555
      Left            =   5160
      TabIndex        =   84
      Top             =   5685
      Width           =   5085
      Begin MwfCtl.MWCTL_BTN1 CmdCBGo 
         Height          =   285
         Left            =   4320
         TabIndex        =   177
         Top             =   3120
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
         Index           =   50
         Left            =   2745
         TabIndex        =   151
         ToolTipText     =   "Enter Packing List Number"
         Top             =   3075
         Width           =   1020
         _ExtentX        =   1799
         _ExtentY        =   503
         Alignment       =   2
         Mask            =   "######0"
         MaxLength       =   7
         DataType        =   1
         IdName          =   "WCBFGPLNO"
      End
      Begin MwfCtl.MWCTL_MED ATXT 
         Height          =   285
         Index           =   51
         Left            =   2040
         TabIndex        =   150
         ToolTipText     =   "Enter Packing List Character"
         Top             =   3075
         Width           =   600
         _ExtentX        =   1058
         _ExtentY        =   503
         MaxLength       =   3
         DataType        =   4
         IdName          =   "WCBFGPLCHR"
      End
      Begin MwfCtl.MWCTL_MED ATXT 
         Height          =   285
         Index           =   49
         Left            =   1575
         TabIndex        =   149
         ToolTipText     =   "Enter Packing List Year"
         Top             =   3075
         Width           =   360
         _ExtentX        =   635
         _ExtentY        =   503
         MaxLength       =   2
         DataType        =   4
         IdName          =   "WCBFGPLYY"
      End
      Begin VB.Frame FraFilt 
         Height          =   2985
         Left            =   0
         TabIndex        =   127
         Top             =   0
         Width           =   5085
         Begin MwfCtl.MWCTL_MED ATXT 
            Height          =   285
            Index           =   61
            Left            =   1575
            TabIndex        =   129
            ToolTipText     =   "Enter From Bag Year"
            Top             =   765
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
            Left            =   2730
            TabIndex        =   131
            ToolTipText     =   "Enter From Bag Number"
            Top             =   765
            Width           =   1020
            _ExtentX        =   1799
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
            Left            =   2035
            TabIndex        =   130
            ToolTipText     =   "Enter From Bag Character"
            Top             =   765
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
            Left            =   1575
            TabIndex        =   132
            ToolTipText     =   "Enter To Bag Year"
            Top             =   1050
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
            Left            =   2730
            TabIndex        =   134
            ToolTipText     =   "Enter To Bag Number"
            Top             =   1050
            Width           =   1020
            _ExtentX        =   1799
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
            Left            =   2035
            TabIndex        =   133
            ToolTipText     =   "Enter To Bag Character"
            Top             =   1050
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
            Left            =   2280
            TabIndex        =   136
            ToolTipText     =   "Enter From Bag Order Year "
            Top             =   1335
            Width           =   360
            _ExtentX        =   635
            _ExtentY        =   503
            MaxLength       =   2
            DataType        =   4
            ReCalcParent    =   "WCBODYYTO"
            IdName          =   "WCBODYYFR"
         End
         Begin MwfCtl.MWCTL_MED ATXT 
            Height          =   285
            Index           =   79
            Left            =   3420
            TabIndex        =   138
            ToolTipText     =   "Enter From Bag Order Number"
            Top             =   1335
            Width           =   780
            _ExtentX        =   1376
            _ExtentY        =   503
            Alignment       =   2
            Mask            =   "#####0"
            MaxLength       =   6
            DataType        =   1
            ReCalcParent    =   "WCBODNOTO"
            IdName          =   "WCBODNOFR"
         End
         Begin MwfCtl.MWCTL_MED ATXT 
            Height          =   285
            Index           =   83
            Left            =   1575
            TabIndex        =   135
            ToolTipText     =   "Enter From Bag Order Tc "
            Top             =   1335
            Width           =   600
            _ExtentX        =   1058
            _ExtentY        =   503
            MaxLength       =   3
            DataType        =   4
            ReCalcParent    =   "WCBODTCTO"
            IdName          =   "WCBODTCFR"
         End
         Begin MwfCtl.MWCTL_MED ATXT 
            Height          =   285
            Index           =   84
            Left            =   2730
            TabIndex        =   137
            ToolTipText     =   "Enter From Bag Order Character"
            Top             =   1335
            Width           =   600
            _ExtentX        =   1058
            _ExtentY        =   503
            MaxLength       =   3
            DataType        =   4
            ReCalcParent    =   "WCBODCHRTO"
            IdName          =   "WCBODCHRFR"
         End
         Begin MwfCtl.MWCTL_MED ATXT 
            Height          =   285
            Index           =   85
            Left            =   4275
            TabIndex        =   139
            ToolTipText     =   "Enter From Bag Order Serial Number"
            Top             =   1335
            Width           =   570
            _ExtentX        =   1005
            _ExtentY        =   503
            Alignment       =   2
            Mask            =   "###0"
            MaxLength       =   4
            DataType        =   1
            ReCalcParent    =   "WCBODSRTO"
            IdName          =   "WCBODSRFR"
         End
         Begin MwfCtl.MWCTL_MED ATXT 
            Height          =   285
            Index           =   86
            Left            =   2280
            TabIndex        =   141
            ToolTipText     =   "Enter To Bag Order Year"
            Top             =   1620
            Width           =   360
            _ExtentX        =   635
            _ExtentY        =   503
            MaxLength       =   2
            DataType        =   4
            ReCalcOn        =   "WCBODYYFR"
            IdName          =   "WCBODYYTO"
         End
         Begin MwfCtl.MWCTL_MED ATXT 
            Height          =   285
            Index           =   87
            Left            =   3420
            TabIndex        =   143
            ToolTipText     =   "Enter To Bag Order Number"
            Top             =   1620
            Width           =   780
            _ExtentX        =   1376
            _ExtentY        =   503
            Alignment       =   2
            Mask            =   "#####0"
            MaxLength       =   6
            DataType        =   1
            ReCalcOn        =   "WCBODNOFR"
            IdName          =   "WCBODNOTO"
         End
         Begin MwfCtl.MWCTL_MED ATXT 
            Height          =   285
            Index           =   88
            Left            =   1575
            TabIndex        =   140
            ToolTipText     =   "Enter To Bag Order Tc"
            Top             =   1620
            Width           =   600
            _ExtentX        =   1058
            _ExtentY        =   503
            MaxLength       =   3
            DataType        =   4
            ReCalcOn        =   "WCBODTCFR"
            IdName          =   "WCBODTCTO"
         End
         Begin MwfCtl.MWCTL_MED ATXT 
            Height          =   285
            Index           =   89
            Left            =   2730
            TabIndex        =   142
            ToolTipText     =   "Enter To Bag Order Character"
            Top             =   1620
            Width           =   600
            _ExtentX        =   1058
            _ExtentY        =   503
            MaxLength       =   3
            DataType        =   4
            ReCalcOn        =   "WCBODCHRFR"
            IdName          =   "WCBODCHRTO"
         End
         Begin MwfCtl.MWCTL_MED ATXT 
            Height          =   285
            Index           =   90
            Left            =   4275
            TabIndex        =   144
            ToolTipText     =   "Enter To Bag Order Serial Number"
            Top             =   1620
            Width           =   570
            _ExtentX        =   1005
            _ExtentY        =   503
            Alignment       =   2
            Mask            =   "###0"
            MaxLength       =   4
            DataType        =   1
            ReCalcOn        =   "WCBODSRFR"
            IdName          =   "WCBODSRTO"
         End
         Begin MwfCtl.MWCTL_MED ATXT 
            Height          =   285
            Index           =   91
            Left            =   1575
            TabIndex        =   145
            ToolTipText     =   "Enter From Design Code"
            Top             =   1905
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
            Left            =   1575
            TabIndex        =   146
            ToolTipText     =   "Enter To Design Code"
            Top             =   2190
            Width           =   2415
            _ExtentX        =   4260
            _ExtentY        =   503
            MaxLength       =   15
            DataType        =   4
            IdName          =   "WCBDMCDTO"
         End
         Begin MwfCtl.MWCTL_MED ATXT 
            Height          =   285
            Index           =   38
            Left            =   1575
            TabIndex        =   128
            ToolTipText     =   "Enter From Bag Character"
            Top             =   480
            Width           =   1320
            _ExtentX        =   2328
            _ExtentY        =   503
            MaxLength       =   8
            DataType        =   4
            IdName          =   "WCBBFGSUBLOC"
         End
         Begin MwfCtl.MWCTL_MED ATXT 
            Height          =   285
            Index           =   43
            Left            =   1575
            TabIndex        =   147
            ToolTipText     =   "Enter From Customer Code"
            Top             =   2475
            Width           =   1455
            _ExtentX        =   2566
            _ExtentY        =   503
            MaxLength       =   8
            DataType        =   4
            IdName          =   "WCBCUSTCDFR"
         End
         Begin MwfCtl.MWCTL_MED ATXT 
            Height          =   285
            Index           =   44
            Left            =   3390
            TabIndex        =   148
            ToolTipText     =   "Enter To Customer Code"
            Top             =   2475
            Width           =   1455
            _ExtentX        =   2566
            _ExtentY        =   503
            MaxLength       =   8
            DataType        =   4
            IdName          =   "WCBCUSTCDTO"
         End
         Begin VB.Label LblMain 
            BackStyle       =   0  'Transparent
            BorderStyle     =   1  'Fixed Single
            Caption         =   "  Filters"
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
            Left            =   0
            TabIndex        =   176
            Top             =   120
            Width           =   5085
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
            Left            =   120
            TabIndex        =   172
            Top             =   2190
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
            Left            =   120
            TabIndex        =   171
            Top             =   1905
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
            Left            =   3345
            TabIndex        =   170
            Top             =   1620
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
            Left            =   2655
            TabIndex        =   169
            Top             =   1620
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
            Left            =   2190
            TabIndex        =   168
            Top             =   1620
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
            Left            =   120
            TabIndex        =   167
            Top             =   1620
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
            Left            =   4200
            TabIndex        =   166
            Top             =   1620
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
            Left            =   3345
            TabIndex        =   165
            Top             =   1335
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
            Left            =   2655
            TabIndex        =   164
            Top             =   1320
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
            Left            =   2190
            TabIndex        =   163
            Top             =   1335
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
            Left            =   120
            TabIndex        =   162
            Top             =   1335
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
            Left            =   4200
            TabIndex        =   161
            Top             =   1335
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
            Left            =   2655
            TabIndex        =   160
            Top             =   1050
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
            Left            =   1950
            TabIndex        =   159
            Top             =   1050
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
            Left            =   120
            TabIndex        =   158
            Top             =   1050
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
            Left            =   2655
            TabIndex        =   157
            Top             =   765
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
            Left            =   1950
            TabIndex        =   156
            Top             =   765
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
            Left            =   120
            TabIndex        =   155
            Top             =   765
            Width           =   855
         End
         Begin VB.Label LblCbSubLoc 
            BackStyle       =   0  'Transparent
            Caption         =   "Sub Loc"
            BeginProperty Font 
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
            TabIndex        =   154
            Top             =   480
            Width           =   855
         End
         Begin VB.Label ALBL 
            BackStyle       =   0  'Transparent
            Caption         =   "Cust Code"
            BeginProperty Font 
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
            Left            =   120
            TabIndex        =   153
            Top             =   2475
            Width           =   1095
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
            Index           =   18
            Left            =   3060
            TabIndex        =   152
            Top             =   2475
            Width           =   375
         End
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
         Left            =   2655
         TabIndex        =   174
         Top             =   3120
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
         Left            =   1950
         TabIndex        =   173
         Top             =   3120
         Width           =   105
      End
      Begin VB.Label LblDefPckLstNo 
         BackStyle       =   0  'Transparent
         Caption         =   "Def Pck List No"
         BeginProperty Font 
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
         TabIndex        =   126
         Top             =   3120
         Width           =   1455
      End
   End
End
Attribute VB_Name = "EmrFrmFg"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim moCn As MwfLib.MDOConnection   '*** (Bef speed) bef 24/11/06

Enum en_FgFra
  ExDet = 0
  Cpy = 1
  Summ = 2
  CpyBag = 3
  CpyRfId = 4     ' ***** Manali 3.10.0 - 19/04/12 - Cpy Bags from RfId
  InvHd = 5     'Sv.22
End Enum
Dim mf_CpyStartedByParent As Boolean
'*** Declare the variables ms_LockDt, mf_DtLocked for the Lock Date Check
Dim ms_LockDt As String
Dim ms_FutureDt As String
Dim ms_PurgeDt As String    '****** Sachin 3.02 - 24-12-07
Dim mf_DtLocked As Boolean
Dim mf_ValidLocSeq As Boolean

Dim ms_TcTyp As String, ms_EntTyp As String, ms_FgCoCd As String
Dim ms_FgTc As String, ms_FgYy As String, ms_FgChr As String
Dim wgrdfgdrow As Integer
Dim mb_ARowExists As Boolean
Dim mb_EnfFrBLocInFg As Boolean

'*** Arr_Cv stores the causes validation state, Arr_Ed the Ena/Disa state of Cmd Buttons
Dim mArr_Cv As Variant, mArr_Ed As Variant

Dim mb_AddRow As Boolean      ' ***** Manali 3.10.0 - 09/05/12 - Bar Code (Ctrl+N)
Dim wStkNoDisp As String, ms_ChkFgWghtYN As String

Private Sub ADC_Load()
  '*** Redim the arrays mArr_Cv and mArr_Ed to the number of command buttons
  '*** Set the First Key Control property to the field FgChr
  '*** Set the First Non Key Control property to FgDt
  '*** Set the Child Property Of ADC to GrdFgd
  '*** Set the Child Property Of GrdFgd to GrdFgRm
  '*** Set the Previous and Next Control Properties for the Grids GrdFgd and GrdFgRm
  '*** Set the Hot Keys for the Grids GrdFgd and GrdFgRm as 'A' and 'B'
  '*** Bring the FraNKeyAll frame to the top
  '*** Get the default values for Tc and Yy from param and head files for the first time the
  '    form is loaded
  '*** Call the Private Sub Procedure PropSet to set the enable and visible properties
  '    for all the controls on the form based on the TcStructure file
  
'  ReDim mArr_Cv(3), mArr_Ed(3)
  ReDim mArr_Cv(5), mArr_Ed(5)      ' ***** Manali 3.10.0 - 19/04/12 - Cpy Bags from RfId
  
  Set moCn = ADC.Connection    '*** (Bef speed) bef 24/11/06
  Set ADC.FirKeyCtl = ADC("FgChr")
  Set ADC.FirNKeyCtl = ADC("FgDt")
  ADC.Child = "GrdFgd"
  GrdFgd.Child = "GrdFgRm"
  Set GrdFgd.PrevCtl = ADC("FgFrBLoc")
  Set GrdFgd.NextCtl = ADC
  Set GrdFgRm.PrevCtl = GrdFgd
  Set GrdFgRm.NextCtl = GrdFgd
  Set GrdSumm.NextCtl = ADC("wSummBagGrp")
  GrdFgd.HotKey = "A"
  GrdFgRm.HotKey = "B"
  FraNKeyAll.ZOrder (0)
  mf_CpyStartedByParent = False

  ms_FgCoCd = gs_CoCd
  ms_TcTyp = ADC.MenuCd
  ms_EntTyp = moCn.GetFldVal("Select PValue1 from Param where PTyp= 'TCTYP' and " + _
            "PMCd= '" + ms_TcTyp + "'")
  ms_FgTc = moCn.GetFldVal("Select PValue from Param where PTyp= 'TCTYP' and " + _
               "PMCd='" + ms_TcTyp + "' ")
  ms_FgYy = moCn.GetFldVal("Select hYy from Head where HCoCd= '" + ms_FgCoCd + "' and HCd='" + ctSelfCmCd + "'")
  Call PropSet
  ' Zubin 212
  If moCn.RecSeek("select 'x' from Head where HCd='" + ctSelfCmCd + "' and HCoCd='" + gs_CoCd + "' and HLotNoYN='N'") Then _
    GrdFgRm.ColProp("FrLotNo").Style = fgcinvisible
  ' Zubin 212
  
  '****** Sachin 3.02 24-12-07 - Setting the Purge Date
  'sv.9 getting purge date from vparam
  ms_PurgeDt = GetPurgeDt(gs_CoCd, "")
  
   '****** Sachin 3.02.0 - Hiding the Partition Key Column based on Partition Selected
  ADC("FgPrtKey").Visible = IIF(gs_Partition = ctCurrPrtn, False, True)
  GrdFgd.ColProp("FdPrtKey").Style = IIF(gs_Partition = ctCurrPrtn, fgcinvisible, fgcReadOnly)
  GrdFgRm.ColProp("FrPrtKey").Style = IIF(gs_Partition = ctCurrPrtn, fgcinvisible, fgcReadOnly)
  
  '  CmdCpyRfId.Enabled = True: CmdCpyRfId.Visible = True
  ' ***** Manali 3.10.0 - 19/04/12 - Cpy Bags from RfId
  '6.1b-22 Rfid module check is removed
'  If Not moCn.RecSeek("Select 'x' from Head Where HCoCd='" + gs_CoCd + "' and HCd='" + ctSelfCmCd + "' and HRfIdMod='Y' ") Then
'    CmdCpyRfId.Enabled = False: CmdCpyRfId.Visible = False
'  End If
  '
  ' ***** Manali 3.10.0 - 19/04/12 - Cpy Bags from RfId
  
  mb_AddRow = True      ' ***** Manali 3.10.0 - 09/05/12 - Bar Code (Ctrl+N)
  
  mb_EnfFrBLocInFg = moCn.RecSeek("Select HEnfFrBLocInFgYN From Head where HCd='" + ctSelfCmCd + "' and HCoCd='" + gs_CoCd + "' and HEnfFrBLocInFgYN='Y'")
  
  Dim mb_SeoUsr As Boolean
  If moCn.GetFldVal("Select hAvgSRtYn from Head where HCoCd= '" + gs_CoCd + "'") = "Y" Then
  'StkRt-Getting stock rate display Y\N from head table
  wStkNoDisp = moCn.GetFldVal("Select hStkRtNoDsp from Head where hcocd='" + gs_CoCd + "'")

  'StkRt-when NoDisplay =NO then stock rate is visible,if NoDisplay =yes & user donot have
  'SEO rights for the form then stock rate column is Invisible else visible for Seo users.
  If wStkNoDisp = "N" Then
    GrdFgRm.ColProp("FRRMSTKRT").Style = fgcReadOnly
  ElseIf wStkNoDisp = "Y" Then
    mb_SeoUsr = moCn.RecSeek("Select 'x' From UsrAccess where UaUsrCd = '" + gs_UsrCd + "' " + _
                        " and UaMnuCd = '" + ADC.MenuCd + "' and UaSeoYn= 'Y' ")
    If mb_SeoUsr = True Then GrdFgRm.ColProp("FRRMSTKRT").Style = fgcReadOnly
    If mb_SeoUsr = False Then GrdFgRm.ColProp("FRRMSTKRT").Style = fgcinvisible
  End If
  End If
  'MW.12- PL No. will be invisible only for FWB entry
  If ADC.MenuCd = "FWB" Then
    LblDefPckLstNo.Visible = False:     lblSlash1.Visible = False:     lblSlash2.Visible = False
    ADC("WCBFGPLYY").Visible = False: ADC("WCBFGPLCHR").Visible = False: ADC("WCBFGPLNO").Visible = False:
    '578-24 copybag/ update plno options will be invisible for FWB
    FraUpdPlNo.Visible = False: OptCpyBag.Visible = False: OptUpdPlNo.Visible = False
    LblPlNo.Caption = "PlNo Filter"
    '5713.59 scale wt not needed for fg return, scale wt frame made invisible
    FraBagScaleWt.Visible = False:   LblBagScaleWt.Visible = False
  End If
    
End Sub
Private Sub ADC_KeyWhen()
  '*** Set the dafault value for the form level variable mf_DtLocked as false.
  '*** Store the sequence of valid Process Locations & Other Locations in the module level
  '    variables gs_ValidPrdLocSeq and gs_ValidOthLocSeq
  '*** Set the default value of form level variable mf_ValidLocSeq as 'True'
  '*** Disable All command buttons
  '*** Hide all frames except Frame FraNKeyAll
  '*** setting the default values for FgTc,FgYy and FgChr from the previous entry
  
  mf_DtLocked = False
  mf_ValidLocSeq = True
  
  CmdCP.Enabled = False: CmdEX.Enabled = False: CmdSumm.Enabled = False: CmdCB.Enabled = False: CmdCpyRfId.Enabled = False
  CmdInvHd.Enabled = False    'Sv.22
  Call HideAllFras
  mb_ARowExists = False
  ADC("wTotQty") = 0: ADC("wTotGrWt") = 0
  If ms_FgCoCd <> "" Then ADC("FgCoCd") = ms_FgCoCd
  If ms_FgTc <> "" Then ADC("FgTc") = ms_FgTc
  If ms_FgYy <> "" Then ADC("FgYy") = ms_FgYy
  If ms_FgChr <> "" Then ADC("FgChr") = ms_FgChr
  
  '****** Sachin 2.14.0 - [24-11-2006] - [Multi-Company] ******
  Me.Caption = GetFrmCaption(gs_CoCd, IIF(ms_TcTyp = "FB", "Finished Goods Entry", "Finished Goods (Bag Return)"))
  '****** Sachin 2.14.0 - [24-11-2006] - [Multi-Company] ******
  '*** Jay 2.14 *** [Over Rule Tolerance Check By Providing SEO PWD]
  TxtPwd.text = ""
  ADC("FgSeoPwd") = ""
  '*** Jay 2.14 *** [Over Rule Tolerance Check By Providing SEO PWD]
End Sub
Private Sub ADC_KeyValid(Cancel As Boolean, ErrMsg As String)
      '*** Check whether the FgYy, FgChr are valid
  If ADC("FgYy") = "" Then Cancel = True: ErrMsg = "Enter Voucher Year": Exit Sub
  If ADC("FgChr") = "" Then Cancel = True: ErrMsg = "Enter Voucher Character": Exit Sub
End Sub

Private Sub ADC_NonKeyWhen(Cancel As Boolean, ErrMsg As String)
  '*** Get the Lock Date from Param('LOCKDT', 'LOCKDT') and store in the form level variable ms_LockDt
  '    If ms_LockDt is blank then store '01/01/80' in ms_LockDt
  '*** Store the FgTc, FgYy and FgChr values for the next entry
  '*** Give default values for the fields FgDt,FgFrRmLoc,FgToRmLoc,wFrRmDc,FgFrBLoc,
      'FgToBLoc of Fg
  '*** If the entry Date is < ms_LockDt then mf_DtLocked= true (i.e the entry
  '    If mf_DtLocked = true then the entry is locked. i.e. the user cannot add, edit or
  '    delete records in the grids GrdFgd & GrdFgRm. Similarly if the user does not
  '    have rights to any of FgFrBLoc, FgToBLoc, FgFrRmLoc or FgFrToLoc locations, then the
  '    user cannot add, edit or delete records in the grids GrdFgd & GrdFgRm.

      If ADC.IsNew And Not moCn.RecSeek("select vPSCd from vParam where " + _
             "vPCoCd = '" + ADC("FgCoCd") + "' and vPTyp= 'CHR' and " + _
             "vPMCd = '" + ADC("FgTc") + "' and vPSCd='" + ADC("FgChr") + "' and vPValidYn<>'N'") Then
       Cancel = True: ErrMsg = "Voucher Character No More Valid": ADC.AllowSave = False: Exit Sub
      Else
        ADC.AllowSave = True
      End If

  '****** Sachin 3.02 26-12-07 - Setting Default Current Partition in ADC for new Records
  If ADC.IsNew And Trim(ADC("FgPrtKey")) = "" Then ADC("FgPrtKey") = ctCurrPrtn

  'ms_LockDt = moCn.GetFldVal("Select PDesc225 from Param where PTyp= 'LOCKDT' and PMCd= 'LOCKDT'")
  'ms_FutureDt = moCn.GetFldVal("Select PDesc225 from Param where PTyp= 'LOCKDT' and PMCd= 'FUTUREDT'")
  
  '6.1_51 wVchrNo added to GetLockDt to check that given voucher number is exist in unlock ptyp,
  'if given voucher is in unlock ptyp it will not check lock date
  Dim wVchrNo As String
  wVchrNo = ADC("FgTc") + "/" + ADC("FgYy") + "/" + ADC("FgChr") + "/" + CStr(ADC("FgNo"))
   Call GetLockDt(ms_FgCoCd, "TC", ms_FgTc, ms_LockDt, ms_FutureDt, wVchrNo)
   
 ' If ms_LockDt = "" Then ms_LockDt = "01/01/80"
  'If ms_FutureDt = "" Then ms_FutureDt = "31/12/28"
  
  '*** Parameters for checking the Valid User Locs in the Txn Entries
  gs_ValidPrdLocSeq = Trim(moCn.GetFldVal("Select PDesc225 from param where PTyp= 'USR' and " + _
                                     "PMCd= '" + EmrFrmMDI.TXT_UCD + "'"))
  gs_ValidOthLocSeq = Trim(moCn.GetFldVal("Select PValue1 from param where PTyp= 'USR' and " + _
                                     "PMCd= '" + EmrFrmMDI.TXT_UCD + "'"))
                                     
  '*** Parameters for checking the Valid User Locs in the Txn Entries

  FraNKeyAll.Enabled = True
  ms_FgTc = ADC("FgTc")
  ms_FgYy = ADC("FgYy")
  ms_FgChr = ADC("FgChr")
  
  '*** Jay 2.14 *** [Over Rule Tolerance Check By Providing SEO PWD]
  ADC("FgSeoPwd") = ""
  TxtPwd.text = ""
  '*** Jay 2.14 *** [Over Rule Tolerance Check By Providing SEO PWD]
  
  'Uma KBS
                      
                      
  If ADC.AddFndMode = xaddmode Then
    
    
    Dim ws_Dt As String
    '53 company code parameter added
    ws_Dt = GetNextDate(ADC("FgCoCd"))
    ws_Dt = IIF((ws_Dt < CDate(ms_LockDt)) _
                               Or _
                (ws_Dt > CDate(ms_FutureDt)), _
                "", ws_Dt)
    If ws_Dt <> "" Then ADC("FgDt") = ws_Dt
    
    'ADC("FgDt") = moCn.SrvrDate
    
    ADC("FgFrRmLoc") = moCn.GetFldVal("Select vPValue from vParam where " + _
                                      "vPCoCd= '" + ADC("FgCoCd") + "' and vPTyp= 'CHR' and " + _
                                      "vPMCd= '" + ADC("FgTc") + "' and vPSCd= '" + ADC("FgChr") + "' and vPValidYn<>'N' ")
    ADC("FgToRmLoc") = moCn.GetFldVal("Select vPDesc225 from vParam where " + _
                                      "vPCoCd= '" + ADC("FgCoCd") + "' and vPTyp= 'CHR' and " + _
                                      "vPMCd= '" + ADC("FgTc") + "' and vPSCd= '" + ADC("FgChr") + "' and vPValidYn<>'N' ")
      
    
    '**** urmi FgSubLoc 10/8/03 ***********************************************'
    If ms_TcTyp = "FB" Then
      ADC("FgSubLoc") = moCn.GetFldVal("Select hFgSubLoc from Head where HCoCd= '" + _
                      ADC("FgCoCd") + "' and HCd='" + ctSelfCmCd + "'")
                            
    End If
  End If
  
  If ADC("wFrRmDc") = "" And (ms_TcTyp = "FR" Or ms_TcTyp = "FWR") Then
    Dim wFrRmDc As String
    wFrRmDc = moCn.GetFldVal("Select vPValue2 from vParam where " + _
                             "vPCoCd= '" + ADC("FgCoCd") + "' and vPTyp= 'CHR' and " + _
                             "vPMCd= '" + ADC("FgTc") + "' and vPSCd= '" + ADC("FgChr") + "' and vPValidYn<>'N' ")
    If Trim(wFrRmDc) <> "" Then ADC("wFrRmDc") = wFrRmDc
  End If
  If ms_TcTyp = "FB" Or ms_TcTyp = "FWB" Then ADC("wFrRmDc") = "C"

  If ms_TcTyp = "FB" And ADC.AddFndMode = xaddmode Then
    ADC("FgFrBLoc") = moCn.GetFldVal("Select vPValue1 from vParam where " + _
                             "vPCoCd= '" + ADC("FgCoCd") + "' and vPTyp='CHR' and " + _
                              "vPMCd= '" + ADC("FgTc") + "' and vPSCd='" + ADC("FgChr") + "' and vPValidYn<>'N' ")
    ADC("FgToBLoc") = ctFgLoc
  End If

  If ms_TcTyp = "FWB" And ADC.AddFndMode = xaddmode Then
    ADC("FgToBLoc") = moCn.GetFldVal("Select vPValue1 from vParam where " + _
                             "vPCoCd= '" + ADC("FgCoCd") + "' and vPTyp= 'CHR' and " + _
                             "vPMCd='" + ADC("FgTc") + "' and vPSCd='" + ADC("FgChr") + "' and vPValidYn<>'N' ")
    ADC("FgFrBLoc") = ctFgLoc
  End If
    
  
  mf_DtLocked = IIF((ADC("FgDt") <> "" And ADC("FgDt") < CDate(ms_LockDt)) _
                                Or _
                   (ADC("FgDt") <> "" And ADC("FgDt") > CDate(ms_FutureDt)) _
                                Or _
                  (ADC("FgDt") <= CDate(ms_PurgeDt)), _
                True, False)
    
  If mf_DtLocked = True Then
    GrdFgd.AllowAdd = False: GrdFgd.AllowDelete = False
    GrdFgRm.AllowAdd = False: GrdFgRm.AllowDelete = False
  Else
    If gs_ValidPrdLocSeq <> "" Then
      If ((ADC("FgFrBLoc") <> "" And InStr(1, "," + UCase(gs_ValidPrdLocSeq) + ",", "," + UCase(ADC("FgFrBLoc")) + ",") <= 0) Or _
         (ADC("FgToBLoc") <> "" And InStr(1, "," + UCase(gs_ValidPrdLocSeq) + ",", "," + UCase(ADC("FgToBLoc")) + ",") <= 0)) And ms_TcTyp <> "FB" Then
         mf_ValidLocSeq = False
        GrdFgd.AllowAdd = False: GrdFgd.AllowDelete = False
        GrdFgRm.AllowAdd = False: GrdFgRm.AllowDelete = False
        ErrMsg = "Cannot Edit records As the Bag Loc. is Not Accessible to the User"    '***********Bhavna
      
      ElseIf ms_TcTyp = "FB" Then
          '**********Bhavna 5.1.0  multiselect for 'FB' TC
          Dim ws_LocArr() As String
          Dim i As Integer
          ws_LocArr = Split(ADC("FgFrBLoc"), ",")
          If IsArray(ws_LocArr) Then
            For i = 0 To UBound(ws_LocArr)
              
              If (gs_ValidPrdLocSeq <> "" And InStr(1, "," + UCase(gs_ValidPrdLocSeq) + ",", "," + ws_LocArr(i) + ",") <= 0) Then
                mf_ValidLocSeq = False
                GrdFgd.AllowAdd = False: GrdFgd.AllowDelete = False
                GrdFgRm.AllowAdd = False: GrdFgRm.AllowDelete = False
                ErrMsg = "Cannot Edit records Loc." + ws_LocArr(i) + " is Not Accessible to the User"
              End If
              
            Next i
          End If
          
          '**********Bhavna
      Else
        '***
        mf_ValidLocSeq = True
        GrdFgd.AllowAdd = True: GrdFgd.AllowDelete = True
        If ms_TcTyp = "FR" Then GrdFgRm.AllowAdd = True: GrdFgRm.AllowDelete = True
        '***
      End If
    Else
      mf_ValidLocSeq = True
      GrdFgd.AllowAdd = True: GrdFgd.AllowDelete = True
      If ms_TcTyp = "FR" Then GrdFgRm.AllowAdd = True: GrdFgRm.AllowDelete = True
    End If
  
    If gs_ValidOthLocSeq <> "" Then
      If (ADC("FgFrRmLoc") <> "" And InStr(1, "," + UCase(gs_ValidOthLocSeq) + ",", "," + UCase(ADC("FgFrRmLoc")) + ",") <= 0) Or _
         (ADC("FgToRmLoc") <> "" And InStr(1, "," + UCase(gs_ValidOthLocSeq) + ",", "," + UCase(ADC("FgToRmLoc")) + ",") <= 0) Then
        mf_ValidLocSeq = False
        GrdFgd.AllowAdd = False: GrdFgd.AllowDelete = False
        GrdFgRm.AllowAdd = False: GrdFgRm.AllowDelete = False
      Else
        '***
        If mf_ValidLocSeq = True Then
          GrdFgd.AllowAdd = True: GrdFgd.AllowDelete = True
          If ms_TcTyp = "FR" Then GrdFgRm.AllowAdd = True: GrdFgRm.AllowDelete = True
        End If
        '***
      End If
    Else
      If mf_ValidLocSeq = True Then
        GrdFgd.AllowAdd = True: GrdFgd.AllowDelete = True
        If ms_TcTyp = "FR" Then GrdFgRm.AllowAdd = True: GrdFgRm.AllowDelete = True
      End If
    End If
  
  If ADC("FgSubLoc") <> "" Then
    If Not moCn.RecSeek("select 'x' From vParam where vPTyp='FgSubLoc' and vPCoCd= '" + gs_CoCd + "' and vPMCd='" + ADC("FgSubLoc") + "'" + _
                    " and (vPValue4='' Or ','+vPValue4+',' Like '%," + gs_UsrCd + ",%')") Then
          mf_ValidLocSeq = False
          GrdFgd.AllowAdd = False: GrdFgd.AllowDelete = False
          GrdFgRm.AllowAdd = False: GrdFgRm.AllowDelete = False
     Else
          If mf_ValidLocSeq = True Then
            GrdFgd.AllowAdd = True: GrdFgd.AllowDelete = True
            If ms_TcTyp = "FR" Then GrdFgRm.AllowAdd = True: GrdFgRm.AllowDelete = True
          End If
    End If
  End If
     CmdCB.Enabled = True
     CmdCpyRfId.Enabled = True
     'Rst.125 getting scale wtYN from head
     ms_ChkFgWghtYN = moCn.GetFldVal("Select hChkFgWghtYN from Head where HCoCd= '" + _
                                ADC("FgCoCd") + "' and HCd='" + ctSelfCmCd + "'")
     'Rst.125 if scale wtYN is Yes then copy buttons will be disable
     '5713.59 scale wt not needed for fg return, scale wt in grid and its related calculations are ignored
     If ms_ChkFgWghtYN = "Y" And ms_TcTyp <> "FWB" Then
        CmdCB.Enabled = False
        CmdCpyRfId.Enabled = False
        GrdFgd.ColProp("wScaleWt").Style = fgcNormal
     Else
        GrdFgd.ColProp("wScaleWt").Style = fgcinvisible
     End If
  End If
     CmdSumm.Enabled = True
     CmdInvHd.Enabled = True    'Sv.22
  'vk.4 setting default values
  OptCpyBagFr(0).Value = True
  ADC("wCpyThruBagNo") = "Y"
  
End Sub
Private Sub ADC_NonKeyValid(Cancel As Boolean, ErrMsg As String)
  '*** Before leaving the entry check if Raw Material have been copied for all the bags in
      'the entry. If not the user has to copy the Raw Materials before he can leave the entry.
      'In case of 'FR' entries the user has to manually enter the Raw Material records
  If ADC("FgDt") > CDate(ms_PurgeDt) And (ms_TcTyp = "FB" Or ms_TcTyp = "FWB" Or ms_TcTyp = "FR" Or ms_TcTyp = "FWR") Then
 '     Dim wFdSrChkSql As String
    '****** Sachin 3.02 - Id fields in Joins
  '  wFdSrChkSql = " ( Select FdSr, Convert(Decimal (14,3), Max(FdGrWt)) as qFdGrWt, " + _
'    "  Convert(Decimal (14,3), Isnull(Sum(FrRmWt / (Case when RmCtg in ('D','C') then 5 else 1 end)),0))  as qFrWt " + _
 '   "  From Fgd " + _
  '  "     Left Outer Join FgRm On FrFdIdNo=FdIdNo and FrPrtKey=FdPrtKey and FrCoCd = FdCoCd and FrTc= FdTc and FrYy= FdYy and FrChr= FdChr and FrNo= FdNo and FrSr= FdSr " + _
'    "     Left Outer Join RmMst on RmCd= FrRmCd and RmPrtKey=FrPrtKey and RmZ= 'N' " + _
 '   "  Where " + IIF(gs_Partition = ctCurrPrtn, " FdPrtKey='" + ctCurrPrtn + "' And ", "") + " FdCoCd= '" + adc("FgCoCd") + "' and FdTc= '" + adc("FgTc") + "' and FdYy='" + adc("FgYy") + "' and " + _
  '  "        FdChr='" + adc("FgChr") + "' and FdNo=" + CStr(adc("FgNo")) + " " + _
'    "  Group By FdCoCd, FdTc, FdYy, FdChr, FdNo, FdSr " + _
 '   ") tbl " + _
'    "Where qFrWt<> qFdGrWt"
    'uma next 2 line to comment to ignore error (sri madam)
    'wFdSrChkSql = moCn.GetStrConcatVal("Ltrim(Str(FdSr))", "", wFdSrChkSql, ", ")
    'If wFdSrChkSql <> "" Then Cancel = True: ErrMsg = "Raw Materials Half Copied /Not Copied For Serial Nos: " + wFdSrChkSql: Exit Sub
    '*** Jay 3.01 *** [FgRmWt <> FdGrWt]
    Dim wFdSrChkSql As String
    '****** Sachin 3.02 - Id fields in Joins
    wFdSrChkSql = " ( Select FdSr, Convert(Decimal (14,3), Max(FdGrWt)) as qFdGrWt, " + _
    "  Convert(Decimal (14,3), Isnull(Sum(FrRmWt / (Case when RmCtg in ('D','C') then 5 else 1 end)),0))  as qFrWt " + _
    "  From Fgd " + _
    "     Left Outer Join FgRm On FrFdIdNo=FdIdNo and FrPrtKey=FdPrtKey and FrCoCd = FdCoCd and FrTc= FdTc and FrYy= FdYy and FrChr= FdChr and FrNo= FdNo and FrSr= FdSr " + _
    "     Left Outer Join RmMst on RmCd= FrRmCd and RmPrtKey=FrPrtKey and RmZ= 'N' " + _
    "  Where " + IIF(gs_Partition = ctCurrPrtn, " FdPrtKey='" + ctCurrPrtn + "' And ", "") + " FdCoCd= '" + ADC("FgCoCd") + "' and FdTc= '" + ADC("FgTc") + "' and FdYy='" + ADC("FgYy") + "' and " + _
    "        FdChr='" + ADC("FgChr") + "' and FdNo=" + CStr(ADC("FgNo")) + " " + _
    "  Group By FdCoCd, FdTc, FdYy, FdChr, FdNo, FdSr " + _
    ") tbl " + _
    "Where qFrWt<> qFdGrWt"
    wFdSrChkSql = moCn.GetStrConcatVal("Ltrim(Str(FdSr))", "", wFdSrChkSql, ", ")
    If wFdSrChkSql <> "" Then Cancel = True: ErrMsg = "Raw Materials Half Copied /Not Copied For Serial Nos: " + wFdSrChkSql: Exit Sub
    '*** Jay 3.01 *** [FgRmWt <> FdGrWt]
  End If
End Sub

Private Sub ATXT_FldWhen(Index As Integer, ByVal IdName As String, Cancel As Boolean, ErrMsg As String)
  '*** Set Helps For Each Field
  '*** Do not allow the user to edit the voucher date if record exists in the 'A' grid
  '*** The help of FgFrRmLoc should show all the locations other than Process locations
  '*** In case of 'FB' entry, the help of FgFrRmLoc should show all the locations of type 'W'
  '*** In case of 'FB' entry, the help of FgToRmLoc should show all the locations of type 'F'
  '*** In case of 'FWB' entry, the help of FgFrRmLoc should show all the locations of type 'F'
  '*** In case of 'FWB' entry, the help of FgToRmLoc should show all the locations of type 'W'
  '*** In case of 'FR' and 'FWR' entries, the help of FgFrRmLoc should show all the locations
      'other than Process locations and 'W' lcations. Similarly for the help of FgToRmLoc
  
  '*** Give help for the field wSummBagGrp in the Summary frame   (PTyp= 'FGBAGGR')
  
  If mf_DtLocked = True And (IdName <> UCase("wSummBagGrp")) Then Cancel = True: ErrMsg = "Cannot Edit Locked Entry": Exit Sub
  If mf_ValidLocSeq = False Then Cancel = True: ErrMsg = "The User Does Not Have Access To One Of The Locs": Exit Sub
  Dim wOthLocBag  As String
  Select Case IdName
  Case Is = UCase("FgYy")
    Call HlpList.PMCd("YY")
  Case Is = UCase("FgChr")
' **** Zubin 211 **** '
    'Call HlpList.vPSCd(ADC("FgCoCd"), "CHR", ADC("FgTc"))
    Call HlpList.vPSCd(ADC("FgCoCd"), "CHR", ADC("FgTc"), "Y")
' **** Zubin 211 **** '
  Case Is = UCase("FgNo")
    Call HlpList.FgNo(ADC("FgCoCd"), ADC("FgTc"), ADC("FgYy"), ADC("FgChr"))
  Case Is = UCase("FgDt")
    If GrdFgd.Rows - 1 > 0 Then Cancel = True: ErrMsg = "Cannot Edit Date": Exit Sub
  Case Is = UCase("FgFrRmLoc")
    If ms_TcTyp = "FB" Then Call HlpList.Loc(ADC("FgCoCd"), "'W'")
    If ms_TcTyp = "FR" Then Call HlpList.Loc(ADC("FgCoCd"), , "'P','W'")
    If ms_TcTyp = "FWB" Then Call HlpList.Loc(ADC("FgCoCd"), "'F'")
    If ms_TcTyp = "FWR" Then Call HlpList.Loc(ADC("FgCoCd"), , "'P','W'")
  Case Is = UCase("FgToRmLoc")
    If ms_TcTyp = "FB" Then Call HlpList.Loc(ADC("FgCoCd"), "'F'")
    If ms_TcTyp = "FR" Then Call HlpList.Loc(ADC("FgCoCd"), , "'P','W'")
    If ms_TcTyp = "FWB" Then Call HlpList.Loc(ADC("FgCoCd"), "'W'")
    If ms_TcTyp = "FWR" Then Call HlpList.Loc(ADC("FgCoCd"), , "'P','W'")
  Case Is = UCase("FgFrBLoc")
    '**********Bhavna 5.1.0
      If ms_TcTyp = "FB" Then
        HlpList.MultiSelect = True
      Else
        HlpList.MultiSelect = False
      End If
      '**********Bhavna 5.1.0
      
    Call HlpList.Loc(ADC("FgCoCd"), "'P'")
  Case Is = UCase("FgToBLoc")
    Call HlpList.Loc(ADC("FgCoCd"), "'P'")
  Case Is = UCase("wExFdExpCmCd")
    Call HlpList.CustCd("C")
  Case Is = UCase("wExFdExpOdTc") 'indigo.164 OS tc added
    Call HlpList.PMCd("TC", "'SO','OS'")
  Case Is = UCase("wExFdExpOdYy")
    Call HlpList.PMCd("YY")
  Case Is = UCase("wExFdExpOdChr")
    Call HlpList.vPSCd(ADC("FgCoCd"), "CHR", ADC("wExFdExpOdTc"))
  Case Is = UCase("wExFdExpOdNo")
    Call HlpList.OmNo(ADC("FgCoCd"), ADC("wExFdExpOdTC"), ADC("wExFdExpOdYy"), ADC("wExFdExpOdChr"))
  Case Is = UCase("wExFdExpOdSr")
    Call HlpList.OdSr(ADC("FgCoCd"), ADC("wExFdExpOdTC"), ADC("wExFdExpOdYy"), ADC("wExFdExpOdChr"), ADC("wExFdExpOdNo"))
  Case Is = UCase("wExFdPlYy")
    Call HlpList.PMCd("YY")
  Case Is = UCase("wExFdPlChr")
    Call HlpList.PMCd("PLCHR")
  Case Is = UCase("wExFdPlNo")
    Call HlpList.PlNo(ADC("FgCoCd"), ADC("wExFdPlYy"), ADC("wExFdPlChr"))
  Case Is = UCase("wExFdInTc")
    Call HlpList.PMCd("TC", "'IN'")
  Case Is = UCase("wExFdInYy")
    Call HlpList.PMCd("YY")
  Case Is = UCase("wExFdInChr")
    Call HlpList.vPSCd(ADC("FgCoCd"), "CHR", ADC("wExFdInTc"))
  Case Is = UCase("wExFdInNo")
    Call HlpList.InNo(ADC("FgCoCd"), ADC("wExFdInTc"), ADC("wExFdInYy"), ADC("wExFdInChr"))
  Case Is = UCase("wFrRmDc")
    Call HlpList.PMCd("DC")
  Case Is = UCase("wExCmName")
    Call HlpList.CmName("C")
  Case Is = UCase("wSummBagGrp")
    Call HlpList.PMCd("FGBAGGR")
  Case Is = UCase("FgSubLoc"), UCase("wCbBFgSubLoc")        ' ***** Manali 3.8.0 - UCase("FgSubLoc") added
    Call HlpList.vPMCd(ADC("FgCoCd"), "FgSubLoc")
  
'Manoj
    Case Is = UCase("wCbBYyFr"), UCase("wCbBYyTo")
    Call HlpList.PMCd("YY")
  Case Is = UCase("wCbBChrFr"), UCase("wCbBChrTo")
' **** Zubin 211 **** '
'''''    Call HlpList.PMCd("BCHR")
    Call HlpList.PMCd("BCHR", "'P'")
' **** Zubin 211 **** '
  Case Is = UCase("wCbBNoFr")
    wOthLocBag = moCn.GetFldVal("Select hOthLocBag from Head where HCoCd= '" + ADC("FgCoCd") + "' and HCd='" + ctSelfCmCd + "'")
    Call HlpList.BNo(ADC("FgCoCd"), ADC("wCbBYyFr"), ADC("wCbBChrFr"), wOthLocBag, ADC("FgFrBLoc"))
  Case Is = UCase("wCbBNoTo")
    wOthLocBag = moCn.GetFldVal("Select hOthLocBag from Head where HCoCd= '" + ADC("FgCoCd") + "' and HCd='" + ctSelfCmCd + "'")
    Call HlpList.BNo(ADC("FgCoCd"), ADC("wCbBYyTo"), ADC("wCbBChrTo"), wOthLocBag, ADC("FgFrBLoc"))
  Case Is = UCase("wCbOdTcFr"), UCase("wCbOdTcTo")  'indigo.164 OS tc added
    Call HlpList.PMCd("TC", "'SO','OS'")
  Case Is = UCase("wCbOdYyFr"), UCase("wCbOdYyTo")
    Call HlpList.PMCd("YY")
  Case Is = UCase("wCbOdChrFr")
    Call HlpList.vPSCd(ADC("FgCoCd"), "CHR", ADC("wCbOdTcFr"))
  Case Is = UCase("wCbOdChrTo")
    Call HlpList.vPSCd(ADC("FgCoCd"), "CHR", ADC("wCbOdTcTo"))
  Case Is = UCase("wCbOdNoFr")
    Call HlpList.OmNo(ADC("FgCoCd"), ADC("wCbOdTcFr"), ADC("wCbOdYyFr"), ADC("wCbOdChrFr"))
  Case Is = UCase("wCbOdNoTo")
    Call HlpList.OmNo(ADC("FgCoCd"), ADC("wCbOdTcFr"), ADC("wCbOdYyFr"), ADC("wCbOdChrFr"))
  Case Is = UCase("wCbOdSrFr")
    Call HlpList.OdSr(ADC("FgCoCd"), ADC("wCbOdTcFr"), ADC("wCbOdYyFr"), ADC("wCbOdChrFr"), ADC("wCbOdNoFr"))
  Case Is = UCase("wCbOdSrTo")
    Call HlpList.OdSr(ADC("FgCoCd"), ADC("wCbOdTcTo"), ADC("wCbOdYyTo"), ADC("wCbOdChrTo"), ADC("wCbOdNoTo"))
  Case Is = UCase("wCbDmCdFr"), UCase("wCbDmCdTo")  'indigo.164 SM design added
    Call HlpList.DmCd("DM','SM")
 '****************Bhavna added cust cd in scope
  Case Is = UCase("wCbCustCdFr"), UCase("wCbCustCdTo")
    Call HlpList.CustCd("C")
  '****************Bhavna added cust cd in scope
   Case Is = UCase("wBrsGrp"), UCase("wVrsGrp")
    Call HlpList.PMCd("TBAGBSGR")
     '*********Bhavna cpy thru bag no
    Case Is = UCase("wCPYTHRUBAGNO")
        If ADC("WCPYTHRUBAGNO") = "" Then
            ADC("WCPYTHRUBAGNO") = "N"
        End If
  'Uma KBS
  Case Is = UCase("WFGPLNO")
    Call HlpList.PlNo(ADC("FgCoCd"), ADC("WFGPLYY"), ADC("WFGPLCHR"))
  Case Is = UCase("WCBFGPLYY"), UCase("WFGPLYY")
      Call HlpList.PMCd("YY")
  Case Is = UCase("WCBFGPLCHR"), UCase("WFGPLCHR")
      Call HlpList.PMCd("PLCHR")
  Case Is = UCase("WCBFGPLNO")
      Call HlpList.PlNo(ADC("FgCoCd"), ADC("WCBFGPLYY"), ADC("WCBFGPLCHR"))
  End Select
End Sub
Private Sub atxt_ReCalculate(Index As Integer, ByVal IdName As String, ByVal OldVal As Variant, ByVal RowNum As Integer, ByVal SourceName As String, ByVal SourceColName As String)
  '*** FgChr on FgTc                Normal
      'PDesc225('TC', adc(FgTc))
  '*** wTotQty on FdQty             Normal, Delete, InitOver
      'wTotQty= Sum(FdQty)
  '*** wTotGrWt on FdGrWt           Normal, Delete, InitOver
      'wTotGrWt= Sum(FdGrWt)
  
  Dim i As Integer
  Dim wCustName As String
  'wCmCd As String
  Select Case UCase(IdName)
  Case Is = UCase("FgChr")
    If ADC.Mode = xNorm Then _
       ADC("FgChr") = moCn.GetFldVal("Select vPValue from vParam where " + _
                                     "vPCoCd= '" + ADC("FgCoCd") + "' and vPTyp= 'DEFCHR' " + _
                                     "and vPMCd= '" + ADC("FgTc") + "'")
       '("Select PDesc225 from Param where PTyp='TC' " + _
                                   "and PMCd='" + ADC("FgTc") + "'")
  Case Is = UCase("wTotQty")
    Dim wTotQty As Single
    If ADC.Mode = xNorm Then
      If GrdFgd.Mode = fgmnorm Or GrdFgd.Mode = fgmdel Or GrdFgd.Mode = fgmInitOver Or GrdFgd.Mode = fgmCopyOver Then
        wTotQty = 0
        For i = 1 To (GrdFgd.Rows - 1)
          If Not GrdFgd.IsDel(i) Then wTotQty = wTotQty + GrdFgd.Value(i, "FdQty")
        Next i
        ADC("wTotQty") = wTotQty
      End If
    End If
  Case Is = UCase("wTotGrWt")
    Dim wTotGrWt As Single
    If ADC.Mode = xNorm Then
      If GrdFgd.Mode = fgmnorm Or GrdFgd.Mode = fgmdel Or GrdFgd.Mode = fgmInitOver Or GrdFgd.Mode = fgmCopyOver Then
        wTotGrWt = 0
        For i = 1 To (GrdFgd.Rows - 1)
          If Not GrdFgd.IsDel(i) Then wTotGrWt = wTotGrWt + GrdFgd.Value(i, "FdGrWt")
        Next i
        ADC("wTotGrWt") = wTotGrWt
      End If
    End If
  
  Case Is = UCase("wExFdExpCmCd")
    If ADC.Mode = xNorm Then
      'wCmCd = moCn.GetFldVal("Select CmCd from CustMst where CmCtg= 'C' and " + _
      '                       "CmName= '" + ADC("wExCmName") + "'")
      'If wCmCd <> ADC("wExFdExpCmCd") Then ADC("wExFdExpCmCd") = wCmCd
      
      '****** Sachin 3.02.0 - Partition Key added
      If ADC("wExFdExpOdNo") <> 0 Then
        ADC("wExFdExpCmCd") = moCn.GetFldVal("Select OmCmCd from OrdMst where " + _
                              "OmCoCd= '" + ADC("FgCoCd") + "' and OmTc= '" + ADC("wExFdExpOdTc") + "' and " + _
                              "OmYy= '" + ADC("wExFdExpOdYy") + "' and OmChr= '" + ADC("wExFdExpOdChr") + "' and " + _
                              "OmNo= " + CStr(ADC("wExFdExpOdNo")) + IIF(gs_Partition = ctCurrPrtn, " and OmPrtKey='" + ctCurrPrtn + "' ", ""))
      Else
        ADC("wExFdExpCmCd") = ""
      End If
    End If
  Case Is = UCase("wExCmName")
    If ADC.Mode = xNorm Then
       wCustName = moCn.GetFldVal("Select CmName from CustMst where CmCtg='C' and " + _
                                  "CmCd='" + ADC("wExFdExpCmCd") + "'")
       If wCustName <> ADC("wExCmName") Then ADC("wExCmName") = wCustName
    End If
  '****** Manali 3.6.0 - 24/11/09 - Defalut To Order Scope for Copy
  Case Is = UCase("wCbOdTcTo")
    If ADC.Mode = xNorm Then ADC("wCbOdTcTo") = ADC("wCbOdTcFr")
  Case Is = UCase("wCbOdYyTo")
    If ADC.Mode = xNorm Then ADC("wCbOdYyTo") = ADC("wCbOdYyFr")
  Case Is = UCase("wCbOdChrTo")
    If ADC.Mode = xNorm Then ADC("wCbOdChrTo") = ADC("wCbOdChrFr")
  Case Is = UCase("wCbOdNoTo")
    If ADC.Mode = xNorm Then ADC("wCbOdNoTo") = ADC("wCbOdNoFr")
  Case Is = UCase("wCbOdSrTo")
    If ADC.Mode = xNorm Then ADC("wCbOdSrTo") = ADC("wCbOdSrFr")
  '****** Manali 3.6.0 - 24/11/09 - Defalut To Order Scope for Copy
  'Rst.125 showing scale wt in text box which is captured in grid
  Case Is = UCase("wBagScaleWt")
    If ADC.Mode = xNorm Then
      If ms_ChkFgWghtYN = "Y" And ms_TcTyp <> "FWB" Then  '5713.59 scale wt not needed for fg return, scale wt in grid and its related calculations are ignored
        If GrdFgd.Row >= 1 Then ADC("wBagScaleWt") = GrdFgd.Value(GrdFgd.Row, "wScaleWt")
      End If
    End If
    
  End Select
End Sub
Private Sub ATXT_FldChange(Index As Integer, ByVal IdName As String, Cancel As Boolean, ErrMsg As String, ByVal pv_NewValue As Variant)
  '*** FgTc             PMCd('TC', PValue= ms_TcTyp) and the type of entry should be 'FgBag' or 'FgRm'
  '*** FgYy             PMCd('YY')
  '*** FgChr            PSCd('CHR', adc(FgTc)) or ''
  '*** FgDt             Not Blank
  '*** FgFrRmLoc        If 'FB' then LocCd from Loc(LocTyp= 'W') or ''
                        'If 'FWB' then LocCd from Loc(LocTyp= 'F') or ''
                        'If 'FR' or 'FWR' then LocCd from Loc(LocTyp<> 'P', 'W') or ''
  '*** FgToRmLoc        If 'FB' then LocCd from Loc(LocTyp= 'F') or ''
                        'If 'FWB' then LocCd from Loc(LocTyp= 'W') or ''
                        'If 'FR' or 'FWR' then LocCd from Loc(LocTyp<> 'P', 'W') or ''
  '*** FgFrBLoc         If 'FB' then LocCd from Loc(LocTyp= 'P' and LocCd<> 'PFG') or ''
                        'If 'FWB' then FgFrBLoc = 'PFG' or ''
  '*** FgToBLoc         If 'FB' then FgToBLoc = 'PFG' or ''
                        'If 'FWB' then LocCd from Loc(LocTyp= 'P' and LocCd<> 'PFG') or ''
  '*** wFrRmDc          PMCd('DC')
  '*** wExFdExpCmCd     CmCd from CustMst(CmCtg= 'C') or ''
  '*** wExFdExpOdTc     PMCd('TC', PValue= 'SO' and the type of entry should be 'OM') or ''
  '*** wExFdExpOdYy     PMCd('YY') or ''
  '*** wExFdExpOdChr    PSCd('CHR', adc(wExFdExpOdTc)) or ''
  '*** wExFdExpOdNo     OdNo from OrdDsg(OdTc= adc(wExFdExpOdTc), OdYy= adc(wExFdExpOdYy)
                        '                OdChr= adc(wExFdExpOdChr)) or 0
  '*** wExFdExpOdSr     OdSr from OrdDsg(OdTc= adc(wExFdExpOdTc), OdYy= adc(wExFdExpOdYy)
                        '                OdChr= adc(wExFdExpOdChr),OdNo= adc(wExFdExpOdNo)) or 0
  '*** wExFdPlChr       PMCd('PLCHR') or ''
  '*** wExFdPlYy        PMCd('YY') or ''
  '*** wExFdInTc        PMCd('TC' and the type of entry should be 'IN') or ''
  '*** wExFdInYy        PMCd('YY') or ''
  '*** wExFdInChr       PSCd('CHR', adc(wExFdInTc)) or ''
  '*** wSummBagGrp  valid PMCd('FGBAGGR') or ''
      '             call the summary routine
  
  Dim i As Integer
  Select Case UCase(IdName)
  Case Is = UCase("FgTc")
    Dim wPValue1 As String
    wPValue1 = moCn.GetFldVal("Select Pvalue1 from Param where PTyp='TCTYP' and " + _
                         "PMCd= '" + ms_TcTyp + "'")
    Cancel = Not moCn.RecSeek("select PMCd from Param where PTyp= 'TC' and " + _
                         "PMCd = '" + pv_NewValue + "' and " + _
                         "PValue= '" + ms_TcTyp + "' and " + _
                         "(Select Pvalue1 from Param where PTyp='TCTYP' and " + _
                         "PMCd= '" + ms_TcTyp + "') ='" + wPValue1 + "'")
    If Cancel = True Then ErrMsg = "Invalid Voucher Tc": Exit Sub
  Case Is = UCase("FgYy")
    Cancel = Not moCn.RecSeek("select PMCd from Param where PTyp= 'YY' and " + _
                         "PMCd = '" + pv_NewValue + "'")
    If Cancel = True Then ErrMsg = "Invalid Voucher Year": Exit Sub
  Case Is = UCase("FgChr")
    Cancel = (pv_NewValue <> "") And (Not moCn.RecSeek("select vPSCd from vParam where " + _
             "vPCoCd = '" + ADC("FgCoCd") + "' and vPTyp= 'CHR' and " + _
             "vPMCd = '" + ADC("FgTc") + "' and vPSCd='" + pv_NewValue + "'"))
' **** Zubin 211 **** '
'''''    ErrMsg = "Invalid Voucher Character": Exit Sub
    If Cancel = True Then ErrMsg = "Invalid Voucher Character": Exit Sub
    Dim ws_UsrArr() As String, wi_Index As Integer
    ws_UsrArr = Split(moCn.GetFldVal("Select vPValue4 from vParam where vPCoCd= '" + ADC("FgCoCd") + "' and " + _
                                        "vPTyp='CHR' and vPMCd='" + ADC("FgTc") + "' and vPSCd= '" + pv_NewValue + "'"), ",")
    If IsArray(ws_UsrArr) Then
      For wi_Index = 0 To UBound(ws_UsrArr)
        Cancel = IIF(UCase(gs_UsrCd) = UCase(ws_UsrArr(wi_Index)), False, True)
        If Cancel = False Then: Exit For
      Next wi_Index
    End If
   If Cancel = True Then ErrMsg = "Access To This Voucher Character Is Denied To User " + gs_UsrCd: Exit Sub
' **** Zubin 211 **** '
  Case Is = UCase("FgDt")
    If pv_NewValue = "" Then Cancel = True: ErrMsg = "Date Cannot Be Blank": Exit Sub
    If pv_NewValue < CDate(ms_LockDt) Then Cancel = True: ErrMsg = "Finished Goods Voucher Date Cannot Be Less Than Lock Date": Exit Sub
    If pv_NewValue > CDate(ms_FutureDt) Then Cancel = True: ErrMsg = "Finished Goods Voucher Date Cannot Be Greater Than Future Date": Exit Sub
    'CC.999.13 check for purge date added
    If pv_NewValue <= CDate(ms_PurgeDt) Then Cancel = True: ErrMsg = "Finished Goods Voucher Date Cannot Be Less Than Or Equal To Purge Date": Exit Sub
    
    
    If moCn.GetFldVal("Select hHolChkYN From Head where HCd='ZSELF' and HCoCd='" + gs_CoCd + "'") = "Y" Then
        '53 company code added
        
        
        If IsHOLYN(ADC("FgCoCd"), pv_NewValue) Then Cancel = True: ErrMsg = "Voucher Cannot be Added on Holidays": Exit Sub
   End If
          
  Case Is = UCase("FgFrRmLoc")
    If ms_TcTyp = "FB" Then
      Cancel = (pv_NewValue <> "") And (Not moCn.RecSeek("Select LocCd from Loc where " + _
                                        "LocCoCd = '" + ADC("FgCoCd") + "' and " + _
                                        "LocCd = '" + pv_NewValue + "' and LocTyp= 'W'"))
    ElseIf ms_TcTyp = "FR" Then
      Cancel = (pv_NewValue <> "") And (Not moCn.RecSeek("Select LocCd from Loc where " + _
                                  "LocCoCd = '" + ADC("FgCoCd") + "' and " + _
                                  "LocCd = '" + pv_NewValue + "' and LocTyp not in ('P', 'W')"))
    ElseIf ms_TcTyp = "FWB" Then
      Cancel = (pv_NewValue <> "") And (Not moCn.RecSeek("Select LocCd from Loc where " + _
                                        "LocCoCd = '" + ADC("FgCoCd") + "' and " + _
                                        "LocCd = '" + pv_NewValue + "' and LocTyp= 'F'"))
    ElseIf ms_TcTyp = "FWR" Then
      Cancel = (pv_NewValue <> "") And (Not moCn.RecSeek("Select LocCd from Loc where " + _
                "LocCoCd = '" + ADC("FgCoCd") + "' and " + _
                "LocCd = '" + pv_NewValue + "' and LocTyp not in('P', 'W')"))
    End If
    If Cancel = True Then ErrMsg = "Invalid Raw Material From Location": Exit Sub
  Case Is = UCase("FgToRmLoc")
    If ms_TcTyp = "FB" Then
      Cancel = (pv_NewValue <> "") And (Not moCn.RecSeek("Select LocCd from Loc where " + _
                                        "LocCoCd = '" + ADC("FgCoCd") + "' and " + _
                                        "LocCd = '" + pv_NewValue + "' and LocTyp= 'F'"))
    ElseIf ms_TcTyp = "FR" Then
      Cancel = (pv_NewValue <> "") And (Not moCn.RecSeek("Select LocCd from Loc where " + _
                "LocCoCd = '" + ADC("FgCoCd") + "' and " + _
                "LocCd = '" + pv_NewValue + "' and LocTyp not in('P', 'W')"))
    ElseIf ms_TcTyp = "FWB" Then
      Cancel = (pv_NewValue <> "") And (Not moCn.RecSeek("Select LocCd from Loc where " + _
                                        "LocCoCd = '" + ADC("FgCoCd") + "' and " + _
                                        "LocCd = '" + pv_NewValue + "' and LocTyp= 'W'"))
    ElseIf ms_TcTyp = "FWR" Then
      Cancel = (pv_NewValue <> "") And (Not moCn.RecSeek("Select LocCd from Loc where " + _
                                  "LocCoCd = '" + ADC("FgCoCd") + "' and " + _
                                  "LocCd = '" + pv_NewValue + "' and LocTyp not in ('P', 'W')"))
    End If
    If Cancel = True Then ErrMsg = "Invalid Raw Material To Location": Exit Sub
  Case Is = UCase("FgFrBLoc")
    If ms_TcTyp = "FB" Then
'        If pv_NewValue = ctFgLoc Then Cancel = True: ErrMsg = "Bag Location Cannot Not Be FG": Exit Sub
        '**********Bhavna 5.1.0  multiselect for 'FB' TC
          Dim ws_LocArr() As String
          ws_LocArr = Split(pv_NewValue, ",")
          If IsArray(ws_LocArr) Then
            For i = 0 To UBound(ws_LocArr)
              Cancel = (ws_LocArr(i) <> "") And (Not moCn.RecSeek("Select LocCd From Loc where " + _
                                        "LocCoCd = '" + ADC("FgCoCd") + "' and " + _
                                        "LocTyp= 'P' and LocCd = '" + ws_LocArr(i) + "'"))
              If Cancel = True Then ErrMsg = "Loc(" + CStr(i + 1) + ") Is An Invalid Process Location": Exit Sub
              
              If (gs_ValidPrdLocSeq <> "" And InStr(1, "," + UCase(gs_ValidPrdLocSeq) + ",", "," + ws_LocArr(i) + ",") <= 0) Then
                  Cancel = True: ErrMsg = "Cannot Edit records Loc." + ws_LocArr(i) + " is Not Accessible to the User": Exit Sub
              End If
              If ws_LocArr(i) = ctFgLoc Then Cancel = True: ErrMsg = "Bag Location Cannot Not Be FG": Exit Sub
            Next i
          End If
          
          '**********Bhavna
          
    Else
        Cancel = (pv_NewValue <> "") And (Not moCn.RecSeek("Select LocCd from Loc where " + _
              "LocCoCd = '" + ADC("FgCoCd") + "' and " + _
              "LocCd = '" + pv_NewValue + "' and LocTyp= 'P'"))
        If Cancel = True Then ErrMsg = "Invalid From Bag Location": Exit Sub
    End If
  Case Is = UCase("FgToBLoc")
    If ms_TcTyp = "FWB" And pv_NewValue = ctFgLoc Then Cancel = True: ErrMsg = "To Bag Location Cannot Not Be FG": Exit Sub
    Cancel = (pv_NewValue <> "") And (Not moCn.RecSeek("Select LocCd from Loc where " + _
              "LocCoCd = '" + ADC("FgCoCd") + "' and " + _
              "LocCd = '" + pv_NewValue + "' and LocTyp= 'P'"))
    If Cancel = True Then ErrMsg = "Invalid To Bag Location": Exit Sub
  Case Is = UCase("wFrRmDc")
    Cancel = Not moCn.RecSeek("select PMCd from Param where PTyp= 'DC' and " + _
                         "PMCd = '" + pv_NewValue + "'")
    If Cancel = True Then ErrMsg = "Invalid From Rm DC": Exit Sub
  Case Is = UCase("wExFdExpCmCd")
    Cancel = (pv_NewValue <> "") And (Not moCn.RecSeek("Select CmCd From CustMst Where CmCtg='C' " + _
                                           "and CmCd='" + pv_NewValue + "'"))
    ErrMsg = "Invalid Exp. Customer Code": Exit Sub
  Case Is = UCase("wExFdExpOdTc")
  'indigo.164 OS tc added
    Cancel = pv_NewValue <> "" And ( _
                    moCn.RecSeek("Select PValue From Param Where PTyp= 'TC' " + _
                    "And PMCd = '" + pv_NewValue + "' And PValue<>'SO' And PValue<>'OS'") _
                    Or _
                    moCn.GetFldVal("select PValue1 from Param where PTyp= 'TcTyp' and " + _
                    "PMCd = (select PValue from Param where PTyp='Tc' and " + _
                    "PMCd='" + pv_NewValue + "')") <> "OM")
    ErrMsg = "Invalid Exp. Order Tc": Exit Sub
  Case Is = UCase("wExFdExpOdYy")
    Cancel = (pv_NewValue <> "") And (Not moCn.RecSeek("Select PMCd From Param Where PTyp= 'Yy' " + _
                                  "And PMCd = '" + pv_NewValue + "'"))
    ErrMsg = "Invalid Exp. Order Yy": Exit Sub
  Case Is = UCase("wExFdExpOdChr")
    Cancel = (pv_NewValue <> "") And (Not moCn.RecSeek("select vPSCd from vParam where " + _
             "vPCoCd = '" + ADC("FgCoCd") + "' and vPTyp= 'CHR' and " + _
             "vPMCd = '" + ADC("wExFdExpOdTc") + "' and vPSCd='" + pv_NewValue + "'"))
    '("Select PMCd From Param Where " + _
                                  "PTyp= 'CHR" + ADC("FgCoCd") + "' " + _
                                  "And PMCd = '" + ADC("wExFdExpOdTc") + "' and PSCd='" + _
                                  pv_NewValue + "'"))
    ErrMsg = "Invalid Exp. Order Chr": Exit Sub
  Case Is = UCase("wExFdExpOdNo")
    '****** Sachin 3.02 - 24-12-07
    Cancel = (pv_NewValue <> 0) And (Not moCn.RecSeek("Select OdNo From OrdDsg Where " + _
                                  "OdCoCd= '" + ADC("FgCoCd") + "' And OdTc= '" + _
                                  ADC("wExFdExpOdTc") + "' And OdYy = '" + _
                                  ADC("wExFdExpOdYy") + "' and OdChr= '" + _
                                  ADC("wExFdExpOdChr") + "' and OdNo= " + _
                                  CStr(pv_NewValue) + IIF(gs_Partition = ctCurrPrtn, " and OdPrtKey='" + ctCurrPrtn + "' ", "")))
    ErrMsg = "Invalid Exp. Order No": Exit Sub
  Case Is = UCase("wExFdExpOdSr")
    '****** Sachin 3.02 - 24-12-07
    Cancel = (pv_NewValue <> 0) And (Not moCn.RecSeek("Select OdNo From OrdDsg Where " + _
                                  "OdCoCd= '" + ADC("FgCoCd") + "' And OdTc= '" + _
                                  ADC("wExFdExpOdTc") + "'And OdYy = '" + _
                                  ADC("wExFdExpOdYy") + "'and OdChr= '" + _
                                  ADC("wExFdExpOdChr") + "'and OdNo= " + _
                                  CStr(ADC("wExFdExpOdNo")) + " and OdSr= " + _
                                  CStr(pv_NewValue) + IIF(gs_Partition = ctCurrPrtn, " and OdPrtKey='" + ctCurrPrtn + "' ", "")))
    ErrMsg = "Invalid Exp. Order Sr": Exit Sub
    
  Case Is = UCase("wExFdPlChr")
    Cancel = (pv_NewValue <> "") And (Not moCn.RecSeek("Select PMCd From Param Where PTyp='PLCHR' " + _
                                  "And PMCd = '" + pv_NewValue + "'"))
    ErrMsg = "Invalid Pl. Chr": Exit Sub
  Case Is = UCase("wExFdPlYy")
    Cancel = (pv_NewValue <> "") And (Not moCn.RecSeek("Select PMCd From Param Where PTyp= 'Yy' " + _
                                  "And PMCd = '" + pv_NewValue + "'"))
    ErrMsg = "Invalid Pl.Yy": Exit Sub
  Case Is = UCase("wExFdInTc")
    Cancel = (pv_NewValue <> "") And (Not moCn.RecSeek("Select PMCd From Param Where PTyp= 'Tc' " + _
                                    "And PMCd = '" + pv_NewValue + "'") _
                                     Or _
                                    (moCn.GetFldVal("select PValue1 from Param where PTyp= 'TcTyp' and " + _
                                    "PMCd = (select PValue from Param where PTyp='Tc' and " + _
                                    "PMCd='" + pv_NewValue + "')") <> "IN"))
    ErrMsg = "Invalid Invoice Tc": Exit Sub
  Case Is = UCase("wExFdInYy")
    Cancel = (pv_NewValue <> "") And (Not moCn.RecSeek("Select PMCd From Param Where PTyp= 'Yy' " + _
                                  "And PMCd = '" + pv_NewValue + "'"))
    ErrMsg = "Invalid Invoice Yy": Exit Sub
  Case Is = UCase("wExFdInChr")
    Cancel = (pv_NewValue <> "") And (Not moCn.RecSeek("select vPSCd from vParam where " + _
             "vPCoCd = '" + ADC("FgCoCd") + "' and vPTyp= 'CHR' and " + _
             "vPMCd = '" + ADC("wExFdInTc") + "' and vPSCd='" + pv_NewValue + "'"))
    '("Select PMCd From Param Where " + _
                                  "PTyp= 'CHR" + ADC("FgCoCd") + "' " + _
                                  "And PMCd = '" + ADC("wExFdInTc") + "' and PSCd='" + _
                                  pv_NewValue + "'"))
    ErrMsg = "Invalid Invoice Chr": Exit Sub
  Case Is = UCase("wExFdInNo")
    '###
    'Cancel = (pv_NewValue <> "") And (Not moCn.RecSeek("Select InNo From InvHd Where " + _
    '                                "InTc= '" + ADC("wExFdInTc") + "' And " + _
    '                                "InYy= '" + ADC("wExFdInYy") + "' and " + _
    '                                "InChr= '" + ADC("wExFdInChr") + "' and " + _
    '                                "InNo= " + CStr(pv_NewValue)))
    'ErrMsg = "Invalid Invoice Number": Exit Sub
    '###
  Case Is = UCase("wExCmName")
    Cancel = pv_NewValue <> "" And (Not moCn.RecSeek("select CmName from CustMst where CmCtg= 'C' and CmName = '" + pv_NewValue + "'"))
    If Cancel = True Then ErrMsg = "Invalid Customer Name": Exit Sub
  Case Is = UCase("wSummBagGrp")
    If pv_NewValue <> "" Then
      Cancel = Not moCn.RecSeek("Select PMCd From Param where PTyp= 'FGBAGGR' " + _
                                "and PMCd = '" + pv_NewValue + "'")
      If Cancel = True Then ErrMsg = "Invalid Summary Group Option": Exit Sub
      Call Summary(pv_NewValue)
    End If
  Case Is = UCase("FgSubLoc")
    If ms_TcTyp = "FB" Then
      Cancel = Not moCn.RecSeek("Select vPmCd from vParam where " + _
              "vPCoCd = '" + ADC("FgCoCd") + "' and " + _
              "vPTyp = 'FGSUBLOC' and vPMcd = '" + pv_NewValue + "'")
      If Cancel = True Then ErrMsg = "Invalid Fg Sub Location": Exit Sub
      
      Cancel = Not moCn.RecSeek("select 'x' From vParam where vPTyp='FgSubLoc' and vPCoCd= '" + gs_CoCd + "' and vPMCd='" + pv_NewValue + "'" + _
                    " and (vPValue4='' Or ','+vPValue4+',' Like '%," + gs_UsrCd + ",%')")
                         
      If Cancel Then ErrMsg = "The User Does Not Have Access To This Sub Location.": Exit Sub
      
      Cancel = moCn.RecSeek("Select 'x' from Fgd Where " + IIF(gs_Partition = ctCurrPrtn, " FdPrtKey='" + ctCurrPrtn + "' And ", "") + " FdCoCd= '" + ADC("FgCoCd") + "' And " + _
              "FdTc= '" + ADC("FgTc") + "' and FdYy= '" + ADC("FgYy") + "' and FdChr= '" + ADC("FgChr") + "' and FdNo= " + CStr(ADC("FgNo")))
      If Cancel = True Then ErrMsg = "Cannot change when Fgd records exist": Exit Sub
    Else
      If pv_NewValue <> "" Then Cancel = True: ErrMsg = "Fg SubLoc Should be Blank": Exit Sub
    End If
  
  ' ***** Manali 3.8.0 -
  Case Is = UCase("wCbBFgSubLoc")
    If ms_TcTyp = "FWB" Then
      Cancel = pv_NewValue <> "" And Not moCn.RecSeek("Select vPmCd from vParam where " + _
              "vPCoCd = '" + ADC("FgCoCd") + "' and " + _
              "vPTyp = 'FGSUBLOC' and vPMcd = '" + pv_NewValue + "'")
      If Cancel = True Then ErrMsg = "Invalid Fg Sub Location": Exit Sub
    Else
      If pv_NewValue <> "" Then Cancel = True: ErrMsg = "Fg SubLoc Should be Blank": Exit Sub
    End If
  
  Case Is = UCase("wBrsGrp")
    Cancel = Not moCn.RecSeek("select PMCd from Param where PTyp= 'TBAGBSGR' and " + _
                         "PMCd = '" + pv_NewValue + "'")
    If Cancel = True Then ErrMsg = "Invalid Group By Option": Exit Sub
    Call RMSummary(pv_NewValue, "N")
    
  Case Is = UCase("wVrsGrp")
    Cancel = Not moCn.RecSeek("select PMCd from Param where PTyp= 'TBAGBSGR' and " + _
                         "PMCd = '" + pv_NewValue + "'")
    If Cancel = True Then ErrMsg = "Invalid Group By Option": Exit Sub
    Call RMSummary(pv_NewValue, "Y")
'**** Bhavna added cpy thru bag no
  Case Is = UCase("wCpyThruBagNo")
    If ADC.Mode = xNorm Then
    'Indigo.84-if Copy RfId is yes then update Pl NO. option will be disable else enable
      If pv_NewValue = "Y" Then OptUpdPlNo.Enabled = True Else OptUpdPlNo.Enabled = False
      Cancel = Not moCn.RecSeek("select PMCd from Param where PTyp= 'YN' and PMCd = '" + pv_NewValue + "'")
      ErrMsg = "Enter 'Y' to Copy thru Bag No Or 'N' to Copy thru RfId No ": Exit Sub
    End If
          
  'Uma KBS
  Case Is = UCase("WFGPLYY")
    If pv_NewValue <> "" Then
      Cancel = Not moCn.RecSeek("Select PMCd From Param Where PTyp= 'Yy' And PMCd = '" + pv_NewValue + "' And PSCd=''")
      ErrMsg = "Invalid Packing List Year": Exit Sub
    End If
  Case Is = UCase("WFGPLCHR")
     If pv_NewValue <> "" Then
      Cancel = Not moCn.RecSeek("Select PMCd From Param Where PTyp= 'PLCHR' And PMCd = '" + pv_NewValue + "' And PSCd=''")
      ErrMsg = "Invalid Packing List Character": Exit Sub
    End If
  Case Is = UCase("WCBFGPLYY")
    If pv_NewValue <> "" Then
      Cancel = Not moCn.RecSeek("Select PMCd From Param Where PTyp= 'Yy' And PMCd = '" + pv_NewValue + "' And PSCd=''")
      ErrMsg = "Invalid Packing List Year": Exit Sub
    End If
  Case Is = UCase("WCBFGPLCHR")
    If pv_NewValue <> "" Then
      Cancel = Not moCn.RecSeek("Select PMCd From Param Where PTyp= 'PLCHR' And PMCd = '" + pv_NewValue + "' And PSCd=''")
      ErrMsg = "Invalid Packing List Character": Exit Sub
    End If
  End Select
End Sub

Private Sub CmdCpyRfId_Click()
  If Not ADC.SaveRec Then
    DispMsg "X Save Failed", etInfo
    Exit Sub
  End If
  If mf_ValidLocSeq = False Then DispMsg "The User Does Not Have Access To One Of The Locs", etError: Exit Sub
  
  'If ADC("FGFRBLOC") = "" Then DispMsg "Enter Bag From Location.", etError: Exit Sub

  Call DispFra(CpyRfId)

End Sub

Private Sub CmdCpyRfIdGo_Click()
  If mf_DtLocked = True Then
    DispMsg "Cannot Copy Bags For Locked Entry", etError
  Else
    If mf_ValidLocSeq = False Then
      DispMsg "The User Does Not Have Access To One Of The Locs", etError
    Else
      'Indigo.84-If CopyBags option is selected system will go for normal routine, if Update PlNo. is selected it calls UpdPlNo
      If OptUpdPlNo.Value = True Then
        Call UpdPLNo
      Else
        Call CopyBag(True)
      End If
    End If
  End If
  Call CmdCpyRfId_Click

End Sub

Private Sub CmdInvHd_Click()
       If Not ADC.SaveRec Then
    DispMsg "X Save Failed", etInfo
    Exit Sub
  End If
  If moCn.RecSeek("Select 'x' From UsrAccess where UaUsrCd='" + ADC.UsrCd + "'  and UaMnuCd='In'") Then
      If gs_InNo = 0 Then
       gs_InCoCd = ADC("FgCoCd")
       gs_InTc = ADC("FgTc")
       gs_InYy = ADC("FgYy")
       gs_InChr = ADC("FgChr")
       gs_InNo = ADC("FgNo")
       FormSelect ("In")
      Else
       DispMsg "Invoice Header already open.", etError
      End If
  Else
      DispMsg "User has no access to Invoice Header menu.", etError
  End If
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
' ***** Manali 3.10.0 - 09/05/12 - Bar Code (Ctrl+N)
  If Not (KeyCode = vbKeyN And Shift = 2) Then Exit Sub
  If UCase(Me.ActiveControl.Name) = UCase("GrdFgd") Then
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
If UCase(Me.ActiveControl.Name) = UCase("GrdFgd") Then
  With Me.ActiveControl
      If .Col = 2 Then
          If .Rows - 1 = .Row Then .Col = 3
    End If
  End With
End If
' ***** Manali 3.10.0 - 12/04/12 - Bar Code
End Sub
Private Sub GrdFgd_RowPostValid(ByVal RowNum As Integer, Cancel As Boolean, ErrMsg As String, ChangedColName As String)
  If moCn.RecSeek("select 'A' from BAG " + _
    " join OrdMst on OmIdNo=BOmIdNo and OmCoCd= BCoCd and OmTc=BOdTc and OmYy= BOdYy and OmChr= BOdChr and OmNo= BOdNo " + _
    " Where '" + ADC("fGcocd") + "' = BCoCd And BYY='" + GrdFgd.Value(RowNum, "FdBYy") + "' and BChr ='" + GrdFgd.Value(RowNum, "FdBChr") + "' and " + _
      CStr(GrdFgd.Value(RowNum, "FdBNo")) + " = BNo  AND " + _
    " ('" + GrdFgd.Value(RowNum, "FdDmCd") + "' <> BOdDmCd or '" + GrdFgd.Value(RowNum, "FdSfx") + "' <> BOdSfx or '" + GrdFgd.Value(RowNum, "FdDmSz") + "' <> BOdDmSz )  ") Then
      ErrMsg = "Design/Customer Code Fields Do Not Match With Bag"
      Cancel = True
  End If
  
End Sub
Private Sub GrdFgd_RowSave(ByVal RowNum As Integer, Cancel As Boolean, ErrMsg As String)
  With GrdFgd
     Dim mb_SeoUsr As Boolean, mb_VldSeoPwd As Boolean, wValChg As Boolean
    'Emr 3.11.0 - Tolerance Check by Pct for 'D' and 'C'
     mb_SeoUsr = moCn.RecSeek(" Select 'x' From Param Join UsrAccess on UaUsrCd = PMCd where PTyp= 'USR' and PMCd= '" + gs_UsrCd + "' " + _
                              " and UaMnuCd = '" + ADC("FgTc") + "' and UaSeoYn= 'Y' and PValue= '" + Trim(TxtPwd.text) + "'")
     mb_VldSeoPwd = moCn.RecSeek(" Select 'x' from Param Where PTyp= 'SEOPWD' and PMCD= 'SEOPWD' and " + _
                                 " PValue1='" + Trim(TxtPwd.text) + "'")
     If Not (mb_SeoUsr Or mb_VldSeoPwd) Then
          Dim wErrDesc As String, wRsOrdRmLst As MwfLib.MDORowSet
            wValChg = Not moCn.RecSeek(" Select 'x' From Fgd Where FdPrtKey='" + .Value(RowNum, "FdPrtKey") + "' " + _
                            " and FdCoCd='" + ADC("FgCoCd") + "' " + _
                            " and FdTc='" + ADC("FgTc") + "' " + _
                            " and FdYy='" + ADC("FgYy") + "' " + _
                            " and FdChr='" + ADC("FgChr") + "' " + _
                            " and FdNo=" + CStr(ADC("FgNo")) + _
                            " and FdSr=" + CStr(.Value(RowNum, "FdSr")) + _
                            " and FdBYy='" + .Value(RowNum, "FdBYy") + "' " + _
                            " and FdBChr='" + .Value(RowNum, "FdBChr") + "' " + _
                            " and FdBNo=" + CStr(.Value(RowNum, "FdBNo")) + _
                            " and FdSubLoc='" + .Value(RowNum, "FdSubLoc") + "' ")

        'Bypass the tolerance check for trial chr=Y
        Dim wChkOdChr As Boolean
        wChkOdChr = moCn.RecSeek("select BOdChr from Bag join vParam on vPCoCd ='" + gs_CoCd + _
                    "' and vPTyp ='CHR' and vPMCd='SO' and vPSCd =BOdChr " + _
                    "Where BCoCd='" + gs_CoCd + "' And BYy='" + .Value(RowNum, "FdBYy") + "' And BChr='" + .Value(RowNum, "FdBChr") + "' And BNo=" + CStr(.Value(RowNum, "FdBNo")) + " and vPValue3 ='N' ")
        If wChkOdChr Then
          If wValChg Then
             Dim wb_DExists As Boolean, wb_CExists As Boolean
             wb_DExists = False: wb_CExists = False
            Set wRsOrdRmLst = moCn.OpenRes(" Select Distinct OrRmCtg,(Case when OrRmCtg In ('D','C') Then OrRmSCtg else '' End) as OrRmSCtg from OrdRm " + _
                                                 "    Join Bag On BCoCd=OrCoCd And BOdTc=OrTc And BOdYy=OrYy And BOdChr=OrChr And BOdNo=OrNo And BOdSr=OrSr " + _
                                                 " Where BCoCd='" + gs_CoCd + "' And BYy='" + .Value(RowNum, "FdBYy") + "' And BChr='" + .Value(RowNum, "FdBChr") + "' And BNo=" + CStr(.Value(RowNum, "FdBNo")))
              Do While Not (wRsOrdRmLst.EOF Or wRsOrdRmLst.BOF)
                  wErrDesc = ChkTolerance(gs_CoCd, .Value(RowNum, "FdBYy"), .Value(RowNum, "FdBChr"), .Value(RowNum, "FdBNo"), wRsOrdRmLst!OrRmCtg, wRsOrdRmLst!OrRmSCtg, True, "", "", "", 0, 0, 0, 0, 0, True, True)
                  If wErrDesc <> "" Then Cancel = True: ErrMsg = wErrDesc: Exit Sub
                  If (Not wb_DExists And wRsOrdRmLst!OrRmCtg = "D") Or (Not wb_CExists And wRsOrdRmLst!OrRmCtg = "C") Then
                      If wRsOrdRmLst!OrRmCtg = "D" Then wb_DExists = True
                      If wRsOrdRmLst!OrRmCtg = "C" Then wb_CExists = True
                      wErrDesc = ChkTolerance(gs_CoCd, .Value(RowNum, "FdBYy"), .Value(RowNum, "FdBChr"), .Value(RowNum, "FdBNo"), wRsOrdRmLst!OrRmCtg, "", True, "", "", "", 0, 0, 0, 0, 0, True, True)
                  End If
                  If wErrDesc <> "" Then Cancel = True: ErrMsg = wErrDesc: Exit Sub
                  wRsOrdRmLst.MoveNext
              Loop
          End If
        End If
      End If
      'Rst.125 checking wt difference of scale wt and existing bag wt, if diff more than 0.099 then donot allow to save record
      '5713.59 scale wt not needed for fg return, scale wt in grid and its related calculations are ignored
      If ms_ChkFgWghtYN = "Y" And ADC("wBagScaleWt") <> 0 And ms_TcTyp <> "FWB" Then
        Dim wDiffWt As Double
        wDiffWt = .Value(RowNum, "wScaleWt") - .Value(RowNum, "FdGrWt")
        
        If wDiffWt > 0.099 Or wDiffWt < -0.099 Then
          Cancel = True: ErrMsg = "Difference is " + CStr(wDiffWt) + ". Allowed Bag difference is +/- 0.099 gms only. Please change the Rm Weight in Bag & then Proceed.": Exit Sub
        End If
      End If
  End With
End Sub
Private Sub GrdFgd_When()
  '*** Enable all the command buttons associated with the Grid 'A'
 If mb_ARowExists = False Then
    '*** Set The Visible, Enable Properties For All The Command Buttons ***
    Call CmdPropSet
    mb_ARowExists = True
  End If
  If CmdCP.Visible = True Then CmdCP.Enabled = True
  If CmdEX.Visible = True Then CmdEX.Enabled = True
 ' If CmdSumm.Visible = True Then CmdSumm.Enabled = True

End Sub
Private Sub GrdFgd_RowWhen(ByVal RowNum As Integer)
  '*** If TcTyp = 'FR' Then set AllowAdd and AllowDelete property of GrdFgRm to 'True'
  '*** If GrdFgd.IsNew Then FdSr = Max + 1 and Autogenerate FdKey and
      'if 'FB' or 'FR' then FdRefYy= FgYy and FdRefKey= FdKey
  '*** If FdDespQty <> 0 Then set AllowAdd and AllowDelete property of GrdFgRm to 'False'
  
  With GrdFgd
    ADC("wFgDesc") = .Value(RowNum, "FgDesc")
    If ms_TcTyp = "FR" Then GrdFgRm.AllowAdd = True: GrdFgRm.AllowDelete = True
   
      If GrdFgd.IsNew(RowNum) Then
      If .Value(RowNum, "FdSr") = 0 Then .Value(RowNum, "FdSr") = .MaxVal("FdSr") + 1
      .Value(RowNum, "FdKey") = moCn.AutoGen("Fgd", "FdKey", ADC("FgCoCd") + ADC("FgYy"))
      If ms_TcTyp = "FB" Or ms_TcTyp = "FR" Then
        .Value(RowNum, "FdRefYy") = ADC("FgYy")
        .Value(RowNum, "FdRefKey") = .Value(RowNum, "FdKey")
      End If
      .Value(RowNum, "FdPrtKey") = ctCurrPrtn       '****** Sachin 3.02
          
      If ms_TcTyp = "FB" Then .Value(RowNum, "FdSubLoc") = ADC("FgSubLoc")
      'Uma KBS
      'MW.12- If the entry is FWB then no need to write Pl no. to grid
      If ms_TcTyp <> "FWB" Then
        If .Value(RowNum, "FdSr") = 1 Then
          .Value(RowNum, "FdPlYy") = ADC("WFGPLYY")
          .Value(RowNum, "FdPlChr") = ADC("WFGPLCHR")
          .Value(RowNum, "FdPlNo") = ADC("WFGPLNO")
        Else
          .Value(RowNum, "FdPlYy") = .Value(RowNum - 1, "FdPlYy")
          .Value(RowNum, "FdPlChr") = .Value(RowNum - 1, "FdPlChr")
          .Value(RowNum, "FdPlNo") = .Value(RowNum - 1, "FdPlNo")
        End If
      End If
                                  
    End If
    If .Value(RowNum, "FdDespQty") <> 0 Then GrdFgRm.AllowAdd = False: GrdFgRm.AllowDelete = False
    'Rst.125 if scale wt is Yes then showing scale wt in text box and disable copy buttons
    '5713.59 scale wt not needed for fg return, scale wt in grid and its related calculations are ignored
    If ms_ChkFgWghtYN = "Y" And ms_TcTyp <> "FWB" Then
      ADC("wBagScaleWt") = .Value(RowNum, "wScaleWt")
      CmdCB.Enabled = False
      CmdCpyRfId.Enabled = False
    End If
  End With
  
  mb_AddRow = True    ' ***** Manali 3.10.0 - 09/05/12 - Bar Code (Ctrl+N)
End Sub
Private Sub GrdFgd_ColWhen(ByVal RowNum As Integer, ByVal ColName As String, Cancel As Boolean, ErrMsg As String, ByVal Action As MwfCtl.en_fgAction)
  '*** Check if Despatch Qty = 0
      'If not then do no allow the user to edit any field in the grid
  '*** Set Helps For Each Field
  '*** If 1st row, then get the default values for FdBYy and FdBChr from the Head file
      'If Row> 1 then get the defa values for FdBYy and FdBChr from the previous row of the grid
      'If not GrdFgd.IsNew the cannot edit FdBYy, FdBChr and FdBNo
      'If the value of the field hOthLocBag of the Head table is 'N' then only bags belonging
      'to FgFrBLoc is shown in the help of FdBNo
      'If hOthLocBag = 'Y' then bags belonging to all the Process locations are shown in the
      'help of FdBNo. The bags belonging to FgFrBLoc Location are shown first.
  
  If mf_DtLocked = True Then Cancel = True: ErrMsg = "Cannot Edit Locked Entry": Exit Sub
  If mf_ValidLocSeq = False Then Cancel = True: ErrMsg = "The User Does Not Have Access To One Of The Locs": Exit Sub
  With GrdFgd
  Dim wBFrLoc As String
  If Not .IsNew(RowNum) Then
    If gs_ValidPrdLocSeq <> "" And .Value(RowNum, "FdBNo") <> 0 Then
          wBFrLoc = moCn.GetFldVal("Select BLoc from Bag where " + _
                                   "BCoCd='" + ADC("FgCoCd") + "' and " + _
                                   "BYy='" + .Value(RowNum, "FdBYy") + "' and " + _
                                   "BChr='" + .Value(RowNum, "FdBChr") + "' and " + _
                                   "BNo= " + CStr(.Value(RowNum, "FdBNo")) + " And BPrtKey='" + ctCurrPrtn + "' ")
          Cancel = IIF((InStr(1, "," + UCase(gs_ValidPrdLocSeq) + ",", "," + UCase(wBFrLoc) + ",") <= 0), True, False)
        Else
          Cancel = False
        End If
        If Cancel = True Then
          ErrMsg = "Cannot Edit record As Bag Loc. is Not Accessible to the User": Exit Sub
         ' GrdFgd.AllowAdd = False
    End If
  End If
          Dim PrevRowNum As Integer
      PrevRowNum = PrevRowNum - 1
      'PrevRowNum > 0 And PrevRowNum <= GrdFgd.Rows - 1
      If (1 = 1) Then
        Dim wFdSrChkSql As String
        wFdSrChkSql = " ( Select FdSr, Convert(Decimal (14,3), Max(FdGrWt)) as qFdGrWt, " + _
        "  Convert(Decimal (14,3), Isnull(Sum(FrRmWt / (Case when RmCtg in ('D','C') then 5 else 1 end)),0))  as qFrWt " + _
        "  From Fgd " + _
        "     Left Outer Join FgRm On FrFdIdNo=FdIdNo and FrPrtKey=FdPrtKey and FrCoCd = FdCoCd and FrTc= FdTc and FrYy= FdYy and FrChr= FdChr and FrNo= FdNo and FrSr= FdSr " + _
        "     Left Outer Join RmMst on RmCd= FrRmCd and RmPrtKey=FrPrtKey and RmZ= 'N' " + _
        "  Where " + IIF(gs_Partition = ctCurrPrtn, " FdPrtKey='" + ctCurrPrtn + "' And ", "") + " FdCoCd= '" + ADC("FgCoCd") + "' and FdTc= '" + ADC("FgTc") + "' and FdYy='" + ADC("FgYy") + "' and " + _
        "        FdChr='" + ADC("FgChr") + "' and FdNo=" + CStr(ADC("FgNo")) + " " + _
        "  Group By FdCoCd, FdTc, FdYy, FdChr, FdNo, FdSr " + _
        ") tbl " + _
        "Where qFrWt<> qFdGrWt"
        '5717 below 2 line comments removed to get this error when press ctrl+N - as like 577
        wFdSrChkSql = moCn.GetStrConcatVal("Ltrim(Str(FdSr))", "", wFdSrChkSql, ", ")
        If wFdSrChkSql <> "" Then Cancel = True: ErrMsg = ("Raw Materials Half Copied /Not Copied For Serial Nos: " + wFdSrChkSql): MsgBox ("Raw Materials Half Copied /Not Copied For Serial Nos: " + wFdSrChkSql): Exit Sub
        Cancel = False
      End If

  
    If .Value(RowNum, "FdDespQty") <> 0 Then Cancel = True: Exit Sub
    Select Case ColName
' **** Zubin 211 **** '
    Case Is = UCase("wFgBNo")
      If Not .IsNew(RowNum) Then Cancel = True: ErrMsg = "Cannot Edit FG Bag Number": Exit Sub
      If Not .IsNew(RowNum) And .Value(RowNum, "FdBNo") <> 0 Then Cancel = True: ErrMsg = "Cannot Edit FG Bag Number": Exit Sub
' **** Zubin 211 **** '
    Case Is = UCase("FdBYy")
' **** Zubin 211 **** '
'''''      If .Value(RowNum, "wFGBNo") <> "" Then Cancel = True: ErrMsg = "Cannot Edit Bag Year When Bag Number Is Read Using Barcode Reader": Exit Sub
' **** Zubin 211 **** '
      If .IsNew(RowNum) And .Value(RowNum, "FdBYy") = "" Then
        If RowNum = 1 Then .Value(RowNum, "FdBYy") = _
           moCn.GetFldVal("Select hYy from Head where HCoCd= '" + ADC("FgCoCd") + "' and HCd='" + ctSelfCmCd + "'")
        If RowNum > 1 Then .Value(RowNum, "FdBYy") = .Value(RowNum - 1, "FdBYy")
      Else
           Cancel = True: ErrMsg = "Cannot Edit Bag Year": Exit Sub
      End If
      Call HlpList.PMCd("YY")
    Case Is = UCase("FdBChr")
' **** Zubin 211 **** '
'''''      If .Value(RowNum, "wFGBNo") <> "" Then Cancel = True: ErrMsg = "Cannot Edit Bag Character When Bag Number Is Read Using Barcode Reader": Exit Sub
' **** Zubin 211 **** '
      If .IsNew(RowNum) And .Value(RowNum, "FdBChr") = "" Then
        If RowNum = 1 Then .Value(RowNum, "FdBChr") = _
           moCn.GetFldVal("Select hBChr from Head where HCoCd= '" + ADC("FgCoCd") + "' and HCd='" + ctSelfCmCd + "'")
        If RowNum > 1 Then .Value(RowNum, "FdBChr") = .Value(RowNum - 1, "FdBChr")
      Else
        Cancel = True: ErrMsg = "Cannot Edit Bag Character": Exit Sub
      End If
' **** Zubin 211 **** '
      Call HlpList.PMCd("BCHR")   'uma *** to add PD bag
      'Call HlpList.PMCd("BCHR", "'P'")
' **** Zubin 211 **** '
    Case Is = UCase("FdBNo")
' **** Zubin 211 **** '
'''''      If .Value(RowNum, "wFGBNo") <> "" Then Cancel = True: ErrMsg = "Cannot Edit Bag Number When Bag Number Is Read Using Barcode Reader": Exit Sub
' **** Zubin 211 **** '
      If Not .IsNew(RowNum) Then _
         Cancel = True: ErrMsg = "Cannot Edit Bag Number": Exit Sub
      '6.1_20 once bagno. entered we cant change it again
      If .IsNew(RowNum) And .Value(RowNum, "FdBNo") <> 0 Then Cancel = True: ErrMsg = "Cannot Edit Bag Number": Exit Sub
      Dim whOthLocBag As String
      whOthLocBag = moCn.GetFldVal("Select hOthLocBag from Head where HCoCd= '" + ADC("FgCoCd") + "' and HCd='" + ctSelfCmCd + "'")
      Call HlpList.BNo(ADC("FgCoCd"), .Value(RowNum, "FdBYy"), .Value(RowNum, "FdBChr"), whOthLocBag, ADC("FgFrBLoc"))
    Case Is = UCase("wFdBaseVchNo")
      Call HlpList.BaseFgdSr(ADC("FgCoCd"), ms_EntTyp)
    Case Is = UCase("FdDmCd") 'indigo.164 SM design added
      Call HlpList.DmCd("DM','SM")
    Case Is = UCase("FdDmSz") 'indigo.164 SM design added
      Call HlpList.DmSz("DM','SM", .Value(RowNum, "FdDmCd"))
    
    '###
    Case Is = UCase("FdPrdOdTc")  'indigo.164 OS tc added
      Call HlpList.PMCd("TC", "'SO','OS'")
    Case Is = UCase("FdPrdOdYy")
      Call HlpList.PMCd("YY")
    Case Is = UCase("FdPrdOdChr")
      Call HlpList.vPSCd(ADC("FgCoCd"), "CHR", ADC("FdPrdOdTc"))
    Case Is = UCase("FdPrdOdNo")
      Call HlpList.OmNo(ADC("FgCoCd"), .Value(RowNum, "FdPrdOdTC"), .Value(RowNum, "FdPrdOdYy"), .Value(RowNum, "FdPrdOdChr"))
    Case Is = UCase("FdPrdOdSr")
      Call HlpList.OdSr(ADC("FgCoCd"), .Value(RowNum, "FdPrdOdTC"), .Value(RowNum, "FdPrdOdYy"), .Value(RowNum, "FdPrdOdChr"), .Value(RowNum, "FdPrdOdNo"))
    '###
    Case Is = UCase("FgDesc")
      ADC("wFgDesc").SetFocus
    
    Case Is = UCase("FdSubLoc")
      Call HlpList.vPMCd(ADC("FgCoCd"), "FgSubLoc")
    
    'Uma KBS
    Case Is = UCase("FdPlYy")
      Call HlpList.PMCd("YY")
    Case Is = UCase("FdPlChr")
      Call HlpList.PMCd("PLCHR")
    Case Is = UCase("FdPlNo")
      Call HlpList.PlNo(ADC("FgCoCd"), .Value(RowNum, "FdPlYy"), .Value(RowNum, "FdPlChr"))
    
    End Select
  End With
End Sub
Private Sub GrdFgd_ReCalc(ByVal RowNum As Integer, ByVal ColName As String, ByVal OldVal As Variant, ByVal SourceName As String, ByVal SourceColName As String)
  '*** FdBNo on FdBYy,FdBChr                                    Normal
      'FdBNo = 0
  '*** wFdBaseVchNo on FdBYy,FdBChr,FdBNo,FdRefYy,FdRefKey      Normal,init
      'If mode= Normal and TcType= 'FWB' and (SourceColName= 'FdBYy' Or 'FdBChr' Or 'FdBNo') then
      '   wFdBaseVchNo = FdTc+'/'+FdYy+'/'+FdChr+'/'+FdNo+'/'+FdSr from
      '                  Fgd(FdBM= 'B', FdBYy= FdBYy of Grid, FdBChr= FdBChr of Grid,
      '                  FdBNo= FdBNo of grid and FdQty> FdDespQty)
      'If mode= init and (TcType= 'FWB' or 'FWR') and (SourceColName= 'FdRefYy' Or 'FdRefYy') then
      '   wFdBaseVchNo = FdTc+'/'+FdYy+'/'+FdChr+'/'+FdNo+'/'+FdSr from
      '                  Fgd(FdBM= 'B', FdYy= FdRefYy of Grid, FdKey= FdRefKey of Grid)
  '*** FdDmCd on FdBYy,FdBChr,FdBNo,wFdBaseVchNo                Normal
      'If TcType= 'FB' and (SourceColName= 'FdBYy' Or 'FdBChr' Or 'FdBNo') then
      '   FdDmCd= BOdDmCd from Bag(BYy= FdBYy, BChr= FdBChr, BNo= FdBNo)
      'If (TcType= 'FWB' or 'FWR') and (SourceColName= 'wFdBaseVchNo') then
      '   FdDmCd= FdDmCd from the Fgd record equal to wFdBaseVchNo
  '*** FdSfx on FdBYy,FdBChr,FdBNo,wFdBaseVchNo                 Normal                              Normal
      'If TcType= 'FB' and (SourceColName= 'FdBYy' Or 'FdBChr' Or 'FdBNo') then
      '   FdSfx= BOdSfx from Bag(BYy= FdBYy, BChr= FdBChr, BNo= FdBNo)
      'If (TcType= 'FWB' or 'FWR') and (SourceColName= 'wFdBaseVchNo') then
      '   FdSfx= FdSfx from the Fgd record equal to wFdBaseVchNo
  '*** FdDmSz on FdBYy,FdBChr,FdBNo,wFdBaseVchNo                Normal
      'If TcType= 'FB' and (SourceColName= 'FdBYy' Or 'FdBChr' Or 'FdBNo') then
      '   FdDmSz= BOdDmSz from Bag(BYy= FdBYy, BChr= FdBChr, BNo= FdBNo)
      'If (TcType= 'FWB' or 'FWR') and (SourceColName= 'wFdBaseVchNo') then
      '   FdDmSz= FdDmSz from the Fgd record equal to wFdBaseVchNo
  '*** FdQty on FdBYy,FdBChr,FdBNo,wFdBaseVchNo                 Normal
      'If TcType= 'FB' and (SourceColName= 'FdBYy' Or 'FdBChr' Or 'FdBNo') then
      '   FdQty= BQty from Bag(BYy= FdBYy, BChr= FdBChr, BNo= FdBNo)
      'If (TcType= 'FWB' or 'FWR') and (SourceColName= 'wFdBaseVchNo') then
      '   FdQty= FdQty from the Fgd record equal to wFdBaseVchNo
  '*** FdGrWt on FdBYy,FdBChr,FdBNo,wFdBaseVchNo                Normal
      'If SourceName = 'GrdFgd' Then
      '   If TcType= 'FB' and (SourceColName= 'FdBYy' Or 'FdBChr' Or 'FdBNo') then
      '     FdGrWt= BGrWt from Bag(BYy= FdBYy, BChr= FdBChr, BNo= FdBNo)
      '   If (TcType= 'FWB' or 'FWR') and (SourceColName= 'wFdBaseVchNo') then
      '     FdGrWt= FdGrWt from the Fgd record equal to wFdBaseVchNo
      'If SourceName = 'GrdFgRm' Then
      '   If TcType = 'FR' And FdBNo= 0 And (GrdFgRm is in Normal or Delete Mode) Then
      '     FdGrWt= sum(FrRmWt) in Gms when the Rm is going into 'FG' loc
      '             minus sum(FrRmWt) in Gms when the Rm is giong out of 'FG' loc
  '*** FdPrdOdTc on FdBYy, FdBChr, FdBNo,wFdBaseVchNo           Normal
      'If TcType= 'FB' and (SourceColName= 'FdBYy' Or 'FdBChr' Or 'FdBNo') then
      '   FdPrdOdTc= BOdTc from Bag(BYy= FdBYy, BChr= FdBChr, BNo= FdBNo)
      'If (TcType= 'FWB' or 'FWR') and (SourceColName= 'wFdBaseVchNo') then
      '   FdPrdOdTc= FdPrdOdTc from the Fgd record equal to wFdBaseVchNo
  '*** FdPrdOdYy on FdBYy, FdBChr, FdBNo,wFdBaseVchNo           Normal
      'If TcType= 'FB' and (SourceColName= 'FdBYy' Or 'FdBChr' Or 'FdBNo') then
      '   FdPrdOdYy= BOdYy from Bag(BYy= FdBYy, BChr= FdBChr, BNo= FdBNo)
      'If (TcType= 'FWB' or 'FWR') and (SourceColName= 'wFdBaseVchNo') then
      '   FdPrdOdYy= FdPrdOdYy from the Fgd record equal to wFdBaseVchNo
  '*** FdPrdOdChr on FdBYy, FdBChr, FdBNo,wFdBaseVchNo          Normal
      'If TcType= 'FB' and (SourceColName= 'FdBYy' Or 'FdBChr' Or 'FdBNo') then
      '   FdPrdOdChr= BOdChr from Bag(BYy= FdBYy, BChr= FdBChr, BNo= FdBNo)
      'If (TcType= 'FWB' or 'FWR') and (SourceColName= 'wFdBaseVchNo') then
      '   FdPrdOdChr= FdPrdOdChr from the Fgd record equal to wFdBaseVchNo
  '*** FdPrdOdNo on FdBYy, FdBChr, FdBNo,wFdBaseVchNo           Normal
      'If TcType= 'FB' and (SourceColName= 'FdBYy' Or 'FdBChr' Or 'FdBNo') then
      '   FdPrdOdNo= BOdNo from Bag(BYy= FdBYy, BChr= FdBChr, BNo= FdBNo)
      'If (TcType= 'FWB' or 'FWR') and (SourceColName= 'wFdBaseVchNo') then
      '   FdPrdOdNo= FdPrdOdNo from the Fgd record equal to wFdBaseVchNo
  '*** FdPrdOdSr on FdBYy, FdBChr, FdBNo,wFdBaseVchNo           Normal
      'If TcType= 'FB' and (SourceColName= 'FdBYy' Or 'FdBChr' Or 'FdBNo') then
      '   FdPrdOdSr= BOdSr from Bag(BYy= FdBYy, BChr= FdBChr, BNo= FdBNo)
      'If (TcType= 'FWB' or 'FWR') and (SourceColName= 'wFdBaseVchNo') then
      '   FdPrdOdSr= FdPrdOdSr from the Fgd record equal to wFdBaseVchNo
  '*** FdPrdCmCd on FdBYy, FdBChr, FdBNo,wFdBaseVchNo           Normal
      'If TcType= 'FB' and (SourceColName= 'FdBYy' Or 'FdBChr' Or 'FdBNo') then
      '   FdPrdCmCd= Customer Code of the Order to which the Bag(FdBYy,FdBChr,FdBNo) belongs
      'If (TcType= 'FWB' or 'FWR') and (SourceColName= 'wFdBaseVchNo') then
      '   FdPrdCmCd= FdPrdCmCd from the Fgd record equal to wFdBaseVchNo
  '*** FdExpCmCd on FdPrdCmCd,wFdBaseVchNo                      Normal
      'If TcType= 'FB' and (SourceColName= 'FdPrdCmCd') then
      '   FdExpCmCd= FdPrdCmCd
      'If (TcType= 'FWB' or 'FWR') and (SourceColName= 'wFdBaseVchNo') then
      '   FdExpCmCd= FdExpCmCd from the Fgd record equal to wFdBaseVchNo
  '*** FdExpOdTc on FdPrdOdTc,wFdBaseVchNo                      Normal
      'If TcType= 'FB' and (SourceColName= 'FdPrdOdTc') then
      '   FdExpOdTc= FdPrdOdTc
      'If (TcType= 'FWB' or 'FWR') and (SourceColName= 'wFdBaseVchNo') then
      '   FdExpOdTc= FdExpOdTc from the Fgd record equal to wFdBaseVchNo
  '*** FdExpOdYy on FdPrdOdYy,wFdBaseVchNo                      Normal
      'If TcType= 'FB' and (SourceColName= 'FdPrdOdYy') then
      '   FdExpOdYy= FdPrdOdYy
      'If (TcType= 'FWB' or 'FWR') and (SourceColName= 'wFdBaseVchNo') then
      '   FdExpOdYy= FdExpOdYy from the Fgd record equal to wFdBaseVchNo
  '*** FdExpOdChr on FdPrdOdChr,wFdBaseVchNo                    Normal
      'If TcType= 'FB' and (SourceColName= 'FdPrdOdChr') then
      '   FdExpOdChr= FdPrdOdChr
      'If (TcType= 'FWB' or 'FWR') and (SourceColName= 'wFdBaseVchNo') then
      '   FdExpOdChr= FdExpOdChr from the Fgd record equal to wFdBaseVchNo
  '*** FdExpOdNo on FdPrdOdNo,wFdBaseVchNo                      Normal
      'If TcType= 'FB' and (SourceColName= 'FdPrdOdNo') then
      '   FdExpOdNo= FdPrdOdNo
      'If (TcType= 'FWB' or 'FWR') and (SourceColName= 'wFdBaseVchNo') then
      '   FdExpOdNo= FdExpOdNo from the Fgd record equal to wFdBaseVchNo
  '*** FdExpOdSr on FdPrdOdSr,wFdBaseVchNo                      Normal
      'If TcType= 'FB' and (SourceColName= 'FdPrdOdSr') then
      '   FdExpOdSr= FdPrdOdSr
      'If (TcType= 'FWB' or 'FWR') and (SourceColName= 'wFdBaseVchNo') then
      '   FdExpOdSr= FdExpOdSr from the Fgd record equal to wFdBaseVchNo
  '*** FdPlYy on wFdBaseVchNo                                   Normal
      'If TcType= 'FWB' or 'FWR' then
      '   FdPlYy= FdPlYy from the Fgd record equal to wFdBaseVchNo
  '*** FdPlChr on wFdBaseVchNo                                  Normal
      'If TcType= 'FWB' or 'FWR' then
      '   FdPlChr= FdPlChr from the Fgd record equal to wFdBaseVchNo
  '*** FdPlNo on wFdBaseVchNo                                   Normal
      'If TcType= 'FWB' or 'FWR' then
      '   FdPlNo= FdPlNo from the Fgd record equal to wFdBaseVchNo
  '*** FdInTc on wFdBaseVchNo                                   Normal
      'If TcType= 'FWB' or 'FWR' then
      '   FdInTc= FdInTc from the Fgd record equal to wFdBaseVchNo
  '*** FdInYy on wFdBaseVchNo                                   Normal
      'If TcType= 'FWB' or 'FWR' then
      '   FdInYy= FdInYy from the Fgd record equal to wFdBaseVchNo
  '*** FdInChr on wFdBaseVchNo                                  Normal
      'If TcType= 'FWB' or 'FWR' then
      '   FdInChr= FdInChr from the Fgd record equal to wFdBaseVchNo
  '*** FdInNo on wFdBaseVchNo                                   Normal
      'If TcType= 'FWB' or 'FWR' then
      '   FdInNo= FdInNo from the Fgd record equal to wFdBaseVchNo
  '*** FdInExpNo on wFdBaseVchNo                                Normal
      'If TcType= 'FWB' or 'FWR' then
      '   FdInExpNo= FdInExpNo from the Fgd record equal to wFdBaseVchNo
  '*** FdRefYy on wFdBaseVchNo                                  Normal
      'If TcType= 'FWB' or 'FWR' then
      '   FdRefYy= FdYy from the Fgd record equal to wFdBaseVchNo
  '*** FdRefKey on wFdBaseVchNo                                  Normal
      'If TcType= 'FWB' or 'FWR' then
      '   FdRefKey= FdKey from the Fgd record equal to wFdBaseVchNo
  '*** FdBYyKey on FdBYy, FdBChr, FdBNo                          Normal
      'If TcType= 'FB' or 'FWB' then
      '   FdBYyKey= BLstYyKey from Bag(BYy= FdBYy, BChr= FdBChr, BNo= FdBNo)
  
  
  Dim wBaseVchSr() As String
' **** Zubin 211 **** '
  Dim ws_FGBNo() As String
' **** Zubin 211 **** '

  With GrdFgd
    Select Case ColName
' **** Zubin 211 **** '
'''''    Case Is = UCase("FdBYy")
'''''      If .Mode = fgmnorm And ADC.Mode = xNorm Then
'''''        If .Value(RowNum, "wFGBNo") <> "" Then
'''''        ws_FGBNo = GetBagNo(.Value(RowNum, "wFGBNo"))
'''''        .Value(RowNum, "FdBYy") = ws_FGBNo(0)
'''''        End If
'''''      End If
'''''    Case Is = UCase("FdBChr")
'''''      If .Mode = fgmnorm And ADC.Mode = xNorm Then
'''''        If .Value(RowNum, "wFGBNo") <> "" Then
'''''        ws_FGBNo = GetBagNo(.Value(RowNum, "wFGBNo"))
'''''        .Value(RowNum, "FdBChr") = ws_FGBNo(1)
'''''        End If
'''''      End If
' **** Zubin 211 **** '
    Case Is = UCase("FdBNo")
' **** Zubin 211 **** '
'''''      If .Mode = fgmnorm Then .Value(RowNum, "FdBNo") = 0
      If RowNum > 0 Then If .Mode = fgmnorm Then If .Value(RowNum, "wFgBNo") = "" Then .Value(RowNum, "FdBNo") = 0
'''''      If .Mode = fgmnorm And ADC.Mode = xNorm Then
'''''        If .Value(RowNum, "wFGBNo") <> "" Then
'''''        ws_FGBNo = GetBagNo(.Value(RowNum, "wFGBNo"))
'''''        .Value(RowNum, "FdBNo") = ws_FGBNo(2)
'''''        End If
'''''      End If
' **** Zubin 211 **** '
    Case Is = UCase("wFdBaseVchNo")       '*** Recalc on FdBYy, FdBChr, FdBNo, FdRefYy, FdRefKey
      '****** Sachin 3.02 - 24-12-07 - Current Partition
      ' ***** Manali 3.8.0 - Copy Bag Option In FWB Entry - Copy mode added in Condition
      If (.Mode = fgmnorm Or .Mode = fgmCopy) And ms_TcTyp = "FWB" And (SourceColName = UCase("FdBYy") Or _
         SourceColName = UCase("FdBChr") Or SourceColName = UCase("FdBNo")) Then _
        .Value(RowNum, "wFdBaseVchNo") = _
               moCn.GetFldVal("Select FdTc+'/'+FdYy+'/'+FdChr+'/'+LTrim(Str(FdNo))+'/'+" + _
                              "LTrim(Str(FdSr)) from Fgd where FdBM= 'B' and " + _
                              "Round(FdQty, 1)> Round(FdDespQty, 1) and " + _
                              "FdCoCd= '" + ADC("FgCoCd") + "' and " + _
                              "FdBYy='" + .Value(RowNum, "FdBYy") + "' and " + _
                              "FdBChr='" + .Value(RowNum, "FdBChr") + " ' and " + _
                              "FdBNo= " + CStr(.Value(RowNum, "FdBNo")) + IIF(gs_Partition = ctCurrPrtn, " and FdPrtKey='" + ctCurrPrtn + "' ", ""))

      '****** Sachin 3.02 - 24-12-07 - Current Partition
      If .Mode = fgminit And (ms_TcTyp = "FWB" Or ms_TcTyp = "FWR") And _
         (SourceColName = UCase("FdRefYy") Or SourceColName = UCase("FdRefKey")) Then _
        .Value(RowNum, "wFdBaseVchNo") = _
               moCn.GetFldVal("Select FdTc+'/'+FdYy+'/'+FdChr+'/'+LTrim(Str(FdNo))+'/'+" + _
                              "LTrim(Str(FdSr)) from Fgd where FdBM= 'B' and " + _
                              "FdCoCd='" + ADC("FgCoCd") + "' and " + _
                              "FdYy='" + .Value(RowNum, "FdRefYy") + "' and " + _
                              "FdKey= " + CStr(.Value(RowNum, "FdRefKey")) + IIF(gs_Partition = ctCurrPrtn, " and FdPrtKey='" + ctCurrPrtn + "' ", ""))
    Case Is = UCase("FdDmCd")
      If .Mode = fgmnorm Then
      '****** Sachin 3.02 - 24-12-07 - Current Partition
        If ms_TcTyp = "FB" And (SourceColName = UCase("FdBYy") Or _
          SourceColName = UCase("FdBChr") Or SourceColName = UCase("FdBNo")) Then
          .Value(RowNum, "FdDmCd") = moCn.GetFldVal("Select BOdDmCd from Bag where " + _
                                       "BCoCd= '" + ADC("FgCoCd") + "' and " + _
                                       "BYy='" + .Value(RowNum, "FdBYy") + "' and " + _
                                       "BChr='" + .Value(RowNum, "FdBChr") + " ' and " + _
                                       "BNo= " + CStr(.Value(RowNum, "FdBNo")) + IIF(gs_Partition = ctCurrPrtn, " and BPrtKey='" + ctCurrPrtn + "' ", ""))
                                       
        ElseIf (ms_TcTyp = "FWB" Or ms_TcTyp = "FWR") And SourceColName = UCase("wFdBaseVchNo") Then
          '****** Sachin 3.02 - 24-12-07 - Current Partition
          wBaseVchSr() = Split(.Value(RowNum, "wFdBaseVchNo"), "/")
          .Value(RowNum, "FdDmCd") = moCn.GetFldVal("Select FdDmCd from Fgd where " + _
                "FdCoCd= '" + ADC("FgCoCd") + "' and " + _
                "FdTc= '" + wBaseVchSr(0) + "' and FdYy= '" + wBaseVchSr(1) + "' and " + _
                "FdChr= '" + wBaseVchSr(2) + "' and FdNo= " + wBaseVchSr(3) + " and " + _
                "FdSr= " + wBaseVchSr(4) + IIF(gs_Partition = ctCurrPrtn, " and FdPrtKey='" + ctCurrPrtn + "'", ""))
        End If
      End If
    Case Is = UCase("FdSfx")
      If .Mode = fgmnorm Then
        If ms_TcTyp = "FB" And (SourceColName = UCase("FdBYy") Or _
           SourceColName = UCase("FdBChr") Or SourceColName = UCase("FdBNo")) Then
           '****** Sachin 3.02 - 24-12-07 - Current Partition
          .Value(RowNum, "FdSfx") = moCn.GetFldVal("Select BOdSfx from Bag where " + _
                                    "BCoCd='" + ADC("FgCoCd") + "' and " + _
                                    "BYy='" + .Value(RowNum, "FdBYy") + "' and " + _
                                    "BChr='" + .Value(RowNum, "FdBChr") + " ' and " + _
                                    "BNo= " + CStr(.Value(RowNum, "FdBNo")) + IIF(gs_Partition = ctCurrPrtn, " and BPrtKey='" + ctCurrPrtn + "' ", ""))
                                    
        ElseIf (ms_TcTyp = "FWB" Or ms_TcTyp = "FWR") And SourceColName = UCase("wFdBaseVchNo") Then
          '****** Sachin 3.02 - 24-12-07 - Current Partition
          wBaseVchSr() = Split(.Value(RowNum, "wFdBaseVchNo"), "/")
          .Value(RowNum, "FdSfx") = moCn.GetFldVal("Select FdSfx from Fgd where " + _
                "FdCoCd= '" + ADC("FgCoCd") + "' and " + _
                "FdTc= '" + wBaseVchSr(0) + "' and FdYy= '" + wBaseVchSr(1) + "' and " + _
                "FdChr= '" + wBaseVchSr(2) + "' and FdNo= " + wBaseVchSr(3) + " and " + _
                "FdSr= " + wBaseVchSr(4) + IIF(gs_Partition = ctCurrPrtn, " and FdPrtKey='" + ctCurrPrtn + "' ", ""))
        End If
      End If
    Case Is = UCase("FdDmSz")
      If .Mode = fgmnorm Then
        If ms_TcTyp = "FB" And (SourceColName = UCase("FdBYy") Or _
           SourceColName = UCase("FdBChr") Or SourceColName = UCase("FdBNo")) Then
           '****** Sachin 3.02 - 24-12-07 - Current Partition
          .Value(RowNum, "FdDmSz") = moCn.GetFldVal("Select BOdDmSz from Bag where " + _
                                     "BCoCd= '" + ADC("FgCoCd") + "' and " + _
                                     "BYy='" + .Value(RowNum, "FdBYy") + "' and " + _
                                     "BChr='" + .Value(RowNum, "FdBChr") + " ' and " + _
                                     "BNo= " + CStr(.Value(RowNum, "FdBNo")) + IIF(gs_Partition = ctCurrPrtn, " and BPrtKey='" + ctCurrPrtn + "' ", ""))
                                     
        ElseIf (ms_TcTyp = "FWB" Or ms_TcTyp = "FWR") And SourceColName = UCase("wFdBaseVchNo") Then
          '****** Sachin 3.02 - 24-12-07 - Current Partition
          wBaseVchSr() = Split(.Value(RowNum, "wFdBaseVchNo"), "/")
          .Value(RowNum, "FdDmSz") = moCn.GetFldVal("Select FdDmSz from Fgd where " + _
                "FdCoCd= '" + ADC("FgCoCd") + "' and " + _
                "FdTc= '" + wBaseVchSr(0) + "' and FdYy= '" + wBaseVchSr(1) + "' and " + _
                "FdChr= '" + wBaseVchSr(2) + "' and FdNo= " + wBaseVchSr(3) + " and " + _
                "FdSr= " + wBaseVchSr(4) + IIF(gs_Partition = ctCurrPrtn, " and FdPrtKey='" + ctCurrPrtn + "' ", ""))
        End If
      End If
    Case Is = UCase("FdQty")
      If .Mode = fgmnorm Then
        If ms_TcTyp = "FB" And (SourceColName = UCase("FdBYy") Or _
           SourceColName = UCase("FdBChr") Or SourceColName = UCase("FdBNo")) Then
           '****** Sachin 3.02 - 24-12-07 - Current Partition
          .Value(RowNum, "FdQty") = moCn.GetFldVal("Select BQty from Bag where " + _
                                    "BCoCd= '" + ADC("FgCoCd") + "' and " + _
                                    "BYy='" + .Value(RowNum, "FdBYy") + "' and " + _
                                    "BChr='" + .Value(RowNum, "FdBChr") + "' and " + _
                                    "BNo= " + CStr(.Value(RowNum, "FdBNo")) + IIF(gs_Partition = ctCurrPrtn, " and BPrtKey='" + ctCurrPrtn + "' ", ""))
                                    
        ElseIf (ms_TcTyp = "FWB" Or ms_TcTyp = "FWR") And SourceColName = UCase("wFdBaseVchNo") Then
          '****** Sachin 3.02 - 24-12-07 - Current Partition
          wBaseVchSr() = Split(.Value(RowNum, "wFdBaseVchNo"), "/")
          .Value(RowNum, "FdQty") = moCn.GetFldVal("Select FdQty from Fgd where " + _
                "FdCoCd= '" + ADC("FgCoCd") + "' and " + _
                "FdTc= '" + wBaseVchSr(0) + "' and FdYy= '" + wBaseVchSr(1) + "' and " + _
                "FdChr= '" + wBaseVchSr(2) + "' and FdNo= " + wBaseVchSr(3) + " and " + _
                "FdSr= " + wBaseVchSr(4) + IIF(gs_Partition = ctCurrPrtn, " and FdPrtKey='" + ctCurrPrtn + "' ", ""))
        End If
      End If
    Case Is = UCase("FdGrWt")
      Dim wFdGrWt As Single, i As Integer, wctg As String, wPurity As Single
      Dim wLocTyp As String, wMul As Integer
      If .Mode = fgmnorm Then
        If UCase(SourceName) = UCase("GrdFgd") Then
          If ms_TcTyp = "FB" And (SourceColName = UCase("FdBYy") Or _
             SourceColName = UCase("FdBChr") Or SourceColName = UCase("FdBNo")) Then
             '****** Sachin 3.02 - 24-12-07 - Current Partition
            .Value(RowNum, "FdGrWt") = moCn.GetFldVal("Select BGrWt from Bag where " + _
                                       "BCoCd= '" + ADC("FgCoCd") + "' and " + _
                                       "BYy='" + .Value(RowNum, "FdBYy") + "' and " + _
                                       "BChr='" + .Value(RowNum, "FdBChr") + "' and " + _
                                       "BNo= " + CStr(.Value(RowNum, "FdBNo")) + IIF(gs_Partition = ctCurrPrtn, " and BPrtKey='" + ctCurrPrtn + "'", ""))
                                       
          ElseIf (ms_TcTyp = "FWB" Or ms_TcTyp = "FWR") And SourceColName = UCase("wFdBaseVchNo") Then
            '****** Sachin 3.02 - 24-12-07 - Current Partition
            wBaseVchSr() = Split(.Value(RowNum, "wFdBaseVchNo"), "/")
            .Value(RowNum, "FdGrWt") = moCn.GetFldVal("Select FdGrWt from Fgd where " + _
                  "FdCoCd= '" + ADC("FgCoCd") + "' and " + _
                  "FdTc= '" + wBaseVchSr(0) + "' and FdYy= '" + wBaseVchSr(1) + "' and " + _
                  "FdChr= '" + wBaseVchSr(2) + "' and FdNo= " + wBaseVchSr(3) + " and " + _
                  "FdSr= " + wBaseVchSr(4) + IIF(gs_Partition = ctCurrPrtn, " and FdPrtKey='" + ctCurrPrtn + "' ", ""))
          End If
          
        End If
        
        If UCase(SourceName) = UCase("GrdFgRm") Then
          If ms_TcTyp = "FR" And .Value(.Row, "FdBNo") = 0 And (GrdFgRm.Mode = fgmnorm Or GrdFgRm.Mode = fgmdel) Then
            With GrdFgRm
              wFdGrWt = 0: wMul = 1
              For i = 1 To (.Rows - 1)
                wctg = moCn.GetFldVal("select RmCtg from RmMst where RmCd='" + .Value(i, "FrRmCd") + "'" + IIF(gs_Partition = ctCurrPrtn, " And RmPrtKey='" + ctCurrPrtn + "' ", ""))
                '*** wt should not be multiplied by wtpurity
                'wpurity = moCn.GetFldVal("select RmPurityWt from RmMst where RmCd='" + .Value(i, "FrRmCd") + "'")
                wLocTyp = moCn.GetFldVal("Select LocTyp From Loc where " + _
                                          "LocCoCd= '" + ADC("FgCoCd") + "' and " + _
                                          "LocCd= '" + .Value(i, "FrFrRmLoc") + "'")
                wMul = IIF(Trim(wLocTyp) = "F", 1, -1)
                'If ms_TcTyp = "FWR" Then wMul = IIf(Trim(wLocTyp) = "F", -1, 1)
                If Not .IsDel(i) Then wFdGrWt = wFdGrWt + .Value(i, "FrRmWt") * _
                        IIF(.Value(i, "FrFrRmDc") = "C", -1, _
                        IIF(.Value(i, "FrFrRmDc") = "D", 1, 0)) / _
                        IIF(wctg = "D" Or wctg = "C", 5, 1) * wMul
              Next i
            End With
            .Value(.Row, "FdGrWt") = wFdGrWt
            GrdFgRm.AddSave "GrdFgd", asCurrent
          End If
        End If
      End If
    Case Is = UCase("FdPrdOdTc")
      If .Mode = fgmnorm Then
        If ms_TcTyp = "FB" And (SourceColName = UCase("FdBYy") Or _
           SourceColName = UCase("FdBChr") Or SourceColName = UCase("FdBNo")) Then
           '****** Sachin 3.02 - 24-12-07 - Current Partition
          .Value(RowNum, "FdPrdOdTc") = moCn.GetFldVal("Select BOdTc from Bag where " + _
                                        "BCoCd= '" + ADC("FgCoCd") + "' and " + _
                                        "BYy='" + .Value(RowNum, "FdBYy") + "' and " + _
                                        "BChr='" + .Value(RowNum, "FdBChr") + " ' and " + _
                                        "BNo= " + CStr(.Value(RowNum, "FdBNo")) + IIF(gs_Partition = ctCurrPrtn, " and BPrtKey='" + ctCurrPrtn + "' ", ""))
                                        
        ElseIf (ms_TcTyp = "FWB" Or ms_TcTyp = "FWR") And SourceColName = UCase("wFdBaseVchNo") Then
          '****** Sachin 3.02 - 24-12-07 - Current Partition
          wBaseVchSr() = Split(.Value(RowNum, "wFdBaseVchNo"), "/")
          .Value(RowNum, "FdPrdOdTc") = moCn.GetFldVal("Select FdPrdOdTc from Fgd where " + _
                "FdCoCd= '" + ADC("FgCoCd") + "' and " + _
                "FdTc= '" + wBaseVchSr(0) + "' and FdYy= '" + wBaseVchSr(1) + "' and " + _
                "FdChr= '" + wBaseVchSr(2) + "' and FdNo= " + wBaseVchSr(3) + " and " + _
                "FdSr= " + wBaseVchSr(4) + IIF(gs_Partition = ctCurrPrtn, " and FdPrtKey='" + ctCurrPrtn + "' ", ""))
        End If
      End If
    Case Is = UCase("FdPrdOdYy")
      If .Mode = fgmnorm Then
        If ms_TcTyp = "FB" And (SourceColName = UCase("FdBYy") Or _
           SourceColName = UCase("FdBChr") Or SourceColName = UCase("FdBNo")) Then
           '****** Sachin 3.02 - 24-12-07 - Current Partition
          .Value(RowNum, "FdPrdOdYy") = moCn.GetFldVal("Select BOdYy from Bag where " + _
                                        "BCoCd= '" + ADC("FgCoCd") + "' and " + _
                                        "BYy='" + .Value(RowNum, "FdBYy") + "' and " + _
                                        "BChr='" + .Value(RowNum, "FdBChr") + " ' and " + _
                                        "BNo= " + CStr(.Value(RowNum, "FdBNo")) + IIF(gs_Partition = ctCurrPrtn, " and BPrtKey='" + ctCurrPrtn + "' ", ""))
                                        
        ElseIf (ms_TcTyp = "FWB" Or ms_TcTyp = "FWR") And SourceColName = UCase("wFdBaseVchNo") Then
          '****** Sachin 3.02 - 24-12-07 - Current Partition
          wBaseVchSr() = Split(.Value(RowNum, "wFdBaseVchNo"), "/")
          .Value(RowNum, "FdPrdOdYy") = moCn.GetFldVal("Select FdPrdOdYy from Fgd where " + _
                "FdCoCd= '" + ADC("FgCoCd") + "' and " + _
                "FdTc= '" + wBaseVchSr(0) + "' and FdYy= '" + wBaseVchSr(1) + "' and " + _
                "FdChr= '" + wBaseVchSr(2) + "' and FdNo= " + wBaseVchSr(3) + " and " + _
                "FdSr= " + wBaseVchSr(4) + IIF(gs_Partition = ctCurrPrtn, " and FdPrtKey='" + ctCurrPrtn + "' ", ""))
        End If
      End If

    Case Is = UCase("FdPrdOdChr")
      If .Mode = fgmnorm Then
        '****** Sachin 3.02 - 24-12-07 - Current Partition
        If ms_TcTyp = "FB" And (SourceColName = UCase("FdBYy") Or _
           SourceColName = UCase("FdBChr") Or SourceColName = UCase("FdBNo")) Then
          .Value(RowNum, "FdPrdOdChr") = moCn.GetFldVal("Select BOdChr from Bag where " + _
                                         "BCoCd= '" + ADC("FgCoCd") + "' and " + _
                                         "BYy='" + .Value(RowNum, "FdBYy") + "' and " + _
                                         "BChr='" + .Value(RowNum, "FdBChr") + " ' and " + _
                                         "BNo= " + CStr(.Value(RowNum, "FdBNo")) + IIF(gs_Partition = ctCurrPrtn, " and BPrtKey='" + ctCurrPrtn + "' ", ""))
                                         
        '****** Sachin 3.02 - 24-12-07 - Current Partition
        ElseIf (ms_TcTyp = "FWB" Or ms_TcTyp = "FWR") And SourceColName = UCase("wFdBaseVchNo") Then
          wBaseVchSr() = Split(.Value(RowNum, "wFdBaseVchNo"), "/")
          .Value(RowNum, "FdPrdOdChr") = moCn.GetFldVal("Select FdPrdOdChr from Fgd where " + _
                "FdCoCd= '" + ADC("FgCoCd") + "' and " + _
                "FdTc= '" + wBaseVchSr(0) + "' and FdYy= '" + wBaseVchSr(1) + "' and " + _
                "FdChr= '" + wBaseVchSr(2) + "' and FdNo= " + wBaseVchSr(3) + " and " + _
                "FdSr= " + wBaseVchSr(4) + IIF(gs_Partition = ctCurrPrtn, " and FdPrtKey='" + ctCurrPrtn + "' ", ""))
        End If
      End If
    Case Is = UCase("FdPrdOdNo")
      If .Mode = fgmnorm Then
        If ms_TcTyp = "FB" And (SourceColName = UCase("FdBYy") Or _
           SourceColName = UCase("FdBChr") Or SourceColName = UCase("FdBNo")) Then
          .Value(RowNum, "FdPrdOdNo") = moCn.GetFldVal("Select BOdNo from Bag where " + _
                                        "BCoCd= '" + ADC("FgCoCd") + "' and " + _
                                        "BYy='" + .Value(RowNum, "FdBYy") + "' and " + _
                                        "BChr='" + .Value(RowNum, "FdBChr") + " ' and " + _
                                        "BNo= " + CStr(.Value(RowNum, "FdBNo")) + IIF(gs_Partition = ctCurrPrtn, " and BPrtKey='" + ctCurrPrtn + "' ", ""))
                                        
        ElseIf (ms_TcTyp = "FWB" Or ms_TcTyp = "FWR") And SourceColName = UCase("wFdBaseVchNo") Then
          wBaseVchSr() = Split(.Value(RowNum, "wFdBaseVchNo"), "/")
          .Value(RowNum, "FdPrdOdNo") = moCn.GetFldVal("Select FdPrdOdNo from Fgd where " + _
                "FdCoCd= '" + ADC("FgCoCd") + "' and " + _
                "FdTc= '" + wBaseVchSr(0) + "' and FdYy= '" + wBaseVchSr(1) + "' and " + _
                "FdChr= '" + wBaseVchSr(2) + "' and FdNo= " + wBaseVchSr(3) + " and " + _
                "FdSr= " + wBaseVchSr(4) + IIF(gs_Partition = ctCurrPrtn, " and FdPrtKey='" + ctCurrPrtn + "'", ""))
        End If
      End If
    Case Is = UCase("FdPrdOdSr")
      If .Mode = fgmnorm Then
        If ms_TcTyp = "FB" And (SourceColName = UCase("FdBYy") Or _
           SourceColName = UCase("FdBChr") Or SourceColName = UCase("FdBNo")) Then
          '****** Sachin 3.02 - 24-12-07 - Current Partition
          .Value(RowNum, "FdPrdOdSr") = moCn.GetFldVal("Select BOdSr from Bag where " + _
                                        "BCoCd= '" + ADC("FgCoCd") + "' and " + _
                                        "BYy='" + .Value(RowNum, "FdBYy") + "' and " + _
                                        "BChr='" + .Value(RowNum, "FdBChr") + " ' and " + _
                                        "BNo= " + CStr(.Value(RowNum, "FdBNo")) + IIF(gs_Partition = ctCurrPrtn, " and BPrtKey='" + ctCurrPrtn + "' ", ""))
                                        
        '****** Sachin 3.02 - 24-12-07 - Current Partition
        ElseIf (ms_TcTyp = "FWB" Or ms_TcTyp = "FWR") And SourceColName = UCase("wFdBaseVchNo") Then
          wBaseVchSr() = Split(.Value(RowNum, "wFdBaseVchNo"), "/")
          .Value(RowNum, "FdPrdOdSr") = moCn.GetFldVal("Select FdPrdOdSr from Fgd where " + _
                "FdCoCd= '" + ADC("FgCoCd") + "' and " + _
                "FdTc= '" + wBaseVchSr(0) + "' and FdYy= '" + wBaseVchSr(1) + "' and " + _
                "FdChr= '" + wBaseVchSr(2) + "' and FdNo= " + wBaseVchSr(3) + " and " + _
                "FdSr= " + wBaseVchSr(4) + IIF(gs_Partition = ctCurrPrtn, " and FdPrtKey='" + ctCurrPrtn + "' ", ""))
        End If
      End If
    Case Is = UCase("FdPrdCmCd")
      If .Mode = fgmnorm Then
        '****** Sachin 3.02 - Id fields In Joins
        If ms_TcTyp = "FB" And (SourceColName = UCase("FdBYy") Or _
           SourceColName = UCase("FdBChr") Or SourceColName = UCase("FdBNo")) Then
           '****** Sachin 3.02 - 24-12-07 - Current Partition
          .Value(RowNum, "FdPrdCmCd") = moCn.GetFldVal("Select OmCmCd from OrdMst join Bag on " + _
                                        "OmIdNo=BOmIdNo and OmCoCd=BCoCd and OmTc=BOdTc and OmYy=BOdYy and " + _
                                        "OmChr=BOdChr and OmNo=BOdNo where " + _
                                        "BCoCd= '" + ADC("FgCoCd") + "' and " + _
                                        "BYy='" + .Value(RowNum, "FdBYy") + "' and " + _
                                        "BChr='" + .Value(RowNum, "FdBChr") + " ' and " + _
                                        "BNo= " + CStr(.Value(RowNum, "FdBNo")) + IIF(gs_Partition = ctCurrPrtn, " and BPrtKey='" + ctCurrPrtn + "' ", ""))
                                        
        ElseIf (ms_TcTyp = "FWB" Or ms_TcTyp = "FWR") And SourceColName = UCase("wFdBaseVchNo") Then
          '****** Sachin 3.02 - 24-12-07 - Current Partition
          wBaseVchSr() = Split(.Value(RowNum, "wFdBaseVchNo"), "/")
          .Value(RowNum, "FdPrdCmCd") = moCn.GetFldVal("Select FdPrdCmCd from Fgd where " + _
                "FdCoCd= '" + ADC("FgCoCd") + "' and " + _
                "FdTc= '" + wBaseVchSr(0) + "' and FdYy= '" + wBaseVchSr(1) + "' and " + _
                "FdChr= '" + wBaseVchSr(2) + "' and FdNo= " + wBaseVchSr(3) + " and " + _
                "FdSr= " + wBaseVchSr(4) + IIF(gs_Partition = ctCurrPrtn, " and FdPrtKey='" + ctCurrPrtn + "' ", ""))
        End If
      End If
    Case Is = UCase("FdExpCmCd")
      If .Mode = fgmnorm Then
        '****** Sachin 3.02 - 24-12-07 - Current Partition
        If ms_TcTyp = "FB" And SourceColName = UCase("FdPrdCmCd") Then
          .Value(RowNum, "FdExpCmCd") = .Value(RowNum, "FdPrdCmCd")
        ElseIf (ms_TcTyp = "FWB" Or ms_TcTyp = "FWR") And SourceColName = UCase("wFdBaseVchNo") Then
          wBaseVchSr() = Split(.Value(RowNum, "wFdBaseVchNo"), "/")
          .Value(RowNum, "FdExpCmCd") = moCn.GetFldVal("Select FdExpCmCd from Fgd where " + _
                "FdCoCd= '" + ADC("FgCoCd") + "' and " + _
                "FdTc= '" + wBaseVchSr(0) + "' and FdYy= '" + wBaseVchSr(1) + "' and " + _
                "FdChr= '" + wBaseVchSr(2) + "' and FdNo= " + wBaseVchSr(3) + " and " + _
                "FdSr= " + wBaseVchSr(4) + IIF(gs_Partition = ctCurrPrtn, " and FdPrtKey='" + ctCurrPrtn + "' ", ""))
        End If
      End If
    Case Is = UCase("FdExpOdTc")
      If .Mode = fgmnorm Then
        If ms_TcTyp = "FB" And SourceColName = UCase("FdPrdOdTc") Then
          .Value(RowNum, "FdExpOdTc") = .Value(RowNum, "FdPrdOdTc")
        ElseIf (ms_TcTyp = "FWB" Or ms_TcTyp = "FWR") And SourceColName = UCase("wFdBaseVchNo") Then
          wBaseVchSr() = Split(.Value(RowNum, "wFdBaseVchNo"), "/")
          '****** Sachin 3.02 - 24-12-07 - Current Partition
          .Value(RowNum, "FdExpOdTc") = moCn.GetFldVal("Select FdExpOdTc from Fgd where " + _
                "FdCoCd= '" + ADC("FgCoCd") + "' and " + _
                "FdTc= '" + wBaseVchSr(0) + "' and FdYy= '" + wBaseVchSr(1) + "' and " + _
                "FdChr= '" + wBaseVchSr(2) + "' and FdNo= " + wBaseVchSr(3) + " and " + _
                "FdSr= " + wBaseVchSr(4) + IIF(gs_Partition = ctCurrPrtn, " and FdPrtKey='" + ctCurrPrtn + "' ", ""))
        End If
      End If
    Case Is = UCase("FdExpOdYy")
      If .Mode = fgmnorm Then
        If ms_TcTyp = "FB" And SourceColName = UCase("FdPrdOdYy") Then
          .Value(RowNum, "FdExpOdYy") = .Value(RowNum, "FdPrdOdYy")
        ElseIf (ms_TcTyp = "FWB" Or ms_TcTyp = "FWR") And SourceColName = UCase("wFdBaseVchNo") Then
          wBaseVchSr() = Split(.Value(RowNum, "wFdBaseVchNo"), "/")
          '****** Sachin 3.02 - 24-12-07 - Current Partition
          .Value(RowNum, "FdExpOdYy") = moCn.GetFldVal("Select FdExpOdYy from Fgd where " + _
                "FdCoCd= '" + ADC("FgCoCd") + "' and " + _
                "FdTc= '" + wBaseVchSr(0) + "' and FdYy= '" + wBaseVchSr(1) + "' and " + _
                "FdChr= '" + wBaseVchSr(2) + "' and FdNo= " + wBaseVchSr(3) + " and " + _
                "FdSr= " + wBaseVchSr(4) + IIF(gs_Partition = ctCurrPrtn, " and FdPrtKey='" + ctCurrPrtn + "' ", ""))
        End If
      End If

    Case Is = UCase("FdExpOdChr")
      If .Mode = fgmnorm Then
        If ms_TcTyp = "FB" And SourceColName = UCase("FdPrdOdChr") Then
          .Value(RowNum, "FdExpOdChr") = .Value(RowNum, "FdPrdOdChr")
        ElseIf (ms_TcTyp = "FWB" Or ms_TcTyp = "FWR") And SourceColName = UCase("wFdBaseVchNo") Then
          wBaseVchSr() = Split(.Value(RowNum, "wFdBaseVchNo"), "/")
          '****** Sachin 3.02 - 24-12-07 - Current Partition
          .Value(RowNum, "FdExpOdChr") = moCn.GetFldVal("Select FdExpOdChr from Fgd where " + _
                "FdCoCd= '" + ADC("FgCoCd") + "' and " + _
                "FdTc= '" + wBaseVchSr(0) + "' and FdYy= '" + wBaseVchSr(1) + "' and " + _
                "FdChr= '" + wBaseVchSr(2) + "' and FdNo= " + wBaseVchSr(3) + " and " + _
                "FdSr= " + wBaseVchSr(4) + IIF(gs_Partition = ctCurrPrtn, " and FdPrtKey='" + ctCurrPrtn + "' ", ""))
        End If
      End If
    Case Is = UCase("FdExpOdNo")
      If .Mode = fgmnorm Then
        If ms_TcTyp = "FB" And SourceColName = UCase("FdPrdOdNo") Then
          .Value(RowNum, "FdExpOdNo") = .Value(RowNum, "FdPrdOdNo")
        ElseIf (ms_TcTyp = "FWB" Or ms_TcTyp = "FWR") And SourceColName = UCase("wFdBaseVchNo") Then
          wBaseVchSr() = Split(.Value(RowNum, "wFdBaseVchNo"), "/")
          '****** Sachin 3.02 - 24-12-07 - Current Partition
          .Value(RowNum, "FdExpOdNo") = moCn.GetFldVal("Select FdExpOdNo from Fgd where " + _
                "FdCoCd= '" + ADC("FgCoCd") + "' and " + _
                "FdTc= '" + wBaseVchSr(0) + "' and FdYy= '" + wBaseVchSr(1) + "' and " + _
                "FdChr= '" + wBaseVchSr(2) + "' and FdNo= " + wBaseVchSr(3) + " and " + _
                "FdSr= " + wBaseVchSr(4) + IIF(gs_Partition = ctCurrPrtn, " and FdPrtKey='" + ctCurrPrtn + "' ", ""))
        End If
      End If
    Case Is = UCase("FdExpOdSr")
      If .Mode = fgmnorm Then
        If ms_TcTyp = "FB" And SourceColName = UCase("FdPrdOdSr") Then
          .Value(RowNum, "FdExpOdSr") = .Value(RowNum, "FdPrdOdSr")
        ElseIf (ms_TcTyp = "FWB" Or ms_TcTyp = "FWR") And SourceColName = UCase("wFdBaseVchNo") Then
          wBaseVchSr() = Split(.Value(RowNum, "wFdBaseVchNo"), "/")
          '****** Sachin 3.02 - 24-12-07 - Current Partition
          .Value(RowNum, "FdExpOdSr") = moCn.GetFldVal("Select FdExpOdSr from Fgd where " + _
                "FdCoCd= '" + ADC("FgCoCd") + "' and " + _
                "FdTc= '" + wBaseVchSr(0) + "' and FdYy= '" + wBaseVchSr(1) + "' and " + _
                "FdChr= '" + wBaseVchSr(2) + "' and FdNo= " + wBaseVchSr(3) + " and " + _
                "FdSr= " + wBaseVchSr(4) + IIF(gs_Partition = ctCurrPrtn, " and FdPrtKey='" + ctCurrPrtn + "' ", ""))
        End If
      End If
    Case Is = UCase("FdPlYy")
    '578-24 fgmcopy added
      If (.Mode = fgmnorm Or .Mode = fgmCopy) And (ms_TcTyp = "FWB" Or ms_TcTyp = "FWR") Then
        wBaseVchSr() = Split(.Value(RowNum, "wFdBaseVchNo"), "/")
        '****** Sachin 3.02 - 24-12-07 - Current Partition
        .Value(RowNum, "FdPlYy") = moCn.GetFldVal("Select FdPlYy from Fgd where " + _
              "FdCoCd= '" + ADC("FgCoCd") + "' and " + _
              "FdTc= '" + wBaseVchSr(0) + "' and FdYy= '" + wBaseVchSr(1) + "' and " + _
              "FdChr= '" + wBaseVchSr(2) + "' and FdNo= " + wBaseVchSr(3) + " and " + _
              "FdSr= " + wBaseVchSr(4) + IIF(gs_Partition = ctCurrPrtn, " and FdPrtKey='" + ctCurrPrtn + "' ", ""))
      End If
    Case Is = UCase("FdPlChr")
    '578-24 fgmcopy added
      If (.Mode = fgmnorm Or .Mode = fgmCopy) And (ms_TcTyp = "FWB" Or ms_TcTyp = "FWR") Then
        wBaseVchSr() = Split(.Value(RowNum, "wFdBaseVchNo"), "/")
        '****** Sachin 3.02 - 24-12-07 - Current Partition
        .Value(RowNum, "FdPlChr") = moCn.GetFldVal("Select FdPlChr from Fgd where " + _
              "FdCoCd= '" + ADC("FgCoCd") + "' and " + _
              "FdTc= '" + wBaseVchSr(0) + "' and FdYy= '" + wBaseVchSr(1) + "' and " + _
              "FdChr= '" + wBaseVchSr(2) + "' and FdNo= " + wBaseVchSr(3) + " and " + _
              "FdSr= " + wBaseVchSr(4) + IIF(gs_Partition = ctCurrPrtn, " and FdPrtKey='" + ctCurrPrtn + "' ", ""))
      End If
    Case Is = UCase("FdPlNo")
    '578-24 fgmcopy added
      If (.Mode = fgmnorm Or .Mode = fgmCopy) And (ms_TcTyp = "FWB" Or ms_TcTyp = "FWR") Then
        wBaseVchSr() = Split(.Value(RowNum, "wFdBaseVchNo"), "/")
        '****** Sachin 3.02 - 24-12-07 - Current Partition
        .Value(RowNum, "FdPlNo") = moCn.GetFldVal("Select FdPlNo from Fgd where " + _
              "FdCoCd= '" + ADC("FgCoCd") + "' and " + _
              "FdTc= '" + wBaseVchSr(0) + "' and FdYy= '" + wBaseVchSr(1) + "' and " + _
              "FdChr= '" + wBaseVchSr(2) + "' and FdNo= " + wBaseVchSr(3) + " and " + _
              "FdSr= " + wBaseVchSr(4) + IIF(gs_Partition = ctCurrPrtn, " and FdPrtKey='" + ctCurrPrtn + "' ", ""))
      End If
    Case Is = UCase("FdInTc")
      If .Mode = fgmnorm And (ms_TcTyp = "FWB" Or ms_TcTyp = "FWR") Then
        wBaseVchSr() = Split(.Value(RowNum, "wFdBaseVchNo"), "/")
        '****** Sachin 3.02 - 24-12-07 - Current Partition
        .Value(RowNum, "FdInTc") = moCn.GetFldVal("Select FdInTc from Fgd where " + _
              "FdCoCd= '" + ADC("FgCoCd") + "' and " + _
              "FdTc= '" + wBaseVchSr(0) + "' and FdYy= '" + wBaseVchSr(1) + "' and " + _
              "FdChr= '" + wBaseVchSr(2) + "' and FdNo= " + wBaseVchSr(3) + " and " + _
              "FdSr= " + wBaseVchSr(4) + IIF(gs_Partition = ctCurrPrtn, " and FdPrtKey='" + ctCurrPrtn + "' ", ""))
      End If
    Case Is = UCase("FdInYy")
      If .Mode = fgmnorm And (ms_TcTyp = "FWB" Or ms_TcTyp = "FWR") Then
        wBaseVchSr() = Split(.Value(RowNum, "wFdBaseVchNo"), "/")
        '****** Sachin 3.02 - 24-12-07 - Current Partition
        .Value(RowNum, "FdInYy") = moCn.GetFldVal("Select FdInYy from Fgd where " + _
              "FdCoCd= '" + ADC("FgCoCd") + "' and " + _
              "FdTc= '" + wBaseVchSr(0) + "' and FdYy= '" + wBaseVchSr(1) + "' and " + _
              "FdChr= '" + wBaseVchSr(2) + "' and FdNo= " + wBaseVchSr(3) + " and " + _
              "FdSr= " + wBaseVchSr(4) + IIF(gs_Partition = ctCurrPrtn, " and FdPrtKey='" + ctCurrPrtn + "' ", ""))
      End If

    Case Is = UCase("FdInChr")
      If .Mode = fgmnorm And (ms_TcTyp = "FWB" Or ms_TcTyp = "FWR") Then
        wBaseVchSr() = Split(.Value(RowNum, "wFdBaseVchNo"), "/")
        '****** Sachin 3.02 - 24-12-07 - Current Partition
        .Value(RowNum, "FdInChr") = moCn.GetFldVal("Select FdInChr from Fgd where " + _
              "FdCoCd= '" + ADC("FgCoCd") + "' and " + _
              "FdTc= '" + wBaseVchSr(0) + "' and FdYy= '" + wBaseVchSr(1) + "' and " + _
              "FdChr= '" + wBaseVchSr(2) + "' and FdNo= " + wBaseVchSr(3) + " and " + _
              "FdSr= " + wBaseVchSr(4) + IIF(gs_Partition = ctCurrPrtn, " and FdPrtKey='" + ctCurrPrtn + "' ", ""))
      End If
    Case Is = UCase("FdInNo")
      If .Mode = fgmnorm And (ms_TcTyp = "FWB" Or ms_TcTyp = "FWR") Then
        wBaseVchSr() = Split(.Value(RowNum, "wFdBaseVchNo"), "/")
        '****** Sachin 3.02 - 24-12-07 - Current Partition
        .Value(RowNum, "FdInNo") = moCn.GetFldVal("Select FdInNo from Fgd where " + _
              "FdCoCd= '" + ADC("FgCoCd") + "' and " + _
              "FdTc= '" + wBaseVchSr(0) + "' and FdYy= '" + wBaseVchSr(1) + "' and " + _
              "FdChr= '" + wBaseVchSr(2) + "' and FdNo= " + wBaseVchSr(3) + " and " + _
              "FdSr= " + wBaseVchSr(4) + IIF(gs_Partition = ctCurrPrtn, " and FdPrtKey='" + ctCurrPrtn + "' ", ""))
      End If
    Case Is = UCase("FdInExpNo")
      If .Mode = fgmnorm And (ms_TcTyp = "FWB" Or ms_TcTyp = "FWR") Then
        wBaseVchSr() = Split(.Value(RowNum, "wFdBaseVchNo"), "/")
        '****** Sachin 3.02 - 24-12-07 - Current Partition
        .Value(RowNum, "FdInExpNo") = moCn.GetFldVal("Select FdInExpNo from Fgd where " + _
              "FdCoCd= '" + ADC("FgCoCd") + "' and " + _
              "FdTc= '" + wBaseVchSr(0) + "' and FdYy= '" + wBaseVchSr(1) + "' and " + _
              "FdChr= '" + wBaseVchSr(2) + "' and FdNo= " + wBaseVchSr(3) + " and " + _
              "FdSr= " + wBaseVchSr(4) + IIF(gs_Partition = ctCurrPrtn, " and FdPrtKey='" + ctCurrPrtn + "' ", ""))
      End If
    Case Is = UCase("FdRefYy")
      ' ***** Manali 3.8.0 - Copy Bag Option In FWB Entry - Copy mode added in Condition
      If (.Mode = fgmnorm Or .Mode = fgmCopy) And (ms_TcTyp = "FWB" Or ms_TcTyp = "FWR") Then
        wBaseVchSr() = Split(.Value(RowNum, "wFdBaseVchNo"), "/")
        '****** Sachin 3.02 - 24-12-07 - Current Partition
        .Value(RowNum, "FdRefYy") = moCn.GetFldVal("Select FdYy from Fgd where " + _
              "FdCoCd= '" + ADC("FgCoCd") + "' and " + _
              "FdTc= '" + wBaseVchSr(0) + "' and FdYy= '" + wBaseVchSr(1) + "' and " + _
              "FdChr= '" + wBaseVchSr(2) + "' and FdNo= " + wBaseVchSr(3) + " and " + _
              "FdSr= " + wBaseVchSr(4) + IIF(gs_Partition = ctCurrPrtn, " and FdPrtKey='" + ctCurrPrtn + "' ", ""))
      End If
    Case Is = UCase("FdRefKey")
      ' ***** Manali 3.8.0 - Copy Bag Option In FWB Entry - Copy mode added in Condition
      If (.Mode = fgmnorm Or .Mode = fgmCopy) And (ms_TcTyp = "FWB" Or ms_TcTyp = "FWR") Then
        wBaseVchSr() = Split(.Value(RowNum, "wFdBaseVchNo"), "/")
        '****** Sachin 3.02 - 24-12-07 - Current Partition
        .Value(RowNum, "FdRefKey") = moCn.GetFldVal("Select FdKey from Fgd where " + _
              "FdCoCd= '" + ADC("FgCoCd") + "' and " + _
              "FdTc= '" + wBaseVchSr(0) + "' and FdYy= '" + wBaseVchSr(1) + "' and " + _
              "FdChr= '" + wBaseVchSr(2) + "' and FdNo= " + wBaseVchSr(3) + " and " + _
              "FdSr= " + wBaseVchSr(4) + IIF(gs_Partition = ctCurrPrtn, " and FdPrtKey='" + ctCurrPrtn + "' ", ""))
      End If
    Case Is = UCase("FdBYyKey")
      If .Mode = fgmnorm And (ms_TcTyp = "FB" Or ms_TcTyp = "FWB") Then
        '****** Sachin 3.02 - 24-12-07 - Current Partition
        .Value(RowNum, "FdBYyKey") = moCn.GetFldVal("Select BLstYyKey from Bag where " + _
                                     "BCoCd= '" + ADC("FgCoCd") + "' and " + _
                                     "BYy='" + .Value(RowNum, "FdBYy") + "' and " + _
                                     "BChr='" + .Value(RowNum, "FdBChr") + " ' and " + _
                                     "BNo= " + CStr(.Value(RowNum, "FdBNo")) + IIF(gs_Partition = ctCurrPrtn, " and BPrtKey='" + ctCurrPrtn + "' ", ""))
      End If
    Case Is = UCase("FgDesc")
    If .Mode = fgmnorm Then
        If .Row > 0 Then
          If GrdFgd.Value(.Row, "FgDesc") <> ADC("wFgDesc") Then
            GrdFgd.Value(.Row, "FgDesc") = ADC("wFgDesc")
            GrdFgd.SaveRec .Row
          End If
        End If
      End If
      
    '****** Sachin 3.02 - Id fields in Joins
    '****** Sachin 3.02 - Shifted to InitKey
    '    Case Is = UCase("FdBIdNo")
    '    If .Mode = fgmnorm Or .Mode = fgmCopy Then _
    '        .Value(RowNum, "FdBIdNo") = moCn.GetFldVal("Select BIdNo From Bag Where " + _
    '                                                    " BCoCd='" + adc("FgCoCd") + "' And BYy='" + .Value(RowNum, "FdBYy") + "' And BChr = '" + .Value(RowNum, "FdBChr") + "' And BNo=" + CStr(.Value(RowNum, "FdBNo")) + " and BPrtKey='" + ctCurrPrtn + "' ")
    '
    '    Case Is = UCase("FdDmIdNo")
    '    If .Mode = fgmnorm Or .Mode = fgmCopy Then _
    '        .Value(RowNum, "FdDmIdNo") = moCn.GetFldVal("Select BDmIdNo From Bag Where " + _
    '                                                    " BCoCd='" + adc("FgCoCd") + "' And BYy='" + .Value(RowNum, "FdBYy") + "' And BChr = '" + .Value(RowNum, "FdBChr") + "' And BNo=" + CStr(.Value(RowNum, "FdBNo")) + " And BPrtKey='" + ctCurrPrtn + "' ")
                                                    
                                                        
'    Case Is = UCase("FdPrdOdIdNo")
'    If .Mode = fgmnorm Or .Mode = fgmCopy Then _
'        .Value(RowNum, "FdPrdOdIdNo") = moCn.GetFldVal("Select OdIdNo From OrdDsg Where " + _
'                                                    " OdCoCd='" + adc("FgCoCd") + "' And OdTc='" + .Value(RowNum, "FdPrdOdTc") + "' And OdYy='" + .Value(RowNum, "FdPrdOdYy") + "' And OdChr = '" + .Value(RowNum, "FdPrdOdChr") + "' And OdNo=" + CStr(.Value(RowNum, "FdPrdOdNo")) + " And OdSr=" + CStr(.Value(RowNum, "FdPrdOdSr")))
'
'    Case Is = UCase("FdExpOdIdNo")
'    If .Mode = fgmnorm Or .Mode = fgmCopy Then _
'        .Value(RowNum, "FdExpOdIdNo") = moCn.GetFldVal("Select OdIdNo From OrdDsg Where " + _
'                                                    " OdCoCd='" + adc("FgCoCd") + "' And OdTc='" + .Value(RowNum, "FdExpOdTc") + "' And OdYy='" + .Value(RowNum, "FdExpOdYy") + "' And OdChr = '" + .Value(RowNum, "FdExpOdChr") + "' And OdNo=" + CStr(.Value(RowNum, "FdExpOdNo")) + " And OdSr=" + CStr(.Value(RowNum, "FdExpOdSr")))
                                                    
    '****** Sachin 3.02 - Id fields in Joins
      
    End Select
  End With
End Sub
Private Sub GrdFgd_ColChange(ByVal RowNum As Integer, ByVal ColName As String, ByVal NewValue As Variant, Cancel As Boolean, ErrMsg As String)
  '*** FdBYy            valid PMCd('YY') or ''
  '*** FdBChr           valid PMCd('BCHR') or ''
  '*** FdBNo            If 'FB' then
  '                       FdBNo= valid BNo from Bag(BYy= FdBYy, BChr= FdBChr, BCls= 'N') or 0
  '                     If 'FWB' then
  '                       FdBNo= valid BNo from Bag(BYy= FdBYy, BChr= FdBChr, BCls= 'N', BLoc= 'PFG') or 0
  '*** wFdBaseVchNo     valid Voucher Sr from Fgd(FdBm= 'M' and Entry Type = Entry Type of the current Entry)
  '*** FdDmCd           valid DmCd from DsgMst(DmTcTyp= 'DM', DmSz= '')
  '*** FdDmSz           valid DmSz from DsgMst(DmTcTyp= 'DM', DmCd= FdDmCd)
  '*** FdQty            >= 0 and >= FdDespQty
  '*** FdPrdOdTc        Blank or valid PMCd('TC') of TcType= 'SO' and EntryType= 'OM'
  '*** FdPrdOdYy        Blank or valid PMCd('YY')
  '*** FdPrdOdChr       Blank or valid PSCd('CHR', FdPrdOdTc)
  '*** FdPrdOdNo        Zero or valid OdNo from OrdDsg(OdTc= FdPrdOdTc, OdYy= FdPrdOdYy,
  '                     OdChr= FdPrdOdChr)
  '*** FdPrdOdSr        Zero or valid OdSr from OrdDsg(OdTc= FdPrdOdTc, OdYy= FdPrdOdYy,
  '                     OdChr= FdPrdOdChr, OdNo= FdPrdOdNo)
  '*** FdPrdCmCd        Blank or valid CmCd from Custst(CmCtg= 'C')
  
' **** Zubin 211 **** '
  Dim ws_FGBNo() As String, ws_BagNoBarCd As String
  Dim ws_OldYy As String, ws_OldChr As String, ws_OldNo As Single
' **** Zubin 211 **** '
  
  With GrdFgd
    Select Case ColName
    Case Is = UCase("FdBYy")
      Cancel = (NewValue <> "") And (Not moCn.RecSeek("Select PMCd From Param Where PTyp= 'YY' And " + _
                                     "PMCd = '" + NewValue + "'"))
      ErrMsg = "Invalid Bag Year": Exit Sub
    Case Is = UCase("FdBChr")
'*** Jay 2.13(CT) ***
' **** Zubin 211 **** '
'''''      Cancel = (NewValue <> "") And (Not moCn.RecSeek("Select PMCd From Param Where PTyp= 'BCHR' And " + _
'''''                                     "PMCd = '" + NewValue + "'"))
'''''      ErrMsg = "Invalid Bag Character": Exit Sub
      'uma *** added D Bags
      Cancel = (NewValue <> "") And ((Not moCn.RecSeek("Select 'x' From Param Where PTyp= 'BCHR' And " + _
                                          "PMCd = '" + NewValue + "' and PSCd= ''")) Or _
                                     (moCn.RecSeek("Select 'x' From Param Where PTyp= 'BCHR' And " + _
                                          "PMCd = '" + NewValue + "' and PSCd= '' and PValue<> 'P' and PValue<> 'D'")))
                                        
      ErrMsg = "Invalid Bag Character / Flute or Component Bag Cannot Be Taken To FG": Exit Sub
' **** Zubin 211 **** '
'*** Jay 2.13(CT) ***
    Case Is = UCase("FdBNo")
      If ms_TcTyp = "FB" Then
        '****** Sachin 3.02 - 24-12-07 - Current Partition
         Cancel = (NewValue <> 0) And (Not moCn.RecSeek("Select BNo From Bag Where " + _
                                      "BCoCd= '" + ADC("FgCoCd") + "' and " + _
                                      "BYy= '" + .Value(RowNum, "FdBYy") + "' And " + _
                                      "BChr = '" + .Value(RowNum, "FdBChr") + "' And " + _
                                      "BNo= " + CStr(NewValue) + " and BPrtKey='" + ctCurrPrtn + "' and BCls='N'"))
        If Cancel = True Then: ErrMsg = "Invalid Bag Number Or Bag Closed": Exit Sub
        
        
        '*****4.1.2
        If mb_EnfFrBLocInFg Then
            Cancel = (NewValue <> 0) And (Not moCn.RecSeek("Select BLoc From Bag Where " + _
                                          "BCoCd= '" + ADC("FgCoCd") + "' and " + _
                                          "BYy= '" + .Value(RowNum, "FdBYy") + "' And " + _
                                          "BChr = '" + .Value(RowNum, "FdBChr") + "' And " + _
                                          "BNo= " + CStr(NewValue) + " and BPrtKey='" + ctCurrPrtn + _
                                          "' and  CHARINDEX(','+BLoc+',', ','+'" + ADC("FgFrBLoc") + "'+',') > 0"))
                                                            
  '                                        BLoc='" + adc("FgFrBLoc") + "'"))
            If Cancel = True Then: ErrMsg = " Bags Cannot be moved From Location Other Than '" + ADC("FgFrBLoc") + "'": Exit Sub
        End If
        
      ElseIf ms_TcTyp = "FWB" Then
        '****** Sachin 3.02 - 24-12-07 - Current Partition
        Cancel = (NewValue <> 0) And (Not moCn.RecSeek("Select BNo From Bag Where " + _
                                      "BCoCd= '" + ADC("FgCoCd") + "' and " + _
                                      "BYy= '" + .Value(RowNum, "FdBYy") + "' And " + _
                                      "BChr = '" + .Value(RowNum, "FdBChr") + "' And " + _
                                      "BNo= " + CStr(NewValue) + " and BPrtKey='" + ctCurrPrtn + "' and BCls='N' and " + _
                                      "BLoc= '" + ctFgLoc + "'"))
        If Cancel = True Then: ErrMsg = "Invalid Bag Number Or Bag Closed Or Bag Not In Fg Loc": Exit Sub
        
        
     If NewValue <> 0 Then
        Dim wBSubLoc As String
          wBSubLoc = moCn.GetFldVal("Select BFgSubLoc from Bag  where " + _
                                     "BCoCd='" + ADC("FgCoCd") + "' and " + _
                                     "BYy='" + GrdFgd.Value(RowNum, "FdBYy") + "' and " + _
                                     "BChr='" + GrdFgd.Value(RowNum, "FdBChr") + "' and " + _
                                     "BNo= " + CStr(NewValue) + " and BPrtKey='" + ctCurrPrtn + "' ")
          Cancel = Not moCn.RecSeek("select 'x' From vParam where vPTyp='FgSubLoc' and vPCoCd= '" + gs_CoCd + "' and vPMCd='" + wBSubLoc + "'" + _
                              " and (vPValue4='' Or ','+vPValue4+',' Like '%," + gs_UsrCd + ",%')")
        If Cancel = True Then ErrMsg = "Invalid Bag Number As Bag Sub Loc. is Not Accessible to the User ": Exit Sub
       End If
      
    End If
     '*******************Geeta************Emr206*******06/04/04
      Dim wDmCd As String
      '****** Sachin 3.02 - 24-12-07 - Current Partition
      wDmCd = moCn.GetFldVal("Select max(OdDmCd)  From  Bag join OrdDsg " + _
                                     "On OdIdNo=BOdIdNo and OdCoCd=BCocd and OdTc=BOdTc and OdYy=BOdYy and OdChr=BOdChr and OdNo=BOdNo and   OdSr=BOdSr and OdPrtKey=BPrtKey " + _
                                     " where OdHld='Y' and BCoCd= '" + ADC("FgCoCd") + "' and BYy= '" + GrdFgd.Value(RowNum, "FdBYy") + "'" + _
                                     "and BChr= '" + GrdFgd.Value(RowNum, "FdBChr") + "' and BNo= " + CStr(NewValue) + " and BPrtKey='" + ctCurrPrtn + "' ")
      If wDmCd <> "" Then DispMsg "Sr : " + CStr(.Value(RowNum, "FdSr")) + " ,Order Design on Hold", etWarning
    '***************Geeta***********

        Dim wBFrLoc As String
        If gs_ValidPrdLocSeq <> "" And NewValue <> 0 Then
          wBFrLoc = moCn.GetFldVal("Select BLoc from Bag where " + _
                                     "BCoCd='" + ADC("FgCoCd") + "' and " + _
                                     "BYy='" + GrdFgd.Value(RowNum, "FdBYy") + "' and " + _
                                     "BChr='" + GrdFgd.Value(RowNum, "FdBChr") + "' and " + _
                                     "BNo= " + CStr(NewValue) + " and BPrtKey='" + ctCurrPrtn + "' ")
          Cancel = IIF((InStr(1, "," + UCase(gs_ValidPrdLocSeq) + ",", "," + UCase(wBFrLoc) + ",") <= 0), True, False)
        Else
          Cancel = False
        End If
        If Cancel = True Then ErrMsg = "Invalid Bag Number As Bag Loc. is Not Accessible to the User ": Exit Sub
    
    ' **** Zubin 211 **** '
    Case Is = UCase("wFGBNo")
      If NewValue <> "" Then
        'uni.156 calling function to ignore CoCd if entered
        ws_BagNoBarCd = ParseBagNoBarCd(NewValue)
        ws_FGBNo = Split(ws_BagNoBarCd, "/")
        If UBound(ws_FGBNo) <> 2 Then Cancel = True: ErrMsg = "Invalid FG Bag No": Exit Sub
        If Not IsNumeric(ws_FGBNo(2)) Then Cancel = True: ErrMsg = "Invalid FG Bag No": Exit Sub
        
        ws_OldYy = .Value(RowNum, "FDBYY")
        ws_OldChr = .Value(RowNum, "FDBCHR")
        ws_OldNo = .Value(RowNum, "FDBNO")
        
        .Value(RowNum, "FDBYY") = UCase(ws_FGBNo(0))
        Call GrdFgd_ColChange(RowNum, "FDBYY", ws_FGBNo(0), Cancel, "")
        If Cancel = True Then ErrMsg = "Invalid Bag Year": .Value(RowNum, "FDBYY") = ws_OldYy: Exit Sub
        
        .Value(RowNum, "FDBCHR") = UCase(ws_FGBNo(1))
        Call GrdFgd_ColChange(RowNum, "FDBCHR", ws_FGBNo(1), Cancel, "")
        If Cancel = True Then ErrMsg = "Invalid Bag Character / Flute or Component Bags Cannot Be Taken To FG": .Value(RowNum, "FDBCHR") = ws_OldChr: Exit Sub
        
        .Value(RowNum, "FDBNO") = ws_FGBNo(2)
        Call GrdFgd_ColChange(RowNum, "FDBNO", ws_FGBNo(2), Cancel, "")
        If Cancel = True Then ErrMsg = "Invalid Bag Number": .Value(RowNum, "FDBNO") = ws_OldNo: Exit Sub

        ' ***** Manali 3.10.0 - 09/05/12 - Bar Code (Ctrl+N)
        If mb_AddRow Then
          If .Rows - 1 = RowNum Then
            .AddItem
            If .SaveRec(RowNum) Then
               Call CpyRm
              .Row = RowNum + 1
              '5713 .Col = 2 is replaced by this if condition to place the cursor at proper position
              .Col = IIF(ms_ChkFgWghtYN = "Y" And ms_TcTyp <> "FWB", 2, 3)
            Else
              .Row = RowNum
              .Col = IIF(ms_ChkFgWghtYN = "Y" And ms_TcTyp <> "FWB", 2, 3)  '5713 .Col = 2 is replaced by this if condition to place the cursor at proper position
            End If
          End If
        Else
          mb_AddRow = True    ' ***** Manali 3.10.0 - 09/05/12 - Bar Code (Ctrl+N)
        End If

''        If .Rows - 1 = RowNum Then
''          If .SaveRec(RowNum) Then
''             Call CpyRm
''            .AddItem
''            .Row = RowNum + 1
''            .Col = 2
''          Else
''            If .Row = 0 Then
''              .AddItem
''              .Row = RowNum
''            Else
''             .Row = RowNum - 1
''            End If
''           .Col = 2
''          End If
''      End If

'''''        If UCase(ADC.MenuCd) = "FB" Then
'''''          .COL = 38
'''''        ElseIf UCase(ADC.MenuCd) = "FWB" Then
'''''          .COL = 3
'''''        End If
      End If
' **** Zubin 211 **** '
    Case Is = UCase("wFdBaseVchNo")
      Dim wBaseVchSr() As String
      wBaseVchSr() = Split(NewValue, "/")
      If UBound(wBaseVchSr) <> 4 Then Cancel = True: ErrMsg = "Invalid Base Voucher Number": Exit Sub
      If Val(wBaseVchSr(3)) = 0 Or Len(Trim(wBaseVchSr(3))) <> Len(CStr(Val(wBaseVchSr(3)))) Then Cancel = True: ErrMsg = "Invalid Base Voucher Number": Exit Sub
      If Val(wBaseVchSr(4)) = 0 Or Len(Trim(wBaseVchSr(4))) <> Len(CStr(Val(wBaseVchSr(4)))) Then Cancel = True: ErrMsg = "Invalid Base Voucher Number": Exit Sub
      
      '****** Sachin 3.02 - 24-12-07 - Current Partition
      Cancel = Not moCn.RecSeek("Select FdTc from Fgd join Param on PTyp= 'TCTYP' and " + _
                   "PMCd= (Select PValue from Param where PTyp= 'TC' and PMCd= FdTc) where " + _
                   "FdCoCd= '" + ADC("FgCoCd") + "' and " + _
                   "FdTc= '" + wBaseVchSr(0) + "' and FdYy= '" + wBaseVchSr(1) + "' and " + _
                   "FdChr= '" + wBaseVchSr(2) + "' and FdNo= " + wBaseVchSr(3) + " and " + _
                   "FdSr= " + wBaseVchSr(4) + " and FdPrtKey='" + ctCurrPrtn + "' and FdBM= 'B' and PValue1= '" + ms_EntTyp + "'")
      If Cancel = True Then ErrMsg = "Invalid Base Voucher Number": Exit Sub
    Case Is = UCase("FdDmCd")
      '****** Sachin 3.02 - 24-12-07 - Current Partition
      'indigo.164 DmTcTyp= 'DM' removed
      Cancel = Not moCn.RecSeek("Select DmCd From DsgMst Where " + _
                           "DmCd='" + NewValue + "' And " + _
                           "DmSz = '' and DmPrtKey='" + ctCurrPrtn + "'")
      ErrMsg = "Invalid Design Code": Exit Sub
    
    Case Is = UCase("FdDmSz")
      '****** Sachin 3.02 - 24-12-07 - Current Partition
      'indigo.164 DmTcTyp= 'DM' removed
      Cancel = (NewValue <> "") And (Not moCn.RecSeek("Select DmSz From DsgMst Where " + _
                                              "DmCd='" + _
                                              .Value(RowNum, "FdDmCd") + "' And " + _
                                               "DmSz = '" + NewValue + "' And DmPrtKey='" + ctCurrPrtn + "'"))
      ErrMsg = "Invalid Design Size Code": Exit Sub
    Case Is = UCase("FdQty")
      Cancel = NewValue < 0 Or NewValue < .Value(RowNum, "FdDespQty")
      ErrMsg = "Invalid Finished Goods Quantity": Exit Sub
    Case Is = UCase("FdPrdOdTc")
      'indigo.164 OS tc added
      Cancel = (NewValue <> "") And (moCn.RecSeek("Select PValue From Param Where PTyp= 'Tc' " + _
                                    "And PMCd = '" + NewValue + "' And PValue<>'SO' And PValue<>'OS'") _
                                     Or _
                                    (moCn.GetFldVal("select PValue1 from Param where PTyp= 'TcTyp' and " + _
                                    "PMCd = (select PValue from Param where PTyp='Tc' and " + _
                                    "PMCd='" + NewValue + "')") <> "OM"))
      ErrMsg = "Invalid Prd. Order Tc": Exit Sub
    Case Is = UCase("FdPrdOdYy")
      Cancel = (NewValue <> "") And (Not moCn.RecSeek("Select PMCd From Param Where PTyp= 'Yy' " + _
                                    "And PMCd = '" + NewValue + "'"))
      ErrMsg = "Invalid Prd. Order Yy": Exit Sub
    Case Is = UCase("FdPrdOdChr")
      Cancel = (NewValue <> "") And (Not moCn.RecSeek("Select vPSCd From vParam Where " + _
                                    "vPCoCd = '" + ADC("FgCoCd") + "' and vPTyp= 'CHR' " + _
                                    "And vPMCd = '" + .Value(RowNum, "FdPrdOdTc") + "' and " + _
                                    "vPSCd='" + NewValue + "'"))
      ErrMsg = "Invalid Prd. Order Chr": Exit Sub
    Case Is = UCase("FdPrdOdNo")
      '****** Sachin 3.02 - 24-12-07 - Current Partition
      Cancel = (NewValue <> 0) And (Not moCn.RecSeek("Select OdNo From OrdDsg Where " + _
                                    "OdCoCd= '" + ADC("FgCoCd") + "' and OdTc='" + _
                                    .Value(RowNum, "FdPrdOdTc") + "'And OdYy = '" + _
                                    .Value(RowNum, "FdPrdOdYy") + "' and OdChr='" + _
                                    .Value(RowNum, "FdPRdOdChr") + "' and OdNo= " + _
                                    CStr(NewValue) + " and OdPrtKey='" + ctCurrPrtn + "' "))
      ErrMsg = "Invalid Prd. Order No": Exit Sub
      
    Case Is = UCase("FdPrdOdSr")
      '****** Sachin 3.02 - 24-12-07 - Current Partition
      Cancel = (NewValue <> 0) And (Not moCn.RecSeek("Select OdNo From OrdDsg Where " + _
                                    "OdCoCd= '" + ADC("FgCoCd") + "' and OdTc='" + _
                                    .Value(RowNum, "FdPrdOdTc") + "'And OdYy = '" + _
                                    .Value(RowNum, "FdPrdOdYy") + "' and OdChr='" + _
                                    .Value(RowNum, "FdPRdOdChr") + "' and OdNo= " + _
                                    CStr(.Value(RowNum, "FdPrdOdNo")) + " and OdSr=" + _
                                    CStr(NewValue) + " and OdPrtKey='" + ctCurrPrtn + "' "))
      ErrMsg = "Invalid Prd. Order Sr": Exit Sub
      
    '****** Sachin 3.02 (14/09/07) Id Fields
      'GrdFgd.Value(RowNum, "FdPrdOdIdNo") = moCn.GetFldVal("Select OdIdNo From OrdDsg " + _
                                        "Where OdCoCd='" + adc("FgCoCd") + "' And " + _
                                        " OdTc='" + .Value(RowNum, "FdPrdOdTc") + "' And " + _
                                        " OdYy='" + .Value(RowNum, "FdPrdOdYy") + "' And " + _
                                        " OdChr='" + .Value(RowNum, "FdPrdOdChr") + "' And " + _
                                        " OdNo=" + CStr(.Value(RowNum, "FdPrdOdNo")) + " And " + _
                                        " OdSr=" + CStr(.Value(RowNum, "FdPrdOdSr")))
    '****** Sachin 3.02 (14/09/07) Id Fields
    
    Case Is = UCase("FdPrdCmCd")
      Cancel = (NewValue <> "") And (Not moCn.RecSeek("Select CmCd From CustMst Where CmCtg='C' " + _
                                     "and CmCd='" + NewValue + "'"))
      ErrMsg = "Invalid Prd. Customer Code": Exit Sub
    
    '***** urmi 2.04 ****************
    Case Is = UCase("FdSubLoc")
      If ms_TcTyp = "FB" Then
        Cancel = Not moCn.RecSeek("Select vPmCd from vParam where " + _
              "vPCoCd = '" + ADC("FgCoCd") + "' and " + _
              "vPTyp = 'FGSUBLOC' and vPMcd = '" + NewValue + "'")
        If Cancel = True Then ErrMsg = "Invalid Fg Sub Location": Exit Sub
                             
        If Not moCn.RecSeek("select 'x' From vParam where vPTyp='FgSubLoc' and vPCoCd= '" + gs_CoCd + "' and vPMCd='" + ADC("FgSubLoc") + "'" + _
                             " and (vPValue1='' Or ','+vPValue1+',' Like '%," + gs_UsrCd + ",%')") Then Cancel = True: ErrMsg = "The User Does Not Have Access To This Fg SubLoc": Exit Sub
        
      Else
        If NewValue <> "" Then Cancel = True: ErrMsg = "Fg SubLoc Should be Blank": Exit Sub
      End If
    '***** urmi 2.04 ****************
    'Uma KBS
    Case Is = UCase("FdPlYy")
    If (NewValue <> "") Then
      Cancel = Not moCn.RecSeek("Select PMCd From Param Where PTyp= 'Yy' And PMCd = '" + NewValue + "' And PSCd=''")
      ErrMsg = "Invalid Packing List Year": Exit Sub
    End If
    Case Is = UCase("FdPlChr")
    If (NewValue <> "") Then
      Cancel = Not moCn.RecSeek("Select PMCd From Param Where PTyp= 'PLCHR' And PMCd = '" + NewValue + "' And PSCd=''")
      ErrMsg = "Invalid Packing List Character": Exit Sub
    End If
    End Select
  End With
End Sub
Private Sub GrdFgd_InitKey(ByVal RowNum As Integer)
  '*** Initialise Key Fields and all the fields which are not included in the grid
  
  'Dim wFdKey As Single, wBaseVchSr() As String, wFdRefYy As String, wFdRefKey As Single
  With GrdFgd
    If .IsNew(RowNum) Then
      .Store "FdCoCd", ADC("FgCoCd")
      .Store "FdTc", ADC("FgTc")
      .Store "FdYy", ADC("FgYy")
      .Store "FdChr", ADC("FgChr")
      .Store "FdNo", ADC("FgNo")
      
      '****** Sachin 3.02 (14/09/07) Id Fields
      '****** Sachin 3.02 - 24-12-07 - Current Partition
      .Store "FdFgIdNo", moCn.GetFldVal("Select FgIdNo From Fg " + _
                                        "Where FgCoCd='" + ADC("FgCoCd") + "' And FgTc='" + ADC("FgTc") + "' And FgYy='" + ADC("FgYy") + "' And FgChr='" + ADC("FgChr") + "' And FgNo=" + CStr(ADC("FgNo")) + " And FgPrtKey='" + ctCurrPrtn + "' ")
      .Store "FdDt", ADC("FgDt")

      '****** Sachin 3.02 (14/09/07) Id Fields
      
      'wFdKey = moCn.AutoGen("Fgd", "FdKey", ADC("FgYy"))
      If ms_TcTyp = "FB" Or ms_TcTyp = "FR" Then
        '.Store "FdKey", wFdKey
        '.Store "FdRefYy", adc("FgYy")
        '.Store "FdRefKey", wFdKey
        .Store "FdBM", "B"
      ElseIf ms_TcTyp = "FWB" Or ms_TcTyp = "FWR" Then
        'wBaseVchSr() = Split(.Value(RowNum, "wFdBaseVchNo"), "/")
        'wFdRefYy = moCn.GetFldVal("Select FdYy from Fgd where " + _
                   "FdTc= '" + wBaseVchSr(0) + "' and FdYy= '" + wBaseVchSr(1) + "' and " + _
                   "FdChr= '" + wBaseVchSr(2) + "' and FdNo= " + wBaseVchSr(3) + " and " + _
                   "FdSr= " + wBaseVchSr(4) + " and FdBM= 'B' ")
        'wFdRefKey = moCn.GetFldVal("Select FdKey from Fgd where " + _
                   "FdTc= '" + wBaseVchSr(0) + "' and FdYy= '" + wBaseVchSr(1) + "' and " + _
                   "FdChr= '" + wBaseVchSr(2) + "' and FdNo= " + wBaseVchSr(3) + " and " + _
                   "FdSr= " + wBaseVchSr(4) + " and FdBM= 'B' ")
        '.Store "FdKey", wFdKey
        '.Store "FdRefYy", wFdRefYy
        '.Store "FdRefKey", wFdRefKey
        .Store "FdBM", "M"
      End If
    End If
    
      '****** Sachin 3.02 (14/09/07) Id Fields
      
      .Store "FdBIdNo", moCn.GetFldVal("Select BIdNo From Bag Where " + _
                                                    " BCoCd='" + ADC("FgCoCd") + "' And BYy='" + .Value(RowNum, "FdBYy") + "' And BChr = '" + .Value(RowNum, "FdBChr") + "' And BNo=" + CStr(.Value(RowNum, "FdBNo")) + IIF(gs_Partition = ctCurrPrtn, " and BPrtKey='" + ctCurrPrtn + "' ", ""))



      '****** Sachin 3.02 (14/09/07) Id Fields
      'indigo.164 FdDmTcTyp added
       Dim wDmTcTyp As String
       wDmTcTyp = moCn.GetFldVal("Select BOdDmTcTyp From Bag Where " + _
                                                    " BCoCd='" + ADC("FgCoCd") + "' And BYy='" + .Value(RowNum, "FdBYy") + "' And BChr = '" + .Value(RowNum, "FdBChr") + "' And BNo=" + CStr(.Value(RowNum, "FdBNo")) + IIF(gs_Partition = ctCurrPrtn, " and BPrtKey='" + ctCurrPrtn + "' ", ""))
      .Store "FdDmTcTyp", wDmTcTyp
    
      'indigo.164 DmTcTyp='DM' removed
      .Store "FdDmIdNo", moCn.GetFldVal("Select DmIdNo From DsgMst " + _
                                        "Where DmTcTyp='" + wDmTcTyp + "' and DmCd='" + GrdFgd.Value(RowNum, "FdDmCd") + "' And DmSz='' " + IIF(gs_Partition = ctCurrPrtn, " And DmPrtKey='" + ctCurrPrtn + "'", ""))
  End With
End Sub
Private Sub GrdFgd_Valid()
  'CmdCP.Enabled = False: CmdEX.Enabled = False
End Sub

Private Sub GrdFgRm_When()
  '### Old Code
  'If GrdFgd.Value(GrdFgd.Row, "FdBNo") <> 0 Then
  '  GrdFgRm.AllowAdd = False
  '  GrdFgRm.AllowDelete = False
  'End If
  '### Old Code
End Sub

Private Sub GrdFgRm_RowWhen(ByVal RowNum As Integer)
  '*** If GrdFgRm.IsNew Then FrSrNo = Max + 1
  With GrdFgRm
    If .IsNew(RowNum) Then .Value(RowNum, "FrSrNo") = .MaxVal("FrSrNo") + 1
  End With
End Sub

Private Sub GrdFgRm_ColWhen(ByVal RowNum As Integer, ByVal ColName As String, Cancel As Boolean, ErrMsg As String, ByVal Action As MwfCtl.en_fgAction)
  '*** Set Helps For Each Field
  '*** Cannot edit any value in this row of the grid if Despatch Quantity exists
  '*** FrFrRmLoc
      'If FrFrRmLoc = Blank then give default value to the FrFrRmLoc field
      'If First Row then FrFrRmLoc = FgFrRmLoc else FrFrRmLoc = value in previous row
  '*** wFrFrRmIr
      'If wFrFrRmIr = Blank then give default value to the working field wFrFrRmIr
      'If First Row then wFrFrRmIr = 'I' else wFrFrRmIr = value in previous row
  '*** FrRmCd
      'Show only those RM Cds in the help which belongs to the Export Order
  '*** FrRmSz
      ' Check whether the Rm records have to be entered Size wise or Not
  '*** FrRmStkRt
      ' Check whether the Rm records have to be entered Stock Rate wise or Not
  '*** FrToRmLoc
      'If FrToRmLoc = Blank then give default value to the FrToRmLoc field
      'If First Row then FrToRmLoc = FgToRmLoc else FrToRmLoc = value in previous row

  If mf_DtLocked = True Then Cancel = True: ErrMsg = "Cannot Edit Locked Entry": Exit Sub
  If mf_ValidLocSeq = False Then Cancel = True: ErrMsg = "The User Does Not Have Access To One Of The Locs": Exit Sub
  If GrdFgd.Value(GrdFgd.Row, "FdDespQty") <> 0 Then Cancel = True: Exit Sub
  
  With GrdFgRm
    Select Case ColName
    Case Is = UCase("FrFrRmLoc")
      If .IsNew(RowNum) Then
        If RowNum = 1 Then .Value(RowNum, "FrFrRmLoc") = ADC("FgFrRmLoc")
        If RowNum > 1 Then .Value(RowNum, "FrFrRmLoc") = .Value(RowNum - 1, "FrFrRmLoc")
      End If
      If ms_TcTyp = "FB" Then Call HlpList.Loc(ADC("FgCoCd"), "'W','F'")
      If ms_TcTyp = "FR" Then Call HlpList.Loc(ADC("FgCoCd"), , "'P','W'")
      
    Case Is = UCase("wFrFrRmIr")
      If .IsNew(RowNum) Then
        If RowNum = 1 Then .Value(RowNum, "wFrFrRmIr") = IIF(ADC("wFrRmDc") = "C", "I", _
                                                         IIF(ADC("wFrRmDc") = "D", "R", ""))
        If RowNum > 1 Then .Value(RowNum, "wFrFrRmIr") = .Value(RowNum - 1, "wFrFrRmIr")
      End If
      Call HlpList.PMCd("IR")
      
    Case Is = UCase("FrRmCd")
      If ms_TcTyp = "FR" Then
        Call HlpList.OrRmCd(ADC("FgCoCd"), GrdFgd.Value(GrdFgd.Row, "FdExpOdTc"), _
                            GrdFgd.Value(GrdFgd.Row, "FdExpOdYy"), _
                            GrdFgd.Value(GrdFgd.Row, "FdExpOdChr"), _
                            GrdFgd.Value(GrdFgd.Row, "FdExpOdNo"), _
                            GrdFgd.Value(GrdFgd.Row, "FdExpOdSr"))
      End If
      'Call HlpList.RmCd("N")
      
    Case Is = UCase("FrRmSz")
      Dim wFdRmSzYn  As String
      '****** Sachin 3.02 - 24-12-07 - Current Partition
      wFdRmSzYn = moCn.GetFldVal("Select Pdesc225 from Param join RmMst on RmCd='" + _
                  .Value(RowNum, "FrRmCd") + "' where PTyp='RmSCtg' and PmCd=RmCtg and RmPrtKey='" + ctCurrPrtn + "' and " + _
                  " PScd = RmSCtg ")
      If wFdRmSzYn = "N" Then ErrMsg = "Size Should Not Be Entered For This Rm SubCtg": Cancel = True
      
    Case Is = UCase("FrRmStkRt")
      Dim wFdRmStkRtYn As String
      '****** Sachin 3.02 - 24-12-07 - Current Partition
      wFdRmStkRtYn = moCn.GetFldVal("Select PValue1 from Param join RmMst on RmCd='" + _
                     .Value(RowNum, "FrRmCd") + "' and RmPrtKey='" + ctCurrPrtn + "' where PTyp='RmSCtg' and PmCd=RmCtg and " + _
                     " PScd = RmSCtg ")
      If wFdRmStkRtYn = "N" Then ErrMsg = "Stock Rate Should Not Be Entered For This Rm SubCtg": Cancel = True
      
    Case Is = UCase("FrToRmLoc")
      If .Value(RowNum, "FrToRmLoc") = "" Then
        If RowNum = 1 Then .Value(RowNum, "FrToRmLoc") = ADC("FgToRmLoc")
        If RowNum > 1 Then .Value(RowNum, "FrToRmLoc") = .Value(RowNum - 1, "FrToRmLoc")
      End If
      If ms_TcTyp = "FB" Then Call HlpList.Loc(ADC("FgCoCd"), "'W','F'")
      If ms_TcTyp = "FR" Then Call HlpList.Loc(ADC("FgCoCd"), , "'P','W'")
    End Select
  End With
End Sub
Private Sub GrdFgRm_ReCalc(ByVal RowNum As Integer, ByVal ColName As String, ByVal OldVal As Variant, ByVal SourceName As String, ByVal SourceColName As String)
  '*** wFrFrRmIr on FrFrRmDc              Init
      'If FrFrRmDc= 'C' Then wFrFrRmIr= 'I'
      'If FrFrRmDc= 'D' Then wFrFrRmIr= 'R'
  '*** FrFrRmDc on wFrFrRmIr              Copy,Normal
      'If wFrFrRmIr= 'I' Then FrFrRmDc= 'C'
      'If wFrFrRmIr= 'R' Then FrFrRmDc= 'D'
  '*** FrRmSz on FrRmCd                   Normal
      ' If the Ctg or SubCtg of the old value of FrRmCd is different from the
      ' Ctg or SubCtg of the new value of FrRmCd then FrRmSz = 0
  '*** FrRmStkRt on FrRmCd                Normal
      ' If the Ctg or SubCtg of the old value of FrRmCd is different from the
      ' Ctg or SubCtg of the new value of FrRmCd then FrRmStkRt= 0
  '*** FrRmQty on FrRmCd                  Normal
      ' If the Ctg or SubCtg of the old value of FrRmCd is different from the
      ' Ctg or SubCtg of the new value of FrRmCd then FrRmQty= 0
  '*** FrRmWt on FrRmCd                   Normal
      ' If the Ctg or SubCtg of the old value of FrRmCd is different from the
      ' Ctg or SubCtg of the new value of FrRmCd then FrRmWt= 0
  
  With GrdFgRm
    Select Case ColName
    Case Is = UCase("wFrFrRmIr")
      If .Mode = fgminit Then
        If .Value(RowNum, "FrFrRmDc") = "C" Then .Value(RowNum, "wFrFrRmIr") = "I"
        If .Value(RowNum, "FrFrRmDc") = "D" Then .Value(RowNum, "wFrFrRmIr") = "R"
      End If
    Case Is = UCase("FrFrRmDc")
      If .Mode = fgmnorm Then
        If .Value(RowNum, "wFrFrRmIr") = "I" Then .Value(RowNum, "FrFrRmDc") = "C"
        If .Value(RowNum, "wFrFrRmIr") = "R" Then .Value(RowNum, "FrFrRmDc") = "D"
      End If
      
' Zubin 212
    Case Is = UCase("FrLotNo")
      If .Mode = fgmnorm Then
        If FnRmCtg(.Value(RowNum, "FrRmCd")) <> FnRmCtg(OldVal) Or _
           FnRmSCtg(.Value(RowNum, "FrRmCd")) <> FnRmSCtg(OldVal) Then _
           .Value(RowNum, "FrLotNo") = ""
      End If
' Zubin 212

    Case Is = UCase("FrRmSz")
      If .Mode = fgmnorm Then
        If FnRmCtg(.Value(RowNum, "FrRmCd")) <> FnRmCtg(OldVal) Or _
           FnRmSCtg(.Value(RowNum, "FrRmCd")) <> FnRmSCtg(OldVal) Then _
           .Value(RowNum, "FrRmSz") = 0
      End If
    Case Is = UCase("FrRmStkRt")
      If .Mode = fgmnorm Then
        If FnRmCtg(.Value(RowNum, "FrRmCd")) <> FnRmCtg(OldVal) Or _
           FnRmSCtg(.Value(RowNum, "FrRmCd")) <> FnRmSCtg(OldVal) Then _
           .Value(RowNum, "FrRmStkRt") = 0
      End If
    Case Is = UCase("FrRmQty")
      If .Mode = fgmnorm Then
        If FnRmCtg(.Value(RowNum, "FrRmCd")) <> FnRmCtg(OldVal) Or _
           FnRmSCtg(.Value(RowNum, "FrRmCd")) <> FnRmSCtg(OldVal) Then _
           .Value(RowNum, "FrRmQty") = 0
      End If
    Case Is = UCase("FrRmWt")
      If .Mode = fgmnorm Then
        If FnRmCtg(.Value(RowNum, "FrRmCd")) <> FnRmCtg(OldVal) Or _
           FnRmSCtg(.Value(RowNum, "FrRmCd")) <> FnRmSCtg(OldVal) Then _
           .Value(RowNum, "FrRmWt") = 0
      End If
      
    '****** Sachin 3.02 - [16/10/07] ******
    '****** Sachin 3.02.0 - Shifted to InitKey
    '    Case Is = UCase("FrRmCtg")
    '      If .Mode = fgmnorm Or .Mode = fgmCopy Or .Mode = fgminit Then
    '         .Value(RowNum, "FrRmCtg") = moCn.GetFldVal("Select RmCtg From RmMst " + _
    '                                                  " Where RmCd='" + .Value(RowNum, "FrRmCd") + "' And RmPrtKey='" + ctCurrPrtn + "'")
    '      End If
    '
    '    Case Is = UCase("FrRmSCtg")
    '      If .Mode = fgmnorm Or .Mode = fgmCopy Or .Mode = fgminit Then
    '         .Value(RowNum, "FrRmSCtg") = moCn.GetFldVal("Select RmSCtg From RmMst " + _
    '                                                  " Where RmCd='" + .Value(RowNum, "FrRmCd") + "' And RmPrtKey='" + ctCurrPrtn + "'")
    '      End If
    '
    '    Case Is = UCase("FrFrRmLocTyp")
    '      If .Mode = fgmnorm Or .Mode = fgmCopy Or .Mode = fgminit Then
    '         .Value(RowNum, "FrFrRmLocTyp") = moCn.GetFldVal("Select LocTyp From Loc Where LocCoCd='" + adc("FgCoCd") + "' " + _
    '                                                  " And LocCd='" + .Value(RowNum, "FrFrRmLoc") + "'")
    '      End If
    '
    '    Case Is = UCase("FrToRmLocTyp")
    '      If .Mode = fgmnorm Or .Mode = fgmCopy Or .Mode = fgminit Then
    '         .Value(RowNum, "FrToRmLocTyp") = moCn.GetFldVal("Select LocTyp From Loc Where LocCoCd='" + adc("FgCoCd") + "' " + _
    '                                                  " And LocCd='" + .Value(RowNum, "FrToRmLoc") + "'")
    '      End If
    '****** Sachin 3.02 - [16/10/07] ******
      
    End Select
  End With
End Sub
Private Sub GrdFgRm_ColChange(ByVal RowNum As Integer, ByVal ColName As String, ByVal NewValue As Variant, Cancel As Boolean, ErrMsg As String)
  '*** FrFrRmLoc          valid LocCd from Loc(LocTyp<> 'P')
  '*** wFrFrRmIr          valid PMCd('IR')
  '*** FrRmCd             valid OrRmCd From OrdRm for the Export Order
  '*** FrToRmLoc          valid LocCd from Loc(LocTyp<> 'P')
  
  With GrdFgRm
    Select Case ColName
    Case Is = UCase("FrFrRmLoc")
      Cancel = Not moCn.RecSeek("Select LocCd From Loc Where LocCoCd= '" + ADC("FgCoCd") + "' and " + _
                           "LocTyp<>'P' and LocCd = '" + NewValue + "'")
      ErrMsg = "Invalid Raw Material From Location": Exit Sub
    Case Is = UCase("wFrFrRmIr")
      Cancel = Not moCn.RecSeek("Select PMCd From Param Where PTyp= 'IR' And " + _
                           "PMCd = '" + NewValue + "'")
      ErrMsg = "Enter Issue Or Receipt (I/R)": Exit Sub
    Case Is = UCase("FrRmCd")
      '****** Sachin 3.02 25-12-07 - Current Partition
      Cancel = Not moCn.RecSeek("Select OrRmCd From OrdRm where " + _
               "OrCoCd= '" + ADC("FgCoCd") + "' and " + _
               "OrTc= '" + GrdFgd.Value(GrdFgd.Row, "FdExpOdTc") + "' and " + _
               "OrYy= '" + GrdFgd.Value(GrdFgd.Row, "FdExpOdYy") + "' and " + _
               "OrChr= '" + GrdFgd.Value(GrdFgd.Row, "FdExpOdChr") + "' and " + _
               "OrNo= " + CStr(GrdFgd.Value(GrdFgd.Row, "FdExpOdNo")) + " and " + _
               "OrSr= " + CStr(GrdFgd.Value(GrdFgd.Row, "FdExpOdSr")) + " and " + _
               "OrRmCd= '" + NewValue + "' and OrPrtKey='" + ctCurrPrtn + "'")
      'Cancel = Not moCn.RecSeek("Select RmCd From RmMst where RmZ= 'N' and RmCd= '" + NewValue + "'")
      ErrMsg = "Invalid Raw Material Code (It Should Belong To Exp Order)": Exit Sub
    Case Is = UCase("FrToRmLoc")
      Cancel = Not moCn.RecSeek("Select LocCd From Loc Where LocCoCd= '" + ADC("FgCoCd") + "' and " + _
                          "LocTyp<>'P' and LocCd = '" + NewValue + "'")
      ErrMsg = "Invalid Raw Material To Location": Exit Sub
    End Select
  End With
End Sub
Private Sub GrdFgRm_InitKey(ByVal RowNum As Integer)
  '*** Initialise Key Fields and all the fields which are not included in the grid
  With GrdFgRm
    If .IsNew(RowNum) Then
      .Store "FrCoCd", ADC("FgCoCd")
      .Store "FrTc", ADC("FgTc")
      .Store "FrYy", ADC("FgYy")
      .Store "FrChr", ADC("FgChr")
      .Store "FrNo", ADC("FgNo")
      .Store "FrSr", IIF(GrdFgd.Row > 0, _
                         GrdFgd.Value(GrdFgd.Row, "FdSr"), GrdFgd.Value(1, "FdSr"))
                               
      '****** Sachin 3.02 (14/09/07) Id Fields
      .Store "FrFdIdNo", moCn.GetFldVal("Select FdIdNo From Fgd " + _
                                        "Where FdCoCd='" + ADC("FgCoCd") + "' And FdTc='" + ADC("FgTc") + "' And FdYy='" + ADC("FgYy") + "' And FdChr='" + ADC("FgChr") + "' And FdNo=" + CStr(ADC("FgNo")) + " And FdSr=" + CStr(IIF(GrdFgd.Row > 0, _
                                        GrdFgd.Value(GrdFgd.Row, "FdSr"), GrdFgd.Value(1, "FdSr"))) + " and FdPrtKey='" + ctCurrPrtn + "'")
      .Store "FrDt", ADC("FgDt")
         
      .Store "FrRmCtg", moCn.GetFldVal("Select RmCtg From RmMst " + _
                                                  " Where RmCd='" + .Value(RowNum, "FrRmCd") + "' And RmPrtKey='" + ctCurrPrtn + "'")
      
      .Store "FrRmSCtg", moCn.GetFldVal("Select RmSCtg From RmMst " + _
                                                  " Where RmCd='" + .Value(RowNum, "FrRmCd") + "' And RmPrtKey='" + ctCurrPrtn + "'")
      
      .Store "FrFrRmLocTyp", moCn.GetFldVal("Select LocTyp From Loc Where LocCoCd='" + ADC("FgCoCd") + "' " + _
                                                  " And LocCd='" + .Value(RowNum, "FrFrRmLoc") + "'")
      
      .Store "FrToRmLocTyp", moCn.GetFldVal("Select LocTyp From Loc Where LocCoCd='" + ADC("FgCoCd") + "' " + _
                                                  " And LocCd='" + .Value(RowNum, "FrToRmLoc") + "'")
      '****** Sachin 3.02 (14/09/07) Id Fields
                         
    End If
    .Store "FrToRmDc", IIF(.Value(RowNum, "FrFrRmDc") = "D", "C", _
                       IIF(.Value(RowNum, "FrFrRmDc") = "C", "D", ""))
  End With
End Sub
Private Sub ADC_SetRecSource()
  '*** Set the RecSource for ADC (Table Fg)
  '****** Sachin 3.02 25-12-07 - Corresponding Partition
    
  ADC.RecSource = " Select * from Fg " + _
                  " Where FgCoCd= '" + ADC("FgCoCd") + "' " + _
                  " And FgTc ='" + ADC("FgTc") + "'" + _
                  " And FgYy = '" + ADC("FgYy") + "'" + _
                  " And FgChr ='" + ADC("FgChr") + "' " + _
                  " And FgNo =" + CStr(ADC("FgNo")) + _
                  " And FgPrtKey=(Case When '" + gs_Partition + "'='" + ctCurrPrtn + "' Then '" + ctCurrPrtn + "' Else FgPrtKey End)" + _
                  IIF(gs_Partition = ctCurrPrtn, " And Fgdt>'" + ms_PurgeDt + "' ", "")
End Sub

Private Sub GrdFgd_SetRecSource()
  '****** Sachin 3.02 25-12-07 - Corresponding Partition
  '*** Set the RecSource for GrdFgd (Table Fgd)
  GrdFgd.RecSource = " Select * from Fgd " + _
                     " Where FdCoCd= '" + ADC("FgCoCd") + "' " + _
                     " And FdTc ='" + ADC("FgTc") + "'" + _
                     " And FdYy = '" + ADC("FgYy") + "'" + _
                     " And FdChr ='" + ADC("FgChr") + "' " + _
                     " And FdNo =" + CStr(ADC("FgNo")) + _
                     IIF(gs_Partition = ctCurrPrtn, " And FdPrtKey ='" + ctCurrPrtn + "' ", "") + _
                     " Order By FdSr"
End Sub

Private Sub GrdFgRm_SetRecSource()
  '****** Sachin 3.02 25-12-07 - Corresponding Partition
  '*** Set the RecSource for GrdTxndB (Table FgRm)
  GrdFgRm.RecSource = " Select * from FgRm " + _
                      " Where FrCoCd= '" + ADC("FgCoCd") + "' " + _
                      " And FrTc ='" + ADC("FgTc") + "'" + _
                      " And FrYy = '" + ADC("FgYy") + "'" + _
                      " And FrChr ='" + ADC("FgChr") + "' " + _
                      " And FrNo =" + CStr(ADC("FgNo")) + _
                      " And FrSr =" + CStr(GrdFgd.Value(GrdFgd.Row, "FdSr")) + _
                      " And FrPrtKey ='" + GrdFgd.Value(GrdFgd.Row, "FdPrtKey") + "'" + _
                      " Order By FrSrNo"
End Sub

Private Sub adc_AutoGen()
  '*** write code for AutoGenerating the Finished Goods Voucher Number
  If ADC("FgNo") = 0 Then
    ADC("FgNo") = moCn.AutoGen("Fg", "FgNo", ADC("FgCoCd") + ADC("FgTc") + ADC("FgYy") + ADC("FgChr"))
  Else
    Call moCn.AutoGen("Fg", "FgNo", ADC("FgCoCd") + ADC("FgTc") + ADC("FgYy") + ADC("FgChr"), ADC("FgNo"))
  End If
End Sub
Private Sub GrdSumm_ColWhen(ByVal RowNum As Integer, ByVal ColName As String, Cancel As Boolean, ErrMsg As String, ByVal Action As MwfCtl.en_fgAction)
  '*** Cannot edit Summary grid values
  Cancel = True
End Sub
Private Sub CmdEX_Click()
  '*** Display or hide the Export Details frame depending on the toggle state
  '*** The CausesValidation property of this command button is false. The purpose of this is
      'when the Export Details option is selected, the grid GrdOrdDsg should not be fully
      'validated i.e. only the validation of the column on which the cursor was previously
      'standing should take place and the row itself should not be saved. When the user exits
      'out of this option, the cursor returns to the same row. The reason is some of the
      'fields can be filled up only by the user when he comes to the Export Details frame.
      'Hence the row should not be saved when the cursor comes to the Export Details frame
      'as it might generate a trigger error
    
  Dim wi_Row As Integer
  If FraEX.Visible = False Then
    If GrdFgd.HalfValidate Then
      GrdFgd.SetFocus
      Exit Sub
    End If
    Call DispFra(ExDet)
    If GrdFgd.Rows - 1 = 0 Then Exit Sub
    wi_Row = GrdFgd.Row
    ADC("wExFdPrdCmCd") = GrdFgd.Value(wi_Row, "FdPrdCmCd")
    ADC("wExFdExpCmCd") = GrdFgd.Value(wi_Row, "FdExpCmCd")
    ADC("wExFdDespQty") = GrdFgd.Value(wi_Row, "FdDespQty")
    ADC("wExFdCTBDespQty") = GrdFgd.Value(wi_Row, "FdCTBDespQty")
    ADC("wExFdInExpNo") = GrdFgd.Value(wi_Row, "FdInExpNo")
    ADC("wExFdExpOdTc") = GrdFgd.Value(wi_Row, "FdExpOdTc")
    ADC("wExFdExpOdYy") = GrdFgd.Value(wi_Row, "FdExpOdYy")
    ADC("wExFdExpOdChr") = GrdFgd.Value(wi_Row, "FdExpOdChr")
    ADC("wExFdExpOdNo") = GrdFgd.Value(wi_Row, "FdExpOdNo")
    ADC("wExFdExpOdSr") = GrdFgd.Value(wi_Row, "FdExpOdSr")
    ADC("wExFdPlChr") = GrdFgd.Value(wi_Row, "FdPlChr")
    ADC("wExFdPlYy") = GrdFgd.Value(wi_Row, "FdPlYy")
    ADC("wExFdPlNo") = GrdFgd.Value(wi_Row, "FdPlNo")
    ADC("wExFdInTc") = GrdFgd.Value(wi_Row, "FdInTc")
    ADC("wExFdInYy") = GrdFgd.Value(wi_Row, "FdInYy")
    ADC("wExFdInChr") = GrdFgd.Value(wi_Row, "FdInChr")
    ADC("wExFdInNo") = GrdFgd.Value(wi_Row, "FdInNo")
    ADC("wExFdBaseVchNo") = GrdFgd.Value(wi_Row, "wFdBaseVchNo")
    'ADC("wExFdInNo").InitialiseFld GrdFgd.Value(wi_Row, "FdInNo")
  Else
    Call DispFra(ExDet)
    If GrdFgd.Rows - 1 = 0 Then Exit Sub
    wi_Row = GrdFgd.Row
    GrdFgd.Value(wi_Row, "FdPrdCmCd") = ADC("wExFdPrdCmCd")
    GrdFgd.Value(wi_Row, "FdExpCmCd") = ADC("wExFdExpCmCd")
    GrdFgd.Value(wi_Row, "FdDespQty") = ADC("wExFdDespQty")
    GrdFgd.Value(wi_Row, "FdCTBDespQty") = ADC("wExFdCTBDespQty")
    GrdFgd.Value(wi_Row, "FdInExpNo") = ADC("wExFdInExpNo")
    GrdFgd.Value(wi_Row, "FdExpOdTc") = ADC("wExFdExpOdTc")
    GrdFgd.Value(wi_Row, "FdExpOdYy") = ADC("wExFdExpOdYy")
    GrdFgd.Value(wi_Row, "FdExpOdChr") = ADC("wExFdExpOdChr")
    GrdFgd.Value(wi_Row, "FdExpOdNo") = ADC("wExFdExpOdNo")
    GrdFgd.Value(wi_Row, "FdExpOdSr") = ADC("wExFdExpOdSr")
    GrdFgd.Value(wi_Row, "FdPlChr") = ADC("wExFdPlChr")
    GrdFgd.Value(wi_Row, "FdPlYy") = ADC("wExFdPlYy")
    GrdFgd.Value(wi_Row, "FdPlNo") = ADC("wExFdPlNo")
    GrdFgd.Value(wi_Row, "FdInTc") = ADC("wExFdInTc")
    GrdFgd.Value(wi_Row, "FdInYy") = ADC("wExFdInYy")
    GrdFgd.Value(wi_Row, "FdInChr") = ADC("wExFdInChr")
    GrdFgd.Value(wi_Row, "FdInNo") = ADC("wExFdInNo")
    GrdFgd.Value(wi_Row, "wFdBaseVchNo") = ADC("wExFdBaseVchNo")
  End If
End Sub
Private Sub CmdCP_Click()
  '*** Call the Sub Procedure CpyRm (Procedure to copy the Raw Material Details)
  
  '****** The focus is set on the grid GrdFgd as the focus tends to getlost when
         'the Copy routine is called.
         
  If mf_DtLocked = True Then Exit Sub
  If mf_ValidLocSeq = False Then DispMsg "The User Does Not Have Access To One Of The Locs", etError: Exit Sub
  MWLib.BeginProcess Me, "Copy in Progress"
  mf_CpyStartedByParent = True
  Call CpyRm
  MWLib.EndProcess Me
  mf_CpyStartedByParent = False
  GrdFgd.SetFocus       'New Code
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
      
     If ADC("wBrsGrp") = "" Then ADC("wBrsGrp") = "RMCTG": ADC("wVrsGrp") = "RMCTG"    '***** Sachin 3.11.0
    Call RMSummary(ADC("wBrsGrp"), "N")
    Call RMSummary(ADC("wVrsGrp"), "Y")
  End If
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
    Case Is = "E", "B", "R"
      FldEna = True
    Case Else
      FldEna = False
  End Select
End Function
Private Function FldVis(ByVal mVis As String) As Boolean
  '*** Function to return the Visible property of the controls on the form
  Select Case mVis
    Case Is = "V", "B", "R"
      FldVis = True
    Case Else
      FldVis = False
  End Select
End Function
Private Sub PropSet()
  '*** Setting Vis and Ena Properties for X Fields ***
  Dim wRsX As MwfLib.MDORowSet
  Dim wPValue1 As String
  wPValue1 = moCn.GetFldVal("Select PValue1 from Param where PTyp= 'TCTYP' and PMCd= '" + ms_TcTyp + "' and PSCd= ''")
  Set wRsX = moCn.OpenRes("Select * from TcStr where TsTbl='" + wPValue1 + "' and " + _
                           "TsTyp='" + ms_TcTyp + "' and TsSTyp='X'")
  Do While Not (wRsX.EOF Or wRsX.BOF)
    ADC(wRsX!TsIdName).Enabled = FldEna(wRsX!TsStyle)
    ADC(wRsX!TsIdName).Visible = FldVis(wRsX!TsStyle)
    wRsX.MoveNext
  Loop

  '*** Setting Vis and Ena Properties for A Fields ***
  Call GrdPropSet(GrdFgd, "A")

  '*** Setting Vis and Ena Properties for B Fields ***
  Call GrdPropSet(GrdFgRm, "B")

  '*** Setting Vis and Ena Properties for Label Conrols ***
  Dim wRsLbl As MwfLib.MDORowSet
  Set wRsLbl = moCn.OpenRes("Select * from TcStr where TsTbl='" + wPValue1 + "' and " + _
                           "TsTyp= '" + ms_TcTyp + "' and TsSTyp='Lbl'")
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
  Call CmdPropSet
  Set wRsX = Nothing
  Set wRsLbl = Nothing
End Sub
Private Sub GrdPropSet(ByVal mGrd As Object, ByVal mTsSTyp As String)
  '*** Procedure to set the Style of the columns in the grids
  Dim wGrdVis As String, wPValue1 As String
  With mGrd
    wPValue1 = moCn.GetFldVal("Select PValue1 from Param where PTyp= 'TCTYP' and PMCd= '" + ms_TcTyp + "' and PSCd= ''")
    '*** Checking If The Grid Is Visible ***
    wGrdVis = moCn.GetFldVal("Select TsStyle from TcStr where TsTbl='" + wPValue1 + "' and " + _
                        "TsTyp='" + ms_TcTyp + "' and TsSTyp= '" + mTsSTyp + "' and " + _
                        "TsIdName=''")
    .Visible = FldVis(wGrdVis)
    
    If wGrdVis = "R" Then .AllowAdd = False: .AllowDelete = False

    '*** If The Grid Is Visible Then Set The Vis And Ena Properties For The Columns For
    '*** Which Record Exists in TcStructure Table
    If FldVis(wGrdVis) = True Then
      Dim wRsGrd As MwfLib.MDORowSet
      Set wRsGrd = moCn.OpenRes("Select * from TcStr where TsTbl='" + wPValue1 + "' and " + _
                                 "TsTyp='" + ms_TcTyp + "' and TsSTyp= '" + mTsSTyp + "' and " + _
                                 "TsIdName<>''")
      Do While Not (wRsGrd.EOF Or wRsGrd.BOF)
        .ColProp(wRsGrd!TsIdName).Style = ColSty(wRsGrd!TsStyle)
        wRsGrd.MoveNext
      Loop
    End If
  End With
  Set wRsGrd = Nothing
End Sub
Private Sub CmdPropSet()
  '*** Setting Vis and Ena Properties for CommandButton Conrols ***
  Dim wRsCmd As MwfLib.MDORowSet, wPValue1 As String
  wPValue1 = moCn.GetFldVal("Select PValue1 from Param where PTyp= 'TCTYP' and PMCd= '" + ms_TcTyp + "' and PSCd= ''")
  Set wRsCmd = moCn.OpenRes("Select * from TcStr where TsTbl='" + wPValue1 + "' and " + _
                           "TsTyp='" + ms_TcTyp + "' and TsSTyp='Cmd'")
  Dim wCmdCtrl As Control
  Do While Not (wRsCmd.EOF Or wRsCmd.BOF)
    For Each wCmdCtrl In Controls
      ' After ClrScr
      'If TypeOf wCmdCtrl Is CommandButton Then
      If TypeOf wCmdCtrl Is MWCTL_BTN1 Then
      ' After ClrScr
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
Private Sub CpyRm()
  '*** Procedure to Enter Raw Material details into 'B' Grid in 'FB', 'FWB' or 'FWR' entries
  '*** Do not copy when record already exists in 'B' grid
  '*** In 'FB' copy the Raw Material Details from the Txnd table
  '*** In 'FWB' and 'FWR' copy the Raw Material Details from FgRm of the base record
  
  If GrdFgd.Rows - 1 = 0 Then Exit Sub
  
  Dim wFgdRow As Single, wRow As Single
  If GrdFgd.Rows - 1 = 0 Then DispMsg "Parent Record Does Not Exist", etError: Exit Sub
  If GrdFgRm.Rows - 1 > 0 Then DispMsg "Record Already Exists", etError: Exit Sub
  wFgdRow = IIF(GrdFgd.Row, GrdFgd.Row, 1)
  Dim wRsCr As MwfLib.MDORowSet
  If ms_TcTyp = "FB" Then
    ' Zubin 212
    '****** Sachin 3.02 25-12-07 - Current Partition
    Set wRsCr = moCn.OpenRes("Select TdRmCd as qRmCd, TdLotNo as qLotNo, TdRmSz as qRmSz,TdRmSz2 as qRmSz2, TdRmSz3 as qRmSz3," + _
                "TdRmStkRt as qRmStkRt, sum((case (case when TdFrRmLoc='W' then TdFrRmDc " + _
                "when TdToRmLoc='W' then TdToRmDc else '' end) when 'C' then -1 " + _
                "when 'D' then +1 else 0 end)* TdRmQty) as qRmQty, " + _
                "sum((case (case when TdFrRmLoc='W' then TdFrRmDc when TdToRmLoc='W' then " + _
                "TdToRmDc end) when 'C' then -1 when 'D' then +1 else 0 end)* " + _
                "TdRmWt) as qRmWt From Txnd " + _
                "Where TdRmCd <> '' and TdCoCd= '" + ADC("FgCoCd") + "' and " + _
                "TdBYy= '" + GrdFgd.Value(wFgdRow, "FdBYy") + "' and " + _
                "TdBChr= '" + GrdFgd.Value(wFgdRow, "FdBChr") + "' and " + _
                "TdBNo= " + CStr(GrdFgd.Value(wFgdRow, "FdBNo")) + " and TdPrtKey='" + ctCurrPrtn + "' " + _
                "group By TdBYy, TdBChr, TdBNo, TdRmCd, TdLotNo, TdRmSz, TdRmSz2, TdRmSz3, TdRmStkRt ")
  ElseIf ms_TcTyp = "FWB" Or ms_TcTyp = "FWR" Then
    ' Zubin 212
    '****** Sachin 3.02 - Id fields In Joins
    Set wRsCr = moCn.OpenRes("Select FrRmCd as qRmCd, FrLotNo as qLotNo, FrRmSz as qRmSz,FrRmSz2 as qRmSz2,FrRmSz3 as qRmSz3, " + _
                "FrRmStkRt as qRmStkRt, FrRmQty as qRmQty, FrRmWt as qRmWt From FgRm " + _
                "join Fgd on FdIdNo=FrFdIdNo and FdCoCd= FrCoCd and FdTc= FrTc and FdYy= FrYy and FdChr= FrChr and " + _
                "FdNo= FrNo and FdSr= FrSr and FdPrtKey=FrPrtKey Where FdCoCd= '" + ADC("FgCoCd") + "' and " + _
                "FdYy= '" + GrdFgd.Value(wFgdRow, "FdRefYy") + _
                "' and FdKey= " + CStr(GrdFgd.Value(wFgdRow, "FdRefKey")) + " and FdPrtKey='" + ctCurrPrtn + "' ")
  End If
  
  GrdFgRm.StartCopy
  If Not mf_CpyStartedByParent Then
  
  MWLib.BeginProcess Me, "Copying Raw Material Details ..."
  '5713 while copy cursor not moving to next and raising error. BeginEmrProcess is the issue, so it is commented
  'BeginEmrProcess Me, "Copying Raw Material Details ..."
  End If
  With GrdFgRm
    .AllowAdd = True
    Do While Not (wRsCr.EOF Or wRsCr.BOF)
      '**** Urmi 2.06 Do not copy RM with 0 qty and 0 wt in FG
      If ms_TcTyp <> "FB" Or (wRsCr!qRmQty <> 0 Or wRsCr!qRmWt <> 0) Then
        
        .AddItem
        wRow = .Rows - 1
        UpdateEmrProcess "Copying Bag Details..." + CStr(GrdFgd.Row) + "::" + "Rm Details " + CStr(wRow)
        
        .Value(wRow, "FrSrNo") = wRow
        .Value(wRow, "FrFrRmLoc") = ADC("FgFrRmLoc")
        If ADC("wFrRmDc") = "D" Then
          .Value(wRow, "wFrFrRmIr") = "R"
          .Value(wRow, "FrFrRmDc") = "D"
        Else
          .Value(wRow, "wFrFrRmIr") = "I"
          .Value(wRow, "FrFrRmDc") = "C"
        End If
        .Value(wRow, "FrRmCd") = wRsCr!qRmCd
        ' Zubin 212
        .Value(wRow, "FrLotNo") = wRsCr!qLotNo
        ' Zubin 212
        .Value(wRow, "FrRmSz") = wRsCr!qRmSz
        .Value(wRow, "FrRmSz2") = wRsCr!qRmSz2
        .Value(wRow, "FrRmSz3") = wRsCr!qRmSz3
        .Value(wRow, "FrRmStkRt") = wRsCr!qRmStkRt
        .Value(wRow, "FrRmQty") = wRsCr!qRmQty
        .Value(wRow, "FrRmWt") = wRsCr!qRmWt
        .Value(wRow, "FrToRmLoc") = ADC("FgToRmLoc")
        .Value(wRow, "FrPrtKey") = ctCurrPrtn     '****** Sachin 3.02 25-12-07
        .SaveRec (wRow)
      End If
      wRsCr.MoveNext
    Loop
    .AllowAdd = False
  
  End With
  If Not mf_CpyStartedByParent Then
  
  MWLib.EndProcess Me
  'EndEmrProcess Me  '5713 while copy cursor not moving to next and raising error. BeginEmrProcess is the issue, so it is commented
  End If
  GrdFgRm.EndCopy
  
  Set wRsCr = Nothing
  
  '*** (Jen 2.14 Next)
  '****** Sachin 3.02 25-12-07 - Current Partition
  If moCn.RecSeek("select * from Fgd where FdCoCd= '" + ADC("FgCoCd") + "' and FdTc= '" + ADC("FgTc") + "' and " + _
                          "FdYy= '" + ADC("FgYy") + "' and FdChr= '" + ADC("FgChr") + "' and " + _
                          "FdNo= " + CStr(ADC("FgNo")) + " and FdSr= " + CStr(GrdFgd.Value(wFgdRow, "FdSr")) + " and FdPrtKey='" + ctCurrPrtn + "' " + _
                  "and convert (decimal (16, 3), FdGrWt)<> " + _
                  "convert (decimal (16, 3), IsNull((Select sum(FrRmWt/ (case when RmCtg in ('C', 'D') then 5 else 1 end)) " + _
                  "from FgRm join RmMst on RmCd= FrRmCd and RmPrtKey=FrPrtKey " + _
                  "where FrFdIdNo=FdIdNo and FrCoCd= FdCoCd and FrTc= FdTc and FrYy= FdYy and FrChr= FdChr and FrNo= FdNo " + _
                  "and FrSr= FdSr and FrPrtKey=FdPrtKey), 0))") Then
      GrdFgRm.DelAllRecs
      MsgBox " Could Not Copy All Records."
      'GrdFgd.SetFocus
      Exit Sub
  End If
  
  If ms_TcTyp = "FWB" Then Exit Sub
  If moCn.RecSeek("select 'a' from Fgd " + _
                  "join Bag on BIdNo=FdBIdNo " + _
                  "join OrdRm on OrOdIdNo=BOdIdNo " + _
                  "left outer join FgRm on FrFdIdNo=FdIdNo and FrRmCd = OrRmCd " + _
                  " where " + _
                  "FrRmCd is null" + _
      " AND BOdTc='SO' and (select hFgRmEqOrdRm from Head where HCoCd=FdCoCd " + _
  " and HCd='ZSELF')='Y' and " + _
  " (select vPValue3 from vParam where vPTyp='CHR' " + _
  " and vPCoCd=OrCoCd and vPMCd=OrTc and vPSCd=OrChr)='N' and convert(Decimal(10,3),OrPrdWt)>0.001 " + _
    "    " + _
    " AND FdCoCd= '" + ADC("FgCoCd") + "' and FdTc= '" + ADC("FgTc") + "' and " + _
                          "FdYy= '" + ADC("FgYy") + "' and FdChr= '" + ADC("FgChr") + "' and " + _
                          "FdNo= " + CStr(ADC("FgNo")) + " and FdSr= " + CStr(GrdFgd.Value(wFgdRow, "FdSr")) + " and FdPrtKey='" + ctCurrPrtn + "' ") Then
      GrdFgRm.DelAllRecs
      MsgBox "Cannot Move the bag to Fg, All the Raw materials Specified In The Order Are Not Used"
      
'      MsgBox " Could Not Copy All Records."
      'GrdFgd.SetFocus
      Exit Sub
  End If
  If moCn.RecSeek("select 'a' from Fgd " + _
                  "join Bag on BIdNo=FdBIdNo " + _
                  "join OrdRm on OrOdIdNo=BOdIdNo " + _
                  "left outer join FgRm on FrFdIdNo=FdIdNo " + _
                  "and FrRmCd = OrRmCd and   SUBSTRING (FrLotNo,1,len(rtrim(ltrim(OrLotNo)))) = rtrim(ltrim(OrLotNo))  where " + _
                  "(FrRmCd) is null" + _
    " and BOdTc='SO' and (select HOrLotYN from Head where HCoCd=FdCoCd and HCd='ZSELF')='Y' and OrRmCtg ='D' and " + _
     " (select vPValue3 from vParam where vPTyp='CHR' and vPCoCd=OrCoCd and vPMCd=OrTc and vPSCd=OrChr)='N'  and " + _
     " BPrtKey=FdPrtKey and BIdNo=FdBIdNo and OrLotNo<>''" + _
     " and FdCoCd= '" + ADC("FgCoCd") + "' and FdTc= '" + ADC("FgTc") + "' and " + _
                          "FdYy= '" + ADC("FgYy") + "' and FdChr= '" + ADC("FgChr") + "' and " + _
                          "FdNo= " + CStr(ADC("FgNo")) + " and FdSr= " + CStr(GrdFgd.Value(wFgdRow, "FdSr")) + " and FdPrtKey='" + ctCurrPrtn + "' ") Then
      GrdFgRm.DelAllRecs
      MsgBox " 'Cannt Move the Bag to Fg, All the Rm Lot Specified In The Order Are Not Used' "
      'GrdFgd.SetFocus
      Exit Sub
  End If
  
  Dim wrs_TotGmChkWt As MDORowSet
  Set wrs_TotGmChkWt = moCn.OpenRes("Select round(MIN(OdMinWt),3) as qOdMinWt,round(MAX(OdMaxWt),3) as qOdMaxWt, round((Select sum(FrRmWt* 1/ (case when FrRmCtg in ('C', 'D') then 5 else 1 end) )),3) as qFgRmWt " + _
 " From FgRm join Fgd on FdIdNo=FrFdIdNo join Bag on BIdNo=FdBIdNo join OrdDsg on OdIdNo=BOdIdNo join RmMst on FrRmCd= RmCd  " + _
 " Where " + _
" ','+ (Select PDesc225 From Param where PTyp= 'CPYRT' and PmCd= OdGmChk)+ ',' like '%,'+ RmCtg+ ',%' " + _
 "  and " + _
      " FdCoCd= '" + ADC("FgCoCd") + "' and FdTc= '" + ADC("FgTc") + "' and " + _
                          "FdYy= '" + ADC("FgYy") + "' and FdChr= '" + ADC("FgChr") + "' and " + _
                          "FdNo= " + CStr(ADC("FgNo")) + " and FdSr= " + CStr(GrdFgd.Value(wFgdRow, "FdSr")) + " and FdPrtKey='" + ctCurrPrtn + "' ")
                          
If ((wrs_TotGmChkWt!qOdMinWt <> 0 And wrs_TotGmChkWt!qFgRmWt < wrs_TotGmChkWt!qOdMinWt * GrdFgd.Value(wFgdRow, "FdQty")) Or (wrs_TotGmChkWt!qOdMaxWt <> 0 And wrs_TotGmChkWt!qFgRmWt > GrdFgd.Value(wFgdRow, "FdQty") * wrs_TotGmChkWt!qOdMaxWt)) Then
      GrdFgRm.DelAllRecs
      MsgBox "     'Gross Gram Wt Is not within the Specified Min & Max Order Gross Wt'"
      'GrdFgd.SetFocus
End If
End Sub

Private Sub oLDCpyRm()
  '*** Procedure to Enter Raw Material details into 'B' Grid in 'FB', 'FWB' or 'FWR' entries
  '*** Do not copy when record already exists in 'B' grid
  '*** In 'FB' copy the Raw Material Details from the Txnd table
  '*** In 'FWB' and 'FWR' copy the Raw Material Details from FgRm of the base record
  
  If GrdFgd.Rows - 1 = 0 Then Exit Sub
  Dim wFgdRow As Single, wRow As Single
  If GrdFgd.Rows - 1 = 0 Then DispMsg "Parent Record Does Not Exist", etError: Exit Sub
  If GrdFgRm.Rows - 1 > 0 Then DispMsg "Record Already Exists", etError: Exit Sub
  wFgdRow = IIF(GrdFgd.Row, GrdFgd.Row, 1)
  Dim wRsCr As MwfLib.MDORowSet
  If ms_TcTyp = "FB" Then
    ' Zubin 212
    '****** Sachin 3.02 25-12-07 - Current Partition
    Set wRsCr = moCn.OpenRes("Select TdRmCd as qRmCd, TdLotNo as qLotNo, TdRmSz as qRmSz,TdRmSz2 as qRmSz2, TdRmSz3 as qRmSz3," + _
                "TdRmStkRt as qRmStkRt, sum((case (case when TdFrRmLoc='W' then TdFrRmDc " + _
                "when TdToRmLoc='W' then TdToRmDc else '' end) when 'C' then -1 " + _
                "when 'D' then +1 else 0 end)* TdRmQty) as qRmQty, " + _
                "sum((case (case when TdFrRmLoc='W' then TdFrRmDc when TdToRmLoc='W' then " + _
                "TdToRmDc end) when 'C' then -1 when 'D' then +1 else 0 end)* " + _
                "TdRmWt) as qRmWt From Txnd " + _
                "Where TdRmCd <> '' and TdCoCd= '" + ADC("FgCoCd") + "' and " + _
                "TdBYy= '" + GrdFgd.Value(wFgdRow, "FdBYy") + "' and " + _
                "TdBChr= '" + GrdFgd.Value(wFgdRow, "FdBChr") + "' and " + _
                "TdBNo= " + CStr(GrdFgd.Value(wFgdRow, "FdBNo")) + " and TdPrtKey='" + ctCurrPrtn + "' " + _
                "group By TdBYy, TdBChr, TdBNo, TdRmCd, TdLotNo, TdRmSz, TdRmSz2, TdRmSz3, TdRmStkRt ")
  ElseIf ms_TcTyp = "FWB" Or ms_TcTyp = "FWR" Then
    ' Zubin 212
    '****** Sachin 3.02 - Id fields In Joins
    Set wRsCr = moCn.OpenRes("Select FrRmCd as qRmCd, FrLotNo as qLotNo, FrRmSz as qRmSz,FrRmSz2 as qRmSz2,FrRmSz3 as qRmSz3, " + _
                "FrRmStkRt as qRmStkRt, FrRmQty as qRmQty, FrRmWt as qRmWt From FgRm " + _
                "join Fgd on FdIdNo=FrFdIdNo and FdCoCd= FrCoCd and FdTc= FrTc and FdYy= FrYy and FdChr= FrChr and " + _
                "FdNo= FrNo and FdSr= FrSr and FdPrtKey=FrPrtKey Where FdCoCd= '" + ADC("FgCoCd") + "' and " + _
                "FdYy= '" + GrdFgd.Value(wFgdRow, "FdRefYy") + _
                "' and FdKey= " + CStr(GrdFgd.Value(wFgdRow, "FdRefKey")) + " and FdPrtKey='" + ctCurrPrtn + "' ")
  End If
  GrdFgRm.StartCopy
'  MWLib.BeginProcess Me, "Copying Raw Material Details ..."
  With GrdFgRm
    .AllowAdd = True
    Do While Not (wRsCr.EOF Or wRsCr.BOF)
      '**** Urmi 2.06 Do not copy RM with 0 qty and 0 wt in FG
      If ms_TcTyp <> "FB" Or (wRsCr!qRmQty <> 0 Or wRsCr!qRmWt <> 0) Then
        .AddItem
        wRow = .Rows - 1
        .Value(wRow, "FrSrNo") = wRow
        .Value(wRow, "FrFrRmLoc") = ADC("FgFrRmLoc")
        If ADC("wFrRmDc") = "D" Then
          .Value(wRow, "wFrFrRmIr") = "R"
          .Value(wRow, "FrFrRmDc") = "D"
        Else
          .Value(wRow, "wFrFrRmIr") = "I"
          .Value(wRow, "FrFrRmDc") = "C"
        End If
        .Value(wRow, "FrRmCd") = wRsCr!qRmCd
        ' Zubin 212
        .Value(wRow, "FrLotNo") = wRsCr!qLotNo
        ' Zubin 212
        .Value(wRow, "FrRmSz") = wRsCr!qRmSz
        .Value(wRow, "FrRmSz2") = wRsCr!qRmSz2
        .Value(wRow, "FrRmSz3") = wRsCr!qRmSz3
        .Value(wRow, "FrRmStkRt") = wRsCr!qRmStkRt
        .Value(wRow, "FrRmQty") = wRsCr!qRmQty
        .Value(wRow, "FrRmWt") = wRsCr!qRmWt
        .Value(wRow, "FrToRmLoc") = ADC("FgToRmLoc")
        .Value(wRow, "FrPrtKey") = ctCurrPrtn     '****** Sachin 3.02 25-12-07
        .SaveRec (wRow)
      End If
      wRsCr.MoveNext
    Loop
    .AllowAdd = False
  
  End With
 ' MWLib.EndProcess Me
  GrdFgRm.EndCopy
  
  Set wRsCr = Nothing
  
  '*** (Jen 2.14 Next)
  '****** Sachin 3.02 25-12-07 - Current Partition
  If moCn.RecSeek("select * from Fgd where FdCoCd= '" + ADC("FgCoCd") + "' and FdTc= '" + ADC("FgTc") + "' and " + _
                          "FdYy= '" + ADC("FgYy") + "' and FdChr= '" + ADC("FgChr") + "' and " + _
                          "FdNo= " + CStr(ADC("FgNo")) + " and FdSr= " + CStr(GrdFgd.Value(wFgdRow, "FdSr")) + " and FdPrtKey='" + ctCurrPrtn + "' " + _
                  "and convert (decimal (16, 3), FdGrWt)<> " + _
                  "convert (decimal (16, 3), IsNull((Select sum(FrRmWt/ (case when RmCtg in ('C', 'D') then 5 else 1 end)) " + _
                  "from FgRm join RmMst on RmCd= FrRmCd and RmPrtKey=FrPrtKey " + _
                  "where FrFdIdNo=FdIdNo and FrCoCd= FdCoCd and FrTc= FdTc and FrYy= FdYy and FrChr= FdChr and FrNo= FdNo " + _
                  "and FrSr= FdSr and FrPrtKey=FdPrtKey), 0))") Then
      GrdFgRm.DelAllRecs
      MsgBox " Could Not Copy All Records."
      GrdFgd.SetFocus
  End If
  '*** (Jen 2.14 Next)
  
  '*** End of the Copying process
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
      'Group By FdInTc,FdInYy,FdInChr,FdInNo
      'Group By FdInExpNo
  '*** Whenever a particular group option is selected, some of the summary grid fields
      'become irrevelent. So these fields should be made invisible.

  Dim wGrpByCd  As String, wGrpBy As String, wCommStr As String
  Dim wTotRmWt As Single, wTotBagQty As Single, wTotBags As Single, wStr As String, wRow As Single
  Dim wTotBalRmWt As Single, wTotBalBagQty As Single
  wGrpByCd = mGrpByCd

  If wGrpByCd = "PRDCMCD" Then wGrpBy = "FdPrdCmCd"
  If wGrpByCd = "EXPCMCD" Then wGrpBy = "FdExpCmCd"
  If wGrpByCd = "PRDORDNO" Then wGrpBy = "FdPrdOdTc,FdPrdOdYy,FdPrdOdChr,FdPrdOdNo"
  If wGrpByCd = "PRDORDSR" Then wGrpBy = "FdPrdOdTc,FdPrdOdYy,FdPrdOdChr,FdPrdOdNo,FdPrdOdSr"
  If wGrpByCd = "EXPORDNO" Then wGrpBy = "FdPrdOdTc,FdPrdOdYy,FdPrdOdChr,FdPrdOdNo"
  If wGrpByCd = "EXPORDSR" Then wGrpBy = "FdPrdOdTc,FdPrdOdYy,FdPrdOdChr,FdPrdOdNo,FdPrdOdSr"
  If wGrpByCd = "DSGCD" Then wGrpBy = "FdDmCd"
  If wGrpByCd = "PCKLST" Then wGrpBy = "FdPlChr,FdPlYy,FdPlNo"
  If wGrpByCd = "INVNO" Then wGrpBy = "FdInTc,FdInYy,FdInChr,FdInNo"
  If wGrpByCd = "INVEXPNO" Then wGrpBy = "FdInExpNo"
  
  wTotRmWt = 0: wTotBagQty = 0: wTotBags = 0: wStr = ""
    
  '****** Sachin 3.02 25-12-07 - Corresponding Partition
  wCommStr = " sum(case when FdBNo<> 0 then 1 else 0 end) as qBags, sum(FdQty) as qBagQty, " + _
             "sum(FdGrWt) as qRmWt, sum(FdQty- FdDespQty) as qBalBagQty, " + _
             "(case when sum(FdQty)> 0 then sum(FdQty- FdDespQty)*sum(FdGrWt)/ " + _
             "sum(FdQty) else 0 end) as qBalRmWt From Fgd " + _
             "join Fg on FgIdNo=FdFgIdNo and FgCoCd= FdCoCd and FgTc=FdTc and FgYy=FdYy and FgChr=FdChr" + _
             " and FgNo=FdNo and FgPrtKey=FdPrtKey Where FdCoCd= '" + ADC("FgCoCd") + "' and FdTc= '" + ADC("FgTc") + "'" + _
             " and FdYy= '" + ADC("FgYy") + "' And " + _
             "FdChr= '" + ADC("FgChr") + "' And FdNo =" + CStr(ADC("FgNo")) + " And FdPrtKey='" + ADC("FgPrtKey") + "' " + _
             " Group By " + wGrpBy
  
  If wGrpBy = "FdPrdCmCd" Then
    wStr = "Select FdPrdCmCd as qPrdCmCd, '*' as qPrdOrd, '*' as qExpCmCd, " + _
           "'*' as qExpOrd, '*' as qDmCd, '*' as qPckLst, '*' as qInv, '*' as qExpNo, " + wCommStr
    GrdSumm.ColProp("wSummPrdCmCd").Style = fgcNormal
    GrdSumm.ColProp("wSummPrdOrd").Style = fgcinvisible
    GrdSumm.ColProp("wSummExpCmCd").Style = fgcinvisible
    GrdSumm.ColProp("wSummExpOrd").Style = fgcinvisible
    GrdSumm.ColProp("wSummDmCd").Style = fgcinvisible
    GrdSumm.ColProp("wSummPl").Style = fgcinvisible
    GrdSumm.ColProp("wSummInv").Style = fgcinvisible
    GrdSumm.ColProp("wSummInvExpNo").Style = fgcinvisible
  ElseIf wGrpBy = "FdExpCmCd" Then
    wStr = "Select '*' as qPrdCmCd, '*' as qPrdOrd, FdExpCmCd as qExpCmCd, " + _
           "'*' as qExpOrd, '*' as qDmCd, '*' as qPckLst, '*' as qInv, '*' as qExpNo, " + wCommStr
    GrdSumm.ColProp("wSummPrdCmCd").Style = fgcinvisible
    GrdSumm.ColProp("wSummPrdOrd").Style = fgcinvisible
    GrdSumm.ColProp("wSummExpCmCd").Style = fgcNormal
    GrdSumm.ColProp("wSummExpOrd").Style = fgcinvisible
    GrdSumm.ColProp("wSummDmCd").Style = fgcinvisible
    GrdSumm.ColProp("wSummPl").Style = fgcinvisible
    GrdSumm.ColProp("wSummInv").Style = fgcinvisible
    GrdSumm.ColProp("wSummInvExpNo").Style = fgcinvisible
  ElseIf wGrpBy = "FdPrdOdTc,FdPrdOdYy,FdPrdOdChr,FdPrdOdNo" Then
    wStr = "Select max(FdPrdCmCd) as qPrdCmCd, (case when FdPrdOdNo> 0 then " + _
           "FdPrdOdTc+'/'+FdPrdOdYy+'/'+FdPrdOdChr+'/'+LTrim(Str(FdPrdOdNo)) else '' end) as qPrdOrd, " + _
           "'*' as qExpCmCd, '*' as qExpOrd, '*' as qDmCd, '*' as qPckLst, '*' as qInv, '*' as qExpNo, " + wCommStr
    GrdSumm.ColProp("wSummPrdCmCd").Style = fgcNormal
    GrdSumm.ColProp("wSummPrdOrd").Style = fgcNormal
    GrdSumm.ColProp("wSummExpCmCd").Style = fgcinvisible
    GrdSumm.ColProp("wSummExpOrd").Style = fgcinvisible
    GrdSumm.ColProp("wSummDmCd").Style = fgcinvisible
    GrdSumm.ColProp("wSummPl").Style = fgcinvisible
    GrdSumm.ColProp("wSummInv").Style = fgcinvisible
    GrdSumm.ColProp("wSummInvExpNo").Style = fgcinvisible
  ElseIf wGrpBy = "FdPrdOdTc,FdPrdOdYy,FdPrdOdChr,FdPrdOdNo,FdPrdOdSr" Then
    wStr = "Select max(FdPrdCmCd) as qPrdCmCd, (case when FdPrdOdNo> 0 then " + _
           "FdPrdOdTc+'/'+FdPrdOdYy+'/'+FdPrdOdChr+'/'+ LTrim(Str(FdPrdOdNo))+" + _
           "'/'+LTrim(Str(FdPrdOdSr)) else '' end) as qPrdOrd, '*' as qExpCmCd, " + _
           "'*' as qExpOrd, max(FdDmCd) as qDmCd, '*' as qPckLst, '*' as qInv, '*' as qExpNo, " + wCommStr
    GrdSumm.ColProp("wSummPrdCmCd").Style = fgcNormal
    GrdSumm.ColProp("wSummPrdOrd").Style = fgcNormal
    GrdSumm.ColProp("wSummExpCmCd").Style = fgcinvisible
    GrdSumm.ColProp("wSummExpOrd").Style = fgcinvisible
    GrdSumm.ColProp("wSummDmCd").Style = fgcNormal
    GrdSumm.ColProp("wSummPl").Style = fgcinvisible
    GrdSumm.ColProp("wSummInv").Style = fgcinvisible
    GrdSumm.ColProp("wSummInvExpNo").Style = fgcinvisible
  ElseIf wGrpBy = "FdExpOdTc,FdExpOdYy,FdExpOdChr,FdExpOdNo" Then
    wStr = "Select '*' as qPrdCmCd, '*' as qPrdOrd, max(FdExpCmCd) as qExpCmCd, " + _
           "(case when FdExpOdNo> 0 then FdExpOdTc+'/'+FdExpOdYy+'/'+FdExpOdChr+'/'+LTrim(Str(FdExpOdNo)) else '' end) " + _
           "as qExpOrd, '*' as qDmCd, '*' as qPckLst, " + _
           "'*' as qInv, '*' as qExpNo, " + wCommStr
    GrdSumm.ColProp("wSummPrdCmCd").Style = fgcinvisible
    GrdSumm.ColProp("wSummPrdOrd").Style = fgcinvisible
    GrdSumm.ColProp("wSummExpCmCd").Style = fgcNormal
    GrdSumm.ColProp("wSummExpOrd").Style = fgcNormal
    GrdSumm.ColProp("wSummDmCd").Style = fgcinvisible
    GrdSumm.ColProp("wSummPl").Style = fgcinvisible
    GrdSumm.ColProp("wSummInv").Style = fgcinvisible
    GrdSumm.ColProp("wSummInvExpNo").Style = fgcinvisible
  ElseIf wGrpBy = "FdExpOdTc,FdExpOdYy,FdExpOdChr,FdExpOdNo,FdExpOdSr" Then
    wStr = "Select '*' as qPrdCmCd, '*' as qPrdOrd, max(FdExpCmCd) as qExpCmCd, " + _
           "(case when FdExpOdNo> 0 then FdExpOdTc+'/'+FdExpOdYy+'/'+FdExpOdChr+'/'+LTrim(Str(FdExpOdNo))+'/'+" + _
           "LTrim(Str(FdExpOdSr)) else '' end) as qExpOrd, max(FdDmCd) as qDmCd, '*' as qPckLst, " + _
           "'*' as qInv, '*' as qExpNo, " + wCommStr
    GrdSumm.ColProp("wSummPrdCmCd").Style = fgcinvisible
    GrdSumm.ColProp("wSummPrdOrd").Style = fgcinvisible
    GrdSumm.ColProp("wSummExpCmCd").Style = fgcNormal
    GrdSumm.ColProp("wSummExpOrd").Style = fgcNormal
    GrdSumm.ColProp("wSummDmCd").Style = fgcNormal
    GrdSumm.ColProp("wSummPl").Style = fgcinvisible
    GrdSumm.ColProp("wSummInv").Style = fgcinvisible
    GrdSumm.ColProp("wSummInvExpNo").Style = fgcinvisible
  ElseIf wGrpBy = "FdDmCd" Then
    wStr = "Select '*' as qPrdCmCd, '*' as qPrdOrd, '*' as qExpCmCd, '*' as qExpOrd, " + _
           "FdDmCd as qDmCd, '*' as qPckLst, '*' as qInv, '*' as qExpNo, " + wCommStr
    GrdSumm.ColProp("wSummPrdCmCd").Style = fgcinvisible
    GrdSumm.ColProp("wSummPrdOrd").Style = fgcinvisible
    GrdSumm.ColProp("wSummExpCmCd").Style = fgcinvisible
    GrdSumm.ColProp("wSummExpOrd").Style = fgcinvisible
    GrdSumm.ColProp("wSummDmCd").Style = fgcNormal
    GrdSumm.ColProp("wSummPl").Style = fgcinvisible
    GrdSumm.ColProp("wSummInv").Style = fgcinvisible
    GrdSumm.ColProp("wSummInvExpNo").Style = fgcinvisible
  ElseIf wGrpBy = "FdPlChr,FdPlYy,FdPlNo" Then
    wStr = "Select '*' as qPrdCmCd, '*' as qPrdOrd, '*' as qExpCmCd, '*' as qExpOrd, " + _
           "'*' as qDmCd, (case when FdPlNo> 0 then FdPlChr+'/'+FdPlYy+'/'+LTrim(Str(FdPlNo)) else '' end) as qPckLst, " + _
           "'*' as qInv, '*' as qExpNo, " + wCommStr
    GrdSumm.ColProp("wSummPrdCmCd").Style = fgcinvisible
    GrdSumm.ColProp("wSummPrdOrd").Style = fgcinvisible
    GrdSumm.ColProp("wSummExpCmCd").Style = fgcinvisible
    GrdSumm.ColProp("wSummExpOrd").Style = fgcinvisible
    GrdSumm.ColProp("wSummDmCd").Style = fgcinvisible
    GrdSumm.ColProp("wSummPl").Style = fgcNormal
    GrdSumm.ColProp("wSummInv").Style = fgcinvisible
    GrdSumm.ColProp("wSummInvExpNo").Style = fgcinvisible
  ElseIf wGrpBy = "FdInTc,FdInYy,FdInChr,FdInNo" Then
    wStr = "Select '*' as qPrdCmCd, '*' as qPrdOrd, '*' as qExpCmCd, '*' as qExpOrd, " + _
           "'*' as qDmCd, '*' as qPckLst, (case when FdInNo> 0 then FdInTc+'/'+FdInYy+'/'" + _
           "+FdInChr+'/'+LTrim(Str(FdInNo)) else '' end) as qInv, " + _
           "max(FdInExpNo) as qExpNo, " + wCommStr
    GrdSumm.ColProp("wSummPrdCmCd").Style = fgcinvisible
    GrdSumm.ColProp("wSummPrdOrd").Style = fgcinvisible
    GrdSumm.ColProp("wSummExpCmCd").Style = fgcinvisible
    GrdSumm.ColProp("wSummExpOrd").Style = fgcinvisible
    GrdSumm.ColProp("wSummDmCd").Style = fgcinvisible
    GrdSumm.ColProp("wSummPl").Style = fgcinvisible
    GrdSumm.ColProp("wSummInv").Style = fgcNormal
    GrdSumm.ColProp("wSummInvExpNo").Style = fgcNormal
  ElseIf wGrpBy = "FdInExpNo" Then
    wStr = "Select '*' as qPrdCmCd, '*' as qPrdOrd, '*' as qExpCmCd, '*' as qExpOrd, " + _
           "'*' as qDmCd, '*' as qPckLst, (case when max(FdInNo)> 0 then " + _
           "max(FdInTc+'/'+FdInYy+'/'+FdInChr+'/'+LTrim(Str(FdInNo))) else '' end) as qInv, " + _
           "FdInExpNo as qExpNo, " + wCommStr
    GrdSumm.ColProp("wSummPrdCmCd").Style = fgcinvisible
    GrdSumm.ColProp("wSummPrdOrd").Style = fgcinvisible
    GrdSumm.ColProp("wSummExpCmCd").Style = fgcinvisible
    GrdSumm.ColProp("wSummExpOrd").Style = fgcinvisible
    GrdSumm.ColProp("wSummDmCd").Style = fgcinvisible
    GrdSumm.ColProp("wSummPl").Style = fgcinvisible
    GrdSumm.ColProp("wSummInv").Style = fgcNormal
    GrdSumm.ColProp("wSummInvExpNo").Style = fgcNormal
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
      .Value(wRow, "wSummBagCount") = wRsSumm!qBags
      .Value(wRow, "wSummQty") = wRsSumm!qBagQty
      .Value(wRow, "wSummWt") = wRsSumm!qRmWt
      .Value(wRow, "wSummBalQty") = wRsSumm!qBalBagQty
      .Value(wRow, "wSummBalWt") = wRsSumm!qBalRmWt
      
      wTotRmWt = wTotRmWt + wRsSumm!qRmWt
      wTotBagQty = wTotBagQty + wRsSumm!qBagQty
      wTotBags = wTotBags + wRsSumm!qBags
      wTotBalRmWt = wTotBalRmWt + wRsSumm!qBalRmWt
      wTotBalBagQty = wTotBalBagQty + wRsSumm!qBalBagQty
      wRsSumm.MoveNext
    Loop
  End With
  ADC("wSummTotWt") = wTotRmWt
  ADC("wSummTotQty") = wTotBagQty
  ADC("wSummTotBags") = wTotBags
  ADC("wSummTotBalQty") = wTotBalBagQty
  ADC("wSummTotBalWt") = wTotBalRmWt
  
  GrdSumm.AllowAdd = False
  GrdSumm.AllowDelete = False
  Set wRsSumm = Nothing
End Sub
Private Sub Summary1(ByVal mGrpByCd As String)
  '*** Summary Group Options :
      'Group By FdPrdCmCd
      'Group By FdExpCmCd
      'Group By FdPrdOdTc,FdPrdOdYy,FdPrdOdChr,FdPrdOdNo
      'Group By FdPrdOdTc,FdPrdOdYy,FdPrdOdChr,FdPrdOdNo,FdPrdOdSr
      'Group By FdExpOdTc,FdExpOdYy,FdExpOdChr,FdExpOdNo
      'Group By FdExpOdTc,FdExpOdYy,FdExpOdChr,FdExpOdNo,FdExpOdSr
      'Group By FdDmCd
      'Group By FdPlChr,FdPlYy,FdPlNo
      'Group By FdInTc,FdInYy,FdInChr,FdInNo
      'Group By FdInExpNo
  '*** Whenever a particular group option is selected, some of the summary grid fields
      'become irrevelent. So these fields should be made invisible.

  Dim wGrpByCd  As String, wGrpBy As String, wCommStr As String
  Dim wTotRmWt As Single, wTotBagQty As Single, wTotBags As Single, wStr As String, wRow As Single
  Dim wTotBalRmWt As Single, wTotBalBagQty As Single
  wGrpByCd = mGrpByCd

  If wGrpByCd = "PRDCMCD" Then wGrpBy = "FdPrdCmCd"
  If wGrpByCd = "EXPCMCD" Then wGrpBy = "FdExpCmCd"
  If wGrpByCd = "PRDORDNO" Then wGrpBy = "FdPrdOdTc,FdPrdOdYy,FdPrdOdChr,FdPrdOdNo"
  If wGrpByCd = "PRDORDSR" Then wGrpBy = "FdPrdOdTc,FdPrdOdYy,FdPrdOdChr,FdPrdOdNo,FdPrdOdSr"
  If wGrpByCd = "EXPORDNO" Then wGrpBy = "FdPrdOdTc,FdPrdOdYy,FdPrdOdChr,FdPrdOdNo"
  If wGrpByCd = "EXPORDSR" Then wGrpBy = "FdPrdOdTc,FdPrdOdYy,FdPrdOdChr,FdPrdOdNo,FdPrdOdSr"
  If wGrpByCd = "DSGCD" Then wGrpBy = "FdDmCd"
  If wGrpByCd = "PCKLST" Then wGrpBy = "FdPlChr,FdPlYy,FdPlNo"
  If wGrpByCd = "INVNO" Then wGrpBy = "FdInTc,FdInYy,FdInChr,FdInNo"
  If wGrpByCd = "INVEXPNO" Then wGrpBy = "FdInExpNo"
  
  wTotRmWt = 0: wTotBagQty = 0: wTotBags = 0: wStr = ""
    
  '****** Sachin 3.02 25-12-07 - Corresponding Partition
  wCommStr = " sum(case when FdBNo<> 0 then 1 else 0 end) as qBags, sum(FdQty) as qBagQty, " + _
             "sum(FdGrWt) as qRmWt, sum(FdQty- FdDespQty) as qBalBagQty, " + _
             "(case when sum(FdQty)> 0 then sum(FdQty- FdDespQty)*sum(FdGrWt)/ " + _
             "sum(FdQty) else 0 end) as qBalRmWt From Fgd " + _
             "join Fg on FgIdNo=FdFgIdNo and FgCoCd= FdCoCd and FgTc=FdTc and FgYy=FdYy and FgChr=FdChr" + _
             " and FgNo=FdNo and FgPrtKey=FdPrtKey Where FdCoCd= '" + ADC("FgCoCd") + "' and FdTc= '" + ADC("FgTc") + "'" + _
             " and FdYy= '" + ADC("FgYy") + "' And " + _
             "FdChr= '" + ADC("FgChr") + "' And FdNo =" + CStr(ADC("FgNo")) + " And FdPrtKey='" + ADC("FgPrtKey") + "' " + _
             " Group By " + wGrpBy
  
  If wGrpBy = "FdPrdCmCd" Then
    wStr = "Select FdPrdCmCd as qPrdCmCd, '*' as qPrdOrd, '*' as qExpCmCd, " + _
           "'*' as qExpOrd, '*' as qDmCd, '*' as qPckLst, '*' as qInv, '*' as qExpNo, " + wCommStr
    GrdSumm.ColProp("wSummPrdCmCd").Style = fgcNormal
    GrdSumm.ColProp("wSummPrdOrd").Style = fgcinvisible
    GrdSumm.ColProp("wSummExpCmCd").Style = fgcinvisible
    GrdSumm.ColProp("wSummExpOrd").Style = fgcinvisible
    GrdSumm.ColProp("wSummDmCd").Style = fgcinvisible
    GrdSumm.ColProp("wSummPl").Style = fgcinvisible
    GrdSumm.ColProp("wSummInv").Style = fgcinvisible
    GrdSumm.ColProp("wSummInvExpNo").Style = fgcinvisible
  ElseIf wGrpBy = "FdExpCmCd" Then
    wStr = "Select '*' as qPrdCmCd, '*' as qPrdOrd, FdExpCmCd as qExpCmCd, " + _
           "'*' as qExpOrd, '*' as qDmCd, '*' as qPckLst, '*' as qInv, '*' as qExpNo, " + wCommStr
    GrdSumm.ColProp("wSummPrdCmCd").Style = fgcinvisible
    GrdSumm.ColProp("wSummPrdOrd").Style = fgcinvisible
    GrdSumm.ColProp("wSummExpCmCd").Style = fgcNormal
    GrdSumm.ColProp("wSummExpOrd").Style = fgcinvisible
    GrdSumm.ColProp("wSummDmCd").Style = fgcinvisible
    GrdSumm.ColProp("wSummPl").Style = fgcinvisible
    GrdSumm.ColProp("wSummInv").Style = fgcinvisible
    GrdSumm.ColProp("wSummInvExpNo").Style = fgcinvisible
  ElseIf wGrpBy = "FdPrdOdTc,FdPrdOdYy,FdPrdOdChr,FdPrdOdNo" Then
    wStr = "Select max(FdPrdCmCd) as qPrdCmCd, (case when FdPrdOdNo> 0 then " + _
           "FdPrdOdTc+'/'+FdPrdOdYy+'/'+FdPrdOdChr+'/'+LTrim(Str(FdPrdOdNo)) else '' end) as qPrdOrd, " + _
           "'*' as qExpCmCd, '*' as qExpOrd, '*' as qDmCd, '*' as qPckLst, '*' as qInv, '*' as qExpNo, " + wCommStr
    GrdSumm.ColProp("wSummPrdCmCd").Style = fgcNormal
    GrdSumm.ColProp("wSummPrdOrd").Style = fgcNormal
    GrdSumm.ColProp("wSummExpCmCd").Style = fgcinvisible
    GrdSumm.ColProp("wSummExpOrd").Style = fgcinvisible
    GrdSumm.ColProp("wSummDmCd").Style = fgcinvisible
    GrdSumm.ColProp("wSummPl").Style = fgcinvisible
    GrdSumm.ColProp("wSummInv").Style = fgcinvisible
    GrdSumm.ColProp("wSummInvExpNo").Style = fgcinvisible
  ElseIf wGrpBy = "FdPrdOdTc,FdPrdOdYy,FdPrdOdChr,FdPrdOdNo,FdPrdOdSr" Then
    wStr = "Select max(FdPrdCmCd) as qPrdCmCd, (case when FdPrdOdNo> 0 then " + _
           "FdPrdOdTc+'/'+FdPrdOdYy+'/'+FdPrdOdChr+'/'+ LTrim(Str(FdPrdOdNo))+" + _
           "'/'+LTrim(Str(FdPrdOdSr)) else '' end) as qPrdOrd, '*' as qExpCmCd, " + _
           "'*' as qExpOrd, max(FdDmCd) as qDmCd, '*' as qPckLst, '*' as qInv, '*' as qExpNo, " + wCommStr
    GrdSumm.ColProp("wSummPrdCmCd").Style = fgcNormal
    GrdSumm.ColProp("wSummPrdOrd").Style = fgcNormal
    GrdSumm.ColProp("wSummExpCmCd").Style = fgcinvisible
    GrdSumm.ColProp("wSummExpOrd").Style = fgcinvisible
    GrdSumm.ColProp("wSummDmCd").Style = fgcNormal
    GrdSumm.ColProp("wSummPl").Style = fgcinvisible
    GrdSumm.ColProp("wSummInv").Style = fgcinvisible
    GrdSumm.ColProp("wSummInvExpNo").Style = fgcinvisible
  ElseIf wGrpBy = "FdExpOdTc,FdExpOdYy,FdExpOdChr,FdExpOdNo" Then
    wStr = "Select '*' as qPrdCmCd, '*' as qPrdOrd, max(FdExpCmCd) as qExpCmCd, " + _
           "(case when FdExpOdNo> 0 then FdExpOdTc+'/'+FdExpOdYy+'/'+FdExpOdChr+'/'+LTrim(Str(FdExpOdNo)) else '' end) " + _
           "as qExpOrd, '*' as qDmCd, '*' as qPckLst, " + _
           "'*' as qInv, '*' as qExpNo, " + wCommStr
    GrdSumm.ColProp("wSummPrdCmCd").Style = fgcinvisible
    GrdSumm.ColProp("wSummPrdOrd").Style = fgcinvisible
    GrdSumm.ColProp("wSummExpCmCd").Style = fgcNormal
    GrdSumm.ColProp("wSummExpOrd").Style = fgcNormal
    GrdSumm.ColProp("wSummDmCd").Style = fgcinvisible
    GrdSumm.ColProp("wSummPl").Style = fgcinvisible
    GrdSumm.ColProp("wSummInv").Style = fgcinvisible
    GrdSumm.ColProp("wSummInvExpNo").Style = fgcinvisible
  ElseIf wGrpBy = "FdExpOdTc,FdExpOdYy,FdExpOdChr,FdExpOdNo,FdExpOdSr" Then
    wStr = "Select '*' as qPrdCmCd, '*' as qPrdOrd, max(FdExpCmCd) as qExpCmCd, " + _
           "(case when FdExpOdNo> 0 then FdExpOdTc+'/'+FdExpOdYy+'/'+FdExpOdChr+'/'+LTrim(Str(FdExpOdNo))+'/'+" + _
           "LTrim(Str(FdExpOdSr)) else '' end) as qExpOrd, max(FdDmCd) as qDmCd, '*' as qPckLst, " + _
           "'*' as qInv, '*' as qExpNo, " + wCommStr
    GrdSumm.ColProp("wSummPrdCmCd").Style = fgcinvisible
    GrdSumm.ColProp("wSummPrdOrd").Style = fgcinvisible
    GrdSumm.ColProp("wSummExpCmCd").Style = fgcNormal
    GrdSumm.ColProp("wSummExpOrd").Style = fgcNormal
    GrdSumm.ColProp("wSummDmCd").Style = fgcNormal
    GrdSumm.ColProp("wSummPl").Style = fgcinvisible
    GrdSumm.ColProp("wSummInv").Style = fgcinvisible
    GrdSumm.ColProp("wSummInvExpNo").Style = fgcinvisible
  ElseIf wGrpBy = "FdDmCd" Then
    wStr = "Select '*' as qPrdCmCd, '*' as qPrdOrd, '*' as qExpCmCd, '*' as qExpOrd, " + _
           "FdDmCd as qDmCd, '*' as qPckLst, '*' as qInv, '*' as qExpNo, " + wCommStr
    GrdSumm.ColProp("wSummPrdCmCd").Style = fgcinvisible
    GrdSumm.ColProp("wSummPrdOrd").Style = fgcinvisible
    GrdSumm.ColProp("wSummExpCmCd").Style = fgcinvisible
    GrdSumm.ColProp("wSummExpOrd").Style = fgcinvisible
    GrdSumm.ColProp("wSummDmCd").Style = fgcNormal
    GrdSumm.ColProp("wSummPl").Style = fgcinvisible
    GrdSumm.ColProp("wSummInv").Style = fgcinvisible
    GrdSumm.ColProp("wSummInvExpNo").Style = fgcinvisible
  ElseIf wGrpBy = "FdPlChr,FdPlYy,FdPlNo" Then
    wStr = "Select '*' as qPrdCmCd, '*' as qPrdOrd, '*' as qExpCmCd, '*' as qExpOrd, " + _
           "'*' as qDmCd, (case when FdPlNo> 0 then FdPlChr+'/'+FdPlYy+'/'+LTrim(Str(FdPlNo)) else '' end) as qPckLst, " + _
           "'*' as qInv, '*' as qExpNo, " + wCommStr
    GrdSumm.ColProp("wSummPrdCmCd").Style = fgcinvisible
    GrdSumm.ColProp("wSummPrdOrd").Style = fgcinvisible
    GrdSumm.ColProp("wSummExpCmCd").Style = fgcinvisible
    GrdSumm.ColProp("wSummExpOrd").Style = fgcinvisible
    GrdSumm.ColProp("wSummDmCd").Style = fgcinvisible
    GrdSumm.ColProp("wSummPl").Style = fgcNormal
    GrdSumm.ColProp("wSummInv").Style = fgcinvisible
    GrdSumm.ColProp("wSummInvExpNo").Style = fgcinvisible
  ElseIf wGrpBy = "FdInTc,FdInYy,FdInChr,FdInNo" Then
    wStr = "Select '*' as qPrdCmCd, '*' as qPrdOrd, '*' as qExpCmCd, '*' as qExpOrd, " + _
           "'*' as qDmCd, '*' as qPckLst, (case when FdInNo> 0 then FdInTc+'/'+FdInYy+'/'" + _
           "+FdInChr+'/'+LTrim(Str(FdInNo)) else '' end) as qInv, " + _
           "max(FdInExpNo) as qExpNo, " + wCommStr
    GrdSumm.ColProp("wSummPrdCmCd").Style = fgcinvisible
    GrdSumm.ColProp("wSummPrdOrd").Style = fgcinvisible
    GrdSumm.ColProp("wSummExpCmCd").Style = fgcinvisible
    GrdSumm.ColProp("wSummExpOrd").Style = fgcinvisible
    GrdSumm.ColProp("wSummDmCd").Style = fgcinvisible
    GrdSumm.ColProp("wSummPl").Style = fgcinvisible
    GrdSumm.ColProp("wSummInv").Style = fgcNormal
    GrdSumm.ColProp("wSummInvExpNo").Style = fgcNormal
  ElseIf wGrpBy = "FdInExpNo" Then
    wStr = "Select '*' as qPrdCmCd, '*' as qPrdOrd, '*' as qExpCmCd, '*' as qExpOrd, " + _
           "'*' as qDmCd, '*' as qPckLst, (case when max(FdInNo)> 0 then " + _
           "max(FdInTc+'/'+FdInYy+'/'+FdInChr+'/'+LTrim(Str(FdInNo))) else '' end) as qInv, " + _
           "FdInExpNo as qExpNo, " + wCommStr
    GrdSumm.ColProp("wSummPrdCmCd").Style = fgcinvisible
    GrdSumm.ColProp("wSummPrdOrd").Style = fgcinvisible
    GrdSumm.ColProp("wSummExpCmCd").Style = fgcinvisible
    GrdSumm.ColProp("wSummExpOrd").Style = fgcinvisible
    GrdSumm.ColProp("wSummDmCd").Style = fgcinvisible
    GrdSumm.ColProp("wSummPl").Style = fgcinvisible
    GrdSumm.ColProp("wSummInv").Style = fgcNormal
    GrdSumm.ColProp("wSummInvExpNo").Style = fgcNormal
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
      .Value(wRow, "wSummBagCount") = wRsSumm!qBags
      .Value(wRow, "wSummQty") = wRsSumm!qBagQty
      .Value(wRow, "wSummWt") = wRsSumm!qRmWt
      .Value(wRow, "wSummBalQty") = wRsSumm!qBalBagQty
      .Value(wRow, "wSummBalWt") = wRsSumm!qBalRmWt
      
      wTotRmWt = wTotRmWt + wRsSumm!qRmWt
      wTotBagQty = wTotBagQty + wRsSumm!qBagQty
      wTotBags = wTotBags + wRsSumm!qBags
      wTotBalRmWt = wTotBalRmWt + wRsSumm!qBalRmWt
      wTotBalBagQty = wTotBalBagQty + wRsSumm!qBalBagQty
      wRsSumm.MoveNext
    Loop
  End With
  ADC("wSummTotWt") = wTotRmWt
  ADC("wSummTotQty") = wTotBagQty
  ADC("wSummTotBags") = wTotBags
  ADC("wSummTotBalQty") = wTotBalBagQty
  ADC("wSummTotBalWt") = wTotBalRmWt
  
  GrdSumm.AllowAdd = False
  GrdSumm.AllowDelete = False
  Set wRsSumm = Nothing
End Sub
Private Sub DispFra(ByVal pv_FgFra As en_FgFra)
  '*** This is the common procedure which is used to toggle between hiding and displaying
      'the frames (i.e. if the frame is visible then make it invisible and viseversa)
      'associated with a command button

  Select Case UCase(pv_FgFra)
  Case Is = ExDet
    If FraEX.Visible = True Then
      FraEX.Visible = False
      FraEX.Enabled = False
      '*** required to pass the array position ExDet-0, to resotre the 'Causes Validation' State
      Call EnaDisaCmds(False, CmdEX, ExDet)
      GrdFgd.SetFocus
    Else
      FraEX.Visible = True
      FraEX.Enabled = True
      If ADC("wExFdExpOdTc").Enabled = True And ADC("wExFdExpOdTc").Visible = True Then
        ADC("wExFdExpOdTc").SetFocus
      Else
        'ADC("wExFdInExpNo").SetFocus
        CmdEX.SetFocus
      End If
      FraEX.ZOrder
      Call EnaDisaCmds(True, CmdEX)
    End If
    
  Case Is = Summ
    If FraSumm.Visible = True Then
      FraSumm.Visible = False
      FraSumm.Enabled = False
      Call EnaDisaCmds(False, CmdSumm, Summ)
      FraSumm.ZOrder
      CmdSumm.SetFocus
    Else
      FraSumm.Visible = True
      FraSumm.Enabled = True
      ADC("wSummBagGrp").SetFocus
      FraSumm.ZOrder
      Call EnaDisaCmds(True, CmdSumm)
    End If
  Case Is = CpyBag
    If FraCB.Visible = True Then
      FraCB.Visible = False
      FraCB.Enabled = False
      Call EnaDisaCmds(False, CmdCB, CpyBag)
      CmdCB.SetFocus
    Else
      FraCB.Visible = True
      FraCB.Enabled = True
      If ADC("wCbBFgSubLoc").Visible = True Then ADC("wCbBFgSubLoc").SetFocus Else ADC("WCBBYYFR").SetFocus   ' ***** Manali 3.8.0 - Copy Bags ib Return Entry
      'adc("WCBBYYFR").SetFocus
      FraCB.ZOrder
      Call EnaDisaCmds(True, CmdCB)
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
      'vk.4 setfocus changed to Bag no./bag alias no. entry text box
'      ADC("wCpyThruBagNo").SetFocus
      ADC("wRfIdSel").SetFocus
      FraCpyRfId.ZOrder
      Call EnaDisaCmds(True, CmdCpyRfId)
    End If
  ' ***** Manali 3.10.0 - 19/04/12 - Cpy Bags from RfId
  End Select
End Sub
Private Sub EnaDisaCmds(ByVal pv_ShowFra As Boolean, Optional ByVal pv_CmdBut As MWCTL_BTN1, Optional ByVal pv_FgFra As en_FgFra)
  '*** This procedure is used to give the toggle option (enable and disable) to
      'the command buttons
  
    '****** storing the state of cmd buttons, when the frame is going to be vis.
    Dim X As Integer
    If pv_ShowFra = True Then
      Call StoreState
    End If
  
  If Not pv_CmdBut Is Nothing Then
    If pv_ShowFra = True Then
      '*** Jenny (Old Code Bef Color)
'        pv_CmdBut.BackColor = vbWhite
'        pv_CmdBut.FontBold = True
      '*** Jenny (Old Code Bef Color)
        If pv_CmdBut.Pic = 0 Then
          pv_CmdBut.BackColor = vbWhite
          'pv_CmdBut.FontBold = True
        End If
        pv_CmdBut.CausesValidation = True
    Else
      '*** Jenny (Old Code Bef Color)
'        pv_CmdBut.BackColor = &H8000000F
'        pv_CmdBut.FontBold = False
      '*** Jenny (Old Code Bef Color)
        If pv_CmdBut.Pic = 0 Then
          pv_CmdBut.BackColor = &H8000000F
          'pv_CmdBut.FontBold = False
        End If
        ' pv_CmdBut.CausesValidation = False
        '*** restoring the 'Causes Validation' state of the current pv_cmdbutton
        If Not IsMissing(pv_FgFra) Then
          pv_CmdBut.CausesValidation = mArr_Cv(pv_FgFra)
        End If
    End If
  End If
  FraNKeyAll.Enabled = Not pv_ShowFra
  ADC.AllowSave = Not pv_ShowFra
  ADC.AllowDelete = Not pv_ShowFra
  
  If pv_ShowFra = True Then
    '** All command buttons are disabled
    CmdCP.Enabled = Not pv_ShowFra
    CmdEX.Enabled = Not pv_ShowFra
    CmdSumm.Enabled = Not pv_ShowFra
    CmdCB.Enabled = Not pv_ShowFra
    CmdCpyRfId.Enabled = Not pv_ShowFra       ' ***** Manali 3.10.0 - 19/04/12 - Cpy Bags from RfId
    CmdInvHd.Enabled = Not pv_ShowFra     'Sv.22
  Else
    '*** Restoring Ena/Disa state of all Command Buttons
    CmdEX.Enabled = mArr_Ed(ExDet)
    CmdCP.Enabled = mArr_Ed(Cpy)
    CmdSumm.Enabled = mArr_Ed(Summ)
    CmdCB.Enabled = mArr_Ed(CpyBag)
    CmdCpyRfId.Enabled = mArr_Ed(CpyRfId)     ' ***** Manali 3.10.0 - 19/04/12 - Cpy Bags from RfId
    CmdInvHd.Enabled = mArr_Ed(InvHd)     'Sv.22
  End If
  
  If Not pv_CmdBut Is Nothing Then pv_CmdBut.Enabled = True
End Sub
Private Sub HideAllFras()
  '*** This procedure hides all the frames other than FraNKeyAll
  FraNKeyAll.Visible = True
  FraEX.Visible = False
  FraEX.Enabled = False
  FraSumm.Visible = False
  FraSumm.Enabled = False
  FraCB.Visible = False
  FraCB.Enabled = False
  FraCpyRfId.Visible = False
  FraCpyRfId.Enabled = False
End Sub

Private Sub StoreState()
  '*** storing the Enabled and CausesValidation state of cmd buttons
  mArr_Cv(ExDet) = CmdEX.CausesValidation
  mArr_Cv(Cpy) = CmdCP.CausesValidation
  mArr_Cv(Summ) = CmdSumm.CausesValidation
  mArr_Cv(CpyBag) = CmdCB.CausesValidation
  mArr_Cv(CpyRfId) = CmdCpyRfId.CausesValidation
  mArr_Cv(InvHd) = CmdInvHd.CausesValidation  'Sv.22
  
  mArr_Ed(ExDet) = CmdEX.Enabled
  mArr_Ed(Cpy) = CmdCP.Enabled
  mArr_Ed(Summ) = CmdSumm.Enabled
  mArr_Ed(CpyBag) = CmdCB.Enabled
  mArr_Ed(CpyRfId) = CmdCpyRfId.Enabled
  mArr_Ed(InvHd) = CmdInvHd.Enabled       'Sv.22
End Sub
Private Function FnRmCtg(ByVal mRmCd As String) As String
    FnRmCtg = Trim(moCn.GetFldVal("Select RmCtg from RmMst where RmCd='" + mRmCd + "'" + IIF(gs_Partition = ctCurrPrtn, " And RmPrtKey='" + ctCurrPrtn + "' ", "")))
End Function
Private Function FnRmSCtg(ByVal mRmCd As String) As String
    FnRmSCtg = Trim(moCn.GetFldVal("Select RmSCtg from RmMst where RmCd='" + mRmCd + "'" + IIF(gs_Partition = ctCurrPrtn, " And RmPrtKey='" + ctCurrPrtn + "' ", "")))
End Function

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
  ws_BtnStr = "CmdCBGo"
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
Private Sub GrdFgd_Validate(Cancel As Boolean)
  Cancel = GrdFgd.Validate
End Sub
Private Sub GrdFgRm_Validate(Cancel As Boolean)
  Cancel = GrdFgRm.Validate
End Sub

Private Sub CmdCB_Click()
  '*** The 'X' record (Txn record) has to be saved before the 'A' records can be entered
  '*** Display or hide the Copy Bag frame depending on the toggle state
  
  If Not ADC.SaveRec Then
    DispMsg "X Save Failed", etInfo
    Exit Sub
  End If
  If ADC("FGFRBLOC") = "" Then DispMsg "Enter Bag From Location.", etError: Exit Sub

  Call DispFra(CpyBag)
  'Uma KBS
  ADC("WCBFGPLYY") = ADC("WFGPLYY")
  ADC("WCBFGPLCHR") = ADC("WFGPLCHR")
  ADC("WCBFGPLNO") = ADC("WFGPLNO")


End Sub

Private Sub CmdCBGo_Click()
  '*** Call the Sub Procedure CopyBag
' ########################  Manoj #### Ver: 2.0.6 #### Date: 25/12/2003  ########################
  If mf_DtLocked = True Then
' ###########################################  Manoj  ###########################################
    DispMsg "Cannot Copy Bags For Locked Entry", etError
  Else
    If mf_ValidLocSeq = False Then
      DispMsg "The User Does Not Have Access To One Of The Locs", etError
    Else
      Call CopyBag
    End If
  End If
  Call CmdCB_Click
End Sub

Private Sub CopyBag(Optional ByVal pb_CpyRfId As Boolean = False)
  '*** Procedure to Copy Bags from one Loc to another in case of 'BV' entry
  ' ***** Manali 3.8.0 - Copy Bag Option In FWB Entry - To Bag Loc Check added
  Dim wRsCB As MwfLib.MDORowSet, wTotRows As Single, wRow As Single, wCnd As String
  Dim wsCmCd As String, wsBag As String
  Dim wMinSr As Integer
  Dim wRfIdStr As String, wRfIdNoArr() As String, wRfIdNo As String, i As Integer, wTotRfid As Integer
  Dim ws_LocStr As String    '*****Bhavna -for multiple bag locns
  '578-24 packing list no in the header will be a filter for FWB
  Dim wPlNoCnd As String
  If ms_TcTyp = "FWB" And ADC("wFgPlYy") <> "" And ADC("wFgPlChr") <> "" And ADC("wFgPlNo") <> 0 Then
    wPlNoCnd = " and exists (Select 'x' from Fgd where FdTc ='FB' and FdCoCd =BCoCd and " + _
               " FdBYy =BYy and FdBChr =BChr and FdBNo =BNo and FdPlYy ='" + ADC("wFgPlYy") + "'" + _
               " and FdPlChr ='" + ADC("wFgPlChr") + "' and FdPlNo =" + CStr(ADC("wFgPlNo")) + " ) "
  End If
  GrdFgd.Col = 0
  ' ***** Manali 3.10.0 - 19/04/12 - Cpy Bags from RfId
  If Not pb_CpyRfId Then
    If ms_TcTyp = "FB" Then
      If ADC("FgFrBLoc") = "" Then DispMsg "Enter From Bag Location", etError: Exit Sub
      If InStr(1, "," + ADC("FgFrBLoc") + ",", ",PFG,") > 0 Then DispMsg "From Bag Location Cannot Be 'PFG'", etError: Exit Sub
'      If adc("FgFrBLoc") = "PFG" Then DispMsg "From Bag Location Cannot Be 'PFG'", etError: Exit Sub
    ElseIf ms_TcTyp = "FWB" Then
        If ADC("wCbBFgSubLoc") <> "" Then
          If Not moCn.RecSeek("select 'x' From vParam where vPTyp='FgSubLoc' and vPCoCd= '" + gs_CoCd + "' and vPMCd='" + ADC("wCbBFgSubLoc") + "'" + _
                              " and (vPValue4='' Or ','+vPValue4+',' Like '%," + gs_UsrCd + ",%')") Then _
          DispMsg "Bag Sub Loc is Not Accessible to the User ", etError: Exit Sub
        End If
        
        
      If ADC("FgToBLoc") = "PFG" Then DispMsg "From Bag Location Cannot Not Be 'PFG'", etError: Exit Sub
      If ADC("FgToBLoc") = "" Then DispMsg "Enter To Bag Location", etError: Exit Sub
      ADC("wCbBFgSubLoc").CmpStr = "BFgSubLoc="
    End If
    
    ADC("wCbBYyFr").CmpStr = "BYy>=": ADC("wCbBChrFr").CmpStr = "BChr>=": ADC("wCbBNoFr").CmpStr = "BNo>="
    ADC("wCbBYyTo").CmpStr = "BYy<=": ADC("wCbBChrTo").CmpStr = "BChr<=": ADC("wCbBNoTo").CmpStr = "BNo<="
    ADC("wCbOdTcFr").CmpStr = "BOdTc>=": ADC("wCbOdYyFr").CmpStr = "BOdYy>=": ADC("wCbOdChrFr").CmpStr = "BOdChr>=": ADC("wCbOdNoFr").CmpStr = "BOdNo>=": ADC("wCbOdSrFr").CmpStr = "BOdSr>="
    ADC("wCbOdTcTo").CmpStr = "BOdTc<=": ADC("wCbOdYyTo").CmpStr = "BOdYy<=": ADC("wCbOdChrTo").CmpStr = "BOdChr<=": ADC("wCbOdNoTo").CmpStr = "BOdNo<=": ADC("wCbOdSrTo").CmpStr = "BOdSr<="
    ADC("wCbDmCdFr").CmpStr = "BOdDmCd>=": ADC("wCbDmCdTo").CmpStr = "BOdDmCd<="
    ADC("wCbCustCdFr").CmpStr = "OmCmCd>=": ADC("wCbCustCdTo").CmpStr = "OmCmCd<="
  
    wCnd = ADC.RepCond
    wCnd = IIF(wCnd <> "", " and " + wCnd, "")

   ' **** Zubin 211 **** '
    '''''  Set wRsCB = moCn.OpenRes("Select * from Bag where BCoCd= '" + ADC("FgCoCd") + "' and BCls= 'N' and " + _
                               "BLoc= '" + ADC("FgFrBLoc") + "' and not exists (Select 'x' from Txnd where " + _
                               "TdCoCd= '" + ADC("FgCoCd") + "' and TdYy= BLstYy and TdKey= BLstKey " + _
                               "and (Select PValue from Param where PTyp= 'TC' and " + _
                               "PMCd= TdTc) in('REJ', 'MLT')) And BLstYyKey <> '' " + wCnd + " Order By BYy, BChr, BNo")
    '*** Jay 2.13(CT) *** (PValue = 'P')
    '*                             "BLoc= '" + ADC("FgFrBLoc") + "'
    '* "(InStr(1,'," + ADC("FgFrBLoc") + ",','+BLoc+',') <> 0) and " + _

    '**************Bhavna added multiple bag locns
          ws_LocStr = "CHARINDEX(','+BLoc+',' , '," + ADC("FgFrBLoc") + ",') <> 0 "
    '**************Bhavna added multiple bag locns and custcd in scope
   '****** Sachin 3.02 25-12-07 - Current Partition
   '578-24 wPlNoCnd added
    Set wRsCB = moCn.OpenRes("Select Bag.*, OmCmCd from Bag " + _
                             "Join OrdMst on OmCoCd = '" + ADC("FgCoCd") + "' And " + _
                             " OmTc = BOdTc And OmYy = BOdYy And OmChr = BOdChr And OmNo = BOdNo And OmPrtKey='" + ctCurrPrtn + "' " + _
                             "where BCoCd = '" + ADC("FgCoCd") + "' and BCls= 'N' and " + _
                             ws_LocStr + " and " + _
                             "BPrtKey='" + ctCurrPrtn + "' and not exists (Select 'x' from Txnd where " + _
                             "TdCoCd= '" + ADC("FgCoCd") + "' and TdYy= BLstYy and TdKey= BLstKey and TdPrtKey=BPrtKey " + _
                             "and (Select PValue from Param where PTyp= 'TC' and " + _
                             "PMCd= TdTc) in('REJ', 'MLT')) And BLstYyKey <> '' " + _
                             "and (Select PValue from Param where PTyp= 'BCHR' and PMCd= BChr and PSCd= '')= 'P' " + _
                             wPlNoCnd + wCnd + " Order By BYy, BChr, BNo")
  Else
    ' ***** Manali 3.10.0 - 19/04/12 - Cpy Bags from RfId
    'vk.4 RfId changed as BagNo. / Bag Alias No.
    If ADC("wRfIdSel") = "" Then DispMsg "Enter Bag No. / Bag Alias No. List", etError: Exit Sub
    'vk.4 calling function to replace enter with ,
    wRfIdStr = PrepareRfIdString(ADC("wRfIdSel"))
       
    
    Dim wInvalRfIdStr As String
    If ADC("wRfIdSel") <> "" Then
      wRfIdStr = ParseBagNoBarCd(wRfIdStr)
      wRfIdNoArr = Split(wRfIdStr, ",")
      wInvalRfIdStr = ""
      If Left(wRfIdStr, 1) = "," Or Right(wRfIdStr, 1) = "," Then
          'vk.4 Tag Input changed as BagNo. / Bag Alias No.
          DispMsg "Invalid BagNo. / Bag Alias No. ", etError: Exit Sub
      End If
      wRfIdStr = ""
      wTotRfid = UBound(wRfIdNoArr) + 1
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
      'vk.4 Tag changed as Bag Alias No.
      If wInvalRfIdStr <> "" Then DispMsg "Bag Alias No. / Bag Not Defined:" + wInvalRfIdStr, etError: Exit Sub
      
      '**************Bhavna added multiple bag locns and cust cd
        ws_LocStr = "CHARINDEX(','+BLoc+',' , '," + ADC("FgFrBLoc") + ",') <> 0 "
        
    '**************Bhavna added multiple bag locns
    '**************Bhavna added cpy thru bag no
      Dim wSelCnd As String
      Dim ws_RfIdJoin As String

      If ADC("WCPYTHRUBAGNO") = "N" Then
        ws_RfIdJoin = " Join RfIdMst On RimBCoCd=BCoCd and RimBYy=BYy and RimBChr=BChr and RimBNo=BNo "
        wSelCnd = " and RimNo In (" + wRfIdStr + ") "
      Else
          ws_RfIdJoin = " "
          wSelCnd = " and  charindex(','+BYy + '/' + BChr + '/' + LTrim(RTrim(str(BNo)))+',' , '," + wRfIdStr + ",') <> 0 "
      End If

'      Set wRsCB = moCn.OpenRes("Select Bag.*, OmCmCd from Bag " + _
                                "Join OrdMst on OmCoCd = '" + ADC("FgCoCd") + "' And " + _
                                " OmTc = BOdTc And OmYy = BOdYy And OmChr = BOdChr And OmNo = BOdNo And OmPrtKey='" + ctCurrPrtn + "' " + _
                                "Join RfIdMst On RimBCoCd=BCoCd and RimBYy=BYy and RimBChr=BChr and RimBNo=BNo " + _
                               "where BCoCd= '" + ADC("FgCoCd") + "' and BCls= 'N' and " + _
                               IIF(mb_EnfFrBLocInFg, ws_LocStr + " and ", "") + _
                               " BLoc " + IIF(ms_TcTyp = "FB", "<>'" + ctFgLoc + "'", "='" + ctFgLoc + "'") + " and  BPrtKey='" + ctCurrPrtn + "' and not exists (Select 'x' from Txnd where " + _
                               "TdCoCd= '" + ADC("FgCoCd") + "' and TdYy= BLstYy and TdKey= BLstKey and TdPrtKey=BPrtKey " + _
                               "and (Select PValue from Param where PTyp= 'TC' and " + _
                               "PMCd= TdTc) in('REJ', 'MLT')) And BLstYyKey <> '' " + _
                               "and (Select PValue from Param where PTyp= 'BCHR' and PMCd= BChr and PSCd= '')= 'P' " + _
                               "and RimNo in (" + wRfIdStr + ") Order By BYy, BChr, BNo")
       '578-24 wPlNoCnd added
       Set wRsCB = moCn.OpenRes("Select Bag.*, OmCmCd from Bag " + _
                                "Join OrdMst on OmCoCd = '" + ADC("FgCoCd") + "' And " + _
                                " OmTc = BOdTc And OmYy = BOdYy And OmChr = BOdChr And OmNo = BOdNo And OmPrtKey='" + ctCurrPrtn + "' " + _
                                ws_RfIdJoin + _
                               "where BCoCd= '" + ADC("FgCoCd") + "' and BCls= 'N' and " + _
                               IIF(mb_EnfFrBLocInFg, ws_LocStr + " and ", "") + _
                               " BLoc " + IIF(ms_TcTyp = "FB", "<>'" + ctFgLoc + "'", "='" + ctFgLoc + "'") + " and  BPrtKey='" + ctCurrPrtn + "' and not exists (Select 'x' from Txnd where " + _
                               "TdCoCd= '" + ADC("FgCoCd") + "' and TdYy= BLstYy and TdKey= BLstKey and TdPrtKey=BPrtKey " + _
                               "and (Select PValue from Param where PTyp= 'TC' and " + _
                               "PMCd= TdTc) in('REJ', 'MLT')) And BLstYyKey <> '' " + _
                               "and (Select PValue from Param where PTyp= 'BCHR' and PMCd= BChr and PSCd= '')= 'P' " + _
                               wPlNoCnd + wSelCnd + " Order By BYy, BChr, BNo")
    End If
    ' ***** Manali 3.10.0 - 19/04/12 - Cpy Bags from RfId
  End If
'*** Jay 2.13(CT) ***
' **** Zubin 211 **** '
 
  wTotRows = wRsCB.RecCount
  If wTotRows <= 0 Then DispMsg "No Bags Found for Copy in given Scope.", etError: GoTo CopyEnd
  GrdFgd.StartCopy
  mf_CpyStartedByParent = True
  
  MWLib.BeginProcess Me, "Copying Bag Details ..."
  BeginEmrProcess Me, "Copying Bag Details..."
  With GrdFgd
   wMinSr = .MaxVal("FdSr") + 1
    Do While Not (wRsCB.EOF Or wRsCB.BOF)
      '****** Sachin 3.02 25-12-07 - Current Partition
'      wsCmCd = moCn.GetFldVal("Select OmCmCd From OrdMst Where OmCoCd = '" & ADC("FgCoCd") & "' And" & _
              " OmTc = '" & wRsCB!BOdTc & "' And OmYy = '" & wRsCB!BOdYy & "' And OmChr = '" & wRsCB!BOdChr & "' And OmNo =" & CStr(wRsCB!BOdNo) & " And OmPrtKey='" + ctCurrPrtn + "' ")
      .AddItem
      wRow = .Rows - 1
      UpdateEmrProcess "Copying Bag Details..." + CStr(wRow)
      .Row = wRow
'     .Value(wRow, "FdSr") = .MaxVal("FdSr") + 1
      .Value(wRow, "FdBYy") = wRsCB!BYy
      .Value(wRow, "FdBChr") = wRsCB!BChr
      .Value(wRow, "FdBNo") = wRsCB!BNo
      .Value(wRow, "FdDmCd") = wRsCB!BOdDmCd
      .Value(wRow, "FdSfx") = wRsCB!BOdSfx
      .Value(wRow, "FdDmSz") = wRsCB!BOdDmSz
      .Value(wRow, "FdQty") = wRsCB!BQty
      .Value(wRow, "FdGrWt") = wRsCB!BGrWt
      .Value(wRow, "FdPrdOdTc") = wRsCB!BOdTc
      .Value(wRow, "FdPrdOdYy") = wRsCB!BOdYy
      .Value(wRow, "FdPrdOdChr") = wRsCB!BOdChr
      .Value(wRow, "FdPrdOdNo") = wRsCB!BOdNo
      .Value(wRow, "FdPrdOdSr") = wRsCB!BOdSr
      .Value(wRow, "FdPrdCmCd") = wRsCB!OmCmCd
      .Value(wRow, "FdExpCmCd") = wRsCB!OmCmCd
      .Value(wRow, "FdExpOdTc") = wRsCB!BOdTc
      .Value(wRow, "FdExpOdYy") = wRsCB!BOdYy
      .Value(wRow, "FdExpOdChr") = wRsCB!BOdChr
      .Value(wRow, "FdExpOdNo") = wRsCB!BOdNo
      .Value(wRow, "FdExpOdSr") = wRsCB!BOdSr
      'MW.12- If the entry is FWB then no need to write Pl no to record
      If ADC.MenuCd <> "FWB" Then
        .Value(wRow, "FdPlYy") = ADC("WCBFGPLYY")
        .Value(wRow, "FdPlChr") = ADC("WCBFGPLCHR")
        .Value(wRow, "FdPlNo") = ADC("WCBFGPLNO")
      End If
      .Value(wRow, "FdInTc") = ""
      .Value(wRow, "FdInYy") = ""
      .Value(wRow, "FdInChr") = ""
      .Value(wRow, "FdInNo") = 0
      .Value(wRow, "FdInExpNo") = ""
      .Value(wRow, "FdDespQty") = 0
      .Value(wRow, "FdBYyKey") = wRsCB!BLstYyKey
      .Value(wRow, "FdSubLoc") = ADC("FgSubLoc")
      .Value(wRow, "FdCTBDespQty") = 0
      If .SaveRec(wRow) = True Then
        Call CpyRm
      Else
        wsBag = wsBag & IIF(wsBag = "", "", ", ") & wRsCB!BYy & "/" & wRsCB!BChr & "/" & CStr(wRsCB!BNo)
' **** Zubin 211 (if first rec does not get saved, then .Row = .Row - 1 will give error) **** '
'''''        .Row = .Row - 1
        If .Row > 0 Then .Row = .Row - 1
' **** Zubin 211 (if first rec does not get saved, then .Row = .Row - 1 will give error) **** '
      End If
      wRsCB.MoveNext
    Loop
  mf_CpyStartedByParent = False
    
  MWLib.EndProcess Me
  EndEmrProcess Me
  GrdFgd.EndCopy
  '*** End of the copy process
   If pb_CpyRfId And ADC("WCPYTHRUBAGNO") = "N" Then
        Dim wRsRfId As MwfLib.MDORowSet
        'vk.4 ADC("wRfIdSel") replaced with wRfIdStr
        Set wRsRfId = moCn.OpenRes("Select RimNo from RfIdMst where Not Exists(Select 'x' From Fgd where " + _
                                        " FdCoCd=RimBCoCd and FdBYy=RimBYY and FdBChr=RimBChr and FdBNo=RimBNo and " + _
                                        " FdCoCd='" + ADC("FgCoCd") + "' and FdTc='" + ADC("FgTc") + "' and FdYy='" + ADC("FgYy") + "' and FdChr='" + ADC("FgChr") + "' and FdNo=" + CStr(ADC("FgNo")) + _
                                        " and  FdSr>=" + CStr(wMinSr) + " and FdSr<=" + CStr(.MaxVal("FdSr")) + ")" + _
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
    If wsBag <> "" Then DispMsg "Bags " & wsBag & " are not Copied.", etWarning
CopyEnd:
  Set wRsCB = Nothing
  mf_CpyStartedByParent = False
  
End Sub

'*** Jay 2.14 *** [Over Rule Tolerance Check By Providing SEO PWD]
Private Sub TxtPwd_GotFocus()
  DispMsg "Enter SEO Password", etInfo
End Sub
'*** Jay 2.14 *** [Over Rule Tolerance Check By Providing SEO PWD]
Private Sub TxtPwd_LostFocus()
Call TxtPwd_Validate(False)
End Sub
Private Sub RMSummary(ByVal mGrpByCd As String, ByVal mVchYN As String)
  
  '***** Sachin 3.11.0 - RM Summary in Invoice Bags entry
  Dim wGrpByCd  As String, wGrpBy As String, wCommStr As String
  Dim wRmTotWt As Single, wRmTotQty As Single, wStr As String, wRow As Single
  Dim wVchSummJoin As String, wSummCnd As String, wCurrRow As Integer
  
  If GrdFgd.Rows <= 1 Then
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
  wCurrRow = IIF(GrdFgd.Row = 0, 1, GrdFgd.Row)
  
  If mVchYN = "N" Then
    wVchSummJoin = ""
    wSummCnd = " FdCoCd ='" + gs_CoCd + "' and  FdBYy ='" + GrdFgd.Value(wCurrRow, "FdBYy") + "' " + _
             " And  FdBChr ='" + GrdFgd.Value(wCurrRow, "FdBChr") + "' And  " + _
             " FdBNo =" + CStr(GrdFgd.Value(wCurrRow, "FdBNo")) + " And FdPrtKey=(Case When '" + gs_Partition + "'='" + ctCurrPrtn + "' Then '" + ctCurrPrtn + "' Else FdPrtKey End)"
    LblMain(2).Caption = "Bag Level Summary for " + GrdFgd.Value(wCurrRow, "FdBYy") + "/" + GrdFgd.Value(wCurrRow, "FdBChr") + "/" + CStr(GrdFgd.Value(wCurrRow, "FdBNo"))
  Else
   ' wVchSummJoin = " Join Fg On FgCoCd=FdCoCd And IfFdBYy=FdBYy And IfFdBChr=FdBChr And IfFdBNo=FdBNo "
    wSummCnd = " FdCoCd='" + gs_CoCd + "' and FdTc='" + ADC("FgTc") + "' and " + _
               " FdYy='" + ADC("FgYy") + "' and FdChr='" + ADC("FgChr") + "' and " + _
               " FdNo= " + CStr(ADC("FgNo")) + " and FdPrtKey=(Case When '" + gs_Partition + "'='" + ctCurrPrtn + "' Then '" + ctCurrPrtn + "' Else FdPrtKey End)"
  End If
  
  wRmTotWt = 0: wRmTotQty = 0: wStr = ""
  wCommStr = " Sum(FrRmQty) as RmQty, Sum(FrRmWt*(Case When FdTc='FB' Then 1 Else -1 End)) as RmWt from Fgrm " + _
             "    Join RmMst On RmCd= FrRmCd " + _
             "    Join Fgd On FdCoCd=FrCoCd And FdTc=FrTc And FdYy=FrYy And FdChr=FrChr And FdNo=FrNo And FdSr=FrSr " + wVchSummJoin + _
             " Where " + wSummCnd + _
             " Group By " + wGrpBy + " " + _
             " Order by RmCtg, RmSCtg, RmCd "
  
  If wGrpBy = "RmCtg" Then
    wStr = " Select RmCtg, '*' as RmSCtg, '*' as RmCd, " + wCommStr
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
'*** Jay 2.14 *** [Over Rule Tolerance Check By Providing SEO PWD]
Private Sub TxtPwd_Validate(Cancel As Boolean)

'mf_DtLocked = IIF((adc("FgDt") <> "" And adc("FgDt") < CDate(ms_LockDt)) _
'                                Or _
'                   (adc("FgDt") <> "" And adc("FgDt") > CDate(ms_FutureDt)) _
'                                Or _
'                  (adc("FgDt") <= CDate(ms_PurgeDt)), _
'                True, False)

Dim ms_PwdFlag  As String
ADC("FgSeoPwd") = ""
  ' ***** Manali 3.6.0 - 12/11/09 - ADC.AddFndMode condition removed from here and added for ADC.SaveRec
  'If ADC.AddFndMode = xFndMode Then
    If Trim(TxtPwd.text) <> "" Then
      If Trim(TxtPwd.text) = moCn.GetFldVal("select IsNull ((select PValue from Param Join UsrAccess on UaUsrCd = PMCd " + _
                   "where PTyp= 'USR' and PMCd= '" + gs_UsrCd + "' and UaMnuCd = '" + ADC.MenuCd + "' and " + _
                   "UaSeoYn= 'Y' and PValue= '" + Trim(TxtPwd.text) + "'), " + _
                   "(Select PValue1 from Param where Ptyp= 'SEOPWD' and PmCd= 'SEOPWD' )) ") Then
          ADC("FgSeoPwd") = TxtPwd.text
      Else
          Cancel = True: DispMsg "Invalid Edit Password / User Does Not Have SEO Rights ", etError
      End If
    Else
      ms_PwdFlag = moCn.RecSeek("Select PValue1 from Param where PTyp= 'SEOPWD' and PmCd= 'SEOPWD' and PValue1= '' ")
      If ms_PwdFlag = True Then ADC("FgSeoPwd") = TxtPwd.text
    End If
  'End If
  If ADC.AddFndMode = xFndMode Then ADC.SaveRec
End Sub

'Function CheckTolErrMsg(ByVal pBYy As String, ByVal pBChr As String, ByVal pBNo As Single, ByVal pTdSr As Integer, _
'                        ByVal pTdSrNo As Integer, ByVal pAddWt As Single, ByVal pRmCtg As String, ByVal pRmSCtg As String, _
'                        ByVal pQWFlag As String) As String
'
'     Dim rChkTol As MwfLib.MDORowSet, wRrMinTol As Single, wRrMaxTol As Single
'     Dim wOrPrdWt As Single, wMinTolDesc As String, wMaxTolDesc As String
'     Dim wOrPrdQty As Single
'
'    If UCase(pQWFlag) = UCase("W") Then
'      CheckTolErrMsg = ""
'      Set rChkTol = moCn.OpenResultset(" Select Top 1 qOrPrdWt, " + _
'                 " (Case when OdMinDiaTolWt<>0 and '" + Trim(pRmSCtg) + "'='' Then OdMinDiaTolWt Else (qOrPrdWt-(Case when (OdIWtEqOrd='Y' or OdFixPrc='Y' Or (OdIWtEqOrd='N' and ','+PDesc225+',' Like '%,D,%')) then RrFixMinTol Else RrOpnMinTol End)) End) as qRrMinTol, " + _
'                 " (Case when OdMaxDiaTolWt<>0 and '" + Trim(pRmSCtg) + "'='' Then OdMaxDiaTolWt Else (qOrPrdWt+(Case when (OdIWtEqOrd='Y' or OdFixPrc='Y' Or (OdIWtEqOrd='N' and ','+PDesc225+',' Like '%,D,%')) then RrFixMaxTol Else RrOpnMaxTol End)) End) as qRrMaxTol, " + _
'                 " (Case when OdMinDiaTolWt<>0 and '" + Trim(pRmSCtg) + "'='' Then 'Order' ELSE 'WTT' End) qMinDesc, " + _
'                 " (Case when OdMaxDiaTolWt<>0 and '" + Trim(pRmSCtg) + "'='' Then 'Order' ELSE 'WTT' End) qMaxDesc " + _
'                 " /* (Case when (OdIWtEqOrd='Y' or OdFixPrc='Y' Or (OdIWtEqOrd='N' and ','+PDesc225+',' Like '%,D,%')) Then RrFixMinTol Else RrOpnMinTol End) as qRrMinTol, " + _
'                 "    (Case when (OdIWtEqOrd='Y' or OdFixPrc='Y' Or (OdIWtEqOrd='N' and ','+PDesc225+',' Like '%,D,%')) Then RrFixMaxTol Else RrOpnMaxTol End) as qRrMaxTol */" + _
'                 " From Bag " + _
'                 "      Join OrdDsg On OdPrtKey=BPrtKey and OdIdNo=BOdIdNo and OdCoCd =BCoCd and OdTc=BOdTc and OdYy=BOdYy and OdChr=BOdChr and OdNo=BOdNo and OdSr=BOdSr " + _
'                 "      Join (Select  OrCoCd,OrTc,OrYy,OrChr,OrNo,OrSr" + IIF(pRmSCtg <> "", ",OrRmCtg,OrRmSCtg", "") + ",sum(OrPrdWt) qOrPrdWt " + _
'                 "               From OrdRm " + _
'                 "            Where OrRmCtg='" + pRmCtg + "' " + _
'                 "            Group By OrCoCd,OrTc,OrYy,OrChr,OrNo,OrSr" + IIF(pRmSCtg <> "", ",OrRmCtg,OrRmSCtg", "") + ") OrRm " + _
'                 "                     On OrCoCd=OdCoCd and OrTc=OdTc and OrYy=OdYy and OrChr=OdChr and OrNo=OdNo and OrSr=OdSr " + IIF(pRmSCtg <> "", "  and  OrRmCtg='" + pRmCtg + "'  and OrRmSCtg='" + pRmSCtg + "' ", "") + _
'                 "      Join RmRt On RrTcTyp='WTT' and RrCmCtg='C' " + IIF(pRmSCtg <> "", " and RrCtg=OrRmCtg and RrSCtg=OrRmSCtg ", "") + _
'                 "                     And round(qOrPrdWt,3) between round(RrFrLn,3) and round(RrToLn,3) " + _
'                 "                     And RrCmCd In (OdOmCmCd,'ZSELF') And RrDmCd In (OdDmCd,'') " + _
'                 "      Left Outer Join Param On  PTyp='IWTFRORD' and PMCd=OdIWtFrOrd " + _
'                 " Where " + _
'                 "      BPrtKey='" + ctCurrPrtn + "' and BCoCd= '" + ADC("TCoCd") + "' and BYy= '" + pBYy + "' and BChr= '" + pBChr + "'  and BNo= " + CStr(pBNo) + _
'                 "      and (select vPValue3 from vParam where vPTyp='CHR' and vPCoCd=OdCoCd and vPMCd=OdTc and vPSCd=OdChr)= 'N' " + _
'                 " Order By (Case when RrCmCd=OdOmCmCd Then 1 Else 2 End), (Case when RrDmCd=OdDmCd Then 1 Else 2 End)")
'
'      If Not (rChkTol.EOF Or rChkTol.BOF) Then
'        wOrPrdWt = rChkTol!qOrPrdWt
'        wRrMinTol = IIF(IsNull(rChkTol!qRrMinTol), -1, rChkTol!qRrMinTol)
'        wRrMaxTol = IIF(IsNull(rChkTol!qRrMaxTol), -1, rChkTol!qRrMaxTol)
'        wMinTolDesc = rChkTol!qMinDesc
'        wMaxTolDesc = rChkTol!qMaxDesc
'        If moCn.RecSeek(" Select 'x' From Bag " + _
'                       "  Left Outer Join (Select  TdCoCd as qTdCoCd,TdBYy as qTdBYy,TdBChr as qTdBChr,TdBNo as qTdBNo," + IIF(pRmSCtg <> "", "TdRmCtg as qTdRmCtg,TdRmSCtg as qTdRmSCtg, ", "") + _
'                       "  sum(TdRmWt* (case when TdFrRmLoc= 'W' and TdFrRmDC= 'C' then -1 " + _
'                       "                    when TdFrRmLoc= 'W' and TdFrRmDC= 'D' then 1 " + _
'                       "                    when TdToRmLoc= 'W' and TdToRmDC= 'C' then -1 " + _
'                       "                    when TdToRmLoc= 'W' and TdToRmDC= 'D' then 1 else 0 end)) qTdRmWt " + _
'                       "  From Txnd where TdRmCtg='" + pRmCtg + "' " + _
'                       "       And Not (TdPrtKey='" + ctCurrPrtn + "' And TdCoCd='" + ADC("TCoCd") + "' And TdTc='" + ADC("TTc") + "' And TdYy='" + ADC("TYy") + "' And TdChr='" + ADC("TChr") + "' and TdNo=" + CStr(ADC("TNo")) + " And TdSr=" + CStr(pTdSr) + " And TdSrNo=" + CStr(pTdSrNo) + ")" + _
'                       "  Group By TdCoCd,TdBYy,TdBChr,TdBNo" + IIF(pRmSCtg <> "", ",TdRmCtg,TdRmSCtg", "") + ") TdRm " + _
'                       "           On bCoCd=qTdCoCd and BYy=qTdBYy and BChr=qTdBChr and BNo=qTdBNo " + IIF(pRmSCtg <> "", " and qTdRmCtg='" + pRmCtg + "' and qTdRmSCtg='" + pRmSCtg + "' ", "") + _
'                       " Where " + _
'                       "      BPrtKey='" + ctCurrPrtn + "' and BCoCd= '" + ADC("TCoCd") + "' and BYy= '" + pBYy + "' and BChr= '" + pBChr + "'  and BNo= " + CStr(pBNo) + _
'                       "      and Cast(BQty as Decimal(8,2))>0   and IsNull(" + IIF(UCase$(ADC.MenuCd) = "BV", CStr(wRrMinTol), CStr(wRrMaxTol)) + ",-1)>-1 " + _
'                       "             and convert(decimal(10,3),((IsNull(qTdRmWt,0)+" + IIF(UCase$(ADC.MenuCd) = "BV", "0", CStr(pAddWt)) + ")/BQty)) " + IIF(UCase$(ADC.MenuCd) = "BV", " < ", " > ") + " convert(decimal(10,3)," + IIF(UCase$(ADC.MenuCd) = "BV", CStr(wRrMinTol), CStr(wRrMaxTol)) + ") ") Then
'
'           CheckTolErrMsg = "Diamond " + IIF(pRmSCtg <> "", "Sub-Ctg", "Ctg") + " Wt " + IIF(UCase$(ADC.MenuCd) = "BV", " is Less Than the Minimum ", " Exceeds the Maximum") + " Tolerance Wt [ in " + IIF(UCase$(ADC.MenuCd) = "BV", wMinTolDesc, wMaxTolDesc) + "], Only Authorised Users Can do this"
'           Set rChkTol = Nothing
'           Exit Function
'         End If
'      End If
'      Set rChkTol = Nothing
'      '****** Sachin 3.5.0 - Tolerance ******
'
'    ElseIf UCase(pQWFlag) = UCase("Q") Then
'      CheckTolErrMsg = ""
'      Set rChkTol = moCn.OpenResultset(" Select Top 1 qOrPrdQty, " + _
'                 " qOrPrdQty-(Case when (OdIWtEqOrd='Y' or OdFixPrc='Y' Or (OdIWtEqOrd='N' and ','+PDesc225+',' Like '%,D,%')) then RrFixMinTol Else RrOpnMinTol End) as qRrMinTol, " + _
'                 " qOrPrdQty+(Case when (OdIWtEqOrd='Y' or OdFixPrc='Y' Or (OdIWtEqOrd='N' and ','+PDesc225+',' Like '%,D,%')) then RrFixMaxTol Else RrOpnMaxTol End) as qRrMaxTol, " + _
'                 " 'QTT' qMinDesc, " + _
'                 " 'QTT' qMaxDesc " + _
'                 " From Bag " + _
'                 "      Join OrdDsg On OdPrtKey=BPrtKey and OdIdNo=BOdIdNo and OdCoCd =BCoCd and OdTc=BOdTc and OdYy=BOdYy and OdChr=BOdChr and OdNo=BOdNo and OdSr=BOdSr " + _
'                 "      Join (Select  OrCoCd,OrTc,OrYy,OrChr,OrNo,OrSr" + IIF(pRmSCtg <> "", ",OrRmCtg,OrRmSCtg", "") + ",sum(OrPrdQty) qOrPrdQty " + _
'                 "               From OrdRm " + _
'                 "            Where OrRmCtg='" + pRmCtg + "' " + _
'                 "            Group By OrCoCd,OrTc,OrYy,OrChr,OrNo,OrSr" + IIF(pRmSCtg <> "", ",OrRmCtg,OrRmSCtg", "") + ") OrRm " + _
'                 "                     On OrCoCd=OdCoCd and OrTc=OdTc and OrYy=OdYy and OrChr=OdChr and OrNo=OdNo and OrSr=OdSr " + IIF(pRmSCtg <> "", "  and  OrRmCtg='" + pRmCtg + "'  and OrRmSCtg='" + pRmSCtg + "' ", "") + _
'                 "      Join RmRt On RrTcTyp='QTT' and RrCmCtg='C' " + IIF(pRmSCtg <> "", " and RrCtg=OrRmCtg and RrSCtg=OrRmSCtg ", "") + _
'                 "                     And round(qOrPrdQty,3) between round(RrFrLn,3) and round(RrToLn,3) " + _
'                 "                     And RrCmCd In (OdOmCmCd,'ZSELF') And RrDmCd In (OdDmCd,'') " + _
'                 "      Left Outer Join Param On  PTyp='IWTFRORD' and PMCd=OdIWtFrOrd " + _
'                 " Where " + _
'                 "      BPrtKey='" + ctCurrPrtn + "' and BCoCd= '" + ADC("TCoCd") + "' and BYy= '" + pBYy + "' and BChr= '" + pBChr + "'  and BNo= " + CStr(pBNo) + _
'                 "      and (select vPValue3 from vParam where vPTyp='CHR' and vPCoCd=OdCoCd and vPMCd=OdTc and vPSCd=OdChr)= 'N' " + _
'                 " Order By (Case when RrCmCd=OdOmCmCd Then 1 Else 2 End), (Case when RrDmCd=OdDmCd Then 1 Else 2 End)")
'
'
'      If Not (rChkTol.EOF Or rChkTol.BOF) Then
'              wOrPrdQty = rChkTol!qOrPrdQty
'              wRrMinTol = IIF(IsNull(rChkTol!qRrMinTol), -1, rChkTol!qRrMinTol)
'              wRrMaxTol = IIF(IsNull(rChkTol!qRrMaxTol), -1, rChkTol!qRrMaxTol)
'              wMinTolDesc = rChkTol!qMinDesc
'              wMaxTolDesc = rChkTol!qMaxDesc
'
'              If moCn.RecSeek(" Select 'x' From Bag " + _
'                             "  Left Outer Join (Select  TdCoCd as qTdCoCd,TdBYy as qTdBYy,TdBChr as qTdBChr,TdBNo as qTdBNo," + IIF(pRmSCtg <> "", "TdRmCtg as qTdRmCtg,TdRmSCtg as qTdRmSCtg, ", "") + _
'                             "  sum(TdRmQty* (case when TdFrRmLoc= 'W' and TdFrRmDC= 'C' then -1 " + _
'                             "                    when TdFrRmLoc= 'W' and TdFrRmDC= 'D' then 1 " + _
'                             "                    when TdToRmLoc= 'W' and TdToRmDC= 'C' then -1 " + _
'                             "                    when TdToRmLoc= 'W' and TdToRmDC= 'D' then 1 else 0 end)) qTdRmQty " + _
'                             "  From Txnd where TdRmCtg='" + pRmCtg + "' " + _
'                             "       And Not (TdPrtKey='" + ctCurrPrtn + "' And TdCoCd='" + ADC("TCoCd") + "' And TdTc='" + ADC("TTc") + "' And TdYy='" + ADC("TYy") + "' And TdChr='" + ADC("TChr") + "' and TdNo=" + CStr(ADC("TNo")) + " And TdSr=" + CStr(pTdSr) + " And TdSrNo=" + CStr(pTdSrNo) + ")" + _
'                             "  Group By TdCoCd,TdBYy,TdBChr,TdBNo" + IIF(pRmSCtg <> "", ",TdRmCtg,TdRmSCtg", "") + ") TdRm " + _
'                             "           On bCoCd=qTdCoCd and BYy=qTdBYy and BChr=qTdBChr and BNo=qTdBNo " + IIF(pRmSCtg <> "", " and qTdRmCtg='" + pRmCtg + "' and qTdRmSCtg='" + pRmSCtg + "' ", "") + _
'                             " Where " + _
'                             "      BPrtKey='" + ctCurrPrtn + "' and BCoCd= '" + ADC("TCoCd") + "' and BYy= '" + pBYy + "' and BChr= '" + pBChr + "'  and BNo= " + CStr(pBNo) + _
'                             "      and Cast(BQty as Decimal(8,1))>0   and IsNull(" + IIF(UCase$(ADC.MenuCd) = "BV", CStr(wRrMinTol), CStr(wRrMaxTol)) + ",-1)>-1 " + _
'                             "             and convert(decimal(10,1),((IsNull(qTdRmQty,0)+" + IIF(UCase$(ADC.MenuCd) = "BV", "0", CStr(pAddWt)) + ")/BQty)) " + IIF(UCase$(ADC.MenuCd) = "BV", " < ", " > ") + " convert(decimal(10,1)," + IIF(UCase$(ADC.MenuCd) = "BV", CStr(wRrMinTol), CStr(wRrMaxTol)) + ") ") Then
'
'                 CheckTolErrMsg = "Diamond " + IIF(pRmSCtg <> "", "Sub-Ctg", "Ctg") + " Qty " + IIF(UCase$(ADC.MenuCd) = "BV", " is Less Than the Minimum ", " Exceeds the Maximum") + " Tolerance Qty [ in " + IIF(UCase$(ADC.MenuCd) = "BV", wMinTolDesc, wMaxTolDesc) + "], Only Authorised Users Can do this"
'                 Set rChkTol = Nothing
'                 Exit Function
'               End If
'      End If
'      Set rChkTol = Nothing
'    End If
'End Function
'

'Indigo.84-Updating Pl No.to Grid for the Bags given in the Copy RFId/Bag frame
Private Sub UpdPLNo()
  If ADC("WFGPLYY") = "" Or ADC("WFGPLCHR") = "" Or ADC("WFGPLNO") = 0 Then _
    DispMsg "Packing List No. Not Entered ", etError: Exit Sub
  
  Dim wRow As Integer, wGrdFgdRows As Integer, i As Integer, wTotRec As Integer
  Dim wRfIdStr As String, wRfIdNoArr() As String, wBagNo As String
    
  'Checking Bag No box whether balnk or not
  If ADC("wRfIdSel") = "" Then DispMsg "Enter Bag List", etError: Exit Sub
  'vk.4 calling function to replace enter with ,
  wRfIdStr = UCase(PrepareRfIdString(ADC("wRfIdSel")))
  
  'Checking Valid input format
  If Left(wRfIdStr, 1) = "," Or Right(wRfIdStr, 1) = "," Then DispMsg "Invalid Bag Input", etError: Exit Sub
  
  wTotRec = 0
  wGrdFgdRows = GrdFgd.Rows
  For wRow = 1 To wGrdFgdRows - 1
    With GrdFgd
      wBagNo = .Value(wRow, "FdBYy") + "/" + .Value(wRow, "FdBChr") + "/" + LTrim(RTrim(Str(.Value(wRow, "FdBNo"))))
          
      If "," + wRfIdStr + "," Like "*," + wBagNo + ",*" Then
            .Value(wRow, "FdPlYy") = ADC("WFGPLYY")
            .Value(wRow, "FdPlChr") = ADC("WFGPLCHR")
            .Value(wRow, "FdPlNo") = ADC("WFGPLNO")
            If .SaveRec(wRow) Then wTotRec = wTotRec + 1
            
      End If
    End With
  Next
  DispMsg "No. of Bags Copied : " + CStr(wTotRec), etWarning

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




