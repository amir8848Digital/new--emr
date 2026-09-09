VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Object = "{AB2CC32A-E652-42A2-B341-26ECE02F6386}#2.1#0"; "MwfCtl.ocx"
Begin VB.Form EmrFrmCertRmLot 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Certificate Detail - RmLot"
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
   Begin MwfCtl.MWCTL_HLP MWCTL_HLP1 
      Height          =   405
      Left            =   8910
      TabIndex        =   27
      TabStop         =   0   'False
      Top             =   9360
      Visible         =   0   'False
      Width           =   555
      _ExtentX        =   979
      _ExtentY        =   714
   End
   Begin VB.Frame FRA_KEYS 
      BorderStyle     =   0  'None
      Height          =   1530
      Left            =   0
      TabIndex        =   29
      Top             =   0
      Width           =   15135
      Begin MwfCtl.MWCTL_MED ATXT 
         Height          =   285
         Index           =   78
         Left            =   2040
         TabIndex        =   2
         ToolTipText     =   "Enter From Certificate Date"
         Top             =   120
         Width           =   885
         _ExtentX        =   1561
         _ExtentY        =   503
         MaxLength       =   8
         DataType        =   3
         IdName          =   "WCDDTFR"
         CmpStr          =   "CdDt>="
      End
      Begin MwfCtl.MWCTL_MED ATXT 
         Height          =   285
         Index           =   28
         Left            =   3735
         TabIndex        =   4
         ToolTipText     =   "Enter From Certificate No."
         Top             =   120
         Width           =   4620
         _ExtentX        =   8149
         _ExtentY        =   503
         MaxLength       =   30
         DataType        =   4
         IdName          =   "WCDCERTNOFR"
         CmpStr          =   "CdCertNo>="
      End
      Begin MwfCtl.MWCTL_MED ATXT 
         Height          =   285
         Index           =   29
         Left            =   3735
         TabIndex        =   5
         ToolTipText     =   "Enter To Certificate No."
         Top             =   405
         Width           =   4620
         _ExtentX        =   8149
         _ExtentY        =   503
         MaxLength       =   30
         DataType        =   4
         IdName          =   "WCDCERTNOTO"
         CmpStr          =   "CdCertNo<="
      End
      Begin MwfCtl.MWCTL_MED ATXT 
         Height          =   285
         Index           =   1
         Left            =   2040
         TabIndex        =   3
         ToolTipText     =   "Enter To Certificate Date"
         Top             =   480
         Width           =   885
         _ExtentX        =   1561
         _ExtentY        =   503
         MaxLength       =   8
         DataType        =   3
         IdName          =   "WCDDTTO"
         CmpStr          =   "CdDt<="
      End
      Begin MwfCtl.MWCTL_MED ATXT 
         Height          =   285
         Index           =   15
         Left            =   720
         TabIndex        =   0
         ToolTipText     =   "Enter Certificate Category"
         Top             =   120
         Width           =   420
         _ExtentX        =   741
         _ExtentY        =   503
         DataType        =   4
         IdName          =   "WCDCTG"
         CmpStr          =   "CdCtg="
      End
      Begin MwfCtl.MWCTL_MED ATXT 
         Height          =   285
         Index           =   16
         Left            =   720
         TabIndex        =   1
         ToolTipText     =   "Enter Rm Category"
         Top             =   450
         Width           =   420
         _ExtentX        =   741
         _ExtentY        =   503
         DataType        =   4
         IdName          =   "WCDRMCTG"
         CmpStr          =   "CdRmCtg="
      End
      Begin MwfCtl.MWCTL_MED ATXT 
         Height          =   285
         Index           =   8
         Left            =   9120
         TabIndex        =   6
         ToolTipText     =   "Enter From Rm Lot"
         Top             =   105
         Width           =   2640
         _ExtentX        =   4657
         _ExtentY        =   503
         MaxLength       =   16
         DataType        =   4
         IdName          =   "WCDRMLOTNOFR"
         CmpStr          =   "CdRmLotNo>="
      End
      Begin MwfCtl.MWCTL_MED ATXT 
         Height          =   285
         Index           =   9
         Left            =   9120
         TabIndex        =   7
         ToolTipText     =   "Enter To Rm Lot"
         Top             =   390
         Width           =   2640
         _ExtentX        =   4657
         _ExtentY        =   503
         MaxLength       =   16
         DataType        =   4
         IdName          =   "WCDRMLOTNOTO"
         CmpStr          =   "CdRmLotNo<="
      End
      Begin MwfCtl.MWCTL_MED ATXT 
         Height          =   285
         Index           =   0
         Left            =   12780
         TabIndex        =   8
         ToolTipText     =   "Enter From Rm Code"
         Top             =   105
         Width           =   2160
         _ExtentX        =   3810
         _ExtentY        =   503
         MaxLength       =   12
         DataType        =   4
         IdName          =   "WCDRMCDFR"
         CmpStr          =   "CdRmCd>="
      End
      Begin MwfCtl.MWCTL_MED ATXT 
         Height          =   285
         Index           =   2
         Left            =   12780
         TabIndex        =   9
         ToolTipText     =   "Enter To Rm Code"
         Top             =   390
         Width           =   2160
         _ExtentX        =   3810
         _ExtentY        =   503
         MaxLength       =   12
         DataType        =   4
         IdName          =   "WCDRMCDTO"
         CmpStr          =   "CdRmCd<="
      End
      Begin MwfCtl.MWCTL_MED ATXT 
         Height          =   285
         Index           =   3
         Left            =   6660
         TabIndex        =   18
         ToolTipText     =   "Enter From Rm Quality"
         Top             =   825
         Width           =   5940
         _ExtentX        =   10478
         _ExtentY        =   503
         MaxLength       =   30
         DataType        =   4
         IdName          =   "WCDRMQLTYFR"
         CmpStr          =   "CdRmQlty>="
      End
      Begin MwfCtl.MWCTL_MED ATXT 
         Height          =   285
         Index           =   4
         Left            =   6660
         TabIndex        =   19
         ToolTipText     =   "Enter To Rm Quality"
         Top             =   1110
         Width           =   5940
         _ExtentX        =   10478
         _ExtentY        =   503
         MaxLength       =   30
         DataType        =   4
         IdName          =   "WCDRMQLTYTO"
         CmpStr          =   "CdRmQlty<="
      End
      Begin MwfCtl.MWCTL_MED ATXT 
         Height          =   285
         Index           =   22
         Left            =   13560
         TabIndex        =   20
         ToolTipText     =   "Enter Sort By Options"
         Top             =   1080
         Width           =   1380
         _ExtentX        =   2434
         _ExtentY        =   503
         MaxLength       =   8
         DataType        =   4
         IdName          =   "WSORTON"
      End
      Begin MwfCtl.MWCTL_MED ATXT 
         Height          =   285
         Index           =   5
         Left            =   960
         TabIndex        =   10
         ToolTipText     =   "Enter From Company Code"
         Top             =   855
         Width           =   660
         _ExtentX        =   1164
         _ExtentY        =   503
         MaxLength       =   3
         DataType        =   4
         IdName          =   "WCDBCOCDFR"
         CmpStr          =   "CdBCoCd>="
      End
      Begin MwfCtl.MWCTL_MED ATXT 
         Height          =   285
         Index           =   6
         Left            =   960
         TabIndex        =   11
         ToolTipText     =   "Enter To Company Code"
         Top             =   1140
         Width           =   660
         _ExtentX        =   1164
         _ExtentY        =   503
         MaxLength       =   3
         DataType        =   4
         IdName          =   "WCDBCOCDTO"
         CmpStr          =   "CdBCoCd<="
      End
      Begin MwfCtl.MWCTL_MED ATXT 
         Height          =   285
         Index           =   7
         Left            =   2130
         TabIndex        =   12
         ToolTipText     =   "Enter From Bag Year"
         Top             =   840
         Width           =   540
         _ExtentX        =   953
         _ExtentY        =   503
         MaxLength       =   2
         DataType        =   4
         IdName          =   "WCDBYYFR"
         CmpStr          =   "CdBYy>="
      End
      Begin MwfCtl.MWCTL_MED ATXT 
         Height          =   285
         Index           =   10
         Left            =   2130
         TabIndex        =   13
         ToolTipText     =   "Enter To Bag Year"
         Top             =   1125
         Width           =   540
         _ExtentX        =   953
         _ExtentY        =   503
         MaxLength       =   2
         DataType        =   4
         IdName          =   "WCDBYYTO"
         CmpStr          =   "CdBYy<="
      End
      Begin MwfCtl.MWCTL_MED ATXT 
         Height          =   285
         Index           =   11
         Left            =   3360
         TabIndex        =   14
         ToolTipText     =   "Enter From Bag Character"
         Top             =   840
         Width           =   660
         _ExtentX        =   1164
         _ExtentY        =   503
         MaxLength       =   3
         DataType        =   4
         IdName          =   "WCDBCHRFR"
         CmpStr          =   "CdBChr>="
      End
      Begin MwfCtl.MWCTL_MED ATXT 
         Height          =   285
         Index           =   12
         Left            =   3360
         TabIndex        =   15
         ToolTipText     =   "Enter To Bag Character"
         Top             =   1125
         Width           =   660
         _ExtentX        =   1164
         _ExtentY        =   503
         MaxLength       =   3
         DataType        =   4
         IdName          =   "WCDBCHRTO"
         CmpStr          =   "CdBChr<="
      End
      Begin MwfCtl.MWCTL_MED ATXT 
         Height          =   285
         Index           =   13
         Left            =   4560
         TabIndex        =   16
         ToolTipText     =   "Enter From Bag Number"
         Top             =   840
         Width           =   900
         _ExtentX        =   1588
         _ExtentY        =   503
         Mask            =   "######0"
         MaxLength       =   7
         DataType        =   1
         IdName          =   "WCDBNOFR"
         CmpStr          =   "CdBNo>="
      End
      Begin MwfCtl.MWCTL_MED ATXT 
         Height          =   285
         Index           =   14
         Left            =   4560
         TabIndex        =   17
         ToolTipText     =   "Enter To Bag Number"
         Top             =   1125
         Width           =   900
         _ExtentX        =   1588
         _ExtentY        =   503
         Mask            =   "######0"
         MaxLength       =   7
         DataType        =   1
         IdName          =   "WCDBNOTO"
         CmpStr          =   "CdBNo<="
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
         Index           =   20
         Left            =   60
         TabIndex        =   65
         Top             =   840
         Width           =   375
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
         Index           =   19
         Left            =   60
         TabIndex        =   64
         Top             =   1200
         Width           =   375
      End
      Begin VB.Label ALBL 
         BackStyle       =   0  'Transparent
         Caption         =   "CoCd"
         BeginProperty Font 
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
         Left            =   360
         TabIndex        =   63
         Top             =   840
         Width           =   555
      End
      Begin VB.Label ALBL 
         BackStyle       =   0  'Transparent
         Caption         =   "BYy "
         BeginProperty Font 
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
         Left            =   1680
         TabIndex        =   62
         Top             =   840
         Width           =   675
      End
      Begin VB.Label ALBL 
         BackStyle       =   0  'Transparent
         Caption         =   "BChr "
         BeginProperty Font 
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
         Left            =   2760
         TabIndex        =   61
         Top             =   840
         Width           =   555
      End
      Begin VB.Label ALBL 
         BackStyle       =   0  'Transparent
         Caption         =   "BNo "
         BeginProperty Font 
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
         Left            =   4080
         TabIndex        =   60
         Top             =   840
         Width           =   555
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
         Index           =   18
         Left            =   1200
         TabIndex        =   59
         Top             =   480
         Width           =   375
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
         Index           =   17
         Left            =   1200
         TabIndex        =   58
         Top             =   120
         Width           =   375
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
         Index           =   16
         Left            =   12720
         TabIndex        =   57
         Top             =   1080
         Width           =   855
      End
      Begin VB.Label ALBL 
         BackStyle       =   0  'Transparent
         Caption         =   "RmLot "
         BeginProperty Font 
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
         Left            =   8460
         TabIndex        =   56
         Top             =   120
         Width           =   675
      End
      Begin VB.Label ALBL 
         BackStyle       =   0  'Transparent
         Caption         =   "RmCode "
         BeginProperty Font 
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
         Left            =   11880
         TabIndex        =   55
         Top             =   120
         Width           =   825
      End
      Begin VB.Label ALBL 
         BackStyle       =   0  'Transparent
         Caption         =   "FMBarcode"
         BeginProperty Font 
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
         Left            =   5550
         TabIndex        =   54
         Top             =   840
         Width           =   1185
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
         Index           =   12
         Left            =   60
         TabIndex        =   47
         Top             =   450
         Width           =   615
      End
      Begin VB.Label ALBL 
         BackStyle       =   0  'Transparent
         Caption         =   "CdCtg"
         BeginProperty Font 
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
         Left            =   60
         TabIndex        =   46
         Top             =   120
         Width           =   615
      End
      Begin VB.Label ALBL 
         BackStyle       =   0  'Transparent
         Caption         =   "CertNo"
         BeginProperty Font 
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
         Left            =   3000
         TabIndex        =   32
         Top             =   120
         Width           =   735
      End
      Begin VB.Label ALBL 
         BackStyle       =   0  'Transparent
         Caption         =   "Date "
         BeginProperty Font 
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
         Left            =   1560
         TabIndex        =   31
         Top             =   120
         Width           =   495
      End
   End
   Begin VB.Frame FRA_NKEYS 
      BorderStyle     =   0  'None
      Enabled         =   0   'False
      Height          =   9300
      Left            =   -120
      TabIndex        =   28
      Top             =   0
      Width           =   15375
      Begin VB.Frame FraNKeyAll 
         BorderStyle     =   0  'None
         Height          =   9165
         Left            =   150
         TabIndex        =   48
         Top             =   90
         Width           =   14940
         Begin MwfCtl.MWCTL_FLX GrdCert 
            Height          =   6135
            Left            =   0
            TabIndex        =   21
            Top             =   1440
            Width           =   14955
            _ExtentX        =   26379
            _ExtentY        =   10821
            Cols            =   20
            colname1        =   "CDBCOCD"
            heading1        =   "CoCd"
            datafld1        =   "CDBCOCD"
            datatype1       =   4
            maxlength1      =   3
            tooltiptext1    =   "Enter Company Code"
            colname2        =   "WBAGNO"
            datatype2       =   4
            maxlength2      =   14
            tooltiptext2    =   "Enter Bag Number Using Barcode Reader"
            colname3        =   "CDBYY"
            heading3        =   "Byy"
            datafld3        =   "CDBYY"
            datatype3       =   4
            maxlength3      =   2
            tooltiptext3    =   "Enter Bag Year"
            colname4        =   "CDBCHR"
            heading4        =   "BChr  "
            datafld4        =   "CDBCHR"
            datatype4       =   4
            maxlength4      =   3
            tooltiptext4    =   "Enter Bag Character"
            colname5        =   "CDBNO"
            heading5        =   "BNo           "
            datafld5        =   "CDBNO"
            datatype5       =   1
            mask5           =   "######0"
            maxlength5      =   7
            tooltiptext5    =   "Enter Bag Number"
            colname6        =   "CDRMCTG"
            heading6        =   "Ctg"
            datafld6        =   "CdRmCtg"
            datatype6       =   4
            maxlength6      =   1
            tooltiptext6    =   "Enter Rm Category"
            colname7        =   "CDRMLOTNO"
            heading7        =   "Lot No.                                       "
            datafld7        =   "CDRMLOTNO"
            datatype7       =   4
            maxlength7      =   16
            tooltiptext7    =   "Enter Rm LotNo."
            colname8        =   "CDCERTNO"
            heading8        =   "Cert No.                                                                                 "
            datafld8        =   "CdCertNo"
            datatype8       =   4
            maxlength8      =   30
            tooltiptext8    =   "Enter Certificate No."
            colname9        =   "CDDT"
            heading9        =   "Date        "
            datafld9        =   "CdDt"
            datatype9       =   3
            maxlength9      =   8
            tooltiptext9    =   "Enter Certificate Date"
            colname10       =   "CDRMCD"
            heading10       =   "Rm Code                     "
            datafld10       =   "CdRmCd"
            datatype10      =   4
            maxlength10     =   12
            tooltiptext10   =   "Enter Rm Code"
            colname11       =   "CDRMCOL"
            heading11       =   "Color       "
            datafld11       =   "CdRmCol"
            datatype11      =   4
            maxlength11     =   5
            tooltiptext11   =   "Enter Rm Color"
            colname12       =   "CDRMQTY"
            heading12       =   "Qty      "
            datafld12       =   "CdRmQty"
            datatype12      =   1
            mask12          =   "####0"
            maxlength12     =   5
            tooltiptext12   =   "Enter Rm Quantity"
            colname13       =   "CDRMWT"
            heading13       =   "Wt                 "
            datafld13       =   "CdRmWt"
            datatype13      =   2
            mask13          =   "#0.0000"
            maxlength13     =   7
            tooltiptext13   =   "Enter Rm Weight"
            colname14       =   "CDFMBARCODE"
            heading14       =   "FMBarCode                                                        "
            datafld14       =   "CdFMBarCode"
            datatype14      =   4
            maxlength14     =   30
            tooltiptext14   =   "Enter BarCode"
            colname15       =   "CDCMNT"
            datafld15       =   "CdCmnt"
            recalcon15      =   "WCMT"
            maxlength15     =   120
            style15         =   0
            colname16       =   "CDDOCNO"
            datafld16       =   "CdDocNo"
            recalcon16      =   "WDOCNO"
            maxlength16     =   45
            style16         =   0
            colname17       =   "CDDOCDT"
            datafld17       =   "CdDocDt"
            datatype17      =   3
            recalcon17      =   "WDOCDT"
            maxlength17     =   8
            style17         =   0
            colname18       =   "CDREM"
            datafld18       =   "CdRem"
            recalcon18      =   "WREMARK"
            maxlength18     =   120
            style18         =   0
            colname19       =   "CDRMQLTY"
            datafld19       =   "CdRmQlty"
            recalcon19      =   "WRMQLTY"
            maxlength19     =   45
            style19         =   0
         End
         Begin MwfCtl.MWCTL_MED ATXT 
            Height          =   285
            Index           =   17
            Left            =   1080
            TabIndex        =   22
            ToolTipText     =   "Enter Comment"
            Top             =   7680
            Width           =   13860
            _ExtentX        =   24448
            _ExtentY        =   503
            MaxLength       =   120
            DataType        =   4
            ReCalcParent    =   "GrdCert(CDCMNT)"
            IdName          =   "WCMT"
         End
         Begin MwfCtl.MWCTL_MED ATXT 
            Height          =   285
            Index           =   18
            Left            =   1080
            TabIndex        =   23
            ToolTipText     =   "Enter Document No."
            Top             =   8040
            Width           =   9420
            _ExtentX        =   16616
            _ExtentY        =   503
            MaxLength       =   45
            DataType        =   4
            ReCalcParent    =   "GrdCert(CDDOCNO)"
            IdName          =   "WDOCNO"
         End
         Begin MwfCtl.MWCTL_MED ATXT 
            Height          =   285
            Index           =   20
            Left            =   1080
            TabIndex        =   25
            ToolTipText     =   "Enter Remarks"
            Top             =   8400
            Width           =   13860
            _ExtentX        =   24448
            _ExtentY        =   503
            MaxLength       =   120
            DataType        =   4
            ReCalcParent    =   "GrdCert(CDREM)"
            IdName          =   "WREMARK"
         End
         Begin MwfCtl.MWCTL_MED ATXT 
            Height          =   285
            Index           =   19
            Left            =   11520
            TabIndex        =   24
            ToolTipText     =   "Enter Document Date"
            Top             =   8040
            Width           =   1485
            _ExtentX        =   2619
            _ExtentY        =   503
            MaxLength       =   8
            DataType        =   3
            ReCalcParent    =   "GrdCert(CDDOCDT)"
            IdName          =   "WDOCDT"
         End
         Begin MwfCtl.MWCTL_MED ATXT 
            Height          =   285
            Index           =   21
            Left            =   1080
            TabIndex        =   26
            ToolTipText     =   "Enter Rm Quality"
            Top             =   8760
            Width           =   9420
            _ExtentX        =   16616
            _ExtentY        =   503
            MaxLength       =   45
            DataType        =   4
            ReCalcParent    =   "GrdCert(CDRMQLTY)"
            IdName          =   "WRMQLTY"
         End
         Begin VB.Label LblRmQlty 
            BackStyle       =   0  'Transparent
            Caption         =   "Rm Quality"
            BeginProperty Font 
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
            TabIndex        =   53
            Top             =   8760
            Width           =   1095
         End
         Begin VB.Label ALBL 
            BackStyle       =   0  'Transparent
            Caption         =   "Comment"
            BeginProperty Font 
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
            Left            =   0
            TabIndex        =   52
            Top             =   7680
            Width           =   1095
         End
         Begin VB.Label ALBL 
            BackStyle       =   0  'Transparent
            Caption         =   "Doc. No"
            BeginProperty Font 
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
            Left            =   0
            TabIndex        =   51
            Top             =   8040
            Width           =   975
         End
         Begin VB.Label ALBL 
            BackStyle       =   0  'Transparent
            Caption         =   "Doc. Dt"
            BeginProperty Font 
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
            Left            =   10680
            TabIndex        =   50
            Top             =   8040
            Width           =   735
         End
         Begin VB.Label ALBL 
            BackStyle       =   0  'Transparent
            Caption         =   "Remarks"
            BeginProperty Font 
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
            Left            =   0
            TabIndex        =   49
            Top             =   8400
            Width           =   1215
         End
      End
   End
   Begin MwfCtl.MWCTL_ADC ADC 
      CausesValidation=   0   'False
      Height          =   555
      Left            =   0
      TabIndex        =   30
      Top             =   9360
      Width           =   14370
      _ExtentX        =   25347
      _ExtentY        =   979
      AllowAdd        =   0   'False
      AllowDelete     =   0   'False
      BtnWidth        =   885
      BtnHeight       =   435
      BtnMaskColor    =   8454143
      FontSz          =   8.25
   End
   Begin MSComDlg.CommonDialog CdImpExc 
      Left            =   2160
      Top             =   9240
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin MwfCtl.MWCTL_BTN1 CmdImpExc 
      Height          =   465
      Left            =   6240
      TabIndex        =   45
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
   Begin VB.Frame FraImpExc 
      Height          =   1515
      Left            =   2760
      TabIndex        =   33
      Top             =   7800
      Width           =   8235
      Begin MwfCtl.MWCTL_MED ATXT 
         Height          =   285
         Index           =   248
         Left            =   660
         TabIndex        =   34
         ToolTipText     =   "Excel File Name and Path"
         Top             =   555
         Width           =   6795
         _ExtentX        =   11986
         _ExtentY        =   503
         Enabled         =   0   'False
         MaxLength       =   250
         IdName          =   "WXLFILE"
      End
      Begin MwfCtl.MWCTL_MED ATXT 
         Height          =   285
         Index           =   249
         Left            =   750
         TabIndex        =   35
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
         Left            =   2340
         TabIndex        =   36
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
         Left            =   3990
         TabIndex        =   37
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
         Left            =   7560
         TabIndex        =   38
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
         Left            =   7560
         TabIndex        =   39
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
      Begin VB.Label ALBL 
         BackStyle       =   0  'Transparent
         Caption         =   "To Row"
         BeginProperty Font 
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
         Left            =   3150
         TabIndex        =   44
         Top             =   1065
         Width           =   765
      End
      Begin VB.Label ALBL 
         BackStyle       =   0  'Transparent
         Caption         =   "Fr Row"
         BeginProperty Font 
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
         TabIndex        =   43
         Top             =   1065
         Width           =   675
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
         TabIndex        =   42
         Top             =   1065
         Width           =   765
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
         TabIndex        =   41
         Top             =   555
         Width           =   615
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
         TabIndex        =   40
         Top             =   120
         Width           =   5640
      End
   End
   Begin MwfCtl.MWCTL_BTN1 CmdOpnPdf 
      Height          =   495
      Left            =   4320
      TabIndex        =   66
      ToolTipText     =   "Option To Open File"
      Top             =   9120
      Width           =   915
      _ExtentX        =   1614
      _ExtentY        =   873
      BackColor       =   -2147483633
      MaskColor       =   8454143
      DisableColor    =   0
      Caption         =   "S&how File"
      ReSizePic       =   0   'False
      BeginProperty LblFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BtnLeftDiff     =   0
      BtnTopDiff      =   0
   End
   Begin VB.OLE OLE1 
      Class           =   "AcroExch.Document.7"
      Enabled         =   0   'False
      Height          =   375
      Left            =   2640
      TabIndex        =   67
      Top             =   9600
      Visible         =   0   'False
      Width           =   975
   End
End
Attribute VB_Name = "EmrFrmCertRmLot"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim moCn As MwfLib.MDOConnection
Enum en_CertFraLot
  ImpExc = 0 'Import from Excel
End Enum

Private Sub ADC_KeyWhen()
  Call EnaDisaCmds(True)
  Call HideAllFras
  'ADC("WDOCDT") = MWLib.EmptyDate
  
End Sub

Private Sub ADC_Load()
  Set moCn = ADC.Connection
  Set ADC.FirKeyCtl = ADC("WCDCTG")
  Set ADC.FirNKeyCtl = GrdCert
  ADC.Child = "GrdCert"
  Set GrdCert.PrevCtl = ADC
  Set GrdCert.NextCtl = ADC
  GrdCert.HotKey = "A"
     
End Sub

Private Sub ADC_NonKeyWhen(pr_Cancel As Boolean, pr_ErrMsg As String)
  If ADC("WCDCTG") = "L" Then
    GrdCert.ColProp("CdBCoCd").Style = fgcinvisible
    GrdCert.ColProp("CdBYy").Style = fgcinvisible
    GrdCert.ColProp("CdBChr").Style = fgcinvisible
    GrdCert.ColProp("CdBNo").Style = fgcinvisible
    GrdCert.ColProp("WBAGNO").Style = fgcinvisible
    GrdCert.ColProp("CDRMCTG").Style = fgcNormal
    GrdCert.ColProp("CDRMLOTNO").Style = fgcNormal
    GrdCert.ColProp("CDRMCD").Style = fgcNormal
    GrdCert.ColProp("CDRMCOL").Style = fgcNormal
    GrdCert.ColProp("CDRMQTY").Style = fgcNormal
    GrdCert.ColProp("CDRMWT").Style = fgcNormal
    GrdCert.ColProp("CDFMBARCODE").Style = fgcNormal
    ADC("WRMQLTY").Visible = True:  LblRmQlty.Visible = True
  ElseIf ADC("WCDCTG") = "B" Then
    GrdCert.ColProp("CDRMCTG").Style = fgcinvisible
    GrdCert.ColProp("CDRMLOTNO").Style = fgcinvisible
    GrdCert.ColProp("CDRMCD").Style = fgcinvisible
    GrdCert.ColProp("CDRMCOL").Style = fgcinvisible
    GrdCert.ColProp("CDRMQTY").Style = fgcinvisible
    GrdCert.ColProp("CDRMWT").Style = fgcinvisible
    GrdCert.ColProp("CDFMBARCODE").Style = fgcinvisible
    GrdCert.ColProp("CdBCoCd").Style = fgcNormal
    GrdCert.ColProp("CdBYy").Style = fgcNormal
    GrdCert.ColProp("CdBChr").Style = fgcNormal
    GrdCert.ColProp("CdBNo").Style = fgcNormal
    GrdCert.ColProp("WBAGNO").Style = fgcNormal
    ADC("WRMQLTY").Visible = False:  LblRmQlty.Visible = False
  End If
  
  Call EnaDisaCmds(False)
  
End Sub

Private Sub atxt_FldChange(Index As Integer, ByVal IdName As String, Cancel As Boolean, ErrMsg As String, ByVal pv_NewValue As Variant)
Select Case UCase(IdName)
  Case Is = UCase("WCDCTG")
    Cancel = Not moCn.RecSeek("Select 'x' From Param Where PTyp= 'CDCTG' And PMCd='" + pv_NewValue + "'")
    If Cancel = True Then ErrMsg = "Invalid Category": Exit Sub
  Case Is = UCase("WCDRMCTG")
    Cancel = pv_NewValue <> "" And Not moCn.RecSeek("Select 'x' From Param Where PTyp= 'RMCTG' And PMCd='" + pv_NewValue + "'")
    If Cancel = True Then ErrMsg = "Invalid Rm Category": Exit Sub
  'RmQuality should not be empty, user must give value
  Case Is = UCase("WRMQLTY")
    If pv_NewValue = "" Then Cancel = True:   ErrMsg = "Rm Quality Should Not Be Empty ": Exit Sub
  Case Is = UCase("WSORTON")
    Cancel = pv_NewValue <> "" And (Not moCn.RecSeek("Select PMCd From Param where PTyp= 'SRTCD' " + _
                           "and PMCd = '" + pv_NewValue + "'"))
    If Cancel = True Then ErrMsg = "Invalid Sort Option": Exit Sub
    
End Select
End Sub

Private Sub ATXT_FldWhen(Index As Integer, ByVal IdName As String, Cancel As Boolean, ErrMsg As String)
Select Case IdName
  Case Is = UCase("WCDCTG")
    Call HlpList.PMCd("CDCTG")
  Case Is = UCase("WCDRMCTG")
    Call HlpList.PMCd("RMCTG")
  Case Is = UCase("WCDRMCDFR"), UCase("WCDRMCDTO")
    Call HlpList.RmCd("N", "'" + ADC("WCDRMCTG") + "'")
  Case Is = UCase("WCDBCOCDFR"), UCase("WCDBCOCDTO")
    Call HlpList.hCoCd
  Case Is = UCase("WCDBYYFR"), UCase("WCDBYYTO")
    Call HlpList.PMCd("YY")
  Case Is = UCase("WCDBCHRFR"), UCase("WCDBCHRTO")
    Call HlpList.PMCd("BCHR")
  Case Is = UCase("WCDBNOFR")
    Call HlpList.BNo(ADC("WCDBCOCDFR"), ADC("WCDBYYFR"), ADC("WCDBCHRFR"), "Y")
  Case Is = UCase("WCDBNOTO")
    Call HlpList.BNo(ADC("WCDBCOCDTO"), ADC("WCDBYYTO"), ADC("WCDBCHRTO"), "Y")
  Case Is = UCase("WCDRMLOTNOFR"), UCase("WCDRMLOTNOTO")
    Call HlpList.RmLotNo(gs_CoCd, ADC("WCDRMCTG"))
  Case Is = UCase("WSORTON")
    Call HlpList.PMCd("SRTCD")
    
End Select
End Sub

Private Sub CmdOpnPdf_Click()
Dim wPath As String, wFolder As String, wFile As String
  On Error GoTo ErrH
  wFolder = IIF(ADC("WCDCTG") = "L", "Cert_LotNo", "Cert_Bag")
  If ADC("WCDCTG") = "L" Then
    wFile = GrdCert.Value(GrdCert.Row, "CdRmCtg") + "." + GrdCert.Value(GrdCert.Row, "CdRmLotNo")
  Else
    wFile = GrdCert.Value(GrdCert.Row, "CdBCoCd") + "." + GrdCert.Value(GrdCert.Row, "CdBYy") + "." + GrdCert.Value(GrdCert.Row, "CdBChr") + "." + CStr(GrdCert.Value(GrdCert.Row, "CdBNo"))
  End If

  wPath = moCn.GetFldVal("Select PDesc225 from Param Where PTyp='PCT' and PMCd='3D'") + "\" + wFolder + "\" + wFile + ".pdf"
  OLE1.CreateLink wPath
  OLE1.DoVerb
  Exit Sub
ErrH:
  If Err.Number = 31031 Then
    DispMsg "File/Folder (" + wPath + ") Not Found", etError
  Else
    DispMsg Err.Description, etError
  End If
End Sub

Private Sub GrdCert_ColChange(ByVal RowNum As Integer, ByVal ColName As String, ByVal NewValue As Variant, Cancel As Boolean, ErrMsg As String)
Select Case ColName
  'RmCtg will allow only 'D'
  Case Is = UCase("CdRmCtg")
    If NewValue <> "D" Then Cancel = True: ErrMsg = "Invalid Value. Rm Category Should Be 'D'": Exit Sub
  Case Is = UCase("CdRmLotNo")
    If NewValue = "" Then Cancel = True: ErrMsg = "Rm LotNo Should Not Be Empty ": Exit Sub
  Case Is = UCase("CDCERTNO")
    If NewValue = "" Then Cancel = True: ErrMsg = "Certificate No. Should Not Be Empty ": Exit Sub
  Case Is = UCase("CdRmCol")
    Cancel = (Not moCn.RecSeek("Select PSCd From Param where PTyp ='RMCOL' " + _
                                    "and PMCd='" + GrdCert.Value(RowNum, "CdRmCtg") + "' " + _
                                    "and PSCd = '" + NewValue + "'"))
    If Cancel = True Then ErrMsg = "Invalid Rm Color": Exit Sub
    
  Case Is = UCase("CdRmCd")
    Cancel = (NewValue <> "") And (Not moCn.RecSeek("Select RmCd From RmMst where RmZ='N' and " + _
                                    "RmCtg='" + GrdCert.Value(RowNum, "CdRmCtg") + "' " + _
                                    "and RmCd = '" + NewValue + "'"))
    If Cancel = True Then ErrMsg = "Invalid Rm Code": Exit Sub
       
  Case Is = UCase("CdBCoCd")
    Cancel = (NewValue <> "") And (Not moCn.RecSeek("Select hCoCd From Head where HCd='" + ctSelfCmCd + "' And HCoCd='" + NewValue + "'"))
    If Cancel = True Then ErrMsg = "Invalid Company Code": Exit Sub
  Case Is = UCase("CdBYy")
    Cancel = (NewValue <> "") And (Not moCn.RecSeek("select PMCd from Param where PTyp= 'YY' and " + _
                         "PMCd = '" + NewValue + "'"))
    If Cancel = True Then ErrMsg = "Invalid Bag Year": Exit Sub
  Case Is = UCase("CdBChr")
    Cancel = (NewValue <> "") And (Not moCn.RecSeek("Select PMCd From Param Where PTyp= 'BCHR' And " + _
                                     "PMCd = '" + NewValue + "' "))
    If Cancel = True Then ErrMsg = "Invalid Bag Character": Exit Sub
  Case Is = UCase("CdBNo")
    Cancel = (NewValue <> 0) And (Not moCn.RecSeek("Select BNo From Bag Where BCoCd='" + GrdCert.Value(RowNum, "CdBCoCd") + _
                                      "' and BYy= '" + GrdCert.Value(RowNum, "CdBYy") + "' And BChr = '" + _
                                      GrdCert.Value(RowNum, "CdBChr") + "' And BNo= " + CStr(NewValue)))
    If Cancel = True Then ErrMsg = "Invalid Bag Number ": Exit Sub
  Case Is = UCase("CDRMQTY")
    If NewValue = 0 Then Cancel = True: ErrMsg = "Rm Qty Should Not Be 0 ": Exit Sub
  Case Is = UCase("CDRMWT")
    If NewValue = 0 Then Cancel = True: ErrMsg = "Rm Wt Should Not Be 0 ": Exit Sub
  Case Is = UCase("wBagNo")
    Dim ws_BagNo() As String
    Dim ws_OldYy As String, ws_OldChr As String, ws_OldNo As Single
    With GrdCert
      If NewValue <> "" Then
        ws_BagNo = Split(NewValue, "/")
        If UBound(ws_BagNo) <> 2 Then Cancel = True: ErrMsg = "Invalid Bag No": Exit Sub
        If Not IsNumeric(ws_BagNo(2)) Then Cancel = True: ErrMsg = "Invalid Bag No": Exit Sub
        
        ws_OldYy = .Value(RowNum, "CDBYY")
        ws_OldChr = .Value(RowNum, "CDBCHR")
        ws_OldNo = .Value(RowNum, "CDBNO")
        
        .Value(RowNum, "CDBYY") = UCase(ws_BagNo(0))
        Call GrdCert_ColChange(RowNum, "CDBYY", ws_BagNo(0), Cancel, "")
        If Cancel = True Then ErrMsg = "Invalid Bag Year": .Value(RowNum, "CDBYY") = ws_OldYy: Exit Sub
        
        .Value(RowNum, "CDBCHR") = UCase(ws_BagNo(1))
        Call GrdCert_ColChange(RowNum, "CDBCHR", ws_BagNo(1), Cancel, "")
        If Cancel = True Then ErrMsg = "Invalid Bag Character": .Value(RowNum, "CDBCHR") = ws_OldChr: Exit Sub

        .Value(RowNum, "CDBNO") = ws_BagNo(2)
        Call GrdCert_ColChange(RowNum, "CDBNO", ws_BagNo(2), Cancel, "")
        If Cancel = True Then ErrMsg = "Invalid Bag Number": .Value(RowNum, "CDBNO") = ws_OldNo: Exit Sub

    End If
  End With

End Select
End Sub

Private Sub GrdCert_ColWhen(ByVal RowNum As Integer, ByVal ColName As String, Cancel As Boolean, ErrMsg As String, ByVal Action As MwfCtl.en_fgAction)
Select Case ColName
  Case Is = UCase("CdRmCtg")
    Cancel = True:     Call HlpList.PMCd("RMCTG")
  Case Is = UCase("CdRmLotNo")
    Call HlpList.RmLotNo(gs_CoCd, GrdCert.Value(RowNum, "CdRmCtg"))
  Case Is = UCase("CdRmCol")
    Call HlpList.PSCd("RMCOL", GrdCert.Value(RowNum, "CdRmCtg"))
  Case Is = UCase("CdRmCd")
    Call HlpList.RmCd("N", "'" + GrdCert.Value(RowNum, "CdRmCtg") + "'")
  Case Is = UCase("CdBCoCd")
    Call HlpList.hCoCd
  Case Is = UCase("CdBYy")
    Call HlpList.PMCd("YY")
  Case Is = UCase("CdBChr")
    Call HlpList.PMCd("BCHR")
  Case Is = UCase("CdBNo")
    Call HlpList.BNo(GrdCert.Value(RowNum, "CdBCoCd"), GrdCert.Value(RowNum, "CdBYy"), GrdCert.Value(RowNum, "CdBChr"), "Y")
  Case Is = UCase("wBagNo")
    If moCn.RecSeek("Select CdCertNo From CertDet Where CdBCoCd='" + GrdCert.Value(RowNum, "CdBCoCd") + "' and " + _
                      "CdBYy='" + GrdCert.Value(RowNum, "CdBYy") + "' and CdBChr='" + GrdCert.Value(RowNum, "CdBChr") + "' and " + _
                      "CdBNo=" + CStr(GrdCert.Value(RowNum, "CdBNo")) + " and CdCertNo='" + GrdCert.Value(RowNum, "CdCertNo") + "'") Then _
    Cancel = True: ErrMsg = "Cannot Edit Bag Number": Exit Sub

End Select
End Sub

Private Sub GrdCert_InitKey(ByVal RowNum As Integer)
  '*** Initialise Key Fields
With GrdCert
  If ADC("WCDCTG") = "L" Then
    .Store "CdCtg", "L"
    .Store "CdBCoCd", ""
    .Store "CdBYy", ""
    .Store "CdBChr", ""
    .Store "CdBNo", 0
  Else
    .Store "CdCtg", "B"
    .Store "CdRmCtg", ""
    .Store "CdRmLotNo", ""
    .Store "CdRmQty", 0
    .Store "CdRmCol", ""
    .Store "CdRmQlty", ""
    .Store "CdRmCd", ""
    .Store "CdRmWt", 0
    .Store "CdFMBarCode", ""
  End If
End With
  
End Sub
Private Sub ADC_SetRecSource()
  '*** Since this is an 'A' Entry, there is no RecSource associated with ADC. So set the
      'DataMode property of ADC control as xdatUnbound

  ADC.DataMode = xdatUnbound
End Sub

Private Sub Form_GotFocus()

  SetMDISysColorsAfterLogin

End Sub
Private Sub Form_Load()
  Dim ws_BtnStr As String
  Call SetObjColors(Me, ws_BtnStr)
  
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

Private Sub GrdCert_ReCalc(ByVal RowNum As Integer, ByVal ColName As String, ByVal OldVal As Variant, ByVal SourceName As String, ByVal SourceColName As String)
With GrdCert
  Select Case UCase(ColName)
    Case Is = UCase("CdCmnt")
      If .Mode = fgmnorm Then
        If .Row > 0 Then
          If .Value(.Row, "CdCmnt") <> ADC("WCMT") Then
            .Value(.Row, "CdCmnt") = ADC("WCMT")
            .SaveRec .Row
          End If
        End If
      End If
    Case Is = UCase("CdDocNo")
      If .Mode = fgmnorm Then
        If .Row > 0 Then
          If .Value(.Row, "CdDocNo") <> ADC("WDOCNO") Then
            .Value(.Row, "CdDocNo") = ADC("WDOCNO")
            .SaveRec .Row
          End If
        End If
      End If
    Case Is = UCase("CdDocDt")
      If .Mode = fgmnorm Then
        If .Row > 0 Then
          If .Value(.Row, "CdDocDt") <> ADC("WDOCDT") Then
            .Value(.Row, "CdDocDt") = ADC("WDOCDT")
            .SaveRec .Row
          End If
        End If
      End If
    Case Is = UCase("CdRem")
      If .Mode = fgmnorm Then
        If .Row > 0 Then
          If .Value(.Row, "CdRem") <> ADC("WREMARK") Then
            .Value(.Row, "CdRem") = ADC("WREMARK")
            .SaveRec .Row
          End If
        End If
      End If
    Case Is = UCase("CdRmQlty")
      If .Mode = fgmnorm Then
        If .Row > 0 Then
          If .Value(.Row, "CdRmQlty") <> ADC("WRMQLTY") Then
            .Value(.Row, "CdRmQlty") = ADC("WRMQLTY")
            .SaveRec .Row
          End If
        End If
      End If
  End Select
End With
  
End Sub

Private Sub GrdCert_RowWhen(ByVal RowNum As Integer)
'If a row new added in grid then RmCtg must be 'D' and RmQty should be 1
With GrdCert
  If .IsNew(RowNum) = True Then
    If RowNum > 1 Then .Value(RowNum, "CdDt") = .Value(RowNum - 1, "CdDt")
    .Value(RowNum, "CdDocDt") = MWLib.EmptyDate
    If ADC("WCDCTG") = "L" Then
      .Value(RowNum, "CDRMCTG") = "D"
      .Value(RowNum, "CDRMQTY") = 1
    ElseIf ADC("WCDCTG") = "B" Then
      If RowNum > 1 Then .Value(RowNum, "CdBCoCd") = .Value(RowNum - 1, "CdBCoCd")
      If RowNum > 1 Then .Value(RowNum, "CdBYy") = .Value(RowNum - 1, "CdBYy")
      If RowNum > 1 Then .Value(RowNum, "CdBChr") = .Value(RowNum - 1, "CdBChr")
    End If
  End If
  ADC("WCMT") = .Value(RowNum, "CdCmnt")
  ADC("WDOCNO") = .Value(RowNum, "CdDocNo")
  ADC("WDOCDT") = .Value(RowNum, "CdDocDt")
  ADC("WREMARK") = .Value(RowNum, "CdRem")
  If ADC("WRMQLTY").Visible = True And ADC("WRMQLTY").Enabled = True Then ADC("WRMQLTY") = .Value(RowNum, "CdRmQlty")
 
End With
 
End Sub

Private Sub GrdCert_SetRecSource()
Dim ws_Cnd As String, ws_OrdBy As String
  ws_Cnd = ADC.RepCond
  ws_Cnd = IIF(ws_Cnd <> "", " and " + ws_Cnd, "")
  If ADC("WSORTON") <> "" Then ws_OrdBy = "Order By " + moCn.GetFldVal("Select PDesc225 from Param " + _
             "where PTyp='SRTCD' and PMCd='" + ADC("WSORTON") + "'")
    
  If moCn.GetFldVal("Select Count(*) from CertDet where CdCtg='" + IIF(ADC("WCDCTG") = "L", "L", "B") + "' " + ws_Cnd) > 2000 Then
      GrdCert.RecSource = " Select * from CertDet Where 1=2"
      MsgBox "More Than 2000 Records to Fetch, Please Reduce Scope !!"
      Exit Sub
  End If
  
  GrdCert.RecSource = " Select * from CertDet where CdCtg='" + IIF(ADC("WCDCTG") = "L", "L", "B") + "' " + ws_Cnd + ws_OrdBy
 
End Sub

Private Sub GrdCert_Validate(Cancel As Boolean)
  Cancel = GrdCert.Validate
End Sub

Private Sub DispFra(ByVal pv_CertFra As en_CertFraLot)
  '*** This is the common procedure which is used to toggle between hiding and displaying
      'the frames (i.e. if the frame is visible then make it invisible and viseversa)
      'associated with a command button

  Select Case pv_CertFra
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

  End Select
End Sub
Private Sub HideAllFras()
  '*** This procedure hides all the frames other than FraNKeyAll
  
  FraNKeyAll.Visible = True
  FraImpExc.Visible = False ' Import from excel
  FraImpExc.Enabled = False ' Import from excel
  
End Sub
Private Sub EnaDisaCmds(ByVal pv_ShowFra As Boolean, Optional ByVal pv_CmdBut As MWCTL_BTN1)
  '*** This procedure is used to give the toggle option (enable and disable) to
      'the command buttons

  If Not pv_CmdBut Is Nothing Then
    If pv_ShowFra = True Then
      pv_CmdBut.CausesValidation = True

    End If
  End If
  FraNKeyAll.Enabled = Not pv_ShowFra
  ADC.AllowSave = Not pv_ShowFra
  CmdImpExc.Enabled = Not pv_ShowFra
  CmdOpnPdf.Enabled = Not pv_ShowFra
  If Not pv_CmdBut Is Nothing Then pv_CmdBut.Enabled = True
End Sub

Private Sub CmdImpExc_Click()
  Call DispFra(ImpExc)
  
End Sub
'Getting excel file from selected path
Private Sub CmdImpExcPath_Click()
  CdImpExc.FileName = ""
  CdImpExc.DialogTitle = "Select Certificate RmLot Excel File "
  CdImpExc.Filter = "Excel Files(*.xls;*.xlsx)|*.xls;*.xlsx"
  CdImpExc.FilterIndex = 1
  CdImpExc.ShowOpen
  ADC("WXLFILE") = CdImpExc.FileName

End Sub
'Checking excel file name, Sheet no., From & To row nos. If all correct then call Import excel funtion
Private Sub CmdImpExcGo_Click()
  If Trim(ADC("WXLFILE")) = "" Then DispMsg "Please Specify a Proper Excel File Name !!", etError: Exit Sub
  If ADC("WXLSHTNO") = 0 Or ADC("WXLFRROW") = 0 Or ADC("WXLTOROW") = 0 Then DispMsg "Please Specify a Proper Sheet No./From Row/To Row !!", etError: Exit Sub

  Call ImportExcel
  Call DispFra(ImpExc)

End Sub
'Store data from excel file to Grid
Private Sub ImportExcel()
On Error GoTo ErrHdlr

    Dim wo_rsExcel As New MDORowSet, i As Integer
    Dim wCdRmCd As String, wCdRmLotNo As String, wCdRmCol As String, wCdRmQlty As String, wCdFMBarCode As String
    Dim wCdRmQty As Integer, wCdRmWt As Double
    Dim wCdCertNo As String, wCdComment As String, wCdDocNo As String, wCdRemark As String, wErrMsg As String
    
    'Checks proper excel file format
    If Not OpenExcelFile(ADC("WXLFILE")) Then DispMsg "Improper File Format", etError: Exit Sub

    'Checking proper sheet no.
    If ChkSheet(ADC("WXLSHTNO")) = False Then DispMsg "This WorkSheet Number Does Not Exist in the specified Excel File", etError: CloseExcel: Exit Sub

    'Creating Tmp file to store data
    moCn.CreateTmpTable "#TmpXls", "Select 0 as qSr, CdDt, CdRmLotNo, CdCertNo, CdRmQty, CdRmCol, CdRmQlty, CdRmCd, CdRmWt, CdFMBarCode, CdCmnt, CdDocNo,CdRem  Into #TmpXls From CertDet Where 1=2"

    ADC("WXLFRROW") = IIF(ADC("WXLFRROW") <= 1, 2, ADC("WXLFRROW"))
    For i = ADC("WXLFRROW") To ADC("WXLTOROW")
        wErrMsg = ""
        'Checking Date
        If ReadCell(ADC("WXLSHTNO"), i, 1) = "" Then Call SetCellVal(ADC("WXLSHTNO"), i, 1, "01/01/80")
                   
        'Checking Rm LotNo
        If (wCdRmLotNo <> "") And (Not moCn.RecSeek("select 'x' from SYyMm where SCoCd= '" + gs_CoCd + "' and " + _
                              IIF(wCdRmCd <> "", "SRmCd='" + wCdRmCd + "' and ", "") + "SLotNo= '" + wCdRmLotNo + "'")) Then _
            wErrMsg = wErrMsg + IIF(wErrMsg <> "", "/", "") + "Invalid LotNo. "
                        
        'Certifictae no.
        wCdCertNo = Trim(ReadCell(ADC("WXLSHTNO"), i, 3))
        If wCdCertNo = "" Then wErrMsg = wErrMsg + IIF(wErrMsg <> "", "/", "") + "Certificate Number cannot be blank "
        
        
        'Checking Rm Qty
        If ReadCell(ADC("WXLSHTNO"), i, 4) <> "" And Not IsNumeric(ReadCell(ADC("WXLSHTNO"), i, 4)) Then _
            wErrMsg = wErrMsg + IIF(wErrMsg <> "", "/", "") + "Rm Qty is not Number "
               
        wCdRmQty = Val(ReadCell(ADC("WXLSHTNO"), i, 4))
        If wCdRmQty < 0 Or wCdRmQty > 99999 Then _
            wErrMsg = wErrMsg + IIF(wErrMsg <> "", "/", "") + "Qty Should be Between 0 and 99999 "
        
        'Checking Rm Color
        wCdRmCol = Trim(ReadCell(ADC("WXLSHTNO"), i, 5))
        If (wCdRmCol <> "") And (Not moCn.RecSeek("Select PSCd From Param where PTyp ='RMCOL' " + _
                              "and PMCd='D' and PSCd = '" + wCdRmCol + "'")) Then _
            wErrMsg = wErrMsg + IIF(wErrMsg <> "", "/", "") + "Invalid Rm Color. "

        'Checking Rm Quality
        wCdRmQlty = Trim(ReadCell(ADC("WXLSHTNO"), i, 6))
        If wCdRmQlty = "" Then wErrMsg = wErrMsg + IIF(wErrMsg <> "", "/", "") + "Rm Quality Should Not Be Empty "
        
        'Checking Rm Code
        wCdRmCd = Trim(ReadCell(ADC("WXLSHTNO"), i, 7))
        If (wCdRmCd <> "") And (Not moCn.RecSeek("Select RmCd From RmMst where RmZ='N' " + _
                               "and RmCtg='D' and RmCd = '" + wCdRmCd + "'")) Then _
            wErrMsg = wErrMsg + IIF(wErrMsg <> "", "/", "") + "Invalid Rm Code "
            
        'Checking Rm Wt
        If ReadCell(ADC("WXLSHTNO"), i, 8) <> "" And Not IsNumeric(ReadCell(ADC("WXLSHTNO"), i, 8)) Then _
            wErrMsg = wErrMsg + IIF(wErrMsg <> "", "/", "") + "Rm Wt is not Number "
               
        wCdRmWt = Val(ReadCell(ADC("WXLSHTNO"), i, 8))
        If wCdRmWt < 0 Or wCdRmWt > 99999.999 Then _
            wErrMsg = wErrMsg + IIF(wErrMsg <> "", "/", "") + "Wt Should be Between 0 and 99999.999 "
        
        wCdFMBarCode = Trim(ReadCell(ADC("WXLSHTNO"), i, 9))
        wCdComment = Trim(ReadCell(ADC("WXLSHTNO"), i, 10))
        wCdDocNo = Trim(ReadCell(ADC("WXLSHTNO"), i, 11))
        wCdRemark = Trim(ReadCell(ADC("WXLSHTNO"), i, 12))
        
        Call SetCellVal(ADC("WXLSHTNO"), i, 14, wErrMsg)
        If wErrMsg = "" Then _
        moCn.Execute ("Insert Into #TmpXls Values (" + CStr(i) + ",'" + ReadCell(ADC("WXLSHTNO"), i, 1) + "','" + wCdRmLotNo + "','" + _
                        wCdCertNo + "'," + CStr(wCdRmQty) + ",'" + wCdRmCol + "','" + wCdRmQlty + "','" + wCdRmCd + "'," + CStr(wCdRmWt) + ",'" + _
                        wCdFMBarCode + "','" + wCdComment + "','" + wCdDocNo + "','" + wCdRemark + "'" + ")")
            
    Next
        
    Set wo_rsExcel = moCn.OpenRes("Select * From #TmpXls Order By qSr")
    CloseExcel
    wo_rsExcel.MoveFirst

    Dim wRow As Double

    If wo_rsExcel.RecCount > 0 Then
      With GrdCert
        .StartCopy
        MWLib.BeginProcess Me, "Copying Details ..."
        Do While Not wo_rsExcel.EOF
          .AddItem
          wRow = .Rows - 1
          .Row = wRow

          .Value(wRow, "CDDT") = wo_rsExcel!CdDt
          .Value(wRow, "CDRMCTG") = "D"
          .Value(wRow, "CDRMCD") = wo_rsExcel!CdRmCd
          .Value(wRow, "CDRMLOTNO") = wo_rsExcel!CdRmLotNo
          .Value(wRow, "CDCERTNO") = wo_rsExcel!CDCERTNO
          .Value(wRow, "CDRMQTY") = wo_rsExcel!CDRMQTY
          .Value(wRow, "CDRMCOL") = wo_rsExcel!CdRmCol
          .Value(wRow, "CDRMQLTY") = wo_rsExcel!CdRmQlty
          .Value(wRow, "CDRMWT") = wo_rsExcel!CDRMWT
          .Value(wRow, "CDFMBARCODE") = wo_rsExcel!CDFMBARCODE
          .Value(wRow, "CDCMNT") = wo_rsExcel!CdCmnt
          .Value(wRow, "CDDOCNO") = wo_rsExcel!CdDocNo
          .Value(wRow, "CDREM") = wo_rsExcel!CdRem

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
  ElseIf Err.Number = 70 Then
    DispMsg "Cannot Access Template File. Close If Already Open and Continue.", etError
  Else
    DispMsg Err.Description, etError
  End If

End Sub

