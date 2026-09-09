VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Object = "{AB2CC32A-E652-42A2-B341-26ECE02F6386}#2.1#0"; "MwfCtl.ocx"
Begin VB.Form EmrFrmCertEntry 
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
      Height          =   1410
      Left            =   0
      TabIndex        =   29
      Top             =   0
      Width           =   15135
      Begin MwfCtl.MWCTL_MED ATXT 
         Height          =   285
         Index           =   78
         Left            =   7320
         TabIndex        =   17
         Top             =   720
         Width           =   1020
         _ExtentX        =   1799
         _ExtentY        =   503
         MaxLength       =   8
         DataType        =   3
         IdName          =   "WCDDTFR"
         CmpStr          =   "CdDt>="
      End
      Begin MwfCtl.MWCTL_MED ATXT 
         Height          =   285
         Index           =   28
         Left            =   1200
         TabIndex        =   15
         Top             =   720
         Width           =   5160
         _ExtentX        =   9102
         _ExtentY        =   503
         MaxLength       =   30
         DataType        =   4
         IdName          =   "WCDCERTNOFR"
         CmpStr          =   "CdCertNo>="
      End
      Begin MwfCtl.MWCTL_MED ATXT 
         Height          =   285
         Index           =   29
         Left            =   1200
         TabIndex        =   16
         Top             =   1005
         Width           =   5160
         _ExtentX        =   9102
         _ExtentY        =   503
         MaxLength       =   30
         DataType        =   4
         IdName          =   "WCDCERTNOTO"
         CmpStr          =   "CdCertNo<="
      End
      Begin MwfCtl.MWCTL_MED ATXT 
         Height          =   285
         Index           =   1
         Left            =   7320
         TabIndex        =   18
         Top             =   1005
         Width           =   1020
         _ExtentX        =   1799
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
         Left            =   2160
         TabIndex        =   1
         ToolTipText     =   "Enter Rm Category"
         Top             =   120
         Width           =   420
         _ExtentX        =   741
         _ExtentY        =   503
         MaxLength       =   2
         DataType        =   4
         IdName          =   "WCDRMCTG"
         CmpStr          =   "CdRmCtg="
      End
      Begin MwfCtl.MWCTL_MED ATXT 
         Height          =   285
         Index           =   8
         Left            =   3720
         TabIndex        =   2
         ToolTipText     =   "Enter From Rm Lot No"
         Top             =   120
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
         Left            =   3720
         TabIndex        =   3
         ToolTipText     =   "Enter To Rm Lot No"
         Top             =   405
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
         Left            =   7320
         TabIndex        =   4
         ToolTipText     =   "Enter From Rm Code"
         Top             =   120
         Width           =   2280
         _ExtentX        =   4022
         _ExtentY        =   503
         MaxLength       =   16
         DataType        =   4
         IdName          =   "WCDRMCDFR"
         CmpStr          =   "CdRmCd>="
      End
      Begin MwfCtl.MWCTL_MED ATXT 
         Height          =   285
         Index           =   2
         Left            =   7320
         TabIndex        =   5
         ToolTipText     =   "Enter To Rm Code"
         Top             =   405
         Width           =   2280
         _ExtentX        =   4022
         _ExtentY        =   503
         MaxLength       =   16
         DataType        =   4
         IdName          =   "WCDRMCDTO"
         CmpStr          =   "CdRmCd<="
      End
      Begin MwfCtl.MWCTL_MED ATXT 
         Height          =   285
         Index           =   3
         Left            =   10320
         TabIndex        =   19
         ToolTipText     =   "Enter From Rm Quality"
         Top             =   720
         Width           =   4605
         _ExtentX        =   8123
         _ExtentY        =   503
         MaxLength       =   30
         DataType        =   4
         IdName          =   "WCDFMBARCDFR"
         CmpStr          =   "CdFMBarCd>="
      End
      Begin MwfCtl.MWCTL_MED ATXT 
         Height          =   285
         Index           =   4
         Left            =   10320
         TabIndex        =   20
         ToolTipText     =   "Enter To Rm Quality"
         Top             =   1005
         Width           =   4605
         _ExtentX        =   8123
         _ExtentY        =   503
         MaxLength       =   30
         DataType        =   4
         IdName          =   "WCDFMBARCDTO"
         CmpStr          =   "CdFMBarCd<="
      End
      Begin MwfCtl.MWCTL_MED ATXT 
         Height          =   285
         Index           =   22
         Left            =   13680
         TabIndex        =   14
         ToolTipText     =   "Enter Sort By Options"
         Top             =   405
         Width           =   1250
         _ExtentX        =   2196
         _ExtentY        =   503
         MaxLength       =   8
         DataType        =   4
         IdName          =   "WSORTON"
      End
      Begin MwfCtl.MWCTL_MED ATXT 
         Height          =   285
         Index           =   5
         Left            =   10320
         TabIndex        =   6
         ToolTipText     =   "Enter From Company Code"
         Top             =   120
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
         Left            =   10320
         TabIndex        =   7
         ToolTipText     =   "Enter To Company Code"
         Top             =   405
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
         Left            =   11130
         TabIndex        =   8
         ToolTipText     =   "Enter From Bag Year"
         Top             =   120
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
         Left            =   11130
         TabIndex        =   9
         ToolTipText     =   "Enter To Bag Year"
         Top             =   405
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
         Left            =   11800
         TabIndex        =   10
         ToolTipText     =   "Enter From Bag Character"
         Top             =   120
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
         Left            =   11800
         TabIndex        =   11
         ToolTipText     =   "Enter To Bag Character"
         Top             =   405
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
         Left            =   12600
         TabIndex        =   12
         ToolTipText     =   "Enter From Bag Number"
         Top             =   120
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
         Left            =   12600
         TabIndex        =   13
         ToolTipText     =   "Enter To Bag Number"
         Top             =   405
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
         Index           =   22
         Left            =   12480
         TabIndex        =   70
         Top             =   360
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
         Index           =   21
         Left            =   12480
         TabIndex        =   69
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
         Index           =   15
         Left            =   11700
         TabIndex        =   68
         Top             =   360
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
         Left            =   11700
         TabIndex        =   67
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
         Index           =   4
         Left            =   11040
         TabIndex        =   66
         Top             =   360
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
         Index           =   2
         Left            =   11040
         TabIndex        =   65
         Top             =   120
         Width           =   135
      End
      Begin VB.Label ALBL 
         BackStyle       =   0  'Transparent
         Caption         =   "Fr"
         BeginProperty Font 
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
         TabIndex        =   62
         Top             =   720
         Width           =   375
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
         Index           =   19
         Left            =   60
         TabIndex        =   61
         Top             =   1080
         Width           =   375
      End
      Begin VB.Label ALBL 
         BackStyle       =   0  'Transparent
         Caption         =   "BagNo"
         BeginProperty Font 
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
         Left            =   9650
         TabIndex        =   60
         Top             =   120
         Width           =   675
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
         Left            =   1200
         TabIndex        =   59
         Top             =   390
         Width           =   375
      End
      Begin VB.Label ALBL 
         BackStyle       =   0  'Transparent
         Caption         =   "Fr"
         BeginProperty Font 
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
         Caption         =   "Sort By"
         BeginProperty Font 
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
         Left            =   13680
         TabIndex        =   57
         Top             =   120
         Width           =   855
      End
      Begin VB.Label ALBL 
         BackStyle       =   0  'Transparent
         Caption         =   "Rm Lot No"
         BeginProperty Font 
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
         Left            =   2700
         TabIndex        =   56
         Top             =   120
         Width           =   1035
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
         Left            =   6480
         TabIndex        =   55
         Top             =   120
         Width           =   825
      End
      Begin VB.Label ALBL 
         BackStyle       =   0  'Transparent
         Caption         =   "FM BarCd"
         BeginProperty Font 
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
         Left            =   9285
         TabIndex        =   54
         Top             =   720
         Width           =   945
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
         Left            =   1500
         TabIndex        =   47
         Top             =   120
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
      Begin VB.Label LblCertNo 
         BackStyle       =   0  'Transparent
         Caption         =   "Cert No"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   420
         TabIndex        =   32
         Top             =   720
         Width           =   735
      End
      Begin VB.Label LblCertDt 
         BackStyle       =   0  'Transparent
         Caption         =   "Cert Dt"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   6480
         TabIndex        =   31
         Top             =   720
         Width           =   735
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
            Height          =   6615
            Left            =   0
            TabIndex        =   21
            Top             =   1320
            Width           =   14955
            _ExtentX        =   26379
            _ExtentY        =   11668
            Cols            =   21
            colname1        =   "CDBCOCD"
            heading1        =   "CoCd"
            datafld1        =   "CdBCoCd"
            datatype1       =   4
            maxlength1      =   3
            tooltiptext1    =   "Enter Company Code"
            colname2        =   "WBAGNO"
            datatype2       =   4
            maxlength2      =   14
            tooltiptext2    =   "Enter Bag Number Using Barcode Reader"
            colname3        =   "CDBYY"
            heading3        =   "BYy"
            datafld3        =   "CdBYy"
            datatype3       =   4
            maxlength3      =   2
            tooltiptext3    =   "Enter Bag Year"
            colname4        =   "CDBCHR"
            heading4        =   "BChr  "
            datafld4        =   "CdBChr"
            datatype4       =   4
            maxlength4      =   3
            tooltiptext4    =   "Enter Bag Character"
            colname5        =   "CDBNO"
            heading5        =   "BNo           "
            datafld5        =   "CdBNo"
            datatype5       =   1
            mask5           =   "######0"
            maxlength5      =   7
            tooltiptext5    =   "Enter Bag Number"
            colname6        =   "CDRMCTG"
            heading6        =   "Ctg"
            datafld6        =   "CdRmCtg"
            datatype6       =   4
            maxlength6      =   2
            tooltiptext6    =   "Enter Rm Category"
            colname7        =   "CDRMLOTNO"
            heading7        =   "Rm Lot No                                      "
            datafld7        =   "CdRmLotNo"
            datatype7       =   4
            maxlength7      =   16
            tooltiptext7    =   "Enter Rm LotNo."
            colname8        =   "CDCERTNO"
            heading8        =   "Cert No                                                                                  "
            datafld8        =   "CdCertNo"
            datatype8       =   4
            maxlength8      =   30
            tooltiptext8    =   "Enter Certificate No."
            colname9        =   "CDDT"
            heading9        =   "Cert Dt     "
            datafld9        =   "CdDt"
            datatype9       =   3
            maxlength9      =   8
            tooltiptext9    =   "Enter Certificate Date"
            colname10       =   "CDRMCD"
            heading10       =   "RmCode                                  "
            datafld10       =   "CdRmCd"
            datatype10      =   4
            maxlength10     =   16
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
            colname14       =   "CDFMBARCD"
            heading14       =   "FM BarCd                                                                                                      "
            datafld14       =   "CdFMBarCd"
            datatype14      =   4
            maxlength14     =   30
            tooltiptext14   =   "Enter FM BarCode"
            colname15       =   "CDSUBSHP"
            heading15       =   "SubShape      "
            datafld15       =   "CdSubShp"
            datatype15      =   4
            maxlength15     =   8
            tooltiptext15   =   "Enter Sub Shape"
            colname16       =   "CDCOMMENT"
            datafld16       =   "CdComment"
            recalcon16      =   "WCMT"
            maxlength16     =   120
            style16         =   0
            colname17       =   "CDDOCNO"
            datafld17       =   "CdDocNo"
            recalcon17      =   "WDOCNO"
            maxlength17     =   45
            style17         =   0
            colname18       =   "CDDOCDT"
            datafld18       =   "CdDocDt"
            datatype18      =   3
            recalcon18      =   "WDOCDT"
            maxlength18     =   8
            style18         =   0
            colname19       =   "CDREM"
            datafld19       =   "CdRem"
            recalcon19      =   "WREMARK"
            maxlength19     =   120
            style19         =   0
            colname20       =   "CDRMQLTY"
            datafld20       =   "CdRmQlty"
            recalcon20      =   "WRMQLTY"
            maxlength20     =   45
            style20         =   0
         End
         Begin MwfCtl.MWCTL_MED ATXT 
            Height          =   285
            Index           =   17
            Left            =   1080
            TabIndex        =   22
            ToolTipText     =   "Enter Comment"
            Top             =   8040
            Width           =   13860
            _ExtentX        =   24448
            _ExtentY        =   503
            MaxLength       =   120
            DataType        =   4
            ReCalcParent    =   "GrdCert(CDCOMMENT)"
            IdName          =   "WCMT"
         End
         Begin MwfCtl.MWCTL_MED ATXT 
            Height          =   285
            Index           =   18
            Left            =   1080
            TabIndex        =   23
            ToolTipText     =   "Enter Document No."
            Top             =   8325
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
            Top             =   8610
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
            Left            =   13440
            TabIndex        =   24
            ToolTipText     =   "Enter Document Date"
            Top             =   8325
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
            Top             =   8895
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
            Top             =   8895
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
            Top             =   8040
            Width           =   1095
         End
         Begin VB.Label ALBL 
            BackStyle       =   0  'Transparent
            Caption         =   "Doc No"
            BeginProperty Font 
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
            Top             =   8325
            Width           =   975
         End
         Begin VB.Label ALBL 
            BackStyle       =   0  'Transparent
            Caption         =   "Doc Dt"
            BeginProperty Font 
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
            Left            =   12600
            TabIndex        =   50
            Top             =   8325
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
            Top             =   8610
            Width           =   975
         End
      End
   End
   Begin MwfCtl.MWCTL_ADC ADC 
      CausesValidation=   0   'False
      Height          =   555
      Left            =   0
      TabIndex        =   30
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
      Left            =   4440
      TabIndex        =   63
      ToolTipText     =   "Option To Open File"
      Top             =   9360
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
      TabIndex        =   64
      Top             =   9600
      Visible         =   0   'False
      Width           =   975
   End
End
Attribute VB_Name = "EmrFrmCertEntry"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim moCn As MwfLib.MDOConnection
Enum en_CertFra
  ImpExc = 0 'Import from Excel
End Enum

Private Sub ADC_KeyWhen()
  Call EnaDisaCmds(True)
  Call HideAllFras
  adc("wCdDtFr") = MWLib.EmptyDate  '99.e to show default date
  adc("wCdDtTo") = MWLib.EmptyDate  '99.e to show default date
  adc("wDocDt") = MWLib.EmptyDate  '99.e to show default date
End Sub

Private Sub ADC_Load()
  Set moCn = adc.Connection
  Set adc.FirKeyCtl = adc("wCdCtg")
  Set adc.FirNKeyCtl = GrdCert
  adc.Child = "GrdCert"
  Set GrdCert.PrevCtl = adc
  Set GrdCert.NextCtl = adc
  GrdCert.HotKey = "A"
     
End Sub

Private Sub ADC_NonKeyWhen(pr_Cancel As Boolean, pr_ErrMsg As String)
  If adc("wCdCtg") = "L" Then
    GrdCert.ColProp("CdBCoCd").Style = fgcinvisible
    GrdCert.ColProp("CdBYy").Style = fgcinvisible
    GrdCert.ColProp("CdBChr").Style = fgcinvisible
    GrdCert.ColProp("CdBNo").Style = fgcinvisible
    GrdCert.ColProp("wBagNo").Style = fgcinvisible
    GrdCert.ColProp("CdRmCtg").Style = fgcNormal
    GrdCert.ColProp("CdRmLotNo").Style = fgcNormal
    GrdCert.ColProp("CdRmCd").Style = fgcNormal
    GrdCert.ColProp("CdRmCol").Style = fgcNormal
    GrdCert.ColProp("CdRmQty").Style = fgcNormal
    GrdCert.ColProp("CdRmWt").Style = fgcNormal
    GrdCert.ColProp("CdFMBarCd").Style = fgcNormal
    GrdCert.ColProp("CdSubShp").Style = fgcNormal     '99.f SubShape visible only for L
    adc("wRmQlty").Visible = True:  LblRmQlty.Visible = True
  ElseIf adc("wCdCtg") = "B" Or adc("wCdCtg") = "H" Then  '62 CdCtg=H added
    GrdCert.ColProp("CdRmCtg").Style = fgcinvisible
    GrdCert.ColProp("CdRmLotNo").Style = fgcinvisible
    GrdCert.ColProp("CdRmCd").Style = fgcinvisible
    GrdCert.ColProp("CdRmCol").Style = fgcinvisible
    GrdCert.ColProp("CdRmQty").Style = fgcinvisible
    GrdCert.ColProp("CdRmWt").Style = fgcinvisible
    GrdCert.ColProp("CdFMBarCd").Style = fgcinvisible
    GrdCert.ColProp("CdBCoCd").Style = fgcNormal
    GrdCert.ColProp("CdBYy").Style = fgcNormal
    GrdCert.ColProp("CdBChr").Style = fgcNormal
    GrdCert.ColProp("CdBNo").Style = fgcNormal
    GrdCert.ColProp("wBagNo").Style = fgcNormal
    GrdCert.ColProp("CdSubShp").Style = fgcinvisible     '99.f SubShape not required in B & in H
    adc("wRmQlty").Visible = False:  LblRmQlty.Visible = False
  End If
  '62 CdCtg=H then certificate heading & tooltip changed as HUID
  If adc("wCdCtg") = "H" Then
    GrdCert.ColProp("CdCertNo").Heading = "HUID No"
    GrdCert.ColProp("CdCertNo").ToolTipText = "Enter HUID No."
    GrdCert.ColProp("CdDt").Heading = "HUID Dt"
    GrdCert.ColProp("CdDt").ToolTipText = "Enter HUID Date"
  Else
    GrdCert.ColProp("CdCertNo").Heading = "Cert No"
    GrdCert.ColProp("CdCertNo").ToolTipText = "Enter Certificate No."
    GrdCert.ColProp("CdDt").Heading = "Cert Dt"
    GrdCert.ColProp("CdDt").ToolTipText = "Enter Certificate Date"
  End If
  
  Call EnaDisaCmds(False)
  
End Sub

Private Sub atxt_FldChange(Index As Integer, ByVal IdName As String, Cancel As Boolean, ErrMsg As String, ByVal pv_NewValue As Variant)
Select Case UCase(IdName)
  Case Is = UCase("wCdCtg")
    Cancel = Not moCn.RecSeek("Select 'x' From Param Where PTyp= 'CDCTG' And PMCd='" + pv_NewValue + "'")
    If Cancel = True Then ErrMsg = "Invalid Category": Exit Sub
    '62 if CdCtg=H then Certificate no/dt scope headings in X replaced with HUID
    LblCertNo.Caption = IIF(pv_NewValue = "H", "HUID No", "Cert No")
    LblCertDt.Caption = IIF(pv_NewValue = "H", "HUID Dt", "Cert Dt")
    '62 Cert no. From/To & Cert Date from/to tooltip changed as per CdCtg
    If pv_NewValue = "H" Then
      adc("wCdCertNoFr").ToolTipText = "Enter From HUID No.": adc("wCdCertNoTo").ToolTipText = "Enter To HUID No."
      adc("wCdDtFr").ToolTipText = "Enter From HUID Date": adc("wCdDtTo").ToolTipText = "Enter To HUID Date"
    Else
      adc("wCdCertNoFr").ToolTipText = "Enter From Certificate No.": adc("wCdCertNoTo").ToolTipText = "Enter To Certificate No."
      adc("wCdDtFr").ToolTipText = "Enter From Certificate Date": adc("wCdDtTo").ToolTipText = "Enter To Certificate Date"
    End If
    
  Case Is = UCase("wSortOn")
    Cancel = pv_NewValue <> "" And (Not moCn.RecSeek("Select PMCd From Param where PTyp= 'SRTCD' " + _
                           "and PMCd = '" + pv_NewValue + "'"))
    If Cancel = True Then ErrMsg = "Invalid Sort Option": Exit Sub

End Select
End Sub

Private Sub ATXT_FldWhen(Index As Integer, ByVal IdName As String, Cancel As Boolean, ErrMsg As String)
Select Case IdName
  Case Is = UCase("wCdCtg")
    Call HlpList.PMCd("CDCTG")
  Case Is = UCase("wCdRmCtg")
    Call HlpList.PMCd("RMCTG")
  Case Is = UCase("wCdRmCdFr"), UCase("wCdRmCdTo")
    Call HlpList.RmCd("N", "'" + adc("wCdRmCtg") + "'")
  Case Is = UCase("wCdBCoCdFr"), UCase("wCdBCoCdTo")
    Call HlpList.hCoCd
  Case Is = UCase("wCdBYyFr"), UCase("wCdBYyTo")
    Call HlpList.PMCd("YY")
  Case Is = UCase("wCdBChrFr"), UCase("wCdBChrTo")
    Call HlpList.PMCd("BCHR")
  Case Is = UCase("wCdBNoFr")
    Call HlpList.BNo(adc("wCdBCoCdFr"), adc("wCdBYyFr"), adc("wCdBChrFr"), "Y")
  Case Is = UCase("wCdBNoTo")
    Call HlpList.BNo(adc("wCdBCoCdTo"), adc("wCdBYyTo"), adc("wCdBChrTo"), "Y")
  Case Is = UCase("wCdRmLotNoFr"), UCase("wCdRmLotNoTo")
    Call HlpList.RmLotNo(adc("wCdRmCtg"))
  Case Is = UCase("wSortOn")
    Call HlpList.PMCd("SRTCD")
  '62 From certificate no. tooltip changed as per CtCtg
  Case Is = UCase("wCdCertNoFr")
    If adc("wCdCtg") = "H" Then DispMsg "Enter From HUID No.", etError Else DispMsg "Enter From Certificate No.", etError
  Case Is = UCase("wCdCertNoTo")
  '62 To certificate no. tooltip changed as per CtCtg
    If adc("wCdCtg") = "H" Then DispMsg "Enter To HUID No.", etError Else DispMsg "Enter To Certificate No.", etError
  '62 From certificate date tooltip changed as per CtCtg
  Case Is = UCase("wCdDtFr")
    If adc("wCdCtg") = "H" Then DispMsg "Enter From HUID Date", etError Else DispMsg "Enter From Certificate Date", etError
  '62 To certificate date tooltip changed as per CtCtg
  Case Is = UCase("wCdDtTo")
    If adc("wCdCtg") = "H" Then DispMsg "Enter To HUID Date", etError Else DispMsg "Enter To Certificate Date", etError
    
End Select
End Sub

Private Sub CmdOpnPdf_Click()
Dim wPath As String, wFolder As String, wFile As String
  On Error GoTo ErrH
  '62 CdCtg=H added
  wFolder = IIF(adc("wCdCtg") = "H", "HUID_Bag", IIF(adc("wCdCtg") = "L", "Cert_LotNo", "Cert_Bag"))
  If adc("wCdCtg") = "L" Then
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
    MsgBox "File/Folder (" + wPath + ") Not Found" ', etError
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

  Case Is = UCase("CdCertNo")
    '62 if CdCtg=H then Certificate replaced with HUID
    If NewValue = "" Then Cancel = True: ErrMsg = IIF(adc("wCdCtg") = "H", "HUID", "Certificate") + " No. Should Not Be Empty ": Exit Sub

  Case Is = UCase("CdRmCol")
    Cancel = (Not moCn.RecSeek("Select PSCd From Param where PTyp ='RMCOL' " + _
                                    "and PMCd='" + GrdCert.Value(RowNum, "CdRmCtg") + "' " + _
                                    "and PSCd = '" + NewValue + "'"))
    If Cancel = True Then ErrMsg = "Invalid Rm Color": Exit Sub

  Case Is = UCase("CdRmCd")
    Cancel = (Not moCn.RecSeek("Select RmCd From RmMst where RmZ='N' and " + _
                                    "RmCtg='" + GrdCert.Value(RowNum, "CdRmCtg") + "' " + _
                                    "and RmCd = '" + NewValue + "'"))
    If Cancel = True Then ErrMsg = "Invalid Rm Code": Exit Sub

  Case Is = UCase("CdBCoCd")
    Cancel = (Not moCn.RecSeek("Select hCoCd From Head where HCd='" + ctSelfCmCd + "' And HCoCd='" + NewValue + "'"))
    If Cancel = True Then ErrMsg = "Invalid Company Code": Exit Sub
  
  Case Is = UCase("CdBYy")
    Cancel = (Not moCn.RecSeek("select PMCd from Param where PTyp= 'YY' and " + _
                         "PMCd = '" + NewValue + "'"))
    If Cancel = True Then ErrMsg = "Invalid Bag Year": Exit Sub
  Case Is = UCase("CdBChr")
    Cancel = (Not moCn.RecSeek("Select PMCd From Param Where PTyp= 'BCHR' And " + _
                                     "PMCd = '" + NewValue + "' "))
    If Cancel = True Then ErrMsg = "Invalid Bag Character": Exit Sub
  Case Is = UCase("CdBNo")
    Cancel = (Not moCn.RecSeek("Select BNo From Bag Where BCoCd='" + GrdCert.Value(RowNum, "CdBCoCd") + _
                                      "' and BYy= '" + GrdCert.Value(RowNum, "CdBYy") + "' And BChr = '" + _
                                      GrdCert.Value(RowNum, "CdBChr") + "' And BNo= " + CStr(NewValue)))
    If Cancel = True Then ErrMsg = "Invalid Bag Number ": Exit Sub

  Case Is = UCase("CdRmQty")
    If NewValue = 0 Then Cancel = True: ErrMsg = "Rm Qty Should Not Be 0 ": Exit Sub
  Case Is = UCase("CdRmWt")
    If NewValue = 0 Then Cancel = True: ErrMsg = "Rm Wt Should Not Be 0 ": Exit Sub
  Case Is = UCase("wBagNo")
    Dim ws_BagNo() As String
    Dim ws_OldYy As String, ws_OldChr As String, ws_OldNo As Single
    With GrdCert
      If NewValue <> "" Then
        ws_BagNo = Split(NewValue, "/")
        If UBound(ws_BagNo) <> 2 Then Cancel = True: ErrMsg = "Invalid Bag No": Exit Sub
        If Not IsNumeric(ws_BagNo(2)) Then Cancel = True: ErrMsg = "Invalid Bag No": Exit Sub
        
        ws_OldYy = .Value(RowNum, "CdBYy")
        ws_OldChr = .Value(RowNum, "CdBChr")
        ws_OldNo = .Value(RowNum, "CdBNo")
        
        .Value(RowNum, "CdBYy") = UCase(ws_BagNo(0))
        Call GrdCert_ColChange(RowNum, "CdBYy", ws_BagNo(0), Cancel, "")
        If Cancel = True Then ErrMsg = "Invalid Bag Year": .Value(RowNum, "CdBYy") = ws_OldYy: Exit Sub
        
        .Value(RowNum, "CdBChr") = UCase(ws_BagNo(1))
        Call GrdCert_ColChange(RowNum, "CdBChr", ws_BagNo(1), Cancel, "")
        If Cancel = True Then ErrMsg = "Invalid Bag Character": .Value(RowNum, "CdBChr") = ws_OldChr: Exit Sub

        .Value(RowNum, "CdBNo") = ws_BagNo(2)
        Call GrdCert_ColChange(RowNum, "CdBNo", ws_BagNo(2), Cancel, "")
        If Cancel = True Then ErrMsg = "Invalid Bag Number": .Value(RowNum, "CdBNo") = ws_OldNo: Exit Sub

    End If
  End With
  
  '11.b validating subshape
  Case Is = UCase("CdSubShp")
    Cancel = NewValue <> "" And (Not moCn.RecSeek("Select PMCd From Param Where PTyp= 'SUBSHP' And " + _
                                     "PMCd = '" + NewValue + "' "))
    If Cancel = True Then ErrMsg = "Invalid SubShape ": Exit Sub
    
End Select
End Sub

Private Sub GrdCert_ColWhen(ByVal RowNum As Integer, ByVal ColName As String, Cancel As Boolean, ErrMsg As String, ByVal Action As MwfCtl.en_fgAction)
Select Case ColName
  Case Is = UCase("CdRmCtg")
    Cancel = True:     Call HlpList.PMCd("RMCTG")
  Case Is = UCase("CdRmLotNo")
    Call HlpList.RmLotNo(GrdCert.Value(RowNum, "CdRmCtg"))
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
    If moCn.RecSeek("Select CdCertNo From CertDet Where  CdBCoCd='" + GrdCert.Value(RowNum, "CdBCoCd") + "' and " + _
                      "CdBYy='" + GrdCert.Value(RowNum, "CdBYy") + "' and CdBChr='" + GrdCert.Value(RowNum, "CdBChr") + "' and " + _
                      "CdBNo=" + CStr(GrdCert.Value(RowNum, "CdBNo")) + " and CdCtg='" + adc("wCdCtg") + "' and CdCertNo='" + GrdCert.Value(RowNum, "CdCertNo") + "'") Then _
    Cancel = True: ErrMsg = "Cannot Edit Bag Number": Exit Sub
  '11.b calling help for subshape
  Case Is = UCase("CdSubShp")
    Call HlpList.PMCd("SUBSHP")
End Select
End Sub

Private Sub GrdCert_InitKey(ByVal RowNum As Integer)
  '*** Initialise Key Fields
With GrdCert
  If adc("wCdCtg") = "L" Then
    .Store "CdCtg", "L"
    .Store "CdBCoCd", ""
    .Store "CdBYy", ""
    .Store "CdBChr", ""
    .Store "CdBNo", 0
  Else  '62 CtCtg=H added
    .Store "CdCtg", IIF(adc("wCdCtg") = "H", "H", "B")
    .Store "CdRmCtg", ""
    .Store "CdRmLotNo", ""
    .Store "CdRmQty", 0
    .Store "CdRmCol", ""
    .Store "CdRmQlty", ""
    .Store "CdRmCd", ""
    .Store "CdRmWt", 0
    .Store "CdFMBarCd", ""
  End If
End With
  
End Sub
Private Sub ADC_SetRecSource()
  '*** Since this is an 'A' Entry, there is no RecSource associated with ADC. So set the
      'DataMode property of ADC control as xdatUnbound

  adc.DataMode = xdatUnbound
End Sub

Private Sub Form_GotFocus()

  SetMDISysColorsAfterLogin

End Sub
Private Sub Form_Load()
  Dim ws_BtnStr As String
  Call SetObjColors(Me, ws_BtnStr)
  
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
    
    Set moCn = Nothing
End Sub

Private Sub GrdCert_ReCalc(ByVal RowNum As Integer, ByVal ColName As String, ByVal OldVal As Variant, ByVal SourceName As String, ByVal SourceColName As String)
With GrdCert
  Select Case UCase(ColName)
    Case Is = UCase("CdComment")
      If .Mode = fgmnorm Then
        If .Row > 0 Then
          If .Value(.Row, "CdComment") <> adc("wCmt") Then
            .Value(.Row, "CdComment") = adc("wCmt")
            .SaveRec .Row
          End If
        End If
      End If
    Case Is = UCase("CdDocNo")
      If .Mode = fgmnorm Then
        If .Row > 0 Then
          If .Value(.Row, "CdDocNo") <> adc("wDocNo") Then
            .Value(.Row, "CdDocNo") = adc("wDocNo")
            .SaveRec .Row
          End If
        End If
      End If
    Case Is = UCase("CdDocDt")
      If .Mode = fgmnorm Then
        If .Row > 0 Then
          If .Value(.Row, "CdDocDt") <> adc("wDocDt") Then
            .Value(.Row, "CdDocDt") = adc("wDocDt")
            .SaveRec .Row
          End If
        End If
      End If
    Case Is = UCase("CdRem")
      If .Mode = fgmnorm Then
        If .Row > 0 Then
          If .Value(.Row, "CdRem") <> adc("wRemark") Then
            .Value(.Row, "CdRem") = adc("wRemark")
            .SaveRec .Row
          End If
        End If
      End If
    Case Is = UCase("CdRmQlty")
      If .Mode = fgmnorm Then
        If .Row > 0 Then
          If .Value(.Row, "CdRmQlty") <> adc("wRmQlty") Then
            .Value(.Row, "CdRmQlty") = adc("wRmQlty")
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
    '99.e first row shows default date other row will show previous row date for certificate date
    If RowNum = 1 Then
      .Value(RowNum, "CdDt") = MWLib.EmptyDate
    Else
      .Value(RowNum, "CdDt") = .Value(RowNum - 1, "CdDt")
    End If
    .Value(RowNum, "CdDocDt") = MWLib.EmptyDate
    If adc("wCdCtg") = "L" Then
      .Value(RowNum, "CdRmCtg") = "D"
      .Value(RowNum, "CdRmQty") = 1
    ElseIf adc("wCdCtg") = "B" Or adc("wCdCtg") = "H" Then  '62 CdCtg=H added
      If RowNum > 1 Then .Value(RowNum, "CdBCoCd") = .Value(RowNum - 1, "CdBCoCd")
      If RowNum > 1 Then .Value(RowNum, "CdBYy") = .Value(RowNum - 1, "CdBYy")
      If RowNum > 1 Then .Value(RowNum, "CdBChr") = .Value(RowNum - 1, "CdBChr")
    End If
    
  End If
  adc("wCmt") = .Value(RowNum, "CdComment")
  adc("wDocNo") = .Value(RowNum, "CdDocNo")
  adc("wDocDt") = .Value(RowNum, "CdDocDt")
  adc("wRemark") = .Value(RowNum, "CdRem")
  If adc("wRmQlty").Visible = True And adc("wRmQlty").Enabled = True Then adc("wRmQlty") = .Value(RowNum, "CdRmQlty")
  
End With
 
End Sub

Private Sub GrdCert_SetRecSource()
Dim ws_Cnd As String, ws_OrdBy As String, ws_DetOrdBy As String
  ws_Cnd = adc.RepCond
  ws_Cnd = IIF(ws_Cnd <> "", " and " + ws_Cnd, "")
  If adc("wSortOn") <> "" Then ws_OrdBy = moCn.GetFldVal("Select PDesc225 from Param " + _
             "where PTyp='SRTCD' and PMCd='" + adc("wSortOn") + "'")
  'Mw.9999 Err.43 max record changed from 2000 to 20000
  '62 IIF(ADC("wCdCtg") = "L", "L", "B") changed as ADC("wCdCtg")
  If moCn.GetFldVal("Select Count(*) from CertDet where CdCtg='" + adc("wCdCtg") + "' " + ws_Cnd) > 20000 Then
      GrdCert.RecSource = " Select * from CertDet Where 1=2"
      MsgBox "More Than 20000 Records to Fetch, Please Reduce Scope !!"
      Exit Sub
  End If
  
  If adc("wSortOn") = "CDBAGNO" Then
    ws_DetOrdBy = " CdRmCtg, CdRmLotNo,CdCertNo"
  ElseIf adc("wSortOn") = "CDCERT" Then
    ws_DetOrdBy = " CdRmCtg , CdRmLotNo , CdBCoCd, CdBYy ,CdBChr, CdBNo"
  ElseIf adc("wSortOn") = "CDDT" Then
    ws_DetOrdBy = " CdRmCtg , CdRmLotNo , CdBCoCd, CdBYy ,CdBChr, CdBNo, CdCertNo"
  ElseIf adc("wSortOn") = "CDRMLOT" Then
    ws_DetOrdBy = "  CdBCoCd, CdBYy ,CdBChr, CdBNo, CdCertNo"
  Else
    ws_DetOrdBy = " CdRmCtg, CdRmLotNo, CdBCoCd, CdBYy ,CdBChr, CdBNo, CdCertNo"
  End If
  
  '62 IIF(ADC("wCdCtg") = "L", "L", "B") changed as ADC("wCdCtg")
  GrdCert.RecSource = "Select * from CertDet where CdCtg='" + adc("wCdCtg") + "' " + ws_Cnd + _
                      " Order By " + IIF(ws_OrdBy = "", "", ws_OrdBy + ",") + ws_DetOrdBy
 
End Sub

Private Sub GrdCert_Validate(Cancel As Boolean)
  Cancel = GrdCert.Validate
End Sub

Private Sub DispFra(ByVal pv_CertFra As en_CertFra)
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
  adc.AllowSave = Not pv_ShowFra
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
  adc("wXlFile") = CdImpExc.FileName

End Sub
'Checking excel file name, Sheet no., From & To row nos. If all correct then call Import excel funtion
Private Sub CmdImpExcGo_Click()
  If Trim(adc("wXlFile")) = "" Then DispMsg "Please Specify a Proper Excel File Name !!", etError: Exit Sub
  If adc("wXlShtNo") = 0 Or adc("wXlFrRow") = 0 Or adc("wXlToRow") = 0 Then DispMsg "Please Specify a Proper Sheet No./From Row/To Row !!", etError: Exit Sub
  
  If adc("wCdCtg") = "L" Then Call ImportExcelLot
  If adc("wCdCtg") = "B" Or adc("wCdCtg") = "H" Then Call ImportExcelBag '62 CdCtg=H added
  Call DispFra(ImpExc)

End Sub
'Store data from excel file to Grid
Private Sub ImportExcelLot()
On Error GoTo ErrHdlr

    Dim wo_rsExcel As New MDORowSet, i As Integer
    Dim wCdRmCd As String, wCdRmLotNo As String, wCdRmCol As String, wCdRmQlty As String, wCdFMBarCode As String
    Dim wCdRmQty As Single, wCdRmWt As Double
    Dim wCdCertNo As String, wCdComment As String, wCdDocNo As String, wCdRemark As String, wErrMsg As String
    Dim wCdSubShp As String
        
    'Checks proper excel file format
    If Not OpenExcelFile(adc("wXlFile")) Then DispMsg "Improper File Format", etError: Exit Sub

    'Checking proper sheet no.
    If ChkSheet(adc("wXlShtNo")) = False Then DispMsg "This WorkSheet Number Does Not Exist in the specified Excel File", etError: CloseExcel: Exit Sub

    'Creating Tmp file to store data
    moCn.CreateTmpTable "#TmpXls", "Select 0 as qSr, CdDt, CdRmLotNo, CdCertNo, CdRmQty, CdRmCol, CdRmQlty, CdRmCd, CdRmWt, CdFMBarCd, CdComment, CdDocNo,CdDocDt, CdRem, CdSubShp Into #TmpXls From CertDet Where 1=2"

    adc("wXlFrRow") = IIF(adc("wXlFrRow") <= 1, 2, adc("wXlFrRow"))
    For i = adc("wXlFrRow") To adc("wXlToRow")
        wErrMsg = ""
        'Checking Date
        If ReadCell(adc("wXlShtNo"), i, 1) = "" Then Call SetCellVal(adc("wXlShtNo"), i, 1, "01/01/80")
                   
        'Checking Rm LotNo
        wCdRmLotNo = Trim(ReadCell(adc("wXlShtNo"), i, 2))
        If wCdRmLotNo = "" Then wErrMsg = wErrMsg + IIF(wErrMsg <> "", "/", "") + "Rm LotNo cannot be blank "
          
                        
        'Certifictae no.
        wCdCertNo = Trim(ReadCell(adc("wXlShtNo"), i, 3))
        If wCdCertNo = "" Then wErrMsg = wErrMsg + IIF(wErrMsg <> "", "/", "") + "Certificate Number cannot be blank "
        
        
        'Checking Rm Qty
        If Not IsNumeric(ReadCell(adc("wXlShtNo"), i, 4)) Then _
            wErrMsg = wErrMsg + IIF(wErrMsg <> "", "/", "") + "Rm Qty is not Number "
               
        wCdRmQty = Val(ReadCell(adc("wXlShtNo"), i, 4))
        If wCdRmQty < 1 Or wCdRmQty > 99999 Then _
            wErrMsg = wErrMsg + IIF(wErrMsg <> "", "/", "") + "Qty Should be Between 1 and 99999 "
        
        'Checking Rm Color
        wCdRmCol = Trim(ReadCell(adc("wXlShtNo"), i, 5))
        If (Not moCn.RecSeek("Select PSCd From Param where PTyp ='RMCOL' " + _
                              "and PMCd='D' and PSCd = '" + wCdRmCol + "'")) Then _
            wErrMsg = wErrMsg + IIF(wErrMsg <> "", "/", "") + "Invalid Rm Color. "

        'Checking Rm Quality
        wCdRmQlty = Trim(ReadCell(adc("wXlShtNo"), i, 6))
        
        'Checking Rm Code
        wCdRmCd = Trim(ReadCell(adc("wXlShtNo"), i, 7))
        If (Not moCn.RecSeek("Select RmCd From RmMst where RmZ='N' " + _
                               "and RmCtg='D' and RmCd = '" + wCdRmCd + "'")) Then _
            wErrMsg = wErrMsg + IIF(wErrMsg <> "", "/", "") + "Invalid RmCode "
            
        'Checking Rm Wt
        If Not IsNumeric(ReadCell(adc("wXlShtNo"), i, 8)) Then _
            wErrMsg = wErrMsg + IIF(wErrMsg <> "", "/", "") + "Rm Wt is not Number "
               
        wCdRmWt = Val(ReadCell(adc("wXlShtNo"), i, 8))
        If wCdRmWt <= 0 Or wCdRmWt > 99.9999 Then _
            wErrMsg = wErrMsg + IIF(wErrMsg <> "", "/", "") + "Wt Should be Between >0 and 99.9999 "
        
        wCdFMBarCode = Trim(ReadCell(adc("wXlShtNo"), i, 9))
        wCdComment = Trim(ReadCell(adc("wXlShtNo"), i, 10))
        wCdDocNo = Trim(ReadCell(adc("wXlShtNo"), i, 11))
        If ReadCell(adc("wXlShtNo"), i, 12) = "" Then Call SetCellVal(adc("wXlShtNo"), i, 12, "01/01/80")
        wCdRemark = Trim(ReadCell(adc("wXlShtNo"), i, 13))
        
        '11.b Checking Subshape
        wCdSubShp = Trim(ReadCell(adc("wXlShtNo"), i, 14))
        If (wCdSubShp <> "") And (Not moCn.RecSeek("Select PMCd From Param where  " + _
                              "PTyp= 'SUBSHP' and PMCd = '" + wCdSubShp + "'")) Then _
            wErrMsg = wErrMsg + IIF(wErrMsg <> "", "/", "") + "Invalid SubShape "
            
        
        Call SetCellVal(adc("wXlShtNo"), i, 16, wErrMsg)
        '11.b subshape added
        If wErrMsg = "" Then _
        moCn.Execute ("Insert Into #TmpXls Values (" + CStr(i) + ",'" + ReadCell(adc("wXlShtNo"), i, 1) + "','" + wCdRmLotNo + "','" + _
                        wCdCertNo + "'," + CStr(wCdRmQty) + ",'" + wCdRmCol + "','" + wCdRmQlty + "','" + wCdRmCd + "'," + CStr(wCdRmWt) + ",'" + _
                        wCdFMBarCode + "','" + wCdComment + "','" + wCdDocNo + "','" + ReadCell(adc("wXlShtNo"), i, 12) + "','" + wCdRemark + "','" + wCdSubShp + "'" + ")")
            
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

          .Value(wRow, "CdDt") = wo_rsExcel!CdDt
          .Value(wRow, "CdRmCtg") = "D"
          .Value(wRow, "CdRmCd") = wo_rsExcel!CdRmCd
          .Value(wRow, "CdRmLotNo") = wo_rsExcel!CdRmLotNo
          .Value(wRow, "CdCertNo") = wo_rsExcel!CDCERTNO
          .Value(wRow, "CdRmQty") = wo_rsExcel!CDRMQTY
          .Value(wRow, "CdRmCol") = wo_rsExcel!CdRmCol
          .Value(wRow, "CdRmWt") = wo_rsExcel!CDRMWT
          .Value(wRow, "CdFMBarCd") = wo_rsExcel!CdFMBarCd
          .Value(wRow, "CdRmQlty") = wo_rsExcel!CdRmQlty
          .Value(wRow, "CdComment") = wo_rsExcel!CdComment
          .Value(wRow, "CdDocNo") = wo_rsExcel!CdDocNo
          .Value(wRow, "CdDocDt") = wo_rsExcel!CdDocDt
          .Value(wRow, "CdRem") = wo_rsExcel!CdRem
          .Value(wRow, "CdSubShp") = wo_rsExcel!CdSubShp    '11.b subshape added
          adc("wCmt") = wo_rsExcel!CdComment
          adc("wDocNo") = wo_rsExcel!CdDocNo
          adc("wDocDt") = wo_rsExcel!CdDocDt
          adc("wRemark") = wo_rsExcel!CdRem
          adc("wRmQlty") = wo_rsExcel!CdRmQlty

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

'Store data from excel file to Grid
Private Sub ImportExcelBag()
On Error GoTo ErrHdlr
    Dim wo_rsExcel As New MDORowSet, i As Integer, wCdBNo As Single
    Dim wCdBCoCd As String, wCdBYy As String, wCdBChr As String
    Dim wCdCertNo As String, wCdComment As String, wCdDocNo As String, wCdRemark As String, wErrMsg As String
        
    'Checks proper excel file format
    If Not OpenExcelFile(adc("wXlFile")) Then DispMsg "Improper File Format", etError: Exit Sub

    'Checking proper sheet no.
    If ChkSheet(adc("wXlShtNo")) = False Then DispMsg "This WorkSheet Number Does Not Exist in the specified Excel File", etError: CloseExcel: Exit Sub

    'Creating Tmp file to store data
    moCn.CreateTmpTable "#TmpXls", "Select 0 as qSr, CdDt, CdBCoCd, CdBYy, CdBChr, CdBNo, CdCertNo, CdComment, CdDocNo, CdDocDt, CdRem Into #TmpXls From CertDet Where 1=2"

    adc("wXlFrRow") = IIF(adc("wXlFrRow") <= 1, 2, adc("wXlFrRow"))
    For i = adc("wXlFrRow") To adc("wXlToRow")
        wErrMsg = ""
        'Checking Date
        If ReadCell(adc("wXlShtNo"), i, 1) = "" Then Call SetCellVal(adc("wXlShtNo"), i, 1, "01/01/80")
                
        'Checking Company code
        wCdBCoCd = Trim(ReadCell(adc("wXlShtNo"), i, 2))
        If wCdBCoCd = "" Then wErrMsg = wErrMsg + IIF(wErrMsg <> "", "/", "") + "Bag Company Code cannot be blank "

        If (wCdBCoCd <> "") And (Not moCn.RecSeek("Select hCoCd From Head where  " + _
                                   "HCd='" + ctSelfCmCd + "' And HCoCd='" + wCdBCoCd + "'")) Then _
            wErrMsg = wErrMsg + IIF(wErrMsg <> "", "/", "") + "Invalid Company Code "
    
        'Checking Bag Year
        wCdBYy = Trim(ReadCell(adc("wXlShtNo"), i, 3))
        If wCdBYy = "" Then wErrMsg = wErrMsg + IIF(wErrMsg <> "", "/", "") + "Bag Year cannot be blank "
        

        If (wCdBYy <> "") And (Not moCn.RecSeek("Select PMCd From Param where PTyp= 'YY' " + _
                               "and PMCd = '" + wCdBYy + "'")) Then _
            wErrMsg = wErrMsg + IIF(wErrMsg <> "", "/", "") + "Invalid Bag Year "
            
        'Checking Bag Character
        wCdBChr = Trim(ReadCell(adc("wXlShtNo"), i, 4))
        If wCdBChr = "" Then wErrMsg = wErrMsg + IIF(wErrMsg <> "", "/", "") + "Bag Character cannot be blank "

        If (wCdBChr <> "") And (Not moCn.RecSeek("Select PMCd From Param where  " + _
                              "PTyp= 'BCHR' and PMCd = '" + wCdBChr + "'")) Then _
            wErrMsg = wErrMsg + IIF(wErrMsg <> "", "/", "") + "Invalid Bag Character "
            
        'Checking valid Bag number in the Excel File
        If Not IsNumeric(ReadCell(adc("wXlShtNo"), i, 5)) Then _
            wErrMsg = wErrMsg + IIF(wErrMsg <> "", "/", "") + "Bag Number is not Number "
                       

        wCdBNo = Val(ReadCell(adc("wXlShtNo"), i, 5))
        
        If (wCdBNo > 0) And (Not moCn.RecSeek("Select BNo From Bag where BCoCd='" + wCdBCoCd + "' and BYy= '" + wCdBYy + _
                              "' And BChr = '" + wCdBChr + "' And BNo= " + CStr(wCdBNo) + "")) Then _
            wErrMsg = wErrMsg + IIF(wErrMsg <> "", "/", "") + "Invalid Bag Number "

        
        wCdCertNo = Trim(ReadCell(adc("wXlShtNo"), i, 6))
        '62 HUID added
        If wCdCertNo = "" Then wErrMsg = wErrMsg + IIF(wErrMsg <> "", "/", "") + "HUID\Certificate Number cannot be blank "
        
        wCdComment = Trim(ReadCell(adc("wXlShtNo"), i, 7))
        wCdDocNo = Trim(ReadCell(adc("wXlShtNo"), i, 8))
        If ReadCell(adc("wXlShtNo"), i, 9) = "" Then Call SetCellVal(adc("wXlShtNo"), i, 9, "01/01/80")
        
        wCdRemark = Trim(ReadCell(adc("wXlShtNo"), i, 10))
        
                 
        
        If wErrMsg <> "" Then Call SetCellVal(adc("wXlShtNo"), i, 12, wErrMsg)
        
        If wErrMsg = "" Then _
        moCn.Execute ("Insert Into #TmpXls Values (" + CStr(i) + ",'" + ReadCell(adc("wXlShtNo"), i, 1) + "', '" + wCdBCoCd + "', '" + wCdBYy + "','" + _
                        wCdBChr + "', " + Str(wCdBNo) + ", '" + wCdCertNo + "', '" + wCdComment + "' , '" + wCdDocNo + "' ,'" + ReadCell(adc("wXlShtNo"), i, 9) + "', '" + wCdRemark + "'" + ")")
        
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

          .Value(wRow, "CdDt") = wo_rsExcel!CdDt
          .Value(wRow, "CdBCoCd") = wo_rsExcel!CdBCoCd
          .Value(wRow, "CdBYy") = wo_rsExcel!CdBYy
          .Value(wRow, "CdBChr") = wo_rsExcel!CdBChr
          .Value(wRow, "CdBNo") = wo_rsExcel!CdBNo
          .Value(wRow, "CdCertNo") = wo_rsExcel!CDCERTNO
          .Value(wRow, "CdComment") = wo_rsExcel!CdComment
          .Value(wRow, "CdDocNo") = wo_rsExcel!CdDocNo
          .Value(wRow, "CdDocDt") = wo_rsExcel!CdDocDt
          .Value(wRow, "CdRem") = wo_rsExcel!CdRem
          adc("wCmt") = wo_rsExcel!CdComment
          adc("wDocNo") = wo_rsExcel!CdDocNo
          adc("wDocDt") = wo_rsExcel!CdDocDt
          adc("wRemark") = wo_rsExcel!CdRem
          
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



