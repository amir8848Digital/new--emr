VERSION 5.00
Object = "{AB2CC32A-E652-42A2-B341-26ECE02F6386}#2.1#0"; "mwfctlforspeed9.ocx"
Begin VB.Form EmrFrmMultiPrcQt 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Multi Price Quotation"
   ClientHeight    =   9825
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   15210
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   9825
   ScaleWidth      =   15210
   WindowState     =   2  'Maximized
   Begin MwfCtl.MWCTL_BTN1 CmdCpyRfId 
      Height          =   435
      Left            =   4330
      TabIndex        =   127
      ToolTipText     =   "Copy Option"
      Top             =   9300
      Width           =   1395
      _ExtentX        =   2461
      _ExtentY        =   767
      BackColor       =   -2147483633
      MaskColor       =   8454143
      Caption         =   "Copy &Bag No / Bag Alias No."
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
   Begin MwfCtl.MWCTL_BTN1 CmdDsgDet 
      Height          =   435
      Left            =   5715
      TabIndex        =   108
      ToolTipText     =   "Option To Copy Grade From Another Customer"
      Top             =   9300
      Width           =   945
      _ExtentX        =   1667
      _ExtentY        =   767
      BackColor       =   -2147483633
      MaskColor       =   8454143
      Caption         =   "&Dsg Det"
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
   Begin MwfCtl.MWCTL_BTN1 CmdCpyDsg 
      Height          =   435
      Left            =   6660
      TabIndex        =   94
      ToolTipText     =   "Option To Copy Grade From Another Customer"
      Top             =   9300
      Width           =   945
      _ExtentX        =   1667
      _ExtentY        =   767
      BackColor       =   -2147483633
      MaskColor       =   8454143
      Caption         =   "&Copy Dsg"
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
   Begin MwfCtl.MWCTL_BTN1 CmdGenExcel 
      Height          =   435
      Left            =   8550
      TabIndex        =   76
      ToolTipText     =   "Option To Copy Grade From Another Customer"
      Top             =   9300
      Width           =   945
      _ExtentX        =   1667
      _ExtentY        =   767
      BackColor       =   -2147483633
      MaskColor       =   8454143
      Caption         =   "&Excel"
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
      Left            =   6180
      TabIndex        =   31
      TabStop         =   0   'False
      Top             =   9360
      Visible         =   0   'False
      Width           =   525
      _ExtentX        =   926
      _ExtentY        =   714
   End
   Begin VB.Frame FRA_KEYS 
      BorderStyle     =   0  'None
      Height          =   300
      Left            =   120
      TabIndex        =   33
      Top             =   90
      Width           =   4080
      Begin MwfCtl.MWCTL_MED ATXT 
         Height          =   285
         Index           =   4
         Left            =   2850
         TabIndex        =   3
         ToolTipText     =   "Enter Number"
         Top             =   0
         Width           =   750
         _ExtentX        =   1323
         _ExtentY        =   503
         Alignment       =   2
         Mask            =   "#####0"
         MaxLength       =   6
         DataType        =   1
         DataField       =   "MqmNo"
         IdName          =   "MQMNO"
      End
      Begin MwfCtl.MWCTL_MED ATXT 
         Height          =   285
         Index           =   37
         Left            =   1020
         TabIndex        =   0
         ToolTipText     =   "Enter Tc"
         Top             =   0
         Width           =   570
         _ExtentX        =   1005
         _ExtentY        =   503
         Enabled         =   0   'False
         MaxLength       =   3
         DataType        =   4
         DataField       =   "MqmTc"
         ReCalcParent    =   "MQMCHR"
         IdName          =   "MQMTC"
      End
      Begin MwfCtl.MWCTL_MED ATXT 
         Height          =   285
         Index           =   38
         Left            =   1680
         TabIndex        =   1
         ToolTipText     =   "Enter YY"
         Top             =   0
         Width           =   330
         _ExtentX        =   582
         _ExtentY        =   503
         MaxLength       =   2
         DataType        =   4
         DataField       =   "MqmYy"
         IdName          =   "MQMYY"
      End
      Begin MwfCtl.MWCTL_MED ATXT 
         Height          =   285
         Index           =   39
         Left            =   2160
         TabIndex        =   2
         ToolTipText     =   "Enter Character"
         Top             =   0
         Width           =   570
         _ExtentX        =   1005
         _ExtentY        =   503
         MaxLength       =   3
         DataType        =   4
         DataField       =   "MqmChr"
         ReCalcOn        =   "MQMTC"
         IdName          =   "MQMCHR"
      End
      Begin MwfCtl.MWCTL_MED ATXT 
         Height          =   285
         Index           =   40
         Left            =   3690
         TabIndex        =   102
         ToolTipText     =   "Enter Company Code"
         Top             =   0
         Visible         =   0   'False
         Width           =   210
         _ExtentX        =   370
         _ExtentY        =   503
         Enabled         =   0   'False
         MaxLength       =   3
         DataType        =   4
         DataField       =   "MqmCoCd"
         IdName          =   "MQMCOCD"
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
         Left            =   1620
         TabIndex        =   101
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
         Left            =   2070
         TabIndex        =   100
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
         Left            =   2760
         TabIndex        =   99
         Top             =   0
         Width           =   105
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
         Index           =   3
         Left            =   90
         TabIndex        =   34
         Top             =   0
         Width           =   795
      End
   End
   Begin MwfCtl.MWCTL_BTN1 CmdCpyVch 
      Height          =   435
      Left            =   7605
      TabIndex        =   36
      ToolTipText     =   "Option To Copy Grade From Another Customer"
      Top             =   9300
      Width           =   945
      _ExtentX        =   1667
      _ExtentY        =   767
      BackColor       =   -2147483633
      MaskColor       =   8454143
      Caption         =   "Copy &Vch"
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
   Begin MwfCtl.MWCTL_ADC ADC 
      CausesValidation=   0   'False
      Height          =   555
      Left            =   -30
      TabIndex        =   29
      Top             =   9300
      Width           =   15060
      _ExtentX        =   26564
      _ExtentY        =   979
      BtnWidth        =   885
      BtnHeight       =   435
      BtnMaskColor    =   8454143
      FontSz          =   8.25
   End
   Begin VB.Frame FRA_NKEYS 
      BorderStyle     =   0  'None
      Enabled         =   0   'False
      Height          =   9495
      Left            =   0
      TabIndex        =   32
      Top             =   0
      Width           =   15285
      Begin VB.Frame FraCpyDsg 
         Height          =   2835
         Left            =   5175
         TabIndex        =   77
         Top             =   6330
         Visible         =   0   'False
         Width           =   4785
         Begin VB.TextBox txtCpyDsgDmCd 
            Height          =   285
            Left            =   1440
            TabIndex        =   80
            ToolTipText     =   "Read Barcode Design"
            Top             =   750
            Width           =   2190
         End
         Begin MwfCtl.MWCTL_BTN1 CmdCpyDsgGo 
            Height          =   495
            Left            =   4140
            TabIndex        =   87
            ToolTipText     =   "Press This Button To Copy"
            Top             =   2265
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
            Index           =   11
            Left            =   1440
            TabIndex        =   81
            ToolTipText     =   "Enter Design Code"
            Top             =   1050
            Width           =   2190
            _ExtentX        =   3863
            _ExtentY        =   503
            MaxLength       =   15
            DataType        =   4
            ReCalcParent    =   "WCPYDSGSRC"
            IdName          =   "WCPYDSGDMCD"
         End
         Begin MwfCtl.MWCTL_MED ATXT 
            Height          =   285
            Index           =   16
            Left            =   1440
            TabIndex        =   85
            ToolTipText     =   "Enter Source Voucher"
            Top             =   2190
            Width           =   2430
            _ExtentX        =   4286
            _ExtentY        =   503
            Enabled         =   0   'False
            MaxLength       =   20
            DataType        =   4
            ReCalcOn        =   "WCPYDSGSRC"
            IdName          =   "WCPYDSGSRCVCH"
         End
         Begin MwfCtl.MWCTL_MED ATXT 
            Height          =   285
            Index           =   31
            Left            =   1440
            TabIndex        =   84
            ToolTipText     =   "Enter Design Source"
            Top             =   1905
            Width           =   1365
            _ExtentX        =   2408
            _ExtentY        =   503
            MaxLength       =   15
            DataType        =   4
            ReCalcParent    =   "WCPYDSGSRCVCH"
            ReCalcOn        =   "WCPYDSGDMCD,WCPYDSGSFX,WCPYDSGDMSZ"
            IdName          =   "WCPYDSGSRC"
         End
         Begin MwfCtl.MWCTL_MED ATXT 
            Height          =   285
            Index           =   32
            Left            =   1440
            TabIndex        =   82
            ToolTipText     =   "Enter Design Suffix"
            Top             =   1335
            Width           =   1635
            _ExtentX        =   2884
            _ExtentY        =   503
            MaxLength       =   15
            DataType        =   4
            ReCalcParent    =   "WCPYDSGSRC"
            IdName          =   "WCPYDSGSFX"
         End
         Begin MwfCtl.MWCTL_MED ATXT 
            Height          =   285
            Index           =   33
            Left            =   1440
            TabIndex        =   83
            ToolTipText     =   "Enter Design Size"
            Top             =   1620
            Width           =   885
            _ExtentX        =   1561
            _ExtentY        =   503
            MaxLength       =   15
            DataType        =   4
            ReCalcParent    =   "WCPYDSGSRC"
            IdName          =   "WCPYDSGDMSZ"
         End
         Begin MwfCtl.MWCTL_MED ATXT 
            Height          =   285
            Index           =   30
            Left            =   1440
            TabIndex        =   86
            ToolTipText     =   "Enter 'Y' to Copy Rates From Source"
            Top             =   2475
            Width           =   315
            _ExtentX        =   556
            _ExtentY        =   503
            DataType        =   4
            IdName          =   "WCPYDSGRT"
         End
         Begin MwfCtl.MWCTL_MED ATXT 
            Height          =   285
            Index           =   43
            Left            =   1440
            TabIndex        =   79
            ToolTipText     =   "Enter 'Y' to Design Using Barcode Else Enter 'N'"
            Top             =   450
            Width           =   315
            _ExtentX        =   556
            _ExtentY        =   503
            DataType        =   4
            IdName          =   "WCPYDSGBARCDYN"
         End
         Begin VB.Label ALBL 
            BackStyle       =   0  'Transparent
            Caption         =   "Use BarCode"
            BeginProperty Font 
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
            Left            =   120
            TabIndex        =   106
            Top             =   450
            Width           =   1455
         End
         Begin VB.Label ALBL 
            BackStyle       =   0  'Transparent
            Caption         =   "Barcode Dsg"
            BeginProperty Font 
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
            TabIndex        =   105
            Top             =   750
            Width           =   1335
         End
         Begin VB.Label ALBL 
            BackStyle       =   0  'Transparent
            Caption         =   "Copy Rt"
            BeginProperty Font 
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
            Left            =   120
            TabIndex        =   93
            Top             =   2490
            Width           =   855
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
            Index           =   10
            Left            =   120
            TabIndex        =   92
            Top             =   1050
            Width           =   855
         End
         Begin VB.Label ALBL 
            BackStyle       =   0  'Transparent
            Caption         =   "Suffix"
            BeginProperty Font 
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
            Left            =   120
            TabIndex        =   91
            Top             =   1335
            Width           =   855
         End
         Begin VB.Label ALBL 
            BackStyle       =   0  'Transparent
            Caption         =   "Size"
            BeginProperty Font 
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
            Left            =   120
            TabIndex        =   90
            Top             =   1635
            Width           =   855
         End
         Begin VB.Label ALBL 
            BackStyle       =   0  'Transparent
            Caption         =   "Src"
            BeginProperty Font 
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
            Left            =   120
            TabIndex        =   89
            Top             =   1935
            Width           =   855
         End
         Begin VB.Label ALBL 
            BackStyle       =   0  'Transparent
            Caption         =   "Src Vch"
            BeginProperty Font 
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
            Left            =   120
            TabIndex        =   88
            Top             =   2190
            Width           =   855
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
            Height          =   300
            Index           =   0
            Left            =   -30
            TabIndex        =   78
            Top             =   90
            Width           =   5325
         End
      End
      Begin VB.Frame FraDsgDet 
         Height          =   3645
         Left            =   180
         TabIndex        =   109
         Top             =   5520
         Width           =   14985
         Begin MwfCtl.MWCTL_FLX GrdMpQtLab 
            Height          =   2520
            Left            =   9150
            TabIndex        =   111
            Top             =   1080
            Width           =   5835
            _ExtentX        =   10292
            _ExtentY        =   4445
            Cols            =   8
            AllowAdd        =   0   'False
            AllowDelete     =   0   'False
            colname1        =   "WMQLSRNO"
            heading1        =   "Sr "
            datatype1       =   1
            mask1           =   "#0"
            maxlength1      =   2
            style1          =   2
            colname2        =   "WMQLMCD"
            heading2        =   "Main Cd"
            datatype2       =   4
            maxlength2      =   5
            style2          =   2
            colname3        =   "WMQLSCD"
            heading3        =   "Sub Cd        "
            datatype3       =   4
            maxlength3      =   8
            style3          =   2
            colname4        =   "WMQLQW"
            heading4        =   "By Q/W"
            datatype4       =   4
            maxlength4      =   1
            style4          =   2
            colname5        =   "WMQLQTY"
            heading5        =   "Qty "
            datatype5       =   1
            mask5           =   "##0"
            maxlength5      =   3
            style5          =   2
            colname6        =   "WMQLSALRT"
            heading6        =   "Rate          "
            datatype6       =   2
            mask6           =   "-#####0.00"
            maxlength6      =   10
            style6          =   2
            colname7        =   "WMQLSALVAL"
            heading7        =   "Value        "
            datatype7       =   2
            mask7           =   "-######0.00"
            maxlength7      =   11
            style7          =   2
         End
         Begin MwfCtl.MWCTL_FLX GrdMpQtRm 
            Height          =   2520
            Left            =   0
            TabIndex        =   110
            Top             =   1080
            Width           =   14655
            _ExtentX        =   25850
            _ExtentY        =   4445
            Cols            =   19
            AllowAdd        =   0   'False
            AllowDelete     =   0   'False
            colname1        =   "WMQRSRNO"
            heading1        =   "Sr "
            datatype1       =   1
            mask1           =   "#0"
            maxlength1      =   2
            style1          =   2
            colname2        =   "WMQRRMCTG"
            heading2        =   "Ctg"
            datatype2       =   4
            maxlength2      =   2
            style2          =   2
            colname3        =   "WMQRRMSCTG"
            heading3        =   "SubCtg"
            datatype3       =   4
            maxlength3      =   5
            style3          =   2
            colname4        =   "WMQRRMCD"
            heading4        =   "Rm Code                                  "
            datatype4       =   4
            maxlength4      =   16
            style4          =   2
            colname5        =   "WMQRLN1"
            heading5        =   "Sv/Ln  "
            datatype5       =   2
            mask5           =   "#0.000"
            maxlength5      =   6
            style5          =   2
            colname6        =   "WMQRLN2"
            heading6        =   "Brdth  "
            datatype6       =   4
            mask6           =   "#0.000"
            maxlength6      =   6
            style6          =   2
            colname7        =   "WMQRQTY"
            heading7        =   "Qty   "
            datatype7       =   1
            mask7           =   "###0"
            maxlength7      =   4
            style7          =   2
            colname8        =   "WMQRRMPTR"
            heading8        =   "Rm Ptr      "
            datatype8       =   2
            mask8           =   "#0.0000"
            maxlength8      =   7
            style8          =   2
            colname9        =   "WMQRWT"
            heading9        =   "Wt          "
            datatype9       =   2
            mask9           =   "-##0.000"
            maxlength9      =   8
            style9          =   2
            colname10       =   "WMQRLMESAL"
            heading10       =   "LME Rt   "
            datatype10      =   2
            mask10          =   "####0.00"
            maxlength10     =   8
            style10         =   2
            colname11       =   "WMQRSALRT"
            heading11       =   "Sale Rt       "
            datatype11      =   2
            mask11          =   "#######0.000"
            maxlength11     =   12
            style11         =   2
            colname12       =   "WMQRSALVAL"
            heading12       =   "Sale Value"
            datatype12      =   2
            mask12          =   "-#######0.000"
            maxlength12     =   13
            style12         =   2
            colname13       =   "WMQRSETSCD"
            heading13       =   "Setting        "
            datatype13      =   4
            maxlength13     =   8
            style13         =   2
            colname14       =   "WMQRSETSALRT"
            heading14       =   "Set Rt   "
            datatype14      =   2
            mask14          =   "#####0.00"
            maxlength14     =   9
            style14         =   2
            colname15       =   "WMQRSETSALVAL"
            heading15       =   "Set Value"
            datatype15      =   2
            mask15          =   "######0.00"
            maxlength15     =   10
            style15         =   2
            colname16       =   "WMQRALYCD"
            heading16       =   "Alloy                      "
            datatype16      =   4
            maxlength16     =   12
            style16         =   2
            colname17       =   "WMQRALYSALRT"
            heading17       =   "Alloy Rt  "
            datatype17      =   2
            mask17          =   "######0.00"
            maxlength17     =   10
            style17         =   2
            colname18       =   "WMQRMAINMET"
            heading18       =   "MMet"
            datatype18      =   4
            maxlength18     =   1
            style18         =   2
         End
         Begin MwfCtl.MWCTL_MED ATXT 
            Height          =   285
            Index           =   45
            Left            =   1050
            TabIndex        =   112
            ToolTipText     =   "Gld As"
            Top             =   420
            Width           =   855
            _ExtentX        =   1508
            _ExtentY        =   503
            MaxLength       =   5
            DataType        =   4
            IdName          =   "WMQDGLDAS"
         End
         Begin MwfCtl.MWCTL_MED ATXT 
            Height          =   285
            Index           =   46
            Left            =   1050
            TabIndex        =   114
            ToolTipText     =   "Gld As Wt"
            Top             =   705
            Width           =   855
            _ExtentX        =   1508
            _ExtentY        =   503
            Alignment       =   2
            Mask            =   "###0.000"
            MaxLength       =   8
            DataType        =   2
            IdName          =   "WMQDGLDASWT"
         End
         Begin MwfCtl.MWCTL_MED ATXT 
            Height          =   285
            Index           =   47
            Left            =   3210
            TabIndex        =   115
            ToolTipText     =   "Lab As Wt"
            Top             =   705
            Width           =   855
            _ExtentX        =   1508
            _ExtentY        =   503
            Alignment       =   2
            Mask            =   "###0.000"
            MaxLength       =   8
            DataType        =   2
            IdName          =   "WMQDLABASWT"
         End
         Begin MwfCtl.MWCTL_MED ATXT 
            Height          =   285
            Index           =   48
            Left            =   3210
            TabIndex        =   113
            ToolTipText     =   "Lab As"
            Top             =   420
            Width           =   855
            _ExtentX        =   1508
            _ExtentY        =   503
            MaxLength       =   5
            DataType        =   4
            IdName          =   "WMQDLABAS"
         End
         Begin VB.Label LblDsgDet 
            BackStyle       =   0  'Transparent
            BorderStyle     =   1  'Fixed Single
            Caption         =   "  Design Details"
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
            Left            =   0
            TabIndex        =   120
            Top             =   60
            Width           =   14955
         End
         Begin VB.Label ALBL 
            BackStyle       =   0  'Transparent
            Caption         =   "LabAs Wt"
            BeginProperty Font 
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
            Left            =   2130
            TabIndex        =   119
            Top             =   705
            Width           =   1050
         End
         Begin VB.Label ALBL 
            BackStyle       =   0  'Transparent
            Caption         =   "Lab As"
            BeginProperty Font 
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
            Left            =   2130
            TabIndex        =   118
            Top             =   420
            Width           =   960
         End
         Begin VB.Label ALBL 
            BackStyle       =   0  'Transparent
            Caption         =   "GldAs Wt"
            BeginProperty Font 
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
            TabIndex        =   117
            Top             =   720
            Width           =   900
         End
         Begin VB.Label ALBL 
            BackStyle       =   0  'Transparent
            Caption         =   "Gld As"
            BeginProperty Font 
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
            TabIndex        =   116
            Top             =   420
            Width           =   930
         End
      End
      Begin VB.Frame FraCpyRfId 
         Height          =   3315
         Left            =   360
         TabIndex        =   128
         Top             =   5760
         Visible         =   0   'False
         Width           =   14235
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
            Left            =   6210
            TabIndex        =   137
            ToolTipText     =   "Specify Copy From Bag Nos / Bag Alias No."
            Top             =   150
            Width           =   1335
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
            Left            =   7620
            TabIndex        =   136
            ToolTipText     =   "Specify Copy From Bag Nos / Bag Alias No."
            Top             =   150
            Width           =   1695
         End
         Begin VB.Frame Frame2 
            Caption         =   "Enter Bag Nos / Bag Alias No."
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
            TabIndex        =   134
            Top             =   375
            Width           =   13455
            Begin MwfCtl.MWCTL_MED ATXT 
               Height          =   2565
               Index           =   52
               Left            =   60
               TabIndex        =   135
               ToolTipText     =   "Enter Bag Alias No./ Bag Nos (Yy/Chr/No,Yy/Chr....,Yy/Chr/No)  String"
               Top             =   240
               Width           =   13335
               _ExtentX        =   23521
               _ExtentY        =   4524
               Alignment       =   3
               MaxLength       =   0
               IdName          =   "WRFIDSEL"
            End
         End
         Begin MwfCtl.MWCTL_BTN1 CmdCpyRfIdGo 
            Height          =   525
            Left            =   13560
            TabIndex        =   131
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
            Left            =   13050
            TabIndex        =   129
            ToolTipText     =   "Enter 'Y' to Copy Thru Bag Nos and  'N' to Copy thru Bag Alternate Id"
            Top             =   165
            Visible         =   0   'False
            Width           =   315
            _ExtentX        =   556
            _ExtentY        =   503
            DataType        =   4
            IdName          =   "WCPYTHRUBAGNO"
         End
         Begin MwfCtl.MWCTL_MED ATXT 
            Height          =   285
            Index           =   53
            Left            =   6960
            TabIndex        =   130
            Top             =   1080
            Width           =   315
            _ExtentX        =   556
            _ExtentY        =   503
            DataType        =   4
            IdName          =   "WRFIDTMPTAB"
         End
         Begin VB.Label Label12 
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
            Left            =   4920
            TabIndex        =   138
            Top             =   140
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
            Left            =   10440
            TabIndex        =   133
            Top             =   405
            Width           =   2325
         End
         Begin VB.Label LblMain 
            BackStyle       =   0  'Transparent
            BorderStyle     =   1  'Fixed Single
            Caption         =   "  Copy using Bag Nos / Bag Alias No."
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
            TabIndex        =   132
            Top             =   120
            Width           =   4650
         End
      End
      Begin VB.Frame FraNKeyAll 
         BorderStyle     =   0  'None
         Height          =   9315
         Left            =   120
         TabIndex        =   35
         Top             =   30
         Width           =   15030
         Begin VB.Frame FraPic3D 
            BorderStyle     =   0  'None
            Height          =   1635
            Left            =   13350
            TabIndex        =   97
            Top             =   360
            Width           =   1575
            Begin VB.Image Pic3D 
               Height          =   1515
               Left            =   0
               ToolTipText     =   "3D - Image"
               Top             =   30
               Width           =   1515
            End
         End
         Begin MwfCtl.MWCTL_FLX GrdMpQtDsg 
            Height          =   6870
            Left            =   -120
            TabIndex        =   30
            Top             =   2280
            Width           =   14985
            _ExtentX        =   26432
            _ExtentY        =   12118
            Cols            =   19
            AllowAdd        =   0   'False
            colname1        =   "MQDSR"
            heading1        =   "Sr       "
            datafld1        =   "MqdSr"
            datatype1       =   1
            mask1           =   "####0"
            recalcparent1   =   "GrdMpQtDsg(WMQDDIAGRD1PRC)"
            maxlength1      =   5
            style1          =   2
            tooltiptext1    =   "Enter Sr"
            colname2        =   "MQDDMCD"
            heading2        =   "Design Cd                    "
            datafld2        =   "MqdDmCd"
            datatype2       =   4
            maxlength2      =   15
            tooltiptext2    =   "Enter Design Code"
            colname3        =   "MQDDMSZ"
            heading3        =   "Size      "
            datafld3        =   "MqdDmSz"
            datatype3       =   4
            maxlength3      =   5
            colname4        =   "MQDMETGRD"
            heading4        =   "KT                 "
            datafld4        =   "MqdMetGrd"
            datatype4       =   4
            maxlength4      =   8
            tooltiptext4    =   "Enter KT"
            colname5        =   "MQDGRSWT"
            heading5        =   "Grs Wt     "
            datafld5        =   "MqdGrsWt"
            datatype5       =   2
            mask5           =   "####0.000"
            maxlength5      =   9
            colname6        =   "MQDDIAWT"
            heading6        =   "Dia Wt   "
            datafld6        =   "MqdDiaWt"
            datatype6       =   2
            mask6           =   "###0.000"
            maxlength6      =   8
            colname7        =   "WMQDDIAGRD1PRC"
            heading7        =   "Grd 1 Prc"
            datatype7       =   2
            mask7           =   "#######0.00"
            recalcon7       =   "GrdMpQtDsg(MQDSR)"
            maxlength7      =   11
            tooltiptext7    =   "Enter Price"
            colname8        =   "WMQDDIAGRD2PRC"
            heading8        =   "Grd 2 prc"
            datatype8       =   2
            mask8           =   "#######0.00"
            maxlength8      =   11
            colname9        =   "WMQDDIAGRD3PRC"
            heading9        =   "Grd 3 Prc"
            datatype9       =   2
            mask9           =   "#######0.00"
            maxlength9      =   11
            colname10       =   "WMQDDIAGRD4PRC"
            heading10       =   "Grd 4 Prc"
            datatype10      =   2
            mask10          =   "#######0.00"
            maxlength10     =   11
            colname11       =   "WMQDDIAGRD5PRC"
            heading11       =   "Grd 5 Prc"
            datatype11      =   2
            mask11          =   "#######0.00"
            maxlength11     =   11
            colname12       =   "WMQDDIAGRD6PRC"
            heading12       =   "Grd 6 Prc"
            datatype12      =   2
            mask12          =   "#######0.00"
            maxlength12     =   11
            colname13       =   "WMQDDIAGRD7PRC"
            heading13       =   "Grd 7 Prc"
            datatype13      =   2
            mask13          =   "#######0.00"
            maxlength13     =   11
            colname14       =   "WMQDDIAGRD8PRC"
            heading14       =   "Grd 8 Prc"
            datatype14      =   2
            mask14          =   "#######0.00"
            maxlength14     =   11
            colname15       =   "MQDGLDAS"
            datafld15       =   "MqdGldAs"
            datatype15      =   4
            maxlength15     =   5
            style15         =   0
            colname16       =   "MQDGLDASWT"
            datafld16       =   "MqdGldAsWt"
            datatype16      =   2
            mask16          =   "###0.000"
            maxlength16     =   8
            style16         =   0
            colname17       =   "MQDLABAS"
            datafld17       =   "MqdLabAs"
            datatype17      =   4
            maxlength17     =   5
            style17         =   0
            colname18       =   "MQDLABASWT"
            datafld18       =   "MqdLabAsWt"
            datatype18      =   2
            mask18          =   "###0.000"
            maxlength18     =   8
            style18         =   0
         End
         Begin MwfCtl.MWCTL_MED ATXT 
            Height          =   285
            Index           =   7
            Left            =   1020
            TabIndex        =   5
            ToolTipText     =   "Enter Customer"
            Top             =   360
            Width           =   1410
            _ExtentX        =   2487
            _ExtentY        =   503
            MaxLength       =   8
            DataType        =   4
            DataField       =   "MqmCmCd"
            ReCalcParent    =   "WCMNAME,MQMMULBY,MQMLMGSAL,MQMCMCURCD"
            IdName          =   "MQMCMCD"
         End
         Begin MwfCtl.MWCTL_MED ATXT 
            Height          =   285
            Index           =   8
            Left            =   4680
            TabIndex        =   4
            ToolTipText     =   "Enter Date"
            Top             =   60
            Width           =   885
            _ExtentX        =   1561
            _ExtentY        =   503
            MaxLength       =   8
            DataType        =   3
            DataField       =   "MqmDt"
            ReCalcParent    =   "MQMLMGSAL,MQMCNVFCT"
            IdName          =   "MQMDT"
         End
         Begin MwfCtl.MWCTL_MED ATXT 
            Height          =   285
            Index           =   9
            Left            =   1020
            TabIndex        =   7
            ToolTipText     =   "Enter Lmg Sales"
            Top             =   660
            Width           =   1050
            _ExtentX        =   1852
            _ExtentY        =   503
            Alignment       =   2
            Mask            =   "####0.000"
            MaxLength       =   9
            DataType        =   2
            DataField       =   "MqmLmgSal"
            ReCalcOn        =   "MQMDT,MQMCMCD,MQMLKUPRMRT,MQMCMCURCD"
            IdName          =   "MQMLMGSAL"
         End
         Begin MwfCtl.MWCTL_MED ATXT 
            Height          =   285
            Index           =   10
            Left            =   3120
            TabIndex        =   8
            ToolTipText     =   "Enter Lmg Sales"
            Top             =   645
            Width           =   1050
            _ExtentX        =   1852
            _ExtentY        =   503
            Alignment       =   2
            Mask            =   "####0.000"
            MaxLength       =   9
            DataType        =   2
            DataField       =   "MqmLmpSal"
            IdName          =   "MQMLMPSAL"
         End
         Begin MwfCtl.MWCTL_MED ATXT 
            Height          =   285
            Index           =   12
            Left            =   1020
            TabIndex        =   11
            ToolTipText     =   "Enter Lmg Sales"
            Top             =   945
            Width           =   1050
            _ExtentX        =   1852
            _ExtentY        =   503
            Alignment       =   2
            Mask            =   "####0.000"
            MaxLength       =   9
            DataType        =   2
            DataField       =   "MqmLmsSal"
            IdName          =   "MQMLMSSAL"
         End
         Begin MwfCtl.MWCTL_MED ATXT 
            Height          =   285
            Index           =   13
            Left            =   3120
            TabIndex        =   12
            ToolTipText     =   "Enter Lmg Sales"
            Top             =   945
            Width           =   1050
            _ExtentX        =   1852
            _ExtentY        =   503
            Alignment       =   2
            Mask            =   "####0.000"
            MaxLength       =   9
            DataType        =   2
            DataField       =   "MqmLmlSal"
            IdName          =   "MQMLMLSAL"
         End
         Begin MwfCtl.MWCTL_MED ATXT 
            Height          =   285
            Index           =   15
            Left            =   1020
            TabIndex        =   22
            ToolTipText     =   "Enter Metal Grade"
            Top             =   1815
            Width           =   1320
            _ExtentX        =   2328
            _ExtentY        =   503
            MaxLength       =   8
            DataType        =   4
            DataField       =   "MqmMetGrd1"
            IdName          =   "MQMMETGRD1"
         End
         Begin MwfCtl.MWCTL_MED ATXT 
            Height          =   285
            Index           =   17
            Left            =   2340
            TabIndex        =   23
            ToolTipText     =   "Enter Metal Grade"
            Top             =   1815
            Width           =   1320
            _ExtentX        =   2328
            _ExtentY        =   503
            MaxLength       =   8
            DataType        =   4
            DataField       =   "MqmMetGrd2"
            IdName          =   "MQMMETGRD2"
         End
         Begin MwfCtl.MWCTL_MED ATXT 
            Height          =   285
            Index           =   18
            Left            =   3660
            TabIndex        =   24
            ToolTipText     =   "Enter Metal Grade"
            Top             =   1815
            Width           =   1320
            _ExtentX        =   2328
            _ExtentY        =   503
            MaxLength       =   8
            DataType        =   4
            DataField       =   "MqmMetGrd3"
            IdName          =   "MQMMETGRD3"
         End
         Begin MwfCtl.MWCTL_MED ATXT 
            Height          =   285
            Index           =   19
            Left            =   4980
            TabIndex        =   25
            ToolTipText     =   "Enter Metal Grade"
            Top             =   1815
            Width           =   1320
            _ExtentX        =   2328
            _ExtentY        =   503
            MaxLength       =   8
            DataType        =   4
            DataField       =   "MqmMetGrd4"
            IdName          =   "MQMMETGRD4"
         End
         Begin MwfCtl.MWCTL_MED ATXT 
            Height          =   285
            Index           =   20
            Left            =   1020
            TabIndex        =   14
            ToolTipText     =   "Enter Metal Grade"
            Top             =   1530
            Width           =   1320
            _ExtentX        =   2328
            _ExtentY        =   503
            MaxLength       =   8
            DataType        =   4
            DataField       =   "MqmDiaGrd1"
            IdName          =   "MQMDIAGRD1"
         End
         Begin MwfCtl.MWCTL_MED ATXT 
            Height          =   285
            Index           =   21
            Left            =   2340
            TabIndex        =   15
            ToolTipText     =   "Enter Metal Grade"
            Top             =   1530
            Width           =   1320
            _ExtentX        =   2328
            _ExtentY        =   503
            MaxLength       =   8
            DataType        =   4
            DataField       =   "MqmDiaGrd2"
            IdName          =   "MQMDIAGRD2"
         End
         Begin MwfCtl.MWCTL_MED ATXT 
            Height          =   285
            Index           =   22
            Left            =   3660
            TabIndex        =   16
            ToolTipText     =   "Enter Metal Grade"
            Top             =   1530
            Width           =   1320
            _ExtentX        =   2328
            _ExtentY        =   503
            MaxLength       =   8
            DataType        =   4
            DataField       =   "MqmDiaGrd3"
            IdName          =   "MQMDIAGRD3"
         End
         Begin MwfCtl.MWCTL_MED ATXT 
            Height          =   285
            Index           =   23
            Left            =   4980
            TabIndex        =   17
            ToolTipText     =   "Enter Metal Grade"
            Top             =   1530
            Width           =   1320
            _ExtentX        =   2328
            _ExtentY        =   503
            MaxLength       =   8
            DataType        =   4
            DataField       =   "MqmDiaGrd4"
            IdName          =   "MQMDIAGRD4"
         End
         Begin MwfCtl.MWCTL_MED ATXT 
            Height          =   285
            Index           =   24
            Left            =   6300
            TabIndex        =   18
            ToolTipText     =   "Enter Metal Grade"
            Top             =   1530
            Width           =   1320
            _ExtentX        =   2328
            _ExtentY        =   503
            MaxLength       =   8
            DataType        =   4
            DataField       =   "MqmDiaGrd5"
            IdName          =   "MQMDIAGRD5"
         End
         Begin MwfCtl.MWCTL_MED ATXT 
            Height          =   285
            Index           =   25
            Left            =   7620
            TabIndex        =   19
            ToolTipText     =   "Enter Metal Grade"
            Top             =   1530
            Width           =   1320
            _ExtentX        =   2328
            _ExtentY        =   503
            MaxLength       =   8
            DataType        =   4
            DataField       =   "MqmDiaGrd6"
            IdName          =   "MQMDIAGRD6"
         End
         Begin MwfCtl.MWCTL_MED ATXT 
            Height          =   285
            Index           =   26
            Left            =   8940
            TabIndex        =   20
            ToolTipText     =   "Enter Metal Grade"
            Top             =   1530
            Width           =   1320
            _ExtentX        =   2328
            _ExtentY        =   503
            MaxLength       =   8
            DataType        =   4
            DataField       =   "MqmDiaGrd7"
            IdName          =   "MQMDIAGRD7"
         End
         Begin MwfCtl.MWCTL_MED ATXT 
            Height          =   285
            Index           =   27
            Left            =   10260
            TabIndex        =   21
            ToolTipText     =   "Enter Metal Grade"
            Top             =   1530
            Width           =   1320
            _ExtentX        =   2328
            _ExtentY        =   503
            MaxLength       =   8
            DataType        =   4
            DataField       =   "MqmDiaGrd8"
            IdName          =   "MQMDIAGRD8"
         End
         Begin MwfCtl.MWCTL_MED ATXT 
            Height          =   285
            Index           =   28
            Left            =   3090
            TabIndex        =   28
            ToolTipText     =   "Enter "
            Top             =   360
            Width           =   4560
            _ExtentX        =   8043
            _ExtentY        =   503
            Enabled         =   0   'False
            MaxLength       =   45
            DataType        =   4
            ReCalcOn        =   "MQMCMCD"
            IdName          =   "WCMNAME"
         End
         Begin MwfCtl.MWCTL_MED ATXT 
            Height          =   285
            Index           =   29
            Left            =   10170
            TabIndex        =   6
            ToolTipText     =   "Enter Price List Customer From Whom RM Rates Have To Be Looked Up For This Customer"
            Top             =   360
            Width           =   1410
            _ExtentX        =   2487
            _ExtentY        =   503
            MaxLength       =   8
            DataType        =   4
            DataField       =   "MqmLkUpRmRt"
            ReCalcParent    =   "MQMLMGSAL"
            IdName          =   "MQMLKUPRMRT"
         End
         Begin MwfCtl.MWCTL_MED ATXT 
            Height          =   285
            Index           =   257
            Left            =   10170
            TabIndex        =   10
            ToolTipText     =   "Enter Price List Customer From Whom Lab Rates Have To Be Looked Up For This Customer"
            Top             =   660
            Width           =   1410
            _ExtentX        =   2487
            _ExtentY        =   503
            MaxLength       =   8
            DataType        =   4
            DataField       =   "MqmLkUpLabRt"
            IdName          =   "MQMLKUPLABRT"
         End
         Begin MwfCtl.MWCTL_MED ATXT 
            Height          =   285
            Index           =   259
            Left            =   10170
            TabIndex        =   13
            ToolTipText     =   "Enter Price List Customer From Whom Metal Loss % Have To Be Looked Up For This Customer"
            Top             =   945
            Width           =   1410
            _ExtentX        =   2487
            _ExtentY        =   503
            MaxLength       =   8
            DataType        =   4
            DataField       =   "MqmLkUpMetLs"
            IdName          =   "MQMLKUPMETLS"
         End
         Begin MwfCtl.MWCTL_MED ATXT 
            Height          =   285
            Index           =   35
            Left            =   6615
            TabIndex        =   95
            ToolTipText     =   "Enter Currency"
            Top             =   75
            Width           =   1035
            _ExtentX        =   1826
            _ExtentY        =   503
            Enabled         =   0   'False
            MaxLength       =   8
            DataType        =   4
            DataField       =   "MqmCmCurCd"
            ReCalcParent    =   "MQMLMGSAL,MQMCNVFCT,MQMMULDIV"
            ReCalcOn        =   "MQMCMCD"
            IdName          =   "MQMCMCURCD"
         End
         Begin MwfCtl.MWCTL_MED ATXT 
            Height          =   285
            Index           =   36
            Left            =   11265
            TabIndex        =   27
            ToolTipText     =   "Enter 'Y' to Display Pic in Excel File"
            Top             =   1815
            Visible         =   0   'False
            Width           =   315
            _ExtentX        =   556
            _ExtentY        =   503
            DataType        =   4
            IdName          =   "WSHOWEXCPIC"
         End
         Begin MwfCtl.MWCTL_MED ATXT 
            Height          =   285
            Index           =   41
            Left            =   7350
            TabIndex        =   26
            ToolTipText     =   "Enter Reference Voucher"
            Top             =   1815
            Width           =   2070
            _ExtentX        =   3651
            _ExtentY        =   503
            MaxLength       =   20
            DataType        =   4
            ReCalcParent    =   "MQMREFVCHIDNO"
            ReCalcOn        =   "MQMREFVCHIDNO"
            IdName          =   "WREFVCH"
         End
         Begin MwfCtl.MWCTL_MED ATXT 
            Height          =   285
            Index           =   42
            Left            =   11700
            TabIndex        =   104
            ToolTipText     =   "Enter Number"
            Top             =   1830
            Visible         =   0   'False
            Width           =   750
            _ExtentX        =   1323
            _ExtentY        =   503
            Alignment       =   2
            Enabled         =   0   'False
            Mask            =   "#####0"
            MaxLength       =   6
            DataType        =   1
            DataField       =   "MqmRefVchIdNo"
            ReCalcParent    =   "WREFVCH"
            ReCalcOn        =   "WREFVCH"
            IdName          =   "MQMREFVCHIDNO"
         End
         Begin MwfCtl.MWCTL_MED ATXT 
            Height          =   285
            Index           =   44
            Left            =   6840
            TabIndex        =   9
            ToolTipText     =   "Enter Multiply By Factor"
            Top             =   660
            Width           =   810
            _ExtentX        =   1429
            _ExtentY        =   503
            Alignment       =   2
            Mask            =   "#0.000"
            MaxLength       =   6
            DataType        =   2
            DataField       =   "MqmMulBy"
            ReCalcOn        =   "MQMCMCD"
            IdName          =   "MQMMULBY"
         End
         Begin MwfCtl.MWCTL_MED ATXT 
            Height          =   285
            Index           =   49
            Left            =   9510
            TabIndex        =   121
            ToolTipText     =   "Enter Conversion Factor"
            Top             =   75
            Width           =   1035
            _ExtentX        =   1826
            _ExtentY        =   503
            Enabled         =   0   'False
            Mask            =   "###0.0000"
            MaxLength       =   9
            DataType        =   2
            DataField       =   "MqmCnvFct"
            ReCalcOn        =   "MQMDT,MQMCMCURCD"
            IdName          =   "MQMCNVFCT"
         End
         Begin MwfCtl.MWCTL_MED ATXT 
            Height          =   285
            Index           =   50
            Left            =   10695
            TabIndex        =   122
            ToolTipText     =   "Enter Conversion Factor Date"
            Top             =   75
            Width           =   885
            _ExtentX        =   1561
            _ExtentY        =   503
            Enabled         =   0   'False
            MaxLength       =   8
            DataType        =   3
            DataField       =   "MqmCnvFctDt"
            IdName          =   "MQMCNVFCTDT"
         End
         Begin MwfCtl.MWCTL_MED ATXT 
            Height          =   285
            Index           =   51
            Left            =   12480
            TabIndex        =   123
            ToolTipText     =   "Enter"
            Top             =   75
            Width           =   345
            _ExtentX        =   609
            _ExtentY        =   503
            Enabled         =   0   'False
            DataType        =   4
            DataField       =   "MqmMulDiv"
            ReCalcOn        =   "MQMCMCURCD"
            IdName          =   "MQMMULDIV"
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
            Left            =   10590
            TabIndex        =   126
            Top             =   75
            Width           =   75
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
            Index           =   45
            Left            =   11655
            TabIndex        =   125
            Top             =   75
            Width           =   825
         End
         Begin VB.Label ALBL 
            BackStyle       =   0  'Transparent
            Caption         =   "Conv Fact / Date"
            BeginProperty Font 
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
            Left            =   7890
            TabIndex        =   124
            Top             =   75
            Width           =   1605
         End
         Begin VB.Label ALBL 
            BackStyle       =   0  'Transparent
            Caption         =   "Multiply By"
            BeginProperty Font 
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
            Left            =   5640
            TabIndex        =   107
            Top             =   660
            Width           =   1095
         End
         Begin VB.Label ALBL 
            BackStyle       =   0  'Transparent
            Caption         =   "Refer  Vch"
            BeginProperty Font 
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
            Left            =   6330
            TabIndex        =   103
            Top             =   1815
            Width           =   1065
         End
         Begin VB.Label ALBL 
            BackStyle       =   0  'Transparent
            Caption         =   "Show Pic in Excel"
            BeginProperty Font 
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
            Left            =   9450
            TabIndex        =   98
            Top             =   1815
            Visible         =   0   'False
            Width           =   1725
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
            Index           =   29
            Left            =   5700
            TabIndex        =   96
            Top             =   75
            Width           =   915
         End
         Begin VB.Label ALBL 
            BackStyle       =   0  'Transparent
            Caption         =   "Metal Loss Look Up"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Index           =   210
            Left            =   7890
            TabIndex        =   75
            Top             =   945
            Width           =   1935
         End
         Begin VB.Label ALBL 
            BackStyle       =   0  'Transparent
            Caption         =   "Lab Rate Look Up"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Index           =   211
            Left            =   7890
            TabIndex        =   74
            Top             =   660
            Width           =   1785
         End
         Begin VB.Label ALBL 
            BackStyle       =   0  'Transparent
            Caption         =   "Rm Rate Look Up"
            BeginProperty Font 
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
            Left            =   7890
            TabIndex        =   73
            Top             =   360
            Width           =   1785
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
            Index           =   26
            Left            =   2460
            TabIndex        =   72
            Top             =   360
            Width           =   645
         End
         Begin VB.Label ALBL 
            BackStyle       =   0  'Transparent
            Caption         =   "Grade 1"
            BeginProperty Font 
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
            Left            =   1020
            TabIndex        =   63
            Top             =   1245
            Width           =   1320
         End
         Begin VB.Label ALBL 
            BackStyle       =   0  'Transparent
            Caption         =   "Grade 8"
            BeginProperty Font 
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
            Left            =   10260
            TabIndex        =   62
            Top             =   1260
            Width           =   1320
         End
         Begin VB.Label ALBL 
            BackStyle       =   0  'Transparent
            Caption         =   "Grade 7"
            BeginProperty Font 
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
            Left            =   8940
            TabIndex        =   61
            Top             =   1245
            Width           =   1320
         End
         Begin VB.Label ALBL 
            BackStyle       =   0  'Transparent
            Caption         =   "Grade 6"
            BeginProperty Font 
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
            Left            =   7620
            TabIndex        =   60
            Top             =   1230
            Width           =   1320
         End
         Begin VB.Label ALBL 
            BackStyle       =   0  'Transparent
            Caption         =   "Grade 5"
            BeginProperty Font 
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
            Left            =   6300
            TabIndex        =   59
            Top             =   1245
            Width           =   1320
         End
         Begin VB.Label ALBL 
            BackStyle       =   0  'Transparent
            Caption         =   "Grade 4"
            BeginProperty Font 
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
            Left            =   4980
            TabIndex        =   58
            Top             =   1245
            Width           =   1320
         End
         Begin VB.Label ALBL 
            BackStyle       =   0  'Transparent
            Caption         =   "Grade 3"
            BeginProperty Font 
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
            Left            =   3660
            TabIndex        =   57
            Top             =   1245
            Width           =   1320
         End
         Begin VB.Label ALBL 
            BackStyle       =   0  'Transparent
            Caption         =   "Grade 2"
            BeginProperty Font 
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
            Left            =   2340
            TabIndex        =   56
            Top             =   1245
            Width           =   1320
         End
         Begin VB.Label ALBL 
            BackStyle       =   0  'Transparent
            Caption         =   "Diamond"
            BeginProperty Font 
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
            Left            =   60
            TabIndex        =   55
            Top             =   1530
            Width           =   1125
         End
         Begin VB.Label ALBL 
            BackStyle       =   0  'Transparent
            Caption         =   "Metal"
            BeginProperty Font 
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
            Left            =   60
            TabIndex        =   54
            Top             =   1815
            Width           =   1155
         End
         Begin VB.Label LblLmlSal 
            BackStyle       =   0  'Transparent
            Caption         =   "LML"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Left            =   2490
            TabIndex        =   53
            Top             =   945
            Width           =   645
         End
         Begin VB.Label ALBL 
            BackStyle       =   0  'Transparent
            Caption         =   "LMP"
            BeginProperty Font 
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
            Left            =   2490
            TabIndex        =   52
            Top             =   660
            Width           =   645
         End
         Begin VB.Label LblLmsSal 
            BackStyle       =   0  'Transparent
            Caption         =   "LMS"
            BeginProperty Font 
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
            TabIndex        =   51
            Top             =   945
            Width           =   645
         End
         Begin VB.Label ALBL 
            BackStyle       =   0  'Transparent
            Caption         =   "LMG"
            BeginProperty Font 
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
            Left            =   60
            TabIndex        =   50
            Top             =   660
            Width           =   435
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
            Index           =   8
            Left            =   4080
            TabIndex        =   49
            Top             =   60
            Width           =   435
         End
         Begin VB.Label ALBL 
            BackStyle       =   0  'Transparent
            Caption         =   "Cust Cd"
            BeginProperty Font 
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
            Left            =   60
            TabIndex        =   48
            Top             =   360
            Width           =   825
         End
      End
      Begin VB.Frame FraCpyVch 
         Height          =   1785
         Left            =   5175
         TabIndex        =   46
         Top             =   7380
         Visible         =   0   'False
         Width           =   4785
         Begin MwfCtl.MWCTL_BTN1 CmdCpyVchGo 
            Height          =   495
            Left            =   3990
            TabIndex        =   45
            ToolTipText     =   "Press This Button To Copy"
            Top             =   1125
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
            Index           =   0
            Left            =   2160
            TabIndex        =   40
            ToolTipText     =   "Enter Voucher Chr"
            Top             =   795
            Width           =   570
            _ExtentX        =   1005
            _ExtentY        =   503
            MaxLength       =   3
            DataType        =   4
            ReCalcParent    =   "WCPYVCHODNO"
            IdName          =   "WCPYVCHODCHR"
         End
         Begin MwfCtl.MWCTL_MED ATXT 
            Height          =   285
            Index           =   1
            Left            =   1725
            TabIndex        =   39
            ToolTipText     =   "Enter Vocuher YY"
            Top             =   795
            Width           =   330
            _ExtentX        =   582
            _ExtentY        =   503
            MaxLength       =   2
            DataType        =   4
            ReCalcParent    =   "WCPYVCHODNO"
            IdName          =   "WCPYVCHODYY"
         End
         Begin MwfCtl.MWCTL_MED ATXT 
            Height          =   285
            Index           =   5
            Left            =   2880
            TabIndex        =   41
            ToolTipText     =   "Enter Voucher Number"
            Top             =   795
            Width           =   870
            _ExtentX        =   1535
            _ExtentY        =   503
            Alignment       =   2
            Mask            =   "#####0"
            MaxLength       =   6
            DataType        =   1
            ReCalcParent    =   "WCPYVCHODSRFR"
            ReCalcOn        =   "WCPYVCHCOCD,WCPYVCHODTC,WCPYVCHODYY,WCPYVCHODCHR"
            IdName          =   "WCPYVCHODNO"
         End
         Begin MwfCtl.MWCTL_MED ATXT 
            Height          =   285
            Index           =   6
            Left            =   1020
            TabIndex        =   42
            ToolTipText     =   "Enter From Voucher Sr"
            Top             =   1080
            Width           =   900
            _ExtentX        =   1588
            _ExtentY        =   503
            Alignment       =   2
            Mask            =   "#####0"
            MaxLength       =   6
            DataType        =   1
            ReCalcOn        =   "WCPYVCHODNO"
            IdName          =   "WCPYVCHODSRFR"
         End
         Begin MwfCtl.MWCTL_MED ATXT 
            Height          =   285
            Index           =   14
            Left            =   2880
            TabIndex        =   43
            ToolTipText     =   "Enter To Voucher Sr"
            Top             =   1080
            Width           =   870
            _ExtentX        =   1535
            _ExtentY        =   503
            Alignment       =   2
            Mask            =   "#####0"
            MaxLength       =   6
            DataType        =   1
            IdName          =   "WCPYVCHODSRTO"
         End
         Begin MwfCtl.MWCTL_MED ATXT 
            Height          =   285
            Index           =   3
            Left            =   1020
            TabIndex        =   38
            ToolTipText     =   "Enter Source Voucher Tc"
            Top             =   795
            Width           =   570
            _ExtentX        =   1005
            _ExtentY        =   503
            MaxLength       =   8
            DataType        =   4
            ReCalcParent    =   "WCPYVCHODNO"
            IdName          =   "WCPYVCHODTC"
         End
         Begin MwfCtl.MWCTL_MED ATXT 
            Height          =   285
            Index           =   34
            Left            =   1020
            TabIndex        =   44
            ToolTipText     =   "Enter 'Y' To Copy Rates from Source"
            Top             =   1365
            Width           =   315
            _ExtentX        =   556
            _ExtentY        =   503
            DataType        =   4
            IdName          =   "WCPYVCHRT"
         End
         Begin MwfCtl.MWCTL_MED ATXT 
            Height          =   285
            Index           =   2
            Left            =   1020
            TabIndex        =   37
            ToolTipText     =   "Enter Source Company Code"
            Top             =   510
            Width           =   555
            _ExtentX        =   979
            _ExtentY        =   503
            MaxLength       =   3
            DataType        =   4
            ReCalcParent    =   "WCPYVCHODNO"
            IdName          =   "WCPYVCHCOCD"
         End
         Begin VB.Label ALBL 
            BackStyle       =   0  'Transparent
            Caption         =   "Copy Rt"
            BeginProperty Font 
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
            Left            =   90
            TabIndex        =   71
            Top             =   1365
            Width           =   855
         End
         Begin VB.Label ALBL 
            Alignment       =   1  'Right Justify
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
            Index           =   6
            Left            =   1980
            TabIndex        =   70
            Top             =   1080
            Width           =   825
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
            Index           =   5
            Left            =   90
            TabIndex        =   69
            Top             =   510
            Width           =   855
         End
         Begin VB.Label ALBL 
            BackStyle       =   0  'Transparent
            Caption         =   "Vch No"
            BeginProperty Font 
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
            TabIndex        =   68
            Top             =   795
            Width           =   855
         End
         Begin VB.Label ALBL 
            BackStyle       =   0  'Transparent
            Caption         =   "Vch Sr Fr"
            BeginProperty Font 
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
            Left            =   90
            TabIndex        =   67
            Top             =   1080
            Width           =   915
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
            Left            =   2775
            TabIndex        =   66
            Top             =   810
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
            Index           =   0
            Left            =   1590
            TabIndex        =   65
            Top             =   810
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
            Index           =   58
            Left            =   2055
            TabIndex        =   64
            Top             =   810
            Width           =   105
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
            Height          =   300
            Index           =   1
            Left            =   -30
            TabIndex        =   47
            Top             =   90
            Width           =   6765
         End
      End
   End
End
Attribute VB_Name = "EmrFrmMultiPrcQt"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim moCn As MwfLib.MDOConnection   '*** (Bef speed) bef 24/11/06
Private Enum en_MpQtFra
  CpyVch = 0
  CpyDsg = 1
  Excel = 2
  DsgDet = 3
  CpyRfId = 4
End Enum

Dim ms_TcTyp As String, ms_CoCd As String, ms_Yy As String, ms_Tc As String, ms_Chr As String

Dim wPic3DHt As Single
Dim wPic3DWdt As Single
Dim wFraPic3DHt As Single
Dim wFraPic3DWdt As Single
Dim wFraPic3DLft As Single
Dim wFraPic3DTop As Single
Dim w3DPth As String

Dim oExcel, oBook, oSheet

Private Const m_xlNone = &HFFFFEFD2
Private Const m_xlEdgeLeft = 7
Private Const m_xlEdgeBottom = 9
Private Const m_xlEdgeTop = 8
Private Const m_xlEdgeRight = 10
Private Const m_xlInsideVertical = 11
Private Const m_xlInsideHorizontal = 12
Private Const m_xlThin = 2
Private Const m_xlMedium = &HFFFFEFD6

'*** Arr_Cv stores the causes validation state, Arr_Ed the Ena/Disa state of Cmd Buttons
Dim mArr_Cv As Variant, mArr_Ed As Variant
Dim mi_LabGrdLeft As Integer

Private Sub adc_AutoGen()
  
  If ADC("MqmNo") = 0 Then
    ADC("MqmNo") = moCn.AutoGen("MultiPrcQtMst", "MqmNo", ADC("MqmCoCd") + ADC("MqmTc") + ADC("MqmYy") + ADC("MqmChr"))
  Else
    Call moCn.AutoGen("MultiPrcQtMst", "MqmNo", ADC("MqmCoCd") + ADC("MqmTc") + ADC("MqmYy") + ADC("MqmChr"), ADC("MqmNo"))
  End If

End Sub

Private Sub ADC_KeyValid(pr_Cancel As Boolean, pr_ErrMsg As String)
  Call ATXT_FldChange(99, "MqmYy", pr_Cancel, pr_ErrMsg, ADC("MqmYy"))
  If pr_Cancel = True Then Exit Sub
  Call ATXT_FldChange(99, "MqmTc", pr_Cancel, pr_ErrMsg, ADC("MqmTc"))
  If pr_Cancel = True Then Exit Sub
  pr_Cancel = (Not moCn.RecSeek("Select vPSCd from vParam where " + _
                         "vPCoCd= '" + ADC("MqmCoCd") + "' and vPTyp= 'CHR' and " + _
                         "vPMCd = '" + ADC("MqmTc") + "' and vPSCd='" + ADC("MqmChr") + "'"))
  If pr_Cancel = True Then pr_ErrMsg = "Invalid Voucher Character": Exit Sub
  Call ATXT_FldChange(99, "MqmChr", pr_Cancel, pr_ErrMsg, ADC("MqmChr"))
  If pr_Cancel = True Then pr_ErrMsg = "User Has No Right To Access This Voucher Character": Exit Sub
  
End Sub

Private Sub ADC_KeyWhen()
  If ms_CoCd <> "" Then ADC("MqmCoCd") = ms_CoCd
  If ms_Yy <> "" Then ADC("MqmYy") = ms_Yy
  If ms_Tc <> "" Then ADC("MqmTc") = ms_Tc
  If ms_Chr <> "" Then ADC("MqmChr") = ms_Chr
  
  Call EnaDisaCmds(True)
  
  Call HideAllFras
  
  Call ShowPic(0, 0, "", Pic3D)

  FraPic3D.Height = wFraPic3DHt
  FraPic3D.Width = wFraPic3DWdt
    
  FraPic3D.Left = wFraPic3DLft
  FraPic3D.Top = wFraPic3DTop

End Sub
Private Sub ADC_Load()
  '*** Set the First Key Control property to the field cuUsrCd
  '*** Set the First Non Key Control property to the grid cuUsrName
  '*** Set the Child Property Of ADC to GrdGrade
  '*** Set the Previous and Next Control Properties for the Grid GrdGrade as ADC
  '*** Set the Hot Key for the Grid GrdGrade as 'A'

  ReDim mArr_Cv(4), mArr_Ed(4)

  Set moCn = ADC.Connection    '*** (Bef speed) bef 24/11/06
  Set ADC.FirKeyCtl = ADC("MqmChr")
  Set ADC.FirNKeyCtl = ADC("MqmDt")
  ADC.Child = "GrdMpQtDsg"
  Set GrdMpQtDsg.PrevCtl = ADC
  Set GrdMpQtDsg.NextCtl = ADC
  GrdMpQtDsg.HotKey = "A"
  
  wPic3DHt = Pic3D.Height
  wPic3DWdt = Pic3D.Width
  
  wFraPic3DHt = FraPic3D.Height
  wFraPic3DWdt = FraPic3D.Width
 
  wFraPic3DLft = FraPic3D.Left
  wFraPic3DTop = FraPic3D.Top
  
  ms_TcTyp = "QT"
  ms_CoCd = gs_CoCd
  ms_Tc = moCn.GetFldVal("Select PValue from Param where PTyp='TCTYP' and PMCd='" + ms_TcTyp + "'")
  ms_Yy = moCn.GetFldVal("Select hYy from Head where HCoCd= '" + ms_CoCd + "' and HCd='" + ctSelfCmCd + "'")
 
  ADC("MqmCmCurCd").Enabled = False: ADC("MqmCnvFct").Enabled = False: ADC("MqmCnvFctDt").Enabled = False
 
  If Not moCn.RecSeek("Select 'x' from Head where HCoCd = '" + gs_CoCd + "' and HCd = '" + ctSelfCmCd + "' and HSlvModYN='Y' ") Then
    LblLmsSal.Visible = False
    ADC("MqmLmsSal").Enabled = False: ADC("MqmLmsSal").Visible = False
  End If
  
  If Not moCn.RecSeek("Select 'x' from Head where HCoCd = '" + gs_CoCd + "' and HCd = '" + ctSelfCmCd + "' and HPdModYN='Y' ") Then
    LblLmlSal.Visible = False
    ADC("MqmLmlSal").Enabled = False: ADC("MqmLmlSal").Visible = False
  End If
  
  FraNKeyAll.ZOrder (0)
   
End Sub
Private Sub ADC_NonKeyValid(pr_Cancel As Boolean, pr_ErrMsg As String)
txtCpyDsgDmCd.text = ""
End Sub
Private Sub ADC_NonKeyWhen(pr_Cancel As Boolean, pr_ErrMsg As String)
      If ADC.IsNew And (Not moCn.RecSeek("Select vPSCd from vParam where " + _
                         "vPCoCd= '" + ADC("MqmCoCd") + "' and vPTyp= 'CHR' and " + _
                         "vPMCd = '" + ADC("MqmTc") + "' and vPSCd='" + ADC("MqmChr") + "' and vPValidYn<>'N'")) Then
       pr_Cancel = True: pr_ErrMsg = "Voucher Character No More Valid": ADC.AllowSave = False: Exit Sub
      Else
        ADC.AllowSave = True
      End If

  Call EnaDisaCmds(False)
  Dim mRsCustMst As MDORowSet
  
  If ADC.AddFndMode = xaddmode Then
    ADC("MqmDt") = GetNextDate(ADC("MqmCoCd"))   '6.1_50 moCn.SrvrDate removed
  End If
  
  ms_Yy = ADC("MqmYy")
  ms_Tc = ADC("MqmTc")
  ms_Chr = ADC("MqmChr")
  
  If ADC.AddFndMode = xFndMode Then
    If ADC("wRefVch") <> "" Then
      Dim wOdKeyArr() As String
      wOdKeyArr = Split(ADC("wRefVch"), "/")
      If UBound(wOdKeyArr) = 3 Then
        ADC("wCpyVchCoCd") = ADC("MqmCoCd")
        ADC("wCpyVchOdTc") = wOdKeyArr(0)
        ADC("wCpyVchOdYy") = wOdKeyArr(1)
        ADC("wCpyVchOdChr") = wOdKeyArr(2)
        ADC("wCpyVchOdNo") = wOdKeyArr(3)
      End If
    End If
  End If
  
  ADC("wShowExcPic") = "Y"
  GrdPropSet
  GrdMpQtDsg.AllowAdd = False
  
  GrdMpQtRm.AllowAdd = False:  GrdMpQtRm.AllowDelete = False
  GrdMpQtLab.AllowAdd = False:  GrdMpQtLab.AllowDelete = False
  CmdDsgDet.Enabled = False
  'vk.4 setting default values
  OptCpyBagFr(0).Value = True
  ADC("wCpyThruBagNo") = "Y"

End Sub

Private Sub ATXT_FldChange(Index As Integer, ByVal IdName As String, Cancel As Boolean, ErrMsg As String, ByVal pv_NewValue As Variant)
  Dim i As Integer, wGrdNo As Integer
  
  Select Case IdName
  Case Is = UCase("MqmTc")
    Cancel = Not moCn.RecSeek("select PMCd from Param where PTyp= 'TC' and " + _
                         "PMCd = '" + pv_NewValue + "' and " + _
                         "PValue= '" + ms_TcTyp + "' ")
    If Cancel = True Then ErrMsg = "Invalid Voucher Tc": Exit Sub
    
  Case Is = UCase("MqmYy")
    Cancel = Not moCn.RecSeek("select PMCd from Param where PTyp= 'YY' and " + _
                         "PMCd = '" + pv_NewValue + "'")
    If Cancel = True Then ErrMsg = "Invalid Voucher Year": Exit Sub
  Case Is = UCase("MqmChr")
    Cancel = pv_NewValue <> "" And (Not moCn.RecSeek("select vPSCd from vParam where " + _
                         "vPCoCd= '" + ADC("MqmCoCd") + "' and vPTyp= 'CHR' and " + _
                         "vPMCd = '" + ADC("mqmTc") + "' and vPSCd='" + pv_NewValue + "'"))
   If Cancel = True Then
        ErrMsg = "Invalid Voucher Character": Exit Sub
   Else
    Dim ws_UsrArr() As String, j As Integer
       ws_UsrArr = Split(moCn.GetFldVal("Select vPValue4 from vParam where vPCoCd= '" + ADC("MqmCoCd") + "' and " + _
                                          "vPTyp='CHR' and vPMCd='" + ADC("MqmTc") + "' and " + _
                                          "vPSCd= '" + pv_NewValue + "'"), ",")
    If IsArray(ws_UsrArr) Then
     For j = 0 To UBound(ws_UsrArr)
         Cancel = IIF(UCase(gs_UsrCd) = UCase(ws_UsrArr(j)), False, True)
           If Cancel = False Then: Exit For
     Next j
    End If
    If Cancel = True Then ErrMsg = "User Has No Right To Access This Voucher Character": Exit Sub
   End If
   
   
  Case Is = UCase("MqmDt")
    Cancel = (pv_NewValue = "")
    If Cancel = True Then ErrMsg = "Voucher Date Cannot Be Blank": Exit Sub
    '6.1_50 checking holiday for voucher
    If moCn.GetFldVal("Select hHolChkYN From Head where HCd='ZSELF' and HCoCd='" + ADC("MqmCoCd") + "'") = "Y" Then
        If IsHOLYN(ADC("MqmCoCd"), pv_NewValue) Then Cancel = True: ErrMsg = "Voucher Cannot be Added on Holidays": Exit Sub
    End If
  Case Is = UCase("MqmDiaGrd1"), UCase("MqmDiaGrd2"), UCase("MqmDiaGrd3"), UCase("MqmDiaGrd4"), _
            UCase("MqmDiaGrd5"), UCase("MqmDiaGrd6"), UCase("MqmDiaGrd7"), UCase("MqmDiaGrd8")
    Cancel = pv_NewValue <> "" And Not moCn.RecSeek("Select 'x' From Param where PTyp= 'GRDCD' and PMCd='DIA' and PSCd = '" + pv_NewValue + "' ")
    If Cancel Then ErrMsg = "Invalid Diamond Grade": Exit Sub
    If pv_NewValue <> "" Then
      For i = 1 To 8
        If UCase(IdName) <> UCase("MqmDiaGrd + Trim(CStr(i))") Then
          Cancel = pv_NewValue = ADC("MqmDiaGrd" + Trim(CStr(i)))
          If Cancel Then ErrMsg = "Diamond Grade Cannot Be Repeated": Exit Sub
        End If
      Next
    Else
      wGrdNo = Right(IdName, 1)
      If wGrdNo < 8 Then
        Cancel = (ADC("MqmDiaGrd" + Trim(CStr(wGrdNo + 1))) <> "")
        If Cancel Then ErrMsg = "Grade Cannoe Be Blank When Next Grade Entered": Exit Sub
      End If
    End If

    
  Case Is = UCase("MqmMetGrd1"), UCase("MqmMetGrd2"), UCase("MqmMetGrd3"), UCase("MqmMetGrd4")
    Cancel = pv_NewValue <> "" And Not moCn.RecSeek("Select 'x' From Param where PTyp= 'GRDCD' and PMCd='MET' and PSCd = '" + pv_NewValue + "' ")
    If Cancel Then ErrMsg = "Invalid Metal Grade": Exit Sub
    If pv_NewValue <> "" Then
      For i = 1 To 4
        If UCase(IdName) <> UCase("MqmMetGrd + Trim(CStr(i))") Then
          Cancel = pv_NewValue = ADC("MqmMetGrd" + Trim(CStr(i)))
          If Cancel Then ErrMsg = "Metal Grade Cannot Be Repeated": Exit Sub
        End If
      Next
    Else
      wGrdNo = Right(IdName, 1)
      If wGrdNo < 4 Then
        Cancel = (ADC("MqmMetGrd" + Trim(CStr(wGrdNo + 1))) <> "")
        If Cancel Then ErrMsg = "Grade Cannoe Be Blank When Next Grade Entered": Exit Sub
      End If
    End If
    
  
  Case Is = UCase("MqmCmCd")
    Cancel = Not moCn.RecSeek("Select 'x' From CustMst where CmCtg= 'C' " + _
                                "and CmCd = '" + pv_NewValue + "' and CmValidYn='Y' ")
    If Cancel Then ErrMsg = "Invalid Customer Code": Exit Sub
  
  Case Is = UCase("MqmMulBy")
    Cancel = (pv_NewValue <= 0)
    ErrMsg = "Multiplying Factor Should Be > Zero": Exit Sub
  
  Case Is = UCase("MqmLkUpRmRt"), UCase("OmLkUpLabRt"), UCase("OmLkUpMetLs")
    Cancel = pv_NewValue <> "" And Not moCn.RecSeek("Select 'x' From CustMst where CmCtg= 'P' " + _
                                "and CmCd = '" + pv_NewValue + "' and CmValidYn='Y' ")
    If Cancel Then ErrMsg = "Invalid Lookup Customer": Exit Sub
  
  Case Is = UCase("wRefVch")
    Dim wOdKeyArr() As String
    If ADC.Mode = xNorm Then
      If pv_NewValue <> "" Then
          wOdKeyArr = Split(pv_NewValue, "/")
          If UBound(wOdKeyArr) <> 3 Then Cancel = True: ErrMsg = "Invalid Reference Voucher": Exit Sub
          If Not IsNumeric(wOdKeyArr(3)) Then Cancel = True: ErrMsg = "Invalid Reference Voucher": Exit Sub
          Cancel = Not moCn.RecSeek("Select 'x' from OrdMst Where OmCoCd='" + ADC("MqmCoCd") + "' and OmTc='" + wOdKeyArr(0) + "' " + _
                                    " and OmYy='" + wOdKeyArr(1) + "' and OmChr='" + wOdKeyArr(2) + "' and OmNo=" + CStr(wOdKeyArr(3)) + " " + _
                                    " and OmPrtKey='" + ctCurrPrtn + "' ")
          If Cancel = True Then ErrMsg = "Invalid Reference Voucher": Exit Sub
      End If
    End If
    
  
  Case Is = UCase("wCpyVchRt"), UCase("wCpyDsgRt")
    Cancel = Not moCn.RecSeek("Select 'x' From Param where PTyp= 'YN' and PmCd = '" + pv_NewValue + "' ")
    If Cancel Then ErrMsg = "Invalid Copy Rates Option": Exit Sub
    
  Case Is = UCase("wCpyDsgDmCd")
    Cancel = pv_NewValue <> "" And Not moCn.RecSeek("Select 'x' From DsgMst where DmTcTyp= 'DM' and DmCd = '" + pv_NewValue + "' and DmSz='' ")
    If Cancel Then ErrMsg = "Invalid Design Code": Exit Sub
    
  Case Is = UCase("wCpyDsgDmSz")
    Cancel = pv_NewValue <> "" And Not moCn.RecSeek("Select 'x' From Param where PTyp= 'DMSZ' and PmCd = '" + pv_NewValue + "' ")
    If Cancel Then ErrMsg = "Invalid Design Size": Exit Sub
    
  Case Is = UCase("wShowExcPic")
    Cancel = Not moCn.RecSeek("Select 'x' From Param where PTyp= 'YN' and PmCd = '" + pv_NewValue + "' ")
    If Cancel Then ErrMsg = "Enter 'Y' to Display Pic in Excel Else Enter 'N'": Exit Sub
    
  Case Is = UCase("wCpyDsgBarCdYn")
    Cancel = Not moCn.RecSeek("Select 'x' From Param where PTyp= 'YN' and PmCd = '" + pv_NewValue + "' ")
    If Cancel Then ErrMsg = "Enter 'Y' to Copy Design Using Bar Code Else Enter 'N'": Exit Sub
    If pv_NewValue = "Y" Then
      ADC("wCpyDsgDmCd").Enabled = False: ADC("wCpyDsgSfx").Enabled = False
      ADC("wCpyDsgDmSz").Enabled = False: ADC("wCpyDsgSrc").Enabled = False
      txtCpyDsgDmCd.Enabled = True
      txtCpyDsgDmCd.SetFocus
    Else
      ADC("wCpyDsgDmCd").Enabled = True: ADC("wCpyDsgSfx").Enabled = True
      ADC("wCpyDsgDmSz").Enabled = True: ADC("wCpyDsgSrc").Enabled = True
    
      txtCpyDsgDmCd.Enabled = False
      If FraCpyDsg.Visible = True Then ADC("wCpyDsgDmCd").SetFocus
    End If
    
'**** Bhavna added cpy thru bag no
  Case Is = UCase("wCpyThruBagNo")
    If ADC.Mode = xNorm Then
      Cancel = Not moCn.RecSeek("select PMCd from Param where PTyp= 'YN' and PMCd = '" + pv_NewValue + "'")
      ErrMsg = "Enter 'Y' to Copy thru Bag No Or 'N' to Copy thru RfId No ": Exit Sub
    End If
      
  End Select
End Sub

Private Sub ATXT_FldWhen(Index As Integer, ByVal IdName As String, Cancel As Boolean, ErrMsg As String)
  '*** Set Helps For Each Field
  
  Dim i As Integer, wGrdNo As Integer
  If GrdMpQtDsg.Rows > 1 Then
    Select Case IdName
       Case Is = UCase("MqmCmCd"), UCase("MqmLkUpRmRt"), UCase("MqmLkUpLabRt"), UCase("MqmLkUpMetLs"), _
          UCase("MqmCmCurCd"), UCase("MqmCnvFct"), UCase("MqmCnvFctDt"), _
          UCase("MqmLmgSal"), UCase("MqmLmpSal"), UCase("MqmLmsSal"), UCase("MqmLmlSal"), UCase("MqmMulBy"), _
          UCase("MqmMetGrd1"), UCase("MqmMetGrd2"), UCase("MqmMetGrd3"), UCase("MqmMetGrd4"), _
          UCase("MqmDiaGrd1"), UCase("MqmDiaGrd2"), UCase("MqmDiaGrd3"), UCase("MqmDiaGrd4"), _
          UCase("MqmDiaGrd5"), UCase("MqmDiaGrd6"), UCase("MqmDiaGrd7"), UCase("MqmDiaGrd8")
      Cancel = True
      ErrMsg = "Cannot Edit Value when Recs exist in Design Box"
      Exit Sub
    End Select
  End If
  
  Select Case IdName
  Case Is = UCase("MqmYy")
    Call HlpList.PMCd("YY")
  
  Case Is = UCase("MqmChr")
    Call HlpList.vPSCd(ADC("MqmCoCd"), "CHR", ADC("MqmTc"), "Y")
  
  Case Is = UCase("MqmNo")
    Call HlpList.MultiPrcQtNo(ADC("MqmCoCd"), ADC("MqmTc"), ADC("MqmYy"), ADC("MqmChr"))
    
    
  Case Is = UCase("MqmCmCd")
    Call HlpList.CustCd("C")
    
  Case Is = UCase("MqmLkUpRmRt"), UCase("MqmLkUpLabRt"), UCase("MqmLkUpMetLs")
    Call HlpList.CustCd("P")
    
  Case Is = UCase("MqmCmCurCd")
    If ADC("MqmCmCd") <> ctSelfCmCd Then Cancel = True: ErrMsg = "Cannot Edit Currency Code For This Customer": Exit Sub
    Call HlpList.PMCd("CURNCY")
   
  Case Is = UCase("MqmMetGrd1"), UCase("MqmMetGrd2"), UCase("MqmMetGrd3"), UCase("MqmMetGrd4")
    wGrdNo = Right(IdName, 1)
    If wGrdNo > 1 Then
        Cancel = (ADC("MqmMetGrd" + Trim(CStr(wGrdNo - 1))) = "")
        If Cancel Then ErrMsg = "Cannot Enter When Earlier Grade Not Entered": Exit Sub
    End If
    Call HlpList.PSCd("GRDCD", "MET")
    
  Case Is = UCase("MqmDiaGrd1"), UCase("MqmDiaGrd2"), UCase("MqmDiaGrd3"), UCase("MqmDiaGrd4"), _
            UCase("MqmDiaGrd5"), UCase("MqmDiaGrd6"), UCase("MqmDiaGrd7"), UCase("MqmDiaGrd8")
    wGrdNo = Right(IdName, 1)
    If wGrdNo > 1 Then
      Cancel = (ADC("MqmDiaGrd" + Trim(CStr(wGrdNo - 1))) = "")
      If Cancel Then ErrMsg = "Cannot Enter When Earlier Grade Not Entered": Exit Sub
    End If
    Call HlpList.PSCd("GRDCD", "DIA")
    
  Case Is = UCase("wRefVch")
    Call HlpList.OrdNo(ADC("MqmCoCd"), ADC("MqmTc"))
    
    
  Case Is = UCase("wCpyDsgDmCd")
    Call HlpList.DmCd("DM")
    
  Case Is = UCase("wCpyDsgDmSz")
    Call HlpList.PMCd("D")
    
  Case Is = UCase("wCpyDsgSfx")
    Call HlpList.OdSfx(ctSelfCoCd, ADC("wCpyDsgDmCd"), "C", ADC("MqmCmCd"))
    
  Case Is = UCase("wCpyVchCoCd")
    Call HlpList.hCoCd
    
  Case Is = UCase("wCpyVchOdTc")
    Call HlpList.PMCd("TC", "'SO','QT'")
    
  Case Is = UCase("wCpyVchOdYy")
    Call HlpList.PMCd("YY")
    
  Case Is = UCase("wCpyVchOdChr")
    Call HlpList.vPSCd(ADC("wCpyVchCoCd"), "CHR", ADC("wCpyVchOdTc"))
    
  Case Is = UCase("wCpyVchOdNo")
    Call HlpList.OmNo(ADC("wCpyVchCoCd"), ADC("wCpyVchOdTc"), ADC("wCpyVchOdYy"), ADC("wCpyVchOdChr"), ADC("MqmCmCurCd"))
    
  Case Is = UCase("wCpyVchOdSrFr"), UCase("wCpyVchOdSrTo")
    If ADC("wCpyVchOdNo") = 0 Then Cancel = True: ErrMsg = "Cannot Edit When Source Order Not Entered": Exit Sub
    Call HlpList.OdSr(ADC("wCpyVchCoCd"), ADC("wCpyVchOdTc"), ADC("wCpyVchOdYy"), ADC("wCpyVchOdChr"), ADC("wCpyVchOdNo"))

  Case Is = UCase("wCpyVchRt"), UCase("wCpyDsgRt")
    Call HlpList.PMCd("YN")
  
  Case Is = UCase("wCpyDsgSrc")
    Call HlpList.PMCd("DSGCPY", , , " (PDesc225= '') ")
  
  Case Is = UCase("wShowExcPic")
    Call HlpList.PMCd("YN")
    
  Case Is = UCase("wCpyDsgBarCdYn")
    Call HlpList.PMCd("YN")
    
  Case Is = UCase("wMqdGldAs"), UCase("wMqdGldAsWt"), UCase("wMqdLabAs"), UCase("wMqdLabAsWt")
    Cancel = True: ErrMsg = "Cannot Edit": Exit Sub
    
     '*********Bhavna cpy thru bag no
    Case Is = UCase("wCPYTHRUBAGNO")
        If ADC("WCPYTHRUBAGNO") = "" Then
            ADC("WCPYTHRUBAGNO") = "N"
        End If
  
  End Select
End Sub

Private Sub atxt_ReCalculate(Index As Integer, ByVal IdName As String, ByVal OldVal As Variant, ByVal RowNum As Integer, ByVal SourceName As String, ByVal SourceColName As String)
  Dim ws_DtDt As String
  
  Select Case UCase(IdName)
    Case Is = UCase("MqmChr")
      If ADC.Mode = xNorm Then
        ADC("MqmChr") = moCn.GetFldVal("Select vPValue from vParam where " + _
                                  "vPCoCd= '" + ADC("MqmCoCd") + "' and vPTyp='DEFCHR' " + _
                                  "and vPMCd='" + ADC("MqmTc") + "'")
      End If

    Case Is = UCase("wCmName")
      If ADC.Mode = xNorm Or ADC.Mode = xInit Then
        ADC("wCmName") = moCn.GetFldVal("Select CmName From CustMst Where CmCtg='C' and CmCd='" + ADC("MqmCmCd") + "' ")
        ADC("MqmCmCurCd") = moCn.GetFldVal("Select CmCurCd From CustMst Where CmCtg='C' and CmCd='" + ADC("MqmCmCd") + "' ")
      End If
    
    Case Is = UCase("MqmCmCurCd")
      If ADC.Mode = xNorm Then
        ADC("MqmCmCurCd") = moCn.GetFldVal("Select CmCurCd From CustMst Where CmCtg='C' and CmCd='" + ADC("MqmCmCd") + "' ")
      End If
    
    Case Is = UCase("MqmCnvFct")
      If ADC.Mode = xNorm Then
        ADC("MqmCnvFct") = CnvRt(moCn.GetFldVal("Select HBaseCurCd from Head where HCd= '" + ctSelfCmCd + "' and HCoCd= '" + ADC("MqmCoCd") + "'"), _
                          ADC("MqmCmCurCd"), ADC("MqmDt"))
        Call GetDtVal("CNV", ADC("MqmDt"), , ws_DtDt, ADC("MqmCmCurCd"), moCn.GetFldVal("Select HBaseCurCd from Head where HCd= '" + ctSelfCmCd + "' and HCoCd= '" + ADC("MqmCoCd") + "'"))
        ADC("MqmCnvFctDt") = ws_DtDt
                                
      End If
    
    Case Is = UCase("MqmMulDiv")
      If ADC.Mode = xNorm Then
        ADC("MqmMulDiv") = MulDiv(moCn.GetFldVal("Select HBaseCurCd from Head where HCd= '" + ctSelfCmCd + "' and HCoCd= '" + ADC("MqmCoCd") + "'"), _
                          ADC("MqmCmCurCd"))
      End If
    
    Case Is = UCase("MqmMulBy")
      If ADC.Mode = xNorm Then
        ADC("MqmMulBy") = moCn.GetFldVal("Select CmMulBy From CustMst Where CmCtg='C' and CmCd='" + ADC("MqmCmCd") + "' ")
      End If
      
    Case Is = UCase("MqmLmgSal")
      Dim wLme As Double
      Dim wDtVal As Double, wDtDt As String
      If ADC.Mode = xNorm Then
        Call GetDtVal("LMG", ADC("MqmDt"), wDtVal, wDtDt, ADC("MqmCmCurCd"))
        '6.1-13 wLme is removed, because wLme is always 0 so MqmLmgSal always save wDtVal
        'ADC("MqmLmgSal") = wLme = GetLmgSal("LMG", ADC("MqmCmCd"), wDtVal, "C", ADC("MqmCmCurCd"), ADC("MqmLkUpRmRt"))
        ADC("MqmLmgSal") = GetLmgSal("LMG", ADC("MqmCmCd"), wDtVal, "C", ADC("MqmCmCurCd"), ADC("MqmLkUpRmRt"))
        If ADC("MqmLmgSal") = 0 Then ADC("MqmLmgSal") = wDtVal
      
        Call GetDtVal("LMP", ADC("MqmDt"), wDtVal, wDtDt, ADC("MqmCmCurCd"))
        ADC("MqmLmpSal") = GetLmgSal("LMP", ADC("MqmCmCd"), wDtVal, "C", ADC("MqmCmCurCd"), ADC("MqmLkUpRmRt"))
        If ADC("MqmLmpSal") = 0 Then ADC("MqmLmpSal") = wDtVal
        
        If ADC("MqmLmsSal").Visible = True Then
          Call GetDtVal("LMS", ADC("MqmDt"), wDtVal, wDtDt, ADC("MqmCmCurCd"))
          ADC("MqmLmsSal") = GetLmgSal("LMS", ADC("MqmCmCd"), wDtVal, "C", ADC("MqmCmCurCd"), ADC("MqmLkUpRmRt"))
          If ADC("MqmLmsSal") = 0 Then ADC("MqmLmsSal") = wDtVal
        Else
          ADC("MqmLmsSal") = 1
        End If
        
        If ADC("MqmLmlSal").Visible = True Then
          Call GetDtVal("LML", ADC("MqmDt"), wDtVal, wDtDt, ADC("MqmCmCurCd"))
          ADC("MqmLmlSal") = GetLmgSal("LML", ADC("MqmCmCd"), wDtVal, "C", ADC("MqmCmCurCd"), ADC("MqmLkUpRmRt"))
          If ADC("MqmLmlSal") = 0 Then ADC("MqmLmlSal") = wDtVal
        Else
          ADC("MqmLmlSal") = 1
        End If
      End If
      
    Case Is = UCase("MqmRefVchIdNo")
    Dim wOdKeyArr() As String
    If ADC.Mode = xNorm Then
        If ADC("wRefVch") <> "" Then
          wOdKeyArr = Split(ADC("wRefVch"), "/")
          If UBound(wOdKeyArr) = 3 Then
            ADC("MqmRefVchIdNo") = moCn.GetFldVal("Select OmIdNo from OrdMst Where OmCoCd='" + ADC("MqmCoCd") + "' and OmTc='" + wOdKeyArr(0) + "' " + _
                                            " and OmYy='" + wOdKeyArr(1) + "' and OmChr='" + wOdKeyArr(2) + "' and OmNo=" + CStr(wOdKeyArr(3)) + " " + _
                                            " and OmPrtKey='" + ctCurrPrtn + "' ")
            ADC("wCpyVchCoCd") = ADC("MqmCoCd")
            ADC("wCpyVchOdTc") = wOdKeyArr(0)
            ADC("wCpyVchOdYy") = wOdKeyArr(1)
            ADC("wCpyVchOdChr") = wOdKeyArr(2)
            ADC("wCpyVchOdNo") = wOdKeyArr(3)
          End If
        Else
          ADC("MqmRefVchIdNo") = 0
        End If
    End If
     
    Case Is = UCase("wRefVch")
    If ADC.Mode = xInit Then
        If ADC("MqmRefVchIdNo") <> 0 Then
          ADC("wRefVch") = moCn.GetFldVal("Select OmTc+'/'+OMYy+'/'+OmChr+'/'+Ltrim(Str(OmNo)) " + _
                                          " from OrdMst Where OmIdNo=" + CStr(ADC("MqmRefVchIdNo")) + " " + _
                                          " and OmPrtKey='" + ctCurrPrtn + "' ")
        End If
    End If

      
    Case Is = UCase("wCpyDsgSrc")
      Dim ws_Source As String
      If ADC.Mode = xNorm Then
        If ADC("wCpyDsgDmCd") = "" Then
          ws_Source = ""
        Else
          If ws_Source = "" Then ws_Source = moCn.GetFldVal("Select '" + ctDsgCpyCust + "' " + _
                     "from OrdMst join OrdDsg " + _
                     "on OdOmIdNo=OmIdNo And OdCoCd= OmCoCd and OdTc= OmTc and OdYy= OmYy and OdChr= OmChr " + _
                     "and OdNo= OmNo and OdPrtKey=OmPrtKey where OmCoCd= '" + ctSelfCoCd + "' and " + _
                     "(Select PValue from Param where PTyp='TC' and PMCd=OmTc)='PL' and OdDmCd='" + ADC("wCpyDsgDmCd") + "' and " + _
                     "OdDmSz='" + ADC("wCpyDsgDmSz") + "' and OdSfx='" + ADC("wCpyDsgSfx") + "' and OmCmCtg='C' and " + _
                     "OmCmCd='" + ADC("MqmCmCd") + "' and OmCmCurCd= '" + ADC("MqmCmCurCd") + "' and OmPrtKey='" + ctCurrPrtn + "'")
    
          If ws_Source = "" Then ws_Source = moCn.GetFldVal("Select '" + ctDsgCpySelf + "' " + _
                    "from OrdMst join OrdDsg " + _
                    "on OdOmIdNo=OmIdNo And OdCoCd= OmCoCd and OdTc=OmTc and OdYy=OmYy and OdChr=OmChr " + _
                    "and OdNo=OmNo and OdPrtKey=OmPrtKey where OmCoCd= '" + ctSelfCoCd + "' and " + _
                    "(Select PValue from Param where PTyp='TC' and PMCd=OmTc)='PL' and " + _
                    "OdDmCd='" + ADC("wCpyDsgDmCd") + "' and " + _
                    "OdDmSz='" + ADC("wCpyDsgDmSz") + "' and OdSfx='" + ADC("wCpyDsgSfx") + "' and OmCmCtg='C' and " + _
                    "OmCmCd='" + ctSelfCmCd + "' and OmCmCurCd= '" + ADC("MqmCmCurCd") + "' and OmPrtKey='" + ctCurrPrtn + "'")
                                        
          If ws_Source = "" Or ADC("wCpyDsgBarCdYn") = "Y" Then ws_Source = moCn.GetFldVal("Select '" + ctDsgCpyMst + "' " + _
                       "from DsgMst where DmTcTyp='DM' and DmCd='" + ADC("wCpyDsgDmCd") + "' " + _
                       "and (DmSz='" + ADC("wCpyDsgDmSz") + "' or DmDefSz='" + ADC("wCpyDsgDmSz") + "') and DmPrtKey='" + ctCurrPrtn + "'")
                       
                                 
        End If
        ADC("wCpyDsgSrc") = ws_Source
      End If
      
    Case Is = UCase("wCpyDsgSrcVch")
      Dim ws_Src As String
      If ADC.Mode = xNorm Then
        'moCn.GetFldVal ("Select *")
        If ADC("wCpyDsgDmCd") <> "" Then
          ws_Src = IIF(UCase(ADC("wCpyDsgSrc")) = UCase(ctDsgCpyCust), ADC("MqmCmCd"), ctSelfCmCd)
          ADC("wCpyDsgSrcVch") = moCn.GetFldVal("Select OdTc+'/'+OdYy+'/'+OdChr+'/'+" + _
                              "LTrim(Str(OdNo))+'/'+LTrim(Str(OdSr)) from OrdMst join OrdDsg " + _
                              "on OdOmIdNo=OmIdNo And OdCoCd=OmCoCd and OdTc=OmTc and OdYy=OmYy and OdChr=OmChr and " + _
                              "OdNo= OmNo and OdPrtKey=OmPrtKey where OmCoCd= '" + ctSelfCoCd + "' and (Select PValue from " + _
                              "Param where PTyp='TC' and " + _
                              "PMCd= OmTc)='PL' and OdDmCd='" + ADC("wCpyDsgDmCd") + "' and " + _
                              "OdDmSz='" + ADC("wCpyDsgDmSz") + "' and OdSfx='" + ADC("wCpyDsgSfx") + "' " + _
                              "and OmCmCtg='C' and OmCmCd='" + ws_Src + "' and OmCmCurCd= '" + ADC("MqmCmCurCd") + "' and OmPrtKey='" + ctCurrPrtn + "' ")
        Else
          ADC("wCpyDsgSrcVch") = ""
        End If
      End If
    
      
    Case Is = UCase("wCpyVchOdNo")
      If ADC.Mode = xNorm Then
        ADC("wCpyVchOdNo") = 0
      End If
  
    Case Is = UCase("wCpyVchOdSrFr")
      If ADC.Mode = xNorm Then
        ADC("wCpyVchOdSrFr") = 0: ADC("wCpyVchOdSrTo") = 0
      End If
  
  End Select
End Sub

Private Sub CmdCpyRfId_Click()
  If Not ADC.SaveRec Then
    DispMsg "X Save Failed", etInfo
    Exit Sub
  End If
  'If ADC("FGFRBLOC") = "" Then DispMsg "Enter Bag From Location.", etError: Exit Sub

  Call DispFra(CpyRfId)

End Sub

Private Sub CmdCpyRfIdGo_Click()
'vk.4 RfId changed as Bag Alias No.
    If ADC("wRfIdSel") = "" Then DispMsg "Enter Bag Alias No. / Bag No List", etError: Exit Sub
    Dim wRfIdStr As String
    Dim wRfIdNoArr() As String
    Dim wTotRfid As Single, i As Single
    Dim wRfIdNo As String
    'vk.4 calling function to replace enter with ,
    wRfIdStr = PrepareRfIdString(ADC("wRfIdSel"))
    Dim wInvalRfIdStr As String
    If ADC("wRfIdSel") <> "" Then
        If ADC("WCPYTHRUBAGNO") = "N" Then
          wRfIdNoArr = Split(wRfIdStr, ",")
          wInvalRfIdStr = ""
          If Left(wRfIdStr, 1) = "," Or Right(wRfIdStr, 1) = "," Then
            'vk.4 Tag Input changed as Bag Alias No.
            DispMsg "Invalid Bag Alias No.", etError: Exit Sub
          End If
      
          wRfIdStr = ""
          wTotRfid = UBound(wRfIdNoArr) + 1
          For i = 0 To UBound(wRfIdNoArr)
            wRfIdNo = wRfIdNoArr(i)
            'vk.4 Tag changed as Bag Alias No.
            If Not IsNumeric(wRfIdNo) Then DispMsg "Bag Alias No. Not Numeric", etError: Exit Sub
            If Not moCn.RecSeek("Select 'x' from RfIdMst Where RimNo=" + wRfIdNo + " ") Then
                wInvalRfIdStr = wInvalRfIdStr + IIF(wInvalRfIdStr = "", "", ",") + wRfIdNo
            End If
            wRfIdStr = wRfIdStr + IIF(wRfIdStr = "", "", ",") + wRfIdNo
          Next
          'vk.4 Tag changed as Bag Alias No.
          If wInvalRfIdStr <> "" Then DispMsg "Bag Alias No. Not Defined:" + wInvalRfIdStr, etError: Exit Sub
      '*******Bhavna added bag no select
    Else
          wInvalRfIdStr = ""
          wRfIdStr = ParseBagNoBarCd(wRfIdStr)
          wRfIdNoArr = Split(wRfIdStr, ",")
        
          If Left(wRfIdStr, 1) = "," Or Right(wRfIdStr, 1) = "," Then
                DispMsg "Invalid Bag No Input", etError: Exit Sub
          End If
          wRfIdStr = ""
            
          For i = 0 To UBound(wRfIdNoArr)
              wRfIdNo = wRfIdNoArr(i)
              If Not moCn.RecSeek("Select 'x' from Bag Where Byy+'/'+BChr+'/'+ltrim(RTrim(str(BNo))) = '" + wRfIdNo + "'") Then
                wInvalRfIdStr = wInvalRfIdStr + IIF(wInvalRfIdStr = "", "", ",") + wRfIdNo
              End If
              wRfIdStr = wRfIdStr + IIF(wRfIdStr = "", "", ",") + wRfIdNo
              
          Next
          wTotRfid = UBound(wRfIdNoArr) + 1
          If wInvalRfIdStr <> "" Then DispMsg "Bag Not Defined:" + wInvalRfIdStr, etError: Exit Sub
        End If

      Dim wRs_OrdDsg As MDORowSet
      Dim wRs_DsgMst As MDORowSet
      If ADC("WCPYTHRUBAGNO") = "Y" Then
            Set wRs_OrdDsg = moCn.OpenRes("Select OdCoCd, OdTc, OdYy, OdChr, OdNo, OdSr, OdDmCd, OdSfx, OdDmSz, OdLabAsWt, OdDmCol  " + _
                               " from Bag Join OrdDsg  " + _
                                       " On BCoCd =OdCoCd and BOdTc=OdTc and BOdYy=OdYy and BOdChr=OdChr and BOdNo=OdNo and OdSr=BOdSr " + _
                                " where charindex(','+BYy + '/' + BChr + '/' + LTrim(RTrim(str(BNo)))+',' , '," + wRfIdStr + ",') <> 0 " + _
                                " Order By BYy, BChr, BNo")
      Else
           Set wRs_OrdDsg = moCn.OpenRes("Select OdCoCd, OdTc, OdYy, OdChr, OdNo, OdSr, OdDmCd, OdSfx, OdDmSz, OdLabAsWt, OdDmCol  " + _
                               " from Bag Join OrdDsg  " + _
                               " On BCoCd =OdCoCd and BOdTc=OdTc and BOdYy=OdYy and BOdChr=OdChr and BOdNo=OdNo and OdSr=BOdSr " + _
                               "Join RfIdMst On RimBCoCd=BCoCd and RimBYy=BYy and RimBChr=BChr and RimBNo=BNo " + _
                               "where RimNo in (" + wRfIdStr + ") Order By BYy, BChr, BNo")
      End If
      If wRs_OrdDsg.RecCount > 0 Then Call CpyDsgGrades(wRs_OrdDsg, False, False)
'vk.4 ADC("wRfIdSel") replaced with wRfIdStr only for ADC("WCPYTHRUBAGNO") = "Y"
      If ADC("WCPYTHRUBAGNO") = "Y" Then
            Set wRs_DsgMst = moCn.OpenRes("Select '' as OdCoCd,'' as OdTc,'' as OdYy,'' as OdChr,0 as OdNo,0 as OdSr,DmCd  as OdDmCd,'' as OdSfx,'' as OdDmSz,0 as OdLabAsWt,DmCol as  OdDmCol  " + _
                               " from DsgMst  " + _
                               "Join bag On  DmTcTyp='DM' and BOdDmCd=DmCd and DmSz='' " + _
                               "where charindex(','+BYy + '/' + BChr + '/' + LTrim(RTrim(str(BNo)))+',' , '," + wRfIdStr + ",') <> 0 " + _
                               "Order By DmCd ")
      Else
            Set wRs_DsgMst = moCn.OpenRes("Select '' as OdCoCd,'' as OdTc,'' as OdYy,'' as OdChr,0 as OdNo,0 as OdSr,DmCd  as OdDmCd,'' as OdSfx,'' as OdDmSz,0 as OdLabAsWt,DmCol as  OdDmCol  " + _
                               " from DsgMst  " + _
                               "Join RfIdMst On  DmTcTyp='DM' and RimDmCd=DmCd and DmSz='' and RimBNo=0 " + _
                               "where RimNo in (" + wRfIdStr + ") Order By DmCd ")
      End If
      If wRs_DsgMst.RecCount > 0 Then Call CpyDsgGrades(wRs_DsgMst, True, False)
  End If
  
  Call CmdCpyRfId_Click

End Sub
Private Sub CmdCpyVch_Click()
If Not ADC.SaveRec Then DispMsg "X Save Failed", etError: Exit Sub

If ADC("wCpyVchRt") = "" Then ADC("wCpyVchRt") = "N"

Call DispFra(CpyVch)
If GrdMpQtDsg.Rows = 1 Then GrdPropSet
End Sub

Private Sub CmdCpyDsg_Click()
If Not ADC.SaveRec Then DispMsg "X Save Failed", etError: Exit Sub
If ADC("wCpyDsgRt") = "" Then ADC("wCpyDsgRt") = "N"
If ADC("wCpyDsgBarCdYn") = "" Then
  ADC("wCpyDsgBarCdYn") = "N"
  Call ATXT_FldChange(99, UCase("wCpyDsgBarCdYn"), False, "", "N")
End If

Call DispFra(CpyDsg)
If GrdMpQtDsg.Rows = 1 Then GrdPropSet
End Sub

Private Sub CmdCpyDsgGo_Click()
  If ADC("wCpyDsgSrc") = "" Then DispMsg "Design Copy Source Should Be Entered", etError: Exit Sub
  
  If UCase(ADC("wCpyDsgSrc")) = UCase(ctDsgCpyMst) And ADC("wCpyDsgRt") = "Y" Then _
     DispMsg "Cannot Copy Rate As Copy Is From Design Master", etError: Exit Sub
  
  Dim wRs_Dsg As MDORowSet, wCpyFrDsgMst As Boolean
  Dim ws_Ord() As String
  
  If ADC("wCpyDsgSrc") = ctDsgCpyMst Then
    Set wRs_Dsg = moCn.OpenRes("Select '' as OdCoCd, '' as OdTc, '' as OdYy, '' as OdChr, 0 as OdNo, 0 as OdSr, DmCd as OdDmCd, '' as OdSfx, DmSz as OdDmSz, 0 as OdLabAsWt, DmCol as OdDmCol From DsgMst " + _
                              "Where DmTcTyp='DM' and DmCd='" + ADC("wCpyDsgDmCd") + "' and DmSz='" + ADC("wCpyDsgDmSz") + "'")
    If wRs_Dsg.RecCount <> 1 Then DispMsg "This Dsg/Sz Combination Does Not Exist In Design Mst", etError: Exit Sub
    wCpyFrDsgMst = True
  Else
    ws_Ord = Split(ADC("wCpyDsgSrcVch"), "/")
    Set wRs_Dsg = moCn.OpenRes("Select OdCoCd, OdTc, OdYy, OdChr, OdNo, OdSr, OdDmCd, OdSfx, OdDmSz, OdLabAsWt,OdDmCol from OrdDsg " + _
                              "Where OdCoCd='" + ctSelfCoCd + "' and OdTc='" + ws_Ord(0) + "' and OdYy='" + ws_Ord(1) + "' and OdChr='" + ws_Ord(2) + "' and OdNo=" + CStr(ws_Ord(3)) + " and OdSr=" + CStr(ws_Ord(4)))
    If wRs_Dsg.RecCount <> 1 Then DispMsg "This Cur/Dsg/Suf/Sz Combination Does Not Exist In Party-Design Combination Master", etError: Exit Sub
    wCpyFrDsgMst = False
  End If
  
  Call CpyDsgGrades(wRs_Dsg, wCpyFrDsgMst, IIF(ADC("wCpyDsgRt") = "Y", True, False))
  If ADC("wCpyDsgBarCdYn") = "N" Then Call CmdCpyDsg_Click

End Sub

Private Sub CmdCpyVchGo_Click()
  If Not moCn.RecSeek("Select 'x' from OrdMst where OmCoCd= '" + ADC("wCpyVchCoCd") + "' and " + _
            "OmTc= '" + ADC("wCpyVchOdTc") + "' and OmYy= '" + ADC("wCpyVchOdYy") + "' and OmChr= '" + ADC("wCpyVchOdChr") + "' and " + _
            "OmNo= " + CStr(ADC("wCpyVchOdNo")) + " and OmCmCurCd= '" + ADC("MqmCmCurCd") + "' And OmPrtKey='" + ctCurrPrtn + "' ") Then
    DispMsg "Invalid Voucher No. / Source And Target Currencies Are Not Matching", etError: Exit Sub
  End If
  
  Dim wRs_OrdDsg As MDORowSet
  Set wRs_OrdDsg = moCn.OpenRes("Select OdCoCd, OdTc, OdYy, OdChr, OdNo, OdSr, OdDmCd, OdSfx, OdDmSz, OdLabAsWt, OdDmCol from OrdDsg " + _
                            "Where OdCoCd='" + ADC("wCpyVchCoCd") + "' and OdTc='" + ADC("wCpyVchOdTc") + "' and OdYy='" + ADC("wCpyVchOdYy") + "' and OdChr='" + ADC("wCpyVchOdChr") + "' and OdNo=" + CStr(ADC("wCpyVchOdNo")) + _
                            IIF(ADC("wCpyVchOdSrFr") <> 0, " and OdSr>=" + CStr(ADC("wCpyVchOdSrFr")) + " ", "") + _
                            IIF(ADC("wCpyVchOdSrTo") <> 0, " and OdSr<=" + CStr(ADC("wCpyVchOdSrTo")) + " ", ""))
  
  
  Call CpyDsgGrades(wRs_OrdDsg, False, IIF(ADC("wCpyVchRt") = "Y", True, False))
  Call CmdCpyVch_Click
  End Sub

Private Sub ADC_SetRecSource()
  '*** Set the RecSource for ADC as xdatUnbound
  ADC.RecSource = " Select * from MultiPrcQtMst " + _
                  " Where MqmCoCd='" + ADC("MqmCoCd") + "' and MqmTc='" + ADC("MqmTc") + "' and MqmYy='" + ADC("MqmYy") + "' " + _
                  " and MqmChr='" + ADC("MqmChr") + "' and MqmNo =" + CStr(ADC("MqmNo"))
                  
End Sub


Private Sub CmdDsgDet_Click()
If GrdMpQtDsg.Rows = 1 Then Exit Sub

If FraDsgDet.Visible = False Then Call ShowDsgDet

Call DispFra(DsgDet)

'If FraDsgDet.Visible = True Then
'  Call SetGrdMpQtLab("S")
'  Call SetGrdMpQtRm("S")
'End If
End Sub
Private Sub ShowDsgDet()
If GrdMpQtDsg.Rows = 1 Then Exit Sub

Dim wDiaGrd As String, wDiaGrdNo As Integer, wColNm As String
Dim wRmStr As String, wLabStr As String
Dim wRs_Rm As MDORowSet, wRs_Lab As MDORowSet
Dim wRmRow As Single, wLabRow As Single

wRmRow = 0: wLabRow = 0

''If GrdMpQtDsg.Col >= 7 And GrdMpQtDsg.Col <= 14 Then
''  wDiaGrdNo = GrdMpQtDsg.Col - 6
''End If

wColNm = UCase(GrdMpQtDsg.ColProp(GrdMpQtDsg.Col).ColName)
If Left(wColNm, 10) = UCase("wMqdDiaGrd") And Right(wColNm, 3) = UCase("Prc") Then
    wDiaGrdNo = Mid(wColNm, 11, Len(wColNm) - 13)
End If
wDiaGrdNo = IIF(wDiaGrdNo = 0, 1, wDiaGrdNo)
wDiaGrd = ADC("MqmDiaGrd" + Trim(Str(wDiaGrdNo)))

LblDsgDet.Caption = "  Design Details (For Diamond Grade " + wDiaGrd + ")"

ADC("wMqdGldAs") = GrdMpQtDsg.Value(GrdMpQtDsg.Row, "MqdGldAs")
ADC("wMqdGldAsWt") = GrdMpQtDsg.Value(GrdMpQtDsg.Row, "MqdGldAsWt")
ADC("wMqdLabAs") = GrdMpQtDsg.Value(GrdMpQtDsg.Row, "MqdLabAs")
ADC("wMqdLabAsWt") = GrdMpQtDsg.Value(GrdMpQtDsg.Row, "MqdLabAsWt")

'3.11.2
wRmStr = " Open Symmetric Key DataEnde Decryption By Password = 'RMeWm'; " + _
         " Select MqrCoCd, MqrTc, MqrYy, MqrChr, MqrNo, MqrSr, MqrSrNo, MqrDiagrd, MqrIdNo, MqrPrtKey, MqrRmCtg, MqrRmSCtg, " + _
         "        Convert(varChar(100), DecryptByKey(Convert(varBinary(100), MqrRmCdBin))) as MqrRmCd, " + _
         "        MqrLn1, MqrLn2, MqrQty, MqrRmPtr, MqrWt, MqrLmeSal, " + _
         "        Cast(Convert(varChar(100), DecryptByKey(Convert(varBinary(100), MqrSalRtBin))) as Decimal(16,2)) as MqrSalRt, " + _
         "        Cast(Convert(varChar(100), DecryptByKey(Convert(varBinary(100), MqrSalValBin))) as Decimal(16,2)) as MqrSalVal, " + _
         "        MqrSetScd, Cast(Convert(varChar(100), DecryptByKey(Convert(varBinary(100), MqrSetSalRtBin))) as Decimal(16,2)) as MqrSetSalRt, " + _
         "        Cast(Convert(varChar(100), DecryptByKey(Convert(varBinary(100), MqrSetSalValBin))) as Decimal(16,2)) as MqrSetSalVal, MqrAlyCd, MqrAlySalRt, MqrMainMet, ModUsr, ModDt, ModTime, MqrRmCdBin, MqrSalRtBin, MqrSalValBin, MqrSetSalRtBin, MqrSetSalValBin " + _
         " from MultiPrcQtRm " + _
         " Where MqrCoCd='" + ADC("MqmCoCd") + "' and MqrTc='" + ADC("MqmTc") + "' and MqrYy='" + ADC("MqmYy") + "' and MqrChr='" + ADC("MqmChr") + "' and MqrNo=" + CStr(ADC("MqmNo")) + " and MqrSr=" + CStr(GrdMpQtDsg.Value(GrdMpQtDsg.Row, "MqdSr")) + " and MqrDiaGrd='" + wDiaGrd + "';Close Symmetric Key DataEnde; "
      
Set wRs_Rm = moCn.OpenRes(wRmStr)

If wRs_Rm.RecCount > 0 Then
  With wRs_Rm
    .MoveFirst
    GrdMpQtRm.Rows = 1
    Do While Not (.EOF Or .BOF)
      wRmRow = wRmRow + 1
      GrdMpQtRm.AddItem
      GrdMpQtRm.Value(wRmRow, "wMqrSrNo") = !MqrSrNo
      GrdMpQtRm.Value(wRmRow, "wMqrRmCtg") = !MqrRmCtg
      GrdMpQtRm.Value(wRmRow, "wMqrRmSCtg") = !MqrRmSCtg
      GrdMpQtRm.Value(wRmRow, "wMqrRmCd") = !MqrRmCd
      GrdMpQtRm.Value(wRmRow, "wMqrLn1") = !MqrLn1
      GrdMpQtRm.Value(wRmRow, "wMqrLn2") = !MqrLn2
      GrdMpQtRm.Value(wRmRow, "wMqrQty") = !MqrQty
      GrdMpQtRm.Value(wRmRow, "wMqrRmPtr") = !MqrRmPtr
      GrdMpQtRm.Value(wRmRow, "wMqrWt") = !MqrWt
      GrdMpQtRm.Value(wRmRow, "wMqrLmeSal") = !MqrLmeSal
      GrdMpQtRm.Value(wRmRow, "wMqrSalRt") = !MqrSalRt
      GrdMpQtRm.Value(wRmRow, "wMqrSalVal") = !MqrSalVal
      GrdMpQtRm.Value(wRmRow, "wMqrSetScd") = !MqrSetScd
      GrdMpQtRm.Value(wRmRow, "wMqrSetSalRt") = !MqrSetSalRt
      GrdMpQtRm.Value(wRmRow, "wMqrSetSalVal") = !MqrSetSalVal
      GrdMpQtRm.Value(wRmRow, "wMqrAlyCd") = !MqrAlyCd
      GrdMpQtRm.Value(wRmRow, "wMqrAlySalRt") = !MqrAlySalRt
      GrdMpQtRm.Value(wRmRow, "wMqrMainMet") = !MqrMainMet
      .MoveNext
    Loop
  End With
End If
'3.11.2
wLabStr = " Open Symmetric Key DataEnde Decryption By Password = 'RMeWm'; " + _
          " Select MqlCoCd, MqlTc, MqlYy, MqlChr, MqlNo, MqlSr, MqlSrNo, MqlIdNo, MqlPrtKey, " + _
          "     Convert(varChar(100), DecryptByKey(Convert(varBinary(100), MqlMcdBin))) as MqlMcd, " + _
          "     Convert(varChar(100), DecryptByKey(Convert(varBinary(100), MqlScdBin))) as MqlScd, " + _
          "     MqlQty, Cast(Convert(varChar(100), DecryptByKey(Convert(varBinary(100), MqlSalRtBin))) as Decimal(16,2)) as MqlSalRt, " + _
          "     Cast(Convert(varChar(100), DecryptByKey(Convert(varBinary(100), MqlSalValBin))) as Decimal(16,2)) as MqlSalVal, MqlQw, ModUsr, ModDt, ModTime, MqlMcdBin, MqlScdBin, MqlSalRtBin, MqlSalValBin " + _
          "  from MultiPrcQtLab " + _
          " Where MqlCoCd='" + ADC("MqmCoCd") + "' and MqlTc='" + ADC("MqmTc") + "' and MqlYy='" + ADC("MqmYy") + "' and MqlChr='" + ADC("MqmChr") + "' and MqlNo=" + CStr(ADC("MqmNo")) + " and MqlSr=" + CStr(GrdMpQtDsg.Value(GrdMpQtDsg.Row, "MqdSr")) + " ;Close Symmetric Key DataEnde; "
Set wRs_Lab = moCn.OpenRes(wLabStr)

If wRs_Lab.RecCount > 0 Then
  With wRs_Lab
    .MoveFirst
    GrdMpQtLab.Rows = 1
    Do While Not (.EOF Or .BOF)
      wLabRow = wLabRow + 1
      GrdMpQtLab.AddItem
      GrdMpQtLab.Value(wLabRow, "wMqlSrNo") = !MqlSrNo
      GrdMpQtLab.Value(wLabRow, "wMqlMCd") = !MqlMcd
      GrdMpQtLab.Value(wLabRow, "wMqlSCd") = !MqlScd
      GrdMpQtLab.Value(wLabRow, "wMqlQw") = !MqlQw
      GrdMpQtLab.Value(wLabRow, "wMqlQty") = !MqlQty
      GrdMpQtLab.Value(wLabRow, "wMqlSalRt") = !MqlSalRt
      GrdMpQtLab.Value(wLabRow, "wMqlSalVal") = !MqlSalVal
      .MoveNext
    Loop
  End With
End If


Set wRs_Rm = Nothing
Set wRs_Lab = Nothing
End Sub


Private Sub CmdGenExcel_Click()
''CmdGenExcel.CausesValidation = True
GenExcel
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
' ***** Manali 3.10.0 - 18/04/12 - Bar Code
Dim wCancel As Boolean
If KeyAscii <> vbKeyReturn Then Exit Sub
If UCase(Me.ActiveControl.Name) = UCase("txtCpyDsgDmCd") Then
  Call txtCpyDsgDmCd_Validate(wCancel)
  If Not wCancel And txtCpyDsgDmCd.text <> "" Then
    CmdCpyDsgGo.SetFocus
    Call CmdCpyDsgGo_Click
    txtCpyDsgDmCd.SetFocus
    txtCpyDsgDmCd.text = ""
    wCancel = False
    Call txtCpyDsgDmCd_Validate(wCancel)
  End If
End If
' ***** Manali 3.10.0 - 18/04/12 - Bar Code
End Sub

Private Sub GrdMpQtDsg_ColWhen(ByVal RowNum As Integer, ByVal ColName As String, Cancel As Boolean, ErrMsg As String, ByVal Action As MwfCtl.en_fgAction)
  
  Cancel = True: ErrMsg = "Cannot Edit": Exit Sub
End Sub

Private Sub GrdMpQtDsg_InitKey(ByVal RowNum As Integer)
 With GrdMpQtDsg
    If .IsNew(RowNum) Then
      .Store "MqdCoCd", ADC("MqmCoCd")
      .Store "MqdTc", ADC("MqmTc")
      .Store "MqdYy", ADC("MqmYy")
      .Store "MqdChr", ADC("MqmChr")
      .Store "MqdNo", ADC("MqmNo")
      
      .Store "MqdMqmIdNo", moCn.GetFldVal("Select MqmIdNo From MultiPrcQtMst " + _
                                        "Where MqmCoCd='" + ADC("MqmCoCd") + "' And MqmTc='" + ADC("MqmTc") + "' " + _
                                        "  And MqmYy='" + ADC("MqmYy") + "' And MqmChr='" + ADC("MqmChr") + "' And MqmNo=" + CStr(ADC("MqmNo")) + " And MqmPrtKey='" + ctCurrPrtn + "'")
      
      .Store "MqdDmIdNo", moCn.GetFldVal("Select DmIdNo From DsgMst Where DmTcTyp='DM' And DmCd='" + .Value(RowNum, "MqDDMCD") + "' " + _
                                    " And DmSz='' And DmPrtKey='" + ctCurrPrtn + "'")
      
      '3.11.2
      .Store "MqdDiaGrd1Prc", .Value(RowNum, "wMqdDiaGrd1Prc")
      .Store "MqdDiaGrd2Prc", .Value(RowNum, "wMqdDiaGrd2Prc")
      .Store "MqdDiaGrd3Prc", .Value(RowNum, "wMqdDiaGrd3Prc")
      .Store "MqdDiaGrd4Prc", .Value(RowNum, "wMqdDiaGrd4Prc")
      .Store "MqdDiaGrd5Prc", .Value(RowNum, "wMqdDiaGrd5Prc")
      .Store "MqdDiaGrd6Prc", .Value(RowNum, "wMqdDiaGrd6Prc")
      .Store "MqdDiaGrd7Prc", .Value(RowNum, "wMqdDiaGrd7Prc")
      .Store "MqdDiaGrd8Prc", .Value(RowNum, "wMqdDiaGrd8Prc")
      
    End If
  End With
End Sub
Private Sub GrdMpQtDsg_ReCalc(ByVal RowNum As Integer, ByVal ColName As String, ByVal OldVal As Variant, ByVal SourceName As String, ByVal SourceColName As String)
'3.11.2
Dim wRsDsgEnde As MDORowSet
With GrdMpQtDsg
If .Rows > 1 Then
   Set wRsDsgEnde = moCn.OpenRes(" Open Symmetric Key DataEnde Decryption By Password = 'RMeWm'; " + _
                               " Select Cast(Convert(varChar(100), DecryptByKey(Convert(varBinary(100), MqdDiaGrd1PrcBin))) as Decimal(16,2)) as MqdDiaGrd1Prc, " + _
                               "        Cast(Convert(varChar(100), DecryptByKey(Convert(varBinary(100), MqdDiaGrd2PrcBin))) as Decimal(16,2)) as MqdDiaGrd2Prc, " + _
                               "        Cast(Convert(varChar(100), DecryptByKey(Convert(varBinary(100), MqdDiaGrd3PrcBin))) as Decimal(16,2)) as MqdDiaGrd3Prc, " + _
                               "        Cast(Convert(varChar(100), DecryptByKey(Convert(varBinary(100), MqdDiaGrd4PrcBin))) as Decimal(16,2)) as MqdDiaGrd4Prc, " + _
                               "        Cast(Convert(varChar(100), DecryptByKey(Convert(varBinary(100), MqdDiaGrd5PrcBin))) as Decimal(16,2)) as MqdDiaGrd5Prc, " + _
                               "        Cast(Convert(varChar(100), DecryptByKey(Convert(varBinary(100), MqdDiaGrd6PrcBin))) as Decimal(16,2)) as MqdDiaGrd6Prc, " + _
                               "        Cast(Convert(varChar(100), DecryptByKey(Convert(varBinary(100), MqdDiaGrd7PrcBin))) as Decimal(16,2)) as MqdDiaGrd7Prc, " + _
                               "        Cast(Convert(varChar(100), DecryptByKey(Convert(varBinary(100), MqdDiaGrd8PrcBin))) as Decimal(16,2)) as MqdDiaGrd8Prc " + _
                               " From MultiPrcQtDsg " + _
                               " Where MqdCoCd='" + ADC("MqmCoCd") + "' and MqdTc='" + ADC("MqmTc") + "' and MqdYy='" + ADC("MqmYy") + "' " + _
                               "       and MqdChr='" + ADC("MqmChr") + "' and MqdNo =" + CStr(ADC("MqmNo")) + " and MqdSr =" + CStr(.Value(RowNum, "MqdSr")) + ";Close Symmetric Key DataEnde;")
    If wRsDsgEnde.RecCount > 0 Then
        Select Case UCase(ColName)
        Case Is = UCase("wMqdDiaGrd1Prc")
            If .Mode = fgminit Then
                 .Value(RowNum, "wMqdDiaGrd1Prc") = wRsDsgEnde!MqdDiaGrd1Prc
                 .Value(RowNum, "wMqdDiaGrd2Prc") = wRsDsgEnde!MqdDiaGrd2Prc
                 .Value(RowNum, "wMqdDiaGrd3Prc") = wRsDsgEnde!MqdDiaGrd3Prc
                 .Value(RowNum, "wMqdDiaGrd4Prc") = wRsDsgEnde!MqdDiaGrd4Prc
                 .Value(RowNum, "wMqdDiaGrd5Prc") = wRsDsgEnde!MqdDiaGrd5Prc
                 .Value(RowNum, "wMqdDiaGrd6Prc") = wRsDsgEnde!MqdDiaGrd6Prc
                 .Value(RowNum, "wMqdDiaGrd7Prc") = wRsDsgEnde!MqdDiaGrd7Prc
                 .Value(RowNum, "wMqdDiaGrd8Prc") = wRsDsgEnde!MqdDiaGrd8Prc
            End If
      End Select
      End If
    End If
 End With
 Set wRsDsgEnde = Nothing
End Sub

Private Sub GrdMpQtDsg_RowWhen(ByVal RowNum As Integer)
  With GrdMpQtDsg
  FraPic3D.Left = wFraPic3DLft: FraPic3D.Top = wFraPic3DTop
  FraPic3D.Height = wFraPic3DHt: FraPic3D.Width = wFraPic3DWdt
      
  w3DPth = GetPictPath("DM", False, .Value(.Row, "MqdDmCd"), .Value(.Row, "MqdDmSz"))
  
  Call ShowPic(wPic3DHt, wPic3DWdt, w3DPth, Pic3D)
  End With
End Sub

Private Sub GrdMpQtDsg_SetRecSource()
  '*** Set the Record Source of the Grid GrdMpQtDsg
  GrdMpQtDsg.RecSource = " Select * from MultiPrcQtDsg " + _
                          " Where MqdCoCd='" + ADC("MqmCoCd") + "' and MqdTc='" + ADC("MqmTc") + "' and MqdYy='" + ADC("MqmYy") + "' " + _
                          " and MqdChr='" + ADC("MqmChr") + "' and MqdNo =" + CStr(ADC("MqmNo")) + _
                          " Order By MqdCoCd, MqdTc, MqdYy, MqdChr, MqdNo, MqdSr "
End Sub

Private Sub CpyDsgGrades(ByVal pv_RsSrcDsg As MDORowSet, ByVal pv_CpyFrDsgMst As Boolean, ByVal pv_CpyRt As Boolean)
'On Error GoTo ErrH
  
  Dim wRs_Dsg As MDORowSet
  Dim wRs_RmDet As MDORowSet, wRs_LabDet As MDORowSet
  Dim wRmStr As String, wLabStr As String
  Dim wRmRt As Single, wSetRt As Single, wLabRt As Single, wAlyRt As Single
  Dim wRmVal As Double, wTotRmVal As Double
  Dim wLabVal As Double, wTotSetVal As Double, wTotLabVal As Double, wDsgVal As Double, wSetVal As Double
  Dim wDRmVal As Double, wCRmVal As Double
  Dim wbRmRt As Boolean
  Dim wGrsWt As Single, wDiaWt As Single
  Dim wLmeSal As Single
  
  Dim wBaseRmCtg As String, wBaseLme As Single
  
  Dim wKt As String, wDiaGrd As String
  
  Dim ws_Ord() As String, wCnd As String
  Dim wCoCd As String, wOdTc As String, wOdYy As String, wOdChr As String, wOdNo As Long, wOdSr As Long
  
  Dim ws_InsertRm As String, ws_InsertLab As String

  
  Dim wDsgSrNotCopied As String
  
  If ADC("MqmDiaGrd1") = "" And ADC("MqmDiaGrd2") = "" And ADC("MqmDiaGrd3") = "" And ADC("MqmDiaGrd4") = "" And _
    ADC("MqmDiaGrd5") = "" And ADC("MqmDiaGrd6") = "" And ADC("MqmDiaGrd7") = "" And ADC("MqmDiaGrd8") = "" Then _
    DispMsg "Enter Diamond Grades", etError: Exit Sub
  
  If ADC("MqmMetGrd1") = "" And ADC("MqmMetGrd2") = "" And ADC("MqmMetGrd3") = "" And ADC("MqmMetGrd4") = "" Then _
    DispMsg "Enter Metal Grades", etError: Exit Sub
  
  If pv_RsSrcDsg.RecCount = 0 Then DispMsg "No Designs Found For Given Scope", etError: Exit Sub
  
  Dim wRsCustMst As MDORowSet
  Set wRsCustMst = moCn.OpenRes("Select * from CustMst Where CmCtg='C' and CmCd='" + ADC("MqmCmCd") + "'")
  wRsCustMst.MoveFirst
  
  Dim wGldAsStr As String, wLabAsStr As String
  Dim wGldAsWt As Double, wLabAsWt As Double
  wGldAsStr = moCn.GetFldVal("Select PDesc225 from Param Where PTyp='GLDAS' and PMCd='" + wRsCustMst!CmGldAs + "' and PSCd=''")
  wLabAsStr = moCn.GetFldVal("Select PDesc225 from Param Where PTyp='GLDAS' and PMCd='" + wRsCustMst!CmLabAs + "' and PSCd=''")
  
  Dim i As Integer, j As Integer, wi_Row As Integer, wi_Sr As Integer
  
  
  GrdMpQtDsg.StartCopy
  MWLib.BeginProcess Me, "Copying Designs ..."
  
  pv_RsSrcDsg.MoveFirst
  Do While Not (pv_RsSrcDsg.EOF Or pv_RsSrcDsg.BOF)
    MWLib.RunningProcess "Copying Design " + pv_RsSrcDsg!OdDmCd
    For i = 1 To 4
      wKt = ADC("MqmMetGrd" + Trim(Str(i)))
      If wKt <> "" Then
'        If moCn.RecSeek("Select 'x' from MultiPrcQtDsg " + _
'                        "Where MqdCoCd='" + ADC("MqmCoCd") + "' and MqdTc='" + ADC("MqmTc") + "' and MqdYy='" + ADC("MqmYy") + "' " + _
'                        "and MqdChr='" + ADC("MqmChr") + "' and MqdNo =" + CStr(ADC("MqmNo")) + _
'                        "and MqdDmCd='" + pv_RsSrcDsg!OdDmCd + "' and MqDDmSz='" + pv_RsSrcDsg!OdDmSz + "' and MqdMetGrd='" + wKt + "' ") Then
'          If pv_CpyFrDsgMst Then
'            wDsgSrNotCopied = wDsgSrNotCopied + IIF(wDsgSrNotCopied = "", "", ", ") + pv_RsSrcDsg!OdDmCd + "-" + pv_RsSrcDsg!OdDmSz + "-" + wKt
'          Else
'            wDsgSrNotCopied = wDsgSrNotCopied + IIF(wDsgSrNotCopied = "", "", ", ") + CStr(pv_RsSrcDsg!OdSr) + " " + pv_RsSrcDsg!OdDmCd + "-" + pv_RsSrcDsg!OdDmSz + "-" + wKt
'          End If
'          GoTo NextMetGrd
'        End If
        
        MWLib.RunningProcess "Copying Design " + pv_RsSrcDsg!OdDmCd + " For MetGrd  " + wKt
        GrdMpQtDsg.AddItem
        
        wi_Row = GrdMpQtDsg.Rows - 1
        wi_Sr = GrdMpQtDsg.MaxVal("MqdSr") + 1
        GrdMpQtDsg.Value(wi_Row, "MqdSr") = wi_Sr
        GrdMpQtDsg.Value(wi_Row, "MqdDmCd") = pv_RsSrcDsg!OdDmCd
        GrdMpQtDsg.Value(wi_Row, "MqdDmSz") = pv_RsSrcDsg!OdDmSz
        GrdMpQtDsg.Value(wi_Row, "MqdMETGRD") = wKt
        GrdMpQtDsg.Value(wi_Row, "MqdGldAs") = wRsCustMst!CmGldAs
        GrdMpQtDsg.Value(wi_Row, "MqdLabAs") = wRsCustMst!CmLabAs
        
        GrdMpQtDsg.Row = wi_Row
        GrdMpQtDsg.Col = 2
        
        ws_InsertRm = "": ws_InsertLab = ""
        
        For j = 1 To 8
          wDiaGrd = ADC("MqmDiaGrd" + Trim(Str(j)))
          wTotRmVal = 0: wTotLabVal = 0: wDsgVal = 0: wTotSetVal = 0
          wGldAsWt = 0: wLabAsWt = 0
          wSetRt = 0: wRmRt = 0: wLabRt = 0
          wDRmVal = 0: wCRmVal = 0
          wSetRt = 0: wRmRt = 0: wSetVal = 0: wLabVal = 0
          wbRmRt = True
          wGrsWt = 0: wDiaWt = 0
          If wDiaGrd = "" Then GoTo NextDiaGrd
          
          If pv_CpyFrDsgMst Then
            wRmStr = " Select DrSr as OrSrNo, DrRmCtg as OrRmCtg, DrRmCd as OrRmCd, DrLn1 as OrLn1, DrLn2 as OrLn2, DrRmPtr as OrRmPtr, DrQty as OrQty, DrWt as OrWt, " + _
                      " IsNull(D.GrCd, '') as qDiaGrd, IsNull(M.GrCd, '') as qMetGrd, " + _
                      " IsNull(D.GrRmCd, '') qNewDiaCd, IsNull(M.GrRmCd, '') qNewMetCd, " + _
                      " (Case When O.RmCtg In ('G', 'P', 'S', 'L', 'X') Then IsNull(nM.RmCtg, O.RmCtg) When O.RmCtg='D' Then IsNull(nD.RmCtg, O.RmCtg) Else O.RmCtg End) as qNewRmCtg, " + _
                      " (Case When O.RmCtg In ('G', 'P', 'S', 'L', 'X') Then IsNull(nM.RmSCtg, O.RmSCtg) When O.RmCtg='D' Then IsNull(nD.RmSCtg, O.RmSCtg) Else O.RmSCtg End) as qNewRmSCtg, " + _
                      " (Case When O.RmCtg In ('G', 'P', 'S', 'L', 'X') Then IsNull(M.GrRmCd, DrRmCd)  When O.RmCtg='D' Then IsNull(D.GrRmCd, DrRmCd) Else DrRmCd End) as qNewRmCd, " + _
                      " (Case When O.RmCtg In ('G', 'P', 'S', 'L', 'X') Then IsNull(nM.RmQW, O.RmQW)  When O.RmCtg='D' Then IsNull(nD.RmQW, O.RmQW) Else O.RmQW End) as qNewRmQW, " + _
                      " (Case When O.RmCtg In ('G', 'P', 'S', 'L', 'X') and IsNull(M.GrRmCd, '')<>'' Then " + _
                      "         (Case When O.RmCtg='X' Then (Case When nM.RmFndWt>0 Then IsNull(nM.RmFndWt, 0)*DrQty Else DrWt End) " + _
                      "               When O.RmCtg In ('G', 'P', 'S', 'L') Then Cast(IsNull(Nm.RmSpGrav, 0)*DrWt/O.RmSpGrav as Decimal(14,3)) Else DrWt End) Else DrWt End) as qNewRmWt, " + _
                      " 0 as OrSalRt, 0 as OrSalVal, DrSetSCd as OrSetSCd, 0 as OrSetSalRt, 0 as OrSetSalVal, 0 As OrLmeSal, " + _
                      " IsNull(nM.RmBaseCd, O.RmBaseCd) as qNewRmBaseCd, DrMainMet as OrMainMet, 0 AS OrAlySalRt, " + _
                      " IsNull((Select CaAlyCd from CustAly where CaCtg='C' and CaCd='" + ADC("MqmCmCd") + "' and CaRmCd= IsNull(M.GrRmCd, '') and CaDmCol='" + pv_RsSrcDsg!OdDmCol + "' and CaAlyCd<> '')" + _
                      " , IsNull((Select CaAlyCd from CustAly where CaCtg='C' and CaCd='" + ctSelfCmCd + "' and CaRmCd= IsNull(M.GrRmCd, '') and CaDmCol='" + pv_RsSrcDsg!OdDmCol + "' and CaAlyCd<> '' ),'')) as OrAlyCd, " + _
                      " 'N' as qCopyRt " + _
                      " From DsgRm " + _
                      " Join RmMst O on O.RmCd=DrRmCd and O.RmPrtKey=DrPrtKey " + _
                      " Left Outer Join Grade D On D.GrTyp='DIA' and D.GrRmCtg=DrRmCtg and D.GrFrLn <= DrLn1 and D.GrToLn >= DrLn1 and D.GrCmCd='" + ADC("MqmCmCd") + "' and D.GrRmId=O.RmId and D.GrPrtKey=DrPrtKey and D.GrCd = '" + wDiaGrd + "' " + _
                      " Left Outer Join RmMst nD on nD.RmCd=D.GrRmCd and nD.RmPrtKey=DrPrtKey " + _
                      " Left Outer Join Grade M On M.GrTyp='MET' and M.GrRmCtg=DrRmCtg and M.GrCmCd='" + ADC("MqmCmCd") + "' and M.GrRmId=O.RmId and M.GrPrtKey=DrPrtKey and M.GrCd = '" + wKt + "' " + _
                      " Left Outer Join RmMst nM on nM.RmCd=M.GrRmCd and nM.RmPrtKey=DrPrtKey " + _
                      " Where DrTcTyp='DM' and DrCd='" + pv_RsSrcDsg!OdDmCd + "' and DrSz=''  and DrPrtKey='" + ctCurrPrtn + "' " + _
                      " Order By OrSrNo "
          Else
            wRmStr = " Select OrSrNo, OrRmCtg, OrRmCd, OrLn1, OrLn2, OrRmPtr, OrQty, OrWt, " + _
                      " IsNull(D.GrCd, '') as qDiaGrd, IsNull(M.GrCd, '') as qMetGrd, " + _
                      " IsNull(D.GrRmCd, '') qNewDiaCd, IsNull(M.GrRmCd, '') qNewMetCd, " + _
                      " (Case When O.RmCtg In ('G', 'P', 'S', 'L', 'X') Then IsNull(nM.RmCtg, O.RmCtg) When O.RmCtg='D' Then IsNull(nD.RmCtg, O.RmCtg) Else O.RmCtg End) as qNewRmCtg, " + _
                      " (Case When O.RmCtg In ('G', 'P', 'S', 'L', 'X') Then IsNull(nM.RmSCtg, O.RmSCtg) When O.RmCtg='D' Then IsNull(nD.RmSCtg, O.RmSCtg) Else O.RmSCtg End) as qNewRmSCtg, " + _
                      " (Case When O.RmCtg In ('G', 'P', 'S', 'L', 'X') Then IsNull(M.GrRmCd, OrRmCd)  When O.RmCtg='D' Then IsNull(D.GrRmCd, OrRmCd) Else OrRmCd End) as qNewRmCd, " + _
                      " (Case When O.RmCtg In ('G', 'P', 'S', 'L', 'X') Then IsNull(nM.RmQW, O.RmQW)  When O.RmCtg='D' Then IsNull(nD.RmQW, O.RmQW) Else O.RmQW End) as qNewRmQW, " + _
                      " (Case When O.RmCtg In ('G', 'P', 'S', 'L', 'X') and IsNull(M.GrRmCd, '')<>'' Then " + _
                      "  (Case When O.RmCtg='X' Then (Case When nM.RmFndWt>0 Then IsNull(nM.RmFndWt, 0)*OrQty Else OrWt End) " + _
                      "        When O.RmCtg In ('G', 'P', 'S', 'L') Then Cast(IsNull(Nm.RmSpGrav, 0)*OrWt/O.RmSpGrav as Decimal(14,3)) Else OrWt End) Else OrWt End) as qNewRmWt, " + _
                      " OrSalRt, OrSalVal, OrSetSCd, OrSetSalRt, OrSetSalVal, OrLmeSal, " + _
                      " IsNull(nM.RmBaseCd, O.RmBaseCd) as qNewRmBaseCd, OrMainMet, OrAlySalRt, " + _
                      " IsNull((Select CaAlyCd from CustAly where CaCtg='C' and CaCd='" + ADC("MqmCmCd") + "' and CaRmCd= IsNull(M.GrRmCd, '') and CaDmCol='" + pv_RsSrcDsg!OdDmCol + "' and CaAlyCd<> '')" + _
                      " , IsNull((Select CaAlyCd from CustAly where CaCtg='C' and CaCd='" + ctSelfCmCd + "' and CaRmCd= IsNull(M.GrRmCd, '') and CaDmCol='" + pv_RsSrcDsg!OdDmCol + "' and CaAlyCd<> '' ),'')) as OrAlyCd, " + _
                      IIF(Not pv_CpyRt, "'N'", _
                      " (Case When O.RmCtg In ('G', 'P', 'S', 'L', 'X') Then (Case When IsNull(M.GrRmCd, OrRmCd) =OrRmCd Then 'Y' Else 'N' End) When O.RmCtg='D' Then (Case When IsNull(D.GrRmCd, OrRmCd) =OrRmCd Then 'Y' Else 'N' End) Else 'Y' End)") + " as qCopyRt " + _
                      " From OrdRm " + _
                      " Join RmMst O on O.RmCd=OrRmCd and O.RmPrtKey=OrPrtKey " + _
                      " Left Outer Join Grade D On D.GrTyp='DIA' and D.GrRmCtg=OrRmCtg and  D.GrFrLn <= OrLn1 and D.GrToLn >= OrLn1 and D.GrCmCd='" + ADC("MqmCmCd") + "' and D.GrRmId=O.RmId and D.GrPrtKey=OrPrtKey and D.GrCd = '" + wDiaGrd + "' " + _
                      " Left Outer Join RmMst nD on nD.RmCd=D.GrRmCd and nD.RmPrtKey=OrPrtKey " + _
                      " Left Outer Join Grade M On M.GrTyp='MET' and M.GrRmCtg=OrRmCtg and M.GrCmCd='" + ADC("MqmCmCd") + "' and M.GrRmId=O.RmId and M.GrPrtKey=OrPrtKey and M.GrCd = '" + wKt + "' " + _
                      " Left Outer Join RmMst nM on nM.RmCd=M.GrRmCd and nM.RmPrtKey=OrPrtKey " + _
                      " Where OrCoCd='" + pv_RsSrcDsg!OdCoCd + "' and OrTc='" + pv_RsSrcDsg!OdTc + "' and OrYy='" + pv_RsSrcDsg!OdYy + "' and OrChr='" + pv_RsSrcDsg!OdChr + "' and OrNo=" + CStr(pv_RsSrcDsg!OdNo) + _
                      " and OrSr=" + CStr(pv_RsSrcDsg!OdSr) + " and OrPrtKey='" + ctCurrPrtn + "' Order By OrSrNo "
          End If
          
          Set wRs_RmDet = moCn.OpenRes(wRmStr)
          If wRs_RmDet.RecCount = 0 Then GoTo NextDiaGrd
          
          With wRs_RmDet
          '****** 1st Loop of Rm to Calculate GldAsWt, LabAsWt
          .MoveFirst
          Do While Not (.EOF Or .BOF)
              If !OrMainMet = "Y" Then
                wGldAsWt = wGldAsWt + !qNewRmWt
              Else
                If InStr(1, "," + wGldAsStr + ",", "," + !qNewRmCtg + ",") <> 0 And !qNewRmCtg <> "G" And !qNewRmCtg <> "P" And !qNewRmCtg <> "S" And !qNewRmCtg <> "L" Then _
                  wGldAsWt = wGldAsWt + MWLib.Div(!qNewRmWt, IIF(!qNewRmCtg = "D" Or !qNewRmCtg = "C", 5, 1))
              End If
              If InStr(1, "," + wLabAsStr + ",", "," + !qNewRmCtg + ",") <> 0 Then wLabAsWt = wLabAsWt + MWLib.Div(!qNewRmWt, IIF(!qNewRmCtg = "D" Or !qNewRmCtg = "C", 5, 1))
              wGrsWt = wGrsWt + MWLib.Div(!qNewRmWt, IIF(!qNewRmCtg = "D" Or !qNewRmCtg = "C", 5, 1))
              wDiaWt = wDiaWt + IIF(!qNewRmCtg = "D", !qNewRmWt, 0)
            .MoveNext
          Loop
          
          If GrdMpQtDsg.Value(wi_Row, "MqdGrsWt") = 0 Then
            GrdMpQtDsg.Value(wi_Row, "MqdGrsWt") = wGrsWt
            GrdMpQtDsg.Value(wi_Row, "MqdDiaWt") = wDiaWt
            GrdMpQtDsg.Value(wi_Row, "MqdGldAsWt") = wGldAsWt
            GrdMpQtDsg.Value(wi_Row, "MqdLabAsWt") = wLabAsWt
          End If
          
          '****** 2nd Loop of Rm to Calculate RmValue and Setting Value
          .MoveFirst
          Do While Not (.EOF Or .BOF)
            wSetRt = 0: wBaseLme = 0: wLmeSal = 0: wAlyRt = 0: wRmRt = 0: wSetVal = 0: wRmVal = 0
            
            Select Case UCase(!qNewRmCtg)
              Case Is = UCase("G"), UCase("P"), UCase("S"), UCase("L")
                wAlyRt = IIF(!qCopyRt = "Y", !OrAlySalRt, GetRmRt(ADC("MqmCmCd"), !OrAlyCd, 0, 0, 0, "S", 0, "C", , , , ADC("MqmCmCurCd"), ADC("MqmLkUpRmRt"), ctCurrPrtn))
                wLmeSal = ADC("MqmLm" + !qNewRmCtg + "Sal")
                wRmRt = IIF(!qCopyRt = "Y", !OrSalRt, Round(GetGldRt(ADC("MqmCmCd"), wLmeSal, !qNewRmCd, wAlyRt, _
                                        "DM", pv_RsSrcDsg!OdDmCd, "S", "C", , ADC("MqmCmCurCd"), "", ADC("MqmLkUpMetLs"), ctCurrPrtn), 3))
                wLmeSal = IIF(!qCopyRt = "Y", !OrLMESal, wLmeSal)
                
              Case Else
                If !qNewRmCtg = "X" Then
                  wBaseRmCtg = moCn.GetFldVal("Select RmCtg From RmMst Where RmCd='" + !qNewRmBaseCd + "'")
                  wBaseLme = IIF(UCase(wBaseRmCtg) = "G", ADC("MqmLmGSal"), IIF(UCase(wBaseRmCtg) = "P", ADC("MqmLmPSal"), IIF(UCase(wBaseRmCtg) = "S", ADC("MqmLmSSal"), IIF(UCase(wBaseRmCtg) = "L", ADC("MqmLmLSal"), 0))))
                ElseIf !qNewRmCtg = "D" Or !qNewRmCtg = "C" Then
                  wSetRt = IIF(!qCopyRt = "Y", !OrSetSalRt, Round(GetLabRt(ADC("MqmCmCd"), "SET", !OrSetSCd, "S", MWLib.Div(!qNewRmWt, !OrQty), "C", , , , ADC("MqmCmCurCd"), ADC("MqmLkUpLabRt")), 3))
                End If
                
                ' ***** Manali 3.10.0 - 18/05/12 - Round for mWtPerDia changed from 4 to 5 (vb rounds of (0.00055, 4) to 0.0005 instead of 0.0006)
                wRmRt = IIF(!qCopyRt = "Y", !OrSalRt, GetRmRt(ADC("MqmCmCd"), !qNewRmCd, !OrLn1, !OrLn2, Round(MWLib.Div(!qNewRmWt, !OrQty), 5), "S", wBaseLme, "C", , , , ADC("MqmCmCurCd"), ADC("MqmLkUpRmRt"), ctCurrPrtn))
            End Select
            
            wRmVal = IIF(!qCopyRt = "Y", !OrSalVal, wRmRt * IIF(!qNewRmQW = "Q", !OrQty, IIF(!OrMainMet = "Y", wGldAsWt, !qNewRmWt)))
            wSetVal = IIF(!qCopyRt = "Y", !OrSetSalVal, wSetRt * !OrQty)
            wTotRmVal = wTotRmVal + wRmVal
            wTotSetVal = wTotSetVal + wSetVal
            If !qNewRmCtg = "D" Then wDRmVal = wDRmVal + wRmVal
            If !qNewRmCtg = "C" Then wCRmVal = wCRmVal + wRmVal
            If !qCopyRt = "Y" Or wRmRt <> 0 Then
              
            Else
              wbRmRt = False
'              wGldAsWt = 0: wLabAsWt = 0: wDRmVal = 0: wCRmVal = 0
'              wTotRmVal = 0: wTotLabVal = 0: wDsgVal = 0: Exit Do
            End If
            
            ws_InsertRm = ws_InsertRm + _
                          " Insert Into MultiPrcQtRm (MqrCoCd, MqrTc, MqrYy, MqrChr, MqrNo, MqrSr, MqrSrNo, MqrDiagrd, MqrRmCtg, MqrRmSCtg, MqrRmCd, MqrLn1, MqrLn2, MqrQty, MqrRmPtr, MqrWt, MqrLmeSal, MqrSalRt, MqrSalVal, MqrSetScd, MqrSetSalRt, MqrSetSalVal, MqrAlyCd, MqrAlySalRt, MqrMainMet, ModUsr, ModDt, ModTime) " + _
                          " Values ('" + ADC("MqmCoCd") + "', '" + ADC("MqmTc") + "', '" + ADC("MqmYy") + "', '" + ADC("MqmChr") + "', " + CStr(ADC("MqmNo")) + ", " + CStr(wi_Sr) + ", " + CStr(!OrSrNo) + ", '" + wDiaGrd + "',  " + _
                          " '" + !qNewRmCtg + "', '" + !qNewRmSCtg + "', '" + !qNewRmCd + "', " + CStr(!OrLn1) + ", " + CStr(!OrLn2) + ", " + CStr(!OrQty) + ", " + CStr(!OrRmPtr) + ", " + CStr(!qNewRmWt) + ", " + CStr(wLmeSal) + ", " + CStr(wRmRt) + ", " + CStr(wRmVal) + ", '" + !OrSetSCd + "', " + CStr(wSetRt) + ", " + CStr(wSetVal) + ", '" + !OrAlyCd + "', " + CStr(wAlyRt) + ", '" + !OrMainMet + "', '" + UCase(gs_UsrCd) + "', '" + CStr(moCn.SrvrDate) + "', '" + CStr(moCn.SrvrTime) + "'); "
            .MoveNext
          Loop
          End With
          
          'If Not wbRmRt Then GoTo NextDiaGrd
          
          If pv_CpyFrDsgMst Then
            wLabStr = "Select DlSr as OlSrNo, DlMCd as OlMCd, DlSCd as OlSCd, PValue as OlQW, DlQty as OlQty, 0 as OlSalRt, 0 as OlSalVal From DsgLab " + _
                      "Join Param on PTyp='LABSCD' and PMCd=DlMCd and PSCd=DlSCd " + _
                      "Where DlTcTyp='DM' and DLCd='" + ADC("wCpyDsgDmCd") + "' and DlSz='' Order By OlSrNo "
          Else
            wLabStr = "Select OlSrNo, OlMCd, OlSCd, OlQW as OlQW, OlQty as OlQty, OlSalRt, OlSalVal From OrdLab " + _
                      "Join Param on PTyp='LABSCD' and PMCd=OlMCd and PSCd=OlSCd " + _
                      "Where OlCoCd='" + pv_RsSrcDsg!OdCoCd + "' and OlTc='" + pv_RsSrcDsg!OdTc + "' and OlYy='" + pv_RsSrcDsg!OdYy + "' and OlChr='" + pv_RsSrcDsg!OdChr + "' and OlNo=" + CStr(pv_RsSrcDsg!OdNo) + _
                      " and OlSr=" + CStr(pv_RsSrcDsg!OdSr) + " Order By OlSrNo "
          End If
          Set wRs_LabDet = moCn.OpenRes(wLabStr)
          
          If wRs_LabDet.RecCount > 0 Then
            With wRs_LabDet
            .MoveFirst
            Do While Not (.EOF Or .BOF)
              Dim wOlQw As String
              wLabRt = 0: wLabVal = 0
              wOlQw = !OlQw
              '5717-19 !OlQw replaced by wOlQw to get QW from lab rate
              wLabRt = IIF(ADC("wCpyVchRt") = "Y" And pv_RsSrcDsg!OdLabAsWt = wLabAsWt, !OlSalRt, GetLabRt(ADC("MqmCmCd"), !OlMCd, !OlSCd, "S", wLabAsWt, "C", wOlQw, , , ADC("MqmCmCurCd"), ADC("MqmLkUpLabRt")))
              If wOlQw = "" Then wOlQw = !OlQw '5717-19 if wOlQw="" then it store the existing value
              
              If ADC("wCpyVchRt") = "Y" And pv_RsSrcDsg!OdLabAsWt = wLabAsWt Then
                wLabVal = !OlSalVal
              Else
                If UCase(!OlMCd) = UCase(ctBrkLMCd) Then
                  wLabVal = wLabRt / 100 * (wDRmVal + wCRmVal)
                ElseIf UCase(!OlMCd) = UCase(ctDHCLMCd) Then
                  wLabVal = wLabRt / 100 * wDRmVal
                ElseIf UCase(!OlMCd) = UCase(ctGHCLMCd) Then
                  wLabVal = wLabRt / 100 * wDRmVal
                ElseIf UCase(!OlMCd) = UCase(ctCHCLMCd) Then
                  wLabVal = wLabRt / 100 * wCRmVal
                ElseIf wLabRt < 0 Then
                  wLabVal = wLabRt * !OlQty
                Else
                  '5717-19 !OlQw replaced by wOlQw to get the value
                  wLabVal = GetLabVal(ADC("MqmCmCd"), !OlMCd, !OlSCd, wLabRt, _
                                      !OlQty, wLabAsWt, wOlQw, "S", "C", ADC("MqmCmCurCd"), ADC("MqmLkUpLabRt"))
                End If
              End If
              wTotLabVal = wTotLabVal + wLabVal
              If j = 1 Then
                ws_InsertLab = ws_InsertLab + _
                              "Insert Into MultiPrcQtLab (MqlCoCd, MqlTc, MqlYy, MqlChr, MqlNo, MqlSr, MqlSrNo, MqlMcd, MqlScd, MqlQty, MqlSalRt, MqlSalVal, MqlQw, ModUsr, ModDt, ModTime) " + _
                              "Values ('" + ADC("MqmCoCd") + "', '" + ADC("MqmTc") + "', '" + ADC("MqmYy") + "', '" + ADC("MqmChr") + "', " + CStr(ADC("MqmNo")) + ", " + CStr(wi_Sr) + ", " + CStr(!OlSrNo) + ", " + _
                              "'" + !OlMCd + "', '" + !OlSCd + "', " + CStr(!OlQty) + ", " + CStr(wLabRt) + ", " + CStr(wLabVal) + ", '" + wOlQw + "', '" + UCase(gs_UsrCd) + "', '" + CStr(moCn.SrvrDate) + "', '" + CStr(moCn.SrvrTime) + "'); "
              End If
              
              .MoveNext
            Loop
            End With
          End If
          wDsgVal = wTotRmVal + wTotSetVal + wTotLabVal
          wDsgVal = FnRndOff(wDsgVal * ADC("MqmMulBy"), 2, wRsCustMst!CmHRndOrdSalPrc)
          wDsgVal = IIF(wbRmRt, wDsgVal, 0)
NextDiaGrd:
          GrdMpQtDsg.Value(wi_Row, "wMqdDiaGrd" + Trim(Str(j)) + "Prc") = wDsgVal
          Set wRs_RmDet = Nothing
          Set wRs_LabDet = Nothing
        Next
        GrdMpQtDsg.SaveRec (wi_Row)
        If ws_InsertRm <> "" Then Call moCn.Execute(ws_InsertRm)
        If ws_InsertLab <> "" Then Call moCn.Execute(ws_InsertLab)
      End If
NextMetGrd:
    Next
  
  pv_RsSrcDsg.MoveNext
  Loop
  MWLib.EndProcess Me
  GrdMpQtDsg.EndCopy
  Set pv_RsSrcDsg = Nothing
  Set wRsCustMst = Nothing
  ''If wDsgSrNotCopied <> "" Then DispMsg "Following Dsg-Sz-Kt Combination Already Exists in Voucher: " + vbCrLf + wDsgSrNotCopied, etError
Exit Sub

ErrH:
  MWLib.EndProcess Me
  GrdMpQtDsg.EndCopy
  Set wRs_RmDet = Nothing
  Set wRs_LabDet = Nothing
  Set pv_RsSrcDsg = Nothing
  Set wRsCustMst = Nothing
  DispMsg Err.Description, etError

End Sub

Private Sub DispFra(ByVal pv_GradeFra As en_MpQtFra)
  '*** This is the common procedure which is used to toggle between hiding and displaying
      'the frames (i.e. if the frame is visible then make it invisible and viseversa)
      'associated with a command button
  
  Select Case pv_GradeFra
  Case Is = CpyVch
    '*** If Frame FraCpy is visible then make it invisible and enable all the command buttons
    '*** If Frame FraCpy is invisible then make it visible and disable all the
        'command buttons except the one which is clicked. Change the background colour of
        'the clicked command button to white and make the font bold
    If FraCpyVch.Visible = True Then
      FraCpyVch.Visible = False
      FraCpyVch.Enabled = False
      Call EnaDisaCmds(False, CmdCpyVch, CpyVch)
      CmdCpyVch.SetFocus
    Else
      FraCpyVch.Visible = True
      FraCpyVch.Enabled = True
      ADC("wCpyVchCoCd").SetFocus
      FraCpyVch.ZOrder
      Call EnaDisaCmds(True, CmdCpyVch)
    End If
  
  Case Is = CpyDsg
    If FraCpyDsg.Visible = True Then
      FraCpyDsg.Visible = False
      FraCpyDsg.Enabled = False
      Call EnaDisaCmds(False, CmdCpyDsg, CpyDsg)
      CmdCpyDsg.SetFocus
    Else
      FraCpyDsg.Visible = True
      FraCpyDsg.Enabled = True
      ADC("wCpyDsgBarCdYn").SetFocus
      ''ADC("wCpyDsgDmCd").SetFocus
      FraCpyDsg.ZOrder
      Call EnaDisaCmds(True, CmdCpyDsg)
    End If
  
  Case Is = DsgDet
    If FraDsgDet.Visible = True Then
      FraDsgDet.Visible = False
      FraDsgDet.Enabled = False
      Call EnaDisaCmds(False, CmdDsgDet, DsgDet)
      CmdDsgDet.SetFocus
    Else
      FraDsgDet.Visible = True
      FraDsgDet.Enabled = True
      ADC("wMqdGldAs").SetFocus
      FraDsgDet.ZOrder
      Call SetGrdMpQtLab("S")
      Call SetGrdMpQtRm("S")
      Call EnaDisaCmds(True, CmdDsgDet)
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
      FraCpyRfId.ZOrder
      Call EnaDisaCmds(True, CmdCpyRfId)
    End If
  End Select
End Sub
Private Sub EnaDisaCmds(ByVal pv_ShowFra As Boolean, Optional ByVal pv_CmdBut As MWCTL_BTN1, Optional ByVal pv_MpQtFra As en_MpQtFra)
  If pv_ShowFra = True Then
    Call StoreState
  End If

  If Not pv_CmdBut Is Nothing Then
    If pv_ShowFra = True Then
      pv_CmdBut.CausesValidation = True
    Else
      If Not IsMissing(pv_MpQtFra) Then
        pv_CmdBut.CausesValidation = mArr_Cv(pv_MpQtFra)
      End If
    End If
  End If
  FraNKeyAll.Enabled = Not pv_ShowFra
  ADC.AllowSave = Not pv_ShowFra
  ADC.AllowDelete = Not pv_ShowFra
  
  If pv_ShowFra = True Then
    '** All command buttons are disabled
    CmdCpyVch.Enabled = Not pv_ShowFra
    CmdCpyDsg.Enabled = Not pv_ShowFra
    CmdGenExcel.Enabled = Not pv_ShowFra
    CmdDsgDet.Enabled = Not pv_ShowFra
    CmdCpyRfId.Enabled = Not pv_ShowFra
  Else
    '*** Restoring Ena/Disa state of all Command Buttons
    CmdCpyVch.Enabled = mArr_Ed(CpyVch)
    CmdCpyDsg.Enabled = mArr_Ed(CpyDsg)
    CmdGenExcel.Enabled = mArr_Ed(Excel)
    CmdDsgDet.Enabled = mArr_Ed(DsgDet)
    CmdCpyRfId.Enabled = mArr_Ed(CpyRfId)
  End If
  
  If Not pv_CmdBut Is Nothing Then pv_CmdBut.Enabled = True
End Sub

Private Sub StoreState()
  mArr_Cv(CpyVch) = CmdCpyVch.CausesValidation
  mArr_Cv(CpyDsg) = CmdCpyDsg.CausesValidation
  mArr_Cv(Excel) = CmdGenExcel.CausesValidation
  mArr_Cv(DsgDet) = CmdDsgDet.CausesValidation
  mArr_Cv(CpyRfId) = CmdCpyRfId.CausesValidation
  
  mArr_Ed(CpyVch) = CmdCpyVch.Enabled
  mArr_Ed(CpyDsg) = CmdCpyDsg.Enabled
  mArr_Ed(Excel) = CmdGenExcel.Enabled
  mArr_Ed(DsgDet) = CmdDsgDet.Enabled
  mArr_Ed(CpyRfId) = CmdCpyRfId.Enabled
  
End Sub


Private Sub HideAllFras()
  '*** This procedure hides all the frames other than FraNKeyAll
  FraNKeyAll.Visible = True
  FraCpyDsg.Visible = False
  FraCpyDsg.Enabled = False
  FraCpyVch.Visible = False
  FraCpyVch.Enabled = False
  FraDsgDet.Visible = False
  FraDsgDet.Enabled = False
  FraCpyRfId.Visible = False
  FraCpyRfId.Enabled = False
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
  ws_BtnStr = "CmdCpyGo"
  Call SetObjColors(Me, ws_BtnStr)
  '*** Jenny Colour
  
  Me.KeyPreview = True
End Sub
Private Sub Form_KeyUp(KeyCode As Integer, Shift As Integer)
  If ADC.HotKeys(KeyCode, Shift) Then
    KeyCode = 0
    Shift = 0
  End If
End Sub
Private Sub Form_Activate()
  Set m_ObjCmdGetRmRt = ADC.Connection.OpenQuery("ForCpy_GetRmRt")
  Set m_ObjCmdGetLabRt = ADC.Connection.OpenQuery("ForCpy_GetLabRt")

  Call FrmActivate(Me)
End Sub
Private Sub Form_Deactivate()
  If Not (m_ObjCmdGetRmRt Is Nothing) Then m_ObjCmdGetRmRt.QryClose
  If Not (m_ObjCmdGetLabRt Is Nothing) Then m_ObjCmdGetLabRt.QryClose

  Call FrmDeActivate(Me)
End Sub
Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    If UnloadMode <> vbFormCode Then Cancel = True: Exit Sub
    If Not m_ObjCmdGetRmRt Is Nothing Then m_ObjCmdGetRmRt.QryClose
    If Not m_ObjCmdGetLabRt Is Nothing Then m_ObjCmdGetLabRt.QryClose

    Set moCn = Nothing    '*** (09/08/05)    '*** (Bef speed) bef 24/11/06
End Sub

Private Sub ATXT_Validate(Index As Integer, Cancel As Boolean)
  Cancel = ATXT(Index).Validate
End Sub

Private Sub GrdMpQtDsg_Valid()
'CmdDsgDet.Enabled = False
End Sub

Private Sub GrdMpQtDsg_Validate(Cancel As Boolean)
Cancel = GrdMpQtDsg.Validate
End Sub

Private Sub GrdMpQtDsg_When()
  If GrdMpQtDsg.Rows = 1 Then GrdPropSet
  CmdDsgDet.Enabled = True
End Sub

Private Sub GrdPropSet()
  Dim wi_SpaceLen As Single, i As Integer
  Dim wDiaGrd As String, wDiaGrdPrc As String
    
  For i = 1 To 8
    wDiaGrd = "MqmDiaGrd" + Trim(Str(i))
    wDiaGrdPrc = "wMqdDiaGrd" + Trim(Str(i)) + "Prc"
    wi_SpaceLen = (10 * 1.7) - (Len(ADC(wDiaGrd)))    '2.5
    GrdMpQtDsg.ColProp(wDiaGrdPrc).Style = IIF(ADC(wDiaGrd) = "", fgcinvisible, fgcNormal)
    GrdMpQtDsg.ColProp(wDiaGrdPrc).Heading = ADC(wDiaGrd) + IIF(wi_SpaceLen > 0, Space(Abs(wi_SpaceLen)), "")
  
  Next
  GrdMpQtDsg.SetColWidthToHd
End Sub

Private Sub GenExcel()
On Error GoTo Err_Hndlr
If GrdMpQtDsg.Rows = 1 Then Exit Sub

Dim wErrMsg As String

Dim wExcVer As String, wExtn As String
Dim wOutputPath As String
Dim sFile, sFilePath As String

Set oExcel = CreateObject("Excel.Application")

wExcVer = oExcel.version
wExtn = IIF(wExcVer >= "12.0", ".xlsx", ".xls")

sFile = gs_CoCd + "-" + gs_UsrCd + "-" + Format(Now, DtFullStr()) + wExtn   '6.1


wOutputPath = moCn.GetFldVal("Select PDesc225 from Param Where PTyp='OUTPUT' and PMCd='OUTPUT' ")
sFilePath = wOutputPath + "\" + sFile

If Dir(sFilePath, vbNormal) <> "" Then
  Kill sFilePath
End If

FileCopy App.Path + "\Template\MultiPrcQt" + wExtn, sFilePath

'Bfore - 4.1.0.0
'Set oBook = oExcel.Workbooks.Open(sFilePath)

'4.1.0.0
Set oBook = oExcel.Workbooks.Open(sFilePath, Password:=ctXCelPwd, WriteResPassword:=ctXCelWritePwd)


Set oSheet = oBook.worksheets(1)

oExcel.Visible = False: oExcel.DisplayAlerts = False
'oExcel.Visible = True

Dim wSqlStr As String, wRsExcel As MDORowSet, wRsParam As MDORowSet
Dim i As Integer, wFRow As Integer, wFPrcCol As Integer
Dim wDmCd As String, wDmSz As String, wDmCdSzSr As Integer, wDiaWt As Single
Dim wCurrRow As Integer, wPictPath As String, wb_Pic As Boolean
Dim wDmCdPicRows As Integer

wSqlStr = " Open Symmetric Key DataEnde Decryption By Password = 'RMeWm'; " + _
        "Select (Case MqdMetGrd When MqmMetGrd1 Then 1 When MqmMetGrd2 Then 2 When MqmMetGrd3 Then 3 When MqmMetGrd4 Then 4 Else 99 End) as qMetGrdNo, " + _
        "MqdNo , MqdSr, MqdDmcd, MqdDmSz, MqdMetGrd, PValue as qMetGrdDesc, MqdGrsWt, MqdDiaWt, " + _
        " Cast(Convert(varChar(100), DecryptByKey(Convert(varBinary(100), MqdDiaGrd1PrcBin))) as Decimal(16,2)) as MqdDiaGrd1Prc, " + _
        " Cast(Convert(varChar(100), DecryptByKey(Convert(varBinary(100), MqdDiaGrd1PrcBin))) as Decimal(16,2)) as MqdDiaGrd1Prc, " + _
        "Cast(Convert(varChar(100), DecryptByKey(Convert(varBinary(100), MqdDiaGrd2PrcBin))) as Decimal(16,2)) as MqdDiaGrd2Prc, " + _
        "Cast(Convert(varChar(100), DecryptByKey(Convert(varBinary(100), MqdDiaGrd3PrcBin))) as Decimal(16,2)) as MqdDiaGrd3Prc, " + _
        "Cast(Convert(varChar(100), DecryptByKey(Convert(varBinary(100), MqdDiaGrd4PrcBin))) as Decimal(16,2)) as MqdDiaGrd4Prc, " + _
        "Cast(Convert(varChar(100), DecryptByKey(Convert(varBinary(100), MqdDiaGrd5PrcBin))) as Decimal(16,2)) as MqdDiaGrd5Prc, " + _
        "Cast(Convert(varChar(100), DecryptByKey(Convert(varBinary(100), MqdDiaGrd6PrcBin))) as Decimal(16,2)) as MqdDiaGrd6Prc, " + _
        "Cast(Convert(varChar(100), DecryptByKey(Convert(varBinary(100), MqdDiaGrd7PrcBin))) as Decimal(16,2)) as MqdDiaGrd7Prc, " + _
        "Cast(Convert(varChar(100), DecryptByKey(Convert(varBinary(100), MqdDiaGrd8PrcBin))) as Decimal(16,2)) as MqdDiaGrd8Prc, " + _
        "MqmCmCd, CmName, MqmDt, MqmLmgSal, MqmLmpSal, MqmLmsSal, MqmLmlSal, MqmLkUpRmRt, MqmLkUpLabRt, MqmLkUpMetLs, " + _
        "MqmMetGrd1, MqmMetGrd2, MqmMetGrd3, MqmMetGrd4, " + _
        "MqmDiaGrd1 , MqmDiaGrd2, MqmDiaGrd3, MqmDiaGrd4, MqmDiaGrd5, MqmDiaGrd6, MqmDiaGrd7, MqmDiaGrd8 " + _
        "From MultiPrcQtDsg " + _
        "Join MultiPrcQtMst On MqmNo=MqdNo " + _
        "Join CustMst On CmCtg='C' and CmCd=MqmCmCd " + _
        "Join Param On PTyp='GRDCD' and PMCd='MET' and PSCd=MqdMetGrd " + _
        "Where MqmCoCd='" + ADC("MqmCoCd") + "' and MqmTc='" + ADC("MqmTc") + "' and MqmYy='" + ADC("MqmYy") + "' and " + _
        "MqmChr='" + ADC("MqmChr") + "' and MqmNo = " + CStr(ADC("MqmNo")) + " " + _
        "Order By MqdDmCd, MqdDmSz, MqdDiaWt, qMetGrdNo, MqdSr ;Close Symmetric Key DataEnde;"
        
               
Set wRsExcel = moCn.OpenRes(wSqlStr)

If wRsExcel.RecCount = 0 Then Exit Sub

MWLib.BeginProcess Me, "Generating Excel ..."

With wRsExcel
.MoveFirst

oSheet.Cells(1, 2) = !CmName
oSheet.Cells(1, 8) = !MqmDt
oSheet.Cells(2, 2) = !MqmLmgSal
oSheet.Cells(2, 4) = !MqmLmpSal
oSheet.Cells(2, 6) = !MqmLmsSal
oSheet.Cells(2, 8) = !MqmLmlSal

wFRow = 5
wCurrRow = wFRow - 1
wFPrcCol = 5
wDmCdPicRows = 5

Dim wDiaGrdNo As Integer
Set wRsParam = moCn.OpenRes("Select PTyp, PMCd, PSCd, PDesc, PValue, " + _
                          " (Case PSCd When '" + !MqmDiaGrd1 + "' Then 1 When '" + !MqmDiaGrd2 + "' Then 2 When '" + !MqmDiaGrd3 + "' Then 3 When '" + !MqmDiaGrd4 + "' Then 4 When '" + !MqmDiaGrd5 + "' Then 5 When '" + !MqmDiaGrd6 + "' Then 6 When '" + !MqmDiaGrd7 + "' Then 7 When '" + !MqmDiaGrd8 + "' Then 8 Else 99 End) as qDiaGrdNo " + _
                          " From Param Where PTyp='GRDCD' and PMCd='DIA' " + _
                          " and PSCd In ('" + !MqmDiaGrd1 + "', '" + !MqmDiaGrd2 + "', '" + !MqmDiaGrd3 + "', '" + !MqmDiaGrd4 + "', '" + !MqmDiaGrd5 + "', '" + !MqmDiaGrd6 + "', '" + !MqmDiaGrd7 + "', '" + !MqmDiaGrd8 + "') " + _
                          " Order By qDiaGrdNo ")

wRsParam.MoveFirst
Do While Not (wRsParam.EOF Or wRsParam.BOF)
  wDiaGrdNo = wRsParam!qDiaGrdNo
  oSheet.Cells(wCurrRow, wDiaGrdNo + wFPrcCol) = wRsParam!pValue
  oSheet.Cells(wCurrRow, wDiaGrdNo + wFPrcCol).Interior.Color = RGB(221, 221, 221)
  wRsParam.MoveNext
Loop

oSheet.Cells.Range(oSheet.Cells(wCurrRow, 1), oSheet.Cells(wCurrRow, wDiaGrdNo + wFPrcCol)).Borders.LineStyle = 1
oSheet.Cells.Range(oSheet.Cells(wCurrRow, 1), oSheet.Cells(wCurrRow, wDiaGrdNo + wFPrcCol)).Borders.Weight = m_xlMedium


wDmCd = "": wDmSz = "": wDiaWt = 0
wb_Pic = False
Do While Not (.EOF Or .BOF)
  If !MqdDmCd + "," + !MqdDmSz + "," + CStr(!MqdDiaWt) <> wDmCd + "," + wDmSz + "," + CStr(wDiaWt) Then
    If wDmCd <> "" Then
        If wb_Pic And wDmCdSzSr - 1 <= wDmCdPicRows Then
          wCurrRow = wCurrRow + wDmCdPicRows - wDmCdSzSr + 1
        End If
        oSheet.Cells.Range(oSheet.Cells(wCurrRow - 1, 1), oSheet.Cells(wCurrRow - 1, wDiaGrdNo + wFPrcCol)).Borders(m_xlEdgeBottom).LineStyle = 1
        oSheet.Cells.Range(oSheet.Cells(wCurrRow - 1, 1), oSheet.Cells(wCurrRow - 1, wDiaGrdNo + wFPrcCol)).Borders(m_xlEdgeBottom).Weight = m_xlThin
        
        oSheet.Range(oSheet.Cells(wFRow, 1), oSheet.Cells(wFRow, wDiaGrdNo + wFPrcCol)).Copy
        oSheet.Range(oSheet.Cells(wCurrRow + 1, 1), oSheet.Cells(wCurrRow + 1, wDiaGrdNo + wFPrcCol)).Select
        oSheet.Paste
        oSheet.Range(oSheet.Cells(wCurrRow + 1, 1), oSheet.Cells(wCurrRow + 1, wDiaGrdNo + wFPrcCol)).Value = ""
    End If
    wDmCdSzSr = 1
    wCurrRow = wCurrRow + 1
    
    oSheet.Cells(wCurrRow, 1) = !MqdDmCd
    oSheet.Cells(wCurrRow, 2) = !MqdDmSz
    oSheet.Cells(wCurrRow, 3) = !MqdDiaWt

    oSheet.Cells.Range(oSheet.Cells(wCurrRow, 1), oSheet.Cells(wCurrRow, wDiaGrdNo + wFPrcCol)).Borders.LineStyle = 1
    oSheet.Cells.Range(oSheet.Cells(wCurrRow, 1), oSheet.Cells(wCurrRow, wDiaGrdNo + wFPrcCol)).Borders.Weight = m_xlThin
    
    If ADC("wShowExcPic") = "Y" Then
      wPictPath = GetPictPath("DM", False, !MqdDmCd, !MqdDmSz)
      If Dir(wPictPath) <> "" And wPictPath <> "" Then
        oSheet.Shapes.AddPicture wPictPath, False, True, oSheet.Cells(wCurrRow, 1).Left, oSheet.Cells(wCurrRow + 1, 1).Top, oSheet.Rows(CStr(wCurrRow) + ":" + CStr(wCurrRow + 3)).Height, oSheet.Rows(CStr(wCurrRow) + ":" + CStr(wCurrRow + 3)).Height
        oSheet.Pictures.ShapeRange.LockAspectRatio = False
        oSheet.Pictures(1).Name = !MqdDmCd + "-" + !MqdDmSz
        oSheet.Pictures.Border.LineStyle = 1
        wb_Pic = True
      Else
        wb_Pic = False
      End If
    End If
  Else
    oSheet.Range(oSheet.Cells(wFRow, 1), oSheet.Cells(wFRow, wDiaGrdNo + wFPrcCol)).Copy
    oSheet.Range(oSheet.Cells(wCurrRow, 1), oSheet.Cells(wCurrRow, wDiaGrdNo + wFPrcCol)).Select
    oSheet.Paste
    oSheet.Range(oSheet.Cells(wCurrRow, 1), oSheet.Cells(wCurrRow, wDiaGrdNo + wFPrcCol)).Value = ""
  
    oSheet.Cells.Range(oSheet.Cells(wCurrRow, 1), oSheet.Cells(wCurrRow, 3)).Borders(m_xlInsideVertical).LineStyle = m_xlNone
    oSheet.Cells.Range(oSheet.Cells(wCurrRow, 1), oSheet.Cells(wCurrRow, 3)).Borders(m_xlInsideHorizontal).LineStyle = m_xlNone
    If wDmCdSzSr > 2 Then oSheet.Cells.Range(oSheet.Cells(wCurrRow, 1), oSheet.Cells(wCurrRow, 3)).Borders(m_xlEdgeTop).LineStyle = m_xlNone
    oSheet.Cells.Range(oSheet.Cells(wCurrRow, 1), oSheet.Cells(wCurrRow, 3)).Borders(m_xlEdgeBottom).LineStyle = m_xlNone
  
  End If
  oSheet.Cells(wCurrRow, 5) = !qMetGrdDesc    '''!MqdMetGrd
  oSheet.Cells(wCurrRow, 5).Interior.Color = RGB(221, 221, 221)
  oSheet.Cells(wCurrRow, 4) = !MqdGrsWt
  oSheet.Cells(wCurrRow, 6) = IIF(!MqmDiaGrd1 <> "", !MqdDiaGrd1Prc, "")
  oSheet.Cells(wCurrRow, 7) = IIF(!MqmDiaGrd2 <> "", !MqdDiaGrd2Prc, "")
  oSheet.Cells(wCurrRow, 8) = IIF(!MqmDiaGrd3 <> "", !MqdDiaGrd3Prc, "")
  oSheet.Cells(wCurrRow, 9) = IIF(!MqmDiaGrd4 <> "", !MqdDiaGrd4Prc, "")
  oSheet.Cells(wCurrRow, 10) = IIF(!MqmDiaGrd5 <> "", !MqdDiaGrd5Prc, "")
  oSheet.Cells(wCurrRow, 11) = IIF(!MqmDiaGrd6 <> "", !MqdDiaGrd6Prc, "")
  oSheet.Cells(wCurrRow, 12) = IIF(!MqmDiaGrd7 <> "", !MqdDiaGrd7Prc, "")
  oSheet.Cells(wCurrRow, 13) = IIF(!MqmDiaGrd8 <> "", !MqdDiaGrd8Prc, "")
  wDmCd = !MqdDmCd: wDmSz = !MqdDmSz: wDiaWt = !MqdDiaWt
  wCurrRow = wCurrRow + 1
  wDmCdSzSr = wDmCdSzSr + 1
  .MoveNext
Loop
End With

If wb_Pic And wDmCdSzSr - 1 <= wDmCdPicRows Then
  wCurrRow = wCurrRow + wDmCdPicRows - wDmCdSzSr + 1
End If

oSheet.Cells.Range(oSheet.Cells(wCurrRow - 1, 1), oSheet.Cells(wCurrRow - 1, wDiaGrdNo + 5)).Borders(m_xlEdgeBottom).LineStyle = 1
oSheet.Cells.Range(oSheet.Cells(wCurrRow - 1, 1), oSheet.Cells(wCurrRow - 1, wDiaGrdNo + 5)).Borders(m_xlEdgeBottom).Weight = m_xlThin
  
oSheet.Cells.Range(oSheet.Cells(wFRow, 1), oSheet.Cells(wCurrRow - 1, wDiaGrdNo + 5)).Borders(m_xlEdgeBottom).LineStyle = 1
oSheet.Cells.Range(oSheet.Cells(wFRow, 1), oSheet.Cells(wCurrRow - 1, wDiaGrdNo + 5)).Borders(m_xlEdgeBottom).Weight = m_xlMedium
oSheet.Cells.Range(oSheet.Cells(wFRow, 1), oSheet.Cells(wCurrRow - 1, wDiaGrdNo + 5)).Borders(m_xlEdgeTop).LineStyle = 1
oSheet.Cells.Range(oSheet.Cells(wFRow, 1), oSheet.Cells(wCurrRow - 1, wDiaGrdNo + 5)).Borders(m_xlEdgeTop).Weight = m_xlMedium
oSheet.Cells.Range(oSheet.Cells(wFRow, 1), oSheet.Cells(wCurrRow - 1, wDiaGrdNo + 5)).Borders(m_xlEdgeLeft).LineStyle = 1
oSheet.Cells.Range(oSheet.Cells(wFRow, 1), oSheet.Cells(wCurrRow - 1, wDiaGrdNo + 5)).Borders(m_xlEdgeLeft).Weight = m_xlMedium
oSheet.Cells.Range(oSheet.Cells(wFRow, 1), oSheet.Cells(wCurrRow - 1, wDiaGrdNo + 5)).Borders(m_xlEdgeRight).LineStyle = 1
oSheet.Cells.Range(oSheet.Cells(wFRow, 1), oSheet.Cells(wCurrRow - 1, wDiaGrdNo + 5)).Borders(m_xlEdgeRight).Weight = m_xlMedium
  
'oExcel.ActiveWorkbook.Save
oExcel.ActiveWorkbook.SaveAs sFilePath, , "", ""       '4.1.0.0 (Pwd Removed)
  
MWLib.EndProcess Me
  
oExcel.Visible = True
  
Set wRsExcel = Nothing
Set wRsParam = Nothing
Set oExcel = Nothing

Exit Sub
Err_Hndlr:
  MWLib.EndProcess Me
  If Err.Number = 32755 Then
    wErrMsg = "Operation Cancelled"
    oExcel.ActiveWorkbook.Close
    oExcel.QUIT
  ElseIf Err.Number = 70 Then       '4.1.4
    wErrMsg = "Cannot Access Template File. Close If Already Open and Continue."
  Else
    wErrMsg = Err.Description
  End If
  Err.Clear
  Set wRsExcel = Nothing
  Set wRsParam = Nothing
  Set oExcel = Nothing
  DispMsg wErrMsg, etError
End Sub
Private Sub SetGrdMpQtRm(mOpt)
  '*** Procedure to make the fields in the grid GrdOrdRm visible or invisible according to
      'requirement for display purposes
  '*** If the parameter passed is 'S' then make the fields invisible
  '*** If the parameter passed is 'E' then make the fields visible except the fields which
      'were kept invisible at design time
  
  If mOpt = "S" Then GrdMpQtLab.ZOrder
  Debug.Print "mOpt " + mOpt
End Sub
Private Sub SetGrdMpQtLab(mOpt)
  '*** Procedure to make the fields if the grid GrdMpQtLab visible or invisible according to
      'requirement for display purposes
  '*** If the parameter passed is 'S' then make the fields invisible
  '*** If the parameter passed is 'E' then make the fields visible except the fields which
      'were kept invisible at design time

  Static WLEFT As Long
  
  Static wtmp As Long
  Dim i As Integer
  
  If mOpt = "S" Then
    If mi_LabGrdLeft = 0 Then mi_LabGrdLeft = GrdMpQtLab.Left
    WLEFT = mi_LabGrdLeft
    wtmp = 0
    For i = 4 To 6
      wtmp = wtmp + GrdMpQtLab.ColProp(i).Width
      GrdMpQtLab.ColProp(i).Style = fgcinvisible
    Next i
    GrdMpQtLab.Left = wtmp + WLEFT
    Debug.Print "GrdMpQtLab.Left S " + CStr(GrdMpQtLab.Left)
    
  End If
  
  If mOpt = "E" Then
    If WLEFT <> 0 Then
      GrdMpQtLab.Left = WLEFT
      GrdMpQtLab.Width = GrdMpQtLab.Width + wtmp
      GrdMpQtLab.ColProp(4).Style = fgcReadOnly
      GrdMpQtLab.ColProp(5).Style = fgcReadOnly
      GrdMpQtLab.ColProp(6).Style = fgcReadOnly
    End If
    Debug.Print "GrdMpQtLab.Left E " + CStr(GrdMpQtLab.Left)
  End If
End Sub


Private Sub GrdMpQtLab_GotFocus()
  SetGrdMpQtLab "E"
  GrdMpQtLab.ZOrder
End Sub
Private Sub GrdMpQtLab_Valid()
SetGrdMpQtLab "S"
End Sub

Private Sub GrdMpQtLab_Validate(Cancel As Boolean)
Cancel = GrdMpQtLab.Validate
End Sub

Private Sub GrdMpQtRm_GotFocus()
SetGrdMpQtRm "E"
GrdMpQtRm.ZOrder
End Sub

Private Sub GrdMpQtRm_Valid()
SetGrdMpQtRm "S"
End Sub

Private Sub GrdMpQtRm_Validate(Cancel As Boolean)
Cancel = GrdMpQtRm.Validate
End Sub

Private Sub txtCpyDsgDmCd_GotFocus()
  DispMsg txtCpyDsgDmCd.ToolTipText, etInfo
End Sub

Private Sub txtCpyDsgDmCd_LostFocus()
Call txtCpyDsgDmCd_Validate(False)
End Sub

Private Sub txtCpyDsgDmCd_Validate(Cancel As Boolean)
On Error GoTo ErrH
    ' ****** Manali 3.10.0
    ' ***** Vb TextBox added for BarCode - Reasons
    ' ***** 1) In case of reading from bar code, Enter Key is trapped, but enter key is also used in ATXT control for selecting value from Help
    ' ***** it was not possible to differentiate the reason for enter key press.
    ' ***** 2) If focus from an ATXT control moved any where else THROUGH CODE, validations for ATXT controls (fld_change event) are not fired
    
    Dim wDmCdOld As String
    
    Cancel = txtCpyDsgDmCd.text <> "" And Not moCn.RecSeek("Select 'x' From DsgMst where DmTcTyp= 'DM' and DmCd = '" + txtCpyDsgDmCd.text + "' and DmSz='' ")
    If Cancel Then DispMsg "Invalid Design Code", etError: Exit Sub
    
    'If txtCpyDsgDmCd.text <> "" Then
      wDmCdOld = ADC("wCpyDsgDmCd")
      ADC("wCpyDsgDmCd") = UCase(txtCpyDsgDmCd.text)
      ADC("wCpyDsgDmSz") = ""
      ADC("wCpyDsgSfx") = ""
      Call ATXT_FldChange(99, UCase("wCpyDsgDmCd"), Cancel, "", txtCpyDsgDmCd.text)
      If Cancel = True Then DispMsg "Invalid Design Code", etError: ADC("wCpyDsgDmCd") = wDmCdOld: Exit Sub
      
    'End If
    Exit Sub
ErrH:
   Cancel = True
   DispMsg "Invalid Design Code", etError: ADC("wCpyDsgDmCd") = "": Exit Sub
End Sub

''  Dim wRs_MpqRm As MDORowSet, wRs_MpqLab As MDORowSet
  
''  Set wRs_MpqRm = moCn.OpenResultset("Select MqrCoCd, MqrTc, MqrYy, MqrChr, MqrNo, MqrSr, MqrSrNo, MqrDiagrd, MqrMqdIdNo, MqrRmCtg, MqrRmSCtg, " + _
''            "MqrRmCd, MqrLn1, MqrLn2, MqrQty, MqrRmPtr, MqrWt, MqrLmeSal, MqrSalRt, MqrSalVal, MqrSetScd, MqrSetSalRt, MqrSetSalVal, " + _
''            "MqrAlyCd, MqrAlySalRt, MqrMainMet, ModUsr, ModDt, ModTime " + _
''            "From MultiPrcQtRm Where MqrCoCd='" + ADC("MqmCoCd") + "' and MqrTc='" + ADC("MqmTc") + "' and MqrYy='" + ADC("MqmYy") + "' and MqrChr='" + ADC("MqmChr") + "' and MqrNo=" + CStr(ADC("MqmNo")) + " ")

''            wRs_MpqRm.Addnew
''            wRs_MpqRm.Update
''            wRs_MpqRm!MqrCoCd = ADC("MqmCoCd")
''            wRs_MpqRm!MqrTc = ADC("MqmTc")
''            wRs_MpqRm!MqrYy = ADC("MqmYy")
''            wRs_MpqRm!MqrChr = ADC("MqmChr")
''            wRs_MpqRm!MqrNo = ADC("MqmNo")
''            wRs_MpqRm!MqrSr = wi_Sr
''            wRs_MpqRm!MqrSrNo = !OrSrNo
''            wRs_MpqRm!MqrDiaGrd = wDiaGrd
''            wRs_MpqRm!MqrRmCtg = !qNewRmCtg
''            wRs_MpqRm!MqrRmSCtg = !qNewRmSCtg
''            wRs_MpqRm!MqrRmCd = !qNewRmCd
''            wRs_MpqRm!MqrLn1 = !OrLn1
''            wRs_MpqRm!MqrLn2 = !OrLn2
''            wRs_MpqRm!MqrQty = !OrQty
''            wRs_MpqRm!MqrRmPtr = !OrRmPtr
''            wRs_MpqRm!MqrWt = !qNewRmWt
''            wRs_MpqRm!MqrLmeSal = !OrLmeSal
''            wRs_MpqRm!MqrSalRt = IIF(!qCopyRt = "Y", !OrSalRt, wRmRt)
''            wRs_MpqRm!MqrSalVal = wRmVal
''            wRs_MpqRm!MqrSetScd = !OrSetSCd
''            wRs_MpqRm!MqrSetSalRt = IIF(!qCopyRt = "Y", !OrSetSalRt, wSetRt)
''            wRs_MpqRm!MqrSetSalVal = wSetVal
''            wRs_MpqRm!MqrAlyCd = !OrAlyCd
''            wRs_MpqRm!MqrAlySalRt = IIF(!qCopyRt = "Y", !OrAlySalRt, wAlyRt)
''            wRs_MpqRm!MqrMainMet = !OrMainMet
''            wRs_MpqRm!ModUsr = UCase(gs_UsrCd)
''            wRs_MpqRm!ModDt = moCn.SrvrDate
''            wRs_MpqRm!ModTime = moCn.SrvrTime
''            wRs_MpqRm!MqrMqdIdNo = 0

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

