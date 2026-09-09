VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "tabctl32.ocx"
Object = "{AB2CC32A-E652-42A2-B341-26ECE02F6386}#2.1#0"; "mwfctlforspeed9.ocx"
Begin VB.Form EmrFrmRetMemo 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Return Memo Entry"
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
   Begin MwfCtl.MWCTL_BTN1 CmdSumm 
      Height          =   495
      Left            =   7890
      TabIndex        =   12
      ToolTipText     =   "Summary Option"
      Top             =   9210
      Width           =   855
      _ExtentX        =   1508
      _ExtentY        =   873
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
   Begin MwfCtl.MWCTL_BTN1 CmdDsgHlp 
      Height          =   495
      Left            =   7035
      TabIndex        =   11
      ToolTipText     =   "Option To View Design Help"
      Top             =   9210
      Width           =   855
      _ExtentX        =   1508
      _ExtentY        =   873
      BackColor       =   -2147483633
      MaskColor       =   8454143
      Caption         =   "Design &Help"
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
      Height          =   495
      Left            =   6180
      TabIndex        =   10
      ToolTipText     =   "Option To Copy Details From Other Design"
      Top             =   9210
      Width           =   855
      _ExtentX        =   1508
      _ExtentY        =   873
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
      Left            =   10380
      TabIndex        =   32
      TabStop         =   0   'False
      Top             =   9390
      Visible         =   0   'False
      Width           =   555
      _ExtentX        =   979
      _ExtentY        =   714
   End
   Begin MwfCtl.MWCTL_ADC ADC 
      CausesValidation=   0   'False
      Height          =   555
      Left            =   60
      TabIndex        =   33
      Top             =   9195
      Width           =   14985
      _ExtentX        =   26432
      _ExtentY        =   979
      BtnWidth        =   885
      BtnHeight       =   435
      BtnMaskColor    =   8454143
      FontSz          =   8.25
   End
   Begin VB.Frame FRA_KEYS 
      BorderStyle     =   0  'None
      Height          =   270
      Left            =   0
      TabIndex        =   35
      Top             =   0
      Width           =   4590
      Begin MwfCtl.MWCTL_MED ATXT 
         Height          =   285
         Index           =   0
         Left            =   2775
         TabIndex        =   1
         ToolTipText     =   "Enter Voucher Character"
         Top             =   0
         Width           =   570
         _ExtentX        =   1005
         _ExtentY        =   503
         MaxLength       =   3
         DataType        =   4
         DataField       =   "ReChr"
         IdName          =   "RECHR"
      End
      Begin MwfCtl.MWCTL_MED ATXT 
         Height          =   285
         Index           =   87
         Left            =   4350
         TabIndex        =   41
         ToolTipText     =   "Enter Company Code"
         Top             =   0
         Visible         =   0   'False
         Width           =   270
         _ExtentX        =   476
         _ExtentY        =   503
         Enabled         =   0   'False
         MaxLength       =   3
         DataType        =   4
         DataField       =   "ReCoCd"
         IdName          =   "RECOCD"
      End
      Begin MwfCtl.MWCTL_MED ATXT 
         Height          =   285
         Index           =   1
         Left            =   2280
         TabIndex        =   0
         ToolTipText     =   "Enter Voucher Year"
         Top             =   0
         Width           =   330
         _ExtentX        =   582
         _ExtentY        =   503
         MaxLength       =   2
         DataType        =   4
         DataField       =   "ReYy"
         IdName          =   "REYY"
      End
      Begin MwfCtl.MWCTL_MED ATXT 
         Height          =   285
         Index           =   42
         Left            =   3450
         TabIndex        =   2
         ToolTipText     =   "Enter Voucher No."
         Top             =   0
         Width           =   750
         _ExtentX        =   1323
         _ExtentY        =   503
         Alignment       =   2
         Mask            =   "#####0"
         MaxLength       =   6
         DataType        =   1
         DataField       =   "ReNo"
         IdName          =   "RENO"
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
         Left            =   2640
         TabIndex        =   43
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
         Left            =   3375
         TabIndex        =   42
         Top             =   0
         Width           =   105
      End
      Begin VB.Label LblDmCd 
         BackStyle       =   0  'Transparent
         Caption         =   "Voucher No"
         BeginProperty Font 
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
         TabIndex        =   36
         Top             =   0
         Width           =   1905
      End
   End
   Begin VB.Frame FRA_NKEYS 
      BorderStyle     =   0  'None
      Enabled         =   0   'False
      Height          =   9510
      Left            =   -120
      TabIndex        =   34
      Top             =   0
      Width           =   15285
      Begin VB.Frame FraDsgHlp 
         Height          =   8190
         Left            =   300
         TabIndex        =   58
         Top             =   840
         Width           =   14730
         Begin MwfCtl.MWCTL_MED ATXT 
            Height          =   285
            Index           =   23
            Left            =   855
            TabIndex        =   26
            ToolTipText     =   "Enter Design Code"
            Top             =   480
            Width           =   2385
            _ExtentX        =   4207
            _ExtentY        =   503
            MaxLength       =   15
            DataType        =   4
            IdName          =   "WHLPDMCD"
         End
         Begin MwfCtl.MWCTL_MED ATXT 
            Height          =   285
            Index           =   17
            Left            =   4185
            TabIndex        =   27
            ToolTipText     =   "Enter Design Suffix"
            Top             =   480
            Width           =   1635
            _ExtentX        =   2884
            _ExtentY        =   503
            MaxLength       =   15
            DataType        =   4
            IdName          =   "WHLPSFX"
         End
         Begin MwfCtl.MWCTL_MED ATXT 
            Height          =   285
            Index           =   18
            Left            =   6705
            TabIndex        =   28
            ToolTipText     =   "Enter Design Size"
            Top             =   480
            Width           =   885
            _ExtentX        =   1561
            _ExtentY        =   503
            MaxLength       =   5
            DataType        =   4
            IdName          =   "WHLPDMSZ"
         End
         Begin MwfCtl.MWCTL_MED ATXT 
            Height          =   285
            Index           =   19
            Left            =   8445
            TabIndex        =   29
            ToolTipText     =   "Enter Help Sort Option"
            Top             =   480
            Width           =   1365
            _ExtentX        =   2408
            _ExtentY        =   503
            MaxLength       =   8
            DataType        =   4
            IdName          =   "WHLPSRT"
         End
         Begin MwfCtl.MWCTL_FLX GrdDsgHlp 
            Height          =   6990
            Left            =   150
            TabIndex        =   31
            ToolTipText     =   "Invoice Design Details"
            Top             =   1050
            Width           =   14475
            _ExtentX        =   25532
            _ExtentY        =   12330
            Cols            =   16
            AllowAdd        =   0   'False
            AllowDelete     =   0   'False
            colname1        =   "WHLPINVDT"
            heading1        =   "Inv Date "
            datatype1       =   3
            maxlength1      =   8
            tooltiptext1    =   "Invoice Date"
            colname2        =   "WHLPDSGCD"
            heading2        =   "Design                            "
            datatype2       =   4
            maxlength2      =   15
            tooltiptext2    =   "Design Code"
            colname3        =   "WHLPDSGSFX"
            heading3        =   "Suffix                 "
            datatype3       =   4
            maxlength3      =   15
            tooltiptext3    =   "Design Suffix"
            colname4        =   "WHLPDSGSZ"
            heading4        =   "Size      "
            datatype4       =   4
            maxlength4      =   5
            tooltiptext4    =   "Design Size"
            colname5        =   "WHLPIDTC"
            heading5        =   "Tc    "
            datatype5       =   4
            maxlength5      =   3
            tooltiptext5    =   "Invoice Tc"
            colname6        =   "WHLPIDYY"
            heading6        =   "Yy  "
            datatype6       =   4
            maxlength6      =   2
            tooltiptext6    =   "Enter Invoice Year"
            colname7        =   "WHLPIDCHR"
            heading7        =   "Chr    "
            datatype7       =   4
            maxlength7      =   3
            tooltiptext7    =   "Invoice Character"
            colname8        =   "WHLPIDNO"
            heading8        =   "No            "
            datatype8       =   1
            mask8           =   "######0"
            maxlength8      =   7
            tooltiptext8    =   "Invoice Number"
            colname9        =   "WHLPIDSR"
            heading9        =   "Sr      "
            datatype9       =   1
            mask9           =   "###0"
            maxlength9      =   4
            tooltiptext9    =   "Invoice Serial"
            colname10       =   "WHLPIDQTY"
            heading10       =   "Qty           "
            datatype10      =   2
            mask10          =   "####0.0"
            maxlength10     =   7
            tooltiptext10   =   "Design Quantity"
            colname11       =   "WHLPIDFOBPPC"
            heading11       =   "FOB Per Pc  "
            datatype11      =   2
            mask11          =   "######0.00"
            maxlength11     =   10
            tooltiptext11   =   "Invoice Design FOB Per Piece"
            colname12       =   "WHLPIDEXPCMCD"
            heading12       =   "Order Cust      "
            datatype12      =   4
            maxlength12     =   8
            tooltiptext12   =   "Export Order Customer"
            colname13       =   "WHLPIDKT"
            heading13       =   "KT         "
            datatype13      =   4
            maxlength13     =   5
            style13         =   2
            colname14       =   "WHLPDMCOL"
            heading14       =   "Colour   "
            datatype14      =   4
            maxlength14     =   5
            style14         =   2
            colname15       =   "WHLPGRSWT"
            heading15       =   "Grs Wt        "
            datatype15      =   2
            mask15          =   "####0.000"
            maxlength15     =   9
            style15         =   2
         End
         Begin MwfCtl.MWCTL_BTN1 CmdDsgHlpGo 
            Height          =   495
            Left            =   10140
            TabIndex        =   30
            ToolTipText     =   "Click Here To Show Help"
            Top             =   480
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
         Begin MwfCtl.MWCTL_BTN1 CmdDsgHlpCpy 
            Height          =   495
            Left            =   13050
            TabIndex        =   88
            ToolTipText     =   "Option To Copy Details From Other Design"
            Top             =   480
            Width           =   855
            _ExtentX        =   1508
            _ExtentY        =   873
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
         Begin VB.Label ALBL 
            BackStyle       =   0  'Transparent
            Caption         =   "Sort"
            BeginProperty Font 
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
            Left            =   7980
            TabIndex        =   63
            Top             =   480
            Width           =   495
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
            Index           =   14
            Left            =   6210
            TabIndex        =   62
            Top             =   480
            Width           =   525
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
            Left            =   3570
            TabIndex        =   61
            Top             =   480
            Width           =   675
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
            Index           =   24
            Left            =   120
            TabIndex        =   60
            Top             =   480
            Width           =   765
         End
         Begin VB.Label LblMain 
            BackStyle       =   0  'Transparent
            BorderStyle     =   1  'Fixed Single
            Caption         =   "  Design Help"
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
            TabIndex        =   59
            Top             =   90
            Width           =   14835
         End
      End
      Begin VB.Frame FraSumm 
         BorderStyle     =   0  'None
         Height          =   3315
         Left            =   180
         TabIndex        =   64
         Top             =   5730
         Width           =   15000
         Begin TabDlg.SSTab TabSumm 
            Height          =   3225
            Left            =   0
            TabIndex        =   65
            Top             =   0
            Width           =   14985
            _ExtentX        =   26432
            _ExtentY        =   5689
            _Version        =   393216
            Tab             =   2
            TabHeight       =   520
            TabCaption(0)   =   "<&1> Design RM Summary"
            TabPicture(0)   =   "EmrFrmRetMemo.frx":0000
            Tab(0).ControlEnabled=   0   'False
            Tab(0).Control(0)=   "FraTabSumm(0)"
            Tab(0).ControlCount=   1
            TabCaption(1)   =   "<&2> Voucher Summary"
            TabPicture(1)   =   "EmrFrmRetMemo.frx":001C
            Tab(1).ControlEnabled=   0   'False
            Tab(1).Control(0)=   "FraTabSumm(1)"
            Tab(1).ControlCount=   1
            TabCaption(2)   =   "<&3> Voucher RM Summary"
            TabPicture(2)   =   "EmrFrmRetMemo.frx":0038
            Tab(2).ControlEnabled=   -1  'True
            Tab(2).Control(0)=   "FraTabSumm(2)"
            Tab(2).Control(0).Enabled=   0   'False
            Tab(2).ControlCount=   1
            Begin VB.Frame FraTabSumm 
               BorderStyle     =   0  'None
               Enabled         =   0   'False
               Height          =   2830
               Index           =   0
               Left            =   -74940
               TabIndex        =   79
               Top             =   360
               Width           =   14895
               Begin MwfCtl.MWCTL_MED ATXT 
                  Height          =   285
                  Index           =   26
                  Left            =   3090
                  TabIndex        =   80
                  ToolTipText     =   "Total Raw Material Quantity"
                  Top             =   2460
                  Width           =   1035
                  _ExtentX        =   1826
                  _ExtentY        =   503
                  Alignment       =   2
                  Enabled         =   0   'False
                  Mask            =   "######0"
                  MaxLength       =   7
                  DataType        =   2
                  IdName          =   "WDRTOTQTY"
               End
               Begin MwfCtl.MWCTL_FLX GrdSummDR 
                  Height          =   2310
                  Left            =   90
                  TabIndex        =   81
                  ToolTipText     =   "Design Raw Material Summary"
                  Top             =   150
                  Width           =   7785
                  _ExtentX        =   13732
                  _ExtentY        =   4075
                  Cols            =   7
                  colname1        =   "WDRCTG"
                  heading1        =   "Ctg"
                  datatype1       =   4
                  maxlength1      =   2
                  tooltiptext1    =   "Raw Material Category"
                  colname2        =   "WDRSCTG"
                  heading2        =   "Sub Ctg"
                  datatype2       =   4
                  maxlength2      =   5
                  tooltiptext2    =   "Raw Material Sub Category"
                  colname3        =   "WDRCD"
                  heading3        =   "Rm Code                                 "
                  datatype3       =   4
                  maxlength3      =   16
                  tooltiptext3    =   "Raw Material Code"
                  colname4        =   "WDRQTY"
                  heading4        =   "Rm Qty     "
                  datatype4       =   1
                  mask4           =   "#####0"
                  maxlength4      =   6
                  tooltiptext4    =   "Raw Material Quantity"
                  colname5        =   "WDRIWT"
                  heading5        =   "Inv Wt.           "
                  datatype5       =   2
                  mask5           =   "#####0.000"
                  maxlength5      =   10
                  tooltiptext5    =   "Invoice Raw Material Weight"
                  colname6        =   "WDRIVAL"
                  heading6        =   "Inv Val                "
                  datatype6       =   2
                  mask6           =   "########0.00"
                  maxlength6      =   12
                  tooltiptext6    =   "Invoice Raw Material Value"
               End
               Begin MwfCtl.MWCTL_MED ATXT 
                  Height          =   285
                  Index           =   34
                  Left            =   4080
                  TabIndex        =   82
                  ToolTipText     =   "Total Invoice Raw Material Weight"
                  Top             =   2460
                  Width           =   1245
                  _ExtentX        =   2196
                  _ExtentY        =   503
                  Alignment       =   2
                  Enabled         =   0   'False
                  Mask            =   "######0.000"
                  MaxLength       =   11
                  DataType        =   2
                  IdName          =   "WDRTOTIWT"
               End
               Begin MwfCtl.MWCTL_MED ATXT 
                  Height          =   285
                  Index           =   29
                  Left            =   5310
                  TabIndex        =   83
                  ToolTipText     =   "Total Invoice Raw Material Value"
                  Top             =   2460
                  Width           =   1515
                  _ExtentX        =   2672
                  _ExtentY        =   503
                  Alignment       =   2
                  Enabled         =   0   'False
                  Mask            =   "########0.000"
                  MaxLength       =   13
                  DataType        =   2
                  IdName          =   "WDRTOTIVAL"
               End
               Begin MwfCtl.MWCTL_MED ATXT 
                  Height          =   285
                  Index           =   56
                  Left            =   1095
                  TabIndex        =   84
                  ToolTipText     =   "Enter Group Option For Design Raw Material Summary"
                  Top             =   2460
                  Width           =   1365
                  _ExtentX        =   2408
                  _ExtentY        =   503
                  MaxLength       =   8
                  DataType        =   4
                  IdName          =   "WDRGRP"
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
                  Index           =   16
                  Left            =   90
                  TabIndex        =   85
                  Top             =   2460
                  Width           =   1005
               End
            End
            Begin VB.Frame FraTabSumm 
               BorderStyle     =   0  'None
               Enabled         =   0   'False
               Height          =   2830
               Index           =   1
               Left            =   -74940
               TabIndex        =   73
               Top             =   360
               Width           =   14895
               Begin MwfCtl.MWCTL_MED ATXT 
                  Height          =   285
                  Index           =   54
                  Left            =   5565
                  TabIndex        =   74
                  ToolTipText     =   "Total Order Quantity"
                  Top             =   2460
                  Width           =   1215
                  _ExtentX        =   2143
                  _ExtentY        =   503
                  Alignment       =   2
                  Enabled         =   0   'False
                  Mask            =   "#######0.0"
                  MaxLength       =   10
                  DataType        =   2
                  IdName          =   "WIDTOTQTY"
               End
               Begin MwfCtl.MWCTL_FLX GrdSummID 
                  Height          =   2310
                  Left            =   90
                  TabIndex        =   75
                  ToolTipText     =   "Order Design Summary"
                  Top             =   150
                  Width           =   9975
                  _ExtentX        =   17595
                  _ExtentY        =   4075
                  Cols            =   7
                  colname1        =   "WIDSR"
                  heading1        =   "Inv Sr.                                      "
                  datatype1       =   4
                  maxlength1      =   25
                  tooltiptext1    =   "Design Serial No"
                  colname2        =   "WIDDMCTG"
                  heading2        =   "Dsg Ctg "
                  datatype2       =   4
                  maxlength2      =   5
                  tooltiptext2    =   "Design Category"
                  colname3        =   "WIDDMCD"
                  heading3        =   "Design Cd                        "
                  datatype3       =   4
                  maxlength3      =   15
                  tooltiptext3    =   "Design Code"
                  colname4        =   "WIDQTY"
                  heading4        =   "Qty               "
                  datatype4       =   2
                  mask4           =   "######0.0"
                  maxlength4      =   9
                  tooltiptext4    =   "Order Quantity"
                  colname5        =   "WIDIPRC"
                  heading5        =   "Inv Price            "
                  datatype5       =   2
                  mask5           =   "########0.00"
                  maxlength5      =   12
                  tooltiptext5    =   "Invoice Price"
                  colname6        =   "WIDIFOB"
                  heading6        =   "Inv FOB                  "
                  datatype6       =   2
                  mask6           =   "##########0.00"
                  maxlength6      =   14
                  tooltiptext6    =   "Invoice FOB"
               End
               Begin MwfCtl.MWCTL_MED ATXT 
                  Height          =   285
                  Index           =   31
                  Left            =   7980
                  TabIndex        =   76
                  ToolTipText     =   "Total Invoice FOB"
                  Top             =   2460
                  Width           =   1725
                  _ExtentX        =   3043
                  _ExtentY        =   503
                  Alignment       =   2
                  Enabled         =   0   'False
                  Mask            =   "###########0.00"
                  MaxLength       =   15
                  DataType        =   2
                  IdName          =   "WIDTOTIFOB"
               End
               Begin MwfCtl.MWCTL_MED ATXT 
                  Height          =   285
                  Index           =   62
                  Left            =   1095
                  TabIndex        =   77
                  ToolTipText     =   "Enter Group Option For Order Design Summary"
                  Top             =   2460
                  Width           =   1365
                  _ExtentX        =   2408
                  _ExtentY        =   503
                  MaxLength       =   8
                  DataType        =   4
                  IdName          =   "WIDGRP"
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
                  Index           =   23
                  Left            =   90
                  TabIndex        =   78
                  Top             =   2460
                  Width           =   1005
               End
            End
            Begin VB.Frame FraTabSumm 
               BorderStyle     =   0  'None
               Enabled         =   0   'False
               Height          =   2830
               Index           =   2
               Left            =   60
               TabIndex        =   66
               Top             =   360
               Width           =   14895
               Begin MwfCtl.MWCTL_MED ATXT 
                  Height          =   285
                  Index           =   27
                  Left            =   2910
                  TabIndex        =   67
                  ToolTipText     =   "Total Raw Material Quantity"
                  Top             =   2460
                  Width           =   945
                  _ExtentX        =   1667
                  _ExtentY        =   503
                  Alignment       =   2
                  Enabled         =   0   'False
                  Mask            =   "#######0"
                  MaxLength       =   8
                  DataType        =   2
                  IdName          =   "WIRTOTQTY"
               End
               Begin MwfCtl.MWCTL_FLX GrdSummIR 
                  Height          =   2310
                  Left            =   90
                  TabIndex        =   68
                  Top             =   150
                  Width           =   7845
                  _ExtentX        =   13838
                  _ExtentY        =   4075
                  Cols            =   7
                  colname1        =   "WIRCTG"
                  heading1        =   "Ctg"
                  datatype1       =   4
                  maxlength1      =   2
                  tooltiptext1    =   "Raw Material Category"
                  colname2        =   "WIRSCTG"
                  heading2        =   "SubCtg"
                  datatype2       =   4
                  maxlength2      =   5
                  tooltiptext2    =   "Raw Material Sub Category"
                  colname3        =   "WIRCD"
                  heading3        =   "Rm Code                                 "
                  datatype3       =   4
                  maxlength3      =   16
                  tooltiptext3    =   "Raw Material Code"
                  colname4        =   "WIRQTY"
                  heading4        =   "Qty           "
                  datatype4       =   1
                  mask4           =   "######0"
                  maxlength4      =   7
                  tooltiptext4    =   "Raw Material Quantity"
                  colname5        =   "WIRIWT"
                  heading5        =   "Inv Wt.            "
                  datatype5       =   2
                  mask5           =   "######0.000"
                  maxlength5      =   11
                  tooltiptext5    =   "Invoice Raw Material Weight"
                  colname6        =   "WIRIVAL"
                  heading6        =   "Inv Val                  "
                  datatype6       =   2
                  mask6           =   "#########0.00"
                  maxlength6      =   13
                  tooltiptext6    =   "Invoice Raw Material Value"
               End
               Begin MwfCtl.MWCTL_MED ATXT 
                  Height          =   285
                  Index           =   28
                  Left            =   3840
                  TabIndex        =   69
                  ToolTipText     =   "Total Raw Material Invoice Weight"
                  Top             =   2460
                  Width           =   1275
                  _ExtentX        =   2249
                  _ExtentY        =   503
                  Alignment       =   2
                  Enabled         =   0   'False
                  Mask            =   "#######0.000"
                  MaxLength       =   12
                  DataType        =   2
                  IdName          =   "WIRTOTIWT"
               End
               Begin MwfCtl.MWCTL_MED ATXT 
                  Height          =   285
                  Index           =   32
                  Left            =   5100
                  TabIndex        =   70
                  ToolTipText     =   "Total Raw Material Invoice Value"
                  Top             =   2460
                  Width           =   1605
                  _ExtentX        =   2831
                  _ExtentY        =   503
                  Alignment       =   2
                  Enabled         =   0   'False
                  Mask            =   "##########0.00"
                  MaxLength       =   14
                  DataType        =   2
                  IdName          =   "WIRTOTIVAL"
               End
               Begin MwfCtl.MWCTL_MED ATXT 
                  Height          =   285
                  Index           =   63
                  Left            =   1095
                  TabIndex        =   71
                  ToolTipText     =   "Enter Group Option For Order Raw Material Summary"
                  Top             =   2460
                  Width           =   1365
                  _ExtentX        =   2408
                  _ExtentY        =   503
                  MaxLength       =   8
                  DataType        =   4
                  IdName          =   "WIRGRP"
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
                  Left            =   90
                  TabIndex        =   72
                  Top             =   2460
                  Width           =   1005
               End
            End
         End
      End
      Begin VB.Frame FraCpy 
         Height          =   1935
         Left            =   4455
         TabIndex        =   39
         Top             =   7095
         Width           =   5310
         Begin MwfCtl.MWCTL_BTN1 CmdCpyGo 
            Height          =   495
            Left            =   4395
            TabIndex        =   25
            ToolTipText     =   "Click Here To Copy"
            Top             =   1140
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
            Index           =   4
            Left            =   2775
            TabIndex        =   15
            ToolTipText     =   "Enter From Invoice Voucher Character"
            Top             =   480
            Width           =   570
            _ExtentX        =   1005
            _ExtentY        =   503
            MaxLength       =   3
            DataType        =   4
            ReCalcParent    =   "WCPINVCHRTO"
            IdName          =   "WCPINVCHRFR"
            CmpStr          =   "InChr >= "
         End
         Begin MwfCtl.MWCTL_MED ATXT 
            Height          =   285
            Index           =   5
            Left            =   2340
            TabIndex        =   14
            ToolTipText     =   "Enter From Invoice Voucher Year"
            Top             =   480
            Width           =   330
            _ExtentX        =   582
            _ExtentY        =   503
            MaxLength       =   2
            DataType        =   4
            ReCalcParent    =   "WCPINVYYTO"
            IdName          =   "WCPINVYYFR"
            CmpStr          =   "InYy >= "
         End
         Begin MwfCtl.MWCTL_MED ATXT 
            Height          =   285
            Index           =   6
            Left            =   3450
            TabIndex        =   16
            ToolTipText     =   "Enter From Invoice Voucher Number"
            Top             =   480
            Width           =   855
            _ExtentX        =   1508
            _ExtentY        =   503
            Alignment       =   2
            Mask            =   "######0"
            MaxLength       =   7
            DataType        =   1
            ReCalcParent    =   "WCPINVNOTO"
            IdName          =   "WCPINVNOFR"
            CmpStr          =   "InNo >= "
         End
         Begin MwfCtl.MWCTL_MED ATXT 
            Height          =   285
            Index           =   7
            Left            =   1650
            TabIndex        =   13
            ToolTipText     =   "Enter From Invoice Voucher Tc"
            Top             =   480
            Width           =   570
            _ExtentX        =   1005
            _ExtentY        =   503
            Enabled         =   0   'False
            MaxLength       =   3
            DataType        =   4
            IdName          =   "WCPINVTCFR"
            CmpStr          =   "InTc >= "
         End
         Begin MwfCtl.MWCTL_MED ATXT 
            Height          =   285
            Index           =   8
            Left            =   4395
            TabIndex        =   17
            ToolTipText     =   "Enter From Invoice Voucher Sr"
            Top             =   480
            Width           =   555
            _ExtentX        =   979
            _ExtentY        =   503
            Alignment       =   2
            Mask            =   "###0"
            MaxLength       =   4
            DataType        =   1
            ReCalcParent    =   "WCPINVSRTO"
            IdName          =   "WCPINVSRFR"
            CmpStr          =   "IdSr >= "
         End
         Begin MwfCtl.MWCTL_MED ATXT 
            Height          =   285
            Index           =   9
            Left            =   2775
            TabIndex        =   20
            ToolTipText     =   "Enter To Invoice Voucher Character"
            Top             =   765
            Width           =   570
            _ExtentX        =   1005
            _ExtentY        =   503
            MaxLength       =   3
            DataType        =   4
            ReCalcOn        =   "WCPINVCHRFR"
            IdName          =   "WCPINVCHRTO"
            CmpStr          =   "InChr <= "
         End
         Begin MwfCtl.MWCTL_MED ATXT 
            Height          =   285
            Index           =   10
            Left            =   2340
            TabIndex        =   19
            ToolTipText     =   "Enter To Invoice Voucher Year"
            Top             =   765
            Width           =   330
            _ExtentX        =   582
            _ExtentY        =   503
            MaxLength       =   2
            DataType        =   4
            ReCalcOn        =   "WCPINVYYFR"
            IdName          =   "WCPINVYYTO"
            CmpStr          =   "InYy <= "
         End
         Begin MwfCtl.MWCTL_MED ATXT 
            Height          =   285
            Index           =   11
            Left            =   3450
            TabIndex        =   21
            ToolTipText     =   "Enter To Invoice Voucher Number"
            Top             =   765
            Width           =   855
            _ExtentX        =   1508
            _ExtentY        =   503
            Alignment       =   2
            Mask            =   "######0"
            MaxLength       =   7
            DataType        =   1
            ReCalcOn        =   "WCPINVNOFR"
            IdName          =   "WCPINVNOTO"
            CmpStr          =   "InNo <= "
         End
         Begin MwfCtl.MWCTL_MED ATXT 
            Height          =   285
            Index           =   12
            Left            =   1650
            TabIndex        =   18
            ToolTipText     =   "Enter To Invoice Voucher Tc"
            Top             =   765
            Width           =   570
            _ExtentX        =   1005
            _ExtentY        =   503
            Enabled         =   0   'False
            MaxLength       =   3
            DataType        =   4
            IdName          =   "WCPINVTCTO"
            CmpStr          =   "InTc <= "
         End
         Begin MwfCtl.MWCTL_MED ATXT 
            Height          =   285
            Index           =   13
            Left            =   4395
            TabIndex        =   22
            ToolTipText     =   "Enter To Invoice Voucher Sr"
            Top             =   765
            Width           =   555
            _ExtentX        =   979
            _ExtentY        =   503
            Alignment       =   2
            Mask            =   "###0"
            MaxLength       =   4
            DataType        =   1
            ReCalcOn        =   "WCPINVSRFR"
            IdName          =   "WCPINVSRTO"
            CmpStr          =   "IdSr <= "
         End
         Begin MwfCtl.MWCTL_MED ATXT 
            Height          =   285
            Index           =   14
            Left            =   1650
            TabIndex        =   23
            ToolTipText     =   "Enter From Invoice Date"
            Top             =   1050
            Width           =   885
            _ExtentX        =   1561
            _ExtentY        =   503
            MaxLength       =   8
            DataType        =   3
            IdName          =   "WCPINVDTFR"
            CmpStr          =   "InDt >="
         End
         Begin MwfCtl.MWCTL_MED ATXT 
            Height          =   285
            Index           =   15
            Left            =   1650
            TabIndex        =   24
            ToolTipText     =   "Enter To Invoice Date"
            Top             =   1335
            Width           =   885
            _ExtentX        =   1561
            _ExtentY        =   503
            MaxLength       =   8
            DataType        =   3
            IdName          =   "WCPINVDTTO"
            CmpStr          =   "InDt <="
         End
         Begin VB.Label ALBL 
            BackStyle       =   0  'Transparent
            Caption         =   "Inv Date To"
            BeginProperty Font 
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
            TabIndex        =   56
            Top             =   1335
            Width           =   1515
         End
         Begin VB.Label ALBL 
            BackStyle       =   0  'Transparent
            Caption         =   "Inv Date Fr"
            BeginProperty Font 
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
            TabIndex        =   55
            Top             =   1050
            Width           =   1515
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
            Left            =   4320
            TabIndex        =   54
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
            Index           =   8
            Left            =   2250
            TabIndex        =   53
            Top             =   765
            Width           =   105
         End
         Begin VB.Label Label2 
            BackStyle       =   0  'Transparent
            Caption         =   "Inv Voucher To"
            BeginProperty Font 
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
            TabIndex        =   52
            Top             =   765
            Width           =   1515
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
            Left            =   3375
            TabIndex        =   51
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
            Index           =   6
            Left            =   2685
            TabIndex        =   50
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
            Index           =   5
            Left            =   4320
            TabIndex        =   49
            Top             =   480
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
            Index           =   4
            Left            =   2250
            TabIndex        =   48
            Top             =   480
            Width           =   105
         End
         Begin VB.Label Label1 
            BackStyle       =   0  'Transparent
            Caption         =   "Inv Voucher Fr"
            BeginProperty Font 
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
            TabIndex        =   47
            Top             =   480
            Width           =   1515
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
            Left            =   3375
            TabIndex        =   46
            Top             =   480
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
            Left            =   2685
            TabIndex        =   45
            Top             =   480
            Width           =   105
         End
         Begin VB.Label LblMain 
            BackStyle       =   0  'Transparent
            BorderStyle     =   1  'Fixed Single
            Caption         =   "  Copy From Invoice"
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
            TabIndex        =   40
            Top             =   90
            Width           =   5325
         End
      End
      Begin VB.Frame FraNKeyAll 
         BorderStyle     =   0  'None
         Height          =   9135
         Left            =   120
         TabIndex        =   37
         Top             =   0
         Width           =   15090
         Begin VB.Frame FraPicHsk 
            BorderStyle     =   0  'None
            Height          =   1635
            Left            =   11700
            TabIndex        =   87
            Top             =   0
            Width           =   1575
            Begin VB.Image PicHsk 
               Height          =   1515
               Left            =   0
               ToolTipText     =   "Hand Sketch Image"
               Top             =   30
               Width           =   1515
            End
         End
         Begin VB.Frame FraPic3D 
            BorderStyle     =   0  'None
            Height          =   1635
            Left            =   13425
            TabIndex        =   86
            Top             =   0
            Width           =   1575
            Begin VB.Image Pic3D 
               Height          =   1515
               Left            =   0
               ToolTipText     =   "3D - Image"
               Top             =   30
               Width           =   1515
            End
         End
         Begin MwfCtl.MWCTL_FLX GrdRetMemDet 
            Height          =   7320
            Left            =   0
            TabIndex        =   9
            Top             =   1680
            Width           =   15075
            _ExtentX        =   26591
            _ExtentY        =   12912
            Cols            =   18
            colname1        =   "REDSR"
            heading1        =   "Sr.       "
            datafld1        =   "RedSr"
            datatype1       =   1
            mask1           =   "####0"
            maxlength1      =   5
            style1          =   2
            colname2        =   "REDIDTC"
            heading2        =   "Tc    "
            datafld2        =   "RedIdTc"
            datatype2       =   4
            recalcparent2   =   $"EmrFrmRetMemo.frx":0054
            maxlength2      =   3
            style2          =   2
            tooltiptext2    =   "Enter Invoice Tc"
            colname3        =   "REDIDYY"
            heading3        =   "Yy  "
            datafld3        =   "RedIdYy"
            datatype3       =   4
            recalcparent3   =   $"EmrFrmRetMemo.frx":00E1
            maxlength3      =   2
            tooltiptext3    =   "Enter Invoice Year"
            colname4        =   "REDIDCHR"
            heading4        =   "Chr    "
            datafld4        =   "RedIdChr"
            datatype4       =   4
            recalcparent4   =   $"EmrFrmRetMemo.frx":016E
            maxlength4      =   3
            tooltiptext4    =   "Enter Invoice Character"
            colname5        =   "REDIDNO"
            heading5        =   "No            "
            datafld5        =   "RedIdNo"
            datatype5       =   1
            mask5           =   "######0"
            recalcon5       =   "GrdRetMemDet(REDIDTC),GrdRetMemDet(REDIDYY),GrdRetMemDet(REDIDCHR)"
            recalcparent5   =   $"EmrFrmRetMemo.frx":01FB
            maxlength5      =   7
            tooltiptext5    =   "Enter Invoice Number"
            colname6        =   "REDIDSR"
            heading6        =   "Sr      "
            datafld6        =   "RedIdSr"
            datatype6       =   1
            mask6           =   "###0"
            recalcon6       =   "GrdRetMemDet(REDIDTC),GrdRetMemDet(REDIDYY),GrdRetMemDet(REDIDCHR),GrdRetMemDet(REDIDNO)"
            recalcparent6   =   "GrdRetMemDet(REDQTY),GrdRetMemDet(REDGRSWT),GrdRetMemDet(REDFOBPERPC),GrdRetMemDet(WIDDMCD)"
            maxlength6      =   4
            tooltiptext6    =   "Enter Invoice Sr."
            colname7        =   "REDQTY"
            heading7        =   "Qty           "
            datafld7        =   "RedQty"
            datatype7       =   2
            mask7           =   "####0.0"
            recalcon7       =   "GrdRetMemDet(REDIDTC),GrdRetMemDet(REDIDYY),GrdRetMemDet(REDIDCHR),GrdRetMemDet(REDIDNO),GrdRetMemDet(REDIDSR)"
            recalcparent7   =   "GrdRetMemDet(REDGRSWT),GrdRetMemDet(WREDFOB)"
            maxlength7      =   7
            tooltiptext7    =   "Enter Invoice Design Return Quantity"
            colname8        =   "REDGRSWT"
            heading8        =   "Grs Wt          "
            datafld8        =   "RedGrsWt"
            datatype8       =   2
            mask8           =   "#####0.000"
            recalcon8       =   $"EmrFrmRetMemo.frx":0286
            recalcparent8   =   "GrdRetMemDet(REDGRSWTDIFF)"
            maxlength8      =   10
            tooltiptext8    =   "Enter Gross Weight Of Returned Pieces"
            colname9        =   "REDRETRSN"
            heading9        =   "Reason Cd"
            datafld9        =   "RedRetRsn"
            maxlength9      =   8
            tooltiptext9    =   "Enter Return Reason Code"
            colname10       =   "REDFOBPERPC"
            heading10       =   "FOB Per Pc  "
            datafld10       =   "RedFOBPerPc"
            datatype10      =   2
            mask10          =   "######0.00"
            recalcon10      =   "GrdRetMemDet(REDIDTC),GrdRetMemDet(REDIDYY),GrdRetMemDet(REDIDCHR),GrdRetMemDet(REDIDNO),GrdRetMemDet(REDIDSR)"
            recalcparent10  =   "GrdRetMemDet(WREDFOB)"
            maxlength10     =   10
            style10         =   2
            tooltiptext10   =   "Invoice FOB Value Per Piece"
            colname11       =   "WREDFOB"
            heading11       =   "FOB                "
            datatype11      =   2
            mask11          =   "######0.00"
            recalcon11      =   "GrdRetMemDet(REDFOBPERPC),GrdRetMemDet(REDQTY)"
            maxlength11     =   10
            style11         =   2
            tooltiptext11   =   "Invoice FOB Value"
            colname12       =   "WIDDMCD"
            heading12       =   "Design                            "
            datatype12      =   4
            recalcon12      =   "GrdRetMemDet(REDIDSR)"
            maxlength12     =   15
            style12         =   2
            tooltiptext12   =   "Design Code"
            colname13       =   "WIDSFX"
            heading13       =   "Suffix                 "
            datatype13      =   4
            maxlength13     =   15
            style13         =   2
            tooltiptext13   =   "Design Suffix"
            colname14       =   "WIDDMSZ"
            heading14       =   "Size      "
            datatype14      =   4
            maxlength14     =   5
            style14         =   2
            tooltiptext14   =   "Design Size"
            colname15       =   "WINDT"
            heading15       =   "Inv Date "
            datatype15      =   3
            recalcon15      =   "GrdRetMemDet(REDIDNO)"
            maxlength15     =   8
            style15         =   2
            tooltiptext15   =   "Invoice Date"
            colname16       =   "WIDEXPCMCD"
            heading16       =   "Order Cust      "
            datatype16      =   4
            maxlength16     =   8
            style16         =   2
            tooltiptext16   =   "Export Order Customer"
            colname17       =   "REDGRSWTDIFF"
            heading17       =   "Diff In Grs Wt"
            datafld17       =   "RedGrsWtDiff"
            datatype17      =   2
            mask17          =   "-####0.000"
            recalcon17      =   "GrdRetMemDet(REDGRSWT)"
            maxlength17     =   10
            style17         =   0
            tooltiptext17   =   "diff in grs wt as edited by user"
         End
         Begin MwfCtl.MWCTL_MED ATXT 
            Height          =   285
            Index           =   2
            Left            =   5250
            TabIndex        =   3
            ToolTipText     =   "Enter Return Memo Date"
            Top             =   15
            Width           =   885
            _ExtentX        =   1561
            _ExtentY        =   503
            MaxLength       =   8
            DataType        =   3
            DataField       =   "ReDt"
            IdName          =   "REDT"
         End
         Begin MwfCtl.MWCTL_MED ATXT 
            Height          =   570
            Index           =   3
            Left            =   2280
            TabIndex        =   5
            ToolTipText     =   "Enter Description"
            Top             =   300
            Width           =   7095
            _ExtentX        =   12515
            _ExtentY        =   1005
            Alignment       =   3
            MaxLength       =   45
            DataField       =   "ReDesc"
            IdName          =   "REDESC"
         End
         Begin MwfCtl.MWCTL_MED ATXT 
            Height          =   285
            Index           =   16
            Left            =   8040
            TabIndex        =   4
            ToolTipText     =   "Enter Return Customer Code"
            Top             =   15
            Width           =   1335
            _ExtentX        =   2355
            _ExtentY        =   503
            MaxLength       =   8
            DataType        =   4
            DataField       =   "ReCmCd"
            IdName          =   "RECMCD"
         End
         Begin MwfCtl.MWCTL_MED ATXT 
            Height          =   285
            Index           =   20
            Left            =   2280
            TabIndex        =   6
            ToolTipText     =   "Enter Default Return Reason  Code"
            Top             =   870
            Width           =   1335
            _ExtentX        =   2355
            _ExtentY        =   503
            MaxLength       =   8
            DataType        =   4
            IdName          =   "WRETRSN"
         End
         Begin MwfCtl.MWCTL_MED ATXT 
            Height          =   285
            Index           =   21
            Left            =   2280
            TabIndex        =   7
            ToolTipText     =   "Enter Default Return Category"
            Top             =   1320
            Width           =   975
            _ExtentX        =   1720
            _ExtentY        =   503
            MaxLength       =   5
            DataType        =   4
            DataField       =   "ReRtnCtg"
            ReCalcParent    =   "WRERTNCTGDESC"
            IdName          =   "RERTNCTG"
         End
         Begin MwfCtl.MWCTL_MED ATXT 
            Height          =   285
            Index           =   22
            Left            =   3360
            TabIndex        =   8
            ToolTipText     =   "Enter Default Return Category Description"
            Top             =   1320
            Width           =   7455
            _ExtentX        =   13150
            _ExtentY        =   503
            MaxLength       =   45
            ReCalcOn        =   "RERTNCTG"
            IdName          =   "WRERTNCTGDESC"
         End
         Begin VB.Label ALBL 
            BackStyle       =   0  'Transparent
            Caption         =   "-"
            BeginProperty Font 
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
            Left            =   3240
            TabIndex        =   91
            Top             =   1320
            Width           =   105
         End
         Begin VB.Label ALBL 
            BackStyle       =   0  'Transparent
            Caption         =   "Return Catogory"
            BeginProperty Font 
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
            Left            =   60
            TabIndex        =   90
            Top             =   1320
            Width           =   2115
         End
         Begin VB.Label ALBL 
            BackStyle       =   0  'Transparent
            Caption         =   "Default Return Reason Code"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   525
            Index           =   17
            Left            =   60
            TabIndex        =   89
            Top             =   800
            Width           =   2055
         End
         Begin VB.Label ALBL 
            BackStyle       =   0  'Transparent
            Caption         =   "Inv Customer"
            BeginProperty Font 
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
            Left            =   6750
            TabIndex        =   57
            Top             =   15
            Width           =   1275
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
            Index           =   0
            Left            =   4740
            TabIndex        =   44
            Top             =   30
            Width           =   585
         End
         Begin VB.Label ALBL 
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
            Index           =   3
            Left            =   60
            TabIndex        =   38
            Top             =   300
            Width           =   855
         End
      End
   End
End
Attribute VB_Name = "EmrFrmRetMemo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'---- General Comments on the Entry --------'
' Image controls  PicHsk for the pencil Sketch, Pic3d for the 3-D image.
'   These image controls are placed on the frames FraPicHsk & FraPic3D respectively
' Enlarged state of the pictures can be obtained by double clicking on the pictures
' The pictures can be brought back to the original size by again double clicking the pictures
'--------------------

Option Explicit
Dim moCn As MwfLib.MDOConnection   '*** (Bef speed) bef 24/11/06

Dim ms_ReYy As String, ms_ReChr As String
Const ctDsgSumm = "D": Const ctInvSumm = "I"

Dim ms_InChr As String  ' **** Manali 3.03 - 24/06/08 - Def Chr for inv

'*** For the Toggle Options of Command Buttons
Private Enum en_RetMemoFra
  Cpy = 0
  Hlp = 1
  Summ = 2
End Enum

'*** (Jen 2.14 Next)
'*** Declare variables wPic3DHt, wPic3DWdt, wPicHskHt, wPicHskWdt as single to store the
    'original Height and Width of the 3D and Sketch Picture controls
'*** Declare variables wFraPic3DLft, wFraPic3DTop, wFraPicHskLft, wFraPicHskTop as single
    'to store the original Positions of the 3D and Sketch Picture controls
Dim wPic3DHt As Double
Dim wPic3DWdt As Double
Dim wPicHskHt As Double
Dim wPicHskWdt As Double

Dim wFraPic3DHt As Single
Dim wFraPic3DWdt As Single
Dim wFraPicHskHt As Single
Dim wFraPicHskWdt As Single

Dim wFraPic3DLft As Double
Dim wFraPic3DTop As Double
Dim wFraPicHskLft As Double
Dim wFraPicHskTop As Double
Dim wPic3DDblClk As Boolean
Dim wPicHskDblClk As Boolean
'*** (Jen 2.14 Next)

Private Sub ADC_Load()
  '*** Set the First Key Control property to Voucher Year
  '*** Set the First Non Key Control property to Date
  '*** Set the Child Property Of ADC to GrdRetMemDet
  '*** Set the Previous and Next Control Properties for the Grid GrdRetMemDet
  '*** Set the Hot Keys for the Grid GrdRetMemDet as 'A'

  Set moCn = ADC.Connection    '*** (Bef speed) bef 24/11/06
  
  'Set ADC.FirKeyCtl = ADC("ReYy")  '*** (Bef 2.14 Next)
  Set ADC.FirKeyCtl = ADC("ReChr")  '*** (Bef 2.14 Next)
  Set ADC.FirNKeyCtl = ADC("ReDt")
  ADC.Child = "GrdRetMemDet"
  Set GrdRetMemDet.PrevCtl = ADC("ReDesc")
  Set GrdRetMemDet.NextCtl = ADC
  Set GrdDsgHlp.NextCtl = CmdDsgHlp
  
  '*** (Jen 2.14 Next)
  Set GrdSummDR.NextCtl = ADC("wDRGrp")
  Set GrdSummID.NextCtl = ADC("wIDGrp")
  Set GrdSummIR.NextCtl = ADC("wIRGrp")
  '*** (Jen 2.14 Next)

  GrdRetMemDet.HotKey = "A"

  ms_ReYy = moCn.GetFldVal("Select hYy from Head where HCoCd= '" + gs_CoCd + "' and HCd= '" + ctSelfCmCd + "'")
  
  ' ***** Manali 3.03 - 24/06/08 - Def chr
  ms_InChr = moCn.GetFldVal("Select vPValue from vParam where vPCoCd='" + gs_CoCd + "' and vPTyp='DEFCHR' and vPMCd='IN'")
  ' ***** Manali 3.03 - 24/06/08 - Def Chr

  '*** (Jen 2.14 Next)
  wPic3DHt = Pic3D.Height
  wPic3DWdt = Pic3D.Width
  wPicHskHt = PicHsk.Height
  wPicHskWdt = PicHsk.Width
  
  wFraPic3DHt = FraPic3D.Height
  wFraPic3DWdt = FraPic3D.Width
  wFraPicHskHt = FraPicHsk.Height
  wFraPicHskWdt = FraPicHsk.Width
  
  wFraPic3DLft = FraPic3D.Left
  wFraPic3DTop = FraPic3D.Top
  wFraPicHskLft = FraPicHsk.Left
  wFraPicHskTop = FraPicHsk.Top
  
  wPic3DDblClk = False
  wPicHskDblClk = False
  '*** (Jen 2.14 Next)

  
End Sub
Private Sub ADC_KeyWhen()
  '*** Disable All command buttons
  '*** Hide all frames except Frame FraNKeyAll

  ADC("ReCoCd") = gs_CoCd
  If ms_ReYy <> "" Then ADC("ReYy") = ms_ReYy
  If ms_ReChr <> "" Then ADC("ReChr") = ms_ReChr

  '****** Sachin 2.14.0 - [24-11-2006] - [Multi-Company] ******
  Me.Caption = GetFrmCaption(gs_CoCd, "Return Memo Entry")
  '****** Sachin 2.14.0 - [24-11-2006] - [Multi-Company] ******

  Call EnaDisaCmds(True)
  Call HideAllFras
  
  '*** (Jen 2.14 Next)
  Call ShowPic(0, 0, "", PicHsk)
  Call ShowPic(0, 0, "", Pic3D)
  FraPic3D.Height = wFraPic3DHt
  FraPic3D.Width = wFraPic3DWdt
  FraPicHsk.Height = wFraPicHskHt
  FraPicHsk.Width = wFraPicHskWdt
  FraPic3D.Left = wFraPic3DLft
  FraPic3D.Top = wFraPic3DTop
  FraPicHsk.Left = wFraPicHskLft
  FraPicHsk.Top = wFraPicHskTop
  '*** (Jen 2.14 Next)
  
End Sub
Private Sub ADC_KeyValid(Cancel As Boolean, ErrMsg As String)
  '*** Force the user to enter proper Yy, Chr values
  
  Call ATXT_FldChange(99, "ReYy", Cancel, ErrMsg, ADC("ReYy"))
  If Cancel = True Then Exit Sub
  Call ATXT_FldChange(99, "ReChr", Cancel, ErrMsg, ADC("ReChr"))
  If Cancel = True Then Exit Sub
End Sub
Private Sub adc_AutoGen()
  '*** write code for AutoGenerating the Return Memo Number
  '    If the ReNo is Zero then generate the number
  '    If the user himself enters the number then write this number in the autogen table (if it is greater than the number in the autogen table)
  
  If ADC("ReNo") = 0 Then
    ADC("ReNo") = moCn.AutoGen("RetMem", "ReNo", ADC("ReCoCd") + ADC("ReYy") + ADC("ReChr"))
  Else
    Call moCn.AutoGen("RetMem", "ReNo", ADC("ReCoCd") + ADC("ReYy") + ADC("ReChr"), ADC("ReNo"))
  End If
End Sub

Private Sub ADC_NonKeyValid(pr_Cancel As Boolean, pr_ErrMsg As String)
  GrdDsgHlp.Rows = 1    ' ***** Manali 3.03 - 24/06/08 - emtying grid after save
End Sub

Private Sub ADC_NonKeyWhen(Cancel As Boolean, ErrMsg As String)
  '*** Enable all the command buttons
  '*** Set the default values for New Record
  
  ms_ReYy = ADC("ReYy")
  ms_ReChr = ADC("ReChr")
  
  ' ***** Manali 3.03 - 24/06/08 - Default values
  
  ' ***** Manali 3.4.1 - 13/05/09 - Default values Set On Button Click
  ''  adc("wCpInvYyFr") = ms_ReYy
  ''  adc("wCpInvChrFr") = ms_InChr
  ' ***** Manali 3.4.1 - 13/05/09 - Default values Set On Button Click
  
  ADC("wCpInvDtFr") = MWLib.EmptyDate
  ADC("wCpInvDtTo") = MWLib.EmptyDate
  
  ADC("wHlpSrt") = UCase(moCn.GetFldVal("Select PMCd from Param where PTyp='SRTRTHLP' and PMCd='INVDT'"))
  ' ***** Manali 3.03 - 24/06/08 - Default values
  ADC("wRetRsn") = "-"     '********Bhavna default return reason is -
  
  If ADC.AddFndMode = xaddmode Then
    ADC("ReDt") = moCn.SrvrDate
    'IG.97 -Default value as - for new voucher
    ADC("RERTNCTG") = "-"
  End If
  
  Call EnaDisaCmds(False)
End Sub

Private Sub ATXT_FldWhen(Index As Integer, ByVal IdName As String, _
                          Cancel As Boolean, ErrMsg As String)
  '*** Set Helps For Each Field
  
  Select Case UCase(IdName)
  Case Is = UCase("ReYy"), UCase("wCpInvYyFr"), UCase("wCpInvYyTo")
    Call HlpList.PMCd("YY")
  Case Is = UCase("ReChr")
    Call HlpList.PMCd("RETCHR")
  Case Is = UCase("ReNo")
    Call HlpList.ReNo(ADC("ReCoCd"), ADC("ReYy"), ADC("ReChr"))
  'Case Is = UCase("ReCurCd")
  '  Call HlpList.PMCd("CURNCY")
  Case Is = UCase("ReCmCd")
    If GrdRetMemDet.Rows - 1 > 0 Then Cancel = True: ErrMsg = "Cannot Edit As Detail Record Exists": Exit Sub
    Call HlpList.CustCd("C")
  Case Is = UCase("wCpInvChrFr")
    Call HlpList.vPSCd(ADC("ReCoCd"), "CHR", ADC("wCpInvTcFr"), "Y")
  Case Is = UCase("wCpInvChrTo")
    Call HlpList.vPSCd(ADC("ReCoCd"), "CHR", ADC("wCpInvTcTo"), "Y")
  Case Is = UCase("wCpInvNoFr")
    Call HlpList.InNo(ADC("ReCoCd"), ADC("wCpInvTcFr"), ADC("wCpInvYyFr"), ADC("wCpInvChrFr"), ADC("ReCmCd"))
  Case Is = UCase("wCpInvNoTo")
    Call HlpList.InNo(ADC("ReCoCd"), ADC("wCpInvTcTo"), ADC("wCpInvYyTo"), ADC("wCpInvChrTo"), ADC("ReCmCd"))
  Case Is = UCase("wCpInvSrFr")
    Call HlpList.InSr(ADC("ReCoCd"), ADC("wCpInvTcFr"), ADC("wCpInvYyFr"), ADC("wCpInvChrFr"), ADC("wCpInvNoFr"))
  Case Is = UCase("wCpInvSrTo")
    Call HlpList.InSr(ADC("ReCoCd"), ADC("wCpInvTcTo"), ADC("wCpInvYyTo"), ADC("wCpInvChrTo"), ADC("wCpInvNoTo"))
  Case Is = UCase("wHlpDmCd")
    If ADC("ReCmCd") = "" Then Cancel = True: ErrMsg = "Enter Customer Code First": Exit Sub
    Call HlpList.DmCd("DM")
  Case Is = UCase("wHlpSfx")
    If ADC("ReCmCd") = "" Then Cancel = True: ErrMsg = "Enter Customer Code First": Exit Sub
    ' ***** Manali Trading Module - CmCtg passed
    Call HlpList.OdSfx(ctSelfCoCd, "C", ADC("wHlpDmCd"))
  Case Is = UCase("wHlpDmSz")
    If ADC("ReCmCd") = "" Then Cancel = True: ErrMsg = "Enter Customer Code First": Exit Sub
    Call HlpList.PMCd("DMSZ")
  Case Is = UCase("wHlpSrt")
    Call HlpList.PMCd("SRTRTHLP")
  
  '*** (Jen 2.14 Next)
  Case Is = UCase("wDRGrp"), UCase("wIRGrp")
   Call HlpList.PMCd("ORDRMGR")
  Case Is = UCase("wIDGrp")
    Call HlpList.PMCd("ORDDSGGR")
  '*** (Jen 2.14 Next)
  
  '*************Bhavna
  Case Is = UCase("wRetRsn")
    Call HlpList.PMCd("RETRSN")
  'IG.97 -help for return category
  Case Is = UCase("ReRtnCtg")
    Call HlpList.PMCd("RTNCTG")
  
  End Select
End Sub

Private Sub atxt_ReCalculate(Index As Integer, ByVal IdName As String, ByVal OldVal As Variant, ByVal RowNum As Integer, ByVal SourceName As String, ByVal SourceColName As String)
  '*** write code for Recalc here
  Select Case UCase(IdName)
  ' ***** Manali 3.03 - 24/06/08 - Recalc for To fields in Cpy frame
    Case Is = UCase("wCpInvYyTo")
      If ADC.Mode = xNorm Then ADC("wCpInvYyTo") = ADC("wCpInvYyFr")
    Case Is = UCase("wCpInvChrTo")
      If ADC.Mode = xNorm Then ADC("wCpInvChrTo") = ADC("wCpInvChrFr")
    Case Is = UCase("wCpInvNoTo")
      If ADC.Mode = xNorm Then ADC("wCpInvNoTo") = ADC("wCpInvNoFr")
    Case Is = UCase("wCpInvSrTo")
      If ADC.Mode = xNorm Then ADC("wCpInvSrTo") = ADC("wCpInvSrFr")
  ' ***** Manali 3.03 - 24/06/08 - Recalc for To fields in Cpy frame
    'IG.97- shows description of the return category code selected
    Case Is = UCase("WRERTNCTGDESC")
      If ADC.Mode = xNorm Or ADC.Mode = xInit Then _
        ADC("WRERTNCTGDESC") = moCn.GetFldVal("Select PDesc225 from Param where PTyp = 'RTNCTG' and PMCd= '" + ADC("RERTNCTG") + "' and PSCd= ''")
      
  End Select
  
End Sub
Private Sub ATXT_FldChange(Index As Integer, ByVal IdName As String, Cancel As Boolean, ErrMsg As String, ByVal pv_NewValue As Variant)
  '*** ReDt         Not Blank
  
  Select Case UCase(IdName)
  Case Is = UCase("ReYy")
    Cancel = (pv_NewValue <> "") And (Not moCn.RecSeek("select PMCd from Param where PTyp= 'YY' and PMCd = '" + pv_NewValue + "'"))
    If Cancel = True Then ErrMsg = "Invalid Return Memo Voucher Year": Exit Sub
  Case Is = UCase("ReChr")
    'Cancel = (pv_NewValue <> "") And (Not moCn.RecSeek("select PMCd from Param where PTyp= 'RETCHR' and PMCd = '" + pv_NewValue + "'"))
    'If Cancel = True Then ErrMsg = "Invalid Return Memo Voucher Character": Exit Sub
    
    Dim ws_UsrArr() As String, wi_Index As Integer
    'ws_UsrArr = Split(moCn.GetFldVal("Select vPValue4 from vParam where vPCoCd= '" + adc("TCoCd") + "' and " + _
    '                                     "vPTyp='CHR' and vPMCd='" + adc("TTc") + "' and vPSCd= '" + pv_NewValue + "'"), ",")
                                         
    '4.1.4
    ws_UsrArr = Split(moCn.GetFldVal("Select PDesc225 from Param where PTyp= 'RETCHR' and PMCd = '" + pv_NewValue + "'"), ",")
    If IsArray(ws_UsrArr) Then
        For wi_Index = 0 To UBound(ws_UsrArr)
          Cancel = IIF(UCase(gs_UsrCd) = UCase(ws_UsrArr(wi_Index)), False, True)
          If Cancel = False Then: Exit For
        Next wi_Index
    End If
    If Cancel = True Then ErrMsg = "Access To This Voucher Character Is Denied To User " + gs_UsrCd: Exit Sub
    
  Case Is = UCase("ReDt")
    Cancel = (pv_NewValue = "")
    If Cancel = True Then ErrMsg = "Memo Date Cannot Be Blank": Exit Sub
  Case Is = UCase("ReCmCd")
    Cancel = Not moCn.RecSeek("select CmCd from CustMst where CmCtg= 'C' and CmCd = '" + pv_NewValue + "'")
    If Cancel = True Then ErrMsg = "Invalid Customer Code": Exit Sub
  'Case Is = UCase("ReCurCd")
  '  Cancel = Not moCn.RecSeek("select PMCd from Param where PTyp= 'CURNCY' and PMCd = '" + pv_NewValue + "'")
  '  If Cancel = True Then ErrMsg = "Invalid Currency Code": Exit Sub
  
  Case Is = UCase("wHlpDmCd")
' ******** Manali 3.2.2 Design Exclusivity check removed from Design help
'    Dim ws_CmRegnCd As String
'    ws_CmRegnCd = moCn.GetFldVal("Select CmRegnCd from CustMst where CmCtg= 'C' and CmCd= '" + ADC("ReCmCd") + "'")
      
    Cancel = (pv_NewValue <> "") And (Not moCn.RecSeek("Select DmCd From DsgMst Where DmTcTyp= 'DM' " + _
                                      "And DmCd = '" + pv_NewValue + "' and DmSz= '' and  DmPrtKey='" + ctCurrPrtn + "' "))
    If Cancel = True Then ErrMsg = "Invalid Dsg Cd": Exit Sub
     '                                 "  and " + _
                                      "(DmCmCd= '' or DmCmCd= '" + ADC("ReCmCd") + "' or " + _
                                      "(DmCmCd<> '" + ADC("ReCmCd") + "' and DmRegnExc= 'Y' and " + _
                                      "(Select CmRegnCd from CustMst where CmCtg= 'C' and CmCd= DmCmCd) <> '" + ws_CmRegnCd + "')) "))
    'If Cancel = True Then ErrMsg = "Invalid Dsg Cd/ Dsg Is Exclusive For Another Cust Or Regn": Exit Sub
  
  'Case Is = UCase("wHlpSfx")
  Case Is = UCase("wHlpDmSz")
    Cancel = (pv_NewValue <> "") And (Not moCn.RecSeek("Select PMCd From Param Where " + _
                                  "PTyp= 'DMSZ' And PMCd = '" + pv_NewValue + "'"))
    ErrMsg = "Invalid Design Size": Exit Sub
  Case Is = UCase("wHlpSrt")
    Cancel = (pv_NewValue <> "") And (Not moCn.RecSeek("Select PMCd From Param Where " + _
                                      "PTyp= 'SRTRTHLP' And PMCd = '" + pv_NewValue + "'"))
    ErrMsg = "Invalid Design Help Sort Option": Exit Sub
  
  '*** (Jen 2.14 Next)
  Case Is = UCase("wDRGrp")
    Cancel = Not moCn.RecSeek("select PMCd from Param " + _
             "where PTyp= 'ORDRMGR' and PMCd = '" + pv_NewValue + "'")
    ErrMsg = "Invalid Group By Value": Exit Sub
  Case Is = UCase("wIRGrp")
    Cancel = Not moCn.RecSeek("select PMCd from Param where PTyp= 'ORDRMGR' and PMCd = '" + pv_NewValue + "'")
    ErrMsg = "Invalid Group By Value": Exit Sub
  Case Is = UCase("wIDGrp")
    Cancel = Not moCn.RecSeek("select PMCd from Param " + _
             "where PTyp= 'ORDDSGGR' and PMCd = '" + pv_NewValue + "'")
    ErrMsg = "Invalid Group By Value": Exit Sub
  '*** (Jen 2.14 Next)
  
  '*************Bhavna
  Case Is = UCase("wRetRsn")
    Cancel = (Not moCn.RecSeek("select PMCd from Param where PTyp= 'RETRSN' and PMCd = '" + pv_NewValue + "'"))
    If Cancel = True Then ErrMsg = "Invalid Return Memo Default Reason Code": Exit Sub
  'IG.97 - Checks valid return category code from param, blank not allowed
  Case Is = UCase("ReRtnCtg")
    Cancel = (Not moCn.RecSeek("select PMCd from Param where PTyp= 'RTNCTG' and PMCd = '" + pv_NewValue + "'"))
    If Cancel = True Then ErrMsg = "Invalid Return Memo Category": Exit Sub
    
  End Select
End Sub
Private Sub atxt_FldValid(Index As Integer, ByVal IdName As String, ByVal OldVal As Variant)
  '*** call the respective summary routines when the user tabs out of the Summary group fields
  
  '*** (Jen 2.14 Next)
  If ADC(IdName) = OldVal Then Exit Sub
  Select Case UCase(IdName)
  Case Is = UCase("wDRGrp")
    Call DsgRmSumm
  Case Is = UCase("wIRGrp")
    Call InvRmSumm
  Case Is = UCase("wIDGrp")
    Call InvDsgSumm
  End Select
  '*** (Jen 2.14 Next)
End Sub

Private Sub CmdDsgHlpCpy_Click()
  Dim ws_IdTc As String, ws_IdYy As String, ws_IdChr As String, wi_IdNo As Integer, wi_IdSr As Integer
  With GrdDsgHlp
      ws_IdTc = .Value(.Row, "wHlpIdTc")
      ws_IdYy = .Value(.Row, "wHlpIdYy")
      ws_IdChr = .Value(.Row, "wHlpIdChr")
      wi_IdNo = .Value(.Row, "wHlpIdNo")
      wi_IdSr = .Value(.Row, "wHlpIdSr")
  End With
  Call CpyInvDets(ws_IdTc, ws_IdYy, ws_IdChr, wi_IdNo, wi_IdSr)
  
End Sub

Private Sub GrdDsgHlp_Valid()
    CmdDsgHlpCpy.Enabled = False    ' ***** Manali 3.03 - 24/06/08 - Disabling button
End Sub

Private Sub GrdDsgHlp_When()
If GrdDsgHlp.Rows > 1 Then CmdDsgHlpCpy.Enabled = True  ' ***** Manali 3.03 - 24/06/08 - Enabling button
End Sub

Private Sub GrdRetMemDet_RowWhen(ByVal RowNum As Integer)
  '*** If GrdRetMemDet.IsNew Then RedSr = Max + 1
  Dim w3DPth As String, wHskPth  As String    '*** (Jen 2.14 Next)
  
  With GrdRetMemDet
    If .IsNew(RowNum) Then
      If .Value(RowNum, "RedSr") = 0 Then .Value(RowNum, "RedSr") = .MaxVal("RedSr") + 1
      .Value(RowNum, "RedIdTc") = ctInTcTyp
      If RowNum = 1 Then
        .Value(RowNum, "RedIdYy") = ADC("ReYy")
        .Value(RowNum, "RedIdChr") = moCn.GetFldVal("Select vPValue from vParam where vPCoCd= '" + ADC("ReCoCd") + "' " + _
                                     "and vPTyp= 'DEFCHR' and vPMCd= '" + ctInTcTyp + "'")
        .Value(RowNum, "RedRetRsn") = ADC("wRetRsn")      '***************bhavna
      Else
        .Value(RowNum, "RedIdYy") = .Value(RowNum - 1, "RedIdYy")
        .Value(RowNum, "RedIdChr") = .Value(RowNum - 1, "RedIdChr")
        .Value(RowNum, "RedRetRsn") = .Value(RowNum - 1, "RedRetRsn")      '***************bhavna
      End If
      
    Else
      '*** (Jen 2.14 Next)
      FraPic3D.Left = wFraPic3DLft: FraPic3D.Top = wFraPic3DTop
      FraPic3D.Height = wFraPic3DHt: FraPic3D.Width = wFraPic3DWdt
      FraPicHsk.Left = wFraPicHskLft: FraPicHsk.Top = wFraPicHskTop
      FraPicHsk.Height = wFraPicHskHt: FraPicHsk.Width = wFraPicHskWdt
      
      w3DPth = GetPictPath("DM", False, .Value(RowNum, "wIdDmCd"))
      wHskPth = GetPictPath("DM", True, .Value(RowNum, "wIdDmCd"))
      
      Call ShowPic(wPic3DHt, wPic3DWdt, w3DPth, Pic3D)
      Call ShowPic(wPicHskHt, wPicHskWdt, wHskPth, PicHsk)
      '*** (Jen 2.14 Next)
    End If
  End With
End Sub
Private Sub GrdRetMemDet_ColWhen(ByVal RowNum As Integer, ByVal ColName As String, Cancel As Boolean, ErrMsg As String, ByVal Action As MwfCtl.en_fgAction)
  '*** Set Helps For Each Field
  If moCn.RecSeek("Select 'x' from OrdDsg join OrdMst on OmCoCd= OdCoCd and OmTc= OdTc and OmYy= OdYy and OmChr= OdChr " + _
                  "and OmNo= OdNo where OmCoCd= '" + ADC("ReCoCd") + "' and OmReYy= '" + ADC("ReYy") + "' " + _
                  "and OmReChr= '" + ADC("ReChr") + "' and OmReNo= " + CStr(ADC("ReNo")) + " " + _
                  "and OdSr= " + CStr(GrdRetMemDet.Value(RowNum, "RedSr")) + " ") Then
    Cancel = True: ErrMsg = "Cannot Edit Record As Used In Order Design": Exit Sub
  End If
  
  With GrdRetMemDet
    Select Case UCase(ColName)
    Case Is = UCase("RedIdYy")
      Call HlpList.PMCd("YY")
    Case Is = UCase("RedIdChr")
      Call HlpList.vPSCd(ADC("ReCoCd"), "CHR", GrdRetMemDet.Value(RowNum, "RedIdTc"), "Y")
    Case Is = UCase("RedIdNo")
      Call HlpList.InNo(ADC("ReCoCd"), .Value(RowNum, "RedIdTc"), .Value(RowNum, "RedIdYy"), .Value(RowNum, "RedIdChr"), ADC("ReCmCd"))
    Case Is = UCase("RedIdSr")
      Call HlpList.InSr(ADC("ReCoCd"), .Value(RowNum, "RedIdTc"), .Value(RowNum, "RedIdYy"), .Value(RowNum, "RedIdChr"), .Value(RowNum, "RedIdNo"))
    Case Is = UCase("RedQty")
      If .Value(RowNum, "RedIdSr") = 0 Then Cancel = True: ErrMsg = "First Enter Valid Invoice Serial": Exit Sub
    Case Is = UCase("RedGrsWt")
      If .Value(RowNum, "RedIdSr") = 0 Then Cancel = True: ErrMsg = "First Enter Valid Invoice Serial": Exit Sub
    '*******************Bhavna
    Case Is = UCase("RedRetRsn")
      Call HlpList.PMCd("RETRSN")
    End Select
  End With
End Sub
Private Sub GrdRetMemDet_ReCalc(ByVal RowNum As Integer, ByVal ColName As String, ByVal OldVal As Variant, ByVal SourceName As String, ByVal SourceColName As String)
  '*** RedIdNo on RedIdTc, RedIdYy, RedIdChr                             Normal
      '0
  '*** RedIdSr on RedIdTc, RedIdYy, RedIdChr, RedIdNo                    Normal
      '0
  '*** RedQty on RedIdTc, RedIdYy, RedIdChr, RedIdNo, RedIdSr            Normal, Copy
      'IdQty - sum(RedQty) for the Inv Design Serial
  '*** RedGrsWt on RedIdTc, RedIdYy, RedIdChr, RedIdNo, RedIdSr          Normal
      '0
  
  '*** (Bef 2.14 Next)
  'Recalc of WIDSFX & WIDDMSZ on GrdRetMemDet(REDIDSR)
  '*** (Bef 2.14 Next)
  
  With GrdRetMemDet
    Select Case UCase(ColName)
    Case Is = UCase("RedIdNo")
      If .Mode = fgmnorm Then
        .Value(RowNum, "RedIdNo") = 0
      End If
    Case Is = UCase("RedIdSr")
      If .Mode = fgmnorm Then
        .Value(RowNum, "RedIdSr") = 0
      End If
    Case Is = UCase("RedQty")
      If .Mode = fgmnorm Then
        Dim wd_Qty As Double
        wd_Qty = moCn.GetFldVal("Select IdQty - IsNull((Select sum(RedQty) from RetMemDet where RedCoCd= IdCoCd and " + _
                                "RedIdTc= IdTc and RedIdYy= IdYy and RedIdChr= IdChr and IdNo= RedIdNo and " + _
                                "IdSr= RedIdSr and not (RedCoCd= '" + ADC("ReCoCd") + "' and " + _
                                "RedYy= '" + ADC("ReYy") + "' and RedChr= '" + ADC("ReChr") + "' and " + _
                                "RedNo= " + CStr(ADC("ReNo")) + " and RedSr= " + CStr(.Value(RowNum, "RedSr")) + ") ), 0) " + _
                                "from InvDsg " + _
                                "where IdCoCd= '" + ADC("ReCoCd") + "' and " + _
                                "IdTc= '" + .Value(RowNum, "RedIdTc") + "' and " + _
                                "IdYy= '" + .Value(RowNum, "RedIdYy") + "' and " + _
                                "IdChr= '" + .Value(RowNum, "RedIdChr") + "' and " + _
                                "IdNo= " + CStr(.Value(RowNum, "RedIdNo")) + " and " + _
                                "IdSr= " + CStr(.Value(RowNum, "RedIdSr")) + " ")
        
        .Value(RowNum, "RedQty") = IIF(wd_Qty > 0, wd_Qty, 0)
      End If
    Case Is = UCase("RedGrsWt")
      If .Mode = fgmnorm Then
        .Value(RowNum, "RedGrsWt") = moCn.GetFldVal("Select IsNull((Select sum(IfGrsWt) from InvFgd " + _
                                     "where IfCoCd= IdCoCd and IfTc= IdTc and IfYy= IdYy " + _
                                     "and IfChr= IdChr and IfNo= IdNo and IfSr= IdSr), 0)/ " + _
                                     "(case when IdQty> 0 then IdQty else 1 end)  " + _
                                     "from InvDsg where IdCoCd= '" + ADC("ReCoCd") + "' and " + _
                                     "IdTc= '" + .Value(RowNum, "RedIdTc") + "' and " + _
                                     "IdYy= '" + .Value(RowNum, "RedIdYy") + "' and " + _
                                     "IdChr= '" + .Value(RowNum, "RedIdChr") + "' and " + _
                                     "IdNo= " + CStr(.Value(RowNum, "RedIdNo")) + " and " + _
                                     "IdSr= " + CStr(.Value(RowNum, "RedIdSr")) + " ") * .Value(RowNum, "RedQty")
      End If
      '************Bhavna
      Case Is = UCase("RedGrsWtDiff")
      If .Mode = fgmnorm Then
        Dim wActGrsWt As Double
        wActGrsWt = moCn.GetFldVal("Select IsNull((Select sum(IfGrsWt) from InvFgd " + _
                                     "where IfCoCd= IdCoCd and IfTc= IdTc and IfYy= IdYy " + _
                                     "and IfChr= IdChr and IfNo= IdNo and IfSr= IdSr), 0)/ " + _
                                     "(case when IdQty> 0 then IdQty else 1 end)  " + _
                                     "from InvDsg where IdCoCd= '" + ADC("ReCoCd") + "' and " + _
                                     "IdTc= '" + .Value(RowNum, "RedIdTc") + "' and " + _
                                     "IdYy= '" + .Value(RowNum, "RedIdYy") + "' and " + _
                                     "IdChr= '" + .Value(RowNum, "RedIdChr") + "' and " + _
                                     "IdNo= " + CStr(.Value(RowNum, "RedIdNo")) + " and " + _
                                     "IdSr= " + CStr(.Value(RowNum, "RedIdSr")) + " ") * .Value(RowNum, "RedQty")
          .Value(RowNum, "RedGrsWtDiff") = (.Value(RowNum, "RedGrsWt") - wActGrsWt)

      End If

    Case Is = UCase("RedFOBPerPc")
      If .Mode = fgmnorm Then
        Dim wd_FOBPerPc As Double
        wd_FOBPerPc = moCn.GetFldVal("Select IdIFob/ (case when IdQty> 0 then IdQty else 1 end) " + _
                                "from InvDsg where IdCoCd= '" + ADC("ReCoCd") + "' and " + _
                                "IdTc= '" + .Value(RowNum, "RedIdTc") + "' and " + _
                                "IdYy= '" + .Value(RowNum, "RedIdYy") + "' and " + _
                                "IdChr= '" + .Value(RowNum, "RedIdChr") + "' and " + _
                                "IdNo= " + CStr(.Value(RowNum, "RedIdNo")) + " and " + _
                                "IdSr= " + CStr(.Value(RowNum, "RedIdSr")) + " ")
        
        .Value(RowNum, "RedFOBPerPc") = wd_FOBPerPc
      End If
    Case Is = UCase("wRedFOB")
      If .Mode = fgmnorm Or .Mode = fgmCopy Or .Mode = fgminit Then
        .Value(RowNum, "wRedFOB") = .Value(RowNum, "RedFOBPerPc") * .Value(RowNum, "RedQty")
      End If
    Case Is = UCase("wIdDmCd")
      '*** (Bef 2.14 Next)
      'If .Mode = fgmnorm Or .Mode = fgminit Then
      '  If .Value(RowNum, "RedIdSr") > 0 Then
      '    .Value(RowNum, "wIdDmCd") = moCn.GetFldVal("Select IdDmCd " + _
      '                                "from InvDsg where IdCoCd= '" + ADC("ReCoCd") + "' and " + _
      '                                "IdTc= '" + .Value(RowNum, "RedIdTc") + "' and " + _
      '                                "IdYy= '" + .Value(RowNum, "RedIdYy") + "' and " + _
      '                                "IdChr= '" + .Value(RowNum, "RedIdChr") + "' and " + _
      '                                "IdNo= " + CStr(.Value(RowNum, "RedIdNo")) + " and " + _
      '                                "IdSr= " + CStr(.Value(RowNum, "RedIdSr")) + " ")
      '  Else
      '    .Value(RowNum, "wIdDmCd") = ""
      '  End If
      'End If
      '*** (Bef 2.14 Next)
      
      '*** (Jen 2.14 Next)
      If .Mode = fgmnorm Or .Mode = fgminit Then
        If .Value(RowNum, "RedIdSr") > 0 Then
          Dim wRsId As MwfLib.MDORowSet
          Set wRsId = moCn.OpenRes("Select IdDmCd, IdSfx, IdDmSz, IdExpCmCd " + _
                                      "from InvDsg where IdCoCd= '" + ADC("ReCoCd") + "' and " + _
                                      "IdTc= '" + .Value(RowNum, "RedIdTc") + "' and " + _
                                      "IdYy= '" + .Value(RowNum, "RedIdYy") + "' and " + _
                                      "IdChr= '" + .Value(RowNum, "RedIdChr") + "' and " + _
                                      "IdNo= " + CStr(.Value(RowNum, "RedIdNo")) + " and " + _
                                      "IdSr= " + CStr(.Value(RowNum, "RedIdSr")) + " ")
          
          .Value(RowNum, "wIdDmCd") = wRsId!IdDmCd
          .Value(RowNum, "wIdSfx") = wRsId!IdSfx
          .Value(RowNum, "wIdDmSz") = wRsId!IdDmSz
          .Value(RowNum, "wIdExpCmCd") = wRsId!IdExpCmCd
        Else
          .Value(RowNum, "wIdDmCd") = ""
          .Value(RowNum, "wIdSfx") = ""
          .Value(RowNum, "wIdDmSz") = ""
          .Value(RowNum, "wIdExpCmCd") = ""
        End If
      End If
      '*** (Jen 2.14 Next)
      
    '*** (Bef 2.14 Next)
    'Case Is = UCase("wIdSfx")
    '  If .Mode = fgmnorm Or .Mode = fgminit Then
    '    If .Value(RowNum, "RedIdSr") > 0 Then
    '      .Value(RowNum, "wIdSfx") = moCn.GetFldVal("Select IdSfx " + _
    '                                  "from InvDsg where IdCoCd= '" + ADC("ReCoCd") + "' and " + _
    '                                  "IdTc= '" + .Value(RowNum, "RedIdTc") + "' and " + _
    '                                  "IdYy= '" + .Value(RowNum, "RedIdYy") + "' and " + _
    '                                  "IdChr= '" + .Value(RowNum, "RedIdChr") + "' and " + _
    '                                  "IdNo= " + CStr(.Value(RowNum, "RedIdNo")) + " and " + _
    '                                  "IdSr= " + CStr(.Value(RowNum, "RedIdSr")) + " ")
    '    Else
    '      .Value(RowNum, "wIdSfx") = ""
    '    End If
    '  End If
    'Case Is = UCase("wIdDmSz")
    '  If .Mode = fgmnorm Or .Mode = fgminit Then
    '    If .Value(RowNum, "RedIdSr") > 0 Then
    '      .Value(RowNum, "wIdDmSz") = moCn.GetFldVal("Select IdDmSz " + _
    '                                  "from InvDsg where IdCoCd= '" + ADC("ReCoCd") + "' and " + _
    '                                  "IdTc= '" + .Value(RowNum, "RedIdTc") + "' and " + _
    '                                  "IdYy= '" + .Value(RowNum, "RedIdYy") + "' and " + _
    '                                  "IdChr= '" + .Value(RowNum, "RedIdChr") + "' and " + _
    '                                  "IdNo= " + CStr(.Value(RowNum, "RedIdNo")) + " and " + _
    '                                  "IdSr= " + CStr(.Value(RowNum, "RedIdSr")) + " ")
    '    Else
    '      .Value(RowNum, "wIdDmSz") = ""
    '    End If
    '  End If
    '*** (Bef 2.14 Next)
    
    Case Is = UCase("wInDt")
      If .Mode = fgmnorm Or .Mode = fgminit Then
        If .Value(RowNum, "RedIdNo") > 0 Then
          .Value(RowNum, "wInDt") = moCn.GetFldVal("Select InDt " + _
                                      "from InvHd where InCoCd= '" + ADC("ReCoCd") + "' and " + _
                                      "InTc= '" + .Value(RowNum, "RedIdTc") + "' and " + _
                                      "InYy= '" + .Value(RowNum, "RedIdYy") + "' and " + _
                                      "InChr= '" + .Value(RowNum, "RedIdChr") + "' and " + _
                                      "InNo= " + CStr(.Value(RowNum, "RedIdNo")) + " ")
        Else
          .Value(RowNum, "wInDt") = "01/01/80"
        End If
      End If
        
      '****** Sachin 3.02 - Filling up the New Fields
      '      Case Is = UCase("RedIdIdNo")
      '      If .Mode = fgmnorm Or .Mode = fgminit Or .Mode = fgmCopy Then
      '        If .Value(RowNum, "RedIdSr") > 0 Then
      '            .Value(RowNum, "RedIdIdNo") = moCn.GetFldVal("Select IdIdNo " + _
      '                                      "from InvDsg where IdCoCd= '" + ADC("ReCoCd") + "' and " + _
      '                                      "IdTc= '" + .Value(RowNum, "RedIdTc") + "' and " + _
      '                                      "IdYy= '" + .Value(RowNum, "RedIdYy") + "' and " + _
      '                                      "IdChr= '" + .Value(RowNum, "RedIdChr") + "' and " + _
      '                                      "IdNo= " + CStr(.Value(RowNum, "RedIdNo")) + _
      '                                      " And IdSr= " + CStr(.Value(RowNum, "RedIdSr")) + " ")
      '        Else
      '            .Value(RowNum, "RedIdIdNo") = 0
      '        End If
      '      End If
     
    End Select
  End With
End Sub
Private Sub GrdRetMemDet_ColChange(ByVal RowNum As Integer, ByVal ColName As String, ByVal NewValue As Variant, Cancel As Boolean, ErrMsg As String)
  '*** RedIdYy          valid PMCd('YY')
  '*** RedIdChr         valid vPSCd('CHR', RedIdTc) or ''
  
  With GrdRetMemDet
    Select Case UCase(ColName)
    Case Is = UCase("RedIdYy")
      Cancel = Not moCn.RecSeek("select PMCd from Param where PTyp= 'YY' and " + _
                           "PMCd = '" + NewValue + "'")
      If Cancel = True Then ErrMsg = "Invalid Invoice Voucher Year": Exit Sub
    Case Is = UCase("RedIdChr")
      Cancel = NewValue <> "" And (Not moCn.RecSeek("select vPSCd from vParam where " + _
                           "vPCoCd= '" + ADC("ReCoCd") + "' and vPTyp= 'CHR' and " + _
                           "vPMCd = '" + .Value(RowNum, "RedIdTc") + "' and vPSCd= '" + NewValue + "'"))
      If Cancel = True Then ErrMsg = "Invalid Invoice Voucher Character": Exit Sub
  
      Dim ws_UsrArr() As String, wi_Index As Integer
      ws_UsrArr = Split(moCn.GetFldVal("Select vPValue4 from vParam where vPCoCd= '" + ADC("ReCoCd") + "' and " + _
                                       "vPTyp= 'CHR' and vPMCd= '" + .Value(RowNum, "RedIdTc") + "' and " + _
                                       "vPSCd= '" + NewValue + "'"), ",")
      If IsArray(ws_UsrArr) Then
        For wi_Index = 0 To UBound(ws_UsrArr)
          Cancel = IIF(UCase(gs_UsrCd) = UCase(ws_UsrArr(wi_Index)), False, True)
          If Cancel = False Then: Exit For
        Next wi_Index
      End If
      If Cancel = True Then ErrMsg = "Access To This Invoice Voucher Character Is Denied To User " + gs_UsrCd: Exit Sub
    Case Is = UCase("RedIdNo")
      Cancel = (NewValue <> 0) And (Not moCn.RecSeek("Select InNo from InvHd " + _
                                "where InCoCd= '" + ADC("ReCoCd") + "' and " + _
                                "InTc= '" + .Value(RowNum, "RedIdTc") + "' and " + _
                                "InYy= '" + .Value(RowNum, "RedIdYy") + "' and " + _
                                "InChr= '" + .Value(RowNum, "RedIdChr") + "' and " + _
                                "InNo= " + CStr(NewValue) + " and InCmCd= '" + ADC("ReCmCd") + "'"))
      ErrMsg = "Invalid Invoice Voucher Number": Exit Sub
    Case Is = UCase("RedIdSr")
      Cancel = (NewValue <> 0) And (Not moCn.RecSeek("Select IdSr from InvDsg " + _
                                "where IdCoCd= '" + ADC("ReCoCd") + "' and " + _
                                "IdTc= '" + .Value(RowNum, "RedIdTc") + "' and " + _
                                "IdYy= '" + .Value(RowNum, "RedIdYy") + "' and " + _
                                "IdChr= '" + .Value(RowNum, "RedIdChr") + "' and " + _
                                "IdNo= " + CStr(.Value(RowNum, "RedIdNo")) + " and " + _
                                "IdSr= " + CStr(NewValue) + " "))
      ErrMsg = "Invalid Invoice Voucher Serial": Exit Sub
    Case Is = UCase("RedQty")
      Dim wd_Qty As Double
      wd_Qty = moCn.GetFldVal("Select IdQty - IsNull((Select sum(RedQty) from RetMemDet where RedCoCd= IdCoCd and " + _
                              "RedIdTc= IdTc and RedIdYy= IdYy and RedIdChr= IdChr and IdNo= RedIdNo and " + _
                              "IdSr= RedIdSr and not (RedCoCd= '" + ADC("ReCoCd") + "' and " + _
                              "RedYy= '" + ADC("ReYy") + "' and RedChr= '" + ADC("ReChr") + "' and " + _
                              "RedNo= " + CStr(ADC("ReNo")) + " and RedSr= " + CStr(.Value(RowNum, "RedSr")) + ") ), 0) " + _
                              "from InvDsg " + _
                              "where IdCoCd= '" + ADC("ReCoCd") + "' and " + _
                              "IdTc= '" + .Value(RowNum, "RedIdTc") + "' and " + _
                              "IdYy= '" + .Value(RowNum, "RedIdYy") + "' and " + _
                              "IdChr= '" + .Value(RowNum, "RedIdChr") + "' and " + _
                              "IdNo= " + CStr(.Value(RowNum, "RedIdNo")) + " and " + _
                              "IdSr= " + CStr(.Value(RowNum, "RedIdSr")) + " ")
        
      'If wd_Qty = 0 Then Cancel = True: ErrMsg = "Entire Quantity For Invoice Design Has Already Been Returned": Exit Sub
      If NewValue > wd_Qty Then Cancel = True: ErrMsg = "Invalid Invoice Design Quantity. Quantity Can Be Maximum " + CStr(wd_Qty): Exit Sub
    
    '***************Bhavna
    Case Is = UCase("RedGrsWt")
     Dim wd_MinWt As Double
      wd_MinWt = moCn.GetFldVal("Select IsNull((Select sum(case when (IrRmCtg='D' or IrRmCtg='C') then IrRmAWt/5 else IrRmAWt end) from Invrm " + _
                                     "where IrCoCd= IdCoCd and IrTc= IdTc and IrYy= IdYy " + _
                                     "and IrChr= IdChr and IrNo= IdNo and IrSr= IdSr and IrMainMet = 'N'), 0)/ " + _
                                     "(case when IdQty> 0 then IdQty else 1 end)  " + _
                                     "from InvDsg where IdCoCd= '" + ADC("ReCoCd") + "' and " + _
                                     "IdTc= '" + .Value(RowNum, "RedIdTc") + "' and " + _
                                     "IdYy= '" + .Value(RowNum, "RedIdYy") + "' and " + _
                                     "IdChr= '" + .Value(RowNum, "RedIdChr") + "' and " + _
                                     "IdNo= " + CStr(.Value(RowNum, "RedIdNo")) + " and " + _
                                     "IdSr= " + CStr(.Value(RowNum, "RedIdSr")) + " ") * .Value(RowNum, "RedQty")
    
      
      If NewValue < wd_MinWt Then Cancel = True: ErrMsg = "Invalid Invoice Design Gross Wt. Gross Wt Can Be Minimum " + CStr(wd_MinWt): Exit Sub
    
    '*********************Bhavna
    Case Is = UCase("RedRetRsn")
      Cancel = Not moCn.RecSeek("select PMCd from Param where PTyp= 'RETRSN' and " + _
                           "PMCd = '" + NewValue + "'")
      If Cancel = True Then ErrMsg = "Invalid Return Reason Code": Exit Sub
    
    '*** (Jen 2.14 Next)
    Case Is = UCase("wIdDmCd")
      'If mf_ShowDsg = True
      If .Mode = fgmnorm Then
        '****** ideally should not do in fldchange , do in fldvalid (but fldvalid is not triggered) ******
        Dim w3DPth As String, wHskPth As String
        w3DPth = GetPictPath("DM", False, NewValue)
        wHskPth = GetPictPath("DM", True, NewValue)
        
        Call ShowPic(wPic3DHt, wPic3DWdt, w3DPth, Pic3D)
        Call ShowPic(wPicHskHt, wPicHskWdt, wHskPth, PicHsk)
      End If
    '*** (Jen 2.14 Next)
    
    End Select
  End With
End Sub
Private Sub GrdRetMemDet_InitKey(ByVal RowNum As Integer)
  '*** Initialise Key Fields
  With GrdRetMemDet
      .Store "RedCoCd", ADC("ReCoCd")
      .Store "RedYy", ADC("ReYy")
      .Store "RedChr", ADC("ReChr")
      .Store "RedNo", ADC("ReNo")
      
      '****** Sachin 3.02 - Filling up the new Id fields
      If .IsNew(RowNum) Then
        .Store "RedReIdNo", moCn.GetFldVal("Select ReIdNo From RetMem" + _
                    " Where ReYy= '" + ADC("ReYy") + "'" + _
                    " And ReChr= '" + ADC("ReChr") + "'" + _
                    " And ReNo= " + CStr(ADC("ReNo")) + " " + _
                    " And ReCoCd= '" + ADC("ReCoCd") + "' ")
        
       .Store "RedPrtKey", "C"
      End If
      .Store "RedIdIdNo", moCn.GetFldVal("Select IdIdNo " + _
                                        "from InvDsg where IdCoCd= '" + ADC("ReCoCd") + "' and " + _
                                        "IdTc= '" + .Value(RowNum, "RedIdTc") + "' and " + _
                                        "IdYy= '" + .Value(RowNum, "RedIdYy") + "' and " + _
                                        "IdChr= '" + .Value(RowNum, "RedIdChr") + "' and " + _
                                        "IdNo= " + CStr(.Value(RowNum, "RedIdNo")) + _
                                        " And IdSr= " + CStr(.Value(RowNum, "RedIdSr")) + " ")
  End With
End Sub
Private Sub GrdDsgHlp_ColWhen(ByVal RowNum As Integer, ByVal ColName As String, Cancel As Boolean, ErrMsg As String, ByVal Action As MwfCtl.en_fgAction)
  '*** No change allowed in the fields of the working grid GrdDsgHlp of the Design Help frame
  Cancel = True
End Sub
Private Sub ADC_SetRecSource()
  '*** Set the RecSource for ADC
  ADC.RecSource = " Select * from RetMem " + _
                  " Where ReYy= '" + ADC("ReYy") + "'" + _
                  " And ReChr= '" + ADC("ReChr") + "'" + _
                  " And ReNo= " + CStr(ADC("ReNo")) + " " + _
                  " And ReCoCd= '" + ADC("ReCoCd") + "' "
End Sub
Private Sub GrdRetMemDet_SetRecSource()
  '*** Set the RecSource for GrdRetMemDet
  GrdRetMemDet.RecSource = " Select * from RetMemDet " + _
                           " Where RedYy= '" + ADC("ReYy") + "'" + _
                           " And RedChr= '" + ADC("ReChr") + "'" + _
                           " And RedNo= " + CStr(ADC("ReNo")) + " " + _
                           " And RedCoCd= '" + ADC("ReCoCd") + "' " + _
                           " Order By RedSr"
End Sub
Private Sub CmdCpy_Click()
  '*** Display or hide the Copy frame depending on the toggle state
  
  ' ***** Manali 3.03 - 24/06/08
  If Not moCn.RecSeek("select CmCd from CustMst where CmCtg= 'C' and CmCd = '" + ADC("ReCmCd") + "'") Then
    DispMsg "Invalid Customer Code", etError: Exit Sub
  End If
  ' ***** Manali 3.03 - 24/06/08
  
  Call DispFra(Cpy)

  If FraCpy.Visible = True Then
    Dim ws_InChr As String
    
    ADC("wCpInvTcFr").CmpStr = "InTc >=": ADC("wCpInvTcTo").CmpStr = "InTc <="
    ADC("wCpInvYyFr").CmpStr = "InYy >=": ADC("wCpInvYyTo").CmpStr = "InYy <="
    ADC("wCpInvChrFr").CmpStr = "InChr >=": ADC("wCpInvChrTo").CmpStr = "InChr <="
    ADC("wCpInvNoFr").CmpStr = "InNo >=": ADC("wCpInvNoTo").CmpStr = "InNo <="
    ADC("wCpInvSrFr").CmpStr = "IdSr >=": ADC("wCpInvSrTo").CmpStr = "IdSr <="
    ADC("wCpInvDtFr").CmpStr = "InDt >=": ADC("wCpInvDtTo").CmpStr = "InDt <="
    
    ADC("wCpInvTcFr") = ctInTcTyp: ADC("wCpInvTcTo") = ctInTcTyp
    ' ***** Manali 3.4.1 - 13/05/09 - Default values Set On button Click
''    ws_InChr = moCn.GetFldVal("Select vPValue from vParam where vPCoCd= '" + ADC("ReCoCd") + "' and vPTyp= 'DEFCHR' " + _
''                              "and vPMCd= '" + ctInTcTyp + "'")
    If ADC("wCpInvChrFr") = "" Then ADC("wCpInvChrFr") = ms_InChr     ' ws_InChr
    If ADC("wCpInvChrTo") = "" Then ADC("wCpInvChrTo") = ms_InChr   '  ws_InChr
    
    If ADC("wCpInvYyFr") = "" Then ADC("wCpInvYyFr") = ms_ReYy
    If ADC("wCpInvYyTo") = "" Then ADC("wCpInvYyTo") = ms_ReYy
    ' ***** Manali 3.4.1 - 13/05/09 - Default values Set On button Click
  Else
    ADC("wCpInvTcFr").CmpStr = "": ADC("wCpInvTcTo").CmpStr = ""
    ADC("wCpInvYyFr").CmpStr = "": ADC("wCpInvYyTo").CmpStr = ""
    ADC("wCpInvChrFr").CmpStr = "": ADC("wCpInvChrTo").CmpStr = ""
    ADC("wCpInvNoFr").CmpStr = "": ADC("wCpInvNoTo").CmpStr = ""
    ADC("wCpInvSrFr").CmpStr = "": ADC("wCpInvSrTo").CmpStr = ""
    ADC("wCpInvDtFr").CmpStr = "": ADC("wCpInvDtTo").CmpStr = ""
  End If
End Sub
Private Sub CmdCpyGo_Click()
  '*** Call the Copy Routine and Close the Copy Frame
  Call CpyInvDets
  Call CmdCpy_Click
End Sub
Private Sub CmdDsgHlp_Click()
  '*** Display or hide the Design Help frame depending on the toggle state
  
  ' ***** Manali 3.03 - 24/06/08
  If Not moCn.RecSeek("select CmCd from CustMst where CmCtg= 'C' and CmCd = '" + ADC("ReCmCd") + "'") Then
    DispMsg "Invalid Customer Code", etError: Exit Sub
  End If
  ' ***** Manali 3.03 - 24/06/08

  Call DispFra(Hlp)

  If FraDsgHlp.Visible = True Then
    ADC("wHlpDmCd").CmpStr = "IdDmCd ="
    ADC("wHlpSfx").CmpStr = "IdSfx ="
    ADC("wHlpDmSz").CmpStr = "IdDmSz ="
  Else
    ADC("wHlpDmCd").CmpStr = ""
    ADC("wHlpSfx").CmpStr = ""
    ADC("wHlpDmSz").CmpStr = ""
  End If
End Sub
Private Sub CmdDsgHlpGo_Click()
  '*** Call the Design Help Routine and Close the Design Help Frame
  
  Call ShowDsgHlp
  'Call CmdDsgHlp_Click
End Sub
Private Sub CmdSumm_Click()
  '*** (Jen 2.14 Next)
  '*** Display or hide the Summary frame depending on the toggle state
  '*** Call the Design Raw Material Summary sub procedure (as it is the first summary tab)
  
  ' ***** Manali 3.03 - 24/06/08
  If Not moCn.RecSeek("select CmCd from CustMst where CmCtg= 'C' and CmCd = '" + ADC("ReCmCd") + "'") Then
    DispMsg "Invalid Customer Code", etError: Exit Sub
  End If
  ' ***** Manali 3.03 - 24/06/08
  
  Call DispFra(Summ)
  If FraSumm.Visible = True Then
    Call InitSummFlds("wDRGrp", "hOrdRmGr")
    Call DsgRmSumm
  End If
  '*** (Jen 2.14 Next)
End Sub
Private Sub TabSumm_Click(PreviousTab As Integer)
  '*** (Jen 2.14 Next)
  '*** Whenever a tab is selected enable the frame on the clicked tab and disable the previous tab frame
  '*** Initialise the group option fields (working fields) in the current tab from the Head table
  '*** Call the Summary sub procedures associated with the current tab
  FraTabSumm(TabSumm.Tab).Enabled = True
  FraTabSumm(PreviousTab).Enabled = False
  Select Case TabSumm.Tab
  Case Is = 0
    Call InitSummFlds("wDRGrp", "hOrdRmGr")
    Call DsgRmSumm
  Case Is = 1
    Call InitSummFlds("wIDGrp", "hOrdDsgGr")
    Call InvDsgSumm
  Case Is = 2
    Call InitSummFlds("wIRGrp", "hOrdRmGr")
    Call InvRmSumm
  End Select
  '*** (Jen 2.14 Next)
End Sub

Private Sub Pic3D_DblClick()
  '*** (Jen 2.14 Next)
  '*** If wPic3DDblClk flag option is False then Enlarge the 3-Dimensional picture to fit the FraNKeyAll Frame
  '*** If wPic3DDblClk flag option is True then bring the 3-Dimensional picture to the original pic control size
  
  If GrdRetMemDet.Rows - 1 = 0 Then Exit Sub
  ADC.SetFocus
  Dim w3DPth As String
  If wPic3DDblClk = False Then
    '*** Enlarge The Picture ***
    w3DPth = GetPictPath("DM", False, GrdRetMemDet.Value(GrdRetMemDet.Row, "wIdDmCd"))
    Call ShowLargePic(IIF(FraNKeyAll.Height - 195 > 0, FraNKeyAll.Height - 195, 0), FraNKeyAll.Width, w3DPth, Pic3D)
   
    FraPic3D.Left = 0: FraPic3D.Top = 195
    FraPic3D.Height = Pic3D.Height + 100: FraPic3D.Width = Pic3D.Width + 100
    FraNKeyAll.ZOrder
    FraPic3D.ZOrder (0)
  Else
    '*** Make The Picture Small ***
    w3DPth = GetPictPath("DM", False, GrdRetMemDet.Value(GrdRetMemDet.Row, "wIdDmCd"))
    Call ShowPic(wPic3DHt, wPic3DWdt, w3DPth, Pic3D)
    FraPic3D.Left = wFraPic3DLft: FraPic3D.Top = wFraPic3DTop
    FraPic3D.Height = Pic3D.Height + 100: FraPic3D.Width = Pic3D.Width + 100
    FraNKeyAll.ZOrder
  End If
  If wPic3DDblClk = False Then
    wPic3DDblClk = True
  Else
    wPic3DDblClk = False
  End If
  '*** (Jen 2.14 Next)
End Sub
Private Sub PicHsk_DblClick()
  '*** (Jen 2.14 Next)
  '*** If wPicHskDblClk flag option is False then Enlarge the Line Drawing to fit the FraNKeyAll Frame
  '*** If wPicHskDblClk flag option is True then bring the Line Drawing to the original pic control size
  
  If GrdRetMemDet.Rows - 1 = 0 Then Exit Sub
  ADC.SetFocus    '*** 2.06
  Dim wHskPth As String
  If wPicHskDblClk = False Then
    '*** Enlarge The Picture ***
    wHskPth = GetPictPath("DM", True, GrdRetMemDet.Value(GrdRetMemDet.Row, "wIdDmCd"))
    Call ShowLargePic(IIF(FraNKeyAll.Height - 195 > 0, FraNKeyAll.Height - 195, 0), FraNKeyAll.Width, wHskPth, PicHsk)
    FraPicHsk.Left = 0: FraPicHsk.Top = 195
    FraPicHsk.Height = PicHsk.Height + 100: FraPicHsk.Width = PicHsk.Width + 100
    FraNKeyAll.ZOrder
    FraPicHsk.ZOrder (0)
  Else
    '*** Make The Picture Small ***
    wHskPth = GetPictPath("DM", True, GrdRetMemDet.Value(GrdRetMemDet.Row, "wIdDmCd"))
    Call ShowPic(wPicHskHt, wPicHskWdt, wHskPth, PicHsk)
    FraPicHsk.Left = wFraPicHskLft: FraPicHsk.Top = wFraPicHskTop
    FraPicHsk.Height = PicHsk.Height + 100: FraPicHsk.Width = PicHsk.Width + 100
    FraNKeyAll.ZOrder
  End If
  If wPicHskDblClk = False Then
    wPicHskDblClk = True
  Else
    wPicHskDblClk = False
  End If
  '*** (Jen 2.14 Next)
End Sub

Private Sub CpyInvDets(Optional ByVal ms_IdTc As String, Optional ByVal ms_IdYy As String, Optional ms_IdChr As String, _
                       Optional mi_IdNo As Integer, Optional mi_IdSr As Integer)
  ' ***** Manali 3.03 - above Optional parameters declared in procedure to enable copy from GrdDsgHlp
  '*** Copy Option is used to copy details from the specified Design Type, Design Cd and Size
  
  Dim wRow As Single, ws_Cnd As String  ', ws_OrdNo As String
  Dim ws_InvNo As String    ' ******** Manali 3.03 - 10/07/08

  If ADC("ReDt") = "" Or ADC("ReDt") = MWLib.EmptyDate Then DispMsg "Invalid Memo Voucher Date", etError: Exit Sub
  'If ADC("ReCurCd") = "" Then DispMsg "Enter Valid Currency Code Before Copying", etError: Exit Sub
  If ADC("ReCmCd") = "" Then DispMsg "Enter Valid Customer Code Before Copying", etError: Exit Sub
  If ADC("wRetRsn") = "" Then DispMsg "Enter Valid Return Reason Code Before Copying", etError: Exit Sub
  
  ' **** Manali 3.03 - 10/07/08 - To distinguish Invoice copy from DsgHlpCpy , ms_IdSr parameter checked if equal to 0
  If mi_IdSr = 0 Then
  
      '*** Check if copy is possible
          '(i.e. check if some source has been mentioned)
      ws_Cnd = ADC.RepCond
      
      'ws_Cnd = IIF(ws_Cnd <> vbNullString, " where " + ws_Cnd, " ")
      If ws_Cnd = vbNullString Then
        DispMsg "Some Scope Has To Be Mentioned", etError: Exit Sub
      'Else
      '  ws_Cnd = " and " + ws_Cnd
      End If
  ' **** Manali 3.03 - 10/07/08 - In case of copy from DsgHlp
  Else
    ws_InvNo = "IdCoCd= '" + ADC("ReCoCd") + "' and IdTc='" + ms_IdTc + "' and IdYy= '" + ms_IdYy + "' and IdChr= '" + ms_IdChr + "' and IdNo ='" + CStr(mi_IdNo) + "' and IdSr='" + CStr(mi_IdSr) + " ' "
  End If
  ' **** Manali 3.03 - 10/07/08 - In case of copy from DsgHlp
      
  '***
  'ws_OrdNo = Trim(moCn.GetFldVal("select OmTc+ '/' + OmYy+ '/' + OmChr+ '/' + LTrim(Str(OmNo)) " + _
                       "from OrdMst where OmCoCd= '" + ADC("ReCoCd") + "' " + _
                       "and OmReYy= '" + ADC("ReYy") + "' and OmReChr= '" + ADC("ReChr") + "' " + _
                       "and OmReNo= " + CStr(ADC("ReNo")) + " "))
  'If ws_OrdNo <> "" Then DispMsg "No More Records Can Be Added As Return Memo Voucher Already Used In Order: " + ws_OrdNo, etError: Exit Sub
  '***
  
  '*** Beginning the copy process
  
  'adc.StartCopy
  If ADC.SaveRec = False Then Exit Sub
  MWLib.BeginProcess Me, "Copying Invoice Design Details ..."
  
  Dim wRs_InvDsg As MwfLib.MDORowSet
  
  ' ***** Manali 3.03 - 10/07/08 - Condition for where clause passed based on if ms_IdSr parameter equal to 0
  Set wRs_InvDsg = moCn.OpenRes("Select InDt, IdTc, IdYy, IdChr, IdNo, IdSr, IdDmCd, IdSfx, IdDmSz, IdQty, IdIFob, IdExpCmCd, " + _
                                "IdQty- IsNull((Select sum(RedQty) from RetMemDet where RedCoCd= IdCoCd and RedIdTc= IdTc " + _
                                "and RedIdYy= IdYy and RedIdChr= IdChr and IdNo= RedIdNo and IdSr= RedIdSr), 0) as qQty, " + _
                                "IsNull((Select sum(IfGrsWt) from InvFgd where IfCoCd= IdCoCd and IfTc= IdTc and IfYy= IdYy " + _
                                "and IfChr= IdChr and IfNo= IdNo and IfSr= IdSr), 0) as qGrsWt " + _
                                "from InvHd join InvDsg On IdCoCd= InCoCd and IdTc= InTc and IdYy= InYy and IdChr= InChr " + _
                                "and IdNo= InNo where " + IIF(mi_IdSr = 0, ws_Cnd, ws_InvNo) + " and InCmCd= '" + ADC("ReCmCd") + "' " + _
                                "Order By IdTc, IdYy, IdChr, IdNo, IdSr")
                                '"and exists(Select 'x' from CustMst where CmCtg= 'C' and CmCd= InCmCd " + _
                                "           and CmCurCd= '" + ADC("ReCurCd") + "') "
                                
                                '"and exists(Select 'x' from InvZRm where IzCoCd= InCoCd and IzTc= InTc and IzYy= InYy " + _
                                "           and IzChr= InChr and IzNo= InNo) "
 
 Dim wIdDmCdStr As String
 If wRs_InvDsg.RecCount > 0 Then
   With wRs_InvDsg
      wIdDmCdStr = ""
      .MoveFirst
      Do While Not (.BOF Or .EOF)
       If Not moCn.RecSeek("Select DmCd From DsgMst where DmTcTyp='DM' " + _
                          " and DmCd = '" + wRs_InvDsg!IdDmCd + "'  and DmSz='' and DmPrtKey='" + ctCurrPrtn + "' ") Then _
        wIdDmCdStr = IIF(wIdDmCdStr = "", "", wIdDmCdStr + ",") + wRs_InvDsg!IdDmCd
       .MoveNext
      Loop
      .MoveFirst
    End With
  End If
  If wIdDmCdStr <> "" Then DispMsg "Cannot Copy Design(s) :" + wIdDmCdStr + " Not In Current Partition.", etError: GoTo EndCopy
                                
  If wRs_InvDsg.RecCount = 0 Then DispMsg "No Record Found For The Given Scope/ Currency Mismatch", etError: GoTo EndCopy
  'If wRs_InvDsg.RecCount = 0 Then DispMsg "No Record Found For The Given Scope/ Custom Matching Not Done For The Invoice Recs/ Currency Mismatch", etError: GoTo EndCopy
    
  ' **** Manali 3.03 - 10/07/08 - Cpy from DsgHlp
  If mi_IdSr <> 0 And wRs_InvDsg!qQty <= 0 Then _
    DispMsg "Cannot Copy Since Invoice Design Quantity Available For Return Is Zero ", etError: GoTo EndCopy
  ' **** Manali 3.03 - 10/07/08 - Cpy from DsgHlp
  
  ' ***** Manali 3.03 - 18/09/08 - to give errmsg in case of copying multiple voucher serials
  Dim wi_RedCnt As Integer
  If mi_IdSr = 0 Then
    wi_RedCnt = moCn.GetFldVal("Select count(*) from InvHd join InvDsg On IdCoCd= InCoCd and " + _
                               "IdTc= InTc and IdYy= InYy and IdChr= InChr and IdNo= InNo where " + _
                               ws_Cnd + " and InCmCd= 'CI' ")
  End If
  ' ***** Manali 3.03
  
  With wRs_InvDsg
    GrdRetMemDet.StartCopy
    Do While Not (.EOF Or .BOF)
      If !qQty > 0 Then
        MWLib.RunningProcess "Copying Details For " + !IdTc + "/" + !IdYy + "/" + !IdChr + "/" + Trim(CStr(!IdNo)) + "/" + Trim(CStr(!IdSr))
        GrdRetMemDet.AddItem
        wRow = GrdRetMemDet.Rows - 1
        If GrdRetMemDet.Value(wRow, "RedSr") = 0 Then GrdRetMemDet.Value(wRow, "RedSr") = GrdRetMemDet.MaxVal("RedSr") + 1
        GrdRetMemDet.Value(wRow, "RedIdTc") = !IdTc
        GrdRetMemDet.Value(wRow, "RedIdYy") = !IdYy
        GrdRetMemDet.Value(wRow, "RedIdChr") = !IdChr
        GrdRetMemDet.Value(wRow, "RedIdNo") = !IdNo
        GrdRetMemDet.Value(wRow, "RedIdSr") = !IdSr
        GrdRetMemDet.Value(wRow, "RedQty") = !qQty
        GrdRetMemDet.Value(wRow, "RedGrsWt") = MWLib.Div(!qGrsWt, !IdQty) * !qQty
        '**********bhavna
        GrdRetMemDet.Value(wRow, "RedGrsWtDiff") = 0
        '***********bhavna
        
        GrdRetMemDet.Value(wRow, "RedFOBPerPc") = MWLib.Div(!IdIFob, !IdQty)
        GrdRetMemDet.Value(wRow, "wIdDmCd") = !IdDmCd
        GrdRetMemDet.Value(wRow, "wIdSfx") = !IdSfx
        GrdRetMemDet.Value(wRow, "wIdDmSz") = !IdDmSz
        GrdRetMemDet.Value(wRow, "wIdExpCmCd") = !IdExpCmCd     '*** (Jen 2.14 Next)
        GrdRetMemDet.Value(wRow, "wInDt") = !InDt
        GrdRetMemDet.Value(wRow, "RedRetRsn") = ADC("wRetRsn")      '***********Bhavna
        
        'GrdRetMemDet.Value(wRow, "wRedFOB") = GrdRetMemDet.Value(wRow, "RedFOBPerPc") * !qQty
        GrdRetMemDet.SaveRec (wRow)
      End If
      .MoveNext
    Loop
    GrdRetMemDet.EndCopy
  End With
  
EndCopy:

  ' ****** Manali 3.03 - 18/09/08
  If wi_RedCnt > 0 Then DispMsg "Some of the Designs from Selected Scope are already returned in Return Memo Voucher ", etWarning
  ' ****** Manali 3.03 - 18/09/08

  MWLib.EndProcess Me
  'adc.EndCopy

  Set wRs_InvDsg = Nothing
  '*** End of the copy process
End Sub


Private Sub ShowDsgHlp()
  '*** Design Help Option is used to display the Invoice Details for the specified Design, Suffix and Size
  
  Dim wRow As Single, ws_Cnd As String, ws_Srt As String

  '*** Check if showing the help is possible
      '(i.e. check if some scope has been mentioned)
  ws_Cnd = ADC.RepCond
  
  If ws_Cnd = vbNullString Then
    DispMsg "Some Scope Has To Be Mentioned For Showing Help", etError: Exit Sub
  End If
  
  ' ***** Manali 3.03 - 14/07/08
  If ADC("wHlpDmCd") = "" Then DispMsg "Design Code Cannot Be Blank", etError: Exit Sub

  If ADC("wHlpSrt") = "" Then DispMsg "Help Sort Option Cannot Be Blank", etError: Exit Sub
  
  
  Select Case UCase(ADC("wHlpSrt"))
  Case Is = UCase("DSG")
    ws_Srt = " Order By IdDmCd, IdSfx, IdDmSz, IdTc, IdYy, IdChr, IdNo, IdSr "
  Case Is = UCase("INV")
    ws_Srt = " Order By IdTc, IdYy, IdChr, IdNo, IdSr "
  Case Is = UCase("INVDT")
    ws_Srt = " Order By InDt, IdTc, IdYy, IdChr, IdNo, IdSr "
  End Select
  
  Dim wRs_DsgHlp As MwfLib.MDORowSet
  
  ' ******* Manali 3.2.2 - Inv Customer condition added, help will now show only vouchers for inv customer
  ' ***** Manali 3.9.0 - Kt, Col, Grs Wt details in Help Grid - OrdDsg Join added
  ' ***** Manali 3.9.0 - 07/02/12 - tcTyp=InTcTyp Cond added
  Set wRs_DsgHlp = moCn.OpenRes("Select InDt, IdTc, IdYy, IdChr, IdNo, IdSr, IdDmCd, IdSfx, IdDmSz, IdQty, IdIFob, IdExpCmCd, OdDmCol, IdKt, " + _
                                "IsNull((Select sum(IfGrsWt) from InvFgd where IfCoCd= IdCoCd and IfTc= IdTc and IfYy= IdYy " + _
                                "and IfChr= IdChr and IfNo= IdNo and IfSr= IdSr), 0) as qGrsWt " + _
                                "from InvHd " + _
                                "join InvDsg On IdCoCd= InCoCd and IdTc= InTc and IdYy= InYy and IdChr= InChr and IdNo= InNo " + _
                                "Join OrdDsg On IdCoCd=OdCoCd and IdExpOdTc=OdTc and IdExpOdYy=OdYy and IdExpOdChr=OdChr and IdExpOdNo=OdNo and IdExpOdSr=OdSr " + _
                                "where IdCoCd= '" + ADC("ReCoCd") + "' and InCmCd= '" + ADC("ReCmCd") + "' and " + _
                                "(select PValue From Param Where PTyp='TC' and PMCd=InTc)='" + ctInTcTyp + "' and " + ws_Cnd + " " + ws_Srt)
                                
  If wRs_DsgHlp.RecCount = 0 Then
    GrdDsgHlp.Rows = 1    ' ***** Manali 3.03 - 24/06/08 - emtying grid
    CmdDsgHlpCpy.Enabled = False  '******** Disabling button when blank grid
    DispMsg "No Record Found For The Given Scope", etError: Exit Sub
  End If
  
  GrdDsgHlp.AllowAdd = True
  GrdDsgHlp.Rows = 1
  
  With wRs_DsgHlp
    Do While Not (.EOF Or .BOF)
      If !IdQty > 0 Then
        GrdDsgHlp.AddItem
        wRow = GrdDsgHlp.Rows - 1
        GrdDsgHlp.Value(wRow, "wHlpDsgCd") = !IdDmCd
        GrdDsgHlp.Value(wRow, "wHlpDsgSfx") = !IdSfx
        GrdDsgHlp.Value(wRow, "wHlpDsgSz") = !IdDmSz
        GrdDsgHlp.Value(wRow, "wHlpIdTc") = !IdTc
        GrdDsgHlp.Value(wRow, "wHlpIdYy") = !IdYy
        GrdDsgHlp.Value(wRow, "wHlpIdChr") = !IdChr
        GrdDsgHlp.Value(wRow, "wHlpIdNo") = !IdNo
        GrdDsgHlp.Value(wRow, "wHlpIdSr") = !IdSr
        GrdDsgHlp.Value(wRow, "wHlpIdQty") = !IdQty
        GrdDsgHlp.Value(wRow, "wHlpInvDt") = !InDt
        GrdDsgHlp.Value(wRow, "wHlpIdExpCmCd") = !IdExpCmCd   '*** (Jen 2.14 Next)
        GrdDsgHlp.Value(wRow, "wHlpIdFOBPPc") = MWLib.Div(!IdIFob, !IdQty)
        ' ***** Manali 3.9.0 - Kt, Col, Grs Wt details in Help Grid
        GrdDsgHlp.Value(wRow, "wHlpIdKt") = !IdKt
        GrdDsgHlp.Value(wRow, "wHlpDmCol") = !OdDmCol
        GrdDsgHlp.Value(wRow, "wHlpGrsWt") = !qGrsWt
        ' ***** Manali 3.9.0 - Kt, Col, Grs Wt details in Help Grid
      End If
      .MoveNext
    Loop
  End With
  GrdDsgHlp.AllowAdd = False

  Set wRs_DsgHlp = Nothing
  '*** End of the copy process
End Sub
Private Sub InitSummFlds(ByVal wReqIdName As String, ByVal wReqFldName As String)
  '*** (Jen 2.14 Next)
  '*** Procedure to initialise the Group Option Fields in the summary frame
  Dim wOrgVal As String
  If ADC(wReqIdName) = "" Then
    wOrgVal = moCn.GetFldVal("Select " + wReqFldName + " from Head where HCoCd= '" + gs_CoCd + "' and HCd='" + ctSelfCmCd + "'")
  Else
    wOrgVal = ADC(wReqIdName)
  End If
  ADC(wReqIdName).InitialiseFld ""
  ADC(wReqIdName) = wOrgVal
  '*** (Jen 2.14 Next)
End Sub
Private Sub DsgRmSumm()
  '*** (Jen 2.14 Next)
  '*** Procedure to fill up the Design Rm Summary grid in the Summary Frame
  '*** The grid is filled on the basis of the Design Rm Group option selected by the user
  If GrdRetMemDet.Row <= 0 Then Exit Sub
  
  Dim wDsgRmTotIWt As Double, wDsgRmTotQty As Double
  Dim wDsgRmTotIVal As Double, wRsDsgRmSumm As MwfLib.MDORowSet
  
  'Dim wDsgRmTotZWt As Double, wDsgRmTotZVal As Double
  
  Dim wRow As Integer, wsql As String
  
  wDsgRmTotIWt = 0: wDsgRmTotQty = 0: wDsgRmTotIVal = 0
  'wDsgRmTotZWt = 0: wDsgRmTotZVal = 0
  wsql = RmSummSql(ctDsgSumm)
  GrdSummDR.Rows = 1
  If wsql = "" Then DispMsg "Invalid GroupBy option", etError: Exit Sub
  Set wRsDsgRmSumm = moCn.OpenRes(wsql)
  GrdSummDR.AllowAdd = True
  With GrdSummDR
    Do While Not (wRsDsgRmSumm.EOF Or wRsDsgRmSumm.BOF)
      .AddItem
      wRow = .Rows - 1
      .Value(wRow, "wDrCtg") = wRsDsgRmSumm!qRmCtg
      .Value(wRow, "wDrSCtg") = wRsDsgRmSumm!qRmSCtg
      .Value(wRow, "wDrCd") = wRsDsgRmSumm!IrRmCd
      .Value(wRow, "wDrQty") = wRsDsgRmSumm!IrRmQty
      .Value(wRow, "wDrIWt") = wRsDsgRmSumm!IrRmIWt
      .Value(wRow, "wDrIVal") = wRsDsgRmSumm!IrRmIVal
      '.Value(wRow, "wDrZWt") = wRsDsgRmSumm!IrRmZWt
      '.Value(wRow, "wDrZVal") = wRsDsgRmSumm!IrRmZVal
      Select Case Trim(wRsDsgRmSumm!qRmCtg)
      ' ***** Manali 3.03 - 23/06/08 - Pd Mod - "L" added
      Case Is = "G", "S", "P", "L", "M", "X"
        wDsgRmTotIWt = wDsgRmTotIWt + wRsDsgRmSumm!IrRmIWt
        'wDsgRmTotZWt = wDsgRmTotZWt + wRsDsgRmSumm!IrRmZWt
      Case Is = "D", "C"
        wDsgRmTotIWt = wDsgRmTotIWt + MWLib.Div(wRsDsgRmSumm!IrRmIWt, 5)
        'wDsgRmTotZWt = wDsgRmTotZWt + MWLib.Div(wRsDsgRmSumm!IrRmZWt, 5)
      End Select
      wDsgRmTotQty = wDsgRmTotQty + wRsDsgRmSumm!IrRmQty
      wDsgRmTotIVal = wDsgRmTotIVal + wRsDsgRmSumm!IrRmIVal
      'wDsgRmTotZVal = wDsgRmTotZVal + wRsDsgRmSumm!IrRmZVal
      wRsDsgRmSumm.MoveNext
    Loop
  End With
  
  ADC("wDRTotIWt") = wDsgRmTotIWt
  'ADC("wDRTotZWt") = wDsgRmTotZWt
  ADC("wDRTotQty") = wDsgRmTotQty
  ADC("wDRTotIVal") = wDsgRmTotIVal
  'ADC("wDRTotZVal") = wDsgRmTotZVal
  GrdSummDR.AllowAdd = False
  '*** (Jen 2.14 Next)
End Sub
Private Sub InvRmSumm()
  '*** (Jen 2.14 Next)
  '*** Procedure to fill up the Invoice Rm Summary grid in the Summary Frame
  '*** The grid is filled on the basis of the Invoice Rm Group option selected by the user
  
  Dim wInvRmTotIWt As Double, wInvRmTotQty As Double
  Dim wInvRmTotIVal As Double, wRsInvRmSumm As MwfLib.MDORowSet
  
  'Dim wInvRmTotZWt As Double, wInvRmTotZVal As Double
  Dim wRow As Integer, wsql As String
  wInvRmTotIWt = 0: wInvRmTotQty = 0: wInvRmTotIVal = 0
  'wInvRmTotZWt = 0: wInvRmTotZVal = 0
  wsql = RmSummSql(ctInvSumm)
  GrdSummIR.Rows = 1
  If wsql = "" Then DispMsg "Invalid GroupBy option", etInfo: Exit Sub
  Set wRsInvRmSumm = moCn.OpenRes(wsql)
  GrdSummIR.AllowAdd = True
  With GrdSummIR
    Do While Not (wRsInvRmSumm.EOF Or wRsInvRmSumm.BOF)
      .AddItem
      wRow = .Rows - 1
      .Value(wRow, "wIrCtg") = wRsInvRmSumm!qRmCtg
      .Value(wRow, "wIrSCtg") = wRsInvRmSumm!qRmSCtg
      .Value(wRow, "wIrCd") = wRsInvRmSumm!IrRmCd
      .Value(wRow, "wIrQty") = wRsInvRmSumm!IrRmQty
      .Value(wRow, "wIrIWt") = wRsInvRmSumm!IrRmIWt
      .Value(wRow, "wIrIVal") = wRsInvRmSumm!IrRmIVal
      '.Value(wRow, "wIrZWt") = wRsInvRmSumm!IrRmZWt
      '.Value(wRow, "wIrZVal") = wRsInvRmSumm!IrRmZVal
      Select Case Trim(wRsInvRmSumm!qRmCtg)
      ' ***** Manali 3.03 - 23/06/08 - Pd Mod - "L" added
      Case Is = "G", "S", "P", "L", "M", "X"
        wInvRmTotIWt = wInvRmTotIWt + wRsInvRmSumm!IrRmIWt
        'wInvRmTotZWt = wInvRmTotZWt + wRsInvRmSumm!IrRmZWt
      Case Is = "D", "C"
        wInvRmTotIWt = wInvRmTotIWt + MWLib.Div(wRsInvRmSumm!IrRmIWt, 5)
        'wInvRmTotZWt = wInvRmTotZWt + MWLib.Div(wRsInvRmSumm!IrRmZWt, 5)
      End Select
      wInvRmTotQty = wInvRmTotQty + wRsInvRmSumm!IrRmQty
      wInvRmTotIVal = wInvRmTotIVal + wRsInvRmSumm!IrRmIVal
      'wInvRmTotZVal = wInvRmTotZVal + wRsInvRmSumm!IrRmZVal
      wRsInvRmSumm.MoveNext
    Loop
  End With
  
  ADC("wIrTotIWt") = wInvRmTotIWt
  'ADC("wIrTotZWt") = wInvRmTotZWt
  ADC("wIrTotQty") = wInvRmTotQty
  ADC("wIrTotIVal") = wInvRmTotIVal
  'ADC("wIrTotZVal") = wInvRmTotZVal
  GrdSummIR.AllowAdd = False
  '*** (Jen 2.14 Next)
End Sub
Private Sub InvDsgSumm()
  '*** (Jen 2.14 Next)
  '*** Procedure to fill up the Invoice Design Summary grid in the Summary Frame
  '*** The grid is filled on the basis of the Invoice Design Group option selected by the user
  
  Dim wInvDsgTotQty As Double, wInvDsgTotiFOB As Double, wInvDsgTotZFOB As Double, wReqGrpBy As String, wStr As String, wCommStr As String, wRow As Integer
  Dim wRsInvDsgSumm As MDORowSet
  Select Case ADC("wIDGrp")
  Case Is = "DMCD"
    wStr = "Select '*' as IdSr, DmCtg as qDmCtg, IdDmCd, 0 as qIPrc, 0 as qZPrc, "
    wReqGrpBy = "DmCtg, IdDmCd"
  Case Is = "DMCTG"
    wStr = "Select '*' as IdSr, DmCtg as qDmCtg, '*' as IdDmCd, 0 as qIPrc, 0 as qZPrc, "
    wReqGrpBy = "DmCtg"
  Case Is = "DSGSR"
    wStr = "Select IdTc+ '/'+ IdYy+ '/'+ IdChr+ '/'+ LTrim(Str(IdNo))+ '/'+ LTrim(Str(IdSr)) as IdSr, " + _
           "max(DmCtg) as qDmCtg, max(IdDmCd) as IdDmCd, max(case when IdQty> 0 then IdIFOB/IdQty else 0 end) as qIPrc, " + _
           "max(case when IdQty> 0 then IdZFOB/IdQty else 0 end) as qZPrc, "
    wReqGrpBy = "IdTc, IdYy, IdChr, IdNo, IdSr"
  Case Else
    DispMsg "Enter Invoice Design Group Option", etError: Exit Sub
  End Select
  
  wInvDsgTotQty = 0: wInvDsgTotiFOB = 0: wInvDsgTotZFOB = 0
  
  '*** (Bef 3.01)
  'wCommStr = "sum(IdQty) as IdQty, sum(IdIFOB) as IdIFOB, sum(IdZFOB) as IdZFOB " + _
  '           "from RetMemDet " + _
  '           "join InvDsg on IdCoCd= RedCoCd and IdTc= RedIdTc and IdYy= RedIdYy and IdChr= RedIdChr " + _
  '           "    and IdNo= RedIdNo and IdSr= RedIdSr " + _
  '           "join DsgMst on DmTcTyp= 'DM' and DmCd= IdDmCd and DmSz= '' " + _
  '           "Where RedCoCd= '" + gs_CoCd + "' and RedYy= '" + ADC("ReYy") + "' " + _
  '           "and RedChr= '" + ADC("ReChr") + "' and RedNo= " + CStr(ADC("ReNo")) + " group by " + wReqGrpBy
  '*** (Bef 3.01)
               
  '*** (Jen 3.01)
  wCommStr = "sum(RedQty) as IdQty, sum(IdIFOB* RedQty/ IdQty) as IdIFOB, sum(IdZFOB* RedQty/ IdQty) as IdZFOB " + _
             "from RetMemDet " + _
             "join InvDsg on IdCoCd= RedCoCd and IdTc= RedIdTc and IdYy= RedIdYy and IdChr= RedIdChr " + _
             "    and IdNo= RedIdNo and IdSr= RedIdSr " + _
             "join DsgMst on DmTcTyp= 'DM' and DmCd= IdDmCd and DmSz= '' " + _
             "Where RedCoCd= '" + gs_CoCd + "' and RedYy= '" + ADC("ReYy") + "' " + _
             "    and RedChr= '" + ADC("ReChr") + "' and RedNo= " + CStr(ADC("ReNo")) + _
             "    and convert (decimal (14, 1), RedQty)> convert (decimal (14, 1), 0) " + _
             "group by " + wReqGrpBy
  '*** (Jen 3.01)
               
  wStr = wStr + wCommStr
  
  Set wRsInvDsgSumm = moCn.OpenRes(wStr)
  
  GrdSummID.AllowAdd = True
  GrdSummID.Rows = 1
  
  With GrdSummID
    Do While Not (wRsInvDsgSumm.EOF Or wRsInvDsgSumm.BOF)
      .AddItem
      wRow = .Rows - 1
      .Value(wRow, "wIdSr") = wRsInvDsgSumm!IdSr
      .Value(wRow, "wIdDmCtg") = wRsInvDsgSumm!qDmCtg
      .Value(wRow, "wIdDmCd") = wRsInvDsgSumm!IdDmCd
      .Value(wRow, "wIdQty") = wRsInvDsgSumm!IdQty
      .Value(wRow, "wIdIPrc") = wRsInvDsgSumm!qIPrc
      .Value(wRow, "wIdIFOB") = wRsInvDsgSumm!IdIFob
      '.Value(wRow, "wIdZPrc") = wRsInvDsgSumm!qZPrc
      '.Value(wRow, "wIdZFOB") = wRsInvDsgSumm!IdZFob
      wInvDsgTotQty = wInvDsgTotQty + wRsInvDsgSumm!IdQty
      wInvDsgTotiFOB = wInvDsgTotiFOB + wRsInvDsgSumm!IdIFob
      'wInvDsgTotZFOB = wInvDsgTotZFOB + wRsInvDsgSumm!IdZFob
      wRsInvDsgSumm.MoveNext
    Loop
  End With
  ADC("wIdTotQty") = wInvDsgTotQty
  ADC("wIdTotiFOB") = wInvDsgTotiFOB
  'ADC("wIdTotZFob") = wInvDsgTotZFOB
  GrdSummID.AllowAdd = False
  '*** (Jen 2.14 Next)
End Sub
Private Function RmSummSql(ByVal mInvDsgSumm As String) As String
  '*** (Jen 2.14 Next)
  '*** Common Procedure used in the Voucher and Design Rm Summary option
  
  Dim wCommStr As String, wCancel As Boolean, wErrMsg As String, wReqCnd As String, wReqIdName As String, wReqGrpBy As String
  
  If mInvDsgSumm = ctDsgSumm Then
    wReqCnd = " and RedIdTc= '" + GrdRetMemDet.Value(GrdRetMemDet.Row, "RedIdTc") + "' " + _
              " and RedIdYy= '" + GrdRetMemDet.Value(GrdRetMemDet.Row, "RedIdYy") + "' " + _
              " and RedIdChr= '" + GrdRetMemDet.Value(GrdRetMemDet.Row, "RedIdChr") + "' " + _
              " and RedIdNo= " + CStr(GrdRetMemDet.Value(GrdRetMemDet.Row, "RedIdNo")) + " " + _
              " and RedIdSr= " + CStr(GrdRetMemDet.Value(GrdRetMemDet.Row, "RedIdSr"))
    wReqIdName = "wDRGrp"
  ElseIf mInvDsgSumm = ctInvSumm Then
    wReqIdName = "wIRGrp"
  End If
  
  Call ATXT_FldChange(99, wReqIdName, wCancel, wErrMsg, ADC(wReqIdName))
  If wCancel = True Then
    DispMsg wErrMsg, etError
    Exit Function
  End If
  Select Case ADC(wReqIdName)
  Case Is = "RMCTG"
    wReqGrpBy = "RmCtg"
    RmSummSql = "Select RmCtg as qRmCtg, '*' as qRmSCtg, '*' as IrRmCd, "
  Case Is = "RMSCTG"
    wReqGrpBy = "RmCtg,RmSCtg"
    RmSummSql = "Select RmCtg as qRmCtg, RmSCtg as qRmSCtg, '*' as IrRmCd, "
  Case Is = "RMCD"
    wReqGrpBy = "RmCtg,RmSCtg,IrRmCd"
    RmSummSql = "Select RmCtg as qRmCtg, RmSCtg as qRmSCtg, IrRmCd, "
  Case Else
    DispMsg "No Grouping Flds defined in system for selected GroupBy value", etError
    Exit Function
  End Select

  '*** (Bef 3.01)
  'wCommStr = "sum(IsNull(IrRmQty, 0)) as IrRmQty, sum(IsNull(IrRmIWt, 0)) as IrRmIWt, " + _
  '           "sum(IsNull(IrRmIVal, 0)) as IrRmIVal, sum(IsNull(IrRmZWt, 0)) as IrRmZWt, " + _
  '           "sum(IsNull(IrRmZVal, 0)) as IrRmZVal " + _
  '           "from RetMemDet join InvRm on IrCoCd= RedCoCd and IrTc= RedIdTc and IrYy= RedIdYy and IrChr= RedIdChr " + _
  '           "    and IrNo= RedIdNo and IrSr= RedIdSr " + _
  '           "join RmMst On RmCd= IrRmCd " + _
  '           "where RedCoCd= '" + gs_CoCd + "' and RedYy= '" + ADC("ReYy") + "' " + _
  '           "and RedChr= '" + ADC("ReChr") + "' and RedNo= " + CStr(ADC("ReNo")) + " " + wReqCnd + _
  '           " Group By " + wReqGrpBy
  '*** (Bef 3.01)
  
  '*** (Jen 3.01)
  wCommStr = "sum(IsNull(IrRmQty, 0)*RedQty/IdQty) as IrRmQty, sum(IsNull(IrRmIWt, 0)*RedQty/IdQty) as IrRmIWt, " + _
             "sum(IsNull(IrRmIVal, 0)*RedQty/IdQty) as IrRmIVal, sum(IsNull(IrRmZWt, 0)*RedQty/IdQty) as IrRmZWt, " + _
             "sum(IsNull(IrRmZVal, 0)*RedQty/IdQty) as IrRmZVal " + _
             "from RetMemDet join InvDsg on IdCoCd= RedCoCd and IdTc= RedIdTc and IdYy= RedIdYy and IdChr= RedIdChr " + _
             "    and IdNo= RedIdNo and IdSr= RedIdSr " + _
             "join InvRm on IrCoCd= RedCoCd and IrTc= RedIdTc and IrYy= RedIdYy and IrChr= RedIdChr " + _
             "    and IrNo= RedIdNo and IrSr= RedIdSr " + _
             "join RmMst On RmCd= IrRmCd " + _
             "where RedCoCd= '" + gs_CoCd + "' and RedYy= '" + ADC("ReYy") + "' " + _
             "and RedChr= '" + ADC("ReChr") + "' and RedNo= " + CStr(ADC("ReNo")) + " " + wReqCnd + _
             " and convert (decimal (14, 1), RedQty)> convert (decimal (14, 1), 0) " + _
             " Group By " + wReqGrpBy
  '*** (Jen 3.01)
  
  RmSummSql = RmSummSql + wCommStr
  '*** (Jen 2.14 Next)
End Function

Private Sub DispFra(ByVal pv_RetMemoFra As en_RetMemoFra)
  '*** This is the common procedure which is used to toggle between hiding and displaying
      'the frames (i.e. if the frame is visible then make it invisible and viseversa)
      'associated with a command button
  Select Case pv_RetMemoFra
  Case Is = Cpy
    '*** If Frame FraCpy is visible then make it invisible and enable all the command buttons
    '*** If Frame FraCpy is invisible then make it visible and disable all the
        'command buttons except the one which is clicked. Change the background colour of
        'the clicked command button to white and make the font bold
    If FraCpy.Visible = True Then
      FraCpy.Visible = False
      FraCpy.Enabled = False
      Call EnaDisaCmds(False, CmdCpy)
      CmdCpy.SetFocus
    Else
      FraCpy.Visible = True
      FraCpy.Enabled = True
      ADC("wCpInvYyFr").SetFocus
      ' ***** Manali 3.3.0
      If ADC("wCpInvDtFr") = "" Then ADC("wCpInvDtFr") = MWLib.EmptyDate
      If ADC("wCpInvDtTo") = "" Then ADC("wCpInvDtTo") = MWLib.EmptyDate
      ' ***** Manali 3.3.0
      FraCpy.ZOrder
      Call EnaDisaCmds(True, CmdCpy)
    End If
  Case Is = Hlp
    '*** If Frame FraDsgHlp is visible then make it invisible and enable all the command buttons
    '*** If Frame FraDsgHlp is invisible then make it visible and disable all the
        'command buttons except the one which is clicked. Change the background colour of
        'the clicked command button to white and make the font bold
    If FraDsgHlp.Visible = True Then
      FraDsgHlp.Visible = False
      FraDsgHlp.Enabled = False
      Call EnaDisaCmds(False, CmdDsgHlp)
      'CmdDsgHlp.SetFocus
    Else
      FraDsgHlp.Visible = True
      FraDsgHlp.Enabled = True
      ADC("wHlpDmCd").SetFocus
      FraDsgHlp.ZOrder
      CmdDsgHlpCpy.Enabled = False    ' ***** Manali 3.03 - 24/06/08
      If ADC("wHlpSrt") = "" Then ADC("wHlpSrt") = UCase(moCn.GetFldVal("Select PMCd from Param where PTyp='SRTRTHLP' and PMCd='INVDT'"))     ' ***** Manali 3.03
      Call EnaDisaCmds(True, CmdDsgHlp)
    End If
  
  '*** (Jen 2.14 Next)
  Case Is = Summ
    If FraSumm.Visible = True Then
      FraSumm.Visible = False
      FraSumm.Enabled = False
      Call EnaDisaCmds(False, CmdSumm)
      CmdSumm.SetFocus
    Else
      FraSumm.Visible = True
      FraSumm.Enabled = True
      TabSumm.SetFocus
      TabSumm.Tab = 0
      FraTabSumm(TabSumm.Tab).Enabled = True
      FraSumm.ZOrder
      Call EnaDisaCmds(True, CmdSumm)
    End If
  '*** (Jen 2.14 Next)
  
  End Select
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
      If pv_CmdBut.Pic = 0 Then
        pv_CmdBut.BackColor = vbWhite
        'pv_CmdBut.FontBold = True
      End If
      pv_CmdBut.CausesValidation = True
    Else
      If pv_CmdBut.Pic = 0 Then
        pv_CmdBut.BackColor = &H8000000F
        'pv_CmdBut.FontBold = False
      End If
      'pv_CmdBut.CausesValidation = False
    End If
  End If
  
  FraNKeyAll.Enabled = Not pv_ShowFra
  ADC.AllowSave = Not pv_ShowFra
  ADC.AllowDelete = Not pv_ShowFra
  CmdCpy.Enabled = Not pv_ShowFra
  CmdDsgHlp.Enabled = Not pv_ShowFra
  CmdSumm.Enabled = Not pv_ShowFra    '*** (Jen 2.14 Next)
  
  If Not pv_CmdBut Is Nothing Then pv_CmdBut.Enabled = True
End Sub
Private Sub HideAllFras()
  '*** This procedure hides all the frames other than FraNKeyAll
  FraNKeyAll.Visible = True
  
  FraCpy.Visible = False
  FraCpy.Enabled = False

  FraDsgHlp.Visible = False
  FraDsgHlp.Enabled = False
  
  '*** (Jen 2.14 Next)
  FraSumm.Visible = False
  FraSumm.Enabled = False
  '*** (Jen 2.14 Next)
End Sub


'''  ***** Std Code not to be changed **********************************
Private Sub Form_GotFocus()
  '*** Jenny Color (For Setting System Colors)
  SetMDISysColorsAfterLogin
  '*** Jenny Color (For Setting System Colors)
End Sub
Private Sub Form_Load()
  '*** Jenny Colour
  Dim ws_BtnStr As String
  ws_BtnStr = "CmdCpyGo,CmdDsgHlpGo"
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
  
  Set moCn = Nothing   '*** (Bef speed) bef 24/11/06
End Sub
Private Sub GrdRetMemDet_Validate(Cancel As Boolean)
  Cancel = GrdRetMemDet.Validate
End Sub
