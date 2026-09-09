VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Object = "{AB2CC32A-E652-42A2-B341-26ECE02F6386}#2.1#0"; "MwfCtlForSpeed8.ocx"
Begin VB.Form EmrFrmTag 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Tag Data Printing"
   ClientHeight    =   9810
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   15210
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   9810
   ScaleMode       =   0  'User
   ScaleWidth      =   15210
   WindowState     =   2  'Maximized
   Begin MwfCtl.MWCTL_HLP MWCTL_HLP1 
      Height          =   405
      Left            =   10650
      TabIndex        =   83
      TabStop         =   0   'False
      Top             =   9300
      Visible         =   0   'False
      Width           =   555
      _ExtentX        =   979
      _ExtentY        =   714
   End
   Begin MwfCtl.MWCTL_BTN1 CmdTagFilePrint 
      Height          =   405
      Left            =   8160
      TabIndex        =   43
      TabStop         =   0   'False
      Top             =   9300
      Width           =   1395
      _ExtentX        =   0
      _ExtentY        =   0
      BackColor       =   -2147483633
      MaskColor       =   8454143
      Caption         =   "Print to &File"
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
   Begin MwfCtl.MWCTL_BTN1 CmdTagPrnPrint 
      Height          =   405
      Left            =   6765
      TabIndex        =   42
      TabStop         =   0   'False
      Top             =   9300
      Width           =   1395
      _ExtentX        =   0
      _ExtentY        =   0
      BackColor       =   -2147483633
      MaskColor       =   8454143
      Caption         =   "Print to &Printer"
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
   Begin VB.Frame FRA_KEYS 
      BorderStyle     =   0  'None
      Height          =   270
      Left            =   30
      TabIndex        =   46
      Top             =   30
      Width           =   5610
      Begin MwfCtl.MWCTL_MED atxt 
         Height          =   285
         Index           =   2
         Left            =   1905
         TabIndex        =   2
         ToolTipText     =   "Enter Invoice Voucher Year"
         Top             =   0
         Width           =   360
         _ExtentX        =   635
         _ExtentY        =   503
         MaxLength       =   2
         DataType        =   4
         DataField       =   "TgYy"
         IdName          =   "TGYY"
      End
      Begin MwfCtl.MWCTL_MED atxt 
         Height          =   285
         Index           =   4
         Left            =   3225
         TabIndex        =   4
         ToolTipText     =   "Enter Invoice Voucher No."
         Top             =   0
         Width           =   930
         _ExtentX        =   1640
         _ExtentY        =   503
         Alignment       =   2
         Mask            =   "#####0"
         MaxLength       =   6
         DataType        =   1
         DataField       =   "TgNo"
         IdName          =   "TGNO"
      End
      Begin MwfCtl.MWCTL_MED atxt 
         Height          =   285
         Index           =   1
         Left            =   1140
         TabIndex        =   1
         ToolTipText     =   "Enter Invoice Voucher Tc"
         Top             =   0
         Width           =   600
         _ExtentX        =   1058
         _ExtentY        =   503
         Enabled         =   0   'False
         MaxLength       =   3
         DataType        =   4
         DataField       =   "TgTc"
         ReCalcParent    =   "TGCHR"
         IdName          =   "TGTC"
      End
      Begin MwfCtl.MWCTL_MED atxt 
         Height          =   285
         Index           =   3
         Left            =   2445
         TabIndex        =   3
         ToolTipText     =   "Enter Invoice Voucher Character"
         Top             =   0
         Width           =   600
         _ExtentX        =   1058
         _ExtentY        =   503
         MaxLength       =   3
         DataType        =   4
         DataField       =   "TgChr"
         ReCalcOn        =   "TGTC"
         IdName          =   "TGCHR"
      End
      Begin MwfCtl.MWCTL_MED atxt 
         Height          =   285
         Index           =   5
         Left            =   4305
         TabIndex        =   5
         ToolTipText     =   "Enter Tag Serial"
         Top             =   0
         Width           =   930
         _ExtentX        =   1640
         _ExtentY        =   503
         Alignment       =   2
         Mask            =   "#####0"
         MaxLength       =   6
         DataType        =   1
         DataField       =   "TgSr"
         IdName          =   "TGSR"
      End
      Begin MwfCtl.MWCTL_MED atxt 
         Height          =   285
         Index           =   0
         Left            =   5250
         TabIndex        =   0
         Top             =   0
         Visible         =   0   'False
         Width           =   120
         _ExtentX        =   212
         _ExtentY        =   503
         Enabled         =   0   'False
         MaxLength       =   3
         DataType        =   4
         DataField       =   "TgCoCd"
         IdName          =   "TGCOCD"
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
         Left            =   4200
         TabIndex        =   52
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
         Left            =   3090
         TabIndex        =   51
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
         Left            =   2310
         TabIndex        =   50
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
         Index           =   38
         Left            =   1785
         TabIndex        =   49
         Top             =   0
         Width           =   105
      End
      Begin VB.Label ALBL 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Tag Serial"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Index           =   0
         Left            =   0
         TabIndex        =   47
         Top             =   0
         Width           =   990
      End
   End
   Begin MwfCtl.MWCTL_BTN1 CmdGenScp 
      Height          =   405
      Left            =   5370
      TabIndex        =   41
      TabStop         =   0   'False
      Top             =   9300
      Width           =   1395
      _ExtentX        =   0
      _ExtentY        =   0
      BackColor       =   -2147483633
      MaskColor       =   8454143
      Caption         =   "Ge&nerate"
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
      Left            =   60
      TabIndex        =   44
      Top             =   9300
      Width           =   15045
      _ExtentX        =   26538
      _ExtentY        =   979
      BtnWidth        =   885
      BtnHeight       =   435
      BtnMaskColor    =   8454143
      FontSz          =   8.25
   End
   Begin VB.Frame FRA_NKEYS 
      BorderStyle     =   0  'None
      Enabled         =   0   'False
      Height          =   9795
      Left            =   0
      TabIndex        =   45
      Top             =   0
      Width           =   15195
      Begin VB.Frame FraScp 
         Enabled         =   0   'False
         Height          =   735
         Left            =   4620
         TabIndex        =   68
         Top             =   8445
         Width           =   5805
         Begin MwfCtl.MWCTL_BTN1 CmdGenGo 
            Height          =   285
            Left            =   4950
            TabIndex        =   32
            Top             =   270
            Width           =   585
            _ExtentX        =   0
            _ExtentY        =   0
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
         Begin MwfCtl.MWCTL_MED TXTScp 
            Height          =   285
            Index           =   2
            Left            =   3480
            TabIndex        =   31
            ToolTipText     =   "Enter IdSr To For Generating Records"
            Top             =   270
            Width           =   1200
            _ExtentX        =   2117
            _ExtentY        =   503
            Alignment       =   2
            Mask            =   "######0"
            MaxLength       =   7
            DataType        =   1
            IdName          =   "WIDSRTO"
         End
         Begin MwfCtl.MWCTL_MED TXTScp 
            Height          =   285
            Index           =   1
            Left            =   1695
            TabIndex        =   30
            ToolTipText     =   "Enter IdSr From For Generating Records"
            Top             =   270
            Width           =   1200
            _ExtentX        =   2117
            _ExtentY        =   503
            Alignment       =   2
            Mask            =   "######0"
            MaxLength       =   7
            DataType        =   1
            IdName          =   "WIDSRFR"
         End
         Begin VB.Label ALBL 
            AutoSize        =   -1  'True
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
            Height          =   240
            Index           =   23
            Left            =   3105
            TabIndex        =   71
            Top             =   270
            Width           =   300
         End
         Begin VB.Label ALBL 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "From"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   240
            Index           =   24
            Left            =   1110
            TabIndex        =   70
            Top             =   270
            Width           =   540
         End
         Begin VB.Label ALBL 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "IdSr"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   240
            Index           =   16
            Left            =   300
            TabIndex        =   69
            Top             =   270
            Width           =   435
         End
      End
      Begin VB.Frame FraPrnFlds 
         Caption         =   "TXTFlds"
         Height          =   4185
         Left            =   6090
         TabIndex        =   58
         Top             =   1470
         Width           =   6735
         Begin MwfCtl.MWCTL_BTN1 CmdFldsOk 
            Height          =   315
            Left            =   5850
            TabIndex        =   29
            Top             =   3720
            Width           =   735
            _ExtentX        =   1296
            _ExtentY        =   556
            BackColor       =   -2147483633
            MaskColor       =   8454143
            Caption         =   "&Ok"
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
         Begin MwfCtl.MWCTL_MED TXTFlds 
            Height          =   285
            Index           =   9
            Left            =   990
            TabIndex        =   25
            ToolTipText     =   "Printing Field 01"
            Top             =   2370
            Width           =   5595
            _ExtentX        =   9869
            _ExtentY        =   503
            Enabled         =   0   'False
            MaxLength       =   255
            IdName          =   "WFLD09"
         End
         Begin MwfCtl.MWCTL_MED TXTFlds 
            Height          =   285
            Index           =   8
            Left            =   990
            TabIndex        =   24
            ToolTipText     =   "Printing Field 01"
            Top             =   2100
            Width           =   5595
            _ExtentX        =   9869
            _ExtentY        =   503
            Enabled         =   0   'False
            MaxLength       =   255
            IdName          =   "WFLD08"
         End
         Begin MwfCtl.MWCTL_MED TXTFlds 
            Height          =   285
            Index           =   7
            Left            =   990
            TabIndex        =   23
            ToolTipText     =   "Printing Field 01"
            Top             =   1830
            Width           =   5595
            _ExtentX        =   9869
            _ExtentY        =   503
            Enabled         =   0   'False
            MaxLength       =   255
            IdName          =   "WFLD07"
         End
         Begin MwfCtl.MWCTL_MED TXTFlds 
            Height          =   285
            Index           =   6
            Left            =   990
            TabIndex        =   22
            ToolTipText     =   "Printing Field 01"
            Top             =   1560
            Width           =   5595
            _ExtentX        =   9869
            _ExtentY        =   503
            Enabled         =   0   'False
            MaxLength       =   255
            IdName          =   "WFLD06"
         End
         Begin MwfCtl.MWCTL_MED TXTFlds 
            Height          =   285
            Index           =   5
            Left            =   990
            TabIndex        =   21
            ToolTipText     =   "Printing Field 01"
            Top             =   1290
            Width           =   5595
            _ExtentX        =   9869
            _ExtentY        =   503
            Enabled         =   0   'False
            MaxLength       =   255
            IdName          =   "WFLD05"
         End
         Begin MwfCtl.MWCTL_MED TXTFlds 
            Height          =   285
            Index           =   4
            Left            =   990
            TabIndex        =   20
            ToolTipText     =   "Printing Field 01"
            Top             =   1020
            Width           =   5595
            _ExtentX        =   9869
            _ExtentY        =   503
            Enabled         =   0   'False
            MaxLength       =   255
            IdName          =   "WFLD04"
         End
         Begin MwfCtl.MWCTL_MED TXTFlds 
            Height          =   285
            Index           =   3
            Left            =   990
            TabIndex        =   19
            ToolTipText     =   "Printing Field 01"
            Top             =   750
            Width           =   5595
            _ExtentX        =   9869
            _ExtentY        =   503
            Enabled         =   0   'False
            MaxLength       =   255
            IdName          =   "WFLD03"
         End
         Begin MwfCtl.MWCTL_MED TXTFlds 
            Height          =   285
            Index           =   2
            Left            =   990
            TabIndex        =   18
            ToolTipText     =   "Printing Field 01"
            Top             =   465
            Width           =   5595
            _ExtentX        =   9869
            _ExtentY        =   503
            Enabled         =   0   'False
            MaxLength       =   255
            IdName          =   "WFLD02"
         End
         Begin MwfCtl.MWCTL_MED TXTFlds 
            Height          =   285
            Index           =   1
            Left            =   990
            TabIndex        =   17
            ToolTipText     =   "Printing Field 01"
            Top             =   180
            Width           =   5595
            _ExtentX        =   9869
            _ExtentY        =   503
            Enabled         =   0   'False
            MaxLength       =   255
            IdName          =   "WFLD01"
         End
         Begin MwfCtl.MWCTL_MED TXTFlds 
            Height          =   285
            Index           =   10
            Left            =   990
            TabIndex        =   26
            ToolTipText     =   "Printing Field 10"
            Top             =   2650
            Width           =   5595
            _ExtentX        =   9869
            _ExtentY        =   503
            Enabled         =   0   'False
            MaxLength       =   255
            IdName          =   "WFLD10"
         End
         Begin MwfCtl.MWCTL_MED TXTFlds 
            Height          =   285
            Index           =   11
            Left            =   990
            TabIndex        =   27
            ToolTipText     =   "Printing Field 11"
            Top             =   2940
            Width           =   5595
            _ExtentX        =   9869
            _ExtentY        =   503
            Enabled         =   0   'False
            MaxLength       =   255
            IdName          =   "WFLD11"
         End
         Begin MwfCtl.MWCTL_MED TXTFlds 
            Height          =   285
            Index           =   12
            Left            =   990
            TabIndex        =   28
            ToolTipText     =   "Printing Field 12"
            Top             =   3230
            Width           =   5595
            _ExtentX        =   9869
            _ExtentY        =   503
            Enabled         =   0   'False
            MaxLength       =   255
            IdName          =   "WFLD12"
         End
         Begin VB.Label ALBL 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Field 12"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   240
            Index           =   32
            Left            =   60
            TabIndex        =   90
            Top             =   3270
            Width           =   750
         End
         Begin VB.Label ALBL 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Field 11"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   240
            Index           =   31
            Left            =   60
            TabIndex        =   89
            Top             =   2970
            Width           =   750
         End
         Begin VB.Label ALBL 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Field 10"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   240
            Index           =   30
            Left            =   60
            TabIndex        =   88
            Top             =   2670
            Width           =   750
         End
         Begin VB.Label ALBL 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Field 01"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   240
            Index           =   7
            Left            =   60
            TabIndex        =   67
            Top             =   180
            Width           =   750
         End
         Begin VB.Label ALBL 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Field 02"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   240
            Index           =   8
            Left            =   60
            TabIndex        =   66
            Top             =   465
            Width           =   750
         End
         Begin VB.Label ALBL 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Field 03"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   240
            Index           =   9
            Left            =   60
            TabIndex        =   65
            Top             =   750
            Width           =   750
         End
         Begin VB.Label ALBL 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Field 04"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   240
            Index           =   10
            Left            =   60
            TabIndex        =   64
            Top             =   1020
            Width           =   750
         End
         Begin VB.Label ALBL 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Field 05"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   240
            Index           =   11
            Left            =   60
            TabIndex        =   63
            Top             =   1290
            Width           =   750
         End
         Begin VB.Label ALBL 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Field 06"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   240
            Index           =   12
            Left            =   60
            TabIndex        =   62
            Top             =   1560
            Width           =   750
         End
         Begin VB.Label ALBL 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Field 07"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   240
            Index           =   13
            Left            =   60
            TabIndex        =   61
            Top             =   1830
            Width           =   750
         End
         Begin VB.Label ALBL 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Field 08"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   240
            Index           =   14
            Left            =   60
            TabIndex        =   60
            Top             =   2100
            Width           =   750
         End
         Begin VB.Label ALBL 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Field 09"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   240
            Index           =   15
            Left            =   60
            TabIndex        =   59
            Top             =   2370
            Width           =   750
         End
      End
      Begin VB.Frame FraNKeyAll 
         BorderStyle     =   0  'None
         Height          =   9300
         Left            =   30
         TabIndex        =   48
         Top             =   30
         Width           =   15120
         Begin VB.TextBox txtPwd 
            Height          =   285
            IMEMode         =   3  'DISABLE
            Left            =   10215
            MaxLength       =   30
            PasswordChar    =   "*"
            TabIndex        =   14
            Top             =   840
            Width           =   2535
         End
         Begin MwfCtl.MWCTL_MED atxt 
            Height          =   285
            Index           =   7
            Left            =   1140
            TabIndex        =   9
            ToolTipText     =   "Enter Tag Level"
            Top             =   555
            Width           =   345
            _ExtentX        =   609
            _ExtentY        =   503
            DataType        =   4
            DataField       =   "TgLevel"
            ReCalcParent    =   "TGAVGTYP"
            IdName          =   "TGLEVEL"
         End
         Begin MwfCtl.MWCTL_MED atxt 
            Height          =   285
            Index           =   6
            Left            =   1140
            TabIndex        =   7
            ToolTipText     =   "Enter Tag Format Code"
            Top             =   270
            Width           =   1005
            _ExtentX        =   1773
            _ExtentY        =   503
            MaxLength       =   5
            DataType        =   4
            DataField       =   "TgCd"
            IdName          =   "TGCD"
         End
         Begin MwfCtl.MWCTL_MED atxt 
            Height          =   285
            Index           =   8
            Left            =   4890
            TabIndex        =   10
            ToolTipText     =   "Enter Tag Average Type"
            Top             =   555
            Width           =   345
            _ExtentX        =   609
            _ExtentY        =   503
            DataType        =   4
            DataField       =   "TgAvgTyp"
            ReCalcParent    =   "TGWT"
            ReCalcOn        =   "TGLEVEL"
            IdName          =   "TGAVGTYP"
         End
         Begin MwfCtl.MWCTL_MED atxt 
            Height          =   285
            Index           =   9
            Left            =   1140
            TabIndex        =   12
            ToolTipText     =   "Enter Tag Weight"
            Top             =   840
            Width           =   345
            _ExtentX        =   609
            _ExtentY        =   503
            DataType        =   4
            DataField       =   "TgWt"
            ReCalcOn        =   "TGAVGTYP"
            IdName          =   "TGWT"
         End
         Begin MwfCtl.MWCTL_MED atxt 
            Height          =   285
            Index           =   10
            Left            =   4890
            TabIndex        =   13
            ToolTipText     =   "Enter Tag Price"
            Top             =   840
            Width           =   345
            _ExtentX        =   609
            _ExtentY        =   503
            DataType        =   4
            DataField       =   "TgPrice"
            IdName          =   "TGPRICE"
         End
         Begin MwfCtl.MWCTL_FLX GrdTagData 
            Height          =   7650
            Left            =   0
            TabIndex        =   16
            Top             =   1500
            Width           =   5985
            _ExtentX        =   10557
            _ExtentY        =   13494
            Cols            =   18
            colname1        =   "TGDSRNO"
            heading1        =   "Sr. No.    "
            datafld1        =   "TgdSrNo"
            datatype1       =   1
            mask1           =   "#####0"
            maxlength1      =   6
            style1          =   2
            tooltiptext1    =   "Serial Number"
            colname2        =   "TGDIDSR"
            heading2        =   "In Dsg Serial"
            datafld2        =   "TgdIdSr"
            datatype2       =   1
            mask2           =   "#####0"
            maxlength2      =   6
            style2          =   2
            tooltiptext2    =   "Invoic Design Serial"
            colname3        =   "TGDBAGNO"
            heading3        =   "Bag No.                    "
            datafld3        =   "TgdBagNo"
            recalcon3       =   "GrdTagData(TGDFLD02)"
            maxlength3      =   12
            style3          =   2
            tooltiptext3    =   "Bag No for which Record is Genrated"
            colname4        =   "TGDQTY"
            heading4        =   "Quantity  "
            datafld4        =   "TgdQty"
            datatype4       =   1
            mask4           =   "#####0"
            maxlength4      =   6
            tooltiptext4    =   "Quantity"
            colname5        =   "TGDFLD01"
            heading5        =   "Field01"
            datafld5        =   "TgdFld01"
            style5          =   0
            tooltiptext5    =   "Printing Field 01"
            colname6        =   "TGDFLD02"
            heading6        =   "Field02"
            datafld6        =   "TgdFld02"
            recalcparent6   =   "GrdTagData(TGDBAGNO)"
            style6          =   0
            tooltiptext6    =   "Printing Field 02"
            colname7        =   "TGDFLD03"
            heading7        =   "Field03"
            datafld7        =   "TgdFld03"
            style7          =   0
            tooltiptext7    =   "Printing Field 03"
            colname8        =   "TGDFLD04"
            heading8        =   "Field04"
            datafld8        =   "TgdFld04"
            style8          =   0
            tooltiptext8    =   "Printing Field 04"
            colname9        =   "TGDFLD05"
            heading9        =   "Field05"
            datafld9        =   "TgdFld05"
            style9          =   0
            tooltiptext9    =   "Printing Field 05"
            colname10       =   "TGDFLD06"
            heading10       =   "Field06"
            datafld10       =   "TgdFld06"
            style10         =   0
            tooltiptext10   =   "Printing Field 06"
            colname11       =   "TGDFLD07"
            heading11       =   "Field07"
            datafld11       =   "TgdFld07"
            style11         =   0
            tooltiptext11   =   "Printing Field 07"
            colname12       =   "TGDFLD08"
            heading12       =   "Field08"
            datafld12       =   "TgdFld08"
            style12         =   0
            tooltiptext12   =   "Printing Field 08"
            colname13       =   "TGDFLD09"
            heading13       =   "Field09"
            datafld13       =   "TgdFld09"
            style13         =   0
            tooltiptext13   =   "Printing Field 09"
            colname14       =   "WFLDS"
            colname15       =   "TGDFLD10"
            datafld15       =   "TgdFld10"
            maxlength15     =   255
            style15         =   0
            tooltiptext15   =   "Printing Field 10"
            colname16       =   "TGDFLD11"
            datafld16       =   "TgdFld11"
            maxlength16     =   255
            style16         =   0
            tooltiptext16   =   "Printing Field 11"
            colname17       =   "TGDFLD12"
            datafld17       =   "TgdFld12"
            style17         =   0
            tooltiptext17   =   "Printing Field 12"
         End
         Begin MSComDlg.CommonDialog CDOpen 
            Left            =   13500
            Top             =   90
            _ExtentX        =   847
            _ExtentY        =   847
            _Version        =   393216
         End
         Begin MwfCtl.MWCTL_MED atxt 
            Height          =   285
            Index           =   11
            Left            =   4530
            TabIndex        =   8
            ToolTipText     =   "Enter Tag Multiply By Factor"
            Top             =   270
            Width           =   705
            _ExtentX        =   1244
            _ExtentY        =   503
            Mask            =   "##0.000"
            MaxLength       =   7
            DataType        =   2
            DataField       =   "TgMulBy"
            ReCalcParent    =   "TGWT"
            ReCalcOn        =   "TGLEVEL"
            IdName          =   "TGMULBY"
         End
         Begin MwfCtl.MWCTL_MED atxt 
            Height          =   285
            Index           =   1000
            Left            =   1140
            TabIndex        =   15
            ToolTipText     =   "Enter Description"
            Top             =   1125
            Width           =   11610
            _ExtentX        =   20479
            _ExtentY        =   503
            MaxLength       =   45
            DataField       =   "TgDesc"
            IdName          =   "TGDESC"
         End
         Begin MwfCtl.MWCTL_MED atxt 
            Height          =   285
            Index           =   12
            Left            =   6780
            TabIndex        =   11
            ToolTipText     =   "Enter 'Y' to Specify Design Master as Source"
            Top             =   555
            Width           =   345
            _ExtentX        =   609
            _ExtentY        =   503
            DataType        =   4
            DataField       =   "TgSrcDsg"
            IdName          =   "TGSRCDSG"
         End
         Begin MwfCtl.MWCTL_MED atxt 
            Height          =   285
            Index           =   13
            Left            =   6780
            TabIndex        =   86
            ToolTipText     =   "Enter 'Y' to Specify Order Rm as Source"
            Top             =   840
            Visible         =   0   'False
            Width           =   345
            _ExtentX        =   609
            _ExtentY        =   503
            Enabled         =   0   'False
            DataType        =   4
            DataField       =   "TgSrcOrd"
            IdName          =   "TGSRCORD"
         End
         Begin MwfCtl.MWCTL_MED atxt 
            Height          =   285
            Index           =   30
            Left            =   6270
            TabIndex        =   6
            Top             =   0
            Width           =   855
            _ExtentX        =   1508
            _ExtentY        =   503
            MaxLength       =   8
            DataType        =   3
            DataField       =   "TgDt"
            IdName          =   "TGDT"
         End
         Begin VB.Label ALBL 
            BackStyle       =   0  'Transparent
            Caption         =   "Seo Pwd"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Index           =   55
            Left            =   9240
            TabIndex        =   92
            Top             =   840
            Width           =   975
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
            Index           =   33
            Left            =   5640
            TabIndex        =   91
            Top             =   0
            Width           =   615
         End
         Begin VB.Label ALBL 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Src Ord Rm"
            Enabled         =   0   'False
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   240
            Index           =   29
            Left            =   5640
            TabIndex        =   87
            Top             =   840
            Visible         =   0   'False
            Width           =   1095
         End
         Begin VB.Label ALBL 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Src DsgMst"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   240
            Index           =   28
            Left            =   5640
            TabIndex        =   85
            Top             =   555
            Width           =   1035
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
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
            Height          =   240
            Left            =   0
            TabIndex        =   84
            Top             =   1125
            Width           =   570
         End
         Begin VB.Label ALBL 
            AutoSize        =   -1  'True
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
            Height          =   240
            Index           =   27
            Left            =   3300
            TabIndex        =   82
            Top             =   270
            Width           =   1050
         End
         Begin VB.Label ALBL 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Tag Price"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   240
            Index           =   6
            Left            =   3750
            TabIndex        =   57
            Top             =   840
            Width           =   915
         End
         Begin VB.Label ALBL 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Tag Wt"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   240
            Index           =   5
            Left            =   0
            TabIndex        =   56
            Top             =   840
            Width           =   675
         End
         Begin VB.Label ALBL 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Tag Avg Type"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   240
            Index           =   4
            Left            =   3300
            TabIndex        =   55
            Top             =   555
            Width           =   1305
         End
         Begin VB.Label ALBL 
            BackStyle       =   0  'Transparent
            Caption         =   "Tag Cd"
            BeginProperty Font 
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
            Left            =   0
            TabIndex        =   54
            Top             =   270
            Width           =   855
         End
         Begin VB.Label ALBL 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Tag Level"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   240
            Index           =   2
            Left            =   0
            TabIndex        =   53
            Top             =   555
            Width           =   945
         End
      End
   End
   Begin VB.Frame FraPrn 
      Enabled         =   0   'False
      Height          =   1395
      Left            =   4620
      TabIndex        =   72
      Top             =   7790
      Width           =   5805
      Begin VB.ComboBox cboPrn 
         Height          =   315
         Left            =   1110
         Style           =   2  'Dropdown List
         TabIndex        =   35
         Top             =   300
         Width           =   3585
      End
      Begin MwfCtl.MWCTL_BTN1 CmdFileOpen 
         Height          =   285
         Left            =   4710
         TabIndex        =   78
         Top             =   300
         Width           =   855
         _ExtentX        =   1508
         _ExtentY        =   503
         BackColor       =   -2147483633
         MaskColor       =   8454143
         Caption         =   "&Browse"
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
      Begin MwfCtl.MWCTL_BTN1 CmdPrnGo 
         Height          =   285
         Left            =   4710
         TabIndex        =   40
         Top             =   930
         Width           =   855
         _ExtentX        =   1508
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
      Begin MwfCtl.MWCTL_MED TXTScp 
         Height          =   285
         Index           =   6
         Left            =   3480
         TabIndex        =   37
         ToolTipText     =   "Enter IdSr To For Printing Records"
         Top             =   630
         Width           =   1200
         _ExtentX        =   2117
         _ExtentY        =   503
         Alignment       =   2
         Mask            =   "######0"
         MaxLength       =   7
         DataType        =   1
         IdName          =   "WIDSRTOPRN"
      End
      Begin MwfCtl.MWCTL_MED TXTScp 
         Height          =   285
         Index           =   5
         Left            =   1695
         TabIndex        =   36
         ToolTipText     =   "Enter IdSr From For Printing Records"
         Top             =   630
         Width           =   1200
         _ExtentX        =   2117
         _ExtentY        =   503
         Alignment       =   2
         Mask            =   "######0"
         MaxLength       =   7
         DataType        =   1
         IdName          =   "WIDSRFRPRN"
      End
      Begin MwfCtl.MWCTL_MED TXTScp 
         Height          =   285
         Index           =   3
         Left            =   1695
         TabIndex        =   33
         ToolTipText     =   "Enter Serial Port No"
         Top             =   300
         Visible         =   0   'False
         Width           =   390
         _ExtentX        =   688
         _ExtentY        =   503
         Alignment       =   2
         Enabled         =   0   'False
         Mask            =   "0"
         DataType        =   1
         IdName          =   "WSRLPORT"
      End
      Begin MwfCtl.MWCTL_MED TXTScp 
         Height          =   285
         Index           =   4
         Left            =   1110
         TabIndex        =   34
         ToolTipText     =   "Printing Field 01"
         Top             =   300
         Width           =   3585
         _ExtentX        =   6324
         _ExtentY        =   503
         MaxLength       =   255
         IdName          =   "WFILENM"
      End
      Begin MwfCtl.MWCTL_MED TXTScp 
         Height          =   285
         Index           =   8
         Left            =   3480
         TabIndex        =   39
         ToolTipText     =   "Enter IdSr To For Printing Records"
         Top             =   930
         Width           =   1200
         _ExtentX        =   2117
         _ExtentY        =   503
         Alignment       =   2
         Mask            =   "######0"
         MaxLength       =   7
         DataType        =   1
         IdName          =   "WTGSRNOTO"
      End
      Begin MwfCtl.MWCTL_MED TXTScp 
         Height          =   285
         Index           =   7
         Left            =   1695
         TabIndex        =   38
         ToolTipText     =   "Enter IdSr From For Printing Records"
         Top             =   930
         Width           =   1200
         _ExtentX        =   2117
         _ExtentY        =   503
         Alignment       =   2
         Mask            =   "######0"
         MaxLength       =   7
         DataType        =   1
         IdName          =   "WTGSRNOFR"
      End
      Begin VB.Label ALBL 
         AutoSize        =   -1  'True
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
         Height          =   240
         Index           =   26
         Left            =   3105
         TabIndex        =   81
         Top             =   930
         Width           =   240
      End
      Begin VB.Label ALBL 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "From"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Index           =   25
         Left            =   1110
         TabIndex        =   80
         Top             =   930
         Width           =   495
      End
      Begin VB.Label ALBL 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Tg SrNo"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Index           =   22
         Left            =   180
         TabIndex        =   79
         Top             =   930
         Width           =   765
      End
      Begin VB.Label ALBL 
         AutoSize        =   -1  'True
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
         Height          =   240
         Index           =   21
         Left            =   180
         TabIndex        =   77
         Top             =   300
         Width           =   360
      End
      Begin VB.Label ALBL 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Printer"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Index           =   20
         Left            =   180
         TabIndex        =   76
         Top             =   300
         Width           =   645
      End
      Begin VB.Label ALBL 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "IdSr"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Index           =   19
         Left            =   180
         TabIndex        =   75
         Top             =   630
         Width           =   390
      End
      Begin VB.Label ALBL 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "From"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Index           =   18
         Left            =   1110
         TabIndex        =   74
         Top             =   630
         Width           =   495
      End
      Begin VB.Label ALBL 
         AutoSize        =   -1  'True
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
         Height          =   240
         Index           =   17
         Left            =   3105
         TabIndex        =   73
         Top             =   630
         Width           =   240
      End
   End
End
Attribute VB_Name = "EmrFrmTag"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'
'----   General Comments on the Entry Form  -------'
'
' ########################  Manoj #### Ver: 2.0.6 #### Date: 02/02/2004  ########################
' # Tag Printing                                                                                #
' # This Form Takes Invoce No. Genrrates Tags for Inv Designs or for Inv Bags                   #
' # Invoice No Along with TgSr is the Key Fields                                                #
' # Tags will be Generated for Invoce No in Key & Inv Serial Whiich are selected in Scope       #
' # Before Generating Tag Record 'X' must be Properly Saved                                     #
' # Depending on TgAvgTyp, Tag Records are generated for Inv Design Or Inv Bag                  #
' # For Bag Records InvDsg Values Are 0 Or Empty                                                #
' # For InvDsg Records Bag Values Are 0 Or Empty                                                #
' # Depending on TagLevel, Piece Values are Set by Piece or Bag or Inv Design Value             #
' # Weight & Values Can be Set As Invoice, Actual Or Customs                                    #
' # But if TgAvgTyp is Bag Then Weight is Always Actual                                         #
' # Using TgCd the Format String & Sring For Evaluation Are Taken                               #
' # While Generating Records First Read All Records of TagFmtSr & Get Strings For Evaluation    #
' # Get All Records for Tag Generation                                                          #
' # For each Record, Set Values for each TagField Defined & Evaluate String red from TagFmtSr   #
' # The Records with Evaluated Tag Fields are Shown to user, these fields are editable.         #
' # The Generated Record Can be Print Directly On Printer Or To File                            #
' ###########################################  Manoj  ###########################################
Option Explicit

' ########################  Manoj #### Ver: 2.0.6 #### Date: 02/02/2004  ########################
' # These Win Api Typpe & Functions are used to print Formated TagData on Printer               #
' ###########################################  Manoj  ###########################################
#If Win32 Then
  Private Type DOC_INFO_1
    pDocName As String
    pOutputFile As String
    pDatatype As String
  End Type
#End If 'WIN32 Types
'** Function Declarations:
#If Win32 Then
  Private Declare Function OpenPrinter& Lib "winspool.drv" Alias "OpenPrinterA" (ByVal pPrinterName As String, phPrinter As Long, ByVal pDefault As Long)    ' Third param changed to long
  Private Declare Function StartDocPrinter& Lib "winspool.drv" Alias "StartDocPrinterA" (ByVal hPrinter As Long, ByVal Level As Long, pDocInfo As DOC_INFO_1)
  Private Declare Function StartPagePrinter& Lib "winspool.drv" (ByVal hPrinter As Long)
  Private Declare Function WritePrinter& Lib "winspool.drv" (ByVal hPrinter As Long, pBuf As Any, ByVal cdBuf As Long, pcWritten As Long)
  Private Declare Function EndDocPrinter& Lib "winspool.drv" (ByVal hPrinter As Long)
  Private Declare Function EndPagePrinter& Lib "winspool.drv" (ByVal hPrinter As Long)
  Private Declare Function ClosePrinter& Lib "winspool.drv" (ByVal hPrinter As Long)
#End If 'WIN32

' ########################  Manoj #### Ver: 2.0.6 #### Date: 02/02/2004  ########################
' # DOWNLOADED FROM SITE www.datamaxcorp.com (Programmers Manual) #
' ###########################################  Manoj  ###########################################
Dim SelPrinter As String        ' Selected Printer Name

Dim moCn As MwfLib.MDOConnection   '*** (Bef speed) bef 24/11/06

' variables for holding key values
Dim ms_TgCoCd As String, ms_TcTyp As String
Dim ms_TgTc As String, ms_TgYy As String, ms_TgChr As String
Dim ms_TgNo As Double, ms_TgSr As Double

Dim scTag As Object                       ' Scripting Object
Dim mColTagFlds As New Collection         ' Tag Fields Collection
Dim mRs_Rep As New MwfLib.MDORowSet       ' Report RecordSet
Dim mRs_FmtStr As New MwfLib.MDORowSet    ' Format String Recordset
Dim OpenFilePath  As String               ' File Name & Path info
Dim mPrntOpt As String                    ' Printing Option
Dim mRunSrOld As Double, mRunSrNew As Double  '578-32
Dim ms_PwdFlag As Boolean                  ' 4.1.2

'*** For the Toggle Options of Command Buttons
Enum en_TagFra
  GenScp = 0
  PrnScp = 1
End Enum
 
Private Sub adc_AutoGen()
  'Auto Generation of TgSr Field
  If adc("TgSr") = 0 Then
    adc("TgSr") = moCn.AutoGen("Tag", "TgSr", adc("TgCoCd") + adc("TgTc") + adc("TgYy") + adc("TgChr") + CStr(adc("TgNo")))
  Else
    Call moCn.AutoGen("Tag", "TgSr", adc("TgCoCd") + adc("TgTc") + adc("TgYy") + adc("TgChr") + CStr(adc("TgNo")), adc("TgSr"))
  End If
End Sub

Private Sub ADC_KeyValid(Cancel As Boolean, ErrMsg As String)
' Checking Validation of Key Fields
  If adc("TgYy") = "" Then Cancel = True: ErrMsg = "Enter Tag Year": Exit Sub
  If adc("TgChr") = "" Then Cancel = True: ErrMsg = "Enter Tag Character": Exit Sub
  If adc("TgNo") = "" Then Cancel = True: ErrMsg = "Enter Tag No": Exit Sub
End Sub

Private Sub ADC_Load()
  '*** Add Printers to CboPrinter
  '*** Set the First Key Control to TgChr
  '*** Set the First Non Key Control to TgCd
  '*** Set the Previous and Next Control Properties for ADC
  Dim ind As Integer
  Dim Prnt As Printer
  
' Finding All Printers Installed on Machine
  For Each Prnt In Printers
      cboPrn.AddItem Prnt.DeviceName
  Next

' Setting Default Printer Selected

  If cboPrn.ListCount > 0 Then
      For ind = 0 To cboPrn.ListCount - 1
          If Printer.DeviceName = cboPrn.List(ind) Then
              cboPrn.ListIndex = ind
              Exit For
          End If
      Next ind
  End If
  
  Set moCn = adc.Connection    '*** (Bef speed) bef 24/11/06
  Set adc.FirKeyCtl = adc("TgChr")
  Set adc.FirNKeyCtl = adc("TgDt")      '4.1.2
  'Set ADC.FirNKeyCtl = ADC("TgCd")
  
  '*** Set Grid Control as Child Control for ADC
  ' Set Previous & Next COntrols of Grid
  adc.Child = "GrdTagData"
  Set GrdTagData.PrevCtl = adc("TgPrice")
  Set GrdTagData.NextCtl = adc
  GrdTagData.HotKey = "A"
  
  ' Set key Variables
  ms_TgCoCd = gs_CoCd
  
  'VK.25 - jEmr TC added
  ms_TcTyp = IIF(UCase(adc.MenuCd) = UCase("RepJtTagPrn"), "JT", ctInTcTyp)
  
  ms_TgTc = moCn.GetFldVal("Select PValue from Param where PTyp='TCTYP' and PMCd='" + ms_TcTyp + "' ")
  ms_TgYy = moCn.GetFldVal("Select hYy from Head where HCoCd= '" + ms_TgCoCd + "' and HCd='" + ctSelfCmCd + "'")

  ' Setting Scripting Control
  Set scTag = CreateObject("MSScriptControl.ScriptControl")
  scTag.Language = "VBScript"
  scTag.Timeout = -1
  SetScriptCtl scTag, mColTagFlds
  
  ' Seting Default Printing Option
  mPrntOpt = "P"
  
  'VK.25 - Enable Tc & Disable Level, AvgTyp, Wt and Price for jEmr
  If UCase(adc.MenuCd) = UCase("RepJtTagPrn") Then
    adc("TgTc").Enabled = True
    adc("TgAvgTyp").Enabled = False
    adc("TgWt").Enabled = False
    adc("TgPrice").Enabled = False
    Set adc.FirKeyCtl = adc("TgTc")
  End If
  
End Sub
Private Sub ADC_KeyWhen()
  'Disable All Buttons
  'Hide All frames except FraNKeyAll
  Call EnaDisaCmds(True)
  Call HideAllFras
  
  ' Set key Fields As Previos Record Values
  If ms_TgCoCd <> "" Then adc("TgCoCd") = ms_TgCoCd
  If ms_TgTc <> "" Then adc("TgTc") = ms_TgTc
  If ms_TgYy <> "" Then adc("TgYy") = ms_TgYy
  If ms_TgChr <> "" Then adc("TgChr") = ms_TgChr
  If ms_TgNo <> 0 Then adc("TgNo") = ms_TgNo
  
  '****** Sachin 2.14.0 - [24-11-2006] - [Multi-Company] ******
  Me.Caption = GetFrmCaption(gs_CoCd, "Tag Data Printing")
  '****** Sachin 2.14.0 - [24-11-2006] - [Multi-Company] ******
    
  TxtPwd.text = ""      '4.1.2
  
End Sub
Private Sub ADC_NonKeyWhen(Cancel As Boolean, ErrMsg As String)
  '*** Enable all the command buttons
       If adc.IsNew And Not moCn.RecSeek("select 1 from vParam where " + _
                         "vPCoCd= '" + ms_TgCoCd + "' and vPTyp= 'CHR' and " + _
                         "vPMCd = '" + ms_TgTc + "' and vPSCd='" + ms_TgChr + "' and vPValidYn<>'N'") Then
       Cancel = True: ErrMsg = "Voucher Character No More Valid": adc.AllowSave = False: Exit Sub
      Else
        adc.AllowSave = True
      End If
  
  Call EnaDisaCmds(False)
  
  ' Set Key Variables
  ms_TgTc = adc("TgTc")
  ms_TgYy = adc("TgYy")
  ms_TgChr = adc("TgChr")
  ms_TgNo = adc("TgNo")
  ms_TgSr = adc("TgSr")
  
  TxtPwd.text = ""  '4.1.2
  
  ' If Grid Having data set Row to First Row
  If GrdTagData.Rows > 1 Then GrdTagData.Row = 1
  If adc.AddFndMode = xaddmode Then
    adc("TgMulBy") = 1
    adc("TgSrcDsg") = "N"     ' ***** Manali 3.10.0 - 03/04/12 - Copy Wts From DsgMst
    adc("TgSrcOrd") = "N"     ' ***** Manali 3.10.0 - 09/05/12 - Copy Wts From Ord
    
    adc("TgDt") = moCn.SrvrDate  '4.1.2
    
    'VK.25 - Set default value to "I" for jEmr
    If UCase(adc.MenuCd) = UCase("RepJtTagPrn") Then
      adc("TgLevel") = "I"
      adc("TgAvgTyp") = "I"
      adc("TgWt") = "I"
      adc("TgPrice") = "I"
    End If
  End If
  
  '4.1.2
  ms_PwdFlag = moCn.RecSeek("Select PValue1 from Param where PTyp= 'SEOPWD' and PmCd= 'SEOPWD' and PValue1= '' ")
  EnaDisaTagFlds (ms_PwdFlag)
   
End Sub
Private Sub ADC_NonKeyValid(pr_Cancel As Boolean, pr_ErrMsg As String)
  '*** Set the current tab of the tabpage as 0
End Sub

Private Sub ATXT_FldWhen(Index As Integer, ByVal IdName As String, Cancel As Boolean, ErrMsg As String)
  ' If Record Exists in Grid Then No Value in 'X' level can be Edited
  ' Zubin (06/12/05)
  If UCase(IdName) <> UCase("TgDesc") Then
  ' Zubin (06/12/05)
    If moCn.RecSeek("Select 1 From TagData Where TgdCoCd = '" + ms_TgCoCd + "'" + _
                    " And TgdTc = '" + ms_TgTc + "' And TgdYy = '" + ms_TgYy + "'" + _
                    " And TgdChr = '" + ms_TgChr + "' And TgdNo = " + CStr(ms_TgNo) + _
                    " And TgdSr = " + CStr(ms_TgSr)) = True And _
                    IdName <> UCase$("TgYy") And IdName <> UCase$("TgChr") And _
                    IdName <> UCase$("TgNo") And IdName <> UCase$("TgSr") Then
      ErrMsg = "Cannot change value while in Tag Data Record Exists"
      Cancel = True: Exit Sub
    End If
  End If
  
  '*** Set Helps For Each Field
  Select Case UCase(IdName)
  'VK.25 - TC help added
  Case Is = UCase("TgTc")
    Call HlpList.PMCd("TC", "'" + ms_TcTyp + "' ")
    
  Case Is = UCase("TgYy")
    Call HlpList.PMCd("YY")
  Case Is = UCase("TgChr")
' **** Zubin 211 **** '
    'Call HlpList.vPSCd(ms_TgCoCd, "CHR", ms_TgTc)
    Call HlpList.vPSCd(ms_TgCoCd, "CHR", ms_TgTc, "Y")
' **** Zubin 211 **** '
  Case Is = UCase("TgNo")
    'VK.25 - OmNo help added for jEmr
    If UCase(adc.MenuCd) = UCase("RepJtTagPrn") Then
      Call HlpList.OmNo(ms_TgCoCd, ms_TgTc, ms_TgYy, ms_TgChr)
    Else
      Call HlpList.InNo(ms_TgCoCd, ms_TgTc, ms_TgYy, ms_TgChr)
    End If
  Case Is = UCase("TgSr")
    Call HlpList.TgSr(ms_TgCoCd, ms_TgTc, ms_TgYy, ms_TgChr, ms_TgNo)
    
  '4.1.2
  Case Is = UCase("TgDt")
    If GrdTagData.Rows - 1 > 0 Then Cancel = True: ErrMsg = "Cannot Edit Date": Exit Sub
    
  Case Is = UCase("TgCd")
    Call HlpList.TgfCd(ms_TgCoCd)
  Case Is = UCase("TgLevel")
    Call HlpList.PMCd("TAGLEVEL")
  Case Is = UCase("TgAvgTyp")
    Cancel = (adc("TgLevel") <> "P")
    If Cancel = True Then ErrMsg = "Cannot Edit This Value": Exit Sub
    Call HlpList.PMCd("TGAVGTYP")
  Case Is = UCase("TgWt")
    Cancel = (adc("TGAVGTYP") = "B")
    If Cancel = True Then ErrMsg = "Cannot Edit This Value": Exit Sub
    Call HlpList.PMCd("INVAL")
  Case Is = UCase("TgPrice")
    Call HlpList.PMCd("INVAL")
  
  ' ***** Manali 3.10.0 - 03/04/12 - Copy Wts From DsgMst
  Case Is = UCase("TgSrcDsg")
    Call HlpList.PMCd("YN")
  End Select
End Sub

Private Sub cboPrn_Click()
  ' Set printer as system default.
  SelPrinter = cboPrn.text
End Sub

Private Sub TxtPwd_Validate(Cancel As Boolean)
  'Sub added in 4.1.2
  If Trim(TxtPwd.text) <> "" Then
      If Trim(TxtPwd.text) = moCn.GetFldVal("select IsNull ((select PValue from Param Join UsrAccess on UaUsrCd = PMCd " + _
                   " Where PTyp= 'USR' and PMCd= '" + gs_UsrCd + "' and UaMnuCd = '" + adc.MenuCd + "' and " + _
                   " UaSeoYn= 'Y' and PValue= '" + Trim(TxtPwd.text) + "'), " + _
                   " (Select PValue1 from Param where Ptyp= 'SEOPWD' and PmCd= 'SEOPWD' )) ") Then
          ms_PwdFlag = True
          'EnaDisaTagFlds (True)
      Else
          ms_PwdFlag = False: Cancel = True: DispMsg "Invalid Edit Password / User Does Not Have SEO Rights ", etError
          'EnaDisaTagFlds (True)
      End If
    Else
      ms_PwdFlag = moCn.RecSeek("Select PValue1 from Param where PTyp= 'SEOPWD' and PmCd= 'SEOPWD' and PValue1= '' ")
      'If ms_PwdFlag = True Then ADC("TSeoPwd") = txtPwd.text
    End If
    EnaDisaTagFlds (ms_PwdFlag)
End Sub

Private Sub EnaDisaTagFlds(pVal As Boolean)
    '4.1.2
    adc("WFLD01").Enabled = pVal
    adc("WFLD02").Enabled = pVal
    adc("WFLD03").Enabled = pVal
    adc("WFLD04").Enabled = pVal
    adc("WFLD05").Enabled = pVal
    adc("WFLD06").Enabled = pVal
    adc("WFLD07").Enabled = pVal
    adc("WFLD08").Enabled = pVal
    adc("WFLD09").Enabled = pVal
    adc("WFLD10").Enabled = pVal
    adc("WFLD11").Enabled = pVal
    adc("WFLD12").Enabled = pVal
End Sub


Private Sub TXTScp_FldWhen(Index As Integer, ByVal IdName As String, Cancel As Boolean, ErrMsg As String)
' Setting Help For Scope Fields
  Select Case UCase(IdName)
    Case Is = UCase("wIdSrFr"), UCase("wIdSrTo"), UCase("wIdSrFrPrn"), UCase("wIdSrToPrn")
      'VK.25 - OdSr help added for jEmr
      If UCase(adc.MenuCd) = UCase("RepJtTagPrn") Then
        Call HlpList.OdSr(ms_TgCoCd, ms_TgTc, ms_TgYy, ms_TgChr, ms_TgNo)
      Else
        Call HlpList.InSr(ms_TgCoCd, ms_TgTc, ms_TgYy, ms_TgChr, ms_TgNo)
      End If
    Case Is = UCase("wTgSrNoFr")
      Call HlpList.TgdSrNo(ms_TgCoCd, ms_TgTc, ms_TgYy, ms_TgChr, ms_TgNo, ms_TgSr)
    Case Is = UCase("wTgSrNoTo")
      Call HlpList.TgdSrNo(ms_TgCoCd, ms_TgTc, ms_TgYy, ms_TgChr, ms_TgNo, ms_TgSr)
  End Select
End Sub
Private Sub atxt_ReCalculate(Index As Integer, ByVal IdName As String, ByVal OldVal As Variant, ByVal RowNum As Integer, ByVal SourceName As String, ByVal SourceColName As String)
  ' Seting ReCalculations
  Select Case UCase(IdName)
  Case Is = UCase("TgAvgTyp")
      If adc("TgLevel") <> "P" Then adc(IdName) = adc("TgLevel")
  Case Is = UCase("TgWt")
      If adc("TgAvgTyp") = "B" Then adc(IdName) = "A"
  ' Zubin
  Case Is = UCase("TgChr")
    If adc.Mode = xNorm Then _
       adc("TgChr") = moCn.GetFldVal("Select vPValue from vParam where " + _
                                     "vPCoCd= '" + adc("TgCoCd") + "' and vPTyp= 'DEFCHR' " + _
                                     "and vPMCd= '" + adc("TgTc") + "' and vPSCd= ''")
  End Select
End Sub
Private Sub ATXT_FldChange(Index As Integer, ByVal IdName As String, Cancel As Boolean, ErrMsg As String, ByVal pv_NewValue As Variant)
  ' Checking Field values for Validity
  Select Case UCase(IdName)
  Case Is = UCase("TgTc")
    Cancel = Not moCn.RecSeek("select 1 from Param where PTyp= 'TC' and " + _
                         "PMCd = '" + pv_NewValue + "' and " + _
                         "PValue= '" + ms_TcTyp + "'")
    If Cancel = True Then ErrMsg = "Invalid Invoice Voucher Tc": Exit Sub
    ms_TgTc = pv_NewValue
  Case Is = UCase("TgYy")
    Cancel = Not moCn.RecSeek("select 1 from Param where PTyp= 'YY' and " + _
                         "PMCd = '" + pv_NewValue + "'")
    If Cancel = True Then ErrMsg = "Invalid Invoice Voucher Year": Exit Sub
    ms_TgYy = pv_NewValue
  Case Is = UCase("TgChr")
    Cancel = pv_NewValue <> "" And (Not moCn.RecSeek("select 1 from vParam where " + _
                         "vPCoCd= '" + ms_TgCoCd + "' and vPTyp= 'CHR' and " + _
                         "vPMCd = '" + ms_TgTc + "' and vPSCd='" + pv_NewValue + "'"))
    If Cancel = True Then ErrMsg = "Invalid Invoice Voucher Character": Exit Sub
' **** Zubin 211 **** '
    Dim ws_UsrArr() As String, wi_Index As Integer
    ws_UsrArr = Split(moCn.GetFldVal("Select vPValue4 from vParam where vPCoCd= '" + ms_TgCoCd + "' and " + _
                                        "vPTyp='CHR' and vPMCd='" + ms_TgTc + "' and vPSCd= '" + pv_NewValue + "'"), ",")
    If IsArray(ws_UsrArr) Then
      For wi_Index = 0 To UBound(ws_UsrArr)
        Cancel = IIF(UCase(gs_UsrCd) = UCase(ws_UsrArr(wi_Index)), False, True)
        If Cancel = False Then: Exit For
      Next wi_Index
    End If
   If Cancel = True Then ErrMsg = "Access To This Invoice Voucher Character Is Denied To User " + gs_UsrCd: Exit Sub
' **** Zubin 211 **** '
    ms_TgChr = pv_NewValue
  Case Is = UCase("TgNo")
    'VK.25 - Changed Order related fields instead of Invoice for jEmr
    If UCase(adc.MenuCd) = UCase("RepJtTagPrn") Then
      Cancel = pv_NewValue <> "0" And (Not moCn.RecSeek("select 1 from OrdMst " + _
                          " Where " + IIF(gs_Partition = ctCurrPrtn, " OmPrtKey='" + ctCurrPrtn + "' and ", "") + " OmCoCd ='" + ms_TgCoCd + "' And OmTc ='" + ms_TgTc + "'" + _
                          " And OmYy ='" + ms_TgYy + "' And OmChr ='" + ms_TgChr + "'" + _
                          " And OmNo =" + CStr(pv_NewValue)))
      If Cancel = True Then ErrMsg = "Invalid Order Voucher Number": Exit Sub
    Else
      Cancel = pv_NewValue <> "0" And (Not moCn.RecSeek("select 1 from InvHd " + _
                          " Where " + IIF(gs_Partition = ctCurrPrtn, " InPrtKey='" + ctCurrPrtn + "' and ", "") + " InCoCd ='" + ms_TgCoCd + "' And InTc ='" + ms_TgTc + "'" + _
                          " And InYy ='" + ms_TgYy + "' And InChr ='" + ms_TgChr + "'" + _
                          " And InNo =" + CStr(pv_NewValue)))
      If Cancel = True Then ErrMsg = "Invalid Invoice Voucher Number": Exit Sub
    End If
    ms_TgNo = pv_NewValue
    
  '4.1.2
  Case Is = UCase("TgDt")
    If pv_NewValue = "" Then Cancel = True: ErrMsg = "Tag Date Cannot Be Blank/Invalid": Exit Sub
    
  Case Is = UCase("TgCd")
    Cancel = Not moCn.RecSeek("select 1 from TagFmt where TgfCoCd= '" + ms_TgCoCd + "' And TgfCd= '" + pv_NewValue + "'")
    If Cancel = True Then ErrMsg = "Invalid Tag Format Code": Exit Sub
  Case Is = UCase("TgLevel")
    Cancel = Not moCn.RecSeek("select PMCd from Param where PTyp= 'TAGLEVEL' and " + _
                         "PMCd = '" + pv_NewValue + "'")
    'VK.25 - Tag Level 'B' not applicable in jEmr
    If UCase(adc.MenuCd) = UCase("RepJtTagPrn") Then
      If UCase(pv_NewValue) = "B" Then
        Cancel = True
      End If
    End If
    If Cancel = True Then ErrMsg = "Invalid Tag Level Option": Exit Sub
  Case Is = UCase("TgAvgTyp")
    Cancel = Not moCn.RecSeek("select PMCd from Param where PTyp= 'TGAVGTYP' and " + _
                         "PMCd = '" + pv_NewValue + "'")
    If Cancel = True Then ErrMsg = "Invalid Tag Average Type Option": Exit Sub
  Case Is = UCase("TgWt")
    Cancel = Not moCn.RecSeek("select PMCd from Param where PTyp= 'INVAL' and " + _
                         "PMCd = '" + pv_NewValue + "' AND '" + pv_NewValue + "' <> 'S'")
    If Cancel = True Then ErrMsg = "Invalid Tag Weight Option": Exit Sub
  Case Is = UCase("TgPrice")
    Cancel = Not moCn.RecSeek("select 1 from Param where PTyp= 'INVAL' and " + _
                         "PMCd = '" + pv_NewValue + "' AND '" + pv_NewValue + "' <> 'S'")
    If Cancel = True Then ErrMsg = "Invalid Tag Price Option": Exit Sub
    
  Case Is = UCase("TgSrcDsg")
    Cancel = Not moCn.RecSeek("select PMCd from Param where PTyp= 'YN' and " + _
                         "PMCd = '" + pv_NewValue + "' AND '" + pv_NewValue + "' <> 'S'")
    If Cancel = True Then ErrMsg = "Enter 'Y' to Specify Design Master as Source Else 'N'": Exit Sub
    
  End Select
End Sub

Private Sub CmdFileOpen_Click()
  ' When Button click dialog box is Open for Opening File for saving Tag Data
  Dim wsFile As String
  If adc("wFileNm") <> "" Then CDOpen.FileName = adc("wFileNm")
  CDOpen.DialogTitle = "Save Text File For Tag"
  CDOpen.InitDir = App.Path
  CDOpen.Filter = "Text Files (*.txt)"
  CDOpen.FilterIndex = 1
  CDOpen.Flags = cdlOFNLongNames Or cdlOFNPathMustExist Or cdlOFNExtensionDifferent
  CDOpen.DefaultExt = "txt"
  CDOpen.ShowOpen
  If CDOpen.FileName <> "" Then adc("wFileNm") = CDOpen.FileName
  CDOpen.FileName = ""
End Sub

Private Sub CmdPrnGo_Click()
' If PrntOpt = "P" Then Check serial No<> 0
' IF PrntOpt = "F" Then Check File Name for validity
' Call Print Report
' Call Display Frame (Print SCope)
On Error GoTo ErrHndlr
  ' If Print To File is selected Then
  If UCase$(mPrntOpt) = UCase$("F") Then
    If Len(adc("wFileNm")) > 4 Then
    ' If FileName is given then Check for Path & FileName
      Open adc("wFileNm") For Output Access Write As #1
      Close #1
      OpenFilePath = adc("wFileNm")         'Set Path of File
    End If
  End If
  ' Select Records To Print with given Scope.
  If GetRepRec = True Then
    'If Records Exists Then Print Tags
    Call PrnRep
    'Hide Print Scope Frame
    Call DispFra(PrnScp)
  End If
  Exit Sub
ErrHndlr:
  If UCase$(mPrntOpt) = UCase$("F") Then
    OpenFilePath = ""
    DispMsg "File Name Or Path Is Wrong", etError: Exit Sub
  End If
End Sub
Private Sub CmdTagFilePrint_Click()
' Set PrntOpt = "F" i.e. Print To File
' Call diplay Frame (Print Scope)
' If frame is Visible Then Call Browse Buttton
  If Not adc.SaveRec Then
    DispMsg "X Save Failed", etInfo
    Exit Sub
  End If
  
  mPrntOpt = "F"
  Call DispFra(PrnScp)
  If FraPrn.Visible = True Then
    Call CmdFileOpen_Click
  End If
End Sub
Private Sub CmdTagPrnPrint_Click()
' Set PrntOpt = "P" i.e. Print On Printer
' Call diplay Frame (Print Scope)
  If Not adc.SaveRec Then
    DispMsg "X Save Failed", etInfo
    Exit Sub
  End If
  
  mPrntOpt = "P"
  Call DispFra(PrnScp)
End Sub

Private Sub GrdTagData_ReCalc(ByVal RowNum As Integer, ByVal ColName As String, ByVal OldVal As Variant, ByVal SourceName As String, ByVal SourceColName As String)
  Select Case UCase$(ColName)
' Zubin (Before 06/12/05) (Recalc on TgdBagNo Not required as value is now stored)
'    'Case UCase$("wBagNo")
'    Case UCase$("TgdBagNo")
'      If (GrdTagData.Mode = fgmnorm Or GrdTagData.Mode = fgminit) And _
'         adc.Mode = xNorm And UCase$(adc("TgAvgTyp")) = "B" Then
'
'         'ADC.AddFndMode= xFndMode And
'        GrdTagData.Value(RowNum, ColName) = moCn.GetFldVal("Select IfFdBYy + '/' + IfFdBChr + '/' + Cast(IfFdBNo As VarChar(6)) As BagNo " + _
'          " from InvDsg Inner Join InvFgd On IdCoCd = IfCoCd And IdTc = IfTc And IdYy = IfYy And" + _
'          " IdChr = IfChr And IdNo = IfNo And IdSr = IfSr Where " + _
'          " IdCoCd = '" + ms_TgCoCd + "' And IdTc = '" + ms_TgTc + "'" + _
'          " And IdYy = '" + ms_TgYy + "' And IdChr = '" + ms_TgChr + "'" + _
'          " And IdNo = " + CStr(ms_TgNo) + " And IdSr = " + CStr(GrdTagData.Value(RowNum, "TgdIdSr")))
'      End If
  End Select
End Sub

Private Sub GrdTagData_Valid()
  ' If Grid doesn't have any row then Clear all Textboxes of Tag Fields
  If GrdTagData.Rows <= 1 Then
    adc("wFLD01") = "": adc("wFLD02") = "": adc("wFLD03") = "": adc("wFLD04") = "": adc("wFLD05") = ""
    adc("wFLD06") = "": adc("wFLD07") = "": adc("wFLD08") = "": adc("wFLD09") = "": adc("wFLD10") = ""
    adc("wFLD11") = "": adc("wFLD12") = ""
  End If
End Sub

Private Sub TXTFlds_FldWhen(Index As Integer, ByVal IdName As String, Cancel As Boolean, ErrMsg As String)
  ' If No record present for Current Key Valus Then Tag Field are not Editable
  If moCn.RecSeek("Select 1 From TagData Where TgdCoCd = '" + ms_TgCoCd + "'" + _
                  " And TgdTc = '" + ms_TgTc + "' And TgdYy = '" + ms_TgYy + "'" + _
                  " And TgdChr = '" + ms_TgChr + "' And TgdNo = " + CStr(ms_TgNo) + _
                  " And TgdSr = " + CStr(ms_TgSr)) = False Then
    Cancel = True
    ErrMsg = "No Record Present for this Tag Serial"
    Exit Sub
  End If
End Sub
Private Sub CmdFldsOk_Click()
  ' Set Changed Values of TextBoxes to Grid Columns
  Dim RowNum As Integer
  FraNKeyAll.Enabled = True
  With GrdTagData
    If .Row > 0 Then
      .SetFocus
      .Col = 1
      RowNum = .Row
      .Value(RowNum, "TgdFld01") = adc("wFLD01")
      .Value(RowNum, "TgdFld02") = adc("wFLD02")
      .Value(RowNum, "TgdFld03") = adc("wFLD03")
      .Value(RowNum, "TgdFld04") = adc("wFLD04")
      .Value(RowNum, "TgdFld05") = adc("wFLD05")
      .Value(RowNum, "TgdFld06") = adc("wFLD06")
      .Value(RowNum, "TgdFld07") = adc("wFLD07")
      .Value(RowNum, "TgdFld08") = adc("wFLD08")
      .Value(RowNum, "TgdFld09") = adc("wFLD09")
      .Value(RowNum, "TgdFld10") = adc("wFLD10")
      .Value(RowNum, "TgdFld11") = adc("wFLD11")
      .Value(RowNum, "TgdFld12") = adc("wFLD12")
    Else
      adc.SetFocus
    End If
  End With
  Call EnaDisaCmds(False)
  adc.AllowSave = True
End Sub

Private Sub CmdGenGo_Click()
'  Save 'X' Record
'  Generate SQL to Find Records
'  Close Scope Frame
  Dim wsSql As String, wsFlds As String, wsJoin As String, wsQty As String, wsScpCnd As String
  Dim wsFld01 As String, wsFld02 As String, wsFld03 As String, wsFld04 As String, wsFld05 As String
  Dim wsFld06 As String, wsFld07 As String, wsFld08 As String, wsFld09 As String, wsFld As String
  Dim wsFld10 As String, wsFld11 As String, wsFld12 As String
  Dim wiSrNo As Double, wRow As Integer
  Dim moRsSr As New MwfLib.MDORowSet, moRsFlds As New MwfLib.MDORowSet

  ' Save ADC Record
  adc.SaveRec
  
  ' Zubin (07/12/05) Not Required
'  'Check Already Record Generated
'  If moCn.RecSeek("Select 1 From TagData Where TgdCoCd = '" + ms_TgCoCd + "'" + _
'                " And TgdTc = '" + ms_TgTc + "' And TgdYy = '" + ms_TgYy + "'" + _
'                " And TgdChr = '" + ms_TgChr + "' And TgdNo = " + CStr(ms_TgNo) + _
'                " And TgdSr = " + CStr(ms_TgSr)) = True Then
'    DispMsg "Cannot Generate Records, Since Records already Present", etError: GoTo CloseScp
'  End If
  
  ' Allow Add, Delete Grid Row
  GrdTagData.AllowAdd = True: GrdTagData.AllowDelete = True
  
  'Start Generating Records
  GrdTagData.StartCopy
  MWLib.BeginProcess Me, "Generating Records For Tag Serial ..."

  ' Read Tag Format Strings which has to be Evaluate
  Set moRsFlds = moCn.OpenResultset("Select TgfsSr, TgfsField From TagFmtSr Where " + _
                  "TgfsCoCd = '" + ms_TgCoCd + "' And TgfsCd = '" + adc("TgCd") + "'")
  If moRsFlds.RecCount > 0 Then
    With moRsFlds
      .MoveFirst
      Do While Not (.EOF Or .BOF)
        wsFld = .FldValue("TgfsField") & ""
        Select Case .FldValue("TgfsSr")
          Case 1: wsFld01 = wsFld
          Case 2: wsFld02 = wsFld
          Case 3: wsFld03 = wsFld
          Case 4: wsFld04 = wsFld
          Case 5: wsFld05 = wsFld
          Case 6: wsFld06 = wsFld
          Case 7: wsFld07 = wsFld
          Case 8: wsFld08 = wsFld
          Case 9: wsFld09 = wsFld
          Case 10: wsFld10 = wsFld
          Case 11: wsFld11 = wsFld
          Case 12: wsFld12 = wsFld
        End Select
        .MoveNext
      Loop
    End With
  End If
  'According to TgAvgTyp Records are generated for Bags or InvDsg
  If UCase$(adc("TgAvgTyp")) = "I" Then
    'VK.25 - Changed OdDmStk instead of Bag for jEmr
    If UCase(adc.MenuCd) = UCase("RepJtTagPrn") Then
      wsFlds = ", '' As OdDmStkYy, '' As OdDmStkChr, 0 as OdDmStkNo"
      wsQty = IIF(adc("TgLevel") = "P", "OdOrdQty", "1")
    Else
      wsFlds = ", '' AS BYy, '' As BChr, 0 as BNo"
      wsQty = IIF(adc("TgLevel") = "P", "IdQty", "1")
    End If
    wsJoin = ""
    
  Else
    wsFlds = ", IfFdBYy AS BYy, IfFdBChr As BChr, IfFdBNo as BNo"
    wsJoin = " Inner Join InvFgd On " + IIF(gs_Partition = ctCurrPrtn, " IdPrtKey='" + ctCurrPrtn + "' and ", "") + " IdCoCd = IfCoCd And IdTc = IfTc And IdYy = IfYy And" + _
          " IdChr = IfChr And IdNo = IfNo And IdSr = IfSr"
    wsQty = IIF(adc("TgLevel") = "P", "IfQty", "1")
  End If
  wsQty = ", " + wsQty + " As qQty"
  wsScpCnd = ""
  'VK.25 - Changed Order related fields instead of Invoice for jEmr
  If UCase(adc.MenuCd) = UCase("RepJtTagPrn") Then
    If adc("wIdSrFr") > 0 Then wsScpCnd = wsScpCnd + " And OdSr >= " + CStr(adc("wIdSrFr"))
    If adc("wIdSrTo") > 0 Then wsScpCnd = wsScpCnd + " And OdSr <= " + CStr(adc("wIdSrTo"))
  Else
    If adc("wIdSrFr") > 0 Then wsScpCnd = wsScpCnd + " And IdSr >= " + CStr(adc("wIdSrFr"))
    If adc("wIdSrTo") > 0 Then wsScpCnd = wsScpCnd + " And IdSr <= " + CStr(adc("wIdSrTo"))
  End If

  ' Zubin (Before 08/12/05)
  'wsSql = "Select IdSr" + wsFlds + wsQty + _
          " from InvDsg " + wsJoin + _
          " Where " + _
          " IdCoCd = '" + ms_TgCoCd + "' And IdTc = '" + ms_TgTc + "'" + _
          " And IdYy = '" + ms_TgYy + "' And IdChr = '" + ms_TgChr + "'" + _
          " And IdNo = " + CStr(ms_TgNo) + wsScpCnd
  ' Zubin (08/12/05)
  
  'VK.25 - Changed Order related fields instead of Invoice for jEmr
  If UCase(adc.MenuCd) = UCase("RepJtTagPrn") Then
    wsSql = "Select OdSr As IdSr " + wsFlds + wsQty + _
          " from OrdDsg " + wsJoin + _
          " Where " + IIF(gs_Partition = ctCurrPrtn, " OdPrtKey='" + ctCurrPrtn + "' and ", "") + _
          " OdCoCd = '" + ms_TgCoCd + "' And OdTc = '" + ms_TgTc + "'" + _
          " And OdYy = '" + ms_TgYy + "' And OdChr = '" + ms_TgChr + "'" + _
          " And OdNo = " + CStr(ms_TgNo) + wsScpCnd + " " + _
          " And Not Exists (Select 'x' from TagData where TgdCoCd= OdCoCd and TgdTc= OdTc and " + _
          " TgdYy= OdYy and TgdChr= OdChr and TgdNo= OdNo and TgdIdSr= OdSr and " + _
          " TgdSr= " + CStr(adc("TgSr")) + " " + _
          " ) order by OdSr "
  Else
    wsSql = "Select IdSr" + wsFlds + wsQty + _
          " from InvDsg " + wsJoin + _
          " Where " + IIF(gs_Partition = ctCurrPrtn, " IdPrtKey='" + ctCurrPrtn + "' and ", "") + _
          " IdCoCd = '" + ms_TgCoCd + "' And IdTc = '" + ms_TgTc + "'" + _
          " And IdYy = '" + ms_TgYy + "' And IdChr = '" + ms_TgChr + "'" + _
          " And IdNo = " + CStr(ms_TgNo) + wsScpCnd + " " + _
          " And Not Exists (Select 'x' from TagData where TgdCoCd= IdCoCd and TgdTc= IdTc and " + _
          " TgdYy= IdYy and TgdChr= IdChr and TgdNo= IdNo and TgdIdSr= IdSr and " + _
          " TgdSr= " + CStr(adc("TgSr")) + " " + _
          IIF(UCase$(adc("TgAvgTyp")) = "I", "", " and TgdBagNo= Case when IfFdBYy= '' then '' else " + _
                  "  IfFdBYy+'/'+IfFdBChr+'/'+LTRIM(CAST(IfFdBNo AS varchar(7))) end ") + _
          " ) order by IdSr "
  End If
  mRunSrOld = 0: mRunSrNew = 0 '578-32 default values for running sr
  
  Set moRsSr = moCn.OpenResultset(wsSql)
   If moRsSr.RecCount > 0 Then
    moRsSr.MoveFirst
    ' Zubin (07/12/05)
    wiSrNo = moCn.GetFldVal("Select IsNull(max(TgdSrNo), 0) from TagData where TgdCoCd= '" + gs_CoCd + "' and " + _
                            "TgdTc= '" + adc("TgTc") + "' and TgdYy= '" + adc("TgYy") + "' and " + _
                            "TgdChr= '" + adc("TgChr") + "' and TgdNo= " + CStr(adc("TgNo")) + " and TgdSr= " + CStr(adc("TgSr")) + "")
    Do While Not (moRsSr.EOF Or moRsSr.BOF)
      ' For Each Record Tag Field Values are Set & Then Tag Strings are Evaluated.
      With GrdTagData
        .AddItem
        wRow = .Rows - 1
        ' Zubin (Before 07/12/05)
        'MWLib.RunningProcess "Copying Record " + CStr(wRow) + "/" + CStr(moRsSr.RecCount) + " In Grid ..."
        'wiSrNo = CStr(wRow)
        
        ' Zubin (07/12/05)
        MWLib.RunningProcess "Copying Record " + CStr(moRsSr.Bookmark) + "/" + CStr(moRsSr.RecCount) + " In Grid ..."
        wiSrNo = wiSrNo + 1
        
        .Value(wRow, "TgdSrNo") = wiSrNo
        .Value(wRow, "TgdIdSr") = moRsSr.FldValue("IdSr")
        If UCase$(adc("TgAvgTyp")) = "B" Then
          ' Zubin (Before 06/12/05)
          '.Value(wRow, "wBagNo") = moRsSr.FldValue("BYy") + "/" + moRsSr.FldValue("BChr") + "/" + CStr(moRsSr.FldValue("BNo"))
          .Value(wRow, "TgdBagNo") = moRsSr.FldValue("BYy") + "/" + moRsSr.FldValue("BChr") + "/" + CStr(moRsSr.FldValue("BNo"))
        Else
          ' Zubin (Before 06/12/05)
          .Value(wRow, "TgdBagNo") = ""
        End If
        .Value(wRow, "TgdQty") = moRsSr.FldValue("qQty")
        
        ' Set Values For Tag Fields for Current InvDsg Or Bag
        'VK.25 - Changed Order related fields instead of Invoice and Bag for jEmr
        If UCase(adc.MenuCd) = UCase("RepJtTagPrn") Then
          Call SetTagFlds(moRsSr.FldValue("IdSr"), moRsSr.FldValue("OdDmStkYy"), moRsSr.FldValue("OdDmStkChr"), moRsSr.FldValue("OdDmStkNo"))
        Else
          Call SetTagFlds(moRsSr.FldValue("IdSr"), moRsSr.FldValue("BYy"), moRsSr.FldValue("BChr"), moRsSr.FldValue("BNo"))
        End If
        
        'Evaluate Tag Strings with new Values
        .Value(wRow, "TgdFld01") = MWEval(wsFld01)
        .Value(wRow, "TgdFld02") = MWEval(wsFld02)
        .Value(wRow, "TgdFld03") = MWEval(wsFld03)
        .Value(wRow, "TgdFld04") = MWEval(wsFld04)
        .Value(wRow, "TgdFld05") = MWEval(wsFld05)
        .Value(wRow, "TgdFld06") = MWEval(wsFld06)
        .Value(wRow, "TgdFld07") = MWEval(wsFld07)
        .Value(wRow, "TgdFld08") = MWEval(wsFld08)
        .Value(wRow, "TgdFld09") = MWEval(wsFld09)
        .Value(wRow, "TgdFld10") = MWEval(wsFld10)
        .Value(wRow, "TgdFld11") = MWEval(wsFld11)
        .Value(wRow, "TgdFld12") = MWEval(wsFld12)
        .SaveRec (wRow)
      End With
      moRsSr.MoveNext
    Loop
  End If
  MWLib.EndProcess Me
  GrdTagData.EndCopy
  ' End Generating Records
  ' GrdTagData.AllowAdd = False
  ' Debug.Print GrdTagData.Rows
  If moRsSr.RecCount > 0 Then GrdTagData.Row = 1
CloseScp:
  Call DispFra(GenScp)
End Sub
Private Sub CmdGenScp_Click()
  ' Open Scope Frame
  ' If Record exists in Grid then Set Min & Max Values of IdSr in IdSrFr & IdSrTo
  Dim moRs As New MwfLib.MDORowSet
  
  If Not adc.SaveRec Then
    DispMsg "X Save Failed", etInfo
    Exit Sub
  End If

  Call DispFra(GenScp)
  If GrdTagData.Rows > 1 And FraScp.Visible = True Then
    Set moRs = moCn.OpenResultset("Select Min(TgdIdSr) As qIdSrFr, Max(TgdIdSr) As qIdSrTo From TagData  Where" + _
                                  " TgdCoCd ='" + ms_TgCoCd + "' And TgdTc ='" + ms_TgTc + "'" + _
                                  " And TgdYy ='" + ms_TgYy + "' And TgdChr ='" + ms_TgChr + "'" + _
                                  " And TgdNo =" + CStr(ms_TgNo) + " And TgdSr =" + CStr(ms_TgSr))
    adc("wIdSrFr") = moRs.FldValue("qIdSrFr") & ""
    adc("wIdSrTo") = moRs.FldValue("qIdSrTo") & ""
  End If
End Sub
Private Sub GrdTagData_ColWhen(ByVal RowNum As Integer, ByVal ColName As String, Cancel As Boolean, ErrMsg As String, ByVal Action As MwfCtl.en_fgAction)
  'If on Last Column of Grid is Selected then Focus Set to FraTagFld
  Select Case UCase$(ColName)
  Case Is = UCase$("wFlds")
      'adc("wFld01").SetFocus
  End Select
End Sub
Private Sub GrdTagData_InitKey(ByVal RowNum As Integer)
  ' Initialize Grid Key Fields
  GrdTagData.Store "TgdCoCd", adc("TgCoCd")
  GrdTagData.Store "TgdTc", adc("TgTc")
  GrdTagData.Store "TgdYy", adc("TgYy")
  GrdTagData.Store "TgdChr", adc("TgChr")
  GrdTagData.Store "TgdNo", adc("TgNo")
  GrdTagData.Store "TgdSr", adc("TgSr")
End Sub
Private Sub GrdTagData_RowWhen(ByVal RowNum As Integer)
  Dim moRsFlds As New MwfLib.MDORowSet
  Dim wsFld As String
  If GrdTagData.IsNew(RowNum) = True And adc.Mode = xNorm Then
    ' Zubin (Before 08/12/05)
    'GrdTagData.Value(RowNum, "TgdSrNo") = CStr(GrdTagData.Rows - 1)
    ' Zubin (08/12/05)
    GrdTagData.Value(RowNum, "TgdSrNo") = moCn.GetFldVal("Select IsNull(max(TgdSrNo), 0)+1 from TagData where " + _
                  "TgdCoCd= '" + gs_CoCd + "' and TgdTc= '" + adc("TgTc") + "' and " + _
                  "TgdYy= '" + adc("TgYy") + "' and TgdChr= '" + adc("TgChr") + "' and " + _
                  "TgdNo= " + CStr(adc("TgNo")) + " and TgdSr= " + CStr(adc("TgSr")) + "")
                            
    GrdTagData.Value(RowNum, "TgdIdSr") = CStr(0)
    GrdTagData.Value(RowNum, "TgdQty") = CStr(1)
    
    Set moRsFlds = moCn.OpenResultset("Select TgfsSr, TgfsField From TagFmtSr Where " + _
                  "TgfsCoCd = '" + ms_TgCoCd + "' And TgfsCd = '" + adc("TgCd") + "'")
    If moRsFlds.RecCount > 0 Then
      With moRsFlds
        .MoveFirst
        Do While Not (.EOF Or .BOF)
          wsFld = Replace(.FldValue("TgfsField") & "", "|", "")
          Select Case .FldValue("TgfsSr")
            Case 1: GrdTagData.Value(RowNum, "TgdFld01") = wsFld
            Case 2: GrdTagData.Value(RowNum, "TgdFld02") = wsFld
            Case 3: GrdTagData.Value(RowNum, "TgdFld03") = wsFld
            Case 4: GrdTagData.Value(RowNum, "TgdFld04") = wsFld
            Case 5: GrdTagData.Value(RowNum, "TgdFld05") = wsFld
            Case 6: GrdTagData.Value(RowNum, "TgdFld06") = wsFld
            Case 7: GrdTagData.Value(RowNum, "TgdFld07") = wsFld
            Case 8: GrdTagData.Value(RowNum, "TgdFld08") = wsFld
            Case 9: GrdTagData.Value(RowNum, "TgdFld09") = wsFld
            Case 10: GrdTagData.Value(RowNum, "TgdFld10") = wsFld
            Case 11: GrdTagData.Value(RowNum, "TgdFld11") = wsFld
            Case 12: GrdTagData.Value(RowNum, "TgdFld12") = wsFld
          End Select
          .MoveNext
        Loop
      End With
    End If
  End If
  Set moRsFlds = Nothing
  ' Set Tag Flds as Current Grid Record
  With GrdTagData
    adc("wFld01") = .Value(RowNum, "TgdFld01")
    adc("wFld02") = .Value(RowNum, "TgdFld02")
    adc("wFld03") = .Value(RowNum, "TgdFld03")
    adc("wFld04") = .Value(RowNum, "TgdFld04")
    adc("wFld05") = .Value(RowNum, "TgdFld05")
    adc("wFld06") = .Value(RowNum, "TgdFld06")
    adc("wFld07") = .Value(RowNum, "TgdFld07")
    adc("wFld08") = .Value(RowNum, "TgdFld08")
    adc("wFld09") = .Value(RowNum, "TgdFld09")
    adc("wFld10") = .Value(RowNum, "TgdFld10")
    adc("wFld11") = .Value(RowNum, "TgdFld11")
    adc("wFld12") = .Value(RowNum, "TgdFld12")
  End With
End Sub
Private Sub GrdTagData_SetRecSource()
  ' Set Recordset for Grid
  GrdTagData.RecSource = " Select * from TagData " + _
                  " Where TgdCoCd ='" + adc("TgCoCd") + "' And TgdTc ='" + adc("TgTc") + "'" + _
                  " And TgdYy ='" + adc("TgYy") + "' And TgdChr ='" + adc("TgChr") + "'" + _
                  " And TgdNo =" + CStr(adc("TgNo")) + " And TgdSr =" + CStr(adc("TgSr"))
End Sub
Private Sub ADC_SetRecSource()
  '*** Set the RecSource for ADC
  adc.RecSource = " Select * from Tag " + _
                  " Where TgCoCd ='" + adc("TgCoCd") + "' And TgTc ='" + adc("TgTc") + "'" + _
                  " And TgYy ='" + adc("TgYy") + "' And TgChr ='" + adc("TgChr") + "'" + _
                  " And TgNo =" + CStr(adc("TgNo")) + " And TgSr =" + CStr(adc("TgSr"))
End Sub
'''  ***** Std Code not to be changed *****
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
  ws_BtnStr = "CmdFldsOk,CmdPrnGo,CmdGenGo"
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
' ########################  Manoj #### Ver: 2.0.6 #### Date: 20/02/2004  ########################
' Initialize Colllection of Tag Field
  Call SetTagFldColl(mColTagFlds)
' ###########################################  Manoj  ###########################################
End Sub
Private Sub Form_Deactivate()
  Call FrmDeActivate(Me)
End Sub
Private Sub ATXT_Validate(Index As Integer, Cancel As Boolean)
  Cancel = atxt(Index).Validate
End Sub
Private Sub TXTScp_Validate(Index As Integer, Cancel As Boolean)
  Cancel = TXTScp(Index).Validate
End Sub
Private Sub TXTFlds_Validate(Index As Integer, Cancel As Boolean)
  Cancel = TXTFlds(Index).Validate
End Sub
Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
  If UnloadMode <> vbFormCode Then Cancel = True: Exit Sub
  '*** (09/08/05)
  Set moCn = Nothing   '*** (Bef speed) bef 24/11/06
  Set mColTagFlds = Nothing
  Set scTag = Nothing
  Set mRs_Rep = Nothing
  Set mRs_FmtStr = Nothing
  '*** (09/08/05)
End Sub
Private Sub GrdTagData_Validate(Cancel As Boolean)
  Cancel = GrdTagData.Validate
End Sub

Private Sub DispFra(ByVal pv_TagFra As en_TagFra)
  '*** This is the common procedure which is used to toggle between hiding and displaying
      'the frames (i.e. if the frame is visible then make it invisible and viseversa)
      'associated with a command button
  Dim wb_Val As Boolean, wb_NVal As Boolean
  Select Case pv_TagFra
  Case Is = GenScp
    '*** If Frame FraScp is visible then make it invisible and enable all the command buttons
    '*** If Frame FraScp is invisible then make it visible and disable all the
        'command buttons except the one which is clicked. Change the background colour of
        'the clicked command button to white and make the font bold
    wb_Val = FraScp.Visible: wb_NVal = Not FraScp.Visible
    
    FraScp.Visible = wb_NVal
    FraScp.Enabled = wb_NVal
    Call EnaDisaCmds(wb_NVal, CmdGenScp)
    
    If wb_Val = True Then
      CmdGenScp.SetFocus
    Else
      adc("wIdSrFr").SetFocus
      FraScp.ZOrder
    End If

Case Is = PrnScp
    '*** If Frame FraPrn is visible then make it invisible and enable all the command buttons
    '*** If Frame FraPrn is invisible then make it visible and disable all the
        'command buttons except the one which is clicked. Change the background colour of
        'the clicked command button to white and make the font bold
    ' If PrntOpt="F" Then Hide Printer Combo
    ' If PrntOpt="P" Then Hide File Name
    wb_Val = FraPrn.Visible: wb_NVal = Not FraPrn.Visible
    
    FraPrn.Visible = wb_NVal
    FraPrn.Enabled = wb_NVal
    
    If mPrntOpt = "F" Then
      ALBL(20).Visible = wb_Val
      cboPrn.Visible = wb_Val
      cboPrn.Enabled = wb_Val
      Call EnaDisaCmds(wb_NVal, CmdTagFilePrint)
    ElseIf mPrntOpt = "P" Then
      ALBL(21).Visible = wb_Val
      adc("wFileNm").Visible = wb_Val
      adc("wFileNm").Enabled = wb_Val
      CmdFileOpen.Visible = wb_Val
      Call EnaDisaCmds(wb_NVal, CmdTagPrnPrint)
    End If
    
    If wb_Val = True Then
      If mPrntOpt = "F" Then
        CmdTagFilePrint.SetFocus
      ElseIf mPrntOpt = "P" Then
        CmdTagPrnPrint.SetFocus
      End If
    Else
      If mPrntOpt = "F" Then
        adc("wFileNm").SetFocus
      ElseIf mPrntOpt = "P" Then
        cboPrn.SetFocus
      End If
      FraPrn.ZOrder
    End If
  End Select
  FraPrnFlds.Enabled = wb_Val
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
  adc.AllowSave = Not pv_ShowFra
  adc.AllowDelete = Not pv_ShowFra
  CmdGenScp.Enabled = Not pv_ShowFra
  CmdGenScp.Enabled = Not pv_ShowFra
  CmdTagFilePrint.Enabled = Not pv_ShowFra
  CmdTagPrnPrint.Enabled = Not pv_ShowFra
  If Not pv_CmdBut Is Nothing Then pv_CmdBut.Enabled = True
End Sub
Private Sub HideAllFras()
  '*** This procedure hides all the frames other than FraNKeyAll
  FraNKeyAll.Visible = True
  FraScp.Visible = False
  FraScp.Enabled = False
  FraPrn.Visible = False
  FraPrn.Enabled = False
End Sub

Private Sub TXTFlds_GotFocus(Index As Integer)
  ' Disable all Cmd Buttons
  ' Disable FraNkeyAll
  If GrdTagData.Row > 0 Then
    adc.AllowSave = False
    Call EnaDisaCmds(True)
    FraNKeyAll.Enabled = False
  End If
End Sub

Private Function MWEval(ps_ExpStr As String) As String
  'Evaluate String
  ' every alternate willbe pass for further Evaluation of Expression
  On Error GoTo ErrHndlr
  Dim wstVal As String, wiStStrt As Integer, wiStEnd As Integer
  Dim wstValSplit() As String, i As Integer, wTemp As Variant
  wstVal = ps_ExpStr
  MWEval = wstVal
  ' Split the String using "|"
  wstValSplit = Split(wstVal, "|")
  wstVal = ""
  If UBound(wstValSplit) >= 0 Then
    For i = 0 To UBound(wstValSplit)
      If (i Mod 2) = 1 Then
        'Evaluate Every Even String
        wstVal = wstVal + mEvalExp(scTag, RemoveWhiteSpaces(wstValSplit(i), True))
      Else
        'Every Odd String Kept as it is
        wstVal = wstVal + RemoveWhiteSpaces(wstValSplit(i), True)
      End If
      Next
  End If
  MWEval = wstVal
  Exit Function
ErrHndlr:
  DispMsg Err.Number & " : " & Err.Description & vbCrLf & Err.Source, etInfo
End Function
Public Sub SetTagFlds(ps_IdSr As Double, ps_BYy As String, ps_BChr As String, ps_BNo As Long)
  Dim mRSTgFld As New MwfLib.MDORowSet
  Dim wsSql As String
  '578-32 creating running serial number
  If mRunSrOld = 0 Then
    mRunSrOld = ps_IdSr
    mRunSrNew = 1
  End If
  If mRunSrOld <> ps_IdSr Then mRunSrNew = mRunSrNew + 1
  mRunSrOld = ps_IdSr

  'Calculating & Setting Tag Field Values
  SetTagFldVal mColTagFlds, adc("TgMulBy"), "TgMulBy"
  
  ' Generating Records of Tag Values Which directly get from the Field
  wsSql = GetMainSql(ps_IdSr, ps_BYy, ps_BChr, ps_BNo)
  Set mRSTgFld = moCn.OpenResultset(wsSql)
  
  If mRSTgFld.RecCount > 0 Then
    ' Setting Tag Field Value
    SetTagFldVal mColTagFlds, mRSTgFld.FldValue("InTc"), "InTc"
    SetTagFldVal mColTagFlds, mRSTgFld.FldValue("InYy"), "InYy"
    SetTagFldVal mColTagFlds, mRSTgFld.FldValue("InChr"), "InChr"
    SetTagFldVal mColTagFlds, mRSTgFld.FldValue("InNo"), "InNo"
    ' Zubin
    SetTagFldVal mColTagFlds, "''", "FdTc"
    SetTagFldVal mColTagFlds, "''", "FdYy"
    SetTagFldVal mColTagFlds, "''", "FdChr"
    SetTagFldVal mColTagFlds, "''", "FdNo"
    ' Zubin
    SetTagFldVal mColTagFlds, mRSTgFld.FldValue("InvNo"), "InvNo"
    SetTagFldVal mColTagFlds, mRSTgFld.FldValue("IdSr"), "IdSr"
    SetTagFldVal mColTagFlds, mRSTgFld.FldValue("InvSr"), "InvSr"
    SetTagFldVal mColTagFlds, mRunSrNew, "RunSr"   '578-32 running serial added
        
    SetTagFldVal mColTagFlds, mRSTgFld.FldValue("IdKt"), "IdKt"
    ' Zubin
    SetTagFldVal mColTagFlds, mRSTgFld.FldValue("IdKt"), "OdKt"
    ' Zubin
    SetTagFldVal mColTagFlds, mRSTgFld.FldValue("DmCol"), "IdCol"
    SetTagFldVal mColTagFlds, mRSTgFld.FldValue("InExpNo"), "InExpNo"
    
    SetTagFldVal mColTagFlds, mRSTgFld.FldValue("IdQty"), "IdQty"
    SetTagFldVal mColTagFlds, mRSTgFld.FldValue("IdDQty"), "IdDQty"
    SetTagFldVal mColTagFlds, mRSTgFld.FldValue("IdCQty"), "IdCQty"
    SetTagFldVal mColTagFlds, mRSTgFld.FldValue("IdDCQty"), "IdDCQty"         ' ***** Manali 3.8.0 - DC added
    SetTagFldVal mColTagFlds, mRSTgFld.FldValue("IdDWt"), "IdDWt"
    SetTagFldVal mColTagFlds, mRSTgFld.FldValue("IdCWt"), "IdCWt"
    SetTagFldVal mColTagFlds, mRSTgFld.FldValue("IdDCWt"), "IdDCWt"         ' ***** Manali 3.8.0 - DC added
    SetTagFldVal mColTagFlds, mRSTgFld.FldValue("IdGWt"), "IdGWt"
    SetTagFldVal mColTagFlds, mRSTgFld.FldValue("IdPWt"), "IdPWt"
    SetTagFldVal mColTagFlds, mRSTgFld.FldValue("IdSWt"), "IdSWt"         ' ***** Manali 3.8.0 - S added
    SetTagFldVal mColTagFlds, mRSTgFld.FldValue("IdLWt"), "IdLWt"         ' ***** Manali 3.8.0 - L added
    SetTagFldVal mColTagFlds, mRSTgFld.FldValue("IdXQty"), "IdXQty"
    SetTagFldVal mColTagFlds, mRSTgFld.FldValue("IdXWt"), "IdXWt"
    SetTagFldVal mColTagFlds, mRSTgFld.FldValue("IdMWt"), "IdMWt"         ' ***** Manali 3.9.0 - Oth Met Added
    SetTagFldVal mColTagFlds, mRSTgFld.FldValue("IdGrsWt"), "IdGrsWt"
    SetTagFldVal mColTagFlds, mRSTgFld.FldValue("IdVal"), "IdVal"
    ' Zubin
    SetTagFldVal mColTagFlds, mRSTgFld.FldValue("IdMuVal"), "IdMuVal"
    
    SetTagFldVal mColTagFlds, mRSTgFld.FldValue("CmCd"), "CmCd"
    SetTagFldVal mColTagFlds, mRSTgFld.FldValue("DmCd"), "DmCd"
    SetTagFldVal mColTagFlds, mRSTgFld.FldValue("CdCustDmCd"), "DmCustCd"
    SetTagFldVal mColTagFlds, mRSTgFld.FldValue("DmSfx"), "DmSfx"
    SetTagFldVal mColTagFlds, mRSTgFld.FldValue("DmSz"), "DmSz"
    SetTagFldVal mColTagFlds, mRSTgFld.FldValue("DmCtg"), "DmCtg"
  
    'VK.25 - BCoCd added
    SetTagFldVal mColTagFlds, mRSTgFld.FldValue("BCoCd"), "BCoCd"
    SetTagFldVal mColTagFlds, mRSTgFld.FldValue("BYy"), "BYy"
    SetTagFldVal mColTagFlds, mRSTgFld.FldValue("BChr"), "BChr"
    SetTagFldVal mColTagFlds, mRSTgFld.FldValue("BNo"), "BNo"
    SetTagFldVal mColTagFlds, mRSTgFld.FldValue("BagNo"), "BagNo"
  
    SetTagFldVal mColTagFlds, mRSTgFld.FldValue("BagQty"), "BagQty"
    SetTagFldVal mColTagFlds, mRSTgFld.FldValue("BagDQty"), "BagDQty"
    SetTagFldVal mColTagFlds, mRSTgFld.FldValue("BagCQty"), "BagCQty"
    SetTagFldVal mColTagFlds, mRSTgFld.FldValue("BagDCQty"), "BagDCQty"         ' ***** Manali 3.8.0 - DC added
    SetTagFldVal mColTagFlds, mRSTgFld.FldValue("BagDWt"), "BagDWt"
    SetTagFldVal mColTagFlds, mRSTgFld.FldValue("BagCWt"), "BagCWt"
    SetTagFldVal mColTagFlds, mRSTgFld.FldValue("BagDCWt"), "BagDCWt"           ' ***** Manali 3.8.0 - DC added
    SetTagFldVal mColTagFlds, mRSTgFld.FldValue("BagGWt"), "BagGWt"
    SetTagFldVal mColTagFlds, mRSTgFld.FldValue("BagPWt"), "BagPWt"
    SetTagFldVal mColTagFlds, mRSTgFld.FldValue("BagSWt"), "BagSWt"        ' ***** Manali 3.8.0 - S added
    SetTagFldVal mColTagFlds, mRSTgFld.FldValue("BagLWt"), "BagLWt"        ' ***** Manali 3.8.0 - L added
    SetTagFldVal mColTagFlds, mRSTgFld.FldValue("BagXQty"), "BagXQty"
    SetTagFldVal mColTagFlds, mRSTgFld.FldValue("BagXWt"), "BagXWt"
    SetTagFldVal mColTagFlds, mRSTgFld.FldValue("BagMWt"), "BagMWt"         ' ***** Manali 3.9.0 - Oth Met Added
    SetTagFldVal mColTagFlds, mRSTgFld.FldValue("BagGrsWt"), "BagGrsWt"
    SetTagFldVal mColTagFlds, mRSTgFld.FldValue("BagVal"), "BagVal"
    ' Zubin
    SetTagFldVal mColTagFlds, mRSTgFld.FldValue("BagMuVal"), "BagMuVal"

' ###########################################  Manoj  ###########################################
    SetTagFldVal mColTagFlds, mRSTgFld.FldValue("ExpOdTc"), "ExpOdTc"
    SetTagFldVal mColTagFlds, mRSTgFld.FldValue("ExpOdYy"), "ExpOdYy"
    SetTagFldVal mColTagFlds, mRSTgFld.FldValue("ExpOdChr"), "ExpOdChr"
    SetTagFldVal mColTagFlds, mRSTgFld.FldValue("ExpOdNo"), "ExpOdNo"
    SetTagFldVal mColTagFlds, mRSTgFld.FldValue("ExpOrdNo"), "ExpOrdNo"
    SetTagFldVal mColTagFlds, mRSTgFld.FldValue("ExpOdSr"), "ExpOdSr"
    SetTagFldVal mColTagFlds, mRSTgFld.FldValue("ExpOrdSr"), "ExpOrdSr"
    ' Manoj 2.9.0 Patch
    SetTagFldVal mColTagFlds, mRSTgFld.FldValue("OmPoNo"), "OmPoNo"
    ' Manoj 2.9.0 Patch
    
    SetTagFldVal mColTagFlds, mRSTgFld.FldValue("OdPoNo"), "OdPoNo"         '3.12.0
    SetTagFldVal mColTagFlds, mRSTgFld.FldValue("OdMrp"), "OdMrp"           '4.1.4
    SetTagFldVal mColTagFlds, mRSTgFld.FldValue("HUIDNO"), "HUIDNO"     '578.45 Huid added
'    SetTagFldVal mColTagFlds, mRSTgFld.FldValue("PrdOdTc"), "PrdOdTc"
'    SetTagFldVal mColTagFlds, mRSTgFld.FldValue("PrdOdYy"), "PrdOdYy"
'    SetTagFldVal mColTagFlds, mRSTgFld.FldValue("PrdOdChr"), "PrdOdChr"
'    SetTagFldVal mColTagFlds, mRSTgFld.FldValue("PrdOdNo"), "PrdOdNo"
'    SetTagFldVal mColTagFlds, mRSTgFld.FldValue("PrdOrdNo"), "PrdOrdNo"
'    SetTagFldVal mColTagFlds, mRSTgFld.FldValue("PrdOdSr"), "PrdOdSr"
'    SetTagFldVal mColTagFlds, mRSTgFld.FldValue("PrdOrdSr"), "PrdOrdSr"
' ###########################################  Manoj  ###########################################
  
    If UCase$(adc("TgLevel")) = UCase$("P") Then
      SetTagFldVal mColTagFlds, mRSTgFld.FldValue("PcDWt"), "PcDWt"
      SetTagFldVal mColTagFlds, mRSTgFld.FldValue("PcCWt"), "PcCWt"
      SetTagFldVal mColTagFlds, mRSTgFld.FldValue("PcDCWt"), "PcDCWt"         ' ***** Manali 3.8.0 - DC added
      SetTagFldVal mColTagFlds, mRSTgFld.FldValue("PcDQty"), "PcDQty"
      SetTagFldVal mColTagFlds, mRSTgFld.FldValue("PcCQty"), "PcCQty"
      SetTagFldVal mColTagFlds, mRSTgFld.FldValue("PcDCQty"), "PcDCQty"         ' ***** Manali 3.8.0 - DC added
      SetTagFldVal mColTagFlds, mRSTgFld.FldValue("PcGWt"), "PcGWt"
      SetTagFldVal mColTagFlds, mRSTgFld.FldValue("PcPWt"), "PcPWt"
      SetTagFldVal mColTagFlds, mRSTgFld.FldValue("PcSWt"), "PcSWt"             ' ***** Manali 3.8.0 - S added
      SetTagFldVal mColTagFlds, mRSTgFld.FldValue("PcLWt"), "PcLWt"             ' ***** Manali 3.8.0 - L added
      SetTagFldVal mColTagFlds, mRSTgFld.FldValue("PcXQty"), "PcXQty"
      SetTagFldVal mColTagFlds, mRSTgFld.FldValue("PcXWt"), "PcXWt"
      SetTagFldVal mColTagFlds, mRSTgFld.FldValue("PcMWt"), "PcMWt"             ' ***** Manali 3.9.0 - Oth Met Added
      SetTagFldVal mColTagFlds, mRSTgFld.FldValue("PcGrsWt"), "PcGrsWt"
      SetTagFldVal mColTagFlds, mRSTgFld.FldValue("PcVal"), "PcVal"
      ' Zubin
      SetTagFldVal mColTagFlds, mRSTgFld.FldValue("PcMuVal"), "PcMuVal"
      
    Else
      If UCase$(adc("TgAvgTyp")) = UCase$("I") Then
        SetTagFldVal mColTagFlds, mRSTgFld.FldValue("IdDWt"), "PcDWt"
        SetTagFldVal mColTagFlds, mRSTgFld.FldValue("IdCWt"), "PcCWt"
        SetTagFldVal mColTagFlds, mRSTgFld.FldValue("IdDCWt"), "PcDCWt"         ' ***** Manali 3.8.0 - DC added
        SetTagFldVal mColTagFlds, mRSTgFld.FldValue("IdDQty"), "PcDQty"
        SetTagFldVal mColTagFlds, mRSTgFld.FldValue("IdCQty"), "PcCQty"
        SetTagFldVal mColTagFlds, mRSTgFld.FldValue("IdDCQty"), "PcDCQty"         ' ***** Manali 3.8.0 - DC added
        SetTagFldVal mColTagFlds, mRSTgFld.FldValue("IdGWt"), "PcGWt"
        SetTagFldVal mColTagFlds, mRSTgFld.FldValue("IdPWt"), "PCPWt"
        SetTagFldVal mColTagFlds, mRSTgFld.FldValue("IdSWt"), "PcSWt"             ' ***** Manali 3.8.0 - S added
        SetTagFldVal mColTagFlds, mRSTgFld.FldValue("IdLWt"), "PCLWt"             ' ***** Manali 3.8.0 - L added
        SetTagFldVal mColTagFlds, mRSTgFld.FldValue("IdXQty"), "PcXQty"
        SetTagFldVal mColTagFlds, mRSTgFld.FldValue("IdXWt"), "PcXWt"
        SetTagFldVal mColTagFlds, mRSTgFld.FldValue("IdMWt"), "PcMWt"             ' ***** Manali 3.9.0 - Oth Met Added
        SetTagFldVal mColTagFlds, mRSTgFld.FldValue("IdGrsWt"), "PcGrsWt"
        SetTagFldVal mColTagFlds, mRSTgFld.FldValue("IdVal"), "PcVal"
      ' Zubin
      SetTagFldVal mColTagFlds, mRSTgFld.FldValue("IdMuVal"), "PcMuVal"
        
      ElseIf UCase$(adc("TgAvgTyp")) = UCase$("B") Then
        SetTagFldVal mColTagFlds, mRSTgFld.FldValue("BagDWt"), "PcDWt"
        SetTagFldVal mColTagFlds, mRSTgFld.FldValue("BagCWt"), "PcCWt"
        SetTagFldVal mColTagFlds, mRSTgFld.FldValue("BagDCWt"), "PcDCWt"         ' ***** Manali 3.8.0 - DC added
        SetTagFldVal mColTagFlds, mRSTgFld.FldValue("BagDQty"), "PcDQty"
        SetTagFldVal mColTagFlds, mRSTgFld.FldValue("BagCQty"), "PcCQty"
        SetTagFldVal mColTagFlds, mRSTgFld.FldValue("BagDCQty"), "PcDCQty"         ' ***** Manali 3.8.0 - DC added
        SetTagFldVal mColTagFlds, mRSTgFld.FldValue("BagGWt"), "PcGWt"
        SetTagFldVal mColTagFlds, mRSTgFld.FldValue("BagPWt"), "PCPWt"
        SetTagFldVal mColTagFlds, mRSTgFld.FldValue("BagSWt"), "PcSWt"             ' ***** Manali 3.8.0 - S added
        SetTagFldVal mColTagFlds, mRSTgFld.FldValue("BagLWt"), "PCLWt"             ' ***** Manali 3.8.0 - L added
        SetTagFldVal mColTagFlds, mRSTgFld.FldValue("BagXQty"), "PcXQty"
        SetTagFldVal mColTagFlds, mRSTgFld.FldValue("BagXWt"), "PcXWt"
        SetTagFldVal mColTagFlds, mRSTgFld.FldValue("BagMWt"), "PcMWt"              ' ***** Manali 3.9.0 - Oth Met Added
        SetTagFldVal mColTagFlds, mRSTgFld.FldValue("BagGrsWt"), "PcGrsWt"
        SetTagFldVal mColTagFlds, mRSTgFld.FldValue("BagVal"), "PcVal"
        ' Zubin
        SetTagFldVal mColTagFlds, mRSTgFld.FldValue("BagMuVal"), "PcMuVal"
      End If
    End If
  End If
  
  ' Generating Records of Tag Values Which Calculated from Function GetTgStrFrRec() for RmSCtg
  wsSql = GetSubSql("RmSCtg", ps_IdSr, ps_BYy, ps_BChr, ps_BNo)
  Set mRSTgFld = moCn.OpenResultset(wsSql)
  
  If mRSTgFld.RecCount > 0 Then
    ' Setting Tag Field Value
    SetTagFldVal mColTagFlds, GetTgStrFrRec(mRSTgFld, "RmSCtg", "RmCtg", "D"), "DRmSCtg"
    SetTagFldVal mColTagFlds, GetTgStrFrRec(mRSTgFld, "RmSCtg", "RmCtg", "C"), "CSRmSCtg"
    SetTagFldVal mColTagFlds, GetTgStrFrRec(mRSTgFld, "RmSCtg", "RmCtg", "G"), "GRmSCtg"
    SetTagFldVal mColTagFlds, GetTgStrFrRec(mRSTgFld, "RmSCtg", "RmCtg", "P"), "PRmSCtg"
    ' ***** Manali 3.03 - Silver and Palladium included
    SetTagFldVal mColTagFlds, GetTgStrFrRec(mRSTgFld, "RmSCtg", "RmCtg", "S"), "SRmSCtg"
    SetTagFldVal mColTagFlds, GetTgStrFrRec(mRSTgFld, "RmSCtg", "RmCtg", "L"), "LRmSCtg"
    ' ***** Manali 3.03 - Silver and Palladium included
    SetTagFldVal mColTagFlds, GetTgStrFrRec(mRSTgFld, "RmSCtg", "RmCtg", "X"), "XRmSCtg"
     SetTagFldVal mColTagFlds, GetTgStrFrRec(mRSTgFld, "RmSCtg", "RmCtg", "M"), "MRmSCtg"           ' ***** Manali 3.9.0 - Oth Met Added
    
    If UCase$(adc("TgLevel")) = UCase$("P") Then
      SetTagFldVal mColTagFlds, GetTgStrFrRec(mRSTgFld, "PcDWt", "RmCtg", "D"), "DWRmSCtg"
      SetTagFldVal mColTagFlds, GetTgStrFrRec(mRSTgFld, "PcCWt", "RmCtg", "C"), "CWRmSCtg"
      SetTagFldVal mColTagFlds, GetTgStrFrRec(mRSTgFld, "PcGWt", "RmCtg", "G"), "GWRmSCtg"
      SetTagFldVal mColTagFlds, GetTgStrFrRec(mRSTgFld, "PcPWt", "RmCtg", "P"), "PWRmSCtg"
      ' ***** Manali 3.03 - Silver and Palladium included
      SetTagFldVal mColTagFlds, GetTgStrFrRec(mRSTgFld, "PcSWt", "RmCtg", "S"), "SWRmSCtg"
      SetTagFldVal mColTagFlds, GetTgStrFrRec(mRSTgFld, "PcLWt", "RmCtg", "L"), "LWRmSCtg"
      ' ***** Manali 3.03 - Silver and Palladium included
      SetTagFldVal mColTagFlds, GetTgStrFrRec(mRSTgFld, "PcXWt", "RmCtg", "X"), "XWRmSCtg"
      SetTagFldVal mColTagFlds, GetTgStrFrRec(mRSTgFld, "PcMWt", "RmCtg", "M"), "MWRmSCtg"          ' ***** Manali 3.9.0 - Oth Met Added
      
      SetTagFldVal mColTagFlds, GetTgStrFrRec(mRSTgFld, "PcDQty", "RmCtg", "D"), "DQRmSCtg"
      SetTagFldVal mColTagFlds, GetTgStrFrRec(mRSTgFld, "PcCQty", "RmCtg", "C"), "CQRmSCtg"
      SetTagFldVal mColTagFlds, GetTgStrFrRec(mRSTgFld, "PcXQty", "RmCtg", "X"), "XQRmSCtg"
    Else
      If UCase$(adc("TgAvgTyp")) = UCase$("I") Then
        SetTagFldVal mColTagFlds, GetTgStrFrRec(mRSTgFld, "IdDWt", "RmCtg", "D"), "DWRmSCtg"
        SetTagFldVal mColTagFlds, GetTgStrFrRec(mRSTgFld, "IdCWt", "RmCtg", "C"), "CWRmSCtg"
        SetTagFldVal mColTagFlds, GetTgStrFrRec(mRSTgFld, "IdGWt", "RmCtg", "G"), "GWRmSCtg"
        SetTagFldVal mColTagFlds, GetTgStrFrRec(mRSTgFld, "IdPWt", "RmCtg", "P"), "PWRmSCtg"
        ' ***** Manali 3.03 - Silver and Palladium included
        SetTagFldVal mColTagFlds, GetTgStrFrRec(mRSTgFld, "IdSWt", "RmCtg", "S"), "SWRmSCtg"
        SetTagFldVal mColTagFlds, GetTgStrFrRec(mRSTgFld, "IdLWt", "RmCtg", "L"), "LWRmSCtg"
        ' ***** Manali 3.03 - Silver and Palladium included
        SetTagFldVal mColTagFlds, GetTgStrFrRec(mRSTgFld, "IdXWt", "RmCtg", "X"), "XWRmSCtg"
        SetTagFldVal mColTagFlds, GetTgStrFrRec(mRSTgFld, "IdMWt", "RmCtg", "M"), "MWRmSCtg"        ' ***** Manali 3.9.0 - Oth Met Added
        
        SetTagFldVal mColTagFlds, GetTgStrFrRec(mRSTgFld, "IdDQty", "RmCtg", "D"), "DQRmSCtg"
        SetTagFldVal mColTagFlds, GetTgStrFrRec(mRSTgFld, "IdCQty", "RmCtg", "C"), "CQRmSCtg"
        SetTagFldVal mColTagFlds, GetTgStrFrRec(mRSTgFld, "IdCQty", "RmCtg", "X"), "XQRmSCtg"
      ElseIf UCase$(adc("TgAvgTyp")) = UCase$("B") Then
        SetTagFldVal mColTagFlds, GetTgStrFrRec(mRSTgFld, "BagDWt", "RmCtg", "D"), "DWRmSCtg"
        SetTagFldVal mColTagFlds, GetTgStrFrRec(mRSTgFld, "BagCWt", "RmCtg", "C"), "CWRmSCtg"
        SetTagFldVal mColTagFlds, GetTgStrFrRec(mRSTgFld, "BagGWt", "RmCtg", "G"), "GWRmSCtg"
        SetTagFldVal mColTagFlds, GetTgStrFrRec(mRSTgFld, "BagPWt", "RmCtg", "P"), "PWRmSCtg"
        ' ***** Manali 3.03 - Silver and Palladium included
        SetTagFldVal mColTagFlds, GetTgStrFrRec(mRSTgFld, "BagSWt", "RmCtg", "S"), "SWRmSCtg"
        SetTagFldVal mColTagFlds, GetTgStrFrRec(mRSTgFld, "BagLWt", "RmCtg", "L"), "LWRmSCtg"
        ' ***** Manali 3.03 - Silver and Palladium included
        SetTagFldVal mColTagFlds, GetTgStrFrRec(mRSTgFld, "BagXWt", "RmCtg", "X"), "XWRmSCtg"
        SetTagFldVal mColTagFlds, GetTgStrFrRec(mRSTgFld, "BagMWt", "RmCtg", "M"), "MWRmSCtg"         ' ***** Manali 3.9.0 - Oth Met Added
        
        SetTagFldVal mColTagFlds, GetTgStrFrRec(mRSTgFld, "BagDQty", "RmCtg", "D"), "DQRmSCtg"
        SetTagFldVal mColTagFlds, GetTgStrFrRec(mRSTgFld, "BagCQty", "RmCtg", "C"), "CQRmSCtg"
        SetTagFldVal mColTagFlds, GetTgStrFrRec(mRSTgFld, "BagXQty", "RmCtg", "X"), "XQRmSCtg"
      End If
    End If
  End If
 
  wsSql = GetSubSql("RmStnCls", ps_IdSr, ps_BYy, ps_BChr, ps_BNo)
  Set mRSTgFld = moCn.OpenResultset(wsSql)
  
  If mRSTgFld.RecCount > 0 Then
    SetTagFldVal mColTagFlds, GetTgStrFrRec(mRSTgFld, "RmStnCls", "RmCtg", "D"), "DRmStnCls"
    SetTagFldVal mColTagFlds, GetTgStrFrRec(mRSTgFld, "RmStnCls", "RmCtg", "C"), "CSRmStnCls"
    
    If UCase$(adc("TgLevel")) = UCase$("P") Then
      SetTagFldVal mColTagFlds, GetTgStrFrRec(mRSTgFld, "PcDWt", "RmCtg", "D"), "DWRmStnCls"
      SetTagFldVal mColTagFlds, GetTgStrFrRec(mRSTgFld, "PcCWt", "RmCtg", "C"), "CWRmStnCls"
      
      SetTagFldVal mColTagFlds, GetTgStrFrRec(mRSTgFld, "PcDQty", "RmCtg", "D"), "DQRmSCtg"
      SetTagFldVal mColTagFlds, GetTgStrFrRec(mRSTgFld, "PcCQty", "RmCtg", "C"), "CQRmSCtg"
    Else
      If UCase$(adc("TgAvgTyp")) = UCase$("I") Then
        SetTagFldVal mColTagFlds, GetTgStrFrRec(mRSTgFld, "IdDWt", "RmCtg", "D"), "DWRmStnCls"
        SetTagFldVal mColTagFlds, GetTgStrFrRec(mRSTgFld, "IdCWt", "RmCtg", "C"), "CWRmStnCls"
        
        SetTagFldVal mColTagFlds, GetTgStrFrRec(mRSTgFld, "IdDQty", "RmCtg", "D"), "DQRmStnCls"
        SetTagFldVal mColTagFlds, GetTgStrFrRec(mRSTgFld, "IdCQty", "RmCtg", "C"), "CQRmStnCls"
      ElseIf UCase$(adc("TgAvgTyp")) = UCase$("B") Then
        SetTagFldVal mColTagFlds, GetTgStrFrRec(mRSTgFld, "BagDWt", "RmCtg", "D"), "DWRmStnCls"
        SetTagFldVal mColTagFlds, GetTgStrFrRec(mRSTgFld, "BagCWt", "RmCtg", "C"), "CWRmStnCls"
        
        SetTagFldVal mColTagFlds, GetTgStrFrRec(mRSTgFld, "BagDQty", "RmCtg", "D"), "DQRmStnCls"
        SetTagFldVal mColTagFlds, GetTgStrFrRec(mRSTgFld, "BagCQty", "RmCtg", "C"), "CQRmStnCls"
      End If
    End If
  End If
  
  ' Generating Records of Tag Values Which Calculated from Function GetTgStrFrRec() for RmCd
  wsSql = GetSubSql("RmCd", ps_IdSr, ps_BYy, ps_BChr, ps_BNo)
  Set mRSTgFld = moCn.OpenResultset(wsSql)
  
  If mRSTgFld.RecCount > 0 Then
    ' Setting Tag Field Value
    SetTagFldVal mColTagFlds, GetTgStrFrRec(mRSTgFld, "RmCd", "RmCtg", "D"), "DiaRmCd"
    SetTagFldVal mColTagFlds, GetTgStrFrRec(mRSTgFld, "RmCd", "RmCtg", "C"), "CSRmCd"
    SetTagFldVal mColTagFlds, GetTgStrFrRec(mRSTgFld, "RmCd", "RmCtg", "G"), "GldRmCd"
    SetTagFldVal mColTagFlds, GetTgStrFrRec(mRSTgFld, "RmCd", "RmCtg", "P"), "PltRmCd"
    ' ***** Manali 3.03 - Silver and Palladium included
    SetTagFldVal mColTagFlds, GetTgStrFrRec(mRSTgFld, "RmCd", "RmCtg", "S"), "SlvRmCd"
    SetTagFldVal mColTagFlds, GetTgStrFrRec(mRSTgFld, "RmCd", "RmCtg", "L"), "PdRmCd"
    ' ***** Manali 3.03 - Silver and Palladium included
    SetTagFldVal mColTagFlds, GetTgStrFrRec(mRSTgFld, "RmCd", "RmCtg", "X"), "XRmCd"
    SetTagFldVal mColTagFlds, GetTgStrFrRec(mRSTgFld, "RmCd", "RmCtg", "M"), "MRmCd"      ' ***** Manali 3.9.0 - Oth Met Added
    
' ########################  Manoj #### Ver: 2.0.7 #### Date: 22/07/2004  ########################
    SetTagFldVal mColTagFlds, GetTgStrFrRec(mRSTgFld, "RmIntQly", "RmCtg", "D"), "DiaRmIQ"
    SetTagFldVal mColTagFlds, GetTgStrFrRec(mRSTgFld, "RmIntQly", "RmCtg", "C"), "CSRmIQ"
    SetTagFldVal mColTagFlds, GetTgStrFrRec(mRSTgFld, "RmIntQly", "RmCtg", "G"), "GldRmIQ"
    SetTagFldVal mColTagFlds, GetTgStrFrRec(mRSTgFld, "RmIntQly", "RmCtg", "P"), "PltRmIQ"
    ' ***** Manali 3.03 - Silver and Palladium included
    SetTagFldVal mColTagFlds, GetTgStrFrRec(mRSTgFld, "RmIntQly", "RmCtg", "S"), "SlvRmIQ"
    SetTagFldVal mColTagFlds, GetTgStrFrRec(mRSTgFld, "RmIntQly", "RmCtg", "L"), "PdRmIQ"
    ' ***** Manali 3.03 - Silver and Palladium included
    SetTagFldVal mColTagFlds, GetTgStrFrRec(mRSTgFld, "RmIntQly", "RmCtg", "X"), "XRmIQ"
    SetTagFldVal mColTagFlds, GetTgStrFrRec(mRSTgFld, "RmIntQly", "RmCtg", "M"), "MRmIQ"      ' ***** Manali 3.9.0 - Oth Met Added
' ###########################################  Manoj  ###########################################
    
    If UCase$(adc("TgLevel")) = UCase$("P") Then
      SetTagFldVal mColTagFlds, GetTgStrFrRec(mRSTgFld, "PcDWt", "RmCtg", "D"), "DWRmCd"
      SetTagFldVal mColTagFlds, GetTgStrFrRec(mRSTgFld, "PcCWt", "RmCtg", "C"), "CWRmCd"
      SetTagFldVal mColTagFlds, GetTgStrFrRec(mRSTgFld, "PcGWt", "RmCtg", "G"), "GWRmCd"
      SetTagFldVal mColTagFlds, GetTgStrFrRec(mRSTgFld, "PcPWt", "RmCtg", "P"), "PWRmCd"
      ' ***** Manali 3.03 - Silver and Palladium included
      SetTagFldVal mColTagFlds, GetTgStrFrRec(mRSTgFld, "PcSWt", "RmCtg", "S"), "SWRmCd"
      SetTagFldVal mColTagFlds, GetTgStrFrRec(mRSTgFld, "PcLWt", "RmCtg", "L"), "LWRmCd"
      ' ***** Manali 3.03 - Silver and Palladium included
      SetTagFldVal mColTagFlds, GetTgStrFrRec(mRSTgFld, "PcXWt", "RmCtg", "X"), "XWRmCd"
      SetTagFldVal mColTagFlds, GetTgStrFrRec(mRSTgFld, "PcMWt", "RmCtg", "M"), "MWRmCd"      ' ***** Manali 3.9.0 - Oth Met Added
      
      SetTagFldVal mColTagFlds, GetTgStrFrRec(mRSTgFld, "PcDQty", "RmCtg", "D"), "DQRmCd"
      SetTagFldVal mColTagFlds, GetTgStrFrRec(mRSTgFld, "PcCQty", "RmCtg", "C"), "CQRmCd"
      SetTagFldVal mColTagFlds, GetTgStrFrRec(mRSTgFld, "PcXQty", "RmCtg", "X"), "XQRmCd"
    Else
      If UCase$(adc("TgAvgTyp")) = UCase$("I") Then
        SetTagFldVal mColTagFlds, GetTgStrFrRec(mRSTgFld, "IdDWt", "RmCtg", "D"), "DWRmCd"
        SetTagFldVal mColTagFlds, GetTgStrFrRec(mRSTgFld, "IdCWt", "RmCtg", "C"), "CWRmCd"
        SetTagFldVal mColTagFlds, GetTgStrFrRec(mRSTgFld, "IdGWt", "RmCtg", "G"), "GWRmCd"
        SetTagFldVal mColTagFlds, GetTgStrFrRec(mRSTgFld, "IdPWt", "RmCtg", "P"), "PWRmCd"
        ' ***** Manali 3.03 - Silver and Palladium included
        SetTagFldVal mColTagFlds, GetTgStrFrRec(mRSTgFld, "IdSWt", "RmCtg", "S"), "SWRmCd"
        SetTagFldVal mColTagFlds, GetTgStrFrRec(mRSTgFld, "IdLWt", "RmCtg", "L"), "LWRmCd"
        ' ***** Manali 3.03 - Silver and Palladium included
        SetTagFldVal mColTagFlds, GetTgStrFrRec(mRSTgFld, "IdXWt", "RmCtg", "X"), "XWRmCd"
        SetTagFldVal mColTagFlds, GetTgStrFrRec(mRSTgFld, "IdMWt", "RmCtg", "M"), "MWRmCd"        ' ***** Manali 3.9.0 - Oth Met Added
        
        SetTagFldVal mColTagFlds, GetTgStrFrRec(mRSTgFld, "IdDQty", "RmCtg", "D"), "DQRmCd"
        SetTagFldVal mColTagFlds, GetTgStrFrRec(mRSTgFld, "IdCQty", "RmCtg", "C"), "CQRmCd"
        SetTagFldVal mColTagFlds, GetTgStrFrRec(mRSTgFld, "IdXQty", "RmCtg", "X"), "XQRmCd"
      ElseIf UCase$(adc("TgAvgTyp")) = UCase$("B") Then
        SetTagFldVal mColTagFlds, GetTgStrFrRec(mRSTgFld, "BagDWt", "RmCtg", "D"), "DWRmCd"
        SetTagFldVal mColTagFlds, GetTgStrFrRec(mRSTgFld, "BagCWt", "RmCtg", "C"), "CWRmCd"
        SetTagFldVal mColTagFlds, GetTgStrFrRec(mRSTgFld, "BagGWt", "RmCtg", "G"), "GWRmCd"
        SetTagFldVal mColTagFlds, GetTgStrFrRec(mRSTgFld, "BagPWt", "RmCtg", "P"), "PWRmCd"
        ' ***** Manali 3.03 - Silver and Palladium included
        SetTagFldVal mColTagFlds, GetTgStrFrRec(mRSTgFld, "BagSWt", "RmCtg", "S"), "SWRmCd"
        SetTagFldVal mColTagFlds, GetTgStrFrRec(mRSTgFld, "BagLWt", "RmCtg", "L"), "LWRmCd"
        ' ***** Manali 3.03 - Silver and Palladium included
        SetTagFldVal mColTagFlds, GetTgStrFrRec(mRSTgFld, "BagXWt", "RmCtg", "X"), "XWRmCd"
        SetTagFldVal mColTagFlds, GetTgStrFrRec(mRSTgFld, "BagMWt", "RmCtg", "M"), "MWRmCd"       ' ***** Manali 3.9.0 - Oth Met Added
        
        SetTagFldVal mColTagFlds, GetTgStrFrRec(mRSTgFld, "BagDQty", "RmCtg", "D"), "DQRmCd"
        SetTagFldVal mColTagFlds, GetTgStrFrRec(mRSTgFld, "BagCQty", "RmCtg", "C"), "CQRmCd"
        SetTagFldVal mColTagFlds, GetTgStrFrRec(mRSTgFld, "BagXQty", "RmCtg", "X"), "XQRmCd"
      End If
    End If
  End If
  
  ' VK.25 - Generating Records of Tag Values Which Calculated from Function GetTgStrFrRec() for RrRngSz
  wsSql = GetSubSql("RrRngSz", ps_IdSr, ps_BYy, ps_BChr, ps_BNo)
  Set mRSTgFld = moCn.OpenResultset(wsSql)

  If mRSTgFld.RecCount > 0 Then
    ' Setting Tag Field Value
    SetTagFldVal mColTagFlds, GetTgStrFrRec(mRSTgFld, "RmCd", "RmCtg", "D"), "DRmCdRsz"
    SetTagFldVal mColTagFlds, GetTgStrFrRec(mRSTgFld, "RmCd", "RmCtg", "C"), "CRmCdRsz"
    SetTagFldVal mColTagFlds, GetTgStrFrRec(mRSTgFld, "RrRngSz", "RmCtg", "D"), "DiaRsz"
    SetTagFldVal mColTagFlds, GetTgStrFrRec(mRSTgFld, "RrRngSz", "RmCtg", "C"), "CSRsz"
    
    If UCase$(adc("TgLevel")) = UCase$("P") Then
      SetTagFldVal mColTagFlds, GetTgStrFrRec(mRSTgFld, "PcDWt", "RmCtg", "D"), "DWRsz"
      SetTagFldVal mColTagFlds, GetTgStrFrRec(mRSTgFld, "PcCWt", "RmCtg", "C"), "CWRsz"
      SetTagFldVal mColTagFlds, GetTgStrFrRec(mRSTgFld, "PcDQty", "RmCtg", "D"), "DQRsz"
      SetTagFldVal mColTagFlds, GetTgStrFrRec(mRSTgFld, "PcCQty", "RmCtg", "C"), "CQRsz"
    Else
      If UCase$(adc("TgAvgTyp")) = UCase$("B") Then
        SetTagFldVal mColTagFlds, GetTgStrFrRec(mRSTgFld, "BagDWt", "RmCtg", "D"), "DWRsz"
        SetTagFldVal mColTagFlds, GetTgStrFrRec(mRSTgFld, "BagCWt", "RmCtg", "C"), "CWRsz"
        SetTagFldVal mColTagFlds, GetTgStrFrRec(mRSTgFld, "BagDQty", "RmCtg", "D"), "DQRsz"
        SetTagFldVal mColTagFlds, GetTgStrFrRec(mRSTgFld, "BagCQty", "RmCtg", "C"), "CQRsz"
      End If
    End If
  End If
  wsSql = GetSubSql("RmStnCls", ps_IdSr, ps_BYy, ps_BChr, ps_BNo)
  Set mRSTgFld = moCn.OpenResultset(wsSql)
  If mRSTgFld.RecCount > 0 Then
    ' Setting Tag Field Value
    SetTagFldVal mColTagFlds, GetTgStrFrRec(mRSTgFld, "RmStnCls", "RmCtg", "D"), "DStnCls"
    SetTagFldVal mColTagFlds, GetTgStrFrRec(mRSTgFld, "RmStnCls", "RmCtg", "C"), "CStnCls"
'    SetTagFldVal mColTagFlds, GetTgStrFrRec(mRSTgFld, "RrRngSz", "RmCtg", "D"), "DiaRsz"
 '   SetTagFldVal mColTagFlds, GetTgStrFrRec(mRSTgFld, "RrRngSz", "RmCtg", "C"), "CSRsz"
    
    If UCase$(adc("TgLevel")) = UCase$("P") Then
      SetTagFldVal mColTagFlds, GetTgStrFrRec(mRSTgFld, "PcDWt", "RmCtg", "D"), "DWStnCls"
      SetTagFldVal mColTagFlds, GetTgStrFrRec(mRSTgFld, "PcCWt", "RmCtg", "C"), "CWStnCls"
      SetTagFldVal mColTagFlds, GetTgStrFrRec(mRSTgFld, "PcDQty", "RmCtg", "D"), "DQStnCls"
      SetTagFldVal mColTagFlds, GetTgStrFrRec(mRSTgFld, "PcCQty", "RmCtg", "C"), "CQStnCls"
    Else
      If UCase$(adc("TgAvgTyp")) = UCase$("B") Then
        SetTagFldVal mColTagFlds, GetTgStrFrRec(mRSTgFld, "BagDWt", "RmCtg", "D"), "DWStnCls"
        SetTagFldVal mColTagFlds, GetTgStrFrRec(mRSTgFld, "BagCWt", "RmCtg", "C"), "CWStnCls"
        SetTagFldVal mColTagFlds, GetTgStrFrRec(mRSTgFld, "BagDQty", "RmCtg", "D"), "DQStnCls"
        SetTagFldVal mColTagFlds, GetTgStrFrRec(mRSTgFld, "BagCQty", "RmCtg", "C"), "CQStnCls"
      End If
    End If
  End If

End Sub
Private Sub SetTagFldVal(ByRef pColVal As Collection, ByVal pValue, Optional ByVal pKey)
  ' According to datatype of TagField Value is Converted & Added to Collection
  If isExists(gColTgFldTyp, pKey) = True Then
    Select Case UCase$(gColTgFldTyp(pKey))
    Case UCase$("C")
      AddItem pColVal, CStr(pValue & ""), pKey
    Case UCase$("U")
      AddItem pColVal, UCase$(CStr(pValue & "")), pKey
    Case UCase$("I")
      AddItem pColVal, CLng("0" & pValue), pKey
    Case UCase$("F")
      AddItem pColVal, CDbl("0" & pValue), pKey
    Case UCase$("B")
      AddItem pColVal, CBool(IIF(pValue = "", "0" & "", pValue)), pKey
    End Select
  Else
    AddItem pColVal, pValue & "", pKey
  End If
  
End Sub
Private Function GetTgStrFrRec(ByRef po_RS As MwfLib.MDORowSet, ByVal ps_Fld As String, Optional ByVal ps_CndFld As String, Optional ByVal ps_CndVal As String) As String
  'Returning Concanated Column Values Sepearated by "/"
  Dim woRS As New MwfLib.MDORowSet
  Dim wsVal As String
  wsVal = ""
  GetTgStrFrRec = ""
  Set woRS = po_RS
  If woRS.RecCount > 0 Then
    With woRS
      .MoveFirst
      Do While Not (.EOF Or .BOF)
        If UCase$(woRS.FldValue(ps_CndFld)) = UCase$(ps_CndVal) Then wsVal = wsVal & Trim$(woRS.FldValue(ps_Fld)) & "/"
        .MoveNext
      Loop
    End With
    If wsVal <> "" Then GetTgStrFrRec = Left$(wsVal, Len(wsVal) - 1)
  End If
End Function
Private Function GetMainSql(ps_IdSr As Double, ps_BYy As String, ps_BChr As String, ps_BNo As Long) As String
  ' Setting MainSQL for Setting Tag Values
  Dim wSqlStr As String
  Dim wSqlGrp As String, wSqlCnd As String, wSqlJoin As String
  Dim wBagFlds As String, wInvFlds As String, wPcFlds As String, wSqlFlds As String
  'Common Sql Fields
  'VK.25 - Changed Order related fields instead of Invoice for jEmr
  If UCase(adc.MenuCd) = UCase("RepJtTagPrn") Then
    wSqlFlds = " OmTc As InTc, OmYy As InYy, OmChr As InChr, OmNo As InNo, " + _
               " OmTc + '/' + OmYy + '/' + OmChr + '/' + Cast(OmNo As VarChar(7)) As InvNo, "
  Else
    wSqlFlds = " InTc, InYy, InChr, InNo, InTc + '/' + InYy + '/' + InChr + '/' + Cast(InNo As VarChar(7)) As InvNo,"
  End If
  
  'Commom Peiece Fields
  ' ***** Manali 3.03 - Silver and Palladium included\
  ' ***** Manali 3.8.0 - DC Wt & Qty added
  ' ***** Manali 3.9.0 - Oth Met Added
  wPcFlds = GetQWSumSql("D", "Q", "P") + "," + GetQWSumSql("C", "Q", "P") + "," + GetQWSumSql("DC", "Q", "P") + "," + _
            GetQWSumSql("D", "W", "P") + "," + GetQWSumSql("C", "W", "P") + "," + GetQWSumSql("DC", "W", "P") + "," + _
            GetQWSumSql("G", "W", "P") + "," + GetQWSumSql("P", "W", "P") + "," + _
            GetQWSumSql("S", "W", "P") + "," + GetQWSumSql("L", "W", "P") + "," + _
            GetQWSumSql("X", "Q", "P") + "," + GetQWSumSql("X", "W", "P") + "," + _
            GetQWSumSql("M", "W", "P") + "," + _
            GetQWSumSql("", "G", "P") + "," + GetValSql("P")
  
  'Common Condition for Sql Statement
  'VK.25 - Changed Order related fields instead of Invoice for jEmr
  If UCase(adc.MenuCd) = UCase("RepJtTagPrn") Then
    wSqlCnd = " WHERE OmCoCd = '" + adc("TgCoCd") + "' And OmTc ='" + adc("TgTc") + "' And OmYy = '" + adc("TgYy") + "' And" + _
            " OmChr = '" + adc("TgChr") + "' And OmNo = " + CStr(adc("TgNo"))
  Else
    wSqlCnd = " WHERE InCoCd = '" + adc("TgCoCd") + "' And InTc ='" + adc("TgTc") + "' And InYy = '" + adc("TgYy") + "' And" + _
            " InChr = '" + adc("TgChr") + "' And InNo = " + CStr(adc("TgNo"))
  End If
  
  'Common Grouping Options
  'VK.25 - Changed Order related fields instead of Invoice for jEmr
  If UCase(adc.MenuCd) = UCase("RepJtTagPrn") Then
    wSqlGrp = " Group By OmCoCd, OmTc, OmYy, OmChr, OmNo"
  Else
    wSqlGrp = " Group By InCoCd, InTc, InYy, InChr, InNo"
  End If
  
  Dim wRmJoin As String   ' ***** Manali 3.10.0 - 03/04/12 - Copy Wts From DsgMst
  
  Select Case UCase$(adc("TgAvgTyp"))
    Case UCase$("B")
      'Sql Field Defintions for Bag Records
' ########################  Manoj #### Ver: 2.0.7 #### Date: 22/07/2004  ########################
'      wSqlFlds = wSqlFlds + _
'                " IfSr  as IdSr, InTc + '/' + InYy + '/' + InChr + '/' + Cast(InNo As VarChar(5)) + '/' + Cast(IfSr As VarChar(5)) As InvSr, " + _
'                " IfFdBYy AS BYy, IfFdBChr AS BChr, IfFdBNo AS BNo, IfFdBYy + '/' + IfFdBChr + '/' + Cast(IfFdBNo As VarChar(5)) AS BagNo," + _
'                " Max(FdDmCd) As DmCd, Max(FdDmSz) As DmSz, Max(FdSfx) As DmSfx, Max(DmCtg) As DmCtg, Max(IdKt) As IdKt," + _
'                " Max(OdDmCol) As DmCol, Max(InExpNo) As InExpNo, Max(InCmCd) As CmCd, isNull((Select CdCustDmCd from CustDsg" + _
'                " Where CdCtg ='C' And CdCd = Max(InCmCd) And CdOurDmCD = Max(FdDmCd) And CdOurSfx = Max(FdSfx) And" + _
'                " CdOurDmSz = Max(FdDmSz)), '')  AS CdCustDmCd,"
'      wSqlFlds = wSqlFlds + _
'                " IfSr  as IdSr, InTc + '/' + InYy + '/' + InChr + '/' + Cast(InNo As VarChar(5)) + '/' + Cast(IfSr As VarChar(5)) As InvSr, " + _
'                " IfFdBYy AS BYy, IfFdBChr AS BChr, IfFdBNo AS BNo, IfFdBYy + '/' + IfFdBChr + '/' + Cast(IfFdBNo As VarChar(5)) AS BagNo," + _
'                " Max(FdExpOdTc) as ExpOdTc, Max(FdExpOdYy) as ExpOdYy, Max(FdExpOdChr) as ExpOdChr, Max(FdExpOdNo) as ExpOdNo," + _
'                " Max(FdExpOdTc) + '/' + Max(FdExpOdYy) + '/' + Max(FdExpOdChr) + '/' + Cast(Max(FdExpOdNo) As VarChar(5)) As ExpOrdNo, Max(FdExpOdSr) as ExpOdSr," + _
'                " Max(FdExpOdTc) + '/' + Max(FdExpOdYy) + '/' + Max(FdExpOdChr) + '/' + Cast(Max(FdExpOdNo) As VarChar(5)) + '/' + Cast(Max(FdExpOdSr) As VarChar(5))As ExpOrdSr," + _
'                " Max(FdPrdOdTc) as PrdOdTc, Max(FdPrdOdYy) as PrdOdYy, Max(FdPrdOdChr) as PrdOdChr, Max(FdPrdOdNo) as PrdOdNo," + _
'                " Max(FdPrdOdTc) + '/' + Max(FdPrdOdYy) + '/' + Max(FdPrdOdChr) + '/' + Cast(Max(FdPrdOdNo) As VarChar(5)) As PrdOrdNo, Max(FdPrdOdSr) as PrdOdSr," + _
'                " Max(FdPrdOdTc) + '/' + Max(FdPrdOdYy) + '/' + Max(FdPrdOdChr) + '/' + Cast(Max(FdPrdOdNo) As VarChar(5)) + '/' + Cast(Max(FdPrdOdSr) As VarChar(5))As PrdOrdSr," + _
'                " Max(FdDmCd) As DmCd, Max(FdDmSz) As DmSz, Max(FdSfx) As DmSfx, Max(DmCtg) As DmCtg, Max(IdKt) As IdKt," + _
'                " Max(OdDmCol) As DmCol, Max(InExpNo) As InExpNo, Max(InCmCd) As CmCd, isNull((Select CdCustDmCd from CustDsg" + _
'                " Where CdCtg ='C' And CdCd = Max(InCmCd) And CdOurDmCD = Max(FdDmCd) And CdOurSfx = Max(FdSfx) And" + _
'                " CdOurDmSz = Max(FdDmSz)), '')  AS CdCustDmCd,"
      
      ' ***** Manali 3.10.0 - 14/05/12 - Varchar for FdExpOdNo changed from 5 to 6 and for IfFdBNo 5 to 7       4.1.4
      '578.45 huidno added
      wSqlFlds = wSqlFlds + _
                " IfSr as IdSr, InTc + '/' + InYy + '/' + InChr + '/' + Cast(InNo As VarChar(7)) + '/' + Cast(IfSr As VarChar(7)) As InvSr, " + _
                " '' AS BCoCd, IfFdBYy AS BYy, IfFdBChr AS BChr, IfFdBNo AS BNo, IfFdBYy + '/' + IfFdBChr + '/' + Cast(IfFdBNo As VarChar(7)) AS BagNo," + _
                " Max(FdExpOdTc) as ExpOdTc, Max(FdExpOdYy) as ExpOdYy, Max(FdExpOdChr) as ExpOdChr, Max(FdExpOdNo) as ExpOdNo," + _
                " Max(FdExpOdTc) + '/' + Max(FdExpOdYy) + '/' + Max(FdExpOdChr) + '/' + Cast(Max(FdExpOdNo) As VarChar(6)) As ExpOrdNo, Max(FdExpOdSr) as ExpOdSr," + _
                " Max(FdExpOdTc) + '/' + Max(FdExpOdYy) + '/' + Max(FdExpOdChr) + '/' + Cast(Max(FdExpOdNo) As VarChar(6)) + '/' + Cast(Max(FdExpOdSr) As VarChar(5))As ExpOrdSr," + _
                " Max(OmPoNo)as OmPoNo, Max(FdDmCd) As DmCd, Max(FdDmSz) As DmSz, Max(FdSfx) As DmSfx, Max(DmCtg) As DmCtg," + IIF(adc("TgSrcDsg") = "Y", "Max(DmKt) ", "Max(IdKt)") + " As IdKt," + _
                IIF(adc("TgSrcDsg") = "Y", "Max(DmCol) ", " Max(OdDmCol)") + " As DmCol, Max(OdPoNo) as OdPoNo, Max(OdMrp) as OdMrp, Max(InExpNo) As InExpNo, Max(InCmCd) As CmCd, isNull((Select CdCustDmCd from CustDsg" + _
                " Where CdPrtKey='" + ctCurrPrtn + "' and CdCtg ='C' And CdCd = Max(InCmCd) And CdOurDmCD = Max(FdDmCd) And CdOurSfx = Max(FdSfx) And" + _
                " CdOurDmSz = Max(FdDmSz)), '')  AS CdCustDmCd," + _
                "(select isnull(CdCertNo,'') from CertDet where CdCtg='H' and CdBCoCd =InCoCd AND CdBYy =IfFdBYy and CdBChr =IfFdBChr and CdBNo =IfFdBNo) as HuidNo,"

' ###########################################  Manoj  ###########################################
      
      'Sql Inoice Field Defintions for Bag Records
      ' Zubin
      ' ***** Manali 3.8.0 - DC, S, L Wt & Qty added
      ' ***** Manali 3.9.0 - Oth Met Added
      wInvFlds = " 0 As IdQty, 0 As IdDQty, 0 As IdCQty, 0 as IdDCQty, 0 As IdDWt, 0 As IdCWt, 0 as IdDCWt, " + _
                 " 0 As IdGWt, 0 As IdPWt, 0 As IdSWt, 0 As IdLWt, 0 As IdXQty, 0 As IdXWt, 0 as IdMWt, 0 as IdGrsWt, 0 As IdVal, 0 as IdMuVal, "
      
      'Sql Bag Field Defintions for Bag Records
      ' ***** Manali 3.03 - Silver and Palladium included
      ' ***** Manali 3.8.0 - DC Wt & Qty added
      ' ***** Manali 3.9.0 - Oth Met Added
      wBagFlds = " Max(FdQty) as BagQty," + _
                  GetQWSumSql("D", "Q", "B") + "," + GetQWSumSql("C", "Q", "B") + "," + GetQWSumSql("DC", "Q", "B") + "," + _
                 GetQWSumSql("D", "W", "B") + "," + GetQWSumSql("C", "W", "B") + "," + GetQWSumSql("DC", "W", "B") + "," + _
                 GetQWSumSql("G", "W", "B") + "," + GetQWSumSql("P", "W", "B") + "," + _
                 GetQWSumSql("S", "W", "B") + "," + GetQWSumSql("L", "W", "B") + "," + _
                 GetQWSumSql("X", "Q", "B") + "," + GetQWSumSql("X", "W", "B") + "," + _
                 GetQWSumSql("M", "W", "B") + "," + _
                 GetQWSumSql("", "G", "B") + "," + GetValSql("B") + ","
      
  ' ***** Manali 3.10.0 - 03/04/12 - Copy Wts From DsgMst
  If adc("TgSrcDsg") = "Y" Then
  wRmJoin = " Join DsgRm On DrPrtKey=FdPrtKey and DrTcTyp=FdDmTcTyp and DrCd=FdDmCd and DrSz=isNull((Select DmSz From DsgMst where DmPrtKey=FdPrtKey and DmTcTyp=FdDmTcTyp and DmCd=FdDmCd and DmSz=FdDmSz),'') " + _
           " INNER JOIN RmMst ON DrPrtKey=RmPrtKey and DrRmCd = RmCd  "
  Else
  ' ***** Manali 3.10.0 - 03/04/12 - Copy Wts From DsgMst
  wRmJoin = " INNER JOIN FgRm ON FdPrtKey=FrPrtKey and FdCoCd = FrCoCd And FdTc = FrTc AND FdYy = FrYy AND FdChr = FrChr AND FdNo = FrNo AND FdSr = FrSr" + _
           " INNER JOIN RmMst ON FrPrtKey=RmPrtKey and FrRmCd = RmCd  "
  End If
      
      'Sql Join for Bag Records
      wSqlJoin = " From InvHd INNER Join InvFgd On InPrtKey=IfPrtKey and InCoCd = IfCoCd And InTc = IfTc And InYy =IfYy And InChr = IfChr And InNo = IfNo  " + _
                " INNER JOIN Fgd ON IfPrtKey=FdPrtKey and IfCoCd = FdCoCd And IfFdTc = FdTc AND IfFdYy = FdYy AND IfFdChr = FdChr AND IfFdNo = FdNo AND IfFdSr = FdSr" + _
                wRmJoin + _
                " INNER JOIN DsgMst ON DmPrtKey=FdPrtKey and DmIdNo=FdDmIdNo and DmTcTyp =FdDmTcTyp And DmCd = FdDmCd And DmSz = ''" + _
                " INNER JOIN OrdDsg ON OdPrtKey=FdPrtKey and OdCoCd=FdCoCd  and OdTc = FdExpOdTc and OdYy = FdExpOdYy" + _
                " And OdChr=FdExpOdChr and OdNo= FdExpOdNo and OdSr=FdExpOdSr" + _
                " INNER JOIN OrdMst ON OmCoCd= OdCoCd And OmTc = OdTc And OmYy = OdYy And OmChr = OdChr And OmNo = OdNo" + _
                " LEFT OUTER JOIN InvDsg On InPrtKey=IdPrtKey and IdInIdNo=InIdNo and InCoCd = IdCoCd And InTc = IdTc And InYy =IdYy And InChr = IdChr And InNo = IdNo And IfSr = IdSr"
      
      'Sql Condition for Bag Records
      wSqlCnd = wSqlCnd + " And IfSr = " + CStr(ps_IdSr) + " And IfFdBYy = '" + CStr(ps_BYy) + _
                "' And IfFdBChr = '" + CStr(ps_BChr) + "' And IfFdBNo = " + CStr(ps_BNo)
      
      'Sql Grouping Fieldds for Bag Records
      wSqlGrp = wSqlGrp + ", IfSr, IfFdBYy, IfFdBChr, IfFdBNo"
    
    Case UCase$("I")
      'Sql Field Defintions for InvDsg Records
' ########################  Manoj #### Ver: 2.0.7 #### Date: 22/07/2004  ########################
'      wSqlFlds = wSqlFlds + _
'                 " IdSr, InTc + '/' + InYy + '/' + InChr + '/' + Cast(InNo As VarChar(5)) + '/' + Cast(IdSr As VarChar(5)) As InvSr, " + _
'                 " '' AS BYy, '' AS BChr, 0 AS BNo, '' AS BagNo, " + _
'                 " Max(IdDmCd) As DmCd, Max(IdDmSz) As DmSz, Max(IdSfx) As DmSfx, Max(DmCtg) As DmCtg, Max(IdKt) As IdKt, " + _
'                 " Max(DmCol) As DmCol, Max(InExpNo) As InExpNo, Max(InCmCd) As CmCd, isNull((Select CdCustDmCd from CustDsg " + _
'                 " Where CdCtg ='C' And CdCd = Max(InCmCd) And CdOurDmCd = Max(IdDmCd) And CdOurSfx = Max(IdSfx) And " + _
'                 " CdOurDmSz = Max(IdDmSz)), '')  AS CdCustDmCd,"
'      wSqlFlds = wSqlFlds + _
'                 " IdSr, InTc + '/' + InYy + '/' + InChr + '/' + Cast(InNo As VarChar(5)) + '/' + Cast(IdSr As VarChar(5)) As InvSr, " + _
'                 " '' AS BYy, '' AS BChr, 0 AS BNo, '' AS BagNo, " + _
'                 " Max(IdExpOdTc) as ExpOdTc, Max(IdExpOdYy) as ExpOdYy, Max(IdExpOdChr) as ExpOdChr, Max(IdExpOdNo) as ExpOdNo," + _
'                 " Max(IdExpOdTc) + '/' + Max(IdExpOdYy) + '/' + Max(IdExpOdChr) + '/' + Cast(Max(IdExpOdNo) As VarChar(5)) As ExpOrdNo, Max(IdExpOdSr) as ExpOdSr," + _
'                 " Max(IdExpOdTc) + '/' + Max(IdExpOdYy) + '/' + Max(IdExpOdChr) + '/' + Cast(Max(IdExpOdNo) As VarChar(5)) + '/' + Cast(Max(IdExpOdSr) As VarChar(5))As ExpOrdSr," + _
'                 " '' as PrdOdTc, '' as PrdOdYy, '' as PrdOdChr, 0 as PrdOdNo, '' As PrdOrdNo, 0 as PrdOdSr, '' as PrdOrdSr," + _
'                 " Max(IdDmCd) As DmCd, Max(IdDmSz) As DmSz, Max(IdSfx) As DmSfx, Max(DmCtg) As DmCtg, Max(IdKt) As IdKt, " + _
'                 " Max(DmCol) As DmCol, Max(InExpNo) As InExpNo, Max(InCmCd) As CmCd, isNull((Select CdCustDmCd from CustDsg " + _
'                 " Where CdCtg ='C' And CdCd = Max(InCmCd) And CdOurDmCd = Max(IdDmCd) And CdOurSfx = Max(IdSfx) And " + _
'                 " CdOurDmSz = Max(IdDmSz)), '')  AS CdCustDmCd,"
      '4.1.4
      'VK.25 - Changed Order related fields instead of Invoice for jEmr
      If UCase(adc.MenuCd) = UCase("RepJtTagPrn") Then
        wSqlFlds = wSqlFlds + _
                 " OdSr As IdSr, OmTc + '/' + OmYy + '/' + OmChr + '/' + Cast(OmNo As VarChar(7)) + '/' + Cast(OdSr As VarChar(7)) As InvSr, " + _
                 " max(OdDmStkCoCd) BCoCd, max(OdDmStkYy) BYy, max(OdDmStkChr) BChr, max(OdDmStkNo) BNo, " + _
                 " case when max(OdDmStkNo)>0 then max(OdDmStkCoCd) + '/' + max(OdDmStkYy) + '/' + max(OdDmStkChr) + '/' + Cast(max(OdDmStkNo) As VarChar(7)) Else '' End As BagNo, " + _
                 " Max(OdTc) as ExpOdTc, Max(OdYy) as ExpOdYy, Max(OdChr) as ExpOdChr, Max(OdNo) as ExpOdNo," + _
                 " Max(OdTc) + '/' + Max(OdYy) + '/' + Max(OdChr) + '/' + Cast(Max(OdNo) As VarChar(5)) As ExpOrdNo, Max(OdSr) as ExpOdSr," + _
                 " Max(OdTc) + '/' + Max(OdYy) + '/' + Max(OdChr) + '/' + Cast(Max(OdNo) As VarChar(5)) + '/' + Cast(Max(OdSr) As VarChar(5))As ExpOrdSr," + _
                 " Max(OmPoNo) As OmPoNo, Max(OdDmCd) As DmCd, Max(OdDmSz) As DmSz, Max(OdSfx) As DmSfx, Max(DmCtg) As DmCtg," + IIF(adc("TgSrcDsg") = "Y", "Max(DmKt)", "Max(OdKt)") + " As IdKt, " + _
                 IIF(adc("TgSrcDsg") = "Y", "Max(DmCol)", "Max(OdDmCol)") + " As DmCol, Max(OdPoNo) as OdPoNo, Max(OdMrp) as OdMrp, '' As InExpNo, Max(OmCmCd) As CmCd, isNull((Select CdCustDmCd from CustDsg " + _
                 " Where CdPrtKey='" + ctCurrPrtn + "' and CdCtg ='C' And CdCd = Max(OmCmCd) And CdOurDmCd = Max(OdDmCd) And CdOurSfx = Max(OdSfx) And " + _
                 " CdOurDmSz = Max(OdDmSz)), '')  AS CdCustDmCd, '' as HuidNo,"
      Else
        wSqlFlds = wSqlFlds + _
                 " IdSr, InTc + '/' + InYy + '/' + InChr + '/' + Cast(InNo As VarChar(7)) + '/' + Cast(IdSr As VarChar(7)) As InvSr, " + _
                 " '' As BCoCd, '' AS BYy, '' AS BChr, 0 AS BNo, '' AS BagNo, " + _
                 " Max(IdExpOdTc) as ExpOdTc, Max(IdExpOdYy) as ExpOdYy, Max(IdExpOdChr) as ExpOdChr, Max(IdExpOdNo) as ExpOdNo," + _
                 " Max(IdExpOdTc) + '/' + Max(IdExpOdYy) + '/' + Max(IdExpOdChr) + '/' + Cast(Max(IdExpOdNo) As VarChar(5)) As ExpOrdNo, Max(IdExpOdSr) as ExpOdSr," + _
                 " Max(IdExpOdTc) + '/' + Max(IdExpOdYy) + '/' + Max(IdExpOdChr) + '/' + Cast(Max(IdExpOdNo) As VarChar(5)) + '/' + Cast(Max(IdExpOdSr) As VarChar(5))As ExpOrdSr," + _
                 " Max(OmPoNo) As OmPoNo, Max(IdDmCd) As DmCd, Max(IdDmSz) As DmSz, Max(IdSfx) As DmSfx, Max(DmCtg) As DmCtg," + IIF(adc("TgSrcDsg") = "Y", "Max(DmKt)", "Max(IdKt)") + " As IdKt, " + _
                 IIF(adc("TgSrcDsg") = "Y", "Max(DmCol)", "Max(OdDmCol)") + " As DmCol, Max(OdPoNo) as OdPoNo, Max(OdMrp) as OdMrp, Max(InExpNo) As InExpNo, Max(InCmCd) As CmCd, isNull((Select CdCustDmCd from CustDsg " + _
                 " Where CdPrtKey='" + ctCurrPrtn + "' and CdCtg ='C' And CdCd = Max(InCmCd) And CdOurDmCd = Max(IdDmCd) And CdOurSfx = Max(IdSfx) And " + _
                 " CdOurDmSz = Max(IdDmSz)), '')  AS CdCustDmCd,'' as HuidNo,"
      End If
                 
      ' ###########################################  Manoj  ###########################################
    
      'Sql Inoice Field Defintions for InvDsg Records
      ' ***** Manali 3.03 - Silver and Palladium included
      ' ***** Manali 3.8.0 - DC Wt & Qty added
      ' ***** Manali 3.9.0 - Oth Met Added
      'VK.25 - Changed Order related fields instead of Invoice for jEmr
      wInvFlds = IIF(UCase(adc.MenuCd) = UCase("RepJtTagPrn"), " Max(OdOrdQty) ", " Max(IdQty) ") + " As IdQty," + _
                 GetQWSumSql("D", "Q", "I") + "," + GetQWSumSql("C", "Q", "I") + "," + GetQWSumSql("DC", "Q", "I") + "," + _
                 GetQWSumSql("D", "W", "I") + "," + GetQWSumSql("C", "W", "I") + "," + GetQWSumSql("DC", "W", "I") + "," + _
                 GetQWSumSql("G", "W", "I") + "," + GetQWSumSql("P", "W", "I") + "," + _
                 GetQWSumSql("S", "W", "I") + "," + GetQWSumSql("L", "W", "I") + "," + _
                 GetQWSumSql("X", "Q", "I") + "," + GetQWSumSql("X", "W", "I") + "," + _
                 GetQWSumSql("M", "W", "I") + "," + _
                 GetQWSumSql("", "G", "I") + "," + GetValSql("I") + ","
      
        'Sql Bag Field Defintions for InvDsg Records
        ' Zubin
        ' ***** Manali 3.8.0 - DC, S, L Wt & Qty added
        ' ***** Manali 3.9.0 - Oth Met Added
        wBagFlds = " 0 As BagQty, 0 As BagDQty, 0 As BagCQty, 0 As BagDCQty, 0 As BagDWt, 0 As BagCWt, 0 As BagDCWt," + _
                 " 0 As BagGWt, 0 As BagPWt, 0 As BagSWt, 0 As BagLWt, 0 As BagXQty, 0 As BagXWt, 0 as BagMWt, 0 as BagGrsWt, 0 As BagVal, 0 As BagMuVal, "
                       
      ' ***** Manali 3.10.0 - 03/04/12 - Copy Wts From DsgMst
      'VK.25 - Changed Order related fields instead of Invoice for jEmr
      If UCase(adc.MenuCd) = UCase("RepJtTagPrn") Then
        If adc("TgSrcDsg") = "Y" Then
          wRmJoin = " Join DsgRm On DrPrtKey=OdPrtKey and DrTcTyp=OdDmTcTyp and DrCd=OdDmCd and DrSz=isNull((Select DmSz From DsgMst where DmPrtKey=OdPrtKey and DmTcTyp=OdDmTcTyp and DmCd=OdDmCd and DmSz=OdDmSz),'') " + _
                    " INNER JOIN RmMst ON DrPrtKey=RmPrtKey and DrRmCd = RmCd  "
        Else
          wRmJoin = " Inner Join OrdRm ON OdPrtKey=OrPrtKey and OdCoCd = OrCoCd And OdTc = OrTc And OdYy =OrYy And" + _
                    " OdChr = OrChr And OdNo = OrNo And OdSr = OrSr INNER JOIN RmMst On OrRmCd = RmCd "
        End If
          wSqlJoin = " From OrdMst INNER Join OrdDsg ON OmPrtKey=OdPrtKey and OmCoCd = OdCoCd And OmTc = OdTc And OmYy =OdYy And OmChr = OdChr And OmNo = OdNo" + _
                       wRmJoin + _
                     " INNER JOIN DsgMst ON DmPrtKey=OdPrtKey and DmTcTyp= OdDmTcTyp And DmCd = OdDmCd And DmSz = ''"
          wSqlCnd = wSqlCnd + " And OdSr = " + CStr(ps_IdSr)
          wSqlGrp = wSqlGrp + ", OdSr"
      Else
        If adc("TgSrcDsg") = "Y" Then
          wRmJoin = " Join DsgRm On DrPrtKey=IdPrtKey and DrTcTyp='DM' and DrCd=IdDmCd and DrSz=isNull((Select DmSz From DsgMst where DmPrtKey=IdPrtKey and DmTcTyp='DM' and DmCd=IdDmCd and DmSz=IdDmSz),'') " + _
                    " INNER JOIN RmMst ON DrPrtKey=RmPrtKey and DrRmCd = RmCd  "
        Else
          ' ***** Manali 3.10.0 - 03/04/12 - Copy Wts From DsgMst
          wRmJoin = " Inner Join InvRm ON IdPrtKey=IrPrtKey and IdCoCd = IrCoCd And IdTc = IrTc And IdYy =IrYy And" + _
                    " IdChr = IrChr And IdNo = IrNo And IdSr = IrSr INNER JOIN RmMst On IrRmCd = RmCd "
        End If
      
        'Sql Join for Bag Records
        wSqlJoin = " From InvHd INNER Join InvDsg ON InPrtKey=IdPrtKey and InCoCd = IdCoCd And InTc = IdTc And InYy =IdYy And InChr = IdChr And InNo = IdNo" + _
                     wRmJoin + _
                   " INNER JOIN OrdDsg On OdPrtKey=IdPrtKey and OdCoCd= IdCoCd and OdTc= IdExpOdTc and OdYy= IdExpOdYy " + _
                   " and OdChr= IdExpOdChr and OdNo= IdExpOdNo and OdSr= IdExpOdSr " + _
                   " INNER JOIN OrdMst ON OmPrtKey=OdPrtKey and OmIdNo=OdOmIdNo and OmCoCd= OdCoCd And OmTc = OdTc And OmYy = OdYy And OmChr = OdChr And OmNo = OdNo" + _
                   " INNER JOIN DsgMst ON DmPrtKey=IdPrtKey and DmTcTyp= 'DM' And DmCd = IdDmCd And DmSz = ''"
      
        'Sql Condition for InvDsg Records
        wSqlCnd = wSqlCnd + " And IdSr = " + CStr(ps_IdSr)
      
        'Sql Grouping Fields for InvDsg Records
        wSqlGrp = wSqlGrp + ", IdSr"
      End If

  End Select
  
  ' Final SQL Statement
  wSqlStr = "Select" + wSqlFlds + wInvFlds + wBagFlds + wPcFlds + wSqlJoin + wSqlCnd + wSqlGrp
  GetMainSql = wSqlStr
End Function

Private Function GetSubSql(ByVal ps_Opt As String, ps_IdSr As Double, ps_BYy As String, ps_BChr As String, ps_BNo As Long) As String
  ' Setting Second SQL for Setting Tag Values
  Dim wSqlStr As String
  Dim wSqlGrp As String, wSqlCnd As String, wSqlJoin As String
  Dim wSqlFlds  As String, wPcFlds As String, wSqlBIFlds As String
  Dim wRmJoin As String
  
  'VK.25 - Added RngSz code for Source as DsgMst
  Dim wSqlRngSz As String
  If adc("TgSrcDsg") = "Y" Then
    wSqlRngSz = " Left Outer Join RmRt On RrTcTyp = 'RSZ' And RrCmCd = '" + ctSelfCmCd + "' And " + _
                "RrCtg = RmCtg And RrSCtg = RmSCtg And RrFrLn <= DrLn1 And RrToLn >= DrLn1 "
  End If
  
  'Sql Common Pcwise Fields
  ' ***** Manali 3.03 - Silver and Palladium included
  ' ***** Manali 3.8.0 - DC Wt & Qty added
  ' ***** Manali 3.9.0 - Oth Met Added
  wPcFlds = GetQWSumSql("D", "Q", "P") + "," + GetQWSumSql("C", "Q", "P") + "," + GetQWSumSql("DC", "Q", "P") + "," + _
            GetQWSumSql("D", "W", "P") + "," + GetQWSumSql("C", "W", "P") + "," + GetQWSumSql("DC", "W", "P") + "," + _
            GetQWSumSql("G", "W", "P") + "," + GetQWSumSql("P", "W", "P") + "," + _
            GetQWSumSql("S", "W", "P") + "," + GetQWSumSql("L", "W", "P") + "," + _
            GetQWSumSql("X", "Q", "P") + "," + GetQWSumSql("X", "W", "P") + "," + _
            GetQWSumSql("M", "W", "P")
  wSqlCnd = " Where 1 = 1"
  
  Select Case UCase$(adc("TgAvgTyp"))
    Case UCase$("B")
      'Sql Fields for Bag Records
      wSqlFlds = " IfTc, IfYy, IfChr, IfNo, IfSr, IfFdBYy, IfFdBChr, IfFdBNo,"
      
      'Sql Fields for Bag Records
      ' ***** Manali 3.03 - Silver and Palladium included
      ' ***** Manali 3.8.0 - DC Wt & Qty added
      ' ***** Manali 3.9.0 - Oth Met Added
      wSqlBIFlds = GetQWSumSql("D", "Q", "B") + "," + GetQWSumSql("C", "Q", "B") + "," + GetQWSumSql("DC", "Q", "B") + "," + _
                  GetQWSumSql("D", "W", "B") + "," + GetQWSumSql("C", "W", "B") + "," + GetQWSumSql("DC", "W", "B") + "," + _
                  GetQWSumSql("G", "W", "B") + "," + GetQWSumSql("P", "W", "B") + "," + _
                  GetQWSumSql("S", "W", "B") + "," + GetQWSumSql("L", "W", "B") + "," + _
                  GetQWSumSql("X", "Q", "B") + "," + GetQWSumSql("X", "W", "B") + "," + _
                  GetQWSumSql("M", "W", "B") + ","
      
      'Sql Join for Bag Records
      ' ***** Manali 3.10.0 - 03/04/12 - Copy Wts From DsgMst
      If adc("TgSrcDsg") = "Y" Then
        wRmJoin = " Join DsgRm On DrPrtKey=FdPrtKey and DrTcTyp=FdDmTcTyp and DrCd=FdDmCd and DrSz=isNull((Select DmSz From DsgMst where DmPrtKey=FdPrtKey and DmTcTyp=FdDmTcTyp and DmCd=FdDmCd and DmSz=FdDmSz),'') " + _
                 " INNER JOIN RmMst ON DrPrtKey=RmPrtKey and DrRmCd = RmCd  "
      Else
      ' ***** Manali 3.10.0 - 03/04/12 - Copy Wts From DsgMst
        wRmJoin = " INNER JOIN FgRm ON FdPrtKey=FrPrtKey and FdCoCd = FrCoCd And FdTc = FrTc AND FdYy = FrYy AND FdChr = FrChr AND FdNo = FrNo AND FdSr = FrSr" + _
                 " INNER JOIN RmMst ON FrPrtKey=RmPrtKey and FrRmCd = RmCd  "
        'VK.25 - Join for RngSz added
        wSqlRngSz = " Left Outer Join RmRt On RrTcTyp = 'RSZ' And RrCmCd = '" + ctSelfCmCd + "' And " + _
                "RrCtg = RmCtg And RrSCtg = RmSCtg And RrFrLn <= FrRmSz And RrToLn >= FrRmSz "
      End If
            
      wSqlJoin = " From InvFgd INNER JOIN Fgd ON IfCoCd = FdCoCd And IfFdTc = FdTc AND IfFdYy = FdYy AND IfFdChr = FdChr AND " + _
                " IfFdNo = FdNo AND IfFdSr = FdSr " + wRmJoin
                
      
      'Sql Condition for Bag Records
      wSqlCnd = wSqlCnd + " And IfCoCd = '" + adc("TgCoCd") + "' And IfTc ='" + adc("TgTc") + "' And IfYy = '" + adc("TgYy") + "'" + _
                " And IfChr = '" + adc("TgChr") + "' And IfNo = " + CStr(adc("TgNo")) + " And IfSr = " + CStr(ps_IdSr) + _
                " And IfFdBYy = '" + CStr(ps_BYy) + "' And IfFdBChr = '" + CStr(ps_BChr) + "' And IfFdBNo = " + CStr(ps_BNo)
      
      'Sql Grouping Fields for Bag Records
      wSqlGrp = " IfCoCd, IfTc, IfYy, IfChr, IfNo, IfSr, IfFdBYy, IfFdBChr, IfFdBNo"
      
    Case UCase$("I")
      'Sql Fields for InvDsg Records
      'VK.25 - Changed Order related fields instead of Invoice for jEmr
      If UCase(adc.MenuCd) = UCase("RepJtTagPrn") Then
        wSqlFlds = " OdYy, OdTc, OdChr, OdNo, OdSr,"
      Else
        wSqlFlds = " IdYy, IdTc, IdChr, IdNo, IdSr,"
      End If
      
      'Sql Fields for InvDsg Records
      ' ***** Manali 3.03 - Silver and Palladium included
      ' ***** Manali 3.8.0 - DC Wt & Qty added
      ' ***** Manali 3.9.0 - Oth Met Added
      wSqlBIFlds = GetQWSumSql("D", "Q", "I") + "," + GetQWSumSql("C", "Q", "I") + "," + GetQWSumSql("DC", "Q", "I") + "," + _
                  GetQWSumSql("D", "W", "I") + "," + GetQWSumSql("C", "W", "I") + "," + GetQWSumSql("DC", "W", "I") + "," + _
                  GetQWSumSql("G", "W", "I") + "," + GetQWSumSql("P", "W", "I") + "," + _
                  GetQWSumSql("S", "W", "I") + "," + GetQWSumSql("L", "W", "I") + "," + _
                  GetQWSumSql("X", "Q", "I") + "," + GetQWSumSql("X", "W", "I") + "," + _
                  GetQWSumSql("M", "W", "I") + ","
      
      'Sql Join for Bag Records
      ' ***** Manali 3.10.0 - 03/04/12 - Copy Wts From DsgMst
            
      ' ***** Manali 3.10.0 - 03/04/12 - Copy Wts From DsgMst
      'VK.25 - Changed Order related fields instead of Invoice for jEmr, join for RngSz added
      If UCase(adc.MenuCd) = UCase("RepJtTagPrn") Then
        If adc("TgSrcDsg") = "Y" Then
          wRmJoin = " Join DsgRm On DrPrtKey=OdPrtKey and DrTcTyp=OdDmTcTyp and DrCd=OdDmCd and DrSz=isNull((Select DmSz From DsgMst where DmPrtKey=OdPrtKey and DmTcTyp=OdDmTcTyp and DmCd=OdDmCd and DmSz=OdDmSz),'') " + _
                   " INNER JOIN RmMst ON DrPrtKey=RmPrtKey and DrRmCd = RmCd  "
        Else
          wRmJoin = " Inner Join OrdRm ON OdPrtKey=OrPrtKey and OdCoCd = OrCoCd And OdTc = OrTc And OdYy =OrYy And" + _
                   " OdChr = OrChr And OdNo = OrNo And OdSr = OrSr INNER JOIN RmMst On OrRmCd = RmCd "
          wSqlRngSz = " Left Outer Join RmRt On RrTcTyp = 'RSZ' And RrCmCd = '" + ctSelfCmCd + "' And " + _
                  "RrCtg = RmCtg And RrSCtg = RmSCtg And RrFrLn <= OrLn1 And RrToLn >= OrLn1 "
        End If
      Else
        If adc("TgSrcDsg") = "Y" Then
          wRmJoin = " Join DsgRm On DrPrtKey=IdPrtKey and DrTcTyp='DM' and DrCd=IdDmCd and DrSz=isNull((Select DmSz From DsgMst where DmPrtKey=IdPrtKey and DmTcTyp='DM' and DmCd=IdDmCd and DmSz=IdDmSz),'') " + _
                   " INNER JOIN RmMst ON DrPrtKey=RmPrtKey and DrRmCd = RmCd  "
        Else
          wRmJoin = " Inner Join InvRm ON IdPrtKey=IrPrtKey and IdCoCd = IrCoCd And IdTc = IrTc And IdYy =IrYy And" + _
                 " IdChr = IrChr And IdNo = IrNo And IdSr = IrSr INNER JOIN RmMst On IrRmCd = RmCd "
          'VK.25 - Join for RngSz added
          wSqlRngSz = " Left Outer Join RmRt On RrTcTyp = 'RSZ' And RrCmCd = '" + ctSelfCmCd + "' And " + _
                "RrCtg = RmCtg And RrSCtg = RmSCtg And RrFrLn <= IrRmSz And RrToLn >= IrRmSz "
        End If
      End If
      
      'VK.25 - Changed OrdDsg instead of InvDsg for jEmr
      If UCase(adc.MenuCd) = UCase("RepJtTagPrn") Then
        wSqlJoin = " From OrdDsg " + wRmJoin
      Else
        wSqlJoin = " From InvDsg " + wRmJoin
      End If
                 
      'VK.25 - Changed Order related fields instead of Invoice for jEmr
      If UCase(adc.MenuCd) = UCase("RepJtTagPrn") Then
        wSqlCnd = wSqlCnd + " And OdCoCd = '" + adc("TgCoCd") + "' And OdTc ='" + adc("TgTc") + "' And OdYy = '" + adc("TgYy") + "'" + _
                " And OdChr = '" + adc("TgChr") + "' And OdNo = " + CStr(adc("TgNo")) + " And OdSr = " + CStr(ps_IdSr)
        wSqlGrp = " OdCoCd, OdTc, OdYy, OdChr, OdNo, OdSr"
      Else
        'Sql Condition for Bag Records
        wSqlCnd = wSqlCnd + " And IdCoCd = '" + adc("TgCoCd") + "' And IdTc ='" + adc("TgTc") + "' And IdYy = '" + adc("TgYy") + "'" + _
                " And IdChr = '" + adc("TgChr") + "' And IdNo = " + CStr(adc("TgNo")) + " And IdSr = " + CStr(ps_IdSr)
      
        'Sql Grouping Fields for InvDsg Records
        wSqlGrp = " IdCoCd, IdTc, IdYy, IdChr, IdNo, IdSr"
      End If
  End Select
  
  ' Final SQL Statement
' ########################  Manoj #### Ver: 2.0.7 #### Date: 22/07/2004  ########################
'  wSqlStr = "Select" + wSqlFlds + _
'            " RmCtg, RmSCtg," + IIF(UCase$(ps_Opt) = UCase$("RMSCtg"), "", " RmCd,") + _
'            wSqlBIFlds + wPcFlds + wSqlJoin + wSqlCnd + _
'            " Group By" + wSqlGrp + ", RmCtg, RmSCtg" + IIF(UCase$(ps_Opt) = UCase$("RMSCtg"), "", ", RmCd")
  '28 StnCls added
  wSqlStr = "Select" + wSqlFlds + _
            " RmCtg, " + _
            IIF(UCase(ps_Opt) = UCase$("RmStnCls"), "", "RmSCtg,") + _
            IIF(UCase$(ps_Opt) = UCase$("RmCd"), " RmCd, Max(RmIntQly) as RmIntQly,", "") + _
            IIF(UCase$(ps_Opt) = UCase$("RrRngSz"), " RmCd, Max(RmIntQly) as RmIntQly, IsNull(RrRngSz, '')RrRngSz, ", "") + _
            IIF(UCase$(ps_Opt) = UCase$("RmStnCls"), " RmStnCls, ", "") + _
            wSqlBIFlds + wPcFlds + wSqlJoin + wSqlRngSz + wSqlCnd + _
            " Group By" + wSqlGrp + ", RmCtg, " + _
            IIF(UCase(ps_Opt) = UCase$("RmSCtg"), " RmSCtg ", "") + _
            IIF(UCase$(ps_Opt) = UCase$("RmCd"), " RmSCtg,RmCd", "") + _
            IIF(UCase(ps_Opt) = UCase$("RmStnCls"), "RmStnCls", "") + _
            IIF(UCase$(ps_Opt) = UCase$("RrRngSz"), " RmSCtg,RmCd, IsNull(RrRngSz, '') ", "")
' ###########################################  Manoj  ###########################################

  GetSubSql = wSqlStr
  
End Function
Private Function GetQWSumSql(ByVal ps_RmCtg As String, ByVal ps_WtQtyGrs As String, ByVal ps_BagInPc As String) As String
  ' Getting SQL subString for Dia, CS wise Wt Qty  & Grs Wt
  Dim wSqlStr As String
  Dim wQtyValStr As String, wWtValStr As String, wQtyFldStr As String, wWtFldStr As String
  Dim wValStr As String, wFldStr As String, wStrTmp As String
  Dim wRmCtgStr As String, i As Integer           ' **** Manali 3.8.0 - DC allowed
  
  
  'Generating Field Name
  'Field For
  Select Case UCase$(ps_WtQtyGrs)
    Case UCase$("G")
      wFldStr = "GrsWt"
    Case UCase$("Q")
      wFldStr = "Qty"
    Case UCase$("W")
      wFldStr = "Wt"
  End Select
  
  'Field for Which Category
  wStrTmp = ps_RmCtg
  wFldStr = wStrTmp + wFldStr
  
  'Field For Level
  wStrTmp = ""
  Select Case UCase$(ps_BagInPc)
    Case UCase$("B")
      wStrTmp = "Bag"
    Case UCase$("I")
      wStrTmp = "Id"
    Case UCase$("P")
      wStrTmp = "Pc"
  End Select
  wFldStr = wStrTmp + wFldStr
  
  'Generating Field Definition
  Select Case UCase$(adc("TgAvgTyp"))
    Case UCase$("B")
      'For Bag Weight Or Quantity
      Select Case UCase$(ps_WtQtyGrs)
        Case UCase$("G"), UCase$("W")
          wValStr = "FrRmWt"
          If adc("TgSrcDsg") = "Y" Then wValStr = "DrWt*FdQty"      ' ***** Manali 3.10.0 - 03/04/12 - Copy Wts From DsgMst
        Case UCase$("Q")
          wValStr = "FrRmQty"
         If adc("TgSrcDsg") = "Y" Then wValStr = "DrQty*FdQty"      ' ***** Manali 3.10.0 - 03/04/12 - Copy Wts From DsgMst
      End Select
      'For Pc Level
      If UCase$(ps_BagInPc) = UCase$("P") Then
        wValStr = "(" + wValStr + " / FdQty)"
      End If
    
    Case UCase$("I")
        'For Bag Weight Or Quantity
      Select Case UCase$(ps_WtQtyGrs)
        Case UCase$("G"), UCase$("W")
        
          ' weight Invoice, Actual or Custom
          'VK.25 - Assigned Order Wt instead of Actual, Custom values for jEmr
          If UCase(adc.MenuCd) = UCase("RepJtTagPrn") Then
            wValStr = "OrWt*OdOrdQty"
            If adc("TgSrcDsg") = "Y" Then wValStr = "DrWt*OdOrdQty"
          Else
            Select Case UCase$(adc("TgWt"))
              Case UCase$("A")
                wValStr = "IrRmAWt"
              Case UCase$("I")
                wValStr = "IrRmIWt"
              Case UCase$("Z")
                wValStr = "IrRmZWt"
            End Select
            If adc("TgSrcDsg") = "Y" Then wValStr = "DrWt*IdQty"     ' ***** Manali 3.10.0 - 03/04/12 - Copy Wts From DsgMst
          End If
          
           
        Case UCase$("Q")
          'VK.25 - Assigned OrQty instead of IrRmQty for jEmr
          If UCase(adc.MenuCd) = UCase("RepJtTagPrn") Then
            wValStr = "OrQty*OdOrdQty"
            If adc("TgSrcDsg") = "Y" Then wValStr = "DrQty*OdOrdQty"
          Else
            wValStr = "IrRmQty"
            If adc("TgSrcDsg") = "Y" Then wValStr = "DrQty*IdQty"     ' ***** Manali 3.10.0 - 03/04/12 - Copy Wts From DsgMst
          End If
      End Select
      ' For Pc Level
      If UCase$(ps_BagInPc) = "P" Then
        'VK.25 - Changed OdOrdQty instead of IdQty for jEmr
        If UCase(adc.MenuCd) = UCase("RepJtTagPrn") Then
          wValStr = "(" + wValStr + " / OdOrdQty)"
        Else
          wValStr = "(" + wValStr + " / IdQty)"
        End If
      End If
  End Select
  
   ' **** Manali 3.8.0 - DC allowed
   For i = 1 To Len(ps_RmCtg)
      wRmCtgStr = wRmCtgStr + IIF(wRmCtgStr <> "", ",", "") + "'" + Mid$(ps_RmCtg, i, 1) + "'"
   Next
   ' **** Manali 3.8.0 - DC allowed
  
  'Final SQL
  Select Case UCase$(ps_WtQtyGrs)
  Case UCase$("Q"), UCase$("W")
    ' **** Manali 3.8.0 - DC allowed
    wSqlStr = " Cast(Sum(Case When RmCtg In (" + wRmCtgStr + ") Then " + wValStr + " Else 0 End) As Decimal(10,3)) As " + wFldStr
    ' BEF 3.8.0
    ''wSqlStr = " Cast(Sum(Case When RmCtg = '" + ps_RmCtg + "' Then " + wValStr + " Else 0 End) As Decimal(10,3)) As " + wFldStr
  Case UCase$("G")
    wSqlStr = " Cast(Sum(" + wValStr + " / Case When RmCtg in ('C','D') Then 5 Else 1 End) As Decimal(10,3)) As " + wFldStr
  End Select
  GetQWSumSql = wSqlStr
End Function
Public Function GetValSql(ByVal ps_BagInPc As String) As String
  ' Getting SQL subString for Value
  Dim wSqlStr As String, wValStr As String, wFldStr As String
  ' Zubin
  Dim ws_MUPValStr As String, ws_MUPSOValStr As String
  Dim ws_MUPFldStr As String
  
  ' Generating Field Definition
  ' Field For Bag, Invoice Or Piece
  Select Case UCase$(adc("TgAvgTyp"))
    Case UCase$("B")
      ' Field For Actual Invoice Or Custom Of Bag
      Select Case UCase$(adc("TgPrice"))
        Case UCase$("A")
          wValStr = "IdAVal"
        Case UCase$("I")
          wValStr = "IdIFob"
        Case UCase$("Z")
          wValStr = "IdZFob"
      End Select
      ' Zubin
        ws_MUPValStr = "Cast(max(" + CStr(wValStr) + ") + max(" + CStr(wValStr) + ") * (IsNull((select IsNull(RrSalRt, 0) from RmRt " + _
            "where RrTcTyp= 'MUP' and RrCmCtg= 'C' and RrCmCd= max(IdExpCmCd) and " + _
            "  RrCtg= '' and RrSCtg= '' and RrSTWGrp= '' and RrCd= '' and RrDmCtg= '' and " + _
            "  RrLsCtg= '' and RrLabMCd= '' and RrCmCurCd= '' and " + _
            "  RrFrLn<= max(" + CStr(wValStr) + ") and RrToLn>= max(" + CStr(wValStr) + ")), " + _
            "  IsNull((select IsNull(RrSalRt, 0) from RmRt " + _
            "    where RrTcTyp= 'MUP' and RrCmCtg= 'C' and RrCmCd= '" + ctSelfCmCd + "' and " + _
            "      RrCtg= '' and RrSCtg= '' and RrSTWGrp= '' and RrCd= '' and RrDmCtg= '' and " + _
            "      RrLsCtg= '' and RrLabMCd= '' and RrCmCurCd= '' and " + _
            "      RrFrLn<= max(" + CStr(wValStr) + ") and RrToLn>= max(" + CStr(wValStr) + ")), 0)))/ 100 As Decimal(10,3)) "

        ws_MUPSOValStr = "Cast(max(OdSalPrc) + max(OdSalPrc) * (IsNull((select IsNull(RrSalRt, 0) from RmRt " + _
            "where RrTcTyp= 'MUP' and RrCmCtg= 'C' and RrCmCd= max(IdExpCmCd) and " + _
            "  RrCtg= '' and RrSCtg= '' and RrSTWGrp= '' and RrCd= '' and RrDmCtg= '' and " + _
            "  RrLsCtg= '' and RrLabMCd= '' and RrCmCurCd= '' and " + _
            "  RrFrLn<= max(OdSalPrc) and RrToLn>= max(OdSalPrc)), " + _
            "  IsNull((select IsNull(RrSalRt, 0) from RmRt " + _
            "    where RrTcTyp= 'MUP' and RrCmCtg= 'C' and RrCmCd= '" + ctSelfCmCd + "' and " + _
            "      RrCtg= '' and RrSCtg= '' and RrSTWGrp= '' and RrCd= '' and RrDmCtg= '' and " + _
            "      RrLsCtg= '' and RrLabMCd= '' and RrCmCurCd= '' and " + _
            "      RrFrLn<= max(OdSalPrc) and RrToLn>= max(OdSalPrc)), 0)))/ 100 As Decimal(10,3)) "
      
      ws_MUPValStr = "(isNull((" + ws_MUPValStr + " / (Case When Max(IdQty) <> 0 Then Max(IdQty) Else 1 End)), " + ws_MUPValStr + ")"
      
      'wValStr = "(isNull((Max(" + wValStr + ") / (Case When Max(IdQty) <> 0 Then Max(IdQty) Else 1 End)), Max(OdSalPrc))"
      wValStr = "(isNull((Max(" + wValStr + ") / (Case When Max(IdQty) <> 0 Then Max(IdQty) Else 1 End)), " + ws_MUPSOValStr + ")"
      
      If UCase$(ps_BagInPc) <> UCase$("P") Then
        wValStr = "" + wValStr + " * Max(FdQty)"
        ws_MUPValStr = "" + ws_MUPValStr + " * Max(FdQty)"
      End If
      ' Zubin
      ws_MUPValStr = "" + ws_MUPValStr + ")"
    Case UCase$("I")
      ' Field For Actual Invoice Or Custom Of InvDsg
      Select Case UCase$(adc("TgPrice"))
        Case UCase$("A")
          wValStr = "IdAVal"
        Case UCase$("I")
          'VK.25 - Changed OdSalPrc instead of IdIFob for jEmr
          If UCase(adc.MenuCd) = UCase("RepJtTagPrn") Then
            wValStr = "OdSalPrc"
          Else
          'SRITMP VK.25 Trying
            wValStr = "(IdIFob*OdOrdQty)"
          End If
        Case UCase$("Z")
          wValStr = "IdZFob"
      End Select
      ' Zubin
      If UCase(adc.MenuCd) = UCase("RepJtTagPrn") Then
        ws_MUPValStr = "Cast(max(" + CStr(wValStr) + ") + max(" + CStr(wValStr) + ") * (IsNull((select IsNull(RrSalRt, 0) from RmRt " + _
            "where RrTcTyp= 'MUP' and RrCmCtg= 'C' and RrCmCd= max(OmCmCd) and " + _
            "  RrCtg= '' and RrSCtg= '' and RrSTWGrp= '' and RrCd= '' and RrDmCtg= '' and " + _
            "  RrLsCtg= '' and RrLabMCd= '' and RrCmCurCd= '' and " + _
            "  RrFrLn<= max(" + CStr(wValStr) + ") and RrToLn>= max(" + CStr(wValStr) + ")), " + _
            "  IsNull((select IsNull(RrSalRt, 0) from RmRt " + _
            "    where RrTcTyp= 'MUP' and RrCmCtg= 'C' and RrCmCd= '" + ctSelfCmCd + "' and " + _
            "      RrCtg= '' and RrSCtg= '' and RrSTWGrp= '' and RrCd= '' and RrDmCtg= '' and " + _
            "      RrLsCtg= '' and RrLabMCd= '' and RrCmCurCd= '' and " + _
            "      RrFrLn<= max(" + CStr(wValStr) + ") and RrToLn>= max(" + CStr(wValStr) + ")), 0)))/ 100 As Decimal(10,3)) "
      
        ws_MUPSOValStr = "Cast(max(OdSalPrc) + max(OdSalPrc) * (IsNull((select IsNull(RrSalRt, 0) from RmRt " + _
            "where RrTcTyp= 'MUP' and RrCmCtg= 'C' and RrCmCd= max(OmCmCd) and " + _
            "  RrCtg= '' and RrSCtg= '' and RrSTWGrp= '' and RrCd= '' and RrDmCtg= '' and " + _
            "  RrLsCtg= '' and RrLabMCd= '' and RrCmCurCd= '' and " + _
            "  RrFrLn<= max(OdSalPrc) and RrToLn>= max(OdSalPrc)), " + _
            "  IsNull((select IsNull(RrSalRt, 0) from RmRt " + _
            "    where RrTcTyp= 'MUP' and RrCmCtg= 'C' and RrCmCd= '" + ctSelfCmCd + "' and " + _
            "      RrCtg= '' and RrSCtg= '' and RrSTWGrp= '' and RrCd= '' and RrDmCtg= '' and " + _
            "      RrLsCtg= '' and RrLabMCd= '' and RrCmCurCd= '' and " + _
            "      RrFrLn<= max(OdSalPrc) and RrToLn>= max(OdSalPrc)), 0)))/ 100 As Decimal(10,3)) "
      Else
        ws_MUPValStr = "Cast(max(" + CStr(wValStr) + ") + max(" + CStr(wValStr) + ") * (IsNull((select IsNull(RrSalRt, 0) from RmRt " + _
            "where RrTcTyp= 'MUP' and RrCmCtg= 'C' and RrCmCd= max(IdExpCmCd) and " + _
            "  RrCtg= '' and RrSCtg= '' and RrSTWGrp= '' and RrCd= '' and RrDmCtg= '' and " + _
            "  RrLsCtg= '' and RrLabMCd= '' and RrCmCurCd= '' and " + _
            "  RrFrLn<= max(" + CStr(wValStr) + ") and RrToLn>= max(" + CStr(wValStr) + ")), " + _
            "  IsNull((select IsNull(RrSalRt, 0) from RmRt " + _
            "    where RrTcTyp= 'MUP' and RrCmCtg= 'C' and RrCmCd= '" + ctSelfCmCd + "' and " + _
            "      RrCtg= '' and RrSCtg= '' and RrSTWGrp= '' and RrCd= '' and RrDmCtg= '' and " + _
            "      RrLsCtg= '' and RrLabMCd= '' and RrCmCurCd= '' and " + _
            "      RrFrLn<= max(" + CStr(wValStr) + ") and RrToLn>= max(" + CStr(wValStr) + ")), 0)))/ 100 As Decimal(10,3)) "
      
        ws_MUPSOValStr = "Cast(max(OdSalPrc) + max(OdSalPrc) * (IsNull((select IsNull(RrSalRt, 0) from RmRt " + _
            "where RrTcTyp= 'MUP' and RrCmCtg= 'C' and RrCmCd= max(IdExpCmCd) and " + _
            "  RrCtg= '' and RrSCtg= '' and RrSTWGrp= '' and RrCd= '' and RrDmCtg= '' and " + _
            "  RrLsCtg= '' and RrLabMCd= '' and RrCmCurCd= '' and " + _
            "  RrFrLn<= max(OdSalPrc) and RrToLn>= max(OdSalPrc)), " + _
            "  IsNull((select IsNull(RrSalRt, 0) from RmRt " + _
            "    where RrTcTyp= 'MUP' and RrCmCtg= 'C' and RrCmCd= '" + ctSelfCmCd + "' and " + _
            "      RrCtg= '' and RrSCtg= '' and RrSTWGrp= '' and RrCd= '' and RrDmCtg= '' and " + _
            "      RrLsCtg= '' and RrLabMCd= '' and RrCmCurCd= '' and " + _
            "      RrFrLn<= max(OdSalPrc) and RrToLn>= max(OdSalPrc)), 0)))/ 100 As Decimal(10,3)) "
      End If
      
      If UCase$(ps_BagInPc) = "P" Then
        'VK.25 - Changed OdOrdQty instead of IdQty for jEmr
        If UCase(adc.MenuCd) <> UCase("RepJtTagPrn") Then
          wValStr = wValStr + ") / Max(IdQty"
          ' Zubin
          ws_MUPValStr = ws_MUPValStr + " / Max(IdQty)"
        Else
          wValStr = wValStr + ") / Max(OdOrdQty"
          ' Zubin
          ws_MUPValStr = ws_MUPValStr + " / Max(OdOrdQty)"
        End If
      End If
      ' Zubin
      'ws_MUPValStr = "" + ws_MUPValStr
      wValStr = "Max(" + wValStr
      
  End Select
  
  ' Generating Field Name for Bag, Invoice Or Piece
  Select Case UCase$(ps_BagInPc)
    Case UCase$("B")
      wFldStr = "BagVal"
      ws_MUPFldStr = "BagMuVal"
    Case UCase$("I")
      wFldStr = "IdVal"
      ws_MUPFldStr = "IdMuVal"
    Case UCase$("P")
      wFldStr = "PcVal"
      ws_MUPFldStr = "PcMuVal"
  End Select
  
  'Final SQL Statement
  'wSqlStr = " Cast(" + wValStr + ") As Decimal(10,3)) As " + wFldStr
  ' Zubin
  wSqlStr = " Cast(" + wValStr + ")*" + CStr(adc("TgMulBy")) + " As Decimal(10,3)) As " + wFldStr + _
            ", Cast(" + ws_MUPValStr + "*" + CStr(adc("TgMulBy")) + " As Decimal(10,3)) As " + ws_MUPFldStr

    ' ***** Manali 3.10.0 - 03/04/12 - Copy Wts From DsgMst
    If adc("TgSrcDsg") = "Y" Then wSqlStr = " Cast(0 As Decimal(10,3)) As " + wFldStr + ",Cast(0 As Decimal(10,3)) As " + ws_MUPFldStr
  GetValSql = wSqlStr
End Function

Private Sub PrnRep()
  Dim ws_FlNm As String
  Dim ws_str As String, wi_Cnt As Integer
  
  Dim hPrinter&
  Dim Res&
  Dim JobId&
  Dim Written&
  Dim PrintData$
  Dim DocInfo As DOC_INFO_1
  
  ' If prntOpt"F" then Open File to save Tag
  If UCase$(mPrntOpt) = UCase$("F") Then
    ws_FlNm = OpenFilePath
    'Open File For Writting
    Open ws_FlNm For Output Access Write As #1
  ElseIf UCase$(mPrntOpt) = UCase$("P") Then
    ' Initialize Printer for Printing
    Res& = OpenPrinter(SelPrinter, hPrinter, 0)
    If Res = 0 Then DispMsg "Unable to open the printer", etError: Exit Sub
    DocInfo.pDocName = "MyDoc"
    DocInfo.pOutputFile = vbNullString
    DocInfo.pDatatype = vbNullString
    JobId = StartDocPrinter(hPrinter, 1, DocInfo)
    Call StartPagePrinter(hPrinter)
    Call WritePrinter(hPrinter, ByVal PrintData$, Len(PrintData$), Written)
  End If
  MWLib.BeginProcess Me, "Printing Records To" + IIF(mPrntOpt = "F", " File " + ws_FlNm, " Printer " + cboPrn.text) & " ..."
  wi_Cnt = 0
  With mRs_Rep
    '*** For each record in the resultset, processing starts
    Do While Not (.EOF Or .BOF)
      '*** For bar code printing, format each label as per the printer specification
      mRs_FmtStr.MoveFirst
      ws_str = Trim$(mRs_FmtStr.FldValue("TgfStartStr")) + vbCrLf
      ws_str = ws_str + IIF(Trim$(mRs_FmtStr.FldValue("TgfDarkness")) <> "", Trim$(mRs_FmtStr.FldValue("TgfDarkness")) + vbCrLf, "")
      ws_str = ws_str + IIF(Trim$(mRs_FmtStr.FldValue("TgfHeat")) <> "", Trim$(mRs_FmtStr.FldValue("TgfHeat")) + vbCrLf, "")
      Do While Not (mRs_FmtStr.EOF Or mRs_FmtStr.BOF)
        ws_str = ws_str + Trim$(mRs_FmtStr.FldValue("TgfFldOrg")) + Trim$(mRs_FmtStr.FldValue("TgfsPrintStr")) + _
                  Trim$(mRs_FmtStr.FldValue("TgfFldData")) + Trim$(.FldValue("TgdFld" + Format$(mRs_FmtStr.FldValue("TgfsSr"), "00"))) + _
                  Trim$(mRs_FmtStr.FldValue("TgfFldEnd")) + vbCrLf
        mRs_FmtStr.MoveNext
      Loop
      mRs_FmtStr.MoveLast
      ws_str = ws_str + Trim$(mRs_FmtStr.FldValue("TgfQtyStr")) + Format(CDbl(.FldValue("TgdQty")), "0000") + vbCrLf
      ws_str = ws_str + Trim$(mRs_FmtStr.FldValue("TgfEndStr"))
          
      If UCase$(mPrntOpt) = UCase$("P") Then
        ' Print data on Printer
        PrintData = ws_str
        Call WritePrinter(hPrinter, ByVal PrintData$, Len(PrintData$), Written)
      ElseIf UCase$(mPrntOpt) = UCase$("F") Then
        ' Print data into File
        Print #1, ws_str
      End If
      wi_Cnt = wi_Cnt + 1
      MWLib.RunningProcess "Records " + CStr(wi_Cnt) + "/" + CStr(.RecCount) + " Are Printed" + _
                          " To" + IIF(mPrntOpt = "F", " File " + ws_FlNm, " Printer " + cboPrn.text) + " ..."
      .MoveNext
    Loop
    MWLib.EndProcess Me
    If UCase$(mPrntOpt) = UCase$("F") Then
      'Close File
      Close #1
    ElseIf UCase$(mPrntOpt) = UCase$("p") Then
      ' Close Printer
      Call EndPagePrinter(hPrinter)
      Call EndDocPrinter(hPrinter)
      Call ClosePrinter(hPrinter) ' Close when done
    End If
  End With
End Sub

Private Function GetRepRec() As Boolean
  ' Set Record sets for Report Printing
  Dim wSqlStr As String, wSqlCnd As String
  GetRepRec = False
  
  wSqlCnd = ""
  
  ' Building Condition for SQL Statement of Printing Record
  If adc("wIdSrFrPrn") > 0 Then wSqlCnd = wSqlCnd + " And TgdIdSr >= " + CStr(adc("wIdSrFrPrn"))
  If adc("wIdSrToPrn") > 0 Then wSqlCnd = wSqlCnd + " And TgdIdSr <= " + CStr(adc("wIdSrToPrn"))
  If adc("wTgSrNoFr") > 0 Then wSqlCnd = wSqlCnd + " And TgdSrNo >= " + CStr(adc("wTgSrNoFr"))
  If adc("wTgSrNoTo") > 0 Then wSqlCnd = wSqlCnd + " And TgdSrNo <= " + CStr(adc("wTgSrNoTo"))
  

  ' SQL Statement Generating Records for Printing
  wSqlStr = "Select TgdQty, TgdFld01, TgdFld02, TgdFld03, TgdFld04, TgdFld05, TgdFld06, TgdFld07, TgdFld08, TgdFld09, TgdFld10, TgdFld11, TgdFld12 " + _
            " From Tag Join TagData On TgCoCd = TgdCoCd And TgTc = TgdTc And TgYy = TgdYy And TgChr = TgdChr And TgNo = TgdNo And TgSr = TgdSr" + _
            " Where TgCoCd = '" + ms_TgCoCd + "' And TgTc = '" + ms_TgTc + "' And TgYy = '" + ms_TgYy + "'" + _
            " And TgChr = '" + ms_TgChr + "' And TgNo = " + CStr(ms_TgNo) + " And TgSr = " + CStr(ms_TgSr) + _
            wSqlCnd
  Set mRs_Rep = moCn.OpenResultset(wSqlStr)
  If mRs_Rep.RecCount = 0 Then
    DispMsg "No Data Found For The Scope", etError: Exit Function
  End If
  
  ' SQL Statement Generating Records for Format Applying to Printing Records
  wSqlStr = "Select TgfStartStr, TgfDarkness, TgfHeat, TgfQtyStr, TgfEndStr, TgfFldOrg, TgfFldData, TgfFldEnd, TgfsSr, TgfsPrintStr, TgfsField" + _
            " From TagFmt Join TagFmtSr On TgfCd =TgfsCd Where TgfCoCd = '" + ms_TgCoCd + "' And TgfCd = '" + adc("TgCd") + "'"
  Set mRs_FmtStr = moCn.OpenResultset(wSqlStr)
  If mRs_FmtStr.RecCount = 0 Then
    DispMsg "No Tag Format Found For Selected Records", etError: Exit Function
  End If
  GetRepRec = True
End Function


