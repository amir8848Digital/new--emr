VERSION 5.00
Object = "{AB2CC32A-E652-42A2-B341-26ECE02F6386}#2.1#0"; "MwfCtlForSpeed8.ocx"
Begin VB.Form EmrFrmCustmPur 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Custom Purchase"
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
   Begin MwfCtl.MWCTL_BTN1 CmdInvHd 
      Height          =   465
      Left            =   6150
      TabIndex        =   137
      ToolTipText     =   "Enter invoice header details"
      Top             =   9300
      Width           =   885
      _ExtentX        =   1561
      _ExtentY        =   820
      BackColor       =   -2147483633
      MaskColor       =   8454143
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
   Begin MwfCtl.MWCTL_BTN1 CmdChgDt 
      Height          =   465
      Left            =   7950
      TabIndex        =   135
      ToolTipText     =   "Change Voucher Date"
      Top             =   9300
      Width           =   885
      _ExtentX        =   1561
      _ExtentY        =   820
      BackColor       =   -2147483633
      MaskColor       =   8454143
      Caption         =   "Chan&ge Date"
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
   Begin MwfCtl.MWCTL_BTN1 CmdCp 
      Height          =   465
      Left            =   7050
      TabIndex        =   33
      ToolTipText     =   "Option To Copy From Raw Material Purchase Voucher"
      Top             =   9300
      Width           =   885
      _ExtentX        =   1561
      _ExtentY        =   820
      BackColor       =   -2147483633
      MaskColor       =   8454143
      Caption         =   "&Copy "
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
      Left            =   8970
      TabIndex        =   43
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
      Left            =   0
      TabIndex        =   44
      Top             =   9300
      Width           =   14865
      _ExtentX        =   26220
      _ExtentY        =   979
      BtnWidth        =   885
      BtnHeight       =   435
      BtnMaskColor    =   8454143
      FontSz          =   8.25
   End
   Begin VB.Frame FRA_KEYS 
      BorderStyle     =   0  'None
      Height          =   280
      Left            =   0
      TabIndex        =   50
      Top             =   0
      Width           =   4095
      Begin MwfCtl.MWCTL_MED ATXT 
         Height          =   285
         Index           =   1
         Left            =   945
         TabIndex        =   48
         ToolTipText     =   "Enter Voucher Tc"
         Top             =   0
         Width           =   570
         _ExtentX        =   1005
         _ExtentY        =   503
         Enabled         =   0   'False
         MaxLength       =   3
         DataType        =   4
         DataField       =   "TTc"
         ReCalcParent    =   "TCHR"
         IdName          =   "TTC"
      End
      Begin MwfCtl.MWCTL_MED ATXT 
         Height          =   285
         Index           =   4
         Left            =   2835
         TabIndex        =   2
         ToolTipText     =   "Enter Voucher Number"
         Top             =   0
         Width           =   930
         _ExtentX        =   1640
         _ExtentY        =   503
         Alignment       =   2
         Mask            =   "######0"
         MaxLength       =   7
         DataType        =   1
         DataField       =   "TNo"
         IdName          =   "TNO"
      End
      Begin MwfCtl.MWCTL_MED ATXT 
         Height          =   285
         Index           =   6
         Left            =   1620
         TabIndex        =   0
         ToolTipText     =   "Enter Voucher Year"
         Top             =   0
         Width           =   435
         _ExtentX        =   767
         _ExtentY        =   503
         MaxLength       =   2
         DataType        =   4
         DataField       =   "TYy"
         IdName          =   "TYY"
      End
      Begin MwfCtl.MWCTL_MED ATXT 
         Height          =   285
         Index           =   41
         Left            =   2160
         TabIndex        =   1
         ToolTipText     =   "Enter Voucher Character"
         Top             =   0
         Width           =   570
         _ExtentX        =   1005
         _ExtentY        =   503
         MaxLength       =   3
         DataType        =   4
         DataField       =   "TChr"
         ReCalcOn        =   "TTC"
         IdName          =   "TCHR"
      End
      Begin MwfCtl.MWCTL_MED ATXT 
         Height          =   345
         Index           =   28
         Left            =   3780
         TabIndex        =   88
         Top             =   0
         Visible         =   0   'False
         Width           =   135
         _ExtentX        =   238
         _ExtentY        =   609
         Enabled         =   0   'False
         MaxLength       =   3
         DataType        =   4
         DataField       =   "TCoCd"
         IdName          =   "TCOCD"
      End
      Begin MwfCtl.MWCTL_MED ATXT 
         Height          =   345
         Index           =   38
         Left            =   3810
         TabIndex        =   102
         Top             =   -30
         Visible         =   0   'False
         Width           =   255
         _ExtentX        =   450
         _ExtentY        =   609
         Enabled         =   0   'False
         MaxLength       =   3
         DataType        =   4
         DataField       =   "TPrtKey"
         IdName          =   "TPRTKEY"
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
         Left            =   30
         TabIndex        =   57
         Top             =   0
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
         Left            =   2070
         TabIndex        =   55
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
         TabIndex        =   54
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
         TabIndex        =   52
         Top             =   0
         Width           =   135
      End
   End
   Begin VB.Frame FRA_NKEYS 
      BorderStyle     =   0  'None
      Enabled         =   0   'False
      Height          =   9930
      Left            =   -120
      TabIndex        =   49
      Top             =   -120
      Width           =   15345
      Begin VB.Frame FraCp 
         Height          =   1185
         Left            =   4800
         TabIndex        =   34
         Top             =   8160
         Width           =   5655
         Begin MwfCtl.MWCTL_BTN1 CmdCpGo 
            Height          =   495
            Left            =   5025
            TabIndex        =   42
            ToolTipText     =   "Click Here To Start Copy"
            Top             =   510
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
            Index           =   57
            Left            =   2235
            TabIndex        =   36
            ToolTipText     =   "Enter Purchase Voucher Tc"
            Top             =   420
            Width           =   570
            _ExtentX        =   1005
            _ExtentY        =   503
            MaxLength       =   3
            DataType        =   4
            ReCalcParent    =   "WCPTCHR"
            IdName          =   "WCPTTC"
         End
         Begin MwfCtl.MWCTL_MED ATXT 
            Height          =   285
            Index           =   58
            Left            =   2910
            TabIndex        =   37
            ToolTipText     =   "Enter Purchase Voucher Year"
            Top             =   420
            Width           =   330
            _ExtentX        =   582
            _ExtentY        =   503
            MaxLength       =   2
            DataType        =   4
            IdName          =   "WCPTYY"
         End
         Begin MwfCtl.MWCTL_MED ATXT 
            Height          =   285
            Index           =   59
            Left            =   3360
            TabIndex        =   38
            ToolTipText     =   "Enter Purchase Voucher Character"
            Top             =   420
            Width           =   570
            _ExtentX        =   1005
            _ExtentY        =   503
            MaxLength       =   3
            DataType        =   4
            ReCalcOn        =   "WCPTTC"
            IdName          =   "WCPTCHR"
         End
         Begin MwfCtl.MWCTL_MED ATXT 
            Height          =   285
            Index           =   60
            Left            =   4050
            TabIndex        =   39
            ToolTipText     =   "Enter Purchase Voucher No."
            Top             =   420
            Width           =   750
            _ExtentX        =   1323
            _ExtentY        =   503
            Alignment       =   2
            Mask            =   "######0"
            MaxLength       =   7
            DataType        =   1
            IdName          =   "WCPTNO"
         End
         Begin MwfCtl.MWCTL_MED ATXT 
            Height          =   285
            Index           =   27
            Left            =   1500
            TabIndex        =   40
            ToolTipText     =   "Enter Option By Which The Source Records Will Be Grouped For Copying"
            Top             =   705
            Width           =   1275
            _ExtentX        =   2249
            _ExtentY        =   503
            MaxLength       =   8
            DataType        =   4
            IdName          =   "WCPGRP"
         End
         Begin MwfCtl.MWCTL_MED ATXT 
            Height          =   285
            Index           =   35
            Left            =   1500
            TabIndex        =   35
            ToolTipText     =   "Enter Option By Which The Source Records Will Be Grouped For Copying"
            Top             =   420
            Width           =   615
            _ExtentX        =   1085
            _ExtentY        =   503
            MaxLength       =   3
            DataType        =   4
            IdName          =   "WCPTCOCD"
         End
         Begin MwfCtl.MWCTL_MED ATXT 
            Height          =   285
            Index           =   51
            Left            =   4455
            TabIndex        =   41
            ToolTipText     =   "Copy Best Use Size Enter Y or N"
            Top             =   720
            Width           =   345
            _ExtentX        =   609
            _ExtentY        =   503
            DataType        =   4
            IdName          =   "WCPBSSZYN"
         End
         Begin VB.Label LblCpBsSz 
            BackStyle       =   0  'Transparent
            Caption         =   "Copy Size"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Left            =   3420
            TabIndex        =   126
            Top             =   735
            Width           =   1020
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
            Index           =   13
            Left            =   2145
            TabIndex        =   98
            Top             =   420
            Width           =   105
         End
         Begin VB.Label LblMain 
            BackStyle       =   0  'Transparent
            BorderStyle     =   1  'Fixed Single
            Caption         =   "  Copy"
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
            TabIndex        =   95
            Top             =   90
            Width           =   5610
         End
         Begin VB.Label ALBL 
            BackStyle       =   0  'Transparent
            Caption         =   "Group Option"
            BeginProperty Font 
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
            TabIndex        =   87
            Top             =   705
            Width           =   1470
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
            Index           =   59
            Left            =   120
            TabIndex        =   86
            Top             =   420
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
            Index           =   58
            Left            =   2820
            TabIndex        =   85
            Top             =   420
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
            Index           =   57
            Left            =   3255
            TabIndex        =   84
            Top             =   420
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
            Index           =   54
            Left            =   3945
            TabIndex        =   83
            Top             =   420
            Width           =   105
         End
      End
      Begin VB.Frame FraGrdTxndZDets 
         Height          =   915
         Left            =   120
         TabIndex        =   103
         Top             =   6930
         Width           =   15195
         Begin VB.Frame FraGrdTxndZDetsDisp 
            Height          =   915
            Left            =   1410
            TabIndex        =   115
            Top             =   0
            Width           =   13785
            Begin MwfCtl.MWCTL_MED ATXT 
               Height          =   285
               Index           =   34
               Left            =   810
               TabIndex        =   105
               ToolTipText     =   "Enter Description"
               Top             =   180
               Width           =   8790
               _ExtentX        =   15505
               _ExtentY        =   503
               MaxLength       =   60
               ReCalcParent    =   "GrdTxndZ(TZDESC)"
               IdName          =   "WTZDESC"
            End
            Begin MwfCtl.MWCTL_MED ATXT 
               Height          =   285
               Index           =   40
               Left            =   1470
               TabIndex        =   107
               ToolTipText     =   "Enter Order Year"
               Top             =   465
               Width           =   330
               _ExtentX        =   582
               _ExtentY        =   503
               MaxLength       =   2
               DataType        =   4
               IdName          =   "WTZODYY"
            End
            Begin MwfCtl.MWCTL_MED ATXT 
               Height          =   285
               Index           =   42
               Left            =   2550
               TabIndex        =   109
               ToolTipText     =   "Enter Order Number"
               Top             =   465
               Width           =   750
               _ExtentX        =   1323
               _ExtentY        =   503
               Alignment       =   2
               Mask            =   "#####0"
               MaxLength       =   6
               DataType        =   1
               IdName          =   "WTZODNO"
            End
            Begin MwfCtl.MWCTL_MED ATXT 
               Height          =   285
               Index           =   44
               Left            =   810
               TabIndex        =   106
               ToolTipText     =   "Enter Order Tc"
               Top             =   465
               Width           =   570
               _ExtentX        =   1005
               _ExtentY        =   503
               MaxLength       =   3
               DataType        =   4
               IdName          =   "WTZODTC"
            End
            Begin MwfCtl.MWCTL_MED ATXT 
               Height          =   285
               Index           =   45
               Left            =   1890
               TabIndex        =   108
               ToolTipText     =   "Enter Order Character"
               Top             =   465
               Width           =   570
               _ExtentX        =   1005
               _ExtentY        =   503
               MaxLength       =   3
               DataType        =   4
               IdName          =   "WTZODCHR"
            End
            Begin MwfCtl.MWCTL_MED ATXT 
               Height          =   285
               Index           =   46
               Left            =   4935
               TabIndex        =   111
               ToolTipText     =   "Enter Design Code"
               Top             =   465
               Width           =   2385
               _ExtentX        =   4207
               _ExtentY        =   503
               MaxLength       =   15
               DataType        =   4
               IdName          =   "WTZDMCD"
            End
            Begin MwfCtl.MWCTL_MED ATXT 
               Height          =   285
               Index           =   47
               Left            =   8100
               TabIndex        =   112
               ToolTipText     =   "Enter RM Code"
               Top             =   465
               Width           =   2445
               _ExtentX        =   4313
               _ExtentY        =   503
               MaxLength       =   16
               DataType        =   4
               IdName          =   "WTZRMCD"
            End
            Begin MwfCtl.MWCTL_MED ATXT 
               Height          =   285
               Index           =   49
               Left            =   11190
               TabIndex        =   113
               ToolTipText     =   "Enter From Size"
               Top             =   465
               Width           =   900
               _ExtentX        =   1588
               _ExtentY        =   503
               Alignment       =   2
               Mask            =   "##0.0000"
               MaxLength       =   8
               DataType        =   2
               IdName          =   "WTZRMSZFR"
            End
            Begin MwfCtl.MWCTL_MED ATXT 
               Height          =   285
               Index           =   50
               Left            =   12780
               TabIndex        =   114
               ToolTipText     =   "Enter To Size"
               Top             =   465
               Width           =   900
               _ExtentX        =   1588
               _ExtentY        =   503
               Alignment       =   2
               Mask            =   "##0.0000"
               MaxLength       =   8
               DataType        =   2
               IdName          =   "WTZRMSZTO"
            End
            Begin MwfCtl.MWCTL_MED ATXT 
               Height          =   285
               Index           =   54
               Left            =   3405
               TabIndex        =   110
               ToolTipText     =   "Enter Order Serial Number"
               Top             =   465
               Width           =   690
               _ExtentX        =   1217
               _ExtentY        =   503
               Alignment       =   2
               Mask            =   "####0"
               MaxLength       =   5
               DataType        =   1
               IdName          =   "WTZODSR"
            End
            Begin VB.Label LblOrdSlash4 
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
               Left            =   3330
               TabIndex        =   136
               Top             =   465
               Width           =   105
            End
            Begin VB.Label Label1 
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
               Left            =   120
               TabIndex        =   124
               Top             =   180
               Width           =   615
            End
            Begin VB.Label LblOrdSlash1 
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
               Left            =   1410
               TabIndex        =   123
               Top             =   465
               Width           =   105
            End
            Begin VB.Label LblOrdSlash2 
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
               Left            =   1815
               TabIndex        =   122
               Top             =   465
               Width           =   105
            End
            Begin VB.Label LblOrdSlash3 
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
               Left            =   2475
               TabIndex        =   121
               Top             =   465
               Width           =   105
            End
            Begin VB.Label LblwTzOrd 
               BackStyle       =   0  'Transparent
               Caption         =   "Order"
               BeginProperty Font 
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
               TabIndex        =   120
               Top             =   465
               Width           =   615
            End
            Begin VB.Label LblwTzDmCd 
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
               Left            =   4230
               TabIndex        =   119
               Top             =   465
               Width           =   750
            End
            Begin VB.Label LblwTzRmCd 
               BackStyle       =   0  'Transparent
               Caption         =   "Rm Cd"
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   285
               Left            =   7440
               TabIndex        =   118
               Top             =   465
               Width           =   750
            End
            Begin VB.Label LblwTzRmSzFr 
               BackStyle       =   0  'Transparent
               Caption         =   "Fr Sz"
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   285
               Left            =   10650
               TabIndex        =   117
               Top             =   465
               Width           =   570
            End
            Begin VB.Label LblwTzRmSzTo 
               BackStyle       =   0  'Transparent
               Caption         =   "To Sz"
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   285
               Left            =   12180
               TabIndex        =   116
               Top             =   465
               Width           =   600
            End
         End
         Begin MwfCtl.MWCTL_BTN1 CmdGrdTxndzDets 
            Height          =   465
            Left            =   135
            TabIndex        =   104
            Top             =   180
            Width           =   1065
            _ExtentX        =   1879
            _ExtentY        =   820
            BackColor       =   -2147483633
            MaskColor       =   8454143
            Caption         =   "&Enter Dets"
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
      End
      Begin VB.Frame FraNKeyAll 
         BorderStyle     =   0  'None
         Height          =   9345
         Left            =   120
         TabIndex        =   51
         Top             =   120
         Width           =   15285
         Begin VB.TextBox TxtPwd 
            Height          =   285
            IMEMode         =   3  'DISABLE
            Left            =   8625
            MaxLength       =   30
            PasswordChar    =   "*"
            TabIndex        =   5
            ToolTipText     =   "Enter SEO Password"
            Top             =   0
            Width           =   2340
         End
         Begin VB.Frame FraCnvTot 
            Height          =   825
            Left            =   0
            TabIndex        =   65
            Top             =   8400
            Width           =   15195
            Begin MwfCtl.MWCTL_MED ATXT 
               Height          =   285
               Index           =   31
               Left            =   1830
               TabIndex        =   30
               Top             =   420
               Width           =   1215
               _ExtentX        =   2143
               _ExtentY        =   503
               Alignment       =   2
               Enabled         =   0   'False
               Mask            =   "######0"
               MaxLength       =   7
               DataType        =   2
               ReCalcOn        =   "GrdTxndZ(TZRMQTY)"
               IdName          =   "WTOTQTY"
            End
            Begin MwfCtl.MWCTL_MED ATXT 
               Height          =   285
               Index           =   32
               Left            =   5160
               TabIndex        =   31
               Top             =   420
               Width           =   1215
               _ExtentX        =   2143
               _ExtentY        =   503
               Alignment       =   2
               Enabled         =   0   'False
               Mask            =   "#####0.000"
               MaxLength       =   10
               DataType        =   2
               ReCalcOn        =   "GrdTxndZ(TZRMWT)"
               IdName          =   "WTOTWT"
            End
            Begin MwfCtl.MWCTL_MED ATXT 
               Height          =   285
               Index           =   3
               Left            =   8310
               TabIndex        =   32
               Top             =   420
               Width           =   1215
               _ExtentX        =   2143
               _ExtentY        =   503
               Alignment       =   2
               Enabled         =   0   'False
               Mask            =   "#####0.000"
               MaxLength       =   10
               DataType        =   2
               ReCalcOn        =   "GrdTxndZ(WTZRMBALWT)"
               IdName          =   "WBALWT"
            End
            Begin VB.Label LblMain 
               BackStyle       =   0  'Transparent
               BorderStyle     =   1  'Fixed Single
               Caption         =   "  Totals"
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
               TabIndex        =   96
               Top             =   90
               Width           =   15195
            End
            Begin VB.Label LblBalWt 
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
               Left            =   7560
               TabIndex        =   80
               Top             =   360
               Width           =   1005
            End
            Begin VB.Label LblTotQty 
               BackStyle       =   0  'Transparent
               Caption         =   "Tot Qty"
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   285
               Left            =   1020
               TabIndex        =   67
               Top             =   420
               Width           =   855
            End
            Begin VB.Label LblTotWt 
               BackStyle       =   0  'Transparent
               Caption         =   "Tot Wt"
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   285
               Left            =   4410
               TabIndex        =   66
               Top             =   420
               Width           =   825
            End
         End
         Begin MwfCtl.MWCTL_MED ATXT 
            Height          =   285
            Index           =   22
            Left            =   13215
            TabIndex        =   46
            ToolTipText     =   "Enter Raw Material Default To Location"
            Top             =   270
            Visible         =   0   'False
            Width           =   210
            _ExtentX        =   370
            _ExtentY        =   503
            Enabled         =   0   'False
            MaxLength       =   5
            DataType        =   4
            DataField       =   "TToRmLoc"
            IdName          =   "TTORMLOC"
         End
         Begin MwfCtl.MWCTL_MED ATXT 
            Height          =   285
            Index           =   30
            Left            =   4635
            TabIndex        =   3
            ToolTipText     =   "Enter Voucher Date"
            Top             =   0
            Width           =   855
            _ExtentX        =   1508
            _ExtentY        =   503
            MaxLength       =   8
            DataType        =   3
            DataField       =   "TDt"
            ReCalcParent    =   "TCNVRT,TCNVRTRS"
            IdName          =   "TDT"
         End
         Begin MwfCtl.MWCTL_MED ATXT 
            Height          =   285
            Index           =   0
            Left            =   13005
            TabIndex        =   45
            ToolTipText     =   "Enter Raw Material Default From Location"
            Top             =   270
            Visible         =   0   'False
            Width           =   210
            _ExtentX        =   370
            _ExtentY        =   503
            Enabled         =   0   'False
            MaxLength       =   5
            DataType        =   4
            DataField       =   "TFrRmLoc"
            IdName          =   "TFRRMLOC"
         End
         Begin MwfCtl.MWCTL_MED ATXT 
            Height          =   285
            Index           =   2
            Left            =   11520
            TabIndex        =   47
            ToolTipText     =   "Bag Location"
            Top             =   270
            Visible         =   0   'False
            Width           =   210
            _ExtentX        =   370
            _ExtentY        =   503
            Enabled         =   0   'False
            MaxLength       =   5
            DataType        =   4
            DataField       =   "TFrBLoc"
            IdName          =   "TFRBLOC"
         End
         Begin MwfCtl.MWCTL_MED ATXT 
            Height          =   285
            Index           =   48
            Left            =   945
            TabIndex        =   6
            ToolTipText     =   "Enter Description"
            Top             =   285
            Width           =   10020
            _ExtentX        =   17674
            _ExtentY        =   503
            MaxLength       =   60
            DataField       =   "TDesc"
            IdName          =   "TDESC"
         End
         Begin MwfCtl.MWCTL_FLX GrdTxndZ 
            Height          =   5100
            Left            =   0
            TabIndex        =   23
            Top             =   1680
            Width           =   15195
            _ExtentX        =   26802
            _ExtentY        =   8996
            Cols            =   24
            colname1        =   "TZSR"
            heading1        =   "Sr.    "
            datafld1        =   "TzSr"
            datatype1       =   1
            mask1           =   "###0"
            maxlength1      =   4
            style1          =   2
            tooltiptext1    =   "Enter Sr No."
            colname2        =   "TZRMZCD"
            heading2        =   "Customs Code                         "
            datafld2        =   "TzRmZCd"
            datatype2       =   4
            recalcparent2   =   "GrdTxndZ(TZRMQTY),GrdTxndZ(TZRMWT),GrdTxndZ(TZPCPERCT),GrdTxndZ(TZPURRT)"
            maxlength2      =   16
            tooltiptext2    =   "Enter Customs Code"
            colname3        =   "TZRMCDDESC"
            heading3        =   "Rm Cd Desc {International Qly}      "
            datafld3        =   "TzRmCdDesc"
            datatype3       =   4
            maxlength3      =   20
            tooltiptext3    =   "Enter Raw Material Code"
            colname4        =   "TZPCPERCT"
            heading4        =   "Pc/Ct "
            datafld4        =   "TzPcPerCt"
            datatype4       =   1
            mask4           =   "##0"
            recalcon4       =   "GrdTxndZ(TZRMZCD),GrdTxndZ(TZRMQTY),GrdTxndZ(TZRMWT)"
            recalcparent4   =   "GrdTxndZ(TZRMQTY)"
            maxlength4      =   3
            tooltiptext4    =   "Enter Average Piece Per Carat"
            colname5        =   "TZRMQTY"
            heading5        =   "Qty      "
            datafld5        =   "TzRmQty"
            datatype5       =   1
            mask5           =   "####0"
            recalcon5       =   "GrdTxndZ(TZPCPERCT),GrdTxndZ(TZRMWT),GrdTxndZ(TZRMZCD)"
            recalcparent5   =   "WTOTQTY,GrdTxndZ(TZPCPERCT)"
            maxlength5      =   5
            tooltiptext5    =   "Enter Raw Material Quantity"
            colname6        =   "TZRMWT"
            heading6        =   "Rm Wt        "
            datafld6        =   "TzRmWt"
            datatype6       =   2
            mask6           =   "####0.000"
            recalcon6       =   "GrdTxndZ(TZRMZCD)"
            recalcparent6   =   "GrdTxndZ(TZRMQTY),GrdTxndZ(TZPURAMT),WTOTWT,GrdTxndZ(WTZRMBALWT),GrdTxndZ(TZPCPERCT),GrdTxndZ(TZCLS)"
            maxlength6      =   9
            tooltiptext6    =   "Enter Raw Material Weight"
            colname7        =   "TZRMDESPWT"
            datafld7        =   "TzRmDespWt"
            datatype7       =   2
            mask7           =   "####0.000"
            recalcparent7   =   "GrdTxndZ(WTZRMBALWT),GrdTxndZ(TZCLS)"
            maxlength7      =   9
            style7          =   0
            tooltiptext7    =   "Raw Material Despatch Weight"
            colname8        =   "WTZRMBALWT"
            heading8        =   "Bal Wt        "
            datatype8       =   2
            mask8           =   "####0.000"
            recalcon8       =   "GrdTxndZ(TZRMWT),GrdTxndZ(TZRMDESPWT)"
            recalcparent8   =   "WBALWT"
            maxlength8      =   9
            style8          =   2
            tooltiptext8    =   "Raw Material Balance Weight"
            colname9        =   "TZPURRT"
            heading9        =   "Pur Rate            "
            datafld9        =   "TzPurRt"
            datatype9       =   2
            mask9           =   "######0.000"
            recalcon9       =   "GrdTxndZ(TZRMZCD),GrdTxndZ(TZPURAMT)"
            recalcparent9   =   "GrdTxndZ(TZPURAMT)"
            maxlength9      =   11
            tooltiptext9    =   "Enter Pur Rate"
            colname10       =   "TZPURAMT"
            heading10       =   "Pur Amount         "
            datafld10       =   "TzPurAmt"
            datatype10      =   2
            mask10          =   "#########0.00"
            recalcon10      =   "GrdTxndZ(TZRMWT),GrdTxndZ(TZPURRT)"
            recalcparent10  =   "TNETAMT,GrdTxndZ(TZPURRT)"
            maxlength10     =   13
            tooltiptext10   =   "Enter Pur Amount"
            colname11       =   "TZCLS"
            heading11       =   "Cls"
            datafld11       =   "TzCls"
            datatype11      =   4
            recalcon11      =   "GrdTxndZ(TZRMWT),GrdTxndZ(TZRMDESPWT)"
            recalcparent11  =   "GrdTxndZ(TZCLSDT)"
            maxlength11     =   1
            tooltiptext11   =   "Enter Whether Purchase is Closed Or Not"
            colname12       =   "TZDESC"
            heading12       =   "Desc"
            datafld12       =   "TzDesc"
            recalcon12      =   "WTZDESC"
            maxlength12     =   60
            style12         =   0
            colname13       =   "TZPRTKEY"
            heading13       =   "Prtn "
            datafld13       =   "TzPrtKey"
            maxlength13     =   1
            style13         =   0
            colname14       =   "TZODTC"
            datafld14       =   "TzOdTc"
            datatype14      =   4
            maxlength14     =   3
            style14         =   0
            colname15       =   "TZODYY"
            datafld15       =   "TzOdYy"
            datatype15      =   4
            maxlength15     =   2
            style15         =   0
            colname16       =   "TZODCHR"
            datafld16       =   "TzOdChr"
            datatype16      =   4
            maxlength16     =   3
            style16         =   0
            colname17       =   "TZODNO"
            datafld17       =   "TzOdNo"
            datatype17      =   1
            mask17          =   "######0"
            maxlength17     =   7
            style17         =   0
            colname18       =   "TZODSR"
            datafld18       =   "TzOdSr"
            datatype18      =   1
            mask18          =   "####0"
            maxlength18     =   5
            style18         =   0
            colname19       =   "TZDMCD"
            datafld19       =   "TzDmCd"
            datatype19      =   4
            maxlength19     =   15
            style19         =   0
            colname20       =   "TZRMCD"
            datafld20       =   "TzRmCd"
            datatype20      =   4
            maxlength20     =   12
            style20         =   0
            colname21       =   "TZRMSZFR"
            datafld21       =   "TzRmSzFr"
            datatype21      =   2
            mask21          =   "#0.0000"
            maxlength21     =   7
            style21         =   0
            colname22       =   "TZRMSZTO"
            datafld22       =   "TzRmSzTo"
            datatype22      =   2
            mask22          =   "#0.0000"
            maxlength22     =   7
            style22         =   0
            colname23       =   "TZCLSDT"
            heading23       =   "Cls Dt      "
            datafld23       =   "TzClsDt"
            datatype23      =   3
            recalcon23      =   "GrdTxndZ(TZCLS)"
            maxlength23     =   8
            tooltiptext23   =   "Purchase Close Date"
         End
         Begin VB.Frame FraPurFot 
            Height          =   855
            Left            =   0
            TabIndex        =   58
            Top             =   7590
            Width           =   15195
            Begin MwfCtl.MWCTL_MED ATXT 
               Height          =   285
               Index           =   18
               Left            =   8310
               TabIndex        =   29
               ToolTipText     =   "Total Amount"
               Top             =   465
               Width           =   1455
               _ExtentX        =   2566
               _ExtentY        =   503
               Alignment       =   2
               Enabled         =   0   'False
               Mask            =   "#########0.00"
               MaxLength       =   13
               DataType        =   2
               DataField       =   "TTotAmt"
               ReCalcOn        =   "TNETAMT,TADDAMT1,TADDAMT2,TADDAMT3,TADDAMT4"
               IdName          =   "TTOTAMT"
            End
            Begin MwfCtl.MWCTL_MED ATXT 
               Height          =   285
               Index           =   19
               Left            =   8310
               TabIndex        =   28
               ToolTipText     =   "Additional Amount 4"
               Top             =   180
               Width           =   1005
               _ExtentX        =   1773
               _ExtentY        =   503
               Alignment       =   2
               Mask            =   "#####0.00"
               MaxLength       =   9
               DataType        =   2
               DataField       =   "TAddAmt4"
               ReCalcParent    =   "TTOTAMT"
               IdName          =   "TADDAMT4"
            End
            Begin MwfCtl.MWCTL_MED ATXT 
               Height          =   285
               Index           =   20
               Left            =   5160
               TabIndex        =   27
               ToolTipText     =   "Additional Amount 3"
               Top             =   465
               Width           =   1005
               _ExtentX        =   1773
               _ExtentY        =   503
               Alignment       =   2
               Mask            =   "#####0.00"
               MaxLength       =   9
               DataType        =   2
               DataField       =   "TAddAmt3"
               ReCalcParent    =   "TTOTAMT"
               IdName          =   "TADDAMT3"
            End
            Begin MwfCtl.MWCTL_MED ATXT 
               Height          =   285
               Index           =   21
               Left            =   5160
               TabIndex        =   26
               ToolTipText     =   "Additional Amount 2"
               Top             =   180
               Width           =   1005
               _ExtentX        =   1773
               _ExtentY        =   503
               Alignment       =   2
               Mask            =   "#####0.00"
               MaxLength       =   9
               DataType        =   2
               DataField       =   "TAddAmt2"
               ReCalcParent    =   "TTOTAMT"
               IdName          =   "TADDAMT2"
            End
            Begin MwfCtl.MWCTL_MED ATXT 
               Height          =   285
               Index           =   23
               Left            =   1830
               TabIndex        =   25
               ToolTipText     =   "Additional Amount 1"
               Top             =   465
               Width           =   1005
               _ExtentX        =   1773
               _ExtentY        =   503
               Alignment       =   2
               Mask            =   "#####0.00"
               MaxLength       =   9
               DataType        =   2
               DataField       =   "TAddAmt1"
               ReCalcParent    =   "TTOTAMT"
               IdName          =   "TADDAMT1"
            End
            Begin MwfCtl.MWCTL_MED ATXT 
               Height          =   285
               Index           =   24
               Left            =   1830
               TabIndex        =   24
               ToolTipText     =   "Net Amount"
               Top             =   180
               Width           =   1455
               _ExtentX        =   2566
               _ExtentY        =   503
               Alignment       =   2
               Enabled         =   0   'False
               Mask            =   "#########0.00"
               MaxLength       =   13
               DataType        =   2
               DataField       =   "TNetAmt"
               ReCalcParent    =   "TTOTAMT"
               ReCalcOn        =   "GrdTxndZ(TZPURAMT)"
               IdName          =   "TNETAMT"
            End
            Begin VB.Label ALBL 
               BackStyle       =   0  'Transparent
               Caption         =   "Net Amount"
               BeginProperty Font 
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
               Left            =   120
               TabIndex        =   64
               Top             =   180
               Width           =   1635
            End
            Begin VB.Label LblTAddAmt1 
               BackStyle       =   0  'Transparent
               Caption         =   "Addn Amount 1"
               BeginProperty Font 
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
               TabIndex        =   63
               Top             =   465
               Width           =   1635
            End
            Begin VB.Label LblTAddAmt2 
               BackStyle       =   0  'Transparent
               Caption         =   "Addn Amount 2"
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   285
               Left            =   3510
               TabIndex        =   62
               Top             =   180
               Width           =   1635
            End
            Begin VB.Label LblTAddAmt3 
               BackStyle       =   0  'Transparent
               Caption         =   "Addn Amount 3"
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   285
               Left            =   3510
               TabIndex        =   61
               Top             =   465
               Width           =   1635
            End
            Begin VB.Label LblTAddAmt4 
               BackStyle       =   0  'Transparent
               Caption         =   "Addn Amount 4"
               BeginProperty Font 
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
               TabIndex        =   60
               Top             =   180
               Width           =   1635
            End
            Begin VB.Label ALBL 
               BackStyle       =   0  'Transparent
               Caption         =   "Tot Amount"
               BeginProperty Font 
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
               Left            =   6660
               TabIndex        =   59
               Top             =   465
               Width           =   1635
            End
         End
         Begin MwfCtl.MWCTL_MED ATXT 
            Height          =   285
            Index           =   9
            Left            =   11940
            TabIndex        =   68
            ToolTipText     =   "Enter Loss From Date"
            Top             =   270
            Visible         =   0   'False
            Width           =   210
            _ExtentX        =   370
            _ExtentY        =   503
            Enabled         =   0   'False
            MaxLength       =   8
            DataType        =   3
            DataField       =   "TLsFrDt"
            IdName          =   "TLSFRDT"
         End
         Begin MwfCtl.MWCTL_MED ATXT 
            Height          =   285
            Index           =   10
            Left            =   12150
            TabIndex        =   69
            ToolTipText     =   "Enter Loss To Date"
            Top             =   270
            Visible         =   0   'False
            Width           =   210
            _ExtentX        =   370
            _ExtentY        =   503
            Enabled         =   0   'False
            MaxLength       =   8
            DataType        =   3
            DataField       =   "TLsToDt"
            IdName          =   "TLSTODT"
         End
         Begin MwfCtl.MWCTL_MED ATXT 
            Height          =   285
            Index           =   7
            Left            =   11730
            TabIndex        =   70
            ToolTipText     =   "Enter Loss Location"
            Top             =   270
            Visible         =   0   'False
            Width           =   210
            _ExtentX        =   370
            _ExtentY        =   503
            Enabled         =   0   'False
            MaxLength       =   5
            DataType        =   4
            DataField       =   "TLsLoc"
            IdName          =   "TLSLOC"
         End
         Begin MwfCtl.MWCTL_MED ATXT 
            Height          =   285
            Index           =   12
            Left            =   12360
            TabIndex        =   71
            ToolTipText     =   "Enter Gold In Loss"
            Top             =   270
            Visible         =   0   'False
            Width           =   210
            _ExtentX        =   370
            _ExtentY        =   503
            Alignment       =   2
            Enabled         =   0   'False
            Mask            =   "####0.000"
            MaxLength       =   9
            DataType        =   2
            DataField       =   "TGldInLs"
            IdName          =   "TGLDINLS"
            Data            =   "99999.999"
         End
         Begin MwfCtl.MWCTL_MED ATXT 
            Height          =   285
            Index           =   8
            Left            =   12570
            TabIndex        =   72
            ToolTipText     =   "Enter Gold Recovered"
            Top             =   270
            Visible         =   0   'False
            Width           =   210
            _ExtentX        =   370
            _ExtentY        =   503
            Alignment       =   2
            Enabled         =   0   'False
            Mask            =   "####0.000"
            MaxLength       =   9
            DataType        =   2
            DataField       =   "TGldRec"
            IdName          =   "TGLDREC"
         End
         Begin MwfCtl.MWCTL_MED ATXT 
            Height          =   285
            Index           =   26
            Left            =   12780
            TabIndex        =   73
            ToolTipText     =   "Enter Dust Weight"
            Top             =   270
            Visible         =   0   'False
            Width           =   210
            _ExtentX        =   370
            _ExtentY        =   503
            Alignment       =   2
            Enabled         =   0   'False
            Mask            =   "####0.000"
            MaxLength       =   9
            DataType        =   2
            DataField       =   "TDustWt"
            IdName          =   "TDUSTWT"
         End
         Begin MwfCtl.MWCTL_MED ATXT 
            Height          =   285
            Index           =   11
            Left            =   945
            TabIndex        =   7
            ToolTipText     =   "Enter Supplier Code"
            Top             =   570
            Visible         =   0   'False
            Width           =   1335
            _ExtentX        =   2355
            _ExtentY        =   503
            MaxLength       =   8
            DataType        =   4
            DataField       =   "TSuppCd"
            ReCalcParent    =   "WCMNAME,TCURRCD"
            ReCalcOn        =   "WCMNAME"
            IdName          =   "TSUPPCD"
         End
         Begin MwfCtl.MWCTL_MED ATXT 
            Height          =   285
            Index           =   13
            Left            =   945
            TabIndex        =   12
            ToolTipText     =   "Enter Bill No"
            Top             =   855
            Visible         =   0   'False
            Width           =   3165
            _ExtentX        =   5583
            _ExtentY        =   503
            MaxLength       =   20
            DataType        =   4
            DataField       =   "TBillNo"
            IdName          =   "TBILLNO"
         End
         Begin MwfCtl.MWCTL_MED ATXT 
            Height          =   285
            Index           =   15
            Left            =   5490
            TabIndex        =   13
            ToolTipText     =   "Enter Bill Date"
            Top             =   855
            Visible         =   0   'False
            Width           =   855
            _ExtentX        =   1508
            _ExtentY        =   503
            MaxLength       =   8
            DataType        =   3
            DataField       =   "TBillDt"
            IdName          =   "TBILLDT"
         End
         Begin MwfCtl.MWCTL_MED ATXT 
            Height          =   285
            Index           =   16
            Left            =   5490
            TabIndex        =   19
            ToolTipText     =   "Enter Bill Of Entry Date"
            Top             =   1140
            Visible         =   0   'False
            Width           =   855
            _ExtentX        =   1508
            _ExtentY        =   503
            Enabled         =   0   'False
            MaxLength       =   8
            DataType        =   3
            DataField       =   "TBOEDt"
            IdName          =   "TBOEDT"
         End
         Begin MwfCtl.MWCTL_MED ATXT 
            Height          =   285
            Index           =   17
            Left            =   12345
            TabIndex        =   10
            ToolTipText     =   "Enter Conversion Rate"
            Top             =   570
            Visible         =   0   'False
            Width           =   1035
            _ExtentX        =   1826
            _ExtentY        =   503
            Alignment       =   2
            Mask            =   "###0.0000"
            MaxLength       =   9
            DataType        =   2
            DataField       =   "TCNVRT"
            ReCalcOn        =   "TCURRCD,TDT"
            IdName          =   "TCNVRT"
         End
         Begin MwfCtl.MWCTL_MED ATXT 
            Height          =   285
            Index           =   14
            Left            =   945
            TabIndex        =   18
            ToolTipText     =   "Enter Bill Of Entry Number"
            Top             =   1140
            Visible         =   0   'False
            Width           =   3165
            _ExtentX        =   5583
            _ExtentY        =   503
            MaxLength       =   20
            DataType        =   4
            DataField       =   "TBOENo"
            IdName          =   "TBOENO"
         End
         Begin MwfCtl.MWCTL_MED ATXT 
            Height          =   285
            Index           =   5
            Left            =   7695
            TabIndex        =   14
            ToolTipText     =   "Enter Date Before Which The Goods Have To Be Used"
            Top             =   855
            Visible         =   0   'False
            Width           =   855
            _ExtentX        =   1508
            _ExtentY        =   503
            MaxLength       =   8
            DataType        =   3
            DataField       =   "TUseBefDt"
            IdName          =   "TUSEBEFDT"
         End
         Begin MwfCtl.MWCTL_MED ATXT 
            Height          =   285
            Index           =   25
            Left            =   5490
            TabIndex        =   9
            ToolTipText     =   "Enter Supplier Name"
            Top             =   570
            Visible         =   0   'False
            Width           =   5475
            _ExtentX        =   9657
            _ExtentY        =   503
            MaxLength       =   45
            ReCalcParent    =   "TSUPPCD"
            ReCalcOn        =   "TSUPPCD"
            IdName          =   "WCMNAME"
         End
         Begin MwfCtl.MWCTL_MED ATXT 
            Height          =   285
            Index           =   29
            Left            =   10110
            TabIndex        =   15
            ToolTipText     =   "Enter Import Type"
            Top             =   855
            Visible         =   0   'False
            Width           =   855
            _ExtentX        =   1508
            _ExtentY        =   503
            MaxLength       =   5
            DataType        =   4
            DataField       =   "TImpTyp"
            IdName          =   "TIMPTYP"
         End
         Begin MwfCtl.MWCTL_MED ATXT 
            Height          =   285
            Index           =   99
            Left            =   13410
            TabIndex        =   90
            ToolTipText     =   "Enter Quantity To Be Transferred"
            Top             =   270
            Visible         =   0   'False
            Width           =   195
            _ExtentX        =   344
            _ExtentY        =   503
            Enabled         =   0   'False
            Mask            =   "##0.0"
            MaxLength       =   5
            DataType        =   2
            DataField       =   "TTRFQTY"
            IdName          =   "TTRFQTY"
         End
         Begin MwfCtl.MWCTL_MED ATXT 
            Height          =   285
            Index           =   100
            Left            =   10950
            TabIndex        =   91
            Top             =   0
            Visible         =   0   'False
            Width           =   240
            _ExtentX        =   423
            _ExtentY        =   503
            Enabled         =   0   'False
            MaxLength       =   30
            DataField       =   "TSEOPWD"
            IdName          =   "TSEOPWD"
         End
         Begin MwfCtl.MWCTL_MED ATXT 
            Height          =   285
            Index           =   109
            Left            =   6885
            TabIndex        =   4
            ToolTipText     =   "Enter Order Lock (Y/N)"
            Top             =   0
            Width           =   345
            _ExtentX        =   609
            _ExtentY        =   503
            DataType        =   4
            DataField       =   "TLockYN"
            IdName          =   "TLOCKYN"
         End
         Begin MwfCtl.MWCTL_MED ATXT 
            Height          =   285
            Index           =   39
            Left            =   3225
            TabIndex        =   8
            ToolTipText     =   "Enter Currency Code"
            Top             =   570
            Width           =   885
            _ExtentX        =   1561
            _ExtentY        =   503
            Enabled         =   0   'False
            MaxLength       =   5
            DataType        =   4
            DataField       =   "TCURRCD"
            ReCalcParent    =   "TCNVRT,TMULDIV,TCNVRTRS,TMULDIVRS"
            ReCalcOn        =   "TSUPPCD"
            IdName          =   "TCURRCD"
         End
         Begin MwfCtl.MWCTL_MED ATXT 
            Height          =   285
            Index           =   33
            Left            =   13620
            TabIndex        =   97
            Top             =   270
            Visible         =   0   'False
            Width           =   195
            _ExtentX        =   344
            _ExtentY        =   503
            Enabled         =   0   'False
            MaxLength       =   12
            DataType        =   4
            DataField       =   "TLOTNO"
            IdName          =   "TLOTNO"
         End
         Begin MwfCtl.MWCTL_MED ATXT 
            Height          =   285
            Index           =   55
            Left            =   14850
            TabIndex        =   11
            ToolTipText     =   "Enter Multiply/Divide (M/D)"
            Top             =   570
            Width           =   345
            _ExtentX        =   609
            _ExtentY        =   503
            Enabled         =   0   'False
            DataType        =   4
            DataField       =   "TMulDiv"
            ReCalcOn        =   "TCURRCD"
            IdName          =   "TMULDIV"
         End
         Begin MwfCtl.MWCTL_MED ATXT 
            Height          =   285
            Index           =   36
            Left            =   12345
            TabIndex        =   16
            ToolTipText     =   "Enter Rupees Conversion Rate"
            Top             =   855
            Width           =   1035
            _ExtentX        =   1826
            _ExtentY        =   503
            Alignment       =   2
            Mask            =   "###0.0000"
            MaxLength       =   9
            DataType        =   2
            DataField       =   "TCNVRTRS"
            ReCalcOn        =   "TCURRCD,TDT"
            IdName          =   "TCNVRTRS"
         End
         Begin MwfCtl.MWCTL_MED ATXT 
            Height          =   285
            Index           =   37
            Left            =   14850
            TabIndex        =   17
            ToolTipText     =   "Enter Rupees Multiply/Divide (M/D)"
            Top             =   855
            Width           =   345
            _ExtentX        =   609
            _ExtentY        =   503
            Enabled         =   0   'False
            DataType        =   4
            DataField       =   "TMulDivRs"
            ReCalcOn        =   "TCURRCD"
            IdName          =   "TMULDIVRS"
         End
         Begin MwfCtl.MWCTL_MED ATXT 
            Height          =   285
            Index           =   43
            Left            =   7695
            TabIndex        =   20
            ToolTipText     =   "Enter Customer Code"
            Top             =   1140
            Width           =   1335
            _ExtentX        =   2355
            _ExtentY        =   503
            MaxLength       =   8
            DataType        =   4
            DataField       =   "TCmCd"
            IdName          =   "TCMCD"
         End
         Begin MwfCtl.MWCTL_MED ATXT 
            Height          =   285
            Index           =   52
            Left            =   10110
            TabIndex        =   21
            ToolTipText     =   "Enter Import Category"
            Top             =   1140
            Width           =   855
            _ExtentX        =   1508
            _ExtentY        =   503
            MaxLength       =   3
            DataType        =   4
            DataField       =   "TImpCtg"
            IdName          =   "TIMPCTG"
         End
         Begin MwfCtl.MWCTL_MED ATXT 
            Height          =   285
            Index           =   53
            Left            =   13800
            TabIndex        =   128
            ToolTipText     =   "Enter Multi Issue/ receipt"
            Top             =   270
            Visible         =   0   'False
            Width           =   195
            _ExtentX        =   344
            _ExtentY        =   503
            Enabled         =   0   'False
            DataType        =   4
            DataField       =   "TMultiIr"
            IdName          =   "TMULTIIR"
         End
         Begin MwfCtl.MWCTL_MED ATXT 
            Height          =   285
            Index           =   86
            Left            =   945
            TabIndex        =   22
            ToolTipText     =   "Enter Pay Terms"
            Top             =   1425
            Width           =   7035
            _ExtentX        =   12409
            _ExtentY        =   503
            MaxLength       =   45
            DataField       =   "TSuppPayTerms"
            IdName          =   "TSUPPPAYTERMS"
         End
         Begin VB.Label ALBL 
            BackStyle       =   0  'Transparent
            Caption         =   "Pay Terms"
            BeginProperty Font 
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
            TabIndex        =   129
            Top             =   1440
            Width           =   1005
         End
         Begin VB.Label ALBL 
            BackStyle       =   0  'Transparent
            Caption         =   "Import Ctg"
            BeginProperty Font 
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
            Left            =   9045
            TabIndex        =   127
            Top             =   1140
            Width           =   1200
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
            Index           =   16
            Left            =   6570
            TabIndex        =   125
            Top             =   1140
            Width           =   930
         End
         Begin VB.Label ALBL 
            BackStyle       =   0  'Transparent
            Caption         =   "Rs. Conv Rt"
            BeginProperty Font 
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
            Left            =   11100
            TabIndex        =   101
            Top             =   855
            Width           =   1170
         End
         Begin VB.Label Label5 
            BackStyle       =   0  'Transparent
            Caption         =   "Rs. Mul/Div"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Left            =   13650
            TabIndex        =   100
            Top             =   855
            Width           =   1155
         End
         Begin VB.Label Label4 
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
            Left            =   13650
            TabIndex        =   99
            Top             =   570
            Width           =   825
         End
         Begin VB.Label LblUsdYn 
            BackStyle       =   0  'Transparent
            Caption         =   "Curr Cd."
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Left            =   2370
            TabIndex        =   94
            Top             =   570
            Width           =   825
         End
         Begin VB.Label ALBL 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "Password"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Index           =   99
            Left            =   7500
            TabIndex        =   93
            Top             =   0
            Width           =   1095
         End
         Begin VB.Label ALBL 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "Pur Lock"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Index           =   98
            Left            =   5880
            TabIndex        =   92
            Top             =   0
            Width           =   975
         End
         Begin VB.Label ALBL 
            BackStyle       =   0  'Transparent
            Caption         =   "Import Typ"
            BeginProperty Font 
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
            Left            =   9045
            TabIndex        =   89
            Top             =   855
            Width           =   1200
         End
         Begin VB.Label ALBL 
            BackStyle       =   0  'Transparent
            Caption         =   "Supplier Nm"
            BeginProperty Font 
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
            Left            =   4230
            TabIndex        =   82
            Top             =   570
            Width           =   1320
         End
         Begin VB.Label ALBL 
            BackStyle       =   0  'Transparent
            Caption         =   "Use Bef Dt"
            BeginProperty Font 
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
            Left            =   6570
            TabIndex        =   81
            Top             =   855
            Width           =   1200
         End
         Begin VB.Label ALBL 
            BackStyle       =   0  'Transparent
            Caption         =   "Supplier"
            BeginProperty Font 
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
            Left            =   30
            TabIndex        =   79
            Top             =   570
            Width           =   930
         End
         Begin VB.Label ALBL 
            BackStyle       =   0  'Transparent
            Caption         =   "Bill No"
            BeginProperty Font 
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
            Left            =   30
            TabIndex        =   78
            Top             =   855
            Width           =   930
         End
         Begin VB.Label ALBL 
            BackStyle       =   0  'Transparent
            Caption         =   "BOE No"
            BeginProperty Font 
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
            Left            =   30
            TabIndex        =   77
            Top             =   1140
            Width           =   930
         End
         Begin VB.Label ALBL 
            BackStyle       =   0  'Transparent
            Caption         =   "Conv Rt"
            BeginProperty Font 
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
            Left            =   11100
            TabIndex        =   76
            Top             =   570
            Width           =   930
         End
         Begin VB.Label ALBL 
            BackStyle       =   0  'Transparent
            Caption         =   "Bill Dt"
            BeginProperty Font 
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
            Left            =   4230
            TabIndex        =   75
            Top             =   855
            Width           =   870
         End
         Begin VB.Label ALBL 
            BackStyle       =   0  'Transparent
            Caption         =   "BOE Dt"
            BeginProperty Font 
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
            Left            =   4230
            TabIndex        =   74
            Top             =   1140
            Width           =   870
         End
         Begin VB.Label LblTDesc 
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
            Left            =   30
            TabIndex        =   56
            Top             =   300
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
            Left            =   4110
            TabIndex        =   53
            Top             =   0
            Width           =   855
         End
      End
   End
   Begin VB.Frame FraCD 
      Height          =   1035
      Left            =   5835
      TabIndex        =   130
      ToolTipText     =   "#0.0000"
      Top             =   8280
      Visible         =   0   'False
      Width           =   3015
      Begin MwfCtl.MWCTL_MED ATXT 
         Height          =   285
         Index           =   212
         Left            =   1350
         TabIndex        =   131
         ToolTipText     =   "Enter New Voucher Date"
         Top             =   465
         Width           =   855
         _ExtentX        =   1508
         _ExtentY        =   503
         MaxLength       =   8
         DataType        =   3
         IdName          =   "WTDT"
      End
      Begin MwfCtl.MWCTL_BTN1 CmdCDGo 
         Height          =   495
         Left            =   2280
         TabIndex        =   132
         ToolTipText     =   "Press This Button To Change Date"
         Top             =   360
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
      Begin VB.Label LblMain 
         BackStyle       =   0  'Transparent
         BorderStyle     =   1  'Fixed Single
         Caption         =   "  Change Date"
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
         Index           =   12
         Left            =   -30
         TabIndex        =   134
         Top             =   90
         Width           =   3285
      End
      Begin VB.Label ALBL 
         BackStyle       =   0  'Transparent
         Caption         =   "New Date"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Index           =   169
         Left            =   150
         TabIndex        =   133
         Top             =   465
         Width           =   1245
      End
   End
End
Attribute VB_Name = "EmrFrmCustmPur"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'----- General comments on the Entry  ----
' X- Txn, A-TxndZ(GrdTxndZ)
'
' This entry is used for entering the Custom Purchase details. These custom purchase records
'   are further matched against an invoice in the Custom Matching entry.
'
' If the Date of the current entry is less than that of the lock date (stored in Param)
'   then the entry is not allowed to be edited or deleted. The user can only read the entry.
'
'-----

'*** Declare the variables ms_LockDt, mf_DtLocked for the Lock Date Check
Dim ms_LockDt As String
Dim ms_FutureDt As String
Dim ms_PurgeDt As String
Dim mf_DtLocked As Boolean

Option Explicit
Dim moCn As MwfLib.MDOConnection   '*** (Bef speed) bef 24/11/06
Dim ms_TcTyp As String, ms_TCoCd As String
Dim ms_TTc As String, ms_TYy As String, ms_TChr As String
Dim ms_PwdFlag As Boolean '*** Jay 2.14 *** (SEO Change)(Name Changed)
Enum en_CustmPurFra
  Cpy = 0
  Pwd = 1
  DET = 2                 '*** (Jen 3.1.2)
  ChgDt = 3               '4.1.3.0
End Enum
Dim mArr_Cv As Variant    '*** (Jen 3.1.2)
'Dim ms_Pwd As String     '*** Jay 2.14 *** (SEO Change)
Dim mb_BestUse As Boolean   '*** (Jen 3.1.2)

Private Sub ADC_Load()
  '*** Set the First Key Control property to the field TChr
  '*** Set the First Non Key Control property to TDt
  '*** Set the Child Property Of ADC to GrdTxndZ
  '*** Set the Previous and Next Control for the Grid GrdTxndZ as TBOEDt and ADC respectively
  '*** Set the Hot Keys for the Grid GrdInvFgd as 'A'
  '*** Store the defa values for Tc,Yy in variables ms_TTc & ms_TYy respectively
  
  ReDim mArr_Cv(3)      '*** (Jen 3.1.2)
  
  Set moCn = ADC.Connection    '*** (Bef speed) bef 24/11/06
  Set ADC.FirKeyCtl = ADC("TChr")
  Set ADC.FirNKeyCtl = ADC("TDt")
  ADC.Child = "GrdTxndZ"
  Set GrdTxndZ.PrevCtl = ADC("TBOEDt")
  Set GrdTxndZ.NextCtl = ADC
  GrdTxndZ.HotKey = "A"
  FraNKeyAll.ZOrder (0)
  FraGrdTxndZDets.ZOrder (0)  '*** (Jen 3.1.2)
  '*** ms_TcTyp = 'Tz'
  
  ADC("TBoeDt").Enabled = False         '4.1.3.0
  
  ms_TCoCd = gs_CoCd
  ms_TcTyp = ADC.MenuCd
  ms_TTc = moCn.GetFldVal("Select PValue from Param where PTyp= 'TCTYP' and " + _
               "PMCd= '" + ms_TcTyp + "' ")
  ms_TYy = moCn.GetFldVal("Select hYy from Head where HCoCd= '" + ms_TCoCd + "' and HCd= '" + ctSelfCmCd + "'")
  
  Call PropSet
  
  ms_PwdFlag = False
  If moCn.GetFldVal("Select PValue1 from Param where " + _
                    "Ptyp = 'SEOPWD' and PmCd = 'SEOPWD'") = "" Then ms_PwdFlag = True
                    
  '****** Sachin 2.14.0 - [24-11-2006] - [Multi-Company] ******
  Me.Caption = GetFrmCaption(gs_CoCd, "Custom Purchase")
  '****** Sachin 2.14.0 - [24-11-2006] - [Multi-Company] ******

  '*** (Jen 3.1.2)
  mb_BestUse = moCn.RecSeek("Select 'x' from Head where HCoCd= '" + gs_CoCd + "' and " + "HCd= '" + ctSelfCmCd + "' and hBestUse= 'Y'")
  If mb_BestUse Then
    LblwTzOrd.Visible = True:: LblOrdSlash1.Visible = True: LblOrdSlash2.Visible = True: LblOrdSlash3.Visible = True
    LblwTzDmCd.Visible = True: LblwTzRmCd.Visible = True: LblwTzRmSzFr.Visible = True: LblwTzRmSzTo.Visible = True
    ADC("wTzOdTc").Visible = True: ADC("wTzOdYy").Visible = True: ADC("wTzOdChr").Visible = True
    ADC("wTzOdNo").Visible = True: ADC("wTzDmCd").Visible = True: ADC("wTzRmCd").Visible = True
    'kbs.132 if bestuse mod is yes then fields will display
    ADC("wTzOdSr").Visible = True: LblOrdSlash4.Visible = True
    ADC("wTzRmSzFr").Visible = True: ADC("wTzRmSzTo").Visible = True
    ADC("wCpBsSzYn").Visible = True: ADC("wCpBsSzYn").Enabled = True: LblCpBsSz.Visible = True   ' **** Manali 3.6.0 - 19/09/09 - Copy Best Use Size fields Y/N
  Else
    LblwTzOrd.Visible = False: LblOrdSlash1.Visible = False: LblOrdSlash2.Visible = False: LblOrdSlash3.Visible = False
    LblwTzDmCd.Visible = False: LblwTzRmCd.Visible = False: LblwTzRmSzFr.Visible = False: LblwTzRmSzTo.Visible = False
    ADC("wTzOdTc").Visible = False: ADC("wTzOdYy").Visible = False: ADC("wTzOdChr").Visible = False
    'kbs.132 if bestuse mod is no then fields will not display
    ADC("wTzOdSr").Visible = False: LblOrdSlash4.Visible = False
    ADC("wTzOdNo").Visible = False: ADC("wTzDmCd").Visible = False: ADC("wTzRmCd").Visible = False
    ADC("wTzRmSzFr").Visible = False: ADC("wTzRmSzTo").Visible = False
    ADC("wCpBsSzYn").Visible = False: ADC("wCpBsSzYn").Enabled = False: LblCpBsSz.Visible = False    ' **** Manali 3.6.0 - 19/09/09 - Copy Best Use Size fields Y/N
  End If
  '*** (Jen 3.1.2)
  
  '****** Sachin 3.02.0 - Hiding the Partition Key Column based on Partition Selected
  ADC("TPRTKEY").Visible = IIF(gs_Partition = ctCurrPrtn, False, True)
  GrdTxndZ.ColProp("TzPrtKey").Style = IIF(gs_Partition = ctCurrPrtn, fgcinvisible, fgcReadOnly)
  'sv.9 getting purge date from vparam
  ms_PurgeDt = GetPurgeDt(gs_CoCd, "")
End Sub

Private Sub ADC_KeyWhen()
  '*** Set the dafault value for the form level variable mf_DtLocked as false.
  '*** setting the default values for TTc, TYy and TChr from the previous entry
  '*** Hide all frames except Frame FraNKeyAll

  Call EnaDisaCmds(True) '###
  mf_DtLocked = False
  Call HideAllFras
  
  If ms_TCoCd <> "" Then ADC("TCoCd") = ms_TCoCd
  If ms_TTc <> "" Then ADC("TTc") = ms_TTc
  If ms_TYy <> "" Then ADC("TYy") = ms_TYy
  If ms_TChr <> "" Then ADC("TChr") = ms_TChr
' ########################  Manoj #### Ver: 2.0.6 #### Date: 25/12/2003  ########################
  TxtPwd.text = ""
' ###########################################  Manoj  ###########################################

End Sub
Private Sub ADC_KeyValid(Cancel As Boolean, ErrMsg As String)
  '*** Checking Key values in Key Valid as they are copied in Init
  '*** Check whether the TYy, TChr are blank
  
  If ADC("TYy") = "" Then Cancel = True: ErrMsg = "Enter Voucher Year": Exit Sub
  If ADC("TChr") = "" Then Cancel = True: ErrMsg = "Enter Voucher Character": Exit Sub
End Sub

Private Sub ADC_NonKeyWhen(Cancel As Boolean, ErrMsg As String)
  '*** Get the Lock Date from Param('LOCKDT', 'LOCKDT') and store in the form level variable ms_LockDt
  '    If ms_LockDt is blank then store '01/01/80' in ms_LockDt
  '*** Store the TTc, TYy and TChr values for the next entry
  '*** Enable All Command Buttons Except Scope Button
  '*** give default values to the fields TDt, TLsFrDt, TLsToDt, TBillDt, TBOEDt,
  '    TFrBLoc, TFrRmLoc, TToRmLoc, TUseBefDt
  '*** Set the Next Key Control Property of the Grid GrdTxndZ based on the visibility of
  '    the First Additional Amount Field(TAddAmt1).
  '*** If the entry Date is < ms_LockDt then mf_DtLocked= true (i.e the entry
  '    If mf_DtLocked = true then the entry is locked. i.e. the user cannot add, edit or
  '    delete records in the grid GrdTxndZ

      If ADC.IsNew And (Not moCn.RecSeek("select vPSCd from vParam where " + _
                         "vPCoCd= '" + ADC("TCoCd") + "' and vPTyp= 'CHR' and " + _
                         "vPMCd = '" + ADC("TTc") + "' and vPSCd='" + ADC("TChr") + "' and vPValidYn<>'N'")) Then
       Cancel = True: ErrMsg = "Voucher Character No More Valid": ADC.AllowSave = False: Exit Sub
      Else
        ADC.AllowSave = True
      End If

  '****** Sachin 3.02 26-12-07 - Setting Partition Key to Current
  If ADC.IsNew And Trim(ADC("TPrtKey")) = "" Then ADC("TPrtKey") = ctCurrPrtn

'  ms_LockDt = moCn.GetFldVal("Select PDesc225 from Param where PTyp= 'LOCKDT' and PMCd= 'LOCKDT'")
'  ms_FutureDt = moCn.GetFldVal("Select PDesc225 from Param where PTyp= 'LOCKDT' and PMCd= 'FUTUREDT'")
  
  '6.1_51 wVchrNo added to GetLockDt to check that given voucher number is exist in unlock ptyp,
  'if given voucher is in unlock ptyp it will not check lock date
  Dim wVchrNo As String
  wVchrNo = ADC("TTc") + "/" + ADC("TYy") + "/" + ADC("TChr") + "/" + CStr(ADC("TNo"))
  
  Call GetLockDt(ms_TCoCd, "TC", ms_TTc, ms_LockDt, ms_FutureDt, wVchrNo)
  
  '****** Sachin 3.02 26-12-07 - Setting Purge Date
  'ms_PurgeDt = moCn.GetFldVal("Select PDesc225 from Param where PTyp= 'PURGEDT' and PMCd= 'PURGEDT'")
  
'  If ms_LockDt = "" Then ms_LockDt = "01/01/80"
'  If ms_FutureDt = "" Then ms_FutureDt = "31/12/28"
  
  '****** Sachin 3.02 26-12-07 - Setting Purge Date
  'If ms_PurgeDt = "" Then ms_PurgeDt = "01/01/80"

  ms_TTc = ADC("TTc")
  ms_TYy = ADC("TYy")
  ms_TChr = ADC("TChr")
  
  ' Set The Visible, Enable Properties For All The command buttons ***
  Call EnaDisaCmds(False) '###

  If ADC.AddFndMode = xaddmode Then
    '53 company code parameter added
    ADC("TDt") = GetNextDate(ADC("TCoCd"))   'moCn.SrvrDate   'Date
    ADC("TBillDt") = MWLib.EmptyDate
    ADC("TBOEDt") = MWLib.EmptyDate
    ADC("TUseBefDt") = moCn.SrvrDate
    
    ADC("TLsFrDt") = MWLib.EmptyDate
    ADC("TLsToDt") = MWLib.EmptyDate
    ADC("TFrBLoc") = ""
    ADC("TFrRmLoc") = ""
    ADC("TToRmLoc") = ""
    '***********Geeta****Emr207*************
    '****** Sachin 2.13.0 [18-07-2006] - Multi-Currency ******
    ADC("TCurrCd") = moCn.GetFldVal("Select hBaseCurCd From Head Where HCd='" + ctSelfCmCd + "' And HCoCd='" + gs_CoCd + "'")
    '****** Sachin 2.13.0 [18-07-2006] - Multi-Currency ******
    '***********Geeta****Emr207*************
    '*** Jay 2.14 *** (SEO Change)
    ADC("TLockYN") = "N"
    '*** Jay 2.14 *** (SEO Change)
    
    ADC("TImpCtg") = ctImpCtgImp ' Manali 3.8.0 - ImpCtg
  End If
  
    
	mf_DtLocked = IIF((ADC("TDt") < CDate(ms_LockDt)) _
                               Or _
                   (ADC("TDt") > CDate(ms_FutureDt)), _
                True, False)
  
    '******************Geeta********Emr206******02/04/04****
  '*****************added adc("TLockYn") check*************
  If ADC("TPrtKey") <> ctCurrPrtn Or mf_DtLocked = True Or ADC("TLockYN") = "Y" Then
    GrdTxndZ.AllowAdd = False: GrdTxndZ.AllowDelete = False
  Else
    GrdTxndZ.AllowAdd = True: GrdTxndZ.AllowDelete = True
  End If
  
  If ADC("TAddAmt1").Visible Then
    Set GrdTxndZ.NextCtl = ADC("TAddAmt1")
  Else
    Set GrdTxndZ.NextCtl = ADC
  End If
'*** Jay 2.14 *** (SEO Change)
  If ADC("TLockYN") = "Y" Then
    'ms_PwdFlag = False     '*** (Bef 3.2.0)
    ADC.AllowDelete = False
  Else
    'ms_PwdFlag = True      '*** (Bef 3.2.0)
    ADC.AllowDelete = True
  End If
'' ########################  Manoj #### Ver: 2.0.6 #### Date: 25/12/2003  ########################
'  If ADC("TLockYN") = "" Then ADC("TLockYN") = "N"
'  ms_Pwd = moCn.GetFldVal("Select PValue1 from Param where " + _
'                               "Ptyp = 'SEOPWD' and PmCd = 'SEOPWD'")
'' ###########################################  Manoj  ###########################################
''******************Geeta********Emr206******02/04/04****
'  '*****************added adc("TLockYn") check*************
'  If ADC("TLockYN") = "Y" Then
'    ADC.AllowDelete = False
'  Else
'    ADC.AllowDelete = True
'  End If
'*** Jay 2.14 *** (SEO Change)

ADC("wCpBsSzYn") = "N" ' **** Manali 3.6.0 - 19/09/09 - Copy Best Use Size fields Y/N
End Sub


Private Sub ATXT_FldWhen(Index As Integer, ByVal IdName As String, Cancel As Boolean, ErrMsg As String)
  '*** Set Helps For Each Field
  '*** Do not allow the user to edit the voucher date if record exists in the grid
  
  If (ADC("TPrtKey") <> ctCurrPrtn And ADC("TPrtKey") <> "") Or mf_DtLocked = True Then
    Cancel = True: ErrMsg = "Child Records Exist / Cannot Edit Locked Entry": Exit Sub
  End If
  
  ' Zubin 213
  If GrdTxndZ.Rows > 1 Then
    Select Case IdName
      Case Is = UCase("TDt"), UCase("TSuppCd"), UCase("wCmName"), UCase("TCnvRt"), UCase("TCmCd")
        '*** (Jen 3.1.1 added TCmCd)
        Cancel = True: ErrMsg = "Cannot Edit Value When Child Records Exist": Exit Sub
    End Select
  End If
  
' ########################  Manoj #### Ver: 2.0.6 #### Date: 25/12/2003  ########################
  If UCase$(ADC("TLockYN")) = UCase$("Y") And UCase$(IdName) <> UCase$("TLockYN") Then
    Cancel = True: ErrMsg = "Cannot Edit Locked Entry": Exit Sub
  End If
' ###########################################  Manoj  ###########################################
  
  Select Case IdName
  Case Is = UCase("wCpTCoCd")
    '*** Jen (2.12 Patch)
    Call HlpList.hCoCd
    '*** Jen (2.12 Patch)
  Case Is = UCase("wCpTTc")
    Call HlpList.PMCd("TC", "'PR'")
  Case Is = UCase("TYy"), UCase("wCpTYy")
    Call HlpList.PMCd("YY")
  Case Is = UCase("TChr")
' zubin
    'Call HlpList.vPSCd(ADC("TCoCd"), "CHR", ADC("TTc"))
    Call HlpList.vPSCd(ADC("TCoCd"), "CHR", ADC("TTc"), gs_UsrCd)
    
  Case Is = UCase("wCpTChr")
    '*** Jen (2.12 Patch)
    Call HlpList.vPSCd(ADC("wCpTCoCd"), "CHR", ADC("wCpTTc"))
    '*** Jen (2.12 Patch)
    '*** (Bef 25/03/06) Call HlpList.vPSCd(ADC("TCoCd"), "CHR", ADC("wCpTTc"))
  Case Is = UCase("TNo")
    Call HlpList.TNo(ADC("TCoCd"), ADC("TTc"), ADC("TYy"), ADC("TChr"))
  Case Is = UCase("wCpTNo")
    '*** Jen (2.12 Patch)
    Call HlpList.TNo(ADC("wCpTCoCd"), ADC("wCpTTc"), ADC("wCpTYy"), ADC("wCpTChr"))
    '*** Jen (2.12 Patch)
    '*** (Bef 25/03/06) Call HlpList.TNo(ADC("TCoCd"), ADC("wCpTTc"), ADC("wCpTYy"), ADC("wCpTChr"))
  
  Case Is = UCase("TDt")
    If GrdTxndZ.Rows - 1 > 0 Then Cancel = True: ErrMsg = "Cannot Edit Date, Child Record Exists": Exit Sub
  Case Is = UCase("TSuppCd")
    Call HlpList.CustCd("S")
' ########################  Manoj #### Ver: 2.0.6 #### Date: 25/12/2003  ########################
  Case Is = UCase("TLockYN")
    '*** Jay 2.14 *** (SEO Change)
    'If Trim$(TxtPwd.text) <> Trim$(ms_Pwd) And ADC("TLockYN") = "Y" Then Cancel = True: ErrMsg = "Enter SEO Password to Edit Value": Exit Sub
    
    '*** (Bef 3.2.0)
    'If ADC("TLockYN") = "Y" And ms_PwdFlag = False And Trim$(TxtPwd.text) = "" Then
    '*** (Bef 3.2.0)
    '****** (Jen 3.2.0)
    If ADC("TLockYN") = "Y" And ms_PwdFlag = False Then _
       Cancel = True: ErrMsg = "Enter SEO Password to Edit Value": Exit Sub
       Call HlpList.PMCd("YN")
    '*** Jay 2.14 *** (SEO Change)
  '###########################################  Manoj  ###########################################

  Case Is = UCase("wCmName")
    Call HlpList.CmName("S")
  Case Is = UCase("wCpGrp")
    Call HlpList.PMCd("PURCPY")
  
  '************************ urmi *****************************
  Case Is = UCase("TImpTyp")
' **** Zubin 211 **** '
'''''    If GrdTxndZ.Rows - 1 > 0 Then Cancel = True: ErrMsg = "Cannot Edit Import Type, Child Record Exists": Exit Sub
' **** Zubin 211 **** '
    Call HlpList.PMCd("ImpTyp")
  '************************ urmi *****************************
  '****** Sachin 2.13.0 [18-07-2006] - Multi-Currency ******
  Case Is = UCase("TCurrCd")
    If GrdTxndZ.Rows - 1 > 0 Then Cancel = True: ErrMsg = "Cannot Edit Currency Code, Child Records Exist": Exit Sub
    Call HlpList.PMCd("CURNCY")
  
  Case Is = UCase("TCnvRt")
    If GrdTxndZ.Rows - 1 > 0 Then Cancel = True: ErrMsg = "Cannot Edit Conversion Factor, Child Records Exist": Exit Sub
  '****** Sachin 2.13.0 [18-07-2006] - Multi-Currency ******
  
  '*** (Jen 3.1.1)
  Case Is = UCase("TCmCd")
    Call HlpList.CustCd("C")
  '*** (Jen 3.1.1)
  
  '*** (Jen 3.1.2)
  Case Is = UCase("wTzOdTc")
    Call HlpList.PMCd("TC", "'SO'")
  Case Is = UCase("wTzOdYy")
    Call HlpList.PMCd("YY")
  Case Is = UCase("wTzOdChr")
    Call HlpList.vPSCd(ADC("TCoCd"), "CHR", ADC("wTzOdTc"), "Y")
  Case Is = UCase("wTzOdNo")
    Call HlpList.OmNo(ADC("TCoCd"), ADC("wTzOdTc"), ADC("wTzOdYy"), ADC("wTzOdChr"))
  'kbs.132 help for Sr number
  Case Is = UCase("wTzOdSr")
    Call HlpList.OdSr(ADC("TCoCd"), ADC("wTzOdTc"), ADC("wTzOdYy"), ADC("wTzOdChr"), ADC("wTzOdNo"))
  Case Is = UCase("wTzDmCd")
    Call HlpList.DmCd("DM")
  Case Is = UCase("wTzRmCd")
    If GrdTxndZ.Row > 0 Then
      If GrdTxndZ.Value(GrdTxndZ.Row, "TzRmZCd") <> "" Then
        Dim ws_Arr() As String
        ws_Arr = Split(moCn.GetFldVal("Select RmCtg+','+RmSCtg From RmMst where " + IIF(gs_Partition = ctCurrPrtn, " RmPrtKey='" + ctCurrPrtn + "' And ", "") + " RmCd= '" + GrdTxndZ.Value(GrdTxndZ.Row, "TzRmZCd") + "'"), ",")
        Call HlpList.RmCd("N", "'" + ws_Arr(0) + "'", ws_Arr(1), True)
      Else
        Call HlpList.RmCd("N", , , True)
      End If
    Else
      Call HlpList.RmCd("N", , , True)
    End If
  '*** (Jen 3.1.2)
  ' **** Manali 3.6.0 - 19/09/09 - Copy Best Use Size fields Y/N
  Case Is = UCase("wCpBsSzYn")
    Call HlpList.PMCd("YN")
  ' **** Manali 3.6.0 - 19/09/09 - Copy Best Use Size fields Y/N
  
  ' ***** Manali 3.8.0 - TImpCtg
  Case Is = UCase("TImpCtg")
    Call HlpList.PMCd("ImpCtg")
  
  End Select
End Sub
Private Sub atxt_ReCalculate(Index As Integer, ByVal IdName As String, ByVal OldVal As Variant, ByVal RowNum As Integer, ByVal SourceName As String, ByVal SourceColName As String)
  '*** TChr on TTc                                                Normal
      'PDesc225('TC', adc(TTc))
  '*** TCnvRt on TDt                                              Normal
      'DtValue corresponding to max(DtDt) from DtTbl(DtTag= 'USD' and DtDt<= TDt)
      'USD stands for 'US Dollars'.
  '*** wTotQty on TzRmQty                                         adc= Normal, GrdTxndZ= Normal,Delete,InitOver
      'If wTotQty is visible then
      'wTotQty = sum(TzRmQty) for the entire grid
  '*** wTotWt on TzRmWt                                           adc= Normal, GrdTxndZ= Normal,Delete,InitOver
      'If wTotWt is visible then
      'wTotWt = sum(TzRmWt) for the entire grid
  '*** wBalWt on wTzRmBalWt                                       adc= Normal, GrdTxndZ= Normal,Delete,InitOver
      'If wBalWt is visible then
      'wBalWt = sum(wTzRmBalWt) for the entire grid
  '*** TNetAmt on TzPurAmt                                        adc= Normal, GrdTxndZ= Normal,Delete
      'If TNetAmt is visible then
      'TNetAmt = sum(TzPurAmt) for the entire grid
  '*** TTotAmt on TNetAmt,TAddAmt1,TAddAmt2,TAddAmt3,TAddAmt4     Normal
      'If TTotAmt is visible then
      'TTotAmt = TNetAmt + TAddAmt1 + TAddAmt2 + TAddAmt3 + TAddAmt4
  '*** TSuppCd on wCmName                                         Normal
      'CmCd from CustMst(CmCtg= 'S', CmName = adc(wCmName))
  '*** wCmName on TSuppCd                                         Normal, Init
      'wCmName= CmName from CustMst(CmCtg= 'S', CmCd = TSuppCd)
  
  Dim i As Integer
  Dim wCmCd As String, wCustName As String
  Dim wSuppPayTerms As String     ' ***** Manali 3.10.0 - 12/03/12 - Supplier Payment Terms
  Dim wCmCurCd As String
  Select Case UCase(IdName)
  Case Is = UCase("TChr")
    If ADC.Mode = xNorm Then
       ADC("TChr") = moCn.GetFldVal("Select vPValue from vParam where " + _
                                "vPCoCd = '" + ADC("TCoCd") + "' and vPTyp= 'DEFCHR' " + _
                                "and vPMCd= '" + ADC("TTc") + "'")
       '("Select PDesc225 from Param where PTyp='TC' " + _
                                    "and PMCd='" + ADC("TTc") + "'")
    End If
    
  Case Is = UCase("wCpTChr")
    If ADC.Mode = xNorm Then
       '*** Jen (2.12 Patch)
       ADC("wCpTChr") = moCn.GetFldVal("Select vPValue from vParam where " + _
                                "vPCoCd = '" + ADC("wCpTCoCd") + "' and vPTyp= 'DEFCHR' " + _
                                "and vPMCd= '" + ADC("wCpTTc") + "'")
       '*** Jen (2.12 Patch)
       
       '*** Bef 25/03/06
       'ADC("wCpTChr") = moCn.GetFldVal("Select vPValue from vParam where " + _
       '                         "vPCoCd = '" + ADC("TCoCd") + "' and vPTyp= 'DEFCHR' " + _
       '                         "and vPMCd= '" + ADC("wCpTTc") + "'")
       '*** Bef 25/03/06
    End If
  
  '****** Sachin 2.13.0 - [18-07-2006] - Multi-Currency ******
  Case Is = UCase("TCnvRt")
    If ADC.Mode = xNorm Then
       ADC("TCnvRt") = CnvRt(moCn.GetFldVal("Select hBaseCurCd From Head Where HCd='" + ctSelfCmCd + "' And HCoCd='" + gs_CoCd + "'"), ADC("TCurrCd"), ADC("TDt"))
    End If

  Case Is = UCase("TMulDiv")
    If ADC.Mode = xNorm Then
       ADC("TMulDiv") = MulDiv(moCn.GetFldVal("Select hBaseCurCd From Head Where HCd='" + ctSelfCmCd + "' And HCoCd='" + gs_CoCd + "'"), ADC("TCurrCd"))
    End If
  '****** Sachin 2.13.0 - [18-07-2006] - Multi-Currency ******
  ' Zubin 213
  Case Is = UCase("TCnvRtRs")
    If ADC.Mode = xNorm Then
      ADC("TCnvRtRs") = CnvRt(ADC("TCurrCd"), ctIndCurncy, ADC("TDt"))
    End If
  Case Is = UCase("TMulDivRs")
    If ADC.Mode = xNorm Then
      ADC("TMulDivRs") = MulDiv(ADC("TCurrCd"), ctIndCurncy)
    End If
  ' Zubin 213
  Case Is = UCase("wTotQty")
    Dim wTotQty As Single
    If ADC.Mode = xNorm And ADC("wTotQty").Visible Then
      If GrdTxndZ.Mode = fgmnorm Or GrdTxndZ.Mode = fgmdel Or GrdTxndZ.Mode = fgmInitOver Or GrdTxndZ.Mode = fgmCopyOver Then
        wTotQty = 0
        For i = 1 To (GrdTxndZ.Rows - 1)
          If Not GrdTxndZ.IsDel(i) Then wTotQty = wTotQty + GrdTxndZ.Value(i, "TzRmQty")
        Next i
        ADC("wTotQty") = wTotQty
      End If
    End If
  Case Is = UCase("wTotWt")
    Dim wTotWt As Double
    If ADC.Mode = xNorm And ADC("wTotWt").Visible Then
      If GrdTxndZ.Mode = fgmnorm Or GrdTxndZ.Mode = fgmdel Or GrdTxndZ.Mode = fgmInitOver Or GrdTxndZ.Mode = fgmCopyOver Then
        wTotWt = 0
        For i = 1 To (GrdTxndZ.Rows - 1)
          If Not GrdTxndZ.IsDel(i) Then wTotWt = wTotWt + GrdTxndZ.Value(i, "TzRmWt")
        Next i
        ADC("wTotWt") = wTotWt
      End If
    End If
  Case Is = UCase("wBalWt")
    Dim wBalWt As Double
    If ADC.Mode = xNorm And ADC("wBalWt").Visible Then
      If GrdTxndZ.Mode = fgmnorm Or GrdTxndZ.Mode = fgmdel Or GrdTxndZ.Mode = fgmInitOver Or GrdTxndZ.Mode = fgmCopyOver Then
        wBalWt = 0
        For i = 1 To (GrdTxndZ.Rows - 1)
          If Not GrdTxndZ.IsDel(i) Then wBalWt = wBalWt + GrdTxndZ.Value(i, "wTzRmBalWt")
        Next i
        ADC("wBalWt") = wBalWt
      End If
    End If
  Case Is = UCase("TNetAmt")
    Dim wTNetAmt  As Double
    If ADC.Mode = xNorm And ADC("TNetAmt").Visible Then
      If GrdTxndZ.Mode = fgmnorm Or GrdTxndZ.Mode = fgmdel Or GrdTxndZ.Mode = fgmCopyOver Then
        wTNetAmt = 0
        For i = 1 To (GrdTxndZ.Rows - 1)
          If Not GrdTxndZ.IsDel(i) Then wTNetAmt = wTNetAmt + GrdTxndZ.Value(i, "TzPurAmt")
          'Debug.Print CStr(GrdTxndZ.Value(i, "TzPurAmt")) & " => " & CStr(wTNetAmt)
        Next i
        'Debug.Print wTNetAmt
        ADC("TNetAmt") = wTNetAmt
      End If
    End If
  Case Is = UCase("TTotAmt")
    If ADC.Mode = xNorm And ADC("TTotAmt").Visible Then _
        ADC("TTotAmt") = ADC("TNetAmt") + ADC("TAddAmt1") + ADC("TAddAmt2") + ADC("TAddAmt3") + ADC("TAddAmt4")
  Case Is = UCase("TSuppCd")
    ' ***** Manali 3.10.0 - 12/03/12 - Supplier Payment Terms
    If ADC.Mode = xNorm Then
      wCmCd = moCn.GetFldVal("Select CmCd from CustMst where CmCtg='S' and " + _
                             "CmName= '" + ADC("wCmName") + "'")
      wSuppPayTerms = moCn.GetFldVal("Select CmPayTerms from CustMst where CmCtg='S' and " + _
                             "CmName= '" + ADC("wCmName") + "'")
      If wCmCd <> ADC("TSuppCd") Then ADC("TSuppCd") = wCmCd: ADC("TSuppPayTerms") = wSuppPayTerms
    End If
  Case Is = UCase("wCmName")
    If ADC.Mode = xNorm Then
      ' ***** Manali 3.10.0 - 12/03/12 - Supplier Payment Terms
       wCustName = moCn.GetFldVal("Select CmName from CustMst where CmCtg='S' and " + _
                                  "CmCd= '" + ADC("TSuppCd") + "'")
        wSuppPayTerms = moCn.GetFldVal("Select CmPayTerms from CustMst where CmCtg='S' and " + _
                               "CmCd= '" + ADC("TSuppCd") + "'")
       If wCustName <> ADC("wCmName") Then ADC("wCmName") = wCustName: ADC("TSuppPayTerms") = wSuppPayTerms
    ElseIf ADC.Mode = xInit Then
       ADC("wCmName") = moCn.GetFldVal("Select CmName from CustMst where CmCtg='S' and " + _
                                  "CmCd= '" + ADC("TSuppCd") + "'")
    End If
    '******************Geeta*Emr207************************
  Case Is = UCase("TCurrCd")
    If ADC.Mode = xNorm Then
       '****** Sachin 2.13.0 [18-07-2006] - Multi Currency ******
       ADC("TCurrCd") = moCn.GetFldVal("Select CmCurCd from CustMst where CmCtg='S' and " + _
                                           "CmCd= '" + ADC("TSuppCd") + "'")
       '****** Sachin 2.13.0 [18-07-2006] - Multi Currency ******
       'ADC("TUSDYn") = IIF(UCase(wCmCurCd) = ctDefCurncy, "Y", "N")
    End If
   '******************Geeta*Emr207************************
  End Select
End Sub
Private Sub ATXT_FldChange(Index As Integer, ByVal IdName As String, Cancel As Boolean, ErrMsg As String, ByVal pv_NewValue As Variant)
  '*** TTc        PMCd('TC', PValue= ms_TcTyp) and the type of entry should be 'TZ'
  '*** TYy        PMCd('YY')
  '*** TChr       PSCd('CHR', adc(TTc)) or ''
  '*** TDt        Not Blank and >= ms_LockDt
  '*** TSuppCd    CmCd from CustMst(CmCtg= 'S')
  '*** wCmName    CmName from CustMst(CmCtg= 'S')
  
  Select Case UCase(IdName)
  Case Is = UCase("TTc")
    Cancel = Not moCn.RecSeek("select PMCd from Param where PTyp= 'TC' and " + _
                         "PMCd = '" + pv_NewValue + "' and " + _
                         "PValue= '" + ms_TcTyp + "' and " + _
                         "(Select Pvalue1 from Param where PTyp='TCTYP' and " + _
                         "PMCd= '" + ms_TcTyp + "') ='TZ'")
    If Cancel = True Then ErrMsg = "Invalid Voucher Tc": Exit Sub
    Case Is = UCase("TYy")
    Cancel = Not moCn.RecSeek("select PMCd from Param where PTyp= 'YY' and " + _
                         "PMCd = '" + pv_NewValue + "'")
    If Cancel = True Then ErrMsg = "Invalid Voucher Year": Exit Sub
  Case Is = UCase("TChr")
    Cancel = (pv_NewValue <> "") And ((Not moCn.RecSeek("select vPSCd from vParam where " + _
             "vPCoCd='" + ADC("TCoCd") + "' and vPTyp= 'CHR' and " + _
             "vPMCd = '" + ADC("TTc") + "' and vPSCd='" + pv_NewValue + "'")))
    If Cancel = True Then ErrMsg = "Invalid Voucher Character": Exit Sub
' **** Zubin 211 **** '
    Dim ws_UsrArr() As String, wi_Index As Integer
    ws_UsrArr = Split(moCn.GetFldVal("Select vPValue4 from vParam where vPCoCd= '" + gs_CoCd + "' and " + _
                                        "vPTyp='CHR' and vPMCd='" + ADC("TTc") + "' and vPSCd= '" + pv_NewValue + "'"), ",")
    If IsArray(ws_UsrArr) Then
      For wi_Index = 0 To UBound(ws_UsrArr)
        Cancel = IIF(UCase(gs_UsrCd) = UCase(ws_UsrArr(wi_Index)), False, True)
        If Cancel = False Then: Exit For
      Next wi_Index
    End If
   If Cancel = True Then ErrMsg = "Access To This Voucher Character Is Denied To User " + gs_UsrCd: Exit Sub
' **** Zubin 211 **** '
    
  Case Is = UCase("TDt")
    If pv_NewValue = "" Then Cancel = True: ErrMsg = "Purchase Date Cannot Be Blank": Exit Sub
    If pv_NewValue < CDate(ms_LockDt) Then Cancel = True: ErrMsg = "Voucher Date Cannot Be Less Than Lock Date": Exit Sub
    If pv_NewValue > CDate(ms_FutureDt) Then Cancel = True: ErrMsg = "Voucher Date Cannot Be Greater Than Future Date": Exit Sub
    'CC.999.13 check for purge date added
    If pv_NewValue <= CDate(ms_PurgeDt) Then Cancel = True: ErrMsg = "Voucher Date Cannot Be Less Than Or Equal To Purge Date": Exit Sub
        
    If moCn.GetFldVal("Select hHolChkYN From Head where HCd='ZSELF' and HCoCd='" + gs_CoCd + "'") = "Y" Then
        '53 company code added
        If IsHOLYN(ADC("TCoCd"), pv_NewValue) Then Cancel = True: ErrMsg = "Voucher Cannot be Added on Holidays": Exit Sub
    End If
  Case Is = UCase("TSuppCd")
    Cancel = Not moCn.RecSeek("Select CmCd From CustMst Where " + _
                         "CmCtg = 'S' and CmCd = '" + pv_NewValue + "'")
    If Cancel = True Then ErrMsg = "Invalid Supplier Code": Exit Sub
' ########################  Manoj #### Ver: 2.0.6 #### Date: 25/12/2003  ########################
  Case Is = UCase("TLockYN")
    Cancel = (pv_NewValue <> "") And ((Not moCn.RecSeek("select PMCd from Param where " + _
            "PTyp= 'YN' and PMCd = '" + pv_NewValue + "'")))
    If Cancel = True Then ErrMsg = "Invalid Custom Purchase Lock (Y/N)": Exit Sub
    '******************Geeta********Emr206******02/04/04****
    If pv_NewValue = "Y" Then
      'ms_PwdFlag = False '*** Jay 2.14 *** (SEO Change)    '*** (Bef 3.2.0)
      ADC.AllowDelete = False: GrdTxndZ.AllowDelete = False: GrdTxndZ.AllowAdd = False
    Else
      'TxtPwd.text = "": ms_PwdFlag = True '*** Jay 2.14 *** (SEO Change)   '*** (Bef 3.2.0)
      ADC.AllowDelete = True: GrdTxndZ.AllowDelete = True: GrdTxndZ.AllowAdd = True
    End If
    '****************************Geeta****************************
' ###########################################  Manoj  ###########################################
  Case Is = UCase("wCmName")
    Cancel = Not moCn.RecSeek("select CmName from CustMst where CmCtg= 'S' and CmName = '" + pv_NewValue + "'")
    If Cancel = True Then ErrMsg = "Invalid Supplier Name": Exit Sub
  Case Is = UCase("wCpTTc")
    Cancel = (pv_NewValue <> "") And (Not moCn.RecSeek("select PMCd from Param where PTyp= 'TC' and " + _
                              "PMCd = '" + pv_NewValue + "' and PValue= 'PR'"))
    If Cancel = True Then ErrMsg = "Invalid Purchase Voucher Tc": Exit Sub
  Case Is = UCase("wCpTYy")
    Cancel = (pv_NewValue <> "") And (Not moCn.RecSeek("select PMCd from Param where " + _
             "PTyp= 'YY' and PMCd = '" + pv_NewValue + "'"))
    If Cancel = True Then ErrMsg = "Invalid Purchase Voucher Year": Exit Sub
  Case Is = UCase("wCpTChr")
    '*** Jen (2.12 Patch)
    Cancel = (pv_NewValue <> "") And (Not moCn.RecSeek("select vPSCd from vParam where " + _
             "vPCoCd= '" + ADC("wCpTCoCd") + "' and vPTyp= 'CHR' and " + _
             "vPMCd= '" + ADC("wCpTTc") + "' and vPSCd= '" + pv_NewValue + "'"))
    '*** Jen (2.12 Patch)
    
    '*** (Bef 25/03/06)
    'Cancel = (pv_NewValue <> "") And (Not moCn.RecSeek("select vPSCd from vParam where " + _
    '         "vPCoCd= '" + ADC("TCoCd") + "' and vPTyp= 'CHR' and " + _
    '         "vPMCd= '" + ADC("wCpTTc") + "' and vPSCd= '" + pv_NewValue + "'"))
    '*** (Bef 25/03/06)
    
    If Cancel = True Then ErrMsg = "Invalid Purchase Voucher Character": Exit Sub
  Case Is = UCase("wCpGrp")
    Cancel = (pv_NewValue <> "") And (Not moCn.RecSeek("select PMCd from Param where " + _
             "PTyp= 'PURCPY' and PMCd = '" + pv_NewValue + "'"))
    If Cancel = True Then ErrMsg = "Invalid Purchase Group Option": Exit Sub
  
  '***************** Urmi 28 Jun 03 ***************************************************'
  Case Is = UCase("TImpTyp")
    Cancel = Not moCn.RecSeek("select PMCd from Param where " + _
             "PTyp= 'ImpTyp' and PMCd = '" + pv_NewValue + "'")
    If Cancel = True Then ErrMsg = "Invalid Import Type": Exit Sub
  '***************** Urmi 28 Jun 03 *************************************************'
    
  '*** (Jen 3.1.1)
  Case Is = UCase("TCmCd")
    Cancel = (pv_NewValue <> "") And (Not moCn.RecSeek("Select CmCd From CustMst Where CmCtg = 'C' and CmCd = '" + pv_NewValue + "'"))
    If Cancel = True Then ErrMsg = "Invalid Customer Code": Exit Sub
  '*** (Jen 3.1.1)
    
  '*** (Jen 3.1.2)
  Case Is = UCase("wTzOdTc")
    Cancel = pv_NewValue <> "" And (Not moCn.RecSeek("select PMCd from Param where PTyp= 'TC' and " + _
                         "PMCd = '" + pv_NewValue + "' and PValue= 'SO' "))
    If Cancel = True Then ErrMsg = "Invalid Order Tc": Exit Sub
  Case Is = UCase("wTzOdYy")
    Cancel = pv_NewValue <> "" And (Not moCn.RecSeek("select PMCd from Param where PTyp= 'YY' and " + _
                         "PMCd = '" + pv_NewValue + "'"))
    If Cancel = True Then ErrMsg = "Invalid Year": Exit Sub
  Case Is = UCase("wTzOdChr")
    Cancel = pv_NewValue <> "" And (Not moCn.RecSeek("select vPSCd from vParam where " + _
                         "vPCoCd= '" + ADC("TCoCd") + "' and vPTyp= 'CHR' and " + _
                         "vPMCd= '" + ADC("wTzOdTc") + "' and vPSCd= '" + pv_NewValue + "'"))
    If Cancel = True Then ErrMsg = "Invalid Order Character": Exit Sub
  Case Is = UCase("wTzOdNo")
    Cancel = pv_NewValue <> 0 And (Not moCn.RecSeek("select OmNo from OrdMst where OmPrtKey='" + ctCurrPrtn + "' and " + _
                         "OmCoCd= '" + ADC("TCoCd") + "' and OmTc= '" + ADC("wTzOdTc") + "' " + _
                         "and OmYy= '" + ADC("wTzOdYy") + "' and OmChr= '" + ADC("wTzOdChr") + "' " + _
                         "and OmNo='" + pv_NewValue + "'"))
    If Cancel = True Then ErrMsg = "Invalid Order Number": Exit Sub
  'kbs.132 validation for Sr number
  Case Is = UCase("wTzOdSr")
    Cancel = pv_NewValue <> 0 And (Not moCn.RecSeek("select OdSr from OrdDsg where OdPrtKey='" + ctCurrPrtn + "' and " + _
                         "OdCoCd= '" + ADC("TCoCd") + "' and OdTc= '" + ADC("wTzOdTc") + "' " + _
                         "and OdYy= '" + ADC("wTzOdYy") + "' and OdChr= '" + ADC("wTzOdChr") + "' " + _
                         "and OdNo=" + CStr(ADC("wTzOdNo")) + " and OdSr=" + CStr(pv_NewValue)))
    If Cancel = True Then ErrMsg = "Invalid Order Serial Number": Exit Sub
  Case Is = UCase("wTzDmCd")
    Cancel = pv_NewValue <> "" And (Not moCn.RecSeek("Select DmCd From DsgMst Where DmPrtKey='" + ctCurrPrtn + "' and DmTcTyp= 'DM' " + _
                              "and DmCd= '" + pv_NewValue + "' and DmSz= '' /* and DmHld= 'N' */ "))
    If Cancel = True Then ErrMsg = "Invalid Design Code": Exit Sub
  Case Is = UCase("wTzRmCd")
    Cancel = pv_NewValue <> "" And (Not moCn.RecSeek("Select RmCd From RmMst Where RmPrtKey='" + ctCurrPrtn + "' and RmCd = '" + pv_NewValue + "' " + _
             "and RmZ= 'N' "))
    If Cancel = True Then ErrMsg = "Invalid Raw Material Code": Exit Sub
  '*** (Jen 3.1.2)
  ' **** Manali 3.6.0 - 19/09/09 - Copy Best Use Size fields Y/N
  Case Is = UCase("wCpBsSzYn")
    Cancel = Not moCn.RecSeek("select PMCd from Param where PTyp= 'YN' and PMCd = '" + pv_NewValue + "'")
    If Cancel = True Then ErrMsg = "Enter 'Y'es or 'N'o ": Exit Sub
  ' **** Manali 3.6.0 - 19/09/09 - Copy Best Use Size fields Y/N
  
  ' *****  Manali 3.8.0 - IMPCTG
  Case Is = UCase("TImpCtg")
    Cancel = Not moCn.RecSeek("select PMCd from Param where " + _
             "PTyp= 'ImpCtg' and PMCd = '" + pv_NewValue + "'")
    If Cancel = True Then ErrMsg = "Invalid Import Category": Exit Sub
  
  End Select
End Sub
Private Sub CmdCDGo_Click()
    '4.3.0.0
    If ADC("wTDt") <> "01/01/80" And ADC("wTDt") <> ADC("TDt") Then
        ADC("TDt") = ADC("wTDt")
        ADC("TBoeDt") = ADC("wTDt")
    End If
    Call DispFra(ChgDt)
End Sub

Private Sub CmdChgDt_Click()
    If ADC("TLockYN") = "N" Then
        Call DispFra(ChgDt)       '4.3.0.0
    Else
        DispMsg "Cannot Change Date as Voucher is Locked", etError
    End If
End Sub

Private Sub CmdInvHd_Click()
'pg.11 calling invoice header form
  If Not ADC.SaveRec Then
    DispMsg "X Save Failed", etInfo
    Exit Sub
  End If
  If moCn.RecSeek("Select 'x' From UsrAccess where UaUsrCd='" + ADC.UsrCd + "'  and UaMnuCd='In'") Then
      If gs_InNo = 0 Then
       gs_InCoCd = ADC("TCoCd")
       gs_InTc = ADC("TTc")
       gs_InYy = ADC("TYy")
       gs_InChr = ADC("TChr")
       gs_InNo = ADC("TNo")
       FormSelect ("In")
      Else
       DispMsg "Invoice Header already open.", etError
      End If
  Else
      DispMsg "User has no access to Invoice Header menu.", etError
  End If
End Sub

Private Sub GrdTxndZ_RowWhen(ByVal RowNum As Integer)
  '*** If GrdTxndA.IsNew Then
      'TzSr = Max of the Grid + 1
      'TzRmDespWt = 0, TzCls= 'N'
  
  With GrdTxndZ
    If .IsNew(RowNum) Then
      .Value(RowNum, "TzSr") = .MaxVal("TzSr") + 1
      .Value(RowNum, "TzRmDespWt") = 0
      .Value(RowNum, "TzCls") = "N"
      .Value(RowNum, "TzPrtKey") = ctCurrPrtn   '****** Sachin 3.02 26-12-07
      .Value(RowNum, "TzClsDt") = MWLib.EmptyDate     ' ***** Manali 350Nxt - Close date
    End If
    ' ****** Sachin 2.12 - Desc in Custom Purchase
    ADC("wTzDesc") = .Value(RowNum, "TzDesc")
    ' ****** Sachin 2.12 - Desc in Custom Purchase
    
    '*** (Jen 3.1.2)
    ADC("wTzOdTc") = .Value(RowNum, "TzOdTc")
    ADC("wTzOdYy") = .Value(RowNum, "TzOdYy")
    ADC("wTzOdChr") = .Value(RowNum, "TzOdChr")
    ADC("wTzOdNo") = .Value(RowNum, "TzOdNo")
    'kbs.132 copying values from grid
    ADC("wTzOdSr") = .Value(RowNum, "TzOdSr")
    ADC("wTzDmCd") = .Value(RowNum, "TzDmCd")
    ADC("wTzRmCd") = .Value(RowNum, "TzRmCd")
    ADC("wTzRmSzFr") = .Value(RowNum, "TzRmSzFr")
    ADC("wTzRmSzTo") = .Value(RowNum, "TzRmSzTo")
    '*** (Jen 3.1.2)
    
  End With
End Sub
Private Sub GrdTxndZ_ColWhen(ByVal RowNum As Integer, ByVal ColName As String, Cancel As Boolean, ErrMsg As String, ByVal Action As MwfCtl.en_fgAction)
  '*** Set Helps For Each Field
  '*** Cannot edit any field other than TzCls if Despatch Wt > 0
  '*** Cannot Edit TzCls when TzRmWt = TzRmDespWt
  
  If GrdTxndZ.Value(RowNum, "TzPrtKey") <> ctCurrPrtn Or mf_DtLocked = True Then Cancel = True: ErrMsg = "Cannot Edit Locked Entry": Exit Sub
' ########################  Manoj #### Ver: 2.0.6 #### Date: 25/12/2003  ########################
  If UCase$(ADC("TLockYN")) = UCase$("Y") Then Cancel = True: ErrMsg = "Cannot Edit Locked Entry": Exit Sub
' ###########################################  Manoj  ###########################################
  
  With GrdTxndZ
    '*** Shilpa ***
    'If ColName <> UCase("TzCls") And ColName <> UCase("TzRmWt") And .Value(RowNum, "TzRmDespWt") > 0 Then Cancel = True: ErrMsg = "Cannot Edit This Value": Exit Sub
    '*** Shilpa ***
    ' ***** Manali 3.6.2 - Added TzClsDt
    If ColName <> UCase("TzCls") And ColName <> UCase("TzRmQty") And ColName <> UCase("TzRmWt") And .Value(RowNum, "TzRmDespWt") > 0 And ColName <> UCase("TzClsDt") Then Cancel = True: ErrMsg = "Cannot Edit This Value": Exit Sub
    If ColName <> UCase("TzCls") And ColName <> UCase("TzClsDt") And .Value(RowNum, "TzCls") = "Y" Then Cancel = True: ErrMsg = "Cannot Edit This Value": Exit Sub
'    If ColName <> UCase("TzCls") And ColName <> UCase("TzRmQty") And ColName <> UCase("TzRmWt") And .Value(RowNum, "TzRmDespWt") > 0 Then Cancel = True: ErrMsg = "Cannot Edit This Value": Exit Sub
'    If ColName <> UCase("TzCls") And .Value(RowNum, "TzCls") = "Y" Then Cancel = True: ErrMsg = "Cannot Edit This Value": Exit Sub
    
    Select Case ColName
    Case Is = UCase("TzRmZCd")
      Call HlpList.RmCd("Y")
    Case Is = UCase("TzRmCdDesc")
      Call HlpList.RmCd("N")
    Case Is = UCase("TzCls")
      If .Value(RowNum, "TzRmWt") = .Value(RowNum, "TzRmDespWt") Then Cancel = True: ErrMsg = "Cannot Edit This Value": Exit Sub
      Call HlpList.PMCd("YN")
    ' ***** Manali 3.6.2 - error Rectified
    Case Is = UCase("TzClsDt")
      If .Value(RowNum, "TzRmWt") = .Value(RowNum, "TzRmDespWt") Or .Value(RowNum, "TzCls") = "N" Then Cancel = True: ErrMsg = "Cannot Edit This Value": Exit Sub
      
    Case Is = UCase("TzRmWt")
      If GrdTxndZ.IsNew(RowNum) = True Then
        Cancel = False
      Else
        'If ms_PwdFlag = True And ADC.AddFndMode = xFndMode Then
        If ms_PwdFlag = True Then
          Cancel = False
        Else
          Cancel = True: ErrMsg = "Cannot Edit This Value, Enter SEO Password": Exit Sub
        End If
      End If
    Case Is = UCase("TzRmQty")
      If GrdTxndZ.IsNew(RowNum) = True Then
        Cancel = False
      Else
        'If ms_PwdFlag = True And ADC.AddFndMode = xFndMode Then
        If ms_PwdFlag = True Then
          Cancel = False
        Else
          Cancel = True: ErrMsg = "Cannot Edit This Value, Enter SEO Password": Exit Sub
        End If
      End If
    End Select
  End With
End Sub
Private Sub GrdTxndZ_ReCalc(ByVal RowNum As Integer, ByVal ColName As String, ByVal OldVal As Variant, ByVal SourceName As String, ByVal SourceColName As String)
  '*** TzRmQty on TzPcPerCt,TzRmWt                  Normal
      'If TzPcPerCt <> Round(TzRmQty/ TzRmWt) Then
      '   TzRmQty = Round(TzPcPerCt * TzRmWt)
  '*** TzPcPerCt on TzRmZCd,TzRmQty,TzRmWt          Normal
      'If SourceColName = TzRmZCd Then
      '   If RmCtg of TzRmZCd <> Old RmCtg or RmSCtg of TzRmZCd <> Old RmSCtg Then
      '       TzPcPerCt = 0
      'Else
      '   TzPcPerCt = Round(TzRmQty/ TzRmWt)
  '*** TzRmWt on TzRmZCd                            Normal
      'If RmCtg of TzRmZCd <> Old RmCtg or RmSCtg of TzRmZCd <> Old RmSCtg Then
      '   TzRmWt = 0
  '*** wTzRmBalWt on TzRmWt,TzRmDespWt              Normal,Init
      '   wTzRmBalWt = TzRmWt - TzRmDespWt
  '*** TzPurRt on TzRmZCd                           Normal
      '   TzPurRt = 0
  '*** TzPurAmt on TzPurRt,TzRmWt                   Normal
      '   TzPurAmt = TzPurRt * TzRmWt
  With GrdTxndZ
    Select Case ColName
    Case Is = UCase("TzRmQty")              '*** recalc on
      If .Mode = fgmnorm Then
        If .Value(RowNum, "TzPcPerCt") <> Round(MWLib.Div(.Value(RowNum, "TzRmQty"), _
           .Value(RowNum, "TzRmWt")), 0) And .Value(RowNum, "TzRmWt") <> 0 Then _
           .Value(RowNum, "TzRmQty") = Round(.Value(RowNum, "TzPcPerCt") * .Value(RowNum, "TzRmWt"), 0)
      End If
    Case Is = UCase("TzPcPerCt")
      If .Mode = fgmnorm Then
        If UCase(SourceColName) = UCase("TzRmZCd") Then
          If FnRmCtg(.Value(RowNum, "TzRmZCd")) <> FnRmCtg(OldVal) Or _
             FnRmSCtg(.Value(RowNum, "TzRmZCd")) <> FnRmSCtg(OldVal) Then _
             .Value(RowNum, "TzPcPerCt") = 0
        Else
          If .Value(RowNum, "TzRmQty") <> 0 And .Value(RowNum, "TzRmWt") <> 0 Then .Value(RowNum, "TzPcPerCt") = Round(MWLib.Div(.Value(RowNum, "TzRmQty"), .Value(RowNum, "TzRmWt")), 0)
        End If
      End If
    Case Is = UCase("TzRmWt")
      If .Mode = fgmnorm Then
        If FnRmCtg(.Value(RowNum, "TzRmZCd")) <> FnRmCtg(OldVal) Or _
           FnRmSCtg(.Value(RowNum, "TzRmZCd")) <> FnRmSCtg(OldVal) Then _
           .Value(RowNum, "TzRmWt") = 0
      End If
    Case Is = UCase("wTzRmBalWt")
      If .Mode = fgmnorm Or .Mode = fgminit Or .Mode = fgmCopy Then _
         .Value(RowNum, "wTzRmBalWt") = .Value(RowNum, "TzRmWt") - .Value(RowNum, "TzRmDespWt")
    Case Is = UCase("TzPurRt")
        '*** (Bef 3.1.1)
        'If .Mode = fgmnorm Then .Value(RowNum, "TzPurRt") = 0
        '*** (Bef 3.1.1)
        
        '*** (Jen 3.1.1)
        If .Mode = fgmnorm And UCase(SourceColName) = UCase("TzRmZCd") Then .Value(RowNum, "TzPurRt") = 0
        If .Mode = fgmnorm And UCase(SourceColName) = UCase("TzPurAmt") Then
          If .Value(RowNum, "TzPurRt") = 0 And .Value(RowNum, "TzPurAmt") > 0 And .Value(RowNum, "TzRmWt") > 0 Then _
             .Value(RowNum, "TzPurRt") = .Value(RowNum, "TzPurAmt") / .Value(RowNum, "TzRmWt")
        End If
        '*** (Jen 3.1.1)
    
    Case Is = UCase("TzPurAmt")
      If .Mode = fgmnorm Then .Value(RowNum, "TzPurAmt") = _
                             .Value(RowNum, "TzPurRt") * .Value(RowNum, "TzRmWt")
    Case Is = UCase("TzCls")
      If .Mode = fgmnorm Then
        If .Value(RowNum, "TzRmWt") > 0 And .Value(RowNum, "TzRmWt") = .Value(RowNum, "TzRmDespWt") Then
          .Value(RowNum, "TzCls") = "Y"
        Else
          .Value(RowNum, "TzCls") = "N"
        End If
      End If
    
      
    ' ****** Sachin 2.12 - Desc in Custom Purchase
    Case Is = UCase("TzDesc")
      If .Mode = fgmnorm Then
        If .Row > 0 Then
          If .Value(.Row, "TzDesc") <> ADC("wTzDesc") Then
            .Value(.Row, "TzDesc") = ADC("wTzDesc")
            .SaveRec .Row
          End If
        End If
     End If
    ' ****** Sachin 2.12 - Desc in Custom Purchase
    
    Case Is = UCase("TzClsDt")
      If .Mode = fgmnorm Then
        If .Value(RowNum, "TzCls") = "Y" And .Value(RowNum, "TzRmWt") <> .Value(RowNum, "TzRmDespWt") Then
          .Value(RowNum, "TzClsDt") = moCn.SrvrDate
        Else
          .Value(RowNum, "TzClsDt") = MWLib.EmptyDate
        End If
      End If
    End Select
    
  End With
End Sub
Private Sub GrdTxndZ_ColChange(ByVal RowNum As Integer, ByVal ColName As String, ByVal NewValue As Variant, Cancel As Boolean, ErrMsg As String)
  '*** TzRmZCd          valid RmCd from RmMst(RmZ= 'Y') (i.e. a valid custom code)
  '*** TzCls            valid PMCd('YN')
  
  With GrdTxndZ
    Select Case ColName
    Case Is = UCase("TzRmZCd")
      '****** Sachin 3.02 26-12-07 - Current Partition
      Cancel = (Not moCn.RecSeek("Select RmCd From RmMst where RmZ='Y' and " + _
                           "RmCd ='" + NewValue + "' and RmPrtKey='" + ctCurrPrtn + "'")) And NewValue <> ""
      ErrMsg = "Invalid Raw Material Customs Code": Exit Sub
      
    Case Is = UCase("TzCls")
      '****** Sachin 3.02 26-12-07 - Current Partition
      Cancel = Not moCn.RecSeek("Select PMCd From Param where PTyp= 'YN' and " + _
                                "PMCd ='" + NewValue + "'")
      ErrMsg = "Invalid Yes/ No Option": Exit Sub
    End Select
  End With
End Sub
Private Sub GrdTxndZ_InitKey(ByVal RowNum As Integer)
  '*** Initialise Key Fields and all the fields which are not included in the grid
  Dim wTzKey As Single
  With GrdTxndZ
    If .IsNew(RowNum) Then
      .Store "TzCoCd", ADC("TCoCd")
      .Store "TzTc", ADC("TTc")
      .Store "TzYy", ADC("TYy")
      .Store "TzChr", ADC("TChr")
      .Store "TzNo", ADC("TNo")
      .Store "TzSrNo", 0
      .Store "TzRmDespPurgeWt", 0

       wTzKey = moCn.AutoGen("TxndZ", "TzKey", ADC("TCoCd") + ADC("TYy"))
      .Store "TzKey", wTzKey
      .Store "TzRefYy", ADC("TYy")
      .Store "TzRefKey", wTzKey
      .Store "TzBM", "B"
      '.Store "TzCls", "N"
      
      '****** Sachin 3.02 - Partition Keys
      .Store "TzTIdNo", moCn.GetFldVal("Select TIdNo From Txn " + _
                                       " Where TPrtKey='" + ctCurrPrtn + "' and TCoCd='" + ADC("TCoCd") + "' And TTc='" + ADC("TTc") + "' And TYy='" + ADC("TYy") + "'" + _
                                       " And TChr='" + ADC("TChr") + "' And TNo=" + CStr(ADC("TNo")))
      .Store "TzInIdNo", 0
      .Store "TzPrtDespQty", 0
      '****** Sachin 3.02 - Partition Keys
    End If
    
        .Store "TzRmCtg", moCn.GetFldVal("Select RmCtg From RmMst Where RmCd='" + .Value(RowNum, "TzRmZCd") + "'" + IIF(gs_Partition = ctCurrPrtn, " and RmPrtKey='" + ctCurrPrtn + "'", ""))
        .Store "TzRmSCtg", moCn.GetFldVal("Select RmSCtg From RmMst Where RmCd='" + .Value(RowNum, "TzRmZCd") + "'" + IIF(gs_Partition = ctCurrPrtn, " and RmPrtKey='" + ctCurrPrtn + "'", ""))
        '****** Sachin 3.02 - Partition Keys

  End With
End Sub
Private Sub ADC_SetRecSource()
  '*** Set the RecSource for ADC (Table Txn)
  '****** Sachin 3.02 26-12-07 - Corresponding Partition
  'sv.12 purge date check addded
  ADC.RecSource = " Select * from Txn " + _
                  " Where TCoCd = '" + ADC("TCoCd") + "'" + _
                  " And TTc ='" + ADC("TTc") + "'" + _
                  " And TYy = '" + ADC("TYy") + "'" + _
                  " And TChr ='" + ADC("TChr") + "' " + _
                  " And TNo =" + CStr(ADC("TNo")) + " And TPrtKey=(Case When '" + gs_Partition + "'='" + ctCurrPrtn + "' Then '" + ctCurrPrtn + "' Else TPrtKey End)" + _
                  IIF(gs_Partition = ctCurrPrtn, " And Tdt>'" + ms_PurgeDt + "' ", "")
                  
End Sub
Private Sub GrdTxndZ_SetRecSource()
  '*** Set the RecSource for GrdTxndZ (Table TxndZ where TzSrNo = 0)
  '****** Will go on the Full Database
  GrdTxndZ.RecSource = " Select * from TxndZ " + _
                       " Where TzCoCd = '" + ADC("TCoCd") + "'" + _
                       " And TzTc ='" + ADC("TTc") + "'" + _
                       " And TzYy = '" + ADC("TYy") + "'" + _
                       " And TzChr ='" + ADC("TChr") + "' " + _
                       " And TzNo =" + CStr(ADC("TNo")) + _
                       " And TzSrNo = 0 " + _
                       " Order By TzSr"
End Sub
Private Sub adc_AutoGen()
  '*** write code for AutoGenerating the Transaction Voucher Number (PreKey= TTc, TYy, TChr)
      'If the TNo is Zero then generate the number
      'If the user himself enters the number then write this number in the autogen table
      '(if it is greater than the number in the autogen table for the given PreKey)

  If ADC("TNo") = 0 Then
    ADC("TNo") = moCn.AutoGen("Txn", "TNo", ADC("TCoCd") + ADC("TTc") + ADC("TYy") + ADC("TChr"))
  Else
    Call moCn.AutoGen("Txn", "TNo", ADC("TCoCd") + ADC("TTc") + ADC("TYy") + ADC("TChr"), ADC("TNo"))
  End If
End Sub
Private Sub CmdCP_Click()
  '*** The 'X' record (Txn record) has to be saved before the 'A' records can be entered
  '*** Display or hide the Bag Generation frame depending on the toggle state
  If Not ADC.SaveRec Then
    DispMsg "X Save Failed", etInfo
    Exit Sub
  End If
  
  Call DispFra(Cpy)
  
  '*** 27/02/06 For Kama (Jen 2.12)
  If ADC("wCpTCoCd") = "" Then ADC("wCpTCoCd") = ADC("TCoCd")
  '*** 27/02/06 For Kama (Jen 2.12)
  If ADC("wCpGrp") = "" Then ADC("wCpGrp") = "OPT3"
End Sub
Private Sub CmdCpGo_Click()
  '*** Call the BagGeneration procedure to automatically generate the bags for the
      'Order Sr range specified by the user
  '*** Close the Bag Generation Frame
  
  Call PurCopy
End Sub

'*** (Jen 3.1.2)
Private Sub CmdGrdTxndzDets_Click()
  
  If GrdTxndZ.Row = 0 Then Exit Sub
  Call DispFra(DET)
  With GrdTxndZ
    If FraGrdTxndZDetsDisp.Enabled = False Then
      .Value(.Row, "TzDesc") = ADC("wTzDesc")
      .Value(.Row, "TzOdTc") = ADC("wTzOdTc")
      .Value(.Row, "TzOdYy") = ADC("wTzOdYy")
      .Value(.Row, "TzOdChr") = ADC("wTzOdChr")
      .Value(.Row, "TzOdNo") = ADC("wTzOdNo")
      'kbs.132 grid getting values from sr. number text box
      .Value(.Row, "TzOdSr") = ADC("wTzOdSr")
      .Value(.Row, "TzDmCd") = ADC("wTzDmCd")
      .Value(.Row, "TzRmCd") = ADC("wTzRmCd")
      .Value(.Row, "TzRmSzFr") = ADC("wTzRmSzFr")
      .Value(.Row, "TzRmSzTo") = ADC("wTzRmSzTo")
      '.SaveRec .Row
    End If
  End With
End Sub

Private Sub FraCP_Click()
  '*** Jen (2.12 Patch)
  If ADC("wCpTCoCd") = "" Then ADC("wCpTCoCd") = ADC("TCoCd")
  '*** Jen (2.12 Patch)
End Sub



' ########################  Manoj #### Ver: 2.0.6 #### Date: 25/12/2003  ########################
Private Sub TxtPwd_GotFocus()
  DispMsg "Enter SEO Password", etInfo
End Sub
'Private Sub CmdPwd_Click()
'  Call DispFra(Pwd)
'  'If ADC.AddFndMode = xFndMode Then
'        If Trim(TxtPwd.Text) = moCn.GetFldVal("Select PValue1 from Param where " + _
'                                       "Ptyp = 'SEOPWD' and PmCd = 'SEOPWD'") Then
'            ms_PwdFlag = True
'        Else
'            ms_PwdFlag = False
'        End If
'   ' End If
'End Sub
' ###########################################  Manoj  ###########################################

Private Sub TxtPwd_LostFocus()
    Call TxtPwd_Validate(False)
End Sub
'*** Jay 2.14 *** (SEO Change)
Private Sub TxtPwd_Validate(Cancel As Boolean)
  '*** (Jen 3.2.0)
  If Trim(TxtPwd.text) <> "" Then
  '*** (Jen 3.2.0)
  'If ADC.AddFndMode = xFndMode And Trim(TxtPwd.text) <> "" Then    '*** (Bef 3.2.0)
    If Trim(TxtPwd.text) = moCn.GetFldVal("select IsNull ((select PValue from Param Join UsrAccess on UaUsrCd = PMCd " + _
                   "where PTyp= 'USR' and PMCd= '" + gs_UsrCd + "' and UaMnuCd = '" + ADC.MenuCd + "' and " + _
                   "UaSeoYn= 'Y' and PValue= '" + Trim(TxtPwd.text) + "'), " + _
                   "(Select PValue1 from Param where Ptyp= 'SEOPWD' and PmCd= 'SEOPWD' )) ") Then
       'If UCase$(ADC("TLockYN")) = UCase$("N") Then ms_PwdFlag = True Else ms_PwdFlag = False    '*** (Bef 3.2.0)
       ms_PwdFlag = True    '*** (Jen 3.2.0)
       Exit Sub
    Else
       ms_PwdFlag = False: Cancel = True: DispMsg "Invalid Edit Password / User Does Not Have SEO Rights ", etError
    End If
  Else
    ms_PwdFlag = moCn.RecSeek("Select PValue1 from Param where PTyp= 'SEOPWD' and PmCd= 'SEOPWD' and PValue1= '' ")
  End If
End Sub
'' ########################  Manoj #### Ver: 2.0.6 #### Date: 25/12/2003  ########################
'Private Sub TxtPwd_Validate(Cancel As Boolean)
'  ms_PwdFlag = False
'  If Trim$(TxtPwd.text) = Trim$(ms_Pwd) Then
'    ms_PwdFlag = True
'    ADC("TSeoPwd") = TxtPwd.text
'  Else
'    If Trim(TxtPwd.text) <> "" Then
'      Cancel = True
'      DispMsg "Invalid SEO Password", etError
'    End If
'  End If
'End Sub
''Private Sub TxtPwd_Validate(Cancel As Boolean)
''    'If ADC.AddFndMode = xFndMode And Trim(TxtPwd.Text) <> "" Then
''    If Trim(TxtPwd.Text) <> "" Then
''        If Trim(TxtPwd.Text) = moCn.GetFldVal("Select PValue1 from Param where " + _
''                               "Ptyp = 'SEOPWD' and PmCd = 'SEOPWD'") Then
''
''            Exit Sub
''        Else
''            Cancel = True
''            DispMsg "Invalid SEO Password", etError
''        End If
''    End If
''End Sub
'' ###########################################  Manoj  ###########################################
'*** Jay 2.14 *** (SEO Change)

Private Function ColSty(ByVal mColSty As String) As en_fgColStyle
  '*** Function to return the Style of columns in the Grid7
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
Private Sub PropSet()
  '*** Setting Vis and Ena Properties for X Fields ***
  Dim wRsX As MwfLib.MDORowSet
  Set wRsX = moCn.OpenRes("Select * from TcStr where TsTbl= 'TZ' and " + _
                           "TsTyp= '" + ms_TcTyp + "' and TsSTyp= 'X'")
  Do While Not (wRsX.EOF Or wRsX.BOF)
    ADC(wRsX!TsIdName).Enabled = FldEna(wRsX!TsStyle)
    ADC(wRsX!TsIdName).Visible = FldVis(wRsX!TsStyle)
    wRsX.MoveNext
  Loop
  
  '*** Setting Vis and Ena Properties for Label Conrols ***
  Dim wRsLbl As MwfLib.MDORowSet
  Set wRsLbl = moCn.OpenRes("Select * from TcStr where TsTbl= 'TZ' and " + _
                           "TsTyp= '" + ms_TcTyp + "' and TsSTyp= 'Lbl'")
  Dim wLblCtrl As Control
  Do While Not (wRsLbl.EOF Or wRsLbl.BOF)
    For Each wLblCtrl In Controls
      If TypeOf wLblCtrl Is Label Then
        If UCase(wLblCtrl.Name) = UCase(wRsLbl!TsIdName) Then
          wLblCtrl.Visible = FldVis(wRsLbl!TsStyle)
          If wRsLbl!TsLbl <> "" Then wLblCtrl.Caption = wRsLbl!TsLbl
        End If
      End If
    Next
    wRsLbl.MoveNext
  Loop
  
  Set wRsX = Nothing
  Set wRsLbl = Nothing
End Sub
Private Function FnRmCtg(ByVal mRmCd As String) As String
  '*** Function to get the RmCtg for the specified RmCode form RmMst table
  FnRmCtg = Trim(moCn.GetFldVal("Select RmCtg from RmMst where RmPrtKey='" + ctCurrPrtn + "' and RmCd='" + mRmCd + "'"))
End Function
Private Function FnRmSCtg(ByVal mRmCd As String) As String
  '*** Function to get the RmSCtg for the specified RmCode form RmMst table
  FnRmSCtg = Trim(moCn.GetFldVal("Select RmSCtg from RmMst where RmPrtKey='" + ctCurrPrtn + "' and RmCd='" + mRmCd + "'"))
End Function
Private Sub PurCopy()
  '*** Copy Option is used to copy Raw Material Purchase Records
  '*** Check if valid Purchase Voucher is specified by the user. If not then copy is not possible
  
  Dim wRsPur As MwfLib.MDORowSet, wRsRmCdStr As MwfLib.MDORowSet
  Dim wRow As Single, wStr As String, wGrp As String, wSr As Single, wRmCdStr As String
  
  If ADC("wCpGrp") = "" Then DispMsg "Enter Valid Option For Grouping The Source Records", etError: Exit Sub
  
  '*** Jen (2.12 Patch)
  '****** Sachin 3.02 26-12-07 - Current Partition
  If Not moCn.RecSeek("Select 'x' From Txn where " + _
     "TCoCd= '" + ADC("wCpTCoCd") + "' and TTc= '" + ADC("wCpTTc") + "' and " + _
     "TYy= '" + ADC("wCpTYy") + "' and TChr= '" + ADC("wCpTChr") + "' and " + _
     "TNo= '" + CStr(ADC("wCpTNo")) + "' and TPrtKey='" + ctCurrPrtn + "' and (Select PValue from Param where PTyp= 'TC' And " + _
     "PMCd= '" + ADC("wCpTTc") + "' and PSCd= '')= 'PR' ") Then DispMsg "Invalid Purchase Voucher", etError: Exit Sub
  '*** Jen (2.12 Patch)
     
  '*** (Bef 25/03/06)
  'If Not moCn.RecSeek("Select 'x' From Txn where " + _
  '   "TCoCd= '" + ADC("TCoCd") + "' and TTc= '" + ADC("wCpTTc") + "' and " + _
  '   "TYy= '" + ADC("wCpTYy") + "' and TChr= '" + ADC("wCpTChr") + "' and " + _
  '   "TNo= '" + CStr(ADC("wCpTNo")) + "' and (Select PValue from Param where PTyp= 'TC' And " + _
  '   "PMCd= '" + ADC("wCpTTc") + "' and PSCd= '')= 'PR' ") Then DispMsg "Invalid Purchase Voucher", etError: Exit Sub
  '*** (Bef 25/03/06)

   '******************Geeta***********Emr207***********************************
     
  '*** Jen (2.12 Patch)
  '****** Sachin 3.02 26-12-07 - Current Partition
  If Not moCn.RecSeek("Select 'x' From Txn where " + _
     "TCoCd= '" + ADC("wCpTCoCd") + "' and TTc= '" + ADC("wCpTTc") + "' and " + _
     "TYy= '" + ADC("wCpTYy") + "' and TChr= '" + ADC("wCpTChr") + "' and " + _
     "TNo= '" + CStr(ADC("wCpTNo")) + "' and TPrtKey='" + ctCurrPrtn + "' and TCurrCd ='" + ADC("TCurrCd") + "'") Then DispMsg "Source && Target Currency Does Not Match", etError: Exit Sub
  '*** Jen (2.12 Patch)
        
  '*** (Bef 25/03/06)
  'If Not moCn.RecSeek("Select 'x' From Txn where " + _
  '  "TCoCd= '" + ADC("TCoCd") + "' and TTc= '" + ADC("wCpTTc") + "' and " + _
  '  "TYy= '" + ADC("wCpTYy") + "' and TChr= '" + ADC("wCpTChr") + "' and " + _
  '  "TNo= '" + CStr(ADC("wCpTNo")) + "' and tUSDYn ='" + ADC("tUSDYn") + "'") Then DispMsg "Source & Target Currency Does Not Match", etError: Exit Sub
  '*** (Bef 25/03/06)
  '***********************************************************
        
  wStr = "": wGrp = ""
  
  'Dim www As String, wRsRmCdStr As MwfLib.MDORowSet
  'www = "declare @wStr varchar(300) " + _
  '      "select @wStr = '' " + _
  '      "select @wStr= @wStr + ', '+ PMCd from param where ptyp= 'rmctg' " + _
  '      "select @wStr AS QQQ "
  'Set wRsRmCdStr = moCn.OpenRes(www)
  'MsgBox wRsRmCdStr!QQQ
  
  
  '*** (Bef 2.13)
  ''****** Sachin 2.12 - 04-01-2006 - TdDesc added to wStr below
  'Select Case UCase(ADC("wCpGrp"))
  'Case Is = UCase("Opt1")
  '  '*** 1 Custm Pur Rec For 1 Pur Rec
  '  wStr = " TdRmCd, TdRmQty, TdRmWt, TdPurRt, TdPurAmt, TdPcPerCt, RmCtg, RmSCtg, RmZCd, TdDesc "
  '  wGrp = " "
  'Case Is = UCase("Opt2")
  '  '*** Grp Src Recs On CustmCd,Rt
  '  wStr = " max(TdRmCd) as TdRmCd, sum(TdRmQty) as TdRmQty, sum(TdRmWt) as TdRmWt, " + _
  '         "TdPurRt, sum(TdPurAmt) as TdPurAmt, (case when sum(TdRmWt)> 0 then " + _
  '         "sum(TdRmQty)/sum(TdRmWt) else max(TdPcPerCt) end) as TdPcPerCt, max(RmCtg) as RmCtg, " + _
  '         "max(RmSCtg) as RmSCtg, RmZCd, '' As TdDesc "
  '  wGrp = " group by RmZCd, TdPurRt "
  'Case Is = UCase("Opt3")
  '  '*** Grp Src Recs On Custm,RmCd,Rt
  '  wStr = " TdRmCd, sum(TdRmQty) as TdRmQty, sum(TdRmWt) as TdRmWt, TdPurRt, sum(TdPurAmt) as TdPurAmt, " + _
  '         "(case when sum(TdRmWt)> 0 then sum(TdRmQty)/sum(TdRmWt) else max(TdPcPerCt) end) as " + _
  '         "TdPcPerCt, max(RmCtg) as RMCtg, max(RmSCtg) as RmSCtg, RmZCd, '' As TdDesc "
  '  wGrp = " group by RmZCd, TdRmCd, TdPurRt "
  'End Select
  '*** (Bef 2.13)
  
  '*** (Jen 2.13)
  Dim ws_Desc As String, ws_Wt As String, ws_PureRt As String, ws_Rt As String
  Dim ws_BestUse As String      ' **** Manali 3.5.0 - 07/11/08 - Copy Best Use fields
  
  ' ***** Manali 3.03 -20/06/08 - Pd Mod - 'L' included
  ws_Wt = " sum(TdRmWt* (case when R1.RmCtg in ('G', 'P', 'S', 'L') then R1.RmPurityZ " + _
          "                        / (case when R2.RmPurityZ> 0 then R2.RmPurityZ else 1 end) " + _
          "              else 1 end)) "
  
  '*** (Bef 24/01/07)
  'ws_PureRt = " max(TdPurRt/ (case when R1.RmCtg in ('G', 'P', 'S') and R1.RmPurityZ > 0 then R1.RmPurityZ " + _
  '        "                         * (case when R2.RmPurityZ > 0 then R2.RmPurityZ else 1 end) " + _
  '        "               else 1 end) ) "
  '*** (Bef 24/01/07)
  
  '*** (24/01/07)
  ' ***** Manali 3.03 -20/06/08 - Pd Mod - 'L' included
  ' ***** Manali 3.8.0 - Purchase Rate By QW
  ws_PureRt = " max(TdPurRt * (Case When TdRtByQW='Q' Then (TdRmQty/Case When TdRmWt=0 Then 1 Else TdRmWt End) Else 1 End) " + _
              "              / (case when R1.RmCtg in ('G', 'P', 'S', 'L') and R1.RmPurityZ > 0 then R1.RmPurityZ else 1 end) " + _
              "              * (case when R1.RmCtg in ('G', 'P', 'S', 'L') and R2.RmPurityZ > 0 then R2.RmPurityZ else 1 end) ) "
    
  ' Bef 3.8.0
  'ws_PureRt = " max(TdPurRt/ (case when R1.RmCtg in ('G', 'P', 'S', 'L') and R1.RmPurityZ > 0 then R1.RmPurityZ else 1 end) " + _
              "              * (case when R1.RmCtg in ('G', 'P', 'S', 'L') and R2.RmPurityZ > 0 then R2.RmPurityZ else 1 end) ) "
  '*** (24/01/07)
  
  ' ***** Manali 3.03 -20/06/08 - Pd Mod - 'L' included
  ws_Rt = " max(TdPurRt/ (case when R1.RmCtg in ('G', 'P', 'S', 'L') and R1.RmPurityZ > 0 " + _
          "                    then R1.RmPurityZ else 1 end)) "
  'kbs.132 best use Sr. number added in all cases
  Select Case UCase(ADC("wCpGrp"))
  Case Is = UCase("Opt1")
    '*** 1 Custm Pur Rec For 1 Pur Rec
    wGrp = " group by TdSr "
    ws_Desc = " max(TdDesc) as TdDesc "
   'ws_BestUse = IIF(mb_BestUse = True And ADC("wCpBsSzYn") = "Y", ", IsNull(max(RrFrLn), max(TdRmSz)) as RrFrLn, IsNull(max(RrToLn), max(TdRmSz)) as RrToLn ", ", 0.0 as RrFrLn, 0.0 as RrToLn ")
    ws_BestUse = IIF(mb_BestUse, ",Max(TdBstOdTc) as TdBstOdTc,Max(TdBstOdYy) as TdBstOdYy,Max(TdBstOdChr) as TdBstOdChr,Max(TdBstOdNo) as TdBstOdNo,Max(TdBstOdSr) as TdBstOdSr,Max(TdBstDmCd) as TdBstDmCd " + _
                    IIF(ADC("wCpBsSzYn") = "Y", ", IsNull(max(RrFrLn), max(TdRmSz)) as RrFrLn, IsNull(max(RrToLn), max(TdRmSz)) as RrToLn ", ", 0.0 as RrFrLn, 0.0 as RrToLn "), ",'' as TdBstOdTc,'' as TdBstOdYy,'' as TdBstOdChr,0 as TdBstOdNo,0 as TdBstOdSr,'' as TdBstDmCd, 0.0 as RrFrLn, 0.0 as RrToLn ")
    ' **** Manali 3.5.0 - 07/11/08 - Copy Best Use fields
    
  Case Is = UCase("Opt2")
    '*** Grp Src Recs On CustmCd,Rt
    ' ***** Manali 3.03 -20/06/08 - Pd Mod - 'L' included
    ' ***** Manali 3.8.0 - Purchase Rate By QW
    wGrp = " group by R1.RmZCd, convert (decimal (16, 3), TdPurRt* (Case When TdRtByQW='Q' Then (TdRmQty/Case When TdRmWt=0 Then 1 Else TdRmWt End) Else 1 End) " + _
           "                    / (case when R1.RmCtg in ('G', 'P', 'S', 'L') " + _
           "                    and R1.RmPurityZ > 0 then R1.RmPurityZ else 1 end)) " + IIF(mb_BestUse, ",TdBstOdTc,TdBstOdYy,TdBstOdChr,TdBstOdNo,TdBstOdSr,TdBstDmCd", "")
    ws_Desc = " '' as TdDesc "
    'ws_BestUse = ", 0.0 as RrFrLn, 0.0 as RrToLn "      ' **** Manali 3.5.0 - 07/11/08 - Copy Best Use fields
    ws_BestUse = IIF(mb_BestUse, ",Max(TdBstOdTc) as TdBstOdTc,Max(TdBstOdYy) as TdBstOdYy,Max(TdBstOdChr) as TdBstOdChr,Max(TdBstOdNo) as TdBstOdNo,Max(TdBstOdSr) as TdBstOdSr,Max(TdBstDmCd) as TdBstDmCd, 0.0 as RrFrLn, 0.0 as RrToLn ", ",'' as TdBstOdTc,'' as TdBstOdYy,'' as TdBstOdChr,0 as TdBstOdNo,0 as TdBstOdSr,'' as TdBstDmCd, 0.0 as RrFrLn, 0.0 as RrToLn ")
    
  Case Is = UCase("Opt3")
    '*** Grp Src Recs On Custm,RmCd,Rt
    ' ***** Manali 3.8.0 - Purchase Rate By QW
    wGrp = " group by R1.RmZCd, TdRmCd, TdPurRt*(Case When TdRtByQW='Q' Then (TdRmQty/Case When TdRmWt=0 Then 1 Else TdRmWt End) Else 1 End) " + IIF(mb_BestUse, ",TdBstOdTc,TdBstOdYy,TdBstOdChr,TdBstOdNo,TdBstOdSr,TdBstDmCd", "")
    ' Bef 3.8.0
    'wGrp = " group by R1.RmZCd, TdRmCd, TdPurRt "
    ws_Desc = " '' as TdDesc "
    'ws_BestUse = ", 0.0 as RrFrLn, 0.0 as RrToLn "      ' **** Manali 3.5.0 - 07/11/08 - Copy Best Use fields
    ws_BestUse = IIF(mb_BestUse, ",Max(TdBstOdTc) as TdBstOdTc,Max(TdBstOdYy) as TdBstOdYy,Max(TdBstOdChr) as TdBstOdChr,Max(TdBstOdNo) as TdBstOdNo,Max(TdBstOdSr) as TdBstOdSr,Max(TdBstDmCd) as TdBstDmCd, 0.0 as RrFrLn, 0.0 as RrToLn ", ",'' as TdBstOdTc,'' as TdBstOdYy,'' as TdBstOdChr,0 as TdBstOdNo,0 as TdBstOdSr,'' as TdBstDmCd, 0.0 as RrFrLn, 0.0 as RrToLn ")
   Case Is = UCase("Opt4")
    wGrp = " group by R1.RmZCd, TdRmCd,RrRngSz,TdPurRt*(Case When TdRtByQW='Q' Then (TdRmQty/Case When TdRmWt=0 Then 1 Else TdRmWt End) Else 1 End) " + IIF(mb_BestUse, ",TdBstOdTc,TdBstOdYy,TdBstOdChr,TdBstOdNo,TdBstOdSr,TdBstDmCd", "")
    ws_Desc = " '' as TdDesc "
    ws_BestUse = IIF(mb_BestUse, ",Max(TdBstOdTc) as TdBstOdTc,Max(TdBstOdYy) as TdBstOdYy,Max(TdBstOdChr) as TdBstOdChr,Max(TdBstOdNo) as TdBstOdNo,Max(TdBstOdSr) as TdBstOdSr,Max(TdBstDmCd) as TdBstDmCd, IsNull(max(RrFrLn), max(TdRmSz))  as RrFrLn, IsNull(max(RrToLn), max(TdRmSz)) as RrToLn ", ",'' as TdBstOdTc,'' as TdBstOdYy,'' as TdBstOdChr,0 as TdBstOdNo,0 as TdBstOdSr,'' as TdBstDmCd, 0.0 as RrFrLn, 0.0 as RrToLn ")
  End Select
  
  ' **** Manali 3.5.0 - 07/11/08 - Copy Best Use fields - ws_BestUse included
  ' ****** Manali 350Nxt - Change in TdPcPerCt
  wStr = " max(TdSr) as TdSr, max(TdRmCd) as TdRmCd, sum(TdRmQty) as TdRmQty, " + ws_Wt + " as TdRmWt, " + _
         ws_PureRt + " as qPurePurRt, " + ws_Rt + " as TdPurRt, sum(TdPurAmt) as TdPurAmt, " + _
         " (case when sum(TdRmWt)> 0 and sum(TdRmQty)> 0 and max(R1.RmCtg) in ('C', 'D') then sum(TdRmQty)/sum(TdRmWt) " + _
         " else max(TdPcPerCt) end) as TdPcPerCt, max(R1.RmCtg) as RmCtg, " + _
         " max(R1.RmSCtg) as RmSCtg, max(R1.RmZCd) as RmZCd, " + ws_Desc + ws_BestUse
  '*** (Jen 2.13)
        
  '*** Beginning the copy process
  GrdTxndZ.StartCopy
  MWLib.BeginProcess Me, "Copying Purchase Details ..."
  '****** (Jen 2.13) added join for Custom Code ******
  '****** Jen (2.12 Patch) changed adc("TCoCd") to adc("wCpTCoCd") ******
  '****** Sachin 3.02 26-12-07 - Current Partition
  ' **** Manali 3.5.0 - 07/11/08 - Copy Best Use fields - Join with RmRt added
  ' **** Manali 3.6.0 - 19/09/09 - Copy Best Use Size fields - ask Y/N
  Set wRsPur = moCn.OpenRes(" Select " + wStr + _
                            " from Txnd join RmMst R1 on R1.RmCd= TdRmCd and R1.RmPrtKey=TdPrtKey " + _
                            " join RmMst R2 on R2.RmCd= R1.RmZCd and R2.RmPrtKey=R1.RmPrtKey " + _
                            " join Loc LFr on LFr.LocCoCd= TdCoCd and LFr.LocCd= TdFrRmLoc " + _
                            " join Loc LTo on LTo.LocCoCd= TdCoCd and LTo.LocCd= TdToRmLoc " + _
                            " Join Txn On TCoCd=TdCoCd and TdYy=TYy and TdChr=TChr and TdNo=TNo and TIdNo=TdTIdNo " + _
                            " left outer join RmRt on RrTcTyp='RSZ' and RrCmCd='" + ctSelfCmCd + "' and RrCtg=R1.RmCtg and RrSCtg=R1.RmSCtg and RrCd='' and RrFrLn<= TdRmSz and RrToLn>= TdRmSz " + _
                            " Where TdCoCd= '" + ADC("wCpTCoCd") + "' and " + _
                            "   TdTc= '" + ADC("wCpTTc") + "' and " + _
                            "   TdYy= '" + ADC("wCpTYy") + "' and " + _
                            "   TdChr= '" + ADC("wCpTChr") + "' and " + _
                            "   TdNo= '" + CStr(ADC("wCpTNo")) + "' and TdPrtKey='" + ctCurrPrtn + "' and " + _
                            "((LFr.LocTyp= 'XP' and TdFrRmDc= 'C') or (LTo.LocTyp= 'XP' and TdToRmDc= 'C')) " + _
                            wGrp + " Order By " + IIF(UCase(ADC("wCpGrp")) = UCase("Opt1"), " TdSr, ", "") + _
                            " RmCtg, RmSCtg, RmZCd, TdRmCd, qPurePurRt ")
                            
  '****** (Bef 2.13) " Order By RmCtg, RmSCtg, RmZCd, TdRmCd, TdPurRt "
  With wRsPur
    wSr = GrdTxndZ.MaxVal("TzSr") + 1
    Do While Not (.EOF Or .BOF)
      MWLib.RunningProcess "Copying Purchase Details "
      GrdTxndZ.AddItem
      wRow = GrdTxndZ.Rows - 1
      GrdTxndZ.Value(wRow, "TzSr") = wSr
      GrdTxndZ.Value(wRow, "TzRmZCd") = !RmZCd
      Select Case UCase(ADC("wCpGrp"))
      Case Is = UCase("Opt1")   ', UCase("Opt3")
        GrdTxndZ.Value(wRow, "TzRmCdDesc") = !TdRmCd
        ' **** Manali 3.5.0 - 07/11/08 - Copy Best Use fields
        If mb_BestUse = True Then GrdTxndZ.Value(wRow, "TzRmCd") = !TdRmCd
        ' **** Manali 3.5.0 - 07/11/08 - Copy Best Use fields
      
      Case Is = UCase("Opt3"), UCase("Opt4")
        GrdTxndZ.Value(wRow, "TzRmCdDesc") = !TdRmCd
  
      Case Is = UCase("Opt2")
        wRmCdStr = ""
        '*** (Bef 2.13)
        'Set wRsRmCdStr = moCn.OpenRes("Select TdRmCd from Txnd join RmMst on RmCd= TdRmCd " + _
                         "where TdCoCd= '" + ADC("wCpTCoCd") + "' and " + _
                         "TdTc= '" + ADC("wCpTTc") + "' and TdYy= '" + ADC("wCpTYy") + "' and " + _
                         "TdChr= '" + ADC("wCpTChr") + "' and TdNo= '" + CStr(ADC("wCpTNo")) + "' and " + _
                         "RmZCd= '" + !RmZCd + "' and TdPurRt= " + CStr(!TdPurRt) + " Order By TdRmCd ")
        ' *** (Bef 2.13)
        
        ' *** (Jen 2.13)
        ' ***** Jen (2.12 Patch) changed adc("TCoCd") to adc("wCpTCoCd") ******
        ' ***** Sachin 3.02 26-12-07 - Current Partition
        ' ***** Manali 3.03 -20/06/08 - Pd Mod - 'L' included
        Set wRsRmCdStr = moCn.OpenRes(" Select TdRmCd from Txnd join RmMst on RmCd= TdRmCd and RmPrtKey=TdPrtKey " + _
                         " where TdCoCd= '" + ADC("wCpTCoCd") + "' and TdTc= '" + ADC("wCpTTc") + "' " + _
                         " and TdYy= '" + ADC("wCpTYy") + "' and TdChr= '" + ADC("wCpTChr") + "' " + _
                         " and TdNo= '" + CStr(ADC("wCpTNo")) + "' and TdPrtKey='" + ctCurrPrtn + "' and RmZCd= '" + !RmZCd + "' " + _
                         " and convert (decimal (16, 3), TdPurRt/ (case when RmCtg in ('G', 'P', 'S', 'L') and RmPurityZ > 0 " + _
                         "    then RmPurityZ else 1 end)) = convert (decimal (16, 3), " + CStr(!TdPurRt) + ") " + _
                         " Order By TdRmCd ")
        '*** (Jen 2.13)
        Do While Not (wRsRmCdStr.EOF Or wRsRmCdStr.BOF)
          wRmCdStr = IIF(Len(wRmCdStr + IIF(wRmCdStr = "", "", ", ") + wRsRmCdStr!TdRmCd) <= 20, wRmCdStr + IIF(wRmCdStr = "", "", ", ") + wRsRmCdStr!TdRmCd, wRmCdStr)
          wRsRmCdStr.MoveNext
        Loop
        GrdTxndZ.Value(wRow, "TzRmCdDesc") = wRmCdStr
      End Select
      
      GrdTxndZ.Value(wRow, "TzPcPerCt") = !TdPcPerCt
      GrdTxndZ.Value(wRow, "TzRmQty") = !TdRmQty
      GrdTxndZ.Value(wRow, "TzRmWt") = !TdRmWt
      GrdTxndZ.Value(wRow, "TzPurRt") = !qPurePurRt   '****** (Jen 2.13)
      'GrdTxndZ.Value(wRow, "TzPurRt") = !TdPurRt    '****** (Bef 2.13)
      GrdTxndZ.Value(wRow, "TzPurAmt") = !TdPurAmt
      GrdTxndZ.Value(wRow, "TzRmDespWt") = 0
      GrdTxndZ.Value(wRow, "TzCls") = "N"
      ' **** Sachin 2.12 - 04-01-2006 - Desc in Custom Purc
      GrdTxndZ.Value(wRow, "TzDesc") = !TdDesc
      ' **** Sachin 2.12  - 04-01-2006 - Desc in Custom Purc
      ' **** Manali 3.5.0 - 07/11/08 - Copy Best Use fields
      GrdTxndZ.Value(wRow, "TzPrtKey") = ctCurrPrtn
      GrdTxndZ.Value(wRow, "TzClsDt") = MWLib.EmptyDate     ' ***** Manali 350Nxt   TzClsDt added
      ' **** Manali 3.5.0 - 07/11/08 - Copy Best Use fields
        GrdTxndZ.Value(wRow, "TzRmSzFr") = !RrFrLn
        GrdTxndZ.Value(wRow, "TzRmSzTo") = !RrToLn
      ' **** Manali 3.5.0 - 07/11/08 - Copy Best Use fields

      
      '4.1.2.0
      If mb_BestUse Then
        GrdTxndZ.Value(wRow, "TzOdTc") = !TdBstOdTc
        GrdTxndZ.Value(wRow, "TzOdYy") = !TdBstOdYy
        GrdTxndZ.Value(wRow, "TzOdChr") = !TdBstOdChr
        GrdTxndZ.Value(wRow, "TzOdNo") = !TdBstOdNo
        GrdTxndZ.Value(wRow, "TzOdSr") = !TdBstOdSr
        GrdTxndZ.Value(wRow, "TzDmCd") = !TdBstDmCd
      End If
      
      GrdTxndZ.SaveRec (wRow)
      wSr = wSr + 1
      .MoveNext
    Loop
  End With

  MWLib.EndProcess Me
  GrdTxndZ.EndCopy
  Dim wRsPurAmt As MwfLib.MDORowSet
  Set wRsPurAmt = moCn.OpenRes("Select TAddAmt1,TAddAmt2,TAddAmt3,TAddAmt4 from Txn " + _
                         " where TCoCd= '" + ADC("wCpTCoCd") + "' and TTc= '" + ADC("wCpTTc") + "' " + _
                         " and TYy= '" + ADC("wCpTYy") + "' and TChr= '" + ADC("wCpTChr") + "' " + _
                         " and TNo= '" + CStr(ADC("wCpTNo")) + "' and TPrtKey='" + ctCurrPrtn + "'")
   
  If wRsPurAmt.RecCount > 0 Then
    ADC("TAddAmt1") = ADC("TAddAmt1") + wRsPurAmt!TAddAmt1
    ADC("TAddAmt2") = ADC("TAddAmt2") + wRsPurAmt!TAddAmt2
    ADC("TAddAmt3") = ADC("TAddAmt3") + wRsPurAmt!TAddAmt3
    ADC("TAddAmt4") = ADC("TAddAmt4") + wRsPurAmt!TAddAmt4
  End If
  
  Set wRsPur = Nothing
  '*** End of the copy process
  Call CmdCP_Click
End Sub
Private Sub DispFra(ByVal pv_CustmPurFra As en_CustmPurFra)
  Select Case pv_CustmPurFra
  Case Is = Cpy
    If FraCP.Visible = True Then
      FraCP.Visible = False
      FraCP.Enabled = False
      'Call EnaDisaCmds(False, CmdCp)   '*** (Bef 3.1.2)
      Call EnaDisaCmds(False, CmdCP, Cpy)   '*** (Jen 3.1.2)
      CmdCP.SetFocus
    Else
      FraCP.Visible = True
      FraCP.Enabled = True
      CmdCP.SetFocus
      FraCP.ZOrder
      Call EnaDisaCmds(True, CmdCP)
    End If
  Case Is = Pwd
'    If FraPwd.Visible = True Then
'      FraPwd.Visible = False
'      FraPwd.Enabled = False
'      Call EnaDisaCmds(False, CmdPwd)
'      CmdPwd.SetFocus
'    Else
'      FraPwd.Visible = True
'      FraPwd.Enabled = True
'      CmdPwd.SetFocus
'      FraPwd.ZOrder
'      Call EnaDisaCmds(True, CmdPwd)
'    End If
  
  '*** (Jen 3.1.2)
  Case Is = DET
    If FraGrdTxndZDetsDisp.Enabled = True Then
      'FraGrdTxndZDetsDisp.Visible = False
      FraGrdTxndZDetsDisp.Enabled = False
      'Call EnaDisaCmds(False, CmdGrdTxndzDets)  '*** (Bef 3.1.2)
      Call EnaDisaCmds(False, CmdGrdTxndzDets, DET)   '*** (Jen 3.1.2)
      GrdTxndZ.SetFocus
      'CmdGrdTxndzDets.SetFocus
    Else
      'FraGrdTxndZDetsDisp.Visible = True
      FraGrdTxndZDetsDisp.Enabled = True
      'CmdGrdTxndzDets.SetFocus
      ADC("wTzDesc").SetFocus
      'FraGrdTxndZDetsDisp.ZOrder
      Call EnaDisaCmds(True, CmdGrdTxndzDets)
    End If
  '*** (Jen 3.1.2)
  
   '4.3.0.0
    Case Is = ChgDt
    If FraCD.Visible = True Then
      FraCD.Visible = False
      FraCD.Enabled = False
      Call EnaDisaCmds(False, CmdChgDt, ChgDt)
      CmdChgDt.SetFocus
    Else
      FraCD.Visible = True
      FraCD.Enabled = True
      CmdChgDt.SetFocus
      FraCD.ZOrder
      ADC("wTDt") = ADC("TDt")
      Call EnaDisaCmds(True, CmdChgDt)
    End If
  End Select
End Sub
Private Sub EnaDisaCmds(ByVal pv_ShowFra As Boolean, Optional ByVal pv_CmdBut As MWCTL_BTN1, Optional ByVal pv_DispFra As en_CustmPurFra)
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

  '*** (Jen 3.1.2)
  If pv_ShowFra = True Then
    Call StoreState
  End If
  '*** (Jen 3.1.2)

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
      '*** (Jen 3.1.2)
      If Not IsMissing(pv_DispFra) Then
        pv_CmdBut.CausesValidation = mArr_Cv(pv_DispFra)
      End If
      '*** (Jen 3.1.2)
    End If
  End If
  FraNKeyAll.Enabled = Not pv_ShowFra
  ADC.AllowSave = Not pv_ShowFra
  ADC.AllowDelete = Not pv_ShowFra
  CmdCP.Enabled = Not pv_ShowFra
  CmdChgDt.Enabled = Not pv_ShowFra         '4.1.3.0
  CmdGrdTxndzDets.Enabled = Not pv_ShowFra  '*** (Jen 3.1.2)
  CmdInvHd.Enabled = Not pv_ShowFra 'pg.11
'  CmdPwd.Enabled = Not pv_ShowFra
  If Not pv_CmdBut Is Nothing Then pv_CmdBut.Enabled = True
End Sub
Private Sub StoreState()
  '*** This procedure is used to store the CausesValidation State
  '*** This is used for restoring the original states in the future
  '*** (Jen 3.1.2)
  mArr_Cv(Cpy) = CmdCP.CausesValidation
  mArr_Cv(DET) = CmdGrdTxndzDets.CausesValidation
  '*** (Jen 3.1.2)
End Sub
Private Sub HideAllFras()
  FraNKeyAll.Visible = True
  FraCP.Visible = False
  FraCP.Enabled = False
  
  '*** (Je 3.1.2)
  'FraGrdTxndZDetsDisp.Visible = False
  FraGrdTxndZDetsDisp.Enabled = False
  '*** (Je 3.1.2)
  
'  FraPwd.Visible = False
'  FraPwd.Enabled = False
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
  ws_BtnStr = "CmdCPGo"
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
    
    Set moCn = Nothing    '*** (09/08/05)    '*** (Bef speed) bef 24/11/06
End Sub
Private Sub GrdTxndZ_Validate(Cancel As Boolean)
  Cancel = GrdTxndZ.Validate
End Sub


'TxndZ
'
'TzTc          TzYy          TzChr          TzNo
'TzSr          TzSrNo        TzRmZCd        TzRmCdDesc
'TzPcPerCt     TzRmQty       TzRmWt         TzKey
'TzRefYy       TzRefKey      TzPurRt        TzPurAmt
'TzBm          TzRmDespWt    TzCls

