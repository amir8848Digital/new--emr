VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Object = "{AB2CC32A-E652-42A2-B341-26ECE02F6386}#2.0#0"; "MwfCtl.ocx"
Begin VB.Form EmrFrmLabRt 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Labour Rate Master"
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
   Begin VB.Frame FraImpExc 
      Height          =   1515
      Left            =   2160
      TabIndex        =   43
      Top             =   7680
      Width           =   8235
      Begin MwfCtl.MWCTL_MED ATXT 
         Height          =   285
         Index           =   248
         Left            =   660
         TabIndex        =   44
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
         TabIndex        =   45
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
         Left            =   2220
         TabIndex        =   46
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
         Left            =   3870
         TabIndex        =   47
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
         TabIndex        =   48
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
         TabIndex        =   49
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
         Caption         =   "To Rw"
         BeginProperty Font 
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
         Left            =   3030
         TabIndex        =   54
         Top             =   1065
         Width           =   645
      End
      Begin VB.Label ALBL 
         BackStyle       =   0  'Transparent
         Caption         =   "Fr Rw"
         BeginProperty Font 
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
         TabIndex        =   53
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
         TabIndex        =   52
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
         TabIndex        =   51
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
         TabIndex        =   50
         Top             =   120
         Width           =   5640
      End
   End
   Begin VB.Frame FraDelRt 
      Height          =   1140
      Left            =   5520
      TabIndex        =   61
      Top             =   8160
      Width           =   4860
      Begin MwfCtl.MWCTL_BTN1 CmdDelRtGo 
         Height          =   465
         Left            =   4050
         TabIndex        =   62
         ToolTipText     =   "Option to Delete Rates"
         Top             =   600
         Width           =   690
         _ExtentX        =   1217
         _ExtentY        =   820
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
         Caption         =   "Delete Rates will Delete the Rate chart for the selected scope, are you sure you want to delete?"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1050
         Index           =   3
         Left            =   30
         TabIndex        =   63
         Top             =   90
         Width           =   4830
      End
   End
   Begin VB.Frame FraPwd 
      Height          =   870
      Left            =   4560
      TabIndex        =   57
      Top             =   8400
      Width           =   4530
      Begin VB.TextBox TxtPwd 
         Height          =   285
         IMEMode         =   3  'DISABLE
         Left            =   2040
         MaxLength       =   30
         PasswordChar    =   "*"
         TabIndex        =   58
         Top             =   405
         Width           =   2340
      End
      Begin VB.Label ALBL 
         BackStyle       =   0  'Transparent
         Caption         =   "Password To Edit"
         BeginProperty Font 
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
         Left            =   60
         TabIndex        =   60
         Top             =   405
         Width           =   1935
      End
      Begin VB.Label LblMain 
         BackStyle       =   0  'Transparent
         BorderStyle     =   1  'Fixed Single
         Caption         =   "  Password"
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
         Index           =   1000
         Left            =   -30
         TabIndex        =   59
         Top             =   90
         Width           =   4530
      End
   End
   Begin MwfCtl.MWCTL_HLP MWCTL_HLP1 
      Height          =   405
      Left            =   8910
      TabIndex        =   18
      TabStop         =   0   'False
      Top             =   9360
      Visible         =   0   'False
      Width           =   555
      _ExtentX        =   979
      _ExtentY        =   714
   End
   Begin VB.Frame FRA_KEYS 
      BorderStyle     =   0  'None
      Height          =   570
      Left            =   -36
      TabIndex        =   20
      Top             =   24
      Width           =   9990
      Begin MwfCtl.MWCTL_MED ATXT 
         Height          =   285
         Index           =   0
         Left            =   1395
         TabIndex        =   0
         ToolTipText     =   "Enter Labour Main Code"
         Top             =   0
         Width           =   885
         _ExtentX        =   1561
         _ExtentY        =   503
         MaxLength       =   5
         DataType        =   4
         IdName          =   "WLRMCD"
         CmpStr          =   "LrMCd="
      End
      Begin MwfCtl.MWCTL_MED ATXT 
         Height          =   285
         Index           =   5
         Left            =   7965
         TabIndex        =   6
         ToolTipText     =   "Enter Record Sort Option"
         Top             =   0
         Width           =   1380
         _ExtentX        =   2434
         _ExtentY        =   503
         MaxLength       =   8
         DataType        =   4
         IdName          =   "WLRSRT"
      End
      Begin MwfCtl.MWCTL_MED ATXT 
         Height          =   285
         Index           =   8
         Left            =   3405
         TabIndex        =   2
         ToolTipText     =   "Enter From Customer"
         Top             =   0
         Width           =   1320
         _ExtentX        =   2328
         _ExtentY        =   503
         MaxLength       =   8
         DataType        =   4
         IdName          =   "WLRCMCDFR"
         CmpStr          =   "LrCmCd>="
      End
      Begin MwfCtl.MWCTL_MED ATXT 
         Height          =   285
         Index           =   9
         Left            =   3405
         TabIndex        =   3
         ToolTipText     =   "Enter To Customer"
         Top             =   285
         Width           =   1320
         _ExtentX        =   2328
         _ExtentY        =   503
         MaxLength       =   8
         DataType        =   4
         IdName          =   "WLRCMCDTO"
         CmpStr          =   "LrCmCd<="
      End
      Begin MwfCtl.MWCTL_MED ATXT 
         Height          =   285
         Index           =   10
         Left            =   1395
         TabIndex        =   1
         ToolTipText     =   "Enter Customer Category"
         Top             =   285
         Width           =   480
         _ExtentX        =   847
         _ExtentY        =   503
         MaxLength       =   2
         DataType        =   4
         IdName          =   "WLRCMCTG"
         CmpStr          =   "LrCmCtg ="
      End
      Begin MwfCtl.MWCTL_MED ATXT 
         Height          =   285
         Index           =   28
         Left            =   6015
         TabIndex        =   4
         ToolTipText     =   "Enter From Currency Code"
         Top             =   0
         Width           =   900
         _ExtentX        =   1588
         _ExtentY        =   503
         MaxLength       =   5
         DataType        =   4
         IdName          =   "WLRCMCURCDFR"
         CmpStr          =   "LrCmCurCd>="
      End
      Begin MwfCtl.MWCTL_MED ATXT 
         Height          =   285
         Index           =   29
         Left            =   6015
         TabIndex        =   5
         ToolTipText     =   "Enter To Currency Code"
         Top             =   285
         Width           =   900
         _ExtentX        =   1588
         _ExtentY        =   503
         MaxLength       =   5
         DataType        =   4
         IdName          =   "WLRCMCURCDTO"
         CmpStr          =   "LrCmCurCd<="
      End
      Begin VB.Label ALBL 
         BackStyle       =   0  'Transparent
         Caption         =   "Curr Cd Fr"
         BeginProperty Font 
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
         Left            =   4950
         TabIndex        =   37
         Top             =   0
         Width           =   1095
      End
      Begin VB.Label ALBL 
         BackStyle       =   0  'Transparent
         Caption         =   "            To"
         BeginProperty Font 
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
         Left            =   4950
         TabIndex        =   36
         Top             =   285
         Width           =   1095
      End
      Begin VB.Label ALBL 
         BackStyle       =   0  'Transparent
         Caption         =   "Cust Ctg"
         BeginProperty Font 
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
         TabIndex        =   31
         Top             =   270
         Width           =   885
      End
      Begin VB.Label ALBL 
         BackStyle       =   0  'Transparent
         Caption         =   "          To"
         BeginProperty Font 
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
         Left            =   2490
         TabIndex        =   25
         Top             =   285
         Width           =   915
      End
      Begin VB.Label ALBL 
         BackStyle       =   0  'Transparent
         Caption         =   "Cust  Fr"
         BeginProperty Font 
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
         Left            =   2580
         TabIndex        =   24
         Top             =   0
         Width           =   795
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
         Index           =   5
         Left            =   7185
         TabIndex        =   23
         Top             =   0
         Width           =   825
      End
      Begin VB.Label ALBL 
         BackStyle       =   0  'Transparent
         Caption         =   "Lab Main Cd"
         BeginProperty Font 
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
         TabIndex        =   22
         Top             =   -15
         Width           =   1335
      End
   End
   Begin VB.Frame FRA_NKEYS 
      BorderStyle     =   0  'None
      Enabled         =   0   'False
      Height          =   9300
      Left            =   -120
      TabIndex        =   19
      Top             =   0
      Width           =   13935
      Begin VB.Frame FraCpy 
         Height          =   1425
         Left            =   450
         TabIndex        =   26
         Top             =   7785
         Width           =   11640
         Begin MwfCtl.MWCTL_BTN1 CmdCpyGo 
            Height          =   285
            Left            =   10890
            TabIndex        =   17
            ToolTipText     =   "Press This Button To Copy"
            Top             =   990
            Width           =   555
            _ExtentX        =   979
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
            Index           =   11
            Left            =   1440
            TabIndex        =   9
            ToolTipText     =   "Enter Copy Source Customer"
            Top             =   705
            Width           =   1320
            _ExtentX        =   2328
            _ExtentY        =   503
            MaxLength       =   8
            DataType        =   4
            ReCalcParent    =   "WCPYCURCDFR"
            IdName          =   "WCUSTFR"
         End
         Begin MwfCtl.MWCTL_MED ATXT 
            Height          =   285
            Index           =   12
            Left            =   1440
            TabIndex        =   11
            ToolTipText     =   "Enter Copy Target Customer"
            Top             =   990
            Width           =   1320
            _ExtentX        =   2328
            _ExtentY        =   503
            MaxLength       =   8
            DataType        =   4
            ReCalcParent    =   "WCPYCURCDTO"
            IdName          =   "WCUSTTO"
         End
         Begin MwfCtl.MWCTL_MED ATXT 
            Height          =   285
            Index           =   2
            Left            =   9390
            TabIndex        =   15
            ToolTipText     =   "Enter From Raw Material Code "
            Top             =   705
            Width           =   1320
            _ExtentX        =   2328
            _ExtentY        =   503
            MaxLength       =   8
            DataType        =   4
            IdName          =   "WSCDFR"
         End
         Begin MwfCtl.MWCTL_MED ATXT 
            Height          =   285
            Index           =   3
            Left            =   9390
            TabIndex        =   16
            ToolTipText     =   "Enter To Raw Material Code "
            Top             =   990
            Width           =   1320
            _ExtentX        =   2328
            _ExtentY        =   503
            MaxLength       =   8
            DataType        =   4
            IdName          =   "WSCDTO"
         End
         Begin MwfCtl.MWCTL_MED ATXT 
            Height          =   285
            Index           =   1
            Left            =   1440
            TabIndex        =   8
            ToolTipText     =   "Enter Source Customer Category"
            Top             =   420
            Width           =   480
            _ExtentX        =   847
            _ExtentY        =   503
            MaxLength       =   2
            DataType        =   4
            IdName          =   "WCMCTGFR"
         End
         Begin MwfCtl.MWCTL_MED ATXT 
            Height          =   285
            Index           =   20
            Left            =   3810
            TabIndex        =   10
            ToolTipText     =   "Enter From Currency  Code"
            Top             =   705
            Width           =   885
            _ExtentX        =   1561
            _ExtentY        =   503
            MaxLength       =   5
            DataType        =   4
            ReCalcParent    =   "WCNVRTFR,WMULDIVFR"
            ReCalcOn        =   "WCUSTFR"
            IdName          =   "WCPYCURCDFR"
         End
         Begin MwfCtl.MWCTL_MED ATXT 
            Height          =   285
            Index           =   23
            Left            =   3810
            TabIndex        =   12
            ToolTipText     =   "Enter To Currency  Code"
            Top             =   990
            Width           =   885
            _ExtentX        =   1561
            _ExtentY        =   503
            MaxLength       =   5
            DataType        =   4
            ReCalcParent    =   "WCNVRTFR,WMULDIVFR"
            ReCalcOn        =   "WCUSTTO"
            IdName          =   "WCPYCURCDTO"
         End
         Begin MwfCtl.MWCTL_MED ATXT 
            Height          =   285
            Index           =   24
            Left            =   5760
            TabIndex        =   13
            ToolTipText     =   "Enter From Conversion Factor"
            Top             =   990
            Width           =   1005
            _ExtentX        =   1773
            _ExtentY        =   503
            Mask            =   "###0.0000"
            MaxLength       =   9
            DataType        =   2
            ReCalcOn        =   "WCPYCURCDFR,WCPYCURCDTO"
            IdName          =   "WCNVRTFR"
         End
         Begin MwfCtl.MWCTL_MED ATXT 
            Height          =   285
            Index           =   26
            Left            =   7740
            TabIndex        =   14
            Top             =   990
            Width           =   375
            _ExtentX        =   661
            _ExtentY        =   503
            Enabled         =   0   'False
            DataType        =   4
            ReCalcOn        =   "WCPYCURCDFR,WCPYCURCDTO"
            IdName          =   "WMULDIVFR"
         End
         Begin VB.Label LblCurCdFr 
            BackStyle       =   0  'Transparent
            Caption         =   "Curr Cd "
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Left            =   2970
            TabIndex        =   41
            Top             =   705
            Width           =   915
         End
         Begin VB.Label LblCurCdTo 
            BackStyle       =   0  'Transparent
            Caption         =   "Curr Cd"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Left            =   2970
            TabIndex        =   40
            Top             =   990
            Width           =   915
         End
         Begin VB.Label LblCnvRtFr 
            BackStyle       =   0  'Transparent
            Caption         =   "Conv Fct"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Left            =   4860
            TabIndex        =   39
            Top             =   990
            Width           =   885
         End
         Begin VB.Label LblMulDivFr 
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
            Left            =   6960
            TabIndex        =   38
            Top             =   990
            Width           =   765
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
            Height          =   270
            Index           =   0
            Left            =   -30
            TabIndex        =   33
            Top             =   90
            Width           =   11670
         End
         Begin VB.Label ALBL 
            BackStyle       =   0  'Transparent
            Caption         =   "Cust Ctg"
            BeginProperty Font 
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
            TabIndex        =   32
            Top             =   420
            Width           =   885
         End
         Begin VB.Label ALBL 
            BackStyle       =   0  'Transparent
            Caption         =   "Source Cust"
            BeginProperty Font 
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
            TabIndex        =   30
            Top             =   705
            Width           =   1395
         End
         Begin VB.Label ALBL 
            BackStyle       =   0  'Transparent
            Caption         =   "Target Cust"
            BeginProperty Font 
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
            TabIndex        =   29
            Top             =   990
            Width           =   1395
         End
         Begin VB.Label ALBL 
            BackStyle       =   0  'Transparent
            Caption         =   "Sub Cd Fr"
            BeginProperty Font 
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
            Left            =   8250
            TabIndex        =   28
            Top             =   705
            Width           =   1095
         End
         Begin VB.Label ALBL 
            BackStyle       =   0  'Transparent
            Caption         =   "             To"
            BeginProperty Font 
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
            Left            =   8250
            TabIndex        =   27
            Top             =   990
            Width           =   1095
         End
      End
      Begin VB.Frame FraNKeyAll 
         BorderStyle     =   0  'None
         Height          =   9165
         Left            =   150
         TabIndex        =   21
         Top             =   90
         Width           =   13740
         Begin MwfCtl.MWCTL_FLX GrdLab 
            Height          =   8550
            Left            =   0
            TabIndex        =   7
            Top             =   570
            Width           =   12075
            _ExtentX        =   21299
            _ExtentY        =   15081
            Cols            =   11
            colname1        =   "LRCMCD"
            heading1        =   "Cust Cd        "
            datafld1        =   "LrCmCd"
            datatype1       =   4
            recalcparent1   =   "GrdLab(LRCSTRT),GrdLab(LRCMCURCD)"
            maxlength1      =   8
            tooltiptext1    =   "Enter Customer Code"
            colname2        =   "LRCMCURCD"
            heading2        =   "Cur Cd     "
            datafld2        =   "LrCmCurCd"
            datatype2       =   4
            recalcon2       =   "GrdLab(LRCMCD)"
            maxlength2      =   5
            tooltiptext2    =   "Enter Currency Code"
            colname3        =   "LRSCD"
            heading3        =   "Lab Sub Cd   "
            datafld3        =   "LrSCd"
            datatype3       =   4
            recalcparent3   =   "GrdLab(LRQW)"
            maxlength3      =   8
            tooltiptext3    =   "Enter Labour Sub Code"
            colname4        =   "LRQW"
            heading4        =   "Lab Q/W"
            datafld4        =   "LrQw"
            datatype4       =   4
            recalcon4       =   "GrdLab(LRSCD)"
            recalcparent4   =   "GrdLab(LRFRWT),GrdLab(LRTOWT),GrdLab(LRSALMIN),GrdLab(LRSALMAX)"
            maxlength4      =   1
            tooltiptext4    =   "Enter Labour Q/W "
            colname5        =   "LRFRWT"
            heading5        =   "From Wt "
            datafld5        =   "LrFrWt"
            datatype5       =   2
            mask5           =   "###0.000"
            recalcon5       =   "GrdLab(LRQW)"
            maxlength5      =   8
            tooltiptext5    =   "Enter From Weight "
            colname6        =   "LRTOWT"
            heading6        =   "To Wt     "
            datafld6        =   "LrToWt"
            datatype6       =   2
            mask6           =   "###0.000"
            recalcon6       =   "GrdLab(LRQW)"
            maxlength6      =   8
            tooltiptext6    =   "Enter To Weight"
            colname7        =   "LRCSTRT"
            heading7        =   "Cost Rate    "
            datafld7        =   "LrCstRt"
            datatype7       =   2
            mask7           =   "#####0.00"
            recalcon7       =   "GrdLab(LRCMCD)"
            maxlength7      =   9
            tooltiptext7    =   "Enter Cost Rate"
            colname8        =   "LRSALRT"
            heading8        =   "Sales Rate  "
            datafld8        =   "LrSalRt"
            datatype8       =   2
            mask8           =   "-#####0.00"
            maxlength8      =   10
            tooltiptext8    =   "Enter Sales Rate"
            colname9        =   "LRSALMIN"
            heading9        =   "Min Sales Val"
            datafld9        =   "LrSalMin"
            datatype9       =   2
            mask9           =   "#####0.00"
            recalcon9       =   "GrdLab(LRQW)"
            maxlength9      =   9
            tooltiptext9    =   "Enter Minimum Sales Rate"
            colname10       =   "LRSALMAX"
            heading10       =   "Max Sales Val"
            datafld10       =   "LrSalMax"
            datatype10      =   2
            mask10          =   "#####0.00"
            recalcon10      =   "GrdLab(LRQW)"
            maxlength10     =   9
            tooltiptext10   =   "Enter Maximum Sales Rate"
         End
      End
   End
   Begin MwfCtl.MWCTL_BTN1 CmdCpy 
      Height          =   465
      Left            =   4830
      TabIndex        =   34
      ToolTipText     =   "Option To Copy The Labour Rates From One Customer To Another"
      Top             =   9360
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
   Begin MwfCtl.MWCTL_ADC ADC 
      CausesValidation=   0   'False
      Height          =   555
      Left            =   0
      TabIndex        =   35
      Top             =   9360
      Width           =   12090
      _ExtentX        =   21325
      _ExtentY        =   979
      AllowAdd        =   0   'False
      AllowDelete     =   0   'False
      BtnWidth        =   885
      BtnHeight       =   435
      BtnMaskColor    =   8454143
      FontSz          =   8.25
   End
   Begin MwfCtl.MWCTL_BTN1 CmdImpExc 
      Height          =   465
      Left            =   5640
      TabIndex        =   42
      ToolTipText     =   "Copy Option"
      Top             =   9360
      Width           =   840
      _ExtentX        =   1482
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
   Begin MSComDlg.CommonDialog CdImpExc 
      Left            =   2160
      Top             =   9240
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin MwfCtl.MWCTL_BTN1 CmdPwd 
      Height          =   465
      Left            =   6480
      TabIndex        =   55
      ToolTipText     =   "Enter Password to Edit Design"
      Top             =   9360
      Width           =   960
      _ExtentX        =   1693
      _ExtentY        =   820
      BackColor       =   -2147483633
      MaskColor       =   8454143
      Caption         =   "&Password"
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
   Begin MwfCtl.MWCTL_BTN1 CmdDelRt 
      Height          =   465
      Left            =   7440
      TabIndex        =   56
      ToolTipText     =   "Option to Delete Rates"
      Top             =   9360
      Width           =   945
      _ExtentX        =   1667
      _ExtentY        =   820
      BackColor       =   -2147483633
      MaskColor       =   8454143
      Caption         =   "&Delete Rates"
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
Attribute VB_Name = "EmrFrmLabRt"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'--- General comments on the entry ----
' FraCpy, with CmdCpy, CmdCpyGo, wCustFr, wCustTo, wSCdFr, wSCdTo
' X-'', A-LabRt
'
' private Sub CmpString to set the values for Compare String property,
'            passing 'K'eyboard, 'C'opy
'
'
'
'---

Option Explicit
Dim moCn As MwfLib.MDOConnection '*** (Bef speed) bef 24/11/06
Dim ms_wLrMCd As String, ms_wLrSrt As String, ms_wLrCmCtg As String
Dim ms_PwdFlag As Boolean
'*** For the Toggle Options of Command Buttons
Enum en_LabRtFra
  Cpy = 0
  Summ = 1
  ImpExc = 2 'Uni.9- Import from Excel
  Pwd = 3 'Uni.24- Password
  DelRt = 5 'Uni.24- Delete rates
End Enum

Private Sub ADC_Load()
  '*** Set the First Key Control property to the working field wLrMCd
  '*** Set the First Non Key Control property to the grid GrdLab
  '*** Set the Child Property Of ADC to GrdLab
  '*** Set the Previous and Next Control Properties for the Grid GrdLab
  '*** Set the Hot Key for the Grid GrdLab as 'A'
  '*** initialise the default values for the working fields wLrMCd and wLrSrt

  Set moCn = ADC.Connection  '*** (Bef speed) bef 24/11/06

  Set ADC.FirKeyCtl = ADC("wLrMCd")
  Set ADC.FirNKeyCtl = GrdLab
  ADC.Child = "GrdLab"
  Set GrdLab.PrevCtl = ADC
  Set GrdLab.NextCtl = ADC
  GrdLab.HotKey = "A"
  
  ms_wLrMCd = "": ms_wLrSrt = "": ms_wLrCmCtg = ""
End Sub
Private Sub ADC_KeyWhen()
  '*** pass 'K' as the parameter for the CmpString procedure. This sets the CmpStr value of
      'the fields in the copy frame as blank and sets the proper cmpstr values for fileds in
      'the key frame (i.e. the range selection fields)
  '*** Disable All command buttons
  '*** Hide all frames except Frame FraNKeyAll
  '*** set the default values for Lab Main Code and Sort Option from the previous entry

  Call CmpString("K")
  Call EnaDisaCmds(True)
  Call HideAllFras
  
  If ms_wLrMCd <> "" Then ADC("wLrMCd") = ms_wLrMCd
  If ms_wLrSrt <> "" Then ADC("wLrSrt") = ms_wLrSrt
  If ms_wLrCmCtg <> "" Then ADC("wLrCmCtg") = ms_wLrCmCtg
  txtPwd.Text = ""  'Uni.24- To set passowrd as blank
  
End Sub
Private Sub ADC_KeyValid(Cancel As Boolean, ErrMsg As String)
  '*** Validating LabMcd, as it is used in Grd Init
  '*** Check whether the LabMCd and the Sort Option is valid
      '(Sort option can be blank in which case the default sorting is on LrCmCd, LrSCd)
  
  Call ATXT_FldChange(0, UCase("wLrMCd"), Cancel, ErrMsg, ADC("wLrMCd"))
  If Cancel = True Then Exit Sub
  
  Call ATXT_FldChange(0, UCase("wLrCmCtg"), Cancel, ErrMsg, ADC("wLrCmCtg"))
  If Cancel = True Then Exit Sub
  
  Cancel = ADC("wLrSrt") <> "" And _
          (Not moCn.RecSeek("Select PMCd From Param Where PTyp = 'SRTLR' " + _
                       "And PMCd = '" + ADC("wLrSrt") + "'"))
  If Cancel = True Then ErrMsg = "Invalid Sort Option": Exit Sub
  
'Indigo.1- Checking user access permission for the given customers
  Dim wErrMsg As String
  If Not ChkCmpUsrAccess(ADC("WLRCMCTG"), ADC.MenuCd, ADC("WLRCMCDFR"), ADC("WLRCMCDTO"), "", wErrMsg) Then
    Cancel = True: ErrMsg = "You Don't Have Permissions To View The Details of some Customers in given Range, example:" + wErrMsg: Exit Sub
  End If
  
End Sub
Private Sub ADC_NonKeyWhen(pr_Cancel As Boolean, pr_ErrMsg As String)
  '*** Enable all the command buttons
  '*** Store the Lab Main Code, Sort Option for the next entry
  
  Call EnaDisaCmds(False)
  ms_wLrMCd = ADC("wLrMCd")
  ms_wLrSrt = ADC("wLrSrt")
  ms_wLrCmCtg = ADC("wLrCmCtg")
      
End Sub
Private Sub atxt_FldWhen(Index As Integer, ByVal IdName As String, Cancel As Boolean, ErrMsg As String)
  '*** Set Helps For Each Field
  Select Case IdName
  Case Is = UCase("wLrMCd")
    Call HlpList.PMCd("LABMCD")
  Case Is = UCase("wLrCmCtg")
    Call HlpList.PMCd("CMCTG")
  Case Is = UCase("wLrCmCdFr")
    Call HlpList.CustCd(ADC("wLrCmCtg"), False) '*** Jay 2.14 ***
  Case Is = UCase("wLrCmCdTo")
    Call HlpList.CustCd(ADC("wLrCmCtg"), False) '*** Jay 2.14 ***
  Case Is = UCase("wLrSrt")
    Call HlpList.PMCd("SRTLR")
  
' fields of Copy frame
  Case Is = UCase("wCmCtgFr")
    Call HlpList.PMCd("CMCTG")
  Case Is = UCase("wCustFr")
    Call HlpList.CustCd(ADC("wCmCtgFr"), False) '*** Jay 2.14 ***
  Case Is = UCase("wCustTo")
    Call HlpList.CustCd(ADC("wLrCmCtg"))
  Case Is = UCase("wSCdFr")
    Call HlpList.PSCd("LABSCD", ADC("wLrMCd"))
  Case Is = UCase("wSCdTo")
    Call HlpList.PSCd("LABSCD", ADC("wLrMCd"))
  ' Zubin 213
  Case Is = UCase("wCpyCurCdFr")
    If UCase(ADC("wCustFr")) = ctSelfCmCd Or UCase(ADC("wCustFr")) = ctValueCmCd Then
      Call HlpList.PMCd("CURNCY")
    Else
      Cancel = True
    End If
  Case Is = UCase("wCpyCurCdTo")
    If UCase(ADC("wCustTo")) = UCase(ctSelfCmCd) Or UCase(ADC("wCustTo")) = UCase(ctValueCmCd) Then
      Call HlpList.PMCd("CURNCY")
    Else
      Cancel = True
    End If
  Case Is = UCase("wLrCmCurCdFr"), UCase("wLrCmCurCdTo")
    Call HlpList.PMCd("CURNCY")
  ' Zubin 213
  End Select
End Sub
Private Sub ATXT_FldChange(Index As Integer, ByVal IdName As String, Cancel As Boolean, ErrMsg As String, ByVal pv_NewValue As Variant)
  '*** wLrMCd       valid PMCd('LABMCD')
  '*** wLrSrt       valid PMCd('SRTLR') or ''
  '*** wCustFr      valid CmCd from CustMst(CmCtg= adc(wCmCtgFr))
  '*** wCustTo      valid CmCd from CustMst(CmCtg= adc(wLrCmCtg))
  
  Select Case IdName
  Case Is = UCase("wLrMCd")
    Cancel = Not moCn.RecSeek("Select PMCd From Param Where PTyp = 'LABMCD' " + _
                         "And PMCd = '" + pv_NewValue + "'")
    If Cancel = True Then ErrMsg = "Invalid Labour Main Code": Exit Sub
  Case Is = UCase("wLrCmCtg")
    Cancel = UCase(pv_NewValue) <> "C" And UCase(pv_NewValue) <> "P" And UCase(pv_NewValue) <> "S" And UCase(pv_NewValue) <> "T"
    ErrMsg = "Invalid Customer Category": Exit Sub
  Case Is = UCase("wLrSrt")
      Cancel = pv_NewValue <> "" And (Not moCn.RecSeek("Select PMCd From Param Where " + _
                                                    "PTyp = 'SRTLR' " + _
                                                    "And PMCd = '" + pv_NewValue + "'"))
      If Cancel = True Then ErrMsg = "Invalid Sort Option": Exit Sub
  Case Is = UCase("wCmCtgFr")
    Cancel = pv_NewValue <> "" And Not moCn.RecSeek("Select PmCd From Param Where PTyp = 'CMCTG' and " + _
                         "PmCd = '" + pv_NewValue + "'")
    If Cancel = True Then ErrMsg = "Invalid Source Customer/ Price List Code": Exit Sub
  Case Is = UCase("wCustFr")
    Cancel = pv_NewValue <> "" And Not moCn.RecSeek("Select CmCd From CustMst Where CmCtg = '" + ADC("wCmCtgFr") + "' and " + _
                         "CmCd = '" + pv_NewValue + "'")
    If Cancel = True Then ErrMsg = "Invalid Source Customer/ Price List Code": Exit Sub
  Case Is = UCase("wCustTo")
    Cancel = pv_NewValue <> "" And Not moCn.RecSeek("Select CmCd From CustMst Where " + _
                         "CmCtg = '" + ADC("wLrCmCtg") + "' and CmCd = '" + pv_NewValue + "'")
    If Cancel = True Then ErrMsg = "Invalid Target Customer/ Price List Code": Exit Sub
  ' Zubin 213
  Case Is = UCase("wCpyCurCdFr")
    If UCase(ADC("wCustFr")) = ctSelfCmCd Or UCase(ADC("wCustFr")) = ctValueCmCd Then
      Cancel = pv_NewValue <> "" And Not moCn.RecSeek("Select 'x' from Param where PTyp= 'CURNCY' and PMCd= '" + pv_NewValue + "' and " + _
              "PSCd= ''")
      If Cancel = True Then ErrMsg = "Invalid Currency Code": Exit Sub
    Else
      Cancel = pv_NewValue <> "" And Not moCn.RecSeek("Select 'x' from CustMst where CmCtg= '" + ADC("wCmCtgFr") + "' and " + _
              "CmCd= '" + ADC("wCustFr") + "' and CmCurCd= '" + pv_NewValue + "' ")
      If Cancel = True Then ErrMsg = "Invalid Currency Code": Exit Sub
    End If
  Case Is = UCase("wCpyCurCdTo")
    If UCase(ADC("wCustTo")) = ctSelfCmCd Or UCase(ADC("wCustTo")) = ctValueCmCd Then
      Cancel = pv_NewValue <> "" And Not moCn.RecSeek("Select 'x' from Param where PTyp= 'CURNCY' and PMCd= '" + pv_NewValue + "' and " + _
              "PSCd= ''")
      If Cancel = True Then ErrMsg = "Invalid Currency Code": Exit Sub
    Else
      Cancel = pv_NewValue <> "" And Not moCn.RecSeek("Select 'x' from CustMst where CmCtg= '" + ADC("wLrCmCtg") + "' and " + _
              "CmCd= '" + ADC("wCustTo") + "' and CmCurCd= '" + pv_NewValue + "' ")
      If Cancel = True Then ErrMsg = "Invalid Currency Code": Exit Sub
    End If
  ' Zubin 213
  End Select
End Sub

Private Sub atxt_ReCalculate(Index As Integer, ByVal IdName As String, ByVal OldVal As Variant, ByVal RowNum As Integer, ByVal SourceName As String, ByVal SourceColName As String)
  ' Zubin 213
  Select Case UCase(IdName)
  Case Is = UCase("wCpyCurCdFr")
    If ADC.Mode = xNorm Then
      ADC("wCpyCurCdFr") = moCn.GetFldVal("Select CmCurCd from CustMst where CmCtg='" + ADC("wCmCtgFr") + "' and " + _
                                   "CmCd='" + ADC("wCustFr") + "'")
    End If
  Case Is = UCase("wCpyCurCdTo")
    If ADC.Mode = xNorm Then
      ADC("wCpyCurCdTo") = moCn.GetFldVal("Select CmCurCd from CustMst where CmCtg='" + ADC("wLrCmCtg") + "' and " + _
                                   "CmCd='" + ADC("wCustTo") + "'")
    End If
  Case Is = UCase("wCnvRtFr")
    If ADC.Mode = xNorm Then ADC("wCnvRtFr") = CnvRt(ADC("wCpyCurCdFr"), ADC("wCpyCurCdTo"))
  Case Is = UCase("wMulDivFr")
    If ADC.Mode = xNorm Then ADC("wMulDivFr") = MulDiv(ADC("wCpyCurCdFr"), ADC("wCpyCurCdTo"))
  End Select
  ' Zubin 213
End Sub

Private Sub GrdLab_RowWhen(ByVal RowNum As Integer)
  With GrdLab
      If ADC("wLrMCd") = "GHC" Then
          GrdLab.Value(RowNum, "LrFrWt") = 0
          GrdLab.Value(RowNum, "LrToWt") = 9999.999
      End If
  End With
End Sub
Private Sub GrdLab_ColWhen(ByVal RowNum As Integer, ByVal ColName As String, Cancel As Boolean, ErrMsg As String, ByVal Action As MwfCtl.en_fgAction)
  '*** Set Helps For Each Field
  '*** LrCstRt can be entered only if LrCmCd= ctSelfCmCd
  
  Select Case ColName
  Case Is = UCase("LrCmCd")
    Call HlpList.CustCd(ADC("wLrCmCtg"))
  Case Is = UCase("LrCmCurCd")
    ' Zubin 213 (ctValueCmCd added)
    If UCase(GrdLab.Value(RowNum, "LrCmCd")) = "ZSELF" Or UCase(GrdLab.Value(RowNum, "LrCmCd")) = ctValueCmCd Then
      Call HlpList.PMCd("CURNCY")
    Else
      Cancel = True
    End If
  Case Is = UCase("LrSCd")
    Call HlpList.PSCd("LABSCD", ADC("wLrMCd"))
  Case Is = UCase("LrQW")
    '*** (Jen 2.12)
    If ADC("wLrMCd") = ctBrkLMCd And GrdLab.Value(RowNum, "LrSCd") = ctBrkLSCd Then
      Cancel = True: ErrMsg = "Cannot Edit Labour Q/W For '" & ADC("wLrMCd") + "'": Exit Sub
    End If
    '*** (Jen 2.12)
    Call HlpList.PMCd("LabQw")
  Case Is = UCase("LrCstRt")
          
'      If GrdLab.Value(RowNum, "LrCmCd") <> ctSelfCmCd Then _
'         Cancel = True: ErrMsg = "Cost Rate Cannot Be Entered For This Customer": Exit Sub
     
  '*** Jay 3.01 *** [Min Max Val Should not Enter when LabQw= 'Q']
  Case Is = UCase("LrSalMax")
      If GrdLab.Value(RowNum, "LrQw") = "Q" Then _
        Cancel = True: ErrMsg = "Max Sales Rate Should Not Be Entered When Lab Q/W is Q": Exit Sub
      
  Case Is = UCase("LrSalMin")
      If GrdLab.Value(RowNum, "LrQw") = "Q" Then _
        Cancel = True: ErrMsg = "Min Sales Rate Should Not Be Entered When Lab Q/W is Q": Exit Sub
  '*** Jay 3.01 *** [Min Max Val Should not Enter when LabQw= 'Q']

 '********************Geeta*****Emr206**02/04/04*********
 '***************From Wt & To wt is not specified for LMCd='SET and  "LrQW"='Q'
' **** Zubin 211 **** '
 Case Is = UCase("LrFrWt")
      If ADC("wLrMCd") = "GHC" Then _
         Cancel = True: ErrMsg = "From Weight Cannot Be Entered for GHC": Exit Sub
  Case Is = UCase("LrToWt")
      If ADC("wLrMCd") = "GHC" Then _
         Cancel = True: ErrMsg = "To Weight Cannot Be Entered for GHC": Exit Sub
' **** Zubin 211 **** '
'******************************
 End Select
End Sub
Private Sub GrdLab_ReCalc(ByVal RowNum As Integer, ByVal ColName As String, ByVal OldVal As Variant, ByVal SourceName As String, ByVal SourceColName As String)
  '*** LrCstRt on LrCmCd         Normal
      'If LrCmCd<> ctSelfCmCd then 0
  '*** LrQW on wLrMCd, LrSCd    Normal, Copy        'Init,
      'PValue('LABSCD', adc(wLrMCd), LrSCd)
  
  With GrdLab
    Select Case ColName
    Case Is = UCase("LrCmCurCd")
      If .Mode = fgmnorm Then
        ' Zubin 213 (ctValueCmCd added)
        If .Value(RowNum, "LrCmCd") <> "ZSELF" And .Value(RowNum, "LrCmCd") <> ctValueCmCd Then _
           GrdLab.Value(RowNum, "LrCmCurCd") = moCn.GetFldVal("select CmCurCd from CustMst where " + _
                                              "CmCtg= '" + ADC("wLrCmCtg") + "' " + _
                                              "and CmCd= '" + .Value(RowNum, "LrCmCd") + " ' ")
      End If
    Case Is = UCase("LrQW")
      If .Mode = fgmnorm Or .Mode = fgmCopy Then
        .Value(RowNum, "LrQW") = moCn.GetFldVal("Select PValue from Param where " + _
                                 "PTyp='LABSCD' and PMCd= '" + ADC("wLrMCd") + "' and " + _
                                 "PSCd='" + .Value(RowNum, "LrSCd") + "'")
        End If
'***** urmila LrRt On OdGldAsWt and LrQw ****
    Case Is = UCase("LrFrWt")
      If .Mode = fgmnorm Or .Mode = fgmCopy Then
         If (.Value(RowNum, "LrQW") = "Q") Then .Value(RowNum, "LrFrWt") = 0#
      End If
    Case Is = UCase("LrToWt")
      If .Mode = fgmnorm Or .Mode = fgmCopy Then
         If (.Value(RowNum, "LrQW") = "Q") Then .Value(RowNum, "LrToWt") = 0#
      End If
'***************
    Case Is = UCase("LrCstRt")
'      If .Mode = fgmnorm Then
 '       If .Value(RowNum, "LrCmCd") <> ctSelfCmCd Then _
 '          .Value(RowNum, "LrCstRt") = 0
 '     End If
    '*** Jay 3.01 *** [Min Max Val Should not Enter when LabQw= 'Q']
    Case Is = UCase("LrSalMax")
      If .Mode = fgmnorm Or .Mode = fgmCopy Then
        If .Value(RowNum, "LrQw") = "Q" Then .Value(RowNum, "LrSalMax") = 0#
      End If
    Case Is = UCase("LrSalMin")
      If .Mode = fgmnorm Or .Mode = fgmCopy Then
        If .Value(RowNum, "LrQw") = "Q" Then .Value(RowNum, "LrSalMin") = 0#
      End If
    End Select
    '*** Jay 3.01 *** [Min Max Val Should not Enter when LabQw= 'Q']
  End With
End Sub
Private Sub GrdLab_ColChange(ByVal RowNum As Integer, ByVal ColName As String, ByVal NewValue As Variant, Cancel As Boolean, ErrMsg As String)
  '*** LrCmCd       valid CmCd from CustMst(CmCtg= adc(wLrCmCtg))
  '*** LrSCd        valid PSCd('LABSCD', adc(wLrMCd))
  
  With GrdLab
    Select Case ColName
    Case Is = UCase("LrCmCd")
      Cancel = Not moCn.RecSeek("Select CmCd From CustMst where CmCtg= '" + ADC("wLrCmCtg") + "' " + _
                           "and CmCd = '" + NewValue + "'")
      If Cancel = True Then ErrMsg = "Invalid Customer Code": Exit Sub
      'Indigo.1- Checking user access permission for the given customers
      Dim wErrMsg As String
      If Not ChkCmpUsrAccess(ADC("wLrCmCtg"), ADC.MenuCd, "", "", NewValue, wErrMsg) Then
        Cancel = True: ErrMsg = "You don't have permissions To add given Customer": Exit Sub
      End If
      
    Case Is = UCase("LrCmCurCd")
     Cancel = (NewValue <> "") And (Not moCn.RecSeek("Select PMCd From Param where " + _
                           "PTyp= 'CURNCY' and PMCd = '" + NewValue + "' and " + _
                           "PSCd = '' "))
      ErrMsg = "Invalid Customer Currency Code": Exit Sub
    Case Is = UCase("LrSCd")
      Cancel = Not moCn.RecSeek("Select PSCd From Param where PTyp= 'LABSCD' " + _
                           "and PMCd = '" + ADC("wLrMCd") + "' and " + _
                           "PSCd = '" + NewValue + "'")
      ErrMsg = "Invalid Labour Sub Category": Exit Sub
    Case Is = UCase("LrQW")
      Cancel = Not moCn.RecSeek("Select PmCd From Param where PTyp= 'LABQW' " + _
                           "and PmCd = '" + NewValue + "'")
      ErrMsg = "Invalid Labour Q/W": Exit Sub
    Case Is = UCase("LrCstRt")
'      If GrdLab.Value(RowNum, "LrCmCd") = ctSelfCmCd And NewValue <= 0 Then _
        Cancel = True: ErrMsg = "Cost Rate Should Be > 0": Exit Sub
    Case Is = UCase("LrSalRt")
'      If NewValue <= 0 Then _
        Cancel = True: ErrMsg = "Sales Rate Should Be > 0": Exit Sub
    Case Is = UCase("LrSalMax")
      If GrdLab.Value(RowNum, "LrQw") = "Q" And NewValue <> 0 Then
        Cancel = True: ErrMsg = "Max Sales Rate Should Not Be Entered When Lab Q/W is Q": Exit Sub
      End If
'      If NewValue < GrdLab.Value(RowNum, "LrSalMin") Then _
        Cancel = True: ErrMsg = "Max Sales Rate Should Be > Min Sales Rate": Exit Sub
    End Select
  End With
End Sub
Private Sub GrdLab_InitKey(ByVal RowNum As Integer)
  '*** Initialise Key Fields
  With GrdLab
    .Store "LrMCd", ADC("wLrMCd")
    .Store "LrCmCtg", ADC("wLrCmCtg")
    
    .Store "LrPrtKey", ctCurrPrtn      '****** Sachin 3.02.0 - Default value for Partition Key
    
  End With
End Sub
Private Sub ADC_SetRecSource()
  '*** Since this is an 'A' Entry, there is no RecSource associated with ADC. So set the
      'DataMode property of ADC control as xdatUnbound

  ADC.DataMode = xdatUnbound
End Sub
Private Sub GrdLab_SetRecSource()
  '*** Set the Record Source of the Grid GrdLab
      'Default sort option is LrCmCd, LrSCd
  
  Dim ws_Cnd  As String, ws_OrdBy As String
  ws_Cnd = ADC.RepCond
  ws_Cnd = IIF(ws_Cnd <> "", " where " + ws_Cnd, "")
  If ADC("wLrSrt") = "" Then
    ws_OrdBy = " Order By LrCmCd, LrSCd "
  Else
    ws_OrdBy = "Order By " + moCn.GetFldVal("Select PDesc225 from Param " + _
             "where PTyp='SRTLR' and PMCd='" + ADC("wLrSrt") + "'")
  End If
  GrdLab.RecSource = " Select * from LabRt " + ws_Cnd + ws_OrdBy
End Sub
Private Sub CmdCpy_Click()
  '*** Display or hide the Copy frame depending on the toggle satate
  Call DispFra(Cpy)
  'Manoj 2.10.0
  If FraCpy.Visible = True Then
    ADC("WCMCTGFR") = ADC("WLRCMCTG")
    If UCase$(ADC("WLRCMCTG")) = "T" Then
      ADC("WCMCTGFR").Enabled = False
    Else
      ADC("WCMCTGFR").Enabled = True
    End If
  End If
  'Manoj 2.10.0
End Sub
Private Sub CmdCpyGo_Click()
  '*** Call the Copy Routine and Close the Copy Frame
  If FraCpy.Visible = True Then Call CpyLabRt
  Call CmdCpy_Click
End Sub
Private Sub CpyLabRt()
  '*** Copy Option will copy the Labour rates from one customer to another customer
      'for the given Lab Main Code and for the given Lab Sub Code range
  '*** pass 'C' as the parameter for the CmpString procedure. This will set the values for the
      'CmpStr option of the key frame fields as blank and set proper values for the CmpStr
      'option of the fields in the copy frame
  '*** Cannot copy if Record exists for the given Lab Main Code, Sub Code range option

  Dim ws_LabCnd As String, wn_Row As Single, ws_TgtCurCd As String
  ' Zubin 213
  Dim ws_SrcMulDiv As String
  
  Call CmpString("C")
  
  'Manoj 2.10.0
  ' ***** Manali Trading Module   - 'S' added below
  If ADC("wLrCmCtg") <> "T" Then
    If ADC("wCmCtgFr") <> "C" And ADC("wCmCtgFr") <> "P" And ADC("wCmCtgFr") <> "S" Then DispMsg "Enter Proper Source Customer Category", etError: Exit Sub
  Else
    If ADC("wCmCtgFr") <> "T" Then DispMsg "Enter Proper Source Customer Category", etError: Exit Sub
  End If
  'Manoj 2.10.0
  
  If Not moCn.RecSeek("Select CmCd From CustMst where CmCtg= '" + ADC("wCmCtgFr") + "' " + _
     "and CmCd = '" + ADC("wCustFr") + "'") Then _
     DispMsg "Enter Proper Source Customer/ Price List Code", etError: Exit Sub
  
  If Not moCn.RecSeek("Select CmCd From CustMst where CmCtg= '" + ADC("wLrCmCtg") + "' " + _
     "and CmCd = '" + ADC("wCustTo") + "'") Then _
     DispMsg "Enter Proper Target Customer/ Price List Code", etError: Exit Sub
  
  ws_LabCnd = ADC.RepCond
  ws_LabCnd = IIF(ws_LabCnd <> "", " and " + ws_LabCnd, "")
  
  ' Zubin 213
  If ADC("wCpyCurCdFr") = "" Or ADC("wCpyCurCdTo") = "" Then _
    DispMsg "Source-Target Currency Code Cannot Be Blank", etWarning: Exit Sub
  
  If ADC("wMulDivFr") = "" Then _
    DispMsg "Source-Target Currency Conversion Is Not Entered In The System", etWarning: Exit Sub
  
  If ADC("wCnvRtFr") = "0" Then _
    DispMsg "Source-Target Currency Conversion Factor Cannot Be Zero", etWarning: Exit Sub
  ' Zubin 213
  
  ' Zubin 213 (Curr Cd added)
  If moCn.RecSeek("Select 'x' from LabRt where LrMCd='" + ADC("wLrMcd") + "' and " + _
             "LrCmCtg= '" + ADC("wLrCmCtg") + "' and LrCmCd='" + ADC("wCustTo") + "' and LrCmCurCd= '" + ADC("wCpyCurCdTo") + "' " + ws_LabCnd) Then
    DispMsg "Cannot Copy as Lab Rates Exist For The Given Customer, Currency And SubCd Range", etError
    Exit Sub
  End If
  
  ' Zubin 213 (Before 213)
  'ws_TgtCurCd = moCn.GetFldVal("Select CmCurCd From CustMst where CmCtg= '" + ADC("wLrCmCtg") + "' " + _
            "and CmCd = '" + ADC("wCustTo") + "' ")
  ' Zubin 213
  ws_TgtCurCd = ADC("wCpyCurCdTo")
  ' Zubin 213
  
' ###################################  Manoj #### Ver: 2.0.6  ###################################
' Date:   31/03/2004
  ' Zubin 213 (ws_TgtCurCd changed to ADC("wCpyCurCdFr"))
  If moCn.RecSeek("Select * from LabRt where LrMCd='" + ADC("wLrMCd") + "' and " + _
                     "LrCmCtg= '" + ADC("wCmCtgFr") + "' and LrCmCd='" + ADC("wCustFr") + "' " + _
                     "and LrCmCurCd= '" + ADC("wCpyCurCdFr") + "' " + ws_LabCnd + _
             " And Exists (Select 1 from Param Where PTyp = 'LABSCD' AND PMCd = LrMCd And PSCd = LrSCd And PValidYN = 'N')") Then
    DispMsg "Cannot Copy as Some Invalid LabSCd are Present", etError
    Exit Sub
  End If
' ###########################################  Manoj  ###########################################
  
  ' Zubin 213
  ws_SrcMulDiv = MulDivOp(ADC("wMulDivFr"))
  
  '*** Beginning the copy process
  GrdLab.StartCopy
  MWLib.BeginProcess Me, "Copying Labour Rates ..."
  Dim wRsCustLabRt As MwfLib.MDORowSet
  ' Zubin 213 (ws_TgtCurCd changed to ADC("wCpyCurCdFr"))
  Set wRsCustLabRt = moCn.OpenRes("Select * from LabRt where LrMCd='" + ADC("wLrMCd") + "' and " + _
                     "LrCmCtg= '" + ADC("wCmCtgFr") + "' and LrCmCd='" + ADC("wCustFr") + "' " + _
                     "and LrCmCurCd= '" + ADC("wCpyCurCdFr") + "' " + ws_LabCnd + " Order By LrSCd")
  With wRsCustLabRt
    Do While Not (.EOF Or .BOF)
      MWLib.RunningProcess "Copying Lab Rates For Lab SubCtg " + !LrSCd
      GrdLab.AddItem
      wn_Row = GrdLab.Rows - 1
      GrdLab.Value(wn_Row, "LrCmCd") = ADC("wCustTo")
      GrdLab.Value(wn_Row, "LrSCd") = !LrSCd
      GrdLab.Value(wn_Row, "LrQw") = !LrQW
      'If ADC("wCustTo") = ctSelfCmCd Then
        ' Zubin 213
        'GrdLab.Value(wn_Row, "LrCstRt") = !LrCstRt
        GrdLab.Value(wn_Row, "LrCstRt") = IIF(moCn.GetFldVal("Select " + CStr(!LrCstRt) + " " + ws_SrcMulDiv + " " + CStr(ADC("wCnvRtFr"))) > ctLabRtLimit, ctLabRtLimit, moCn.GetFldVal("Select " + CStr(!LrCstRt) + " " + ws_SrcMulDiv + " " + CStr(ADC("wCnvRtFr"))))
      'End If
      ' Zubin 213 (Before 213)
      'GrdLab.Value(wn_Row, "LrSalRt") = !LrSalRt
      'GrdLab.Value(wn_Row, "LrSalMin") = !LrSalMin
      'GrdLab.Value(wn_Row, "LrSalMax") = !LrSalMax
      'GrdLab.Value(wn_Row, "LrCmCurCd") = !LrCmCurCd
      ' Zubin 213
      GrdLab.Value(wn_Row, "LrSalRt") = IIF(moCn.GetFldVal("Select " + CStr(!LrSalRt) + " " + ws_SrcMulDiv + " " + CStr(ADC("wCnvRtFr"))) > ctLabRtLimit, ctLabRtLimit, moCn.GetFldVal("Select " + CStr(!LrSalRt) + " " + ws_SrcMulDiv + " " + CStr(ADC("wCnvRtFr"))))
      GrdLab.Value(wn_Row, "LrSalMin") = IIF(moCn.GetFldVal("Select " + CStr(!LrSalMin) + " " + ws_SrcMulDiv + " " + CStr(ADC("wCnvRtFr"))) > ctLabRtLimit, ctLabRtLimit, moCn.GetFldVal("Select " + CStr(!LrSalMin) + " " + ws_SrcMulDiv + " " + CStr(ADC("wCnvRtFr"))))
      GrdLab.Value(wn_Row, "LrSalMax") = IIF(moCn.GetFldVal("Select " + CStr(!LrSalMax) + " " + ws_SrcMulDiv + " " + CStr(ADC("wCnvRtFr"))) > ctLabRtLimit, ctLabRtLimit, moCn.GetFldVal("Select " + CStr(!LrSalMax) + " " + ws_SrcMulDiv + " " + CStr(ADC("wCnvRtFr"))))
      
      GrdLab.Value(wn_Row, "LrCmCurCd") = ADC("wCpyCurCdTo")
      
      GrdLab.Value(wn_Row, "LrFrWt") = !LrFrWt
      GrdLab.Value(wn_Row, "LrToWt") = !LrToWt
      GrdLab.SaveRec (wn_Row)
      .MoveNext
    Loop
  End With
  MWLib.EndProcess Me
  GrdLab.EndCopy
  '*** End of the copy process
  Set wRsCustLabRt = Nothing
End Sub
Private Sub CmpString(ByVal mTyp As String)
  '*** Sets the CmpStr properties for the fields in the Key Frame and the Copy Frame
  If mTyp = "K" Then
    ADC("wLrMCd").CmpStr = "LrMCd= "
    ADC("wLrCmCtg").CmpStr = "LrCmCtg = "
    ADC("wLrCmCdFr").CmpStr = "LrCmCd>= "
    ADC("wLrCmCdTo").CmpStr = "LrCmCd<= "
    ADC("wSCdFr").CmpStr = ""
    ADC("wSCdTo").CmpStr = ""
    ' Zubin 213
    ADC("wLrCmCurCdFr").CmpStr = "LrCmCurCd>= "
    ADC("wLrCmCurCdTo").CmpStr = "LrCmCurCd<= "
    ' Zubin 213
  ElseIf mTyp = "C" Then
    ADC("wLrMCd").CmpStr = ""
    ADC("wLrCmCtg").CmpStr = ""
    ADC("wLrCmCdFr").CmpStr = ""
    ADC("wLrCmCdTo").CmpStr = ""
    ADC("wSCdFr").CmpStr = "LrSCd>= "
    ADC("wSCdTo").CmpStr = "LrSCd<= "
    ' Zubin 213
    ADC("wLrCmCurCdFr").CmpStr = ""
    ADC("wLrCmCurCdTo").CmpStr = ""
    ' Zubin 213
  End If
End Sub
Private Sub DispFra(ByVal pv_LabRtFra As en_LabRtFra)
  '*** This is the common procedure which is used to toggle between hiding and displaying
      'the frames (i.e. if the frame is visible then make it invisible and viseversa)
      'associated with a command button

  Select Case pv_LabRtFra
  Case Is = Cpy
    If FraCpy.Visible = True Then
      FraCpy.Visible = False
      FraCpy.Enabled = False
      Call EnaDisaCmds(False, CmdCpy)
      CmdCpy.SetFocus
    Else
      FraCpy.Visible = True
      FraCpy.Enabled = True
      If ADC("wCmCtgFr").Enabled = True Then
        ADC("wCmCtgFr").SetFocus
      Else
        ADC("wCustFr").SetFocus
      End If
      FraCpy.ZOrder
      Call EnaDisaCmds(True, CmdCpy)
    End If
  'Uni.9- if the Import excel frame is visible then will be invisible and viseversa
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
  'Uni.24-if password frame is visible then will be invisible and viseversa
  Case Is = Pwd
    If FraPwd.Visible = True Then
      FraPwd.Visible = False
      FraPwd.Enabled = False
      Call EnaDisaCmds(False, CmdPwd)
      CmdPwd.SetFocus
    Else
      FraPwd.Visible = True
      FraPwd.Enabled = True
      txtPwd.SetFocus
      FraPwd.ZOrder
      Call EnaDisaCmds(True, CmdPwd)
    End If
  'Uni.24- if Delete rates frame is visible then will be invisible and viseversa
    Case Is = DelRt
    If FraDelRt.Visible = True Then
      FraDelRt.Visible = False
      FraDelRt.Enabled = False
      Call EnaDisaCmds(False, CmdDelRt)
      CmdDelRt.SetFocus
    Else
      FraDelRt.Visible = True
      FraDelRt.Enabled = True
      CmdDelRtGo.SetFocus
      FraDelRt.ZOrder
      Call EnaDisaCmds(True, CmdDelRt)
    End If
  End Select
End Sub
Private Sub EnaDisaCmds(ByVal pv_ShowFra As Boolean, Optional ByVal pv_CmdBut As MWCTL_BTN1)
  '*** This procedure is used to give the toggle option (enable and disable) to
      'the command buttons
  
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
  CmdCpy.Enabled = Not pv_ShowFra
  CmdImpExc.Enabled = Not pv_ShowFra
  CmdPwd.Enabled = Not pv_ShowFra
  CmdDelRt.Enabled = Not pv_ShowFra
  If Not pv_CmdBut Is Nothing Then pv_CmdBut.Enabled = True
End Sub
Private Sub HideAllFras()
  '*** This procedure hides all the frames other than FraNKeyAll
  
  FraNKeyAll.Visible = True
  FraCpy.Visible = False
  FraCpy.Enabled = False
  FraImpExc.Visible = False 'Uni.9- Import from excel
  FraImpExc.Enabled = False 'Uni.9- Import from excel
  FraPwd.Visible = False  'Uni.24- Seo Password
  FraPwd.Enabled = False  'Uni.24- Seo Password
  FraDelRt.Visible = False 'Uni.24- Delete rates
  FraDelRt.Enabled = False 'Uni.24-Delete rates
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
    
    Set moCn = Nothing    '*** (09/08/05)  '*** (Bef speed) bef 24/11/06
End Sub
Private Sub GrdLab_Validate(Cancel As Boolean)
  Cancel = GrdLab.Validate
End Sub
'Uni.9-Calling visible / invisible stage of Import excel frame
Private Sub CmdImpExc_Click()
  Call DispFra(ImpExc)
  
End Sub
'Uni.9- Getting excel file from selected path
Private Sub CmdImpExcPath_Click()
  CdImpExc.FileName = ""
  CdImpExc.DialogTitle = "Select Lab Rate Excel File "
  CdImpExc.Filter = "Excel Files(*.xls;*.xlsx)|*.xls;*.xlsx"
  CdImpExc.FilterIndex = 1
  CdImpExc.ShowOpen
  ADC("WXLFILE") = CdImpExc.FileName

End Sub
'Uni.9- Checking excel file name, Sheet no., From & To row nos. If all correct then call Import excel funtion
Private Sub CmdImpExcGo_Click()
  If Trim(ADC("WXLFILE")) = "" Then DispMsg "Please Specify a Proper Excel File Name !!", etError: Exit Sub
  If ADC("WXLSHTNO") = 0 Or ADC("WXLFRROW") = 0 Or ADC("WXLTOROW") = 0 Then DispMsg "Please Specify a Proper Sheet No./From Row/To Row !!", etError: Exit Sub

  Call ImportExcel
  Call DispFra(ImpExc)

End Sub
'Uni.9-Store data from excel file to Grid
Private Sub ImportExcel()
On Error GoTo ErrHdlr
    
    Dim wo_rsExcel As New MDORowSet, i As Integer
    Dim wLrCmCd As String, wLrCmCurCd As String, wLrSCd As String, wLrQW As String
            
    Dim wLrFrWt As Double, wLrToWt As Double, wLrCstRt As Double, wLrSalRt As Double
    Dim wLrMinSalVal As Double, wLrMaxSalVal As Double

    'Checks proper excel file format
    If Not OpenExcelFile(ADC("WXLFILE")) Then DispMsg "Improper File Format", etError: Exit Sub
    
    'Checking proper sheet no.
    If ChkSheet(ADC("WXLSHTNO")) = False Then DispMsg "This WorkSheet Number Does Not Exist in the specified Excel File", etError: CloseExcel: Exit Sub
    
    'Creating Tmp file to store data
    moCn.CreateTmpTable "#TmpXls", "Select 0 as qSr, LrCmCd , LrCmCurCd, LrSCd, LrQw, LrFrWt, LrToWt, LrCstRt, LrSalRt, LrSalMin, LrSalMax Into #TmpXls From LabRt Where 1=2"
            
    ADC("WXLFRROW") = IIF(ADC("WXLFRROW") <= 1, 2, ADC("WXLFRROW"))
    For i = ADC("WXLFRROW") To ADC("WXLTOROW")
        'Checking Customer code
        wLrCmCd = Trim(ReadCell(ADC("WXLSHTNO"), i, 1))
        If wLrCmCd = "" Then DispMsg "Cannot Copy as Customer Code cannot be blank [Row :" + CStr(i) + "]", etError: CloseExcel: Exit Sub
                
        If (wLrCmCd <> "") And (Not moCn.RecSeek("Select CmCd From CustMst where CmCtg= '" + ADC("WLRCMCTG") + "' " + _
                             "and CmCd = '" + wLrCmCd + "'")) Then _
                DispMsg "Cannot Copy as Invalid Customer Code [Row :" + CStr(i) + "]", etError: CloseExcel: Exit Sub
        
        'Indigo.1- Checking user access permission for the given customer
        Dim wErrMsg As String
        If Not ChkCmpUsrAccess(ADC("wLrCmCtg"), ADC.MenuCd, "", "", wLrCmCd, wErrMsg) Then _
             DispMsg "You Don't Have Permissions To Add This Customer [Row :" + CStr(i) + "]", etError: CloseExcel: Exit Sub
                
        'Checking Currency code
        wLrCmCurCd = Trim(ReadCell(ADC("WXLSHTNO"), i, 2))
        If wLrCmCurCd = "" Then DispMsg "Cannot Copy as Customer Currency Code cannot be blank [Row :" + CStr(i) + "]", etError: CloseExcel: Exit Sub
                
        If (wLrCmCurCd <> "") And (Not moCn.RecSeek("Select PMCd From Param where  " + _
                                   "PTyp= 'CURNCY' and PMCd = '" + wLrCmCurCd + "' and PSCd = '' ")) Then _
            DispMsg "Cannot Copy as Invalid Currency Code [Row :" + CStr(i) + "]", etError: CloseExcel: Exit Sub
                            
      
        'Checking Lab Sub Code
        wLrSCd = Trim(ReadCell(ADC("WXLSHTNO"), i, 3))
        If wLrSCd = "" Then DispMsg "Cannot Copy as Lab Sub Code cannot be blank [Row :" + CStr(i) + "]", etError: CloseExcel: Exit Sub
        
        If (wLrSCd <> "") And (Not moCn.RecSeek("Select PSCd From Param where PTyp= 'LABSCD' " + _
                               "and PMCd = '" + ADC("WLRMCD") + "' and PSCd = '" + wLrSCd + "'")) Then _
            DispMsg "Cannot Copy as Invalid Lab Sub Code [Row :" + CStr(i) + "]", etError: CloseExcel: Exit Sub
      
        'Checking Lab Q/W
        wLrQW = Trim(ReadCell(ADC("WXLSHTNO"), i, 4))
        If wLrQW = "" Then DispMsg "Cannot Copy as Lab Q/W cannot be blank [Row :" + CStr(i) + "]", etError: CloseExcel: Exit Sub
        
        If (wLrQW <> "") And (Not moCn.RecSeek("Select PMCd From Param where  " + _
                              "PTyp= 'LABQW' and PMCd = '" + wLrQW + "' and PSCd = '' ")) Then _
            DispMsg "Cannot Copy as Invalid Lab Q/W [Row :" + CStr(i) + "]", etError: CloseExcel: Exit Sub
    
                
        'Checking valid Wt,Rates in the Excel File
        If ReadCell(ADC("WXLSHTNO"), i, 5) <> "" And Not IsNumeric(ReadCell(ADC("WXLSHTNO"), i, 5)) Then
            DispMsg "Cannot Copy as From Wt is not Number [Row :" + CStr(i) + "]", etError: CloseExcel: Exit Sub
        End If
        If ReadCell(ADC("WXLSHTNO"), i, 6) <> "" And Not IsNumeric(ReadCell(ADC("WXLSHTNO"), i, 6)) Then
            DispMsg "Cannot Copy as To Wt is not Number [Row :" + CStr(i) + "]", etError: CloseExcel: Exit Sub
        End If
        If ReadCell(ADC("WXLSHTNO"), i, 7) <> "" And Not IsNumeric(ReadCell(ADC("WXLSHTNO"), i, 7)) Then
            DispMsg "Cannot Copy as Cost Rate is not Number [Row :" + CStr(i) + "]", etError: CloseExcel: Exit Sub
        End If
        If ReadCell(ADC("WXLSHTNO"), i, 8) <> "" And Not IsNumeric(ReadCell(ADC("WXLSHTNO"), i, 8)) Then
            DispMsg "Cannot Copy as Sales Rate is not Number [Row :" + CStr(i) + "]", etError: CloseExcel: Exit Sub
        End If
        If ReadCell(ADC("WXLSHTNO"), i, 9) <> "" And Not IsNumeric(ReadCell(ADC("WXLSHTNO"), i, 9)) Then
            DispMsg "Cannot Copy as Min. Sales Value is not Number [Row :" + CStr(i) + "]", etError: CloseExcel: Exit Sub
        End If
        If ReadCell(ADC("WXLSHTNO"), i, 10) <> "" And Not IsNumeric(ReadCell(ADC("WXLSHTNO"), i, 10)) Then
            DispMsg "Cannot Copy as Max. Sales Value is not Number [Row :" + CStr(i) + "]", etError: CloseExcel: Exit Sub
        End If
                    
        wLrFrWt = Val(ReadCell(ADC("WXLSHTNO"), i, 5))
        wLrToWt = Val(ReadCell(ADC("WXLSHTNO"), i, 6))
        wLrCstRt = Val(ReadCell(ADC("WXLSHTNO"), i, 7))
        wLrSalRt = Val(ReadCell(ADC("WXLSHTNO"), i, 8))
        wLrMinSalVal = Val(ReadCell(ADC("WXLSHTNO"), i, 9))
        wLrMaxSalVal = Val(ReadCell(ADC("WXLSHTNO"), i, 10))
                       
        'Checking valid From & To Wt
        If wLrFrWt < 0 Or wLrFrWt > 9999.999 Then _
            DispMsg "Cannot Copy as Invalid From Weight Entered, Should be Between 0 and fr 9999.999 [Row :" + CStr(i) + "]", etError: CloseExcel: Exit Sub
        If wLrToWt < 0 Or wLrToWt > 9999.999 Then _
            DispMsg "Cannot Copy as Invalid To Weight Entered, Should be Between 0 and to 9999.999 [Row :" + CStr(i) + "]", etError: CloseExcel: Exit Sub

        'Checking valid Cost Rate & Sale rate
        'If wLrCstRt <> 0 And wLrCmCd <> ctSelfCmCd Then _
        '    DispMsg "Cannot Copy as Cost Rate, Should not be Entered For This Customer [Row :" + CStr(i) + "]", etError: CloseExcel: Exit Sub
        
        If wLrCstRt < 0 Or wLrCstRt > 999999.99 Then _
            DispMsg "Cannot Copy as Invalid Cost Rate Entered, Should be Between 0 and 999999.99 [Row :" + CStr(i) + "]", etError: CloseExcel: Exit Sub
        If wLrSalRt < 0 Or wLrSalRt > 999999.99 Then _
            DispMsg "Cannot Copy as Invalid Sale Rate Entered, Should be Between 0 and 999999.99 [Row :" + CStr(i) + "]", etError: CloseExcel: Exit Sub

        'If Lab Q/W is Q then can't enter Min/Max sales rate
        If wLrMinSalVal <> 0 And wLrQW = "Q" Then _
            DispMsg "Cannot Copy as Min Sales Rate, Should Not Be Entered When Lab Q/W is Q [Row :" + CStr(i) + "]", etError: CloseExcel: Exit Sub
  
        If wLrMaxSalVal <> 0 And wLrQW = "Q" Then _
            DispMsg "Cannot Copy as Max Sales Rate, Should Not Be Entered When Lab Q/W is Q [Row :" + CStr(i) + "]", etError: CloseExcel: Exit Sub
           
        'Checking proper Min/Max sales rate
        If wLrMinSalVal < 0 Or wLrMinSalVal > 999999.99 Then _
            DispMsg "Cannot Copy as Invalid Min. Sales Rate Entered, Should be Between 0 and 999999.99 [Row :" + CStr(i) + "]", etError: CloseExcel: Exit Sub
        If wLrMaxSalVal < 0 Or wLrMaxSalVal > 999999.99 Then _
            DispMsg "Cannot Copy as Invalid Max. Sales Rate Entered, Should be Between 0 and 999999.99 [Row :" + CStr(i) + "]", etError: CloseExcel: Exit Sub
        
        moCn.Execute ("Insert Into #TmpXls Values (" + CStr(i) + ", '" + wLrCmCd + "', '" + wLrCmCurCd + "', '" + wLrSCd + "'," + _
                        "'" + wLrQW + "', " + Str(wLrFrWt) + ", " + Str(wLrToWt) + ", " + _
                        Str(wLrCstRt) + " , " + Str(wLrSalRt) + " , " + Str(wLrMinSalVal) + " , " + Str(wLrMaxSalVal) + ")")
         
    Next

    CloseExcel
    
    Set wo_rsExcel = moCn.OpenRes("Select * From #TmpXls Order By qSr")
    wo_rsExcel.MoveFirst
    
    Dim wRow As Double
    
    If wo_rsExcel.RecCount > 0 Then
      With GrdLab
        .StartCopy
        MWLib.BeginProcess Me, "Copying Details ..."
        Do While Not wo_rsExcel.EOF
          .AddItem
          wRow = .Rows - 1
          .Row = wRow
                  
          .Value(wRow, "LrCmCd") = wo_rsExcel!LrCmCd
          .Value(wRow, "LrCmCurCd") = wo_rsExcel!LrCmCurCd
          .Value(wRow, "LrSCd") = wo_rsExcel!LrSCd
          .Value(wRow, "LrQw") = wo_rsExcel!LrQW
          .Value(wRow, "LrFrWt") = wo_rsExcel!LrFrWt
          .Value(wRow, "LrToWt") = wo_rsExcel!LrToWt
          .Value(wRow, "LrCstRt") = wo_rsExcel!LrCstRt
          .Value(wRow, "LrSalRt") = wo_rsExcel!LrSalRt
          .Value(wRow, "LrSalMin") = wo_rsExcel!LrSalMin
          .Value(wRow, "LrSalMax") = wo_rsExcel!LrSalMax

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

'Uni.24-When user click Delete rates button system will check the Seo password entered, if not then gives error message
Private Sub CmdDelRt_Click()
  If Not ms_PwdFlag Then DispMsg "SEO Password Required", etError: Exit Sub
  Call DispFra(DelRt)
End Sub
'Uni.24-Calling DelRts function
Private Sub CmdDelRtGo_Click()
  Call CmdDelRt_Click
  Call DelRts
End Sub
'Uni.24- Deletes all the rows in the Grid
Private Sub DelRts()
  Dim ws_Cnd As String, wsql As String
  
  ws_Cnd = ADC.RepCond
  ws_Cnd = IIF(ws_Cnd <> "", " where ", " and ") + ws_Cnd

  wsql = "Delete from LabRt " + ws_Cnd
  moCn.Execute (wsql)
  GrdLab.Rows = 1

End Sub
'Uni.24-Validating Seo Password entered
Private Sub TxtPwd_Validate(Cancel As Boolean)
  If txtPwd.Text = "" Then
      ms_PwdFlag = moCn.RecSeek("Select PValue1 from Param where PTyp= 'SEOPWD' and PmCd= 'SEOPWD' and PValue1= ''")
  Else
      If Trim(txtPwd.Text) = moCn.GetFldVal("select IsNull ((select PValue from Param Join UsrAccess on UaUsrCd = PMCd " + _
                       "where PTyp= 'USR' and PMCd= '" + gs_UsrCd + "' and UaMnuCd = '" + ADC.MenuCd + "' and " + _
                       "UaSeoYn= 'Y' and PValue= '" + Trim(txtPwd.Text) + "'), " + _
                       "(Select PValue1 from Param where Ptyp= 'SEOPWD' and PmCd= 'SEOPWD' )) ") Then
        ms_PwdFlag = True: Exit Sub
      Else
        ms_PwdFlag = False
        Cancel = True: DispMsg "Invalid SEO Password / User Does Not Have SEO Rights", etError
      End If
  End If
End Sub
'Uni.24-Invisible /invisible stage of password frame
Private Sub CmdPwd_Click()
  Call DispFra(Pwd)
End Sub


