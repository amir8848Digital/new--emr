VERSION 5.00
Object = "{AB2CC32A-E652-42A2-B341-26ECE02F6386}#2.1#0"; "mwfctlforspeed9.ocx"
Begin VB.Form EmrSuppFrmUpd 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Emperor Updates"
   ClientHeight    =   9810
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   12165
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   9810
   ScaleWidth      =   12165
   WindowState     =   2  'Maximized
   Begin MwfCtl.MWCTL_HLP MWCTL_HLP1 
      Height          =   435
      Left            =   11130
      TabIndex        =   135
      TabStop         =   0   'False
      Top             =   6510
      Visible         =   0   'False
      Width           =   435
      _ExtentX        =   767
      _ExtentY        =   767
   End
   Begin MwfCtl.MWCTL_ADC ADC 
      Height          =   555
      Left            =   11070
      TabIndex        =   27
      Top             =   9300
      Width           =   5100
      _ExtentX        =   8996
      _ExtentY        =   979
      AllowAdd        =   0   'False
      AllowFind       =   0   'False
      AllowSave       =   0   'False
      AllowDelete     =   0   'False
      BtnWidth        =   885
      BtnHeight       =   435
      BtnMaskColor    =   8454143
      FontSz          =   8.25
   End
   Begin VB.Frame FraMtnPwd 
      Height          =   1125
      Left            =   8190
      TabIndex        =   132
      Top             =   2820
      Visible         =   0   'False
      Width           =   3840
      Begin VB.TextBox TxtMTNPwd 
         Height          =   285
         IMEMode         =   3  'DISABLE
         Left            =   1380
         MaxLength       =   20
         PasswordChar    =   "*"
         TabIndex        =   0
         ToolTipText     =   "Enter Password To Login To Maintenance"
         Top             =   600
         Width           =   1425
      End
      Begin MwfCtl.MWCTL_BTN1 CmdMtnPwdGo 
         Height          =   285
         Left            =   3000
         TabIndex        =   1
         ToolTipText     =   "Press This Button To Delete The Orders In The Scope"
         Top             =   600
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
      Begin VB.Label ALBL 
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
         Index           =   50
         Left            =   210
         TabIndex        =   134
         Top             =   600
         Width           =   1065
      End
      Begin VB.Label LblMain 
         BackStyle       =   0  'Transparent
         BorderStyle     =   1  'Fixed Single
         Caption         =   "  Enter Password For Maintenance"
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
         Index           =   10
         Left            =   -30
         TabIndex        =   133
         Top             =   90
         Width           =   3840
      End
   End
   Begin VB.Frame FraMain 
      BorderStyle     =   0  'None
      Height          =   9975
      Left            =   -60
      TabIndex        =   29
      Top             =   -150
      Width           =   12090
      Begin VB.Frame FraEnaDisaMod 
         Height          =   5505
         Left            =   1920
         TabIndex        =   208
         Top             =   3240
         Visible         =   0   'False
         Width           =   7770
         Begin VB.TextBox TxtModulePwd 
            Height          =   285
            IMEMode         =   3  'DISABLE
            Left            =   1770
            MaxLength       =   20
            PasswordChar    =   "*"
            TabIndex        =   238
            Top             =   5040
            Width           =   1425
         End
         Begin MwfCtl.MWCTL_MED atxt 
            Height          =   285
            Index           =   40
            Left            =   3090
            TabIndex        =   209
            ToolTipText     =   "Enable Flute Bag Module (Y/N)"
            Top             =   480
            Width           =   285
            _ExtentX        =   503
            _ExtentY        =   503
            DataType        =   4
            ReCalcParent    =   "WUCOLDCONAME,WUCNEWCONAME"
            IdName          =   "WFLUTEBAG"
         End
         Begin MwfCtl.MWCTL_MED atxt 
            Height          =   285
            Index           =   27
            Left            =   3090
            TabIndex        =   211
            ToolTipText     =   "Enable Silver Module (Y/N)"
            Top             =   1050
            Width           =   285
            _ExtentX        =   503
            _ExtentY        =   503
            DataType        =   4
            ReCalcParent    =   "WUCOLDCONAME,WUCNEWCONAME"
            IdName          =   "WSLV"
         End
         Begin MwfCtl.MWCTL_MED atxt 
            Height          =   285
            Index           =   28
            Left            =   3090
            TabIndex        =   210
            ToolTipText     =   "Enable Lot No. Module (Y/N)"
            Top             =   765
            Width           =   285
            _ExtentX        =   503
            _ExtentY        =   503
            DataType        =   4
            ReCalcParent    =   "WUCOLDCONAME,WUCNEWCONAME"
            IdName          =   "WLOTNO"
         End
         Begin MwfCtl.MWCTL_MED atxt 
            Height          =   285
            Index           =   29
            Left            =   3090
            TabIndex        =   214
            ToolTipText     =   "Enable Silver Module (Y/N)"
            Top             =   1905
            Width           =   285
            _ExtentX        =   503
            _ExtentY        =   503
            DataType        =   4
            ReCalcParent    =   "WUCOLDCONAME,WUCNEWCONAME"
            IdName          =   "WFGQUOT"
         End
         Begin MwfCtl.MWCTL_BTN1 CmdEnaDisaModGO 
            Height          =   375
            Left            =   6900
            TabIndex        =   239
            ToolTipText     =   "Press This Button To Delete The Orders In The Scope"
            Top             =   4980
            Width           =   585
            _ExtentX        =   1032
            _ExtentY        =   661
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
         Begin MwfCtl.MWCTL_MED atxt 
            Height          =   285
            Index           =   30
            Left            =   3090
            TabIndex        =   212
            ToolTipText     =   "Enable PPC Module (Y/N)"
            Top             =   1335
            Width           =   285
            _ExtentX        =   503
            _ExtentY        =   503
            DataType        =   4
            ReCalcParent    =   "WUCOLDCONAME,WUCNEWCONAME"
            IdName          =   "WPPC"
         End
         Begin MwfCtl.MWCTL_MED atxt 
            Height          =   285
            Index           =   31
            Left            =   3090
            TabIndex        =   213
            ToolTipText     =   "Enable CTB Module (Y/N)"
            Top             =   1620
            Width           =   285
            _ExtentX        =   503
            _ExtentY        =   503
            DataType        =   4
            ReCalcParent    =   "WUCOLDCONAME,WUCNEWCONAME"
            IdName          =   "WCTB"
         End
         Begin MwfCtl.MWCTL_MED atxt 
            Height          =   285
            Index           =   133
            Left            =   3090
            TabIndex        =   218
            ToolTipText     =   "Specify Whether To Show Menu Options Related To Loose RM Invoicing Reports? (Y/N)"
            Top             =   3045
            Width           =   285
            _ExtentX        =   503
            _ExtentY        =   503
            DataType        =   4
            IdName          =   "WLOOSEINV"
         End
         Begin MwfCtl.MWCTL_MED atxt 
            Height          =   285
            Index           =   35
            Left            =   7215
            TabIndex        =   224
            ToolTipText     =   "Specify Whether To Show Menu Options Related To Consolidated Invoicing Reports? (Y/N)"
            Top             =   525
            Width           =   285
            _ExtentX        =   503
            _ExtentY        =   503
            DataType        =   4
            IdName          =   "WCONSOLIDATEDINV"
         End
         Begin MwfCtl.MWCTL_MED atxt 
            Height          =   285
            Index           =   36
            Left            =   7215
            TabIndex        =   225
            ToolTipText     =   "Specify Whether To Automate Invoice Returns?  (Y/N)"
            Top             =   810
            Width           =   285
            _ExtentX        =   503
            _ExtentY        =   503
            DataType        =   4
            IdName          =   "WAUTOMATERETURNS"
         End
         Begin MwfCtl.MWCTL_MED atxt 
            Height          =   285
            Index           =   37
            Left            =   7215
            TabIndex        =   226
            ToolTipText     =   "Specify Whether QW Module to be Enabled?  (Y/N)"
            Top             =   1095
            Width           =   285
            _ExtentX        =   503
            _ExtentY        =   503
            DataType        =   4
            IdName          =   "WQWMOD"
         End
         Begin MwfCtl.MWCTL_MED atxt 
            Height          =   285
            Index           =   140
            Left            =   7215
            TabIndex        =   227
            ToolTipText     =   "Enter 'Y' To Implement Multi Currency Module Or 'N' To Disable It"
            Top             =   1380
            Width           =   285
            _ExtentX        =   503
            _ExtentY        =   503
            DataType        =   4
            IdName          =   "WMULTICUR"
         End
         Begin MwfCtl.MWCTL_MED atxt 
            Height          =   285
            Index           =   43
            Left            =   7215
            TabIndex        =   228
            ToolTipText     =   "Implement Common Locations Module "
            Top             =   1665
            Width           =   285
            _ExtentX        =   503
            _ExtentY        =   503
            DataType        =   4
            IdName          =   "WIMPCOMLOCS"
         End
         Begin MwfCtl.MWCTL_MED atxt 
            Height          =   285
            Index           =   46
            Left            =   7215
            TabIndex        =   229
            ToolTipText     =   "Implement Common Locations Module "
            Top             =   1950
            Width           =   285
            _ExtentX        =   503
            _ExtentY        =   503
            DataType        =   4
            IdName          =   "WINREPMETLSRMCD"
         End
         Begin MwfCtl.MWCTL_MED atxt 
            Height          =   285
            Index           =   48
            Left            =   3090
            TabIndex        =   215
            ToolTipText     =   "Implement Local Invoicing Module "
            Top             =   2190
            Visible         =   0   'False
            Width           =   285
            _ExtentX        =   503
            _ExtentY        =   503
            Enabled         =   0   'False
            DataType        =   4
            IdName          =   "WHLCLINVYN"
         End
         Begin MwfCtl.MWCTL_MED atxt 
            Height          =   285
            Index           =   49
            Left            =   3090
            TabIndex        =   216
            ToolTipText     =   "Implement Fg Tag Printing Module "
            Top             =   2475
            Visible         =   0   'False
            Width           =   285
            _ExtentX        =   503
            _ExtentY        =   503
            Enabled         =   0   'False
            DataType        =   4
            IdName          =   "WHFGTAGMOD"
         End
         Begin MwfCtl.MWCTL_MED atxt 
            Height          =   285
            Index           =   50
            Left            =   7215
            TabIndex        =   230
            ToolTipText     =   "Implement Customise Order Entry Module"
            Top             =   2235
            Width           =   285
            _ExtentX        =   503
            _ExtentY        =   503
            DataType        =   4
            IdName          =   "WHCUSTOMISEORD"
         End
         Begin MwfCtl.MWCTL_MED atxt 
            Height          =   285
            Index           =   51
            Left            =   3090
            TabIndex        =   217
            ToolTipText     =   "Enable Auto Custom Matching Option"
            Top             =   2760
            Width           =   285
            _ExtentX        =   503
            _ExtentY        =   503
            DataType        =   4
            IdName          =   "WHAUTOCUSTMMTCH"
         End
         Begin MwfCtl.MWCTL_MED atxt 
            Height          =   285
            Index           =   54
            Left            =   7215
            TabIndex        =   231
            ToolTipText     =   "Enable Best Use Fields In Custom Purchase For Auto Custom Matching"
            Top             =   2520
            Width           =   285
            _ExtentX        =   503
            _ExtentY        =   503
            DataType        =   4
            IdName          =   "WHBESTUSE"
         End
         Begin MwfCtl.MWCTL_MED atxt 
            Height          =   285
            Index           =   55
            Left            =   7215
            TabIndex        =   232
            ToolTipText     =   "Specify 'Y' if you want to Enable Order Copy from Excel"
            Top             =   2805
            Width           =   285
            _ExtentX        =   503
            _ExtentY        =   503
            DataType        =   4
            IdName          =   "WORDCPYFRMEXCEL"
         End
         Begin MwfCtl.MWCTL_MED atxt 
            Height          =   285
            Index           =   56
            Left            =   7215
            TabIndex        =   233
            ToolTipText     =   "Enable Palladium Module (Y/N)"
            Top             =   3090
            Width           =   285
            _ExtentX        =   503
            _ExtentY        =   503
            DataType        =   4
            ReCalcParent    =   "WUCOLDCONAME,WUCNEWCONAME"
            IdName          =   "WPDMOD"
         End
         Begin MwfCtl.MWCTL_MED atxt 
            Height          =   285
            Index           =   57
            Left            =   7215
            TabIndex        =   234
            ToolTipText     =   "Enable Travel Excel  (Y/N)"
            Top             =   3375
            Width           =   285
            _ExtentX        =   503
            _ExtentY        =   503
            DataType        =   4
            ReCalcParent    =   "WUCOLDCONAME,WUCNEWCONAME"
            IdName          =   "WTRAVELEXL"
         End
         Begin MwfCtl.MWCTL_MED atxt 
            Height          =   285
            Index           =   58
            Left            =   3090
            TabIndex        =   219
            ToolTipText     =   "Enable Finish Goods Transaction (Yes/No)"
            Top             =   3330
            Width           =   285
            _ExtentX        =   503
            _ExtentY        =   503
            DataType        =   4
            IdName          =   "WHFGIOYN"
         End
         Begin MwfCtl.MWCTL_MED atxt 
            Height          =   285
            Index           =   59
            Left            =   3090
            TabIndex        =   220
            ToolTipText     =   "Enable Design Catalogue Modulel  (Y/N)"
            Top             =   3630
            Width           =   285
            _ExtentX        =   503
            _ExtentY        =   503
            DataType        =   4
            ReCalcParent    =   "WUCOLDCONAME,WUCNEWCONAME"
            IdName          =   "WDSGCAT"
         End
         Begin MwfCtl.MWCTL_MED atxt 
            Height          =   285
            Index           =   60
            Left            =   7215
            TabIndex        =   235
            ToolTipText     =   "Enable Order and Invoice Costing Reports in Excel Format"
            Top             =   3660
            Width           =   285
            _ExtentX        =   503
            _ExtentY        =   503
            DataType        =   4
            IdName          =   "WHORDEXCELREPYN"
         End
         Begin MwfCtl.MWCTL_MED atxt 
            Height          =   285
            Index           =   61
            Left            =   3090
            TabIndex        =   221
            ToolTipText     =   "Enable Bag Alias No. Module in Emr (Yes/No)"
            Top             =   3930
            Width           =   285
            _ExtentX        =   503
            _ExtentY        =   503
            DataType        =   4
            IdName          =   "WHRFIDMODYN"
         End
         Begin MwfCtl.MWCTL_MED atxt 
            Height          =   285
            Index           =   62
            Left            =   7215
            TabIndex        =   236
            ToolTipText     =   "Enable Multi Price Quotation  Module in Emr (Yes/No)"
            Top             =   3945
            Width           =   285
            _ExtentX        =   503
            _ExtentY        =   503
            DataType        =   4
            IdName          =   "WHMULTIPRCQTMODYN"
         End
         Begin MwfCtl.MWCTL_MED atxt 
            Height          =   285
            Index           =   63
            Left            =   7215
            TabIndex        =   237
            ToolTipText     =   "Enable Rate from RateChart in Invoice  Module in Emr (Yes/No)"
            Top             =   4245
            Width           =   285
            _ExtentX        =   503
            _ExtentY        =   503
            DataType        =   4
            IdName          =   "WHRTFRRTCHRTYN"
         End
         Begin MwfCtl.MWCTL_MED atxt 
            Height          =   285
            Index           =   64
            Left            =   3090
            TabIndex        =   222
            ToolTipText     =   "Cost Rate Editable Yes/No"
            Top             =   4215
            Width           =   285
            _ExtentX        =   503
            _ExtentY        =   503
            DataType        =   4
            IdName          =   "WHCSTRTEDTYN"
         End
         Begin MwfCtl.MWCTL_MED atxt 
            Height          =   285
            Index           =   65
            Left            =   3090
            TabIndex        =   223
            ToolTipText     =   "Rework Module Yes/No"
            Top             =   4500
            Width           =   285
            _ExtentX        =   503
            _ExtentY        =   503
            DataType        =   4
            IdName          =   "WHRWRKMODYN"
         End
         Begin VB.Label Label21 
            BackStyle       =   0  'Transparent
            Caption         =   "Rework Module"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Left            =   150
            TabIndex        =   271
            Top             =   4500
            Width           =   2895
         End
         Begin VB.Label Label20 
            BackStyle       =   0  'Transparent
            Caption         =   "Cost Rate Editable"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Left            =   150
            TabIndex        =   270
            Top             =   4215
            Width           =   2895
         End
         Begin VB.Label Label19 
            BackStyle       =   0  'Transparent
            Caption         =   "Rate from Rate Chrt in Invoice"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Left            =   3720
            TabIndex        =   269
            Top             =   4260
            Width           =   2895
         End
         Begin VB.Label Label18 
            BackStyle       =   0  'Transparent
            Caption         =   "Multi Price Quotation "
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Left            =   3720
            TabIndex        =   268
            Top             =   3945
            Width           =   2895
         End
         Begin VB.Label Label17 
            BackStyle       =   0  'Transparent
            Caption         =   "Bag Alias No. Module In Emr"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Left            =   150
            TabIndex        =   267
            Top             =   3930
            Width           =   2895
         End
         Begin VB.Label Label15 
            BackStyle       =   0  'Transparent
            Caption         =   "Ord/Inv Printing in Excel Format"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Left            =   3690
            TabIndex        =   265
            Top             =   3660
            Width           =   3735
         End
         Begin VB.Label ALBL 
            BackStyle       =   0  'Transparent
            Caption         =   "CTB"
            BeginProperty Font 
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
            Left            =   150
            TabIndex        =   264
            Top             =   1620
            Width           =   1245
         End
         Begin VB.Label ALBL 
            BackStyle       =   0  'Transparent
            Caption         =   "PPC"
            BeginProperty Font 
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
            Left            =   150
            TabIndex        =   263
            Top             =   1335
            Width           =   1245
         End
         Begin VB.Label ALBL 
            BackStyle       =   0  'Transparent
            Caption         =   "Enter Password"
            BeginProperty Font 
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
            Left            =   120
            TabIndex        =   262
            Top             =   5040
            Width           =   1575
         End
         Begin VB.Label ALBL 
            BackStyle       =   0  'Transparent
            Caption         =   "Lot No."
            BeginProperty Font 
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
            Left            =   150
            TabIndex        =   261
            Top             =   765
            Width           =   1365
         End
         Begin VB.Label ALBL 
            BackStyle       =   0  'Transparent
            Caption         =   "Quotation From FG Bag"
            BeginProperty Font 
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
            Left            =   150
            TabIndex        =   260
            Top             =   1905
            Width           =   2415
         End
         Begin VB.Label ALBL 
            BackStyle       =   0  'Transparent
            Caption         =   "Silver"
            BeginProperty Font 
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
            Left            =   150
            TabIndex        =   259
            Top             =   1050
            Width           =   1245
         End
         Begin VB.Label ALBL 
            BackStyle       =   0  'Transparent
            Caption         =   "Flute Bag"
            BeginProperty Font 
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
            Left            =   150
            TabIndex        =   258
            Top             =   480
            Width           =   1365
         End
         Begin VB.Label LblMain 
            BackStyle       =   0  'Transparent
            BorderStyle     =   1  'Fixed Single
            Caption         =   "  Enable/Disable Modules"
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
            Index           =   9
            Left            =   -30
            TabIndex        =   257
            Top             =   90
            Width           =   7845
         End
         Begin VB.Label ALBL 
            BackStyle       =   0  'Transparent
            Caption         =   "Loose Invoice Menus"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Index           =   115
            Left            =   150
            TabIndex        =   256
            Top             =   3045
            Width           =   2115
         End
         Begin VB.Label ALBL 
            BackStyle       =   0  'Transparent
            Caption         =   "Consolidated Invoice Menus"
            BeginProperty Font 
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
            Left            =   3690
            TabIndex        =   255
            Top             =   525
            Width           =   2775
         End
         Begin VB.Label ALBL 
            BackStyle       =   0  'Transparent
            Caption         =   "Automate Returns"
            BeginProperty Font 
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
            Left            =   3690
            TabIndex        =   254
            Top             =   810
            Width           =   2775
         End
         Begin VB.Label ALBL 
            BackStyle       =   0  'Transparent
            Caption         =   "Q/W For All RM Categories"
            BeginProperty Font 
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
            Left            =   3690
            TabIndex        =   253
            Top             =   1095
            Width           =   2775
         End
         Begin VB.Label Label16 
            BackStyle       =   0  'Transparent
            Caption         =   "Multi Currency"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Left            =   3690
            TabIndex        =   252
            Top             =   1380
            Width           =   2775
         End
         Begin VB.Label Label2 
            BackStyle       =   0  'Transparent
            Caption         =   "Implement Common Locs"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Left            =   3690
            TabIndex        =   251
            Top             =   1665
            Width           =   2775
         End
         Begin VB.Label Label4 
            BackStyle       =   0  'Transparent
            Caption         =   "Inv Rep Met Loss On RmCd"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Left            =   3690
            TabIndex        =   250
            Top             =   1950
            Width           =   2775
         End
         Begin VB.Label Label5 
            BackStyle       =   0  'Transparent
            Caption         =   "Local Invoice"
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
            Height          =   285
            Left            =   150
            TabIndex        =   249
            Top             =   2190
            Visible         =   0   'False
            Width           =   2295
         End
         Begin VB.Label Label6 
            BackStyle       =   0  'Transparent
            Caption         =   "FG Tag Printing"
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
            Height          =   285
            Left            =   150
            TabIndex        =   248
            Top             =   2475
            Visible         =   0   'False
            Width           =   2295
         End
         Begin VB.Label Label9 
            BackStyle       =   0  'Transparent
            Caption         =   "IDE (Customise Order Entry)"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Left            =   3690
            TabIndex        =   247
            Top             =   2235
            Width           =   2775
         End
         Begin VB.Label Label10 
            BackStyle       =   0  'Transparent
            Caption         =   "Enable Auto Custom Matching"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Left            =   150
            TabIndex        =   246
            Top             =   2760
            Width           =   2895
         End
         Begin VB.Label Label13 
            BackStyle       =   0  'Transparent
            Caption         =   "Enable Best Use In Custom Purchase"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Left            =   3690
            TabIndex        =   245
            Top             =   2520
            Width           =   3525
         End
         Begin VB.Label ALBL 
            BackStyle       =   0  'Transparent
            Caption         =   "Enable Order Copy from Excel"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Index           =   66
            Left            =   3690
            TabIndex        =   244
            Top             =   2805
            Width           =   3495
         End
         Begin VB.Label ALBL 
            BackStyle       =   0  'Transparent
            Caption         =   "Palladium"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Index           =   68
            Left            =   3690
            TabIndex        =   243
            Top             =   3090
            Width           =   1245
         End
         Begin VB.Label ALBL 
            BackStyle       =   0  'Transparent
            Caption         =   "Enable Travel Excel"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Index           =   69
            Left            =   3690
            TabIndex        =   242
            Top             =   3375
            Width           =   3525
         End
         Begin VB.Label Label14 
            BackStyle       =   0  'Transparent
            Caption         =   "Finsh Goods Transaction"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Left            =   150
            TabIndex        =   241
            Top             =   3330
            Width           =   2895
         End
         Begin VB.Label ALBL 
            BackStyle       =   0  'Transparent
            Caption         =   "Design Catalogue Module"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Index           =   70
            Left            =   150
            TabIndex        =   240
            Top             =   3660
            Width           =   2985
         End
      End
      Begin VB.Frame FraUpdPrdStg 
         Height          =   1395
         Left            =   2655
         TabIndex        =   56
         Top             =   7230
         Visible         =   0   'False
         Width           =   6300
         Begin VB.TextBox TxtUpdPrdStgPwd 
            Height          =   285
            IMEMode         =   3  'DISABLE
            Left            =   1320
            MaxLength       =   20
            PasswordChar    =   "*"
            TabIndex        =   55
            ToolTipText     =   "Enter Password To Login To Maintenance"
            Top             =   960
            Width           =   1425
         End
         Begin MwfCtl.MWCTL_BTN1 CmdUpdPrdStgGo 
            Height          =   285
            Left            =   5640
            TabIndex        =   57
            ToolTipText     =   "Press This Button To Update Company Code"
            Top             =   930
            Width           =   555
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
         Begin VB.Label ALBL 
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
            Index           =   38
            Left            =   150
            TabIndex        =   137
            Top             =   960
            Width           =   1065
         End
         Begin VB.Label LblPrdStgUpd 
            BackStyle       =   0  'Transparent
            Caption         =   "Before Running This Script Update The LocPrdStg Field In The Location Master Table For LocType = 'R'"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   495
            Index           =   2
            Left            =   150
            TabIndex        =   59
            Top             =   405
            Width           =   6105
         End
         Begin VB.Label LblMain 
            BackStyle       =   0  'Transparent
            BorderStyle     =   1  'Fixed Single
            Caption         =   "  Update Production Stage"
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
            Index           =   8
            Left            =   -30
            TabIndex        =   58
            Top             =   60
            Width           =   6300
         End
      End
      Begin VB.Frame FraCmpCodeIns 
         Height          =   2355
         Left            =   2737
         TabIndex        =   197
         Top             =   6240
         Visible         =   0   'False
         Width           =   6690
         Begin VB.TextBox TxtInsCompPwd 
            Height          =   285
            IMEMode         =   3  'DISABLE
            Left            =   1560
            MaxLength       =   20
            PasswordChar    =   "*"
            TabIndex        =   203
            ToolTipText     =   "Enter Password"
            Top             =   1845
            Width           =   1425
         End
         Begin VB.TextBox TxtCmpCode 
            Height          =   315
            Left            =   1560
            MaxLength       =   3
            TabIndex        =   201
            Top             =   1215
            Width           =   795
         End
         Begin VB.TextBox TxtCmpName 
            Height          =   315
            Left            =   1560
            MaxLength       =   30
            TabIndex        =   202
            Top             =   1530
            Width           =   4395
         End
         Begin VB.CommandButton CmdNewCmpGo 
            Caption         =   "&Go"
            Height          =   375
            Left            =   6030
            TabIndex        =   204
            ToolTipText     =   "Press This Button To Update Database To The Required Version"
            Top             =   1755
            Width           =   525
         End
         Begin MwfCtl.MWCTL_MED atxt 
            Height          =   285
            Index           =   53
            Left            =   1560
            TabIndex        =   200
            ToolTipText     =   "Enter Source Company Code"
            Top             =   930
            Width           =   795
            _ExtentX        =   1402
            _ExtentY        =   503
            MaxLength       =   5
            DataType        =   4
            IdName          =   "WSRCCOCD"
         End
         Begin VB.Label ALBL 
            BackStyle       =   0  'Transparent
            Caption         =   "Enter Password"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Index           =   71
            Left            =   30
            TabIndex        =   266
            Top             =   1845
            Width           =   1485
         End
         Begin VB.Label ALBL 
            BackStyle       =   0  'Transparent
            Caption         =   "Src Cmp. Code"
            BeginProperty Font 
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
            Left            =   60
            TabIndex        =   207
            Top             =   930
            Width           =   1425
         End
         Begin VB.Label ALBL 
            BackStyle       =   0  'Transparent
            Caption         =   "Cmp. Code"
            BeginProperty Font 
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
            Left            =   60
            TabIndex        =   206
            Top             =   1215
            Width           =   1425
         End
         Begin VB.Label ALBL 
            BackStyle       =   0  'Transparent
            Caption         =   "Cmp. Name"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Index           =   65
            Left            =   30
            TabIndex        =   205
            Top             =   1530
            Width           =   1425
         End
         Begin VB.Label ALBL 
            BackStyle       =   0  'Transparent
            Caption         =   "No User Should Be Currently Accessing This Database."
            BeginProperty Font 
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
            Left            =   150
            TabIndex        =   199
            Top             =   450
            Width           =   6255
         End
         Begin VB.Label Label12 
            BackStyle       =   0  'Transparent
            BorderStyle     =   1  'Fixed Single
            Caption         =   "  Insert New Company"
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
            Left            =   0
            TabIndex        =   198
            Top             =   60
            Width           =   6870
         End
      End
      Begin VB.Frame FraUpdCoCd 
         Height          =   2175
         Left            =   1620
         TabIndex        =   118
         Top             =   6450
         Visible         =   0   'False
         Width           =   8370
         Begin VB.TextBox TxtPwd 
            Height          =   285
            IMEMode         =   3  'DISABLE
            Left            =   2760
            MaxLength       =   20
            PasswordChar    =   "*"
            TabIndex        =   123
            ToolTipText     =   "Enter Password"
            Top             =   1680
            Width           =   1425
         End
         Begin MwfCtl.MWCTL_BTN1 CmdUpdCoCdGo 
            Height          =   285
            Left            =   7650
            TabIndex        =   124
            ToolTipText     =   "Press This Button To Update Company Code"
            Top             =   1710
            Width           =   555
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
         Begin MwfCtl.MWCTL_MED atxt 
            Height          =   285
            Index           =   32
            Left            =   3570
            TabIndex        =   120
            ToolTipText     =   "Old Company Name"
            Top             =   1110
            Width           =   4665
            _ExtentX        =   8229
            _ExtentY        =   503
            Enabled         =   0   'False
            MaxLength       =   30
            ReCalcOn        =   "WUCOLDCOCD"
            IdName          =   "WUCOLDCONAME"
         End
         Begin MwfCtl.MWCTL_MED atxt 
            Height          =   285
            Index           =   33
            Left            =   2760
            TabIndex        =   121
            ToolTipText     =   "Enter New Company Code"
            Top             =   1395
            Width           =   705
            _ExtentX        =   1244
            _ExtentY        =   503
            MaxLength       =   3
            DataType        =   4
            IdName          =   "WUCNEWCOCD"
         End
         Begin MwfCtl.MWCTL_MED atxt 
            Height          =   285
            Index           =   34
            Left            =   2760
            TabIndex        =   119
            ToolTipText     =   "Enter Old Company Code"
            Top             =   1110
            Width           =   705
            _ExtentX        =   1244
            _ExtentY        =   503
            MaxLength       =   3
            DataType        =   4
            ReCalcParent    =   "WUCOLDCONAME,WUCNEWCONAME"
            IdName          =   "WUCOLDCOCD"
            Data            =   "WWW"
         End
         Begin MwfCtl.MWCTL_MED atxt 
            Height          =   285
            Index           =   0
            Left            =   3570
            TabIndex        =   122
            ToolTipText     =   "Enter New Company Name"
            Top             =   1395
            Width           =   4665
            _ExtentX        =   8229
            _ExtentY        =   503
            MaxLength       =   30
            ReCalcOn        =   "WUCOLDCOCD"
            IdName          =   "WUCNEWCONAME"
         End
         Begin VB.Label LblUpdCoCd 
            BackStyle       =   0  'Transparent
            Caption         =   "Take A BackUp Of The Current Database. No User Should Be Currently Accessing This Database."
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   495
            Index           =   10
            Left            =   150
            TabIndex        =   131
            Top             =   405
            Width           =   6585
         End
         Begin VB.Label ALBL 
            BackStyle       =   0  'Transparent
            Caption         =   "New Company Cd/ Name"
            BeginProperty Font 
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
            Left            =   150
            TabIndex        =   130
            Top             =   1395
            Width           =   2565
         End
         Begin VB.Label ALBL 
            BackStyle       =   0  'Transparent
            Caption         =   "Old Company Cd/ Name"
            BeginProperty Font 
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
            Left            =   150
            TabIndex        =   129
            Top             =   1110
            Width           =   2565
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
            Left            =   3480
            TabIndex        =   128
            Top             =   1110
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
            Index           =   1
            Left            =   3480
            TabIndex        =   127
            Top             =   1395
            Width           =   135
         End
         Begin VB.Label ALBL 
            BackStyle       =   0  'Transparent
            Caption         =   "Enter Password"
            BeginProperty Font 
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
            Left            =   150
            TabIndex        =   126
            Top             =   1680
            Width           =   2565
         End
         Begin VB.Label LblMain 
            BackStyle       =   0  'Transparent
            BorderStyle     =   1  'Fixed Single
            Caption         =   "  Update Company Code"
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
            TabIndex        =   125
            Top             =   90
            Width           =   8370
         End
      End
      Begin VB.Frame FraCompUpd 
         Height          =   6045
         Left            =   3000
         TabIndex        =   177
         Top             =   2580
         Visible         =   0   'False
         Width           =   5430
         Begin VB.TextBox TxtCompUpdPwd 
            Height          =   285
            IMEMode         =   3  'DISABLE
            Left            =   2760
            MaxLength       =   20
            PasswordChar    =   "*"
            TabIndex        =   22
            ToolTipText     =   "Enter Password"
            Top             =   1635
            Width           =   1425
         End
         Begin MwfCtl.MWCTL_BTN1 CmdCompUpdGo 
            Height          =   375
            Left            =   4470
            TabIndex        =   23
            ToolTipText     =   "Press This Button To Update The System With Component Module To 'Y'"
            Top             =   1530
            Width           =   585
            _ExtentX        =   1032
            _ExtentY        =   661
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
         Begin MwfCtl.MWCTL_MED atxt 
            Height          =   285
            Index           =   38
            Left            =   2760
            TabIndex        =   19
            ToolTipText     =   "Enter 'Y' To Implement Component Bag Module Or 'N' To Disable It (If Bags are open then Can not Disable it)"
            Top             =   780
            Width           =   285
            _ExtentX        =   503
            _ExtentY        =   503
            DataType        =   4
            ReCalcParent    =   "WCOMPUPDCHR"
            IdName          =   "WCOMPBAGS"
         End
         Begin MwfCtl.MWCTL_MED atxt 
            Height          =   285
            Index           =   47
            Left            =   2760
            TabIndex        =   21
            ToolTipText     =   "Enter New Component Character For Default Update"
            Top             =   1365
            Width           =   735
            _ExtentX        =   1296
            _ExtentY        =   503
            MaxLength       =   3
            DataType        =   4
            ReCalcOn        =   "WCOMPBAGS"
            IdName          =   "WCOMPUPDCHR"
         End
         Begin VB.Frame FraCompUpdDispErrGrd 
            Height          =   4095
            Left            =   -30
            TabIndex        =   181
            Top             =   1950
            Visible         =   0   'False
            Width           =   5460
            Begin MwfCtl.MWCTL_FLX GrdCompUpd 
               Height          =   3330
               Left            =   60
               TabIndex        =   24
               ToolTipText     =   "Component Designs"
               Top             =   660
               Width           =   3585
               _ExtentX        =   6324
               _ExtentY        =   5874
               Cols            =   3
               AllowDelete     =   0   'False
               colname1        =   "WCOMPUPDDSGCD"
               heading1        =   "Design Cd                        "
               datatype1       =   4
               maxlength1      =   15
               tooltiptext1    =   "Design Component Code"
               colname2        =   "WCOMPUPDDSGSZ"
               heading2        =   "Size       "
               datatype2       =   4
               maxlength2      =   5
               tooltiptext2    =   "Design Component Size"
            End
            Begin VB.Label LblMain 
               BackStyle       =   0  'Transparent
               BorderStyle     =   1  'Fixed Single
               Caption         =   "  Following Are The List Of Top 300 Design Components With           Missing  Main Metal Details or Production Sequence:"
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   9
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   540
               Index           =   13
               Left            =   0
               TabIndex        =   182
               Top             =   90
               Width           =   5460
            End
         End
         Begin MwfCtl.MWCTL_MED atxt 
            Height          =   285
            Index           =   52
            Left            =   2760
            TabIndex        =   20
            ToolTipText     =   "Enter Maximum No of Comp BChr To Be Opened"
            Top             =   1080
            Width           =   525
            _ExtentX        =   926
            _ExtentY        =   503
            Alignment       =   2
            Mask            =   "#0"
            MaxLength       =   2
            DataType        =   1
            IdName          =   "WBCHRCNT"
         End
         Begin VB.Label Label11 
            BackStyle       =   0  'Transparent
            Caption         =   "Total Comp BChr's Allowed"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Left            =   150
            TabIndex        =   185
            Top             =   1080
            Width           =   2775
         End
         Begin VB.Label LblMain 
            BackStyle       =   0  'Transparent
            Caption         =   "This Update Should Be Performed when No One Else Is Using This Database."
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   480
            Index           =   14
            Left            =   30
            TabIndex        =   184
            Top             =   330
            Width           =   5400
            WordWrap        =   -1  'True
         End
         Begin VB.Label Label1 
            BackStyle       =   0  'Transparent
            Caption         =   "Component Bags"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Left            =   150
            TabIndex        =   183
            Top             =   780
            Width           =   2505
         End
         Begin VB.Label ALBL 
            BackStyle       =   0  'Transparent
            Caption         =   "Enter Password"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Index           =   67
            Left            =   150
            TabIndex        =   180
            Top             =   1635
            Width           =   1575
         End
         Begin VB.Label ALBL 
            BackStyle       =   0  'Transparent
            Caption         =   "Enter New Component Chr"
            BeginProperty Font 
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
            Left            =   150
            TabIndex        =   179
            Top             =   1350
            Width           =   2565
         End
         Begin VB.Label LblMain 
            BackStyle       =   0  'Transparent
            Caption         =   "Design Component Module With Default Updates."
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
            Left            =   30
            TabIndex        =   178
            Top             =   90
            Width           =   7020
            WordWrap        =   -1  'True
         End
      End
      Begin VB.CommandButton CmdInsComp 
         Caption         =   "<&m> Add New Company"
         CausesValidation=   0   'False
         Height          =   600
         Left            =   4890
         Style           =   1  'Graphical
         TabIndex        =   196
         ToolTipText     =   "Update From Version 2.10 Patch 1 To 2.11 Patch 2"
         Top             =   1380
         Width           =   2400
      End
      Begin VB.Frame FraUpdRmQw 
         Height          =   3675
         Left            =   3030
         TabIndex        =   192
         Top             =   4950
         Visible         =   0   'False
         Width           =   5430
         Begin VB.TextBox TxtSuppPwd_UpdRmQw 
            Height          =   285
            IMEMode         =   3  'DISABLE
            Left            =   1770
            PasswordChar    =   "*"
            TabIndex        =   25
            ToolTipText     =   "Enter Password"
            Top             =   3195
            Width           =   2025
         End
         Begin MwfCtl.MWCTL_BTN1 CmdUpdRmQwGo 
            Height          =   375
            Left            =   4470
            TabIndex        =   26
            ToolTipText     =   "Press This Button To Update The System With Component Module To 'Y'"
            Top             =   3180
            Width           =   585
            _ExtentX        =   1032
            _ExtentY        =   661
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
            Caption         =   "Update Rm Qw For Diamonds And Color Stones."
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
            Index           =   18
            Left            =   30
            TabIndex        =   195
            Top             =   90
            Width           =   7020
            WordWrap        =   -1  'True
         End
         Begin VB.Label ALBL 
            BackStyle       =   0  'Transparent
            Caption         =   "Enter Password"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Index           =   60
            Left            =   150
            TabIndex        =   194
            Top             =   3195
            Width           =   1575
         End
         Begin VB.Label LblMain 
            BackStyle       =   0  'Transparent
            Caption         =   $"EmrSuppFrmUpd.frx":0000
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   2610
            Index           =   17
            Left            =   30
            TabIndex        =   193
            Top             =   450
            Width           =   5400
            WordWrap        =   -1  'True
         End
      End
      Begin VB.CommandButton CmdUpdRmQw 
         Caption         =   "<&l> Update Rm Cst And Sal Qw (D And C Only)"
         CausesValidation=   0   'False
         Height          =   600
         Left            =   2490
         Style           =   1  'Graphical
         TabIndex        =   18
         ToolTipText     =   "Option To Update the Production Purity of G or P or S"
         Top             =   1380
         Width           =   2400
      End
      Begin VB.Frame FraInvFgRmMismatch 
         Height          =   4605
         Left            =   2010
         TabIndex        =   186
         Top             =   4020
         Visible         =   0   'False
         Width           =   7680
         Begin MwfCtl.MWCTL_BTN1 CmdInvFgRmMisMatchGo 
            Height          =   285
            Left            =   6930
            TabIndex        =   187
            ToolTipText     =   "Press This Button To Update User Access"
            Top             =   405
            Width           =   555
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
         Begin MwfCtl.MWCTL_FLX GrdInvFgdDelRecs 
            Height          =   3360
            Left            =   30
            TabIndex        =   190
            ToolTipText     =   "Deleted Fg Bags from Invoice"
            Top             =   1200
            Visible         =   0   'False
            Width           =   7635
            _ExtentX        =   13467
            _ExtentY        =   5927
            Cols            =   4
            AllowDelete     =   0   'False
            colname1        =   "QIFNO"
            heading1        =   "InNo                                                 "
            datatype1       =   4
            style1          =   2
            tooltiptext1    =   "Deleted Record from Invoice"
            colname2        =   "QBNO"
            heading2        =   "Bag No                      "
            datatype2       =   4
            style2          =   2
            tooltiptext2    =   "Deleted Bags"
            colname3        =   "QFDNO"
            heading3        =   "Fg Entry                                 "
            datatype3       =   4
            style3          =   2
            tooltiptext3    =   "Fg Entry Record for Bag"
         End
         Begin VB.Label LblTmpInvFgd 
            BackStyle       =   0  'Transparent
            BorderStyle     =   1  'Fixed Single
            Caption         =   "  Deleted Bags From Invoice Which are Not Considered In Invoice"
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
            Left            =   0
            TabIndex        =   191
            Top             =   840
            Visible         =   0   'False
            Width           =   7680
         End
         Begin VB.Label LblMain 
            BackStyle       =   0  'Transparent
            BorderStyle     =   1  'Fixed Single
            Caption         =   "  Repair Invoice and  FG Rm Mismatch"
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
            Index           =   15
            Left            =   -30
            TabIndex        =   189
            Top             =   90
            Width           =   7680
         End
         Begin VB.Label ALBL 
            BackStyle       =   0  'Transparent
            Caption         =   "Please take a backup of the database and ensure Nobody is using this database."
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   495
            Index           =   59
            Left            =   120
            TabIndex        =   188
            Top             =   360
            Width           =   6585
         End
      End
      Begin VB.Frame FraUpdUA 
         Height          =   975
         Left            =   1965
         TabIndex        =   93
         Top             =   7650
         Visible         =   0   'False
         Width           =   7680
         Begin MwfCtl.MWCTL_BTN1 CmdUpdUAGo 
            Height          =   285
            Left            =   6930
            TabIndex        =   94
            ToolTipText     =   "Press This Button To Update User Access"
            Top             =   405
            Width           =   555
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
         Begin VB.Label ALBL 
            BackStyle       =   0  'Transparent
            Caption         =   "Execute This Update after Generating Menus in MDI and Report -> Listing Forms"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   495
            Index           =   13
            Left            =   150
            TabIndex        =   96
            Top             =   405
            Width           =   6585
         End
         Begin VB.Label LblMain 
            BackStyle       =   0  'Transparent
            BorderStyle     =   1  'Fixed Single
            Caption         =   "  Update User Access"
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
            Index           =   7
            Left            =   -30
            TabIndex        =   95
            Top             =   90
            Width           =   7680
         End
      End
      Begin VB.CommandButton CmdInvFgRmMisMatch 
         Caption         =   "<&k> Repair Invoice and Fg Rm Mismatch"
         CausesValidation=   0   'False
         Height          =   600
         Left            =   90
         Style           =   1  'Graphical
         TabIndex        =   17
         ToolTipText     =   "Option To Update the Production Purity of G or P or S"
         Top             =   1380
         Width           =   2400
      End
      Begin VB.Frame FraUpdSYyMm 
         Height          =   1215
         Left            =   2655
         TabIndex        =   89
         Top             =   7410
         Visible         =   0   'False
         Width           =   6300
         Begin MwfCtl.MWCTL_BTN1 CmdUpdSYyMmGo 
            Height          =   285
            Left            =   5640
            TabIndex        =   90
            ToolTipText     =   "Press This Button To Update Company Code"
            Top             =   840
            Width           =   555
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
         Begin VB.Label LblUpdCoCd 
            BackStyle       =   0  'Transparent
            Caption         =   "Take A BackUp Of The Current Database. No User Should Be Currently Accessing This Database."
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   495
            Index           =   1
            Left            =   150
            TabIndex        =   92
            Top             =   405
            Width           =   6105
         End
         Begin VB.Label LblMain 
            BackStyle       =   0  'Transparent
            BorderStyle     =   1  'Fixed Single
            Caption         =   "  Update SYyMm"
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
            Index           =   5
            Left            =   -30
            TabIndex        =   91
            Top             =   90
            Width           =   6300
         End
      End
      Begin VB.Frame FraUpdRmZCd 
         Height          =   1815
         Left            =   2655
         TabIndex        =   101
         Top             =   6810
         Visible         =   0   'False
         Width           =   6300
         Begin MwfCtl.MWCTL_BTN1 CmdUpdRmZCdGo 
            Height          =   285
            Left            =   5040
            TabIndex        =   104
            ToolTipText     =   "Press This Button To Update Company Code"
            Top             =   1395
            Width           =   555
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
         Begin MwfCtl.MWCTL_MED atxt 
            Height          =   285
            Index           =   14
            Left            =   2040
            TabIndex        =   103
            ToolTipText     =   "Enter New Customs Code"
            Top             =   1395
            Width           =   2385
            _ExtentX        =   4207
            _ExtentY        =   503
            MaxLength       =   12
            DataType        =   4
            IdName          =   "WNEWRMZCD"
         End
         Begin MwfCtl.MWCTL_MED atxt 
            Height          =   285
            Index           =   15
            Left            =   2040
            TabIndex        =   102
            ToolTipText     =   "Enter Old Customs Code"
            Top             =   1110
            Width           =   2385
            _ExtentX        =   4207
            _ExtentY        =   503
            MaxLength       =   12
            DataType        =   4
            ReCalcParent    =   "WUCOLDCONAME,WUCNEWCONAME"
            IdName          =   "WOLDRMZCD"
            Data            =   "WWW"
         End
         Begin VB.Label ALBL 
            BackStyle       =   0  'Transparent
            Caption         =   "Old Customs Cd"
            BeginProperty Font 
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
            Left            =   120
            TabIndex        =   108
            Top             =   1110
            Width           =   2205
         End
         Begin VB.Label ALBL 
            BackStyle       =   0  'Transparent
            Caption         =   "New Customs Cd"
            BeginProperty Font 
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
            Left            =   120
            TabIndex        =   107
            Top             =   1395
            Width           =   2205
         End
         Begin VB.Label LblUpdCoCd 
            BackStyle       =   0  'Transparent
            Caption         =   "Take A BackUp Of The Current Database. No User Should Be Currently Accessing This Database."
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   495
            Index           =   0
            Left            =   150
            TabIndex        =   106
            Top             =   405
            Width           =   6105
         End
         Begin VB.Label LblMain 
            BackStyle       =   0  'Transparent
            BorderStyle     =   1  'Fixed Single
            Caption         =   "  Update Customs Code"
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
            Left            =   -30
            TabIndex        =   105
            Top             =   90
            Width           =   6300
         End
      End
      Begin VB.Frame FraUpdPrdPurity 
         Height          =   2415
         Left            =   3990
         TabIndex        =   155
         Top             =   6210
         Visible         =   0   'False
         Width           =   4470
         Begin VB.TextBox TxtPrdPurUpdPwd 
            Height          =   285
            IMEMode         =   3  'DISABLE
            Left            =   1710
            MaxLength       =   20
            PasswordChar    =   "*"
            TabIndex        =   161
            ToolTipText     =   "Enter Password To Login To Maintenance"
            Top             =   1890
            Width           =   1425
         End
         Begin MwfCtl.MWCTL_BTN1 RmPrdPurityUpdGo 
            Height          =   345
            Left            =   3300
            TabIndex        =   162
            ToolTipText     =   "Press This Button To Delete The Orders In The Scope"
            Top             =   1830
            Width           =   615
            _ExtentX        =   1085
            _ExtentY        =   609
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
         Begin MwfCtl.MWCTL_MED atxt 
            Height          =   285
            Index           =   39
            Left            =   1710
            TabIndex        =   158
            ToolTipText     =   "Enter Raw Material Category"
            Top             =   990
            Width           =   495
            _ExtentX        =   873
            _ExtentY        =   503
            DataType        =   4
            IdName          =   "WRMCTGPPUR"
         End
         Begin MwfCtl.MWCTL_MED atxt 
            Height          =   285
            Index           =   41
            Left            =   1710
            TabIndex        =   160
            ToolTipText     =   "Enter Raw Material Category"
            Top             =   1590
            Width           =   1425
            _ExtentX        =   2514
            _ExtentY        =   503
            Mask            =   "###0.0000"
            MaxLength       =   9
            DataType        =   2
            IdName          =   "WRMPRDPURITY"
         End
         Begin MwfCtl.MWCTL_MED atxt 
            Height          =   285
            Index           =   42
            Left            =   1710
            TabIndex        =   159
            ToolTipText     =   "Enter Raw Material Category"
            Top             =   1290
            Width           =   1785
            _ExtentX        =   3149
            _ExtentY        =   503
            MaxLength       =   12
            DataType        =   4
            IdName          =   "WRMCDPPUR"
         End
         Begin VB.Label LblDelOrd 
            BackStyle       =   0  'Transparent
            Caption         =   "No User Should Be Currently Accessing This Database."
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   495
            Index           =   2
            Left            =   90
            TabIndex        =   166
            Top             =   450
            Width           =   4875
         End
         Begin VB.Label ALBL 
            BackStyle       =   0  'Transparent
            Caption         =   "RM Code"
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
            Index           =   54
            Left            =   90
            TabIndex        =   165
            Top             =   1290
            Width           =   885
         End
         Begin VB.Label ALBL 
            BackStyle       =   0  'Transparent
            Caption         =   "New Prd. Purity"
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
            Index           =   53
            Left            =   90
            TabIndex        =   164
            Top             =   1590
            Width           =   1845
         End
         Begin VB.Label ALBL 
            BackStyle       =   0  'Transparent
            Caption         =   "RM Ctg "
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
            Index           =   52
            Left            =   90
            TabIndex        =   163
            Top             =   990
            Width           =   885
         End
         Begin VB.Label LblMain 
            BackStyle       =   0  'Transparent
            BorderStyle     =   1  'Fixed Single
            Caption         =   "  Production Purity Update"
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
            Index           =   11
            Left            =   -30
            TabIndex        =   157
            Top             =   90
            Width           =   5130
         End
         Begin VB.Label ALBL 
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
            Index           =   51
            Left            =   90
            TabIndex        =   156
            Top             =   1890
            Width           =   1065
         End
      End
      Begin VB.Frame FraPurPlt 
         Height          =   1395
         Left            =   1965
         TabIndex        =   109
         Top             =   7230
         Visible         =   0   'False
         Width           =   7680
         Begin VB.TextBox TxtPwdPltUpd 
            Height          =   285
            IMEMode         =   3  'DISABLE
            Left            =   3090
            MaxLength       =   30
            PasswordChar    =   "*"
            TabIndex        =   112
            ToolTipText     =   "Enter SEO Password "
            Top             =   930
            Width           =   1950
         End
         Begin MwfCtl.MWCTL_BTN1 CmdPurPltGo 
            Height          =   285
            Left            =   6930
            TabIndex        =   113
            ToolTipText     =   "Press This Button To Set Pure Platinum Code"
            Top             =   960
            Width           =   555
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
         Begin MwfCtl.MWCTL_MED atxt 
            Height          =   285
            Index           =   12
            Left            =   3090
            TabIndex        =   111
            ToolTipText     =   "Enter Pure Platinum Or Silver Code"
            Top             =   645
            Width           =   2175
            _ExtentX        =   3836
            _ExtentY        =   503
            MaxLength       =   12
            DataType        =   4
            IdName          =   "WRMCD"
         End
         Begin MwfCtl.MWCTL_MED atxt 
            Height          =   285
            Index           =   26
            Left            =   3090
            TabIndex        =   110
            ToolTipText     =   "Enter Raw Material Category"
            Top             =   360
            Width           =   495
            _ExtentX        =   873
            _ExtentY        =   503
            MaxLength       =   12
            DataType        =   4
            IdName          =   "WRMCTG"
         End
         Begin VB.Label ALBL 
            BackStyle       =   0  'Transparent
            Caption         =   "Enter Pure Plat./Silver Code"
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
            Index           =   12
            Left            =   150
            TabIndex        =   117
            Top             =   675
            Width           =   2745
         End
         Begin VB.Label ALBL 
            BackStyle       =   0  'Transparent
            Caption         =   "Enter Support Password"
            BeginProperty Font 
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
            Left            =   150
            TabIndex        =   116
            Top             =   960
            Width           =   2745
         End
         Begin VB.Label LblMain 
            BackStyle       =   0  'Transparent
            BorderStyle     =   1  'Fixed Single
            Caption         =   "  Set Pure Platinum && Silver Code"
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
            Index           =   3
            Left            =   -30
            TabIndex        =   115
            Top             =   90
            Width           =   7680
         End
         Begin VB.Label ALBL 
            BackStyle       =   0  'Transparent
            Caption         =   "Raw Material Category"
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
            Index           =   32
            Left            =   150
            TabIndex        =   114
            Top             =   360
            Width           =   2745
         End
      End
      Begin VB.Frame FraDsgAna 
         Height          =   3975
         Left            =   2610
         TabIndex        =   138
         Top             =   4650
         Visible         =   0   'False
         Width           =   6390
         Begin VB.CommandButton CmdDsgAnaUpdGo 
            Caption         =   "&Go"
            Height          =   495
            Left            =   5610
            TabIndex        =   141
            ToolTipText     =   "Press This Button To Update Database To The Required Version"
            Top             =   3360
            Width           =   645
         End
         Begin VB.TextBox txtseopwd 
            Height          =   315
            IMEMode         =   3  'DISABLE
            Left            =   1320
            MaxLength       =   20
            PasswordChar    =   "*"
            TabIndex        =   140
            Top             =   3420
            Width           =   2115
         End
         Begin MwfCtl.MWCTL_FLX GrdDsgAnlyFlds 
            Height          =   2820
            Left            =   90
            TabIndex        =   139
            Top             =   480
            Width           =   6165
            _ExtentX        =   10874
            _ExtentY        =   4974
            Cols            =   4
            colname1        =   "WDANASR"
            heading1        =   "Fld. Sr."
            datatype1       =   4
            maxlength1      =   3
            style1          =   2
            tooltiptext1    =   "Design Analysis Field"
            colname2        =   "WDAANAFLDDESC"
            heading2        =   "Ana. Fld Desc                      "
            datatype2       =   4
            maxlength2      =   50
            style2          =   2
            tooltiptext2    =   "Analysis Field Description"
            colname3        =   "WDANADEFVAL"
            heading3        =   "Def. Value                     "
            datatype3       =   4
            maxlength3      =   15
            tooltiptext3    =   "Enter Default Analysis Code Value"
         End
         Begin VB.Label Label8 
            BackStyle       =   0  'Transparent
            BorderStyle     =   1  'Fixed Single
            Caption         =   "Design Analysis Update"
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
            Left            =   -30
            TabIndex        =   143
            Top             =   90
            Width           =   7680
         End
         Begin VB.Label ALBL 
            BackStyle       =   0  'Transparent
            Caption         =   "SEO Pwd :"
            BeginProperty Font 
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
            Left            =   120
            TabIndex        =   142
            Top             =   3420
            Width           =   1425
         End
      End
      Begin VB.Frame FraDelOrd 
         Height          =   2295
         Left            =   2820
         TabIndex        =   30
         Top             =   6330
         Visible         =   0   'False
         Width           =   5970
         Begin VB.TextBox TxtDelOrdPwd 
            Height          =   285
            IMEMode         =   3  'DISABLE
            Left            =   1440
            MaxLength       =   20
            PasswordChar    =   "*"
            TabIndex        =   41
            ToolTipText     =   "Enter Password To Login To Maintenance"
            Top             =   1740
            Width           =   1425
         End
         Begin MwfCtl.MWCTL_BTN1 CmdDelOrdGo 
            Height          =   285
            Left            =   5220
            TabIndex        =   42
            ToolTipText     =   "Press This Button To Delete The Orders In The Scope"
            Top             =   1755
            Width           =   555
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
         Begin MwfCtl.MWCTL_MED atxt 
            Height          =   285
            Index           =   1
            Left            =   1440
            TabIndex        =   31
            ToolTipText     =   "Enter From Order Tc"
            Top             =   1170
            Width           =   615
            _ExtentX        =   1085
            _ExtentY        =   503
            MaxLength       =   3
            DataType        =   4
            IdName          =   "WDOOMTCFR"
         End
         Begin MwfCtl.MWCTL_MED atxt 
            Height          =   285
            Index           =   2
            Left            =   2145
            TabIndex        =   32
            ToolTipText     =   "Enter From Order Year"
            Top             =   1170
            Width           =   330
            _ExtentX        =   582
            _ExtentY        =   503
            MaxLength       =   2
            DataType        =   4
            IdName          =   "WDOOMYYFR"
         End
         Begin MwfCtl.MWCTL_MED atxt 
            Height          =   285
            Index           =   3
            Left            =   2580
            TabIndex        =   33
            ToolTipText     =   "Enter From Order Character"
            Top             =   1170
            Width           =   615
            _ExtentX        =   1085
            _ExtentY        =   503
            MaxLength       =   3
            DataType        =   4
            IdName          =   "WDOOMCHRFR"
         End
         Begin MwfCtl.MWCTL_MED atxt 
            Height          =   285
            Index           =   4
            Left            =   3285
            TabIndex        =   34
            ToolTipText     =   "Enter From Order Number"
            Top             =   1170
            Width           =   855
            _ExtentX        =   1508
            _ExtentY        =   503
            Alignment       =   2
            Mask            =   "######0"
            MaxLength       =   7
            DataType        =   1
            IdName          =   "WDOOMNOFR"
         End
         Begin MwfCtl.MWCTL_MED atxt 
            Height          =   285
            Index           =   5
            Left            =   1440
            TabIndex        =   36
            ToolTipText     =   "Enter To Order Tc"
            Top             =   1455
            Width           =   615
            _ExtentX        =   1085
            _ExtentY        =   503
            MaxLength       =   3
            DataType        =   4
            IdName          =   "WDOOMTCTO"
         End
         Begin MwfCtl.MWCTL_MED atxt 
            Height          =   285
            Index           =   6
            Left            =   2145
            TabIndex        =   37
            ToolTipText     =   "Enter To Order Year"
            Top             =   1455
            Width           =   330
            _ExtentX        =   582
            _ExtentY        =   503
            MaxLength       =   2
            DataType        =   4
            IdName          =   "WDOOMYYTO"
         End
         Begin MwfCtl.MWCTL_MED atxt 
            Height          =   285
            Index           =   7
            Left            =   2580
            TabIndex        =   38
            ToolTipText     =   "Enter To Order Character"
            Top             =   1455
            Width           =   615
            _ExtentX        =   1085
            _ExtentY        =   503
            MaxLength       =   3
            DataType        =   4
            IdName          =   "WDOOMCHRTO"
         End
         Begin MwfCtl.MWCTL_MED atxt 
            Height          =   285
            Index           =   9
            Left            =   3285
            TabIndex        =   39
            ToolTipText     =   "Enter To Order Number"
            Top             =   1455
            Width           =   855
            _ExtentX        =   1508
            _ExtentY        =   503
            Alignment       =   2
            Mask            =   "######0"
            MaxLength       =   7
            DataType        =   1
            IdName          =   "WDOOMNOTO"
         End
         Begin MwfCtl.MWCTL_MED atxt 
            Height          =   285
            Index           =   10
            Left            =   4215
            TabIndex        =   35
            ToolTipText     =   "Enter From Order Serial"
            Top             =   1170
            Width           =   855
            _ExtentX        =   1508
            _ExtentY        =   503
            Alignment       =   2
            Mask            =   "###0"
            MaxLength       =   4
            DataType        =   1
            IdName          =   "WDOOMSRFR"
         End
         Begin MwfCtl.MWCTL_MED atxt 
            Height          =   285
            Index           =   11
            Left            =   4215
            TabIndex        =   40
            ToolTipText     =   "Enter To Order Serial"
            Top             =   1455
            Width           =   855
            _ExtentX        =   1508
            _ExtentY        =   503
            Alignment       =   2
            Mask            =   "###0"
            MaxLength       =   4
            DataType        =   1
            IdName          =   "WDOOMSRTO"
         End
         Begin VB.Label ALBL 
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
            Index           =   37
            Left            =   240
            TabIndex        =   136
            Top             =   1740
            Width           =   1065
         End
         Begin VB.Label LblDelOrd 
            BackStyle       =   0  'Transparent
            Caption         =   "Take A BackUp Of The Current Database. No User Should Be Currently Accessing This Database."
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   495
            Index           =   0
            Left            =   150
            TabIndex        =   54
            Top             =   405
            Width           =   5655
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
            Left            =   3210
            TabIndex        =   53
            Top             =   1455
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
            Index           =   3
            Left            =   2490
            TabIndex        =   52
            Top             =   1455
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
            Left            =   2070
            TabIndex        =   51
            Top             =   1455
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
            Index           =   0
            Left            =   3210
            TabIndex        =   50
            Top             =   1170
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
            Index           =   5
            Left            =   2490
            TabIndex        =   49
            Top             =   1170
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
            Index           =   6
            Left            =   2070
            TabIndex        =   48
            Top             =   1170
            Width           =   135
         End
         Begin VB.Label ALBL 
            BackStyle       =   0  'Transparent
            Caption         =   "Order Sr Fr"
            BeginProperty Font 
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
            Left            =   150
            TabIndex        =   47
            Top             =   1170
            Width           =   1245
         End
         Begin VB.Label ALBL 
            BackStyle       =   0  'Transparent
            Caption         =   "               To"
            BeginProperty Font 
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
            Left            =   150
            TabIndex        =   46
            Top             =   1455
            Width           =   1245
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
            Left            =   4140
            TabIndex        =   45
            Top             =   1170
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
            Index           =   10
            Left            =   4140
            TabIndex        =   44
            Top             =   1455
            Width           =   135
         End
         Begin VB.Label LblMain 
            BackStyle       =   0  'Transparent
            BorderStyle     =   1  'Fixed Single
            Caption         =   "  Delete Orders"
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
            TabIndex        =   43
            Top             =   90
            Width           =   5970
         End
      End
      Begin VB.Frame FraDelInv 
         Height          =   2625
         Left            =   2760
         TabIndex        =   82
         Top             =   4080
         Visible         =   0   'False
         Width           =   6090
         Begin MwfCtl.MWCTL_BTN1 CmdDelInvGo 
            Height          =   285
            Left            =   5220
            TabIndex        =   60
            ToolTipText     =   "Press This Button To Delete The Orders In The Scope"
            Top             =   1935
            Width           =   555
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
         Begin MwfCtl.MWCTL_MED atxt 
            Height          =   285
            Index           =   13
            Left            =   1470
            TabIndex        =   62
            ToolTipText     =   "Enter From Order Tc"
            Top             =   1620
            Width           =   615
            _ExtentX        =   1085
            _ExtentY        =   503
            MaxLength       =   3
            DataType        =   4
            IdName          =   "WDOOMTCFR1"
         End
         Begin MwfCtl.MWCTL_MED atxt 
            Height          =   285
            Index           =   16
            Left            =   2175
            TabIndex        =   63
            ToolTipText     =   "Enter From Order Year"
            Top             =   1620
            Width           =   330
            _ExtentX        =   582
            _ExtentY        =   503
            MaxLength       =   2
            DataType        =   4
            IdName          =   "WDOOMYYFR1"
         End
         Begin MwfCtl.MWCTL_MED atxt 
            Height          =   285
            Index           =   17
            Left            =   2610
            TabIndex        =   64
            ToolTipText     =   "Enter From Order Character"
            Top             =   1620
            Width           =   615
            _ExtentX        =   1085
            _ExtentY        =   503
            MaxLength       =   3
            DataType        =   4
            IdName          =   "WDOOMCHRFR1"
         End
         Begin MwfCtl.MWCTL_MED atxt 
            Height          =   285
            Index           =   18
            Left            =   3315
            TabIndex        =   65
            ToolTipText     =   "Enter From Order Number"
            Top             =   1620
            Width           =   855
            _ExtentX        =   1508
            _ExtentY        =   503
            Alignment       =   2
            Mask            =   "######0"
            MaxLength       =   7
            DataType        =   1
            IdName          =   "WDOOMNOFR1"
         End
         Begin MwfCtl.MWCTL_MED atxt 
            Height          =   285
            Index           =   19
            Left            =   1470
            TabIndex        =   67
            ToolTipText     =   "Enter To Order Tc"
            Top             =   1905
            Width           =   615
            _ExtentX        =   1085
            _ExtentY        =   503
            MaxLength       =   3
            DataType        =   4
            IdName          =   "WDOOMTCTO1"
         End
         Begin MwfCtl.MWCTL_MED atxt 
            Height          =   285
            Index           =   20
            Left            =   2175
            TabIndex        =   68
            ToolTipText     =   "Enter To Order Year"
            Top             =   1905
            Width           =   330
            _ExtentX        =   582
            _ExtentY        =   503
            MaxLength       =   2
            DataType        =   4
            IdName          =   "WDOOMYYTO1"
         End
         Begin MwfCtl.MWCTL_MED atxt 
            Height          =   285
            Index           =   21
            Left            =   2610
            TabIndex        =   69
            ToolTipText     =   "Enter To Order Character"
            Top             =   1905
            Width           =   615
            _ExtentX        =   1085
            _ExtentY        =   503
            MaxLength       =   3
            DataType        =   4
            IdName          =   "WDOOMCHRTO1"
         End
         Begin MwfCtl.MWCTL_MED atxt 
            Height          =   285
            Index           =   22
            Left            =   3315
            TabIndex        =   70
            ToolTipText     =   "Enter To Order Number"
            Top             =   1905
            Width           =   855
            _ExtentX        =   1508
            _ExtentY        =   503
            Alignment       =   2
            Mask            =   "######0"
            MaxLength       =   7
            DataType        =   1
            IdName          =   "WDOOMNOTO1"
         End
         Begin MwfCtl.MWCTL_MED atxt 
            Height          =   285
            Index           =   23
            Left            =   4245
            TabIndex        =   66
            ToolTipText     =   "Enter From Order Serial"
            Top             =   1620
            Width           =   855
            _ExtentX        =   1508
            _ExtentY        =   503
            Alignment       =   2
            Mask            =   "###0"
            MaxLength       =   4
            DataType        =   1
            IdName          =   "WDOOMSRFR1"
         End
         Begin MwfCtl.MWCTL_MED atxt 
            Height          =   285
            Index           =   24
            Left            =   4245
            TabIndex        =   71
            ToolTipText     =   "Enter To Order Serial"
            Top             =   1905
            Width           =   855
            _ExtentX        =   1508
            _ExtentY        =   503
            Alignment       =   2
            Mask            =   "###0"
            MaxLength       =   4
            DataType        =   1
            IdName          =   "WDOOMSRTO1"
         End
         Begin MwfCtl.MWCTL_MED atxt 
            Height          =   285
            Index           =   25
            Left            =   1470
            TabIndex        =   61
            ToolTipText     =   "Enter Old Company Code"
            Top             =   1140
            Width           =   705
            _ExtentX        =   1244
            _ExtentY        =   503
            MaxLength       =   3
            DataType        =   4
            ReCalcParent    =   "WUCOLDCONAME,WUCNEWCONAME"
            IdName          =   "WCOCD"
            Data            =   "WWW"
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
            Left            =   4170
            TabIndex        =   85
            Top             =   1905
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
            Index           =   19
            Left            =   4170
            TabIndex        =   83
            Top             =   1620
            Width           =   135
         End
         Begin VB.Label ALBL 
            BackStyle       =   0  'Transparent
            Caption         =   "               To"
            BeginProperty Font 
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
            Left            =   180
            TabIndex        =   81
            Top             =   1905
            Width           =   1245
         End
         Begin VB.Label ALBL 
            BackStyle       =   0  'Transparent
            Caption         =   "Order Sr Fr"
            BeginProperty Font 
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
            Left            =   180
            TabIndex        =   80
            Top             =   1620
            Width           =   1245
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
            Index           =   25
            Left            =   2100
            TabIndex        =   79
            Top             =   1620
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
            Index           =   26
            Left            =   2520
            TabIndex        =   78
            Top             =   1620
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
            Index           =   27
            Left            =   3240
            TabIndex        =   77
            Top             =   1620
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
            Index           =   28
            Left            =   2100
            TabIndex        =   76
            Top             =   1905
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
            Index           =   29
            Left            =   2520
            TabIndex        =   75
            Top             =   1905
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
            Index           =   30
            Left            =   3240
            TabIndex        =   74
            Top             =   1905
            Width           =   135
         End
         Begin VB.Label LblDelOrd 
            BackStyle       =   0  'Transparent
            Caption         =   "Take A BackUp Of The Current Database. No User Should Be Currently Accessing This Database."
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   495
            Index           =   1
            Left            =   150
            TabIndex        =   84
            Top             =   405
            Width           =   5655
         End
         Begin VB.Label ALBL 
            BackStyle       =   0  'Transparent
            Caption         =   "Company Cd"
            BeginProperty Font 
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
            Left            =   150
            TabIndex        =   73
            Top             =   1140
            Width           =   1365
         End
         Begin VB.Label LblMain 
            BackStyle       =   0  'Transparent
            BorderStyle     =   1  'Fixed Single
            Caption         =   "  Delete Invoices"
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
            TabIndex        =   72
            Top             =   90
            Width           =   6090
         End
      End
      Begin VB.Frame FraCoName 
         Height          =   2025
         Left            =   3150
         TabIndex        =   167
         Top             =   6600
         Visible         =   0   'False
         Width           =   6330
         Begin VB.TextBox TxtCoNmSuppPwd 
            Height          =   285
            IMEMode         =   3  'DISABLE
            Left            =   1380
            MaxLength       =   20
            PasswordChar    =   "*"
            TabIndex        =   170
            ToolTipText     =   "Enter Password To Login To Maintenance"
            Top             =   1560
            Width           =   2625
         End
         Begin VB.CommandButton CmdCoNameUpdGo 
            Caption         =   "&Go"
            Height          =   315
            Left            =   5310
            TabIndex        =   172
            ToolTipText     =   "Press This Button To Check Dlls"
            Top             =   1545
            Width           =   735
         End
         Begin MwfCtl.MWCTL_MED atxt 
            Height          =   285
            Index           =   44
            Left            =   1380
            TabIndex        =   169
            ToolTipText     =   "Enter New Company Name"
            Top             =   1260
            Width           =   4695
            _ExtentX        =   8281
            _ExtentY        =   503
            MaxLength       =   45
            ReCalcParent    =   "WUCOLDCONAME,WUCNEWCONAME"
            IdName          =   "WCONAME"
         End
         Begin MwfCtl.MWCTL_MED atxt 
            Height          =   285
            Index           =   45
            Left            =   1380
            TabIndex        =   168
            ToolTipText     =   "Enter New Company Name"
            Top             =   990
            Width           =   945
            _ExtentX        =   1667
            _ExtentY        =   503
            MaxLength       =   45
            ReCalcParent    =   "WUCOLDCONAME,WUCNEWCONAME"
            IdName          =   "WCOMCD"
         End
         Begin VB.Label ALBL 
            BackStyle       =   0  'Transparent
            Caption         =   "Enter CoCd :"
            BeginProperty Font 
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
            Left            =   90
            TabIndex        =   176
            Top             =   960
            Width           =   1425
         End
         Begin VB.Label ALBL 
            BackStyle       =   0  'Transparent
            Caption         =   "Enter Name :"
            BeginProperty Font 
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
            Left            =   90
            TabIndex        =   175
            Top             =   1275
            Width           =   1425
         End
         Begin VB.Label Label3 
            BackStyle       =   0  'Transparent
            BorderStyle     =   1  'Fixed Single
            Caption         =   " Update Company Name"
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
            Left            =   -30
            TabIndex        =   174
            Top             =   90
            Width           =   6330
         End
         Begin VB.Label ALBL 
            BackStyle       =   0  'Transparent
            Caption         =   "Supp Pwd :"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Index           =   56
            Left            =   90
            TabIndex        =   173
            Top             =   1590
            Width           =   1425
         End
         Begin VB.Label ALBL 
            BackStyle       =   0  'Transparent
            Caption         =   "Take a Backup. No User should be accessing the Database during this update."
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   495
            Index           =   55
            Left            =   60
            TabIndex        =   171
            Top             =   420
            Width           =   5985
         End
      End
      Begin VB.Frame FraAvgStkRt 
         Height          =   1815
         Left            =   1965
         TabIndex        =   97
         Top             =   6810
         Visible         =   0   'False
         Width           =   7680
         Begin MwfCtl.MWCTL_BTN1 CmdAvgStkRtGo 
            Height          =   285
            Left            =   6930
            TabIndex        =   98
            ToolTipText     =   "Press This Button To Update Database To The Required Version"
            Top             =   405
            Width           =   555
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
         Begin VB.Label ALBL 
            BackStyle       =   0  'Transparent
            Caption         =   $"EmrSuppFrmUpd.frx":032C
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   735
            Index           =   11
            Left            =   150
            TabIndex        =   100
            Top             =   405
            Width           =   6585
         End
         Begin VB.Label LblMain 
            BackStyle       =   0  'Transparent
            BorderStyle     =   1  'Fixed Single
            Caption         =   "  Update Average Stock Rate"
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
            Left            =   -30
            TabIndex        =   99
            Top             =   90
            Width           =   7680
         End
      End
      Begin VB.CommandButton CmdCompUpd 
         Caption         =   "<&j> Enable/ Disable Design Component Module"
         CausesValidation=   0   'False
         Height          =   600
         Left            =   9690
         Style           =   1  'Graphical
         TabIndex        =   16
         ToolTipText     =   "Option To Enable Design Component Module"
         Top             =   780
         Width           =   2400
      End
      Begin VB.CommandButton CmdUpdCoName 
         Caption         =   "<&i> Update Company Name"
         CausesValidation=   0   'False
         Height          =   600
         Left            =   7290
         Style           =   1  'Graphical
         TabIndex        =   14
         ToolTipText     =   "Option To Reset The Autogen Table"
         Top             =   780
         Width           =   2400
      End
      Begin VB.CommandButton CmdUpdPrdPurity 
         Caption         =   "<&h> Update Prd Purity"
         CausesValidation=   0   'False
         Height          =   600
         Left            =   4890
         Style           =   1  'Graphical
         TabIndex        =   13
         ToolTipText     =   "Option To Update the Production Purity of G or P or S"
         Top             =   780
         Width           =   2400
      End
      Begin VB.CommandButton CmdDllsUpd 
         Caption         =   "<&d> Check Dlls"
         CausesValidation=   0   'False
         Height          =   600
         Left            =   7290
         Style           =   1  'Graphical
         TabIndex        =   5
         ToolTipText     =   "Option To Check the Dlls"
         Top             =   180
         Width           =   2400
      End
      Begin VB.CommandButton CmdUpdDsgAna 
         Caption         =   "<&g> Update Design Analysis"
         CausesValidation=   0   'False
         Height          =   600
         Left            =   2490
         Style           =   1  'Graphical
         TabIndex        =   12
         ToolTipText     =   "Option To Update Design Rm Trigger"
         Top             =   780
         Width           =   2400
      End
      Begin VB.CommandButton CmdUpdCoCd 
         Caption         =   "<&a> Change Company Code"
         CausesValidation=   0   'False
         Height          =   600
         Left            =   90
         Style           =   1  'Graphical
         TabIndex        =   2
         ToolTipText     =   "Option To Reset The Autogen Table"
         Top             =   180
         Width           =   2400
      End
      Begin VB.TextBox TxtDisp 
         Height          =   1185
         Left            =   105
         Locked          =   -1  'True
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   28
         Top             =   8730
         Width           =   10920
      End
      Begin VB.CommandButton CmdDelOrd 
         Caption         =   "<&c> Delete Orders"
         CausesValidation=   0   'False
         Height          =   600
         Left            =   4890
         Style           =   1  'Graphical
         TabIndex        =   4
         ToolTipText     =   "Option To Reset The Autogen Table"
         Top             =   180
         Width           =   2400
      End
      Begin VB.CommandButton CmdAvgStkRt 
         Caption         =   "<&m> Update Average Stock Rate"
         CausesValidation=   0   'False
         Height          =   600
         Left            =   210
         Style           =   1  'Graphical
         TabIndex        =   15
         ToolTipText     =   "Option To Reset The Autogen Table"
         Top             =   2130
         Visible         =   0   'False
         Width           =   2400
      End
      Begin VB.CommandButton CmdPurPlt 
         Caption         =   "<&b> Set Pure RM Code"
         CausesValidation=   0   'False
         Height          =   600
         Left            =   2490
         Style           =   1  'Graphical
         TabIndex        =   3
         ToolTipText     =   "Set The Pure Raw Material Code"
         Top             =   180
         Width           =   2400
      End
      Begin VB.CommandButton CmdUpdUA 
         Caption         =   "<&n> Update UserAccess"
         CausesValidation=   0   'False
         Height          =   600
         Left            =   450
         Style           =   1  'Graphical
         TabIndex        =   6
         ToolTipText     =   "Option To Reset The Autogen Table"
         Top             =   2340
         Visible         =   0   'False
         Width           =   2400
      End
      Begin VB.CommandButton CmdUpdRmZCd 
         Caption         =   "<&o> Change Customs Code"
         CausesValidation=   0   'False
         Height          =   600
         Left            =   660
         Style           =   1  'Graphical
         TabIndex        =   7
         ToolTipText     =   "Option To Update Customs Code"
         Top             =   2580
         Visible         =   0   'False
         Width           =   2400
      End
      Begin VB.CommandButton CmdUpdSYyMm 
         Caption         =   "<&p> Update SYyMm"
         CausesValidation=   0   'False
         Height          =   600
         Left            =   90
         Style           =   1  'Graphical
         TabIndex        =   8
         ToolTipText     =   "Option To Update SYyMm"
         Top             =   1980
         Visible         =   0   'False
         Width           =   2400
      End
      Begin VB.CommandButton CmdUpdTrgDsgRm 
         Caption         =   "<&q> Update DsgRm Trigger"
         CausesValidation=   0   'False
         Height          =   600
         Left            =   780
         Style           =   1  'Graphical
         TabIndex        =   9
         ToolTipText     =   "Option To Update Design Rm Trigger"
         Top             =   2820
         Visible         =   0   'False
         Width           =   2400
      End
      Begin VB.CommandButton CmdUpdPrdStg 
         Caption         =   "<&e> Update Production Stage"
         CausesValidation=   0   'False
         Height          =   600
         Left            =   9690
         Style           =   1  'Graphical
         TabIndex        =   10
         ToolTipText     =   "Option To Update Design Rm Trigger"
         Top             =   180
         Width           =   2400
      End
      Begin VB.CommandButton CmdEnaDisaMod 
         Caption         =   "<&f> Enable/Disable Modules"
         CausesValidation=   0   'False
         Height          =   600
         Left            =   90
         Style           =   1  'Graphical
         TabIndex        =   11
         ToolTipText     =   "Option To Update Design Rm Trigger"
         Top             =   780
         Width           =   2400
      End
      Begin MwfCtl.MWCTL_MED atxt 
         Height          =   570
         Index           =   8
         Left            =   105
         TabIndex        =   88
         ToolTipText     =   "Message"
         Top             =   2670
         Visible         =   0   'False
         Width           =   540
         _ExtentX        =   953
         _ExtentY        =   1005
         Alignment       =   3
         Enabled         =   0   'False
         MaxLength       =   255
         DataType        =   4
         IdName          =   "WDISP"
      End
      Begin VB.Frame FRA_KEYS 
         BorderStyle     =   0  'None
         Height          =   390
         Left            =   90
         TabIndex        =   86
         Top             =   210
         Width           =   1350
      End
      Begin VB.Frame FRA_NKEYS 
         BorderStyle     =   0  'None
         Enabled         =   0   'False
         Height          =   390
         Left            =   1905
         TabIndex        =   87
         Top             =   210
         Width           =   1425
      End
      Begin VB.Frame FraUpdIni 
         Height          =   2505
         Left            =   4140
         TabIndex        =   144
         Top             =   6120
         Visible         =   0   'False
         Width           =   4860
         Begin VB.TextBox TxtDllUpdPwd 
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   360
            IMEMode         =   3  'DISABLE
            Left            =   1470
            MaxLength       =   15
            PasswordChar    =   "*"
            TabIndex        =   149
            Top             =   1950
            Width           =   1995
         End
         Begin VB.CommandButton CmdBrow 
            Caption         =   "..."
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   4290
            TabIndex        =   148
            ToolTipText     =   "Press This Button To Update Database To The Required Version"
            Top             =   1560
            Width           =   375
         End
         Begin VB.ComboBox CmbProj 
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   360
            Left            =   1470
            Style           =   2  'Dropdown List
            TabIndex        =   147
            Top             =   1200
            Width           =   1125
         End
         Begin VB.CommandButton CmdUpdIniGo 
            Caption         =   "&Go"
            Height          =   375
            Left            =   3720
            TabIndex        =   146
            ToolTipText     =   "Press This Button To Check Dlls"
            Top             =   1965
            Width           =   975
         End
         Begin VB.TextBox TxtDllPath 
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   360
            Left            =   1470
            MaxLength       =   260
            TabIndex        =   145
            Top             =   1575
            Width           =   2805
         End
         Begin VB.Label ALBL 
            BackStyle       =   0  'Transparent
            Caption         =   "You should have Read And Write Access to the Dlls Path Specified Below for this Check."
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   495
            Index           =   49
            Left            =   90
            TabIndex        =   154
            Top             =   450
            Width           =   4635
         End
         Begin VB.Label ALBL 
            BackStyle       =   0  'Transparent
            Caption         =   "Supp Pwd :"
            BeginProperty Font 
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
            Left            =   120
            TabIndex        =   153
            Top             =   1950
            Width           =   1425
         End
         Begin VB.Label Label7 
            BackStyle       =   0  'Transparent
            BorderStyle     =   1  'Fixed Single
            Caption         =   " Check Dlls"
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
            Left            =   -30
            TabIndex        =   152
            Top             =   90
            Width           =   5490
         End
         Begin VB.Label ALBL 
            BackStyle       =   0  'Transparent
            Caption         =   "Dlls Path :"
            BeginProperty Font 
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
            Left            =   120
            TabIndex        =   151
            Top             =   1575
            Width           =   1425
         End
         Begin VB.Label ALBL 
            BackStyle       =   0  'Transparent
            Caption         =   "Proj. Name :"
            BeginProperty Font 
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
            TabIndex        =   150
            Top             =   1215
            Width           =   1425
         End
      End
      Begin VB.Line Line1 
         Visible         =   0   'False
         X1              =   5805
         X2              =   5805
         Y1              =   4080
         Y2              =   5670
      End
   End
End
Attribute VB_Name = "EmrSuppFrmUpd"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim moCn As MDOConnection    '*** (Bef speed) bef 24/11/06
Dim ms_Opt As String, ws_str As String
'Const mcs_Pwd = "OrcimIsMicro"
Const mcs_Pwd = "SdaSdjPst"
'*** For the Toggle Options of Command Buttons
Private Enum en_Upd
  UpdCoCd = 0
  DelOrd = 1
  AvgStkRt = 2
  PurPlt = 3
  UPDUA = 4
  UpdRmZCd = 5
  UpdSYyMm = 6
  UpdPrdStg = 7
  ' Zubin 212 (23/02/06)
  EnaDisaMod = 8
  ' Zubin 212 (23/02/06)
  UpdDsgAnaCd = 9       ' ****** Sachin 212 - Patch 03 - 29-04-2006
  UpdIni = 10           ' ****** Sachin [31/08/06] - Ini Update from Maintenance
  UpdPrdPurity = 11     ' ****** Sachin 2.14.0 [07/10/2006] - Prd Purity Update
  UpdCoName = 12        ' ****** Sachin 2.14.0 Next - Update Company Name
  UpdCompDsg = 13       ' ****** (Jen 2.14 Next)
  InvFgRm = 14
  UpdRmQw = 15          ' *** (Jen 3.01)
  InsertComp = 16       ' *** Jay 3.1.0_1
End Enum

' ****** Sachin [31-08-06] Update Ini Date from Maintenance Exe ******
Private Declare Sub CoTaskMemFree Lib "ole32.dll" (ByVal hMem As Long)
Private Declare Function lstrcat Lib "kernel32" Alias "lstrcatA" (ByVal lpString1 As String, ByVal lpString2 As String) As Long
Private Declare Function SHBrowseForFolder Lib "shell32" (lpbi As BrowseInfo) As Long
Private Declare Function SHGetPathFromIDList Lib "shell32" (ByVal pidList As Long, ByVal lpBuffer As String) As Long

Private Type BrowseInfo
     hwndOwner As Long
     pIDLRoot As Long
     pszDisplayName As Long
     lpszTitle As Long
     ulFlags As Long
     lpfnCallback As Long
     lParam As Long
     iImage As Long
End Type

Private Const BIF_RETURNONLYFSDIRS = 1
Private Const MAX_PATH = 260

Const NewEmrLckDt = "01/10/2009"        'The New Lock Date will be Random between 14/07/2007 And 15/11/2007
Const NewStoLckDt = "01/10/2009"        'The New Lock Date Set to 01/10/2007

Private Type CoDet
  ts_CoCd As String
  ts_CoNm As String
End Type

Private Type IniDet
  ts_FlNm As String
  ts_PrjNm As String
  ts_SrvrNm As String
  ti_Wait As Integer
  ti_Retry As Integer
  ts_SrlNo As String
  tl_TmOut As Long
  ti_NoOfUsr As Integer
  td_EndDt As Date
  tCoDet() As CoDet
End Type

Dim mt As IniDet, ws_dllpath As String
' ****** Sachin [31-08-06] Update Ini Date from Maintenance Exe ******

Dim ms_Ver As String
Dim ms_Pwd As String

'*** (Jen 3.01)
Private Enum en_IniUpdTyp
  UpdDt = 1
  UpdMod = 2
End Enum
'*** (Jen 3.01)


Private Sub ADC_Load()
  Set moCn = adc.Connection    '*** (Bef speed) bef 24/11/06
  Call HideAllFras
  Call EnaDisaCmds(False)
  CmdUpdPrdPurity.Enabled = False   'Sachin 4.1.0.0
  ms_Ver = Trim(moCn.GetFldVal("Select PDesc225 from Param where PTyp= 'VER' and PMCd= 'DATVER'"))
  'SEO Pwd
  ms_Pwd = moCn.GetFldVal("Select PValue1 from Param where Ptyp = 'SEOPWD' and PmCd = 'SEOPWD'")
End Sub

Private Sub ADC_KeyWhen()
  ' Zubin 212 (20/02/06)
  adc("wFluteBag") = moCn.GetFldVal("Select hFluteBagYN from Head where HCoCd= '" + gs_CoCd + "' and HCd= '" + ctSelfCmCd + "' ")
  adc("wLotNo") = moCn.GetFldVal("Select HLotNoYN from Head where HCoCd= '" + gs_CoCd + "' and HCd= '" + ctSelfCmCd + "' ")
  adc("wSlv") = moCn.GetFldVal("Select HSlvModYN from Head where HCoCd= '" + gs_CoCd + "' and HCd= '" + ctSelfCmCd + "' ")
  adc("wFgQuot") = moCn.GetFldVal("Select HFgQuotYN from Head where HCoCd= '" + gs_CoCd + "' and HCd= '" + ctSelfCmCd + "' ")
  adc("wPpc") = moCn.GetFldVal("Select hPpcYN from Head where HCoCd= '" + gs_CoCd + "' and HCd= '" + ctSelfCmCd + "' ")
  adc("wCTB") = moCn.GetFldVal("Select HCtbMod from Head where HCoCd= '" + gs_CoCd + "' and HCd= '" + ctSelfCmCd + "' ")
  ' Zubin 212 (20/02/06)
  
  '*** (Jen 2.12 Patch 3)
  adc("wLooseInv") = moCn.GetFldVal("Select hLooseInv from Head where HCoCd= '" + gs_CoCd + "' and HCd= '" + ctSelfCmCd + "' ")
  '*** (Jen 2.12 Patch 3)
  
  '*** (Jen 2.13)
  adc("wConsolidatedInv") = moCn.GetFldVal("Select hConsolidatedInv from Head where HCoCd= '" + gs_CoCd + "' and HCd= '" + ctSelfCmCd + "' ")
  adc("wAutomateReturns") = moCn.GetFldVal("Select hAutomateReturns from Head where HCoCd= '" + gs_CoCd + "' and HCd= '" + ctSelfCmCd + "' ")
  '*** (Jen 2.13)
  
  ' ****** Sachin 2.13.0  ******
  adc("wQWMod") = moCn.GetFldVal("Select hQWModYn from Head where HCoCd= '" + gs_CoCd + "' and HCd= '" + ctSelfCmCd + "' ")
  ' ****** Sachin 2.13.0  ******
  
  ' ****** Sachin 2.14.0 - [Multi-Company] ******
  adc("wImpComLocs") = moCn.GetFldVal("Select hImpComLoc from Head where HCoCd= '" + gs_CoCd + "' and HCd= '" + ctSelfCmCd + "' ")
  adc("WINREPMETLSRMCD") = moCn.GetFldVal("Select hInRepMetLsRmCd from Head where HCoCd= '" + gs_CoCd + "' and HCd= '" + ctSelfCmCd + "' ")
  ' ****** Sachin 2.14.0 - [Multi-Company] ******
  
  ' Zubin 213
  adc("wMultiCur") = moCn.GetFldVal("Select hMultiCurYN from Head where HCoCd= '" + gs_CoCd + "' and HCd= '" + ctSelfCmCd + "' ")
  
  ' ****** Sachin [31-08-06] Update Ini Date from Maintenance Exe ******
  CmbProj.Clear
  CmbProj.AddItem "EMR"
  CmbProj.AddItem "STO"
  ' ****** Sachin [31-08-06] Update Ini Date from Maintenance Exe ******
  
  '*** jay 2.13(CT) ***
  adc("wCompBags") = moCn.GetFldVal("Select hCompBagsYn from Head where HCoCd= '" + gs_CoCd + "' and HCd= '" + ctSelfCmCd + "' ")
  '*** jay 2.13(CT) ***
  
  '*** Jay 2.14(LclInv) ***
  adc("wHLclInvYN") = moCn.GetFldVal("Select HLclInvYN from Head where HCoCd= '" + gs_CoCd + "' and HCd= '" + ctSelfCmCd + "' ")
  '*** Jay 2.14(LclInv) ***
  '*** Jay 3.01 *** [FgTag Module]
  adc("wHFgTagMod") = moCn.GetFldVal("Select HFgTagMod from Head where HCoCd= '" + gs_CoCd + "' and HCd= '" + ctSelfCmCd + "' ")
  '*** Jay 3.01 *** [FgTag Module]
  
  '*** (Jen 3.01)
  adc("wHCustomiseOrd") = moCn.GetFldVal("Select HCustomiseOrd from Head where HCoCd= '" + gs_CoCd + "' and HCd= '" + ctSelfCmCd + "' ")
  adc("wHAutoCustmMtch") = moCn.GetFldVal("Select HAutoCustmMtch from Head where HCoCd= '" + gs_CoCd + "' and HCd= '" + ctSelfCmCd + "' ")
  '*** (Jen 3.01)
  
  '*** (Jen 3.1.2)
  adc("whBestUse") = moCn.GetFldVal("Select hBestUse from Head where HCoCd= '" + gs_CoCd + "' and HCd= '" + ctSelfCmCd + "' ")
  '*** (Jen 3.1.2)
  
  '****** Sachin 3.3.0 - [Copy From Excel]******
  adc("wOrdCpyFrmExcel") = moCn.GetFldVal("Select hExcelImpInOrdYN from Head where HCoCd= '" + gs_CoCd + "' and HCd= '" + ctSelfCmCd + "' ")
  adc("wTravelExl") = moCn.GetFldVal("Select hTravelExlYN from Head where HCoCd= '" + gs_CoCd + "' and HCd= '" + ctSelfCmCd + "' ")
  '****** Sachin 3.3.0 - [Copy From Excel]******
  
  ' ***** Manali 3.3.0 - Palladium Module
  adc("wPdMod") = moCn.GetFldVal("Select HPdModYN from Head where HCoCd= '" + gs_CoCd + "' and HCd= '" + ctSelfCmCd + "' ")
  ' ***** Manali 3.3.0 - Palladium Module
  
  '*** Jay 3.3(FG)
  adc("wHFGIOYN") = moCn.GetFldVal("Select HFGIOYN from Head where HCoCd= '" + gs_CoCd + "' and HCd= '" + ctSelfCmCd + "' ")
  '*** Jay 3.3(FG)
  
  ' ****** Manali 3.4.1 - 22/10/08 - Design Catalogue Module
  adc("wDsgCat") = moCn.GetFldVal("Select HDsgCatYN from Head where HCoCd= '" + gs_CoCd + "' and HCd= '" + ctSelfCmCd + "' ")
  ' ****** Manali 3.4.1 - 22/10/08 - Design Catalogue Module
  
  ' ***** Manali 3.6.0 - Order/Invoice Printing in Excel Format
  adc("wHOrdExcelRepYn") = moCn.GetFldVal("Select HOrdExcelRepYn from Head where HCoCd= '" + gs_CoCd + "' and HCd= '" + ctSelfCmCd + "' ")
  ' ***** Manali 3.6.0 - Order/Invoice Printing in Excel Format

  ' ***** Manali 3.9.1 - 21/02/12 - RfId Module
  adc("wHRfIdModYn") = moCn.GetFldVal("Select HRfIdMod from Head where HCoCd= '" + gs_CoCd + "' and HCd= '" + ctSelfCmCd + "' ")
  ' ***** Manali 3.9.1 - 21/02/12 - RfId Module

  '***** Bhavna added Cost Rt Editable module
  adc("wHCstRtEdtYn") = moCn.GetFldVal("Select HCstrtEdtYn from Head where HCoCd= '" + gs_CoCd + "' and HCd= '" + ctSelfCmCd + "' ")
  
  '***** Bhavna added Rework module
  adc("whRwrkModYn") = moCn.GetFldVal("Select hRwrkModYn from Head where HCoCd= '" + gs_CoCd + "' and HCd= '" + ctSelfCmCd + "' ")
  
  ' ***** Manali 3.10.0 - 03/04/12 - Multi Price Quotation Module
  adc("wHMultiPrcQtModYn") = moCn.GetFldVal("Select HMultiPrcQtMod from Head where HCoCd= '" + gs_CoCd + "' and HCd= '" + ctSelfCmCd + "' ")
  ' ***** Manali 3.10.0 - 03/04/12 - Multi Price Quotation Module

  ' ***** Sachin 3.11.0 - Rate from Rate Chart in Invoice Module
  adc("whRtfrRtChrtYN") = moCn.GetFldVal("Select hRtfrRtChrtYN from Head where HCoCd= '" + gs_CoCd + "' and HCd= '" + ctSelfCmCd + "' ")

End Sub

Private Sub ATXT_FldWhen(Index As Integer, ByVal IdName As String, Cancel As Boolean, ErrMsg As String)
  Select Case UCase(IdName)
  '*** Update CoCd
  Case Is = UCase("wUcOldCoCd"), UCase("wComCd"), UCase("wSrcCoCd")
    Call HlpList.hCoCd
  '*** Update CoCd
  '*** Order ***
  Case Is = UCase("wDoOmTcFr"), UCase("wDoOmTcTo")
    Call HlpList.PMCd("TC", "'SO'")
  Case Is = UCase("wDoOmYyFr"), UCase("wDoOmYyTo")
    Call HlpList.PMCd("YY")
  Case Is = UCase("wDoOmChrFr")
    Call HlpList.vPSCd(gs_CoCd, "CHR", adc("wDoOmTcFr"))
  Case Is = UCase("wDoOmChrTo")
    Call HlpList.vPSCd(gs_CoCd, "CHR", adc("wDoOmTcTo"))
  Case Is = UCase("wDoOmNoFr")
    Call HlpList.OmNo(gs_CoCd, adc("wDoOmTcFr"), adc("wDoOmYyFr"), adc("wDoOmChrFr"))
  Case Is = UCase("wDoOmNoTo")
    Call HlpList.OmNo(gs_CoCd, adc("wDoOmTcTo"), adc("wDoOmYyTo"), adc("wDoOmChrTo"))
  Case Is = UCase("wDoOmSrFr")
    Call HlpList.OdSr(gs_CoCd, adc("wDoOmTcFr"), adc("wDoOmYyFr"), adc("wDoOmChrFr"), adc("wDoOmNoFr"))
  Case Is = UCase("wDoOmSrTo")
    Call HlpList.OdSr(gs_CoCd, adc("wDoOmTcTo"), adc("wDoOmYyTo"), adc("wDoOmChrTo"), adc("wDoOmNoTo"))
  '*** Order ***
  '*********Geeta***********
  Case Is = UCase("wRmCd")
    Call HlpList.RmCd("N", "'" + adc("WRMCTG") + "'")
  Case Is = UCase("wOldRmZCd")
    Call HlpList.RmCd("Y")
  '****** Sachin 2.12 - Silver Pure Code
  Case Is = UCase("wRmCtg"), UCase("WRMCTGPPUR")
    Call HlpList.PMCd("RmCtg")
  '****** (Jen 2.13) added wConsolidatedInv, wAutomateReturns ******
  '****** Sachin 2.12 - Silver Pure Code
  ' Zubin 212 (23/02/06)
  ' Zubin 213 (wMultiCur added)
  '*** jay 2.13(CT) *** (WCompBags added)
  '****** Sachin - 26-10-2006 [Multi-Company] ******
  '*** Jay 2.14(LclInv) *** (wHLclInvYN added)
  '*** Jay 3.01 *** [FgTag Module] (HFgTagMod added)
  '****** (Jen 3.01 added wHCustomiseOrd, wHAutoCustmMtch) ******
  '****** (Jen 3.1.2) added whBestUse ******
  '****** Sachin 3.3.0 - added wOrdCpyFrmExcel
  ' ***** Manali 3.3.0 - added wPdMod [Palladium Module]
  '*** Jay 3.3(FG)
  ' *** Manali 3.4.1 - 22/10/08 - HDsgCatYN added
  ' ***** Manali 3.6.0 - Order Printing in Excel Format
  ' ***** Manali 3.9.0 - 21/02/12 Rfid Module
  ' ***** Manali 3.10.0 - 03/04/12 - Multi Price Quotation Module
  ' *********Bhavna added Cost Rt editable module
  Case Is = UCase("wFluteBag"), UCase("wLotNo"), UCase("wSlv"), UCase("wFgQuot"), _
            UCase("wPPC"), UCase("wCTB"), UCase("wLooseInv"), UCase("wConsolidatedInv"), UCase("wAutomateReturns"), _
            UCase("wQWMod"), UCase("wMultiCur"), UCase("WCompBags"), UCase("wImpComLocs"), UCase("WINREPMETLSRMCD"), _
            UCase("wHLclInvYN"), UCase("wHFgTagMod"), UCase("wHCustomiseOrd"), UCase("wHAutoCustmMtch"), _
            UCase("whBestUse"), UCase("wOrdCpyFrmExcel"), UCase("wPdMod"), UCase("wTravelExl"), UCase("wHFGIOYN"), _
            UCase("wDsgCat"), UCase("wHOrdExcelRepYn"), UCase("wHRfIdModYn"), UCase("wHMultiPrcQtModYn"), _
            UCase("wHRtFrRtChrtYN"), UCase("wHCstRtEdtYn"), UCase("whRwrkModYn")
    Call HlpList.PMCd("YN")
  ' Zubin 212 (23/02/06)
  
  ' ****** Sachin 2.14.0 [12-10-06] Production Purity Update ******
  Case Is = UCase("wRmCdPPur")
    Call HlpList.RmCd("N", "'" + adc("WRMCTGPPUR") + "'")
  ' ****** Sachin 2.14.0 [12-10-06] Production Purity Update ******
  '*** Jay 2.14Next ***
  Case Is = UCase("WCOMPUPDCHR")
      If adc.Mode = xNorm And UCase(adc("wCompBags")) = "N" Then Cancel = True: ErrMsg = "Can Not Edit This Field."
  '*** Jay 2.14Next ***
  End Select
End Sub
Private Sub ATXT_FldChange(Index As Integer, ByVal IdName As String, Cancel As Boolean, ErrMsg As String, ByVal pv_NewValue As Variant)
  Select Case UCase(IdName)
  Case Is = UCase("wUcOldCoCd")
    Cancel = pv_NewValue <> "" And Not moCn.RecSeek("Select HCoCd from Head where HCoCd= '" + pv_NewValue + "' and HCd= '" + ctSelfCmCd + "'")
    If Cancel = True Then ErrMsg = "Invalid Company Code": Exit Sub
  Case Is = UCase("wUcNewCoCd")
    Cancel = pv_NewValue <> "" And pv_NewValue <> adc("wUcOldCoCd") And _
             moCn.RecSeek("Select HCoCd from Head where HCoCd= '" + pv_NewValue + "' and HCd= '" + ctSelfCmCd + "'")
    If Cancel = True Then ErrMsg = "Company Code Already Exists": Exit Sub
  '*************Geeta*************Emr206**********
  Case Is = UCase("wRmCd")
    Cancel = pv_NewValue <> "" And _
            Not moCn.RecSeek("Select RmCd from RmMst where RmCd= '" + pv_NewValue + "' and RmCtg = '" + adc("WRMCTG") + "'")
    If Cancel = True Then ErrMsg = "Invalid RmCd For Selected Category": Exit Sub
    
    '****** Sachin 2.13.0 - [12-08-2006] - [Q/W Module]
    Cancel = pv_NewValue <> "" And _
            Not moCn.RecSeek("Select RmCd from RmMst where RmCd= '" + pv_NewValue + "' and RmCtg = '" + adc("WRMCTG") + "' And RmQw='W' And RmQwCst='W' ")
    If Cancel = True Then ErrMsg = "Invalid RmCd As the Q/W Cost/Sale Specified in RmMst is not Weight Wise": Exit Sub
    '****** Sachin 2.13.0 - [12-08-2006] - [Q/W Module]
   
  Case Is = UCase("wOldRmZCd")
    Cancel = pv_NewValue <> "" And _
            Not moCn.RecSeek("Select RmCd from RmMst where RmCd= '" + pv_NewValue + "' and RmZ= 'Y'")
    If Cancel = True Then ErrMsg = "Invalid Customs Code": Exit Sub
  Case Is = UCase("WRMCTG")
    Cancel = Not (pv_NewValue = "P" Or pv_NewValue = "S")
    If Cancel = True Then ErrMsg = "Only Platinum Or Silver Category Should be Selected": Exit Sub
    
  '****** Sachin 2.14.0 - [07/10/06] - [Update Prd Purity for G,P,S]
  Case Is = UCase("WRMCTGPPUR")
    Cancel = pv_NewValue <> "" And Not (pv_NewValue = "G" Or pv_NewValue = "P" Or pv_NewValue = "S" Or pv_NewValue = "L")
    If Cancel = True Then ErrMsg = "Only Gold, Platinum, Silver or Palladium Category Should be Selected": Exit Sub
    
  Case Is = UCase("wRmCdPPur")
    Cancel = pv_NewValue <> "" And Not moCn.RecSeek("Select 'x' From RmMst Where RmCtg='" + adc("wRmCtgPPur") + "' And RmCd='" + pv_NewValue + "'")
    If Cancel = True Then ErrMsg = "Invalid Raw Material Code": Exit Sub
  '****** Sachin 2.14.0 - [07/10/06] - [Update Prd Purity for G,P,S]
    
'  Case Is = UCase("wNewRmZCd")
'    Cancel = pv_NewValue <> "" And pv_NewValue <> ADC("wUcOldCoCd") And _
'             moCn.RecSeek("Select RmCd from RmMst where RmCd= '" + pv_NewValue + "'")
'    If Cancel = True Then ErrMsg = "Customs/Raw Material Code Already Exists": Exit Sub
'

  '****** (Jen 2.13) added wConsolidatedInv, wAutomateReturns ******
  ' Zubin 212 (23/02/06)
  ' Zubin 213 (wMultiCur added)
  '*** jay 2.13(CT) *** (WCompBags added)
  '****** Sachin 2.14.0 [wImpComLoc] added ******
  '*** Jay 2.14(LclInv) *** (wHclInv added)
  '*** Jay 3.01 *** [FgTag Module]
  '*** (Jen 3.01 added wHCustomiseOrd, wHAutoCustmMtch
  '****** (Jen 3.1.2) added whBestUse
  '****** Sachin 3.3.0 - added wOrdCpyFrmExcel
  ' ***** Manali 3.3.0 - added wPdMod [Palladium Module]
  '*** Jay 3.3(FG)
  ' ***** Manali 3.4.1 - 22/10/08 - HDsgCat added
  ' ***** Manali 3.6.0 - Order Printing in Excel Format
  ' ***** Manali 3.9.1 - 20/02/12 - RfId Module
  ' ***** Manali 3.10.0 - 03/04/12 - Multi Price Quotation Module
  ' ******Bhavna added cost rt editable module
  Case Is = UCase("wFluteBag"), UCase("wLotNo"), UCase("wSlv"), UCase("wFgQuot"), _
            UCase("wPPC"), UCase("wCTB"), UCase("wLooseInv"), UCase("wConsolidatedInv"), UCase("wAutomateReturns"), _
            UCase("wQWMod"), UCase("wMultiCur"), UCase("WCompBags"), UCase("WImpComLocs"), UCase("WINREPMETLSRMCD"), _
            UCase("wHLclInvYN"), UCase("wHFgTagMod"), UCase("wHCustomiseOrd"), UCase("wHAutoCustmMtch"), _
            UCase("whBestUse"), UCase("wOrdCpyFrmExcel"), UCase("wPdMod"), UCase("wTravelExl"), UCase("wHFGIOYN"), _
            UCase("wDsgCat"), UCase("wHOrdExcelRepYn"), UCase("wHRfIdModYn"), UCase("wHMultiPrcQtModYn"), _
            UCase("wHRtFrRtChrtYN"), UCase("wHCstRtEdtYn"), UCase("whRwrkModYn")
    Cancel = Not moCn.RecSeek("Select 'x' from Param where PTyp= 'YN' and PMCd= '" + pv_NewValue + "' and " + _
                              "PSCd= '' ")
    If Cancel = True Then ErrMsg = "Invalid Option. Enter Y/N.": Exit Sub
  ' Zubin 212 (23/02/06)
  '****** Sachin 2.14.0 Next [12-02-07] ******
  Case Is = UCase("WCOMCD")
    Cancel = pv_NewValue <> "" And Not moCn.RecSeek("Select 'x' from CustMst Where CmCtg='Z' And CmCd= '" + pv_NewValue + "' ")
    If Cancel = True Then ErrMsg = "Invalid Company Code/Company Code Not Defined in Company Master": Exit Sub
  '****** Sachin 2.14.0 Next [12-02-07] ******

  '*** (Jen 2.14 Next)
  Case Is = UCase("wCompUpdChr")
    Cancel = pv_NewValue <> "" And moCn.RecSeek("Select 'x' from Param where PTyp= 'BCHR' and PMCd= '" + pv_NewValue + "' ")
    If Cancel = True Then ErrMsg = "Bag Character Already Exists": Exit Sub
  '*** (Jen 2.14 Next)
  '*** Jay 3.1.0_1
  Case Is = UCase("wSrcCoCd")
    If Not moCn.RecSeek("select 'x' from Head where HCd= '" + ctSelfCmCd + "' and HCoCd= '" + pv_NewValue + "' ") Then
      Cancel = True: ErrMsg = "Invalid Company Code": Exit Sub
    End If
  End Select
  '*** Jay 3.1.0_1
End Sub
Private Sub atxt_ReCalculate(Index As Integer, ByVal IdName As String, ByVal OldVal As Variant, ByVal RowNum As Integer, ByVal SourceName As String, ByVal SourceColName As String)
  Select Case IdName
  Case Is = UCase("wUcOldCoName"), UCase("wUcNewCoName")
    If adc.Mode = xNorm Then
      adc(IdName) = moCn.GetFldVal("Select HName from Head where HCoCd= '" + adc("wUcOldCoCd") + "' and HCd= '" + ctSelfCmCd + "'")
    End If
    
    '*** Jay 2.14Next ***
    Case Is = UCase("WCOMPUPDCHR")
      If adc.Mode = xNorm And UCase(adc("wCompBags")) = "N" Then adc("WCOMPUPDCHR") = ""
    '*** Jay 2.14Next ***
  End Select
End Sub
Private Sub CmdInvFgRmMisMatch_Click()
    Call DispFra(InvFgRm)
    If FraUpdUA.Visible = True Then
     ms_Opt = "UPDUA"
    End If
End Sub
Private Sub CmdInvFgRmMisMatchGo_Click()
  Call InvFgRmMisMatch
  If FraUpdUA.Visible = True Then ms_Opt = "InvFgRmErr"
  'Call DispFra(InvFgRm)
End Sub
Private Sub GrdDsgAnlyFlds_ColChange(ByVal RowNum As Integer, ByVal ColName As String, ByVal NewValue As Variant, Cancel As Boolean, ErrMsg As String)
    ' ****** Sachin 2.12 Patch 02 - Default Analysis Code Update for Unidesign - 10/04/06 ******
    Select Case UCase(ColName)
      Case Is = UCase("WDANADEFVAL")
        Dim ws_PValue As String
            Cancel = Not moCn.RecSeek("Select PSCd From Param Where " + _
                                          "PTyp= 'DAANACD' And PMCd= '" + GrdDsgAnlyFlds.Value(RowNum, "WDANASR") + "' and " + _
                                          "PSCd = '" + NewValue + "'")
        If Cancel = True Then ErrMsg = "Invalid Analysis Code": Exit Sub
    End Select
End Sub
Private Sub GrdDsgAnlyFlds_ColWhen(ByVal RowNum As Integer, ByVal ColName As String, Cancel As Boolean, ErrMsg As String, ByVal Action As MwfCtl.en_fgAction)
    Select Case UCase(ColName)
      Case Is = UCase("WDANADEFVAL")
        Call HlpList.PSCd("DAANACD", GrdDsgAnlyFlds.Value(RowNum, "WDANASR"))
    End Select
    ' ****** Sachin 2.12 Patch 02 - Default Analysis Code Update for Unidesign - 10/04/06 ******
End Sub
Private Sub GrdCompUpd_ColWhen(ByVal RowNum As Integer, ByVal ColName As String, Cancel As Boolean, ErrMsg As String, ByVal Action As MwfCtl.en_fgAction)
  '*** (Jen 2.14 Next)
  Cancel = True: ErrMsg = "Cannot Edit Record": Exit Sub
  '*** (Jen 2.14 Next)
End Sub
Private Sub CmdBrow_Click()
  ' ****** Sachin [31-08-06] Update Ini Date from Maintenance ******
    TxtDllPath.text = BrowseForFolder(hwnd, "Please select a Dlls folder.")
  ' ****** Sachin [31-08-06] Update Ini Date from Maintenance ******
End Sub
Private Sub CmdCoNameUpdGo_Click()

' ****** Sachin 2.14.0 Next [12-02-07] ******
  On Error GoTo ErrHandler
  
    If adc("WCONAME") = "" Or adc("WCOMCD") = "" Then _
        DispMsg "New Company Code Cannot Be Blank", etError: Exit Sub
    
    If Trim$(TxtCoNmSuppPwd.text) = "" Then
      DispMsg "Enter Support Password for this Update", etError
    ElseIf Trim$(TxtCoNmSuppPwd.text) <> mcs_Pwd Then
      DispMsg "Not a Valid Support Password", etError
    Else
      moCn.Execute (" Insert Into TrgDisa Values ('CustMst') " + vbCrLf + _
                    " Update CustMst Set CmName='" + CStr(adc("WCONAME")) + "' Where CmCtg='Z' And CmCd='" + adc("WCOMCD") + "' " + _
                    vbCrLf + " Delete From TrgDisa ")
      TxtDisp.text = "Company Name Updated Successfully"
      Call DispFra(UpdCoName)
    End If
  Exit Sub

ErrHandler:
   If moCn.GetFldVal("Select 'x' From TrgDisa") <> 0 Then moCn.Execute ("Delete From TrgDisa")
   TxtDisp.text = Err.Description
' ****** Sachin 2.14.0 Next [12-02-07] ******

End Sub

Private Sub CmdDllsUpd_Click()
  ' ****** Sachin [31-08-06] Update Ini Date from Maintenance ******
    Call DispFra(UpdIni)
    CmbProj.ListIndex = 0
  ' ****** Sachin [31-08-06] Update Ini Date from Maintenance ******
End Sub

Private Sub CmdUpdDsgAna_Click()
    Call DispFra(UpdDsgAnaCd)
    Call LstDsgAnaCds
    GrdDsgAnlyFlds.AllowAdd = False
    GrdDsgAnlyFlds.AllowDelete = False
End Sub
Private Sub CmdDsgAnaUpdGo_Click()
  Dim ws_cnt As Integer
  ws_cnt = 1
  ms_Pwd = moCn.GetFldVal("Select PValue1 from Param where Ptyp = 'SEOPWD' and PmCd = 'SEOPWD'")
  If Trim$(txtseopwd.text) = "" Then
    DispMsg "Enter SEO Password for this Update", etError
  ElseIf Trim$(txtseopwd.text) <> ms_Pwd Then
    DispMsg "Not a Valid SEO Password", etError
  Else
       Do While ws_cnt <= GrdDsgAnlyFlds.Rows - 1
            If Trim(GrdDsgAnlyFlds.Value(ws_cnt, "WDANADEFVAL")) = "" Then
                TxtDisp.text = "All the Analysis Codes Must Be Entered for this Update!!": Exit Sub
            End If
            ws_cnt = ws_cnt + 1
       Loop
       Call UpdDefDsgAnaCodes
       Call DispFra(UpdDsgAnaCd)
  End If
End Sub
' Zubin 212 (23/02/06)
Private Sub CmdEnaDisaMod_Click()
  '*** Display or hide the Copy frame depending on the toggle state
  Call DispFra(EnaDisaMod)
  If FraEnaDisaMod.Visible = True Then adc("wFluteBag").SetFocus
End Sub
' Zubin 212 (23/02/06)
Private Sub CmdEnaDisaModGO_Click()
On Error GoTo ErrHndlr
  Dim rs_ModuleYN As MwfLib.MDORowSet
  If TxtModulePwd <> mcs_Pwd Then DispMsg "Invalid Support Password", etError: Exit Sub
  moCn.BeginTrans (mdoCommitRec)
  '*** Jay 3.3(FG)
  ' **** Manali 3.4.1 - 22/10/08 - Design Catalogue Module
  ' **** Manali 3.6.0 - [HOrdExcelRepYn] added
  ' ***** Manali 3.9.1 - 21/02/12 - RfId Module
  ' ***** Manali 3.10.0 - 03/04/12 - Multi Price Quotation Module
  ' ****** Bhavna added cost rt editable mod
  Set rs_ModuleYN = adc.Connection.OpenResultset("select HCoCd, hPPCYN, hFluteBagYN, hSlvModYN, hCtbMod, hLooseInv, " + _
                    "hConsolidatedInv, hAutomateReturns, hQWModYn, hMultiCurYN, hCompBagsYn, hImpComLoc, " + _
                    "hInRepMetLsRmCd, HLclInvYN, HFgTagMod, HCustomiseOrd, hCustomizeRmCtgExp, HAutoCustmMtch, hBestUse, " + _
                    "hExcelImpInOrdYN, HPdModYN, hTravelExlYN, hChkIntgty,HFGIOYN, HDsgCatYN, HOrdExcelRepYn, HRfIdMod, " + _
                    "HMultiPrcQtMod,hRtfrRtChrtYN,hCstRtEdtYn,hRwrkModYn from Head " + _
                    "where HCoCd= '" + ctSelfCoCd + "' and HCd= '" + ctSelfCmCd + "' ")
                    
  If Not (rs_ModuleYN Is Nothing) Then
    With rs_ModuleYN
      .FldValue("hFluteBagYN") = adc("wFluteBag")
      .FldValue("hPpcYN") = adc("wPPC")
      .FldValue("hSlvModYN") = adc("wSlv")
      .FldValue("hCtbMod") = adc("wCTB")
      .FldValue("hLooseInv") = adc("wLooseInv")   '*** (Jen 2.12 Patch 3)
      '*** (Jen 2.13)
      .FldValue("hConsolidatedInv") = adc("wConsolidatedInv")
      .FldValue("hAutomateReturns") = adc("wAutomateReturns")
      '*** (Jen 2.13)
      .FldValue("hQWModYn") = adc("wQWMod")   '****** Sachin 2.13.0 - 05-06-2006 ******
      ' Zubin 213
      .FldValue("hMultiCurYN") = adc("wMultiCur")
      .FldValue("hCompBagsYn") = adc("wCompBags") '*** jay 2.13(CT) ***
      .FldValue("hImpComLoc") = adc("wImpComLocs")   '****** Sachin 2.14.0 - 26-10-2006 - [Multi-Company] ******
      .FldValue("hInRepMetLsRmCd") = adc("WINREPMETLSRMCD")   '****** Sachin 2.14.0 Next ******
      .FldValue("HLclInvYN") = adc("WHLCLINVYN") '*** Jay 2.14(LclInv) ***
      .FldValue("HFgTagMod") = adc("wHFgTagMod") '*** Jay 3.01 *** [FgTag Module]
      '*** (Jen 3.01)
      .FldValue("HCustomiseOrd") = adc("wHCustomiseOrd")
      .FldValue("HAutoCustmMtch") = adc("wHAutoCustmMtch")
      '*** (Jen 3.01)
      .FldValue("hBestUse") = adc("whBestUse")    '*** (Jen 3.1.2)
      .FldValue("hExcelImpInOrdYN") = adc("WORDCPYFRMEXCEL")    '****** Sachin 3.3.0
      .FldValue("HPdModYN") = adc("wPdMOd")       ' ****** Manali 3.3.0
      .FldValue("hTravelExlYN") = adc("WTRAVELEXL") '****** Sachin 3.3.0 - Travel Excel
      .FldValue("HFGIOYN") = adc("wHFGIOYN")    '*** Jay 3.3(FG)
      .FldValue("HDsgCatYN") = adc("wDsgCat") ' ***** Manali 3.4.1 - 22/10/08 - Design Catalogue Module
      .FldValue("HOrdExcelRepYn") = adc("wHOrdExcelRepYn") '***** Manali 3.6.0 - [HOrdExcelRepYn] - Order/Inv Printing in Excel Format
      .FldValue("HRfIdMod") = adc("wHRfIdModYn") ' ***** Manali 3.9.1 - 21/02/12 - RfId Module
      .FldValue("HCstRtEdtYn") = adc("wHCstRtEdtYn")  '**Bhavna Cost rt editable y/n
      .FldValue("hRwrkModYn") = adc("whRwrkModYn")  '**Bhavna Rework Module  y/n
      .FldValue("HMultiPrcQtMod") = adc("wHMultiPrcQtModYn")      ' ***** Manali 3.10.0 - 03/04/12 - Multi Price Quotation Module
      .FldValue("hRtfrRtChrtYN") = adc("wHRtFrRtChrtYN")      ' ***** Sachin 3.11.0 - Rate from Rate Chart in Invoice Module
      .Update
    End With
  End If
  
  Set rs_ModuleYN = Nothing
  Set rs_ModuleYN = adc.Connection.OpenResultset("select HCoCd, hLotNoYN, hFgQuotYN from Head " + _
                                                  "where HCoCd= '" + gs_CoCd + "' and HCd= '" + ctSelfCmCd + "' ")
  If Not (rs_ModuleYN Is Nothing) Then
    With rs_ModuleYN
      .FldValue("hLotNoYN") = adc("wLotNo")
      .FldValue("hFgQuotYN") = adc("wFgQuot")
      .Update
    End With
  End If
  Set rs_ModuleYN = Nothing
  
  moCn.CommitTrans mdoCommitRec
  
  Call UpdIntegrity(moCn)     '****** Sachin 3.3.0
  
  DispMsg "Module Rights Updated", etInfo
  Call DispFra(EnaDisaMod)
Exit Sub
ErrHndlr:
  If Err.Number = -2147217900 Then
    rs_ModuleYN.CancelUpdate: moCn.RollbackTrans mdoCommitNone
    DispMsg Mid(Err.Description, InStr(1, Err.Description, "|", vbTextCompare)), etError
    Err.Clear
    Exit Sub
  End If
  DispMsg Err.Description, etError
  Err.Clear
End Sub
'Private Sub CmdMtnPwdGo_Click()
'  ' Zubin 212
'  If TxtMTNPwd <> "MWTEAM" Then
'    DispMsg "Invalid Password", etError: Exit Sub
'  Else
'    FraMtnPwd.Visible = False: FraMain.Visible = True
'  End If
'End Sub

Private Sub CmdCompUpd_Click()
  '*** (Jen 2.14 Next)
  '*** Display or hide the Copy frame depending on the toggle state
  Call DispFra(UpdCompDsg)
  If FraCompUpd.Visible = True Then
    If GrdCompUpd.Rows - 1 = 0 Then FraCompUpdDispErrGrd.Visible = False
    'adc("wCompUpdChr").SetFocus
  End If
  '*** (Jen 2.14 Next)
End Sub
Private Sub CmdCompUpdGo_Click()
  '*** (Jen 2.14 Next)
  On Error GoTo ErrHndlr
  Dim wb_UpdBChrCnt As Boolean    '*** Jay 3.01 *** [Dynamic BCHR]
  Dim Rs_InvalidComp As MwfLib.MDORowSet
  Dim ws_str As String, ws_hCompBagsYn As String
  Dim wi_Row As Integer
  
  TxtDisp.text = ""
  DispMsg "", etInfo
  
  '*** Jay 2.14Next ***
  If TxtCompUpdPwd <> mcs_Pwd Then DispMsg "Invalid Support Password", etError: Call DispFra(UpdCompDsg): Exit Sub
  '*** Jay 3.01 *** [Dynamic BCHR]
  ws_hCompBagsYn = UCase(moCn.GetFldVal("Select hCompBagsYn from Head where HCoCd= '" + ctSelfCoCd + "' and HCd= '" + ctSelfCmCd + "' "))
  If adc("WCompBags") = "N" And ws_hCompBagsYn = "N" Then
    DispMsg "Design Component Module Is Already OFF", etError
    TxtDisp.text = "Design Component Module Is Already OFF"
  ElseIf adc("WCompBags") = "N" And ws_hCompBagsYn = "Y" Then
    moCn.Execute ("Update Head Set hCompBagsYn= 'N' where HCoCd= '" + ctSelfCoCd + "' and HCd= '" + ctSelfCmCd + "'")
    Call UpdIntegrity(moCn)     '****** Manali 3.6.0 - 16/10/09
  ElseIf adc("WCompBags") = "Y" And ws_hCompBagsYn = "Y" Then
    If CInt(moCn.GetFldVal("Select Count('x') from Param where Ptyp= 'BCHR' and PValue= 'C'")) <= CInt(Trim(adc("wBChrCnt"))) Then
        ws_str = "Insert into TrgDisa values ('Param') " + _
                 "Update Param Set PNum= " + CStr(Trim(adc("wBChrCnt"))) + " Where PTyp= 'BTYP' and PMCd= 'C' " + _
                 "delete TrgDisa"
        moCn.Execute (ws_str)
        DispMsg "Maximum No of Comp BChr's Can be Open Upto: " + CStr(adc("wBChrCnt")), etInfo
        TxtDisp.text = "Maximum No of Comp BChr's Can be Open Upto: " + CStr(adc("wBChrCnt"))
    Else
        DispMsg "Entered No of Comp BChr's " + CStr(adc("wBChrCnt")) + " are Lesser Than Current No of Comp BChr's", etInfo
        TxtDisp.text = "Entered No of Comp BChr's " + CStr(adc("wBChrCnt")) + " are Lesser Than Current No of Comp BChr's"
    End If
  ElseIf adc("WCompBags") = "Y" And ws_hCompBagsYn = "N" Then
  '*** Jay 3.01 *** [Dynamic BCHR]
    If adc("wCompUpdChr") = "" Then DispMsg "New Component Character For Default Update Cannot Be Blank", etError: Call DispFra(UpdCompDsg): Exit Sub
    
    Set Rs_InvalidComp = moCn.OpenRes("Select Top 300 DctCd, DctSz from DsgCt " + _
                                      "where exists(Select 'x' from DsgMst where DmTcTyp= 'PM' and DmCd= DctCd and " + _
                                              "DmSz= DctSz and (DmPrdSeq = '' or not exists(Select 'x' From DsgRm " + _
                                              "where DrTcTyp= DmTcTyp and DrCd= DmCd and DrSz= DmSz and " + _
                                              "DrMainMet= 'Y' and DrRmCtg in ('G','P','S', 'L')) ))")
    '*** Jay 2.14Next ***
    GrdCompUpd.Rows = 1
    If Rs_InvalidComp.RecCount > 0 Then
      FraCompUpdDispErrGrd.Visible = True
      With GrdCompUpd
        .AllowAdd = True
        Do While Not (Rs_InvalidComp.EOF Or Rs_InvalidComp.BOF)
          .AddItem
          wi_Row = .Rows - 1
          .Value(wi_Row, "wCompUpdDsgCd") = Rs_InvalidComp!DctCd
          .Value(wi_Row, "wCompUpdDsgSz") = Rs_InvalidComp!DctSz
          Rs_InvalidComp.MoveNext
        Loop
        .AllowAdd = False
      End With
      DispMsg "Cannot Update As Design Component RM Records Do Not Have Main Metal Details", etError: Exit Sub
    Else
      FraCompUpdDispErrGrd.Visible = False
    End If
    DoEvents
    MWLib.BeginProcess Me, "Starting With Design Component Update"
  '6.1
    ws_str = " Set XACT_ABORT ON " + _
             "BEGIN TRAN UpdDsgComp " + _
             "set dateformat " + DtFmtStr() + " " + _
             "SET ANSI_WARNINGS OFF "
    '*** Jay 2.14Next ***
    '*** Jay 3.01 *** [Dynamic BCHR] [(Update Param Set PNum) Part Added]
    ws_str = ws_str + "Insert into TrgDisa values ('HEAD') " + _
            "Insert into TrgDisa values ('Param') " + _
            "Update Head Set hCompBagsYn= 'Y' " + _
            "Update Param Set PNum= " + CStr(Trim(adc("wBChrCnt"))) + " Where PTyp= 'BTYP' and PMCd= 'C' " + _
            "Delete TrgDisa " + _
            "Insert into Param (PTyp, PMCd, PSCd, PDesc, PDesc225, PValue, PNum, PValue1, PNum1, PValue2, " + _
            "ModUsr, ModDt, ModTime, PValue3, PValidYn) " + _
            "values ('BCHR', '" + Trim(adc("wCompUpdChr")) + "', '', 'Default Component Chr', '', 'C', 0, '', 0, '', 'MIC', '01/01/80', 12.00, '', '') " + _
            "Update DsgCt set DctChr= '" + Trim(adc("wCompUpdChr")) + "', DctRmCd= DrRmCd, DctDmCol= DmCol, DctRmWt= DrPrdWt, " + _
            "   DctDc= IsNull((Select (case when sum(case when Dc.DrRmCtg= 'D' then 1 else 0 end)> 0 then 'D' else '' end)+ " + _
            "                         (case when sum(case when Dc.DrRmCtg= 'C' then 1 else 0 end)> 0 then 'C' else '' end) " + _
            "                  From DsgRm Dc where Dc.DrTcTyp= DmTcTyp and Dc.DrCd= DmCd and Dc.DrSz= DmSz and Dc.DrRmCtg in ('C','D') ), ''), " + _
            "   DctPrdSeq = DmPrdSeq " + _
            "From DsgMst join DsgRm Dr on DrTcTyp= DmTcTyp and DrCd= DmCd and DrSz= DmSz " + _
            "where DmTcTyp= 'PM' and DmCd= DctCd and DmSz= DctSz and DrMainMet= 'Y' and DrRmCtg in ('G','P','S') and DctChr='' "
  '*** Jay 2.14Next ***
    ws_str = ws_str + " SET ANSI_WARNINGS ON " + _
                      "/*** @@Error ***/ " + _
                      "COMMIT TRAN UpdDsgComp " + _
                      "Set XACT_ABORT OFF "
    
    moCn.Execute (ws_str)
    
    Call UpdIntegrity(moCn)     '****** Manali 3.6.0 - 16/10/09
    
    MWLib.EndProcess Me
    
    If moCn.RecSeek("Select 'x' from Head where hCompBagsYn= 'N' ") Then
      DispMsg "Design Component Module Could Not Be Updated", etInfo
      TxtDisp.text = "Design Component Module Could Not Be Updated"
    Else
      DispMsg "Design Component Module Updated Successfully", etInfo
      TxtDisp.text = "Design Component Module Updated Successfully"
    End If
  End If
  Call DispFra(UpdCompDsg)
  Exit Sub
  '*** Jay 3.01 *** [Dynamic BCHR]
ErrHndlr:
  DispMsg Err.Description, etError
  TxtDisp.text = Err.Description
  Err.Clear
  '*** (Jen 2.14 Next)
End Sub
Private Sub CmdPurPlt_Click()
  Call DispFra(PurPlt)
  If FraPurPlt.Visible = True Then
    adc("wRmCtg").SetFocus
    ms_Opt = "PurPlt"
  End If
End Sub
Private Sub CmdPurPltGo_Click()
  On Error GoTo ErrHandler
  ' Zubin 212 (24/02/06)
  If UCase(adc("wRmCtg")) = "S" And Not moCn.RecSeek("select 'x' from Head " + _
                            "where HCoCd= '" + gs_CoCd + "' and HCd= '" + ctSelfCmCd + "' and HSlvModYN= 'Y'") Then _
    DispMsg "Silver Category Cannot Be Selected As Silver Module Is Disabled", etError: Exit Sub
  ' Zubin 212 (24/02/06)
  '*** Sachin *** Emr 2.11.0
    ' Zubin 212 Patch 02 (30/03/06)
    'Retrieving the SEO Password
    'ms_Pwd = moCn.GetFldVal("Select PValue1 from Param where " + "Ptyp = 'SEOPWD' and PmCd = 'SEOPWD'")
    'ms_Pwd = "MWTEAM"
    ' Zubin 212 Patch 02 (30/03/06)
  '*** Call the 'Update Platinum Code' Routine
  If Trim$(TxtPwdPltUpd.text) = "" Then
    DispMsg "Enter Support Password for this Update", etError
  ElseIf Trim$(TxtPwdPltUpd.text) <> mcs_Pwd Then
    DispMsg "Not a Valid Support Password", etError
  Else
    moCn.Execute ("Insert Into TrgDisa Values ('Param') " + vbCrLf + _
                 "Update Param Set PValue='" + adc("wRmcd") + "' where PTyp='RMCTG' and PMCd='" + adc("WRMCTG") + "' and PSCd='' " + _
                 vbCrLf + " Delete From TrgDisa ")
    TxtDisp.text = "Pure Code Updated Successfully"
    Call DispFra(PurPlt)
  End If
   '*** Sachin *** Emr 2.11.0
   Exit Sub
ErrHandler:
   If moCn.GetFldVal("Select 'x' From TrgDisa") <> 0 Then moCn.Execute ("Delete From TrgDisa")
End Sub
Private Sub CmdUpdCoCd_Click()
  '*** Display or hide the Copy frame depending on the toggle state
  Call DispFra(UpdCoCd)
  If FraUpdCoCd.Visible = True Then
    'After Clr Scrn
    'FraUpdCoCd.Caption = "Update Company Code"
    LblMain(2).Caption = "Update Company Code"
    'After Clr Scrn
    ms_Opt = "UpdCoCd"
  End If
End Sub
Private Sub CmdUpdCoCdGo_Click()
  '*** Sachin *** Emr 2.11.0
    'Retrieving the SEO Password
    ' Zubin 212 (Before 28/02/06)
    'ms_Pwd = moCn.GetFldVal("Select PValue1 from Param where " + "Ptyp = 'SEOPWD' and PmCd = 'SEOPWD'")
    'ms_Pwd = "MWTEAM"
    ' Zubin 212
  '*** Sachin *** Emr 2.11.0
  '*** Call the 'Update Company Code' Routine
  If Trim$(TxtPwd.text) = "" Then
    DispMsg "Enter Support Password for this Update", etError
  ElseIf Trim$(TxtPwd.text) <> mcs_Pwd Then
    DispMsg "Not a Valid Support Password", etError
  Else
    Call UpdateCoCd
  End If
End Sub
Private Sub CmdDelOrd_Click()
  'Exit Sub
  '*** Display or hide the Copy frame depending on the toggle state
  Call DispFra(DelOrd)
  If FraDelOrd.Visible = True Then
    ms_Opt = "DelOrd"
  End If
End Sub
Private Sub CmdDelOrdGo_Click()
  If Trim$(TxtDelOrdPwd.text) = "" Then
    DispMsg "Enter SEO Password for this Update", etError
  ElseIf Trim$(TxtDelOrdPwd.text) <> ms_Pwd Then
    DispMsg "Not a Valid SEO Password", etError
  Else
    '*** Call the Delete Order Routine
    Call DeleteOrder
  End If
End Sub
'************Geeta***************24/04/04
Private Sub CmdAvgStkRt_Click()
  '*** Display or hide the Copy frame depending on the toggle state
  Call DispFra(AvgStkRt)
  If FraAvgStkRt.Visible = True Then
    ms_Opt = "AvgStkRt"
  End If
End Sub
Private Sub CmdAvgStkRtGo_Click()
  If Not moCn.RecSeek("select 'x' from dbo.sysobjects where " + _
         "id = object_id(N'[dbo].[MwSp_205_AvgStkRt]') and " + _
         "OBJECTPROPERTY(id, N'IsProcedure') = 1") Then
    TxtDisp.text = "Procedure 'MwSp_205_AvgStkRt' Does Not Exist In The Database "
    DispMsg "Procedure 'MwSp_205_AvgStkRt' Does Not Exist In The Database ", etError
    Exit Sub
  End If
  
  DoEvents
  MWLib.BeginProcess Me, "Starting With The 205 Average Stock Rate Update"
  moCn.Execute ("MwSp_205_AvgStkRt")
  MWLib.EndProcess Me
  TxtDisp.text = "Database Has Been Successfully Updated"
  DispMsg "Database Has Been Successfully Updated", etInfo
  moCn.Execute ("drop procedure MwSp_205_AvgStkRt")
  Call CmdAvgStkRt_Click
End Sub
'************Geeta**********************

Private Sub CmdUpdCoName_Click()
  ' ****** Sachin 2.14.0 Next ******
    Call DispFra(UpdCoName)
  ' ****** Sachin 2.14.0 Next ******
End Sub
Private Sub CmdUpdIniGo_Click()
' ****** Sachin [31-08-06] Update Ini Date from Maintenance Exe ******
On Error GoTo ErrHd
Dim DtTime As String
If Trim$(TxtDllUpdPwd.text) = mcs_Pwd Then
  DtTime = CStr(Format(Time, "hhmmss"))
  'Backing up the Original Emr.ini/Sto.ini File.
  FileCopy TxtDllPath.text + IIF(CmbProj.text = "EMR", "\Emr.ini ", "\Sto.ini "), TxtDllPath.text + IIF(CmbProj.text = "EMR", "\EmrBackUp_" + DtTime + ".ini", "\StoBackUp_" + DtTime + ".ini")
  
  If SaveFile = True Then
    If TxtDllPath.text <> "" Then
      'Resetting the ModTime value to 0 after updating with new Date in Ini File
      moCn.Execute ("Update Param set PNum1= 99, ModTime= 0 where PTyp= 'USR' and PMCd= 'MIC'")
      TxtDisp.text = IIF(CmbProj.text = "EMR", "Emperor ", "Stores ") + "Dlls Checked Successfully."
    End If
  Else
    If TxtDllPath.text <> "" Then
      TxtDisp.text = "Dlls Path/File Error"
    Else
      TxtDisp.text = "Dlls Path Not Specified"
    End If
  End If
  
  Call DispFra(UpdIni)
Else
  TxtDisp.text = "Invalid Support Password Entered, Cannot Continue with Dlls Check!!"
End If
' ****** Sachin [31-08-06] Update Ini Date from Maintenance Exe ******
Exit Sub
ErrHd:
  TxtDisp.text = Err.Description
End Sub
Private Sub CmdUpdPrdPurity_Click()
  ' ****** Sachin 2.14.0 [07/10/06] - Update Production Purity For G,P,S ******
  ' Call DispFra(UpdPrdPurity)     'Sachin - 4.1.0.0 (Option Removed)
  ' ****** Sachin 2.14.0 [07/10/06] - Update Production Purity For G,P,S ******
End Sub
Private Sub CmdUpdRmZCd_Click()
 '*** Display or hide the Copy frame depending on the toggle state
  Call DispFra(UpdRmZCd)
  If FraUpdRmZCd.Visible = True Then
    ms_Opt = "UpdRmZCd"
  End If
End Sub
'***********Geeta************Emr207************
Private Sub CmdUpdRmZCdGo_Click()
On Error GoTo ErrH
Dim wStr As String
   wStr = "Insert Into TrgDisa Values ('RmMst') " + _
           "Update RmMst WITH (TABLOCKX) Set  RmCd ='" + adc("wNewRmZCd") + "' where  RmZ='Y' and RmCd='" + adc("wOldRmZCd") + "'  " + _
           " and Not Exists(Select * From RmMst where RmCd= '" + adc("wNewRmZCd") + "') " + _
           "Delete From TrgDisa " + _
           "Insert Into TrgDisa Values ('RmMst') " + _
           "Update RmMst WITH (TABLOCKX) Set  RmZCd ='" + adc("wNewRmZCd") + "' where  RmZCd='" + adc("wOldRmZCd") + "'  " + _
           "Delete From TrgDisa " + _
           "Insert Into TrgDisa Values ('Txndz') " + _
           "Update  TxndZ WITH (TABLOCKX) Set  TzRmZCd='" + adc("wNewRmZCd") + "' where   TzRmZCd='" + adc("wOldRmZCd") + "'  " + _
           "Delete From TrgDisa" + _
           "Insert Into TrgDisa Values ('InvRm') " + _
           "Update InvRm WITH (TABLOCKX) Set IrRmZCd ='" + adc("wNewRmZCd") + "' where IrRmZCd='" + adc("wOldRmZCd") + "'  " + _
           "Delete From TrgDisa " + _
           "Insert Into TrgDisa Values ('InvZRm') " + _
           "Update InvZRm WITH (TABLOCKX) Set IzRmZCd ='" + adc("wNewRmZCd") + "' where IzRmZCd='" + adc("wOldRmZCd") + "'  " + _
           "Delete From TrgDisa "
   moCn.Execute (wStr)
   TxtDisp.text = "Update Of Customs Code Successfully Completed"
Call DispFra(UpdRmZCd)
Exit Sub
ErrH:
  ws_str = "Delete From TrgDisa"
  moCn.Execute (ws_str)
End Sub
Private Sub CmdUpdSYyMm_Click()
'*** Display or hide the Copy frame depending on the toggle state
  Call DispFra(UpdSYyMm)
  If FraUpdSYyMm.Visible = True Then
    ms_Opt = "UpdSYyMm"
  End If
End Sub
Private Sub CmdUpdSYyMmGo_Click()
 Call UpdateSYyMm
 ''''Call UpdateSYyMmPart    'Temporary
 Call DispFra(UpdSYyMm)
End Sub
Private Sub CmdUpdPrdStg_Click()
'*** Display or hide the Copy frame depending on the toggle state
  Call DispFra(UpdPrdStg)
  If FraUpdPrdStg.Visible = True Then
    ms_Opt = "UpdPrdStg"
  End If
End Sub
Private Sub CmdUpdPrdStgGo_Click()
  If Trim$(TxtUpdPrdStgPwd.text) = "" Then
    DispMsg "Enter Support Password for this Update", etError
  ElseIf Trim$(TxtUpdPrdStgPwd.text) <> mcs_Pwd Then
    DispMsg "Not a Valid Support Password", etError
  Else
    Call UpdatePrdStage
    'Call DispFra(UpdPrdStg)
  End If
End Sub
Private Sub CmdUpdTrgDsgRm_Click()
  Dim Rs_DsgRm  As MDORowSet
  Dim wStrDsgSr As String
  wStrDsgSr = ""
  Set Rs_DsgRm = moCn.OpenResultset("select DrCd as qDrCd, DrSr as qDrSr from DsgRm where " + _
                     " not exists (select 'x' from RmMst where " + _
                     " RmCtg = DrRmCtg And RmSCtg = DrRmSCtg " + _
                     " and RmCd= DrRmCd and RmZ= 'N' )")
  Do While Not (Rs_DsgRm.EOF Or Rs_DsgRm.BOF)
      wStrDsgSr = wStrDsgSr + Rs_DsgRm!qDrCd + " / " + CStr(Rs_DsgRm!qDrSr) + ", "
      Rs_DsgRm.MoveNext
  Loop
  If wStrDsgSr <> "" Then
   TxtDisp.text = "Raw Material Category, SubCategory & RmCode Does Not Exists In DsgMst For the following DmCd and Sr  " + wStrDsgSr
   DispMsg "Correct The Design Master Details and Execute Update DsgRm Trigger Option Again", etError: Exit Sub
  End If
  If Not moCn.RecSeek("select 'x' from dbo.sysobjects where " + _
             "id = object_id(N'[dbo].[MwSp_UpdateDsgRmTrig]') and " + _
             "OBJECTPROPERTY(id, N'IsProcedure') = 1") Then
        TxtDisp.text = "Procedure 'MwSp_UpdateDsgRmTrig' Does Not Exist In The Database "
        DispMsg "Procedure 'MwSp_UpdateDsgRmTrig' Does Not Exist In The Database ", etError: Exit Sub
  End If
DoEvents
      MWLib.BeginProcess Me, "Starting With Design Rm Trigger Update "
      moCn.Execute ("MwSp_UpdateDsgRmTrig")
      MWLib.EndProcess Me
      TxtDisp.text = "Design Rm Trigger  Has Been Successfully Updated"
      DispMsg "Design Rm Trigger  Has Been Successfully Updated", etInfo
      moCn.Execute ("drop procedure MwSp_UpdateDsgRmTrig")
End Sub
Private Sub CmdUpdUA_Click()
    Call DispFra(UPDUA)
    If FraUpdUA.Visible = True Then
     ms_Opt = "UPDUA"
    End If
End Sub
Private Sub CmdUpdUAGo_Click()
If ms_Ver = "2.6.0" Then
    moCn.Execute ("UPDATE UsrAccess Set UaMnuForm = T.UaMnuForm  " + _
                  "From (Select UaMnuCd, UaMnuForm FROM  UsrAccess UA2 WHERE  " + _
                  "UA2.UaUsrCd = 'MIC' ) T WHERE UsrAccess.UaMnuCd = T.UaMnuCd And  " + _
                  "UsrAccess.UaUsrCd <> 'MIC' ")
Else
  TxtDisp.text = "You can Not use This Option on " + ms_Ver + " Version."
End If
  Call DispFra(UPDUA)
End Sub

'*** (Jen 3.01)
Private Sub CmdUpdRmQw_Click()
    Call DispFra(UpdRmQw)
    If FraUpdRmQw.Visible = True Then
     ms_Opt = "UpdRmQw"
    End If
End Sub
Private Sub CmdUpdRmQwGo_Click()
  Call UpdateRmQw
  If FraUpdRmQw.Visible = True Then ms_Opt = "UpdRmQw"
End Sub
'*** (Jen 3.01)

Private Sub GrdInvFgdDelRecs_ColWhen(ByVal RowNum As Integer, ByVal ColName As String, Cancel As Boolean, ErrMsg As String, ByVal Action As MwfCtl.en_fgAction)
  Cancel = True
End Sub

Private Sub RmPrdPurityUpdGo_Click()
' ****** Sachin 2.14.0 [12-10-06] Production Purity Update ******
  On Error GoTo ErrHandler
    If adc("WRMCTGPPUR") = "" Or adc("WRMCDPPUR") = "" Or (Val(adc("WRMPRDPURITY")) = 0 Or adc("WRMPRDPURITY") = "") Then
        DispMsg "RM Category, Rm Code And a Valid Production Purity Should be Entered For This Update", etError
        Exit Sub
    End If
    
    ' ***** Manali 3.9.0 - Module Check added for Silver and Palladium
    If adc("WRMCTGPPUR") = "S" And Not moCn.RecSeek("Select 'x' from Head Where HCoCd='" + gs_CoCd + "' and HCd='" + ctSelfCmCd + "' and hSlvModYN='Y' ") Then
        DispMsg "Purity Cannot be Updated as Silver Module Not available", etError
        Exit Sub
    End If
    
    If adc("WRMCTGPPUR") = "L" And Not moCn.RecSeek("Select 'x' from Head Where HCoCd='" + gs_CoCd + "' and HCd='" + ctSelfCmCd + "' and hPdModYN='Y' ") Then
        DispMsg "Purity Cannot be Updated as Palladium Module Not available", etError
        Exit Sub
    End If
    ' ***** Manali 3.9.0 - Module Check added for Silver and Palladium
    
    If Trim$(TxtPrdPurUpdPwd.text) = "" Then
      DispMsg "Enter Support Password for this Update", etError
    ElseIf Trim$(TxtPrdPurUpdPwd.text) <> mcs_Pwd Then
      DispMsg "Not a Valid Support Password", etError
    Else
      moCn.Execute (" Insert Into TrgDisa Values ('RmMst') " + vbCrLf + _
                    " Update RmMst Set RmPurityWt=" + CStr(adc("WRMPRDPURITY")) + " Where RmCtg='" + adc("WRMCTGPPUR") + "' And RmCd='" + adc("WRMCDPPUR") + "' " + _
                    vbCrLf + " Delete From TrgDisa ")
      TxtDisp.text = "Production Purity Updated Successfully"
      Call DispFra(UpdPrdPurity)
    End If
  Exit Sub
ErrHandler:
   If moCn.GetFldVal("Select 'x' From TrgDisa") <> 0 Then moCn.Execute ("Delete From TrgDisa")
   TxtDisp.text = Err.Description
' ****** Sachin 2.14.0 [12-10-06] Production Purity Update ******
End Sub

' ***** Manali - 3.5.0 - 04/06/09 - New UpdateSYyMm Proc for versions after 3.2.0
Private Sub UpdateSYyMm()
  Dim ws_str As String, ws_SqlStr As String
  On Error GoTo ErrH
' Zubin 212 (LotNo included)
'*** Jay 3.01 *** [Inv-Fg Rm Mismatch] [TrgDisa & SysDb Chk Inroduced]
'ws_Sqlstr = "Set dateformat dmy  " + _
             "delete from SYyMm "
'6.1
ws_SqlStr = " Set XACT_ABORT ON " + _
            "BEGIN TRAN UpdateSYyMm  " + _
            "set dateformat " + DtFmtStr() + " " + _
            "SET ANSI_WARNINGS OFF "
  
ws_SqlStr = ws_SqlStr + " Insert Into TrgDisa Values ('SYyMm') " + _
            "Delete from SYyMm " + _
            "Where IsNull((Select ','+SdCoCd+',' From Sysdb Where SysId='EMR'),','+SCoCd+',') Like '%,'+SCoCd+',%' " + _
            "SELECT TdCoCd, convert(int,convert(varchar(6),Tdt,112)) as qYyMm, TdFrRmLoc as qLoc, " + _
            "TdRmCd, TdLotNo,  TdRmSz, TdRmStkRt, " + _
            "sum(TdRmQty * (case when TdFrRmDc= 'D' then 1 else 0 end)) as qDrQty, " + _
            "sum(TdRmWt * (case when TdFrRmDc= 'D' then 1 else 0 end)) as qDrWt, " + _
            "sum(TdRmQty * (case when TdFrRmDc= 'C' then 1 else 0 end)) as qCrQty, " + _
            "sum(TdRmWt * (case when TdFrRmDc= 'C' then 1 else 0 end)) as qCrWt " + _
            "Into TmpSYYMM from Txnd Join Txn on TCoCd= TdCoCd and TTc= TdTc and " + _
            "TYy = TdYy And TChr = TdChr And TNo = TDNo " + _
            "where /* convert(int,convert(varchar(6),Tdt,112))> 190001 and */ " + _
            "TdRmCd<> '' and TdFrRmLoc<> '' and TdFrRmDc<> '' " + _
            "group by TdCoCd, convert(int,convert(varchar(6),Tdt,112)), TdFrRmLoc, TdRmCd, TdLotNo, TDRmSz,TdRmStkRt"


ws_SqlStr = ws_SqlStr + " insert into TmpSYyMm " + _
            "SELECT TdCoCd, convert(int,convert(varchar(6),Tdt,112)) as qYyMm, TdToRmLoc as qLoc," + _
            "TdRmCd, TdLotNo, TdRmSz, TdRmStkRt, " + _
            "sum(TdRmQty * (case when TdToRmDc= 'D' then 1 else 0 end)) as qDrQty, " + _
            "sum(TdRmWt * (case when TdToRmDc= 'D' then 1 else 0 end)) as qDrWt, " + _
            "sum(TdRmQty * (case when TdToRmDc= 'C' then 1 else 0 end)) as qCrQty, " + _
            "sum(TdRmWt * (case when TdToRmDc= 'C' then 1 else 0 end)) as qCrWt " + _
            "from Txnd Join Txn on TCoCd= TdCoCd and TTc= TdTc and TYy= TdYy and TChr= TdChr " + _
            "and TNo= TDNo where /* convert(int, convert(varchar(6), Tdt, 112))> 190001 and */ " + _
            "TdRmCd<> '' and TdToRmLoc<> '' and TdToRmDc<> '' " + _
            "group by TdCoCd, convert(int,convert(varchar(6),Tdt,112)), TdToRmLoc, TdRmCd, TdLotNo, TDRmSz,TdRmStkRt "

ws_SqlStr = ws_SqlStr + " insert into TmpSYyMm " + _
            "SELECT FrCoCd as TdCoCd, convert(int,convert(varchar(6),Fgdt,112)) as qYyMm, FrFrRmLoc as qLoc," + _
            "FrRmCd as TdRmCd, FrLotNo as TdLotNo, FrRmSz as TdRmSz, FrRmStkRt as TdRmStkRt, " + _
            "sum(FrRmQty * (case when FrFrRmDc= 'D' then 1 else 0 end)) as qDrQty, " + _
            "sum(FrRmWt * (case when FrFrRmDc= 'D' then 1 else 0 end)) as qDrWt, " + _
            "sum(FrRmQty * (case when FrFrRmDc= 'C' then 1 else 0 end)) as qCrQty, " + _
            "sum(FrRmWt * (case when FrFrRmDc= 'C' then 1 else 0 end)) as qCrWt " + _
            "from FgRm Join Fg on FgCoCd= FrCoCd and FgTc= FrTc and fgYy= FrYy and FgChr= FrChr " + _
            "and FgNo= FrNo /* where convert(int, convert(varchar(6), FgDt, 112))> 190001 */ " + _
            "group by FrCoCd, convert(int,convert(varchar(6), Fgdt,112)), FrFrRmLoc, FrRmCd, FrLotNo, FrRmSz,FrRmStkRt "


ws_SqlStr = ws_SqlStr + " insert into TmpSYyMm " + _
            "SELECT FrCoCd, convert(int,convert(varchar(6),Fgdt,112)) as qYyMm, FrToRmLoc as qLoc," + _
            "FrRmCd, FrLotNo, FrRmSz, FrRmStkRt, " + _
            "sum(FrRmQty * (case when FrToRmDc= 'D' then 1 else 0 end)) as qDrQty, " + _
            "sum(FrRmWt * (case when FrToRmDc= 'D' then 1 else 0 end)) as qDrWt, " + _
            "sum(FrRmQty * (case when FrToRmDc= 'C' then 1 else 0 end)) as qCrQty, " + _
            "sum(FrRmWt * (case when FrToRmDc= 'C' then 1 else 0 end)) as qCrWt " + _
            "from FgRm Join Fg on FgCoCd= FrCoCd and FgTc= FrTc and fgYy= FrYy and FgChr= FrChr " + _
            "and FgNo= FrNo /* where convert(int, convert(varchar(6), FgDt, 112))> 190001 */ " + _
            "group by FrCoCd, convert(int,convert(varchar(6), Fgdt,112)), FrToRmLoc, FrRmCd, FrLotNo, FrRmSz,FrRmStkRt "



ws_SqlStr = ws_SqlStr + " insert into TmpSYyMm " + _
            "SELECT IrCoCd, convert(int, convert(varchar(6), InDt, 112)) as qYyMm, IrFrRmLoc as qLoc, " + _
            "IrRmCd, IrLotNo, IrRmSz, IrStkRt, " + _
            "sum(IrRmQty * (case when IrFrRmDc= 'D' then 1 else 0 end)) as qDrQty, " + _
            "sum(IrRmAWt * (case when IrFrRmDc= 'D' then 1 else 0 end)) as qDrWt, " + _
            "sum(IrRmQty * (case when IrFrRmDc= 'C' then 1 else 0 end)) as qCrQty, " + _
            "sum(IrRmAWt * (case when IrFrRmDc= 'C' then 1 else 0 end)) as qCrWt " + _
            "from InvRm join InvHd on IrCoCd= InCoCd and IrTc= InTc and IrYy= InYy and IrChr= InChr " + _
            "and IrNo= InNo /* where convert(int, convert(varchar(6), InDt, 112))> 190001 */ " + _
            "group by IrCoCd, convert(int,convert(varchar(6), InDt,112)), IrFrRmLoc, IrRmCd, IrLotNo, IrRmSz,IrStkRt "



ws_SqlStr = ws_SqlStr + " insert into TmpSYyMm " + _
            "SELECT IrCoCd, convert(int, convert(varchar(6), InDt, 112)) as qYyMm, IrToRmLoc as qLoc, " + _
            "IrRmCd, IrLotNo, IrRmSz, IrStkRt, " + _
            "sum(IrRmQty * (case when IrToRmDc= 'D' then 1 else 0 end)) as qDrQty, " + _
            "sum(IrRmAWt * (case when IrToRmDc= 'D' then 1 else 0 end)) as qDrWt, " + _
            "sum(IrRmQty * (case when IrToRmDc= 'C' then 1 else 0 end)) as qCrQty, " + _
            "sum(IrRmAWt * (case when IrToRmDc= 'C' then 1 else 0 end)) as qCrWt " + _
            "from InvRm join InvHd on IrCoCd= InCoCd and IrTc= InTc and IrYy= InYy and IrChr= InChr and IrNo= InNo " + _
            "/* where convert(int, convert(varchar(6), InDt, 112))> 190001 */ " + _
            "group by IrCoCd, convert(int,convert(varchar(6), InDt,112)), IrToRmLoc, IrRmCd, IrLotNo, IrRmSz,IrStkRt "

          
' ***** Manali 3.9.0 - Update For Trading - Afet 3.7.0
ws_SqlStr = ws_SqlStr + " insert into TmpSYyMm " + _
            "SELECT OmCoCd as TdCoCd, convert(int,convert(varchar(6),OmDt,112)) as qYyMm, OmFrRmLoc as qLoc, " + _
            "FrRmCd as TdRmCd, FrLotNo as TdLotNo, FrRmSz as TdRmSz, FrRmStkRt as TdRmStkRt, " + _
            "sum(FrRmQty * (case when OmFrRmDc= 'D' then 1 else 0 end)) as qDrQty, " + _
            "sum(FrRmWt * (case when OmFrRmDc= 'D' then 1 else 0 end)) as qDrWt, " + _
            "sum(FrRmQty * (case when OmFrRmDc= 'C' then 1 else 0 end)) as qCrQty, " + _
            "sum(FrRmWt * (case when OmFrRmDc= 'C' then 1 else 0 end)) as qCrWt " + _
            "From (Select FdCoCd, FdbYy, FdBChr, FdBNo, FdBIdNo, FdPrtKey, FrRmCd, FrLotNo, FrRmSz,FrRmStkRt, " + _
            "  Sum(FrRmWt * (Case When FrFrRmLoc='FG' and FrFrRmDc= 'D' Then 1 " + _
            "           When FrFrRmLoc='FG' and FrFrRmDc= 'C' Then -1 " + _
            "           When FrToRmLoc='FG' and FrToRmDc= 'D' Then 1 " + _
            "           When FrToRmLoc='FG' and FrToRmDc= 'C' Then -1 Else 0 End)) FrRmWt, " + _
            "   Sum(FrRmQty * (Case When FrFrRmLoc='FG' and FrFrRmDc= 'D' Then 1 " + _
                        "When FrFrRmLoc='FG' and FrFrRmDc= 'C' Then -1 " + _
                        "When FrToRmLoc='FG' and FrToRmDc= 'D' Then 1 " + _
                        "When FrToRmLoc='FG' and FrToRmDc= 'C' Then -1 Else 0 End)) FrRmQty " + _
            "From FgRm " + _
            "Join Fgd on FdCoCd= FrCoCd and FdTc= FrTc and FdYy= FrYy and FdChr= FrChr and FdNo= FrNo and FdSr=FrSr and FrPrtKey=FdPrtKey " + _
            "Group By FdCoCd, FdbYy, FdBChr, FdBNo, FdBIdNo, FdPrtKey, FrRmCd, FrLotNo, FrRmSz,FrRmStkRt) Fgd " + _
            "Join OrdDsg On OdCoCd=FdCoCd and OdBYy=FdBYy and OdBChr=FdBChr and OdBNo=FdBNo " + _
            "Join OrdMst On OmIdNo=OdOmIdNo and OmCoCd=OdCoCd and OmTc=OdTc and OmYy=OdYy and OmChr=OdChr and OmNo=OdNo " + _
            "Where OmFrRmLoc<> '' and OmFrRmDc<> '' " + _
            "group by OmCoCd, convert(int,convert(varchar(6), OmDt,112)), OmFrRmLoc, FrRmCd, FrLotNo, FrRmSz,FrRmStkRt "
          
ws_SqlStr = ws_SqlStr + " insert into TmpSYyMm " + _
            "SELECT OmCoCd as TdCoCd, convert(int,convert(varchar(6),OmDt,112)) as qYyMm, OmToRmLoc as qLoc, " + _
            "FrRmCd as TdRmCd, FrLotNo as TdLotNo, FrRmSz as TdRmSz, FrRmStkRt as TdRmStkRt, " + _
            "sum(FrRmQty * (case when OmToRmDc= 'D' then 1 else 0 end)) as qDrQty, " + _
            "sum(FrRmWt * (case when OmToRmDc= 'D' then 1 else 0 end)) as qDrWt, " + _
            "sum(FrRmQty * (case when OmToRmDc= 'C' then 1 else 0 end)) as qCrQty, " + _
            "sum(FrRmWt * (case when OmToRmDc= 'C' then 1 else 0 end)) as qCrWt " + _
            "From " + _
            "(Select FdCoCd, FdbYy, FdBChr, FdBNo, FdBIdNo, FdPrtKey, FrRmCd, FrLotNo, FrRmSz,FrRmStkRt, " + _
            "  Sum(FrRmWt * (Case When FrFrRmLoc='FG' and FrFrRmDc= 'D' Then 1 " + _
            "            When FrFrRmLoc='FG' and FrFrRmDc= 'C' Then -1 " + _
            "            When FrToRmLoc='FG' and FrToRmDc= 'D' Then 1 " + _
            "            When FrToRmLoc='FG' and FrToRmDc= 'C' Then -1 Else 0 End)) FrRmWt, " + _
            "  Sum(FrRmQty * (Case When FrFrRmLoc='FG' and FrFrRmDc= 'D' Then 1 " + _
            "            When FrFrRmLoc='FG' and FrFrRmDc= 'C' Then -1 " + _
            "            When FrToRmLoc='FG' and FrToRmDc= 'D' Then 1 " + _
            "            When FrToRmLoc='FG' and FrToRmDc= 'C' Then -1 Else 0 End)) FrRmQty " + _
            "From FgRm " + _
            "Join Fgd on FdCoCd= FrCoCd and FdTc= FrTc and FdYy= FrYy and FdChr= FrChr and FdNo= FrNo and FdSr=FrSr and FrPrtKey=FdPrtKey " + _
            "Group By FdCoCd, FdbYy, FdBChr, FdBNo, FdBIdNo, FdPrtKey, FrRmCd, FrLotNo, FrRmSz,FrRmStkRt ) " + _
            "Fgd Join OrdDsg On OdCoCd=FdCoCd and OdBYy=FdBYy and OdBChr=FdBChr and OdBNo=FdBNo " + _
            "Join OrdMst On OmIdNo=OdOmIdNo and OmCoCd=OdCoCd and OmTc=OdTc and OmYy=OdYy and OmChr=OdChr and OmNo=OdNo " + _
            "Where OmToRmLoc<> '' and OmToRmDc<> '' " + _
            "group by OmCoCd, convert(int,convert(varchar(6), OmDt,112)), OmToRmLoc, FrRmCd, FrLotNo, FrRmSz,FrRmStkRt "
          
          
'*** Jay 3.01 *** [Inv-Fg Rm Mismatch] [TrgDisa & SysDb Chk Inroduced]
ws_SqlStr = ws_SqlStr + " Insert Into SYyMm(SCoCd, SYyMm, SLoc, SRmCd, SLotNo, SRmSz, SRmStkRt, SRmDrQty, SRmDrWt, SRmCrQty, SRmCrWt,SPrtKey) " + _
            "SELECT TdCoCd, qYyMm, qLoc, TdRmCd, TdLotNo, TdRmSz, TdRmStkRt," + _
            "  Sum(qDrQty) as qDrQty, Sum(qDrWt) as qDrWt, Sum(qCrQty) as qCrQty, Sum(qCrWt) as qCrWt , 'C' " + _
            "From TmpSYyMm " + _
            "Where IsNull((Select ','+SdCoCd+',' From Sysdb Where SysId='EMR'),','+TdCoCd+',') Like '%,'+TdCoCd+',%' " + _
            "Group BY TdCoCd, qYyMm, qLoc, TdRmCd, TdLotNo, TdRmSz, TdRmStkRt " + _
            "ORDER BY TdCoCd, qYyMm, qLoc, TdRmCd, TdLotNo, TdRmSz, TdRmStkRt " + _
            "Delete From TrgDisa "

 
 ws_SqlStr = ws_SqlStr + " SET ANSI_WARNINGS ON " + _
                         "COMMIT TRAN UpdateSYyMm " + _
                         "Set XACT_ABORT OFF "
DoEvents
      MWLib.BeginProcess Me, "Starting With SYyMm Update"
        moCn.Execute ("if exists (select * from dbo.sysobjects where id = object_id(N'[TmpSYyMm]') and OBJECTPROPERTY(id, N'IsUserTable') = 1)  " + _
             "drop table [TmpSYyMm] ")
        moCn.Execute (ws_SqlStr)
        moCn.Execute ("if exists (select * from dbo.sysobjects where id = object_id(N' [TmpSYyMm]') and OBJECTPROPERTY(id, N'IsUserTable') = 1) " + _
             " drop table [TmpSYyMm] ")
 MWLib.EndProcess Me

  TxtDisp.text = "SYyMm Has Been Successfully Updated"
  DispMsg "SYyMm Has Been Successfully Updated", etInfo
  'Call CmdUpdSYyMm_Click     '*** Jay 3.01 *** [Inv-Fg Rm Mismatch]
  Exit Sub
ErrH:
  Call ErrHndlr
  MWLib.EndProcess Me
  'Call CmdUpdSYyMm_Click     '*** Jay 3.01 *** [Inv-Fg Rm Mismatch]
End Sub


Private Sub UpdateSYyMm_Bef320()
''  Dim ws_str As String, ws_SqlStr As String
''  On Error GoTo ErrH
''' Zubin 212 (LotNo included)
'''*** Jay 3.01 *** [Inv-Fg Rm Mismatch] [TrgDisa & SysDb Chk Inroduced]
'''ws_Sqlstr = "Set dateformat dmy  " + _
''             "delete from SYyMm "
''
''ws_SqlStr = " Set XACT_ABORT ON " + _
''            "BEGIN TRAN UpdateSYyMm  " + _
''            "set dateformat dmy " + _
''            "SET ANSI_WARNINGS OFF "
''
''ws_SqlStr = ws_SqlStr + " Insert Into TrgDisa Values ('SYyMm') " + _
''            "Delete from SYyMm " + _
''            "Where IsNull((Select ','+SdCoCd+',' From Sysdb Where SysId='EMR'),','+SCoCd+',') Like '%,'+SCoCd+',%' " + _
''            "SELECT TdCoCd, convert(int,convert(varchar(6),Tdt,112)) as qYyMm, TdFrRmLoc as qLoc, " + _
''            "TdRmCd, TdLotNo, TdRmSz, TdRmStkRt, " + _
''            "sum(TdRmQty * (case when TdFrRmDc= 'D' then 1 else 0 end)) as qDrQty, " + _
''            "sum(TdRmWt * (case when TdFrRmDc= 'D' then 1 else 0 end)) as qDrWt, " + _
''            "sum(TdRmQty * (case when TdFrRmDc= 'C' then 1 else 0 end)) as qCrQty, " + _
''            "sum(TdRmWt * (case when TdFrRmDc= 'C' then 1 else 0 end)) as qCrWt " + _
''            "Into TmpSYYMM from Txnd Join Txn on TCoCd= TdCoCd and TTc= TdTc and  " + _
''            "TYy= TdYy and TChr= TdChr   and TNo= TDNo " + _
''            "where /* convert(int,convert(varchar(6),Tdt,112))> 190001 and */ " + _
''            "  TdRmCd<> '' and TdFrRmLoc<> '' and TdFrRmDc<> '' " + _
''            "group by TdCoCd, convert(int,convert(varchar(6),Tdt,112)), TdFrRmLoc, TdRmCd, TdLotNo, TDRmSz,TdRmStkRt  "
''
''ws_SqlStr = ws_SqlStr + " insert into TmpSYyMm " + _
''          "SELECT TdCoCd, convert(int,convert(varchar(6),Tdt,112)) as qYyMm, TdToRmLoc as qLoc, " + _
''          "TdRmCd, TdLotNo, TdRmSz, TdRmStkRt,sum(TdRmQty * (case when TdToRmDc= 'D' then 1 else 0 end)) as qDrQty, " + _
''          "sum(TdRmWt * (case when TdToRmDc= 'D' then 1 else 0 end)) as qDrWt, " + _
''          " sum(TdRmQty * (case when TdToRmDc= 'C' then 1 else 0 end)) as qCrQty, " + _
''          " sum(TdRmWt * (case when TdToRmDc= 'C' then 1 else 0 end)) as qCrWt " + _
''          " from Txnd Join Txn on TCoCd= TdCoCd and TTc= TdTc and TYy= TdYy and TChr= TdChr " + _
''          " and TNo= TDNo where /* convert(int, convert(varchar(6), Tdt, 112))> 190001 and */ " + _
''          "TdRmCd<> '' and TdToRmLoc<> '' and TdToRmDc<> '' " + _
''          "group by TdCoCd, convert(int,convert(varchar(6),Tdt,112)), TdToRmLoc, TdRmCd, TdLotNo, TDRmSz,TdRmStkRt  "
''
''ws_SqlStr = ws_SqlStr + " insert into TmpSYyMm " + _
''          "SELECT FrCoCd as TdCoCd, convert(int,convert(varchar(6),Fgdt,112)) as qYyMm, FrFrRmLoc as qLoc, " + _
''          " FrRmCd as TdRmCd, FrLotNo as TdLotNo, FrRmSz as TdRmSz, FrRmStkRt as TdRmStkRt," + _
''          " sum(FrRmQty * (case when FrFrRmDc= 'D' then 1 else 0 end)) as qDrQty, " + _
''          " sum(FrRmWt * (case when FrFrRmDc= 'D' then 1 else 0 end)) as qDrWt, " + _
''          " sum(FrRmQty * (case when FrFrRmDc= 'C' then 1 else 0 end)) as qCrQty, " + _
''          " sum(FrRmWt * (case when FrFrRmDc= 'C' then 1 else 0 end)) as qCrWt " + _
''          " from FgRm Join Fg on FgCoCd= FrCoCd and FgTc= FrTc and fgYy= FrYy and FgChr= FrChr " + _
''          " and FgNo= FrNo /* where convert(int, convert(varchar(6), FgDt, 112))> 190001 */ " + _
''          " group by FrCoCd, convert(int,convert(varchar(6), Fgdt,112)), FrFrRmLoc, FrRmCd, FrLotNo, FrRmSz,FrRmStkRt  "
''
''ws_SqlStr = ws_SqlStr + " insert into TmpSYyMm " + _
''          " SELECT FrCoCd, convert(int,convert(varchar(6),Fgdt,112)) as qYyMm, FrToRmLoc as qLoc, " + _
''          " FrRmCd, FrLotNo, FrRmSz, FrRmStkRt,sum(FrRmQty * (case when FrToRmDc= 'D' then 1 else 0 end)) as qDrQty," + _
''          " sum(FrRmWt * (case when FrToRmDc= 'D' then 1 else 0 end)) as qDrWt, " + _
''          " sum(FrRmQty * (case when FrToRmDc= 'C' then 1 else 0 end)) as qCrQty, " + _
''          " sum(FrRmWt * (case when FrToRmDc= 'C' then 1 else 0 end)) as qCrWt  " + _
''          " from FgRm Join Fg on FgCoCd= FrCoCd and FgTc= FrTc and fgYy= FrYy and FgChr= FrChr " + _
''          " and FgNo= FrNo /* where convert(int, convert(varchar(6), FgDt, 112))> 190001 */ " + _
''          " group by FrCoCd, convert(int,convert(varchar(6), Fgdt,112)), FrToRmLoc, FrRmCd, FrLotNo, FrRmSz,FrRmStkRt  "
''
''ws_SqlStr = ws_SqlStr + " insert into TmpSYyMm " + _
''          " SELECT IrCoCd, convert(int, convert(varchar(6), InDt, 112)) as qYyMm, IrFrRmLoc as qLoc, " + _
''          " IrRmCd, IrLotNo, IrRmSz, IrStkRt,sum(IrRmQty * (case when IrFrRmDc= 'D' then 1 else 0 end)) as qDrQty, " + _
''          " sum(IrRmAWt * (case when IrFrRmDc= 'D' then 1 else 0 end)) as qDrWt, " + _
''          " sum(IrRmQty * (case when IrFrRmDc= 'C' then 1 else 0 end)) as qCrQty, " + _
''          " sum(IrRmAWt * (case when IrFrRmDc= 'C' then 1 else 0 end)) as qCrWt  " + _
''          " from InvRm join InvHd on IrCoCd= InCoCd and IrTc= InTc and IrYy= InYy and IrChr= InChr  " + _
''          " and IrNo= InNo /* where convert(int, convert(varchar(6), InDt, 112))> 190001 */ " + _
''          " group by IrCoCd, convert(int,convert(varchar(6), InDt,112)), IrFrRmLoc, IrRmCd, IrLotNo, IrRmSz,IrStkRt "
''
''ws_SqlStr = ws_SqlStr + " insert into TmpSYyMm " + _
''          "SELECT IrCoCd, convert(int, convert(varchar(6), InDt, 112)) as qYyMm, IrToRmLoc as qLoc," + _
''          "IrRmCd, IrLotNo, IrRmSz, IrStkRt,sum(IrRmQty * (case when IrToRmDc= 'D' then 1 else 0 end)) as qDrQty," + _
''          "sum(IrRmAWt * (case when IrToRmDc= 'D' then 1 else 0 end)) as qDrWt, " + _
''          "sum(IrRmQty * (case when IrToRmDc= 'C' then 1 else 0 end)) as qCrQty, " + _
''          "sum(IrRmAWt * (case when IrToRmDc= 'C' then 1 else 0 end)) as qCrWt  " + _
''          "from InvRm join InvHd on IrCoCd= InCoCd and IrTc= InTc and IrYy= InYy and IrChr= InChr and IrNo= InNo  " + _
''          "/* where convert(int, convert(varchar(6), InDt, 112))> 190001 */ " + _
''          "group by IrCoCd, convert(int,convert(varchar(6), InDt,112)), IrToRmLoc, IrRmCd, IrLotNo, IrRmSz,IrStkRt "
''
'''*** Jay 3.01 *** [Inv-Fg Rm Mismatch] [TrgDisa & SysDb Chk Inroduced]
''ws_SqlStr = ws_SqlStr + " insert into SYyMm(SCoCd, SYyMm, SLoc, SRmCd, SLotNo, SRmSz, SRmStkRt, SRmDrQty, SRmDrWt, SRmCrQty, SRmCrWt) " + _
''          "SELECT TdCoCd, qYyMm, qLoc, TdRmCd, TdLotNo, TdRmSz, TdRmStkRt," + _
''          "sum(qDrQty) as qDrQty, sum(qDrWt) as qDrWt, sum(qCrQty) as qCrQty, sum(qCrWt) as qCrWt " + _
''          "From TmpSYyMm " + _
''          "Where IsNull((Select ','+SdCoCd+',' From Sysdb Where SysId='EMR'),','+TdCoCd+',') Like '%,'+TdCoCd+',%' " + _
''          "Group BY TdCoCd, qYyMm, qLoc, TdRmCd, TdLotNo, TdRmSz, TdRmStkRt " + _
''          "ORDER BY TdCoCd, qYyMm, qLoc, TdRmCd, TdLotNo, TdRmSz, TdRmStkRt " + _
''          "Delete From TrgDisa "
''
'' ws_SqlStr = ws_SqlStr + " SET ANSI_WARNINGS ON " + _
''                         "COMMIT TRAN UpdateSYyMm " + _
''                         "Set XACT_ABORT OFF "
''DoEvents
''      MWLib.BeginProcess Me, "Starting With SYyMm Update"
''        moCn.Execute ("if exists (select * from dbo.sysobjects where id = object_id(N'[TmpSYyMm]') and OBJECTPROPERTY(id, N'IsUserTable') = 1)  " + _
''             "drop table [TmpSYyMm] ")
''        moCn.Execute (ws_SqlStr)
''        moCn.Execute ("if exists (select * from dbo.sysobjects where id = object_id(N' [TmpSYyMm]') and OBJECTPROPERTY(id, N'IsUserTable') = 1) " + _
''             " drop table [TmpSYyMm] ")
'' MWLib.EndProcess Me
''
''  TxtDisp.text = "SYyMm Has Been Successfully Updated"
''  DispMsg "SYyMm Has Been Successfully Updated", etInfo
''  'Call CmdUpdSYyMm_Click     '*** Jay 3.01 *** [Inv-Fg Rm Mismatch]
''  Exit Sub
''ErrH:
''  Call ErrHndlr
''  MWLib.EndProcess Me
''  'Call CmdUpdSYyMm_Click     '*** Jay 3.01 *** [Inv-Fg Rm Mismatch]
End Sub

Private Sub UpdateSYyMmPart()
  '*** (This procedure is for temporary use.) This can be used to reverse the effect of Txns
  Dim ws_str As String, ws_SqlStr As String
  On Error GoTo ErrH
  ' Zubin 212 (Lot No. included)
  '6.1
ws_SqlStr = "Set dateformat " + DtFmtStr() + "  " + _
             "delete from SYyMm where SYyMm= 200505 " + _
             " SELECT TdCoCd, convert(int,convert(varchar(6),Tdt,112)) as qYyMm, TdFrRmLoc as qLoc, " + _
             " TdRmCd, TdLotNo, TdRmSz, TdRmStkRt, " + _
             " sum(TdRmQty * (case when TdFrRmDc= 'D' then 1 else 0 end)) as qDrQty, " + _
             " sum(TdRmWt * (case when TdFrRmDc= 'D' then 1 else 0 end)) as qDrWt, " + _
             " sum(TdRmQty * (case when TdFrRmDc= 'C' then 1 else 0 end)) as qCrQty, " + _
             " sum(TdRmWt * (case when TdFrRmDc= 'C' then 1 else 0 end)) as qCrWt " + _
             "Into TmpSYYMM from Txnd Join Txn on TCoCd= TdCoCd and TTc= TdTc and  " + _
             "TYy= TdYy and TChr= TdChr   and TNo= TDNo " + _
             "where convert(int, convert(varchar(8), Tdt, 112))>= 20050501 and " + _
             " convert(int, convert(varchar(8), Tdt, 112))<= 20050515 and " + _
             "  TdRmCd<> '' and TdFrRmLoc<> '' and TdFrRmDc<> '' " + _
             "group by TdCoCd, convert(int,convert(varchar(6),Tdt,112)), TdFrRmLoc, TdRmCd, TdLotNo, TDRmSz,TdRmStkRt  "

ws_SqlStr = ws_SqlStr + " insert into TmpSYyMm " + _
         "SELECT TdCoCd, convert(int,convert(varchar(6),Tdt,112)) as qYyMm, TdToRmLoc as qLoc, " + _
         "TdRmCd, TdLotNo, TdRmSz, TdRmStkRt,sum(TdRmQty * (case when TdToRmDc= 'D' then 1 else 0 end)) as qDrQty, " + _
         "sum(TdRmWt * (case when TdToRmDc= 'D' then 1 else 0 end)) as qDrWt, " + _
         " sum(TdRmQty * (case when TdToRmDc= 'C' then 1 else 0 end)) as qCrQty, " + _
         " sum(TdRmWt * (case when TdToRmDc= 'C' then 1 else 0 end)) as qCrWt " + _
         "from Txnd Join Txn on TCoCd= TdCoCd and TTc= TdTc and TYy= TdYy and TChr= TdChr " + _
         " and TNo= TDNo where convert(int, convert(varchar(8), Tdt, 112))>= 20050501 and " + _
         " convert(int, convert(varchar(8), Tdt, 112))<= 20050515 and " + _
         "TdRmCd<> '' and TdToRmLoc<> '' and TdToRmDc<> '' " + _
         "group by TdCoCd, convert(int,convert(varchar(6),Tdt,112)), TdToRmLoc, TdRmCd, TdLotNo, TDRmSz,TdRmStkRt  "

ws_SqlStr = ws_SqlStr + " insert into TmpSYyMm " + _
          "SELECT FrCoCd as TdCoCd, convert(int,convert(varchar(6),Fgdt,112)) as qYyMm, FrFrRmLoc as qLoc, " + _
          " FrRmCd as TdRmCd, FrLotNo as TdLotNo, FrRmSz as TdRmSz, FrRmStkRt as TdRmStkRt," + _
          " sum(FrRmQty * (case when FrFrRmDc= 'D' then 1 else 0 end)) as qDrQty, " + _
          " sum(FrRmWt * (case when FrFrRmDc= 'D' then 1 else 0 end)) as qDrWt, " + _
          " sum(FrRmQty * (case when FrFrRmDc= 'C' then 1 else 0 end)) as qCrQty, " + _
          " sum(FrRmWt * (case when FrFrRmDc= 'C' then 1 else 0 end)) as qCrWt " + _
          " from FgRm Join Fg on FgCoCd= FrCoCd and FgTc= FrTc and fgYy= FrYy and FgChr= FrChr " + _
          " and FgNo= FrNo where convert(int, convert(varchar(8), FgDt, 112))>= 20050501 " + _
          " and convert(int, convert(varchar(8), FgDt, 112))<= 20050515 " + _
          " group by FrCoCd, convert(int,convert(varchar(6), Fgdt,112)), FrFrRmLoc, FrRmCd, FrLotNo, FrRmSz,FrRmStkRt  "


ws_SqlStr = ws_SqlStr + "insert into TmpSYyMm " + _
          " SELECT FrCoCd, convert(int,convert(varchar(6),Fgdt,112)) as qYyMm, FrToRmLoc as qLoc, " + _
          " FrRmCd, FrLotNo, FrRmSz, FrRmStkRt,sum(FrRmQty * (case when FrToRmDc= 'D' then 1 else 0 end)) as qDrQty," + _
          " sum(FrRmWt * (case when FrToRmDc= 'D' then 1 else 0 end)) as qDrWt, " + _
          " sum(FrRmQty * (case when FrToRmDc= 'C' then 1 else 0 end)) as qCrQty, " + _
          " sum(FrRmWt * (case when FrToRmDc= 'C' then 1 else 0 end)) as qCrWt  " + _
          " from FgRm Join Fg on FgCoCd= FrCoCd and FgTc= FrTc and fgYy= FrYy and FgChr= FrChr " + _
          " and FgNo= FrNo where convert(int, convert(varchar(8), FgDt, 112))>= 20050501 " + _
          " and convert(int, convert(varchar(8), FgDt, 112))<= 20050515 " + _
          " group by FrCoCd, convert(int,convert(varchar(6), Fgdt,112)), FrToRmLoc, FrRmCd, FrLotNo, FrRmSz,FrRmStkRt  "

ws_SqlStr = ws_SqlStr + "insert into TmpSYyMm " + _
          " SELECT IrCoCd, convert(int, convert(varchar(6), InDt, 112)) as qYyMm, IrFrRmLoc as qLoc, " + _
          " IrRmCd, IrLotNo, IrRmSz, IrStkRt,sum(IrRmQty * (case when IrFrRmDc= 'D' then 1 else 0 end)) as qDrQty, " + _
          " sum(IrRmAWt * (case when IrFrRmDc= 'D' then 1 else 0 end)) as qDrWt, " + _
          " sum(IrRmQty * (case when IrFrRmDc= 'C' then 1 else 0 end)) as qCrQty, " + _
          " sum(IrRmAWt * (case when IrFrRmDc= 'C' then 1 else 0 end)) as qCrWt  " + _
          " from InvRm join InvHd on IrCoCd= InCoCd and IrTc= InTc and IrYy= InYy and IrChr= InChr  " + _
          " and IrNo= InNo where convert(int, convert(varchar(8), InDt, 112))>= 20050501 " + _
          " and convert(int, convert(varchar(8), InDt, 112))<= 20050515 " + _
          " group by IrCoCd, convert(int,convert(varchar(6), InDt,112)), IrFrRmLoc, IrRmCd, IrLotNo, IrRmSz,IrStkRt "

ws_SqlStr = ws_SqlStr + "insert into TmpSYyMm " + _
          "SELECT IrCoCd, convert(int, convert(varchar(6), InDt, 112)) as qYyMm, IrToRmLoc as qLoc," + _
          "IrRmCd, IrLotNo, IrRmSz, IrStkRt,sum(IrRmQty * (case when IrToRmDc= 'D' then 1 else 0 end)) as qDrQty," + _
          "sum(IrRmAWt * (case when IrToRmDc= 'D' then 1 else 0 end)) as qDrWt, " + _
          "sum(IrRmQty * (case when IrToRmDc= 'C' then 1 else 0 end)) as qCrQty, " + _
          "sum(IrRmAWt * (case when IrToRmDc= 'C' then 1 else 0 end)) as qCrWt  " + _
          "from InvRm join InvHd on IrCoCd= InCoCd and IrTc= InTc and IrYy= InYy and IrChr= InChr and IrNo= InNo  " + _
          " where convert(int, convert(varchar(8), InDt, 112))>= 20050501 " + _
          " and convert(int, convert(varchar(8), InDt, 112))<= 20050515 " + _
          "group by IrCoCd, convert(int,convert(varchar(6), InDt,112)), IrToRmLoc, IrRmCd, IrLotNo, IrRmSz,IrStkRt "
ws_SqlStr = ws_SqlStr + "insert into SYyMm(SCoCd, SYyMm, SLoc, SRmCd, SLotNo, SRmSz, SRmStkRt, SRmDrQty, SRmDrWt, SRmCrQty, SRmCrWt) " + _
          "SELECT TdCoCd, qYyMm, qLoc, TdRmCd, TdLotNo, TdRmSz, TdRmStkRt," + _
          "sum(qDrQty) as qDrQty, sum(qDrWt) as qDrWt, sum(qCrQty) as qCrQty, sum(qCrWt) as qCrWt " + _
          "From TmpSYyMm group BY TdCoCd, qYyMm, qLoc, TdRmCd, TdLotNo, TdRmSz, TdRmStkRt " + _
          "ORDER BY TdCoCd, qYyMm, qLoc, TdRmCd, TdLotNo, TdRmSz, TdRmStkRt "
DoEvents
      MWLib.BeginProcess Me, "Starting With SYyMm Update"
        moCn.Execute ("if exists (select * from dbo.sysobjects where id = object_id(N'[TmpSYyMm]') and OBJECTPROPERTY(id, N'IsUserTable') = 1)  " + _
             "drop table [TmpSYyMm] ")
        moCn.Execute (ws_SqlStr)
        moCn.Execute ("if exists (select * from dbo.sysobjects where id = object_id(N' [TmpSYyMm]') and OBJECTPROPERTY(id, N'IsUserTable') = 1) " + _
                  " drop table [TmpSYyMm] ")
 MWLib.EndProcess Me

  TxtDisp.text = "SYyMm Has Been Successfully Updated"
  DispMsg "SYyMm Has Been Successfully Updated", etInfo
  Call CmdUpdSYyMm_Click
  Exit Sub
ErrH:
  Call ErrHndlr
  MWLib.EndProcess Me
  Call CmdUpdSYyMm_Click
End Sub
Private Sub UpdatePrdStage()
  '*** 2.11
  Dim ws_str As String
  On Error GoTo ErrH

  TxtDisp.text = ""
  DispMsg "", etInfo

  If UCase(moCn.GetFldVal("Select hPpcYN from Head where HCoCd= 'ZZZ' and HCd= 'ZSELF'")) = "N" Then
    TxtDisp.text = "PPC Module Has To Be Enabled in Head Table For This Option"
    DispMsg "PPC Module Has To Be Enabled in Head Table For This Option", etError
    Exit Sub
  End If

  If moCn.GetFldVal("Select max(LocPrdStg) From Loc where LocTyp= 'R'") = "" Then
    TxtDisp.text = "The LocPrdStg Field In The Location Master Table For LocType = 'R' Has Not Been Updated"
    DispMsg "The LocPrdStg Field In The Location Master Table For LocType = 'R' Has Not Been Updated", etError
    Exit Sub
  End If

  DoEvents
  MWLib.BeginProcess Me, "Starting With Production Stage Update"
  moCn.BeginTrans (mdoCommitRec)

  '*** (Jen 2.14 Next)
  ws_str = " if IsNull((Select hPpcYN from Head where HCoCd= 'ZZZ' and HCd= 'ZSELF'), '') = 'Y' " + _
           "Begin " + _
           "    update Bag set BPrdStg= IsNull((Select PMCd from Param P1 where PTyp= 'PRDSTG' " + _
           "           and PNum= IsNull((Select max(P2.PNum) from Param P2 where P2.PTyp= 'PRDSTG' " + _
           "                            and P2.PNum< IsNull((Select P3.PNum from Param P3 where P3.PTyp= 'PRDSTG' " + _
           "                            and P3.PmCd= L2.LocPrdStg), 0) " + _
           "                      ), 0)          ), '') " + _
           "    from Loc L1 JOIN Loc L2 ON L2.LocCoCd= L1.LocCoCd and L2.LocTyp = 'R' and L2.LocCd=L1.LocPrcs" + _
           " where L1.LocTyp= 'P' and L1.LocCoCd= BCoCd and L1.LocCd= BLoc and BLstKey = 0 " + _
           "    " + _
           "    update Bag set BPrdStg= (Select PMCd from Param where PTyp= 'PRDSTG' and PNum= 5) where BLoc= 'PFG' " + _
           "    " + _
           "    update Bag set " + _
           "      BPrdStg= IsNull((Select PMCd from Param P1 where PTyp= 'PRDSTG' " + _
           "                        and PNum= IsNull((Select max(P2.PNum) from Txnd TdLoc  " + _
           "                                  join Loc L1 on L1.LocCoCd= TdLoc.TdCoCd and L1.LocTyp= 'P' and L1.LocCd= TdLoc.TdByLoc " + _
           "                                  join Loc L2 on L2.LocCoCd= TdLoc.TdCoCd and L2.LocTyp= 'R' and L2.LocCd= L1.LocPrcs " + _
           "                                  join Param P2 on P2.PTyp= 'PRDSTG' and P2.PMCd= L2.LocPrdStg " + _
           "                                  Where TdLoc.TdCoCd = BCoCd And TdLoc.TdBYy = BYy " + _
           "                                  and TdLoc.TdBChr= BChr and TdLoc.TdBNo= BNo " + _
           "                                  and TdLoc.TdSrNo= 0 and TdLoc.TdByLoc<> '' and TdLoc.TdPrdYN= 'Y' " + _
           "                                  ), 0) " + _
           "                ), '') " + _
           "    where BLoc<> 'PFG' and BLstKey<> 0 " + _
           "End "

  '*** (Jen 2.14 Next)

  '*** (Bef 2.14 Next)
  'ws_str = " if IsNull((Select hPpcYN from Head where HCoCd= 'ZZZ' and HCd= 'ZSELF'), '') = 'Y' " + _
  '         "Begin " + _
  '         "    update Bag set BPrdStg= IsNull((Select PMCd from Param P1 where PTyp= 'PRDSTG' " + _
  '         "           and PNum= IsNull((Select max(P2.PNum) from Param P2 where P2.PTyp= 'PRDSTG' " + _
  '         "                            and P2.PNum< IsNull((Select P3.PNum from Param P3 where P3.PTyp= 'PRDSTG' " + _
  '         "                            and P3.PmCd= LocPrdStg), 0) " + _
  '         "                      ), 0) " + _
  '         "           ), '') " + _
  '         "    from Loc where LocTyp= 'P' and LocCoCd= BCoCd and LocCd= BLoc and BLstKey = 0 " + _
  '         "    " + _
  '         "    update Bag set BPrdStg= 'EI' where BLoc= 'PFG' " + _
  '         "    " + _
  '         "    update Bag set " + _
  '         "      BPrdStg= IsNull((Select PMCd from Param P1 where PTyp= 'PRDSTG' " + _
  '         "                        and PNum= IsNull((Select max(P2.PNum) from Txnd TdLoc  " + _
  '         "                                  join Loc on LocCoCd= TdLoc.TdCoCd and LocTyp= 'P' and LocCd= TdLoc.TdByLoc " + _
  '         "                                  join Param P2 on P2.PTyp= 'PRDSTG' and P2.PMCd= LocPrdStg " + _
  '         "                                  Where TdLoc.TdCoCd = BCoCd And TdLoc.TdBYy = BYy " + _
  '         "                                  and TdLoc.TdBChr= BChr and TdLoc.TdBNo= BNo " + _
  '         "                                  and TdLoc.TdSrNo= 0 and TdLoc.TdByLoc<> '' and TdLoc.TdPrdYN= 'Y' " + _
  '         "                                  ), 0) " + _
  '         "                ), '') " + _
  '         "    where BLoc<> 'PFG' and BLstKey<> 0 " + _
  '         "End "
  '*** (Bef 2.14 Next)
  moCn.Execute (ws_str)

  moCn.CommitTrans mdoCommitRec
  MWLib.EndProcess Me

  TxtDisp.text = "Bags Have Been Successfully Updated For Production Stage"
  DispMsg "Bags Have Been Successfully Updated For Production Stage", etInfo
  Call CmdUpdPrdStg_Click

  Exit Sub
  
ErrH:
  Call ErrHndlr
  MWLib.EndProcess Me
  Call CmdUpdPrdStg_Click
  '*** 2.11
End Sub
Private Sub LstDsgAnaCds()
  ' ****** Sachin 2.12 Patch 02 - Default Analysis Code Update for Unidesign - 10/04/06 ******
  Dim i As Integer
  Dim wRsDsgAna As MDORowSet
  i = 1
  Set wRsDsgAna = moCn.OpenRes("Select PMCd, PDesc From Param Where PTyp='DAANAFLD' And PValue='Y' And PValue1='Y'")
  GrdDsgAnlyFlds.Rows = 1
  With GrdDsgAnlyFlds
    GrdDsgAnlyFlds.StartCopy
    Do While Not (wRsDsgAna.EOF Or wRsDsgAna.BOF)
      .AddItem
      i = .Rows - 1
      .Value(i, "WDANASR") = wRsDsgAna!PMCd
      .Value(i, "WDAANAFLDDESC") = wRsDsgAna!PDesc
      .Value(i, "WDANADEFVAL") = ""
      wRsDsgAna.MoveNext
    Loop
    GrdDsgAnlyFlds.EndCopy
  End With
  ' ****** Sachin 2.12 Patch 02 - Default Analysis Code Update for Unidesign - 10/04/06 ******
End Sub
Private Sub UpdDefDsgAnaCodes()
  ' ****** Sachin 2.12 Patch 02 - Default Analysis Code Update for Unidesign - 10/04/06 ******
  Dim ws_AnaFlds As Integer
  ws_AnaFlds = 1
  Me.MousePointer = vbHourglass
  Do While ws_AnaFlds <= GrdDsgAnlyFlds.Rows - 1
    '6.1
    moCn.Execute "Insert Into DsgAna (DaTcTyp,DaCd,DaSz,DaAnaSr,DaAnaCd,DaAnaDesc,ModUsr,ModDt,ModTime,DaDmIdNo,DaPrtKey) " + _
                 " Select DmTcTyp, DmCd , DmSz , '" + GrdDsgAnlyFlds.Value(ws_AnaFlds, "WDANASR") + "', '" + GrdDsgAnlyFlds.Value(ws_AnaFlds, "WDANADEFVAL") + "', '', 'MIC', '" + Format(moCn.SrvrDate, DtShortStr()) + "', 0.0,DmIdNo,'C' " + _
                 " From DsgMst " + _
                 " Where DmTcTyp='DM' And DmPrtKey='" + ctCurrPrtn + "' " + _
                 "          And Not Exists (Select 'x' From DsgAna Where DaTcTyp=DmTcTyp And DaCd=DmCd And DaSz=DmSz " + _
                 "          And DaAnaSr='" + GrdDsgAnlyFlds.Value(ws_AnaFlds, "WDANASR") + "')"
    ws_AnaFlds = ws_AnaFlds + 1
  Loop
  Me.MousePointer = vbDefault
  TxtDisp.text = "Design Analysis Codes Updated Successfully !!"
  ' ****** Sachin 2.12 Patch 02 - Default Analysis Code Update for Unidesign - 10/04/06 ******
End Sub

' ****** Sachin [31-08-06] Update Ini Date from Maintenance Exe ******
Private Function FileLoad() As Boolean
  FileLoad = False
  If TxtDllPath.text <> "" And TxtDllPath.text <> "" Then
    If LoadFile = True Then FileLoad = True
  End If
End Function

Private Function LoadFile() As Boolean
  
  On Error GoTo ErrHndl

  Dim ws_IniStr As String, wi_Pos As Integer
  Dim wo_FS As FileSystemObject, wo_File As file, wo_tsIni As TextStream
  Dim ws_Arr() As String, ws_IniArr() As String, ws_CoDet() As String, i As Integer
  Dim wDt As Date, wsDt As String, wDtAndCoChk As Variant, ws_Tmp As String, Ws_TmpArr() As String
  Dim wt_Ini As IniDet, ws_ArrCo() As CoDet, wCrypt As enGENCryptFormula, wiCnt As Integer
  
  LoadFile = False

  '***** Reading Ini File
  wt_Ini.ts_FlNm = TxtDllPath.text + IIF(CmbProj.text = "EMR", "\EMR.INI", "\STO.INI")
  wt_Ini.ts_PrjNm = Trim(CmbProj.text)
  
  Set wo_FS = New FileSystemObject
  Set wo_File = wo_FS.GetFile(wt_Ini.ts_FlNm)
  Set wo_tsIni = wo_File.OpenAsTextStream(ForReading)
  
  ws_IniStr = wo_tsIni.ReadAll
  ws_IniArr = Split(ws_IniStr, vbCrLf)
  wo_tsIni.Close

  Set wo_tsIni = Nothing
  Set wo_File = Nothing
  Set wo_FS = Nothing
  
  If UBound(ws_IniArr) < 1 Then GoTo ErrHndl     'No lines in the Ini
  If ws_IniArr(0) = "" Then GoTo ErrHndl         'No lines in the Ini

  '***** Validating values in first line of .Ini File
  wi_Pos = InStr(1, ws_IniArr(0), "MwMstSrv=")
  If wi_Pos = 0 Then GoTo ErrHndl
  
  ws_IniArr(0) = Mid(ws_IniArr(0), wi_Pos + Len("MwMstSrv="))
  ws_Arr = Split(ws_IniArr(0), ",")
  
  If Not IsArray(ws_Arr) Then GoTo ErrHndl
  
  wCrypt = GenCryptNew     'Valid only for New Encryption
   
  If UBound(ws_Arr) <> 4 Then GoTo ErrHndl
  If ws_IniArr(0) = "" Then GoTo ErrHndl

  '***** Validating values in first line of .Ini File
  wt_Ini.ts_SrvrNm = Trim(ws_Arr(0))
  wt_Ini.ts_SrlNo = UCase(Decrypt(Trim$(ws_Arr(3)), wCrypt))
      
      If IsNumeric(ws_Arr(4)) Then
        wt_Ini.tl_TmOut = CLng(ws_Arr(4))
      Else
        wt_Ini.tl_TmOut = -1
      End If
  
  wiCnt = UBound(ws_IniArr) - 1
  
  ReDim ws_ArrCo(wiCnt)
  For i = 1 To UBound(ws_IniArr)
    If Trim$(ws_IniArr(i)) = "" Then
      wiCnt = wiCnt - 1
      ReDim Preserve ws_ArrCo(wiCnt)
      GoTo NextI
    End If
    
    ws_CoDet = Split(ws_IniArr(i), ",")
    If Not IsArray(ws_CoDet) Then GoTo ErrHndl
    If UBound(ws_CoDet) < 1 Then GoTo ErrHndl

    ws_ArrCo(i - 1).ts_CoCd = UCase(Decrypt(Trim(ws_CoDet(0)), wCrypt))
    ws_Tmp = Decrypt(Trim(ws_CoDet(1)), wCrypt)
    wDtAndCoChk = Left$(ws_Tmp, Len(ws_Tmp) - 3)
    ws_Tmp = Right$(ws_Tmp, 3)
    
    If IsNumeric(ws_Tmp) Then
      If CInt(ws_Tmp) >= wt_Ini.ti_NoOfUsr Then wt_Ini.ti_NoOfUsr = CInt(ws_Tmp)
    Else
      wt_Ini.ti_NoOfUsr = -1
    End If

    If Len(wDtAndCoChk) >= 8 Then
      wsDt = IIF(InStr(1, wDtAndCoChk, "/") = 2, "0" & Mid(wDtAndCoChk, 1, 7), Mid(wDtAndCoChk, 1, 8))
      wDt = CDate(wsDt)
      ws_ArrCo(i - 1).ts_CoNm = IIF(InStr(1, wDtAndCoChk, "/") = 2, Mid(wDtAndCoChk, 8), IIF(Len(wDtAndCoChk) > 8, Mid(wDtAndCoChk, 9), ""))
      If wDt > wt_Ini.td_EndDt Then wt_Ini.td_EndDt = wDt
    End If
    
     'For Emr any Random Date will be generated between 15/03/07 and 15/07/07
     'For Stores it will be 01/10/2007
    Randomize
    wt_Ini.td_EndDt = IIF(CmbProj.text = "EMR", CStr(CDate(NewEmrLckDt) + Int((120 * Rnd) + 1)), NewStoLckDt)
    
NextI:
  Next i
  wt_Ini.tCoDet = ws_ArrCo

  If ValidtIni(wt_Ini, False) = True Then
    mt = GetIniDet(wt_Ini)
  Else
    Exit Function
  End If

  LoadFile = True
  Exit Function
ErrHndl:
  TxtDisp.text = "Dll Load Error Occurred"
End Function

Private Function SaveFile() As Boolean

  On Error GoTo ErrHndl
  Dim wId As IniDet
  Dim ws_IniStr As String, wi_Pos As Integer
  Dim wo_FS As FileSystemObject, wo_File As file, wo_tsIni As TextStream
  Dim ws_Arr() As String, ws_IniArr() As String, ws_CoDet() As String, i As Integer
  Dim wCrypt As enGENCryptFormula, ws_Tmp As String, ws_Tmp1 As String
  
  If FileLoad = False Then Exit Function
    SaveFile = False

  Set wo_FS = New FileSystemObject
  Set wo_tsIni = wo_FS.CreateTextFile(mt.ts_FlNm, True)
  
    wCrypt = GenCryptNew
    
    ws_Tmp = "MwMstSrv=" & mt.ts_SrvrNm & ","
    ws_Tmp = ws_Tmp & Encrypt("UsrAppPwd", wCrypt) & "," & Encrypt("MwAppRolePwd", wCrypt) & ","
    ws_Tmp = ws_Tmp & Encrypt(Trim$(mt.ts_SrlNo), wCrypt) & ","
    ws_Tmp = ws_Tmp & Trim$(CStr(mt.tl_TmOut))
    wo_tsIni.WriteLine ws_Tmp
  
  For i = LBound(mt.tCoDet) To UBound(mt.tCoDet)
    ws_Tmp = Encrypt(mt.tCoDet(i).ts_CoCd, wCrypt) & ","
    ws_Tmp1 = Format$(mt.td_EndDt, DtShortStr()) & Trim$(mt.tCoDet(i).ts_CoNm)    '6.1
    ws_Tmp1 = ws_Tmp1 & Format$(mt.ti_NoOfUsr, "000")
    ws_Tmp = ws_Tmp & Encrypt(ws_Tmp1, wCrypt)
    wo_tsIni.Write ws_Tmp + IIF(i <> UBound(mt.tCoDet), vbCrLf, "")
  Next
        
    SaveFile = True
  
ErrH:
      wo_tsIni.Close
  Set wo_tsIni = Nothing
  Set wo_FS = Nothing
  Exit Function
  
ErrHndl:
    TxtDisp.text = "Dll Error Occurred "
  GoTo ErrH
  
End Function

Private Function ValidtIni(pid As IniDet, Optional pbDtChk As Boolean = True) As Boolean
Dim ErrMsg As String, i As Integer
  ValidtIni = False
  If pid.ts_FlNm = "" Then ErrMsg = "Dll Folder Not Specified.": GoTo ErrHndl
  If pid.ts_PrjNm = "" Then ErrMsg = "Project Name Not Specified.": GoTo ErrHndl
  If pid.ts_SrvrNm = "" Then ErrMsg = "DataBase Server Name Not Specified.": GoTo ErrHndl
  If pid.ti_Wait > 60 Then ErrMsg = "Login Wait Time Can Not Be More Than 60 Seconds": GoTo ErrHndl
  If pid.ti_Retry > 10 Then ErrMsg = "Login Retry Count Can Not Be More Than 10 Times": GoTo ErrHndl
  If pid.ts_SrlNo = "" Then ErrMsg = "Hard Disk Serial No. Not Specified.": GoTo ErrHndl
  If pid.tl_TmOut = -1 Then ErrMsg = "Time Out Value is Not Proper.": GoTo ErrHndl
  If pid.ti_NoOfUsr > 999 Or pid.ti_NoOfUsr < 0 Then ErrMsg = "No. of Users Must be In Range (0 To 999).": GoTo ErrHndl
  If pid.td_EndDt <= CDate(Now) And pbDtChk = True Then ErrMsg = "Date is Not Proper.": GoTo ErrHndl
  For i = LBound(pid.tCoDet) To UBound(pid.tCoDet)
    If pid.tCoDet(i).ts_CoCd = "" Then ErrMsg = "Company Code Cannot be Blank.": GoTo ErrHndl
    If Len(pid.tCoDet(i).ts_CoCd) > 3 Then ErrMsg = "Company Code Not Proper.": GoTo ErrHndl
    If pid.tCoDet(i).ts_CoNm = "" Then ErrMsg = "Company Name Cannot be Blank.": GoTo ErrHndl
    If Len(pid.tCoDet(i).ts_CoNm) > 30 Then ErrMsg = "Company Name Not Proper.": GoTo ErrHndl
  Next
  ValidtIni = True
  Exit Function
ErrHndl:
  TxtDisp.text = "Dlls Validation Error Occurred"
End Function

Private Function GetIniDet(pid As IniDet) As IniDet
  GetIniDet.tCoDet = pid.tCoDet
  GetIniDet.ti_NoOfUsr = pid.ti_NoOfUsr
  GetIniDet.tl_TmOut = pid.tl_TmOut
  GetIniDet.ts_FlNm = pid.ts_FlNm
  GetIniDet.ts_PrjNm = pid.ts_PrjNm
  GetIniDet.ti_Wait = pid.ti_Wait
  GetIniDet.ti_Retry = pid.ti_Retry
  GetIniDet.ts_SrlNo = pid.ts_SrlNo
  GetIniDet.ts_SrvrNm = pid.ts_SrvrNm
  GetIniDet.td_EndDt = pid.td_EndDt
End Function

Public Function BrowseForFolder(hwndOwner As Long, sPrompt As String) As String
    'declare variables to be used
     Dim iNull As Integer
     Dim lpIDList As Long
     Dim lResult As Long
     Dim sPath As String
     Dim udtBI As BrowseInfo

    'initialise variables
     With udtBI
        .hwndOwner = hwndOwner
        .lpszTitle = lstrcat(sPrompt, "")
        .ulFlags = BIF_RETURNONLYFSDIRS
     End With

    'Call the browse for folder API
     lpIDList = SHBrowseForFolder(udtBI)
      
    'get the resulting string path
     If lpIDList Then
        sPath = String$(MAX_PATH, 0)
        lResult = SHGetPathFromIDList(lpIDList, sPath)
        Call CoTaskMemFree(lpIDList)
        iNull = InStr(sPath, vbNullChar)
        If iNull Then sPath = Left$(sPath, iNull - 1)
     End If

     BrowseForFolder = sPath
End Function
' ****** Sachin [31-08-06] Update Ini Date from Maintenance Exe ******

Private Sub UpdateCoCd()
  
  On Error GoTo ErrH

  TxtDisp.text = ""
  DispMsg "", etInfo

  'Dim ws_NetId As String
  'ws_NetId = NetDets.IdNo

  'If moCn.RecSeek("select count('x') From master.dbo.sysprocesses " + _
  '                "where sid= suser_sid('UsrApp') and db_name(dbid)= '" + MWLib.DbNm + "' and " + _
  '                "net_address<> '" + ws_NetId + "'") Then
  '  TxtDisp.Text = "All Users Should First Log Out Of The System For This Option To Work"
  '  DispMsg "All Users Should First Log Out Of The System For This Option To Work", etError
  '  Exit Sub
  'End If

  '*** Check if Company Code update is possible
  If adc("wUcOldCoCd") = "" Then
    TxtDisp.text = "Old Company Code Should Be Entered"
    DispMsg "Old Company Code Should Be Entered", etError
    Exit Sub
  End If
  If adc("wUcOldCoCd") = "" Or adc("wUcNewCoCd") = "" Or adc("wUcNewCoName") = "" Then
    TxtDisp.text = "New Company Code/ Name Should Be Entered"
    DispMsg "New Company Code/ Name Should Be Entered", etError
    Exit Sub
  End If

  '***
'  --select @@SERVERNAME
'declare @sid varbinary(85)
'select @sid = suser_sid('UsrApp')
'if (@sid is null)
'begin
'  raiserror(15007,-1,-1,'UsrApp')
'End
'select spid, ecid, status,
'     loginame= rtrim(loginame),
'     hostname, blk= convert(char(5),blocked),
'     dbname = (case when dbid = 0 then null
'      when dbid <> 0 then db_name(dbid)
'      end), cmd
'From master.dbo.sysprocesses
'where sid = @sid and hostname<> 'MW6'
  '***

  DoEvents
  MWLib.BeginProcess Me, "Starting With Company Code Update"
  moCn.BeginTrans (mdoCommitRec)

  If adc("wUcOldCoCd") = adc("wUcNewCoCd") Then
    If adc("wUcNewCoName") <> adc("wUcOldCoName") Then
      ws_str = "Insert Into TrgDisa Values ('Head') " + _
               "update Head set HName= '" + adc("wUcNewCoName") + "' " + _
               "where HCoCd= '" + adc("wUcOldCoCd") + "' " + _
               "Delete From TrgDisa "
      moCn.Execute (ws_str)
    End If
  Else
    ws_str = "Insert Into TrgDisa Values ('Head') " + _
             "Insert Into TrgDisa Values ('CustMst') " + _
             "Insert Into TrgDisa Values ('Bag') " + _
             "Insert Into TrgDisa Values ('BgRecoMst') " + _
             "Insert Into TrgDisa Values ('BgRecoDet') " + _
             "Insert Into TrgDisa Values ('Fg') " + _
             "Insert Into TrgDisa Values ('Fgd') " + _
             "Insert Into TrgDisa Values ('FgRm')" + _
             "Insert Into TrgDisa Values ('InvDsg') " + _
             "Insert Into TrgDisa Values ('InvFgd') " + _
             "Insert Into TrgDisa Values ('InvHd') " + _
             "Insert Into TrgDisa Values ('InvLab') " + _
             "Insert Into TrgDisa Values ('InvRm') " + _
             "Insert Into TrgDisa Values ('InvZRm') " + _
             "Insert Into TrgDisa Values ('Loc') " + _
             "Insert Into TrgDisa Values ('LsRec') " + _
             "Insert Into TrgDisa Values ('OrdAmd') "
    ws_str = ws_str + "Insert Into TrgDisa Values ('OrdDsg') " + _
             "Insert Into TrgDisa Values ('OrdLab') " + _
             "Insert Into TrgDisa Values ('OrdMst') " + _
             "Insert Into TrgDisa Values ('OrdPln') " + _
             "Insert Into TrgDisa Values ('OrdPlnDt') " + _
             "Insert Into TrgDisa Values ('OrdRm') " + _
             "Insert Into TrgDisa Values ('OrdTgt') " + _
             "Insert Into TrgDisa Values ('PpcCap') " + _
             "Insert Into TrgDisa Values ('Txn') " + _
             "Insert Into TrgDisa Values ('Txnd') " + _
             "Insert Into TrgDisa Values ('TxndPrd') " + _
             "Insert Into TrgDisa Values ('Txndz') " + _
             "Insert Into TrgDisa Values ('vParam') "

    '********Geeta********Emr207***************
    ws_str = ws_str + "Insert Into TrgDisa Values ('TreeMst') " + _
             "Insert Into TrgDisa Values ('TreeDet') " + _
             "Insert Into TrgDisa Values ('Tag') " + _
             "Insert Into TrgDisa Values ('TagData') " + _
             "Insert Into TrgDisa Values ('TagFmt') " + _
             "Insert Into TrgDisa Values ('TagFmtSr') " + _
             "Insert Into TrgDisa Values ('UsrLogin') " + _
             "Insert Into TrgDisa Values ('LossMst') "
    '********Geeta********Emr207***************


    '*** (Jen 2.13)
    ' ***** Manali 3.8.0 - TcMst Added
    ws_str = ws_str + " Insert Into TrgDisa Values ('RetMem') " + _
            "Insert Into TrgDisa Values ('RetMemDet') " + _
             "Insert Into TrgDisa Values ('TcMst') " + _
             "Insert Into TrgDisa Values ('DmYyMm') " + _
             "Insert Into TrgDisa Values ('RfIdMst') " + _
             "Insert Into TrgDisa Values ('MultiPrcQtMst') " + _
             "Insert Into TrgDisa Values ('MultiPrcQtDsg') "

    '*** (Jen 2.13)

    moCn.Execute (ws_str)

    ws_str = "update Head set HCoCd= '" + adc("wUcNewCoCd") + "', HName= '" + adc("wUcNewCoName") + "' " + _
             "where HCoCd= '" + adc("wUcOldCoCd") + "' " + _
             "update CustMst set CmCd= '" + adc("wUcNewCoCd") + "', CmName= '" + adc("wUcNewCoName") + "' " + _
             "where CmCtg='Z' and CmCd= '" + adc("wUcOldCoCd") + "' " + _
             "update Bag set BCoCd= '" + adc("wUcNewCoCd") + "' where BCoCd= '" + adc("wUcOldCoCd") + "' " + _
             "update BgRecoMst set BmCoCd= '" + adc("wUcNewCoCd") + "' where BmCoCd= '" + adc("wUcOldCoCd") + "' " + _
             "update BgRecoDet set BdCoCd= '" + adc("wUcNewCoCd") + "' where BdCoCd= '" + adc("wUcOldCoCd") + "' " + _
             "update Fg set FgCoCd= '" + adc("wUcNewCoCd") + "' where FgCoCd= '" + adc("wUcOldCoCd") + "' " + _
             "update Fgd set FdCoCd= '" + adc("wUcNewCoCd") + "' where FdCoCd= '" + adc("wUcOldCoCd") + "' " + _
             "update FgRm set FrCoCd= '" + adc("wUcNewCoCd") + "' where FrCoCd= '" + adc("wUcOldCoCd") + "' " + _
             "update InvHd set InCoCd= '" + adc("wUcNewCoCd") + "' where InCoCd= '" + adc("wUcOldCoCd") + "' " + _
             "update InvDsg set IdCoCd= '" + adc("wUcNewCoCd") + "' where IdCoCd= '" + adc("wUcOldCoCd") + "' " + _
             "update InvFgd set IfCoCd= '" + adc("wUcNewCoCd") + "' where IfCoCd= '" + adc("wUcOldCoCd") + "' " + _
             "update InvLab set iLCoCd= '" + adc("wUcNewCoCd") + "' where iLCoCd= '" + adc("wUcOldCoCd") + "' " + _
             "update InvRm set IrCoCd= '" + adc("wUcNewCoCd") + "' where IrCoCd= '" + adc("wUcOldCoCd") + "' " + _
             "update InvZRm set IzCoCd= '" + adc("wUcNewCoCd") + "' where IzCoCd= '" + adc("wUcOldCoCd") + "' " + _
             "update Loc set LocCoCd= '" + adc("wUcNewCoCd") + "' where LocCoCd= '" + adc("wUcOldCoCd") + "' " + _
             "update LsRec set LrcCoCd= '" + adc("wUcNewCoCd") + "' where LrcCoCd= '" + adc("wUcOldCoCd") + "' " + _
             "update OrdAmd set OaCoCd= '" + adc("wUcNewCoCd") + "' where OaCoCd= '" + adc("wUcOldCoCd") + "' " + _
             "update OrdDsg set OdCoCd= '" + adc("wUcNewCoCd") + "' where OdCoCd= '" + adc("wUcOldCoCd") + "' " + _
             "update OrdLab set OlCoCd= '" + adc("wUcNewCoCd") + "' where OlCoCd= '" + adc("wUcOldCoCd") + "' "
    ws_str = ws_str + "update OrdMst set OmCoCd= '" + adc("wUcNewCoCd") + "' where OmCoCd= '" + adc("wUcOldCoCd") + "' " + _
             "update OrdPln set OpCoCd= '" + adc("wUcNewCoCd") + "' where OpCoCd= '" + adc("wUcOldCoCd") + "' " + _
             "update OrdPlnDt set OpdCoCd= '" + adc("wUcNewCoCd") + "' where OpdCoCd= '" + adc("wUcOldCoCd") + "' " + _
             "update OrdRm set OrCoCd= '" + adc("wUcNewCoCd") + "' where OrCoCd= '" + adc("wUcOldCoCd") + "' " + _
             "update OrdTgt set OtCoCd= '" + adc("wUcNewCoCd") + "' where OtCoCd= '" + adc("wUcOldCoCd") + "' " + _
             "update PpcCap set PcCoCd= '" + adc("wUcNewCoCd") + "' where PcCoCd= '" + adc("wUcOldCoCd") + "' " + _
             "update SYyMm set SCoCd= '" + adc("wUcNewCoCd") + "' where SCoCd= '" + adc("wUcOldCoCd") + "' " + _
             "update Txn set TCoCd= '" + adc("wUcNewCoCd") + "' where TCoCd= '" + adc("wUcOldCoCd") + "' " + _
             "update Txnd set TdCoCd= '" + adc("wUcNewCoCd") + "' where TdCoCd= '" + adc("wUcOldCoCd") + "' " + _
             "update TxndPrd set TpCoCd= '" + adc("wUcNewCoCd") + "' where TpCoCd= '" + adc("wUcOldCoCd") + "' " + _
             "update TxndZ set TzCoCd= '" + adc("wUcNewCoCd") + "' where TzCoCd= '" + adc("wUcOldCoCd") + "' " + _
             "update vParam set vPCoCd= '" + adc("wUcNewCoCd") + "' where vPCoCd= '" + adc("wUcOldCoCd") + "' " + _
             "update AuToGen set APREKEY= '" + adc("wUcNewCoCd") + "' + SubString(APREKEY, 4, len(APREKEY)- 3) " + _
             "where CHARINDEX ( '" + adc("wUcOldCoCd") + "' , APREKEY, 1)> 0  "

    '********Geeta********Emr207***************
    ws_str = ws_str + "update TreeMst set TrCoCd= '" + adc("wUcNewCoCd") + "' where TrCoCd= '" + adc("wUcOldCoCd") + "' " + _
             "update TreeDet set TrdCoCd= '" + adc("wUcNewCoCd") + "' where TrdCoCd= '" + adc("wUcOldCoCd") + "' " + _
             "update Tag set TgCoCd= '" + adc("wUcNewCoCd") + "' where TgCoCd= '" + adc("wUcOldCoCd") + "' " + _
             "update TagData set TgdCoCd= '" + adc("wUcNewCoCd") + "' where TgdCoCd= '" + adc("wUcOldCoCd") + "' " + _
             "update TagFmt set TgfCoCd= '" + adc("wUcNewCoCd") + "' where TgfCoCd= '" + adc("wUcOldCoCd") + "' " + _
             "update TagFmtSr set TgfsCoCd= '" + adc("wUcNewCoCd") + "' where TgfsCoCd= '" + adc("wUcOldCoCd") + "' " + _
             "update UsrLogin set UlCoCd= '" + adc("wUcNewCoCd") + "' where UlCoCd= '" + adc("wUcOldCoCd") + "' " + _
             "update LossMst set LmCoCd= '" + adc("wUcNewCoCd") + "' where LmCoCd= '" + adc("wUcOldCoCd") + "' "
    '********Geeta********Emr207***************
   
    '*** (Jen 2.13)
    ws_str = ws_str + "update RetMem set ReCoCd= '" + adc("wUcNewCoCd") + "' where ReCoCd= '" + adc("wUcOldCoCd") + "' " + _
             "update RetMemDet set RedCoCd= '" + adc("wUcNewCoCd") + "' where RedCoCd= '" + adc("wUcOldCoCd") + "' " + _
             "update TcMst set TmCoCd= '" + adc("wUcNewCoCd") + "' where TmCoCd= '" + adc("wUcOldCoCd") + "' " + _
             "update DmYyMm set DCoCd= '" + adc("wUcNewCoCd") + "' where DCoCd= '" + adc("wUcOldCoCd") + "' " + _
             "update RfIdMst set RimBCoCd= '" + adc("wUcNewCoCd") + "' where RimBCoCd= '" + adc("wUcOldCoCd") + "' " + _
             "update MultiPrcQtMst set MqmCoCd= '" + adc("wUcNewCoCd") + "' where MqmCoCd= '" + adc("wUcOldCoCd") + "' " + _
             "update MultiPrcQtDsg set MqdCoCd= '" + adc("wUcNewCoCd") + "' where MqdCoCd= '" + adc("wUcOldCoCd") + "' "
    '*** (Jen 2.13)
   
    moCn.Execute (ws_str)

    ws_str = "Delete From TrgDisa "

'    ws_str = "alter table Head enable Trigger Head_InUpDl " + _
'             "alter table CustMst enable trigger CustMst_InUpDl " + _
'             "alter table Bag enable trigger Bag_InUpDl " + _
'             "alter table Fg enable trigger Fg_InUpDl " + _
'             "alter table Fgd enable trigger Fgd_InUpDl " + _
'             "alter table FgRm enable trigger FgRm_InUpDl " + _
'             "alter table InvDsg enable trigger InvDsg_InUpDl " + _
'             "alter table InvFgd enable trigger InvFgd_InUpDl " + _
'             "alter table InvHd enable trigger InvHd_InUpDl " + _
'             "alter table InvLab enable trigger InvLab_InUpDl " + _
'             "alter table InvRm enable trigger InvRm_InUpDl " + _
'             "alter table InvZRm enable trigger InvZRm_InUpDl " + _
'             "alter table Loc enable trigger Loc_InUpDl " + _
'             "alter table LsRec enable trigger LsRec_InUpDl " + _
'             "alter table OrdAmd enable trigger OrdAmd_InUpDl " + _
'             "alter table OrdDsg enable trigger OrdDsg_InUpDl " + _
'             "alter table OrdLab enable trigger OrdLab_InUpDl "
'    ws_str = ws_str + "alter table OrdMst enable trigger OrdMst_InUpDl " + _
'             "alter table OrdPln enable trigger OrdPln_InUpDl " + _
'             "alter table OrdPlnDt enable trigger OrdPlnDt_InUpDl " + _
'             "alter table OrdRm enable trigger OrdRm_InUpDl " + _
'             "alter table OrdTgt enable trigger OrdTgt_InUpDl " + _
'             "alter table PpcCap enable trigger PpcCap_InUpDl " + _
'             "alter table Txn enable trigger Txn_InUpDl " + _
'             "alter table Txnd enable trigger Txnd_InUpDl " + _
'             "alter table TxndPrd enable trigger TxndPrd_InUpDl " + _
'             "alter table TxndZ enable trigger TxndZ_InUpDl " + _
'             "alter table vParam enable trigger vParam_InUpDl "
'
'   '********Geeta********Emr207***************
'   ws_str = ws_str + "alter table TreeMst Enable trigger TreeMst_InUpDl " + _
'                        "alter table TreeDet Enable trigger TreeDet_InUpDl " + _
'                        "alter table Tag Enable trigger Tag_InUpDl " + _
'                        "alter table TagData Enable trigger TagData_InUpDl " + _
'                        "alter table TagFmt Enable trigger TagFmt_InUpDl " + _
'                        "alter table TagFmtSr Enable trigger TagFmtSr_InUpDl " + _
'                        "alter table UsrLogin Enable trigger UsrLogin_InUpDl " + _
'                        "alter table LossMst Enable trigger LossMst_InUpDl "

    '********Geeta********Emr207***************

    moCn.Execute (ws_str)
    ws_str = "Insert Into TrgDisa Values ('Fm') " + _
             "Insert Into TrgDisa Values ('Fmd') " + _
             "Insert Into TrgDisa Values ('OrdScop') "
    moCn.Execute (ws_str)
    ws_str = "update Fm set FmCoCd= '" + adc("wUcNewCoCd") + "' where FmCoCd= '" + adc("wUcOldCoCd") + "' " + _
             "update Fmd set FmdCoCd= '" + adc("wUcNewCoCd") + "' where FmdCoCd= '" + adc("wUcOldCoCd") + "' " + _
             "update OrdScop set OsCoCd= '" + adc("wUcNewCoCd") + "' where OsCoCd= '" + adc("wUcOldCoCd") + "'"
    moCn.Execute (ws_str)
    ws_str = "Delete From TrgDisa "
    'ws_str = "alter table Fm enable trigger Fm_InUpDl " + _
    '         "alter table Fmd enable trigger Fmd_InUpDl " + _
    '         "alter table OrdScop enable trigger OrdScop_InUpDl"
    moCn.Execute (ws_str)

''    moCn.CommitTrans mdoCommitRec
''    MWLib.EndProcess Me
''
''    TxtDisp.text = "Company Code Has Been Successfully Updated"
''    DispMsg "Company Code Has Been Successfully Updated", etInfo
''    Call CmdUpdCoCd_Click
''
''    Exit Sub
    
  End If
  
  Call UpdIntegrity(moCn)
  moCn.CommitTrans mdoCommitRec
  MWLib.EndProcess Me
  
  TxtDisp.text = "Company Code/Name Has Been Successfully Updated"
  DispMsg "Company Code/Name Has Been Successfully Updated", etInfo
  Call CmdUpdCoCd_Click

  Exit Sub
  
ErrH:
  Call ErrHndlr
  moCn.CommitTrans mdoCommitNone
  MWLib.EndProcess Me
  ws_str = "Delete From TrgDisa"
  moCn.Execute (ws_str)
  Call CmdUpdCoCd_Click
End Sub

Private Sub DeleteOrder()
  '****** The Before and After Trigger Procedures should be on the path D:\MWAPP\MWSQL ******
  '*** This procedure was used for version 2.05

  Dim ws_str As String, ws_FgCnd As String, ws_BagCnd As String, ws_InvCnd As String
  Dim ws_OrdCnd As String, ws_pth As String
  On Error GoTo ErrH

  TxtDisp.text = ""
  DispMsg "", etInfo

  '*** Check if Order For The Specified Scope Can Be Deleted
  '*** Very Important
  If Not moCn.RecSeek("Select HCoCd from Head where HCoCd= '" + gs_CoCd + "' and HCd= '" + ctSelfCmCd + "'") Then
    TxtDisp.text = "The Company Code '" + gs_CoCd + "' No Longer Exists In The Database. Login Again Using Another Company Code"
    DispMsg "The Company Code '" + gs_CoCd + "' No Longer Exists In The Database. Login Again Using Another Company Code", etError
    Call CmdDelOrd_Click: Exit Sub
  End If
  '*** Very Important
  
  If adc("wDoOmTcFr") = "" And adc("wDoOmYyFr") = "" And adc("wDoOmChrFr") = "" _
     And adc("wDoOmNoFr") = 0 And adc("wDoOmSrFr") = 0 And adc("wDoOmTcTo") = "" _
     And adc("wDoOmYyTo") = "" And adc("wDoOmChrTo") = "" And adc("wDoOmNoTo") = 0 _
     And adc("wDoOmSrTo") = 0 Then
    TxtDisp.text = "Some Scope Should Be Entered"
    DispMsg "Some Scope Should Be Entered", etError
    Exit Sub
  End If

  '  If Not moCn.RecSeek("select 'x' from dbo.sysobjects where " + _
  '         "id = object_id(N'[dbo].[MwSp_BefTrgs]') and " + _
  '         "OBJECTPROPERTY(id, N'IsProcedure') = 1") Then
  '    TxtDisp.text = "Procedure 'MwSp_BefTrgs' Does Not Exist In The Database "
  '    DispMsg "Procedure 'MwSp_BefTrgs' Does Not Exist In The Database ", etError
  '    Call CmdDelOrd_Click: Exit Sub
  '  End If
  '
  '  If Not moCn.RecSeek("select 'x' from dbo.sysobjects where " + _
  '         "id = object_id(N'[dbo].[MwSp_CorrectTrgs]') and " + _
  '         "OBJECTPROPERTY(id, N'IsProcedure') = 1") Then
  '    TxtDisp.text = "Procedure 'MwSp_CorrectTrgs' Does Not Exist In The Database "
  '    DispMsg "Procedure 'MwSp_CorrectTrgs' Does Not Exist In The Database ", etError
  '    Call CmdDelOrd_Click: Exit Sub
  '  End If

  '****** Sachin 2.12 - Delete Order  - 16-02-2006
  ws_str = "if exists (select * from sysobjects where id = object_id('dbo.MwSp_BefTrgs') and sysstat & 0xf = 4)" + _
           "  drop procedure [dbo].[MwSp_BefTrgs]"
  moCn.Execute (ws_str)
  ws_str = " if exists (select * from sysobjects where id = object_id('dbo.MwSp_CorrectTrgs') and sysstat & 0xf = 4) " + _
           " drop procedure [dbo].[MwSp_CorrectTrgs]"
  moCn.Execute (ws_str)
  
  '****** Checking whether the Required Files Exists
  ws_pth = "D:\MWAPP\MWSQL\MwSp_BefTrgs.Sql"
  If Dir(ws_pth, vbNormal) = "" Then
      TxtDisp.text = "File Not Found : '" + ws_pth + "', Cannot Continue...'"
      DispMsg "File Not Found : '" + ws_pth + "'", etError
      Exit Sub
  End If
  ws_pth = "D:\MWAPP\MWSQL\MwSp_CorrectTrgs.Sql"
  If Dir(ws_pth, vbNormal) = "" Then
      TxtDisp.text = "File Not Found : '" + ws_pth + "', Cannot Continue...'"
      DispMsg "File Not Found : '" + ws_pth + "'", etError
      Exit Sub
  End If
  '****** If Found Execute the Files
  ws_str = ReadFileContents("D:\MWAPP\MWSQL\MwSp_BefTrgs.Sql")
  moCn.Execute (ws_str)
  ws_str = ReadFileContents("D:\MWAPP\MWSQL\MwSp_CorrectTrgs.Sql")
  moCn.Execute (ws_str)
  '****** Sachin 2.12 - Delete Order  - 16-02-2006

  adc("wDoOmTcFr").CmpStr = "OdTc>= ": adc("wDoOmTcTo").CmpStr = "OdTc<= "
  adc("wDoOmYyFr").CmpStr = "OdYy>= ": adc("wDoOmYyTo").CmpStr = "OdYy<= "
  adc("wDoOmChrFr").CmpStr = "OdChr>= ": adc("wDoOmChrTo").CmpStr = "OdChr<= "
  adc("wDoOmNoFr").CmpStr = "OdNo>= ": adc("wDoOmNoTo").CmpStr = "OdNo<= "
  adc("wDoOmSrFr").CmpStr = "OdSr>= ": adc("wDoOmSrTo").CmpStr = "OdSr<= "
  ws_OrdCnd = adc.RepCond
  ws_OrdCnd = IIF(ws_OrdCnd = "", "", " and " + ws_OrdCnd)

  adc("wDoOmTcFr").CmpStr = "FdPrdOdTc>= ": adc("wDoOmTcTo").CmpStr = "FdPrdOdTc<= "
  adc("wDoOmYyFr").CmpStr = "FdPrdOdYy>= ": adc("wDoOmYyTo").CmpStr = "FdPrdOdYy<= "
  adc("wDoOmChrFr").CmpStr = "FdPrdOdChr>= ": adc("wDoOmChrTo").CmpStr = "FdPrdOdChr<= "
  adc("wDoOmNoFr").CmpStr = "FdPrdOdNo>= ": adc("wDoOmNoTo").CmpStr = "FdPrdOdNo<= "
  adc("wDoOmSrFr").CmpStr = "FdPrdOdSr>= ": adc("wDoOmSrTo").CmpStr = "FdPrdOdSr<= "
  ws_FgCnd = adc.RepCond
  ws_FgCnd = IIF(ws_FgCnd = "", "", " and " + ws_FgCnd)

  If moCn.RecSeek("Select 'x' From Fgd where FdCoCd= '" + gs_CoCd + "' " + ws_FgCnd + _
                  "and (FdPrdOdTc<> FdExpOdTc or FdPrdOdYy<> FdExpOdYy or " + _
                  "FdPrdOdChr<> FdExpOdChr or FdPrdOdNo<> FdExpOdNo or FdPrdOdSr<> FdExpOdSr) " + _
                  "and not exists(Select 'x' From OrdDsg where OdCoCd= FdCoCd and " + _
                  "OdTc= FdExpOdTc and OdYy= FdExpOdYy and OdChr= FdExpOdChr and " + _
                  "OdNo= FdExpOdNo and OdSr= FdExpOdSr " + ws_OrdCnd + ")") Then
    TxtDisp.text = "Some Orders In The Scope Have Fg Recs Where PrdOrd <> ExpOrd and ExpOrd Not In Scope Of Deletion. Such Orders Cannot Be Deleted"
    DispMsg "Some Orders In The Scope Have Fg Recs Where PrdOrd <> ExpOrd and ExpOrd Not In Scope Of Deletion. Such Orders Cannot Be Deleted", etError
    Exit Sub
  End If

  adc("wDoOmTcFr").CmpStr = "BOdTc>= ": adc("wDoOmTcTo").CmpStr = "BOdTc<= "
  adc("wDoOmYyFr").CmpStr = "BOdYy>= ": adc("wDoOmYyTo").CmpStr = "BOdYy<= "
  adc("wDoOmChrFr").CmpStr = "BOdChr>= ": adc("wDoOmChrTo").CmpStr = "BOdChr<= "
  adc("wDoOmNoFr").CmpStr = "BOdNo>= ": adc("wDoOmNoTo").CmpStr = "BOdNo<= "
  adc("wDoOmSrFr").CmpStr = "BOdSr>= ": adc("wDoOmSrTo").CmpStr = "BOdSr<= "
  ws_BagCnd = adc.RepCond
  ws_BagCnd = IIF(ws_BagCnd = "", "", " and " + ws_BagCnd)
  
  '*** (Bef 2.12 Patch 4)
  'If moCn.RecSeek("Select 'x' From Bag where BCoCd= '" + gs_CoCd + "' " + ws_BagCnd + _
  '              "and exists(Select 'x' From Txnd Td where TdCoCd= BCoCd and TdBYy= BYy " + _
  '              "and TdBChr= BChr and TdBNo= BNo and (Select PValue From Param " + _
  '              "where PTyp= 'TC' and PMCd= TdTc)= 'BS' " + _
  '              "and exists(Select 'x' From Txnd Td1 where Td1.TdCoCd= Td.TdCoCd " + _
  '              "and Td1.TdTc= Td.TdTc and Td1.TdYy= Td.TdYy and Td1.TdChr= Td.TdChr " + _
  '              "and Td1.TdNo= Td.TdNo and not exists(Select 'x' From Bag B join OrdDsg on " + _
  '              "OdCOCd= B.BCoCd and OdTc= B.BOdTc and OdYy= B.BOdYy and OdChr= B.BOdChr " + _
  '              "and OdNo= B.BOdNo and OdSr= B.BOdSr where B.BCoCd= Td1.TdCoCd and " + _
  '              "B.BYy= Td1.TdBYy and B.BChr= Td1.TdBChr and B.BNo= Td1.TdBNo " + ws_OrdCnd + "))) ") Then
  '  TxtDisp.text = "Some Orders In The Scope Have Bag Split Recs Whose Partner Recs Cannot Be Deleted. Such Orders Cannot Be Deleted"
  '  DispMsg "Some Orders In The Scope Have Bag Split Recs Whose Partner Recs Cannot Be Deleted. Such Orders Cannot Be Deleted", etError
  '  Exit Sub
  'End If
  '*** (Bef 2.12 Patch 4)
  
  '*** (Bef 2.13)
  ''*** (Jen 2.12 Patch 4)
  'If moCn.RecSeek("Select 'x' From Bag where BCoCd= '" + gs_CoCd + "' " + ws_BagCnd + _
  '                "and exists(Select 'x' From Txnd Td where TdCoCd= BCoCd and TdBYy= BYy " + _
  '                "and TdBChr= BChr and TdBNo= BNo and (Select PValue From Param " + _
  '                "where PTyp= 'TC' and PMCd= TdTc) in ('BS', 'BM') " + _
  '                "and exists(Select 'x' From Txnd Td1 where Td1.TdCoCd= Td.TdCoCd " + _
  '                "and Td1.TdTc= Td.TdTc and Td1.TdYy= Td.TdYy and Td1.TdChr= Td.TdChr " + _
  '                "and Td1.TdNo= Td.TdNo and not exists(Select 'x' From Bag B join OrdDsg on " + _
  '                "OdCoCd= B.BCoCd and OdTc= B.BOdTc and OdYy= B.BOdYy and OdChr= B.BOdChr " + _
  '                "and OdNo= B.BOdNo and OdSr= B.BOdSr where B.BCoCd= Td1.TdCoCd and " + _
  '                "B.BYy= Td1.TdBYy and B.BChr= Td1.TdBChr and B.BNo= Td1.TdBNo " + ws_OrdCnd + "))) ") Then
  '  TxtDisp.text = "Some Orders In The Scope Have Bag Split/ Bag Merge Recs Whose Partner Recs Cannot Be Deleted. Such Orders Cannot Be Deleted"
  '  DispMsg "Some Orders In The Scope Have Bag Split/ Bag Merge Recs Whose Partner Recs Cannot Be Deleted. Such Orders Cannot Be Deleted", etError
  '  Exit Sub
  'End If
  ''*** (Jen 2.12 Patch 4)
  '*** (Bef 2.13)
  
  '*** (Jen 2.13)
  If moCn.RecSeek("Select 'x' From Bag where BCoCd= '" + gs_CoCd + "' " + ws_BagCnd + _
                  "and exists(Select 'x' From Txnd Td where TdCoCd= BCoCd and TdBYy= BYy " + _
                  "and TdBChr= BChr and TdBNo= BNo and (Select PValue From Param " + _
                  "where PTyp= 'TC' and PMCd= TdTc) in ('BS') " + _
                  "and exists(Select 'x' From Txnd Td1 where Td1.TdCoCd= Td.TdCoCd " + _
                  "and Td1.TdTc= Td.TdTc and Td1.TdYy= Td.TdYy and Td1.TdChr= Td.TdChr " + _
                  "and Td1.TdNo= Td.TdNo and not exists(Select 'x' From Bag B join OrdDsg on " + _
                  "OdCoCd= B.BCoCd and OdTc= B.BOdTc and OdYy= B.BOdYy and OdChr= B.BOdChr " + _
                  "and OdNo= B.BOdNo and OdSr= B.BOdSr where B.BCoCd= Td1.TdCoCd and " + _
                  "B.BYy= Td1.TdBYy and B.BChr= Td1.TdBChr and B.BNo= Td1.TdBNo " + ws_OrdCnd + "))) ") Then
    TxtDisp.text = "Some Orders In The Scope Have Bag Split Recs Whose Partner Recs Cannot Be Deleted. Such Orders Cannot Be Deleted"
    DispMsg "Some Orders In The Scope Have Bag Split Recs Whose Partner Recs Cannot Be Deleted. Such Orders Cannot Be Deleted", etError
    Exit Sub
  End If
  
  '****** (added 'and Td1.TdRefKey= Td.TdRefKey' ) ******
  If moCn.RecSeek("Select 'x' From Bag where BCoCd= '" + gs_CoCd + "' " + ws_BagCnd + _
                  "and exists(Select 'x' From Txnd Td where TdCoCd= BCoCd and TdBYy= BYy " + _
                  "and TdBChr= BChr and TdBNo= BNo and (Select PValue From Param " + _
                  "where PTyp= 'TC' and PMCd= TdTc) in ('BM') " + _
                  "and exists(Select 'x' From Txnd Td1 where Td1.TdCoCd= Td.TdCoCd " + _
                  "and Td1.TdTc= Td.TdTc and Td1.TdYy= Td.TdYy and Td1.TdChr= Td.TdChr " + _
                  "and Td1.TdNo= Td.TdNo and Td1.TdRefKey= Td.TdRefKey and not exists(Select 'x' From Bag B join OrdDsg on " + _
                  "OdCoCd= B.BCoCd and OdTc= B.BOdTc and OdYy= B.BOdYy and OdChr= B.BOdChr " + _
                  "and OdNo= B.BOdNo and OdSr= B.BOdSr where B.BCoCd= Td1.TdCoCd and " + _
                  "B.BYy= Td1.TdBYy and B.BChr= Td1.TdBChr and B.BNo= Td1.TdBNo " + ws_OrdCnd + "))) ") Then
    TxtDisp.text = "Some Orders In The Scope Have Bag Merge Recs Whose Partner Recs Cannot Be Deleted. Such Orders Cannot Be Deleted"
    DispMsg "Some Orders In The Scope Have Bag Merge Recs Whose Partner Recs Cannot Be Deleted. Such Orders Cannot Be Deleted", etError
    Exit Sub
  End If
  '*** (Jen 2.13)
  
  ' ***** Manali 3.8.0 - Checked if Bag is Used in trading Conversion
  If moCn.RecSeek("Select 'x' from Bag Where BCnvJtYn='Y' " + ws_BagCnd) Then
    TxtDisp.text = "Some Orders In The Scope Have Bag(s) Used In Trading Conversion, Such Orders Cannor Be Deleted"
    DispMsg "Some Orders In The Scope Have Bag(s) Used In Trading Conversion, Such Orders Cannor Be Deleted", etError
    Exit Sub
  End If
  
  
  '*** Check if Order For The Specified Scope Can Be Deleted

  adc("wDoOmTcFr").CmpStr = "IdExpOdTc>= ": adc("wDoOmTcTo").CmpStr = "IdExpOdTc<= "
  adc("wDoOmYyFr").CmpStr = "IdExpOdYy>= ": adc("wDoOmYyTo").CmpStr = "IdExpOdYy<= "
  adc("wDoOmChrFr").CmpStr = "IdExpOdChr>= ": adc("wDoOmChrTo").CmpStr = "IdExpOdChr<= "
  adc("wDoOmNoFr").CmpStr = "IdExpOdNo>= ": adc("wDoOmNoTo").CmpStr = "IdExpOdNo<= "
  adc("wDoOmSrFr").CmpStr = "IdExpOdSr>= ": adc("wDoOmSrTo").CmpStr = "IdExpOdSr<= "
  ws_InvCnd = adc.RepCond
  ws_InvCnd = IIF(ws_InvCnd = "", "", " and " + ws_InvCnd)

  '*** (Jen 2.13)
  '*** Jay 2.14 Next Patch01 *** [Removed Version Chk]
  'If ms_Ver = "2.13.0" Then
    If moCn.RecSeek("Select 'x' from InvDsg where IdCoCd= '" + gs_CoCd + "' " + ws_InvCnd + _
                    " And exists (Select 'x' from RetMemDet Where RedCoCd= IdCoCd and RedIdTc= IdTc and RedIdYy= IdYy " + _
                    " And RedIdChr= IdChr and RedIdNo= IdNo and RedIdSr= IdSr) ") Then
      TxtDisp.text = "The Invoice Designs Of Some Orders In The Scope Have Return Memo Recs Associated With Them. Such Orders Cannot Be Deleted"
      DispMsg "The Invoice Designs Of Some Orders In The Scope Have Return Memo Recs Associated With Them. Such Orders Cannot Be Deleted", etError
      Exit Sub
    End If
  'End If
  '*** (Jen 2.13)
  
  DoEvents
  MWLib.BeginProcess Me, "Starting With Order Deletion"
  'moCn.BeginTrans (mdoCommitRec)

  moCn.Execute ("MwSp_BefTrgs")

  moCn.Execute ("if exists (select * from sysobjects where id = object_id('dbo.MwSp_DelOrdBags') " + _
                "and sysstat & 0xf = 4) " + _
                "drop procedure [dbo].[MwSp_DelOrdBags] ")

  '*** (Bef 2.12 Patch 4)
  '"Select BYy, BChr, BNo, 3 as qFlag, TdTc as qTc, TdYy as qYy, TdChr as qChr, " + _
  '"TdNo as qNo, TdSr as qSr, TdYy as qYy1, TdKey as qKey, 0 as qFmdKey, " + _
  '"TdYyKey as qYyKey into ##TmpBagTxn "
  '*** (Bef 2.12 Patch 4)

  ws_str = "CREATE PROCEDURE [dbo].[MwSp_DelOrdBags] @ms_CoCd varchar(3), @ms_Tc varchar(3), " + _
           "@ms_Yy varchar(2), @ms_Chr varchar(3), @mi_No int, @mi_Sr int " + _
           "/*With Encryption*/ As Begin " + _
           "if exists(select 'x' from tempdb..sysobjects where name like '##TmpBagTxn' " + _
           "and xtype= 'U') " + _
           "Begin " + _
           "drop table ##TmpBagTxn " + _
           "End " + _
           "Select BYy, BChr, BNo, 3 as qFlag, TdTc as qTc, TdYy as qYy, TdChr as qChr, " + _
           "TdNo as qNo, TdSr as qSr, space(2) as qYy1, 0 as qKey, 0 as qFmdKey, " + _
           "TdYyKey as qYyKey into ##TmpBagTxn " + _
           "From Txnd join Bag on BCoCd= TdCoCd and BYy= TdBYy and BChr= TdBChr and BNo= TdBNo " + _
           "where BCoCd= @ms_CoCd and BOdTc= @ms_Tc and BOdYy= @ms_Yy and BOdChr= @ms_Chr " + _
           "and BOdNo= @mi_No and BOdSr= @mi_Sr and TdSrNo= 0 " + _
           "ORDER BY BOdTc, BOdYy, BOdChr, BOdNo, BOdSr, BYy, BChr, BNo " + _
           "insert into ##TmpBagTxn " + _
           "Select BYy, BChr, BNo, 2 as qFlag, FdTc as qTc, FdYy as qYy, FdChr as qChr, " + _
           "FdNo as qNo, FdSr as qSr, FdYy as qYy1, FdKey as qKey, 0 as qFmdKey, " + _
           "FdBYyKey As qYyKey From Fgd " + _
           "join Bag on BCoCd= FdCoCd and BYy= FdBYy and BChr= FdBChr and BNo= FdBNo " + _
           "where BCoCd= @ms_CoCd and BOdTc= @ms_Tc and BOdYy= @ms_Yy and BOdChr= @ms_Chr " + _
           "and BOdNo= @mi_No and BOdSr= @mi_Sr " + _
           "ORDER BY BOdTc, BOdYy, BOdChr, BOdNo, BOdSr, BYy, BChr, BNo "

  '*** (Bef 2.12 Patch 4)
  '"SELECT qFlag, qTc, qYy, qChr, qNo, qSr, qYy1, qKey, qFmdKey, qYyKey " + _
  '"FROM ##TmpBagTxn Order By BYy, BChr, BNo, qYyKey desc, qFlag asc, " + _
  '"qYy1 desc, qKey desc, qYy desc, qFmdKey desc "
  '*** (Bef 2.12 Patch 4)

  ws_str = ws_str + _
           "insert into ##TmpBagTxn " + _
           "Select BYy, BChr, BNo, 1 as qFlag, FmdTc as qTc, FmdYy as qYy, FmdChr as qChr, " + _
           "FmdNo as qNo, FmdSr as qSr, FmdFdYy as qYy1, FmdFdKey as qKey, " + _
           "FmdKey As qFmdKey, FdBYyKey As qYyKey " + _
           "From Fmd join Fgd on FdCoCd= FmdCoCd and FdYy= FmdFdYy and FdKey= FmdFdKey " + _
           "join Bag on BCoCd= FdCoCd and BYy= FdBYy and BChr= FdBChr and BNo= FdBNo " + _
           "where BCoCd= @ms_CoCd and BOdTc= @ms_Tc and BOdYy= @ms_Yy " + _
           "and BOdChr= @ms_Chr and BOdNo= @mi_No and BOdSr= @mi_Sr " + _
           "ORDER BY BOdTc, BOdYy, BOdChr, BOdNo, BOdSr, BYy, BChr, BNo " + _
           "/*declare @wSqlStr nvarchar(4000) " + _
           " Set @wSqlStr= */ DECLARE Txn_Cursor SCROLL CURSOR FOR " + _
           "SELECT qFlag, qTc, qYy, qChr, qNo, qSr, qYy1, qKey, qFmdKey, qYyKey " + _
           "FROM ##TmpBagTxn Order By BYy, BChr, BNo, qYyKey desc, " + _
           "qYy1 desc, qKey desc, qFlag asc, qYy desc, qFmdKey desc " + _
           "declare @wi_Flag int, @ws_Tc varchar(3), @ws_Yy varchar(2), " + _
           "@ws_Chr varchar(3), @wi_No int, @wi_Sr int, @ws_Yy1 varchar(2), " + _
           "@wi_Key int, @wi_FmdKey int, @ws_YyKey varchar(9) "
  ws_str = ws_str + "OPEN Txn_Cursor " + _
           "/*** Fetch the first row in the cursor. ***/ " + _
           "FETCH FIRST FROM Txn_Cursor " + _
           "INTO @wi_Flag, @ws_Tc, @ws_Yy, @ws_Chr, @wi_No, @wi_Sr, @ws_Yy1, " + _
           "@wi_Key, @wi_FmdKey, @ws_YyKey " + _
           "WHILE @@FETCH_STATUS = 0 " + _
           "Begin " + _
           "if @wi_Flag= 1 " + _
           "Begin " + _
           "delete from Fmd where FmdCoCd= @ms_CoCd and FmdTc= @ws_Tc and FmdYy= @ws_Yy " + _
           "and FmdChr= @ws_Chr and FmdNo= @wi_No and FmdSr= @wi_Sr " + _
           "End " + _
           "if @wi_Flag= 2 " + _
           "Begin " + _
           "delete from FgRm where FrCoCd= @ms_CoCd and FrTc= @ws_Tc and FrYy= @ws_Yy " + _
           "and FrChr= @ws_Chr and FrNo= @wi_No and FrSr= @wi_Sr " + _
           "delete from Fgd where FdCoCd= @ms_CoCd and FdTc= @ws_Tc and FdYy= @ws_Yy " + _
           "and FdChr= @ws_Chr and FdNo= @wi_No and FdSr= @wi_Sr " + _
           "End " + _
           "if @wi_Flag= 3 " + _
           "Begin " + _
           "delete from TxndPrd where TpCoCd= @ms_CoCd and TpTc= @ws_Tc and TpYy= @ws_Yy " + _
           "and TpChr= @ws_Chr and TpNo= @wi_No and TpSr= @wi_Sr and TpSrNo> 0 "
  ws_str = ws_str + "delete from Txnd where TdCoCd= @ms_CoCd and TdTc= @ws_Tc " + _
           "and TdYy= @ws_Yy and TdChr= @ws_Chr and TdNo= @wi_No and TdSr= @wi_Sr " + _
           "and TdSrNo> 0 " + _
           "delete from Txnd where TdCoCd= @ms_CoCd and TdTc= @ws_Tc and TdYy= @ws_Yy " + _
           "and TdChr= @ws_Chr and TdNo= @wi_No and TdSr= @wi_Sr and TdSrNo= 0 " + _
           "End " + _
           "/*** This is executed as long as the previous fetch succeeds. ***/ " + _
           "FETCH NEXT FROM Txn_Cursor " + _
           "INTO @wi_Flag, @ws_Tc, @ws_Yy, @ws_Chr, @wi_No, @wi_Sr, @ws_Yy1, @wi_Key, " + _
           "@wi_FmdKey, @ws_YyKey " + _
           "End " + _
           "Close Txn_Cursor " + _
           "DEALLOCATE Txn_Cursor " + _
           "/* EXEC sp_executesql @wSqlStr, N '@ms_CoCd varchar(10)', @ms_CoCd */ " + _
           "if exists(select 'x' from tempdb..sysobjects where name like '##TmpBagTxn' " + _
           "and xtype= 'U') " + _
           "Begin " + _
           "drop table ##TmpBagTxn " + _
           "End " + _
           "End "

  moCn.Execute (ws_str)

  '### Not required *** Run Temperory Triggers For TxndZ, InvZRm, InvRm, InvLab, InvFgd, InvDsg to take care of InCls= 'Y'

  '*** (Bef 2.12 Patch 4)
  'ws_str = "DECLARE Inv_cursor SCROLL CURSOR FOR " + _
  '         "SELECT IdCoCd, IdTc, IdYy, IdChr, IdNo, IdSr " + _
  '         " " + _
  '         "FROM InvDsg where IdCoCd= '" + gs_CoCd + "' " + ws_InvCnd + _
  '         " ORDER BY IdExpOdTc, IdExpOdYy, IdExpOdChr, IdExpOdNo, IdExpOdSr " + _
  '         "declare @ws_CoCd varchar(3), @ws_Tc varchar(3), @ws_Yy varchar(2), " + _
  '         "@ws_Chr varchar(3), @wi_No int, @wi_Sr int " + _
  '         "OPEN Inv_cursor " + _
  '         "/*** Fetch the first row in the cursor. ***/ " + _
  '         "FETCH FIRST FROM Inv_cursor " + _
  '         "INTO @ws_CoCd, @ws_Tc, @ws_Yy, @ws_Chr, @wi_No, @wi_Sr " + _
  '         "WHILE @@FETCH_STATUS = 0 "
  '
  'ws_str = ws_str + "Begin " + _
  '         "delete from TxndZ where TzCoCd= @ws_CoCd and TzTc= @ws_Tc " + _
  '         "and TzYy= @ws_Yy and TzChr= @ws_Chr and TzNo= @wi_No /*and TzSr= @wi_Sr*/ " + _
  '         "delete from InvZRm where IzCoCd= @ws_CoCd and IzTc= @ws_Tc and IzYy= @ws_Yy " + _
  '         "and IzChr= @ws_Chr and IzNo= @wi_No /*and IzSr= @wi_Sr */ " + _
  '         "delete from InvLab where iLCoCd= @ws_CoCd and iLTc= @ws_Tc and iLYy= @ws_Yy " + _
  '         "and iLChr= @ws_Chr and iLNo= @wi_No and iLSr= @wi_Sr " + _
  '         "delete from InvRm where IrCoCd= @ws_CoCd and IrTc= @ws_Tc and IrYy= @ws_Yy " + _
  '         "and IrChr= @ws_Chr and IrNo= @wi_No and IrSr= @wi_Sr " + _
  '         "delete from InvFgd where IfCoCd= @ws_CoCd and IfTc= @ws_Tc and IfYy= @ws_Yy " + _
  '         "and IfChr= @ws_Chr and IfNo= @wi_No and IfSr= @wi_Sr " + _
  '         "delete from InvDsg where IdCoCd= @ws_CoCd and IdTc= @ws_Tc and IdYy= @ws_Yy " + _
  '         "and IdChr= @ws_Chr and IdNo= @wi_No and IdSr= @wi_Sr " + _
  '         "/*** This is executed as long as the previous fetch succeeds. ***/ " + _
  '         "FETCH NEXT FROM Inv_cursor " + _
  '         "INTO @ws_CoCd, @ws_Tc, @ws_Yy, @ws_Chr, @wi_No, @wi_Sr " + _
  '         "End " + _
  '         "Close Inv_cursor " + _
  '         "DEALLOCATE Inv_cursor "
  '
  'moCn.Execute (ws_str)
  '*** (Bef 2.12 Patch 4)

  '*** (Jen 2.12 Patch 4)
  '6.1
  ws_str = " Set XACT_ABORT ON " + _
           "BEGIN TRAN TxnDelOrd " + _
           "set dateformat " + DtFmtStr() + " " + _
           "SET ANSI_WARNINGS OFF "
  
  ws_str = ws_str + " DECLARE Inv_cursor SCROLL CURSOR FOR " + _
           "SELECT IdCoCd, IdTc, IdYy, IdChr, IdNo, IdSr " + _
           " " + _
           "FROM InvDsg where IdCoCd= '" + gs_CoCd + "' " + ws_InvCnd + _
           " ORDER BY IdExpOdTc, IdExpOdYy, IdExpOdChr, IdExpOdNo, IdExpOdSr " + _
           "declare @ws_InCoCd varchar(3), @ws_InTc varchar(3), @ws_InYy varchar(2), " + _
           "@ws_InChr varchar(3), @wi_InNo int, @wi_InSr int " + _
           "OPEN Inv_cursor " + _
           "/*** Fetch the first row in the cursor. ***/ " + _
           "FETCH FIRST FROM Inv_cursor " + _
           "INTO @ws_InCoCd, @ws_InTc, @ws_InYy, @ws_InChr, @wi_InNo, @wi_InSr " + _
           "WHILE @@FETCH_STATUS = 0 "

  ws_str = ws_str + " Begin " + _
           "delete from TxndZ where TzCoCd= @ws_InCoCd and TzTc= @ws_InTc " + _
           "and TzYy= @ws_InYy and TzChr= @ws_InChr and TzNo= @wi_InNo /*and TzSr= @wi_InSr*/ " + _
           "delete from InvZRm where IzCoCd= @ws_InCoCd and IzTc= @ws_InTc and IzYy= @ws_InYy " + _
           "and IzChr= @ws_InChr and IzNo= @wi_InNo /*and IzSr= @wi_InSr */ " + _
           "delete from InvLab where iLCoCd= @ws_InCoCd and iLTc= @ws_InTc and iLYy= @ws_InYy " + _
           "and iLChr= @ws_InChr and iLNo= @wi_InNo and iLSr= @wi_InSr " + _
           "delete from InvRm where IrCoCd= @ws_InCoCd and IrTc= @ws_InTc and IrYy= @ws_InYy " + _
           "and IrChr= @ws_InChr and IrNo= @wi_InNo and IrSr= @wi_InSr " + _
           "delete from InvFgd where IfCoCd= @ws_InCoCd and IfTc= @ws_InTc and IfYy= @ws_InYy " + _
           "and IfChr= @ws_InChr and IfNo= @wi_InNo and IfSr= @wi_InSr " + _
           "delete from InvDsg where IdCoCd= @ws_InCoCd and IdTc= @ws_InTc and IdYy= @ws_InYy " + _
           "and IdChr= @ws_InChr and IdNo= @wi_InNo and IdSr= @wi_InSr " + _
           "/*** This is executed as long as the previous fetch succeeds. ***/ " + _
           "FETCH NEXT FROM Inv_cursor " + _
           "INTO @ws_InCoCd, @ws_InTc, @ws_InYy, @ws_InChr, @wi_InNo, @wi_InSr " + _
           "End " + _
           "Close Inv_cursor " + _
           "DEALLOCATE Inv_cursor "
  ''moCn.Execute (ws_str)
  '*** (Jen 2.12 Patch 4)

  '##### Not required *** Run Proper Triggers For TxndZ, InvZRm, InvRm, InvLab, InvFgd, InvDsg
  
  '****** (Bef 2.12 Patch 4) ws_str = " DECLARE Ord_Cursor SCROLL CURSOR FOR "  ******
  
  ws_str = ws_str + " DECLARE Ord_Cursor SCROLL CURSOR FOR " + _
           "SELECT OdCoCd, OdTc, OdYy, OdChr, OdNo, OdSr FROM OrdDsg " + _
           "where OdCoCd= '" + gs_CoCd + "' " + ws_OrdCnd + _
           " and exists(Select 'x' from Param where PTyp= 'TC' and PMCd= OdTc and PSCd= '' " + _
           "and PValue= 'SO') ORDER BY OdTc, OdYy, OdChr, OdNo, OdSr " + _
           "declare @ws_CoCd varchar(3), @ws_Tc varchar(3), @ws_Yy varchar(2), " + _
           "@ws_Chr varchar(3), @wi_No int, @wi_Sr int " + _
           "OPEN Ord_Cursor " + _
           "/*** Fetch the first row in the cursor. ***/ " + _
           "FETCH FIRST FROM Ord_Cursor " + _
           "INTO @ws_CoCd, @ws_Tc, @ws_Yy, @ws_Chr, @wi_No, @wi_Sr " + _
           "WHILE @@FETCH_STATUS = 0 " + _
           "Begin " + _
           "delete from InvFgd where IfSrNo= 0 and IfCoCd= @ws_CoCd and " + _
           "exists(Select 'x' from Bag where BCoCd= IfCoCd and BYy= IfFdBYy " + _
           "and BChr= IfFdBChr and BNo= IfFdBNo and BOdTc= @ws_Tc and BOdYy= @ws_Yy " + _
           "and BOdChr= @ws_Chr and BOdNo= @wi_No and BOdSr= @wi_Sr) "
  ws_str = ws_str + "Delete from OrdTgt where OtCoCd= @ws_CoCd and OtTc= @ws_Tc " + _
           "and OtYy= @ws_Yy and OtChr= @ws_Chr and OtNo= @wi_No and OtSr= @wi_Sr " + _
           "execute dbo.MwSp_DelOrdBags @ws_CoCd, @ws_Tc, @ws_Yy, @ws_Chr, @wi_No, @wi_Sr " + _
           "Delete from Bag where BCoCd= @ws_CoCd and BOdTc= @ws_Tc and BOdYy= @ws_Yy " + _
           "and BOdChr= @ws_Chr and BOdNo= @wi_No and BOdSr= @wi_Sr " + _
           "Delete from OrdPlnDt where OpdCoCd= @ws_CoCd and OpdTc= @ws_Tc and " + _
           "OpdYy= @ws_Yy and OpdChr= @ws_Chr and OpdNo= @wi_No and OpdSr= @wi_Sr " + _
           "Delete from OrdPln where OpCoCd= @ws_CoCd and OpTc= @ws_Tc and " + _
           "OpYy= @ws_Yy and OpChr= @ws_Chr and OpNo= @wi_No and OpSr= @wi_Sr " + _
           "Delete from OrdLab where OlCoCd= @ws_CoCd and OlTc= @ws_Tc and " + _
           "OlYy= @ws_Yy and OlChr= @ws_Chr and OlNo= @wi_No and OlSr= @wi_Sr " + _
           "Delete from OrdRm where OrCoCd= @ws_CoCd and OrTc= @ws_Tc and " + _
           "OrYy= @ws_Yy and OrChr= @ws_Chr and OrNo= @wi_No and OrSr= @wi_Sr " + _
           "Delete from OrdAmd where OaCoCd= @ws_CoCd and OaTc= @ws_Tc and " + _
           "OaYy= @ws_Yy and OaChr= @ws_Chr and OaNo= @wi_No and OaSr= @wi_Sr " + _
           "Delete from OrdCt where OctCoCd= @ws_CoCd and OctTc= @ws_Tc and " + _
           "OctYy= @ws_Yy and OctChr= @ws_Chr and OctNo= @wi_No and OctSr= @wi_Sr "
  ws_str = ws_str + "Delete from OrdDsg where OdCoCd= @ws_CoCd and OdTc= @ws_Tc and " + _
           "OdYy= @ws_Yy and OdChr= @ws_Chr and OdNo= @wi_No and OdSr= @wi_Sr " + _
           "Delete from OrdScop where OsCoCd= @ws_CoCd and OsTc= @ws_Tc and " + _
           "OsYy= @ws_Yy and OsChr= @ws_Chr and OsNo= @wi_No " + _
           "Delete from OrdMst where OmCoCd= @ws_CoCd and OmTc= @ws_Tc and " + _
           "OmYy= @ws_Yy and OmChr= @ws_Chr and OmNo= @wi_No and " + _
           "not exists(Select 'x' From OrdDsg where OdCoCd= OmCoCd and " + _
           "OdTc= OmTc and OdYy= OmYy and OdChr= OmChr and OdNo= OmNo) " + _
           "/*** This is executed as long as the previous fetch succeeds. ***/ " + _
           "FETCH NEXT FROM Ord_Cursor " + _
           "INTO @ws_CoCd, @ws_Tc, @ws_Yy, @ws_Chr, @wi_No, @wi_Sr " + _
           "End " + _
           "Close Ord_Cursor " + _
           "DEALLOCATE Ord_Cursor "
    
  '*** (Jen 2.12 Patch 4)
  ws_str = ws_str + " SET ANSI_WARNINGS ON " + _
                    "/*** @@Error ***/ " + _
                    "COMMIT TRAN TxnDelOrd " + _
                    "Set XACT_ABORT OFF "
  '*** (Jen 2.12 Patch 4)
    
  moCn.Execute (ws_str)
  moCn.Execute ("MwSp_CorrectTrgs")

  moCn.Execute ("if exists (select * from sysobjects where id = object_id('dbo.MwSp_DelOrdBags') " + _
                "and sysstat & 0xf = 4) " + _
                "drop procedure [dbo].[MwSp_DelOrdBags] ")

  moCn.Execute ("if exists (select * from sysobjects where id = object_id('dbo.MwSp_BefTrgs') " + _
                "and sysstat & 0xf = 4) " + _
                "drop procedure [dbo].[MwSp_BefTrgs] ")

  'moCn.Execute ("if exists (select * from sysobjects where id = object_id('dbo.MwSp_CorrectTrgs') " + _
                "and sysstat & 0xf = 4) " + _
                "drop procedure [dbo].[MwSp_CorrectTrgs] ")

  'moCn.CommitTrans mdoCommitRec
  MWLib.EndProcess Me

  '*** (Jen 2.12 Patch 4)
  If moCn.RecSeek("Select 'x' From OrdDsg Where OdCoCd= '" + gs_CoCd + "' " + ws_OrdCnd + _
           " and exists(Select 'x' from Param where PTyp= 'TC' and PMCd= OdTc and PSCd= '' " + _
           " and PValue= 'SO')") Then
    TxtDisp.text = "Orders Could Not Be Deleted..."
    DispMsg "Orders Could Not Be Deleted...", etInfo
  Else
    TxtDisp.text = "Orders Have Been Deleted Successfully"
    DispMsg "Orders Have Been Deleted Successfully", etInfo
  End If
  '*** (Jen 2.12 Patch 4)
  
  '*** (Bef 2.12 Patch 4)
  'TxtDisp.text = "Orders Have Been Deleted Successfully"
  'DispMsg "Orders Have Been Deleted Successfully", etInfo
  '*** (Bef 2.12 Patch 4)
  
  Call CmdDelOrd_Click

  Exit Sub
ErrH:
  Call ErrHndlr
  MWLib.EndProcess Me
  ws_str = "Delete From TrgDisa"
  moCn.Execute (ws_str)
  
  '*** (Jen 2.12 Patch 4)
  On Error Resume Next
  moCn.Execute ("MwSp_CorrectTrgs")
  TxtDisp.text = "Trigger Restored Successfully"
  '*** (Jen 2.12 Patch 4)
  
  Call CmdDelOrd_Click
End Sub
Private Sub DeleteOrderWithoutDisablingTrg()
'  Dim ws_str As String, ws_FgCnd As String, ws_BagCnd As String, ws_InvCnd As String
'  Dim ws_OrdCnd As String
'  On Error GoTo errh
'
'  '*** Check if Order For The Specified Scope Can Be Deleted
'  '*** Very Important
'  If Not moCn.RecSeek("Select HCoCd from Head where HCoCd= '" + gs_CoCd + "' and HCd= '" + ctSelfCmCd + "'") Then
'    TxtDisp.Text = "The Company Code '" + gs_CoCd + "' No Longer Exists In The Database. Login Again Using Another Company Code"
'    DispMsg "The Company Code '" + gs_CoCd + "' No Longer Exists In The Database. Login Again Using Another Company Code", etError
'    Call CmdDelOrd_Click: Exit Sub
'  End If
'  '*** Very Important
'  If ms_Ver <> "2.3.0" And ms_Ver <> "2.4.0" And ms_Ver <> "2.5.0" Then
'    TxtDisp.Text = "Current Database Version Is " + ms_Ver + ". This Option Can Be Run Only If The Database Is 2.3.0 or 2.4.0 or 2.5.0"
'    DispMsg "Current Database Version Is " + ms_Ver + ". This Option Can Be Run Only If The Database Is 2.3.0 or 2.4.0 or 2.5.0", etError
'    Call CmdUpdCoCd_Click: Exit Sub
'  End If
'  If ADC("wDoOmTcFr") = "" And ADC("wDoOmYyFr") = "" And ADC("wDoOmChrFr") = "" _
'     And ADC("wDoOmNoFr") = 0 And ADC("wDoOmSrFr") = 0 And ADC("wDoOmTcTo") = "" _
'     And ADC("wDoOmYyTo") = "" And ADC("wDoOmChrTo") = "" And ADC("wDoOmNoTo") = 0 _
'     And ADC("wDoOmSrTo") = 0 Then
'    TxtDisp.Text = "Some Scope Should Be Entered"
'    DispMsg "Some Scope Should Be Entered", etError
'    Exit Sub
'  End If
'
'  'If Not moCn.RecSeek("select 'x' from dbo.sysobjects where " + _
'  '       "id = object_id(N'[dbo].[MwSp_DelOrdBags]') and " + _
'  '       "OBJECTPROPERTY(id, N'IsProcedure') = 1") Then
'  '  TxtDisp.Text = "Procedure 'MwSp_DelOrdBags' Does Not Exist In The Database "
'  '  DispMsg "Procedure 'MwSp_DelOrdBags' Does Not Exist In The Database ", etError
'  '  Call CmdDelOrd_Click: Exit Sub
'  'End If
'
'  ADC("wDoOmTcFr").CmpStr = "FdPrdOdTc>= ": ADC("wDoOmTcTo").CmpStr = "FdPrdOdTc<= "
'  ADC("wDoOmYyFr").CmpStr = "FdPrdOdYy>= ": ADC("wDoOmYyTo").CmpStr = "FdPrdOdYy<= "
'  ADC("wDoOmChrFr").CmpStr = "FdPrdOdChr>= ": ADC("wDoOmChrTo").CmpStr = "FdPrdOdChr<= "
'  ADC("wDoOmNoFr").CmpStr = "FdPrdOdNo>= ": ADC("wDoOmNoTo").CmpStr = "FdPrdOdNo<= "
'  ADC("wDoOmSrFr").CmpStr = "FdPrdOdSr>= ": ADC("wDoOmSrTo").CmpStr = "FdPrdOdSr<= "
'  ws_FgCnd = ADC.RepCond
'  ws_FgCnd = IIf(ws_FgCnd = "", "", " and " + ws_FgCnd)
'
'  If moCn.RecSeek("Select 'x' From Fgd where FdCoCd= '" + gs_CoCd + "' " + ws_FgCnd + _
'                  "and (FdPrdOdTc<> FdExpOdTc or FdPrdOdYy<> FdExpOdYy or " + _
'                  "FdPrdOdChr<> FdExpOdChr or FdPrdOdNo<> FdExpOdNo or FdPrdOdSr<> FdExpOdSr) ") Then
'    TxtDisp.Text = "Some Orders In The Scope Have Fg Records Where Prd Ord <> Exp Ord. Such Orders Cannot Be Deleted"
'    DispMsg "Some Orders In The Scope Have Fg Records Where Prd Ord <> Exp Ord. Such Orders Cannot Be Deleted", etError
'    Exit Sub
'  End If
'
'  ADC("wDoOmTcFr").CmpStr = "BOdTc>= ": ADC("wDoOmTcTo").CmpStr = "BOdTc<= "
'  ADC("wDoOmYyFr").CmpStr = "BOdYy>= ": ADC("wDoOmYyTo").CmpStr = "BOdYy<= "
'  ADC("wDoOmChrFr").CmpStr = "BOdChr>= ": ADC("wDoOmChrTo").CmpStr = "BOdChr<= "
'  ADC("wDoOmNoFr").CmpStr = "BOdNo>= ": ADC("wDoOmNoTo").CmpStr = "BOdNo<= "
'  ADC("wDoOmSrFr").CmpStr = "BOdSr>= ": ADC("wDoOmSrTo").CmpStr = "BOdSr<= "
'  ws_BagCnd = ADC.RepCond
'  ws_BagCnd = IIf(ws_BagCnd = "", "", " and " + ws_BagCnd)
'  If moCn.RecSeek("Select 'x' From Bag where BCoCd= '" + gs_CoCd + "' " + ws_BagCnd + _
'                  "and exists(Select 'x' From Txnd where TdCoCd= BCoCd and TdBYy= BYy " + _
'                  "and TdBChr= BChr and TdBNo= BNo and (Select PValue From Param " + _
'                  "where PTyp= 'TC' and PMCd= TdTc)= 'BS') ") Then
'    TxtDisp.Text = "Some Orders In The Scope Have Bag Split Entries. Such Orders Cannot Be Deleted"
'    DispMsg "Some Orders In The Scope Have Bag Split Entries. Such Orders Cannot Be Deleted", etError
'    Exit Sub
'  End If
'  '*** Check if Order For The Specified Scope Can Be Deleted
'
'  ADC("wDoOmTcFr").CmpStr = "IdExpOdTc>= ": ADC("wDoOmTcTo").CmpStr = "IdExpOdTc<= "
'  ADC("wDoOmYyFr").CmpStr = "IdExpOdYy>= ": ADC("wDoOmYyTo").CmpStr = "IdExpOdYy<= "
'  ADC("wDoOmChrFr").CmpStr = "IdExpOdChr>= ": ADC("wDoOmChrTo").CmpStr = "IdExpOdChr<= "
'  ADC("wDoOmNoFr").CmpStr = "IdExpOdNo>= ": ADC("wDoOmNoTo").CmpStr = "IdExpOdNo<= "
'  ADC("wDoOmSrFr").CmpStr = "IdExpOdSr>= ": ADC("wDoOmSrTo").CmpStr = "IdExpOdSr<= "
'  ws_InvCnd = ADC.RepCond
'  ws_InvCnd = IIf(ws_InvCnd = "", "", " and " + ws_InvCnd)
'
'  ADC("wDoOmTcFr").CmpStr = "OdTc>= ": ADC("wDoOmTcTo").CmpStr = "OdTc<= "
'  ADC("wDoOmYyFr").CmpStr = "OdYy>= ": ADC("wDoOmYyTo").CmpStr = "OdYy<= "
'  ADC("wDoOmChrFr").CmpStr = "OdChr>= ": ADC("wDoOmChrTo").CmpStr = "OdChr<= "
'  ADC("wDoOmNoFr").CmpStr = "OdNo>= ": ADC("wDoOmNoTo").CmpStr = "OdNo<= "
'  ADC("wDoOmSrFr").CmpStr = "OdSr>= ": ADC("wDoOmSrTo").CmpStr = "OdSr<= "
'  ws_OrdCnd = ADC.RepCond
'  ws_OrdCnd = IIf(ws_OrdCnd = "", "", " and " + ws_OrdCnd)
'
'  DoEvents
'  MWLib.BeginProcess Me, "Starting With Order Deletion"
'  'moCn.BeginTrans (mdoCommitRec)
'
'  moCn.Execute ("if exists (select * from sysobjects where id = object_id('dbo.MwSp_DelOrdBags') " + _
'                "and sysstat & 0xf = 4) " + _
'                "drop procedure [dbo].[MwSp_DelOrdBags] ")
'
'  ws_str = "CREATE PROCEDURE [dbo].[MwSp_DelOrdBags] @ms_CoCd varchar(3), @ms_Tc varchar(3), " + _
'           "@ms_Yy varchar(2), @ms_Chr varchar(3), @mi_No int, @mi_Sr int " + _
'           "With Encryption As " + _
'           "if exists(select 'x' from tempdb..sysobjects where name like '##TmpBagTxn' " + _
'           "and xtype= 'U') " + _
'           "Begin " + _
'           "drop table ##TmpBagTxn " + _
'           "End " + _
'           "Select BYy, BChr, BNo, 3 as qFlag, TdTc as qTc, TdYy as qYy, TdChr as qChr, " + _
'           "TdNo as qNo, TdSr as qSr, TdYy as qYy1, TdKey as qKey, 0 as qFmdKey, " + _
'           "TdYyKey as qYyKey into ##TmpBagTxn " + _
'           "From Txnd join Bag on BCoCd= TdCoCd and BYy= TdBYy and BChr= TdBChr and BNo= TdBNo " + _
'           "where BCoCd= @ms_CoCd and BOdTc= @ms_Tc and BOdYy= @ms_Yy and BOdChr= @ms_Chr " + _
'           "and BOdNo= @mi_No and BOdSr= @mi_Sr and TdSrNo= 0 " + _
'           "ORDER BY BOdTc, BOdYy, BOdChr, BOdNo, BOdSr, BYy, BChr, BNo " + _
'           "insert into ##TmpBagTxn " + _
'           "Select BYy, BChr, BNo, 2 as qFlag, FdTc as qTc, FdYy as qYy, FdChr as qChr, " + _
'           "FdNo as qNo, FdSr as qSr, FdYy as qYy1, FdKey as qKey, 0 as qFmdKey, " + _
'           "FdBYyKey As qYyKey From Fgd " + _
'           "join Bag on BCoCd= FdCoCd and BYy= FdBYy and BChr= FdBChr and BNo= FdBNo " + _
'           "where BCoCd= @ms_CoCd and BOdTc= @ms_Tc and BOdYy= @ms_Yy and BOdChr= @ms_Chr " + _
'           "and BOdNo= @mi_No and BOdSr= @mi_Sr " + _
'           "ORDER BY BOdTc, BOdYy, BOdChr, BOdNo, BOdSr, BYy, BChr, BNo "
'
'  ws_str = ws_str + IIf(ms_Ver = "2.4.0" Or ms_Ver = "2.5.0", _
'           "insert into ##TmpBagTxn " + _
'           "Select BYy, BChr, BNo, 1 as qFlag, FmdTc as qTc, FmdYy as qYy, FmdChr as qChr, " + _
'           "FmdNo as qNo, FmdSr as qSr, FmdFdYy as qYy1, FmdFdKey as qKey, " + _
'           "FmdKey As qFmdKey, FdBYyKey As qYyKey " + _
'           "From Fmd join Fgd on FdCoCd= FmdCoCd and FdYy= FmdFdYy and FdKey= FmdFdKey " + _
'           "join Bag on BCoCd= FdCoCd and BYy= FdBYy and BChr= FdBChr and BNo= FdBNo " + _
'           "where BCoCd= @ms_CoCd and BOdTc= @ms_Tc and BOdYy= @ms_Yy " + _
'           "and BOdChr= @ms_Chr and BOdNo= @mi_No and BOdSr= @mi_Sr " + _
'           "ORDER BY BOdTc, BOdYy, BOdChr, BOdNo, BOdSr, BYy, BChr, BNo ", " ") + _
'           "/*declare @wSqlStr nvarchar(4000) " + _
'           " Set @wSqlStr= */ DECLARE Txn_Cursor SCROLL CURSOR FOR " + _
'           "SELECT qFlag, qTc, qYy, qChr, qNo, qSr, qYy1, qKey, qFmdKey, qYyKey " + _
'           "FROM ##TmpBagTxn Order By BYy, BChr, BNo, qYyKey desc, qFlag asc, " + _
'           "qYy1 desc, qKey desc, qYy desc, qFmdKey desc " + _
'           "declare @wi_Flag int, @ws_Tc varchar(3), @ws_Yy varchar(2), " + _
'           "@ws_Chr varchar(3), @wi_No int, @wi_Sr int, @ws_Yy1 varchar(2), " + _
'           "@wi_Key int, @wi_FmdKey int, @ws_YyKey varchar(9) "
'  ws_str = ws_str + "OPEN Txn_Cursor " + _
'           "/*** Fetch the first row in the cursor. ***/ " + _
'           "FETCH FIRST FROM Txn_Cursor " + _
'           "INTO @wi_Flag, @ws_Tc, @ws_Yy, @ws_Chr, @wi_No, @wi_Sr, @ws_Yy1, " + _
'           "@wi_Key, @wi_FmdKey, @ws_YyKey " + _
'           "WHILE @@FETCH_STATUS = 0 " + _
'           "Begin " + _
'           IIf(ms_Ver = "2.4.0" Or ms_Ver = "2.5.0", "if @wi_Flag= 1 " + _
'           "Begin " + _
'           "delete from Fmd where FmdCoCd= @ms_CoCd and FmdTc= @ws_Tc and FmdYy= @ws_Yy " + _
'           "and FmdChr= @ws_Chr and FmdNo= @wi_No and FmdSr= @wi_Sr " + _
'           "End ", " ") + _
'           "if @wi_Flag= 2 " + _
'           "Begin " + _
'           "delete from FgRm where FrCoCd= @ms_CoCd and FrTc= @ws_Tc and FrYy= @ws_Yy " + _
'           "and FrChr= @ws_Chr and FrNo= @wi_No and FrSr= @wi_Sr " + _
'           "delete from Fgd where FdCoCd= @ms_CoCd and FdTc= @ws_Tc and FdYy= @ws_Yy " + _
'           "and FdChr= @ws_Chr and FdNo= @wi_No and FdSr= @wi_Sr " + _
'           "End " + _
'           "if @wi_Flag= 3 " + _
'           "Begin " + _
'           "delete from TxndPrd where TpCoCd= @ms_CoCd and TpTc= @ws_Tc and TpYy= @ws_Yy " + _
'           "and TpChr= @ws_Chr and TpNo= @wi_No and TpSr= @wi_Sr and TpSrNo> 0 "
'  ws_str = ws_str + "delete from Txnd where TdCoCd= @ms_CoCd and TdTc= @ws_Tc " + _
'           "and TdYy= @ws_Yy and TdChr= @ws_Chr and TdNo= @wi_No and TdSr= @wi_Sr " + _
'           "and TdSrNo> 0 " + _
'           "delete from Txnd where TdCoCd= @ms_CoCd and TdTc= @ws_Tc and TdYy= @ws_Yy " + _
'           "and TdChr= @ws_Chr and TdNo= @wi_No and TdSr= @wi_Sr and TdSrNo= 0 " + _
'           "End " + _
'           "/*** This is executed as long as the previous fetch succeeds. ***/ " + _
'           "FETCH NEXT FROM Txn_Cursor " + _
'           "INTO @wi_Flag, @ws_Tc, @ws_Yy, @ws_Chr, @wi_No, @wi_Sr, @ws_Yy1, @wi_Key, " + _
'           "@wi_FmdKey, @ws_YyKey " + _
'           "End " + _
'           "Close Txn_Cursor " + _
'           "DEALLOCATE Txn_Cursor " + _
'           "/* EXEC sp_executesql @wSqlStr, " + _
'           "N '@ms_CoCd varchar(10)', @ms_CoCd */ " + _
'           "if exists(select 'x' from tempdb..sysobjects where name like '##TmpBagTxn' " + _
'           "and xtype= 'U') " + _
'           "Begin " + _
'           "drop table ##TmpBagTxn " + _
'           "End "
'
'  moCn.Execute (ws_str)
'
'  '###Not required ***  Run Temperory Triggers For TxndZ, InvZRm, InvRm, InvLab, InvFgd, InvDsg to take care of InCls= 'Y'
'  ws_str = "Alter table InvHd disable trigger InvHd_InUpDl " + _
'           "Update InvHd set InCls= '' where InCls= 'Y' and exists(Select 'x' from InvDsg " + _
'           "where IdCoCd= InCoCd and IdTc= InTc and IdYy= InYy and IdChr= InChr " + _
'           "and IdNo= InNo " + ws_InvCnd + ") " + _
'           "Alter table InvHd enable trigger InvHd_InUpDl " + _
'           "DECLARE Inv_cursor SCROLL CURSOR FOR " + _
'           "SELECT IdCoCd, IdTc, IdYy, IdChr, IdNo, IdSr " + _
'           " " + _
'           "FROM InvDsg where IdCoCd= '" + gs_CoCd + "' " + ws_InvCnd + _
'           " ORDER BY IdExpOdTc, IdExpOdYy, IdExpOdChr, IdExpOdNo, IdExpOdSr " + _
'           "declare @ws_CoCd varchar(3), @ws_Tc varchar(3), @ws_Yy varchar(2), " + _
'           "@ws_Chr varchar(3), @wi_No int, @wi_Sr int " + _
'           "OPEN Inv_cursor " + _
'           "/*** Fetch the first row in the cursor. ***/ " + _
'           "FETCH FIRST FROM Inv_cursor " + _
'           "INTO @ws_CoCd, @ws_Tc, @ws_Yy, @ws_Chr, @wi_No, @wi_Sr " + _
'           "WHILE @@FETCH_STATUS = 0 "
'
'  ws_str = ws_str + "Begin " + _
'           "delete from TxndZ where TzCoCd= @ws_CoCd and TzTc= @ws_Tc " + _
'           "and TzYy= @ws_Yy and TzChr= @ws_Chr and TzNo= @wi_No and TzSr= @wi_Sr " + _
'           "delete from InvZRm where IzCoCd= @ws_CoCd and IzTc= @ws_Tc and IzYy= @ws_Yy " + _
'           "and IzChr= @ws_Chr and IzNo= @wi_No and IzSr= @wi_Sr " + _
'           "delete from InvLab where iLCoCd= @ws_CoCd and iLTc= @ws_Tc and iLYy= @ws_Yy " + _
'           "and iLChr= @ws_Chr and iLNo= @wi_No and iLSr= @wi_Sr " + _
'           "delete from InvRm where IrCoCd= @ws_CoCd and IrTc= @ws_Tc and IrYy= @ws_Yy " + _
'           "and IrChr= @ws_Chr and IrNo= @wi_No and IrSr= @wi_Sr " + _
'           "delete from InvFgd where IfCoCd= @ws_CoCd and IfTc= @ws_Tc and IfYy= @ws_Yy " + _
'           "and IfChr= @ws_Chr and IfNo= @wi_No and IfSr= @wi_Sr " + _
'           "delete from InvDsg where IdCoCd= @ws_CoCd and IdTc= @ws_Tc and IdYy= @ws_Yy " + _
'           "and IdChr= @ws_Chr and IdNo= @wi_No and IdSr= @wi_Sr " + _
'           "/*** This is executed as long as the previous fetch succeeds. ***/ " + _
'           "FETCH NEXT FROM Inv_cursor " + _
'           "INTO @ws_CoCd, @ws_Tc, @ws_Yy, @ws_Chr, @wi_No, @wi_Sr " + _
'           "End " + _
'           "Close Inv_cursor " + _
'           "DEALLOCATE Inv_cursor " + _
'           "Alter table InvHd disable trigger InvHd_InUpDl " + _
'           "Update InvHd set InCls= 'Y' where InCls= '' " + _
'           "Alter table InvHd enable trigger InvHd_InUpDl "
'
'  moCn.Execute (ws_str)
'  '###Not required *** Run Proper Triggers For TxndZ, InvZRm, InvRm, InvLab, InvFgd, InvDsg
'
'  ws_str = "DECLARE Ord_Cursor SCROLL CURSOR FOR " + _
'           "SELECT OdCoCd, OdTc, OdYy, OdChr, OdNo, OdSr FROM OrdDsg " + _
'           "where OdCoCd= '" + gs_CoCd + "' " + ws_OrdCnd + _
'           " and exists(Select 'x' from Param where PTyp= 'TC' and PMCd= OdTc and PSCd= '' " + _
'           "and PValue= 'SO') ORDER BY OdTc, OdYy, OdChr, OdNo, OdSr " + _
'           "declare @ws_CoCd varchar(3), @ws_Tc varchar(3), @ws_Yy varchar(2), " + _
'           "@ws_Chr varchar(3), @wi_No int, @wi_Sr int " + _
'           "OPEN Ord_Cursor " + _
'           "/*** Fetch the first row in the cursor. ***/ " + _
'           "FETCH FIRST FROM Ord_Cursor " + _
'           "INTO @ws_CoCd, @ws_Tc, @ws_Yy, @ws_Chr, @wi_No, @wi_Sr " + _
'           "WHILE @@FETCH_STATUS = 0 " + _
'           "Begin " + _
'           "delete from InvFgd where IfSrNo= 0 and IfCoCd= @ws_CoCd and " + _
'           "exists(Select 'x' from Bag where BCoCd= IfCoCd and BYy= IfFdBYy " + _
'           "and BChr= IfFdBChr and BNo= IfFdBNo and BOdTc= @ws_Tc and BOdYy= @ws_Yy " + _
'           "and BOdChr= @ws_Chr and BOdNo= @wi_No and BOdSr= @wi_Sr) "
'  ws_str = ws_str + "Delete from OrdTgt where OtCoCd= @ws_CoCd and OtTc= @ws_Tc " + _
'           "and OtYy= @ws_Yy and OtChr= @ws_Chr and OtNo= @wi_No and OtSr= @wi_Sr " + _
'           "execute dbo.MwSp_DelOrdBags @ws_CoCd, @ws_Tc, @ws_Yy, @ws_Chr, @wi_No, @wi_Sr " + _
'           "Delete from Bag where BCoCd= @ws_CoCd and BOdTc= @ws_Tc and BOdYy= @ws_Yy " + _
'           "and BOdChr= @ws_Chr and BOdNo= @wi_No and BOdSr= @wi_Sr " + _
'           "Delete from OrdPlnDt where OpdCoCd= @ws_CoCd and OpdTc= @ws_Tc and " + _
'           "OpdYy= @ws_Yy and OpdChr= @ws_Chr and OpdNo= @wi_No and OpdSr= @wi_Sr " + _
'           "Delete from OrdPln where OpCoCd= @ws_CoCd and OpTc= @ws_Tc and " + _
'           "OpYy= @ws_Yy and OpChr= @ws_Chr and OpNo= @wi_No and OpSr= @wi_Sr " + _
'           "Delete from OrdLab where OlCoCd= @ws_CoCd and OlTc= @ws_Tc and " + _
'           "OlYy= @ws_Yy and OlChr= @ws_Chr and OlNo= @wi_No and OlSr= @wi_Sr " + _
'           "Delete from OrdRm where OrCoCd= @ws_CoCd and OrTc= @ws_Tc and " + _
'           "OrYy= @ws_Yy and OrChr= @ws_Chr and OrNo= @wi_No and OrSr= @wi_Sr " + _
'           "Delete from OrdAmd where OaCoCd= @ws_CoCd and OaTc= @ws_Tc and " + _
'           "OaYy= @ws_Yy and OaChr= @ws_Chr and OaNo= @wi_No and OaSr= @wi_Sr "
'  ws_str = ws_str + "Delete from OrdDsg where OdCoCd= @ws_CoCd and OdTc= @ws_Tc and " + _
'           "OdYy= @ws_Yy and OdChr= @ws_Chr and OdNo= @wi_No and OdSr= @wi_Sr " + _
'           "/*Delete from OrdScop where OsCoCd= @ws_CoCd and OsTc= @ws_Tc and " + _
'           "OsYy= @ws_Yy and OsChr= @ws_Chr and OsNo= @wi_No */ " + _
'           "Delete from OrdMst where OmCoCd= @ws_CoCd and OmTc= @ws_Tc and " + _
'           "OmYy= @ws_Yy and OmChr= @ws_Chr and OmNo= @wi_No and " + _
'           "not exists(Select 'x' From OrdDsg where OdCoCd= OmCoCd and " + _
'           "OdTc= OmTc and OdYy= OmYy and OdChr= OmChr and OdNo= OmNo) " + _
'           "/*** This is executed as long as the previous fetch succeeds. ***/ " + _
'           "FETCH NEXT FROM Ord_Cursor " + _
'           "INTO @ws_CoCd, @ws_Tc, @ws_Yy, @ws_Chr, @wi_No, @wi_Sr " + _
'           "End " + _
'           "Close Ord_Cursor " + _
'           "DEALLOCATE Ord_Cursor "
'
'  moCn.Execute (ws_str)
'
'  moCn.Execute ("if exists (select * from sysobjects where id = object_id('dbo.MwSp_DelOrdBags') " + _
'                "and sysstat & 0xf = 4) " + _
'                "drop procedure [dbo].[MwSp_DelOrdBags] ")
'
'  'moCn.CommitTrans mdoCommitRec
'  MWLib.EndProcess Me
'
'  TxtDisp.Text = "Orders Have Been Deleted Successfully"
'  DispMsg "Orders Have Been Deleted Successfully", etInfo
'  Call CmdDelOrd_Click
'
'  Exit Sub
'errh:
'  Call ErrHndlr
'  MWLib.EndProcess Me
'  Call CmdDelOrd_Click
End Sub

'*** (Jen 3.01)
Private Sub UpdateRmQw()
  '*** For this update to function the user has to temporarily update the STW

  Dim ws_str As String, ws_Disp As String, ws_LoginUsers As String, ws_ArrUsrDet() As String

  TxtDisp.text = ""
  DispMsg "", etInfo

  If ms_Ver <> "2.14.0" And ms_Ver <> "2.14.2" And ms_Ver <> "3.1.0" Then
    ws_Disp = "Current Database Version Is " + ms_Ver + ". This Option Can Be Run Only If The Database Ver Is 2.14.0, 2.14.2 Or 3.1.0"
    TxtDisp.text = ws_Disp
    DispMsg ws_Disp, etError
    Exit Sub
  End If

  If Trim$(TxtSuppPwd_UpdRmQw.text) <> mcs_Pwd Then
    ws_Disp = "Not a Valid Support Password"
    TxtDisp.text = ws_Disp
    DispMsg ws_Disp, etError
    Exit Sub
  End If

  If moCn.RecSeek("Select 'x' from Head where HCd= '" + ctSelfCmCd + "' and hQwModYN= 'N'") Then
    '/* HCoCd= '" + gs_CoCd + "' and */
    ws_Disp = "The Q/W Module For All RM Categories Has Not Been Enabled In This Database"
    TxtDisp.text = ws_Disp
    DispMsg ws_Disp, etError
    Exit Sub
  End If
  
  If Not moCn.RecSeek("Select 'x' from RmMst where RmCtg in ('D', 'C') " + _
                  "and   ( (RmSTWGrp= 'QQ' and (RmQw= 'W' or RmQwCst= 'W')) " + _
                  "         or (RmSTWGrp= 'QW' and RmQw= 'W') " + _
                  "         or (RmSTWGrp= 'WQ' and RmQwCst= 'W') " + _
                  "      )  ") Then
  
    ws_Disp = "No RM Code has been marked for update. First Temporarily Update The STW Group of the required RM Codes as specified in the Instruction Above."
    TxtDisp.text = ws_Disp
    DispMsg ws_Disp, etError
    Exit Sub
  End If

  On Error GoTo ErrH
  DoEvents
  MWLib.BeginProcess Me, "Starting With Rm QW Update"
  'moCn.BeginTrans (mdoCommitRec)

  'moCn.Execute ("update MwMst..MSys set MsLoginNm= MsLoginNm+ '#' where MsSrvrNm= 'EMR' and MsLoginNm= '" + MWLib.DbNm + "'")
  moCn.Execute ("update Param set PValidYN= 'N' where PTyp= 'VER' and PMCd= 'DATVER' and PSCd= ''")

  'ws_LoginUsers = moCn.GetFldVal("SELECT dbo.MwFn_LoginDets ('Emr301', 'UsrApp')")
  ws_LoginUsers = moCn.GetFldVal("SELECT dbo.MwFn_LoginDets ('" + MWLib.DbNm + "', 'UsrApp')")
  
  ws_ArrUsrDet = Split(ws_LoginUsers, "|")

  If UBound(ws_ArrUsrDet) = 2 Then
    If Val(ws_ArrUsrDet(0)) > 1 Then
      ws_Disp = "More Than One User Is Logged On Through Emperor To The Database '" + MWLib.DbNm + "'. " + _
                "This Update Can Proceed Only If All Users Log Out Of The System." + _
                "The List Of Machines Logged On To The System Are: " + ws_ArrUsrDet(1)
      TxtDisp.text = ws_Disp
      DispMsg ws_Disp, etError
      'moCn.Execute ("update MwMst..MSys set MsLoginNm= replace(MsLoginNm, '#', '') where MsSrvrNm= 'EMR' and MsLoginNm= '" + MWLib.DbNm + "'+ '#'")
      moCn.Execute ("update Param set PValidYN= '' where PTyp= 'VER' and PMCd= 'DATVER' and PSCd= ''")
      Exit Sub
    'Else
    '  moCn.Execute ("update MwMst..MSys set MsLoginNm= MsLoginNm+ '#' where MsSrvrNm= 'EMR' and MsLoginNm= '" + MWLib.DbNm + "'")
    End If
  End If

  If moCn.GetFldVal("Select count('x') from UsrLogin") > 1 Then
      ws_Disp = "More Than One User Is Logged On Through Emperor To The Database '" + MWLib.DbNm + "'. " + _
                "This Update Can Proceed Only If All Users Log Out Of The System."
      TxtDisp.text = ws_Disp
      DispMsg ws_Disp, etError
      moCn.Execute ("update Param set PValidYN= '' where PTyp= 'VER' and PMCd= 'DATVER' and PSCd= ''")
      Exit Sub
  End If

  'ws_Str = "if exists (select * from sysobjects where id = object_id('dbo.MwSp_UpdateQw') and OBJECTPROPERTY(Id, N'IsProcedure') = 1)" + _
  '         "  drop procedure [dbo].[MwSp_UpdateQw]"
  'moCn.Execute (ws_Str)
'6.1
  ws_str = " Set XACT_ABORT ON " + _
           "BEGIN TRAN TUpdQw " + _
           "set dateformat " + DtFmtStr() + " " + _
           "SET ANSI_WARNINGS OFF "

  ws_str = ws_str + " if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[TmpRmMstQw]') " + _
           "and OBJECTPROPERTY(id, N'IsUserTable') = 1) " + _
           "drop table [dbo].[TmpRmMstQw] "

  ws_str = ws_str + " Select *, (case when RmSTWGrp in ('QQ', 'QW') and RmQw= 'W' then 'Y' else 'N' end) as qUpdSal, " + _
           "(case when RmSTWGrp in ('QQ', 'WQ') and RmQwCst= 'W' then 'Y' else 'N' end) as qUpdCst " + _
           "Into dbo.TmpRmMstQw " + _
           "from RmMst where RmCtg in ('D', 'C') " + _
           "  and   ( (RmSTWGrp= 'QQ' and (RmQw= 'W' or RmQwCst= 'W')) " + _
           "          or (RmSTWGrp= 'QW' and RmQw= 'W') " + _
           "          or (RmSTWGrp= 'WQ' and RmQwCst= 'W') " + _
           "        ) "

  ws_str = ws_str + " insert into TrgDisa values ('RmMst') " + _
           "update RmMst set RmQw= (case when qUpdSal= 'Y' then 'Q' else RmMst.RmQw end), " + _
           "RmQwCst= (case when qUpdCst= 'Y' then 'Q' else RmMst.RmQwCst end) " + _
           "from TmpRmMstQw Tp  " + _
           "where RmMst.RmCd= Tp.RmCd And (qUpdSal= 'Y' or qUpdCst= 'Y') " + _
           "delete from TrgDisa "

  ws_str = ws_str + " insert into TrgDisa values ('OrdRm') " + _
           "declare @wi_RmSalQty int, @wi_RmPrdQty int " + _
           "set @wi_RmSalQty= 0 " + _
           "set @wi_RmPrdQty= 0 " + _
           "update OrdRm set " + _
           "@wi_RmSalQty= (case when qUpdSal= 'Y' and OrQty= 0 then 1 else OrQty end), " + _
           "@wi_RmPrdQty= (case when qUpdCst= 'Y' and OrPrdQty= 0 then 1 else OrPrdQty end), " + _
           "OrQty= (case when HShowPrdRmDets= 'N' then (case when @wi_RmSalQty>= @wi_RmPrdQty then @wi_RmSalQty else @wi_RmPrdQty end) else @wi_RmSalQty end), " + _
           "OrPrdQty= (case when HShowPrdRmDets= 'N' then (case when @wi_RmSalQty>= @wi_RmPrdQty then @wi_RmSalQty else @wi_RmPrdQty end) else @wi_RmPrdQty end) " + _
           "from TmpRmMstQw Tp join Head on HCoCd= '" + ctSelfCoCd + "' " + _
           "where OrdRm.OrRmCd= Tp.RmCd And (qUpdSal= 'Y' or qUpdCst= 'Y') "

  ws_str = ws_str + " update OrdRm set " + _
           "OrSalRt= (case when qUpdSal= 'Y' then convert (decimal (14, 3), (case when OrQty= 0 then OrSalVal else OrSalVal/ convert (decimal (14, 2), OrQty) end)) else OrSalRt end), " + _
           "OrCstRt= (case when qUpdCst= 'Y' then convert (decimal (14, 3), (case when OrPrdQty= 0 then OrCstVal else OrCstVal/ convert (decimal (14, 2), OrPrdQty) end)) else OrCstRt end) " + _
           "from TmpRmMstQw Tp " + _
           "where OrdRm.OrRmCd= Tp.RmCd And (qUpdSal= 'Y' or qUpdCst= 'Y') " + _
           "delete from TrgDisa "

  ws_str = ws_str + " insert into TrgDisa values ('InvRm') " + _
           "update InvRm set " + _
           "IrRmARt= (case when qUpdCst= 'Y' then convert (decimal (14, 3), (case when IrRmQty= 0 then IrRmAVal else IrRmAVal/ convert (decimal (14, 2), IrRmQty) end)) else IrRmARt end), " + _
           "IrRmIRt= (case when qUpdSal= 'Y' then convert (decimal (14, 3), (case when IrRmQty= 0 then IrRmIVal else IrRmIVal/ convert (decimal (14, 2), IrRmQty) end)) else IrRmIRt end), " + _
           "IrRmZRt= (case when qUpdSal= 'Y' then convert (decimal (14, 3), (case when IrRmQty= 0 then IrRmZVal else IrRmZVal/ convert (decimal (14, 2), IrRmQty) end)) else IrRmZRt end), " + _
           "IrRmQty= (case when (qUpdSal= 'Y' or qUpdCst= 'Y') and IrRmQty= 0 then 1 else IrRmQty end) " + _
           "from TmpRmMstQw Tp " + _
           "where InvRm.IrRmCd= Tp.RmCd And (qUpdSal= 'Y' or qUpdCst= 'Y') " + _
           "delete from TrgDisa "

  ws_str = ws_str + " insert into TrgDisa values ('InvZRm') " + _
           "update InvZRm set IzRmQty= IsNull((Select sum(IrRmQty) From InvRm " + _
           "Where IrCoCd = IzCoCd And IrTc = IzTc And IrYy = IzYy And IrChr = IzChr And IrNo = IzNo " + _
           "and IrRmZCd= IzRmZCd and convert (decimal (14, 3), IrRmZPureRt)= convert (decimal (14, 3), IzRmZRt) " + _
           "), 0) " + _
           "where IzRmQty<> IsNull((Select sum(IrRmQty) From InvRm " + _
           "Where IrCoCd = IzCoCd And IrTc = IzTc And IrYy = IzYy And IrChr = IzChr And IrNo = IzNo " + _
           "and IrRmZCd= IzRmZCd and convert (decimal (14, 3), IrRmZPureRt)= convert (decimal (14, 3), IzRmZRt) " + _
           "), 0) " + _
           "delete from TrgDisa "

  ws_str = ws_str + " if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[TmpRmMstQw]') " + _
           "and OBJECTPROPERTY(id, N'IsUserTable') = 1) " + _
           "drop table [dbo].[TmpRmMstQw] "

  ws_str = ws_str + " SET ANSI_WARNINGS ON " + _
                    "COMMIT TRAN TUpdQw " + _
                    "Set XACT_ABORT OFF "

  moCn.Execute (ws_str)

  'moCn.Execute ("update MwMst..MSys set MsLoginNm= replace(MsLoginNm, '#', '') where MsSrvrNm= 'EMR' and MsLoginNm= '" + MWLib.DbNm + "'+ '#'")
  moCn.Execute ("update Param set PValidYN= '' where PTyp= 'VER' and PMCd= 'DATVER' and PSCd= ''")

  'moCn.CommitTrans mdoCommitRec
  MWLib.EndProcess Me

  If moCn.RecSeek("Select 'x' from RmMst where RmCtg in ('D', 'C') " + _
                  "and   ( (RmSTWGrp= 'QQ' and (RmQw= 'W' or RmQwCst= 'W')) " + _
                  "         or (RmSTWGrp= 'QW' and RmQw= 'W') " + _
                  "         or (RmSTWGrp= 'WQ' and RmQwCst= 'W') " + _
                  "      )  ") Then
    ws_Disp = "Rm QW Could Not Be Updated..."
    TxtDisp.text = ws_Disp
    DispMsg ws_Disp, etInfo
  Else
    ws_Disp = "Rm QW Have Been Successfully Updated"
    TxtDisp.text = ws_Disp
    DispMsg ws_Disp, etInfo
  End If

  Call CmdUpdRmQw_Click

  Exit Sub
ErrH:
  Call ErrHndlr
  MWLib.EndProcess Me
  On Error Resume Next
  'moCn.Execute ("update MwMst..MSys set MsLoginNm= replace(MsLoginNm, '#', '') where MsSrvrNm= 'EMR' and MsLoginNm= '" + MWLib.DbNm + "'+ '#'")
  moCn.Execute ("update Param set PValidYN= '' where PTyp= 'VER' and PMCd= 'DATVER' and PSCd= ''")
  moCn.Execute ("Delete From TrgDisa")

  Call CmdUpdRmQw_Click
End Sub
'*** (Jen 3.01)

'****** Sachin 2.12 - For Delete Order
Public Function ReadFileContents(FilePath As String) As String

    On Error GoTo ErrorHandler
    Dim iFileNumber As Integer
    Dim sAns As Variant
    
    If Dir(FilePath) = "" Then Exit Function
    iFileNumber = FreeFile
    Open FilePath For Input As #iFileNumber
    sAns = Input(LOF(iFileNumber), #iFileNumber)
    ReadFileContents = sAns
    
ErrorHandler:
    Close #iFileNumber
    
End Function
'****** Sachin 2.12 - For Delete Order


Private Sub ErrHndlr()
  Dim ws_ReturnStr As String
  If InStr(1, Err.Description, "|") > 0 Then
    ws_ReturnStr = Mid(Err.Description, InStr(1, Err.Description, "|") + 1)
  Else
    Dim wf_Ctr As Single, ws_ErrMsg As String
    For wf_Ctr = 0 To moCn.ErrCount - 1
      If InStr(1, moCn.ErrObj(wf_Ctr).Description, "|") > 0 Then
        ws_ErrMsg = Mid(moCn.ErrObj(wf_Ctr).Description, InStr(1, moCn.ErrObj(wf_Ctr).Description, "|") + 1)
      Else
        ws_ErrMsg = moCn.ErrObj(wf_Ctr).Description
      End If
      ws_ReturnStr = ws_ReturnStr + CStr(wf_Ctr + 1) + "> " + ws_ErrMsg + vbCrLf
    Next wf_Ctr
  End If
  If ws_ReturnStr = "" Then ws_ReturnStr = Err.Description
  TxtDisp.text = ws_ReturnStr

  DoEvents

  DispMsg ws_ReturnStr, etInfo
  moCn.ErrClear
End Sub
Private Sub DispFra(ByVal pv_VerUpdRtFra As en_Upd)
  '*** This is the common procedure which is used to toggle between hiding and displaying
      'the frames (i.e. if the frame is visible then make it invisible and viseversa)
      'associated with a command button
  Select Case pv_VerUpdRtFra
  Case Is = UpdCoCd
    '*** If Frame FraUpdCoCd is visible then make it invisible and enable all the command buttons
    '*** If Frame FraUpdCoCd is invisible then make it visible and disable all the
        'command buttons except the one which is clicked. Change the background colour of
        'the clicked command button to white and make the font bold
    If FraUpdCoCd.Visible = True Then
      FraUpdCoCd.Visible = False
      FraUpdCoCd.Enabled = False
      Call EnaDisaCmds(False, CmdUpdCoCd)
      CmdUpdCoCd.SetFocus
    Else
      FraUpdCoCd.Visible = True
      FraUpdCoCd.Enabled = True
      CmdUpdCoCdGo.SetFocus
      FraUpdCoCd.ZOrder
      Call EnaDisaCmds(True, CmdUpdCoCd)
    End If
  Case Is = DelOrd
    If FraDelOrd.Visible = True Then
      FraDelOrd.Visible = False
      FraDelOrd.Enabled = False
      Call EnaDisaCmds(False, CmdDelOrd)
      CmdDelOrd.SetFocus
    Else
      FraDelOrd.Visible = True
      FraDelOrd.Enabled = True
      CmdDelOrdGo.SetFocus
      FraDelOrd.ZOrder
      Call EnaDisaCmds(True, CmdDelOrd)
    End If
  Case Is = AvgStkRt
    If FraAvgStkRt.Visible = True Then
      FraAvgStkRt.Visible = False
      FraAvgStkRt.Enabled = False
      Call EnaDisaCmds(False, CmdAvgStkRt)
      CmdAvgStkRt.SetFocus
    Else
      FraAvgStkRt.Visible = True
      FraAvgStkRt.Enabled = True
      CmdAvgStkRtGo.SetFocus
      FraAvgStkRt.ZOrder
      Call EnaDisaCmds(True, CmdAvgStkRt)
    End If
  Case Is = PurPlt
    If FraPurPlt.Visible = True Then
      FraPurPlt.Visible = False
      FraPurPlt.Enabled = False
      Call EnaDisaCmds(False, CmdPurPlt)
      CmdPurPlt.SetFocus
    Else
      FraPurPlt.Visible = True
      FraPurPlt.Enabled = True
      CmdPurPltGo.SetFocus
      FraPurPlt.ZOrder
      Call EnaDisaCmds(True, CmdPurPlt)
    End If
  Case Is = UPDUA
    If FraUpdUA.Visible = True Then
      FraUpdUA.Visible = False
      FraUpdUA.Enabled = False
      Call EnaDisaCmds(False, CmdUpdUA)
      CmdUpdUA.SetFocus
    Else
      FraUpdUA.Visible = True
      FraUpdUA.Enabled = True
      CmdUpdUAGo.SetFocus
      FraUpdUA.ZOrder
      Call EnaDisaCmds(True, CmdUpdUA)
    End If
  '***********Geeta********Emr207*******
  Case Is = UpdRmZCd
    If FraUpdRmZCd.Visible = True Then
      FraUpdRmZCd.Visible = False
      FraUpdRmZCd.Enabled = False
      Call EnaDisaCmds(False, CmdUpdRmZCd)
      CmdUpdRmZCd.SetFocus
    Else
      FraUpdRmZCd.Visible = True
      FraUpdRmZCd.Enabled = True
      CmdUpdRmZCdGo.SetFocus
      FraUpdRmZCd.ZOrder
      Call EnaDisaCmds(True, CmdUpdRmZCd)
    End If
  Case Is = UpdSYyMm
    If FraUpdSYyMm.Visible = True Then
      FraUpdSYyMm.Visible = False
      FraUpdSYyMm.Enabled = False
      Call EnaDisaCmds(False, CmdUpdSYyMm)
      CmdUpdSYyMm.SetFocus
    Else
      FraUpdSYyMm.Visible = True
      FraUpdSYyMm.Enabled = True
      CmdUpdSYyMmGo.SetFocus
      FraUpdSYyMm.ZOrder
      Call EnaDisaCmds(True, CmdUpdSYyMm)
    End If
  Case Is = UpdPrdStg
    If FraUpdPrdStg.Visible = True Then
      FraUpdPrdStg.Visible = False
      FraUpdPrdStg.Enabled = False
      Call EnaDisaCmds(False, CmdUpdPrdStg)
      CmdUpdPrdStg.SetFocus
    Else
      FraUpdPrdStg.Visible = True
      FraUpdPrdStg.Enabled = True
      CmdUpdPrdStgGo.SetFocus
      FraUpdPrdStg.ZOrder
      Call EnaDisaCmds(True, CmdUpdPrdStg)
    End If
' Zubin 212 (23/02/06)
  Case Is = EnaDisaMod
    If FraEnaDisaMod.Visible = True Then
      FraEnaDisaMod.Visible = False
      FraEnaDisaMod.Enabled = False
      Call EnaDisaCmds(False, CmdEnaDisaMod)
      CmdEnaDisaMod.SetFocus
    Else
      FraEnaDisaMod.Visible = True
      FraEnaDisaMod.Enabled = True
      CmdEnaDisaModGO.SetFocus
      FraEnaDisaMod.ZOrder
      Call EnaDisaCmds(True, CmdEnaDisaMod)
    End If
    ' Zubin 212 (23/02/06)

  '****** Sachin 212 Patch 03 - 29-04-2006
  Case Is = UpdDsgAnaCd
    If FraDsgAna.Visible = True Then
      FraDsgAna.Visible = False
      FraDsgAna.Enabled = False
      Call EnaDisaCmds(False, CmdUpdDsgAna)
      'GrdDsgAnlyFlds.SetFocus
    Else
      FraDsgAna.Visible = True
      FraDsgAna.Enabled = True
      CmdDsgAnaUpdGo.SetFocus
      FraDsgAna.ZOrder
      Call EnaDisaCmds(True, CmdUpdDsgAna)
    End If
    '****** Sachin 212 Patch 03 - 29-04-2006
  
  ' ****** Sachin [31-08-06] Update Ini Date from Maintenance ******
  Case Is = UpdIni
    If FraUpdIni.Visible = True Then
      FraUpdIni.Visible = False
      FraUpdIni.Enabled = False
      Call EnaDisaCmds(False, CmdDllsUpd)
    Else
      FraUpdIni.Visible = True
      FraUpdIni.Enabled = True
      CmdUpdIniGo.SetFocus
      FraUpdIni.ZOrder
      Call EnaDisaCmds(True, CmdDllsUpd)
    End If
  ' ****** Sachin [31-08-06] Update Ini Date from Maintenance ******
  
  ' ****** Sachin 2.14.0 [07/10/06] Prd Purity Update For G,P,S ******
  Case Is = UpdPrdPurity
    If FraUpdPrdPurity.Visible = True Then
      FraUpdPrdPurity.Visible = False
      FraUpdPrdPurity.Enabled = False
      Call EnaDisaCmds(False, CmdUpdPrdPurity)
    Else
      FraUpdPrdPurity.Visible = True
      FraUpdPrdPurity.Enabled = True
      RmPrdPurityUpdGo.SetFocus
      FraUpdPrdPurity.ZOrder
      Call EnaDisaCmds(True, CmdUpdPrdPurity)
    End If
    
  '****** Sachin 2.14.0 Next ******
  Case Is = UpdCoName
    If FraCoName.Visible = True Then
      FraCoName.Visible = False
      FraCoName.Enabled = False
      Call EnaDisaCmds(False, CmdUpdCoName)
    Else
      FraCoName.Visible = True
      FraCoName.Enabled = True
      CmdCoNameUpdGo.SetFocus
      FraCoName.ZOrder
      Call EnaDisaCmds(True, CmdUpdCoName)
    End If
  '****** Sachin 2.14.0 Next ******
  ' ****** Sachin 2.14.0 [07/10/06] Prd Purity Update For G,P,S ******
  
  '*** (Jen 2.14 Next)
  Case Is = UpdCompDsg
    If FraCompUpd.Visible = True Then
      FraCompUpd.Visible = False
      FraCompUpd.Enabled = False
      Call EnaDisaCmds(False, CmdCompUpd)
      CmdCompUpd.SetFocus
    Else
      FraCompUpd.Visible = True
      FraCompUpd.Enabled = True
      adc("wBChrCnt").SetFocus '*** Jay 3.01 *** [Dynamic BCHR]
      FraCompUpd.ZOrder
      Call EnaDisaCmds(True, CmdCompUpd)
    End If
  '*** (Jen 2.14 Next)
  '*** Jay 3.01 *** [Inv-Fg Rm Mismatch]
    Case Is = InvFgRm
    If FraInvFgRmMismatch.Visible = True Then
      FraInvFgRmMismatch.Visible = False
      FraInvFgRmMismatch.Enabled = False
      Call EnaDisaCmds(False, CmdInvFgRmMisMatch)
      CmdInvFgRmMisMatch.SetFocus
    Else
      FraInvFgRmMismatch.Visible = True
      FraInvFgRmMismatch.Enabled = True
      CmdInvFgRmMisMatchGo.SetFocus
      FraInvFgRmMismatch.ZOrder
      Call EnaDisaCmds(True, CmdInvFgRmMisMatch)
    End If
  '*** Jay 3.01 *** [Inv-Fg Rm Mismatch]
  
    '*** (Jen 3.01)
    Case Is = UpdRmQw
    If FraUpdRmQw.Visible = True Then
      FraUpdRmQw.Visible = False
      FraUpdRmQw.Enabled = False
      Call EnaDisaCmds(False, CmdUpdRmQw)
      CmdUpdRmQw.SetFocus
    Else
      FraUpdRmQw.Visible = True
      FraUpdRmQw.Enabled = True
      'CmdUpdRmQwGo.SetFocus
      TxtSuppPwd_UpdRmQw.SetFocus
      FraUpdRmQw.ZOrder
      Call EnaDisaCmds(True, CmdUpdRmQw)
    End If
    '*** (Jen 3.01)
    '*** Jay 3.1.0_1
    Case Is = InsertComp
    If FraCmpCodeIns.Visible = True Then
      FraCmpCodeIns.Visible = False
      FraCmpCodeIns.Enabled = False
      Call EnaDisaCmds(False, CmdInsComp)
      CmdInsComp.SetFocus
    Else
      FraCmpCodeIns.Visible = True
      FraCmpCodeIns.Enabled = True
      'CmdNewCmpGo.SetFocus
      adc("wSrcCoCd").SetFocus
      FraCmpCodeIns.ZOrder
      Call EnaDisaCmds(True, CmdInsComp)
    End If
    '*** Jay 3.1.0_1
  End Select
End Sub
Private Sub EnaDisaCmds(ByVal pv_ShowFra As Boolean, Optional ByVal pv_CmdBut As CommandButton)
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
      pv_CmdBut.BackColor = vbWhite
      pv_CmdBut.FontBold = True
      pv_CmdBut.CausesValidation = True
    Else
      pv_CmdBut.BackColor = &H8000000F
      pv_CmdBut.FontBold = False
      'pv_CmdBut.CausesValidation = False
    End If
  End If
  CmdUpdCoCd.Enabled = Not pv_ShowFra
  CmdDelOrd.Enabled = Not pv_ShowFra
  CmdUpdRmZCd.Enabled = Not pv_ShowFra
  CmdAvgStkRt.Enabled = Not pv_ShowFra
  CmdUpdUA.Enabled = Not pv_ShowFra
  CmdPurPlt.Enabled = Not pv_ShowFra
  CmdUpdSYyMm.Enabled = Not pv_ShowFra
  CmdUpdPrdStg.Enabled = Not pv_ShowFra
  CmdEnaDisaMod.Enabled = Not pv_ShowFra
  CmdUpdDsgAna.Enabled = Not pv_ShowFra
  CmdDllsUpd.Enabled = Not pv_ShowFra
  CmdUpdPrdPurity.Enabled = Not pv_ShowFra
  CmdUpdCoName.Enabled = Not pv_ShowFra         '****** Sachin 2.14.0 Next *****
  CmdCompUpd.Enabled = Not pv_ShowFra           '*** (Jen 2.14 Next)
  CmdInvFgRmMisMatch.Enabled = Not pv_ShowFra   '*** Jay 3.01 *** [Inv-Fg Rm Mismatch]
  CmdUpdRmQw.Enabled = Not pv_ShowFra           '*** (Jen 3.01)
  CmdInsComp.Enabled = Not pv_ShowFra           '*** Jay 3.1.0_1
  TxtDisp.Enabled = Not pv_ShowFra
  If Not pv_CmdBut Is Nothing Then pv_CmdBut.Enabled = True
End Sub
Private Sub HideAllFras()
  '*** This procedure hides all the frames
  FraUpdCoCd.Visible = False
  FraUpdCoCd.Enabled = False
  FraDelOrd.Visible = False
  FraDelOrd.Enabled = False
  FraAvgStkRt.Visible = False
  FraAvgStkRt.Enabled = False
  FraPurPlt.Visible = False
  FraPurPlt.Enabled = False
  FraUpdUA.Visible = False
  FraUpdUA.Enabled = False
  FraUpdRmZCd.Visible = False
  FraUpdRmZCd.Enabled = False
  FraUpdSYyMm.Visible = False
  FraUpdSYyMm.Enabled = False
  FraDsgAna.Visible = False
  FraDsgAna.Enabled = False
  
  '*** (Jen 2.14 Next)
  FraCompUpd.Visible = False
  FraCompUpd.Enabled = False
  '*** (Jen 2.14 Next)
  
  '*** (Jen 3.01)
  FraUpdRmQw.Visible = False
  FraUpdRmQw.Enabled = False
  '*** (Jen 3.01)
  '*** Jay 3.1.0_1
  FraCmpCodeIns.Visible = False
  FraCmpCodeIns.Enabled = False
  '*** Jay 3.1.0_1
End Sub

'*** Std Code not to be changed
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
  ws_BtnStr = "CmdUpdPrdStgGo,CmdUpdCoCdGo,CmdDelOrdGo,CmdPurPltGo,CmdAvgStkRtGo,CmdUpdSYyMmGo,CmdUpdRmZCdGo,CmdUpdUAGo,CmdDelInvGo,CmdUpdCoName"
  Call SetObjColors(Me, ws_BtnStr)
  '*** Jenny Colour
  ' Zubin 212
  'FraMtnPwd.Move (Me.Width - FraMtnPwd.Width) \ 2, _
                  (Me.Height - FraMtnPwd.Height) \ 2, FraMtnPwd.Width, FraMtnPwd.Height
  ' Zubin 212
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
Private Sub GrdDsgAnlyFlds_Validate(Cancel As Boolean)
  Cancel = GrdDsgAnlyFlds.Validate
End Sub

'*** (Jen 2.14 Next)
Private Sub GrdCompUpd_Validate(Cancel As Boolean)
  Cancel = GrdCompUpd.Validate
End Sub
'*** (Jen 2.14 Next)

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    If UnloadMode <> vbFormCode Then Cancel = True: Exit Sub
    
    Set moCn = Nothing    '*** (09/08/05)    '*** (Bef speed) bef 24/11/06
End Sub
'*** Jay 3.01 *** [Inv-Fg Rm Mismatch]
Private Sub InvFgRmMisMatch()
  '****** The Procedures should be on the path D:\MWAPP\MWSQL ******
  '*** This procedure is used for Correction of Invoice & Fg Rm MisMatch
  '*** Procedure Include JewelMark Script & Update SYyMm Script

  Dim ws_str As String, ws_StartTime As String
  Dim ws_OrdCnd As String, ws_pth As String
  On Error GoTo ErrH
  TxtDisp.text = ""
  DispMsg "", etInfo

  ws_StartTime = "Start Time: " + CStr(Format(moCn.SrvrDate, IIF(UCase(MWLib.DtSetting) = "MM/DD/YY", "MMM DD YYYY", "DD MMM YYYY"))) + " / " + Format(moCn.SrvrTime, "#0.#0")
  DispMsg "Start Time: " + CStr(Format(moCn.SrvrDate, IIF(UCase(MWLib.DtSetting) = "MM/DD/YY", "MMM DD YYYY", "DD MMM YYYY"))) + " / " + Format(moCn.SrvrTime, "#0.#0"), etInfo

  '****** Checking whether the Required Files Exists
  ws_pth = "D:\MWAPP\MWSQL\Repair Invoice Fg Rm MisMatch.sql"
  If Dir(ws_pth, vbNormal) = "" Then
      TxtDisp.text = "File Not Found : '" + ws_pth + "', Cannot Continue...'"
      DispMsg "File Not Found : '" + ws_pth + "'", etError
      Exit Sub
  End If
  
  'JewelMark Script
  moCn.Execute ("If Exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[UpdFGStkMismatch]') and OBJECTPROPERTY(id, N'IsProcedure') = 1) " + _
            "Drop Procedure [Dbo].[UpdFGStkMismatch]")
  moCn.Execute ("If exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[TmpInvFgd]') and OBJECTPROPERTY(id, N'IsTable') = 1) " + _
            "Drop Table [dbo].[TmpInvFgd]")
  
  ws_str = ReadFileContents("D:\MWAPP\MWSQL\Repair Invoice Fg Rm MisMatch.sql")
  moCn.Execute (ws_str)

  DoEvents
  MWLib.BeginProcess Me, "Mismatch Error Is Under Process, "
  
  ws_str = "[dbo].[UpdFGStkMismatch]"
  moCn.Execute (ws_str)

  'SYyMm Update
  Call UpdateSYyMm
  
  'Display Deleted Records From InvFgd
  'select * From TmpInvFgd
  Dim Rs_TmpInvFgd As MwfLib.MDORowSet
  Dim wi_Row  As Integer
  Set Rs_TmpInvFgd = moCn.OpenRes("select IfCocd+ '/'+ IfTc+ '/'+ IfYy+ '/'+ IfChr+ '/'+ Ltrim(Rtrim(Str(IfNo)))+ " + _
          " '/'+ Ltrim(Rtrim(Str(IfSr))) + '/'+ Ltrim(Rtrim(Str(IfSrNo))) as qIfNo,  " + _
          "IfFdTc + '/' + IfFdYy + '/' + IfFdChr + '/' + Ltrim(Rtrim(Str(IfFdNo))) + '/' + " + _
          "Ltrim(Rtrim(Str(IfFdSr))) as qFdNo, IfFdBYy + '/' + IfFdBChr + '/' + Ltrim(Rtrim(Str(IfFdBNo))) as qBNo " + _
          "From TmpInvFgd ")
    GrdInvFgdDelRecs.Rows = 1
    If Rs_TmpInvFgd.RecCount > 0 Then
      GrdInvFgdDelRecs.Visible = True
      LblTmpInvFgd.Visible = True
      With GrdInvFgdDelRecs
        .AllowAdd = True
        Do While Not (Rs_TmpInvFgd.EOF Or Rs_TmpInvFgd.BOF)
          .AddItem
          wi_Row = .Rows - 1
          .Value(wi_Row, "qIfNo") = Rs_TmpInvFgd!qIfNo
          .Value(wi_Row, "qFdNo") = Rs_TmpInvFgd!qfdNo
          .Value(wi_Row, "qBNo") = Rs_TmpInvFgd!qBNo
          Rs_TmpInvFgd.MoveNext
        Loop
        .AllowAdd = False
      End With
      TxtDisp.text = TxtDisp.text + Chr(13) + "Data in The Grid is Deleted from Invoice during Repairing process."
    Else
      GrdInvFgdDelRecs.Visible = False
      LblTmpInvFgd.Visible = False
    End If
  MWLib.EndProcess Me
  TxtDisp.text = ws_StartTime + Chr(13) + " " + TxtDisp.text + Chr(13) + _
                " Invoice & Fg Rm Mismatch Error Has Been Rectified From Database " + Chr(13) + _
                " End Time: " + CStr(Format(moCn.SrvrDate, IIF(UCase(MWLib.DtSetting) = "MM/DD/YY", "MMM DD YYYY", "DD MMM YYYY"))) + " / " + Format(moCn.SrvrTime, "#0.#0")
  DispMsg "Invoice & Fg Rm Mismatch Error Has Been Rectified From Database", etInfo
  Exit Sub
ErrH:
  Call ErrHndlr
  MWLib.EndProcess Me
  ws_str = "Delete TrgDisa"
  moCn.Execute (ws_str)
End Sub
'*** Jay 3.1.0_1
Private Sub CmdInsComp_Click()
  Call DispFra(InsertComp)
End Sub
'*** Jay 3.1.0_1
Private Sub CmdNewCmpGo_Click()
On Error GoTo ErrHandler
  Dim SqlStr As String
    
  ' ***** Manali 3.8.0 - Supp Pwd checked for Adding New Company
  If TxtInsCompPwd <> mcs_Pwd Then DispMsg "Invalid Support Password", etError: Exit Sub

  If moCn.RecSeek("Select 'x' from Head where HCd= '" + ctSelfCmCd + "' and HCoCd= '" + Trim(TxtCmpCode) + "'") Then
    TxtDisp.text = "Company Code Already Exists!!"
    DispMsg "Company Code Already Exists!!", etInfo: Exit Sub
    Exit Sub
  End If
  
  If moCn.RecSeek("Select 'x' from Head where HCd= '" + ctSelfCmCd + "' and HName= '" + Trim(TxtCmpName) + "'") Then
    TxtDisp.text = "Company Name Already Exists!!"
    DispMsg "Company Name Already Exists!!", etInfo: Exit Sub
    Exit Sub
  End If
  
  
  If adc("wSrcCoCd") = "" Or Trim(TxtCmpCode.text) = "" Or Trim(TxtCmpName.text) = "" Then
    TxtDisp.text = "Source Company Code, New Company Code And Company Name should be entered for this Update"
    DispMsg "Both Company Code And Company Name should be entered for this Update", etError: Exit Sub
  End If

  ' Add record in Head Master
  '******* Manali 3.2.2 - hChgPtrOnStwYN, hCmILabWtFrOrd and Imp Before date fields for auto custom matching added
  '****** Sachin 3.5.0 - hCustomizeRmCtgExp added
  SqlStr = " Begin Transaction "
  SqlStr = SqlStr + " Insert Into TrgDisa Values ('Head') "
  SqlStr = SqlStr + " Insert Into Head " + _
      "(HAgAdd1,HAgAdd2,HAgAdd3,HAgName,HAR4No,hBChr,HBinNo,HBkAc,HBkAdd1,HBkAdd2,HBkAdd3,HBkName,HBnkInstr,hBOpnLoc,HCd,HCmCdHlp,hCmCsRtFA,hCmDiaRtFA,hCmFixPrc,hCmGldAs,hCmIGldRtEqOrd,hCmIWtEqOrd,hCmIWtFrOrd,HCoCd,HConOfOrg,HConsignee,hDmPrdSeq,hDmVaCtg,HEtcSection,HETCSvg,HEtcZone,hEximPolicy,hExpLocal,HExpOff,hFgPlChr,HFgSubLoc,HFrgnBkAc,HFrgnBkName,HFtRem,hGrsWtInBsYN,hGspEligibility,HGSPOrgCrit,HIecNo,hInRepCnvRt,hInRepMetLs,hInRepRtFld,hInRepWtFld,hInRepWtValAs,HInsBy,HInsNo,HItcNo,HLicNo,hLine,HMarks,hMaxLineUtil,hMaxOrdQty,hMinOrdPrcnt,hMinOrdQty,Hname,HOdAmdQty,hOmPrtCd,HOrdCpyRtOpt,hOrdDsgGr,hOrdLabGr,hOrdRmGr,hOthLocBag,hPdcSfx,HPkgDesc,HPortofLd,hPpcYN,HPrdDesc,HPreCarriage,HPreCarrier,HPtrYN,HQCertNo,HRbiNo,HRndOffFOB,HRndOrdSalPrc,HShowActWt,HShowPrdRmDets,HSrtBasePurVch,hStatCd,hTdAck,HTotAs,HUnResolutions,HVaDc,HValAddnFt,HVaX,Hvessel,hWH,hYy,hZValEqInv,hZWtEq,ModDt,ModTime,ModUsr" + _
",HAlyCnvDiff,HAuthSgn,HAutoCustmMtch,HAutomateReturns,HAvgSRtYn,HBagNoEditYN,HBaseCurCd,HChkDCQtyAtIssYN,HChkDCWtAtIssYN,HChkGPWtAtIssYN,hCmLabAs,HCnvDiffEffDt,hCompBagsYn,HConsolidatedInv,HCstNo,HCstPer,HCtbMod,HCtCd,HCustomiseOrd,HCustRem,HDTPrdPtr,hETCTxt,hFgQuotYN,hFgRmEqOrdRm,hFluteBagYN,hGenExpNoYN,hGjepcNo,hGjepcValidUpTo,hGldFiness,HGldRtDAYN,HGspPkgDesc,HImpAftC,HImpAftD,HImpAftG,HImpAftP,HImpAftS,HImpAftX,hImpComLoc,HImpTypNotInSelC,HImpTypNotInSelD,HImpTypNotInSelG,HImpTypNotInSelP,HImpTypNotInSelS,HImpTypNotInSelX,hImpValTxt,hInRepMetLsRmCd,HInvTyp,HLclInvFT,HLclInvVATFT,HLclInvYN,HLooseInv,HLotNoYN,hMetalBOETxt,hMMTc,hModAvlYN,HMultiCurYN,hNegStkRmCtg,HNegStkYn,HOdGmChk,HOrdPrdPts,hPlFiness,HPrdPtsYN,HPrtAllotOnPrdYN,HPtrMinusC,HPtrMinusD,HPtrMinusG,HPtrMinusP,HPtrMinusS,HPtrMinusX,HPtrPlusC,HPtrPlusD,HPtrPlusG,HPtrPlusP,HPtrPlusS,HPtrPlusX,hQWModYN,hRbiPolicy,HRtMinusC,HRtMinusD,HRtMinusG,HRtMinusP" + _
",HRtMinusS,HRtMinusX,HRtPlusC,HRtPlusD,HRtPlusG,HRtPlusP,HRtPlusS,HRtPlusX,hShpBillSgnBy,HSizeStkRt,hSlFiness,hSlvModYN,hStatOfOrg,HStkRtPassReq,HValAddMtlColYN,HValBOEDetYN,HVATNo,HVatPer" + _
",hBestUse,HCHALicenseCd,hChgPtrOnStwYN,hChkCastLocs,hChkIntgty,HChkOrRtYN,hChkVldRmLnYN,hChkWaxPullOutLocs,hChnSubCtgForGld,hChnSubCtgForLab,hCmILabWtFrOrd,hCpyActCalcWtYN,HCustomHouseCd" + _
",hCustomizeRmCtgExp,HDsgCatYN,HEnfFrBLocInFgYN,hExcelImpInOrdYN,hFgIOYn,HFgTagMod,HHideCommYN,HHolChkYN,hHSCode,HImpAftL,HImpBefC,HImpBefD,HImpBefG,HImpBefL,HImpBefP,HImpBefS,HImpBefX" + _
",HImpTypNotInSelL,HInAmt1Per,HInAmt2Per,HInAmt3Per,HInAmt4Per,HInAmt5Per,HInAmt6Per,HInAmtDesc1,HInAmtDesc10,HInAmtDesc11,HInAmtDesc12,HInAmtDesc2,HInAmtDesc3,HInAmtDesc4,HInAmtDesc5" + _
",HInAmtDesc6,HInAmtDesc7,HInAmtDesc8,HInAmtDesc9,hInCircularNo,hInGspAdd1,hInGspAdd2,hInGspAdd3,HInInvGrsWtEqActual,HInInvTermCond,hInIRtwrtStkYN,HInLclInvTyp,HInMemoTermCond,hInNotionRtYN" + _
",hInRegNo,HInSrt,hInXIRtwrtStkYN,hjStkDefChr,HjStkNumjFrBagYN,HLclInvCSTFT,HMaxDsgViewRec,hMetIRtwrtStkYN,HMetlrCd,HMinDiaProcess,HMultiPrcQtMod,HOrdExcelRepYn,HOrdRtRndOffRmCtg,HPAN,HPdFiness" + _
",HPdModYN,HPoL,hPrdPtsOnBQtyYN,hPrtKey,HPtrMinusL,HPtrPlusL,HRfIdMod,HRndOffOdTolWtYn,hRtfrRtChrtYN,HRtMinusL,HRtPlusL,hRwrkModYn,HSbUom,HSelfDeductYN,HSelfDscYN,hSetTypOnlyFrOrdYN,hShowBgHistDetsYN " + _
",HShowRepLogo , HTinNo, hTolQtyPctDC, hTolWtPctDC, HTotNoOfPack, hTravelExlYN, HTypOfExporter, HUnitCd, HValidDsgSetCd" + _
",hCell,HCERegnNo,HCHCDHCOnCustRm,hClstrCd,HCstRtEdtYn,HDefRepBagLoc,hDmPrcsSeq,HInAmt10Per,HInAmt11Per,HInAmt8Per,HInAmt9Per,hLocSeq,HRmStkBy,MqrLn3,hStkRtNoDsp,hStkRtNoEdit)"

SqlStr = SqlStr + "Select HAgAdd1,HAgAdd2,HAgAdd3,HAgName,HAR4No,hBChr,HBinNo,HBkAc,HBkAdd1,HBkAdd2,HBkAdd3,HBkName,HBnkInstr,hBOpnLoc,HCd,HCmCdHlp,hCmCsRtFA,hCmDiaRtFA,hCmFixPrc,hCmGldAs,hCmIGldRtEqOrd,hCmIWtEqOrd,hCmIWtFrOrd,'" + Trim(TxtCmpCode) + "', HConOfOrg,HConsignee,hDmPrdSeq,hDmVaCtg,HEtcSection,HETCSvg,HEtcZone,hEximPolicy,hExpLocal,HExpOff,hFgPlChr,HFgSubLoc,HFrgnBkAc,HFrgnBkName,HFtRem,hGrsWtInBsYN,hGspEligibility,HGSPOrgCrit,HIecNo,hInRepCnvRt,hInRepMetLs,hInRepRtFld,hInRepWtFld,hInRepWtValAs,HInsBy,HInsNo,HItcNo,HLicNo,hLine,HMarks,hMaxLineUtil,hMaxOrdQty,hMinOrdPrcnt,hMinOrdQty, '" + Trim(TxtCmpName) + "',HOdAmdQty,hOmPrtCd,HOrdCpyRtOpt,hOrdDsgGr,hOrdLabGr,hOrdRmGr,hOthLocBag,hPdcSfx,HPkgDesc,HPortofLd,hPpcYN,HPrdDesc,HPreCarriage,HPreCarrier,HPtrYN,HQCertNo,HRbiNo,HRndOffFOB,HRndOrdSalPrc,HShowActWt,HShowPrdRmDets,HSrtBasePurVch,hStatCd,hTdAck,HTotAs,HUnResolutions,HVaDc,HValAddnFt,HVaX,Hvessel,hWH,hYy,hZValEqInv,hZWtEq,ModDt,ModTime,ModUsr" + _
",HAlyCnvDiff,HAuthSgn,HAutoCustmMtch,HAutomateReturns,HAvgSRtYn,HBagNoEditYN,HBaseCurCd,HChkDCQtyAtIssYN,HChkDCWtAtIssYN,HChkGPWtAtIssYN,hCmLabAs,HCnvDiffEffDt,hCompBagsYn,HConsolidatedInv,HCstNo,HCstPer,HCtbMod,HCtCd,HCustomiseOrd,HCustRem,HDTPrdPtr,hETCTxt,hFgQuotYN,hFgRmEqOrdRm,hFluteBagYN,hGenExpNoYN,hGjepcNo,hGjepcValidUpTo,hGldFiness,HGldRtDAYN,HGspPkgDesc,HImpAftC,HImpAftD,HImpAftG,HImpAftP,HImpAftS,HImpAftX,hImpComLoc,HImpTypNotInSelC,HImpTypNotInSelD,HImpTypNotInSelG,HImpTypNotInSelP,HImpTypNotInSelS,HImpTypNotInSelX,hImpValTxt,hInRepMetLsRmCd,HInvTyp,HLclInvFT,HLclInvVATFT,HLclInvYN,HLooseInv,HLotNoYN,hMetalBOETxt,hMMTc,hModAvlYN,HMultiCurYN,hNegStkRmCtg,HNegStkYn,HOdGmChk,HOrdPrdPts,hPlFiness,HPrdPtsYN,HPrtAllotOnPrdYN,HPtrMinusC,HPtrMinusD,HPtrMinusG,HPtrMinusP,HPtrMinusS,HPtrMinusX,HPtrPlusC,HPtrPlusD,HPtrPlusG,HPtrPlusP,HPtrPlusS,HPtrPlusX,hQWModYN,hRbiPolicy,HRtMinusC,HRtMinusD,HRtMinusG,HRtMinusP" + _
",HRtMinusS,HRtMinusX,HRtPlusC,HRtPlusD,HRtPlusG,HRtPlusP,HRtPlusS,HRtPlusX,hShpBillSgnBy,HSizeStkRt,hSlFiness,hSlvModYN,hStatOfOrg,HStkRtPassReq,HValAddMtlColYN,HValBOEDetYN,HVATNo,HVatPer" + _
",hBestUse,HCHALicenseCd,hChgPtrOnStwYN,hChkCastLocs,'',HChkOrRtYN,hChkVldRmLnYN,hChkWaxPullOutLocs,hChnSubCtgForGld,hChnSubCtgForLab,hCmILabWtFrOrd,hCpyActCalcWtYN,HCustomHouseCd" + _
",hCustomizeRmCtgExp,HDsgCatYN,HEnfFrBLocInFgYN,hExcelImpInOrdYN,hFgIOYn,HFgTagMod,HHideCommYN,HHolChkYN,hHSCode,HImpAftL,HImpBefC,HImpBefD,HImpBefG,HImpBefL,HImpBefP,HImpBefS,HImpBefX" + _
",HImpTypNotInSelL,HInAmt1Per,HInAmt2Per,HInAmt3Per,HInAmt4Per,HInAmt5Per,HInAmt6Per,HInAmtDesc1,HInAmtDesc10,HInAmtDesc11,HInAmtDesc12,HInAmtDesc2,HInAmtDesc3,HInAmtDesc4,HInAmtDesc5" + _
",HInAmtDesc6,HInAmtDesc7,HInAmtDesc8,HInAmtDesc9,hInCircularNo,hInGspAdd1,hInGspAdd2,hInGspAdd3,HInInvGrsWtEqActual,HInInvTermCond,hInIRtwrtStkYN,HInLclInvTyp,HInMemoTermCond,hInNotionRtYN" + _
",hInRegNo,HInSrt,hInXIRtwrtStkYN,hjStkDefChr,HjStkNumjFrBagYN,HLclInvCSTFT,HMaxDsgViewRec,hMetIRtwrtStkYN,HMetlrCd,HMinDiaProcess,HMultiPrcQtMod,HOrdExcelRepYn,HOrdRtRndOffRmCtg,HPAN,HPdFiness" + _
",HPdModYN,HPoL,hPrdPtsOnBQtyYN,hPrtKey,HPtrMinusL,HPtrPlusL,HRfIdMod,HRndOffOdTolWtYn,hRtfrRtChrtYN,HRtMinusL,HRtPlusL,hRwrkModYn,HSbUom,HSelfDeductYN,HSelfDscYN,hSetTypOnlyFrOrdYN,hShowBgHistDetsYN " + _
",HShowRepLogo , HTinNo, hTolQtyPctDC, hTolWtPctDC, HTotNoOfPack, hTravelExlYN, HTypOfExporter, HUnitCd, HValidDsgSetCd" + _
",hCell,HCERegnNo,HCHCDHCOnCustRm,hClstrCd,HCstRtEdtYn,HDefRepBagLoc,hDmPrcsSeq,HInAmt10Per,HInAmt11Per,HInAmt8Per,HInAmt9Per,hLocSeq,HRmStkBy,MqrLn3,hStkRtNoDsp,hStkRtNoEdit" + _
" From Head where HCd= '" + ctSelfCmCd + "' and HCoCd= '" + adc("wSrcCoCd") + "' "
  
     '****** Sachin 3.3.0 - Added new fields here
   '*** Jay 3.3(FG)
   
  ' Add record in Company Master
  ' ***** Manali 3.2.2 - CmILabWtFrOrd added
  ' ***** Manali 3.9.0 - CmTinNo added
  ' **** Manali 3.10.0 - 03/04/12 - CmCstNo, CmVATNo, CmPAN added
  SqlStr = SqlStr + " Insert Into TrgDisa Values ('CustMst') "
  SqlStr = SqlStr + " Insert Into CustMst " + _
            " Select 'Z', '" + Trim(TxtCmpCode) + "', '" + Trim(TxtCmpName) + "', CmMulBy, CmFixPrc, " + _
            " CmGldAs, CmIWtEqOrd, CmIWtFrOrd, CmIGldRtEqOrd, CmDiaRtFA, CmCSRtFA, CmPrdInst, CmStmpInst, " + _
            " CmPayTerms, CmDelTerms, CmAdd1, CmAdd2, CmAdd3, CmTel, CmFax, CmEMail, CmCpName, CmPortofDisch, " + _
            " CmFinDest, CmBuyer, ModUsr, ModDt, ModTime, CmLkUpRmRt, CmLkUpLabRt, CmLkUpMetLs, CmRegnCd, CmHInsBy, " + _
            " CmHRndOrdSalPrc, CmHistory, CmCurCd, CmSalPer, CmBuyAdd1, CmBuyAdd2, CmBuyAdd3, CmBkCd, CmAgCd, " + _
            " CmValidYN, CmVessel, CmLabAs, CmBillCo, CmVatYn, CmCstYn, CmOctroiYn, CmOctroiPer, CmOurBkCd, CmPrtKey, CmILabWtFrOrd, CmTinNo, " + _
            " CmCstNo, CmVATNo, CmPAN,CmGrp,CmTyp,CmPod,CmFinDstCd,CmInvShowGrsWtYN,CmInvShowDiaWtYN,CmInvFtRem,CmMrpMulby,CmSeekZSelfPDCM,CmSeekZSelfRt, " + _
            " CmOmSplRem,CmInvAllBags,CmValidCoCd,CmCell            " + _
            "from CustMst where CmCtg= 'Z' and CmCd= '" + adc("wSrcCoCd") + "' "

  ' Add records in vParam
  SqlStr = SqlStr + "Insert Into TrgDisa Values ('vParam') "
  SqlStr = SqlStr + " Insert Into vParam " + _
            " Select '" + Trim(TxtCmpCode) + "', vPTyp, vPMCd, vPSCd, vPDesc, vPDesc225, vPValue, " + _
            " vPNum, vPValue1, vPNum1, vPValue2, vPValue3, ModUsr, ModDt, ModTime, vPValidYn, vPValue4, vPPrtKey " + _
            " from vParam where vPCoCd= '" + adc("wSrcCoCd") + "' "

  ' Add records in Loc
  ' **** Manali 3.10.0 - 03/04/12 - LocAllowLsPer added
  SqlStr = SqlStr + "Insert Into TrgDisa Values ('Loc') "
  SqlStr = SqlStr + " Insert Into Loc " + _
            " Select LocCd, LocDesc, LocTyp, LocPrdTyp, LocLoss, LocBrk, LocMsg, " + _
            " ModUsr, ModDt, ModTime, '" + Trim(TxtCmpCode) + "', LocPrdStg, LocValidYN, " + _
            " LocLsTyp, LocPrntYN, LocPrntCd, LocPrtKey, LocAllowLsPer,LocHideYN, " + _
            " LocPrcs , LocUseCapOf " + _
            " from Loc where LocCoCd= '" + adc("wSrcCoCd") + "' "

  '*** Jay 3.01 *** [Tag Entry Lock for DFDMX, DFZEB] (Middle of 3.01)
  SqlStr = SqlStr + " Insert Into TrgDisa Values ('TagFmt') Insert Into TagFmt " + _
            " Select '" + Trim(TxtCmpCode) + "', TgfCd, TgfDesc, TgfStartStr, TgfDarkness, TgfHeat, TgfQtyStr, " + _
            " TgfEndStr, TgfFldOrg, TgfFldData, TgfFldEnd, ModUsr, ModDt, ModTime " + _
            " from TagFmt where TgfCoCd= '" + adc("wSrcCoCd") + "' and TgfCd in ('DFDMX', 'DFZEB') "
              
  SqlStr = SqlStr + " Insert Into TrgDisa Values ('TagFmtSr')  Insert Into TagFmtSr " + _
            " Select '" + Trim(TxtCmpCode) + "', TgfsCd, TgfsSr, TgfsPrintStr, TgfsField, ModUsr, ModDt, ModTime " + _
            " from TagFmtSr where TgfsCoCd= '" + adc("wSrcCoCd") + "' and TgfsCd in ('DFDMX', 'DFZEB') "
  '*** Jay 3.01 *** [Tag Entry Lock for DFDMX, DFZEB] (Middle of 3.01)
  
  '*** JAY 3.2.0 [AutoTree]
  ' ****** Manali 3.03 -  Not Required
'  SqlStr = SqlStr + " Insert Into TrgDisa Values ('vParam')  Insert into vparam " + _
'            "Select '" + Trim(TxtCmpCode) + "', vPTyp, vPMCd, vPSCd, vPDesc, vPDesc225, vPValue, vPNum, vPValue1, " + _
'            "vPNum1, vPValue2, vPValue3, vParam.ModUsr, vParam.ModDt, vParam.ModTime, vPValidYn, vPValue4, vPPrtKey " + _
'            "From vParam where vPCoCd= '" + ADC("wSrcCoCd") + "' and vPTyp = 'TrDef'"
  '*** JAY 3.2.0 [AutoTree]
  
  ' ***** Manali Trading Module - TcMst
  SqlStr = SqlStr + " Insert Into TrgDisa Values ('TcMst')  Insert Into TcMst " + _
            " Select '" + Trim(TxtCmpCode) + "', TmTc, TmDesc, TmBaseTcs, TmRefTcs, TmSrcTcs, TmSrcCmCtg, " + _
            " TmLocTypFr, TmFrDmDc, TmLocTypTo, TmToDmDc, TmCmCtg, TmBaseMustYN, ModUsr, ModDt, ModTime,TmInvHdMustYn, TmChkDuplicateBagYN,TmDmStkNumYN  " + _
            " ,TmJLooseTC   from TcMst where TmCoCd= '" + adc("wSrcCoCd") + "'  "
  
  SqlStr = SqlStr + " Delete from Trgdisa "
  SqlStr = SqlStr + " Commit Transaction "
  moCn.Execute (SqlStr)
  'Debug.Print SqlStr
  
  If moCn.RecSeek("Select 'x' from Head where HCd= '" + ctSelfCmCd + "' and HCoCd= '" + Trim(TxtCmpCode) + "'") Then
    Call UpdIntegrity(moCn, Trim(TxtCmpCode))    '****** Sachin 3.3.0
    TxtDisp.text = "Company Code Added Successfully!!"
    DispMsg "Company Code Added Successfully!!", etInfo
  Else
    TxtDisp.text = "Company Not Added !!"
    DispMsg "Company Not Added !!", etInfo
  End If
  
  CmdInsComp_Click
  Exit Sub
ErrHandler:
    moCn.RollbackTrans mdoCommitNone
    TxtDisp.text = Err.Description
    DispMsg Err.Description, etWarning
End Sub

Private Sub TxtCmpCode_KeyPress(KeyAscii As Integer)
  KeyAscii = Asc(UCase$(Chr$(KeyAscii)))
End Sub



