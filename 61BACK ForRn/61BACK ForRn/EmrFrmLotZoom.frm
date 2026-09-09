VERSION 5.00
Object = "{AB2CC32A-E652-42A2-B341-26ECE02F6386}#2.1#0"; "MwfCtl.ocx"
Begin VB.Form EmrFrmLotZoom 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Lot Zoom"
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
   Begin VB.Frame FRA_KEYS 
      BorderStyle     =   0  'None
      Height          =   1935
      Left            =   0
      TabIndex        =   33
      Top             =   0
      Width           =   14490
      Begin VB.Frame Frame2 
         Height          =   795
         Left            =   120
         TabIndex        =   77
         Top             =   1150
         Width           =   13245
         Begin VB.CheckBox ChkBoxArr 
            Alignment       =   1  'Right Justify
            Caption         =   "Orphans"
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
            Left            =   11520
            TabIndex        =   88
            Tag             =   "ShowVal"
            ToolTipText     =   "Check To Show Orphans"
            Top             =   120
            Width           =   1245
         End
         Begin MwfCtl.MWCTL_MED atxt 
            Height          =   285
            Index           =   0
            Left            =   1080
            TabIndex        =   17
            ToolTipText     =   "Enter From Transaction Tc"
            Top             =   120
            Width           =   675
            _ExtentX        =   1191
            _ExtentY        =   503
            MaxLength       =   3
            DataType        =   4
            IdName          =   "WTDTCFR"
         End
         Begin MwfCtl.MWCTL_MED atxt 
            Height          =   285
            Index           =   1
            Left            =   1860
            TabIndex        =   19
            ToolTipText     =   "Enter From Transaction Year"
            Top             =   120
            Width           =   345
            _ExtentX        =   609
            _ExtentY        =   503
            MaxLength       =   2
            DataType        =   4
            IdName          =   "WTDYYFR"
         End
         Begin MwfCtl.MWCTL_MED atxt 
            Height          =   285
            Index           =   2
            Left            =   2310
            TabIndex        =   21
            ToolTipText     =   "Enter From Transaction Chr"
            Top             =   120
            Width           =   675
            _ExtentX        =   1191
            _ExtentY        =   503
            MaxLength       =   3
            DataType        =   4
            IdName          =   "WTDCHRFR"
         End
         Begin MwfCtl.MWCTL_MED atxt 
            Height          =   285
            Index           =   3
            Left            =   4140
            TabIndex        =   25
            ToolTipText     =   "Enter From Transaction Sr"
            Top             =   120
            Width           =   555
            _ExtentX        =   979
            _ExtentY        =   503
            Alignment       =   2
            Mask            =   "###0"
            MaxLength       =   4
            DataType        =   1
            IdName          =   "WTDSRFR"
         End
         Begin MwfCtl.MWCTL_MED atxt 
            Height          =   285
            Index           =   4
            Left            =   3090
            TabIndex        =   23
            ToolTipText     =   "Enter From Transaction No"
            Top             =   120
            Width           =   980
            _ExtentX        =   1720
            _ExtentY        =   503
            Alignment       =   2
            Mask            =   "######0"
            MaxLength       =   7
            DataType        =   1
            IdName          =   "WTDNOFR"
         End
         Begin MwfCtl.MWCTL_MED atxt 
            Height          =   285
            Index           =   5
            Left            =   1080
            TabIndex        =   18
            ToolTipText     =   "Enter To Transaction Tc"
            Top             =   400
            Width           =   675
            _ExtentX        =   1191
            _ExtentY        =   503
            MaxLength       =   3
            DataType        =   4
            IdName          =   "WTDTCTO"
         End
         Begin MwfCtl.MWCTL_MED atxt 
            Height          =   285
            Index           =   7
            Left            =   2310
            TabIndex        =   22
            ToolTipText     =   "Enter To Transaction Chr"
            Top             =   400
            Width           =   675
            _ExtentX        =   1191
            _ExtentY        =   503
            MaxLength       =   3
            DataType        =   4
            IdName          =   "WTDCHRTO"
         End
         Begin MwfCtl.MWCTL_MED atxt 
            Height          =   285
            Index           =   8
            Left            =   4140
            TabIndex        =   26
            ToolTipText     =   "Enter To Transaction Sr"
            Top             =   400
            Width           =   555
            _ExtentX        =   979
            _ExtentY        =   503
            Alignment       =   2
            Mask            =   "###0"
            MaxLength       =   4
            DataType        =   1
            IdName          =   "WTDSRTO"
         End
         Begin MwfCtl.MWCTL_MED atxt 
            Height          =   285
            Index           =   9
            Left            =   3090
            TabIndex        =   24
            ToolTipText     =   "Enter To Transaction No"
            Top             =   400
            Width           =   980
            _ExtentX        =   1720
            _ExtentY        =   503
            Alignment       =   2
            Mask            =   "######0"
            MaxLength       =   7
            DataType        =   1
            IdName          =   "WTDNOTO"
         End
         Begin MwfCtl.MWCTL_MED atxt 
            Height          =   285
            Index           =   17
            Left            =   7920
            TabIndex        =   27
            ToolTipText     =   "Enter From Rm Category"
            Top             =   120
            Width           =   555
            _ExtentX        =   979
            _ExtentY        =   503
            MaxLength       =   2
            DataType        =   4
            IdName          =   "WTDRMCTGFR"
         End
         Begin MwfCtl.MWCTL_MED atxt 
            Height          =   285
            Index           =   18
            Left            =   7920
            TabIndex        =   28
            ToolTipText     =   "Enter To Rm Category"
            Top             =   405
            Width           =   555
            _ExtentX        =   979
            _ExtentY        =   503
            MaxLength       =   2
            DataType        =   4
            IdName          =   "WTDRMCTGTO"
         End
         Begin MwfCtl.MWCTL_MED atxt 
            Height          =   285
            Index           =   19
            Left            =   8760
            TabIndex        =   29
            ToolTipText     =   "Enter From Currency"
            Top             =   120
            Width           =   915
            _ExtentX        =   1614
            _ExtentY        =   503
            MaxLength       =   5
            DataType        =   4
            IdName          =   "WTDCURRFR"
         End
         Begin MwfCtl.MWCTL_MED atxt 
            Height          =   285
            Index           =   20
            Left            =   8760
            TabIndex        =   30
            ToolTipText     =   "Enter To Currency"
            Top             =   405
            Width           =   915
            _ExtentX        =   1614
            _ExtentY        =   503
            MaxLength       =   5
            DataType        =   4
            IdName          =   "WTDCURRTO"
         End
         Begin MwfCtl.MWCTL_MED atxt 
            Height          =   285
            Index           =   6
            Left            =   1860
            TabIndex        =   20
            ToolTipText     =   "Enter To Transaction Year"
            Top             =   400
            Width           =   345
            _ExtentX        =   609
            _ExtentY        =   503
            MaxLength       =   2
            DataType        =   4
            IdName          =   "WTDYYTO"
         End
         Begin MwfCtl.MWCTL_MED atxt 
            Height          =   285
            Index           =   42
            Left            =   9960
            TabIndex        =   31
            ToolTipText     =   "Enter Transaction Sort By Option"
            Top             =   120
            Width           =   1335
            _ExtentX        =   2355
            _ExtentY        =   503
            MaxLength       =   8
            DataType        =   4
            IdName          =   "WTDSORT"
         End
         Begin MwfCtl.MWCTL_MED atxt 
            Height          =   285
            Index           =   16
            Left            =   12840
            TabIndex        =   89
            Top             =   120
            Visible         =   0   'False
            Width           =   345
            _ExtentX        =   609
            _ExtentY        =   503
            Enabled         =   0   'False
            DataType        =   4
            IdName          =   "WORPHYN"
         End
         Begin VB.Label ALBL 
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
            Height          =   285
            Index           =   3
            Left            =   120
            TabIndex        =   87
            Top             =   120
            Width           =   825
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
            Index           =   2
            Left            =   120
            TabIndex        =   86
            Top             =   480
            Width           =   825
         End
         Begin VB.Label Label12 
            AutoSize        =   -1  'True
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
            Height          =   240
            Left            =   4065
            TabIndex        =   85
            Top             =   120
            Width           =   60
         End
         Begin VB.Label Label8 
            AutoSize        =   -1  'True
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
            Height          =   240
            Index           =   6
            Left            =   3030
            TabIndex        =   84
            Top             =   120
            Width           =   60
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
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
            Height          =   240
            Index           =   2
            Left            =   2250
            TabIndex        =   83
            Top             =   120
            Width           =   60
         End
         Begin VB.Label Label11 
            AutoSize        =   -1  'True
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
            Height          =   240
            Left            =   1800
            TabIndex        =   82
            Top             =   120
            Width           =   60
         End
         Begin VB.Label Label8 
            AutoSize        =   -1  'True
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
            Height          =   240
            Index           =   5
            Left            =   4065
            TabIndex        =   81
            Top             =   480
            Width           =   60
         End
         Begin VB.Label Label8 
            AutoSize        =   -1  'True
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
            Height          =   240
            Index           =   4
            Left            =   3030
            TabIndex        =   80
            Top             =   480
            Width           =   60
         End
         Begin VB.Label Label10 
            AutoSize        =   -1  'True
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
            Height          =   240
            Left            =   1800
            TabIndex        =   79
            Top             =   480
            Width           =   60
         End
         Begin VB.Label Label9 
            AutoSize        =   -1  'True
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
            Height          =   240
            Left            =   2250
            TabIndex        =   78
            Top             =   480
            Width           =   60
         End
      End
      Begin VB.Frame Frame1 
         Height          =   795
         Left            =   120
         TabIndex        =   65
         Top             =   360
         Width           =   13245
         Begin MwfCtl.MWCTL_MED atxt 
            Height          =   285
            Index           =   32
            Left            =   1080
            TabIndex        =   0
            ToolTipText     =   "Enter From Lot Tc"
            Top             =   120
            Width           =   675
            _ExtentX        =   1191
            _ExtentY        =   503
            MaxLength       =   3
            DataType        =   4
            IdName          =   "WLOTTCFR"
         End
         Begin MwfCtl.MWCTL_MED atxt 
            Height          =   285
            Index           =   33
            Left            =   1860
            TabIndex        =   2
            ToolTipText     =   "Enter From Lot Year"
            Top             =   120
            Width           =   345
            _ExtentX        =   609
            _ExtentY        =   503
            MaxLength       =   2
            DataType        =   4
            IdName          =   "WLOTYYFR"
         End
         Begin MwfCtl.MWCTL_MED atxt 
            Height          =   285
            Index           =   34
            Left            =   2310
            TabIndex        =   4
            ToolTipText     =   "Enter From Lot Chr"
            Top             =   120
            Width           =   675
            _ExtentX        =   1191
            _ExtentY        =   503
            MaxLength       =   3
            DataType        =   4
            IdName          =   "WLOTCHRFR"
         End
         Begin MwfCtl.MWCTL_MED atxt 
            Height          =   285
            Index           =   36
            Left            =   4140
            TabIndex        =   8
            ToolTipText     =   "Enter From Lot Sr"
            Top             =   120
            Width           =   555
            _ExtentX        =   979
            _ExtentY        =   503
            Alignment       =   2
            Mask            =   "###0"
            MaxLength       =   4
            DataType        =   1
            IdName          =   "WLOTSRFR"
         End
         Begin MwfCtl.MWCTL_MED atxt 
            Height          =   285
            Index           =   40
            Left            =   3090
            TabIndex        =   6
            ToolTipText     =   "Enter From Lot No"
            Top             =   120
            Width           =   980
            _ExtentX        =   1720
            _ExtentY        =   503
            Alignment       =   2
            Mask            =   "######0"
            MaxLength       =   7
            DataType        =   1
            IdName          =   "WLOTNOFR"
         End
         Begin MwfCtl.MWCTL_MED atxt 
            Height          =   285
            Index           =   35
            Left            =   1080
            TabIndex        =   1
            ToolTipText     =   "Enter To Lot Tc "
            Top             =   400
            Width           =   675
            _ExtentX        =   1191
            _ExtentY        =   503
            MaxLength       =   3
            DataType        =   4
            IdName          =   "WLOTTCTO"
         End
         Begin MwfCtl.MWCTL_MED atxt 
            Height          =   285
            Index           =   41
            Left            =   1860
            TabIndex        =   3
            ToolTipText     =   "Enter To Lot Year"
            Top             =   400
            Width           =   345
            _ExtentX        =   609
            _ExtentY        =   503
            MaxLength       =   2
            DataType        =   4
            IdName          =   "WLOTYYTO"
         End
         Begin MwfCtl.MWCTL_MED atxt 
            Height          =   285
            Index           =   43
            Left            =   2310
            TabIndex        =   5
            ToolTipText     =   "Enter To Lot Chr"
            Top             =   400
            Width           =   675
            _ExtentX        =   1191
            _ExtentY        =   503
            MaxLength       =   3
            DataType        =   4
            IdName          =   "WLOTCHRTO"
         End
         Begin MwfCtl.MWCTL_MED atxt 
            Height          =   285
            Index           =   44
            Left            =   4140
            TabIndex        =   9
            ToolTipText     =   "Enter To Lot Sr"
            Top             =   400
            Width           =   555
            _ExtentX        =   979
            _ExtentY        =   503
            Alignment       =   2
            Mask            =   "###0"
            MaxLength       =   4
            DataType        =   1
            IdName          =   "WLOTSRTO"
         End
         Begin MwfCtl.MWCTL_MED atxt 
            Height          =   285
            Index           =   45
            Left            =   3090
            TabIndex        =   7
            ToolTipText     =   "EnterTo Lot No"
            Top             =   400
            Width           =   980
            _ExtentX        =   1720
            _ExtentY        =   503
            Alignment       =   2
            Mask            =   "######0"
            MaxLength       =   7
            DataType        =   1
            IdName          =   "WLOTNOTO"
         End
         Begin MwfCtl.MWCTL_MED atxt 
            Height          =   285
            Index           =   54
            Left            =   4920
            TabIndex        =   10
            ToolTipText     =   "Enter From Lot Code"
            Top             =   120
            Width           =   2715
            _ExtentX        =   4789
            _ExtentY        =   503
            MaxLength       =   16
            DataType        =   4
            IdName          =   "WLOTCDFR"
         End
         Begin MwfCtl.MWCTL_MED atxt 
            Height          =   285
            Index           =   55
            Left            =   4920
            TabIndex        =   11
            ToolTipText     =   "Enter To Lot Code"
            Top             =   405
            Width           =   2715
            _ExtentX        =   4789
            _ExtentY        =   503
            MaxLength       =   16
            DataType        =   4
            IdName          =   "WLOTCDTO"
         End
         Begin MwfCtl.MWCTL_MED atxt 
            Height          =   285
            Index           =   46
            Left            =   7920
            TabIndex        =   12
            ToolTipText     =   "Enter From Rm Category"
            Top             =   120
            Width           =   555
            _ExtentX        =   979
            _ExtentY        =   503
            MaxLength       =   2
            DataType        =   4
            IdName          =   "WLOTRMCTGFR"
         End
         Begin MwfCtl.MWCTL_MED atxt 
            Height          =   285
            Index           =   49
            Left            =   7920
            TabIndex        =   13
            ToolTipText     =   "Enter To Rm Category"
            Top             =   405
            Width           =   555
            _ExtentX        =   979
            _ExtentY        =   503
            MaxLength       =   2
            DataType        =   4
            IdName          =   "WLOTRMCTGTO"
         End
         Begin MwfCtl.MWCTL_MED atxt 
            Height          =   285
            Index           =   50
            Left            =   8760
            TabIndex        =   14
            ToolTipText     =   "Enter From Currency"
            Top             =   120
            Width           =   915
            _ExtentX        =   1614
            _ExtentY        =   503
            MaxLength       =   5
            DataType        =   4
            IdName          =   "WLOTCURRFR"
         End
         Begin MwfCtl.MWCTL_MED atxt 
            Height          =   285
            Index           =   51
            Left            =   8760
            TabIndex        =   15
            ToolTipText     =   "Enter To Currency"
            Top             =   405
            Width           =   915
            _ExtentX        =   1614
            _ExtentY        =   503
            MaxLength       =   5
            DataType        =   4
            IdName          =   "WLOTCURRTO"
         End
         Begin MwfCtl.MWCTL_MED atxt 
            Height          =   285
            Index           =   15
            Left            =   9960
            TabIndex        =   16
            ToolTipText     =   "Enter Lot Sort By Option"
            Top             =   120
            Width           =   1335
            _ExtentX        =   2355
            _ExtentY        =   503
            MaxLength       =   8
            DataType        =   4
            IdName          =   "WLOTSORT"
         End
         Begin VB.Label Label3 
            AutoSize        =   -1  'True
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
            Height          =   240
            Left            =   2250
            TabIndex        =   75
            Top             =   480
            Width           =   60
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
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
            Height          =   240
            Left            =   1800
            TabIndex        =   74
            Top             =   480
            Width           =   60
         End
         Begin VB.Label Label8 
            AutoSize        =   -1  'True
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
            Height          =   240
            Index           =   2
            Left            =   3030
            TabIndex        =   73
            Top             =   480
            Width           =   60
         End
         Begin VB.Label Label8 
            AutoSize        =   -1  'True
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
            Height          =   240
            Index           =   0
            Left            =   4065
            TabIndex        =   72
            Top             =   480
            Width           =   60
         End
         Begin VB.Label LblSlash 
            AutoSize        =   -1  'True
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
            Height          =   240
            Left            =   1800
            TabIndex        =   71
            Top             =   120
            Width           =   60
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
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
            Height          =   240
            Index           =   0
            Left            =   2250
            TabIndex        =   70
            Top             =   120
            Width           =   60
         End
         Begin VB.Label Label8 
            AutoSize        =   -1  'True
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
            Height          =   240
            Index           =   1
            Left            =   3030
            TabIndex        =   69
            Top             =   120
            Width           =   60
         End
         Begin VB.Label Label7 
            AutoSize        =   -1  'True
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
            Height          =   240
            Left            =   4065
            TabIndex        =   68
            Top             =   120
            Width           =   60
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
            Index           =   1
            Left            =   120
            TabIndex        =   67
            Top             =   480
            Width           =   825
         End
         Begin VB.Label ALBL 
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
            Height          =   285
            Index           =   0
            Left            =   120
            TabIndex        =   66
            Top             =   120
            Width           =   825
         End
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
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
         Height          =   240
         Index           =   1
         Left            =   10080
         TabIndex        =   76
         Top             =   120
         Width           =   690
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "LotCode"
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
         Left            =   5040
         TabIndex        =   64
         Top             =   120
         Width           =   795
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
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
         Height          =   240
         Index           =   12
         Left            =   8880
         TabIndex        =   63
         Top             =   120
         Width           =   720
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
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
         Height          =   240
         Index           =   11
         Left            =   8040
         TabIndex        =   62
         Top             =   120
         Width           =   630
      End
      Begin VB.Label ALBL 
         BackStyle       =   0  'Transparent
         Caption         =   "Scope"
         BeginProperty Font 
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
         Left            =   240
         TabIndex        =   52
         Top             =   120
         Width           =   825
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Tc/Yy/Chr/No/Sr"
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
         Left            =   1200
         TabIndex        =   35
         Top             =   120
         Width           =   1485
      End
   End
   Begin MwfCtl.MWCTL_HLP MWCTL_HLP1 
      Height          =   435
      Left            =   9750
      TabIndex        =   34
      TabStop         =   0   'False
      Top             =   9270
      Visible         =   0   'False
      Width           =   645
      _ExtentX        =   1138
      _ExtentY        =   767
   End
   Begin MwfCtl.MWCTL_ADC ADC 
      CausesValidation=   0   'False
      Height          =   555
      Left            =   30
      TabIndex        =   32
      Top             =   9270
      Width           =   14985
      _ExtentX        =   26432
      _ExtentY        =   979
      AllowAdd        =   0   'False
      AllowDelete     =   0   'False
      BtnWidth        =   885
      BtnHeight       =   435
      BtnMaskColor    =   8454143
      FontSz          =   8.25
   End
   Begin VB.Frame FRA_NKEYS 
      BorderStyle     =   0  'None
      Enabled         =   0   'False
      Height          =   8910
      Left            =   -150
      TabIndex        =   36
      Top             =   495
      Width           =   15315
      Begin VB.Frame FraNKeyAll 
         BorderStyle     =   0  'None
         Height          =   8745
         Left            =   60
         TabIndex        =   37
         Top             =   0
         Width           =   15195
         Begin VB.Frame FraTxnd 
            BorderStyle     =   0  'None
            Height          =   3555
            Left            =   120
            TabIndex        =   60
            Top             =   5160
            Width           =   14925
            Begin MwfCtl.MWCTL_FLX GrdTxnd 
               Height          =   3600
               Left            =   0
               TabIndex        =   61
               Top             =   0
               Width           =   14835
               _ExtentX        =   26167
               _ExtentY        =   6350
               Cols            =   12
               AllowAdd        =   0   'False
               AllowDelete     =   0   'False
               colname1        =   "WTDLOTCD"
               heading1        =   "LotCode                                     "
               maxlength1      =   16
               tooltiptext1    =   "Lot Code"
               colname2        =   "WTDVCHNO"
               heading2        =   "Voucher No.                                "
               maxlength2      =   25
               tooltiptext2    =   "Voucher No."
               colname3        =   "WTDDATE"
               heading3        =   "Date         "
               datatype3       =   3
               maxlength3      =   8
               tooltiptext3    =   "Voucher Date"
               colname4        =   "WTDRMCTG"
               heading4        =   "RmCtg"
               datatype4       =   4
               maxlength4      =   3
               tooltiptext4    =   "Rm Category"
               colname5        =   "WTDQW"
               heading5        =   "QW"
               datatype5       =   4
               maxlength5      =   1
               tooltiptext5    =   "QW"
               colname6        =   "WTDCURRCD"
               heading6        =   "Currency"
               datatype6       =   4
               maxlength6      =   5
               tooltiptext6    =   "Currency"
               colname7        =   "WTDRTBYW"
               heading7        =   "Rate W           "
               datatype7       =   2
               mask7           =   "######0.##0"
               maxlength7      =   11
               tooltiptext7    =   "Rate W"
               colname8        =   "WTDRTBYQ"
               heading8        =   "Rate Q            "
               datatype8       =   2
               mask8           =   "######0.##0"
               maxlength8      =   11
               tooltiptext8    =   "Rate Q"
               colname9        =   "WTDRMCD"
               heading9        =   "RmCode                                  "
               datatype9       =   4
               maxlength9      =   16
               tooltiptext9    =   "Rm Code"
               colname10       =   "WTDRMQTY"
               heading10       =   "Qty      "
               datatype10      =   1
               mask10          =   "####0"
               maxlength10     =   5
               tooltiptext10   =   "Rm Qty"
               colname11       =   "WTDRMWT"
               heading11       =   "Rm Wt         "
               datatype11      =   2
               mask11          =   "####0.000"
               maxlength11     =   9
               tooltiptext11   =   "Rm Wt"
            End
         End
         Begin VB.Frame FraLotMst 
            BorderStyle     =   0  'None
            Height          =   3675
            Left            =   0
            TabIndex        =   53
            Top             =   1440
            Width           =   14925
            Begin MwfCtl.MWCTL_FLX GrdLotMst 
               Height          =   3360
               Left            =   120
               TabIndex        =   54
               Top             =   0
               Width           =   14835
               _ExtentX        =   26167
               _ExtentY        =   5927
               Cols            =   13
               AllowAdd        =   0   'False
               AllowDelete     =   0   'False
               colname1        =   "WLMLOTCD"
               heading1        =   "LotCode                                     "
               datatype1       =   4
               maxlength1      =   16
               tooltiptext1    =   "Lot Code"
               colname2        =   "WLMVCHNO"
               heading2        =   "Voucher No.                                "
               datatype2       =   4
               maxlength2      =   25
               tooltiptext2    =   "Lot Voucher No."
               colname3        =   "WLMDATE"
               heading3        =   "Date         "
               datatype3       =   3
               maxlength3      =   8
               tooltiptext3    =   "Lot Voucher Date"
               colname4        =   "WLMRMCTG"
               heading4        =   "RmCtg"
               datatype4       =   4
               maxlength4      =   3
               tooltiptext4    =   "Lot Rm Category"
               colname5        =   "WLMQW"
               heading5        =   "QW"
               datatype5       =   4
               maxlength5      =   1
               tooltiptext5    =   "Lot QW"
               colname6        =   "WLMCURRCD"
               heading6        =   "Currency"
               datatype6       =   4
               maxlength6      =   5
               tooltiptext6    =   "Lot Currency"
               colname7        =   "WLMRTBYW"
               heading7        =   "Rate W           "
               datatype7       =   2
               mask7           =   "######0.##0"
               maxlength7      =   11
               tooltiptext7    =   "Lot Rate W"
               colname8        =   "WLMRTBYQ"
               heading8        =   "Rate Q            "
               datatype8       =   2
               mask8           =   "######0.##0"
               maxlength8      =   11
               tooltiptext8    =   "Lot Rate Q"
               colname9        =   "WLMBASERTBYW"
               datatype9       =   2
               mask9           =   "######0.##0"
               maxlength9      =   11
               style9          =   0
               colname10       =   "WLMBASERTBYQ"
               heading10       =   " "
               datatype10      =   2
               mask10          =   "######0.##0"
               maxlength10     =   11
               style10         =   0
               colname11       =   "WLMRSRTBYW"
               heading11       =   "   "
               datatype11      =   2
               mask11          =   "######0.##0"
               maxlength11     =   11
               style11         =   0
               colname12       =   "WLMRSRTBYQ"
               datatype12      =   2
               mask12          =   "######0.##0"
               maxlength12     =   11
               style12         =   0
            End
            Begin MwfCtl.MWCTL_MED atxt 
               Height          =   285
               Index           =   53
               Left            =   6120
               TabIndex        =   55
               ToolTipText     =   "Rs. Rate W"
               Top             =   3360
               Width           =   1275
               _ExtentX        =   2249
               _ExtentY        =   503
               Alignment       =   2
               Enabled         =   0   'False
               Mask            =   "######0.##0"
               MaxLength       =   11
               DataType        =   2
               IdName          =   "WLOTRSRTBYW"
            End
            Begin MwfCtl.MWCTL_MED atxt 
               Height          =   285
               Index           =   56
               Left            =   8520
               TabIndex        =   56
               ToolTipText     =   "Rs. Rate Q"
               Top             =   3360
               Width           =   1275
               _ExtentX        =   2249
               _ExtentY        =   503
               Alignment       =   2
               Enabled         =   0   'False
               Mask            =   "######0.##0"
               MaxLength       =   11
               DataType        =   2
               IdName          =   "WLOTRSRTBYQ"
            End
            Begin MwfCtl.MWCTL_MED atxt 
               Height          =   285
               Index           =   21
               Left            =   1080
               TabIndex        =   90
               ToolTipText     =   "Currency Code"
               Top             =   0
               Width           =   1275
               _ExtentX        =   2249
               _ExtentY        =   503
               Enabled         =   0   'False
               MaxLength       =   5
               DataType        =   4
               IdName          =   "WLOTCURR"
            End
            Begin MwfCtl.MWCTL_MED atxt 
               Height          =   285
               Index           =   22
               Left            =   3480
               TabIndex        =   93
               ToolTipText     =   "Base Rate Q"
               Top             =   3360
               Width           =   1275
               _ExtentX        =   2249
               _ExtentY        =   503
               Alignment       =   2
               Enabled         =   0   'False
               Mask            =   "######0.##0"
               MaxLength       =   11
               DataType        =   2
               IdName          =   "WLOTBASERTBYQ"
            End
            Begin MwfCtl.MWCTL_MED atxt 
               Height          =   285
               Index           =   23
               Left            =   1200
               TabIndex        =   94
               ToolTipText     =   "Base Rate W"
               Top             =   3360
               Width           =   1275
               _ExtentX        =   2249
               _ExtentY        =   503
               Alignment       =   2
               Enabled         =   0   'False
               Mask            =   "######0.##0"
               MaxLength       =   11
               DataType        =   2
               IdName          =   "WLOTBASERTBYW"
            End
            Begin VB.Label Label6 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Base Rt Q"
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
               Left            =   2520
               TabIndex        =   92
               Top             =   3360
               Width           =   930
            End
            Begin VB.Label Label6 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Base Rt-W"
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
               Index           =   3
               Left            =   0
               TabIndex        =   91
               Top             =   0
               Width           =   975
            End
            Begin VB.Label Label6 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Base Rt W"
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
               Left            =   120
               TabIndex        =   59
               Top             =   3360
               Width           =   975
            End
            Begin VB.Label Label6 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Rs. Rate W"
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
               Left            =   5040
               TabIndex        =   58
               Top             =   3360
               Width           =   1035
            End
            Begin VB.Label Label6 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Rs. Rate Q"
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
               Left            =   7440
               TabIndex        =   57
               Top             =   3360
               Width           =   990
            End
         End
      End
      Begin VB.Frame FraInvDet 
         BorderStyle     =   0  'None
         Enabled         =   0   'False
         Height          =   8235
         Left            =   150
         TabIndex        =   38
         Top             =   510
         Width           =   14985
         Begin MwfCtl.MWCTL_MED atxt 
            Height          =   285
            Index           =   10
            Left            =   840
            TabIndex        =   39
            ToolTipText     =   "Enter Fr Voucher Tc"
            Top             =   150
            Width           =   675
            _ExtentX        =   1191
            _ExtentY        =   503
            Enabled         =   0   'False
            MaxLength       =   3
            DataType        =   4
            IdName          =   "WODTC"
         End
         Begin MwfCtl.MWCTL_MED atxt 
            Height          =   285
            Index           =   11
            Left            =   1620
            TabIndex        =   40
            ToolTipText     =   "Enter Fr Voucher Year"
            Top             =   150
            Width           =   345
            _ExtentX        =   609
            _ExtentY        =   503
            Enabled         =   0   'False
            MaxLength       =   2
            DataType        =   4
            IdName          =   "WODYY"
         End
         Begin MwfCtl.MWCTL_MED atxt 
            Height          =   285
            Index           =   12
            Left            =   2070
            TabIndex        =   41
            ToolTipText     =   "Enter Fr Voucher Chr"
            Top             =   150
            Width           =   675
            _ExtentX        =   1191
            _ExtentY        =   503
            Enabled         =   0   'False
            MaxLength       =   3
            DataType        =   4
            IdName          =   "WODCHR"
         End
         Begin MwfCtl.MWCTL_MED atxt 
            Height          =   285
            Index           =   13
            Left            =   3780
            TabIndex        =   42
            ToolTipText     =   "Enter Fr Voucher Sr"
            Top             =   150
            Width           =   555
            _ExtentX        =   979
            _ExtentY        =   503
            Enabled         =   0   'False
            Mask            =   "###0"
            MaxLength       =   4
            DataType        =   1
            IdName          =   "WODSR"
         End
         Begin MwfCtl.MWCTL_MED atxt 
            Height          =   285
            Index           =   14
            Left            =   2850
            TabIndex        =   43
            ToolTipText     =   "Enter Fr Voucher No"
            Top             =   150
            Width           =   795
            _ExtentX        =   1402
            _ExtentY        =   503
            Enabled         =   0   'False
            Mask            =   "#####0"
            MaxLength       =   6
            DataType        =   1
            IdName          =   "WODNO"
         End
         Begin VB.Label Label4 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "/"
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
            Left            =   3645
            TabIndex        =   48
            Top             =   195
            Width           =   135
         End
         Begin VB.Label Label8 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "/"
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
            Index           =   3
            Left            =   2730
            TabIndex        =   47
            Top             =   195
            Width           =   135
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "/"
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
            Index           =   1
            Left            =   1950
            TabIndex        =   46
            Top             =   195
            Width           =   135
         End
         Begin VB.Label Label5 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "/"
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
            Left            =   1500
            TabIndex        =   45
            Top             =   195
            Width           =   135
         End
         Begin VB.Label Label6 
            AutoSize        =   -1  'True
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
            Height          =   240
            Index           =   2
            Left            =   90
            TabIndex        =   44
            Top             =   150
            Width           =   750
         End
      End
      Begin VB.Frame FraOrdDsg 
         BorderStyle     =   0  'None
         Height          =   6480
         Left            =   90
         TabIndex        =   49
         Top             =   2250
         Width           =   15270
      End
   End
   Begin VB.Frame FraCpySel 
      Height          =   2415
      Left            =   7680
      TabIndex        =   50
      Top             =   6810
      Visible         =   0   'False
      Width           =   2655
      Begin MwfCtl.MWCTL_MED atxt 
         Height          =   2205
         Index           =   28
         Left            =   60
         TabIndex        =   51
         ToolTipText     =   "ZOOM Selection"
         Top             =   120
         Width           =   2535
         _ExtentX        =   4471
         _ExtentY        =   3889
         Alignment       =   3
         MaxLength       =   1000
         IdName          =   "WZOOMSEL"
      End
   End
End
Attribute VB_Name = "EmrFrmLotZoom"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim mf_LotMstNo As Single
Dim moCn As MwfLib.MDOConnection

Private Sub ADC_Load()
  Set moCn = ADC.Connection
  FRA_KEYS.ZOrder (0)
  
  Set ADC.FirKeyCtl = ADC("WLOTTCFR")
  Set ADC.FirNKeyCtl = GrdLotMst
  
    
End Sub
Private Sub ADC_KeyWhen()
  Me.Caption = GetFrmCaption(gs_CoCd, "Lot Zoom")
  
  FraNKeyAll.ZOrder (0)
  GrdLotMst.Rows = 1
  GrdTxnd.Rows = 1
    
  mf_LotMstNo = 0
  ADC("WORPHYN") = "N"
  If ADC("WORPHYN") = "Y" Then ChkBoxArr(0).Value = Checked Else ChkBoxArr(0).Value = Unchecked

End Sub

Private Sub ADC_NonKeyWhen(pr_Cancel As Boolean, pr_ErrMsg As String)
'Set the Values to GrdLot
  Call FillLotMst
  If GrdLotMst.Rows > 1 Then
    GrdLotMst.Row = 1: GrdLotMst.SetFocus
  End If

End Sub
Private Sub ADC_SetRecSource()
    '*** Since this is an 'A' Entry, there is no RecSource associated with ADC. So set the
    'DataMode property of ADC control as xdatUnbound
  ADC.DataMode = xdatUnbound
End Sub

Private Sub ATXT_FldWhen(Index As Integer, ByVal IdName As String, Cancel As Boolean, ErrMsg As String)
Select Case UCase(IdName)
  Case Is = UCase("WLOTTCFR"), UCase("WLOTTCTO"), UCase("WTDTCFR"), UCase("WTDTCTO")
    Call HlpList.PMCd("TC")
  Case Is = UCase("WLOTYYFR"), UCase("WLOTYYTO"), UCase("WTDYYFR"), UCase("WTDYYTO")
    Call HlpList.PMCd("YY")
  Case Is = UCase("WLOTCHRFR")
    Call HlpList.vPSCd(gs_CoCd, "CHR", ADC("WLOTTCFR"))
  Case Is = UCase("WLOTCHRTO")
    Call HlpList.vPSCd(gs_CoCd, "CHR", ADC("WLOTTCTO"))
  Case Is = UCase("WTDCHRFR")
    Call HlpList.vPSCd(gs_CoCd, "CHR", ADC("WTDTCFR"))
  Case Is = UCase("WTDCHRTO")
    Call HlpList.vPSCd(gs_CoCd, "CHR", ADC("WTDTCTO"))
  Case Is = UCase("WLOTRMCTGFR"), UCase("WLOTRMCTGTO"), UCase("WTDRMCTGFR"), UCase("WTDRMCTGTO")
    Call HlpList.PMCd("RMCTG")
  Case Is = UCase("WLOTCURRFR"), UCase("WLOTCURRTO"), UCase("WTDCURRFR"), UCase("WTDCURRTO")
    Call HlpList.PMCd("CURNCY")
  Case Is = UCase("WLOTSORT")
    Call HlpList.PMCd("SRTMLOT")
  Case Is = UCase("WTDSORT")
    Call HlpList.PMCd("SRTTDLOT")
       
End Select
End Sub

Private Sub ATXT_Validate(Index As Integer, Cancel As Boolean)
  Cancel = ATXT(Index).Validate
End Sub

Private Sub ChkBoxArr_Click(Index As Integer)
If ChkBoxArr(0).Value = Checked Then ADC("WORPHYN") = "Y" Else ADC("WORPHYN") = "N"

End Sub

Private Sub ChkBoxArr_GotFocus(Index As Integer)
DispMsg ChkBoxArr(Index).ToolTipText, etInfo
End Sub

Private Sub Form_GotFocus()

  SetMDISysColorsAfterLogin

End Sub
Private Sub Form_Load()
  Dim ws_BtnStr As String
  ws_BtnStr = ""
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
Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
  If UnloadMode <> vbFormCode Then Cancel = True: Exit Sub
    Set moCn = Nothing
End Sub

Public Sub FillLotMst()
'Gets values from LotMst as per the values in scope and display the same in Grid LotMst.
'If Lotmst grid has value then it shows the transaction details in Grid Txnd
  Dim wCnd As String, wsql As String, wsort As String
  Dim wo_rsLotMst As MwfLib.MDORowSet, wi_Row As Integer
    
  If ADC("WLOTTCFR") <> "" Then wCnd = wCnd + " and LotVchTc>='" + ADC("WLOTTCFR") + "' "
  If ADC("WLOTTCTO") <> "" Then wCnd = wCnd + " and LotVchTc<='" + ADC("WLOTTCTO") + "' "
  If ADC("WLOTYYFR") <> "" Then wCnd = wCnd + " and LotVchYy>='" + ADC("WLOTYYFR") + "' "
  If ADC("WLOTYYTO") <> "" Then wCnd = wCnd + " and LotVchYy<='" + ADC("WLOTYYTO") + "' "
  If ADC("WLOTCHRFR") <> "" Then wCnd = wCnd + " and LotVchChr>='" + ADC("WLOTCHRFR") + "' "
  If ADC("WLOTCHRTO") <> "" Then wCnd = wCnd + " and LotVchChr<='" + ADC("WLOTCHRTO") + "' "
  If ADC("WLOTNOFR") > 0 Then wCnd = wCnd + " and LotVchNo>=" + CStr(ADC("WLOTNOFR")) + " "
  If ADC("WLOTNOTO") > 0 Then wCnd = wCnd + " and LotVchNo<=" + CStr(ADC("WLOTNOTO")) + " "
  If ADC("WLOTSRFR") > 0 Then wCnd = wCnd + " and LotVchSr>=" + CStr(ADC("WLOTSRFR")) + " "
  If ADC("WLOTSRTO") > 0 Then wCnd = wCnd + " and LotVchSr<=" + CStr(ADC("WLOTSRTO")) + " "
  If ADC("WLOTCDFR") <> "" Then wCnd = wCnd + " and LotNo>='" + ADC("WLOTCDFR") + "' "
  If ADC("WLOTCDTO") <> "" Then wCnd = wCnd + " and LotNo<='" + ADC("WLOTCDTO") + "' "
  If ADC("WLOTRMCTGFR") <> "" Then wCnd = wCnd + " and LotRmCtg>='" + ADC("WLOTRMCTGFR") + "' "
  If ADC("WLOTRMCTGTO") <> "" Then wCnd = wCnd + " and LotRmCtg<='" + ADC("WLOTRMCTGTO") + "' "
  If ADC("WLOTCURRFR") <> "" Then wCnd = wCnd + " and LotCurrCd>='" + ADC("WLOTCURRFR") + "' "
  If ADC("WLOTCURRTO") <> "" Then wCnd = wCnd + " and LotCurrCd<='" + ADC("WLOTCURRTO") + "' "
  
  wsort = moCn.GetFldVal("Select PDesc225 from Param where PTyp = 'SRTMLOT' and PmCd= '" + ADC("WLOTSORT") + "' and PSCd= ''")

  If ADC("WORPHYN") = "Y" Then
  wsql = "Select 'OrPhans' as LotNo, '' as LotVchCoCd, '' as LotVchTc, '' as LotVchYy, '' as LotVchChr, '' as LotVchNo, '' as LotVchSr, " + _
         "'' as LotVchDt, '' as LotRmCtg, '' as LotQW, 0 as LotBaseRtByW, 0 as LotBaseRtByQ, " + _
         "'' as LotCurrCd, 0 as LotRtByW, 0 as LotRtByQ, 0 as LotRsRtByW, 0 as LotRsRtByQ"
  
  Else
  wsql = "Select LotNo, LotVchCoCd, LotVchTc, LotVchYy, LotVchChr, LotVchNo, LotVchSr, " + _
         "LotVchDt, LotRmCtg,  LotQW,LotBaseRtByW,LotBaseRtByQ, " + _
         "LotCurrCd, LotRtByW,LotRtByQ,LotRsRtByW,LotRsRtByQ " + _
         "From LotMst Where LotVchCoCd='" + gs_CoCd + "'" + IIF(wCnd <> "", wCnd, "") + " Order by " + _
         IIF(ADC("WLOTSORT") <> "", wsort + IIF(ADC("WLOTSORT") = "LOTVCH", "", ",LotVchTc, LotVchYy, LotVchChr, LotVchNo, LotVchSr"), "LotVchTc, LotVchYy, LotVchChr, LotVchNo, LotVchSr")
            
  End If
  GrdLotMst.Rows = 1
  Set wo_rsLotMst = moCn.OpenRes(wsql)
  If wo_rsLotMst.RecCount <= 2500 Then
    GrdLotMst.AllowAdd = True
    With GrdLotMst
      Do While Not (wo_rsLotMst.EOF Or wo_rsLotMst.BOF)
        .AddItem
        wi_Row = .Rows - 1
        .Value(wi_Row, "WLMLOTCD") = Trim(wo_rsLotMst!LotNo)
        .Value(wi_Row, "WLMVCHNO") = Trim(wo_rsLotMst!LotVchTc) + "/" + Trim(wo_rsLotMst!LotVchYy) + "/" + Trim(wo_rsLotMst!LotVchChr) + "/" + CStr(Trim(wo_rsLotMst!LotVchNo)) + "/" + CStr(Trim(wo_rsLotMst!LotVchSr))
        .Value(wi_Row, "WLMDATE") = Trim(wo_rsLotMst!LotVchDt)
        .Value(wi_Row, "WLMRMCTG") = Trim(wo_rsLotMst!LotRmCtg)
        .Value(wi_Row, "WLMQW") = Trim(wo_rsLotMst!LotQW)
        .Value(wi_Row, "WLMBASERTBYW") = Trim(wo_rsLotMst!LotBaseRtByW)
        .Value(wi_Row, "WLMBASERTBYQ") = Trim(wo_rsLotMst!LotBaseRtByQ)
        
        .Value(wi_Row, "WLMCURRCD") = Trim(wo_rsLotMst!LotCurrCd)
        .Value(wi_Row, "WLMRTBYW") = Trim(wo_rsLotMst!LotRtByW)
        .Value(wi_Row, "WLMRTBYQ") = Trim(wo_rsLotMst!LotRtByQ)
        .Value(wi_Row, "WLMRSRTBYW") = Trim(wo_rsLotMst!LotRsRtByW)
        .Value(wi_Row, "WLMRSRTBYQ") = Trim(wo_rsLotMst!LotRsRtByQ)
                        
        wo_rsLotMst.MoveNext
      Loop
    End With

    If GrdLotMst.Rows > 1 Then
      GrdLotMst.Row = 1
      Call FillTxnd(1)
    End If
    mf_LotMstNo = GrdLotMst.Row
    
  Else
    DispMsg "No. of Records are more than 2500, Plz Limit the Scope", etError
  End If
  Set wo_rsLotMst = Nothing
   
End Sub

Public Sub FillTxnd(ByVal RowNum As Integer)
'Gets values from Txnd as per the values in scope and display the same in Grid.

  Dim wCnd As String, wsql As String, wsort As String
  Dim wo_rsTxnd As MwfLib.MDORowSet, wi_Row As Integer
  
  If ADC("WTDTCFR") <> "" Then wCnd = wCnd + " and TdTc>='" + ADC("WTDTCFR") + "' "
  If ADC("WTDTCTO") <> "" Then wCnd = wCnd + " and TdTc<='" + ADC("WTDTCTO") + "' "
  If ADC("WTDYYFR") <> "" Then wCnd = wCnd + " and TdYy>='" + ADC("WTDYYFR") + "' "
  If ADC("WTDYYTO") <> "" Then wCnd = wCnd + " and TdYy<='" + ADC("WTDYYTO") + "' "
  If ADC("WTDCHRFR") <> "" Then wCnd = wCnd + " and TdChr>='" + ADC("WTDCHRFR") + "' "
  If ADC("WTDCHRTO") <> "" Then wCnd = wCnd + " and TdChr<='" + ADC("WTDCHRTO") + "' "
  If ADC("WTDNOFR") > 0 Then wCnd = wCnd + " and TdNo>=" + CStr(ADC("WTDNOFR")) + " "
  If ADC("WTDNOTO") > 0 Then wCnd = wCnd + " and TdNo<=" + CStr(ADC("WTDNOTO")) + " "
  If ADC("WTDSRFR") > 0 Then wCnd = wCnd + " and TdSr>=" + CStr(ADC("WTDSRFR")) + " "
  If ADC("WTDSRTO") > 0 Then wCnd = wCnd + " and TdSr<=" + CStr(ADC("WTDSRTO")) + " "
  If ADC("WTDRMCTGFR") <> "" Then wCnd = wCnd + " and TdRmCtg>='" + ADC("WTDRMCTGFR") + "' "
  If ADC("WTDRMCTGTO") <> "" Then wCnd = wCnd + " and TdRmCtg<='" + ADC("WTDRMCTGTO") + "' "
  If ADC("WTDCURRFR") <> "" Then wCnd = wCnd + " and TCurrCd>='" + ADC("WTDCURRFR") + "' "
  If ADC("WTDCURRTO") <> "" Then wCnd = wCnd + " and TCurrCd<='" + ADC("WTDCURRTO") + "' "

  wsort = moCn.GetFldVal("Select PDesc225 from Param where PTyp = 'SRTTDLOT' and PmCd= '" + ADC("WTDSORT") + "' and PSCd= ''")
  
  If ADC("WORPHYN") = "Y" Then
  wsql = "Select TdLotNo, TdCoCd, TdTc, TdYy, TdChr, TdNo, TdSr, TdDt, TdRmCtg, TdRtByQW, " + _
         "TdPurRt, TdRmCd, TdRmQty, TdRmWt, TCurrCd from Txnd " + _
         "Join Txn on TCoCd = TdCoCd and TTc= TdTc and TYy= TdYy and TChr=TdChr " + _
         "and TNo=TdNo and TIdNo=TdTIdNo and TPrtKey=TdPrtKey " + _
         "where TdLotMstRecYN='N' and TdCoCd='" + gs_CoCd + "' and TdLotNo <> '' and " + _
         "not exists (Select TdLotNo from LotMst where LotMst.LotNo = TdLotNo)" + _
         IIF(wCnd <> "", wCnd, "") + " Order by " + _
         IIF(ADC("WTDSORT") <> "", wsort + IIF(ADC("WTDSORT") = "TDVCHR", "", ",TdTc, TdYy, TdChr, TdNo, TdSr"), "TdTc, TdYy, TdChr, TdNo, TdSr")
                  
  Else
  wsql = "Select TdLotNo, TdCoCd, TdTc, TdYy, TdChr, TdNo, TdSr, TdDt, TdRmCtg, TdRtByQW, " + _
         "TdPurRt, TdRmCd, TdRmQty, TdRmWt, TCurrCd From Txnd " + _
         "Join LotMst on LotNo=TdLotNo " + _
         "Join Txn on TCoCd = TdCoCd and TTc= TdTc and TYy= TdYy and TChr=TdChr " + _
         "and TNo=TdNo and TIdNo=TdTIdNo and TPrtKey=TdPrtKey " + _
         "Where TdLotMstRecYN='N' and LotVchCoCd='" + gs_CoCd + "' and LotNo='" + GrdLotMst.Value(RowNum, "WLMLOTCD") + "'" + _
         IIF(wCnd <> "", wCnd, "") + " Order by " + _
         IIF(ADC("WTDSORT") <> "", wsort + IIF(ADC("WTDSORT") = "TDVCHR", "", ",TdTc, TdYy, TdChr, TdNo, TdSr"), "TdTc, TdYy, TdChr, TdNo, TdSr")
         
  End If
  
  GrdTxnd.Rows = 1
  Set wo_rsTxnd = moCn.OpenRes(wsql)
  If wo_rsTxnd.RecCount <= 2500 Then
    GrdTxnd.AllowAdd = True
    With GrdTxnd
      Do While Not (wo_rsTxnd.EOF Or wo_rsTxnd.BOF)
        .AddItem
        wi_Row = .Rows - 1
        .Value(wi_Row, "WTDLOTCD") = Trim(wo_rsTxnd!TdLotNo)
        .Value(wi_Row, "WTDVCHNO") = Trim(wo_rsTxnd!TdTc) + "/" + Trim(wo_rsTxnd!TdYy) + "/" + Trim(wo_rsTxnd!TdChr) + "/" + CStr(Trim(wo_rsTxnd!TdNo)) + "/" + CStr(Trim(wo_rsTxnd!TdSr))
        .Value(wi_Row, "WTDDATE") = Trim(wo_rsTxnd!TdDt)
        .Value(wi_Row, "WTDRMCTG") = Trim(wo_rsTxnd!TdRmCtg)
        .Value(wi_Row, "WTDQW") = Trim(wo_rsTxnd!TdRtByQW)
        
        .Value(wi_Row, "WTDCURRCD") = Trim(wo_rsTxnd!TCurrCd)
        If (Trim(wo_rsTxnd!TdRtByQW)) = "W" Then
          .Value(wi_Row, "WTDRTBYW") = Trim(wo_rsTxnd!TdPurRt)
        Else
          .Value(wi_Row, "WTDRTBYQ") = Trim(wo_rsTxnd!TdPurRt)
        End If
        .Value(wi_Row, "WTDRMCD") = Trim(wo_rsTxnd!TdRmCd)
        .Value(wi_Row, "WTDRMQTY") = Trim(wo_rsTxnd!TdRmQty)
        .Value(wi_Row, "WTDRMWT") = Trim(wo_rsTxnd!TdRmWt)
             
        wo_rsTxnd.MoveNext
      Loop
    End With

    If GrdTxnd.Rows > 1 Then
        GrdTxnd.Row = 1
    End If
        
  Else
    DispMsg "No. of Records are more than 2500, Plz Limit the Scope", etError
  End If
  Set wo_rsTxnd = Nothing
End Sub

Private Sub GrdLotMst_ColWhen(ByVal RowNum As Integer, ByVal ColName As String, Cancel As Boolean, ErrMsg As String, ByVal Action As MwfCtl.en_fgAction)
 'can't enter values to Grid
  Cancel = True
End Sub

Private Sub GrdLotMst_RowWhen(ByVal RowNum As Integer)
'When the row change, the values in txt boxes below LotMst grid & Txnd Grid  will change
If RowNum > 0 Then
  With GrdLotMst
    ADC("WLOTBASERTBYW") = .Value(RowNum, "WLMBASERTBYW")
    ADC("WLOTBASERTBYQ") = .Value(RowNum, "WLMBASERTBYQ")
    ADC("WLOTRSRTBYW") = .Value(RowNum, "WLMRSRTBYW")
    ADC("WLOTRSRTBYQ") = .Value(RowNum, "WLMRSRTBYQ")
    
  End With
  
  If mf_LotMstNo <> GrdLotMst.Row Then
    If GrdLotMst.Value(RowNum, "WLMLOTCD") <> "" Then Call FillTxnd(RowNum)
    mf_LotMstNo = GrdLotMst.Row
  End If

End If

End Sub

Private Sub GrdTxnd_ColWhen(ByVal RowNum As Integer, ByVal ColName As String, Cancel As Boolean, ErrMsg As String, ByVal Action As MwfCtl.en_fgAction)
 'can't enter values to Grid
  Cancel = True
End Sub

