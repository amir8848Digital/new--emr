VERSION 5.00
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#2.0#0"; "mscomctl.ocx"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Object = "{6A2CE895-C1AC-48AF-B558-FB94BBE12499}#1.0#0"; "mwfctl.ocx"
Begin VB.Form EmrFrmTrfOut 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Emr To XML Transfer"
   ClientHeight    =   9810
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   15210
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MDIChild        =   -1  'True
   ScaleHeight     =   9810
   ScaleWidth      =   15210
   WindowState     =   2  'Maximized
   Begin MSComDlg.CommonDialog ComDialogOpn 
      Left            =   12210
      Top             =   5190
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.Frame FRA_KEYS 
      BorderStyle     =   0  'None
      Height          =   450
      Left            =   12060
      TabIndex        =   182
      Top             =   5910
      Width           =   825
   End
   Begin VB.Frame FRA_NKEYS 
      BorderStyle     =   0  'None
      Enabled         =   0   'False
      Height          =   420
      Left            =   12045
      TabIndex        =   180
      Top             =   6450
      Width           =   1125
      Begin VB.Frame FraNKeyAll 
         BorderStyle     =   0  'None
         Height          =   3705
         Left            =   6030
         TabIndex        =   181
         Top             =   1530
         Width           =   5790
      End
   End
   Begin MwfCtl.MWCTL_HLP MWCTL_HLP1 
      Height          =   405
      Left            =   10005
      TabIndex        =   157
      TabStop         =   0   'False
      Top             =   6690
      Visible         =   0   'False
      Width           =   555
      _ExtentX        =   979
      _ExtentY        =   714
   End
   Begin MwfCtl.MWCTL_ADC ADC 
      CausesValidation=   0   'False
      Height          =   555
      Left            =   11100
      TabIndex        =   179
      Top             =   9180
      Width           =   5415
      _ExtentX        =   9551
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
   Begin MwfCtl.MWCTL_MED ATXT 
      Height          =   570
      Index           =   8
      Left            =   30
      TabIndex        =   185
      ToolTipText     =   "Message"
      Top             =   9165
      Width           =   10980
      _ExtentX        =   19368
      _ExtentY        =   1005
      Alignment       =   3
      Enabled         =   0   'False
      MaxLength       =   255
      DataType        =   4
      IdName          =   "WDISP"
   End
   Begin MwfCtl.MWCTL_MED ATXT 
      Height          =   285
      Index           =   142
      Left            =   2220
      TabIndex        =   0
      ToolTipText     =   "Specify The Target Exe ('Y' For EMR And 'N' For XEMR)"
      Top             =   0
      Width           =   315
      _ExtentX        =   556
      _ExtentY        =   503
      DataType        =   4
      ReCalcParent    =   "WTGTVER"
      IdName          =   "WTGTEXE"
   End
   Begin MwfCtl.MWCTL_MED ATXT 
      Height          =   285
      Index           =   143
      Left            =   4380
      TabIndex        =   1
      ToolTipText     =   "Enter The Version Of the Target Exe"
      Top             =   0
      Width           =   885
      _ExtentX        =   1561
      _ExtentY        =   503
      MaxLength       =   8
      DataType        =   4
      ReCalcOn        =   "WTGTEXE"
      IdName          =   "WTGTVER"
   End
   Begin VB.CommandButton CmdOrd 
      Caption         =   "<&k> Order"
      Height          =   600
      Left            =   9630
      Style           =   1  'Graphical
      TabIndex        =   12
      ToolTipText     =   "Option To Transfer Customer Master"
      Top             =   930
      Width           =   2400
   End
   Begin VB.CommandButton CmdDtTbl 
      Caption         =   "<&h> Date Table"
      Height          =   600
      Left            =   2430
      Style           =   1  'Graphical
      TabIndex        =   9
      ToolTipText     =   "Option To Transfer Customer Master"
      Top             =   930
      Width           =   2400
   End
   Begin VB.CommandButton CmdInv 
      Caption         =   "<&l> Invoice"
      CausesValidation=   0   'False
      Height          =   600
      Left            =   12030
      Style           =   1  'Graphical
      TabIndex        =   13
      ToolTipText     =   "Option To Transfer Customer Master"
      Top             =   930
      Width           =   2400
   End
   Begin VB.CommandButton CmdCust 
      Caption         =   "<&b> Customer"
      Height          =   600
      Left            =   2430
      Style           =   1  'Graphical
      TabIndex        =   3
      ToolTipText     =   "Option To Transfer Customer Master"
      Top             =   330
      Width           =   2400
   End
   Begin VB.CommandButton CmdRm 
      Caption         =   "<&a> Raw Material"
      Height          =   600
      Left            =   30
      Style           =   1  'Graphical
      TabIndex        =   2
      ToolTipText     =   "Option To Transfer Customer Master"
      Top             =   330
      Width           =   2400
   End
   Begin VB.CommandButton CmdDsg 
      Caption         =   "<&c> Design"
      Height          =   600
      Left            =   4830
      Style           =   1  'Graphical
      TabIndex        =   4
      ToolTipText     =   "Option To Transfer Customer Master"
      Top             =   330
      Width           =   2400
   End
   Begin VB.CommandButton CmdRmRt 
      Caption         =   "<&d> Rm Rate"
      Height          =   600
      Left            =   7230
      Style           =   1  'Graphical
      TabIndex        =   5
      ToolTipText     =   "Option To Transfer Customer Master"
      Top             =   330
      Width           =   2400
   End
   Begin VB.CommandButton CmdLabRt 
      Caption         =   "<&e> Labour Rate"
      Height          =   600
      Left            =   9630
      Style           =   1  'Graphical
      TabIndex        =   6
      ToolTipText     =   "Option To Transfer Customer Master"
      Top             =   330
      Width           =   2400
   End
   Begin VB.CommandButton CmdBkAg 
      Caption         =   "<&i> Bank/ Agent"
      Height          =   600
      Left            =   4830
      Style           =   1  'Graphical
      TabIndex        =   10
      ToolTipText     =   "Option To Transfer Customer Master"
      Top             =   930
      Width           =   2400
   End
   Begin VB.CommandButton CmdLoc 
      Caption         =   "<&j> Location"
      Height          =   600
      Left            =   7230
      Style           =   1  'Graphical
      TabIndex        =   11
      ToolTipText     =   "Option To Transfer Customer Master"
      Top             =   930
      Width           =   2400
   End
   Begin VB.CommandButton CmdVPrm 
      Caption         =   "<&g> Variable Param"
      Height          =   600
      Left            =   30
      Style           =   1  'Graphical
      TabIndex        =   8
      ToolTipText     =   "Option To Transfer Customer Master"
      Top             =   930
      Width           =   2400
   End
   Begin VB.CommandButton CmdPrm 
      Caption         =   "<&f> Param"
      Height          =   600
      Left            =   12030
      Style           =   1  'Graphical
      TabIndex        =   7
      ToolTipText     =   "Option To Transfer Customer Master"
      Top             =   330
      Width           =   2400
   End
   Begin VB.CommandButton CmdAllMst 
      Caption         =   "<&m> All Masters"
      CausesValidation=   0   'False
      Height          =   600
      Left            =   30
      Style           =   1  'Graphical
      TabIndex        =   14
      ToolTipText     =   "Option To Transfer All Masters"
      Top             =   1530
      Width           =   2400
   End
   Begin VB.CommandButton CmdLabRtHist 
      Caption         =   "<&o> Labour Rate History"
      Enabled         =   0   'False
      Height          =   600
      Left            =   4830
      Style           =   1  'Graphical
      TabIndex        =   16
      ToolTipText     =   "Option To Transfer Customer Master"
      Top             =   1530
      Visible         =   0   'False
      Width           =   2400
   End
   Begin VB.CommandButton CmdRmRtHist 
      Caption         =   "<&n> Rm Rate History"
      Enabled         =   0   'False
      Height          =   600
      Left            =   2430
      Style           =   1  'Graphical
      TabIndex        =   15
      ToolTipText     =   "Option To Transfer Customer Master"
      Top             =   1530
      Visible         =   0   'False
      Width           =   2400
   End
   Begin VB.Frame FraDtTbl 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1950
      Left            =   5310
      TabIndex        =   282
      Top             =   7125
      Visible         =   0   'False
      Width           =   4005
      Begin VB.CommandButton CmdDtTblTrfOut 
         Caption         =   "&Transfer Out"
         Height          =   315
         Left            =   2490
         TabIndex        =   135
         ToolTipText     =   "Date Table Transfer Out"
         Top             =   1455
         Width           =   1230
      End
      Begin MwfCtl.MWCTL_MED ATXT 
         Height          =   285
         Index           =   112
         Left            =   1350
         TabIndex        =   132
         ToolTipText     =   "Enter From Date"
         Top             =   1155
         Width           =   885
         _ExtentX        =   1561
         _ExtentY        =   503
         MaxLength       =   8
         DataType        =   3
         IdName          =   "WDTDTFR"
      End
      Begin MwfCtl.MWCTL_MED ATXT 
         Height          =   285
         Index           =   113
         Left            =   2835
         TabIndex        =   133
         ToolTipText     =   "Enter To Date"
         Top             =   1155
         Width           =   885
         _ExtentX        =   1561
         _ExtentY        =   503
         MaxLength       =   8
         DataType        =   3
         IdName          =   "WDTDTTO"
      End
      Begin MwfCtl.MWCTL_MED ATXT 
         Height          =   285
         Index           =   114
         Left            =   1350
         TabIndex        =   131
         ToolTipText     =   "Enter Date Tag"
         Top             =   870
         Width           =   885
         _ExtentX        =   1561
         _ExtentY        =   503
         MaxLength       =   5
         DataType        =   4
         IdName          =   "WDTTAG"
      End
      Begin MwfCtl.MWCTL_MED ATXT 
         Height          =   285
         Index           =   115
         Left            =   1350
         TabIndex        =   134
         ToolTipText     =   "Enter From Date Table Modified Date"
         Top             =   1440
         Width           =   885
         _ExtentX        =   1561
         _ExtentY        =   503
         MaxLength       =   8
         DataType        =   3
         IdName          =   "WDTMODDTFR"
      End
      Begin VB.Label LblMain 
         BackStyle       =   0  'Transparent
         BorderStyle     =   1  'Fixed Single
         Caption         =   "  Date Table Transfer"
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
         TabIndex        =   333
         Top             =   90
         Width           =   4005
      End
      Begin VB.Label ALBL 
         BackStyle       =   0  'Transparent
         Caption         =   "Modified Dt"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Index           =   88
         Left            =   90
         TabIndex        =   287
         Top             =   1440
         Width           =   1245
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
         Index           =   94
         Left            =   1350
         TabIndex        =   286
         Top             =   480
         Width           =   1035
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
         Index           =   93
         Left            =   2820
         TabIndex        =   285
         Top             =   480
         Width           =   1035
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
         Index           =   92
         Left            =   90
         TabIndex        =   284
         Top             =   1155
         Width           =   1245
      End
      Begin VB.Label ALBL 
         BackStyle       =   0  'Transparent
         Caption         =   "Date Tag"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Index           =   91
         Left            =   90
         TabIndex        =   283
         Top             =   870
         Width           =   1245
      End
   End
   Begin VB.Frame FraBkAg 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1950
      Left            =   5280
      TabIndex        =   288
      Top             =   7125
      Visible         =   0   'False
      Width           =   4095
      Begin VB.CommandButton CmdBkAgTrfOut 
         Caption         =   "&Transfer Out"
         Height          =   315
         Left            =   2550
         TabIndex        =   140
         ToolTipText     =   "Bank/ gent Transfer Out"
         Top             =   1455
         Width           =   1230
      End
      Begin MwfCtl.MWCTL_MED ATXT 
         Height          =   285
         Index           =   116
         Left            =   1410
         TabIndex        =   137
         ToolTipText     =   "Enter From Bank/ Agent Code"
         Top             =   1155
         Width           =   885
         _ExtentX        =   1561
         _ExtentY        =   503
         MaxLength       =   5
         DataType        =   4
         IdName          =   "WBKCDFR"
      End
      Begin MwfCtl.MWCTL_MED ATXT 
         Height          =   285
         Index           =   117
         Left            =   2895
         TabIndex        =   138
         ToolTipText     =   "Enter To Bank/ Agent Code"
         Top             =   1155
         Width           =   885
         _ExtentX        =   1561
         _ExtentY        =   503
         MaxLength       =   5
         DataType        =   4
         IdName          =   "WBKCDTO"
      End
      Begin MwfCtl.MWCTL_MED ATXT 
         Height          =   285
         Index           =   118
         Left            =   1410
         TabIndex        =   136
         ToolTipText     =   "Enter Bank/ Agent Category"
         Top             =   870
         Width           =   885
         _ExtentX        =   1561
         _ExtentY        =   503
         DataType        =   4
         IdName          =   "WBKCTG"
      End
      Begin MwfCtl.MWCTL_MED ATXT 
         Height          =   285
         Index           =   119
         Left            =   1410
         TabIndex        =   139
         ToolTipText     =   "Enter From Bank/ Agent Modified Date"
         Top             =   1440
         Width           =   885
         _ExtentX        =   1561
         _ExtentY        =   503
         MaxLength       =   8
         DataType        =   3
         IdName          =   "WBKMODDTFR"
      End
      Begin VB.Label LblMain 
         BackStyle       =   0  'Transparent
         BorderStyle     =   1  'Fixed Single
         Caption         =   "  Bank/ Agent Transfer"
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
         TabIndex        =   332
         Top             =   90
         Width           =   4095
      End
      Begin VB.Label ALBL 
         BackStyle       =   0  'Transparent
         Caption         =   "Bank/ Agent"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Index           =   97
         Left            =   90
         TabIndex        =   293
         Top             =   870
         Width           =   1305
      End
      Begin VB.Label ALBL 
         BackStyle       =   0  'Transparent
         Caption         =   "Code"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Index           =   96
         Left            =   90
         TabIndex        =   292
         Top             =   1155
         Width           =   1245
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
         Index           =   95
         Left            =   2880
         TabIndex        =   291
         Top             =   480
         Width           =   1035
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
         Index           =   90
         Left            =   1410
         TabIndex        =   290
         Top             =   480
         Width           =   1035
      End
      Begin VB.Label ALBL 
         BackStyle       =   0  'Transparent
         Caption         =   "Modified Dt"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Index           =   89
         Left            =   90
         TabIndex        =   289
         Top             =   1440
         Width           =   1245
      End
   End
   Begin VB.Frame FraLoc 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1830
      Left            =   5280
      TabIndex        =   294
      Top             =   7245
      Visible         =   0   'False
      Width           =   4095
      Begin VB.CommandButton CmdLocTrfOut 
         Caption         =   "&Transfer Out"
         Height          =   315
         Left            =   2550
         TabIndex        =   162
         ToolTipText     =   "Location Transfer Out"
         Top             =   1455
         Width           =   1230
      End
      Begin MwfCtl.MWCTL_MED ATXT 
         Height          =   285
         Index           =   120
         Left            =   1410
         TabIndex        =   159
         ToolTipText     =   "Enter From Location Code"
         Top             =   1155
         Width           =   885
         _ExtentX        =   1561
         _ExtentY        =   503
         MaxLength       =   5
         DataType        =   4
         IdName          =   "WLOCCDFR"
      End
      Begin MwfCtl.MWCTL_MED ATXT 
         Height          =   285
         Index           =   121
         Left            =   2895
         TabIndex        =   160
         ToolTipText     =   "Enter To Location Code"
         Top             =   1155
         Width           =   885
         _ExtentX        =   1561
         _ExtentY        =   503
         MaxLength       =   5
         DataType        =   4
         IdName          =   "WLOCCDTO"
      End
      Begin MwfCtl.MWCTL_MED ATXT 
         Height          =   285
         Index           =   122
         Left            =   1410
         TabIndex        =   158
         ToolTipText     =   "Enter Location Type"
         Top             =   870
         Width           =   465
         _ExtentX        =   820
         _ExtentY        =   503
         MaxLength       =   2
         DataType        =   4
         IdName          =   "WLOCTYP"
      End
      Begin MwfCtl.MWCTL_MED ATXT 
         Height          =   285
         Index           =   123
         Left            =   1410
         TabIndex        =   161
         ToolTipText     =   "Enter From Location Modified Date"
         Top             =   1440
         Width           =   885
         _ExtentX        =   1561
         _ExtentY        =   503
         MaxLength       =   8
         DataType        =   3
         IdName          =   "WLOCMODDTFR"
      End
      Begin VB.Label LblMain 
         BackStyle       =   0  'Transparent
         BorderStyle     =   1  'Fixed Single
         Caption         =   "  Location Transfer"
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
         TabIndex        =   331
         Top             =   90
         Width           =   4095
      End
      Begin VB.Label ALBL 
         BackStyle       =   0  'Transparent
         Caption         =   "Modified Dt"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Index           =   102
         Left            =   90
         TabIndex        =   299
         Top             =   1440
         Width           =   1245
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
         Index           =   101
         Left            =   1410
         TabIndex        =   298
         Top             =   480
         Width           =   1035
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
         Index           =   100
         Left            =   2880
         TabIndex        =   297
         Top             =   480
         Width           =   1035
      End
      Begin VB.Label ALBL 
         BackStyle       =   0  'Transparent
         Caption         =   "Loc Code"
         BeginProperty Font 
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
         Left            =   90
         TabIndex        =   296
         Top             =   1155
         Width           =   1245
      End
      Begin VB.Label ALBL 
         BackStyle       =   0  'Transparent
         Caption         =   "Loc Type"
         BeginProperty Font 
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
         Left            =   90
         TabIndex        =   295
         Top             =   870
         Width           =   1305
      End
   End
   Begin VB.Frame FraCust 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2310
      Left            =   5010
      TabIndex        =   209
      Top             =   6765
      Visible         =   0   'False
      Width           =   4635
      Begin VB.CommandButton CmdCustTrfOut 
         Caption         =   "&Transfer Out"
         Height          =   315
         Left            =   3180
         TabIndex        =   49
         ToolTipText     =   "Customer Transfer Out"
         Top             =   1725
         Width           =   1230
      End
      Begin MwfCtl.MWCTL_MED ATXT 
         Height          =   285
         Index           =   32
         Left            =   1350
         TabIndex        =   45
         ToolTipText     =   "Enter From Customer Code"
         Top             =   1155
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   503
         MaxLength       =   8
         DataType        =   4
         IdName          =   "WCMCDFR"
      End
      Begin MwfCtl.MWCTL_MED ATXT 
         Height          =   285
         Index           =   33
         Left            =   3075
         TabIndex        =   46
         ToolTipText     =   "Enter From Customer Code"
         Top             =   1155
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   503
         MaxLength       =   8
         DataType        =   4
         IdName          =   "WCMCDTO"
      End
      Begin MwfCtl.MWCTL_MED ATXT 
         Height          =   285
         Index           =   34
         Left            =   1350
         TabIndex        =   44
         ToolTipText     =   "Enter Customer Category"
         Top             =   870
         Width           =   495
         _ExtentX        =   873
         _ExtentY        =   503
         MaxLength       =   2
         DataType        =   4
         IdName          =   "WCMCTG"
      End
      Begin MwfCtl.MWCTL_MED ATXT 
         Height          =   285
         Index           =   35
         Left            =   1350
         TabIndex        =   47
         ToolTipText     =   "Enter From Order Modified Date"
         Top             =   1440
         Width           =   885
         _ExtentX        =   1561
         _ExtentY        =   503
         MaxLength       =   8
         DataType        =   3
         IdName          =   "WCMMODDTFR"
      End
      Begin MwfCtl.MWCTL_MED ATXT 
         Height          =   285
         Index           =   138
         Left            =   1350
         TabIndex        =   48
         ToolTipText     =   "Specify Whether To Copy The Child Records (Y/ N)"
         Top             =   1725
         Width           =   315
         _ExtentX        =   556
         _ExtentY        =   503
         DataType        =   4
         IdName          =   "WCMCPYCHILD"
      End
      Begin VB.Label LblMain 
         BackStyle       =   0  'Transparent
         BorderStyle     =   1  'Fixed Single
         Caption         =   "  Customer Transfer"
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
         TabIndex        =   336
         Top             =   90
         Width           =   4635
      End
      Begin VB.Label ALBL 
         BackStyle       =   0  'Transparent
         Caption         =   "Copy Child Recs (Y/ N)"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   465
         Index           =   118
         Left            =   90
         TabIndex        =   316
         Top             =   1725
         Width           =   1305
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
         Index           =   27
         Left            =   1350
         TabIndex        =   214
         Top             =   480
         Width           =   1035
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
         Index           =   26
         Left            =   3060
         TabIndex        =   213
         Top             =   480
         Width           =   1035
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
         Index           =   24
         Left            =   90
         TabIndex        =   212
         Top             =   1155
         Width           =   1245
      End
      Begin VB.Label ALBL 
         BackStyle       =   0  'Transparent
         Caption         =   "Category"
         BeginProperty Font 
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
         TabIndex        =   211
         Top             =   870
         Width           =   1245
      End
      Begin VB.Label ALBL 
         BackStyle       =   0  'Transparent
         Caption         =   "Modified Dt"
         BeginProperty Font 
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
         Left            =   90
         TabIndex        =   210
         Top             =   1440
         Width           =   1245
      End
   End
   Begin VB.Frame FraLabRt 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2790
      Left            =   4860
      TabIndex        =   247
      Top             =   6285
      Visible         =   0   'False
      Width           =   4905
      Begin VB.CommandButton CmdLabRtTrfOut 
         Caption         =   "&Transfer Out"
         Height          =   315
         Left            =   3360
         TabIndex        =   93
         ToolTipText     =   "Labour Rate Transfer Out"
         Top             =   2265
         Width           =   1230
      End
      Begin MwfCtl.MWCTL_MED ATXT 
         Height          =   285
         Index           =   68
         Left            =   1500
         TabIndex        =   84
         ToolTipText     =   "Enter From Customer Code"
         Top             =   1095
         Width           =   1305
         _ExtentX        =   2302
         _ExtentY        =   503
         MaxLength       =   8
         DataType        =   4
         IdName          =   "WLRCMCDFR"
      End
      Begin MwfCtl.MWCTL_MED ATXT 
         Height          =   285
         Index           =   69
         Left            =   3285
         TabIndex        =   85
         ToolTipText     =   "Enter To Customer Code"
         Top             =   1110
         Width           =   1305
         _ExtentX        =   2302
         _ExtentY        =   503
         MaxLength       =   8
         DataType        =   4
         IdName          =   "WLRCMCDTO"
      End
      Begin MwfCtl.MWCTL_MED ATXT 
         Height          =   285
         Index           =   70
         Left            =   1500
         TabIndex        =   92
         ToolTipText     =   "Enter From Labour Rate Modified Date"
         Top             =   2235
         Width           =   885
         _ExtentX        =   1561
         _ExtentY        =   503
         MaxLength       =   8
         DataType        =   3
         IdName          =   "WLRMODDTFR"
      End
      Begin MwfCtl.MWCTL_MED ATXT 
         Height          =   285
         Index           =   72
         Left            =   1500
         TabIndex        =   82
         ToolTipText     =   "Enter From Customer Category"
         Top             =   810
         Width           =   435
         _ExtentX        =   767
         _ExtentY        =   503
         MaxLength       =   2
         DataType        =   4
         IdName          =   "WLRCMCTGFR"
      End
      Begin MwfCtl.MWCTL_MED ATXT 
         Height          =   285
         Index           =   73
         Left            =   3285
         TabIndex        =   83
         ToolTipText     =   "Enter To Customer Category"
         Top             =   825
         Width           =   435
         _ExtentX        =   767
         _ExtentY        =   503
         MaxLength       =   2
         DataType        =   4
         IdName          =   "WLRCMCTGTO"
      End
      Begin MwfCtl.MWCTL_MED ATXT 
         Height          =   285
         Index           =   74
         Left            =   1500
         TabIndex        =   86
         ToolTipText     =   "Enter From Currency Code"
         Top             =   1380
         Width           =   885
         _ExtentX        =   1561
         _ExtentY        =   503
         MaxLength       =   5
         DataType        =   4
         IdName          =   "WLRCMCURCDFR"
      End
      Begin MwfCtl.MWCTL_MED ATXT 
         Height          =   285
         Index           =   75
         Left            =   3285
         TabIndex        =   87
         ToolTipText     =   "Enter To Currency Code"
         Top             =   1395
         Width           =   885
         _ExtentX        =   1561
         _ExtentY        =   503
         MaxLength       =   5
         DataType        =   4
         IdName          =   "WLRCMCURCDTO"
      End
      Begin MwfCtl.MWCTL_MED ATXT 
         Height          =   285
         Index           =   86
         Left            =   1500
         TabIndex        =   90
         ToolTipText     =   "Enter From Labour Sub Code"
         Top             =   1950
         Width           =   885
         _ExtentX        =   1561
         _ExtentY        =   503
         MaxLength       =   5
         DataType        =   4
         IdName          =   "WLRSCDFR"
      End
      Begin MwfCtl.MWCTL_MED ATXT 
         Height          =   285
         Index           =   87
         Left            =   3285
         TabIndex        =   91
         ToolTipText     =   "Enter To Labour Sub Code"
         Top             =   1965
         Width           =   885
         _ExtentX        =   1561
         _ExtentY        =   503
         MaxLength       =   5
         DataType        =   4
         IdName          =   "WLRSCDTO"
      End
      Begin MwfCtl.MWCTL_MED ATXT 
         Height          =   285
         Index           =   88
         Left            =   1500
         TabIndex        =   88
         ToolTipText     =   "Enter From Labour Main Code"
         Top             =   1665
         Width           =   885
         _ExtentX        =   1561
         _ExtentY        =   503
         MaxLength       =   5
         DataType        =   4
         IdName          =   "WLRMCDFR"
      End
      Begin MwfCtl.MWCTL_MED ATXT 
         Height          =   285
         Index           =   89
         Left            =   3285
         TabIndex        =   89
         ToolTipText     =   "Enter To Labour Main Code"
         Top             =   1680
         Width           =   885
         _ExtentX        =   1561
         _ExtentY        =   503
         MaxLength       =   5
         DataType        =   4
         IdName          =   "WLRMCDTO"
      End
      Begin VB.Label LblMain 
         BackStyle       =   0  'Transparent
         BorderStyle     =   1  'Fixed Single
         Caption         =   "  Labour Rate Transfer"
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
         TabIndex        =   335
         Top             =   90
         Width           =   4905
      End
      Begin VB.Label ALBL 
         BackStyle       =   0  'Transparent
         Caption         =   "Currency Cd"
         BeginProperty Font 
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
         Left            =   90
         TabIndex        =   255
         Top             =   1380
         Width           =   1305
      End
      Begin VB.Label ALBL 
         BackStyle       =   0  'Transparent
         Caption         =   "Customer Ctg"
         BeginProperty Font 
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
         Left            =   90
         TabIndex        =   254
         Top             =   810
         Width           =   1425
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
         Index           =   68
         Left            =   90
         TabIndex        =   253
         Top             =   1095
         Width           =   1335
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
         Index           =   67
         Left            =   3270
         TabIndex        =   252
         Top             =   480
         Width           =   1035
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
         Index           =   66
         Left            =   1500
         TabIndex        =   251
         Top             =   480
         Width           =   1035
      End
      Begin VB.Label ALBL 
         BackStyle       =   0  'Transparent
         Caption         =   "Modified Dt"
         BeginProperty Font 
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
         Left            =   90
         TabIndex        =   250
         Top             =   2235
         Width           =   1245
      End
      Begin VB.Label ALBL 
         BackStyle       =   0  'Transparent
         Caption         =   "Lab Sub Cd"
         BeginProperty Font 
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
         Left            =   90
         TabIndex        =   249
         Top             =   1950
         Width           =   1455
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
         Index           =   58
         Left            =   90
         TabIndex        =   248
         Top             =   1665
         Width           =   1455
      End
   End
   Begin VB.Frame FraLabRtHist 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2970
      Left            =   4860
      TabIndex        =   272
      Top             =   6105
      Visible         =   0   'False
      Width           =   4905
      Begin VB.CommandButton CmdLabRtHistTrfOut 
         Caption         =   "&Transfer Out"
         Height          =   315
         Left            =   3360
         TabIndex        =   130
         ToolTipText     =   "Labour Rate History Transfer Out"
         Top             =   2565
         Width           =   1230
      End
      Begin MwfCtl.MWCTL_MED ATXT 
         Height          =   285
         Index           =   100
         Left            =   1500
         TabIndex        =   121
         ToolTipText     =   "Enter From Customer Code"
         Top             =   1395
         Width           =   1305
         _ExtentX        =   2302
         _ExtentY        =   503
         MaxLength       =   8
         DataType        =   4
         IdName          =   "WLHCMCDFR"
      End
      Begin MwfCtl.MWCTL_MED ATXT 
         Height          =   285
         Index           =   101
         Left            =   3285
         TabIndex        =   122
         ToolTipText     =   "Enter To Customer Code"
         Top             =   1410
         Width           =   1305
         _ExtentX        =   2302
         _ExtentY        =   503
         MaxLength       =   8
         DataType        =   4
         IdName          =   "WLHCMCDTO"
      End
      Begin MwfCtl.MWCTL_MED ATXT 
         Height          =   285
         Index           =   102
         Left            =   1500
         TabIndex        =   129
         ToolTipText     =   "Enter From Labour Rate Modified Date"
         Top             =   2535
         Width           =   885
         _ExtentX        =   1561
         _ExtentY        =   503
         MaxLength       =   8
         DataType        =   3
         IdName          =   "WLHMODDTFR"
      End
      Begin MwfCtl.MWCTL_MED ATXT 
         Height          =   285
         Index           =   103
         Left            =   1500
         TabIndex        =   119
         ToolTipText     =   "Enter From Customer Category"
         Top             =   1110
         Width           =   435
         _ExtentX        =   767
         _ExtentY        =   503
         MaxLength       =   2
         DataType        =   4
         IdName          =   "WLHCMCTGFR"
      End
      Begin MwfCtl.MWCTL_MED ATXT 
         Height          =   285
         Index           =   104
         Left            =   3285
         TabIndex        =   120
         ToolTipText     =   "Enter To Customer Category"
         Top             =   1125
         Width           =   435
         _ExtentX        =   767
         _ExtentY        =   503
         MaxLength       =   2
         DataType        =   4
         IdName          =   "WLHCMCTGTO"
      End
      Begin MwfCtl.MWCTL_MED ATXT 
         Height          =   285
         Index           =   105
         Left            =   1500
         TabIndex        =   123
         ToolTipText     =   "Enter From Currency Code"
         Top             =   1680
         Width           =   885
         _ExtentX        =   1561
         _ExtentY        =   503
         MaxLength       =   5
         DataType        =   4
         IdName          =   "WLHCMCURCDFR"
      End
      Begin MwfCtl.MWCTL_MED ATXT 
         Height          =   285
         Index           =   106
         Left            =   3285
         TabIndex        =   124
         ToolTipText     =   "Enter To Currency Code"
         Top             =   1695
         Width           =   885
         _ExtentX        =   1561
         _ExtentY        =   503
         MaxLength       =   5
         DataType        =   4
         IdName          =   "WLHCMCURCDTO"
      End
      Begin MwfCtl.MWCTL_MED ATXT 
         Height          =   285
         Index           =   107
         Left            =   1500
         TabIndex        =   127
         ToolTipText     =   "Enter From Labour Sub Code"
         Top             =   2250
         Width           =   885
         _ExtentX        =   1561
         _ExtentY        =   503
         MaxLength       =   5
         DataType        =   4
         IdName          =   "WLHSCDFR"
      End
      Begin MwfCtl.MWCTL_MED ATXT 
         Height          =   285
         Index           =   108
         Left            =   3285
         TabIndex        =   128
         ToolTipText     =   "Enter To Labour Sub Code"
         Top             =   2265
         Width           =   885
         _ExtentX        =   1561
         _ExtentY        =   503
         MaxLength       =   5
         DataType        =   4
         IdName          =   "WLHSCDTO"
      End
      Begin MwfCtl.MWCTL_MED ATXT 
         Height          =   285
         Index           =   109
         Left            =   1500
         TabIndex        =   125
         ToolTipText     =   "Enter From Labour Main Code"
         Top             =   1965
         Width           =   885
         _ExtentX        =   1561
         _ExtentY        =   503
         MaxLength       =   5
         DataType        =   4
         IdName          =   "WLHMCDFR"
      End
      Begin MwfCtl.MWCTL_MED ATXT 
         Height          =   285
         Index           =   110
         Left            =   3285
         TabIndex        =   126
         ToolTipText     =   "Enter To Labour Main Code"
         Top             =   1980
         Width           =   885
         _ExtentX        =   1561
         _ExtentY        =   503
         MaxLength       =   5
         DataType        =   4
         IdName          =   "WLHMCDTO"
      End
      Begin MwfCtl.MWCTL_MED ATXT 
         Height          =   285
         Index           =   111
         Left            =   1500
         TabIndex        =   118
         ToolTipText     =   "Enter Year And Month in 'YyyyMm' Format"
         Top             =   825
         Width           =   1320
         _ExtentX        =   2328
         _ExtentY        =   503
         Alignment       =   2
         Mask            =   "#####0"
         MaxLength       =   6
         DataType        =   1
         IdName          =   "WLHYYYYMM"
      End
      Begin VB.Label LblMain 
         BackStyle       =   0  'Transparent
         BorderStyle     =   1  'Fixed Single
         Caption         =   "  Labour Rate History Transfer"
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
         TabIndex        =   334
         Top             =   90
         Width           =   4905
      End
      Begin VB.Label ALBL 
         BackStyle       =   0  'Transparent
         Caption         =   "YyyyMm"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Index           =   87
         Left            =   90
         TabIndex        =   281
         Top             =   710
         Width           =   975
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
         Index           =   86
         Left            =   90
         TabIndex        =   280
         Top             =   1965
         Width           =   1455
      End
      Begin VB.Label ALBL 
         BackStyle       =   0  'Transparent
         Caption         =   "Lab Sub Cd"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Index           =   85
         Left            =   90
         TabIndex        =   279
         Top             =   2250
         Width           =   1455
      End
      Begin VB.Label ALBL 
         BackStyle       =   0  'Transparent
         Caption         =   "Modified Dt"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Index           =   84
         Left            =   90
         TabIndex        =   278
         Top             =   2415
         Width           =   1245
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
         Index           =   83
         Left            =   1500
         TabIndex        =   277
         Top             =   480
         Width           =   1035
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
         Index           =   82
         Left            =   3270
         TabIndex        =   276
         Top             =   480
         Width           =   1035
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
         Index           =   81
         Left            =   90
         TabIndex        =   275
         Top             =   1395
         Width           =   1335
      End
      Begin VB.Label ALBL 
         BackStyle       =   0  'Transparent
         Caption         =   "Customer Ctg"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Index           =   80
         Left            =   90
         TabIndex        =   274
         Top             =   1110
         Width           =   1425
      End
      Begin VB.Label ALBL 
         BackStyle       =   0  'Transparent
         Caption         =   "Currency Cd"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Index           =   79
         Left            =   90
         TabIndex        =   273
         Top             =   1680
         Width           =   1305
      End
   End
   Begin VB.Frame FraDsg 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2400
      Left            =   3840
      TabIndex        =   223
      Top             =   6675
      Visible         =   0   'False
      Width           =   6705
      Begin VB.CommandButton CmdDsgTrfOut 
         Caption         =   "&Transfer Out"
         Height          =   315
         Left            =   5190
         TabIndex        =   224
         ToolTipText     =   "Design Transfer Out"
         Top             =   1995
         Width           =   1230
      End
      Begin MwfCtl.MWCTL_MED ATXT 
         Height          =   285
         Index           =   44
         Left            =   1380
         TabIndex        =   164
         ToolTipText     =   "Enter From Design Code"
         Top             =   1140
         Width           =   2355
         _ExtentX        =   4154
         _ExtentY        =   503
         MaxLength       =   15
         DataType        =   4
         IdName          =   "WDMCDFR"
      End
      Begin MwfCtl.MWCTL_MED ATXT 
         Height          =   285
         Index           =   45
         Left            =   4065
         TabIndex        =   165
         ToolTipText     =   "Enter To Design Code"
         Top             =   1140
         Width           =   2355
         _ExtentX        =   4154
         _ExtentY        =   503
         MaxLength       =   15
         DataType        =   4
         IdName          =   "WDMCDTO"
      End
      Begin MwfCtl.MWCTL_MED ATXT 
         Height          =   285
         Index           =   46
         Left            =   1380
         TabIndex        =   170
         ToolTipText     =   "Enter From Design Modified Date"
         Top             =   1995
         Width           =   885
         _ExtentX        =   1561
         _ExtentY        =   503
         MaxLength       =   8
         DataType        =   3
         IdName          =   "WDMMODDTFR"
      End
      Begin MwfCtl.MWCTL_MED ATXT 
         Height          =   285
         Index           =   36
         Left            =   1380
         TabIndex        =   163
         ToolTipText     =   "Enter Design TcType"
         Top             =   855
         Width           =   555
         _ExtentX        =   979
         _ExtentY        =   503
         MaxLength       =   3
         DataType        =   4
         IdName          =   "WDMTCTYP"
      End
      Begin MwfCtl.MWCTL_MED ATXT 
         Height          =   285
         Index           =   37
         Left            =   1380
         TabIndex        =   166
         ToolTipText     =   "Enter From Design Size"
         Top             =   1425
         Width           =   885
         _ExtentX        =   1561
         _ExtentY        =   503
         MaxLength       =   5
         DataType        =   4
         IdName          =   "WDMSZFR"
      End
      Begin MwfCtl.MWCTL_MED ATXT 
         Height          =   285
         Index           =   38
         Left            =   4065
         TabIndex        =   167
         ToolTipText     =   "Enter To Design Size"
         Top             =   1425
         Width           =   885
         _ExtentX        =   1561
         _ExtentY        =   503
         MaxLength       =   5
         DataType        =   4
         IdName          =   "WDMSZTO"
      End
      Begin MwfCtl.MWCTL_MED ATXT 
         Height          =   285
         Index           =   39
         Left            =   1380
         TabIndex        =   168
         ToolTipText     =   "Enter From Design Category"
         Top             =   1710
         Width           =   885
         _ExtentX        =   1561
         _ExtentY        =   503
         MaxLength       =   5
         DataType        =   4
         IdName          =   "WDMCTGFR"
         Data            =   "ABCDW"
      End
      Begin MwfCtl.MWCTL_MED ATXT 
         Height          =   285
         Index           =   40
         Left            =   4065
         TabIndex        =   169
         ToolTipText     =   "Enter To Design Category"
         Top             =   1710
         Width           =   885
         _ExtentX        =   1561
         _ExtentY        =   503
         MaxLength       =   5
         DataType        =   4
         IdName          =   "WDMCTGTO"
      End
      Begin VB.Label LblMain 
         BackStyle       =   0  'Transparent
         BorderStyle     =   1  'Fixed Single
         Caption         =   "  Design Transfer"
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
         Index           =   13
         Left            =   -30
         TabIndex        =   338
         Top             =   90
         Width           =   6705
      End
      Begin VB.Label ALBL 
         BackStyle       =   0  'Transparent
         Caption         =   "Design Ctg"
         BeginProperty Font 
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
         Left            =   90
         TabIndex        =   231
         Top             =   1710
         Width           =   1305
      End
      Begin VB.Label ALBL 
         BackStyle       =   0  'Transparent
         Caption         =   "Design Size"
         BeginProperty Font 
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
         TabIndex        =   230
         Top             =   1425
         Width           =   1305
      End
      Begin VB.Label ALBL 
         BackStyle       =   0  'Transparent
         Caption         =   "Tc Type"
         BeginProperty Font 
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
         TabIndex        =   229
         Top             =   855
         Width           =   1245
      End
      Begin VB.Label ALBL 
         BackStyle       =   0  'Transparent
         Caption         =   "Design Code"
         BeginProperty Font 
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
         Left            =   90
         TabIndex        =   228
         Top             =   1140
         Width           =   1245
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
         Index           =   44
         Left            =   4050
         TabIndex        =   227
         Top             =   480
         Width           =   1035
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
         Index           =   43
         Left            =   1380
         TabIndex        =   226
         Top             =   480
         Width           =   1035
      End
      Begin VB.Label ALBL 
         BackStyle       =   0  'Transparent
         Caption         =   "Modified Dt"
         BeginProperty Font 
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
         Left            =   90
         TabIndex        =   225
         Top             =   1995
         Width           =   1245
      End
   End
   Begin VB.Frame FraPrm 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2520
      Left            =   4830
      TabIndex        =   300
      Top             =   6555
      Visible         =   0   'False
      Width           =   4665
      Begin VB.CommandButton CmdPrmTrfOut 
         Caption         =   "&Transfer Out"
         Height          =   315
         Left            =   3180
         TabIndex        =   148
         ToolTipText     =   "Param Transfer Out"
         Top             =   2010
         Width           =   1230
      End
      Begin MwfCtl.MWCTL_MED ATXT 
         Height          =   285
         Index           =   124
         Left            =   1410
         TabIndex        =   143
         ToolTipText     =   "Enter From Main Code"
         Top             =   1440
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   503
         MaxLength       =   8
         DataType        =   4
         IdName          =   "WPMCDFR"
      End
      Begin MwfCtl.MWCTL_MED ATXT 
         Height          =   285
         Index           =   125
         Left            =   3075
         TabIndex        =   144
         ToolTipText     =   "Enter To Main Code"
         Top             =   1440
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   503
         MaxLength       =   8
         DataType        =   4
         IdName          =   "WPMCDTO"
      End
      Begin MwfCtl.MWCTL_MED ATXT 
         Height          =   285
         Index           =   126
         Left            =   1410
         TabIndex        =   142
         ToolTipText     =   "Enter Parameter Type"
         Top             =   1155
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   503
         MaxLength       =   8
         DataType        =   4
         IdName          =   "WPTYP"
      End
      Begin MwfCtl.MWCTL_MED ATXT 
         Height          =   285
         Index           =   127
         Left            =   1410
         TabIndex        =   147
         ToolTipText     =   "Enter From Param Modified Date"
         Top             =   2010
         Width           =   885
         _ExtentX        =   1561
         _ExtentY        =   503
         MaxLength       =   8
         DataType        =   3
         IdName          =   "WPMODDTFR"
      End
      Begin MwfCtl.MWCTL_MED ATXT 
         Height          =   285
         Index           =   129
         Left            =   1410
         TabIndex        =   145
         ToolTipText     =   "Enter From Sub Code"
         Top             =   1725
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   503
         MaxLength       =   8
         DataType        =   4
         IdName          =   "WPSCDFR"
      End
      Begin MwfCtl.MWCTL_MED ATXT 
         Height          =   285
         Index           =   130
         Left            =   3075
         TabIndex        =   146
         ToolTipText     =   "Enter To Sub Code"
         Top             =   1725
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   503
         MaxLength       =   8
         DataType        =   4
         IdName          =   "WPSCDTO"
      End
      Begin MwfCtl.MWCTL_MED ATXT 
         Height          =   285
         Index           =   137
         Left            =   1410
         TabIndex        =   141
         ToolTipText     =   "Specify Whether The Param To Be Transferred Is A System/ Other Param ('Y' For Sys, 'N' For Oth)"
         Top             =   870
         Width           =   315
         _ExtentX        =   556
         _ExtentY        =   503
         DataType        =   4
         IdName          =   "WPCTG"
      End
      Begin VB.Label LblMain 
         BackStyle       =   0  'Transparent
         BorderStyle     =   1  'Fixed Single
         Caption         =   "  Param Transfer"
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
         TabIndex        =   330
         Top             =   90
         Width           =   4665
      End
      Begin VB.Label ALBL 
         BackStyle       =   0  'Transparent
         Caption         =   "Main Code"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Index           =   117
         Left            =   90
         TabIndex        =   315
         Top             =   1440
         Width           =   1305
      End
      Begin VB.Label ALBL 
         BackStyle       =   0  'Transparent
         Caption         =   "Sys/ Oth"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Index           =   116
         Left            =   90
         TabIndex        =   314
         Top             =   870
         Width           =   1305
      End
      Begin VB.Label ALBL 
         BackStyle       =   0  'Transparent
         Caption         =   "Sub Code"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Index           =   109
         Left            =   90
         TabIndex        =   305
         Top             =   1725
         Width           =   1245
      End
      Begin VB.Label ALBL 
         BackStyle       =   0  'Transparent
         Caption         =   "Param Type"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Index           =   107
         Left            =   90
         TabIndex        =   304
         Top             =   1155
         Width           =   1305
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
         Index           =   105
         Left            =   3060
         TabIndex        =   303
         Top             =   480
         Width           =   1035
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
         Index           =   104
         Left            =   1410
         TabIndex        =   302
         Top             =   480
         Width           =   1035
      End
      Begin VB.Label ALBL 
         BackStyle       =   0  'Transparent
         Caption         =   "Modified Dt"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Index           =   103
         Left            =   90
         TabIndex        =   301
         Top             =   2010
         Width           =   1245
      End
   End
   Begin VB.Frame FraRm 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2520
      Left            =   4200
      TabIndex        =   215
      Top             =   6555
      Visible         =   0   'False
      Width           =   5835
      Begin VB.CommandButton CmdRmTrfOut 
         Caption         =   "&Transfer Out"
         Height          =   315
         Left            =   4350
         TabIndex        =   58
         ToolTipText     =   "Rm Transfer Out"
         Top             =   2025
         Width           =   1230
      End
      Begin MwfCtl.MWCTL_MED ATXT 
         Height          =   285
         Index           =   23
         Left            =   1350
         TabIndex        =   51
         ToolTipText     =   "Enter From Raw Material Category"
         Top             =   1155
         Width           =   495
         _ExtentX        =   873
         _ExtentY        =   503
         MaxLength       =   2
         DataType        =   4
         IdName          =   "WRMCTGFR"
      End
      Begin MwfCtl.MWCTL_MED ATXT 
         Height          =   285
         Index           =   25
         Left            =   3645
         TabIndex        =   52
         ToolTipText     =   "Enter To Raw Material Category"
         Top             =   1155
         Width           =   495
         _ExtentX        =   873
         _ExtentY        =   503
         MaxLength       =   2
         DataType        =   4
         IdName          =   "WRMCTGTO"
      End
      Begin MwfCtl.MWCTL_MED ATXT 
         Height          =   285
         Index           =   26
         Left            =   1350
         TabIndex        =   50
         ToolTipText     =   "Transfer Customs Cd 'Y' Or Normal Cd 'N' "
         Top             =   870
         Width           =   315
         _ExtentX        =   556
         _ExtentY        =   503
         DataType        =   4
         IdName          =   "WRMZ"
      End
      Begin MwfCtl.MWCTL_MED ATXT 
         Height          =   285
         Index           =   27
         Left            =   1350
         TabIndex        =   57
         ToolTipText     =   "Enter From Raw Material Modified Date"
         Top             =   2010
         Width           =   885
         _ExtentX        =   1561
         _ExtentY        =   503
         MaxLength       =   8
         DataType        =   3
         IdName          =   "WRMMODDTFR"
      End
      Begin MwfCtl.MWCTL_MED ATXT 
         Height          =   285
         Index           =   28
         Left            =   1350
         TabIndex        =   53
         ToolTipText     =   "Enter From Raw Material Sub Category"
         Top             =   1440
         Width           =   885
         _ExtentX        =   1561
         _ExtentY        =   503
         MaxLength       =   5
         DataType        =   4
         IdName          =   "WRMSCTGFR"
      End
      Begin MwfCtl.MWCTL_MED ATXT 
         Height          =   285
         Index           =   29
         Left            =   3645
         TabIndex        =   54
         ToolTipText     =   "Enter To Raw Material Sub Category"
         Top             =   1440
         Width           =   885
         _ExtentX        =   1561
         _ExtentY        =   503
         MaxLength       =   5
         DataType        =   4
         IdName          =   "WRMSCTGTO"
      End
      Begin MwfCtl.MWCTL_MED ATXT 
         Height          =   285
         Index           =   30
         Left            =   1350
         TabIndex        =   55
         ToolTipText     =   "Enter From Raw Material Code"
         Top             =   1725
         Width           =   1935
         _ExtentX        =   3413
         _ExtentY        =   503
         MaxLength       =   12
         DataType        =   4
         IdName          =   "WRMCDFR"
      End
      Begin MwfCtl.MWCTL_MED ATXT 
         Height          =   285
         Index           =   31
         Left            =   3645
         TabIndex        =   56
         ToolTipText     =   "Enter To Raw Material Code"
         Top             =   1725
         Width           =   1935
         _ExtentX        =   3413
         _ExtentY        =   503
         MaxLength       =   12
         DataType        =   4
         IdName          =   "WRMCDTO"
      End
      Begin VB.Label LblMain 
         BackStyle       =   0  'Transparent
         BorderStyle     =   1  'Fixed Single
         Caption         =   "  Raw Material Transfer"
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
         TabIndex        =   328
         Top             =   90
         Width           =   5835
      End
      Begin VB.Label ALBL 
         BackStyle       =   0  'Transparent
         Caption         =   "Rm Code"
         BeginProperty Font 
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
         Left            =   90
         TabIndex        =   222
         Top             =   1725
         Width           =   1245
      End
      Begin VB.Label ALBL 
         BackStyle       =   0  'Transparent
         Caption         =   "Rm Sub Ctg"
         BeginProperty Font 
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
         TabIndex        =   221
         Top             =   1440
         Width           =   1245
      End
      Begin VB.Label ALBL 
         BackStyle       =   0  'Transparent
         Caption         =   "Modified Dt"
         BeginProperty Font 
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
         Left            =   90
         TabIndex        =   220
         Top             =   2010
         Width           =   1245
      End
      Begin VB.Label ALBL 
         BackStyle       =   0  'Transparent
         Caption         =   "Customs Cd"
         BeginProperty Font 
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
         TabIndex        =   219
         Top             =   870
         Width           =   1245
      End
      Begin VB.Label ALBL 
         BackStyle       =   0  'Transparent
         Caption         =   "Rm Ctg"
         BeginProperty Font 
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
         Left            =   90
         TabIndex        =   218
         Top             =   1155
         Width           =   1245
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
         Index           =   29
         Left            =   3630
         TabIndex        =   217
         Top             =   480
         Width           =   1035
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
         Index           =   28
         Left            =   1350
         TabIndex        =   216
         Top             =   480
         Width           =   1035
      End
   End
   Begin VB.Frame FraVPrm 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2520
      Left            =   4860
      TabIndex        =   306
      Top             =   6555
      Visible         =   0   'False
      Width           =   4665
      Begin VB.CommandButton CmdVPrmTrfOut 
         Caption         =   "&Transfer Out"
         Height          =   315
         Left            =   3180
         TabIndex        =   156
         ToolTipText     =   "Variable Param Transfer Out"
         Top             =   2010
         Width           =   1230
      End
      Begin MwfCtl.MWCTL_MED ATXT 
         Height          =   285
         Index           =   128
         Left            =   1410
         TabIndex        =   151
         ToolTipText     =   "Enter From Variable Main Code"
         Top             =   1440
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   503
         MaxLength       =   8
         DataType        =   4
         IdName          =   "WVPMCDFR"
      End
      Begin MwfCtl.MWCTL_MED ATXT 
         Height          =   285
         Index           =   131
         Left            =   3075
         TabIndex        =   152
         ToolTipText     =   "Enter To Variable Main Code"
         Top             =   1440
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   503
         MaxLength       =   8
         DataType        =   4
         IdName          =   "WVPMCDTO"
      End
      Begin MwfCtl.MWCTL_MED ATXT 
         Height          =   285
         Index           =   132
         Left            =   1410
         TabIndex        =   150
         ToolTipText     =   "Enter Variable Parameter Type (Other Param)"
         Top             =   1155
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   503
         MaxLength       =   8
         DataType        =   4
         IdName          =   "WVPTYP"
      End
      Begin MwfCtl.MWCTL_MED ATXT 
         Height          =   285
         Index           =   133
         Left            =   1410
         TabIndex        =   155
         ToolTipText     =   "Enter From Variable Param Modified Date"
         Top             =   2010
         Width           =   885
         _ExtentX        =   1561
         _ExtentY        =   503
         MaxLength       =   8
         DataType        =   3
         IdName          =   "WVPMODDTFR"
      End
      Begin MwfCtl.MWCTL_MED ATXT 
         Height          =   285
         Index           =   134
         Left            =   1410
         TabIndex        =   153
         ToolTipText     =   "Enter From Variable Sub Code"
         Top             =   1725
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   503
         MaxLength       =   8
         DataType        =   4
         IdName          =   "WVPSCDFR"
      End
      Begin MwfCtl.MWCTL_MED ATXT 
         Height          =   285
         Index           =   135
         Left            =   3075
         TabIndex        =   154
         ToolTipText     =   "Enter To Variable Sub Code"
         Top             =   1725
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   503
         MaxLength       =   8
         DataType        =   4
         IdName          =   "WVPSCDTO"
      End
      Begin MwfCtl.MWCTL_MED ATXT 
         Height          =   285
         Index           =   136
         Left            =   1410
         TabIndex        =   149
         ToolTipText     =   "Specify Whether The Param To Be Transferred Is A System/ Other Param ('Y' For Sys, 'N' For Oth)"
         Top             =   870
         Width           =   315
         _ExtentX        =   556
         _ExtentY        =   503
         DataType        =   4
         IdName          =   "WVPCTG"
      End
      Begin VB.Label LblMain 
         BackStyle       =   0  'Transparent
         BorderStyle     =   1  'Fixed Single
         Caption         =   "  Variable Param Transfer"
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
         TabIndex        =   329
         Top             =   90
         Width           =   4665
      End
      Begin VB.Label ALBL 
         BackStyle       =   0  'Transparent
         Caption         =   "Sys/ Oth"
         BeginProperty Font 
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
         Left            =   90
         TabIndex        =   313
         Top             =   870
         Width           =   1305
      End
      Begin VB.Label ALBL 
         BackStyle       =   0  'Transparent
         Caption         =   "Modified Dt"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Index           =   114
         Left            =   90
         TabIndex        =   312
         Top             =   2010
         Width           =   1245
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
         Index           =   113
         Left            =   1410
         TabIndex        =   311
         Top             =   480
         Width           =   1035
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
         Index           =   112
         Left            =   3060
         TabIndex        =   310
         Top             =   480
         Width           =   1035
      End
      Begin VB.Label ALBL 
         BackStyle       =   0  'Transparent
         Caption         =   "Main Code"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Index           =   111
         Left            =   90
         TabIndex        =   309
         Top             =   1440
         Width           =   1245
      End
      Begin VB.Label ALBL 
         BackStyle       =   0  'Transparent
         Caption         =   "Param Type"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Index           =   110
         Left            =   90
         TabIndex        =   308
         Top             =   1155
         Width           =   1305
      End
      Begin VB.Label ALBL 
         BackStyle       =   0  'Transparent
         Caption         =   "Sub Code"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Index           =   108
         Left            =   90
         TabIndex        =   307
         Top             =   1725
         Width           =   1245
      End
   End
   Begin VB.Frame FraRmRt 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   4530
      Left            =   4110
      TabIndex        =   232
      Top             =   4545
      Visible         =   0   'False
      Width           =   6045
      Begin VB.CommandButton CmdRmRtTrfOut 
         Caption         =   "&Transfer Out"
         Height          =   315
         Left            =   4530
         TabIndex        =   81
         ToolTipText     =   "Rm Rate Transfer Out"
         Top             =   3990
         Width           =   1230
      End
      Begin MwfCtl.MWCTL_MED ATXT 
         Height          =   285
         Index           =   41
         Left            =   1500
         TabIndex        =   62
         ToolTipText     =   "Enter From Customer Code"
         Top             =   1425
         Width           =   1305
         _ExtentX        =   2302
         _ExtentY        =   503
         MaxLength       =   8
         DataType        =   4
         IdName          =   "WRRCMCDFR"
      End
      Begin MwfCtl.MWCTL_MED ATXT 
         Height          =   285
         Index           =   42
         Left            =   3825
         TabIndex        =   63
         ToolTipText     =   "Enter To Customer Code"
         Top             =   1440
         Width           =   1305
         _ExtentX        =   2302
         _ExtentY        =   503
         MaxLength       =   8
         DataType        =   4
         IdName          =   "WRRCMCDTO"
      End
      Begin MwfCtl.MWCTL_MED ATXT 
         Height          =   285
         Index           =   43
         Left            =   1500
         TabIndex        =   80
         ToolTipText     =   "Enter From Rm Rate Modified Date"
         Top             =   3990
         Width           =   885
         _ExtentX        =   1561
         _ExtentY        =   503
         MaxLength       =   8
         DataType        =   3
         IdName          =   "WRRMODDTFR"
      End
      Begin MwfCtl.MWCTL_MED ATXT 
         Height          =   285
         Index           =   47
         Left            =   1500
         TabIndex        =   59
         ToolTipText     =   "Enter Rm Rate Tc Type"
         Top             =   855
         Width           =   555
         _ExtentX        =   979
         _ExtentY        =   503
         MaxLength       =   3
         DataType        =   4
         IdName          =   "WRRTCTYP"
      End
      Begin MwfCtl.MWCTL_MED ATXT 
         Height          =   285
         Index           =   48
         Left            =   1500
         TabIndex        =   60
         ToolTipText     =   "Enter From Customer Category"
         Top             =   1140
         Width           =   435
         _ExtentX        =   767
         _ExtentY        =   503
         MaxLength       =   2
         DataType        =   4
         IdName          =   "WRRCMCTGFR"
      End
      Begin MwfCtl.MWCTL_MED ATXT 
         Height          =   285
         Index           =   49
         Left            =   3825
         TabIndex        =   61
         ToolTipText     =   "Enter To Customer Category"
         Top             =   1155
         Width           =   435
         _ExtentX        =   767
         _ExtentY        =   503
         MaxLength       =   2
         DataType        =   4
         IdName          =   "WRRCMCTGTO"
      End
      Begin MwfCtl.MWCTL_MED ATXT 
         Height          =   285
         Index           =   50
         Left            =   1500
         TabIndex        =   64
         ToolTipText     =   "Enter From Currency Code"
         Top             =   1710
         Width           =   885
         _ExtentX        =   1561
         _ExtentY        =   503
         MaxLength       =   5
         DataType        =   4
         IdName          =   "WRRCMCURCDFR"
      End
      Begin MwfCtl.MWCTL_MED ATXT 
         Height          =   285
         Index           =   51
         Left            =   3825
         TabIndex        =   65
         ToolTipText     =   "Enter To Currency Code"
         Top             =   1725
         Width           =   885
         _ExtentX        =   1561
         _ExtentY        =   503
         MaxLength       =   5
         DataType        =   4
         IdName          =   "WRRCMCURCDTO"
      End
      Begin MwfCtl.MWCTL_MED ATXT 
         Height          =   285
         Index           =   52
         Left            =   1500
         TabIndex        =   66
         ToolTipText     =   "Enter From Raw Material Category"
         Top             =   1995
         Width           =   495
         _ExtentX        =   873
         _ExtentY        =   503
         MaxLength       =   2
         DataType        =   4
         IdName          =   "WRRCTGFR"
      End
      Begin MwfCtl.MWCTL_MED ATXT 
         Height          =   285
         Index           =   53
         Left            =   3825
         TabIndex        =   67
         ToolTipText     =   "Enter To Raw Material Category"
         Top             =   2010
         Width           =   495
         _ExtentX        =   873
         _ExtentY        =   503
         MaxLength       =   2
         DataType        =   4
         IdName          =   "WRRCTGTO"
      End
      Begin MwfCtl.MWCTL_MED ATXT 
         Height          =   285
         Index           =   54
         Left            =   1500
         TabIndex        =   68
         ToolTipText     =   "Enter From Raw Material Sub Category"
         Top             =   2280
         Width           =   885
         _ExtentX        =   1561
         _ExtentY        =   503
         MaxLength       =   5
         DataType        =   4
         IdName          =   "WRRSCTGFR"
      End
      Begin MwfCtl.MWCTL_MED ATXT 
         Height          =   285
         Index           =   55
         Left            =   3825
         TabIndex        =   69
         ToolTipText     =   "Enter To Raw Material Sub Category"
         Top             =   2295
         Width           =   885
         _ExtentX        =   1561
         _ExtentY        =   503
         MaxLength       =   5
         DataType        =   4
         IdName          =   "WRRSCTGTO"
      End
      Begin MwfCtl.MWCTL_MED ATXT 
         Height          =   285
         Index           =   56
         Left            =   1500
         TabIndex        =   70
         ToolTipText     =   "Enter From Raw Material Code"
         Top             =   2565
         Width           =   1935
         _ExtentX        =   3413
         _ExtentY        =   503
         MaxLength       =   12
         DataType        =   4
         IdName          =   "WRRCDFR"
      End
      Begin MwfCtl.MWCTL_MED ATXT 
         Height          =   285
         Index           =   57
         Left            =   3825
         TabIndex        =   71
         ToolTipText     =   "Enter To Raw Material Code"
         Top             =   2580
         Width           =   1935
         _ExtentX        =   3413
         _ExtentY        =   503
         MaxLength       =   12
         DataType        =   4
         IdName          =   "WRRCDTO"
      End
      Begin MwfCtl.MWCTL_MED ATXT 
         Height          =   285
         Index           =   58
         Left            =   1500
         TabIndex        =   74
         ToolTipText     =   "Enter From Design Category"
         Top             =   3135
         Width           =   885
         _ExtentX        =   1561
         _ExtentY        =   503
         MaxLength       =   5
         DataType        =   4
         IdName          =   "WRRDMCTGFR"
      End
      Begin MwfCtl.MWCTL_MED ATXT 
         Height          =   285
         Index           =   59
         Left            =   3825
         TabIndex        =   75
         ToolTipText     =   "Enter To Design Category"
         Top             =   3150
         Width           =   885
         _ExtentX        =   1561
         _ExtentY        =   503
         MaxLength       =   5
         DataType        =   4
         IdName          =   "WRRDMCTGTO"
      End
      Begin MwfCtl.MWCTL_MED ATXT 
         Height          =   285
         Index           =   60
         Left            =   1500
         TabIndex        =   72
         ToolTipText     =   "Enter From Size To Weight Group"
         Top             =   2850
         Width           =   885
         _ExtentX        =   1561
         _ExtentY        =   503
         MaxLength       =   5
         DataType        =   4
         IdName          =   "WRRSTWGRPFR"
      End
      Begin MwfCtl.MWCTL_MED ATXT 
         Height          =   285
         Index           =   61
         Left            =   3825
         TabIndex        =   73
         ToolTipText     =   "Enter To Size To Weight Group"
         Top             =   2865
         Width           =   885
         _ExtentX        =   1561
         _ExtentY        =   503
         MaxLength       =   5
         DataType        =   4
         IdName          =   "WRRSTWGRPTO"
      End
      Begin MwfCtl.MWCTL_MED ATXT 
         Height          =   285
         Index           =   62
         Left            =   1500
         TabIndex        =   76
         ToolTipText     =   "Enter From Design Loss Category"
         Top             =   3420
         Width           =   885
         _ExtentX        =   1561
         _ExtentY        =   503
         MaxLength       =   5
         DataType        =   4
         IdName          =   "WRRLSCTGFR"
      End
      Begin MwfCtl.MWCTL_MED ATXT 
         Height          =   285
         Index           =   63
         Left            =   3825
         TabIndex        =   77
         ToolTipText     =   "Enter To Design Loss Category"
         Top             =   3435
         Width           =   885
         _ExtentX        =   1561
         _ExtentY        =   503
         MaxLength       =   5
         DataType        =   4
         IdName          =   "WRRLSCTGTO"
      End
      Begin MwfCtl.MWCTL_MED ATXT 
         Height          =   285
         Index           =   64
         Left            =   1500
         TabIndex        =   78
         ToolTipText     =   "Enter From Labour Main Code"
         Top             =   3705
         Width           =   885
         _ExtentX        =   1561
         _ExtentY        =   503
         MaxLength       =   5
         DataType        =   4
         IdName          =   "WRRLABMCDFR"
      End
      Begin MwfCtl.MWCTL_MED ATXT 
         Height          =   285
         Index           =   65
         Left            =   3825
         TabIndex        =   79
         ToolTipText     =   "Enter To Labour Main Code"
         Top             =   3720
         Width           =   885
         _ExtentX        =   1561
         _ExtentY        =   503
         MaxLength       =   5
         DataType        =   4
         IdName          =   "WRRLABMCDTO"
      End
      Begin VB.Label LblMain 
         BackStyle       =   0  'Transparent
         BorderStyle     =   1  'Fixed Single
         Caption         =   "  Rm Rate Transfer"
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
         TabIndex        =   327
         Top             =   90
         Width           =   6045
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
         Index           =   55
         Left            =   90
         TabIndex        =   246
         Top             =   3705
         Width           =   1455
      End
      Begin VB.Label ALBL 
         BackStyle       =   0  'Transparent
         Caption         =   "Dsg Loss Ctg"
         BeginProperty Font 
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
         Left            =   90
         TabIndex        =   245
         Top             =   3420
         Width           =   1455
      End
      Begin VB.Label ALBL 
         BackStyle       =   0  'Transparent
         Caption         =   "SzToWt Grp"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Index           =   53
         Left            =   90
         TabIndex        =   244
         Top             =   2850
         Width           =   1305
      End
      Begin VB.Label ALBL 
         BackStyle       =   0  'Transparent
         Caption         =   "Design Ctg"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Index           =   52
         Left            =   90
         TabIndex        =   243
         Top             =   3135
         Width           =   1305
      End
      Begin VB.Label ALBL 
         BackStyle       =   0  'Transparent
         Caption         =   "Rm Ctg"
         BeginProperty Font 
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
         TabIndex        =   242
         Top             =   1995
         Width           =   1245
      End
      Begin VB.Label ALBL 
         BackStyle       =   0  'Transparent
         Caption         =   "Rm Sub Ctg"
         BeginProperty Font 
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
         Left            =   90
         TabIndex        =   241
         Top             =   2280
         Width           =   1245
      End
      Begin VB.Label ALBL 
         BackStyle       =   0  'Transparent
         Caption         =   "Rm Code"
         BeginProperty Font 
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
         Left            =   90
         TabIndex        =   240
         Top             =   2565
         Width           =   1245
      End
      Begin VB.Label ALBL 
         BackStyle       =   0  'Transparent
         Caption         =   "Modified Dt"
         BeginProperty Font 
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
         Left            =   90
         TabIndex        =   239
         Top             =   3990
         Width           =   1245
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
         Index           =   47
         Left            =   1500
         TabIndex        =   238
         Top             =   480
         Width           =   1035
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
         Index           =   46
         Left            =   3810
         TabIndex        =   237
         Top             =   480
         Width           =   1035
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
         Index           =   42
         Left            =   90
         TabIndex        =   236
         Top             =   1425
         Width           =   1335
      End
      Begin VB.Label ALBL 
         BackStyle       =   0  'Transparent
         Caption         =   "Tc Type"
         BeginProperty Font 
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
         Left            =   90
         TabIndex        =   235
         Top             =   855
         Width           =   1245
      End
      Begin VB.Label ALBL 
         BackStyle       =   0  'Transparent
         Caption         =   "Customer Ctg"
         BeginProperty Font 
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
         Left            =   90
         TabIndex        =   234
         Top             =   1140
         Width           =   1425
      End
      Begin VB.Label ALBL 
         BackStyle       =   0  'Transparent
         Caption         =   "Currency Cd"
         BeginProperty Font 
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
         Left            =   90
         TabIndex        =   233
         Top             =   1710
         Width           =   1305
      End
   End
   Begin VB.Frame FraRmRtHist 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   4830
      Left            =   4110
      TabIndex        =   256
      Top             =   4245
      Visible         =   0   'False
      Width           =   6045
      Begin VB.CommandButton CmdRmRtHistTrfOut 
         Caption         =   "&Transfer Out"
         Height          =   315
         Left            =   4530
         TabIndex        =   117
         ToolTipText     =   "Rm Rate History Transfer Out"
         Top             =   4290
         Width           =   1230
      End
      Begin MwfCtl.MWCTL_MED ATXT 
         Height          =   285
         Index           =   66
         Left            =   1500
         TabIndex        =   98
         ToolTipText     =   "Enter From Customer Code"
         Top             =   1710
         Width           =   1305
         _ExtentX        =   2302
         _ExtentY        =   503
         MaxLength       =   8
         DataType        =   4
         IdName          =   "WRHCMCDFR"
      End
      Begin MwfCtl.MWCTL_MED ATXT 
         Height          =   285
         Index           =   67
         Left            =   3825
         TabIndex        =   99
         ToolTipText     =   "Enter To Customer Code"
         Top             =   1710
         Width           =   1305
         _ExtentX        =   2302
         _ExtentY        =   503
         MaxLength       =   8
         DataType        =   4
         IdName          =   "WRHCMCDTO"
      End
      Begin MwfCtl.MWCTL_MED ATXT 
         Height          =   285
         Index           =   71
         Left            =   1500
         TabIndex        =   116
         ToolTipText     =   "Enter From Rm Rate Modified Date"
         Top             =   4275
         Width           =   885
         _ExtentX        =   1561
         _ExtentY        =   503
         MaxLength       =   8
         DataType        =   3
         IdName          =   "WRHMODDTFR"
      End
      Begin MwfCtl.MWCTL_MED ATXT 
         Height          =   285
         Index           =   76
         Left            =   1500
         TabIndex        =   94
         ToolTipText     =   "Enter Rm Rate Tc Type"
         Top             =   855
         Width           =   555
         _ExtentX        =   979
         _ExtentY        =   503
         MaxLength       =   3
         DataType        =   4
         IdName          =   "WRHTCTYP"
      End
      Begin MwfCtl.MWCTL_MED ATXT 
         Height          =   285
         Index           =   77
         Left            =   1500
         TabIndex        =   96
         ToolTipText     =   "Enter From Customer Category"
         Top             =   1425
         Width           =   435
         _ExtentX        =   767
         _ExtentY        =   503
         MaxLength       =   2
         DataType        =   4
         IdName          =   "WRHCMCTGFR"
      End
      Begin MwfCtl.MWCTL_MED ATXT 
         Height          =   285
         Index           =   78
         Left            =   3825
         TabIndex        =   97
         ToolTipText     =   "Enter To Customer Category"
         Top             =   1425
         Width           =   435
         _ExtentX        =   767
         _ExtentY        =   503
         MaxLength       =   2
         DataType        =   4
         IdName          =   "WRHCMCTGTO"
      End
      Begin MwfCtl.MWCTL_MED ATXT 
         Height          =   285
         Index           =   79
         Left            =   1500
         TabIndex        =   100
         ToolTipText     =   "Enter From Currency Code"
         Top             =   1995
         Width           =   885
         _ExtentX        =   1561
         _ExtentY        =   503
         MaxLength       =   5
         DataType        =   4
         IdName          =   "WRHCMCURCDFR"
      End
      Begin MwfCtl.MWCTL_MED ATXT 
         Height          =   285
         Index           =   80
         Left            =   3825
         TabIndex        =   101
         ToolTipText     =   "Enter To Currency Code"
         Top             =   1995
         Width           =   885
         _ExtentX        =   1561
         _ExtentY        =   503
         MaxLength       =   5
         DataType        =   4
         IdName          =   "WRHCMCURCDTO"
      End
      Begin MwfCtl.MWCTL_MED ATXT 
         Height          =   285
         Index           =   81
         Left            =   1500
         TabIndex        =   102
         ToolTipText     =   "Enter From Raw Material Category"
         Top             =   2280
         Width           =   495
         _ExtentX        =   873
         _ExtentY        =   503
         MaxLength       =   2
         DataType        =   4
         IdName          =   "WRHCTGFR"
      End
      Begin MwfCtl.MWCTL_MED ATXT 
         Height          =   285
         Index           =   82
         Left            =   3825
         TabIndex        =   103
         ToolTipText     =   "Enter To Raw Material Category"
         Top             =   2280
         Width           =   495
         _ExtentX        =   873
         _ExtentY        =   503
         MaxLength       =   2
         DataType        =   4
         IdName          =   "WRHCTGTO"
      End
      Begin MwfCtl.MWCTL_MED ATXT 
         Height          =   285
         Index           =   83
         Left            =   1500
         TabIndex        =   104
         ToolTipText     =   "Enter From Raw Material Sub Category"
         Top             =   2565
         Width           =   885
         _ExtentX        =   1561
         _ExtentY        =   503
         MaxLength       =   5
         DataType        =   4
         IdName          =   "WRHSCTGFR"
      End
      Begin MwfCtl.MWCTL_MED ATXT 
         Height          =   285
         Index           =   84
         Left            =   3825
         TabIndex        =   105
         ToolTipText     =   "Enter To Raw Material Sub Category"
         Top             =   2565
         Width           =   885
         _ExtentX        =   1561
         _ExtentY        =   503
         MaxLength       =   5
         DataType        =   4
         IdName          =   "WRHSCTGTO"
      End
      Begin MwfCtl.MWCTL_MED ATXT 
         Height          =   285
         Index           =   85
         Left            =   1500
         TabIndex        =   106
         ToolTipText     =   "Enter From Raw Material Code"
         Top             =   2850
         Width           =   1935
         _ExtentX        =   3413
         _ExtentY        =   503
         MaxLength       =   12
         DataType        =   4
         IdName          =   "WRHCDFR"
      End
      Begin MwfCtl.MWCTL_MED ATXT 
         Height          =   285
         Index           =   90
         Left            =   3825
         TabIndex        =   107
         ToolTipText     =   "Enter To Raw Material Code"
         Top             =   2850
         Width           =   1935
         _ExtentX        =   3413
         _ExtentY        =   503
         MaxLength       =   12
         DataType        =   4
         IdName          =   "WRHCDTO"
      End
      Begin MwfCtl.MWCTL_MED ATXT 
         Height          =   285
         Index           =   91
         Left            =   1500
         TabIndex        =   110
         ToolTipText     =   "Enter From Design Category"
         Top             =   3420
         Width           =   885
         _ExtentX        =   1561
         _ExtentY        =   503
         MaxLength       =   5
         DataType        =   4
         IdName          =   "WRHDMCTGFR"
      End
      Begin MwfCtl.MWCTL_MED ATXT 
         Height          =   285
         Index           =   92
         Left            =   3825
         TabIndex        =   111
         ToolTipText     =   "Enter To Design Category"
         Top             =   3420
         Width           =   885
         _ExtentX        =   1561
         _ExtentY        =   503
         MaxLength       =   5
         DataType        =   4
         IdName          =   "WRHDMCTGTO"
      End
      Begin MwfCtl.MWCTL_MED ATXT 
         Height          =   285
         Index           =   93
         Left            =   1500
         TabIndex        =   108
         ToolTipText     =   "Enter From Size To Weight Group"
         Top             =   3135
         Width           =   885
         _ExtentX        =   1561
         _ExtentY        =   503
         MaxLength       =   5
         DataType        =   4
         IdName          =   "WRHSTWGRPFR"
      End
      Begin MwfCtl.MWCTL_MED ATXT 
         Height          =   285
         Index           =   94
         Left            =   3825
         TabIndex        =   109
         ToolTipText     =   "Enter To Size To Weight Group"
         Top             =   3135
         Width           =   885
         _ExtentX        =   1561
         _ExtentY        =   503
         MaxLength       =   5
         DataType        =   4
         IdName          =   "WRHSTWGRPTO"
      End
      Begin MwfCtl.MWCTL_MED ATXT 
         Height          =   285
         Index           =   95
         Left            =   1500
         TabIndex        =   112
         ToolTipText     =   "Enter From Design Loss Category"
         Top             =   3705
         Width           =   885
         _ExtentX        =   1561
         _ExtentY        =   503
         MaxLength       =   5
         DataType        =   4
         IdName          =   "WRHLSCTGFR"
      End
      Begin MwfCtl.MWCTL_MED ATXT 
         Height          =   285
         Index           =   96
         Left            =   3825
         TabIndex        =   113
         ToolTipText     =   "Enter To Design Loss Category"
         Top             =   3705
         Width           =   885
         _ExtentX        =   1561
         _ExtentY        =   503
         MaxLength       =   5
         DataType        =   4
         IdName          =   "WRHLSCTGTO"
      End
      Begin MwfCtl.MWCTL_MED ATXT 
         Height          =   285
         Index           =   97
         Left            =   1500
         TabIndex        =   114
         ToolTipText     =   "Enter From Labour Main Code"
         Top             =   3990
         Width           =   885
         _ExtentX        =   1561
         _ExtentY        =   503
         MaxLength       =   5
         DataType        =   4
         IdName          =   "WRHLABMCDFR"
      End
      Begin MwfCtl.MWCTL_MED ATXT 
         Height          =   285
         Index           =   98
         Left            =   3825
         TabIndex        =   115
         ToolTipText     =   "Enter To Labour Main Code"
         Top             =   3990
         Width           =   885
         _ExtentX        =   1561
         _ExtentY        =   503
         MaxLength       =   5
         DataType        =   4
         IdName          =   "WRHLABMCDTO"
      End
      Begin MwfCtl.MWCTL_MED ATXT 
         Height          =   285
         Index           =   99
         Left            =   1500
         TabIndex        =   95
         ToolTipText     =   "Enter Year And Month in 'YyyyMm' Format"
         Top             =   1140
         Width           =   1320
         _ExtentX        =   2328
         _ExtentY        =   503
         Alignment       =   2
         Mask            =   "#####0"
         MaxLength       =   6
         DataType        =   1
         IdName          =   "WRHYYYYMM"
      End
      Begin VB.Label LblMain 
         BackStyle       =   0  'Transparent
         BorderStyle     =   1  'Fixed Single
         Caption         =   "  Rm Rate History Transfer"
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
         TabIndex        =   326
         Top             =   90
         Width           =   6045
      End
      Begin VB.Label ALBL 
         BackStyle       =   0  'Transparent
         Caption         =   "YyyyMm"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Index           =   78
         Left            =   90
         TabIndex        =   271
         Top             =   1140
         Width           =   975
      End
      Begin VB.Label ALBL 
         BackStyle       =   0  'Transparent
         Caption         =   "Currency Cd"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Index           =   77
         Left            =   90
         TabIndex        =   270
         Top             =   1995
         Width           =   1305
      End
      Begin VB.Label ALBL 
         BackStyle       =   0  'Transparent
         Caption         =   "Customer Ctg"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Index           =   76
         Left            =   90
         TabIndex        =   269
         Top             =   1425
         Width           =   1425
      End
      Begin VB.Label ALBL 
         BackStyle       =   0  'Transparent
         Caption         =   "Tc Type"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Index           =   75
         Left            =   90
         TabIndex        =   268
         Top             =   855
         Width           =   1245
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
         Index           =   74
         Left            =   90
         TabIndex        =   267
         Top             =   1710
         Width           =   1335
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
         Index           =   73
         Left            =   3810
         TabIndex        =   266
         Top             =   480
         Width           =   1035
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
         Index           =   72
         Left            =   1500
         TabIndex        =   265
         Top             =   480
         Width           =   1035
      End
      Begin VB.Label ALBL 
         BackStyle       =   0  'Transparent
         Caption         =   "Modified Dt"
         BeginProperty Font 
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
         Left            =   90
         TabIndex        =   264
         Top             =   4275
         Width           =   1245
      End
      Begin VB.Label ALBL 
         BackStyle       =   0  'Transparent
         Caption         =   "Rm Code"
         BeginProperty Font 
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
         Left            =   90
         TabIndex        =   263
         Top             =   2850
         Width           =   1245
      End
      Begin VB.Label ALBL 
         BackStyle       =   0  'Transparent
         Caption         =   "Rm Sub Ctg"
         BeginProperty Font 
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
         Left            =   90
         TabIndex        =   262
         Top             =   2565
         Width           =   1245
      End
      Begin VB.Label ALBL 
         BackStyle       =   0  'Transparent
         Caption         =   "Rm Ctg"
         BeginProperty Font 
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
         Left            =   90
         TabIndex        =   261
         Top             =   2280
         Width           =   1245
      End
      Begin VB.Label ALBL 
         BackStyle       =   0  'Transparent
         Caption         =   "Design Ctg"
         BeginProperty Font 
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
         Left            =   90
         TabIndex        =   260
         Top             =   3420
         Width           =   1305
      End
      Begin VB.Label ALBL 
         BackStyle       =   0  'Transparent
         Caption         =   "SzToWt Grp"
         BeginProperty Font 
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
         Left            =   90
         TabIndex        =   259
         Top             =   3135
         Width           =   1305
      End
      Begin VB.Label ALBL 
         BackStyle       =   0  'Transparent
         Caption         =   "Dsg Loss Ctg"
         BeginProperty Font 
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
         TabIndex        =   258
         Top             =   3705
         Width           =   1455
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
         Index           =   56
         Left            =   90
         TabIndex        =   257
         Top             =   3990
         Width           =   1455
      End
   End
   Begin VB.Frame FraInv 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2220
      Left            =   3570
      TabIndex        =   195
      Top             =   6855
      Visible         =   0   'False
      Width           =   7305
      Begin VB.CommandButton CmdInvTrfOut 
         Caption         =   "&Transfer Out"
         Height          =   315
         Left            =   5820
         TabIndex        =   43
         ToolTipText     =   "Invoice Transfer Out"
         Top             =   1710
         Width           =   1230
      End
      Begin MwfCtl.MWCTL_MED ATXT 
         Height          =   285
         Index           =   11
         Left            =   1380
         TabIndex        =   30
         ToolTipText     =   "Enter From Invoice Tc"
         Top             =   855
         Width           =   615
         _ExtentX        =   1085
         _ExtentY        =   503
         MaxLength       =   3
         DataType        =   4
         IdName          =   "WINTCFR"
      End
      Begin MwfCtl.MWCTL_MED ATXT 
         Height          =   285
         Index           =   12
         Left            =   2085
         TabIndex        =   31
         ToolTipText     =   "Enter From Invoice Year"
         Top             =   855
         Width           =   330
         _ExtentX        =   582
         _ExtentY        =   503
         MaxLength       =   2
         DataType        =   4
         IdName          =   "WINYYFR"
      End
      Begin MwfCtl.MWCTL_MED ATXT 
         Height          =   285
         Index           =   13
         Left            =   2520
         TabIndex        =   32
         ToolTipText     =   "Enter From Invoice Character"
         Top             =   855
         Width           =   615
         _ExtentX        =   1085
         _ExtentY        =   503
         MaxLength       =   3
         DataType        =   4
         IdName          =   "WINCHRFR"
      End
      Begin MwfCtl.MWCTL_MED ATXT 
         Height          =   285
         Index           =   14
         Left            =   3225
         TabIndex        =   33
         ToolTipText     =   "Enter From Invoice Number"
         Top             =   855
         Width           =   855
         _ExtentX        =   1508
         _ExtentY        =   503
         Alignment       =   2
         Mask            =   "######0"
         MaxLength       =   7
         DataType        =   1
         IdName          =   "WINNOFR"
      End
      Begin MwfCtl.MWCTL_MED ATXT 
         Height          =   285
         Index           =   15
         Left            =   4365
         TabIndex        =   34
         ToolTipText     =   "Enter To Invoice Tc"
         Top             =   855
         Width           =   615
         _ExtentX        =   1085
         _ExtentY        =   503
         MaxLength       =   3
         DataType        =   4
         IdName          =   "WINTCTO"
      End
      Begin MwfCtl.MWCTL_MED ATXT 
         Height          =   285
         Index           =   16
         Left            =   5070
         TabIndex        =   35
         ToolTipText     =   "Enter To Invoice Year"
         Top             =   855
         Width           =   330
         _ExtentX        =   582
         _ExtentY        =   503
         MaxLength       =   2
         DataType        =   4
         IdName          =   "WINYYTO"
      End
      Begin MwfCtl.MWCTL_MED ATXT 
         Height          =   285
         Index           =   17
         Left            =   5475
         TabIndex        =   36
         ToolTipText     =   "Enter To Invoice Character"
         Top             =   855
         Width           =   615
         _ExtentX        =   1085
         _ExtentY        =   503
         MaxLength       =   3
         DataType        =   4
         IdName          =   "WINCHRTO"
      End
      Begin MwfCtl.MWCTL_MED ATXT 
         Height          =   285
         Index           =   18
         Left            =   6180
         TabIndex        =   37
         ToolTipText     =   "Enter To Invoice Number"
         Top             =   855
         Width           =   855
         _ExtentX        =   1508
         _ExtentY        =   503
         Alignment       =   2
         Mask            =   "######0"
         MaxLength       =   7
         DataType        =   1
         IdName          =   "WINNOTO"
      End
      Begin MwfCtl.MWCTL_MED ATXT 
         Height          =   285
         Index           =   19
         Left            =   1380
         TabIndex        =   38
         ToolTipText     =   "Enter From Customer Code"
         Top             =   1140
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   503
         MaxLength       =   8
         DataType        =   4
         IdName          =   "WINCMCDFR"
      End
      Begin MwfCtl.MWCTL_MED ATXT 
         Height          =   285
         Index           =   20
         Left            =   4365
         TabIndex        =   39
         ToolTipText     =   "Enter To Customer Code"
         Top             =   1140
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   503
         MaxLength       =   8
         DataType        =   4
         IdName          =   "WINCMCDTO"
      End
      Begin MwfCtl.MWCTL_MED ATXT 
         Height          =   285
         Index           =   22
         Left            =   1380
         TabIndex        =   42
         ToolTipText     =   "Enter From Invoice Modified Date"
         Top             =   1710
         Width           =   885
         _ExtentX        =   1561
         _ExtentY        =   503
         MaxLength       =   8
         DataType        =   3
         IdName          =   "WINMODDTFR"
      End
      Begin MwfCtl.MWCTL_MED ATXT 
         Height          =   285
         Index           =   139
         Left            =   1380
         TabIndex        =   40
         ToolTipText     =   "Enter From Invoice Date"
         Top             =   1425
         Width           =   885
         _ExtentX        =   1561
         _ExtentY        =   503
         MaxLength       =   8
         DataType        =   3
         IdName          =   "WINDTFR"
      End
      Begin MwfCtl.MWCTL_MED ATXT 
         Height          =   285
         Index           =   140
         Left            =   4365
         TabIndex        =   41
         ToolTipText     =   "Enter To Invoice Date"
         Top             =   1425
         Width           =   885
         _ExtentX        =   1561
         _ExtentY        =   503
         MaxLength       =   8
         DataType        =   3
         IdName          =   "WINDTTO"
      End
      Begin VB.Label LblMain 
         BackStyle       =   0  'Transparent
         BorderStyle     =   1  'Fixed Single
         Caption         =   "  Invoice Transfer"
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
         TabIndex        =   325
         Top             =   90
         Width           =   7305
      End
      Begin VB.Label ALBL 
         BackStyle       =   0  'Transparent
         Caption         =   "Invoice Dt"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Index           =   119
         Left            =   90
         TabIndex        =   317
         Top             =   1425
         Width           =   1245
      End
      Begin VB.Label ALBL 
         BackStyle       =   0  'Transparent
         Caption         =   "Modified Dt"
         BeginProperty Font 
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
         Left            =   90
         TabIndex        =   208
         Top             =   1710
         Width           =   1245
      End
      Begin VB.Label ALBL 
         BackStyle       =   0  'Transparent
         Caption         =   "Invoice No"
         BeginProperty Font 
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
         Left            =   90
         TabIndex        =   205
         Top             =   855
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
         Index           =   18
         Left            =   2010
         TabIndex        =   204
         Top             =   855
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
         Index           =   17
         Left            =   2430
         TabIndex        =   203
         Top             =   855
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
         Index           =   16
         Left            =   3150
         TabIndex        =   202
         Top             =   855
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
         Left            =   4980
         TabIndex        =   201
         Top             =   855
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
         Index           =   13
         Left            =   5400
         TabIndex        =   200
         Top             =   855
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
         Index           =   12
         Left            =   6090
         TabIndex        =   199
         Top             =   855
         Width           =   135
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
         Index           =   11
         Left            =   1380
         TabIndex        =   198
         Top             =   480
         Width           =   1035
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
         Index           =   10
         Left            =   4350
         TabIndex        =   197
         Top             =   480
         Width           =   1035
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
         Index           =   9
         Left            =   90
         TabIndex        =   196
         Top             =   1140
         Width           =   1245
      End
   End
   Begin VB.Frame FraOrd 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2220
      Left            =   3240
      TabIndex        =   183
      Top             =   6855
      Visible         =   0   'False
      Width           =   7305
      Begin VB.CommandButton CmdOrdTrfOut 
         Caption         =   "&Transfer Out"
         Height          =   315
         Left            =   5820
         TabIndex        =   29
         ToolTipText     =   "Order Transfer Out"
         Top             =   1725
         Width           =   1230
      End
      Begin MwfCtl.MWCTL_MED ATXT 
         Height          =   285
         Index           =   0
         Left            =   1350
         TabIndex        =   18
         ToolTipText     =   "Enter From Order Tc"
         Top             =   1155
         Width           =   615
         _ExtentX        =   1085
         _ExtentY        =   503
         MaxLength       =   3
         DataType        =   4
         IdName          =   "WOMTCFR"
      End
      Begin MwfCtl.MWCTL_MED ATXT 
         Height          =   285
         Index           =   1
         Left            =   2055
         TabIndex        =   19
         ToolTipText     =   "Enter From Order Year"
         Top             =   1155
         Width           =   330
         _ExtentX        =   582
         _ExtentY        =   503
         MaxLength       =   2
         DataType        =   4
         IdName          =   "WOMYYFR"
      End
      Begin MwfCtl.MWCTL_MED ATXT 
         Height          =   285
         Index           =   2
         Left            =   2490
         TabIndex        =   20
         ToolTipText     =   "Enter From Order Character"
         Top             =   1155
         Width           =   615
         _ExtentX        =   1085
         _ExtentY        =   503
         MaxLength       =   3
         DataType        =   4
         IdName          =   "WOMCHRFR"
      End
      Begin MwfCtl.MWCTL_MED ATXT 
         Height          =   285
         Index           =   3
         Left            =   3195
         TabIndex        =   21
         ToolTipText     =   "Enter From Order Number"
         Top             =   1155
         Width           =   855
         _ExtentX        =   1508
         _ExtentY        =   503
         Alignment       =   2
         Mask            =   "######0"
         MaxLength       =   7
         DataType        =   1
         IdName          =   "WOMNOFR"
      End
      Begin MwfCtl.MWCTL_MED ATXT 
         Height          =   285
         Index           =   4
         Left            =   4365
         TabIndex        =   22
         ToolTipText     =   "Enter To Order Tc"
         Top             =   1155
         Width           =   615
         _ExtentX        =   1085
         _ExtentY        =   503
         MaxLength       =   3
         DataType        =   4
         IdName          =   "WOMTCTO"
      End
      Begin MwfCtl.MWCTL_MED ATXT 
         Height          =   285
         Index           =   5
         Left            =   5070
         TabIndex        =   23
         ToolTipText     =   "Enter To Order Year"
         Top             =   1155
         Width           =   330
         _ExtentX        =   582
         _ExtentY        =   503
         MaxLength       =   2
         DataType        =   4
         IdName          =   "WOMYYTO"
      End
      Begin MwfCtl.MWCTL_MED ATXT 
         Height          =   285
         Index           =   6
         Left            =   5475
         TabIndex        =   24
         ToolTipText     =   "Enter To Order Character"
         Top             =   1155
         Width           =   615
         _ExtentX        =   1085
         _ExtentY        =   503
         MaxLength       =   3
         DataType        =   4
         IdName          =   "WOMCHRTO"
      End
      Begin MwfCtl.MWCTL_MED ATXT 
         Height          =   285
         Index           =   7
         Left            =   6180
         TabIndex        =   25
         ToolTipText     =   "Enter To Order Number"
         Top             =   1155
         Width           =   855
         _ExtentX        =   1508
         _ExtentY        =   503
         Alignment       =   2
         Mask            =   "######0"
         MaxLength       =   7
         DataType        =   1
         IdName          =   "WOMNOTO"
      End
      Begin MwfCtl.MWCTL_MED ATXT 
         Height          =   285
         Index           =   9
         Left            =   1350
         TabIndex        =   26
         ToolTipText     =   "Enter From Customer Code"
         Top             =   1440
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   503
         MaxLength       =   8
         DataType        =   4
         IdName          =   "WOMCMCDFR"
      End
      Begin MwfCtl.MWCTL_MED ATXT 
         Height          =   285
         Index           =   10
         Left            =   4365
         TabIndex        =   27
         ToolTipText     =   "Enter From Customer Code"
         Top             =   1440
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   503
         MaxLength       =   8
         DataType        =   4
         IdName          =   "WOMCMCDTO"
      End
      Begin MwfCtl.MWCTL_MED ATXT 
         Height          =   285
         Index           =   21
         Left            =   1350
         TabIndex        =   17
         ToolTipText     =   "Enter Transaction Type"
         Top             =   870
         Width           =   615
         _ExtentX        =   1085
         _ExtentY        =   503
         MaxLength       =   3
         DataType        =   4
         IdName          =   "WOMTCTYP"
      End
      Begin MwfCtl.MWCTL_MED ATXT 
         Height          =   285
         Index           =   24
         Left            =   1350
         TabIndex        =   28
         ToolTipText     =   "Enter From Order Modified Date"
         Top             =   1725
         Width           =   885
         _ExtentX        =   1561
         _ExtentY        =   503
         MaxLength       =   8
         DataType        =   3
         IdName          =   "WOMMODDTFR"
      End
      Begin VB.Label LblMain 
         BackStyle       =   0  'Transparent
         BorderStyle     =   1  'Fixed Single
         Caption         =   "  Order Transfer"
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
         TabIndex        =   337
         Top             =   90
         Width           =   7305
      End
      Begin VB.Label ALBL 
         BackStyle       =   0  'Transparent
         Caption         =   "Modified Dt"
         BeginProperty Font 
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
         Left            =   90
         TabIndex        =   207
         Top             =   1725
         Width           =   1245
      End
      Begin VB.Label ALBL 
         BackStyle       =   0  'Transparent
         Caption         =   "Tc Type"
         BeginProperty Font 
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
         Left            =   90
         TabIndex        =   206
         Top             =   870
         Width           =   1245
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
         Index           =   8
         Left            =   90
         TabIndex        =   194
         Top             =   1440
         Width           =   1245
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
         Index           =   7
         Left            =   4350
         TabIndex        =   193
         Top             =   480
         Width           =   1035
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
         Index           =   6
         Left            =   1350
         TabIndex        =   192
         Top             =   480
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
         Index           =   4
         Left            =   6090
         TabIndex        =   191
         Top             =   1155
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
         Left            =   5400
         TabIndex        =   190
         Top             =   1155
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
         Left            =   4980
         TabIndex        =   189
         Top             =   1155
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
         Left            =   3120
         TabIndex        =   188
         Top             =   1155
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
         Left            =   2400
         TabIndex        =   187
         Top             =   1155
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
         Left            =   1980
         TabIndex        =   186
         Top             =   1155
         Width           =   135
      End
      Begin VB.Label ALBL 
         BackStyle       =   0  'Transparent
         Caption         =   "Order No"
         BeginProperty Font 
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
         TabIndex        =   184
         Top             =   1155
         Width           =   1245
      End
   End
   Begin VB.Frame FraAllMst 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   6915
      Left            =   30
      TabIndex        =   171
      Top             =   2160
      Visible         =   0   'False
      Width           =   12000
      Begin VB.DriveListBox DriveLst 
         Height          =   315
         Left            =   1380
         TabIndex        =   175
         Top             =   5820
         Width           =   10395
      End
      Begin VB.TextBox TxtNewFolderName 
         Height          =   285
         Left            =   1380
         TabIndex        =   177
         Top             =   6420
         Width           =   3675
      End
      Begin VB.TextBox TxtFolderPath 
         Enabled         =   0   'False
         Height          =   285
         Left            =   1380
         TabIndex        =   176
         Top             =   6135
         Width           =   10395
      End
      Begin VB.CommandButton CmdAllMstTrfOut 
         Caption         =   "&Transfer Out"
         Height          =   315
         Left            =   10550
         TabIndex        =   178
         ToolTipText     =   "All Masters Transfer Out"
         Top             =   6420
         Width           =   1230
      End
      Begin MwfCtl.MWCTL_MED ATXT 
         Height          =   285
         Index           =   151
         Left            =   1380
         TabIndex        =   172
         ToolTipText     =   "Enter From Modified Date"
         Top             =   780
         Width           =   885
         _ExtentX        =   1561
         _ExtentY        =   503
         MaxLength       =   8
         DataType        =   3
         IdName          =   "WALLMSTMODDTFR"
      End
      Begin MwfCtl.MWCTL_MED ATXT 
         Height          =   285
         Index           =   141
         Left            =   3480
         TabIndex        =   173
         ToolTipText     =   "Enter To Modified Date"
         Top             =   780
         Width           =   885
         _ExtentX        =   1561
         _ExtentY        =   503
         MaxLength       =   8
         DataType        =   3
         IdName          =   "WALLMSTMODDTTO"
      End
      Begin ComctlLib.TreeView TreeViewSel 
         Height          =   4605
         Left            =   1380
         TabIndex        =   174
         Top             =   1200
         Width           =   10395
         _ExtentX        =   18336
         _ExtentY        =   8123
         _Version        =   393217
         Indentation     =   635
         LabelEdit       =   1
         Style           =   7
         Appearance      =   1
      End
      Begin ComctlLib.ImageList ImgList 
         Left            =   6870
         Top             =   6390
         _ExtentX        =   1005
         _ExtentY        =   1005
         BackColor       =   -2147483643
         ImageWidth      =   16
         ImageHeight     =   16
         MaskColor       =   12632256
         _Version        =   393216
         BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
            NumListImages   =   3
            BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "EmrFrmTrfOut.frx":0000
               Key             =   ""
            EndProperty
            BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "EmrFrmTrfOut.frx":0452
               Key             =   ""
            EndProperty
            BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "EmrFrmTrfOut.frx":08A4
               Key             =   ""
            EndProperty
         EndProperty
      End
      Begin VB.Label LblMain 
         BackStyle       =   0  'Transparent
         BorderStyle     =   1  'Fixed Single
         Caption         =   "  Transfer Of All Masters"
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
         Index           =   14
         Left            =   -30
         TabIndex        =   339
         Top             =   90
         Width           =   12000
      End
      Begin VB.Label ALBL 
         BackStyle       =   0  'Transparent
         Caption         =   "New Folder"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Index           =   122
         Left            =   90
         TabIndex        =   322
         Top             =   6420
         Width           =   1305
      End
      Begin VB.Label ALBL 
         BackStyle       =   0  'Transparent
         Caption         =   "Folder Path"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Index           =   120
         Left            =   90
         TabIndex        =   321
         Top             =   6135
         Width           =   1305
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
         Index           =   129
         Left            =   1380
         TabIndex        =   320
         Top             =   480
         Width           =   1035
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
         Index           =   130
         Left            =   3480
         TabIndex        =   319
         Top             =   480
         Width           =   1035
      End
      Begin VB.Label ALBL 
         BackStyle       =   0  'Transparent
         Caption         =   "Modified Dt"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Index           =   121
         Left            =   90
         TabIndex        =   318
         Top             =   780
         Width           =   1245
      End
   End
   Begin VB.Label ALBL 
      BackStyle       =   0  'Transparent
      Caption         =   "Target Version"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Index           =   124
      Left            =   2790
      TabIndex        =   324
      Top             =   0
      Width           =   1605
   End
   Begin VB.Label ALBL 
      BackStyle       =   0  'Transparent
      Caption         =   "Target (EMR/ XEMR)"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Index           =   123
      Left            =   30
      TabIndex        =   323
      Top             =   0
      Width           =   2205
   End
   Begin VB.Line Line1 
      Visible         =   0   'False
      X1              =   5940
      X2              =   5940
      Y1              =   1860
      Y2              =   2640
   End
End
Attribute VB_Name = "EmrFrmTrfOut"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'*****************************************************************************************
'*** Data Porting From Emr To XML
'*** Emr Ver 2.9.0 to xEmr Ver 2.6.0
'*** xEmr Ver 2.6.0 to Emr Ver 2.6.0, 2.7.0, 2.8.0, 2.9.0
'*****************************************************************************************

Option Explicit

Dim WithEvents Explorer As CExplorer2
Attribute Explorer.VB_VarHelpID = -1

Dim moCn As MwfLib.MDOConnection   '*** (Bef speed) bef 24/11/06

'*** Declarations For The Source RecordSets
Dim Rs_SrcX As MwfLib.MDORowSet
Dim Rs_SrcA As MwfLib.MDORowSet
Dim Rs_SrcB As MwfLib.MDORowSet
Dim Rs_SrcC As MwfLib.MDORowSet
Dim Rs_SrcB1 As MwfLib.MDORowSet
Dim Rs_SrcC1 As MwfLib.MDORowSet
Dim Rs_SrcD As MwfLib.MDORowSet
Dim Rs_SrcE As MwfLib.MDORowSet
Dim Rs_SrcG As MwfLib.MDORowSet

'*** Declarations For The XML Document
Dim XmlDoc As MSXML2.DOMDocument
Dim xmlRootElement As MSXML2.IXMLDOMElement
Dim xmlChildX As MSXML2.IXMLDOMElement
Dim xmlChildA As MSXML2.IXMLDOMElement
Dim xmlChildB As MSXML2.IXMLDOMElement
Dim xmlChildC As MSXML2.IXMLDOMElement
Dim xmlChildB1 As MSXML2.IXMLDOMElement
Dim xmlChildC1 As MSXML2.IXMLDOMElement
Dim xmlChildD As MSXML2.IXMLDOMElement
Dim xmlChildE As MSXML2.IXMLDOMElement
Dim xmlChildG As MSXML2.IXMLDOMElement

'*** Arr_Cv stores the causes validation state, Arr_Ed the Ena/Disa state of Cmd Buttons
Dim mArr_Cv As Variant, mArr_Ed As Variant

Dim ms_Cnd As String, ms_Str As String
Dim mf_Cancel As Boolean, ms_ErrMsg As String

'*** 2.11 (02/08/05)
Dim ms_ChildModDtCnd As String
'*** 2.11 (02/08/05)

Dim mf_AllMst As Boolean

'Dim ms_TrfOutVer As String

Private Enum en_EmrTrfOutDispFra
  Rm = 0
  Cust = 1
  Dsg = 2
  RmRt = 3
  LabRt = 4
  RmRtHist = 5
  LabRtHist = 6
  DtTbl = 7
  BkAg = 8
  Loc = 9
  Prm = 10
  VPrm = 11
  Ord = 12
  Inv = 13
  AllMst = 14
End Enum

' Zubin 213
Dim mi_MajVer As Integer, mi_MinVer As Integer
Dim mi_TgtMajVer As Integer, mi_TgtMinVer As Integer
Private Sub ADC_Load()
  'If ctProjectName = "XEMR" Then
  '  ms_TrfOutVer = "2.9.0"
  '  ''ms_TrfOutVer = "2.6.0"    'used For testing
  'ElseIf ctProjectName = "EMR" Then
  '  ms_TrfOutVer = "2.6.0"
  'End If
  
  ReDim mArr_Cv(15), mArr_Ed(15)
  Set moCn = adc.Connection    '*** (Bef speed) bef 24/11/06
  Set adc.FirKeyCtl = adc("wTgtExe")
  Call HideAllFras
  
  mf_AllMst = False
  
  ' Zubin 213
  Dim ws_TmpVer As String
  ws_TmpVer = gs_Ver
  mi_MajVer = Mid(ws_TmpVer, 1, InStr(1, ws_TmpVer, ".") - 1)
  ws_TmpVer = Mid(ws_TmpVer, InStr(1, ws_TmpVer, ".") + 1)
  mi_MinVer = Mid(ws_TmpVer, 1, InStr(1, ws_TmpVer, ".") - 1)

  '''Call EnaDisaCmds(False)
  
  '*** Done ***
  'Bank
  'CustAly, CustDsg, CustMst, CustRm
  'DsgAna, DsgLab, DsgMst, DsgRm
  'DtTbl
  'LabRt
  'LabRtHist
  'Loc
  'Param
  'RmAlyCmp , RmMst
  'RmRt
  'RmRtHist
  'vParam

  '*** Not To Be Done ***
  'DsgLoad
  'Head
  'LsRec
  'PpcCap
  'PrdPts
End Sub
Private Sub ADC_KeyWhen()
  adc("wTgtExe") = "N"
  adc("wTgtVer") = "2.6.0"
End Sub
Private Sub ATXT_FldWhen(Index As Integer, ByVal IdName As String, Cancel As Boolean, ErrMsg As String)
  Select Case UCase(IdName)
  Case Is = UCase("wTgtExe")
    Call HlpList.PMCd("YN")
  
  '*** Rm ***
  Case Is = UCase("wRmZ")
    Call HlpList.PMCd("YN")
  Case Is = UCase("wRmCtgFr"), UCase("wRmCtgTo")
    Call HlpList.PMCd("RMCTG")
  Case Is = UCase("wRmSCtgFr")
    Call HlpList.PSCd("RMSCTG", adc("wRmCtgFr"))
  Case Is = UCase("wRmSCtgTo")
    Call HlpList.PSCd("RMSCTG", adc("wRmCtgTo"))
  Case Is = UCase("wRmCdFr")
    Call HlpList.RmCd(adc("wRmZ"), "'" + adc("wRmCtgFr") + "'", adc("wRmSCtgFr"))
  Case Is = UCase("wRmCdTo")
    Call HlpList.RmCd(adc("wRmZ"), "'" + adc("wRmCtgTo") + "'", adc("wRmSCtgTo"))
  '*** Rm ***
  
  '*** Customer ***
  Case Is = UCase("wCmCtg")
    Call HlpList.PMCd("CMCTG")
  Case Is = UCase("wCmCpyChild")
    Call HlpList.PMCd("YN")
  Case Is = UCase("wCmCdFr"), UCase("wCmCdTo")
    Call HlpList.CustCd("C")
  '*** Customer ***
  
  '*** Design ***
  Case Is = UCase("wDmTcTyp")
    Call HlpList.PMCd("DMTCTYP")
  Case Is = UCase("wDmCdFr"), UCase("wDmCdTo")
    Call HlpList.DmCd(adc("wDmTcTyp"))
  Case Is = UCase("wDmCdFr"), UCase("wDmCdTo")
    Call HlpList.DmCd(adc("wDmTcTyp"))
  Case Is = UCase("wDmSzFr"), UCase("wDmSzTo")
    Call HlpList.PMCd("DMSZ")
  Case Is = UCase("wDmCtgFr"), UCase("wDmCtgTo")
    Call HlpList.PMCd("DMCTG")
  '*** Design ***
  
  '*** LabRt ***
  Case Is = UCase("wLrCmCtgFr"), UCase("wLrCmCtgTo")
    Call HlpList.PMCd("CMCTG")
  Case Is = UCase("wLrCmCdFr")
    Call HlpList.CustCd(adc("wLrCmCtgFr"))
  Case Is = UCase("wLrCmCdTo")
    Call HlpList.CustCd(adc("wLrCmCtgTo"))
  Case Is = UCase("wLrCmCurCdFr"), UCase("wLrCmCurCdTo")
    Call HlpList.PMCd("CURNCY")
  Case Is = UCase("wLrMCdFr"), UCase("wLrMCdTo")
    Call HlpList.PMCd("LABMCD")
  Case Is = UCase("wLrSCdFr")
    Call HlpList.PSCd("LABSCD", adc("wLrMCdFr"))
  Case Is = UCase("wLrSCdTo")
    Call HlpList.PSCd("LABSCD", adc("wLrMCdTo"))
  '*** LabRt ***
  
  '*** RmRt ***
  Case Is = UCase("wRrTcTyp")
    Call HlpList.PMCd("RRTCTYP")
  Case Is = UCase("wRrCmCtgFr"), UCase("wRrCmCtgTo")
    Call HlpList.PMCd("CMCTG")
  Case Is = UCase("wRrCmCdFr")
    Call HlpList.CustCd(adc("wRrCmCtgFr"))
  Case Is = UCase("wRrCmCdTo")
    Call HlpList.CustCd(adc("wRrCmCtgTo"))
  Case Is = UCase("wRrCmCurCdFr"), UCase("wRrCmCurCdTo")
    Call HlpList.PMCd("CURNCY")
  Case Is = UCase("wRrCtgFr"), UCase("wRrCtgTo")
    Call HlpList.PMCd("RMCTG")
  Case Is = UCase("wRrSCtgFr")
    Call HlpList.PSCd("RMSCTG", adc("wRrCtgFr"))
  Case Is = UCase("wRrSCtgTo")
    Call HlpList.PSCd("RMSCTG", adc("wRrCtgTo"))
  Case Is = UCase("wRrCdFr")
    Call HlpList.RmCd("N", "'" + adc("wRrCtgFr") + "'", adc("wRrSCtgFr"))
  Case Is = UCase("wRrCdTo")
    Call HlpList.RmCd("N", "'" + adc("wRrCtgTo") + "'", adc("wRrSCtgTo"))
  Case Is = UCase("wRrSTWGrpFr"), UCase("wRrSTWGrpTo")
    Call HlpList.PMCd("STWGRP")
  Case Is = UCase("wRrDmCtgFr"), UCase("wRrDmCtgTo")
    Call HlpList.PMCd("DMCTG")
  Case Is = UCase("wRrLsCtgFr"), UCase("wRrLsCtgTo")
    Call HlpList.PMCd("LSCTG")
  Case Is = UCase("wRrLabMCdFr"), UCase("wRrLabMCdTo")
    Call HlpList.PMCd("LABMCD")
  '*** RmRt ***
  
  '*** LabRtHist ***
  Case Is = UCase("wLhYyyyMm")
    HlpList.YyyyMm
  Case Is = UCase("wLhCmCtgFr"), UCase("wLhCmCtgTo")
    Call HlpList.PMCd("CMCTG")
  Case Is = UCase("wLrCmCdFr")
    Call HlpList.CustCd(adc("wLhCmCtgFr"))
  Case Is = UCase("wLhCmCdTo")
    Call HlpList.CustCd(adc("wLhCmCtgTo"))
  Case Is = UCase("wLhCmCurCdFr"), UCase("wLhCmCurCdTo")
    Call HlpList.PMCd("CURNCY")
  Case Is = UCase("wLhMCdFr"), UCase("wLhMCdTo")
    Call HlpList.PMCd("LABMCD")
  Case Is = UCase("wLhSCdFr")
    Call HlpList.PSCd("LABSCD", adc("wLhMCdFr"))
  Case Is = UCase("wLhSCdTo")
    Call HlpList.PSCd("LABSCD", adc("wLhMCdTo"))
  '*** LabRtHist ***
  
  '*** RmRtHist ***
  Case Is = UCase("wRhYyyyMm")
    HlpList.YyyyMm
  Case Is = UCase("wRhTcTyp")
    Call HlpList.PMCd("RRTCTYP")
  Case Is = UCase("wRhCmCtgFr"), UCase("wRhCmCtgTo")
    Call HlpList.PMCd("CMCTG")
  Case Is = UCase("wRhCmCdFr")
    Call HlpList.CustCd(adc("wRhCmCtgFr"))
  Case Is = UCase("wRhCmCdTo")
    Call HlpList.CustCd(adc("wRhCmCtgTo"))
  Case Is = UCase("wRhCmCurCdFr"), UCase("wRhCmCurCdTo")
    Call HlpList.PMCd("CURNCY")
  Case Is = UCase("wRhCtgFr"), UCase("wRhCtgTo")
    Call HlpList.PMCd("RMCTG")
  Case Is = UCase("wRhSCtgFr")
    Call HlpList.PSCd("RMSCTG", adc("wRhCtgFr"))
  Case Is = UCase("wRrSCtgTo")
    Call HlpList.PSCd("RMSCTG", adc("wRhCtgTo"))
  Case Is = UCase("wRhCdFr")
    Call HlpList.RmCd("N", "'" + adc("wRhCtgFr") + "'", adc("wRhSCtgFr"))
  Case Is = UCase("wRmCdTo")
    Call HlpList.RmCd("N", "'" + adc("wRhCtgTo") + "'", adc("wRhSCtgTo"))
  Case Is = UCase("wRhSTWGrpFr"), UCase("wRhSTWGrpTo")
    Call HlpList.PMCd("STWGRP")
  Case Is = UCase("wRhDmCtgFr"), UCase("wRhDmCtgTo")
    Call HlpList.PMCd("DMCTG")
  Case Is = UCase("wRhLsCtgFr"), UCase("wRhLsCtgTo")
    Call HlpList.PMCd("LSCTG")
  Case Is = UCase("wRhLabMCdFr"), UCase("wRhLabMCdTo")
    Call HlpList.PMCd("LABMCD")
  '*** RmRtHist ***
  
  '*** DtTbl ***
  Case Is = UCase("wDtTag")
    Call HlpList.PMCd("DTTAG")
  '*** DtTbl ***
  
  '*** BnkAgMst ***
  Case Is = UCase("wBkCtg")
    Call HlpList.PMCd("BKCTG")
  '*** BnkAgMst ***
  
  '*** Loc ***
  Case Is = UCase("wLocTyp")
    Call HlpList.PMCd("LOCTYP")
  Case Is = UCase("wLocCdFr"), UCase("wLocCdTo")
    Call HlpList.Loc(gs_CoCd, "'" + adc("wLocTyp") + "'")
  '*** Loc ***
  
  '*** Param ***
  Case Is = UCase("wPCtg")
    Call HlpList.PMCd("YN")
  Case Is = UCase("wPTyp")
    Call HlpList.PMCd("PTYP", "'" + adc("wPCtg") + "'")
  Case Is = UCase("wPMCdFr"), UCase("wPMCdTo")
    Call HlpList.PMCd(adc("wPTyp"))
  Case Is = UCase("wPSCdFr")
    Call HlpList.PSCd(adc("wPTyp"), adc("wPMCdFr"))
  Case Is = UCase("wPSCdTo")
    Call HlpList.PSCd(adc("wPTyp"), adc("wPMCdTo"))
  '*** Param ***
  
  '*** vParam ***
  Case Is = UCase("wvPCtg")
    Call HlpList.PMCd("YN")
  Case Is = UCase("wvPTyp")
    Call HlpList.vPMCd(ctSelfCoCd, "PTYP", "'" + adc("wvPCtg") + "'")
  Case Is = UCase("wvPMCdFr"), UCase("wvPMCdTo")
    Call HlpList.vPMCd(gs_CoCd, adc("wvPTyp"))
  Case Is = UCase("wvPSCdFr")
    Call HlpList.vPSCd(gs_CoCd, adc("wvPTyp"), adc("wvPMCdFr"))
  Case Is = UCase("wvPSCdTo")
    Call HlpList.vPSCd(gs_CoCd, adc("wvPTyp"), adc("wvPMCdTo"))
  '*** vParam ***
  
  '*** Order ***
  Case Is = UCase("wOmTcTyp")
    Call HlpList.PMCd("TCTYP", , "OM")
  Case Is = UCase("wOmTcFr"), UCase("wOmTcTo")
    Call HlpList.PMCd("TC", "'" + adc("wOmTcTyp") + "'")
  Case Is = UCase("wOmYyFr"), UCase("wOmYyTo")
    Call HlpList.PMCd("YY")
  Case Is = UCase("wOmChrFr")
    Call HlpList.vPSCd(gs_CoCd, "CHR", adc("wOmTcFr"))
  Case Is = UCase("wOmChrTo")
    Call HlpList.vPSCd(gs_CoCd, "CHR", adc("wOmTcTo"))
  Case Is = UCase("wOmNoFr")
    Call HlpList.OmNo(gs_CoCd, adc("wOmTcFr"), adc("wOmYyFr"), adc("wOmChrFr"))
  Case Is = UCase("wOmNoTo")
    Call HlpList.OmNo(gs_CoCd, adc("wOmTcTo"), adc("wOmYyTo"), adc("wOmChrTo"))
  Case Is = UCase("wOmCmCdFr"), UCase("wOmCmCdTo")
    Call HlpList.CustCd("C")
  '*** Order ***
  
  '*** Invoice ***
  Case Is = UCase("wInTcFr"), UCase("wInTcTo")
    Call HlpList.PMCd("TC", "'IN'")
  Case Is = UCase("wInYyFr"), UCase("wInYyTo")
    Call HlpList.PMCd("YY")
  Case Is = UCase("wInChrFr")
    Call HlpList.vPSCd(gs_CoCd, "CHR", adc("wInTcFr"))
  Case Is = UCase("wInChrTo")
    Call HlpList.vPSCd(gs_CoCd, "CHR", adc("wInTcTo"))
  Case Is = UCase("wInNoFr")
    Call HlpList.InNo(gs_CoCd, adc("wInTcFr"), adc("wInYyFr"), adc("wInChrFr"))
  Case Is = UCase("wInNoTo")
    Call HlpList.InNo(gs_CoCd, adc("wInTcTo"), adc("wInYyTo"), adc("wInChrTo"))
  Case Is = UCase("wInCmCdFr"), UCase("wInCmCdTo")
    Call HlpList.CustCd("C")
  '*** Invoice ***
  End Select
End Sub
Private Sub atxt_FldChange(Index As Integer, ByVal IdName As String, Cancel As Boolean, ErrMsg As String, ByVal pv_NewValue As Variant)
  '*** wLhYyyyMm, wRhYyyyMm       Valid Year & Month (Year should be between 1900 & 2099, Month between 1 & 12)
  Dim wi_Year As Integer, wi_Month As Integer
  
  Select Case UCase(IdName)
  Case Is = UCase("wTgtExe")
    Cancel = Not moCn.RecSeek("Select PMCd from Param where PTyp= 'YN' and PMCd = '" + pv_NewValue + "'")
    If Cancel = True Then ErrMsg = "Specify Whether The Target Exe Is 'EMR' Or 'XEMR' ('Y'/ 'N')": Exit Sub
  Case Is = UCase("wTgtVer")
    '****** (Jen 2.13) ******
    If adc("wTgtExe") = "Y" Then
      Cancel = (pv_NewValue <> "2.6.0" And pv_NewValue <> "2.7.0" And _
                pv_NewValue <> "2.8.0" And pv_NewValue <> "2.9.0" And pv_NewValue <> "2.10.0" And pv_NewValue <> "2.11.0" And pv_NewValue <> "2.12.0" And pv_NewValue <> "2.13.0")
      If Cancel = True Then ErrMsg = "Invalid Target Version (Possible Values: '2.6.0', '2.7.0', '2.8.0', '2.9.0' Or '2.10.0' Or '2.11.0' Or '2.12.0'  Or '2.13.0')": Exit Sub
    ElseIf adc("wTgtExe") = "N" Then
      Cancel = (pv_NewValue <> "2.6.0")
      If Cancel = True Then ErrMsg = "Invalid Target Version (Possible Values: '2.6.0')": Exit Sub
    End If
    ' Zubin 213
    Dim ws_TmpVer As String
    ws_TmpVer = pv_NewValue
    mi_TgtMajVer = Mid(ws_TmpVer, 1, InStr(1, ws_TmpVer, ".") - 1)
    ws_TmpVer = Mid(ws_TmpVer, InStr(1, ws_TmpVer, ".") + 1)
    mi_TgtMinVer = Mid(ws_TmpVer, 1, InStr(1, ws_TmpVer, ".") - 1)
    
  Case Is = UCase("wRmZ")
    Cancel = (pv_NewValue <> "") And (Not moCn.RecSeek("Select PMCd from Param where PTyp= 'YN' and PMCd = '" + pv_NewValue + "'"))
    If Cancel = True Then ErrMsg = "Specify Whether To Transfer Customs Code Or Normal Code": Exit Sub
  Case Is = UCase("wCmCtg")
    Cancel = (pv_NewValue <> "") And (Not moCn.RecSeek("Select PMCd from Param where PTyp= 'CMCTG' and PMCd = '" + pv_NewValue + "'"))
    If Cancel = True Then ErrMsg = "Invalid Customer Category": Exit Sub
  Case Is = UCase("wCmCpyChild")
    Cancel = (pv_NewValue <> "") And (Not moCn.RecSeek("Select PMCd from Param where PTyp= 'YN' and PMCd = '" + pv_NewValue + "'"))
    If Cancel = True Then ErrMsg = "Specify Whether To Copy Customer Child Records (Y/ N)": Exit Sub
  Case Is = UCase("wOmTcTyp")
    Cancel = (pv_NewValue <> "") And (Not moCn.RecSeek("Select PMCd from Param where PTyp= 'TCTYP' and " + _
                              "PMCd = '" + pv_NewValue + "' and PValue1= 'OM'"))
    If Cancel = True Then ErrMsg = "Invalid Order Tc Type": Exit Sub
  Case Is = UCase("wDmTcTyp")
    Cancel = (pv_NewValue <> "") And (Not moCn.RecSeek("Select PMCd from Param where PTyp= 'DMTCTYP' and PMCd = '" + pv_NewValue + "'"))
    If Cancel = True Then ErrMsg = "Invalid Design Type": Exit Sub
  Case Is = UCase("wRrTcTyp"), UCase("wRhTcTyp")
    Cancel = (pv_NewValue <> "") And (Not moCn.RecSeek("Select PMCd from Param where PTyp= 'RRTCTYP' and PMCd = '" + pv_NewValue + "'"))
    If Cancel = True Then ErrMsg = "Invalid Raw Material Rate Type": Exit Sub
  Case Is = UCase("wLhYyyyMm"), UCase("wRhYyyyMm")
    If pv_NewValue = 0 Then Exit Sub
    If Len(CStr(pv_NewValue)) <> 6 Then Cancel = True: ErrMsg = "Invalid Year & Month": Exit Sub
    wi_Year = Val(Mid(CStr(pv_NewValue), 1, 4))
    If wi_Year < 1900 Or wi_Year > 2099 Then Cancel = True: ErrMsg = "Invalid Year": Exit Sub
    wi_Month = Val(Mid(CStr(pv_NewValue), 5, 2))
    If wi_Month < 1 Or wi_Month > 12 Then Cancel = True: ErrMsg = "Invalid Month": Exit Sub
  Case Is = UCase("wDtTag")
    Cancel = (pv_NewValue <> "") And (Not moCn.RecSeek("Select PMCd from Param where PTyp= 'DTTAG' and PMCd = '" + pv_NewValue + "'"))
    If Cancel = True Then ErrMsg = "Invalid Date Tag": Exit Sub
  Case Is = UCase("wBkCtg")
    Cancel = (pv_NewValue <> "") And (Not moCn.RecSeek("Select PMCd from Param where PTyp= 'BKCTG' and PMCd = '" + pv_NewValue + "'"))
    If Cancel = True Then ErrMsg = "Invalid Bank/ Agent Category": Exit Sub
  Case Is = UCase("wLocTyp")
    Cancel = (pv_NewValue <> "") And (Not moCn.RecSeek("Select PMCd from Param where PTyp= 'LOCTYP' and PMCd = '" + pv_NewValue + "'"))
    If Cancel = True Then ErrMsg = "Invalid Location Type": Exit Sub
  Case Is = UCase("wPCtg")
    Cancel = (pv_NewValue <> "") And (Not moCn.RecSeek("Select PMCd from Param where PTyp= 'YN' and PMCd = '" + pv_NewValue + "'"))
    If Cancel = True Then ErrMsg = "Specify Whether The Transfer Is For A System Or Other Parameter": Exit Sub
  Case Is = UCase("wPTyp")
    Cancel = (pv_NewValue <> "") And (Not moCn.RecSeek("Select PMCd from Param where PTyp= 'PTYP' and PMCd = '" + pv_NewValue + "' and PValue= '" + adc("wPCtg") + "'"))
    If Cancel = True Then ErrMsg = "Invalid " + IIF(adc("wPCtg") = "Y", "System", IIF(adc("wPCtg") = "N", "Other", "")) + " Parameter Type": Exit Sub
  Case Is = UCase("wvPCtg")
    Cancel = (pv_NewValue <> "") And (Not moCn.RecSeek("Select PMCd from Param where PTyp= 'YN' and PMCd = '" + pv_NewValue + "'"))
    If Cancel = True Then ErrMsg = "Specify Whether The Transfer Is For A System Or Other Variable Parameter": Exit Sub
  Case Is = UCase("wvPTyp")
    Cancel = (pv_NewValue <> "") And (Not moCn.RecSeek("Select vPMCd from vParam where vPCoCd= '" + ctSelfCoCd + "' and vPTyp= 'PTYP' and vPMCd = '" + pv_NewValue + "' and vPValue= '" + adc("wvPCtg") + "'"))
    If Cancel = True Then ErrMsg = "Invalid " + IIF(adc("wvPCtg") = "Y", "System", IIF(adc("wvPCtg") = "N", "Other", "")) + " Variable Parameter Type": Exit Sub
  End Select
End Sub
Private Sub atxt_ReCalculate(Index As Integer, ByVal IdName As String, ByVal OldVal As Variant, ByVal RowNum As Integer, ByVal SourceName As String, ByVal SourceColName As String)
  Select Case UCase(IdName)
  Case Is = UCase("wTgtVer")
    If adc.Mode = xNorm Then
      If adc("wTgtExe") = "N" Then adc("wTgtVer") = "2.6.0"
    End If
  End Select
End Sub
Private Sub CmdBkAg_Click()
  Call DispFra(BkAg)
  If FraBkAg.Visible = True Then
    If adc("wBkModDtFr") = "" Then adc("wBkModDtFr") = "01/01/80"
  End If
End Sub
Private Sub CmdBkAgTrfOut_Click()
  If adc("wBkCtg") = "" Then DispMsg "Invalid Bank/ Agent Category", etError: Exit Sub

  MWLib.BeginProcess Me, "Bank / Agent Master Tranfer Out is In Progress"
  
  adc("wBkCtg").CmpStr = "BkCtg= ": adc("wBkModDtFr").CmpStr = "ModDt>= "
  adc("wBkCdFr").CmpStr = "BkCd>= ": adc("wBkCdTo").CmpStr = "BkCd<= "
  
  ms_Cnd = adc.RepCond
  ms_Cnd = IIF(ms_Cnd <> "", " where " + ms_Cnd, "")
  ms_Str = "Select BkCtg, BkCd, BkName, BkAdd1, BkAdd2, BkAdd3, " + _
           "convert(varchar(5800), BkInstr) as BkInstr, ModUsr, ModDt, " + _
           "ModTime From BnkAgMst " + ms_Cnd + " Order By BkCd"
  Call CreateXmlDocMst("BnkAgMst", ms_Str, adc("wBkCtg"))
  
  adc("wBkCtg").CmpStr = "": adc("wBkModDtFr").CmpStr = ""
  adc("wBkCdFr").CmpStr = "": adc("wBkCdTo").CmpStr = ""
  
  MWLib.EndProcess Me
  
  If FraBkAg.Visible = True Then
    Call CmdBkAg_Click
  End If
End Sub
Private Sub CmdRm_Click()
  Call DispFra(Rm)
  If FraRm.Visible = True Then
    If adc("wRmModDtFr") = "" Then adc("wRmModDtFr") = "01/01/80"
  End If
End Sub
Private Sub CmdRmTrfOut_Click()
  If adc("wRmZ") = "" Then DispMsg "Specify Whether To Transfer Customs Code Or Normal Code", etError: Exit Sub

  MWLib.BeginProcess Me, "Raw Material Master Tranfer Out is In Progress"
  
  '*** (Bef 2.11) (Bef 02/08/05)
  'ADC("wRmModDtFr").CmpStr = "ModDt>= "
  '*** (Bef 2.11) (Bef 02/08/05)
  
  adc("wRmZ").CmpStr = "RmZ= "
  adc("wRmCtgFr").CmpStr = "RmCtg>= ": adc("wRmCtgTo").CmpStr = "RmCtg<= "
  adc("wRmSCtgFr").CmpStr = "RmSCtg>= ": adc("wRmSCtgTo").CmpStr = "RmSCtg<= "
  adc("wRmCdFr").CmpStr = "RmCd>= ": adc("wRmCdTo").CmpStr = "RmCd<= "
  
  Call CreateXmlDocRm(adc("wRmZ"))
  
  adc("wRmModDtFr").CmpStr = "": adc("wRmZ").CmpStr = ""
  adc("wRmCtgFr").CmpStr = "": adc("wRmCtgTo").CmpStr = ""
  adc("wRmSCtgFr").CmpStr = "": adc("wRmSCtgTo").CmpStr = ""
  adc("wRmCdFr").CmpStr = "": adc("wRmCdTo").CmpStr = ""
  
  MWLib.EndProcess Me
  
  If FraRm.Visible = True Then
    Call CmdRm_Click
  End If
End Sub
Private Sub CmdCust_Click()
  Call DispFra(Cust)
  If FraCust.Visible = True Then
    If adc("wCmCpyChild") = "" Then adc("wCmCpyChild") = "Y"
    If adc("wCmModDtFr") = "" Then adc("wCmModDtFr") = "01/01/80"
  End If
End Sub
Private Sub CmdCustTrfOut_Click()
  If adc("wCmCtg") = "" Then DispMsg "Invalid Customer Category", etError: Exit Sub
  If adc("wCmCpyChild") = "" Then DispMsg "Specify Whether To Copy Customer Child Records (Y/ N)", etError: Exit Sub

  MWLib.BeginProcess Me, "Customer Master Tranfer Out is In Progress"
  
  '*** (Bef 2.11) (Bef 02/08/05)
  'ADC("wCmModDtFr").CmpStr = "ModDt>= "
  '*** (Bef 2.11) (Bef 02/08/05)
  
  adc("wCmCtg").CmpStr = "CmCtg= "
  adc("wCmCdFr").CmpStr = "CmCd>= ": adc("wCmCdTo").CmpStr = "CmCd<= "
  
  Call CreateXmlDocCust(adc("wCmCtg"))
  
  adc("wCmModDtFr").CmpStr = "": adc("wCmCtg").CmpStr = ""
  adc("wCmCdFr").CmpStr = "": adc("wCmCdTo").CmpStr = ""
   
  MWLib.EndProcess Me
  
  If FraCust.Visible = True Then
    Call CmdCust_Click
  End If
End Sub
Private Sub CmdDsg_Click()
  Call DispFra(Dsg)
  If FraDsg.Visible = True Then
    If adc("wDmModDtFr") = "" Then adc("wDmModDtFr") = "01/01/80"
  End If
End Sub
Private Sub CmdDsgTrfOut_Click()
  If adc("wDmTcTyp") = "" Then DispMsg "Invalid Design Type", etError: Exit Sub
  
  MWLib.BeginProcess Me, "Design Master Tranfer Out is In Progress"
  
  '*** (Bef 2.11) (Bef 02/08/05)
  'ADC("wDmModDtFr").CmpStr = "ModDt>= "
  '*** (Bef 2.11) (Bef 02/08/05)
  
  adc("wDmTcTyp").CmpStr = "DmTcTyp= "
  adc("wDmCdFr").CmpStr = "DmCd>= ": adc("wDmCdTo").CmpStr = "DmCd<= "
  adc("wDmSzFr").CmpStr = "DmSz>= ": adc("wDmSzTo").CmpStr = "DmSz<= "
  adc("wDmCtgFr").CmpStr = "DmCtg>= ": adc("wDmCtgTo").CmpStr = "DmCtg<= "
  
  Call CreateXmlDocDsg(adc("wDmTcTyp"))
  
  adc("wDmModDtFr").CmpStr = "": adc("wDmTcTyp").CmpStr = ""
  adc("wDmCdFr").CmpStr = "": adc("wDmCdTo").CmpStr = ""
  adc("wDmSzFr").CmpStr = "": adc("wDmSzTo").CmpStr = ""
  adc("wDmCtgFr").CmpStr = "": adc("wDmCtgTo").CmpStr = ""
  
  MWLib.EndProcess Me
  
  If FraDsg.Visible = True Then
    Call CmdDsg_Click
  End If
End Sub
Private Sub CmdLabRt_Click()
  Call DispFra(LabRt)
  If FraLabRt.Visible = True Then
    If adc("wLrModDtFr") = "" Then adc("wLrModDtFr") = "01/01/80"
  End If
End Sub
Private Sub CmdLabRtTrfOut_Click()
  MWLib.BeginProcess Me, "Labour Rate Tranfer Out is In Progress"
  
  adc("wLrModDtFr").CmpStr = "ModDt>= "
  adc("wLrCmCtgFr").CmpStr = "LrCmCtg>= ": adc("wLrCmCtgTo").CmpStr = "LrCmCtg<= "
  adc("wLrCmCdFr").CmpStr = "LrCmCd>= ": adc("wLrCmCdTo").CmpStr = "LrCmCd<= "
  adc("wLrCmCurCdFr").CmpStr = "LrCmCurCd>= ": adc("wLrCmCurCdTo").CmpStr = "LrCmCurCd<= "
  adc("wLrMCdFr").CmpStr = "LrMCd>= ": adc("wLrMCdTo").CmpStr = "LrMCd<= "
  adc("wLrSCdFr").CmpStr = "LrSCd>= ": adc("wLrSCdTo").CmpStr = "LrSCd<= "
  
  ms_Cnd = adc.RepCond
  ms_Cnd = IIF(ms_Cnd <> "", " where " + ms_Cnd, "")
  ms_Str = "Select * from LabRt " + ms_Cnd + " Order By LrCmCtg, LrCmCd, LrCmCurCd, LrMCd, LrSCd"
  Call CreateXmlDocMst("LabRt", ms_Str)
  
  adc("wLrModDtFr").CmpStr = ""
  adc("wLrCmCtgFr").CmpStr = "": adc("wLrCmCtgTo").CmpStr = ""
  adc("wLrCmCdFr").CmpStr = "": adc("wLrCmCdTo").CmpStr = ""
  adc("wLrCmCurCdFr").CmpStr = "": adc("wLrCmCurCdTo").CmpStr = ""
  adc("wLrMCdFr").CmpStr = "": adc("wLrMCdTo").CmpStr = ""
  adc("wLrSCdFr").CmpStr = "": adc("wLrSCdTo").CmpStr = ""
  
  MWLib.EndProcess Me
  
  If FraLabRt.Visible = True Then
    Call CmdLabRt_Click
  End If
End Sub
Private Sub CmdRmRt_Click()
  Call DispFra(RmRt)
  If FraRmRt.Visible = True Then
    If adc("wRrModDtFr") = "" Then adc("wRrModDtFr") = "01/01/80"
  End If
End Sub
Private Sub CmdRmRtTrfOut_Click()
  If adc("wRrTcTyp") = "" Then DispMsg "Invalid Raw Material Rate Type", etError: Exit Sub

  MWLib.BeginProcess Me, "RM Rate Master Tranfer Out is In Progress"
  
  adc("wRrModDtFr").CmpStr = "ModDt>= ": adc("wRrTcTyp").CmpStr = "RrTcTyp= "
  adc("wRrCmCtgFr").CmpStr = "RrCmCtg>= ": adc("wRrCmCtgTo").CmpStr = "RrCmCtg<= "
  adc("wRrCmCdFr").CmpStr = "RrCmCd>= ": adc("wRrCmCdTo").CmpStr = "RrCmCd<= "
  adc("wRrCmCurCdFr").CmpStr = "RrCmCurCd>= ": adc("wRrCmCurCdTo").CmpStr = "RrCmCurCd<= "
  adc("wRrCtgFr").CmpStr = "RrCtg>= ": adc("wRrCtgTo").CmpStr = "RrCtg<= "
  adc("wRrSCtgFr").CmpStr = "RrSCtg>= ": adc("wRrSCtgTo").CmpStr = "RrSCtg<= "
  adc("wRrCdFr").CmpStr = "RrCd>= ": adc("wRrCdTo").CmpStr = "RrCd<= "
  adc("wRrSTWGrpFr").CmpStr = "RrSTWGrp>= ": adc("wRrSTWGrpTo").CmpStr = "RrSTWGrp<= "
  adc("wRrDmCtgFr").CmpStr = "RrDmCtg>= ": adc("wRrDmCtgTo").CmpStr = "RrDmCtg<= "
  adc("wRrLsCtgFr").CmpStr = "RrLsCtg>= ": adc("wRrLsCtgTo").CmpStr = "RrLsCtg<= "
  adc("wRrLabMCdFr").CmpStr = "RrLabMCd>= ": adc("wRrLabMCdTo").CmpStr = "RrLabMCd<= "
  
  ms_Cnd = adc.RepCond
  ms_Cnd = IIF(ms_Cnd <> "", " where " + ms_Cnd, "")
  ' Zubin 212 (Before 06/01/06)
  'ms_Str = "Select * from RmRt " + ms_Cnd + " Order By RrTcTyp, RrCmCtg, RrCmCd, RrCmCurCd, RrCtg, RrSctg, RrSTWGrp, RrCd, RrDmCtg, RrLsCtg, RrLabMCd, RrFrLn, RrToLn"
  
  Dim ws_NewFlds As String
  
  '*** (Jen 2.12)
  If adc("wTgtVer") >= "2.11.0" And gs_Ver >= "2.11.0" And adc("wTgtVer") <> "2.6.0" Then
    ws_NewFlds = ws_NewFlds + ", RrFixMinTol, RrFixMaxTol, RrOpnMinTol, RrOpnMaxTol "
  End If
  '*** (Jen 2.12)
  
  '****** (Jen 2.12) ******
  If adc("wTgtVer") >= "2.12.0" And gs_Ver >= "2.12.0" And adc("wTgtVer") <> "2.6.0" Then
    ws_NewFlds = ws_NewFlds + ", RrDmCd "
  End If
  ms_Str = "Select RrTcTyp, RrCmCd, RrCtg, RrSctg, RrCd, RrDmCtg, RrLabMCd, RrFrLn, RrToLn, RrCstRt, RrSalRt, " + _
          "ModUsr, ModDt, ModTime, RrRngSz, RrCmCtg, RrLsCtg, RrSTWGrp, RrCmCurCd, RrXWtPer " + ws_NewFlds + _
          " from RmRt " + ms_Cnd + _
          " Order By RrTcTyp, RrCmCtg, RrCmCd, RrCmCurCd, RrCtg, RrSctg, RrSTWGrp, RrCd, RrDmCtg, RrLsCtg, " + _
          "RrLabMCd, RrFrLn, RrToLn"
  
  Call CreateXmlDocMst("RmRt", ms_Str)
  
  adc("wRrModDtFr").CmpStr = "": adc("wRrTcTyp").CmpStr = ""
  adc("wRrCmCtgFr").CmpStr = "": adc("wRrCmCtgTo").CmpStr = ""
  adc("wRrCmCdFr").CmpStr = "": adc("wRrCmCdTo").CmpStr = ""
  adc("wRrCmCurCdFr").CmpStr = "": adc("wRrCmCurCdTo").CmpStr = ""
  adc("wRrCtgFr").CmpStr = "": adc("wRrCtgTo").CmpStr = ""
  adc("wRrSCtgFr").CmpStr = "": adc("wRrSCtgTo").CmpStr = ""
  adc("wRrCdFr").CmpStr = "": adc("wRrCdTo").CmpStr = ""
  adc("wRrSTWGrpFr").CmpStr = "": adc("wRrSTWGrpTo").CmpStr = ""
  adc("wRrDmCtgFr").CmpStr = "": adc("wRrDmCtgTo").CmpStr = ""
  adc("wRrLsCtgFr").CmpStr = "": adc("wRrLsCtgTo").CmpStr = ""
  adc("wRrLabMCdFr").CmpStr = "": adc("wRrLabMCdTo").CmpStr = ""
  
  MWLib.EndProcess Me
  
  If FraRmRt.Visible = True Then
    Call CmdRmRt_Click
  End If
End Sub
Private Sub CmdLabRtHist_Click()
  Call DispFra(LabRtHist)
  If FraLabRtHist.Visible = True Then
    If adc("wLhModDtFr") = "" Then adc("wLhModDtFr") = "01/01/80"
  End If
End Sub
Private Sub CmdLabRtHistTrfOut_Click()
  MWLib.BeginProcess Me, "Labour Rate History Transfer Out is In Progress"
  
  adc("wLhModDtFr").CmpStr = "ModDt>= "
  adc("wLhCmCtgFr").CmpStr = "LhCmCtg>= ": adc("wLhCmCtgTo").CmpStr = "LhCmCtg<= "
  adc("wLhCmCdFr").CmpStr = "LhCmCd>= ": adc("wLhCmCdTo").CmpStr = "LhCmCd<= "
  adc("wLhCmCurCdFr").CmpStr = "LhCmCurCd>= ": adc("wLhCmCurCdTo").CmpStr = "LhCmCurCd<= "
  adc("wLhMCdFr").CmpStr = "LhMCd>= ": adc("wLhMCdTo").CmpStr = "LhMCd<= "
  adc("wLhSCdFr").CmpStr = "LhSCd>= ": adc("wLhSCdTo").CmpStr = "LhSCd<= "
  
  ms_Cnd = adc.RepCond
  ms_Cnd = IIF(ms_Cnd <> "", " where " + ms_Cnd, "")
  ms_Str = "Select * from LabRtHist " + ms_Cnd + " Order By LhCmCtg, LhCmCd, LhCmCurCd, LhMCd, LhSCd"
  Call CreateXmlDocMst("LabRtHist", ms_Str)
  
  adc("wLhModDtFr").CmpStr = ""
  adc("wLhCmCtgFr").CmpStr = "": adc("wLhCmCtgTo").CmpStr = ""
  adc("wLhCmCdFr").CmpStr = "": adc("wLhCmCdTo").CmpStr = ""
  adc("wLhCmCurCdFr").CmpStr = "": adc("wLhCmCurCdTo").CmpStr = ""
  adc("wLhMCdFr").CmpStr = "": adc("wLhMCdTo").CmpStr = ""
  adc("wLhSCdFr").CmpStr = "": adc("wLhSCdTo").CmpStr = ""
  
  MWLib.EndProcess Me
  
  If FraLabRtHist.Visible = True Then
    Call CmdLabRtHist_Click
  End If
End Sub
Private Sub CmdRmRtHist_Click()
  Call DispFra(RmRtHist)
  If FraRmRtHist.Visible = True Then
    If adc("wRhModDtFr") = "" Then adc("wRhModDtFr") = "01/01/80"
  End If
End Sub
Private Sub CmdRmRtHistTrfOut_Click()
  If adc("wRhTcTyp") = "" Then DispMsg "Invalid Raw Material Rate Type", etError: Exit Sub

  Call atxt_FldChange(0, "wRhYyyyMm", mf_Cancel, ms_ErrMsg, adc("wRhYyyyMm"))
  If mf_Cancel = True Then DispMsg ms_ErrMsg, etError: Exit Sub

  MWLib.BeginProcess Me, "RM Rate History Tranfer Out is In Progress"
  
  adc("wRhModDtFr").CmpStr = "ModDt>= ": adc("wRhTcTyp").CmpStr = "RhTcTyp= "
  adc("wRhCmCtgFr").CmpStr = "RhCmCtg>= ": adc("wRhCmCtgTo").CmpStr = "RhCmCtg<= "
  adc("wRhCmCdFr").CmpStr = "RhCmCd>= ": adc("wRhCmCdTo").CmpStr = "RhCmCd<= "
  adc("wRhCmCurCdFr").CmpStr = "RhCmCurCd>= ": adc("wRhCmCurCdTo").CmpStr = "RhCmCurCd<= "
  adc("wRhCtgFr").CmpStr = "RhCtg>= ": adc("wRhCtgTo").CmpStr = "RhCtg<= "
  adc("wRhSCtgFr").CmpStr = "RhSCtg>= ": adc("wRhSCtgTo").CmpStr = "RhSCtg<= "
  adc("wRhCdFr").CmpStr = "RhCd>= ": adc("wRhCdTo").CmpStr = "RhCd<= "
  adc("wRhSTWGrpFr").CmpStr = "RhSTWGrp>= ": adc("wRhSTWGrpTo").CmpStr = "RhSTWGrp<= "
  adc("wRhDmCtgFr").CmpStr = "RhDmCtg>= ": adc("wRhDmCtgTo").CmpStr = "RhDmCtg<= "
  adc("wRhLsCtgFr").CmpStr = "RhLsCtg>= ": adc("wRhLsCtgTo").CmpStr = "RhLsCtg<= "
  adc("wRhLabMCdFr").CmpStr = "RhLabMCd>= ": adc("wRhLabMCdTo").CmpStr = "RhLabMCd<= "
  
  ms_Cnd = adc.RepCond
  ms_Cnd = IIF(ms_Cnd <> "", " where " + ms_Cnd, "")
  ms_Str = "Select * from RmRtHist " + ms_Cnd + " Order By RhTcTyp, RhCmCtg, RhCmCd, RhCmCurCd, RhCtg, RhSCtg, RhSTWGrp, RhCd, RhDmCtg, RhLsCtg, RhLabMCd, RhFrLn, RhToLn"
  Call CreateXmlDocMst("RmRtHist", ms_Str)
  
  adc("wRhModDtFr").CmpStr = "": adc("wRhTcTyp").CmpStr = ""
  adc("wRhCmCtgFr").CmpStr = "": adc("wRhCmCtgTo").CmpStr = ""
  adc("wRhCmCdFr").CmpStr = "": adc("wRhCmCdTo").CmpStr = ""
  adc("wRhCmCurCdFr").CmpStr = "": adc("wRhCmCurCdTo").CmpStr = ""
  adc("wRhCtgFr").CmpStr = "": adc("wRhCtgTo").CmpStr = ""
  adc("wRhSCtgFr").CmpStr = "": adc("wRhSCtgTo").CmpStr = ""
  adc("wRhCdFr").CmpStr = "": adc("wRhCdTo").CmpStr = ""
  adc("wRhSTWGrpFr").CmpStr = "": adc("wRhSTWGrpTo").CmpStr = ""
  adc("wRhDmCtgFr").CmpStr = "": adc("wRhDmCtgTo").CmpStr = ""
  adc("wRhLsCtgFr").CmpStr = "": adc("wRhLsCtgTo").CmpStr = ""
  adc("wRhLabMCdFr").CmpStr = "": adc("wRhLabMCdTo").CmpStr = ""
  
  MWLib.EndProcess Me
  
  If FraRmRtHist.Visible = True Then
    Call CmdRmRtHist_Click
  End If
End Sub
Private Sub CmdDtTbl_Click()
  Call DispFra(DtTbl)
  If FraDtTbl.Visible = True Then
    If adc("wDtModDtFr") = "" Then adc("wDtModDtFr") = "01/01/80"
    If adc("wDtDtFr") = "" Then adc("wDtDtFr") = "01/01/80"
    If adc("wDtDtTo") = "" Then adc("wDtDtTo") = "01/01/80"
  End If
End Sub
Private Sub CmdDtTblTrfOut_Click()
  If adc("wDtTag") = "" Then DispMsg "Invalid Date Tag", etError: Exit Sub

  MWLib.BeginProcess Me, "Date Table Tranfer Out is In Progress"
  
  adc("wDtModDtFr").CmpStr = "ModDt>= ": adc("wDtTag").CmpStr = "DtTag= "
  adc("wDtDtFr").CmpStr = "DtDt>= ": adc("wDtDtTo").CmpStr = "DtDt<= "
  
  ms_Cnd = adc.RepCond
  ms_Cnd = IIF(ms_Cnd <> "", " where " + ms_Cnd, "")
  ' Zubin 213 (Bef 213)
  'ms_Str = "Select * from DtTbl " + ms_Cnd + " Order By DtDt"
  ' Zubin 213
  ms_Str = "Select DtTag, DtDt, DtVal, ModUsr, ModDt, ModTime, DtHolYN " + _
            IIF(ValidateCurrVer(2, 13) And ValidateTgtVer(2, 13), ", DtCurCd1, DtCurCd2 ", "") + _
            " from DtTbl " + ms_Cnd + " Order By DtDt "
  ' Zubin 213

  Call CreateXmlDocMst("DtTbl", ms_Str, adc("wDtTag"))
  
  adc("wDtModDtFr").CmpStr = "": adc("wDtTag").CmpStr = ""
  adc("wDtDtFr").CmpStr = "": adc("wDtDtTo").CmpStr = ""
  
  MWLib.EndProcess Me
  
  If FraDtTbl.Visible = True Then
    Call CmdDtTbl_Click
  End If
End Sub
Private Sub CmdLoc_Click()
  Call DispFra(Loc)
  If FraLoc.Visible = True Then
    If adc("wLocModDtFr") = "" Then adc("wLocModDtFr") = "01/01/80"
  End If
End Sub
Private Sub CmdLocTrfOut_Click()
  If adc("wLocTyp") = "" Then DispMsg "Invalid Location Type", etError: Exit Sub

  MWLib.BeginProcess Me, "Location Master Tranfer Out is In Progress"
  
  adc("wLocModDtFr").CmpStr = "ModDt>= ": adc("wLocTyp").CmpStr = "LocTyp= "
  adc("wLocCdFr").CmpStr = "LocCd>= ": adc("wLocCdTo").CmpStr = "LocCd<= "
  
  ms_Cnd = adc.RepCond
  ms_Cnd = IIF(ms_Cnd <> "", " and " + ms_Cnd, "")
  
  If adc("wTgtVer") = "2.6.0" Then
  'If ms_TrfOutVer = "2.6.0" Then
    ms_Str = "Select LocCd, LocDesc, LocTyp, LocPrdTyp, LocLoss, LocBrk, LocMsg, ModUsr, " + _
             "ModDt, ModTime, LocCoCd, LocPrdStg, LocValidYN from Loc " + _
             "where LocCoCd= '" + gs_CoCd + "'" + ms_Cnd + " Order By LocCd"
  Else
    'If ms_TrfOutVer = "2.9.0" Then
    ms_Str = "Select LocCd, LocDesc, LocTyp, LocPrdTyp, LocLoss, LocBrk, LocMsg, ModUsr, " + _
             "ModDt, ModTime, LocCoCd, LocPrdStg, LocValidYN, " + _
             IIF(gs_Ver = "2.6.0" Or gs_Ver = "2.7.0", " '' as LocLsTyp ", " LocLsTyp ") + " from Loc " + _
             "where LocCoCd= '" + gs_CoCd + "'" + ms_Cnd + " Order By LocCd"
  End If
  Call CreateXmlDocMst("Loc", ms_Str, adc("wLocTyp"))
  
  adc("wLocModDtFr").CmpStr = "": adc("wLocTyp").CmpStr = ""
  adc("wLocCdFr").CmpStr = "": adc("wLocCdTo").CmpStr = ""
  
  MWLib.EndProcess Me
  
  If FraLoc.Visible = True Then
    Call CmdLoc_Click
  End If
End Sub
Private Sub CmdPrm_Click()
  Call DispFra(Prm)
  If FraPrm.Visible = True Then
    If adc("wPModDtFr") = "" Then adc("wPModDtFr") = "01/01/80"
  End If
End Sub
Private Sub CmdPrmTrfOut_Click()
  If adc("wPCtg") = "" Then DispMsg "Specify Whether The Transfer Is For A System Or Other Parameter", etError: Exit Sub

  If adc("wPTyp") = "" Then DispMsg "Enter Valid Param Type", etError: Exit Sub
  Call atxt_FldChange(0, "wPTyp", mf_Cancel, ms_ErrMsg, adc("wPTyp"))
  If mf_Cancel = True Then DispMsg ms_ErrMsg, etError: Exit Sub

  MWLib.BeginProcess Me, "Parameter Master Tranfer Out is In Progress"
  
  adc("wPModDtFr").CmpStr = "ModDt>= ": adc("wPTyp").CmpStr = "PTyp= "
  adc("wPMCdFr").CmpStr = "PMCd>= ": adc("wPMCdTo").CmpStr = "PMCd<= "
  adc("wPSCdFr").CmpStr = "PSCd>= ": adc("wPSCdTo").CmpStr = "PSCd<= "
  
  ms_Cnd = adc.RepCond
  ms_Cnd = IIF(ms_Cnd <> "", " where " + ms_Cnd, "")
  ms_Str = "Select * from Param " + ms_Cnd + " Order By PMCd, PSCd"
  Call CreateXmlDocMst("Param", ms_Str, adc("wPTyp"))
  
  adc("wPModDtFr").CmpStr = "": adc("wPTyp").CmpStr = ""
  adc("wPMCdFr").CmpStr = "": adc("wPMCdTo").CmpStr = ""
  adc("wPSCdFr").CmpStr = "": adc("wPSCdTo").CmpStr = ""
  
  MWLib.EndProcess Me
  
  If FraPrm.Visible = True Then
    Call CmdPrm_Click
  End If
End Sub
Private Sub CmdVPrm_Click()
  Call DispFra(VPrm)
  If FraVPrm.Visible = True Then
    If adc("wvPModDtFr") = "" Then adc("wvPModDtFr") = "01/01/80"
  End If
End Sub
Private Sub CmdVPrmTrfOut_Click()
  If adc("wVPCtg") = "" Then DispMsg "Specify Whether The Transfer Is For A System Or Other Variable Parameter", etError: Exit Sub

  If adc("wVPTyp") = "" Then DispMsg "Enter Valid Variable Param Type", etError: Exit Sub
  Call atxt_FldChange(0, "wVPTyp", mf_Cancel, ms_ErrMsg, adc("wVPTyp"))
  If mf_Cancel = True Then DispMsg ms_ErrMsg, etError: Exit Sub

  MWLib.BeginProcess Me, "Variable Parameter Master Tranfer Out is In Progress"
  
  adc("wVPModDtFr").CmpStr = "ModDt>= ": adc("wVPTyp").CmpStr = "vPTyp= "
  adc("wVPMCdFr").CmpStr = "vPMCd>= ": adc("wVPMCdTo").CmpStr = "vPMCd<= "
  adc("wVPSCdFr").CmpStr = "vPSCd>= ": adc("wVPSCdTo").CmpStr = "vPSCd<= "
  
  ms_Cnd = adc.RepCond
  ms_Cnd = IIF(ms_Cnd <> "", " and " + ms_Cnd, "")
  If adc("wTgtVer") = "2.6.0" Then
  'If ms_TrfOutVer = "2.6.0" Then
    ms_Str = "Select vPCoCd, vPTyp, vPMCd, vPSCd, vPDesc, vPDesc225, vPValue, vPNum, " + _
             "vPValue1, vPNum1, vPValue2, vPValue3, ModUsr, ModDt, ModTime, vPValidYn " + _
             "from vParam where vPCoCd= '" + gs_CoCd + "' " + ms_Cnd + " Order By vPMCd, vPSCd"
  Else
    'If ms_TrfOutVer = "2.9.0" Then
    ms_Str = "Select vPCoCd, vPTyp, vPMCd, vPSCd, vPDesc, vPDesc225, vPValue, vPNum, " + _
             "vPValue1, vPNum1, vPValue2, vPValue3, ModUsr, ModDt, ModTime, vPValidYn, " + _
             IIF(gs_Ver = "2.6.0", " '' as vPValue4 ", " vPValue4 ") + _
             "from vParam where vPCoCd= '" + gs_CoCd + "' " + ms_Cnd + " Order By vPMCd, vPSCd"
  End If
  Call CreateXmlDocMst("vParam", ms_Str, adc("wVPTyp"))
  
  adc("wVPModDtFr").CmpStr = "": adc("wVPTyp").CmpStr = ""
  adc("wVPMCdFr").CmpStr = "": adc("wVPMCdTo").CmpStr = ""
  adc("wVPSCdFr").CmpStr = "": adc("wVPSCdTo").CmpStr = ""
  
  MWLib.EndProcess Me
  
  If FraVPrm.Visible = True Then
    Call CmdVPrm_Click
  End If
End Sub
Private Sub CmdOrd_Click()
  Call DispFra(Ord)
  If FraOrd.Visible = True Then
    If adc("wOmModDtFr") = "" Then adc("wOmModDtFr") = "01/01/80"
  End If
End Sub
Private Sub CmdOrdTrfOut_Click()
  If adc("wOmTcTyp") = "" Then DispMsg "Invalid Order Tc Type", etError: Exit Sub

  MWLib.BeginProcess Me, "Order Tranfer Out is In Progress"
  
  adc("wOmModDtFr").CmpStr = "ModDt>= "
  adc("wOmTcFr").CmpStr = "OmTc>= ": adc("wOmTcTo").CmpStr = "OmTc<= "
  adc("wOmYyFr").CmpStr = "OmYy>= ": adc("wOmYyTo").CmpStr = "OmYy<= "
  adc("wOmChrFr").CmpStr = "OmChr>= ": adc("wOmChrTo").CmpStr = "OmChr<= "
  adc("wOmNoFr").CmpStr = "OmNo>= ": adc("wOmNoTo").CmpStr = "OmNo<= "
  adc("wOmCmCdFr").CmpStr = "OmCmCd>= ": adc("wOmCmCdTo").CmpStr = "OmCmCd<= "

  Call CreateXmlDocOrd
  
  adc("wOmModDtFr").CmpStr = ""
  adc("wOmTcFr").CmpStr = "": adc("wOmTcTo").CmpStr = ""
  adc("wOmYyFr").CmpStr = "": adc("wOmYyTo").CmpStr = ""
  adc("wOmChrFr").CmpStr = "": adc("wOmChrTo").CmpStr = ""
  adc("wOmNoFr").CmpStr = "": adc("wOmNoTo").CmpStr = ""
  adc("wOmCmCdFr").CmpStr = "": adc("wOmCmCdTo").CmpStr = ""
  
  MWLib.EndProcess Me
  
  If FraOrd.Visible = True Then
    Call CmdOrd_Click
  End If
End Sub
Private Sub CmdInv_Click()
  Call DispFra(Inv)
  If FraInv.Visible = True Then
    If adc("wInModDtFr") = "" Then adc("wInModDtFr") = "01/01/80"
  End If
End Sub
Private Sub CmdInvTrfOut_Click()
  'Before 030604
  'ADC("wInModDtFr").CmpStr = "ModDt>= "
  'Zubin 030604
  MWLib.BeginProcess Me, "Invoice Tranfer Out is In Progress"
  
  adc("wInModDtFr").CmpStr = "InvHd.ModDt>= "
  adc("wInTcFr").CmpStr = "InTc>= ": adc("wInTcTo").CmpStr = "InTc<= "
  adc("wInYyFr").CmpStr = "InYy>= ": adc("wInYyTo").CmpStr = "InYy<= "
  adc("wInChrFr").CmpStr = "InChr>= ": adc("wInChrTo").CmpStr = "InChr<= "
  adc("wInNoFr").CmpStr = "InNo>= ": adc("wInNoTo").CmpStr = "InNo<= "
  adc("wInCmCdFr").CmpStr = "InCmCd>= ": adc("wInCmCdTo").CmpStr = "InCmCd<= "
  adc("wInDtFr").CmpStr = "InDt>= ": adc("wInDtTo").CmpStr = "InDt<= "
 
  Call CreateXmlDocInv
  
  adc("wOmModDtFr").CmpStr = ""
  adc("wInTcFr").CmpStr = "": adc("wInTcTo").CmpStr = ""
  adc("wInYyFr").CmpStr = "": adc("wInYyTo").CmpStr = ""
  adc("wInChrFr").CmpStr = "": adc("wInChrTo").CmpStr = ""
  adc("wInNoFr").CmpStr = "": adc("wInNoTo").CmpStr = ""
  adc("wInCmCdFr").CmpStr = "": adc("wInCmCdTo").CmpStr = ""
  adc("wInDtFr").CmpStr = "": adc("wInDtTo").CmpStr = ""
  
  MWLib.EndProcess Me
  
  If FraInv.Visible = True Then
    Call CmdInv_Click
  End If
End Sub
Private Sub CmdAllMst_Click()
  Call DispFra(AllMst)
  If FraAllMst.Visible = True Then
    If adc("wAllMstModDtFr") = "" Then adc("wAllMstModDtFr") = moCn.SrvrDate
    If adc("wAllMstModDtTo") = "" Then adc("wAllMstModDtTo") = "01/01/80"
    If TxtFolderPath.text = "" Then TxtFolderPath.text = moCn.GetFldVal("Select PDesc225 from Param where PTyp= 'TRFDIR' and PMCd= 'TRFDIR' ")
    TreeViewSel.ImageList = ImgList
    RefreshTree
  End If
End Sub
Sub RefreshTree()
    Set Explorer = New CExplorer2
    Explorer.Init TreeViewSel, , 1, 2, 3
    
    On Error Resume Next
    
    '** add the "C:\" root
    Dim rootDir As New CDirectory2
    rootDir.Path = Mid(DriveLst.Drive, 1, InStr(1, DriveLst.Drive, ":"))
    Set Explorer.Root = rootDir
End Sub
Private Sub DriveLst_Change()
  RefreshTree
End Sub
Private Sub TreeViewSel_NodeClick(ByVal Node As ComctlLib.Node)
  TxtFolderPath.text = TreeViewSel.SelectedItem.FullPath
End Sub
Private Sub CmdAllMstTrfOut_Click()
  If Trim(CStr(adc("wAllMstModDtFr"))) = "" Or CStr(adc("wAllMstModDtFr")) = "01/01/80" Then
    DispMsg "From Modified Date Is Compulsory", etError: Exit Sub
  End If
  
  If Trim(TxtFolderPath.text) = "" Then
    adc("wDisp") = "Enter Proper Folder Path"
    DispMsg "Enter Proper Folder Path", etError: Exit Sub
  End If
  
  '***###
  Dim wo_FileSysObj, wo_Folder, wo_FolderColl, wo_NewFolder
  'Dim ws_FolderPath As String

  On Error Resume Next
  'ws_FolderPath = TxtFolderPath.Text
  Set wo_FileSysObj = CreateObject("Scripting.FileSystemObject")
  Set wo_Folder = wo_FileSysObj.GetFolder(TxtFolderPath.text)
  
  TxtNewFolderName.text = Trim(TxtNewFolderName.text)
  '*** If No Such Folder Exists Then Create One.
  If TxtNewFolderName.text <> "" Then
    Set wo_FolderColl = wo_Folder.SubFolders
    Set wo_NewFolder = wo_FolderColl.Add(TxtNewFolderName.text)
   
    If Err <> 0 Then
      adc("wDisp") = "Folder Already Exists/ Folder Name Is Invalid"
      DispMsg "Folder Already Exists/ Folder Name Is Invalid", etError: Exit Sub
      
      'Dim Response
      'Response = MsgBox("Folder Does Not Exist. Do You Want To Create The Folder?", vbYesNo)
      'If Response = vbYes Then
      '  '*** Create Folder
      'Else
      '  '*** User chose No.
      '  Exit Sub
      'End If
    End If
    On Error GoTo 0
   
    TxtFolderPath.text = TxtFolderPath.text + "\" + TxtNewFolderName.text
  End If
  
  If UCase(wo_Folder.Type) <> UCase("File Folder") Then
    adc("wDisp") = "Specified File Is Not A Folder"
    DispMsg "Specified File Is Not A Folder", etError: Exit Sub
  End If
  '***###
  
  MWLib.BeginProcess Me, "All Masters Tranfer Out is In Progress"
  
  adc("wAllMstModDtFr").CmpStr = "ModDt>= "
  adc("wAllMstModDtTo").CmpStr = "ModDt<= "
    
  Call CreateXmlDocAllMst
  
  adc("wAllMstModDtFr").CmpStr = ""
  adc("wAllMstModDtTo").CmpStr = ""

  MWLib.EndProcess Me
  
  Call CmdAllMst_Click
End Sub
Private Sub CreateXmlDocAllMst()
  Dim Rs_Scope As MwfLib.MDORowSet
  
  mf_AllMst = True
  
  Call VerMatch
  
  Set Rs_Scope = moCn.OpenRes("Select PMCd from Param where PTyp= 'BKCTG'")
  Do While Not (Rs_Scope.EOF Or Rs_Scope.BOF)
    adc("wBkCtg") = Rs_Scope!PMCd
    Call CmdBkAgTrfOut_Click
    Rs_Scope.MoveNext
  Loop
  
  Set Rs_Scope = moCn.OpenRes("Select PMCd from Param where PTyp= 'YN'")
  Do While Not (Rs_Scope.EOF Or Rs_Scope.BOF)
    adc("wRmZ") = Rs_Scope!PMCd
    Call CmdRmTrfOut_Click
    Rs_Scope.MoveNext
  Loop
 
  Set Rs_Scope = moCn.OpenRes("Select PMCd from Param where PTyp= 'CMCTG'")
  adc("wCmCpyChild") = "Y"
  Do While Not (Rs_Scope.EOF Or Rs_Scope.BOF)
    adc("wCmCtg") = Rs_Scope!PMCd
    Call CmdCustTrfOut_Click
    Rs_Scope.MoveNext
  Loop
 
  Set Rs_Scope = moCn.OpenRes("Select PMCd from Param where PTyp= 'DMTCTYP'")
  Do While Not (Rs_Scope.EOF Or Rs_Scope.BOF)
    adc("wDmTcTyp") = Rs_Scope!PMCd
    Call CmdDsgTrfOut_Click
    Rs_Scope.MoveNext
  Loop
 
  Set Rs_Scope = moCn.OpenRes("Select PMCd from Param where PTyp= 'DTTAG'")
  Do While Not (Rs_Scope.EOF Or Rs_Scope.BOF)
    adc("wDtTag") = Rs_Scope!PMCd
    Call CmdDtTblTrfOut_Click
    Rs_Scope.MoveNext
  Loop
 
  Set Rs_Scope = moCn.OpenRes("Select PMCd from Param where PTyp= 'LOCTYP'")
  Do While Not (Rs_Scope.EOF Or Rs_Scope.BOF)
    adc("wLocTyp") = Rs_Scope!PMCd
    Call CmdLocTrfOut_Click
    Rs_Scope.MoveNext
  Loop
  
  Set Rs_Scope = moCn.OpenRes("Select PMCd, PValue from Param where PTyp= 'PTYP' and PValue in ('Y', 'N') Order By PValue, PMCd")
  Do While Not (Rs_Scope.EOF Or Rs_Scope.BOF)
    adc("wPCtg") = Rs_Scope!pValue
    adc("wPTyp") = Rs_Scope!PMCd
    Call CmdPrmTrfOut_Click
    Rs_Scope.MoveNext
  Loop
 
  Set Rs_Scope = moCn.OpenRes("Select vPMCd, vPValue from vParam where vPTyp= 'PTYP' and vPValue in ('Y', 'N') Order By vPValue, vPMCd")
  Do While Not (Rs_Scope.EOF Or Rs_Scope.BOF)
    adc("wVPCtg") = Rs_Scope!vPValue
    adc("wVPTyp") = Rs_Scope!vPMCd
    Call CmdVPrmTrfOut_Click
    Rs_Scope.MoveNext
  Loop
 
  'Call CmdLabRtTrfOut_Click
  'Call CmdRmRtTrfOut_Click
  'Call CmdLabRtHistTrfOut_Click
  'Call CmdRmRtHistTrfOut_Click
  
  adc("wDisp") = "Transfer Out Of All Masters Completed Successfully"
  DispMsg "Transfer Out Of All Masters Completed Successfully", etInfo
  
  mf_AllMst = False
  Set Rs_Scope = Nothing
End Sub
Private Sub CreateXmlDocMst(ms_XTbl As String, ms_XStr As String, Optional ByVal ms_Identification As String)
  Dim ws_Identification As String
  
  Call VerMatch

  ws_Identification = ""
  If Not IsMissing(ms_Identification) Then ws_Identification = "_" + ms_Identification
  
  Set Rs_SrcX = moCn.OpenRes(ms_XStr)
  If Rs_SrcX.RecCount <= 0 Then
    If mf_AllMst = False Then DispMsg "No Record Exists For This Scope", etError
    Exit Sub
  End If
  adc("wDisp") = ms_XTbl + " Transfer Out In Process ..."
  
  Set XmlDoc = New MSXML2.DOMDocument
  Set xmlRootElement = XmlDoc.createElement("Emr" + ms_XTbl + "Transfer")
  XmlDoc.appendChild xmlRootElement

  xmlRootElement.setAttribute "Ver", adc("wTgtVer")
  'xmlRootElement.setAttribute "Ver", ms_TrfOutVer
   
  Do While Not (Rs_SrcX.EOF Or Rs_SrcX.BOF)
    '### ADC("wDisp") = ms_XTbl + " : " + Rs_SrcX!CmCd
    Set xmlChildX = XmlDoc.createElement(ms_XTbl)
    Call WriteDocument(xmlChildX, Rs_SrcX)
    xmlRootElement.appendChild xmlChildX
    Rs_SrcX.MoveNext
  Loop
  
  Call SaveDocument(XmlDoc, ms_XTbl + ws_Identification)
  adc("wDisp") = ms_XTbl + " Transfer Out Completed"
  Set Rs_SrcX = Nothing
End Sub
Private Sub CreateXmlDocRm(Optional ByVal ms_Identification As String)
  Dim ws_Identification As String
  
  Call VerMatch
  
  ws_Identification = ""
  If Not IsMissing(ms_Identification) Then ws_Identification = "_" + ms_Identification
  
  ms_Cnd = adc.RepCond
  ms_Cnd = IIF(ms_Cnd <> "", " where " + ms_Cnd, "")

  '*** 2.11 (02/08/05)
  If adc("wRmModDtFr") <> "01/01/80" Then
    ms_Cnd = ms_Cnd + IIF(ms_Cnd <> "", " and ", " where ") + " (ModDt>= '" + CStr(adc("wRmModDtFr")) + "' or " + _
             "exists(Select 'x' from RmAlyCmp Ra where RaCd= RmCd and Ra.ModDt>= '" + CStr(adc("wRmModDtFr")) + "')) "
    ms_ChildModDtCnd = " and ModDt >= '" + CStr(adc("wRmModDtFr")) + "' "
  Else
    ms_ChildModDtCnd = ""
  End If
  '*** 2.11 (02/08/05)
  
  ' Zubin 212 (Before 06/01/06)
  'Set Rs_SrcX = moCn.OpenRes("Select * from RmMst " + ms_Cnd + " Order By RmCd")
  ' Zubin 212
  '****** (Jen 2.13) ******
  If adc("wTgtVer") >= "2.12.0" And gs_Ver >= "2.12.0" And adc("wTgtVer") <> "2.6.0" Then
    Dim ws_NewFlds As String
    ws_NewFlds = ", RmId, RmFndWt "
  End If
  
  ' Zubin 213
  'If (mi_TgtMajVer > 2 Or (mi_TgtMajVer = 2 And mi_TgtMinVer >= 13)) And _
      (mi_MajVer > 2 Or (mi_MajVer = 2 And mi_MinVer >= 13)) Then 'And ADC("wTgtVer") <> "2.6.0" Then
  If ValidateCurrVer(2, 13) = True And ValidateTgtVer(2, 13) = True Then
    ws_NewFlds = ws_NewFlds + ", RmQwCst "
  End If
  
  Set Rs_SrcX = moCn.OpenRes("Select RmCd, RmCtg, RmSCtg, RmDesc, RmSpGrav, RmIntQly, RmCmCd, RmPurityRt, RmUom, " + _
                  "RmCol, RmKt, RmZ, RmZCd, RmPurityWt, ModUsr, ModDt, ModTime, RmPurityZ, RmSTWGrp, RmQw, " + _
                  "RmBaseCd, RmBasePurityWt, RmParentCd, RmPrtDesc, RmValidYN " + ws_NewFlds + " from RmMst " + _
                  ms_Cnd + " Order By RmCd")
  If Rs_SrcX.RecCount <= 0 Then
    If mf_AllMst = False Then DispMsg "No Record Exists For This Scope", etError
    Exit Sub
  End If
  adc("wDisp") = "Raw Material Master Transfer Out In Process ..."

  Set XmlDoc = New MSXML2.DOMDocument
  Set xmlRootElement = XmlDoc.createElement("EmrRmMstTransfer")
  XmlDoc.appendChild xmlRootElement
  xmlRootElement.setAttribute "Ver", adc("wTgtVer")
  'xmlRootElement.setAttribute "Ver", ms_TrfOutVer

  Do While Not (Rs_SrcX.EOF Or Rs_SrcX.BOF)
    adc("wDisp") = "Rm : " + Rs_SrcX!RmCd
    Set xmlChildX = XmlDoc.createElement("RmMst")
    Call WriteDocument(xmlChildX, Rs_SrcX)
    xmlRootElement.appendChild xmlChildX
    
    Set Rs_SrcA = moCn.OpenRes("Select * from RmAlyCmp where " + _
                  "RaCd= '" + Rs_SrcX!RmCd + "' " + ms_ChildModDtCnd + " Order By RaCmpCd")
    Do While Not (Rs_SrcA.EOF Or Rs_SrcA.BOF)
      Set xmlChildA = XmlDoc.createElement("RmAlyCmp")
      Call WriteDocument(xmlChildA, Rs_SrcA)
      xmlChildX.appendChild xmlChildA
      Rs_SrcA.MoveNext
    Loop
    
    Rs_SrcX.MoveNext
  Loop
  
  Call SaveDocument(XmlDoc, "RmMst" + ws_Identification)
  adc("wDisp") = "Raw Material Transfer Out Completed"
  Set Rs_SrcX = Nothing
End Sub
Private Sub CreateXmlDocCust(Optional ByVal ms_Identification As String)
  Dim ws_Identification As String
  
  Call VerMatch
  
  ws_Identification = ""
  If Not IsMissing(ms_Identification) Then ws_Identification = "_" + ms_Identification
  
  ms_Cnd = adc.RepCond
  ms_Cnd = IIF(ms_Cnd <> "", " where " + ms_Cnd, "")

  '*** 2.11 (02/08/05)
  If adc("wCmModDtFr") <> "01/01/80" Then
    ms_Cnd = ms_Cnd + IIF(ms_Cnd <> "", " and ", " where ") + " (ModDt>= '" + CStr(adc("wCmModDtFr")) + "' " + _
             "or exists(Select 'x' from CustDsg Cd where CdCtg= CmCtg and CdCd= CmCd " + _
             "   and Cd.ModDt>= '" + CStr(adc("wCmModDtFr")) + "') " + _
             "or exists(Select 'x' from CustAly Ca where CaCtg= CmCtg and CaCd= CmCd " + _
             "   and Ca.ModDt>= '" + CStr(adc("wCmModDtFr")) + "') " + _
             "or exists(Select 'x' from CustRm Cr where CrCtg= CmCtg and CrCd= CmCd " + _
             "   and Cr.ModDt>= '" + CStr(adc("wCmModDtFr")) + "') " + _
             ") "
    ms_ChildModDtCnd = " and ModDt >= '" + CStr(adc("wCmModDtFr")) + "' "
  Else
    ms_ChildModDtCnd = ""
  End If
  '*** 2.11 (02/08/05)

  If adc("wTgtVer") = "2.6.0" Then
  'If ms_TrfOutVer = "2.6.0" Then
    Set Rs_SrcX = moCn.OpenRes("Select CmCtg, CmCd, CmName, CmMulBy, CmFixPrc, CmGldAs, " + _
                  "CmIWtEqOrd, CmIWtFrOrd, CmIGldRtEqOrd, CmDiaRtFA, CmCSRtFA, CmPrdInst, " + _
                  "CmStmpInst, CmPayTerms, CmDelTerms, CmAdd1, CmAdd2, CmAdd3, CmTel, CmFax, " + _
                  "CmEMail, CmCpName, CmPortofDisch, CmFinDest, CmBuyer, ModUsr, " + _
                  "ModDt, ModTime, CmLkUpRmRt, CmLkUpLabRt, CmLkUpMetLs, CmRegnCd, CmHInsBy, " + _
                  "CmHRndOrdSalPrc, convert(varchar(5800), CmHistory) as CmHistory, CmCurCd, " + _
                  "CmSalPer, CmBuyAdd1, CmBuyAdd2, CmBuyAdd3, CmBkCd, CmAgCd, CmValidYN " + _
                  "from CustMst " + ms_Cnd + " Order By CmCd")
  Else
    ' Zubin 213 (CmLabAs Added)
    'If ms_TrfOutVer = "2.9.0" Then
    Set Rs_SrcX = moCn.OpenRes("Select CmCtg, CmCd, CmName, CmMulBy, CmFixPrc, CmGldAs, " + _
                  "CmIWtEqOrd, CmIWtFrOrd, CmIGldRtEqOrd, CmDiaRtFA, CmCSRtFA, CmPrdInst, " + _
                  "CmStmpInst, CmPayTerms, CmDelTerms, CmAdd1, CmAdd2, CmAdd3, CmTel, CmFax, " + _
                  "CmEMail, CmCpName, CmPortofDisch, CmFinDest, CmBuyer, ModUsr, " + _
                  "ModDt, ModTime, CmLkUpRmRt, CmLkUpLabRt, CmLkUpMetLs, CmRegnCd, CmHInsBy, " + _
                  "CmHRndOrdSalPrc, convert(varchar(5800), CmHistory) as CmHistory, CmCurCd, " + _
                  "CmSalPer, CmBuyAdd1, CmBuyAdd2, CmBuyAdd3, CmBkCd, CmAgCd, CmValidYN, " + _
                  IIF(gs_Ver = "2.6.0", " '' as CmVessel ", " CmVessel ") + _
                  IIF(ValidateCurrVer(2, 13) And ValidateTgtVer(2, 13), ", CmLabAs ", "") + _
                  " from CustMst " + _
                  ms_Cnd + " Order By CmCd")
  End If
                
  If Rs_SrcX.RecCount <= 0 Then
    If mf_AllMst = False Then DispMsg "No Record Exists For This Scope", etError
    Exit Sub
  End If
  adc("wDisp") = "Customer Master Transfer Out In Process ..."

  Set XmlDoc = New MSXML2.DOMDocument
  Set xmlRootElement = XmlDoc.createElement("EmrCustMstTransfer")
  XmlDoc.appendChild xmlRootElement
  xmlRootElement.setAttribute "Ver", adc("wTgtVer")
  'xmlRootElement.setAttribute "Ver", ms_TrfOutVer

  Do While Not (Rs_SrcX.EOF Or Rs_SrcX.BOF)
    adc("wDisp") = "Customer : " + Rs_SrcX!CmCd
    Set xmlChildX = XmlDoc.createElement("CustMst")
    Call WriteDocument(xmlChildX, Rs_SrcX)
    xmlRootElement.appendChild xmlChildX
    
    If adc("wCmCpyChild") = "Y" Then
      Set Rs_SrcA = moCn.OpenRes("Select * from CustDsg where " + _
                    "CdCtg= '" + Rs_SrcX!CmCtg + "' and CdCd= '" + Rs_SrcX!CmCd + "' " + ms_ChildModDtCnd + " " + _
                    " Order By CdOurDmCd, CdOurSfx, CdOurDmSz")
      Do While Not (Rs_SrcA.EOF Or Rs_SrcA.BOF)
        Set xmlChildA = XmlDoc.createElement("CustDsg")
        Call WriteDocument(xmlChildA, Rs_SrcA)
        xmlChildX.appendChild xmlChildA
        Rs_SrcA.MoveNext
      Loop
      
      Set Rs_SrcD = moCn.OpenRes("Select * from CustAly where " + _
                    "CaCtg= '" + Rs_SrcX!CmCtg + "' and CaCd= '" + Rs_SrcX!CmCd + "' " + ms_ChildModDtCnd + " " + _
                    " Order By CaRmCd, CaDmCol")
      Do While Not (Rs_SrcD.EOF Or Rs_SrcD.BOF)
        Set xmlChildD = XmlDoc.createElement("CustAly")
        Call WriteDocument(xmlChildD, Rs_SrcD)
        xmlChildX.appendChild xmlChildD
        Rs_SrcD.MoveNext
      Loop
      
      Set Rs_SrcG = moCn.OpenRes("Select * from CustRm where " + _
                    "CrCtg= '" + Rs_SrcX!CmCtg + "' and CrCd= '" + Rs_SrcX!CmCd + "' " + ms_ChildModDtCnd + " " + _
                    " Order By CrOurRmCd, CrFrLn, CrToLn")
      Do While Not (Rs_SrcG.EOF Or Rs_SrcG.BOF)
        Set xmlChildG = XmlDoc.createElement("CustRm")
        Call WriteDocument(xmlChildG, Rs_SrcG)
        xmlChildX.appendChild xmlChildG
        Rs_SrcG.MoveNext
      Loop
      'If Source and Target Versions are >= 2.13.0 then transfer out DsgModAvl
      
    End If
    
    Rs_SrcX.MoveNext
  Loop
  
  Call SaveDocument(XmlDoc, "CustMst" + ws_Identification)
  adc("wDisp") = "Customer Transfer Out Completed"
  Set Rs_SrcX = Nothing
End Sub
Private Sub CreateXmlDocDsg(Optional ByVal ms_Identification As String)
  Dim ws_Identification As String
  
  Call VerMatch
  
  ws_Identification = ""
  If Not IsMissing(ms_Identification) Then ws_Identification = "_" + ms_Identification
  
  ms_Cnd = adc.RepCond
  ms_Cnd = IIF(ms_Cnd <> "", " where " + ms_Cnd, "")

  '*** 2.11 (02/08/05)
  Dim ws_Cnd As String
  ' Zubin 212 (Before 06/01/06)
  'If (ADC("wTgtVer") = "2.10.0" Or ADC("wTgtVer") = "2.11.0") And gs_Ver = "2.11.0" Then
  ' Zubin 212
  If adc("wTgtVer") >= "2.10.0" And gs_Ver >= "2.11.0" Then
    ws_Cnd = " or exists(Select 'x' from DsgPrt Dp where DpTcTyp= DmTcTyp and DpCd= DmCd and DpSz= DmSz " + _
               "   and Dp.ModDt>= '" + CStr(adc("wDmModDtFr")) + "') "
  Else
    ws_Cnd = " "
  End If
  
  If adc("wDmModDtFr") <> "01/01/80" Then
    ' Zubin 213 (DsgAvl added)
    ms_Cnd = ms_Cnd + IIF(ms_Cnd <> "", " and ", " where ") + " (ModDt>= '" + CStr(adc("wDmModDtFr")) + "' " + _
             "or exists(Select 'x' from DsgRm Dr where DrTcTyp= DmTcTyp and DrCd= DmCd and DrSz= DmSz " + _
             "   and Dr.ModDt>= '" + CStr(adc("wDmModDtFr")) + "') " + _
             "or exists(Select 'x' from DsgLab Dl where DlTcTyp= DmTcTyp and DlCd= DmCd and DlSz= DmSz " + _
             "   and Dl.ModDt>= '" + CStr(adc("wDmModDtFr")) + "') " + _
             "or exists(Select 'x' from DsgAna Da where DaTcTyp= DmTcTyp and DaCd= DmCd and DaSz= DmSz " + _
             "   and Da.ModDt>= '" + CStr(adc("wDmModDtFr")) + "') " + _
             IIF(ValidateCurrVer(2, 13) And ValidateTgtVer(2, 13), _
             " or exists(Select 'x' from DsgModAvl Dma where DmaTcTyp= DmTcTyp and DmaCd= DmCd and DmaSz= DmSz " + _
             "and Dma.ModDt>= '" + CStr(adc("wDmModDtFr")) + "') ", "") + _
             ws_Cnd + ") "
             
    ms_ChildModDtCnd = " and ModDt >= '" + CStr(adc("wDmModDtFr")) + "' "
  Else
    ms_ChildModDtCnd = ""
  End If
  '*** 2.11 (02/08/05)
  
  ' Zubin 212 '****** (Jen 2.13) ******
  If adc("wTgtVer") >= "2.12.0" And gs_Ver >= "2.12.0" And adc("wTgtVer") <> "2.6.0" Then
    Dim ws_NewFlds As String
    ws_NewFlds = ", DmWaxWt, DmModRunWt, DmSilModWt, DmCasPcWt "
  End If
  
  Set Rs_SrcX = moCn.OpenRes("Select DmTcTyp, DmCd, DmSz, DmDesc, DmCtg, DmPrdCtg, " + _
                "DmSalCtg, DmBagPcs, '' as DmCmCd, DmHld, DmHldDesc, DmParts, DmPartDesc, " + _
                "DmOldCd, DmCol, DmPrdSeq, DmPrdInst, DmDsgBy, DmDsgDt, DmModMkr, " + _
                "convert(varchar(5800), DmHistory) as DmHistory, " + _
                "convert(varchar(5800), DmFaults) as DmFaults, DmLstMdf, DmUom, DmVaCtg, " + _
                "DmKt, ModUsr, ModDt, ModTime, DmLsCtg, DmSetCd, '' as DmRegnExc, DmLockYN " + ws_NewFlds + _
                "from DsgMst " + ms_Cnd + " Order By DmCd, DmSz")
  If Rs_SrcX.RecCount <= 0 Then
    If mf_AllMst = False Then DispMsg "No Record Exists For This Scope", etError
    Exit Sub
  End If
  adc("wDisp") = "Design Master Transfer Out In Process ..."

  Set XmlDoc = New MSXML2.DOMDocument
  Set xmlRootElement = XmlDoc.createElement("EmrDsgMstTransfer")
  XmlDoc.appendChild xmlRootElement
  xmlRootElement.setAttribute "Ver", adc("wTgtVer")
  'xmlRootElement.setAttribute "Ver", ms_TrfOutVer

  Do While Not (Rs_SrcX.EOF Or Rs_SrcX.BOF)
    adc("wDisp") = "Design/ Size : " + Rs_SrcX!DmCd + "/ " + Rs_SrcX!DmSz
    Set xmlChildX = XmlDoc.createElement("DsgMst")
    Call WriteDocument(xmlChildX, Rs_SrcX)
    xmlRootElement.appendChild xmlChildX
    
    Set Rs_SrcA = moCn.OpenRes("Select * from DsgRm where " + _
                  "DrTcTyp= '" + Rs_SrcX!DmTcTyp + "' and DrCd= '" + Rs_SrcX!DmCd + "' and " + _
                  "DrSz= '" + Rs_SrcX!DmSz + "' " + ms_ChildModDtCnd + " Order By DrSr")
    Do While Not (Rs_SrcA.EOF Or Rs_SrcA.BOF)
      Set xmlChildA = XmlDoc.createElement("DsgRm")
      Call WriteDocument(xmlChildA, Rs_SrcA)
      xmlChildX.appendChild xmlChildA
      Rs_SrcA.MoveNext
    Loop
    
    Set Rs_SrcD = moCn.OpenRes("Select * from DsgLab where " + _
                  "DlTcTyp= '" + Rs_SrcX!DmTcTyp + "' and DlCd= '" + Rs_SrcX!DmCd + "' and " + _
                  "DlSz= '" + Rs_SrcX!DmSz + "' " + ms_ChildModDtCnd + " Order By DlSr")
    Do While Not (Rs_SrcD.EOF Or Rs_SrcD.BOF)
      Set xmlChildD = XmlDoc.createElement("DsgLab")
      Call WriteDocument(xmlChildD, Rs_SrcD)
      xmlChildX.appendChild xmlChildD
      Rs_SrcD.MoveNext
    Loop
    
    Set Rs_SrcG = moCn.OpenRes("Select * from DsgAna where " + _
                  "DaTcTyp= '" + Rs_SrcX!DmTcTyp + "' and DaCd= '" + Rs_SrcX!DmCd + "' and " + _
                  "DaSz= '" + Rs_SrcX!DmSz + "' " + ms_ChildModDtCnd + " Order By DaAnaSr")
    Do While Not (Rs_SrcG.EOF Or Rs_SrcG.BOF)
      Set xmlChildG = XmlDoc.createElement("DsgAna")
      Call WriteDocument(xmlChildG, Rs_SrcG)
      xmlChildX.appendChild xmlChildG
      Rs_SrcG.MoveNext
    Loop
    
    If (adc("wTgtVer") = "2.10.0" Or adc("wTgtVer") = "2.11.0") And gs_Ver = "2.11.0" Then
      Set Rs_SrcE = moCn.OpenRes("Select * from DsgPrt where " + _
                    "DpTcTyp= '" + Rs_SrcX!DmTcTyp + "' and DpCd= '" + Rs_SrcX!DmCd + "' and " + _
                    "DpSz= '" + Rs_SrcX!DmSz + "' " + ms_ChildModDtCnd + " Order By DpSr")
      Do While Not (Rs_SrcE.EOF Or Rs_SrcE.BOF)
        Set xmlChildE = XmlDoc.createElement("DsgAna")
        Call WriteDocument(xmlChildE, Rs_SrcE)
        xmlChildX.appendChild xmlChildE
        Rs_SrcE.MoveNext
      Loop
    End If
    
    ' Zubin 213
    If ValidateCurrVer(2, 13) And ValidateTgtVer(2, 13) Then
      Set Rs_SrcC = moCn.OpenRes("Select * from DsgModAvl where " + _
                    "DmaTcTyp= '" + Rs_SrcX!DmTcTyp + "' and DmaCd= '" + Rs_SrcX!DmCd + "' and " + _
                    "DmaSz= '" + Rs_SrcX!DmSz + "' " + ms_ChildModDtCnd + " Order By DmaSr")
      Do While Not (Rs_SrcC.EOF Or Rs_SrcC.BOF)
        Set xmlChildB = XmlDoc.createElement("DsgModAvl")
        Call WriteDocument(xmlChildB, Rs_SrcC)
        xmlChildX.appendChild xmlChildB
        Rs_SrcC.MoveNext
      Loop
    End If
    
    Rs_SrcX.MoveNext
  Loop
  
  Call SaveDocument(XmlDoc, "DsgMst" + ws_Identification)
  adc("wDisp") = "Design Transfer Out Completed"

  Set Rs_SrcX = Nothing
End Sub
Private Sub CreateXmlDocOrd()
  Call VerMatch
  
  Dim wOmTc As String, wOmYY As String, wOmChr As String, wOmNo As Single
  Dim wOdSr As Single, wOdOrdQty As Single
  
  ms_Cnd = adc.RepCond
  ms_Cnd = IIF(ms_Cnd <> "", " and " + ms_Cnd, "")
  
  Set Rs_SrcX = moCn.OpenRes("Select OmCoCd as xTCoCd, OmTc as xTTc, OmYy as xTYy, " + _
                "OmChr as xTChr, OmNo as xTNo, OmDt as xTDt, '' as xTFrRmLoc, " + _
                "'' xTToRmLoc, '' as xTFrRmDc, '' as xTFrDsgLoc, '' as xTToDsgLoc, " + _
                "'' as xTFrDsgDc, '' as xTCmCtg, OmCmCd as xTCmCd, OmPoNo as xTDocNo, " + _
                "OmPoDt as xTDocDt, '' as xTBOENo, '01/01/80' as xTBOEDt, '' as xTRefCmCtg, " + _
                "'' as xTRefCmCd, '' as xTRefDocNo, '01/01/80' as xTRefDocDt, " + _
                "'US$' as xTCurCd, (Select PNum From Param where PTyp= 'CURNCY' " + _
                "and PMCd= 'US$' and PSCd= '') as xTCnvRt, 0 as xTTotAmt, 0 as xTAmt01, " + _
                "0 as xTAmt02, 0 as xTAmt03, 0 as xTAmt04, 0 as xTAmt05, 0 as xTAmt06, " + _
                "0 as xTAmt07, 0 as xTAmt08, 0 as xTNetAmt, '' as xTAna01, '' as xTAna02, " + _
                "'' as xTAna03, '' as xTAna04, '' as xTAna05, '' as xTAna06, '' as xTAna07, " + _
                "'' as xTAna08, '' as xTDesc, 0 as xTKey, '01/01/80' as xTUseBefDt, " + _
                "OmLmgCst as xTLmgCst, OmLmgDt as xTLmgDt, OmLmgSal as xTLmgSal, " + _
                "OmLmpCst as xTLmpCst, OmLmpDt as xTLmpDt, OmLmpSal as xTLmpSal, " + _
                "OmExpDelDt as xTExpDelDt, OmDelDt as xTPrdDelDt, " + _
                "OmCmDelTerms as xTCmDelTerms, OmCmPayTerms as xTCmPayTerms, " + _
                "OmPerson as xTPerson, OmPrtCd as xTPrtCd, OmMulby as xTMulBy, " + _
                "OmFixPrc as xTFixPrc, OmGldAs as xTGldAs, OmIWtEqOrd as xTiWtEqOrd, " + _
                "OmIWtFrOrd as xTiWtFrOrd, OmIGldRtEqOrd as xTiGldRtEqOrd, " + _
                "'' as xTOwnBy, '' as xTAppTo, ModUsr, ModDt, ModTime " + _
                "from OrdMst where OmCoCd= '" + gs_CoCd + "' and OmTc in (Select PMCd from " + _
                "Param where PTyp= 'TC' and PValue= '" + adc("wOmTcTyp") + "') " + ms_Cnd + _
                "Order By OmTc, OmYy, OmChr, OmNo")
  
  If Rs_SrcX.RecCount <= 0 Then DispMsg "No Record Exists For This Scope", etError: Exit Sub
  adc("wDisp") = "Order Transfer Out In Process ..."
  
  Set XmlDoc = New MSXML2.DOMDocument
  Set xmlRootElement = XmlDoc.createElement("EmrOrdMstTransfer")
  XmlDoc.appendChild xmlRootElement
  xmlRootElement.setAttribute "Ver", adc("wTgtVer")
  'xmlRootElement.setAttribute "Ver", ms_TrfOutVer
   
  Do While Not (Rs_SrcX.EOF Or Rs_SrcX.BOF)
      wOmTc = Rs_SrcX!xTTc
      wOmYY = Rs_SrcX!xTYy
      wOmChr = Rs_SrcX!xTChr
      wOmNo = Rs_SrcX!xTNo
      
      adc("wDisp") = "Order : " + wOmTc + "/" + wOmYY + "/" + wOmChr + "/" + Trim(CStr(wOmNo))
      Set xmlChildX = XmlDoc.createElement("xTxn")
      Call WriteDocument(xmlChildX, Rs_SrcX)
      xmlRootElement.appendChild xmlChildX
      
      Set Rs_SrcD = moCn.OpenRes("Select OtCoCd as xTtCoCd, OtTc as xTtTc, OtYy as xTtYy, " + _
                    "OtChr as xTtChr, OtNo as xTtNo, OtSr as xTtSr, OtRmCtg as xTtRmCtg, " + _
                    "OtRmSCtg as xTtRmSCtg, OtRmCd as xTtRmCd, ModUsr, ModDt, ModTime " + _
                    "from OrdTgt where OtCoCd= '" + gs_CoCd + "' and " + _
                    "OtTc= '" + wOmTc + "' and OtYy ='" + wOmYY + "' and " + _
                    "OtChr= '" + wOmChr + "' and OtNo= " + CStr(wOmNo) + " and " + _
                    "OtSr= 0 Order By OtRmCtg, OtRmSCtg")
      Do While Not (Rs_SrcD.EOF Or Rs_SrcD.BOF)
        Set xmlChildD = XmlDoc.createElement("xTxnTgt")
        Call WriteDocument(xmlChildD, Rs_SrcD)
        xmlChildX.appendChild xmlChildD
        Rs_SrcD.MoveNext
      Loop
      
      Set Rs_SrcA = moCn.OpenRes("Select OdCoCd as xTdCoCd, OdTc as xTdTc, OdYy as xTdYy, " + _
                    "OdChr as xTdChr, OdNo as xTdNo, OdSr as xTdSr, 0 as xTdSrNo, " + _
                    "'' as xTdIdOwnBy, '' as xTdIdYy, 0 as xTdIdNo, '' as xTdFrDsgLoc, " + _
                    "'' as xTdToDsgLoc, '' as xTdFrDsgDc, '' as xTdToDsgDc, 0 as xTdKey, " + _
                    "'' as xTdYyKey, '' as xTdMemRefYy, 0 as xTdMemRefKey, '' as xTdTxnRefYy, " + _
                    "0 as xTdTxnRefKey, '' as xTdDmTyp, OdDmCd as xTdDmCd, OdSfx as xTdSfx, " + _
                    "OdDmSz as xTdDmSz, OdExpDelDt as xTdExpDelDt, OdDelDt as xTdPrdDelDt, " + _
                    "OdOrdEnt as xTdQty, OdCalcPrc as xTdSalValPpc, OdSalPrc as xTdOffValPpc, " + _
                    "OdCstPrc as xTdCstValPpc, OdOrdQty* OdSalPrc as xTdOffVal, " + _
                    "OdOrdQty* OdCalcPrc as xTdSalVal, OdOrdQty* OdCstPrc as xTdCstVal, " + _
                    "OdDmPrdInst as xTdDmPrdInst, OdCmPrdInst as xTdCmPrdInst, OdCmStmpInst as xTdCmStmpInst, " + _
                    "OdSzInst as xTdSzInst, OdPrtCd as xTdPrtCd, OdHld as xTdHld, " + _
                    "OdHldDesc as xTdHldDesc, OdKt as xTdKt, OdMulby as xTdMulBy, OdFixPrc as xTdFixPrc, " + _
                    "OdGldAs as xTdGldAs, OdIWtEqOrd as xTdIWtEqOrd, OdIWtFrOrd as xTdIWtFrOrd, " + _
                    "OdIGldRtEqOrd as xTdIGldRtEqOrd, OdDmCol as xTdDmCol, OdGldAsWt as xTdGldAsWt, " + _
                    "OdSfxDesc as xTdSfxDesc, 0 as xTdAmt01, 0 as xTdAmt02, 0 as xTdAmt03, 0 as xTdAmt04, " + _
                    "0 as xTdTotAmt, '' as xTdAna01, '' as xTdAna02, '' as xTdAna03, '' as xTdAna04, " + _
                    "'' as xTdOwnBy, '' as xTdAppTo, IsNull((Select sum(OrWt/(case " + _
                    "when OrRmCtg in ('C', 'D') then 5 else 1 end)) From OrdRm where " + _
                    "OrCoCd= OdCoCd and OrTc= OdTc and OrYy= OdYy and OrChr= OdChr " + _
                    "and OrNo= OdNo and OrSr= OdSr), 0)* OdOrdQty as xTdGrsWt, ModUsr, ModDt, ModTime " + _
                    "from OrdDsg where OdCoCd= '" + gs_CoCd + "' and " + _
                    "OdTc= '" + wOmTc + "' and OdYy= '" + wOmYY + "' and OdChr= '" + wOmChr + "' and " + _
                    "OdNo= " + CStr(wOmNo) + " Order By OdSr")
      
      '*** Order Design Level
      Do While Not (Rs_SrcA.EOF Or Rs_SrcA.BOF)
          wOdSr = Rs_SrcA("xTdSr")
          wOdOrdQty = moCn.GetFldVal("Select OdOrdQty from OrdDsg where OdCoCd= '" + gs_CoCd + "' and " + _
                      "OdTc= '" + wOmTc + "' and OdYy= '" + wOmYY + "' and OdChr= '" + wOmChr + "' and " + _
                      "OdNo= " + CStr(wOmNo) + " and OdSr= " + CStr(wOdSr) + " ")
          adc("wDisp") = "Design :" + wOmTc + "/" + wOmYY + "/" + wOmChr + "/" + Trim(CStr(wOmNo)) + "/" + Trim(CStr(wOdSr))
          Set xmlChildA = XmlDoc.createElement("xTxnDsg")
          Call WriteDocument(xmlChildA, Rs_SrcA)
          xmlChildX.appendChild xmlChildA
          
          Set Rs_SrcB = moCn.OpenRes("Select OrCoCd as xTrCoCd, OrTc as xTrTc, " + _
                        "OrYy as xTrYy, OrChr as xTrChr, OrNo as xTrNo, OrSr as xTrSr, " + _
                        "OrSrNo as xTrSrNo, OrRmCtg as xTrRmCtg, OrRmSCtg as xTrRmSCtg, " + _
                        "OrRmCd as xTrRmCd, OrLn1 as xTrLn1, OrLn2 as xTrLn2, OrRmPtr as xTrRmPtr, " + _
                        "OrQty as xTrQtyPPc, OrWt as xTrWtPPc, OrSalRt as xTrSalRt, " + _
                        "OrLmeSal as xTrLmeSal, OrSalVal as xTrSalValPPc, OrCstRt as xTrCstRt, " + _
                        "OrCstVal as xTrCstValPPc, OrSetSCd as xTrSetSCd, OrSetSalRt as xTrSetSalRt, " + _
                        "OrSetSalVal as xTrSetSalValPPc, OrSetCstRt as xTrSetCstRt, " + _
                        "OrSetCstVal as xTrSetCstValPPc, OrWsQty as xTrWsQtyPpc, OrHsQty as xTrHsQtyPpc, " + _
                        "OrSubShp as xTrSubShp, OrAlyCd as xTrAlyCd, OrAlySalRt as xTrAlySalRt, " + _
                        "OrAlyCstRt as xTrAlyCstRt, OrMainMet as xTrMainMet, '' as xTrFrRmLoc, " + _
                        "'' as xTrFrRmDc, 0 as xTrRmSz, 0 as xTrRmStkRt, OrQty* " + CStr(wOdOrdQty) + " as xTrRmQty, " + _
                        "OrWt* " + CStr(wOdOrdQty) + " as xTrRmWt, '' as xTrToRmLoc, '' as xTrToRmDc, " + _
                        "'' as xTrDesc, '' as xTrOwnBy, '' as xTrAppTo, ModUsr, " + _
                        "ModDt, ModTime from OrdRm where OrCoCd= '" + gs_CoCd + "' and " + _
                        "OrTc= '" + wOmTc + "' and OrYy= '" + wOmYY + "' and OrChr= '" + wOmChr + "' and " + _
                        "OrNo= " + CStr(wOmNo) + " and OrSr= " + CStr(wOdSr) + " Order By OrSrNo")
          Do While Not (Rs_SrcB.EOF Or Rs_SrcB.BOF)
            Set xmlChildB = XmlDoc.createElement("xTxnRm")
            Call WriteDocument(xmlChildB, Rs_SrcB)
            xmlChildA.appendChild xmlChildB
            Rs_SrcB.MoveNext
          Loop
          
          Set Rs_SrcC = moCn.OpenRes("Select OlCoCd as xTLCoCd, OlTc as xTLTc, " + _
                        "OlYy as xTLYy, OlChr as xTLChr, OlNo as xTLNo, " + _
                        "OlSr as xTLSr, OlSrNo as xTLSrNo, OlMCd as xTLMCd, OlSCd as xTLSCd, " + _
                        "OlQty as xTLQtyPPc, OlSalRt as xTLSalRt, OlSalVal as xTLSalValPPc, " + _
                        "OlCstRt as xTLCstRt, OlCstVal as xTLCstValPPc, OlQw as xTLSalQw, " + _
                        "OlCstQw as xTLCstQw, '' as xTLOwnBy, '' as xTLAppTo, ModUsr, " + _
                        "ModDt, ModTime  from OrdLab where OlCoCd= '" + gs_CoCd + "' and " + _
                        "OlTc= '" + wOmTc + "' and OlYy= '" + wOmYY + "' and " + _
                        "OlChr= '" + wOmChr + "' and OlNo= " + CStr(wOmNo) + " and " + _
                        "OlSr= " + CStr(wOdSr) + " Order By OlSrNo")
          Do While Not (Rs_SrcC.EOF Or Rs_SrcC.BOF)
            Set xmlChildC = XmlDoc.createElement("xTxnLab")
            Call WriteDocument(xmlChildC, Rs_SrcC)
            xmlChildA.appendChild xmlChildC
            Rs_SrcC.MoveNext
          Loop
          
          Set Rs_SrcB1 = moCn.OpenRes("Select OaCoCd as xTdaCoCd, OaTc as xTdaTc, " + _
                         "OaYy as xTdaYy, OaChr as xTdaChr, OaNo as xTdaNo, OaSr as xTdaSr, " + _
                         "OaSrNo as xTdaSrNo, OaOrdQty as xTdaQty, OaDt as xTdaDt, " + _
                         "OaDesc as xTdaDesc, '' As xTdaOwnBy, '' As xTdaAppTo, " + _
                         "ModUsr, ModDt, ModTime from OrdAmd where OaCoCd= '" + gs_CoCd + "' and " + _
                         "OaTc= '" + wOmTc + "' and OaYy ='" + wOmYY + "' and " + _
                         "OaChr= '" + wOmChr + "' and OaNo= " + CStr(wOmNo) + " and " + _
                         "OaSr= " + CStr(wOdSr) + " Order By OaSrNo")
          Do While Not (Rs_SrcB1.EOF Or Rs_SrcB1.BOF)
            Set xmlChildB1 = XmlDoc.createElement("xTxnDsgAmd")
            Call WriteDocument(xmlChildB1, Rs_SrcB1)
            xmlChildA.appendChild xmlChildB1
            Rs_SrcB1.MoveNext
          Loop
          
          Set Rs_SrcC1 = moCn.OpenRes("Select OtCoCd as xTtCoCd, OtTc as xTtTc, OtYy as xTtYy, " + _
                         "OtChr as xTtChr, OtNo as xTtNo, OtSr as xTtSr, OtRmCtg as xTtRmCtg, " + _
                         "OtRmSCtg as xTtRmSCtg, OtRmCd as xTtRmCd, ModUsr, ModDt, ModTime " + _
                         "from OrdTgt where OtCoCd= '" + gs_CoCd + "' and OtTc= '" + wOmTc + "' and " + _
                         "OtYy ='" + wOmYY + "' and OtChr= '" + wOmChr + "' and " + _
                         "OtNo= " + CStr(wOmNo) + " and OtSr= " + CStr(wOdSr) + " Order By OtRmCtg, OtRmSCtg")
          Do While Not (Rs_SrcC1.EOF Or Rs_SrcC1.BOF)
            'Set xmlChildB1 = XmlDoc.createElement("xTxnTgt")
            Set xmlChildC1 = XmlDoc.createElement("xTxnTgt")        '*** changed on 15/01/05
            Call WriteDocument(xmlChildC1, Rs_SrcC1)
            xmlChildA.appendChild xmlChildC1
            Rs_SrcC1.MoveNext
          Loop
          
          Rs_SrcA.MoveNext
      Loop
      Rs_SrcX.MoveNext
  Loop
  
  Call SaveDocument(XmlDoc, "OrdMst")
  adc("wDisp") = "Order Transfer Out Completed"
  
  Set Rs_SrcX = Nothing
  Set Rs_SrcA = Nothing
  Set Rs_SrcB = Nothing
  Set Rs_SrcC = Nothing
  Set Rs_SrcB1 = Nothing
  Set Rs_SrcC1 = Nothing
  Set Rs_SrcD = Nothing
End Sub
Private Sub CreateXmlDocInv()
  Call VerMatch
  
  Dim wInTc As String, wInYy As String, wInChr As String, wInNo As Single, wCmCd As String
  Dim wIdSr As Single, wIdQty As Single
  
  Dim ws_str As String
  
  ms_Cnd = adc.RepCond
  ms_Cnd = IIF(ms_Cnd <> "", " and " + ms_Cnd, "")
  
  Set Rs_SrcX = moCn.OpenRes("set dateformat dmy Select InCoCd as xTCoCd, InTc as xTTc, InYy as xTYy, " + _
                "InChr as xTChr, InNo as xTNo, InDt as xTDt, '' as xTFrRmLoc, " + _
                "'' xTToRmLoc, '' as xTFrRmDc, '' as xTFrDsgLoc, '' as xTToDsgLoc, " + _
                "'' as xTFrDsgDc, '' as xTCmCtg, InCmCd as xTCmCd, InExpNo as xTDocNo, " + _
                "'01/01/80' as xTDocDt, InPoNo as xTBOENo, InPoDt as xTBOEDt, '' as xTRefCmCtg, " + _
                "'' as xTRefCmCd, '' as xTRefDocNo, '01/01/80' as xTRefDocDt, " + _
                "'US$' as xTCurCd, (Select PNum From Param where PTyp= 'CURNCY' " + _
                "and PMCd= 'US$' and PSCd= '') as xTCnvRt, InFOB as xTTotAmt, 0 as xTAmt01, " + _
                "0 as xTAmt02, 0 as xTAmt03, 0 as xTAmt04, 0 as xTAmt05, 0 as xTAmt06, " + _
                "0 as xTAmt07, 0 as xTAmt08, InFOB as xTNetAmt, '' as xTAna01, '' as xTAna02, " + _
                "'' as xTAna03, '' as xTAna04, '' as xTAna05, '' as xTAna06, '' as xTAna07, " + _
                "'' as xTAna08, '' as xTDesc, 0 as xTKey, '01/01/80' as xTUseBefDt, " + _
                "InLmgCst as xTLmgCst, InLmgDt as xTLmgDt, InLmgSal as xTLmgSal, " + _
                "InLmpCst as xTLmpCst, InLmpDt as xTLmpDt, InLmpSal as xTLmpSal, " + _
                "'01/01/80' as xTExpDelDt, '01/01/80' as xTPrdDelDt, CMDelTerms as xTCmDelTerms, " + _
                "InCmPayTerms as xTCmPayTerms, CmCpName as xTPerson, '' as xTPrtCd, 1 as xTMulBy, " + _
                "CmFixPrc as xTFixPrc, CmGldAs as xTGldAs, '' as xTiWtEqOrd, '' as xTiWtFrOrd, " + _
                "'' as xTiGldRtEqOrd, '' as xTOwnBy, '' as xTAppTo, InvHd.ModUsr, " + _
                "InvHd.ModDt, InvHd.ModTime " + _
                "from InvHd join CustMst On CmCtg= 'C' and CmCd= InCmCd " + _
                "where InCoCd= '" + gs_CoCd + "' and InTc in (Select PMCd from " + _
                "Param where PTyp= 'TC' and PValue= 'IN') " + ms_Cnd + _
                "Order By InTc, InYy, InChr, InNo")
  
  If Rs_SrcX.RecCount <= 0 Then DispMsg "No Record Exists For This Scope", etError: Exit Sub
  adc("wDisp") = "Invoice Transfer Out In Process ..."
  
  Set XmlDoc = New MSXML2.DOMDocument
  Set xmlRootElement = XmlDoc.createElement("EmrInvHdTransfer")
  XmlDoc.appendChild xmlRootElement
  xmlRootElement.setAttribute "Ver", adc("wTgtVer")
  'xmlRootElement.setAttribute "Ver", ms_TrfOutVer
   
  Do While Not (Rs_SrcX.EOF Or Rs_SrcX.BOF)
      wInTc = Rs_SrcX!xTTc
      wInYy = Rs_SrcX!xTYy
      wInChr = Rs_SrcX!xTChr
      wInNo = Rs_SrcX!xTNo
      wCmCd = Rs_SrcX!xTCmCd
      
      adc("wDisp") = "Invoice : " + wInTc + "/" + wInYy + "/" + wInChr + "/" + Trim(CStr(wInNo))
      Set xmlChildX = XmlDoc.createElement("xTxn")
      Call WriteDocument(xmlChildX, Rs_SrcX)
      xmlRootElement.appendChild xmlChildX
      
      ws_str = "Select IdCoCd as xTdCoCd, IdTc as xTdTc, IdYy as xTdYy, " + _
               "IdChr as xTdChr, IdNo as xTdNo, IdSr as xTdSr, 0 as xTdSrNo, " + _
               "'' as xTdIdOwnBy, IdYy as xTdIdYy, 0 as xTdIdNo, '' as xTdFrDsgLoc, " + _
               "'' as xTdToDsgLoc, '' as xTdFrDsgDc, '' as xTdToDsgDc, 0 as xTdKey, " + _
               "'' as xTdYyKey, '' as xTdMemRefYy, 0 as xTdMemRefKey, '' as xTdTxnRefYy, " + _
               "0 as xTdTxnRefKey, '' as xTdDmTyp, IdDmCd as xTdDmCd, IdSfx as xTdSfx, " + _
               "IdDmSz as xTdDmSz, '01/01/80' as xTdExpDelDt, '01/01/80' as xTdPrdDelDt, " + _
               "IdQty as xTdQty, (case when IdQty> 0 then IdIVal/IdQty else 0 end) as xTdSalValPpc, " + _
               "(case when IdQty> 0 then IdIVal/IdQty else 0 end) as xTdOffValPpc, " + _
               "0 as xTdCstValPpc, IdIVal as xTdOffVal, IdIVal as xTdSalVal, 0 as xTdCstVal, " + _
               "DmPrdInst as xTdDmPrdInst, CmPrdInst as xTdCmPrdInst, CmStmpInst as xTdCmStmpInst, " + _
               "OdSzInst as xTdSzInst, '' as xTdPrtCd, 'N' as xTdHld, '' as xTdHldDesc, "
      ws_str = ws_str + "IdKt as xTdKt, CmMulby as xTdMulBy, IdFixPrc as xTdFixPrc, " + _
               "IdGldAs as xTdGldAs, '' as xTdIWtEqOrd, '' as xTdIWtFrOrd, " + _
               "'' as xTdIGldRtEqOrd, OdDmCol as xTdDmCol, IdiGldAsWt as xTdGldAsWt, " + _
               "OdSfxDesc as xTdSfxDesc, 0 as xTdAmt01, 0 as xTdAmt02, 0 as xTdAmt03, 0 as xTdAmt04, " + _
               "IdIVal as xTdTotAmt, '' as xTdAna01, '' as xTdAna02, '' as xTdAna03, '' as xTdAna04, " + _
               "'' as xTdOwnBy, '' as xTdAppTo, IsNull((Select sum(IrRmIWt/" + _
               "(case when RmCtg in ('C', 'D') then 5 else 1 end)) From InvRm, " + _
               "RmMst where RMCd= IrRmCd and IrCoCd= IdCoCd and IrTc= IdTc " + _
               "and IrYy= IdYy and IrChr= IdChr and IrNo= IdNo and IrSr= IdSr), 0) as xTdGrsWt, " + _
               "InvDsg.ModUsr, InvDsg.ModDt, InvDsg.ModTime " + _
               "from InvDsg join DsgMst on DmTcTyp= 'DM' and DmCd= IdDmCd and DmSz= '' " + _
               "join OrdDsg on OdCoCd= IdCoCd and OdTc= IdExpOdTc and OdYy= IdExpOdYy " + _
               "and OdChr= IdExpOdChr and OdNo= IdExpOdNo and OdSr= IdExpOdSr " + _
               "join CustMst on CmCtg= 'C' and CmCd= '" + wCmCd + "' " + _
               "where IdCoCd= '" + gs_CoCd + "' and IdTc= '" + wInTc + "' and IdYy= '" + wInYy + "' " + _
               "and IdChr= '" + wInChr + "' and IdNo= " + CStr(wInNo) + " Order By IdSr"
      Set Rs_SrcA = moCn.OpenRes(ws_str)
      
      '*** Invoice Design Level
      Do While Not (Rs_SrcA.EOF Or Rs_SrcA.BOF)
          wIdSr = Rs_SrcA("xTdSr")
          wIdQty = Rs_SrcA("xTdQty")
          adc("wDisp") = "Design :" + wInTc + "/" + wInYy + "/" + wInChr + "/" + Trim(CStr(wInNo)) + "/" + Trim(CStr(wIdSr))
          Set xmlChildA = XmlDoc.createElement("xTxnDsg")
          Call WriteDocument(xmlChildA, Rs_SrcA)
          xmlChildX.appendChild xmlChildA
          
          ws_str = "Select IrCoCd as xTrCoCd, IrTc as xTrTc, IrYy as xTrYy, IrChr as xTrChr, " + _
                   "IrNo as xTrNo, IrSr as xTrSr, IrSrNo as xTrSrNo, RmCtg as xTrRmCtg, " + _
                   "RmSCtg as xTrRmSCtg, IrRmCd as xTrRmCd, IrRmSz as xTrLn1, 0 as xTrLn2, " + _
                   "(case when RmCtg in ('C', 'D') and IrRmQty> 0 then IrRmIWt/ IrRmQty " + _
                   "else 0 end) as xTrRmPtr, IrRmQty" + IIF(wIdQty > 0, "/" + CStr(wIdQty), " ") + _
                   " as xTrQtyPPc, IrRmIWt" + IIF(wIdQty > 0, "/" + CStr(wIdQty), " ") + _
                   " as xTrWtPPc, IrRmIRt as xTrSalRt, " + _
                   "(case when RmCtg= 'G' then " + CStr(Rs_SrcX!xTLmgSal) + " when " + _
                   "RmCtg= 'P' then " + CStr(Rs_SrcX!xTLmpSal) + " else 0 end)as xTrLmeSal, " + _
                   "IrRmIVal" + IIF(wIdQty > 0, "/" + CStr(wIdQty), " ") + " as xTrSalValPPc, " + _
                   "0 as xTrCstRt, 0 as xTrCstValPPc, IrSetSCd as xTrSetSCd, IrSetIRt as xTrSetSalRt, " + _
                   "IrSetIVal" + IIF(wIdQty > 0, "/" + CStr(wIdQty), " ") + " as xTrSetSalValPPc, " + _
                   "0 as xTrSetCstRt, 0 as xTrSetCstValPPc, "
          ws_str = ws_str + "(case when RmCtg in ('C', 'D') and OdWh= 'W' then " + _
                   "IrRmQty" + IIF(wIdQty > 0, "/" + CStr(wIdQty), " ") + " else 0 end) as xTrWsQtyPpc, " + _
                   "(case when RmCtg in ('C', 'D') and OdWh in ('H', 'WH') then " + _
                   "IrRmQty" + IIF(wIdQty > 0, "/" + CStr(wIdQty), " ") + " else 0 end) as xTrHsQtyPpc, " + _
                   "'' as xTrSubShp, '' as xTrAlyCd, 0 as xTrAlySalRt, " + _
                   "0 as xTrAlyCstRt, IrMainMet as xTrMainMet, '' as xTrFrRmLoc, " + _
                   "'' as xTrFrRmDc, IrRmSz as xTrRmSz, IrStkRt as xTrRmStkRt, " + _
                   "IrRmQty as xTrRmQty, IrRmIWt as xTrRmWt, '' as xTrToRmLoc, '' as xTrToRmDc, " + _
                   "'' as xTrDesc, '' as xTrOwnBy, '' as xTrAppTo, InvRm.ModUsr, " + _
                   "InvRm.ModDt, InvRm.ModTime from InvRm join InvDsg on IdCoCd= IrCoCd and " + _
                   "IdTc= IrTc and IdYy= IrYy and IdChr= IrChr and IdNo= IrNo and IdSr= IrSr " + _
                   "join OrdDsg on OdCoCd= IdCoCd and OdTc= IdExpOdTc and OdYy= IdExpOdYy " + _
                   "and OdChr= IdExpOdChr and OdNo= IdExpOdNo and OdSr= IdExpOdSr " + _
                   "join RmMst on RmCd= IrRmCd where IrCoCd= '" + gs_CoCd + "' and IrTc= '" + wInTc + "' and " + _
                   "IrYy= '" + wInYy + "' and IrChr= '" + wInChr + "' and " + _
                   "IrNo= " + CStr(wInNo) + " and IrSr= " + CStr(wIdSr) + " Order By IrSrNo"

          Set Rs_SrcB = moCn.OpenRes(ws_str)
          
          Do While Not (Rs_SrcB.EOF Or Rs_SrcB.BOF)
            Set xmlChildB = XmlDoc.createElement("xTxnRm")
            Call WriteDocument(xmlChildB, Rs_SrcB)
            xmlChildA.appendChild xmlChildB
            Rs_SrcB.MoveNext
          Loop
          
          Set Rs_SrcC = moCn.OpenRes("Select iLCoCd as xTLCoCd, iLTc as xTLTc, iLYy as xTLYy, " + _
                        "iLChr as xTLChr, iLNo as xTLNo, iLSr as xTLSr, iLSrNo as xTLSrNo, " + _
                        "iLMCd as xTLMCd, iLSCd as xTLSCd, iLQty as xTLQtyPPc, iLiRt as xTLSalRt, " + _
                        "iLiVal" + IIF(wIdQty > 0, "/" + CStr(wIdQty), " ") + " as xTLSalValPPc, " + _
                        "0 as xTLCstRt, 0 as xTLCstValPPc, IlQw as xTLSalQw, " + _
                        "IlQw as xTLCstQw, '' as xTLOwnBy, '' as xTLAppTo, " + _
                        "InvLab.ModUsr, InvLab.ModDt, InvLab.ModTime " + _
                        "from InvDsg join InvLab on iLCoCd= IdCoCd and iLTc= IdTc and " + _
                        "iLYy= IdYy and iLChr= IdChr and iLNo= IdNo and iLSr= IdSr " + _
                        "where IdCoCd= '" + gs_CoCd + "' and IdTc= '" + wInTc + "' and " + _
                        "IdYy= '" + wInYy + "' and IdChr= '" + wInChr + "' and " + _
                        "IdNo= " + CStr(wInNo) + " and IdSr= " + CStr(wIdSr) + " " + _
                        "Order By iLSrNo")
          Do While Not (Rs_SrcC.EOF Or Rs_SrcC.BOF)
            Set xmlChildC = XmlDoc.createElement("xTxnLab")
            Call WriteDocument(xmlChildC, Rs_SrcC)
            xmlChildA.appendChild xmlChildC
            Rs_SrcC.MoveNext
          Loop
          
          Rs_SrcA.MoveNext
      Loop
      Rs_SrcX.MoveNext
  Loop
  
  Call SaveDocument(XmlDoc, "InvHd")
  adc("wDisp") = "Invoice Transfer Out Completed"
  
  Set Rs_SrcX = Nothing
  Set Rs_SrcA = Nothing
  Set Rs_SrcB = Nothing
  Set Rs_SrcC = Nothing
End Sub
Private Sub WriteDocument(mo_DocElement As MSXML2.IXMLDOMElement, mRs_RecSet As MwfLib.MDORowSet)
  Dim i As Integer
  For i = 0 To mRs_RecSet.FldCount - 1
    mo_DocElement.setAttribute mRs_RecSet.FldDets(i).fdFldName, mRs_RecSet.FldValue(i)
  Next i
End Sub
Private Sub SaveDocument(ByRef mo_XmlDoc As MSXML2.DOMDocument, ms_XTblName As String)
'*** Changed on 02-06-04
On Error GoTo Err_Hndlr
  '### Dim wComDialog As CommonDialog
  '### Set wComDialog = CreateObject(CommonDialog)
  '### wComDialog.ShowOpen
  
  If mf_AllMst = True Then
    If Right(TxtFolderPath.text, 1) <> "\" Then TxtFolderPath.text = TxtFolderPath.text + "\"
    mo_XmlDoc.Save TxtFolderPath.text + ms_XTblName + CStr(Year(moCn.SrvrDate)) + Format(Month(moCn.SrvrDate), "00") + Format(Day(moCn.SrvrDate), "00") + ".xml"
    'mo_XmlDoc.Save "C:\Temp\AllMst\" + ms_XTblName + CStr(Year(moCn.SrvrDate)) + Format(Month(moCn.SrvrDate), "00") + Format(Day(moCn.SrvrDate), "00") + ".xml"
    Exit Sub
  End If
  
  ComDialogOpn.DialogTitle = "Save " + ms_XTblName + " XML File As"
  ComDialogOpn.InitDir = moCn.GetFldVal("Select PDesc225 from Param where PTyp= 'TRFDIR' and PMCd= 'TRFDIR' ")
  ComDialogOpn.Filter = "XML Files (*.xml)"
  ComDialogOpn.FilterIndex = 1
  ComDialogOpn.CancelError = True
  
  '*** cdlOFNOverwritePrompt specifies whether the user should be alerted if he is overwriting and existing file
  '*** cdlOFNPathMustExist makes sure that the the user should specify an existing path
  '*** cdlOFNLongNames specifies that long names can be used for the files
  '*** cdlOFNExtensionDifferent indicates that the user has selected a different extension than the one specified through the DefaultExt property
  ComDialogOpn.Flags = cdlOFNOverwritePrompt Or cdlOFNLongNames Or cdlOFNPathMustExist Or cdlOFNExtensionDifferent
  ComDialogOpn.DefaultExt = "xml"
  ComDialogOpn.FileName = ms_XTblName + CStr(Year(moCn.SrvrDate)) + Format(Month(moCn.SrvrDate), "00") + Format(Day(moCn.SrvrDate), "00")
  ComDialogOpn.ShowSave
  ''' If cdlOFNExtensionDifferent Then DispMsg "Cannot Change Extension", etError: Exit Sub
    
  If Len(ComDialogOpn.FileName) > 4 Then
    If UCase(Right(ComDialogOpn.FileName, 4)) = ".XML" Then
      mo_XmlDoc.Save ComDialogOpn.FileName
    Else
      DispMsg "File Extension Is Wrong", etError: Exit Sub
    End If
  End If
  
  If Trim(ComDialogOpn.FileName) <> "" And Dir(ComDialogOpn.FileName) <> "" Then
    Dim wl_XMLSizeInKB As Single, wo_FileSysObj, wo_MyFile
    Set wo_FileSysObj = CreateObject("Scripting.FileSystemObject")
    Set wo_MyFile = wo_FileSysObj.GetFile(ComDialogOpn.FileName)
    wl_XMLSizeInKB = MWLib.Div(wo_MyFile.Size, 1024)
    '*** For Rounding Up By One
    wl_XMLSizeInKB = IIF(wl_XMLSizeInKB > Round(wl_XMLSizeInKB), 1, 0) + Round(wl_XMLSizeInKB)
    
    '### (Temporarily Commented)
    'If wl_XMLSizeInKB > 1200 Then
    '  wo_MyFile.Delete
    '  DispMsg "The File Size Was Exceeding The Maximum Permissible Limits Hence Was Not Saved", etError
    'End If
    '### (Temporarily Commented)
  End If
  
  
  '********************************* (Correct Code) ************************************
  '  If Trim(ComDialogOpn.FileName) <> "" And Dir(ComDialogOpn.FileName) <> "" Then
  '    Dim wXMLFile As Integer, wXMLSizeInKB As Long, wFileSysObj, wMyFile
  '    wXMLFile = FreeFile(0)
  '    Open ComDialogOpn.FileName For Binary Access Read As wXMLFile
  '    wXMLSizeInKB = MWLib.Div(LOF(wXMLFile), 1024)
  '    Close wXMLFile
  '    wXMLSizeInKB = IIf(wXMLSizeInKB <> Round(wXMLSizeInKB), Int(wXMLSizeInKB + 1), wXMLSizeInKB)
  '    If wXMLSizeInKB > 1200 Then
  '      Set wFileSysObj = CreateObject("Scripting.FileSystemObject")
  '      Set wMyFile = wFileSysObj.GetFile(ComDialogOpn.FileName)
  '      wMyFile.Delete
  '    End If
  '  End If
  '********************************* (Correct Code) ************************************


  '  Shell "D:\Installations\WinZip\winzip32.exe c:\temp\aaa.txt c:\temp\aaa.zip", vbHide


  '*** For the following code to work, the refrences:
  '   1.  Microsoft Script Control 1.0
  '   2.  Microsoft Scripting Runtime
  '   have to be added to the project
  '************************************************************************************
  '  Dim wfilesys As FileSystemObject
  '  Dim newfile As File
  '  Set wfilesys = CreateObject("Scripting.FileSystemObject")
  '  Set newfile = wfilesys.GetFile("C:\temp\aaa.txt")
  '  newfile.Size
  '  Set newfile.Attributes = VbFileAttribute.vbHidden
  '************************************************************************************


  '***
  '  Dim fso, MyFile
  '  Set fso = CreateObject("Scripting.FileSystemObject")
  '  Set MyFile = fso.CreateTextFile(ComDialogOpn.FileName, True)
  '  MyFile.WriteLine ("This is a test.")
  '  MyFile.Close
  '  Set MyFile = fso.GetFile("c:\testfile.txt")
  '  MyFile.Delete
  '***

'*** Changed on 02-06-04
Err_Hndlr:
  If Err.Number = 32755 Then
    adc("wDisp") = "Operation Cancelled "
    DispMsg "Operation Cancelled ", etInfo
    Err.Clear
  End If
End Sub
Private Sub DispFra(ByVal pv_DispFra As en_EmrTrfOutDispFra)
  '*** This is the common procedure which is used to toggle between hiding and displaying
      'the frames (i.e. if the frame is visible then make it invisible and viseversa)
      'associated with a command button

  Select Case pv_DispFra
  Case Is = Rm
    If FraRm.Visible = True Then
      FraRm.Visible = False
      '*** required to pass the array position Rm-0, to resotre the 'Causes Validation' State
      Call EnaDisaCmds(False, CmdRm, Rm)
      CmdRm.SetFocus
    Else
      FraRm.Visible = True
      adc("wRmZ").SetFocus
      FraRm.ZOrder
      Call EnaDisaCmds(True, CmdRm)
    End If
  Case Is = Cust
    If FraCust.Visible = True Then
      FraCust.Visible = False
      Call EnaDisaCmds(False, CmdCust, Cust)
      CmdCust.SetFocus
    Else
      FraCust.Visible = True
      adc("wCmCtg").SetFocus
      FraCust.ZOrder
      Call EnaDisaCmds(True, CmdCust)
    End If
  Case Is = Dsg
    If FraDsg.Visible = True Then
      FraDsg.Visible = False
      Call EnaDisaCmds(False, CmdDsg, Dsg)
      CmdDsg.SetFocus
    Else
      FraDsg.Visible = True
      adc("wDmTcTyp").SetFocus
      FraDsg.ZOrder
      Call EnaDisaCmds(True, CmdDsg)
    End If
  Case Is = LabRt
    If FraLabRt.Visible = True Then
      FraLabRt.Visible = False
      Call EnaDisaCmds(False, CmdLabRt, LabRt)
      CmdLabRt.SetFocus
    Else
      FraLabRt.Visible = True
      adc("wLrCmCtgFr").SetFocus
      FraLabRt.ZOrder
      Call EnaDisaCmds(True, CmdLabRt)
    End If
  Case Is = RmRt
    If FraRmRt.Visible = True Then
      FraRmRt.Visible = False
      Call EnaDisaCmds(False, CmdRmRt, RmRt)
      CmdRmRt.SetFocus
    Else
      FraRmRt.Visible = True
      adc("wRrTcTyp").SetFocus
      FraRmRt.ZOrder
      Call EnaDisaCmds(True, CmdRmRt)
    End If
  Case Is = LabRtHist
    If FraLabRtHist.Visible = True Then
      FraLabRtHist.Visible = False
      Call EnaDisaCmds(False, CmdLabRtHist, LabRtHist)
      CmdLabRtHist.SetFocus
    Else
      FraLabRtHist.Visible = True
      adc("wLhYyyyMm").SetFocus
      FraLabRtHist.ZOrder
      Call EnaDisaCmds(True, CmdLabRtHist)
    End If
  Case Is = RmRtHist
    If FraRmRtHist.Visible = True Then
      FraRmRtHist.Visible = False
      Call EnaDisaCmds(False, CmdRmRtHist, RmRtHist)
      CmdRmRtHist.SetFocus
    Else
      FraRmRtHist.Visible = True
      adc("wRhTcTyp").SetFocus
      FraRmRtHist.ZOrder
      Call EnaDisaCmds(True, CmdRmRtHist)
    End If
  Case Is = DtTbl
    If FraDtTbl.Visible = True Then
      FraDtTbl.Visible = False
      Call EnaDisaCmds(False, CmdDtTbl, DtTbl)
      CmdDtTbl.SetFocus
    Else
      FraDtTbl.Visible = True
      adc("wDtTag").SetFocus
      FraDtTbl.ZOrder
      Call EnaDisaCmds(True, CmdDtTbl)
    End If
  Case Is = BkAg
    If FraBkAg.Visible = True Then
      FraBkAg.Visible = False
      Call EnaDisaCmds(False, CmdBkAg, BkAg)
      CmdBkAg.SetFocus
    Else
      FraBkAg.Visible = True
      adc("wBkCtg").SetFocus
      FraBkAg.ZOrder
      Call EnaDisaCmds(True, CmdBkAg)
    End If
  Case Is = Loc
    If FraLoc.Visible = True Then
      FraLoc.Visible = False
      Call EnaDisaCmds(False, CmdLoc, Loc)
      CmdLoc.SetFocus
    Else
      FraLoc.Visible = True
      adc("wLocTyp").SetFocus
      FraLoc.ZOrder
      Call EnaDisaCmds(True, CmdLoc)
    End If
  Case Is = Prm
    If FraPrm.Visible = True Then
      FraPrm.Visible = False
      Call EnaDisaCmds(False, CmdPrm, Prm)
      CmdPrm.SetFocus
    Else
      FraPrm.Visible = True
      adc("wPCtg").SetFocus
      FraPrm.ZOrder
      Call EnaDisaCmds(True, CmdPrm)
    End If
  Case Is = VPrm
    If FraVPrm.Visible = True Then
      FraVPrm.Visible = False
      Call EnaDisaCmds(False, CmdVPrm, VPrm)
      CmdVPrm.SetFocus
    Else
      FraVPrm.Visible = True
      adc("wvPCtg").SetFocus
      FraVPrm.ZOrder
      Call EnaDisaCmds(True, CmdVPrm)
    End If
  Case Is = Ord
    If FraOrd.Visible = True Then
      FraOrd.Visible = False
      Call EnaDisaCmds(False, CmdOrd, Ord)
      CmdOrd.SetFocus
    Else
      FraOrd.Visible = True
      adc("wOmTcTyp").SetFocus
      FraOrd.ZOrder
      Call EnaDisaCmds(True, CmdOrd)
    End If
  Case Is = Inv
    If FraInv.Visible = True Then
      FraInv.Visible = False
      Call EnaDisaCmds(False, CmdInv, Inv)
      CmdInv.SetFocus
    Else
      FraInv.Visible = True
      adc("wInTcFr").SetFocus
      FraInv.ZOrder
      Call EnaDisaCmds(True, CmdInv)
    End If
  Case Is = AllMst
    If FraAllMst.Visible = True Then
      FraAllMst.Visible = False
      Call EnaDisaCmds(False, CmdAllMst, AllMst)
      CmdAllMst.SetFocus
    Else
      FraAllMst.Visible = True
      adc("wAllMstModDtFr").SetFocus
      FraAllMst.ZOrder
      Call EnaDisaCmds(True, CmdAllMst)
    End If
  End Select
End Sub
Private Sub EnaDisaCmds(ByVal pv_ShowFra As Boolean, Optional ByVal pv_CmdBut As CommandButton, Optional ByVal pv_DispFra As en_EmrTrfOutDispFra)
  '*** This procedure is used to give the toggle option (enable and disable) to
      'the command buttons

  '*** storing the state of all cmd buttons, when the frame is going to be visible.
  Dim X As Integer
  If pv_ShowFra = True Then
    Call StoreState
  End If

  If Not pv_CmdBut Is Nothing Then
    If pv_ShowFra = True Then
        pv_CmdBut.BackColor = vbWhite
        pv_CmdBut.FontBold = True
        pv_CmdBut.CausesValidation = True
    Else
        pv_CmdBut.BackColor = &H8000000F
        pv_CmdBut.FontBold = False
        '*** restoring the 'Causes Validation' state of the current pv_cmdbutton
        If Not IsMissing(pv_DispFra) Then
          pv_CmdBut.CausesValidation = mArr_Cv(pv_DispFra)
        End If
    End If
  End If
  
  ''FraNKeyAll.Enabled = Not pv_ShowFra
  ''ADC.AllowSave = Not pv_ShowFra
  ''ADC.AllowDelete = Not pv_ShowFra
  
  If pv_ShowFra = True Then
    '** All command buttons are disabled
    CmdRm.Enabled = Not pv_ShowFra
    CmdCust.Enabled = Not pv_ShowFra
    CmdDsg.Enabled = Not pv_ShowFra
    CmdRmRt.Enabled = Not pv_ShowFra
    CmdLabRt.Enabled = Not pv_ShowFra
    CmdRmRtHist.Enabled = Not pv_ShowFra
    CmdLabRtHist.Enabled = Not pv_ShowFra
    CmdDtTbl.Enabled = Not pv_ShowFra
    CmdBkAg.Enabled = Not pv_ShowFra
    CmdLoc.Enabled = Not pv_ShowFra
    CmdPrm.Enabled = Not pv_ShowFra
    CmdVPrm.Enabled = Not pv_ShowFra
    CmdOrd.Enabled = Not pv_ShowFra
    CmdInv.Enabled = Not pv_ShowFra
    CmdAllMst.Enabled = Not pv_ShowFra
  Else
    '*** Restoring Ena/Disa state of all Command Buttons
    CmdRm.Enabled = mArr_Ed(Rm)
    CmdCust.Enabled = mArr_Ed(Cust)
    CmdDsg.Enabled = mArr_Ed(Dsg)
    CmdRmRt.Enabled = mArr_Ed(RmRt)
    CmdLabRt.Enabled = mArr_Ed(LabRt)
    CmdRmRtHist.Enabled = mArr_Ed(RmRtHist)
    CmdLabRtHist.Enabled = mArr_Ed(LabRtHist)
    CmdDtTbl.Enabled = mArr_Ed(DtTbl)
    CmdBkAg.Enabled = mArr_Ed(BkAg)
    CmdLoc.Enabled = mArr_Ed(Loc)
    CmdPrm.Enabled = mArr_Ed(Prm)
    CmdVPrm.Enabled = mArr_Ed(VPrm)
    CmdOrd.Enabled = mArr_Ed(Ord)
    CmdInv.Enabled = mArr_Ed(Inv)
    CmdAllMst.Enabled = mArr_Ed(AllMst)
  End If

  If Not pv_CmdBut Is Nothing Then pv_CmdBut.Enabled = True
End Sub
Private Sub HideAllFras()
  '*** This procedure hides all the frames other than FraNKeyAll
  'FraNKeyAll.Visible = True
  FraRm.Visible = False
  FraCust.Visible = False
  FraDsg.Visible = False
  FraRmRt.Visible = False
  FraLabRt.Visible = False
  FraRmRtHist.Visible = False
  FraLabRtHist.Visible = False
  FraDtTbl.Visible = False
  FraBkAg.Visible = False
  FraLoc.Visible = False
  FraPrm.Visible = False
  FraVPrm.Visible = False
  FraOrd.Visible = False
  FraInv.Visible = False
  FraAllMst.Visible = False
End Sub
Private Sub StoreState()
  '*** This procedure is used to store the CausesValidation State and the Enabled/Disabled
      'states of all the Command Buttons
  '*** This is used for restoring the original states in the future
  mArr_Cv(Rm) = CmdRm.CausesValidation
  mArr_Cv(Cust) = CmdCust.CausesValidation
  mArr_Cv(Dsg) = CmdDsg.CausesValidation
  mArr_Cv(RmRt) = CmdRmRt.CausesValidation
  mArr_Cv(LabRt) = CmdLabRt.CausesValidation
  mArr_Cv(RmRtHist) = CmdRmRtHist.CausesValidation
  mArr_Cv(LabRtHist) = CmdLabRtHist.CausesValidation
  mArr_Cv(DtTbl) = CmdDtTbl.CausesValidation
  mArr_Cv(BkAg) = CmdBkAg.CausesValidation
  mArr_Cv(Loc) = CmdLoc.CausesValidation
  mArr_Cv(Prm) = CmdPrm.CausesValidation
  mArr_Cv(VPrm) = CmdVPrm.CausesValidation
  mArr_Cv(Ord) = CmdOrd.CausesValidation
  mArr_Cv(Inv) = CmdInv.CausesValidation
  mArr_Cv(AllMst) = CmdAllMst.CausesValidation
  
  mArr_Ed(Rm) = CmdRm.Enabled
  mArr_Ed(Cust) = CmdCust.Enabled
  mArr_Ed(Dsg) = CmdDsg.Enabled
  mArr_Ed(RmRt) = CmdRmRt.Enabled
  mArr_Ed(LabRt) = CmdLabRt.Enabled
  mArr_Ed(RmRtHist) = CmdRmRtHist.Enabled
  mArr_Ed(LabRtHist) = CmdLabRtHist.Enabled
  mArr_Ed(DtTbl) = CmdDtTbl.Enabled
  mArr_Ed(BkAg) = CmdBkAg.Enabled
  mArr_Ed(Loc) = CmdLoc.Enabled
  mArr_Ed(Prm) = CmdPrm.Enabled
  mArr_Ed(VPrm) = CmdVPrm.Enabled
  mArr_Ed(Ord) = CmdOrd.Enabled
  mArr_Ed(Inv) = CmdInv.Enabled
  mArr_Ed(AllMst) = CmdAllMst.Enabled
End Sub

'*** Std Code Not To Be Changed
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
  ws_BtnStr = ""
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
End Sub
Private Sub Form_Deactivate()
  Call FrmDeActivate(Me)
End Sub
Private Sub ATXT_Validate(Index As Integer, Cancel As Boolean)
  Cancel = ATXT(Index).Validate
End Sub
Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    If UnloadMode <> vbFormCode Then Cancel = True: Exit Sub
    
    '*** (09/08/05)
    Set moCn = Nothing   '*** (Bef speed) bef 24/11/06
    Set XmlDoc = Nothing
    Set xmlRootElement = Nothing
    Set xmlChildX = Nothing
    Set xmlChildA = Nothing
    Set xmlChildB = Nothing
    Set xmlChildC = Nothing
    Set xmlChildB1 = Nothing
    Set xmlChildC1 = Nothing
    Set xmlChildD = Nothing
    Set xmlChildE = Nothing
    Set xmlChildG = Nothing
    '*** (09/08/05)

End Sub
'***

' Zubin 213
Private Function ValidateCurrVer(ByVal mi_MajorVer As Integer, ByVal mi_MinorVer As Integer) As Boolean
  ValidateCurrVer = IIF((mi_MajVer > mi_MajorVer Or (mi_MajVer = mi_MajorVer And mi_MinVer >= mi_MinorVer)), True, False)
End Function

' Zubin 213
Private Function ValidateTgtVer(ByVal mi_TgtMajorVer As Integer, ByVal mi_TgtMinorVer As Integer) As Boolean
  ValidateTgtVer = IIF((mi_TgtMajVer > mi_TgtMajorVer Or (mi_TgtMajVer = mi_TgtMajorVer And mi_TgtMinVer >= mi_TgtMinorVer)), True, False)
End Function

