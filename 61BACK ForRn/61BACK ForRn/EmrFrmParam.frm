VERSION 5.00
Object = "{AB2CC32A-E652-42A2-B341-26ECE02F6386}#2.0#0"; "MwfCtl.ocx"
Begin VB.Form EmrFrmParam 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Parameter Entry"
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
   Begin VB.Frame FraUpdDsgLab 
      Enabled         =   0   'False
      Height          =   1470
      Left            =   4470
      TabIndex        =   60
      Top             =   7680
      Visible         =   0   'False
      Width           =   7260
      Begin MwfCtl.MWCTL_MED atxt 
         Height          =   285
         Index           =   9
         Left            =   4770
         TabIndex        =   67
         ToolTipText     =   "Enter Lab SCd"
         Top             =   510
         Width           =   1440
         _ExtentX        =   2540
         _ExtentY        =   503
         Enabled         =   0   'False
         MaxLength       =   8
         DataType        =   4
         IdName          =   "WDSGLABSCD"
      End
      Begin MwfCtl.MWCTL_BTN1 cmdUpdDsgLabGo 
         Height          =   435
         Left            =   6510
         TabIndex        =   72
         Top             =   510
         Width           =   615
         _ExtentX        =   1085
         _ExtentY        =   767
         BackColor       =   -2147483633
         MaskColor       =   8454143
         Caption         =   "Go"
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
         Index           =   10
         Left            =   4770
         TabIndex        =   69
         ToolTipText     =   "Enter Design Category"
         Top             =   795
         Width           =   1200
         _ExtentX        =   2117
         _ExtentY        =   503
         MaxLength       =   5
         DataType        =   4
         IdName          =   "WDSGLABDMCTG"
      End
      Begin MwfCtl.MWCTL_MED atxt 
         Height          =   285
         Index           =   11
         Left            =   1620
         TabIndex        =   70
         ToolTipText     =   "Enter From  Design Code"
         Top             =   1050
         Width           =   2130
         _ExtentX        =   3757
         _ExtentY        =   503
         MaxLength       =   15
         DataType        =   4
         IdName          =   "WDSGLABDMCDFR"
      End
      Begin MwfCtl.MWCTL_MED atxt 
         Height          =   285
         Index           =   12
         Left            =   4770
         TabIndex        =   71
         ToolTipText     =   "Enter To Design Code"
         Top             =   1080
         Width           =   2340
         _ExtentX        =   4128
         _ExtentY        =   503
         MaxLength       =   15
         DataType        =   4
         IdName          =   "WDSGLABDMCDTO"
      End
      Begin MwfCtl.MWCTL_MED atxt 
         Height          =   285
         Index           =   13
         Left            =   1620
         TabIndex        =   66
         ToolTipText     =   "Enter Lab MCd"
         Top             =   480
         Width           =   1620
         _ExtentX        =   2858
         _ExtentY        =   503
         Enabled         =   0   'False
         MaxLength       =   8
         DataType        =   4
         IdName          =   "WDSGLABMCD"
      End
      Begin MwfCtl.MWCTL_MED atxt 
         Height          =   285
         Index           =   14
         Left            =   1620
         TabIndex        =   68
         ToolTipText     =   "Enter DM TcTyp"
         Top             =   765
         Width           =   720
         _ExtentX        =   1270
         _ExtentY        =   503
         MaxLength       =   3
         DataType        =   4
         IdName          =   "WDSGLABDMTCTYP"
      End
      Begin VB.Label ALBL 
         BackStyle       =   0  'Transparent
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Add  Lab Sub Cd to Design Master"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000E&
         Height          =   285
         Index           =   11
         Left            =   30
         TabIndex        =   74
         Top             =   150
         Width           =   6045
      End
      Begin VB.Label ALBL 
         BackStyle       =   0  'Transparent
         Caption         =   "Design TcTyp"
         BeginProperty Font 
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
         Left            =   60
         TabIndex        =   73
         Top             =   810
         Width           =   1515
      End
      Begin VB.Label ALBL 
         BackStyle       =   0  'Transparent
         Caption         =   "Lab MainCd"
         BeginProperty Font 
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
         Left            =   60
         TabIndex        =   65
         Top             =   510
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
         Height          =   345
         Index           =   5
         Left            =   4500
         TabIndex        =   64
         Top             =   1140
         Width           =   465
      End
      Begin VB.Label ALBL 
         BackStyle       =   0  'Transparent
         Caption         =   "Design Cd From"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Index           =   4
         Left            =   60
         TabIndex        =   63
         Top             =   1080
         Width           =   1665
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
         Index           =   3
         Left            =   3660
         TabIndex        =   62
         Top             =   810
         Width           =   1635
      End
      Begin VB.Label ALBL 
         BackStyle       =   0  'Transparent
         Caption         =   "Lab SubCd"
         BeginProperty Font 
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
         Left            =   3660
         TabIndex        =   61
         Top             =   540
         Width           =   1275
      End
   End
   Begin MwfCtl.MWCTL_BTN1 CmdDelDsgAna 
      Height          =   435
      Left            =   4320
      TabIndex        =   56
      Top             =   9300
      Visible         =   0   'False
      Width           =   990
      _ExtentX        =   1746
      _ExtentY        =   767
      BackColor       =   -2147483633
      MaskColor       =   8454143
      Caption         =   "Delete DsgAna"
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
   Begin MwfCtl.MWCTL_BTN1 CmdUpdDsgAna 
      Height          =   435
      Left            =   5400
      TabIndex        =   51
      Top             =   9300
      Visible         =   0   'False
      Width           =   990
      _ExtentX        =   1746
      _ExtentY        =   767
      BackColor       =   -2147483633
      MaskColor       =   8454143
      Caption         =   "Update DsgAna"
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
   Begin MwfCtl.MWCTL_BTN1 cmdReAllotPrt 
      Height          =   435
      Left            =   6510
      TabIndex        =   18
      Top             =   9300
      Visible         =   0   'False
      Width           =   990
      _ExtentX        =   1746
      _ExtentY        =   767
      BackColor       =   -2147483633
      MaskColor       =   8454143
      Caption         =   "&Reallot Prt"
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
   Begin MwfCtl.MWCTL_BTN1 CmdPwd 
      Height          =   435
      Left            =   7500
      TabIndex        =   23
      ToolTipText     =   "Enter Password to Edit Design"
      Top             =   9300
      Width           =   960
      _ExtentX        =   1693
      _ExtentY        =   767
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
   Begin MwfCtl.MWCTL_HLP MWCTL_HLP1 
      Height          =   435
      Left            =   9960
      TabIndex        =   22
      TabStop         =   0   'False
      Top             =   9315
      Visible         =   0   'False
      Width           =   435
      _ExtentX        =   767
      _ExtentY        =   767
   End
   Begin MwfCtl.MWCTL_ADC ADC 
      CausesValidation=   0   'False
      Height          =   555
      Left            =   60
      TabIndex        =   26
      Top             =   9300
      Width           =   15060
      _ExtentX        =   26564
      _ExtentY        =   979
      AllowAdd        =   0   'False
      AllowDelete     =   0   'False
      BtnWidth        =   885
      BtnHeight       =   435
      BtnMaskColor    =   8454143
      FontSz          =   8.25
   End
   Begin VB.Frame FRA_KEYS 
      BorderStyle     =   0  'None
      Height          =   300
      Left            =   -36
      TabIndex        =   28
      Top             =   24
      Width           =   5700
      Begin MwfCtl.MWCTL_MED atxt 
         Height          =   285
         Index           =   1
         Left            =   3225
         TabIndex        =   20
         ToolTipText     =   "Enter Parameter Type"
         Top             =   0
         Width           =   1380
         _ExtentX        =   2434
         _ExtentY        =   503
         MaxLength       =   8
         DataType        =   4
         ReCalcOn        =   "WENTRYTYP"
         IdName          =   "WPTYP"
         CmpStr          =   "PTyp="
      End
      Begin MwfCtl.MWCTL_MED atxt 
         Height          =   285
         Index           =   0
         Left            =   1305
         TabIndex        =   19
         ToolTipText     =   "Specify Entry Type"
         Top             =   0
         Width           =   360
         _ExtentX        =   635
         _ExtentY        =   503
         DataType        =   4
         ReCalcParent    =   "WPTYP"
         IdName          =   "WENTRYTYP"
      End
      Begin VB.Label ALBL 
         BackStyle       =   0  'Transparent
         Caption         =   "Entry Type"
         BeginProperty Font 
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
         Left            =   150
         TabIndex        =   31
         Top             =   0
         Width           =   1185
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
         Index           =   7
         Left            =   1950
         TabIndex        =   30
         Top             =   0
         Width           =   1305
      End
   End
   Begin VB.Frame FRA_NKEYS 
      BorderStyle     =   0  'None
      Enabled         =   0   'False
      Height          =   9360
      Left            =   -120
      TabIndex        =   27
      Top             =   0
      Width           =   15345
      Begin VB.Frame FraUpdDsgAna 
         Height          =   750
         Left            =   4560
         TabIndex        =   52
         Top             =   8400
         Width           =   4020
         Begin MwfCtl.MWCTL_MED atxt 
            Height          =   285
            Index           =   8
            Left            =   1200
            TabIndex        =   53
            ToolTipText     =   "Enter From  Order Tc "
            Top             =   240
            Width           =   1440
            _ExtentX        =   2540
            _ExtentY        =   503
            MaxLength       =   8
            DataType        =   4
            IdName          =   "WDEFDAANACD"
         End
         Begin MwfCtl.MWCTL_BTN1 CmdUpdDsgAnaGo 
            Height          =   435
            Left            =   3240
            TabIndex        =   54
            Top             =   240
            Width           =   615
            _ExtentX        =   1085
            _ExtentY        =   767
            BackColor       =   -2147483633
            MaskColor       =   8454143
            Caption         =   "Go"
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
            Caption         =   "Def AnaCd"
            BeginProperty Font 
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
            TabIndex        =   55
            Top             =   285
            Width           =   1335
         End
      End
      Begin VB.Frame FraNKeyAll 
         BorderStyle     =   0  'None
         Height          =   9225
         Left            =   240
         TabIndex        =   29
         Top             =   90
         Width           =   15150
         Begin VB.Frame FraRTCLst 
            BorderStyle     =   0  'None
            Height          =   2655
            Left            =   0
            TabIndex        =   57
            Top             =   6480
            Visible         =   0   'False
            Width           =   15135
            Begin VB.TextBox TxtRTCLst 
               BackColor       =   &H8000000F&
               BorderStyle     =   0  'None
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   2415
               Left            =   0
               MultiLine       =   -1  'True
               TabIndex        =   58
               Text            =   "EmrFrmParam.frx":0000
               Top             =   160
               Width           =   3495
            End
         End
         Begin MwfCtl.MWCTL_FLX GrdParam 
            Height          =   8295
            Left            =   0
            TabIndex        =   21
            Top             =   240
            Width           =   15135
            _ExtentX        =   26696
            _ExtentY        =   14631
            Cols            =   12
            colname1        =   "PMCD"
            heading1        =   "Main Cd         "
            datafld1        =   "PMCd"
            datatype1       =   4
            maxlength1      =   8
            tooltiptext1    =   "Enter Param Main Code"
            colname2        =   "PSCD"
            heading2        =   "Sub Cd           "
            datafld2        =   "PSCd"
            datatype2       =   4
            maxlength2      =   8
            tooltiptext2    =   "Enter Param Sub Code"
            colname3        =   "PDESC"
            heading3        =   "Desc                     "
            datafld3        =   "PDesc"
            recalcparent3   =   "GrdParam(PVALUE1)"
            maxlength3      =   30
            tooltiptext3    =   "Enter Description"
            colname4        =   "PNUM"
            heading4        =   "Number"
            datafld4        =   "PNum"
            datatype4       =   2
            mask4           =   "######0.0000"
            maxlength4      =   12
            tooltiptext4    =   "Enter PNum"
            colname5        =   "PDESC225"
            heading5        =   "Desc225                "
            datafld5        =   "PDesc225"
            datatype5       =   4
            recalcon5       =   "WINVDESC"
            maxlength5      =   225
            tooltiptext5    =   "Enter PDesc225"
            colname6        =   "PVALUE"
            heading6        =   "Value                     "
            datafld6        =   "PValue"
            datatype6       =   4
            maxlength6      =   30
            tooltiptext6    =   "Enter PValue"
            colname7        =   "PVALUE1"
            heading7        =   "Value1                   "
            datafld7        =   "PValue1"
            datatype7       =   4
            recalcon7       =   "GrdParam(PDESC)"
            maxlength7      =   120
            tooltiptext7    =   "Enter PValue1"
            colname8        =   "PNUM1"
            heading8        =   "Number1           "
            datafld8        =   "PNum1"
            datatype8       =   2
            mask8           =   "######0.0000"
            maxlength8      =   12
            tooltiptext8    =   "Enter PNum1"
            colname9        =   "PVALUE2"
            heading9        =   "Value2                   "
            datafld9        =   "PValue2"
            datatype9       =   4
            maxlength9      =   30
            tooltiptext9    =   "Enter PValue2"
            colname10       =   "PVALUE3"
            heading10       =   "Value3                 "
            datafld10       =   "PValue3"
            datatype10      =   4
            maxlength10     =   30
            tooltiptext10   =   "Enter PValue3"
            colname11       =   "PVALIDYN"
            heading11       =   "Yn"
            datafld11       =   "PValidYN"
            datatype11      =   4
            maxlength11     =   1
            style11         =   0
            tooltiptext11   =   "Enter Parameter Valid (Y/N)"
         End
         Begin MwfCtl.MWCTL_MED atxt 
            Height          =   600
            Index           =   23
            Left            =   2130
            TabIndex        =   32
            ToolTipText     =   "Enter Purchase Order No"
            Top             =   8520
            Visible         =   0   'False
            Width           =   9750
            _ExtentX        =   17198
            _ExtentY        =   1058
            Alignment       =   3
            Enabled         =   0   'False
            MaxLength       =   180
            ReCalcParent    =   "GrdParam(PDESC225)"
            IdName          =   "WINVDESC"
         End
         Begin VB.Label lblDesc 
            BackStyle       =   0  'Transparent
            Caption         =   "Detail Product Desc"
            BeginProperty Font 
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
            TabIndex        =   33
            Top             =   8490
            Visible         =   0   'False
            Width           =   2130
         End
      End
      Begin VB.Frame FraReAllotPrt 
         Height          =   2685
         Left            =   4590
         TabIndex        =   35
         Top             =   6525
         Visible         =   0   'False
         Width           =   5940
         Begin MwfCtl.MWCTL_BTN1 CmdRpGo 
            Height          =   285
            Left            =   5040
            TabIndex        =   17
            Top             =   2130
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
         Begin MwfCtl.MWCTL_MED atxt 
            Height          =   285
            Index           =   77
            Left            =   2580
            TabIndex        =   1
            ToolTipText     =   "Enter From  Order Year "
            Top             =   390
            Width           =   360
            _ExtentX        =   635
            _ExtentY        =   503
            MaxLength       =   2
            DataType        =   4
            IdName          =   "WODYYFR"
         End
         Begin MwfCtl.MWCTL_MED atxt 
            Height          =   285
            Index           =   79
            Left            =   1860
            TabIndex        =   6
            ToolTipText     =   "Enter From  Order Number"
            Top             =   675
            Width           =   780
            _ExtentX        =   1376
            _ExtentY        =   503
            Alignment       =   2
            Mask            =   "#####0"
            MaxLength       =   6
            DataType        =   1
            IdName          =   "WODNOFR"
         End
         Begin MwfCtl.MWCTL_MED atxt 
            Height          =   285
            Index           =   83
            Left            =   1860
            TabIndex        =   0
            ToolTipText     =   "Enter From  Order Tc "
            Top             =   390
            Width           =   600
            _ExtentX        =   1058
            _ExtentY        =   503
            MaxLength       =   3
            DataType        =   4
            IdName          =   "WODTCFR"
         End
         Begin MwfCtl.MWCTL_MED atxt 
            Height          =   285
            Index           =   84
            Left            =   3030
            TabIndex        =   2
            ToolTipText     =   "Enter From  Order Character"
            Top             =   390
            Width           =   600
            _ExtentX        =   1058
            _ExtentY        =   503
            MaxLength       =   3
            DataType        =   4
            IdName          =   "WODCHRFR"
         End
         Begin MwfCtl.MWCTL_MED atxt 
            Height          =   285
            Index           =   85
            Left            =   2715
            TabIndex        =   7
            ToolTipText     =   "Enter From  Order Serial Number"
            Top             =   675
            Width           =   570
            _ExtentX        =   1005
            _ExtentY        =   503
            Alignment       =   2
            Mask            =   "###0"
            MaxLength       =   4
            DataType        =   1
            IdName          =   "WODSRFR"
         End
         Begin MwfCtl.MWCTL_MED atxt 
            Height          =   285
            Index           =   86
            Left            =   4590
            TabIndex        =   4
            ToolTipText     =   "Enter To  Order Year"
            Top             =   390
            Width           =   360
            _ExtentX        =   635
            _ExtentY        =   503
            MaxLength       =   2
            DataType        =   4
            IdName          =   "WODYYTO"
         End
         Begin MwfCtl.MWCTL_MED atxt 
            Height          =   285
            Index           =   87
            Left            =   3900
            TabIndex        =   8
            ToolTipText     =   "Enter To  Order Number"
            Top             =   675
            Width           =   780
            _ExtentX        =   1376
            _ExtentY        =   503
            Alignment       =   2
            Mask            =   "#####0"
            MaxLength       =   6
            DataType        =   1
            IdName          =   "WODNOTO"
         End
         Begin MwfCtl.MWCTL_MED atxt 
            Height          =   285
            Index           =   88
            Left            =   3900
            TabIndex        =   3
            ToolTipText     =   "Enter To Order Tc"
            Top             =   390
            Width           =   600
            _ExtentX        =   1058
            _ExtentY        =   503
            MaxLength       =   3
            DataType        =   4
            IdName          =   "WODTCTO"
         End
         Begin MwfCtl.MWCTL_MED atxt 
            Height          =   285
            Index           =   89
            Left            =   5040
            TabIndex        =   5
            ToolTipText     =   "Enter To  Order Character"
            Top             =   390
            Width           =   600
            _ExtentX        =   1058
            _ExtentY        =   503
            MaxLength       =   3
            DataType        =   4
            IdName          =   "WODCHRTO"
         End
         Begin MwfCtl.MWCTL_MED atxt 
            Height          =   285
            Index           =   90
            Left            =   4785
            TabIndex        =   9
            ToolTipText     =   "Enter To  Order Serial Number"
            Top             =   675
            Width           =   570
            _ExtentX        =   1005
            _ExtentY        =   503
            Alignment       =   2
            Mask            =   "###0"
            MaxLength       =   4
            DataType        =   1
            IdName          =   "WODSRTO"
         End
         Begin MwfCtl.MWCTL_MED atxt 
            Height          =   315
            Index           =   51
            Left            =   1860
            TabIndex        =   10
            ToolTipText     =   "Enter From Export Delivery Date"
            Top             =   960
            Visible         =   0   'False
            Width           =   885
            _ExtentX        =   1561
            _ExtentY        =   556
            MaxLength       =   8
            DataType        =   3
            IdName          =   "WODEXPDELDTFR"
         End
         Begin MwfCtl.MWCTL_MED atxt 
            Height          =   315
            Index           =   4
            Left            =   3900
            TabIndex        =   11
            ToolTipText     =   "Enter To Export Delivery Date"
            Top             =   960
            Visible         =   0   'False
            Width           =   885
            _ExtentX        =   1561
            _ExtentY        =   556
            MaxLength       =   8
            DataType        =   3
            IdName          =   "WODEXPDELDTTO"
         End
         Begin MwfCtl.MWCTL_MED atxt 
            Height          =   315
            Index           =   2
            Left            =   1860
            TabIndex        =   12
            ToolTipText     =   "Enter From  Production Delivery Date"
            Top             =   1275
            Visible         =   0   'False
            Width           =   885
            _ExtentX        =   1561
            _ExtentY        =   556
            MaxLength       =   8
            DataType        =   3
            IdName          =   "WODDELDTFR"
         End
         Begin MwfCtl.MWCTL_MED atxt 
            Height          =   315
            Index           =   3
            Left            =   3900
            TabIndex        =   13
            ToolTipText     =   "Enter To  Production Delivery Date"
            Top             =   1275
            Visible         =   0   'False
            Width           =   885
            _ExtentX        =   1561
            _ExtentY        =   556
            MaxLength       =   8
            DataType        =   3
            IdName          =   "WODDELDTTO"
         End
         Begin MwfCtl.MWCTL_MED atxt 
            Height          =   315
            Index           =   5
            Left            =   1860
            TabIndex        =   14
            ToolTipText     =   "Apply Diamond Percentage Logic ?(Enter Y/N)"
            Top             =   1590
            Width           =   285
            _ExtentX        =   503
            _ExtentY        =   556
            DataType        =   4
            IdName          =   "WDIAISSYN"
         End
         Begin MwfCtl.MWCTL_MED atxt 
            Height          =   315
            Index           =   6
            Left            =   4440
            TabIndex        =   15
            ToolTipText     =   "Apply Date Logic ?(Enter Y/N)"
            Top             =   1590
            Width           =   285
            _ExtentX        =   503
            _ExtentY        =   556
            DataType        =   4
            IdName          =   "WDATEYN"
         End
         Begin MwfCtl.MWCTL_MED atxt 
            Height          =   315
            Index           =   7
            Left            =   1830
            TabIndex        =   16
            ToolTipText     =   "Allot Lower Priority? (Y/N)"
            Top             =   2100
            Width           =   285
            _ExtentX        =   503
            _ExtentY        =   556
            DataType        =   4
            IdName          =   "WALLOTLOWPRT"
         End
         Begin VB.Label LblMain 
            BackStyle       =   0  'Transparent
            BorderStyle     =   1  'Fixed Single
            Caption         =   "  Priority Reallotment"
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
            TabIndex        =   49
            Top             =   90
            Width           =   5940
         End
         Begin VB.Label Label3 
            BackStyle       =   0  'Transparent
            Caption         =   " Allot Lower Priority"
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
            Left            =   30
            TabIndex        =   48
            Top             =   2100
            Width           =   1725
         End
         Begin VB.Label Label2 
            BackStyle       =   0  'Transparent
            Caption         =   " Apply Date Logic "
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
            Left            =   3150
            TabIndex        =   47
            Top             =   1590
            Width           =   1335
         End
         Begin VB.Label Label1 
            BackStyle       =   0  'Transparent
            Caption         =   " Apply Diamond %age Logic "
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
            Left            =   30
            TabIndex        =   46
            Top             =   1590
            Width           =   1725
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
            Index           =   48
            Left            =   4965
            TabIndex        =   45
            Top             =   390
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
            Index           =   47
            Left            =   4500
            TabIndex        =   44
            Top             =   390
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
            Index           =   45
            Left            =   4710
            TabIndex        =   43
            Top             =   675
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
            Index           =   43
            Left            =   2955
            TabIndex        =   42
            Top             =   390
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
            Index           =   26
            Left            =   2490
            TabIndex        =   41
            Top             =   390
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
            Index           =   20
            Left            =   2640
            TabIndex        =   40
            Top             =   675
            Width           =   105
         End
         Begin VB.Label LblPrdDt 
            BackStyle       =   0  'Transparent
            Caption         =   "Prd Del Date"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   30
            TabIndex        =   39
            Top             =   1275
            Visible         =   0   'False
            Width           =   1725
         End
         Begin VB.Label LblExpDt 
            BackStyle       =   0  'Transparent
            Caption         =   "Exp Del Date"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Left            =   30
            TabIndex        =   38
            Top             =   960
            Visible         =   0   'False
            Width           =   1725
         End
         Begin VB.Label ALBL 
            BackStyle       =   0  'Transparent
            Caption         =   "Order Tc/Yy/Chr"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   285
            Index           =   9
            Left            =   30
            TabIndex        =   37
            Top             =   420
            Width           =   1845
         End
         Begin VB.Label ALBL 
            BackStyle       =   0  'Transparent
            Caption         =   "Order No/Sr"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   285
            Index           =   10
            Left            =   30
            TabIndex        =   36
            Top             =   675
            Width           =   1845
         End
      End
   End
   Begin VB.Frame FraPwd 
      Height          =   870
      Left            =   5250
      TabIndex        =   24
      Top             =   8340
      Width           =   4500
      Begin VB.TextBox TxtPwd 
         Height          =   285
         IMEMode         =   3  'DISABLE
         Left            =   2040
         MaxLength       =   30
         PasswordChar    =   "*"
         TabIndex        =   25
         Top             =   405
         Width           =   2340
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
         TabIndex        =   50
         Top             =   90
         Width           =   4530
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
         TabIndex        =   34
         Top             =   405
         Width           =   1935
      End
   End
   Begin MwfCtl.MWCTL_BTN1 CmdUpdDsgLab 
      Height          =   435
      Left            =   3180
      TabIndex        =   59
      Top             =   9300
      Visible         =   0   'False
      Width           =   990
      _ExtentX        =   1746
      _ExtentY        =   767
      BackColor       =   -2147483633
      MaskColor       =   8454143
      Caption         =   "Update Dsg&Lab"
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
Attribute VB_Name = "EmrFrmParam"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'*** (A hyphen and three stars are used for general commenting
'****** (A hyphen and six stars are used for commenting difficult syntax)
'### (A hyphen and three hashes are used when there is some problem which has to be looked into)
'##### (A hyphen and five hashes are used for alternative codes)
'@@@@@ (A hyphen and five 'AtTheRates' are used for temporarily commenting code which is useful)

'--- General Comments ----
' There are two types of parameters under two different menu options :
'    1. System Parameters (These are high level parameters which the normal user has no rights to)
'    2. Parameters (These parameters can be changed by normal users)
'Whether a parameter is a System Parameter or not is specified in the field 'SysParam(Y/N)'
'    of the Parameter Type 'PTYP'
' A- Entry, GrdParam
'TcStr table is used to set the properties of all the fields
'    in the grid GrdParam
' fn GrdPropSet called from Adc_recsource, sets the properties like enable, visible,
'    heading, length and mask of each column by reading from TcStr table.
'
'Once the System Patrameters are entered no further Add,Change or Delete Operations are allowed on them
'----

Option Explicit
Dim moCn As MwfLib.MDOConnection   '*** (Bef speed) bef 24/11/06

Dim ms_EntryTyp As String
Dim ms_PwdFlag As Boolean

'Dim wi_CharWdth As Integer

Enum en_PrmFra
  Pwd = 0
  ReAlloc = 1
  UpdDsgAna = 2
  DelDsgAna = 3 '***********Bhavna
  UpdDsgLab = 4
End Enum

Private Sub ADC_Load()
  '*** Set the First Key Control to the field wPTyp (working field for Parameter Type)
  '*** Set the First Non Key Control to GrdParam
  '*** Set the Child Property Of ADC to GrdParam
  '*** Set the Previous and Next Control Properties for Grid GrdParam as ADC
  '*** Set the Hot Keys for the Grids GrdParam as 'A'
  
  Set moCn = ADC.Connection    '*** (Bef speed) bef 24/11/06
  'Set adc.FirKeyCtl = adc("wPTyp")
  Set ADC.FirKeyCtl = ADC("wEntryTyp")
  Set ADC.FirNKeyCtl = GrdParam
  ADC.Child = "GrdParam"
  Set GrdParam.PrevCtl = ADC
  
  'urmi 2.06 changed for InvDesc To be an x entry
  'Set GrdParam.NextCtl = ADC
  Set GrdParam.NextCtl = ADC("wInvDesc")
  GrdParam.HotKey = "A"
  '*** Jay 2.14 *** (SEO Change)
  ms_PwdFlag = moCn.RecSeek("Select PValue1 from Param where PTyp= 'SEOPWD' and PmCd= 'SEOPWD' and PValue1= ''")
  '*** Jay 2.14 *** (SEO Change)
  
End Sub
Private Sub ADC_KeyWhen()
  '*** Set the Frame FraNKeyAll as the first frame
  If ms_EntryTyp <> "" Then ADC("wEntryTyp") = ms_EntryTyp
  Call EnaDisaCmds(True)
  If CmdPwd.Visible = True Then CmdPwd.Enabled = True
  Call HideAllFras
  FraNKeyAll.ZOrder (0)
  '****************Geeta****************
  If ADC("WPTYP") <> "PRTCD" Then
   cmdReAllotPrt.Visible = False
  End If
  '*************************************
  If ADC("WPTYP") <> "DAANAFLD" Then
   CmdUpdDsgAna.Visible = False
   CmdDelDsgAna.Visible = False         '**********Bhavna
  End If
  If ADC("WPTYP") <> "LABSCD" Then
   CmdUpdDsgLab.Visible = False
  End If
End Sub
Private Sub ADC_KeyValid(pr_Cancel As Boolean, pr_ErrMsg As String)
  '*** Validate PTyp on Menu Cd (i.e. If The Parameter Is System Param Or Other Param) ***
  '*** If the Index of the field is not going to be used then pass 0 as the first
      'parameter of the atxt_FldChange Sub procedure, pass the IdName of the
      'field whose value has to be validated as the second parameter,
      'pass pr_Cancel and pr_ErrMsg as the 3rd and 4th parameters,
      'pass the value which has to be validated as the 5th parameter
  
  'Call atxt_FldChange(0, UCase("wPTyp"), pr_Cancel, pr_ErrMsg, adc("wPTyp"))
  'If pr_Cancel = True Then Exit Sub
  
  If ADC("wEntryTyp") = "" Or ADC("wPTyp") = "" Then pr_Cancel = True: pr_ErrMsg = "Entry Type Or Param Type Cannot Be Blank": Exit Sub
  
  '*** Jay 2.14 *** (SEO Change)
  '  If Trim(TxtPwd.text) = moCn.GetFldVal("Select PValue1 from Param where PTyp= 'SEOPWD' " + _
                         "and PmCd= 'SEOPWD'") Then ms_PwdFlag = True
  If (ADC("wEntryTyp") = "M" Or ADC("wEntryTyp") = "Y") And ms_PwdFlag = False Then
    pr_Cancel = True: pr_ErrMsg = "Enter Proper SEO Password To Edit": Exit Sub
  End If
  '*** Jay 2.14 *** (SEO Change)
  
  '****************Geeta****************
  If ADC("WPTYP") = "PRTCD" Then
  cmdReAllotPrt.Visible = True
  End If
  '*************************************
  If ADC("WPTYP") = "DAANAFLD" Then
  CmdUpdDsgAna.Visible = True
  CmdDelDsgAna.Visible = True    '**bhavna
  End If

  If ADC("WPTYP") = "LABSCD" Then
    CmdUpdDsgLab.Visible = True
  End If


End Sub

Private Sub ADC_NonKeyWhen(pr_Cancel As Boolean, pr_ErrMsg As String)
  
  ms_EntryTyp = ADC("wEntryTyp")
  pr_Cancel = False
  
  If FraPwd.Visible = True Then
    '*** Display or hide the Password frame depending on the toggle satate
    Call DispFra(Pwd)
  End If
  
  Call EnaDisaCmds(False)
  CmdPwd.Enabled = False
  
  '***** urmi 2.06 for InvDesc to be an X Entry
  If UCase(ADC("wPTyp")) = UCase("InvDesc") Then
    lblDesc.Visible = True: ADC("WINVDESC").Visible = True: ADC("WINVDESC").Enabled = True
    GrdParam.Height = 8250 '5880
  Else
    lblDesc.Visible = False: ADC("WINVDESC").Visible = False: ADC("WINVDESC").Enabled = False
    'After ClrScr
    'GrdParam.Height = 6580
    GrdParam.Height = 8880 '6500
    'After ClrScr
  End If
  '*******************************************************
   
  If ADC("WPTYP") = "RTC" Then  'MW.105 - Report names are shown which are having RTC
    FraRTCLst.Visible = True
    TxtRTCLst.BackColor = vbButtonFace
  End If
    
End Sub

Private Sub ATXT_FldWhen(Index As Integer, ByVal IdName As String, Cancel As Boolean, ErrMsg As String)
  '*** Give help for the working field wPTyp depending on the menucode (System Parameters or Other Parameters)
  Select Case IdName
  Case Is = UCase("wDsgLabSCd")
        Call HlpList.PSCd("LABSCD", ADC("wDsgLabMCd"))
  Case Is = UCase("wDsgLabMCd")
        Call HlpList.PMCd("LABMCD")
  Case Is = UCase("wDsgLabDmTcTyp")
        Call HlpList.PMCd("DMTCTYP")
  Case Is = UCase("wDsgLabDmCtg")
        Call HlpList.PMCd("DmCtg")
  Case Is = UCase("wEntryTyp")
    HlpList.PMCd ("SYSDFN")
  Case Is = UCase("wDsgLabDmCd")
    HlpList.DmCd ADC("wDsgLabDmTcTyp"), ADC("wDsgLabDmCtg")
  Case Is = UCase("wPTyp")
    Call HlpList.PMCd("PTYP", "'" + ADC("wEntryTyp") + "'")
    'If UCase(adc.MenuCd) = UCase("SysParam") Then Call HlpList.PMCd("PTYP", "'Y'")
    'If UCase(adc.MenuCd) = UCase("OthParam") Then Call HlpList.PMCd("PTYP", "'N'")
    'If UCase(adc.MenuCd) = UCase("MicParam") Then Call HlpList.PMCd("PTYP", "'M'")
  '********************Geeta******************************
  '****************Emr206***************************
  Case Is = UCase("wOdTcFr"), UCase("wOdTcTo")
    Call HlpList.PMCd("TC", "'SO'")
  Case Is = UCase("wOdYyFr"), UCase("wOdYyTo")
    Call HlpList.PMCd("YY")
  Case Is = UCase("wOdChrFr")
    Call HlpList.vPSCd(gs_CoCd, "CHR", ADC("wOdTcFr"))
  Case Is = UCase("wOdChrTo")
    Call HlpList.vPSCd(gs_CoCd, "CHR", ADC("wOdTcTo"))
  Case Is = UCase("wOdNoFr")
    Call HlpList.OmNo(gs_CoCd, ADC("wOdTcFr"), ADC("wOdYyFr"), ADC("wOdChrFr"))
  Case Is = UCase("wOdNoTo")
    Call HlpList.OmNo(gs_CoCd, ADC("wOdTcFr"), ADC("wOdYyFr"), ADC("wOdChrFr"))
  Case Is = UCase("wOdSrFr")
    Call HlpList.OdSr(gs_CoCd, ADC("wOdTcFr"), ADC("wOdYyFr"), ADC("wOdChrFr"), ADC("wOdNoFr"))
  Case Is = UCase("wOdSrTo")
    Call HlpList.OdSr(gs_CoCd, ADC("wOdTcTo"), ADC("wOdYyTo"), ADC("wOdChrTo"), ADC("wOdNoTo"))
  Case Is = UCase("wDiaIssYN"), UCase("wDateYN"), UCase("WAllotLowPrt")
    Call HlpList.PMCd("YN")
 '**********************************Geeta*********************************************
 Case Is = UCase("WDEFDAANACD")
    If GrdParam.Value(GrdParam.Row, "PValue") = "N" Then
        Cancel = True: ErrMsg = "Cannot Enter Code For This Analysis Field": Exit Sub
     Else
       Call HlpList.PSCd("DAANACD", GrdParam.Value(GrdParam.Row, "PMCd"))
     End If
  End Select
End Sub
Private Sub atxt_ReCalculate(Index As Integer, ByVal IdName As String, ByVal OldVal As Variant, ByVal RowNum As Integer, ByVal SourceName As String, ByVal SourceColName As String)
  Select Case IdName
  Case Is = UCase("wPTyp")
    If ADC.Mode = xNorm Then
      ADC("wPTyp") = ""
    End If
  End Select
End Sub
Private Sub atxt_FldChange(Index As Integer, ByVal IdName As String, Cancel As Boolean, ErrMsg As String, ByVal pv_NewValue As Variant)
  '*** Validating PTyp on Menu Cd (i.e. If The Parameter Is System Param Or Simple Param) ***
  'Dim ws_Sys As String
  Select Case IdName
  Case Is = UCase("wEntryTyp")
    Cancel = Not moCn.RecSeek("Select PMCd From Param where PTyp= 'SYSDFN' and " + _
                              "PMCd = '" + pv_NewValue + "'")
    ErrMsg = "Invalid Entry Type": Exit Sub
  Case Is = UCase("wPTyp")
    'If UCase(adc.MenuCd) = UCase("SysParam") Then ws_Sys = "Y"
    'If UCase(adc.MenuCd) = UCase("OthParam") Then ws_Sys = "N"
    'If UCase(adc.MenuCd) = UCase("MicParam") Then ws_Sys = "M"
    Cancel = (pv_NewValue <> "") And (Not moCn.RecSeek("Select PMCd From Param where " + _
             "PTyp= 'PTYP' and PValue= '" + ADC("wEntryTyp") + "' and PMCd = '" + pv_NewValue + "'"))
    If Cancel = True Then ErrMsg = "Invalid Parameter Type": Exit Sub
    
    '****** Sachin 2.14.0 - [Restriction To Microway Users]
    If UCase(gs_UsrCd) = "MIC" And (pv_NewValue <> "" And UCase(pv_NewValue) <> "USR") Then Cancel = True: ErrMsg = "The User Has No Rights To View This Parameter": Exit Sub
    '****** Sachin 2.14.0 - [Restriction To Microway Users]
    
  '***********************Geeta****************Emr206**********************
  Case Is = UCase("wDiaIssYN"), UCase("wDateYN"), UCase("WAllotLowPrt")
    Cancel = (pv_NewValue <> "") And (Not moCn.RecSeek("Select PMCd From Param where " + _
             "PTyp= 'YN' and  PMCd = '" + pv_NewValue + "'"))
    ErrMsg = "Invalid Option (Enter (Y/N))": Exit Sub
 '***********************Geeta****************Emr206**********************
 Case Is = UCase("WDEFDAANACD")
  Cancel = (pv_NewValue <> "") And (Not moCn.RecSeek("Select PSCd From Param Where " + _
                                    "PTyp= 'DAANACD' And PMCd= '" + GrdParam.Value(GrdParam.Row, "PMCd") + "' and " + _
                                    "PSCd = '" + pv_NewValue + "'"))
      If Cancel = True Then ErrMsg = "Invalid Analysis Code": Exit Sub
  End Select
End Sub
Private Sub cmdReAllotPrt_Click()

 Dim ws_PrtAllotOnPrd As String
 Dim ws_DelDt As String
Call DispFra(ReAlloc)
 ws_PrtAllotOnPrd = moCn.GetFldVal("Select HPrtAllotOnPrdYN from Head " + _
                    "where HCoCd = '" + gs_CoCd + "' " + _
                    "and HCd = '" + ctSelfCmCd + "'")
 If ws_PrtAllotOnPrd = "Y" Then
  ' After Clr Scrn
  'FraReAllotPrt.Caption = "Priority Reallotment Based On Dia/Cs Issued And Prd Del Date"
  LblMain(0).Caption = "Priority Reallotment Based On Dia/Cs Issued And Prd Del Date"
  ' After Clr Scrn
  ADC("wOdDelDtFr").Visible = True
  ADC("wOdDelDtTo").Visible = True
  LblPrdDt.Visible = True
  ADC("wOdExpDelDtFr").Visible = False
  ADC("wOdExpDelDtTo").Visible = False
  LblExpDt.Visible = False
 ElseIf ws_PrtAllotOnPrd = "N" Then
  ' After Clr Scrn
  'FraReAllotPrt.Caption = "Priority Reallotment Based On Dia/Cs Issued And Exp Del Date"
  LblMain(0).Caption = "Priority Reallotment Based On Dia/Cs Issued And Exp Del Date"
  ' After Clr Scrn
  ADC("wOdExpDelDtFr").Visible = True
  ADC("wOdExpDelDtTo").Visible = True
  LblExpDt.Visible = True
  ADC("wOdDelDtFr").Visible = False
  ADC("WOdDelDtTo").Visible = False
  LblPrdDt.Visible = False
 End If
 End Sub

Private Sub CmdRpGo_Click()
'*************************** Geeta**************************
'***Emr206******************************************
 ' This cmd is available only for PTyp 'PRTCD'
 ' if HPrtAllotOnPrdYN from Head is '' then do nothing
 ' if HPrtAllotOnPrdYN is 'Y' then Allot Priority on Production
 ' if HPrtAllotOnPrdYN is 'N' then Allot Priority on Export

  
 Dim ws_PrtAllotOnPrd As String
 Dim ws_DelDt As String
 Dim ws_Prt As String
 Dim ws_AllotLowPrt As String
 ws_PrtAllotOnPrd = moCn.GetFldVal("Select HPrtAllotOnPrdYN from Head " + _
                    "where HCoCd = '" + gs_CoCd + "' " + _
                    "and HCd = '" + ctSelfCmCd + "'")

 If ws_PrtAllotOnPrd = "" Then
  DispMsg "Reallotment Option Set To None In Head", etInfo
  Exit Sub
 ElseIf ws_PrtAllotOnPrd = "Y" Then
  ws_DelDt = "OdDelDt"
 ElseIf ws_PrtAllotOnPrd = "N" Then
  ws_DelDt = "OdExpDelDt"
 End If
  ws_AllotLowPrt = ""
 If ADC("wAllotLowPrt") = "N" Then _
    ws_AllotLowPrt = " and qOdPrtCd>qPrtCd"
 
 If ADC("wDiaIssYN") <> "Y" And ADC("wDateYN") <> "Y" Then
  DispMsg "Any one of Apply Diamond Percentage Logic or Delivery Date Logic must be 'Y'  ", etInfo
  Exit Sub
 End If
 If ADC("wAllotLowPrt") = "" Then
  DispMsg "Allot Lower Priority Cannot be Blank ,Enter (Y/N)  ", etInfo
  Exit Sub
 End If
 
 If ADC("wDiaIssYN") = "Y" Then _
  ws_Prt = "(isnull(qBQty,0)/case when OdPrdQty=0 then 1 else OdPrdQty end *100) between PValue and PValue1 "
 If ADC("wDateYN") = "Y" Then _
   ws_Prt = IIF(ws_Prt <> "", ws_Prt + "or  DateDiff(Day,GetDate()," + ws_DelDt + ") between PValue2 and PValue3 ", " DateDiff(Day,GetDate()," + ws_DelDt + ") between PValue2 and PValue3 ")
  

  Dim wCnd As String
  ADC("WPTYP").CmpStr = ""
  ADC("wOdTcFr").CmpStr = "OdTc>=": ADC("wOdYyFr").CmpStr = "OdYy>=": ADC("wOdChrFr").CmpStr = "OdChr>=": ADC("wOdNoFr").CmpStr = "OdNo>=": ADC("wOdSrFr").CmpStr = "OdSr>="
  ADC("wOdTcTo").CmpStr = "OdTc<=": ADC("wOdYyTo").CmpStr = "OdYy<=": ADC("wOdChrTo").CmpStr = "OdChr<=": ADC("wOdNoTo").CmpStr = "OdNo<=": ADC("wOdSrTo").CmpStr = "OdSr<="
  ADC("wOdDelDtFr").CmpStr = "OdDelDt>=": ADC("wOdDelDtTo").CmpStr = "OdDelDt<="
  ADC("wOdExpDelDtFr").CmpStr = "OdExpDelDt>=": ADC("wOdExpDelDtTo").CmpStr = "OdExpDelDt<="
  wCnd = ADC.RepCond
  wCnd = IIF(wCnd <> "", " and " + wCnd, "")
  ADC("WPTYP").CmpStr = "PTyp="
  
'****** Sachin 3.2.0 - Current Partition Only ******
MWLib.BeginProcess Me, "ReAlloCating Order Priority ..."
moCn.Execute ("Update OrdDsg Set OdPrtCd = qPrtCd from  " + _
                "(select OdPrtKey as qPrtKey, OdIdNo as qOdIdNo, OdCoCd as qOdCoCd,OdTc as qOdTc,OdYy as qOdYy,OdChr as qOdChr,OdNo as qOdNo," + _
                  "OdSr as qOdSr,OdPrtCd as qOdPrtCd,isnull((select Min(PMCD) from Param  " + _
                   "where PTyp='PRTCd' and PSCd=''  and (" + ws_Prt + ")),OdPrtCd) as qPrtCd " + _
                 "from OrdDsg Left Outer join " + _
                 "(select BPrtKey,BOdIdNo,BCoCd,BOdTc,BOdYy,BOdChr,BOdNo,BOdSr,IsNull(sum(BQty),0) qBQty From  Bag  " + _
                "where isnull((Select sum(TdRmQty*(case when TdFrRmLoc= 'W' and TdFrRmDc= 'C' then -1  " + _
                                                 "when TdFrRmLoc= 'W' and TdFrRmDc= 'D' then 1  " + _
                                                 "when TdToRmLoc= 'W' and TdToRmDc= 'C' then -1  " + _
                                                 "when TdToRmLoc= 'W' and TdToRmDc= 'D' then 1   " + _
                                                 "  else 0 end)) " + _
                                "from Txnd  Join RmMst on RmPrtKey=TdPrtKey And RmCd=TdRmCd  " + _
                                "Where TdPrtKey=BPrtKey And TdBIdNo=BIdNo And TdCoCd = BCocd And TdBYy = BYy And TdBChr = BChr  " + _
                                "and TdBNo= BNo and   RmCtg  in ('D','C')),0)>0  " + _
                                "group by BPrtKey,BOdIdNo,BCoCd,BOdTc,BOdYy,BOdChr,BOdNo,BOdSr) Bag  " + _
                 "On BPrtKey=OdPrtKey And BOdIdNo=OdIdNo And BCocd=OdCoCd and BOdTc=OdTc  and BOdYy=OdYy and BOdChr = OdChr  " + _
                 "And BOdNo = OdNo And BOdSr = OdSr  " + _
                 "Where OdPrtKey='" + ctCurrPrtn + "' And OdExpQty<OdPrdQty  and  (select PDesc225 from Param where PTYp='PRTCD' and PMCd=OdPrtCd)='Y' and  " + _
                 "(select PValue from Param where PTyp='TC' and PMCd=OdTc and PSCd='')='SO' and OdCocd='" + gs_CoCd + "'" + wCnd + " ) Ord  " + _
                 "Where OdPrtKey='" + ctCurrPrtn + "' And (select PValue from Param where PTyp='TC' and PMCd=OdTc and PSCd='')='SO' and OdCoCd = qOdCocd And OdTc = qOdTc  " + _
                 "And OdYy = qOdYy And  OdChr = qOdChr And OdNo = qOdNo " + _
                 "and OdSr=qOdSr and OdExpQty<OdPrdQty  and OdCocd='" + gs_CoCd + "'" + ws_AllotLowPrt)
                        
MWLib.EndProcess Me
Call DispFra(ReAlloc)
'****************************************Geeta**************************************
End Sub

Private Sub CmdUpdDsgAna_Click()
    ADC("wDefDAANACD") = ""
    Call DispFra(UpdDsgAna)
End Sub
Private Sub CmdUpdDsgLab_Click()
    ADC("wDsgLabSCd") = ""

  If GrdParam.Row >= 1 And GrdParam.Row <= GrdParam.Rows - 1 Then
    ADC("wDsgLabMCd") = GrdParam.Value(GrdParam.Row, "PMCd")
    ADC("wDsgLabSCd") = GrdParam.Value(GrdParam.Row, "PSCd")
    
  End If
  If ADC("wDsgLabDmTcTyp") = "" Then
    ADC("wDsgLabDmTcTyp") = "DM"
  End If
    Call DispFra(UpdDsgLab)
End Sub

'**************Bhavna
Private Sub CmdDelDsgAna_Click()
    If MsgBox("Delete Design Analysis Records For Analysis Code " + CStr(GrdParam.Value(GrdParam.Row, "PMCd")) + "?", vbQuestion + vbYesNo, "Delete Design Analysis") = vbYes Then
'            Call DispFra(Pwd)
'            If ms_PwdFlag Then
          Me.MousePointer = vbHourglass
        '************ deleting records from DsgAna
        moCn.Execute "DELETE FROM DsgAna where " + _
                    "(DaTcTyp='DM' or DaTcTyp='SM') And DaAnaSr=" + CStr(GrdParam.Value(GrdParam.Row, "PMCd")) + _
                    "And DaPrtKey='" + ctCurrPrtn + "' "
        '****************deleting records from param 'N' (normal)
        moCn.Execute "DELETE FROM Param where " + _
                    "Ptyp = 'DAANACD' And PMCd = " + CStr(GrdParam.Value(GrdParam.Row, "PMCd")) + _
                    "And PPrtKey='" + ctCurrPrtn + "' "

        Me.MousePointer = vbDefault
        MsgBox "Design Analysis Codes Deleted Successfully !! Delete this Parent Code Now."
'        End If
    End If
End Sub
Private Sub CmdUpdDsgLabGo_Click()

  If GrdParam.Row > 1 And GrdParam.Row < GrdParam.Rows - 1 Then
    ADC("wDsgLabMCd") = GrdParam.Value(GrdParam.Row, "PMCd")
    ADC("wDsgLabSCd") = GrdParam.Value(GrdParam.Row, "PSCd")
    
  End If
  If ADC("wDsgLabDmTcTyp") = "" Then
    ADC("wDsgLabDmTcTyp") = "DM"
  End If
  
 If ADC("wDsgLabDmTcTyp") = "" Or Not moCn.RecSeek("Select * from Param where PTyp='DMTCTYP' and '" + ADC("wDsgLabDmTcTyp") + "' = PMcd and PSCd=''") Then
    DispMsg "Please enter a valid DmTcTyp to proceed", etInfo
    Exit Sub
  End If
  If ADC("wDsgLabMCd") = "" Or Not moCn.RecSeek("Select * from Param where PTyp='LABMCD' and '" + ADC("wDsgLabMCd") + "' = PMcd and PSCd=''") Then
    DispMsg "Please enter a valid Lab MCd to proceed", etInfo
    Exit Sub
  End If
  If ADC("wDsgLabDmTcTyp") <> "DM" And ADC("wDsgLabDmTcTyp") <> "SM" Then
    DispMsg "Please enter DM or SM for DmTcTyp", etInfo
    Exit Sub
  End If
  If ADC("wDsgLabSCd") = "" Or Not moCn.RecSeek("Select * from Param where PTyp='LABSCD' and '" + ADC("wDsgLabMCd") + "' = PMcd and PSCd='" + ADC("wDsgLabSCd") + "' and PValidYN='Y'") Then
    DispMsg "Please enter a valid Lab SCd to proceed", etInfo
    Exit Sub
  End If
    
  If ADC("wDsgLabDmCdTo") = "" Then ADC("wDsgLabDmCdTo") = "ZZZZZZZZZZZZZZZ"
     Dim wSqlStr1 As String
     wSqlStr1 = "   Select count('a') " + _
                "  From DsgMst " + _
                 " Where DmTcTyp='" + ADC("WDSGLABDMTCTYP") + "' and " + IIF(ADC("wDsgLabDmCtg") = "", "1=1", "DmCtg='" + ADC("wDsgLabDmCtg") + "'") + " and DmCd>='" + ADC("wDsgLabDmCdFr") + "' and DmCd <= '" + ADC("wDsgLabDmCdTo") + "' and DmPrtKey='" + ctCurrPrtn + "' " + _
                 "   And Not Exists (Select 'x' From DsgLab Where DlTcTyp=DmTcTyp And DlCd=DmCd And DlSz=DmSz " + _
                 "          And DlMCd='" + ADC("WDSGLABMCD") + "' and DlSCd='" + ADC("WDSGLABSCD") + "')"
      
     Dim wRoughNoOfInserts As Long
     wRoughNoOfInserts = moCn.GetFldVal(wSqlStr1)
     If MsgBox("This will update " + CStr(wRoughNoOfInserts) + " Designs. Proceed?", vbYesNo) = vbNo Then
      Exit Sub
    End If
     Me.MousePointer = vbHourglass
     Dim wSqlStr As String
      wSqlStr = "   Insert into DsgLab (DlTcTyp, DlCd, DlSz, DlSr, DlMCd, DlSCd, DlQty, ModUsr, ModDt, ModTime, DlDmIdNo, DlPrtkey) " + _
                "   Select DmTcTyp, DmCd, DmSz, isnull((select max(DlSr) from DsgLab ForDlIdNo where DlDmIdNo=DsgMst.DmIdNo),0)+1,'" + ADC("WDSGLABMCD") + "','" + ADC("WDSGLABSCD") + "' ,1,'MIC','" + Format(moCn.SrvrDate, "DD/MM/YY") + "', 0.0,DsgMst.DmIdNo,'C'" + _
                "  From DsgMst " + _
                 " Where DmTcTyp='" + ADC("WDSGLABDMTCTYP") + " ' and " + IIF(ADC("wDsgLabDmCtg") = "", "1=1", "DmCtg='" + ADC("wDsgLabDmCtg") + "'") + " and DmCd>='" + ADC("wDsgLabDmCdFr") + "' and DmCd <= '" + ADC("wDsgLabDmCdTo") + "' and DmPrtKey='" + ctCurrPrtn + "' " + _
                 "   And Not Exists (Select 'x' From DsgLab Where DlTcTyp=DmTcTyp And DlCd=DmCd And DlSz=DmSz " + _
                 "          And DlMCd='" + ADC("WDSGLABMCD") + "' and DlSCd='" + ADC("WDSGLABSCD") + "')"
                 
    moCn.Execute wSqlStr
  Me.MousePointer = vbDefault
  MsgBox "Updated around " + CStr(wRoughNoOfInserts) + ". Labor Code inserted Successfully !!"
  
End Sub
Private Sub CmdUpdDsgAnaGo_Click()
If ADC("wDefDAANACD") = "" And GrdParam.Value(GrdParam.Row, "PValue") = "Y" Then
  DispMsg "Default Analysis Code  Cannot be blank.  ", etInfo
  Exit Sub
 End If
 Me.MousePointer = vbHourglass
    moCn.Execute "Insert Into DsgAna (DaTcTyp,DaCd,DaSz,DaAnaSr,DaAnaCd,DaAnaDesc,ModUsr,ModDt,ModTime,DaDmIdNo,DaPrtKey) " + _
                 " Select DmTcTyp, DmCd , DmSz ,'" + CStr(GrdParam.Value(GrdParam.Row, "PMCd")) + "' ,'" + ADC("wDefDAANACD") + "'," + IIF(ADC("wDefDAANACD") = "", "'" + GrdParam.Value(GrdParam.Row, "PDesc225") + "'", "''") + ", 'MIC', '" + Format(moCn.SrvrDate, "DD/MM/YY") + "', 0.0,DmIdNo,'C' " + _
                 " From DsgMst " + _
                 " Where DmTcTyp='" + ADC("WDSGLABDMTCTYP") + " ' and " + IIF(ADC("wDsgLabDmCtg") = "", "1=1", "DmCtg='" + ADC("wDsgLabDmCtg")) + "' and + And DmCdFr'" + ADC("wDsgLabDmCdFr") + " ' and DmCdTo = '" + ADC("wDsgLabDmCdTo") + "' and DmPrtKey='" + ctCurrPrtn + "' " + _
                 "   And Not Exists (Select 'x' From DsgAna Where DaTcTyp=DmTcTyp And DaCd=DmCd And DaSz=DmSz " + _
                 "          And DaAnaSr=" + CStr(GrdParam.Value(GrdParam.Row, "PMCd")) + ")"
    
  Me.MousePointer = vbDefault
  MsgBox "Design Analysis Codes Updated Successfully !!"
  Call DispFra(UpdDsgAna)
End Sub
'*********************Bhavna
Private Sub CmdDelDsgAnaGo_Click()
'If ADC("wDefDAANACD") = "" And GrdParam.Value(GrdParam.Row, "PValue") = "Y" Then
'  DispMsg "Default Analysis Code  Cannot be blank.  ", etInfo
'  Exit Sub
' End If
 
     Me.MousePointer = vbHourglass
    '************ deleting records from DsgAna
    moCn.Execute "DELETE FROM DsgAna where " + _
                    "(DaTcTyp='DM' or DaTcTyp='SM') And DaAnaSr=" + CStr(GrdParam.Value(GrdParam.Row, "PMCd")) + _
                    "And DaPrtKey='" + ctCurrPrtn + "' "
    '****************deleting records from param 'N' (normal)
    moCn.Execute "DELETE FROM Param where " + _
                    "Ptyp = 'DAANACD' And PMCd = " + CStr(GrdParam.Value(GrdParam.Row, "PMCd")) + _
                    "And PPrtKey='" + ctCurrPrtn + "' "

  Me.MousePointer = vbDefault
  MsgBox "Design Analysis Codes Deleted Successfully !! Delete this Parent Code Now."
  Call DispFra(DelDsgAna)
End Sub


Private Sub GrdParam_ColWhen(ByVal RowNum As Integer, ByVal ColName As String, Cancel As Boolean, ErrMsg As String, ByVal Action As MwfCtl.en_fgAction)
  '*** Set Helps For Each Field
  '*** PMCd
      'PTyp= 'CHR'
      '     Show help of all the Tc's
      'PTyp= 'DAANACD'
      '     Show help for Design Analysis Fields
      'PTyp= 'LABSCD'
      '     Show help for Labour Main Code
      'PTyp= 'RMSCTG' or 'RMCOl'
      '     Show all the Raw Material Categories available
      'PTyp= 'SUBPRC'
      '     Show help of all the Process Locations
      'PTyp= 'RWRK'            Bhavna
      '     Show help of all the Process Locations
  '*** PDesc225
      'PTyp= 'CHR'
      '     Show help of all the Locations which are not Process Locations
      'PTyp= 'GLDAS', 'IWTFRORD', 'CPYRT'
      '     MuliSelect= True    (This option allows the user to select multiple values from the help)
      '     Show help for Raw Material Category
      'PTyp= 'PRDSEQ', 'WIWSEQ', 'WIWLONG'
      '     MuliSelect= True
      '     Show help of all the Locations which are Process Locations
      'PTyp= 'RMSCTG'
      '     Show help for 'YN' (Yes/ No). Here the user has to specify whether the
      '     raw materail transactions are Size wise or not
      'PTyp= 'TC'
      '     Show help of all the Characters for the selected 'TC'. This field is used to
      '     store the dafault 'CHR' value for the current 'TC'
      'PTyp= 'INGR'
      '     MuliSelect= True
      '     Show help of all the Group Participants
      'PTyp= 'SRTBAG'
      '     MuliSelect= True
      '     Show help of all the fields in the Bag Table. This field stores the fields
      '     of Bag Table by which sorting has to be done in the Bag Generation Entry
      'PTyp= 'SRTINBAG'
      '     MuliSelect= True
      '     Show help of all the fields in the InvFgd Table. This field stores the fields
      '     of InvFgd Table by which sorting has to be done in the Invoice All Bags Entry
      'PTyp= 'SRTLR'
      '     MuliSelect= True
      '     Show help of all the fields in the LabRt Table. This field stores the fields
      '     of LabRt Table by which sorting has to be done in the Labour Rate Entry
      'PTyp= 'SRTPP'
      '     MuliSelect= True
      '     Show help of all the fields in the PrdPts Table. This field stores the fields
      '     of PrdPts Table by which sorting has to be done in the Production Points Master Entry
      'PTyp= 'SRTPRHLP'
      '     MuliSelect= True
      '     Show help of all the fields in the Txndz Table. This field stores the fields
      '     of Txndz Table by which sorting has to be done in the help of Base Voucher No
      '     field of Custom Matching Entry
      'PTyp= 'SRTRM'
      '     MuliSelect= True
      '     Show help of all the fields in the RmMst Table. This field stores the fields
      '     of RmMst Table by which sorting has to be done in the Raw Material Master Entry
      'PTyp= 'SRTRR'
      '     MuliSelect= True
      '     Show help of all the fields in the RmRt Table. This field stores the fields
      '     of RmRt Table by which sorting has to be done in the Raw Material Rate Master Entry
      'PTyp= 'USR'
      '     MuliSelect= True
      '     Show help of all the production Locs
  '*** PValue
      'PTyp= 'CHR'
      '     Show help of all the Locations which are not Process Locations
      '     This field is used to store the default 'From RmLoc' for that Character. It will
      '     be used in the Transaction Entry for the TC-CHR combination
      'PTyp= 'DAANAFLD'
      '     Show help for Param('YN')
      'PTyp= 'DMCTG'
      '     Show help for Design Unit Of Measurement
      'PTyp= 'LABSCD'
      '     Show help for LABQW. This field specifies whether the calculation of Labour
      '     Value is based on Qty Or Weight
      'PTyp= 'PARENTCD'
      '     Show help for RmCtg
      'PTyp= 'PRDCTG'
      '     Show help for PRDSEQ. This field stores the Sequence of locations the Designs
      '     of this Prd Category has to pass through while in production
      'PTyp= 'PRDSEQ'
      'PTyp= 'REMAKE'
      '     MuliSelect= True    (This option allows the user to select multiple values from the help)
      '     Show help for Raw Material Category
      'PTyp= 'RMCTG'
      '     Show all the Raw Material Codes of the category which are not custom codes.
      '     This field is used to store the Pure Rm Code for this category. This RmCd is
      '     used to calculate the pure wt in many reports.
      'PTyp= 'RMSCTG'
      '     Show help for the 'RATEBY' option. This field specifies how the Rate for
      '     raw materials other than Gold & Platinum are calculated
      '     (i.e By Size, Pointer or Both Size)
      'PTyp= 'TC'
      '     Show the help for TC Type. This field specifies the Tc Type to which this
      '     this field belongs
      'PTyp= 'TCTYP'
      '     Show all the TC's where TcType= PMCd. This field is used to store the default
      '     TC for the current Tc Type
      'PTyp= 'WORK'
      '     Show help of all the Process Locations (Default Process Loc)
  '*** PValue1
      'PTyp= 'CHR'
      '     Show help of all the Process Locations
      'PTyp= 'DAANAFLD'
      '     Show help for Param('YN')
      'PTyp= 'RMSCTG'
      '     Show help for 'YN' (Yes/ No). Here the user has to specify whether the
      '     raw materail transactions are Stock Rate wise or not
      'PTyp= 'PARENTCD'
      '     Show help for RmSCtg
      'PTyp= 'TCTYP'
      '     This field is used to specify the Entry Type of the current Tc Type
      '     The Entry Types are hard coded in the System
      'PTyp= 'USR'
      '     MuliSelect= True
      '     Show help of all the non production Locs
  '*** PValue2
      'PTyp= 'CHR'
      '     Show help of 'DC' (Debit/Credit). This field is used to store the default
      '     Raw Material 'DC' to be used in the Transaction Entry for the TC-CHR combination
      'PTyp= 'RMSCTG'
      '     Show help for 'YN' (Yes/ No). Here the user has to specify whether the
      '     Rm Ptr field in the Design Entry is editable or not
  '*** PValue3
      'PTyp= 'RMSCTG'
      '     Show help for 'YN' (Yes/ No). Here the user has to specify whether the
      '     raw materail transactions are Pointer wise or not
Dim i As Integer
  GrdParam.ToolTipText = ""
  Select Case ColName
  
  Case Is = UCase("PNUM")    'MW.105 - This case provides FROM hours tool tip for RTC only
    Select Case UCase(ADC("wPTyp"))
      Case Is = "RTC"
        DispMsg "Enter hours in 24 hrs format.", etWarning
    End Select
    
  '****** Sachin 3.02.0 - 3 Users Only For Both Partitions
  Case Is = UCase("PNum1")
    Select Case UCase(ADC("wPTyp"))
    Case Is = "USR"
      Call HlpList.PrtnSr
    Case Is = "RTC" 'MW.105 - This case provides TO HRS tool tip for RTC only
      DispMsg "Enter hours in 24 hrs format.", etWarning
    
  '****** Sachin 3.02.0 - 3 Users Only For Both Partitions
    Case Is = "RMSCTG"
      If GrdParam.Value(RowNum, "PDESC225") = "N" Then
        GrdParam.Value(RowNum, "PNum1") = 0
        Cancel = True
        ErrMsg = "Cannot enter NoOfSz when Stk By Sz is No"
      End If
    End Select
  '4.1.2
  'Case Is = UCase("PDesc")
  '  If UCase(ADC("wPTyp")) = "DMMATX" Then _
  '    Call HlpList.Loc("", "'P'")
  
  Case Is = UCase("PMCd")
    Select Case UCase(ADC("wPTyp"))
'    Case Is = "CHR"
'      Call HlpList.PMCd("TC")
    Case Is = "DAANACD"
      Call HlpList.PMCd("DAANAFLD", "'Y'")
    Case Is = "LABSCD", "LABLOGIC"
      Call HlpList.PMCd("LABMCD")
    Case Is = "RMSCTG", "RMCOL"
      Call HlpList.PMCd("RMCTG")
    Case Is = "SUBPRC"
    'uma **** P replaced by R
      Call HlpList.Loc("", "'R'")
  '*********Bhavna
    Case Is = "RWRK"
      Call HlpList.Loc("", "'R'") 'uma **** P replaced by R
   '**********Bhavna
   '*******Geeta********2.10
    Case Is = "STNCLS"
       Call HlpList.PMCd("RMCTG")
   '*******Geeta********2.10
   '*********Geeta***CRM**2.12**
    Case Is = "RMID"
      Call HlpList.PSCd("GRDRMCTG", "", True)
    Case Is = "GRDCD"
      Call HlpList.PMCd("GRDTYP")
    '*********Geeta***CRM**2.12**
    ' Zubin 212 (07/12/05)
    Case Is = "SALCTG"
      Call HlpList.PMCd("DMCTG")
    ' Zubin 212 (07/12/05)
    ' Zubin 213 (was missing Before 2.13)
    Case Is = "GMFACTOR"
      Call HlpList.PMCd("RMCTG")
    ' Zubin 213 (was missing Before 2.13)
    Case Is = "WIWCOL"    ' ****** Manali 3.5.0 - 12/11/08 - WIWCol added
      Call HlpList.PMCd("REPCOL")
    'uma ***
    'Case Is = "CSEQ"
    '  Call HlpList.PMCd("CELL")
    'pg.11-hlp for GstCd
    Case Is = "GSTFLDS"
      Call HlpList.PMCd("GSTCD")
    End Select
  Case Is = UCase("PSCd")
    ' Zubin 213
    Select Case UCase(ADC("wPTyp"))
    Case Is = "GMFACTOR"
      Call HlpList.PMCd("CURNCY")
    Case Is = "RTC"     'MW.105 - This case provides Sr. Number tool tip for RTC only
      DispMsg "Enter Sr. number 1 to 99", etWarning
    'uma ***
    'Case Is = "CSEQ"
    '  Call HlpList.vPMCd("", "PRCSEQ")
    'pg.11-hlp for GstSr
    Case Is = "GSTFLDS"
      Call HlpList.PMCd("GSTSR")
    End Select
    ' Zubin 213
  Case Is = UCase("PDesc225")
    Select Case UCase(ADC("wPTyp"))
    Case Is = "BCHR"
      'IG BCHR- The Value in this field will be allowed only if
      'RMCtg field will have the value as ‘D’ Or ‘C’ or both.
      Call HlpList.PMCd("SUBSHP")
    Case Is = "DSGCPY"
      Call HlpList.PMCd("TCTYP", , "OM")
'    Case Is = "CHR"
'      Call HlpList.Loc(, , "'P'")
    Case Is = "GLDAS", "IWTFRORD", "CPYRT"
      HlpList.MultiSelect = True
      Call HlpList.PMCd("RMCTG")
'    Case Is = "PRDSEQ", "WIWSEQ", "WIWLONG"
'      HlpList.MultiSelect = True
'      Call HlpList.Loc(, "'P'")
    Case Is = "RMSCTG"
      Call HlpList.PMCd("YN")
    'EmrSv - Uma
    Case Is = "DMCTG"
      'HlpList.MultiSelect = True
      Call HlpList.vPMCd("", "PRCSSEQ")
      
    'Case Is = "TC"
    '  Call HlpList.vPMCd(gs_CoCd, "CHR", GrdParam.Value(RowNum, "PMCd"))
    Case Is = "INGR"
      HlpList.MultiSelect = True
      Call HlpList.PMCd("INGRPRCP")
    Case Is = "SRTBAG"
      HlpList.MultiSelect = True
      Call HlpList.TableFields("'Bag'")
    Case Is = "SRTDLD"
      HlpList.MultiSelect = True
      Call HlpList.TableFields("'DsgLoad'")
    Case Is = "SRTINBAG"
      HlpList.MultiSelect = True
      Call HlpList.TableFields("'InvFgd'")
    Case Is = "SRTPP"
      HlpList.MultiSelect = True
      Call HlpList.TableFields("'PrdPts'")
    Case Is = "SRTPRHLP"
      HlpList.MultiSelect = True
      Call HlpList.TableFields("'TxndZ'")
    Case Is = "SRTLR"
      HlpList.MultiSelect = True
      Call HlpList.TableFields("'LabRt'")
    Case Is = "SRTRM"
      HlpList.MultiSelect = True
      Call HlpList.TableFields("'RmMst'")
    Case Is = "SRTRR"
      HlpList.MultiSelect = True
      Call HlpList.TableFields("'RmRt'")
                  
    '*************************** ZUBIN **************************
    ' 29th Nov 2003, EMR206
    Case Is = "SRTRH"
      HlpList.MultiSelect = True
      Call HlpList.TableFields("RmRtHist")
    '*************************** ZUBIN **************************
    
    Case Is = "USR"
      HlpList.MultiSelect = True
      Call HlpList.Loc("", "'P'")
      
    '*************************** ZUBIN **************************
    ' 10th Dec 2003, EMR206 - For Priority Code Gen
    Case Is = "PRTCD"
      Call HlpList.PMCd("YN")
    '*************************** ZUBIN **************************
    'JewelSmith.2 RmSctg help for Ptyp=STNCLS is removed
        
    ' ***** Manali 3.8.0 - Metler
    Case Is = "METLR"
      Call HlpList.PMCd("USR")
      HlpList.MultiSelect = True
      
    Case Is = "RETCHR"      '4.1.4
      Call HlpList.PMCd("USR")
      HlpList.MultiSelect = True
      
    End Select
    
  Case Is = UCase("PValue")
    Select Case UCase(ADC("wPTyp"))
'    Case Is = "CHR"
'      Call HlpList.Loc(, , "'P'")
    Case Is = "DAANAFLD"
      Call HlpList.PMCd("YN")
    Case Is = "DMCTG"
      Call HlpList.PMCd("DMUOM")
    Case Is = "LABSCD"
      Call HlpList.PMCd("LABQW")
    Case Is = "PARENTCD"
      Call HlpList.PMCd("RMCTG")
'    Case Is = "PRDCTG"
'      Call HlpList.PSCd("PRDSEQ", gs_CoCd)
'    Case Is = "PRDSEQ"
'      '$$$ Call HlpList.PMCd("PLNSEQ")
    Case Is = "PTYP"
      Call HlpList.PMCd("SYSDFN")
    'uma ***
    'Case Is = "CSEQ"
    '  Call HlpList.vPMCd("", "PRDSEQ")
      
    '*** (Jen 3.01.1)
    Case Is = "REMAKE"
      HlpList.MultiSelect = True
      Call HlpList.PMCd("RMCTG")
    '*** (Jen 3.01.1)
    'EmrSv - Uma
    Case Is = "PDMGR"
      Call HlpList.PMCd("CELL")
    
    Case Is = "RMCTG"
      Call HlpList.RmCd("N", "'" + GrdParam.Value(RowNum, "PMCd") + "'")
    Case Is = "RMSCTG"
      Call HlpList.PMCd("RATEBY")
    Case Is = "TC"
      Call HlpList.PMCd("TCTYP")
    Case Is = "TCTYP"
      Call HlpList.PMCd("TC", "'" + GrdParam.Value(RowNum, "PMCd") + "'")
'    Case Is = "WORK"
'      Call HlpList.Loc(, "'P'")
    
    '*************************** ZUBIN **************************
    ' 10th Dec 2003, EMR206 - For Priority Code Gen
'    Case Is = "PRTCD"
'      If GrdParam.Value(RowNum, "PDesc225") = "N" Then
'        Cancel = True: ErrMsg = "Cannot Edit Value": Exit Sub
'      End If
    '*************************** ZUBIN **************************
' ########################  Manoj #### Ver: 2.0.6 #### Date: 23/02/2004  ########################
    Case Is = "TAGFLDS"
      Call HlpList.PMCd("TGFLDTYP")
' ###########################################  Manoj  ###########################################
'*** (Jen 2.13) (07/11/06)
' **** Zubin 211 **** '
'''''    '*****Geeta***2.10
    Case Is = "SYSOPT"
      Call HlpList.PMCd("YN")
'''''    '*****Geeta***2.10
' **** Zubin 211 **** '
'*****Geeta***2.11
'*** (Jen 2.13) (07/11/06)

'*** Jay 2.13(CT) ***
    Case Is = "BCHR"
      Call HlpList.PMCd("BTyp")
'*** Jay 2.13(CT) ***
'*****Geeta***2.11
    ' Jen ClrScr
    Case Is = "DEFCLR"
      Call HlpList.CsCd
    ' Jen ClrScr
    'pg.11-hlp for GstBase,GStOn
    Case Is = "GSTCD"
      Call HlpList.PMCd("GSTBASE")
    Case Is = "GSTFLDS"
      Call HlpList.PMCd("GSTON")
    End Select
  Case Is = UCase("PValue1")
    Select Case UCase(ADC("wPTyp"))
'    Case Is = "CHR"
'      Call HlpList.Loc(, "'P'")
    Case Is = "DAANAFLD"
      Call HlpList.PMCd("YN")
    Case Is = "PARENTCD"
      Call HlpList.PSCd("RMSCTG", GrdParam.Value(RowNum, "PValue"))
    Case Is = "RMSCTG"
      Call HlpList.PMCd("YN")
    Case Is = "TCTYP"
      '$$$
    Case Is = "USR"
      HlpList.MultiSelect = True
      Call HlpList.Loc("", , "'P'")
    
    '*************************** ZUBIN **************************
    ' 10th Dec 2003, EMR206 - For Priority Code Gen
'    Case Is = "PRTCD"
'      If GrdParam.Value(RowNum, "PDesc225") = "N" Then
'        Cancel = True: ErrMsg = "Cannot Edit Value": Exit Sub
'      End If
    '*************************** ZUBIN **************************
    ' **** Zubin 211 **** '
    Case Is = "BCHR"
      HlpList.MultiSelect = True  'IG BCHR-Added MultiSelect for RmCtg
      Call HlpList.PMCd("RMCTG")
    ' **** Zubin 211 **** '
    'pg.11-hlp for GstFor
    Case Is = "GSTFLDS"
      Call HlpList.PMCd("GSTFOR")

    End Select
  Case Is = UCase("PValue2")
    Select Case UCase(ADC("wPTyp"))
'    Case Is = "CHR"
'      Call HlpList.PMCd("DC")
    Case Is = "RMSCTG"
      Call HlpList.PMCd("YN")
    '*************************** ZUBIN **************************
    ' 10th Dec 2003, EMR206 - For Priority Code Gen
'    Case Is = "PRTCD"
'      If GrdParam.Value(RowNum, "PDesc225") = "N" Then
'        Cancel = True: ErrMsg = "Cannot Edit Value": Exit Sub
'      End If
    '*************************** ZUBIN **************************
    ' **** Zubin 211 **** '
    Case Is = "BCHR"
'  ####################  Manoj # Ver: 2.0.11 Patch 01 # Date: 17/08/2005  ####################
      HlpList.MultiSelect = True
'  #########################################  Manoj  #########################################
      'IG BCHR-RmSCtg need not be validated as belonging to RmCtg
      Call HlpList.PSCd("RMSCTG", "", True)
    ' **** Zubin 211 **** '
    'JencClr Scrn
    Case Is = "USR"
      Call HlpList.CsCd
    'JencClr Scrn
    End Select
  Case Is = UCase("PValue3")
    Select Case UCase(ADC("wPTyp"))
    Case Is = "RMSCTG"
      Call HlpList.PMCd("YN")
    Case Is = "LABSCD"
      Call HlpList.PSCd("LABLOGIC", GrdParam.Value(RowNum, "PMCd"))
      
    '*************************** ZUBIN **************************
    ' 10th Dec 2003, EMR206 - For Priority Code Gen
'    Case Is = "PRTCD"
'      If GrdParam.Value(RowNum, "PDesc225") = "N" Then
'        Cancel = True: ErrMsg = "Cannot Edit Value": Exit Sub
'      End If
    '*************************** ZUBIN **************************
    
    '*** (Jen 3.01)
    Case Is = "USR"
      Call HlpList.PMCd("USRGRP")
    '*** (Jen 3.01)
    'IG BCHR-PRDTYP added
    Case Is = "BCHR"
      'IG BCHR- The Value in this field will be allowed only if
      'RMCtg field will have the value as ‘D’ Or ‘C’ or both.
      Cancel = Not (("," + GrdParam.Value(RowNum, "PValue1") + "," Like "*,C,*") Or ("," + GrdParam.Value(RowNum, "PValue1") + "," Like "*,D,*"))
      If Cancel = True Then ErrMsg = "Cannot Enter Value if D Or C not in Rm Category": Exit Sub
      Call HlpList.PMCd("PRDTYP")
            
    End Select
    
' ###################################  Manoj #### Ver: 2.0.6  ###################################
' Date:   20/01/2004
  Case Is = UCase("PValidYN")
    Select Case UCase(ADC("wPTyp"))
    ' Manoj 2.10.0
    'Case Is = "LABSCD"
    'Sachin 4.1.0.0 - New fields added (CmGrp, CmTyp)
    Case Is = "LABSCD", "DMCOL", "DMCTG", "DMSZ", "PRDCTG", "RMCOL", "SALCTG", "USR", "CMGRP", "CMTYP", "DTCYC", "CELL", "PRDLINE", "PDMGR", "MKTMGR" 'EmrSv - Uma
    ' Manoj 2.10.0
      Call HlpList.PMCd("YN")
    End Select
' ###########################################  Manoj  ###########################################
  End Select
End Sub

Private Sub GrdParam_ColChange(ByVal RowNum As Integer, ByVal ColName As String, ByVal NewValue As Variant, Cancel As Boolean, ErrMsg As String)
  '*** Check From TcStr If Blank Is Allowed ***
      'ws_BlnkYN = TsBlank from TcStr(TsTbl='Param' and TsTyp= PTyp and TsSTyp=''
      '                               and TsIdName= ColName)
  '*** PMCd
      'PTyp is 'CHR'
      '     If ws_BlnkYN= 'N' then valid PMCd('TC')
      '     If ws_BlnkYN= 'Y' then '' or valid PMCd('TC')
      'PTyp is 'DAANACD'
      '     If ws_BlnkYN= 'N' then valid PMCd('DAANAFLD')
      '     If ws_BlnkYN= 'Y' then '' or valid PMCd('DAANAFLD')
      'PTyp is 'DAANAFLD'
      '     Enter A Proper Integer > 0 & < 100
      'PTyp is 'LABSCD'
      '     If ws_BlnkYN= 'N' then valid PMCd('LABMCD')
      '     If ws_BlnkYN= 'Y' then '' or valid PMCd('LABMCD')
      'PTyp is 'RMSCTG', 'RMCOL'
      '     If ws_BlnkYN= 'N' then valid PMCd('RMCTG')
      '     If ws_BlnkYN= 'Y' then '' or valid PMCd('RMCTG')
      'PTyp is 'SUBPRC'
      '     If ws_BlnkYN= 'N' then valid LocCd from Loc(LocTyp= 'P')
      '     If ws_BlnkYN= 'Y' then '' or valid LocCd from Loc(LocTyp= 'P')
      'PTyp is 'RWRK'       bhavna
      '     If ws_BlnkYN= 'N' then valid LocCd from Loc(LocTyp= 'P')
      '     If ws_BlnkYN= 'Y' then '' or valid LocCd from Loc(LocTyp= 'P')
      'PTyp is 'PRTCD'
      '     If ws_BlnkYN= 'N' then valid two digit number
      '     If ws_BlnkYN= 'Y' then '' or valid two digit number
  '*** PSCd
      '     If ws_BlnkYN= 'N' then this field cannot be Blank
  '*** PDesc225
      'PTyp is 'CHR'
      '     If ws_BlnkYN= 'N' then valid LocCd from Loc(LocTyp<> 'P')
      '     If ws_BlnkYN= 'Y' then '' or valid LocCd from Loc(LocTyp<> 'P')
      'PTyp is 'GLDAS', 'IWTFRORD', 'CPYRT'
      '     valid sequence of PMCd('RMCTG') (separated by comma)
      'PTyp is 'PRDSEQ', 'WIWSEQ', 'WIWLONG'
      '     valid sequence of Process Locations i.e. valid LocCd from Loc(LocTyp= 'P')
      '     separated by commas
      'PTyp is 'RMSCTG'
      '     If ws_BlnkYN= 'N' then valid PMCd('YN')
      '     If ws_BlnkYN= 'Y' then '' or valid PMCd('YN')
      'PTyp is 'SRTBAG'
      '     If ws_BlnkYN= 'N' then valid sequence of fields from the Bag table
      '     If ws_BlnkYN= 'Y' then '' or valid sequence of fields from the Bag table
      'PTyp is 'SRTINBAG'
      '     If ws_BlnkYN= 'N' then valid sequence of fields from the InvFgd table
      '     If ws_BlnkYN= 'Y' then '' or valid sequence of fields from the InvFgd table
      'PTyp is 'SRTLR'
      '     If ws_BlnkYN= 'N' then valid sequence of fields from the LabRt table
      '     If ws_BlnkYN= 'Y' then '' or valid sequence of fields from the LabRt table
      'PTyp is 'SRTPP'
      '     If ws_BlnkYN= 'N' then valid sequence of fields from the PrdPts table
      '     If ws_BlnkYN= 'Y' then '' or valid sequence of fields from the PrdPts table
      'PTyp is 'SRTPRHLP'
      '     If ws_BlnkYN= 'N' then valid sequence of fields from the TxndZ table
      '     If ws_BlnkYN= 'Y' then '' or valid sequence of fields from the TxndZ table
      'PTyp is 'SRTRM'
      '     If ws_BlnkYN= 'N' then valid sequence of fields from the RmMst table
      '     If ws_BlnkYN= 'Y' then '' or valid sequence of fields from the RmMst table
      'PTyp is 'SRTRR'
      '     If ws_BlnkYN= 'N' then valid sequence of fields from the RmRt table
      '     If ws_BlnkYN= 'Y' then '' or valid sequence of fields from the RmRt table
      'PTyp is 'TC'
      '     If ws_BlnkYN= 'N' then valid PSCd('CHR', NewValue)
      '     If ws_BlnkYN= 'Y' then '' or valid PSCd('CHR', NewValue)
      'PTyp is 'INGR'
      '     If ws_BlnkYN= 'N' then valid PMCd('INGRPRCP', NewValue)
      '     If ws_BlnkYN= 'Y' then '' or valid PMCd('INGRPRCP', NewValue)
      'PTyp is 'USR'
      '     If ws_BlnkYN= 'N' then valid sequence of locations where LocTyp= 'P'
      '     If ws_BlnkYN= 'Y' then '' or valid sequence of locations where LocTyp= 'P'
  '*** PValue
      'PTyp is 'CHR'
      '     If ws_BlnkYN= 'N' then valid LocCd from Loc(LocTyp<> 'P')
      '     If ws_BlnkYN= 'Y' then '' or valid LocCd from Loc(LocTyp<> 'P')
      'PTyp is 'DAANAFLD'
      '     If ws_BlnkYN= 'N' then valid PMCd('YN')
      '     If ws_BlnkYN= 'Y' then '' or valid PMCd('YN')
      'PTyp is 'DMCTG'
      '     If ws_BlnkYN= 'N' then valid PMCd('DMUOM')
      '     If ws_BlnkYN= 'Y' then '' or valid PMCd('DMUOM')
      'PTyp is 'LABSCD'
      '     If ws_BlnkYN= 'N' then valid PMCd('LABQW')
      '     If ws_BlnkYN= 'Y' then '' or valid PMCd('LABQW')
      'PTyp is 'PARENTCD'
      '     If ws_BlnkYN= 'N' then valid PMCd('RMCTG')
      '     If ws_BlnkYN= 'Y' then '' or valid PMCd('RMCTG')
      'PTyp is 'PRDCTG'
      '     If ws_BlnkYN= 'N' then valid PMCd('PRDSEQ')
      '     If ws_BlnkYN= 'Y' then '' or valid PMCd('PRDSEQ')
      'PTyp is 'PRDSEQ'
      '     If ws_BlnkYN= 'N' then valid PMCd('PLNSEQ')
      '     If ws_BlnkYN= 'Y' then '' or valid PMCd('PLNSEQ')
      'PTyp is 'REMAKE'
      '     valid sequence of PMCd('RMCTG') (separated by comma)
      'PTyp is 'RMCTG'
      '     If ws_BlnkYN= 'N' then valid RmCd from RmMst(RmZ= 'N')
      '     If ws_BlnkYN= 'Y' then '' or valid RmCd from RmMst(RmZ= 'N')
      'PTyp is 'RMSCTG'
      '     If ws_BlnkYN= 'N' then valid PMCd('RATEBY')
      '     If ws_BlnkYN= 'Y' then '' or valid PMCd('RATEBY')
      'PTyp is 'TC'
      '     If ws_BlnkYN= 'N' then valid PMCd('TCTYP')
      '     If ws_BlnkYN= 'Y' then '' or valid PMCd('TCTYP')
      'PTyp is 'TCTYP'
      '     If ws_BlnkYN= 'N' then valid PMCd('TC')
      '     If ws_BlnkYN= 'Y' then '' or valid PMCd('TC')
      'PTyp is 'WORK'
      '     If ws_BlnkYN= 'N' then valid LocCd from Loc(LocTyp= 'P')
      '     If ws_BlnkYN= 'Y' then '' or valid LocCd from Loc(LocTyp= 'P')
  '*** PNum
      'PTyp is 'DMCTG'
      '     If ws_BlnkYN= 'N' then >0
      'PTyp is 'OUNCE'
      '     If ws_BlnkYN= 'N' then <> 0
      'PTyp is 'VACTG'
      '     If ws_BlnkYN= 'N' then <> 0
  '*** PValue1
      'PTyp is 'CHR'
      '     If ws_BlnkYN= 'N' then valid LocCd from Loc(LocTyp= 'P')
      '     If ws_BlnkYN= 'Y' then '' or valid LocCd from Loc(LocTyp= 'P')
      'PTyp is 'DAANAFLD'
      '     If ws_BlnkYN= 'N' then valid PMCd('YN')
      '     If ws_BlnkYN= 'Y' then '' or valid PMCd('YN')
      'PTyp is 'PARENTCD'
      '     If ws_BlnkYN= 'N' then valid PSCd('RMSCTG', PMCd)
      '     If ws_BlnkYN= 'Y' then '' or valid PMCd('RMSCTG', PMCd)
      'PTyp is 'RMSCTG'
      '     If ws_BlnkYN= 'N' then valid PMCd('YN')
      '     If ws_BlnkYN= 'Y' then '' or valid PMCd('YN')
      'PTyp is 'USR'
      '     If ws_BlnkYN= 'N' then valid sequence of locations where LocTyp<> 'P'
      '     If ws_BlnkYN= 'Y' then '' or valid sequence of locations where LocTyp<> 'P'
  '*** PValue2
      'PTyp is 'CHR'
      '     If ws_BlnkYN= 'N' then valid PMCd('DC')
      '     If ws_BlnkYN= 'Y' then '' or valid PMCd('DC')
      'PTyp is 'RMSCTG'
      '     If ws_BlnkYN= 'N' then valid PMCd('YN')
      '     If ws_BlnkYN= 'Y' then '' or valid PMCd('YN')
  '*** PValue3
      'PTyp is 'RMSCTG'
      '     If ws_BlnkYN= 'N' then valid PMCd('YN')
      '     If ws_BlnkYN= 'Y' then '' or valid PMCd('YN')

  Dim ws_BlnkYN As String, i As Integer, ws_Loc() As String, ws_Usr() As String
  Dim ws_RmCtgArr() As String, ws_PrcsSeqArr() As String, wi_WHUsrs As Integer
  
  With GrdParam
    ws_BlnkYN = moCn.GetFldVal("Select TsBlank from TcStr where TsTbl='Param' and " + _
                            "TsTyp='" + ADC("wPTyp") + "' and TsSTyp='' and " + _
                            "TsIdName='" + ColName + "'")
    Select Case ColName
    
    '****** Sachin 3.02.0 - Checking the For the Number Of Users to Full Db To be As per the value specified in parameter 'WHUSERS' ******
    Case Is = UCase("PNum1")
      Select Case UCase(ADC("wPTyp"))
      Case Is = "USR"
        If NewValue = 1 _
          And moCn.RecSeek("Select 'x' From Param Where PTyp='USR' And PMCd='" + GrdParam.Value(RowNum, "PMCd") + "' And PSCd='' And PNum1=2 ") _
          And moCn.RecSeek("Select 'x' From Usrlogin Where UlUsrCd='" + GrdParam.Value(RowNum, "PMCd") + "' ") Then
          Cancel = True
          If Cancel = True Then ErrMsg = "Cannot Change Partition Key to Current as User has logged on"
        End If
        
        wi_WHUsrs = moCn.GetFldVal("Select PNum From Param Where PTyp='WHUSERS' And PMCd='' and PSCd=''")
        If NewValue = 2 And moCn.GetFldVal("Select count('x') from Param where PTyp= 'USR' and PNum1=2 ") = wi_WHUsrs Then
          Cancel = True
          If Cancel = True Then ErrMsg = "Only " + CStr(wi_WHUsrs) + " Users Can Have Access To Complete Database"
        End If
       Case Is = "RMSCTG"
        If GrdParam.Value(RowNum, "pDESC225") = "Y" Then
          If NewValue < 1 Or NewValue > 3 Then
            Cancel = True
            ErrMsg = "NoOfSz can only be 1, 2 or 3"
          End If
          If Cancel = False Then
            If NewValue <> 1 And NewValue > moCn.GetFldVal("Select hRmStkBy from Head where HCoCd='" + gs_CoCd + "'") Then
              Cancel = True
              ErrMsg = "NoOfSz cannot be greater than what is in Head"
            End If
          End If
          
          Dim wRmSzCndToChk As String
          wRmSzCndToChk = "1 != 1"
          If Cancel = False Then
            If NewValue = 2 Then wRmSzCndToChk = "(TdRmSz3 <> 0)"
            If NewValue = 1 Then wRmSzCndToChk = "(TdRmSz2<> 0 or TdRmSz3 <> 0)"
            If moCn.RecSeek("Select 'a' from Txnd join RmMst on TdRmCd = RmCd where RmSCtg = '" + GrdParam.Value(RowNum, "Pscd") + "' and TdRmCd <>'' and " + _
                           wRmSzCndToChk) Then
              Cancel = True
              ErrMsg = "Transactions exist with non zero values for Sizes greater than given number"
            End If
          End If
        Else
          If NewValue <> 0 Then
            Cancel = True
            ErrMsg = "NoOfSz can only be 0"
          End If
        End If
      End Select
    '****** Sachin 3.02.0 - Checking the For the Number Of Users to Full Db To be As per the value specified in parameter 'WHUSERS' ******
    
    Case Is = UCase("PMCd")
    
      Select Case UCase(ADC("wPTyp"))

'      Case Is = "CHR" + gs_CoCd
'        If ws_BlnkYN = "N" And NewValue = "" Then _
'           Cancel = True: ErrMsg = "Tc Should Be Entered": Exit Sub
'        Cancel = (NewValue <> "") And (Not moCn.RecSeek("Select PMCd From Param where PTyp= 'TC' " + _
'                                      "and PMCd = '" + NewValue + "'"))
'        ErrMsg = "Invalid Tc": Exit Sub
      Case Is = "CNVCDCHK"
        Cancel = NewValue <> "CNVCDCHK"
        ErrMsg = "This value must be CnvCdChk"
      Case Is = "DAANACD"
        If ws_BlnkYN = "N" And NewValue = "" Then _
           Cancel = True: ErrMsg = "Analysis Field Should Be Entered": Exit Sub
        Cancel = (NewValue <> "") And (Not moCn.RecSeek("Select PMCd From Param where PTyp= 'DAANAFLD' " + _
                                      "and PMCd = '" + NewValue + "' and PValue= 'Y'"))
        ErrMsg = "Invalid Design Analysis Field": Exit Sub
      Case Is = "DAANAFLD"
        If ws_BlnkYN = "N" And NewValue = "" Then _
           Cancel = True: ErrMsg = "Analysis Field Should Be Entered": Exit Sub
        If Val(NewValue) < 1 Or Val(NewValue) > 100 Then Cancel = True: ErrMsg = "Analysis Field Value Should Be An Integer > 0 And < 100": Exit Sub
        For i = 1 To Len(NewValue)
          If Mid(NewValue, i, 1) <> "0" Then
            If Asc(Mid(NewValue, i, 1)) < 49 Or Asc(Mid(NewValue, i, 1)) > 57 Then Cancel = True: ErrMsg = "Analysis Field Value Should Be An Integer > 0 And < 100": Exit Sub
          End If
        Next
      Case Is = "LABSCD"
        If ws_BlnkYN = "N" And NewValue = "" Then _
        Cancel = True: ErrMsg = "Lab Main Code Should Be Entered": Exit Sub
        Cancel = (NewValue <> "") And (Not moCn.RecSeek("Select PMCd From Param where " + _
                                      "PTyp= 'LABMCD' and PMCd = '" + NewValue + "'"))
        ErrMsg = "Invalid Lab Main Code": Exit Sub
      Case Is = "LABLOGIC"
        Cancel = True: ErrMsg = "Lab Main Code Should Be Entered"
        Cancel = (NewValue <> "") And (Not moCn.RecSeek("Select PMCd From Param where " + _
                                      "PTyp= 'LABMCD' and PMCd = '" + NewValue + "'"))
        Exit Sub
      
      Case Is = "RMSCTG", "RMCOL"
        If ws_BlnkYN = "N" And NewValue = "" Then _
           Cancel = True: ErrMsg = "Raw Material Category Should Be Entered": Exit Sub
        Cancel = (NewValue <> "") And (Not moCn.RecSeek("Select PMCd From Param where " + _
                                      "PTyp= 'RMCTG' and PMCd = '" + NewValue + "'"))
        ErrMsg = "Invalid Raw Material Category": Exit Sub
      Case Is = "SUBPRC"
      'uma **** P replaced by R
        If ws_BlnkYN = "N" And NewValue = "" Then _
           Cancel = True: ErrMsg = "Process Should Be Entered": Exit Sub
        Cancel = (NewValue <> "") And (Not moCn.RecSeek("Select LocCd From Loc where " + _
                                      "LocTyp= 'R' and LocCd = '" + NewValue + "'"))
        ErrMsg = "Invalid Process ": Exit Sub
      '*********Bhavna
      Case Is = "RWRK"  'uma **** P replaced by R
        If ws_BlnkYN = "N" And NewValue = "" Then _
           Cancel = True: ErrMsg = "Process Location Should Be Entered": Exit Sub
        Cancel = (NewValue <> "") And (Not moCn.RecSeek("Select LocCd From Loc where " + _
                                      "LocTyp= 'R' and LocCd = '" + NewValue + "'"))
        ErrMsg = "Invalid Process Location": Exit Sub
      '**********Bhavna
      Case Is = "PRTCD"
        For i = 1 To Len(NewValue)
          If Asc(Mid(NewValue, i, 1)) < 48 Or Asc(Mid(NewValue, i, 1)) > 57 Then _
             Cancel = True: ErrMsg = "Priority Code Should Be A Number": Exit Sub
        Next i
        '******Geeta****CRM*****2.12************
      Case Is = "RMID"
         If ws_BlnkYN = "N" And NewValue = "" Then _
         Cancel = True: ErrMsg = "Grade RmCtg Should Be Entered": Exit Sub
         Cancel = (NewValue <> "") And (Not moCn.RecSeek("Select PMCd From Param where " + _
                                      "PTyp= 'GRDRMCTG' and PSCd = '" + NewValue + "'"))
        ErrMsg = "Invalid Grade RmCtg": Exit Sub
      Case Is = "GRDCD"
         If ws_BlnkYN = "N" And NewValue = "" Then _
         Cancel = True: ErrMsg = "Grade Type Should Be Entered": Exit Sub
         Cancel = (NewValue <> "") And (Not moCn.RecSeek("Select PMCd From Param where " + _
                                      "PTyp= 'GRDTYP' and PMCd = '" + NewValue + "'"))
        ErrMsg = "Invalid Grade Type": Exit Sub
      ' Zubin 212 (07/12/05)
      Case Is = "SALCTG"
         If ws_BlnkYN = "N" And NewValue = "" Then Cancel = True: ErrMsg = "Design Category Should Be Entered": Exit Sub
         Cancel = (NewValue <> "") And (Not moCn.RecSeek("Select 'x' From Param where " + _
                                      "PTyp= 'DMCTG' and PMCd = '" + NewValue + "'"))
        ErrMsg = "Invalid Design Category": Exit Sub
      ' Zubin 212 (07/12/05)
      ' Zubin 212 (24/02/06)
      '*** Jen 2.14 Next Patch 1 [No of Usrs = 4]
      Case Is = "USR"
        If GrdParam.IsNew(RowNum) = True Then
          ' ***** Manali 3.7.1 - JMIC included
          '4.1.4
          Cancel = moCn.GetFldVal("Select count('x') from Param where PTyp= 'USR' and PValidYN= 'Y'") + 1 > MWLib.NoOfUsers + 1
          ErrMsg = "The Total No. Of Users Cannot Exceed " + CStr(MWLib.NoOfUsers + 1) + "."
          
          'before 4.1.4
          'Cancel = moCn.GetFldVal("Select count('x') from Param where PTyp= 'USR' and PValidYN= 'Y'") + 1 > MWLib.NoOfUsers + 5
          'ErrMsg = "The Total No. Of Users Cannot Exceed " + CStr(MWLib.NoOfUsers + 5) + "."
          
          ' Bef 3.7.1
          'Cancel = moCn.GetFldVal("Select count('x') from Param where PTyp= 'USR' and PValidYN= 'Y'") + 1 > MWLib.NoOfUsers + 4
          'ErrMsg = "The Total No. Of Users Cannot Exceed " + CStr(MWLib.NoOfUsers + 4) + "."
        End If
      '*** Jen 2.14 Next Patch 1 [No of Usrs = 4]
      ' Zubin 212 (24/02/06)
      ' Zubin 213 (was missing Before 2.13)
      Case Is = "GMFACTOR"
         If ws_BlnkYN = "N" And NewValue = "" Then Cancel = True: ErrMsg = "Raw Material Category Should Be Entered": Exit Sub
         Cancel = (NewValue <> "") And (Not moCn.RecSeek("Select 'x' From Param where " + _
                                      "PTyp= 'RMCTG' and PMCd = '" + NewValue + "' and PSCd= ''"))
        ErrMsg = "Invalid Raw Material Category": Exit Sub
      ' Zubin 213 (was missing Before 2.13)
      '*** JAY 2.14 *** (Yy Change)
      Case Is = "YY"
        If Not IsNumeric(NewValue) Then Cancel = True: ErrMsg = "Yy Should Be A Number": Exit Sub
      '*** JAY 2.14 *** (Yy Change)
      ' ***** Manali 3.5.0 - 12/11/08 - Color applied to WIW Reports
      Case Is = "WIWCOL"
         If ws_BlnkYN = "N" And NewValue = "" Then _
         Cancel = True: ErrMsg = "Report Color Should Be Entered": Exit Sub
         Cancel = (NewValue <> "") And (Not moCn.RecSeek("Select PMCd From Param where " + _
                                      "PTyp= 'REPCOL' and PMCd = '" + NewValue + "'"))
        ErrMsg = "Invalid Color": Exit Sub
      Case Is = "SKL"
        If Not IsNumeric(NewValue) Then
          Cancel = True: ErrMsg = "Skill must be a number": Exit Sub
        End If
        If CInt(NewValue) <= 0 Then
          Cancel = True: ErrMsg = "Skill must be a positive number": Exit Sub
        End If
        
      ' ***** Manali 3.5.0 - 12/11/08 - Color applied to WIW Reports
      'uma *****
      'Case Is = "CSEQ"
      ' If ws_BlnkYN = "N" And NewValue = "" Then _
      '     Cancel = True: ErrMsg = "Cell Name Should Be Entered": Exit Sub
      '  Cancel = (NewValue <> "") And (Not moCn.RecSeek("Select PMCd From Param where PTyp= 'CELL' " + _
      '                                "and PMCd = '" + NewValue + "'"))
      '  ErrMsg = "Invalid Cell Name ": Exit Sub
      'pg.11-Checking valid GstCode
      Case Is = "GSTFLDS"
         If ws_BlnkYN = "N" And NewValue = "" Then _
         Cancel = True: ErrMsg = "GST Code Should Be Entered": Exit Sub
         Cancel = (NewValue <> "") And (Not moCn.RecSeek("Select PMCd From Param where " + _
                                      "PTyp= 'GSTCD' and PMCd = '" + NewValue + "'"))
         ErrMsg = "Invalid GST Code": Exit Sub
      End Select
      If ws_BlnkYN = "N" And NewValue = "" Then _
         Cancel = True: ErrMsg = "Value Should Be Entered": Exit Sub
    Case Is = UCase("PSCd")
      ' Zubin 213
      Select Case UCase(ADC("wPTyp"))
      Case Is = "GMFACTOR"
        Cancel = (NewValue <> "") And Not moCn.RecSeek("Select 'x' from Param where PTyp= 'CURNCY' and PMCd= '" + NewValue + "' and PSCd= ''")
        ErrMsg = "Invalid Currency Code": Exit Sub
      'uma ***
      'Case Is = "CSEQ"
      '  If ws_BlnkYN = "N" And NewValue = "" Then _
      '     Cancel = True: ErrMsg = "Process Sequence Should Be Entered": Exit Sub
      '  Cancel = (NewValue <> "") And Not moCn.RecSeek("Select vPMCd from vParam where " + _
       '         "vPTyp= 'PRCSEQ' and vPMCd= '" + NewValue + "'")
        'ErrMsg = "Invalid Process Sequence Code": Exit Sub
      'pg.11-Checking valid Gst serial
      Case Is = "GSTFLDS"
        Cancel = Not moCn.RecSeek("Select PMCd From Param where " + _
                                      "PTyp= 'GSTSR' and PMCd = " + CStr(NewValue) + "")
        ErrMsg = "Invalid GST Serial Number": Exit Sub

      End Select
      ' Zubin 213
      If ws_BlnkYN = "N" And NewValue = "" Then _
         Cancel = True: ErrMsg = "Value Should Be Entered": Exit Sub
    Case Is = UCase("PDesc225")
      Select Case UCase(ADC("wPTyp"))
'      Case Is = "CHR" + gs_CoCd
'        If ws_BlnkYN = "N" And NewValue = "" Then _
'           Cancel = True: ErrMsg = "To RM Location Should Be Entered": Exit Sub
'        Cancel = (NewValue <> "") And (Not moCn.RecSeek("Select LocCd From Loc where " + _
'                                      "LocTyp<> 'P' and LocCd = '" + NewValue + "'"))
'        ErrMsg = "Invalid To RM Location": Exit Sub
 Case Is = "BCHR"
        Cancel = (NewValue <> "") And Not (moCn.RecSeek("Select PMCd from Param where PTyp='SUBSHP' AND PMCd='" + NewValue + "'"))
        ErrMsg = "SubShp must be valid Or Blank": Exit Sub
      Case Is = "DSGCPY"
        Call HlpList.PMCd("TCTYP", , "OM")
        Cancel = (NewValue <> "") And (Not moCn.RecSeek("Select PMCd From Param where PTyp= 'TCTYP' " + _
                 "and PMCd = '" + NewValue + "' and PValue1= 'OM'"))
        If Cancel = True Then ErrMsg = "Invalid 'OM' TcType": Exit Sub
      Case Is = "GLDAS", "IWTFRORD", "CPYRT"
        '*** To Check If All The RmCtgs In The Seq Are Valid
            'Use the Split function to break up the sequence and store the individual values
            'in an array
            'Use the For-Next Loop to validate the individual values in the array
        ws_RmCtgArr = Split(NewValue, ",")
        If IsArray(ws_RmCtgArr) Then
          For i = 0 To UBound(ws_RmCtgArr)
            Cancel = Not moCn.RecSeek("Select PMCd From Param where PTyp= 'RMCTG' " + _
                     "and PMCd = '" + ws_RmCtgArr(i) + "'")
            If Cancel = True Then ErrMsg = "Value(" + CStr(i + 1) + ") Is An Invalid Raw Material Category": Exit Sub
          Next i
        End If
'      Case Is = "PRDSEQ", "WIWSEQ", "WIWLONG"
'        Dim ws_LocArr() As String
'        ws_LocArr = Split(NewValue, ",")
'        If IsArray(ws_LocArr) Then
'          For i = 0 To UBound(ws_LocArr)
'            Cancel = Not moCn.RecSeek("Select LocCd From Loc where LocTyp= 'P' " + _
'                                      "and LocCd = '" + ws_LocArr(i) + "'")
'            If Cancel = True Then ErrMsg = "Loc(" + CStr(i + 1) + ") Is An Invalid Production Location": Exit Sub
'          Next i
'        End If
      Case Is = "RMSCTG"
        If ws_BlnkYN = "N" And NewValue = "" Then _
           Cancel = True: ErrMsg = "Yes/No Should Be Entered": Exit Sub
        Cancel = (NewValue <> "") And (Not moCn.RecSeek("Select PMCd From Param where " + _
                                      "PTyp= 'YN' and PMCd = '" + NewValue + "'"))
        ErrMsg = "Accept Size In Transaction (Y/N)": Exit Sub
      
      'EmrSv - Uma
      Case Is = "DMCTG"
        If ws_BlnkYN = "N" And NewValue = "" Then _
           Cancel = True: ErrMsg = "Process Should Be Entered": Exit Sub
        ws_PrcsSeqArr = Split(NewValue, ",")
        If UBound(ws_PrcsSeqArr) > 0 Then
           Cancel = True: ErrMsg = "Only one Process Sequence should Be Entered": Exit Sub
        End If
        
        For i = 0 To UBound(ws_PrcsSeqArr)
          If IsArray(ws_PrcsSeqArr) Then
          Cancel = (NewValue <> "") And (Not moCn.RecSeek("Select vPMCd From vParam where " + _
                                      "vPTyp= 'PRCSSEQ' and vPMCd = '" + ws_PrcsSeqArr(i) + "'"))
            If Cancel = True Then ErrMsg = "Value(" + CStr(i + 1) + ") Is An Invalid Process Sequence": Exit Sub
          End If
        Next i
'        ErrMsg = "Invalid Process Sequence": Exit Sub
            
      Case Is = "SRTBAG"
        '*** To Check If All The Fields In The Sort Seq Are Valid Fields In Bag Table
        '*** INFORMATION_SCHEMA.COLUMNS is a view which contains one row for each column
            'accessible to the current user in the current database.
        '*** OBJECT_ID is a system function which returns the database
            'object identification number. (You can also append ' asc' or ' desc' after the fields)
        Dim ws_BagFld() As String
        If ws_BlnkYN = "N" And NewValue = "" Then _
           Cancel = True: ErrMsg = "Sort Field/s Should Be Entered": Exit Sub
        ws_BagFld = Split(NewValue, ",")
        If IsArray(ws_BagFld) Then
          For i = 0 To UBound(ws_BagFld)
            Cancel = Not moCn.RecSeek("Select 'x' From INFORMATION_SCHEMA.Columns " + _
                     "where TABLE_NAME = 'Bag' and (" + _
                     "COL_NAME((Select OBJECT_ID('Bag')), ORDINAL_POSITION)= '" + ws_BagFld(i) + "'" + _
                     "or COL_NAME((Select OBJECT_ID('Bag')), ORDINAL_POSITION)+ ' desc'= '" + ws_BagFld(i) + "'" + _
                     "or COL_NAME((Select OBJECT_ID('Bag')), ORDINAL_POSITION)+ ' asc'= '" + ws_BagFld(i) + "')")
            If Cancel = True Then ErrMsg = "Field(" + CStr(i + 1) + ") Is Not A Field In Bag Table": Exit Sub
          Next i
        End If
      Case Is = "SRTDLD"
        Dim ws_DldFld() As String
        If ws_BlnkYN = "N" And NewValue = "" Then _
           Cancel = True: ErrMsg = "Sort Field/s Should Be Entered": Exit Sub
        ws_DldFld = Split(NewValue, ",")
        If IsArray(ws_DldFld) Then
          For i = 0 To UBound(ws_DldFld)
            Cancel = Not moCn.RecSeek("Select 'x' From INFORMATION_SCHEMA.Columns " + _
                     "where TABLE_NAME = 'DsgLoad' and (" + _
                     "COL_NAME((Select OBJECT_ID('DsgLoad')), ORDINAL_POSITION)= '" + ws_DldFld(i) + "'" + _
                     "or COL_NAME((Select OBJECT_ID('DsgLoad')), ORDINAL_POSITION)+ ' desc'= '" + ws_DldFld(i) + "'" + _
                     "or COL_NAME((Select OBJECT_ID('DsgLoad')), ORDINAL_POSITION)+ ' asc'= '" + ws_DldFld(i) + "')")
            If Cancel = True Then ErrMsg = "Field(" + CStr(i + 1) + ") Is Not A Field In DsgLoad Table": Exit Sub
          Next i
        End If
      Case Is = "SRTINBAG"
        Dim ws_InvFgdFld() As String
        If ws_BlnkYN = "N" And NewValue = "" Then _
           Cancel = True: ErrMsg = "Sort Field/s Should Be Entered": Exit Sub
        ws_InvFgdFld = Split(NewValue, ",")
        If IsArray(ws_InvFgdFld) Then
          For i = 0 To UBound(ws_InvFgdFld)
            Cancel = Not moCn.RecSeek("Select 'x' From INFORMATION_SCHEMA.Columns " + _
                     "where TABLE_NAME = 'InvFgd' and (" + _
                     "COL_NAME((Select OBJECT_ID('InvFgd')), ORDINAL_POSITION)= '" + ws_InvFgdFld(i) + "'" + _
                     "or COL_NAME((Select OBJECT_ID('InvFgd')), ORDINAL_POSITION)+ ' desc'= '" + ws_InvFgdFld(i) + "'" + _
                     "or COL_NAME((Select OBJECT_ID('InvFgd')), ORDINAL_POSITION)+ ' asc'= '" + ws_InvFgdFld(i) + "')")
            If Cancel = True Then ErrMsg = "Field(" + CStr(i + 1) + ") Is Not A Field In Invoice Bag Table": Exit Sub
          Next i
        End If
      Case Is = "SRTLR"
        Dim ws_LabRtFld() As String
        If ws_BlnkYN = "N" And NewValue = "" Then _
           Cancel = True: ErrMsg = "Sort Field/s Should Be Entered": Exit Sub
        ws_LabRtFld = Split(NewValue, ",")
        If IsArray(ws_LabRtFld) Then
          For i = 0 To UBound(ws_LabRtFld)
            Cancel = Not moCn.RecSeek("Select 'x' From INFORMATION_SCHEMA.Columns " + _
                     "where TABLE_NAME = 'LabRt' and (" + _
                     "COL_NAME((Select OBJECT_ID('LabRt')), ORDINAL_POSITION)= '" + ws_LabRtFld(i) + "'" + _
                     "or COL_NAME((Select OBJECT_ID('LabRt')), ORDINAL_POSITION)+ ' desc'= '" + ws_LabRtFld(i) + "'" + _
                     "or COL_NAME((Select OBJECT_ID('LabRt')), ORDINAL_POSITION)+ ' asc'= '" + ws_LabRtFld(i) + "')")
            If Cancel = True Then ErrMsg = "Field(" + CStr(i + 1) + ") Is Not A Field In LabRt Table": Exit Sub
          Next i
        End If
      Case Is = "SRTPP"
        Dim ws_PpFld() As String
        If ws_BlnkYN = "N" And NewValue = "" Then _
           Cancel = True: ErrMsg = "Sort Field/s Should Be Entered": Exit Sub
        ws_PpFld = Split(NewValue, ",")
        If IsArray(ws_PpFld) Then
          For i = 0 To UBound(ws_PpFld)
            Cancel = Not moCn.RecSeek("Select 'x' From INFORMATION_SCHEMA.Columns " + _
                     "where TABLE_NAME = 'PrdPts' and (" + _
                     "COL_NAME((Select OBJECT_ID('PrdPts')), ORDINAL_POSITION)= '" + ws_PpFld(i) + "'" + _
                     "or COL_NAME((Select OBJECT_ID('PrdPts')), ORDINAL_POSITION)+ ' desc'= '" + ws_PpFld(i) + "'" + _
                     "or COL_NAME((Select OBJECT_ID('PrdPts')), ORDINAL_POSITION)+ ' asc'= '" + ws_PpFld(i) + "')")
            If Cancel = True Then ErrMsg = "Field(" + CStr(i + 1) + ") Is Not A Field In PrdPts Table": Exit Sub
          Next i
        End If
      Case Is = "SRTPRHLP"
        Dim ws_TzFld() As String
        If ws_BlnkYN = "N" And NewValue = "" Then _
           Cancel = True: ErrMsg = "Sort Field/s Should Be Entered": Exit Sub
        ws_TzFld = Split(NewValue, ",")
        If IsArray(ws_TzFld) Then
          For i = 0 To UBound(ws_TzFld)
            Cancel = Not moCn.RecSeek("Select 'x' From INFORMATION_SCHEMA.Columns " + _
                     "where TABLE_NAME = 'TxndZ' and (" + _
                     "COL_NAME((Select OBJECT_ID('TxndZ')), ORDINAL_POSITION)= '" + ws_TzFld(i) + "'" + _
                     "or COL_NAME((Select OBJECT_ID('TxndZ')), ORDINAL_POSITION)+ ' desc'= '" + ws_TzFld(i) + "'" + _
                     "or COL_NAME((Select OBJECT_ID('TxndZ')), ORDINAL_POSITION)+ ' asc'= '" + ws_TzFld(i) + "')")
            If Cancel = True Then ErrMsg = "Field(" + CStr(i + 1) + ") Is Not A Field In TxndZ Table": Exit Sub
          Next i
        End If
      Case Is = "SRTRM"
        Dim ws_RmFld() As String
        If ws_BlnkYN = "N" And NewValue = "" Then _
           Cancel = True: ErrMsg = "Sort Field/s Should Be Entered": Exit Sub
        ws_RmFld = Split(NewValue, ",")
        If IsArray(ws_RmFld) Then
          For i = 0 To UBound(ws_RmFld)
            Cancel = Not moCn.RecSeek("Select 'x' From INFORMATION_SCHEMA.Columns " + _
                     "where TABLE_NAME = 'RmMst' and (" + _
                     "COL_NAME((Select OBJECT_ID('RmMst')), ORDINAL_POSITION)= '" + ws_RmFld(i) + "'" + _
                     "or COL_NAME((Select OBJECT_ID('RmMst')), ORDINAL_POSITION)+ ' desc'= '" + ws_RmFld(i) + "'" + _
                     "or COL_NAME((Select OBJECT_ID('RmMst')), ORDINAL_POSITION)+ ' asc'= '" + ws_RmFld(i) + "')")
            If Cancel = True Then ErrMsg = "Field(" + CStr(i + 1) + ") Is Not A Field In RmMst Table": Exit Sub
          Next i
        End If
      Case Is = "SRTRR"
        Dim ws_RmRtFld() As String
        If ws_BlnkYN = "N" And NewValue = "" Then _
           Cancel = True: ErrMsg = "Sort Field/s Should Be Entered": Exit Sub
        ws_RmRtFld = Split(NewValue, ",")
        If IsArray(ws_RmRtFld) Then
          For i = 0 To UBound(ws_RmRtFld)
            Cancel = Not moCn.RecSeek("Select 'x' From INFORMATION_SCHEMA.Columns " + _
                     "where TABLE_NAME = 'RmRt' and (" + _
                     "COL_NAME((Select OBJECT_ID('RmRt')), ORDINAL_POSITION)= '" + ws_RmRtFld(i) + "'" + _
                     "or COL_NAME((Select OBJECT_ID('RmRt')), ORDINAL_POSITION)+ ' desc'= '" + ws_RmRtFld(i) + "'" + _
                     "or COL_NAME((Select OBJECT_ID('RmRt')), ORDINAL_POSITION)+ ' asc'= '" + ws_RmRtFld(i) + "')")
            If Cancel = True Then ErrMsg = "Field(" + CStr(i + 1) + ") Is Not A Field In RmRt Table": Exit Sub
          Next i
          End If
        
      '*************************** ZUBIN **************************
      ' 01st Dec 2003, EMR206
      Case Is = "SRTRH"
        Dim ws_RmRtHistFld() As String
        If ws_BlnkYN = "N" And NewValue = "" Then _
           Cancel = True: ErrMsg = "Sort Field/s Should Be Entered": Exit Sub
        ws_RmRtHistFld = Split(NewValue, ",")
        If IsArray(ws_RmRtHistFld) Then
          For i = 0 To UBound(ws_RmRtHistFld)
            Cancel = Not moCn.RecSeek("Select 'x' From INFORMATION_SCHEMA.Columns " + _
                     "where TABLE_NAME = 'RmRtHist' and (" + _
                     "COL_NAME((Select OBJECT_ID('RmRtHist')), ORDINAL_POSITION)= '" + ws_RmRtHistFld(i) + "'" + _
                     "or COL_NAME((Select OBJECT_ID('RmRtHist')), ORDINAL_POSITION)+ ' desc'= '" + ws_RmRtHistFld(i) + "'" + _
                     "or COL_NAME((Select OBJECT_ID('RmRtHist')), ORDINAL_POSITION)+ ' asc'= '" + ws_RmRtHistFld(i) + "')")
            If Cancel = True Then ErrMsg = "Field(" + CStr(i + 1) + ") Is Not A Field In RmRtHist Table": Exit Sub
          Next i
        End If
      '*************************** ZUBIN **************************
        
'      Case Is = "TC"
'        If ws_BlnkYN = "N" And NewValue = "" Then _
'           Cancel = True: ErrMsg = "Default Character Should Be Entered": Exit Sub
'        Cancel = (NewValue <> "") And (Not moCn.RecSeek("Select PSCd From Param where PTyp= 'CHR' and " + _
'                                      "PMCd= '" + GrdParam.Value(RowNum, "PMCd") + "' and " + _
'                                      "PSCd = '" + NewValue + "'"))
'        ErrMsg = "Invalid Default Character": Exit Sub

      Case Is = "RETCHR"                    '4.1.4
        ws_Usr = Split(NewValue, ",")
        If IsArray(ws_Usr) Then
          For i = 0 To UBound(ws_Usr)
            Cancel = Not moCn.RecSeek(" Select PMCd From Param Where " + _
                                      " PTyp='USR' And PMCd = '" + ws_Usr(i) + "' And PSCd=''")
            If Cancel = True Then ErrMsg = "Value(" + CStr(i + 1) + ") Is An Invalid User": Exit Sub
          Next i
        End If

      Case Is = "USR"
        '*** To Check If All The Production Locs In The Seq Are Valid
        ws_Loc = Split(NewValue, ",")
        If IsArray(ws_Loc) Then
          For i = 0 To UBound(ws_Loc)
            Cancel = Not moCn.RecSeek("Select LocCd From Loc where " + _
                                      "LocTyp in ('P') and LocCd = '" + ws_Loc(i) + "'")
            If Cancel = True Then ErrMsg = "Value(" + CStr(i + 1) + ") Is An Invalid Prd Location": Exit Sub
          Next i
        End If
      Case Is = "INGR"
        '*** All the Values should be a valid group participant
        Dim ws_GrpArr() As String
        ws_GrpArr = Split(NewValue, ",")
        If IsArray(ws_GrpArr) Then
          For i = 0 To UBound(ws_GrpArr)
            Cancel = Not moCn.RecSeek("Select PMCd From Param where PTyp= 'INGRPRCP' " + _
                     "and PMCd = '" + ws_GrpArr(i) + "'")
            If Cancel = True Then ErrMsg = "Value(" + CStr(i + 1) + ") Is An Invalid Invoice Report Group Participant": Exit Sub
          Next i
        End If
   '*************************** ZUBIN **************************
      ' 10th Dec 2003, EMR206 - For Priority Reallotment
      Case Is = "PRTCD"
        Cancel = NewValue = vbNullString Or (Not moCn.RecSeek("Select PMCd From Param " + _
                                      "where PTyp= 'YN' and PMCd = '" + NewValue + "'"))
        If Cancel Then
          ErrMsg = "Specify Whether Priority Can Be ReAlloted"
          Exit Sub
'        Else
'          If UCase(NewValue) = "N" Then
'            GrdParam.Value(RowNum, "PValue") = 0
'            GrdParam.Value(RowNum, "PValue1") = 0
'            GrdParam.Value(RowNum, "PValue2") = 0
'            GrdParam.Value(RowNum, "PValue3") = 0
'          End If
        End If
      '*************************** ZUBIN **************************
      'JewelSmith.2 RmSctg validation for Ptyp=STNCLS is removed
            
      ' ***** Manali 3.8.0 - Metler Added
      Case Is = "METLR"
          Dim ws_UsrArr() As String
          ws_UsrArr = Split(NewValue, ",")
          If IsArray(ws_UsrArr) Then
            For i = 0 To UBound(ws_UsrArr)
              Cancel = Not moCn.RecSeek("Select PMCd From Param where PTyp= 'USR' and " + _
                                        "PMCd = '" + ws_UsrArr(i) + "'")
              If Cancel = True Then ErrMsg = "User(" + CStr(i + 1) + ") Is Not Defined": Exit Sub
            Next i
          End If
      
      End Select
      
      
      If ws_BlnkYN = "N" And NewValue = "" Then _
         Cancel = True: ErrMsg = "Value Should Be Entered": Exit Sub
    Case Is = UCase("PValue")
      Select Case UCase(ADC("wPTyp"))
      
'      Case Is = "CHR" + gs_CoCd
'        If ws_BlnkYN = "N" And NewValue = "" Then _
'           Cancel = True: ErrMsg = "From RM Location Should Be Entered": Exit Sub
'        Cancel = (NewValue <> "") And (Not moCn.RecSeek("Select LocCd From Loc where " + _
'                                      "LocTyp<> 'P' and LocCd = '" + NewValue + "'"))
'        ErrMsg = "Invalid From RM Location": Exit Sub
      'uma ***
      'Case Is = "CSEQ"
      '  If ws_BlnkYN = "N" And NewValue = "" Then _
      '     Cancel = True: ErrMsg = "Production Sequence Should Be Entered": Exit Sub
      '  Cancel = (NewValue <> "") And Not moCn.RecSeek("Select vPMCd from vParam where " + _
       '         "vPTyp= 'PRDSEQ' and vPMCd= '" + NewValue + "'")
        'ErrMsg = "Invalid Production Sequence Code": Exit Sub
        
      Case Is = "DAANAFLD"
        If ws_BlnkYN = "N" And NewValue = "" Then _
           Cancel = True: ErrMsg = "Specify Whether The Analysis Field Will Contain The Code Or Description": Exit Sub
        Cancel = (NewValue <> "") And (Not moCn.RecSeek("Select PMCd From Param where " + _
                                      "PTyp= 'YN' and PMCd = '" + NewValue + "'"))
        ErrMsg = "Specify Whether The Analysis Field Will Contain The Code Or Description": Exit Sub
      
      'Jen Clr Scr
      Case Is = "DEFCLR"
        If ws_BlnkYN = "N" And NewValue = "" Then _
           Cancel = True: ErrMsg = "Default Color Code Should Be Entered": Exit Sub
        Cancel = (NewValue <> "") And (Not moCn.RecSeek("Select 'x' From ClrSet where CsCd= '" + NewValue + "' "))
        ErrMsg = "Invalid Color Code": Exit Sub
      
      'Jen Clr Scr
      Case Is = "DMCTG"
        If ws_BlnkYN = "N" And NewValue = "" Then _
           Cancel = True: ErrMsg = "Design UOM Should Be Entered": Exit Sub
        Cancel = (NewValue <> "") And (Not moCn.RecSeek("Select PMCd From Param where " + _
                                      "PTyp= 'DMUOM' and PMCd = '" + NewValue + "'"))
        ErrMsg = "Invalid Design Unit Of Measurement": Exit Sub
        
      Case Is = "LABSCD"
        If ws_BlnkYN = "N" And NewValue = "" Then _
           Cancel = True: ErrMsg = "Labour Q/W Should Be Entered": Exit Sub
        Cancel = (NewValue <> "") And (Not moCn.RecSeek("Select PMCd From Param where " + _
                                      "PTyp= 'LABQW' and PMCd = '" + NewValue + "'"))
        ErrMsg = "Invalid Labour Q/W": Exit Sub
        
      Case Is = "PARENTCD"
        If ws_BlnkYN = "N" And NewValue = "" Then _
           Cancel = True: ErrMsg = "Rm Ctg Should Be Entered": Exit Sub
        Cancel = (NewValue <> "") And (Not moCn.RecSeek("Select PMCd From Param where " + _
                                      "PTyp= 'RMCTG' and PMCd = '" + NewValue + "'"))
        ErrMsg = "Invalid Rm Ctg": Exit Sub
      'EmrSv - Uma
      Case Is = "PDMGR"
        If ws_BlnkYN = "N" And NewValue = "" Then _
           Cancel = True: ErrMsg = "Cell Should Be Entered": Exit Sub
        Cancel = (NewValue <> "") And (Not moCn.RecSeek("Select PMCd From Param where " + _
                                      "PTyp= 'Cell' and PMCd = '" + NewValue + "'"))
        ErrMsg = "Invalid Cell": Exit Sub
        
'      Case Is = "PRDCTG"
'        If ws_BlnkYN = "N" And NewValue = "" Then _
'           Cancel = True: ErrMsg = "Production Sequence Should Be Entered": Exit Sub
'        Cancel = (NewValue <> "") And (Not moCn.RecSeek("Select PSCd From Param where " + _
'                                      "PTyp= 'PRDSEQ' and PMCd = '" + gs_CoCd + "' and " + _
'                                      "PSCd = '" + NewValue + "'"))
'        ErrMsg = "Invalid Production Sequence": Exit Sub
''      Case Is = "PRDSEQ"
'        '$$$
'        'If ws_BlnkYN = "N" And NewValue = "" Then _
'        '   Cancel = True: ErrMsg = "Plan Sequence Should Be Entered": Exit Sub
'        'Cancel = (NewValue <> "") And (Not moCn.RecSeek("Select PMCd From Param where " + _
'        '                              "PTyp= 'PLNSEQ' and PMCd = '" + NewValue + "'"))
'        'ErrMsg = "Invalid Plan Sequence": Exit Sub
'        '$$$
      Case Is = "PTYP"
        If ws_BlnkYN = "N" And NewValue = "" Then _
           Cancel = True: ErrMsg = "Specify Whether The Param Type Is Microway/System/Other": Exit Sub
        Cancel = (NewValue <> "") And (Not moCn.RecSeek("Select PMCd From Param where " + _
                                      "PTyp= 'SYSDFN' and PMCd = '" + NewValue + "'"))
        ErrMsg = "Invalid Patameter Category": Exit Sub
      
      '*** (Jen 3.01.1)
      Case Is = "REMAKE"
        '*** To Check If All The RmCtgs In The Seq Are Valid
            'Use the Split function to break up the sequence and store the individual values
            'in an array
            'Use the For-Next Loop to validate the individual values in the array
        ws_RmCtgArr = Split(NewValue, ",")
        If IsArray(ws_RmCtgArr) Then
          For i = 0 To UBound(ws_RmCtgArr)
            Cancel = Not moCn.RecSeek("Select PMCd From Param where PTyp= 'RMCTG' " + _
                     "and PMCd = '" + ws_RmCtgArr(i) + "'")
            If Cancel = True Then ErrMsg = "Value(" + CStr(i + 1) + ") Is An Invalid Raw Material Category": Exit Sub
          Next i
        End If
      '*** (Jen 3.01.1)
      
      Case Is = "RMCTG"
        If ws_BlnkYN = "N" And NewValue = "" Then _
           Cancel = True: ErrMsg = "Rm Code For Calculation Of Pure Wt Should Be Entered": Exit Sub
        Cancel = (NewValue <> "") And (Not moCn.RecSeek("Select RmCd From RmMst where " + IIF(gs_Partition = ctCurrPrtn, " RmPrtKey='" + ctCurrPrtn + "' And ", "") + " RmZ= 'N' and " + _
                                      "RmCtg= '" + GrdParam.Value(RowNum, "PMCd") + "' and " + _
                                      "RmCd = '" + NewValue + "'"))
        If Cancel = True Then ErrMsg = "Enter Correct Rm Code For Calculation Of Pure Wt": Exit Sub
      
      Case Is = "RMSCTG"
        If ws_BlnkYN = "N" And NewValue = "" Then _
           Cancel = True: ErrMsg = "Rate By Option Should Be Entered": Exit Sub
        Cancel = (NewValue <> "") And (Not moCn.RecSeek("Select PMCd From Param where " + _
                                      "PTyp= 'RATEBY' and PMCd = '" + NewValue + "'"))
        ErrMsg = "Invalid Rate By Option": Exit Sub
      
      Case Is = "TC"
        If ws_BlnkYN = "N" And NewValue = "" Then _
           Cancel = True: ErrMsg = "Tc Type Should Be Entered": Exit Sub
        Cancel = (NewValue <> "") And (Not moCn.RecSeek("Select PMCd From Param where " + _
                                      "PTyp= 'TCTYP' and PMCd = '" + NewValue + "'"))
        ErrMsg = "Invalid Tc Type": Exit Sub
      
      Case Is = "TCTYP"
        If ws_BlnkYN = "N" And NewValue = "" Then _
           Cancel = True: ErrMsg = "Default Tc Should Be Entered": Exit Sub
        Cancel = (NewValue <> "") And (Not moCn.RecSeek("Select PMCd From Param where PTyp= 'TC' and " + _
                                      "PMCd= '" + NewValue + "' and " + _
                                      "PValue= '" + GrdParam.Value(RowNum, "PMCd") + "'"))
        ErrMsg = "Invalid Default Tc": Exit Sub
      '*************************** ZUBIN **************************
      ' 11th Dec 2003, EMR206
      Case Is = "PRTCD"
        Cancel = NewValue > 100
        ErrMsg = "Dia% Fr Cannot Be > 100": Exit Sub
      '***************************   **************************
      '      Case Is = "WORK"
'        If ws_BlnkYN = "N" And NewValue = "" Then _
'           Cancel = True: ErrMsg = "Process Location Should Be Entered": Exit Sub
'        Cancel = (NewValue <> "") And (Not moCn.RecSeek("Select LocCd From Loc where " + _
'                                      "LocTyp= 'P' and LocCd = '" + NewValue + "'"))
'        ErrMsg = "Invalid Process Location": Exit Sub

        '*** (Jen 2.13) (07/11/06)
        ' **** Zubin 211 **** '
        '******Geeta****2.10
        Case Is = "SYSOPT"
          If ws_BlnkYN = "N" And NewValue = "" Then _
             Cancel = True: ErrMsg = "System Option Should Be Entered (Enter Y/N?)": Exit Sub
          Cancel = (NewValue <> "") And (Not moCn.RecSeek("Select PMCd From Param where PTyp= 'YN' and " + _
                                        "PMCd= '" + NewValue + "' and  PSCd=''"))
          ErrMsg = "Invalid System Option": Exit Sub
        '******Geeta****2.10
        ' **** Zubin 211 **** '
        '*** (Jen 2.13) (07/11/06)
        
        '*** Jay 2.13(CT) ***
        Case Is = "BCHR"
          Cancel = Not moCn.RecSeek("Select 'x' From Param where PTyp= 'BTYP' and " + _
                        "PMCd= '" + NewValue + "' and PSCd= '' ")
          ErrMsg = "Invalid Bag Type Option": Exit Sub
        '*** Jay 2.13(CT) ***

        '*** Jay 2.14 *** (SEO Change)
        Case Is = "USR"
          If NewValue = "" Then
            Cancel = moCn.RecSeek("Select 'x' From UsrAccess where " + _
                                  "UaUsrCd = '" + GrdParam.Value(RowNum, "PMCD") + "' and UaSeoYn = 'Y'")
            ErrMsg = "Password Can Not Be Blank When User Has Seo Rights": Exit Sub
          End If
        '*** Jay 2.14 *** (SEO Change)
        'pg.11-Checking valid Gst base & GstOn
        Case Is = "GSTCD"
          If ws_BlnkYN = "N" And NewValue = "" Then _
            Cancel = True: ErrMsg = "GST Base Should Be Entered": Exit Sub
          Cancel = (NewValue <> "") And (Not moCn.RecSeek("Select PMCd From Param where " + _
                                      "PTyp= 'GSTBASE' and PMCd = '" + NewValue + "'"))
          ErrMsg = "Invalid GST Base": Exit Sub
        Case Is = "GSTFLDS"
          If ws_BlnkYN = "N" And NewValue = "" Then _
            Cancel = True: ErrMsg = "GST On Should Be Entered": Exit Sub
          Cancel = (NewValue <> "") And (Not moCn.RecSeek("Select PMCd From Param where " + _
                                      "PTyp= 'GSTON' and PMCd = '" + NewValue + "'"))
          ErrMsg = "Invalid GST On": Exit Sub
      End Select
      
      If ws_BlnkYN = "N" And NewValue = "" Then _
         Cancel = True: ErrMsg = "Value Should Be Entered": Exit Sub
    Case Is = UCase("PNum")
      Select Case UCase(ADC("wPTyp"))
      Case Is = "DMCTG"
        If ws_BlnkYN = "N" And NewValue = 0 Then _
           Cancel = True: ErrMsg = "Enter Default Bag Opening Pieces": Exit Sub
      ' Zubin 213
      'Case Is = "OUNCE"
      Case Is = "GMFACTOR"
      ' Zubin 213
        If ws_BlnkYN = "N" And NewValue = 0 Then _
           Cancel = True: ErrMsg = "Enter Value": Exit Sub
      Case Is = "VACTG"
        If ws_BlnkYN = "N" And NewValue = 0 Then _
           Cancel = True: ErrMsg = "Enter Gold Loss % For Value Addition Category": Exit Sub
'***** Shilpa *****
      Case Is = "USR"
        If ws_BlnkYN = "N" And (NewValue < 1 And NewValue > 9) Then _
            Cancel = True: ErrMsg = "Enter User Level Between 1 to 9": Exit Sub
'***** Shilpa *****
      End Select
    Case Is = UCase("PValue1")
      Select Case UCase(ADC("wPTyp"))
'      Case Is = "CHR" + gs_CoCd
'        If ws_BlnkYN = "N" And NewValue = "" Then _
'           Cancel = True: ErrMsg = "Process Location Should Be Entered": Exit Sub
'        Cancel = (NewValue <> "") And (Not moCn.RecSeek("Select LocCd From Loc where " + _
'                                      "LocTyp= 'P' and LocCd = '" + NewValue + "'"))
'        ErrMsg = "Invalid Process Location": Exit Sub
      Case Is = "DAANAFLD"
        If ws_BlnkYN = "N" And NewValue = "" Then _
           Cancel = True: ErrMsg = "Specify Whether The Analysis Field Is Compulsory Or Not": Exit Sub
        Cancel = (NewValue <> "") And (Not moCn.RecSeek("Select PMCd From Param where " + _
                                      "PTyp= 'YN' and PMCd = '" + NewValue + "'"))
        ErrMsg = "Specify Whether The Analysis Field Is Compulsory Or Not": Exit Sub
      Case Is = "PARENTCD"
        If ws_BlnkYN = "N" And NewValue = "" Then _
           Cancel = True: ErrMsg = "Rm Sub Ctg Should Be Entered": Exit Sub
        Cancel = (NewValue <> "") And (Not moCn.RecSeek("Select PSCd From Param where " + _
                                      "PTyp= 'RMSCTG' and PMCd= '" + GrdParam.Value(RowNum, "PValue") + "' and " + _
                                      "PSCd = '" + NewValue + "'"))
        ErrMsg = "Invalid Rm Sub Ctg": Exit Sub
      Case Is = "RMSCTG"
        If ws_BlnkYN = "N" And NewValue = "" Then _
           Cancel = True: ErrMsg = "Enter Yes/No": Exit Sub
        Cancel = (NewValue <> "") And (Not moCn.RecSeek("Select PMCd From Param where " + _
                                      "PTyp= 'YN' and PMCd = '" + NewValue + "'"))
        ErrMsg = "Invalid Yes/No Option"
        
        If NewValue = "N" Then
          If moCn.RecSeek("Select 'a' From SYyMm " + _
                         "  Join " + _
                         "   RmMst on RmCd = SRmCd " + _
                         "  Join " + _
                         "   Param on PTyp = 'RMSCTG' and PMCd = RmMst.RmCtg and PSCd = RmMst.RMSCtg " + _
                         "  Where " + _
                         " (SRmDrQty <>0 or SRmCrQty <>0 or SRmDrWt <> 0 or SRmCrWt <> 0) " + _
                         "  and SRmStkRt <> 0 " + _
                         "  and RmCtg='" + GrdParam.Value(RowNum, "PMCd") + "'" + _
                         "  and RmSCtg='" + GrdParam.Value(RowNum, "PSCd") + "'" + _
                         "") Then
            Cancel = True
            ErrMsg = "StkByRt cannot be set to 'N' as there are Stock records with non zero StkRt values"
          End If
        End If
      Case Is = "USR"
        '*** To Check If All The Non Production Locs In The Seq Are Valid
        ws_Loc = Split(NewValue, ",")
        If IsArray(ws_Loc) Then
          For i = 0 To UBound(ws_Loc)
            Cancel = Not moCn.RecSeek("Select LocCd From Loc where LocTyp not in ('P') " + _
                                      "and LocCd = '" + ws_Loc(i) + "'")
            If Cancel = True Then ErrMsg = "Value(" + CStr(i + 1) + ") Is An Invalid Non Prd Location": Exit Sub
          Next i
        End If
      '*************************** ZUBIN **************************
      ' 11th Dec 2003, EMR206
      Case Is = "PRTCD"
        Cancel = NewValue > 100
        ErrMsg = "Dia% To Cannot Be > 100": Exit Sub
      '*************************** ZUBIN **************************
      
      ' **** Zubin 211 **** '
      Case Is = "BCHR"
        'IG BCHR- Validation for multiselect RmCtg
        ws_RmCtgArr = Split(NewValue, ",")
        If IsArray(ws_RmCtgArr) Then
          For i = 0 To UBound(ws_RmCtgArr)
            Cancel = Not moCn.RecSeek("Select 1 From Param where " + _
                                      "PTyp = 'RMCTG' and PMCd = '" + ws_RmCtgArr(i) + "' And PSCd = ''")
            If Cancel = True Then ErrMsg = "Value ' " + ws_RmCtgArr(i) + " ' is An Invalid RmCtg": Exit Sub
          Next i
        End If

      ' **** Zubin 211 **** '
      'pg.11-Checking valid GstFor
      Case Is = "GSTFLDS"
        If ws_BlnkYN = "N" And NewValue = "" Then _
           Cancel = True: ErrMsg = "GST For Should Be Entered": Exit Sub
        Cancel = (NewValue <> "") And (Not moCn.RecSeek("Select PMCd From Param where " + _
                                      "PTyp= 'GSTFOR' and PMCd = '" + NewValue + "'"))
        ErrMsg = "Invalid GST For": Exit Sub

      End Select
      If ws_BlnkYN = "N" And NewValue = "" Then _
        Cancel = True: ErrMsg = "Value Should Be Entered": Exit Sub
    Case Is = UCase("PValue2")
      Select Case UCase(ADC("wPTyp"))
'      Case Is = "CHR" + gs_CoCd
'        If ws_BlnkYN = "N" And NewValue = "" Then _
'           Cancel = True: ErrMsg = "Enter Debit/Credit": Exit Sub
'        Cancel = (NewValue <> "") And (Not moCn.RecSeek("Select PMCd From Param where " + _
'                                      "PTyp= 'DC' and PMCd = '" + NewValue + "'"))
'        ErrMsg = "Invalid Debit/Credit Option": Exit Sub
      Case Is = "RMSCTG"
        If ws_BlnkYN = "N" And NewValue = "" Then _
           Cancel = True: ErrMsg = "Yes/No Should Be Entered": Exit Sub
        Cancel = (NewValue <> "") And (Not moCn.RecSeek("Select PMCd From Param where " + _
                                      "PTyp= 'YN' and PMCd = '" + NewValue + "'"))
        ErrMsg = "RmPtr Field In The Design Master Entry Is Editable Or Not (Y/N)": Exit Sub
      ' **** Zubin 211 **** '
      Case Is = "BCHR"  'IG BCHR- Validation for RmSCtg, check belong to RmCtg is removed
        Dim ws_RmSCtg() As String
        ws_RmSCtg = Split(NewValue, ",")
        If IsArray(ws_RmSCtg) Then
          For i = 0 To UBound(ws_RmSCtg)
            Cancel = Not moCn.RecSeek("Select 1 From Param where " + _
                                      "PTyp = 'RMSCTG'  And PSCd = '" + ws_RmSCtg(i) + "'")
            If Cancel = True Then ErrMsg = "Value ' " + ws_RmSCtg(i) + " ' Is An Invalid RmSCtg": Exit Sub
          Next i
        End If

      ' **** Zubin 211 **** '
      'Jen Clr Scr
      Case Is = "USR"
        Cancel = (NewValue <> "") And (Not moCn.RecSeek("Select 'x' From ClrSet where CsCd= '" + NewValue + "' "))
        ErrMsg = "Invalid Color Code": Exit Sub
      'Jen Clr Scr
      End Select
      If ws_BlnkYN = "N" And NewValue = "" Then _
         Cancel = True: ErrMsg = "Value Should Be Entered": Exit Sub
    Case Is = UCase("PValue3")
      Select Case UCase(ADC("wPTyp"))
      Case Is = "RMSCTG"
        If ws_BlnkYN = "N" And NewValue = "" Then _
           Cancel = True: ErrMsg = "Yes/No Should Be Entered": Exit Sub
        Cancel = (NewValue <> "") And (Not moCn.RecSeek("Select PMCd From Param where " + _
                                      "PTyp= 'YN' and PMCd = '" + NewValue + "'"))
        ErrMsg = "Accept Ptr Wise Stock In Transaction (Y/N)": Exit Sub
      
      '*** (Jen 3.01)
      Case Is = "USR"
        If ws_BlnkYN = "N" And NewValue = "" Then _
           Cancel = True: ErrMsg = "User Group Should Be Entered": Exit Sub
        Cancel = (NewValue <> "") And (Not moCn.RecSeek("Select PMCd From Param where " + _
                                      "PTyp= 'USRGRP' and PMCd = '" + NewValue + "'"))
        ErrMsg = "Invalid User Group": Exit Sub
      '*** (Jen 3.01)
      'IG BCHR-Production Type should be 'GS' Or 'WS' Or Blank and Rm Category must be D or C
      Case Is = "BCHR"
        If ws_BlnkYN = "N" And NewValue = "" Then _
           Cancel = True: ErrMsg = "Production Type Should Be Entered": Exit Sub
        Cancel = (NewValue <> "") And Not ((NewValue = "GS" Or NewValue = "WS") And (("," + GrdParam.Value(RowNum, "PValue1") + "," Like "*,C,*") Or ("," + GrdParam.Value(RowNum, "PValue1") + "," Like "*,D,*")))
        ErrMsg = "Production Type should be 'GS' Or 'WS' Or Blank \ Rm Category should be D or C": Exit Sub
      Case Is = "LABSCD"
          Cancel = (NewValue <> "") And (Not moCn.RecSeek("Select PSCd From Param where " + _
                                      "PTyp= 'LABLOGIC' and PSCd = '" + NewValue + "' and PMCd='" + GrdParam.Value(RowNum, "PMCd") + "'"))
          ErrMsg = "Invalid LabLogic value"
          
          If GrdParam.Value(RowNum, "PMCd") = "CDW" Or GrdParam.Value(RowNum, "PMCd") = "GHC" Then
            Cancel = (NewValue = "")
            ErrMsg = "LabLogic cannot be blank for this Labor Main Code"
        End If
        
       End Select
    'pg.11-Description for GstFlds should not be blank
    Case Is = UCase("PDesc")
      Select Case UCase(ADC("wPTyp"))
      Case Is = "GSTFLDS"
        Cancel = (NewValue = ""): ErrMsg = "Description Should Be Entered": Exit Sub
      End Select

' ###################################  Manoj #### Ver: 2.0.6  ###################################
' Date:   20/01/2004
    Case Is = UCase("PValidYN")
      Select Case UCase(ADC("wPTyp"))
      ' Manoj 2.10.0
'      Case Is = "LABSCD"
      Case Is = "LABSCD", "DMCOL", "DMCTG", "DMSZ", "PRDCTG", "RMCOL", "SALCTG", "USR"
      ' Manoj 2.10.0
        If ws_BlnkYN = "N" And NewValue = "" Then _
           Cancel = True: ErrMsg = "Yes/No Should Be Entered": Exit Sub
        ' Zubin 212 (24/02/06)
        '*** Jen 2.14 Next Patch 1 [No of Usrs = 4]
        If UCase(ADC("wPTyp")) = "USR" Then
          If NewValue = "Y" And moCn.GetFldVal("Select PValidYN from Param where PTyp= 'USR' and PMCd= '" + GrdParam.Value(RowNum, "PMCd") + "' ") = "Y" Then Exit Sub
          ' ***** Manali 3.7.1 - JMIC included
          '4.1.4 - only MIC allowed as additional user
          Cancel = NewValue = "Y" And moCn.GetFldVal("Select count('x') from Param where PTyp= 'USR' and PValidYN= 'Y'") + 1 > MWLib.NoOfUsers + 1
          If Cancel = True Then ErrMsg = "The Total No. Of Users Cannot Exceed " + CStr(MWLib.NoOfUsers + 1) + ".": Exit Sub

          'Before 4.1.4
          'Cancel = NewValue = "Y" And moCn.GetFldVal("Select count('x') from Param where PTyp= 'USR' and PValidYN= 'Y'") + 1 > MWLib.NoOfUsers + 5
          'If Cancel = True Then ErrMsg = "The Total No. Of Users Cannot Exceed " + CStr(MWLib.NoOfUsers + 5) + ".": Exit Sub
          
          ' Bef 3.7.1
'          Cancel = NewValue = "Y" And moCn.GetFldVal("Select count('x') from Param where PTyp= 'USR' and PValidYN= 'Y'") + 1 > MWLib.NoOfUsers + 4
'          If Cancel = True Then ErrMsg = "The Total No. Of Users Cannot Exceed " + CStr(MWLib.NoOfUsers + 4) + ".": Exit Sub
        End If
        '*** Jen 2.14 Next Patch 1 [No of Usrs = 4]
        ' Zubin 212 (24/02/06)
        Cancel = (NewValue <> "") And (Not moCn.RecSeek("Select PMCd From Param where " + _
                                      "PTyp= 'YN' and PMCd = '" + NewValue + "'"))
        ' Before 2.12
        'If Cancel = True Then ErrMsg = "Labour Sub Code Is Valid Or Not (Y/N)": Exit Sub
        ' Zubin 212 (24/02/06)
        If Cancel = True Then ErrMsg = "Invalid Option. Enter Y/N.": Exit Sub
      End Select
' ###########################################  Manoj  ###########################################
    End Select
  End With
End Sub

Private Sub ADC_SetRecSource()
  '*** Since this is an 'A' Entry, there is no RecSource associated with ADC. So set the
      'DataMode property of ADC control as xdatUnbound
  ADC.DataMode = xdatUnbound
  'wi_CharWdth = 315
  Call GrdPropSet
End Sub

Private Sub GrdParam_ReCalc(ByVal RowNum As Integer, ByVal ColName As String, ByVal OldVal As Variant, ByVal SourceName As String, ByVal SourceColName As String)
  With GrdParam
    Select Case ColName
    Case Is = UCase("PValue1")
      Select Case UCase(ADC("wPTyp"))
      Case Is = "LABSCD"
        If .Mode = fgmnorm Then
          If .Value(RowNum, "PValue1") = "" Then .Value(RowNum, "PValue1") = .Value(RowNum, "PDesc")
        End If
      End Select
    
    
    'URMI 2.06 iNVdESC TO BE AN x ENTRY
    Case Is = UCase("PDesc225")
      Select Case UCase(ADC("wPTyp"))
      Case Is = "INVDESC"
        If .Mode = fgmnorm And .Row > 0 Then
            If ADC("wInvDesc") <> .Value(.Row, "PDesc225") Then
                    .Value(.Row, "PDesc225") = ADC("wInvDesc")
                    .SaveRec .Row
            End If
        End If
      End Select
    '**********************************
    
    End Select
  End With
End Sub


'****************************************
Private Sub GrdParam_RowWhen(ByVal RowNum As Integer)
    
    If UCase(ADC("wPTyp")) = UCase("RTC") Then  'MW.105 - Used to display default PDesc and disable PDesc
      GrdParam.ColProp("PDesc").Style = fgcReadOnly
      If GrdParam.IsNew(RowNum) Then
        GrdParam.Value(RowNum, "PDesc") = "Reports not shown in this time"
      End If
    End If
    
    If UCase(ADC("wPTyp")) = UCase("InvDesc") Then
      If GrdParam.IsNew(RowNum) Then GrdParam.Value(RowNum, "PDesc225") = "Default"
      ADC("WINVDESC") = GrdParam.Value(RowNum, "PDesc225")
    End If
' ###################################  Manoj #### Ver: 2.0.6  ###################################
' Date:   24/04/2004
    If GrdParam.ColProp("PValidYn").Style = fgcNormal And GrdParam.IsNew(RowNum) = True Then
      GrdParam.Value(RowNum, "PValidYn") = "Y"
    End If
' ###########################################  Manoj  ###########################################
  '*************************** ZUBIN **************************
  ' 10th Dec 2003, EMR206
  If UCase(ADC("WPTYP")) = "PRTCD" And GrdParam.IsNew(RowNum) Then
    With GrdParam
      .Value(RowNum, "PValue") = 0
      .Value(RowNum, "PValue1") = 0
      .Value(RowNum, "PValue2") = 0
      .Value(RowNum, "PValue3") = 0
    End With
  End If
  '*************************** ZUBIN **************************
  
  '****** Sachin 3.2.0 - Default Value of PNum1=1(Curr. Partition)
  If UCase(ADC("WPTYP")) = "USR" And GrdParam.IsNew(RowNum) Then
    With GrdParam
      .Value(RowNum, "PNum1") = 1
    End With
  End If
  'EmrSv - Uma
  If UCase(ADC("wPTyp")) = "DMCTG" And GrdParam.IsNew(RowNum) Then
  GrdParam.Value(RowNum, "PDesc225") = moCn.GetFldVal("Select hDmPrcsSeq from Head where HCoCd='" + gs_CoCd + "'")
  End If
  
End Sub

Private Sub GrdParam_SetRecSource()
  '*** Set the Record Source of the Grid GrdParam
      'Sort the Recs in the grid By PMCd, PSCd
  Dim ws_Cnd As String, ws_OrdBy As String
  ws_Cnd = ADC.RepCond
  ws_Cnd = IIF(ws_Cnd <> "", " where " + ws_Cnd, "")
  ws_OrdBy = " Order By PMCd, PSCd "
  GrdParam.RecSource = " Select * from Param " + ws_Cnd + ws_OrdBy
End Sub
Private Sub GrdParam_InitKey(ByVal RowNum As Integer)
  '*** Initialise Key Field PTyp
  With GrdParam
    If .IsNew(RowNum) Then _
      .Store "PTyp", ADC("wPTyp")
  End With

End Sub
Private Sub CmdPwd_Click()
  Call DispFra(Pwd)
End Sub

Private Sub MWCTL_BTN11_Click()

End Sub

Private Sub TxtPwd_GotFocus()
  DispMsg "Enter SEO Password", etInfo
End Sub

Private Sub TxtPwd_LostFocus()
Call TxtPwd_Validate(False)
End Sub

Private Sub TxtPwd_Validate(Cancel As Boolean)
'*** Jay 2.14 *** (SEO Change)
  If txtPwd.Text = "" Then
      ms_PwdFlag = moCn.RecSeek("Select PValue1 from Param where PTyp= 'SEOPWD' and PmCd= 'SEOPWD' and PValue1= ''")
  Else
      If Trim(txtPwd.Text) = moCn.GetFldVal("Select IsNull ((select PValue from Param Join UsrAccess on UaUsrCd = PMCd " + _
                       "where PTyp= 'USR' and PMCd= '" + gs_UsrCd + "' and UaMnuCd = '" + ADC.MenuCd + "' and " + _
                       "UaSeoYn= 'Y' and PValue= '" + Trim(txtPwd.Text) + "'), " + _
                       "(Select PValue1 from Param where Ptyp= 'SEOPWD' and PmCd= 'SEOPWD' )) ") Then
        ms_PwdFlag = True: Exit Sub
      Else
        ms_PwdFlag = False
        Cancel = True: DispMsg "Invalid SEO Password / User Does Not Have SEO Rights", etError
      End If
  End If
'  If Trim(TxtPwd.text) = "" Then Exit Sub
'  If Not Trim(TxtPwd.text) = moCn.GetFldVal("Select PValue1 from Param where " + _
'                                            "PTyp= 'SEOPWD' and PmCd= 'SEOPWD'") Then
'    Cancel = True: DispMsg "Invalid SEO Password", etError
'  End If
'*** Jay 2.14 *** (SEO Change)
End Sub
Private Sub GrdPropSet()
  '*** Write a procedure to set the properties like Style, Heading, DataType, Mask,
      'MaxLength of the Grid GrdParam using the TcStr table
      'Note that the heading has to be appended by spaces when necessary if the
      'column width is not accomodating the length of the field

  Dim wo_RsTc As MwfLib.MDORowSet
  Dim wi_SpaceLen As Single
  Set wo_RsTc = moCn.OpenRes("Select * from TcStr where TsTbl='Param' and " + _
                            "TsTyp='" + ADC("wPTyp") + "'")
  With GrdParam
    Do While Not (wo_RsTc.EOF Or wo_RsTc.BOF)
      .ColProp(wo_RsTc!TsIdName).Style = ColSty(wo_RsTc!TsStyle)
      '***** To Find Out The Space To Be Appended To The Heading
            '(This Is To Set The Proper Width For The Field)
      '***** 2.5 is taken as the approximate value which stands for the space occupied by one
            'character in Upper Case
      wi_SpaceLen = (wo_RsTc!TsMaxLen * 2.5) - (Len(wo_RsTc!TsLbl))
      .ColProp(wo_RsTc!TsIdName).Heading = wo_RsTc!TsLbl + IIF(wi_SpaceLen > 0, Space(Abs(wi_SpaceLen)), "")
      
      .ColProp(wo_RsTc!TsIdName).DataType = DataTyp(wo_RsTc!TsDataTyp)
      .ColProp(wo_RsTc!TsIdName).Mask = wo_RsTc!TsMask
      .ColProp(wo_RsTc!TsIdName).MaxLength = wo_RsTc!TsMaxLen
      
      '***** Note: Width Property cannot be used to set the width
      '*****  .ColProp(wo_RsTc!TsIdName).Width = wi_CharWdth * wo_RsTc!TsMaxLen
      wo_RsTc.MoveNext
    Loop
  End With
  GrdParam.SetColWidthToHd
  Set wo_RsTc = Nothing
End Sub
Private Function ColSty(ByVal mColSty As String) As en_fgColStyle
  '*** Write a function to return the Grid Column Style based on the value of the
      'TsStyle field of the TcStr table
  Select Case mColSty
    Case Is = "N"
      ColSty = fgcNormal
    Case Is = "I"
      ColSty = fgcinvisible
    Case Is = "R"
      ColSty = fgcReadOnly
  End Select
End Function
Private Function DataTyp(ByVal mDataTyp As String) As en_FldDataType
  '*** Write a function to return the datatype of a column based on the value of the
      'TsDataTyp field of the TcStr table
  Select Case mDataTyp
    Case Is = "C"
      DataTyp = fdtCharType
    Case Is = "N"
      DataTyp = fdtNumber
    Case Is = "F"
      DataTyp = fdtFloat
    Case Is = "D"
      DataTyp = fdtDateType
    Case Is = "U"
      DataTyp = fdtConvertToUpper
    Case Is = "T"
      DataTyp = fdtTimeType
  End Select
End Function
Private Sub DispFra(ByVal pv_PrmFra As en_PrmFra)
  '*** This is the common procedure which is used to toggle between hiding and displaying
      'the frames (i.e. if the frame is visible then make it invisible and viseversa)
      'associated with a command button
  Select Case pv_PrmFra
  Case Is = Pwd
    If FraPwd.Visible = True Then
      FraPwd.Visible = False
      FraPwd.Enabled = False
      Call EnaDisaCmds(False, CmdPwd)
      CmdPwd.SetFocus
    Else
      FraPwd.Visible = True
      FraPwd.Enabled = True
      CmdPwd.SetFocus
      FraPwd.ZOrder
      Call EnaDisaCmds(True, CmdPwd)
    End If
  Case Is = ReAlloc
   If FraReAllotPrt.Visible = True Then
      FraReAllotPrt.Enabled = False
      FraReAllotPrt.Visible = False
      Call EnaDisaCmds(False, cmdReAllotPrt)
      cmdReAllotPrt.SetFocus
    Else
      FraReAllotPrt.Enabled = True
      FraReAllotPrt.Visible = True
      cmdReAllotPrt.SetFocus
      FraReAllotPrt.ZOrder
      Call EnaDisaCmds(True, cmdReAllotPrt)
    End If
  Case Is = UpdDsgAna
   If FraUpdDsgAna.Visible = True Then
      FraUpdDsgAna.Enabled = False
      FraUpdDsgAna.Visible = False
      Call EnaDisaCmds(False, CmdUpdDsgAna)
      CmdUpdDsgAna.SetFocus
    Else
      FraUpdDsgAna.Enabled = True
      FraUpdDsgAna.Visible = True
      CmdUpdDsgAna.SetFocus
      FraUpdDsgAna.ZOrder
      Call EnaDisaCmds(True, CmdUpdDsgAna)
    End If
  Case Is = UpdDsgLab
   If FraUpdDsgLab.Visible = True Then
      FraUpdDsgLab.Enabled = False
      FraUpdDsgLab.Visible = False
      Call EnaDisaCmds(False, CmdUpdDsgLab)
      CmdUpdDsgLab.SetFocus
    Else
      FraUpdDsgLab.Enabled = True
      FraUpdDsgLab.Visible = True
      CmdUpdDsgLab.SetFocus
      FraUpdDsgLab.ZOrder
      Call EnaDisaCmds(True, CmdUpdDsgLab)
    End If
    
    '*************Bhavna
'  Case Is = DelDsgAna
'   If FraDelDsgAna.Visible = True Then
'      FraDelDsgAna.Enabled = False
'      FraDelDsgAna.Visible = False
'      Call EnaDisaCmds(False, CmdDelDsgAna)
'      CmdDelDsgAna.SetFocus
'    Else
'      FraDelDsgAna.Enabled = True
'      FraDelDsgAna.Visible = True
'      CmdDelDsgAna.SetFocus
'      FraDelDsgAna.ZOrder
'      Call EnaDisaCmds(True, CmdDelDsgAna)
'    End If
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
  'CmdPwd.Enabled = Not pv_ShowFra
  cmdReAllotPrt.Enabled = Not pv_ShowFra
  CmdUpdDsgAna.Enabled = Not pv_ShowFra
  CmdUpdDsgLab.Enabled = Not pv_ShowFra
  
  CmdDelDsgAna.Enabled = Not pv_ShowFra     '**Bhavna
  If Not pv_CmdBut Is Nothing Then pv_CmdBut.Enabled = True
End Sub
Private Sub HideAllFras()
  '*** This procedure hides all the frames other than FraNKeyAll
  FraNKeyAll.Visible = True
  FraPwd.Visible = False
  FraPwd.Enabled = False
  FraUpdDsgAna.Visible = False
  FraUpdDsgAna.Enabled = False
  FraRTCLst.Visible = False
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
  ws_BtnStr = "CmdRpGo"
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
Private Sub GrdParam_Validate(Cancel As Boolean)
  Cancel = GrdParam.Validate
End Sub





