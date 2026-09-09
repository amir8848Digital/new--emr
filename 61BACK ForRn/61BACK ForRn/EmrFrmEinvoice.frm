VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Object = "{AB2CC32A-E652-42A2-B341-26ECE02F6386}#2.1#0"; "MwfCtl.ocx"
Begin VB.Form EmrFrmEinvoice 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "eInvoice"
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
      Left            =   4560
      TabIndex        =   28
      Top             =   7680
      Width           =   5955
      Begin MwfCtl.MWCTL_MED atxt 
         Height          =   285
         Index           =   248
         Left            =   540
         TabIndex        =   29
         ToolTipText     =   "Excel File Name and Path"
         Top             =   555
         Width           =   4635
         _ExtentX        =   8176
         _ExtentY        =   503
         Enabled         =   0   'False
         MaxLength       =   250
         IdName          =   "WXLFILE"
      End
      Begin MwfCtl.MWCTL_MED atxt 
         Height          =   285
         Index           =   249
         Left            =   800
         TabIndex        =   30
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
      Begin MwfCtl.MWCTL_MED atxt 
         Height          =   285
         Index           =   250
         Left            =   2460
         TabIndex        =   31
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
      Begin MwfCtl.MWCTL_MED atxt 
         Height          =   285
         Index           =   251
         Left            =   4350
         TabIndex        =   32
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
         Left            =   5280
         TabIndex        =   33
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
         Left            =   5280
         TabIndex        =   34
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
         Left            =   3630
         TabIndex        =   39
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
         Left            =   1800
         TabIndex        =   38
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
         Left            =   60
         TabIndex        =   37
         Top             =   1065
         Width           =   735
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
         Left            =   60
         TabIndex        =   36
         Top             =   555
         Width           =   375
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
         TabIndex        =   35
         Top             =   120
         Width           =   5950
      End
   End
   Begin VB.Frame FRA_KEYS 
      BorderStyle     =   0  'None
      Height          =   9255
      Left            =   0
      TabIndex        =   12
      Top             =   0
      Width           =   15210
      Begin VB.Frame FraCpyErr 
         Height          =   2775
         Left            =   0
         TabIndex        =   41
         Top             =   0
         Width           =   13815
         Begin MwfCtl.MWCTL_FLX GrdCpyErr 
            Height          =   2625
            Left            =   60
            TabIndex        =   42
            Top             =   120
            Width           =   13665
            _ExtentX        =   24104
            _ExtentY        =   4630
            Cols            =   9
            colname1        =   "OCESRCCOCD"
            datafld1        =   "OceSrcCoCd"
            recalcparent1   =   "GrdOrdCpyErr(),GrdOrdCpyErr(WSRCVCH)"
            maxlength1      =   3
            style1          =   0
            colname2        =   "OCESRCTC"
            heading2        =   " "
            datafld2        =   "OceSrcTc"
            recalcparent2   =   "GrdOrdCpyErr(),GrdOrdCpyErr(WSRCVCH)"
            maxlength2      =   3
            style2          =   0
            colname3        =   "OCESRCYY"
            datafld3        =   "OceSrcYy"
            recalcparent3   =   "GrdOrdCpyErr(),GrdOrdCpyErr(WSRCVCH)"
            maxlength3      =   2
            style3          =   0
            colname4        =   "OCESRCCHR"
            datafld4        =   "OceSrcChr"
            recalcparent4   =   "GrdOrdCpyErr(),GrdOrdCpyErr(WSRCVCH)"
            maxlength4      =   3
            style4          =   0
            colname5        =   "OCESRCNO"
            datafld5        =   "OceSrcNo"
            datatype5       =   1
            mask5           =   "#####0"
            recalcparent5   =   "GrdOrdCpyErr(),GrdOrdCpyErr(WSRCVCH)"
            maxlength5      =   6
            style5          =   0
            colname6        =   "WSRCVCH"
            heading6        =   "Source Voucher                         "
            recalcon6       =   "GrdOrdCpyErr(OCESRCCOCD),GrdOrdCpyErr(OCESRCTC),GrdOrdCpyErr(OCESRCYY),GrdOrdCpyErr(OCESRCCHR),GrdOrdCpyErr(OCESRCNO)"
            maxlength6      =   25
            style6          =   2
            colname7        =   "OCESRCSR"
            heading7        =   "Src Sr "
            datafld7        =   "OceSrcSr"
            datatype7       =   1
            mask7           =   "####0"
            maxlength7      =   5
            style7          =   2
            colname8        =   "OCEERR"
            heading8        =   $"EmrFrmEinvoice.frx":0000
            datafld8        =   "OceErr"
            maxlength8      =   250
            style8          =   2
         End
      End
      Begin MwfCtl.MWCTL_MED atxt 
         Height          =   285
         Index           =   0
         Left            =   2880
         TabIndex        =   1
         ToolTipText     =   "Enter From Invoice Tc"
         Top             =   870
         Width           =   675
         _ExtentX        =   1191
         _ExtentY        =   503
         MaxLength       =   3
         DataType        =   4
         IdName          =   "WIDTCFR"
         CmpStr          =   "IdTc>="
      End
      Begin MwfCtl.MWCTL_MED atxt 
         Height          =   285
         Index           =   1
         Left            =   3660
         TabIndex        =   2
         ToolTipText     =   "Enter From Invoice Year"
         Top             =   870
         Width           =   345
         _ExtentX        =   609
         _ExtentY        =   503
         MaxLength       =   2
         DataType        =   4
         IdName          =   "WIDYYFR"
         CmpStr          =   "IdYy>="
      End
      Begin MwfCtl.MWCTL_MED atxt 
         Height          =   285
         Index           =   2
         Left            =   4110
         TabIndex        =   3
         ToolTipText     =   "Enter From Invoice Character"
         Top             =   870
         Width           =   675
         _ExtentX        =   1191
         _ExtentY        =   503
         MaxLength       =   3
         DataType        =   4
         IdName          =   "WIDCHRFR"
         CmpStr          =   "IdChr>="
      End
      Begin MwfCtl.MWCTL_MED atxt 
         Height          =   285
         Index           =   3
         Left            =   6810
         TabIndex        =   5
         ToolTipText     =   "Enter To Invoice Tc "
         Top             =   870
         Width           =   675
         _ExtentX        =   1191
         _ExtentY        =   503
         MaxLength       =   3
         DataType        =   4
         IdName          =   "WIDTCTO"
         CmpStr          =   "IdTc<="
      End
      Begin MwfCtl.MWCTL_MED atxt 
         Height          =   285
         Index           =   6
         Left            =   4890
         TabIndex        =   4
         ToolTipText     =   "Enter From Invoice No"
         Top             =   870
         Width           =   795
         _ExtentX        =   1402
         _ExtentY        =   503
         Alignment       =   2
         Mask            =   "#####0"
         MaxLength       =   6
         DataType        =   1
         IdName          =   "WIDNOFR"
         CmpStr          =   "IdNo>="
      End
      Begin MwfCtl.MWCTL_MED atxt 
         Height          =   285
         Index           =   4
         Left            =   7590
         TabIndex        =   6
         ToolTipText     =   "Enter To Invoice Year"
         Top             =   870
         Width           =   345
         _ExtentX        =   609
         _ExtentY        =   503
         MaxLength       =   2
         DataType        =   4
         IdName          =   "WIDYYTO"
         CmpStr          =   "IdYy<="
      End
      Begin MwfCtl.MWCTL_MED atxt 
         Height          =   285
         Index           =   5
         Left            =   8040
         TabIndex        =   7
         ToolTipText     =   "Enter To Invoice Character"
         Top             =   870
         Width           =   675
         _ExtentX        =   1191
         _ExtentY        =   503
         MaxLength       =   3
         DataType        =   4
         IdName          =   "WIDCHRTO"
         CmpStr          =   "IdChr<="
      End
      Begin MwfCtl.MWCTL_MED atxt 
         Height          =   285
         Index           =   8
         Left            =   8820
         TabIndex        =   8
         ToolTipText     =   "Enter To Invoice No"
         Top             =   870
         Width           =   795
         _ExtentX        =   1402
         _ExtentY        =   503
         Alignment       =   2
         Mask            =   "#####0"
         MaxLength       =   6
         DataType        =   1
         IdName          =   "WIDNOTO"
         CmpStr          =   "IdNo<="
      End
      Begin MwfCtl.MWCTL_MED atxt 
         Height          =   285
         Index           =   31
         Left            =   2880
         TabIndex        =   0
         ToolTipText     =   "Enter Company Code"
         Top             =   600
         Width           =   675
         _ExtentX        =   1191
         _ExtentY        =   503
         MaxLength       =   3
         DataType        =   4
         IdName          =   "WIDCOCD"
         CmpStr          =   "IdCoCd = "
      End
      Begin MwfCtl.MWCTL_MED atxt 
         Height          =   285
         Index           =   7
         Left            =   2880
         TabIndex        =   9
         ToolTipText     =   "Enter From Invoice Date"
         Top             =   1155
         Width           =   915
         _ExtentX        =   1614
         _ExtentY        =   503
         MaxLength       =   8
         DataType        =   3
         IdName          =   "WINDTFR"
         CmpStr          =   "InDt >= "
      End
      Begin MwfCtl.MWCTL_MED atxt 
         Height          =   285
         Index           =   9
         Left            =   6810
         TabIndex        =   10
         ToolTipText     =   "Enter To Invoice Date"
         Top             =   1155
         Width           =   915
         _ExtentX        =   1614
         _ExtentY        =   503
         MaxLength       =   8
         DataType        =   3
         IdName          =   "WINDTTO"
         CmpStr          =   "InDt <= "
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Inv Date"
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
         Left            =   180
         TabIndex        =   40
         Top             =   1170
         Width           =   780
      End
      Begin VB.Label ALBL 
         BackStyle       =   0  'Transparent
         Caption         =   "Company Code"
         BeginProperty Font 
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
         Left            =   180
         TabIndex        =   27
         Top             =   600
         Width           =   2145
      End
      Begin VB.Label Label6 
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
         Index           =   3
         Left            =   6840
         TabIndex        =   25
         Top             =   240
         Width           =   240
      End
      Begin VB.Label Label6 
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
         Index           =   4
         Left            =   2880
         TabIndex        =   24
         Top             =   240
         Width           =   615
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
         Left            =   8740
         TabIndex        =   22
         Top             =   870
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
         Left            =   7510
         TabIndex        =   21
         Top             =   870
         Width           =   60
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
         Left            =   7960
         TabIndex        =   20
         Top             =   870
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
         Left            =   4810
         TabIndex        =   19
         Top             =   870
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
         Left            =   4030
         TabIndex        =   17
         Top             =   870
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
         Left            =   3590
         TabIndex        =   16
         Top             =   870
         Width           =   60
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Inv Tc/ Yy/ Chr/ No"
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
         Left            =   180
         TabIndex        =   15
         Top             =   885
         Width           =   1740
      End
   End
   Begin MwfCtl.MWCTL_BTN1 CmdExpExc 
      Height          =   420
      Left            =   6300
      TabIndex        =   26
      Top             =   9270
      Width           =   945
      _ExtentX        =   1667
      _ExtentY        =   741
      BackColor       =   -2147483633
      MaskColor       =   8454143
      Caption         =   "Export"
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
   Begin MwfCtl.MWCTL_BTN1 CmdImpExc 
      Height          =   420
      Left            =   7260
      TabIndex        =   14
      Top             =   9270
      Width           =   945
      _ExtentX        =   1667
      _ExtentY        =   741
      BackColor       =   -2147483633
      MaskColor       =   8454143
      Caption         =   "Import"
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
      Left            =   9750
      TabIndex        =   13
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
      TabIndex        =   11
      Top             =   9270
      Width           =   14985
      _ExtentX        =   26432
      _ExtentY        =   979
      AllowAdd        =   0   'False
      AllowFind       =   0   'False
      AllowDelete     =   0   'False
      BtnWidth        =   885
      BtnHeight       =   435
      BtnMaskColor    =   8454143
      FontSz          =   8.25
   End
   Begin VB.Frame FRA_NKEYS 
      BorderStyle     =   0  'None
      Enabled         =   0   'False
      Height          =   9030
      Left            =   -150
      TabIndex        =   18
      Top             =   495
      Width           =   15315
      Begin VB.Frame FraNKeyAll 
         BorderStyle     =   0  'None
         Height          =   8745
         Left            =   60
         TabIndex        =   23
         Top             =   30
         Width           =   15210
      End
   End
   Begin MSComDlg.CommonDialog CD1 
      Left            =   15480
      Top             =   0
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
End
Attribute VB_Name = "EmrFrmEinvoice"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'---- General Comments on the Entry --------'
Option Explicit
Dim moCn As MwfLib.MDOConnection
Dim ws_TcFr As String, ws_TcTo As String, ws_YyFr As String, ws_YyTo As String
Dim ws_CoCdFr As String, ws_CoCdTo As String
Dim ws_ChrFr As String, ws_ChrTo As String
Dim mi_FraKeysHt As Integer
'Dim ws_CoCd As String, ws_Tc As String, ws_Yy As String

Enum en_eInvFra
  ImpExc = 1
End Enum

Private Sub ADC_Load()
  '*** Set the First Key Control property to the working field
    
  '*** Set the Previous and Next Control Properties for the Grid GrdDt
  
  Set moCn = ADC.Connection    '*** (Bef speed) bef 24/11/06
  
  FRA_KEYS.ZOrder (0)
  Set ADC.FirKeyCtl = ADC("wIdCoCd")
  
 
  ws_CoCdFr = gs_CoCd: ws_CoCdTo = gs_CoCd
  ws_TcFr = ctInTcTyp: ws_TcTo = ctInTcTyp
  ws_YyFr = moCn.GetFldVal("Select HYy From Head Where HCd='" + ctSelfCmCd + "' And HCoCd='" + gs_CoCd + "'")
  ws_YyTo = moCn.GetFldVal("Select HYy From Head Where HCd='" + ctSelfCmCd + "' And HCoCd='" + gs_CoCd + "'")
  'ws_ChrFr = moCn.GetFldVal("Select vPValue from vParam where vPCoCd='" + gs_CoCd + "' and vPTyp='DEFCHR' and vPMCd='" + ws_TcFr + "' and vPValue<> ''")
  'ws_ChrTo = moCn.GetFldVal("Select vPValue from vParam where vPCoCd='" + gs_CoCd + "' and vPTyp='DEFCHR' and vPMCd='" + ws_TcTo + "' and vPValue<> ''")
    
End Sub
Private Sub ADC_KeyWhen()

  
  
  '****** Sachin 2.14.0 - [24-11-2006] - [Multi-Company] ******
  Me.Caption = GetFrmCaption(gs_CoCd, "eInvoice")
  '****** Sachin 2.14.0 - [24-11-2006] - [Multi-Company] ******
  
  FraNKeyAll.ZOrder (0)
 
   
  ADC("wIdCoCd") = ws_CoCdFr
  'ADC("wIdCoCdTo") = ws_CoCdTo
  ADC("wIdTcFr") = ws_TcFr
  ADC("wIdTcTo") = ws_TcTo
  ADC("wIdYyFr") = ws_YyFr
  ADC("wIdYyTo") = ws_YyTo
  ADC("wInDtFr") = IIF(ADC("wInDtFr") <> "", ADC("wInDtFr"), "01/01/80")
  ADC("wInDtTo") = IIF(ADC("wInDtTo") <> "", ADC("wInDtTo"), "01/01/80")
'  ADC("wIdChrFr") = ws_ChrFr
'  ADC("wIdChrTo") = ws_ChrTo
  
End Sub

Private Sub ADC_SetRecSource()
    '*** Since this is an 'A' Entry, there is no RecSource associated with ADC. So set the
    'DataMode property of ADC control as xdatUnbound
  ADC.DataMode = xdatUnbound
  
   
End Sub

Private Sub ATXT_FldChange(Index As Integer, ByVal IdName As String, Cancel As Boolean, ErrMsg As String, ByVal pv_NewValue As Variant)
Select Case IdName
  'Case Is = UCase("wIdCoCdFr")
  '  ADC("wIdCoCdTo") = pv_NewValue
  Case Is = UCase("wIdTcFr")
    ADC("wIdTcTo") = pv_NewValue
  Case Is = UCase("wIdYyFr")
    ADC("wIdYyTo") = pv_NewValue
  Case Is = UCase("wIdChrFr")
    ADC("wIdChrTo") = pv_NewValue
  Case Is = UCase("wIdNoFr")
    ADC("wIdNoTo") = pv_NewValue
End Select
End Sub

Private Sub atxt_FldWhen(Index As Integer, ByVal IdName As String, Cancel As Boolean, ErrMsg As String)
Select Case IdName
  Case Is = UCase("wIdCoCd") ', UCase("wIdCoCdTo")
    Call HlpList.hCoCd
  Case Is = UCase("wIdTcFr"), UCase("wIdTcTo")
    Call HlpList.PMCd("TC", "'IN'")
  Case Is = UCase("wIdYyFr"), UCase("wIdYyTo")
    Call HlpList.PMCd("YY")
  Case Is = UCase("wIdChrFr")
    Call HlpList.vPSCd(ADC("wIdCoCd"), "CHR", ADC("wIdTcFr"), "Y")
  Case Is = UCase("wIdChrTo")
    'Call HlpList.vPSCd(ADC("wIdCoCdTo"), "CHR", ADC("wIdTcTo"), "Y")
    Call HlpList.vPSCd(ADC("wIdCoCd"), "CHR", ADC("wIdTcTo"), "Y")
  Case Is = UCase("wIdNoFr")
    Call HlpList.InNo(ADC("wIdCoCd"), ADC("wIdTcFr"), ADC("wIdYyFr"), ADC("wIdChrFr"))
  Case Is = UCase("wIdNoTo")
    Call HlpList.InNo(ADC("wIdCoCd"), ADC("wIdTcTo"), ADC("wIdYyTo"), ADC("wIdChrTo"))
    
End Select
End Sub


Private Sub CmdExpExc_Click()
'copy records from result grid to excel file
  On Error GoTo Err_Hndlr
  Dim wExcVer As String, wExtn As String, sFilePath As String, wErrMsg As String
  Dim oExcel, oBook, oSheet
  Set oExcel = GetExcelObj()
  wExcVer = oExcel.version
 
  
'  wExtn = ".xlsm"
'  CD1.DialogTitle = "Open Excel File"
'  CD1.Filter = "Excel Files (*" + wExtn + ") |*" + wExtn
'  CD1.Flags = cdlOFNOverwritePrompt Or cdlOFNLongNames Or cdlOFNPathMustExist Or cdlOFNExtensionDifferent
'  CD1.CancelError = True
'  If (SelectFileName(wExtn, sFilePath)) = False Then Exit Sub
'
'
''  CD1.ShowOpen
''  sFilePath = CD1.FileName
'
'  'FileCopy App.Path + "\Template\eInvoiceTmp" + wExtn, sFilePath
'  oExcel.Visible = False: oExcel.DisplayAlerts = False
'
'If Dir(sFilePath) = "" Then MsgBox ("File Not Found, Please Check the File Name"): Exit Sub
'
'
''  Set oBook = oExcel.Workbooks.Open("C:\MwApp\MwPic\Template\JSON-Copy.xlsm")
'  Set oBook = oExcel.Workbooks.Open(sFilePath)
'  DoEvents
'  'Set oBook = oExcel.Workbooks.Load(sFilePath)
'  'Set oBook = CD1.FileName
''xlApp.Application.OleRequestPendingTimeout = 0
''oExcel.Application.OleRequestPendingTimeout = 0
'
'  Set oSheet = oBook.worksheets(3)
'********************************************************************************
  
  'wExtn = ".xlsm"
  'CD1.DialogTitle = "Save Excel File As"
  'CD1.FileName = ADC("wIdTcFr") + "-" + ADC("wIdYyFr") + "-" + ADC("wIdChrFr") + "-" + CStr(ADC("wIdNoFr")) + wExtn
  'CD1.FileName = Format(Now(), "YYYYmmDD_HHMMss") + "_01"
  'CD1.Filter = "Excel Files (*" + wExtn + ") |*" + wExtn
  'CD1.Flags = cdlOFNOverwritePrompt Or cdlOFNLongNames Or cdlOFNPathMustExist Or cdlOFNExtensionDifferent
  'CD1.CancelError = True
  'If (SelectFileName(wExtn, sFilePath)) = False Then Exit Sub
  'FileCopy App.Path + "\Template\JSON-Copy" + wExtn, sFilePath
  Dim wSrcFile As String, wTgtFile As String, wFileNm As String, wPth As String
  wPth = moCn.GetFldVal("Select PDesc225 from Param Where PTyp='PCT' and PMCd='3D'")
           
  wSrcFile = wPth + "\eInvoice\Template\JSON-Copy.xlsm"
  wFileNm = Format(Now(), "YYYYmmDD_HHMMss") + "_01.xlsm"
  'wFileNm = Format(Now(), "YYYYmmDD") + "_01.xlsm"
  wTgtFile = wPth + "\eInvoice\" + wFileNm
  
  If Dir(wTgtFile) <> "" Then Kill wTgtFile
  
  FileCopy wSrcFile, wTgtFile
  oExcel.Visible = False: oExcel.DisplayAlerts = False

  Set oBook = oExcel.Workbooks.Open(wTgtFile)
  Set oSheet = oBook.worksheets(3)
  
  
'********************************************************************************


  Dim wGrdRows As Integer, wi_Row As Integer, wExlRw As Integer
  Dim wRsInv As MwfLib.MDORowSet, wsql As String, wrepcnd As String
  wrepcnd = ADC.RepCond
  wsql = "Select * From InvDsg join InvHd on InIdNo=IdInIdNo and InCoCd= IdCoCd and InTc= IdTc and InYy= IdYy " + _
          " and InChr= IdChr and InNo= IdNo " + IIF(gs_Partition = ctCurrPrtn, " and InPrtKey=IdPrtKey ", "") + _
          " where " + wrepcnd
       
  Set wRsInv = moCn.OpenRes(wsql)
'new
If wRsInv.RecCount = 0 Then DispMsg "No Record Found for Given Scope.", etInfo: CloseExcel: Exit Sub
   wGrdRows = 1
    wExlRw = 15
    'For wi_Row = 1 To wGrdRows
    Do While Not (wRsInv.EOF Or wRsInv.BOF)
      'oSheet.Cells(wExlRw, 33) = "b2b"
      oSheet.Cells(wExlRw, 33) = wRsInv!IdNo
      oSheet.Cells(wExlRw, 34) = wRsInv!IdQty
      wExlRw = wExlRw + 1
      wRsInv.MoveNext
    'Next
    Loop
    
   Set wRsInv = Nothing
  Set oSheet = oBook.worksheets(3)
  DispMsg "Updated Excel Successfully !!", etInfo
  MsgBox "Updated Excel Successfully !!"
  oExcel.ActiveWorkbook.Close True
  oExcel.QUIT

  oExcel.DisplayAlerts = True
  Set oExcel = Nothing
  Set oBook = Nothing
  Set oSheet = Nothing
  Exit Sub
Err_Hndlr:
  If Err.Number = 32755 Then
    wErrMsg = "Operation Cancelled"
    oExcel.ActiveWorkbook.Close
    oExcel.QUIT
  ElseIf Err.Number = 429 Then
    DispMsg "MS Excel is either Corrupt Or Missing on this PC. Please Install/Re-Install MS Excel ! ", etError
    
  ElseIf Err.Number = 70 Then
    DispMsg "Cannot Access Template File. Close If Already Open and Continue.", etError
  Else
    wErrMsg = Err.Description
    DispMsg wErrMsg, etError
  End If


End Sub

Private Sub CmdImpExc_Click()
Call DispFra(ImpExc)
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
  ws_BtnStr = ""

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
  Cancel = atxt(Index).Validate
End Sub
Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
  If UnloadMode <> vbFormCode Then Cancel = True: Exit Sub
    Set moCn = Nothing    '*** (09/08/05)    '*** (Bef speed) bef 24/11/06
End Sub

Private Sub DispFra(ByVal pv_eInvFra As en_eInvFra)
  Select Case pv_eInvFra
  Case Is = ImpExc
    If FraImpExc.Visible = True Then
      FraImpExc.Visible = False
      FraImpExc.Enabled = False

    Else
      FraImpExc.Visible = True
      FraImpExc.Enabled = True
      CmdImpExcPath.SetFocus
      FraImpExc.ZOrder

    End If
  End Select
End Sub

Private Function SelectFileName(ByVal wExtn As String, ByRef ChosenFileName) As Boolean
'svv.

  Do While True
    On Error GoTo ErrH
    'CD1.ShowOpen
    Cd1.ShowSave
    If Len(Cd1.FileName) > 4 And ((UCase(Right(Cd1.FileName, 4)) <> UCase(wExtn) And UCase(Right(Cd1.FileName, 5)) <> UCase(wExtn))) Then
        MsgBox ("File Extension Is Wrong. Please re-select")
    Else
      If (CheckAndDeleteFile(Cd1.FileName)) Then
        ChosenFileName = Cd1.FileName
        SelectFileName = True
        Exit Function
      End If
    End If
  Loop
  Exit Function
ErrH:
  If Err.Number = cdlCancel Then
    SelectFileName = False
  End If
End Function
Private Function CheckAndDeleteFile(FileName As String) As Boolean
  If Dir(FileName, vbNormal) <> "" Then
    On Error GoTo errDeletingFile
    Kill FileName
    CheckAndDeleteFile = True
    Exit Function
errDeletingFile:
    MsgBox ("System is unable to delete file " + FileName + ". Please delete file or choose a different path or name")
    CheckAndDeleteFile = False
  Else
    CheckAndDeleteFile = True
  End If
End Function

Private Sub CmdImpExcPath_Click()
  Cd1.FileName = ""
  Cd1.DialogTitle = "Select Excel File "
  Cd1.Filter = "Excel Files(*.xls;*.xlsx)|*.xls;*.xlsx"
  Cd1.FilterIndex = 1
  Cd1.ShowOpen
  ADC("WXLFILE") = Cd1.FileName
  
End Sub
Private Sub CmdImpExcGo_Click()
If Trim(ADC("WXLFILE")) = "" Then DispMsg "Please Specify a Proper Excel File Name !!", etError: Exit Sub
If ADC("WXLSHTNO") = 0 Or ADC("WXLFRROW") = 0 Or ADC("WXLTOROW") = 0 Then DispMsg "Please Specify a Proper Sheet No./From Row/To Row !!", etError: Exit Sub

Call ImportExcel
Call DispFra(ImpExc)

End Sub
Private Sub ImportExcel()
'On Error GoTo ErrHdlr
    
    Dim wo_rsExcel As New MDORowSet, i As Integer
    
    Dim wRrCmCd As String, wRrCmCurCd As String, wRrSCtg As String, wQWCst As String, wQWSale As String
    Dim wRrCd As String, wRmCtg As String, wRrFrLn As Double, wRrToLn As Double
    Dim wRrCstRt As Double, wRrSalRt As Double, ws_BlnkYN As String, wRrCstRtInc As Double, wRrSalRtInc As Double




    If Not OpenExcelFile(ADC("WXLFILE")) Then DispMsg "Improper File Format", etError: Exit Sub
'new
    If IsExcelFileReadOnly() Then
      CloseExcel
      MsgBox ("File Is Already Open. Cannot Import. Close the file and try again")
      Exit Sub
    End If
    

    If ChkSheet(ADC("WXLSHTNO")) = False Then DispMsg "This WorkSheet Number Does Not Exist in the specified Excel File", etError: CloseExcel: Exit Sub
  
    
    moCn.CreateTmpTable "#TmpXls", "Select 0 as qSr, InIRN, space(20) as qInvNo Into #TmpXls From InvHd Where 1=2"
    
     Dim wInIRN As String, wInvNo As String
    ADC("WXLFRROW") = IIF(ADC("WXLFRROW") <= 1, 2, ADC("WXLFRROW"))
    For i = ADC("WXLFRROW") To ADC("WXLTOROW")
        
        wInIRN = Trim(ReadCell(ADC("WXLSHTNO"), i, 2))
                
      
        wInvNo = Trim(ReadCell(ADC("WXLSHTNO"), i, 5))
             
       moCn.Execute ("Insert Into #TmpXls Values (" + CStr(i) + ", '" + wInIRN + "', '" + wInvNo + "')")
    
    Next

    CloseExcel
    
    Set wo_rsExcel = moCn.OpenRes("Select * From #TmpXls Order By qSr")
    wo_rsExcel.MoveFirst


  If wo_rsExcel.RecCount = 0 Then DispMsg "No Record Found for Given Scope.", etInfo:  Exit Sub

Dim w_InvNo() As String
    moCn.Execute ("Disable trigger InvHd_inupdl on InvHd")
    If wo_rsExcel.RecCount > 0 Then

        Do While Not wo_rsExcel.EOF
        w_InvNo = Split(wo_rsExcel!qInvNo, "/")
        Dim wsql As String
        wsql = "Update InvHd set InIRN='" + wo_rsExcel!InIRN + "' where InCoCd='" + gs_CoCd + "' and InTc='" + w_InvNo(0) + _
                  "' and InYy='" + w_InvNo(1) + "' and InChr='" + w_InvNo(2) + "' and InNo=" + CStr(w_InvNo(3))
        moCn.Execute (wsql)
        

         wo_rsExcel.MoveNext
        Loop


    End If
    moCn.Execute ("Enable trigger InvHd_inupdl on InvHd")
DispMsg "IRN Updated Successfully !!", etInfo
  MsgBox "IRN Updated Successfully !!"
Exit Sub


End Sub

