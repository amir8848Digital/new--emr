VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Object = "{AB2CC32A-E652-42A2-B341-26ECE02F6386}#2.1#0"; "MwfCtl.ocx"
Begin VB.Form EmrFrmOrdCpy 
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
   Begin VB.Frame FRA_KEYS 
      BorderStyle     =   0  'None
      Height          =   290
      Left            =   120
      TabIndex        =   7
      Top             =   0
      Width           =   4920
      Begin MwfCtl.MWCTL_MED atxt 
         Height          =   285
         Index           =   41
         Left            =   3180
         TabIndex        =   8
         ToolTipText     =   "Enter Voucher Character"
         Top             =   0
         Width           =   570
         _ExtentX        =   1005
         _ExtentY        =   503
         MaxLength       =   3
         DataType        =   4
         DataField       =   "OmChr"
         ReCalcOn        =   "OMTC"
         IdName          =   "OMCHR"
      End
      Begin MwfCtl.MWCTL_MED atxt 
         Height          =   285
         Index           =   42
         Left            =   3840
         TabIndex        =   9
         ToolTipText     =   "Enter Voucher No."
         Top             =   0
         Width           =   750
         _ExtentX        =   1323
         _ExtentY        =   503
         Alignment       =   2
         Mask            =   "#####0"
         MaxLength       =   6
         DataType        =   1
         DataField       =   "OmNo"
         IdName          =   "OMNO"
      End
      Begin MwfCtl.MWCTL_MED atxt 
         Height          =   285
         Index           =   87
         Left            =   1275
         TabIndex        =   10
         ToolTipText     =   "Enter Voucher Company Code"
         Top             =   0
         Width           =   630
         _ExtentX        =   1111
         _ExtentY        =   503
         MaxLength       =   3
         DataType        =   4
         DataField       =   "OmCoCd"
         IdName          =   "OMCOCD"
      End
      Begin MwfCtl.MWCTL_MED atxt 
         Height          =   285
         Index           =   229
         Left            =   4770
         TabIndex        =   11
         Top             =   0
         Visible         =   0   'False
         Width           =   390
         _ExtentX        =   688
         _ExtentY        =   503
         Enabled         =   0   'False
         MaxLength       =   3
         DataType        =   4
         DataField       =   "OmPrtKey"
         IdName          =   "OMPRTKEY"
      End
      Begin MwfCtl.MWCTL_MED atxt 
         Height          =   285
         Index           =   40
         Left            =   2730
         TabIndex        =   12
         ToolTipText     =   "Enter Voucher Year"
         Top             =   0
         Width           =   330
         _ExtentX        =   582
         _ExtentY        =   503
         MaxLength       =   2
         DataType        =   4
         DataField       =   "OmYy"
         IdName          =   "OMYY"
      End
      Begin MwfCtl.MWCTL_MED atxt 
         Height          =   285
         Index           =   7
         Left            =   2020
         TabIndex        =   13
         ToolTipText     =   "Enter Voucher Type"
         Top             =   0
         Width           =   570
         _ExtentX        =   1005
         _ExtentY        =   503
         MaxLength       =   3
         DataType        =   4
         DataField       =   "OmTc"
         ReCalcParent    =   "OMCHR"
         IdName          =   "OMTC"
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
         Left            =   100
         TabIndex        =   18
         Top             =   0
         Width           =   1065
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
         Left            =   2625
         TabIndex        =   17
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
         Left            =   3075
         TabIndex        =   16
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
         Left            =   3765
         TabIndex        =   15
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
         Index           =   298
         Left            =   1920
         TabIndex        =   14
         Top             =   0
         Width           =   105
      End
   End
   Begin VB.Frame FRA_KEY2S 
      BorderStyle     =   0  'None
      Height          =   615
      Left            =   5160
      TabIndex        =   1
      Top             =   0
      Width           =   7245
   End
   Begin MwfCtl.MWCTL_BTN1 CmdExpExc 
      Height          =   420
      Left            =   6300
      TabIndex        =   6
      Top             =   9270
      Visible         =   0   'False
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
      TabIndex        =   3
      Top             =   9270
      Visible         =   0   'False
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
      TabIndex        =   2
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
      TabIndex        =   0
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
      Height          =   9030
      Left            =   0
      TabIndex        =   4
      Top             =   135
      Width           =   13875
      Begin VB.Frame FraNKeyAll 
         BorderStyle     =   0  'None
         Height          =   8265
         Left            =   120
         TabIndex        =   5
         Top             =   360
         Width           =   13410
         Begin VB.Frame Fr1 
            Height          =   8085
            Left            =   30
            TabIndex        =   20
            Top             =   120
            Width           =   13305
            Begin MwfCtl.MWCTL_MED atxt 
               Height          =   285
               Index           =   0
               Left            =   2820
               TabIndex        =   21
               ToolTipText     =   "Enter From Invoice Tc"
               Top             =   1110
               Width           =   675
               _ExtentX        =   1191
               _ExtentY        =   503
               MaxLength       =   3
               DataType        =   4
               IdName          =   "WODTCFR"
               CmpStr          =   "OdTc>="
            End
            Begin MwfCtl.MWCTL_MED atxt 
               Height          =   285
               Index           =   1
               Left            =   3600
               TabIndex        =   22
               ToolTipText     =   "Enter From Invoice Year"
               Top             =   1110
               Width           =   345
               _ExtentX        =   609
               _ExtentY        =   503
               MaxLength       =   2
               DataType        =   4
               IdName          =   "WODYYFR"
               CmpStr          =   "OdYy>="
            End
            Begin MwfCtl.MWCTL_MED atxt 
               Height          =   285
               Index           =   2
               Left            =   4050
               TabIndex        =   23
               ToolTipText     =   "Enter From Invoice Character"
               Top             =   1110
               Width           =   675
               _ExtentX        =   1191
               _ExtentY        =   503
               MaxLength       =   3
               DataType        =   4
               IdName          =   "WODCHRFR"
               CmpStr          =   "OdChr>="
            End
            Begin MwfCtl.MWCTL_MED atxt 
               Height          =   285
               Index           =   3
               Left            =   6750
               TabIndex        =   24
               ToolTipText     =   "Enter To Invoice Tc "
               Top             =   1110
               Width           =   675
               _ExtentX        =   1191
               _ExtentY        =   503
               MaxLength       =   3
               DataType        =   4
               IdName          =   "WODTCTO"
               CmpStr          =   "OdTc<="
            End
            Begin MwfCtl.MWCTL_MED atxt 
               Height          =   285
               Index           =   6
               Left            =   4830
               TabIndex        =   25
               ToolTipText     =   "Enter From Invoice No"
               Top             =   1110
               Width           =   795
               _ExtentX        =   1402
               _ExtentY        =   503
               Alignment       =   2
               Mask            =   "#####0"
               MaxLength       =   6
               DataType        =   1
               IdName          =   "WODNOFR"
               CmpStr          =   "OdNo>="
            End
            Begin MwfCtl.MWCTL_MED atxt 
               Height          =   285
               Index           =   4
               Left            =   7530
               TabIndex        =   26
               ToolTipText     =   "Enter To Invoice Year"
               Top             =   1110
               Width           =   345
               _ExtentX        =   609
               _ExtentY        =   503
               MaxLength       =   2
               DataType        =   4
               IdName          =   "WODYYTO"
               CmpStr          =   "OdYy<="
            End
            Begin MwfCtl.MWCTL_MED atxt 
               Height          =   285
               Index           =   5
               Left            =   7980
               TabIndex        =   27
               ToolTipText     =   "Enter To Invoice Character"
               Top             =   1110
               Width           =   675
               _ExtentX        =   1191
               _ExtentY        =   503
               MaxLength       =   3
               DataType        =   4
               IdName          =   "WODCHRTO"
               CmpStr          =   "OdChr<="
            End
            Begin MwfCtl.MWCTL_MED atxt 
               Height          =   285
               Index           =   8
               Left            =   8760
               TabIndex        =   28
               ToolTipText     =   "Enter To Invoice No"
               Top             =   1110
               Width           =   795
               _ExtentX        =   1402
               _ExtentY        =   503
               Alignment       =   2
               Mask            =   "#####0"
               MaxLength       =   6
               DataType        =   1
               IdName          =   "WODNOTO"
               CmpStr          =   "OdNo<="
            End
            Begin MwfCtl.MWCTL_MED atxt 
               Height          =   285
               Index           =   11
               Left            =   2820
               TabIndex        =   29
               ToolTipText     =   "Enter Company Code"
               Top             =   840
               Width           =   675
               _ExtentX        =   1191
               _ExtentY        =   503
               MaxLength       =   3
               DataType        =   4
               IdName          =   "WODCOCD"
               CmpStr          =   "OdCoCd = "
            End
            Begin MwfCtl.MWCTL_BTN1 CmdFillGrid 
               Height          =   495
               Left            =   9960
               TabIndex        =   30
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
            Begin MwfCtl.MWCTL_FLX GrdOrdDsg 
               Height          =   6285
               Left            =   120
               TabIndex        =   40
               Top             =   1680
               Width           =   12540
               _ExtentX        =   22119
               _ExtentY        =   11086
               Cols            =   5
               AllowAdd        =   0   'False
               AllowDelete     =   0   'False
               colname1        =   "WODYN"
               heading1        =   "YN  "
               datatype1       =   4
               maxlength1      =   1
               colname2        =   "WODNO"
               heading2        =   "Sr.    "
               datatype2       =   1
               mask2           =   "###0"
               maxlength2      =   4
               style2          =   0
               tooltiptext2    =   "Enter Order Design Sr. No."
               colname3        =   "WODSRCVCH"
               heading3        =   "Source Vch                                                 "
               datatype3       =   4
               maxlength3      =   30
               style3          =   2
               colname4        =   "WODDMCD"
               heading4        =   "Design Code                                                            "
               datatype4       =   4
               maxlength4      =   15
               style4          =   2
               tooltiptext4    =   "Design Suffix"
            End
            Begin MwfCtl.MWCTL_BTN1 CmdCpy 
               Height          =   495
               Left            =   12600
               TabIndex        =   41
               ToolTipText     =   "Click Here To Copy"
               Top             =   1680
               Width           =   555
               _ExtentX        =   979
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
               Index           =   1
               Left            =   6750
               TabIndex        =   43
               Top             =   480
               Width           =   240
            End
            Begin VB.Label LblMain 
               BackStyle       =   0  'Transparent
               BorderStyle     =   1  'Fixed Single
               Caption         =   "  Scope Option For Copy"
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
               Index           =   23
               Left            =   0
               TabIndex        =   42
               Top             =   120
               Width           =   13230
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
               Left            =   120
               TabIndex        =   39
               Top             =   840
               Width           =   2145
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
               Left            =   2820
               TabIndex        =   38
               Top             =   480
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
               Left            =   8685
               TabIndex        =   37
               Top             =   1110
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
               Left            =   7455
               TabIndex        =   36
               Top             =   1110
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
               Left            =   7905
               TabIndex        =   35
               Top             =   1110
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
               Left            =   4755
               TabIndex        =   34
               Top             =   1110
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
               Left            =   3975
               TabIndex        =   33
               Top             =   1110
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
               Left            =   3525
               TabIndex        =   32
               Top             =   1110
               Width           =   60
            End
            Begin VB.Label Label6 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Vch Tc/ Yy/ Chr/ No"
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
               Left            =   120
               TabIndex        =   31
               Top             =   1125
               Width           =   1815
            End
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
            Left            =   6660
            TabIndex        =   19
            Top             =   0
            Width           =   240
         End
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
Attribute VB_Name = "EmrFrmOrdCpy"
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
Dim ws_CoCd As String, ws_Tc As String, ws_Yy As String

Enum en_eInvFra
  ImpExc = 1
End Enum

Private Sub ADC_Load()
 ' *** Set the First Key Control property to the working field
  '*** Set the Previous and Next Control Properties for the Grid GrdDt

  Set moCn = ADC.Connection    '*** (Bef speed) bef 24/11/06
  FRA_KEYS.ZOrder (0)
  'Set ADC.FirKeyCtl = ADC("wOdCoCd")
    Set ADC.FirKeyCtl = ADC("OmCoCd")
    
    ADC("OmCoCd") = "UD2":  ADC("OmTc") = "SO": ADC("OmYy") = "05"
    ADC("OmChr") = "SVT":  ADC("OmNo") = 215
    
'  ws_CoCdFr = gs_CoCd: ws_CoCdTo = gs_CoCd
'  ws_TcFr = ctInTcTyp: ws_TcTo = ctInTcTyp
'  ws_YyFr = moCn.GetFldVal("Select HYy From Head Where HCd='" + ctSelfCmCd + "' And HCoCd='" + gs_CoCd + "'")
'  ws_YyTo = moCn.GetFldVal("Select HYy From Head Where HCd='" + ctSelfCmCd + "' And HCoCd='" + gs_CoCd + "'")
'  ws_ChrFr = moCn.GetFldVal("Select vPValue from vParam where vPCoCd='" + gs_CoCd + "' and vPTyp='DEFCHR' and vPMCd='" + ws_TcFr + "' and vPValue<> ''")
'  ws_ChrTo = moCn.GetFldVal("Select vPValue from vParam where vPCoCd='" + gs_CoCd + "' and vPTyp='DEFCHR' and vPMCd='" + ws_TcTo + "' and vPValue<> ''")

End Sub
Private Sub ADC_KeyWhen()
  '****** Sachin 2.14.0 - [24-11-2006] - [Multi-Company] ******
  Me.Caption = GetFrmCaption(gs_CoCd, "eInvoice")
  '****** Sachin 2.14.0 - [24-11-2006] - [Multi-Company] ******

  'FraNKeyAll.ZOrder (0)


'  ADC("wIdCoCd") = ws_CoCdFr
'  ADC("wIdCoCdTo") = ws_CoCdTo
'  ADC("wIdTcFr") = ws_TcFr
'  ADC("wIdTcTo") = ws_TcTo
'  ADC("wIdYyFr") = ws_YyFr
'  ADC("wIdYyTo") = ws_YyTo
'  ADC("wInDtFr") = IIF(ADC("wInDtFr") <> "", ADC("wInDtFr"), "01/01/80")
'  ADC("wInDtTo") = IIF(ADC("wInDtTo") <> "", ADC("wInDtTo"), "01/01/80")
'  ADC("wIdChrFr") = ws_ChrFr
'  ADC("wIdChrTo") = ws_ChrTo
  ADC("OmCoCd") = "UD2":  ADC("OmTc") = "SO": ADC("OmYy") = "05"
    ADC("OmChr") = "SVT":  ADC("OmNo") = 215
End Sub

Private Sub ADC_NonKeyWhen(pr_Cancel As Boolean, pr_ErrMsg As String)
ADC("wOdCoCd") = "UD2"
ADC("wOdTcFr") = "SO"
ADC("wOdYyFr") = "05"
ADC("wOdChrFr") = "SVT"
ADC("wOdNoFr") = 216
ADC("wOdTcTo") = "SO"
ADC("wOdYyTo") = "05"
ADC("wOdChrTo") = "SVT"
ADC("wOdNoTo") = 216

End Sub

Private Sub ADC_SetRecSource()
   ' *** Since this is an 'A' Entry, there is no RecSource associated with ADC. So set the
    'DataMode property of ADC control as xdatUnbound
  ADC.DataMode = xdatUnbound

End Sub

Private Sub ATXT_FldChange(Index As Integer, ByVal IdName As String, Cancel As Boolean, ErrMsg As String, ByVal pv_NewValue As Variant)
Select Case IdName
  Case Is = UCase("wOdTcFr")
    ADC("wOdTcTo") = pv_NewValue
  Case Is = UCase("wOdYyFr")
    ADC("wOdYyTo") = pv_NewValue
  Case Is = UCase("wOdChrFr")
    ADC("wOdChrTo") = pv_NewValue
  Case Is = UCase("wOdNoFr")
    ADC("wOdNoTo") = pv_NewValue
End Select
End Sub

Private Sub atxt_FldWhen(Index As Integer, ByVal IdName As String, Cancel As Boolean, ErrMsg As String)
Select Case IdName
  Case Is = UCase("wOdCoCd"), UCase("OmCoCd")
    Call HlpList.hCoCd
  Case Is = UCase("wOdTcFr"), UCase("wOdTcTo"), UCase("OmTc")
    Call HlpList.PMCd("TC", "'SO'")
  Case Is = UCase("wOdYyFr"), UCase("wOdYyTo"), UCase("OmYy")
    Call HlpList.PMCd("YY")
  Case Is = UCase("wOdChrFr")
    Call HlpList.vPSCd(ADC("wOdCoCd"), "CHR", ADC("wOdTcFr"), "Y")
  Case Is = UCase("wOdChrTo")
    Call HlpList.vPSCd(ADC("wOdCoCd"), "CHR", ADC("wOdTcTo"), "Y")
  Case Is = UCase("wOdNoFr")
    Call HlpList.OmNo(ADC("wOdCoCd"), ADC("wOdTcFr"), ADC("wOdYyFr"), ADC("wOdChrFr"))
  Case Is = UCase("wOdNoTo")
    Call HlpList.OmNo(ADC("wOdCoCd"), ADC("wOdTcTo"), ADC("wOdYyTo"), ADC("wOdChrTo"))
  Case Is = UCase("OmChr")
    Call HlpList.vPSCd(ADC("OmCoCd"), "CHR", ADC("OmTc"), "Y")
  Case Is = UCase("OmNo")
    Call HlpList.OmNo(ADC("OmCoCd"), ADC("OmTc"), ADC("OmYy"), ADC("OmChr"))
End Select
End Sub

Private Sub CmdCpy_Click()
  Dim wOdOmIdNo As Double, wOdOmCmCd As String, wOdOmDt As String
  Dim wOdSr As Integer, wSqlDsg As String, wSqlRm As String, wSqlLab As String
  Dim wOrOdIdNo As Double
  
  wOdOmIdNo = moCn.GetFldVal("Select OmIdNo From OrdMst Where OmCoCd='" + ADC("OmCoCd") + "' And OmTc='" + ADC("OmTc") + "' " + _
                            " And OmYy='" + ADC("OmYy") + "' And OmChr='" + ADC("OmChr") + "' And OmNo=" + CStr(ADC("OmNo")) + " And OmPrtKey='" + ctCurrPrtn + "'")
        
  wOdOmCmCd = moCn.GetFldVal("Select OmCmCd From OrdMst Where OmCoCd='" + ADC("OmCoCd") + "' And OmTc='" + ADC("OmTc") + "' " + _
                            " And OmYy='" + ADC("OmYy") + "' And OmChr='" + ADC("OmChr") + "' And OmNo=" + CStr(ADC("OmNo")) + " And OmPrtKey='" + ctCurrPrtn + "'")
                                          
  wOdOmDt = moCn.GetFldVal("Select OmDt From OrdMst Where OmCoCd='" + ADC("OmCoCd") + "' And OmTc='" + ADC("OmTc") + "' " + _
                          " And OmYy='" + ADC("OmYy") + "' And OmChr='" + ADC("OmChr") + "' And OmNo=" + CStr(ADC("OmNo")) + " And OmPrtKey='" + ctCurrPrtn + "'")
                          
  wOdSr = moCn.GetFldVal("Select max(OdSr) From OrdDsg Where OdCoCd='" + ADC("OmCoCd") + "' And OdTc='" + ADC("OmTc") + "' " + _
                          " And OdYy='" + ADC("OmYy") + "' And OdChr='" + ADC("OmChr") + "' And OdNo=" + CStr(ADC("OmNo")) + " And OdPrtKey='" + ctCurrPrtn + "'")
  
                            
  Dim wOrdVch As String, wsOrd() As String, i As Integer
  For i = 1 To GrdOrdDsg.Rows - 1
  If GrdOrdDsg.Value(i, "WODYN") = "Y" Then
    wOdSr = wOdSr + 1
    wOrdVch = GrdOrdDsg.Value(i, "WODSRCVCH")
    wsOrd = Split(wOrdVch, "/")
  
  wSqlDsg = "insert into OrdDsg " + _
            "select '" + ADC("OmTc") + "','" + ADC("OmYy") + "','" + ADC("OmChr") + "'," + CStr(ADC("OmNo")) + "," + CStr(wOdSr) + "," + _
            "OdDmCd,OdSfx,OdDmSz,OdPrdSeq,OdDelDt,OdOrdEnt,OdOrdQty,OdExpQty,OdPrdEnt,OdPrdQty,OdFgQty," + _
            "OdCalcPrc,OdSalPrc,OdCstPrc,OdDmPrdInst,OdCmPrdInst,OdCmStmpInst,OdSzInst,OdPrtCd,OdHld,OdHldDesc," + _
            "OdVaCtg,OdKt,OdMulby,OdFixPrc,OdGldAs,OdIWtEqOrd,OdIWtFrOrd,OdIGldRtEqOrd,OdCls,OdBagPcs,OdDmCol,OdGldAsWt," + _
            "ModUsr,ModDt,ModTime,OdExpDelDt,OdSfxDesc,'" + ADC("OmCoCd") + "',OdWh,OdWDiaAvlblDt,OdHDiaAvlblDt,OdLine," + _
            "OdSalRem,OdOmCtCd,OdMinWt,OdMaxWt,OdBYy,OdBChr,OdBNo,OdGmChk,OdChkTol,OdGrMet,OdGrDia,OdGrCS,OdGrFin,OdCrmFixPrcYN," + _
            "OdSubRem,OdLabAs,OdLabAsWt,'" + wOdOmCmCd + "','" + wOdOmDt + "'," + CStr(wOdOmIdNo) + ",OdDmIdNo,OdPrtKey,0,0," + _
            "OdPicNm,OdPoNo,OdILabWtFrOrd,OdChgPtrOnStwYN,InsDt,OdMinDiaTolWt,OdMaxDiaTolWt,OdBaseIdKey,OdRefIdKey,OdMrpMulby,OdMrp," + _
            "OdDmStkYy,OdDmStkChr,OdDmStkNo,OdJLRmCd,OdJLLn1,OdJLLn2,OdJLLn3,OdJLLotNo,OdJLQty,OdJLWt,OdJLExpWt,OdJLVchRt,OdJLRateByQW," + _
            "OdJLVchVal,OdInvAllBags,OdCell,OdRepGrWt,OdDmStkCoCd,OdMrpDisc,OdSubCust,OdDmTcTyp,OdFndAvlblDt " + _
            "from OrdDsg where OdCoCd ='" + wsOrd(0) + "' and OdTc='" + wsOrd(1) + "' and OdYy='" + wsOrd(2) + "' " + _
            "and OdChr='" + wsOrd(3) + "' and OdNo =" + CStr(wsOrd(4)) + " and OdSr =" + CStr(wsOrd(5))
   
   moCn.Execute (wSqlDsg)
   
   wOrOdIdNo = moCn.GetFldVal("Select OdIdNo From OrdDsg Where OdCoCd='" + ADC("OmCoCd") + "' And OdTc='" + ADC("OmTc") + "' " + _
                            " And OdYy='" + ADC("OmYy") + "' And OdChr='" + ADC("OmChr") + "' And OdNo=" + CStr(ADC("OmNo")) + _
                            " and OdSr=" + CStr(wOdSr) + " And OdPrtKey='" + ctCurrPrtn + "'")
    
   wSqlRm = "insert into OrdRm " + _
            "select '" + ADC("OmTc") + "','" + ADC("OmYy") + "','" + ADC("OmChr") + "'," + CStr(ADC("OmNo")) + "," + CStr(wOdSr) + "," + _
            "OrSrNo,OrRmCtg,OrRmSCtg,OrRmCd,OrLn1,OrLn2,OrQty,OrWt,OrSalRt,OrSalVal,OrCstRt,OrCstVal," + _
            "OrSetScd,OrSetSalRt,OrSetSalVal,OrSetCstRt,OrSetCstVal,OrWsQty,OrHsQty,OrSubShp," + _
            "OrAlyCd,OrAlySalRt,OrAlyCstRt,OrMainMet,ModUsr,ModDt,ModTime,OrRmPtr,OrLmeSal," + _
            "OrPrdQty,OrPrdWt,'" + ADC("OmCoCd") + "'," + CStr(wOrOdIdNo) + "," + _
            "OrPrtKey,OrCustRmCd,ORLN3,OrLotNo,OrOdMstIdNo " + _
            "from OrdRm where OrCoCd ='" + wsOrd(0) + "' and OrTc='" + wsOrd(1) + "' and OrYy='" + wsOrd(2) + "' " + _
            "and OrChr='" + wsOrd(3) + "' and OrNo =" + CStr(wsOrd(4)) + " and OrSr =" + CStr(wsOrd(5))
   
   moCn.Execute (wSqlRm)
   wSqlLab = "insert into OrdLab " + _
            "select '" + ADC("OmTc") + "','" + ADC("OmYy") + "','" + ADC("OmChr") + "'," + CStr(ADC("OmNo")) + "," + CStr(wOdSr) + "," + _
            "OlSrNo,OlMcd,OlScd,OlQty,OlSalRt,OlSalVal,OlCstRt,OlCstVal,OlQw,ModUsr,ModDt,ModTime,OlCstQw," + _
            "'" + ADC("OmCoCd") + "'," + CStr(wOrOdIdNo) + ",OlPrtKey " + _
            "from OrdLab where OlCoCd ='" + wsOrd(0) + "' and OlTc='" + wsOrd(1) + "' and OlYy='" + wsOrd(2) + "' " + _
            "and OlChr='" + wsOrd(3) + "' and OlNo =" + CStr(wsOrd(4)) + " and OlSr =" + CStr(wsOrd(5))
   moCn.Execute (wSqlLab)
   
   End If
   Next
   MsgBox "Records Copied"
End Sub

Private Sub CmdFillGrid_Click()
  Dim wCnd As String, wSql As String
  Dim wo_rsOrdDsg As MwfLib.MDORowSet, wi_Row As Integer

  wCnd = ADC.RepCond
  'wsql = "Select * from orddsg where odtc='so' and odyy='05' and odchr='svt' and odno=215"
  wSql = "Select * from OrdMst join OrdDsg on OmIdNo=OdOmIdNo And OmCoCd= OdCoCd and OmTc= OdTc " + _
         "and OmYy= OdYy and OmChr= OdChr and OmNo= OdNo " + IIF(gs_Partition = ctCurrPrtn, " And OmPrtKey=OdPrtKey ", "") + _
         IIF(wCnd <> "", "Where " + wCnd, "") + " Order By OdTc, OdYy, OdChr, OdNo, OdSr  "

  GrdOrdDsg.Rows = 1
  Set wo_rsOrdDsg = moCn.OpenRes(wSql)
    GrdOrdDsg.AllowAdd = True
    With GrdOrdDsg
      Do While Not (wo_rsOrdDsg.EOF Or wo_rsOrdDsg.BOF)
        .AddItem
        wi_Row = .Rows - 1
        '.Value(wi_Row, "wOdno") = Trim(wo_rsOrdDsg!OdNo)
        .Value(wi_Row, "wOdSrcVch") = Trim(wo_rsOrdDsg!OdCoCd) + "/" + Trim(wo_rsOrdDsg!OdTc) + "/" + Trim(wo_rsOrdDsg!OdYy) + "/" + Trim(wo_rsOrdDsg!OdChr) + "/" + CStr(Trim(wo_rsOrdDsg!OdNo)) + "/" + CStr(Trim(wo_rsOrdDsg!OdSr))
        .Value(wi_Row, "wOdDmCd") = Trim(wo_rsOrdDsg!OdDmCd)
        wo_rsOrdDsg.MoveNext
      Loop
    End With
  Set wo_rsOrdDsg = Nothing
End Sub

''  ***** Std Code not to be changed **********************************
Private Sub Form_GotFocus()
 ' *** Jenny Color (For Setting System Colors)
  SetMDISysColorsAfterLogin
  '*** Jenny Color (For Setting System Colors)

End Sub
Private Sub Form_Load()
  '*** Jenny Colour
  '(Bef 13/08/05)
  SetMDISysColorsAfterLogin   '*** Jenny Color (For Setting System Colors)
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

