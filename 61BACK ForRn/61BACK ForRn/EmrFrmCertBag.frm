VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Object = "{AB2CC32A-E652-42A2-B341-26ECE02F6386}#2.1#0"; "MwfCtl.ocx"
Begin VB.Form EmrFrmCertBag 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Certificate Detail - Bag"
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
      Left            =   2640
      TabIndex        =   26
      Top             =   7640
      Width           =   8235
      Begin MwfCtl.MWCTL_MED ATXT 
         Height          =   285
         Index           =   248
         Left            =   660
         TabIndex        =   27
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
         TabIndex        =   28
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
         Left            =   2340
         TabIndex        =   29
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
         Left            =   3990
         TabIndex        =   30
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
         TabIndex        =   31
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
         TabIndex        =   32
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
         TabIndex        =   37
         Top             =   120
         Width           =   5640
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
         TabIndex        =   36
         Top             =   555
         Width           =   615
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
         TabIndex        =   35
         Top             =   1065
         Width           =   765
      End
      Begin VB.Label ALBL 
         BackStyle       =   0  'Transparent
         Caption         =   "Fr Row"
         BeginProperty Font 
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
         TabIndex        =   34
         Top             =   1065
         Width           =   675
      End
      Begin VB.Label ALBL 
         BackStyle       =   0  'Transparent
         Caption         =   "To Row"
         BeginProperty Font 
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
         Left            =   3150
         TabIndex        =   33
         Top             =   1065
         Width           =   765
      End
   End
   Begin MwfCtl.MWCTL_HLP MWCTL_HLP1 
      Height          =   405
      Left            =   8910
      TabIndex        =   12
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
      TabIndex        =   14
      Top             =   24
      Width           =   14055
      Begin MwfCtl.MWCTL_MED ATXT 
         Height          =   285
         Index           =   78
         Left            =   855
         TabIndex        =   0
         ToolTipText     =   "Enter From Certificate Date"
         Top             =   0
         Width           =   885
         _ExtentX        =   1561
         _ExtentY        =   503
         MaxLength       =   8
         DataType        =   3
         IdName          =   "WCDDTFR"
         CmpStr          =   "CdDt>="
      End
      Begin MwfCtl.MWCTL_MED ATXT 
         Height          =   285
         Index           =   1
         Left            =   855
         TabIndex        =   1
         ToolTipText     =   "Enter To Certificate Date"
         Top             =   285
         Width           =   885
         _ExtentX        =   1561
         _ExtentY        =   503
         MaxLength       =   8
         DataType        =   3
         IdName          =   "WCDDTTO"
         CmpStr          =   "CdDt<="
      End
      Begin MwfCtl.MWCTL_MED ATXT 
         Height          =   285
         Index           =   28
         Left            =   8760
         TabIndex        =   10
         ToolTipText     =   "Enter From Certificate No."
         Top             =   0
         Width           =   5100
         _ExtentX        =   8996
         _ExtentY        =   503
         MaxLength       =   30
         DataType        =   4
         IdName          =   "WCDCERTNOFR"
         CmpStr          =   "CdCertNo>="
      End
      Begin MwfCtl.MWCTL_MED ATXT 
         Height          =   285
         Index           =   29
         Left            =   8760
         TabIndex        =   11
         ToolTipText     =   "Enter To Certificate No."
         Top             =   285
         Width           =   5100
         _ExtentX        =   8996
         _ExtentY        =   503
         MaxLength       =   30
         DataType        =   4
         IdName          =   "WCDCERTNOTO"
         CmpStr          =   "CdCertNo<="
      End
      Begin MwfCtl.MWCTL_MED ATXT 
         Height          =   285
         Index           =   0
         Left            =   2595
         TabIndex        =   2
         ToolTipText     =   "Enter From Company Code"
         Top             =   0
         Width           =   660
         _ExtentX        =   1164
         _ExtentY        =   503
         MaxLength       =   3
         DataType        =   4
         IdName          =   "WCDBCOCDFR"
         CmpStr          =   "CdBCoCd>="
      End
      Begin MwfCtl.MWCTL_MED ATXT 
         Height          =   285
         Index           =   2
         Left            =   2595
         TabIndex        =   3
         ToolTipText     =   "Enter To Company Code"
         Top             =   285
         Width           =   660
         _ExtentX        =   1164
         _ExtentY        =   503
         MaxLength       =   3
         DataType        =   4
         IdName          =   "WCDBCOCDTO"
         CmpStr          =   "CdBCoCd<="
      End
      Begin MwfCtl.MWCTL_MED ATXT 
         Height          =   285
         Index           =   3
         Left            =   4005
         TabIndex        =   4
         ToolTipText     =   "Enter From Bag Year"
         Top             =   -15
         Width           =   540
         _ExtentX        =   953
         _ExtentY        =   503
         MaxLength       =   2
         DataType        =   4
         IdName          =   "WCDBYYFR"
         CmpStr          =   "CdBYy>="
      End
      Begin MwfCtl.MWCTL_MED ATXT 
         Height          =   285
         Index           =   4
         Left            =   4005
         TabIndex        =   5
         ToolTipText     =   "Enter To Bag Year"
         Top             =   270
         Width           =   540
         _ExtentX        =   953
         _ExtentY        =   503
         MaxLength       =   2
         DataType        =   4
         IdName          =   "WCDBYYTO"
         CmpStr          =   "CdBYy<="
      End
      Begin MwfCtl.MWCTL_MED ATXT 
         Height          =   285
         Index           =   5
         Left            =   5355
         TabIndex        =   6
         ToolTipText     =   "Enter From Bag Character"
         Top             =   -15
         Width           =   660
         _ExtentX        =   1164
         _ExtentY        =   503
         MaxLength       =   3
         DataType        =   4
         IdName          =   "WCDBCHRFR"
         CmpStr          =   "CdBChr>="
      End
      Begin MwfCtl.MWCTL_MED ATXT 
         Height          =   285
         Index           =   6
         Left            =   5355
         TabIndex        =   7
         ToolTipText     =   "Enter To Bag Character"
         Top             =   270
         Width           =   660
         _ExtentX        =   1164
         _ExtentY        =   503
         MaxLength       =   3
         DataType        =   4
         IdName          =   "WCDBCHRTO"
         CmpStr          =   "CdBChr<="
      End
      Begin MwfCtl.MWCTL_MED ATXT 
         Height          =   285
         Index           =   7
         Left            =   6795
         TabIndex        =   8
         ToolTipText     =   "Enter From Bag Number"
         Top             =   -15
         Width           =   900
         _ExtentX        =   1588
         _ExtentY        =   503
         MaxLength       =   7
         DataType        =   4
         IdName          =   "WCDBNOFR"
         CmpStr          =   "CdBNo>="
      End
      Begin MwfCtl.MWCTL_MED ATXT 
         Height          =   285
         Index           =   8
         Left            =   6795
         TabIndex        =   9
         ToolTipText     =   "Enter To Bag Number"
         Top             =   270
         Width           =   900
         _ExtentX        =   1588
         _ExtentY        =   503
         MaxLength       =   7
         DataType        =   4
         IdName          =   "WCDBNOTO"
         CmpStr          =   "CdBNo<="
      End
      Begin VB.Label ALBL 
         BackStyle       =   0  'Transparent
         Caption         =   "BNo Fr"
         BeginProperty Font 
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
         Left            =   6075
         TabIndex        =   24
         Top             =   0
         Width           =   675
      End
      Begin VB.Label ALBL 
         BackStyle       =   0  'Transparent
         Caption         =   "BChr Fr"
         BeginProperty Font 
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
         Left            =   4605
         TabIndex        =   23
         Top             =   0
         Width           =   795
      End
      Begin VB.Label ALBL 
         BackStyle       =   0  'Transparent
         Caption         =   "BYy Fr"
         BeginProperty Font 
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
         Left            =   3315
         TabIndex        =   22
         Top             =   0
         Width           =   675
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
         Index           =   6
         Left            =   120
         TabIndex        =   21
         Top             =   285
         Width           =   735
      End
      Begin VB.Label ALBL 
         BackStyle       =   0  'Transparent
         Caption         =   "CertNo Fr"
         BeginProperty Font 
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
         Left            =   7770
         TabIndex        =   19
         Top             =   15
         Width           =   975
      End
      Begin VB.Label ALBL 
         BackStyle       =   0  'Transparent
         Caption         =   "CoCd Fr"
         BeginProperty Font 
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
         Left            =   1785
         TabIndex        =   18
         Top             =   15
         Width           =   795
      End
      Begin VB.Label ALBL 
         BackStyle       =   0  'Transparent
         Caption         =   "Date Fr"
         BeginProperty Font 
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
         TabIndex        =   17
         Top             =   0
         Width           =   735
      End
   End
   Begin VB.Frame FRA_NKEYS 
      BorderStyle     =   0  'None
      Enabled         =   0   'False
      Height          =   9300
      Left            =   -120
      TabIndex        =   13
      Top             =   0
      Width           =   14415
      Begin VB.Frame FraNKeyAll 
         BorderStyle     =   0  'None
         Height          =   9165
         Left            =   150
         TabIndex        =   15
         Top             =   90
         Width           =   14100
         Begin MwfCtl.MWCTL_FLX GrdCert 
            Height          =   8430
            Left            =   0
            TabIndex        =   20
            Top             =   600
            Width           =   13995
            _ExtentX        =   24686
            _ExtentY        =   14870
            Cols            =   10
            colname1        =   "CDDT"
            heading1        =   "Date        "
            datafld1        =   "CdDt"
            datatype1       =   3
            maxlength1      =   8
            tooltiptext1    =   "Enter Certificate Date"
            colname2        =   "CDBCOCD"
            heading2        =   "CoCd"
            datafld2        =   "CdBCoCd"
            datatype2       =   4
            maxlength2      =   3
            tooltiptext2    =   "Enter Company Code"
            colname3        =   "CDBYY"
            heading3        =   "BYy"
            datafld3        =   "CdBYy"
            datatype3       =   4
            maxlength3      =   2
            tooltiptext3    =   "Enter Bag Year"
            colname4        =   "CDBCHR"
            heading4        =   "BChr  "
            datafld4        =   "CdBChr"
            datatype4       =   4
            maxlength4      =   3
            tooltiptext4    =   "Enter Bag Character"
            colname5        =   "CDBNO"
            heading5        =   "BNo       "
            datafld5        =   "CdBNo"
            datatype5       =   1
            mask5           =   "######0"
            maxlength5      =   7
            tooltiptext5    =   "Enter Bag Number"
            colname6        =   "CDCERTNO"
            heading6        =   "Cert No.                                                                                 "
            datafld6        =   "CdCertNo"
            datatype6       =   4
            maxlength6      =   30
            tooltiptext6    =   "Enter Certificate No."
            colname7        =   "CDCMNT"
            heading7        =   "Comment                                                                             "
            datafld7        =   "CdCmnt"
            datatype7       =   4
            maxlength7      =   30
            tooltiptext7    =   "Enter Comment"
            colname8        =   "CDDOCNO"
            heading8        =   "DocNo                                                                                                                   "
            datafld8        =   "CdDocNo"
            datatype8       =   4
            maxlength8      =   45
            tooltiptext8    =   "Enter Document No."
            colname9        =   "CDREM"
            heading9        =   "Remarks                                                                                                     "
            datafld9        =   "CdRem"
            datatype9       =   4
            maxlength9      =   120
            tooltiptext9    =   "Enter Remarks"
         End
      End
   End
   Begin MwfCtl.MWCTL_ADC ADC 
      CausesValidation=   0   'False
      Height          =   555
      Left            =   0
      TabIndex        =   16
      Top             =   9360
      Width           =   14010
      _ExtentX        =   24712
      _ExtentY        =   979
      AllowAdd        =   0   'False
      AllowDelete     =   0   'False
      BtnWidth        =   885
      BtnHeight       =   435
      BtnMaskColor    =   8454143
      FontSz          =   8.25
   End
   Begin MSComDlg.CommonDialog CdImpExc 
      Left            =   2160
      Top             =   9240
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin MwfCtl.MWCTL_BTN1 CmdImpExc 
      Height          =   465
      Left            =   6240
      TabIndex        =   25
      ToolTipText     =   "Copy Option"
      Top             =   9360
      Width           =   960
      _ExtentX        =   1693
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
End
Attribute VB_Name = "EmrFrmCertBag"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim moCn As MwfLib.MDOConnection
Enum en_CertFraBag
  ImpExcBag = 0 'Import from Excel
End Enum

Private Sub ADC_KeyWhen()
  Call EnaDisaCmds(True)
  Call HideAllFras
End Sub

Private Sub ADC_Load()
  Set moCn = adc.Connection
  Set adc.FirKeyCtl = adc("WCDDTFR")
  Set adc.FirNKeyCtl = GrdCert
  adc.Child = "GrdCert"
  Set GrdCert.PrevCtl = adc
  Set GrdCert.NextCtl = adc
  GrdCert.HotKey = "A"
  
End Sub

Private Sub ADC_NonKeyWhen(pr_Cancel As Boolean, pr_ErrMsg As String)
  Call EnaDisaCmds(False)
End Sub

Private Sub GrdCert_ColChange(ByVal RowNum As Integer, ByVal ColName As String, ByVal NewValue As Variant, Cancel As Boolean, ErrMsg As String)
Select Case ColName
  Case Is = UCase("CdBCoCd")
    Cancel = (NewValue <> "") And (Not moCn.RecSeek("Select hCoCd From Head where HCd='" + ctSelfCmCd + "' And HCoCd='" + NewValue + "'"))
    If Cancel = True Then ErrMsg = "Invalid Company Code": Exit Sub
  Case Is = UCase("CdBYy")
    Cancel = (NewValue <> "") And (Not moCn.RecSeek("select PMCd from Param where PTyp= 'YY' and " + _
                         "PMCd = '" + NewValue + "'"))
    If Cancel = True Then ErrMsg = "Invalid Bag Year": Exit Sub
  Case Is = UCase("CdBChr")
    Cancel = (NewValue <> "") And (Not moCn.RecSeek("Select PMCd From Param Where PTyp= 'BCHR' And " + _
                                     "PMCd = '" + NewValue + "' "))
    If Cancel = True Then ErrMsg = "Invalid Bag Character": Exit Sub
  
  Case Is = UCase("CdBNo")
    Cancel = (NewValue <> 0) And (Not moCn.RecSeek("Select BNo From Bag Where BCoCd='" + GrdCert.Value(RowNum, "CdBCoCd") + _
                                      "' and BYy= '" + GrdCert.Value(RowNum, "CdBYy") + "' And BChr = '" + _
                                      GrdCert.Value(RowNum, "CdBChr") + "' And BNo= " + CStr(NewValue)))
    If Cancel = True Then ErrMsg = "Invalid Bag Number ": Exit Sub
  Case Is = UCase("CdDt")
    If NewValue = "" Then Cancel = True:  ErrMsg = "Invalid Date ": Exit Sub
End Select
End Sub

Private Sub GrdCert_ColWhen(ByVal RowNum As Integer, ByVal ColName As String, Cancel As Boolean, ErrMsg As String, ByVal Action As MwfCtl.en_fgAction)
Select Case ColName
  Case Is = UCase("CdBCoCd")
    Call HlpList.hCoCd
  Case Is = UCase("CdBYy")
    Call HlpList.PMCd("YY")
  Case Is = UCase("CdBChr")
    Call HlpList.PMCd("BCHR")
  Case Is = UCase("CdBNo")
    Call HlpList.BNo(GrdCert.Value(RowNum, "CdBCoCd"), GrdCert.Value(RowNum, "CdBYy"), GrdCert.Value(RowNum, "CdBChr"), "Y")
    
End Select
End Sub

Private Sub GrdCert_InitKey(ByVal RowNum As Integer)
  '*** Initialise Key Fields
  With GrdCert
    .Store "CdCtg", "B"
    .Store "CdRmCtg", ""
    .Store "CdRmLotNo", ""
    .Store "CdRmQty", 0
    .Store "CdRmCol", ""
    .Store "CdRmQlty", ""
    .Store "CdRmCd", ""
    .Store "CdRmWt", 0
    .Store "CdFMBarCode", ""
  End With
End Sub
Private Sub ADC_SetRecSource()
  '*** Since this is an 'A' Entry, there is no RecSource associated with ADC. So set the
      'DataMode property of ADC control as xdatUnbound

  adc.DataMode = xdatUnbound
End Sub

Private Sub Form_GotFocus()

  SetMDISysColorsAfterLogin

End Sub
Private Sub Form_Load()
  Dim ws_BtnStr As String
  Call SetObjColors(Me, ws_BtnStr)
  
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
Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    If UnloadMode <> vbFormCode Then Cancel = True: Exit Sub
    
    Set moCn = Nothing
End Sub

Private Sub GrdCert_RowWhen(ByVal RowNum As Integer)
'If a new row added in grid and it is not first row then values will get from the previous row for CoCd, Byy, Bchr
With GrdCert
  If .IsNew(RowNum) = True Then
    .Value(RowNum, "CdBCoCd") = .Value(RowNum - 1, "CdBCoCd")
    .Value(RowNum, "CdBYy") = .Value(RowNum - 1, "CdBYy")
    .Value(RowNum, "CdBChr") = .Value(RowNum - 1, "CdBChr")

  End If
End With

End Sub

Private Sub GrdCert_SetRecSource()
Dim ws_Cnd As String, ws_OrdBy As String
  ws_Cnd = adc.RepCond
  ws_Cnd = IIF(ws_Cnd <> "", " and " + ws_Cnd, "")
  ws_OrdBy = " Order By CdCertNo "
  GrdCert.RecSource = " Select * from CertDet where CdCtg='B' " + ws_Cnd + ws_OrdBy
End Sub

Private Sub GrdCert_Validate(Cancel As Boolean)
  Cancel = GrdCert.Validate
End Sub

Private Sub DispFra(ByVal pv_CertFra As en_CertFraBag)
  '*** This is the common procedure which is used to toggle between hiding and displaying
      'the frames (i.e. if the frame is visible then make it invisible and viseversa)
      'associated with a command button

  Select Case pv_CertFra
  Case Is = ImpExcBag
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

  End Select
End Sub
Private Sub HideAllFras()
  '*** This procedure hides all the frames other than FraNKeyAll
  
  FraNKeyAll.Visible = True
  FraImpExc.Visible = False ' Import from excel
  FraImpExc.Enabled = False ' Import from excel
  
End Sub
Private Sub EnaDisaCmds(ByVal pv_ShowFra As Boolean, Optional ByVal pv_CmdBut As MWCTL_BTN1)
  '*** This procedure is used to give the toggle option (enable and disable) to
      'the command buttons

  If Not pv_CmdBut Is Nothing Then
    If pv_ShowFra = True Then
      pv_CmdBut.CausesValidation = True

    End If
  End If
  FraNKeyAll.Enabled = Not pv_ShowFra
  adc.AllowSave = Not pv_ShowFra
  CmdImpExc.Enabled = Not pv_ShowFra
  If Not pv_CmdBut Is Nothing Then pv_CmdBut.Enabled = True
End Sub

Private Sub CmdImpExc_Click()
  Call DispFra(ImpExcBag)
  
End Sub
'Getting excel file from selected path
Private Sub CmdImpExcPath_Click()
  CdImpExc.FileName = ""
  CdImpExc.DialogTitle = "Select Certificate Bag Excel File "
  CdImpExc.Filter = "Excel Files(*.xls;*.xlsx)|*.xls;*.xlsx"
  CdImpExc.FilterIndex = 1
  CdImpExc.ShowOpen
  adc("WXLFILE") = CdImpExc.FileName

End Sub
'Checking excel file name, Sheet no., From & To row nos. If all correct then call Import excel funtion
Private Sub CmdImpExcGo_Click()
  If Trim(adc("WXLFILE")) = "" Then DispMsg "Please Specify a Proper Excel File Name !!", etError: Exit Sub
  If adc("WXLSHTNO") = 0 Or adc("WXLFRROW") = 0 Or adc("WXLTOROW") = 0 Then DispMsg "Please Specify a Proper Sheet No./From Row/To Row !!", etError: Exit Sub

  Call ImportExcel
  Call DispFra(ImpExcBag)

End Sub
'Store data from excel file to Grid
Private Sub ImportExcel()
On Error GoTo ErrHdlr

    Dim wo_rsExcel As New MDORowSet, i As Integer
    Dim wCdBCoCd As String, wCdBYy As String, wCdBChr As String, wCdBNo As Integer
    Dim wCdCertNo As String, wCdComment As String, wCdDocNo As String, wCdRemark As String, wErrMsg As String
    
    'Checks proper excel file format
    If Not OpenExcelFile(adc("WXLFILE")) Then DispMsg "Improper File Format", etError: Exit Sub

    'Checking proper sheet no.
    If ChkSheet(adc("WXLSHTNO")) = False Then DispMsg "This WorkSheet Number Does Not Exist in the specified Excel File", etError: CloseExcel: Exit Sub

    'Creating Tmp file to store data
    moCn.CreateTmpTable "#TmpXls", "Select 0 as qSr, CdDt, CdBCoCd, CdBYy, CdBChr, CdBNo, CdCertNo, CdCmnt, CdDocNo, CdRem Into #TmpXls From CertDet Where 1=2"

    adc("WXLFRROW") = IIF(adc("WXLFRROW") <= 1, 2, adc("WXLFRROW"))
    For i = adc("WXLFRROW") To adc("WXLTOROW")
        wErrMsg = ""
        'Checking Date
        If ReadCell(adc("WXLSHTNO"), i, 1) = "" Then Call SetCellVal(adc("WXLSHTNO"), i, 1, "01/01/80")
                
        'Checking Company code
        wCdBCoCd = Trim(ReadCell(adc("WXLSHTNO"), i, 2))
        If wCdBCoCd = "" Then wErrMsg = wErrMsg + IIF(wErrMsg <> "", "/", "") + "Bag Company Code cannot be blank "

        If (wCdBCoCd <> "") And (Not moCn.RecSeek("Select hCoCd From Head where  " + _
                                   "HCd='" + ctSelfCmCd + "' And HCoCd='" + wCdBCoCd + "'")) Then _
            wErrMsg = wErrMsg + IIF(wErrMsg <> "", "/", "") + "Invalid Company Code "
    
        'Checking Bag Year
        wCdBYy = Trim(ReadCell(adc("WXLSHTNO"), i, 3))
        If wCdBYy = "" Then wErrMsg = wErrMsg + IIF(wErrMsg <> "", "/", "") + "Bag Year cannot be blank "
        

        If (wCdBYy <> "") And (Not moCn.RecSeek("Select PMCd From Param where PTyp= 'YY' " + _
                               "and PMCd = '" + wCdBYy + "'")) Then _
            wErrMsg = wErrMsg + IIF(wErrMsg <> "", "/", "") + "Invalid Bag Year "
            
        'Checking Bag Character
        wCdBChr = Trim(ReadCell(adc("WXLSHTNO"), i, 4))
        If wCdBChr = "" Then wErrMsg = wErrMsg + IIF(wErrMsg <> "", "/", "") + "Bag Character cannot be blank "

        If (wCdBChr <> "") And (Not moCn.RecSeek("Select PMCd From Param where  " + _
                              "PTyp= 'BCHR' and PMCd = '" + wCdBChr + "'")) Then _
            wErrMsg = wErrMsg + IIF(wErrMsg <> "", "/", "") + "Invalid Bag Character "
            
        'Checking valid Bag number in the Excel File
        If ReadCell(adc("WXLSHTNO"), i, 5) <> "" And Not IsNumeric(ReadCell(adc("WXLSHTNO"), i, 5)) Then
            wErrMsg = wErrMsg + IIF(wErrMsg <> "", "/", "") + "Bag Number is not Number "
        End If

        wCdBNo = Val(ReadCell(adc("WXLSHTNO"), i, 5))
        If (wCdBNo > 0) And (Not moCn.RecSeek("Select BNo From Bag where BCoCd='" + wCdBCoCd + "' and BYy= '" + wCdBYy + _
                              "' And BChr = '" + wCdBChr + "' And BNo= " + CStr(wCdBNo) + "")) Then _
            wErrMsg = wErrMsg + IIF(wErrMsg <> "", "/", "") + "Invalid Bag Number "
                    
        
        wCdCertNo = Trim(ReadCell(adc("WXLSHTNO"), i, 6))
        If wCdCertNo = "" Then wErrMsg = wErrMsg + IIF(wErrMsg <> "", "/", "") + "Certificate Number cannot be blank "
        
        wCdComment = Trim(ReadCell(adc("WXLSHTNO"), i, 7))
        wCdDocNo = Trim(ReadCell(adc("WXLSHTNO"), i, 8))
        wCdRemark = Trim(ReadCell(adc("WXLSHTNO"), i, 9))
        

       If wErrMsg <> "" Then Call SetCellVal(adc("WXLSHTNO"), i, 11, wErrMsg)
       If wErrMsg = "" Then _
        moCn.Execute ("Insert Into #TmpXls Values (" + CStr(i) + ",'" + ReadCell(adc("WXLSHTNO"), i, 1) + "', '" + wCdBCoCd + "', '" + wCdBYy + "','" + _
                        wCdBChr + "', " + Str(wCdBNo) + ", '" + wCdCertNo + "', '" + wCdComment + "' , '" + wCdDocNo + "' , '" + wCdRemark + "' " + ")")
        
    Next
  

    Set wo_rsExcel = moCn.OpenRes("Select * From #TmpXls Order By qSr")
    CloseExcel
    wo_rsExcel.MoveFirst

    Dim wRow As Double

    If wo_rsExcel.RecCount > 0 Then
      With GrdCert
        .StartCopy
        MWLib.BeginProcess Me, "Copying Details ..."
        Do While Not wo_rsExcel.EOF
          .AddItem
          wRow = .Rows - 1
          .Row = wRow

          .Value(wRow, "CDDT") = wo_rsExcel!CdDt
          .Value(wRow, "CDBCOCD") = wo_rsExcel!CdBCoCd
          .Value(wRow, "CDBYY") = wo_rsExcel!CdBYy
          .Value(wRow, "CDBCHR") = wo_rsExcel!CdBChr
          .Value(wRow, "CDBNO") = wo_rsExcel!CdBNo
          .Value(wRow, "CDCERTNO") = wo_rsExcel!CDCERTNO
          .Value(wRow, "CDCMNT") = wo_rsExcel!CdCmnt
          .Value(wRow, "CDDOCNO") = wo_rsExcel!CdDocNo
          .Value(wRow, "CDREM") = wo_rsExcel!CdRem

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

