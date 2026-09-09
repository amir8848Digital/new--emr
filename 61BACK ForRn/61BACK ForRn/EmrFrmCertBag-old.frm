VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Object = "{AB2CC32A-E652-42A2-B341-26ECE02F6386}#2.1#0"; "MwfCtl.ocx"
Begin VB.Form EmrFrmCertBag 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Certificate Details – Bag"
   ClientHeight    =   3090
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   4680
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   3090
   ScaleWidth      =   4680
   WindowState     =   2  'Maximized
   Begin VB.Frame FRA_KEYS 
      BorderStyle     =   0  'None
      Height          =   930
      Left            =   84
      TabIndex        =   1
      Top             =   24
      Width           =   14775
      Begin MwfCtl.MWCTL_MED atxt 
         Height          =   285
         Index           =   78
         Left            =   660
         TabIndex        =   2
         ToolTipText     =   "Enter From Certificate Date"
         Top             =   360
         Width           =   885
         _ExtentX        =   1561
         _ExtentY        =   503
         MaxLength       =   8
         DataType        =   3
         IdName          =   "WCDDTFR"
         CmpStr          =   "CdDt>="
      End
      Begin MwfCtl.MWCTL_MED atxt 
         Height          =   285
         Index           =   8
         Left            =   1575
         TabIndex        =   3
         ToolTipText     =   "Enter From Rm Lot"
         Top             =   360
         Width           =   2400
         _ExtentX        =   4233
         _ExtentY        =   503
         MaxLength       =   16
         DataType        =   4
         IdName          =   "WCDRMLOTNOFR"
         CmpStr          =   "CdRmLotNo>="
      End
      Begin MwfCtl.MWCTL_MED atxt 
         Height          =   285
         Index           =   9
         Left            =   1575
         TabIndex        =   4
         ToolTipText     =   "Enter To Rm Lot"
         Top             =   645
         Width           =   2400
         _ExtentX        =   4233
         _ExtentY        =   503
         MaxLength       =   16
         DataType        =   4
         IdName          =   "WCDRMLOTNOTO"
         CmpStr          =   "CdRmLotNo<="
      End
      Begin MwfCtl.MWCTL_MED atxt 
         Height          =   285
         Index           =   28
         Left            =   3975
         TabIndex        =   5
         ToolTipText     =   "Enter From Certificate No."
         Top             =   360
         Width           =   4380
         _ExtentX        =   7726
         _ExtentY        =   503
         MaxLength       =   30
         DataType        =   4
         IdName          =   "WCDCERTNOFR"
         CmpStr          =   "CdCertNo>="
      End
      Begin MwfCtl.MWCTL_MED atxt 
         Height          =   285
         Index           =   29
         Left            =   3975
         TabIndex        =   6
         ToolTipText     =   "Enter To Certificate No."
         Top             =   645
         Width           =   4380
         _ExtentX        =   7726
         _ExtentY        =   503
         MaxLength       =   30
         DataType        =   4
         IdName          =   "WCDCERTNOTO"
         CmpStr          =   "CdCertNo<="
      End
      Begin MwfCtl.MWCTL_MED atxt 
         Height          =   285
         Index           =   1
         Left            =   660
         TabIndex        =   7
         ToolTipText     =   "Enter To Certificate Date"
         Top             =   645
         Width           =   885
         _ExtentX        =   1561
         _ExtentY        =   503
         MaxLength       =   8
         DataType        =   3
         IdName          =   "WCDDTTO"
         CmpStr          =   "CdDt<="
      End
      Begin MwfCtl.MWCTL_MED atxt 
         Height          =   285
         Index           =   0
         Left            =   8355
         TabIndex        =   8
         ToolTipText     =   "Enter From Rm Code"
         Top             =   360
         Width           =   1920
         _ExtentX        =   3387
         _ExtentY        =   503
         MaxLength       =   12
         DataType        =   4
         IdName          =   "WCDRMCDFR"
         CmpStr          =   "CdRmCd>="
      End
      Begin MwfCtl.MWCTL_MED atxt 
         Height          =   285
         Index           =   2
         Left            =   8355
         TabIndex        =   9
         ToolTipText     =   "Enter To Rm Code"
         Top             =   645
         Width           =   1920
         _ExtentX        =   3387
         _ExtentY        =   503
         MaxLength       =   12
         DataType        =   4
         IdName          =   "WCDRMCDTO"
         CmpStr          =   "CdRmCd<="
      End
      Begin MwfCtl.MWCTL_MED atxt 
         Height          =   285
         Index           =   3
         Left            =   10280
         TabIndex        =   10
         ToolTipText     =   "Enter From Rm Quality"
         Top             =   360
         Width           =   4380
         _ExtentX        =   7726
         _ExtentY        =   503
         MaxLength       =   30
         DataType        =   4
         IdName          =   "WCDRMQLTYFR"
         CmpStr          =   "CdRmQlty>="
      End
      Begin MwfCtl.MWCTL_MED atxt 
         Height          =   285
         Index           =   4
         Left            =   10280
         TabIndex        =   11
         ToolTipText     =   "Enter To Rm Quality"
         Top             =   645
         Width           =   4380
         _ExtentX        =   7726
         _ExtentY        =   503
         MaxLength       =   30
         DataType        =   4
         IdName          =   "WCDRMQLTYTO"
         CmpStr          =   "CdRmQlty<="
      End
      Begin VB.Label ALBL 
         BackStyle       =   0  'Transparent
         Caption         =   "Rm Quality "
         BeginProperty Font 
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
         Left            =   10365
         TabIndex        =   18
         Top             =   15
         Width           =   1185
      End
      Begin VB.Label ALBL 
         BackStyle       =   0  'Transparent
         Caption         =   "Date "
         BeginProperty Font 
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
         Left            =   720
         TabIndex        =   17
         Top             =   0
         Width           =   735
      End
      Begin VB.Label ALBL 
         BackStyle       =   0  'Transparent
         Caption         =   "Rm Code "
         BeginProperty Font 
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
         Left            =   8445
         TabIndex        =   16
         Top             =   15
         Width           =   1065
      End
      Begin VB.Label ALBL 
         BackStyle       =   0  'Transparent
         Caption         =   "Rm Lot "
         BeginProperty Font 
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
         Left            =   1635
         TabIndex        =   15
         Top             =   15
         Width           =   915
      End
      Begin VB.Label ALBL 
         BackStyle       =   0  'Transparent
         Caption         =   "Certificate No"
         BeginProperty Font 
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
         Left            =   4050
         TabIndex        =   14
         Top             =   15
         Width           =   2055
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
         Index           =   1
         Left            =   120
         TabIndex        =   13
         Top             =   360
         Width           =   615
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
         TabIndex        =   12
         Top             =   645
         Width           =   735
      End
   End
   Begin VB.Frame FRA_NKEYS 
      BorderStyle     =   0  'None
      Enabled         =   0   'False
      Height          =   9300
      Left            =   0
      TabIndex        =   19
      Top             =   0
      Width           =   14775
      Begin VB.Frame FraNKeyAll 
         BorderStyle     =   0  'None
         Height          =   9165
         Left            =   150
         TabIndex        =   20
         Top             =   90
         Width           =   14580
         Begin MwfCtl.MWCTL_FLX GrdCert 
            Height          =   8190
            Left            =   0
            TabIndex        =   21
            Top             =   960
            Width           =   14595
            _ExtentX        =   25744
            _ExtentY        =   14446
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
            heading4        =   "BChr   "
            datafld4        =   "CdBChr"
            datatype4       =   4
            maxlength4      =   3
            tooltiptext4    =   "Enter bag Character"
            colname5        =   "CDBNO"
            heading5        =   "BNo        "
            datafld5        =   "CdBNo"
            datatype5       =   1
            mask5           =   "######0"
            maxlength5      =   7
            tooltiptext5    =   "Enter Bag No."
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
   Begin MwfCtl.MWCTL_HLP MWCTL_HLP1 
      Height          =   405
      Left            =   9030
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   9360
      Visible         =   0   'False
      Width           =   555
      _ExtentX        =   979
      _ExtentY        =   714
   End
   Begin MwfCtl.MWCTL_ADC ADC 
      CausesValidation=   0   'False
      Height          =   555
      Left            =   120
      TabIndex        =   22
      Top             =   9360
      Width           =   14370
      _ExtentX        =   25347
      _ExtentY        =   979
      AllowAdd        =   0   'False
      AllowDelete     =   0   'False
      BtnWidth        =   885
      BtnHeight       =   435
      BtnMaskColor    =   8454143
      FontSz          =   8.25
   End
   Begin MSComDlg.CommonDialog CdImpExc 
      Left            =   2280
      Top             =   9240
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
End
Attribute VB_Name = "EmrFrmCertBag"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim moCn As MwfLib.MDOConnection

Private Sub ADC_Load()
  Set moCn = ADC.Connection
  Set ADC.FirKeyCtl = ADC("WCDDTFR")
  Set ADC.FirNKeyCtl = GrdCert
  ADC.Child = "GrdCert"
  Set GrdCert.PrevCtl = ADC
  Set GrdCert.NextCtl = ADC
  GrdCert.HotKey = "A"
  
End Sub

Private Sub GrdCert_ColChange(ByVal RowNum As Integer, ByVal ColName As String, ByVal NewValue As Variant, Cancel As Boolean, ErrMsg As String)
Select Case ColName
  Case Is = UCase("CdRmCtg")
    If NewValue <> "D" Then Cancel = True: ErrMsg = "Invalid Value. Rm Category Should Be 'D'": Exit Sub
   
  Case Is = UCase("CdRmLotNo")
    Cancel = NewValue <> "" And (Not moCn.RecSeek("select 'x' from SYyMm where SCoCd= '" + gs_CoCd + "' and " + _
                                  IIF(GrdCert.Value(RowNum, "CdRmCd") <> "", "SRmCd='" + GrdCert.Value(RowNum, "CdRmCd") + "' and ", "") + _
                                  "SLotNo= '" + NewValue + "'"))
    ErrMsg = "Invalid Lot No.": Exit Sub
      
  Case Is = UCase("CdRmCol")
    Cancel = (NewValue <> "") And (Not moCn.RecSeek("Select PSCd From Param where PTyp ='RMCOL' " + _
                                    "and PMCd='" + GrdCert.Value(RowNum, "CdRmCtg") + "' " + _
                                    "and PSCd = '" + NewValue + "'"))
    If Cancel = True Then ErrMsg = "Invalid Rm Color": Exit Sub
  Case Is = UCase("CDRMQLTY")
    If NewValue = "" Then Cancel = True: ErrMsg = "Rm Quality Should Not Be Empty ": Exit Sub
   
  Case Is = UCase("CdRmCd")
    Cancel = (NewValue <> "") And (Not moCn.RecSeek("Select RmCd From RmMst where RmZ='N' and " + _
                                    "RmCtg='" + GrdCert.Value(RowNum, "CdRmCtg") + "' " + _
                                    "and RmCd = '" + NewValue + "'"))
    If Cancel = True Then ErrMsg = "Invalid Rm Code": Exit Sub
End Select
End Sub

Private Sub GrdCert_ColWhen(ByVal RowNum As Integer, ByVal ColName As String, Cancel As Boolean, ErrMsg As String, ByVal Action As MwfCtl.en_fgAction)
Select Case ColName
  Case Is = UCase("CdRmCtg")
    Call HlpList.PMCd("RMCTG")
  Case Is = UCase("CdRmLotNo")
    Call HlpList.LotNo(gs_CoCd, , GrdCert.Value(RowNum, "CdRmCd"))
  Case Is = UCase("CdRmCol")
    Call HlpList.PSCd("RMCOL", GrdCert.Value(RowNum, "CdRmCtg"))
  Case Is = UCase("CdRmCd")
    Call HlpList.RmCd("N", "'" + GrdCert.Value(RowNum, "CdRmCtg") + "'")
End Select
End Sub

Private Sub GrdCert_InitKey(ByVal RowNum As Integer)
  '*** Initialise Key Fields
  With GrdCert
    .Store "CdCtg", "L"
    .Store "CdBCoCd", ""
    .Store "CdBYy", ""
    .Store "CdBChr", ""
    .Store "CdBNo", 0

  End With
End Sub
Private Sub ADC_SetRecSource()
  '*** Since this is an 'A' Entry, there is no RecSource associated with ADC. So set the
      'DataMode property of ADC control as xdatUnbound

  ADC.DataMode = xdatUnbound
End Sub

Private Sub Form_GotFocus()

  SetMDISysColorsAfterLogin

End Sub
Private Sub Form_Load()
  Dim ws_BtnStr As String
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
Private Sub ATXT_Validate(Index As Integer, Cancel As Boolean)
  Cancel = atxt(Index).Validate
End Sub
Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    If UnloadMode <> vbFormCode Then Cancel = True: Exit Sub
    
    Set moCn = Nothing
End Sub

Private Sub GrdCert_RowWhen(ByVal RowNum As Integer)
If GrdCert.IsNew(RowNum) = True Then
  GrdCert.Value(RowNum, "CDRMCTG") = "D"
  GrdCert.Value(RowNum, "CDRMQTY") = 1
  
End If
End Sub

Private Sub GrdCert_SetRecSource()
Dim ws_Cnd As String, ws_OrdBy As String
  ws_Cnd = ADC.RepCond
  ws_Cnd = IIF(ws_Cnd <> "", " where " + ws_Cnd, "")
  ws_OrdBy = " Order By CdCertNo "
  GrdCert.RecSource = " Select * from CertDet " + ws_Cnd + ws_OrdBy
End Sub

Private Sub GrdCert_Validate(Cancel As Boolean)
  Cancel = GrdCert.Validate
End Sub



