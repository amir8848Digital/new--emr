VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Object = "{6949E11D-3FEA-49A2-8D98-05EF092A8A24}#1.0#0"; "mwfctl.ocx"
Begin VB.Form EmrFrmTrfIn 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "XML To Emr Transfer"
   ClientHeight    =   8475
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   12420
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   8475
   ScaleWidth      =   12420
   Begin VB.CommandButton CmdSO 
      Caption         =   "<&a> Sales Order"
      Height          =   600
      Left            =   30
      Style           =   1  'Graphical
      TabIndex        =   26
      ToolTipText     =   "Option To Transfer Customer Master"
      Top             =   0
      Width           =   2400
   End
   Begin VB.CommandButton CmdClearErrLog 
      Caption         =   "clea<&R> log"
      Height          =   405
      Left            =   10230
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   7980
      Width           =   1155
   End
   Begin VB.CommandButton CmdShowErrLog 
      Caption         =   "<&S>how log"
      Height          =   405
      Left            =   10230
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   7575
      Width           =   1155
   End
   Begin MSComDlg.CommonDialog ComDialogOpn 
      Left            =   10110
      Top             =   1530
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.Frame FRA_NKEYS 
      BorderStyle     =   0  'None
      Enabled         =   0   'False
      Height          =   300
      Left            =   10725
      TabIndex        =   10
      Top             =   1080
      Width           =   1125
      Begin VB.Frame FraNKeyAll 
         BorderStyle     =   0  'None
         Height          =   3705
         Left            =   6030
         TabIndex        =   11
         Top             =   1530
         Width           =   5790
      End
   End
   Begin MwfCtl.MWCTL_HLP MWCTL_HLP1 
      Height          =   405
      Left            =   10035
      TabIndex        =   2
      TabStop         =   0   'False
      Top             =   990
      Visible         =   0   'False
      Width           =   555
      _ExtentX        =   979
      _ExtentY        =   714
   End
   Begin MwfCtl.MWCTL_ADC ADC 
      CausesValidation=   0   'False
      Height          =   555
      Left            =   11460
      TabIndex        =   3
      Top             =   7950
      Width           =   2085
      _ExtentX        =   3678
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
      Height          =   810
      Index           =   8
      Left            =   30
      TabIndex        =   13
      ToolTipText     =   "Message"
      Top             =   7575
      Width           =   10140
      _ExtentX        =   17886
      _ExtentY        =   1429
      Alignment       =   3
      Enabled         =   0   'False
      MaxLength       =   255
      DataType        =   4
      IdName          =   "WDISP"
   End
   Begin VB.Frame FRA_KEYS 
      BorderStyle     =   0  'None
      Height          =   450
      Left            =   10770
      TabIndex        =   12
      Top             =   1560
      Width           =   825
   End
   Begin VB.Frame FraSO 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1680
      Left            =   30
      TabIndex        =   14
      Top             =   5805
      Visible         =   0   'False
      Width           =   12300
      Begin VB.CommandButton CmdSOTrfIn 
         Caption         =   "&Transfer In"
         Height          =   315
         Left            =   10920
         TabIndex        =   9
         ToolTipText     =   "Transfer In"
         Top             =   900
         Width           =   1170
      End
      Begin MwfCtl.MWCTL_MED ATXT 
         Height          =   570
         Index           =   34
         Left            =   1410
         TabIndex        =   8
         ToolTipText     =   "Enter File Path"
         Top             =   900
         Width           =   9345
         _ExtentX        =   16484
         _ExtentY        =   1005
         Alignment       =   3
         Enabled         =   0   'False
         MaxLength       =   0
         IdName          =   "WFILEPATH"
      End
      Begin MwfCtl.MWCTL_MED ATXT 
         Height          =   285
         Index           =   19
         Left            =   1410
         TabIndex        =   7
         ToolTipText     =   "Enter Target Supplier Code"
         Top             =   615
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   503
         MaxLength       =   8
         DataType        =   4
         IdName          =   "WSOCMCDTGT"
      End
      Begin MwfCtl.MWCTL_MED ATXT 
         Height          =   285
         Index           =   5
         Left            =   2640
         TabIndex        =   6
         ToolTipText     =   "Enter Target Voucher Tc"
         Top             =   330
         Width           =   570
         _ExtentX        =   1005
         _ExtentY        =   503
         MaxLength       =   3
         DataType        =   4
         ReCalcOn        =   "WSOVCHTCTGT"
         IdName          =   "WSOVCHCHRTGT"
      End
      Begin MwfCtl.MWCTL_MED ATXT 
         Height          =   285
         Index           =   6
         Left            =   2100
         TabIndex        =   5
         ToolTipText     =   "Enter Target Voucher Year"
         Top             =   330
         Width           =   435
         _ExtentX        =   767
         _ExtentY        =   503
         MaxLength       =   2
         DataType        =   4
         IdName          =   "WSOVCHYYTGT"
      End
      Begin MwfCtl.MWCTL_MED ATXT 
         Height          =   285
         Index           =   117
         Left            =   1410
         TabIndex        =   4
         ToolTipText     =   "Enter Target Voucher Tc"
         Top             =   330
         Width           =   570
         _ExtentX        =   1005
         _ExtentY        =   503
         MaxLength       =   3
         DataType        =   4
         ReCalcParent    =   "WSOVCHCHRTGT"
         IdName          =   "WSOVCHTCTGT"
      End
      Begin VB.Label LblMain 
         BackStyle       =   0  'Transparent
         BorderStyle     =   1  'Fixed Single
         Caption         =   "  Sales Order Transfer In"
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
         TabIndex        =   32
         Top             =   -30
         Width           =   12300
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
         Index           =   3
         Left            =   120
         TabIndex        =   31
         Top             =   330
         Width           =   1065
      End
      Begin VB.Label ALBL 
         BackStyle       =   0  'Transparent
         Caption         =   "File Path"
         BeginProperty Font 
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
         Left            =   120
         TabIndex        =   30
         Top             =   900
         Width           =   1065
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
         Index           =   9
         Left            =   120
         TabIndex        =   29
         Top             =   615
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
         Index           =   14
         Left            =   2010
         TabIndex        =   28
         Top             =   330
         Width           =   135
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
         Left            =   2550
         TabIndex        =   27
         Top             =   330
         Width           =   135
      End
   End
   Begin VB.Frame FraErrLog 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   5355
      Left            =   30
      TabIndex        =   15
      Top             =   2130
      Width           =   12315
      Begin MwfCtl.MWCTL_BTN1 CmdErrLogGo 
         Height          =   285
         Left            =   4335
         TabIndex        =   25
         ToolTipText     =   "Press This Button To Copy"
         Top             =   4830
         Width           =   585
         _ExtentX        =   0
         _ExtentY        =   0
         BackColor       =   -2147483633
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
      Begin MwfCtl.MWCTL_FLX GrdErrLog 
         Height          =   3840
         Left            =   90
         TabIndex        =   20
         Top             =   330
         Width           =   12105
         _ExtentX        =   21352
         _ExtentY        =   6773
         Cols            =   5
         colname1        =   "WSRNO"
         heading1        =   "Sr."
         datatype1       =   1
         mask1           =   "###0"
         maxlength1      =   4
         style1          =   2
         colname2        =   "WELTABLE"
         heading2        =   "Table                     "
         maxlength2      =   30
         style2          =   2
         colname3        =   "WELKEY"
         heading3        =   "Key                                     "
         maxlength3      =   30
         style3          =   2
         colname4        =   "WELERRMSG"
         heading4        =   $"EmrFrmTrfIn.frx":0000
         datatype4       =   4
         maxlength4      =   255
         style4          =   2
      End
      Begin MwfCtl.MWCTL_MED ATXT 
         Height          =   285
         Index           =   0
         Left            =   1305
         TabIndex        =   16
         ToolTipText     =   "Enter Table From For Scope"
         Top             =   4230
         Width           =   1095
         _ExtentX        =   1931
         _ExtentY        =   503
         MaxLength       =   0
         IdName          =   "WTBLFR"
      End
      Begin MwfCtl.MWCTL_MED ATXT 
         Height          =   285
         Index           =   1
         Left            =   1305
         TabIndex        =   17
         ToolTipText     =   "Enter Table From To Scope"
         Top             =   4515
         Width           =   1095
         _ExtentX        =   1931
         _ExtentY        =   503
         MaxLength       =   0
         IdName          =   "WTBLTO"
      End
      Begin MwfCtl.MWCTL_MED ATXT 
         Height          =   285
         Index           =   2
         Left            =   3825
         TabIndex        =   18
         ToolTipText     =   "Enter DdMm For Scope"
         Top             =   4230
         Width           =   1095
         _ExtentX        =   1931
         _ExtentY        =   503
         MaxLength       =   0
         IdName          =   "WDDMMFR"
      End
      Begin MwfCtl.MWCTL_MED ATXT 
         Height          =   285
         Index           =   3
         Left            =   3825
         TabIndex        =   19
         ToolTipText     =   "Enter DdMm To Scope"
         Top             =   4515
         Width           =   1095
         _ExtentX        =   1931
         _ExtentY        =   503
         MaxLength       =   0
         IdName          =   "WDDMMTO"
      End
      Begin VB.Label LblMain 
         BackStyle       =   0  'Transparent
         BorderStyle     =   1  'Fixed Single
         Caption         =   "  Error Log"
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
         TabIndex        =   33
         Top             =   -30
         Width           =   12315
      End
      Begin VB.Label ALBL 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00C0C0C0&
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
         ForeColor       =   &H00000000&
         Height          =   285
         Index           =   2
         Left            =   2610
         TabIndex        =   24
         Top             =   4515
         Width           =   1095
      End
      Begin VB.Label ALBL 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00C0C0C0&
         BackStyle       =   0  'Transparent
         Caption         =   "DdMm Fr"
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
         Index           =   1
         Left            =   2610
         TabIndex        =   23
         Top             =   4230
         Width           =   1095
      End
      Begin VB.Label ALBL 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00C0C0C0&
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
         ForeColor       =   &H00000000&
         Height          =   285
         Index           =   0
         Left            =   90
         TabIndex        =   22
         Top             =   4515
         Width           =   1095
      End
      Begin VB.Label ALBL 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00C0C0C0&
         BackStyle       =   0  'Transparent
         Caption         =   "Table Fr"
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
         Index           =   5
         Left            =   90
         TabIndex        =   21
         Top             =   4230
         Width           =   1095
      End
   End
   Begin VB.Line Line1 
      Visible         =   0   'False
      X1              =   5940
      X2              =   5940
      Y1              =   1860
      Y2              =   2640
   End
End
Attribute VB_Name = "EmrFrmTrfIn"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'*****************************************************************************************
'*** Data Porting From XML To Emr
'*** xEmr Ver 2.06 To Emr Ver 2.06, 2.07, 2.08, 2.09, 2.10, 2.11
'*****************************************************************************************

Option Explicit
Dim moCn As MwfLib.MDOConnection

'*** Declarations For The Target RecordSets (Add)
Dim Rs_TgtAddX As MwfLib.MDORowSet
Dim Rs_TgtAddA As MwfLib.MDORowSet
Dim Rs_TgtAddB As MwfLib.MDORowSet
Dim Rs_TgtAddC As MwfLib.MDORowSet
Dim Rs_TgtAddB1 As MwfLib.MDORowSet
Dim Rs_TgtAddC1 As MwfLib.MDORowSet
Dim Rs_TgtAddD As MwfLib.MDORowSet
Dim Rs_TgtAddG As MwfLib.MDORowSet

'*** Declarations For The Target RecordSets (Update)
Dim Rs_TgtUpdX As MwfLib.MDORowSet
Dim Rs_TgtUpdA As MwfLib.MDORowSet
Dim Rs_TgtUpdB As MwfLib.MDORowSet
Dim Rs_TgtUpdC As MwfLib.MDORowSet
Dim Rs_TgtUpdB1 As MwfLib.MDORowSet
Dim Rs_TgtUpdC1 As MwfLib.MDORowSet
Dim Rs_TgtUpdD As MwfLib.MDORowSet
Dim Rs_TgtUpdG As MwfLib.MDORowSet

'*** Declarations For The XML Document
Dim XmlDoc As MSXML2.DOMDocument
Dim xmlRootElement As MSXML2.IXMLDOMElement
Dim xmlChildX As MSXML2.IXMLDOMElement
Dim xmlChildLevel1 As MSXML2.IXMLDOMElement
Dim xmlChildLevel2 As MSXML2.IXMLDOMElement

'*** Arr_Cv stores the causes validation state, Arr_Ed the Ena/Disa state of Cmd Buttons
Dim mArr_Cv As Variant, mArr_Ed As Variant

Dim mf_i As Single, mf_j As Single, mf_k As Single
Dim mi_i As Integer
Dim ms_FilePth As String
Dim ms_Opt As String
Dim mf_Cancel As Boolean, ms_ErrMsg As String
Dim mf_SrcCurCnvRt As Single

Dim ms_ErrTbl As String
Dim xmlErrRecSet As MSXML2.IXMLDOMElement

'*** Var Declaration for Tc/Yy/Chr/No Flds
Dim ms_OmTc As String, ms_OmYy As String, ms_OmChr As String, mf_OmNo As Single

Private Enum en_EmrTrfInDispFra
  SO = 0
  ErrLog = 1
End Enum
Private Sub ADC_Load()
  ReDim mArr_Cv(2), mArr_Ed(2)
  Set moCn = ADC.Connection
  Set ADC.FirKeyCtl = ADC
  Call HideAllFras
  GrdErrLog.AllowAdd = False
  GrdErrLog.AllowDelete = False
End Sub
Private Sub ATXT_FldWhen(Index As Integer, ByVal IdName As String, Cancel As Boolean, ErrMsg As String)
  Select Case UCase(IdName)
  Case Is = UCase("wSOVchTcTgt")
    HlpList.PMCd "TC", "'SO'"
  Case Is = UCase("wSOVchYyTgt")
    Call HlpList.PMCd("YY")
  Case Is = UCase("wSOVchChrTgt")
    Call HlpList.vPSCd(gs_CoCd, "CHR", ADC("wSOVchTcTgt"))
  Case Is = UCase("wTblFr"), UCase("wTblTo")
    Call HlpList.Tables
  Case Is = UCase("wSOCmCdTgt")
    Call HlpList.CustCd("C")
  End Select
End Sub
Private Sub atxt_FldChange(Index As Integer, ByVal IdName As String, Cancel As Boolean, ErrMsg As String, ByVal pv_NewValue As Variant)
  Select Case UCase(IdName)
  Case Is = UCase("wSOVchTcTgt")
    Cancel = (pv_NewValue <> "") And (Not moCn.RecSeek("select PMCd from Param where PTyp= 'TC' and " + _
                         "PMCd = '" + pv_NewValue + "' and PValue= 'SO'"))
    If Cancel = True Then ErrMsg = "Invalid Voucher Tc": Exit Sub
  Case Is = UCase("wSOVchYyTgt")
    Cancel = (pv_NewValue <> "") And (Not moCn.RecSeek("select PMCd from Param where PTyp= 'YY' and " + _
                         "PMCd = '" + pv_NewValue + "'"))
    If Cancel = True Then ErrMsg = "Invalid Voucher Year": Exit Sub
  Case Is = UCase("wSOVchChrTgt")
    Cancel = pv_NewValue <> "" And (Not moCn.RecSeek("select vPSCd from vParam where " + _
                         "vPCoCd= '" + gs_CoCd + "' and vPTyp= 'CHR' and " + _
                         "vPMCd = '" + ADC("wSOVchTcTgt") + "' and vPSCd='" + pv_NewValue + "'"))
    ErrMsg = "Invalid Voucher Character": Exit Sub
  Case Is = UCase("wDdMmFr"), UCase("wDdMmTo")
    Cancel = (pv_NewValue <> "" And _
                (Not IsNumeric(pv_NewValue) _
                  Or Len(pv_NewValue) <> 4 _
                  Or (Mid(pv_NewValue, 1, 2) < 1 Or Mid(pv_NewValue, 1, 2) > 31) _
                  Or (Mid(pv_NewValue, 3, 2) < 1 Or Mid(pv_NewValue, 3, 2) > 12)))
    ErrMsg = "Invalid DdMm Name": Exit Sub
  Case Is = UCase("wSOCmCdTgt")
    Cancel = (pv_NewValue <> "") And Not moCn.RecSeek("Select CmCd from CustMst where CmCtg= 'C' and CmCd = '" + pv_NewValue + "' ")
    If Cancel = True Then ErrMsg = "Invalid Target Customer Code": Exit Sub
  End Select
End Sub
Private Sub atxt_ReCalculate(Index As Integer, ByVal IdName As String, ByVal OldVal As Variant, ByVal RowNum As Integer, ByVal SourceName As String, ByVal SourceColName As String)
  Select Case UCase(IdName)
    Case Is = UCase("wSOVchChrTgt")
      If ADC.Mode = xNorm Then
       ADC("wSOVchChrTgt") = moCn.GetFldVal("Select vPValue from vParam where " + _
                                  "vPCoCd= '" + gs_CoCd + "' and vPTyp='DEFCHR' " + _
                                  "and vPMCd='" + ADC("wSOVchTcTgt") + "'")
      End If
  End Select
End Sub
Private Sub CmdClearErrLog_Click()
  If MsgBox("Are you sure you want to clear Transfer-In Log?", vbYesNo + vbDefaultButton2, "XML To XEmr Transfer") = vbYes Then
    moCn.Execute ("Delete from ErrLog ")
    GrdErrLog.Rows = 1
  End If
End Sub
Private Sub CmdErrLogGo_Click()
  Call ShowErrLog
End Sub
Private Sub CmdShowErrLog_Click()
  Call DispFra(ErrLog)
  If FraErrLog.Visible = True Then
    'FraErrLog.Caption = "Error Log"
    ADC("wTblFr").SetFocus
  End If
End Sub
Private Sub CmdSOTrfIn_Click()
  If ADC("wSOVchTcTgt") = "" Then DispMsg "Target Voucher Tc Cannot Be Blank", etError: ADC("wSOVchTcTgt").SetFocus: Exit Sub
  If ADC("wSOVchYyTgt") = "" Then DispMsg "Target Voucher Yy Cannot Be Blank", etError: ADC("wSOVchYyTgt").SetFocus: Exit Sub
  If ADC("wSOVchChrTgt") = "" Then DispMsg "Target Voucher Chr Cannot Be Blank", etError: ADC("wSOVchChrTgt").SetFocus: Exit Sub
  If ADC("wSOCmCdTgt") = "" Then DispMsg "Target Customer Code Cannot Be Blank", etError: ADC("wSOCmCdTgt").SetFocus: Exit Sub
  If ms_FilePth = "" Then DispMsg "Invalid File Name/ File Name Cannot Be Blank", etError: Exit Sub
  DispMsg "", etInfo
  
  ms_OmTc = ADC("wSOVchTcTgt")
  ms_OmYy = ADC("wSOVchYyTgt")
  ms_OmChr = ADC("wSOVchChrTgt")

  Call TransferInRoutine("OrdMst", "OrdDsg", "OrdRm", "OrdLab", , "OrdTgt", "OrdTgt", , "SORD")
  Call CmdSO_Click
End Sub
Private Sub TransferInRoutine(ps_TblX As String, Optional ByVal ps_TblA As Variant, _
                              Optional ByVal ps_TblB As Variant, _
                              Optional ByVal ps_TblC As Variant, _
                              Optional ByVal ps_TblB1 As Variant, _
                              Optional ByVal ps_TblC1 As Variant, _
                              Optional ByVal ps_TblD As Variant, _
                              Optional ByVal ps_TblG As Variant, _
                              Optional ps_VchTyp As Variant)
  
  On Error GoTo errh
  
  ms_ErrTbl = ""
  Set xmlErrRecSet = Nothing
  
  Dim ws_TblA As String, ws_TblB As String, ws_TblC As String, ws_TblB1 As String
  Dim ws_TblC1 As String, ws_TblD As String, ws_TblG As String
  Dim ws_TcTyp As String, ws_Cnd As String, wb_Flag As Boolean
  Dim wb_SourceAvlbl As Boolean, ws_KeyVal As String, ws_KeyStr As String
  Dim wf_xTdTotAmt As Single, ws_VchTyp As String
    
  '*** Keeps count of Records successfully transferred in and the total no. of records
  '    to be transfered
  Dim wi_Cnt As Integer, wi_TotCnt As Integer
  Dim ws_TrnfCntStr As String
    
  mf_SrcCurCnvRt = 0 '*** For xTxn
  wb_SourceAvlbl = False: ws_KeyVal = "": ws_KeyStr = ""

  Call VerMatch
  
  Set XmlDoc = New MSXML2.DOMDocument
  '*** Load The XML Document InTo Our Document Object
  XmlDoc.Load ms_FilePth
  ADC("wDisp") = ps_TblX + " Transfer-In Is In Process ..."
  '*** Accessing THe RootElement Which Is A Child Of The Document Object
  Set xmlRootElement = XmlDoc.childNodes(0)
    
  If Not IsMissing(ps_VchTyp) Then
    Select Case UCase(ps_VchTyp)
    Case Is = "SORD"
      If UCase(xmlRootElement.baseName) <> UCase("xEmrPOTransfer") Then
        DispMsg "File Does Not Contain Proper Data", etError
        ADC("wDisp") = "File Does Not Contain Proper Data"
        Exit Sub
      End If
      '*** 20/12/04
      If xmlRootElement.Attributes.length - 1 < 0 Then DispMsg "XML File Does Not Have A Proper Version", etError: ADC("wDisp") = "XML File Does Not Have A Proper Version": Exit Sub
      If UCase(xmlRootElement.Attributes.Item(0).nodeName) <> UCase("Ver") Then DispMsg "XML File Does Not Have A Proper Version", etError: ADC("wDisp") = "XML File Does Not Have A Proper Version": Exit Sub
      If Trim(xmlRootElement.Attributes.Item(0).nodeValue) <> gs_Ver Then DispMsg "XML Version Not Proper", etError: ADC("wDisp") = "XML Version Not Proper": Exit Sub
      '*** 20/12/04
    End Select
    ws_VchTyp = ps_VchTyp
'''''  Else
'''''    If UCase(xmlRootElement.baseName) <> UCase("xEmr" + ps_TblX + "Transfer") Then
'''''      DispMsg "File Does Not Contain Proper Data", etError
'''''      ADC("wDisp") = "File Does Not Contain Proper Data"
'''''      Exit Sub
'''''    End If
'''''   ws_VchTyp = ""
  End If
  
  '*** Opening Targer RecordSet And XML Doc(RecSource) For Transfer To Table
  Set Rs_TgtAddX = moCn.OpenResultset("Select * from " + ps_TblX + " where 1= 2")
  If Not IsMissing(ps_TblA) Then Set Rs_TgtAddA = moCn.OpenResultset("Select * from " + ps_TblA + " where 1= 2"): ws_TblA = ps_TblA
  If Not IsMissing(ps_TblB) Then Set Rs_TgtAddB = moCn.OpenResultset("Select * from " + ps_TblB + " where 1= 2"): ws_TblB = ps_TblB
  If Not IsMissing(ps_TblC) Then Set Rs_TgtAddC = moCn.OpenResultset("Select * from " + ps_TblC + " where 1= 2"): ws_TblC = ps_TblC
  If Not IsMissing(ps_TblB1) Then Set Rs_TgtAddB1 = moCn.OpenResultset("Select * from " + ps_TblB1 + " where 1= 2"): ws_TblB1 = ps_TblB1
  If Not IsMissing(ps_TblC1) Then Set Rs_TgtAddC1 = moCn.OpenResultset("Select * from " + ps_TblC1 + " where 1= 2"): ws_TblC1 = ps_TblC1
  If Not IsMissing(ps_TblD) Then Set Rs_TgtAddD = moCn.OpenResultset("Select * from " + ps_TblD + " where 1= 2"): ws_TblD = ps_TblD
  If Not IsMissing(ps_TblG) Then Set Rs_TgtAddG = moCn.OpenResultset("Select * from " + ps_TblG + " where 1= 2"): ws_TblG = ps_TblG
  
  '*** Total Records to be transferred
  wi_TotCnt = xmlRootElement.childNodes.length
  
  '*** Loop For Retrieving X Recs - Child Elements Of RootElement
  For mf_i = 0 To xmlRootElement.childNodes.length - 1
    DoEvents
    moCn.BeginTrans (mdoCommitRec)
    '*** Take A Reference Of The Current X Rec In The Document
    'Dim xmlChildX As IXMLDOMElement
    Set xmlChildX = xmlRootElement.childNodes(mf_i)
    
    ms_ErrTbl = ps_TblX
    Set xmlErrRecSet = xmlChildX
    
    If CheckForRecExists(ps_TblX, xmlChildX, ws_TcTyp, ws_Cnd, ws_KeyVal) = True Then
      If ws_VchTyp = "SORD" Then moCn.CommitTrans mdoCommitRec: GoTo XNext
      ADC("wDisp") = ps_TblX + ": " + ws_KeyVal
      Set Rs_TgtUpdX = moCn.OpenResultset("Select * from " + ps_TblX + " " + ws_Cnd)
      If Rs_TgtUpdX.RecCount > 0 Then
        wb_Flag = True
        Rs_TgtUpdX.edit
        MoveAttributesToFields Rs_TgtUpdX, xmlChildX
        If Not IsMissing(ps_VchTyp) Then Call GenFldVal(ws_TcTyp, ps_TblX, Rs_TgtUpdX)
        Rs_TgtUpdX.Update
      End If
      '''Set Rs_TgtUpdX = Nothing
    Else
      ADC("wDisp") = ps_TblX + ": " + ws_KeyVal
      wb_Flag = False
      '*** Add A New Record In The Target X RecordSet
      Rs_TgtAddX.Addnew
      '*** Read Attributes Of Current X Element And Move It To Fields Of RecordSet
      MoveAttributesToFields Rs_TgtAddX, xmlChildX
      If Not IsMissing(ps_VchTyp) Then Call GenFldVal(ws_TcTyp, ps_TblX, Rs_TgtAddX)
      Rs_TgtAddX.Update
    End If
    '*** Update The X Record To The
    
    '*** Loop For A_Level Records - ChildNodes Of Current X Element
    For mf_j = 0 To xmlChildX.childNodes.length - 1
      '*** Take A Reference Of The Current A Rec In The Document
      Set xmlChildLevel1 = xmlChildX.childNodes(mf_j)
      
      If UCase(xmlChildLevel1.baseName) = UCase(ws_TblA) Then
        ms_ErrTbl = ws_TblA
        Set xmlErrRecSet = xmlChildLevel1
        
        If UCase(ws_TcTyp) = "SO" Then
          'Check for Region Exclusive
          If CheckForRegExcl(ws_TblA, xmlChildLevel1, Rs_TgtAddX!OmCmCd) = True Then
            moCn.RollbackTrans mdoCommitNone: GoTo XNext
          End If
        End If
        
        If wb_Flag = True Then
          If CheckForRecExists(ws_TblA, xmlChildLevel1, ws_TcTyp, ws_Cnd) = True Then
            Set Rs_TgtUpdA = moCn.OpenResultset("Select * from " + ws_TblA + " " + ws_Cnd)
            If Rs_TgtUpdA.RecCount > 0 Then
              Rs_TgtUpdA.edit
              MoveAttributesToFields Rs_TgtUpdA, xmlChildLevel1
              If Not IsMissing(ps_VchTyp) Then Call GenFldVal(ws_TcTyp, ws_TblA, Rs_TgtUpdA)
              Rs_TgtUpdA.Update
            End If
            '''Set Rs_TgtUpdA = Nothing
          Else
            Rs_TgtAddA.Addnew
            MoveAttributesToFields Rs_TgtAddA, xmlChildLevel1
            If Not IsMissing(ps_VchTyp) Then Call GenFldVal(ws_TcTyp, ws_TblA, Rs_TgtAddA)
            Rs_TgtAddA.Update
          End If
        Else
          Rs_TgtAddA.Addnew
          MoveAttributesToFields Rs_TgtAddA, xmlChildLevel1
          If Not IsMissing(ps_VchTyp) Then Call GenFldVal(ws_TcTyp, ws_TblA, Rs_TgtAddA)
          Rs_TgtAddA.Update
        End If
      ElseIf UCase(xmlChildLevel1.baseName) = UCase(ws_TblD) Then
        ms_ErrTbl = ws_TblD
        Set xmlErrRecSet = xmlChildLevel1
     
        If wb_Flag = True Then
          If CheckForRecExists(ws_TblD, xmlChildLevel1, ws_TcTyp, ws_Cnd) = True Then
            Set Rs_TgtUpdD = moCn.OpenResultset("Select * from " + ws_TblD + " " + ws_Cnd)
            If Rs_TgtUpdD.RecCount > 0 Then
              Rs_TgtUpdD.edit
              MoveAttributesToFields Rs_TgtUpdD, xmlChildLevel1
              If Not IsMissing(ps_VchTyp) Then Call GenFldVal(ws_TcTyp, ws_TblD, Rs_TgtUpdD)
              Rs_TgtUpdD.Update
            End If
            '''Set Rs_TgtUpdD = Nothing
          Else
            Rs_TgtAddD.Addnew
            MoveAttributesToFields Rs_TgtAddD, xmlChildLevel1
            If Not IsMissing(ps_VchTyp) Then Call GenFldVal(ws_TcTyp, ws_TblD, Rs_TgtAddD)
            Rs_TgtAddD.Update
          End If
        Else
          Rs_TgtAddD.Addnew
          MoveAttributesToFields Rs_TgtAddD, xmlChildLevel1
          If Not IsMissing(ps_VchTyp) Then Call GenFldVal(ws_TcTyp, ws_TblD, Rs_TgtAddD)
          Rs_TgtAddD.Update
        End If
      ElseIf UCase(xmlChildLevel1.baseName) = UCase(ws_TblG) Then
        ms_ErrTbl = ws_TblG
        Set xmlErrRecSet = xmlChildLevel1
      
        If wb_Flag = True Then
          If CheckForRecExists(ws_TblG, xmlChildLevel1, ws_TcTyp, ws_Cnd) = True Then
            Set Rs_TgtUpdG = moCn.OpenResultset("Select * from " + ws_TblG + " " + ws_Cnd)
            If Rs_TgtUpdG.RecCount > 0 Then
              Rs_TgtUpdG.edit
              MoveAttributesToFields Rs_TgtUpdG, xmlChildLevel1
              If Not IsMissing(ps_VchTyp) Then Call GenFldVal(ws_TcTyp, ws_TblG, Rs_TgtUpdG)
              Rs_TgtUpdG.Update
            End If
            '''Set Rs_TgtUpdG = Nothing
          Else
            Rs_TgtAddG.Addnew
            MoveAttributesToFields Rs_TgtAddG, xmlChildLevel1
            If Not IsMissing(ps_VchTyp) Then Call GenFldVal(ws_TcTyp, ws_TblG, Rs_TgtAddG)
            Rs_TgtAddG.Update
          End If
        Else
          Rs_TgtAddG.Addnew
          MoveAttributesToFields Rs_TgtAddG, xmlChildLevel1
          If Not IsMissing(ps_VchTyp) Then Call GenFldVal(ws_TcTyp, ws_TblG, Rs_TgtAddG)
          Rs_TgtAddG.Update
        End If
      End If
      
      For mf_k = 0 To xmlChildLevel1.childNodes.length - 1
        Set xmlChildLevel2 = xmlChildLevel1.childNodes(mf_k)
        
        If UCase(xmlChildLevel2.baseName) = UCase(ws_TblB) Then
          ms_ErrTbl = ws_TblB
          Set xmlErrRecSet = xmlChildLevel2
        
          If wb_Flag = True Then
            If CheckForRecExists(ws_TblB, xmlChildLevel2, ws_TcTyp, ws_Cnd) = True Then
              Set Rs_TgtUpdB = moCn.OpenResultset("Select * from " + ws_TblB + " " + ws_Cnd)
              If Rs_TgtUpdB.RecCount > 0 Then
                Rs_TgtUpdB.edit
                MoveAttributesToFields Rs_TgtUpdB, xmlChildLevel2
                If Not IsMissing(ps_VchTyp) Then Call GenFldVal(ws_TcTyp, ws_TblB, Rs_TgtUpdB)
                Rs_TgtUpdB.Update
              End If
              '''Set Rs_TgtUpdB = Nothing
            Else
              Rs_TgtAddB.Addnew
              MoveAttributesToFields Rs_TgtAddB, xmlChildLevel2
              If Not IsMissing(ps_VchTyp) Then Call GenFldVal(ws_TcTyp, ws_TblB, Rs_TgtAddB)
              Rs_TgtAddB.Update
            End If
          Else
            Rs_TgtAddB.Addnew
            MoveAttributesToFields Rs_TgtAddB, xmlChildLevel2
            If Not IsMissing(ps_VchTyp) Then Call GenFldVal(ws_TcTyp, ws_TblB, Rs_TgtAddB)
            Rs_TgtAddB.Update
          End If
        ElseIf UCase(xmlChildLevel2.baseName) = UCase(ws_TblC) Then
          ms_ErrTbl = ws_TblC
          Set xmlErrRecSet = xmlChildLevel2
          
          If wb_Flag = True Then
            If CheckForRecExists(ws_TblC, xmlChildLevel2, ws_TcTyp, ws_Cnd) = True Then
              Set Rs_TgtUpdC = moCn.OpenResultset("Select * from " + ws_TblC + " " + ws_Cnd)
              If Rs_TgtUpdC.RecCount > 0 Then
                Rs_TgtUpdC.edit
                MoveAttributesToFields Rs_TgtUpdC, xmlChildLevel2
                If Not IsMissing(ps_VchTyp) Then Call GenFldVal(ws_TcTyp, ws_TblC, Rs_TgtUpdC)
                Rs_TgtUpdC.Update
              End If
              '''Set Rs_TgtUpdC = Nothing
            Else
              Rs_TgtAddC.Addnew
              MoveAttributesToFields Rs_TgtAddC, xmlChildLevel2
              If Not IsMissing(ps_VchTyp) Then Call GenFldVal(ws_TcTyp, ws_TblC, Rs_TgtAddC)
              Rs_TgtAddC.Update
            End If
          Else
            Rs_TgtAddC.Addnew
            MoveAttributesToFields Rs_TgtAddC, xmlChildLevel2
            If Not IsMissing(ps_VchTyp) Then Call GenFldVal(ws_TcTyp, ws_TblC, Rs_TgtAddC)
            Rs_TgtAddC.Update
          End If
        ElseIf UCase(xmlChildLevel2.baseName) = UCase(ws_TblB1) Then
          ms_ErrTbl = ws_TblB1
          Set xmlErrRecSet = xmlChildLevel2
        
          If wb_Flag = True Then
            If CheckForRecExists(ws_TblB1, xmlChildLevel2, ws_TcTyp, ws_Cnd) = True Then
              Set Rs_TgtUpdB1 = moCn.OpenResultset("Select * from " + ws_TblB1 + " " + ws_Cnd)
              If Rs_TgtUpdB1.RecCount > 0 Then
                Rs_TgtUpdB1.edit
                MoveAttributesToFields Rs_TgtUpdB1, xmlChildLevel2
                If Not IsMissing(ps_VchTyp) Then Call GenFldVal(ws_TcTyp, ws_TblB1, Rs_TgtUpdB1)
                Rs_TgtUpdB1.Update
              End If
              '''Set Rs_TgtUpdB1 = Nothing
            Else
              Rs_TgtAddB1.Addnew
              MoveAttributesToFields Rs_TgtAddB1, xmlChildLevel2
              If Not IsMissing(ps_VchTyp) Then Call GenFldVal(ws_TcTyp, ws_TblB1, Rs_TgtAddB1)
              Rs_TgtAddB1.Update
            End If
          Else
            Rs_TgtAddB1.Addnew
            MoveAttributesToFields Rs_TgtAddB1, xmlChildLevel2
            If Not IsMissing(ps_VchTyp) Then Call GenFldVal(ws_TcTyp, ws_TblB1, Rs_TgtAddB1)
            Rs_TgtAddB1.Update
          End If
        ElseIf UCase(xmlChildLevel2.baseName) = UCase(ws_TblC1) Then
          ms_ErrTbl = ws_TblC1
          Set xmlErrRecSet = xmlChildLevel2
        
          If wb_Flag = True Then
            If CheckForRecExists(ws_TblC1, xmlChildLevel2, ws_TcTyp, ws_Cnd) = True Then
              Set Rs_TgtUpdC1 = moCn.OpenResultset("Select * from " + ws_TblC1 + " " + ws_Cnd)
              If Rs_TgtUpdC1.RecCount > 0 Then
                Rs_TgtUpdC1.edit
                MoveAttributesToFields Rs_TgtUpdC1, xmlChildLevel2
                If Not IsMissing(ps_VchTyp) Then Call GenFldVal(ws_TcTyp, ws_TblC1, Rs_TgtUpdC1)
                Rs_TgtUpdC1.Update
              End If
              '''Set Rs_TgtUpdC1 = Nothing
            Else
              Rs_TgtAddC1.Addnew
              MoveAttributesToFields Rs_TgtAddC1, xmlChildLevel2
              If Not IsMissing(ps_VchTyp) Then Call GenFldVal(ws_TcTyp, ws_TblC1, Rs_TgtAddC1)
              Rs_TgtAddC1.Update
            End If
          Else
            Rs_TgtAddC1.Addnew
            MoveAttributesToFields Rs_TgtAddC1, xmlChildLevel2
            If Not IsMissing(ps_VchTyp) Then Call GenFldVal(ws_TcTyp, ws_TblC1, Rs_TgtAddC1)
            Rs_TgtAddC1.Update
          End If
        End If
        Set xmlChildLevel2 = Nothing
        '*** GoTo Next Rec In 2nd-Level Element
      Next mf_k
      
      ' Calculate OrdDsg Cst
      If UCase(ps_TblA) = UCase("OrdDsg") Then
        ' CheckForRecExists() Not Required
        If CheckForRecExists(ws_TblA, xmlChildLevel1, ws_TcTyp, ws_Cnd) = True Then
          ADC("wDisp") = ps_TblA + ": " + ws_KeyVal
          Set Rs_TgtUpdA = moCn.OpenResultset("Select * from " + ps_TblA + " " + ws_Cnd)
          If Rs_TgtUpdA.RecCount > 0 Then
            Rs_TgtUpdA.edit
            If Not IsMissing(ps_VchTyp) Then Call GenFldVal(ws_TcTyp, ws_TblA, Rs_TgtUpdA, ws_TblA)
            Rs_TgtUpdA.Update
          End If
        End If
      End If
      Set xmlChildLevel1 = Nothing
      '*** GoTo Next Rec In 1st-Level Element
    Next mf_j
    
    moCn.CommitTrans mdoCommitRec
    
    Set xmlChildX = Nothing
    '*** GoTo Next X Element
    
    wi_Cnt = wi_Cnt + 1
    
XNext:
  Next mf_i
XClose:
  ws_TrnfCntStr = "(" + CStr(wi_Cnt) + "/" + CStr(wi_TotCnt) + " Records Transferred)"

  If wi_TotCnt <> wi_Cnt Then
    ADC("wDisp") = ps_TblX + " Transfer-In Completed With Errors. " + ws_TrnfCntStr
  Else
    ADC("wDisp") = ps_TblX + " Transfer-In Completed Successfully. " + ws_TrnfCntStr
  End If
  
  Set Rs_TgtAddX = Nothing: Set Rs_TgtAddA = Nothing
  Set Rs_TgtAddB = Nothing: Set Rs_TgtAddC = Nothing
  Set Rs_TgtAddB1 = Nothing: Set Rs_TgtAddC1 = Nothing
  Set Rs_TgtAddD = Nothing: Set Rs_TgtAddG = Nothing
  
  Set Rs_TgtUpdX = Nothing: Set Rs_TgtUpdA = Nothing
  Set Rs_TgtUpdB = Nothing: Set Rs_TgtUpdC = Nothing
  Set Rs_TgtUpdB1 = Nothing: Set Rs_TgtUpdC1 = Nothing
  Set Rs_TgtUpdD = Nothing: Set Rs_TgtUpdG = Nothing
  
  Set xmlErrRecSet = Nothing
  
  Exit Sub
errh:
  If Not (Rs_TgtAddX Is Nothing) Then
    If Not (Rs_TgtAddX.EOF Or Rs_TgtAddX.BOF) Then Rs_TgtAddX.CancelUpdate
  End If
  If Not (Rs_TgtAddA Is Nothing) Then
    If Not (Rs_TgtAddA.EOF Or Rs_TgtAddA.BOF) Then Rs_TgtAddA.CancelUpdate
  End If
  If Not (Rs_TgtAddB Is Nothing) Then
    If Not (Rs_TgtAddB.EOF Or Rs_TgtAddB.BOF) Then Rs_TgtAddB.CancelUpdate
  End If
  If Not (Rs_TgtAddC Is Nothing) Then
    If Not (Rs_TgtAddC.EOF Or Rs_TgtAddC.BOF) Then Rs_TgtAddC.CancelUpdate
  End If
  If Not (Rs_TgtAddB1 Is Nothing) Then
    If Not (Rs_TgtAddB1.EOF Or Rs_TgtAddB1.BOF) Then Rs_TgtAddB1.CancelUpdate
  End If
  If Not (Rs_TgtAddC1 Is Nothing) Then
    If Not (Rs_TgtAddC1.EOF Or Rs_TgtAddC1.BOF) Then Rs_TgtAddC1.CancelUpdate
  End If
  If Not (Rs_TgtAddD Is Nothing) Then
    If Not (Rs_TgtAddD.EOF Or Rs_TgtAddD.BOF) Then Rs_TgtAddD.CancelUpdate
  End If
  If Not (Rs_TgtAddG Is Nothing) Then
    If Not (Rs_TgtAddG.EOF Or Rs_TgtAddG.BOF) Then Rs_TgtAddG.CancelUpdate
  End If
  
  If Not (Rs_TgtUpdX Is Nothing) Then
    If Not (Rs_TgtUpdX.EOF Or Rs_TgtUpdX.BOF) Then Rs_TgtUpdX.CancelUpdate
  End If
  If Not (Rs_TgtUpdA Is Nothing) Then
    If Not (Rs_TgtUpdA.EOF Or Rs_TgtUpdA.BOF) Then Rs_TgtUpdA.CancelUpdate
  End If
  If Not (Rs_TgtUpdB Is Nothing) Then
    If Not (Rs_TgtUpdB.EOF Or Rs_TgtUpdB.BOF) Then Rs_TgtUpdB.CancelUpdate
  End If
  If Not (Rs_TgtUpdC Is Nothing) Then
    If Not (Rs_TgtUpdC.EOF Or Rs_TgtUpdC.BOF) Then Rs_TgtUpdC.CancelUpdate
  End If
  If Not (Rs_TgtUpdB1 Is Nothing) Then
    If Not (Rs_TgtUpdB1.EOF Or Rs_TgtUpdB1.BOF) Then Rs_TgtUpdB1.CancelUpdate
  End If
  If Not (Rs_TgtUpdC1 Is Nothing) Then
    If Not (Rs_TgtUpdC1.EOF Or Rs_TgtUpdC1.BOF) Then Rs_TgtUpdC1.CancelUpdate
  End If
  If Not (Rs_TgtUpdD Is Nothing) Then
    If Not (Rs_TgtUpdD.EOF Or Rs_TgtUpdD.BOF) Then Rs_TgtUpdD.CancelUpdate
  End If
  If Not (Rs_TgtUpdG Is Nothing) Then
    If Not (Rs_TgtUpdG.EOF Or Rs_TgtUpdG.BOF) Then Rs_TgtUpdG.CancelUpdate
  End If
  
  wb_SourceAvlbl = Not (xmlErrRecSet Is Nothing)
  
  If wb_SourceAvlbl Then
    Call CheckForRecExists(ms_ErrTbl, xmlErrRecSet, , , ws_KeyVal)
    ws_KeyStr = ws_KeyVal
  Else
    ws_KeyStr = "No Source Data"
  End If
  
  Call ErrHndlr(CStr(Format(moCn.SrvrDate, "DDMM")) + ms_ErrTbl, ws_KeyStr)
  
  If wb_SourceAvlbl Then
    Resume XNext
  Else
    Resume XClose
  End If
End Sub
Private Sub MoveAttributesToFields(ByRef pRs_Tgt As MwfLib.MDORowSet, po_ReqXmlElement As IXMLDOMElement)
  '*** Loop Through Attributes Of The Current Element
  '*** The Name Of THe Attribute Is Assumed To Be The Same As Field Of The RecordSet
  '*** Moving The Values
  For mi_i = 0 To po_ReqXmlElement.Attributes.length - 1
    pRs_Tgt(po_ReqXmlElement.Attributes.Item(mi_i).nodeName) = po_ReqXmlElement.Attributes.Item(mi_i).nodeValue
  Next mi_i
End Sub
Private Function CheckForRecExists(fs_Tbl As String, fo_ReqXmlElement As IXMLDOMElement, Optional ByRef fs_TcTyp As Variant, Optional ByRef fs_Cnd As Variant, Optional ByRef fs_KeyVal As Variant) As Boolean
  Dim ws_TcTyp As String, ws_Cnd As String, ws_KeyVal As String
  CheckForRecExists = False
  
  Select Case UCase(fs_Tbl)
    Case Is = UCase("OrdMst")
      mf_OmNo = moCn.AutoGen("OrdMst", "OmNo", gs_CoCd + ms_OmTc + ms_OmYy + ms_OmChr)
      ws_TcTyp = moCn.GetFldVal("Select PValue from Param where PTyp= 'TC' and PMCd= '" + ms_OmTc + "'")
      ws_Cnd = " where OmCoCd= '" + gs_CoCd + "' and OmTc= '" + ms_OmTc + "' and " + _
               "OmYy= '" + ms_OmYy + "' and OmChr= '" + ms_OmChr + "' and OmNo= " + CStr(mf_OmNo) + " "
      If Not IsMissing(fs_Cnd) Then CheckForRecExists = moCn.RecSeek("Select 'x' from " + fs_Tbl + " " + ws_Cnd)
      ws_KeyVal = gs_CoCd + "/" + ms_OmTc + "/" + ms_OmYy + "/" + ms_OmChr + "/" + Trim(CStr(mf_OmNo))
    Case Is = UCase("OrdDsg")
      Dim wf_OdSr As Single
      For mi_i = 0 To fo_ReqXmlElement.Attributes.length - 1
        If UCase(fo_ReqXmlElement.Attributes.Item(mi_i).nodeName) = UCase("OdSr") Then wf_OdSr = fo_ReqXmlElement.Attributes.Item(mi_i).nodeValue
        If wf_OdSr <> 0 Then Exit For
      Next mi_i
      ws_TcTyp = moCn.GetFldVal("Select PValue from Param where PTyp= 'TC' and PMCd= '" + ms_OmTc + "'")
      ws_Cnd = " where OdCoCd= '" + gs_CoCd + "' and OdTc= '" + ms_OmTc + "' " + _
               "and OdYy= '" + ms_OmYy + "' and OdChr= '" + ms_OmChr + "' and OdNo= " + CStr(mf_OmNo) + " " + _
               "and OdSr= " + CStr(wf_OdSr) + " "
      If Not IsMissing(fs_Cnd) Then CheckForRecExists = moCn.RecSeek("Select 'x' from " + fs_Tbl + " " + ws_Cnd)
      ws_KeyVal = gs_CoCd + "/" + ms_OmTc + "/" + ms_OmYy + "/" + ms_OmChr + "/" + Trim(CStr(mf_OmNo)) + "/" + Trim(CStr(wf_OdSr))
    Case Is = UCase("OrdRm")
      Dim wf_OrSr As Single, wf_OrSrNo As Single
      For mi_i = 0 To fo_ReqXmlElement.Attributes.length - 1
        If UCase(fo_ReqXmlElement.Attributes.Item(mi_i).nodeName) = UCase("OrSr") Then wf_OrSr = fo_ReqXmlElement.Attributes.Item(mi_i).nodeValue
        If UCase(fo_ReqXmlElement.Attributes.Item(mi_i).nodeName) = UCase("OrSrNo") Then wf_OrSrNo = fo_ReqXmlElement.Attributes.Item(mi_i).nodeValue
        If wf_OrSr <> 0 And wf_OrSrNo <> 0 Then Exit For
      Next mi_i
      ws_TcTyp = moCn.GetFldVal("Select PValue from Param where PTyp= 'TC' and PMCd= '" + ms_OmTc + "'")
      ws_Cnd = " where OrCoCd= '" + gs_CoCd + "' and OrTc= '" + ms_OmTc + "' " + _
               "and OrYy= '" + ms_OmYy + "' and OrChr= '" + ms_OmChr + "' and OrNo= " + CStr(mf_OmNo) + " " + _
               "and OrSr= " + CStr(wf_OrSr) + " and OrSrNo= " + CStr(wf_OrSrNo) + " "
      If Not IsMissing(fs_Cnd) Then CheckForRecExists = moCn.RecSeek("Select 'x' from " + fs_Tbl + " " + ws_Cnd)
      ws_KeyVal = gs_CoCd + "/" + ms_OmTc + "/" + ms_OmYy + "/" + ms_OmChr + "/" + Trim(CStr(mf_OmNo)) + "/" + Trim(CStr(wf_OrSr)) + "/" + Trim(CStr(wf_OrSrNo))
    Case Is = UCase("OrdLab")
      Dim wf_OlSr As Single, wf_OlSrNo As Single
      For mi_i = 0 To fo_ReqXmlElement.Attributes.length - 1
        If UCase(fo_ReqXmlElement.Attributes.Item(mi_i).nodeName) = UCase("OlSr") Then wf_OlSr = fo_ReqXmlElement.Attributes.Item(mi_i).nodeValue
        If UCase(fo_ReqXmlElement.Attributes.Item(mi_i).nodeName) = UCase("OlSrNo") Then wf_OlSrNo = fo_ReqXmlElement.Attributes.Item(mi_i).nodeValue
        If wf_OlSr <> 0 And wf_OlSrNo <> 0 Then Exit For
      Next mi_i
      ws_TcTyp = moCn.GetFldVal("Select PValue from Param where PTyp= 'TC' and PMCd= '" + ms_OmTc + "'")
      ws_Cnd = " where OlCoCd= '" + gs_CoCd + "' and OlTc= '" + ms_OmTc + "' " + _
               "and OlYy= '" + ms_OmYy + "' and OlChr= '" + ms_OmChr + "' and OlNo= " + CStr(mf_OmNo) + " " + _
               "and OlSr= " + CStr(wf_OlSr) + " and OlSrNo= " + CStr(wf_OlSrNo) + " "
      If Not IsMissing(fs_Cnd) Then CheckForRecExists = moCn.RecSeek("Select 'x' from " + fs_Tbl + " " + ws_Cnd)
      ws_KeyVal = gs_CoCd + "/" + ms_OmTc + "/" + ms_OmYy + "/" + ms_OmChr + "/" + Trim(CStr(mf_OmNo)) + "/" + Trim(CStr(wf_OlSr)) + "/" + Trim(CStr(wf_OlSrNo))
    Case Is = UCase("OrdTgt")
      Dim wf_OtSr As Single, ws_OtRmCtg As String, ws_OtRmSCtg As String
      For mi_i = 0 To fo_ReqXmlElement.Attributes.length - 1
        If UCase(fo_ReqXmlElement.Attributes.Item(mi_i).nodeName) = UCase("OtSr") Then wf_OtSr = fo_ReqXmlElement.Attributes.Item(mi_i).nodeValue
        If UCase(fo_ReqXmlElement.Attributes.Item(mi_i).nodeName) = UCase("OtRmCtg") Then ws_OtRmCtg = fo_ReqXmlElement.Attributes.Item(mi_i).nodeValue
        If UCase(fo_ReqXmlElement.Attributes.Item(mi_i).nodeName) = UCase("OtRmSCtg") Then ws_OtRmSCtg = fo_ReqXmlElement.Attributes.Item(mi_i).nodeValue
        If wf_OtSr <> 0 And ws_OtRmCtg <> "" And ws_OtRmSCtg <> "" Then Exit For
      Next mi_i
      ws_TcTyp = moCn.GetFldVal("Select xPValue from xParam where xPTyp= 'TC' and xPMCd= '" + ms_OmTc + "'")
      ws_Cnd = " where OtCoCd= '" + gs_CoCd + "' and OtTc= '" + ms_OmTc + "' " + _
               "and OtYy= '" + ms_OmYy + "' and OtChr= '" + ms_OmChr + "' and OtNo= " + CStr(mf_OmNo) + " " + _
               "and OtSr= " + CStr(wf_OtSr) + " and OtRmCtg= " + CStr(ws_OtRmCtg) + " " + _
               "and OtRmSCtg= " + CStr(ws_OtRmSCtg) + " "
      If Not IsMissing(fs_Cnd) Then CheckForRecExists = moCn.RecSeek("Select 'x' from " + fs_Tbl + " " + ws_Cnd)
      ws_KeyVal = gs_CoCd + "/" + ms_OmTc + "/" + ms_OmYy + "/" + ms_OmChr + "/" + Trim(CStr(mf_OmNo)) + "/" + Trim(CStr(wf_OtSr)) + "/" + ws_OtRmCtg + "/" + ws_OtRmSCtg
  End Select

  If Not IsMissing(fs_TcTyp) Then fs_TcTyp = ws_TcTyp
  If Not IsMissing(fs_Cnd) Then fs_Cnd = ws_Cnd
  If Not IsMissing(fs_KeyVal) Then fs_KeyVal = ws_KeyVal
End Function
Private Sub GenFldVal(fs_TcTyp As String, fs_Tbl As String, ByRef fRs_Tbl As MwfLib.MDORowSet, Optional ByRef fs_CalcTbl As Variant)
  Dim ws_TgtCmCd As String
  Dim wRs_Head As MDORowSet, wRs_OrdMst As MDORowSet, wRs_OrdDsg As MDORowSet
  Dim wRs_OrdRm As MDORowSet, wRs_OrdLab As MDORowSet
  Dim ws_TgtCurCd As String
  Dim wf_TgtCurrCnvRt As Single
  
  Set wRs_Head = moCn.OpenRes("Select * from Head where HCoCd= '" + gs_CoCd + "' ")
  If UCase(fs_TcTyp) = "SO" Then
    ws_TgtCmCd = ADC("wSOCmCdTgt")
    ws_TgtCurCd = moCn.GetFldVal("Select CmCurCd from CustMst where CmCtg= 'C' and CmCd= '" + ws_TgtCmCd + "'")
    wf_TgtCurrCnvRt = moCn.GetFldVal("Select PNum from Param where PTyp= 'CURNCY' and PMCd= '" + ws_TgtCurCd + "' and PSCd= ''")
    wf_TgtCurrCnvRt = IIF(wf_TgtCurrCnvRt > 0, wf_TgtCurrCnvRt, 1)
  End If
  
  Select Case UCase(fs_Tbl)
  Case Is = UCase("OrdMst")
    fRs_Tbl!OmCoCd = gs_CoCd
    fRs_Tbl!OmTc = ms_OmTc
    fRs_Tbl!OmYy = ms_OmYy
    fRs_Tbl!OmChr = ms_OmChr
    fRs_Tbl!OmNo = mf_OmNo
    
    fRs_Tbl!OmCmCd = ws_TgtCmCd
    fRs_Tbl!OmCmPayTerms = moCn.GetFldVal("Select CmPayTerms from CustMst where CmCtg= '' and CmCd= '" + ws_TgtCmCd + "' ")
    fRs_Tbl!OmCmDelTerms = moCn.GetFldVal("Select CmPayTerms from CustMst where CmCtg= '' and CmCd= '" + ws_TgtCmCd + "' ")
    If fRs_Tbl!OmExpDelDt < moCn.SrvrDate Then fRs_Tbl!OmExpDelDt = moCn.SrvrDate
    If fRs_Tbl!OmDelDt < moCn.SrvrDate Then fRs_Tbl!OmDelDt = moCn.SrvrDate
    
    'If gs_Ver = "2.7.0" Or gs_Ver = "2.8.0" Or gs_Ver = "2.9.0" Then
    '  fRs_Tbl!OmLockSalPrcYN = "N"
    'End If
    
    If Not (fRs_Tbl.EOF Or fRs_Tbl.BOF) Then
      fRs_Tbl!OmLine = wRs_Head!HLine
      fRs_Tbl!OmHCtCd = wRs_Head!HCtCd
      fRs_Tbl!OmFixPrc = wRs_Head!hCmFixPrc
    End If
  Case Is = UCase("OrdDsg")
    If IsMissing(fs_CalcTbl) Then
      fRs_Tbl!OdCoCd = gs_CoCd
      fRs_Tbl!OdTc = ms_OmTc
      fRs_Tbl!OdYy = ms_OmYy
      fRs_Tbl!OdChr = ms_OmChr
      fRs_Tbl!OdNo = mf_OmNo
      
      fRs_Tbl!OdSalPrc = fRs_Tbl!OdSalPrc * wf_TgtCurrCnvRt
      
      If fRs_Tbl!OdExpDelDt < moCn.SrvrDate Then fRs_Tbl!OdExpDelDt = moCn.SrvrDate
      If fRs_Tbl!OdDelDt < moCn.SrvrDate Then fRs_Tbl!OdDelDt = moCn.SrvrDate

      fRs_Tbl!OdBagPcs = moCn.GetFldVal("Select DmBagPcs from DsgMst where DmTcTyp= 'DM' and DmCd= '" + fRs_Tbl!OdDmCd + "' " + _
                         " and DmSz= '" + fRs_Tbl!OdDmSz + "'")
      If Not (fRs_Tbl.EOF Or fRs_Tbl.BOF) Then
        fRs_Tbl!OdLine = wRs_Head!HLine
        fRs_Tbl!OdFixPrc = wRs_Head!hCmFixPrc
        fRs_Tbl!OdPrdSeq = wRs_Head!HDmPrdSeq
        fRs_Tbl!OdVaCtg = wRs_Head!HDmVaCtg
        fRs_Tbl!OdOmCtCd = wRs_Head!HCtCd
        fRs_Tbl!OdWh = wRs_Head!HWh
      End If
      
      'xTdOffValPpc As OdSalPrc
      'xTdSalValPpc As OdCalcPrc
    End If

    If Not IsMissing(fs_CalcTbl) Then
      Dim wd_RmVal As Double, wd_LabVal As Double
      Dim wd_RmCstVal As Double, wd_LabCstVal As Double

      Set wRs_OrdRm = moCn.OpenRes("Select Sum(OrSalVal)+Sum(OrSetSalVal) as qRmCalcVal, " + _
                     "Sum(OrCstVal)+ Sum(OrSetCstVal) as qRmCstVal, " + _
                     "(case when sum(OrWsQty)> 0 and sum(OrHsQty)= 0 then 'W' " + _
                            "when sum(OrWsQty)= 0 and sum(OrHsQty)> 0 then 'H' " + _
                            "when sum(OrWsQty)> 0 and sum(OrHsQty)> 0 then 'WH' else '' end) as qOdWh " + _
                     "from OrdRm where OrCoCd= '" + gs_CoCd + "' " + _
                     "and OrTc= '" + fRs_Tbl!OdTc + "' and OrYy= '" + fRs_Tbl!OdYy + "' " + _
                     "and OrChr= '" + fRs_Tbl!OdChr + "' and OrNo= " & fRs_Tbl!OdNo & " " + _
                     "and OrSr= " & fRs_Tbl!OdSr & " ")
      Set wRs_OrdLab = moCn.OpenRes("Select Sum(OlSalVal) as qLabSalVal, " + _
                      "Sum(OlCstVal) as qLabCstVal from OrdLab where " + _
                      "OlCoCd= '" + gs_CoCd + "' and OlTc= '" + fRs_Tbl!OdTc + "' " + _
                      "and OlYy= '" + fRs_Tbl!OdYy + "' and OlChr= '" + fRs_Tbl!OdChr + "' " + _
                      "and OlNo= " & fRs_Tbl!OdNo & " and OlSr= " & fRs_Tbl!OdSr & " ")
      
      If Not (wRs_OrdRm.EOF Or wRs_OrdRm.BOF) Then
        wd_RmVal = wRs_OrdRm!qRmCalcVal
        wd_RmCstVal = wRs_OrdRm!qRmCstVal
        fRs_Tbl!OdWh = IIF(wRs_OrdRm!qOdWh <> "", wRs_OrdRm!qOdWh, fRs_Tbl!OdWh)
      End If
      If Not (wRs_OrdLab.EOF Or wRs_OrdLab.BOF) Then
        wd_LabVal = wRs_OrdLab!qLabSalVal
        wd_LabCstVal = wRs_OrdLab!qLabCstVal
      End If
      
      fRs_Tbl!OdCalcPrc = IIF(wd_RmVal + wd_LabVal < 0, 0, wd_RmVal + wd_LabVal)
      fRs_Tbl!OdCstPrc = IIF(wd_RmCstVal + wd_LabCstVal < 0, 0, wd_RmCstVal + wd_LabCstVal) * fRs_Tbl!OdMulBy
    End If
  Case Is = UCase("OrdRm")
    fRs_Tbl!OrCoCd = gs_CoCd
    fRs_Tbl!OrTc = ms_OmTc
    fRs_Tbl!OrYy = ms_OmYy
    fRs_Tbl!OrChr = ms_OmChr
    fRs_Tbl!OrNo = mf_OmNo
    
    Set wRs_OrdMst = moCn.OpenRes("Select * from OrdMst where OmCoCd= '" + gs_CoCd + "' " + _
                     "and OmTc= '" + fRs_Tbl!OrTc + "' and OmYy= '" + fRs_Tbl!OrYy + "' " + _
                     "and OmChr= '" + fRs_Tbl!OrChr + "' and OmNo=" & fRs_Tbl!OrNo & " ")
    Set wRs_OrdDsg = moCn.OpenRes("Select * from OrdDsg where OdCoCd= '" + gs_CoCd + "' " + _
                     "and OdTc= '" + fRs_Tbl!OrTc + "' and OdYy= '" + fRs_Tbl!OrYy + "' " + _
                     "and OdChr= '" + fRs_Tbl!OrChr + "' and OdNo= " & fRs_Tbl!OrNo & " " + _
                      "and OdSr= " & fRs_Tbl!OrSr & " ")
    ' Zubin 213
    If fRs_Tbl!OrRmCtg = "G" Then
      fRs_Tbl!OrCstRt = GetGldRt(wRs_OrdMst!OmCmCd, wRs_OrdMst!OmLmgCst, fRs_Tbl!OrRmCd, fRs_Tbl!OrAlyCstRt, _
                                  "DM", wRs_OrdDsg!OdDmCd, "C", "C", , wRs_OrdMst!OmCmCurCd)
    ElseIf fRs_Tbl!OrRmCtg = "P" Then
      fRs_Tbl!OrCstRt = GetGldRt(wRs_OrdMst!OmCmCd, wRs_OrdMst!OmLmpCst, fRs_Tbl!OrRmCd, fRs_Tbl!OrAlyCstRt, _
                                 "DM", wRs_OrdDsg!OdDmCd, "C", "C", , wRs_OrdMst!OmCmCurCd)
    ElseIf fRs_Tbl!OrRmCtg = "S" Then
      fRs_Tbl!OrCstRt = GetGldRt(wRs_OrdMst!OmCmCd, wRs_OrdMst!OmLmsCst, fRs_Tbl!OrRmCd, fRs_Tbl!OrAlyCstRt, _
                                 "DM", wRs_OrdDsg!OdDmCd, "C", "C", , wRs_OrdMst!OmCmCurCd)
    Else
      '*** (Bef 2.13)
      'fRs_Tbl!OrCstRt = GetRmRt(wRs_OrdMst!OmCmCd, fRs_Tbl!OrRmCd, fRs_Tbl!OrLn1, fRs_Tbl!OrLn2, _
      '                          Round(MWLib.Div(fRs_Tbl!OrPrdWt, fRs_Tbl!OrPrdQty), 4), "C", 0)
      '*** (Bef 2.13)
      
      '*** (Jen 2.13)
      Dim wf_BaseLme As Single
      If Trim(UCase(fRs_Tbl!OrRmCtg)) = "X" Then
        wf_BaseLme = moCn.GetFldVal("Select (case B.RmCtg when 'G' then " + CStr(wRs_OrdMst!OmLmgCst) + " " + _
                     "when 'P' then " + CStr(wRs_OrdMst!OmLmpCst) + " when 'S' then " + CStr(wRs_OrdMst!OmLmsCst) + " " + _
                     "else 0 end) from RmMst B where " + _
                     "RmCd= (Select Rm.RmBaseCd from RmMst Rm where Rm.RmCd= '" + fRs_Tbl!OrRmCd + "') ")
      Else
        wf_BaseLme = 0
      End If
      fRs_Tbl!OrCstRt = GetRmRt(wRs_OrdMst!OmCmCd, fRs_Tbl!OrRmCd, fRs_Tbl!OrLn1, fRs_Tbl!OrLn2, _
                        Round(MWLib.Div(fRs_Tbl!OrPrdWt, fRs_Tbl!OrPrdQty), 4), "C", wf_BaseLme, , , , , wRs_OrdMst!OmCmCurCd)
      '*** (Jen 2.13)
      
    End If
    fRs_Tbl!OrCstVal = IIF(moCn.GetFldVal("Select RmQw from RmMst where RmCd= '" + fRs_Tbl!OrRmCd + "' ") = "Q", _
                           fRs_Tbl!OrPrdQty, fRs_Tbl!OrPrdWt) * fRs_Tbl!OrCstRt
    ' Zubin 213
    fRs_Tbl!OrSetCstRt = GetLabRt(wRs_OrdMst!OmCmCd, "SET", fRs_Tbl!OrSetSCd, "C", wRs_OrdDsg!OdGldAsWt, "C", , , , wRs_OrdMst!OmCmCurCd)
    ' Zubin 213
    fRs_Tbl!OrSetCstVal = GetLabVal(wRs_OrdMst!OmCmCd, "SET", fRs_Tbl!OrSetSCd, fRs_Tbl!OrSetCstRt, _
                                    fRs_Tbl!OrPrdQty, 0, "Q", "C", "C", wRs_OrdMst!OmCmCurCd)
    'fRs_Tbl!OrAlyCstRt = GetRmRt(wRs_OrdMst!OmCmCd, fRs_Tbl!OrAlyCd, 0, 0, 0, "C")   '*** (Bef 2.13)
    ' Zubin 213
    fRs_Tbl!OrAlyCstRt = GetRmRt(wRs_OrdMst!OmCmCd, fRs_Tbl!OrAlyCd, 0, 0, 0, "C", 0, , , , , wRs_OrdMst!OmCmCurCd)   '*** (Jen 2.13)

    fRs_Tbl!OrSalRt = fRs_Tbl!OrSalRt * wf_TgtCurrCnvRt
    fRs_Tbl!OrSalVal = fRs_Tbl!OrSalVal * wf_TgtCurrCnvRt
    fRs_Tbl!OrSetSalRt = fRs_Tbl!OrSetSalRt * wf_TgtCurrCnvRt
    fRs_Tbl!OrSetSalVal = fRs_Tbl!OrSetSalVal * wf_TgtCurrCnvRt
    fRs_Tbl!OrAlySalRt = fRs_Tbl!OrAlySalRt * wf_TgtCurrCnvRt
    
'    xTrSalRt As OrSalRt
'    xTrSalValPPc As OrSalVal
'    xTrSetSalRt As OrSetSalRt
'    xTrSetSalValPPc As OrSetSalVal
'    xTrAlySalRt As OrAlySalRt

  Case Is = UCase("OrdLab")
    Dim ws_OlCstQw As String
    fRs_Tbl!OlCoCd = gs_CoCd
    fRs_Tbl!OlTc = ms_OmTc
    fRs_Tbl!OlYy = ms_OmYy
    fRs_Tbl!OlChr = ms_OmChr
    fRs_Tbl!OlNo = mf_OmNo
    
    Set wRs_OrdMst = moCn.OpenRes("Select * from OrdMst where OmCoCd= '" + gs_CoCd + "' " + _
                     "and OmTc='" + fRs_Tbl!OlTc + "' and OmYy= '" + fRs_Tbl!OlYy + "' " + _
                     "and OmChr= '" + fRs_Tbl!OlChr + "' and OmNo= " & fRs_Tbl!OlNo & " ")
    Set wRs_OrdDsg = moCn.OpenRes("Select * from OrdDsg where OdCoCd= '" + gs_CoCd + "' " + _
                     "and OdTc= '" + fRs_Tbl!OlTc + "' and OdYy= '" + fRs_Tbl!OlYy + "' " + _
                     "and OdChr= '" + fRs_Tbl!OlChr + "' and OdNo= " & fRs_Tbl!OlNo & " " + _
                     "and OdSr= " & fRs_Tbl!OlSr & " ")
    fRs_Tbl!OlCstQw = moCn.GetFldVal("Select PValue from Param where PTyp= 'LABSCD' " + _
                      "and PMCd= '" + fRs_Tbl!OlMCd + "' and PSCd= '" + fRs_Tbl!OlSCd + "'")
    
    ' Zubin 213
    fRs_Tbl!OlCstRt = GetLabRt(wRs_OrdMst!OmCmCd, fRs_Tbl!OlMCd, fRs_Tbl!OlSCd, "C", wRs_OrdDsg!OdGldAsWt, "C", ws_OlCstQw, , , wRs_OrdMst!OmCmCurCd)
    If ws_OlCstQw <> "" Then fRs_Tbl!OlCstQw = ws_OlCstQw
    If fRs_Tbl!OlQty = 0 Then fRs_Tbl!OlQty = 1
    
    ' Zubin 213
    fRs_Tbl!OlCstVal = GetLabVal(wRs_OrdMst!OmCmCd, fRs_Tbl!OlMCd, fRs_Tbl!OlSCd, fRs_Tbl!OlCstRt, _
                                 fRs_Tbl!OlQty, wRs_OrdDsg!OdGldAsWt, fRs_Tbl!OlCstQw, "C", "C", wRs_OrdMst!OmCmCurCd)
    
    fRs_Tbl!OlSalRt = fRs_Tbl!OlSalRt * wf_TgtCurrCnvRt
    fRs_Tbl!OlSalVal = fRs_Tbl!OlSalVal * wf_TgtCurrCnvRt
  
    'xTLSalRt As OlSalRt
    'xTLSalValPPc As OlSalVal
  
  Case Is = UCase("OrdTgt")
    fRs_Tbl!OtCoCd = gs_CoCd
    fRs_Tbl!OtTc = ms_OmTc
    fRs_Tbl!OtYy = ms_OmYy
    fRs_Tbl!OtChr = ms_OmChr
    fRs_Tbl!OtNo = mf_OmNo
  End Select
End Sub
Private Function OpenDocFilePath(fs_XTblName As String) As String
  Dim wl_XMLSizeInKB As Single, wo_FileSysObj, wo_MyFile
  
  On Error GoTo ErrHndlr
  
  ComDialogOpn.DialogTitle = "Open " + fs_XTblName + " XML File As"
  ComDialogOpn.InitDir = moCn.GetFldVal("Select PDesc225 from Param where PTyp= 'TRFDIR' and PMCd= 'TRFDIR' ")
  ComDialogOpn.Filter = "XML Files (*.xml)"
  ComDialogOpn.FilterIndex = 1
  '*** cdlOFNOverwritePrompt:     specifies whether the user should be alerted if he is overwriting and existing file
  '*** cdlOFNFileMustExist:       Specifies that the user can enter only names of existing files in the File Name text box
  '*** cdlOFNPathMustExist:       makes sure that the the user should specify an existing path
  '*** cdlOFNLongNames:           specifies that long names can be used for the files
  '*** cdlOFNExtensionDifferent:  indicates that the user has selected a different extension than the one specified through the DefaultExt property
  ComDialogOpn.Flags = cdlOFNLongNames Or cdlOFNFileMustExist Or cdlOFNPathMustExist Or cdlOFNExtensionDifferent
  ComDialogOpn.DefaultExt = "xml"
  ComDialogOpn.CancelError = True
  ComDialogOpn.ShowOpen
  
  If Len(ComDialogOpn.FileName) > 4 And Trim(ComDialogOpn.FileName) <> "" And Dir(ComDialogOpn.FileName) <> "" Then
    If UCase(Right(ComDialogOpn.FileName, 4)) = ".XML" Then
      OpenDocFilePath = ComDialogOpn.FileName
    Else
      OpenDocFilePath = ""
      DispMsg "File Extension Is Wrong", etError: Exit Function
    End If
  End If

  If OpenDocFilePath <> "" Then
    Set wo_FileSysObj = CreateObject("Scripting.FileSystemObject")
    Set wo_MyFile = wo_FileSysObj.GetFile(ComDialogOpn.FileName)
    wl_XMLSizeInKB = MWLib.Div(wo_MyFile.Size, 1024)
    '*** For Rounding Up By One
    wl_XMLSizeInKB = IIF(wl_XMLSizeInKB > Round(wl_XMLSizeInKB), 1, 0) + Round(wl_XMLSizeInKB)
    If wl_XMLSizeInKB > 1200 Then
      OpenDocFilePath = ""
      DispMsg "The File Size Was Exceeding The Maximum Permissible Limits Hence Cannot Be Opened", etError
    End If
  End If
Exit Function

ErrHndlr:
  If Err.Number = cdlCancel Then
    ADC("wDisp") = ""
    DispMsg "Transfer Cancelled", etInfo
    Err.Clear
  End If
  OpenDocFilePath = ""

End Function
Private Sub DispFra(ByVal pv_DispFra As en_EmrTrfInDispFra)
  '*** This is the common procedure which is used to toggle between hiding and displaying
      'the frames (i.e. if the frame is visible then make it invisible and viseversa)
      'associated with a command button

  Select Case pv_DispFra
    Case Is = SO
      If FraSO.Visible = True Then
        FraSO.Visible = False
        '*** required to pass the array position SO-0, to restore the 'Causes Validation' State
        Call EnaDisaCmds(False, CmdSO, SO)
        CmdSO.SetFocus
      Else
        FraSO.Visible = True
        CmdSOTrfIn.SetFocus
        FraSO.ZOrder
        Call EnaDisaCmds(True, CmdSO)
      End If
    Case Is = ErrLog
      If FraErrLog.Visible = True Then
        FraErrLog.Visible = False
        Call EnaDisaCmds(False, CmdShowErrLog, ErrLog)
        CmdShowErrLog.SetFocus
      Else
        FraErrLog.Visible = True
        CmdShowErrLog.SetFocus
        FraErrLog.ZOrder
        Call EnaDisaCmds(True, CmdShowErrLog)
      End If
  End Select
End Sub
Private Sub EnaDisaCmds(ByVal pv_ShowFra As Boolean, Optional ByVal pv_CmdBut As CommandButton, Optional ByVal pv_DispFra As en_EmrTrfInDispFra)
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
    CmdSO.Enabled = Not pv_ShowFra
    CmdClearErrLog.Enabled = Not pv_ShowFra
    CmdShowErrLog.Enabled = Not pv_ShowFra
  Else
    '*** Restoring Ena/Disa state of all Command Buttons
    CmdSO.Enabled = mArr_Ed(SO)
    CmdClearErrLog.Enabled = mArr_Ed(ErrLog)
    CmdShowErrLog.Enabled = mArr_Ed(ErrLog)
  End If

  If Not pv_CmdBut Is Nothing Then pv_CmdBut.Enabled = True
End Sub
Private Sub HideAllFras()
  '*** This procedure hides all the frames other than FraNKeyAll
  'FraNKeyAll.Visible = True
  FraSO.Visible = False
  FraErrLog.Visible = False
End Sub
Private Sub StoreState()
  '*** This procedure is used to store the CausesValidation State and the Enabled/Disabled
      'states of all the Command Buttons
  '*** This is used for restoring the original states in the future
  mArr_Cv(SO) = CmdSO.CausesValidation
  mArr_Cv(ErrLog) = CmdShowErrLog.CausesValidation
  
  mArr_Ed(SO) = CmdSO.Enabled
  mArr_Ed(ErrLog) = CmdShowErrLog.Enabled
End Sub
Private Sub ErrHndlr(ByVal TableName As String, ByVal KeyVal As String)
  Dim ws_ReturnStr As String
    
  If InStr(1, Err.Description, "|") > 0 Then
    ws_ReturnStr = Mid(Err.Description, InStr(1, Err.Description, "|") + 1)
  Else
    ws_ReturnStr = Err.Description
    Dim wf_Ctr As Single, ws_ErrMsg As String, ws_ExecStr As String
    For wf_Ctr = 0 To moCn.ErrCount - 1
      If moCn.ErrObj(wf_Ctr).Number = -2147217900 Then
        ws_ReturnStr = ""
        ws_ErrMsg = "Record Already Exists"
      Else
        If InStr(1, moCn.ErrObj(wf_Ctr).Description, "|") > 0 Then
          ws_ErrMsg = Mid(moCn.ErrObj(wf_Ctr).Description, InStr(1, moCn.ErrObj(wf_Ctr).Description, "|") + 1)
        Else
          ws_ErrMsg = moCn.ErrObj(wf_Ctr).Description
        End If
      End If
      ws_ReturnStr = ws_ReturnStr + "::" + ws_ErrMsg
    Next wf_Ctr
  End If
  ws_ReturnStr = ws_ReturnStr + "   " + KeyVal

  ADC("wDisp") = ws_ReturnStr
  DoEvents

  '*** This statement is not needed here as code for Rollback exists in the trigger
  'moCn.RollbackTrans mdoCommitRec
  '*** This statement is not needed here as code for Rollback exists in the trigger
  
  If InStr(1, TableName + KeyVal + ws_ReturnStr, "'") <> 0 Then
    If InStr(1, KeyVal, "'") = 0 Then
      ws_ExecStr = "insert into ErrLog values('" + TableName + "','" + KeyVal + "','Cannot Display Msg As It Contains Single Quotes')"
    Else
      ws_ExecStr = "insert into ErrLog values('" + TableName + "','Key Contains Quotes','Cannot Display Msg As It Contains Single Quotes')"
    End If
  Else
    ws_ExecStr = "insert into ErrLog values('" + TableName + "','" + KeyVal + "','" + ws_ReturnStr + "')"
  End If

  moCn.Execute (ws_ExecStr)
  DispMsg ws_ReturnStr, etInfo
  moCn.ErrClear
End Sub
Private Sub CmdSO_Click()
  Call DispFra(SO)
  If FraSO.Visible = True Then
    ' **** After Clr Scrn **** '
    'FraSO.Caption = "Sales Order Transfer In"
    LblMain(0).Caption = "  Sales Order Transfer In"
    ' **** After Clr Scrn **** '
    ms_FilePth = OpenDocFilePath("Order")
    ADC("wFilePath") = ms_FilePth
    ms_Opt = "SO"
    If ms_FilePth = "" Then Call DispFra(SO)
    If FraSO.Visible Then ADC("wSOVchTcTgt").SetFocus
  End If
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
  ws_BtnStr = "CmdErrLogGo"
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
    
    '*** (09/08/05)
    Set moCn = Nothing
    
    Set XmlDoc = Nothing
    Set xmlRootElement = Nothing
    Set xmlChildX = Nothing
    Set xmlChildLevel1 = Nothing
    Set xmlChildLevel2 = Nothing
    '*** (09/08/05)
End Sub
'***
Private Sub ShowErrLog()
  Dim wRsErrLog As MwfLib.MDORowSet, wi_Row As Integer
  Dim wQuryStr As String, ws_Cnd As String
  Dim wi_SrNo As Integer
  
  If ADC("wTblFr") <> "" Then ws_Cnd = " substring(ElTable, 5, Len(ElTable)-4) >= '" + ADC("wTblFr") + "' and "
  If ADC("wTblTo") <> "" Then ws_Cnd = ws_Cnd + " substring(ElTable, 5, Len(ElTable)-4) <= '" + ADC("wTblTo") + "' and "
  If ADC("wDdMmFr") <> "" Then ws_Cnd = ws_Cnd + " substring(ElTable, 1, 4) >= '" + ADC("wDdMmFr") + "' and "
  If ADC("wDdMmTo") <> "" Then ws_Cnd = ws_Cnd + " substring(ElTable, 1, 4) <= '" + ADC("wDdMmTo") + "' and "
  
  If Right(ws_Cnd, 4) = "and " Then ws_Cnd = Mid(ws_Cnd, 1, Len(ws_Cnd) - 4)
  
  If ws_Cnd <> "" Then ws_Cnd = " where " + ws_Cnd
  wQuryStr = "Select * from ErrLog " + ws_Cnd + _
                " Order By ElTable, ElKey, ElErrMsg "
  Set wRsErrLog = moCn.OpenRes(wQuryStr)
  
  With GrdErrLog
    .Rows = 1
    .AllowAdd = False
    .AllowDelete = False
  End With
  
  With wRsErrLog
    Do While Not (.EOF Or .BOF)
      GrdErrLog.AddItem
      wi_Row = GrdErrLog.Rows - 1
      wi_SrNo = wi_SrNo + 1
      GrdErrLog.Value(wi_Row, "wSrNo") = wi_SrNo
      GrdErrLog.Value(wi_Row, "wElTable") = Mid(!ElTable, 5)
      GrdErrLog.Value(wi_Row, "wElKey") = !ElKey
      GrdErrLog.Value(wi_Row, "wElErrMsg") = !ElErrMsg
      .MoveNext
    Loop
  End With
  Set wRsErrLog = Nothing

End Sub
Private Function CheckForRegExcl(fs_Tbl As String, fo_ReqXmlElement As IXMLDOMElement, ByRef fs_OmCmCd As String) As Boolean
  Dim ws_DmCd As String, ws_CmRegnCd As String
  
  ws_CmRegnCd = moCn.GetFldVal("Select CmRegnCd from CustMst where CmCtg= 'C' and CmCd= '" + fs_OmCmCd + "'")
  Select Case UCase(fs_Tbl)
    Case Is = UCase("OrdDsg")
      Dim ws_OdTc As String, ws_OdYy As String, ws_OdChr As String, wf_OdNo As Single
      Dim wf_OdSr As Single, wf_OdSrNo As Single
      For mi_i = 0 To fo_ReqXmlElement.Attributes.length - 1
        If UCase(fo_ReqXmlElement.Attributes.Item(mi_i).nodeName) = UCase("OdTc") Then ws_OdTc = fo_ReqXmlElement.Attributes.Item(mi_i).nodeValue
        If UCase(fo_ReqXmlElement.Attributes.Item(mi_i).nodeName) = UCase("OdYy") Then ws_OdYy = fo_ReqXmlElement.Attributes.Item(mi_i).nodeValue
        If UCase(fo_ReqXmlElement.Attributes.Item(mi_i).nodeName) = UCase("OdChr") Then ws_OdChr = fo_ReqXmlElement.Attributes.Item(mi_i).nodeValue
        If UCase(fo_ReqXmlElement.Attributes.Item(mi_i).nodeName) = UCase("OdNo") Then wf_OdNo = fo_ReqXmlElement.Attributes.Item(mi_i).nodeValue
        If UCase(fo_ReqXmlElement.Attributes.Item(mi_i).nodeName) = UCase("OdSr") Then wf_OdSr = fo_ReqXmlElement.Attributes.Item(mi_i).nodeValue
        If UCase(fo_ReqXmlElement.Attributes.Item(mi_i).nodeName) = UCase("OdSrNo") Then wf_OdSrNo = fo_ReqXmlElement.Attributes.Item(mi_i).nodeValue
        If UCase(fo_ReqXmlElement.Attributes.Item(mi_i).nodeName) = UCase("OdDmCd") Then ws_DmCd = fo_ReqXmlElement.Attributes.Item(mi_i).nodeValue
        If ws_OdTc <> "" And ws_OdYy <> "" And ws_OdChr <> "" And wf_OdNo <> 0 And wf_OdSr <> 0 And ws_DmCd <> "" Then Exit For
      Next mi_i
      
      CheckForRegExcl = Not moCn.RecSeek("Select 'x' From DsgMst Where " + _
                 "DmTcTyp= 'DM' And DmHld = 'N' and DmCd = '" + ws_DmCd + "' and DmSz= '' and " + _
                 "(DmCmCd= '' or DmCmCd= '" + fs_OmCmCd + "' or " + _
                 "(DmCmCd<> '" + fs_OmCmCd + "' and DmRegnExc= 'Y' and " + _
                 "(Select CmRegnCd from CustMst where CmCtg= 'C' and CmCd= DmCmCd) <> '" + ws_CmRegnCd + "')) ")
  End Select

End Function


