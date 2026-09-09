VERSION 5.00
Object = "{6A2CE895-C1AC-48AF-B558-FB94BBE12499}#1.0#0"; "mwfctl.ocx"
Begin VB.Form EmrFrmVldRmLn 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Valid RM Ln Definition"
   ClientHeight    =   5235
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7920
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MDIChild        =   -1  'True
   ScaleHeight     =   5235
   ScaleWidth      =   7920
   Begin MwfCtl.MWCTL_HLP MWCTL_HLP1 
      Height          =   435
      Left            =   2430
      TabIndex        =   2
      TabStop         =   0   'False
      Top             =   4680
      Visible         =   0   'False
      Width           =   435
      _ExtentX        =   767
      _ExtentY        =   767
   End
   Begin MwfCtl.MWCTL_ADC ADC 
      CausesValidation=   0   'False
      Height          =   555
      Left            =   105
      TabIndex        =   3
      Top             =   4710
      Width           =   7365
      _ExtentX        =   12991
      _ExtentY        =   979
      AllowAdd        =   0   'False
      AllowDelete     =   0   'False
      BtnWidth        =   885
      BtnHeight       =   435
      BtnMaskColor    =   12648447
      FontSz          =   8.25
   End
   Begin VB.Frame FRA_KEYS 
      BorderStyle     =   0  'None
      Height          =   330
      Left            =   -36
      TabIndex        =   5
      Top             =   24
      Width           =   5220
      Begin MwfCtl.MWCTL_MED ATXT 
         Height          =   285
         Index           =   7
         Left            =   1005
         TabIndex        =   0
         ToolTipText     =   "Enter RM Ctg"
         Top             =   0
         Width           =   540
         _ExtentX        =   953
         _ExtentY        =   503
         DataType        =   4
         IdName          =   "WVRSRMCTG"
         CmpStr          =   "VrsRmCtg ="
      End
      Begin MwfCtl.MWCTL_MED ATXT 
         Height          =   285
         Index           =   28
         Left            =   2625
         TabIndex        =   1
         ToolTipText     =   "Enter RM ID"
         Top             =   0
         Width           =   900
         _ExtentX        =   1588
         _ExtentY        =   503
         MaxLength       =   5
         DataType        =   4
         IdName          =   "WVRSRMID"
         CmpStr          =   "VrsRmID ="
      End
      Begin VB.Label ALBL 
         BackStyle       =   0  'Transparent
         Caption         =   "Rm ID"
         BeginProperty Font 
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
         Left            =   1800
         TabIndex        =   8
         Top             =   0
         Width           =   1125
      End
      Begin VB.Label ALBL 
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
         Height          =   285
         Index           =   7
         Left            =   270
         TabIndex        =   7
         Top             =   0
         Width           =   915
      End
   End
   Begin VB.Frame FRA_NKEYS 
      BorderStyle     =   0  'None
      Enabled         =   0   'False
      Height          =   5040
      Left            =   -120
      TabIndex        =   4
      Top             =   0
      Width           =   7935
      Begin VB.Frame FraNKeyAll 
         BorderStyle     =   0  'None
         Height          =   4785
         Left            =   150
         TabIndex        =   6
         Top             =   90
         Width           =   7740
         Begin MwfCtl.MWCTL_FLX GrdVldLn 
            Height          =   4260
            Left            =   120
            TabIndex        =   9
            Top             =   360
            Width           =   7560
            _ExtentX        =   13335
            _ExtentY        =   7514
            Cols            =   6
            colname1        =   "VRSRMCTG"
            heading1        =   "Rm Ctg  "
            datafld1        =   "VrsRmCtg"
            datatype1       =   4
            maxlength1      =   8
            tooltiptext1    =   "Enter Date"
            colname2        =   "VRSRMID"
            heading2        =   "Rm ID     "
            datafld2        =   "VrsRmID"
            datatype2       =   4
            maxlength2      =   8
            colname3        =   "VRSLN1"
            heading3        =   "Ln 1               "
            datafld3        =   "VrsLn1"
            datatype3       =   2
            mask3           =   "##0.000"
            maxlength3      =   7
            colname4        =   "VRSLN2"
            heading4        =   "Ln 2               "
            datafld4        =   "VrsLn2"
            datatype4       =   2
            mask4           =   "##0.000"
            maxlength4      =   7
            colname5        =   "VRSLNDESC"
            heading5        =   " Rm Ln Desc                                        "
            datafld5        =   "VrsLnDesc"
            maxlength5      =   15
         End
      End
   End
End
Attribute VB_Name = "EmrFrmVldRmLn"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
  '---- General Comments on the Entry --------'
  Option Explicit
  Dim moCn As MwfLib.MDOConnection   '*** (Bef speed) bef 24/11/06
  
  Private Sub ADC_Load()
    Set moCn = adc.Connection
    Set adc.FirKeyCtl = adc("wVrsRmCtg")
    Set adc.FirNKeyCtl = GrdVldLn
    adc.Child = "GrdVldLn"
    Set GrdVldLn.PrevCtl = adc
    Set GrdVldLn.NextCtl = adc
    GrdVldLn.HotKey = "A"
  End Sub
  
  Private Sub ADC_KeyWhen()
    '*** Set the frame FraNKeyAll as the top most frame
    FraNKeyAll.ZOrder (0)
  End Sub
  
  Private Sub ADC_KeyValid(pr_Cancel As Boolean, pr_ErrMsg As String)
    '*** Check whether the Date Tag is valid
    'Call atxt_FldChange(0, UCase("wDtTag"), pr_Cancel, pr_ErrMsg, ADC("wDtTag"))
    'If pr_Cancel = True Then Exit Sub
  End Sub
  
  Private Sub ATXT_FldWhen(Index As Integer, ByVal IdName As String, Cancel As Boolean, ErrMsg As String)
    Select Case IdName
    Case Is = UCase("wVrsRmCtg")
      Call HlpList.PMCd("RMCTG")
    
    Case Is = UCase("wVrsRmId")
      Call HlpList.PSCd("RMID", adc("WVRSRMCTG"))
    End Select
    
  End Sub
  
  Private Sub atxt_FldChange(Index As Integer, ByVal IdName As String, Cancel As Boolean, ErrMsg As String, ByVal pv_NewValue As Variant)
    Select Case IdName
    Case Is = UCase("wVrsRmCtg")
      Cancel = pv_NewValue <> "" And Not moCn.RecSeek("Select PMCd From Param where PTyp= 'RMCTG' and " + _
              "PMCd = '" + pv_NewValue + "' And PMCd In ('D','C')")
      If Cancel = True Then ErrMsg = "RM Category Not Defined/Only 'D' Or 'C' can be Entered": Exit Sub
    
    Case Is = UCase("wVrsRmId")
      Cancel = pv_NewValue <> "" And Not moCn.RecSeek("Select PSCd From Param where PTyp= 'RMID' and " + _
              " PMCD='" + adc("wVrsRmCtg") + "' And PSCd = '" + pv_NewValue + "'")
      If Cancel = True Then ErrMsg = "RM ID Not Defined": Exit Sub

    End Select
  End Sub

  Private Sub GrdVldLn_ColWhen(ByVal RowNum As Integer, ByVal ColName As String, Cancel As Boolean, ErrMsg As String, ByVal Action As MwfCtl.en_fgAction)
    Select Case ColName
      Case Is = UCase("VrsRmCtg")
        Call HlpList.PMCd("RMCTG")
  
      Case Is = UCase("VrsRmID")
        Call HlpList.PSCd("RMID", GrdVldLn.Value(RowNum, "VRSRMCTG"))
  
    End Select
    
  End Sub
  
  Private Sub GrdVldLn_ColChange(ByVal RowNum As Integer, ByVal ColName As String, ByVal NewValue As Variant, Cancel As Boolean, ErrMsg As String)
    With GrdVldLn
      Select Case ColName
      Case Is = UCase("VrsRmCtg")
        If Not moCn.RecSeek("Select 'x' From Param Where PTyp='RMCTG' And PMCd='" + NewValue + "' And PMCd In ('D','C') ") Then Cancel = True: ErrMsg = "Raw Material Category Not Defined/Only 'D' Or 'C' can be Entered": Exit Sub
        
      Case Is = UCase("VrsRmSCtg")
        If Not moCn.RecSeek("Select 'x' From Param Where PTyp='RMID' And PMCd='" + .Value(RowNum, "VRSRMCTG") + "' And PSCd='" + NewValue + "' ") Then Cancel = True: ErrMsg = "Raw Material ID Not Defined": Exit Sub
      
      Case Is = UCase("VrsLn1")
        If NewValue = 0 Then Cancel = True: ErrMsg = "RM Ln1 Cannot Zero": Exit Sub
      
      End Select
    End With
    
  End Sub
  
  Private Sub ADC_SetRecSource()
    '*** Since this is an 'A' Entry, there is no RecSource associated with ADC. So set the
        'DataMode property of ADC control as xdatUnbound
    adc.DataMode = xdatUnbound
  End Sub
  
  Private Sub GrdVldLn_SetRecSource()
    '*** Set the Record Source of the Grid GrdVldLn
        'The sort option DtDt
        
    Dim ws_Cnd As String, ws_OrdBy As String
    ws_Cnd = adc.RepCond
    ws_Cnd = IIF(ws_Cnd <> "", " where " + ws_Cnd, "")
    ws_OrdBy = " Order By VrsRmCtg, VrsRmId"
    GrdVldLn.RecSource = " Select * from VldRmLn " + ws_Cnd + ws_OrdBy
    
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
    Set moCn = Nothing    '*** (09/08/05)    '*** (Bef speed) bef 24/11/06
  End Sub
  Private Sub GrdVldLn_Validate(Cancel As Boolean)
    Cancel = GrdVldLn.Validate
  End Sub
