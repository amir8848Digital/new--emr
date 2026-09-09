VERSION 5.00
Object = "{45ED3C8A-4E24-4B3A-9BBF-206FD4D635C7}#1.0#0"; "mwfctl.ocx"
Begin VB.Form EmrFrmRmIdSz 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "RmID Size Description"
   ClientHeight    =   7710
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5760
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MDIChild        =   -1  'True
   ScaleHeight     =   7710
   ScaleWidth      =   5760
   Begin MwfCtl.MWCTL_HLP MWCTL_HLP1 
      Height          =   435
      Left            =   2550
      TabIndex        =   3
      TabStop         =   0   'False
      Top             =   7050
      Visible         =   0   'False
      Width           =   435
      _ExtentX        =   767
      _ExtentY        =   767
   End
   Begin MwfCtl.MWCTL_ADC ADC 
      CausesValidation=   0   'False
      Height          =   555
      Left            =   105
      TabIndex        =   4
      Top             =   7110
      Width           =   5505
      _ExtentX        =   9710
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
      Height          =   630
      Left            =   60
      TabIndex        =   6
      Top             =   0
      Width           =   5445
      Begin MwfCtl.MWCTL_MED ATXT 
         Height          =   285
         Index           =   28
         Left            =   2460
         TabIndex        =   1
         ToolTipText     =   "Enter From RM Id"
         Top             =   60
         Width           =   1320
         _ExtentX        =   2328
         _ExtentY        =   503
         MaxLength       =   8
         DataType        =   4
         IdName          =   "WRISRMIDFR"
         CmpStr          =   "RisRmId >="
      End
      Begin MwfCtl.MWCTL_MED ATXT 
         Height          =   285
         Index           =   7
         Left            =   810
         TabIndex        =   0
         ToolTipText     =   "Enter RM Ctg"
         Top             =   60
         Width           =   540
         _ExtentX        =   953
         _ExtentY        =   503
         DataType        =   4
         IdName          =   "WRISRMCTG"
         CmpStr          =   "RisRmCtg ="
      End
      Begin MwfCtl.MWCTL_MED ATXT 
         Height          =   285
         Index           =   0
         Left            =   2460
         TabIndex        =   2
         ToolTipText     =   "Enter To RM Id"
         Top             =   360
         Width           =   1320
         _ExtentX        =   2328
         _ExtentY        =   503
         MaxLength       =   8
         DataType        =   4
         IdName          =   "WRISRMIDTO"
         CmpStr          =   "RisRmId <="
      End
      Begin VB.Label ALBL 
         Alignment       =   1  'Right Justify
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
         Left            =   1500
         TabIndex        =   11
         Top             =   360
         Width           =   765
      End
      Begin VB.Label ALBL 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "RmId Fr"
         BeginProperty Font 
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
         Left            =   1500
         TabIndex        =   9
         Top             =   60
         Width           =   765
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
         Left            =   60
         TabIndex        =   8
         Top             =   60
         Width           =   675
      End
   End
   Begin VB.Frame FRA_NKEYS 
      BorderStyle     =   0  'None
      Enabled         =   0   'False
      Height          =   7710
      Left            =   0
      TabIndex        =   5
      Top             =   0
      Width           =   5805
      Begin VB.Frame FraNKeyAll 
         BorderStyle     =   0  'None
         Height          =   7695
         Left            =   30
         TabIndex        =   7
         Top             =   0
         Width           =   5700
         Begin MwfCtl.MWCTL_FLX GrdRmIdSz 
            Height          =   6240
            Left            =   90
            TabIndex        =   10
            Top             =   750
            Width           =   5520
            _ExtentX        =   9737
            _ExtentY        =   11007
            Cols            =   5
            colname1        =   "RISRMCTG"
            heading1        =   "Rm Ctg  "
            datafld1        =   "RisRmCtg"
            datatype1       =   4
            maxlength1      =   1
            colname2        =   "RISRMID"
            heading2        =   "Rm Id                   "
            datafld2        =   "RisRmId"
            datatype2       =   4
            maxlength2      =   8
            colname3        =   "RISSZ1"
            heading3        =   "Sz 1               "
            datafld3        =   "RisSz1"
            datatype3       =   2
            mask3           =   "##0.000"
            maxlength3      =   7
            colname4        =   "RISRMSZDESC"
            heading4        =   "RmId Sz Desc        "
            datafld4        =   "RisRmSzDesc"
            datatype4       =   4
            maxlength4      =   9
         End
      End
   End
End
Attribute VB_Name = "EmrFrmRmIdSz"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
  '---- General Comments on the Entry --------'
  Option Explicit
  Dim moCn As MwfLib.MDOConnection   '*** (Bef speed) bef 24/11/06
  
  Private Sub ADC_Load()
    Set moCn = ADC.Connection
    Set ADC.FirKeyCtl = ADC("wRisRmCtg")
    Set ADC.FirNKeyCtl = GrdRmIdSz
    ADC.Child = "GrdRmIdSz"
    Set GrdRmIdSz.PrevCtl = ADC
    Set GrdRmIdSz.NextCtl = ADC
    GrdRmIdSz.HotKey = "A"
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
    Case Is = UCase("wRisRmCtg")
      Call HlpList.PMCd("RMCTG")
    
    Case Is = UCase("wRisRmIdFr"), UCase("wRisRmIdTo")
      Call HlpList.PSCd("RMID", ADC("wRisRmCtg"))
    End Select
    
  End Sub
  
  Private Sub atxt_FldChange(Index As Integer, ByVal IdName As String, Cancel As Boolean, ErrMsg As String, ByVal pv_NewValue As Variant)
    Select Case IdName
    Case Is = UCase("wRisRmCtg")
      Cancel = pv_NewValue <> "" And Not moCn.RecSeek("Select PMCd From Param where PTyp= 'RMCTG' and " + _
              "PMCd = '" + pv_NewValue + "' And PMCd In ('D','C')")
      If Cancel = True Then ErrMsg = "RM Category Not Defined/Only 'D' Or 'C' can be Entered": Exit Sub
    
    Case Is = UCase("wRisRmIdFr"), UCase("wRisRmIdTo")
      Cancel = pv_NewValue <> "" And Not moCn.RecSeek("Select PSCd From Param where PTyp= 'RMID' and " + _
              " PMCD='" + ADC("wRisRmCtg") + "' And PSCd = '" + pv_NewValue + "'")
      If Cancel = True Then ErrMsg = "RM Id Not Defined": Exit Sub

    End Select
  End Sub

  Private Sub GrdRmIdSz_ColWhen(ByVal RowNum As Integer, ByVal ColName As String, Cancel As Boolean, ErrMsg As String, ByVal Action As MwfCtl.en_fgAction)
    Select Case ColName
      Case Is = UCase("RisRmCtg")
        Call HlpList.PMCd("RMCTG")
  
      Case Is = UCase("RisRmId")
        Call HlpList.PSCd("RMID", GrdRmIdSz.Value(RowNum, "RisRmCtg"))
  
    End Select
    
  End Sub
  
  Private Sub GrdRmIdSz_ColChange(ByVal RowNum As Integer, ByVal ColName As String, ByVal NewValue As Variant, Cancel As Boolean, ErrMsg As String)
    With GrdRmIdSz
      Select Case ColName
      Case Is = UCase("RisRmCtg")
        If Not moCn.RecSeek("Select 'x' From Param Where PTyp='RMCTG' And PMCd='" + NewValue + "' And PMCd In ('D','C') ") Then Cancel = True
        ErrMsg = "Raw Material Category Not Defined/Only 'D' Or 'C' can be Entered": Exit Sub
        
      Case Is = UCase("RisRmId")
        If Not moCn.RecSeek("Select 'x' From Param Where PTyp='RMID' And PMCd='" + .Value(RowNum, "RisRmCtg") + "' And PSCd='" + NewValue + "' ") Then Cancel = True
        ErrMsg = "RmId Not Defined": Exit Sub
      
      Case Is = UCase("RisSz1")
        If NewValue = 0 Then Cancel = True: ErrMsg = "RM Size Cannot Zero": Exit Sub
      
      Case Is = UCase("RisRmSzDesc")
        If NewValue = "" Then Cancel = True: ErrMsg = "RmId Description Cannot Be Blank": Exit Sub
      End Select
    End With
    
  End Sub
  
  Private Sub ADC_SetRecSource()
    '*** Since this is an 'A' Entry, there is no RecSource associated with ADC. So set the
        'DataMode property of ADC control as xdatUnbound
    ADC.DataMode = xdatUnbound
  End Sub
  
  Private Sub GrdRmIdSz_SetRecSource()
    '*** Set the Record Source of the Grid GrdRmIdSz
        
    Dim ws_Cnd As String, ws_OrdBy As String

    ws_Cnd = ADC.RepCond
    ws_Cnd = IIF(ws_Cnd <> "", " where " + ws_Cnd, "")

    ws_OrdBy = " Order By RisRmCtg, RisRmId"
    GrdRmIdSz.RecSource = " Select * from RmIdSz " + ws_Cnd + ws_OrdBy
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
    Cancel = ATXT(Index).Validate
  End Sub
  Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    If UnloadMode <> vbFormCode Then Cancel = True: Exit Sub
    Set moCn = Nothing    '*** (09/08/05)    '*** (Bef speed) bef 24/11/06
  End Sub
  Private Sub GrdRmIdSz_Validate(Cancel As Boolean)
    Cancel = GrdRmIdSz.Validate
  End Sub
