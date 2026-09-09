VERSION 5.00
Object = "{6A2CE895-C1AC-48AF-B558-FB94BBE12499}#1.0#0"; "mwfctl.ocx"
Begin VB.Form EmrFrmDsgPrm 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Design Parameters"
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
   Begin MwfCtl.MWCTL_HLP MWCTL_HLP1 
      Height          =   405
      Left            =   6510
      TabIndex        =   7
      TabStop         =   0   'False
      Top             =   9360
      Visible         =   0   'False
      Width           =   555
      _ExtentX        =   979
      _ExtentY        =   714
   End
   Begin VB.Frame FRA_KEYS 
      BorderStyle     =   0  'None
      Height          =   330
      Left            =   -36
      TabIndex        =   9
      Top             =   24
      Width           =   8370
      Begin MwfCtl.MWCTL_MED ATXT 
         Height          =   285
         Index           =   0
         Left            =   975
         TabIndex        =   0
         ToolTipText     =   "Enter Labour Main Code"
         Top             =   0
         Width           =   885
         _ExtentX        =   1561
         _ExtentY        =   503
         MaxLength       =   5
         DataType        =   4
         IdName          =   "WDPTYP"
         CmpStr          =   "DpTyp="
      End
      Begin MwfCtl.MWCTL_MED ATXT 
         Height          =   285
         Index           =   8
         Left            =   2805
         TabIndex        =   1
         ToolTipText     =   "Enter From Code"
         Top             =   0
         Width           =   2460
         _ExtentX        =   4339
         _ExtentY        =   503
         MaxLength       =   16
         DataType        =   4
         IdName          =   "WDPCDFR"
         CmpStr          =   "DpCd>="
      End
      Begin MwfCtl.MWCTL_MED ATXT 
         Height          =   285
         Index           =   9
         Left            =   5865
         TabIndex        =   2
         ToolTipText     =   "Enter To Code"
         Top             =   -15
         Width           =   2460
         _ExtentX        =   4339
         _ExtentY        =   503
         MaxLength       =   16
         DataType        =   4
         IdName          =   "WDPCDTO"
         CmpStr          =   "DpCd<="
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
         Index           =   8
         Left            =   5370
         TabIndex        =   13
         Top             =   0
         Width           =   435
      End
      Begin VB.Label ALBL 
         BackStyle       =   0  'Transparent
         Caption         =   "Code  Fr"
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
         Left            =   1920
         TabIndex        =   12
         Top             =   0
         Width           =   915
      End
      Begin VB.Label LBLDpTyp 
         BackStyle       =   0  'Transparent
         Caption         =   "Type"
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
         Left            =   240
         TabIndex        =   11
         Top             =   0
         Width           =   735
      End
   End
   Begin VB.Frame FRA_NKEYS 
      BorderStyle     =   0  'None
      Enabled         =   0   'False
      Height          =   9300
      Left            =   -120
      TabIndex        =   8
      Top             =   0
      Width           =   15315
      Begin VB.Frame FraNKeyAll 
         BorderStyle     =   0  'None
         Height          =   9165
         Left            =   120
         TabIndex        =   10
         Top             =   0
         Width           =   15060
         Begin VB.TextBox TxtPwd 
            Height          =   285
            IMEMode         =   3  'DISABLE
            Left            =   11520
            MaxLength       =   30
            PasswordChar    =   "*"
            TabIndex        =   3
            Top             =   0
            Width           =   2340
         End
         Begin VB.Frame FraPic3D 
            BorderStyle     =   0  'None
            Height          =   4695
            Left            =   960
            TabIndex        =   14
            Top             =   4200
            Width           =   5235
            Begin VB.Image Pic3D 
               Height          =   4590
               Left            =   30
               Top             =   30
               Width           =   5130
            End
         End
         Begin MwfCtl.MWCTL_FLX GrdDsgPrm 
            Height          =   3750
            Left            =   60
            TabIndex        =   4
            Top             =   420
            Width           =   7215
            _ExtentX        =   12726
            _ExtentY        =   6615
            Cols            =   4
            colname1        =   "DPCD"
            heading1        =   "Code                                          "
            datafld1        =   "DpCd"
            datatype1       =   4
            maxlength1      =   16
            tooltiptext1    =   "Enter Code"
            colname2        =   "DPDESC"
            heading2        =   "Description                                               "
            datafld2        =   "DpDesc"
            maxlength2      =   45
            colname3        =   "DPDT"
            heading3        =   "Date         "
            datafld3        =   "DpDt"
            datatype3       =   3
            maxlength3      =   8
         End
         Begin MwfCtl.MWCTL_FLX GrdDsgPrmDet 
            Height          =   8670
            Left            =   7260
            TabIndex        =   5
            Top             =   420
            Width           =   7755
            _ExtentX        =   13679
            _ExtentY        =   15293
            Cols            =   6
            colname1        =   "DPDMNO"
            heading1        =   "Sr      "
            datafld1        =   "DpDmNo"
            datatype1       =   1
            mask1           =   "###0"
            maxlength1      =   4
            colname2        =   "DPDMCD"
            heading2        =   "Design Code              "
            datafld2        =   "DpDmCd"
            datatype2       =   4
            maxlength2      =   15
            tooltiptext2    =   "Enter Design Code"
            colname3        =   "DPPCTNM"
            heading3        =   "Picture Nm                                    "
            datafld3        =   "DpPctNm"
            datatype3       =   4
            maxlength3      =   50
            colname4        =   "DPREM1"
            heading4        =   "Rem 1       "
            datafld4        =   "DpRem1"
            maxlength4      =   15
            colname5        =   "DPREM2"
            heading5        =   "Rem 2       "
            datafld5        =   "DpRem2"
            maxlength5      =   15
         End
         Begin VB.Label ALBL 
            BackStyle       =   0  'Transparent
            Caption         =   "SEO Password"
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
            Left            =   10020
            TabIndex        =   15
            Top             =   0
            Width           =   1545
         End
      End
   End
   Begin MwfCtl.MWCTL_ADC ADC 
      CausesValidation=   0   'False
      Height          =   555
      Left            =   15
      TabIndex        =   6
      Top             =   9360
      Width           =   14910
      _ExtentX        =   26300
      _ExtentY        =   979
      AllowAdd        =   0   'False
      AllowDelete     =   0   'False
      BtnWidth        =   885
      BtnHeight       =   435
      BtnMaskColor    =   8454143
      FontSz          =   8.25
   End
End
Attribute VB_Name = "EmrFrmDsgPrm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'--- General comments on the entry ----
' X-'', A-DsgPrm
Option Explicit
Dim moCn As MwfLib.MDOConnection
Dim ms_wDpTyp As String
Dim wPic3DHt As Single
Dim wPic3DWdt As Single
Dim ms_PwdFlag As Boolean
Private Sub ADC_Load()
  '*** Set the First Key Control property to the working field wDpTyp
  '*** Set the First Non Key Control property to the grid GrdDsgPrm
  '*** Set the Child Property Of ADC to GrdDsgPrm
  '*** Set the Previous and Next Control Properties for the Grid GrdDsgPrm
  '*** Set the Hot Key for the Grid GrdDsgPrm as 'A'
  '*** initialise the default values for the working fields wDpTyp
  Set moCn = adc.Connection
  Set adc.FirKeyCtl = adc("WDPTYP")
  Set adc.FirNKeyCtl = IIF(adc.MenuCd <> UCase("DsgPrmCat"), TxtPwd, GrdDsgPrm)
  adc.Child = "GrdDsgPrm"
  GrdDsgPrm.Child = "GrdDsgPrmDet"
  Set GrdDsgPrm.PrevCtl = adc
  Set GrdDsgPrm.NextCtl = adc
  Set GrdDsgPrmDet.PrevCtl = GrdDsgPrm
  Set GrdDsgPrmDet.NextCtl = adc
  GrdDsgPrm.HotKey = "A"
  GrdDsgPrmDet.HotKey = "B"
  
  Me.Caption = IIF(adc.MenuCd <> UCase("DsgPrmCat"), "Design Parameters", "Design Catalog")
  If adc.MenuCd <> UCase("DsgPrmCat") Then
    GrdDsgPrm.ColProp("DpDt").Style = fgcinvisible
    GrdDsgPrmDet.ColProp("DpDmNo").Style = fgcinvisible
    GrdDsgPrmDet.ColProp("DpPctNm").Style = fgcinvisible
    GrdDsgPrmDet.ColProp("DpRem1").Style = fgcinvisible
    GrdDsgPrmDet.ColProp("DpRem2").Style = fgcinvisible
  Else
    adc("wDpTyp").Enabled = False
    ALBL(20).Visible = False
    TxtPwd.Enabled = False: TxtPwd.Visible = False
  End If
  ms_wDpTyp = ""
  wPic3DHt = Pic3D.Height
  wPic3DWdt = Pic3D.Width
  
End Sub
Private Sub ADC_KeyWhen()
  '*** set the default values for DpTyp from the previous entry
  If adc.MenuCd = UCase("DsgPrmCat") Then adc("wDpTyp") = "CAT"
  If ms_wDpTyp <> "" Then adc("wDpTyp") = ms_wDpTyp
  TxtPwd.text = ""
End Sub
Private Sub ADC_KeyValid(Cancel As Boolean, ErrMsg As String)
  '*** Validating DpTyp as it is used in Grd Init
  '*** Check whether the DpTyp is valid
  If UCase(adc.MenuCd) <> UCase("DsgPrmCat") Then Call atxt_FldChange(0, UCase("wDpTyp"), Cancel, ErrMsg, adc("wDpTyp"))
  If Cancel = True Then Exit Sub
End Sub

Private Sub ADC_NonKeyValid(pr_Cancel As Boolean, pr_ErrMsg As String)
  If adc.MenuCd = UCase("DsgPrmCat") Then go_HlpCn.Execute ("BEGIN TRY Drop Table #TmpPictFileLst END TRY BEGIN CATCH END CATCH")
End Sub

Private Sub ADC_NonKeyWhen(pr_Cancel As Boolean, pr_ErrMsg As String)
  '*** Store the DpTyp for the next entry
  ms_wDpTyp = adc("wDpTyp")
  If adc.MenuCd <> UCase("DsgPrmCat") Then
    ms_PwdFlag = False:  TxtPwd.text = ""
    GrdDsgPrm.AllowAdd = False: GrdDsgPrm.AllowDelete = False
  ElseIf adc.MenuCd = UCase("DsgPrmCat") Then
    'Call GenFileList   ' ****** Manali 3.5.0 - 09/06/09 - Not reqd now Files Names refresfhed in Head Master
    
    '****** Sachin 3.4.1 - Commented Code below as it fails with Windows 2003(only works in Win 2000) - FSO Used Instead ******
    'go_HlpCn.Execute ("Select Space(100) as PfFileName,0 as Pfdepth,0 as PfFile Into #TmpPictFileLst; " + _
                      "Insert Into #TmpPictFileLst " + _
                      "Exec xp_dirtree '" + wPictPath + "', 0, 1")
  End If
End Sub
Private Sub ATXT_FldWhen(Index As Integer, ByVal IdName As String, Cancel As Boolean, ErrMsg As String)
  '*** Set Helps For Each Field
  Select Case IdName
  Case Is = UCase("wDpTyp")
    Call HlpList.PMCd("DSGPRM")
  Case Is = UCase("wDpCdFr"), UCase("wDpCdTo")
    Call HlpList.DsgPrmCd(adc("wDpTyp"))
  End Select
End Sub
Private Sub atxt_FldChange(Index As Integer, ByVal IdName As String, Cancel As Boolean, ErrMsg As String, ByVal pv_NewValue As Variant)
  '*** wDpTyp       valid PMCd('DSGPRM')
  Select Case IdName
  Case Is = UCase("wDpTyp")
  'pv_NewValue = "CAT" Or
    Cancel = adc.Mode <> xInit And ((adc.MenuCd = UCase("DsgPrmCat") And pv_NewValue <> "CAT") Or (adc.MenuCd <> UCase("DsgPrmCat") And Not moCn.RecSeek("Select PMCd From Param Where PTyp = 'DSGPRM' " + _
                         "And PMCd = '" + pv_NewValue + "'")))
    If Cancel = True Then ErrMsg = "Invalid Design Parameter Type": Exit Sub
  End Select
End Sub

Private Sub GrdDsgPrm_ColChange(ByVal RowNum As Integer, ByVal ColName As String, ByVal NewValue As Variant, Cancel As Boolean, ErrMsg As String)
  Select Case ColName
  Case Is = UCase("DpCd")
     Cancel = (NewValue = "" Or NewValue = 0): ErrMsg = "Parameter Code Can Not Be Blank/ Zero": Exit Sub
  End Select
End Sub

Private Sub GrdDsgPrm_ColWhen(ByVal RowNum As Integer, ByVal ColName As String, Cancel As Boolean, ErrMsg As String, ByVal Action As MwfCtl.en_fgAction)
  '*** Set Helps For Each Field
  If adc.MenuCd <> UCase("DsgPrmCat") And Not ms_PwdFlag Then Cancel = True: ErrMsg = "Enter Password To Edit Parameter Code"
  
  Select Case ColName
  Case Is = UCase("DpCd")
    If Not GrdDsgPrm.IsNew(RowNum) Then
      Cancel = True: ErrMsg = "Design Count For This Param Code: " + CStr(GrdDsgPrmDet.Rows - 1): Exit Sub
    ElseIf (UCase(adc("wDpTyp")) = "SER" Or UCase(adc("wDpTyp")) = "SET" Or UCase(adc("wDpTyp")) = "VER") Then
      Cancel = True: ErrMsg = "Can Not Change Parameter Code": Exit Sub
    End If
  End Select
End Sub

Private Sub GrdDsgPrm_RowWhen(ByVal RowNum As Integer)
  With GrdDsgPrm
    If GrdDsgPrm.IsNew(RowNum) Then
      If (UCase(adc("wDpTyp")) <> "COLL" And UCase(adc("wDpTyp")) <> "CAT") Then
        .Value(.Row, "DpCd") = moCn.AutoGen("DsgPrm", "DpCd", adc("wDpTyp"), -1)
      End If
      .Value(RowNum, "DpDt") = moCn.SrvrDate
    End If
  End With
End Sub

Private Sub GrdDsgPrmDet_ColWhen(ByVal RowNum As Integer, ByVal ColName As String, Cancel As Boolean, ErrMsg As String, ByVal Action As MwfCtl.en_fgAction)
  '*** Set Helps For Each Field
  Select Case ColName
  Case Is = UCase("DpDmCd")
    Call HlpList.DmCd("DM")
  Case Is = UCase("DpPctNm")
    Call HlpList.PictLst(GrdDsgPrmDet.Value(RowNum, "DpDmCd"))
  End Select
End Sub
Private Sub GrdDsgPrmDet_ColChange(ByVal RowNum As Integer, ByVal ColName As String, ByVal NewValue As Variant, Cancel As Boolean, ErrMsg As String)
  With GrdDsgPrmDet
    Select Case ColName
    Case Is = UCase("DpDmCd")
      Cancel = Not moCn.RecSeek("Select 'x' from DsgMst where DmTcTyp='DM' and " + _
                            "DmCd='" + NewValue + "' and DmSz='' and DmHld='N' and DmValidYN='Y' ")
      ErrMsg = "Design Code does not Exist/ On Hold/ Not Valid": Exit Sub
      
    Case Is = UCase("DpDmNo")
      If adc.MenuCd = UCase("DsgPrmCat") And NewValue <= 0 Then
        Cancel = True: ErrMsg = "Design Sr Should Be > 0 ": Exit Sub
      End If
    End Select
  End With
End Sub
Private Sub GrdDsgPrm_InitKey(ByVal RowNum As Integer)
  '*** Initialise Key Fields
  With GrdDsgPrm
    If .IsNew(RowNum) Then
      .Store "DpTyp", adc("wDpTyp")
      .Store "DpNo", 0
      .Store "DpDmCd", ""
      .Store "DpDmNo", 0
      .Store "DpPctNm", ""
      .Store "DpRem1", ""
      .Store "DpRem2", ""
    End If
  End With
End Sub
Private Sub GrdDsgPrmDet_InitKey(ByVal RowNum As Integer)
  '*** Initialise Key Fields
  With GrdDsgPrmDet
    If .IsNew(RowNum) Then
     .Store "DpTyp", adc("wDpTyp")
     .Store "DpCd", GrdDsgPrm.Value(GrdDsgPrm.Row, "DpCd")
     .Store "DpDesc", ""
     .Store "DpDt", MWLib.EmptyDate
     If adc("wDpTyp") <> "CAT" Then
        .Store "DpNo", 0
     Else
        .Store "DpNo", moCn.GetFldVal("Select Max(DpNo) from DsgPrm where DpTyp='" + adc("wDpTyp") + "' and DpCd='" + GrdDsgPrm.Value(GrdDsgPrm.Row, "DpCd") + "'") + 1
      End If
    End If
  End With
End Sub
Private Sub ADC_SetRecSource()
  '*** Since this is an 'A' Entry, there is no RecSource associated with ADC. So set the
      'DataMode property of ADC control as xdatUnbound
  adc.DataMode = xdatUnbound
End Sub
Private Sub GrdDsgPrm_SetRecSource()
  '*** Set the Record Source of the Grid GrdDsgPrm
  Dim ws_Cnd As String, wOrdByStr As String
  wOrdByStr = IIF((UCase(adc("wDpTyp")) = "COLL" Or UCase(adc("wDpTyp")) = "CAT"), " DpCd", " Convert(Int, DpCd)")
  ws_Cnd = adc.RepCond
  ws_Cnd = " where DpDmCd='' " + IIF(ws_Cnd <> "", " and " + ws_Cnd, "")
  GrdDsgPrm.RecSource = "Select * from DsgPrm " + ws_Cnd + " Order By " + wOrdByStr
End Sub

Private Sub GrdDsgPrmDet_LostFocus()
  Call ShowPic(wPic3DHt, wPic3DWdt, "", Pic3D)
End Sub

Private Sub GrdDsgPrmDet_RowWhen(ByVal RowNum As Integer)
  If GrdDsgPrm.Rows > 1 Then
    If GrdDsgPrmDet.IsNew(RowNum) Then
      If adc.MenuCd <> UCase("DsgPrmCat") Then
        GrdDsgPrmDet.Value(RowNum, "DpDmNo") = 0
        GrdDsgPrmDet.Value(RowNum, "DpPctNm") = ""
        GrdDsgPrmDet.Value(RowNum, "DpRem1") = ""
        GrdDsgPrmDet.Value(RowNum, "DpRem2") = ""
      Else
        GrdDsgPrmDet.Value(RowNum, "DpDmNo") = moCn.GetFldVal("Select Max(DpDmNo) from DsgPrm where DpTyp='" + adc("wDpTyp") + "' and DpCd='" + GrdDsgPrm.Value(GrdDsgPrm.Row, "DpCd") + "'") + 1
      End If
    Else
      Dim w3DPth As String
      w3DPth = GetPictPath("DM", False, GrdDsgPrmDet.Value(RowNum, "DpDmCd"), , , , , , , , GrdDsgPrmDet.Value(RowNum, "DpPctNm"))
      Call ShowPic(wPic3DHt, wPic3DWdt, w3DPth, Pic3D)
    End If
  End If
End Sub

Private Sub GrdDsgPrmDet_SetRecSource()
  '*** Set the Record Source of the Grid GrdDsgPrm
  GrdDsgPrmDet.RecSource = "Select * from DsgPrm where " + _
        "DpTyp= '" + adc("wDpTyp") + "' and DpCd='" + GrdDsgPrm.Value(GrdDsgPrm.Row, "DpCd") + "' and DpDmCd<>'' " + _
        "Order By " + IIF(adc("wDpTyp") = "CAT", "DpDmNo", "DpDmCd")
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
  ws_BtnStr = "CmdCpyGo"
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
    Set moCn = Nothing    '*** (09/08/05)  '*** (Bef speed) bef 24/11/06
End Sub
Private Sub GrdDsgPrm_Validate(Cancel As Boolean)
  Cancel = GrdDsgPrm.Validate
End Sub
Private Sub GrdDsgPrmDet_Validate(Cancel As Boolean)
  Cancel = GrdDsgPrmDet.Validate
End Sub
Private Sub TxtPwd_GotFocus()
  DispMsg "Enter SEO Password", etInfo
End Sub
Private Sub TxtPwd_LostFocus()
  Call TxtPwd_Validate(False)
End Sub
Private Sub TxtPwd_Validate(Cancel As Boolean)
  If adc.AddFndMode = xFndMode Then
    If Trim(TxtPwd.text) <> "" Then
      If Trim(TxtPwd.text) = moCn.GetFldVal("select IsNull ((select PValue from Param Join UsrAccess on UaUsrCd = PMCd " + _
                   "where PTyp= 'USR' and PMCd= '" + gs_UsrCd + "' and UaMnuCd = '" + adc.MenuCd + "' and " + _
                   "UaSeoYn= 'Y' and PValue= '" + Trim(TxtPwd.text) + "'), " + _
                   "(Select PValue1 from Param where Ptyp= 'SEOPWD' and PmCd= 'SEOPWD' )) ") Then
        ms_PwdFlag = True
        GrdDsgPrm.AllowAdd = True: GrdDsgPrm.AllowDelete = True
        Exit Sub
      Else
        ms_PwdFlag = False
        GrdDsgPrm.AllowAdd = False: GrdDsgPrm.AllowDelete = False
        Cancel = True: DispMsg "Invalid Edit Password / User Does Not Have SEO Rights ", etError
      End If
    Else
      ms_PwdFlag = False
      GrdDsgPrm.AllowAdd = False: GrdDsgPrm.AllowDelete = False
    End If
  End If
End Sub

Private Sub GenFileList()
 '****** Sachin 3.4.1 - Filling the Picture Files List in Temporary Table  ******
 On Error GoTo errhld
 Dim wfso As New FileSystemObject, wFile As file, wRsDmCtg As MwfLib.MDORowSet, wPctPath As String
  go_HlpCn.Execute ("Select Space(8) as PDmCtg, Space(100) as PfFileName Into #TmpPictFileLst ")
  wPctPath = moCn.GetFldVal("Select PDesc225 from Param Where PTyp='PCT' and PMCd='3D'")
  If Not wfso.FolderExists(wPctPath) Then DispMsg " Folder Does Not Exist ", etInfo:  Exit Sub  '****** Exit Sub if Path does not exist
  Set wRsDmCtg = moCn.OpenRes("Select PMCd From Param Where PTyp='DMCTG'")
    With wRsDmCtg
      Do While Not (.EOF Or .BOF)
        wPctPath = moCn.GetFldVal("Select PDesc225 from Param Where PTyp='PCT' and PMCd='3D'") + "\" + !PMCd
        If wfso.FolderExists(wPctPath) Then
            For Each wFile In wfso.GetFolder(wPctPath).Files
              go_HlpCn.Execute (" Insert Into #TmpPictFileLst Values ('" + !PMCd + "','" + wFile.Name + "') ")
            Next
        End If
        .MoveNext
      Loop
    End With
Exit Sub
errhld:
    DispMsg Err.Description, etError
'****** Sachin 3.4.1 - Filling the Picture Files List in Temporary Table  ******
End Sub
