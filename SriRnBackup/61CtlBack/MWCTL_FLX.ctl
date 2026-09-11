VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "msflxgrd.ocx"
Begin VB.UserControl MWCTL_FLX 
   BackStyle       =   0  'Transparent
   ClientHeight    =   3600
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   5055
   EditAtDesignTime=   -1  'True
   PropertyPages   =   "MWCTL_FLX.ctx":0000
   ScaleHeight     =   3600
   ScaleLeft       =   100
   ScaleMode       =   0  'User
   ScaleWidth      =   5055
   Begin VB.TextBox TxtFnd 
      ForeColor       =   &H80000007&
      Height          =   288
      Left            =   1032
      TabIndex        =   4
      Top             =   624
      Visible         =   0   'False
      Width           =   168
   End
   Begin MwfCtl.MWCTL_MED TXT_FLX 
      Height          =   345
      Left            =   2220
      TabIndex        =   2
      Tag             =   "FG"
      Top             =   510
      Visible         =   0   'False
      Width           =   915
      _ExtentX        =   1614
      _ExtentY        =   609
      MaxLength       =   0
   End
   Begin MSFlexGridLib.MSFlexGrid MSFg 
      CausesValidation=   0   'False
      Height          =   1080
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   3495
      _ExtentX        =   6165
      _ExtentY        =   1905
      _Version        =   393216
      AllowBigSelection=   0   'False
      FormatString    =   " |                              "
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.TextBox TXT_PREV 
      Height          =   300
      Left            =   885
      TabIndex        =   1
      Text            =   "PREV"
      Top             =   705
      Width           =   525
   End
   Begin VB.TextBox TXT_NEXT 
      Height          =   300
      Left            =   1680
      TabIndex        =   3
      Text            =   "NEXT"
      Top             =   690
      Width           =   540
   End
   Begin VB.Menu MWMnu_Flx 
      Caption         =   ""
      Visible         =   0   'False
      Begin VB.Menu Add 
         Caption         =   "Add"
      End
      Begin VB.Menu Del 
         Caption         =   "Delete"
      End
      Begin VB.Menu Ins 
         Caption         =   "Insert"
      End
      Begin VB.Menu Data 
         Caption         =   "Data"
      End
   End
End
Attribute VB_Name = "MWCTL_FLX"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
Dim mo_PrevCtl As Object
Dim mo_NextCtl As Object
Dim mi_PrevRow As Integer
Dim mi_ReCalcRow As Integer     'in let value, rownum
Dim mi_PrevCol As Integer
Dim ms_RecSource As String
Dim mo_Recset As MDORowSet
Dim ms_child As String
Dim mf_AllowAdd As Boolean
Dim mf_AllowDelete As Boolean
Dim mi_PgSize As Integer        'not a property, no. of records
Dim mf_ForceLost As Boolean       'no significance
Dim mf_ReloadDataOnEnd As Boolean
Dim mf_StopRecursiveRecalc As Boolean
Dim mo_BckColorSelRow As OLE_COLOR      '*** Jenny
Dim mb_SelFlag As Boolean               '*** Jenny
Dim ms_StoredProcedure As String  '*** Jenny
Dim ms_ValidateStoredProcedure As String  '*** Jenny
Dim mf_SpecialCopy As Boolean
Dim mo_QueryObject As MDOQuery
Dim mo_QueryObjectValidate As MDOQuery

'??? used in multi form
Dim me_Action As en_fgAction      'tab , shft tab, click, .col , (if shft Tab then no focus to B)
Const ctm_Read = "R"
Const ctm_Write = "W"
Event SetRecSource()
Event InitKey(ByVal RowNum As Integer)
Event ColWhen(ByVal RowNum As Integer, ByVal ColName As String, Cancel As Boolean, ErrMsg As String, ByVal Action As en_fgAction)
Event ColChange(ByVal RowNum As Integer, ByVal ColName As String, ByVal NewValue As Variant, Cancel As Boolean, ErrMsg As String)
Event ColValid(ByVal RowNum As Integer, ByVal ColName As String, ByVal NewValue As Variant, Cancel As Boolean, ErrMsg As String)
Event RowWhen(ByVal RowNum As Integer)
Event OnDelete(ByVal RowNum As Integer, Cancel As Boolean, ErrMsg As String)
Event RowSave(ByVal RowNum As Integer, Cancel As Boolean, ErrMsg As String)
Event RowValid(ByVal RowNum As Integer, Cancel As Boolean, ErrMsg As String)
Event ReCalc(ByVal RowNum As Integer, ByVal ColName As String, ByVal OldVal As Variant, ByVal SourceName As String, ByVal SourceColName As String)
Event When()
Event Valid()
Event BeforeServerCopy()
Event AfterServerCopy()

Dim mc_Columns As New Collection
Dim mf_FrmTxt As Boolean
Dim mf_Val_failed As Boolean
Dim mf_Additem As Boolean
Dim mf_IniProp As Boolean
Dim mf_DataChanged As Boolean
Dim mf_stRowCol As Boolean
Dim mf_stRowColIfOnGrd As Boolean
Dim me_Mode As en_fgMode
Dim mc_AddSave As Collection
Dim mf_SaveByCode As Boolean
Dim mf_AddSave As Boolean
Dim mf_ColValByCode As Boolean
Dim mf_ChkData As Boolean
Dim ml_MaxRowBeforeCopy As Long


'***************** Props Begin ************************************************************'
Public Property Get Row() As Integer
    Row = mi_PrevRow
End Property
Public Property Let Row(ByVal New_Row As Integer)
'??? let row
    mf_SaveByCode = True
    If MSFg.Row = New_Row Then
        Call Whn_Col
        Call whn_row
    Else
        MSFg.Row() = New_Row
    End If
    mf_SaveByCode = False
End Property
Public Sub StartRecursiveRecalc()
  mf_StopRecursiveRecalc = False

End Sub
Public Sub StopRecursiveRecalc()
  mf_StopRecursiveRecalc = True
End Sub


Public Property Get Rows() As Integer
    Rows = MSFg.Rows
End Property
Public Property Let Rows(vNewValue As Integer)
    MSFg.Rows = vNewValue
End Property
Public Property Get Cols() As Integer
    Cols = MSFg.Cols
End Property
Public Property Let Cols(ByVal New_Cols As Integer)
    Dim ws_formatstring As String, i As Integer, wo_reqClsFlx As MWCLS_FLX
    MSFg.Cols() = New_Cols
'???
    ws_formatstring = Space(2) + "|"
    For i = 1 To mc_Columns.Count
        Set wo_reqClsFlx = mc_Columns(i)
        ws_formatstring = ws_formatstring + wo_reqClsFlx.Heading + "|"
    Next i
    FormatString = Mid(ws_formatstring, 1, Len(ws_formatstring) - 1)
End Property
Public Property Get COL() As Integer
    COL = MSFg.COL
End Property
Public Property Let COL(ByVal vNewValue As Integer)
    MSFg.COL() = vNewValue
    me_Action = fgaViaCode
End Property
Public Property Get Enabled() As Boolean
    Enabled = UserControl.Enabled
End Property
Public Property Let Enabled(ByVal vNewValue As Boolean)
    UserControl.Enabled = vNewValue
    PropertyChanged "Enabled"
End Property

'*** Jenny
Public Property Let BckColor(ByVal vNewValue As OLE_COLOR)
    MSFg.BackColor = vNewValue
    PropertyChanged "BckColor"
End Property
Public Property Get BckColor() As OLE_COLOR
    BckColor = MSFg.BackColor
End Property
Public Property Let BckColorBkg(ByVal vNewValue As OLE_COLOR)
    MSFg.BackColorBkg = vNewValue
    PropertyChanged "BckColorBkg"
End Property
Public Property Get BckColorBkg() As OLE_COLOR
    BckColorBkg = MSFg.BackColorBkg
End Property
Public Property Let BckColorFixed(ByVal vNewValue As OLE_COLOR)
    MSFg.BackColorFixed = vNewValue
    PropertyChanged "BckColorBkg"
End Property
Public Property Get BckColorFixed() As OLE_COLOR
    BckColorFixed = MSFg.BackColorFixed
End Property
Public Property Let BckColorSel(ByVal vNewValue As OLE_COLOR)
    MSFg.BackColorSel = vNewValue
    PropertyChanged "BckColorSel"
End Property
Public Property Get BckColorSel() As OLE_COLOR
    BckColorSel = MSFg.BackColorSel
End Property
Public Property Let ForColor(ByVal vNewValue As OLE_COLOR)
    MSFg.ForeColor = vNewValue
    PropertyChanged "ForColor"
End Property
Public Property Get ForColor() As OLE_COLOR
    ForColor = MSFg.ForeColor
End Property
Public Property Let ForColorFixed(ByVal vNewValue As OLE_COLOR)
    MSFg.ForeColorFixed = vNewValue
    PropertyChanged "ForColorFixed"
End Property
Public Property Get ForColorFixed() As OLE_COLOR
    ForColorFixed = MSFg.ForeColorFixed
End Property
Public Property Let ForColorSel(ByVal vNewValue As OLE_COLOR)
    MSFg.ForeColorSel = vNewValue
    PropertyChanged "ForColorSel"
End Property
Public Property Get ForColorSel() As OLE_COLOR
    ForColorSel = MSFg.ForeColorSel
End Property
Public Property Let GrdColor(ByVal vNewValue As OLE_COLOR)
    MSFg.GridColor = vNewValue
    PropertyChanged "GrdColor"
End Property
Public Property Get GrdColor() As OLE_COLOR
    GrdColor = MSFg.GridColor
End Property
Public Property Let GrdColorFixed(ByVal vNewValue As OLE_COLOR)
    MSFg.GridColorFixed = vNewValue
    PropertyChanged "GrdColorFixed"
End Property
Public Property Get GrdColorFixed() As OLE_COLOR
    GrdColorFixed = MSFg.GridColorFixed
End Property
Public Property Let GrdLines(ByVal vNewValue As en_GrdLineStyle)
    MSFg.GridLines = vNewValue
    PropertyChanged "GrdLines"
End Property
Public Property Get GrdLines() As en_GrdLineStyle
    GrdLines = MSFg.GridLines
End Property
Public Property Let GrdLinesFixed(ByVal vNewValue As en_GrdLineStyle)
    MSFg.GridLinesFixed = vNewValue
    PropertyChanged "GrdLinesFixed"
End Property
Public Property Get GrdLinesFixed() As en_GrdLineStyle
    GrdLinesFixed = MSFg.GridLinesFixed
End Property
Public Property Let GrdAppearance(ByVal vNewValue As en_GrdAppearance)
    MSFg.Appearance = vNewValue
    PropertyChanged "GrdAppearance"
End Property
Public Property Get GrdAppearance() As en_GrdAppearance
    GrdAppearance = MSFg.Appearance
End Property


Public Property Let BckColorSelTxt(ByVal vNewValue As OLE_COLOR)
    TXT_FLX.BckColor = vNewValue
    PropertyChanged "BckColorSelTxt"
End Property
Public Property Get BckColorSelTxt() As OLE_COLOR
    BckColorSelTxt = TXT_FLX.BckColor
End Property
Public Property Let ForColorSelTxt(ByVal vNewValue As OLE_COLOR)
    TXT_FLX.ForColor = vNewValue
    PropertyChanged "ForColorSelTxt"
End Property
Public Property Get ForColorSelTxt() As OLE_COLOR
    ForColorSelTxt = TXT_FLX.ForColor
End Property
Public Property Let BckColorSelRow(ByVal vNewValue As OLE_COLOR)
    mo_BckColorSelRow = vNewValue
    PropertyChanged "BckColorSelRow"
End Property
Public Property Get BckColorSelRow() As OLE_COLOR
    BckColorSelRow = mo_BckColorSelRow
End Property
'Public Property Get Pic() As Variant
'    Pic = UserControl.Picture
'End Property
'Public Property Let Pic(ByVal vNewValue As Variant)
'    If CStr(vNewValue) <> "" And Dir(CStr(vNewValue)) <> "" Then
'      UserControl.Picture = LoadPicture(CStr(vNewValue))
'    Else
'      UserControl.Picture = vNewValue
'    End If
'    PropertyChanged "Pic"
'End Property
Public Property Get Pic() As Picture
    Set Pic = UserControl.Picture
End Property
Public Property Let Pic(ByVal vNewValue As Picture)
    Set UserControl.Picture = vNewValue
    PropertyChanged "Pic"
End Property
Public Property Set Pic(ByVal vNewValue As Picture)
   Set UserControl.Picture = vNewValue
   PropertyChanged "Pic"
End Property
'*** Jenny
Public Property Get StoredProcedure() As String
  StoredProcedure = ms_StoredProcedure
End Property
Public Property Let StoredProcedure(ByVal vNewValue As String)
  ms_StoredProcedure = vNewValue
  
  Set mo_QueryObject = Parent.adc.Connection.OpenQuery(vNewValue)
'  Set m_ObjCmd = ADC.Connection.OpenQuery("InsOrderCpy")
End Property
Public Property Get ValidateProcedure() As String
  ValidateProcedure = ms_ValidateStoredProcedure
End Property
Public Property Let ValidateProcedure(ByVal vNewValue As String)
  ms_ValidateStoredProcedure = vNewValue
  Set mo_QueryObjectValidate = Parent.adc.Connection.OpenQuery(ms_ValidateStoredProcedure)
End Property
Public Property Let FormatString(ByVal vNewValue As String)
    MSFg.FormatString = vNewValue
End Property
Public Property Get FormatString() As String
    FormatString = MSFg.FormatString
    PropertyChanged "formatstring"
End Property
Public Property Get ColProp(ByVal pv_ColRef As Variant) As MWCLS_FLX
    Dim wi_Col As Integer
    wi_Col = GetCol(pv_ColRef)
    If wi_Col <= 0 Or wi_Col > mc_Columns.Count Then
        Set ColProp = Nothing
    Else
        Set ColProp = mc_Columns(wi_Col)
    End If
End Property
Public Property Get Value(ByVal pv_RowNum As Integer, ByVal pv_ColRef As Variant, Optional ByRef pr_Cancel As Boolean) As Variant
    Dim wi_Col As Integer, wo_ReqCol As MWCLS_FLX, wv_val As Variant, we_reqDataType As en_FldDataType
    If (me_Mode <> fgmServerCopy) And pv_RowNum <= 0 Or pv_RowNum >= MSFg.Rows Then Err.Raise vbObjectError + 512, "", "Invalid Rownum"
    If IsNumeric(pv_ColRef) Then
          wi_Col = CInt(pv_ColRef)
    Else
          wi_Col = GetCol(pv_ColRef)
    End If
    
    If me_Mode = fgmServerCopy Then
      wv_val = mo_QueryObject.QryParameters(wi_Col)
    Else
      wv_val = MSFg.TextMatrix(pv_RowNum, wi_Col)
    End If
    Set wo_ReqCol = mc_Columns(wi_Col)
    we_reqDataType = wo_ReqCol.DataType
    wv_val = convValue(we_reqDataType, wv_val)
    Set wo_ReqCol = Nothing
    Value = wv_val
End Property
'Public Property Get DelRow(ByVal pv_ColRef As Variant) As Variant
'    On Error GoTo errh
'    Dim wi_Col As Integer, wv_val As Variant, wo_ReqCol As MWCLS_FLX
'    If UBound(mv_delrow) = 0 Then Err.Raise vbObjectError + 512, "", "No Deleted Record"
'    wi_Col = GetCol(pv_ColRef)
'    If wi_Col <= 0 Or wi_Col >= MSFg.Cols Then Err.Raise vbObjectError + 512, "", "Invalid ColNum"
'    wv_val = m_delrow(0, wi_Col)
'    Set wo_ReqCol = mc_Columns(wi_Col)
'    we_reqDataType = wo_ReqCol.DataType
'    If we_reqDataType = fdtFloat Or we_reqDataType = fdtNumber Or we_rqDataType = fdtTimeType Then
'        wv_val = Val(wv_val)
'    ElseIf we_reqDataType = fdtDateType Then
'        If IsDate(wv_val) Then
'            wv_val = CDate(wv_val)
'        Else
'            wv_val = ""
'        End If
'    End If
'    Set wo_ReqCol = Nothing
'    DelRow = wv_val
'End Property
Public Property Let Value(ByVal pv_RowNum As Integer, ByVal pv_ColRef As Variant, Optional ByRef pr_Cancel As Boolean, ByVal pv_vNewValue As Variant)
    On Error GoTo errh
    Dim wi_Col As Integer, wo_ReqCol As MWCLS_FLX, wv_OldVal As Variant, we_reqDataType As en_FldDataType, ws_OvFlwVal As String, ws_reqFormatMask As String, ws_reqColName As String, pr_ErrMsg As String
    If IsNumeric(pv_ColRef) Then
          wi_Col = CInt(pv_ColRef)
    Else
          wi_Col = GetCol(pv_ColRef)
    End If

    If me_Mode = fgmServerCopy Then
      Set wo_ReqCol = mc_Columns(wi_Col)
      we_reqDataType = wo_ReqCol.DataType
      ws_reqFormatMask = wo_ReqCol.FormatMask
      ws_reqColName = wo_ReqCol.ColName
      If we_reqDataType = fdtFloat Or we_reqDataType = fdtNumber Or we_reqDataType = fdtTimeType Then
        ws_OvFlwVal = NumOvrFlwVal(we_reqDataType, ws_reqFormatMask)
    
        pv_vNewValue = convValue(we_reqDataType, pv_vNewValue)
        If we_reqDataType = fdtNumber Or we_reqDataType = fdtFloat Or we_reqDataType = fdtTimeType Then
          pv_vNewValue = Format(MWLib.NVal(pv_vNewValue), ws_reqFormatMask)
        End If
      End If
        mo_QueryObject.QryParameters(wi_Col) = pv_vNewValue
  
      Exit Property
    End If
    
    If wi_Col <= 0 Or wi_Col >= MSFg.Cols Then Err.Raise vbObjectError + 513, "", "Invalid ColNum"
    If pv_RowNum <= 0 Or pv_RowNum >= MSFg.Rows Then Err.Raise vbObjectError + 513, "", "Invalid RowNum"
    
    pr_Cancel = False
    
    Set wo_ReqCol = mc_Columns(wi_Col)
    we_reqDataType = wo_ReqCol.DataType
    ws_reqFormatMask = wo_ReqCol.FormatMask
    ws_reqColName = wo_ReqCol.ColName
 
    wv_OldVal = Value(pv_RowNum, wi_Col)
       
    If we_reqDataType = fdtFloat Or we_reqDataType = fdtNumber Or we_reqDataType = fdtTimeType Then
        ws_OvFlwVal = NumOvrFlwVal(we_reqDataType, ws_reqFormatMask)
        If MSFg.TextMatrix(pv_RowNum, wi_Col) = ws_OvFlwVal Then
            wv_OldVal = ws_OvFlwVal
        End If
    End If
    pv_vNewValue = convValue(we_reqDataType, pv_vNewValue)
    If (pv_vNewValue <> wv_OldVal) Or (wo_ReqCol.inInitMode) Then
        If we_reqDataType = fdtNumber Or we_reqDataType = fdtFloat Or we_reqDataType = fdtTimeType Then
            pv_vNewValue = Format(MWLib.NVal(pv_vNewValue), ws_reqFormatMask)
            Call ChkValidFormats(ws_reqColName, we_reqDataType, ws_reqFormatMask, wo_ReqCol.AllowNegatives, pv_vNewValue, pr_Cancel, pr_ErrMsg)
            If pr_Cancel = True And mi_PrevCol = wi_Col And (Not mf_ColValByCode) Then Err.Raise vbObjectError + 513, "", pr_ErrMsg
            If pr_Cancel = True Then MSFg.TextMatrix(pv_RowNum, wi_Col) = ws_OvFlwVal
        End If
        If pr_Cancel = False Then
          If Not mf_SpecialCopy Then
              RaiseEvent ColChange(pv_RowNum, ws_reqColName, convValue(we_reqDataType, pv_vNewValue), pr_Cancel, pr_ErrMsg)
            End If
            If pr_Cancel = True And mi_PrevCol = wi_Col And (Not mf_ColValByCode) Then Err.Raise vbObjectError + 513, "", pr_ErrMsg
            MSFg.TextMatrix(pv_RowNum, wi_Col) = FormatText(we_reqDataType, ws_reqFormatMask, pv_vNewValue)
        End If
        If pr_Cancel = False Then
            mi_ReCalcRow = pv_RowNum
            If Not mf_StopRecursiveRecalc Then
                Call ReCalcProcV(pv_RowNum, wi_Col, wv_OldVal)
            End If
            If wo_ReqCol.inInitMode = True Then wo_ReqCol.inInitMode = False
        End If
    End If
    Exit Property
errh:
    Call Parent.adc.HandleError
    If EnAndVis And (m_ActScrn.ActiveForm Is UserControl.Parent) Then
        mf_Val_failed = True
        Parent.adc.ValFailed = "col"
        MSFg.COL = mi_PrevCol
        MSFg.Row = mi_PrevRow
        If TXT_FLX.Enabled And TXT_FLX.Visible Then TXT_FLX.SetFocus
    Else
        Parent.MWCTL_HLP1.Visible = False
        Parent.MWCTL_HLP1.Mode = hmHlpNotavlbl
    End If
'??? where to setfocus
'    If m_ActScrn.ActiveControl Is Not Me And EnAndVis Then
 '       Me.SetFocus
  '  End If
End Property
Public Property Let RecSource(ByVal vNewValue As String)
    ms_RecSource = vNewValue
    mf_StopRecursiveRecalc = False
End Property
Public Property Get Child() As String
    Child = ms_child
End Property
Public Property Let Child(ByVal vNewValue As String)
    ms_child = vNewValue
End Property
Public Property Get TxtText() As Variant
    TxtText = TXT_FLX.TxtText
End Property

'Geeta****
Public Property Get MaxLength() As Integer
    MaxLength = TXT_FLX.MaxLength
End Property
Public Property Get Alignment() As en_TxtAlign
    Alignment = TXT_FLX.Alignment
End Property
'Geeta****

Public Property Let TxtText(vNewValue As Variant)
    TXT_FLX.TxtText = vNewValue
End Property
Public Property Get DataType() As en_FldDataType
    DataType = TXT_FLX.DataType
End Property
Public Property Get TxtIsBlank() As Boolean
    TxtIsBlank = TXT_FLX.TxtIsBlank
End Property
Public Property Get ObjNm() As String
    ObjNm = UserControl.Extender.Name
End Property
Property Get ParentRef() As Object
    Set ParentRef = Parent
End Property
Public Property Get Top() As Long
    Top = TXT_FLX.Top
End Property
Public Property Get Left() As Long
    Left = TXT_FLX.Left
End Property
Public Property Get Width() As Long
    Width = TXT_FLX.Width
End Property
Public Property Get Height() As Long
    Height = TXT_FLX.Height
End Property
Public Property Get Mode() As en_fgMode
    Mode = me_Mode
End Property
Public Property Get IsNew(ByVal pv_RowNum As Integer) As Boolean
    IsNew = IIf(MSFg.RowData(pv_RowNum) = 0, True, False)
End Property
Public Property Get IsDel(ByVal pv_RowNum As Integer) As Boolean
    IsDel = IIf(MSFg.RowData(pv_RowNum) = -1, True, False)
End Property
Public Property Set PrevCtl(ByVal vNewValue As Object)
    Set mo_PrevCtl = vNewValue
End Property
Public Property Set NextCtl(ByVal vNewValue As Object)
    Set mo_NextCtl = vNewValue
End Property
'Public Property Let ColWidth(ByVal pv_ColNum As Integer, ByVal vNewValue As Long)
'    MSFg.ColWidth(pv_ColNum) = vNewValue
'End Property
'Public Property Get ColWidth(ByVal ColNum As Integer) As Long
'    ColWidth = MSFg.ColWidth(ColNum)
'End Property
Public Property Get EnAndVis() As Boolean
    Dim wobj As Variant
    If Not (Enabled Or UserControl.Extender.Visible) Then
        EnAndVis = False
        Exit Function
    End If
    EnAndVis = True
    Set wobj = UserControl.Extender.Container
    Do While True
        If wobj.Name = ParentRef.Name Then
            Exit Do
        End If
        If (Not wobj.Enabled) Or (Not wobj.Visible) Then
            EnAndVis = False
            Exit Do
        Else
            Set wobj = wobj.Container
        End If
    Loop
End Property
Public Property Get AllowAdd() As Boolean
    AllowAdd = mf_AllowAdd
End Property
Public Property Let AllowAdd(ByVal vNewValue As Boolean)
    mf_AllowAdd = vNewValue
    PropertyChanged "AllowAdd"
End Property
Public Property Get AllowDelete() As Boolean
    AllowDelete = mf_AllowDelete
End Property
Public Property Let AllowDelete(ByVal vNewValue As Boolean)
    mf_AllowDelete = vNewValue
    PropertyChanged "AllowDelete"
End Property
Private Sub MSFg_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Y > MSFg.RowHeight(0) Then
        MSFg.ToolTipText = ""
        Exit Sub
    End If
    Dim wtot As Integer, i As Integer
    wtot = 0
    For i = 0 To MSFg.Cols - 1
        If X > wtot And X <= wtot + MSFg.ColWidth(i) Then
            If i > 0 And i < mc_Columns.Count - 1 Then MSFg.ToolTipText = mc_Columns(i).ToolTipText
            Exit Sub
        End If
        wtot = wtot + MSFg.ColWidth(i)
    Next i
End Sub
Public Property Let HotKey(vNewValue As String)
    UserControl.AccessKeys = vNewValue
End Property
Private Sub TxtFnd_GotFocus()
    mf_FrmTxt = False
End Sub

'***************** Props End ************************************************************'
'***************** Events Begin ************************************************************'
Private Sub UserControl_InitProperties()
    MSFg.Cols = 1
End Sub
Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
    Dim ws_formatstring As String, wi_Ctr As Integer, wo_reqClsFlx As MWCLS_FLX, ws_ReqStr As String
    mb_SelFlag = False    '*** Jenny
    MSFg.Font.Size = 10
    If Ambient.UserMode Then MSFg.Rows = 1
    mf_SaveByCode = False
    mf_AddSave = False
    mf_FrmTxt = False
    mf_Val_failed = False
    mf_ChkData = False
    mf_Additem = False
    mf_IniProp = True
    me_Mode = fgmnorm
    MSFg.Cols = PropBag.ReadProperty("Cols", 1)
    Enabled = PropBag.ReadProperty("Enabled", True)
    
    '*** Jenny
    BckColor = PropBag.ReadProperty("BckColor", &H80000005)
    BckColorBkg = PropBag.ReadProperty("BckColorBkg", &H808080)
    BckColorFixed = PropBag.ReadProperty("BckColorFixed", &H8000000F)
    BckColorSel = PropBag.ReadProperty("BckColorSel", &H8000000D)
    ForColor = PropBag.ReadProperty("ForColor", &H80000008)
    ForColorFixed = PropBag.ReadProperty("ForColorFixed", &H80000012)
    ForColorSel = PropBag.ReadProperty("ForColorSel", &H8000000E)
    GrdColor = PropBag.ReadProperty("GrdColor", &HC0C0C0)
    GrdColorFixed = PropBag.ReadProperty("GrdColorFixed", &H0&)
    GrdLines = PropBag.ReadProperty("GrdLines", flxGrdFlat)
    GrdLinesFixed = PropBag.ReadProperty("GrdLinesFixed", flxGrdInset)
    GrdAppearance = PropBag.ReadProperty("GrdAppearance", flxGrdApp3D)
    
    '*** (Bef 24/03/06)
    'BckColorSelTxt = PropBag.ReadProperty("BckColorSelTxt", &H8000000D)
    'ForColorSelTxt = PropBag.ReadProperty("ForColorSelTxt", &H8000000E)
    '*** (Bef 24/03/06)

    '*** (24/03/06)
    BckColorSelTxt = PropBag.ReadProperty("BckColorSelTxt", &H80000005)
    ForColorSelTxt = PropBag.ReadProperty("ForColorSelTxt", &H80000008)
    '*** (24/03/06)

    mo_BckColorSelRow = PropBag.ReadProperty("BckColorSelRow", &H80000005)
    Pic = PropBag.ReadProperty("Pic", UserControl.Picture)
    '*** Jenny
    
    mf_AllowAdd = PropBag.ReadProperty("AllowAdd", True)
    mf_AllowDelete = PropBag.ReadProperty("AllowDelete", True)
    ws_formatstring = Space(2) + "|"
    MSFg.Row = 0              '*** (Jenny)   Set the current Row as 0 (heading) for setting the font properties of heading
    For wi_Ctr = 1 To MSFg.Cols - 1
      Set wo_reqClsFlx = New MWCLS_FLX
      Set wo_reqClsFlx.Parent = Me
      mc_Columns.Add wo_reqClsFlx
      With wo_reqClsFlx
            ws_ReqStr = "colname" + CStr(wi_Ctr)
            .ColName = PropBag.ReadProperty(ws_ReqStr, "")
            ws_ReqStr = "heading" + CStr(wi_Ctr)
            .Heading = PropBag.ReadProperty(ws_ReqStr, "")
            ws_formatstring = ws_formatstring + .Heading + "|"
            ws_ReqStr = "datafld" + CStr(wi_Ctr)
            .DataFld = PropBag.ReadProperty(ws_ReqStr, "")
            ws_ReqStr = "datatype" + CStr(wi_Ctr)
            .DataType = PropBag.ReadProperty(ws_ReqStr, 0)
            ws_ReqStr = "mask" + CStr(wi_Ctr)
            .Mask = PropBag.ReadProperty(ws_ReqStr, "")
            ws_ReqStr = "recalcon" + CStr(wi_Ctr)
            .ReCalcOn = PropBag.ReadProperty(ws_ReqStr, "")
            ws_ReqStr = "recalcparent" + CStr(wi_Ctr)
            .ReCalcParent = PropBag.ReadProperty(ws_ReqStr, "")
            ws_ReqStr = "maxlength" + CStr(wi_Ctr)
            mc_Columns(wi_Ctr).MaxLength = PropBag.ReadProperty(ws_ReqStr, 0)
            ws_ReqStr = "style" + CStr(wi_Ctr)
            mc_Columns(wi_Ctr).Style = PropBag.ReadProperty(ws_ReqStr, fgcNormal)
            ws_ReqStr = "tooltiptext" + CStr(wi_Ctr)
            .ToolTipText = PropBag.ReadProperty(ws_ReqStr, "")
            Set wo_reqClsFlx = Nothing
            
            '*** (Jenny)
            MSFg.COL = wi_Ctr             '*** Set the current Column as wi_Ctr
            MSFg.CellFontBold = True      '*** CellFontBold affects only the currently active cell and once set will remain set.
            'MSFg.CellFontName = "Arial"
            'MSFg.CellFontSize = 9
            MSFg.CellFontSize = 8
            '*** (Jenny)
        End With
    Next wi_Ctr
    FormatString = Mid(ws_formatstring, 1, Len(ws_formatstring) - 1)
    If Ambient.UserMode Then
        For wi_Ctr = 1 To MSFg.Cols - 1
            Set wo_reqClsFlx = mc_Columns(wi_Ctr)
            wo_reqClsFlx.Width = MSFg.ColWidth(wi_Ctr)
            If wo_reqClsFlx.Style = fgcinvisible Then
                MSFg.ColWidth(wi_Ctr) = 0
            End If
        Next
    End If
    Set wo_reqClsFlx = Nothing
    me_Action = fgaOthers
    mf_stRowCol = False
    mf_IniProp = False
    MSFg.FillStyle = flexFillSingle
    mi_PgSize = (UserControl.Extender.Height / MSFg.RowHeight(0)) - 1
End Sub
Private Sub UserControl_Resize()
    MSFg.Height = UserControl.Height
    MSFg.Width = UserControl.Width
End Sub
Private Sub UserControl_WriteProperties(PropBag As PropertyBag)
    Dim wi_Ctr As Integer, ws_ReqStr As String, wo_reqClsFlx
    Call PropBag.WriteProperty("Cols", MSFg.Cols, 1)
    Call PropBag.WriteProperty("Enabled", Enabled, True)
    Call PropBag.WriteProperty("AllowAdd", mf_AllowAdd, True)
    Call PropBag.WriteProperty("AllowDelete", mf_AllowDelete, True)
    
    '*** Jenny
    Call PropBag.WriteProperty("BckColor", MSFg.BackColor, &H80000005)
    Call PropBag.WriteProperty("BckColorBkg", MSFg.BackColorBkg, &H808080)
    Call PropBag.WriteProperty("BckColorFixed", MSFg.BackColorFixed, &H8000000F)
    Call PropBag.WriteProperty("BckColorSel", MSFg.BackColorSel, &H8000000D)
    Call PropBag.WriteProperty("ForColor", MSFg.ForeColor, &H80000008)
    Call PropBag.WriteProperty("ForColorFixed", MSFg.ForeColorFixed, &H80000012)
    Call PropBag.WriteProperty("ForColorSel", MSFg.ForeColorSel, &H8000000E)
    Call PropBag.WriteProperty("GrdColor", MSFg.GridColor, &HC0C0C0)
    Call PropBag.WriteProperty("GrdColorFixed", MSFg.GridColorFixed, &H0&)
    Call PropBag.WriteProperty("GrdLines", MSFg.GridLines, flxGrdFlat)
    Call PropBag.WriteProperty("GrdLinesFixed", MSFg.GridLinesFixed, flxGrdInset)
    Call PropBag.WriteProperty("GrdAppearance", MSFg.Appearance, flxGrdApp3D)
    
    '*** (Bef 24/03/06)
    'Call PropBag.WriteProperty("BckColorSelTxt", TXT_FLX.BckColor, &H8000000D)
    'Call PropBag.WriteProperty("ForColorSelTxt", TXT_FLX.ForColor, &H8000000E)
    '*** (Bef 24/03/06)
    
    '*** (24/03/06)
    Call PropBag.WriteProperty("BckColorSelTxt", TXT_FLX.BckColor, &H80000005)
    Call PropBag.WriteProperty("ForColorSelTxt", TXT_FLX.ForColor, &H80000008)
    '*** (24/03/06)
    
    Call PropBag.WriteProperty("BckColorSelRow", mo_BckColorSelRow, &H80000005)
    Call PropBag.WriteProperty("Pic", UserControl.Picture, LoadPicture())
    '*** Jenny
    
    For wi_Ctr = 1 To MSFg.Cols - 1
        Set wo_reqClsFlx = mc_Columns(wi_Ctr)
        With wo_reqClsFlx
            ws_ReqStr = "colname" + CStr(wi_Ctr)
            Call PropBag.WriteProperty(ws_ReqStr, .ColName, "")
            ws_ReqStr = "heading" + CStr(wi_Ctr)
            Call PropBag.WriteProperty(ws_ReqStr, .Heading, "")
            ws_ReqStr = "datafld" + CStr(wi_Ctr)
            Call PropBag.WriteProperty(ws_ReqStr, .DataFld, "")
            ws_ReqStr = "datatype" + CStr(wi_Ctr)
            Call PropBag.WriteProperty(ws_ReqStr, .DataType, 0)
            ws_ReqStr = "mask" + CStr(wi_Ctr)
            Call PropBag.WriteProperty(ws_ReqStr, .Mask, "")
            ws_ReqStr = "recalcon" + CStr(wi_Ctr)
            Call PropBag.WriteProperty(ws_ReqStr, .ReCalcOn, "")
            ws_ReqStr = "recalcparent" + CStr(wi_Ctr)
            Call PropBag.WriteProperty(ws_ReqStr, .ReCalcParent, "")
            ws_ReqStr = "maxlength" + CStr(wi_Ctr)
            Call PropBag.WriteProperty(ws_ReqStr, .MaxLength, 0)
            ws_ReqStr = "style" + CStr(wi_Ctr)
            Call PropBag.WriteProperty(ws_ReqStr, .Style, fgcNormal)
            ws_ReqStr = "tooltiptext" + CStr(wi_Ctr)
            Call PropBag.WriteProperty(ws_ReqStr, .ToolTipText, "")
        End With
    Next wi_Ctr
    Set wo_reqClsFlx = Nothing
End Sub
Private Sub Add_Click()
    Dim wo_reqClsFlx As MWCLS_FLX
    Cols = Cols + 1
    Set wo_reqClsFlx = New MWCLS_FLX
    Set wo_reqClsFlx.Parent = Me
    mc_Columns.Add wo_reqClsFlx
    Set wo_reqClsFlx = Nothing
End Sub
Private Sub Data_Click()
    Dim wo_reqClsFlx As MWCLS_FLX, ws_Mask As String, wi_DecPos As Integer, ws_reqTxt As String, i As Integer
    MSFg.Redraw = False
    For i = 1 To MSFg.Cols - 1
        Set wo_reqClsFlx = mc_Columns(i)
        With wo_reqClsFlx
            Select Case .DataType
            Case Is = fdtFloat, fdtTimeType
                ws_Mask = .Mask
                wi_DecPos = InStr(1, ws_Mask, ".")
                ws_reqTxt = String(wi_DecPos - 1, "9") + "." + String(Len(ws_Mask) - wi_DecPos, "9")
            Case Is = fdtNumber
                ws_reqTxt = String(.MaxLength, "9")
            Case Is = fdtDateType
                ws_reqTxt = "12/12/12"
            Case Else
                ws_reqTxt = String(.MaxLength, "X")
            End Select
            MSFg.TextMatrix(1, i) = ws_reqTxt
        End With
    Next i
    MSFg.Redraw = True
End Sub
Private Sub Del_Click()
    Dim wo_reqClsFlx As MWCLS_FLX, ws_RecalcOn As String, ws_RecalcParent As String
    Set wo_reqClsFlx = mc_Columns(MSFg.COL)
    With wo_reqClsFlx
        ws_RecalcOn = .ReCalcOn
        ws_RecalcParent = .ReCalcParent
    End With
'    If ws_RecalcOn <> "" Or ws_RecalcParent <> "" Then
'        ErrMsg = IIf(wrecalcon <> "", "Fld Recalculated on change of " + wrecalcon, "")
'        ErrMsg = ErrMsg + IIf(wrecalcparent <> "", "Triggering Recalc of " + wrecalcparent, "")
'        MsgBox "Cannot Delete as  " + ErrMsg
'        Exit Sub
'    End If
'??? ErrHndling at Design Time
    Set wo_reqClsFlx = Nothing
    mc_Columns.Remove (MSFg.COL)
    Cols = Cols - 1
End Sub
Private Sub Ins_Click()
    Dim wo_reqClsFlx As MWCLS_FLX
    Set wo_reqClsFlx = New MWCLS_FLX
    Set wo_reqClsFlx.Parent = Me
    mc_Columns.Add wo_reqClsFlx, , MSFg.COL
    Set wo_reqClsFlx = Nothing
    Cols = Cols + 1
End Sub
Private Sub msfg_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Dim wo_reqClsFlx As MWCLS_FLX
    If MSFg.COL <= 0 Or MSFg.COL >= MSFg.Cols Then Exit Sub
    If Not Ambient.UserMode And Button = vbRightButton And Shift = 0 Then
        PopupMenu MWMnu_Flx
    End If
    If Ambient.UserMode And Button = vbRightButton And Shift = 0 Then
        Call Val_Col
        If mf_Val_failed Then
            mf_Val_failed = False
            mf_Additem = False
            Exit Sub
        End If
        Call val_row
        If mf_Val_failed Then
            mf_Val_failed = False
            mf_Additem = False
            Exit Sub
        End If
        Set wo_reqClsFlx = mc_Columns(MSFg.COL)
        With TxtFnd
          .MaxLength = wo_reqClsFlx.MaxLength
          .Visible = True
          .Text = ""
          .Left = MSFg.CellLeft + MSFg.Left
          .Top = MSFg.Top
          .Height = MSFg.CellHeight
          .Width = MSFg.CellWidth
          .SetFocus
        End With
        Set wo_reqClsFlx = Nothing
    End If
End Sub
Private Sub MSFg_SelChange()
    '*** (Bef 13/09/05)  Jenny  (This Code Is Required To Give Background Color To The Selected Row. But It Makes Scrolling Slow)
    ''*** Jenny
    'If MSFg.Row <> MSFg.RowSel Then MSFg.RowSel = MSFg.Row
    'If mb_SelFlag = False Then
    '  If MSFg.COL <> MSFg.ColSel Then MSFg.ColSel = MSFg.COL
    'End If
    '*** Jenny
    '*** (Bef 13/09/05)  Jenny  (This Code Is Required To Give Background Color To The Selected Row. But It Makes Scrolling Slow)

    '*** Bef 07/02/05  (Original Code)
    If MSFg.Row <> MSFg.RowSel Then MSFg.RowSel = MSFg.Row
    If MSFg.COL <> MSFg.ColSel Then MSFg.ColSel = MSFg.COL
    '*** Bef 07/02/05  (Original Code)
    ' DO NOT REMOVE THIS IF
    ' SELCHANGE CALLED ITERATIVELY WHEN
    ' ONLY ONE ROW IN GRID;SETG MSFG.ROWSEL=MSFG.ROW ie equal to 1
    ' /// with only one column in grd;setg msfg.colsel=msfg.col ie equal to 1
End Sub
Private Sub TXT_FLX_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyUp And MSFg.Row > 1 Then MSFg.Row = MSFg.Row - 1
    If KeyCode = vbKeyDown Then
        If MSFg.Row < MSFg.Rows - 1 Then MSFg.Row = MSFg.Row + 1
    End If
    If KeyCode = vbKeyPageUp Then MSFg.Row = IIf(MSFg.Row - mi_PgSize >= 1, MSFg.Row - mi_PgSize, 1)
    If KeyCode = vbKeyPageDown Then MSFg.Row = IIf(MSFg.Row + mi_PgSize <= MSFg.Rows - 1, MSFg.Row + mi_PgSize, MSFg.Rows - 1)
    If Shift = 2 And KeyCode = vbKeyW Then SendToNextCtl
    If Shift = 2 And KeyCode = vbKeyQ Then SendToPrevCtl
    If mf_AllowDelete And (Shift = 2 And KeyCode = vbKeyT) Then
        If Not InSelMode Then Call Del_Rec(True)
    End If
    If mf_AllowAdd And (Shift = 2 And KeyCode = 78) Then
        If Not InSelMode Then
            mf_Additem = True
            Call Val_Col
            If mf_Val_failed Then
                mf_Val_failed = False
                mf_Additem = False
                Exit Sub
            End If
            Call val_row
            If mf_Val_failed Then
                mf_Val_failed = False
                mf_Additem = False
                Exit Sub
            End If
        End If
    End If
End Sub
Private Sub TXT_FLX_KeyUp(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn And Shift = 2 Then
        Call ReCalc(mc_Columns(MSFg.COL).ColName, "", "ZUDEF", "ZUDEF", mi_PrevRow)
        TXT_FLX.TxtText = MSFg
    End If
End Sub
Private Sub TXT_NEXT_GotFocus()
    Dim wi_NxtCol As Integer, i As Integer, wo_reqClsFlx As MWCLS_FLX
    mf_FrmTxt = False
    wi_NxtCol = MSFg.Cols
    For i = MSFg.COL + 1 To mc_Columns.Count
        Set wo_reqClsFlx = mc_Columns(i)
         If wo_reqClsFlx.Style = fgcNormal Then
            wi_NxtCol = i
            Exit For
         End If
     Next i
     Set wo_reqClsFlx = Nothing
     If wi_NxtCol >= MSFg.Cols Then
         If TXT_FLX.Visible And TXT_FLX.Enabled Then TXT_FLX.SetFocus
     Else
         me_Action = fgaTabFrmCol
         MSFg.COL = wi_NxtCol
     End If
End Sub
Private Sub TXT_PREV_GotFocus()
    Dim mi_PrvCol As Integer, wo_reqClsFlx As MWCLS_FLX, i As Integer
    mf_FrmTxt = False
    mi_PrvCol = 0
    For i = MSFg.COL - 1 To 1 Step -1
        Set wo_reqClsFlx = mc_Columns(i)
        If wo_reqClsFlx.Style = fgcNormal Then
            mi_PrvCol = i
            Exit For
        End If
    Next i
    If mi_PrvCol <= 0 Then
        If TXT_FLX.Visible And TXT_FLX.Enabled Then TXT_FLX.SetFocus
    Else
        me_Action = fgaShiftTabFrmCol
        MSFg.COL = mi_PrvCol
    End If
End Sub
Private Sub TxtFnd_LostFocus()
    mf_FrmTxt = False
    Dim i As Integer, wo_reqClsFlx As MWCLS_FLX, wv_GrdVal As Variant, wv_reqVal As Variant
    Dim we_reqDataType As en_FldDataType
    MSFg.SetFocus
    TxtFnd.Visible = False
    Set wo_reqClsFlx = mc_Columns(MSFg.ColSel)
    we_reqDataType = wo_reqClsFlx.DataType
    Set wo_reqClsFlx = Nothing
    If TxtFnd <> "" Then
        For i = MSFg.RowSel + 1 To MSFg.Rows
            wv_GrdVal = Value(i - 1, MSFg.ColSel)
            wv_reqVal = convValue(we_reqDataType, TxtFnd)
            'If MSFg.TextMatrix(i - 1, MSFg.ColSel) = TxtFnd Then
            If we_reqDataType = fdtCharType Or we_reqDataType = fdtConvertToUpper Then
                wv_GrdVal = UCase(wv_GrdVal)
                wv_reqVal = UCase(wv_reqVal)
            End If
            If wv_GrdVal = wv_reqVal Then
                MSFg.Row = i - 1
                Exit Sub
            End If
        Next i
    End If
End Sub
Private Sub MSFG_RowColChange()
    If Not Ambient.UserMode Then Exit Sub
    If mf_stRowCol Then Exit Sub
    If mf_stRowColIfOnGrd Then
        If Not (UserControl.ActiveControl Is Nothing) Then
            If UCase(UserControl.ActiveControl.Name) = UCase("Msfg") Then
                Exit Sub
            End If
        End If
    End If
    If MSFg.COL > 0 And MSFg.COL < MSFg.Cols Then
        If mc_Columns(MSFg.COL).Style = fgcinvisible Then
            mf_stRowCol = True
            MSFg.COL = mi_PrevCol
            MSFg.Row = mi_PrevRow
            mf_stRowCol = False
        End If
    End If
    If Parent.MWCTL_HLP1.Mode = hmHlpBrowsing Then
        If TXT_FLX.Visible Then
            mf_stRowCol = True
            MSFg.Row = mi_PrevRow
            MSFg.COL = mi_PrevCol
            TXT_FLX.SetFocus
            mf_stRowCol = False
            Exit Sub
        End If
    End If
    If Not (mf_Val_failed Or mf_Additem) Then
        If mi_PrevCol <> MSFg.COL Then Call Val_Col
        If mf_Val_failed Then
            mf_Val_failed = False
            mf_Additem = False
            Exit Sub
        End If
        
        If mi_PrevRow <> MSFg.Row Then
            If mi_PrevCol = MSFg.COL Then Call Val_Col
            If mf_Val_failed Then
                mf_Val_failed = False
                mf_Additem = False
                Exit Sub
            End If
            
            Call val_row
            If mf_Val_failed Then
                mf_Val_failed = False
                mf_Additem = False
                Exit Sub
            End If
            
            Call whn_row
            If mi_PrevCol = MSFg.COL Then Call Whn_Col
        End If
        If mi_PrevCol <> MSFg.COL Then Call Whn_Col
    End If
End Sub
Private Sub MSFg_GotFocus()
    On Error GoTo errh
    Dim wf_XSaveFailed As Boolean
    If Not mf_FrmTxt Then
        If Ambient.UserMode Then
            If Parent.MWCTL_HLP1.Mode = hmHlpBrowsing Then
                If TXT_FLX.Visible Then TXT_FLX.SetFocus
                Exit Sub
            End If
            If (Not mf_Val_failed) And Parent.adc.ValFailed = "" Then
                If Parent.adc.XChild(UserControl.Extender.Name) Then
                    wf_XSaveFailed = True
                    If Not Parent.adc.SaveRec Then Err.Raise vbObjectError + 514, 0, "X Save Failed"
                    wf_XSaveFailed = False
                End If
                If mf_AllowAdd And MSFg.Rows = 1 Then
                    If MSFg.Cols > 1 Then MSFg.COL = FirstCol
                    Call Additem
                End If
                mf_stRowCol = True
                If MSFg.Row = 0 And MSFg.Rows > 1 Then MSFg.Row = 1
                If MSFg.COL = 0 And MSFg.Cols > 1 Then MSFg.COL = FirstCol()
                If MSFg.COL > 0 And MSFg.COL < MSFg.Cols Then
                  If mc_Columns(MSFg.COL).Style = fgcinvisible And MSFg.Cols > 1 Then
                    MSFg.COL = FirstCol()
                  End If
                End If
                mf_stRowCol = False
                'sri
                  If Not Mode = fgmServerCopy Then
                      RaiseEvent When
                  End If
                Call ChildsEnaDisa(True)
                Call whn_row
                Call Whn_Col
                mf_stRowColIfOnGrd = False
            End If
        End If
    End If
    If TXT_FLX.Visible And TXT_FLX.Enabled Then
        TXT_FLX.SetFocus
    End If
    mf_FrmTxt = False
    Exit Sub
errh:
    Call Parent.adc.HandleError
    If wf_XSaveFailed Then Parent.adc.SetFocus
'??? otherwise no setfocus
End Sub
Private Sub TXT_FLX_LostFocus()
    If m_ActScrn.ActiveControl Is Nothing Then Exit Sub
    If (Not mf_ForceLost) And (Not (UserControl.Parent Is m_ActScrn.ActiveForm)) Then Exit Sub
    If TypeOf m_ActScrn.ActiveControl Is MWCTL_HLP Then
        Parent.MWCTL_HLP1.Mode = hmHlpBrowsing
    End If
    Dim ROBJ As Object
    Set ROBJ = m_ActScrn.ActiveControl
    If Parent.MWCTL_HLP1.Mode = hmHlpBrowsing Then
        mf_FrmTxt = True
        Exit Sub
    End If
    If ROBJ.Name = UserControl.Extender.Name Then
         If Not (UserControl.ActiveControl Is Nothing) Then
'             If UserControl.ActiveControl.Name <> "TXT_FLX" Then
'    to avoid happening of grd when while click on grid from txt_flx
                mf_FrmTxt = True
'             End If
         Else
            mf_FrmTxt = True
         End If
         Exit Sub
    End If
    If ROBJ.Name = "ADC" Then
        If (ROBJ.OptSel = "C") Then
            Parent.MWCTL_HLP1.Visible = False
            Parent.MWCTL_HLP1.Mode = hmHlpNotavlbl
            TXT_FLX.Visible = False
            mf_stRowColIfOnGrd = True
            Parent.adc.ValFailed = ""
            Exit Sub
        ElseIf ROBJ.OptSel = "S" Or ROBJ.OptSel = "D" Then
            If Validate Then Exit Sub
        End If
    End If
    TXT_FLX.Visible = False
    mf_stRowColIfOnGrd = True
End Sub
Private Sub MSFG_Scroll()
    If Ambient.UserMode Then TXT_FLX.Visible = False
End Sub
'***************** Events End ************************************************************'
'***************** Methods Begin ************************************************************'
Public Function FormatText(ByVal pv_DataType As en_FldDataType, ByVal pv_FormatMask As String, ByVal pv_vNewValue As Variant)
    Select Case pv_DataType
    Case Is = fdtNumber, fdtFloat, fdtTimeType
        FormatText = Format(Val(pv_vNewValue), pv_FormatMask)
    Case Is = fdtDateType
        FormatText = Format(pv_vNewValue, "dd/mm/yy")
    Case Else
        FormatText = pv_vNewValue
    End Select
End Function
Public Function Del_Rec(ByVal pv_UserDel As Boolean, Optional ByVal pv_RowNum As Variant) As Boolean
    Dim pr_Cancel As Boolean, pr_ErrMsg As String, wl_DelBkMrk As Double, i As Integer
    If mo_Recset Is Nothing Then Exit Function
    On Error GoTo errh

' ########################  Manoj #### Ver: 2.0.6 #### Date: 16/06/2004  ########################
' check for Login session is Working or Not
    If Not ChkLogin(Parent.adc) Then Exit Function
' ###########################################  Manoj  ###########################################
    
    Parent.adc.ValFailed = ""
    If IsMissing(pv_RowNum) Then pv_RowNum = MSFg.Row
    
    wl_DelBkMrk = MSFg.RowData(pv_RowNum)
    
    '??? should it be .row or mi_prevrow
    'raiseerror if pv_rownum is not within permitted limits
    pr_Cancel = False
    pr_ErrMsg = ""
    If Not IsNew(pv_RowNum) Then RaiseEvent OnDelete(pv_RowNum, pr_Cancel, pr_ErrMsg)
    
    If pr_Cancel = True Then Err.Raise vbObjectError + 512, "", pr_ErrMsg
    
    me_Mode = fgmdel

    MSFg.RowData(pv_RowNum) = -1
    
    Set mc_AddSave = Nothing
    ' to remove addsave put becos of recalcs occured before this while editing the rec
    mi_ReCalcRow = pv_RowNum
    For i = 1 To MSFg.Cols - 1
        Call ReCalcProc(i, 0)
    Next i
    Parent.adc.Connection.BeginTrans (mdoCommitNone)
    If Not AddSaveProc Then Err.Raise vbObjectError + 514, 0, "AddSave Failed"
    If wl_DelBkMrk <> 0 Then
        mo_Recset.Bookmark = wl_DelBkMrk
        mo_Recset.Delete
    End If
    Parent.adc.Connection.CommitTrans (mdoCommitNone)
    Set mc_AddSave = Nothing
    Parent.MWCTL_HLP1.Mode = hmHlpNotavlbl
    If pv_RowNum = 1 And MSFg.Rows = 2 Then
        MSFg.Rows = 1
        TXT_FLX.Visible = False
    Else
        MSFg.RemoveItem (pv_RowNum)
    End If
    
    If wl_DelBkMrk = 0 Then Call SetColsInitMode(False)
    
    me_Mode = fgmnorm
    If MSFg.Rows = 1 Then
        If pv_UserDel Then SendToNextCtl
        Call Init
        Call ChildsEnaDisa(False)
    Else
        If pv_UserDel Then
            Call whn_row
            Call Whn_Col
        End If
    End If
    If Parent.MWCTL_HLP1.Visible Then Parent.MWCTL_HLP1.Visible = False
    Del_Rec = True
    DispMsg "", etInfo
    Exit Function
errh:
    Call Parent.adc.HandleError
    Set mc_AddSave = Nothing
    mo_Recset.CancelUpdate
    Del_Rec = False
    If TXT_FLX.Visible And TXT_FLX.Enabled Then TXT_FLX.SetFocus
    MSFg.RowData(pv_RowNum) = wl_DelBkMrk
    mi_ReCalcRow = pv_RowNum
    For i = 1 To MSFg.Cols - 1
        Call ReCalcProc(i, 0)
    Next i
    me_Mode = fgmnorm
End Function
Public Sub RefreshData()
  LoadData (False)
End Sub
Private Sub LoadData(ByVal Reload As Boolean)
    On Error GoTo errh
    Dim i As Integer, ws_OrgRecSource As String
    ws_OrgRecSource = ms_RecSource
    
    RaiseEvent SetRecSource
    If ms_RecSource = "" Then Exit Sub
    If Reload = False And ms_RecSource = ws_OrgRecSource Then Exit Sub
    
    MSFg.Redraw = False
    me_Mode = fgminit
    Set mo_Recset = Nothing
    
    Set mo_Recset = Parent.adc.Connection.OpenResultset(ms_RecSource)
    If Reload = False Then
        MSFg.Rows = 1
    End If
    Dim j As Long
    j = 1
    Do While Not (mo_Recset.EOF Or mo_Recset.BOF)
            If (Not Reload) Or (j >= ml_MaxRowBeforeCopy) Then
              MSFg.Additem ""
              Call MoveData(MSFg.Rows - 1, ctm_Read)
            '  j = MSFg.Rows - 1
  '          Else
  '            j = j + 1
            End If
            MSFg.RowData(j) = mo_Recset.Bookmark
            j = j + 1
            mo_Recset.MoveNext
    Loop

    me_Mode = fgmInitOver
    For i = 1 To mc_Columns.Count
'sri        Call ReCalcProc(i, 0)
    Next i
    me_Mode = fgminit

    If Not Reload Then
      If mo_Recset.RecCount <> 0 Then mo_Recset.MoveFirst
      mf_stRowCol = True
    
      mi_PrevRow = 0
      mi_PrevCol = 0

    '********* added later as child's recsource refer to current row
      If MSFg.Rows > 1 Then
        mi_PrevRow = 1
        MSFg.Row = IIf(Reload, Rows - 1, 1)
      Else
        MSFg.Row = 0
      End If
    
    
    ''***********************************************
    If MSFg.Rows > 1 Then Call RefChild
    
    mi_PrevRow = 0
    MSFg.Row = 0
    MSFg.COL = 0
    
    mf_stRowCol = False
    
    If mo_Recset.RecCount <> 0 Then mo_Recset.MoveFirst
    End If
    me_Mode = fgmnorm
'sri
    If Not me_Mode = fgmServerCopy Then
      MSFg.Redraw = True
    End If
    Exit Sub
errh:
    DispMsg "Error", etError
End Sub
Private Sub ReloadData(ByVal Reload As Boolean)
    Dim i As Integer, ws_OrgRecSource As String
    ws_OrgRecSource = ms_RecSource
    RaiseEvent AfterServerCopy
   ' If ms_RecSource = "" Then Exit Sub
   ' If Reload = False And ms_RecSource = ws_OrgRecSource Then Exit Sub
    
    MSFg.Redraw = False
    me_Mode = fgminit
   ' Set mo_Recset = Nothing
    Dim wo_recset As MDORowSet
    Set wo_recset = Parent.adc.Connection.OpenResultset(ms_RecSource)
'    Do While Not (wo_recset.EOF)
        ' You can place if condition here
 '       wo_recset.MoveNext
 '   Loop
    
'    mo_Recset.Addnew
  '  MSFg.Rows = 1
    Do While Not (wo_recset.EOF Or wo_recset.BOF)
        mo_Recset.Addnew
        For i = 0 To wo_recset.FldCount - 1
            mo_Recset.FldValue(i) = (wo_recset.FldValue(i))
           ' (i).Name) = wo_recset.Fields(i).Value
        Next i
        'mo_Recset.UpdateNoDB

            MSFg.Additem ""
            
    '        mo_Recset.AddNew
            MSFg.RowData(MSFg.Rows - 1) = mo_Recset.Bookmark
         '    wo_recset.
         '   If MSFg.Rows - 1 > mf_ReloadDataOnEnd - 1 Then
              Call MoveData(MSFg.Rows - 1, ctm_Read)
              '   End If
            'mo_Recset.UpdateNoDB
            wo_recset.MoveNext
    Loop
    wo_recset.CloseRes
    Set wo_recset = Nothing
    
    me_Mode = fgmInitOver
    For i = 1 To mc_Columns.Count
        Call ReCalcProc(i, 0)
    Next i
    me_Mode = fgminit

    If mo_Recset.RecCount <> 0 Then mo_Recset.MoveFirst
    mf_stRowCol = True
    
    mi_PrevRow = 0
    mi_PrevCol = 0

    '********* added later as child's recsource refer to current row
      If Not Reload Then
      If MSFg.Rows > 1 Then
        mi_PrevRow = 1
        MSFg.Row = IIf(Reload, Rows - 1, 1)
      Else
        MSFg.Row = 0
      End If
End If
    ''***********************************************
    If MSFg.Rows > 1 Then Call RefChild
    
    mi_PrevRow = 0
    MSFg.Row = 0
    MSFg.COL = 0
    
    mf_stRowCol = False
    
    If mo_Recset.RecCount <> 0 Then mo_Recset.MoveFirst

    me_Mode = fgmnorm
    'sri
    If Not me_Mode = fgmServerCopy Then
      MSFg.Redraw = True
    End If
End Sub

Public Sub Store(ByVal pv_Fldname As String, ByVal pv_vNewValue As Variant)
    On Error GoTo errh
    If me_Mode = fgmServerCopy Then
          mo_QueryObject.QryParametersByName(pv_Fldname) = pv_vNewValue
          Exit Sub
    End If
    'Debug.Print pv_Fldname, pv_vNewValue
    If mo_Recset Is Nothing Then Exit Sub
    Dim wv_cmpval As Variant
    wv_cmpval = IIf(IsNull(mo_Recset.FldValue(pv_Fldname)), "", mo_Recset.FldValue(pv_Fldname))
    
    Select Case mo_Recset.FldType(pv_Fldname)
    Case Is = mdoftDateTime
        If IsDate(pv_vNewValue) Then
            pv_vNewValue = CDate(pv_vNewValue)
        Else
            pv_vNewValue = Null
        End If
    Case Is = mdoftString
        wv_cmpval = Trim(wv_cmpval)
        pv_vNewValue = Trim(pv_vNewValue)
    End Select
    If wv_cmpval <> pv_vNewValue Then
        mf_DataChanged = True
    End If
'    If mo_Recset.FldType(pv_Fldname) = mdoftString Then
 '       If pv_vNewValue = "" Then pv_vNewValue = " "
  '  End If
    
    '*** (Bef 14/07/05) If Not mf_ChkData Then mo_Recset.FldValue(pv_Fldname) = pv_vNewValue
    
    ''*** Manoj (14/07/05)
    If Not mf_ChkData And (wv_cmpval <> pv_vNewValue Or IsEmpty(mo_Recset.FldValue(pv_Fldname))) Then
           mo_Recset.FldValue(pv_Fldname) = pv_vNewValue
    End If
    
    If ms_ValidateStoredProcedure <> "" Then
        mo_QueryObjectValidate.QryParametersByName("@" + pv_Fldname) = pv_vNewValue
    End If
    
    Exit Sub
errh:
    Err.Raise vbObjectError + 512, "", Err.Description + " (" + pv_Fldname + ") "
End Sub
Public Function SaveRec(Optional ByVal pv_RowNum As Variant) As Boolean
    mf_SaveByCode = True
    SaveRec = RowSave(pv_RowNum)
    mf_SaveByCode = False
End Function
Public Function AddSaveRec(Optional ByVal pv_RowNum As Variant) As Boolean
    mf_AddSave = True
    AddSaveRec = RowSave(pv_RowNum)
    mf_AddSave = False
End Function
Public Sub ReCalc(ByVal pv_ColName As String, ByVal pv_OldVal As Variant, ByVal pv_SourceName As String, Optional ByVal pv_SourceColName As String, Optional ByVal pv_RowNum As Integer)
    If IsMissing(pv_RowNum) Then pv_RowNum = 0
    If Not mf_SpecialCopy Then
        RaiseEvent ReCalc(pv_RowNum, pv_ColName, pv_OldVal, pv_SourceName, pv_SourceColName)
        End If
End Sub
Public Function MaxVal(ByVal pv_ColName As String) As Integer
    Dim mi_MaxVal As Integer, i As Integer
    mi_MaxVal = 0
    For i = 1 To MSFg.Rows - 1
        If Value(i, pv_ColName) > mi_MaxVal Then
            mi_MaxVal = Value(i, pv_ColName)
        End If
    Next i
    MaxVal = mi_MaxVal
End Function
Public Sub Additem()
    Dim wi_CurRows As Integer
' ########################  Manoj #### Ver: 2.0.6 #### Date: 16/06/2004  ########################
  'Check Login session is Working or Not
    If Not ChkLogin(Parent.adc) Then Exit Sub
' ###########################################  Manoj  ###########################################
    wi_CurRows = MSFg.Rows
    MSFg.Additem ""
    If wi_CurRows = 1 Then
        mf_stRowCol = True
        MSFg.Row = 0
        mf_stRowCol = False
    End If
    Call SetColsInitMode(True)
    '?? when msfg.rows=1(msfg.row=0), calling additem sets .row to 1
    ' not desirable . hence put back to 0
End Sub
Public Sub PropertyModified(ByVal pv_PropertyName As String, ByVal pv_ColName As String)
    If mf_IniProp Then Exit Sub
    Dim wi_ColNum As Integer, wo_reqClsFlx As MWCLS_FLX
    wi_ColNum = GetCol(pv_ColName)
    Set wo_reqClsFlx = mc_Columns(wi_ColNum)
    With wo_reqClsFlx
    Select Case pv_PropertyName
    Case Is = "heading"
        MSFg.TextMatrix(0, wi_ColNum) = .Heading
    Case Is = "style"
        If .Style = fgcNormal Then
            MSFg.ColWidth(wi_ColNum) = .Width
        ElseIf .Style = fgcinvisible Then
            MSFg.ColWidth(wi_ColNum) = 0
        ElseIf .Style = fgcReadOnly Then
            MSFg.ColWidth(wi_ColNum) = .Width
        End If
    End Select
    End With
End Sub
Public Sub EnterFocus()
    Call MSFg_GotFocus
End Sub
Public Sub LeaveFocus()
    mf_ForceLost = True
    Call TXT_FLX_LostFocus
'    TXT_FLX.Visible = True
    mf_ForceLost = False
End Sub
Public Function DelAllRecs()
    Dim wi_TotRec As Integer, i As Integer
    DelAllRecs = True
    wi_TotRec = MSFg.Rows - 1
    For i = wi_TotRec To 1 Step -1
        If Not Del_Rec(False, i) Then
            DelAllRecs = False
            Exit For
        End If
    Next i
End Function
Public Sub Init()
    Dim wv_rArr() As String, i As Integer
    MSFg.Rows = 1
    ms_RecSource = ""
    mf_stRowCol = True
    MSFg.Row = 0
    MSFg.COL = 0
    mi_PrevRow = 0
    mi_PrevCol = 0
    mf_stRowCol = False
    wv_rArr = Split(ms_child, ",")
    If IsArray(wv_rArr) Then
        For i = 0 To UBound(wv_rArr)
            Call Parent.Controls(wv_rArr(i)).Init
        Next i
    End If
End Sub
Public Sub SetCellColour(ByVal pv_RowNum As Integer, ByVal pv_ColRef As Variant, ByVal pv_ReqColour As Variant)
    On Error GoTo errh
    If pv_RowNum <= 0 Or pv_RowNum >= MSFg.Rows Then Err.Raise vbObjectError + 512, "", "Invalid RowNum"
    If pv_ColRef <= 0 Or pv_ColRef >= MSFg.Cols Then Err.Raise vbObjectError + 512, "", "Invalid ColNum"
    Dim mi_Row As Integer, mi_Col As Integer
    mi_Row = MSFg.Row
    mi_Col = MSFg.COL
    mf_stRowCol = True
    MSFg.Row = pv_RowNum
    MSFg.COL = pv_ColRef
    MSFg.CellBackColor = pv_ReqColour
    MSFg.Row = mi_Row
    MSFg.COL = mi_Col
    mf_stRowCol = False
    Exit Sub
errh:
    Call Parent.adc.HandleError
End Sub
Public Sub SetTopRow()
    If mi_PrevRow > 1 And mi_PrevRow < MSFg.Rows Then
        MSFg.TopRow = mi_PrevRow
    End If
End Sub
Public Function HalfValidate() As Boolean
    If TypeOf m_ActScrn.ActiveControl Is MWCTL_HLP Then
        Parent.MWCTL_HLP1.Mode = hmHlpBrowsing
    End If
    Call Val_Col
    If mf_Val_failed Then
        mf_Val_failed = False
        TXT_FLX.Visible = True
        mf_stRowColIfOnGrd = False
        HalfValidate = True
        Exit Function
    End If
End Function
Public Function Validate() As Boolean
    If TypeOf m_ActScrn.ActiveControl Is MWCTL_HLP Then
        Parent.MWCTL_HLP1.Mode = hmHlpBrowsing
    End If
    Call Val_Col
    If mf_Val_failed Then
        mf_Val_failed = False
        Validate = True
        Exit Function
    End If
    Call val_row
    If mf_Val_failed Then
        mf_Val_failed = False
        Validate = True
        Exit Function
    End If
'sri    RaiseEvent Valid
'    TXT_FLX.Visible = False
'    mf_stRowColIfOnGrd = True
End Function
Public Sub AddSave(ByVal pv_ObjNm As String, ByVal pv_AddOpt As en_AddSaveOpt)
    Dim wa_arr As Variant
    On Error GoTo errh
    If mc_AddSave Is Nothing Then Set mc_AddSave = New Collection
    wa_arr = Array(pv_ObjNm, pv_AddOpt)
    mc_AddSave.Add wa_arr, pv_ObjNm
    Exit Sub
errh:
    If Err.Number = 457 Then
        Exit Sub
    Else
        Err.Raise Err.Number, "", Err.Description
    End If
End Sub
Public Sub StartCopy()
    me_Mode = fgmCopy
    MSFg.Redraw = False
End Sub

Public Sub StartServerCopy(ByVal ReloadDataOnEnd As Boolean)
    MSFg.Redraw = False
    RaiseEvent BeforeServerCopy
    ml_MaxRowBeforeCopy = MSFg.Rows
    mf_ReloadDataOnEnd = ReloadDataOnEnd
    me_Mode = fgmServerCopy
End Sub
Public Sub EndServerCopy()
    RaiseEvent AfterServerCopy

   If mf_ReloadDataOnEnd Then
      LoadData (True)
    End If
    Row = Rows - 1
    MSFg.Redraw = True
    me_Mode = fgmnorm
End Sub
Public Sub EndCopy()
    Dim i As Integer
    me_Mode = fgmCopyOver
    If Not mf_StopRecursiveRecalc Then
    For i = 1 To mc_Columns.Count
        Call ReCalcProc(i, 0)
    Next i
    End If
    me_Mode = fgmnorm
    'sri
    If Not me_Mode = fgmServerCopy Then

      MSFg.Redraw = True
    End If
End Sub
Public Sub ChildsEnaDisa(ByVal pv_Enable As Boolean)
    Dim wv_rArr() As String, i As Integer
    wv_rArr = Split(ms_child, ",")
    If IsArray(wv_rArr) Then
        For i = 0 To UBound(wv_rArr)
            Parent.Controls(wv_rArr(i)).Enabled = pv_Enable
        Next i
    End If
End Sub
Public Sub Resync(ByVal pv_AddSaveOpt As en_AddSaveOpt)
    If pv_AddSaveOpt = asAll Then mo_Recset.Resync mdoryAll
    If pv_AddSaveOpt = asCurrent Then mo_Recset.Resync mdoryCurrent
End Sub
Public Sub CloseRes()
    Dim i As Integer
    Dim ws_rArr() As String
    Set mo_Recset = Nothing
    ws_rArr = Split(ms_child, ",")
    If IsArray(ws_rArr) Then
        For i = 0 To UBound(ws_rArr)
            Call Parent.Controls(ws_rArr(i)).CloseRes
        Next i
    End If
End Sub
Public Sub SetColWidthToHd()
    Dim ws_formatstring As String, i As Integer, wo_reqClsFlx As MWCLS_FLX
    ws_formatstring = Space(2) + "|"
    For i = 1 To mc_Columns.Count
        Set wo_reqClsFlx = mc_Columns(i)
        ws_formatstring = ws_formatstring + wo_reqClsFlx.Heading + "|"
    Next i
    FormatString = Mid(ws_formatstring, 1, Len(ws_formatstring) - 1)
    For i = 1 To MSFg.Cols - 1
        Set wo_reqClsFlx = mc_Columns(i)
        wo_reqClsFlx.Width = MSFg.ColWidth(i)
        If wo_reqClsFlx.Style = fgcinvisible Then
            MSFg.ColWidth(i) = 0
        End If
    Next
    Set wo_reqClsFlx = Nothing
End Sub
Public Sub TermCtl()
    Set mo_NextCtl = Nothing
    Set mo_PrevCtl = Nothing
    Set mc_Columns = Nothing
End Sub
'***************** Methods End ************************************************************'
'***************** Subs Begin ************************************************************'
Private Sub SendToNextCtl()
    If mo_NextCtl Is Nothing Then Set mo_NextCtl = Parent.adc
    If mo_NextCtl.Visible And mo_NextCtl.Enabled Then
        MWLib.SendFocus m_ActScrn, mo_NextCtl
    End If
'??? check en and vis before setfocus
End Sub
Private Sub SendToPrevCtl()
    If mo_PrevCtl Is Nothing Then Set mo_PrevCtl = Parent.adc
    If mo_PrevCtl.Visible And mo_PrevCtl.Enabled Then
        MWLib.SendFocus m_ActScrn, mo_PrevCtl
    End If
'???check en and vis before setfocus
End Sub
Private Function GetCol(ByVal pv_ColRef As String) As Integer
    Dim i As Integer, wo_reqClsFlx As MWCLS_FLX
    If IsNumeric(pv_ColRef) Then
        GetCol = pv_ColRef
    Else
        For i = 1 To mc_Columns.Count
            Set wo_reqClsFlx = mc_Columns(i)
            If wo_reqClsFlx.ColName = UCase(pv_ColRef) Then
                GetCol = i
                Exit Function
            End If
        Next i
    End If
    Set wo_reqClsFlx = Nothing
End Function
Private Sub RefChild()
    Dim wv_rArr() As String, i As Integer
    wv_rArr = Split(ms_child, ",")
    If IsArray(wv_rArr) Then
        For i = 0 To UBound(wv_rArr)
            If IsNew(mi_PrevRow) Then Call Parent.Controls(wv_rArr(i)).Init
            Call Parent.Controls(wv_rArr(i)).RefreshData
        Next i
    End If
End Sub
Private Sub MoveData(ByVal pv_RowNum As Integer, ByVal pv_Opt As String)
    Dim ws_DataFld As String, i As Integer, wo_reqClsFlx As MWCLS_FLX, we_reqDataType As en_FldDataType, ws_reqFormatMask As String
    For i = 1 To mc_Columns.Count
        Set wo_reqClsFlx = mc_Columns(i)
        ws_DataFld = wo_reqClsFlx.DataFld
        we_reqDataType = wo_reqClsFlx.DataType
        ws_reqFormatMask = wo_reqClsFlx.FormatMask
        If ws_DataFld <> "" Then
            If pv_Opt = ctm_Read Then
                Select Case we_reqDataType
                Case Is = fdtFloat, fdtNumber, fdtTimeType
                  If ws_reqFormatMask <> "" Then
                    MSFg.TextMatrix(pv_RowNum, i) = Format(IIf(IsNull(mo_Recset.FldValue(ws_DataFld)), 0, mo_Recset.FldValue(ws_DataFld)), ws_reqFormatMask)
                  End If
                Case Is = fdtDateType
                    MSFg.TextMatrix(pv_RowNum, i) = Format(IIf(IsNull(mo_Recset.FldValue(ws_DataFld)), "", mo_Recset.FldValue(ws_DataFld)), "dd/mm/yy")
                Case Else
                    MSFg.TextMatrix(pv_RowNum, i) = IIf(IsNull(mo_Recset.FldValue(ws_DataFld)), "", Trim(mo_Recset.FldValue(ws_DataFld)))
                End Select
                mi_ReCalcRow = pv_RowNum
                Call ReCalcProc(i, 0)
            Else
                Store ws_DataFld, Value(pv_RowNum, i)
            End If
        End If
    Next i
    Set wo_reqClsFlx = Nothing
End Sub
Private Sub Reload_MoveData(ByVal pv_RowNum As Integer, ByVal pv_Opt As String, po_rowset As MDORowSet)
    Dim ws_DataFld As String, i As Integer, wo_reqClsFlx As MWCLS_FLX, we_reqDataType As en_FldDataType, ws_reqFormatMask As String
    For i = 1 To mc_Columns.Count
        Set wo_reqClsFlx = mc_Columns(i)
        ws_DataFld = wo_reqClsFlx.DataFld
        we_reqDataType = wo_reqClsFlx.DataType
        ws_reqFormatMask = wo_reqClsFlx.FormatMask
        If ws_DataFld <> "" Then
            Store ws_DataFld, mo_Recset.FldValue(ws_DataFld)

            If pv_Opt = ctm_Read Then
                Select Case we_reqDataType
                Case Is = fdtFloat, fdtNumber, fdtTimeType
                  If ws_reqFormatMask <> "" Then
                    MSFg.TextMatrix(pv_RowNum, i) = Format(IIf(IsNull(mo_Recset.FldValue(ws_DataFld)), 0, mo_Recset.FldValue(ws_DataFld)), ws_reqFormatMask)
                  End If
                Case Is = fdtDateType
                    MSFg.TextMatrix(pv_RowNum, i) = Format(IIf(IsNull(mo_Recset.FldValue(ws_DataFld)), "", mo_Recset.FldValue(ws_DataFld)), "dd/mm/yy")
                Case Else
                    MSFg.TextMatrix(pv_RowNum, i) = IIf(IsNull(mo_Recset.FldValue(ws_DataFld)), "", Trim(mo_Recset.FldValue(ws_DataFld)))
                End Select
                mi_ReCalcRow = pv_RowNum
                Call ReCalcProc(i, 0)
            'Else
            End If
        End If
    Next i
    Set wo_reqClsFlx = Nothing
End Sub

Private Sub Whn_Col()
    On Error GoTo errh
    Dim wo_reqClsFlx As MWCLS_FLX, we_reqDataType As en_FldDataType, pr_Cancel As Boolean, pr_ErrMsg As String
    If MSFg.COL <= 0 Or MSFg.COL >= MSFg.Cols Then Exit Sub
    If MSFg.Row <= 0 Or MSFg.Row >= MSFg.Rows Then Exit Sub
    mi_PrevCol = MSFg.COL
    Set wo_reqClsFlx = mc_Columns(mi_PrevCol)
    With wo_reqClsFlx
        DispMsg .ToolTipText, etInfo
        we_reqDataType = .DataType
        If we_reqDataType = fdtFloat Or we_reqDataType = fdtNumber Or we_reqDataType = fdtTimeType Then
            TXT_FLX.Alignment = en_TxtAlign.tRightAlign
        Else
            TXT_FLX.Alignment = en_TxtAlign.tLeftAlign
        End If
        TXT_FLX.DataType = we_reqDataType
        TXT_FLX.Mask = .Mask
        TXT_FLX.MaxLength = .MaxLength
        TXT_FLX.Locked = False
        '***(Bef 14/06/05) TXT_FLX.Move MSFg.Left + MSFg.CellLeft, (MSFg.Top + MSFg.CellTop) - 15, MSFg.CellWidth, MSFg.RowHeight(0)
        
        '*** Jenny (14/06/05)
        If (me_Mode <> fgmCopy) Then
          If GrdAppearance = flxGrdAppFlat Then
            TXT_FLX.Move MSFg.Left + MSFg.CellLeft, (MSFg.Top + MSFg.CellTop) - 15, MSFg.CellWidth, MSFg.RowHeight(0)
          Else
            TXT_FLX.Move MSFg.Left + MSFg.CellLeft, (MSFg.Top + MSFg.CellTop) - 30, MSFg.CellWidth, MSFg.RowHeight(0)
          End If
        End If
        '*** Jenny (14/06/05)
        
        If Not mf_SaveByCode Then
            If Not TXT_FLX.Visible Then TXT_FLX.Visible = True
            If UserControl.Enabled Then
              TXT_FLX.SetFocus
              '*** (Bef 13/09/05)  Jenny  (This Code Is Required To Give Background Color To The Selected Row. But It Makes Scrolling Slow)
              'SetCellColour MSFg.Row, MSFg.COL, mo_BckColorSelRow     '*** Jenny
              '*** (Bef 13/09/05)  Jenny  (This Code Is Required To Give Background Color To The Selected Row. But It Makes Scrolling Slow)
            End If
        End If
'??? check only my enabled - nahi chale
        pr_Cancel = False
        If Not mf_SpecialCopy Then
          If .Style <> fgcReadOnly Then RaiseEvent ColWhen(MSFg.Row, .ColName, pr_Cancel, pr_ErrMsg, me_Action)
        End If
        If .Style = fgcReadOnly Then
            pr_Cancel = True
            pr_ErrMsg = "Cannot modify Readonly column"
        End If
        me_Action = fgaOthers
        TXT_FLX.TxtText = MSFg
        Call TXT_FLX.InitVal
        TXT_FLX.SelStart = 0
        If pr_Cancel Then
            TXT_FLX.Locked = True
            Err.Raise vbObjectError + 514, "", pr_ErrMsg
        End If
    End With
    Exit Sub
errh:
    Call Parent.adc.HandleError
End Sub
Private Sub Val_Col()
    Dim wf_Cancel As Boolean
    If mi_PrevRow <= 0 Or mi_PrevRow >= MSFg.Rows Then Exit Sub
    If mi_PrevCol <= 0 Or mi_PrevCol >= MSFg.Cols Then Exit Sub
    If Not TXT_FLX.Locked Then
        Parent.adc.ValFailed = ""
        mf_ColValByCode = False
        wf_Cancel = False
        Value(mi_PrevRow, mi_PrevCol, wf_Cancel) = TXT_FLX.TxtText
        mf_ColValByCode = True
        If Not wf_Cancel Then
            Parent.MWCTL_HLP1.Visible = False
            Parent.MWCTL_HLP1.Mode = hmHlpNotavlbl
        End If
     Else
        Parent.MWCTL_HLP1.Visible = False
        Parent.MWCTL_HLP1.Mode = hmHlpNotavlbl
    End If
End Sub
Private Sub ReCalcProc(ByVal pv_ReqCol As Integer, ByVal pv_OldVal As Variant)
'    On Error GoTo ERRH
    Dim wo_reqClsFlx As MWCLS_FLX, wv_rArr As Variant, ws_ReqStr As String, wi_FirBr As Integer, wi_SecBr As Integer, ws_reqPrm As String, wo_ReqObj As Object, ws_RecalcParent As String, i As Integer
    Set wo_reqClsFlx = mc_Columns(pv_ReqCol)
    ws_RecalcParent = wo_reqClsFlx.ReCalcParent
    wv_rArr = Split(ws_RecalcParent, ",")
    If Not IsArray(wv_rArr) Then Exit Sub
    For i = 0 To UBound(wv_rArr, 1)
        ws_ReqStr = wv_rArr(i)
        wi_FirBr = InStr(1, ws_ReqStr, "(")
        wi_SecBr = InStr(1, ws_ReqStr, ")")
        If wi_FirBr <> 0 Then
            ws_reqPrm = Mid(ws_ReqStr, wi_FirBr + 1, wi_SecBr - wi_FirBr - 1)
            Set wo_ReqObj = Parent.Controls(Mid(ws_ReqStr, 1, wi_FirBr - 1))
        Else
            Set wo_ReqObj = Parent.adc(ws_ReqStr)
        End If
        If TypeOf wo_ReqObj Is MWCTL_FLX Or UCase(TypeName(wo_ReqObj)) = "MWCTL_FLX" Then
            wo_ReqObj.ReCalc ws_reqPrm, pv_OldVal, UCase(UserControl.Extender.Name), UCase(wo_reqClsFlx.ColName), mi_ReCalcRow
        Else
            wo_ReqObj.ReCalc pv_OldVal, UCase(UserControl.Extender.Name), UCase(wo_reqClsFlx.ColName), mi_ReCalcRow
        End If
    Next i
'??? consider for speed
End Sub
Private Sub ReCalcProcV(ByVal pv_RecalcRow As Integer, ByVal pv_ReqCol As Integer, ByVal pv_OldVal As Variant)
'    On Error GoTo ERRH
    Dim wo_reqClsFlx As MWCLS_FLX, wv_rArr As Variant, ws_ReqStr As String, wi_FirBr As Integer, wi_SecBr As Integer, ws_reqPrm As String, wo_ReqObj As Object, ws_RecalcParent As String, i As Integer
    Set wo_reqClsFlx = mc_Columns(pv_ReqCol)
    ws_RecalcParent = wo_reqClsFlx.ReCalcParent
    wv_rArr = Split(ws_RecalcParent, ",")
    If Not IsArray(wv_rArr) Then Exit Sub
    For i = 0 To UBound(wv_rArr, 1)
        ws_ReqStr = wv_rArr(i)
        wi_FirBr = InStr(1, ws_ReqStr, "(")
        wi_SecBr = InStr(1, ws_ReqStr, ")")
        If wi_FirBr <> 0 Then
            ws_reqPrm = Mid(ws_ReqStr, wi_FirBr + 1, wi_SecBr - wi_FirBr - 1)
            Set wo_ReqObj = Parent.Controls(Mid(ws_ReqStr, 1, wi_FirBr - 1))
        Else
            Set wo_ReqObj = Parent.adc(ws_ReqStr)
        End If
        If (TypeOf wo_ReqObj Is MWCTL_FLX) Or UCase(TypeName(wo_ReqObj)) = "MWCTL_FLX" Then
            wo_ReqObj.ReCalc ws_reqPrm, pv_OldVal, UCase(UserControl.Extender.Name), UCase(wo_reqClsFlx.ColName), pv_RecalcRow
        Else
            wo_ReqObj.ReCalc pv_OldVal, UCase(UserControl.Extender.Name), UCase(wo_reqClsFlx.ColName), pv_RecalcRow
        End If
    Next i
'??? consider for speed
End Sub
Private Sub whn_row()
    '*** (13/09/05)  Jenny
    If MSFg.Row <= 0 Or MSFg.Row >= MSFg.Rows Then Exit Sub
    
    If mi_PrevRow > 0 And mi_PrevRow < MSFg.Rows Then
      MSFg.TextMatrix(mi_PrevRow, 0) = ""
    End If

    mi_PrevRow = MSFg.Row
    If ms_child <> "" Then Call RefChild
    RaiseEvent RowWhen(MSFg.Row)
    
    MSFg.TextMatrix(mi_PrevRow, 0) = "¶"
    '*** (13/09/05)  Jenny
    
    
    
    '*** (Bef 13/09/05)  Jenny  (This Code Is Required To Give Background Color To The Selected Row. But It Makes Scrolling Slow)
    'If MSFg.Row <= 0 Or MSFg.Row >= MSFg.Rows Then Exit Sub
    '
    ''*** Jenny
    'Dim wi_PrevRow As Integer
    'wi_PrevRow = -1
    ''*** Jenny
    '
    'If mi_PrevRow > 0 And mi_PrevRow < MSFg.Rows Then
    '  MSFg.TextMatrix(mi_PrevRow, 0) = ""
    '  wi_PrevRow = mi_PrevRow   '*** Jenny
    'End If
    '
    'mi_PrevRow = MSFg.Row
    'If ms_child <> "" Then Call RefChild
    'RaiseEvent RowWhen(MSFg.Row)
    '
    ''MSFg.TextMatrix(mi_PrevRow, 0) = "¶"    '*** Original Code   (Bef 14/01/2005)
    '
    ''****** Jenny
    'If UserControl.Picture = LoadPicture() Then
    '  MSFg.TextMatrix(mi_PrevRow, 0) = "¶"
    'Else
    '  SetCellPic wi_PrevRow, 0, LoadPicture()
    '  SetCellPic mi_PrevRow, 0, UserControl.Picture
    'End If
    '
    'If wi_PrevRow <> mi_PrevRow Then
    '  If wi_PrevRow > 0 Then SetRowColour wi_PrevRow, BckColor
    '  SetRowColour mi_PrevRow, mo_BckColorSelRow
    'End If
    ''****** Jenny
    '*** (Bef 13/09/05)  Jenny  (This Code Is Required To Give Background Color To The Selected Row. But It Makes Scrolling Slow)
End Sub
Public Sub SetRowColour(ByVal pv_RowNum As Integer, ByVal pv_ReqColour As Variant)
    '*** Jenny
    On Error GoTo errh
    If pv_RowNum <= 0 Or pv_RowNum >= MSFg.Rows Then Err.Raise vbObjectError + 512, "", "Invalid RowNum For Setting Row Color"
    Dim mi_Row As Integer, mi_Col As Integer
    mb_SelFlag = True
    mi_Row = MSFg.Row
    mi_Col = MSFg.COL
    mf_stRowCol = True
    MSFg.FillStyle = flexFillRepeat
    'MSFg.SelectionMode = flexSelectionByRow
    MSFg.Row = pv_RowNum
    MSFg.COL = 1
    MSFg.RowSel = pv_RowNum
    MSFg.ColSel = MSFg.Cols - 1
    MSFg.CellBackColor = pv_ReqColour
    
    mb_SelFlag = False
    MSFg.FillStyle = flexFillSingle
    'MSFg.SelectionMode = flexSelectionFree
    MSFg.Row = mi_Row
    MSFg.COL = mi_Col
    MSFg.RowSel = mi_Row
    MSFg.ColSel = mi_Col
    mf_stRowCol = False
    Exit Sub
errh:
    Call Parent.adc.HandleError
    '*** Jenny
End Sub
Public Sub SetCellPic(ByVal pv_RowNum As Integer, ByVal pv_ColRef As Variant, ByVal pv_ReqPic As Variant)
    '*** Jenny
    On Error GoTo errh
    'If pv_RowNum <= 0 Or pv_RowNum >= MSFg.Rows Then Err.Raise vbObjectError + 512, "", "Invalid RowNum"
    If pv_RowNum <= 0 Or pv_RowNum >= MSFg.Rows Then Exit Sub
    If pv_ColRef >= MSFg.Cols Then Err.Raise vbObjectError + 512, "", "Invalid ColNum"
    'If pv_ColRef <= 0 Or pv_ColRef >= MSFg.Cols Then Err.Raise vbObjectError + 512, "", "Invalid ColNum"
    Dim mi_Row As Integer, mi_Col As Integer
    mi_Row = MSFg.Row
    mi_Col = MSFg.COL
    mf_stRowCol = True
    MSFg.Row = pv_RowNum
    MSFg.COL = pv_ColRef
    If CStr(pv_ReqPic) <> "" And Dir(CStr(pv_ReqPic)) <> "" Then
      Set MSFg.CellPicture = LoadPicture(CStr(pv_ReqPic))
    Else
      Set MSFg.CellPicture = pv_ReqPic
    End If
    MSFg.CellPictureAlignment = flexAlignCenterCenter
    MSFg.Row = mi_Row
    MSFg.COL = mi_Col
    mf_stRowCol = False
    Exit Sub
errh:
    Call Parent.adc.HandleError
    '*** Jenny
End Sub
Private Sub val_row()
    On Error GoTo errh
    Dim pr_Cancel As Boolean, pr_ErrMsg As String
    If mi_PrevRow <= 0 Or mi_PrevRow >= MSFg.Rows Then Exit Sub
    pr_Cancel = False
    Parent.adc.ValFailed = ""
    If Not mf_SpecialCopy Then
      RaiseEvent RowValid(mi_PrevRow, pr_Cancel, pr_ErrMsg)
    End If
    If pr_Cancel = True Then Err.Raise vbObjectError + 513, 0, pr_ErrMsg
    Call RowSave
    If mf_AllowAdd And mf_Additem Then
        Call Additem
        MSFg.Row = MSFg.Rows - 1
        MSFg.COL = FirstCol
        Call whn_row
        Call Whn_Col
    End If
    mf_Additem = False
    Exit Sub
errh:
   Call Parent.adc.HandleError
End Sub
Private Sub errhndlr(ByVal RowNum As Integer)
    If EnAndVis() And (Not mf_AddSave) And (Not mf_SaveByCode) And m_ActScrn.ActiveForm Is UserControl.Parent Then
        mf_Val_failed = True
        mf_Additem = False
        Parent.adc.ValFailed = "row"
        Call Parent.adc.HandleError
        mo_Recset.CancelUpdate
' parent.adc.valfailed has to occur before handleerror - do not change order of execution
        If TXT_FLX.Enabled And TXT_FLX.Visible Then TXT_FLX.SetFocus
        MSFg.Row = mi_PrevRow
        MSFg.COL = mi_PrevCol
        Call Whn_Col
    Else
        Call Parent.adc.HandleError
        mo_Recset.CancelUpdate
        If Not mf_AddSave Then Call RefRow(RowNum)
    End If
End Sub
Private Sub RefRow(ByVal pv_RowNum As Variant)
    TXT_FLX.Visible = False
    If IsNew(pv_RowNum) Then
        Call Del_Rec(False, pv_RowNum)
    Else
        mo_Recset.Bookmark = MSFg.RowData(pv_RowNum)
        Call MoveData(pv_RowNum, ctm_Read)
    End If
End Sub

Private Function FirstCol()
    Dim wi_FirCol As Integer, i As Integer, wo_reqClsFlx As MWCLS_FLX
    FirstCol = 0
    For i = 1 To mc_Columns.Count
        Set wo_reqClsFlx = mc_Columns(i)
        If wo_reqClsFlx.Style = fgcNormal Then
            Exit For
        End If
    Next i
    FirstCol = i
End Function

Private Function InSelMode()
    If (MSFg.RowSel <> MSFg.Row Or MSFg.ColSel <> MSFg.COL) Then
        InSelMode = True
    Else
        InSelMode = False
    End If
End Function

Private Function RowSave(Optional ByVal pv_RowNum As Variant) As Boolean

    On Error GoTo errh
    RowSave = False
    If me_Mode = fgmServerCopy Then
      mo_QueryObject.QryExecute
      RowSave = True
      Exit Function
    End If
       
    If mo_Recset Is Nothing Then
        RowSave = True
        ParentRef.adc.BtnValFailed = ""     '*** (14/09/05)  Jenny
        Exit Function
    End If
    
    Dim pr_Cancel As Boolean, pr_ErrMsg As String
    If IsMissing(pv_RowNum) Then pv_RowNum = mi_PrevRow
    If mf_AddSave = True And (IsNew(pv_RowNum) Or IsDel(pv_RowNum)) Then
        RowSave = True
        Exit Function
    End If

    pr_Cancel = False
 'sri
    If Not mf_SpecialCopy Then
      RaiseEvent RowSave(pv_RowNum, pr_Cancel, pr_ErrMsg)
    End If
    If pr_Cancel = True Then Err.Raise vbObjectError + 513, 0, pr_ErrMsg
    mf_DataChanged = False
    
    If MSFg.RowData(pv_RowNum) <> 0 Then
        mo_Recset.Bookmark = MSFg.RowData(pv_RowNum)
        mf_ChkData = True
        
        Call MoveData(pv_RowNum, ctm_Write)
        RaiseEvent InitKey(pv_RowNum)
        mf_ChkData = False
    End If
    
    Parent.adc.Connection.BeginTrans (mdoCommitNone)
    If MSFg.RowData(pv_RowNum) = 0 Or (mf_DataChanged) Then
        If Not AddSaveProc Then Err.Raise vbObjectError + 514, 0, "AddSave Failed"
        If MSFg.RowData(pv_RowNum) <> 0 Then
            mo_Recset.Bookmark = MSFg.RowData(pv_RowNum)
            mo_Recset.edit
        Else
            mo_Recset.Addnew
        End If
        
        Call MoveData(pv_RowNum, ctm_Write)
        RaiseEvent InitKey(pv_RowNum)
        If Not MoveUsrDets(ctm_Write) Then Err.Raise vbObjectError + 514, 0, "Error on Updating UserDets"
        
        ValidateServerMode (pv_RowNum)
        mo_Recset.Update
        
        '''''''''''''''''''''''''''' Jen Tp
'        Dim www As MDORowSet
'        Set www = Parent.adc.Connection.OpenResultset("Select SCOPE_IDENTITY() as qNum")
    
'        If Not (www.EOF Or www.BOF) Then
'          Debug.Print www!qNum
'          'www.MoveNext
'        End If
'        Set www = Nothing
        '''''''''''''''''''''''''''' Jen Tp
        
        If MSFg.RowData(pv_RowNum) = 0 Then
            MSFg.RowData(pv_RowNum) = mo_Recset.Bookmark
            Call SetColsInitMode(False)
        End If
        '???If mf_XChild Then Call AutoCommit
        Set mc_AddSave = Nothing
    Else
    
        mo_Recset.CancelUpdate
    End If
    
    Parent.adc.Connection.CommitTrans (mdoCommitNone)
    DispMsg "", etInfo
    RowSave = True
    ParentRef.adc.BtnValFailed = ""     '*** (14/09/05)  Jenny
    Exit Function
    
errh:
   Call errhndlr(pv_RowNum)
   
End Function
Private Function ValidateServerMode(ByVal pv_RowNum As Integer)
  If ms_ValidateStoredProcedure = "" Then
    Exit Function
  End If
  If IsNew(pv_RowNum) Then
    mo_QueryObjectValidate.QryParametersByName("@Mode") = "A"
  Else
    mo_QueryObjectValidate.QryParametersByName("@Mode") = "C"
  End If
    
  If ms_ValidateStoredProcedure <> "" Then
    mo_QueryObjectValidate.QryExecute
  End If
End Function
Private Function AddSaveProc() As Boolean
    Dim i As Integer, ws_ObjNm As String, we_AddSaveOpt As en_AddSaveOpt, j As Integer, e As Integer
    AddSaveProc = True
    If Not (mc_AddSave Is Nothing) Then
        For i = 1 To mc_AddSave.Count
            ws_ObjNm = mc_AddSave(i)(0)
            we_AddSaveOpt = mc_AddSave(i)(1)
            If we_AddSaveOpt = asAll Then
                For j = 1 To Parent.Controls(ws_ObjNm).Rows - 1
                    If Not Parent.Controls(ws_ObjNm).AddSaveRec(j) Then
                        Err.Raise vbObjectError + 514, "", "Save of " + ws_ObjNm + " failed "
                    End If
                Next j
            Else
                If Not Parent.Controls(ws_ObjNm).AddSaveRec Then
                    Err.Raise vbObjectError + 514, "", "Save of " + ws_ObjNm + " failed "
                End If
            End If
'??? consider speed while calling saverec
'??? check for type before using .rows
'??? should I check for isdel status when in ascurrent mode
        Next i
    End If
    Exit Function
errh:
    AddSaveProc = False
    If Not (mc_AddSave Is Nothing) Then
        For e = 1 To i - 1
            ws_ObjNm = mc_AddSave(e)(0)
            we_AddSaveOpt = mc_AddSave(e)(1)
            Parent.Controls(ws_ObjNm).Resync we_AddSaveOpt
'??? consider speed while calling resync
   Next e
    End If
End Function

Private Sub SetColsInitMode(ByVal pv_inInitMode As Boolean)
    Dim wo_reqClsFlx As MWCLS_FLX, i As Integer
    For i = 1 To mc_Columns.Count
        Set wo_reqClsFlx = mc_Columns(i)
        wo_reqClsFlx.inInitMode = pv_inInitMode
    Next i
    Set wo_reqClsFlx = Nothing
End Sub
Private Function convValue(pv_reqDataType As en_FldDataType, pv_Val As Variant)
    If pv_reqDataType = fdtFloat Or pv_reqDataType = fdtNumber Or pv_reqDataType = fdtTimeType Then
        convValue = Val(pv_Val)
    ElseIf pv_reqDataType = fdtDateType Then
        If IsDate(pv_Val) Then
            convValue = CDate(pv_Val)
        Else
            convValue = ""
        End If
    ElseIf pv_reqDataType = fdtCharType Or pv_reqDataType = fdtConvertToUpper Then
        convValue = Trim(pv_Val)
    End If
End Function
Private Function MoveUsrDets(ByVal pv_Opt As String) As Boolean
    On Error GoTo errh
    MoveUsrDets = False
    If pv_Opt = ctm_Write Then
        Store "ModUsr", Parent.adc.UsrCd
        Store "ModDt", Parent.adc.Connection.SrvrDate
        Store "ModTime", Parent.adc.Connection.SrvrTime
    End If
    MoveUsrDets = True
    Exit Function
errh:
    Call Parent.adc.HandleError
End Function
'***************** Subs End ****************************************************************'







