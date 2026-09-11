VERSION 5.00
Begin VB.UserControl MWCTL_MED 
   ClientHeight    =   3600
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4800
   KeyPreview      =   -1  'True
   ScaleHeight     =   3600
   ScaleWidth      =   4800
   Begin VB.Timer TIM_MED 
      Enabled         =   0   'False
      Interval        =   500
      Left            =   2400
      Top             =   1080
   End
   Begin VB.TextBox txt 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   405
      Left            =   0
      MultiLine       =   -1  'True
      TabIndex        =   0
      Top             =   0
      Width           =   1395
   End
End
Attribute VB_Name = "MWCTL_MED"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
Dim mf_ForceLost As Boolean
Dim mv_Data As Variant
Dim mf_MyKey As Boolean
Dim ms_Recalcparent As String
Dim ms_Recalcon As String
Dim mv_Value As Variant
Dim me_DataType As en_FldDataType
Dim ms_DataField As String
Dim me_Align As en_TxtAlign
Dim ms_Mask As String
Dim ms_FormatMask As String
Dim ms_IdName As String
Dim ms_CmpStr As String
Dim mf_AllowNegatives As Boolean
Dim mf_inInitMode As Boolean
Dim mf_InitialiseFld As Boolean

Dim mo_BckColor As OLE_COLOR          '*** Jenny
Dim mo_BckColorSel As OLE_COLOR       '*** Jenny
Dim mo_BckColorDisable As OLE_COLOR   '*** Jenny
Dim mo_ForColor As OLE_COLOR          '*** Jenny
Dim mo_ForColorDisable As OLE_COLOR   '*** Jenny

Event KeyDown(ByRef KeyCode As Integer, ByRef Shift As Integer)
Event KeyUp(ByRef KeyCode As Integer, ByRef Shift As Integer)
Event ReCalculate(ByVal IdName As String, ByVal OldVal As Variant, ByVal RowNum As Integer, ByVal SourceName As String, ByVal SourceColName As String)
Event FldWhen(ByVal IdName As String, ByRef Cancel As Boolean, ByRef ErrMsg As String)
Event FldChange(ByVal IdName As String, ByRef Cancel As Boolean, ByRef ErrMsg As String, ByVal pv_NewValue As Variant)
Event FldValid(ByVal IdName As String, ByVal OldVal As Variant)
'***************** Props Begin ************************************************************'
Public Property Get Enabled() As Boolean
Attribute Enabled.VB_Description = "Returns/sets a value that determines whether an object can respond to user-generated events."
Attribute Enabled.VB_ProcData.VB_Invoke_Property = ";Additional"
    Enabled = txt.Enabled
End Property
Public Property Let Enabled(ByVal vNewValue As Boolean)
    txt.Enabled() = vNewValue
    If vNewValue Then
        '*** (Old Code)
        'txt.BackColor = &H80000005
        'txt.ForeColor = &H80000008
        '*** (Old Code)
        
        '*** Jenny Color
        txt.BackColor = BckColor
        txt.ForeColor = ForColor
        '*** Jenny Color
    Else
        '*** (Old Code)
        'txt.BackColor = &H8000000C
        'txt.ForeColor = &H80000009
        '*** (Old Code)
        
        '*** Jenny Color
        txt.BackColor = BckColorDisable
        txt.ForeColor = ForColorDisable
        '*** Jenny Color
    End If
    UserControl.Enabled = vNewValue
    PropertyChanged "Enabled"
End Property
Public Property Let MaxLength(ByVal vNewValue As Integer)
    txt.MaxLength = vNewValue
    Call Setmaxlength
    PropertyChanged "MaxLength"
End Property
Public Property Get MaxLength() As Integer
Attribute MaxLength.VB_ProcData.VB_Invoke_Property = ";Additional"
    MaxLength = txt.MaxLength
End Property

'*** Jenny
Public Property Let BckColor(ByVal vNewValue As OLE_COLOR)
    mo_BckColor = vNewValue
    If Enabled = True Then txt.BackColor = vNewValue
    'PropertyChanged "BckColor"
End Property
Public Property Get BckColor() As OLE_COLOR
Attribute BckColor.VB_ProcData.VB_Invoke_Property = ";Additional"
    BckColor = mo_BckColor
    'BckColor = txt.BackColor
End Property
Public Property Let ForColor(ByVal vNewValue As OLE_COLOR)
    mo_ForColor = vNewValue
    If Enabled = True Then txt.ForeColor = vNewValue
    'PropertyChanged "ForColor"
End Property
Public Property Get ForColor() As OLE_COLOR
Attribute ForColor.VB_ProcData.VB_Invoke_Property = ";Additional"
    ForColor = mo_ForColor
    'ForColor = txt.ForeColor
End Property
Public Property Let BckColorDisable(ByVal vNewValue As OLE_COLOR)
    mo_BckColorDisable = vNewValue
    If Enabled = False Then txt.BackColor = vNewValue
    'PropertyChanged "BckColorDisable"
End Property
Public Property Get BckColorDisable() As OLE_COLOR
Attribute BckColorDisable.VB_ProcData.VB_Invoke_Property = ";Additional"
    BckColorDisable = mo_BckColorDisable
End Property
Public Property Let ForColorDisable(ByVal vNewValue As OLE_COLOR)
    mo_ForColorDisable = vNewValue
    If Enabled = False Then txt.ForeColor = vNewValue
    'PropertyChanged "ForColorDisable"
End Property
Public Property Get ForColorDisable() As OLE_COLOR
Attribute ForColorDisable.VB_ProcData.VB_Invoke_Property = ";Additional"
    ForColorDisable = mo_ForColorDisable
End Property
Public Property Let BckColorSel(ByVal vNewValue As OLE_COLOR)
    mo_BckColorSel = vNewValue
    PropertyChanged "BckColorSel"
End Property
Public Property Get BckColorSel() As OLE_COLOR
    BckColorSel = mo_BckColorSel
End Property
'*** Jenny

Public Property Let SelStart(New_Value As Long)
    txt.SelStart = New_Value
End Property
Public Property Get SelStart() As Long
Attribute SelStart.VB_MemberFlags = "400"
    SelStart = txt.SelStart
End Property
Public Property Get Value(Optional ByRef pr_Cancel As Boolean) As Variant
Attribute Value.VB_UserMemId = 0
Attribute Value.VB_MemberFlags = "400"
    Value = convValue(mv_Value)
End Property
Public Property Get TxtText() As Variant
Attribute TxtText.VB_MemberFlags = "400"
  If Ambient.UserMode Then TxtText = convValue(txt.Text)
End Property
Public Property Let TxtText(ByVal vNewValue As Variant)
    txt.Text = FormatText(vNewValue)
End Property
Public Property Let Value(Optional ByRef pr_Cancel As Boolean, ByVal vNewValue As Variant)
    On Error GoTo errh
    Dim pr_ErrMsg As String, wv_OldVal As Variant, wo_ActCtl As Control, we_ADCMode As en_XMode, ws_OvFlwVal As String
    we_ADCMode = ParentRef.adc.Mode
        
    'DispMsg "BtnValFailed: " & ParentRef.adc.BtnValFailed, etError    '*** Jenny (Temp)
    pr_Cancel = False
    wv_OldVal = mv_Value
    If we_ADCMode = xClear Or we_ADCMode = xInit Or mf_InitialiseFld Then GoTo AcceptValue
    If me_DataType = fdtFloat Or me_DataType = fdtNumber Or me_DataType = fdtTimeType Then
        ws_OvFlwVal = NumOvrFlwVal(me_DataType, ms_FormatMask)
        If txt = ws_OvFlwVal Then
            wv_OldVal = ws_OvFlwVal
        End If
    End If
    vNewValue = convValue(vNewValue)
    If (Value <> vNewValue) Or (mf_inInitMode) Then
'??? use mv_value or get text
        If me_DataType = fdtNumber Or me_DataType = fdtFloat Or me_DataType = fdtTimeType Then
            vNewValue = Format(MWLib.NVal(vNewValue), ms_FormatMask)
            Call ChkValidFormats(ms_IdName, me_DataType, ms_FormatMask, True, vNewValue, pr_Cancel, pr_ErrMsg)
            If pr_Cancel = True And EnAndVis Then Err.Raise vbObjectError + 513, "", pr_ErrMsg
            If pr_Cancel = True Then vNewValue = NumOvrFlwVal(me_DataType, ms_FormatMask)
        End If
        If pr_Cancel = False Then
            RaiseEvent FldChange(ms_IdName, pr_Cancel, pr_ErrMsg, vNewValue)
            If pr_Cancel = True Then Err.Raise vbObjectError + 513, "", pr_ErrMsg
        End If
AcceptValue:
        vNewValue = FormatText(vNewValue)
        mv_Value = vNewValue
        If pr_Cancel = False And we_ADCMode <> xClear Then ReCalcProc (wv_OldVal)
        ' for key part validations
        
        'ParentRef.adc.BtnValFailed = ""     '*** (13/09/05)  Jenny
        
        If we_ADCMode = xClear Then mf_inInitMode = True
        ' for non key part validataions
        If we_ADCMode = xInit And ParentRef.adc.AddFndMode = xaddmode Then mf_inInitMode = True
        If we_ADCMode = xInit And ParentRef.adc.AddFndMode = xFndMode Then mf_inInitMode = False
        ' do not change the order of setting mf_inInitMode above. (has to be for xClear mode first,then xInitMode )
        If mf_inInitMode = True And (we_ADCMode = xNorm Or we_ADCMode = xCopy) Then mf_inInitMode = False
    End If
    ParentRef.adc.BtnValFailed = ""     '*** (13/09/05)  Jenny
    
    txt.Text = FormatText(vNewValue)
    Exit Property
errh:
    Call ParentRef.adc.HandleError
    Set wo_ActCtl = m_ActScrn.ActiveControl
    If (wo_ActCtl Is Nothing) Then Exit Property
    If TypeOf wo_ActCtl Is MWCTL_MED Then
      If wo_ActCtl.IdName = ms_IdName Then Exit Property
    End If
    If Not EnAndVis Then Exit Property
 '   Me.SetFocus
    UserControl.Extender.SetFocus
End Property
Public Property Get EnAndVis() As Boolean
Attribute EnAndVis.VB_MemberFlags = "400"
                          ' returns true if the Med or its Container till Form is...
    Dim wobj As Variant
    If Not (Enabled Or Visible) Then
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
Public Property Get DataType() As en_FldDataType
Attribute DataType.VB_ProcData.VB_Invoke_Property = ";Additional"
    DataType = me_DataType
End Property
Public Property Let DataType(ByVal vNewValue As en_FldDataType)
    me_DataType = vNewValue
    Call Setmaxlength
    Call SetMask                ' for timetype
    PropertyChanged "DataType"
End Property
Public Property Get DataField() As String
Attribute DataField.VB_ProcData.VB_Invoke_Property = "MWPAG_DATA;Additional"
    DataField = ms_DataField
End Property
Public Property Let DataField(ByVal vNewValue As String)
    ms_DataField = Trim(vNewValue)
    If ms_IdName = "" Then ms_IdName = vNewValue
    If ParentRef.adc.EntType = en_XEntType.xetReport Then
        ms_DataField = ms_IdName
    End If
    PropertyChanged "DataField"
End Property
Public Property Get ReCalcParent() As String
Attribute ReCalcParent.VB_MemberFlags = "400"
    ReCalcParent = ms_Recalcparent
End Property
Public Property Let ReCalcParent(ByVal vNewValue As String)
    ms_Recalcparent = vNewValue
    PropertyChanged "ReCalcParent"
End Property
Public Property Get ReCalcOn() As String
Attribute ReCalcOn.VB_ProcData.VB_Invoke_Property = "MWPAG_MED;Additional"
    ReCalcOn = ms_Recalcon
End Property
Public Property Let ReCalcOn(ByVal vNewValue As String)
    On Error GoTo errh
    Dim ws_OldVal As String, wv_rArr As Variant, ws_ReqStr As String, wi_FirBr As Integer, wi_SecBr As Integer, ws_reqPrm As String, wo_ReqObj As Object, ws_rStr As String, ws_AddStr As String, i As Integer, wo_ReqFlx As MWCLS_FLX, wo_reqMED As MWCTL_MED
    If ms_Recalcon = vNewValue Then Exit Property
    If ms_Recalcon <> vNewValue And ms_Recalcon <> "" Then
        Call Rmv_Frm_Parent
    End If
    ws_OldVal = ms_Recalcon
    ms_Recalcon = vNewValue
    PropertyChanged "RecalcOn"
    If ms_Recalcon = "" Then Exit Property
    wv_rArr = Split(ms_Recalcon, ",")
    If Not IsArray(wv_rArr) Then Exit Property
    For i = 0 To UBound(wv_rArr)
        ws_ReqStr = wv_rArr(i)
        wi_FirBr = InStr(1, ws_ReqStr, "(")
        wi_SecBr = InStr(1, ws_ReqStr, ")")
        If wi_FirBr <> 0 Then
            ws_reqPrm = Mid(ws_ReqStr, wi_FirBr + 1, wi_SecBr - wi_FirBr - 1)
            Set wo_ReqObj = ParentRef.Controls(Mid(ws_ReqStr, 1, wi_FirBr - 1))
        Else
            Set wo_ReqObj = ParentRef.adc(ws_ReqStr)
        End If
        If TypeOf wo_ReqObj Is MWCTL_FLX Then
            Set wo_ReqFlx = wo_ReqObj.ColProp(ws_reqPrm)
            ws_rStr = wo_ReqFlx.ReCalcParent
'??? no check for typeof
        Else
            Set wo_reqMED = wo_ReqObj
            ws_rStr = wo_reqMED.ReCalcParent
        End If
        ws_AddStr = ms_IdName
        If TypeOf wo_ReqObj Is MWCTL_FLX Then
            wo_ReqFlx.ReCalcParent = IIf(ws_rStr = "", "", ws_rStr + ",") + ws_AddStr
        Else
            wo_reqMED.ReCalcParent = IIf(ws_rStr = "", "", ws_rStr + ",") + ws_AddStr
        End If
    Next i
    Exit Property
errh:
'??? ErrHndling
    MsgBox "Cannot find obj" + ws_ReqStr
    ms_Recalcon = ""
End Property
Property Get ParentRef() As Object
Attribute ParentRef.VB_MemberFlags = "400"
    If TypeOf Parent Is MWCTL_FLX Then
        Set ParentRef = Parent.ParentRef
    Else
        Set ParentRef = Parent
    End If
End Property
Public Property Get Alignment() As en_TxtAlign
Attribute Alignment.VB_ProcData.VB_Invoke_Property = ";Additional"
    Alignment = me_Align
End Property
Public Property Let Alignment(ByVal vNewValue As en_TxtAlign)
    me_Align = vNewValue
    Select Case me_Align
    Case Is = tLeftAlign
        txt.Alignment = 0
    Case Is = tRightAlign
        txt.Alignment = 1
    Case Is = tMultiLine
        txt.Alignment = 0
    End Select
    PropertyChanged "Alignment"
End Property
Public Property Get IdName() As String
Attribute IdName.VB_ProcData.VB_Invoke_Property = ";Additional"
    IdName = ms_IdName
End Property
Public Property Let IdName(ByVal vNewValue As String)
    If Ambient.UserMode Then
        MsgBox "Cannot set at RunTime"
        Exit Property
    End If
'errorhndlng
'???If ms_IdName <> "" And UCase(ms_IdName) <> UCase(vNewValue) Then
'        If ms_recalcon <> "" Or ms_recalcparent <> "" Then
'            ErrMsg = "Cannot change Idname as "
'            ErrMsg = ErrMsg + IIf(ms_RecalcOn <> "", "Fld Recalculated on change of " + ms_RecalcOn, "")
'            ErrMsg = ErrMsg + IIf(ms_Recalcparent <> "", "Triggering Recalc of " + ms_Recalcparent, "")
'            MsgBox ErrMsg
'            Exit Property
'        End If
'    End If
    ms_IdName = UCase(vNewValue)
    If ParentRef.adc.EntType = en_XEntType.xetReport Then
        ms_DataField = ms_IdName
    End If
    PropertyChanged "IdName"
End Property
Public Property Get Data() As Variant
Attribute Data.VB_ProcData.VB_Invoke_Property = ";Additional"
    Data = mv_Data
End Property
Public Property Let Data(ByVal vNewValue As Variant)
    If Ambient.UserMode Then Exit Property
    mv_Data = vNewValue
    txt.Text = mv_Data
    PropertyChanged "Data"
End Property
Public Property Get Visible() As Boolean
    Visible = UserControl.Extender.Visible
End Property
Public Property Let Visible(ByVal vNewValue As Boolean)
    UserControl.Extender.Visible = vNewValue
End Property
Public Property Get IsBlank() As Boolean
Attribute IsBlank.VB_MemberFlags = "400"
    Select Case me_DataType
    Case Is = fdtCharType, fdtConvertToUpper, fdtDateType
        If Value = "" Then IsBlank = True
    Case Is = fdtFloat, fdtNumber, fdtTimeType
        If Value = 0 Then IsBlank = True
    End Select
End Property
Public Property Get TxtIsBlank() As Boolean
Attribute TxtIsBlank.VB_MemberFlags = "400"
                          'returns from Txt.Text, mainly used by Hlp
    Select Case me_DataType
    Case Is = fdtCharType, fdtConvertToUpper, fdtDateType
        If TxtText = "" Then TxtIsBlank = True
    Case Is = fdtFloat, fdtNumber, fdtTimeType
        If TxtText = 0 Then TxtIsBlank = True
    End Select
End Property
Public Property Get CmpStr() As String
Attribute CmpStr.VB_ProcData.VB_Invoke_Property = "MWPAG_REP;Additional"
    CmpStr = ms_CmpStr
End Property
Public Property Let CmpStr(ByVal vNewValue As String)
    ms_CmpStr = vNewValue
    PropertyChanged "CmpStr"
End Property
Public Property Get Mask() As String
Attribute Mask.VB_ProcData.VB_Invoke_Property = ";Additional"
    Mask = ms_Mask
End Property
Public Property Let Mask(ByVal vNewValue As String)
    mf_AllowNegatives = False
    ms_FormatMask = ""            ' is mask w/o -ve
    ms_Mask = vNewValue
    Call SetMask
    If Mid(ms_Mask, 1, 1) = "-" Then
      ms_FormatMask = Mid(ms_Mask, 2)
      mf_AllowNegatives = True
    Else
      ms_FormatMask = ms_Mask
    End If
    Call Setmaxlength
    PropertyChanged "Mask"
End Property
Public Property Get Locked() As Boolean
Attribute Locked.VB_MemberFlags = "400"
    Locked = txt.Locked
End Property
Public Property Let Locked(ByVal vNewValue As Boolean)
    txt.Locked = vNewValue            ' readonly
End Property
'***************** Props End ************************************************************'
'***************** Events Begin ************************************************************'
Private Sub TIM_MED_Timer()
    If ParentRef.MWCTL_HLP1.Mode = hmHlpAvlbl And ParentRef.MWCTL_HLP1.Visible Then ParentRef.MWCTL_HLP1.FndMatch
    TIM_MED.Enabled = False
End Sub
Private Sub Txt_KeyDown(KeyCode As Integer, Shift As Integer)
    If m_ActScrn.ActiveForm.WindowState = vbMinimized Then
        KeyCode = 0
        Shift = 0
       Exit Sub
    End If
    mf_MyKey = True
    If me_DataType = fdtFloat Or me_DataType = fdtDateType Or me_DataType = fdtTimeType Then
        If KeyCode = vbKeyLeft Or KeyCode = vbKeyBack Then
            If txt.SelStart > 0 Then
                If Mid(txt, txt.SelStart, 1) = "/" Or Mid(txt, txt.SelStart, 1) = "." Then
                    txt.SelStart = txt.SelStart - 1
                    KeyCode = 0
                    Shift = 0
                End If
            End If
        Else
            If KeyCode <> vbKeyHome And KeyCode <> vbKeyEnd Then
                If Mid(txt, txt.SelStart + 1, 1) = "/" Or Mid(txt, txt.SelStart + 1, 1) = "." Then
                    txt.SelStart = txt.SelStart + 1
                    KeyCode = 0
                    Shift = 0
                End If
            End If
        End If
    End If
    If (KeyCode = vbKeyUp Or KeyCode = vbKeyDown Or KeyCode = vbKeyPageDown Or KeyCode = vbKeyPageUp) And ParentRef.MWCTL_HLP1.Visible Then
        Call ParentRef.MWCTL_HLP1.FndMatch(KeyCode)
        KeyCode = 0
    End If
    RaiseEvent KeyDown(KeyCode, Shift)      'in Grd
End Sub
Private Sub Txt_KeyPress(KeyAscii As Integer)
    If m_ActScrn.ActiveForm.WindowState = vbMinimized Then
        KeyAscii = 0
        Exit Sub
    End If
    If Not mf_MyKey Then Exit Sub
    
    If KeyAscii = vbKeyReturn And (Not ParentRef.MWCTL_HLP1.Visible) And me_Align <> tMultiLine Then
        KeyAscii = 0
        Exit Sub
    End If
    
    '*** (Jenny 23/05/07) (Do not allow Single Quots in 'Convert To Upper' MED controls)
    If KeyAscii = 39 And me_DataType = fdtConvertToUpper Then
      If ParentRef.adc.EntType = en_XEntType.xetEntry Then
        KeyAscii = 0
        Exit Sub
      End If
    End If
    '*** (Jenny 23/05/07) (Do not allow Single Quots in 'Convert To Upper' MED controls)
    
    Dim wf_ValidKey As Boolean, wi_DecPos As Integer, wi_IntLen As Integer, wi_NegPos As Integer, wv_rStr As Variant, wi_Pos As Integer
    If me_DataType = fdtFloat Or me_DataType = fdtNumber Or me_DataType = fdtDateType Or me_DataType = fdtTimeType Then
        wf_ValidKey = False
        If (Chr(KeyAscii) Like "[0-9]") Then wf_ValidKey = True
        If mf_AllowNegatives Then
          If Chr(KeyAscii) = "-" Then
            If InStr(1, txt, "-") = 0 And Trim(Left(txt, txt.SelStart)) = "" Then
              wf_ValidKey = True
            End If
          Else
            If (Chr(KeyAscii)) Like "[0-9]" Then
                wf_ValidKey = True
                wi_DecPos = InStr(1, ms_FormatMask, ".")
                wi_IntLen = IIf(wi_DecPos = 0, Len(ms_FormatMask), wi_DecPos - 1)
                If InStr(1, txt, "-") = 0 And (txt.SelStart >= wi_IntLen And (wi_DecPos = 0 Or txt.SelStart < wi_DecPos)) Then
                  wf_ValidKey = False
                End If
                wi_NegPos = InStr(1, txt, "-")
                If wi_NegPos <> 0 And txt.SelStart < wi_NegPos Then
                  wf_ValidKey = False
                End If
            End If
          End If
        End If
        If (Not txt.Locked) And wf_ValidKey Then
            wv_rStr = txt.Text
            If wv_rStr <> "" Then
                wi_Pos = txt.SelStart + 1
                If wi_Pos <= Len(txt) Then
                    Mid(wv_rStr, wi_Pos) = Chr(KeyAscii)
                    txt = wv_rStr
                    txt.SelStart = wi_Pos
                    KeyAscii = 0
                End If
            End If
        Else
            If Chr(KeyAscii) Like "[A-z]" Or Chr(KeyAscii) = "-" Then
                KeyAscii = 0
            End If
        End If
        If (me_DataType = fdtFloat Or me_DataType = fdtTimeType) And KeyAscii = 46 Then
           txt = FormatText
           wi_DecPos = InStr(1, txt, ".")
           If wi_DecPos > 0 Then
               txt.SelStart = wi_DecPos
               KeyAscii = 0
            End If
        End If
    End If
    If me_DataType = fdtConvertToUpper Then
        KeyAscii = Asc(UCase(Chr(KeyAscii)))
    End If
End Sub
Private Sub txt_KeyUp(KeyCode As Integer, Shift As Integer)
    Dim wi_Cur As Integer, wi_IntLen As Integer, wi_DecPos As Integer, ws_IntEnt As String, ws_DecEnt As String, ws_rStr As String, i As Integer
    RaiseEvent KeyUp(KeyCode, Shift)
    If Not mf_MyKey Then Exit Sub
    mf_MyKey = False
    If KeyCode = vbKeyReturn And Shift = 2 Then Call ReCalc("", "ZUDEF", "")
    
    '*** (Jenny 24/05/07) (Do not allow Single Quots in 'Convert To Upper' MED controls if the user presses Ctrl+V)
    If KeyCode = 86 And Shift = 2 And me_DataType = fdtConvertToUpper Then
      If ParentRef.adc.EntType = en_XEntType.xetEntry Then txt = Replace(txt.Text, "'", "")
    End If
    '*** (Jenny 24/05/07) (Do not allow Single Quots in 'Convert To Upper' MED controls if the user presses Ctrl+V)
    
    Select Case me_DataType
    Case Is = fdtDateType
        If KeyCode = vbKeyDelete Or KeyCode = vbKeyBack Then
            wi_Cur = txt.SelStart
            For i = 1 To 8
                If Len(ws_rStr) = 2 Or Len(ws_rStr) = 5 Then
                    ws_rStr = ws_rStr + "/"
                End If
                If i <= Len(txt) Then
                    If Mid(txt, i, 1) <> "/" Then
                       ws_rStr = ws_rStr + Mid(txt, i, 1)
                    End If
                Else
                    ws_rStr = ws_rStr + " "
                End If
            Next i
            txt = ws_rStr
            txt.SelStart = wi_Cur
        End If
    Case Is = fdtFloat, fdtTimeType
        If KeyCode = vbKeyDelete Or KeyCode = vbKeyBack Then
            wi_Cur = txt.SelStart
            wi_IntLen = InStr(1, ms_Mask, ".") - 1
            wi_DecPos = InStr(1, txt, ".")
            ws_IntEnt = txt
            ws_DecEnt = ""
            If wi_DecPos > 0 Then
                ws_IntEnt = Mid(txt, 1, wi_DecPos - 1)
                ws_DecEnt = Mid(txt, wi_DecPos + 1)
            End If
            ws_DecEnt = ws_DecEnt + Space((Len(ms_Mask) - InStr(1, ms_Mask, ".")) - Len(ws_DecEnt))
            ws_rStr = ws_IntEnt + Space(wi_IntLen - Len(ws_IntEnt)) + "." + ws_DecEnt
            txt = ws_rStr
            txt.SelStart = wi_Cur
        End If
    Case Is = fdtNumber
        If KeyCode = vbKeyDelete Or KeyCode = vbKeyBack Then
            wi_Cur = txt.SelStart
            wi_IntLen = Len(ms_Mask)
            For i = 1 To Len(ms_Mask)
                If i <= Len(txt) Then
                    ws_rStr = ws_rStr + Mid(txt, i, 1)
                Else
                    ws_rStr = ws_rStr + " "
                End If
            Next i
            txt = ws_rStr
            txt.SelStart = wi_Cur
        End If
    End Select
    If ParentRef.MWCTL_HLP1.Mode = hmHlpAvlbl Then
        If KeyCode = vbKeyF2 Then
            Call StartHlp
        End If
        If KeyCode <> vbKeyF2 And (KeyCode = vbKeyDelete Or KeyCode = vbKeyBack Or (Shift = 0 And ((KeyCode >= 96 And KeyCode <= 105) Or Chr(KeyCode) Like "[A-z]" Or Chr(KeyCode) Like "[0-9]"))) Then
            TIM_MED.Enabled = False
            TIM_MED.Enabled = True
        End If
        If (Not txt.Locked) And ParentRef.MWCTL_HLP1.Visible And KeyCode = vbKeyReturn Then
            Call ParentRef.MWCTL_HLP1.ReplSelVal
            txt.SelStart = Len(txt.Text)
        End If
        If ParentRef.MWCTL_HLP1.MultiSelect Then
            If KeyCode = vbKeyEscape Then ParentRef.MWCTL_HLP1.Visible = False
        Else
            If KeyCode = vbKeyReturn Or KeyCode = vbKeyEscape Then ParentRef.MWCTL_HLP1.Visible = False
        End If
    End If
End Sub
Private Sub txt_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
  '*** (Jenny 24/05/07) (Do not allow Single Quots in 'Convert To Upper' MED controls if the user does copy & paste using mouse)
  If txt.Text <> "" And InStr(txt.Text, "'") > 0 And me_DataType = fdtConvertToUpper Then
    If ParentRef.adc.EntType = en_XEntType.xetEntry Then txt.Text = Replace(txt.Text, "'", "")
  End If
  '*** (Jenny 24/05/07) (Do not allow Single Quots in 'Convert To Upper' MED controls if the user does copy & paste using mouse)
End Sub
Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
    mf_AllowNegatives = False
    mf_inInitMode = False
    ms_FormatMask = ""
    mf_InitialiseFld = False
    Alignment = PropBag.ReadProperty("Alignment", en_TxtAlign.tLeftAlign)
    DataType = PropBag.ReadProperty("DataType", en_FldDataType.fdtCharType)
    Enabled = PropBag.ReadProperty("Enabled", True)
    MaxLength = PropBag.ReadProperty("MaxLength", 1)
    
    '*** Jenny
    BckColor = PropBag.ReadProperty("BckColor", &H80000005)
    mo_BckColor = BckColor
    ForColor = PropBag.ReadProperty("ForColor", &H80000008)
    mo_ForColor = ForColor
    BckColorDisable = PropBag.ReadProperty("BckColorDisable", &H8000000C)
    mo_BckColorDisable = BckColorDisable
    ForColorDisable = PropBag.ReadProperty("ForColorDisable", &H80000009)
    mo_ForColorDisable = ForColorDisable
    mo_BckColorSel = PropBag.ReadProperty("BckColorSel", &H80000005)
    '*** Jenny
    
    Mask = PropBag.ReadProperty("Mask", "")
    ms_DataField = PropBag.ReadProperty("DataField", "")
    ms_Recalcparent = PropBag.ReadProperty("ReCalcParent", "")
    ms_Recalcon = PropBag.ReadProperty("ReCalcOn", "")
    ms_IdName = PropBag.ReadProperty("IdName", "")
    mv_Data = PropBag.ReadProperty("Data", "")
    ms_CmpStr = PropBag.ReadProperty("CmpStr", "")
'???    Call InitVal
' is this reqd inspite of being there in gotfocus
End Sub
Private Sub UserControl_Resize()
    txt.Left = 0
    txt.Top = 0
    txt.Height = UserControl.Height
    txt.Width = UserControl.Width
End Sub
'Write property values to storage
Private Sub UserControl_WriteProperties(PropBag As PropertyBag)
    Call PropBag.WriteProperty("Alignment", me_Align, tLeftAlign)
    Call PropBag.WriteProperty("Enabled", Enabled, True)
    Call PropBag.WriteProperty("Mask", ms_Mask, "")
    Call PropBag.WriteProperty("MaxLength", MaxLength, 1)
    
    '*** Jenny
    'Call PropBag.WriteProperty("BckColor", txt.BackColor, &H80000005)
    'Call PropBag.WriteProperty("ForColor", txt.ForeColor, &H80000008)
    Call PropBag.WriteProperty("BckColor", mo_BckColor, &H80000005)
    Call PropBag.WriteProperty("ForColor", mo_ForColor, &H80000008)
    Call PropBag.WriteProperty("BckColorSel", mo_BckColorSel, &H80000005)
    Call PropBag.WriteProperty("BckColorDisable", mo_BckColorDisable, &H8000000C)
    Call PropBag.WriteProperty("ForColorDisable", mo_ForColorDisable, &H80000009)
    '*** Jenny
    
    Call PropBag.WriteProperty("DataType", me_DataType, fdtCharType)
    Call PropBag.WriteProperty("DataField", ms_DataField, "")
    Call PropBag.WriteProperty("ReCalcParent", ms_Recalcparent, "")
    Call PropBag.WriteProperty("ReCalcOn", ms_Recalcon, "")
    Call PropBag.WriteProperty("IdName", ms_IdName, "")
    Call PropBag.WriteProperty("Data", mv_Data, "")
    Call PropBag.WriteProperty("CmpStr", ms_CmpStr, "")
End Sub
Private Sub txt_GotFocus()
    Dim pr_Cancel As Boolean, pr_ErrMsg As String
    On Error GoTo errh
    If TypeOf Parent Is MWCTL_FLX Then Exit Sub
    If ParentRef.adc.ValFailed <> "" Then
        If ms_IdName = ParentRef.adc.ValFailed Then ParentRef.adc.ValFailed = ""
        Exit Sub
    End If
    If ParentRef.MWCTL_HLP1.Mode = hmHlpNotavlbl Then
        DispMsg txt.ToolTipText, etInfo
        Call InitVal
        txt.SelStart = 0
        pr_Cancel = False
        pr_ErrMsg = ""
        txt.Locked = False
        
        '*** Jenny
        'txt.BackColor = mo_BckColorSel     '*** This code is proper (this is to highlight the selected field)
        '*** Jenny
        RaiseEvent FldWhen(UCase(ms_IdName), pr_Cancel, pr_ErrMsg)
        If pr_Cancel = True Then Err.Raise vbObjectError + 514, "", pr_ErrMsg
    End If
    Exit Sub
errh:
    Call ParentRef.adc.HandleError
    txt.Locked = True
End Sub
Private Sub txt_LostFocus()
    If m_ActScrn.ActiveControl Is Nothing Then Exit Sub
    If TypeOf Parent Is MWCTL_FLX Then Exit Sub
    If TypeOf m_ActScrn.ActiveControl Is MWCTL_HLP Then
        ParentRef.MWCTL_HLP1.Mode = hmHlpBrowsing
        Exit Sub
    End If
    
    '*** Jenny
    'txt.BackColor = mo_BckColor      '*** This code is proper (this is to unselect a field on lostfocus)
    '*** Jenny
    
    If TypeOf m_ActScrn.ActiveControl Is MWCTL_ADC Then
        Select Case ParentRef.adc.OptSel
        Case Is = "A", "F", "S"
            If Validate = True Then
                ParentRef.adc.ValFailed = ms_IdName
            End If
        End Select
    End If
    ParentRef.MWCTL_HLP1.Mode = hmHlpNotavlbl
    ParentRef.MWCTL_HLP1.Visible = False
End Sub
'************************** Events End ******************************************************'
'***************** Methods Begin ************************************************************'
Public Sub StartHlp()
    If m_ActScrn.ActiveForm.WindowState = vbMinimized Then Exit Sub
    If ParentRef.MWCTL_HLP1.Mode = hmHlpAvlbl And (Not ParentRef.MWCTL_HLP1.Visible) Then
        Set ParentRef.MWCTL_HLP1.RetInto = m_ActScrn.ActiveControl
        ParentRef.MWCTL_HLP1.Gen_Hlp
        ParentRef.MWCTL_HLP1.ZOrder
        ParentRef.MWCTL_HLP1.Visible = True
    End If
End Sub
Public Function FormatText(Optional vNewValue As Variant)
    If IsMissing(vNewValue) Then vNewValue = txt
    Select Case me_DataType
        Case Is = fdtNumber, fdtFloat, fdtTimeType
            vNewValue = Format(MWLib.NVal(vNewValue), ms_FormatMask)
            vNewValue = Space(Len(ms_Mask) - Len(vNewValue)) + Trim(vNewValue)
        Case Is = fdtDateType
            vNewValue = IIf(IsDate(vNewValue), Format(vNewValue, MWLib.GetDateFormatForEntry), "  /  /  ")
    End Select
    FormatText = vNewValue
End Function
Public Sub InitialiseFld(ByVal vNewValue As Variant)
    If Not Ambient.UserMode Then Exit Sub
    mf_InitialiseFld = True
    Value = vNewValue
    mf_InitialiseFld = False
End Sub
Public Sub InitVal()
                        ' called by GotFocus
    If Not TxtIsBlank Then Exit Sub
    Dim wi_DecPos As Integer
    If Not Ambient.UserMode Then Exit Sub
    Select Case me_DataType
        Case Is = fdtDateType
            txt = "  /  /  "
        Case Is = fdtFloat, fdtTimeType
            wi_DecPos = InStr(1, ms_Mask, ".")
            If wi_DecPos <> 0 Then
                txt = Space(wi_DecPos - 1) + "." + Space(Len(ms_Mask) - wi_DecPos)
            Else
                txt = ""
            End If
        Case Is = fdtNumber
            txt = Space(Len(ms_Mask))
    End Select
End Sub
Public Sub ReCalc(ByVal pv_OldVal As Variant, ByVal pv_SourceName As String, Optional ByVal pv_SourceColName As String, Optional ByVal pv_RowNum As Integer)
    RaiseEvent ReCalculate(ms_IdName, pv_OldVal, pv_RowNum, pv_SourceName, pv_SourceColName)
End Sub
Public Sub SetFocus()
    UserControl.Extender.SetFocus
End Sub
Public Function Validate() As Boolean
    Dim wf_Cancel As Boolean, wv_OldVal As Variant
    wv_OldVal = Value
    Value(wf_Cancel) = TxtText
'    If wf_Cancel = False Then
'        Parent.MWCTL_HLP1.Mode = hmHlpNotavlbl
'        ParentRef.MWCTL_HLP1.Visible = False
'    End If
    Validate = wf_Cancel
    If wf_Cancel = False Then RaiseEvent FldValid(ms_IdName, wv_OldVal)
End Function
Public Sub SetToolTip()
                        ' sets the tooltip from Med to txt, Called by Adc.InitRtn
    txt.ToolTipText = UserControl.Extender.ToolTipText
End Sub
'***************** Methods End ************************************************************'
'***************** Subs Begin ************************************************************'
Private Sub AutomateAdFnd()
    If ParentRef.adc.Automate Then
        If UserControl.Extender.Container.Name = "FRA_KEYS" Then
            If m_ActScrn.ActiveControl.Container.Name <> "FRA_KEYS" And m_ActScrn.ActiveControl.Name <> "ADC" Then
                Call ParentRef.adc.AutomateAdFnd
            End If
        End If
    End If
End Sub
Private Sub Setmaxlength()
    Select Case me_DataType
      Case Is = fdtDateType
          txt.MaxLength = 8
      Case Is = fdtFloat, fdtNumber, fdtTimeType
          txt.MaxLength = Len(ms_Mask)
    End Select
End Sub
Private Sub SetMask()
    If me_DataType = fdtTimeType Then ms_Mask = "#0.00"
End Sub
Private Sub ReCalcProc(ByVal pv_OldVal As Variant)
    Dim wv_rArr As Variant, ws_ReqStr As String, wi_FirBr As Integer, wi_SecBr As Integer, ws_reqPrm As String, wo_ReqObj As Object, wo_ReqFlx As MWCTL_FLX, wo_reqMED As MWCTL_MED, i As Integer, j As Integer
    wv_rArr = Split(ms_Recalcparent, ",")
    If Not IsArray(wv_rArr) Then Exit Sub
    For i = 0 To UBound(wv_rArr)
        ws_ReqStr = wv_rArr(i)
        wi_FirBr = InStr(1, ws_ReqStr, "(")
        wi_SecBr = InStr(1, ws_ReqStr, ")")
        If wi_FirBr <> 0 Then
            ws_reqPrm = Mid(ws_ReqStr, wi_FirBr + 1, wi_SecBr - wi_FirBr - 1)
            Set wo_ReqObj = ParentRef.Controls(Mid(ws_ReqStr, 1, wi_FirBr - 1))
        Else
            Set wo_ReqObj = ParentRef.adc(ws_ReqStr)
        End If
        If TypeOf wo_ReqObj Is MWCTL_FLX Then
            Set wo_ReqFlx = wo_ReqObj
            wo_ReqFlx.ReCalc ws_reqPrm, pv_OldVal, ms_IdName, ""
        ElseIf TypeOf wo_ReqObj Is MWCTL_MED Then
            Set wo_reqMED = wo_ReqObj
            wo_reqMED.ReCalc pv_OldVal, ms_IdName, ""
        End If
    Next i
End Sub
Private Sub Rmv_Frm_Parent()
    On Error GoTo errh
    Dim wv_rArr As Variant, ws_ReqStr As String, wi_FirBr As Integer, wi_SecBr As Integer, ws_reqPrm As String, wo_ReqObj As Object, wo_ReqFlx As MWCLS_FLX, wo_reqMED As MWCTL_MED, ws_rStr As String, ws_AddStr As String, wi_rPos As Integer, i As Integer
    wv_rArr = Split(ms_Recalcon, ",")
    If Not IsArray(wv_rArr) Then Exit Sub
    For i = 0 To UBound(wv_rArr)
        ws_ReqStr = wv_rArr(i)
        wi_FirBr = InStr(1, ws_ReqStr, "(")
        wi_SecBr = InStr(1, ws_ReqStr, ")")
        If wi_FirBr <> 0 Then
            ws_reqPrm = Mid(ws_ReqStr, wi_FirBr + 1, wi_SecBr - wi_FirBr - 1)
            Set wo_ReqObj = ParentRef.Controls(Mid(ws_ReqStr, 1, wi_FirBr - 1))
        Else
            Set wo_ReqObj = ParentRef.adc(ws_ReqStr)
        End If
        If TypeOf wo_ReqObj Is MWCTL_FLX Then
            Set wo_ReqFlx = wo_ReqObj.ColProp(ws_reqPrm)
            ws_rStr = "," + wo_ReqFlx.ReCalcParent + ","
        ElseIf TypeOf wo_ReqObj Is MWCTL_MED Then
            Set wo_reqMED = wo_ReqObj
            ws_rStr = "," + wo_reqMED.ReCalcParent + ","
        End If
        ws_AddStr = ms_IdName
        wi_rPos = InStr(1, ws_rStr, "," + ws_AddStr + ",")
        If wi_rPos <> 0 Then
            ws_rStr = Mid(ws_rStr, 1, wi_rPos - 1) + Mid(ws_rStr, wi_rPos + Len(ws_AddStr) + 1)
            If ws_rStr <> "" Then
                If Left(ws_rStr, 1) = "," Then ws_rStr = Mid(ws_rStr, 2)
                If Right(ws_rStr, 1) = "," Then ws_rStr = Left(ws_rStr, Len(ws_rStr) - 1)
            End If
            If TypeOf wo_ReqObj Is MWCTL_FLX Then
                wo_ReqFlx.ReCalcParent = ws_rStr
            ElseIf TypeOf wo_ReqObj Is MWCTL_MED Then
                wo_reqMED.ReCalcParent = ws_rStr
            End If
        End If
    Next i
    Exit Sub
errh:
    Resume Next
End Sub
Private Function convValue(ByVal vNewValue As Variant) As Variant
    Dim wi_Pos As Integer
    If me_DataType = fdtFloat Or me_DataType = fdtNumber Or me_DataType = fdtTimeType Then
        convValue = MWLib.NVal(vNewValue)
    ElseIf me_DataType = fdtDateType Then
        If IsDate(vNewValue) Then
          convValue = CDate(vNewValue)
        Else
          convValue = ""
        End If
    ElseIf me_DataType = fdtCharType Or me_DataType = fdtConvertToUpper Then
        convValue = Trim(vNewValue)
        If me_Align <> tMultiLine Then
            wi_Pos = InStr(1, vNewValue, Chr(13))
            If wi_Pos <> 0 Then
                convValue = Mid(convValue, 1, wi_Pos - 1)
            End If
        End If
    End If
End Function
'***************** Subs End ****************************************************************'
