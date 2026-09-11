VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "MWCLS_FLX"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
Dim ms_Mask As String
Dim me_DataType As en_FldDataType
Dim ms_Colname As String
Dim ms_Datafld As String
Dim ms_Heading As String
Dim mi_Maxlength As Integer
Dim ms_Recalcon As String
Dim ms_Recalcparent As String
Dim mf_ReadOnly As Boolean        'not used anywhere
Dim mo_Parent As Object
Dim ml_Width As Long
Dim ms_ToolTipText As String
Dim me_Style As en_fgColStyle
Dim ms_FormatMask As String
Dim mf_AllowNegatives As Boolean
Dim mf_inInitMode As Boolean
Public Property Get Mask() As String
    Mask = ms_Mask
End Property
Public Property Get FormatMask() As String
    FormatMask = ms_FormatMask
End Property
Public Property Let Mask(ByVal vNewValue As String)
    If me_DataType = fdtTimeType Then vNewValue = "#0.00"
    ms_Mask = vNewValue
    ms_FormatMask = ""
    mf_AllowNegatives = False
    If Mid(ms_Mask, 1, 1) = "-" Then
      ms_FormatMask = Mid(ms_Mask, 2)
      mf_AllowNegatives = True
    Else
      ms_FormatMask = ms_Mask
    End If
End Property
Public Property Get Style() As en_fgColStyle
    Style = me_Style
End Property
Public Property Let Style(ByVal vNewValue As en_fgColStyle)
    me_Style = vNewValue
    mo_Parent.PropertyModified "style", ms_Colname
End Property
Public Property Get DataType() As en_FldDataType
    DataType = me_DataType
End Property
Public Property Let DataType(ByVal vNewValue As en_FldDataType)
    me_DataType = vNewValue
    If me_DataType = fdtTimeType Then
        Mask = "#0.00"
    End If
    Call set_maxl
End Property
Public Property Get ColName() As String
    ColName = ms_Colname
End Property
Public Property Let ColName(ByVal vNewValue As String)
    If ms_Colname <> "" And UCase(ms_Colname) <> UCase(vNewValue) Then
        If ms_Recalcon <> "" Or ms_Recalcparent <> "" Then
'??? errhndling at design time
'            ErrMsg = "Cannot change Colname of " + ms_colname + "as "
 '           ErrMsg = ErrMsg + IIf(m_recalcon <> "", "Fld Recalculated on change of " + m_recalcon, "")
  '          ErrMsg = ErrMsg + IIf(m_recalcparent <> "", "Triggering Recalc of " + m_recalcparent, "")
   '         MsgBox ErrMsg
            Exit Property
        End If
    End If
    ms_Colname = UCase(vNewValue)
End Property
Public Property Get DataFld() As String
    DataFld = ms_Datafld
End Property
Public Property Let DataFld(ByVal vNewValue As String)
    ms_Datafld = vNewValue
End Property
Public Property Get Heading() As String
    Heading = ms_Heading
End Property
Public Property Let Heading(ByVal vNewValue As String)
    ms_Heading = vNewValue
    mo_Parent.PropertyModified "heading", ms_Colname
End Property
Public Property Get ReCalcOn(Optional pv_Curobj As Object) As String
    ReCalcOn = ms_Recalcon
End Property
Public Property Let ReCalcOn(Optional pv_Curobj As Object, ByVal vNewValue As String)
    Dim ws_CurObjNm As String, wo_ParObj As Object, ws_OldVal As String, wv_rArr As Variant, i As Integer, ws_ReqStr As String, wi_FirBr As Integer, wi_SecBr As Integer, ws_reqPrm As String, wo_ReqObj As Object, wo_reqClsFlx As MWCLS_FLX, wo_reqMED As MWCTL_MED, ws_rStr As String
    On Error GoTo errh
    If pv_Curobj Is Nothing Then
        ms_Recalcon = vNewValue
        Exit Property
    End If
    If ms_Recalcon = vNewValue Then Exit Property
    Set wo_ParObj = pv_Curobj.ParentRef
'??? can u not use just name for objnm
    ws_CurObjNm = pv_Curobj.ObjNm + "(" + ms_Colname + ")"
    If ms_Recalcon <> vNewValue And ms_Recalcon <> "" Then
         Call Rmv_Frm_Parent(ws_CurObjNm, wo_ParObj)
    End If
    ws_OldVal = ms_Recalcon
    ms_Recalcon = vNewValue
    If ms_Recalcon = "" Then Exit Property
    wv_rArr = Split(ms_Recalcon, ",")
    If Not IsArray(wv_rArr) Then Exit Property
    For i = 0 To UBound(wv_rArr)
        ws_ReqStr = wv_rArr(i)
        wi_FirBr = InStr(1, ws_ReqStr, "(")
        wi_SecBr = InStr(1, ws_ReqStr, ")")
        If wi_FirBr <> 0 Then
            ws_reqPrm = Mid(ws_ReqStr, wi_FirBr + 1, wi_SecBr - wi_FirBr - 1)
            Set wo_ReqObj = wo_ParObj.Controls(Mid(ws_ReqStr, 1, wi_FirBr - 1))
        Else
            Set wo_ReqObj = wo_ParObj.adc(ws_ReqStr)
        End If
        If TypeOf wo_ReqObj Is MWCTL_FLX Then
            Set wo_reqClsFlx = wo_ReqObj.ColProp(ws_reqPrm)
            ws_rStr = wo_reqClsFlx.ReCalcParent
            wo_reqClsFlx.ReCalcParent = IIf(ws_rStr = "", "", ws_rStr + ",") + ws_CurObjNm
        Else
            Set wo_reqMED = wo_ReqObj
            ws_rStr = wo_reqMED.ReCalcParent
            wo_reqMED.ReCalcParent = IIf(ws_rStr = "", "", ws_rStr + ",") + ws_CurObjNm
        End If
    Next i
    Exit Property
errh:
'??? error handling
    MsgBox "Cannot find obj" + ws_ReqStr
    ms_Recalcon = ""
End Property
Private Sub Rmv_Frm_Parent(ByVal pv_CurObjNm As String, pv_Parobj As Object)
    On Error GoTo errh
    Dim wv_rArr As Variant, ws_ReqStr As String, wi_FirBr As Integer, wi_SecBr As Integer, wo_ReqObj As Object, wo_reqClsFlx As MWCLS_FLX, wo_reqMED As MWCTL_MED, ws_rStr As String, wi_rPos As Integer, i As Integer, ws_reqPrm As String
    wv_rArr = Split(ms_Recalcon, ",")
    If Not IsArray(wv_rArr) Then Exit Sub
    For i = 0 To UBound(wv_rArr)
        ws_ReqStr = wv_rArr(i)
        wi_FirBr = InStr(1, ws_ReqStr, "(")
        wi_SecBr = InStr(1, ws_ReqStr, ")")
        If wi_FirBr <> 0 Then
            ws_reqPrm = Mid(ws_ReqStr, wi_FirBr + 1, wi_SecBr - wi_FirBr - 1)
            Set wo_ReqObj = pv_Parobj.Controls(Mid(ws_ReqStr, 1, wi_FirBr - 1))
        Else
            Set wo_ReqObj = pv_Parobj.adc(ws_ReqStr)
        End If
        If TypeOf wo_ReqObj Is MWCTL_FLX Then
            Set wo_reqClsFlx = wo_ReqObj.ColProp(ws_reqPrm)
            ws_rStr = "," + wo_reqClsFlx.ReCalcParent + ","
        Else
            Set wo_reqMED = wo_ReqObj
            ws_rStr = "," + wo_reqMED.ReCalcParent + ","
        End If
        wi_rPos = InStr(1, ws_rStr, "," + pv_CurObjNm + ",")
        If wi_rPos <> 0 Then
            ws_rStr = Mid(ws_rStr, 1, wi_rPos - 1) + Mid(ws_rStr, wi_rPos + Len(pv_CurObjNm) + 1)
            If ws_rStr <> "" Then
                If Left(ws_rStr, 1) = "," Then ws_rStr = Mid(ws_rStr, 2)
                If Right(ws_rStr, 1) = "," Then ws_rStr = Left(ws_rStr, Len(ws_rStr) - 1)
            End If
            If TypeOf wo_ReqObj Is MWCTL_FLX Then
                wo_reqClsFlx.ReCalcParent = ws_rStr
            Else
                wo_reqMED.ReCalcParent = ws_rStr
            End If
        End If
    Next i
    Exit Sub
errh:
    Resume Next
End Sub
Public Property Get ReCalcParent() As String
    ReCalcParent = ms_Recalcparent
End Property
Public Property Let ReCalcParent(ByVal vNewValue As String)
    ms_Recalcparent = vNewValue
End Property
Public Property Get MaxLength() As Integer
    MaxLength = mi_Maxlength
End Property
Public Property Let MaxLength(ByVal vNewValue As Integer)
    mi_Maxlength = vNewValue
    Call set_maxl
End Property
Public Property Set Parent(ByVal vNewValue As Object)
    Set mo_Parent = vNewValue
End Property
Public Property Get Width() As Long
    Width = ml_Width
End Property
Public Property Let Width(ByVal vNewValue As Long)
    ml_Width = vNewValue
End Property
Public Property Get ToolTipText() As String
    ToolTipText = ms_ToolTipText
End Property
Public Property Let ToolTipText(ByVal vNewValue As String)
    ms_ToolTipText = vNewValue
End Property
Public Property Get AllowNegatives() As Boolean
    AllowNegatives = mf_AllowNegatives
End Property
Private Sub set_maxl()
    Select Case me_DataType
    Case Is = fdtDateType
        mi_Maxlength = 8
    Case Is = fdtFloat, fdtNumber, fdtTimeType
        mi_Maxlength = Len(ms_Mask)
    End Select
End Sub
Public Property Get inInitMode() As Boolean
    inInitMode = mf_inInitMode
End Property
Public Property Let inInitMode(ByVal vNewValue As Boolean)
    mf_inInitMode = vNewValue
End Property

