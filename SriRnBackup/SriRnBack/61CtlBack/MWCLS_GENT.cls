VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "MWCLS_GEN"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False

Dim m_procFrm As Form
Dim qd_autogen As rdoQuery
Dim m_DSNName As String
Dim m_CurrConn As rdoConnection
Public Function SetConn(Optional ByVal DSNName As Variant) As rdoConnection
    If IsMissing(DSNName) Then DSNName = m_DSNName
    Set env = rdoEngine.rdoCreateEnvironment("", "sa", "")
   env.CursorDriver = rdUseClientBatch
'    env.CursorDriver = rdUseServer
   Set Conn = env.OpenConnection(DSNName, rdDriverNoPrompt, False, "UID=sa;")
    Set qd_autogen = Conn.CreateQuery("", "{?=call sp_autogen(?,?,?,?)}")
    Conn.Execute ("SET DATEFORMAT DMY")
     Set m_CurrConn = Conn
    Set SetConn = Conn
End Function
Public Function SetFoxConn(DSNName As String) As rdoConnection
    rdoEngine.rdoEnvironments(0).CursorDriver = rdUseClientBatch
'    rdoEngine.rdoEnvironments(0).CursorDriver = rdUseServer
    Set Conn = rdoEngine.rdoEnvironments(0).OpenConnection(DSNName, rdDriverNoPrompt, False, "UID=sa;")
    Set SetFoxConn = Conn
End Function
Public Function ResultSets() As rdoResultsets
    Set ResultSets = m_CurrConn.rdoResultsets
End Function
Public Function RollBackTrans()
    m_CurrConn.RollBackTrans
End Function
Public Function BeginTrans()
    m_CurrConn.BeginTrans
    Set wtmp = OpenRes("select * from dum")
End Function
Public Function CommitTrans()
    m_CurrConn.CommitTrans
End Function
Public Function SaveTrans()
    m_CurrConn.Execute ("SAVE TRANSACTION SV")
End Function
Public Function AutoGen(TableName As String, Fieldname As String, PreKey As String, Optional ByVal EntNum As Variant) As Long
    If IsMissing(EntNum) Then EntNum = -1
    qd_autogen.rdoParameters(1) = TableName
    qd_autogen.rdoParameters(2) = Fieldname
    qd_autogen.rdoParameters(3) = PreKey
    qd_autogen.rdoParameters(4) = EntNum
    qd_autogen.Execute
    AutoGen = qd_autogen.rdoParameters(0)
End Function
Public Function OpenResultset(ByVal wsqlstrg As String) As rdoResultset
    Set OpenResultset = m_CurrConn.OpenResultset(wsqlstrg, rdOpenKeyset, rdConcurValues, rdExecDirect)
'    Set OpenResultset = m_CurrConn.OpenResultset(wsqlstrg, rdOpenForwardOnly, rdConcurValues, rdExecDirect)
End Function
Public Function Execute(ByVal wsqlstrg As String)
    m_CurrConn.Execute wsqlstrg
End Function
Public Function OpenRes(ByVal wsqlstrg As String) As rdoResultset
    Set OpenRes = m_CurrConn.OpenResultset(wsqlstrg, rdOpenKeyset, rdConcurReadOnly, rdExecDirect)
End Function
Public Function NVal(New_Value As Variant) As Double
    New_Value = Trim(New_Value)
    wdecpos = InStr(1, New_Value, ".")
    wintpor = New_Value
    wdecpor = ""
    If wdecpos > 0 Then
        wintpor = Mid(New_Value, 1, wdecpos - 1)
        wdecpor = Mid(New_Value, wdecpos + 1)
    End If
    WREQSTR = NNVal(wintpor) + IIf(NNVal(wdecpor) <> "", "." + wdecpor, "")
    NVal = Val(WREQSTR)
End Function
Private Function NNVal(New_Value As Variant) As String
    wsPcpos = InStr(1, Trim(New_Value), " ")
   wvalchars = IIf(wsPcpos >= 1, wsPcpos - 1, Len(New_Value))
    NNVal = Mid(Trim(New_Value), 1, wvalchars)
End Function
Public Function Div(ByVal Numerator As Double, ByVal Denominator As Double)
    If Denominator = 0 Then
        Div = 0
    Else
        Div = Numerator / Denominator
    End If
End Function
Public Function GetDefa(ByVal Sqlstrg As String) As Variant
    Dim rs_tmp As rdoResultset
    Set rs_tmp = MWLib.OpenRes(Sqlstrg)
    wreqval = ""
    If rs_tmp.RowCount > 0 Then wreqval = IIf(IsNull(rs_tmp.rdoColumns(0)), "", rs_tmp.rdoColumns(0))
    If rs_tmp.rdoColumns(0).Type = rdTypeINTEGER Or rs_tmp.rdoColumns(0).Type = rdTypeFLOAT Then
        GetDefa = Val(wreqval)
    Else
        GetDefa = wreqval
    End If
    rs_tmp.Close
    Set rs_tmp = Nothing
End Function
Public Sub DispMsg(ByVal Message As String, Optional NoBeep As Boolean)
    If Message = "" Then Exit Sub
    SUNMDI.Caption = Message
    If Not NoBeep Then
        For i = 2 To 4
            T = Beep(i * 100, 200)
        Next i
    End If
End Sub
Public Sub EraseMsg()
    SUNMDI.Caption = m_OrigCaption
End Sub
Public Sub FrmActivate(CurrForm As Form)
    Set MWHlp = CurrForm.MWCTL_HLP1
    Set m_CurrConn = CurrForm.ADC.Connection
    If CurrForm.ADC.ExecGot Then
        If Not (CurrForm.ActiveControl Is Nothing) Then
            If TypeOf CurrForm.ActiveControl Is MWCTL_MED Or TypeOf CurrForm.ActiveControl Is MWCTL_FLX Then
                Call CurrForm.ActiveControl.EnterFocus
            End If
        End If
    End If
    CurrForm.ADC.ExecGot = False
End Sub
Public Sub FrmDeActivate(CurrForm As Form)
    If Not (CurrForm.ActiveControl Is Nothing) Then
        If TypeOf CurrForm.ActiveControl Is MWCTL_MED Or TypeOf CurrForm.ActiveControl Is MWCTL_FLX Then
            Call CurrForm.ActiveControl.LeaveFocus
        End If
    End If
End Sub
Public Property Let DSNName(ByVal vNewValue As Variant)
    m_DSNName = vNewValue
End Property
Public Property Get CurrConn() As rdoConnection
    Set CurrConn = m_CurrConn
End Property
Public Property Set CurrConn(ByVal vNewValue As rdoConnection)
    Set m_CurrConn = vNewValue
End Property

Public Sub BeginProcess(ByVal wReqMsg As String)
    Screen.MousePointer = vbHourglass
    If m_procFrm Is Nothing Then
        Set m_procFrm = New MWFRM_PROC
    End If
    m_procFrm.Label1.Caption = wReqMsg + ".Please Wait..."
    m_procFrm.Show False, SUNMDI
'    DoEvents
End Sub

Public Sub EndProcess()
    If Not (m_procFrm Is Nothing) Then
        m_procFrm.Visible = False
        Unload m_procFrm
    End If
    Set m_procFrm = Nothing
    Screen.MousePointer = vbNormal
End Sub
Public Sub SetFormula(ByVal ReqReport As CRAXDRT.Report, ByVal FormulaName As String, ByVal Expn As String)
    FormulaName = "@" + Trim(FormulaName)
    Dim wform As CRAXDRT.FormulaFieldDefinition
    For Each wform In ReqReport.FormulaFields
        If UCase(wform.FormulaFieldName) = UCase(FormulaName) Then
            wform.text = Expn
            Exit Sub
        End If
    Next wform
End Sub
Public Sub Update(ByVal Resultset As rdoResultset)
    SaveTrans
    Resultset.Update
End Sub
Public Function RmvEmbSpace(ByVal ReqStrg As String)
    ReqStrg = Trim(ReqStrg)
    Do While True
        wsPcpos = InStr(1, ReqStrg, " ")
        If wsPcpos = 0 Then Exit Do
        ReqStrg = Mid(ReqStrg, 1, wsPcpos - 1) + Mid(ReqStrg, wsPcpos + 1)
    Loop
    RmvEmbSpace = ReqStrg
End Function


