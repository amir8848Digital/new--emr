VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "MDOConnection"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
Dim m_adoconn As ADODB.Connection
Dim mo_AutoGen As ADODB.Command
Enum enMDOCommitType
    mdoCommitNone = 0
    mdoCommitRec = 1
End Enum
Public Type MDOErrObj
    Description As String
    Number As Long
End Type
Public Type SQLLogInDets
    m_DSNName As String
    m_SrvrNm As String
    m_DbNm As String
    m_LoginId As String
    m_Pwd As String
    m_AppPwd As String
    m_AppRole As String
    '*** Jenny
    m_TimeOut As Long
    '*** Jenny
    '*** Sachin 3.02 - 06/09/07
    m_RepSrvrNm As String
    m_RepDbNm As String
    '*** Sachin 3.02 - 06/09/07
End Type
Const adoErrCanDrop = -2147217865
Const adoErrSetRole = -2147217900
Dim m_defaultrs As New MDORowSet
Dim m_defaultrsado As New ADODB.Recordset

'*** (Jenny speed) 16/11/06
'Dim wRs_Hlp As New ADODB.Recordset
'*** (Jenny speed) 16/11/06

'*** (Bef speed) bef 16/11/06
'Public Function SetConn(m_LoginDets As SQLLogInDets) As Boolean
'*** (Bef speed) bef 16/11/06

'*** (Jenny speed) 16/11/06
'****** Sachin 3.02 06/09/07 - (additional optional parameter mb_rep for report connection ******
Public Function SetConn(m_LoginDets As SQLLogInDets, Optional ByVal mb_Hlp As Boolean = False, Optional ByVal mb_Rep As Boolean = False) As Boolean
'*** (Jenny speed) 16/11/06
    
    Dim i As Integer
    On Error GoTo errh:
    SetConn = False
    Dim ws_Connstr As String
    
    With m_LoginDets
        If .m_DSNName <> "" Then
            ws_Connstr = .m_DSNName
        Else
            
            '*** (Bef speed) bef 16/11/06
            'ws_Connstr = "driver={SQL Server};" _
                        & "server=" + .m_SrvrNm + ";" _
                        & "uid=" + .m_LoginId + ";" _
                        & "pwd=" + .m_Pwd + ";" _
                        & "database=" + .m_DbNm
            '*** (Bef speed) bef 16/11/06
                        
            '*** (Jenny speed) 16/11/06
            If mb_Hlp = True Then
              '*** (24/11/06)
              ' sqloledb.1
              ws_Connstr = "Provider=SQLOLEDB.1;" _
                          & "Data Source=" + .m_SrvrNm + ";" _
                          & "Initial Catalog=" + .m_DbNm + ";" _
                          & "User Id=" + .m_LoginId + ";" _
                          & "Password=" + .m_Pwd + ";"
              '*** (24/11/06)
            
              '*** (Bef 24/11/06)
              'ws_Connstr = "Provider=sqloledb;" _
                          & "Data Source=" + .m_SrvrNm + ";" _
                          & "Initial Catalog=" + .m_DbNm + ";" _
                          & "User Id=" + .m_LoginId + ";" _
                          & "Password=" + .m_Pwd + ";"
              '*** (Bef 24/11/06)
              
            '****** Sachin 3.02 06/09/07 - Multi Server ******
            ElseIf mb_Rep = True Then
              ws_Connstr = "driver={SQL Server};" _
                          & "server=" + .m_RepSrvrNm + ";" _
                          & "uid=" + .m_LoginId + ";" _
                          & "pwd=" + .m_Pwd + ";" _
                          & "database=" + .m_RepDbNm
            '****** Sachin 3.02 06/09/07 - Multi Server ******
              
            Else
              ws_Connstr = "driver={SQL Server};" _
                          & "server=" + .m_SrvrNm + ";" _
                          & "uid=" + .m_LoginId + ";" _
                          & "pwd=" + .m_Pwd + ";" _
                          & "database=" + .m_DbNm
            End If
            ''wStrCnn = "Provider=sqloledb;" & _
                      "Data Source=mw6;Initial Catalog=Emr201;User Id=sa;Password=; "
            '*** (Jenny speed) 16/11/06
        
        End If
    End With
    Set m_adoconn = New ADODB.Connection
    m_adoconn.ConnectionString = ws_Connstr
    
    '*** (Jenny speed) 24/11/06
    If mb_Hlp = True Then
      m_adoconn.ConnectionTimeout = 20    '*** default is 15 secs
    End If
    '*** (Jenny speed) 24/11/06
    
    '*** Indicates how long to wait while executing a command before terminating the attempt and generating an error.
    'm_adoconn.CommandTimeout = 300   '*** Jenny ***   'Bef 05/04/03
    m_adoconn.CommandTimeout = m_LoginDets.m_TimeOut   '*** Jenny ***
'    m_adoconn.ConnectionString = "driver={SQL Server};server=MW9;database=MwfTry;uid=sa;pwd="
    m_adoconn.CursorLocation = adUseClient
    '''m_adoconn.CursorLocation = adUseServer   '*** (speed)
    m_adoconn.Open
    If (m_LoginDets.m_AppRole <> "") Then
      m_adoconn.Execute ("sp_SetAppRole '" + m_LoginDets.m_AppRole + "' , '" + m_LoginDets.m_AppPwd + "'")
    End If
    If ms_UsrDtFormat = ms_dmyformat Then
      m_adoconn.Execute ("Set DateFormat dmy")
    Else
      m_adoconn.Execute ("Set DateFormat mdy")
    End If
    
    '****** Sachin 3.2.0 - Temporarily Full Partition to have Top Locking(instruction by Sir)
    If ms_IsoLvl = "SR" Then
        m_adoconn.Execute ("Set Transaction Isolation Level Serializable")
        m_adoconn.IsolationLevel = adXactSerializable
    Else
        m_adoconn.Execute ("Set Transaction Isolation Level Read Uncommitted")
        m_adoconn.IsolationLevel = adXactReadUncommitted
    End If
    
    'm_adoconn.Execute ("Set transaction isolation level read uncommitted")
    'm_adoconn.IsolationLevel = adXactReadUncommitted
    
    SetConn = True
    
    
    Exit Function
errh:
    For i = 0 To ErrCount - 1
        If ErrObj(i).Number = adoErrSetRole Then
            Resume Next
            Exit For
        End If
    Next i

End Function
Public Sub BeginTrans(Optional ByVal pv_CommitType As Variant)
    If Not IsMissing(pv_CommitType) And pv_CommitType = mdoCommitNone Then Exit Sub
    m_adoconn.Execute ("Begin Transaction")
End Sub
Public Sub CommitTrans(Optional ByVal pv_CommitType As Variant)
    If Not IsMissing(pv_CommitType) And pv_CommitType = mdoCommitNone Then Exit Sub
    m_adoconn.Execute ("Commit Transaction")
End Sub
Public Sub RollbackTrans(Optional ByVal pv_CommitType As Variant)
    If Not IsMissing(pv_CommitType) And pv_CommitType = mdoCommitNone Then Exit Sub
    m_adoconn.Execute ("Rollback Transaction")
End Sub
Public Function OpenResultset(ByVal wsqlstrg As String) As MDORowSet
    Dim wo_rsmdo As MDORowSet
    Set wo_rsmdo = New MDORowSet
    Dim wrs As New ADODB.Recordset
    
    wrs.Open wsqlstrg, m_adoconn, adOpenKeyset, adLockOptimistic
    
    Set wo_rsmdo.adorecordset = wrs
    Set wrs = Nothing
    Set OpenResultset = wo_rsmdo
    Set wo_rsmdo = Nothing
End Function
Public Function OpenQuery(ByVal wsqlstrg As String) As MDOQuery
    Dim wo_mdoqry As MDOQuery
    Set wo_mdoqry = New MDOQuery
    
    Dim wo_Command As New ADODB.Command
    wo_Command.ActiveConnection = m_adoconn
    wo_Command.CommandText = wsqlstrg
    wo_Command.CommandType = adCmdStoredProc
    wo_Command.Parameters.Refresh
    If wo_Command.Parameters.Count = 0 Then
        MsgBox wsqlstrg
        MsgBox wo_Command.ActiveConnection.ConnectionString
    End If
    
    Set wo_mdoqry.adoCommand = wo_Command
'    Set wo_Command = Nothing
    
    Set OpenQuery = wo_mdoqry
    Set wo_mdoqry = Nothing
End Function
Public Function OpenRes(ByVal wsqlstrg As String) As MDORowSet
    Dim wo_rsmdo As MDORowSet
    Set wo_rsmdo = New MDORowSet
    Dim wrs As New ADODB.Recordset
    wrs.Open wsqlstrg, m_adoconn, adOpenKeyset, adLockReadOnly

    Set wo_rsmdo.adorecordset = wrs
    Set wrs = Nothing
    Set OpenRes = wo_rsmdo
    Set wo_rsmdo = Nothing
End Function
'*** (Bef speed) bef 16/11/06
'*** (Jen 04/02/08)
Public Function OpenResExisting(ByVal wsqlstrg As String) As MDORowSet
'   Dim wo_rsmdo As MDORowSet
'    Set wo_rsmdo = New MDORowSet
'    Dim wrs As New ADODB.Recordset
    m_defaultrsado.Open wsqlstrg, m_adoconn, adOpenKeyset, adLockReadOnly

    Set m_defaultrs.adorecordset = m_defaultrsado
'    Set wrs = Nothing
    Set OpenResExisting = m_defaultrs
 '   Set wo_rsmdo = Nothing
End Function
'*** (Bef speed) bef 16/11/06
'*** (Jen 04/02/08)

'*** (Bef 04/02/08)
''*** (Jenny speed) 16/11/06
'Public Function OpenRes(ByVal wsqlstrg As String, Optional ByVal wb_rep As Boolean) As MDORowSet
'    Dim wo_rsmdo As MDORowSet
'    Set wo_rsmdo = New MDORowSet
'
'    If wb_rep = True Then
'
'      '*** (20/11/06)
'      Dim wrsHlp As New ADODB.Recordset
'      wrsHlp.CursorLocation = adUseClient
'      wrsHlp.Open wsqlstrg, m_adoconn, adOpenStatic, adLockReadOnly
'      ''wrs1.Open wsqlstrg, m_adoconn, adOpenKeyset, adLockReadOnly
'      Set wo_rsmdo.adorecordset = wrsHlp
'      Set wrsHlp = Nothing
'      '*** (20/11/06)
'
'      Set OpenRes = wo_rsmdo
'      Set wo_rsmdo = Nothing
'    Else
'      Dim wrs As New ADODB.Recordset
'      wrs.Open wsqlstrg, m_adoconn, adOpenKeyset, adLockReadOnly
'
'      Set wo_rsmdo.adorecordset = wrs
'      Set wrs = Nothing
'      Set OpenRes = wo_rsmdo
'      Set wo_rsmdo = Nothing
'    End If
'End Function
''*** (Jenny speed) 16/11/06
'*** (Bef 04/02/08)

Public Function RepRes(ByVal pv_ReportObj As CRAXDRT.Report, ByVal wsqlstrg As String) As Long
    Dim wrs As New ADODB.Recordset
    wrs.Open wsqlstrg, m_adoconn, adOpenKeyset, adLockReadOnly
    pv_ReportObj.Database.SetDataSource wrs
    RepRes = wrs.RecordCount
    Set wrs = Nothing
End Function
Public Function GetFldVal(ByVal pv_Sqlstrg As String) As Variant
'    Dim wo_rsmdo As MDORowSet
    Set m_defaultrs = OpenResExisting(pv_Sqlstrg)
    GetFldVal = ""
    If m_defaultrs.RecCount > 0 Then _
        GetFldVal = IIf(IsNull(m_defaultrs.FldValue(0)), "", m_defaultrs.FldValue(0))
    If m_defaultrs.FldType(0) = mdoftInt Or m_defaultrs.FldType(0) = mdoftFloat Then
        GetFldVal = Val(GetFldVal)
    Else
        GetFldVal = GetFldVal
    End If
    m_defaultrs.CloseRes
 '   Set m_defaultrs = Nothing
End Function
Public Function RecSeek(ByVal pv_Sqlstrg As String) As Boolean
    RecSeek = False
    pv_Sqlstrg = "Select 'x' where exists (" + pv_Sqlstrg + ")"
'    Dim wo_rsmdo As MDORowSet
    Set m_defaultrs = OpenResExisting(pv_Sqlstrg)
    If m_defaultrs.RecCount > 0 Then RecSeek = True
    m_defaultrs.CloseRes
'    Set wo_rsmdo = Nothing
End Function

Private Sub Class_Terminate()
    Set mo_AutoGen = Nothing
    Set m_defaultrsado = Nothing
    Set m_defaultrs = Nothing
    '*** (Jenny speed) 20/11/06
    If Not (m_adoconn Is Nothing) Then
      If m_adoconn.State = adStateOpen Then m_adoconn.Close
    End If
    '*** (Jenny speed) 20/11/06
   
    'If m_adoconn.State = adStateOpen Then m_adoconn.Close     '*** (Bef speed)
   
    Set m_adoconn = Nothing
End Sub
Public Property Get ErrCount() As Long
    ErrCount = m_adoconn.Errors.Count
End Property
Public Property Get ErrObj(ByVal ErrNum As Long) As MDOErrObj
    ErrObj.Description = m_adoconn.Errors(ErrNum).Description
    ErrObj.Number = m_adoconn.Errors(ErrNum).Number
End Property

'*** (Bef 04/02/08)
''*** (Jenny speed) 20/11/06
'Public Property Get ConnState() As Long
'    If m_adoconn Is Nothing Then
'      ConnState = 0
'    Else
'      ConnState = m_adoconn.State
'    End If
'End Property
''*** (Jenny speed) 20/11/06
'*** (Bef 04/02/08)

Public Sub ErrClear()
    m_adoconn.Errors.Clear
End Sub
Public Sub Execute(ByVal ExecStr As String)
    m_adoconn.Execute ExecStr
End Sub
Public Function AutoGen(TableName As String, FieldName As String, PreKey As String, Optional ByVal EntNum As Variant) As Long
  Dim wo_param As ADODB.Parameter
  If mo_AutoGen Is Nothing Then
    Set mo_AutoGen = New ADODB.Command
    mo_AutoGen.ActiveConnection = m_adoconn
    mo_AutoGen.CommandText = "sp_autogen"
    mo_AutoGen.CommandType = adCmdStoredProc
    mo_AutoGen.CommandTimeout = 300     '*** Jenny ***
    mo_AutoGen.Parameters.Refresh
  End If
  With mo_AutoGen
    If IsMissing(EntNum) Then EntNum = -1
    .Parameters(1) = TableName
    .Parameters(2) = FieldName
    .Parameters(3) = PreKey
    .Parameters(4) = EntNum
    .Execute
    AutoGen = .Parameters(0)
  End With
End Function

'*** (Jenny 2.13)   26/05/06
Public Function GetStrConcatVal(FieldName As String, TableName As String, JoinAndCondition As String, ByVal Delimiter As String) As String
  Dim wo_Cmd As ADODB.Command
  Dim wo_param As ADODB.Parameter
  
  'If wo_Cmd Is Nothing Then
    Set wo_Cmd = New ADODB.Command
    wo_Cmd.ActiveConnection = m_adoconn
    wo_Cmd.CommandText = "MwSp_ColConcat"
    wo_Cmd.CommandType = adCmdStoredProc
    wo_Cmd.CommandTimeout = m_adoconn.CommandTimeout
    wo_Cmd.Parameters.Refresh
  'End If
  With wo_Cmd
    .Parameters(1) = FieldName
    .Parameters(2) = TableName
    .Parameters(3) = JoinAndCondition
    .Parameters(4) = Delimiter
    '.Parameters(5) = ""    '*** (This line is not required)
    .Execute
    GetStrConcatVal = .Parameters(5)
  End With
  Set wo_Cmd = Nothing
End Function
'*** (Jenny 2.13)   26/05/06

Public Property Get OpenTableRes() As MDORowSet
    Dim wo_rsmdo As MDORowSet
    Set wo_rsmdo = New MDORowSet
    
    Dim wrs As New ADODB.Recordset
    Set wrs = m_adoconn.OpenSchema(adSchemaTables)
    
    Set wo_rsmdo.adorecordset = wrs
    Set wrs = Nothing
    
    Set OpenTableRes = wo_rsmdo
    Set wo_rsmdo = Nothing
End Property
Public Property Get SrvrDate() As Date
    SrvrDate = GetFldVal("select convert(varchar(8),getdate()," + IIf(ms_UsrDtFormat = "dd/MM/yy", "3", "1") + ") ")
End Property
Public Property Get SrvrTime() As Double
    SrvrTime = GetFldVal("select datepart(hh,getdate())+(convert(float,datepart(mi,getdate()))/100)")
End Property
Public Function CreateTmpTable(ByVal mTableName As String, ByVal mSql As String)
    Dim i As Integer
    On Error GoTo errh
    Execute ("Drop Table " + mTableName)
    Execute (mSql)
    Exit Function
errh:
    For i = 0 To ErrCount - 1
        If ErrObj(i).Number = adoErrCanDrop Then
            Resume Next
            Exit For
        End If
    Next i
End Function





