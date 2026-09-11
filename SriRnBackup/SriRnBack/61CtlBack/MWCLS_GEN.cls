VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "MWCLS_GEN"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
Dim m_procFrm As MWFRM_PROC
Dim m_ErrFrm As MWFRM_PROC
Dim m_LoginDets As SQLLogInDets
Const ct_INI_NotEncryptedServerName = "DJEMRDB"
Const ct_INI_NotEncryptedExpiryDt = "30/09/15"
Const ct_AppUser = "D6CED2D13EDC"           '"UsrApp"  'Geeta  3.11.4
Const ct_AppUserNotEncrypted = "UsrApp"
Const ct_MwAppRole = "CECAE1E03EFE09398F"    '"MwAppRole" 'Geeta  3.11.4
Const ct_MwAppRoleNotEncrypted = "MwAppRole"    '"MwAppRole" 'Geeta  3.11.4
Const ct_AppRole = "C2CDD0C221C003"           '"AppRole" 'Geeta  3.11.4
Const ct_AppRoleNotEncrypted = "AppRole"           '"AppRole" 'Geeta  3.11.4
Public Enum enGENCryptFormula
  GenCryptOld = 0
  GenCryptNew = 1
End Enum
Dim mo_HlpConn As MDOConnection
Dim mo_RepConn As MDOConnection
Dim m_RepLoginDets As SQLLogInDets
Dim m_INI_NotEncrypted As Boolean
Dim m_INISQLLogin As String
Dim m_INISQLPwd As String
Dim m_INIAppRole As String
Dim m_INIAppRolePwd As String
Private Enum enINITErr
  INIFileNotFound = 1
  ININoOfLinesMisMatch = 2
  INIValuesMissing = 3
  ConnectionFailedToMasterServer = 4
  HarddiskcheckFailed = 5
End Enum
Dim m_INITErrMsg(6) As String
Private Sub Class_Initialize()
    m_INITErrMsg(1) = "INI File Not Found"
    m_INITErrMsg(2) = "INI file - no of lines or content mismatch"
    m_INITErrMsg(3) = "INI Content Invalid"
    m_INITErrMsg(4) = "Connection to Master Server failed"
    m_INITErrMsg(5) = "Harddisk error check failed"
End Sub
Public Function DtSetting() As String
  DtSetting = ms_UsrDtFormat
End Function

Public Function EmptyDate() As Date
    EmptyDate = "01/01/80"
End Function
Public Function DataPurged(ByVal SysId As String) As Boolean
'*** In the ini file the first line contains the Server Name, UsrApp Pwd, AppRolePwd, NetWork Card Number, TimeOut
'    Here only the Time Out Option is not Encrypted. The rest all are in encrypted format.
'    The Second line onwards contains the Company code & the corresponding company name.
'    Both the parameters are in encrypted format.
'*** This function returns the Company Code if the ini file contains only one company code
'    Else this function returns Blank
    DataPurged = False
    On Error GoTo errh
    Dim ws_FlNm As String, ws_IniStr As String, wo_FS As Object, wo_File As Object, wo_tsIni As Object, wi_TotLines As Integer
    Dim wv_LineArr() As String, ws_Arr() As String, ws_IniArr() As String
    
    '***** Reading Ini File
    ws_FlNm = App.Path + "\" + SysId + ".Ini"
    If Dir(ws_FlNm) = "" Then err.Raise vbObjectError + 512, "", "Error Initialising the application"
    Set wo_FS = CreateObject("Scripting.FileSystemObject")
    Set wo_File = wo_FS.GetFile(ws_FlNm)
    Set wo_tsIni = wo_File.OpenAsTextStream(1)  '*** 1 - means forReading
    ws_IniStr = wo_tsIni.ReadAll                '*** Reading all the lines in the ini file
    ws_IniArr = Split(ws_IniStr, Chr(13))
    wo_tsIni.Close
    '****** Done

    If UBound(ws_IniArr) < 1 Then GoTo err
    If ws_IniArr(0) = "" Then GoTo err
       ws_Arr = Split(ws_IniArr(0), ",")
    
    If Not IsArray(ws_Arr) Then GoTo err

    '****** Sachin 3.02.0
    If (UBound(ws_Arr)) >= 5 Then
        If DecryptPwd(Trim$(ws_Arr(5)), GenCryptNew) = "PRG" Then
            DataPurged = True
        Else
            DataPurged = False
        End If
    End If

    Set wo_FS = Nothing
    Set wo_File = Nothing
    Set wo_tsIni = Nothing
    
    Exit Function
err:
    err.Raise vbObjectError + 512, "", "Master Server Dets incorrect"
errh:
    DataPurged = False
    MsgBox err.Description
End Function

Public Function SingleCoCd(ByVal SysId As String) As String
'*** In the ini file the first line contains the Server Name, UsrApp Pwd, AppRolePwd, NetWork Card Number, TimeOut
'    Here only the Time Out Option is not Encrypted. The rest all are in encrypted format.
'    The Second line onwards contains the Company code & the corresponding company name.
'    Both the parameters are in encrypted format.
'*** This function returns the Company Code if the ini file contains only one company code
'    Else this function returns Blank

    SingleCoCd = ""
    On Error GoTo errh
    Dim ws_FlNm As String, ws_IniStr As String, wo_FS As Object, wo_File As Object, wo_tsIni As Object, wi_TotLines As Integer, ws_IniArr() As String, wi_Pos As Integer, ws_Arr() As String, ws_SrvrNm As String
    Dim wv_LineArr() As String, ws_CoDet() As String
'***** Reading Ini File
    ws_FlNm = App.Path + "\" + SysId + ".Ini"
    If Dir(ws_FlNm) = "" Then err.Raise vbObjectError + 512, "", "Error Initialising the application"
    Set wo_FS = CreateObject("Scripting.FileSystemObject")
    Set wo_File = wo_FS.GetFile(ws_FlNm)
    Set wo_tsIni = wo_File.OpenAsTextStream(1)  '*** 1 - means forReading
    ws_IniStr = wo_tsIni.ReadAll                '*** Reading all the lines in the ini file
    wi_TotLines = wo_tsIni.Line                 '*** Returns the total number of lines in the file
    ws_IniArr = Split(ws_IniStr, Chr(13))
    wo_tsIni.Close

    If UBound(ws_IniArr) < 1 Then GoTo err
    If ws_IniArr(0) = "" Then GoTo err
    
'***** Validating values in first line of .Ini File
    wi_Pos = InStr(1, ws_IniArr(0), "MwMstSrv=")
    If wi_Pos = 0 Then GoTo err
    ws_IniArr(0) = Mid(ws_IniArr(0), wi_Pos + Len("MwMstSrv="))
    ws_Arr = Split(ws_IniArr(0), ",")
    If Not IsArray(ws_Arr) Then GoTo err
    
' ########################  Manoj #### Ver: 2.0.6 #### Date: 17/06/2004  ########################
' ########################  Manoj #### Ver: 2.0.8 #### Date: 07/10/2004  ########################
    If UBound(ws_Arr) < 4 Then GoTo err
'   If UBound(ws_Arr) < 6 Then GoTo err
    
    ws_SrvrNm = Trim(ws_Arr(0))
    
   ' If UCase(ws_SrvrNm) = ct_INI_NotEncryptedServerName Then m_INI_NotEncrypted = True

'****** Done
    Set wo_FS = Nothing
    Set wo_File = Nothing
    Set wo_tsIni = Nothing
    
    '*** There should be a minimum of two lines in the ini file.
    If wi_TotLines < 2 Then GoTo err
    
    '*** If there is only one Company Code, in the ini file, then return the Company Code
    If wi_TotLines = 2 Then
      '*** Store all the lines in the ini file in the array wv_LineArr
      '*** Also note that the second line onwards will contain the New Line Character also.
      '    So the first character in the 2nd, 3rd, .... lines will be the New Line Character.
      '    So for extracting the Company Code we have to extract form the secong character onwards in each line.
      wv_LineArr = Split(ws_IniStr, Chr(13))
      If Len(wv_LineArr(1)) > 1 Then
        ws_CoDet = Split(Mid(wv_LineArr(1), 2), ",")
      Else
        ws_CoDet = Split(wv_LineArr(1), ",")
      End If
      '*** There should be two parameters in each line from the second
      If Not IsArray(ws_CoDet) Then GoTo err
      If UBound(ws_CoDet) = 1 Then SingleCoCd = DecryptPwd(Trim(ws_CoDet(0)), GenCryptNew)
    End If
    Exit Function
err:
    err.Raise vbObjectError + 512, "", "Master Server Dets incorrect"
errh:
    SingleCoCd = ""
    MsgBox err.Description
End Function
Public Function SetLoginParameters(ByVal SrvrSQLLogin As String, ByVal SrvrSQLPwd As String, ByVal AppUser As String, ByVal AppPwd As String) As Boolean
    m_INISQLLogin = SrvrSQLLogin
    m_INISQLPwd = SrvrSQLPwd
    m_INIAppRole = AppUser
    m_INIAppRolePwd = AppPwd
    SetLoginParameters = True
End Function
Public Function SetLogInDets(ByVal UId As String, ByVal SysId As String, ByVal CoCd As String, Optional ByVal SrvrNm As String = "", Optional ByVal DbNm As String = "") As Boolean
    SetLogInDets = False
    On Error GoTo errh

    Dim wo_MstConn As MDOConnection
    Dim ws_IniArr() As String
    Set wo_MstConn = GetMwMstConnectionAndIniLines(SysId, ws_IniArr)
    
    If wo_MstConn Is Nothing Then _
      RaiseInitErr enINITErr.ConnectionFailedToMasterServer, ""
      
    Dim ws_Str As String, wi_Pos As Integer, ws_CoCdStr As String, ws_CoDet(2) As String
    Dim i As Integer

    ms_CoNm = ""
    Dim ws_Tmp As String
    For i = 1 To UBound(ws_IniArr)
    ' Move on to Company lines of INI file
      
      If Trim(ws_IniArr(i)) = "" Then RaiseInitErr enINITErr.ININoOfLinesMisMatch, "Line empty"

      Dim wCompanyFound As Boolean
      wCompanyFound = False
      ws_Str = Mid(ws_IniArr(i), 1)
      wi_Pos = InStr(1, ws_Str, ",")
      If wi_Pos > 1 Then
        ws_CoDet(0) = Mid(ws_Str, 1, wi_Pos - 1)
        ws_CoDet(1) = Mid(ws_Str, wi_Pos + 1)
        Dim wCompanyNm As String
        If UCase(CoCd) = UCase(DecryptPwd(Trim(ws_CoDet(0)), GenCryptNew)) Then
          wCompanyFound = True
          ws_Tmp = DecryptPwd(Trim(ws_CoDet(1)), GenCryptNew)
          ms_CoNm = ws_Tmp
          'Old way of extracting conm  -          ws_CoNm = Left$(ws_Tmp, Len(ws_Tmp) - 3)
          Exit For
        End If
      Else
        RaiseInitErr enINITErr.INIValuesMissing, "Company Line has no comma. Line Number:" + CStr(i)
      End If
    Next i
        
    If wCompanyFound = False Then RaiseInitErr enINITErr.INIValuesMissing, "Missing line for Company " + UCase(CoCd) + " in INI "

'Geeta 3.11.4
    '****** Sachin 3.02 (06/09/07) - Added MuRepSvrNm,MuRepDbNm ******
'    Set wo_MstRs = wo_MstConn.OpenRes("Select MuSrvrNm,MuDbNm,MuRepSvrNm,MuRepDbNm from MUsrMst where " + _
'                                       "MuSysId='" + SysId + "' and " + _
'                                       "MuUsrId='" + UId + "'" + IIf(SrvrNm <> "", " and MuSrvrNm='" + SrvrNm + "' ", "") + _
'                                       IIf(DbNm <> "", " and MuDbNm='" + DbNm + "' ", ""))
'
'    If wo_MstRs.RecCount = 0 Then
'        Set wo_MstRs = wo_MstConn.OpenRes("Select MuSrvrNm,MuDbNm,MuRepSvrNm,MuRepDbNm from MUsrMst where " + _
'                                       "MuSysId='" + SysId + "' and " + _
'                                       "MuUsrId=''" + IIf(SrvrNm <> "", " and MuSrvrNm='" + SrvrNm + "' ", "") + _
'                                       IIf(DbNm <> "", " and MuDbNm='" + DbNm + "' ", ""))
'    End If
'    '****** Sachin 3.02 (06/09/07) - Added MuRepSvrNm,MuRepDbNm ******
'
'    If wo_MstRs.RecCount = 0 Then err.Raise vbObjectError + 512, "", "Application Record missing in Master Server"
'
    '****** Sachin 3.02 (06/09/07) - Report Server CHK ******
'    mb_ReplStat = IIf((UCase(wo_MstRs!MuSrvrNm) = UCase(wo_MstRs!MuRepSvrNm) And UCase(wo_MstRs!MuDbNm) = UCase(wo_MstRs!MuRepDbNm)) Or _
'                                  (Trim(wo_MstRs!MuRepSvrNm) = "" And Trim(wo_MstRs!MuRepDbNm) = ""), False, True)
'    '****** Sachin 3.02 (06/09/07) - Report Server ******

    With m_LoginDets
        .m_SrvrNm = SrvrNm  'wo_MstRs!MuSrvrNm
        .m_DbNm = DbNm  'wo_MstRs!MuDbNm
        .m_RepSrvrNm = SrvrNm
        .m_RepDbNm = DbNm
        '****** Get LoginId,Pwd for gn Srvr from SysTable
'        Set wo_MstRs = wo_MstConn.OpenRes("Select MsLoginNm,MsPwd from MSys where " + _
 '                                      "MsSrvrNm='" + .m_SrvrNm + "'")
  '      If wo_MstRs.RecCount = 0 Then err.Raise vbObjectError + 512, "", "Server Dets missing in Sys Table"
        
        .m_LoginId = m_INISQLLogin
        'old way .m_LoginId = IIf(m_INI_NotEncrypted, ct_AppUserNotEncrypted, DecryptPwd(wo_MstRs!MsLoginNm, GenCryptNew))
        'old way .m_Pwd = IIf(m_INI_NotEncrypted, "UsrAppPwd", DecryptPwd(wo_MstRs!MsPwd, GenCryptNew))
        .m_Pwd = m_INISQLPwd
        
        '****** Get Application Pwd for gn SysId from SysTable
        ' old way .m_AppRole = IIf(m_INI_NotEncrypted, ct_AppRoleNotEncrypted, DecryptPwd(ct_AppRole, GenCryptNew))
        
        ' old way .m_AppPwd = IIf(m_INI_NotEncrypted, "AppRolePwd", DecryptPwd(wo_MstConn.GetFldVal("Select MsPwd from MSys where MsSrvrNm='" + SysId + "'" + _
                                                " and MsLoginNm='" + .m_DbNm + "' and MsDbSrvrNm='" + .m_SrvrNm + "'"), GenCryptNew))
        .m_AppRole = m_INIAppRole
        .m_AppPwd = m_INIAppRolePwd
        .m_TimeOut = ml_Timeout
    End With
    
    ms_DbSrvrNm = m_LoginDets.m_SrvrNm
    ms_DbNm = m_LoginDets.m_DbNm
    ms_RepDbSrvrNm = m_LoginDets.m_RepSrvrNm
    ms_RepDbNm = m_LoginDets.m_RepDbNm
    Set wo_MstConn = Nothing
    SetLogInDets = True
    Exit Function
errh:
    SetLogInDets = False
    MsgBox err.Description
End Function
Private Sub RaiseInitErr(ErrNumber As enINITErr, ErrMessage As String)
  err.Raise vbObjectError + ErrNumber, "", m_INITErrMsg(ErrNumber) + "." + vbCrLf + ErrMessage + "." + vbCrLf + "INI File:" + ms_INIFlNm + "."
End Sub

'*** Jenny ***
Public Property Get DbSrvrNm() As String
  DbSrvrNm = ms_DbSrvrNm
End Property
Public Property Get DbNm() As String
  DbNm = ms_DbNm
End Property
Public Property Get CoNm() As String
  CoNm = ms_CoNm
End Property

'****** Sachin 3.02 (06/09/07) ******
Public Property Get RepDbSrvrNm() As String
  RepDbSrvrNm = ms_RepDbSrvrNm
End Property

Public Property Get RepDbNm() As String
  RepDbNm = ms_RepDbNm
End Property
'****** Sachin 3.02 (06/09/07) ******

' ########################  Manoj #### Ver: 2.0.6 #### Date: 07/06/2004  ########################
'Public Property Get NetWrkCrdNo() As String
'  NetWrkCrdNo = ms_NetWrkCrd
'End Property
Public Property Get HDDSrlNo() As String
  HDDSrlNo = ms_HDDSrlNo
End Property
' ###########################################  Manoj  ###########################################
'*** Jenny ***

'*** (Jenny 27/06/07)
Public Property Get GetDllPath() As String
  GetDllPath = App.Path
End Property
'*** (Jenny 27/06/07)
Public Function GetDateFormat() As String
    Dim ws_CurSet As String, wf_Res As Boolean, wo_reg As REGTool5.Registry
    Set wo_reg = New REGTool5.Registry
    wf_Res = True
    wf_Res = wo_reg.GetKeyValue(HKEY_CURRENT_USER, "Control Panel\International", "sShortDate", ws_CurSet)
    Set wo_reg = Nothing
    If Not wf_Res Then GoTo err
    If (ws_CurSet <> "dd/MM/yy" And ws_CurSet <> "MM/dd/yy") Then GoTo err
    GetDateFormat = ws_CurSet
    Exit Function
err:
    GetDateFormat = ""
    Exit Function
End Function

Public Function SetDateFormat() As Boolean
    SetDateFormat = False
    Dim ws_CurSet As String, wf_Res As Boolean, wo_reg As REGTool5.Registry
    Set wo_reg = New REGTool5.Registry
    wf_Res = True
    wf_Res = wo_reg.GetKeyValue(HKEY_CURRENT_USER, "Control Panel\International", "sShortDate", ws_CurSet)
    Set wo_reg = Nothing
    If Not wf_Res Then GoTo err
    ms_UsrDtFormat = ws_CurSet
    If (ws_CurSet <> "dd/MM/yy" And ws_CurSet <> "MM/dd/yy") Then GoTo err
    SetDateFormat = True
    Exit Function
err:
    MsgBox "Change DateFormat in RegionalSettings to dd/MM/yy or MM/dd/yy"
End Function
'Public Function SetLogInDets(ByVal UId As String, ByVal SysId As String) As Boolean
'    With m_LoginDets
'        .m_SrvrNm = "MW9"
'        .m_DbNm = "MwfTry"
'        .m_LoginId = "sa"
'        .m_Pwd = ""
'        .m_AppPwd = ""
'    End With
'    SetLogInDets = True
'End Function

'*** (Jen 04/02/08)
'*** (Bef speed) bef 24/11/06
Public Function SetConn() As MDOConnection
'???    If m_LoginDets <> "" Then
        Set SetConn = New MDOConnection
        SetConn.SetConn m_LoginDets
'    End If
End Function
'*** (Bef speed) bef 24/11/06
'*** (Jen 04/02/08)

'*** (Jenny Hlp speed) 04/02/08
Public Function SetRepConn() As MDOConnection
  Set SetRepConn = New MDOConnection
  SetRepConn.SetConn m_LoginDets, , True
End Function
Public Sub SetHlpConn()
  Dim wb_Connected As Boolean
  wb_Connected = False

  Set mo_HlpConn = New MDOConnection
  On Error GoTo errHlpConn

  'wb_Connected = mo_HlpConn.SetConn(m_LoginDets, True)
  wb_Connected = mo_HlpConn.SetConn(m_LoginDets)
  If wb_Connected = False Then GoTo errHlpConn

  Exit Sub

errHlpConn:
  Set mo_HlpConn = Nothing
  
End Sub
Public Property Get GetHlpConn() As MDOConnection
  Set GetHlpConn = mo_HlpConn
End Property
Public Sub SetHlpConnNothing()
  Set mo_HlpConn = Nothing
End Sub
'*** (Jenny Hlp speed) 04/02/08


'*** (Bef 04/02/08)
''*** (Jenny speed) 24/11/06
''*** for a single connection
'Public Sub SetConn()
'  'Dim wb_Connected As Boolean
'  'wb_Connected = False
'
'  Set mo_Conn = New MDOConnection
'  On Error GoTo errConn
'  'wb_Connected = mo_Conn.SetConn(m_LoginDets)
'  'If wb_Connected = False Then GoTo errConn
'
'  mo_Conn.SetConn m_LoginDets
'  Exit Sub
'
'errConn:
'  Set mo_Conn = Nothing
'End Sub
'Public Property Get GetConn() As MDOConnection
'  Set GetConn = mo_Conn
'End Property
'Public Sub SetConnNothing()
'  Set mo_Conn = Nothing
'End Sub
''*** for a single connection
''*** (Jenny speed) 24/11/06
'
''****** Sachin 3.02 (06/09/07) ******
'Public Sub SetRepConn()
'
'  Dim wb_Connected As Boolean
'  wb_Connected = False
'
'  On Error GoTo errRepConn
'  Set mo_RepConn = New MDOConnection
'
'  wb_Connected = mo_RepConn.SetConn(m_LoginDets, , True)
'  If wb_Connected = False Then GoTo errRepConn
'  Exit Sub
'
'errRepConn:
'  Set mo_Conn = Nothing
'End Sub
'
'Public Property Get GetRepConn() As MDOConnection
'  Set GetRepConn = mo_RepConn
'End Property
'
'Public Sub SetRepConnNothing()
'  Set mo_RepConn = Nothing
'End Sub
''****** Sachin 3.02 (06/09/07) ******
'
''*** (Jenny speed) 16/11/06
'Public Sub SetHlpConn()
'  '*** (24/11/06)
'  Dim wb_Connected As Boolean
'  wb_Connected = False
'  '*** (24/11/06)
'
'  Set mo_HlpConn = New MDOConnection
'  On Error GoTo errHlpConn
'  'mo_HlpConn.SetConn m_LoginDets, True   '*** (Bef 24/11/06)
'
'  '*** (24/11/06)
'  wb_Connected = mo_HlpConn.SetConn(m_LoginDets, True)
'  If wb_Connected = False Then GoTo errHlpConn
'  '*** (24/11/06)
'
'  Exit Sub
'
'errHlpConn:
'  Set mo_HlpConn = Nothing
'End Sub
'Public Property Get GetHlpConn() As MDOConnection
'  Set GetHlpConn = mo_HlpConn
'End Property
'Public Sub SetHlpConnNothing()
'  Set mo_HlpConn = Nothing
'End Sub
''*** (Jenny speed) 16/11/06
'*** (Bef 04/02/08)

Public Function NVal(ByVal pv_NewValue As Variant) As Double
    Dim wi_decpos As Integer, ws_intpor As String, ws_decpor As String, ws_reqStr As String
    pv_NewValue = Trim(pv_NewValue)
    wi_decpos = InStr(1, pv_NewValue, ".")
    ws_intpor = pv_NewValue
    ws_decpor = ""
    If wi_decpos > 0 Then
        ws_intpor = Mid(pv_NewValue, 1, wi_decpos - 1)
        ws_decpor = Mid(pv_NewValue, wi_decpos + 1)
    End If
    ws_reqStr = NNVal(ws_intpor) + IIf(NNVal(ws_decpor) <> "", "." + ws_decpor, "")
    NVal = Val(ws_reqStr)
End Function
Private Function NNVal(ByVal pv_NewValue As Variant) As String
    Dim wi_spcPos As Integer, wi_valchars As Integer
    wi_spcPos = InStr(1, Trim(pv_NewValue), " ")
    wi_valchars = IIf(wi_spcPos >= 1, wi_spcPos - 1, Len(pv_NewValue))
    NNVal = Mid(Trim(pv_NewValue), 1, wi_valchars)
End Function
Public Function Div(ByVal Numerator As Double, ByVal Denominator As Double)
    If Denominator = 0 Then
        Div = 0
    Else
        Div = Numerator / Denominator
    End If
End Function
Public Sub GenError(ByVal Message As String, Optional ByVal NoDisp As Variant, Optional ByVal NoBeep As Variant)
    On Error GoTo 0
    Dim i As Integer, t As Long
    If IsMissing(NoDisp) Then NoDisp = False
    If IsMissing(NoBeep) Then NoBeep = False
    If m_ErrFrm Is Nothing Then
        Set m_ErrFrm = New MWFRM_PROC
    End If
    If Message <> "" And NoBeep <> True Then
        For i = 2 To 4
            t = Beep(i * 100, 200)
        Next i
    End If
' junk lines leftover from debugging
    If InStr(1, Message, "modal") <> 0 Then
        i = 3
    End If
    m_ErrFrm.Caption = "Error Occured"
    m_ErrFrm.Label1 = Message
    m_ErrFrm.Height = m_ErrFrm.Label1.Height + 700
    m_ErrFrm.Width = m_ErrFrm.Label1.Width + 90
'    m_ErrFrm.Height = Len(Message) / 50
    If NoDisp <> True Then
'        m_procFrm.DispMode = 1
'        If m_procFrm.Visible = False Then m_procFrm.Show
    End If
'??? reqd to check visible or else runtime error ave che,generror kiya hai and not closd
End Sub

'*** Jenny
Public Sub ViewError(ByVal pv_Frm As Object)
    If m_ErrFrm Is Nothing Then Call GenError("")
    m_ErrFrm.Picture = pv_Frm.Picture
    If Not (TypeOf pv_Frm Is MDIForm) Then
      m_ErrFrm.Label1.ForeColor = pv_Frm.ForeColor
      m_ErrFrm.BackColor = pv_Frm.BackColor   '*** (Jenny 23/04/07)
    Else
      m_ErrFrm.Label1.ForeColor = pv_Frm.BackColor    '&H0&
      'm_ErrFrm.Label1.ForeColor = &H0&
    End If
    m_ErrFrm.Show 1
End Sub
'*** Jenny

'*** Old Code
'Public Sub ViewError1()
'    If m_ErrFrm Is Nothing Then Call GenError("")
'    m_ErrFrm.Show 1
'End Sub
'*** Old Code

Public Sub BeginProcess(ByVal pv_Frm As Object, ByVal wreqmsg As String)
    pv_Frm.MousePointer = vbHourglass
    If Not (TypeOf pv_Frm Is MDIForm) Then pv_Frm.Enabled = False
    If m_procFrm Is Nothing Then
        Set m_procFrm = New MWFRM_PROC
    End If
    m_procFrm.DispMode = 2
    m_procFrm.Label1 = wreqmsg + ".Please Wait..."
    m_procFrm.Height = m_procFrm.Label1.Height + 700
    m_procFrm.Width = m_procFrm.Label1.Width + 90
'??? consider speed (use with)
      
    '*** Jenny
    m_procFrm.Picture = pv_Frm.Picture
    If Not (TypeOf pv_Frm Is MDIForm) Then
      m_procFrm.Label1.ForeColor = pv_Frm.ForeColor
      m_procFrm.BackColor = pv_Frm.BackColor   '*** (Jenny 23/04/07)
    Else
      m_procFrm.Label1.ForeColor = pv_Frm.BackColor   '&H0&
      'm_procFrm.Label1.ForeColor = &H0&
    End If
    '*** Jenny
    m_procFrm.Show
    m_procFrm.Refresh
    Set pv_Frm = Nothing
End Sub
Public Sub RunningProcess(ByVal wreqmsg As String)
    If Not (m_procFrm Is Nothing) Then
        m_procFrm.Label1.Caption = wreqmsg + ".Please Wait..."
    End If
    'm_procFrm.Label1.Refresh
End Sub
Public Sub EndProcess(ByVal pv_Frm As Object)
    If Not (m_procFrm Is Nothing) Then
        m_procFrm.Visible = False
        Unload m_procFrm
    End If
    Set m_procFrm = Nothing
    pv_Frm.MousePointer = vbNormal
    If Not (TypeOf pv_Frm Is MDIForm) Then pv_Frm.Enabled = True
    Set pv_Frm = Nothing
End Sub
'Public Sub SetFormula(ByVal ReqReport As Object, ByVal FormulaName As String, ByVal Expn As String)
''    FormulaName = "@" + Trim(FormulaName)
'    FormulaName = Trim(FormulaName)
'    Dim wForm As CRAXDRT.FormulaFieldDefinition
'    For Each wForm In ReqReport.FormulaFields
'        If UCase(wForm.FormulaFieldName) = UCase(FormulaName) Then
'            wForm.Text = Expn
'            Exit Sub
'        End If
'    Next wForm
'End Sub
Public Function RmvEmbSpace(ByVal ReqStrg As String)
    Dim wsPcPos As Integer
    ReqStrg = Trim(ReqStrg)
    Do While True
        wsPcPos = InStr(1, ReqStrg, " ")
        If wsPcPos = 0 Then Exit Do
        ReqStrg = Mid(ReqStrg, 1, wsPcPos - 1) + Mid(ReqStrg, wsPcPos + 1)
    Loop
    RmvEmbSpace = ReqStrg
End Function
Public Function NulVal(mVal As Variant, mRetVal As Variant)
    If IsNull(mVal) Then
        NulVal = mRetVal
    Else
        NulVal = mVal
    End If
End Function
Public Function SendFocus(ByVal po_Scrn As Object, ByVal po_reqobj As Object)
    Dim X As Boolean
    Dim po_reqScrn As Screen
    Set po_reqScrn = po_Scrn
    If CtlEnAndVis(po_reqobj) Then
        If UCase(TypeName(po_reqScrn.ActiveControl)) = "MWCTL_MED" Or UCase(TypeName(po_reqScrn.ActiveControl)) = "MWCTL_FLX" Then
            X = False
            X = CallByName(po_reqScrn.ActiveControl, "Validate", VbMethod)
            If X = False Then po_reqobj.SetFocus
        End If
    End If
End Function
Public Function CtlEnAndVis(ByVal po_reqobj As Object)
    CtlEnAndVis = True
    Do While True
        If (Not po_reqobj.Enabled) Or (Not po_reqobj.Visible) Then
            CtlEnAndVis = False
            Exit Do
        Else
            If TypeOf po_reqobj.Container Is VB.Form Then Exit Do
            Set po_reqobj = po_reqobj.Container
        End If
    Loop
End Function
Private Sub Class_Terminate()
    If Not m_ErrFrm Is Nothing Then
        Unload m_ErrFrm
        Set m_ErrFrm = Nothing
    End If
    If Not m_procFrm Is Nothing Then
        Unload m_procFrm
        Set m_procFrm = Nothing
    End If
End Sub
Private Function DecryptPwd(ByVal wTxt As String, ByVal ps_Flg As enGENCryptFormula) As String
   Dim wstr As String, wStr1 As String, wStr2 As String
   Dim i As Integer, j As Integer, k As Integer
   Dim wStrArr() As String, wStr1Arr() As String
   wStrArr = Split(wTxt, Chr(13))
   wStr1 = ""
   
   For i = 0 To UBound(wStrArr)
    wStr2 = ""
    wStr1Arr = Split(wStrArr(i), ",")
     For j = 0 To UBound(wStr1Arr)
        k = 1: wstr = ""
    
        Do While k <= Len(wStr1Arr(j))
            If AscW(Mid$(wStr1Arr(j), k, 1)) <> 10 Then wstr = wstr + Mid$(wStr1Arr(j), k, 1)
             k = k + 1
        Loop
'        If m_INI_NotEncrypted Then
 '           wStr2 = wStr2 + wstr
  '      Else
            'wStr2 = wStr2 + IIf(wStr2 <> "", ",", "") + CryptRC4(FromHexDump(wstr), "RonakRoshni")
            wStr2 = wStr2 + IIf(wStr2 <> "", ",", "") + CryptRC4(FromHexDump(wstr), "AhujaKoila")
   '     End If
      Next
      wStr1 = wStr1 + IIf(wStr1 <> "", vbCrLf, "") + wStr2
   Next
   DecryptPwd = wStr1
End Function
Public Function CryptRC4(sText As String, sKey As String) As String  'Geeta  3.11.4
    Dim baS(0 To 255) As Byte
    Dim baK(0 To 255) As Byte
    Dim bytSwap     As Byte
    Dim lI          As Long
    Dim lJ          As Long
    Dim lIdx        As Long

    For lIdx = 0 To 255
        baS(lIdx) = lIdx
        baK(lIdx) = AscW(Mid$(sKey, 1 + (lIdx Mod Len(sKey)), 1))
    Next
    For lI = 0 To 255
        lJ = (lJ + baS(lI) + baK(lI)) Mod 256
        bytSwap = baS(lI)
        baS(lI) = baS(lJ)
        baS(lJ) = bytSwap
    Next
    lI = 0
    lJ = 0
    For lIdx = 1 To Len(sText)
        lI = (lI + 1) Mod 256
        lJ = (lJ + baS(lI)) Mod 256
        bytSwap = baS(lI)
        baS(lI) = baS(lJ)
        baS(lJ) = bytSwap
        CryptRC4 = CryptRC4 & ChrW$((pvCryptXor(baS((CLng(baS(lI)) + baS(lJ)) Mod 256), AscW(Mid$(sText, lIdx, 1)))))
    Next
End Function


Private Function pvCryptXor(ByVal lI As Long, ByVal lJ As Long) As Long  'Geeta  3.11.4
    If lI = lJ Then
        pvCryptXor = lJ
    Else
        pvCryptXor = lI Xor lJ
    End If
End Function

Public Function ToHexDump(sText As String) As String  'Geeta  3.11.4
    Dim lIdx            As Long

    For lIdx = 1 To Len(sText)
        ToHexDump = ToHexDump & Right$("0" & Hex(AscW(Mid(sText, lIdx, 1))), 2)
    Next
End Function

Public Function FromHexDump(sText As String) As String  'Geeta  3.11.4
    Dim lIdx            As Long
    For lIdx = 1 To Len(sText) Step 2
        FromHexDump = FromHexDump & ChrW$(CLng("&H" & Mid(sText, lIdx, 2)))
    Next
End Function
'Private Function GetPwd(ByVal wTxt As String) As String
''''''''''''''''''' numbers for required password ''''''''''''
'  The number should be of type string & not numeric
'  Copy both the functions. You call GetPwd passing the
'  number as string. It will return Password as string
'-------------------------------------------------------------
'  AppRole 170486177.3001
'  AppRolePwd 170486177.750494
'  UsrApp 486302406.51
'  UsrAppPwd 486302406.128553
'  MwAppRole 361110822.695272
'  MwAppRolePwd  361110822.438021385.3693
''''''Both functions GetPwd & GetSrt as submitted by Up, 25.7.01'''''''''''''''
'
'    Dim wPwd As String
'    Dim wPosDot As Integer
'    wPwd = ""
'    Do While True
'        wPosDot = InStr(1, wTxt, ".")
'        If wPosDot = 0 Then
'            wPwd = wPwd + GetStr(wTxt)
'            Exit Do
'        Else
'            wPwd = wPwd + GetStr(Mid(wTxt, 1, wPosDot - 1))
'            wTxt = Mid(wTxt, wPosDot + 1)
'        End If
'    Loop
'    GetPwd = wPwd
'End Function
'
'Private Function GetStr(ByVal wTxt As String) As String
'    Dim wnum As Variant
'    Dim wchr As Integer
'    Dim wtmp As Variant
'    Dim i As Integer
'    Dim j As Integer
'    Dim wmst As String
'
'    wmst = "0123456789" + "ABCDEFGHIJKLMNOPQRSTUVWXYZ" + _
'                          "abcdefghijklmnopqrstuvwxyz" + "-"
'
'    GetStr = ""
'    wnum = CDec(wTxt)
'    i = 1
'    Do While True
'        If wnum < 63 ^ i Then Exit Do
'        i = i + 1
'    Loop
'    For j = i - 1 To 0 Step -1
'        wtmp = CDec(63 ^ j)
'        wchr = Int(wnum / wtmp)
'        GetStr = GetStr + Mid(wmst, wchr + 1, 1)
'        wnum = wnum - CDec(wtmp * wchr)
'    Next j
'End Function

Private Function DecryptPwd_Old(ByVal wTxt As String, ByVal ps_Flg As enGENCryptFormula) As String  'Geeta  3.11.4
''''''This function is used to decrypt the passwords which have been encrypted using the
''''''GenNo function
''''''''''''''''''' numbers for required password ''''''''''''
'  The number should be of type string & not numeric
'  Copy both the functions. You call DecryptPwd passing the
'  number as string. It will return Password as string
'-------------------------------------------------------------
'  AppRole          980769459.4649
'  AppRolePwd       980769459.411338
'  UsrApp           2682133588.52
'  UsrAppPwd        2682133588.46252
'  MwAppRole        2005802932.25247
'  MwAppRolePwd     2005802932.2423738234.5704
''''''Both functions DecryptPwd & GetSrt as submitted by Up, 25.7.01'''''''''''''''
''''''Modified By Jenny on 02/04/02 '''''''''''''''''''''''''''''''''''''''''''''''
''''''Modified By Manoj on 07/06/04 '''''''''''''''''''''''''''''''''''''''''''''''
    
    Dim wPwd As String
    Dim wPosDot As Integer
    wPwd = ""
    Do While True
        wPosDot = InStr(1, wTxt, ".")
        If wPosDot = 0 Then
            wPwd = wPwd + GetStr(wTxt, ps_Flg)
            Exit Do
        Else
            wPwd = wPwd + GetStr(Mid(wTxt, 1, wPosDot - 1), ps_Flg)
            wTxt = Mid(wTxt, wPosDot + 1)
        End If
    Loop
    DecryptPwd_Old = wPwd
End Function
Private Function GetStr(ByVal wTxt As String, ByVal ps_Flg As enGENCryptFormula) As String
    ''''''Modified By Jenny on 02/04/02 '''''''''''''''''''''''''''''''''''''''''''''''
    Dim wnum As String
    Dim wchr As Long
    Dim wtmp As String
    Dim i As Long
    Dim j As Long
    Dim wMst As String
    
    '63 characters
    '*** (Old Code) *** wMst = "0123456789" + "ABCDEFGHIJKLMNOPQRSTUVWXYZ" + _
                          "abcdefghijklmnopqrstuvwxyz" + "-"
    
    '*** 95 characters
    wMst = "0123456789" + "ABCDEFGHIJKLMNOPQRSTUVWXYZ" + _
                          "abcdefghijklmnopqrstuvwxyz" + _
                          "`~!@#$%^&*()-_=+\|[{]};:'" + """" + ",<.>/?" + " "
    
    GetStr = ""
    wnum = CDec(wTxt)

' ########################  Manoj #### Ver: 2.0.6 #### Date: 07/06/2004  ########################
    i = 1
    Do While True
        If ps_Flg = GenCryptOld Then
          If wnum < 95 ^ i Then Exit Do
        ElseIf ps_Flg = GenCryptNew Then
          If wnum < 96 ^ i Then Exit Do
        End If
        i = i + 1
    Loop
    For j = i - 1 To 0 Step -1
        If ps_Flg = GenCryptOld Then
          wtmp = CDec(95 ^ j)
        ElseIf ps_Flg = GenCryptNew Then
          wtmp = CDec(96 ^ j)
        End If
        wchr = Int(wnum / wtmp)
        If ps_Flg = GenCryptOld Then
          GetStr = GetStr + Mid(wMst, wchr + 1, 1)
        ElseIf ps_Flg = GenCryptNew Then
          GetStr = GetStr + Mid(wMst, CStr(wchr), 1)
        End If
        wnum = wnum - CDec(wtmp * wchr)
    Next j
' ###########################################  Manoj  ###########################################
End Function
Private Function GenNo(ByVal wPwd As String, ByVal ps_Flg As enGENCryptFormula) As String
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'****** This function is used to encrypt any password.
'****** This function should never be given to the user
'****** Ideally a separate exe should be made for this function so that if the
       'text or number to be encrypted is typed in a text box, the encrypted number
       'should be displayed in another text box.
       'The text to be encrypted should be sent to Microway and the encrypted number
       'should be sent back to the user so that it can be stored in the ini file or
       'Param table etc. wherever it is required
'******* This function cannot be used to encrypt any character which is a special symbol
'''''' Modified By Jenny on 02/04/02 ''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
           
    Dim wMst As String
    Dim i As Long
    Dim j As Long
    Dim k As Long
    Dim wAns As String
    Dim wVal As Double
    Dim wTxt As String
    
    '*** (Old Code) *** wMst = "0123456789" + "ABCDEFGHIJKLMNOPQRSTUVWXYZ" + _
                          "abcdefghijklmnopqrstuvwxyz" + " "
    
    '*** 10 + 26 + 26 + 32 + 1 = 95 characters
    wMst = "0123456789" + "ABCDEFGHIJKLMNOPQRSTUVWXYZ" + _
                          "abcdefghijklmnopqrstuvwxyz" + _
                          "`~!@#$%^&*()-_=+\|[{]};:'" + """" + ",<.>/?" + " "
        
    i = 1
    wAns = ""
    Do While True
        If i > Len(wPwd) Then Exit Do
        wTxt = Mid(wPwd, i, 5)
        wVal = 0
        j = 0
' ########################  Manoj #### Ver: 2.0.6 #### Date: 07/06/2004  ########################
        For k = Len(wTxt) To 1 Step -1
            If ps_Flg = GenCryptOld Then
              wVal = wVal + 95 ^ j * (InStr(1, wMst, Mid(wTxt, k, 1)) - 1)
            ElseIf ps_Flg = GenCryptNew Then
              wVal = wVal + 96 ^ j * (InStr(1, wMst, Mid(wTxt, k, 1)))
            End If
            j = j + 1
        Next
' ###########################################  Manoj  ###########################################
        If i = 1 Then
            wAns = CStr(wVal)
        Else
            wAns = wAns + "." + CStr(wVal)
        End If
        i = i + 5
    Loop
    GenNo = wAns
End Function

' ########################  Manoj #### Ver: 2.0.6 #### Date: 07/06/2004  ########################
Public Property Get NoOfUsers() As Variant
  NoOfUsers = ml_NoOfUsers
End Property

Public Property Get IsItLaptop() As Boolean
  IsItLaptop = mb_IsItLaptop
End Property
Public Function GetDateFormatForEntry()
  If ms_UsrDtFormat = ms_dmyformat Then
    GetDateFormatForEntry = "dd/MM/yy"
  Else
    GetDateFormatForEntry = "MM/dd/yy"
  End If
End Function


' Function to Access Server Hdd Volume Serali Number
Public Function GetSrvrHddSrlNo(po_MstConnection As MDOConnection, ByRef HDDSrlNo As String, ByRef ErrMsg As String) As Boolean
  On Error GoTo errh
  Dim wMConn As MDOConnection, wMRs As MDORowSet, wsExecSql As String
  Dim w_MstLogin As SQLLogInDets
  
  GetSrvrHddSrlNo = "-3"
  
  Set wMConn = po_MstConnection
   wsExecSql = "IF OBJECT_ID('tempdb..#TmpProc', 'p') IS NOT NULL " + vbCrLf + _
               " DROP Procedure tempdb..#TmpProc "
   wMConn.Execute (wsExecSql)
               
   wsExecSql = " Create Procedure #TmpProc With Encryption  " + vbCrLf + _
               " As " + _
               " Begin " + _
               " Declare @wCnt as Integer;Declare @Temp VarChar(1000);Declare @KeepValues as varchar(50);  " + vbCrLf + _
               " IF OBJECT_ID('tempdb..#tbl1', 'TABLE') IS NOT NULL   " + vbCrLf + _
               " DROP TABLE #tbl1; " + vbCrLf + _
               " CREATE TABLE #tbl1 (Idx INT IDENTITY(1, 1), Txt NVARCHAR(4000)) " + vbCrLf + _
               " INSERT INTO #tbl1 (Txt) EXEC xp_cmdshell 'vol c:';Select @wCnt=Max(Idx) From #Tbl1; Delete From #Tbl1 where Idx<>2;" + vbCrLf + _
               " Update #Tbl1 Set Txt=Replace(Txt, 'Volume Serial Number is','') ;   " + vbCrLf + _
               " INSERT INTO #tbl1 (Txt) EXEC xp_cmdshell 'wmic diskdrive GET caption' ; Delete From #Tbl1 where Idx<>2 and Idx<>@wCnt+2 ;  " + vbCrLf + _
               " Select @Temp=Txt From #Tbl1 where Idx=2;Select @Temp=@Temp+Txt From #Tbl1 where Idx=@wCnt+2 ;   " + vbCrLf + _
               "  Set @KeepValues = '%[^a-z,0-9,.,-]%'  While PatIndex(@KeepValues, @Temp) > 0 Set @Temp = Stuff(@Temp, PatIndex(@KeepValues, @Temp), 1, '');Select @Temp as Ret;  " + vbCrLf + _
               " End ;"
   wMConn.Execute (wsExecSql)
              
  Set wMRs = wMConn.OpenRes(" Exec #TmpProc")

  If wMRs.RecCount > 0 Then
    HDDSrlNo = wMRs.FldValue("Ret")
  Else
    HDDSrlNo = ""
  End If
  GetSrvrHddSrlNo = True
CloseAll:
  If Not (wMRs Is Nothing) Then Set wMRs = Nothing
  If Not (wMConn Is Nothing) Then Set wMConn = Nothing
  Exit Function
errh:
  ErrMsg = err.Description
  err.Clear
  GetSrvrHddSrlNo = False
  GoTo CloseAll
End Function
Public Property Get ReplRng() As Boolean
  ReplRng = mb_ReplStat
End Property
Public Function GetMwMstConnection(ByVal SysId As String) As MDOConnection
' The idea is to provide a public function with simply minimum no of parameters but calling overloaded function below with right parameters
' We dont have a need here for using the INI values but just using the common function to avoid repeating code
  Dim ws_Arr() As String
  Set GetMwMstConnection = GetMwMstConnectionAndIniLines(SysId, ws_Arr)
End Function
Private Function GetMwMstConnectionAndIniLines(ByVal SysId As String, ByRef IniFileArr() As String) As MDOConnection
    Dim w_MstLogin As SQLLogInDets
    Dim ws_ExpiryDt As String
    
    Dim ws_IniStr As String, ws_SrvrNm As String, wo_FS As Object, wo_File As Object, _
    wo_tsIni As Object, wo_MstConn As MDOConnection, wo_MstRs As MDORowSet, ws_Arr() As String, ws_LoginId As String, ws_Pwd As String, _
    ws_AppPwd As String, ws_IniArr() As String, ws_Str As String, _
    i As Integer, wctr As Integer, wi_Pos As Integer
    
    Set wo_MstConn = New MDOConnection
    Dim ws_Tmp As String, Ws_TmpArr() As String
    ' For China Customers - this flag is used to set value to True or False to identify if INI file is Encrypted or Decrypted
    m_INI_NotEncrypted = False
    ml_NoOfUsers = -1
    
    ' Find and Open INI file
    ms_INIFlNm = App.Path + "\" + SysId + ".Ini"

    If Dir(ms_INIFlNm) = "" Then RaiseInitErr enINITErr.INIFileNotFound, ""
    
    Set wo_FS = CreateObject("Scripting.FileSystemObject")
    Set wo_File = wo_FS.GetFile(ms_INIFlNm)
    Set wo_tsIni = wo_File.OpenAsTextStream(1)
    ws_IniStr = wo_tsIni.ReadAll
    Set wo_FS = Nothing
    Set wo_File = Nothing
    Set wo_tsIni = Nothing
    
    ' Break INI file values
    ws_IniArr = Split(ws_IniStr, Chr(13))
    IniFileArr = ws_IniArr
    If UBound(ws_IniArr) < 1 Then RaiseInitErr enINITErr.ININoOfLinesMisMatch, "No of lines in INI file less than one."
    If ws_IniArr(0) = "" Then RaiseInitErr enINITErr.ININoOfLinesMisMatch, "First INI line empty"
    
    ' Break first line of .Ini File
    wi_Pos = InStr(1, ws_IniArr(0), "MwMstSrv=")
    If wi_Pos = 0 Then RaiseInitErr enINITErr.INIValuesMissing, "Word MwMstSrv= missing in first line of INI File"
    ws_IniArr(0) = Mid(ws_IniArr(0), wi_Pos + Len("MwMstSrv="))
    ws_Arr = Split(ws_IniArr(0), ",")
    
    ' Check its an array and no of values reqd are there in INI file
    If UBound(ws_Arr) = 0 Or (Not IsArray(ws_Arr)) Then RaiseInitErr enINITErr.ININoOfLinesMisMatch, "First INI line missing comma"
    If UBound(ws_Arr) < 4 Then RaiseInitErr enINITErr.ININoOfLinesMisMatch, "Fewer than 5 comma separated values in first line"
    
    ' Extract Values from Line 1
    ws_SrvrNm = Trim(ws_Arr(0))
    If ws_SrvrNm = "" Then RaiseInitErr enINITErr.INIValuesMissing, "Master Server Name is blank"
    
    ' Based on Server name, for Servers specifially marked (for China Customers) set flag that indicates the INI file is not encrypted
'    If UCase(ws_SrvrNm) = ct_INI_NotEncryptedServerName Then _
 '       m_INI_NotEncrypted = True
    ' Validate Timeout and No Of Users are numeric
        
    ' Assign values to right variables
    ws_LoginId = m_INISQLLogin 'DecryptPwd(IIf(m_INI_NotEncrypted, ct_AppUserNotEncrypted, ct_AppUser), GenCryptNew)
    ws_Pwd = m_INISQLPwd ' DecryptPwd(Trim(ws_Arr(1)), GenCryptNew)
    ws_AppPwd = m_INIAppRole ' DecryptPwd(Trim(ws_Arr(2)), GenCryptNew)
    ms_HDDSrlNo = DecryptPwd(Trim$(ws_Arr(1)), GenCryptNew)
    
    ws_Tmp = DecryptPwd(Trim$(ws_Arr(2)), GenCryptNew)
    If Not IsNumeric(ws_Tmp) Then
        RaiseInitErr enINITErr.INIValuesMissing, "No of Users given is a non zero value;Value:" + ws_Tmp
    Else
        ml_NoOfUsers = CInt(ws_Tmp)
    End If
    
    ' For Non China customers take date as in INI for rest take hardcoded date
   ' If UCase(ws_SrvrNm) = ct_INI_NotEncryptedServerName Then
'        ms_ExpiryDate = ct_INI_NotEncryptedExpiryDt
   ' Else
        ms_ExpiryDate = DecryptPwd(Trim$(ws_Arr(3)), GenCryptNew)
   ' End If
    ms_ExpiryDate = GetExpiryDateConv(ms_ExpiryDate)
    
    If Not IsDate(ms_ExpiryDate) Then
      RaiseInitErr enINITErr.INIValuesMissing, " Expiry Date given is a non date value;Value:" + ms_ExpiryDate
    End If
    
    mb_IsItLaptop = IIf(UCase(DecryptPwd(Trim$(ws_Arr(4)), GenCryptNew)) = "Y", True, False)
    
    ' Time out is an optional setting in INI file, so setting it to zero first then reading what is INI file
    ml_Timeout = 0
    If UBound(ws_Arr) = 5 Then
      ws_Tmp = DecryptPwd(Trim$(ws_Arr(5)), GenCryptNew)
      If IsNumeric(ws_Tmp) Then
        ml_Timeout = CLng(ws_Tmp)
      Else
        RaiseInitErr enINITErr.INIValuesMissing, " Timeout given is a non zero value;Value:" + ws_Tmp
      End If
    End If
    
    If ms_HDDSrlNo = "" Then RaiseInitErr enINITErr.INIValuesMissing, "HardDisk Serial Number blank"
    If ml_NoOfUsers < 0 Then RaiseInitErr enINITErr.INIValuesMissing, "No Of Users given is less than zero value;Value:" & ml_NoOfUsers
    
    '****** Connect to MasterDatabase with details in .Ini file
    With w_MstLogin
        .m_SrvrNm = ws_SrvrNm
        .m_DbNm = "MWMst"
        .m_LoginId = ws_LoginId
        .m_Pwd = ws_Pwd
        ' Old way
        '.m_AppPwd = ws_AppPwd
        '.m_AppRole = DecryptPwd(IIf(m_INI_NotEncrypted, ct_MwAppRoleNotEncrypted, ct_MwAppRole), GenCryptNew)
        .m_AppRole = ""
        .m_AppPwd = ""
        .m_TimeOut = 60
    End With

    ms_MstSrvrNm = w_MstLogin.m_SrvrNm
    
    If Not wo_MstConn.SetConn(w_MstLogin) Then RaiseInitErr enINITErr.ConnectionFailedToMasterServer, "Server:" + ws_SrvrNm

    Dim wDt As Date
    wDt = CDate(ms_ExpiryDate)
    If (wo_MstConn.SrvrDate > wDt) Then RaiseInitErr enINITErr.INIValuesMissing, "CALL 98204 90905"
    
        
    
    '*** (Bef 2.13) Harddisk Serial Number Check Temporarily commented since the code is unpredictable at some places (especially on SQL Server 2005). Please verify this check thoroughly
    If HDDSrlNo <> "56789" Then
      Dim wErrMsg As String
      If GetSrvrHddSrlNo(wo_MstConn, ws_Tmp, wErrMsg) Then
        Select Case ws_Tmp
          Case Is = "-1", "-2", "-3"
            RaiseInitErr enINITErr.HarddiskcheckFailed, "Harddisk check failed. Returned HdSerialNo= " + ws_Tmp
          Case Else
            If Mid(ws_Tmp, 1, 9) <> Mid(HDDSrlNo, 1, 9) Then
              RaiseInitErr enINITErr.HarddiskcheckFailed, "Harddisk check failed. Server HdSerialNo= " + ws_Tmp + ". Ini File HdSerialNo: " + HDDSrlNo + " The first 9 characters dont match"
            End If
        End Select
      Else
        RaiseInitErr enINITErr.HarddiskcheckFailed, "Error trying to read Server Harddisk SerialNo. Error Message: " + wErrMsg
      End If
    End If

    Set GetMwMstConnectionAndIniLines = wo_MstConn
    Set wo_MstRs = Nothing
    Exit Function
End Function
Private Function GetExpiryDateConv(ms_ExpiryDate) As String
    If ms_UsrDtFormat = ms_mdyformat Then
      If Len(ms_ExpiryDate) = 8 Then
        GetExpiryDateConv = Mid(ms_ExpiryDate, 4, 2) + "/" + Mid(ms_ExpiryDate, 1, 2) + "/" + Mid(ms_ExpiryDate, 7, 4)
      Else
        GetExpiryDateConv = ms_ExpiryDate
      End If
    Else
        GetExpiryDateConv = ms_ExpiryDate
    End If
End Function

Public Function OldGetMwMstConn(ByVal SysId As String) As MDOConnection
  On Error GoTo errh
  Dim wMConn As MDOConnection, wMRs As MDORowSet, wsExecSql As String
  Dim w_MstLogin As SQLLogInDets
  Set wMConn = New MDOConnection
  On Error GoTo errh
    Dim ws_FlNm As String, ws_IniStr As String, ws_SrvrNm As String, wi_Pos As Integer, wo_FS As Object, wo_File As Object, wo_tsIni As Object, wo_MstConn As MDOConnection, wo_MstRs As MDORowSet, ws_Arr() As String, ws_LoginId As String, ws_Pwd As String, ws_AppPwd As String, ws_CoCdStr As String, ws_IniArr() As String, ws_CoDet(2) As String, ws_Str As String, i As Integer, wctr As Integer
    Dim ws_Tmp As String, Ws_TmpArr() As String
'***** Reading Ini File
    ws_FlNm = App.Path + "\" + SysId + ".Ini"
    If Dir(ws_FlNm) = "" Then err.Raise vbObjectError + 512, "", "Error Initialising the application"
    Set wo_FS = CreateObject("Scripting.FileSystemObject")
    Set wo_File = wo_FS.GetFile(ws_FlNm)
    Set wo_tsIni = wo_File.OpenAsTextStream(1) ' 1 - means forReading
    ws_IniStr = wo_tsIni.ReadAll
    ws_IniArr = Split(ws_IniStr, Chr(13))
    wo_tsIni.Close

    If UBound(ws_IniArr) < 1 Then GoTo errh
    If ws_IniArr(0) = "" Then GoTo errh
    
    '***** Validating values in first line of .Ini File
    wi_Pos = InStr(1, ws_IniArr(0), "MwMstSrv=")
    If wi_Pos = 0 Then GoTo errh
    ws_IniArr(0) = Mid(ws_IniArr(0), wi_Pos + Len("MwMstSrv="))
    ws_Arr = Split(ws_IniArr(0), ",")
    If Not IsArray(ws_Arr) Then GoTo errh
    
    If UBound(ws_Arr) < 4 Then GoTo errh
    
    ws_SrvrNm = Trim(ws_Arr(0))

  With w_MstLogin
    .m_SrvrNm = ws_SrvrNm
    .m_DbNm = "MwMst"
    .m_LoginId = m_INISQLLogin
    .m_Pwd = m_INISQLPwd
        .m_AppRole = "MwAppRole"
        .m_AppPwd = "MwAppRolePwd"
    .m_TimeOut = 0
  End With
  Dim wRet As Boolean
  wRet = wMConn.SetConn(w_MstLogin)
  If wRet = False Then _
      GoTo errh

  Set OldGetMwMstConn = wMConn
  Exit Function
errh:
      err.Raise vbObjectError + 512, "", "Unable to Connect to Master Server. Server: " + ws_SrvrNm + ". Ini File: " + ws_FlNm
End Function
