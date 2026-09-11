VERSION 5.00
Begin VB.UserControl MWCTL_CTL 
   ClientHeight    =   3600
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4800
   InvisibleAtRuntime=   -1  'True
   ScaleHeight     =   3600
   ScaleWidth      =   4800
End
Attribute VB_Name = "MWCTL_CTL"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
Public Event HandleError(ByVal pv_ErrMsg As String, ByVal pv_ErrorType As en_ErrorType)
Public Function SetLogInDets(ByVal pv_UsrId As String, ByVal pv_SysId As String, ByVal pv_CoCd As String, Optional ByVal SrvrNm As String = "", Optional ByVal DbNm As String = "") As Boolean
    SetLogInDets = MWLib.SetLogInDets(pv_UsrId, pv_SysId, pv_CoCd, SrvrNm, DbNm) ''3.11.4
    'SetLogInDets = MWLib.SetLogInDets(pv_UsrId, pv_SysId)  '*** Bef 05/04/03
End Function
Public Function SetLogInParameters(ByVal pv_SrvrSQLLogin As String, ByVal pv_SrvrSQLPwd As String, ByVal pv_AppUser As String, ByVal pv_AppPwd As String) As Boolean
    SetLogInParameters = MWLib.SetLogInParameters(pv_SrvrSQLLogin, pv_SrvrSQLPwd, pv_AppUser, pv_AppPwd)
End Function
Public Property Set MDIApp(ByVal vNewValue As Object)
    Set g_MDIApp = vNewValue
End Property

''*** (Jenny speed) 16/11/06
Public Sub SetHlpConnection()
  Call MWLib.SetHlpConn
End Sub
Public Property Get GetHlpConnection() As MwfLib.MDOConnection
  Set GetHlpConnection = MWLib.GetHlpConn
End Property
Public Sub CloseConnections()
  '*** (Bef 04/02/08)
  'Call MWLib.SetConnNothing    'for a single connection   '24/11/06
  'Call MWLib.SetRepConnNothing  'Sachin 3.02 06/09/07 - Report Server Connection
  '*** (Bef 04/02/08)
  Call MWLib.SetHlpConnNothing
End Sub
'*** (Jenny speed) 16/11/06
'3.11.4
Public Property Get GetMwMstConn(ByVal SysId As String)
  Set GetMwMstConn = Nothing
  'MWLib.GetMwMstConn(SysId)
End Property
'*** (Bef 04/02/08)
''****** Sachin 3.02 ******
'Public Sub SetRepConnection()
'  Call MWLib.SetRepConn
'End Sub
'
'Public Property Get GetRepConnection() As MDOConnection
'  Set GetRepConnection = MWLib.GetRepConn
'End Property
''****** Sachin 3.02 ******
'
''*** (Jenny speed) 24/11/06
'Public Sub SetConnection()
'  Call MWLib.SetConn
'End Sub
'Public Property Get GetConnection() As MDOConnection
'  Set GetConnection = MWLib.GetConn
'End Property
''*** (Jenny speed) 24/11/06
'*** (Bef 04/02/08)

Public Sub HandleError(ByVal pv_ErrMsg As String, Optional ByVal pv_ErrorType As en_ErrorType)
    RaiseEvent HandleError(pv_ErrMsg, pv_ErrorType)
End Sub

Public Property Set ActScrn(ByVal vNewValue As Object)
    Set m_ActScrn = vNewValue
End Property

Public Sub FrmActivate(CurrForm As Object)
'    If CurrForm.ADC.ExecGot Then
 '       If Not (CurrForm.ActiveControl Is Nothing) Then
  '          If typeof CurrForm.ActiveControl  is MwCtl_Med Or typeof CurrForm.ActiveControl  is MwCtl_Flx Then
   '             Call CurrForm.ActiveControl.EnterFocus
    '        End If
     '   End If
'    End If
 '   CurrForm.ADC.ExecGot = False
End Sub
Public Sub FrmDeActivate(CurrForm As Object)
  '  If Not (CurrForm.ActiveControl Is Nothing) Then
'        If TypeOf CurrForm.ActiveControl Is MWCTL_MED Or TypeOf CurrForm.ActiveControl Is MWCTL_FLX Then
 '           Call CurrForm.ActiveControl.LeaveFocus
  '      End If
   ' End If
End Sub


