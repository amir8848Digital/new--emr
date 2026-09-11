Attribute VB_Name = "MwcMod"
Public MWLib As MWCLS_GEN
Public g_MDIApp As MDIForm
Public m_ActScrn As Screen
'**** FOR API CALLS
Declare Function GetSystemMenu Lib "user32" (ByVal hwnd As Long, ByVal bRevert As Long) As Long
Declare Function RemoveMenu Lib "user32" (ByVal hMenu As Long, ByVal nPosition As Long, ByVal wFlags As Long) As Long
Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Long) As Long
Public Const MF_BYPOSITION = &H400
Declare Function Beep Lib "kernel32" (ByVal dwFreq As Long, ByVal dwDuration As Long) As Long
Declare Function FlashWindow Lib "user32" (ByVal hwnd As Long, ByVal bInvert As Long) As Long
Public Const ctErrUpdConflict = -2147217864
Public Const ctErrDtFormat = -2147217869
Public Sub Main()
  Set MWLib = New MWCLS_GEN
End Sub
Public Sub DispMsg(ByVal pv_rStr As String, ByVal pv_ErrorType As en_ErrorType)
  g_MDIApp.Controls("CTL").HandleError pv_rStr, pv_ErrorType
End Sub
Public Sub ChkValidFormats(ByVal mName As String, ByVal mDataType As en_FldDataType, ByVal mMask As String, ByVal mAllowNegatives As Boolean, ByVal mText As Variant, ByRef Cancel As Boolean, ByRef ErrMsg As String)
    Cancel = False
    ErrMsg = ""
    If mDataType = fdtFloat Or mDataType = fdtNumber Or mDataType = fdtTimeType Then
        wnegpos = InStr(1, mText, "-")
        If (Not mAllowNegatives) And (wnegpos <> 0) Then
            Cancel = True
            ErrMsg = "Negative Values not permitted in " + mName
            Exit Sub
        End If
        wdeclen = IIf(InStr(1, mMask, ".") = 0, 0, Len(mMask) - InStr(1, mMask, "."))
        wintlen = Len(mMask) - IIf(wdeclen = 0, 0, wdeclen + 1)
        wintent = Mid(mText, IIf(wnegpos <> 0, wnegpos + 1, 1))
        wdecpos = InStr(1, wintent, ".")
        wdecent = ""
        If wdecpos > 0 Then
            wdecent = Mid(wintent, wdecpos + 1)
            wintent = Mid(wintent, 1, wdecpos - 1)
        End If
        If mDataType = fdtFloat Or mDataType = fdtNumber Then
            If (Len(wintent) > wintlen Or Len(wdecent) > wdeclen) Then
                Cancel = True
                ErrMsg = "Value in " + mName + " beyond permitted limits of " + CStr(wintlen) + "." + CStr(wdeclen)
            End If
        End If
        If mDataType = fdtTimeType Then
            wintent = Val(wintent)
            wdecent = Val(mText) - Val(wintent)
            If wintent > 23 Or wdecent > 0.59 Then
                Cancel = True
                ErrMsg = "Value in " + mName + " not a Valid Time "
            End If
        End If
    End If
End Sub
Public Function NumOvrFlwVal(mDataType As en_FldDataType, mMask As String) As String
    Select Case mDataType
    Case Is = fdtFloat, fdtTimeType
        wdecpos = InStr(1, mMask, ".")
        If wdecpos <> 0 Then
            wStr = String((wdecpos - 1), "*") + "." + String(Len(mMask) - wdecpos, "*")
        Else
            wStr = String(Len(mMask), "*")
        End If
    Case Is = fdtNumber
        wStr = String(Len(mMask), "*")
    End Select
    NumOvrFlwVal = wStr
End Function

' ########################  Manoj #### Ver: 2.0.6 #### Date: 16/06/2004  ########################
'Function For Checking Login Session is Working or Not
' Used By MWCTL_ADC & MWCTL_FLX Controls
' In MWCTL_ADC -> arr_cmd_adfnd_Click,arr_cmd_svcndl_Click(Only Report Save & Entry Delete)
' In MWCTL_FLX -> Del_Rec, AddItem
Public Function ChkLogin(ByRef po_ADC As Object) As Boolean
  If po_ADC.ExcelReportMode Then
    ChkLogin = True
  Exit Function
  End If
  
  '*** (Bef 11/08/07)
  'ChkLogin = po_ADC.Connection.RecSeek("Select 1 from UsrLogin Where UlUsrCd='" + _
  '                      UCase$(Trim$(po_ADC.UsrCd)) + "' And UlHostNm = '" + po_ADC.HostNm + "'")
  '*** (Bef 11/08/07)

  '*** (Jenny 11/08/07)
  '****** Sachin 3.5.0 [11/11/08] - UlIdNo Added ******
  ChkLogin = po_ADC.Connection.RecSeek("Select 1 from UsrLogin Where UlUsrCd='" + _
             UCase$(Trim$(po_ADC.UsrCd)) + "' And UlHostNm = '" + po_ADC.HostNm + "' and UlCoCd= '" + po_ADC.CompanyCd + "' And UlIdNo=" + CStr(po_ADC.UlIdNo))
  '*** (Jenny 11/08/07)
  
  If Not ChkLogin Then DispMsg "Your Login Session is Terminated, Please Login Again", etError
End Function
' ###########################################  Manoj  ###########################################

