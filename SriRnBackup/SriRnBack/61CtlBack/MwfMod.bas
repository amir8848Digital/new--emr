Attribute VB_Name = "MwfMod"
Public MWLib As MwfLib.MWCLS_GEN
Public m_ActScrn As Screen
'**** FOR API CALLS
Declare Function GetSystemMenu Lib "user32" (ByVal hwnd As Long, ByVal bRevert As Long) As Long
Declare Function RemoveMenu Lib "user32" (ByVal hMenu As Long, ByVal nPosition As Long, ByVal wFlags As Long) As Long
Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Long) As Long
Public Const MF_BYPOSITION = &H400
Declare Function Beep Lib "kernel32" (ByVal dwFreq As Long, ByVal dwDuration As Long) As Long
Declare Function FlashWindow Lib "user32" (ByVal hwnd As Long, ByVal bInvert As Long) As Long
Public Sub Main()
  Set MWLib = New MwfLib.MWCLS_GEN
End Sub
Public Sub ChkValidFormats(ByVal mName As String, ByVal mDataType As En_DataType, ByVal mMask As String, ByVal mAllowNegatives As Boolean, ByVal mText As Variant, ByRef Cancel As Boolean, ByRef ErrMsg As String)
    Cancel = False
    ErrMsg = ""
    If mDataType = Float Or mDataType = Number Or mDataType = TimeType Then
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
        If mDataType = Float Or mDataType = Number Then
            If (Len(wintent) > wintlen Or Len(wdecent) > wdeclen) Then
                Cancel = True
                ErrMsg = "Value in " + mName + " beyond permitted limits of " + CStr(wintlen) + "." + CStr(wdeclen)
            End If
        End If
        If mDataType = TimeType Then
            wintent = Val(wintent)
            wdecent = Val(mText) - Val(wintent)
            If wintent > 23 Or wdecent > 0.59 Then
                Cancel = True
                ErrMsg = "Value in " + mName + " not a Valid Time "
            End If
        End If
    End If
End Sub
Public Function NumOvrFlwVal(mDataType As En_DataType, mMask As String) As String
    Select Case mDataType
    Case Is = Float, TimeType
        wdecpos = InStr(1, mMask, ".")
        If wdecpos <> 0 Then
            wStr = String((wdecpos - 1), "*") + "." + String(Len(mMask) - wdecpos, "*")
        Else
            wStr = String(Len(mMask), "*")
        End If
    Case Is = Number
        wStr = String(Len(mMask), "*")
    End Select
    NumOvrFlwVal = wStr
End Function

