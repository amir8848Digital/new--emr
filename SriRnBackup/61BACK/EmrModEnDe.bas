Attribute VB_Name = "EmrModEnDe"
' ****** Sachin [31-08-06] This Module has been added to Generate New Ini Date From Maintenance ******
Public Enum enGENCryptFormula
  GenCryptOld = 0
  GenCryptNew = 1
End Enum

Const msMstStr = "0123456789" + "ABCDEFGHIJKLMNOPQRSTUVWXYZ" + _
        "abcdefghijklmnopqrstuvwxyz" + _
        "`~!@#$%^&*()-_=+\|[{]};:'" + """" + ",<.>/?" + " "

Public Function Decrypt(ByVal ps_Txt As String, ByVal ps_Flg As enGENCryptFormula) As String
  Dim wPwd As String
  Dim wPosDot As Integer
  wPwd = ""
  Do While True
    wPosDot = InStr(1, ps_Txt, ".")
    If wPosDot = 0 Then
      wPwd = wPwd + GetStr(ps_Txt, ps_Flg)
      Exit Do
    Else
      wPwd = wPwd + GetStr(Mid(ps_Txt, 1, wPosDot - 1), ps_Flg)
      ps_Txt = Mid(ps_Txt, wPosDot + 1)
    End If
  Loop
  Decrypt = wPwd
End Function

Private Function GetStr(ByVal ps_Txt As String, ByVal ps_Flg As enGENCryptFormula) As String
  Dim wnum As String
  Dim wchr As Long
  Dim wtmp As String
  Dim i As Long
  Dim j As Long
  Dim wMst As String

  On Error GoTo ErrHndl
  '*** 95 characters
'  wMst = "0123456789" + "ABCDEFGHIJKLMNOPQRSTUVWXYZ" + _
'  "abcdefghijklmnopqrstuvwxyz" + _
'  "`~!@#$%^&*()-_=+\|[{]};:'" + """" + ",<.>/?" + " "
  wMst = msMstStr
  'wMst = msMstOldStr
  GetStr = ""
  If IsNumeric(ps_Txt) = False Then GoTo ErrHndl
  If CDec(ps_Txt) = 0 And ps_Flg = False Then
    GetStr = ""
    Exit Function
  End If
  
  wnum = CDec(ps_Txt)
  i = 1
  Do While True
    If ps_Flg = GenCryptOld Then
      If wnum < (Len(wMst) ^ i) Then Exit Do
    Else
      If wnum < ((Len(wMst) + 1) ^ i) Then Exit Do
    End If
    i = i + 1
  Loop
  
  For j = i - 1 To 0 Step -1
    If ps_Flg = GenCryptOld Then
      wtmp = CDec(Len(wMst) ^ j)
      wchr = Int(wnum / wtmp)
      GetStr = GetStr + Mid(wMst, wchr + 1, 1)
    Else
      wtmp = CDec((Len(wMst) + 1) ^ j)
      wchr = Int(wnum / wtmp)
      GetStr = GetStr + Mid(wMst, CStr(wchr), 1)
    End If
    wnum = wnum - CDec(wtmp * wchr)
  Next j
  Exit Function
ErrHndl:
  GetStr = "***"
End Function

Public Function Encrypt(ByVal ps_Txt As String, ByVal ps_Flg As enGENCryptFormula) As String
  Dim wMst As String
  Dim i As Long
  Dim j As Long
  Dim k As Long
  Dim wAns As String
  Dim wVal As Double
  Dim wTxt As String

'  '*** 10 + 26 + 26 + 32 + 1 = 95 characters
'  wMst = "0123456789" + "ABCDEFGHIJKLMNOPQRSTUVWXYZ" + _
'        "abcdefghijklmnopqrstuvwxyz" + _
'        "`~!@#$%^&*()-_=+\|[{]};:'" + """" + ",<.>/?" + " "
  wMst = msMstStr
  'wMst = msMstOldStr
  i = 1
  wAns = ""
  Do While True
    If i > Len(ps_Txt) Then Exit Do
    wTxt = Mid(ps_Txt, i, 5)
    wVal = 0
    j = 0
    For k = Len(wTxt) To 1 Step -1
      If ps_Flg = GenCryptOld Then
        wVal = wVal + 95 ^ j * (InStr(1, wMst, Mid(wTxt, k, 1)) - 1)
      Else
        wVal = wVal + 96 ^ j * (InStr(1, wMst, Mid(wTxt, k, 1)))
      End If
      j = j + 1
    Next
    If i = 1 Then
      wAns = CStr(wVal)
    Else
      wAns = wAns + "." + CStr(wVal)
    End If
    i = i + 5
  Loop
  Encrypt = wAns
End Function





