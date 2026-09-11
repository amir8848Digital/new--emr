VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "EmrClsFun"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' ########################  Manoj #### Ver: 2.0.6 #### Date: 02/02/2004  ########################
' # This Class is used to avail some functions in the Evaluationa of Tag Fields                 #
' ###########################################  Manoj  ###########################################

Option Explicit
' ########################  Manoj #### Ver: 2.0.6 #### Date: 02/02/2004  ########################
' # This Enum Data tpe is Used to Define the Rounding Direction                                 #
' # i.e. 0 = No Rounding, 1 = Simple Rounding, 2 = Rounding Up, 3 = Rounding Down               #
' ###########################################  Manoj  ###########################################
Enum RndDirOpt
  RndNo = 0
  RndSmp = 1
  RndUp = 2
  RndDn = 3
End Enum

' ########################  Manoj #### Ver: 2.0.6 #### Date: 02/02/2004  ########################
' # Rounding Function Which Can be Round Numeric Value                                          #
' # This Can be Round to Any No. of Digits                                                      #
' # This Can be Round To Any Value - Result will be Multiple of that Value                      #
' ###########################################  Manoj  ###########################################
Public Function RndOff(ByVal mNo As Double, Optional ByVal mDigits As Integer = 0, Optional ByVal mVal As Integer = 1, Optional ByVal mRndDir As RndDirOpt = 1) As Double
On Error GoTo ErrHndlr
  Dim wFinalNo As Double, wlMul As Long
  wFinalNo = 0
  Dim wNoStr As String
  If mNo <> 0 Then
    wlMul = 10 ^ mDigits
    mNo = mNo * wlMul
    wNoStr = mNo
    wNoStr = Replace(wNoStr, ".5", ".6")
    mNo = wNoStr
    Select Case UCase(mRndDir)
      Case Is = RndNo
        wFinalNo = mNo
      Case Is = RndSmp
        wFinalNo = IIF(mNo / mVal <> mNo \ mVal, Round(mNo / mVal) * mVal, mNo)
      Case Is = RndUp
        wFinalNo = IIF(mNo / mVal <> mNo \ mVal, Int((mNo + mVal) / mVal) * mVal, mNo)
      Case Is = RndDn
        wFinalNo = IIF(mNo / mVal <> mNo \ mVal, Int(mNo / mVal) * mVal, mNo)
    End Select
    wFinalNo = wFinalNo / wlMul
  End If
  RndOff = wFinalNo
ErrHndlr:
End Function

' ########################  Manoj #### Ver: 2.0.6 #### Date: 02/02/2004  ########################
' # Ths Function Formats Data To Given Format (Numeric & Date values)                           #
' ###########################################  Manoj  ###########################################
Public Function StrFmt(ByVal pv_Val As Variant, ByVal pv_Fmt As Variant) As String
  StrFmt = Format$(pv_Val, pv_Fmt)
End Function

' ########################  Manoj #### Ver: 2.0.6 #### Date: 02/02/2004  ########################
' # This Function Returns One Value According to Condition                                      #
' ###########################################  Manoj  ###########################################
Public Function IIF(pb_Cnd As Boolean, pv_TrueVal As Variant, pv_FalseVal As Variant) As Variant
On Error GoTo ErrHndlr
  If pb_Cnd = True Then
    If IsObject(pv_TrueVal) = True Then
      Set IIF = pv_TrueVal
    Else
      IIF = pv_TrueVal
    End If
  Else
    If IsObject(pv_FalseVal) = True Then
      Set IIF = pv_FalseVal
    Else
      IIF = pv_FalseVal
    End If
  End If
ErrHndlr:
End Function

