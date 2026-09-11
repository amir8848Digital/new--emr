Attribute VB_Name = "MwlMod"
Declare Function Beep Lib "kernel32" (ByVal dwFreq As Long, ByVal dwDuration As Long) As Long

'*** Jenny ***
Public ms_DbSrvrNm As String
Public ms_DbNm As String
Public ms_CoNm As String
' ########################  Manoj #### Ver: 2.0.8 #### Date: 15/11/2004  ########################
Public ms_MstSrvrNm As String
' ###########################################  Manoj  ###########################################
'Public ms_IniArr() As String
'Public ms_NetWrkCrd As String
'*** Jenny ***

'****** Sachin 3.02 06/09/07 ******
Public mb_ReplStat As Boolean
Public ms_RepDbSrvrNm As String
Public ms_RepDbNm As String
Public ms_IsoLvl As String
'****** Sachin 3.02 06/09/07 ******

' ########################  Manoj #### Ver: 2.0.6 #### Date: 07/06/2004  ########################
Public ms_HDDSrlNo As String
Public ml_NoOfUsers As Long
Public ml_Timeout As Long
Public mi_LoginWaitTime As Integer
Public mi_LoginRetryCnt As Integer
' ###########################################  Manoj  ###########################################
Public mb_IsItLaptop As Boolean
Public ms_ExpiryDate As String
Public ml_DBPurged As Boolean
Public ms_INIFlNm As String
Public ms_UsrDtFormat As String

Public Const ms_dmyformat = "dd/MM/yy"
Public Const ms_mdyformat = "MM/dd/yy"


