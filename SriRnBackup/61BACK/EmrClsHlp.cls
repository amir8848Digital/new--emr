VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "EmrClsHlp"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'*** Jay 2.14 ***
' ########################  Manoj #### Ver: 2.11.0 #### Date: 26/05/2005  ########################
'Public Sub PMCd(ByVal ps_PTyp As String, Optional ByVal pv_PValue As Variant, Optional ByVal pv_PValue1 As Variant, Optional pv_Filter As Variant)
'Public Sub PMCd(ByVal ps_PTyp As String, Optional ByVal pv_PValue As Variant, Optional ByVal pv_PValue1 As Variant)
' ###########################################  Manoj  ###########################################
Public Sub PMCd(ByVal ps_PTyp As String, Optional ByVal pv_PValue As Variant, Optional ByVal pv_PValue1 As Variant, _
                Optional pv_Filter As Variant, Optional ByVal pb_ShowValidOnly As Boolean = True)
'*** Jay 2.14 ***
  
  'Dim wPTyp As String
  
' ########################  Manoj #### Ver: 2.0.6 #### Date: 30/03/2004  ########################
  Dim ws_HdPMCd As String, ws_HdPSCd As String, ws_HdPDesc As String
  Dim ws_HdPDesc225 As String, ws_HdPValue As String, ws_HdPNum As String, ws_HdPValue1  As String
  Dim ws_HdPNum1 As String, ws_HdPValue2 As String, ws_HdPValue3 As String, ws_HdPValidYN As String
  Dim ws_ColHeadg As String, ws_ColFlds As String
  Dim ws_Cnd As String
  
  ws_Cnd = ""
  If Not IsMissing(pv_PValue) Then ws_Cnd = IIF(Trim(pv_PValue) <> "", " and PValue in (" + pv_PValue + ") ", " ")
  If Not IsMissing(pv_PValue1) Then ws_Cnd = ws_Cnd + " and PValue1= '" + pv_PValue1 + "' "
' ########################  Manoj #### Ver: 2.11.0 #### Date: 26/05/2005  ########################
  If Not IsMissing(pv_Filter) Then ws_Cnd = IIF(Trim$(ws_Cnd) <> "", ws_Cnd, "") + IIF(Trim$(pv_Filter) <> "", " and " + pv_Filter, "")
' ###########################################  Manoj  ###########################################
  '*** JAY 2.14 ***
  ws_Cnd = IIF(Trim$(ws_Cnd) <> "", ws_Cnd, "") + IIF(pb_ShowValidOnly = True, " and PValidYN <> 'N' ", " ")
  '*** JAY 2.14 ***

  Call GetParamHd("Param", ps_PTyp, ws_HdPMCd, ws_HdPSCd, ws_HdPDesc, ws_HdPDesc225, ws_HdPValue, ws_HdPNum, _
                  ws_HdPValue1, ws_HdPNum1, ws_HdPValue2, ws_HdPValue3, ws_HdPValidYN)
  
  ws_ColHeadg = IIF(ws_HdPMCd = "", "", ws_HdPMCd + "|") + IIF(ws_HdPSCd = "", "", ws_HdPSCd + "|") + _
                IIF(ws_HdPDesc = "", "", ws_HdPDesc + "|") + IIF(ws_HdPDesc225 = "", "", ws_HdPDesc225 + "|") + _
                IIF(ws_HdPValue = "", "", ws_HdPValue + "|") + IIF(ws_HdPNum = "", "", ws_HdPNum + "|") + _
                IIF(ws_HdPValue1 = "", "", ws_HdPValue1 + "|") + IIF(ws_HdPNum1 = "", "", ws_HdPNum1 + "|") + _
                IIF(ws_HdPValue2 = "", "", ws_HdPValue2 + "|") + IIF(ws_HdPValue3 = "", "", ws_HdPValue3 + "|") + _
                IIF(ws_HdPValidYN = "", "", ws_HdPValidYN + "|")
  ws_ColFlds = IIF(ws_HdPMCd = "", "", "PMCd" + "|") + IIF(ws_HdPSCd = "", "", "PSCd" + "|") + _
                IIF(ws_HdPDesc = "", "", "PDesc" + "|") + IIF(ws_HdPDesc225 = "", "", "PDesc225" + "|") + _
                IIF(ws_HdPValue = "", "", "PValue" + "|") + IIF(ws_HdPNum = "", "", "PNum" + "|") + _
                IIF(ws_HdPValue1 = "", "", "PValue1" + "|") + IIF(ws_HdPNum1 = "", "", "PNum1" + "|") + _
                IIF(ws_HdPValue2 = "", "", "PValue2" + "|") + IIF(ws_HdPValue3 = "", "", "PValue3" + "|") + _
                IIF(ws_HdPValidYN = "", "", "PValidYN" + "|")
' ###########################################  Manoj  ###########################################
  If ws_ColHeadg = "" Then
    ws_ColHeadg = "|Code             |Desc                                                  |"
    ws_ColFlds = "|PMCd|PDesc|"
  Else
    ws_ColHeadg = "|" + ws_ColHeadg
    ws_ColFlds = "|" + ws_ColFlds
  End If
  
  With mwhlp
    If ps_PTyp = "DAANAFLD" Then
        .Sqlstrg = " Select * From Param " + _
               " Where PTyp ='" + ps_PTyp + "' " + ws_Cnd + " order by (case when len(ltrim(rtrim(PMCd))) = 1 then '0' else '' end) + PMCd"
               
    Else
        .Sqlstrg = " Select * From Param " + _
               " Where PTyp ='" + ps_PTyp + "' " + ws_Cnd + " order by PMCd " + IIF(UCase(ps_PTyp) = "YY", " desc", "")
    End If
    .ColHeadg = ws_ColHeadg
    .ColFlds = ws_ColFlds
    .RetFldNm = "PMCd"
    .SearchColNum = 1
    If UCase(ps_PTyp) = "YY" Then
      .SortAsc = False
    Else
      .SortAsc = True
    End If
  End With
End Sub

'****** Sachin 3.2.0 - Help for Partition Sr ******
Public Sub PrtnSr()
  With mwhlp
    .Sqlstrg = " Select PNum1, PMCd, PDesc From Param " + _
               " Where PTyp ='PTNTYP' order by PNum1 "
    .ColHeadg = "|Sr.|Prtn |Desc                                                |"
    .ColFlds = "|PNum1|PMCd|PDesc|"
    .RetFldNm = "PNum1"
    .SearchColNum = 1
    .SortAsc = True
  End With
End Sub
'****** Sachin 3.2.0 - Help for Partition Sr ******

'*** jay 2.14 *** (pb_ShowValidOnly ADDED)
'Geeta ***Crm***2.12*********pb_ShowAll
'Used in param where PTyp="RMID"  and PMCD=GRDRMCTG
Public Sub PSCd(ByVal ps_PTyp As String, ByVal ps_PMCd As String, _
                Optional ByVal pb_ShowAll_PMCd As Boolean = False, Optional ByVal pb_ShowValidOnly As Boolean = True)
'Public Sub PSCd(ByVal ps_PTyp As String, ByVal ps_PMCd As String, Optional ByVal pb_ShowAll As Boolean = False)
'*** jay 2.14 ***
  'Dim wPTyp As String
' ########################  Manoj #### Ver: 2.0.6 #### Date: 30/03/2004  ########################
  Dim ws_HdPMCd As String, ws_HdPSCd As String, ws_HdPDesc As String
  Dim ws_HdPDesc225 As String, ws_HdPValue As String, ws_HdPNum As String, ws_HdPValue1  As String
  Dim ws_HdPNum1 As String, ws_HdPValue2 As String, ws_HdPValue3 As String, ws_HdPValidYN As String
  Dim ws_ColHeadg As String, ws_ColFlds As String
  
  Call GetParamHd("Param", ps_PTyp, ws_HdPMCd, ws_HdPSCd, ws_HdPDesc, ws_HdPDesc225, ws_HdPValue, ws_HdPNum, _
                  ws_HdPValue1, ws_HdPNum1, ws_HdPValue2, ws_HdPValue3, ws_HdPValidYN)
  
  ws_ColHeadg = IIF(ws_HdPSCd = "", "", ws_HdPSCd + "|") + _
                IIF(ws_HdPDesc = "", "", ws_HdPDesc + "|") + IIF(ws_HdPDesc225 = "", "", ws_HdPDesc225 + "|") + _
                IIF(ws_HdPValue = "", "", ws_HdPValue + "|") + IIF(ws_HdPNum = "", "", ws_HdPNum + "|") + _
                IIF(ws_HdPValue1 = "", "", ws_HdPValue1 + "|") + IIF(ws_HdPNum1 = "", "", ws_HdPNum1 + "|") + _
                IIF(ws_HdPValue2 = "", "", ws_HdPValue2 + "|") + IIF(ws_HdPValue3 = "", "", ws_HdPValue3 + "|") + _
                IIF(ws_HdPValidYN = "", "", ws_HdPValidYN + "|")
  ws_ColFlds = IIF(ws_HdPSCd = "", "", "PSCd" + "|") + _
               IIF(ws_HdPDesc = "", "", "PDesc" + "|") + IIF(ws_HdPDesc225 = "", "", "PDesc225" + "|") + _
               IIF(ws_HdPValue = "", "", "PValue" + "|") + IIF(ws_HdPNum = "", "", "PNum" + "|") + _
               IIF(ws_HdPValue1 = "", "", "PValue1" + "|") + IIF(ws_HdPNum1 = "", "", "PNum1" + "|") + _
               IIF(ws_HdPValue2 = "", "", "PValue2" + "|") + IIF(ws_HdPValue3 = "", "", "PValue3" + "|") + _
               IIF(ws_HdPValidYN = "", "", "PValidYN" + "|")
' ###########################################  Manoj  ###########################################
  
  If ws_ColHeadg = "" Then
    ws_ColHeadg = "|Code             |Desc                                                |"
    ws_ColFlds = "|PSCd|PDesc|"
  Else
    ws_ColHeadg = "|" + ws_ColHeadg
    ws_ColFlds = "|" + ws_ColFlds
  End If
  
  With mwhlp
     '*** jay 2.14 ***
     .Sqlstrg = " Select * From Param " + _
                " Where PTyp ='" + ps_PTyp + "' " + _
                IIF(pb_ShowAll_PMCd = False, " and  PMCd ='" + ps_PMCd + "' ", "") + _
                IIF(pb_ShowValidOnly = True, " and PValidYN <> 'N' ", " ") + " order by PSCd "
     '.Sqlstrg = " Select * From Param " + _
             " Where PTyp ='" + ps_PTyp + "' " + _
             IIF(pb_ShowAll = False, " and  PMCd ='" + ps_PMCd + "' ", "") + " order by PSCd "
     '*** jay 2.14 ***
    .ColHeadg = ws_ColHeadg
    .ColFlds = ws_ColFlds
    .RetFldNm = "PSCd"
    .SearchColNum = 1
    .SortAsc = True
  End With
End Sub

'*** JAY 2.14 ***(pb_ShowValidOnly ADDED)
Public Sub vPMCd(ByVal ps_cocd As String, ByVal ps_vPTyp As String, Optional ByVal pv_vPValue As Variant, _
                 Optional ByVal pv_vPValue1 As Variant, Optional ByVal pb_ShowValidOnly As Boolean = True, Optional ByVal pv_vPValue3 As Variant, Optional ByVal pv_LocToSortByForWorkerHelp As String = "")
'Public Sub vPMCd(ByVal ps_CoCd As String, ByVal ps_vPTyp As String, Optional ByVal pv_vPValue As Variant, Optional ByVal pv_vPValue1 As Variant)

'*** JAY 2.14 ***
  'Dim wPTyp As String
' ########################  Manoj #### Ver: 2.0.6 #### Date: 30/03/2004  ########################
  Dim ws_HdPMCd As String, ws_HdPSCd As String, ws_HdPDesc As String
  Dim ws_HdPDesc225 As String, ws_HdPValue As String, wHdPNum As String, ws_HdPValue1  As String
  Dim wHdPNum1 As String, ws_HdPValue2 As String, ws_HdPValue3 As String, ws_HdPValue4 As String, ws_HdPValidYN As String
  Dim ws_ColHeadg As String, ws_ColFlds As String
  Dim ws_Head As String, ws_Flds As String
  Dim ws_Cnd As String
  Dim ws_OrdBy As String, ws_Wrk As String
  
  ' ***** Manali 3.6.2 - Following Change Reverted
  'ws_Wrk = "WORK"   ' ****  Manali 3.6.0 - Help for vPTyp='WORK' changed - By PLoc
  
  If Trim(ps_cocd) <> "" Then
      ws_Cnd = ws_Cnd + " and vPCoCd = '" + ps_cocd + "'"
        ws_Head = "": ws_Flds = ""
  ElseIf Trim(ps_cocd) = "" Then
      ws_Head = "Company |": ws_Flds = "vPCoCd|"
  End If
        
  ' ***** Manali 3.6.2 - Following Change Reverted
  ' ****  Manali 3.6.0 - Help for vPTyp='WORK' changed - By PLoc -
  'If Not IsMissing(pv_vPValue) And ps_vPTyp <> ws_Wrk Then ws_Cnd = ws_Cnd + IIF(Trim(pv_vPValue) <> "", " and vPValue in (" + pv_vPValue + ") ", " ")
  ''' Bef 3.6.0
  
  If Not IsMissing(pv_vPValue) Then ws_Cnd = ws_Cnd + IIF(Trim(pv_vPValue) <> "", " and vPValue in (" + pv_vPValue + ") ", " ")
  If Not IsMissing(pv_vPValue1) Then ws_Cnd = ws_Cnd + " and vPValue1= '" + pv_vPValue1 + "' "
  
  '****** Sachin 2.14.0 - [23-11-2006] - [Multi-Company] ******
  If Not IsMissing(pv_vPValue3) Then ws_Cnd = ws_Cnd + " and vPValue3= '" + pv_vPValue3 + "' "
  '****** Sachin 2.14.0 - [23-11-2006] - [Multi-Company] ******
  
  '*** JAY 2.14 ***
  ws_Cnd = ws_Cnd + IIF(pb_ShowValidOnly = True, " and vPValidYn <> 'N' ", " ")
  '*** JAY 2.14 ***
 
  '************Geeta******Emr207 ****New Field vPValue4 in vParam
  Call GetParamHd("vParam", ps_vPTyp, ws_HdPMCd, ws_HdPSCd, ws_HdPDesc, ws_HdPDesc225, ws_HdPValue, wHdPNum, _
                  ws_HdPValue1, wHdPNum1, ws_HdPValue2, ws_HdPValue3, ws_HdPValidYN, ws_HdPValue4)
  
  ws_ColHeadg = ws_Head + IIF(ws_HdPMCd = "", "", ws_HdPMCd + "|") + IIF(ws_HdPSCd = "", "", ws_HdPSCd + "|") + _
                IIF(ws_HdPDesc = "", "", ws_HdPDesc + "|") + IIF(ws_HdPDesc225 = "", "", ws_HdPDesc225 + "|") + _
                IIF(ws_HdPValue = "", "", ws_HdPValue + "|") + IIF(wHdPNum = "", "", wHdPNum + "|") + _
                IIF(ws_HdPValue1 = "", "", ws_HdPValue1 + "|") + IIF(wHdPNum1 = "", "", wHdPNum1 + "|") + _
                IIF(ws_HdPValue2 = "", "", ws_HdPValue2 + "|") + IIF(ws_HdPValue3 = "", "", ws_HdPValue3 + "|") + _
                IIF(ws_HdPValidYN = "", "", ws_HdPValidYN + "|") + IIF(ws_HdPValue4 = "", "", ws_HdPValue4 + "|")
                
  ws_ColFlds = ws_Flds + IIF(ws_HdPMCd = "", "", "vPMCd" + "|") + IIF(ws_HdPSCd = "", "", "vPSCd" + "|") + _
                IIF(ws_HdPDesc = "", "", "vPDesc" + "|") + IIF(ws_HdPDesc225 = "", "", "vPDesc225" + "|") + _
                IIF(ws_HdPValue = "", "", "vPValue" + "|") + IIF(wHdPNum = "", "", "vPNum" + "|") + _
                IIF(ws_HdPValue1 = "", "", "vPValue1" + "|") + IIF(wHdPNum1 = "", "", "vPNum1" + "|") + _
                IIF(ws_HdPValue2 = "", "", "vPValue2" + "|") + IIF(ws_HdPValue3 = "", "", "vPValue3" + "|") + _
                IIF(ws_HdPValidYN = "", "", "vPValidYN" + "|") + IIF(ws_HdPValue4 = "", "", "vPValue4" + "|")
' ###########################################  Manoj  ###########################################
  
  If ws_ColHeadg = "" Then
    ws_ColHeadg = "|Code             |Desc                                                  |"
    ws_ColFlds = "|vPMCd|vPDesc|"
  Else
    ws_ColHeadg = "|" + ws_ColHeadg
    ws_ColFlds = "|" + ws_ColFlds
  End If
  
  ' ***** Manali - Following Change Reverted
  ' ****  Manali 3.6.0 - - 01/09/09 - Help for vPTyp='WORK' changed - By PLoc
  If ps_vPTyp = "WORK" And pv_LocToSortByForWorkerHelp <> "" Then
    ws_OrdBy = " Order by vPCoCd, Case  vPValue When '" + pv_LocToSortByForWorkerHelp + "' Then 1 When '' Then 3 Else 2 End, vPValue, vPMCd "
  Else
    ws_OrdBy = " Order by vPCoCd, vPMCd "
  End If
  ' ****  Manali 3.6.0 - Help for vPTyp='WORK' changed - By PLoc
  
  With mwhlp
    .Sqlstrg = " Select * From vParam " + _
               " Where vPTyp ='" + ps_vPTyp + "' " + ws_Cnd + ws_OrdBy
    .ColHeadg = ws_ColHeadg
    .ColFlds = ws_ColFlds
    .RetFldNm = "vPMCd"
    '**********Geeta**Emr206 *****15/04/04
    'If Trim(ps_CoCdFr) = "" Then     '*** (Bef 2.14 Next Patch 1)
    '*** (Jen 2.14 Next Patch 1)
    If Trim(ps_cocd) = "" Then
    '*** (Jen 2.14 Next Patch 1)
      .SearchColNum = 2
    Else
      .SearchColNum = 1
    End If
    .SortAsc = True
    '**********Geeta**Emr206 *****15/04/04
  End With
End Sub

' **** Zubin 211 **** '
'''''Public Sub vPSCd(ByVal ps_CoCd As String, ByVal ps_vPTyp As String, ByVal ps_vPMCd As String)
'''''  'Dim wPTyp As String
'''''' ########################  Manoj #### Ver: 2.0.6 #### Date: 30/03/2004  ########################
'''''  Dim ws_HdPMCd As String, ws_HdPSCd As String, ws_HdPDesc As String
'''''  Dim ws_HdPDesc225 As String, ws_HdPValue As String, ws_HdPNum As String, ws_HdPValue1  As String
'''''  Dim ws_HdPNum1 As String, ws_HdPValue2 As String, ws_HdPValue3 As String, ws_HdPValue4 As String, ws_HdPValidYN As String
'''''  Dim ws_ColHeadg As String, ws_ColFlds As String
'''''  Dim ws_Head As String, ws_Flds As String
'''''  Dim ws_Cnd As String
'''''
'''''  If Trim(ps_CoCd) <> "" Then
'''''      ws_Cnd = ws_Cnd + " and vPCoCd = '" + ps_CoCd + "'"
'''''      ws_Head = "": ws_Flds = ""
'''''  ElseIf Trim(ps_CoCd) = "" Then
'''''      ws_Head = "Company |": ws_Flds = "vPCoCd|"
'''''  End If
'''''
'''''  '******************************Geeta**********Emr205****6/05/04**********
'''''    If Trim(ps_vPTyp) = "CHR" Then _
'''''       ws_Cnd = ws_Cnd + " and Upper(substring(vPDesc,LEN (Rtrim(vPDesc))-2,3))<>Upper('XXX') "
'''''  '********************************Geeta************************************
'''''
'''''  Call GetParamHd("vParam", ps_vPTyp, ws_HdPMCd, ws_HdPSCd, ws_HdPDesc, ws_HdPDesc225, ws_HdPValue, ws_HdPNum, _
'''''                  ws_HdPValue1, ws_HdPNum1, ws_HdPValue2, ws_HdPValue3, ws_HdPValidYN, ws_HdPValue4)
'''''
'''''  ws_ColHeadg = ws_Head + IIF(ws_HdPSCd = "", "", ws_HdPSCd + "|") + _
'''''                IIF(ws_HdPDesc = "", "", ws_HdPDesc + "|") + IIF(ws_HdPDesc225 = "", "", ws_HdPDesc225 + "|") + _
'''''                IIF(ws_HdPValue = "", "", ws_HdPValue + "|") + IIF(ws_HdPNum = "", "", ws_HdPNum + "|") + _
'''''                IIF(ws_HdPValue1 = "", "", ws_HdPValue1 + "|") + IIF(ws_HdPNum1 = "", "", ws_HdPNum1 + "|") + _
'''''                IIF(ws_HdPValue2 = "", "", ws_HdPValue2 + "|") + IIF(ws_HdPValue3 = "", "", ws_HdPValue3 + "|") + _
'''''                IIF(ws_HdPValidYN = "", "", ws_HdPValidYN + "|") + IIF(ws_HdPValue4 = "", "", ws_HdPValue4 + "|")
'''''  ws_ColFlds = ws_Flds + IIF(ws_HdPSCd = "", "", "vPSCd" + "|") + _
'''''               IIF(ws_HdPDesc = "", "", "vPDesc" + "|") + IIF(ws_HdPDesc225 = "", "", "vPDesc225" + "|") + _
'''''               IIF(ws_HdPValue = "", "", "vPValue" + "|") + IIF(ws_HdPNum = "", "", "vPNum" + "|") + _
'''''               IIF(ws_HdPValue1 = "", "", "vPValue1" + "|") + IIF(ws_HdPNum1 = "", "", "vPNum1" + "|") + _
'''''               IIF(ws_HdPValue2 = "", "", "vPValue2" + "|") + IIF(ws_HdPValue3 = "", "", "vPValue3" + "|") + _
'''''               IIF(ws_HdPValidYN = "", "", "vPValidYN" + "|") + IIF(ws_HdPValue4 = "", "", "vPValue4" + "|")
'''''' ###########################################  Manoj  ###########################################
'''''
'''''  If ws_ColHeadg = "" Then
'''''    ws_ColHeadg = "|Code             |Desc                                                |"
'''''    ws_ColFlds = "|vPSCd|vPDesc|"
'''''  Else
'''''    ws_ColHeadg = "|" + ws_ColHeadg
'''''    ws_ColFlds = "|" + ws_ColFlds
'''''  End If
'''''
'''''  With mwhlp
'''''    .Sqlstrg = " Select * From vParam " + _
'''''               " Where vPTyp ='" + ps_vPTyp + "' and " + _
'''''               " vPMCd ='" + ps_vPMCd + "' " + ws_Cnd + " order by vPCoCd, vPSCd "
'''''    .ColHeadg = ws_ColHeadg
'''''    .ColFlds = ws_ColFlds
'''''    .RetFldNm = "vPSCd"
'''''    '**********Geeta**Emr206 *****15/04/04
'''''    If Trim(ps_CoCd) = "" Then
'''''      .SearchColNum = 2
'''''    Else
'''''      .SearchColNum = 1
'''''    End If
'''''    '**********Geeta**Emr206 *****15/04/04
'''''  End With
'''''End Sub

'*** JAY 2.14 *** (pb_ShowValidOnly ADDED)
Public Sub vPSCd(ByVal ps_cocd As String, ByVal ps_vPTyp As String, ByVal ps_vPMCd As String, _
                 Optional ByVal ps_vPValue4YN As Variant, Optional ByVal pb_ShowValidOnly As Boolean = False, Optional ByVal ps_vPValue3 As Variant)
'Public Sub vPSCd(ByVal ps_CoCd As String, ByVal ps_vPTyp As String, ByVal ps_vPMCd As String, _
                                            Optional ByVal ps_vPValue4YN As Variant)
'*** JAY 2.14 ***

' **** Zubin 211 **** '
' **** New optional parameter ps_vPValue4YN is added.
' **** This is mainly for vPTyp = 'CHR'. CHR's shld be available for users in ps_vPValue4
' **** Zubin 211 **** '

  'Dim wPTyp As String
' ########################  Manoj #### Ver: 2.0.6 #### Date: 30/03/2004  ########################
  Dim ws_HdPMCd As String, ws_HdPSCd As String, ws_HdPDesc As String
  Dim ws_HdPDesc225 As String, ws_HdPValue As String, ws_HdPNum As String, ws_HdPValue1  As String
  Dim ws_HdPNum1 As String, ws_HdPValue2 As String, ws_HdPValue3 As String, ws_HdPValue4 As String, ws_HdPValidYN As String
  Dim ws_ColHeadg As String, ws_ColFlds As String
  Dim ws_Head As String, ws_Flds As String
  Dim ws_Cnd As String
  
  If Trim(ps_cocd) <> "" Then
      ws_Cnd = ws_Cnd + " and vPCoCd = '" + ps_cocd + "'"
      ws_Head = "": ws_Flds = ""
  ElseIf Trim(ps_cocd) = "" Then
      ws_Head = "Company |": ws_Flds = "vPCoCd|"
  End If
  
  '******************************Geeta**********Emr205****6/05/04**********
    If Trim(ps_vPTyp) = "CHR" Then _
       ws_Cnd = ws_Cnd + " and Upper(substring(vPDesc,LEN (Rtrim(vPDesc))-2,3))<>Upper('XXX') "
  '********************************Geeta************************************
  
  Call GetParamHd("vParam", ps_vPTyp, ws_HdPMCd, ws_HdPSCd, ws_HdPDesc, ws_HdPDesc225, ws_HdPValue, ws_HdPNum, _
                  ws_HdPValue1, ws_HdPNum1, ws_HdPValue2, ws_HdPValue3, ws_HdPValidYN, ws_HdPValue4)
  
  ws_ColHeadg = ws_Head + IIF(ws_HdPSCd = "", "", ws_HdPSCd + "|") + _
                IIF(ws_HdPDesc = "", "", ws_HdPDesc + "|") + IIF(ws_HdPDesc225 = "", "", ws_HdPDesc225 + "|") + _
                IIF(ws_HdPValue = "", "", ws_HdPValue + "|") + IIF(ws_HdPNum = "", "", ws_HdPNum + "|") + _
                IIF(ws_HdPValue1 = "", "", ws_HdPValue1 + "|") + IIF(ws_HdPNum1 = "", "", ws_HdPNum1 + "|") + _
                IIF(ws_HdPValue2 = "", "", ws_HdPValue2 + "|") + IIF(ws_HdPValue3 = "", "", ws_HdPValue3 + "|") + _
                IIF(ws_HdPValidYN = "", "", ws_HdPValidYN + "|") + IIF(ws_HdPValue4 = "", "", ws_HdPValue4 + "|")
  ws_ColFlds = ws_Flds + IIF(ws_HdPSCd = "", "", "vPSCd" + "|") + _
               IIF(ws_HdPDesc = "", "", "vPDesc" + "|") + IIF(ws_HdPDesc225 = "", "", "vPDesc225" + "|") + _
               IIF(ws_HdPValue = "", "", "vPValue" + "|") + IIF(ws_HdPNum = "", "", "vPNum" + "|") + _
               IIF(ws_HdPValue1 = "", "", "vPValue1" + "|") + IIF(ws_HdPNum1 = "", "", "vPNum1" + "|") + _
               IIF(ws_HdPValue2 = "", "", "vPValue2" + "|") + IIF(ws_HdPValue3 = "", "", "vPValue3" + "|") + _
               IIF(ws_HdPValidYN = "", "", "vPValidYN" + "|") + IIF(ws_HdPValue4 = "", "", "vPValue4" + "|")
' ###########################################  Manoj  ###########################################
  
  If ws_ColHeadg = "" Then
    ws_ColHeadg = "|Code             |Desc                                                |"
    ws_ColFlds = "|vPSCd|vPDesc|"
  Else
    ws_ColHeadg = "|" + ws_ColHeadg
    ws_ColFlds = "|" + ws_ColFlds
  End If
  
 ' **** Zubin 211 **** '
 If UCase(ps_vPTyp) = "CHR" And Not IsMissing(ps_vPValue4YN) Then
    ws_Cnd = ws_Cnd + " and (vPValue4= '' Or ',' + vPValue4 + ',' LIKE '%," + gs_UsrCd + ",%') "
 End If
 ' **** Zubin 211 **** '

 'uma *** Repair Chr
  If UCase(ps_vPTyp) = "CHR" And Not IsMissing(ps_vPValue3) Then
    ws_Cnd = ws_Cnd + IIF(ps_vPValue3 = "Y", " and vPValue3 = 'Y'", "")
    ws_Cnd = ws_Cnd + IIF(ps_vPValue3 = "N", " and vPValue3 = 'N'", "")
  End If

  '*** JAY 2.14 ***
  ws_Cnd = ws_Cnd + IIF(pb_ShowValidOnly = False, " ", " and vPValidYn <> 'N' ")
  '*** JAY 2.14 ***
     
  With mwhlp
    .Sqlstrg = " Select * From vParam " + _
               " Where vPTyp ='" + ps_vPTyp + "' and " + _
               " vPMCd ='" + ps_vPMCd + "' " + ws_Cnd + " order by vPCoCd, vPSCd "
      
    .ColHeadg = ws_ColHeadg
    .ColFlds = ws_ColFlds
    .RetFldNm = "vPSCd"
    '**********Geeta**Emr206 *****15/04/04
    If Trim(ps_cocd) = "" Then
      .SearchColNum = 2
    Else
      .SearchColNum = 1
    End If
    .SortAsc = True
    '**********Geeta**Emr206 *****15/04/04
  End With
End Sub
' **** Zubin 211 **** '

Public Sub xPMCd(ByVal ps_xPTyp As String, Optional ByVal pv_xPValue As Variant, Optional ByVal pv_xPValue1 As Variant)
  'Dim wPTyp As String
  Dim ws_HdPMCd As String, ws_HdPSCd As String, ws_HdPDesc As String
  Dim ws_HdPDesc225 As String, ws_HdPValue As String, ws_HdPNum As String, ws_HdPValue1  As String
  Dim ws_HdPNum1 As String, ws_HdPValue2 As String, ws_HdPValue3 As String
  Dim ws_ColHeadg As String, ws_ColFlds As String
  Dim ws_Cnd As String
  
  ws_Cnd = " and (xPCoCd= '" + ctSelfCoCd + "' or xPAppTo= '' or CHARINDEX(','+ '" + gs_CoCd + "' + ',', ','+xPAppTo+',', 1) > 0) "
  If Not IsMissing(pv_xPValue) Then ws_Cnd = IIF(Trim(pv_xPValue) <> "", " and xPValue in (" + pv_xPValue + ") ", " ")
  If Not IsMissing(pv_xPValue1) Then ws_Cnd = ws_Cnd + " and xPValue1= '" + pv_xPValue1 + "' "
  
  Call GetParamHd("xParam", ps_xPTyp, ws_HdPMCd, ws_HdPSCd, ws_HdPDesc, ws_HdPDesc225, ws_HdPValue, ws_HdPNum, _
                  ws_HdPValue1, ws_HdPNum1, ws_HdPValue2, ws_HdPValue3)
  
  ws_ColHeadg = IIF(ws_HdPMCd = "", "", ws_HdPMCd + "|") + IIF(ws_HdPSCd = "", "", ws_HdPSCd + "|") + _
                IIF(ws_HdPDesc = "", "", ws_HdPDesc + "|") + IIF(ws_HdPDesc225 = "", "", ws_HdPDesc225 + "|") + _
                IIF(ws_HdPValue = "", "", ws_HdPValue + "|") + IIF(ws_HdPNum = "", "", ws_HdPNum + "|") + _
                IIF(ws_HdPValue1 = "", "", ws_HdPValue1 + "|") + IIF(ws_HdPNum1 = "", "", ws_HdPNum1 + "|") + _
                IIF(ws_HdPValue2 = "", "", ws_HdPValue2 + "|") + IIF(ws_HdPValue3 = "", "", ws_HdPValue3 + "|")
  ws_ColFlds = IIF(ws_HdPMCd = "", "", "xPMCd" + "|") + IIF(ws_HdPSCd = "", "", "xPSCd" + "|") + _
               IIF(ws_HdPDesc = "", "", "xPDesc" + "|") + IIF(ws_HdPDesc225 = "", "", "xPDesc225" + "|") + _
               IIF(ws_HdPValue = "", "", "xPValue" + "|") + IIF(ws_HdPNum = "", "", "xPNum" + "|") + _
               IIF(ws_HdPValue1 = "", "", "xPValue1" + "|") + IIF(ws_HdPNum1 = "", "", "xPNum1" + "|") + _
               IIF(ws_HdPValue2 = "", "", "xPValue2" + "|") + IIF(ws_HdPValue3 = "", "", "xPValue3" + "|")
  
  If ws_ColHeadg = "" Then
    ws_ColHeadg = "|Code             |Desc                                                  |"
    ws_ColFlds = "|xPMCd|xPDesc|"
  Else
    ws_ColHeadg = "|" + ws_ColHeadg
    ws_ColFlds = "|" + ws_ColFlds
  End If
  
  With mwhlp
    .Sqlstrg = " Select * From xParam " + _
               " Where xPTyp ='" + ps_xPTyp + "' " + ws_Cnd + " order by xPMCd "
    .ColHeadg = ws_ColHeadg
    .ColFlds = ws_ColFlds
    .RetFldNm = "xPMCd"
    .SearchColNum = 1
    .SortAsc = True
  End With
End Sub
Public Sub xPSCd(ByVal ps_xPTyp As String, ByVal ps_xPMCd As String)
  'Dim wPTyp As String
  Dim ws_HdPMCd As String, ws_HdPSCd As String, ws_HdPDesc As String
  Dim ws_HdPDesc225 As String, ws_HdPValue As String, ws_HdPNum As String, ws_HdPValue1  As String
  Dim ws_HdPNum1 As String, ws_HdPValue2 As String, ws_HdPValue3 As String
  Dim ws_ColHeadg As String, ws_ColFlds As String
  Dim ws_Cnd As String
  
  ws_Cnd = " and (xPCoCd= '" + ctSelfCoCd + "' or xPAppTo= '' or CHARINDEX(','+ '" + gs_CoCd + "' + ',', ','+xPAppTo+',', 1) > 0) "
  Call GetParamHd("xParam", ps_xPTyp, ws_HdPMCd, ws_HdPSCd, ws_HdPDesc, ws_HdPDesc225, ws_HdPValue, ws_HdPNum, _
                  ws_HdPValue1, ws_HdPNum1, ws_HdPValue2, ws_HdPValue3)
  
  ws_ColHeadg = IIF(ws_HdPSCd = "", "", ws_HdPSCd + "|") + _
                IIF(ws_HdPDesc = "", "", ws_HdPDesc + "|") + IIF(ws_HdPDesc225 = "", "", ws_HdPDesc225 + "|") + _
                IIF(ws_HdPValue = "", "", ws_HdPValue + "|") + IIF(ws_HdPNum = "", "", ws_HdPNum + "|") + _
                IIF(ws_HdPValue1 = "", "", ws_HdPValue1 + "|") + IIF(ws_HdPNum1 = "", "", ws_HdPNum1 + "|") + _
                IIF(ws_HdPValue2 = "", "", ws_HdPValue2 + "|") + IIF(ws_HdPValue3 = "", "", ws_HdPValue3 + "|")
  ws_ColFlds = IIF(ws_HdPSCd = "", "", "xPSCd" + "|") + _
               IIF(ws_HdPDesc = "", "", "xPDesc" + "|") + IIF(ws_HdPDesc225 = "", "", "xPDesc225" + "|") + _
               IIF(ws_HdPValue = "", "", "xPValue" + "|") + IIF(ws_HdPNum = "", "", "xPNum" + "|") + _
               IIF(ws_HdPValue1 = "", "", "xPValue1" + "|") + IIF(ws_HdPNum1 = "", "", "xPNum1" + "|") + _
               IIF(ws_HdPValue2 = "", "", "xPValue2" + "|") + IIF(ws_HdPValue3 = "", "", "xPValue3" + "|")
  
  If ws_ColHeadg = "" Then
    ws_ColHeadg = "|Code             |Desc                                                |"
    ws_ColFlds = "|xPSCd|xPDesc|"
  Else
    ws_ColHeadg = "|" + ws_ColHeadg
    ws_ColFlds = "|" + ws_ColFlds
  End If
  
  With mwhlp
    .Sqlstrg = " Select * From xParam " + _
               " Where xPTyp ='" + ps_xPTyp + "' and " + _
               " xPMCd ='" + ps_xPMCd + "' order by xPSCd "
    .ColHeadg = ws_ColHeadg
    .ColFlds = ws_ColFlds
    .RetFldNm = "xPSCd"
    .SearchColNum = 1
    .SortAsc = True
  End With
End Sub
Public Sub xvPMCd(ByVal ps_cocd As String, ByVal ps_xvPTyp As String, Optional ByVal pv_xvPValue As Variant, Optional ByVal pv_xvPValue1 As Variant)
  'Dim wPTyp As String
  Dim ws_HdPMCd As String, ws_HdPSCd As String, ws_HdPDesc As String
  Dim ws_HdPDesc225 As String, ws_HdPValue As String, wHdPNum As String, ws_HdPValue1  As String
  Dim wHdPNum1 As String, ws_HdPValue2 As String, ws_HdPValue3 As String
  Dim ws_ColHeadg As String, ws_ColFlds As String
  Dim ws_Head As String, ws_Flds As String
  Dim ws_Cnd As String
  
  If Trim(ps_cocd) <> "" Then
    ws_Cnd = " and (xvPCoCd= '" + ps_cocd + "' or xvPAppTo= '' or CHARINDEX(','+ '" + gs_CoCd + "' + ',', ','+xvPAppTo+',', 1) > 0) "
    ws_Head = "": ws_Flds = ""
  ElseIf Trim(ps_cocd) = "" Then
    ws_Cnd = " and (xvPAppTo= '' or CHARINDEX(','+ '" + gs_CoCd + "' + ',', ','+xvPAppTo+',', 1) > 0) "
    ws_Head = "Company |": ws_Flds = "xvPCoCd|"
  End If

  If Not IsMissing(pv_xvPValue) Then ws_Cnd = ws_Cnd + IIF(Trim(pv_xvPValue) <> "", " and xvPValue in (" + pv_xvPValue + ") ", " ")
  If Not IsMissing(pv_xvPValue1) Then ws_Cnd = ws_Cnd + " and xvPValue1= '" + pv_xvPValue1 + "' "
  
  Call GetParamHd("xvParam", ps_xvPTyp, ws_HdPMCd, ws_HdPSCd, ws_HdPDesc, ws_HdPDesc225, ws_HdPValue, wHdPNum, _
                  ws_HdPValue1, wHdPNum1, ws_HdPValue2, ws_HdPValue3)
  
  ws_ColHeadg = ws_Head + IIF(ws_HdPMCd = "", "", ws_HdPMCd + "|") + IIF(ws_HdPSCd = "", "", ws_HdPSCd + "|") + _
                IIF(ws_HdPDesc = "", "", ws_HdPDesc + "|") + IIF(ws_HdPDesc225 = "", "", ws_HdPDesc225 + "|") + _
                IIF(ws_HdPValue = "", "", ws_HdPValue + "|") + IIF(wHdPNum = "", "", wHdPNum + "|") + _
                IIF(ws_HdPValue1 = "", "", ws_HdPValue1 + "|") + IIF(wHdPNum1 = "", "", wHdPNum1 + "|") + _
                IIF(ws_HdPValue2 = "", "", ws_HdPValue2 + "|") + IIF(ws_HdPValue3 = "", "", ws_HdPValue3 + "|")
  ws_ColFlds = ws_Flds + IIF(ws_HdPMCd = "", "", "xvPMCd" + "|") + IIF(ws_HdPSCd = "", "", "xvPSCd" + "|") + _
               IIF(ws_HdPDesc = "", "", "xvPDesc" + "|") + IIF(ws_HdPDesc225 = "", "", "xvPDesc225" + "|") + _
               IIF(ws_HdPValue = "", "", "xvPValue" + "|") + IIF(wHdPNum = "", "", "xvPNum" + "|") + _
               IIF(ws_HdPValue1 = "", "", "xvPValue1" + "|") + IIF(wHdPNum1 = "", "", "xvPNum1" + "|") + _
               IIF(ws_HdPValue2 = "", "", "xvPValue2" + "|") + IIF(ws_HdPValue3 = "", "", "xvPValue3" + "|")
  
  If ws_ColHeadg = "" Then
    ws_ColHeadg = "|Code             |Desc                                                  |"
    ws_ColFlds = "|xvPMCd|xvPDesc|"
  Else
    ws_ColHeadg = "|" + ws_ColHeadg
    ws_ColFlds = "|" + ws_ColFlds
  End If
  
  With mwhlp
    .Sqlstrg = " Select * From xvParam " + _
               " Where xvPTyp ='" + ps_xvPTyp + "' " + ws_Cnd + " order by xvPCoCd, xvPMCd "
    .ColHeadg = ws_ColHeadg
    .ColFlds = ws_ColFlds
    .RetFldNm = "xvPMCd"
    .SearchColNum = 1
    .SortAsc = True
  End With
End Sub
Public Sub xvPSCd(ByVal ps_cocd As String, ByVal ps_xvPTyp As String, ByVal ps_xvPMCd As String)
  'Dim wPTyp As String
  Dim ws_HdPMCd As String, ws_HdPSCd As String, ws_HdPDesc As String
  Dim ws_HdPDesc225 As String, ws_HdPValue As String, ws_HdPNum As String, ws_HdPValue1  As String
  Dim ws_HdPNum1 As String, ws_HdPValue2 As String, ws_HdPValue3 As String
  Dim ws_ColHeadg As String, ws_ColFlds As String
  Dim ws_Head As String, ws_Flds As String
  Dim ws_Cnd As String
  
  If Trim(ps_cocd) <> "" Then
    ws_Cnd = " and (xvPCoCd= '" + ps_cocd + "' or xvPAppTo= '' or CHARINDEX(','+ '" + gs_CoCd + "' + ',', ','+xvPAppTo+',', 1) > 0) "
    'ws_Cnd = ws_Cnd + " and vPCoCd = '" + ps_CoCd + "'"
    ws_Head = "": ws_Flds = ""
  ElseIf Trim(ps_cocd) = "" Then
    ws_Cnd = " and (xvPAppTo= '' or CHARINDEX(','+ '" + gs_CoCd + "' + ',', ','+xvPAppTo+',', 1) > 0) "
    ws_Head = "Company |": ws_Flds = "xvPCoCd|"
  End If
  
  Call GetParamHd("xvParam", ps_xvPTyp, ws_HdPMCd, ws_HdPSCd, ws_HdPDesc, ws_HdPDesc225, ws_HdPValue, ws_HdPNum, _
                  ws_HdPValue1, ws_HdPNum1, ws_HdPValue2, ws_HdPValue3)
  
  ws_ColHeadg = ws_Head + IIF(ws_HdPSCd = "", "", ws_HdPSCd + "|") + _
                IIF(ws_HdPDesc = "", "", ws_HdPDesc + "|") + IIF(ws_HdPDesc225 = "", "", ws_HdPDesc225 + "|") + _
                IIF(ws_HdPValue = "", "", ws_HdPValue + "|") + IIF(ws_HdPNum = "", "", ws_HdPNum + "|") + _
                IIF(ws_HdPValue1 = "", "", ws_HdPValue1 + "|") + IIF(ws_HdPNum1 = "", "", ws_HdPNum1 + "|") + _
                IIF(ws_HdPValue2 = "", "", ws_HdPValue2 + "|") + IIF(ws_HdPValue3 = "", "", ws_HdPValue3 + "|")
  ws_ColFlds = ws_Flds + IIF(ws_HdPSCd = "", "", "xvPSCd" + "|") + _
               IIF(ws_HdPDesc = "", "", "xvPDesc" + "|") + IIF(ws_HdPDesc225 = "", "", "xvPDesc225" + "|") + _
               IIF(ws_HdPValue = "", "", "xvPValue" + "|") + IIF(ws_HdPNum = "", "", "xvPNum" + "|") + _
               IIF(ws_HdPValue1 = "", "", "xvPValue1" + "|") + IIF(ws_HdPNum1 = "", "", "xvPNum1" + "|") + _
               IIF(ws_HdPValue2 = "", "", "xvPValue2" + "|") + IIF(ws_HdPValue3 = "", "", "xvPValue3" + "|")
  
  If ws_ColHeadg = "" Then
    ws_ColHeadg = "|Code             |Desc                                                |"
    ws_ColFlds = "|xvPSCd|xvPDesc|"
  Else
    ws_ColHeadg = "|" + ws_ColHeadg
    ws_ColFlds = "|" + ws_ColFlds
  End If
  
  With mwhlp
    .Sqlstrg = " Select * From xvParam " + _
               " Where xvPTyp ='" + ps_xvPTyp + "' and " + _
               " xvPMCd ='" + ps_xvPMCd + "' " + ws_Cnd + " order by xvPCoCd, xvPSCd "
    .ColHeadg = ws_ColHeadg
    .ColFlds = ws_ColFlds
    .RetFldNm = "xvPSCd"
    .SearchColNum = 1
    .SortAsc = True
  End With
End Sub
'*** Jay 2.14 *** (pb_ShowValidOnly added)
'****** Sachin 3.02 - Partition Key added
Public Sub RmCd(ByVal mRmZ As String, Optional ByVal mRmCtg As String, _
                Optional ByVal mRmSCtg As String, Optional ByVal pb_ShowValidOnly As Boolean = True, Optional ByVal mPrtKey As String, Optional ByVal mCmCd As String)
'Public Sub RmCd(ByVal mRmZ As String, Optional ByVal mRmCtg As String, _
                Optional ByVal mRmSCtg As String)

Dim wCnd As String, wCmCnd As String
  wCnd = IIF(pb_ShowValidOnly = True, " and RmValidYn = 'Y' ", " ")
  
  'Sachin 3.11.0 - Customer RM Code in help
  'wCmCnd = IIF(mCmCd = "", "", " Left Outer Join CustRm On CrCd='" + mCmCd + "' And CrOurRmCd=RmCd ")
  
  'Sachin - 4.1.0.0 - Max Customer RM Code to be displayed in help (not all)
  wCmCnd = IIF(mCmCd = "", "", " Left Outer Join (Select CrCd, CrOurRmCd, Max(CrCustRmCd) CrCustRmCd " + _
                               "                  From CustRm " + _
                               "                  Group by CrCd, CrOurRmCd) qCustRm On CrCd='" + mCmCd + "' And CrOurRmCd=RmCd ")

'*** Jay 2.14 ***
  '****** Sachin 2.13.0 [12-08-2006] - [Q/W Module]
  Dim ws_QwMod As Boolean
  ws_QwMod = GetQWMod()
  '****** Sachin 2.13.0 [12-08-2006] - [Q/W Module]
  With mwhlp
' ########################  Manoj #### Ver: 2.0.6 #### Date: 30/03/2004  ########################
'*** Jay 2.14 ***
'***** Sachin 3.11.0 - Customer RM in help
    If mRmCtg = "" Then
      .Sqlstrg = " Select RmCd," + IIF(mCmCd = "", "", "IsNull(CrCustRmCd,'') as CrCustRmCd,") + IIF(ws_QwMod = True, "RmQw, RmQwCst,", "") + " RmCtg, RmSCtg, RmIntQly, RmDesc, RmValidYN, RmPrtKey " + _
                 " from RmMst " + wCmCnd + _
                 " where RmZ='" + mRmZ + "' " + wCnd + _
                 IIF(gs_Partition = ctCurrPrtn, " And RmPrtKey= '" + ctCurrPrtn + "' ", "") + _
                 " Order by RmCd "
    Else
      If mRmSCtg = "" Then
        .Sqlstrg = "Select RmCd," + IIF(mCmCd = "", "", "IsNull(CrCustRmCd,'') as CrCustRmCd,") + IIF(ws_QwMod = True, "RmQw, RmQwCst,", "") + " RmCtg, RmSCtg, RmIntQly, RmDesc, RmValidYN, RmPrtKey  " + _
                   "from RmMst " + wCmCnd + _
                   "where RmZ='" + mRmZ + "' and RmCtg in (" + mRmCtg + ") " + wCnd + _
                   IIF(gs_Partition = ctCurrPrtn, " And RmPrtKey= '" + ctCurrPrtn + "' ", "") + _
                   "Order by RmCd "
      Else
        .Sqlstrg = "Select RmCd," + IIF(mCmCd = "", "", "IsNull(CrCustRmCd,'') as CrCustRmCd,") + IIF(ws_QwMod = True, "RmQw, RmQwCst,", "") + " RmCtg, RmSCtg, RmIntQly, RmDesc, RmValidYN, RmPrtKey  " + _
                   "from RmMst " + wCmCnd + _
                   "where RmZ='" + mRmZ + "' and RmCtg in (" + mRmCtg + ") " + _
                   IIF(gs_Partition = ctCurrPrtn, " And RmPrtKey= '" + ctCurrPrtn + "' ", "") + _
                   "and RmSCtg ='" + mRmSCtg + "' " + wCnd + " order by RmCd "
      End If
    End If
    '*** Jay 2.14 ***
    '578.13 Rmcd length increased
    '6.1B-25 Description length increased
    .ColHeadg = "|Code                                  |" + IIF(mCmCd = "", "", " Cust RM Code               |") + IIF(gs_Partition = ctCurrPrtn, "", "Prtn |") + IIF(ws_QwMod = True, "Q/W Sal.   |Q/W Cst.   |", "") + "Ctg |Sub Ctg|International Qly|Description                                                                                                                                                                          |Yn|"
    .ColFlds = "|RmCd|" + IIF(mCmCd = "", "", "CrCustRmCd|") + IIF(gs_Partition = ctCurrPrtn, "", "RmPrtKey|") + IIF(ws_QwMod = True, "RmQw|RmQwCst|", "") + "RmCtg|RmSCtg|RmIntQly|RmDesc|RmValidYN|"
    .RetFldNm = "RmCd"
    .SearchColNum = 1
    .SortAsc = True
' ###########################################  Manoj  ###########################################
  End With
End Sub
Public Sub RmIntQly(Optional ByVal mRmCtg As String, Optional ByVal mRmSCtg As String)
  '****** Sachin 3.02 - Partition Key In Help added
  With mwhlp
    '*** jay 2.14 *** (["RmValidYn = 'Y'" ]ADDED)
    If mRmCtg = "" Then
      .Sqlstrg = "SELECT RmIntQly, RmCtg, RmSCtg, RmCd from RmMst " + _
                 "where RmValidYn = 'Y' " + _
                 IIF(gs_Partition = ctCurrPrtn, " And RmPrtKey= '" + ctCurrPrtn + "' ", "") + _
                 " Order by RmIntQly"
    Else
      If mRmSCtg = "" Then
        .Sqlstrg = "SELECT RmIntQly, RmCtg, RmSCtg, RmCd from RmMst " + _
                   "where RmCtg in (" + mRmCtg + ") " + _
                   IIF(gs_Partition = ctCurrPrtn, " And RmPrtKey= '" + ctCurrPrtn + "' ", "") + _
                   "and RmValidYn = 'Y' order by RmIntQly"
      Else
        .Sqlstrg = "SELECT RmIntQly, RmCtg, RmSCtg, RmCd from RmMst where " + _
                   "RmCtg in (" + mRmCtg + ") and RmSCtg ='" + mRmSCtg + "' " + _
                   IIF(gs_Partition = ctCurrPrtn, " And RmPrtKey= '" + ctCurrPrtn + "' ", "") + _
                   "and RmValidYn = 'Y' order by RmIntQly"
      End If
    End If
    '*** jay 2.14 *** (["RmValidYn = 'Y'" ]ADDED)
    '578.13 Rmcd length increased
    .ColHeadg = "|International Qly|Rm Code                               |Ctg |Sub Ctg|" + IIF(gs_Partition <> ctCurrPrtn, "Prtn |", "")
    .ColFlds = "|RmIntQly|RmCd|RmCtg|RmSCtg|" + IIF(gs_Partition <> ctCurrPrtn, "RmPrtKey|", "")
    .RetFldNm = "RmIntQly"
    .SearchColNum = 1
    .SortAsc = True
  End With
End Sub
'*** Jay 2.14 ***
Public Sub CustCd(ByVal mCmCtg As String, Optional ByVal pb_ShowValidOnly As Boolean = True, Optional ByVal pb_CmCtg As String _
, Optional ByVal mValidCmCd As String)
'Public Sub CustCd(ByVal mCmCtg As String)

  Dim wCnd As String
  'Dim wSrtOpt As String, wOrdBy As String, wSrchCol As Integer
  wCnd = IIF(pb_ShowValidOnly = True, "And CmValidYN = 'Y' ", " ")
'*** Jay 2.14 ***
  '6.1b-26
  If mValidCmCd <> "" Then wCnd = " and '," + mValidCmCd + ",' LIKE '%,'+CmCd+',%' "
  
  '###
  'wSrtOpt = GetCmCdHlpSrt
  'If wSrtOpt = "Y" Then
  '  wOrdBy = " Order by CmName ": wSrchCol = 2
  'Else
  '  wOrdBy = " Order by CmCd ": wSrchCol = 1
  'End If
  '###
  
  With mwhlp
      '### .Sqlstrg = "SELECT CmCd, CmName from CustMst where " + _
                 "CmCtg='" + mCmCtg + "' " + wOrdBy
' ########################  Manoj #### Ver: 2.0.6 #### Date: 30/03/2004  ########################

'*** Jay 2.14 ***
    ' Zubin 213
    If pb_CmCtg = "N" Then
        .Sqlstrg = "SELECT CmCd, CmName, CmCurCd, CmValidYN from CustMst where " + _
                  "1=1 " + wCnd + " order by CmCd "
    
    Else
        .Sqlstrg = "SELECT CmCd, CmName, CmCurCd, CmValidYN from CustMst where " + _
                 "CmCtg='" + mCmCtg + "' " + wCnd + " order by CmCd "
    End If
    '.Sqlstrg = "SELECT CmCd, CmName, CmCurCd, CmValidYN from CustMst where " + _
                 "CmCtg='" + mCmCtg + "' order by CmCd "
'*** Jay 2.14 ***

    '**** Shilpa ****'
    If GetUsrLevel <= 5 Then
        .ColHeadg = "|Code                      |Name                                         |Currency        |Yn|"
        .ColFlds = "|CmCd|CmName|CmCurCd|CmValidYN|"
    Else
        .ColHeadg = "|Code       |Currency        |Yn|"
        .ColFlds = "|CmCd|CmCurCd|CmValidYN|"
    End If
' ###########################################  Manoj  ###########################################
    '**** Shilpa ****'
    .RetFldNm = "CmCd"
    .SearchColNum = 1
    .SortAsc = True
    '### .SearchColNum = wSrchCol
  End With
End Sub
Public Sub DmCd(ByVal mDmTcTyp As String, Optional ByVal mDmCtg As String, Optional ByVal mPrdCtg As String, Optional ByVal mPrtKey As String, Optional ByVal mShowValidOnly As Boolean = False)
  
  With mwhlp
      '*** (Bef 2.14 Next)
      '.Sqlstrg = "SELECT DmCd, DmSz, DmCtg, DmPrdCtg, DmSetCd, DmDesc from DsgMst " + _
      '           "Where DmTcTyp='" + mDmTcTyp + "'" + _
      '           IIF(mDmCtg <> "", " And DmCtg='" + mDmCtg + "'", "") + " order by DmCd,DmSz "
      '*** (Bef 2.14 Next)
      
      '*** (Jen 2.14 Next)
      ' 4.1.3.0 - [Valid YN added]
      'indigo.164 DmTcTyp= '" + mDmTcTyp + "'" changed
      .Sqlstrg = " Select DmCd, DmSz, DmCtg, DmPrdCtg, DmSetCd, DmValidYn, DmDesc, DmPrtKey from DsgMst " + _
                 " Where DmTcTyp In('" + mDmTcTyp + "')" + _
                 IIF(mDmCtg <> "", " And DmCtg= '" + mDmCtg + "'", "") + _
                 IIF(mPrdCtg <> "", " And DmPrdCtg= '" + mPrdCtg + "'", "") + _
                 IIF(gs_Partition = ctCurrPrtn, " And DmPrtKey= '" + ctCurrPrtn + "'", "") + _
                 IIF(mShowValidOnly = True, " And DmValidYn = 'Y' ", "") + _
                 " Order by DmCd, DmSz "
      '*** (Jen 2.14 Next)
    
    ' ***** Manali 3.9.0 - Description displayed long
    '7 degcd width increased
    If mDmTcTyp = "PM" Then
      .ColHeadg = "|Comp Cd                              |Size       |Dsg Ctg|Prd Ctg|Set Cd                 |Description                                                                                                                                                   |Vld  |" + IIF(gs_Partition <> ctCurrPrtn, "Prtn |", "")
    Else
      .ColHeadg = "|Design                               |Size       |Dsg Ctg|Prd Ctg|Set Cd                 |Description                                                                                                                                                   |Vld  |" + IIF(gs_Partition <> ctCurrPrtn, "Prtn |", "")
    End If
    
    ' Manoj 2.10.0
    .ColFlds = "|DmCd|DmSz|DmCtg|DmPrdCtg|DmSetCd|DmDesc|DmValidYn|" + IIF(gs_Partition <> ctCurrPrtn, "DmPrtKey|", "")
    .RetFldNm = "DmCd"
    .SearchColNum = 1
    .SortAsc = True
  End With
End Sub

Public Sub DmOldCd(ByVal mDmTcTyp As String, Optional ByVal mDmCtg As String, Optional ByVal mPrdCtg As String, Optional ByVal mPrtKey As String, Optional ByVal mShowValidOnly As Boolean = False)
  
  With mwhlp
      '*** (Bef 2.14 Next)
      '.Sqlstrg = "SELECT DmCd, DmSz, DmCtg, DmPrdCtg, DmSetCd, DmDesc from DsgMst " + _
      '           "Where DmTcTyp='" + mDmTcTyp + "'" + _
      '           IIF(mDmCtg <> "", " And DmCtg='" + mDmCtg + "'", "") + " order by DmCd,DmSz "
      '*** (Bef 2.14 Next)
      
      '*** (Jen 2.14 Next)
      ' 4.1.3.0 - [Valid YN added]
      .Sqlstrg = " Select DmOldCd, DmCd, DmSz, DmCtg, DmPrdCtg, DmSetCd, DmValidYn, DmDesc, DmPrtKey from DsgMst " + _
                 " Where DmTcTyp= '" + mDmTcTyp + "'" + _
                 IIF(mDmCtg <> "", " And DmCtg= '" + mDmCtg + "'", "") + _
                 IIF(mPrdCtg <> "", " And DmPrdCtg= '" + mPrdCtg + "'", "") + _
                 IIF(gs_Partition = ctCurrPrtn, " And DmPrtKey= '" + ctCurrPrtn + "'", "") + _
                 IIF(mShowValidOnly = True, " And DmValidYn = 'Y' ", "") + _
                 " Order by DmOldCd, DmCd, DmSz "
      '*** (Jen 2.14 Next)
    
    ' ***** Manali 3.9.0 - Description displayed long
    If mDmTcTyp = "PM" Then
      .ColHeadg = "|Old Code            |Comp Cd                |Size  |Dsg Ctg|Prd Ctg|Set Cd                 |Description                                                                                                                                                   |Vld  |" + IIF(gs_Partition <> ctCurrPrtn, "Prtn |", "")
    Else
      .ColHeadg = "|Old Code            |Design                 |Size  |Dsg Ctg|Prd Ctg|Set Cd                 |Description                                                                                                                                                   |Vld  |" + IIF(gs_Partition <> ctCurrPrtn, "Prtn |", "")
    End If
    
    ' Manoj 2.10.0
    .ColFlds = "|DmOldCd|DmCd|DmSz|DmCtg|DmPrdCtg|DmSetCd|DmDesc|DmValidYn|" + IIF(gs_Partition <> ctCurrPrtn, "DmPrtKey|", "")
    .RetFldNm = "DmOldCd"
    .SearchColNum = 1
    .SortAsc = True
  End With
End Sub

'*** Jay 2.14 ***
' ***** Manali Trading Module
'*** Jay 2.14 ***
' ***** Manali Trading Module
Public Sub OdSfx(ByVal mCoCd As String, ByVal mOdDmCd As String, ByVal mOmCmCtg As String, Optional ByVal mOmCmCd As String)
'Public Sub OdSfx(ByVal mCoCd As String, ByVal mOdDmCd As String, Optional ByVal mOmCmCd As String)
'Public Sub OdSfx(ByVal mCoCd As String, ByVal mOdDmCd As String)
'*** Jay 2.14 ***
  
  'ByVal mOmCmCd As String
  With mwhlp
    '******************Geeta*****************Emr205*********************
'*** Jay 2.14 ***
   ' If Not IsMissing(mOmCmCd) Then
      '****** Sachin 3.02 - Partition Key Added ******
      ' ***** Manali Trding Module - mOmCmCtg added
      .Sqlstrg = "SELECT OdPrtKey, OdSfx, OdDmSz, isnull(OdDmCol, '') as qDmCol, " + _
               "isnull(OdKt, '') as qDmKt,OmCmCurCd, OmCmCd, OdSalPrc, OmDt, OdSfxDesc from OrdDsg " + _
               "join Ordmst on OdOmIdNo=OmIdNo and OdCoCd= OmCoCd and OdTc = OmTc and OdYy= OmYy and " + _
               "OdChr = OmChr and OdNo = OmNo " + _
               "where " + _
               IIF(gs_Partition = ctCurrPrtn, " OmPrtKey= '" + ctCurrPrtn + "' And ", "") + _
               " OmCoCd= '" + mCoCd + "' and OmCmCtg='" + mOmCmCtg + "' and " + _
               "OmTc in (Select PMCd from Param where PTyp='TC' and " + _
               "PValue='PL') and OdDmCd='" + mOdDmCd + "' " + _
               "order by (case when OmCmCd = '" + mOmCmCd + "' then ' ' else OmCmCd  end), OdSfx, OdDmSz "
'    Else
'      .Sqlstrg = "SELECT OdSfx, OdDmSz, isnull(OdDmCol, '') as qDmCol, " + _
'               "isnull(OdKt, '') as qDmKt, OmCmCd, OdSalPrc, OmDt, OdSfxDesc from OrdDsg " + _
'               "join Ordmst on OdCoCd= OmCoCd and OdTc = OmTc and OdYy= OmYy and " + _
'               "OdChr = OmChr and OdNo = OmNo " + _
'               "where OmCoCd= '" + mCoCd + "' and " + _
'               "OmTc in (Select PMCd from Param where PTyp='TC' and " + _
'               "PValue='PL') and OdDmCd='" + mOdDmCd + "' order by OmCmCd, OdSfx, OdDmSz "
'               '"join DsgMst On DmTcTyp ='DM' and " + _
'               '"DmCd=OdDmCd and DmSz=''
'               'and OmCmCd in ('" + mOmCmCd + "', '" + ctSelfCmCd + "')
'               '(index=OrdMst) '(index=DsgMst)
'    End If
    '*** Jay 2.14 ***
    .ColHeadg = "|Suffix                                   |Size  |" + IIF(gs_Partition <> ctCurrPrtn, "Prtn |", "") + "Colour |Karat  |Customer     |Sales Price|Currency|Date        |Desc                                                        |"
    .ColFlds = "|OdSfx|OdDmSz|" + IIF(gs_Partition <> ctCurrPrtn, "OdPrtKey|", "") + "qDmCol|qDmKt|OmCmCd|OdSalPrc|OmCmCurCd|OmDt|OdSfxDesc|"
    .RetFldNm = "OdSfx"
    .SearchColNum = 1
    .SortAsc = True
  End With
End Sub
' Zubin 213
' ***** Manali Trading Module - mCmCd as added as parameter
Public Sub OmNoJL(ByVal mCoCd As String, ByVal mTc As String, ByVal mYy As String, ByVal mChr As String, Optional mCurCd As Variant, Optional ByVal mPrtKey As String, _
    Optional ByVal mCmCtg As String, Optional ByVal mCmCd As String, Optional ByVal mb_ShowBalQty As Boolean = False, Optional ByVal mb_BaseTc As Boolean = False)
'Prev
'wSqlStrg = "SELECT OmNo,OmDt,OmDelDt,OmCmCd, OmPrtCd from OrdMst " + _
'             "where OmCoCd= '" + mCoCd + "' and " + _
'             "OmTc= '" + mTc + "' and OmYy='" + mYy + "' and " + _
'             "OmChr='" + mChr + "' Order By OmNo"
'  With mwhlp
'    .Sqlstrg = wSqlStrg
'    .ColHeadg = "|Number|Date        |Del Date    |Customer  |Priority|"
'    .ColFlds = "|OmNo|OmDt|OmDelDt|OmCmCd|OmPrtCd|"
'    .ColMask = "||dd/mm/yy|dd/mm/yy|||"
'    .SearchColNum = 1
'    .RetFldNm = "OmNo"
'  End With
'****Geeta***Crm**212

  Dim ws_Cnd As String
  ws_Cnd = IIF(IsMissing(mCurCd) = True, "", " and OmCmCurCd= '" + CStr(mCurCd) + "' ")
  ' ***** Manali Trading Module - mCmCd as added as parameter
  ws_Cnd = ws_Cnd + IIF(mCmCtg = "", "", " and OmCmCtg= '" + mCmCtg + "' ")
  ws_Cnd = ws_Cnd + IIF(mCmCd = "", "", " and OmCmCd= '" + mCmCd + "' ")
  
  Dim wTcTyp As String
  wTcTyp = go_HlpCn.GetFldVal("Select PValue from Param Where PTyp='TC' and PMCd='" + mTc + "' ")
  ' ***** Manali Trading Module - mCmCd as added as parameter
  
' Zubin 213 (OmCmCurCd added)
' ****** Sachin 2.14.0 - [25-11-2006] - [Multi-Company] - Blank Company Code Handled ******
'****** Sachin 3.02.0 - Partition Key added
' ***** Manali - 3.8.0 - OmPoNo added
If UCase(mChr) = "WEB" Then
  wSqlStrg = "SELECT OmPrtKey,OmNo,OmDt,OmDelDt,OmCmCd, OmPoNo, OmCmCurCd, OmPrtKey, OmPrtCd,OmCrmYy+'/'+OmCrmChr+'/'+LTrim(OmCrmNo) as qCrmVch, " + _
             "OmSubmitYN,OmAckYN  from OrdMst " + _
             "where " + _
             IIF(gs_Partition = ctCurrPrtn Or mPrtKey = ctCurrPrtn, " OmPrtKey= '" + ctCurrPrtn + "' And ", "") + _
             " OmCoCd= '" + mCoCd + "' And " + _
             "OmTc= '" + mTc + "' and OmYy='" + mYy + "' and " + _
             "OmChr='" + mChr + "' " + _
             ws_Cnd + " Order By OmNo Desc"
ElseIf wTcTyp = ctJtTcTyp And mb_ShowBalQty Then
  wSqlStrg = "SELECT OmPrtKey,OmNo,OmDt,OmDelDt,OmCmCd, OmPoNo, OmCmCurCd, OmPrtKey, OmPrtCd, qBalQty from OrdMst " + _
             "Join (Select OdOmIdNo, OdPrtKey, OdCoCd, OdTc, OdYy, OdChr, OdNo, Convert(Decimal(14,1), Sum(OdJLWt - OdJLExpWt)) as qBalQty " + _
             "      from OrdDsg Group By OdOmIdNo, OdPrtKey, OdCoCd, OdTc, OdYy, OdChr, OdNo ) " + _
             "OrdDsg On OdCoCd=OmCoCd and OdTc=OmTc and OdYy=OmYy and OdChr=OmChr and OdNo=OmNo and OdOmIdNo=OmIdNo and OdPrtKey=OmPrtKey " + _
             "where " + _
             IIF(gs_Partition = ctCurrPrtn Or mPrtKey = ctCurrPrtn, " OmPrtKey= '" + ctCurrPrtn + "' And ", "") + _
             " OmCoCd= '" + mCoCd + "' And " + _
             "OmTc= '" + mTc + "' and OmYy='" + mYy + "' and " + _
             "OmChr='" + mChr + "' " + ws_Cnd + _
             IIF(mb_BaseTc, " and qBalQty > 0.0 ", "") + _
             " Order By OmNo Desc"
Else
  wSqlStrg = "SELECT OmPrtKey,OmNo,OmDt,OmDelDt,OmCmCd, OmPoNo, OmCmCurCd, OmPrtKey, OmPrtCd from OrdMst " + _
             "where " + _
             IIF(gs_Partition = ctCurrPrtn Or mPrtKey = ctCurrPrtn, " OmPrtKey= '" + ctCurrPrtn + "' And ", "") + _
             " OmCoCd= '" + mCoCd + "' And " + _
             "OmTc= '" + mTc + "' and OmYy='" + mYy + "' and " + _
             "OmChr='" + mChr + "' " + _
             ws_Cnd + " Order By OmNo Desc"
  
End If

  With mwhlp
    .Sqlstrg = wSqlStrg
    .ColHeadg = "|Number|Date        |Del Date    |" + IIF(gs_Partition <> ctCurrPrtn, "Prtn |", "") + "Customer       |Po No                                          |Currency   |Priority|" + IIF(UCase(mChr) = "WEB", "CrmVchNo             |Submit|Ack|", "") + IIF(wTcTyp = ctJtTcTyp And mb_ShowBalQty, "Bal Qty  |", "")
    .ColFlds = "|OmNo|OmDt|OmDelDt|" + IIF(gs_Partition <> ctCurrPrtn, "OmPrtKey|", "") + "OmCmCd|OmPoNo|OmCmCurCd|OmPrtCd|" + IIF(UCase(mChr) = "WEB", "qCrmVch|OmSubmitYN|OmAckYN|", "") + IIF(wTcTyp = ctJtTcTyp And mb_ShowBalQty, "qBalQty|", "")
    '6.1
    .ColMask = "||" + DtShortStr() + "|" + DtShortStr() + "|" + IIF(gs_Partition <> ctCurrPrtn, "|", "") + "||||" + IIF(UCase(mChr) = "WEB", "|||", "") + IIF(wTcTyp = ctJtTcTyp And mb_ShowBalQty, "|", "")
    .SearchColNum = 1
    .RetFldNm = "OmNo"
    .SortAsc = False
  End With
End Sub
Public Sub OmNo(ByVal mCoCd As String, ByVal mTc As String, ByVal mYy As String, ByVal mChr As String, Optional mCurCd As Variant, Optional ByVal mPrtKey As String, _
    Optional ByVal mCmCtg As String, Optional ByVal mCmCd As String, Optional ByVal mb_ShowBalQty As Boolean = False, Optional ByVal mb_BaseTc As Boolean = False)
'Prev
'wSqlStrg = "SELECT OmNo,OmDt,OmDelDt,OmCmCd, OmPrtCd from OrdMst " + _
'             "where OmCoCd= '" + mCoCd + "' and " + _
'             "OmTc= '" + mTc + "' and OmYy='" + mYy + "' and " + _
'             "OmChr='" + mChr + "' Order By OmNo"
'  With mwhlp
'    .Sqlstrg = wSqlStrg
'    .ColHeadg = "|Number|Date        |Del Date    |Customer  |Priority|"
'    .ColFlds = "|OmNo|OmDt|OmDelDt|OmCmCd|OmPrtCd|"
'    .ColMask = "||dd/mm/yy|dd/mm/yy|||"
'    .SearchColNum = 1
'    .RetFldNm = "OmNo"
'  End With
'****Geeta***Crm**212

  Dim ws_Cnd As String
  ws_Cnd = IIF(IsMissing(mCurCd) = True, "", " and OmCmCurCd= '" + CStr(mCurCd) + "' ")
  ' ***** Manali Trading Module - mCmCd as added as parameter
  ws_Cnd = ws_Cnd + IIF(mCmCtg = "", "", " and OmCmCtg= '" + mCmCtg + "' ")
  ws_Cnd = ws_Cnd + IIF(mCmCd = "", "", " and OmCmCd= '" + mCmCd + "' ")
  
  Dim wTcTyp As String
  wTcTyp = go_HlpCn.GetFldVal("Select PValue from Param Where PTyp='TC' and PMCd='" + mTc + "' ")
  ' ***** Manali Trading Module - mCmCd as added as parameter
  
' Zubin 213 (OmCmCurCd added)
' ****** Sachin 2.14.0 - [25-11-2006] - [Multi-Company] - Blank Company Code Handled ******
'****** Sachin 3.02.0 - Partition Key added
' ***** Manali - 3.8.0 - OmPoNo added
If UCase(mChr) = "WEB" Then
  wSqlStrg = "SELECT OmPrtKey,OmNo,OmDt,OmDelDt,OmCmCd, OmPoNo, OmCmCurCd, OmPrtKey, OmPrtCd,OmCrmYy+'/'+OmCrmChr+'/'+LTrim(OmCrmNo) as qCrmVch, " + _
             "OmSubmitYN,OmAckYN  from OrdMst " + _
             "where " + _
             IIF(gs_Partition = ctCurrPrtn Or mPrtKey = ctCurrPrtn, " OmPrtKey= '" + ctCurrPrtn + "' And ", "") + _
             " OmCoCd= '" + mCoCd + "' And " + _
             "OmTc= '" + mTc + "' and OmYy='" + mYy + "' and " + _
             "OmChr='" + mChr + "' " + _
             ws_Cnd + " Order By OmNo Desc"
ElseIf wTcTyp = ctJtTcTyp And mb_ShowBalQty Then
  wSqlStrg = "SELECT OmPrtKey,OmNo,OmDt,OmDelDt,OmCmCd, OmPoNo, OmCmCurCd, OmPrtKey, OmPrtCd, qBalQty from OrdMst " + _
             "Join (Select OdOmIdNo, OdPrtKey, OdCoCd, OdTc, OdYy, OdChr, OdNo, Convert(Decimal(14,1), Sum(OdOrdQty - OdExpQty)) as qBalQty " + _
             "      from OrdDsg Group By OdOmIdNo, OdPrtKey, OdCoCd, OdTc, OdYy, OdChr, OdNo ) " + _
             "OrdDsg On OdCoCd=OmCoCd and OdTc=OmTc and OdYy=OmYy and OdChr=OmChr and OdNo=OmNo and OdOmIdNo=OmIdNo and OdPrtKey=OmPrtKey " + _
             "where " + _
             IIF(gs_Partition = ctCurrPrtn Or mPrtKey = ctCurrPrtn, " OmPrtKey= '" + ctCurrPrtn + "' And ", "") + _
             " OmCoCd= '" + mCoCd + "' And " + _
             "OmTc= '" + mTc + "' and OmYy='" + mYy + "' and " + _
             "OmChr='" + mChr + "' " + ws_Cnd + _
             IIF(mb_BaseTc, " and qBalQty > 0.0 ", "") + _
             " Order By OmNo Desc"
Else
  wSqlStrg = "SELECT OmPrtKey,OmNo,OmDt,OmDelDt,OmCmCd, OmPoNo, OmCmCurCd, OmPrtKey, OmPrtCd from OrdMst " + _
             "where " + _
             IIF(gs_Partition = ctCurrPrtn Or mPrtKey = ctCurrPrtn, " OmPrtKey= '" + ctCurrPrtn + "' And ", "") + _
             " OmCoCd= '" + mCoCd + "' And " + _
             "OmTc= '" + mTc + "' and OmYy='" + mYy + "' and " + _
             "OmChr='" + mChr + "' " + _
             ws_Cnd + " Order By OmNo Desc"
  
End If

  With mwhlp
    .Sqlstrg = wSqlStrg
    .ColHeadg = "|Number|Date        |Del Date    |" + IIF(gs_Partition <> ctCurrPrtn, "Prtn |", "") + "Customer       |Po No                                          |Currency   |Priority|" + IIF(UCase(mChr) = "WEB", "CrmVchNo             |Submit|Ack|", "") + IIF(wTcTyp = ctJtTcTyp And mb_ShowBalQty, "Bal Qty  |", "")
    .ColFlds = "|OmNo|OmDt|OmDelDt|" + IIF(gs_Partition <> ctCurrPrtn, "OmPrtKey|", "") + "OmCmCd|OmPoNo|OmCmCurCd|OmPrtCd|" + IIF(UCase(mChr) = "WEB", "qCrmVch|OmSubmitYN|OmAckYN|", "") + IIF(wTcTyp = ctJtTcTyp And mb_ShowBalQty, "qBalQty|", "")
    '6.1
    .ColMask = "||" + DtShortStr() + "|" + DtShortStr() + "|" + IIF(gs_Partition <> ctCurrPrtn, "|", "") + "||||" + IIF(UCase(mChr) = "WEB", "|||", "") + IIF(wTcTyp = ctJtTcTyp And mb_ShowBalQty, "|", "")
    .SearchColNum = 1
    .RetFldNm = "OmNo"
    .SortAsc = False
  End With
End Sub



Public Sub ReNo(ByVal mCoCd As String, ByVal mYy As String, ByVal mChr As String)
  '*** (Jen 2.13)
  '****** Sachin 3.02.0 - Partition Key added
  wSqlStrg = "SELECT ReNo, ReDt, ReCmCd from RetMem where " + _
            " ReCoCd= '" + mCoCd + "' and ReYy= '" + mYy + "' and ReChr= '" + mChr + "' Order By ReNo Desc"
  
  With mwhlp
    .Sqlstrg = wSqlStrg
    .ColHeadg = "|Number      |Date        |Customer       |"
    .ColFlds = "|ReNo|ReDt|ReCmCd|"
    .ColMask = "||" + DtShortStr() + "||"   '6.1
    .SearchColNum = 1
    .RetFldNm = "ReNo"
    .SortAsc = False
  End With
  '*** (Jen 2.13)
End Sub

Public Sub TNo(ByVal mCoCd As String, ByVal mTc As String, ByVal mYy As String, ByVal mChr As String, Optional ByVal mPrtKey As String, Optional ByVal mPrgDt As Boolean = False)
  If mTc = "FM" Then Call FmNo(mCoCd, mTc, mYy, mChr): Exit Sub
  Dim wPurgeDt As String
  wPurgeDt = GetPurgeDt(mCoCd, "")  'sv.9 getting purge date
  '****** Sachin 3.02 24-12-07 - Partition Key added
  '****** Manali 3.5.0 - 07/11/08 - Supplier Code added
  wSqlStrg = "SELECT TNo,TDt,TPrtKey, TSuppCd from Txn where TCoCd= '" + mCoCd + "' " + _
             "and TTc= '" + mTc + "' and TYy='" + mYy + "' and " + _
             "TChr='" + mChr + "' " + _
             IIF(gs_Partition = ctCurrPrtn Or mPrtKey = ctCurrPrtn, " And TPrtKey= '" + ctCurrPrtn + "' ", "") + _
             IIF(gs_Partition = ctCurrPrtn, IIF(mPrgDt = False, " And Tdt>'" + wPurgeDt + "' ", ""), "") + _
             " Order By TNo Desc"
  With mwhlp
    .Sqlstrg = wSqlStrg
    .ColHeadg = "|Number|Date        |" + IIF(mTc = "PR" Or mTc = "SA" Or mTc = "TZ", "Supplier       |", "") + IIF(gs_Partition <> ctCurrPrtn, "Prtn |", "")
    .ColFlds = "|TNo|TDt|" + IIF(mTc = "PR" Or mTc = "SA" Or mTc = "TZ", "TSuppCd|", "") + IIF(gs_Partition <> ctCurrPrtn, "TPrtKey|", "")
    '6.1
    .ColMask = "||" + DtShortStr() + "|" + IIF(mTc = "PR" Or mTc = "SA" Or mTc = "TZ", "|", "") + IIF(gs_Partition <> ctCurrPrtn, "|", "")
    .SearchColNum = 1
    .RetFldNm = "TNo"
    .SortAsc = False
  End With

End Sub
'************************Geeta******Emr206****18/05/04**********************************
Public Sub TrYmd(ByVal mCoCd As String)
 ' wSqlStrg = "SELECT  TrYmd  from (select distinct TrYmd from TreeMst where TrCoCd='" + mCoCd + "' ) Tr Order By cast(TrYmd as integer)"
  wSqlStrg = "SELECT distinct  TrYmd  from TreeMst where TrCoCd='" + mCoCd + "' Order By TrYmd Desc"
  With mwhlp
    .Sqlstrg = wSqlStrg
    .ColHeadg = "|Tree YyMmDd|"
    .ColFlds = "|TrYmd|"
    .ColMask = "||"
    .SearchColNum = 1
    .RetFldNm = "TrYmd"
    .SortAsc = False
  End With
End Sub
'******************************Geeta******************************************
Public Sub TrNo(ByVal mCoCd As String, ByVal mYmd As String, ByVal mChr As String)
  wSqlStrg = "SELECT TrNo,TrDt,TrCstDt from TreeMst where TrCoCd= '" + mCoCd + "' " + _
             "and TrYmd='" + mYmd + "' and " + _
             "TrChr='" + mChr + "' Order By TrNo Desc"
  With mwhlp
    .Sqlstrg = wSqlStrg
    .ColHeadg = "|Number|Tree Date     |Cast Date  |"
    .ColFlds = "|TrNo|TrDt|TrCstDt|"
    .ColMask = "||" + DtShortStr() + "|" + DtShortStr() + "|"   '6.1
    .SearchColNum = 1
    .RetFldNm = "TrNo"
    .SortAsc = False
  End With
End Sub


'Public Sub BNo(ByVal mYy As String, ByVal mChr As String, ByVal mhOthLocBag As String, Optional ByVal mLoc As Variant)
''Bef 03/05/02
'  If Not IsMissing(mLoc) Then
'    wLocStr = IIf(mLoc = "", " ", " and BLoc= '" + mLoc + "' ")
'    wSqlStrg = "SELECT 'A' as qTag, BNo, BOdDmCd, BOdSfx, BOdDmSz, " + _
'               "BOdTc+'/'+BOdYy+'/'+BOdChr+'/'+LTrim(Str(BOdNo))+'/'+LTrim(Str(BOdSr)) " + _
'               "as qOrdNo, BLoc, BQty from Bag " + _
'               "where BYy= '" + mYy + "' and BChr='" + mChr + "' " + wLocStr
'    If mhOthLocBag = "Y" Then
'      wStr = " Union All " + _
'             "SELECT 'B' as qTag, BNo, BOdDmCd, BOdSfx, BOdDmSz, " + _
'             "BOdTc+'/'+BOdYy+'/'+BOdChr+'/'+LTrim(Str(BOdNo))+'/'+LTrim(Str(BOdSr)) " + _
'             "as qOrdNo, BLoc, BQty from Bag " + _
'             "where BYy= '" + mYy + "' and BChr='" + mChr + "' and BLoc<> '" + mLoc + "' "
'    Else
'      wStr = " "
'    End If
'    wSqlStrg = wSqlStrg + wStr + " Order By qTag, BNo"
'  Else
'    wSqlStrg = "SELECT BNo, BOdDmCd, BOdSfx, BOdDmSz, " + _
'               "BOdTc+'/'+BOdYy+'/'+BOdChr+'/'+LTrim(Str(BOdNo))+'/'+LTrim(Str(BOdSr)) " + _
'               "as qOrdNo, BLoc, BQty from Bag " + _
'               "where BYy= '" + mYy + "' and BChr='" + mChr + "' Order By BNo"
'  End If
'  With mwhlp
'    .Sqlstrg = wSqlStrg
'    .ColHeadg = "|Bag Number|Design           |Suffix              |Size |Order                                  |Loc    |Qty     |"
'    .ColFlds = "|BNo|BOdDmCd|BOdSfx|BOdDmSz|qOrdNo|BLoc|BQty|"
'    .SearchColNum = 1
'    .RetFldNm = "BNo"
'  End With
'End Sub
'5717-20 help is coming when dsgcd entered in prdpts table, wDmCd changed as mDmCd
Public Sub PRDPTS(ByVal mPrdTyp As String, ByVal mPrcLoc As String, ByVal mSubPrc As String, ByVal mDmTcTyp As String, ByVal mDmCd As String, ByVal mSetCd As String, ByVal mRmQty As Long)
            wSqlStrg = "Select PpLoc,PpSetTYp,PpFrRmWt, PpToRmWt,PpPts,PpPts*" + CStr(mRmQty) + " as TotPts from PrdPts where PpTyp= '" + mPrdTyp + "' and " + _
                      "PpLoc= '" + mPrcLoc + "' and " + _
                      "PpSubPrc= '" + mSubPrc + "' and (PpDmTcTyp='' or PpDmTcTyp='" + mDmTcTyp + "') and (PpDmCd='' or PpDmCd= '" + mDmCd + "') and " + _
                      "PpSetTyp = '" + mSetCd + "'"

  With mwhlp
    .Sqlstrg = wSqlStrg
    .ColHeadg = "|Process    |Setting   |       Fr Wt|         To Wt|    Points|     Tot Points|"
    .ColFlds = "|PpLoc|PpSetTyp|PpFrRmWt|PpToRmWt|PpPts|TotPts|"
    .ColMask = "|||###0.000|###0.000|####0|####0|"
    .SearchColNum = 6
    .RetFldNm = "TotPts"
    .SortAsc = False
  End With
End Sub



Public Sub BNo(ByVal mCoCd As String, ByVal mYy As String, ByVal mChr As String, ByVal mhOthLocBag As String, Optional ByVal mLoc As Variant, Optional ByVal mOdTc As Variant, Optional ByVal mOdYy As Variant, Optional ByVal mOdChr As Variant, Optional ByVal mOdNo As Variant, Optional ByVal mOdDmCd As Variant, Optional ByVal mSameDesignBags As Boolean = False, Optional ByVal mPrtKey As String, Optional ByVal mRprYN As String, _
Optional ByVal mOdSfx As Variant, Optional ByVal mOdDmSz As Variant, Optional ByVal mCmCd As Variant, Optional ByVal mBagAheadLocPrcs As Variant, Optional ByVal mLocTyp As Variant)
'uma **** Repair added RepairYN parameter

  Dim wOrdCnd1 As String, wOrdCnd2 As String, wLocStr As String, wSameDsgBag As String, wRprYN As String
  Dim wOdSfx As String, wOdDmSz As String, wOrdCnd As String, wCmCd As String, wCmCdJn As String, wBagAheadLocPrcs As String
  
  '****** Sachin 3.02.0 - Partition Key added
  If Not (IsMissing(mOdTc) Or IsMissing(mOdYy) Or IsMissing(mOdChr) Or IsMissing(mOdNo) Or IsMissing(mOdDmCd)) Then
    'Ec.4- added suffix, size, customer, bag ahead location checks
    If Not IsMissing(mOdSfx) Then
      If mOdSfx <> "" Then wOdSfx = " and BOdSfx= '" + mOdSfx + "' "
    End If
    If Not IsMissing(mOdDmSz) Then
      If mOdDmSz <> "" Then wOdDmSz = " and BOdDmSz= '" + mOdDmSz + "' "
    End If
    If Not IsMissing(mCmCd) Then
      If mCmCd <> "" Then
        wCmCd = " and OmCmCd= '" + mCmCd + "' "
        wCmCdJn = " join OrdMst on OmCoCd =BCoCd and OmTc =BOdTc and OmYy =BOdYy and OmChr =BOdChr and OmNo =BOdNo "
      End If
    End If
    If Not IsMissing(mBagAheadLocPrcs) Then
      If mBagAheadLocPrcs <> "" And mLocTyp = "P" Then wBagAheadLocPrcs = " and BLoc in ('" + mBagAheadLocPrcs + "')"
      If mBagAheadLocPrcs <> "" And mLocTyp = "R" Then wBagAheadLocPrcs = " and BLoc in (select LocCd from Loc where LocCoCd ='" + mCoCd + "' and LocTyp='P' " + _
                          "and LocPrcs in ('" + mBagAheadLocPrcs + "'))"
    End If
    
    If mOdTc <> "" And mOdYy <> "" And mOdChr <> "" And mOdNo <> "" Then
      wOrdCnd1 = " and BCoCd='" + mCoCd + "' and BOdTc= '" + mOdTc + "' and" + _
                 " BOdYy= '" + mOdYy + "' and BOdChr= '" + mOdChr + "' and" + _
                 " BOdNo= " + CStr(mOdNo) + " and BOdDmCd= '" + mOdDmCd + "' "
    Else
      wOrdCnd1 = " and BCoCd='" + mCoCd + "' and BOdDmCd= '" + mOdDmCd + "' "
    End If
    wOrdCnd2 = " and not (BCoCd='" + mCoCd + "' and BOdTc= '" + mOdTc + "' and" + _
               " BOdYy= '" + mOdYy + "' and BOdChr= '" + mOdChr + "' and" + _
               " BOdNo= " + CStr(mOdNo) + " and BOdDmCd= '" + mOdDmCd + "') "
    If Not IsMissing(mLoc) And mhOthLocBag = "N" Then
            wLocStr = " and CHARINDEX(','+BLoc+',' , '," + mLoc + ",') <> 0 "          '***Bhavna help for bags in multiple locations -- order by locations selected
'            wLocStr = IIF(mLoc = "", " ", " and BLoc= '" + mLoc + "' ")
    Else
      wLocStr = " "
    End If
'    wSqlStrg = "SELECT 'A' + dbo.CStr(dbo.InStr('" + mLoc + "', BLOC)) as qTag, BPrtKey, BNo, BOdDmCd, BOdSfx, BOdDmSz, " + _
               "BOdTc+'/'+BOdYy+'/'+BOdChr+'/'+LTrim(Str(BOdNo))+'/'+LTrim(Str(BOdSr)) " + _
               "as qOrdNo, BLoc, BQty, BFgSubLoc from Bag " + _
               "where BCoCd='" + mCoCd + "' and BYy= '" + mYy + "' and BChr='" + mChr + "' " + _
               IIF(gs_Partition = ctCurrPrtn, " And BPrtKey='" + ctCurrPrtn + "'", "") + _
                wLocStr + wOrdCnd1 + _
               " Union All " + _
               "SELECT 'B' as qTag, BPrtKey, BNo, BOdDmCd, BOdSfx, BOdDmSz, " + _
               "BOdTc+'/'+BOdYy+'/'+BOdChr+'/'+LTrim(Str(BOdNo))+'/'+LTrim(Str(BOdSr)) " + _
               "as qOrdNo, BLoc, BQty, BFgSubLoc from Bag " + _
               "where " + _
               IIF(gs_Partition = ctCurrPrtn, " BPrtKey='" + ctCurrPrtn + "' And ", "") + _
               " BCoCd='" + mCoCd + "' and BYy='" + mYy + "' and BChr='" + mChr + "' " + _
                wLocStr + wOrdCnd2
              
    '*******************Bhavna help for bags in multiple locations -- order by locations selected
    wSqlStrg = "Select case CHARINDEX(','+BLoc+',' , '," + mLoc + ",') when 0 then 99 else CHARINDEX(','+BLoc+',' , '," + mLoc + ",') end as qTagNo, " + _
               "BPrtKey, BNo, BOdDmCd, BOdSfx, BOdDmSz, " + _
               "BOdTc+'/'+BOdYy+'/'+BOdChr+'/'+LTrim(Str(BOdNo))+'/'+LTrim(Str(BOdSr)) " + _
               "as qOrdNo, BLoc, BQty, BFgSubLoc from Bag " + wCmCdJn + _
               "where BCoCd='" + mCoCd + "' and BYy= '" + mYy + "' and BChr='" + mChr + "' " + _
               IIF(gs_Partition = ctCurrPrtn, " And BPrtKey='" + ctCurrPrtn + "'", "") + _
                wLocStr + wOrdCnd1 + wOdSfx + wOdDmSz + wCmCd + wBagAheadLocPrcs
    
    wSqlStrg = wSqlStrg + " Order By qTagNo, BNo Desc, BQty "
    '*******************Bhavna help for bags in multiple locations -- order by locations selected
  Else
    If Not IsMissing(mLoc) Then
      'wLocStr = IIf(mLoc = "", " ", " and BLoc= '" + mLoc + "' ")
           '*******************Bhavna help for bags in multiple locations -- order by locations selected
      If mhOthLocBag = "Y" Then
            wLocStr = ""
      Else
            wLocStr = " and CHARINDEX(','+BLoc+',' , '," + mLoc + ",') <> 0 "
      End If
           '*******************Bhavna help for bags in multiple locations -- order by locations selected
           
      '*** Jay 2.14 *** [FG Bag Copy Into DsgMst]
      wSameDsgBag = IIF(mSameDesignBags = True And (Not IsMissing(mOdDmCd)), " and BOdDmCd = '" + IIF(Not IsMissing(mOdDmCd), mOdDmCd, "") + "'", " ")
      '*** Jay 2.14 *** [FG Bag Copy Into DsgMst]
      
      'uma **** for Repair bag
      'wRprYN = IIF(mRprYN = "" Or IsMissing(mRprYN), " and BRepYN = 'N'", " and BRepYN = 'Y'")
      If mRprYN = "Y" Then
      wSqlStrg = "Select case CHARINDEX(','+BLoc+',' , '," + mLoc + ",') when 0 then 99 else CHARINDEX(','+BLoc+',' , '," + mLoc + ",') end as qTagNo, " + _
                 "BPrtKey, BNo, BOdDmCd, BOdSfx, BOdDmSz, " + _
                 "BOdTc+'/'+BOdYy+'/'+BOdChr+'/'+LTrim(Str(BOdNo))+'/'+LTrim(Str(BOdSr)) " + _
                 "as qOrdNo, BLoc, BQty, BFgSubLoc, BRepQty, BRepLoc from Bag " + _
                 "where " + _
                 IIF(gs_Partition = ctCurrPrtn Or mPrtKey = ctCurrPrtn, " BPrtKey='" + ctCurrPrtn + "' And ", "") + _
                 "BCoCd='" + mCoCd + "' and BYy= '" + mYy + "' and BChr='" + mChr + "'  and BRepYN = 'Y'" + wLocStr + wSameDsgBag
               
      Else
      '*******************Bhavna help for bags in multiple locations -- order by locations selected
      wSqlStrg = "Select case CHARINDEX(','+BLoc+',' , '," + mLoc + ",') when 0 then 99 else CHARINDEX(','+BLoc+',' , '," + mLoc + ",') end as qTagNo, " + _
                 "BPrtKey, BNo, BOdDmCd, BOdSfx, BOdDmSz, " + _
                 "BOdTc+'/'+BOdYy+'/'+BOdChr+'/'+LTrim(Str(BOdNo))+'/'+LTrim(Str(BOdSr)) " + _
                 "as qOrdNo, BLoc, BQty, BFgSubLoc from Bag " + _
                 "where " + _
                 IIF(gs_Partition = ctCurrPrtn Or mPrtKey = ctCurrPrtn, " BPrtKey='" + ctCurrPrtn + "' And ", "") + _
                 "BCoCd='" + mCoCd + "' and BYy= '" + mYy + "' and BChr='" + mChr + "' " + wLocStr + wSameDsgBag
       '*******************Bhavna help for bags in multiple locations -- order by locations selected
      End If
'      If mhOthLocBag = "Y" Then
'        wLocStr = " and CHARINDEX(BLoc,'" + mLoc + "') = 0) "
'        wStr = " Union All " + _
'               "SELECT 'B' as qTag, 0 as qTagNo, BPRtKey, BNo, BOdDmCd, BOdSfx, BOdDmSz, " + _
'               "BOdTc+'/'+BOdYy+'/'+BOdChr+'/'+LTrim(Str(BOdNo))+'/'+LTrim(Str(BOdSr)) " + _
'               "as qOrdNo, BLoc, BQty, BFgSubLoc from Bag " + _
'               "where " + _
'               IIF(gs_Partition = ctCurrPrtn Or mPrtKey = ctCurrPrtn, " BPrtKey='" + ctCurrPrtn + "' And ", "") + _
'               "BCoCd='" + mCoCd + "' and BYy= '" + mYy + "' and BChr='" + mChr + _
'               " " + wLocStr + wSameDsgBag
'      Else
'        wStr = " "
'      End If
      wSqlStrg = wSqlStrg + " Order By qTagNo, BNo Desc"
    Else
      If mRprYN = "Y" Then
      wSqlStrg = "SELECT BPrtKey, BNo, BOdDmCd, BOdSfx, BOdDmSz, " + _
                 "BOdTc+'/'+BOdYy+'/'+BOdChr+'/'+LTrim(Str(BOdNo))+'/'+LTrim(Str(BOdSr)) " + _
                 "as qOrdNo, BLoc, BQty, BFgSubLoc from Bag " + _
                 "where " + _
                 IIF(gs_Partition = ctCurrPrtn Or mPrtKey = ctCurrPrtn, " BPrtKey='" + ctCurrPrtn + "' And ", "") + _
                 "BCoCd='" + mCoCd + "' and BYy= '" + mYy + "' and BChr='" + mChr + _
                 "' and BRepYN = 'Y' Order By BNo Desc"
      Else
      wSqlStrg = "SELECT BPrtKey, BNo, BOdDmCd, BOdSfx, BOdDmSz, " + _
                 "BOdTc+'/'+BOdYy+'/'+BOdChr+'/'+LTrim(Str(BOdNo))+'/'+LTrim(Str(BOdSr)) " + _
                 "as qOrdNo, BLoc, BQty, BFgSubLoc from Bag " + _
                 "where " + _
                 IIF(gs_Partition = ctCurrPrtn Or mPrtKey = ctCurrPrtn, " BPrtKey='" + ctCurrPrtn + "' And ", "") + _
                 "BCoCd='" + mCoCd + "' and BYy= '" + mYy + "' and BChr='" + mChr + _
                 "' Order By BNo Desc"
      End If
    End If
  End If
  With mwhlp
    .Sqlstrg = wSqlStrg
    'uma *** Repair - help for Repair Bag
    If mRprYN = "Y" Then
    .ColHeadg = "|Bag Number|" + IIF(gs_Partition <> ctCurrPrtn, "Prtn |", "") + "Design           |Suffix                                   |Size |Loc    |Qty     |"
    .ColFlds = "|BNo|" + IIF(gs_Partition <> ctCurrPrtn, "BPrtKey|", "") + "BOdDmCd|BOdSfx|BOdDmSz|BRepLoc|BRepQty|"
    Else
    .ColHeadg = "|Bag Number|" + IIF(gs_Partition <> ctCurrPrtn, "Prtn |", "") + "Design           |Suffix                                   |Size |Order                                  |Loc    |FgSubLoc    |Qty     |"
    .ColFlds = "|BNo|" + IIF(gs_Partition <> ctCurrPrtn, "BPrtKey|", "") + "BOdDmCd|BOdSfx|BOdDmSz|qOrdNo|BLoc|BFgSubLoc|BQty|"
    End If
    .SearchColNum = 1
    .RetFldNm = "BNo"
    .SortAsc = False
  End With
End Sub
'*** Jay 2.14 *** [pb_ShowValidOnly added]
Public Sub Loc(ByVal mCoCd As Variant, Optional ByVal mLocTyp As Variant, Optional ByVal mNotInLocTyp As Variant, Optional ByVal mLocPrntYn As Variant, Optional ByVal pb_ShowValidOnly As Boolean = True _
, Optional ByVal mValidLocs As String)
  Dim wHead As String, wFlds As String
  'Debug.Print "in help"
  wCnd = ""
  If Not IsMissing(mCoCd) Then
    If Trim(mCoCd) <> "" Then
      wCnd = " where LocCoCd= '" + mCoCd + "' "
      wHead = "": wFlds = ""
    ElseIf Trim(mCoCd) = "" Then
      wHead = "Company |": wFlds = "LocCoCd|"
    End If
  End If
  
  If Not IsMissing(mLocTyp) Then
    If Trim(mLocTyp) <> "" Then wCnd = wCnd + IIF(wCnd = "", " where ", " And ") + " LocTyp in (" + mLocTyp + ") "
  End If
  If Not IsMissing(mNotInLocTyp) Then
    If Trim(mNotInLocTyp) <> "" Then wCnd = wCnd + IIF(wCnd = "", " where ", " And ") + " LocTyp not in (" + mNotInLocTyp + ") "
  End If
  
  '****** Sachin 2.14.0 27-10-2006 [Multi-Company] ******
  If Not IsMissing(mLocPrntYn) Then
    If Trim(mLocPrntYn) <> "" Then wCnd = wCnd + IIF(wCnd = "", " where ", " And ") + " LocPrntYN in (" + mLocPrntYn + ") "
  End If
  '****** Sachin 2.14.0 27-10-2006 [Multi-Company] ******
  '6.1B-16 valid loc added
  If Not IsMissing(mValidLocs) Then
    If Trim(mValidLocs) <> "" Then wCnd = wCnd + IIF(wCnd = "", " where ", " And ") + " LocCd in (" + mValidLocs + ") "
  End If
  '*** jay 2.14 *** ([and LocValidYN = 'Y'] added)
    If IsMissing(pb_ShowValidOnly) = True Then
        pb_ShowValidOnly = True
    Else
        pb_ShowValidOnly = False
    End If
  
    wCnd = wCnd + IIF(pb_ShowValidOnly = True, " and LocValidYN= 'Y' ", " ")
  '*** jay 2.14 *** ([and LocValidYN = 'Y'] added)
  With mwhlp
' ########################  Manoj #### Ver: 2.0.6 #### Date: 30/03/2004  ########################
    .Sqlstrg = "Select LocCoCd, LocTyp, LocCd, LocDesc, LocPrdStg, LocValidYN From Loc " + _
               wCnd + " Order by LocCoCd, LocTyp, LocCd "
    .ColHeadg = "|Code            |" + wHead + "Type|Desc                                                  |PrdStg     |Yn|"
    .ColFlds = "|LocCd|" + wFlds + "LocTyp|LocDesc|LocPrdStg|LocValidYN|"
' ###########################################  Manoj  ###########################################
    .RetFldNm = "LocCd"
    .SearchColNum = 1
    .SortAsc = True
  End With
End Sub
Public Sub DmSz(ByVal mDmTcTyp As String, ByVal mDmCd As String)
  With mwhlp
    '****** Sachin 3.02.0 - Partition Key added
    'indigo.164 DmTcTyp parameter value splitted and build where condition
    Dim wDmTcTyp As String, wDmTcTypArr() As String
    Dim i As Integer
    wDmTcTypArr = Split(mDmTcTyp, "','")
    For i = 0 To UBound(wDmTcTypArr)
      wDmTcTyp = wDmTcTyp + IIF(wDmTcTyp = "", "", " Or ") + "DmTcTyp='" + wDmTcTypArr(i) + "'"
    Next
    'indigo.164 DmTcTyp ='" + mDmTcTyp + "'" in where clause changed
    .Sqlstrg = "Select DmSz, DmDesc From DsgMst Where " + _
                IIF(gs_Partition = ctCurrPrtn, " DmPrtKey='" + ctCurrPrtn + "' And ", "") + _
               " DMCd ='" + mDmCd + "' and (" + wDmTcTyp + ") order by DmSz"
               
    If mDmTcTyp = "PM" Then
      .ColHeadg = "|Comp Size  |Desc                                                |" + IIF(gs_Partition <> ctCurrPrtn, "Prtn |", "")
    Else
      .ColHeadg = "|Size       |Desc                                                |" + IIF(gs_Partition <> ctCurrPrtn, "Prtn |", "")
    End If
    
    .ColFlds = "|DmSz|DmDesc|" + IIF(gs_Partition <> ctCurrPrtn, "DmPrtKey|", "")
    .RetFldNm = "DmSz"
    .SearchColNum = 1
    .SortAsc = True
  End With
End Sub
Public Sub HCd(ByVal mCoCd As String)
  With mwhlp
    .Sqlstrg = "SELECT HCd, HName from Head where HCoCd = '" + mCoCd + "' order by HCd "
    .ColHeadg = "|Code       |Name                             |"
    .ColFlds = "|HCd|HName|"
    .RetFldNm = "HCd"
    .SearchColNum = 1
    .SortAsc = True
  End With
End Sub
Public Sub hCoCd()
  With mwhlp
    .Sqlstrg = "SELECT HCoCd, HCd, HName from Head order by HCoCd "
    .ColHeadg = "|Company Code |Code       |Name                             |"
    .ColFlds = "|HCoCd|HCd|HName|"
    .RetFldNm = "HCoCd"
    .SearchColNum = 1
    .SortAsc = True
  End With
End Sub
Public Sub OdSr(ByVal mCoCd As String, ByVal mTc As String, ByVal mYy As String, ByVal mChr As String, ByVal mNo As Single, Optional ByVal mShowPrdQty As Boolean = False, Optional ByVal mPrtKey As String)
   '****** Sachin 3.02 24-12-07 - Partition Key added
   wSqlStrg = "SELECT " + IIF(gs_Partition = ctCurrPrtn, "", "OdPrtKey,") + "OdSr, OdDmCd, OdSfx, OdDmSz " + IIF(mShowPrdQty = True, ", OdPrdQty ", "") + ", OdPoNo " + _
             "From OrdDsg where OdCoCd= '" + mCoCd + "' and " + _
             "OdTc= '" + mTc + "' and OdYy='" + mYy + "' and " + _
             "OdChr='" + mChr + "' and OdNo = " + CStr(mNo) + _
             IIF(gs_Partition = ctCurrPrtn Or mPrtKey = ctCurrPrtn, " and OdPrtKey='" + ctCurrPrtn + "'", "") + _
             " Order By OdSr"
  With mwhlp
    .Sqlstrg = wSqlStrg
    .ColHeadg = "|Sr.         |Design                          |Suffix                                   |Size |SubPO               |" + IIF(gs_Partition <> ctCurrPrtn, "Prtn |", "") + IIF(mShowPrdQty = True, "Prd Qty   |", "")
    .ColFlds = "|OdSr|OdDmCd|OdSfx|OdDmSz|OdPONo|" + IIF(gs_Partition <> ctCurrPrtn, "OdPrtKey|", "") + IIF(mShowPrdQty = True, "OdPrdQty|", "")
    .ColMask = "||||||" + IIF(gs_Partition <> ctCurrPrtn, "|", "") + IIF(mShowPrdQty = True, "|", "")
    .SearchColNum = 1
    .RetFldNm = "OdSr"
    .SortAsc = True
  End With
End Sub
Public Sub OdPONo(ByVal mCoCd As String, ByVal mTc As String, ByVal mYy As String, ByVal mChr As String, ByVal mNo As Single, Optional ByVal mShowPrdQty As Boolean = False, Optional ByVal mPrtKey As String)
   '****** Sachin 3.02 24-12-07 - Partition Key added
   wSqlStrg = "SELECT " + IIF(gs_Partition = ctCurrPrtn, "", "OdPrtKey,") + "OdSr, OdDmCd, OdSfx, OdDmSz " + IIF(mShowPrdQty = True, ", OdPrdQty ", "") + ", OdPoNo " + _
             "From OrdDsg where OdCoCd= '" + mCoCd + "' and " + _
             "OdTc= '" + mTc + "' and OdYy='" + mYy + "' and " + _
             "OdChr='" + mChr + "' and OdNo = " + CStr(mNo) + _
             IIF(gs_Partition = ctCurrPrtn Or mPrtKey = ctCurrPrtn, " and OdPrtKey='" + ctCurrPrtn + "'", "") + _
             " Order By OdSr"
  With mwhlp
    .Sqlstrg = wSqlStrg
    .ColHeadg = "|Sr.     |Design                          |Suffix                                   |Size |SubPO               |" + IIF(gs_Partition <> ctCurrPrtn, "Prtn |", "") + IIF(mShowPrdQty = True, "Prd Qty   |", "")
    .ColFlds = "|OdSr|OdDmCd|OdSfx|OdDmSz|OdPONo|" + IIF(gs_Partition <> ctCurrPrtn, "OdPrtKey|", "") + IIF(mShowPrdQty = True, "OdPrdQty|", "")
    .ColMask = "||||||" + IIF(gs_Partition <> ctCurrPrtn, "|", "") + IIF(mShowPrdQty = True, "|", "")
    .SearchColNum = 1
    .RetFldNm = "OdPONo"
    .SortAsc = True
  End With
End Sub
Public Sub OdDmCd(ByVal mOdCoCd As String, ByVal mOdTc As String, ByVal mOdYy As String, _
                  ByVal mOdChr As String, ByVal mOdNo As Single)
  '****** Sachin 3.02 - Partition Key added ******
  With mwhlp
    .Sqlstrg = "SELECT " + IIF(gs_Partition = ctCurrPrtn, "", "OdPrtKey,") + "OdSr, OdDmCd, OdDmSz, DmCtg, DmPrdCtg, DmDesc from OrdDsg " + _
               "Join DsgMst on DmIdNo=OdDmIdNo and DmTcTyp= OdDmTcTyp and DmCd= OdDmCd and DmSz= '' " + _
               "where " + _
               IIF(gs_Partition = ctCurrPrtn, " OdPrtKey='" + ctCurrPrtn + "' And ", "") + _
               " OdCoCd = '" + mOdCoCd + "' and " + _
               "OdTc= '" + mOdTc + "' and OdYy='" + mOdYy + "' and " + _
               "OdChr= '" + mOdChr + "' and OdNo= " + CStr(mOdNo) + " order by OdDmCd, OdDmSz "
    .ColHeadg = "|Sr   |" + IIF(gs_Partition <> ctCurrPrtn, "Prtn |", "") + "Design           |Size |Dsg Ctg|Prd Ctg|Description                          |"
    .ColFlds = "|OdSr|" + IIF(gs_Partition <> ctCurrPrtn, "OdPrtKey", "") + "OdDmCd|OdDmSz|DmCtg|DmPrdCtg|DmDesc|"
    .RetFldNm = "OdDmCd"
    .SearchColNum = 1
    .SortAsc = True
  End With
End Sub
Public Sub OrRmCd(ByVal mOrCoCd As String, ByVal mOrTc As String, ByVal mOrYy As String, _
                  ByVal mOrChr As String, ByVal mOrNo As Single, ByVal mOrSr As Single, _
                  Optional ByVal mRmCtg As Variant, Optional ByVal mRmSCtg As Variant)
  Dim wCnd As String
  '****** Sachin 3.02.0 - Partition Key added
  With mwhlp
    If Not IsMissing(mRmCtg) Then wCnd = " and RmCtg in (" + mRmCtg + ") "
    If Not IsMissing(mRmSCtg) Then _
       wCnd = wCnd + " and  RmSCtg in (" + mRmSCtg + ") "
    
    'Manoj 2.10.0
'    .Sqlstrg = "Select OrRmCd, max(OrRmCtg) as OrRmCtg, max(OrRmSCtg) as OrRmSCtg, " + _
'               "sum(OrPrdQty) as OrQty, sum(OrPrdWt) as OrWt from OrdRm " + _
'               "where OrCoCd='" + mOrCoCd + "' and " + _
'               "OrTc='" + mOrTc + "' and OrYy='" + mOrYy + "' and " + _
'               "OrChr='" + mOrChr + "' and OrNo=" + CStr(mOrNo) + " and " + _
'               "OrSr=" + CStr(mOrSr) + " Group By OrRmCd Order by OrRmCd "

    .Sqlstrg = "Select OrRmCd, max(OrRmCtg) as OrRmCtg, max(OrRmSCtg) as OrRmSCtg, " + _
               "sum(OrPrdQty) as OrQty, sum(OrPrdWt) as OrWt from OrdRm " + _
               "Join RmMst On OrRmCd = RmCd " + IIF(gs_Partition = ctCurrPrtn, " And OrPrtKey=RmPrtKey ", "") + " where " + _
                IIF(gs_Partition = ctCurrPrtn, " OrPrtKey='" + ctCurrPrtn + "' and ", "") + _
               "OrCoCd='" + mOrCoCd + "' and " + _
               "OrTc='" + mOrTc + "' and OrYy='" + mOrYy + "' and " + _
               "OrChr='" + mOrChr + "' and OrNo=" + CStr(mOrNo) + " and " + _
               "OrSr=" + CStr(mOrSr) + " " + wCnd + " Group By OrRmCd Order by OrRmCd "
    'Manoj 2.10.0
    .ColHeadg = "|Code                      |Ctg |Sub Ctg |Qty    |Wt            |"
    .ColFlds = "|OrRmCd|OrRmCtg|OrRmSCtg|OrQty|OrWt|"
    .RetFldNm = "OrRmCd"
    .SearchColNum = 1
    .SortAsc = True
  End With
End Sub
Public Sub OrRmCdForIssueRm(ByVal mOrCoCd As String, ByVal mOrTc As String, ByVal mOrYy As String, _
                  ByVal mOrChr As String, ByVal mOrNo As Single, ByVal mOrSr As Single, _
                  Optional ByVal mRmCtg As Variant, Optional ByVal mRmSCtg As Variant, Optional mBChr As String)
  Dim wCnd As String
  Dim wRsBagChr As MDORowSet
  Dim wAddCnd As String
  Dim wSubShp As String
  Dim wPrdTyp As String
  '****** Sachin 3.02.0 - Partition Key added
  If mBChr <> "" Then
    Set wRsBagChr = go_HlpCn.OpenRes("Select * from Param where pTyp='BCHR' and pMCd= '" + mBChr + "'")
    If wRsBagChr.RecCount > 0 Then
      wSubShp = wRsBagChr!PDesc225
      wPrdTyp = wRsBagChr!PValue3
      wAddCnd = wAddCnd + IIF(wRsBagChr!PDesc225 <> "", " and OrSubShp = '" + wRsBagChr!PDesc225 + "'", "") + _
      " " + IIF(wRsBagChr!PValue3 = "WS", " and OrWsQty>0 ", IIF(wRsBagChr!PValue3 = "GS", " and OrHsQty > 0", "")) + _
      " " + IIF(wRsBagChr!PValue1 <> "", " and charindex(','+OrRmCtg+',',','+ '" + wRsBagChr!PValue1 + "'+',',0)>0", "") + _
      " " + IIF(wRsBagChr!PValue2 <> "", " and charindex(','+OrRmSCtg+',',','+'" + wRsBagChr!PValue2 + "'+',',0)>0", "")
    End If
  End If
  
  
  With mwhlp
    If Not IsMissing(mRmCtg) Then wCnd = " and RmCtg in (" + mRmCtg + ") "
    If Not IsMissing(mRmSCtg) Then _
       wCnd = wCnd + " and  RmSCtg in (" + mRmSCtg + ") "
    
    'Manoj 2.10.0
'    .Sqlstrg = "Select OrRmCd, max(OrRmCtg) as OrRmCtg, max(OrRmSCtg) as OrRmSCtg, " + _
'               "sum(OrPrdQty) as OrQty, sum(OrPrdWt) as OrWt from OrdRm " + _
'               "where OrCoCd='" + mOrCoCd + "' and " + _
'               "OrTc='" + mOrTc + "' and OrYy='" + mOrYy + "' and " + _
'               "OrChr='" + mOrChr + "' and OrNo=" + CStr(mOrNo) + " and " + _
'               "OrSr=" + CStr(mOrSr) + " Group By OrRmCd Order by OrRmCd "

    .Sqlstrg = "Select OrRmCd, max(OrRmCtg) as OrRmCtg, max(OrRmSCtg) as OrRmSCtg, " + _
               "sum(OrPrdQty) as OrQty, sum(OrPrdWt) as OrWt,'" + wSubShp + "' as wOrSubShp,'" + wPrdTyp + "' as wPrdTyp from OrdRm " + _
               "Join RmMst On OrRmCd = RmCd " + IIF(gs_Partition = ctCurrPrtn, " And OrPrtKey=RmPrtKey ", "") + _
               " where " + _
                IIF(gs_Partition = ctCurrPrtn, " OrPrtKey='" + ctCurrPrtn + "' and ", "") + _
               "OrCoCd='" + mOrCoCd + "' and " + _
               "OrTc='" + mOrTc + "' and OrYy='" + mOrYy + "' and " + _
               "OrChr='" + mOrChr + "' and OrNo=" + CStr(mOrNo) + " and " + _
               "OrSr=" + CStr(mOrSr) + " " + wCnd + wAddCnd + " Group By OrRmCd Order by OrRmCd "
    'Manoj 2.10.0
    .ColHeadg = "|Code                      |Ctg |Sub Ctg |SubShp|PrdTyp|Qty    |Wt            |"
    .ColFlds = "|OrRmCd|OrRmCtg|OrRmSCtg|wOrSubShp|wPrdTyp|OrQty|OrWt|"
    .RetFldNm = "OrRmCd"
    .SearchColNum = 1
    .SortAsc = True
  End With
End Sub

'/* urmila help of BRm_Sz/ln
Public Sub BRmSz(ByVal mBCoCd As String, ByVal mBYy As String, ByVal mBChr As String, _
                  ByVal mBNo As Single, ByVal mRmCd As String, ByVal mSeoYn As Boolean)
  
  Dim wTemp As String
  ' Zubin 212 (TdLotNo included only for viewing)
  ' ***** Manali 3.03 - 19/07/08 - included RmSz Desc for viewing
  Dim ws_SqlFlds As String, ws_GrpBy As String, ws_ColHeadg As String, ws_ColFlds As String
  'Dim moCn As MwfLib.MDOConnection   '*** (Bef speed) bef 24/11/06

  '****** Sachin 3.02.0 - Partition Key added
  If ImplementLotNo(mBCoCd) Then _
    ws_SqlFld = IIF(gs_Partition = ctCurrPrtn, "", "TdPrtKey,") + "TdLotNo, ": ws_GrpBy = IIF(gs_Partition = ctCurrPrtn, "TdPrtKey,", "") + "TdLotNo, ": ws_ColHeadg = IIF(gs_Partition <> ctCurrPrtn, "|Prtn ", "") + "|LotNo                   ": ws_ColFlds = IIF(gs_Partition <> ctCurrPrtn, "|TdPrtKey", "") + "|TdLotNo"
  
  With mwhlp
    ' ***** Manali 3.03 - 19/07/08 - RmSz Desc added, join with RmMst and RmIdSz Added
    wTemp = "Select " + ws_SqlFld + " TdRmSz, " + _
                "IsNull(max(RisRmSzDesc), '') as qRmSzDesc, " + _
                "TdRmStkRt, sum((case (case when TdFrRmLoc='W' then TdFrRmDc " + _
                "when TdToRmLoc='W' then TdToRmDc else '' end) when 'C' then -1 " + _
                "when 'D' then +1 else 0 end)* TdRmQty) as qRmQty, " + _
                "sum((case (case when TdFrRmLoc='W' then TdFrRmDc when TdToRmLoc='W' then " + _
                "TdToRmDc end) when 'C' then -1 when 'D' then +1 else 0 end)* " + _
                "TdRmWt) as qRmWt From Txnd " + _
                "Join RmMst on RmCd=TdRmCd Left Outer Join RmIdSz on RisRmCtg=TdRmCtg and RisRmId=RmId and RisSz1=TdRmSz " + _
                "Where TdRmCd <> '' And " + _
                IIF(gs_Partition = ctCurrPrtn, " TdPrtKey='" + ctCurrPrtn + "' And ", "") + _
                " TdCoCd= '" + mBCoCd + "' and " + _
                " TdBYy= '" + mBYy + "' and " + _
                " TdBChr= '" + mBChr + "' and " + _
                " TdBNo= " + CStr(mBNo) + " and " + _
                " TdRmCd= '" + mRmCd + "' " + _
                " group By TdBYy, TdBChr, TdBNo, TdRmCd, " + ws_GrpBy + " TdRmSz, TdRmStkRt having " + _
                "sum((case (case when TdFrRmLoc='W' then TdFrRmDc when TdToRmLoc='W' then " + _
                "TdToRmDc end) when 'C' then -1 when 'D' then +1 else 0 end)* " + _
                "TdRmWt) <> 0"
                
'StkRt-if Seo rights is true stock rate will display else not show stock rate
    .Sqlstrg = wTemp
    .ColHeadg = "|Ln/Sz/Ptr   |RmSz Desc    " + ws_ColHeadg + IIF(mSeoYn = True, "|StkRt   ", "") + "|Qty    |Wt            |"
    .ColFlds = "|TdRmSz|qRmSzDesc" + ws_ColFlds + IIF(mSeoYn = True, "|TdRMStkRt", "") + "|qRmQty|qRmWt|"
    .RetFldNm = "TdRmSz"
    .SearchColNum = 1
    .SortAsc = True
  End With
End Sub


Public Sub FgNo(ByVal mCoCd As String, ByVal mTc As String, ByVal mYy As String, ByVal mChr As String)
  Dim wPurgeDt As String
  wPurgeDt = GetPurgeDt(mCoCd, "")  'sv.9 getting purge date
  '****** Sachin 3.02.0 - Partition Key added
  wSqlStrg = "SELECT " + IIF(gs_Partition = ctCurrPrtn, "", "FgPrtKey,") + " FgNo,FgDt from Fg where " + _
             IIF(gs_Partition = ctCurrPrtn, " FgPrtKey='" + ctCurrPrtn + "' And ", "") + _
             " FgCoCd= '" + mCoCd + "' " + _
             "and FgTc= '" + mTc + "' and FgYy='" + mYy + "' and " + _
             "FgChr='" + mChr + "' " + _
             IIF(gs_Partition = ctCurrPrtn, " And FgPrtKey= '" + ctCurrPrtn + "'", "") + _
             IIF(gs_Partition = ctCurrPrtn, " And Fgdt>'" + wPurgeDt + "' ", "") + _
             " Order By FgNo Desc"
  With mwhlp
    .Sqlstrg = wSqlStrg
    .ColHeadg = "|Number   |Date        |" + IIF(gs_Partition <> ctCurrPrtn, "Prtn |", "")
    .ColFlds = "|FgNo|FgDt|" + IIF(gs_Partition <> ctCurrPrtn, "FgPrtKey|", "")
    .ColMask = "||" + DtShortStr() + "|" + IIF(gs_Partition <> ctCurrPrtn, "|", "")   '6.1
    .SearchColNum = 1
    .RetFldNm = "FgNo"
    .SortAsc = False
  End With
End Sub
Public Sub FmNo(ByVal mCoCd As String, ByVal mTc As String, ByVal mYy As String, ByVal mChr As String)

  '****** Sachin 3.02.0 - Partition Key added
  wSqlStrg = "SELECT " + IIF(gs_Partition = ctCurrPrtn, "", "FmPrtKey,") + "FmNo,FmDt from Fm Where " + _
             IIF(gs_Partition = ctCurrPrtn, " FmPrtKey='" + ctCurrPrtn + "' And ", "") + _
             " FmCoCd= '" + mCoCd + "' " + _
             " and FmTc= '" + mTc + "' and FmYy='" + mYy + "' and " + _
             " FmChr='" + mChr + "' Order By FmNo Desc"
  With mwhlp
    .Sqlstrg = wSqlStrg
    .ColHeadg = IIF(gs_Partition <> ctCurrPrtn, "|Prtn ", "") + "|Number   |Date        |"
    .ColFlds = IIF(gs_Partition <> ctCurrPrtn, "|FmPrtKey", "") + "|FmNo|FmDt|"
    .ColMask = IIF(gs_Partition <> ctCurrPrtn, "|", "") + "||" + DtShortStr() + "|"   '6.1
    .SearchColNum = IIF(gs_Partition <> ctCurrPrtn, 2, 1)
    .RetFldNm = "FmNo"
    .SortAsc = False
  End With
End Sub
Public Sub FmdSr(ByVal mCoCd As String, ByVal mTc As String, ByVal mYy As String, ByVal mChr As String, ByVal mNo As Single)

  wSqlStrg = "SELECT " + IIF(gs_Partition = ctCurrPrtn, "", "FmPrtKey,") + " FmdSr from Fmd where " + _
            IIF(gs_Partition = ctCurrPrtn, " FmdPrtKey='" + ctCurrPrtn + "' And ", "") + _
             " FmdCoCd= '" + mCoCd + "' " + _
             " and FmdTc= '" + mTc + "' and FmdYy='" + mYy + "' and " + _
             " FmdChr='" + mChr + "' and FmdNo = " + CStr(mNo) + "  Order By FmdSr"
  With mwhlp
    .Sqlstrg = wSqlStrg
    .ColHeadg = "|Sr.    |" + IIF(gs_Partition <> ctCurrPrtn, "Prtn |", "")
    .ColFlds = "|FmdSr|" + IIF(gs_Partition <> ctCurrPrtn, "FmdPrtKey|", "")
    .ColMask = "||" + IIF(gs_Partition <> ctCurrPrtn, "|", "")
    .SearchColNum = 1
    .RetFldNm = "FmdSr"
    .SortAsc = True
  End With
End Sub

Public Sub TdSr(ByVal mCoCd As String, ByVal mTc As String, ByVal mYy As String, ByVal mChr As String, ByVal mNo As Single)
    If mTc = "FM" Then Call FmdSr(mCoCd, mTc, mYy, mChr, mNo): Exit Sub

  '****** Sachin 3.02.0 - Partition Key added
  wSqlStrg = "SELECT " + IIF(gs_Partition = ctCurrPrtn, "", "TdPrtKey,") + " TdSr from Txnd where " + _
             IIF(gs_Partition = ctCurrPrtn, " TdPrtKey='" + ctCurrPrtn + "' And ", "") + _
             " TdCoCd= '" + mCoCd + "' and " + _
             "TdTc= '" + mTc + "' and TdYy='" + mYy + "' and " + _
             "TdChr='" + mChr + "' and TdNo = " + CStr(mNo) + " and TdSrNo= 0 " + _
             "Order By TdSr"
  With mwhlp
    .Sqlstrg = wSqlStrg
    .ColHeadg = "|Sr.    |" + IIF(gs_Partition <> ctCurrPrtn, "Prtn |", "")
    .ColFlds = "|TdSr|" + IIF(gs_Partition <> ctCurrPrtn, "TdPrtKey|", "")
    .ColMask = "||" + IIF(gs_Partition <> ctCurrPrtn, "|", "")
    .SearchColNum = 1
    .RetFldNm = "TdSr"
    .SortAsc = True
  End With
End Sub
Public Sub FdSr(ByVal mCoCd As String, ByVal mTc As String, ByVal mYy As String, ByVal mChr As String, ByVal mNo As Single)
  '****** Sachin 3.02.0 - Partition Key added
  wSqlStrg = "SELECT " + IIF(gs_Partition = ctCurrPrtn, "", "FdPrtKey,") + " FdSr from Fgd " + _
             "where " + _
             IIF(gs_Partition = ctCurrPrtn, " FdPrtKey='" + ctCurrPrtn + "' And ", "") + _
             " FdCoCd= '" + mCoCd + "' and FdTc= '" + mTc + "' and " + _
             "FdYy='" + mYy + "' and FdChr='" + mChr + "' and " + _
             "FdNo = " + CStr(mNo) + " Order By FdSr"
  With mwhlp
    .Sqlstrg = wSqlStrg
    .ColHeadg = "|Sr.    |" + IIF(gs_Partition <> ctCurrPrtn, "Prtn |", "")
    .ColFlds = "|FdSr|" + IIF(gs_Partition <> ctCurrPrtn, "TdPrtKey|", "")
    .ColMask = "||" + IIF(gs_Partition <> ctCurrPrtn, "|", "")
    .SearchColNum = 1
    .RetFldNm = "FdSr"
    .SortAsc = True
  End With
End Sub
Public Sub InNo(ByVal mCoCd As String, ByVal mTc As String, ByVal mYy As String, ByVal mChr As String, Optional ByVal mCmCd As Variant, Optional ByVal mSkipPrgDt As Boolean = False)
  '*** (Jen 2.13)
  Dim wCnd As String
  If Not IsMissing(mCmCd) Then
    wCnd = " and InCmCd= '" + mCmCd + "' "
  Else
    wCnd = " "
  End If
  '*** (Jen 2.13)
  Dim wPurgeDt As String
  wPurgeDt = GetPurgeDt(mCoCd, "")  'sv.9 getting purge date
  '*** (Bef 2.13)
  'wSqlStrg = "SELECT InNo, InDt, InCmCd, InExpNo from InvHd where InCoCd= '" + mCoCd + "' " + _
  '           "and InTc= '" + mTc + "' and InYy= '" + mYy + "' and " + _
  '           "InChr= '" + mChr + "' " + wCnd + " Order By InNo"
  '*** (Bef 2.13)
  
  '*** (Jen 2.13)
  '****** Sachin 3.02 - Partition Key added
  '*** Jay 3.3(FG) [Case modified for FG]
  ' ***** Manali Trading Module - [Case added for 'JT']
  wSqlStrg = "SELECT " + IIF(gs_Partition = ctCurrPrtn, "", "InPrtKey,") + " InNo, InDt, InCmCd, InExpNo, CmCurCd from InvHd join Param on PTyp= 'TC' and PMCd= InTc and PSCd= '' " + _
             "Join CustMst on CmCtg= (case when PValue in ('FG','IN') then 'C' when PValue in ('CTB') then 'T' " + _
             "                        When PValue In ('JT') Then (Select OmCmCtg from OrdMst Where OmYy=InYy and OmCoCd=InCoCd and OmTc=InTc and OmChr=InChr and OmNo=InNo and OmPrtKey=InPrtKey) else '' end) and CmCd= InCmCd " + _
             " Where " + _
             IIF(gs_Partition = ctCurrPrtn, " InPrtKey= '" + ctCurrPrtn + "' and ", "") + _
             IIF(gs_Partition = ctCurrPrtn And mSkipPrgDt = False, " Indt>'" + wPurgeDt + "' And ", "") + _
             " InCoCd= '" + mCoCd + "' " + _
             "and InTc= '" + mTc + "' and InYy= '" + mYy + "' and " + _
             "InChr= '" + mChr + "' " + wCnd + " Order By InNo Desc"
  '*** (Jen 2.13)
           
  With mwhlp
    .Sqlstrg = wSqlStrg
    .ColHeadg = "|Number   |Date        |Customer               |Inv Exp No                               |Currency   |" + IIF(gs_Partition <> ctCurrPrtn, "Prtn |", "")
    .ColFlds = "|InNo|InDt|InCmCd|InExpNo|CmCurCd|" + IIF(gs_Partition <> ctCurrPrtn, "InPrtKey|", "")
    .ColMask = "||" + DtShortStr() + "||||" + IIF(gs_Partition <> ctCurrPrtn, "|", "")  '6.1
    .SearchColNum = 1
    .RetFldNm = "InNo"
    .SortAsc = False
  End With
End Sub
Public Sub InSr(ByVal mCoCd As String, ByVal mTc As String, ByVal mYy As String, ByVal mChr As String, ByVal mNo As Single)
  '****** Sachin 3.02.0 - Partition Key added
  wSqlStrg = "SELECT " + IIF(gs_Partition = ctCurrPrtn, "", "IdPrtKey,") + " IdSr, IdDmCd, IdSfx, IdDmSz from InvDsg " + _
             "where " + _
             IIF(gs_Partition = ctCurrPrtn, " IdPrtKey='" + ctCurrPrtn + "' and ", "") + _
             " IdCoCd= '" + mCoCd + "' and IdTc= '" + mTc + "' and " + _
             "IdYy='" + mYy + "' and IdChr='" + mChr + "' and IdNo = " + CStr(mNo) + _
             " Order By IdSr"
  With mwhlp
    .Sqlstrg = wSqlStrg
    .ColHeadg = IIF(gs_Partition <> ctCurrPrtn, "|Prtn ", "") + "|Sr.  |Design                          |Suffix                                   |Size |"
    .ColFlds = IIF(gs_Partition <> ctCurrPrtn, "|IdPrtKey", "") + "|IdSr|IdDmCd|IdSfx|IdDmSz|"
    .ColMask = IIF(gs_Partition <> ctCurrPrtn, "|", "") + "|||||"
    .SearchColNum = IIF(gs_Partition <> ctCurrPrtn, 2, 1)
    .RetFldNm = "IdSr"
    .SortAsc = True
  End With
End Sub
'*** Jay 3.3(FG)
Public Sub InExpNo(ByVal mCoCd As String, Optional mInTcTyp As String = "IN")
'*** Jay 3.3(FG)
  '*** (Bef 2.13)
  'wSqlStrg = "SELECT InExpNo, InTc, InYy, InChr, InNo, InDt ,InCmCd, CmCurCd from InvHd " + _
  '           "Where InCoCd= '" + mCoCd + "' Order By InExpNo"
  '*** (Bef 2.13)

  '*** Jay 3.3(FG)
  '*** (Jen 2.13)
  '****** Partition Key Not added as Export Number Is Unique Across the Database
  wSqlStrg = "SELECT InPrtKey, InExpNo, InTc, InYy, InChr, InNo, InDt, InCmCd, CmCurCd, InCnvRt from InvHd " + _
             "join Param on PTyp= 'TC' and PMCd= InTc and PSCd= '' " + _
             "join CustMst on CmCtg= (case when PValue in ('FG','IN') then 'C' when PValue in ('CTB') then 'T' else '' end) and CmCd= InCmCd " + _
             "Where " + _
             IIF(gs_Partition = ctCurrPrtn, " InPrtKey='" + ctCurrPrtn + "' and ", "") + _
             " InCoCd= '" + mCoCd + "' " + wCnd + " Order By InExpNo"
  '*** (Jen 2.13)
  '*** Jay 3.3(FG)

  With mwhlp
    .Sqlstrg = wSqlStrg
    .ColHeadg = IIF(gs_Partition <> ctCurrPrtn, "|Prtn ", "") + "|Inv Export Number          |Tc     |Year|Chr    |Number|Date        |Customer   |Currency     |Cnv Rt         |"
    .ColFlds = IIF(gs_Partition <> ctCurrPrtn, "|InPrtKey", "") + "|InExpNo|InTc|InYy|InChr|InNo|InDt|InCmCd|CmCurCd|InCnvRt|"
    .ColMask = IIF(gs_Partition <> ctCurrPrtn, "|", "") + "||||||" + DtShortStr() + "|||####0.0000|"    '6.1
    .SearchColNum = IIF(gs_Partition <> ctCurrPrtn, 2, 1)
    .RetFldNm = "InExpNo"
    .SortAsc = True
  End With
End Sub
Public Sub OdDmStkNo(ByVal ps_cocd As String, ByVal mOdDmCd As String, ByVal mOdSfx As String, ByVal mOdDmSz As String, ByVal mOdDmStkCoCd As String, ByVal mOdDmStkYy As String, ByVal mOdDmStkChr As String, Optional ByVal mOmTC As String)
  With mwhlp
  If mOmTC = "JRI" Then
    .Sqlstrg = "Select Byy,BChr,BNo,BGrWt from Bag where " + IIF(gs_Partition = ctCurrPrtn, " BPrtKey= '" + ctCurrPrtn + "' And ", "") + _
               "BCoCd = '" + mOdDmStkCoCd + "' and BOdDmCd='" + mOdDmCd + "' and BOdSfx='" + mOdSfx + "' and BOdDmSz='" + mOdDmSz + _
               "' and BYy = '" + mOdDmStkYy + "' and BChr ='" + mOdDmStkChr + " ' and BCls='Y' and BLoc='PFG' and BRepYN='N'" + _
               " order by BYy, BChr, BNo Desc "
    .ColHeadg = "|Yy |Chr  |DmStkNo     |Grs Wt    |"
    .ColFlds = "|BYy|BChr|BNo|BGrWt|"
    .RetFldNm = "BNo"
    .SearchColNum = 1
    .SortAsc = True
  ElseIf mOmTC = "JRO" Then
    .Sqlstrg = "Select Byy,BChr,BNo,BGrWt from Bag where " + IIF(gs_Partition = ctCurrPrtn, " BPrtKey= '" + ctCurrPrtn + "' And ", "") + _
               "BCoCd = '" + mOdDmStkCoCd + "' and BOdDmCd='" + mOdDmCd + "' and BOdSfx='" + mOdSfx + "' and BOdDmSz='" + mOdDmSz + _
               "' and BYy = '" + mOdDmStkYy + "' and BChr ='" + mOdDmStkChr + " ' and BRepYN='Y' " + _
               " order by BYy, BChr, BNo Desc "
    .ColHeadg = "|Yy |Chr  |DmStkNo     |Grs Wt    |"
    .ColFlds = "|BYy|BChr|BNo|BGrWt|"
    .RetFldNm = "BNo"
    .SearchColNum = 1
    .SortAsc = True
  Else
  'sritmp
     .Sqlstrg = "Select Distinct OdDmStkCoCd,OdDmStkYy, OdDmStkChr, OdDmStkNo,  " + _
                "IsNull((Select Top 1 Sum(OrWt/ (Case When OrRmCtg In ('D', 'C') Then 5 Else 1 End)) " + _
                "From OrdRm " + _
                "Join OrdDsg Od1 On Od1.OdCoCd=OrCoCd and Od1.OdTc=OrTc and Od1.OdYy=OrYy and Od1.OdChr=OrChr  and Od1.OdNo=OrNo and Od1.OdSr=OrSr " + _
                "Where Od1.OdCoCd='" + ps_cocd + "' and Od1.OdDmCd='" + mOdDmCd + "' and Od1.OdSfx='" + mOdSfx + "' and Od1.OdDmSz='" + mOdDmSz + "' and Od1.OdDmStkNo=Od2.OdDmStkNo and Od1.OdDmStkYy =  Od2.OdDmStkYy and Od1.OdDmStkChr = Od2.OdDmStkChr and Od1.OdDmStkCoCd = Od2.OdDmStkCoCd " + _
                "and Od1.OdTc In (Select PMCd from Param where Ptyp= 'TC' and PValue = 'JT') and Od1.OdPrtKey='C' " + _
                "Group By Od1.OdIdNo " + _
                "Order By max(Od1.OdOmDt) Desc, Od1.OdIdNo Desc), 0) as qGrsWt " + _
                "from OrdDsg Od2 where " + _
                IIF(gs_Partition = ctCurrPrtn, " OdPrtKey= '" + ctCurrPrtn + "' And ", "") + _
                " OdDmCd='" + mOdDmCd + "' and OdSfx='" + mOdSfx + "' and OdDmSz='" + mOdDmSz + "' and OdDmStkCoCd = '" + mOdDmStkCoCd + "' and OdDmStkYy = '" + mOdDmStkYy + "' and OdDmStkChr ='" + mOdDmStkChr + " '" + _
                " order by OdDmStkCoCd,OdDmStkYy, OdDmStkChr, OdDmStkNo Desc "
             
    .ColHeadg = "|StkCoCd|Yy |Chr  |DmStkNo     |Grs Wt    |"
    .ColFlds = "|OdDmStkCoCd|OdDmStkYy|OdDmStkChr|OdDmStkNo|qGrsWt|"
    .RetFldNm = "OdDmStkNo"
    .SearchColNum = 1
    .SortAsc = True
  End If
  End With
End Sub
'*** Jay 3.3(FG)

Public Sub PlNo(ByVal mCoCd As String, ByVal mYy As String, ByVal mChr As String)
  '****** Sachin 3.02.0  - Partition Key added
  wSqlStrg = "SELECT distinct FdPlNo from Fgd " + _
             "where " + _
             IIF(gs_Partition = ctCurrPrtn, " FdPrtKey='" + ctCurrPrtn + "' And ", "") + _
             " FdCoCd='" + mCoCd + "' and FdPlYy='" + mYy + "' and " + _
             "FdPlChr='" + mChr + "' Order By FdPlNo Desc"
  With mwhlp
    .Sqlstrg = wSqlStrg
    .ColHeadg = "|Pck List Number|"
    .ColFlds = "|FdPlNo|"
    .ColMask = "||"
    .SearchColNum = 1
    .RetFldNm = "FdPlNo"
    .SortAsc = False
  End With
End Sub
Public Sub BaseCustmPurNo(ByVal mCoCd As String, ByVal mRmZCd As String, ByVal mSrt As String, ByVal mCnvFct As Double, ByVal mMulDiv As String, Optional ByVal mPurRtFr As Double = 0#, Optional ByVal mPurRtTo As Double = 0#)
  '*** (Jen 2.13 speed) , Optional ByVal mPurRtTo As Double = 0#
  '*** iF HSrtBasePurVch= 'VCH' then Sort Help By 'TzTc, TzYy, TzChr, TzNo, TzSr'
  '*** iF HSrtBasePurVch= 'RTVCH' then Sort Help By 'TzPurRt, TzTc, TzYy, TzChr, TzNo, TzSr'
  '*** iF HSrtBasePurVch= 'CDRTVCH' then Sort Help By 'TzRmCdDesc, TzPurRt, TzTc, TzYy, TzChr, TzNo, TzSr'
  
  Dim wSRt As String
  
  '****** Sachin 3.02.0 - Partition Key added
  '*** (Jen 2.13 speed)
  Dim ws_Rt As String, ws_RtRange As String
  ws_Rt = " convert (decimal (16, 3), TzPurRt* (case when TMulDiv= 'D' then TCnvRt else 1 end)/ (case when TMulDiv= 'M' then TCnvRt else 1 end) " + _
             "* " + IIF(mMulDiv = "M", CStr(mCnvFct), CStr(1)) + _
             "/ " + IIF(mMulDiv = "D", CStr(mCnvFct), CStr(1)) + ") "
  ws_RtRange = " "
  If Round(mPurRtFr, 2) > 0 Then
    ws_RtRange = " and " + ws_Rt + " >= convert (decimal (16, 3), " + CStr(mPurRtFr) + ") "
  End If
  If Round(mPurRtTo, 2) > 0 Then
    ws_RtRange = ws_RtRange + " and " + ws_Rt + " <= convert (decimal (16, 3), " + CStr(mPurRtTo) + ") "
  End If
  '*** (Jen 2.13 speed)
  
  wSRt = " Order By " + mSrt + " "
  'wSRt = Replace(wSRt, "TzPurRt", "(case when TUSDYn='N' then Round(TzPurRt/TCnvRt, 3) else Round(TzPurRt, 3) end) ")
  'wSRt = Replace(wSRt, "TzPurRt", "Round(TzPurRt, 3)")   '*** (Bef 2.13)
  wSRt = Replace(wSRt, "TzPurRt", "convert (decimal (16, 3), TzPurRt)")
  'wSqlStrg = "SELECT TzTc+'/'+TzYy+'/'+TzChr+'/'+LTrim(Str(TzNo))+'/'+LTrim(Str(TzSr)) as qVch, " + _
             "TBOENo, TBOEDt, TUseBefDt, Round((TzRmWt- TzRmDespWt), 3) as qBalWt, TzPcPerCt, Round(TzPurRt, 2) As TzPurRt, TzRmCdDesc, " + _
             "TzYy, TzKey from TxndZ " + _
             "join Txn on TCoCd= TzCoCd and TTc= TzTc and TYy= TzYy and TChr= TzChr and TNo= TzNo " + _
             "where TzCoCd= '" + mCoCd + "' and TzBM= 'B' and TzCls= 'N' and TzRmZCd= '" + mRmZCd + "' and " + _
             "TzRmWt> TzRmDespWt " + wSRt
               
  '*** (Bef 2.13)
  ''*************************Geeta**************Emr207**************
  'wSqlStrg = "SELECT TzTc+ '/'+ TzYy+ '/'+ TzChr+ '/'+ LTrim(Str(TzNo))+ '/'+ LTrim(Str(TzSr)) as qVch, " + _
  '           "TBOENo, TBOEDt, TUseBefDt, Round((TzRmWt- TzRmDespWt), 3) as qBalWt, TzPcPerCt, " + _
  '           "Round(TzPurRt, 3) As TzPurRt, TzRmCdDesc, " + _
  '           "TzYy, TzKey, TzDesc from TxndZ " + _
  '           "join Txn on TCoCd= TzCoCd and TTc= TzTc and TYy= TzYy and TChr= TzChr and TNo= TzNo " + _
  '           "where TzCoCd= '" + mCoCd + "' and TzBM= 'B' and TzCls= 'N' and TzRmZCd= '" + mRmZCd + "' and " + _
  '           "TzRmWt> TzRmDespWt " + " and Round(TzPurRt, 3)>=" + CStr(mPurRtFr) + wSRt
  ''*************************Geeta**************Emr207**************
  '*** (Bef 2.13)
  
  '*** (Jen 2.13)
  '****** Sachin 3.02 - Id fields In Joins and Where Clause ******
  '*** 3.11.2
  mMulDiv = UCase(Trim(mMulDiv))
  wSqlStrg = "SELECT (case when (TCmCd<> '' or TzOdNo<> 0 or TzDmCd<> '' or TzRmCd<> '' or TzRmSzFr<> 0 " + _
             "         or TzRmSzTo<> 0) then 'Y' else 'N' end) as qBestUse, " + _
             "TzTc+ '/'+ TzYy+ '/'+ TzChr+ '/'+ LTrim(Str(TzNo))+ '/'+ LTrim(Str(TzSr)) as qVch, " + _
             "TBOENo, TBOEDt, TBillNo, TUseBefDt, convert (decimal (16, 3), (TzRmWt- TzRmDespWt)) as qBalWt, TzPcPerCt, " + _
             "TzPurRt, TCurrCd, " + _
             "convert (decimal (16, 3), TzPurRt " + _
             "* (case when TMulDiv= 'D' then TCnvRt else 1 end)/ (case when TMulDiv= 'M' then TCnvRt else 1 end) " + _
             "* " + IIF(mMulDiv = "M", CStr(mCnvFct), CStr(1)) + _
             "/ " + IIF(mMulDiv = "D", CStr(mCnvFct), CStr(1)) + ") As qTzPurRtInInvCur, " + _
             "convert (decimal (16, 3), TzPurRt " + _
             "* (case when TMulDiv= 'D' then TCnvRt else 1 end)/ (case when TMulDiv= 'M' then TCnvRt else 1 end)) " + _
             " as qTzPurRtInBaseCur, TzRmCdDesc, TzDesc " + IIF(gs_Partition = ctCurrPrtn, "", ", TzPrtKey ") + _
             "from TxndZ /* with (index(xTzCls))*/ " + _
             "Join Txn on TIdNo=TzTIdNo and TCoCd= TzCoCd and TTc= TzTc and TYy= TzYy and TChr= TzChr and TNo= TzNo " + _
             "Where TzCls= 'N' and TzRmZCd= '" + mRmZCd + "'" + _
             IIF(gs_Partition = ctCurrPrtn, " and TzPrtKey='" + ctCurrPrtn + "' ", "") + _
             " and TzCoCd= '" + mCoCd + "' and TzBM= 'B' and " + _
             "TzRmWt> TzRmDespWt " + ws_RtRange + wSRt
             
             '*** (Bef speed)
             '"TzRmWt> TzRmDespWt " + " and convert (decimal (16, 3), TzPurRt* (case when TMulDiv= 'D' then TCnvRt else 1 end)/ (case when TMulDiv= 'M' then TCnvRt else 1 end) " + _
             "* " + IIF(mMulDiv = "M", CStr(mCnvFct), CStr(1)) + _
             "/ " + IIF(mMulDiv = "D", CStr(mCnvFct), CStr(1)) + ")>= convert (decimal (16, 3), " + CStr(mPurRtFr) + ") " + wSRt
             '*** (Bef speed)
  '*** (Jen 2.13)
   
   
  With mwhlp
    .Sqlstrg = wSqlStrg
        
    '*** (Bef 2.13)
    '.ColHeadg = "|Base Voucher Number       |BOE Number                         |BOE Dt   |UseBef Dt    |" + _
    '            "Bal Wt       |Pc Per Ct    |Purchase Rt     |" + _
    '            "Rm Code                           |Yy |Key        |Description                       |"
    '.ColFlds = "|qVch|TBOENo|TBOEDt|TUseBefDt|qBalWt|TzPcPerCt|TzPurRt|TzRmCdDesc|TzYy|TzKey|TzDesc|"
    '*** (Bef 2.13)
    
    '*** (Jen 2.13)
    '*** (Bef 06/11/06)
    '.ColHeadg = "|Base Voucher Number      |BOE Number                      |BOE Dt   |UseBef Dt   |" + _
    '            "Bal Wt       |Pc Per Ct   |Rt (Inv Cur)    |Purchase Rt    |Currency   |" + _
    '            "Rm Code                         |Description                       |"
    '*** (Bef 06/11/06)
    
    '*** (06/11/06)
    .ColHeadg = "|Bst |Base Vch Number          |BOE Number                    |BOE Dt   |Bill Number                   |UseBef    |" + _
                "Bal Wt       |Pc/Ct |Rt (Inv Cur)     |Pur Rt         |Cur  |" + _
                "Rm Code               |Description               |" + _
                IIF(gs_Partition <> ctCurrPrtn, "Prtn |", "")
    '*** (06/11/06)
    .ColFlds = "|qBestUse|qVch|TBOENo|TBOEDt|TBillNo|TUseBefDt|qBalWt|TzPcPerCt|qTzPurRtInInvCur|TzPurRt|TCurrCd|TzRmCdDesc|TzDesc|" + _
               IIF(gs_Partition <> ctCurrPrtn, "TzPrtKey|", "")
    '*** (Jen 2.13)
    '.ColMask = "||||dd/mm/yy|dd/mm/yy||||||"
    .RetFldNm = "qVch"
    .SearchColNum = 2
    .SortAsc = True
  End With
End Sub
Public Sub BaseFgdSr(ByVal mCoCd As String, ByVal mEntTyp As String)
  '****** Sachin 3.02.0 - Partition Key added
  wSqlStrg = "SELECT " + IIF(gs_Partition = ctCurrPrtn, "", "FdPrtKey,") + " FdTc+'/'+FdYy+'/'+FdChr+'/'+LTrim(Str(FdNo))+'/'+LTrim(Str(FdSr)) as qVch, " + _
             "FdExpOdTc+'/'+FdExpOdYy+'/'+FdExpOdChr+'/'+LTrim(Str(FdExpOdNo))+'/'+" + _
             "LTrim(Str(FdExpOdSr)) as qExpOrd, (FdQty- FdDespQty) as qBalQty, " + _
             "FdYy, FdKey, FdInExpNo, FdDmCd, FdDmSz, FdSfx from Fgd " + _
             "join Param on PTyp= 'TCTYP' and PMCd= (Select PValue from Param where " + _
             "PTyp= 'TC' and PMCd= FdTc) " + _
             "where PValue1= '" + mEntTyp + "' and FdBM= 'B' and " + _
             "FdQty- FdDespQty> 0 and FdCoCd = '" + mCoCd + "' " + _
             IIF(gs_Partition = ctCurrPrtn, " And FdPrtKey='" + ctCurrPrtn + "'", "") + _
             "Order By FdTc, FdYy, FdChr, FdNo, FdSr"
  With mwhlp
    .Sqlstrg = wSqlStrg
    .ColHeadg = IIF(gs_Partition <> ctCurrPrtn, "|Prtn ", "") + "|Base Voucher Number                   |Balance Qty |" + _
                "Design                          |Suffix                                   |Size |" + _
                "Export Order                        |Yy|Key        |"
    .ColFlds = IIF(gs_Partition <> ctCurrPrtn, "|FdPrtKey", "") + "|qVch|qBalQty|FdDmCd|FdDmSz|FdSfx|qExpOrd|FdYy|FdKey|"
    .ColMask = IIF(gs_Partition <> ctCurrPrtn, "|", "") + "|||||||||"
    .SearchColNum = IIF(gs_Partition <> ctCurrPrtn, 2, 1)
    .RetFldNm = "qVch"
    .SortAsc = True
  End With
End Sub
Public Sub Tables()
  wSqlStrg = "select Name as qTableName from sysobjects where xtype= 'U' Order By Name "
  With mwhlp
    .Sqlstrg = wSqlStrg
    .ColHeadg = "|Table Name                               |"
    .ColFlds = "|qTableName|"
    .ColMask = "||"
    .SearchColNum = 1
    .RetFldNm = "qTableName"
    .SortAsc = True
  End With
End Sub
Public Sub TableFields(ByVal mTblName As String)
  wSqlStrg = "Select COLUMN_NAME From INFORMATION_SCHEMA.Columns where TABLE_NAME = '" + mTblName + "' Order By COLUMN_NAME "
  With mwhlp
    .Sqlstrg = wSqlStrg
    .ColHeadg = "|Column Name                               |"
    .ColFlds = "|COLUMN_NAME|"
    .ColMask = "||"
    .SearchColNum = 1
    .RetFldNm = "COLUMN_NAME"
    .SortAsc = True
  End With
End Sub
Public Sub BkCd(ByVal mBkCtg As String)
  With mwhlp
    .Sqlstrg = "SELECT BkCd, BkName from BnkAgMst where BkCtg ='" + Trim(mBkCtg) + "' order by BkCtg,BkCd, BkName "
    .ColHeadg = "|Bank/Agent Code |Bank/Agent Name                                                     |"
    .ColFlds = "|BkCd|BkName|"
    .RetFldNm = "BkCd"
    .SearchColNum = 1
    .SortAsc = True
  End With
End Sub
Public Sub RangeSz(ByVal mRmCtg As String, ByVal mRmSCtg As String)
  With mwhlp
    .Sqlstrg = "SELECT RrRngSz, RrFrLn, RrToLn from RmRt where RrTcTyp= 'RSZ' and " + _
               "RrCmCtg= 'C' and RrCtg= '" + mRmCtg + "' and RrSCtg= '" + mRmSCtg + "' " + _
               "order by RrRngSz, RrFrLn, RrToLn "
    .ColHeadg = "|Range Size          |From Size  |To Size     |"
    .ColFlds = "|RrRngSz|RrFrLn|RrToLn|"
    .RetFldNm = "RrRngSz"
    .SearchColNum = 1
    .SortAsc = True
  End With
End Sub
Public Sub DmSetCd()
  '*** Jay 3.1.0_1 [Dsg SetCd From Param]
  '****** Sachin 3.02.0 - Partition Key added
  If go_HlpCn.RecSeek("Select 'x' from Head where HCoCd= '" + ctSelfCoCd + "' and  HValidDsgSetCd = 'N'") Then
    With mwhlp
      .Sqlstrg = "SELECT distinct DmSetCd from DsgMst " + _
                 "Where " + _
                 IIF(gs_Partition = ctCurrPrtn, "DmPrtKey='" + ctCurrPrtn + "' And  ", "") + _
                 " DmSetCd <> '' Order by DmSetCd "
      .ColHeadg = IIF(gs_Partition <> ctCurrPrtn, "|Prtn ", "") + "|Set Cd                      |"
      .ColFlds = IIF(gs_Partition <> ctCurrPrtn, "|DmPrtKey", "") + "|DmSetCd|"
      .RetFldNm = "DmSetCd"
      .SearchColNum = IIF(gs_Partition <> ctCurrPrtn, 2, 1)
      .SortAsc = True
    End With
  Else
    Call PMCd("SETCD")
  End If
  '*** Jay 3.1.0_1 [Dsg SetCd From Param]
End Sub
Public Sub YyyyMm()
  With mwhlp
    .Sqlstrg = "SELECT distinct RhYyyyMm from RmRtHist where RhYyyyMm <> '' Order by RhYyyyMm Desc "
    .ColHeadg = "|YyyyMm            |"
    .ColFlds = "|RhYyyyMm|"
    .RetFldNm = "RhYyyyMm"
    .SearchColNum = 1
    .SortAsc = False
  End With
End Sub
Public Sub CustYyyyMm(ByVal ps_CmCtg As String, ByVal ps_CmCd As String, ByVal ps_CmLkUpRmRt As String)
  Dim wCmCdStr As String
  '*************************** ZUBIN **************************
  ' 08th Dec 2003, EMR206
  ' Shows YyyyMm for Specific CmCtg and CmCd
  With mwhlp
   ' .Sqlstrg = "SELECT distinct RhYyyyMm, RhCmCtg, RhCmCd from RmRtHist where RhYyyyMm <> '' and " + _
                "RhCmCtg = '" + ps_CmCtg + "' and RhCmCd = '" + ps_CmCd + "' " + _
                "Order by RhYyyyMm "
                
      '******************************Geeta***********Emr206***********************
      '***********************If customer records are not present then Zself Year will be given******
      '6.1-13 (select CmLkUpRmRt from CustMst where CmCtg='C' and CmCd='" + ps_CmCd + "' ) removed and ps_CmLkUpRmRt added
       wCmCdStr = IIF(ps_CmCtg = "P", "'" + ps_CmLkUpRmRt + "'", "'" + ps_CmCd + "'")
       .Sqlstrg = "select Ym.RhYyYyMm qRhYyYyMm,isNull(RhCmCtg,'" + ps_CmCtg + "') qRhCmCtg,isNull(Rh.RhCmCd,'" + ctSelfCmCd + "')  qRhCmCd  " + _
                  "from  (SELECT Distinct RhYyYyMm from RmRtHist  where RhCmctg='" + ps_CmCtg + "'" + _
                  " and RhcmCd in (" + wCmCdStr + ",'" + ctSelfCmCd + "')) Ym " + _
                  "left Outer join (select distinct RhYyYyMm,RhCmctg ,RhCmCd from RmRtHist " + _
                  "where RhCmCtg ='" + ps_CmCtg + "' and RhCmCd =" + wCmCdStr + ") Rh  on  Rh.RhYyYyMm=Ym.RhYyYyMm " + _
                  "Order by Ym.RhYyyyMm  Desc "
  
    .ColHeadg = "|CustCtg  |CustCd                |YyyyMm            |"
    .ColFlds = "|qRhCmCtg|qRhCmCd|qRhYyyyMm|"
    .RetFldNm = "qRhYyyyMm"
    .SearchColNum = 1
    .SortAsc = False
  End With
   '******************************Geeta***********Emr206***********************
End Sub
'*** Jay 2.14 ***(pb_ShowValidOnly )
Public Sub CmName(ByVal mCmCtg As String, Optional ByVal pb_ShowValidOnly As Boolean = True)
'Public Sub CmName(ByVal mCmCtg As String)
'*** Jay 2.14 ***
  With mwhlp
    '*** Jay 2.14 ***
    .Sqlstrg = "SELECT CmCd, CmName, CmCurCd from CustMst where CmCtg= '" + mCmCtg + "'" + _
                IIF(pb_ShowValidOnly = True, " And CmValidYN = 'Y' ", " ") + " order by CmName "
    '.Sqlstrg = "SELECT CmCd, CmName, CmCurCd from CustMst where CmCtg= '" + mCmCtg + "' order by CmName "
    '*** Jay 2.14 ***
    .ColHeadg = "|Code       |Name                                                   |Currency        |"
    .ColFlds = "|CmCd|CmName|CmCurCd|"
    .RetFldNm = "CmName"
    .SearchColNum = 2
    .SortAsc = True
  End With
End Sub
Public Sub RpYyyyMm()
  With mwhlp
    .Sqlstrg = "Select distinct RpYyyyMm from RecPrcYyMm Order by RpYyyyMm Desc "
    .ColHeadg = "|YyyyMm            |"
    .ColFlds = "|RpYyyyMm|"
    .RetFldNm = "RpYyyyMm"
    .SearchColNum = 1
    .SortAsc = False
  End With
End Sub
Public Sub LrcYyMm(ByVal mCoCd As String)
  With mwhlp
    .Sqlstrg = "Select distinct LrcYyMm from LsRec where LrcCoCd = '" + mCoCd + "' Order by LrcYyMm Desc "
    .ColHeadg = "|YyMm            |"
    .ColFlds = "|LrcYyMm|"
    .RetFldNm = "LrcYyMm"
    .SearchColNum = 1
    .SortAsc = False
  End With
End Sub

'*************************************** XEMR *********************************************
'*************************************** XEMR *********************************************
'*************************************** XEMR *********************************************
Public Sub xHCd()
  With mwhlp
    .Sqlstrg = "SELECT xHCd, xHName from xHead order by xHCd "
    .ColHeadg = "|Company Code|Name                       |"
    .ColFlds = "|xHCd|xHName|"
    .RetFldNm = "xHCd"
    .SearchColNum = 1
    .SortAsc = True
  End With
End Sub
Public Sub xTpTcTyp(Optional ByVal pv_VchTyp As Variant)
  Dim ws_Cnd As String
  ws_Cnd = ""
  
  If Not IsMissing(pv_VchTyp) Then
    If Trim(pv_VchTyp) <> "" Then ws_Cnd = " Where xTpVchTyp in (" + pv_VchTyp + ") "
  End If
  
  With mwhlp
    .Sqlstrg = "SELECT xTpTcTyp, xTpDesc from xTcTyp " + ws_Cnd + "order by xTpTcTyp "
    .ColHeadg = "|Tc Type|Description                     |"
    .ColFlds = "|xTpTcTyp|xTpDesc|"
    .RetFldNm = "xTpTcTyp"
    .SearchColNum = 1
    .SortAsc = True
  End With
End Sub
Public Sub xDocNo(ByVal mCoCd As String)
  With mwhlp
    .Sqlstrg = "SELECT xTDocNo, xTDocDt, xTTc+ '/'+ xTYy+ '/'+ xTChr+ '/'+ LTrim(Str(xTNo)) " + _
                "as qVchNo, xTCmCd from xTxn where xTCoCd= '" + mCoCd + "' order by xTDocNo"
    .ColHeadg = "|DocNo          |DocDt     |VchNo               |CmCd         |"
    .ColFlds = "|xTDocNo|xTDocDt|qVchNo|xTCmCd|"
    .ColMask = "||" + DtShortStr() + "|||"    '6.1
    .RetFldNm = "xTDocNo"
    .SearchColNum = 1
    .SortAsc = True
  End With
End Sub
Public Sub xDsgLoc(Optional ByVal pv_LocTyp As Variant, Optional ByVal pv_NotInLocTyp As Variant)
  Dim ws_Cnd As String
  ws_Cnd = ""
  If Not IsMissing(pv_LocTyp) Then
    If Trim(pv_LocTyp) <> "" Then ws_Cnd = " Where xDlLocTyp in (" + pv_LocTyp + ") "
  End If
  If Not IsMissing(pv_NotInLocTyp) Then
    ws_Cnd = IIF(ws_Cnd = "", ws_Cnd + " where ", ws_Cnd + " and ")
    ws_Cnd = ws_Cnd + " xDlLocTyp not in (" + pv_NotInLocTyp + ") "
  End If
  With mwhlp
    .Sqlstrg = "Select xDlLocTyp, xDlLocCd, xDlDesc From xDsgLoc " + ws_Cnd + _
               "order by xDlLocTyp, xDlLocCd "
    .ColHeadg = "|Code     |Type|Desc                                                  |"
    .ColFlds = "|xDlLocCd|xDlLocTyp|xDlDesc|"
    .RetFldNm = "xDlLocCd"
    .SearchColNum = 1
    .SortAsc = True
  End With
End Sub
Public Sub xTNo(ByVal ps_Tc As String, ByVal ps_Yy As String, ByVal ps_Chr As String)
  Dim ws_SqlStrg As String
  ws_SqlStrg = "Select xTNo, xTDt, xTCmCd from xTxn " + _
               "where xTCoCd= '" + gs_CoCd + "' and xTTc= '" + ps_Tc + "' and " + _
               "xTYy= '" + ps_Yy + "' and xTChr= '" + ps_Chr + "' Order By xTNo"
  With mwhlp
    .Sqlstrg = ws_SqlStrg
    .ColHeadg = "|Number|Date        |Customer    |"
    .ColFlds = "|xTNo|xTDt|xTCmCd|"
    .ColMask = "||" + DtShortStr() + "||"   '6.1
    .SearchColNum = 1
    .RetFldNm = "xTNo"
    .SortAsc = True
  End With
End Sub
Public Sub xTdSr(ByVal ps_Tc As String, ByVal ps_Yy As String, ByVal ps_Chr As String, ByVal pf_No As Single)
  Dim wSqlStrg As String
  wSqlStrg = "SELECT xTdSr, xTdDmCd, xTdSfx, xTdDmSz from xTxnDsg " + _
             "where xTdTc= '" + ps_Tc + "' and xTdYy= '" + ps_Yy + "' and " + _
             "xTdChr= '" + ps_Chr + "' and xTdNo= " + CStr(pf_No) + " Order By xTdSr"
  With mwhlp
    .Sqlstrg = wSqlStrg
    .ColHeadg = "|Sr.  |Design                          |Suffix                                   |Size |"
    .ColFlds = "|xTdSr|xTdDmCd|xTdSfx|xTdDmSz|"
    .ColMask = "|||||"
    .SearchColNum = 1
    .RetFldNm = "xTdSr"
    .SortAsc = True
  End With
End Sub
Public Sub xImIdNo(ByVal ps_ImIdOwn As String, ByVal ps_ImIdYy As String, Optional ByVal ps_DmCd As Variant)
  Dim ws_SqlStrg As String, ws_Cnd As String
  If Not IsMissing(ps_DmCd) Then
    ws_Cnd = " and xTdDmCd= '" + ps_DmCd + "' "
  End If
  ws_SqlStrg = "SELECT xImIdNo, xTdSfx, xTdDmSz, xTdTc+ '/'+ xTdYy+ '/'+ xTdChr+ '/'+ LTrim(Str(xTdNo))+ '/'+ " + _
               "LTrim(Str(xTdSr)) as qSrcVch, (Select xTdTc+ '/'+ xTdYy+ '/'+ xTdChr+ '/'+ LTrim(Str(xTdNo))+ '/'+ " + _
               "LTrim(Str(xTdSr)) from xTxnDsg where xTdCoCd= xImCoCd and xTdYy= xImLstTdYy and " + _
               "xTdKey= xImLstTdKey) as qLstVch " + _
               "from xIdMst join xTxnDsg xTd on xTdCoCd= xImCoCd and xTdYy= xImSrcTdYy and " + _
               "xTdKey= xImSrcTdKey " + _
               "where xImCoCd= '" + gs_CoCd + "' and xImIdOwn= '" + ps_ImIdOwn + "' and " + _
               "xImIdYy= '" + ps_ImIdYy + "' " + ws_Cnd + " Order By xImIdNo"
  With mwhlp
    .Sqlstrg = ws_SqlStrg
    .ColHeadg = "|Id Number|Suffix                                   |Size  |Source Voucher               |Last Voucher                 |"
    .ColFlds = "|xImIdNo|xTdSfx|xTdDmSz|qSrcVch|qLstVch|"
    .ColMask = "||||||"
    .SearchColNum = 1
    .RetFldNm = "xImIdNo"
    .SortAsc = True
  End With
End Sub
Public Sub xTdSfx(ByVal ps_xTdDmCd As String, Optional ByVal ps_CmCd As Variant)
  Dim ws_Cnd As String
  If Not IsMissing(ps_CmCd) Then
    ws_Cnd = " and xTCmCd= '" + ps_CmCd + "' "
  End If
  With mwhlp
    .Sqlstrg = "Select xTdSfx, xTdDmSz, xTdDmCol, " + _
               "xTdKt, xTCmCd, xTdOffValPpc, xTDt, xTdSfxDesc from xTxnDsg " + _
               "join xTxn on xTCoCd= xTdCoCd and xTTc= xTdTc and xTYy= xTdYy and " + _
               "xTChr= xTdChr and xTNo= xTdNo " + _
               "where xTdCoCd= '" + gs_CoCd + "' and " + _
               "xTdTc in (Select xPMCd from xParam where xPTyp= 'TC' and " + _
               "xPValue in('PDC')) and xTdDmCd= '" + ps_xTdDmCd + "' " + ws_Cnd + _
               "order by xTdSfx, xTdDmSz "
    .ColHeadg = "|Suffix                                   |Size  |Colour |Karat  |Customer     |Sales Price|Date        |Desc                                                        |"
    .ColFlds = "|xTdSfx|xTdDmSz|xTdDmCol|xTdKt|xTCmCd|xTdOffValPpc|xTDt|xTdSfxDesc|"
    .RetFldNm = "xTdSfx"
    .SearchColNum = 1
    .SortAsc = True
  End With
End Sub
Public Sub MemBaseVchSr(ByVal ps_TcTyp As String, ByVal ps_CmCtg As String, _
                        ByVal ps_CmCd As String, ByVal ps_DmTyp As String, _
                        Optional ByVal ps_DmCd As Variant, _
                        Optional ByVal ps_ImIdOwn As Variant, _
                        Optional ByVal ps_ImIdYy As Variant, _
                        Optional ByVal ps_ImIdNo As Variant)
  
  Dim ws_SqlStrg As String, ws_Cnd As String
  If Not IsMissing(ps_DmCd) Then
    ws_Cnd = " and xTdDmCd= '" + ps_DmCd + "' "
  End If
  If Not (IsMissing(ps_ImIdOwn) Or IsMissing(ps_ImIdYy) Or IsMissing(ps_ImIdNo)) Then
    ws_Cnd = ws_Cnd + " and ((xTdIdOwnBy= '" + ps_ImIdOwn + "' and xTdIdYy= '" + ps_ImIdYy + "' " + _
             " and xTdIdNo= " + CStr(ps_ImIdNo) + ") or (xTdIdOwnBy= '' and xTdIdYy= '' and xTdIdNo= 0)) "
  End If
  ws_SqlStrg = "Select xTdTc+'/'+xTdYy+'/'+xTdChr+'/'+LTrim(Str(xTdNo))+'/'+LTrim(Str(xTdSr)) as qVch, " + _
               "(xTduQty- xTduMemMtchQty) as qBalQty, " + _
               "xTdYy, xTdKey, xTdDmCd, xTdDmSz, xTdSfx ,(case when xTdIdNo> 0 then " + _
               "xTdIdOwnBy+'/'+xTdIdYy+'/'+LTrim(Str(xTdIdNo)) else '' end) as qId from xTxnDsg " + _
               "join xTxnDsgUpd on xTduCoCd= xTdCoCd and xTduTc= xTdTc and xTduYy= xTdYy and " + _
               "xTduChr= xTdChr and xTduNo= xTdNo and xTduSr= xTdSr and xTduSrNo= xTdSrNo " + _
               "join xTxn on xTCoCd= xTdCoCd and xTTc= xTdTc and xTYy= xTdYy and xTChr= xTdChr and " + _
               "xTNo= xTdNo join xParam on xPTyp= 'TC' and xPMCd= xTdTc " + _
               "join xTcTyp on xTpTcTyp= xPValue " + _
               "where xTdCoCd= '" + gs_CoCd + "' and xTdSrNo= 0 and xPValue= '" + ps_TcTyp + "' and " + _
               "xTpMemBM= 'B' and xTCmCtg= '" + ps_CmCtg + "' and xTCmCd= '" + ps_CmCd + "' and " + _
               "xTdDmTyp= '" + ps_DmTyp + "' and xTduQty- xTduMemMtchQty> 0 " + ws_Cnd + _
               " Order By xTdTc, xTdYy, xTdChr, xTdNo, xTdSr "
  With mwhlp
    .Sqlstrg = ws_SqlStrg
    .ColHeadg = "|Base Voucher Number         |Bal Qty  |" + _
                "Design                  |Suffix                                   |Size |Id Number      |" + _
                "Yy|Key     |"
    .ColFlds = "|qVch|qBalQty|xTdDmCd|xTdDmSz|xTdSfx|qId|xTdYy|xTdKey|"
    .ColMask = "|||||||||"
    .SearchColNum = 1
    .RetFldNm = "qVch"
    .SortAsc = True
  End With
End Sub
Public Sub TxnBaseVchSr(ByVal ps_TcTyp As String, ByVal ps_CmCtg As String, _
                        ByVal ps_CmCd As String, ByVal ps_DmTyp As String, _
                        Optional ByVal ps_DmCd As Variant, _
                        Optional ByVal ps_ImIdOwn As Variant, _
                        Optional ByVal ps_ImIdYy As Variant, _
                        Optional ByVal ps_ImIdNo As Variant)
  Dim ws_SqlStrg As String, ws_Cnd As String
  If Not IsMissing(ps_DmCd) Then
    ws_Cnd = " and xTdDmCd= '" + ps_DmCd + "' "
  End If
  If Not (IsMissing(ps_ImIdOwn) Or IsMissing(ps_ImIdYy) Or IsMissing(ps_ImIdNo)) Then
    ws_Cnd = ws_Cnd + " and ((xTdIdOwnBy= '" + ps_ImIdOwn + "' and xTdIdYy= '" + ps_ImIdYy + "' " + _
             " and xTdIdNo= " + CStr(ps_ImIdNo) + ") or (xTdIdOwnBy= '' and xTdIdYy= '' and xTdIdNo= 0)) "
  End If
  
  ws_SqlStrg = "Select xTdTc+'/'+xTdYy+'/'+xTdChr+'/'+LTrim(Str(xTdNo))+'/'+LTrim(Str(xTdSr)) as qVch, " + _
               "(xTduQty- xTduTxnMtchQty) as qBalQty, xTdYy, xTdKey, xTdDmCd, xTdDmSz, " + _
               "xTdSfx, (case when xTdIdNo> 0 then xTdIdOwnBy+'/'+xTdIdYy+'/'+LTrim(Str(xTdIdNo)) else '' end) as qId from xTxnDsg " + _
               "join xTxnDsgUpd on xTduCoCd= xTdCoCd and xTduTc= xTdTc and xTduYy= xTdYy and " + _
               "xTduChr= xTdChr and xTduNo= xTdNo and xTduSr= xTdSr and xTduSrNo= xTdSrNo " + _
               "join xTxn on xTCoCd= xTdCoCd and xTTc= xTdTc and xTYy= xTdYy and xTChr= xTdChr and " + _
               "xTNo= xTdNo join xParam on xPTyp= 'TC' and xPMCd= xTdTc " + _
               "join xTcTyp on xTpTcTyp= xPValue " + _
               "where xTdCoCd= '" + gs_CoCd + "' and xTdSrNo= 0 and xPValue= '" + ps_TcTyp + "' and " + _
               "xTpTxnBM= 'B' and xTCmCtg= '" + ps_CmCtg + "' and xTCmCd= '" + ps_CmCd + "' and " + _
               "xTdDmTyp= '" + ps_DmTyp + "' and xTduQty- xTduTxnMtchQty> 0 " + ws_Cnd + _
               " Order By xTdTc, xTdYy, xTdChr, xTdNo, xTdSr "
  With mwhlp
    .Sqlstrg = ws_SqlStrg
    .ColHeadg = "|Base Voucher Number         |Bal Qty  |" + _
                "Design                  |Suffix                                   |Size |Id Number       |" + _
                "Yy|Key    |"
    .ColFlds = "|qVch|qBalQty|xTdDmCd|xTdDmSz|xTdSfx|qId|xTdYy|xTdKey|"
    .ColMask = "|||||||||"
    .SearchColNum = 1
    .RetFldNm = "qVch"
    .SortAsc = True
  End With
End Sub
Public Sub CpySrcVchNo(ByVal ps_Tc As String, ByVal ps_Yy As String, ByVal ps_Chr As String, ByVal ps_CustCd As String, ByVal ps_NotInCustCd As String)
  Dim ws_SqlStrg As String, ws_Cnd As String
  If ps_CustCd <> "" Then
    ws_Cnd = " and xTCmCd in (" + ps_CustCd + ") "
  End If
  If ps_NotInCustCd <> "" Then
    ws_Cnd = ws_Cnd + " and xTCmCd not in (" + ps_NotInCustCd + ") "
  End If
  
  ws_SqlStrg = "Select xTNo,xTCmCd, xTDt from xTxn " + _
               "where xTCoCd= '" + gs_CoCd + "' and xTTc= '" + ps_Tc + "' and " + _
               "xTYy= '" + ps_Yy + "' and xTChr= '" + ps_Chr + "' " + ws_Cnd + _
               " Order By xTNo"
  With mwhlp
    .Sqlstrg = ws_SqlStrg
    .ColHeadg = "|Number|Customer      |Date        |"
    .ColFlds = "|xTNo|xTCmCd|xTDt|"
    .ColMask = "|||" + DtShortStr() + "|"   '6.1
    .SearchColNum = 1
    .RetFldNm = "xTNo"
    .SortAsc = True
  End With
End Sub
Public Sub CpySrcVchSr(ByVal ps_SrcTcTyp As String, ByVal ps_DmTyp As String, _
                       ByVal ps_DmCd As String, ByVal ps_Sfx As String, _
                       ByVal ps_DmSz As String, ByVal ps_CmCtg As String, _
                       ByVal ps_CustCd As String, ByVal ps_NotInCustCd As String, _
                       Optional ByVal ps_IdOwnBy As Variant, Optional ByVal ps_IdYy As Variant, _
                       Optional ByVal ps_IdNo As Variant)
    
  Dim ws_SqlStrg As String, ws_Cnd As String
  
  If ps_CustCd <> "" Then
    ws_Cnd = " and xTCmCd in (" + ps_CustCd + ") "
  End If
  If ps_NotInCustCd <> "" Then
    ws_Cnd = ws_Cnd + " and xTCmCd not in (" + ps_NotInCustCd + ") "
  End If
  
  If Not IsMissing(ps_IdOwnBy) Then
    ws_Cnd = ws_Cnd + " and xTdIdOwnBy= '" + ps_IdOwnBy + "' "
  End If
  If Not IsMissing(ps_IdYy) Then
    ws_Cnd = ws_Cnd + " and xTdIdYy= '" + ps_IdYy + "' "
  End If
  If Not IsMissing(ps_IdNo) Then
    ws_Cnd = ws_Cnd + " and xTdIdNo= " + CStr(ps_IdNo) + " "
  End If
  
  If ps_Sfx <> "" Then ws_Cnd = ws_Cnd + " and xTdSfx= '" + ps_Sfx + "' "
  If ps_DmSz <> "" Then ws_Cnd = ws_Cnd + " and xTdDmSz= '" + ps_DmSz + "' "
  
  'and xTdSfx= '" + ps_Sfx + "' and xTdDmSz= '" + ps_DmSz + "'
  
  ws_SqlStrg = "Select xTdTc+'/'+xTdYy+'/'+xTdChr+'/'+LTrim(Str(xTdNo))+'/'+LTrim(Str(xTdSr)) as qVch, " + _
               "xTdSfx, xTdDmSz, xTdIdOwnBy, xTdIdYy, xTdIdNo, xTduQty, (xTduQty- xTduMemMtchQty) as qMemBalQty, " + _
               "(xTduQty- xTduTxnMtchQty) as qTxnBalQty, xTDt from xTxnDsg " + _
               "join xTxnDsgUpd on xTduCoCd= xTdCoCd and xTduTc= xTdTc and xTduYy= xTdYy and " + _
               "xTduChr= xTdChr and xTduNo= xTdNo and xTduSr= xTdSr and xTduSrNo= xTdSrNo " + _
               "join xTxn on xTCoCd= xTdCoCd and xTTc= xTdTc and xTYy= xTdYy and xTChr= xTdChr and " + _
               "xTNo= xTdNo join xParam on xPTyp= 'TC' and xPMCd= xTdTc " + _
               "join xTcTyp on xTpTcTyp= xPValue " + _
               "where xTdCoCd= '" + gs_CoCd + "' and xTdSrNo= 0 and xPValue= '" + ps_SrcTcTyp + "' and " + _
               "xTCmCtg= '" + ps_CmCtg + "' and xTdDmTyp= '" + ps_DmTyp + "' and " + _
               "xTdDmCd= '" + ps_DmCd + "' " + ws_Cnd + _
               " Order By xTDt desc, xTdTc, xTdYy, xTdChr, xTdNo desc, xTdSr desc"
  With mwhlp
    .Sqlstrg = ws_SqlStrg
    .ColHeadg = "|Src  Voucher Number           |" + _
                "Suffix                                   |Size |" + _
                "IdOwn|IdYy|IdNo        |       Qty |   Mem Bal|   Txn Bal|Date        |"
    .ColFlds = "|qVch|xTdSfx|xTdDmSz|xTdIdOwnBy|xTdIdYy|xTdIdNo|xTduQty|qMemBalQty|qTxnBalQty|xTDt|"
    '.ColMask = "||||||||dd/mm/yy|"
    .SearchColNum = 1
    .RetFldNm = "qVch"
    .SortAsc = True
  End With
End Sub
'*************************************** XEMR *********************************************
'*************************************** XEMR *********************************************
'*************************************** XEMR *********************************************

Public Sub MenuCode()
    With mwhlp
        .Sqlstrg = "SELECT UaMnuCd,UaMnuDesc from UsrAccess where UaUsrCd= '" + ctGenMenuUsr + "' order by UaMnuCd"
        .ColHeadg = "|MenuCode                 |Description                                           |"
        .ColFlds = "|UaMnuCd|UaMnuDesc|"
        .RetFldNm = "UaMnuCd"
        .SearchColNum = 1
        .SortAsc = True
    End With
End Sub

'****** Sachin 2.14.0 Next [03/02/2007] ******
Public Sub MenuTag(ByVal mMnuFrm As String)
    With mwhlp
        .Sqlstrg = "SELECT UaTag,UaMnuCd,UaMnuDesc from UsrAccess where UaUsrCd= '" + ctGenMenuUsr + "' And UaMnuForm='" + mMnuFrm + "'  Order By UaTag"
        .ColHeadg = "|MenuTag|MenuCode                 |Description                                           |"
        .ColFlds = "|UaTag|UaMnuCd|UaMnuDesc|"
        .RetFldNm = "UaTag"
        .SearchColNum = 1
        .SortAsc = True
    End With
End Sub
'****** Sachin 2.14.0 Next [03/02/2007] ******

Public Property Let MultiSelect(ByVal vNewValue As Boolean)
  mwhlp.MultiSelect = vNewValue
End Property

' ########################  Manoj #### Ver: 2.0.6 #### Date: 02/02/2004  ########################
Public Sub TgfCd(ByVal mCoCd As String)
  With mwhlp
    .Sqlstrg = "SELECT TgfCd, TgfDesc, TgfStartStr, TgfEndStr  from TagFmt Where TgfCoCd = '" + mCoCd + "'"
    .ColHeadg = "|Tag Cd       |Description                                     |Start String|End String  |"
    .ColFlds = "|TgfCd|TgfDesc|TgfStartStr|TgfEndStr|"
    .RetFldNm = "TgfCd"
    .SearchColNum = 1
    .SortAsc = True
  End With
End Sub
' ###########################################  Manoj  ###########################################

' ########################  Manoj #### Ver: 2.0.6 #### Date: 02/02/2004  ########################
Public Sub TgSr(ByVal mCoCd As String, ByVal mTc As String, ByVal mYy As String, ByVal mChr As String, ByVal mNo As Single)
  wSqlStrg = "SELECT TgSr, TgCd, TgLevel, TgAvgTyp, TgWt, TgPrice from Tag " + _
             "where TgCoCd= '" + mCoCd + "' and TgTc= '" + mTc + "' and " + _
             "TgYy='" + mYy + "' and TgChr='" + mChr + "' and TgNo = " + CStr(mNo) + _
             " Order By TgSr Desc"
  With mwhlp
    .Sqlstrg = wSqlStrg
    .ColHeadg = "|Sr.   |Tag Cd  |Tag Level    |Avg Type   |Weight  |Price  |"
    .ColFlds = "|TgSr|TgCd|TgLevel|TgAvgTyp|TgWt|TgPrice|"
    .ColMask = "|||||||"
    .SearchColNum = 1
    .RetFldNm = "TgSr"
    .SortAsc = True
  End With
End Sub
Public Sub TgdSrNo(ByVal mCoCd As String, ByVal mTc As String, ByVal mYy As String, ByVal mChr As String, ByVal mNo As Single, ByVal mSr As Single)
  wSqlStrg = "SELECT TgdSrNo, TgdIdSr, TgdQty from TagData " + _
             "where TgdCoCd= '" + mCoCd + "' and TgdTc= '" + mTc + "' and " + _
             "TgdYy='" + mYy + "' and TgdChr='" + mChr + "' and TgdNo = " + CStr(mNo) + _
             " and TgdSr = " + CStr(mSr) + _
             " Order By TgdSrNo"
  With mwhlp
    .Sqlstrg = wSqlStrg
    .ColHeadg = "|Sr.No  |Id Sr. |Qty   |"
    .ColFlds = "|TgdSrNo|TgdIdSr|TgdQty|"
    .ColMask = "||||"
    .SearchColNum = 1
    .RetFldNm = "TgdSrNo"
    .SortAsc = True
  End With
End Sub
' ###########################################  Manoj  ###########################################
'****Geeta***2.11
Public Sub OrSetCd(ByVal ps_OdCoCd As String, ByVal ps_OdTc As String, ByVal ps_OdYy As String, ByVal ps_OdChr As String, ByVal ps_OdNo As Single, ByVal ps_OdSr As Single, Optional ByVal ps_RmCtg As Variant)
' This procedure returns all the 'SET' Sub Cd that exist in the given OdSr
  Dim wRmStr As String
  With mwhlp
       If IsMissing(ps_RmCtg) Then
            wRmStr = " "
       Else
            wRmStr = "  and OrRmCtg= '" + ps_RmCtg + " ' "
       End If
       
      .Sqlstrg = " Select Distinct OrRmCtg,OrSetSCd from OrdRm " + _
                 " Where OrRmCtg in ('D','C') and " + _
                 IIF(gs_Partition = ctCurrPrtn, " OrPrtKey='" + ctCurrPrtn + "' And ", "") + _
                 " OrCoCd='" + ps_OdCoCd + "'" + _
                 " and OrTc= '" + ps_OdTc + "' and OrYy='" + ps_OdYy + "' and OrChr='" + ps_OdChr + "'" + _
                 " and OrNo=" + Str(ps_OdNo) + " And OrSr = " + Str(ps_OdSr) + wRmStr + _
                 " Order by OrRmCtg, OrSetSCd"

    .ColHeadg = "|RmCtg |Setting Code|"
    .ColFlds = "|OrRmCtg|OrSetSCd|"
    .RetFldNm = "OrSetSCd"
    .SearchColNum = 1
    .SortAsc = True
  End With
End Sub
' **** Zubin 211 **** '
' **** Zubin 211 **** '

' ########################  Manoj #### Ver: 2.11.0 #### Date: 01/06/2005  ########################
Public Sub EmailAddr()
' This procedure returns all the 'SET' Sub Cd that exist in the Order of the Bag No. passed as parameter
  With mwhlp
    .Sqlstrg = "Select UmUsrCd, PDesc, UmMailAdd From " + _
                "UsrMail Join Param On PTyp ='USR' And PMCd = UmUsrCd " + _
                "Order By UmUsrCd"
    .ColHeadg = "|User Code|User Name                |Email Address                                                           |"
    .ColFlds = "|UmUsrCd|PDesc|UmMailAdd|"
    .RetFldNm = "UmMailAdd"
    .SearchColNum = 3
    .SortAsc = True
  End With
End Sub
' ###########################################  Manoj  ###########################################

' ******* Sachin for Bag Reconciliation ******* '
'pg.19 -stock Reco- Order by Desc added
Public Sub BrNo(ByVal mCoCd As String, ByVal mTc As String, ByVal mYy As String, ByVal mChr As String)
  wSqlStrg = "Select BmNo,BmDt,BmBagCoCdFr,BmBagCoCdTo,BmBagCoCdSel,BmBagLocFr,BmBagLocTo,BmBagLocSel from BgRecoMst where BmCoCd= '" + mCoCd + "' " + _
             "and BmTc='" + mTc + "' " + _
             "and BmYy='" + mYy + "' " + _
             "and BmChr='" + mChr + "' Order By BmNo Desc"
  With mwhlp
    .Sqlstrg = wSqlStrg
    .ColHeadg = "|Number|Br Date  |Co Fr |Co To |Co Sel          |Loc Fr |Loc To |Loc Sel          |"
    .ColFlds = "|BmNo|BmDt|BmBagCoCdFr|BmBagCoCdTo|BmBagCoCdSel|BmBagLocFr|BmBagLocTo|BmBagLocSel|"
    .ColMask = "||" + DtShortStr() + "|||||||"    '6.1
    .SearchColNum = 1
    .RetFldNm = "BmNo"
    .SortAsc = True
  End With
End Sub

Public Sub BrSr(ByVal mCoCd As String, ByVal mTc As String, ByVal mYy As String, ByVal mChr As String, ByVal mNo As Single)
  
  wSqlStrg = "SELECT BdSr from BgRecoDet Where BdCoCd= '" + mCoCd + "' " + _
             "and BdTc= '" + mTc + "' and BdYy='" + mYy + "' and " + _
             "BdChr='" + mChr + "' And BdNo='" + CStr(mNo) + "' Order By BdSr"

  With mwhlp
    .Sqlstrg = wSqlStrg
    .ColHeadg = "|Reco Sr.|"
    .ColFlds = "|BdSr|"
    .ColMask = "||"
    .SearchColNum = 1
    .RetFldNm = "BdSr"
    .SortAsc = True
  End With
End Sub
' ******* Sachin for Bag Reconciliation ******* '

Public Sub CsCd()
  '*** Jenny Color
  With mwhlp
    .Sqlstrg = "SELECT CsCd, CsName, CsEdit From ClrSet Order By CsCd "
    .ColHeadg = "|Template       |Description                             |Editable    |"
    .ColFlds = "|CsCd|CsName|CsEdit|"
    .RetFldNm = "CsCd"
    .SearchColNum = 1
    .SortAsc = True
  End With
  '*** Jenny Color
End Sub

'***********Geeta*****************Crm***2.12*******
Public Sub CustGrade(ByVal mCmCd As String, ByVal mGrTyp As String)
    wSqlStrg = "SELECT Distinct GrCd,PDesc  from Grade Join Param On PTyp='GRDCD' and PMCd=GrTyp and PSCd=GrCd " + _
               "Where GrCmCd= '" + mCmCd + "'" + _
               "and GrTyp= '" + mGrTyp + "' Order By GrCd"
  With mwhlp
    .Sqlstrg = wSqlStrg
    .ColHeadg = "|GrCd |Description          |"
    .ColFlds = "|GrCd|PDesc|"
    .SearchColNum = 1
    .RetFldNm = "GrCd"
    .SortAsc = True
  End With
End Sub
Public Sub GrRmCd(ByVal mGrTyp As String, ByVal mGrRmId As String)
  With mwhlp
     '****** Sachin 3.02.0 - Partition Key
     ' ***** Manali 3.10.0 - 18/05/12 - 'S','L' added for 'MET'
     .Sqlstrg = "Select " + IIF(gs_Partition = ctCurrPrtn, "", "RmPrtKey,") + " RmCd, RmCtg, RmSCtg, RmIntQly, RmDesc, RmValidYN " + _
                 "from RmMst " + _
                 "where " + _
                 IIF(gs_Partition = ctCurrPrtn, "RmPrtKey='" + ctCurrPrtn + "' and ", "") + _
                 "RmZ='N' and RmValidYN='Y' and " + _
                               "(( '" + mGrTyp + "'='MET' and RmCtg in ('G','P','X','S','L') ) Or " + _
                                          "('" + mGrTyp + "'='DIA' and RmCtg ='D' ) Or  " + _
                                          "('" + mGrTyp + "'='CS' and RmCtg ='C')) " + _
                 "order by RmCd "
    '578.13 Rmcd length increased
    .ColHeadg = IIF(gs_Partition <> ctCurrPrtn, "|Prtn ", "") + "|Code                                  |Ctg |Sub Ctg|International Qly|Desc                                                                              |Yn|"
    .ColFlds = IIF(gs_Partition <> ctCurrPrtn, "|RmPrtKey", "") + "|RmCd|RmCtg|RmSCtg|RmIntQly|RmDesc|RmValidYN|"
    .RetFldNm = "RmCd"
    .SearchColNum = IIF(gs_Partition <> ctCurrPrtn, 2, 1)
    .SortAsc = True
  End With
End Sub
'***********Geeta*****************Crm***2.12*******

' Zubin 212
'*** Jay 3.01 **** [LotNo Err][mLocTyp added]
'SRITODO
Public Sub LotNo(ByVal mCoCd As String, Optional ByVal mLoc As Variant, Optional ByVal mRmCd As Variant, _
                                    Optional ByVal mRmSz As Variant, Optional ByVal mRmSz2 As Variant, Optional ByVal mRmSz3 As Variant, Optional ByVal mLocTyp As Variant)
  Dim ws_Cnd As String, ws_GrpBy As String, ws_OrdBy As String
  Dim ws_ColHeadg As String, ws_ColFlds As String
  Dim ws_SqlFlds As String, ws_Join As String
  
  '*** Jay 3.01 **** [LotNo Err]
  ws_SqlFlds = ws_SqlFlds + ", SLoc "
  ws_GrpBy = ws_GrpBy + ", SLoc "
  ws_ColHeadg = ws_ColHeadg + "Loc        |": ws_ColFlds = ws_ColFlds + "SLoc|"
  ws_OrdBy = ws_OrdBy + ", SLoc "
  ws_Join = " "
  
  If Not IsMissing(mLocTyp) Then
    If mLocTyp <> "" Then
'      ws_Cnd = ws_Cnd + " and LocTyp = '" + mLocTyp + "' "
      ws_Join = " Join Loc On LocCoCd = SCoCd and LocCd= SLoc "
    End If
  End If
  '*** Jay 3.01 **** [LotNo Err]
  If Not IsMissing(mLoc) Then
    If mLoc <> "" Then
      'ws_SqlFlds = ws_SqlFlds + ", SLoc "        '*** Jay 3.01 Bef **** [LotNo Err]
      ws_Cnd = ws_Cnd + " and SLoc= '" + mLoc + "' "
      '*** Jay 3.01 Bef **** [LotNo Err]
      'ws_GrpBy = ws_GrpBy + ", SLoc "
      'ws_ColHeadg = ws_ColHeadg + "Loc        |": ws_ColFlds = ws_ColFlds + "SLoc|"
      'ws_OrdBy = ws_OrdBy + ", SLoc "
      '*** Jay 3.01 Bef **** [LotNo Err]
    End If
  End If
  
  If Not IsMissing(mRmCd) Then
    If mRmCd <> "" Then
      ws_SqlFlds = ws_SqlFlds + ", SRmCd "
      ws_Cnd = ws_Cnd + " and SRmCd= '" + mRmCd + "' ": ws_GrpBy = ws_GrpBy + ", SRmCd "
      ws_ColHeadg = ws_ColHeadg + "RmCd        |": ws_ColFlds = ws_ColFlds + "SRmCd|"
      ws_OrdBy = ws_OrdBy + ", SRmCd "
    End If
  End If
    
  If Not IsMissing(mRmSz) Then
    If mRmSz <> 0 Then
      ws_SqlFlds = ws_SqlFlds + ", SRmSz "
      ws_Cnd = ws_Cnd + " and SRmSz= '" + CStr(mRmSz) + "' ": ws_GrpBy = ws_GrpBy + ", SRmSz "
      ws_ColHeadg = ws_ColHeadg + "RmSz        |": ws_ColFlds = ws_ColFlds + "SRmSz|"
      ws_OrdBy = ws_OrdBy + ", SRmSz "
    End If
  End If
  
  If Not IsMissing(mRmSz2) Then
    If mRmSz2 <> 0 Then
      ws_SqlFlds = ws_SqlFlds + ", SRmSz2 "
      ws_Cnd = ws_Cnd + " and SRmSz2= '" + CStr(mRmSz2) + "' ": ws_GrpBy = ws_GrpBy + ", SRmSz, SRmSz2 "
      ws_ColHeadg = ws_ColHeadg + "RmSz2        |": ws_ColFlds = ws_ColFlds + "SRmSz2|"
      ws_OrdBy = ws_OrdBy + ", SRmSz2 "
    End If
  End If
  If Not IsMissing(mRmSz3) Then
    If mRmSz3 <> 0 Then
      ws_SqlFlds = ws_SqlFlds + ", SRmSz3 "
      ws_Cnd = ws_Cnd + " and SRmSz3= '" + CStr(mRmSz3) + "' ": ws_GrpBy = ws_GrpBy + ", SRmSz, SRmSz2, SRmSz3 "
      ws_ColHeadg = ws_ColHeadg + "RmSz3        |": ws_ColFlds = ws_ColFlds + "SRmSz3|"
      ws_OrdBy = ws_OrdBy + ", SRmSz3 "
    End If
  End If
  
  With mwhlp
  '*** Jay 3.01 ****** [LotNo Err](ws_Join ADDED)
    .Sqlstrg = " Select distinct SLotNo " + ws_SqlFlds + ", sum(SRmDrQty - SRmCrQty) as qStkQty, " + _
                " Convert(decimal(10, 3), sum(SRmDrWt - SRmCrWt)) as qStkWt " + _
                " From SYyMm " + ws_Join + _
                " Where SCoCd='" + mCoCd + "' and SLotNo <> '' " + ws_Cnd + _
                " Group By SLotNo " + ws_GrpBy + _
                " having Convert(decimal(10, 3), sum(SRmDrWt - SRmCrWt))<> 0 " + _
                " Order By SLotNo " + ws_OrdBy
    .ColHeadg = "|Lot No.                      |" + ws_ColHeadg + "StkQty         |StkWt         |"
    .ColFlds = "|SLotNo|" + ws_ColFlds + "qStkQty|qStkWt|"
    .RetFldNm = "SLotNo"
    .SearchColNum = 1
    .SortAsc = True
  End With
End Sub

Public Sub LotNoJL(ByVal mCoCd As String, Optional ByVal mLoc As Variant, Optional ByVal mRmCd As Variant, _
                                    Optional ByVal mRmSz As Variant, Optional ByVal mRmSz2 As Variant, Optional ByVal mRmSz3 As Variant, Optional ByVal mLocTyp As Variant)
  Dim ws_Cnd As String, ws_GrpBy As String, ws_OrdBy As String
  Dim ws_ColHeadg As String, ws_ColFlds As String
  Dim ws_SqlFlds As String, ws_Join As String
  
  '*** Jay 3.01 **** [LotNo Err]
  ws_SqlFlds = ws_SqlFlds + ", DLoc "
  ws_GrpBy = ws_GrpBy + ", DLoc "
  ws_ColHeadg = ws_ColHeadg + "Loc        |": ws_ColFlds = ws_ColFlds + "DLoc|"
  ws_OrdBy = ws_OrdBy + ", DLoc "
  ws_Join = " "
  
  If Not IsMissing(mLocTyp) Then
    If mLocTyp <> "" Then
      ws_Cnd = ws_Cnd + " and LocTyp = '" + mLocTyp + "' "
      ws_Join = " Join Loc On LocCd= DLoc "
    End If
  End If
  '*** Jay 3.01 **** [LotNo Err]
  If Not IsMissing(mLoc) Then
    If mLoc <> "" Then
      'ws_SqlFlds = ws_SqlFlds + ", SLoc "        '*** Jay 3.01 Bef **** [LotNo Err]
      ws_Cnd = ws_Cnd + " and DLoc= '" + mLoc + "' "
      '*** Jay 3.01 Bef **** [LotNo Err]
      'ws_GrpBy = ws_GrpBy + ", SLoc "
      'ws_ColHeadg = ws_ColHeadg + "Loc        |": ws_ColFlds = ws_ColFlds + "SLoc|"
      'ws_OrdBy = ws_OrdBy + ", SLoc "
      '*** Jay 3.01 Bef **** [LotNo Err]
    End If
  End If
  
  If Not IsMissing(mRmCd) Then
    If mRmCd <> "" Then
      ws_SqlFlds = ws_SqlFlds + ", DJLRmCd "
      ws_Cnd = ws_Cnd + " and DJLRmCd= '" + mRmCd + "' ": ws_GrpBy = ws_GrpBy + ", DJLRmCd "
      ws_ColHeadg = ws_ColHeadg + "RmCd        |": ws_ColFlds = ws_ColFlds + "DJLRmCd|"
      ws_OrdBy = ws_OrdBy + ", DJLRmCd "
    End If
  End If
    
  If Not IsMissing(mRmSz) Then
    If mRmSz <> 0 Then
      ws_SqlFlds = ws_SqlFlds + ", DJLLn1 "
      ws_Cnd = ws_Cnd + " and DJLLn1= " + CStr(mRmSz) + " ": ws_GrpBy = ws_GrpBy + ", DJLLn1 "
      ws_ColHeadg = ws_ColHeadg + "DJLLn1 |": ws_ColFlds = ws_ColFlds + "DJLLn1|"
      ws_OrdBy = ws_OrdBy + ", DJLLn1 "
    End If
  End If
  
  If Not IsMissing(mRmSz2) Then
    If mRmSz2 <> 0 Then
      ws_SqlFlds = ws_SqlFlds + ", DJLLn2 "
      ws_Cnd = ws_Cnd + " and DJLLn2= " + CStr(mRmSz2) + " ": ws_GrpBy = ws_GrpBy + ", DJLLn1, DJLLn2 "
      ws_ColHeadg = ws_ColHeadg + "DJLLn2        |": ws_ColFlds = ws_ColFlds + "DJLLn2|"
      ws_OrdBy = ws_OrdBy + ", DJLLn2 "
    End If
  End If
  If Not IsMissing(mRmSz3) Then
    If mRmSz3 <> 0 Then
      ws_SqlFlds = ws_SqlFlds + ", DJLLn3 "
      ws_Cnd = ws_Cnd + " and DJLLn3= " + CStr(mRmSz3) + " ": ws_GrpBy = ws_GrpBy + ", DJLLn1, DJLLn2, DJLLn3 "
      ws_ColHeadg = ws_ColHeadg + "RmSz3        |": ws_ColFlds = ws_ColFlds + "DJLLn3|"
      ws_OrdBy = ws_OrdBy + ", DJLLn3 "
    End If
  End If
  
  With mwhlp
  '*** Jay 3.01 ****** [LotNo Err](ws_Join ADDED)
    .Sqlstrg = " Select distinct DJLLotNo " + ws_SqlFlds + ", " + _
                " Convert(decimal(10, 3), sum(DDmDrWt - DDmCrWt)) as qStkWt " + _
                " From DmYyMm " + ws_Join + _
                " Where DCoCd='" + mCoCd + "' and DJLLotNo <> '' " + ws_Cnd + _
                " Group By DJLLotNo " + ws_GrpBy + _
                " having Convert(decimal(10, 3), sum(DDmDrWt - DDmCrWt))<> 0 " + _
                " Order By DJLLotNo " + ws_OrdBy
    .ColHeadg = "|Lot No.                  |" + ws_ColHeadg + "StkWt         |"
    .ColFlds = "|DJLLotNo|" + ws_ColFlds + "qStkWt|"
    .RetFldNm = "DJLLotNo"
    .SearchColNum = 1
    .SortAsc = True
  End With
End Sub
' Zubin 212
'*** Jay 3.01 **** [LotNo Err][mLocTyp added]
'SRITODO
' ***** Manali 3.6.0 - 27/10/09 - Stk Rt Help
Public Sub StkRt(ByVal mCoCd As String, Optional ByVal mLoc As Variant, Optional ByVal mRmCd As Variant, _
                                    Optional ByVal mRmSz As Variant, Optional ByVal mRmSz2 As Variant, Optional ByVal mRmSz3 As Variant, Optional ByVal mLocTyp As Variant)
  Dim ws_Cnd As String, ws_GrpBy As String, ws_OrdBy As String
  Dim ws_ColHeadg As String, ws_ColFlds As String
  Dim ws_SqlFlds As String, ws_Join As String
  
  ws_SqlFlds = ws_SqlFlds + ", SLoc "
  ws_GrpBy = ws_GrpBy + ", SLoc "
  ws_ColHeadg = ws_ColHeadg + "Loc        |": ws_ColFlds = ws_ColFlds + "SLoc|"
  ws_OrdBy = ws_OrdBy + ", SLoc "
  ws_Join = " "
  
  If Not IsMissing(mLocTyp) Then
    If mLocTyp <> "" Then
      ws_Cnd = ws_Cnd + " and LocTyp = '" + mLocTyp + "' "
      ws_Join = " Join Loc On LocCd= SLoc "
    End If
  End If
  
  If Not IsMissing(mLoc) Then
    If mLoc <> "" Then
      ws_Cnd = ws_Cnd + " and SLoc= '" + mLoc + "' "
    End If
  End If
  
  If Not IsMissing(mRmCd) Then
    If mRmCd <> "" Then
      ws_SqlFlds = ws_SqlFlds + ", SRmCd "
      ws_Cnd = ws_Cnd + " and SRmCd= '" + mRmCd + "' ": ws_GrpBy = ws_GrpBy + ", SRmCd "
      ws_ColHeadg = ws_ColHeadg + "RmCd        |": ws_ColFlds = ws_ColFlds + "SRmCd|"
      ws_OrdBy = ws_OrdBy + ", SRmCd "
    End If
  End If
    
  If Not IsMissing(mRmSz) Then
    If mRmSz <> 0 Then
      ws_SqlFlds = ws_SqlFlds + ", SRmSz "
      ws_Cnd = ws_Cnd + " and SRmSz= '" + CStr(mRmSz) + "' ": ws_GrpBy = ws_GrpBy + ", SRmSz "
      ws_ColHeadg = ws_ColHeadg + "RmSz        |": ws_ColFlds = ws_ColFlds + "SRmSz|"
      ws_OrdBy = ws_OrdBy + ", SRmSz "
    End If
  End If
  If Not IsMissing(mRmSz2) Then
    If mRmSz2 <> 0 Then
      ws_SqlFlds = ws_SqlFlds + ", SRmSz2 "
      ws_Cnd = ws_Cnd + " and SRmSz2= '" + CStr(mRmSz2) + "' ": ws_GrpBy = ws_GrpBy + ", SRmSz2 "
      ws_ColHeadg = ws_ColHeadg + "RmSz2        |": ws_ColFlds = ws_ColFlds + "SRmSz2|"
      ws_OrdBy = ws_OrdBy + ", SRmSz2 "
    End If
  End If
  If Not IsMissing(mRmSz3) Then
    If mRmSz2 <> 0 Then
      ws_SqlFlds = ws_SqlFlds + ", SRmSz3 "
      ws_Cnd = ws_Cnd + " and SRmSz3= '" + CStr(mRmSz3) + "' ": ws_GrpBy = ws_GrpBy + ", SRmSz3 "
      ws_ColHeadg = ws_ColHeadg + "RmSz3        |": ws_ColFlds = ws_ColFlds + "SRmSz3|"
      ws_OrdBy = ws_OrdBy + ", SRmSz3 "
    End If
  End If
  
  With mwhlp

    .Sqlstrg = " Select distinct SRmStkRt " + ws_SqlFlds + ", sum(SRmDrQty - SRmCrQty) as qStkQty, " + _
                " Convert(decimal(10, 3), sum(SRmDrWt - SRmCrWt)) as qStkWt " + _
                " From SYyMm " + ws_Join + _
                " Where SCoCd='" + mCoCd + "' and SRmStkRt <> 0 " + ws_Cnd + _
                " Group By SRmStkRt " + ws_GrpBy + _
                " having Convert(decimal(10, 3), sum(SRmDrWt - SRmCrWt))<> 0 " + _
                " Order By SRmStkRt " + ws_OrdBy
    .ColHeadg = "|StkRt                  |" + ws_ColHeadg + "StkQty         |StkWt         |"
    .ColFlds = "|SRmStkRt|" + ws_ColFlds + "qStkQty|qStkWt|"
    .RetFldNm = "SRmStkRt"
    .SearchColNum = 1
    .SortAsc = True
  End With
End Sub
' ***** Manali 3.6.0 - 27/10/09 - Stk Rt Help

' Zubin 212
'sritodo
Public Sub BagLotNo(ByVal mCoCd As String, ByVal mBYy As String, ByVal mBChr As String, _
                ByVal mBNo As Single, ByVal mRmCd As String, ByVal mRmSz As Single, ByVal mRmSz2 As Single, ByVal mRmSz3 As Single, ByVal mSeoYn As Boolean)
  
  Dim ws_Cnd As String, ws_GrpBy As String, ws_OrdBy As String
  Dim ws_ColHeadg As String, ws_ColFlds As String
  Dim ws_SqlFlds As String
  
  If Not IsMissing(mRmCd) Then
    If mRmCd <> "" Then
      ws_SqlFlds = ws_SqlFlds + ", TdRmCd "
      ws_Cnd = ws_Cnd + " and TdRmCd= '" + mRmCd + "' ": ws_GrpBy = ws_GrpBy + ", TdRmCd "
      ws_ColHeadg = ws_ColHeadg + "RmCd        |": ws_ColFlds = ws_ColFlds + "TdRmCd|"
      ws_OrdBy = ws_OrdBy + ", TdRmCd "
    End If
  End If
    
  If Not IsMissing(mRmSz) Then
    If mRmSz <> 0 Then
      ws_SqlFlds = ws_SqlFlds + ", TdRmSz "
      ws_Cnd = ws_Cnd + " and TdRmSz= '" + CStr(mRmSz) + "' ": ws_GrpBy = ws_GrpBy + ", TdRmSz "
      ws_ColHeadg = ws_ColHeadg + "RmSz        |": ws_ColFlds = ws_ColFlds + "TdRmSz|"
      ws_OrdBy = ws_OrdBy + ", TdRmSz "
    End If
  End If
  
  '****** Sachin 3.02.0 Partition Key added
  With mwhlp
    .Sqlstrg = "Select " + IIF(gs_Partition = ctCurrPrtn, "", "TdPrtKey,") + " TdLotNo, TdRmStkRt, " + _
                "sum((case (case when TdFrRmLoc='W' then TdFrRmDc " + _
                "when TdToRmLoc='W' then TdToRmDc else '' end) when 'C' then -1 " + _
                "when 'D' then 1 else 0 end)* TdRmQty) as qRmQty, " + _
                "sum((case (case when TdFrRmLoc='W' then TdFrRmDc when TdToRmLoc='W' then " + _
                "TdToRmDc end) when 'C' then -1 when 'D' then 1 else 0 end)* " + _
                "TdRmWt) as qRmWt " + ws_SqlFlds + " From Txnd " + _
                "Where TdRmCd <> '' and TdLotNo <> '' " + _
                IIF(gs_Partition = ctCurrPrtn, " And TdPrtKey='" + ctCurrPrtn + "' ", "") + _
                " and TdCoCd= '" + mCoCd + "' and " + _
                "TdBYy= '" + mBYy + "' and TdBChr= '" + mBChr + "' and " + _
                "TdBNo= " + CStr(mBNo) + ws_Cnd + " " + _
                "Group By TdPrtKey,TdBYy, TdBChr, TdBNo, TdLotNo, TdRmStkRt " + ws_GrpBy + _
                " Order By TdLotNo " + ws_OrdBy + ", TdRmStkRt "
'StkRt-if Seo rights is true stock rate will display else not show stock rate
    .ColHeadg = IIF(gs_Partition <> ctCurrPrtn, "|Prtn ", "") + "|Lot No.                  |" + ws_ColHeadg + IIF(mSeoYn = True, "Stk Rt.   |", "") + "BalQty         |BalWt         |"
    .ColFlds = IIF(gs_Partition <> ctCurrPrtn, "|TdPrtKey", "") + "|TdLotNo|" + ws_ColFlds + IIF(mSeoYn = True, "TdRmStkRt|", "") + "qRmQty|qRmWt|"
    .RetFldNm = "TdLotNo"
    .SearchColNum = IIF(gs_Partition <> ctCurrPrtn, 2, 1)
    .SortAsc = True
  End With
End Sub
'SRITODO
Public Sub BagStkRt(ByVal mCoCd As String, ByVal mBYy As String, ByVal mBChr As String, _
                ByVal mBNo As Single, ByVal mRmCd As String, ByVal mRmSz As Single, ByVal mRmSz2 As Single, ByVal mRmSz3 As Single)
' ***** Manali 3.6.0 - 27/10/09 - Stk Rt Help

  Dim ws_Cnd As String, ws_GrpBy As String, ws_OrdBy As String
  Dim ws_ColHeadg As String, ws_ColFlds As String
  Dim ws_SqlFlds As String
  
  If Not IsMissing(mRmCd) Then
    If mRmCd <> "" Then
      ws_SqlFlds = ws_SqlFlds + ", TdRmCd "
      ws_Cnd = ws_Cnd + " and TdRmCd= '" + mRmCd + "' ": ws_GrpBy = ws_GrpBy + ", TdRmCd "
      ws_ColHeadg = ws_ColHeadg + "RmCd        |": ws_ColFlds = ws_ColFlds + "TdRmCd|"
      ws_OrdBy = ws_OrdBy + ", TdRmCd "
    End If
  End If
    
  If Not IsMissing(mRmSz) Then
    If mRmSz <> 0 Then
      ws_SqlFlds = ws_SqlFlds + ", TdRmSz "
      ws_Cnd = ws_Cnd + " and TdRmSz= '" + CStr(mRmSz) + "' ": ws_GrpBy = ws_GrpBy + ", TdRmSz "
      ws_ColHeadg = ws_ColHeadg + "RmSz        |": ws_ColFlds = ws_ColFlds + "TdRmSz|"
      ws_OrdBy = ws_OrdBy + ", TdRmSz "
    End If
  End If
  
  '****** Sachin 3.02.0 Partition Key added
  With mwhlp
    .Sqlstrg = "Select " + IIF(gs_Partition = ctCurrPrtn, "", "TdPrtKey,") + " TdRmStkRt, TdLotNo, " + _
                "sum((case (case when TdFrRmLoc='W' then TdFrRmDc " + _
                "when TdToRmLoc='W' then TdToRmDc else '' end) when 'C' then -1 " + _
                "when 'D' then 1 else 0 end)* TdRmQty) as qRmQty, " + _
                "sum((case (case when TdFrRmLoc='W' then TdFrRmDc when TdToRmLoc='W' then " + _
                "TdToRmDc end) when 'C' then -1 when 'D' then 1 else 0 end)* " + _
                "TdRmWt) as qRmWt " + ws_SqlFlds + " From Txnd " + _
                "Where TdRmCd <> '' and TdRmStkRt <> 0 " + _
                IIF(gs_Partition = ctCurrPrtn, " And TdPrtKey='" + ctCurrPrtn + "' ", "") + _
                " and TdCoCd= '" + mCoCd + "' and " + _
                "TdBYy= '" + mBYy + "' and TdBChr= '" + mBChr + "' and " + _
                "TdBNo= " + CStr(mBNo) + ws_Cnd + " " + _
                "Group By TdPrtKey,TdBYy, TdBChr, TdBNo, TdRmStkRt, TdLotNo " + ws_GrpBy + _
                " Order By TdRmStkRt " + ws_OrdBy + ", TdLotNo "
    .ColHeadg = IIF(gs_Partition <> ctCurrPrtn, "|Prtn ", "") + "|Stk Rt.                  |" + ws_ColHeadg + "Lot No.   |BalQty         |BalWt         |"
    .ColFlds = IIF(gs_Partition <> ctCurrPrtn, "|TdPrtKey", "") + "|TdRmStkRt|" + ws_ColFlds + "TdLotNo|qRmQty|qRmWt|"
    .RetFldNm = "TdRmStkRt"
    .SearchColNum = IIF(gs_Partition <> ctCurrPrtn, 2, 1)
    .SortAsc = True
  End With
' ***** Manali 3.6.0 - 27/10/09 - Stk Rt Help
End Sub

' ******* Sachin for Customer Design ******* '
Public Sub CustDmCd(ByVal mCmCd As String, Optional ByVal ps_ArrRetOth As Variant)
  '*** (The optional parameter ps_ArrRetOth was introduced in 2.12)
  
  '****** Sachin 3.02.0 - Partition Key added
  wSqlStrg = "SELECT " + IIF(gs_Partition = ctCurrPrtn, "", "CdPrtKey,") + "CdCustDmCd, CdOurDmCd, CdOurSfx, CdOurDmSz From CustDsg " + _
             " Where CdCd= '" + mCmCd + "' " + IIF(gs_Partition = ctCurrPrtn, " and CdPrtKey='" + ctCurrPrtn + "'", "") + " Order By CdCustDmCd, CdOurDmCd, CdOurSfx, CdOurDmSz"

  With mwhlp
    .Sqlstrg = wSqlStrg
    .ColHeadg = IIF(gs_Partition <> ctCurrPrtn, "|Prtn ", "") + "|Cust Dsg.         |Our Dsg.            |Our Sfx.       |Our Size     |"
    .ColFlds = IIF(gs_Partition <> ctCurrPrtn, "|CdPrtKey", "") + "|CdCustDmCd|CdOurDmCd|CdOurSfx|CdOurDmSz|"
    .SearchColNum = 1
    .RetFldNm = "CdCustDmCd"
    .SortAsc = True
    '*** (Jen 2.12)
    If Not IsMissing(ps_ArrRetOth) Then
      .RetOth = ps_ArrRetOth
    End If
    '*** (Jen 2.12)
  End With
  
End Sub
' ******* Sachin for Customer Design ******* '

' **** Zubin (06/12/05) **** '
Public Sub TagNo(ByVal mCoCd As String, ByVal mTc As String, ByVal mYy As String, ByVal mChr As String)
  Dim wSqlStrg
  wSqlStrg = "SELECT TgNo from Tag where TgCoCd= '" + mCoCd + "' " + _
             "and TgTc= '" + mTc + "' and TgYy='" + mYy + "' and " + _
             "TgChr='" + mChr + "' Order By TgNo"
  With mwhlp
    .Sqlstrg = wSqlStrg
    .ColHeadg = "|Number     |"
    .ColFlds = "|TgNo|"
    .SearchColNum = 1
    .RetFldNm = "TgNo"
    .SortAsc = True
  End With
End Sub

'*** (Jen 3.01)
Public Sub CtrlSubTyp(ByVal ms_TcTyp As String, ByVal ms_CtlTyp As String)
  Dim ws_SqlStrg
  
  ws_SqlStrg = "Select distinct UrCtlSTyp as UrCtlSTyp, UrDesc From UsrRight Where UrTcTyp= '" + ms_TcTyp + "' " + _
               "and UrGrp= '' and UrCtlTyp= '" + ms_CtlTyp + "' and UrCtlSTyp<> '' and UrName= '' " + _
               "Order By UrCtlSTyp"
  
  With mwhlp
    .Sqlstrg = ws_SqlStrg
    .ColHeadg = "|Sub Type                              |Description                                           |"
    .ColFlds = "|UrCtlSTyp|UrDesc|"
    .SearchColNum = 1
    .RetFldNm = "UrCtlSTyp"
    .SortAsc = True
  End With
End Sub
Public Sub CtrlName(ByVal ms_TcTyp As String, ByVal ms_CtlTyp As String, ByVal ms_CtlSubTyp As String)
  Dim ws_SqlStrg
  
  ws_SqlStrg = "Select UrName, UrStyle, UrDesc From UsrRight Where UrTcTyp= '" + ms_TcTyp + "' " + _
               "and UrGrp= '' and UrCtlTyp= '" + ms_CtlTyp + "' and UrCtlSTyp= '" + ms_CtlSubTyp + "' " + _
               "and UrName<> '' Order By UrName"
  
  With mwhlp
    .Sqlstrg = ws_SqlStrg
    .ColHeadg = "|Name                                       |Def Right    |Description                                           |"
    .ColFlds = "|UrName|UrStyle|UrDesc|"
    .SearchColNum = 1
    .RetFldNm = "UrName"
    .SortAsc = True
  End With
End Sub
'*** (Jen 3.01)

' ******* Sachin 3.3.0 (28-06-08) - [28. Valid RmSz Master]  ******* '
Public Sub OrLn1(ByVal ms_RmCtg As String, ByVal ms_RmId As String, Optional ByVal ps_ArrRetOth As Variant)
  
  wSqlStrg = " Select VrsLn1, VrsLn2 From VldRmLn " + _
             " Where VrsRmCtg='" + ms_RmCtg + "' And VrsRmId='" + ms_RmId + "' " + _
             " Order By VrsLn1, VrsLn2"
               
  With mwhlp
    .Sqlstrg = wSqlStrg
    .ColHeadg = "|Ln 1     |Ln 2     |"
    .ColFlds = "|VrsLn1|VrsLn2|"
    .SearchColNum = 1
    .RetFldNm = "VrsLn1"
    .SortAsc = True
    '*** (Jen 2.12)
    If Not IsMissing(ps_ArrRetOth) Then _
      .RetOth = ps_ArrRetOth
    '*** (Jen 2.12)
  End With
  
End Sub
' ******* Sachin 3.3.0 (28-06-08) - [28. Valid RmSz Master]  ******* '

Public Sub OdlSrNo(ByVal mTc As String, ByVal mYy As String, ByVal mChr As String, ByVal mNo As Integer, ByVal mSrFor As Integer, Optional ByVal mSrFr As Integer, Optional ByVal mSrTo As Integer)
 Dim wCnd As String
 wSqlStrg = "Select IdKey,OdlTc,OdlYy,OdlChr,OdlNo,OdlSr,OdlSrNo from OrdLn where OdlCoCd ='" + gs_CoCd + "' and " + _
                " OdlTc ='" + mTc + "' and OdlYy= '" + mYy + "' and " + _
                " OdlChr='" + mChr + "' and OdlNo= " + CStr(mNo)
 
 If mSrFor <> 0 Then
    wCnd = "  OdlSr= " + CStr(mSrFor)
 End If
 If mSrFr <> 0 Then
    wCnd = wCnd + IIF(wCnd <> "", " and ", " ") + " OdlSr >=" + CStr(mSrFr)
 End If
 If mSrTo <> 0 Then
    wCnd = wCnd + IIF(wCnd <> "", " and ", " ") + " OdlSr <=" + CStr(mSrTo)
 End If
 wSqlStrg = wSqlStrg + " and " + wCnd
 
 
 With mwhlp
    .Sqlstrg = wSqlStrg
    .ColHeadg = "|Tc  |Yy   |Chr  |OrdNo  |OrdSr  |OdlSrNo  |"
    .ColFlds = "|OdlTc|OdlYy|OdlChr|OdlNo|OdlSr|OdlSrNo|"
    .SearchColNum = 6
    .RetFldNm = "OdlSrNo"
    .SortAsc = True
  End With
    
End Sub

'*** Jay 3.4[DP]
Public Sub DsgPrmCd(ByVal mDpTyp As String, Optional ByVal mbSearchDmCd As Boolean = False)
  Dim wOrdByStr As String
  wOrdByStr = IIF((UCase(mDpTyp) = "COLL" Or UCase(mDpTyp) = "CAT"), "DpCd", "Convert(Int, DpCd)")
  With mwhlp
  If Not mbSearchDmCd Then
    .Sqlstrg = "SELECT DpCd, DpDesc from DsgPrm where DpTyp = '" + mDpTyp + "' and DpDmCd='' order by " + wOrdByStr
    .ColHeadg = "|Code       |Description                             |"
    .ColFlds = "|DpCd|DpDesc|"
    .RetFldNm = "DpCd"
    .SearchColNum = 1
  Else
    .Sqlstrg = "SELECT DpDmCd, DpCd from DsgPrm where DpTyp = '" + mDpTyp + "' and DpDmCd<>'' order by DpDmCd," + wOrdByStr
    .ColHeadg = "|Design Code    |Code       |"
    .ColFlds = "|DpDmCd|DpCd|"
    .RetFldNm = "DpCd"
    .SearchColNum = 1
  End If
  .SortAsc = True
  End With
'*** Jay 3.4[DP]
End Sub
'*** Jay 3.4[DP]


'*** Jay 3.4[DP]
Public Sub PictLst(ByVal mDesignCd As String)
  '#### With '#' demarcation Design Files will be displayed.
  Dim wDmTcTyp As String, wLineDraw As String, wPictPath As String, wDmCtg As String
  wDmTcTyp = "DM": wLineDraw = "3D"
  
  'wPictPath = go_HlpCn.GetFldVal("Select PDesc225 from Param Where PTyp='PCT' and " + _
                 "PMCd='" + wLineDraw + "'") + "\"
   wDmCtg = go_HlpCn.GetFldVal("Select DmCtg from DsgMst where DmTcTyp='" + wDmTcTyp + "' and DmCd='" + mDesignCd + "' " + _
                 "and DmSz= ''" + IIF(gs_Partition = ctCurrPrtn, " and DmPrtKey='" + ctCurrPrtn + "'", ""))
    
'  go_HlpCn.Execute ("Drop Table #TmpPictFileLst;")
'  go_HlpCn.Execute ("select Space(100) as PfFileName,0 as Pfdepth,0 as PfFile Into #TmpPictFileLst;" + _
           " Insert Into #TmpPictFileLst " + _
           " Exec xp_dirtree '" + wPictPath + "', 0, 1")
  With mwhlp
    ' Bef 3.5.0
    '''.Sqlstrg = "SELECT REPLACE(SUBSTRING(PfFileName,7, Len(PfFileName)),'.jpg','') as PfFileName FROM #TmpPictFileLst WHERE pDmCtg='" + UCase(wDmCtg) + "' and PfFileName LIKE '" + wDmTcTyp + " " + wLineDraw + " " + mDesignCd + "%.jpg' Order By PfFileName"
    ' **** Manali 3.5.0 - 09/06/09 - Pict File Names from PictFileLst (Permanent Table)
    .Sqlstrg = "SELECT REPLACE(SUBSTRING(PfFileName,7, Len(PfFileName)),'.jpg','') as PfFileName FROM PictFileLst WHERE PfDmCtg='" + UCase(wDmCtg) + "' and PfFileName LIKE '" + wDmTcTyp + " " + wLineDraw + " " + mDesignCd + "%.jpg' Order By PfFileName"
    .ColHeadg = "|Design Code                       |"
    .ColFlds = "|PfFileName|"
    .RetFldNm = "PfFileName"
    .SearchColNum = 1
    .SortAsc = True
  End With

                            
'All folders and files in the folder (& not sub folders)
'-- exec xp_dirtree 'D:\MwPict\BRh', 1, 1
'  go_HlpCn.Execute ("Truncate Table PictFileLst " + _
'                    "Insert PictFileLst " + _
'                    "Exec xp_dirtree '" + wPictPath + "', 1, 1")
'  Debug.Print go_HlpCn.GetFldVal("Select Count(*) From PictFileLst ")
'  With mwhlp
'    .Sqlstrg = "SELECT SUBSTRING(PfFileName,7, Len(PfFileName) - 10) as PfFileName FROM PictFileLst WHERE PfFile= 1 and PfFileName LIKE '" + wDmTcTyp + " " + wLineDraw + " " + mDesignCd + "#%.jpg'"
'    .ColHeadg = "|Design Code                       |"
'    .ColFlds = "|PfFileName|"
'    .RetFldNm = "PfFileName"
'    .SearchColNum = 1
'    .SortAsc = True
'  End With
'*** Jay 3.4[DP]
End Sub
Public Sub BaseRefVchJL(ByVal ps_cocd As String, ByVal mBaseRef As String, ByVal mOmTC As String, ByVal mJLRmCd As String, ByVal mJLLotNo As String, _
                      ByVal mJLLn1 As Double, ByVal mJLLn2 As Double, ByVal mJLLn3 As Double, Optional ByVal mCmCtg As String, Optional ByVal mCmCd As String, Optional ByVal mBaseCocd As String = "")
Dim ws_SqlStrg As String, ws_BaseRefTc As String

Dim wsCnd As String
wsCnd = wsCnd + " and OdJLRmCd ='" + mJLRmCd + "'"

ws_BaseRefTc = go_HlpCn.GetFldVal("Select " + IIF(mBaseRef = "B", "TmBaseTcs", "TmRefTcs") + " from TcMst Where TmCoCd='" + ps_cocd + "' and TmTc='" + mOmTC + "'")
    
wsCnd = wsCnd + IIF(ws_BaseRefTc = "LSO" Or ws_BaseRefTc = "LPO", "", " and OdJLLotNo='" + mJLLotNo + "'")
wsCnd = wsCnd + "  and OdJLLn1=" + CStr(mJLLn1) + " and OdJLLn2= " + CStr(mJLLn2) + " and OdJLLn3 = " + CStr(mJLLn3)

If mBaseRef = "B" Then
  ws_SqlStrg = "Select OdTc+'/'+OdYy+'/'+OdChr+'/'+Ltrim(Str(OdNo))+'/'+Ltrim(Str(OdSr)) as qBaseRefVch,OdJLRmCd, OdJLLotNo, OdJLLn1, ODJLLn2,OdJLLn3, CAST((OdJLWt-OdJLExpWt) AS DECIMAL(10,3)) as qBalWt " + _
                " from OrdDsg Join OrdMst on OmCoCd=OdCoCd and OmTc=OdTc and OmYy=OdYy and OmChr=OdChr and OmNo=OdNo and OmPrtKey=OdPrtKey and OmIdNo=OdOmIdNo " + _
                " Where OdCoCd=" + IIF(mOmTC = "LIN", "'" + mBaseCocd + "'", "'" + ps_cocd + "'") + "  and OmCmCtg ='" + mCmCtg + "' and OmCmCd='" + mCmCd + "' " + _
                " and OdTc='" + ws_BaseRefTc + "' and OdDmCd='" + mDmCd + "' " + wsCnd + _
                " and OmCmCtg ='" + mCmCtg + "' and OmCmCd='" + mCmCd + "' and CAST((OdJLWt-OdJLExpWt) AS DECIMAL(10,3))> 0 " + _
                "Order By OdJLRmCd, OdJLLotNo, OdJLLn1, OdJLLn2,OdJLLn3"
ElseIf mBaseRef = "R" Then
  ws_SqlStrg = "Select OdTc+'/'+OdYy+'/'+OdChr+'/'+Ltrim(Str(OdNo))+'/'+Ltrim(Str(OdSr)) as qBaseRefVch, OdJLRmCd, OdJLLotNo, OdJLLn1, OdJLLn2, OdJLLn3,CAST((OdJLWt-OdJLExpWt) AS DECIMAL(10,3)) as qBalWt " + _
                " from OrdDsg Where OdCoCd='" + ps_cocd + "' and '," + ws_BaseRefTc + ",' like '%,'+OdTc+',%' and OdJLRmCd='" + mJLRmCd + "' " + _
                wsCnd + " Order By OdJLRmCd, OdJLLotNo, OdJLLn1, OdJLLn2, OdJLLn3"
End If

  With mwhlp
    .Sqlstrg = ws_SqlStrg
    '578.13 Rmcd length increased
    .ColHeadg = "|Base Vch Number          |RmCd                                  |LotNo             |Ln1      |Ln2      |Ln3      |Bal Wt     |"
    .ColFlds = "|qBaseRefVch|OdJLRmCd|OdJLLotNo|OdJLLn1|OdJLLn2|OdJLLn3|qBalWt|"
    .RetFldNm = "qBaseRefVch"
    .SearchColNum = 1
    .SortAsc = True
  End With
End Sub
Public Sub BaseRefVch(ByVal ps_cocd As String, ByVal mBaseRef As String, ByVal mOmTC As String, ByVal mDmCd As String, ByVal mDmSfx As String, _
                      ByVal mDmSz As String, ByVal mDmYy As String, ByVal mDmChr As String, ByVal mDmNo As String, Optional ByVal mCmCtg As String, Optional ByVal mCmCd As String, Optional ByVal mBaseCocd As String = "")
Dim ws_SqlStrg As String, ws_BaseRefTc As String

Dim wsCnd As String
wsCnd = IIF(mDmSfx <> "", " and OdSfx='" + mDmSfx + "' ", "")
wsCnd = wsCnd + " and OdDmSz='" + mDmSz + "' " ' IIF(mDmSz <> "", " and OdDmSz='" + mDmSz + "' ", "")

ws_BaseRefTc = go_HlpCn.GetFldVal("Select " + IIF(mBaseRef = "B", "TmBaseTcs", "TmRefTcs") + " from TcMst Where TmCoCd='" + ps_cocd + "' and TmTc='" + mOmTC + "'")
    
wsCnd = wsCnd + IIF(mDmId = 0 Or ws_BaseRefTc = "JSO" Or ws_BaseRefTc = "JPO", "", " and OdDmStkYy='" + mDmYy + "' OdDmStkChr='" + mDmChr + "' and OdDmStkNo = " + mDmNo)

If mBaseRef = "B" Then
  ' Conditional 1=1 below is by Sri - error was JPR was not showing help of JPO ... it turns out OmCmCd was compared to CoCd in other Tcs except JIN
  ' Original line was 'IIF (OmTc = 'JIN'...)
  ws_SqlStrg = "Select OmCmCd, OdTc+'/'+OdYy+'/'+OdChr+'/'+Ltrim(Str(OdNo))+'/'+Ltrim(Str(OdSr)) as qBaseRefVch, OdDmCd as DmCd, OdSfx as DmSfx, OdDmSz as DmSz, OdDmStkYy as DmStkYy, OdDmStkChr as DmStkChr, OdDmStkNo as DmStkNo, OdOrdQty-OdExpQty as qBalQty " + _
                " from OrdDsg Join OrdMst on OmCoCd=OdCoCd and OmTc=OdTc and OmYy=OdYy and OmChr=OdChr and OmNo=OdNo and OmPrtKey=OdPrtKey and OmIdNo=OdOmIdNo " + _
                " Where OdCoCd='" + IIF(mOmTC = "JIN" Or mOmTC = "LIN", mBaseCocd, ps_cocd) + "'  and OmCmCtg ='" + mCmCtg + "'" + _
                " and OdTc='" + ws_BaseRefTc + "' and OdDmCd='" + mDmCd + "' " + wsCnd + _
                " and OmCmCtg ='" + mCmCtg + "'" + IIF(mCmCd <> "", " and OmCmCd='" + mCmCd + "'", "") + " and Convert(Decimal(14,1), OdOrdQty-OdExpQty)> 0.0 " + _
                "Order By OdDmCd, OdSfx, OdDmSz, OdDmStkYy,OdDmStkChr, OdDmStkNo"
                
                'and OdSfx='" + mDmSfx + "' and OdDmSz='" + mDmSz + "' " + _
                IIF(mDmId = 0, "", " and OdDmId In (0, " + CStr(mDmId) + ") ") + _
                " and OmCmCtg ='" + mCmCtg + "' and OmCmCd='" + mCmCd + "' "
ElseIf mBaseRef = "R" Then
  ws_SqlStrg = "Select OmCmCd, OdTc+'/'+OdYy+'/'+OdChr+'/'+Ltrim(Str(OdNo))+'/'+Ltrim(Str(OdSr)) as qBaseRefVch, OdDmCd as DmCd, OdSfx as DmSfx, OdDmSz as DmSz, OdDmStkYy as DmStkYy, OdDmStkChr as DmStkChr, OdDmStkNo As DmStkNo, OdOrdQty-OdExpQty as qBalQty " + _
                " from OrdDsg join OrdMst on OmCoCd=OdCoCd and OmTc=OdTc and OmYy=OdYy and OmChr=OdChr and OmNo=OdNo and OmPrtKey=OdPrtKey and OmIdNo=OdOmIdNo Where OdCoCd='" + ps_cocd + "' and '," + ws_BaseRefTc + ",' like '%,'+OdTc+',%' and OdDmCd='" + mDmCd + "' " + _
                wsCnd + " Order By OdDmCd, OdSfx, OdDmSz, OdDmStkYy, OdDmStkChr, OdDmStkNo"
                
                'and OdSfx='" + mDmSfx + "' and OdDmSz='" + mDmSz + "' "
                'IIF(mDmId = 0, "", " and OdDmId In (0, " + CStr(mDmId) + ") ")
End If

  With mwhlp
    .Sqlstrg = ws_SqlStrg
    .ColHeadg = "|Base Vch Number          |Design                |Suffix                                   |Size     |Yy|DmChr|DmNo  |Bal Qty  |Customer     |"
    .ColFlds = "|qBaseRefVch|DmCd|DmSfx|DmSz|DmStkYy|DmStkChr|DmStkNo|qBalQty|OmCmCd|"
    .RetFldNm = "qBaseRefVch"
    .SearchColNum = 1
    .SortAsc = True
  End With

End Sub

          
          
Public Sub BagNo(ByVal mCoCd As String, ByVal mhOthLocBag As String, Optional ByVal mLoc As Variant, Optional ByVal mOdTc As Variant, Optional ByVal mOdYy As Variant, Optional ByVal mOdChr As Variant, Optional ByVal mOdNo As Variant, Optional ByVal mOdDmCd As Variant, Optional ByVal mSameDesignBags As Boolean = False, Optional ByVal mPrtKey As String, Optional ByVal mBCls As String)
  Dim wOrdCnd1 As String, wOrdCnd2 As String, wLocStr As String, wSameDsgBag As String, wsBCls As String
  
  If mBCls <> "" Then
    wsBCls = " and BCls='" + mBCls + "' "
  Else
    wsBCls = ""
  End If
  
  If Not (IsMissing(mOdTc) Or IsMissing(mOdYy) Or IsMissing(mOdChr) Or IsMissing(mOdNo) Or IsMissing(mOdDmCd)) Then
    wOrdCnd1 = " and BCoCd='" + mCoCd + "' and BOdTc= '" + mOdTc + "' and" + _
               " BOdYy= '" + mOdYy + "' and BOdChr= '" + mOdChr + "' and" + _
               " BOdNo= " + CStr(mOdNo) + " and BOdDmCd= '" + mOdDmCd + "' "
    wOrdCnd2 = " and not (BCoCd='" + mCoCd + "' and BOdTc= '" + mOdTc + "' and" + _
               " BOdYy= '" + mOdYy + "' and BOdChr= '" + mOdChr + "' and" + _
               " BOdNo= " + CStr(mOdNo) + " and BOdDmCd= '" + mOdDmCd + "') "
    If Not IsMissing(mLoc) And mhOthLocBag = "N" Then
      wLocStr = IIF(mLoc = "", " ", " and BLoc= '" + mLoc + "' ")
    Else
      wLocStr = " "
    End If
    wSqlStrg = "SELECT 'A' as qTag, BPrtKey, BYy+'/'+BChr+'/'+LTrim(RTrim(str(BNo))) as qBagNo, BOdDmCd, BOdSfx, BOdDmSz, " + _
               "BOdTc+'/'+BOdYy+'/'+BOdChr+'/'+LTrim(Str(BOdNo))+'/'+LTrim(Str(BOdSr)) " + _
               "as qOrdNo, BLoc, BQty, BFgSubLoc from Bag " + _
               "where BCoCd='" + mCoCd + "' " + _
               IIF(gs_Partition = ctCurrPrtn, " And BPrtKey='" + ctCurrPrtn + "'", "") + _
               wsBCls + wLocStr + wOrdCnd1 + _
               " Union All " + _
               "SELECT 'B' as qTag, BPrtKey, BYy+'/'+BChr+'/'+LTrim(RTrim(str(BNo))) as qBagNo, BOdDmCd, BOdSfx, BOdDmSz, " + _
               "BOdTc+'/'+BOdYy+'/'+BOdChr+'/'+LTrim(Str(BOdNo))+'/'+LTrim(Str(BOdSr)) " + _
               "as qOrdNo, BLoc, BQty, BFgSubLoc from Bag " + _
               "where " + _
               IIF(gs_Partition = ctCurrPrtn, " BPrtKey='" + ctCurrPrtn + "' And ", "") + _
               " BCoCd='" + mCoCd + "' " + _
               wsBCls + wLocStr + wOrdCnd2
    wSqlStrg = wSqlStrg + " Order By qTag, qBagNo Desc, BQty "
  Else
    If Not IsMissing(mLoc) Then
      wLocStr = " and BLoc= '" + mLoc + "' "
      wSameDsgBag = IIF(mSameDesignBags = True And (Not IsMissing(mOdDmCd)), " and BOdDmCd = '" + IIF(Not IsMissing(mOdDmCd), mOdDmCd, "") + "'", " ")
      wSqlStrg = "SELECT 'A' as qTag, BPrtKey, BYy+'/'+BChr+'/'+LTrim(RTrim(str(BNo))) as qBagNo, BOdDmCd, BOdSfx, BOdDmSz, " + _
                 "BOdTc+'/'+BOdYy+'/'+BOdChr+'/'+LTrim(Str(BOdNo))+'/'+LTrim(Str(BOdSr)) " + _
                 "as qOrdNo, BLoc, BQty, BFgSubLoc from Bag " + _
                 "where " + _
                 IIF(gs_Partition = ctCurrPrtn Or mPrtKey = ctCurrPrtn, " BPrtKey='" + ctCurrPrtn + "' And ", "") + _
                 "BCoCd='" + mCoCd + "' " + _
                 wsBCls + wLocStr + wSameDsgBag
      
      If mhOthLocBag = "Y" Then
        wStr = " Union All " + _
               "SELECT 'B' as qTag, BPRtKey, BYy+'/'+BChr+'/'+LTrim(RTrim(str(BNo))) as qBagNo, BOdDmCd, BOdSfx, BOdDmSz, " + _
               "BOdTc+'/'+BOdYy+'/'+BOdChr+'/'+LTrim(Str(BOdNo))+'/'+LTrim(Str(BOdSr)) " + _
               "as qOrdNo, BLoc, BQty, BFgSubLoc from Bag " + _
               "where " + _
               IIF(gs_Partition = ctCurrPrtn Or mPrtKey = ctCurrPrtn, " BPrtKey='" + ctCurrPrtn + "' And ", "") + _
               "BCoCd='" + mCoCd + "' " + _
               wsBCls + " and BLoc<> '" + mLoc + "' " + wSameDsgBag
      Else
        wStr = " "
      End If
      wSqlStrg = wSqlStrg + wStr + " Order By qTag, qBagNo Desc"
    Else
      wSqlStrg = "SELECT BPrtKey, BYy+'/'+BChr+'/'+LTrim(RTrim(str(BNo))) as qBagNo, BOdDmCd, BOdSfx, BOdDmSz, " + _
                 "BOdTc+'/'+BOdYy+'/'+BOdChr+'/'+LTrim(Str(BOdNo))+'/'+LTrim(Str(BOdSr)) " + _
                 "as qOrdNo, BLoc, BQty, BFgSubLoc from Bag " + _
                 "where " + _
                 IIF(gs_Partition = ctCurrPrtn Or mPrtKey = ctCurrPrtn, " BPrtKey='" + ctCurrPrtn + "' And ", "") + _
                 "BCoCd='" + mCoCd + "' " + _
                 wsBCls + " Order By BYy+'/'+BChr+'/'+LTrim(RTrim(str(BNo))) Desc"
    End If
  End If
  With mwhlp
    .Sqlstrg = wSqlStrg
    .ColHeadg = "|Bag Number      |" + IIF(gs_Partition <> ctCurrPrtn, "Prtn |", "") + "Design           |Suffix                                   |Size |Order                                  |Loc    |FgSubLoc    |Qty     |"
    .ColFlds = "|qBagNo|" + IIF(gs_Partition <> ctCurrPrtn, "BPrtKey|", "") + "BOdDmCd|BOdSfx|BOdDmSz|qOrdNo|BLoc|BFgSubLoc|BQty|"
    .SearchColNum = 1
    .RetFldNm = "qBagNo"
    .SortAsc = False
  End With
End Sub



Public Sub TmTc(ByVal mCoCd As String, Optional ByVal mCmCtg As String)
Dim ws_SqlStrg As String, wsCnd As String

wsCnd = IIF(mCmCtg <> "", " and TmCmCtg='" + mCmCtg + "' ", "")

ws_SqlStrg = "Select TmTc, TmDesc from TcMst Where TmCoCd='" + mCoCd + "' " + wsCnd

  With mwhlp
    .Sqlstrg = ws_SqlStrg
    .ColHeadg = "|Tc       |TmDesc                               |"
    .ColFlds = "|TmTc|TmDesc|"
    .RetFldNm = "TmTc"
    .SearchColNum = 1
    .SortAsc = True
  End With

End Sub


' ***** Manali 3.9.1 - RfId Master - TagNo Help
Public Sub RimNo()
Dim ws_SqlStrg As String

ws_SqlStrg = "Select RimNo, RimBCoCd, (Case When RimBNo<>0 Then RimBYy+'/'+RimBChr+'/'+Cast(RimBNo as Varchar(7)) Else '' End) as qBagNo, RimDmCd from RfIdMst Order By  RimNo "

  With mwhlp
    .Sqlstrg = ws_SqlStrg
    .ColHeadg = "|Tag No  |Co Cd  |Bag No            |Dsg Cd          |"
    .ColFlds = "|RimNo|RimBCoCd|qBagNo|RimDmCd|"
    .RetFldNm = "RimNo"
    .SearchColNum = 1
    .SortAsc = False
  End With

End Sub


' ***** Manali 3.10.0
Public Sub MultiPrcQtNo(ByVal mCoCd As String, ByVal mTc As String, ByVal mYy As String, ByVal mChr As String, Optional mCurCd As Variant, Optional ByVal mPrtKey As String)

  Dim ws_Cnd As String
  ws_Cnd = IIF(IsMissing(mCurCd) = True, "", " and OmCmCurCd= '" + CStr(mCurCd) + "' ")
  wSqlStrg = "SELECT MqmNo,MqmDt,MqmCmCd, CmCurCd From MultiPrcQtMst " + _
              "Join CustMst on CmCtg='C' and CmCd=MqmCmCd " + _
             "where " + _
             IIF(gs_Partition = ctCurrPrtn Or mPrtKey = ctCurrPrtn, " MqmPrtKey= '" + ctCurrPrtn + "' And ", "") + _
             " MqmCoCd= '" + mCoCd + "' And " + _
             "MqmTc= '" + mTc + "' and MqmYy='" + mYy + "' and " + _
             "MqmChr='" + mChr + "' " + _
             ws_Cnd + " Order By MqmNo Desc"

  With mwhlp
    .Sqlstrg = wSqlStrg
    .ColHeadg = "|Vch No  |Customer        |Date         |Currency|"
    .ColFlds = "|MqmNo|MqmCmCd|MqmDt|CmCurCd|"
    .ColMask = "|||" + DtShortStr() + "||"    '6.1
    .RetFldNm = "MqmNo"
    .SearchColNum = 1
    .SortAsc = False
  End With

End Sub

Public Sub MultiPrcQtSr(ByVal mCoCd As String, ByVal mTc As String, ByVal mYy As String, ByVal mChr As String, ByVal mNo As Single, Optional ByVal mPrtKey As String)
   wSqlStrg = "SELECT MqdSr, MqdDmCd, MqdDmSz, MqdMetGrd " + _
             "From MultiPrcQtDsg where MqdCoCd= '" + mCoCd + "' and " + _
             "MqdTc= '" + mTc + "' and MqdYy='" + mYy + "' and " + _
             "MqdChr='" + mChr + "' and MqdNo = " + CStr(mNo) + _
             IIF(gs_Partition = ctCurrPrtn Or mPrtKey = ctCurrPrtn, " and MqdPrtKey='" + ctCurrPrtn + "'", "") + _
             " Order By MqdSr"
  With mwhlp
    .Sqlstrg = wSqlStrg
    .ColHeadg = "|Sr.     |Design                          |Size |KT          |"
    .ColFlds = "|MqdSr|MqdDmCd|MqdDmSz|MqdMetGrd|"
    .ColMask = "|||||"
    .SearchColNum = 1
    .RetFldNm = "MqdSr"
    .SortAsc = True
  End With
End Sub

Public Sub OrdNo(ByVal mCoCd As String, ByVal mOmTC As String, Optional ByVal mCmCd As String)
Dim ws_SqlStrg As String

Dim wsCnd As String

  ws_SqlStrg = "Select OmTc+'/'+OmYy+'/'+OmChr+'/'+Ltrim(Str(OmNo)) as qOrdNo, OmCmCd, OmDt, OmCmCurCd, OmPoNo " + _
              " From OrdMst " + _
              " Where OmCoCd='" + mCoCd + "'   and OmTc='" + mOmTC + "' " + _
              IIF(mCmCd <> "", " and OmCmCd='" + mCmCd + "' ", "") + _
              "Order By OmDt Desc, OmNo Desc "

  With mwhlp
    .Sqlstrg = ws_SqlStrg
    .ColHeadg = "|Ord No                     |Date     |Customer       |Currency|PO No                  |"
    .ColFlds = "|qOrdNo|OmCmCd|OmDt|OmCmCurCd|OmPoNo|"
    .ColMask = "|||" + DtShortStr() + "|||"   '6.1
    .RetFldNm = "qOrdNo"
    .SearchColNum = 1
    .SortAsc = False
  End With
End Sub

' ******* Sachin for Customer RM Code - 4.1.0.0 ******* '
Public Sub CustRmCd(ByVal pCmCd As String, Optional ByVal pRmCtg As String, Optional ByVal pRmSCtg As String, Optional ByVal pSvLn As Double, Optional ByVal ps_ArrRetOth As Variant)
  '
  Dim wRmDetsCnd, wSvLnCnd As String
  wRmDetsCnd = IIF(IsMissing(pRmCtg) Or pRmCtg = "", "", " And CrRmCtg='" + pRmCtg + "' ")
  wRmDetsCnd = wRmDetsCnd + IIF(IsMissing(pRmSCtg) Or pRmSCtg = "", "", " And RmSCtg='" + pRmSCtg + "' ")
  wSvLnCnd = IIF(IsMissing(pSvLn) Or pSvLn = 0, "", " And " + CStr(pSvLn) + ">=CrFrLn And " + CStr(pSvLn) + "<=CrToLn ")
  
  wSqlStrg = " Select CrRmCtg, RmSCtg, CrCustRmCd, CrOurRmCd, CrFrLn, CrToLn " + IIF(gs_Partition = ctCurrPrtn, "", ",CrPrtKey ") + _
             " From CustRm " + _
             "     Join RmMst On RmCd=CrOurRmCd " + _
             " Where CrCd= '" + pCmCd + "' " + wRmDetsCnd + wSvLnCnd + IIF(gs_Partition = ctCurrPrtn, " and CrPrtKey='" + ctCurrPrtn + "'", "") + _
             " Order By CrRmCtg, RmSCtg, CrCustRmCd, CrOurRmCd, CrFrLn, CrToLn "
            
  With mwhlp
    .Sqlstrg = wSqlStrg
    .ColHeadg = "|Rm Ctg|Cust RM Cd.             |Our RM Cd.               |Fr Ln.       |To Ln.     |"
    .ColFlds = "|CrRmCtg|CrCustRmCd|CrOurRmCd|CrFrLn|CrToLn|"
    .SearchColNum = 1
    .RetFldNm = "CrCustRmCd"
    .SortAsc = True
    If Not IsMissing(ps_ArrRetOth) Then _
      .RetOth = ps_ArrRetOth
  End With
End Sub
' ******* Sachin for Customer RM Code ******* '
'vk.5- help for RmLotno
'Public Sub RmLotNo(ByVal mCoCd As String, Optional ByVal mRmCtg As Variant)
'  Dim ws_Cnd As String, ws_GrpBy As String, ws_OrdBy As String
'  Dim ws_ColHeadg As String, ws_ColFlds As String
'  Dim ws_SqlFlds As String, ws_Join As String
'
'  ws_SqlFlds = ws_SqlFlds + ", SLoc "
'  ws_ColHeadg = ws_ColHeadg + "Loc        |": ws_ColFlds = ws_ColFlds + "SLoc|"
'  ws_OrdBy = ws_OrdBy + ", SLoc "
'  ws_Join = " "
'
'  With mwhlp
'    .Sqlstrg = " Select distinct SLotNo " + ws_SqlFlds + ", sum(SRmDrQty - SRmCrQty) as qStkQty, " + _
'                " Convert(decimal(10, 3), sum(SRmDrWt - SRmCrWt)) as qStkWt " + _
'                " From SYyMm " + _
'                " Join RmMst On RmCd= SRmCd and RmCtg='" + mRmCtg + "' " + _
'                " Where SCoCd='" + mCoCd + "' and SLotNo <> '' " + ws_Cnd + _
'                " Group By SLotNo , SLoc" + _
'                " having Convert(decimal(10, 3), sum(SRmDrWt - SRmCrWt))<> 0 " + _
'                " Order By SLotNo " + ws_OrdBy
'    .ColHeadg = "|Lot No.                  |" + ws_ColHeadg + "StkQty         |StkWt         |"
'    .ColFlds = "|SLotNo|" + ws_ColFlds + "qStkQty|qStkWt|"
'    .RetFldNm = "SLotNo"
'    .SearchColNum = 1
'    .SortAsc = True
'  End With
'End Sub

Public Sub RmLotNo(ByVal mRmCtg As String)

  With mwhlp
     .Sqlstrg = " Select distinct SLotNo " + _
                " From SYyMm " + _
                " Join RmMst On RmCd= SRmCd and RmCtg='" + mRmCtg + "' " + _
                " Where SCoCd='" + gs_CoCd + "' and SLotNo <> ''   "

    .ColHeadg = "|Lot No.                        |"
    .ColFlds = "|SLotNo|"
    .RetFldNm = "SLotNo"
    .SearchColNum = 1
    .SortAsc = True
  End With
End Sub
'84
Public Sub CertRmLotNo(ByVal mRmCtg As String)
  With mwhlp
     .Sqlstrg = " Select CdRmLotNo, CdRmCd, CdFMBarCd, CdRmQty, CdRmWt from CertDet " + _
                " Where CdCtg='L' and CdRmCtg='" + mRmCtg + "' and CdRmLotNo <>'' " + _
                " Order By CdRmLotNo "
    '578.13 Rmcd length increased
    .ColHeadg = "|RmLotNo                              |RmCd                                  |RmFMCode                                                           |RmQty  |RmWt    |"
    .ColFlds = "|CdRmLotNo|CdRmCd|CdFMBarCd|CdRmQty|CdRmWt|"
    .RetFldNm = "CdRmLotNo"
    .SearchColNum = 1
    .SortAsc = True
  End With
 
End Sub
Public Sub OdNoJNo(ByVal mOdCoCd, ByVal mOdTc As String, ByVal mOdYy As String, ByVal mOdChr As String, _
ByVal mOdDmCd As String, ByVal mOdSfx As String, ByVal mOdDmSz As String, ByVal mOdDmStkCoCd As String, _
ByVal mOdDmStkYy As String, ByVal mOdDmStkChr As String, ByVal mOdDmStkNo As Single)
  With mwhlp
    .Sqlstrg = "Select OdCoCd, OdNo,OmDt,OmDelDt,OmPrtKey,OmCmCd,OmPoNo,OmCmCurCd,OmPrtCd from OrdDsg " + _
               "Join OrdMst on OdOmIdNo=OmIdNo and OdCoCd= OmCoCd and " + _
               "OdTc = OmTc and OdYy= OmYy and OdChr = OmChr and OdNo = OmNo " + _
               "where " + IIF(gs_Partition = ctCurrPrtn, " OdPrtKey= '" + ctCurrPrtn + "' And ", "") + _
               "OdCoCd='" + mOdCoCd + "' and OdTc='" + mOdTc + "' and OdYy='" + mOdYy + "' and OdChr='" + mOdChr + "' and " + _
               "OdDmCd='" + mOdDmCd + "' and OdSfx='" + mOdSfx + "' and OdDmSz='" + mOdDmSz + "' and " + _
               "OdDmStkCoCd='" + mOdDmStkCoCd + "' and OdDmStkYy='" + mOdDmStkYy + "' and " + _
               "OdDmStkChr='" + mOdDmStkChr + "' and OdDmStkNo=" + CStr(mOdDmStkNo) + _
               " Order by OdNo Desc "
    
    .ColHeadg = "|OrdNo  |Date        |Del Date    |" + IIF(gs_Partition <> ctCurrPrtn, "Prtn |", "") + "Customer       |Po No                                          |Currency   |Priority |)"
    .ColFlds = "|OdNo|OmDt|OmDelDt|" + IIF(gs_Partition <> ctCurrPrtn, "OmPrtKey|", "") + "OmCmCd|OmPoNo|OmCmCurCd|OmPrtCd|"
    .RetFldNm = "OdNo"
    .SearchColNum = 1
    .SortAsc = True
        
  End With
End Sub
'KC.8 new help added
Public Sub OdSrJNo(ByVal mOdCoCd As String, ByVal mOdTc As String, ByVal mOdYy As String, ByVal mOdChr As String, ByVal mOdNo As Single, _
ByVal mOdDmCd As String, ByVal mOdSfx As String, ByVal mOdDmSz As String, ByVal mOdDmStkCoCd As String, _
ByVal mOdDmStkYy As String, ByVal mOdDmStkChr As String, ByVal mOdDmStkNo As Single)
  With mwhlp
    .Sqlstrg = "SELECT " + IIF(gs_Partition = ctCurrPrtn, "", "OdPrtKey,") + "OdSr, OdDmCd, OdSfx, OdDmSz " + IIF(mShowPrdQty = True, ", OdPrdQty ", "") + ", OdPoNo from OrdDsg " + _
               "Join OrdMst on OdOmIdNo=OmIdNo and OdCoCd= OmCoCd and " + _
               "OdTc = OmTc and OdYy= OmYy and OdChr = OmChr and OdNo = OmNo " + _
               "where " + IIF(gs_Partition = ctCurrPrtn, " OdPrtKey= '" + ctCurrPrtn + "' And ", "") + _
               "OdCoCd='" + mOdCoCd + "' and OdTc='" + mOdTc + "' and OdYy='" + mOdYy + "' " + _
               "and OdChr='" + mOdChr + "' and OdNo = " + CStr(mOdNo) + " and " + _
               "OdDmCd='" + mOdDmCd + "' and OdSfx='" + mOdSfx + "' and OdDmSz='" + mOdDmSz + "' and " + _
               "OdDmStkCoCd='" + mOdDmStkCoCd + "' and OdDmStkYy='" + mOdDmStkYy + "' and " + _
               "OdDmStkChr='" + mOdDmStkChr + "' and OdDmStkNo=" + CStr(mOdDmStkNo) + _
               " Order by OdSr "
    .ColHeadg = "|Sr.     |Design                          |Suffix                                   |Size |SubPO               |" + IIF(gs_Partition <> ctCurrPrtn, "Prtn |", "") + IIF(mShowPrdQty = True, "Prd Qty   |", "")
    .ColFlds = "|OdSr|OdDmCd|OdSfx|OdDmSz|OdPONo|" + IIF(gs_Partition <> ctCurrPrtn, "OdPrtKey|", "") + IIF(mShowPrdQty = True, "OdPrdQty|", "")
    .ColMask = "||||||" + IIF(gs_Partition <> ctCurrPrtn, "|", "") + IIF(mShowPrdQty = True, "|", "")
    .SearchColNum = 1
    .RetFldNm = "OdSr"
    .SortAsc = True
  End With
  
End Sub

'578-20 calling help of customer store code
Public Sub CustStoreCd(ByVal mCmCtg As String, ByVal mCmCd As String)
  
  wSqlStrg = "Select CstCd ,CstName,CstAdd1+';'+CstAdd2+';'+CstAdd3 as qCstAdd from CustStore " + _
             " Where CstCmCtg= '" + mCmCtg + "' and CstCmCd= '" + mCmCd + "' " + _
             IIF(gs_Partition = ctCurrPrtn, " and CstPrtKey='" + ctCurrPrtn + "'", "") + " Order By CstCd "
      
  With mwhlp
    .Sqlstrg = wSqlStrg
    .ColHeadg = "|Store Code                    |Name                                                                             |Address                                                                                                    |"
    .ColFlds = "|CstCd|CstName|qCstAdd|"
    .SearchColNum = 1
    .RetFldNm = "CstCd"
    .SortAsc = True
  End With
  
End Sub

'6.1.11 calling help of shipment note number
Public Sub ShpNoteNo(ByVal mCoCd As String, ByVal mTc As String, ByVal mYy As String, ByVal mChr As String)
  Dim wSqlStrg As String
  wSqlStrg = "SELECT SnNo,SnDt,SnCmCd,SnCurCd from ShpNote " + _
             "where SnCoCd= '" + mCoCd + "' And " + _
             "SnTc= '" + mTc + "' and SnYy='" + mYy + "' and " + _
             "SnChr='" + mChr + "' Order By SnNo Desc"
  
  With mwhlp
    .Sqlstrg = wSqlStrg
    .ColHeadg = "|Number|Date          |Customer               |Currency   |"
    .ColFlds = "|SnNo|SnDt|SnCmCd|SnCurCd|"
    .ColMask = "|||" + DtShortStr() + "|||"   '6.1
    .SearchColNum = 1
    .RetFldNm = "SnNo"
    .SortAsc = False
  End With
End Sub


