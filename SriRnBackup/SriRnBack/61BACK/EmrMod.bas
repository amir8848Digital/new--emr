Attribute VB_Name = "EmrMod"
Option Explicit
Public Const ctGenMenuUsr = "MIC"
Public mFrmTrackStatus As New EmrFrmTrackStatus
Public Sub FormSelect(ByVal ps_Mnu As String)
  '*** Jenny Color
  Call SetMnuCd(ps_Mnu)   '*** 13/06/05
  'Dim wfrm As New EmrFrmCustMst
  Dim wfrm As Form
  Set wfrm = Nothing

  '*** (Jen 2.13)
  If Not (GetRecClrSet Is Nothing) Then
  '*** (Jen 2.13)
    If GetRecClrSet.RecCount > 0 Then
      GetRecClrSet.MoveFirst
      '*** For Setting System Colors (For Tab Page)
      SetTabSysColors Val(GetRecClrSet!CsTabPgClr), Val(GetRecClrSet!CsTabForClr), Val(GetRecClrSet!CsTabBorderClr), Val(GetRecClrSet!CsTabShadowClr), Val(GetRecClrSet!CsTabHighLightClr)
      '*** For Setting System Colors (For Tab Page)
    End If
  '*** (Jen 2.13)
  End If
  '*** (Jen 2.13)


  '****** Sachin 2.14.0 - [Restrictions To Microway Users]
  '****** User CRM And SUPP Has No Access To Any Menus ******
  If UCase(gs_UsrCd) = "CRM" Or UCase(gs_UsrCd) = "SUPP" Then DispMsg "The User Has No Rights To View This Menu", etError: Exit Sub
  '****** User MIC Can Create Users Grant UserAccess And Generate Menus Only ******
  'If UCase(gs_UsrCd) = "MIC" And (UCase(ps_Mnu) <> UCase("UsrAcss") And UCase(ps_Mnu) <> UCase("Param") And UCase(ps_Mnu) <> UCase("UsrAdm") And UCase(ps_Mnu) <> UCase("RepLst")) Then DispMsg "The User Has No Rights To View This Menu", etError: Exit Sub
  ' ***** Manali 3.5.0 - 12/11/08 - Maintenance menus added below For "MIC"
  If (UCase(gs_UsrCd) = "MIC" Or UCase(gs_UsrCd) = "JMIC") And (UCase(ps_Mnu) <> UCase("UsrAcss") And UCase(ps_Mnu) <> UCase("Param") And UCase(ps_Mnu) <> UCase("UsrAdm") And UCase(ps_Mnu) <> UCase("RepLst") _
      And UCase(ps_Mnu) <> UCase("EmrUpd") And UCase(ps_Mnu) <> UCase("RepDscRep") And UCase(ps_Mnu) <> UCase("EmrToXML") And UCase(ps_Mnu) <> UCase("TagFmt") And UCase(ps_Mnu) <> UCase("CustMaint")) Then _
      DispMsg "The User Has No Rights To View This Menu" + ps_Mnu, etError: Exit Sub
  '****** User ADM Can Only View the Usr Administration Form ******
  If UCase(gs_UsrCd) = "ADM" And UCase(ps_Mnu) <> UCase("UsrAdm") Then DispMsg "The User Has No Rights To View This Menu", etError: Exit Sub
  '****** Sachin 2.14.0 - [Restrictions To Microway Users]

  '4.1.0.0
  'If Trim(DiskSpaceMsg) <> "" Then MsgBox DiskSpaceMsg, vbCritical + vbOKOnly, "Disk Space Error"

  Select Case UCase(ps_Mnu)
  Case Is = UCase("DM")
    Set wfrm = New EmrFrmDsgMst
  End Select

  If Not (wfrm Is Nothing) Then
    Dim ws_UaMnuForm As String
    ''ws_UaMnuForm = UCase(gocn.GetFldVal("Select UaMnuForm From UsrAccess where UaUsrCd= '" + gs_UsrCd + "' and UaMnuCd= '" + ps_Mnu + "' "))
    
    ws_UaMnuForm = UCase(GetMDIFrm.ADC.Connection.GetFldVal("Select UaMnuForm From UsrAccess where UaUsrCd= '" + gs_UsrCd + "' and UaMnuCd= '" + ps_Mnu + "' "))
    
    If ws_UaMnuForm = "MDI" Then
      ShowForm GetMDIFrm, wfrm, ps_Mnu
    Else
      '''Dim go_FrmLst As New EmrFrmList
   '   If (GetFrmLst Is Nothing) Then
        'Set go_FrmLst = New EmrFrmList
    '    Call SetFrmLst(New EmrFrmList)
     '   ShowForm GetMDIFrm, GetFrmLst, "RepLst"
      'End If
      'ShowForm GetFrmLst, wfrm, ps_Mnu
    End If
    'ShowForm GetMDIFrm, wfrm, ps_Mnu   '*** (Bef 12/08/05)
    'wfrm.WindowState = 0
    'wfrm.Top = 0
    'wfrm.Left = 0
    'wfrm.Height = 10185
    'wfrm.Width = 15300
  End If
  
  SetMDISysColorsAfterLogin   '*** Jenny Color (For Setting System Colors) (12/08/05)
  '*** Jenny Color
End Sub

Public Function DiskSpaceMsg() As String
'3.11.2 - Disk Space Check
    Dim wDSpace As MDORowSet, wRsDrvLet As MDORowSet
   ' If Not GetMDIFrm.ADC.Connection.RecSeek("Select 'x' from diskinfo Where ModDt=Convert(varChar(8), GetDate(),3)") Then GetMDIFrm.ADC.Connection.Execute ("EXECUTE [dbo].[MwSp_UpdDiskInfo] '" + pDbName + "'")
    
    
    '6.1
    Set wDSpace = GetMDIFrm.ADC.Connection.OpenRes(" Set DateFormat " + DtFmtStr() + "; " + _
                                         " Select *,(Case When DiDataDrvFreeSpace<DiDbSize*3 Then 'Data Drive' Else '' End) as data_drive, (Case When DiLogDrvFreeSpace<DiDbSize*3 Then 'Log Drive' Else '' End) as Log_Drive  from diskinfo Where (DiDataDrvFreeSpace<DiDbSize*3 Or DiLogDrvFreeSpace<DiDbSize*3) And ModDt=(Select Max(ModDt) from diskinfo) And DateDiff(day, ModDt, GetDate())<10")
    
    If wDSpace.RecCount > 0 Then
        DiskSpaceMsg = " Low Disk Space on " + IIF(wDSpace!Data_Drive <> "", wDSpace!Data_Drive, "") + IIF(wDSpace!Data_Drive <> "" And wDSpace!Log_Drive <> "", " And ", "") + _
                            IIF(wDSpace!Log_Drive <> "", wDSpace!Log_Drive, "") + " on the Server. " + vbCrLf + _
                            " Free Space Reqd : " + CStr(Round((wDSpace!DiDbSize * 3) / 1024, 2)) + " GB." + vbCrLf + _
                            " Current Free Space : (Data Drive) " + CStr(Round((wDSpace!DiDataDrvFreeSpace) / 1024, 2)) + " GB And (Log Drive) " + CStr(Round((wDSpace!DiLogDrvFreeSpace) / 1024, 2)) + " GB" + vbCrLf + _
                            " Performance May Suffer. Contact Administrator."
    End If
    Set wDSpace = Nothing
    '6.1
    Set wDSpace = GetMDIFrm.ADC.Connection.OpenRes(" Set DateFormat " + DtFmtStr() + "; " + _
                                         " Select 'x' from diskinfo Where ModDt=(Select Max(ModDt) from diskinfo) And DateDiff(day, ModDt, GetDate())>10")
    If wDSpace.RecCount > 0 Or Not GetMDIFrm.ADC.Connection.RecSeek("Select 'x' from diskinfo") Then _
        DiskSpaceMsg = " Disk Status Not Updated. Contact Administrator."
     DiskSpaceMsg = ""
    Set wDSpace = Nothing
End Function



