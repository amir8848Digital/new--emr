Attribute VB_Name = "Mod"
'Manoj
Option Explicit
Public MWLib As MwfLib.MWCLS_GEN
Public HlpList As New EmrClsHlp
Public Const ctlCoCd As Integer = 3
Public Const ctlCmCd As Integer = 8
Public Const ctlCmCtg As Integer = 1
'*** (Jen 04/02/08)
Dim goCn As MDOConnection      '*** (Bef speed) bef 24/11/06
'*** (Jen 04/02/08)

'*** (Bef 04/02/08)
''*** (Jenny speed 24/11/06)
'Public goCn As MDOConnection
''*** (Jenny speed 24/11/06)
'*** (Bef 04/02/08)
Public m_ObjCmdGetRmRt  As MDOQuery
Public m_ObjCmdGetLabRt  As MDOQuery
Public m_ObjCmdGetRmRt_IPD  As MDOQuery
Public m_objCmdGetRmRt_Base As MDOQuery
Public gs_DateFormatAtLogin As String

''*** (Jenny speed 04/02/08)
Public go_HlpCn As MDOConnection
''*** (Jenny speed 04/02/08)

'****** Sachin 3.02 - Report Server Connection - [22/09/2007]
'*** (Bef 04/02/08)
'Public groCn As MDOConnection
'*** (Bef 04/02/08)
Public gbSvrChk As Boolean
Public gs_PurgeDt As String, gs_Partition As String
Public Const ctCurrPrtn = "C"
'****** Sachin 3.02 - Report Server Connection - [22/09/2007]
Public gs_InCoCd As String, gs_InTc As String, gs_InYy As String, gs_InChr As String, gs_InNo As Long

Public gs_OrdZoomSel As String    '****** Sachin 3.3.0 - (26-06-08) - [21. Order Copy From Zoom] ******

'Public gs_ModulesStr(22) As String  ' ***** Sachin 3.3.0 - (26-06-08) - [16. Module and Ini Check] ****** '*** Jay 3.4(FG)
'Public gs_ModulesStr(23) As String  ' ***** Manali 3.4.1 - 22/10/08 - Design Catalogue Module
'Public gs_ModulesStr(24) As String  ' ***** Manali 3.6.0 - Order Printing In Excel Format
'Public gs_ModulesStr(25) As String  ' ***** Manali 3.9.1 - 21/02/12 - RfId Module
'Public gs_ModulesStr(26) As String  ' ***** Manali 3.10.0 - 04/04/12 - Multi Price Quotation Module
Public gs_ModulesStr(29) As String   ' ***** 3.11.0 - Refresh by Rate Chart
Public gs_UlIdNo As Long        '****** Sachin 3.5.0 [11/11/08] ******
Private Enum TErrorCorretion
    QualityLow
    QualityMedium
    QualityStandard
    QualityHigh
End Enum
 
'**** FOR API CALLS
'??? reqd api calls
Private Declare Sub GenerateBMP _
                Lib "C:\Temp\quricol32.dll" _
                Alias "GenerateBMPW" ( _
                ByVal FileName As Long, _
                ByVal text As Long, _
                ByVal Margin As Long, _
                ByVal Size As Long, _
                ByVal Level As TErrorCorretion)

Declare Function GetSystemMenu Lib "user32" (ByVal hwnd As Long, ByVal bRevert As Long) As Long
Declare Function RemoveMenu Lib "user32" (ByVal hMenu As Long, ByVal nPosition As Long, ByVal wFlags As Long) As Long
Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Long) As Long
Public Const MF_BYPOSITION = &H400
Declare Function Beep Lib "kernel32" (ByVal dwFreq As Long, ByVal dwDuration As Long) As Long
Declare Function FlashWindow Lib "user32" (ByVal hwnd As Long, ByVal bInvert As Long) As Long
'*************

' ########################  Manoj #### Ver: 2.0.6 #### Date: 16/06/2004  ########################
' API Used For Accessing Host Computer Name
Public Declare Function GetComputerName Lib "kernel32" Alias "GetComputerNameA" (ByVal lpBuffer As String, nSize As Long) As Long
' ###########################################  Manoj  ###########################################

Enum En_opt
    Initialise
    When
    Valid
End Enum

'*************
' ########################  Manoj #### Ver: 2.0.6 #### Date: 21/02/2004  ########################
  Public gColTgFldTyp As New Collection
  'variable to save Host Machine Name
  Public gs_HostNm As String
' ###########################################  Manoj  ###########################################

Public mwhlp As Object

'*** For multi company
Public gs_CoCd As String
Public gs_CoNm As String
Public gs_UsrCd As String
Public Const ctSelfCoCd = "ZZZ"
Public gs_Ver As String
'*** For multi company

'****** Sachin 3.01 - MaxDop In All Reports ******
Public Const ctMaxDopOpt = " Option (MaxDop 1)"
'****** Sachin 3.01 - MaxDop In All Reports ******

Public Const ctProjectName = "EMR"
''''Public Const ctGenMenuUsr = "MIC"
Public Const ctSelfCmCd = "ZSELF"
Public Const ctDsgCpySelf = "ZSELF"
Public Const ctDsgCpyCust = "CUST"
Public Const ctDsgCpyMst = "DSGMST"
Public Const ctDsgCpyBag = "BAG"
Public Const ctDsgCpyDmId = "DMID"     '*** Jay 3.3(FG)

' **** Zubin 211 (include below when TcTyp= 'QF' is included) **** '
'''''Public Const ctDsgCpyFG = "FG"
' **** Zubin 211 **** '
Public Const ctTdWrk = "ZSELF"
Public Const ctFgLoc = "PFG"
Public Const ctOmTcTyp = "OM"
Public Const ctInTcTyp = "IN"
Public Const ctCTBTcTyp = "CTB"
Public Const ctValueCmCd = "ZVALUE"
Public Const ctDefCurncy = "US$"
Public Const ctIndCurncy = "RS"
Public Const ctFotTcTyp = "FG"     '*** Jay 3.3(FG)
Public Const ctJtTcTyp = "JT" ' ***** Manali Trading Module

'*** urmi AvgStkrt *****
Public Const ctAvgLocTypSeq = "('W', 'S', 'L', 'M', 'B')"
Public Const ctLotNoRmCtg = "('D', 'C', 'X','G','P','S','L','M')"
'**********************
' ########################  Manoj #### Ver: 2.0.8 #### Date: 07/10/2004  ########################
  Public Const MWSrvrClsID = "{1E774FFA-40BF-4B1D-B4A0-B95BC24B3766}"
' ###########################################  Manoj  ###########################################

'*** (Jen 2.12)
Public Const ctBrkLMCd = "ZBRK"
Public Const ctBrkLSCd = "-"
'*** (Jen 2.12)


'*** Manali 3.5.0

Public Const ctDHCLMCd = "DHC"
Public Const ctCHCLMCd = "CHC"
Public Const ctDCHCLSCd = "-"
Public Const ctCDWLMCd = "CDW"
Public Const ctGHCLMCd = "GHC"
Public Const ctCBRKLMCd = "CBRK"

'*** Manali 3.5.0

Public gs_LocTyp As String
Public gb_RmCtgFor As Boolean
Public gb_RmZ As Boolean
Public gb_RmZ1 As Boolean     '*** (Jen 3.2.0)
Public gs_DmTcTyp As String
Public gs_TxnTcTyp As String
Public gs_OmTcTyp As String
Public gs_CmCtg As String
Public gs_CmCtg1 As String     ' ***** Manali 3.8.0 - For Help Of UoOmTc
Public gs_RmGrp As String
Public gs_LabGrp As String
Public gs_ParamOpt As String
Public gs_PTyp As String
Public gs_PTyp1 As String
Public gs_LabMCd As String
Public gs_FgTcTyp As String
Public gs_ValidPrdLocSeq As String
Public gs_ValidOthLocSeq As String
Public gs_ValidFgSubLoc As String
Public gb_CmCtgFor As Boolean
Public gb_ShowYyMmHlp As Boolean
Public gs_Tbl As String
Public gs_InTcTyp As String
Public gs_DmCtg As String 'Geeta****Crm***212
Public gb_CrmDb As Boolean 'Geeta****Crm***212
Public gb_CoCdFor As Boolean  '****** Sachin 2.14 - [Multi-Company] ******

Public Const ctAllLineCd = "ZZZ"

Public gs_DsgLocTyp As String
Public gs_TcTyp As String
Public gs_VchTyp As String
Public gs_EntTyp As String    ' **** Manali Trading Module
Public gb_InHlpFrOm As Boolean   ' ***** Manali 3.7.1
Public gs_ActFrmMenucd As String  '6.1 hlp
'urmi 2.04-3 SoPnd Slow, Option to load Picture
Public gb_ShowPic As Boolean

' ########################  Manoj #### Ver: 2.0.9 #### Date: 22/11/2004  ########################
Public gb_SingleCoCd As Boolean
' ###########################################  Manoj  ###########################################

' ########################  Manoj #### Ver: 2.11.0 #### Date: 26/05/2005  ########################
Public gs_GrphGrpNo As Integer
Public gb_MailAct As Boolean
' ###########################################  Manoj  ###########################################
' Zubin 213
Public Const ctLmeLimit = 99999.999
Public Const ctLmeSalRtLimit = 999999.999
Public Const ctLabRtLimit = 99999.99
' Zubin 213

' ***** Manali 3.8.0 - ImpCtg
Public Const ctImpCtgImp = "IMP"
Public Const ctImpCtgInd = "IND"
' ***** Manali 3.8.0 - ImpCtg

'*** (Jen 2.13) ***
Public gs_ErrMsg As String
'*** (Jen 2.13) ***

'*** Jay 3.01 **** [LotNo Err]
Public gs_Loc As String
Public gs_RmLocTyp As String

' Zubin (3.02) - 250308
Public gs_PMCd As String
' Zubin (3.02) - 250308

'****** Sachin 3.3.0 (25-06-08) - [12. Import From Excel] ******
Dim ms_objExcel, ms_xlWBook

' ***** 4.1.0.0 - Excel Pwd
Public Const ctXCelPwd = "1234"
Public Const ctXCelWritePwd = "1234"
' ***** 4.1.0.0 - Excel Pwd
Public Const ctBTyp = "D" 'Uma EmrSv2
Public g_InProcessingMode As Boolean
Public g_DateSetting As String
Declare Function LockWindowUpdate Lib "user32" (ByVal hwndLock As Long) As Long
Public Sub DispMsg(ByVal pv_MsgStr As String, ByVal pv_ErrorType As en_ErrorType)
    Call GetMDIFrm.Ctl_HandleError(pv_MsgStr, pv_ErrorType)
End Sub
'Public Function GetFldVal_FromHead(ByVal pv_FldName As String, ByVal pv_CoCd As String) As String
'    GetFldVal_FromHead = goCn.GetFldVal_FromDynamic(pv_FldName, "Select Hname from Head where HCoCd=?", "@HCoCd", enMDOQueryDataType.mdovarchar, ctlCoCd, pv_CoCd)
'End Function
'Public Function GetFldVal_FromCustMst(ByVal pv_FldName As String, ByVal pv_CmCtg As String, ByVal pv_CmCd As String) As String
'    GetFldVal_FromCustMst = goCn.GetFldVal_FromDynamic(pv_FldName, "Select CmName from Custmst where CmCtg=? and CmCd=?", "@CmCtg", enMDOQueryDataType.mdovarchar, ctlCmCtg, pv_CmCtg, _
                                                                                                                            "@CmCd", enMDOQueryDataType.mdovarchar, ctlCmCd, pv_CmCd)
'End Function
Public Sub FrmActivate(CurrForm As Object)
    Call GetMDIFrm.Ctl.FrmActivate(CurrForm)
    Set mwhlp = Screen.ActiveForm.MWCTL_HLP1
    Set goCn = CurrForm.ADC.Connection
    gs_ActFrmMenucd = CurrForm.ADC.MenuCd   '6.1 hlp
    CurrForm.ADC.RepExportFileDirectory = GetExportFilePath()

End Sub
Public Sub FrmDeActivate(CurrForm As Object)
    '*** (Jen 2.13) This code works. Check it properly.
    '****** This piece of code is introduced because if the user is in the scope of a particular report and he types
    '       some value in a scope field and switches to another form and comes back to the original report scope,
    '       the value typed was not getting saved in the field but it was being displayed. Now when the report is taken,
    '       the user will not get the report according to the scope he has entered which is confusing.
    '       With the below specified code, the value in the field will be accepted if valid or will be initialised to
    '       blank if the user switches to another entry form.
    '******
    If Not (CurrForm.ADC Is Nothing) Then
      '****** currently this piece of code is applicable only for report forms
      If Not (CurrForm.ActiveControl Is Nothing) And CurrForm.ADC.EntType = xetReport Then
        '****** This check is done onl for MED user controls
        If TypeOf CurrForm.ActiveControl Is MwfCtl.MWCTL_MED Then
          Dim wVar_Val As Variant
          wVar_Val = CurrForm.ActiveControl.Value
          
          ''CurrForm.ActiveControl.Validate
          
          '************ This can be used if the original value of the field is to be retained. ******************
          '****** If the value in the active field is invalid then it is reverted back to the original value before
          '       the user changed it
          If CurrForm.ActiveControl.Validate = True Then
            CurrForm.ActiveControl.InitialiseFld wVar_Val
          End If
          '************ This can be used if the original value of the field is to be retained. ******************
          
          '************ This can be used if the original value of the field is not to be retained. (This works) ********
          ''****** If the value in the active field is invalid then initialise it to the appropriate blank value
          ''****** InitialiseFld is used so that the validation of the field is not triggered again
          'If CurrForm.ActiveControl.Validate = True Then
          '  Select Case CurrForm.ActiveControl.DataType
          '  Case Is = fdtFloat, fdtNumber, fdtTimeType
          '    CurrForm.ActiveControl.InitialiseFld CurrForm.ActiveControl.formatText(0)
          '  Case Is = fdtCharType, fdtConvertToUpper
          '    CurrForm.ActiveControl.InitialiseFld ""
          '  Case Is = fdtDateType
          '    CurrForm.ActiveControl.InitialiseFld "01/01/80"
          '  End Select
          '  'CurrForm.ActiveControl.Value(False) = ""   '*** This might work (not tested)
          'End If
          '************ This can be used if the original value of the field is not to be retained. (This works) ********
        End If
      End If
    End If
    
    ''If Not (CurrForm.ActiveControl Is Nothing) Then
    ''  If TypeOf CurrForm.ActiveControl Is MwfCtl.MWCTL_MED Then
    ''    If CurrForm.ActiveControl.Validate = True Then
    ''      'DispMsg "aaa", etError
    ''      Exit Sub
    ''    'Else
    ''    '  DispMsg "bbb", etError
    ''    End If
    ''  End If
    ''End If
    '*** (Jen 2.13) This code works. Check it properly.
    
    Call GetMDIFrm.Ctl.FrmDeActivate(CurrForm)
    'Call EmrFrmMDI.Ctl.FrmDeActivate(CurrForm)
    '*** (Jen 04/02/08)
    Set goCn = Nothing     '*** (Bef speed) bef 24/11/06
    '*** (Jen 04/02/08)
End Sub

Public Sub ShowForm(ByVal ContainerForm As Form, ByVal Frm As Form, MenuCd As String)
    Dim wf_ExecGot  As Boolean
    On Error GoTo ErrH
' ########################  Manoj #### Ver: 2.0.6 #### Date: 16/06/2004  ########################
' Check Login Session for Working or Not
    'If Not ChkLogin(ADC.Connection) Then End  '*** (Bef 12/08/05)
    If Not ChkLogin(GetMDIFrm.ADC.Connection) Then End  '*** (12/08/05)
' ###########################################  Manoj  ###########################################
    'MWLib.BeginProcess Me, "Loading " + Controls(MenuCd).Caption   '*** (Bef 12/08/05)
    MWLib.BeginProcess ContainerForm, "Loading " + ContainerForm.Controls(MenuCd).Caption '*** (12/08/05)
    wf_ExecGot = False
    If Not (GetMDIFrm.ActiveForm Is Nothing) Then
        GetMDIFrm.ActiveForm.ADC.ExecGot = True
        wf_ExecGot = True
    End If
    Load Frm
    'takecloseoff Frm.hWnd    '*** (Bef 12/08/05)
    GetMDIFrm.takecloseoff Frm.hwnd '*** (12/08/05)
'    'Frm.Icon = GetMDIFrm.Icon
    If wf_ExecGot Then Frm.ADC.ExecGot = True
    If Frm.WindowState = 0 Then
      Frm.Left = 0
      Frm.Top = 0
    End If
    Frm.ADC.MenuCd = MenuCd
    Frm.ADC.UsrCd = gs_UsrCd
' ########################  Manoj #### Ver: 2.0.6 #### Date: 03/06/2004  ########################
  ' Set HostName  as ADC.HostNm for User Check
    Frm.ADC.HostNm = gs_HostNm
    Frm.ADC.UlIdNo = gs_UlIdNo      '****** Sachin 3.5.0 [11/11/08]
    Frm.ADC.CompanyCd = gs_CoCd     '*** Jay 3.1.1
  ' set Lising Form AS MenuForm for Loading Form
    'Set Frm.ADC.MnuForm = Me   '*** (Bef 12/08/05)
    Set Frm.ADC.MnuForm = ContainerForm   '*** (12/08/05)
    
    '*** Jenny Color (12/08/05)
    Frm.ADC.ToolBarCtrlName = "CmdCtlToolBar"
    'Frm.ADC.ToolBarCtrlIndx = GetMDIFrm.mo_TlBarColl.Item(MenuCd)
    Frm.ADC.ToolBarCtrlIndx = GetMDIFrm.TlBarColl.Item(MenuCd)
    '*** Jenny Color (12/08/05)
    
    If Not Frm.ADC.InitRtn Then End
    'ADC.Visible = False   '*** (Bef 12/08/05)  (Was used in EmrFrmList)
    
' ###########################################  Manoj  ###########################################
    Set Frm = Nothing
    'MWLib.EndProcess Me    '*** (Bef 12/08/05)
    MWLib.EndProcess ContainerForm   '*** (12/08/05)
    Exit Sub
ErrH:
    If Err.Number = 40002 Then
        Resume Next
    '*** Jenny Color (12/08/05)
    ElseIf Err.Number = 9 Or Err.Number = 5 Then
        '*** 9: If the reqd Menucode does not exist in the collection then this error is generated
        Frm.ADC.ToolBarCtrlName = ""
        Resume Next
    '*** Jenny Color (12/08/05)
    Else
        MsgBox Err.Description
    End If
End Sub
Public Function GetExportFilePath() As String
    GetExportFilePath = goCn.GetFldVal("Select PDesc225 from Param Where PTyp='PCT' and " + _
           "PMCd='3D'") + "\ExportReports\"
End Function
Public Sub InitProp(ByVal ActiveFrm As Form)
  
  '*** Parameters Used For Setting Properties Of UsrOpt Fields ***
  gs_LocTyp = "'P'"
  gb_RmCtgFor = True
  gb_RmZ = False
  gb_RmZ1 = False   '*** (Jen 3.2.0)
  gs_DmTcTyp = "DM"
  gs_TxnTcTyp = "BV"
  gs_OmTcTyp = "SO"
  gs_CmCtg = "C"
  gs_CmCtg1 = "C"     ' ***** Manali 3.8.0 - For Help Of UoOmTc
  gs_RmGrp = "ORDDET"
  gs_LabGrp = "ORDDET"
  gs_ParamOpt = ""
  gs_PTyp = ""
  gs_PTyp1 = ""
  gs_LabMCd = ""
  gs_FgTcTyp = "FB"
  gb_CmCtgFor = True
  gb_ShowYyMmHlp = False
  gs_Tbl = "Param"
  
  gs_DsgLocTyp = ""
  gs_TcTyp = ""
  gs_VchTyp = ""
  gs_InTcTyp = "IN"
  gs_EntTyp = ""  ' ***** Manali Trading Module
  
  gs_DmCtg = ""
  
  ' Zubin (3.02) - 250308
  gs_PMCd = ""
  ' Zubin (3.02) - 250308
  
  'Defa val by urmi to shoe Pic
  gb_ShowPic = True
' ########################  Manoj #### Ver: 2.11.0 #### Date: 26/05/2005  ########################
  gs_GrphGrpNo = 0
' ###########################################  Manoj  ###########################################
  
  gb_CoCdFor = True '****** Sachin 2.14 - [Multi-Company] ******
  '*** Parameters Used For Setting Properties Of UsrOpt Fields ***
  
  gb_InHlpFrOm = False ' **** Manali 3.7.1

  Dim wo_Ctl As Control
  For Each wo_Ctl In ActiveFrm.Controls
'??? consider speed
    If TypeName(wo_Ctl) = "MWCTL_MED" Then
        SetProp ActiveFrm, wo_Ctl.IdName, Initialise
    End If
  Next
End Sub

' ************************************ Zubin 211 ************************************ '
' ************************************ Before 211 ************************************ '
' ************************************ Zubin 211 ************************************ '
'''''Public Function GetPictPath(ByVal mTcType As String, _
'''''                            ByVal mLineDraw As Boolean, _
'''''                            ByVal mDesignCd As String, _
'''''                            Optional ByVal mDesignSz As String)
'''''  '*** Naming Conventions for 3D Images
'''''      '*** Dsg with TcTyp 'DM' or 'PM' and Sz<> '':    DM 3D DsgCd Sz.jpg
'''''      '*** Dsg with TcTyp 'DM' or 'PM' and Sz= '' :    DM 3D DsgCd.jpg
'''''  '*** Naming Conventions for Line Drawings
'''''      '*** Dsg with TcTyp 'DM' or 'PM' and Sz<> '':    DM LD DsgCd Sz.jpg
'''''      '*** Dsg with TcTyp 'DM' or 'PM' and Sz= '' :    DM LD DsgCd.jpg
'''''      '*** Dsg with TcTyp 'SM' and Sz<> ''        :    SM LD DsgCd Sz.jpg
'''''      '*** Dsg with TcTyp 'SM' and Sz= ''         :    SM LD DsgCd.jpg
'''''
'''''  '*** 3D and LD Pics can be in different folders
'''''  '*** If Picture with Size is not found then the function searches for Picture without Size
'''''
'''''  Dim wFlNm  As String, wDsgCtg As String, wPth As String, wTyp As String
'''''  Dim wSz As String, wDir  As String
'''''  On Error GoTo errh
'''''  wFlNm = MWLib.RmvEmbSpace(mDesignCd)
'''''  ' Manoj 2.10.0
'''''  If mTcType = "DM" Then
'''''  'If mTcType = "DM" Or mTcType = "PM" Then
'''''  ' Manoj 2.10.0
'''''    wDsgCtg = gocn.GetFldVal("Select DmCtg from DsgMst where DmTcTyp='DM' and DmCd='" + mDesignCd + "' and DmSz= ''")
'''''    '*** Get the picture path from param depending on whether '3D' pic or 'Line Drawing' is to be shown
'''''    wPth = gocn.GetFldVal("Select PDesc225 from Param Where PTyp='PCT' and " + _
'''''           "PMCd='" + IIF(mLineDraw = False, "3D", "HSK") + "'") + "\" + wDsgCtg
'''''  ElseIf mTcType = "SM" Then
'''''    wDsgCtg = gocn.GetFldVal("Select DmCtg from DsgMst where DmTcTyp='SM' and DmCd='" + mDesignCd + "' and DmSz= ''")
'''''    '*** Get the 'Line Drawing' picture path from param for Sketch Master
'''''    wPth = gocn.GetFldVal("Select PDesc225 from Param Where PTyp='PCT' and PMCd='HSK'") + "\" + wDsgCtg
'''''  ' Manoj 2.10.0
'''''  ElseIf mTcType = "PM" Then
'''''    wDsgCtg = gocn.GetFldVal("Select DmCtg from DsgMst where DmTcTyp='PM' and DmCd='" + mDesignCd + "' and DmSz= ''")
'''''    wPth = gocn.GetFldVal("Select PDesc225 from Param Where PTyp='PCT' and " + _
'''''           "PMCd='" + IIF(mLineDraw = False, "3D", "HSK") + "'") + "\" + wDsgCtg
'''''  ' Manoj 2.10.0
'''''  End If
'''''  wSz = ""
'''''
'''''  If Not IsMissing(mDesignSz) Then
'''''    If mDesignSz <> "" Then wSz = " " + mDesignSz
'''''  End If
'''''
'''''' Manoj 2.10.0
''''''  wTyp = IIF(mTcType = "DM" Or mTcType = "PM", "DM", mTcType)
'''''  wTyp = mTcType
''''''Manoj 2.10.0
'''''  If mLineDraw = True Then
'''''    wDir = Dir(wPth + "\" + wTyp + " LD " + wFlNm + wSz + ".jpg")
'''''    If wDir = "" Then wSz = ""
'''''    GetPictPath = wPth + "\" + wTyp + " LD " + wFlNm + wSz + ".jpg"
'''''  Else
'''''    wDir = Dir(wPth + "\" + wTyp + " 3D " + wFlNm + wSz + ".jpg")
'''''    If wDir = "" Then wSz = ""
'''''    GetPictPath = wPth + "\" + wTyp + " 3D " + wFlNm + wSz + ".jpg"
'''''  End If
'''''  Exit Function
'''''errh:
'''''  GetPictPath = ""
'''''  Exit Function
'''''End Function
' ************************************ Zubin 211 ************************************ '

'*** Jay 3.2.0 [PicNm] (OrdDets Parameter added)
Public Function GetPictPath(ByVal mTcType As String, _
                            ByVal mLineDraw As Boolean, _
                            ByVal mDesignCd As String, _
                            Optional ByVal mDesignSz As String, _
                            Optional mOdCoCd As String, Optional mOdTc As String, Optional mOdYy As String, _
                            Optional mOdChr As String, Optional mOdNo As Double, Optional mOdSr As Double, Optional mOdPicNm As String, Optional mPicNo As String, Optional mHDImg As Boolean)
  '*** Naming Conventions for 3D Images
      '*** Dsg with TcTyp 'DM' or 'PM' and Sz<> '':    DM 3D DsgCd Sz.jpg
      '*** Dsg with TcTyp 'DM' or 'PM' and Sz= '' :    DM 3D DsgCd.jpg
  '*** Naming Conventions for Line Drawings
      '*** Dsg with TcTyp 'DM' or 'PM' and Sz<> '':    DM LD DsgCd Sz.jpg
      '*** Dsg with TcTyp 'DM' or 'PM' and Sz= '' :    DM LD DsgCd.jpg
      '*** Dsg with TcTyp 'SM' and Sz<> ''        :    SM LD DsgCd Sz.jpg
      '*** Dsg with TcTyp 'SM' and Sz= ''         :    SM LD DsgCd.jpg
  
  '*** 3D and LD Pics can be in different folders
  '*** If Picture with Size is not found then the function searches for Picture without Size
  
  Dim wFlNm  As String, wDsgCtg As String, wPth As String, wTyp As String
  Dim wSz As String, wDir  As String
  Dim wPicNm As String '*** Jay 3.2.0 [PicNm]
  Dim wPicNo As String
  Dim wFlg As Boolean
  On Error GoTo ErrH
  wPicNo = IIF(mPicNo <> "", "_" + mPicNo, "") 'Sachin 3.11.0 - 3 additional images
  wFlg = False
PicH:
  If wFlg Then
    wFlNm = MWLib.RmvEmbSpace(mDesignCd)
  Else
    wFlNm = mDesignCd
  End If
  
  ' Manoj 2.10.0
  If mTcType = "DM" Then
  'If mTcType = "DM" Or mTcType = "PM" Then
  ' Manoj 2.10.0
    wDsgCtg = goCn.GetFldVal("Select DmCtg from DsgMst where DmTcTyp='DM' and DmCd='" + mDesignCd + "' and DmSz= ''" + IIF(gs_Partition = ctCurrPrtn, " and DmPrtKey='" + ctCurrPrtn + "'", ""))
    '*** Get the picture path from param depending on whether '3D' pic or 'Line Drawing' is to be shown
    wPth = goCn.GetFldVal("Select PDesc225 from Param Where PTyp='PCT' and " + _
           "PMCd='" + IIF(mLineDraw = False, "3D", "HSK") + "'") + "\" + wDsgCtg
  ElseIf mTcType = "SM" Then
    wDsgCtg = goCn.GetFldVal("Select DmCtg from DsgMst where DmTcTyp='SM' and DmCd='" + mDesignCd + "' and DmSz= ''" + IIF(gs_Partition = ctCurrPrtn, " and DmPrtKey='" + ctCurrPrtn + "'", ""))
    '*** Get the 'Line Drawing' picture path from param for Sketch Master
    wPth = goCn.GetFldVal("Select PDesc225 from Param Where PTyp='PCT' and PMCd='HSK'") + "\" + wDsgCtg
  ' Manoj 2.10.0
  ElseIf mTcType = "PM" Then
    wDsgCtg = goCn.GetFldVal("Select DmCtg from DsgMst where DmTcTyp='PM' and DmCd='" + mDesignCd + "' and DmSz= ''" + IIF(gs_Partition = ctCurrPrtn, " and DmPrtKey='" + ctCurrPrtn + "'", ""))
    wPth = goCn.GetFldVal("Select PDesc225 from Param Where PTyp='PCT' and " + _
           "PMCd='" + IIF(mLineDraw = False, "3D", "HSK") + "'") + "\" + wDsgCtg
  ' Manoj 2.10.0
  End If
  wSz = ""
  
  If Not IsMissing(mDesignSz) Then
    If mDesignSz <> "" Then wSz = " " + mDesignSz
  End If
  
' Manoj 2.10.0
'  wTyp = IIF(mTcType = "DM" Or mTcType = "PM", "DM", mTcType)
  wTyp = mTcType
'Manoj 2.10.0
  
  '*** Jay 3.2.0 [PicNm]
  If mOdPicNm <> "" Then
    wPicNm = mOdPicNm
  ElseIf Not (IsMissing(mOdCoCd) Or IsMissing(mOdTc) Or IsMissing(mOdYy) Or IsMissing(mOdChr) Or IsMissing(mOdNo) Or IsMissing(mOdSr)) Then
    wPicNm = goCn.GetFldVal("Select OdPicNm from OrdDsg where OdCoCd= '" + mOdCoCd + "' and OdTc= '" + mOdTc + "' and " + _
            "OdYy= '" + mOdYy + "' and OdChr= '" + mOdChr + "' and OdNo= " + CStr(mOdNo) + " and OdSr= " + CStr(mOdSr) + IIF(gs_Partition = ctCurrPrtn, " and OdPrtKey='" + ctCurrPrtn + "'", ""))
  End If
  
  'Sachin - 4.1.0.0
  Dim wPicTyp As String
  If IsMissing(mHDImg) Then
     wPicTyp = " 3D "
  Else
     wPicTyp = IIF(mHDImg, " HD ", " 3D ")
  End If
  
  '*** Jay 3.2.0 [PicNm]
  If mLineDraw = True Then
    '*** Jay 3.2.0 [PicNm]
    If wPicNm <> "" Then wDir = Dir(wPth + "\" + wTyp + " LD " + wPicNm + ".jpg")
    If wDir = "" Then
    '*** Jay 3.2.0 [PicNm]
      wDir = Dir(wPth + "\" + wTyp + " LD " + wFlNm + wSz + ".jpg")
      If wDir = "" Then wSz = ""
      wDir = Dir(wPth + "\" + wTyp + " LD " + wFlNm + wSz + ".jpg")
      If wDir = "" And Not wFlg Then wFlg = True: GoTo PicH:
      
      GetPictPath = wPth + "\" + wTyp + " LD " + wFlNm + wSz + ".jpg"
    '*** Jay 3.2.0 [PicNm]
    Else
      GetPictPath = wPth + "\" + wTyp + " LD " + wPicNm + ".jpg"
    End If
    '*** Jay 3.2.0 [PicNm]
  Else
    '*** Jay 3.2.0 [PicNm]
    If wPicNm <> "" Then wDir = Dir(wPth + "\" + wTyp + wPicTyp + wPicNm + ".jpg")
    
    If wDir = "" Then
    '*** Jay 3.2.0 [PicNm]
      wDir = Dir(wPth + "\" + wTyp + wPicTyp + wFlNm + wSz + wPicNo + ".jpg")
      If wDir = "" Then wSz = ""
      
      wDir = Dir(wPth + "\" + wTyp + wPicTyp + wFlNm + wSz + wPicNo + ".jpg")
      If wDir = "" And Not wFlg Then wFlg = True: GoTo PicH:
    
      GetPictPath = wPth + "\" + wTyp + wPicTyp + wFlNm + wSz + wPicNo + ".jpg"
    '*** Jay 3.2.0 [PicNm]
    Else
      GetPictPath = wPth + "\" + wTyp + wPicTyp + wPicNm + wPicNo + ".jpg"
    End If
    '*** Jay 3.2.0 [PicNm]
  End If
  Exit Function
ErrH:
  
  GetPictPath = ""
  Exit Function
End Function

'*** (Bef 20/06/05)
'Public Function GetPictPath(ByVal mTcType As String, ByVal mDesignCd As String, Optional ByVal mDesignSz As String, _
'                            Optional ByVal mLineDraw As Variant)
'  '*** Naming Conventions for 3D Images
'      '*** Dsg with TcTyp 'DM' or 'PM' and Sz<> '':    DM 3D DsgCd Sz.jpg
'      '*** Dsg with TcTyp 'DM' or 'PM' and Sz= '' :    DM 3D DsgCd.jpg
'  '*** Naming Conventions for Line Drawings
'      '*** Dsg with TcTyp 'DM' or 'PM' and Sz<> '':    DM LD DsgCd Sz.jpg
'      '*** Dsg with TcTyp 'DM' or 'PM' and Sz= '' :    DM LD DsgCd.jpg
'      '*** Dsg with TcTyp 'SM' and Sz<> ''        :    SM LD DsgCd Sz.jpg
'      '*** Dsg with TcTyp 'SM' and Sz= ''         :    SM LD DsgCd.jpg
'
'  '*** 3D and LD Pics can be in different folders
'  '*** If Picture with Size is not found then the function searches for Picture without Size
'
'  '*** If mLineDraw parameter is passed, then '3D' or 'HSK' pic paths are returned depending upon their values.
'  '*** If mLineDraw parameter is not passed, then '3D' pic path is returned if found else 'HSK' pic path is returned.
'
'  Dim wFlNm  As String, wDsgCtg As String, wPth As String, wTyp As String
'  Dim wSz As String, wDir  As String
'  On Error GoTo errh
'  wFlNm = MWLib.RmvEmbSpace(mDesignCd)
'  ' Manoj 2.10.0
'  If mTcType = "DM" Then
'  'If mTcType = "DM" Or mTcType = "PM" Then
'  ' Manoj 2.10.0
'    wDsgCtg = gocn.GetFldVal("Select DmCtg from DsgMst where DmTcTyp='DM' and DmCd='" + mDesignCd + "' and DmSz= ''")
'    '*** Get the picture path from param depending on whether '3D' pic or 'Line Drawing' is to be shown
'
'    ' **** Zubin 211 **** '
''''''    wPth = gocn.GetFldVal("Select PDesc225 from Param Where PTyp='PCT' and " + _
''''''           "PMCd='" + IIF(mLineDraw = False, "3D", "HSK") + "'") + "\" + wDsgCtg
'    If Not IsMissing(mLineDraw) Then
'      If mLineDraw = False Then
'        wPth = gocn.GetFldVal("Select PDesc225 from Param Where PTyp='PCT' and PMCd='3D'") + "\" + wDsgCtg
'      Else
'        wPth = gocn.GetFldVal("Select PDesc225 from Param Where PTyp='PCT' and PMCd='HSK'") + "\" + wDsgCtg
'      End If
'    End If
'    ' **** Zubin 211 **** '
'
'  ElseIf mTcType = "SM" Then
'    wDsgCtg = gocn.GetFldVal("Select DmCtg from DsgMst where DmTcTyp='SM' and DmCd='" + mDesignCd + "' and DmSz= ''")
'    '*** Get the 'Line Drawing' picture path from param for Sketch Master
'    wPth = gocn.GetFldVal("Select PDesc225 from Param Where PTyp='PCT' and PMCd='HSK'") + "\" + wDsgCtg
'  ' Manoj 2.10.0
'  ElseIf mTcType = "PM" Then
'    wDsgCtg = gocn.GetFldVal("Select DmCtg from DsgMst where DmTcTyp='PM' and DmCd='" + mDesignCd + "' and DmSz= ''")
'    ' **** Zubin 211 **** '
''''''    wPth = gocn.GetFldVal("Select PDesc225 from Param Where PTyp='PCT' and " + _
''''''           "PMCd='" + IIF(mLineDraw = False, "3D", "HSK") + "'") + "\" + wDsgCtg
'    If Not IsMissing(mLineDraw) Then
'      If mLineDraw = False Then
'        wPth = gocn.GetFldVal("Select PDesc225 from Param Where PTyp='PCT' and PMCd='3D'") + "\" + wDsgCtg
'      Else
'        wPth = gocn.GetFldVal("Select PDesc225 from Param Where PTyp='PCT' and PMCd='HSK'") + "\" + wDsgCtg
'      End If
'    End If
'    ' **** Zubin 211 **** '
'  ' Manoj 2.10.0
'  End If
'  wSz = ""
'
'  If Not IsMissing(mDesignSz) Then
'    If mDesignSz <> "" Then wSz = " " + mDesignSz
'  End If
'
'' Manoj 2.10.0
''  wTyp = IIF(mTcType = "DM" Or mTcType = "PM", "DM", mTcType)
'  wTyp = mTcType
''Manoj 2.10.0
'
'  ' **** Zubin 211 **** '
'  If IsMissing(mLineDraw) Then
'    wPth = gocn.GetFldVal("Select PDesc225 from Param Where PTyp='PCT' and PMCd='3D'") + "\" + wDsgCtg
'    wDir = Dir(wPth + "\" + wTyp + " 3D " + wFlNm + wSz + ".jpg")
'    'If wDir = "" Then wSz = ""
'
'    If wDir = "" Then
'      If Not IsMissing(mDesignSz) Then
'        wSz = ""
'        wDir = Dir(wPth + "\" + wTyp + " 3D " + wFlNm + wSz + ".jpg")
'        If wDir = "" Then
'          wPth = gocn.GetFldVal("Select PDesc225 from Param Where PTyp='PCT' and PMCd='HSK'") + "\" + wDsgCtg
'          If Not IsMissing(mDesignSz) Then If mDesignSz <> "" Then wSz = " " + mDesignSz
'          'If wDir = "" Then wSz = ""
'          wDir = Dir(wPth + "\" + wTyp + " LD " + wFlNm + wSz + ".jpg")
'          If wDir = "" Then wSz = ""
'
'          If wDir <> "" Then
'            GetPictPath = wPth + "\" + wTyp + " LD " + wFlNm + wSz + ".jpg"
'            Exit Function
'          End If
'        Else
'          GetPictPath = wPth + "\" + wTyp + " 3D " + wFlNm + wSz + ".jpg"
'          Exit Function
'        End If
'      End If
'
'      wSz = ""
'      wPth = gocn.GetFldVal("Select PDesc225 from Param Where PTyp='PCT' and PMCd='HSK'") + "\" + wDsgCtg
'      GetPictPath = wPth + "\" + wTyp + " LD " + wFlNm + wSz + ".jpg"
'    Else
'      GetPictPath = wPth + "\" + wTyp + " 3D " + wFlNm + wSz + ".jpg"
'    End If
'    Exit Function
'  End If
'  ' **** Zubin 211 **** '
'
'  If mLineDraw = True Then
'    wDir = Dir(wPth + "\" + wTyp + " LD " + wFlNm + wSz + ".jpg")
'    If wDir = "" Then wSz = ""
'    GetPictPath = wPth + "\" + wTyp + " LD " + wFlNm + wSz + ".jpg"
'  Else
'    wDir = Dir(wPth + "\" + wTyp + " 3D " + wFlNm + wSz + ".jpg")
'    If wDir = "" Then wSz = ""
'    GetPictPath = wPth + "\" + wTyp + " 3D " + wFlNm + wSz + ".jpg"
'  End If
'
'  Exit Function
'errh:
'  GetPictPath = ""
'  Exit Function
'End Function
'*** (Bef 20/06/05)


Public Sub ShowPic(ByVal mOldHt As Single, ByVal mOldWdt As Single, ByVal mPth As String, _
                   ByRef mPic As Image)
  Dim wPicNewHt As Single, wPicNewWdt As Single
  mPic.Visible = False                '*** First Make The Image Control Invisible ***
  mPic.Stretch = True                 '*** Set Stretch Property To True So That The Image Control Takes The Actual Picture Size ***
  If Dir(mPth) = "" Then Exit Sub     '*** Checking If The Picture Exists ***
  
  '*** (Jen 27/01/07)
  On Error GoTo ErrPic
  '*** (Jen 27/01/07)
  
  mPic.Picture = LoadPicture(mPth)  '*** Loading The Picture ***
  
  '*** Storing The Actual Picture Height And Width ***
  wPicNewHt = mPic.Picture.Height: wPicNewWdt = mPic.Picture.Width
  
  '*** If Pic Ht > Wdth Then The Pic Ht Is Set To Original Pic Control Ht And The Pic Wdth Is Set Proportionately ***
  '*** Follow Same Procedure If Wdth > Ht ***
  If wPicNewHt > wPicNewWdt Then
    mPic.Height = mOldHt
    mPic.Width = mOldWdt * MWLib.Div(wPicNewWdt, wPicNewHt)
  Else
    mPic.Width = mOldWdt
    mPic.Height = mOldHt * MWLib.Div(wPicNewHt, wPicNewWdt)
  End If
  mPic.Visible = True

Exit Sub
'*** (Jen 27/01/07)
ErrPic:
  If Err.Number = 481 Then
     DispMsg "Improper Picture Format, Please Change the Picture File", etError
  Else
     DispMsg Err.Description, etError
  End If
  mPic.Picture = LoadPicture("")
'*** (Jen 27/01/07)

End Sub

Public Sub ShowLargePic(ByVal mHt As Single, ByVal mWdt As Single, ByVal mPth As String, _
                        ByRef mPic As Image)
                        
  On Error GoTo ErrLrgPic
                        
  Dim wPicNewHt As Single, wPicNewWdt As Single
  mPic.Stretch = True                 '*** Set Stretch Property To True So That The Image Control Takes The Actual Picture Size ***
  If Dir(mPth) = "" Then Exit Sub     '*** Checking If The Picture Exists ***
  mPic.Picture = LoadPicture(mPth)    '*** Loading The Picture ***

  '*** Setting Picture Height And Width To The Actual Picture Height And Width ***
  wPicNewHt = mPic.Picture.Height: wPicNewWdt = mPic.Picture.Width
  
  '*** If Pic Ht > Wdth Then The Pic Ht Is Set To FraNKeyAll Ht And The Pic Wdth Is Set Proportionately ***
  If wPicNewHt > wPicNewWdt Then
    mPic.Height = mHt
    mPic.Width = mHt * MWLib.Div(wPicNewWdt, wPicNewHt)
  Else
    '*** If FraNKeyAll Ht >= Proportionate Ht when Wdth is taken as FraNKeyAll Wdth then Pic Wdth is Set to FraNKeyAll Wdth and The Height is Set Proportionately ***
    If mHt >= mWdt * MWLib.Div(wPicNewHt, wPicNewWdt) Then
      mPic.Width = mWdt
      mPic.Height = mWdt * MWLib.Div(wPicNewHt, wPicNewWdt)
    '*** If FraNKeyAll Ht < Proportionate Ht when Wdth is taken as FraNKeyAll Wdth then Pic Ht is Set to FraNKeyAll Ht and The Width is Set Proportionately ***
    Else
      mPic.Width = mHt * MWLib.Div(wPicNewWdt, wPicNewHt)
      mPic.Height = mHt
    End If
  End If
  'mPic.Move mLeft, mTop
  Exit Sub
ErrLrgPic:
  If Err.Number = 481 Then
     DispMsg "Improper Picture Format, Please Change the Picture File", etError
  Else
     DispMsg Err.Description, etError
  End If
  mPic.Picture = LoadPicture("")
  
End Sub

' Zubin 213 (Before 213)
'Public Sub GetDtVal(ByVal mDtTag As String, ByVal mOmDt As Variant, _
'                    Optional ByRef mDtVal As Variant, Optional ByRef mDtDt As Variant)
'  '*** Returns the max Date and corresponding Value from the Date Table (DtTbl) for the
'      'specified Date Tag and max date less than the specified order date
'  mDtVal = gocn.GetFldVal("Select DtVal from DtTbl where DtTag='" + mDtTag + "' and " + _
'           "DtDt=(Select max(DtDt) from DtTbl where DtTag='" + mDtTag + "' and " + _
'           "DtDt <= '" + CStr(mOmDt) + "') ")
'  mDtDt = gocn.GetFldVal("Select max(DtDt) from DtTbl where " + _
'          "DtTag='" + mDtTag + "' and DtDt <= '" + CStr(mOmDt) + "'")
'  If mDtDt = "" Then mDtDt = "01/01/80"
'
'End Sub

' Zubin 213
Public Sub GetDtVal(ByVal mDtTag As String, ByVal mOmDt As Variant, _
                    Optional ByRef mDtVal As Variant, Optional ByRef mDtDt As Variant, _
                    Optional ByVal mCurCd1 As Variant, Optional ByVal mCurCd2 As Variant)
  '*** Returns the max Date and corresponding Value from the Date Table (DtTbl) for the
      'specified Date Tag and max date less than the specified order date
  Dim ws_CurCnd As String
  ' ****** Manali 3.03 - 18/06/08 - "LML" included
  If Not IsMissing(mCurCd1) And (UCase(mDtTag) = "LMG" Or UCase(mDtTag) = "LMP" Or UCase(mDtTag) = "LMS" Or UCase(mDtTag) = "LML") Then
    ws_CurCnd = " and DtCurCd1= '" + mCurCd1 + "' "
  ElseIf Not IsMissing(mCurCd1) And Not IsMissing(mCurCd2) And UCase(mDtTag) = "CNV" Then
    ws_CurCnd = " and ((DtCurCd1= '" + mCurCd1 + "' and DtCurCd2= '" + mCurCd2 + "') Or " + _
              "(DtCurCd1= '" + mCurCd2 + "' and DtCurCd2= '" + mCurCd1 + "'))"
  End If

  mDtVal = goCn.GetFldVal("Select DtVal from DtTbl where DtTag='" + mDtTag + "' " + ws_CurCnd + " and " + _
           "DtDt=(Select max(DtDt) from DtTbl where DtTag='" + mDtTag + "' " + ws_CurCnd + " and " + _
           "DtDt <= '" + CStr(mOmDt) + "') ")
  mDtDt = goCn.GetFldVal("Select max(DtDt) from DtTbl where " + _
          "DtTag='" + mDtTag + "' " + ws_CurCnd + " and DtDt <= '" + CStr(mOmDt) + "'")
          
  If mDtDt = "" Then mDtDt = "01/01/80"

End Sub

'Public Sub GetDtVal(ByVal mDtTag As String, ByVal mOmDt As Variant, _
'                    Optional ByRef mDtVal As Variant, Optional ByRef mDtDt As Variant)
'  '*** Returns the max Date and corresponding Value from the Date Table (DtTbl) for the
'      'specified Date Tag and max date less than the specified order date
'  ' Zubin 213 (CurCd added)
'  mDtVal = gocn.GetFldVal("Select DtVal from DtTbl where DtTag='" + mDtTag + "' and " + _
'           "DtDt=(Select max(DtDt) from DtTbl where DtTag='" + mDtTag + "' " + _
'           IIF(UCase(mDtTag) = "LMG" Or UCase(mDtTag) = "LMP" Or UCase(mDtTag) = "LMS", " and DtCurCd1= '" + gocn.GetFldVal("Select HBaseCurCd from Head where HCd= '" + ctSelfCmCd + "' and HCoCd= '" + gs_CoCd + "'") + "' ", " ") + _
'           " and DtDt <= '" + CStr(mOmDt) + "') " + _
'           IIF(UCase(mDtTag) = "LMG" Or UCase(mDtTag) = "LMP" Or UCase(mDtTag) = "LMS", " and DtCurCd1= '" + gocn.GetFldVal("Select HBaseCurCd from Head where HCd= '" + ctSelfCmCd + "' and HCoCd= '" + gs_CoCd + "'") + "'", ""))
'  mDtDt = gocn.GetFldVal("Select max(DtDt) from DtTbl where " + _
'          "DtTag='" + mDtTag + "' and DtDt <= '" + CStr(mOmDt) + "' " + _
'          IIF(UCase(mDtTag) = "LMG" Or UCase(mDtTag) = "LMP" Or UCase(mDtTag) = "LMS", " and DtCurCd1= '" + gocn.GetFldVal("Select HBaseCurCd from Head where HCd= '" + ctSelfCmCd + "' and HCoCd= '" + gs_CoCd + "'") + "' ", " "))
'  If mDtDt = "" Then mDtDt = "01/01/80"
'
'End Sub

Public Function GetLmgSal(ByVal mTcTyp As String, ByVal mRrCmCd As String, _
                          ByVal mGldRt As Single, ByVal mCmCtg As String, _
                          ByVal mCmCurCd As String, _
                          Optional ByVal mCmLkUpRmRt As String) As Single
  Dim wInc As Double, wRrCtg As String, wCommStr As String
  Dim wCmSeekZSelfRt As Boolean
  
  ' **** Manali 3.03 - 02/07/08 - wCmLkUpRmRt shifted to function declaration
  'Dim wCmLkUpRmRt As String
  ' **** Manali 3.03 - 02/07/08 - wCmLkUpRmRt shifted to function declaration
  
  ' ***** Manali 3.03 - 18/06/08 - "L" and "LML" included
  mTcTyp = UCase(mTcTyp)
  wRrCtg = IIF(mTcTyp = "LMG", "G", IIF(mTcTyp = "LMP", "P", IIF(mTcTyp = "LMS", "S", IIF(mTcTyp = "LML", "L", ""))))
  GetLmgSal = 0
  
  '********** Increments on Self Record *********
  wInc = goCn.GetFldVal("Select RrSalRt from RmRt where RrTcTyp = 'INC' " + _
         "and RrCtg = '" + wRrCtg + "' and RrCmCtg= '" + mCmCtg + "' " + _
         "and RrCmCd = '" + mRrCmCd + "'")
'Manoj
'  wCmLkUpRmRt = gocn.GetFldVal("Select CmLkUpRmRt from CustMst where CmCtg= 'C' and CmCd= '" + mRrCmCd + "'")
  
  ' ***** Bef 3.03 - Manali
'  If UCase(mCmCtg) = "C" Or UCase(mCmCtg) = "S" Then
'    wCmLkUpRmRt = goCn.GetFldVal("Select CmLkUpRmRt from CustMst where CmCtg= '" & mCmCtg & "' and CmCd= '" + mRrCmCd + "'")
'  End If
  
  ' **** Manali 3.03 - 02/07/08
  '6.1-13 seek from custmst is removed
'  If mCmLkUpRmRt = "" And (UCase(mCmCtg) = "C" Or UCase(mCmCtg) = "S") Then
'    mCmLkUpRmRt = goCn.GetFldVal("Select CmLkUpRmRt from CustMst where CmCtg= '" & mCmCtg & "' and CmCd= '" + mRrCmCd + "'")
'  End If
  
  wCmSeekZSelfRt = goCn.RecSeek("Select 'x' from CustMst where CmCtg= '" & mCmCtg & "' and CmCd= '" + mRrCmCd + "' and CmSeekZSelfRt='Y'")
'Manoj
  
  ' Zubin 213 (mCmCurCd added)
  wCommStr = " Select RrSalRt from RmRt where RrTcTyp = '" + mTcTyp + "' and RrCtg= '" + wRrCtg + "' " + _
             "and RrCmCurCd= '" + mCmCurCd + "' and RrFrLn <= " + CStr(mGldRt) + " and RrToLn >= " + CStr(mGldRt) + " "
    
  '********** 1st get sale rate from PLCust (CmLkUpRmRt), then from Cust record;
  '********** then from Self (applying increments for the customer)
'Manoj
'  If UCase(mCmCtg) = "C" Then
  If UCase(mCmCtg) = "C" Or UCase(mCmCtg) = "S" Then
'Manoj
    GetLmgSal = goCn.GetFldVal("Select IsNull" + _
                "((" + wCommStr + _
                   "and RrCmCtg= 'P' and RrCmCd = '" + mCmLkUpRmRt + "' and RrSalRt<>0) " + _
                 ", " + _
                "IsNull((" + wCommStr + _
                   "and RrCmCtg= '" & mCmCtg & "' and RrCmCd = '" + mRrCmCd + "' and RrSalRt<>0) " + _
                 ", " + _
   IIF(wCmSeekZSelfRt, "IsNull((" + wCommStr + _
                   "and RrCmCtg= '" & mCmCtg & "' and RrCmCd  = '" + ctSelfCmCd + "' and RrSalRt<>0)*" + CStr(1 + wInc / 100) + ",0)", "0") + _
                 "))")
  ElseIf UCase(mCmCtg) = "T" Then
    GetLmgSal = goCn.GetFldVal(wCommStr + " and RrCmCtg= 'T' and RrCmCd = '" + mRrCmCd + "' and RrSalRt<> 0 ")
  End If

'*** Bef ver 1.10 ***
'  GetLmgSal = gocn.GetFldVal("Select IsNull" + _
'              "((Select RrSalRt from RmRt where RrTcTyp = 'LMG' and RrCtg= 'G' " + _
'                 "and RrFrLn <= " + CStr(mGldRt) + " and RrToLn >= " + CStr(mGldRt) + " " + _
'                 "and RrCmCd = '" + mRrCmCd + "' and RrSalRt<>0) " + _
'               ", " + _
'               "IsNull((Select RrSalRt from RmRt where RrTcTyp = 'LMG' and RrCtg= 'G' " + _
'                 "and RrFrLn <= " + CStr(mGldRt) + " and RrToLn >= " + CStr(mGldRt) + " " + _
'                 "and RrCmCd  = '" + ctSelfCmCd + "' and RrSalRt<>0)*" + CStr(1 + wInc / 100) + ",0)" + _
'               ")")
End Function
Public Function GetGldRt1(ByVal mCmCd As String, ByVal mLmg As Single, _
                         ByVal mRmCd As String, ByVal mAlyRt As Single, _
                         ByVal mDmCd As String, ByVal mOpt As String) As Single
'  '*** Bef ver 1.10 ***
'  Dim wDmCtg As String, wRmCtg As String, wRmSCtg As String, wRmPurity As String, wRt As String, wAlyRt As Double, wInc As Double, wGldRt As Double, wGldLs As Double
'  If (mOpt <> "C") And (mOpt <> "S") Then GetGldRt = 0: Exit Function
'  wDmCtg = gocn.GetFldVal("Select DmCtg from DsgMst where DmTcTyp= 'DM' " + _
'                         "and DmCd='" + mDmCd + "'")
'  wRmCtg = gocn.GetFldVal("Select RmCtg from RmMst where RmCd='" + mRmCd + "'")
'  wRmSCtg = gocn.GetFldVal("Select RmSCtg from RmMst where RmCd='" + mRmCd + "'")
'  wRmPurity = IIf(mOpt = "S", "RmPurityRt", "RmPurityWt")
'  wRt = IIf(mOpt = "S", "RrSalRt", "RrCstRt")
'
'  '******** AlyRt *********
'    '**** for Sale  = (1-RmPurityRt)* mAlyRt
'    '**** for Cost  = (1-RmPurityWt)* mAlyRt
'    wAlyRt = (1 - gocn.GetFldVal("Select " + wRmPurity + " from RmMst where RmCd='" + mRmCd + "'")) _
'             * mAlyRt
'
'  '********** Increments, only if mOpt='S'
'  If mOpt = "S" Then
'    wInc = gocn.GetFldVal("Select RrSalRt from RmRt where RrTcTyp= 'INC' " + _
'           "and RrCtg= '" + wRmCtg + "' and RrCmCd='" + mCmCd + "'")
'  Else
'    wInc = 0
'  End If
'
'  '********* looking for per gm record ('RM') for defined customer ***********
'  '***** Look for record with the specified DmCtg. If not found then look for DmCtg= ''
'  wGldRt = gocn.GetFldVal("Select isnull" + _
'           "((select " + wRt + "  from RmRt where RrTcTyp= 'RM' and " + _
'              "RrCtg= '" + wRmCtg + "' and RrSCtg= '" + wRmSCtg + "' and " + _
'              "RrCd='" + mRmCd + "' and RrFrLn <= " + CStr(mLmg) + " And " + _
'              "RrToLn >= " + CStr(mLmg) + " and RrCmCd = '" + mCmCd + "' and " + _
'              "RrDmCtg='" + wDmCtg + "' and " + wRt + " <> 0)" + _
'            ", " + _
'            "IsNull((Select " + wRt + "  from RmRt where RrTcTyp='RM' and " + _
'                     "RrCtg= '" + wRmCtg + "' and RrSCtg= '" + wRmSCtg + "' and " + _
'                     "RrCd='" + mRmCd + "' and RrFrLn <= " + CStr(mLmg) + " And " + _
'                     "RrToLn >= " + CStr(mLmg) + " and RrCmCd = '" + mCmCd + "' and " + _
'                     "RrDmCtg='' and " + wRt + " <> 0),0)" + _
'           ")")
'
'  '********* per gm record is found from Customer record,
'  If wGldRt <> 0 Then GetGldRt = (wGldRt + wAlyRt): Exit Function
'
'  '********* looking for per gm record for Self ***********
'  '***** Look for record with the specified DmCtg. If not found then look for DmCtg= ''
'  wGldRt = gocn.GetFldVal("Select isnull" + _
'           "((select " + wRt + " from RmRt where RrTcTyp='RM' and " + _
'              "RrCtg= '" + wRmCtg + "' and RrSCtg= '" + wRmSCtg + "' and " + _
'              "RrCd='" + mRmCd + "' and RrFrLn <= " + CStr(mLmg) + " And " + _
'              "RrToLn >= " + CStr(mLmg) + " and RrCmCd = '" + ctSelfCmCd + "' and " + _
'              "RrDmCtg='" + wDmCtg + "' and " + wRt + " <> 0)" + _
'             ", " + _
'             "IsNull((select " + wRt + " from RmRt where RrTcTyp='RM' and " + _
'                      "RrCtg= '" + wRmCtg + "' and RrSCtg= '" + wRmSCtg + "' and " + _
'                      "RrCd='" + mRmCd + "' and RrFrLn <= " + CStr(mLmg) + " And " + _
'                      "RrToLn >= " + CStr(mLmg) + " and RrCmCd ='" + ctSelfCmCd + "' and " + _
'                      "RrDmCtg='' and " + wRt + "<> 0),0)" + _
'           ")")
'
'  '********* per gm record is found from Self record,
'  If wGldRt <> 0 Then GetGldRt = (wGldRt * (1 + wInc / 100) + wAlyRt): Exit Function
'
'  '******** finding Gold Loss%age to be applied *********
'  '*** ( Look Cust Rate LS G record for DsgCtg
'  '***   then Cust Rate LS G record for DsgCtg = ''
'  '***   then Self Rate LS G record for DsgCtg
'  '***   then Self Rate LS G record for DsgCtg = '')
'
'  wGldLs = gocn.GetFldVal("Select isnull" + _
'           "((select " + wRt + " from RmRt where RrTcTyp='LS' and " + _
'             "RrCtg='" + wRmCtg + "' and RrCmCd = '" + mCmCd + "' and " + _
'             "RrDmCtg='" + wDmCtg + "' and " + wRt + "<> 0)" + _
'            ", " + _
'            "IsNull((select " + wRt + " from RmRt where RrTcTyp= 'LS' and " + _
'                    "RrCtg='" + wRmCtg + "' and RrCmCd = '" + mCmCd + "' and " + _
'                    "RrDmCtg='' and " + wRt + "<> 0)" + _
'                   ", " + _
'                   "IsNull((select " + wRt + " from RmRt where RrTcTyp='LS' and " + _
'                           "RrCtg='" + wRmCtg + "'and RrCmCd ='" + ctSelfCmCd + "' and " + _
'                           "RrDmCtg= '" + wDmCtg + "' and " + wRt + "<> 0)" + _
'                          ", " + _
'                          "IsNull((select " + wRt + " from RmRt where RrTcTyp='LS' and " + _
'                                  "RrCtg='" + wRmCtg + "' and RrCmCd = '" + ctSelfCmCd + "' and " + _
'                                  "RrDmCtg='' and " + wRt + " <> 0),0)" + _
'                          ")" + _
'                   ")" + _
'            ")")
'
'  '**** for Sale gold rate = (Lmg / Ounce) * RmPurityRt * (1+WGldLs) + wAlyRt
'  '**** for Cost gold rate = (Lmg / Ounce) * RmPurityWt * (1+WGldLs) + wAlyRt
'
'  wRmPurity = IIf(mOpt = "S", "RmPurityRt", "RmPurityWt")
'  wGldRt = MWLib.Div(mLmg, gocn.GetFldVal("Select PNum from param where PTyp=' ' " + _
'                           "and PMCd='" + wRmCtg + "'")) _
'           * gocn.GetFldVal("Select " + wRmPurity + " from RmMst where " + _
'             "RmCd='" + mRmCd + "'") _
'           * (1 + wGldLs / 100) + wAlyRt
'  GetGldRt = wGldRt
End Function

'*************Geeta**********Emr206********05/04/04
'''''  Public Function OrgGetRmRt(ByVal mCmCd As String, ByVal mRmCd As String, _
'''''                        ByVal mLn1 As Single, ByVal mLn2 As Single, _
'''''                        ByVal mWtPerDia As Single, ByVal mOpt As String, _
'''''                        Optional mCmCtg As String = "C") As Single
'''''  Dim wRmCtg As String, wRmSCtg As String, wRtBy As String, wRt As String, wInc As Double, wCnd As String
'''''  Dim wRtChrtByPtr As String, wCmLkUpRmRt As String, wRtChrtByBs As String
'''''  If (mOpt <> "C") And (mOpt <> "S") Then GetRmRt = 0: Exit Function
'''''  If (mCmCtg <> "C") And (mCmCtg <> "S") Then GetRmRt = 0: Exit Function
'''''
'''''
'''''  wRmCtg = gocn.GetFldVal("Select RmCtg from RmMst where RMCd='" + mRmCd + "'")
'''''  If wRmCtg = "G" Or wRmCtg = "P" Then GetRmRt = 0: Exit Function
'''''  wRmSCtg = gocn.GetFldVal("Select RmSCtg from RmMst where RmCd='" + mRmCd + "'")
'''''
'''''  wRtChrtByPtr = gocn.GetFldVal("Select 'P' from RmRt where RrTcTyp= 'PTR' and RrCmCtg= '" & mCmCtg & "' " + _
'''''                 "And RrCmCd = '" + mCmCd + "' and RrCtg= '" + wRmCtg + "' and " + _
'''''                 "RrSCtg= '" + wRmSCtg + "' and RrSTWGrp= '' and RrCd= '' and RrDmCtg= '' " + _
'''''                 "and RrLsCtg= '' and RrLabMCd= '' and RrFrLn= 0 and RrToLn= 0 ")
'''''  wRtChrtByBs = gocn.GetFldVal("Select 'BS' from RmRt where RrTcTyp= 'BS' and RrCmCtg= '" & mCmCtg & "' " + _
'''''                 "And RrCmCd = '" + mCmCd + "' and RrCtg= '" + wRmCtg + "' and " + _
'''''                 "RrSCtg= '" + wRmSCtg + "' and RrSTWGrp= '' and RrCd= '' and RrDmCtg= '' " + _
'''''                 "and RrLsCtg= '' and RrLabMCd= '' and RrFrLn= 0 and RrToLn= 0 ")
'''''
'''''  wRtBy = gocn.GetFldVal("Select PValue from Param where PTyp= 'RMSCTG' and " + _
'''''          "PMCd= '" + wRmCtg + "' and PSCd= '" + wRmSCtg + "'")
'''''  If wRtChrtByPtr <> "" Then
'''''    wRtBy = wRtChrtByPtr
'''''  ElseIf wRtChrtByBs <> "" Then
'''''    wRtBy = wRtChrtByBs
'''''  End If
'''''  wRt = IIF(mOpt = "S", "RrSalRt", "RrCstRt")
'''''
'''''  '********** Increments, only if mOpt='S' & Self Record *********
'''''  If mOpt = "S" Then
'''''    wInc = gocn.GetFldVal("Select RrSalRt from RmRt where RrTcTyp= 'INC' and RrCmCtg= '" & mCmCtg & "' and " + _
'''''           "RrCtg= '" + wRmCtg + "' and RrSCtg= '" + wRmSCtg + "' and RrCmCd= '" + mCmCd + "'")
'''''  Else
'''''    wInc = 0
'''''  End If
'''''
'''''  '****** 'S'  - By Size     (condition is RrFrLn <=  Ln1  <= RrToLn)
'''''  '****** 'P'  - By Pointer  (condition is RrFrLn <=  WtPerDia  <= RrToLn)
'''''  '****** 'BS' - Both Size   (condition is RrFrLn = Ln1  and  RrToLn = Ln2)
'''''
'''''  '*** By Size      - is used by diamonds i.e. only Len is used to calculate the rate
'''''  '*** By Pointer   - is used by Pointers i.e. Wt Per Diamond is used to calculate the rate
'''''  '*** By Both Size - is used by Col Stones i.e. Both Length and Breadth are used to calculate the rate
'''''
'''''  Select Case wRtBy
'''''  Case Is = "S"
'''''    wCnd = " Round(RrFrLn, 4) <= Round(" + CStr(mLn1) + ", 3) And Round(RrToLn, 4) >= Round(" + CStr(mLn1) + ", 3) and "
'''''  Case Is = "W", "P"
'''''    wCnd = " Round(RrFrLn, 4) <= Round(" + CStr(mWtPerDia) + ", 4) And Round(RrToLn, 4) >= Round(" + CStr(mWtPerDia) + ", 4) and "
'''''  Case Is = "BS"
'''''    wCnd = " Round(RrFrLn, 4) = Round(" + CStr(mLn1) + ", 3) And Round(RrToLn, 4) = Round(" + CStr(mLn2) + ", 3) and "
'''''  End Select
'''''
''''''  Select Case wRtBy
''''''  Case Is = "S"
''''''    wCnd = " RrFrLn <= " + CStr(mLn1) + " And RrToLn >= " + CStr(mLn1) + " and "
''''''  Case Is = "W", "P"
''''''    wCnd = " RrFrLn <= " + CStr(mWtPerDia) + " And RrToLn >= " + CStr(mWtPerDia) + " and "
''''''  Case Is = "BS"
''''''    wCnd = " RrFrLn = " + CStr(mLn1) + " And RrToLn = " + CStr(mLn2) + " and "
''''''  End Select
'''''
'''''
'''''  wCmLkUpRmRt = gocn.GetFldVal("Select CmLkUpRmRt from CustMst where CmCtg= '" & mCmCtg & "' and CmCd= '" + mCmCd + "'")
'''''  '********** 1st look for sale rate from PLCust, then from Cust record;
'''''  '********** if not found then from Self (applying increments for the customer)
'''''
'''''  If wRtChrtByPtr = "" And wRtChrtByBs = "" Then
'''''    ' Zubin 213 (ws_CmCurCd added)
'''''    GetRmRt = gocn.GetFldVal("Select " + _
'''''             "isnull((Select " + wRt + " from RmRt where RrTcTyp='RM' and " + _
'''''                     "RrCtg= '" + wRmCtg + "' and RrSCtg= '" + wRmSCtg + "' and " + _
'''''                     "RrCd='" + mRmCd + "' and " + wCnd + " RrCmCtg= 'P' and RrCmCd = '" + wCmLkUpRmRt + "' and " + _
'''''                     "RrCmCurCd= '" + ws_CmCurCd + "' and " + _
'''''                      wRt + " <> 0)" + _
'''''                    ", " + _
'''''                    "isnull((Select " + wRt + " from RmRt where RrTcTyp='RM' and " + _
'''''                            "RrCtg= '" + wRmCtg + "' and RrSCtg= '" + wRmSCtg + "' and " + _
'''''                            "RrCd='" + mRmCd + "' and " + wCnd + " RrCmCtg= '" & mCmCtg & "' and " + _
'''''                            "RrCmCd = '" + mCmCd + "' and RrCmCurCd= '" + ws_CmCurCd + "' and " + _
'''''                             wRt + " <> 0)" + _
'''''                           ", " + _
'''''                           "IsNull((Select " + wRt + " from RmRt where RrTcTyp= 'RM' and " + _
'''''                                    "RrCtg= '" + wRmCtg + "' and RrSCtg= '" + wRmSCtg + "' and " + _
'''''                                    "RrCd= '" + mRmCd + "' and " + wCnd + " RrCmCtg= '" & mCmCtg & "' and " + _
'''''                                    "RrCmCd = '" + ctSelfCmCd + "' and RrCmCurCd= '" + ws_CmCurCd + "' and " + _
'''''                                    wRt + " <> 0)*" + CStr(1 + wInc / 100) + ",0)" + _
'''''                    ")" + _
'''''             ")")
'''''  ElseIf UCase(wRtChrtByPtr) = "P" Or UCase(wRtChrtByBs) = "BS" Then
'''''    ' Zubin 213 (ws_CmCurCd added)
'''''    GetRmRt = gocn.GetFldVal("Select " + wRt + " from RmRt where RrTcTyp= 'RM' and " + _
'''''              "RrCtg= '" + wRmCtg + "' and RrSCtg= '" + wRmSCtg + "' and " + _
'''''              "RrCd='" + mRmCd + "' and " + wCnd + " RrCmCtg= '" & mCmCtg & "' and RrCmCd = '" + mCmCd + "' and " + _
'''''              "RrCmCurCd= '" + ws_CmCurCd + "' and " + _
'''''               wRt + " <> 0")
'''''  End If
'''''End Function
                         

Public Function GetRmRt(ByVal mCmCd As String, ByVal mRmCd As String, _
                        ByVal mLn1 As Single, ByVal mLn2 As Single, _
                        ByVal mWtPerDia As Single, ByVal mOpt As String, _
                        ByVal mLme As Single, Optional mCmCtg As String = "C", _
                        Optional ByVal mHistCmCtg As String, Optional ByVal mHistCmCd As String, _
                        Optional ByVal mHistYyyyMm As Double, _
                        Optional ByVal mCmCurCd As Variant, _
                        Optional ByVal mCmLkUpRmRt As String, _
                        Optional ByVal mPrtKey As String) As Single
  
  'ByVal mLme As Single    '*** (Jen 2.13)

  Dim wRmCtg As String, wRmSCtg As String, wRtBy As String, wRt As String, wInc As Double, wCnd As String, wRtByForCostRt As String
  Dim wRtChrtByPtr As String, wRtChrtByBs As String, wcndForCostRt As String
  Dim wCmSeekZSelfRt As Boolean
  Dim wInp As Double
  Dim wFrLn As String, wToLn As String
  m_ObjCmdGetRmRt.QryParametersByName("@mCmCd") = mCmCd
m_ObjCmdGetRmRt.QryParametersByName("@mRmCd") = mRmCd
m_ObjCmdGetRmRt.QryParametersByName("@mLn1") = mLn1
m_ObjCmdGetRmRt.QryParametersByName("@mLn2") = mLn2
m_ObjCmdGetRmRt.QryParametersByName("@mLme") = mLme
m_ObjCmdGetRmRt.QryParametersByName("@mCmCtg") = mCmCtg
m_ObjCmdGetRmRt.QryParametersByName("@mHistCmCtg") = mHistCmCtg
m_ObjCmdGetRmRt.QryParametersByName("@mHistCmCd") = mHistCmCd
m_ObjCmdGetRmRt.QryParametersByName("@mHistYyyyMm") = mHistYyyyMm
m_ObjCmdGetRmRt.QryParametersByName("@mCmCurCd") = mCmCurCd
m_ObjCmdGetRmRt.QryParametersByName("@mCmLkUpRmRt") = mCmLkUpRmRt
m_ObjCmdGetRmRt.QryParametersByName("@mPrtKey") = "C"
m_ObjCmdGetRmRt.QryParametersByName("@mWtPerDia") = mWtPerDia
m_ObjCmdGetRmRt.QryParametersByName("@mOpt") = mOpt
m_ObjCmdGetRmRt.QryParametersByName("@mrt") = 0

m_ObjCmdGetRmRt.QryExecute
GetRmRt = m_ObjCmdGetRmRt.QryParametersByName("@mrt")
If GetRmRt = 0 Then
  GetRmRt = 0
  Exit Function
End If

Dim wIpd As Double
wIpd = GetRmRt_IPD(mCmCd, mRmCd, _
                         mLn1, mLn2, _
                        mWtPerDia, mOpt, _
                        mLme, mCmCtg, _
                        mHistCmCtg, mHistCmCd, _
                        mHistYyyyMm, _
                        mCmCurCd, _
                        mCmLkUpRmRt, _
                        mPrtKey)
'  If wIpd < 0 Then wIpd = 0
  
  GetRmRt = GetRmRt + wIpd
  If GetRmRt < 0 Then GetRmRt = 0
Exit Function






  
  
  If (mOpt <> "C") And (mOpt <> "S") Then GetRmRt = 0: Exit Function
  If (mCmCtg <> "C") And (mCmCtg <> "S") Then GetRmRt = 0: Exit Function

  ' Zubin 213
  If IsMissing(mCmCurCd) Then mCmCurCd = goCn.GetFldVal("Select CmCurCd from CustMst where CmCtg= '" + mCmCtg + "' and CmCd= '" + mCmCd + "'")

  '*************************** ZUBIN **************************
   If (mHistCmCtg <> "" And (mHistYyyyMm = 0 Or mHistCmCd = "")) _
      Or (mHistCmCd <> "" And (mHistYyyyMm = 0 Or mHistCmCtg = "")) _
      Or (mHistYyyyMm <> 0 And (mHistCmCtg = "" Or mHistCmCd = "")) Then
    DispMsg "Proper Rate From History Cust Ctg, Cust Cd and YyyyMm Combination should be passed", etError
    Exit Function
  End If
  '*************************** ZUBIN **************************

  ' ***** Manali 3.03 - 02/07/08
  '6.1-13 seek from custmst is removed
'  If mCmLkUpRmRt = "" Then
'    mCmLkUpRmRt = goCn.GetFldVal("Select CmLkUpRmRt from CustMst where CmCtg= '" & mCmCtg & "' and CmCd= '" + mCmCd + "'")
'  End If
    wCmSeekZSelfRt = goCn.RecSeek("Select 'x' from CustMst where CmCtg= '" & mCmCtg & "' and CmCd= '" + mCmCd + "' and CmSeekZSelfRt='Y'")


  '****** Sachin 3.02.0 - Partition Key added
  ' ***** Manali 3.03 - 18/06/08 - RmCtg = "L" , "S" included
  wRmCtg = UCase(Trim(goCn.GetFldVal("Select RmCtg from RmMst where RMCd='" + mRmCd + "'" + IIF(mPrtKey = ctCurrPrtn, " and RmPrtKey='" + ctCurrPrtn + "'", ""))))
  If wRmCtg = "G" Or wRmCtg = "P" Or wRmCtg = "S" Or wRmCtg = "L" Then GetRmRt = 0: Exit Function
  wRmSCtg = goCn.GetFldVal("Select RmSCtg from RmMst where RmCd='" + mRmCd + "'" + IIF(mPrtKey = ctCurrPrtn, " and RmPrtKey='" + ctCurrPrtn + "'", ""))

  ' ***** Manali 3.9.1 - 20/02/12 - Round Off on incremented rate based on Head
  'SRITMP
  Dim ws_RndOff As String
  ws_RndOff = IIF(goCn.RecSeek("Select 'x' from Head " + _
                                "Join Param on PTyp='CPYRT' and PMCd=HOrdRtRndOffRmCtg " + _
                                "Where HCoCd='" + ctSelfCoCd + "' and hCd='" + ctSelfCmCd + "' " + _
                                "and ','+PDesc225+',' like '%," + wRmCtg + ",%' "), 0, 3)
  ' ***** Manali 3.9.1 - 20/02/12 - Round Off on incremented rate based on Head
   
   
  If mHistCmCtg = vbNullString Or mHistCmCtg <> "P" Then
    wRtChrtByPtr = goCn.GetFldVal("Select 'P' from RmRt where RrTcTyp= 'PTR' and RrCmCtg= '" & mCmCtg & "' " + _
                 "And RrCmCd = '" + mCmCd + "' and RrCtg= '" + wRmCtg + "' and " + _
                 "RrSCtg= '" + wRmSCtg + "' and RrSTWGrp= '' and RrCd= '' and RrDmCtg= '' " + _
                 "and RrLsCtg= '' and RrLabMCd= '' and RrFrLn= 0 and RrToLn= 0 " + IIF(mPrtKey = ctCurrPrtn, " and RrPrtKey='" + ctCurrPrtn + "'", ""))

    wRtChrtByBs = goCn.GetFldVal("Select 'BS' from RmRt where RrTcTyp= 'BS' and RrCmCtg= '" & mCmCtg & "' " + _
                 "And RrCmCd = '" + mCmCd + "' and RrCtg= '" + wRmCtg + "' and " + _
                 "RrSCtg= '" + wRmSCtg + "' and RrSTWGrp= '' and RrCd= '' and RrDmCtg= '' " + _
                 "and RrLsCtg= '' and RrLabMCd= '' and RrFrLn= 0 and RrToLn= 0 " + IIF(mPrtKey = ctCurrPrtn, " and RrPrtKey='" + ctCurrPrtn + "'", ""))
  End If

  wRtBy = goCn.GetFldVal("Select PValue from Param where PTyp= 'RMSCTG' and " + _
          "PMCd= '" + wRmCtg + "' and PSCd= '" + wRmSCtg + "'")

  wRtByForCostRt = wRtBy
  
  If wRtChrtByPtr <> "" Then
    wRtBy = wRtChrtByPtr
  ElseIf wRtChrtByBs <> "" Then
    wRtBy = wRtChrtByBs
  End If

  '*************************** ZUBIN **************************
  If mHistCmCtg = vbNullString Then
    wRt = IIF(mOpt = "S", "RrSalRt", "RrCstRt")
  Else
    wRt = IIF(mOpt = "S", "RhSalRt", "RhCstRt")
  End If
  '*************************** ZUBIN **************************

  ' ***** Manali 3.8.0 - Cost Prc for CmCtg='S' from CmCtg='C'
  If mCmCtg = "S" And mOpt = "C" Then mCmCtg = "C"

  '********** Increments *********
  If mOpt = "S" Then
    wInc = goCn.GetFldVal("Select RrSalRt from RmRt where RrTcTyp= 'INC' and RrCmCtg= '" & mCmCtg & "' and " + _
           "RrCtg= '" + wRmCtg + "' and RrSCtg= '" + wRmSCtg + "' and RrCmCd= '" + mCmCd + "'" + IIF(mPrtKey = ctCurrPrtn, " and RrPrtKey='" + ctCurrPrtn + "'", ""))
    If wInc = 0 Then _
      wInc = goCn.GetFldVal("Select RrSalRt from RmRt where RrTcTyp= 'INC' and RrCmCtg= '" & mCmCtg & "' and " + _
           "RrCtg= '" + wRmCtg + "' and RrSCtg= '" + "" + "' and RrCmCd= '" + mCmCd + "'" + IIF(mPrtKey = ctCurrPrtn, " and RrPrtKey='" + ctCurrPrtn + "'", ""))
    If wInc = 0 Then _
      wInc = goCn.GetFldVal("Select RrSalRt from RmRt where RrTcTyp= 'INC' and RrCmCtg= '" & mCmCtg & "' and " + _
           "RrCtg= '" + wRmCtg + "' and RrSCtg= '" + wRmSCtg + "' and RrCmCd= '" + ctSelfCmCd + "'" + IIF(mPrtKey = ctCurrPrtn, " and RrPrtKey='" + ctCurrPrtn + "'", ""))
    If wInc = 0 Then _
      wInc = goCn.GetFldVal("Select RrSalRt from RmRt where RrTcTyp= 'INC' and RrCmCtg= '" & mCmCtg & "' and " + _
           "RrCtg= '" + wRmCtg + "' and RrSCtg= '" + "" + "' and RrCmCd= '" + ctSelfCmCd + "'" + IIF(mPrtKey = ctCurrPrtn, " and RrPrtKey='" + ctCurrPrtn + "'", ""))
  Else
    wInc = 0
  End If
  If mOpt = "S" Then
    wInp = goCn.GetFldVal("Select RrSalRt from RmRt where RrTcTyp= 'INP' and RrCmCtg= '" & "P" & "' and " + _
           "RrCtg= '" + wRmCtg + "' and RrSCtg= '" + wRmSCtg + "' and RrCmCd= '" + mCmLkUpRmRt + "'" + IIF(mPrtKey = ctCurrPrtn, " and RrPrtKey='" + ctCurrPrtn + "'", ""))
    If wInp = 0 Then _
      wInp = goCn.GetFldVal("Select RrSalRt from RmRt where RrTcTyp= 'INP' and RrCmCtg= '" & "P" & "' and " + _
           "RrCtg= '" + wRmCtg + "' and RrSCtg= '" + "" + "' and RrCmCd= '" + mCmLkUpRmRt + "'" + IIF(mPrtKey = ctCurrPrtn, " and RrPrtKey='" + ctCurrPrtn + "'", ""))
  Else
    wInp = 0
  End If

  '*************************** ZUBIN **************************
  'Dim wFrLn As String, wToLn As String

  If mHistCmCtg = vbNullString Then
    wFrLn = "RrFrLn": wToLn = "RrToLn"
  Else
    wFrLn = "RhFrLn": wToLn = "RhToLn"
  End If

  Select Case wRtBy
  Case Is = "S"
    '*** (Jen 2.13)
    If wRmCtg = "X" Then mLn1 = mLme
    '*** (Jen 2.13)
  
    wCnd = " Round(" + wFrLn + ", 4) <= Round(" + CStr(mLn1) + ", 3) And Round(" + wToLn + ", 4) >= Round(" + CStr(mLn1) + ", 3) and "
  Case Is = "W", "P"
    wCnd = " Round(" + wFrLn + ", 4) <= Round(" + CStr(mWtPerDia) + ", 4) And Round(" + wToLn + ", 4) >= Round(" + CStr(mWtPerDia) + ", 4) and "
  Case Is = "BS"
    wCnd = " Round(" + wFrLn + ", 4) = Round(" + CStr(mLn1) + ", 3) And Round(" + wToLn + ", 4) = Round(" + CStr(mLn2) + ", 3) and "
  End Select

  Select Case wRtByForCostRt
  Case Is = "S"
    '*** (Jen 2.13)
    If wRmCtg = "X" Then mLn1 = mLme

    wcndForCostRt = " Round(" + wFrLn + ", 4) <= Round(" + CStr(mLn1) + ", 3) And Round(" + wToLn + ", 4) >= Round(" + CStr(mLn1) + ", 3) and "
  Case Is = "W", "P"
    wcndForCostRt = " Round(" + wFrLn + ", 4) <= Round(" + CStr(mWtPerDia) + ", 4) And Round(" + wToLn + ", 4) >= Round(" + CStr(mWtPerDia) + ", 4) and "
  Case Is = "BS"
    wcndForCostRt = " Round(" + wFrLn + ", 4) = Round(" + CStr(mLn1) + ", 3) And Round(" + wToLn + ", 4) = Round(" + CStr(mLn2) + ", 3) and "
  End Select

  If wRtChrtByPtr = "" And wRtChrtByBs = "" Then
    If mHistCmCtg <> vbNullString Then
      GetRmRt = goCn.GetFldVal("Select " + _
              "isnull((Select " + wRt + " from RmRtHist where " + _
                      "RhYyyyMm= " + CStr(mHistYyyyMm) + " and RhTcTyp='RM' and " + _
                      "RhCtg= '" + wRmCtg + "' and RhSCtg= '" + wRmSCtg + "' and " + _
                      "RhCd='" + mRmCd + "' and " + wCnd + " RhCmCtg= '" + mHistCmCtg + "' and " + _
                      "RhCmCd='" + mHistCmCd + "' and RhCmCurCd= '" + mCmCurCd + "' and " + wRt + " <> 0)" + _
              ", " + _
      IIF(wCmSeekZSelfRt Or mOpt = "C", _
                      "Convert(Decimal(14," + IIF(wInc = 0, "3", ws_RndOff) + "),IsNull((Select " + wRt + " from RmRtHist where " + _
                      "RhYyyyMm= " + CStr(mHistYyyyMm) + " and RhTcTyp= 'RM' and " + _
                      "RhCtg= '" + wRmCtg + "' and RhSCtg= '" + wRmSCtg + "' and " + _
                      "RhCd= '" + mRmCd + "' and " + wCnd + " RhCmCtg= 'C' and RhCmCd = '" + ctSelfCmCd + "' and " + _
                      "RhCmCurCd= '" + mCmCurCd + "' and " + _
                      wRt + " <> 0)*" + CStr(1 + wInc / 100) + ",0))", "0") + _
                      ")")
    Else
      Dim wAccessoriesRmIncFac As String
      wAccessoriesRmIncFac = "(case when '" + wRmCtg + "' ='X' then ((" + CStr(mLme) + " -RrFrLn) * " + IIF(mOpt = "C", "RrFixMinTol", "RrFixMaxTol") + ") else 0 end) "
      GetRmRt = goCn.GetFldVal("Select " + _
             "isnull((Select " + wRt + "+" + wAccessoriesRmIncFac + _
                     " from RmRt where RrTcTyp='RM' and " + IIF(mPrtKey = ctCurrPrtn, " RrPrtKey='" + ctCurrPrtn + "' and ", "") + _
                     "RrCtg= '" + wRmCtg + "' and RrSCtg= '" + wRmSCtg + "' and " + _
                     "RrCd='" + mRmCd + "' and " + wCnd + " RrCmCtg= 'P' and RrCmCd = '" + mCmLkUpRmRt + "' and " + _
                     "RrCmCurCd= '" + mCmCurCd + "' and " + _
                      wRt + " <> 0)*" + CStr(1 + wInp / 100) + _
                    ", " + _
                    "isnull((Select " + wRt + "+" + wAccessoriesRmIncFac + _
                    " from RmRt where RrTcTyp='RM' and " + IIF(mPrtKey = ctCurrPrtn, " RrPrtKey='" + ctCurrPrtn + "' and ", "") + _
                            "RrCtg= '" + wRmCtg + "' and RrSCtg= '" + wRmSCtg + "' and " + _
                            "RrCd='" + mRmCd + "' and " + wCnd + " RrCmCtg= '" & mCmCtg & "' and " + _
                            "RrCmCd = '" + mCmCd + "' and RrCmCurCd= '" + mCmCurCd + "' and " + _
                             wRt + " <> 0)" + _
                           ", " + _
                  IIF(wCmSeekZSelfRt Or mOpt = "C", "Convert(Decimal(14," + IIF(wInc = 0, "3", ws_RndOff) + "),IsNull((Select " + wRt + "+" + wAccessoriesRmIncFac + _
                  " from RmRt where RrTcTyp= 'RM' and " + IIF(mPrtKey = ctCurrPrtn, " RrPrtKey='" + ctCurrPrtn + "' and ", "") + _
                                    "RrCtg= '" + wRmCtg + "' and RrSCtg= '" + wRmSCtg + "' and " + _
                                    "RrCd= '" + mRmCd + "' and " + wCnd + " RrCmCtg= '" & mCmCtg & "' and " + _
                                    "RrCmCd = '" + ctSelfCmCd + "' and RrCmCurCd= '" + mCmCurCd + "' and " + _
                                    wRt + " <> 0)*" + CStr(1 + wInc / 100) + ",0))", "0") + _
                    ")" + _
             ")")
    End If
  ElseIf UCase(wRtChrtByPtr) = "P" Or UCase(wRtChrtByBs) = "BS" Then
    If mHistCmCtg <> vbNullString Then
      If (mOpt = "C") Then
        GetRmRt = goCn.GetFldVal("Select " + wRt + " from RmRtHist " + _
                "where RhYyyyMm= " + CStr(mHistYyyyMm) + " and RhTcTyp= 'RM' and " + _
                "RhCtg= '" + wRmCtg + "' and RhSCtg= '" + wRmSCtg + "' and " + _
                "RhCd='" + mRmCd + "' and " + wcndForCostRt + " RhCmCtg= 'C' and RhCmCd = '" + _
                ctSelfCmCd + "' and RhCmCurCd= '" + mCmCurCd + "' and " + wRt + " <> 0")
      
      Else
        GetRmRt = goCn.GetFldVal("Select " + wRt + " from RmRtHist " + _
                "where RhYyyyMm= " + CStr(mHistYyyyMm) + " and RhTcTyp= 'RM' and " + _
                "RhCtg= '" + wRmCtg + "' and RhSCtg= '" + wRmSCtg + "' and " + _
                "RhCd='" + mRmCd + "' and " + wCnd + " RhCmCtg= 'C' and RhCmCd = '" + _
                mCmCd + "' and RhCmCurCd= '" + mCmCurCd + "' and " + wRt + " <> 0")
      End If
    Else
      If (mOpt = "C") Then
        GetRmRt = goCn.GetFldVal("Select isnull((Select " + wRt + " from RmRt where RrTcTyp='RM' and " + IIF(mPrtKey = ctCurrPrtn, " RrPrtKey='" + ctCurrPrtn + "' and ", "") + _
                            "RrCtg= '" + wRmCtg + "' and RrSCtg= '" + wRmSCtg + "' and " + _
                            "RrCd='" + mRmCd + "' and " + wCnd + " RrCmCtg= '" & mCmCtg & "' and " + _
                            "RrCmCd = '" + mCmCd + "' and RrCmCurCd= '" + mCmCurCd + "' and " + _
                             wRt + " <> 0)," + _
                        "isnull((Select " + wRt + " from RmRt where RrTcTyp= 'RM' and " + IIF(mPrtKey = ctCurrPrtn, " RrPrtKey='" + ctCurrPrtn + "' and ", "") + _
              "RrCtg= '" + wRmCtg + "' and RrSCtg= '" + wRmSCtg + "' and " + _
              "RrCd='" + mRmCd + "' and " + wcndForCostRt + " RrCmCtg= '" & mCmCtg & "' and RrCmCd = '" + ctSelfCmCd + "' and " + _
              "RrCmCurCd= '" + mCmCurCd + "' and " + _
               wRt + " <> 0),0))")
      Else
        GetRmRt = goCn.GetFldVal("Select " + wRt + " from RmRt where RrTcTyp= 'RM' and " + IIF(mPrtKey = ctCurrPrtn, " RrPrtKey='" + ctCurrPrtn + "' and ", "") + _
              "RrCtg= '" + wRmCtg + "' and RrSCtg= '" + wRmSCtg + "' and " + _
              "RrCd='" + mRmCd + "' and " + wCnd + " RrCmCtg= '" & mCmCtg & "' and RrCmCd = '" + mCmCd + "' and " + _
              "RrCmCurCd= '" + mCmCurCd + "' and " + _
               wRt + " <> 0")
      End If
    End If
  End If
End Function
Public Function GetRmRt_base(ByVal mCmCd As String, ByVal mRmCd As String, _
                        ByVal mLn1 As Single, ByVal mLn2 As Single, _
                        ByVal mWtPerDia As Single, ByVal mOpt As String, _
                        ByVal mLme As Single, Optional mCmCtg As String = "C", _
                        Optional ByVal mHistCmCtg As String, Optional ByVal mHistCmCd As String, _
                        Optional ByVal mHistYyyyMm As Double, _
                        Optional ByVal mCmCurCd As Variant, _
                        Optional ByVal mCmLkUpRmRt As String, _
                        Optional ByVal mPrtKey As String) As Single
  'ByVal mLme As Single    '*** (Jen 2.13)
  Dim wRmCtg As String, wRmSCtg As String, wRtBy As String, wRt As String, wInc As Double, wCnd As String, wRtByForCostRt As String
  Dim wRtChrtByPtr As String, wRtChrtByBs As String, wcndForCostRt As String
  Dim wCmSeekZSelfRt As Boolean
  Dim wInp As Double
  Dim wFrLn As String, wToLn As String
  m_objCmdGetRmRt_Base.QryParametersByName("@mCmCd") = mCmCd
  m_objCmdGetRmRt_Base.QryParametersByName("@mRmCd") = mRmCd
  m_objCmdGetRmRt_Base.QryParametersByName("@mLn1") = mLn1
  m_objCmdGetRmRt_Base.QryParametersByName("@mLn2") = mLn2
  m_objCmdGetRmRt_Base.QryParametersByName("@mLme") = mLme
m_objCmdGetRmRt_Base.QryParametersByName("@mCmCtg") = mCmCtg
m_objCmdGetRmRt_Base.QryParametersByName("@mHistCmCtg") = mHistCmCtg
m_objCmdGetRmRt_Base.QryParametersByName("@mHistCmCd") = mHistCmCd
m_objCmdGetRmRt_Base.QryParametersByName("@mHistYyyyMm") = mHistYyyyMm
m_objCmdGetRmRt_Base.QryParametersByName("@mCmCurCd") = mCmCurCd
m_objCmdGetRmRt_Base.QryParametersByName("@mCmLkUpRmRt") = mCmLkUpRmRt
m_objCmdGetRmRt_Base.QryParametersByName("@mPrtKey") = "C"
m_objCmdGetRmRt_Base.QryParametersByName("@mWtPerDia") = mWtPerDia
m_objCmdGetRmRt_Base.QryParametersByName("@mOpt") = mOpt
m_objCmdGetRmRt_Base.QryParametersByName("@mrt") = 0

m_objCmdGetRmRt_Base.QryExecute
GetRmRt_base = m_objCmdGetRmRt_Base.QryParametersByName("@mrt")

End Function


Public Function GetRmRt_IPD(ByVal mCmCd As String, ByVal mRmCd As String, _
                        ByVal mLn1 As Single, ByVal mLn2 As Single, _
                        ByVal mWtPerDia As Single, ByVal mOpt As String, _
                        ByVal mLme As Single, Optional mCmCtg As String = "C", _
                        Optional ByVal mHistCmCtg As String, Optional ByVal mHistCmCd As String, _
                        Optional ByVal mHistYyyyMm As Double, _
                        Optional ByVal mCmCurCd As Variant, _
                        Optional ByVal mCmLkUpRmRt As String, _
                        Optional ByVal mPrtKey As String) As Single
  
  'ByVal mLme As Single    '*** (Jen 2.13)

  Dim wRmCtg As String, wRmSCtg As String, wRtBy As String, wRt As String, wInc As Double, wCnd As String, wRtByForCostRt As String
  Dim wRtChrtByPtr As String, wRtChrtByBs As String, wcndForCostRt As String
  Dim wCmSeekZSelfRt As Boolean
  Dim wInp As Double
  Dim wFrLn As String, wToLn As String
  Dim wBase As Double
  
  wBase = GetRmRt_base(mCmCd, mRmCd, _
                        mLn1, mLn2, _
                        mWtPerDia, mOpt, _
                        mLme, mCmCtg, _
                        mHistCmCtg, mHistCmCd, _
                        mHistYyyyMm, _
                        mCmCurCd, _
                        mCmLkUpRmRt, _
                        mPrtKey)
                        
                        
  
  m_ObjCmdGetRmRt_IPD.QryParametersByName("@mCmCd") = mCmCd
m_ObjCmdGetRmRt_IPD.QryParametersByName("@mRmCd") = mRmCd
m_ObjCmdGetRmRt_IPD.QryParametersByName("@mLn1") = mLn1
m_ObjCmdGetRmRt_IPD.QryParametersByName("@mLn2") = mLn2
m_ObjCmdGetRmRt_IPD.QryParametersByName("@mLme") = mLme
m_ObjCmdGetRmRt_IPD.QryParametersByName("@mCmCtg") = mCmCtg
m_ObjCmdGetRmRt_IPD.QryParametersByName("@mHistCmCtg") = mHistCmCtg
m_ObjCmdGetRmRt_IPD.QryParametersByName("@mHistCmCd") = mHistCmCd
m_ObjCmdGetRmRt_IPD.QryParametersByName("@mHistYyyyMm") = mHistYyyyMm
m_ObjCmdGetRmRt_IPD.QryParametersByName("@mCmCurCd") = mCmCurCd
m_ObjCmdGetRmRt_IPD.QryParametersByName("@mCmLkUpRmRt") = mCmLkUpRmRt
m_ObjCmdGetRmRt_IPD.QryParametersByName("@mPrtKey") = "C"
m_ObjCmdGetRmRt_IPD.QryParametersByName("@mWtPerDia") = mWtPerDia
m_ObjCmdGetRmRt_IPD.QryParametersByName("@mOpt") = mOpt
m_ObjCmdGetRmRt_IPD.QryParametersByName("@mrt") = 0

m_ObjCmdGetRmRt_IPD.QryExecute
GetRmRt_IPD = m_ObjCmdGetRmRt_IPD.QryParametersByName("@mrt")
Exit Function






  
End Function






'**********************<<Previous Emr205***********
Public Function GetRmRt2(ByVal mCmCd As String, ByVal mRmCd As String, _
                        ByVal mLn1 As Single, ByVal mLn2 As Single, _
                        ByVal mWtPerDia As Single, ByVal mOpt As String) As Single
'  Dim wRmCtg As String, wRmSCtg As String, wRtBy As String, wRt As String, wInc As Double, wCnd As String
'  Dim wRtChrtByPtr As String, wCmLkUpRmRt As String
'  If (mOpt <> "C") And (mOpt <> "S") Then GetRmRt = 0: Exit Function
'
'  wRmCtg = gocn.GetFldVal("Select RmCtg from RmMst where RMCd='" + mRmCd + "'")
'  If wRmCtg = "G" Or wRmCtg = "P" Then GetRmRt = 0: Exit Function
'  wRmSCtg = gocn.GetFldVal("Select RmSCtg from RmMst where RmCd='" + mRmCd + "'")
'
'  wRtChrtByPtr = gocn.GetFldVal("Select 'P' from RmRt where RrTcTyp= 'PTR' and RrCmCtg= 'C' " + _
'                 "And RrCmCd = '" + mCmCd + "' and RrCtg= '" + wRmCtg + "' and " + _
'                 "RrSCtg= '" + wRmSCtg + "' and RrSTWGrp= '' and RrCd= '' and RrDmCtg= '' " + _
'                 "and RrLsCtg= '' and RrLabMCd= '' and RrFrLn= 0 and RrToLn= 0 ")
'
'  wRtBy = gocn.GetFldVal("Select PValue from Param where PTyp= 'RMSCTG' and " + _
'          "PMCd= '" + wRmCtg + "' and PSCd= '" + wRmSCtg + "'")
'  If wRtChrtByPtr <> "" Then wRtBy = wRtChrtByPtr
'  wRt = IIf(mOpt = "S", "RrSalRt", "RrCstRt")
'
'  '********** Increments, only if mOpt='S' & Self Record *********
'  If mOpt = "S" Then
'    wInc = gocn.GetFldVal("Select RrSalRt from RmRt where RrTcTyp= 'INC' and RrCmCtg= 'C' and " + _
'           "RrCtg= '" + wRmCtg + "' and RrSCtg= '" + wRmSCtg + "' and RrCmCd= '" + mCmCd + "'")
'  Else
'    wInc = 0
'  End If
'
'  '****** 'S'  - By Size     (condition is RrFrLn <=  Ln1  <= RrToLn)
'  '****** 'P'  - By Pointer  (condition is RrFrLn <=  WtPerDia  <= RrToLn)
'  '****** 'BS' - Both Size   (condition is RrFrLn = Ln1  and  RrToLn = Ln2)
'
'  '*** By Size      - is used by diamonds i.e. only Len is used to calculate the rate
'  '*** By Pointer   - is used by Pointers i.e. Wt Per Diamond is used to calculate the rate
'  '*** By Both Size - is used by Col Stones i.e. Both Length and Breadth are used to calculate the rate
'
'  Select Case wRtBy
'  Case Is = "S"
'    wCnd = " Round(RrFrLn, 4) <= Round(" + CStr(mLn1) + ", 3) And Round(RrToLn, 4) >= Round(" + CStr(mLn1) + ", 3) and "
'  Case Is = "W", "P"
'    wCnd = " Round(RrFrLn, 4) <= Round(" + CStr(mWtPerDia) + ", 4) And Round(RrToLn, 4) >= Round(" + CStr(mWtPerDia) + ", 4) and "
'  Case Is = "BS"
'    wCnd = " Round(RrFrLn, 4) = Round(" + CStr(mLn1) + ", 3) And Round(RrToLn, 4) = Round(" + CStr(mLn2) + ", 3) and "
'  End Select
'
''  Select Case wRtBy
''  Case Is = "S"
''    wCnd = " RrFrLn <= " + CStr(mLn1) + " And RrToLn >= " + CStr(mLn1) + " and "
''  Case Is = "W", "P"
''    wCnd = " RrFrLn <= " + CStr(mWtPerDia) + " And RrToLn >= " + CStr(mWtPerDia) + " and "
''  Case Is = "BS"
''    wCnd = " RrFrLn = " + CStr(mLn1) + " And RrToLn = " + CStr(mLn2) + " and "
''  End Select
'
'
'  wCmLkUpRmRt = gocn.GetFldVal("Select CmLkUpRmRt from CustMst where CmCtg= 'C' and CmCd= '" + mCmCd + "'")
'  '********** 1st look for sale rate from PLCust, then from Cust record;
'  '********** if not found then from Self (applying increments for the customer)
'
'  If wRtChrtByPtr = "" Then
'    GetRmRt = gocn.GetFldVal("Select " + _
'             "isnull((Select " + wRt + " from RmRt where RrTcTyp='RM' and " + _
'                     "RrCtg= '" + wRmCtg + "' and RrSCtg= '" + wRmSCtg + "' and " + _
'                     "RrCd='" + mRmCd + "' and " + wCnd + " RrCmCtg= 'P' and RrCmCd = '" + wCmLkUpRmRt + "' and " + _
'                      wRt + " <> 0)" + _
'                    ", " + _
'                    "isnull((Select " + wRt + " from RmRt where RrTcTyp='RM' and " + _
'                            "RrCtg= '" + wRmCtg + "' and RrSCtg= '" + wRmSCtg + "' and " + _
'                            "RrCd='" + mRmCd + "' and " + wCnd + " RrCmCtg= 'C' and RrCmCd = '" + mCmCd + "' and " + _
'                             wRt + " <> 0)" + _
'                           ", " + _
'                           "IsNull((Select " + wRt + " from RmRt where RrTcTyp= 'RM' and " + _
'                                    "RrCtg= '" + wRmCtg + "' and RrSCtg= '" + wRmSCtg + "' and " + _
'                                    "RrCd= '" + mRmCd + "' and " + wCnd + " RrCmCtg= 'C' and RrCmCd = '" + ctSelfCmCd + "' and " + _
'                                    wRt + " <> 0)*" + CStr(1 + wInc / 100) + ",0)" + _
'                    ")" + _
'             ")")
'  ElseIf UCase(wRtChrtByPtr) = "P" Then
'    GetRmRt = gocn.GetFldVal("Select " + wRt + " from RmRt where RrTcTyp= 'RM' and " + _
'              "RrCtg= '" + wRmCtg + "' and RrSCtg= '" + wRmSCtg + "' and " + _
'              "RrCd='" + mRmCd + "' and " + wCnd + " RrCmCtg= 'C' and RrCmCd = '" + mCmCd + "' and " + _
'               wRt + " <> 0")
'  End If
End Function
Public Function GetWtFrSz(ByVal mRmCtg As String, ByVal mRmSCtg As String, _
                          ByVal mLn1 As Single, ByVal mLn2 As Single, _
                          ByVal mRmSTWGrp As String) As Single
  '****** This function is used to get the per dia Wt based on the size entered by the user
  '****** Calculation of Wt Per Qty is also dependent on the RATEBY option of specified Rm Sub Ctg
  '*** If 'S'(By Size)  - is used by diamonds i.e. only Len is used to calculate the Wt per Qty
  '*** If 'BS'(By Both Size)  - is used by Col Stones i.e. Both Length and Breadth are used to calculate the Wt per Qty
  
  Dim wRt As String, wRtBy As String, wCnd As String
  wRt = "RrSalRt"
  ' **** Manali 3.03 - 18/06/08 - RmCtg = "S" and "L" included
  '6.1 new below commented
  'If mRmCtg = "G" Or mRmCtg = "P" Or mRmCtg = "S" Or mRmCtg = "L" Then GetWtFrSz = 0: Exit Function
  
  wRtBy = goCn.GetFldVal("Select PValue from Param where PTyp='RMSCTG' and " + _
          "PMCd='" + mRmCtg + "' and PSCd='" + mRmSCtg + "'")
  If wRtBy <> "S" And wRtBy <> "BS" And wRtBy <> "P" Then GetWtFrSz = 0: Exit Function
  
  Select Case wRtBy
  Case Is = "S", "P"
    wCnd = " RrFrLn <= " + CStr(mLn1) + " And RrToLn >= " + CStr(mLn1) + " "
  Case Is = "BS"
    wCnd = " RrFrLn = " + CStr(mLn1) + " And RrToLn = " + CStr(mLn2) + " "
  End Select

'  GetWtFrSz = gocn.GetFldVal("Select RrSalRt from RmRt where RrTcTyp='STW' and " + _
'                    "RrCmCd= '" + ctSelfCmCd + "' and RrCtg='" + mRmCtg + "' and " + _
'                    "RrSCtg='" + mRmSCtg + "' and " + wCnd)
    '**********Geeta*****Emr207*******Fetchs wt from 'BSW' for wRtBy='S'
  '  If wRtBy = "S" Then
    GetWtFrSz = goCn.GetFldVal("Select IsNull" + _
                    "((Select RrSalRt from RmRt where RrTcTyp='BSW' and RrCmCtg= 'C' and " + _
                    "RrCmCd= '" + ctSelfCmCd + "' and RrCtg='" + mRmCtg + "' and " + _
                    "RrSCtg='" + mRmSCtg + "' and RrSTWGrp= '" + mRmSTWGrp + "' and " + _
                    "RrFrLn = " + CStr(mLn1) + " And RrToLn = " + CStr(mLn2) + ") , " + _
                    "IsNull((Select RrSalRt from RmRt where RrTcTyp='BSW' and RrCmCtg= 'C' and " + _
                            "RrCmCd= '" + ctSelfCmCd + "' and RrCtg='" + mRmCtg + "' and " + _
                            "RrSCtg='" + mRmSCtg + "' and RrSTWGrp= '' and " + _
                            "RrFrLn = " + CStr(mLn1) + " And RrToLn = " + CStr(mLn2) + "), 0))")
   '**********Geeta*****Emr207*******
      
  If GetWtFrSz = 0 Then _
  GetWtFrSz = goCn.GetFldVal("Select IsNull" + _
                    "((Select RrSalRt from RmRt where RrTcTyp='STW' and RrCmCtg= 'C' and " + _
                    "RrCmCd= '" + ctSelfCmCd + "' and RrCtg='" + mRmCtg + "' and " + _
                    "RrSCtg='" + mRmSCtg + "' and RrSTWGrp= '" + mRmSTWGrp + "' and " + wCnd + ") " + _
                    ", " + _
                    "IsNull((Select RrSalRt from RmRt where RrTcTyp='STW' and RrCmCtg= 'C' and " + _
                            "RrCmCd= '" + ctSelfCmCd + "' and RrCtg='" + mRmCtg + "' and " + _
                            "RrSCtg='" + mRmSCtg + "' and RrSTWGrp= '' and " + wCnd + "), 0))")
End Function
Public Function GetWtFrSz1(ByVal mRmCtg As String, ByVal mRmSCtg As String, _
                          ByVal mLn1 As Single, ByVal mLn2 As Single) As Single
'  ''*** Code Bef ver 1.10 ***
'  '****** This function is used to get the per dia Wt based on the size entered by the user
'  '****** Calculation of Wt Per Qty is also dependent on the RATEBY option of specified Rm Sub Ctg
'  '*** If 'S'(By Size)  - is used by diamonds i.e. only Len is used to calculate the Wt per Qty
'  '*** If 'BS'(By Both Size)  - is used by Col Stones i.e. Both Length and Breadth are used to calculate the Wt per Qty
'
'  Dim wRt As String, wRtBy As String, wCnd As String
'  wRt = "RrSalRt"
'  If mRmCtg = "G" Or mRmCtg = "P" Then GetWtFrSz = 0: Exit Function
'
'  wRtBy = gocn.GetFldVal("Select PValue from Param where PTyp='RMSCTG' and " + _
'          "PMCd='" + mRmCtg + "' and PSCd='" + mRmSCtg + "'")
'  If wRtBy <> "S" And wRtBy <> "BS" Then GetWtFrSz = 0: Exit Function
'
'  Select Case wRtBy
'  Case Is = "S"
'    wCnd = " RrFrLn <= " + CStr(mLn1) + " And RrToLn >= " + CStr(mLn1) + " "
'  Case Is = "BS"
'    wCnd = " RrFrLn = " + CStr(mLn1) + " And RrToLn = " + CStr(mLn2) + " "
'  End Select
'
''  GetWtFrSz = gocn.GetFldVal("Select RrSalRt from RmRt where RrTcTyp='STW' and " + _
''                    "RrCmCd= '" + ctSelfCmCd + "' and RrCtg='" + mRmCtg + "' and " + _
''                    "RrSCtg='" + mRmSCtg + "' and " + wCnd)
End Function
'**************Geeta************Emr206********02/04/04****
Public Function GetLabRt(ByVal mCmCd As String, ByVal mLabMCd As String, _
                         ByVal mLabSCd As String, ByVal mOpt As String, ByVal mGldAsWt As Single, _
                         ByVal mCmCtg As String, Optional ByRef mLabQw As Variant, Optional ByVal mOldGldAsWt As Single = 0#, Optional ByRef mWtRngChg As Variant, _
                         Optional ByVal mCmCurCd As Variant, _
                         Optional ByVal mCmLkUpLabRt As String) As Single
  Dim wRt As String, wInc As Double, wCommStr As String, wCommQwStr As String
  Dim wCnd As String, wCnd1 As String, wCnd2 As String
  Dim wLabQw As String
  Dim wRsOldLabRt As MDORowSet, wRsNewLabRt As MDORowSet
  'Dim wQRt As Single, wWRt As Single, wQW As String
  Dim wCmSeekZSelfRt As Boolean
  Dim wRmCtg As String, wRmSCtg As String, wRtBy As String, wRtByForCostRt As String
  Dim wRtChrtByPtr As String, wRtChrtByBs As String, wcndForCostRt As String
  'Dim wCmSeekZSelfRt As Boolean
  Dim wInp As Double
  'Dim wCmLkUpRmRt As String    ' **** Manali 3.03 - 02/07/08 - variable shifted to function parameters
  
    '*************************** ZUBIN **************************
    ' 05th Dec 2003, EMR206

  Dim wFrLn As String, wToLn As String

'  If m_ObjCmdGetLabRt Is Nothing Then
'  End If
  m_ObjCmdGetLabRt.QryParametersByName("@mCmCd") = mCmCd
m_ObjCmdGetLabRt.QryParametersByName("@mCmCtg") = mCmCtg

m_ObjCmdGetLabRt.QryParametersByName("@mLabMCd") = mLabMCd
m_ObjCmdGetLabRt.QryParametersByName("@mLabSCd") = mLabSCd
m_ObjCmdGetLabRt.QryParametersByName("@mLabQw") = IIF(IsMissing(mLabQw), "", mLabQw)
m_ObjCmdGetLabRt.QryParametersByName("@mOpt") = mOpt
m_ObjCmdGetLabRt.QryParametersByName("@mOldGldAsWt") = mOldGldAsWt

m_ObjCmdGetLabRt.QryParametersByName("@mGldAsWt") = mGldAsWt
m_ObjCmdGetLabRt.QryParametersByName("@mWtRngChg") = ""
m_ObjCmdGetLabRt.QryParametersByName("@mCmCurCd") = mCmCurCd
m_ObjCmdGetLabRt.QryParametersByName("@mCmLkUpLabRt") = mCmLkUpLabRt
m_ObjCmdGetLabRt.QryParametersByName("@mrt") = 0


m_ObjCmdGetLabRt.QryExecute
mWtRngChg = m_ObjCmdGetLabRt.QryParametersByName("@mWtRngChg")
GetLabRt = m_ObjCmdGetLabRt.QryParametersByName("@mrt")
'Exit Function
  
  
  ' **** Manali 3.03 - 02/07/08 - variable shifted to func declaration
  'Dim wCmLkUpLabRt As String
  ' **** Manali 3.03 - 02/07/08 - variable shifted to func declaration
  
  If (mOpt <> "C") And (mOpt <> "S") Then GetLabRt = 0: Exit Function
  wRt = IIF(mOpt = "S", "LrSalRt", "LrCstRt")

  ' Zubin 213
  If IsMissing(mCmCurCd) Then mCmCurCd = goCn.GetFldVal("Select CmCurCd from CustMst where CmCtg= '" + mCmCtg + "' and CmCd= '" + mCmCd + "'")
  
  ' *** Bef 3.03 - Manali *****
  'wCmLkUpLabRt = goCn.GetFldVal("Select CmLkUpLabRt from CustMst where CmCtg= '" + mCmCtg + "' and CmCd= '" + mCmCd + "'")
  
  ' ****** Manali 3.03 - 02/07/08
  '6.1-13 seek from custmst is removed
  'If mCmLkUpLabRt = "" Then _
    mCmLkUpLabRt = goCn.GetFldVal("Select CmLkUpLabRt from CustMst where CmCtg= '" + mCmCtg + "' and CmCd= '" + mCmCd + "'")
    
  wCmSeekZSelfRt = goCn.RecSeek("Select 'x' from CustMst where CmCtg= '" & mCmCtg & "' and CmCd= '" + mCmCd + "' and CmSeekZSelfRt='Y'")
  
  ' ***** Manali 3.8.0 - Cost Prc for CmCtg='S' from CmCtg='C'
  If mCmCtg = "S" And mOpt = "C" Then mCmCtg = "C"
  
  '********** 1st look for Lab sale rate from Cust record;
  '********** if not found then from Self (applying increments for the customer)
  '********** Increments on Self Record applied on Sale Rate *********
  ' Zubin 213 (mCmCurCd cnd added)
  wCommStr = " Select " + wRt + " from LabRt where LrMCd='" + mLabMCd + "' and " + _
             "LrSCd= '" + mLabSCd + "' and LrCmCurCd= '" + mCmCurCd + "' "
' **** Zubin 211 **** '
'''''  wCnd = IIF(mLabMCd = "SET", "and ((LrFrWt<= " + CStr(mGldAsWt) + " and LrToWt>= " + CStr(mGldAsWt) + " and LrQW='W') Or LrQW='Q') ", " and LrFrWt<= " + CStr(mGldAsWt) + " and LrToWt>= " + CStr(mGldAsWt))
  wCnd = " and LrFrWt<= " + CStr(mGldAsWt) + " and LrToWt>= " + CStr(mGldAsWt)
 ' **** Zubin 211 **** '
  If mLabMCd = "GHC" Then
    wCnd = " and 1=1"
  End If
  If mOpt = "S" Then
    wInc = goCn.GetFldVal("Select RrSalRt from RmRt where RrTcTyp= 'INL' and " + _
           "RrLabMcd= '" + mLabMCd + "' and RrCmCtg= '" + mCmCtg + "' and RrCmCd= '" + mCmCd + "'")
     
    ' ***** Manali Trading Module - 'S' added below - RrCmCtg='C' for 'ZSELF' replaced by mCmCtg (for Suppier)
    If UCase(mCmCtg) = "C" Or UCase(mCmCtg) = "S" Then
      GetLabRt = goCn.GetFldVal("Select " + _
                  "IsNull ((" + wCommStr + " and LrCmCtg= 'P' and LrCmCd='" + mCmLkUpLabRt + "' " + _
                          wCnd + "), " + _
                 "IsNull((" + wCommStr + " and LrCmCtg='" + mCmCtg + "' /*'C'*/ and LrCmCd='" + mCmCd + "' " + _
                         wCnd + "), " + _
                  IIF(wCmSeekZSelfRt Or mOpt = "C", "IsNull((" + wCommStr + " and LrCmCtg='" + mCmCtg + "' /*'C'*/ and LrCmCd='" + ctSelfCmCd + "' " + _
                          wCnd + ")" + _
                          "* " + CStr(1 + wInc / 100) + ",0)", "0") + " ))")
    ElseIf UCase(mCmCtg) = "T" Then
      GetLabRt = goCn.GetFldVal("Select " + _
                 "IsNull((" + wCommStr + " and LrCmCtg= 'T' and LrCmCd= '" + mCmCd + "' " + _
                          wCnd + "),'')")
    End If
  Else
    'wInc = 0
'    sritmp
'GetLabRt = goCn.GetFldVal("Select " + _
               "IsNull((" + wCommStr + " and LrCmCtg= '" + mCmCtg + "' and LrCmCd= '" + ctSelfCmCd + "' " + _
                        wCnd + "),'')")
      GetLabRt = goCn.GetFldVal("Select " + _
                  "IsNull ((" + wCommStr + " and LrCmCtg= 'P' and LrCmCd='" + mCmLkUpLabRt + "' " + _
                          wCnd + " and LrCstRt <> 0 " + "), " + _
                 "IsNull((" + wCommStr + " and LrCmCtg='" + mCmCtg + "' /*'C'*/ and LrCmCd='" + mCmCd + "' " + _
                         wCnd + " and LrCstRt <> 0 " + "), " + _
                  IIF(wCmSeekZSelfRt Or mOpt = "C", "IsNull((" + wCommStr + " and LrCmCtg='" + mCmCtg + "' /*'C'*/ and LrCmCd='" + ctSelfCmCd + "' " + _
                          wCnd + " and LrCstRt <> 0 " + ")" + _
                          "* " + CStr(1 + wInc / 100) + ",0)", "0") + " ))")
  
  End If
                      
  If (Not IsMissing(mLabQw)) Or (Not IsMissing(mWtRngChg)) Then
    ' Zubin 213 (mCmCurCd cnd added)
    wCommQwStr = " Select LrQw from LabRt where LrMCd= '" + mLabMCd + "' and " + _
                 "LrSCd= '" + mLabSCd + "' and LrCmCurCd= '" + mCmCurCd + "' "
    If mOpt = "S" Then
      ' ***** Manali Trading Module - 'S' added below - RrCmCtg='C' for 'ZSELF' replaced by mCmCtg (for Suppier)
      If UCase(mCmCtg) = "C" Or UCase(mCmCtg) = "S" Then
        wLabQw = goCn.GetFldVal("Select " + _
                  "IsNull((" + wCommQwStr + " and LrCmCtg= 'P' and LrCmCd='" + mCmLkUpLabRt + "' " + _
                         wCnd + "), " + _
                 "IsNull((" + wCommQwStr + " and LrCmCtg='" + mCmCtg + "' /*'C'*/ and LrCmCd='" + mCmCd + "' " + _
                         wCnd + "), " + _
    IIF(wCmSeekZSelfRt Or mOpt = "C", "IsNull((" + wCommQwStr + " and LrCmCtg='" + mCmCtg + "' /*'C'*/ and LrCmCd='" + ctSelfCmCd + "' " + _
                         wCnd + "),'')", "''") + " ))")
      ElseIf UCase(mCmCtg) = "T" Then
        wLabQw = goCn.GetFldVal("Select " + _
                 "IsNull((" + wCommQwStr + " and LrCmCtg= 'T' and LrCmCd= '" + mCmCd + "' " + _
                          wCnd + "),'')")
      End If
    ElseIf wCmSeekZSelfRt Or mOpt = "C" Then
      wLabQw = goCn.GetFldVal("Select " + _
               "IsNull((" + wCommQwStr + " and LrCmCtg= '" + mCmCtg + "' and LrCmCd='" + _
                        ctSelfCmCd + "' " + wCnd + "),'')")
    End If
  End If
 '******************Geeta********Emr207*****
 '******************Check Wt Range Changed/Not
 If (Not IsMissing(mLabQw)) Then _
    mLabQw = wLabQw
    
  If Not IsMissing(mWtRngChg) Then    'And mOldGldAsWt <> 0 And mGldAsWt <> mOldGldAsWt
    If wLabQw = "Q" Or wLabQw = "W" Then
     ' Zubin 213 (mCmCurCd cnd added)
     wCommQwStr = " Select LrCmCtg, LrCmCd,LrMCd, LrSCd, LrFrWt, LrToWt from LabRt where LrMCd= '" + mLabMCd + "' and " + _
                 "LrSCd= '" + mLabSCd + "' and LrCmCurCd= '" + mCmCurCd + "' "
     wCnd1 = " and LrFrWt<= " + CStr(mOldGldAsWt) + " and LrToWt>= " + CStr(mOldGldAsWt)
     wCnd2 = " and LrFrWt<= " + CStr(mGldAsWt) + " and LrToWt>= " + CStr(mGldAsWt)
    If mOpt = "S" Then
      ' ***** Manali Trading Module - 'S' added below - RrCmCtg='C' for 'ZSELF' replaced by mCmCtg (for Suppier)
      If UCase(mCmCtg) = "C" Or UCase(mCmCtg) = "S" Then
        Set wRsOldLabRt = goCn.OpenResultset(wCommQwStr + " and LrCmCtg= 'P' and LrCmCd='" + mCmLkUpLabRt + "' " + wCnd1)
        If (wRsOldLabRt.EOF Or wRsOldLabRt.BOF) Then _
          Set wRsOldLabRt = goCn.OpenResultset(wCommQwStr + " and LrCmCtg='" + mCmCtg + "' /*'C'*/ and LrCmCd='" + mCmCd + "' " + wCnd1)
        If (wRsOldLabRt.EOF Or wRsOldLabRt.BOF) And wCmSeekZSelfRt Then _
           Set wRsOldLabRt = goCn.OpenResultset(wCommQwStr + " and LrCmCtg='" + mCmCtg + "' /*'C'*/ and LrCmCd='" + ctSelfCmCd + "' " + wCnd1)
       
       
       Set wRsNewLabRt = goCn.OpenResultset(wCommQwStr + " and LrCmCtg= 'P' and LrCmCd='" + mCmLkUpLabRt + "' " + wCnd2)
        If (wRsNewLabRt.EOF Or wRsNewLabRt.BOF) Then _
          Set wRsNewLabRt = goCn.OpenResultset(wCommQwStr + " and LrCmCtg='" + mCmCtg + "' /*'C'*/ and LrCmCd='" + mCmCd + "' " + wCnd2)
        If (wRsNewLabRt.EOF Or wRsNewLabRt.BOF) And wCmSeekZSelfRt Then _
           Set wRsNewLabRt = goCn.OpenResultset(wCommQwStr + " and LrCmCtg='" + mCmCtg + "' /*'C'*/ and LrCmCd='" + ctSelfCmCd + "' " + wCnd2)
       
      '  mWtRngChg = gocn.GetFldVal("Select " + _
                  "IsNull((" + wCommQwStr + " and LrCmCtg= 'P' and LrCmCd='" + mCmLkUpLabRt + "' " + _
                         wCnd + "), " + _
                 "IsNull((" + wCommQwStr + " and LrCmCtg= 'C' and LrCmCd='" + mCmCd + "' " + _
                         wCnd + "), " + _
                 "IsNull((" + wCommQwStr + " and LrCmCtg= 'C' and LrCmCd='" + ctSelfCmCd + "' " + _
                         wCnd + "),'T')))")
      ElseIf UCase(mCmCtg) = "T" Then
        Set wRsOldLabRt = goCn.OpenResultset(wCommQwStr + " and LrCmCtg= 'T' and LrCmCd= '" + mCmCd + "' " + wCnd1)
        Set wRsNewLabRt = goCn.OpenResultset(wCommQwStr + " and LrCmCtg= 'T' and LrCmCd= '" + mCmCd + "' " + wCnd2)
       ' mWtRngChg = gocn.GetFldVal("Select " + _
                 "IsNull((" + wCommQwStr + " and LrCmCtg= 'T' and LrCmCd= '" + mCmCd + "' " + _
                          wCnd + "),'T')")
                          
      End If
    Else
    Set wRsOldLabRt = goCn.OpenResultset(wCommQwStr + " and LrCmCtg= '" + mCmCtg + "' and LrCmCd='" + ctSelfCmCd + "' " + wCnd1)
    Set wRsNewLabRt = goCn.OpenResultset(wCommQwStr + " and LrCmCtg= '" + mCmCtg + "' and LrCmCd='" + ctSelfCmCd + "' " + wCnd2)
      ' mWtRngChg = gocn.GetFldVal("Select " + _
               "IsNull((" + wCommQwStr + " and LrCmCtg= '" + mCmCtg + "' and LrCmCd='" + _
                        ctSelfCmCd + "'" + wCnd + "),'T')")
    End If
    If wRsOldLabRt.RecCount > 0 And wRsNewLabRt.RecCount > 0 Then
      If wRsOldLabRt!LrCmCtg = wRsNewLabRt!LrCmCtg And wRsOldLabRt!LrCmCd = wRsNewLabRt!LrCmCd And _
          wRsOldLabRt!LrMCd = wRsNewLabRt!LrMCd And wRsNewLabRt!LrSCd = wRsNewLabRt!LrSCd And _
          wRsOldLabRt!LrFrWt = wRsNewLabRt!LrFrWt And wRsOldLabRt!LrToWt = wRsNewLabRt!LrToWt Then
          mWtRngChg = "F"
       Else
          mWtRngChg = "T"
       End If
    Else
      If wRsOldLabRt.RecCount > 0 Then
       mWtRngChg = "F"
      Else
       mWtRngChg = "T"
      End If
    End If
    Set wRsOldLabRt = Nothing
    Set wRsNewLabRt = Nothing
  Else
     mWtRngChg = "F"
 End If
End If
'******************geeta********emr207*****
End Function
'**********************Geeta**********************
Public Sub BeginEmrProcess(pv_Frm As Form, pv_Msg As String)
  g_InProcessingMode = True
  LockWindowUpdate pv_Frm.hwnd
  pv_Frm.Enabled = False
  GetMDIFrm.Enabled = False
'  BlockInput True
  'MWLib.BeginProcess pv_Frm, pv_Msg
End Sub
Public Sub EndEmrProcess(pv_Frm As Form)
  LockWindowUpdate 0
  'MWLib.EndProcess pv_Frm
  g_InProcessingMode = False
  pv_Frm.Enabled = True
  GetMDIFrm.Enabled = True
'  BlockInput False
End Sub
Public Sub UpdateEmrProcess(pv_Msg As String)
  DispMsg "[PrcsUpdate]" + pv_Msg, etInfo
  DoEvents
End Sub

Public Function IsItUpdateEmrProcessMsg(ByRef pv_Msg As String)
  If InStr(1, pv_Msg, "[PrcsUpdate]") > 0 Then
    pv_Msg = Mid(pv_Msg, 13)
    IsItUpdateEmrProcessMsg = True
  Else
    IsItUpdateEmrProcessMsg = False
  End If
End Function


'**********************<<Previous Emr205***********
Public Function GetLabRt2(ByVal mCmCd As String, ByVal mLabMCd As String, _
                         ByVal mLabSCd As String, ByVal mOpt As String, ByVal mGldAsWt As Single, _
                         ByVal mCmCtg As String, Optional ByRef mLabQw As Variant) As Single
'  Dim wRt As String, wInc As Double, wCmLkUpLabRt As String, wCommStr As String, wCommQwStr As String
'  'Dim wQRt As Single, wWRt As Single, wQW As String
'
'
'  If (mOpt <> "C") And (mOpt <> "S") Then GetLabRt = 0: Exit Function
'  wRt = IIf(mOpt = "S", "LrSalRt", "LrCstRt")
'
'  wCmLkUpLabRt = gocn.GetFldVal("Select CmLkUpLabRt from CustMst where CmCtg= '" + mCmCtg + "' and CmCd= '" + mCmCd + "'")
'  '********** 1st look for Lab sale rate from Cust record;
'  '********** if not found then from Self (applying increments for the customer)
'  '********** Increments on Self Record applied on Sale Rate *********
'  wCommStr = " Select " + wRt + " from LabRt where LrMCd='" + mLabMCd + "' and " + _
'             "LrSCd= '" + mLabSCd + "' "
'
'  If mOpt = "S" Then
'    wInc = gocn.GetFldVal("Select RrSalRt from RmRt where RrTcTyp= 'INL' and " + _
'           "RrLabMcd= '" + mLabMCd + "' and RrCmCtg= '" + mCmCtg + "' and RrCmCd= '" + mCmCd + "'")
'
'    If UCase(mCmCtg) = "C" Then
'      GetLabRt = gocn.GetFldVal("Select " + _
'                  "IsNull ((" + wCommStr + " and LrCmCtg= 'P' and LrCmCd='" + wCmLkUpLabRt + "'" + _
'                      " and LrQW= 'Q'), " + _
'                 "IsNull ((" + wCommStr + " and LrCmCtg= 'P' and LrCmCd='" + wCmLkUpLabRt + "'" + _
'                          " and LrFrWt<= " + CStr(mGldAsWt) + " and LrToWt>= " + CStr(mGldAsWt) + ")" + _
'                          ", " + _
'                 "IsNull((" + wCommStr + " and LrCmCtg= 'C' and LrCmCd='" + mCmCd + "'" + _
'                         " and LrQW= 'Q'), " + _
'                 "IsNull((" + wCommStr + " and LrCmCtg= 'C' and LrCmCd='" + mCmCd + "'" + _
'                         " and LrFrWt<= " + CStr(mGldAsWt) + " and LrToWt>= " + CStr(mGldAsWt) + ")" + _
'                         ", " + _
'                  "IsNull((" + wCommStr + " and LrCmCtg= 'C' and LrCmCd='" + ctSelfCmCd + "' " + _
'                         " and LrQW= 'Q') * " + CStr(1 + wInc / 100) + ", " + _
'                 "IsNull((" + wCommStr + " and LrCmCtg= 'C' and LrCmCd='" + ctSelfCmCd + "' " + _
'                          " and LrFrWt<= " + CStr(mGldAsWt) + " and LrToWt>= " + CStr(mGldAsWt) + ")" + _
'                          "* " + CStr(1 + wInc / 100) + ",0))))))")
'    ElseIf UCase(mCmCtg) = "T" Then
'      GetLabRt = gocn.GetFldVal("Select " + _
'                 "IsNull((" + wCommStr + " and LrCmCtg= 'T' and LrCmCd= '" + mCmCd + "' " + _
'                          "and LrFrWt<= " + CStr(mGldAsWt) + " and LrToWt>= " + CStr(mGldAsWt) + "), " + _
'                 "IsNull((" + wCommStr + " and LrCmCtg= 'T' and LrCmCd= '" + mCmCd + "' " + _
'                          "and LrQW= 'Q'), ''))")
'    End If
'  Else
'    'wInc = 0
'      GetLabRt = gocn.GetFldVal("Select " + _
'                 "IsNull((" + wCommStr + " and LrCmCtg= '" + mCmCtg + "' and LrCmCd= '" + ctSelfCmCd + "' " + _
'                          "and LrFrWt<= " + CStr(mGldAsWt) + " and LrToWt>= " + CStr(mGldAsWt) + "), " + _
'                 "IsNull((" + wCommStr + " and LrCmCtg= '" + mCmCtg + "' and LrCmCd= '" + ctSelfCmCd + "' " + _
'                          "and LrQW= 'Q'), ''))")
'  End If
'
'  If Not IsMissing(mLabQw) Then
'    wCommQwStr = " Select LrQw from LabRt where LrMCd= '" + mLabMCd + "' and " + _
'                 "LrSCd= '" + mLabSCd + "' "
'    If mOpt = "S" Then
'      If UCase(mCmCtg) = "C" Then
'        mLabQw = gocn.GetFldVal("Select " + _
'                  "IsNull((" + wCommQwStr + " and LrCmCtg= 'P' and LrCmCd='" + wCmLkUpLabRt + "' " + _
'                         " and LrQW= 'Q'), " + _
'                 "IsNull((" + wCommQwStr + " and LrCmCtg= 'P' and LrCmCd='" + wCmLkUpLabRt + "' " + _
'                         " and LrFrWt<= " + CStr(mGldAsWt) + " and LrToWt>= " + CStr(mGldAsWt) + "), " + _
'                 "IsNull((" + wCommQwStr + " and LrCmCtg= 'C' and LrCmCd='" + mCmCd + "' " + _
'                         " and LrQW= 'Q'), " + _
'                 "IsNull((" + wCommQwStr + " and LrCmCtg= 'C' and LrCmCd='" + mCmCd + "' " + _
'                         " and LrFrWt<= " + CStr(mGldAsWt) + " and LrToWt>= " + CStr(mGldAsWt) + "), " + _
'                 "IsNull((" + wCommQwStr + " and LrCmCtg= 'C' and LrCmCd='" + ctSelfCmCd + "' " + _
'                         " and LrQW= 'Q'), " + _
'                "IsNull((" + wCommQwStr + " and LrCmCtg= 'C' and LrCmCd='" + ctSelfCmCd + "' " + _
'                         " and LrFrWt<= " + CStr(mGldAsWt) + " and LrToWt>= " + CStr(mGldAsWt) + "), " + _
'                " ''))))))")
'      ElseIf UCase(mCmCtg) = "T" Then
'        mLabQw = gocn.GetFldVal("Select " + _
'                 "IsNull((" + wCommQwStr + " and LrCmCtg= 'T' and LrCmCd= '" + mCmCd + "' " + _
'                          "and LrFrWt<= " + CStr(mGldAsWt) + " and LrToWt>= " + CStr(mGldAsWt) + "), " + _
'                 "IsNull((" + wCommQwStr + " and LrCmCtg= 'T' and LrCmCd= '" + mCmCd + "' " + _
'                          "and LrQW= 'Q'), ''))")
'      End If
'    Else
'      mLabQw = gocn.GetFldVal("Select " + _
'               "IsNull((" + wCommQwStr + " and LrCmCtg= '" + mCmCtg + "' and LrCmCd='" + _
'                        ctSelfCmCd + "' and LrFrWt<= " + CStr(mGldAsWt) + _
'                        " and LrToWt>= " + CStr(mGldAsWt) + "), " + _
'               "IsNull((" + wCommQwStr + " and LrCmCtg= '" + mCmCtg + "' and LrCmCd='" + _
'                        ctSelfCmCd + "' and LrQW= 'Q'), ''))")
'    End If
'  End If
End Function
Public Function GetLabRt1(ByVal mCmCd As String, ByVal mLabMCd As String, _
                         ByVal mLabSCd As String, ByVal mOpt As String) As Single
'  '*** Before 1.10 release
'  Dim wRt As String, wInc As Double
'  If (mOpt <> "C") And (mOpt <> "S") Then GetLabRt = 0: Exit Function
'  wRt = IIf(mOpt = "S", "LrSalRt", "LrCstRt")
'
'  '********** 1st look for Lab sale rate from Cust record;
'  '********** if not found then from Self (applying increments for the customer)
'  '********** Increments on Self Record applied on Sale Rate *********
'  If mOpt = "S" Then
'    wInc = gocn.GetFldVal("Select RrSalRt from RmRt where RrTcTyp= 'INL' and " + _
'           "RrLabMcd= '" + mLabMCd + "' and RrCmCd= '" + mCmCd + "'")
'  Else
'    wInc = 0
'  End If
'
''*** Old Code (Does not consider records with Rate= 0)
''  GetLabRt = gocn.GetFldVal("Select IsNull" + _
''             "((Select " + wRt + " from LabRt where LrMCd='" + mLabMCd + "' and " + _
''               "LrSCd='" + mLabSCd + "' and LrCmCd='" + mCmCd + "' and " + wRt + "<> 0)" + _
''              ", " + _
''              "IsNull((Select " + wRt + " from LabRt where LrMCd='" + mLabMCd + "' and " + _
''                      "LrSCd='" + mLabSCd + "' and LrCmCd='" + ctSelfCmCd + "' and " + _
''                      wRt + "<> 0)*" + CStr(1 + wInc / 100) + ",0))")
'
'  '*** New Code (Considers records with Rate= 0)
'  GetLabRt = gocn.GetFldVal("Select IsNull" + _
'             "((Select " + wRt + " from LabRt where LrMCd='" + mLabMCd + "' and " + _
'               "LrSCd='" + mLabSCd + "' and LrCmCd='" + mCmCd + "' )" + _
'              ", " + _
'              "IsNull((Select " + wRt + " from LabRt where LrMCd='" + mLabMCd + "' and " + _
'                      "LrSCd='" + mLabSCd + "' and LrCmCd='" + ctSelfCmCd + "') " + _
'                      "* " + CStr(1 + wInc / 100) + ",0))")
'
End Function
Public Function GetLabVal(ByVal mCmCd As String, ByVal mLabMCd As String, _
                          ByVal mLabSCd As String, ByVal mLabRt As Single, _
                          ByVal mLabQty As Single, ByVal mGldAsWt As Single, _
                          ByVal mLabQw As String, ByVal mOpt As String, _
                          ByVal mCmCtg As String, ByVal mCmCurCd As String, _
                          Optional ByVal mCmLkUpLabRt As String) As Single
  Dim wLabVal As Double, wMin As Double, wMax As Double
  Dim wMaxCommStr As String, wMinCommStr As String
  Dim wCmSeekZSelfRt As Boolean
  
  
  ' **** Manali 3.03 - variables shifted to func declaration
  '  Dim wCmLkUpLabRt As String
  ' **** Manali 3.03 - variables shifted to func declaration
  
  If (mOpt <> "C") And (mOpt <> "S") Then GetLabVal = 0: Exit Function
  If (mLabQw <> "W") And (mLabQw <> "Q") Then GetLabVal = 0: Exit Function
  wLabVal = mLabRt * IIF(mLabQw = "Q", mLabQty, mGldAsWt)
  If mOpt = "C" Then GetLabVal = wLabVal: Exit Function
  
  wMax = 0: wMin = 0
  
  ' **** Bef 3.03 *****
  'wCmLkUpLabRt = goCn.GetFldVal("Select CmLkUpLabRt from CustMst where CmCtg= '" + mCmCtg + "' and CmCd= '" + mCmCd + "'")
  
  ' **** Manali 3.03 - 02/07/08
  '6.1-13 seek from custmst is removed
  'If mCmLkUpLabRt = "" Then _
  mCmLkUpLabRt = goCn.GetFldVal("Select CmLkUpLabRt from CustMst where CmCtg= '" + mCmCtg + "' and CmCd= '" + mCmCd + "'")
  ' **** Manali 3.03 - 02/07/08
  
  wCmSeekZSelfRt = goCn.RecSeek("Select 'x' from CustMst where CmCtg= '" & mCmCtg & "' and CmCd= '" + mCmCd + "' and CmSeekZSelfRt='Y'")
  
  ' Zubin 213 (mCmCurCd added)
  wMaxCommStr = " Select LrSalMax from LabRt where LrMCd= '" + mLabMCd + "' and LrSCd= '" + mLabSCd + "' and LrCmCurCd= '" + mCmCurCd + "' "
  wMinCommStr = " Select LrSalMin from LabRt where LrMCd= '" + mLabMCd + "' and LrSCd= '" + mLabSCd + "' and LrCmCurCd= '" + mCmCurCd + "' "
  
  ' ***** Manali Trading Module - 'S' added below - RrCmCtg='C' for 'ZSELF' replaced by mCmCtg (for Suppier)
  If UCase(mCmCtg) = "C" Or UCase(mCmCtg) = "S" Then
    wMax = goCn.GetFldVal("Select IsNull" + _
           "((" + wMaxCommStr + " and LrCmCtg= 'P' and LrCmCd='" + mCmLkUpLabRt + "'" + _
             " and LrFrWt<= " + CStr(mGldAsWt) + " and LrToWt>= " + CStr(mGldAsWt) + ")" + _
            ", " + _
           "IsNull((" + wMaxCommStr + " and LrCmCtg='" + mCmCtg + "' /*'C'*/ and LrCmCd='" + mCmCd + "'" + _
                   " and LrFrWt<= " + CStr(mGldAsWt) + " and LrToWt>= " + CStr(mGldAsWt) + ")" + _
                  ", " + _
  IIF(wCmSeekZSelfRt Or mOpt = "C", "IsNull((" + wMaxCommStr + " and LrCmCtg='" + mCmCtg + "' /*'C'*/ and LrCmCd='" + ctSelfCmCd + "' " + _
                          " and LrFrWt<= " + CStr(mGldAsWt) + " and LrToWt>= " + CStr(mGldAsWt) + ")" + _
                          ",0)", "0") + "))")
    wMin = goCn.GetFldVal("select IsNull" + _
           "((" + wMinCommStr + " and LrCmCtg= 'P' and LrCmCd='" + mCmLkUpLabRt + "' " + _
             " and LrFrWt<= " + CStr(mGldAsWt) + " and LrToWt>= " + CStr(mGldAsWt) + ")" + _
            ", " + _
           "IsNull((" + wMinCommStr + " and LrCmCtg='" + mCmCtg + "' /*'C'*/ and LrCmCd='" + mCmCd + "' " + _
                    " and LrFrWt<= " + CStr(mGldAsWt) + " and LrToWt>= " + CStr(mGldAsWt) + ")" + _
                    ", " + _
   IIF(wCmSeekZSelfRt Or mOpt = "C", "IsNull((" + wMinCommStr + " and LrCmCtg='" + mCmCtg + "' /*'C'*/ and LrCmCd='" + ctSelfCmCd + "' " + _
                            " and LrFrWt<= " + CStr(mGldAsWt) + " and LrToWt>= " + CStr(mGldAsWt) + ")" + _
                                                           " ,0)", "0") + "))")
  ElseIf UCase(mCmCtg) = "T" Then
    wMax = goCn.GetFldVal(wMaxCommStr + " and LrCmCtg= 'T' and LrCmCd= '" + mCmCd + "' " + _
                "and LrFrWt<= " + CStr(mGldAsWt) + " and LrToWt>= " + CStr(mGldAsWt) + " ")
    wMin = goCn.GetFldVal(wMinCommStr + " and LrCmCtg= 'C' and LrCmCd= '" + mCmCd + "' " + _
                "and LrFrWt<= " + CStr(mGldAsWt) + " and LrToWt>= " + CStr(mGldAsWt) + " ")
  End If
    
  wLabVal = IIF(wMin <> 0 And wLabVal < wMin, wMin, wLabVal)
  wLabVal = IIF(wMax <> 0 And wLabVal > wMax, wMax, wLabVal)
  GetLabVal = wLabVal
End Function

Public Function GetAlyCd(ByVal mCmCd As String, ByVal mRmCd As String, _
                         ByVal mDmCol As String) As String
  GetAlyCd = goCn.GetFldVal("Select IsNull" + _
            "((Select CaAlyCd from CustAly where CaCtg='C' and " + _
              "CaCd='" + mCmCd + "' and CaRmCd= '" + mRmCd + "' and " + _
              "CaDmCol='" + mDmCol + "' and CaAlyCd<> '')" + _
             ", " + _
             "IsNull((Select CaAlyCd from CustAly where CaCtg='C' and " + _
                     "CaCd='" + ctSelfCmCd + "' and CaRmCd= '" + mRmCd + "' and " + _
                     "CaDmCol='" + mDmCol + "' and CaAlyCd<> '' ),''))")
End Function

'Public Sub GetXtraWt1(ByVal mRrTcTyp As String, ByVal mSrcStr As String, _
'                     ByVal mCmDiaRtFA As String, ByVal mCmCSRtFA As String, _
'                     ByVal mCmCd As String, ByVal mDmCtg As String, _
'                     ByRef mRsFinalResSet As MDORowSet)
''### for testing
'''  Dim wRs As MwfLib.MDORowSet
'''  Call GetXtraWt("XDM", "Select DrSr as qSrNo, DrRmCtg as qRmCtg, DrRmSCtg as qRmSCtg, " + _
'''                 "DrRmCd as qRmCd, DrWt as qRmWt, 9.999 as qXtraWt, 0 as qRmRt into #TmpDsgRm from DsgRm " + _
'''                 "where DrTcTyp= 'DM' and DrCd= 'R 338 R' and DrSz='' AND DrRmCtg in ('D')", "A", "A", _
'''                 "A1", "RNG", wRs)
'''
''### for testing
'
'
'  '*** Select qSrNo, qRmCtg, qRmSCtg, qRmCd, qRmWt, 9.999 as qXtraWt, qRmRt into #TmpTbl from SourceTbl
'
'  Dim wTmpTbl As String, wflagXtraSCtg As Boolean, wRetVal As Boolean, wRsRmGrp As MwfLib.MDORowSet
'
'  '*** Creating Temporary Table #TmpDsgRm ***
'  wTmpTbl = "#TmpDsgRm"
'  wflagXtraSCtg = False
'  gocn.CreateTmpTable wTmpTbl, mSrcStr
'
'  '*** If Rec for Ctg-SCtg Combination is found in RtMst then Extra Wt is applied to only those Recs
'  Set wRsRmGrp = gocn.OpenRes("Select qRmCtg, qRmSCtg, sum(qRmWt) as qRmWt " + _
'                              "from " + wTmpTbl + " where qRmCtg in ('C', 'D') " + _
'                              "Group By qRmCtg, qRmSCtg Order By qRmCtg, qRmSCtg")
'  Do While Not (wRsRmGrp.EOF Or wRsRmGrp.BOF)
'    wRetVal = FnGetXtraWt(mRrTcTyp, wTmpTbl, mCmDiaRtFA, mCmCSRtFA, mCmCd, mDmCtg, _
'                          wRsRmGrp!qRmCtg, wRsRmGrp!qRmSCtg, wRsRmGrp!qRmWt)
'    If wRetVal = True Then wflagXtraSCtg = True
'    wRsRmGrp.MoveNext
'  Loop
'
'  If wflagXtraSCtg = True Then
'    Set mRsFinalResSet = gocn.OpenRes("Select * from " + wTmpTbl)
'    Exit Sub
'  End If
'
'  '*** If no Rec is found for any one Ctg-SCtg Combination,
'       'then search for RmRt rec with SCtg blank
'       'If found then Extra Wt is applied to all recs of the Ctg
'  Set wRsRmGrp = gocn.OpenRes("Select qRmCtg, qRmSCtg, sum(qRmWt) as qRmWt " + _
'                              "from " + wTmpTbl + " where qRmCtg in ('C', 'D') " + _
'                              "Group By qRmCtg Order By qRmCtg")
'  Do While Not (wRsRmGrp.EOF Or wRsRmGrp.BOF)
'    Call FnGetXtraWt(mRrTcTyp, wTmpTbl, mCmDiaRtFA, mCmCSRtFA, mCmCd, mDmCtg, _
'                     wRsRmGrp!qRmCtg, "", wRsRmGrp!qRmWt)
'    wRsRmGrp.MoveNext
'  Loop
'  Set mRsFinalResSet = gocn.OpenRes("Select * from " + wTmpTbl)
'End Sub

Public Sub GetXtraWt(ByVal mDRrTcTyp As String, ByVal mCRrTcTyp As String, _
                     ByVal mTmpTbl As String, ByVal mCmDiaRtFA As String, _
                     ByVal mCmCSRtFA As String, ByVal mCmCd As String, _
                     ByVal mDmCtg As String, ByVal mWtFld As String, _
                     ByVal mXtraWtFld As String, ByVal mRtFld As String, _
                     ByVal mDsgQty As Single)
'### for testing
''  Dim wRs As MwfLib.MDORowSet
''  Call GetXtraWt("XDM", "Select DrSr as qSrNo, DrRmCtg as qRmCtg, DrRmSCtg as qRmSCtg, " + _
''                 "DrRmCd as qRmCd, DrWt as qRmWt, 9999.999 as qXtraWt, 0 as qRmRt into #TmpDsgRm from DsgRm " + _
''                 "where DrTcTyp= 'DM' and DrCd= 'R 338 R' and DrSz='' AND DrRmCtg in ('D')", "A", "A", _
''                 "A1", "RNG", wRs)
''
'### for testing
  
  
  '*** Select qSrNo, qRmCtg, qRmSCtg, qRmCd, qRmWt, 999.999 as qXtraWt, qRmRt into #TmpTbl from SourceTbl
  
  Dim wflagXtraSCtg As Boolean, wRetVal As Boolean, wRsRmGrp As MwfLib.MDORowSet
 
  
  '************** For Dia *******************
  wflagXtraSCtg = False
  '*** If Rec for Ctg-SCtg Combination is found in RtMst then Extra Wt is applied to only those Recs
  Set wRsRmGrp = goCn.OpenRes("Select qRmCtg, qRmSCtg, sum(" + mWtFld + ") as qRmWt " + _
                              "from " + mTmpTbl + " where qRmCtg in ('D') " + _
                              "Group By qRmCtg, qRmSCtg Order By qRmCtg, qRmSCtg")
  Do While Not (wRsRmGrp.EOF Or wRsRmGrp.BOF)
    wRetVal = FnGetXtraWt(mDRrTcTyp, mTmpTbl, mCmDiaRtFA, mCmCSRtFA, mCmCd, mDmCtg, _
                          wRsRmGrp!qRmCtg, wRsRmGrp!qRmSCtg, MWLib.Div(wRsRmGrp!qRmWt, mDsgQty), _
                          mWtFld, mXtraWtFld, mDsgQty, mRtFld)
    If wRetVal = True Then wflagXtraSCtg = True
    wRsRmGrp.MoveNext
  Loop
  
  If wflagXtraSCtg = False Then
    '*** If no Rec is found for any one Ctg-SCtg Combination,
         'then search for RmRt rec with SCtg blank
         'If found then Extra Wt is applied to all recs of the Ctg
    Set wRsRmGrp = goCn.OpenRes("Select qRmCtg, '', sum(" + mWtFld + ") as qRmWt " + _
                                "from " + mTmpTbl + " where qRmCtg in ('D') " + _
                                "Group By qRmCtg Order By qRmCtg")
    Do While Not (wRsRmGrp.EOF Or wRsRmGrp.BOF)
      Call FnGetXtraWt(mDRrTcTyp, mTmpTbl, mCmDiaRtFA, mCmCSRtFA, mCmCd, mDmCtg, _
                       wRsRmGrp!qRmCtg, "", MWLib.Div(wRsRmGrp!qRmWt, mDsgQty), mWtFld, mXtraWtFld, mDsgQty, mRtFld)
      wRsRmGrp.MoveNext
    Loop
  End If

  '************** For Col Stn *******************
  wflagXtraSCtg = False
  '*** If Rec for Ctg-SCtg Combination is found in RtMst then Extra Wt is applied to only those Recs
  Set wRsRmGrp = goCn.OpenRes("Select qRmCtg, qRmSCtg, sum(" + mWtFld + ") as qRmWt " + _
                              "from " + mTmpTbl + " where qRmCtg in ('C') and " + mWtFld + "> 0 " + _
                              "Group By qRmCtg, qRmSCtg Order By qRmCtg, qRmSCtg")
  Do While Not (wRsRmGrp.EOF Or wRsRmGrp.BOF)
    wRetVal = FnGetXtraWt(mCRrTcTyp, mTmpTbl, mCmDiaRtFA, mCmCSRtFA, mCmCd, mDmCtg, _
                          wRsRmGrp!qRmCtg, wRsRmGrp!qRmSCtg, MWLib.Div(wRsRmGrp!qRmWt, mDsgQty), mWtFld, _
                          mXtraWtFld, mDsgQty, mRtFld)
    If wRetVal = True Then wflagXtraSCtg = True
    wRsRmGrp.MoveNext
  Loop
  
  If wflagXtraSCtg = False Then
    '*** If no Rec is found for any one Ctg-SCtg Combination,
         'then search for RmRt rec with SCtg blank
         'If found then Extra Wt is applied to all recs of the Ctg
    Set wRsRmGrp = goCn.OpenRes("Select qRmCtg, '', sum(" + mWtFld + ") as qRmWt " + _
                                "from " + mTmpTbl + " where qRmCtg in ('C') " + _
                                "Group By qRmCtg Order By qRmCtg")
    Do While Not (wRsRmGrp.EOF Or wRsRmGrp.BOF)
      Call FnGetXtraWt(mCRrTcTyp, mTmpTbl, mCmDiaRtFA, mCmCSRtFA, mCmCd, mDmCtg, _
                       wRsRmGrp!qRmCtg, "", MWLib.Div(wRsRmGrp!qRmWt, mDsgQty), mWtFld, mXtraWtFld, mDsgQty, mRtFld)
      wRsRmGrp.MoveNext
    Loop
  End If
End Sub
Public Function FnGetXtraWt(ByVal mRrTcTyp As String, ByVal mTmpTbl As String, _
                            ByVal mCmDiaRtFA As String, ByVal mCmCSRtFA As String, _
                            ByVal mCmCd As String, ByVal mDmCtg As String, _
                            ByVal mRmCtg As String, ByVal mRmSCtg As String, _
                            ByVal mRmWt As Single, ByVal mWtFld As String, _
                            ByVal mXtraWtFld As String, ByVal mDsgQty As Single, _
                            ByVal mRtFld As String) As Boolean
  
  Dim wRsRm As MwfLib.MDORowSet, wXtraWt As Single, wCommStr As String, wRtFA As String
  Dim wTotRecs As Integer, wMulFact As Single, wTotRmWt As Single, wRow As Integer
  Dim wMaxRt As Single, wStr As String, wRmWt As Single

  FnGetXtraWt = False
  
'  wCommStr = "select (case when RrCstRt>0 then RrCstRt else " + _
'             "RrSalRt- " + CStr(mRmWt) + " end) from RmRt where " + _
'             "RrTcTyp= '" + mRrTcTyp + "' and RrCtg='" + mRmCtg + "' and " + _
'             IIf(mRmSCtg = "", "", "RrSCtg= '" + mRmSCtg + "' and ") + " " + _
'             "RrFrLn<= " + CStr(mRmWt) + " and RrToLn>= " + CStr(mRmWt) + " "

'before ECH
'  wCommStr = "select (case when RrCstRt>0 then RrCstRt else " + _
'             "RrSalRt- " + CStr(mRmWt) + " end) from RmRt where " + _
'             "RrTcTyp= '" + mRrTcTyp + "' and RrCtg='" + mRmCtg + "' and " + _
'             "RrSCtg= '" + mRmSCtg + "' and " + _
'             "RrFrLn<= " + CStr(mRmWt) + " and RrToLn>= " + CStr(mRmWt) + " "

  wCommStr = "select (case when RrCstRt>0 then RrCstRt " + _
                          "when RrSalRt>0 then RrSalRt - " + CStr(mRmWt) + " " + _
                          "when RrXWtPer> 0 then (RrXWtPer/100) * " + CStr(mRmWt) + _
                          " else 0 end) from RmRt where " + _
             "RrTcTyp= '" + mRrTcTyp + "' and RrCtg='" + mRmCtg + "' and " + _
             "RrSCtg= '" + mRmSCtg + "' and " + _
             "RrFrLn<= " + CStr(mRmWt) + " and RrToLn>= " + CStr(mRmWt) + " " + IIF(gs_Partition = ctCurrPrtn, " and RrPrtKey='" + ctCurrPrtn + "' ", "")

  wXtraWt = goCn.GetFldVal("Select isnull" + _
         "((" + wCommStr + " and RrCmCtg= 'C' and RrCmCd = '" + mCmCd + "' and RrDmCtg='" + mDmCtg + "')" + _
          ", " + _
          "IsNull((" + wCommStr + " and RrCmCtg= 'C' and RrCmCd = '" + mCmCd + "' and RrDmCtg='')" + _
                 ", " + _
                 "IsNull((" + wCommStr + " and RrCmCtg= 'C' and RrCmCd = '" + ctSelfCmCd + "' and RrDmCtg='" + mDmCtg + "')" + _
                        ", " + _
                        "IsNull((" + wCommStr + " and RrCmCtg= 'C' and RrCmCd = '" + ctSelfCmCd + "' and RrDmCtg=''),0)" + _
                        ")" + _
                 ")" + _
          ")")
  If wXtraWt <> 0 Then
    FnGetXtraWt = True
    wRtFA = IIF(mRmCtg = "D", mCmDiaRtFA, mCmCSRtFA)

    If wRtFA = "A" Then
      wStr = ""
    ElseIf wRtFA = "F" Then
      wMaxRt = goCn.GetFldVal("Select max(" + mRtFld + ") from " + mTmpTbl + " where " + _
                              "qRmCtg='" + mRmCtg + "' and " + mWtFld + "> 0 " + _
                              IIF(mRmSCtg = "", "", "and qRmSCtg= '" + mRmSCtg + "'"))
      wStr = " and " + mRtFld + "= " + CStr(wMaxRt)
    End If
    Set wRsRm = goCn.OpenResultset("Select * from " + mTmpTbl + " where " + _
                                   "qRmCtg='" + mRmCtg + "' and " + mWtFld + "> 0 " + _
                                   IIF(mRmSCtg = "", "", "and qRmSCtg= '" + mRmSCtg + "'") + wStr)
    
    wXtraWt = wXtraWt * mDsgQty
    wTotRecs = wRsRm.RecCount
    wTotRmWt = mRmWt * mDsgQty
    If wRsRm.RecCount > 0 Then wRsRm.MoveFirst
    wRow = 0
    
    Do While Not (wRsRm.EOF Or wRsRm.BOF)
      wMulFact = MWLib.Div(wXtraWt, wTotRmWt)
      'DoEvents
      'gocn.BeginTrans (mdoCommitRec)
      wRsRm.edit
      wRow = wRow + 1
      If wRow < wTotRecs Then
        '*** If Not The Last Record ***
        wRmWt = wRsRm(mWtFld) * wMulFact
      Else
        '*** If Last Record ***
        wRmWt = wXtraWt
      End If
      wRsRm(mXtraWtFld) = Round(wRmWt, 3) ' ****** Manali 3.5.0 - 15/12/08 - Rounded wRmWt
      'wRsRm(mXtraWtFld) = wRmWt
      wRsRm.Update
      'gocn.CommitTrans mdoCommitRec

      wTotRmWt = wTotRmWt - wRsRm(mWtFld)
      wXtraWt = wXtraWt - wRmWt
      wRsRm.MoveNext
    Loop
  
  End If
  
  If FnGetXtraWt = False Then
    wCommStr = "select 'x' from RmRt where RrTcTyp= '" + mRrTcTyp + "' and " + _
               "RrCtg='" + mRmCtg + "' and RrSCtg<> '' and RrFrLn<= " + CStr(mRmWt) + " and " + _
               "RrToLn>= " + CStr(mRmWt) + " and (case when RrCstRt>0 then RrCstRt else " + _
               "RrSalRt- " + CStr(mRmWt) + " end)> 0 " + IIF(gs_Partition = ctCurrPrtn, " and RrPrtKey='" + ctCurrPrtn + "' ", "")
      
      If goCn.RecSeek(wCommStr + " and RrCmCtg= 'C' and RrCmCd = '" + mCmCd + "' and " + _
                      "RrDmCtg='" + mDmCtg + "'") Then FnGetXtraWt = True
      If FnGetXtraWt = False And goCn.RecSeek(wCommStr + " and RrCmCtg= 'C' and RrCmCd = '" + mCmCd + "' and " + _
                       "RrDmCtg=''") Then FnGetXtraWt = True
      If FnGetXtraWt = False And goCn.RecSeek(wCommStr + " and RrCmCtg= 'C' and RrCmCd = '" + ctSelfCmCd + "' " + _
                       "and RrDmCtg='" + mDmCtg + "'") Then FnGetXtraWt = True
      If FnGetXtraWt = False And goCn.RecSeek(wCommStr + " and RrCmCtg= 'C' and RrCmCd = '" + ctSelfCmCd + "' " + _
                       "and RrDmCtg=''") Then FnGetXtraWt = True
  End If
End Function
Public Function BagsExist(ByVal mOrdCoCd As String, ByVal mOrdTc As String, ByVal mOrdYy As String, _
                          ByVal mOrdChr As String, ByVal mOrdNo As Single, _
                          ByVal mOrdSr As Single) As Boolean
  BagsExist = goCn.RecSeek("Select 'x' from Bag where BCoCd= '" + mOrdCoCd + "' and " + _
                           "BOdTc = '" + mOrdTc + "' and BOdYy = '" + mOrdYy + "' and " + _
                           "BOdChr = '" + mOrdChr + "' and BOdNo = " + CStr(mOrdNo) + " and " + _
                           "BOdSr = " + CStr(mOrdSr))
End Function
Public Function GetOrdBagQty(ByVal mOrdCoCd As String, ByVal mOrdTc As String, _
                             ByVal mOrdYy As String, ByVal mOrdChr As String, _
                             ByVal mOrdNo As Single, ByVal mOrdSr As Single) As Single
' **** Zubin 211 **** '
'''''  GetOrdBagQty = gocn.GetFldVal("Select sum(BQty) from Bag where " + _
                 "BCls= 'N' and BCoCd= '" + mOrdCoCd + "' and BOdTc = '" + mOrdTc + "' and " + _
                 "BOdYy = '" + mOrdYy + "' and BOdChr = '" + mOrdChr + "' and " + _
                 "BOdNo = " + CStr(mOrdNo) + " and BOdSr = " + CStr(mOrdSr))
  GetOrdBagQty = goCn.GetFldVal("Select sum(BQty) from Bag where " + _
                 "BCls= 'N' and BCoCd= '" + mOrdCoCd + "' and BOdTc = '" + mOrdTc + "' and " + _
                 "BOdYy = '" + mOrdYy + "' and BOdChr = '" + mOrdChr + "' and " + _
                 "BOdNo = " + CStr(mOrdNo) + " and BOdSr = " + CStr(mOrdSr) + " and " + _
                 "(select PValue from Param where PTyp= 'BCHR' and PMCd= BChr and PSCd= '')= 'N'")
' **** Zubin 211 **** '
End Function
Public Function GetCmCdHlpSrt() As String
  GetCmCdHlpSrt = goCn.GetFldVal("Select hCmCdHlp from Head where HCoCd= '" + ctSelfCoCd + "' and HCd='" + ctSelfCmCd + "'")
End Function
Public Function Find(ByVal RecSet As MDORowSet, ByVal SeekFld As String, _
                    ByVal SeekVal As Variant, ByVal RetFld As String) As Variant
'*** This func is used in WIW Rep ***
  If RecSet.RecCount > 0 Then RecSet.MoveFirst
  Do While Not RecSet.EOF Or RecSet.BOF
    If RecSet(SeekFld) = SeekVal Then
      Find = RecSet(RetFld)
      Exit Function
    Else
      RecSet.MoveNext
    End If
  Loop
  Find = False
End Function

Public Sub GetParamHd(ByVal ps_Tbl As String, ByVal ps_PTyp As String, ByRef pv_HdPMCd As Variant, _
                      ByRef pv_HdPSCd As Variant, ByRef pv_HdPDesc As Variant, _
                      ByRef pv_HdPDesc225 As Variant, ByRef pv_HdPValue As Variant, _
                      ByRef pv_HdPNum As Variant, ByRef pv_HdPValue1 As Variant, _
                      ByRef pv_HdPNum1 As Variant, ByRef pv_HdPValue2 As Variant, _
                      ByRef pv_HdPValue3 As Variant, Optional ByRef pv_HdPValidYN As Variant = "", _
                      Optional pv_HdPValue4 As Variant = "")
  
  Dim wPMCdMaxLen As Single, wPSCdMaxLen As Single, wPDescMaxLen As Single
  Dim wPDesc225MaxLen As Single, wPValueMaxLen As Single, wPNumMaxLen As Single
  Dim wPValue1MaxLen As Single, wPNum1MaxLen As Single, wPValue2MaxLen As Single
  Dim wPValue3MaxLen As Single, wPValidYNMaxLen As Single, wPValue4MaxLen As Single
  Dim wSpaceLen As Single
  
  Dim ws_TcStr As String, ws_TsTbl As String, ws_TsTyp As String, ws_TsIdName As String
  Dim ws_TsLbl As String, ws_TsMaxLen As String, ws_PMCd As String, ws_PSCd As String
  Dim ws_PDesc As String, ws_PDesc225 As String, ws_PValue As String, ws_PNum As String
  Dim ws_PValue1 As String, ws_PNum1 As String, ws_PValue2 As String, ws_PValue3 As String
  Dim ws_PValidYN As String, ws_PValue4 As String
  
  If UCase(ps_Tbl) = UCase("xParam") Then
    ws_TcStr = "xTcStr": ws_TsTbl = "xTsTbl": ws_TsTyp = "xTsTyp": ws_TsIdName = "xTsIdName"
    ws_TsLbl = "xTsLbl": ws_TsMaxLen = "xTsMaxLen"
    ws_PMCd = "xPMCd": ws_PSCd = "xPSCd": ws_PDesc = "xPDesc"
    ws_PValue = "xPValue": ws_PDesc225 = "xPDesc225": ws_PNum = "xPNum"
    ws_PValue1 = "xPValue1": ws_PNum1 = "xPNum1": ws_PValue2 = "xPValue2": ws_PValue3 = "xPValue3"
  ElseIf UCase(ps_Tbl) = UCase("xvParam") Then
    ws_TcStr = "xTcStr": ws_TsTbl = "xTsTbl": ws_TsTyp = "xTsTyp": ws_TsIdName = "xTsIdName"
    ws_TsLbl = "xTsLbl": ws_TsMaxLen = "xTsMaxLen"
    ws_PMCd = "xvPMCd": ws_PSCd = "xvPSCd": ws_PDesc = "xvPDesc"
    ws_PValue = "xvPValue": ws_PDesc225 = "xvPDesc225": ws_PNum = "xvPNum"
    ws_PValue1 = "xvPValue1": ws_PNum1 = "xvPNum1": ws_PValue2 = "xvPValue2": ws_PValue3 = "xvPValue3"
  ElseIf UCase(ps_Tbl) = UCase("Param") Then
    ws_TcStr = "TcStr": ws_TsTbl = "TsTbl": ws_TsTyp = "TsTyp": ws_TsIdName = "TsIdName"
    ws_TsLbl = "TsLbl": ws_TsMaxLen = "TsMaxLen"
    ws_PMCd = "PMCd": ws_PSCd = "PSCd": ws_PDesc = "PDesc"
    ws_PValue = "PValue": ws_PDesc225 = "PDesc225": ws_PNum = "PNum": ws_PValue1 = "PValue1"
    ws_PNum1 = "PNum1": ws_PValue2 = "PValue2": ws_PValue3 = "PValue3": ws_PValidYN = "PValidYN"
  ElseIf UCase(ps_Tbl) = UCase("vParam") Then
    ws_TcStr = "TcStr": ws_TsTbl = "TsTbl": ws_TsTyp = "TsTyp": ws_TsIdName = "TsIdName"
    ws_TsLbl = "TsLbl": ws_TsMaxLen = "TsMaxLen"
    ws_PMCd = "vPMCd": ws_PSCd = "vPSCd": ws_PDesc = "vPDesc"
    ws_PValue = "vPValue": ws_PDesc225 = "vPDesc225": ws_PNum = "vPNum": ws_PValue1 = "vPValue1"
    ws_PNum1 = "vPNum1": ws_PValue2 = "vPValue2": ws_PValue3 = "vPValue3": ws_PValidYN = "vPValidYN"
    ws_PValue4 = "vPValue4"
  End If
  
  pv_HdPMCd = goCn.GetFldVal("Select " + ws_TsLbl + " From " + ws_TcStr + " where " + ws_TsTbl + "= '" + ps_Tbl + "' and " + ws_TsTyp + "= '" + ps_PTyp + "' and " + ws_TsIdName + "= '" + ws_PMCd + "'")
  pv_HdPSCd = goCn.GetFldVal("Select " + ws_TsLbl + " From " + ws_TcStr + "  where " + ws_TsTbl + "= '" + ps_Tbl + "' and " + ws_TsTyp + "= '" + ps_PTyp + "' and " + ws_TsIdName + "= '" + ws_PSCd + "'")
  pv_HdPDesc = goCn.GetFldVal("Select " + ws_TsLbl + " From " + ws_TcStr + "  where " + ws_TsTbl + "= '" + ps_Tbl + "' and " + ws_TsTyp + "= '" + ps_PTyp + "' and " + ws_TsIdName + "= '" + ws_PDesc + "'")
  pv_HdPDesc225 = goCn.GetFldVal("Select " + ws_TsLbl + " From " + ws_TcStr + "  where " + ws_TsTbl + "= '" + ps_Tbl + "' and " + ws_TsTyp + "= '" + ps_PTyp + "' and " + ws_TsIdName + "= '" + ws_PDesc225 + "'")
  pv_HdPValue = goCn.GetFldVal("Select " + ws_TsLbl + " From " + ws_TcStr + "  where " + ws_TsTbl + "= '" + ps_Tbl + "' and " + ws_TsTyp + "= '" + ps_PTyp + "' and " + ws_TsIdName + "= '" + ws_PValue + "'")
  pv_HdPNum = goCn.GetFldVal("Select " + ws_TsLbl + " From " + ws_TcStr + "  where " + ws_TsTbl + "= '" + ps_Tbl + "' and " + ws_TsTyp + "= '" + ps_PTyp + "' and " + ws_TsIdName + "= '" + ws_PNum + "'")
  pv_HdPValue1 = goCn.GetFldVal("Select " + ws_TsLbl + " From " + ws_TcStr + "  where " + ws_TsTbl + "= '" + ps_Tbl + "' and " + ws_TsTyp + "= '" + ps_PTyp + "' and " + ws_TsIdName + "= '" + ws_PValue1 + "'")
  pv_HdPNum1 = goCn.GetFldVal("Select " + ws_TsLbl + " From " + ws_TcStr + "  where " + ws_TsTbl + "= '" + ps_Tbl + "' and " + ws_TsTyp + "= '" + ps_PTyp + "' and " + ws_TsIdName + "= '" + ws_PNum1 + "'")
  pv_HdPValue2 = goCn.GetFldVal("Select " + ws_TsLbl + " From " + ws_TcStr + "  where " + ws_TsTbl + "= '" + ps_Tbl + "' and " + ws_TsTyp + "= '" + ps_PTyp + "' and " + ws_TsIdName + "= '" + ws_PValue2 + "'")
  pv_HdPValue3 = goCn.GetFldVal("Select " + ws_TsLbl + " From " + ws_TcStr + "  where " + ws_TsTbl + "= '" + ps_Tbl + "' and " + ws_TsTyp + "= '" + ps_PTyp + "' and " + ws_TsIdName + "= '" + ws_PValue3 + "'")
  pv_HdPValidYN = goCn.GetFldVal("Select " + ws_TsLbl + " From " + ws_TcStr + "  where " + ws_TsTbl + "= '" + ps_Tbl + "' and " + ws_TsTyp + "= '" + ps_PTyp + "' and " + ws_TsIdName + "= '" + ws_PValidYN + "'")
  pv_HdPValue4 = goCn.GetFldVal("Select " + ws_TsLbl + " From " + ws_TcStr + "  where " + ws_TsTbl + "= '" + ps_Tbl + "' and " + ws_TsTyp + "= '" + ps_PTyp + "' and " + ws_TsIdName + "= '" + ws_PValue4 + "'")

  wPMCdMaxLen = goCn.GetFldVal("Select " + ws_TsMaxLen + " From " + ws_TcStr + "  where " + ws_TsTbl + "= '" + ps_Tbl + "' and " + ws_TsTyp + "= '" + ps_PTyp + "' and " + ws_TsIdName + "= '" + ws_PMCd + "'")
  wPSCdMaxLen = goCn.GetFldVal("Select " + ws_TsMaxLen + " From " + ws_TcStr + "  where " + ws_TsTbl + "= '" + ps_Tbl + "' and " + ws_TsTyp + "= '" + ps_PTyp + "' and " + ws_TsIdName + "= '" + ws_PSCd + "'")
  wPDescMaxLen = goCn.GetFldVal("Select " + ws_TsMaxLen + " From " + ws_TcStr + "  where " + ws_TsTbl + "= '" + ps_Tbl + "' and " + ws_TsTyp + "= '" + ps_PTyp + "' and " + ws_TsIdName + "= '" + ws_PDesc + "'")
  wPDesc225MaxLen = goCn.GetFldVal("Select " + ws_TsMaxLen + " From " + ws_TcStr + "  where " + ws_TsTbl + "= '" + ps_Tbl + "' and " + ws_TsTyp + "= '" + ps_PTyp + "' and " + ws_TsIdName + "= '" + ws_PDesc225 + "'")
  wPValueMaxLen = goCn.GetFldVal("Select " + ws_TsMaxLen + " From " + ws_TcStr + "  where " + ws_TsTbl + "= '" + ps_Tbl + "' and " + ws_TsTyp + "= '" + ps_PTyp + "' and " + ws_TsIdName + "= '" + ws_PValue + "'")
  wPNumMaxLen = goCn.GetFldVal("Select " + ws_TsMaxLen + " From " + ws_TcStr + "  where " + ws_TsTbl + "= '" + ps_Tbl + "' and " + ws_TsTyp + "= '" + ps_PTyp + "' and " + ws_TsIdName + "= '" + ws_PNum + "'")
  wPValue1MaxLen = goCn.GetFldVal("Select " + ws_TsMaxLen + " From " + ws_TcStr + "  where " + ws_TsTbl + "= '" + ps_Tbl + "' and " + ws_TsTyp + "= '" + ps_PTyp + "' and " + ws_TsIdName + "= '" + ws_PValue1 + "'")
  wPNum1MaxLen = goCn.GetFldVal("Select " + ws_TsMaxLen + " From " + ws_TcStr + "  where " + ws_TsTbl + "= '" + ps_Tbl + "' and " + ws_TsTyp + "= '" + ps_PTyp + "' and " + ws_TsIdName + "= '" + ws_PNum1 + "'")
  wPValue2MaxLen = goCn.GetFldVal("Select " + ws_TsMaxLen + " From " + ws_TcStr + "  where " + ws_TsTbl + "= '" + ps_Tbl + "' and " + ws_TsTyp + "= '" + ps_PTyp + "' and " + ws_TsIdName + "= '" + ws_PValue2 + "'")
  wPValue3MaxLen = goCn.GetFldVal("Select " + ws_TsMaxLen + " From " + ws_TcStr + "  where " + ws_TsTbl + "= '" + ps_Tbl + "' and " + ws_TsTyp + "= '" + ps_PTyp + "' and " + ws_TsIdName + "= '" + ws_PValue3 + "'")
  wPValidYNMaxLen = goCn.GetFldVal("Select " + ws_TsMaxLen + " From " + ws_TcStr + "  where " + ws_TsTbl + "= '" + ps_Tbl + "' and " + ws_TsTyp + "= '" + ps_PTyp + "' and " + ws_TsIdName + "= '" + ws_PValidYN + "'")
  wPValue4MaxLen = goCn.GetFldVal("Select " + ws_TsMaxLen + " From " + ws_TcStr + "  where " + ws_TsTbl + "= '" + ps_Tbl + "' and " + ws_TsTyp + "= '" + ps_PTyp + "' and " + ws_TsIdName + "= '" + ws_PValue4 + "'")

  '*** To Find Out The Space To Be Appended To The Heading
      '(This Is To Set The Proper Width For The Field)
  '*** 2.5 is taken as the approximate value which stands for the space occupied by one
      'character in Upper Case
      
  wSpaceLen = (wPMCdMaxLen * 2.5) - (Len(pv_HdPMCd))
  pv_HdPMCd = IIF(pv_HdPMCd = "", "", pv_HdPMCd + IIF(wSpaceLen > 0, Space(Abs(wSpaceLen)), ""))

  wSpaceLen = (wPSCdMaxLen * 2.5) - (Len(pv_HdPSCd))
  pv_HdPSCd = IIF(pv_HdPSCd = "", "", pv_HdPSCd + IIF(wSpaceLen > 0, Space(Abs(wSpaceLen)), ""))

  wSpaceLen = (wPDescMaxLen * 2.5) - (Len(pv_HdPDesc))
  pv_HdPDesc = IIF(pv_HdPDesc = "", "", pv_HdPDesc + IIF(wSpaceLen > 0, Space(Abs(wSpaceLen)), ""))

  wSpaceLen = (wPDesc225MaxLen * 2.5) - (Len(pv_HdPDesc225))
  pv_HdPDesc225 = IIF(pv_HdPDesc225 = "", "", pv_HdPDesc225 + IIF(wSpaceLen > 0, Space(Abs(wSpaceLen)), ""))

  wSpaceLen = (wPValueMaxLen * 2.5) - (Len(pv_HdPValue))
  pv_HdPValue = IIF(pv_HdPValue = "", "", pv_HdPValue + IIF(wSpaceLen > 0, Space(Abs(wSpaceLen)), ""))

  wSpaceLen = (wPNumMaxLen * 2.5) - (Len(pv_HdPNum))
  pv_HdPNum = IIF(pv_HdPNum = "", "", pv_HdPNum + IIF(wSpaceLen > 0, Space(Abs(wSpaceLen)), ""))

  wSpaceLen = (wPValue1MaxLen * 2.5) - (Len(pv_HdPValue1))
  pv_HdPValue1 = IIF(pv_HdPValue1 = "", "", pv_HdPValue1 + IIF(wSpaceLen > 0, Space(Abs(wSpaceLen)), ""))

  wSpaceLen = (wPNum1MaxLen * 2.5) - (Len(pv_HdPNum1))
  pv_HdPNum1 = IIF(pv_HdPNum1 = "", "", pv_HdPNum1 + IIF(wSpaceLen > 0, Space(Abs(wSpaceLen)), ""))

  wSpaceLen = (wPValue2MaxLen * 2.5) - (Len(pv_HdPValue2))
  pv_HdPValue2 = IIF(pv_HdPValue2 = "", "", pv_HdPValue2 + IIF(wSpaceLen > 0, Space(Abs(wSpaceLen)), ""))

  wSpaceLen = (wPValue3MaxLen * 2.5) - (Len(pv_HdPValue3))
  pv_HdPValue3 = IIF(pv_HdPValue3 = "", "", pv_HdPValue3 + IIF(wSpaceLen > 0, Space(Abs(wSpaceLen)), ""))
  
  wSpaceLen = (wPValidYNMaxLen * 2.5) - (Len(pv_HdPValidYN))
  pv_HdPValidYN = IIF(pv_HdPValidYN = "", "", pv_HdPValidYN + IIF(wSpaceLen > 0, Space(Abs(wSpaceLen)), ""))
  
  wSpaceLen = (wPValue4MaxLen * 2.5) - (Len(pv_HdPValue4))
  pv_HdPValue4 = IIF(pv_HdPValue4 = "", "", pv_HdPValue4 + IIF(wSpaceLen > 0, Space(Abs(wSpaceLen)), ""))
End Sub
Public Function FnNumToWords(ByVal mNo As Double, ByVal mType As String) As String
  Dim wNumWord As Collection
  Dim wStrNo As String, wDecPos As Integer, wDecNo As String
  Dim wNoWrd As String, wNoDecWord As String, wZrStr As String
  Set wNumWord = New Collection
  
  Dim ws_IntPartDesc As String, ws_DecPartDesc As String  '*** (Jen 2.13)
  
  wNumWord.Add "Zero", "0"
  wNumWord.Add "One", "1"
  wNumWord.Add "Two", "2"
  wNumWord.Add "Three", "3"
  wNumWord.Add "Four", "4"
  wNumWord.Add "Five", "5"
  wNumWord.Add "Six", "6"
  wNumWord.Add "Seven", "7"
  wNumWord.Add "Eight", "8"
  wNumWord.Add "Nine", "9"
  wNumWord.Add "Ten", "10"
  wNumWord.Add "Eleven", "11"
  wNumWord.Add "Twelve", "12"
  wNumWord.Add "Thirteen", "13"
  wNumWord.Add "Fourteen", "14"
  wNumWord.Add "Fifteen", "15"
  wNumWord.Add "Sixteen", "16"
  wNumWord.Add "Seventeen", "17"
  wNumWord.Add "Eighteen", "18"
  wNumWord.Add "Nineteen", "19"
  wNumWord.Add "Twenty", "20"
  wNumWord.Add "Thirty", "30"
  wNumWord.Add "Forty", "40"
  wNumWord.Add "Fifty", "50"
  wNumWord.Add "Sixty", "60"
  wNumWord.Add "Seventy", "70"
  wNumWord.Add "Eighty", "80"
  wNumWord.Add "Ninety", "90"
  If mNo > 99999999999.99 Then
      FnNumToWords = "Value cannot be converted"
      Exit Function
  End If
  wStrNo = CStr(mNo)
  wDecPos = InStr(1, wStrNo, ".")
  If wDecPos > 0 Then
      wStrNo = Mid(CStr(mNo), 1, wDecPos - 1)
      wDecNo = Mid(CStr(mNo), wDecPos + 1, Len(CStr(mNo)) - wDecPos)
  End If
  wNoWrd = FnIntToWord(wStrNo, wNumWord)
  
  '*** (Bef 2.13)
  'wNoDecWord = FnIntToWord(wDecNo, wNumWord, Switch(mType = "USD" Or mType = "INR", 2))
  'If Val(wDecNo) <> 0 Then wNoWrd = wNoWrd + " and " + _
  '            Switch(mType = "USD", "cents", mType = "INR", "paise") + _
  '                                    wNoDecWord
  'If wNoWrd = "" Then
  '    wZrStr = FnNumWrd("0", wNumWord)
  '    wNoWrd = wZrStr + " and cents " + wZrStr
  'End If
  'wNoWrd = Switch(mType = "USD", "US Dollars ", mType = "INR", "Rs ") + _
  '         wNoWrd + " Only"
  '*** (Bef 2.13)
  
  '*** (Jen 2.13)
  ws_IntPartDesc = goCn.GetFldVal("Select PValue from Param where PTyp= 'CURNCY' and PmCd= '" + mType + "'")
  ws_DecPartDesc = goCn.GetFldVal("Select PValue1 from Param where PTyp= 'CURNCY' and PmCd= '" + mType + "'")
  wNoDecWord = FnIntToWord(wDecNo, wNumWord, Switch(Trim(mType) <> "", 2))
  If Val(wDecNo) <> 0 Then wNoWrd = wNoWrd + " and " + ws_DecPartDesc + wNoDecWord
  If wNoWrd = "" Then
      wZrStr = FnNumWrd("0", wNumWord)
      wNoWrd = wZrStr + " and cents " + wZrStr
  End If
  wNoWrd = ws_IntPartDesc + " " + wNoWrd + " Only"
  '*** (Jen 2.13)
  
  FnNumToWords = wNoWrd
End Function
Private Function FnNumWrd(ByVal mNo As String, ByVal mNumWord As Collection) As String
  Dim wLstChr As String
  If Val(mNo) >= 0 And Val(mNo) < 21 Then
      FnNumWrd = mNumWord(Trim(CStr(Val(mNo))))
  Else
      wLstChr = Mid(mNo, 2, 1)
      FnNumWrd = mNumWord(CStr(Int(Val(mNo) / 10) * 10))
      If Val(wLstChr) > 0 Then
          FnNumWrd = FnNumWrd + " " + mNumWord(wLstChr)
      End If
  End If
End Function
Private Function FnIntToWord(ByVal mNo As String, ByVal mNumWord As Collection, _
                             Optional ByVal mDecDigit As Variant) As String
  Dim wStrNo As String, wNoCr As String, wNoLac As String, wNoTh As String, wNoHnd As String
  Dim wNoTen As String, wNoWrd As String
  wStrNo = CStr(mNo)
  If Not IsMissing(mDecDigit) Then _
              wStrNo = wStrNo + String(mDecDigit - Len(wStrNo), "0")
  wStrNo = Space(11 - Len(wStrNo)) + wStrNo
  
  wNoCr = Mid(wStrNo, 1, 4)
  wNoLac = Mid(wStrNo, 5, 2)
  wNoTh = Mid(wStrNo, 7, 2)
  wNoHnd = Mid(wStrNo, 9, 1)
  wNoTen = Mid(wStrNo, 10, 2)
  
  wNoWrd = ""
  'If Val(wNoBil) <> 0 Then wNoWrd = FnNumWrd(wNoBil, mNumWord) + " Billion"
  'If Val(wNoMil) <> 0 Then wNoWrd = wNoWrd + " " + FnNumWrd(wNoMil, mNumWord) + " Million"
  If Val(wNoCr) <> 0 Then wNoWrd = wNoWrd + " " + FnNumWrd(wNoCr, mNumWord) + " Crore"
  If Val(wNoLac) <> 0 Then wNoWrd = wNoWrd + " " + FnNumWrd(wNoLac, mNumWord) + " Lac"
  If Val(wNoTh) <> 0 Then wNoWrd = wNoWrd + " " + FnNumWrd(wNoTh, mNumWord) + " Thousand"
  If Val(wNoHnd) <> 0 Then wNoWrd = wNoWrd + " " + FnNumWrd(wNoHnd, mNumWord) + " Hundred"
  If Val(wNoTen) <> 0 Then wNoWrd = wNoWrd + " " + FnNumWrd(wNoTen, mNumWord)
  FnIntToWord = wNoWrd
End Function
Public Function FnRndOff(ByVal mNo As Double, ByVal mDigits As Integer, ByVal mRndOffOpt As String) As Double
  Dim wFinalNo As Double
  mNo = Round(mNo, mDigits)
 
  Select Case UCase(mRndOffOpt)
  Case Is = "RNDUP1"
    wFinalNo = IIF(mNo <> Round(mNo), Int(mNo + 1), mNo)
  Case Is = "RNDUP5"
    wFinalNo = IIF(mNo / 5 <> mNo \ 5, Int((mNo + 5) / 5) * 5, mNo)
  Case Is = "RNDUP10"
    wFinalNo = IIF(mNo / 10 <> mNo \ 10, Int((mNo + 10) / 10) * 10, mNo)
  Case Is = "RNDUP50"
    wFinalNo = IIF(mNo / 50 <> mNo \ 50, Int((mNo + 50) / 50) * 50, mNo)
  Case Is = "RNDUP100"
    wFinalNo = IIF(mNo / 100 <> mNo \ 100, Int((mNo + 100) / 100) * 100, mNo)
  Case Is = "RNDDN1"
    wFinalNo = IIF(mNo <> Round(mNo), Int(mNo), mNo)
  Case Is = "RNDDN5"
    wFinalNo = IIF(mNo / 5 <> mNo \ 5, Int(mNo / 5) * 5, mNo)
  Case Is = "RNDDN10"
    wFinalNo = IIF(mNo / 10 <> mNo \ 10, Int(mNo / 10) * 10, mNo)
  Case Is = "RNDSMP1"
    wFinalNo = IIF(mNo <> Round(mNo), Round(mNo), mNo)
  Case Is = "RNDSMP5"
    wFinalNo = IIF(mNo / 5 <> mNo \ 5, Round(mNo / 5) * 5, mNo)
  Case Is = "RNDSMP10"
    wFinalNo = IIF(mNo / 10 <> mNo \ 10, Round(mNo / 10) * 10, mNo)
  ' ***** Manali 3.6.0 -  New Round Off Options added
  Case Is = "RNDUP.5"
    wFinalNo = IIF(mNo / 0.5 <> Round(mNo / 0.5), Int((mNo + 0.5) / 0.5) * 0.5, mNo)
  Case Is = "RNDUP.05"
    wFinalNo = IIF(mNo / 0.05 <> Round(mNo / 0.05), Int((mNo + 0.05) / 0.05) * 0.05, mNo)
  Case Is = "RNDDN.5"
    wFinalNo = IIF(mNo / 0.5 <> Round(mNo / 0.5), Int(mNo / 0.5) * 0.5, mNo)
  Case Is = "RNDDN.05"
    wFinalNo = IIF(mNo / 0.05 <> Round(mNo / 0.05), Int(mNo / 0.05) * 0.05, mNo)
  Case Is = "RNDSM.5"
    wFinalNo = IIF(mNo / 0.5 <> Round(mNo / 0.5), Round(mNo / 0.5) * 0.5, mNo)
  Case Is = "RNDSM.05"
    wFinalNo = IIF(mNo / 0.05 <> Round(mNo / 0.05), Round(mNo / 0.05) * 0.05, mNo)
  ' ***** Manali 3.6.0 -  New Round Off Options added
  Case Is = "NORND"
    wFinalNo = mNo
  End Select
  
  FnRndOff = wFinalNo
End Function
Public Function SetProp(ByVal ActiveFrm As Form, ByVal IdName As String, ByVal Opt As En_opt, Optional Cancel As Boolean, Optional ErrMsg As String) As Boolean
  Dim wctl As MWCTL_MED, ADC As MWCTL_ADC
  Set wctl = ActiveFrm.ADC(IdName)
  Set ADC = ActiveFrm.ADC
  Dim wRmCtg  As String, wRmSCtg  As String, wRmZ As String
  With wctl
    Select Case UCase(IdName)
    Case Is = UCase("UoRmCtgFr"), UCase("UoRmCtgTo")
        Select Case Opt
        Case Is = Initialise
          .DataType = fdtConvertToUpper
          .Alignment = tLeftAlign
          .Mask = ""
          .MaxLength = 2
        Case Is = When
          HlpList.PMCd "RMCTG"
        End Select
    Case Is = UCase("UoRmCtgSel")
        Select Case Opt
        Case Is = Initialise
          .DataType = fdtConvertToUpper
          .Alignment = tLeftAlign
          .Mask = ""
          .MaxLength = 255
        Case Is = When
          HlpList.MultiSelect = True
          HlpList.PMCd "RMCTG"
        End Select
    Case Is = UCase("UoRmSCtgFr"), UCase("UoRmSCtgTo")
        Select Case Opt
        Case Is = Initialise
          .DataType = fdtConvertToUpper
          .Alignment = tLeftAlign
          .Mask = ""
          .MaxLength = 5
        Case Is = When
          If gb_RmCtgFor = True Then
            HlpList.PSCd "RMSCTG", ADC("UoRmCtgFr")
          Else
            If UCase(IdName) = "UORMSCTGFR" Then HlpList.PSCd "RMSCTG", ADC("UoRmCtgFr")
            If UCase(IdName) = "UORMSCTGTO" Then HlpList.PSCd "RMSCTG", ADC("UoRmCtgTo")
          End If
        End Select
    Case Is = UCase("UoRmSCtgSel")
        Select Case Opt
        Case Is = Initialise
          .DataType = fdtConvertToUpper
          .Alignment = tLeftAlign
          .Mask = ""
          .MaxLength = 255
        Case Is = When
          HlpList.MultiSelect = True
          HlpList.PSCd "RMSCTG", ADC("UoRmCtgFr")
        End Select
    Case Is = UCase("UoRmCdFr"), UCase("UoRmCdTo")
        Select Case Opt
        Case Is = Initialise
          .DataType = fdtConvertToUpper
          .Alignment = tLeftAlign
          .Mask = ""
          .MaxLength = 16
        Case Is = When
          If gb_RmCtgFor = True Then
            wRmCtg = ADC("UoRmCtgFr")
            wRmSCtg = IIF(UCase(IdName) = UCase("UoRmCdFr"), ADC("UoRmSCtgFr"), ADC("UoRmSCtgTo"))
          Else
            If UCase(IdName) = UCase("UoRmCdFr") Then wRmCtg = ADC("UoRmCtgFr"): wRmSCtg = ADC("UoRmSCtgFr")
            If UCase(IdName) = UCase("UoRmCdTo") Then wRmCtg = ADC("UoRmCtgTo"): wRmSCtg = ADC("UoRmSCtgTo")
          End If
          If gb_RmZ = False Then wRmZ = "N"
          If gb_RmZ = True Then wRmZ = "Y"
          If Trim(wRmCtg) <> "" Then
            If Trim(wRmSCtg) <> "" Then
              HlpList.RmCd wRmZ, "'" + wRmCtg + "'", wRmSCtg
            Else
              HlpList.RmCd wRmZ, "'" + wRmCtg + "'"
            End If
          Else
            HlpList.RmCd wRmZ
          End If
        End Select
    Case Is = UCase("UoRmCdSel")
        Select Case Opt
        Case Is = Initialise
          .DataType = fdtConvertToUpper
          .Alignment = tLeftAlign
          .Mask = ""
          .MaxLength = 255
        Case Is = When
          HlpList.MultiSelect = True
          wRmCtg = ADC("UoRmCtgFr")
          wRmSCtg = ADC("UoRmSCtgFr")
          If gb_RmZ = False Then wRmZ = "N"
          If gb_RmZ = True Then wRmZ = "Y"
          If Trim(wRmCtg) <> "" Then
            If Trim(wRmSCtg) <> "" Then
              HlpList.RmCd wRmZ, "'" + wRmCtg + "'", wRmSCtg
            Else
              HlpList.RmCd wRmZ, "'" + wRmCtg + "'"
            End If
          Else
            HlpList.RmCd wRmZ
          End If
        End Select
    Case Is = UCase("UoRmCd1Fr"), UCase("UoRmCd1To")
        Select Case Opt
        Case Is = Initialise
          .DataType = fdtConvertToUpper
          .Alignment = tLeftAlign
          .Mask = ""
          .MaxLength = 16
        Case Is = When
          If gb_RmCtgFor = True Then
            wRmCtg = ADC("UoRmCtgFr")
          Else
            If UCase(IdName) = UCase("UoRmCd1Fr") Then wRmCtg = ADC("UoRmCtgFr")
            If UCase(IdName) = UCase("UoRmCd1To") Then wRmCtg = ADC("UoRmCtgTo")
          End If
          If gb_RmZ1 = False Then wRmZ = "N"
          If gb_RmZ1 = True Then wRmZ = "Y"
          If Trim(wRmCtg) <> "" Then
            HlpList.RmCd wRmZ, "'" + wRmCtg + "'"
          Else
            HlpList.RmCd wRmZ
          End If
        End Select
    Case Is = UCase("UoRmCd1Sel")
        Select Case Opt
        Case Is = Initialise
          .DataType = fdtConvertToUpper
          .Alignment = tLeftAlign
          .Mask = ""
          .MaxLength = 255
        Case Is = When
          HlpList.MultiSelect = True
          wRmCtg = ADC("UoRmCtgFr")
          If gb_RmZ1 = False Then wRmZ = "N"
          If gb_RmZ1 = True Then wRmZ = "Y"
          If Trim(wRmCtg) <> "" Then
            HlpList.RmCd wRmZ, "'" + wRmCtg + "'"
          Else
            HlpList.RmCd wRmZ
          End If
        End Select
    Case Is = UCase("UoBaseCdFr"), UCase("UoBaseCdTo")
        Select Case Opt
        Case Is = Initialise
          .DataType = fdtConvertToUpper
          .Alignment = tLeftAlign
          .Mask = ""
          .MaxLength = 16
        Case Is = When
          Call HlpList.RmCd("N", "'G','P'")
        End Select
    Case Is = UCase("UoBaseCdSel")
        Select Case Opt
        Case Is = Initialise
          .DataType = fdtConvertToUpper
          .Alignment = tLeftAlign
          .Mask = ""
          .MaxLength = 255
        Case Is = When
          HlpList.MultiSelect = True
          Call HlpList.RmCd("N", "'G','P'")
        End Select
     Case Is = UCase("UoFgSubLocFr"), UCase("UoFgSubLocTo")
        Select Case Opt
        Case Is = Initialise
          .DataType = fdtConvertToUpper
          .Alignment = tLeftAlign
          .Mask = ""
          .MaxLength = 8
        Case Is = When
            '****** Sachin 2.14.0 - [Multi-Company] ******
            If gb_CoCdFor = True Then
              HlpList.vPMCd ADC("UoCoCdFr"), "FGSUBLOC"
            Else
              If UCase(IdName) = UCase("UoFgSubLocFr") Then HlpList.vPMCd ADC("UoCoCdFr"), "FGSUBLOC"
              If UCase(IdName) = UCase("UoFgSubLocTo") Then HlpList.vPMCd ADC("UoCoCdTo"), "FGSUBLOC"
            End If
            '****** Sachin 2.14.0 - [Multi-Company] ******
        End Select
    Case Is = UCase("UoFgSubLocSel")
        Select Case Opt
        Case Is = Initialise
          .DataType = fdtConvertToUpper
          .Alignment = tLeftAlign
          .Mask = ""
          .MaxLength = 255
        Case Is = When
          HlpList.MultiSelect = True
          HlpList.vPMCd ADC("UoCoCdFr"), "FGSUBLOC"
        End Select
    Case Is = UCase("UoRecFgSubLocFr"), UCase("UoRecFgSubLocTo")
        Select Case Opt
        Case Is = Initialise
          .DataType = fdtConvertToUpper
          .Alignment = tLeftAlign
          .Mask = ""
          .MaxLength = 8
        Case Is = When
          HlpList.vPMCd ADC("UoCoCdFr"), "FGSUBLOC"
        End Select
    Case Is = UCase("UoRecFgSubLocSel")
        Select Case Opt
        Case Is = Initialise
          .DataType = fdtConvertToUpper
          .Alignment = tLeftAlign
          .Mask = ""
          .MaxLength = 255
        Case Is = When
          HlpList.MultiSelect = True
          HlpList.vPMCd ADC("UoCoCdFr"), "FGSUBLOC"
        End Select
    '***************** urmi Base Cd 2.04 ***************************
        
    Case Is = UCase("UoLocTypFr"), UCase("UoLocTypTo")
        Select Case Opt
        Case Is = Initialise
          .DataType = fdtConvertToUpper
          .Alignment = tLeftAlign
          .Mask = ""
          .MaxLength = 2
        Case Is = When
          HlpList.PMCd "LOCTYP"
        End Select
' ###################################  Manoj #### Ver: 2.0.6  ###################################
' Date: 08/12/2003
' Change: Settings Properties of UoLocTypSel
    Case Is = UCase("UoLocTypSel")
        Select Case Opt
        Case Is = Initialise
          .DataType = fdtConvertToUpper
          .Alignment = tLeftAlign
          .Mask = ""
          .MaxLength = 255
        Case Is = When
          HlpList.MultiSelect = True
          HlpList.PMCd "LOCTYP"
        End Select
' ###########################################  Manoj  ###########################################
    '****** Sachin 2.14.0 - [Multi-Company] ******
    Case Is = UCase("UoPrntCdFr"), UCase("UoPrntCdTo")
        Select Case Opt
        Case Is = Initialise
          .DataType = fdtConvertToUpper
          .Alignment = tLeftAlign
          .Mask = ""
          .MaxLength = 8
        Case Is = When
          If Trim(gs_LocTyp) <> "" Then
              HlpList.Loc ctSelfCoCd, gs_LocTyp, , "'Y'"
          Else
              HlpList.Loc ctSelfCoCd, , , "'Y'"
          End If
        End Select
    Case Is = UCase("UoPrntCdSel")
        Select Case Opt
        Case Is = Initialise
          .DataType = fdtConvertToUpper
          .Alignment = tLeftAlign
          .Mask = ""
          .MaxLength = 255
        Case Is = When
          HlpList.MultiSelect = True
          If Trim(gs_LocTyp) <> "" Then
            HlpList.Loc ctSelfCoCd, gs_LocTyp, , "'Y'"
          Else
            HlpList.Loc ctSelfCoCd, , , "'Y'"
          End If
        End Select
    '****** Sachin 2.14.0 - [Multi-Company] ******
    Case Is = UCase("UoRecLocFr"), UCase("UoRecLocTo")
        Select Case Opt
        Case Is = Initialise
          .DataType = fdtConvertToUpper
          .Alignment = tLeftAlign
          .Mask = ""
          .MaxLength = 8
        Case Is = When
          'HlpList.Loc adc("UoCoCdFr"), "'P'"
          If Trim(gs_LocTyp) <> "" Then
            HlpList.Loc ADC("UoCoCdFr"), gs_LocTyp
          Else
            HlpList.Loc ADC("UoCoCdFr")
          End If
        End Select
    Case Is = UCase("UoRecLocSel")
        Select Case Opt
        Case Is = Initialise
          .DataType = fdtConvertToUpper
          .Alignment = tLeftAlign
          .Mask = ""
          .MaxLength = 255
        Case Is = When
          HlpList.MultiSelect = True
          'HlpList.Loc adc("UoCoCdFr"), "'P'"          '###
          If Trim(gs_LocTyp) <> "" Then
            HlpList.Loc ADC("UoCoCdFr"), gs_LocTyp
          Else
            HlpList.Loc ADC("UoCoCdFr")
          End If
        End Select
    Case Is = UCase("UoDtFr"), UCase("UoDtTo")
        Select Case Opt
        Case Is = Initialise
          .DataType = fdtDateType
          .Alignment = tLeftAlign
          .Mask = ""
          .MaxLength = 8
        Case Is = When
          '***
        End Select
    Case Is = UCase("UoDtSel")
        Select Case Opt
        Case Is = Initialise
          .DataType = fdtConvertToUpper
          .Alignment = tLeftAlign
          .Mask = ""
          .MaxLength = 255
        Case Is = When
          '***
        End Select
    Case Is = UCase("UoBYyFr"), UCase("UoBYyTo")
        Select Case Opt
        Case Is = Initialise
          .DataType = fdtConvertToUpper
          .Alignment = tLeftAlign
          .Mask = ""
          .MaxLength = 2
        Case Is = When
          HlpList.PMCd "YY"
        End Select
    Case Is = UCase("UoBChrFr"), UCase("UoBChrTo"), UCase("UoBChr2Fr"), UCase("UoBChr2To")
        Select Case Opt
        Case Is = Initialise
          .DataType = fdtConvertToUpper
          .Alignment = tLeftAlign
          .Mask = ""
          .MaxLength = 3
        Case Is = When
          HlpList.PMCd "BCHR"
        End Select
    Case Is = UCase("UoBNoFr"), UCase("UoBNoTo")
        Select Case Opt
        Case Is = Initialise
          .DataType = fdtNumber
          .Alignment = tRightAlign
          .Mask = "######0"
          .MaxLength = 7
        Case Is = When
          Dim whOthLocBag As String
          whOthLocBag = goCn.GetFldVal("Select hOthLocBag from Head where HCoCd= '" + ADC("UoCoCdFr") + "' and HCd='" + ctSelfCmCd + "'")
          If UCase(IdName) = UCase("UoBNoFr") Then HlpList.BNo ADC("UoCoCdFr"), ADC("UoBYyFr"), ADC("UoBChrFr"), whOthLocBag
          If UCase(IdName) = UCase("UoBNoTo") Then HlpList.BNo ADC("UoCoCdFr"), ADC("UoBYyTo"), ADC("UoBChrTo"), whOthLocBag
        End Select
    
  '/* URMI 2.06 BagDelayDays */
    '578-29 mask, max length changed from 3 to 4
    Case Is = UCase("UoBDlDays")
        Select Case Opt
        Case Is = Initialise
          .DataType = fdtNumber
          .Alignment = tRightAlign
          .Mask = "###0"
          .MaxLength = 4
        End Select
    '578-29 UoBDlDaysTo added
    Case Is = UCase("UoBDlDaysTo")
        Select Case Opt
        Case Is = Initialise
          .DataType = fdtNumber
          .Alignment = tRightAlign
          .Mask = "###0"
          .MaxLength = 4
        End Select
        
    Case Is = UCase("UoStkRtFr"), UCase("UoStkRtTo")
        Select Case Opt
        Case Is = Initialise
          .DataType = fdtFloat
          .Alignment = tRightAlign
          .Mask = "######0.000"
          .MaxLength = 11
        End Select
    Case Is = UCase("UoDmCdFr"), UCase("UoDmCdTo")
        Select Case Opt
        Case Is = Initialise
          .DataType = fdtConvertToUpper
          .Alignment = tLeftAlign
          .Mask = ""
          .MaxLength = 15
        Case Is = When
          HlpList.DmCd gs_DmTcTyp
        End Select
    Case Is = UCase("UoDmCdSel")
        Select Case Opt
        Case Is = Initialise
          .DataType = fdtConvertToUpper
          .Alignment = tLeftAlign
          .Mask = ""
          .MaxLength = 255
        Case Is = When
          HlpList.MultiSelect = True
          HlpList.DmCd gs_DmTcTyp
        End Select
    Case Is = UCase("UoDmSzFr"), UCase("UoDmSzTo")
        Select Case Opt
        Case Is = Initialise
          .DataType = fdtConvertToUpper
          .Alignment = tLeftAlign
          .Mask = ""
          .MaxLength = 5
        Case Is = When
          If UCase(IdName) = UCase("UoDmSzFr") Then HlpList.DmSz gs_DmTcTyp, ADC("UoDmCdFr")
          If UCase(IdName) = UCase("UoDmSzTo") Then HlpList.DmSz gs_DmTcTyp, ADC("UoDmCdTo")
        End Select
    Case Is = UCase("UoOdSfxFr"), UCase("UoOdSfxTo")
        Select Case Opt
        Case Is = Initialise
          .DataType = fdtConvertToUpper
          .Alignment = tLeftAlign
          .Mask = ""
          .MaxLength = 15 '6.1-83
        Case Is = When
          ' ***** Manali Trading Module - CmCtg passed
          If UCase(IdName) = UCase("UoOdSfxFr") Then HlpList.OdSfx ADC("UoCoCdFr"), ADC("UoDmCdFr"), gs_CmCtg
          'adc("UoCmCdFr")
          If UCase(IdName) = UCase("UoOdSfxTo") Then HlpList.OdSfx ADC("UoCoCdFr"), ADC("UoDmCdTo"), gs_CmCtg
          'adc("UoCmCdTo"),
        End Select
    Case Is = UCase("UoTcTypFr")
        Select Case Opt
        Case Is = Initialise
          .DataType = fdtConvertToUpper
          .Alignment = tLeftAlign
          .Mask = ""
          .MaxLength = 3
        Case Is = When
          HlpList.PMCd "TCTYP", , gs_EntTyp   ' **** Manali Trading Module
          ' Bef Trading Module
          'HlpList.PMCd "TCTYP", , "OM"
        End Select
    Case Is = UCase("UoOmTcFr"), UCase("UoOmTcTo")
        Select Case Opt
        Case Is = Initialise
          .DataType = fdtConvertToUpper
          .Alignment = tLeftAlign
          .Mask = ""
          .MaxLength = 3
        Case Is = When
          '6.1 menu code added
          If UCase(ActiveFrm.ADC.MenuCd) = UCase("RepRmReqStat") Then
            HlpList.PMCd "TC", "'" + gs_OmTcTyp + "','JT'", , " PMCd =(case when PValue='JT' then 'JSO' else PMCd  end) "
          Else
            HlpList.PMCd "TC", "'" + gs_OmTcTyp + "'"
          End If
        End Select
    Case Is = UCase("UoOmYyFr"), UCase("UoOmYyTo")
        Select Case Opt
        Case Is = Initialise
          .DataType = fdtConvertToUpper
          .Alignment = tLeftAlign
          .Mask = ""
          .MaxLength = 2
        Case Is = When
          HlpList.PMCd "YY"
        End Select
    Case Is = UCase("UoOmChrFr"), UCase("UoOmChrTo")
       Select Case Opt
        Case Is = Initialise
          .DataType = fdtConvertToUpper
          .Alignment = tLeftAlign
          .Mask = ""
          .MaxLength = 3
        Case Is = When
          '6.1B menu code section added
          If UCase(ActiveFrm.ADC.MenuCd) = UCase("RepWipBPcs") Or UCase(ActiveFrm.ADC.MenuCd) = UCase("RepPrd") Then
            If UCase(IdName) = UCase("UoOmChrFr") Then HlpList.vPSCd ADC("UoCoCdFr"), "CHR", ADC("UoOmTcFr"), "Y"
            If gb_CoCdFor = False Then
              If UCase(IdName) = UCase("UoOmChrTo") Then HlpList.vPSCd ADC("UoCoCdTo"), "CHR", ADC("UoOmTcTo"), "Y"
            Else
              If UCase(IdName) = UCase("UoOmChrTo") Then HlpList.vPSCd ADC("UoCoCdFr"), "CHR", ADC("UoOmTcTo"), "Y"
            End If
          Else
            If UCase(IdName) = UCase("UoOmChrFr") Then HlpList.vPSCd ADC("UoCoCdFr"), "CHR", ADC("UoOmTcFr")
            If gb_CoCdFor = False Then
              If UCase(IdName) = UCase("UoOmChrTo") Then HlpList.vPSCd ADC("UoCoCdTo"), "CHR", ADC("UoOmTcTo")
            Else
              If UCase(IdName) = UCase("UoOmChrTo") Then HlpList.vPSCd ADC("UoCoCdFr"), "CHR", ADC("UoOmTcTo")
            End If
          End If
        End Select
    Case Is = UCase("UoOmNoFr"), UCase("UoOmNoTo")
        Select Case Opt
        Case Is = Initialise
          .DataType = fdtNumber
          .Alignment = tRightAlign
          .Mask = "######0"
          .MaxLength = 7
        Case Is = When
          If UCase(IdName) = UCase("UoOmNoFr") Then HlpList.OmNo ADC("UoCoCdFr"), ADC("UoOmTcFr"), ADC("UoOmYyFr"), ADC("UoOmChrFr")
          If gb_CoCdFor = False Then
            If UCase(IdName) = UCase("UoOmNoTo") Then HlpList.OmNo ADC("UoCoCdTo"), ADC("UoOmTcTo"), ADC("UoOmYyTo"), ADC("UoOmChrTo")
          Else
            If UCase(IdName) = UCase("UoOmNoTo") Then HlpList.OmNo ADC("UoCoCdFr"), ADC("UoOmTcTo"), ADC("UoOmYyTo"), ADC("UoOmChrTo")
          End If
        End Select
    Case Is = UCase("UoOdSrFr"), UCase("UoOdSrTo")
        Select Case Opt
        Case Is = Initialise
          .DataType = fdtNumber
          .Alignment = tRightAlign
          .Mask = "####0"
          .MaxLength = 5
        Case Is = When
          If UCase(IdName) = UCase("UoOdSrFr") Then HlpList.OdSr ADC("UoCoCdFr"), ADC("UoOmTcFr"), ADC("UoOmYyFr"), ADC("UoOmChrFr"), ADC("UoOmNoFr")
          If gb_CoCdFor = False Then
            If UCase(IdName) = UCase("UoOdSrTo") Then HlpList.OdSr ADC("UoCoCdTo"), ADC("UoOmTcTo"), ADC("UoOmYyTo"), ADC("UoOmChrTo"), ADC("UoOmNoTo")
          Else
            If UCase(IdName) = UCase("UoOdSrTo") Then HlpList.OdSr ADC("UoCoCdFr"), ADC("UoOmTcTo"), ADC("UoOmYyTo"), ADC("UoOmChrTo"), ADC("UoOmNoTo")
          End If
        End Select
    Case Is = UCase("UoTTcFr"), UCase("UoTTcTo")
        Select Case Opt
        Case Is = Initialise
          .DataType = fdtConvertToUpper
          .Alignment = tLeftAlign
          .Mask = ""
          .MaxLength = 3
        Case Is = When
          HlpList.PMCd "TC", "'" + gs_TxnTcTyp + "'"
        End Select
    Case Is = UCase("UoTYyFr"), UCase("UoTYyTo")
        Select Case Opt
        Case Is = Initialise
          .DataType = fdtConvertToUpper
          .Alignment = tLeftAlign
          .Mask = ""
          .MaxLength = 2
        Case Is = When
          HlpList.PMCd "YY"
        End Select
    Case Is = UCase("UoTChrFr"), UCase("UoTChrTo")
        Select Case Opt
        Case Is = Initialise
          .DataType = fdtConvertToUpper
          .Alignment = tLeftAlign
          .Mask = ""
          .MaxLength = 3
        Case Is = When          '6.1B menu code section added
          If UCase(ActiveFrm.ADC.MenuCd) = UCase("RepPrd") Then
            If UCase(IdName) = UCase("UoTChrFr") Then HlpList.vPSCd ADC("UoCoCdFr"), "CHR", ADC("UoTTcFr"), "Y"
            If UCase(IdName) = UCase("UoTChrTo") Then HlpList.vPSCd ADC("UoCoCdFr"), "CHR", ADC("UoTTcTo"), "Y"
          Else
            If UCase(IdName) = UCase("UoTChrFr") Then HlpList.vPSCd ADC("UoCoCdFr"), "CHR", ADC("UoTTcFr")
            If UCase(IdName) = UCase("UoTChrTo") Then HlpList.vPSCd ADC("UoCoCdFr"), "CHR", ADC("UoTTcTo")
          End If
        End Select
    Case Is = UCase("UoTNoFr"), UCase("UoTNoTo")
        Select Case Opt
        Case Is = Initialise
          .DataType = fdtNumber
          .Alignment = tRightAlign
          .Mask = "######0"
          .MaxLength = 7
        Case Is = When
          If UCase(IdName) = UCase("UoTNoFr") Then HlpList.TNo ADC("UoCoCdFr"), ADC("UoTTcFr"), ADC("UoTYyFr"), ADC("UoTChrFr")
          If UCase(IdName) = UCase("UoTNoTo") Then HlpList.TNo ADC("UoCoCdFr"), ADC("UoTTcTo"), ADC("UoTYyTo"), ADC("UoTChrTo")
        End Select
    Case Is = UCase("UoTdSrFr"), UCase("UoTdSrTo")
        Select Case Opt
        Case Is = Initialise
          .DataType = fdtNumber
          .Alignment = tRightAlign
          .Mask = "###0"
          .MaxLength = 4
        Case Is = When
          If UCase(IdName) = UCase("UoTdSrFr") Then HlpList.TdSr ADC("UoCoCdFr"), ADC("UoTTcFr"), ADC("UoTYyFr"), ADC("UoTChrFr"), ADC("UoTNoFr")
          If UCase(IdName) = UCase("UoTdSrTo") Then HlpList.TdSr ADC("UoCoCdFr"), ADC("UoTTcTo"), ADC("UoTYyTo"), ADC("UoTChrTo"), ADC("UoTNoTo")
        End Select
    Case Is = UCase("UoDmCtgFr"), UCase("UoDmCtgTo")
        Select Case Opt
        Case Is = Initialise
          .DataType = fdtConvertToUpper
          .Alignment = tLeftAlign
          .Mask = ""
          .MaxLength = 5
        Case Is = When
          HlpList.PMCd "DMCTG"
        End Select
    Case Is = UCase("UoDmCtgSel")
        Select Case Opt
        Case Is = Initialise
          .DataType = fdtConvertToUpper
          .Alignment = tLeftAlign
          .Mask = ""
          .MaxLength = 255
        Case Is = When
          HlpList.MultiSelect = True
          HlpList.PMCd "DMCTG"
        End Select
    Case Is = UCase("UoDmPrdCtgFr"), UCase("UoDmPrdCtgTo")
        Select Case Opt
        Case Is = Initialise
          .DataType = fdtConvertToUpper
          .Alignment = tLeftAlign
          .Mask = ""
          .MaxLength = 5
        Case Is = When
          HlpList.PMCd "PRDCTG"
        End Select
    Case Is = UCase("UoDmPrdCtgSel")
        Select Case Opt
        Case Is = Initialise
          .DataType = fdtConvertToUpper
          .Alignment = tLeftAlign
          .Mask = ""
          .MaxLength = 255
        Case Is = When
          HlpList.MultiSelect = True
          HlpList.PMCd "PRDCTG"
        End Select
    Case Is = UCase("UoDmSalCtgFr"), UCase("UoDmSalCtgTo")
        Select Case Opt
        Case Is = Initialise
          .DataType = fdtConvertToUpper
          .Alignment = tLeftAlign
          .Mask = ""
          .MaxLength = 5
        Case Is = When
        If IdName = UCase("UoDmSalCtgFr") Then
          HlpList.PSCd "SALCTG", ADC("UODMCTGFR") 'Geeta***Crm***212
        ElseIf IdName = UCase("UoDmSalCtgTo") Then
          HlpList.PSCd "SALCTG", ADC("UODMCTGTo") 'Geeta***Crm***212
        End If
        End Select
    Case Is = UCase("UoDmSalCtgSel")
        Select Case Opt
        Case Is = Initialise
          .DataType = fdtConvertToUpper
          .Alignment = tLeftAlign
          .Mask = ""
          .MaxLength = 255
        Case Is = When
          HlpList.MultiSelect = True
          HlpList.PSCd "SALCTG", ADC("UODMCTGFR")  'Geeta****Crm***212
        End Select
     '*******Geeta***********Emr206 ********10/04/04**********
     Case Is = UCase("UoTrYmdFr"), UCase("UoTrYmdTo")
        Select Case Opt
        Case Is = Initialise
          .DataType = fdtConvertToUpper
          .Alignment = tLeftAlign
          .Mask = ""
          .MaxLength = 6
        Case Is = When
          HlpList.TrYmd ADC("UoCoCdFr")
        End Select
    Case Is = UCase("UoTrNoFr"), UCase("UoTrNoTo")
        Select Case Opt
        Case Is = Initialise
          .DataType = fdtNumber
          .Alignment = tRightAlign
          .Mask = "######0"
          .MaxLength = 7
        Case Is = When
          If UCase(IdName) = UCase("UoTrNoFr") Then HlpList.TrNo ADC("UoCoCdFr"), ADC("UoTrYmdFr"), ADC("UoPMCdFr")
          If UCase(IdName) = UCase("UoTrNoTo") Then HlpList.TrNo ADC("UoCoCdFr"), ADC("UoTrYmdTo"), ADC("UoPMCdTo")
        End Select
    Case Is = UCase("UoDmPrdSeqFr"), UCase("UoDmPrdSeqTo")
      Select Case Opt
        Case Is = Initialise
          .DataType = fdtConvertToUpper
          .Alignment = tLeftAlign
          .Mask = ""
        .MaxLength = 8
        Case Is = When
          HlpList.vPMCd "", "PRDSEQ"
        End Select
     Case Is = UCase("UoDmPrdSeqSel")
      Select Case Opt
        Case Is = Initialise
          .DataType = fdtConvertToUpper
          .Alignment = tLeftAlign
          .Mask = ""
          .MaxLength = 255
      Case Is = When
         HlpList.MultiSelect = True
         HlpList.vPMCd "", "PRDSEQ"
      End Select
      
Case Is = UCase("UoMnuCdFr"), UCase("UoMnuCdTo")
      Select Case Opt
        Case Is = Initialise
          .DataType = fdtConvertToUpper
          .Alignment = tLeftAlign
          .Mask = ""
        .MaxLength = 20
        Case Is = When
          HlpList.MenuCode
        End Select
     Case Is = UCase("UoMnuCdSel")
      Select Case Opt
        Case Is = Initialise
          .DataType = fdtConvertToUpper
          .Alignment = tLeftAlign
          .Mask = ""
          .MaxLength = 255
      Case Is = When
         HlpList.MultiSelect = True
         HlpList.MenuCode
      End Select
  '*************Geeta****************
    Case Is = UCase("UoKtFr"), UCase("UoKtTo")
        Select Case Opt
        Case Is = Initialise
          .DataType = fdtConvertToUpper
          .Alignment = tLeftAlign
          .Mask = ""
          .MaxLength = 5
        Case Is = When
          HlpList.PMCd "KT"
        End Select
    Case Is = UCase("UoKtSel")
        Select Case Opt
        Case Is = Initialise
          .DataType = fdtConvertToUpper
          .Alignment = tLeftAlign
          .Mask = ""
          .MaxLength = 255
        Case Is = When
          HlpList.MultiSelect = True
          HlpList.PMCd "KT"
          'HlpList.OmNo "MW", "SO", "15", "REG"
    End Select
 
    Case Is = UCase("UoModTimeFr"), UCase("UoModTimeTo")
        Select Case Opt
        Case Is = Initialise
          .DataType = fdtTimeType
          .Alignment = tLeftAlign
          .Mask = ""
          .MaxLength = 8
        Case Is = When
          '***
        End Select
    Case Is = UCase("UoRmGrp")
        Select Case Opt
        Case Is = Initialise
          .DataType = fdtConvertToUpper
          .Alignment = tLeftAlign
          .Mask = ""
          .MaxLength = 8
        Case Is = When
          HlpList.PSCd "RMGRP", gs_RmGrp
        End Select
        
       
  '*************************Geeta*******************
  '********************Emr206*************************
      Case Is = UCase("UoLMCd1"), UCase("UoLMCD2"), UCase("UoLMCD3")
        Select Case Opt
        Case Is = Initialise
          .DataType = fdtConvertToUpper
          .Alignment = tLeftAlign
          .Mask = ""
          .MaxLength = 5
        Case Is = When
          HlpList.PMCd "LABMCD"
        End Select
  '*********************************************************************
    Case Is = UCase("UoLabGrp")
        Select Case Opt
        Case Is = Initialise
          .DataType = fdtConvertToUpper
          .Alignment = tLeftAlign
          .Mask = ""
          .MaxLength = 8
        Case Is = When
          HlpList.PSCd "LABGRP", gs_LabGrp
        End Select
    '***********Geeta ************ UoYn5, UoYN6,UoYn7 & UoYN8 Entered
    ' Manoj 2.10.0
    ' *** Jay 2.14Next ***  [UoYn14 Entered]
'    Case Is = UCase("UoYN"), UCase("UoYN5"), UCase("UoYN6"), UCase("UoYN7"), UCase("UoYN8")
    Case Is = UCase("UoYN"), UCase("UoYN5"), UCase("UoYN6"), UCase("UoYN7"), UCase("UoYN8"), UCase("UoYN9"), UCase("UoYN10"), UCase("UoYN11"), UCase("UoYN12"), UCase("UoYN13"), UCase("UoYN14")
    ' Manoj 2.10.0
      Select Case Opt
       Case Is = Initialise
          .DataType = fdtConvertToUpper
          .Alignment = tLeftAlign
          .Mask = ""
          .MaxLength = 1
        Case Is = When
          HlpList.PMCd "YN"
        End Select
    Case Is = UCase("UoYN1"), UCase("UoYN2"), UCase("UoYN3"), UCase("UoYN4"), UCase("UoYN29")   '4.1.0.0
        Select Case Opt
        Case Is = Initialise
          .DataType = fdtConvertToUpper
          .Alignment = tLeftAlign
          .Mask = ""
          .MaxLength = 1
        Case Is = When
          HlpList.PMCd "YN"
        End Select
    Case Is = UCase("UoLn1Fr"), UCase("UoLn1To")
        Select Case Opt
        Case Is = Initialise
          .DataType = fdtFloat
          .Alignment = tRightAlign
          .Mask = "##0.0000"    ' ***** Manali - 3.6.0 - 24/11/09 - Mask increased for pointer
          .MaxLength = 7
        Case Is = When
          '***
        End Select
    Case Is = UCase("UoLn2Fr"), UCase("UoLn2To")
        Select Case Opt
        Case Is = Initialise
          .DataType = fdtFloat
          .Alignment = tRightAlign
          .Mask = "#0.000"
          .MaxLength = 6
        Case Is = When
          '***
        End Select
    Case Is = UCase("UoSalRtFr"), UCase("UoSalRtTo")
        Select Case Opt
        Case Is = Initialise
          .DataType = fdtFloat
          .Alignment = tRightAlign
          .Mask = "######0.00"
          .MaxLength = 10
        Case Is = When
          '***
        End Select
    Case Is = UCase("UoCstRtFr"), UCase("UoCstRtTo")
        Select Case Opt
        Case Is = Initialise
          .DataType = fdtFloat
          .Alignment = tRightAlign
          .Mask = "-######0.000"
          .MaxLength = 12
        Case Is = When
          '***
        End Select
    Case Is = UCase("UoPrdQtyFr"), UCase("UoPrdQtyTo")
        Select Case Opt
        Case Is = Initialise
          .DataType = fdtFloat
          .Alignment = tRightAlign
          .Mask = "####0.0"
          .MaxLength = 7
        Case Is = When
          '***
        End Select
    Case Is = UCase("UoBalPrdFr"), UCase("UoBalPrdTo")
        Select Case Opt
        Case Is = Initialise
          .DataType = fdtFloat
          .Alignment = tRightAlign
          .Mask = "####0.0"
          .MaxLength = 7
        Case Is = When
          '***
        End Select
    '*** Jay 3.2.0 [Sub Po]
    Case Is = UCase("UoPoNoFr"), UCase("UoPoNoTo"), UCase("UoOdPoNoFr"), UCase("UoOdPoNoTo")
    '*** Jay 3.2.0 [Sub Po]
        Select Case Opt
        Case Is = Initialise
          .DataType = fdtConvertToUpper
          .Alignment = tLeftAlign
          .Mask = ""
          .MaxLength = 20
        Case Is = When
          '### HlpList
        End Select
    '*** Jay 3.2.0 [Sub Po]
    Case Is = UCase("UoPoNoSel"), UCase("UoOdPoNoSel")
    '*** Jay 3.2.0 [Sub Po]
        Select Case Opt
        Case Is = Initialise
          .DataType = fdtConvertToUpper
          .Alignment = tLeftAlign
          .Mask = ""
          .MaxLength = 255
        Case Is = When
          HlpList.MultiSelect = True
          '### HlpList
        End Select
    Case Is = UCase("UoDelDtFr"), UCase("UoDelDtTo")
        Select Case Opt
        Case Is = Initialise
          .DataType = fdtDateType
          .Alignment = tLeftAlign
          .Mask = ""
          .MaxLength = 8
        Case Is = When
          '***
        End Select
    Case Is = UCase("UoDelDtSel")
        Select Case Opt
        Case Is = Initialise
          .DataType = fdtConvertToUpper
          .Alignment = tLeftAlign
          .Mask = ""
          .MaxLength = 255
        Case Is = When
          '***
        End Select
    Case Is = UCase("UoDelDt1Fr"), UCase("UoDelDt1To")
        Select Case Opt
        Case Is = Initialise
          .DataType = fdtDateType
          .Alignment = tLeftAlign
          .Mask = ""
          .MaxLength = 8
        Case Is = When
          '***
        End Select
    Case Is = UCase("UoDelDt1Sel")
        Select Case Opt
        Case Is = Initialise
          .DataType = fdtConvertToUpper
          .Alignment = tLeftAlign
          .Mask = ""
          .MaxLength = 255
        Case Is = When
          '***
        End Select
    Case Is = UCase("UoWrkFr"), UCase("UoWrkTo")
        Select Case Opt
        Case Is = Initialise
          .DataType = fdtConvertToUpper
          .Alignment = tLeftAlign
          .Mask = ""
          .MaxLength = 8
        Case Is = When
          If gb_CoCdFor = True Then
            HlpList.vPMCd ADC("UoCoCdFr"), "WORK"
          Else
            If UCase(IdName) = UCase("UoWrkFr") Then HlpList.vPMCd ADC("UoCoCdFr"), "WORK"
            If UCase(IdName) = UCase("UoWrkTo") Then HlpList.vPMCd ADC("UoCoCdTo"), "WORK"
          End If
        End Select
    Case Is = UCase("UoWrkSel")
        Select Case Opt
        Case Is = Initialise
          .DataType = fdtConvertToUpper
          .Alignment = tLeftAlign
          .Mask = ""
          .MaxLength = 255
        Case Is = When
          HlpList.MultiSelect = True
          HlpList.vPMCd ADC("UoCoCdFr"), "WORK"
        End Select
    Case Is = UCase("UoSubPrcFr"), UCase("UoSubPrcTo")
        Select Case Opt
        Case Is = Initialise
          .DataType = fdtConvertToUpper
          .Alignment = tLeftAlign
          .Mask = ""
          .MaxLength = 8
        Case Is = When
          If UCase(IdName) = "UOSUBPRCFR" Then HlpList.PSCd "SUBPRC", ADC("UoIssLocFr")
          If UCase(IdName) = "UOSUBPRCTO" Then HlpList.PSCd "SUBPRC", ADC("UoIssLocTo")
        End Select
    Case Is = UCase("UoSubPrcSel")
        Select Case Opt
        Case Is = Initialise
          .DataType = fdtConvertToUpper
          .Alignment = tLeftAlign
          .Mask = ""
          .MaxLength = 255
        Case Is = When
          HlpList.MultiSelect = True
          HlpList.PSCd "SUBPRC", ADC("UoIssLocFr")
        End Select
    Case Is = UCase("UoBagPcs")
        Select Case Opt
        Case Is = Initialise
          .DataType = fdtConvertToUpper
          .Alignment = tLeftAlign
          .Mask = ""
          .MaxLength = 1
        Case Is = When
          HlpList.PMCd "BAGPCS"
        End Select
    Case Is = UCase("UoBrkMsg")
        Select Case Opt
        Case Is = Initialise
          .DataType = fdtConvertToUpper
          .Alignment = tLeftAlign
          .Mask = ""
          .MaxLength = 1
        Case Is = When
          HlpList.PMCd "BRKMSG"
        End Select
    Case Is = UCase("UoShpCd")
        Select Case Opt
        Case Is = Initialise
          .DataType = fdtConvertToUpper
          .Alignment = tLeftAlign
          .Mask = ""
          .MaxLength = 1
        Case Is = When
          HlpList.PMCd "SHPCDGR"
        End Select
    Case Is = UCase("UoPrtCdFr"), UCase("UoPrtCdTo")
        Select Case Opt
        Case Is = Initialise
          .DataType = fdtConvertToUpper
          .Alignment = tLeftAlign
          .Mask = ""
          .MaxLength = 2
        Case Is = When
          HlpList.PMCd "PRTCD"
        End Select
    Case Is = UCase("UoPrtCdSel")
        Select Case Opt
        Case Is = Initialise
          .DataType = fdtConvertToUpper
          .Alignment = tLeftAlign
          .Mask = ""
          .MaxLength = 255
        Case Is = When
          HlpList.MultiSelect = True
          HlpList.PMCd "PRTCD"
        End Select
    Case Is = UCase("UoSeqCdFr")
        Select Case Opt
        Case Is = Initialise
          .DataType = fdtConvertToUpper
          .Alignment = tLeftAlign
          .Mask = ""
          .MaxLength = 8
        Case Is = When
          HlpList.vPMCd ctSelfCoCd, gs_PTyp, , , , "N"            '"WIWSEQ"
        End Select
    '****** Sachin 2.14.0 - [23-11-2006] - [Multi-Company] ******
    Case Is = UCase("UoPrntSeqCdFr")
        Select Case Opt
        Case Is = Initialise
          .DataType = fdtConvertToUpper
          .Alignment = tLeftAlign
          .Mask = ""
          .MaxLength = 5
        Case Is = When
          HlpList.vPMCd ctSelfCoCd, gs_PTyp, , , , "Y"                '"WIWSEQ"
        End Select
    '****** Sachin 2.14.0 - [23-11-2006] - [Multi-Company] ******
    Case Is = UCase("UoYyMmFr"), UCase("UoYyMmTo")
        Select Case Opt
        Case Is = Initialise
          .DataType = fdtNumber
          .Alignment = tRightAlign
          .Mask = "#####0"
          .MaxLength = 6
        Case Is = When
          If gb_ShowYyMmHlp = True Then
            HlpList.YyyyMm
          End If
        End Select
    Case Is = UCase("UoOpnYyMm")
        Select Case Opt
        Case Is = Initialise
          .DataType = fdtNumber
          .Alignment = tRightAlign
          .Mask = "#####0"
          .MaxLength = 6
        Case Is = When
          If gb_ShowYyMmHlp = True Then
            HlpList.YyyyMm
          End If
        End Select
    Case Is = UCase("UoRejCdFr"), UCase("UoRejCdTo")
        Select Case Opt
        Case Is = Initialise
          .DataType = fdtConvertToUpper
          .Alignment = tLeftAlign
          .Mask = ""
          .MaxLength = 8
        Case Is = When
          HlpList.PMCd "REJCD"
        End Select
    Case Is = UCase("UoRejCdSel")
        Select Case Opt
        Case Is = Initialise
          .DataType = fdtConvertToUpper
          .Alignment = tLeftAlign
          .Mask = ""
          .MaxLength = 255
        Case Is = When
          HlpList.MultiSelect = True
          HlpList.PMCd "REJCD"
        End Select
    Case Is = UCase("UoLsDtFr"), UCase("UoLsDtTo")
        Select Case Opt
        Case Is = Initialise
          .DataType = fdtDateType
          .Alignment = tLeftAlign
          .Mask = ""
          .MaxLength = 8
        End Select
    Case Is = UCase("UoBOEFr"), UCase("UoBOETo")
        Select Case Opt
        Case Is = Initialise
          .DataType = fdtConvertToUpper
          .Alignment = tLeftAlign
          .Mask = ""
          .MaxLength = 20
        Case Is = When
          '###
        End Select
    Case Is = UCase("UoBOESel")
        Select Case Opt
        Case Is = Initialise
          .DataType = fdtConvertToUpper
          .Alignment = tLeftAlign
          .Mask = ""
          .MaxLength = 255
        Case Is = When
          HlpList.MultiSelect = True
          '###
        End Select
    Case Is = UCase("UoVchOpt")
        Select Case Opt
        Case Is = Initialise
          .DataType = fdtConvertToUpper
          .Alignment = tLeftAlign
          .Mask = ""
          .MaxLength = 8
        Case Is = When
          HlpList.PMCd "VCHOPT"
        End Select
    Case Is = UCase("UoDmColFr"), UCase("UoDmColTo")
        Select Case Opt
        Case Is = Initialise
          .DataType = fdtConvertToUpper
          .Alignment = tLeftAlign
          .Mask = ""
          .MaxLength = 5
        Case Is = When
          HlpList.PMCd "DMCOL"
        End Select
    Case Is = UCase("UoDmColSel")
        Select Case Opt
        Case Is = Initialise
          .DataType = fdtConvertToUpper
          .Alignment = tLeftAlign
          .Mask = ""
          .MaxLength = 255
        Case Is = When
          HlpList.MultiSelect = True
          HlpList.PMCd "DMCOL"
        End Select
    Case Is = UCase("UoPTypFr"), UCase("UoPTypTo")
        Select Case Opt
        Case Is = Initialise
          .DataType = fdtConvertToUpper
          .Alignment = tLeftAlign
          .Mask = ""
          .MaxLength = 8
        Case Is = When
          Select Case UCase(gs_Tbl)
          Case Is = UCase("Param")
            HlpList.PMCd "PTYP", "'" + gs_ParamOpt + "'"
          Case Is = UCase("vParam")
            'If ctProjectName = "EMR" Then
            '  HlpList.vPMCd ADC("UoCoCdFr"), "PTYP", "'" + gs_ParamOpt + "'"
            'ElseIf ctProjectName = "XEMR" Then
            '  HlpList.vPMCd ADC("UoxCoCdFr"), "PTYP", "'" + gs_ParamOpt + "'"
            'End If
            HlpList.vPMCd "", "PTYP", "'" + gs_ParamOpt + "'"
          Case Is = UCase("xParam")
            HlpList.xPMCd "PTYP", "'" + gs_ParamOpt + "'"
          Case Is = UCase("xvParam")
            'HlpList.xvPMCd ctSelfCoCd, "PTYP", "'" + gs_ParamOpt + "'"
            HlpList.xvPMCd "", "PTYP", "'" + gs_ParamOpt + "'"
          End Select
          'HlpList.PMCd "PTYP", "'" + gs_ParamOpt + "'"
        End Select
    Case Is = UCase("UoPTypSel")
        Select Case Opt
        Case Is = Initialise
          .DataType = fdtConvertToUpper
          .Alignment = tLeftAlign
          .Mask = ""
          .MaxLength = 255
        Case Is = When
          HlpList.MultiSelect = True
          Select Case UCase(gs_Tbl)
          Case Is = UCase("Param")
            HlpList.PMCd "PTYP", "'" + gs_ParamOpt + "'"
          Case Is = UCase("vParam")
            HlpList.vPMCd "", "PTYP", "'" + gs_ParamOpt + "'"
          Case Is = UCase("xParam")
            HlpList.xPMCd "PTYP", "'" + gs_ParamOpt + "'"
          Case Is = UCase("xvParam")
            HlpList.xvPMCd ctSelfCoCd, "PTYP", "'" + gs_ParamOpt + "'"
          End Select
          'HlpList.MultiSelect = True
          'HlpList.PMCd "PTYP", "'" + gs_ParamOpt + "'"
        End Select
    Case Is = UCase("UoPMCdFr"), UCase("UoPMCdTo"), UCase("UoPmCd1")    '4.1.4
        Select Case Opt
        Case Is = Initialise
          .DataType = fdtConvertToUpper
          .Alignment = tLeftAlign
          .Mask = ""
          .MaxLength = 8
        Case Is = When
          Select Case UCase(gs_Tbl)
          Case Is = UCase("Param")
            HlpList.PMCd gs_PTyp
          Case Is = UCase("vParam")
            HlpList.vPMCd "", gs_PTyp
          Case Is = UCase("xParam")
            HlpList.xPMCd gs_PTyp
          Case Is = UCase("xvParam")
            HlpList.xvPMCd ADC("UoCoCdFr"), gs_PTyp
          End Select
        End Select
    Case Is = UCase("UoPMCdSel")
        Select Case Opt
        Case Is = Initialise
          .DataType = fdtConvertToUpper
          .Alignment = tLeftAlign
          .Mask = ""
          .MaxLength = 255
        Case Is = When
          HlpList.MultiSelect = True
          Select Case UCase(gs_Tbl)
          Case Is = UCase("Param")
            HlpList.PMCd gs_PTyp
          Case Is = UCase("vParam")
            HlpList.vPMCd "", gs_PTyp
          Case Is = UCase("xParam")
            HlpList.xPMCd gs_PTyp
          Case Is = UCase("xvParam")
            HlpList.xvPMCd ADC("UoCoCdFr"), gs_PTyp
          End Select
          'HlpList.PMCd gs_PTyp
        End Select
    Case Is = UCase("UoLSCdFr"), UCase("UoLSCdTo")
        Select Case Opt
        Case Is = Initialise
          .DataType = fdtConvertToUpper
          .Alignment = tLeftAlign
          .Mask = ""
          .MaxLength = 8
        Case Is = When
          HlpList.PSCd "LABSCD", gs_LabMCd
        End Select
        'Amore.1 - Help added for Set Cd selection
    Case Is = UCase("UoLSCdSel")
        Select Case Opt
        Case Is = Initialise
          .DataType = fdtConvertToUpper
          .Alignment = tLeftAlign
          .Mask = ""
          .MaxLength = 255
        Case Is = When
          HlpList.MultiSelect = True
          HlpList.PSCd "LABSCD", gs_LabMCd
        End Select
    Case Is = UCase("UoFgTcFr"), UCase("UoFgTcTo")
        Select Case Opt
        Case Is = Initialise
          .DataType = fdtConvertToUpper
          .Alignment = tLeftAlign
          .Mask = ""
          .MaxLength = 3
        Case Is = When
          HlpList.PMCd "TC", "'" + gs_FgTcTyp + "'"
        End Select
    Case Is = UCase("UoFgYyFr"), UCase("UoFgYyTo")
        Select Case Opt
        Case Is = Initialise
          .DataType = fdtConvertToUpper
          .Alignment = tLeftAlign
          .Mask = ""
          .MaxLength = 2
        Case Is = When
          HlpList.PMCd "YY"
        End Select
    Case Is = UCase("UoFgChrFr"), UCase("UoFgChrTo")
        Select Case Opt
        Case Is = Initialise
          .DataType = fdtConvertToUpper
          .Alignment = tLeftAlign
          .Mask = ""
          .MaxLength = 3
        Case Is = When
          If UCase(IdName) = UCase("UoFgChrFr") Then HlpList.vPSCd ADC("UoCoCdFr"), "CHR", ADC("UoFgTcFr")
          If UCase(IdName) = UCase("UoFgChrTo") Then HlpList.vPSCd ADC("UoCoCdFr"), "CHR", ADC("UoFgTcTo")
        End Select
    Case Is = UCase("UoFgNoFr"), UCase("UoFgNoTo")
        Select Case Opt
        Case Is = Initialise
          .DataType = fdtNumber
          .Alignment = tRightAlign
          .Mask = "######0"
          .MaxLength = 7
        Case Is = When
          If UCase(IdName) = UCase("UoFgNoFr") Then HlpList.FgNo ADC("UoCoCdFr"), ADC("UoFgTcFr"), ADC("UoFgYyFr"), ADC("UoFgChrFr")
          If UCase(IdName) = UCase("UoFgNoTo") Then HlpList.FgNo ADC("UoCoCdFr"), ADC("UoFgTcTo"), ADC("UoFgYyTo"), ADC("UoFgChrTo")
        End Select
    Case Is = UCase("UoFdSrFr"), UCase("UoFdSrTo")
        Select Case Opt
        Case Is = Initialise
          .DataType = fdtNumber
          .Alignment = tRightAlign
          .Mask = "###0"
          .MaxLength = 4
        Case Is = When
          If UCase(IdName) = UCase("UoFdSrFr") Then HlpList.FdSr ADC("UoCoCdFr"), ADC("UoFgTcFr"), ADC("UoFgYyFr"), ADC("UoFgChrFr"), ADC("UoFgNoFr")
          If UCase(IdName) = UCase("UoFdSrTo") Then HlpList.FdSr ADC("UoCoCdFr"), ADC("UoFgTcTo"), ADC("UoFgYyTo"), ADC("UoFgChrTo"), ADC("UoFgNoTo")
        End Select
    
    '*** (Jen 2.14 Next)
    Case Is = UCase("UoReYyFr"), UCase("UoReYyTo")
        Select Case Opt
        Case Is = Initialise
          .DataType = fdtConvertToUpper
          .Alignment = tLeftAlign
          .Mask = ""
          .MaxLength = 2
        Case Is = When
          HlpList.PMCd "YY"
        End Select
    Case Is = UCase("UoReChrFr"), UCase("UoReChrTo")
        Select Case Opt
        Case Is = Initialise
          .DataType = fdtConvertToUpper
          .Alignment = tLeftAlign
          .Mask = ""
          .MaxLength = 3
        Case Is = When
          HlpList.PMCd "RETCHR"
        End Select
    Case Is = UCase("UoReNoFr"), UCase("UoReNoTo")
        Select Case Opt
        Case Is = Initialise
          .DataType = fdtNumber
          .Alignment = tRightAlign
          .Mask = "######0"
          .MaxLength = 7
        Case Is = When
          If UCase(IdName) = UCase("UoReNoFr") Then HlpList.ReNo ADC("UoCoCdFr"), ADC("UoReYyFr"), ADC("UoReChrFr")
          If UCase(IdName) = UCase("UoReNoTo") Then HlpList.ReNo ADC("UoCoCdFr"), ADC("UoReYyTo"), ADC("UoReChrTo")
        End Select
    '*** (Jen 2.14 Next)

    
    Case Is = UCase("UoExpOdTcFr"), UCase("UoExpOdTcTo")
        Select Case Opt
        Case Is = Initialise
          .DataType = fdtConvertToUpper
          .Alignment = tLeftAlign
          .Mask = ""
          .MaxLength = 3
        Case Is = When
          HlpList.PMCd "TC", "'" + gs_OmTcTyp + "'"
        End Select
    Case Is = UCase("UoExpOdYyFr"), UCase("UoExpOdYyTo")
        Select Case Opt
        Case Is = Initialise
          .DataType = fdtConvertToUpper
          .Alignment = tLeftAlign
          .Mask = ""
          .MaxLength = 2
        Case Is = When
          HlpList.PMCd "YY"
        End Select
    Case Is = UCase("UoExpOdChrFr"), UCase("UoExpOdChrTo")
        Select Case Opt
        Case Is = Initialise
          .DataType = fdtConvertToUpper
          .Alignment = tLeftAlign
          .Mask = ""
          .MaxLength = 3
        Case Is = When
          If UCase(IdName) = UCase("UoExpOdChrFr") Then HlpList.vPSCd ADC("UoCoCdFr"), "CHR", ADC("UoExpOdTcFr")
          If UCase(IdName) = UCase("UoExpOdChrTo") Then HlpList.vPSCd ADC("UoCoCdFr"), "CHR", ADC("UoExpOdTcTo")
        End Select
    Case Is = UCase("UoExpOdNoFr"), UCase("UoExpOdNoTo")
        Select Case Opt
        Case Is = Initialise
          .DataType = fdtNumber
          .Alignment = tRightAlign
          .Mask = "######0"
          .MaxLength = 7
        Case Is = When
          If UCase(IdName) = UCase("UoExpOdNoFr") Then HlpList.OmNo ADC("UoCoCdFr"), ADC("UoExpOdTcFr"), ADC("UoExpOdYyFr"), ADC("UoExpOdChrFr")
          If UCase(IdName) = UCase("UoExpOdNoTo") Then HlpList.OmNo ADC("UoCoCdFr"), ADC("UoExpOdTcFr"), ADC("UoExpOdYyFr"), ADC("UoExpOdChrFr")
        End Select
    Case Is = UCase("UoExpOdSrFr"), UCase("UoExpOdSrTo")
        Select Case Opt
        Case Is = Initialise
          .DataType = fdtNumber
          .Alignment = tRightAlign
          .Mask = "###0"
          .MaxLength = 4
        Case Is = When
          If UCase(IdName) = UCase("UoExpOdSrFr") Then HlpList.OdSr ADC("UoCoCdFr"), ADC("UoExpOdTcFr"), ADC("UoExpOdYyFr"), ADC("UoExpOdChrFr"), ADC("UoExpOdNoFr")
          If UCase(IdName) = UCase("UoExpOdSrTo") Then HlpList.OdSr ADC("UoCoCdFr"), ADC("UoExpOdTcTo"), ADC("UoExpOdYyTo"), ADC("UoExpOdChrTo"), ADC("UoExpOdNoTo")
        End Select
    Case Is = UCase("UoExpCmCdFr"), UCase("UoExpCmCdTo")
        Select Case Opt
        Case Is = Initialise
          .DataType = fdtConvertToUpper
          .Alignment = tLeftAlign
          .Mask = ""
          .MaxLength = 8
        Case Is = When
          HlpList.CustCd gs_CmCtg
        End Select
    Case Is = UCase("UoExpCmCdSel")
        Select Case Opt
        Case Is = Initialise
          .DataType = fdtConvertToUpper
          .Alignment = tLeftAlign
          .Mask = ""
          .MaxLength = 255
        Case Is = When
          HlpList.MultiSelect = True
          HlpList.CustCd gs_CmCtg
        End Select
    Case Is = UCase("UoFdPlYyFr"), UCase("UoFdPlYyTo")
        Select Case Opt
        Case Is = Initialise
          .DataType = fdtConvertToUpper
          .Alignment = tLeftAlign
          .Mask = ""
          .MaxLength = 2
        Case Is = When
          HlpList.PMCd "YY"
        End Select
    Case Is = UCase("UoFdPlChrFr"), UCase("UoFdPlChrTo")
        Select Case Opt
        Case Is = Initialise
          .DataType = fdtConvertToUpper
          .Alignment = tLeftAlign
          .Mask = ""
          .MaxLength = 3
        Case Is = When
          HlpList.PMCd "PLCHR"
        End Select
    Case Is = UCase("UoFdPlNoFr"), UCase("UoFdPlNoTo")
        Select Case Opt
        Case Is = Initialise
          .DataType = fdtNumber
          .Alignment = tRightAlign
          .Mask = "######0"
          .MaxLength = 7
        Case Is = When
          If UCase(IdName) = UCase("UoFdPlNoFr") Then HlpList.PlNo ADC("UoCoCdFr"), ADC("UoFdPlYyFr"), ADC("UoFdPlChrFr")
          If gb_CoCdFor = True Then
            If UCase(IdName) = UCase("UoFdPlNoTo") Then HlpList.PlNo ADC("UoCoCdFr"), ADC("UoFdPlYyTo"), ADC("UoFdPlChrTo")
          Else
            If UCase(IdName) = UCase("UoFdPlNoTo") Then HlpList.PlNo ADC("UoCoCdTo"), ADC("UoFdPlYyTo"), ADC("UoFdPlChrTo")
          End If
          '### If UCase(IdName) = UCase("UoFdPlNoFr") Then HlpList.PlNo ADC("UoFdPlYyFr"), ADC("UoFdPlChrFr")
          '### If UCase(IdName) = UCase("UoFdPlNoTo") Then HlpList.PlNo ADC("UoFdPlYyTo"), ADC("UoFdPlChrTo")
        End Select
    Case Is = UCase("UoInTcFr"), UCase("UoInTcTo")
        Select Case Opt
        Case Is = Initialise
          .DataType = fdtConvertToUpper
          .Alignment = tLeftAlign
          .Mask = ""
          .MaxLength = 3
        Case Is = When
          HlpList.PMCd "TC", "'" + gs_InTcTyp + "'"
        End Select
    Case Is = UCase("UoInYyFr"), UCase("UoInYyTo")
        Select Case Opt
        Case Is = Initialise
          .DataType = fdtConvertToUpper
          .Alignment = tLeftAlign
          .Mask = ""
          .MaxLength = 2
        Case Is = When
          HlpList.PMCd "YY"
        End Select
    Case Is = UCase("UoInChrFr"), UCase("UoInChrTo")
        Select Case Opt
        Case Is = Initialise
          .DataType = fdtConvertToUpper
          .Alignment = tLeftAlign
          .Mask = ""
          .MaxLength = 3
        Case Is = When
          '****** Sachin 2.14.0 - [27-11-2006] - [Multi-Company] ******
          If UCase(IdName) = UCase("UoInChrFr") Then HlpList.vPSCd ADC("UoCoCdFr"), "CHR", ADC("UoInTcFr")
          If gb_CoCdFor = False Then
            If UCase(IdName) = UCase("UoInChrTo") Then HlpList.vPSCd ADC("UoCoCdTo"), "CHR", ADC("UoInTcTo")
          Else
            If UCase(IdName) = UCase("UoInChrTo") Then HlpList.vPSCd ADC("UoCoCdFr"), "CHR", ADC("UoInTcTo")
          End If
          '****** Sachin 2.14.0 - [27-11-2006] - [Multi-Company] ******
        End Select
    Case Is = UCase("UoInNoFr"), UCase("UoInNoTo")
        Select Case Opt
        Case Is = Initialise
          .DataType = fdtNumber
          .Alignment = tRightAlign
          .Mask = "######0"
          .MaxLength = 7
        Case Is = When
          If Not gb_InHlpFrOm Then
            If UCase(IdName) = UCase("UoInNoFr") Then HlpList.InNo ADC("UoCoCdFr"), ADC("UoInTcFr"), ADC("UoInYyFr"), ADC("UoInChrFr")
            '****** Sachin 2.14.0 - [27-11-2006] - [Multi-Company] ******
            If gb_CoCdFor = False Then
              If UCase(IdName) = UCase("UoInNoTo") Then HlpList.InNo ADC("UoCoCdTo"), ADC("UoInTcTo"), ADC("UoInYyTo"), ADC("UoInChrTo")
            Else
              If UCase(IdName) = UCase("UoInNoTo") Then HlpList.InNo ADC("UoCoCdFr"), ADC("UoInTcTo"), ADC("UoInYyTo"), ADC("UoInChrTo")
            End If
            '****** Sachin 2.14.0 - [27-11-2006] - [Multi-Company] ******
          ' ***** Manali 3.7.1
          Else
            If UCase(IdName) = UCase("UoInNoFr") Then HlpList.OmNo ADC("UoCoCdFr"), ADC("UoInTcFr"), ADC("UoInYyFr"), ADC("UoInChrFr")
            If gb_CoCdFor = False Then
              If UCase(IdName) = UCase("UoInNoTo") Then HlpList.OmNo ADC("UoCoCdTo"), ADC("UoInTcTo"), ADC("UoInYyTo"), ADC("UoInChrTo")
            Else
              If UCase(IdName) = UCase("UoInNoTo") Then HlpList.OmNo ADC("UoCoCdFr"), ADC("UoInTcTo"), ADC("UoInYyTo"), ADC("UoInChrTo")
            End If
          End If
          ' ***** Manali 3.7.1
        End Select
    Case Is = UCase("UoInExpNoFr"), UCase("UoInExpNoTo")
        Select Case Opt
        Case Is = Initialise
          .DataType = fdtConvertToUpper
          .Alignment = tLeftAlign
          .Mask = ""
          .MaxLength = 20
        Case Is = When
          '### HlpList
          '****** Sachin 2.14.0 - [27-11-2006] - [Multi-Company] ******
          '*** Jay 3.3(FG)
          If UCase(IdName) = UCase("UoInExpNoFr") Then HlpList.InExpNo ADC("UoCoCdFr"), gs_InTcTyp
          If gb_CoCdFor = False Then
            If UCase(IdName) = UCase("UoInExpNoTo") Then HlpList.InExpNo ADC("UoCoCdTo"), gs_InTcTyp
          Else
            If UCase(IdName) = UCase("UoInExpNoTo") Then HlpList.InExpNo ADC("UoCoCdFr"), gs_InTcTyp
          End If
          '*** Jay 3.3(FG)
        End Select
    Case Is = UCase("UoInExpNoSel")
        Select Case Opt
        Case Is = Initialise
          .DataType = fdtConvertToUpper
          .Alignment = tLeftAlign
          .Mask = ""
          .MaxLength = 255
        Case Is = When
          HlpList.MultiSelect = True
          '### HlpList
          HlpList.InExpNo ADC("UoCoCdFr"), gs_InTcTyp     '*** Jay 3.3(FG)
        End Select
    Case Is = UCase("UoInDet")
        Select Case Opt
        Case Is = Initialise
          .DataType = fdtConvertToUpper
          .Alignment = tLeftAlign
          .Mask = ""
          .MaxLength = 8
        Case Is = When
          HlpList.PMCd "INDET"
        End Select
    Case Is = UCase("UoInVal")
        Select Case Opt
        Case Is = Initialise
          .DataType = fdtConvertToUpper
          .Alignment = tLeftAlign
          .Mask = ""
          .MaxLength = 1
        Case Is = When
          HlpList.PMCd "INVAL"
        End Select
    Case Is = UCase("UoInVal1"), UCase("UoInVal2")  '4.1.4
        Select Case Opt
        Case Is = Initialise
          .DataType = fdtConvertToUpper
          .Alignment = tLeftAlign
          .Mask = ""
          .MaxLength = 1
        Case Is = When
          HlpList.PMCd "INVAL"
        End Select
    Case Is = UCase("UoInWt")
        Select Case Opt
        Case Is = Initialise
          .DataType = fdtConvertToUpper
          .Alignment = tLeftAlign
          .Mask = ""
          .MaxLength = 1
        Case Is = When
          HlpList.PMCd "INWT"
        End Select
    Case Is = UCase("UoBOEDtFr"), UCase("UoBOEDtFr")
        Select Case Opt
        Case Is = Initialise
          .DataType = fdtDateType
          .Alignment = tLeftAlign
          .Mask = ""
          .MaxLength = 8
        Case Is = When
          '***
        End Select
    Case Is = UCase("UoBillNoFr"), UCase("UoBillNoFr")
        Select Case Opt
        Case Is = Initialise
          .DataType = fdtConvertToUpper
          .Alignment = tLeftAlign
          .Mask = ""
          .MaxLength = 20
        Case Is = When
          '###
        End Select
    Case Is = UCase("UoAsOfDt")
        Select Case Opt
        Case Is = Initialise
          .DataType = fdtDateType
          .Alignment = tLeftAlign
          .Mask = ""
          .MaxLength = 8
        Case Is = When
          '***
        End Select
    Case Is = UCase("UoPcPerCtFr"), UCase("UoPcPerCtTo")
        Select Case Opt
        Case Is = Initialise
          .DataType = fdtNumber
          .Alignment = tRightAlign
          .Mask = "##0"
          .MaxLength = 3
        End Select
    Case Is = UCase("UoSubShpFr"), UCase("UoSubShpTo")
        Select Case Opt
        Case Is = Initialise
          .DataType = fdtConvertToUpper
          .Alignment = tLeftAlign
          .Mask = ""
          .MaxLength = 8    '11.b size changed from 3 to 8
        Case Is = When
          HlpList.PMCd "SUBSHP"
        End Select
    Case Is = UCase("UoSubShpSel")
        Select Case Opt
        Case Is = Initialise
          .DataType = fdtConvertToUpper
          .Alignment = tLeftAlign
          .Mask = ""
          .MaxLength = 255
        Case Is = When
          HlpList.MultiSelect = True
          HlpList.PMCd "SUBSHP"
        End Select
    Case Is = UCase("UoGrsWtFr"), UCase("UoGrsWtTo"), UCase("UoMainMetWtFr"), UCase("UoMainMetWtTo")
        Select Case Opt
        Case Is = Initialise
          .DataType = fdtFloat
          .Alignment = tRightAlign
          .Mask = "########0.000"
          .MaxLength = 13
        Case Is = When
          '***
        End Select
    Case Is = UCase("UoRmWtFr"), UCase("UoRmWtTo")
        Select Case Opt
        Case Is = Initialise
          .DataType = fdtFloat
          .Alignment = tRightAlign
          .Mask = "########0.000"
          .MaxLength = 13
        Case Is = When
          '***
        End Select
    Case Is = UCase("UoSaRmGrp")
        Select Case Opt
        Case Is = Initialise
          .DataType = fdtConvertToUpper
          .Alignment = tLeftAlign
          .Mask = ""
          .MaxLength = 8
        Case Is = When
          HlpList.PMCd "SARMGR"
        End Select
    Case Is = UCase("UoSaLabGrp")
        Select Case Opt
        Case Is = Initialise
          .DataType = fdtConvertToUpper
          .Alignment = tLeftAlign
          .Mask = ""
          .MaxLength = 8
        Case Is = When
          HlpList.PMCd "SALABGR"
        End Select
    Case Is = UCase("UoRngSzFr"), UCase("UoRngSzTo")
        Select Case Opt
        Case Is = Initialise
          .DataType = fdtConvertToUpper
          .Alignment = tLeftAlign
          .Mask = ""
          .MaxLength = 10
        Case Is = When
          If gb_RmCtgFor = True Then
            If UCase(IdName) = "UORNGSZFR" Then HlpList.RangeSz ADC("UoRmCtgFr"), ADC("UoRmSCtgFr")
            If UCase(IdName) = "UORNGSZTO" Then HlpList.RangeSz ADC("UoRmCtgFr"), ADC("UoRmSCtgTo")
          Else
            If UCase(IdName) = "UORNGSZFR" Then HlpList.RangeSz ADC("UoRmCtgFr"), ADC("UoRmSCtgFr")
            If UCase(IdName) = "UORNGSZTO" Then HlpList.RangeSz ADC("UoRmCtgTo"), ADC("UoRmSCtgTo")
          End If
        End Select
    Case Is = UCase("UoRngSzSel")
        Select Case Opt
        Case Is = Initialise
          .DataType = fdtConvertToUpper
          .Alignment = tLeftAlign
          .Mask = ""
          .MaxLength = 255
        Case Is = When
          HlpList.MultiSelect = True
          HlpList.RangeSz ADC("UoRmCtgFr"), ADC("UoRmSCtgFr")
        End Select
    Case Is = UCase("UoCmCtgFr"), UCase("UoCmCtgTo")
        Select Case Opt
        Case Is = Initialise
          .DataType = fdtConvertToUpper
          .Alignment = tLeftAlign
          .Mask = ""
          .MaxLength = 2
        Case Is = When
          HlpList.PMCd "CMCTG"
        End Select
    Case Is = UCase("UoSetCdFr"), UCase("UoSetCdTo")
        Select Case Opt
        Case Is = Initialise
          .DataType = fdtConvertToUpper
          .Alignment = tLeftAlign
          .Mask = ""
          .MaxLength = 15
        Case Is = When
          HlpList.DmSetCd
        End Select
    Case Is = UCase("UoSetCdSel")
        Select Case Opt
        Case Is = Initialise
          .DataType = fdtConvertToUpper
          .Alignment = tLeftAlign
          .Mask = ""
          .MaxLength = 255
        Case Is = When
          HlpList.MultiSelect = True
          HlpList.DmSetCd
        End Select
    Case Is = UCase("UoLsCtgFr"), UCase("UoLsCtgTo")
        Select Case Opt
        Case Is = Initialise
          .DataType = fdtConvertToUpper
          .Alignment = tLeftAlign
          .Mask = ""
          .MaxLength = 5
        Case Is = When
          HlpList.PMCd ("LSCTG")
        End Select
    Case Is = UCase("UoLsCtgSel")
        Select Case Opt
        Case Is = Initialise
          .DataType = fdtConvertToUpper
          .Alignment = tLeftAlign
          .Mask = ""
          .MaxLength = 255
        Case Is = When
          HlpList.MultiSelect = True
          HlpList.PMCd ("LSCTG")
        End Select
    Case Is = UCase("UoSTWGrpFr"), UCase("UoSTWGrpTo")
        Select Case Opt
        Case Is = Initialise
          .DataType = fdtConvertToUpper
          .Alignment = tLeftAlign
          .Mask = ""
          .MaxLength = 5
        Case Is = When
          HlpList.PMCd ("STWGRP")
        End Select
    Case Is = UCase("UoSTWGrpSel")
        Select Case Opt
        Case Is = Initialise
          .DataType = fdtConvertToUpper
          .Alignment = tLeftAlign
          .Mask = ""
          .MaxLength = 255
        Case Is = When
          HlpList.MultiSelect = True
          HlpList.PMCd ("STWGRP")
        End Select
    Case Is = UCase("UoParentCdFr"), UCase("UoParentCdTo")
        Select Case Opt
        Case Is = Initialise
          .DataType = fdtConvertToUpper
          .Alignment = tLeftAlign
          .Mask = ""
          .MaxLength = 8
        Case Is = When
          HlpList.PMCd ("PARENTCD")
        End Select
    Case Is = UCase("UoParentCdSel")
        Select Case Opt
        Case Is = Initialise
          .DataType = fdtConvertToUpper
          .Alignment = tLeftAlign
          .Mask = ""
          .MaxLength = 255
        Case Is = When
          HlpList.MultiSelect = True
          HlpList.PMCd ("PARENTCD")
        End Select
    Case Is = UCase("UoPrc")
        Select Case Opt
        Case Is = Initialise
          .DataType = fdtFloat
          .Alignment = tRightAlign
          .Mask = "-##0.000"
          .MaxLength = 8
        Case Is = When
          '***
        End Select
    Case Is = UCase("UoOrdVal")
        Select Case Opt
        Case Is = Initialise
          .DataType = fdtConvertToUpper
          .Alignment = tLeftAlign
          .Mask = ""
          .MaxLength = 1
        Case Is = When
          HlpList.PMCd "ORDVAL"
        End Select
    Case Is = UCase("UoOrdVal1")
        Select Case Opt
        Case Is = Initialise
          .DataType = fdtConvertToUpper
          .Alignment = tLeftAlign
          .Mask = ""
          .MaxLength = 1
        Case Is = When
          HlpList.PMCd "ORDVAL"
        End Select
    Case Is = UCase("UoCoCdFr"), UCase("UoCoCdTo")
        Select Case Opt
        Case Is = Initialise
          .DataType = fdtConvertToUpper
          .Alignment = tLeftAlign
          .Mask = ""
          .MaxLength = 3
        Case Is = When
          HlpList.hCoCd
        End Select
    Case Is = UCase("UoCoCdSel")
        Select Case Opt
        Case Is = Initialise
          .DataType = fdtConvertToUpper
          .Alignment = tLeftAlign
          .Mask = ""
          .MaxLength = 255
        Case Is = When
          HlpList.MultiSelect = True
          HlpList.hCoCd
        
        '*** (Jen 2.14)
        Case Is = Valid
          If .TxtText <> "" Then
            Cancel = Not (ChkVal("Select 'x' from Head where HCoCd in (" + .TxtText + ")", ErrMsg))
            If Cancel = True Then ErrMsg = "Invalid Company Code Selection"
          End If
        '*** (Jen 2.14)
        
        End Select
    Case Is = UCase("UoSalPerFr"), UCase("UoSalPerTo")
        Select Case Opt
        Case Is = Initialise
          .DataType = fdtConvertToUpper
          .Alignment = tLeftAlign
          .Mask = ""
          .MaxLength = 8
        Case Is = When
          HlpList.PMCd "SP"
        End Select
    Case Is = UCase("UoSalPerSel")
        Select Case Opt
        Case Is = Initialise
          .DataType = fdtConvertToUpper
          .Alignment = tLeftAlign
          .Mask = ""
          .MaxLength = 255
        Case Is = When
          HlpList.MultiSelect = True
          HlpList.PMCd "SP"
        End Select
    Case Is = UCase("UoPMCd")
        Select Case Opt
        Case Is = Initialise
          .DataType = fdtConvertToUpper
          .Alignment = tLeftAlign
          .Mask = ""
          .MaxLength = 8
        Case Is = When
          HlpList.PMCd gs_PTyp1
        End Select
    Case Is = UCase("UoLineFr"), UCase("UoLineTo")
        Select Case Opt
        Case Is = Initialise
          .DataType = fdtConvertToUpper
          .Alignment = tLeftAlign
          .Mask = ""
          .MaxLength = 5
        Case Is = When
          HlpList.vPMCd ADC("UoCoCdFr"), "LINE"
        End Select
    Case Is = UCase("UoLineSel")
        Select Case Opt
        Case Is = Initialise
          .DataType = fdtConvertToUpper
          .Alignment = tLeftAlign
          .Mask = ""
          .MaxLength = 255
        Case Is = When
          HlpList.MultiSelect = True
          HlpList.vPMCd ADC("UoCoCdFr"), "LINE"
        End Select
    
' ########################  Manoj #### Ver: 2.11 Graph #### Date: 26/05/2005  ########################
    Case Is = UCase$("UoShowGrph")
        Select Case Opt
        Case Is = Initialise
          .DataType = fdtConvertToUpper
          .Alignment = tLeftAlign
          .Mask = ""
          .MaxLength = 1
        Case Is = When
          HlpList.PMCd "YN"
        End Select
    Case Is = UCase$("UoGrphTyp")
        Select Case Opt
        Case Is = Initialise
          .DataType = fdtConvertToUpper
          .Alignment = tLeftAlign
          .Mask = ""
          .MaxLength = 8
        Case Is = When
          If gs_GrphGrpNo <> 0 Then
            HlpList.PMCd "GrTyp", , , "(PNum1 = 0 Or PNum1 = " + CStr(gs_GrphGrpNo) + ")"
          Else
            HlpList.PMCd "GrTyp", , , "(PNum1 >= 0) "
          End If
        End Select
    Case Is = UCase$("UoGrphDir")
        Select Case Opt
        Case Is = Initialise
          .DataType = fdtConvertToUpper
          .Alignment = tLeftAlign
          .Mask = ""
          .MaxLength = 8
        Case Is = When
          HlpList.PMCd "GrDir"
        End Select
    Case Is = UCase$("UoGrphDataLb")
        Select Case Opt
        Case Is = Initialise
          .DataType = fdtConvertToUpper
          .Alignment = tLeftAlign
          .Mask = ""
          .MaxLength = 8
        Case Is = When
          HlpList.PMCd "GrDataLb"
        End Select
    Case Is = UCase$("UoGrphVwAngl")
        Select Case Opt
        Case Is = Initialise
          .DataType = fdtConvertToUpper
          .Alignment = tLeftAlign
          .Mask = ""
          .MaxLength = 8
        Case Is = When
          HlpList.PMCd "GrVwAngl"
        End Select
    Case Is = UCase$("UoGrphG1TopN"), UCase$("UoGrphG2TopN")
        Select Case Opt
        Case Is = Initialise
          .DataType = fdtConvertToUpper
          .Alignment = tLeftAlign
          .Mask = ""
          .MaxLength = 8
        Case Is = When
          HlpList.PMCd "GRTOPN"
        End Select
    Case Is = UCase$("UoGrphG1NVal"), UCase$("UoGrphG2NVal")
        Select Case Opt
        Case Is = Initialise
          .DataType = fdtNumber
          .Alignment = tRightAlign
          .Mask = "####0"
          .MaxLength = 5
        End Select
    Case Is = UCase$("UoGrphSumm")
        Select Case Opt
        Case Is = Initialise
          .DataType = fdtConvertToUpper
          .Alignment = tLeftAlign
          .Mask = ""
          .MaxLength = 8
        Case Is = When
          HlpList.PMCd "GRSUMM"
        End Select
    Case Is = UCase$("UoGrphClr")
        Select Case Opt
        Case Is = Initialise
          .DataType = fdtConvertToUpper
          .Alignment = tLeftAlign
          .Mask = ""
          .MaxLength = 1
        Case Is = When
          HlpList.PMCd "YN"
        End Select
    Case Is = UCase$("UoGrphOrient")
        Select Case Opt
        Case Is = Initialise
          .DataType = fdtConvertToUpper
          .Alignment = tLeftAlign
          .Mask = ""
          .MaxLength = 8
        Case Is = When
          HlpList.PMCd "CRORIENT"
        End Select
    Case Is = UCase$("UoGrphPSZ")
        Select Case Opt
        Case Is = Initialise
          .DataType = fdtConvertToUpper
          .Alignment = tLeftAlign
          .Mask = ""
          .MaxLength = 8
        Case Is = When
          HlpList.PMCd "CRPPRSZ"
        End Select

' ###########################################  Manoj  ###########################################
    
'    Case Is = UCase("UoSrlPort")
'        Select Case Opt
'        Case Is = Initialise
'          .DataType = fdtNumber
'          .Alignment = tRightAlign
'          .Mask = "#0"
'          .MaxLength = 2
'        End Select
    
    Case Is = UCase("UoBrNoFr"), UCase("UoBrNoTo")
        Select Case Opt
        Case Is = Initialise
          .DataType = fdtNumber
          .Alignment = tRightAlign
          .Mask = "######0"
          .MaxLength = 7
        Case Is = When
          If UCase(IdName) = UCase("UoBrNoFr") Then HlpList.BrNo ADC("UoCoCdFr"), ADC("UoTTcFr"), ADC("UoTYyFr"), ADC("UoTChrFr")
          If UCase(IdName) = UCase("UoBrNoTo") Then HlpList.BrNo ADC("UoCoCdFr"), ADC("UoTTcTo"), ADC("UoTYyTo"), ADC("UoTChrTo")
        End Select
    Case Is = UCase("UoBrSrFr"), UCase("UoBrSrTo")
        Select Case Opt
        Case Is = Initialise
          .DataType = fdtNumber
          .Alignment = tRightAlign
          .Mask = "###0"
          .MaxLength = 4
'        Case Is = When
'          If UCase(IdName) = UCase("UoBrSrFr") Then HlpList.BrSr adc("UoCoCdFr"), adc("UoTTcFr"), adc("UoTYyFr"), adc("UoTChrFr"), adc("UoBrNoFr")
'          If UCase(IdName) = UCase("UoBrSrTo") Then HlpList.BrSr adc("UoCoCdFr"), adc("UoTTcTo"), adc("UoTYyTo"), adc("UoTChrTo"), adc("UoBrNoTo")
        End Select
    ' Zubin 212
    Case Is = UCase("UoLotNoFr"), UCase("UoLotNoTo")
        Select Case Opt
        Case Is = Initialise
          .DataType = fdtConvertToUpper
          .Alignment = tLeftAlign
          .Mask = ""
          .MaxLength = 16
        Case Is = When
          'If UCase(IdName) = UCase("UoLotNoFr") Then HlpList.LotNo ADC("UoCoCdFr"), ADC("UoIssLocFr"), ADC("UoRmCdFr"), ADC("UoLn1Fr")
          'If UCase(IdName) = UCase("UoLotNoTo") Then HlpList.LotNo ADC("UoCoCdFr"), ADC("UoIssLocTo"), ADC("UoRmCdTo"), ADC("UoLn1To")
          '*** Jay 3.01 Bef **** [LotNo Err]
          'If UCase(IdName) = UCase("UoLotNoFr") Then HlpList.LotNo ADC("UoCoCdFr"), ADC("UoIssLocFr"), ADC("UoRmCdFr")
          'If UCase(IdName) = UCase("UoLotNoTo") Then HlpList.LotNo ADC("UoCoCdFr"), ADC("UoIssLocTo"), ADC("UoRmCdTo")
          '*** Jay 3.01 Bef **** [LotNo Err]
          '*** Jay 3.01 **** [LotNo Err]
          If UCase(IdName) = UCase("UoLotNoFr") Then HlpList.LotNo ADC("UoCoCdFr"), gs_Loc, ADC("UoRmCdFr"), , , , gs_RmLocTyp
          If UCase(IdName) = UCase("UoLotNoTo") Then HlpList.LotNo ADC("UoCoCdFr"), gs_Loc, ADC("UoRmCdTo"), , , , gs_RmLocTyp
        End Select
    Case Is = UCase("UoLotNoSel")
        Select Case Opt
        Case Is = Initialise
          .DataType = fdtConvertToUpper
          .Alignment = tLeftAlign
          .Mask = ""
          .MaxLength = 255
        Case Is = When
          HlpList.MultiSelect = True
          'HlpList.LotNo ADC("UoCoCdFr"), ADC("UoIssLocFr"), ADC("UoRmCdFr"), ADC("UoLn1Fr")
          'HlpList.LotNo adc("UoCoCdFr"), adc("UoIssLocFr"), adc("UoRmCdFr")
          ' ******* Manali 350Nxt
          HlpList.LotNo ADC("UoCoCdFr")
          ' ******* Manali 350Nxt
        End Select
    ' Zubin 212

    '*** Jay 2.13(CT) ***
    Case Is = UCase("UoBagTyp")
        Select Case Opt
        Case Is = Initialise
          .DataType = fdtConvertToUpper
          .Alignment = tLeftAlign
          .Mask = ""
          .MaxLength = 1
        Case Is = When
          HlpList.PMCd "BTYP"
        End Select
    '*** Jay 2.13(CT) ***
    
    ' Zubin (3.02) - 250308
    Case Is = UCase("UoPSCdSel")
        Select Case Opt
        Case Is = Initialise
          .DataType = fdtConvertToUpper
          .Alignment = tLeftAlign
          .Mask = ""
          .MaxLength = 255
        Case Is = When
          HlpList.MultiSelect = True
          HlpList.PSCd gs_PTyp, gs_PMCd
        End Select
    ' Zubin (3.02) - 250308


'****************************** XEMR **************************************************
'****************************** XEMR **************************************************
'****************************** XEMR **************************************************
    Case Is = UCase("UoxCoCdFr"), UCase("UoxCoCdTo")
        Select Case Opt
        Case Is = Initialise
          .DataType = fdtConvertToUpper
          .Alignment = tLeftAlign
          .Mask = ""
          .MaxLength = 3
        Case Is = When
          HlpList.xHCd
        End Select
    
    Case Is = UCase("UoxCoCdSel")
        Select Case Opt
        Case Is = Initialise
          .DataType = fdtConvertToUpper
          .Alignment = tLeftAlign
          .Mask = ""
          .MaxLength = 255
        Case Is = When
          HlpList.MultiSelect = True
          HlpList.xHCd
        End Select
    Case Is = UCase("UoDsgLocTypFr"), UCase("UoDsgLocTypTo")
        Select Case Opt
        Case Is = Initialise
          .DataType = fdtConvertToUpper
          .Alignment = tLeftAlign
          .Mask = ""
          .MaxLength = 2
        Case Is = When
          HlpList.xPMCd "LOCDSGTY"
        End Select
        
    Case Is = UCase("UoDsgLocTypSel")
        Select Case Opt
        Case Is = Initialise
          .DataType = fdtConvertToUpper
          .Alignment = tLeftAlign
          .Mask = ""
          .MaxLength = 255
        Case Is = When
          HlpList.MultiSelect = True
          HlpList.xPMCd "LOCDSGTY"
        End Select
    Case Is = UCase("UoDsgLocFr"), UCase("UoDsgLocTo")
        Select Case Opt
        Case Is = Initialise
          .DataType = fdtConvertToUpper
          .Alignment = tLeftAlign
          .Mask = ""
          .MaxLength = 5
        Case Is = When
          If Trim(gs_DsgLocTyp) <> "" Then
            HlpList.xDsgLoc gs_DsgLocTyp
          Else
            HlpList.xDsgLoc
          End If
        End Select
    
    Case Is = UCase("UoDsgLocSel")
        Select Case Opt
        Case Is = Initialise
          .DataType = fdtConvertToUpper
          .Alignment = tLeftAlign
          .Mask = ""
          .MaxLength = 255
        Case Is = When
          HlpList.MultiSelect = True
            HlpList.xDsgLoc
        End Select
    Case Is = UCase("UoDmSfxFr"), UCase("UoDmSfxTo")
        Select Case Opt
        Case Is = Initialise
          .DataType = fdtConvertToUpper
          .Alignment = tLeftAlign
          .Mask = ""
          .MaxLength = 15 '6.1-83
        Case Is = When
          If UCase(IdName) = UCase("UoDmSfxFr") Then HlpList.xTdSfx ADC("UoDmCdFr") ', ADC("UoCmCdFr")
          If UCase(IdName) = UCase("UoDmSfxTo") Then HlpList.xTdSfx ADC("UoDmCdTo") ', ADC("UoCmCdTo")
        End Select
    Case Is = UCase("UoDmSfxSel")
        Select Case Opt
        Case Is = Initialise
          .DataType = fdtConvertToUpper
          .Alignment = tLeftAlign
          .Mask = ""
          .MaxLength = 255
        Case Is = When
          'HlpList.xTdSfx
        End Select
   Case Is = UCase("UoxTTcFr"), UCase("UoxTTcTo")
        Select Case Opt
        Case Is = Initialise
          .DataType = fdtConvertToUpper
          .Alignment = tLeftAlign
          .Mask = ""
          .MaxLength = 3
        Case Is = When
          If Trim(gs_TcTyp) <> "" Then
           HlpList.xPMCd "TC", gs_TcTyp
          Else
           HlpList.xPMCd "TC"
          End If
        End Select
    Case Is = UCase("UoxTYyFr"), UCase("UoxTYyTo")
        Select Case Opt
        Case Is = Initialise
          .DataType = fdtConvertToUpper
          .Alignment = tLeftAlign
          .Mask = ""
          .MaxLength = 2
        Case Is = When
          HlpList.PMCd "YY"
        End Select
    Case Is = UCase("UoxTChrFr"), UCase("UoxTChrTo")
        Select Case Opt
        Case Is = Initialise
          .DataType = fdtConvertToUpper
          .Alignment = tLeftAlign
          .Mask = ""
          .MaxLength = 3
        Case Is = When
          If UCase(IdName) = UCase("UoxTChrFr") Then HlpList.xPSCd "CHR", ADC("UoxTTcFr")
          If UCase(IdName) = UCase("UoxTChrTo") Then HlpList.xPSCd "CHR", ADC("UoxTTcTo")
        End Select
    Case Is = UCase("UoxTNoFr"), UCase("UoxTNoTo")
        Select Case Opt
        Case Is = Initialise
          .DataType = fdtNumber
          .Alignment = tRightAlign
          .Mask = "######0"
          .MaxLength = 7
        Case Is = When
          If UCase(IdName) = UCase("UoxTNoFr") Then HlpList.xTNo ADC("UoxTTcFr"), ADC("UoxTYyFr"), ADC("UoxTChrFr")
          If UCase(IdName) = UCase("UoxTNoTo") Then HlpList.xTNo ADC("UoxTTcTo"), ADC("UoxTYyTo"), ADC("UoxTChrTo")
        End Select
    Case Is = UCase("UoxTdSrFr"), UCase("UoxTdSrTo")
        Select Case Opt
        Case Is = Initialise
          .DataType = fdtNumber
          .Alignment = tRightAlign
          .Mask = "###0"
          .MaxLength = 4
        Case Is = When
          If UCase(IdName) = UCase("UoxTdSrFr") Then HlpList.xTdSr ADC("UoxTTcFr"), ADC("UoxTYyFr"), ADC("UoxTChrFr"), ADC("UoxTNoFr")
          If UCase(IdName) = UCase("UoxTdSrTo") Then HlpList.xTdSr ADC("UoxTTcTo"), ADC("UoxTYyTo"), ADC("UoxTChrTo"), ADC("UoxTNoTo")
        End Select
    Case Is = UCase("UoIdOwnByFr"), UCase("UoIdOwnByTo")
        Select Case Opt
        Case Is = Initialise
          .DataType = fdtConvertToUpper
          .Alignment = tLeftAlign
          .Mask = ""
          .MaxLength = 3
        Case Is = When
          HlpList.xPMCd "IDOWN"
        End Select
    Case Is = UCase("UoIdYyFr"), UCase("UoIdYyTo")
        Select Case Opt
        Case Is = Initialise
          .DataType = fdtConvertToUpper
          .Alignment = tLeftAlign
          .Mask = ""
          .MaxLength = 2
        Case Is = When
          HlpList.PMCd "YY"
        End Select
    Case Is = UCase("UoIdNoFr"), UCase("UoIdNoTo")
        Select Case Opt
        Case Is = Initialise
          .DataType = fdtNumber
          .Alignment = tRightAlign
          .Mask = "######0"
          .MaxLength = 7
        Case Is = When
          If UCase(IdName) = UCase("UoIdNoFr") Then HlpList.xImIdNo ADC("UoIdOwnByFr"), ADC("UoIdYyFr")   ', adc("UoDmCdFr")
          If UCase(IdName) = UCase("UoIdNoTo") Then HlpList.xImIdNo ADC("UoIdOwnByTo"), ADC("UoIdYyTo")   ', adc("UoDmCdTo")
        End Select
    Case Is = UCase("UoDocNoFr"), UCase("UoDocNoTo")
        Select Case Opt
        Case Is = Initialise
          .DataType = fdtConvertToUpper
          .Alignment = tLeftAlign
          .Mask = ""
          .MaxLength = 20
        Case Is = When
          HlpList.xDocNo ADC("UoxCoCdFr")
        End Select
    Case Is = UCase("UoDocDtFr"), UCase("UoDocDtTo")
        Select Case Opt
        Case Is = Initialise
          .DataType = fdtDateType
          .Alignment = tLeftAlign
          .Mask = ""
          .MaxLength = 8
        Case Is = When
          '***
        End Select
    ' Zubin 213 (UoCurCdTo included)
    Case Is = UCase("UoCurCdFr"), UCase("UoCurCdTo")
        Select Case Opt
        Case Is = Initialise
          .DataType = fdtConvertToUpper
          .Alignment = tLeftAlign
          .Mask = ""
          .MaxLength = 5
        Case Is = When
          HlpList.PMCd "Curncy"
        End Select
    ' Zubin 213
    Case Is = UCase("UoCurCdSel")
        Select Case Opt
        Case Is = Initialise
          .DataType = fdtConvertToUpper
          .Alignment = tLeftAlign
          .Mask = ""
          .MaxLength = 255
        Case Is = When
          HlpList.MultiSelect = True
          HlpList.PMCd ("CURNCY")
        End Select
    ' Zubin 213
    Case Is = UCase("UoCurTyp")
        Select Case Opt
        Case Is = Initialise
          .DataType = fdtConvertToUpper
          .Alignment = tLeftAlign
          .Mask = ""
          .MaxLength = 1
        Case Is = When
          HlpList.xPMCd "CurTyp"
        End Select
    Case Is = UCase("UoMatchTyp")
        Select Case Opt
        Case Is = Initialise
          .DataType = fdtConvertToUpper
          .Alignment = tLeftAlign
          .Mask = ""
          .MaxLength = 1
        Case Is = When
          HlpList.xPMCd "MatchTyp"
        End Select
    Case Is = UCase("UoxTAna01Fr"), UCase("UoxTAna01To")
        Select Case Opt
        Case Is = Initialise
          .DataType = fdtConvertToUpper
          .Alignment = tLeftAlign
          .Mask = ""
          .MaxLength = 8
        Case Is = When
          HlpList.xPSCd "AnaCd", "xTAna01"
        End Select
    Case Is = UCase("UoxTAna02Fr"), UCase("UoxTAna02To")
        Select Case Opt
        Case Is = Initialise
          .DataType = fdtConvertToUpper
          .Alignment = tLeftAlign
          .Mask = ""
          .MaxLength = 8
        Case Is = When
          HlpList.xPSCd "AnaCd", "xTAna02"
        End Select
    Case Is = UCase("UoxTAna03Fr"), UCase("UoxTAna03To")
        Select Case Opt
        Case Is = Initialise
          .DataType = fdtConvertToUpper
          .Alignment = tLeftAlign
          .Mask = ""
          .MaxLength = 8
        Case Is = When
          HlpList.xPSCd "AnaCd", "xTAna03"
        End Select
    Case Is = UCase("UoxTAna04Fr"), UCase("UoxTAna04To")
        Select Case Opt
        Case Is = Initialise
          .DataType = fdtConvertToUpper
          .Alignment = tLeftAlign
          .Mask = ""
          .MaxLength = 8
        Case Is = When
          HlpList.xPSCd "AnaCd", "xTAna04"
        End Select
    Case Is = UCase("UoxTAna05Fr"), UCase("UoxTAna05To")
        Select Case Opt
        Case Is = Initialise
          .DataType = fdtConvertToUpper
          .Alignment = tLeftAlign
          .Mask = ""
          .MaxLength = 8
        Case Is = When
          HlpList.xPSCd "AnaCd", "xTAna05"
        End Select
    Case Is = UCase("UoxTAna06Fr"), UCase("UoxTAna06To")
        Select Case Opt
        Case Is = Initialise
          .DataType = fdtConvertToUpper
          .Alignment = tLeftAlign
          .Mask = ""
          .MaxLength = 8
        Case Is = When
          HlpList.xPSCd "AnaCd", "xTAna06"
        End Select
    Case Is = UCase("UoxTAna07Fr"), UCase("UoxTAna07To")
        Select Case Opt
        Case Is = Initialise
          .DataType = fdtConvertToUpper
          .Alignment = tLeftAlign
          .Mask = ""
          .MaxLength = 8
        Case Is = When
          HlpList.xPSCd "AnaCd", "xTAna07"
        End Select
    Case Is = UCase("UoxTAna08Fr"), UCase("UoxTAna08To")
        Select Case Opt
        Case Is = Initialise
          .DataType = fdtConvertToUpper
          .Alignment = tLeftAlign
          .Mask = ""
          .MaxLength = 8
        Case Is = When
          HlpList.xPSCd "AnaCd", "xTAna08"
        End Select
    Case Is = UCase("UoxTdAna01Fr"), UCase("UoxTdAna01To")
        Select Case Opt
        Case Is = Initialise
          .DataType = fdtConvertToUpper
          .Alignment = tLeftAlign
          .Mask = ""
          .MaxLength = 8
        Case Is = When
          HlpList.xPSCd "AnaCd", "xTdAna01"
        End Select
    Case Is = UCase("UoxTdAna02Fr"), UCase("UoxTdAna02To")
        Select Case Opt
        Case Is = Initialise
          .DataType = fdtConvertToUpper
          .Alignment = tLeftAlign
          .Mask = ""
          .MaxLength = 8
        Case Is = When
          HlpList.xPSCd "AnaCd", "xTdAna02"
        End Select
    Case Is = UCase("UoxTdAna03Fr"), UCase("UoxTdAna03To")
        Select Case Opt
        Case Is = Initialise
          .DataType = fdtConvertToUpper
          .Alignment = tLeftAlign
          .Mask = ""
          .MaxLength = 8
        Case Is = When
          HlpList.xPSCd "AnaCd", "xTdAna03"
        End Select
    Case Is = UCase("UoxTdAna04Fr"), UCase("UoxTdAna04To")
        Select Case Opt
        Case Is = Initialise
          .DataType = fdtConvertToUpper
          .Alignment = tLeftAlign
          .Mask = ""
          .MaxLength = 8
        Case Is = When
          HlpList.xPSCd "AnaCd", "xTdAna04"
        End Select
    Case Is = UCase("UoVchTypFr")
        Select Case Opt
        Case Is = Initialise
          .DataType = fdtConvertToUpper
          .Alignment = tLeftAlign
          .Mask = ""
          .MaxLength = 8
        Case Is = When
          HlpList.xPMCd "VchTyp"
        End Select
    Case Is = UCase("UoDmTypFr")
        Select Case Opt
        Case Is = Initialise
          .DataType = fdtConvertToUpper
          .Alignment = tLeftAlign
          .Mask = ""
          .MaxLength = 2
        Case Is = When
          HlpList.PMCd "DmTcTyp"
        End Select
    End Select
  End With
  
  Call SetPropNew(ActiveFrm, IdName, Opt, Cancel, ErrMsg)
  ' ***** Manali 3.5.0 - 14/11/08 - On Compile, VB gave error Procedure too long, thus SetPropNew Procedure called which will have similar functions but for new fields
  
  Set wctl = Nothing
  Set ADC = Nothing
     
End Function

' ****** Manali 3.5.0 - 14/11/08 - SetProp Contd.
Public Function SetPropNew(ByVal ActiveFrm As Form, ByVal IdName As String, ByVal Opt As En_opt, Optional Cancel As Boolean, Optional ErrMsg As String) As Boolean
  
  Dim wctl As MWCTL_MED, ADC As MWCTL_ADC
  Set wctl = ActiveFrm.ADC(IdName)
  Set ADC = ActiveFrm.ADC
  Dim wRmCtg  As String, wRmSCtg  As String, wRmZ As String
  Dim wValidPrdLocSeq() As String, j As Integer, wLocs As String, wValidNonPrdLocSeq() As String  '6.1B
  With wctl
    Select Case UCase(IdName)
    Case Is = UCase("UoCmCdFr"), UCase("UoCmCdTo")
        Select Case Opt
        Case Is = Initialise
          .DataType = fdtConvertToUpper
          .Alignment = tLeftAlign
          .Mask = ""
          .MaxLength = 8
        Case Is = When
          '6.1B-26 valid customer code for the user is added
          Dim wValidCmCd As String
          wValidCmCd = goCn.GetFldVal("Select PDesc225 from Param where PTyp= 'USRCMCD' and PMCd= '" + Trim(gs_UsrCd) + "'")
          
          If gs_CmCtg <> "" Then
            If UCase(ActiveFrm.ADC.MenuCd) = UCase("RepDetOrd") Or UCase(ActiveFrm.ADC.MenuCd) = UCase("RepAvgOrd") Or UCase(ActiveFrm.ADC.MenuCd) = UCase("RepOrdCat") Or _
              UCase(ActiveFrm.ADC.MenuCd) = UCase("RepWIWPrd") Or UCase(ActiveFrm.ADC.MenuCd) = UCase("RepWIWExp") Or _
              UCase(ActiveFrm.ADC.MenuCd) = UCase("RepWIWLongPrd") Or UCase(ActiveFrm.ADC.MenuCd) = UCase("RepWIWLongExp") Then
              HlpList.CustCd gs_CmCtg, , , wValidCmCd
            Else
              HlpList.CustCd gs_CmCtg
            End If
          Else
            If gb_CmCtgFor = True Then
              If UCase(ActiveFrm.ADC.MenuCd) = UCase("RepDetOrd") Or UCase(ActiveFrm.ADC.MenuCd) = UCase("RepAvgOrd") Or UCase(ActiveFrm.ADC.MenuCd) = UCase("RepOrdCat") Or _
                UCase(ActiveFrm.ADC.MenuCd) = UCase("RepWIWPrd") Or UCase(ActiveFrm.ADC.MenuCd) = UCase("RepWIWExp") Or _
                UCase(ActiveFrm.ADC.MenuCd) = UCase("RepWIWLongPrd") Or UCase(ActiveFrm.ADC.MenuCd) = UCase("RepWIWLongExp") Then
                HlpList.CustCd ADC("UoCmCtgFr"), , , wValidCmCd
              Else
                HlpList.CustCd ADC("UoCmCtgFr")
              End If
            Else
'              If adc("UoCmCtgFr") <> "" Or adc("UoCmCtgTo") <> "" Then
'                If UCase(IdName) = "UOCMCDFR" Then HlpList.CustCd adc("UoCmCtgFr")
'                If UCase(IdName) = "UOCMCDTO" Then HlpList.CustCd adc("UoCmCtgTo")
'              Else
                HlpList.CustCd "", , "N"
 '             End If
            End If
          End If
        End Select
    Case Is = UCase("UoCmCdSel")
        Select Case Opt
        Case Is = Initialise
          .DataType = fdtConvertToUpper
          .Alignment = tLeftAlign
          .Mask = ""
          .MaxLength = 255
        Case Is = When
          HlpList.MultiSelect = True
          '6.1B-26 valid customer code for the user is added
          Dim wValidCmCdSel As String
          wValidCmCdSel = goCn.GetFldVal("Select PDesc225 from Param where PTyp= 'USRCMCD' and PMCd= '" + Trim(gs_UsrCd) + "'")
          
          If gs_CmCtg <> "" Then
            If UCase(ActiveFrm.ADC.MenuCd) = UCase("RepWIWPrd") Or UCase(ActiveFrm.ADC.MenuCd) = UCase("RepWIWExp") Or _
              UCase(ActiveFrm.ADC.MenuCd) = UCase("RepWIWLongPrd") Or UCase(ActiveFrm.ADC.MenuCd) = UCase("RepWIWLongExp") Then
              HlpList.CustCd gs_CmCtg, , , wValidCmCdSel
            Else
              HlpList.CustCd gs_CmCtg
            End If
          Else
            If gb_CmCtgFor = True Then
              If UCase(ActiveFrm.ADC.MenuCd) = UCase("RepWIWPrd") Or UCase(ActiveFrm.ADC.MenuCd) = UCase("RepWIWExp") Or _
              UCase(ActiveFrm.ADC.MenuCd) = UCase("RepWIWLongPrd") Or UCase(ActiveFrm.ADC.MenuCd) = UCase("RepWIWLongExp") Then
                HlpList.CustCd ADC("UoCmCtgFr"), , , wValidCmCdSel
              Else
                HlpList.CustCd ADC("UoCmCtgFr")
              End If
            Else
             HlpList.CustCd "", , "N"
            End If
'            HlpList.CustCd adc("UoCmCtgFr")
          End If
        End Select
    Case Is = UCase("UoIssLocFr"), UCase("UoIssLocTo")
        Select Case Opt
        Case Is = Initialise
          .DataType = fdtConvertToUpper
          .Alignment = tLeftAlign
          .Mask = ""
          .MaxLength = 8
        Case Is = When
          j = 0: wLocs = ""   '6.1B
          If Trim(gs_LocTyp) <> "" Then
            If UCase(ActiveFrm.ADC.MenuCd) = UCase("RepWipBPcs") Or UCase(ActiveFrm.ADC.MenuCd) = UCase("RepPrd") Then
              wValidPrdLocSeq = Split(Trim(goCn.GetFldVal("Select PDesc225 from param where PTyp= 'USR' and PMCd= '" + Trim(gs_UsrCd) + "'")), ",")
              wValidNonPrdLocSeq = Split(Trim(goCn.GetFldVal("Select PValue1 from param where PTyp= 'USR' and PMCd= '" + Trim(gs_UsrCd) + "'")), ",")
              If gs_LocTyp = "'P'" Then
                If IsArray(wValidPrdLocSeq) Then
                  For j = 0 To UBound(wValidPrdLocSeq)
                    wLocs = wLocs + IIF(wLocs <> "", ",", "") + "'" + UCase(wValidPrdLocSeq(j)) + "'"
                  Next j
                End If
              Else
                If IsArray(wValidNonPrdLocSeq) Then
                  For j = 0 To UBound(wValidNonPrdLocSeq)
                    wLocs = wLocs + IIF(wLocs <> "", ",", "") + "'" + UCase(wValidNonPrdLocSeq(j)) + "'"
                  Next j
                End If
              End If
            End If
            If gb_CoCdFor = True Then
              HlpList.Loc ADC("UoCoCdFr"), gs_LocTyp, , , , wLocs
            Else
              If UCase(IdName) = UCase("UoIssLocFr") Then HlpList.Loc ADC("UoCoCdFr"), gs_LocTyp, , , , wLocs
              If UCase(IdName) = UCase("UoIssLocTo") Then HlpList.Loc ADC("UoCoCdTo"), gs_LocTyp, , , , wLocs
            End If
          Else
            If gb_CoCdFor = True Then
              HlpList.Loc ADC("UoCoCdFr")
            Else
              If UCase(IdName) = UCase("UoIssLocFr") Then HlpList.Loc ADC("UoCoCdFr")
              If UCase(IdName) = UCase("UoIssLocTo") Then HlpList.Loc ADC("UoCoCdTo")
            End If
          End If
        End Select
    Case Is = UCase("UoIssLocSel")
        Select Case Opt
        Case Is = Initialise
          .DataType = fdtConvertToUpper
          .Alignment = tLeftAlign
          .Mask = ""
          .MaxLength = 255
        Case Is = When
          HlpList.MultiSelect = True
          If Trim(gs_LocTyp) <> "" Then
            j = 0: wLocs = ""   '6.1B
            If UCase(ActiveFrm.ADC.MenuCd) = UCase("RepWipBPcs") Or UCase(ActiveFrm.ADC.MenuCd) = UCase("RepPrd") Then
              wValidPrdLocSeq = Split(Trim(goCn.GetFldVal("Select PDesc225 from param where PTyp= 'USR' and PMCd= '" + Trim(gs_UsrCd) + "'")), ",")
              wValidNonPrdLocSeq = Split(Trim(goCn.GetFldVal("Select PValue1 from param where PTyp= 'USR' and PMCd= '" + Trim(gs_UsrCd) + "'")), ",")
              If gs_LocTyp = "'P'" Then
                If IsArray(wValidPrdLocSeq) Then
                  For j = 0 To UBound(wValidPrdLocSeq)
                    wLocs = wLocs + IIF(wLocs <> "", ",", "") + "'" + UCase(wValidPrdLocSeq(j)) + "'"
                  Next j
                End If
              Else
                If IsArray(wValidNonPrdLocSeq) Then
                  For j = 0 To UBound(wValidNonPrdLocSeq)
                    wLocs = wLocs + IIF(wLocs <> "", ",", "") + "'" + UCase(wValidNonPrdLocSeq(j)) + "'"
                  Next j
                End If
              End If
            End If
            HlpList.Loc ADC("UoCoCdFr"), gs_LocTyp, , , , wLocs
          Else
            HlpList.Loc ADC("UoCoCdFr")
          End If
        End Select
    Case Is = UCase("UoxTcTypFr"), UCase("UoxTcTypTo")
        Select Case Opt
        Case Is = Initialise
          .DataType = fdtConvertToUpper
          .Alignment = tLeftAlign
          .Mask = ""
          .MaxLength = 3
        Case Is = When
          HlpList.xTpTcTyp gs_VchTyp
        End Select
    Case Is = UCase("UoxTcTypSel")
        Select Case Opt
        Case Is = Initialise
          .DataType = fdtConvertToUpper
          .Alignment = tLeftAlign
          .Mask = ""
          .MaxLength = 255
        Case Is = When
          HlpList.MultiSelect = True
          HlpList.xTpTcTyp gs_VchTyp
        End Select
    '****** Sachin 3.02 - Purge Partition (Speed Optimization)
        Case Is = UCase("UoxIssLocFr"), UCase("UoxIssLocTo")
        Select Case Opt
        Case Is = Initialise
          .DataType = fdtConvertToUpper
          .Alignment = tLeftAlign
          .Mask = ""
          .MaxLength = 8
        Case Is = When
          If Trim(gs_LocTyp) <> "" Then
            HlpList.Loc ADC("UoxCoCdFr"), gs_LocTyp
          Else
            HlpList.Loc ADC("UoxCoCdFr")
          End If
        End Select
    Case Is = UCase("UoxIssLocSel")
        Select Case Opt
        Case Is = Initialise
          .DataType = fdtConvertToUpper
          .Alignment = tLeftAlign
          .Mask = ""
          .MaxLength = 255
        Case Is = When
          HlpList.MultiSelect = True
          If Trim(gs_LocTyp) <> "" Then
            HlpList.Loc ADC("UoxCoCdFr"), gs_LocTyp
          Else
            HlpList.Loc ADC("UoxCoCdFr")
          End If
        End Select
    'Zubin xEmr205 ************************************************************
    Case Is = UCase("UoxDc")
        Select Case Opt
        Case Is = Initialise
          .DataType = fdtConvertToUpper
          .Alignment = tLeftAlign
          .Mask = ""
          .MaxLength = 1
        Case Is = When
          HlpList.PMCd "DC"
        End Select
    Case Is = UCase("UoxRepGrp")
        Select Case Opt
          Case Is = Initialise
            .DataType = fdtConvertToUpper
            .Alignment = tLeftAlign
            .Mask = ""
            .MaxLength = 3
          Case Is = When
            HlpList.xPMCd "REPGRP"
        End Select
    Case Is = UCase("UoPrtKeyFr")
        Select Case Opt
        Case Is = Initialise
          .DataType = fdtConvertToUpper
          .Alignment = tLeftAlign
          .Mask = ""
          .MaxLength = 255
        Case Is = When
          Call HlpList.PMCd("PTNTYP")
        End Select
    '****** Sachin 3.02 - Purge Partition (Speed Optimization)
    '****** Sachin 2.13.0 - 30/06/2006 ******
    Case Is = UCase("UoStnClsFr"), UCase("UoStnClsTo")
        Select Case Opt
        Case Is = Initialise
          .DataType = fdtConvertToUpper
          .Alignment = tLeftAlign
          .Mask = ""
          .MaxLength = 8
        Case Is = When
          'jewelsmith.2 default colorstone removed & RmCtg added
          If gb_RmCtgFor = True Then
            HlpList.PSCd "STNCLS", ADC("UoRmCtgFr")
          Else
            If UCase(IdName) = "UOSTNCLSFR" Then HlpList.PSCd "STNCLS", ADC("UoRmCtgFr")
            If UCase(IdName) = "UOSTNCLSTO" Then HlpList.PSCd "STNCLS", ADC("UoRmCtgTo")
          End If
        End Select
    Case Is = UCase("UoStnClsSel")
        Select Case Opt
        Case Is = Initialise
          .DataType = fdtConvertToUpper
          .Alignment = tLeftAlign
          .Mask = ""
          .MaxLength = 255
        Case Is = When
          HlpList.MultiSelect = True
          'jewelsmith.2 default colorstone removed & RmCtg added
          HlpList.PSCd "STNCLS", ADC("UoRmCtgFr")
        End Select
     Case Is = UCase("UoOdDmIdFr"), UCase("UoOdDmIdTo")
        Select Case Opt
        Case Is = Initialise
          .DataType = fdtNumber
          .Alignment = tRightAlign
          .Mask = "######0"
          .MaxLength = 7
        Case Is = When
          ' ***** Manali Trading Module - [DmSz] added
          If UCase(IdName) = UCase("UoOdDmIdFr") Then HlpList.OdDmStkNo gs_CoCd, ADC("UoDmCdFr"), ADC("UoOdSfxFr"), ADC("UoDmSzFr"), ADC("UOSTKCOCDFR"), ADC("UOBYyFr"), ADC("UoBChrFr")
          If UCase(IdName) = UCase("UoOdDmIdTo") Then HlpList.OdDmStkNo gs_CoCd, ADC("UoDmCdTo"), ADC("UoOdSfxTo"), ADC("UoDmSzTo"), ADC("UOSTKCOCDTO"), ADC("UOBYyTo"), ADC("UoBChrTo")
        End Select
    '*** Jay 3.3(FG)
    
      Case Is = UCase("UoRmIdFr"), UCase("UoRmIdTo")
          Select Case Opt
          Case Is = Initialise
            .DataType = fdtConvertToUpper
            .Alignment = tLeftAlign
            .Mask = ""
            .MaxLength = 5
          Case Is = When
            If UCase(IdName) = UCase("UoRmIdFr") Then HlpList.PSCd "RMID", ADC("UoRmCtgFr")
            If UCase(IdName) = UCase("UoRmIdTo") Then HlpList.PSCd "RMID", ADC("UoRmCtgTo")
          End Select
      Case Is = UCase("UoRmIdSel")
          Select Case Opt
          Case Is = Initialise
            .DataType = fdtConvertToUpper
            .Alignment = tLeftAlign
            .Mask = ""
            .MaxLength = 255
          Case Is = When
            HlpList.MultiSelect = True
            HlpList.PSCd "RMID", ADC("UoRmCtgSel")
          End Select
      Case Is = UCase("UOREGNCDFR"), UCase("UOREGNCDTO")
          Select Case Opt
          Case Is = Initialise
            .DataType = fdtConvertToUpper
            .Alignment = tLeftAlign
            .Mask = ""
            .MaxLength = 5
          Case Is = When
            HlpList.PMCd "REGION"
          End Select
      Case Is = UCase("UOREGNCDSEL")
          Select Case Opt
          Case Is = Initialise
            .DataType = fdtConvertToUpper
            .Alignment = tLeftAlign
            .Mask = ""
            .MaxLength = 255
          Case Is = When
            HlpList.MultiSelect = True
            HlpList.PMCd "REGION"
          End Select
  ' ****** Manali 3.5.0 - 08/12/08 - DsgPrm fields
      Case Is = UCase("UoDpCatFr"), UCase("UoDpCatTo")
          Select Case Opt
          Case Is = Initialise
            .DataType = fdtConvertToUpper
            .Alignment = tLeftAlign
            .Mask = ""
            .MaxLength = 16
          Case Is = When
            HlpList.DsgPrmCd ("CAT")
          End Select
      Case Is = UCase("UoDpCatSel")
          Select Case Opt
          Case Is = Initialise
            .DataType = fdtConvertToUpper
            .Alignment = tLeftAlign
            .Mask = ""
            .MaxLength = 255
          Case Is = When
            HlpList.MultiSelect = True
            HlpList.DsgPrmCd ("CAT")
          End Select
      Case Is = UCase("UoDpCollFr"), UCase("UoDpCollTo")
          Select Case Opt
          Case Is = Initialise
            .DataType = fdtConvertToUpper
            .Alignment = tLeftAlign
            .Mask = ""
            .MaxLength = 16
          Case Is = When
            HlpList.DsgPrmCd ("COLL")
          End Select
      Case Is = UCase("UoDpCollSel")
          Select Case Opt
          Case Is = Initialise
            .DataType = fdtConvertToUpper
            .Alignment = tLeftAlign
            .Mask = ""
            .MaxLength = 255
          Case Is = When
            HlpList.MultiSelect = True
            HlpList.DsgPrmCd ("COLL")
          End Select
      Case Is = UCase("UoDpSerFr"), UCase("UoDpSerTo")
          Select Case Opt
          Case Is = Initialise
            .DataType = fdtNumber
            .Alignment = tRightAlign
            .Mask = "####0"
            .MaxLength = 5
          Case Is = When
            Call HlpList.DsgPrmCd("SER", True)
          End Select
      Case Is = UCase("UoDpSerSel")
          Select Case Opt
          Case Is = Initialise
            .DataType = fdtConvertToUpper
            .Alignment = tLeftAlign
            .Mask = ""
            .MaxLength = 255
          Case Is = When
            HlpList.MultiSelect = True
            Call HlpList.DsgPrmCd("SER", True)
          End Select
      Case Is = UCase("UoDpSetFr"), UCase("UoDpSetTo")
          Select Case Opt
          Case Is = Initialise
            .DataType = fdtNumber
            .Alignment = tRightAlign
            .Mask = "####0"
            .MaxLength = 5
          Case Is = When
            Call HlpList.DsgPrmCd("SET", True)
          End Select
      Case Is = UCase("UoDpSetSel")
          Select Case Opt
          Case Is = Initialise
            .DataType = fdtConvertToUpper
            .Alignment = tLeftAlign
            .Mask = ""
            .MaxLength = 255
          Case Is = When
            HlpList.MultiSelect = True
            Call HlpList.DsgPrmCd("SET", True)
          End Select
      Case Is = UCase("UoDpVerFr"), UCase("UoDpVerTo")
          Select Case Opt
          Case Is = Initialise
            .DataType = fdtNumber
            .Alignment = tRightAlign
            .Mask = "####0"
            .MaxLength = 5
          Case Is = When
            Call HlpList.DsgPrmCd("VER", True)
          End Select
      Case Is = UCase("UoDpVerSel")
          Select Case Opt
          Case Is = Initialise
            .DataType = fdtConvertToUpper
            .Alignment = tLeftAlign
            .Mask = ""
            .MaxLength = 255
          Case Is = When
            HlpList.MultiSelect = True
            Call HlpList.DsgPrmCd("VER", True)
          End Select
  ' ****** Manali 3.5.0 - 08/12/08 - DsgPrm fields
  ' ****** Manali 350Nxt - DSGEXC
        Case Is = UCase("UODSGEXC")
          Select Case Opt
          Case Is = Initialise
            .DataType = fdtConvertToUpper
            .Alignment = tLeftAlign
            .Mask = ""
            .MaxLength = 1
          Case Is = When
            HlpList.PMCd "DSGEXC"
          End Select
  ' ****** Manali 350Nxt - DSGEXC
  ' **** Manali 3.6.0 - 08/09/09 - IdSr Level Scope added to Inv Costing Reports
        Case Is = UCase("UoIdSrFr"), UCase("UoIdSrTo")
          Select Case Opt
          Case Is = Initialise
            .DataType = fdtNumber
            .Alignment = tRightAlign
            .Mask = "###0"
            .MaxLength = 4
          Case Is = When
            If UCase(IdName) = UCase("UoIdSrFr") Then HlpList.InSr ADC("UoCoCdFr"), ADC("UoInTcFr"), ADC("UoInYyFr"), ADC("UoInChrFr"), ADC("UoInNoFr")
            If gb_CoCdFor = False Then
              If UCase(IdName) = UCase("UoIdSrTo") Then HlpList.InSr ADC("UoCoCdTo"), ADC("UoInTcTo"), ADC("UoInYyTo"), ADC("UoInChrTo"), ADC("UoInNoTo")
            Else
              If UCase(IdName) = UCase("UoIdSrTo") Then HlpList.InSr ADC("UoCoCdFr"), ADC("UoInTcTo"), ADC("UoInYyTo"), ADC("UoInChrTo"), ADC("UoInNoTo")
            End If
          End Select
  ' **** Manali 3.6.0 - 08/09/09 - IdSr Level Scope added to Inv Costing Reports
  ' ***** Manali Trading Module
        Case Is = UCase("UoOmTc"), UCase("UoOmTc1"), UCase("UoOmTc2")
            Select Case Opt
            Case Is = Initialise
              .DataType = fdtConvertToUpper
              .Alignment = tLeftAlign
              .Mask = ""
              .MaxLength = 3
            Case Is = When
              'HlpList.PMCd "TC", "'" + gs_OmTcTyp + "'"
              'If gb_CmCtgFor Then
              ' ***** Manali 3.8.0 - For Help Of UoOmTc
              If gs_CmCtg1 <> "" Then
                HlpList.TmTc ADC("UoCoCdFr"), gs_CmCtg1
              Else
                HlpList.TmTc ADC("UoCoCdFr")
              End If
            End Select
  ' ***** Manali Trading Module
        Case Is = UCase("UoSbICC"), UCase("UoSbCOD")
            Select Case Opt
            Case Is = Initialise
              .DataType = fdtConvertToUpper
              .Alignment = tLeftAlign
              .Mask = ""
              .MaxLength = 3
            Case Is = When
               HlpList.PMCd "CONTRYCD"
            End Select
        Case Is = UCase("UoSbPOD")
            Select Case Opt
            Case Is = Initialise
              .DataType = fdtConvertToUpper
              .Alignment = tLeftAlign
              .Mask = ""
              .MaxLength = 6
            Case Is = When
               HlpList.PMCd "PORTCD"
            End Select
        Case Is = UCase("UoSbAF")
            Select Case Opt
            Case Is = Initialise
              .DataType = fdtConvertToUpper
              .Alignment = tLeftAlign
              .Mask = ""
              .MaxLength = 1
            Case Is = When
               HlpList.PMCd "AF"
            End Select
        Case Is = UCase("UoSbNOC")
            Select Case Opt
            Case Is = Initialise
              .DataType = fdtConvertToUpper
              .Alignment = tLeftAlign
              .Mask = ""
              .MaxLength = 1
            Case Is = When
               HlpList.PMCd "INVTOTAS"
            End Select

        ' **** Manali 3.9.1 - 23/02/12 - RfId Module
        Case Is = UCase("UoRimNoFr"), UCase("UoRimNoTo")
          Select Case Opt
          Case Is = Initialise
            .DataType = fdtNumber
            .Alignment = tRightAlign
            .Mask = "#####0"
            .MaxLength = 6
          Case Is = When
             HlpList.RimNo
          End Select
        ' **** Manali 3.9.1 - 23/02/12 - RfId Module
    
    Case Is = UCase("UoMqmTcFr"), UCase("UoMqmTcTo")
        Select Case Opt
        Case Is = Initialise
          .DataType = fdtConvertToUpper
          .Alignment = tLeftAlign
          .Mask = ""
          .MaxLength = 3
        Case Is = When
          HlpList.PMCd "TC", "'QT'"
        End Select
    Case Is = UCase("UoMqmYyFr"), UCase("UoMqmYyTo")
        Select Case Opt
        Case Is = Initialise
          .DataType = fdtConvertToUpper
          .Alignment = tLeftAlign
          .Mask = ""
          .MaxLength = 2
        Case Is = When
          HlpList.PMCd "YY"
        End Select
    Case Is = UCase("UoMqmChrFr"), UCase("UoMqmChrTo")
        Select Case Opt
        Case Is = Initialise
          .DataType = fdtConvertToUpper
          .Alignment = tLeftAlign
          .Mask = ""
          .MaxLength = 3
        Case Is = When
          If UCase(IdName) = UCase("UoMqmChrFr") Then HlpList.vPSCd ADC("UoCoCdFr"), "CHR", ADC("UoMqmTcFr")
          If gb_CoCdFor = False Then
            If UCase(IdName) = UCase("UoMqmChrTo") Then HlpList.vPSCd ADC("UoCoCdTo"), "CHR", ADC("UoMqmTcTo")
          Else
            If UCase(IdName) = UCase("UoMqmChrTo") Then HlpList.vPSCd ADC("UoCoCdFr"), "CHR", ADC("UoMqmTcTo")
          End If
        End Select
    Case Is = UCase("UoMqmNoFr"), UCase("UoMqmNoTo")
        Select Case Opt
        Case Is = Initialise
          .DataType = fdtNumber
          .Alignment = tRightAlign
          .Mask = "######0"
          .MaxLength = 7
        Case Is = When
          If UCase(IdName) = UCase("UoMqmNoFr") Then HlpList.MultiPrcQtNo ADC("UoCoCdFr"), ADC("UoMqmTcFr"), ADC("UoMqmYyFr"), ADC("UoMqmChrFr")
          If gb_CoCdFor = False Then
            If UCase(IdName) = UCase("UoMqmNoTo") Then HlpList.MultiPrcQtNo ADC("UoCoCdTo"), ADC("UoMqmTcTo"), ADC("UoMqmYyTo"), ADC("UoMqmChrTo")
          Else
            If UCase(IdName) = UCase("UoMqmNoTo") Then HlpList.MultiPrcQtNo ADC("UoCoCdFr"), ADC("UoMqmTcTo"), ADC("UoMqmYyTo"), ADC("UoMqmChrTo")
          End If
        End Select
    Case Is = UCase("UoMqdSrFr"), UCase("UoMqdSrTo")
        Select Case Opt
        Case Is = Initialise
          .DataType = fdtNumber
          .Alignment = tRightAlign
          .Mask = "####0"
          .MaxLength = 5
        Case Is = When
          If UCase(IdName) = UCase("UoMqdSrFr") Then HlpList.MultiPrcQtSr ADC("UoCoCdFr"), ADC("UoMqmTcFr"), ADC("UoMqmYyFr"), ADC("UoMqmChrFr"), ADC("UoMqmNoFr")
          If gb_CoCdFor = False Then
            If UCase(IdName) = UCase("UoMqdSrTo") Then HlpList.MultiPrcQtSr ADC("UoCoCdTo"), ADC("UoMqmTcTo"), ADC("UoMqmYyTo"), ADC("UoMqmChrTo"), ADC("UoMqmNoTo")
          Else
            If UCase(IdName) = UCase("UoMqdSrTo") Then HlpList.MultiPrcQtSr ADC("UoCoCdFr"), ADC("UoMqmTcTo"), ADC("UoMqmYyTo"), ADC("UoMqmChrTo"), ADC("UoMqmNoTo")
          End If
        End Select
    Case Is = UCase("UoMetGrdFr"), UCase("UoMetGrdTo")
        Select Case Opt
        Case Is = Initialise
          .DataType = fdtConvertToUpper
          .Alignment = tLeftAlign
          .Mask = ""
          .MaxLength = 8
        Case Is = When
          HlpList.PSCd "GRDCD", "MET"
        End Select
    Case Is = UCase("UoMetGrdSel")
        Select Case Opt
        Case Is = Initialise
          .DataType = fdtConvertToUpper
          .Alignment = tLeftAlign
          .Mask = ""
          .MaxLength = 255
        Case Is = When
          HlpList.MultiSelect = True
          HlpList.PSCd "GRDCD", "MET"
        End Select
    Case Is = UCase("UoDiaGrdFr"), UCase("UoDiaGrdTo")
        Select Case Opt
        Case Is = Initialise
          .DataType = fdtConvertToUpper
          .Alignment = tLeftAlign
          .Mask = ""
          .MaxLength = 8
        Case Is = When
          HlpList.PSCd "GRDCD", "DIA"
        End Select
    Case Is = UCase("UoDiaGrdSel")
        Select Case Opt
        Case Is = Initialise
          .DataType = fdtConvertToUpper
          .Alignment = tLeftAlign
          .Mask = ""
          .MaxLength = 255
        Case Is = When
          HlpList.MultiSelect = True
          HlpList.PSCd "GRDCD", "DIA"
        End Select
    Case Is = UCase("UoVaCtgFr"), UCase("UoVaCtgTo")  '3.11.2
        Select Case Opt
        Case Is = Initialise
          .DataType = fdtConvertToUpper
          .Alignment = tLeftAlign
          .Mask = ""
          .MaxLength = 5
        Case Is = When
          HlpList.PMCd "VACTG", ""                   '###
        End Select
    Case Is = UCase("UoVaCtgSel")                   '3.11.2
        Select Case Opt
        Case Is = Initialise
          .DataType = fdtConvertToUpper
          .Alignment = tLeftAlign
          .Mask = ""
          .MaxLength = 255
        Case Is = When
          HlpList.MultiSelect = True
          HlpList.PMCd "VACTG", ""
        End Select
  
    Case Is = UCase("UoCmGrpFr"), UCase("UoCmGrpTo")  'Sachin - 4.1.0.0
        Select Case Opt
        Case Is = Initialise
          .DataType = fdtConvertToUpper
          .Alignment = tLeftAlign
          .Mask = ""
          .MaxLength = 8
        Case Is = When
          HlpList.PMCd "CMGRP", ""
        End Select
    Case Is = UCase("UoCmGrpSel")
        Select Case Opt
        Case Is = Initialise
          .DataType = fdtConvertToUpper
          .Alignment = tLeftAlign
          .Mask = ""
          .MaxLength = 255
        Case Is = When
          HlpList.MultiSelect = True
          HlpList.PMCd "CMGRP", ""
        End Select
  
    Case Is = UCase("UoCmTypFr"), UCase("UoCmTypTo")  'Sachin - 4.1.0.0
        Select Case Opt
        Case Is = Initialise
          .DataType = fdtConvertToUpper
          .Alignment = tLeftAlign
          .Mask = ""
          .MaxLength = 8
        Case Is = When
          HlpList.PMCd "CMTYP", ""
        End Select
    Case Is = UCase("UoCmTypSel")
        Select Case Opt
        Case Is = Initialise
          .DataType = fdtConvertToUpper
          .Alignment = tLeftAlign
          .Mask = ""
          .MaxLength = 255
        Case Is = When
          HlpList.MultiSelect = True
          HlpList.PMCd "CMTYP", ""
        End Select
  
'***********************Bhavna memo return reason
    Case Is = UCase("UoRetRsnFr"), UCase("UoRetRsnTo")
        Select Case Opt
        Case Is = Initialise
          .DataType = fdtConvertToUpper
          .Alignment = tLeftAlign
          .Mask = ""
          .MaxLength = 2
        Case Is = When
          HlpList.PMCd "RETRSN"
        End Select
    Case Is = UCase("UoRetRsnSel")
        Select Case Opt
        Case Is = Initialise
          .DataType = fdtConvertToUpper
          .Alignment = tLeftAlign
          .Mask = ""
          .MaxLength = 255
        Case Is = When
          HlpList.MultiSelect = True
          HlpList.PMCd "RETRSN"
        End Select

'***********************Bhavna rework reason
    Case Is = UCase("UoRwrkFr"), UCase("UoRwrkTo")
        Select Case Opt
        Case Is = Initialise
          .DataType = fdtConvertToUpper
          .Alignment = tLeftAlign
          .Mask = ""
          .MaxLength = 8
        Case Is = When
         
          If UCase(IdName) = "UORWRKFR" Then HlpList.PSCd "RWRK", ADC("UoIssLocFr")
          If UCase(IdName) = "UORWRKTO" Then HlpList.PSCd "RWRK", ADC("UoIssLocTo")
        End Select
    Case Is = UCase("UoRwrkSel")
        Select Case Opt
        Case Is = Initialise
          .DataType = fdtConvertToUpper
          .Alignment = tLeftAlign
          .Mask = ""
          .MaxLength = 255
        Case Is = When
          HlpList.MultiSelect = True
          HlpList.PSCd "RWRK", ADC("UoIssLocFr")
          
        End Select
    Case Is = UCase("UoStkCoCdFr"), UCase("UoStkCoCdTo")
        Select Case Opt
        Case Is = Initialise
          .DataType = fdtConvertToUpper
          .Alignment = tLeftAlign
          .Mask = ""
          .MaxLength = 3
        Case Is = When
          HlpList.hCoCd
        End Select
  'Thiru
    Case Is = UCase("UoCdCtg")
       Select Case Opt
       Case Is = Initialise
         .DataType = fdtConvertToUpper
         .Alignment = tLeftAlign
         .Mask = ""
         .MaxLength = 1
       Case Is = When
         HlpList.PMCd "CDCTG"
       End Select
    'vk.19 Sub customer From & To added
    Case Is = UCase("UoOdSubCustFr"), UCase("UoOdSubCustTo")
        Select Case Opt
        Case Is = Initialise
          .DataType = fdtConvertToUpper
          .Alignment = tLeftAlign
          .Mask = ""
          .MaxLength = 8
        Case Is = When
          HlpList.PMCd "SUBCUST"
        End Select
    ' MW.107 - Help added for OdCellFr, OdCellTo and OdCellSel for Shipment Analysis Report
    Case Is = UCase("UoOdCellFr"), UCase("UoOdCellTo")
          Select Case Opt
          Case Is = Initialise
            .DataType = fdtConvertToUpper
            .Alignment = tLeftAlign
            .Mask = ""
            .MaxLength = 8
          Case Is = When
            HlpList.PMCd "CELL"
          End Select
    Case Is = UCase("UoOdCellSel")
          Select Case Opt
          Case Is = Initialise
            .DataType = fdtConvertToUpper
            .Alignment = tLeftAlign
            .Mask = ""
            .MaxLength = 255
          Case Is = When
            HlpList.MultiSelect = True
            HlpList.PMCd "CELL"
          End Select
    Case Is = UCase("UoInGrp")
        Select Case Opt
        Case Is = Initialise
          .DataType = fdtConvertToUpper
          .Alignment = tLeftAlign
          .Mask = ""
          .MaxLength = 8
        Case Is = When
          HlpList.PMCd "INGR", , , " (PValue='' or ','+PValue+',' like ('%," + ActiveFrm.ADC.MenuCd + ",%')) "
        End Select
    'sv.53 PValue filter added to show OmPo option only to some reports
    Case Is = UCase("UoInGrp1")
        Select Case Opt
        Case Is = Initialise
          .DataType = fdtConvertToUpper
          .Alignment = tLeftAlign
          .Mask = ""
          .MaxLength = 8
        Case Is = When
          HlpList.PMCd "INGR", , , " (PValue='' or ','+PValue+',' like ('%," + ActiveFrm.ADC.MenuCd + ",%')) "
        End Select
 End Select
End With
  
  Set wctl = Nothing
  Set ADC = Nothing

End Function


'SRI -> THIS FUNCTION IS NEVER USED
Public Function zGetRMRt(ByVal mCmCd As String, ByVal mRmCd As String, _
              ByVal mLn1 As Single, ByVal mLn2 As Single, ByVal mWtPerDia As Single, _
              ByVal mOpt As String, Optional ByVal mHistCmCtg As String, _
              Optional ByVal mHistCmCd As String, _
              Optional ByVal mHistYyyyMm As Double) As Single
              
  Dim wRmCtg As String, wRmSCtg As String, wRtBy As String, wRt As String, wCnd As String
  Dim wInc As Double
  Dim wRtChrtByPtr As String, wCmLkUpRmRt As String, wRtChrtByBs As String
  '*************************** ZUBIN **************************
  ' 05th Dec 2003, EMR206
  Dim wFrLn As String, wToLn As String
  
  If (mOpt <> "C") And (mOpt <> "S") Then zGetRMRt = 0: Exit Function
    
  '*************************** ZUBIN **************************
   If (mHistCmCtg <> "" And (mHistYyyyMm = 0 Or mHistCmCd = "")) _
      Or (mHistCmCd <> "" And (mHistYyyyMm = 0 Or mHistCmCtg = "")) _
      Or (mHistYyyyMm <> 0 And (mHistCmCtg = "" Or mHistCmCd = "")) Then
    DispMsg "Proper Rate From History Cust Ctg, Cust Cd and YyyyMm Combination should be passed", etError
    Exit Function
  End If
  '*************************** ZUBIN **************************
  
  ' **** Manali 3.03 -18/06/08 - RmCtg ="S" and "L" included
  wRmCtg = goCn.GetFldVal("Select RmCtg from RmMst where RMCd='" + mRmCd + "'" + IIF(gs_Partition = ctCurrPrtn, " and RmPrtKey='" + ctCurrPrtn + "' ", ""))
  If wRmCtg = "G" Or wRmCtg = "P" Or wRmCtg = "S" Or wRmCtg = "L" Then zGetRMRt = 0: Exit Function

  wRmSCtg = goCn.GetFldVal("Select RmSCtg from RmMst where RmCd='" + mRmCd + "'" + IIF(gs_Partition = ctCurrPrtn, " and RmPrtKey='" + ctCurrPrtn + "' ", ""))
  wRtChrtByPtr = ""
  wRtChrtByBs = ""
  If mHistCmCtg = vbNullString Or mHistCmCtg <> "P" Then
  wRtChrtByPtr = goCn.GetFldVal("Select 'P' from RmRt where RrTcTyp= 'PTR' and RrCmCtg= 'C' " + _
                 "And RrCmCd = '" + mCmCd + "' and RrCtg= '" + wRmCtg + "' and " + _
                 "RrSCtg= '" + wRmSCtg + "' and RrSTWGrp= '' and RrCd= '' and RrDmCtg= '' " + _
                 "and RrLsCtg= '' and RrLabMCd= '' and RrFrLn= 0 and RrToLn= 0 " + IIF(gs_Partition = ctCurrPrtn, " and RrPrtKey='" + ctCurrPrtn + "' ", ""))
  
  '***************Geeta*************Emr206*********05/04/04
  wRtChrtByBs = goCn.GetFldVal("Select 'BS' from RmRt where RrTcTyp= 'BS' and RrCmCtg= 'C' " + _
                 "And RrCmCd = '" + mCmCd + "' and RrCtg= '" + wRmCtg + "' and " + _
                 "RrSCtg= '" + wRmSCtg + "' and RrSTWGrp= '' and RrCd= '' and RrDmCtg= '' " + _
                 "and RrLsCtg= '' and RrLabMCd= '' and RrFrLn= 0 and RrToLn= 0 " + IIF(gs_Partition = ctCurrPrtn, " and RrPrtKey='" + ctCurrPrtn + "' ", ""))
 End If
  '*****************
  wRtBy = goCn.GetFldVal("Select PValue from Param where PTyp= 'RMSCTG' and " + _
          "PMCd= '" + wRmCtg + "' and PSCd= '" + wRmSCtg + "'")
          
  If wRtChrtByPtr <> "" Then
    wRtBy = wRtChrtByPtr
 '***************Geeta*************Emr206*********05/04/04
  ElseIf wRtChrtByBs <> "" Then
    wRtBy = wRtChrtByBs
  End If
  
  '*************************** ZUBIN **************************
  If mHistCmCtg = vbNullString Then
    wRt = IIF(mOpt = "S", "RrSalRt", "RrCstRt")
  Else
    wRt = IIF(mOpt = "S", "RhSalRt", "RhCstRt")
  End If
  '*************************** ZUBIN **************************
  
  '********** Increments, only if mOpt='S' & Self Record *********
  If mOpt = "S" Then
    wInc = goCn.GetFldVal("Select RrSalRt from RmRt where RrTcTyp= 'INC' and RrCmCtg= 'C' and " + _
           "RrCtg= '" + wRmCtg + "' and RrSCtg= '" + wRmSCtg + "' and RrCmCd= '" + mCmCd + "'" + IIF(gs_Partition = ctCurrPrtn, " and RrPrtKey='" + ctCurrPrtn + "' ", ""))
  Else
    wInc = 0
  End If

  '****** 'S'  - By Size     (condition is RrFrLn <=  Ln1  <= RrToLn)
  '****** 'P'  - By Pointer  (condition is RrFrLn <=  WtPerDia  <= RrToLn)
  '****** 'BS' - Both Size   (condition is RrFrLn = Ln1  and  RrToLn = Ln2)
  
  '*** By Size      - is used by diamonds i.e. only Len is used to calculate the rate
  '*** By Pointer   - is used by Pointers i.e. Wt Per Diamond is used to calculate the rate
  '*** By Both Size - is used by Col Stones i.e. Both Length and Breadth are used to calculate the rate
  
  '*************************** ZUBIN **************************
  'Dim wFrLn As String, wToLn As String
  
  If mHistCmCtg = vbNullString Then
    wFrLn = "RrFrLn": wToLn = "RrToLn"
  Else
    wFrLn = "RhFrLn": wToLn = "RhToLn"
  End If
  
  Select Case wRtBy
    Case Is = "S"
      wCnd = " Round(" + wFrLn + ", 4) <= Round(" + CStr(mLn1) + ", 3) And Round(" + wToLn + ", 4) >= Round(" + CStr(mLn1) + ", 3) and "
    Case Is = "W", "P"
      wCnd = " Round(" + wFrLn + ", 4) <= Round(" + CStr(mWtPerDia) + ", 4) And Round(" + wToLn + ", 4) >= Round(" + CStr(mWtPerDia) + ", 4) and "
    Case Is = "BS"
      wCnd = " Round(" + wFrLn + ", 4) = Round(" + CStr(mLn1) + ", 3) And Round(" + wToLn + ", 4) = Round(" + CStr(mLn2) + ", 3) and "
  End Select
  
'  Select Case wRtBy
'  Case Is = "S"
'    wCnd = " RrFrLn <= " + CStr(mLn1) + " And RrToLn >= " + CStr(mLn1) + " and "
'  Case Is = "W", "P"
'    wCnd = " RrFrLn <= " + CStr(mWtPerDia) + " And RrToLn >= " + CStr(mWtPerDia) + " and "
'  Case Is = "BS"
'    wCnd = " RrFrLn = " + CStr(mLn1) + " And RrToLn = " + CStr(mLn2) + " and "
'  End Select

  wCmLkUpRmRt = goCn.GetFldVal("Select CmLkUpRmRt from CustMst where CmCtg= 'C' and CmCd= '" + mCmCd + "'")
  '********** 1st look for sale rate from PLCust, then from Cust record;
  '********** if not found then from Self (applying increments for the customer)
  
  'CHECK IF mHistCmCtg IsMissing().
  'IF mHistCmCtg IsMissing() Is True then Let the Rates come from RmRt (As before)
  
  If wRtChrtByPtr = "" And wRtChrtByBs = "" Then
    If mHistCmCtg <> vbNullString Then
      zGetRMRt = goCn.GetFldVal("Select " + _
              "isnull((Select " + wRt + " from RmRtHist where " + _
                      "RhYyyyMm= " + CStr(mHistYyyyMm) + " and RhTcTyp='RM' and " + _
                      "RhCtg= '" + wRmCtg + "' and RhSCtg= '" + wRmSCtg + "' and " + _
                      "RhCd='" + mRmCd + "' and " + wCnd + " RhCmCtg= '" + mHistCmCtg + "' and " + _
                      "RhCmCd='" + mHistCmCd + "' and " + wRt + " <> 0" + IIF(gs_Partition = ctCurrPrtn, " and RrPrtKey='" + ctCurrPrtn + "' ", "") + ")" + _
              ", " + _
              "IsNull((Select " + wRt + " from RmRtHist " + _
                      "where RhYyyyMm= " + CStr(mHistYyyyMm) + " and RhTcTyp= 'RM' and " + _
                      "RhCtg= '" + wRmCtg + "' and RhSCtg= '" + wRmSCtg + "' and " + _
                      "RhCd= '" + mRmCd + "' and " + wCnd + " RhCmCtg= 'C' and RhCmCd = '" + ctSelfCmCd + "' and " + _
                      wRt + " <> 0" + IIF(gs_Partition = ctCurrPrtn, " and RrPrtKey='" + ctCurrPrtn + "' ", "") + ")*" + CStr(1 + wInc / 100) + ",0)" + _
                      ")")
    Else
      zGetRMRt = goCn.GetFldVal("Select " + _
             "isnull((Select " + wRt + " from RmRt where RrTcTyp='RM' and " + _
                     "RrCtg= '" + wRmCtg + "' and RrSCtg= '" + wRmSCtg + "' and " + _
                     "RrCd='" + mRmCd + "' and " + wCnd + " RrCmCtg= 'P' and RrCmCd = '" + wCmLkUpRmRt + "' and " + _
                      wRt + " <> 0" + IIF(gs_Partition = ctCurrPrtn, " and RrPrtKey='" + ctCurrPrtn + "' ", "") + ")" + _
                    ", " + _
                    "isnull((Select " + wRt + " from RmRt where RrTcTyp='RM' and " + _
                            "RrCtg= '" + wRmCtg + "' and RrSCtg= '" + wRmSCtg + "' and " + _
                            "RrCd='" + mRmCd + "' and " + wCnd + " RrCmCtg= 'C' and RrCmCd = '" + mCmCd + "' and " + _
                             wRt + " <> 0" + IIF(gs_Partition = ctCurrPrtn, " and RrPrtKey='" + ctCurrPrtn + "' ", "") + ")" + _
                           ", " + _
                           "IsNull((Select " + wRt + " from RmRt where RrTcTyp= 'RM' and " + _
                                    "RrCtg= '" + wRmCtg + "' and RrSCtg= '" + wRmSCtg + "' and " + _
                                    "RrCd= '" + mRmCd + "' and " + wCnd + " RrCmCtg= 'C' and RrCmCd = '" + ctSelfCmCd + "' and " + _
                                    wRt + " <> 0" + IIF(gs_Partition = ctCurrPrtn, " and RrPrtKey='" + ctCurrPrtn + "' ", "") + ")*" + CStr(1 + wInc / 100) + ",0)" + _
                    ")" + _
             ")")
    End If
  ElseIf UCase(wRtChrtByPtr) = "P" Or UCase(wRtChrtByBs) = "BS" Then
    If mHistCmCtg <> vbNullString Then
      zGetRMRt = goCn.GetFldVal("Select " + wRt + " from RmRtHist " + _
                "where RhYyyyMm= " + CStr(mHistYyyyMm) + " and RhTcTyp= 'RM' and " + _
                "RhCtg= '" + wRmCtg + "' and RhSCtg= '" + wRmSCtg + "' and " + _
                "RhCd='" + mRmCd + "' and " + wCnd + " RhCmCtg= 'C' and RhCmCd = '" + _
                mCmCd + "' and " + wRt + " <> 0" + IIF(gs_Partition = ctCurrPrtn, " and RrPrtKey='" + ctCurrPrtn + "' ", ""))
    Else
      zGetRMRt = goCn.GetFldVal("Select " + wRt + " from RmRt where RrTcTyp= 'RM' and " + _
                "RrCtg= '" + wRmCtg + "' and RrSCtg= '" + wRmSCtg + "' and " + _
                "RrCd='" + mRmCd + "' and " + wCnd + " RrCmCtg= 'C' and RrCmCd = '" + mCmCd + "' and " + _
                wRt + " <> 0" + IIF(gs_Partition = ctCurrPrtn, " and RrPrtKey='" + ctCurrPrtn + "' ", ""))
    End If
  End If

End Function
'Print SRITMP
'Public Function CoNm() As String
'  CoNm = goCn.GetFldVal("Select HName from Head where HCoCd= '" + gs_CoCd + "' and HCd='" + ctSelfCmCd + "'")
'End Function
Public Sub SetRepText(ByVal Rep As CRAXDRT.Report, _
                    ByVal TxtFld As String, _
                    ByVal TxtVal As String, _
                    Optional ByVal Frmt As Variant)
  Dim Section1 As Section
  Dim RptObj1 As Object
  For Each Section1 In Rep.Sections
    For Each RptObj1 In Section1.ReportObjects
      Select Case RptObj1.Kind
      Case crTextObject
        If UCase(RptObj1.Name) = UCase(TxtFld) Then
          If Not IsMissing(Frmt) Then
            RptObj1.SetText Format(TxtVal, Frmt)
            Exit Sub
          Else
            RptObj1.SetText TxtVal
            Exit Sub
          End If
        End If
      End Select
    Next RptObj1
  Next Section1
End Sub


'***** Shilpa *******
Public Function GetUsrLevel() As Single
    GetUsrLevel = goCn.GetFldVal("Select PNum from Param where Ptyp = 'USR' and PmCd = '" + Trim(gs_UsrCd) + "'")
    'GetUsrLevel = goCn.GetFldVal("Select PNum from Param where Ptyp = 'USR' and PmCd = '" + Trim(EmrFrmMDI.TXT_UCD.text) + "'")
End Function
'***** Shilpa *******


'**** urmi AvgStkRt *********
' Zubin 212
'SRITODO
Public Function GetAvgStkRt(ByVal mCoCd As String, ByVal mRmLoc As String, ByVal mRmCd As String, _
                            ByVal mLotNo As String, ByVal mRmSz As Single, Optional ByVal mRmSz2 As Single, Optional ByVal mRmSz3 As Single, Optional ByVal mBYy As String, _
                            Optional ByVal mBChr As String, Optional ByVal mBNo As Single) As Single
  Dim wRmLocTyp As String, wTotalWt As String, wRmCtg As String
  wRmLocTyp = goCn.GetFldVal("Select LocTyp from Loc where LocCoCd= '" + mCoCd + "' " + _
                                    "and LocCd= '" + mRmLoc + "' ")
  'SRITODO
  '*** Jen 2.12 (30/12/05) (Trim added)
  If Trim(wRmLocTyp) <> "W" Then
    '*** (Bef 2.14 Next Patch 2)
    'wTotalWt = goCn.GetFldVal("Select sum(SRmDrWt-SRmCrWt) from SYyMm " + _
    '               "where SCoCd= '" + mCoCd + "' and SLoc= '" + mRmLoc + "' and " + _
    '               "SRmCd= '" + mRmCd + "' and SLotNo= '" + mLotNo + "' and SRmSz= " + CStr(mRmSz))
    '*** (Bef 2.14 Next Patch 2)
    
    '*** (Jen 2.14 Next Patch 2)
    wTotalWt = goCn.GetFldVal("Select sum(SRmDrWt- SRmCrWt) from SYyMm " + _
                    "where " + IIF(gs_Partition = ctCurrPrtn, " SPrtKey='" + ctCurrPrtn + "' And ", " SYyMm<>'0' And ") + " SCoCd= '" + mCoCd + "' and SLoc= '" + mRmLoc + "' and " + _
                    "SRmCd= '" + mRmCd + "' and SLotNo= '" + mLotNo + "' " + _
                    "and convert (decimal (14, 4), SRmSz)= convert (decimal (14, 4), " + CStr(mRmSz) + ")" + _
                    "and convert (decimal (14, 4), SRmSz2)= convert (decimal (14, 4), " + CStr(mRmSz2) + ")" + _
                    "and convert (decimal (14, 4), SRmSz3)= convert (decimal (14, 4), " + CStr(mRmSz3) + ")")
                    
    '*** (Jen 2.14 Next Patch 2)
    
    'If Round(wTotalWt, 3) > 0 Then
    If goCn.GetFldVal("Select convert (decimal (10, 3), " + CStr(wTotalWt) + ") ") > 0 Then
      '*** (Bef 2.14 Next Patch 2)
      'GetAvgStkRt = goCn.GetFldVal("Select sum((SRmDrWt-SRmCrWt)* SRmStkRt) / Cast(" + CStr(wTotalWt) + " As Decimal(10,3)) " + _
      '         "from SYyMm where SCoCd= '" + mCoCd + "' and SLoc= '" + mRmLoc + "' and " + _
      '         "SRmCd= '" + mRmCd + "' and SLotNo= '" + mLotNo + "' and SRmSz= " + CStr(mRmSz))
      '*** (Bef 2.14 Next Patch 2)

      '*** (Jen 2.14 Next Patch 2)
      GetAvgStkRt = goCn.GetFldVal("Select convert (decimal (10, 3), sum((SRmDrWt-SRmCrWt)* SRmStkRt) / " + CStr(wTotalWt) + ") " + _
               "from SYyMm where " + IIF(gs_Partition = ctCurrPrtn, " SPrtKey='" + ctCurrPrtn + "' And ", " SYyMm<>'0' And ") + " SCoCd= '" + mCoCd + "' and SLoc= '" + mRmLoc + "' and " + _
               "SRmCd= '" + mRmCd + "' and SLotNo= '" + mLotNo + "' " + _
               "and convert (decimal (14, 4), SRmSz)= convert (decimal (14, 4), " + CStr(mRmSz) + ") " + _
               "and convert (decimal (14, 4), SRmSz2)= convert (decimal (14, 4), " + CStr(mRmSz2) + ")" + _
               "and convert (decimal (14, 4), SRmSz3)= convert (decimal (14, 4), " + CStr(mRmSz3) + ") ")
      '*** (Jen 2.14 Next Patch 2)
                   
    Else
      GetAvgStkRt = 0#
    End If

' Manoj 2.9.0 Patch
    wRmCtg = goCn.GetFldVal("Select RmCtg From RmMst Where RmCd = '" & mRmCd & "' And RmPrtKey='" + ctCurrPrtn + "' ")
    If GetAvgStkRt <= 0 And (UCase$(wRmCtg) = "C" Or UCase$(wRmCtg) = "D") Then
      '*** (Bef 2.14 Next Patch 2)
      'wTotalWt = goCn.GetFldVal("Select sum(SRmDrWt-SRmCrWt) from SYyMm " + _
      '                "Join RmMst On RmCd = SRmCd Join RmRt On RrTcTyp = 'RSZ' And RrCmCd = 'ZSELF' And " + _
      '                "RrCtg = RmCtg And RrSCtg = RmSCtg And RrFrLn <= SRmSz And RrToLn >= SRmSz " + _
      '               "where SCoCd= '" + mCoCd + "' and SLoc= '" + mRmLoc + "' and " + _
      '               "SRmCd= '" + mRmCd + "' and SLotNo= '" + mLotNo + "' and " + _
      '               "RrFrLn <= " + CStr(mRmSz) + " And RrToLn >= " + CStr(mRmSz))
      '*** (Bef 2.14 Next Patch 2)
      
      
      '*** (Jen 2.14 Next Patch 2)
      wTotalWt = goCn.GetFldVal("Select sum(SRmDrWt-SRmCrWt) from SYyMm " + _
                      " Join RmMst On RmCd = SRmCd " + IIF(gs_Partition = ctCurrPrtn, " and RmPrtKey=SPrtKey ", "") + _
                      " Join RmRt On RrTcTyp = 'RSZ' And RrCmCd = 'ZSELF' And " + _
                      "RrCtg = RmCtg And RrSCtg = RmSCtg " + IIF(gs_Partition = ctCurrPrtn, " and RmPrtKey=RrPrtKey ", "") + _
                      "And convert (decimal (14, 4), RrFrLn) <= convert (decimal (14, 4), SRmSz) " + _
                      "And convert (decimal (14, 4), RrToLn) >= convert (decimal (14, 4), SRmSz) " + _
                      "where SCoCd= '" + mCoCd + "' and SLoc= '" + mRmLoc + "' " + _
                      "and SRmCd= '" + mRmCd + "' and SLotNo= '" + mLotNo + "' " + _
                      "and convert (decimal (14, 4), RrFrLn) <= convert (decimal (14, 4), " + CStr(mRmSz) + ") " + _
                      "and convert (decimal (14, 4), RrToLn) >= convert (decimal (14, 4), " + CStr(mRmSz) + ") " + _
                      IIF(gs_Partition = ctCurrPrtn, " and RmPrtKey='" + ctCurrPrtn + "'", " and SYyMm<>'0'"))
      '*** (Jen 2.14 Next Patch 2)
      
      'If Round(wTotalWt, 3) > 0 Then
      If goCn.GetFldVal("Select convert (decimal (10, 3), " + CStr(wTotalWt) + ") ") > 0 Then
        '*** (Bef 2.14 Next Patch 2)
        'GetAvgStkRt = goCn.GetFldVal("Select sum((SRmDrWt-SRmCrWt)* SRmStkRt) / Cast(" + CStr(wTotalWt) + " As Decimal(10,3))  from SYyMm " + _
        '          "Join RmMst On RmCd = SRmCd Join RmRt On RrTcTyp = 'RSZ' And RrCmCd = 'ZSELF' And " + _
        '          "RrCtg = RmCtg And RrSCtg = RmSCtg And RrFrLn <= SRmSz And RrToLn >= SRmSz " + _
        '         "where SCoCd= '" + mCoCd + "' and SLoc= '" + mRmLoc + "' and " + _
        '         "SRmCd= '" + mRmCd + "' and SLotNo= '" + mLotNo + "' and " + _
        '         "RrFrLn <= " + CStr(mRmSz) + " And RrToLn >= " + CStr(mRmSz))
        '*** (Bef 2.14 Next Patch 2)
        
        '*** (Jen 2.14 Next Patch 2)
        GetAvgStkRt = goCn.GetFldVal("Select convert (decimal (10, 3), sum((SRmDrWt-SRmCrWt)* SRmStkRt) / " + CStr(wTotalWt) + ") from SYyMm " + _
                  "Join RmMst On RmCd = SRmCd " + IIF(gs_Partition = ctCurrPrtn, " and RmPrtKey=SPrtKey  ", "") + _
                  "Join RmRt On RrTcTyp = 'RSZ' And RrCmCd = 'ZSELF' And " + _
                  "RrCtg = RmCtg And RrSCtg = RmSCtg " + IIF(gs_Partition = ctCurrPrtn, " and RmPrtKey=RrPrtKey ", "") + _
                  "And convert (decimal (14, 4), RrFrLn) <= convert (decimal (14, 4), SRmSz) " + _
                  "And convert (decimal (14, 4), RrToLn) >= convert (decimal (14, 4), SRmSz) " + _
                  "where SCoCd= '" + mCoCd + "' and SLoc= '" + mRmLoc + "' and " + _
                  "SRmCd= '" + mRmCd + "' and SLotNo= '" + mLotNo + "' and " + _
                  "convert (decimal (14, 4), RrFrLn)<= convert (decimal (14, 4), " + CStr(mRmSz) + ") And " + _
                  "convert (decimal (14, 4), RrToLn)>= convert (decimal (14, 4), " + CStr(mRmSz) + ")" + _
                  IIF(gs_Partition = ctCurrPrtn, " and RmPrtKey='" + ctCurrPrtn + "'", " and SYyMm<>'0' "))
        '*** (Jen 2.14 Next Patch 2)
        If GetAvgStkRt < 0 Then GetAvgStkRt = 0#
      Else
        GetAvgStkRt = 0#
      End If
    End If
' Manoj 2.9.0 Patch

  Else
    '*** (Bef 2.14 Next Patch 2)
    'GetAvgStkRt = goCn.GetFldVal("Select min(TdRmStkRt) from TxnD where TdCoCd= '" + mCoCd + _
    '                             "' and TdByy= '" + mBYy + "' and TdBChr= '" + mBChr + _
    '                             "' and TdBNo = " + CStr(mBNo) + " and TdRmCd= '" + mRmCd + _
    '                             "' and TdLotNo= '" + mLotNo + "' and TdRmSz= " + CStr(mRmSz))
    '*** (Bef 2.14 Next Patch 2)
    
    '*** (Jen 2.14 Next Patch 2)
    GetAvgStkRt = goCn.GetFldVal("Select convert (decimal (10, 3), IsNull(min(TdRmStkRt),0)) From Txnd Where " + _
                                IIF(gs_Partition = ctCurrPrtn, " TdPrtKey='" + ctCurrPrtn + "' And ", "") + " TdCoCd= '" + mCoCd + _
                                 "' and TdBYy= '" + mBYy + "' and TdBChr= '" + mBChr + _
                                 "' and TdBNo = " + CStr(mBNo) + " and TdRmCd= '" + mRmCd + _
                                 "' and TdLotNo= '" + mLotNo + "' " + _
                                 " and convert (decimal (14, 4), TdRmSz)= convert (decimal (14, 4), " + CStr(mRmSz) + ")" + _
                                 " and convert (decimal (14, 4), TdRmSz2)= convert (decimal (14, 4), " + CStr(mRmSz2) + ")" + _
                                 " and convert (decimal (14, 4), TdRmSz3)= convert (decimal (14, 4), " + CStr(mRmSz3) + ")")

  End If
End Function
Public Function GetAvgStkRtJL(ByVal mCoCd As String, ByVal mRmLoc As String, ByVal mRmCd As String, _
                            ByVal mLotNo As String, ByVal mRmSz As Single, Optional ByVal mRmSz2 As Single, Optional ByVal mRmSz3 As Single)
    Dim wRes As MDORowSet
    
    Set wRes = goCn.OpenRes("Select isnull(sum((case when ((OmToDmDc='D' and OmToDmLoc='" + mRmLoc + "') or (OmFrDmDc='D' and OmFrDmLoc='" + mRmLoc + "')) then +1 " + _
            " when ((OmToDmDc='C' and OmToDmLoc='" + mRmLoc + "') or (OmFrDmDc='C' and OmFrDmLoc='" + mRmLoc + "')) then -1 " + _
            " else 0 end) * OdJLVchVal),0) as Val, " + _
                               "isnull(sum((case when ((OmToDmDc='D' and OmToDmLoc='" + mRmLoc + "') or (OmFrDmDc='D' and OmFrDmLoc='" + mRmLoc + "')) then +1 " + _
            "  when ((OmToDmDc='C' and OmToDmLoc='" + mRmLoc + "') or (OmFrDmDc='C' and OmFrDmLoc='" + mRmLoc + "')) then -1 " + _
            " else 0 end) * OdJLWt),0) as Wt " + _
          " from OrdDsg Ins Join OrdMst " + _
          " on OmPrtKey=Ins.OdPrtKey and OmIdNo=Ins.OdOmIdNo and OmCoCd= Ins.OdCoCd and OmTc= Ins.OdTc and OmYy= Ins.OdYy and OmChr= Ins.OdChr and OmNo= Ins.OdNo " + _
          " Where " + _
          " ((OmToDmDc='D' and OmToDmLoc='" + mRmLoc + "') or (OmFrDmDc='D' and OmFrDmLoc='" + mRmLoc + "'))" + _
          " and Ins.OdCoCd= '" + mCoCd + "' and Ins.OdJLRmCd = '" + mRmCd + "' and Ins.ODJLLotNo = '" + mLotNo + "' and Ins.ODJLLn1 = " + _
           CStr(mRmSz) + " and Ins.ODJLLn2 = " + CStr(mRmSz2) + " and Ins.ODJLLn3 = " + CStr(mRmSz3) + "")
    If (wRes.RecCount > 0) Then
      If wRes!Wt > 0 Then
        GetAvgStkRtJL = wRes!Val / wRes!Wt
      End If
    End If
End Function
' ########################  Manoj #### Ver: 2.0.6 #### Date: 25/12/2003  ########################
Public Sub SetScriptCtl(ByRef objScrpt As Object, ByRef pcol_TagFlds As Collection)
  objScrpt.Reset
  objScrpt.addObject "Fld", pcol_TagFlds
  objScrpt.addObject "Fun", New EmrClsFun
End Sub
Public Function mEvalExp(ByRef ScptCtl As Object, ByVal wstValExp As String) As String
On Error GoTo ErrHndlr
  wstValExp = CStr(ScptCtl.Eval(wstValExp))
  mEvalExp = wstValExp
  Exit Function
ErrHndlr:
  If ScptCtl.Error.Number = 6 Or ScptCtl.Error.Number = 11 Then
    mEvalExp = "*.*"
  Else
    mEvalExp = "***"
  End If
  DispMsg ScptCtl.Error.Description & _
      IIF(ScptCtl.Error.text <> "", " : " & ScptCtl.Error.text, ""), etWarning
End Function
Public Function mEvalChk(ByRef wstRetVal As String, ByRef ScptCtl As Object, ByRef pcol_TgFlds As Collection, ByVal wstValExp As String) As Boolean
On Error GoTo ErrHndlr
  mEvalChk = True
  wstRetVal = ""
  ScptCtl.AddCode "Public Sub Check()" & vbCrLf & _
                  "Dim Exp1 " & vbCrLf & _
                  "Exp1 = " & CStr(wstValExp) & vbCrLf & _
                  "End Sub"
  SetScriptCtl ScptCtl, pcol_TgFlds
  wstValExp = CStr(ScptCtl.Eval(wstValExp))
  Exit Function
ErrHndlr:
  Select Case ScptCtl.Error.Number
    Case 5, 13, 438, 450
    mEvalChk = False
  End Select
  If ScptCtl.Error.text <> "" Then
    mEvalChk = False
  End If
  wstRetVal = ScptCtl.Error.Description & _
      IIF(ScptCtl.Error.text <> "", " : " & Right$(ScptCtl.Error.text, IIF(ScptCtl.Error.text <> "", (Len(ScptCtl.Error.text) - 7), 0)), "")
'  wstRetVal = ScptCtl.Error.Number & " : " & ScptCtl.Error.Description & " " & _
'      IIf(ScptCtl.Error.Text <> "", " : " & Right$(ScptCtl.Error.Text, IIf(ScptCtl.Error.Text <> "", (Len(ScptCtl.Error.Text) - 7), 0)), "")
End Function
Public Function mEvFndStrCnt(ByRef pStrMVal As String, ByRef pStrFVal As String) As Integer
  Dim wiCnt As Integer, wiPos As Integer
  mEvFndStrCnt = 0
  wiCnt = 1
  Do
    wiPos = InStr(wiCnt, UCase$(pStrMVal), UCase$(pStrFVal))
    If wiPos > 0 Then mEvFndStrCnt = mEvFndStrCnt + 1: wiCnt = wiPos + Len(pStrFVal)
  Loop While (wiPos > 0)
End Function
Public Function RemoveWhiteSpaces(ByVal psVal As String, Optional ByVal pbFormat As Boolean = False) As String
  If pbFormat = True Then
    Do While InStr(1, psVal, vbCrLf) > 0
      psVal = Replace(psVal, vbCrLf, " ")
    Loop
  End If
  Do While InStr(1, psVal, vbTab) > 0
    psVal = Replace(psVal, vbTab, " ")
  Loop
  Do While InStr(1, psVal, "  ") > 0
    psVal = Replace(psVal, "  ", " ")
  Loop
  RemoveWhiteSpaces = psVal
End Function
Public Function isExists(ByRef pColVal As Collection, ByVal pKey) As Boolean
On Error GoTo ErrHndlr
  Dim wVarTemp As Variant
  wVarTemp = pColVal.Item(pKey)
  isExists = True
  Exit Function
ErrHndlr:
  isExists = False
End Function
Public Sub RemoveItem(ByRef pColVal As Collection, Optional ByVal pKey)
On Error Resume Next
  Dim wiCnt As Integer
  If IsMissing(pKey) = True Then
    For wiCnt = 0 To pColVal.Count - 1
      pColVal.Remove 1
    Next
  Else
    If isExists(pColVal, pKey) = True Then pColVal.Remove pKey
  End If
End Sub
Public Sub AddItem(ByRef pColVal As Collection, ByVal pValue, Optional ByVal pKey, Optional ByVal pBefore, Optional ByVal pAfter)
  If IsMissing(pKey) = False Then RemoveItem pColVal, pKey
  pColVal.Add pValue, pKey, pBefore, pAfter
End Sub

Sub ShowAllItem(ByRef pColVal As Object)
  Dim wiCnt
'  For wiCnt = 1 To pColVal.Count
'    Debug.Print CStr(pColVal.Key(wiCnt)) & " " & CStr(pColVal.Item(wiCnt))
'  Next
End Sub
Public Sub SetTgFldTypCol()
  Dim woRw As MwfLib.MDORowSet
  Set woRw = goCn.OpenResultset("SELECT PMCD, PValue FROM PARAM WHERE PTYP = 'TAGFLDS'")
  With woRw
    If Not (.EOF And .BOF) Then
      .MoveFirst
      While Not .EOF
        AddItem gColTgFldTyp, CStr(.FldValue("PValue")), CStr(.FldValue("PMCD"))
        .MoveNext
      Wend
    End If
  End With
End Sub
Public Sub SetTagFldColl(ByRef pcol_TgFld As Collection)
  Dim woRw As MwfLib.MDORowSet
  RemoveItem pcol_TgFld
  Set woRw = goCn.OpenResultset("SELECT PMCD, PValue FROM PARAM WHERE PTYP = 'TAGFLDS'")
  With woRw
    If Not (.EOF And .BOF) Then
      .MoveFirst
      While Not .EOF
        Select Case UCase$(CStr(.FldValue("PValue")))
        Case UCase$("C"), UCase$("U")
          AddItem pcol_TgFld, "", CStr(.FldValue("PMCD"))
        Case UCase$("I"), UCase$("F")
          AddItem pcol_TgFld, 1, CStr(.FldValue("PMCD"))
        Case UCase$("B")
          AddItem pcol_TgFld, False, CStr(.FldValue("PMCD"))
        End Select
        .MoveNext
      Wend
    End If
  End With
End Sub

' ###########################################  Manoj  ###########################################

' ########################  Manoj #### Ver: 2.0.6 #### Date: 16/06/2004  ########################
' Function to Check Login Session is working or Not
' Used By MDI Form & Listing Form in ShowForm
Public Function ChkLogin(poCn As MwfLib.MDOConnection) As Boolean
  '*** Jay 3.1.1 [UlCoCd Added]
  '6.1-49 UlSysCd added
  ChkLogin = poCn.RecSeek("Select 1 from UsrLogin Where UlUsrCd='" + UCase$(Trim$(gs_UsrCd)) + "' And " + _
                                  "UlCoCd='" + UCase$(Trim$(gs_CoCd)) + "' And UlHostNm = '" + gs_HostNm + "' and UlSysCd='" + IIF(ctGenMenuUsr = "MIC", "EMR", "JEMR") + "'")
  If Not ChkLogin Then MsgBox "Your Login Session is Terminated, Please Login Again"
End Function
' ###########################################  Manoj  ###########################################

Public Sub SetSubFormula(ByVal ReqReport As CRAXDRT.SubreportObject, ByVal FormulaName As String, ByVal Expn As String)
    FormulaName = Trim(FormulaName)
    Dim wForm As CRAXDRT.FormulaFieldDefinition
    For Each wForm In ReqReport.OpenSubreport.FormulaFields
        If UCase(wForm.FormulaFieldName) = UCase(FormulaName) Then
            wForm.text = Expn
            Exit Sub
        End If
    Next wForm
End Sub

' ########################  Manoj #### Ver: 2.0.10 #### Date: 21/02/2005  ########################
Public Sub ShowRepSummFld(FldG1F As FieldObject, FldG2F As FieldObject, FldG3F As FieldObject, FldRF As FieldObject, GrpOpts As String, Grps As String, Optional HideCurrent As Boolean = False, _
    Optional HideAllButCurrent As Boolean = False)
  ' ***** Manali 3.10.0 - Change in routine- added new parameter - This Change is done to show Summary field for desired Group Only n Suppress for others
  Dim ws_GrpOpt() As String, i As Integer, j As Integer, wColFld As Collection
  Set wColFld = New Collection
  ws_GrpOpt = Split(GrpOpts, ",")
  wColFld.Add FldRF, "RF"
  wColFld.Add FldG1F, "G1F"
  wColFld.Add FldG2F, "G2F"
  wColFld.Add FldG3F, "G3F"
  If UBound(ws_GrpOpt) >= 2 Then
    For i = 0 To 2
      If InStr(1, UCase$("," + Grps + ","), UCase$("," + ws_GrpOpt(i) + ",")) > 0 Then
        For j = 1 To wColFld.Count
          ' ***** Manali 3.10.0
          If HideAllButCurrent = True Then
            If (i + 2) <> j Then wColFld(j).Suppress = True
          Else
          ' ***** Manali 3.10.0
            If HideCurrent = True Then
              If (i + 3) > j Then wColFld(j).Suppress = True
            Else
              If (i + 2) > j Then wColFld(j).Suppress = True
            End If
          End If
        Next
      End If
    Next
  End If
  Do While wColFld.Count <= 0
    wColFld.Remove 1
  Loop
  Set wColFld = Nothing
End Sub
' ###########################################  Manoj  ###########################################

'*** Jay 3.2.0 [PicNm]
Public Sub SetRptPic(ByVal mf_OriginalHt As Single, ByVal mf_OriginalWdth As Single, _
                      ByVal mo_PicObj As Object, ByVal ms_TcTyp As String, _
                      ByVal mb_LineDraw As Boolean, ByVal mb_GetNxtTypOfPic As Boolean, _
                      ByVal ms_DesignCd As String, Optional ByVal ms_DesignSz As Variant, _
                      Optional ByVal mo_PicSec As Variant, _
                      Optional ms_OdCoCd As String, Optional ms_mOdTc As String, Optional ms_mOdYy As String, _
                      Optional ms_OdChr As String, Optional ml_mOdNo As Double, Optional ml_mOdSr As Double, Optional ms_mOdPicNm As String, Optional ms_ImgNo As String, Optional mHDImg As Boolean)

  '*** Jen 2.12 (11/08/05)
  
  On Error GoTo ErrRptPic
  
  Dim ws_pth As String, ws_Sz As String
  
  'Sachin - 3.11.2
  Dim wHdImg As Boolean
  If Not IsMissing(mHDImg) And Not mb_LineDraw Then
    wHdImg = mHDImg
  Else
    wHdImg = False
  End If
  
  If Not IsMissing(mo_PicSec) Then If mo_PicSec.Suppress = True Then Exit Sub
  Set mo_PicObj.FormattedPicture = LoadPicture("")
  
  If IsMissing(ms_DesignSz) Then
    ws_Sz = ""
  Else
    ws_Sz = ms_DesignSz
  End If
    
  '*** Jay 3.2.0 [PicNm]
  If ms_mOdPicNm <> "" Then
    ws_pth = GetPictPath(ms_TcTyp, mb_LineDraw, ms_DesignCd, ws_Sz, , , , , , , ms_mOdPicNm, , wHdImg)
    If Dir(ws_pth) = "" Then ws_pth = GetPictPath(ms_TcTyp, mb_LineDraw, ms_DesignCd, ws_Sz, , , , , , , ms_mOdPicNm)
  ElseIf Not (IsMissing(ms_OdCoCd) Or IsMissing(ms_mOdTc) Or IsMissing(ms_mOdYy) Or IsMissing(ms_OdChr) Or IsMissing(ml_mOdSr) Or IsMissing(ml_mOdSr)) Then
    ws_pth = GetPictPath(ms_TcTyp, mb_LineDraw, ms_DesignCd, ws_Sz, ms_OdCoCd, ms_mOdTc, ms_mOdYy, ms_OdChr, ml_mOdNo, ml_mOdSr, , , wHdImg)
    If Dir(ws_pth) = "" Then ws_pth = GetPictPath(ms_TcTyp, mb_LineDraw, ms_DesignCd, ws_Sz, ms_OdCoCd, ms_mOdTc, ms_mOdYy, ms_OdChr, ml_mOdNo, ml_mOdSr)
  Else
    ws_pth = GetPictPath(ms_TcTyp, mb_LineDraw, ms_DesignCd, ws_Sz, , , , , , , , , wHdImg)
    If Dir(ws_pth) = "" Then ws_pth = GetPictPath(ms_TcTyp, mb_LineDraw, ms_DesignCd, ws_Sz)
  End If
  'ws_pth = GetPictPath(ms_TcTyp, mb_LineDraw, ms_DesignCd, ws_Sz)
  '*** Jay 3.2.0 [PicNm]
  
  If ms_ImgNo <> "" Then ws_pth = Replace(ws_pth, ".jpg", "_" + ms_ImgNo + ".jpg") 'Sachin 3.11.0 - 3 additional images

  If ws_pth = "" Or Dir(ws_pth) = "" Then
    If mb_GetNxtTypOfPic = True Then
      '*** Jay 3.2.0 [PicNm]
        If Not IsMissing(ms_mOdPicNm) Then
          ws_pth = GetPictPath(ms_TcTyp, Not (mb_LineDraw), ms_DesignCd, ws_Sz, , , , , , , ms_mOdPicNm, , wHdImg)
          If Dir(ws_pth) = "" Then ws_pth = GetPictPath(ms_TcTyp, Not (mb_LineDraw), ms_DesignCd, ws_Sz, , , , , , , ms_mOdPicNm)
        ElseIf Not (IsMissing(ms_OdCoCd) Or IsMissing(ms_mOdTc) Or IsMissing(ms_mOdYy) Or IsMissing(ms_OdChr) Or IsMissing(ml_mOdSr) Or IsMissing(ml_mOdSr)) Then
          ws_pth = GetPictPath(ms_TcTyp, Not (mb_LineDraw), ms_DesignCd, ws_Sz, ms_OdCoCd, ms_mOdTc, ms_mOdYy, ms_OdChr, ml_mOdNo, ml_mOdSr, , , wHdImg)
          If Dir(ws_pth) = "" Then ws_pth = GetPictPath(ms_TcTyp, Not (mb_LineDraw), ms_DesignCd, ws_Sz, ms_OdCoCd, ms_mOdTc, ms_mOdYy, ms_OdChr, ml_mOdNo, ml_mOdSr)
        Else
          ws_pth = GetPictPath(ms_TcTyp, Not (mb_LineDraw), ms_DesignCd, ws_Sz, , , , , , , , , wHdImg)
          If Dir(ws_pth) = "" Then ws_pth = GetPictPath(ms_TcTyp, Not (mb_LineDraw), ms_DesignCd, ws_Sz)
        End If
      'ws_pth = GetPictPath(ms_TcTyp, Not (mb_LineDraw), ms_DesignCd, ws_Sz)
      '*** Jay 3.2.0 [PicNm]
    End If
  End If
  
  If ws_pth <> "" And Dir(ws_pth) <> "" Then
    Set mo_PicObj.FormattedPicture = LoadPicture(ws_pth)
    
    Dim wPicNewHt As Single, wPicNewWdt As Single
  
    '*** Storing The Actual Picture Height And Width ***
    wPicNewHt = mo_PicObj.FormattedPicture.Height: wPicNewWdt = mo_PicObj.FormattedPicture.Width
    
    '*** If Pic Ht > Wdth Then The Pic Ht Is Set To Original Pic Control Ht And The Pic Wdth Is Set Proportionately ***
    '*** Follow Same Procedure If Wdth > Ht ***
    If wPicNewHt > wPicNewWdt Then
      mo_PicObj.Height = mf_OriginalHt
      mo_PicObj.Width = mf_OriginalWdth * MWLib.Div(wPicNewWdt, wPicNewHt)
    Else
      mo_PicObj.Width = mf_OriginalWdth
      mo_PicObj.Height = mf_OriginalHt * MWLib.Div(wPicNewHt, wPicNewWdt)
    End If
    
    Exit Sub
  End If
  
  If Not IsMissing(mo_PicSec) Then
    mo_PicSec.Suppress = True
  End If
  
  Exit Sub

ErrRptPic:
  If Err.Number = 481 Then
      DispMsg "Improper Picture Format, Please Change the Picture File", etInfo
  Else
      DispMsg Err.Description, etInfo
  End If
  Set mo_PicObj.FormattedPicture = LoadPicture("")
  '*** Jen 2.12 (11/08/05)
End Sub
Public Sub SetRptPic1(ByVal mo_PicObj As Object, ByVal ms_TcTyp As String, _
                     ByVal mb_LineDraw As Boolean, ByVal mb_GetNxtTypOfPic As Boolean, _
                     ByVal ms_DesignCd As String, Optional ByVal ms_DesignSz As Variant, _
                     Optional ByVal mo_PicSec As Variant)
  
  '*** (Bef 2.12)
'  '*** Jen 2.11 (20/06/05)
'  Dim ws_Pth As String, ws_Sz As String
'
'  'Manoj 2.11.0 Correction
'  If Not IsMissing(mo_PicSec) Then If mo_PicSec.Suppress = True Then Exit Sub
'  'Manoj 2.11.0 Correction
'
'  Set mo_PicObj.FormattedPicture = LoadPicture("")
'  If IsMissing(ms_DesignSz) Then
'    ws_Sz = ""
'  Else
'    ws_Sz = ms_DesignSz
'  End If
'
'  ws_Pth = GetPictPath(ms_TcTyp, mb_LineDraw, ms_DesignCd, ws_Sz)
'
'  If ws_Pth = "" Or Dir(ws_Pth) = "" Then
'    If mb_GetNxtTypOfPic = True Then
'      ws_Pth = GetPictPath(ms_TcTyp, Not (mb_LineDraw), ms_DesignCd, ws_Sz)
'    End If
'  End If
'
'  If ws_Pth <> "" And Dir(ws_Pth) <> "" Then
'    Set mo_PicObj.FormattedPicture = LoadPicture(ws_Pth)
'    Exit Sub
'  End If
'
'  If Not IsMissing(mo_PicSec) Then
'    mo_PicSec.Suppress = True
'  End If
'  '*** Jen 2.11 (20/06/05)
  '*** (Bef 2.12)
End Sub


'*** (Bef 20/06/05)
' **** Zubin 211 **** '
' **** Zubin 211 **** '
'Public Sub SetRptPic(ByVal mPicObj As Object, ByVal mTcType As String, ByVal mDesignCd As String, Optional ByVal mDesignSz As Variant, _
'                            Optional ByVal mLineDraw As Variant, Optional ByVal mPicSec As Object)
'
'  Dim wPth As String
'
'  If Not IsMissing(mPicSec) Then If mPicSec.Suppress = True Then Exit Sub
'
'  If IsMissing(mDesignSz) Then
'    If IsMissing(mLineDraw) Then
'      wPth = GetPictPath(mTcType, mDesignCd)
'    Else
'      wPth = GetPictPath(mTcType, mDesignCd, , mLineDraw)
'    End If
'  Else
'    If IsMissing(mLineDraw) Then
'      wPth = GetPictPath(mTcType, mDesignCd, mDesignSz)
'    Else
'      wPth = GetPictPath(mTcType, mDesignCd, mDesignSz, mLineDraw)
'    End If
'  End If
'
'  If wPth <> "" And Dir(wPth) <> "" Then
'    Set mPicObj.FormattedPicture = LoadPicture(wPth)
'    Exit Sub
'  End If
'
'  If Not IsMissing(mPicSec) Then mPicSec.Suppress = True
'End Sub
'*** (Bef 20/06/05)


' **** Zubin 211 **** '
' **** Zubin 211 **** '

' **** Zubin 211 **** '
' **** Zubin 211 **** '
'''''Public Function IsValidUsr(ByVal mCoCd As String, ByVal mTc As String, ByVal mSubTc As String, ByVal mUsrCd As String) As Boolean
'''''' **** This function returns true if User has rights to a 'CHR'
'''''  Dim ws_UsrArr() As String, wi_Index As Integer
'''''  ws_UsrArr = Split(gocn.GetFldVal("Select vPValue4 from vParam where vPCoCd= '" + mCoCd + "' and " + _
'''''                                      "vPTyp='CHR' and vPMCd='" + mTc + "' and vPSCd= '" + mSubTc + "'"), ",")
'''''
'''''  If UBound(ws_UsrArr) = -1 Then IsValidUsr = True: Exit Function
'''''  If IsArray(ws_UsrArr) Then
'''''    For wi_Index = 0 To UBound(ws_UsrArr)
'''''      IsValidUsr = IIF(UCase(mUsrCd) = UCase(ws_UsrArr(wi_Index)), True, False)
'''''      If IsValidUsr = True Then: Exit Function
'''''    Next wi_Index
'''''  End If
'''''  IsValidUsr = False
'''''End Function
' **** Zubin 211 **** '
' **** Zubin 211 **** '

Public Sub SetRptLogoPic(ByVal mf_OriginalHt As Single, ByVal mf_OriginalWdth As Single, _
                      ByVal mo_PicObj As Object, ByVal ms_CoCd As String, _
                      Optional ByVal mo_PicSec As Variant, Optional ByVal mb_SuppSec As Boolean = False, Optional ByVal mb_Dtc As Boolean = False)

  On Error GoTo ErrRptPic
  
  Dim ws_pth As String, wLogo As String
  Dim wFlNm  As String, wDsgCtg As String, wPth As String, wTyp As String
  Dim wSz As String, wDir  As String
  Dim wPicNewHt As Single, wPicNewWdt As Single
    
  If Not IsMissing(mo_PicSec) Then If mo_PicSec.Suppress = True Then Exit Sub
  Set mo_PicObj.FormattedPicture = LoadPicture("")

  wFlNm = MWLib.RmvEmbSpace(ms_CoCd)
  If mb_Dtc Then wFlNm = "Logo_DTC_Tm"      '4.1.3
  
  wLogo = "Logo"
  '*** Get the picture path from param depending on '3D' pic is to be shown
  wPth = goCn.GetFldVal("Select PDesc225 from Param Where PTyp='PCT' and PMCd='3D'") + "\" + wLogo
  
  ws_pth = wPth + "\" + wFlNm + ".jpg"

  If ws_pth <> "" And Dir(ws_pth) <> "" Then
    Set mo_PicObj.FormattedPicture = LoadPicture(ws_pth)
    '*** Storing The Actual Picture Height And Width ***
    wPicNewHt = mo_PicObj.FormattedPicture.Height: wPicNewWdt = mo_PicObj.FormattedPicture.Width

    If wPicNewHt > wPicNewWdt Then
      mo_PicObj.Height = mf_OriginalHt
      mo_PicObj.Width = mf_OriginalWdth * MWLib.Div(wPicNewWdt, wPicNewHt)
    Else
      mo_PicObj.Width = mf_OriginalWdth
      mo_PicObj.Height = mf_OriginalHt * MWLib.Div(wPicNewHt, wPicNewWdt)
    End If
    Exit Sub
  End If

  ' **** Manali 3.8.0 - mb_SuppSec added
  If Not IsMissing(mo_PicSec) And mb_SuppSec Then
    mo_PicSec.Suppress = True
  End If

Exit Sub
ErrRptPic:
  If Err.Number = 481 Then
      DispMsg "Improper Picture Format, Please Change the Picture File", etInfo
  Else
      DispMsg Err.Description, etInfo
  End If
  Set mo_PicObj.FormattedPicture = LoadPicture("")
  '*** Jen 2.12 (11/08/05)
End Sub

Public Sub GetRmSzDisp(ByVal wRmSz1 As Double, ByVal wRmSz2 As Double, ByVal wRmSz3 As Double, ByVal DispFldNm As TextObject)

Dim n1 As Double, n2 As Double, n3 As Double
Dim wRmSz123 As String
 
n1 = wRmSz1: n2 = wRmSz2: n3 = wRmSz3

If n1 <> 0 Or n2 <> 0 Or n3 <> 0 Then

wRmSz123 = CStr(n1) + "*" + CStr(n2) + "*" + CStr(n3)

DispFldNm.SetText (Left(wRmSz123, 14))

Else

DispFldNm.SetText (Space(14))
End If
End Sub

Public Function GetRmSzDisp1(ByVal wRmSz1 As Double, ByVal wRmSz2 As Double, ByVal wRmSz3 As Double) As String

Dim n1 As Double, n2 As Double, n3 As Double
Dim wRmSz123 As String
 
n1 = wRmSz1: n2 = wRmSz2: n3 = wRmSz3

If n1 <> 0 Or n2 <> 0 Or n3 <> 0 Then

wRmSz123 = CStr(n1) + "*" + CStr(n2) + "*" + CStr(n3)

GetRmSzDisp1 = Left(wRmSz123, 14)

Else

GetRmSzDisp1 = Space(14)
End If
End Function

Public Sub GetRmSzGrpSrt(ByVal wRmSz1 As String, ByVal wRmSz2 As String, ByVal wRmSz3 As String, ByVal wGrpRep As MWCTL_GRP)
    wGrpRep.Add "Rm Sz", " str(" + wRmSz1 + ",8,4 )+STR(" + wRmSz2 + ",8,4)+STR(" + wRmSz3 + ",8,4)", "(case when " + wRmSz1 + "+" + wRmSz2 + "+" + wRmSz3 + "=0 then '' else (str(" + wRmSz1 + ",8,4 )+' * '+STR(" + wRmSz2 + ",8,4)+' * '+STR(" + wRmSz3 + ",8,4)) end)", "wRmSz", "hRmsz", "Rm Ctg,Rm SubCtg,Rm Cd,Range Size,Rm Ptr", "", "", ""
End Sub

Public Function GetRmSzDBFldNm(ByVal wRmSz1 As String, ByVal wRmSz2 As String, ByVal wRmSz3 As String)
    GetRmSzDBFldNm = " str(" + wRmSz1 + ",8,4 )+STR(" + wRmSz2 + ",8,4)+STR(" + wRmSz3 + ",8,4)"
End Function
Public Function GetRmSzDispFldNm(ByVal wRmSz1 As String, ByVal wRmSz2 As String, ByVal wRmSz3 As String)
'    GetRmSzDispFldNm = "(case when " + wRmSz1 + "+" + wRmSz2 + "+" + wRmsz3 + "=0 then '' else " + _
    "replicate('0',8 - len(ltrim(str(" + wRmSz1 + ",8,4)))) + ltrim(str(" + wRmSz1 + ",8,4)) + ' * ' + " + _
    "replicate('0',8 - len(ltrim(str(" + wRmSz2 + ",8,4)))) + ltrim(str(" + wRmSz2 + ",8,4)) + ' * ' + " + _
    "replicate('0',8 - len(ltrim(str(" + wRmsz3 + ",8,4)))) + ltrim(str(" + wRmsz3 + ",8,4)) end) "
'    GetRmSzDispFldNm = "(case when " + wRmSz1 + "+" + wRmSz2 + "+" + wRmsz3 + "=0 then '' else " + _
    "(case when  " + wRmSz2 + "+" + wRmsz3 + "=0 then str(" + wRmSz1 + ",8,4 ) else " + _
    "(case when  " + wRmsz3 + "=0 then (str(" + wRmSz1 + ",8,4 )+' * '+STR(" + wRmSz2 + ",8,4)) else " + _
    "(str(" + wRmSz1 + ",8,4 )+' * '+STR(" + wRmSz2 + ",8,4)+' * '+STR(" + wRmsz3 + ",8,4)) end)"
'    GetRmSzDispFldNm = "(case when " + wRmSz1 + "+" + wRmSz2 + "+" + wRmsz3 + "=0 then '' " + _
    "when  " + wRmSz2 + "+" + wRmsz3 + "=0 then str(" + wRmSz1 + ",8,4 ) " + _
    "when  " + wRmsz3 + "=0 then (str(" + wRmSz1 + ",8,4 )+' * '+STR(" + wRmSz2 + ",8,4)) " + _
    "else (str(" + wRmSz1 + ",8,4 )+' * '+STR(" + wRmSz2 + ",8,4)+' * '+STR(" + wRmsz3 + ",8,4)) end)"
    
    
    GetRmSzDispFldNm = "(case when " + wRmSz1 + "+" + wRmSz2 + "+" + wRmSz3 + "=0 then '' " + _
    " when  " + wRmSz2 + "+" + wRmSz3 + "=0 then " + _
        " case " + wRmSz1 + " when 0.01 then '+0' when 0.02 then '+00' when 0.03 then '+000' when 0.04 then '+0000' else " + " LTrim(RTrim(STR(convert(int," + wRmSz1 + "))+(case " + wRmSz1 + "-convert(int," + wRmSz1 + ") when 0 then '' else '.'+reverse(STR(reverse(" + wRmSz1 + "-convert(int," + wRmSz1 + ")))) end ))) end " + _
    "when  " + wRmSz3 + "=0 then " + _
        " case " + wRmSz1 + " when 0.01 then '+0' when 0.02 then '+00' when 0.03 then '+000' when 0.04 then '+0000' else " + " LTrim(RTrim(STR(convert(int," + wRmSz1 + "))+(case " + wRmSz1 + "-convert(int," + wRmSz1 + ") when 0 then '' else '.'+reverse(STR(reverse(" + wRmSz1 + "-convert(int," + wRmSz1 + ")))) end ))) end +' * ' +" + _
        "LTrim(RTrim(STR(convert(int," + wRmSz2 + "))+(case " + wRmSz2 + "-convert(int," + wRmSz2 + ") when 0 then '' else '.'+reverse(STR(reverse(" + wRmSz2 + "-convert(int," + wRmSz2 + ")))) end ))) " + _
    "else " + _
        " case " + wRmSz1 + " when 0.01 then '+0' when 0.02 then '+00' when 0.03 then '+000' when 0.04 then '+0000' else " + " LTrim(RTrim(STR(convert(int," + wRmSz1 + "))+(case " + wRmSz1 + "-convert(int," + wRmSz1 + ") when 0 then '' else '.'+reverse(STR(reverse(" + wRmSz1 + "-convert(int," + wRmSz1 + ")))) end ))) end + ' * ' +" + _
        "LTrim(RTrim(STR(convert(int," + wRmSz2 + "))+(case " + wRmSz2 + "-convert(int," + wRmSz2 + ") when 0 then '' else '.'+reverse(STR(reverse(" + wRmSz2 + "-convert(int," + wRmSz2 + ")))) end ))) + ' * ' + " + _
        "LTrim(RTrim(STR(convert(int," + wRmSz3 + "))+(case " + wRmSz3 + "-convert(int," + wRmSz3 + ") when 0 then '' else '.'+reverse(STR(reverse(" + wRmSz3 + "-convert(int," + wRmSz3 + ")))) end ))) " + _
    " end )"
        
' chk in sql : - LTrim(RTrim(STR(123.00450)+(case 123.00450-convert(int,123.00450) when 0 then '' else '.'+reverse(STR(reverse(123.00450-convert(int,123.00450)))) end )))

End Function
Public Function GetBagNo(ByVal fs_FGBNo As String) As String()
  ' This function returns BYy, BChr, BNo
  ' It assumes that the bag No is valid and BLoc is not considered
  Dim ws_FGBNo() As String
  
  ws_FGBNo = Split(fs_FGBNo, "/")
  GetBagNo = ws_FGBNo
End Function
' **** Zubin 211 **** '
' **** Zubin 211 **** '

Public Sub VerMatch()
  If gs_Ver <> Trim(goCn.GetFldVal("Select PDesc225 from Param where PTyp= 'VER' and " + _
     "PMCd= 'DATVER'")) Then DispMsg "The Version Has Changed After Login. Please Login Again", etError: Exit Sub
End Sub
' Zubin 212
Public Function ImplementLotNo(ByVal ps_cocd) As Boolean
' This function returns true if LotNo is implemented (used mainly for BRmSz hlp)
  If goCn.RecSeek("select 'x' from Head where HCd= '" + ctSelfCmCd + "' and HCoCd= '" + ps_cocd + "' and " + _
                  "HLotNoYN= 'Y'") Then ImplementLotNo = True
End Function
' Zubin 212
'*** (Jen 2.12)
Public Sub ChkSelScope(ByVal pd_ErrNum As Double, ByVal ps_ErrDesc As Variant)
  If pd_ErrNum <> 0 Then
    If pd_ErrNum = -2147217900 Then
      DispMsg "Error In Report Scope Selection. ", etError
    Else
      DispMsg ps_ErrDesc + ". Cannot Show Report.", etError
    End If
  End If
End Sub
'*** (Jen 2.12)

'*** (Jen 2.13)
Public Function ChkVal(ByVal ps_Str As String, ByRef ps_Msg As String) As Boolean
  ChkVal = False
  On Error GoTo ErrH
  
  If goCn.GetFldVal(ps_Str) >= 0 Then ChkVal = True
  ps_Msg = ""
  Exit Function
ErrH:
  ps_Msg = "Validation Query Not Proper"
End Function
'*** (Jen 2.13)

' Zubin 213
Public Function CnvRt(ByVal ps_CurCd1 As String, ByVal ps_CurCd2 As String, Optional ps_Dt As Variant) As Double
  Dim wd_CnvRt As Double
  
  If ps_CurCd1 <> "" And ps_CurCd1 = ps_CurCd2 Then CnvRt = 1: Exit Function

  If Not IsMissing(ps_Dt) Then
    wd_CnvRt = goCn.GetFldVal("Select IsNull((Select DtVal from DtTbl where DtTag= 'CNV' and " + _
                "   DtCurCd1= '" + ps_CurCd1 + "' and DtCurCd2= '" + ps_CurCd2 + "' and " + _
                "   DtDt= (Select max(DtDt) from DtTbl where DtTag= 'CNV' and DtCurCd1= '" + ps_CurCd1 + "' and " + _
                "   DtCurCd2= '" + ps_CurCd2 + "' and DtDt<= '" + CStr(ps_Dt) + "')), " + _
                "       IsNull((Select DtVal from DtTbl where DtTag= 'CNV' and " + _
                "           DtCurCd1= '" + ps_CurCd2 + "' and DtCurCd2= '" + ps_CurCd1 + "' and " + _
                "           DtDt= (Select max(DtDt) from DtTbl where DtTag= 'CNV' and " + _
                "           DtCurCd1= '" + ps_CurCd2 + "' and DtCurCd2= '" + ps_CurCd1 + "' and " + _
                "           DtDt<= '" + CStr(ps_Dt) + "')), 0))")
  Else
    wd_CnvRt = goCn.GetFldVal("Select IsNull((Select DtVal from DtTbl where DtTag= 'CNV' and " + _
                "   DtCurCd1= '" + ps_CurCd1 + "' and DtCurCd2= '" + ps_CurCd2 + "' and " + _
                "   DtDt= (Select max(DtDt) from DtTbl where DtTag= 'CNV' and DtCurCd1= '" + ps_CurCd1 + "' and " + _
                "   DtCurCd2= '" + ps_CurCd2 + "')), " + _
                "         IsNull((Select DtVal from DtTbl where DtTag= 'CNV' and " + _
                "           DtCurCd1= '" + ps_CurCd2 + "' and DtCurCd2= '" + ps_CurCd1 + "' and " + _
                "           DtDt= (Select max(DtDt) from DtTbl where DtTag= 'CNV' and " + _
                "           DtCurCd1= '" + ps_CurCd2 + "' and DtCurCd2= '" + ps_CurCd1 + "' )), 0))")
  End If
  CnvRt = wd_CnvRt
End Function

Public Function MulDiv(ByVal ps_CurCd1 As String, ByVal ps_CurCd2 As String) As String
  Dim ws_MulDiv As String

  If ps_CurCd1 <> "" And ps_CurCd1 = ps_CurCd2 Then MulDiv = "M": Exit Function
  ws_MulDiv = goCn.GetFldVal("Select IsNull((Select 'M' from DtTbl where DtTag= 'CNV' and " + _
              "   DtCurCd1= '" + ps_CurCd1 + "' and DtCurCd2= '" + ps_CurCd2 + "' and " + _
              "   DtDt= (Select max(DtDt) from DtTbl where DtTag= 'CNV' and DtCurCd1= '" + ps_CurCd1 + "' and " + _
              "   DtCurCd2= '" + ps_CurCd2 + "')), " + _
              "         IsNull((Select 'D' from DtTbl where DtTag= 'CNV' and " + _
              "           DtCurCd1= '" + ps_CurCd2 + "' and DtCurCd2= '" + ps_CurCd1 + "' and " + _
              "           DtDt= (Select max(DtDt) from DtTbl where DtTag= 'CNV' and " + _
              "           DtCurCd1= '" + ps_CurCd2 + "' and DtCurCd2= '" + ps_CurCd1 + "' )), ''))")
  MulDiv = ws_MulDiv
End Function
Public Function MulDivOp(ByVal ps_MulDiv As String) As String
  Select Case UCase(ps_MulDiv)
    Case "M"
      MulDivOp = "*"
    Case "D"
      MulDivOp = "/"
  End Select
End Function
'****** Sachin 2.13.0 [12-08-2006] - [Q/W Module]
Public Function GetQWMod() As Boolean
  GetQWMod = IIF(goCn.GetFldVal("Select hQwModYN From Head Where hCd= '" + ctSelfCmCd + "' And hCoCd='" + ctSelfCoCd + "'") = "Y", True, False)
End Function
'****** Sachin 2.13.0 [12-08-2006] - [Q/W Module]
Public Function GetFrmCaption(ByVal ps_cocd As String, ByVal ps_FrmNm As String) As String
  If GetMDIFrm.ADC.Connection.RecSeek("Select 'x' From Head Where hCd= '" + ctSelfCmCd + "' And hCoCd='" + UCase(ps_cocd) + "' and hImpComLoc='Y'") Then
    GetFrmCaption = " " + ps_cocd + " - (" + ps_FrmNm + ")"
  Else
    GetFrmCaption = ps_FrmNm
  End If
End Function

'*** (Jen 2.14 Next)
Public Function oRIGNWithinRTCRange() As Boolean
  oRIGNWithinRTCRange = False
'  If go_HlpCn.RecSeek("Select 'x' from Param where PTyp= 'RTC' and PMCd= 'RTC' and PSCd= '' " + _
                  "and (convert (int, PNum)= 0 or convert (int, PNum) <= DATEPART (hh, getdate()) ) " + _
                  "and (convert (int, PNum1)= 0 or convert (int, PNum1) >= DATEPART (hh, getdate()) ) ") Then
  If go_HlpCn.RecSeek("Select 'x' from Param where PTyp= 'RTC' and PMCd= 'RTC' and PSCd= '' " + _
                  "and ((convert (int, PNum)= 0 or convert (int, PNum) >= DATEPART (hh, getdate()) ) " + _
                  "Or (convert (int, PNum1)= 0 or convert (int, PNum1) <= DATEPART (hh, getdate()) )) ") Then
    oRIGNWithinRTCRange = True
  End If
End Function
Public Sub GrdPropSetPPC(wg_Grd As MWCTL_FLX, ws_PreStr As String)
  '*** Procedure to set the Heading property of the fields of the grid
      'Note that the heading has to be appended by spaces when necessary if the
      'column width is not accomodating the length of the field

  Dim wo_RsColHd As MwfLib.MDORowSet
  Dim wi_SpaceLen As Single
  Set wo_RsColHd = goCn.OpenRes("Select PMCd, PValue, PDesc, convert(int, PNum) as qSeq from Param where PTyp= 'PRDSTG' Order By PNum ")
  With wg_Grd
    Do While Not (wo_RsColHd.EOF Or wo_RsColHd.BOF)
      '***** To Find Out The Space To Be Appended To The Heading
            '(This Is To Set The Proper Width For The Field)
      '***** 2.5 is taken as the approximate value which stands for the space occupied by one
            'character in Upper Case
      'wi_SpaceLen = (.ColProp(ws_PreStr + Trim(wo_RsColHd!PMCd)).MaxLength * 2.5) - (Len(wo_RsColHd!pValue))
      'wi_SpaceLen = (.ColProp(ws_PreStr + Trim(wo_RsColHd!PMCd)).MaxLength * (2.5 - IIF(Len(wo_RsColHd!pValue) / 5 < 2.5, Len(wo_RsColHd!pValue) / 5, 0))) - Len(wo_RsColHd!pValue)
      wi_SpaceLen = (.ColProp(ws_PreStr + Trim(wo_RsColHd!PMCd)).MaxLength * 2.5) - (Len(wo_RsColHd!pValue) * 3.3)
      
      .ColProp(ws_PreStr + Trim(wo_RsColHd!PMCd)).Heading = wo_RsColHd!pValue + IIF(wi_SpaceLen > 0, Space(Abs(wi_SpaceLen)), "")
      .ColProp(ws_PreStr + Trim(wo_RsColHd!PMCd)).ToolTipText = "Enter " + Trim(wo_RsColHd!PDesc) + " Capacity Per Day For The Line"
      
      '***** Note: Width Property cannot be used to set the width
      '*****  .ColProp(.ColProp(ws_PreStr + Trim(CStr(wo_RsColHd!qSeq))).Width = wi_CharWdth * .ColProp(ws_PreStr + Trim(CStr(wo_RsColHd!qSeq))).MaxLength
      wo_RsColHd.MoveNext
    Loop
    .SetColWidthToHd
  End With

  Set wo_RsColHd = Nothing
End Sub
'*** (Jen 2.14 Next)

'*** (Jen 3.01)
Public Function GetModVal(ByVal ps_Val As String, ByVal pb_FromIni As Boolean) As String
  Dim wi_i As Integer, ws_Val As String, wi_Rnd As Integer, wColl_YesVal As New Collection
  
  wColl_YesVal.Add "YA", CStr(1)
  wColl_YesVal.Add "YJ", CStr(2)
  wColl_YesVal.Add "YX", CStr(3)
  
  If pb_FromIni = True Then
    '*** From Ini
    ws_Val = "N"
    For wi_i = 1 To wColl_YesVal.Count
      If ws_Val = wColl_YesVal(wi_i) Then ws_Val = "Y": Exit For
    Next wi_i
  Else
    '*** To Ini
    Randomize
    ws_Val = ""
    If UCase(ps_Val) = "Y" Then
      '*** Formula to generate a random number within a range is : Int((upperbound - lowerbound + 1) * Rnd + lowerbound)
      wi_Rnd = Int((wColl_YesVal.Count - 1 + 1) * Rnd + 1)
      ws_Val = wColl_YesVal(CStr(wi_Rnd))
    Else
      ws_Val = ""
      Do While ws_Val = ""
        wi_Rnd = 0
        '*** Range from 32 to 126 (Not 124, 44)
        Do While wi_Rnd = 0
          wi_Rnd = Int((126 - 32 + 1) * Rnd + 32)
          If wi_Rnd = 44 Or wi_Rnd = 124 Then wi_Rnd = 0
        Loop
        ws_Val = Chr(wi_Rnd)
        
        wi_Rnd = 0
        '*** Range from 32 to 126 (Not 124, 44)
        Do While wi_Rnd = 0
          wi_Rnd = Int((126 - 32 + 1) * Rnd + 32)
          If wi_Rnd = 44 Or wi_Rnd = 124 Then wi_Rnd = 0
        Loop
        ws_Val = ws_Val + Chr(wi_Rnd)
        For wi_i = 1 To wColl_YesVal.Count
          If ws_Val = wColl_YesVal(wi_i) Then ws_Val = "": Exit For
        Next wi_i
      Loop
    End If
  End If
  
  GetModVal = ws_Val
End Function
'*** (Jen 3.01)

'*** (Bef 04/02/08)
'Public Function ChkMwTime() As Boolean
'*** (Bef 04/02/08)

'*** (Jenny speed) 04/02/08
Public Function ChkMwTime(ByVal po_Conn As MwfLib.MDOConnection) As Boolean
'*** (Jenny speed) 04/02/08

'****** Sachin 3.02.0 - Report Server (Replication Check) - [22-09-2007]
'6.1
po_Conn.Execute ("Set DateFormat " + DtFmtStr() + " ")
'****** Seconds Part Removed Later as SrvrTime Does Not Return Seconds (Sachin)
'If po_Conn.RecSeek("Select 'x' From Param " + _
                " Where PTyp='MWTIME' and PMCd='MWTIME' " + _
                "    And convert(smalldatetime, PValue) " + _
                " + (Case When Len(convert(varchar(10),(cast(PNum as int)/(60*60))))<2 Then '0'+ convert(varchar(10),(cast(PNum as int)/(60*60))) " + _
                " Else Convert(varchar(10),(cast(PNum as int)/(60*60))) End + ':' " + _
                " + Case When Len(convert(varchar(10),(cast(PNum as int)%(60*60)/60)))<2 Then '0'+ convert(varchar(10),(cast(PNum as int)%(60*60)/60)) " + _
                " Else Convert(varchar(10),(cast(PNum as int)%(60*60)/60)) End + ':' " + _
                " + Case When Len(convert(varchar(10),(cast(PNum as int)%(60*60)%60)))<2 Then '0'+ Convert(varchar(10),(cast(PNum as int)%(60*60)%60)) " + _
                " Else Convert(varchar(10),(cast(PNum as int)%(60*60)%60)) End) >= convert(smalldatetime,'" + Format$(goCn.SrvrDate, "DD-MM-YY") + " " + Replace(Format(goCn.SrvrTime, "00.00"), ".", ":") + ":00" + "')") Then
                '6.1
If po_Conn.RecSeek("Select 'x' From Param " + _
                " Where PTyp='MWTIME' and PMCd='MWTIME' " + _
                "    And convert(smalldatetime, PValue) " + _
                " + (Case When Len(convert(varchar(10),(cast(PNum as int)/(60*60))))<2 Then '0'+ convert(varchar(10),(cast(PNum as int)/(60*60))) " + _
                " Else Convert(varchar(10),(cast(PNum as int)/(60*60))) End + ':' " + _
                " + Case When Len(convert(varchar(10),(cast(PNum as int)%(60*60)/60)))<2 Then '0'+ convert(varchar(10),(cast(PNum as int)%(60*60)/60)) " + _
                " Else Convert(varchar(10),(cast(PNum as int)%(60*60)/60)) End + ':00') " + _
                " >= convert(smalldatetime,'" + Format$(goCn.SrvrDate, DtShortStrWithDash()) + " " + Replace(Format(goCn.SrvrTime, "00.00"), ".", ":") + ":00" + "')") Then
    ChkMwTime = True
Else
    ChkMwTime = False
End If
'****** Sachin 3.02.0 - Report Server (Replication Check) - [22-09-2007]
End Function
Public Sub DispReport(ByRef po_adcObj As Object)
  '*** (Jenny speed) 04/02/08
  If ChkMwTime(po_adcObj.Connection) = True Or gbSvrChk = True Then
  '*** (Jenny speed) 04/02/08

  '*** (Bef 04/02/08)
  ''****** Sachin 3.02 - Report Server ******
  'If ChkMwTime = True Or gbSvrChk = True Then
  '*** (Bef 04/02/08)
    po_adcObj.ShowRepDone
  Else
    DispMsg "Replication Status Not available, Cannot Display Report", etError
  End If
  '****** Sachin 3.02 - Report Server ******
End Sub

'****** Sachin 3.3.0 (24-06-08) - [14. Cost Markup In Invoicing] ******
' ***** Manali 3.8.0 - DmCtg and DmCd added
Public Function GetMarkUpRt(ByVal mCmCd As String, ByVal mRmCtg As String, ByVal mSubCtg As String, ByVal mStkRt As Single, _
                            ByVal mCmCurCd As String, ByVal mDmCtg As String, ByVal mDmCd As String, ByVal mIsCustRmCd As Boolean) As Single

Dim wo_rsMrkUp As MwfLib.MDORowSet, ws_MrkUpStr As String, ws_RoundOff As String
Dim ws_MrkUpTc As String
ws_MrkUpTc = IIF(mIsCustRmCd, "IMC", "IMK")

' ***** Manali 3.6.1 - 15/02/10 - CmCurCd added
' ***** Manali 3.8.0 - DmCd, DmCtg added
ws_MrkUpStr = "Select Top 1 RrXWtPer as qPrcnt, RrSalRt as qRndOff, RrFixMinTol as qMinMrkUp, RrFixMaxTol as qMaxMrkUp, " + _
                            " (" + CStr(mStkRt) + "*RrXWtPer/100) as qMarkUp " + _
                            " From RmRt " + _
                            " Where RrTcTyp='" + ws_MrkUpTc + "' And RrCmCd In ('" + mCmCd + "','" + ctSelfCmCd + "') and RrCmCurCd='" + mCmCurCd + "' " + _
                            " And RrDmCtg In ('" + mDmCtg + "', '') And RrDmCD In ('" + mDmCd + "', '') " + _
                            "     And RrCtg='" + mRmCtg + "' And RrSCtg='" + mSubCtg + "' And " + CStr(mStkRt) + " Between RrFrLn And RrToLn " + _
                            " Order By (Case When RrCmCd <> '" + ctSelfCmCd + "' Then 1 Else 2 End), (Case When RrDmCtg <> '' Then 1 Else 2 End), (Case When RrDmCd <> '' Then 1 Else 2 End)   "
Set wo_rsMrkUp = goCn.OpenResultset(ws_MrkUpStr)
With wo_rsMrkUp
    If Not (.EOF And .BOF) Then
        If !qRndOff = 1 Then
           ws_RoundOff = "RNDUP1"
        ElseIf !qRndOff = 5 Then
           ws_RoundOff = "RNDUP5"
        ' ***** Manali 3.6.0 - 28/10/09 - IMK for Accessories - new Round Of Option Added
        ElseIf !qRndOff = 0.5 Then
           ws_RoundOff = "RNDUP.5"
        ElseIf !qRndOff = 0.05 Then
           ws_RoundOff = "RNDUP.05"
        ' ***** Manali 3.6.0 - 28/10/09 - IMK for Accessories - new Round Of Option Added
        Else
            ws_RoundOff = "NORND"
        End If
        GetMarkUpRt = FnRndOff(!qMarkUp, 2, ws_RoundOff)
        GetMarkUpRt = IIF(!qMinMrkUp <> 0 And !qMinMrkUp > GetMarkUpRt, !qMinMrkUp, IIF(!qMaxMrkUp <> 0 And !qMaxMrkUp < GetMarkUpRt, !qMaxMrkUp, GetMarkUpRt))
    Else
        GetMarkUpRt = 0
    End If
 End With
End Function
'****** Sachin 3.3.0 (24-06-08) - [14. Cost Markup In Invoicing] ******

'****** Sachin 3.3.0 - Excel Import Functions *******************
Public Function IsExcelFileReadOnly() As Boolean
     IsExcelFileReadOnly = ms_xlWBook.ReadOnly
End Function

' ***** Manali 3.9.0 - Following function modified to accept Excel 2007 (.xlsx) files
Public Function OpenExcelFile(ByVal psFileName As String) As Boolean
On Error GoTo ErrHdlr
    Set ms_objExcel = CreateObject("Excel.Application.14")
    If ms_objExcel Is Nothing Then Set ms_objExcel = CreateObject("Excel.Application.12")
    If ms_objExcel Is Nothing Then Set ms_objExcel = CreateObject("Excel.Application")
    If ms_objExcel Is Nothing Then OpenExcelFile = False: Exit Function
    
    If ms_objExcel.version < "12.0" And UCase(Right(psFileName, 4)) <> UCase(".xls") Then
      OpenExcelFile = False: Exit Function
    End If
    Set ms_xlWBook = ms_objExcel.Workbooks.Open(psFileName)
    OpenExcelFile = True
  Exit Function
ErrHdlr:
  'If ms_objExcel Is Nothing Or ms_objExcel Is Empty Then MsgBox "EMPTY"
  If Err.Number = 429 Or Err.Number = 424 Then Resume Next
  OpenExcelFile = False
  DispMsg Err.Description, etError
End Function

' Bef 3.9.0
Public Function OpenExcelFile_Bef390(ByVal psFileName As String)
    Set ms_objExcel = CreateObject("Excel.Application")
    Set ms_xlWBook = ms_objExcel.Workbooks.Open(psFileName)
End Function

Public Function ChkSheet(ByVal Sheet As Integer) As Boolean
On Error GoTo ErrHdlr
  ChkSheet = IIF(Sheet > ms_xlWBook.worksheets.Count, False, True)
  Exit Function
ErrHdlr:
  DispMsg Err.Description, etError
End Function

Public Function ReadCell(ByVal Sheet As Integer, ByVal Row As Integer, ByVal Col As Integer) As String
  On Error GoTo ErrHdlr
   ReadCell = ms_xlWBook.worksheets(Sheet).Cells(Row, Col).Value
  Exit Function
ErrHdlr:
  DispMsg Err.Description, etError
End Function

Public Sub SetCellVal(ByVal Sheet As Integer, ByVal Row As Integer, ByVal Col As Integer, ByVal Val As String)
  On Error GoTo ErrHdlr
   ms_xlWBook.worksheets(Sheet).Cells(Row, Col) = Val
  Exit Sub
ErrHdlr:
  DispMsg Err.Description, etError
End Sub

Public Function CloseExcel()
  On Error GoTo ErrHdlr
    ms_xlWBook.Close savechanges:=True  'False
    ms_objExcel.QUIT
    Set ms_objExcel = Nothing
    Set ms_xlWBook = Nothing
  Exit Function
ErrHdlr:
  DispMsg Err.Description, etError
End Function

Public Sub SetModuleSeq()
'****** EVERY ADDITIONAL MODULE'S HEAD FIELD HAS TO ADDED HERE ******
'****** Sachin 3.3.0 - (26-06-08) - [16. Modules and Ini Check Integrity] ******
  gs_ModulesStr(0) = "HLotNoYN"               'Lot Number Module
  gs_ModulesStr(1) = "hFluteBagYN"            'Flute Bag Module
  gs_ModulesStr(2) = "HSlvModYN"              'Silver Module
  gs_ModulesStr(3) = "hCompBagsYn"            'Component Bags Module
  gs_ModulesStr(4) = "HLclInvYN"              'Local Invoicing Module
  gs_ModulesStr(5) = "HFgQuotYN"              'FG Quotation Module
  gs_ModulesStr(6) = "HFgTagMod"              'FG Tag Printing Module
  gs_ModulesStr(7) = "HAutoCustmMtch"         'Automated Custom Matching Module
  gs_ModulesStr(8) = "hBestUse"               'Best Use In Auto. Custom Matching Module
  gs_ModulesStr(9) = "hMultiCurYN"            'Multi-Currency Module
  gs_ModulesStr(10) = "hAutomateReturns"      'Repair-Return Module
  gs_ModulesStr(11) = "hImpComLoc"            'Multi-Company Module
  gs_ModulesStr(12) = "HCustomiseOrd"         'IDE-Customize Order Module
  gs_ModulesStr(13) = "hConsolidatedInv"      'Consolidated Invoicing Reports Module
  gs_ModulesStr(14) = "hLooseInv"             'Loose Invoicing Module
  gs_ModulesStr(15) = "hQWModYn"              'QW Module YN
  gs_ModulesStr(16) = "hInRepMetLsRmCd"       'Invoice Report Metal Loss By RmCd Module
  gs_ModulesStr(17) = "hExcelImpInOrdYN"      'Enable Import In Order From Excel
  gs_ModulesStr(18) = "HPdModYN"              'Palladium Module [Manali]
  gs_ModulesStr(19) = "hPpcYN"                'Production Planning Module
  gs_ModulesStr(20) = "HCtbMod"               'Contractor Module
  gs_ModulesStr(21) = "hTravelExlYN"          'Travel Excel
  gs_ModulesStr(22) = "hFgIOYn"               'Finish Goods IN-Out  '*** Jay 3.4(FG)
  gs_ModulesStr(23) = "HDsgCatYN"             'Design Catalogue Module - Manali 3.4.1
  gs_ModulesStr(24) = "HOrdExcelRepYn"        'Order Printing In Excel Format - Manali 3.6.0
  gs_ModulesStr(25) = "HRfIdMod"              'RfId Module - Manali 3.9.1
  gs_ModulesStr(26) = "HMultiPrcQtMod"        'Multi Price Quotation Module - Manali 3.10.0 - 04/04/12
  gs_ModulesStr(27) = "hRtfrRtChrtYN"         'Refresh Rate Module
  gs_ModulesStr(28) = "hCstRtEdtYn"           'Cost Rate Editable Module - Bhavna
  gs_ModulesStr(29) = "hRwrkModYn"            'Rework Module y/n - Bhavna
  
'****** Sachin 3.3.0 - (26-06-08) - [16. Modules and Ini Check Integrity] ******
End Sub

' ***** Manali 3.6.1 - Microway Name in Report
Public Sub SetMwName(pRep As CRAXDRT.Report)
  pRep.TxtMwRepFt.SetText "Emperor System by Microway (www.microwaysoftware.com)"
  pRep.TxtMwRepFt.HorAlignment = crRightAlign
  pRep.TxtMwRepFt.Font.Size = 8
  pRep.TxtMwRepFt.Font.Italic = True
  pRep.TxtMwRepFt.TextColor = RGB(128, 128, 128)    ' Gray
End Sub

Public Function ShowRepLogo() As Boolean

  ShowRepLogo = goCn.RecSeek("Select 'x' From Head Where HCoCd='" + gs_CoCd + "' and HCd='" + ctSelfCmCd + "' " + _
                  "and HShowRepLogo='Y' ")
End Function

Public Function ChkTolerance(ByVal pBCoCd As String, ByVal pBYy As String, ByVal pBChr As String, ByVal pBNo As Single, _
                             ByVal pRmCtg As String, ByVal pRmSCtg As String, ByVal pMinChk As Boolean, _
                             ByVal pTdTc As String, ByVal pTdYy As String, ByVal pTdChr As String, ByVal pTdNo As Single, _
                             ByVal pTdSr As Single, ByVal pTdSrNo As Single, ByVal pAddWt As Single, ByVal pAddQty As Single, ByVal pChkWt As Boolean, ByVal pChkQty As Boolean, Optional ByVal pTolTyp As String = "", Optional ByVal SpclParamTolChk As Boolean = False, Optional ByVal SubShp As String = "", Optional ByVal PrdTyp As String = "") As String

'When we pass SpclParamTolChk it will consider SubCtg for RmCtg = 'G P S L' as well
'This was originally not considered after introduction of PTyp = 'CHKTOL' for IG we want RmSCtg to be checked separately for GPSL as well
'depending on user defined options in the param
'Also added was SubShp and WS HS check for CHKTOL param so just select from OrdRm has to vary for that

'This Check is for both Transactions and Finished Goods
'***********************************************************************************
If IsMissing(pTolTyp) Or pTolTyp = "" Then pTolTyp = "WTT"
If Not SpclParamTolChk Then
  If (pRmCtg = "G" Or pRmCtg = "P" Or pRmCtg = "S" Or pRmCtg = "L") Then pRmSCtg = ""
End If
If pTolTyp = "WTT" And Not goCn.RecSeek(" Select 'x' from RmRt Where RrTcTyp In ('WTT','QTT') And RrCtg='" + pRmCtg + "' And RrSCtg='" + pRmSCtg + "' ") Then Exit Function

Dim wOrdRs As MDORowSet, wTolWtDefRs As MDORowSet, wTolQtyDefRs As MDORowSet, wActRs As MDORowSet
Dim ws_hTolWtPctDC As String, ws_hTolQtyPctDC As String, wWtPct As Boolean, wQtyPct As Boolean, wQtyWtStr As String
Dim wMinTolWt As Double, wMaxTolWt As Double, wMinTolQty As Double, wMaxTolQty As Double

Dim wRsTol As MDORowSet
Set wRsTol = goCn.OpenRes("Select hTolWtPctDC, hTolQtyPctDC From Head Where HCd='" + ctSelfCmCd + "' And HCoCd='" + pBCoCd + "'")
ws_hTolWtPctDC = wRsTol!hTolWtPctDC
ws_hTolQtyPctDC = wRsTol!hTolQtyPctDC
Set wRsTol = Nothing


wMinTolWt = -99999: wMaxTolWt = -99999: wMinTolQty = -99999: wMaxTolQty = -99999: wWtPct = False
If (InStr(1, "," + ws_hTolWtPctDC + ",", "," + UCase(pRmCtg) + ",", vbTextCompare) > 0) Or pRmCtg = "G" Or pRmCtg = "P" Or pRmCtg = "S" Or pRmCtg = "L" Then wWtPct = True
If InStr(1, "," + ws_hTolQtyPctDC + ",", "," + UCase(pRmCtg) + ",", vbTextCompare) > 0 Then wQtyPct = True
If pTolTyp = "WXT" Or pTolTyp = "WCT" Then
  wWtPct = True
End If
  
If goCn.GetFldVal("Select BOdTc from Bag where BCoCd='" + UCase(pBCoCd) + "' and BYy='" + pBYy + "' and BChr ='" + pBChr + "' and BNo=" + CStr(pBNo)) = "OS" Then
  Exit Function
End If

If pTolTyp = "WXT" Then
     Set wOrdRs = goCn.OpenResultset("Select OdYy, OdTc, OdChr, OdNo, OdSr, OdOmCmCd as qCmCd, DmCtg as qDmCtg,OdDmCd as qDmCd,BQty as qBQty, DmWaxWt as qPrdWt, 0 as qPrdQty, " + _
                            "   (Case When (OdIWtEqOrd='Y' Or OdFixPrc='Y' Or (OdIWtEqOrd='N' and ','+IsNull(PDesc225,'')+',' Like '%,G,P,S,L,%')) Then 'Y' Else 'N' End) as qFixYN, 0 as OdMinDiaTolWt, 0 as OdMaxDiaTolWt " + _
                            " From OrdDsg Join Bag On BCoCd=OdCoCd And BOdTc=OdTc And BOdYy=OdYy And BOdChr=OdChr And BOdNo=OdNo and BOdSr=OdSr " + _
                            "   Join DsgMst On DmTcTyp=OdDmTcTyp And DmCd=OdDmCd And DmSz='' " + _
                            "   Left Outer Join Param On  PTyp='IWTFRORD' and PMCd=OdIWtFrOrd    " + _
                            " Where BCoCd='" + UCase(pBCoCd) + "' And BYy='" + pBYy + "' And BChr='" + pBChr + "' And BNo=" + CStr(pBNo))
ElseIf pTolTyp = "WCT" Then
      Set wOrdRs = goCn.OpenResultset("Select OdYy, OdTc, OdChr, OdNo, OdSr, Max(OdOmCmCd) as qCmCd, Max(DmCtg) as qDmCtg, Max(OdDmCd) as qDmCd,Max(BQty) as qBQty," + _
                            " Cast(Max(DmCasPcWt)* Max(OrPrdWt)/Sum(DrPrdWt) as Decimal(10,3))  as qPrdWt, " + _
                            "  0 as qPrdQty, " + _
                            "   Max((Case When (OdIWtEqOrd='Y' Or OdFixPrc='Y' Or (OdIWtEqOrd='N' and ','+IsNull(PDesc225,'')+',' Like '%,G,P,S,L,%')) Then 'Y' Else 'N' End)) as qFixYN, 0 as OdMinDiaTolWt,0 as OdMaxDiaTolWt " + _
                            " From OrdDsg Join Bag On BCoCd=OdCoCd And BOdTc=OdTc And BOdYy=OdYy And BOdChr=OdChr And BOdNo=OdNo and BOdSr=OdSr " + _
                            " Join (Select OrCoCd,OrTc,OrYy,OrChr,OrNo,OrSr, Sum(OrPrdWt) as OrBaseWt, Sum(OrPrdWt) as OrPrdWt " + _
                            "       From OrdRm Join RmMst On RmCd=OrRmCd " + _
                            "       where RmCtg In ('G','P','S','L') Group By OrCoCd,OrTc,OrYy,OrChr,OrNo,OrSr ) Rm1  " + _
                            " On OdCoCd=OrCoCd And OdTc=OrTc And OdYy=OrYy And OdChr=OrChr And OdNo=OrNo and OdSr=OrSr " + _
                            "  Join DsgMst On DmTcTyp=OdDmTcTyp And DmCd=OdDmCd And DmSz='' " + _
                            "  Join DsgRm On DrTcTyp=DmTcTyp and DrCd=DmCd and DrSz=DmSz " + _
                            "  Join RmMst On RmCd=DrRmCd " + _
                            "  Left Outer Join Param On  PTyp='IWTFRORD' and PMCd=OdIWtFrOrd    " + _
                            " Where BCoCd='" + UCase(pBCoCd) + "' And BYy='" + pBYy + "' And BChr='" + pBChr + "' And BNo=" + CStr(pBNo) + " and RmCtg in ('G','P','S','L')" + _
                            " Group By OdCoCd,OdTc, OdYy, OdChr, OdNo, OdSr")

Else
    Set wOrdRs = goCn.OpenResultset("Select OdYy, OdTc, OdChr, OdNo, OdSr, Max(OdOmCmCd) as qCmCd, Max(DmCtg) as qDmCtg, Max(OdDmCd) as qDmCd,Max(BQty) as qBQty, Sum(OrPrdWt) as qPrdWt, Sum(OrPrdQty) as qPrdQty, " + _
                            "   Max((Case When (OdIWtEqOrd='Y' Or OdFixPrc='Y' Or (OdIWtEqOrd='N' and ','+IsNull(PDesc225,'')+',' Like '%," + pRmCtg + ",%')) Then 'Y' Else 'N' End)) as qFixYN, Max(OdMinDiaTolWt) as OdMinDiaTolWt, Max(OdMaxDiaTolWt) as OdMaxDiaTolWt " + _
                            " From Ordrm " + _
                            "   Join OrdDsg On OdCoCd=OrCoCd And OdTc=OrTc And OdYy=OrYy And OdChr=OrChr And OdNo=OrNo And OdSr=OrSr " + _
                            "   Join Bag On BCoCd=OdCoCd And BOdTc=OdTc And BOdYy=OdYy And BOdChr=OdChr And BOdNo=OdNo and BOdSr=OdSr " + _
                            "   Join DsgMst On DmTcTyp=OdDmTcTyp And DmCd=OdDmCd And DmSz='' " + _
                            "   Left Outer Join Param On  PTyp='IWTFRORD' and PMCd=OdIWtFrOrd    " + _
                            " Where BCoCd='" + UCase(pBCoCd) + "' And BYy='" + pBYy + "' And BChr='" + pBChr + "' And BNo=" + CStr(pBNo) + _
                            "   And OrRmCtg='" + pRmCtg + "'  " + IIF(pRmSCtg <> "", "And OrRmSCtg='" + pRmSCtg + "' ", "") + _
                            IIF(SubShp <> "", "And OrSubShp='" + SubShp + "' ", "") + _
                            IIF(PrdTyp = "WS", "And OrWsQty>0 ", "") + _
                            IIF(PrdTyp = "GS", "And OrHsQty>0 ", "") + _
                            " Group by OdCoCd, OdTc, OdYy, OdChr, OdNo, OdSr, OrRmCtg " + IIF(pRmSCtg <> "", ",OrRmSCtg", ""))
                            
End If
If wOrdRs.RecCount <= 0 Then Exit Function

Dim wOmCmCurCd As String
wOmCmCurCd = goCn.GetFldVal("Select OmCmCurCd from OrdMst where OmCoCd ='" + UCase(pBCoCd) + "' and OmYy ='" + wOrdRs!OdYy + "' and OmTc='" + wOrdRs!OdTc + "' and OmChr='" + wOrdRs!OdChr + "' and OmNo=" + CStr(wOrdRs!OdNo))

Set wTolWtDefRs = goCn.OpenResultset(" Select Top 1 RrFixMinTol, RrFixMaxTol, RrOpnMinTol, RrOpnMaxTol " + _
                             " From RmRt " + _
                            " Where (RrCmCurCd='" + wOmCmCurCd + "' Or RrCmCurCd='') and " + _
                            " RrTcTyp='" + pTolTyp + "' And RrCtg='" + pRmCtg + "' and RrCmCtg='C' " + _
                            "   And RrCmCd In ('" + wOrdRs!qCmCd + "','ZSELF') And RrSCtg ='" + pRmSCtg + "' And RrDmCtg in ('" + wOrdRs!qDmCtg + "','') And RrDmCd In ('" + wOrdRs!qDmCd + "','') " + _
                            "   And round(" + CStr(wOrdRs!qPrdWt) + ",3) between round(RrFrLn,3) and round(RrToLn,3) " + _
                            " Order By (Case when RrCmCd='" + wOrdRs!qCmCd + "' Then 1 Else 2 End),(Case when RrCmCurCd<>'' Then 1 else 2 end), (Case when RrDmCtg='" + wOrdRs!qDmCtg + "' Then 1 Else 2 End), " + _
                            "          (Case when RrDmCd='" + wOrdRs!qDmCd + "' Then 1 Else 2 End) ")

Set wTolQtyDefRs = goCn.OpenResultset(" Select Top 1 RrFixMinTol, RrFixMaxTol, RrOpnMinTol, RrOpnMaxTol " + _
                            " From RmRt " + _
                            " Where (RrCmCurCd='" + wOmCmCurCd + "' Or RrCmCurCd='') and " + _
                            " RrTcTyp='QTT' And RrCtg='" + pRmCtg + "' and RrCmCtg='C' " + _
                            "   And RrCmCd In ('" + wOrdRs!qCmCd + "','ZSELF') And RrSCtg ='" + pRmSCtg + "' And RrDmCtg in ('" + wOrdRs!qDmCtg + "','') And RrDmCd In ('" + wOrdRs!qDmCd + "','') " + _
                            "   And " + CStr(wOrdRs!qPrdQty) + " between RrFrLn and RrToLn " + _
                            " Order By (Case when RrCmCd='" + wOrdRs!qCmCd + "' Then 1 Else 2 End),(Case when RrCmCurCd<>'' Then 1 else 2 end),  (Case when RrDmCtg='" + wOrdRs!qDmCtg + "' Then 1 Else 2 End), " + _
                            "          (Case when RrDmCd='" + wOrdRs!qDmCd + "' Then 1 Else 2 End) ")
                            
If wTolWtDefRs.RecCount <= 0 And wTolQtyDefRs.RecCount <= 0 And Not (UCase(pRmCtg) = "D" And UCase(Trim(pRmSCtg)) = "") Then Exit Function

If wTolWtDefRs.RecCount > 0 Then
  wMinTolWt = IIF(wOrdRs!qFixYn = "Y", wTolWtDefRs!RrFixMinTol, wTolWtDefRs!RrOpnMinTol)
  wMaxTolWt = IIF(wOrdRs!qFixYn = "Y", wTolWtDefRs!RrFixMaxTol, wTolWtDefRs!RrOpnMaxTol)
  wMinTolWt = IIF(wWtPct, wOrdRs!qPrdWt * wMinTolWt / 100, wMinTolWt)
  wMaxTolWt = IIF(wWtPct, wOrdRs!qPrdWt * wMaxTolWt / 100, wMaxTolWt)
  wMinTolWt = wOrdRs!qPrdWt - wMinTolWt: wMaxTolWt = wOrdRs!qPrdWt + wMaxTolWt
End If

If UCase(pRmCtg) = "D" And UCase(Trim(pRmSCtg)) = "" Then
    If wOrdRs!OdMinDiaTolWt <> 0 Then wMinTolWt = wOrdRs!OdMinDiaTolWt
    If wOrdRs!OdMaxDiaTolWt <> 0 Then wMaxTolWt = wOrdRs!OdMaxDiaTolWt
End If

If wTolQtyDefRs.RecCount > 0 Then
  wMinTolQty = IIF(wOrdRs!qFixYn = "Y", wTolQtyDefRs!RrFixMinTol, wTolQtyDefRs!RrOpnMinTol)
  wMaxTolQty = IIF(wOrdRs!qFixYn = "Y", wTolQtyDefRs!RrFixMaxTol, wTolQtyDefRs!RrOpnMaxTol)
  wMinTolQty = IIF(wQtyPct, wOrdRs!qPrdQty * wMinTolQty / 100, wMinTolQty)
  wMaxTolQty = IIF(wQtyPct, wOrdRs!qPrdQty * wMaxTolQty / 100, wMaxTolQty)
  wMinTolQty = wOrdRs!qPrdQty - wMinTolQty: wMaxTolQty = wOrdRs!qPrdQty + wMaxTolQty
End If
  
If pTolTyp = "WXT" Then
    Set wActRs = goCn.OpenResultset("Select  IsNull(Sum(TdWaxWt),0)+ " + CStr(pAddWt) + "  as qTdRmWt,0 as qTdRmQty " + _
                                " From Txnd Where TdCoCd = '" + pBCoCd + "' " + _
                                "   And TdBYy='" + pBYy + "' And TdBChr='" + pBChr + "' And TdBNo=" + CStr(pBNo) + _
                                "   And Not(TdTc='" + pTdTc + "' And TdYy='" + pTdYy + "' And TdChr='" + pTdChr + "' And TdNo=" + CStr(pTdNo) + " And TdSr=" + CStr(pTdSr) + " and TdSrNo=" + CStr(pTdSrNo) + ") ")
                                
ElseIf pTolTyp = "WCT" Then
  Set wActRs = goCn.OpenResultset("Select IsNull(sum(TdRmWt* (case when TdFrRmLoc= 'W' and TdFrRmDC= 'C' then -1 " + _
                                "   when TdFrRmLoc= 'W' and TdFrRmDC= 'D' then 1 " + _
                                "   when TdToRmLoc= 'W' and TdToRmDC= 'C' then -1 " + _
                                "   when  TdToRmLoc='W' and TdToRmDC= 'D' then 1 else 0 end)),0)+" + CStr(pAddWt) + " as qTdRmWt, " + _
                                " IsNull(sum(TdRmQty* (case when TdFrRmLoc= 'W' and TdFrRmDC= 'C' then -1 " + _
                                "   when TdFrRmLoc= 'W' and TdFrRmDC= 'D' then 1 " + _
                                "   when TdToRmLoc= 'W' and TdToRmDC= 'C' then -1 " + _
                                "   when  TdToRmLoc='W' and TdToRmDC= 'D' then 1 else 0 end)),0)+" + CStr(pAddQty) + " as qTdRmQty " + _
                                " From Txnd Where TdRmCtg in ('G','P','S','L') And TdCoCd = '" + pBCoCd + "' " + _
                                "   And TdBYy='" + pBYy + "' And TdBChr='" + pBChr + "' And TdBNo=" + CStr(pBNo) + _
                                "   And Not (TdTc='" + pTdTc + "' And TdYy='" + pTdYy + "' And TdChr='" + pTdChr + "' And TdNo=" + CStr(pTdNo) + " And TdSr=" + CStr(pTdSr) + " and TdSrNo=" + CStr(pTdSrNo) + ") ")
    
Else
    Set wActRs = goCn.OpenResultset("Select IsNull(sum(TdRmWt* (case when TdFrRmLoc= 'W' and TdFrRmDC= 'C' then -1 " + _
                                "   when TdFrRmLoc= 'W' and TdFrRmDC= 'D' then 1 " + _
                                "   when TdToRmLoc= 'W' and TdToRmDC= 'C' then -1 " + _
                                "   when  TdToRmLoc='W' and TdToRmDC= 'D' then 1 else 0 end)),0)+" + CStr(pAddWt) + " as qTdRmWt, " + _
                                " IsNull(sum(TdRmQty* (case when TdFrRmLoc= 'W' and TdFrRmDC= 'C' then -1 " + _
                                "   when TdFrRmLoc= 'W' and TdFrRmDC= 'D' then 1 " + _
                                "   when TdToRmLoc= 'W' and TdToRmDC= 'C' then -1 " + _
                                "   when  TdToRmLoc='W' and TdToRmDC= 'D' then 1 else 0 end)),0)+" + CStr(pAddQty) + " as qTdRmQty " + _
                                " From Txnd Where TdRmCtg='" + pRmCtg + "'   " + IIF(pRmSCtg <> "", " And TdRmSCtg='" + pRmSCtg + "' ", "") + " And TdCoCd = '" + pBCoCd + "' " + _
                                "   And TdBYy='" + pBYy + "' And TdBChr='" + pBChr + "' And TdBNo=" + CStr(pBNo) + _
                                "   And Not (TdTc='" + pTdTc + "' And TdYy='" + pTdYy + "' And TdChr='" + pTdChr + "' And TdNo=" + CStr(pTdNo) + " And TdSr=" + CStr(pTdSr) + " and TdSrNo=" + CStr(pTdSrNo) + ") ")
End If
If wActRs.RecCount <= 0 Then Exit Function
   
'Dim wErrStr As String
'If wActRs.RecCount > 0 And wOrdRs!qBQty > 0 Then
'      If wActRs!qTdRmWt = 0 Then Exit Function
'      If pMinChk And wMinTolWt > -1 And pChkWt Then _
'            If wActRs!qTdRmWt / wOrdRs!qBQty < wMinTolWt Then wErrStr = "[Min Weight]"
'      If wMaxTolWt > -1 And pChkWt Then _
'            If wActRs!qTdRmWt / wOrdRs!qBQty > wMaxTolWt Then wErrStr = wErrStr + "[Max Weight]"
'      If pMinChk And wMinTolQty > -1 And pChkQty Then _
'            If wActRs!qTdRmQty / wOrdRs!qBQty < wMinTolQty Then wErrStr = wErrStr + "[Min Quantity]"
'      If wMaxTolQty > -1 And pChkQty Then _
'            If wActRs!qTdRmQty / wOrdRs!qBQty > wMaxTolQty Then wErrStr = wErrStr + "[Max Quantity]"
'End If
'
'If pTolTyp = "WXT" Or pTolTyp = "WCT" Then
'    If wErrStr <> "" Then ChkTolerance = IIF(pTolTyp = "WXT", "Wax Pullout Weight", " Cast Weight ") + "Tolerance Check Fails in " + wErrStr
'Else
'    If wErrStr <> "" Then ChkTolerance = "Bag Tolerance Check Fails in " + wErrStr + " for " + pRmCtg + IIF(pRmSCtg <> "", "(" + pRmSCtg + ")", "")
'End If

Dim wErrStr As String, wDetErrMsg As String
If wActRs.RecCount > 0 And wOrdRs!qBQty > 0 Then
      'If wActRs!qTdRmWt = 0 Then Exit Function
      If pMinChk And wMinTolWt <> -99999 And pChkWt Then
            If wActRs!qTdRmWt / wOrdRs!qBQty < wMinTolWt Then
                    wErrStr = "[Min Wt]"
                    wDetErrMsg = " Min Wt. Allowed : " + CStr(wMinTolWt) + ", Entered Wt : " + CStr(wActRs!qTdRmWt / wOrdRs!qBQty) + vbCrLf
            End If
      End If
            
      If wMaxTolWt <> -99999 And pChkWt Then
            If wActRs!qTdRmWt / wOrdRs!qBQty > wMaxTolWt Then
                wErrStr = wErrStr + "[Max Wt]"
                wDetErrMsg = wDetErrMsg + " Max Wt. Allowed : " + CStr(wMaxTolWt) + ", Entered Wt : " + CStr(wActRs!qTdRmWt / wOrdRs!qBQty) + vbCrLf
            End If
       End If
            
      If pMinChk And wMinTolQty <> -99999 And pChkQty Then
            If wActRs!qTdRmQty / wOrdRs!qBQty < wMinTolQty Then
                wErrStr = wErrStr + "[Min Qty]"
                wDetErrMsg = wDetErrMsg + " Min Qty. Allowed : " + CStr(wMinTolQty) + ", Entered Qty : " + CStr(wActRs!qTdRmQty / wOrdRs!qBQty) + vbCrLf
            End If
      End If
      
      If wMaxTolQty > -99999 And pChkQty Then
            If wActRs!qTdRmQty / wOrdRs!qBQty > wMaxTolQty Then
                wErrStr = wErrStr + "[Max Qty]"
                wDetErrMsg = wDetErrMsg + " Max Qty. Allowed : " + CStr(wMaxTolQty) + ", Entered Qty : " + CStr(wActRs!qTdRmQty / wOrdRs!qBQty) + vbCrLf
            End If
      End If
End If

If pTolTyp = "WXT" Or pTolTyp = "WCT" Then
    'If wErrStr <> "" Then ChkTolerance = IIF(pTolTyp = "WXT", "Wax Pullout Weight", " Cast Weight ") + "Tolerance Check Fails in " + wErrStr
    If wErrStr <> "" Then ChkTolerance = IIF(pTolTyp = "WXT", "Wax Pullout Weight", " Cast Weight ") + " Tolerance for " + wErrStr + " Failed for : " + vbCrLf + _
            " RmCtg : " + pRmCtg + IIF(pRmSCtg <> "", ", RmSCtg : " + pRmSCtg, "") + vbCrLf + wDetErrMsg
Else
    If wErrStr <> "" Then ChkTolerance = " Bag Level Tolerance for " + wErrStr + " Failed for : " + vbCrLf + _
            " RmCtg : " + pRmCtg + IIF(pRmSCtg <> "", ", RmSCtg : " + pRmSCtg, "") + vbCrLf + wDetErrMsg
End If

End Function

Public Function ReplXlImg(pExcelPath As String, prs_Pic As MDORowSet, pPicHeight As Single, pPicWidth As Single) As Boolean
On Error GoTo errhld

 Dim oExcel, oBook, oSheet, wSp, sp
 Dim wPictPath As String
 Dim wDmTcTyp As String
 Dim t, l, h, W As Integer
 Dim sFile As String
 Dim wOutputPath As String, wPos As Integer, wExtn As String, sFilePath As String

 ReplXlImg = False
 If prs_Pic Is Nothing Then Exit Function
 
 Set oExcel = GetExcelObj()                         'Sachin 4.4.2
 'Set oExcel = CreateObject("Excel.Application")
 If oExcel.version < "12.0" Then DispMsg "This Option is only available for Excel 2007 or above", etError: Exit Function
   
 wPos = InStr(pExcelPath, ".")
 wExtn = Mid(pExcelPath, wPos, Len(pExcelPath))
 sFilePath = Replace(pExcelPath, wExtn, "_" + Format(Now, DtFullStr) + wExtn)   '6.1
 
 sFilePath = Replace(pExcelPath, ".xls", ".xlsx")
 If Dir(sFilePath, vbNormal) <> "" Then Kill sFilePath
  
 Set oBook = oExcel.Workbooks.Open(pExcelPath)
 oExcel.ActiveWorkbook.SaveAs sFilePath, FileFormat:=51     'Save in .xlsx format (as Crystal is damaging the .xls file while exporting)

 If Dir(pExcelPath, vbNormal) <> "" Then Kill pExcelPath    'Delete Original Excel File (created by Crystal Reports export)

 Set oBook = oExcel.Workbooks.Open(sFilePath)
 Set oSheet = oBook.worksheets(1)
    If prs_Pic.RecCount <> oSheet.Shapes.Count Then DispMsg "Scope and File Images Mismatch. Cannot Continue.", etError: GoTo errhld 'Exit Function
    If prs_Pic.RecCount = 0 Then Exit Function
    
    prs_Pic.MoveFirst
    For Each sp In oSheet.Shapes
       sp.Select
       With sp
        t = .Top
        l = .Left
        'h = .Height
        'w = .Width
    End With
    'sp.Delete
        With prs_Pic
            wDmTcTyp = IIF(prs_Pic!DmTcTyp <> "", prs_Pic!DmTcTyp, IIF(prs_Pic!OmTc = "QS", "SM", "DM"))
            wPictPath = GetPictPath(wDmTcTyp, IIF(wDmTcTyp <> "DM", True, False), !DmCd, !DmSz, !OmCoCd, !OmTc, !OmYy, !OmChr, !OmNo, !OdSr, !PicNm, , True)
                If Dir(wPictPath) = "" Then wPictPath = GetPictPath(wDmTcTyp, IIF(wDmTcTyp <> "DM", True, False), !DmCd, !DmSz, !OmCoCd, !OmTc, !OmYy, !OmChr, !OmNo, !OdSr, !PicNm)
                    If wPictPath <> "" And Dir(wPictPath) <> "" Then
                        sp.Delete
                        Set wSp = oSheet.Shapes.AddPicture(wPictPath, False, True, l, t, pPicWidth, pPicHeight)
                        'wSp.ShapeRange.LockAspectRatio = True
                        wSp.Line.Visible = True
                    End If
               If Not (prs_Pic.BOF Or prs_Pic.EOF) Then prs_Pic.MoveNext
          End With
    Next
    oExcel.Visible = True
    oExcel.DisplayAlerts = True
    Set oExcel = Nothing
    Set oBook = Nothing
    Set oSheet = Nothing
    Set prs_Pic = Nothing
    ReplXlImg = True
Exit Function
errhld:
   If Err.Number = 70 Then       '4.1.4
     DispMsg "Cannot Access Template File. Close If Already Open and Continue.", etError
   Else
     DispMsg Err.Description, etError
   End If
  
   If Not oExcel = Empty Then
    oExcel.QUIT
    oExcel.DisplayAlerts = True
   End If
    Set oExcel = Nothing
    Set oBook = Nothing
    Set oSheet = Nothing
    Set prs_Pic = Nothing
End Function
Public Sub SetStockSizeVisible(SizeNum As Integer, GrdColumn As MWCLS_FLX, Conn As MDOConnection)
  If Conn.RecSeek("Select 'x' From Head Where HCd='" + ctSelfCmCd + "' And HCoCd='" + ctSelfCoCd + "' " + _
                            " and HRmStkBy >= " + CStr(SizeNum)) Then
    GrdColumn.Style = fgcNormal
  Else
    GrdColumn.Style = fgcinvisible
  End If
End Sub
Public Function ColWhenStkSz(SizeNum As Integer, RmCtg As String, RmSCtg As String, ByRef ErrMsg As String) As Boolean
  ErrMsg = "Cannot enter Sz" + CStr(SizeNum) + " for this Rm Sub Category"
  ColWhenStkSz = Not goCn.RecSeek("Select 'x' From Param where PTyp = 'RMSCTG' and PMCd = '" + RmCtg + "' and PSCd = '" + RmSCtg + "' and PNum1 >=" + CStr(SizeNum))
End Function
Public Function ValidateStkSize(Ln1 As Double, Ln2 As Double, Ln3 As Double, ByRef ErrMsg) As Boolean
    Dim Cancel As Boolean
    Dim e As String
    e = CStr(Ln1) + IIF(Ln2 <> 0, "*" + CStr(Ln2), "") + IIF(Ln3 <> 0, "*" + CStr(Ln3), "")
    If Len(e) > 22 Then
      Cancel = True
      ErrMsg = "Sz1 * Sz2 * Sz3 must be less than 21 characters. Current value is: " + e
    End If
  ValidateStkSize = Cancel
End Function
'53 company code parameter added
Public Function GetNextDate(ByVal pCoCd As String, Optional ByVal pDate As String) As String
    Dim wb_HolFlg As Boolean
    wb_HolFlg = False
    If IsMissing(pDate) Then pDate = goCn.SrvrDate
    If pDate = "" Then pDate = goCn.SrvrDate
    
    '6.1_50 if no check for holiday then no need to check date table, it will return server date
    If goCn.GetFldVal("Select hHolChkYN From Head where HCd='ZSELF' and HCoCd='" + gs_CoCd + "'") = "N" Then GetNextDate = pDate: Exit Function
  
 '   If goCn.GetFldVal("Select top 1 hHolChkYN From Head where HCd='ZSELF' and HCoCd in('" + ctSelfCoCd + "','" + pCoCd + "') order by HCoCd desc ") = "N" Then GetNextDate = pDate: Exit Function
    
    '53 company code added
'      wb_HolFlg = goCn.RecSeek("Select top 1 'x' From DtTbl where DtTag='HOL' and DtCoCd in('" + ctSelfCoCd + "','" + pCoCd + "') and convert(varchar(15),DtDt,3)= '" + pDate + "' and DtHolYN='Y' order by DTCoCd asc ")
    Dim wEr As String
    '6.1
    wEr = goCn.GetFldVal("Select DtHolYn From DtTbl where DtTag='HOL' and DtCoCd in('" + ctSelfCoCd + "','" + pCoCd + "') and convert(varchar(15),DtDt," + CStr(DtFmtSlahVal()) + ")= '" + pDate + "' order by dtCoCd asc")

      Do While wEr = "Y"
        pDate = goCn.GetFldVal(" Select DateAdd(dd,1,'" + pDate + "')")
        '53 company code added
        'wb_HolFlg = goCn.RecSeek("Select top 1 'x' From DtTbl where DtTag='HOL' and DtCoCd in('" + ctSelfCoCd + "','" + pCoCd + "') and convert(varchar(15),DtDt,3)= '" + pDate + "' and DtHolYN='Y' order by DTCoCd asc")
        '6.1
        wEr = goCn.GetFldVal("Select DtHolYn From DtTbl where DtTag='HOL' and DtCoCd in('" + ctSelfCoCd + "','" + pCoCd + "') and convert(varchar(15),DtDt," + CStr(DtFmtSlahVal()) + ")= '" + pDate + "' order by dtCoCd asc")
    Loop
    GetNextDate = pDate
End Function
'53 company code parameter added
Public Function IsHOLYN(ByVal pCoCd As String, Optional ByVal pDate As String) As Boolean
'    If goCn.GetFldVal("Select hHolChkYN From Head where HCd='ZSELF' and HCoCd='" + ctSelfCoCd + "'") = "N" Then IsHOLYN = "N = pDate: Exit Function"
    Dim wEr As String
    '6.1
    wEr = goCn.GetFldVal("Select DtHolYn From DtTbl where DtTag='HOL' and DtCoCd in('" + ctSelfCoCd + "','" + pCoCd + "') and convert(varchar(15),DtDt," + CStr(DtFmtSlahVal()) + ")= '" + pDate + "' order by dtCoCd asc")
    If wEr = "" Or wEr = "N" Then
      IsHOLYN = False
    Else
      IsHOLYN = True
    End If
    
End Function
Public Sub GetLockDt(ByVal pCoCd As String, ByVal pTcTyp As String, ByVal pTc As String, ByRef pBackDt As String, ByRef pFutureDt As String, Optional pVchrNo As String)
    Dim wCurrDt As String
    wCurrDt = goCn.GetFldVal("Select convert(varchar(15),GetDate()," + CStr(DtFmtSlahVal()) + ")")  '6.1
    '6.1_51 pVchrNo added as parameter to check that given voucher number is exist in unlock ptyp,
    'if given voucher is in unlock ptyp it will not check lock date
    If pVchrNo <> "" Then _
      If ChkUnLockVch(pCoCd, pVchrNo) = True Then pBackDt = "01/01/80": pFutureDt = IIF(UCase(MWLib.DtSetting) = "MM/DD/YY", "12/31/28", "31/12/28"): Exit Sub
    '6.1
    Dim wLockDtArr() As String
    wLockDtArr = Split(goCn.GetFldVal("Select Str(vPNum)+','+Str(vPNum1) From vParam where vPCoCd = '" + pCoCd + "' and vPTyp='LOCKDT' and vPMCd='" + pTcTyp + "' and vPSCd='" + pTc + "'"), ",")
    If UBound(wLockDtArr) < 0 Then pBackDt = "01/01/80": pFutureDt = IIF(UCase(MWLib.DtSetting) = "MM/DD/YY", "12/31/28", "31/12/28"): Exit Sub
    pBackDt = IIF(wLockDtArr(0) = 999, "01/01/80", goCn.GetFldVal(" Select DateAdd(dd,-1*(" + wLockDtArr(0) + "-1),'" + wCurrDt + "')"))
    pFutureDt = IIF(wLockDtArr(1) = 999, IIF(UCase(MWLib.DtSetting) = "MM/DD/YY", "12/31/28", "31/12/28"), goCn.GetFldVal(" Select DateAdd(dd,(" + wLockDtArr(1) + "-1),'" + wCurrDt + "')"))
End Sub
Public Function IsDateFormatChanged() As Boolean
  If gs_DateFormatAtLogin <> MWLib.GetDateFormat Then
    IsDateFormatChanged = False
  Else
    IsDateFormatChanged = True
  End If
End Function
Public Function IsDateValidAgainstLockDt(ByVal pCoCd As String, ByVal pTcTyp As String, ByVal pTc As String, ByVal pv_NewValue As String) As Boolean
  IsDateValidAgainstLockDt = goCn.GetFldVal("Select 1 where exists (Select * from vParam where VpCoCd='" + pCoCd + "' and vPTyp='LOCKDT' and vPMCd='" + pTcTyp + "' and vPSCd='" + pTc + "'" + " and " + _
                                                                      " (('" + pv_NewValue + "'   < dateadd(dd,-1*vpnum,getdate()))" + _
                                                                      " or ('" + pv_NewValue + "' > dateadd(dd,+1*vpnum1,getdate())))) ")
                                                                        
End Function
Public Function GetExcelObj() As Variant
On Error GoTo ErrHdlr
    Set GetExcelObj = CreateObject("Excel.Application")
    If GetExcelObj Is Nothing Then Set GetExcelObj = CreateObject("Excel.Application.12")
    If GetExcelObj Is Nothing Then Set GetExcelObj = CreateObject("Excel.Application.14")
  Exit Function
ErrHdlr:
  If Err.Number = 429 Or Err.Number = 424 Then Resume Next
  DispMsg Err.Description, etError
End Function
Public Sub UpdIntegrity(Optional ByVal Conn As MDOConnection, Optional ByVal mCoCd As String)
  '****** Sachin 3.3.0 - Updating Integrity field on Update of Modules ******
  Dim rs_intgty As MwfLib.MDORowSet, ws_ModulesFldStr As String, i As Integer, ws_CoCd As String
  On Error GoTo Errhnd
  Dim wtmpconn As MDOConnection
  If Conn Is Nothing Then
    Set wtmpconn = goCn
  Else
    Set wtmpconn = Conn
  End If
  ws_ModulesFldStr = ""
  ws_CoCd = IIF(mCoCd <> "", " Where HChkIntgty='' And hCoCd='" + mCoCd + "'", "")
  Do While i <= UBound(gs_ModulesStr)
    ws_ModulesFldStr = ws_ModulesFldStr + "+" + gs_ModulesStr(i)
    i = i + 1
  Loop
  Set rs_intgty = wtmpconn.OpenResultset("Select HCoCd, HChkIntgty From Head " + ws_CoCd)
  wtmpconn.Execute ("Insert Into TrgDisa Values ('Head')")
'  wtmpconn.BeginTrans (mdoCommitRec)
  If Not (rs_intgty.EOF Or rs_intgty.BOF) Then
    rs_intgty.MoveFirst
    With rs_intgty
      Do While Not rs_intgty.EOF
        .FldValue("HChkIntgty") = Encrypt(wtmpconn.GetFldVal("Select (HCoCd+':'+HName+':'" + ws_ModulesFldStr + ") " + _
                                                         "  From Head " + _
                                                         " Where HCd='" + ctSelfCmCd + "' And HCoCd='" + !hCoCd + "'"), GenCryptNew)
        .Update
        .MoveNext
      Loop
    End With
  End If
 ' wtmpconn.CommitTrans mdoCommitRec
  wtmpconn.Execute ("Delete From TrgDisa")
  Set rs_intgty = Nothing
  '****** Sachin 3.3.0 - Updating Integrity field on Update of Modules ******
  Exit Sub
Errhnd:
  MsgBox ("Error in UpdIntegrity :" + Err.Description)
  wtmpconn.Execute ("Delete From TrgDisa")
End Sub
Public Function ChkCmpUsrAccess(ByVal pCmCtg As String, ByVal pMnuCd As String, ByVal pCmCdFr As String, ByVal pCmCdTo As String, ByVal pCmCdFor As String, ByRef ErrMsg As String) As Boolean
'Indigo.1-Getting valid company code from custmst,checking that this Cocd should match with valid user access CoCd,
'if match then user can access the records else user cant view the records

  Dim ws_CoCd() As String, i As Integer, ws_CmValidCoCd As String
  Dim wCmValidCoCd As MDORowSet
  ChkCmpUsrAccess = True
  ws_CmValidCoCd = ""
  If Not IIF(goCn.GetFldVal("Select hImpComLoc From Head Where HCd='" + ctSelfCmCd + "' And HCoCd='" + ctSelfCoCd + "'") = "Y", True, False) Then
    ChkCmpUsrAccess = True
    Exit Function
  End If
  Set wCmValidCoCd = goCn.OpenRes("Select CmValidCoCd,CmCd from CustMst Where CmCtg ='" + pCmCtg + "'" + _
                        IIF(pCmCdFr <> "", " And CmCd >= '" + pCmCdFr + "'", "") + _
                        IIF(pCmCdTo <> "", " And CmCd <= '" + pCmCdTo + "'", "") + _
                        IIF(pCmCdFor <> "", " And CmCd = '" + pCmCdFor + "'", ""))
                        
  Dim ws_UaValidCoCd As String
  ws_UaValidCoCd = goCn.GetFldVal("Select UaValidCoCd From UsrAccess Where UaMnuCd= '" + pMnuCd + "' And UaUsrCd = '" + gs_UsrCd + "' And UaMnuForm='MDI'")
          
  Do While Not (wCmValidCoCd.BOF Or wCmValidCoCd.EOF)
    ws_CmValidCoCd = wCmValidCoCd!CmValidCoCd
      
    If ws_CmValidCoCd <> "" Then
      ws_CoCd = Split(ws_CmValidCoCd, ",")
      If IsArray(ws_CoCd) Then
        For i = 0 To UBound(ws_CoCd)
          If ws_UaValidCoCd = "" Or "," + ws_UaValidCoCd + "," Like "*," + ws_CoCd(i) + ",*" Then
            ChkCmpUsrAccess = True
            Exit For
          Else
            ChkCmpUsrAccess = False
            ErrMsg = wCmValidCoCd!CmCd
          End If
        Next i
        
        If Not ChkCmpUsrAccess Then
          Exit Function
        End If
      End If
    End If
    wCmValidCoCd.MoveNext
  Loop

End Function
Public Function GetMRD(ByVal mCmCtg As String, ByVal mCmCd As String, ByVal mDmCd As String, _
                          ByVal mDmSz As String, ByVal mDmSfx As String)
    
  Dim wMrd As Double
  Dim wCmLkUpCmCd As String
  
  wCmLkUpCmCd = goCn.GetFldVal("Select CmLkUpMRPDisc from CustMst where CmCtg= '" + mCmCtg + "' and CmCd= '" + mCmCd + "'")
  wMrd = goCn.GetFldVal("Select isnull" + _
             "((select " + "RrSalRt" + "  from RmRt where RrTcTyp= 'MRD' and " + IIF(1 = 1, " RrPrtKey='" + ctCurrPrtn + "' and ", "") + _
             "  RrCmCtg= 'P' and RrCmCd= '" + wCmLkUpCmCd + "' and " + _
               "RrDmCd = '" + mDmCd + "' and RrDmSfx = '" + mDmSfx + "' and RrDmSz=' " + mDmSz + "')" + _
              ", " + _
              "IsNull((Select RrSalRt" + "  from RmRt where RrTcTyp= 'MRD' and " + IIF(1 = 1, " RrPrtKey='" + ctCurrPrtn + "' and ", "") + _
                       "RrCmCtg= 'P' and RrCmCd='" + wCmLkUpCmCd + "' and " + _
                           "RrDmCd= '" + mDmCd + "' and RrDmSfx ='" + mDmSfx + "' and RrDmSz='')" + _
              ", " + _
              "IsNull((Select RrSalRt" + "  from RmRt where RrTcTyp= 'MRD' and " + IIF(1 = 1, " RrPrtKey='" + ctCurrPrtn + "' and ", "") + _
                       "RrCmCtg= 'P' and RrCmCd='" + wCmLkUpCmCd + "' and " + _
                       "RrDmCd= '" + mDmCd + "' and RrDmSfx='' and RrDmSz='')" + _
              ", " + _
              "IsNull((Select RrSalRt" + "  from RmRt where RrTcTyp= 'MRD' and " + IIF(1 = 1, " RrPrtKey='" + ctCurrPrtn + "' and ", "") + _
                       "RrCmCtg= 'P' and RrCmCd='" + wCmLkUpCmCd + "' and RrDmCd='' and RrDmSfx='' and RrDmSz=''),0) " + _
             ")))")
  GetMRD = wMrd
End Function
Public Function GetMRP(ByVal mCmCtg As String, ByVal mCmCd As String, ByVal mCmCurCd, ByVal mDmCd As String, _
                          ByVal mDmSz As String, ByVal mDmSfx As String)
  Dim wMrp As Double
  
  wMrp = goCn.GetFldVal("Select isnull" + _
             "((select " + "RrSalRt" + "  from RmRt where RrTcTyp= 'MRP' and " + IIF(1 = 1, " RrPrtKey='" + ctCurrPrtn + "' and ", "") + _
             " RrCmCtg= 'C' and RrCmCd= 'ZSELF' and RrCmCurCd ='" + mCmCurCd + "' and " + _
             " RrDmCd = '" + mDmCd + "' and RrDmSfx = '" + mDmSfx + "' and RrDmSz='" + mDmSz + "')" + _
              ", " + _
              "IsNull((Select RrSalRt" + "  from RmRt where RrTcTyp= 'MRP' and " + IIF(1 = 1, " RrPrtKey='" + ctCurrPrtn + "' and ", "") + _
                       "RrCmCtg= 'C' and RrCmCd='ZSELF' and RrCmCurCd='" + mCmCurCd + "' and " + _
                           "RrDmCd= '" + mDmCd + "' and RrDmSz='' and RrDmSfx ='" + mDmSfx + "')" + _
              ", " + _
              "IsNull((Select RrSalRt" + "  from RmRt where RrTcTyp= 'MRP' and " + IIF(1 = 1, " RrPrtKey='" + ctCurrPrtn + "' and ", "") + _
                       "RrCmCtg= 'C' and RrCmCd='ZSELF' and RrCmCurCd='" + mCmCurCd + "' and " + _
                       "RrDmCd= '" + mDmCd + "' and RrDmSz='' and RrDmSfx=''),0))" + _
             ")")
  
  GetMRP = wMrp
End Function
'uni.156 if bagno entered with company code then it will ignore the company code and send back the rest of them
Public Function ParseBagNoBarCd(ByVal pBarCd As String) As String
    Dim ws_BagNo() As String, wRfIdNoArr() As String, wRfIdNo As String, i As Integer
    wRfIdNoArr = Split(pBarCd, ",")
    For i = 0 To UBound(wRfIdNoArr)
      wRfIdNo = wRfIdNoArr(i)
      ws_BagNo = Split(wRfIdNo, "/")
      If UBound(ws_BagNo) = 3 Then
        ParseBagNoBarCd = ParseBagNoBarCd + IIF(ParseBagNoBarCd = "", "", ",") + ws_BagNo(1) + "/" + ws_BagNo(2) + "/" + ws_BagNo(3)
      Else
        ParseBagNoBarCd = ParseBagNoBarCd + IIF(ParseBagNoBarCd = "", "", ",") + wRfIdNo
      End If
    Next
End Function

Public Function InBaseCalc(ByVal pGstBase As String, ByVal pCoCd As String, ByVal pInTc As String, ByVal pInYy As String, ByVal pInChr As String, ByVal pInNo As Single) As Double
  If pGstBase = "INFOB" Then
    InBaseCalc = goCn.GetFldVal("select InFOB from InvHd where InCoCd ='" + pCoCd + _
                            "' and InTc ='" + pInTc + "' and InYy ='" + pInYy + "' and InChr ='" + _
                            pInChr + "' and InNo =" + CStr(pInNo) + "")
                            
  ElseIf pGstBase = "OMCS" Then
    InBaseCalc = goCn.GetFldVal("select Sum(OrSalVal*OdOrdQty) from OrdRm join OrdMst on OmCoCd =OrCoCd and OmTc =OrTc " + _
                            "and OmYy =OrYy and OmChr =OrChr and OmNo =OrNo join OrdDsg on OdCoCd =OrCoCd and OdTc =OrTc " + _
                            "and OdYy =OrYy and OdChr =OrChr and OdNo =OrNo and OdSr =OrSr where OmCoCd='" + pCoCd + _
                            "' and OmTc ='" + pInTc + "' and OmYy='" + pInYy + "' and OmChr ='" + _
                            pInChr + "' and OmNo =" + CStr(pInNo) + " and OrRmCtg ='C'")
                            
  ElseIf pGstBase = "OMDIA" Then
    InBaseCalc = goCn.GetFldVal("select Sum(OrSalVal*OdOrdQty) from OrdRm join OrdMst on OmCoCd =OrCoCd and OmTc =OrTc " + _
                            "and OmYy =OrYy and OmChr =OrChr and OmNo =OrNo join OrdDsg on OdCoCd =OrCoCd and OdTc =OrTc " + _
                            "and OdYy =OrYy and OdChr =OrChr and OdNo =OrNo and OdSr =OrSr where OmCoCd='" + pCoCd + _
                            "' and OmTc ='" + pInTc + "' and OmYy='" + pInYy + "' and OmChr ='" + _
                            pInChr + "' and OmNo =" + CStr(pInNo) + " and OrRmCtg ='D'")

  ElseIf pGstBase = "OMDIACS" Then
    InBaseCalc = goCn.GetFldVal("select Sum(OrSalVal*OdOrdQty) from OrdRm join OrdMst on OmCoCd =OrCoCd and OmTc =OrTc " + _
                            "and OmYy =OrYy and OmChr =OrChr and OmNo =OrNo join OrdDsg on OdCoCd =OrCoCd and OdTc =OrTc " + _
                            "and OdYy =OrYy and OdChr =OrChr and OdNo =OrNo and OdSr =OrSr where OmCoCd='" + pCoCd + _
                            "' and OmTc ='" + pInTc + "' and OmYy='" + pInYy + "' and OmChr ='" + _
                            pInChr + "' and OmNo =" + CStr(pInNo) + " and (OrRmCtg ='D' or OrRmCtg ='C')")

  ElseIf pGstBase = "OMJLVAL" Then
    InBaseCalc = goCn.GetFldVal("select Sum(OdJLVchVal) from OrdDsg join OrdMst on OmCoCd =OdCoCd and OmTc =OdTc " + _
                            "and OmYy =OdYy and OmChr =OdChr and OmNo =OdNo where OmCoCd='" + pCoCd + _
                            "' and OmTc ='" + pInTc + "' and OmYy='" + pInYy + "' and OmChr ='" + _
                            pInChr + "' and OmNo =" + CStr(pInNo))

  ElseIf pGstBase = "OMLABVAL" Then
    InBaseCalc = goCn.GetFldVal("select Sum(OlSalVal*OdOrdQty) from OrdLab join OrdMst on OmCoCd =OlCoCd and OmTc =OlTc " + _
                            "and OmYy =OlYy and OmChr =OlChr and OmNo =OlNo join OrdDsg on OdCoCd =OlCoCd and OdTc =OlTc " + _
                            "and OdYy =OlYy and OdChr =OlChr and OdNo =OlNo and OdSr =OlSr  where OmCoCd='" + pCoCd + _
                            "' and OmTc ='" + pInTc + "' and OmYy='" + pInYy + "' and OmChr ='" + _
                            pInChr + "' and OmNo =" + CStr(pInNo))

  ElseIf pGstBase = "OMRMVAL" Then
    InBaseCalc = goCn.GetFldVal("select Sum(OrSalVal*OdOrdQty) from OrdRm join OrdMst on OmCoCd =OrCoCd and OmTc =OrTc " + _
                            "and OmYy =OrYy and OmChr =OrChr and OmNo =OrNo join OrdDsg on OdCoCd =OrCoCd and OdTc =OrTc " + _
                            "and OdYy =OrYy and OdChr =OrChr and OdNo =OrNo and OdSr =OrSr  where OmCoCd='" + pCoCd + _
                            "' and OmTc ='" + pInTc + "' and OmYy='" + pInYy + "' and OmChr ='" + _
                            pInChr + "' and OmNo =" + CStr(pInNo))
  
  ElseIf pGstBase = "OMSALVAL" Then
    InBaseCalc = goCn.GetFldVal("select Sum(OdSalPrc*OdOrdQty) from OrdDsg join OrdMst on OmCoCd =OdCoCd and OmTc =OdTc " + _
                            "and OmYy =OdYy and OmChr =OdChr and OmNo =OdNo where OmCoCd='" + pCoCd + _
                            "' and OmTc ='" + pInTc + "' and OmYy='" + pInYy + "' and OmChr ='" + _
                            pInChr + "' and OmNo =" + CStr(pInNo))
  
  ElseIf pGstBase = "TTOTAMT" Then
    InBaseCalc = goCn.GetFldVal("select TTotAmt from Txn where TCoCd='" + pCoCd + _
                            "' and TTc ='" + pInTc + "' and TYy='" + pInYy + "' and TChr ='" + _
                            pInChr + "' and TNo =" + CStr(pInNo))
  End If
    
End Function
  
Public Function GetGldRtSri(ByVal mCmCd As String, ByVal mLme As Single, _
                         ByVal mRmCd As String, ByVal mAlyRt As Single, _
                         ByVal mDmTcTyp As String, ByVal mDmCd As String, _
                         ByVal mOpt As String, ByVal ms_CmCtg As String, _
                         Optional ByVal mf_CTBCurCnv As Variant, _
                         Optional ByVal mf_CmCurCd As Variant, _
                         Optional ByVal mCmLkUpRmRt As String, _
                         Optional ByVal mCmLkUpMetLs As String, _
                         Optional ByVal mPrtKey As String, Optional ByRef mGldLsPer As Variant) As Single
                         
  'Optional ByVal mf_CurCd As Variant = "RS"
  
  Dim wDmCtg As String, wLsCtg As String, wRmCtg As String, wRmSCtg As String
  Dim wRmPurity As String, wRt As String, wAlyRt As Double, wInc As Double
  Dim wGldRt As Double, wGldLs As Double, wRmKT As String
  Dim wCustRtPurity As Double '*** Jay 3.1.3 [CRP]
  Dim wCmSeekZSelfRt As Boolean
  
  '****** Sachin 2.13.0 - 29-05-2006 ******
  Dim wQwCS As String, wQw As String
  '****** Sachin 2.13.0 - 29-05-2006 ******
  
  ' ***** Manali 3.03 - 02/07/08 - variables shifted to func declaration
  'Dim wCmLkUpRmRt As String, wCmLkUpMetLs As String
  ' ***** Manali 3.03 - 02/07/08 - variables shifted to func declaration
  
  Dim wGldLsStr As String
  
  If (mOpt <> "C") And (mOpt <> "S") Then GetGldRtSri = 0: Exit Function
  
  ' Zubin 213
  If IsMissing(mf_CmCurCd) Then mf_CmCurCd = goCn.GetFldVal("Select CmCurCd from CustMst where CmCtg= '" + ms_CmCtg + "' and CmCd= '" + mCmCd + "'")

  wDmCtg = goCn.GetFldVal("Select DmCtg from DsgMst where DmTcTyp= '" + mDmTcTyp + "' and DmSz= '' " + _
                         "and DmCd= '" + mDmCd + "'")
  wRmCtg = goCn.GetFldVal("Select RmCtg from RmMst where RmCd='" + mRmCd + "'" + IIF(mPrtKey = ctCurrPrtn, " and RmPrtKey='" + ctCurrPrtn + "'", ""))
  wRmSCtg = goCn.GetFldVal("Select RmSCtg from RmMst where RmCd='" + mRmCd + "'" + IIF(mPrtKey = ctCurrPrtn, " and RmPrtKey='" + ctCurrPrtn + "'", ""))
  wRmKT = goCn.GetFldVal("Select RmKT from RmMst where RmCd='" + mRmCd + "'" + IIF(mPrtKey = ctCurrPrtn, " and RmPrtKey='" + ctCurrPrtn + "'", ""))
  wRmPurity = IIF(mOpt = "S", "RmPurityRt", "RmPurityWt")
  wRt = IIF(mOpt = "S", "RrSalRt", "RrCstRt")
  
  '****** Sachin 2.13.0 - 29-05-2006 ******
  wQwCS = IIF(mOpt = "S", "RmQw", "RmQwCst")
  wQw = goCn.GetFldVal("Select " + wQwCS + " From RmMst Where RmCd='" + mRmCd + "'" + IIF(mPrtKey = ctCurrPrtn, " and RmPrtKey='" + ctCurrPrtn + "'", ""))
  '****** Sachin 2.13.0 - 29-05-2006 ******
  
  '*** Jay 3.1.3 [CRP]
  If mOpt = "S" Then
    wCustRtPurity = goCn.GetFldVal("Select RrSalRt from RmRt where RrTcTyp = 'CRP' and RrCd= '" + mRmCd + "' " + _
                                        "and RrCmCtg= '" + ms_CmCtg + "' and RrCmCd= '" + mCmCd + "'" + IIF(mPrtKey = ctCurrPrtn, " and RrPrtKey='" + ctCurrPrtn + "'", ""))
  End If
  '*** Jay 3.1.3 [CRP]

  '******** AlyRt *********
  '**** for Sale  = (1-RmPurityRt)* mAlyRt
  '**** for Cost  = (1-RmPurityWt)* mAlyRt
  '*** Jay 3.1.3 [CRP]
  wAlyRt = (1 - IIF(wCustRtPurity <> 0 And mOpt = "S", wCustRtPurity, goCn.GetFldVal("Select " + wRmPurity + " from RmMst where RmCd='" + mRmCd + "'" + IIF(mPrtKey = ctCurrPrtn, " and RmPrtKey='" + ctCurrPrtn + "'", "")))) _
           * mAlyRt
  '*** Jay 3.1.3 [CRP]
  '********** Increments, only if mOpt='S'
'Manoj
'  If mOpt = "S" And Trim(ms_CmCtg) = "C" Then
  If mOpt = "S" And ((Trim(ms_CmCtg) = "C") Or (Trim(ms_CmCtg) = "S")) Then
'    wInc = gocn.GetFldVal("Select RrSalRt from RmRt where RrTcTyp= 'INC' " + _
'           "and RrCtg= '" + wRmCtg + "' and RrCmCtg= 'C' and RrCmCd='" + mCmCd + "'")
    wInc = goCn.GetFldVal("Select RrSalRt from RmRt where RrTcTyp= 'INC' " + _
           "and RrCtg= '" + wRmCtg + "' and RrCmCtg= '" & Trim$(ms_CmCtg) & "' and RrCmCd='" + mCmCd + "'")
'Manoj
  Else
    wInc = 0
  End If
  
  ' ****** Bef 3.03 - Manali
  'wCmLkUpRmRt = goCn.GetFldVal("Select CmLkUpRmRt from CustMst where CmCtg= '" + ms_CmCtg + "' and CmCd= '" + mCmCd + "'")
  
  ' ***** Manali - 3.03 - 02/07/08
  If mCmLkUpRmRt = "" Then _
  mCmLkUpRmRt = goCn.GetFldVal("Select CmLkUpRmRt from CustMst where CmCtg= '" + ms_CmCtg + "' and CmCd= '" + mCmCd + "'")
  
  ' ***** Manali 3.8.0 - Cost Prc for CmCtg='S' from CmCtg='C'
  If ms_CmCtg = "S" And mOpt = "C" Then ms_CmCtg = "C"

  '********* looking for per gm record ('RM') for price List customer ***********
  '***** Look for record with the specified DmCtg. If not found then look for DmCtg= ''
'Manoj
'  If Trim(ms_CmCtg) = "C" Then
  If (Trim(ms_CmCtg) = "C") Or (Trim(ms_CmCtg) = "S") Then
'Manoj
    ' Zubin 213 (RrCmCurCd Cnd added)
    wGldRt = goCn.GetFldVal("Select isnull" + _
             "((select " + wRt + "  from RmRt where RrTcTyp= 'RM' and " + IIF(mPrtKey = ctCurrPrtn, " RrPrtKey='" + ctCurrPrtn + "' and ", "") + _
               "RrCtg= '" + wRmCtg + "' and RrSCtg= '" + wRmSCtg + "' and " + _
               "RrCd= '" + mRmCd + "' and RrFrLn<= " + CStr(mLme) + " And " + _
               "RrToLn>= " + CStr(mLme) + " and RrCmCtg= 'P' and RrCmCd= '" + mCmLkUpRmRt + "' and " + _
               "RrCmCurCd= '" + mf_CmCurCd + "' and " + _
               "RrDmCtg= '" + wDmCtg + "' and " + wRt + " <> 0)" + _
              ", " + _
              "IsNull((Select " + wRt + "  from RmRt where RrTcTyp= 'RM' and " + IIF(mPrtKey = ctCurrPrtn, " RrPrtKey='" + ctCurrPrtn + "' and ", "") + _
                       "RrCtg= '" + wRmCtg + "' and RrSCtg= '" + wRmSCtg + "' and " + _
                       "RrCd= '" + mRmCd + "' and RrFrLn <= " + CStr(mLme) + " And " + _
                       "RrToLn>= " + CStr(mLme) + " and RrCmCtg= 'P' and RrCmCd= '" + mCmLkUpRmRt + "' and " + _
                       "RrCmCurCd= '" + mf_CmCurCd + "' and " + _
                       "RrDmCtg= '' and " + wRt + " <> 0),0)" + _
             ")")
  End If
  '********* per gm record is found from Customer record,
  If wGldRt <> 0 Then GetGldRtSri = (wGldRt + IIF(wQw = "W", wAlyRt, 0)): Exit Function
  
  '********* looking for per gm record ('RM') for defined customer ***********
  '***** Look for record with the specified DmCtg. If not found then look for DmCtg= ''
  ' Zubin 213 (RrCmCurCd Cnd added)
  wGldRt = goCn.GetFldVal("Select isnull" + _
           "((select " + wRt + "  from RmRt where RrTcTyp= 'RM' and " + IIF(mPrtKey = ctCurrPrtn, " RrPrtKey='" + ctCurrPrtn + "' and ", "") + _
              "RrCtg= '" + wRmCtg + "' and RrSCtg= '" + wRmSCtg + "' and " + _
              "RrCd= '" + mRmCd + "' and RrFrLn <= " + CStr(mLme) + " And " + _
              "RrToLn>= " + CStr(mLme) + " and RrCmCtg= '" + ms_CmCtg + "' and RrCmCd = '" + mCmCd + "' and " + _
              "RrCmCurCd= '" + mf_CmCurCd + "' and " + _
              "RrDmCtg= '" + wDmCtg + "' and " + wRt + " <> 0)" + _
            ", " + _
            "IsNull((Select " + wRt + "  from RmRt where RrTcTyp='RM' and " + IIF(mPrtKey = ctCurrPrtn, " RrPrtKey='" + ctCurrPrtn + "' and ", "") + _
                     "RrCtg= '" + wRmCtg + "' and RrSCtg= '" + wRmSCtg + "' and " + _
                     "RrCd='" + mRmCd + "' and RrFrLn <= " + CStr(mLme) + " And " + _
                     "RrToLn >= " + CStr(mLme) + " and RrCmCtg= '" + ms_CmCtg + "' and RrCmCd = '" + mCmCd + "' and " + _
                     "RrCmCurCd= '" + mf_CmCurCd + "' and " + _
                     "RrDmCtg='' and " + wRt + " <> 0),0)" + _
           ")")

  '********* per gm record is found from Customer record,
  If wGldRt <> 0 Then GetGldRtSri = (wGldRt + IIF(wQw = "W", wAlyRt, 0)): Exit Function

  '********* looking for per gm record for Self ***********
  '***** Look for record with the specified DmCtg. If not found then look for DmCtg= ''
'Manoj
'  If Trim(ms_CmCtg) = "C" Then
  If (Trim(ms_CmCtg) = "C") Or (Trim(ms_CmCtg) = "S") And (mOpt = "S" And wCmSeekZSelfRt) Then
'    wGldRt = gocn.GetFldVal("Select isnull" + _
'             "((select " + wRt + " from RmRt where RrTcTyp='RM' and " + _
'                "RrCtg= '" + wRmCtg + "' and RrSCtg= '" + wRmSCtg + "' and " + _
'                "RrCd='" + mRmCd + "' and RrFrLn <= " + CStr(mLme) + " And " + _
'                "RrToLn >= " + CStr(mLme) + " and RrCmCtg= 'C' and RrCmCd = '" + ctSelfCmCd + "' and " + _
'                "RrDmCtg='" + wDmCtg + "' and " + wRt + " <> 0)" + _
'               ", " + _
'               "IsNull((select " + wRt + " from RmRt where RrTcTyp='RM' and " + _
'                        "RrCtg= '" + wRmCtg + "' and RrSCtg= '" + wRmSCtg + "' and " + _
'                        "RrCd='" + mRmCd + "' and RrFrLn <= " + CStr(mLme) + " And " + _
'                        "RrToLn >= " + CStr(mLme) + " and RrCmCtg= 'C' and RrCmCd ='" + ctSelfCmCd + "' and " + _
'                        "RrDmCtg='' and " + wRt + "<> 0),0)" + _
'             ")")
    ' Zubin 213 (RrCmCurCd Cnd added)
    wGldRt = goCn.GetFldVal("Select isnull" + _
             "((select " + wRt + " from RmRt where RrTcTyp='RM' and " + IIF(mPrtKey = ctCurrPrtn, " RrPrtKey='" + ctCurrPrtn + "' and ", "") + _
                "RrCtg= '" + wRmCtg + "' and RrSCtg= '" + wRmSCtg + "' and " + _
                "RrCd='" + mRmCd + "' and RrFrLn <= " + CStr(mLme) + " And " + _
                "RrToLn >= " + CStr(mLme) + " and RrCmCtg= '" & ms_CmCtg & "' and RrCmCd = '" + ctSelfCmCd + "' and " + _
                "RrCmCurCd= '" + mf_CmCurCd + "' and " + _
                "RrDmCtg='" + wDmCtg + "' and " + wRt + " <> 0)" + _
               ", " + _
               "IsNull((select " + wRt + " from RmRt where RrTcTyp='RM' and " + IIF(mPrtKey = ctCurrPrtn, " RrPrtKey='" + ctCurrPrtn + "' and ", "") + _
                        "RrCtg= '" + wRmCtg + "' and RrSCtg= '" + wRmSCtg + "' and " + _
                        "RrCd='" + mRmCd + "' and RrFrLn <= " + CStr(mLme) + " And " + _
                        "RrToLn >= " + CStr(mLme) + " and RrCmCtg= '" & ms_CmCtg & "' and RrCmCd ='" + ctSelfCmCd + "' and " + _
                        "RrCmCurCd= '" + mf_CmCurCd + "' and " + _
                        "RrDmCtg='' and " + wRt + "<> 0),0)" + _
             ")")
'Manoj
  End If
  '********* per gm record is found from Self record,
  If wGldRt <> 0 Then GetGldRtSri = (wGldRt * (1 + wInc / 100) + IIF(wQw = "W", wAlyRt, 0)): Exit Function
  If wQw = "W" Then
    wLsCtg = goCn.GetFldVal("Select DmLsCtg from DsgMst where DmTcTyp= '" + mDmTcTyp + "' and " + _
                            "DmSz= '' and DmCd= '" + mDmCd + "'")
                            
    If mCmLkUpMetLs = "" Then _
      mCmLkUpMetLs = goCn.GetFldVal("Select CmLkUpMetLs from CustMst where CmCtg= '" + ms_CmCtg + "' and CmCd= '" + mCmCd + "'")
    
    '**********Bhavna added DsgCd, rmkt for loss%
    wGldLsStr = "Select IsNull" + _
            "((select " + wRt + " from RmRt where RrTcTyp='LS' and RrCtg='" + wRmCtg + "' and " + IIF(mPrtKey = ctCurrPrtn, " RrPrtKey='" + ctCurrPrtn + "' and ", "") + _
               "RrCmCtg= 'P' and RrCmCd = '" + mCmLkUpMetLs + "' and RrDmCd= '" + mDmCd + "'" + _
               " and RrRmKT = '" + wRmKT + "' and RrCmCurCd='" + "" + "' and " + wRt + "<> 0), " + _
              "IsNull ((select " + wRt + " from RmRt where RrTcTyp='LS' and RrCtg='" + wRmCtg + "' and " + IIF(mPrtKey = ctCurrPrtn, " RrPrtKey='" + ctCurrPrtn + "' and ", "") + _
                 "RrCmCtg= 'P' and RrCmCd = '" + mCmLkUpMetLs + "' and RrDmCd= '" + mDmCd + "'" + _
                 " and RrRmKT = '' and RrCmCurCd='" + "" + "' and " + wRt + "<> 0), "
                 
    '**********Bhavna added DsgCd, rmkt for loss%
    
    wGldLsStr = wGldLsStr + "IsNull" + _
             "((select " + wRt + " from RmRt where RrTcTyp='LS' and RrCtg='" + wRmCtg + "' and " + IIF(mPrtKey = ctCurrPrtn, " RrPrtKey='" + ctCurrPrtn + "' and ", "") + _
             "RrCmCtg= 'P' and RrCmCd = '" + mCmLkUpMetLs + "' and RrDmCtg= '" + wDmCtg + "' and RrDmCd = '' and " + _
             "RrLsCtg= '" + wLsCtg + "' and RrRmKT = '" + wRmKT + "' and RrCmCurCd='" + "" + "' and " + wRt + "<> 0), " + _
               "IsNull((select " + wRt + " from RmRt where RrTcTyp='LS' and RrCtg='" + wRmCtg + "' and " + IIF(mPrtKey = ctCurrPrtn, " RrPrtKey='" + ctCurrPrtn + "' and ", "") + _
               "RrCmCtg= 'P' and RrCmCd = '" + mCmLkUpMetLs + "' and RrDmCtg= '" + wDmCtg + "' and RrDmCd = '' and " + _
               "RrLsCtg= '" + wLsCtg + "' and RrRmKT = '' and RrCmCurCd='" + "" + "' and " + wRt + "<> 0), " + _
                "IsNull((select " + wRt + " from RmRt where RrTcTyp= 'LS' and RrCtg='" + wRmCtg + "' and " + IIF(mPrtKey = ctCurrPrtn, " RrPrtKey='" + ctCurrPrtn + "' and ", "") + _
                      "RrCmCtg= 'P' and RrCmCd = '" + mCmLkUpMetLs + "' and RrDmCtg= '" + wDmCtg + "' and RrDmCd = '' and " + _
                      "RrLsCtg= '' and RrRmKT = '" + wRmKT + "' and RrCmCurCd='" + "" + "' and " + wRt + "<> 0), " + _
                 "IsNull((select " + wRt + " from RmRt where RrTcTyp= 'LS' and RrCtg='" + wRmCtg + "' and " + IIF(mPrtKey = ctCurrPrtn, " RrPrtKey='" + ctCurrPrtn + "' and ", "") + _
                      "RrCmCtg= 'P' and RrCmCd = '" + mCmLkUpMetLs + "' and RrDmCtg= '" + wDmCtg + "' and RrDmCd = '' and " + _
                      "RrLsCtg= '' and RrRmKT = '' and RrCmCurCd='" + "" + "' and " + wRt + "<> 0), " + _
                     "IsNull((select " + wRt + " from RmRt where RrTcTyp= 'LS' and RrCtg='" + wRmCtg + "' and " + IIF(mPrtKey = ctCurrPrtn, " RrPrtKey='" + ctCurrPrtn + "' and ", "") + _
                             "RrCmCtg= 'P' and RrCmCd = '" + mCmLkUpMetLs + "' and RrDmCtg='' and RrDmCd = '' and " + _
                             "RrLsCtg= '' and RrRmKT = '' and RrCmCurCd='" + "" + "' and " + wRt + "<> 0),0)))))))"
                             
    wGldLs = goCn.GetFldVal(wGldLsStr)
    
    If wGldLs = 0 Then
        wGldLsStr = "Select IsNull" + _
            "((select " + wRt + " from RmRt where RrTcTyp='LS' and RrCtg='" + wRmCtg + "' and " + IIF(mPrtKey = ctCurrPrtn, " RrPrtKey='" + ctCurrPrtn + "' and ", "") + _
              "RrCmCtg= '" + ms_CmCtg + "' and RrCmCd = '" + mCmCd + "' and RrDmCd= '" + mDmCd + "'" + _
              " and RrRmKT = '" + wRmKT + "' and RrCmCurCd='" + "" + "' and " + wRt + "<> 0), " + _
              "IsNull ((select " + wRt + " from RmRt where RrTcTyp='LS' and RrCtg='" + wRmCtg + "' and " + IIF(mPrtKey = ctCurrPrtn, " RrPrtKey='" + ctCurrPrtn + "' and ", "") + _
                 "RrCmCtg= '" + ms_CmCtg + "' and RrCmCd = '" + mCmCd + "' and RrDmCd= '" + mDmCd + "'" + _
                 " and RrRmKT = '' and RrCmCurCd='" + "" + "' and " + wRt + "<> 0), "

        wGldLsStr = wGldLsStr + "IsNull" + _
                "((select " + wRt + " from RmRt where RrTcTyp= 'LS' and RrCtg='" + wRmCtg + "' and " + IIF(mPrtKey = ctCurrPrtn, " RrPrtKey='" + ctCurrPrtn + "' and ", "") + _
                "RrCmCtg= '" + ms_CmCtg + "' and RrCmCd = '" + mCmCd + "' and RrDmCtg= '" + wDmCtg + "' and RrDmCd = '' and " + _
                "RrLsCtg= '" + wLsCtg + "' and RrRmKT = '" + wRmKT + "' and RrCmCurCd='" + "" + "' and " + wRt + "<> 0), " + _
                    "IsNull ((select " + wRt + " from RmRt where RrTcTyp= 'LS' and RrCtg='" + wRmCtg + "' and " + IIF(mPrtKey = ctCurrPrtn, " RrPrtKey='" + ctCurrPrtn + "' and ", "") + _
                    "RrCmCtg= '" + ms_CmCtg + "' and RrCmCd = '" + mCmCd + "' and RrDmCtg= '" + wDmCtg + "' and RrDmCd = '' and " + _
                    "RrLsCtg= '" + wLsCtg + "' and RrRmKT = '' and RrCmCurCd='" + "" + "' and " + wRt + "<> 0), " + _
                      "IsNull ((select " + wRt + " from RmRt where RrTcTyp= 'LS' and RrCtg='" + wRmCtg + "' and " + IIF(mPrtKey = ctCurrPrtn, " RrPrtKey='" + ctCurrPrtn + "' and ", "") + _
                      "RrCmCtg= '" + ms_CmCtg + "' and RrCmCd = '" + mCmCd + "' and RrDmCtg= '" + wDmCtg + "' and RrDmCd = '' and " + _
                      "RrLsCtg= '' and RrRmKT = '" + wRmKT + "' and RrCmCurCd='" + "" + "' and " + wRt + "<> 0), " + _
                        "IsNull ((select " + wRt + " from RmRt where RrTcTyp= 'LS' and RrCtg='" + wRmCtg + "' and " + IIF(mPrtKey = ctCurrPrtn, " RrPrtKey='" + ctCurrPrtn + "' and ", "") + _
                        "RrCmCtg= '" + ms_CmCtg + "' and RrCmCd = '" + mCmCd + "' and RrDmCtg= '" + wDmCtg + "' and RrDmCd = '' and " + _
                        "RrLsCtg= '' and RrRmKT = '' and RrCmCurCd='" + "" + "' and " + wRt + "<> 0), " + _
                            "IsNull ((select " + wRt + " from RmRt where RrTcTyp= 'LS' and RrCtg='" + wRmCtg + "' and " + IIF(mPrtKey = ctCurrPrtn, " RrPrtKey='" + ctCurrPrtn + "' and ", "") + _
                            "RrCmCtg= '" + ms_CmCtg + "' and RrCmCd = '" + mCmCd + "' and RrDmCtg= '' and RrDmCd = '' and " + _
                            "RrLsCtg= '' and RrRmKT = '' and RrCmCurCd='" + "" + "' and " + wRt + "<> 0),0)))))))"
        
        
        wGldLs = goCn.GetFldVal(wGldLsStr)
    End If

    If wGldLs = 0 Then
        wGldLsStr = "Select IsNull" + _
            "((select " + wRt + " from RmRt where RrTcTyp='LS' and RrCtg='" + wRmCtg + "' and " + IIF(mPrtKey = ctCurrPrtn, " RrPrtKey='" + ctCurrPrtn + "' and ", "") + _
              "RrCmCtg= '" + ms_CmCtg + "' and RrCmCd = '" + ctSelfCmCd + "' and RrDmCd= '" + mDmCd + "'" + _
              " and RrRmKT = '" + wRmKT + "' and RrCmCurCd='" + mf_CmCurCd + "' and " + wRt + "<> 0), " + _
              "IsNull ((select " + wRt + " from RmRt where RrTcTyp='LS' and RrCtg='" + wRmCtg + "' and " + IIF(mPrtKey = ctCurrPrtn, " RrPrtKey='" + ctCurrPrtn + "' and ", "") + _
                 "RrCmCtg= '" + ms_CmCtg + "' and RrCmCd = '" + ctSelfCmCd + "' and RrDmCd= '" + mDmCd + "'" + _
                 " and RrRmKT = '' and RrCmCurCd='" + mf_CmCurCd + "' and " + wRt + "<> 0), "

        wGldLsStr = wGldLsStr + "IsNull" + _
                "((select " + wRt + " from RmRt where RrTcTyp= 'LS' and RrCtg='" + wRmCtg + "' and " + IIF(mPrtKey = ctCurrPrtn, " RrPrtKey='" + ctCurrPrtn + "' and ", "") + _
                "RrCmCtg= '" + ms_CmCtg + "' and RrCmCd = '" + ctSelfCmCd + "' and RrDmCtg= '" + wDmCtg + "' and RrDmCd = '' and " + _
                "RrLsCtg= '" + wLsCtg + "' and RrRmKT = '" + wRmKT + "' and RrCmCurCd='" + mf_CmCurCd + "' and " + wRt + "<> 0), " + _
                    "IsNull ((select " + wRt + " from RmRt where RrTcTyp= 'LS' and RrCtg='" + wRmCtg + "' and " + IIF(mPrtKey = ctCurrPrtn, " RrPrtKey='" + ctCurrPrtn + "' and ", "") + _
                    "RrCmCtg= '" + ms_CmCtg + "' and RrCmCd = '" + ctSelfCmCd + "' and RrDmCtg= '" + wDmCtg + "' and RrDmCd = '' and " + _
                    "RrLsCtg= '" + wLsCtg + "' and RrRmKT = '' and RrCmCurCd='" + mf_CmCurCd + "' and " + wRt + "<> 0), " + _
                      "IsNull ((select " + wRt + " from RmRt where RrTcTyp= 'LS' and RrCtg='" + wRmCtg + "' and " + IIF(mPrtKey = ctCurrPrtn, " RrPrtKey='" + ctCurrPrtn + "' and ", "") + _
                      "RrCmCtg= '" + ms_CmCtg + "' and RrCmCd = '" + ctSelfCmCd + "' and RrDmCtg= '" + wDmCtg + "' and RrDmCd = '' and " + _
                      "RrLsCtg= '' and RrRmKT = '" + wRmKT + "' and RrCmCurCd='" + mf_CmCurCd + "' and " + wRt + "<> 0), " + _
                        "IsNull ((select " + wRt + " from RmRt where RrTcTyp= 'LS' and RrCtg='" + wRmCtg + "' and " + IIF(mPrtKey = ctCurrPrtn, " RrPrtKey='" + ctCurrPrtn + "' and ", "") + _
                        "RrCmCtg= '" + ms_CmCtg + "' and RrCmCd = '" + ctSelfCmCd + "' and RrDmCtg= '" + wDmCtg + "' and RrDmCd = '' and " + _
                        "RrLsCtg= '' and RrRmKT = '' and RrCmCurCd='" + mf_CmCurCd + "' and " + wRt + "<> 0), " + _
                            "IsNull ((select " + wRt + " from RmRt where RrTcTyp= 'LS' and RrCtg='" + wRmCtg + "' and " + IIF(mPrtKey = ctCurrPrtn, " RrPrtKey='" + ctCurrPrtn + "' and ", "") + _
                            "RrCmCtg= '" + ms_CmCtg + "' and RrCmCd = '" + ctSelfCmCd + "' and RrDmCtg= '' and RrDmCd = '' and " + _
                            "RrLsCtg= '' and RrRmKT = '' and RrCmCurCd='" + mf_CmCurCd + "' and " + wRt + "<> 0),0)))))))"
        wGldLs = goCn.GetFldVal(wGldLsStr)
    End If
    If wGldLs = 0 Then
        wGldLsStr = "Select IsNull" + _
            "((select " + wRt + " from RmRt where RrTcTyp='LS' and RrCtg='" + wRmCtg + "' and " + IIF(mPrtKey = ctCurrPrtn, " RrPrtKey='" + ctCurrPrtn + "' and ", "") + _
              "RrCmCtg= '" + ms_CmCtg + "' and RrCmCd = '" + ctSelfCmCd + "' and RrDmCd= '" + mDmCd + "'" + _
              " and RrRmKT = '" + wRmKT + "' and RrCmCurCd='" + "" + "' and " + wRt + "<> 0), " + _
              "IsNull ((select " + wRt + " from RmRt where RrTcTyp='LS' and RrCtg='" + wRmCtg + "' and " + IIF(mPrtKey = ctCurrPrtn, " RrPrtKey='" + ctCurrPrtn + "' and ", "") + _
                 "RrCmCtg= '" + ms_CmCtg + "' and RrCmCd = '" + ctSelfCmCd + "' and RrDmCd= '" + mDmCd + "'" + _
                 " and RrRmKT = '' and RrCmCurCd='" + "" + "' and " + wRt + "<> 0), "

        wGldLsStr = wGldLsStr + "IsNull" + _
                "((select " + wRt + " from RmRt where RrTcTyp= 'LS' and RrCtg='" + wRmCtg + "' and " + IIF(mPrtKey = ctCurrPrtn, " RrPrtKey='" + ctCurrPrtn + "' and ", "") + _
                "RrCmCtg= '" + ms_CmCtg + "' and RrCmCd = '" + ctSelfCmCd + "' and RrDmCtg= '" + wDmCtg + "' and RrDmCd = '' and " + _
                "RrLsCtg= '" + wLsCtg + "' and RrRmKT = '" + wRmKT + "' and RrCmCurCd='" + "" + "' and " + wRt + "<> 0), " + _
                    "IsNull ((select " + wRt + " from RmRt where RrTcTyp= 'LS' and RrCtg='" + wRmCtg + "' and " + IIF(mPrtKey = ctCurrPrtn, " RrPrtKey='" + ctCurrPrtn + "' and ", "") + _
                    "RrCmCtg= '" + ms_CmCtg + "' and RrCmCd = '" + ctSelfCmCd + "' and RrDmCtg= '" + wDmCtg + "' and RrDmCd = '' and " + _
                    "RrLsCtg= '" + wLsCtg + "' and RrRmKT = '' and RrCmCurCd='" + "" + "' and " + wRt + "<> 0), " + _
                      "IsNull ((select " + wRt + " from RmRt where RrTcTyp= 'LS' and RrCtg='" + wRmCtg + "' and " + IIF(mPrtKey = ctCurrPrtn, " RrPrtKey='" + ctCurrPrtn + "' and ", "") + _
                      "RrCmCtg= '" + ms_CmCtg + "' and RrCmCd = '" + ctSelfCmCd + "' and RrDmCtg= '" + wDmCtg + "' and RrDmCd = '' and " + _
                      "RrLsCtg= '' and RrRmKT = '" + wRmKT + "' and RrCmCurCd='" + "" + "' and " + wRt + "<> 0), " + _
                        "IsNull ((select " + wRt + " from RmRt where RrTcTyp= 'LS' and RrCtg='" + wRmCtg + "' and " + IIF(mPrtKey = ctCurrPrtn, " RrPrtKey='" + ctCurrPrtn + "' and ", "") + _
                        "RrCmCtg= '" + ms_CmCtg + "' and RrCmCd = '" + ctSelfCmCd + "' and RrDmCtg= '" + wDmCtg + "' and RrDmCd = '' and " + _
                        "RrLsCtg= '' and RrRmKT = '' and RrCmCurCd='" + "" + "' and " + wRt + "<> 0), " + _
                            "IsNull ((select " + wRt + " from RmRt where RrTcTyp= 'LS' and RrCtg='" + wRmCtg + "' and " + IIF(mPrtKey = ctCurrPrtn, " RrPrtKey='" + ctCurrPrtn + "' and ", "") + _
                            "RrCmCtg= '" + ms_CmCtg + "' and RrCmCd = '" + ctSelfCmCd + "' and RrDmCtg= '' and RrDmCd = '' and " + _
                            "RrLsCtg= '' and RrRmKT = '' and RrCmCurCd='" + "" + "' and " + wRt + "<> 0),0)))))))"
        wGldLs = goCn.GetFldVal(wGldLsStr)
    End If
    'indigo.162- passed mGldLsPer parameter to store Ls % value
    mGldLsPer = wGldLs
    
    wRmPurity = IIF(mOpt = "S", "RmPurityRt", "RmPurityWt")
    If IsMissing(mf_CTBCurCnv) Then
      ' Zubin 213 (Before 213)
      'wGldRt = MWLib.Div(mLme, gocn.GetFldVal("Select PNum from param where PTyp= 'OUNCE' " + _
                               "and PMCd='" + wRmCtg + "'")) _
               * gocn.GetFldVal("Select " + wRmPurity + " from RmMst where " + _
                 "RmCd='" + mRmCd + "'") _
               * (1 + wGldLs / 100) + wAlyRt
      ' Zubin 213
      '*** Jay 3.1.3 [CRP]
      wGldRt = MWLib.Div(mLme, goCn.GetFldVal("Select PNum from param where PTyp= 'GMFACTOR' " + _
                               "and PMCd='" + wRmCtg + "' and PSCd= '" + mf_CmCurCd + "' ")) _
               * IIF(wCustRtPurity <> 0 And mOpt = "S", wCustRtPurity, goCn.GetFldVal("Select " + wRmPurity + " from RmMst where " + _
                 "RmCd='" + mRmCd + "'")) _
               * (1 + wGldLs / 100) + wAlyRt
      '*** Jay 3.1.3 [CRP]
    Else
      ' Zubin 213 (Before 213)
      'wGldRt = MWLib.Div(mLme, gocn.GetFldVal("Select PNum from param where PTyp= 'OUNCE' " + _
                               "and PMCd= '" + wRmCtg + "'") * gocn.GetFldVal("Select PNum from " + _
                           "Param where PTyp= 'CURNCY' and PMCd= '" + ctDefCurncy + "'")) _
               * gocn.GetFldVal("Select " + wRmPurity + " from RmMst where " + _
                 "RmCd= '" + mRmCd + "'") _
               * (1 + wGldLs / 100) * mf_CTBCurCnv + wAlyRt
      ' Zubin 213
      '*** Jay 3.1.3 [CRP]
      wGldRt = MWLib.Div(mLme, goCn.GetFldVal("Select PNum from param where PTyp= 'GMFACTOR' " + _
                               "and PMCd= '" + wRmCtg + "'and PSCd= '" + mf_CmCurCd + "' ") _
                          * goCn.GetFldVal("Select PNum from Param where PTyp= 'CURNCY' and " + _
                              "PMCd= '" + mf_CmCurCd + "'")) _
               * IIF(wCustRtPurity <> 0 And mOpt = "S", wCustRtPurity, goCn.GetFldVal("Select " + wRmPurity + " from RmMst where " + _
                 "RmCd= '" + mRmCd + "'")) _
               * (1 + wGldLs / 100) * mf_CTBCurCnv + wAlyRt
      '*** Jay 3.1.3 [CRP]
       
    End If
    GetGldRtSri = wGldRt
  Else
    GetGldRtSri = 0
  End If
End Function
'6.1-85 mLsRtApplicableYN parameter added
Public Function GetGldRt(ByVal mCmCd As String, ByVal mLme As Single, _
                         ByVal mRmCd As String, ByVal mAlyRt As Single, _
                         ByVal mDmTcTyp As String, ByVal mDmCd As String, _
                         ByVal mOpt As String, ByVal ms_CmCtg As String, _
                         Optional ByVal mf_CTBCurCnv As Variant, _
                         Optional ByVal mf_CmCurCd As Variant, _
                         Optional ByVal mCmLkUpRmRt As String, _
                         Optional ByVal mCmLkUpMetLs As String, _
                         Optional ByVal mPrtKey As String, Optional ByVal mLsRtApplicableYN As String = "Y", _
                         Optional ByVal mLn1 As Single, Optional ByVal mLn2 As Single, _
                         Optional ByVal mWtPerDia As Single) As Single
  
    Dim wRtBy As String
    'wRtBy = goCn.GetFldVal("Select pvALUE from RmMst join Param on RmCtg = PmCd and PTyp='RMSCTG' and PSCd=RmSCtg where RmCd='" + mRmCd + "'")
    GetGldRt = GetRmRt(mCmCd, mRmCd, mLn1, mLn2, mWtPerDia, mOpt, mLme, ms_CmCtg, "", "", 0, mf_CmCurCd, mCmLkUpRmRt, mPrtKey)
    If GetGldRt <> 0 Then
      Exit Function
    End If
    
  'Optional ByVal mf_CurCd As Variant = "RS"
  Dim wDmCtg As String, wLsCtg As String, wRmCtg As String, wRmSCtg As String
  Dim wRmPurity As String, wRt As String, wAlyRt As Double, wInc As Double
  Dim wGldRt As Double, wGldLs As Double, wRmKT As String
  Dim wCustRtPurity As Double '*** Jay 3.1.3 [CRP]
  Dim wCmSeekZSelfRt As Boolean
  '****** Sachin 2.13.0 - 29-05-2006 ******
  Dim wQwCS As String, wQw As String
  '****** Sachin 2.13.0 - 29-05-2006 ******
  ' ***** Manali 3.03 - 02/07/08 - variables shifted to func declaration
  'Dim wCmLkUpRmRt As String, wCmLkUpMetLs As String
  ' ***** Manali 3.03 - 02/07/08 - variables shifted to func declaration
  Dim wGldLsStr As String
  If (mOpt <> "C") And (mOpt <> "S") Then GetGldRt = 0: Exit Function
  ' Zubin 213
  If IsMissing(mf_CmCurCd) Then mf_CmCurCd = goCn.GetFldVal("Select CmCurCd from CustMst where CmCtg= '" + ms_CmCtg + "' and CmCd= '" + mCmCd + "'")
  wDmCtg = goCn.GetFldVal("Select DmCtg from DsgMst where DmTcTyp= '" + mDmTcTyp + "' and DmSz= '' " + _
                         "and DmCd= '" + mDmCd + "'")
  wRmCtg = goCn.GetFldVal("Select RmCtg from RmMst where RmCd='" + mRmCd + "'" + IIF(mPrtKey = ctCurrPrtn, " and RmPrtKey='" + ctCurrPrtn + "'", ""))
  wRmSCtg = goCn.GetFldVal("Select RmSCtg from RmMst where RmCd='" + mRmCd + "'" + IIF(mPrtKey = ctCurrPrtn, " and RmPrtKey='" + ctCurrPrtn + "'", ""))
  wRmKT = goCn.GetFldVal("Select RmKT from RmMst where RmCd='" + mRmCd + "'" + IIF(mPrtKey = ctCurrPrtn, " and RmPrtKey='" + ctCurrPrtn + "'", ""))
  wRmPurity = IIF(mOpt = "S", "RmPurityRt", "RmPurityWt")
  wRt = IIF(mOpt = "S", "RrSalRt", "RrCstRt")
  
  '****** Sachin 2.13.0 - 29-05-2006 ******
  wQwCS = IIF(mOpt = "S", "RmQw", "RmQwCst")
  wQw = goCn.GetFldVal("Select " + wQwCS + " From RmMst Where RmCd='" + mRmCd + "'" + IIF(mPrtKey = ctCurrPrtn, " and RmPrtKey='" + ctCurrPrtn + "'", ""))
  '****** Sachin 2.13.0 - 29-05-2006 ******
  
  '*** Jay 3.1.3 [CRP]
  If mOpt = "S" Then
    wCustRtPurity = goCn.GetFldVal("Select RrSalRt from RmRt where RrTcTyp = 'CRP' and RrCd= '" + mRmCd + "' " + _
                                        "and RrCmCtg= '" + ms_CmCtg + "' and RrCmCd= '" + mCmCd + "'" + IIF(mPrtKey = ctCurrPrtn, " and RrPrtKey='" + ctCurrPrtn + "'", ""))
  End If
  '*** Jay 3.1.3 [CRP]

  '******** AlyRt *********
  '**** for Sale  = (1-RmPurityRt)* mAlyRt
  '**** for Cost  = (1-RmPurityWt)* mAlyRt
  '*** Jay 3.1.3 [CRP]
  wAlyRt = (1 - IIF(wCustRtPurity <> 0 And mOpt = "S", wCustRtPurity, goCn.GetFldVal("Select " + wRmPurity + " from RmMst where RmCd='" + mRmCd + "'" + IIF(mPrtKey = ctCurrPrtn, " and RmPrtKey='" + ctCurrPrtn + "'", "")))) _
           * mAlyRt
  '*** Jay 3.1.3 [CRP]
  '********** Increments, only if mOpt='S'
'Manoj
'  If mOpt = "S" And Trim(ms_CmCtg) = "C" Then
  If mOpt = "S" And ((Trim(ms_CmCtg) = "C") Or (Trim(ms_CmCtg) = "S")) Then
'    wInc = gocn.GetFldVal("Select RrSalRt from RmRt where RrTcTyp= 'INC' " + _
'           "and RrCtg= '" + wRmCtg + "' and RrCmCtg= 'C' and RrCmCd='" + mCmCd + "'")
    wInc = goCn.GetFldVal("Select RrSalRt from RmRt where RrTcTyp= 'INC' " + _
           "and RrCtg= '" + wRmCtg + "' and RrCmCtg= '" & Trim$(ms_CmCtg) & "' and RrCmCd='" + mCmCd + "'")
'Manoj
  Else
    wInc = 0
  End If
  
  ' ****** Bef 3.03 - Manali
  'wCmLkUpRmRt = goCn.GetFldVal("Select CmLkUpRmRt from CustMst where CmCtg= '" + ms_CmCtg + "' and CmCd= '" + mCmCd + "'")

  ' ***** Manali - 3.03 - 02/07/08
  '6.1-13 seek from custmst is removed
  'If mCmLkUpRmRt = "" Then _
  mCmLkUpRmRt = goCn.GetFldVal("Select CmLkUpRmRt from CustMst where CmCtg= '" + ms_CmCtg + "' and CmCd= '" + mCmCd + "'")
  
  ' ***** Manali 3.8.0 - Cost Prc for CmCtg='S' from CmCtg='C'
  If ms_CmCtg = "S" And mOpt = "C" Then ms_CmCtg = "C"

  '********* looking for per gm record ('RM') for price List customer ***********
  '***** Look for record with the specified DmCtg. If not found then look for DmCtg= ''
'Manoj
'  If Trim(ms_CmCtg) = "C" Then
  If (Trim(ms_CmCtg) = "C") Or (Trim(ms_CmCtg) = "S") Then
'Manoj
    ' Zubin 213 (RrCmCurCd Cnd added)
    wGldRt = goCn.GetFldVal("Select isnull" + _
             "((select " + wRt + "  from RmRt where RrTcTyp= 'RM' and " + IIF(mPrtKey = ctCurrPrtn, " RrPrtKey='" + ctCurrPrtn + "' and ", "") + _
               "RrCtg= '" + wRmCtg + "' and RrSCtg= '" + wRmSCtg + "' and " + _
               "RrCd= '" + mRmCd + "' and RrFrLn<= " + CStr(mLme) + " And " + _
               "RrToLn>= " + CStr(mLme) + " and RrCmCtg= 'P' and RrCmCd= '" + mCmLkUpRmRt + "' and " + _
               "RrCmCurCd= '" + mf_CmCurCd + "' and " + _
               "RrDmCtg= '" + wDmCtg + "' and " + wRt + " <> 0)" + _
              ", " + _
              "IsNull((Select " + wRt + "  from RmRt where RrTcTyp= 'RM' and " + IIF(mPrtKey = ctCurrPrtn, " RrPrtKey='" + ctCurrPrtn + "' and ", "") + _
                       "RrCtg= '" + wRmCtg + "' and RrSCtg= '" + wRmSCtg + "' and " + _
                       "RrCd= '" + mRmCd + "' and RrFrLn <= " + CStr(mLme) + " And " + _
                       "RrToLn>= " + CStr(mLme) + " and RrCmCtg= 'P' and RrCmCd= '" + mCmLkUpRmRt + "' and " + _
                       "RrCmCurCd= '" + mf_CmCurCd + "' and " + _
                       "RrDmCtg= '' and " + wRt + " <> 0),0)" + _
             ")")
  End If
  '********* per gm record is found from Customer record,
  If wGldRt <> 0 Then GetGldRt = (wGldRt + IIF(wQw = "W", wAlyRt, 0)): Exit Function
  
  '********* looking for per gm record ('RM') for defined customer ***********
  '***** Look for record with the specified DmCtg. If not found then look for DmCtg= ''
  ' Zubin 213 (RrCmCurCd Cnd added)
  wGldRt = goCn.GetFldVal("Select isnull" + _
           "((select " + wRt + "  from RmRt where RrTcTyp= 'RM' and " + IIF(mPrtKey = ctCurrPrtn, " RrPrtKey='" + ctCurrPrtn + "' and ", "") + _
              "RrCtg= '" + wRmCtg + "' and RrSCtg= '" + wRmSCtg + "' and " + _
              "RrCd= '" + mRmCd + "' and RrFrLn <= " + CStr(mLme) + " And " + _
              "RrToLn>= " + CStr(mLme) + " and RrCmCtg= '" + ms_CmCtg + "' and RrCmCd = '" + mCmCd + "' and " + _
              "RrCmCurCd= '" + mf_CmCurCd + "' and " + _
              "RrDmCtg= '" + wDmCtg + "' and " + wRt + " <> 0)" + _
            ", " + _
            "IsNull((Select " + wRt + "  from RmRt where RrTcTyp='RM' and " + IIF(mPrtKey = ctCurrPrtn, " RrPrtKey='" + ctCurrPrtn + "' and ", "") + _
                     "RrCtg= '" + wRmCtg + "' and RrSCtg= '" + wRmSCtg + "' and " + _
                     "RrCd='" + mRmCd + "' and RrFrLn <= " + CStr(mLme) + " And " + _
                     "RrToLn >= " + CStr(mLme) + " and RrCmCtg= '" + ms_CmCtg + "' and RrCmCd = '" + mCmCd + "' and " + _
                     "RrCmCurCd= '" + mf_CmCurCd + "' and " + _
                     "RrDmCtg='' and " + wRt + " <> 0),0)" + _
           ")")

  '********* per gm record is found from Customer record,
  If wGldRt <> 0 Then GetGldRt = (wGldRt + IIF(wQw = "W", wAlyRt, 0)): Exit Function

  '********* looking for per gm record for Self ***********
  '***** Look for record with the specified DmCtg. If not found then look for DmCtg= ''
'Manoj
'  If Trim(ms_CmCtg) = "C" Then
  If (Trim(ms_CmCtg) = "C") Or (Trim(ms_CmCtg) = "S") And (mOpt = "S" And wCmSeekZSelfRt) Then
'    wGldRt = gocn.GetFldVal("Select isnull" + _
'             "((select " + wRt + " from RmRt where RrTcTyp='RM' and " + _
'                "RrCtg= '" + wRmCtg + "' and RrSCtg= '" + wRmSCtg + "' and " + _
'                "RrCd='" + mRmCd + "' and RrFrLn <= " + CStr(mLme) + " And " + _
'                "RrToLn >= " + CStr(mLme) + " and RrCmCtg= 'C' and RrCmCd = '" + ctSelfCmCd + "' and " + _
'                "RrDmCtg='" + wDmCtg + "' and " + wRt + " <> 0)" + _
'               ", " + _
'               "IsNull((select " + wRt + " from RmRt where RrTcTyp='RM' and " + _
'                        "RrCtg= '" + wRmCtg + "' and RrSCtg= '" + wRmSCtg + "' and " + _
'                        "RrCd='" + mRmCd + "' and RrFrLn <= " + CStr(mLme) + " And " + _
'                        "RrToLn >= " + CStr(mLme) + " and RrCmCtg= 'C' and RrCmCd ='" + ctSelfCmCd + "' and " + _
'                        "RrDmCtg='' and " + wRt + "<> 0),0)" + _
'             ")")
    ' Zubin 213 (RrCmCurCd Cnd added)
    wGldRt = goCn.GetFldVal("Select isnull" + _
             "((select " + wRt + " from RmRt where RrTcTyp='RM' and " + IIF(mPrtKey = ctCurrPrtn, " RrPrtKey='" + ctCurrPrtn + "' and ", "") + _
                "RrCtg= '" + wRmCtg + "' and RrSCtg= '" + wRmSCtg + "' and " + _
                "RrCd='" + mRmCd + "' and RrFrLn <= " + CStr(mLme) + " And " + _
                "RrToLn >= " + CStr(mLme) + " and RrCmCtg= '" & ms_CmCtg & "' and RrCmCd = '" + ctSelfCmCd + "' and " + _
                "RrCmCurCd= '" + mf_CmCurCd + "' and " + _
                "RrDmCtg='" + wDmCtg + "' and " + wRt + " <> 0)" + _
               ", " + _
               "IsNull((select " + wRt + " from RmRt where RrTcTyp='RM' and " + IIF(mPrtKey = ctCurrPrtn, " RrPrtKey='" + ctCurrPrtn + "' and ", "") + _
                        "RrCtg= '" + wRmCtg + "' and RrSCtg= '" + wRmSCtg + "' and " + _
                        "RrCd='" + mRmCd + "' and RrFrLn <= " + CStr(mLme) + " And " + _
                        "RrToLn >= " + CStr(mLme) + " and RrCmCtg= '" & ms_CmCtg & "' and RrCmCd ='" + ctSelfCmCd + "' and " + _
                        "RrCmCurCd= '" + mf_CmCurCd + "' and " + _
                        "RrDmCtg='' and " + wRt + "<> 0),0)" + _
             ")")
'Manoj
  End If
  '********* per gm record is found from Self record,
  If wGldRt <> 0 Then GetGldRt = (wGldRt * (1 + wInc / 100) + IIF(wQw = "W", wAlyRt, 0)): Exit Function
  If wQw = "W" Then
    wLsCtg = goCn.GetFldVal("Select DmLsCtg from DsgMst where DmTcTyp= '" + mDmTcTyp + "' and " + _
                            "DmSz= '' and DmCd= '" + mDmCd + "'")
    '6.1-13 seek from custmst is removed
    'If mCmLkUpMetLs = "" Then _
      mCmLkUpMetLs = goCn.GetFldVal("Select CmLkUpMetLs from CustMst where CmCtg= '" + ms_CmCtg + "' and CmCd= '" + mCmCd + "'")
    
    'indigo.162-Ls calculation is removed from here and added as a new function - GetGldRtLsPer and getting Ls %
    '6.1-85 if mLsRtApplicableYN<>Y then it will ignore loss calculation
    If mLsRtApplicableYN = "Y" Then _
    wGldLs = GetGldRtLsPer(wRmCtg, ms_CmCtg, mCmCd, mDmCd, wRmKT, wDmCtg, wLsCtg, wRt, mf_CmCurCd, mCmLkUpMetLs, mPrtKey)
      
    wRmPurity = IIF(mOpt = "S", "RmPurityRt", "RmPurityWt")
    If IsMissing(mf_CTBCurCnv) Then
      ' Zubin 213 (Before 213)
      'wGldRt = MWLib.Div(mLme, gocn.GetFldVal("Select PNum from param where PTyp= 'OUNCE' " + _
                               "and PMCd='" + wRmCtg + "'")) _
               * gocn.GetFldVal("Select " + wRmPurity + " from RmMst where " + _
                 "RmCd='" + mRmCd + "'") _
               * (1 + wGldLs / 100) + wAlyRt
      ' Zubin 213
      '*** Jay 3.1.3 [CRP]
      wGldRt = MWLib.Div(mLme, goCn.GetFldVal("Select PNum from param where PTyp= 'GMFACTOR' " + _
                               "and PMCd='" + wRmCtg + "' and PSCd= '" + mf_CmCurCd + "' ")) _
               * IIF(wCustRtPurity <> 0 And mOpt = "S", wCustRtPurity, goCn.GetFldVal("Select " + wRmPurity + " from RmMst where " + _
                 "RmCd='" + mRmCd + "'")) _
               * (1 + wGldLs / 100) + wAlyRt
      '*** Jay 3.1.3 [CRP]
    Else
      ' Zubin 213 (Before 213)
      'wGldRt = MWLib.Div(mLme, gocn.GetFldVal("Select PNum from param where PTyp= 'OUNCE' " + _
                               "and PMCd= '" + wRmCtg + "'") * gocn.GetFldVal("Select PNum from " + _
                           "Param where PTyp= 'CURNCY' and PMCd= '" + ctDefCurncy + "'")) _
               * gocn.GetFldVal("Select " + wRmPurity + " from RmMst where " + _
                 "RmCd= '" + mRmCd + "'") _
               * (1 + wGldLs / 100) * mf_CTBCurCnv + wAlyRt
      ' Zubin 213
      '*** Jay 3.1.3 [CRP]
      wGldRt = MWLib.Div(mLme, goCn.GetFldVal("Select PNum from param where PTyp= 'GMFACTOR' " + _
                               "and PMCd= '" + wRmCtg + "'and PSCd= '" + mf_CmCurCd + "' ") _
                          * goCn.GetFldVal("Select PNum from Param where PTyp= 'CURNCY' and " + _
                              "PMCd= '" + mf_CmCurCd + "'")) _
               * IIF(wCustRtPurity <> 0 And mOpt = "S", wCustRtPurity, goCn.GetFldVal("Select " + wRmPurity + " from RmMst where " + _
                 "RmCd= '" + mRmCd + "'")) _
               * (1 + wGldLs / 100) * mf_CTBCurCnv + wAlyRt
      '*** Jay 3.1.3 [CRP]
       
    End If
    GetGldRt = wGldRt
  Else
    GetGldRt = 0
  End If
  
  
End Function
Public Function GetGldRtForLabCdPure(ByVal mCmCd As String, ByVal mLme As Single, _
                         ByVal mRmCd As String, ByVal mAlyRt As Single, _
                         ByVal mDmTcTyp As String, ByVal mDmCd As String, _
                         ByVal mOpt As String, ByVal ms_CmCtg As String, _
                         Optional ByVal mf_CTBCurCnv As Variant, _
                         Optional ByVal mf_CmCurCd As Variant, _
                         Optional ByVal mCmLkUpRmRt As String, _
                         Optional ByVal mCmLkUpMetLs As String, _
                         Optional ByVal mPrtKey As String) As Single
                         
  'Optional ByVal mf_CurCd As Variant = "RS"
  
  Dim wDmCtg As String, wLsCtg As String, wRmCtg As String, wRmSCtg As String
  Dim wRmPurity As String, wRt As String, wAlyRt As Double, wInc As Double
  Dim wGldRt As Double, wGldLs As Double, wRmKT As String
  Dim wCustRtPurity As Double '*** Jay 3.1.3 [CRP]
  Dim wCmSeekZSelfRt As Boolean
  
  '****** Sachin 2.13.0 - 29-05-2006 ******
  Dim wQwCS As String, wQw As String
  '****** Sachin 2.13.0 - 29-05-2006 ******
  
  ' ***** Manali 3.03 - 02/07/08 - variables shifted to func declaration
  'Dim wCmLkUpRmRt As String, wCmLkUpMetLs As String
  ' ***** Manali 3.03 - 02/07/08 - variables shifted to func declaration
  
  Dim wGldLsStr As String
  
  If (mOpt <> "C") And (mOpt <> "S") Then GetGldRtForLabCdPure = 0: Exit Function
  
  ' Zubin 213
  If IsMissing(mf_CmCurCd) Then mf_CmCurCd = goCn.GetFldVal("Select CmCurCd from CustMst where CmCtg= '" + ms_CmCtg + "' and CmCd= '" + mCmCd + "'")

  wDmCtg = goCn.GetFldVal("Select DmCtg from DsgMst where DmTcTyp= '" + mDmTcTyp + "' and DmSz= '' " + _
                         "and DmCd= '" + mDmCd + "'")
  wRmCtg = goCn.GetFldVal("Select RmCtg from RmMst where RmCd='" + mRmCd + "'" + IIF(mPrtKey = ctCurrPrtn, " and RmPrtKey='" + ctCurrPrtn + "'", ""))
  wRmSCtg = goCn.GetFldVal("Select RmSCtg from RmMst where RmCd='" + mRmCd + "'" + IIF(mPrtKey = ctCurrPrtn, " and RmPrtKey='" + ctCurrPrtn + "'", ""))
  wRmKT = goCn.GetFldVal("Select RmKT from RmMst where RmCd='" + mRmCd + "'" + IIF(mPrtKey = ctCurrPrtn, " and RmPrtKey='" + ctCurrPrtn + "'", ""))
  wRmPurity = IIF(mOpt = "S", "RmPurityRt", "RmPurityWt")
  wRt = IIF(mOpt = "S", "RrSalRt", "RrCstRt")
  
  '****** Sachin 2.13.0 - 29-05-2006 ******
  wQwCS = IIF(mOpt = "S", "RmQw", "RmQwCst")
  wQw = goCn.GetFldVal("Select " + wQwCS + " From RmMst Where RmCd='" + mRmCd + "'" + IIF(mPrtKey = ctCurrPrtn, " and RmPrtKey='" + ctCurrPrtn + "'", ""))
  '****** Sachin 2.13.0 - 29-05-2006 ******
  wQw = "W"
  '*** Jay 3.1.3 [CRP]
  If mOpt = "S" Then
    wCustRtPurity = goCn.GetFldVal("Select RrSalRt from RmRt where RrTcTyp = 'CRP' and RrCd= '" + mRmCd + "' " + _
                                        "and RrCmCtg= '" + ms_CmCtg + "' and RrCmCd= '" + mCmCd + "'" + IIF(mPrtKey = ctCurrPrtn, " and RrPrtKey='" + ctCurrPrtn + "'", ""))
  End If
  '*** Jay 3.1.3 [CRP]

  '******** AlyRt *********
  '**** for Sale  = (1-RmPurityRt)* mAlyRt
  '**** for Cost  = (1-RmPurityWt)* mAlyRt
  '*** Jay 3.1.3 [CRP]
  wAlyRt = (1 - IIF(wCustRtPurity <> 0 And mOpt = "S", wCustRtPurity, goCn.GetFldVal("Select " + wRmPurity + " from RmMst where RmCd='" + mRmCd + "'" + IIF(mPrtKey = ctCurrPrtn, " and RmPrtKey='" + ctCurrPrtn + "'", "")))) _
           * mAlyRt
  '*** Jay 3.1.3 [CRP]
  '********** Increments, only if mOpt='S'
'Manoj
'  If mOpt = "S" And Trim(ms_CmCtg) = "C" Then
  If mOpt = "S" And ((Trim(ms_CmCtg) = "C") Or (Trim(ms_CmCtg) = "S")) Then
'    wInc = gocn.GetFldVal("Select RrSalRt from RmRt where RrTcTyp= 'INC' " + _
'           "and RrCtg= '" + wRmCtg + "' and RrCmCtg= 'C' and RrCmCd='" + mCmCd + "'")
    wInc = goCn.GetFldVal("Select RrSalRt from RmRt where RrTcTyp= 'INC' " + _
           "and RrCtg= '" + wRmCtg + "' and RrCmCtg= '" & Trim$(ms_CmCtg) & "' and RrCmCd='" + mCmCd + "'")
'Manoj
  Else
    wInc = 0
  End If
  
  ' ****** Bef 3.03 - Manali
  'wCmLkUpRmRt = goCn.GetFldVal("Select CmLkUpRmRt from CustMst where CmCtg= '" + ms_CmCtg + "' and CmCd= '" + mCmCd + "'")
  
  ' ***** Manali - 3.03 - 02/07/08
  '6.1-13 seek from custmst is removed
  'If mCmLkUpRmRt = "" Then _
  mCmLkUpRmRt = goCn.GetFldVal("Select CmLkUpRmRt from CustMst where CmCtg= '" + ms_CmCtg + "' and CmCd= '" + mCmCd + "'")
  
  ' ***** Manali 3.8.0 - Cost Prc for CmCtg='S' from CmCtg='C'
  If ms_CmCtg = "S" And mOpt = "C" Then ms_CmCtg = "C"

  '********* looking for per gm record ('RM') for price List customer ***********
  '***** Look for record with the specified DmCtg. If not found then look for DmCtg= ''
'Manoj
'  If Trim(ms_CmCtg) = "C" Then
  If (Trim(ms_CmCtg) = "C") Or (Trim(ms_CmCtg) = "S") Then
'Manoj
    ' Zubin 213 (RrCmCurCd Cnd added)
'    wGldRt = goCn.GetFldVal("Select isnull" + _
             "((select " + wRt + "  from RmRt where RrTcTyp= 'RM' and " + IIF(mPrtKey = ctCurrPrtn, " RrPrtKey='" + ctCurrPrtn + "' and ", "") + _
               "RrCtg= '" + wRmCtg + "' and RrSCtg= '" + wRmSCtg + "' and " + _
               "RrCd= '" + mRmCd + "' and RrFrLn<= " + CStr(mLme) + " And " + _
               "RrToLn>= " + CStr(mLme) + " and RrCmCtg= 'P' and RrCmCd= '" + mCmLkUpRmRt + "' and " + _
               "RrCmCurCd= '" + mf_CmCurCd + "' and " + _
               "RrDmCtg= '" + wDmCtg + "' and " + wRt + " <> 0)" + _
              ", " + _
              "IsNull((Select " + wRt + "  from RmRt where RrTcTyp= 'RM' and " + IIF(mPrtKey = ctCurrPrtn, " RrPrtKey='" + ctCurrPrtn + "' and ", "") + _
                       "RrCtg= '" + wRmCtg + "' and RrSCtg= '" + wRmSCtg + "' and " + _
                       "RrCd= '" + mRmCd + "' and RrFrLn <= " + CStr(mLme) + " And " + _
                       "RrToLn>= " + CStr(mLme) + " and RrCmCtg= 'P' and RrCmCd= '" + mCmLkUpRmRt + "' and " + _
                       "RrCmCurCd= '" + mf_CmCurCd + "' and " + _
                       "RrDmCtg= '' and " + wRt + " <> 0),0)" + _
             ")")
  End If
  '********* per gm record is found from Customer record,
  If wGldRt <> 0 Then GetGldRtForLabCdPure = (wGldRt + IIF(wQw = "W", wAlyRt, 0)): Exit Function
  
  '********* looking for per gm record ('RM') for defined customer ***********
  '***** Look for record with the specified DmCtg. If not found then look for DmCtg= ''
  ' Zubin 213 (RrCmCurCd Cnd added)
 ' wGldRt = goCn.GetFldVal("Select isnull" + _
           "((select " + wRt + "  from RmRt where RrTcTyp= 'RM' and " + IIF(mPrtKey = ctCurrPrtn, " RrPrtKey='" + ctCurrPrtn + "' and ", "") + _
              "RrCtg= '" + wRmCtg + "' and RrSCtg= '" + wRmSCtg + "' and " + _
              "RrCd= '" + mRmCd + "' and RrFrLn <= " + CStr(mLme) + " And " + _
              "RrToLn>= " + CStr(mLme) + " and RrCmCtg= '" + ms_CmCtg + "' and RrCmCd = '" + mCmCd + "' and " + _
              "RrCmCurCd= '" + mf_CmCurCd + "' and " + _
              "RrDmCtg= '" + wDmCtg + "' and " + wRt + " <> 0)" + _
            ", " + _
            "IsNull((Select " + wRt + "  from RmRt where RrTcTyp='RM' and " + IIF(mPrtKey = ctCurrPrtn, " RrPrtKey='" + ctCurrPrtn + "' and ", "") + _
                     "RrCtg= '" + wRmCtg + "' and RrSCtg= '" + wRmSCtg + "' and " + _
                     "RrCd='" + mRmCd + "' and RrFrLn <= " + CStr(mLme) + " And " + _
                     "RrToLn >= " + CStr(mLme) + " and RrCmCtg= '" + ms_CmCtg + "' and RrCmCd = '" + mCmCd + "' and " + _
                     "RrCmCurCd= '" + mf_CmCurCd + "' and " + _
                     "RrDmCtg='' and " + wRt + " <> 0),0)" + _
           ")")

  '********* per gm record is found from Customer record,
  If wGldRt <> 0 Then GetGldRtForLabCdPure = (wGldRt + IIF(wQw = "W", wAlyRt, 0)): Exit Function

  '********* looking for per gm record for Self ***********
  '***** Look for record with the specified DmCtg. If not found then look for DmCtg= ''
'Manoj
'  If Trim(ms_CmCtg) = "C" Then
  If (Trim(ms_CmCtg) = "C") Or (Trim(ms_CmCtg) = "S") And (mOpt = "S" And wCmSeekZSelfRt) Then
'    wGldRt = gocn.GetFldVal("Select isnull" + _
'             "((select " + wRt + " from RmRt where RrTcTyp='RM' and " + _
'                "RrCtg= '" + wRmCtg + "' and RrSCtg= '" + wRmSCtg + "' and " + _
'                "RrCd='" + mRmCd + "' and RrFrLn <= " + CStr(mLme) + " And " + _
'                "RrToLn >= " + CStr(mLme) + " and RrCmCtg= 'C' and RrCmCd = '" + ctSelfCmCd + "' and " + _
'                "RrDmCtg='" + wDmCtg + "' and " + wRt + " <> 0)" + _
'               ", " + _
'               "IsNull((select " + wRt + " from RmRt where RrTcTyp='RM' and " + _
'                        "RrCtg= '" + wRmCtg + "' and RrSCtg= '" + wRmSCtg + "' and " + _
'                        "RrCd='" + mRmCd + "' and RrFrLn <= " + CStr(mLme) + " And " + _
'                        "RrToLn >= " + CStr(mLme) + " and RrCmCtg= 'C' and RrCmCd ='" + ctSelfCmCd + "' and " + _
'                        "RrDmCtg='' and " + wRt + "<> 0),0)" + _
'             ")")
    ' Zubin 213 (RrCmCurCd Cnd added)
  '  wGldRt = goCn.GetFldVal("Select isnull" + _
             "((select " + wRt + " from RmRt where RrTcTyp='RM' and " + IIF(mPrtKey = ctCurrPrtn, " RrPrtKey='" + ctCurrPrtn + "' and ", "") + _
                "RrCtg= '" + wRmCtg + "' and RrSCtg= '" + wRmSCtg + "' and " + _
                "RrCd='" + mRmCd + "' and RrFrLn <= " + CStr(mLme) + " And " + _
                "RrToLn >= " + CStr(mLme) + " and RrCmCtg= '" & ms_CmCtg & "' and RrCmCd = '" + ctSelfCmCd + "' and " + _
                "RrCmCurCd= '" + mf_CmCurCd + "' and " + _
                "RrDmCtg='" + wDmCtg + "' and " + wRt + " <> 0)" + _
               ", " + _
               "IsNull((select " + wRt + " from RmRt where RrTcTyp='RM' and " + IIF(mPrtKey = ctCurrPrtn, " RrPrtKey='" + ctCurrPrtn + "' and ", "") + _
                        "RrCtg= '" + wRmCtg + "' and RrSCtg= '" + wRmSCtg + "' and " + _
                        "RrCd='" + mRmCd + "' and RrFrLn <= " + CStr(mLme) + " And " + _
                        "RrToLn >= " + CStr(mLme) + " and RrCmCtg= '" & ms_CmCtg & "' and RrCmCd ='" + ctSelfCmCd + "' and " + _
                        "RrCmCurCd= '" + mf_CmCurCd + "' and " + _
                        "RrDmCtg='' and " + wRt + "<> 0),0)" + _
             ")")
'Manoj
  End If
  '********* per gm record is found from Self record,
  If wGldRt <> 0 Then GetGldRtForLabCdPure = (wGldRt * (1 + wInc / 100) + IIF(wQw = "W", wAlyRt, 0)): Exit Function
  If wQw = "W" Then
    wLsCtg = goCn.GetFldVal("Select DmLsCtg from DsgMst where DmTcTyp= '" + mDmTcTyp + "' and " + _
                            "DmSz= '' and DmCd= '" + mDmCd + "'")
    '6.1-13 seek from custmst is removed
    'If mCmLkUpMetLs = "" Then _
      mCmLkUpMetLs = goCn.GetFldVal("Select CmLkUpMetLs from CustMst where CmCtg= '" + ms_CmCtg + "' and CmCd= '" + mCmCd + "'")
    
    'indigo.162-Ls calculation is removed from here and added as a new function - GetGldRtLsPer and getting Ls %
    wGldLs = GetGldRtLsPer(wRmCtg, ms_CmCtg, mCmCd, mDmCd, wRmKT, wDmCtg, wLsCtg, wRt, mf_CmCurCd, mCmLkUpMetLs, mPrtKey)
      
    wRmPurity = IIF(mOpt = "S", "RmPurityRt", "RmPurityWt")
    If IsMissing(mf_CTBCurCnv) Then
      ' Zubin 213 (Before 213)
      'wGldRt = MWLib.Div(mLme, gocn.GetFldVal("Select PNum from param where PTyp= 'OUNCE' " + _
                               "and PMCd='" + wRmCtg + "'")) _
               * gocn.GetFldVal("Select " + wRmPurity + " from RmMst where " + _
                 "RmCd='" + mRmCd + "'") _
               * (1 + wGldLs / 100) + wAlyRt
      ' Zubin 213
      '*** Jay 3.1.3 [CRP]
      wGldLs = 0
      wAlyRt = 0
      wGldRt = MWLib.Div(mLme, goCn.GetFldVal("Select PNum from param where PTyp= 'GMFACTOR' " + _
                               "and PMCd='" + wRmCtg + "' and PSCd= '" + mf_CmCurCd + "' ")) _
               * IIF(wCustRtPurity <> 0 And mOpt = "S", 1, 1) _
               * (1 + wGldLs / 100) + wAlyRt
      '*** Jay 3.1.3 [CRP]
    Else
      ' Zubin 213 (Before 213)
      'wGldRt = MWLib.Div(mLme, gocn.GetFldVal("Select PNum from param where PTyp= 'OUNCE' " + _
                               "and PMCd= '" + wRmCtg + "'") * gocn.GetFldVal("Select PNum from " + _
                           "Param where PTyp= 'CURNCY' and PMCd= '" + ctDefCurncy + "'")) _
               * gocn.GetFldVal("Select " + wRmPurity + " from RmMst where " + _
                 "RmCd= '" + mRmCd + "'") _
               * (1 + wGldLs / 100) * mf_CTBCurCnv + wAlyRt
      ' Zubin 213
      '*** Jay 3.1.3 [CRP]
      wGldLs = 0
      wAlyRt = 0
      wCustRtPurity = 1
      wGldRt = MWLib.Div(mLme, goCn.GetFldVal("Select PNum from param where PTyp= 'GMFACTOR' " + _
                               "and PMCd= '" + wRmCtg + "'and PSCd= '" + mf_CmCurCd + "' ") _
                          * goCn.GetFldVal("Select PNum from Param where PTyp= 'CURNCY' and " + _
                              "PMCd= '" + mf_CmCurCd + "'")) _
               * IIF(wCustRtPurity <> 0 And mOpt = "S", wCustRtPurity, 1) * (1 + wGldLs / 100) * mf_CTBCurCnv + wAlyRt
               
     
    End If
    GetGldRtForLabCdPure = wGldRt
  Else
    GetGldRtForLabCdPure = 0
  End If
End Function

Public Function GetGldRtLsPer(ByVal mRmCtg As String, ByVal ms_CmCtg As String, ByVal mCmCd As String, ByVal mDmCd As String, _
                        ByVal mRmKt As String, ByVal mDmCtg As String, ByVal mLsCtg As String, ByVal mOpt As String, _
                        Optional ByVal mf_CmCurCd As Variant, Optional ByVal mCmLkUpMetLs As String, Optional ByVal mPrtKey As String) As Single
  
  Dim wGldLsStr As String
  
  Dim wRt As String
  wRt = mOpt
'  If wRt = "S" Then wRt = "RrSalRt"
'  If wRt = "C" Then wRt = "RrCstRt"
  
'  wRt = IIF(wRt = "S", "RrSalRt", "RrCstRt")
  
  
  Dim wGldLs As Double
  wGldLsStr = "Select IsNull" + _
          "((select " + wRt + " from RmRt where RrTcTyp='LS' and RrCtg='" + mRmCtg + "' and " + IIF(mPrtKey = ctCurrPrtn, " RrPrtKey='" + ctCurrPrtn + "' and ", "") + _
             "RrCmCtg= 'P' and RrCmCd = '" + mCmLkUpMetLs + "' and RrDmCd= '" + mDmCd + "'" + _
             " and RrRmKT = '" + mRmKt + "' and RrCmCurCd='" + "" + "' and " + wRt + "<> 0), " + _
            "IsNull ((select " + wRt + " from RmRt where RrTcTyp='LS' and RrCtg='" + mRmCtg + "' and " + IIF(mPrtKey = ctCurrPrtn, " RrPrtKey='" + ctCurrPrtn + "' and ", "") + _
               "RrCmCtg= 'P' and RrCmCd = '" + mCmLkUpMetLs + "' and RrDmCd= '" + mDmCd + "'" + _
               " and RrRmKT = '' and RrCmCurCd='" + "" + "' and " + wRt + "<> 0), "
        
    
  wGldLsStr = wGldLsStr + "IsNull" + _
           "((select " + wRt + " from RmRt where RrTcTyp='LS' and RrCtg='" + mRmCtg + "' and " + IIF(mPrtKey = ctCurrPrtn, " RrPrtKey='" + ctCurrPrtn + "' and ", "") + _
           "RrCmCtg= 'P' and RrCmCd = '" + mCmLkUpMetLs + "' and RrDmCtg= '" + mDmCtg + "' and RrDmCd = '' and " + _
           "RrLsCtg= '" + mLsCtg + "' and RrRmKT = '" + mRmKt + "' and RrCmCurCd='" + "" + "' and " + wRt + "<> 0), " + _
             "IsNull((select " + wRt + " from RmRt where RrTcTyp='LS' and RrCtg='" + mRmCtg + "' and " + IIF(mPrtKey = ctCurrPrtn, " RrPrtKey='" + ctCurrPrtn + "' and ", "") + _
             "RrCmCtg= 'P' and RrCmCd = '" + mCmLkUpMetLs + "' and RrDmCtg= '" + mDmCtg + "' and RrDmCd = '' and " + _
             "RrLsCtg= '" + mLsCtg + "' and RrRmKT = '' and RrCmCurCd='" + "" + "' and " + wRt + "<> 0), " + _
              "IsNull((select " + wRt + " from RmRt where RrTcTyp= 'LS' and RrCtg='" + mRmCtg + "' and " + IIF(mPrtKey = ctCurrPrtn, " RrPrtKey='" + ctCurrPrtn + "' and ", "") + _
                    "RrCmCtg= 'P' and RrCmCd = '" + mCmLkUpMetLs + "' and RrDmCtg= '" + mDmCtg + "' and RrDmCd = '' and " + _
                    "RrLsCtg= '' and RrRmKT = '" + mRmKt + "' and RrCmCurCd='" + "" + "' and " + wRt + "<> 0), " + _
               "IsNull((select " + wRt + " from RmRt where RrTcTyp= 'LS' and RrCtg='" + mRmCtg + "' and " + IIF(mPrtKey = ctCurrPrtn, " RrPrtKey='" + ctCurrPrtn + "' and ", "") + _
                    "RrCmCtg= 'P' and RrCmCd = '" + mCmLkUpMetLs + "' and RrDmCtg= '" + mDmCtg + "' and RrDmCd = '' and " + _
                    "RrLsCtg= '' and RrRmKT = '' and RrCmCurCd='" + "" + "' and " + wRt + "<> 0), " + _
                   "IsNull((select " + wRt + " from RmRt where RrTcTyp= 'LS' and RrCtg='" + mRmCtg + "' and " + IIF(mPrtKey = ctCurrPrtn, " RrPrtKey='" + ctCurrPrtn + "' and ", "") + _
                           "RrCmCtg= 'P' and RrCmCd = '" + mCmLkUpMetLs + "' and RrDmCtg='' and RrDmCd = '' and " + _
                           "RrLsCtg= '' and RrRmKT = '' and RrCmCurCd='" + "" + "' and " + wRt + "<> 0),0)))))))"
                           
  wGldLs = goCn.GetFldVal(wGldLsStr)
  
  If wGldLs = 0 Then
      wGldLsStr = "Select IsNull" + _
          "((select " + wRt + " from RmRt where RrTcTyp='LS' and RrCtg='" + mRmCtg + "' and " + IIF(mPrtKey = ctCurrPrtn, " RrPrtKey='" + ctCurrPrtn + "' and ", "") + _
            "RrCmCtg= '" + ms_CmCtg + "' and RrCmCd = '" + mCmCd + "' and RrDmCd= '" + mDmCd + "'" + _
            " and RrRmKT = '" + mRmKt + "' and RrCmCurCd='" + "" + "' and " + wRt + "<> 0), " + _
            "IsNull ((select " + wRt + " from RmRt where RrTcTyp='LS' and RrCtg='" + mRmCtg + "' and " + IIF(mPrtKey = ctCurrPrtn, " RrPrtKey='" + ctCurrPrtn + "' and ", "") + _
               "RrCmCtg= '" + ms_CmCtg + "' and RrCmCd = '" + mCmCd + "' and RrDmCd= '" + mDmCd + "'" + _
               " and RrRmKT = '' and RrCmCurCd='" + "" + "' and " + wRt + "<> 0), "
  
      wGldLsStr = wGldLsStr + "IsNull" + _
              "((select " + wRt + " from RmRt where RrTcTyp= 'LS' and RrCtg='" + mRmCtg + "' and " + IIF(mPrtKey = ctCurrPrtn, " RrPrtKey='" + ctCurrPrtn + "' and ", "") + _
              "RrCmCtg= '" + ms_CmCtg + "' and RrCmCd = '" + mCmCd + "' and RrDmCtg= '" + mDmCtg + "' and RrDmCd = '' and " + _
              "RrLsCtg= '" + mLsCtg + "' and RrRmKT = '" + mRmKt + "' and RrCmCurCd='" + "" + "' and " + wRt + "<> 0), " + _
                  "IsNull ((select " + wRt + " from RmRt where RrTcTyp= 'LS' and RrCtg='" + mRmCtg + "' and " + IIF(mPrtKey = ctCurrPrtn, " RrPrtKey='" + ctCurrPrtn + "' and ", "") + _
                  "RrCmCtg= '" + ms_CmCtg + "' and RrCmCd = '" + mCmCd + "' and RrDmCtg= '" + mDmCtg + "' and RrDmCd = '' and " + _
                  "RrLsCtg= '" + mLsCtg + "' and RrRmKT = '' and RrCmCurCd='" + "" + "' and " + wRt + "<> 0), " + _
                    "IsNull ((select " + wRt + " from RmRt where RrTcTyp= 'LS' and RrCtg='" + mRmCtg + "' and " + IIF(mPrtKey = ctCurrPrtn, " RrPrtKey='" + ctCurrPrtn + "' and ", "") + _
                    "RrCmCtg= '" + ms_CmCtg + "' and RrCmCd = '" + mCmCd + "' and RrDmCtg= '" + mDmCtg + "' and RrDmCd = '' and " + _
                    "RrLsCtg= '' and RrRmKT = '" + mRmKt + "' and RrCmCurCd='" + "" + "' and " + wRt + "<> 0), " + _
                      "IsNull ((select " + wRt + " from RmRt where RrTcTyp= 'LS' and RrCtg='" + mRmCtg + "' and " + IIF(mPrtKey = ctCurrPrtn, " RrPrtKey='" + ctCurrPrtn + "' and ", "") + _
                      "RrCmCtg= '" + ms_CmCtg + "' and RrCmCd = '" + mCmCd + "' and RrDmCtg= '" + mDmCtg + "' and RrDmCd = '' and " + _
                      "RrLsCtg= '' and RrRmKT = '' and RrCmCurCd='" + "" + "' and " + wRt + "<> 0), " + _
                          "IsNull ((select " + wRt + " from RmRt where RrTcTyp= 'LS' and RrCtg='" + mRmCtg + "' and " + IIF(mPrtKey = ctCurrPrtn, " RrPrtKey='" + ctCurrPrtn + "' and ", "") + _
                          "RrCmCtg= '" + ms_CmCtg + "' and RrCmCd = '" + mCmCd + "' and RrDmCtg= '' and RrDmCd = '' and " + _
                          "RrLsCtg= '' and RrRmKT = '' and RrCmCurCd='" + "" + "' and " + wRt + "<> 0),0)))))))"
      
      
      wGldLs = goCn.GetFldVal(wGldLsStr)
  End If

  If wGldLs = 0 Then
      wGldLsStr = "Select IsNull" + _
          "((select " + wRt + " from RmRt where RrTcTyp='LS' and RrCtg='" + mRmCtg + "' and " + IIF(mPrtKey = ctCurrPrtn, " RrPrtKey='" + ctCurrPrtn + "' and ", "") + _
            "RrCmCtg= '" + ms_CmCtg + "' and RrCmCd = '" + ctSelfCmCd + "' and RrDmCd= '" + mDmCd + "'" + _
            " and RrRmKT = '" + mRmKt + "' and RrCmCurCd='" + mf_CmCurCd + "' and " + wRt + "<> 0), " + _
            "IsNull ((select " + wRt + " from RmRt where RrTcTyp='LS' and RrCtg='" + mRmCtg + "' and " + IIF(mPrtKey = ctCurrPrtn, " RrPrtKey='" + ctCurrPrtn + "' and ", "") + _
               "RrCmCtg= '" + ms_CmCtg + "' and RrCmCd = '" + ctSelfCmCd + "' and RrDmCd= '" + mDmCd + "'" + _
               " and RrRmKT = '' and RrCmCurCd='" + mf_CmCurCd + "' and " + wRt + "<> 0), "
  
      wGldLsStr = wGldLsStr + "IsNull" + _
              "((select " + wRt + " from RmRt where RrTcTyp= 'LS' and RrCtg='" + mRmCtg + "' and " + IIF(mPrtKey = ctCurrPrtn, " RrPrtKey='" + ctCurrPrtn + "' and ", "") + _
              "RrCmCtg= '" + ms_CmCtg + "' and RrCmCd = '" + ctSelfCmCd + "' and RrDmCtg= '" + mDmCtg + "' and RrDmCd = '' and " + _
              "RrLsCtg= '" + mLsCtg + "' and RrRmKT = '" + mRmKt + "' and RrCmCurCd='" + mf_CmCurCd + "' and " + wRt + "<> 0), " + _
                  "IsNull ((select " + wRt + " from RmRt where RrTcTyp= 'LS' and RrCtg='" + mRmCtg + "' and " + IIF(mPrtKey = ctCurrPrtn, " RrPrtKey='" + ctCurrPrtn + "' and ", "") + _
                  "RrCmCtg= '" + ms_CmCtg + "' and RrCmCd = '" + ctSelfCmCd + "' and RrDmCtg= '" + mDmCtg + "' and RrDmCd = '' and " + _
                  "RrLsCtg= '" + mLsCtg + "' and RrRmKT = '' and RrCmCurCd='" + mf_CmCurCd + "' and " + wRt + "<> 0), " + _
                    "IsNull ((select " + wRt + " from RmRt where RrTcTyp= 'LS' and RrCtg='" + mRmCtg + "' and " + IIF(mPrtKey = ctCurrPrtn, " RrPrtKey='" + ctCurrPrtn + "' and ", "") + _
                    "RrCmCtg= '" + ms_CmCtg + "' and RrCmCd = '" + ctSelfCmCd + "' and RrDmCtg= '" + mDmCtg + "' and RrDmCd = '' and " + _
                    "RrLsCtg= '' and RrRmKT = '" + mRmKt + "' and RrCmCurCd='" + mf_CmCurCd + "' and " + wRt + "<> 0), " + _
                      "IsNull ((select " + wRt + " from RmRt where RrTcTyp= 'LS' and RrCtg='" + mRmCtg + "' and " + IIF(mPrtKey = ctCurrPrtn, " RrPrtKey='" + ctCurrPrtn + "' and ", "") + _
                      "RrCmCtg= '" + ms_CmCtg + "' and RrCmCd = '" + ctSelfCmCd + "' and RrDmCtg= '" + mDmCtg + "' and RrDmCd = '' and " + _
                      "RrLsCtg= '' and RrRmKT = '' and RrCmCurCd='" + mf_CmCurCd + "' and " + wRt + "<> 0), " + _
                          "IsNull ((select " + wRt + " from RmRt where RrTcTyp= 'LS' and RrCtg='" + mRmCtg + "' and " + IIF(mPrtKey = ctCurrPrtn, " RrPrtKey='" + ctCurrPrtn + "' and ", "") + _
                          "RrCmCtg= '" + ms_CmCtg + "' and RrCmCd = '" + ctSelfCmCd + "' and RrDmCtg= '' and RrDmCd = '' and " + _
                          "RrLsCtg= '' and RrRmKT = '' and RrCmCurCd='" + mf_CmCurCd + "' and " + wRt + "<> 0),0)))))))"
      wGldLs = goCn.GetFldVal(wGldLsStr)
  End If
  If wGldLs = 0 Then
      wGldLsStr = "Select IsNull" + _
          "((select " + wRt + " from RmRt where RrTcTyp='LS' and RrCtg='" + mRmCtg + "' and " + IIF(mPrtKey = ctCurrPrtn, " RrPrtKey='" + ctCurrPrtn + "' and ", "") + _
            "RrCmCtg= '" + ms_CmCtg + "' and RrCmCd = '" + ctSelfCmCd + "' and RrDmCd= '" + mDmCd + "'" + _
            " and RrRmKT = '" + mRmKt + "' and RrCmCurCd='" + "" + "' and " + wRt + "<> 0), " + _
            "IsNull ((select " + wRt + " from RmRt where RrTcTyp='LS' and RrCtg='" + mRmCtg + "' and " + IIF(mPrtKey = ctCurrPrtn, " RrPrtKey='" + ctCurrPrtn + "' and ", "") + _
               "RrCmCtg= '" + ms_CmCtg + "' and RrCmCd = '" + ctSelfCmCd + "' and RrDmCd= '" + mDmCd + "'" + _
               " and RrRmKT = '' and RrCmCurCd='" + "" + "' and " + wRt + "<> 0), "
  
      wGldLsStr = wGldLsStr + "IsNull" + _
              "((select " + wRt + " from RmRt where RrTcTyp= 'LS' and RrCtg='" + mRmCtg + "' and " + IIF(mPrtKey = ctCurrPrtn, " RrPrtKey='" + ctCurrPrtn + "' and ", "") + _
              "RrCmCtg= '" + ms_CmCtg + "' and RrCmCd = '" + ctSelfCmCd + "' and RrDmCtg= '" + mDmCtg + "' and RrDmCd = '' and " + _
              "RrLsCtg= '" + mLsCtg + "' and RrRmKT = '" + mRmKt + "' and RrCmCurCd='" + "" + "' and " + wRt + "<> 0), " + _
                  "IsNull ((select " + wRt + " from RmRt where RrTcTyp= 'LS' and RrCtg='" + mRmCtg + "' and " + IIF(mPrtKey = ctCurrPrtn, " RrPrtKey='" + ctCurrPrtn + "' and ", "") + _
                  "RrCmCtg= '" + ms_CmCtg + "' and RrCmCd = '" + ctSelfCmCd + "' and RrDmCtg= '" + mDmCtg + "' and RrDmCd = '' and " + _
                  "RrLsCtg= '" + mLsCtg + "' and RrRmKT = '' and RrCmCurCd='" + "" + "' and " + wRt + "<> 0), " + _
                    "IsNull ((select " + wRt + " from RmRt where RrTcTyp= 'LS' and RrCtg='" + mRmCtg + "' and " + IIF(mPrtKey = ctCurrPrtn, " RrPrtKey='" + ctCurrPrtn + "' and ", "") + _
                    "RrCmCtg= '" + ms_CmCtg + "' and RrCmCd = '" + ctSelfCmCd + "' and RrDmCtg= '" + mDmCtg + "' and RrDmCd = '' and " + _
                    "RrLsCtg= '' and RrRmKT = '" + mRmKt + "' and RrCmCurCd='" + "" + "' and " + wRt + "<> 0), " + _
                      "IsNull ((select " + wRt + " from RmRt where RrTcTyp= 'LS' and RrCtg='" + mRmCtg + "' and " + IIF(mPrtKey = ctCurrPrtn, " RrPrtKey='" + ctCurrPrtn + "' and ", "") + _
                      "RrCmCtg= '" + ms_CmCtg + "' and RrCmCd = '" + ctSelfCmCd + "' and RrDmCtg= '" + mDmCtg + "' and RrDmCd = '' and " + _
                      "RrLsCtg= '' and RrRmKT = '' and RrCmCurCd='" + "" + "' and " + wRt + "<> 0), " + _
                          "IsNull ((select " + wRt + " from RmRt where RrTcTyp= 'LS' and RrCtg='" + mRmCtg + "' and " + IIF(mPrtKey = ctCurrPrtn, " RrPrtKey='" + ctCurrPrtn + "' and ", "") + _
                          "RrCmCtg= '" + ms_CmCtg + "' and RrCmCd = '" + ctSelfCmCd + "' and RrDmCtg= '' and RrDmCd = '' and " + _
                          "RrLsCtg= '' and RrRmKT = '' and RrCmCurCd='" + "" + "' and " + wRt + "<> 0),0)))))))"
      wGldLs = goCn.GetFldVal(wGldLsStr)
  End If
  GetGldRtLsPer = wGldLs
End Function
'Rn.158 added optional parameter to show message when RTC
Public Function WithinRTCRange(Optional ByVal pShowMsgYN As String) As Boolean
  Dim wRsRTC As MwfLib.MDORowSet, wmsgstr As String, wStr As String
  WithinRTCRange = False
  wStr = "Select * from Param where PTyp= 'RTC' and PMCd= 'RTC' and " + _
          "(convert (int, PNum) <= DATEPART (hh, getdate()) ) and (convert (int, PNum1) > DATEPART (hh, getdate()) )"
  'MW.105 - Checking RTC time peroid
  If Not go_HlpCn.RecSeek(wStr) Then
    WithinRTCRange = True
  End If
  'MW.105 - Display message when RTC
  If WithinRTCRange = False Then
    'Rn.158 if ShowMsg is No then does not display error message
    If pShowMsgYN <> "N" Then
      Set wRsRTC = go_HlpCn.OpenResultset(wStr)
      Do While Not (wRsRTC.EOF Or wRsRTC.BOF)
        wmsgstr = wmsgstr & "          " & "Between " & _
                IIF(wRsRTC!PNum >= 13, wRsRTC!PNum - 12, IIF((wRsRTC!PNum) < 1, wRsRTC!PNum + 12, wRsRTC!PNum)) & _
                IIF(wRsRTC!PNum >= 12, " PM", " AM") & " to " & _
                IIF(wRsRTC!PNum1 >= 13, wRsRTC!PNum1 - 12, IIF((wRsRTC!PNum1) < 1, wRsRTC!PNum1 + 12, wRsRTC!PNum1)) & _
                IIF(wRsRTC!PNum1 >= 12, " PM.", " AM.") & vbCrLf
        wRsRTC.MoveNext
      Loop
      MsgBox "Report Cannot Be Taken During These Times. " & vbNewLine & wmsgstr & "Contact Your EMR Admin."
    End If
  End If

End Function
'stellar.1- it checks whether Quotes are exists in text box or not
Public Function ChkQuotes(ByVal ActiveFrm As Form, ByVal pChkQtsIdName As String, ByVal pChkQtsVal As String) As String
  Dim wChkQtsStr As String, wChkQtsIdNameArr() As String, i As Integer, wChkQtsIdName As String
  Dim ADC As MWCTL_ADC
  Set ADC = ActiveFrm.ADC
  If pChkQtsVal <> "" Then
    If InStr(1, pChkQtsVal, "'") Or InStr(1, pChkQtsVal, """") > 0 Then
      wChkQtsIdName = pChkQtsIdName
    End If
  Else
    wChkQtsIdNameArr = Split(pChkQtsIdName, "|")
    For i = 0 To UBound(wChkQtsIdNameArr)
      wChkQtsStr = ADC(wChkQtsIdNameArr(i))
      If InStr(1, wChkQtsStr, "'") Or InStr(1, wChkQtsStr, """") > 0 Then
        wChkQtsIdName = wChkQtsIdNameArr(i)
        Exit For
      End If
    Next
  End If
  If wChkQtsIdName <> "" Then
    ChkQuotes = ADC(wChkQtsIdName).ToolTipText + " Without Quotes"
  Else
    ChkQuotes = ""
  End If
  Set ADC = Nothing
End Function

'sv.9 getting value for purge date
Public Function GetPurgeDt(ByVal pCoCd As String, ByVal pSysCd As String) As String
  If pSysCd = "JT" Then
    GetPurgeDt = GetMDIFrm.ADC.Connection.GetFldVal("Select vPDesc225 from vParam where vPCoCd='ZZZ' and vPTyp= 'JPURGEDT' and vPMCd= 'JPURGEDT' and vPSCd =''")
  Else
    GetPurgeDt = GetMDIFrm.ADC.Connection.GetFldVal("Select vPDesc225 from vParam where vPCoCd='" + pCoCd + "' and vPTyp= 'PURGEDT' and vPMCd= 'PURGEDT' and vPSCd =''")
  End If
  If GetPurgeDt = "" Then GetPurgeDt = "01/01/80"
     
End Function

'Uni.4 - SalCtg scope will check all 3 Sal Ctg(SalCtg, SalCtg2 & SalCtg3).
'if include SalCtg2 and SalCtg3 is YEs then checks all SalCtg else behave as it is "
Public Function GetDmSalCtgWhrCnd(ByVal pSalCtgFr As String, ByVal pSalCtgTo As String, ByVal pSalCtgSel As String, ByVal pAllSalCtgYN As String) As String
  Dim wSalCtgTo As String, wSalCtgSel As String
  wSalCtgTo = IIF(pSalCtgTo <> "", pSalCtgTo, "ZZZZZ")
          
  If pAllSalCtgYN = "Y" Then
    GetDmSalCtgWhrCnd = " and ((DmSalCtg>='" + pSalCtgFr + "' and DmSalCtg<='" + wSalCtgTo + "' " + IIF(pSalCtgSel <> "", " and DmSalCtg in(" + pSalCtgSel + ")", "") + ") " + _
                        "Or (DmSalCtg2>='" + pSalCtgFr + "' and DmSalCtg2<='" + wSalCtgTo + "' " + IIF(pSalCtgSel <> "", " and DmSalCtg2 in(" + pSalCtgSel + ")", "") + ") " + _
                        "Or (DmSalCtg3>='" + pSalCtgFr + "' and DmSalCtg3<='" + wSalCtgTo + "' " + IIF(pSalCtgSel <> "", " and DmSalCtg3 in(" + pSalCtgSel + ")", "") + ") ) "
  Else
    GetDmSalCtgWhrCnd = " and DmSalCtg>='" + pSalCtgFr + "' and DmSalCtg<='" + wSalCtgTo + "' " + IIF(pSalCtgSel <> "", " and DmSalCtg in(" + pSalCtgSel + ")", "")
    
  End If
End Function

'Uni.4 selecting Dmsalctg either from SalCtg,SalCtg2 or SalCtg3
'if include SalCtg2 and SalCtg3 is YEs then checks all SalCtg else behave as it is
Public Function GetDmSalCtg(ByVal pSalCtgFr As String, ByVal pSalCtgTo As String, ByVal pSalCtgSel As String, ByVal pAllSalCtgYN As String, ByVal pGrpYN As String) As String
  Dim wSalCtgTo As String, wDmSalCtg1Cnd As String, wDmSalCtg2Cnd As String, wDmSalCtg3Cnd As String
  Dim wDmSalCtg As String, wDmSalCtg2 As String, wDmSalCtg3 As String
  
  wSalCtgTo = IIF(pSalCtgTo <> "", pSalCtgTo, "ZZZZZ")
  wDmSalCtg = IIF(pGrpYN = "Y", "max(DmSalCtg)", "DmSalCtg")
  wDmSalCtg2 = IIF(pGrpYN = "Y", "max(DmSalCtg2)", "DmSalCtg2")
  wDmSalCtg3 = IIF(pGrpYN = "Y", "max(DmSalCtg3)", "DmSalCtg3")
                      
  If pAllSalCtgYN = "Y" And (pSalCtgFr <> "" Or pSalCtgTo <> "" Or pSalCtgSel <> "") Then
    wDmSalCtg1Cnd = wDmSalCtg + ">='" + pSalCtgFr + "' and " + wDmSalCtg + "<='" + wSalCtgTo + "'" + IIF(pSalCtgSel <> "", " and " + wDmSalCtg + " In(" + pSalCtgSel + ")", "")
    wDmSalCtg2Cnd = wDmSalCtg2 + ">='" + pSalCtgFr + "' and " + wDmSalCtg2 + "<='" + wSalCtgTo + "'" + IIF(pSalCtgSel <> "", " and " + wDmSalCtg2 + " In(" + pSalCtgSel + ")", "")
    wDmSalCtg3Cnd = wDmSalCtg3 + ">='" + pSalCtgFr + "' and " + wDmSalCtg3 + "<='" + wSalCtgTo + "'" + IIF(pSalCtgSel <> "", " and " + wDmSalCtg3 + " In(" + pSalCtgSel + ")", "")

    GetDmSalCtg = "(case when " + wDmSalCtg1Cnd + " then " + wDmSalCtg + " else " + _
                  "(case when " + wDmSalCtg2Cnd + " then " + wDmSalCtg2 + " else " + _
                  "(case when " + wDmSalCtg3Cnd + " then " + wDmSalCtg3 + " end)end)end)"
  Else
    GetDmSalCtg = wDmSalCtg
  End If

End Function

'Public Function PrepareRfIdString(ByVal pRfId As String) As String
 ' pRfId = Replace(pRfId, "'", "")
 ' pRfId = Replace(pRfId, vbCrLf, ",")
 ' Do While InStr(1, pRfId, ",,") <> 0
 '   pRfId = Replace(pRfId, ",,", ",")
 ' Loop
 ' pRfId = IIF(Left(pRfId, 1) = ",", Mid(pRfId, 2, Len(pRfId) - 1), pRfId)
 ' pRfId = IIF(Right(pRfId, 1) = ",", Left(pRfId, Len(pRfId) - 1), pRfId)
  
  'pRfId = Replace(pRfId, ",", "','")
 
  'pRfId = IIF(Left(pRfId, 1) <> "'", "'", "") + pRfId
  'pRfId = pRfId + IIF(Right(pRfId, 1) <> "'", "'", "")
 
  'pRfId = Replace(pRfId, ",'',", ",")
  'PrepareRfIdString = pRfId

'End Function

Public Function PrepareRfIdString(ByVal pRfId As String, Optional ByVal pQuotesYN As String) As String
  pRfId = Replace(pRfId, "'", "")
  pRfId = Replace(pRfId, vbCrLf, ",")
  Do While InStr(1, pRfId, ",,") <> 0
    pRfId = Replace(pRfId, ",,", ",")
  Loop
  pRfId = IIF(Left(pRfId, 1) = ",", Mid(pRfId, 2, Len(pRfId) - 1), pRfId)
  pRfId = IIF(Right(pRfId, 1) = ",", Left(pRfId, Len(pRfId) - 1), pRfId)
  
  If pQuotesYN = "Y" Then
    pRfId = Replace(pRfId, ",", "','")
 
    pRfId = IIF(Left(pRfId, 1) <> "'", "'", "") + pRfId
    pRfId = pRfId + IIF(Right(pRfId, 1) <> "'", "'", "")
  End If
 
  pRfId = Replace(pRfId, ",'',", ",")
  PrepareRfIdString = pRfId

End Function
'sv.63 it will take QR code for given invoice
Public Sub SetRptQRCOde(ByVal mf_OriginalHt As Single, ByVal mf_OriginalWdth As Single, _
                      ByVal mo_PicObj As Object, ByVal ms_CoCd As String, _
                      ByVal ms_InTc As String, ByVal ms_InYy As String, ByVal ms_InChr As String, ByVal ms_InNo As String)
                      
On Error GoTo ErrRptPic
    
  Dim wPth As String, wFlNm As String, wFlPth As String
  Dim wPicNewHt As Single, wPicNewWdt As Single
      
  Set mo_PicObj.FormattedPicture = LoadPicture("")
  wFlNm = ms_CoCd + "_" + ms_InTc + "_" + ms_InYy + "_" + ms_InChr + "_" + CStr(ms_InNo)
  
  wPth = goCn.GetFldVal("Select PDesc225 from Param Where PTyp='PCT' and PMCd='3D'") + "\eInvoice\" + ms_CoCd
  wFlPth = wPth + "\" + wFlNm + ".bmp"

  If wFlPth <> "" And Dir(wFlPth) <> "" Then
    Set mo_PicObj.FormattedPicture = LoadPicture(wFlPth)
    wPicNewHt = mo_PicObj.FormattedPicture.Height: wPicNewWdt = mo_PicObj.FormattedPicture.Width

    If wPicNewHt > wPicNewWdt Then
      mo_PicObj.Height = mf_OriginalHt
      mo_PicObj.Width = mf_OriginalWdth * MWLib.Div(wPicNewWdt, wPicNewHt)
    Else
      mo_PicObj.Width = mf_OriginalWdth
      mo_PicObj.Height = mf_OriginalHt * MWLib.Div(wPicNewHt, wPicNewWdt)
    End If
    Exit Sub
  End If
  
  Exit Sub
ErrRptPic:
  If Err.Number = 481 Then
      DispMsg "Improper Picture Format, Please Change the Picture File", etInfo
  Else
      DispMsg Err.Description, etInfo
  End If
  Set mo_PicObj.FormattedPicture = LoadPicture("")
End Sub
'578.19 if Rmcode length more than 12 then font size will be 5 else existing font size
Public Function GetRmcdFontSz(ByVal pRmCd As String, ByVal pFontSz As Double) As String
  GetRmcdFontSz = " if Len(" + pRmCd + ")>12 then 5 else " + CStr(pFontSz) + ""

End Function

'6.1_27 copy picture to the given target path
Public Sub PicCopy(ByVal pRsPicCpy As MDORowSet, ByVal pTarPicPath As String, Optional ByVal pDsgMstYN As String, Optional ByVal p3DHsk As String)
Dim i As Integer, wSrcPicPath As String, wSrcPic As String, wTarPic As String
  With pRsPicCpy
    .MoveFirst
    Do While Not (pRsPicCpy.EOF Or pRsPicCpy.BOF)
      'it will copy 250 picture only and then exit
      If i > 250 Then MsgBox "Copied 250 Picures Only ": Exit Sub
      'if source is dsgmst then takes values from design table else other tables
      If pDsgMstYN = "Y" Then
        'it checks picture format
        wSrcPicPath = GetPictPath(!DmTcTyp, IIF(UCase(p3DHsk) = "3D" Or p3DHsk = "", False, True), !DmCd, !DmSz)
        If wSrcPicPath = "" Or Dir(wSrcPicPath) = "" Then wSrcPicPath = GetPictPath(!DmTcTyp, IIF(UCase(p3DHsk) = "3D" Or p3DHsk = "", True, False), !DmCd, !DmSz)
      Else
        'it checks picture format
        wSrcPicPath = GetPictPath(!DmTcTyp, IIF(UCase(p3DHsk) = "3D" Or p3DHsk = "", False, True), !OdDmCd, !OdDmSz, !OmCoCd, !OmTc, !OmYy, !OmChr, !OmNo, !OdSr, !OdPicNm)
        If wSrcPicPath = "" Or Dir(wSrcPicPath) = "" Then wSrcPicPath = GetPictPath(!DmTcTyp, IIF(UCase(p3DHsk) = "3D" Or p3DHsk = "", True, False), !OdDmCd, !OdDmSz, !OmCoCd, !OmTc, !OmYy, !OmChr, !OmNo, !OdSr, !OdPicNm)
      End If
           
      'if source file exists then start copy
      If Dir(wSrcPicPath) <> "" Then
        wSrcPic = goCn.GetFldVal("SELECT RIGHT('" + wSrcPicPath + "',CHARINDEX('\', (REVERSE('" + wSrcPicPath + "')) )  )")
        'taking source file name and it is used as target picture name
        wTarPic = pTarPicPath + wSrcPic
        'if picture already copied then ignore the copy
        If Dir(wTarPic) = "" Then
          FileCopy wSrcPicPath, wTarPic
          i = i + 1
        End If
      End If
      pRsPicCpy.MoveNext
    Loop
  End With
    
End Sub

'6.1_51 checking that given voucher number is exist in unlock ptyp,
'if given voucher is in unlock ptyp it will not check lock date
Public Function ChkUnLockVch(ByVal pCoCd As String, ByVal pVchrNo As String) As Boolean
  ChkUnLockVch = goCn.RecSeek("select 'x' from vParam where vPCoCd ='" + pCoCd + "'" + _
                            " and vPTyp ='UNLOCK' and vPMCd ='UNLOCK' and vPSCd =''" + _
                            " and ',' + vPDesc225 + ',' like '%," + pVchrNo + ",%'")

End Function

Public Function DtFmtSlahVal() As Integer
DtFmtSlahVal = IIF(UCase(MWLib.DtSetting) = "MM/DD/YY", 1, 3)

End Function
Public Function DtFmtSlahYyyyVal() As Integer
DtFmtSlahYyyyVal = IIF(UCase(MWLib.DtSetting) = "MM/DD/YY", 101, 103)

End Function
Public Function DtFmtDashVal() As Integer
DtFmtDashVal = IIF(UCase(MWLib.DtSetting) = "MM/DD/YY", 10, 5)

End Function

Public Function DtShortStr() As String
DtShortStr = IIF(UCase(MWLib.DtSetting) = "MM/DD/YY", "MM/DD/YY", "DD/MM/YY")

End Function
Public Function DtFullStr() As String
DtFullStr = IIF(UCase(MWLib.DtSetting) = "MM/DD/YY", "MMddyyHHmmss", "ddMMyyHHmmss")

End Function
Public Function DtShortStrWithDash() As String
DtShortStrWithDash = IIF(UCase(MWLib.DtSetting) = "MM/DD/YY", "MM-DD-YY", "DD-MM-YY")

End Function

Public Function DtShortStrYyyy() As String
DtShortStrYyyy = IIF(UCase(MWLib.DtSetting) = "MM/DD/YY", "MM/DD/YYYY", "DD/MM/YYYY")

End Function
Public Function DtFmtStr() As String
DtFmtStr = IIF(UCase(MWLib.DtSetting) = "MM/DD/YY", "MDY", "DMY")

End Function

