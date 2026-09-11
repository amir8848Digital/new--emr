VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "EmrClsGrph"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit
Enum en_GrphValFldLst
  DispNm = 0
  DBFldNm = 1
  RepFormula = 2
  TFilter = 3
End Enum

'local variable(s) to hold property value(s)
Private mc_Vals As New Collection
Private mGrphFrm As Form 'local copy
Private mGrphRep As Report 'local copy
'*** (Jen 04/02/08)
Private mClsCn As MDOConnection 'local copy    '*** (Bef speed) bef 24/11/06
'*** (Jen 04/02/08)
Private Const INCHTWIPS = 1440         '  Twips/inch
Private Const CENTITWIPS = 567         '  Twips/cm

'*** (Jen 04/02/08)
Public Sub SetFormDets(ByRef pFrm As Form, ByRef pConn As MDOConnection) '*** (Bef speed) bef 24/11/06
'*** (Jen 04/02/08)
'*** (Bef 04/02/08)
'Public Sub SetFormDets(ByRef pFrm As Form)
'*** (Bef 04/02/08)
    Set mGrphFrm = pFrm
    '*** (Jen 04/02/08)
    Set mClsCn = pConn   '*** (Bef speed) bef 24/11/06
    '*** (Jen 04/02/08)
End Sub
Public Property Set GrphRep(ByRef vData As Report)
'used when assigning an Object to the property, on the left side of a Set statement.
'Syntax: Set x.GrphRep = Form1
    Set mGrphRep = vData
End Property
Public Sub AddItem(ByVal DispStrg As String, ByVal DBFldNm As String, ByVal RepFormula As String, ByVal TFilter As String)
    Dim wv_ArrObj As Variant
    wv_ArrObj = Array(DispStrg, DBFldNm, RepFormula, TFilter)
    mc_Vals.Add wv_ArrObj, DispStrg
End Sub
Public Function GetVal(ByVal DispStrg As String, ByVal ValFld As en_GrphValFldLst) As String
  GetVal = mc_Vals(DispStrg)(ValFld)
End Function
Public Sub FillValCmb(ByRef ValCmb As ComboBox)
  Dim ws_AddStrg As String, i As Integer
  ValCmb.Clear
  For i = 1 To mc_Vals.Count
    ws_AddStrg = mc_Vals(i)(0)
    If ws_AddStrg <> "" Then
      ValCmb.AddItem ws_AddStrg
    End If
  Next i
End Sub

Public Sub SetSecVisible()
  Dim wSec As Section
  If mGrphFrm.adc("UoShowGrph") = "Y" Then
    For Each wSec In mGrphRep.Sections
      If wSec.Name <> "RH" And wSec.Name <> "SECGRPH" Then wSec.Suppress = True
    Next
  ElseIf mGrphFrm.adc("UoShowGrph") = "N" Then
    mGrphRep.Sections("SECGRPH").Suppress = True
  End If
End Sub
Public Sub NonKeyWhen()
  If mGrphFrm.adc("UoShowGrph") = "Y" Then
    mGrphFrm.ChkGrph.Value = vbChecked
  ElseIf mGrphFrm.adc("UoShowGrph") = "N" Then
    mGrphFrm.ChkGrph.Value = vbUnchecked
  End If
  FillValCmb mGrphFrm.cboGrphValue
  mGrphFrm.cboGrphValue.ListIndex = mGrphFrm.adc("UoGrphValInd")
  mGrphFrm.TAB_GRPH.Tab = 0
  mGrphFrm.FraTabGrph(1).Enabled = False
End Sub
Public Sub NonKeyValid(ByRef pr_Cancel As Boolean, ByRef pr_ErrMsg As String)
  Dim wiGrpNo As Integer
  Call SetDefaVal
  If UCase$(mGrphFrm.adc("UoShowGrph")) = "Y" Then
    wiGrpNo = SetGrphGrpNo
    If (Not mClsCn.RecSeek("Select * from Param Where PTyp = 'GRTYP' And PMCD = '" + mGrphFrm!adc("UoGrphTyp") + "' And " + _
                " (PNum1 =" + CStr(wiGrpNo) + " Or PNum1 = 0)")) Or wiGrpNo = 0 Then _
      pr_Cancel = True: pr_ErrMsg = "No of Groups & Graph Type Combination Not Proper."
    If pr_Cancel = False And _
      Not (mClsCn.RecSeek("Select * from Param Where PTyp = 'GRTYP' And PMCD = '" + mGrphFrm!adc("UoGrphTyp") + "'")) Then _
        pr_Cancel = True: pr_ErrMsg = "Graph Type Not Defined."
    If pr_Cancel = False And _
      Not (mClsCn.RecSeek("Select * from Param Where PTyp = 'GRDIR' And PMCD = '" + mGrphFrm!adc("UoGrphDir") + "'")) Then _
        pr_Cancel = True: pr_ErrMsg = "Graph Direction Not Defined."
    If pr_Cancel = False And _
      Not (mClsCn.RecSeek("Select * from Param Where PTyp = 'GRDATALB' And PMCD = '" + mGrphFrm!adc("UoGrphDataLb") + "'")) Then _
        pr_Cancel = True: pr_ErrMsg = "Graph Data Label Not Defined."
    If pr_Cancel = False And _
      Not (mClsCn.RecSeek("Select * from Param Where PTyp = 'GRVWANGL' And PMCD = '" + mGrphFrm!adc("UoGrphVwAngl") + "'")) Then _
        pr_Cancel = True: pr_ErrMsg = "Graph View Angle Not Defined."
    If pr_Cancel = False And _
      Not (mClsCn.RecSeek("Select * from Param Where PTyp = 'CRPPRSZ' And PMCD = '" + mGrphFrm!adc("UoGrphPSZ") + "'")) Then _
        pr_Cancel = True: pr_ErrMsg = "Graph Paper Size Not Defined."
    If pr_Cancel = False And _
      Not (mClsCn.RecSeek("Select * from Param Where PTyp = 'CRORIENT' And PMCD = '" + mGrphFrm!adc("UoGrphOrient") + "'")) Then _
        pr_Cancel = True: pr_ErrMsg = "Graph Paper Orientation Not Defined."
    If pr_Cancel = False And _
      Not (mClsCn.RecSeek("Select * from Param Where PTyp = 'YN' And PMCD = '" + mGrphFrm!adc("UoGrphClr") + "'")) Then _
        pr_Cancel = True: pr_ErrMsg = "Graph Color (Y/N) Not Defined."
    
    If pr_Cancel = False And mGrphFrm!adc("UoGrphSumm") <> "" And _
      Not (mClsCn.RecSeek("Select * from Param Where PTyp = 'GRSUMM' And PMCD = '" + mGrphFrm!adc("UoGrphSumm") + "'")) Then _
        pr_Cancel = True: pr_ErrMsg = "Group Top N Summary Not Defined."
    If pr_Cancel = False And mGrphFrm!adc("UoGrphSumm") = "" And _
      Not (mClsCn.RecSeek("Select * from Param Where PTyp = 'GRTOPN' And PMCD = '" + mGrphFrm!adc("UoGrphG1TopN") + "' And PNum = -1")) Then _
        pr_Cancel = True: pr_ErrMsg = "Group 1 Top N Should be 'GRTNNO'."
    If pr_Cancel = False And mGrphFrm!adc("UoGrphSumm") <> "" And _
      Not (mClsCn.RecSeek("Select * from Param Where PTyp = 'GRTOPN' And PMCD = '" + mGrphFrm!adc("UoGrphG1TopN") + "'")) Then _
        pr_Cancel = True: pr_ErrMsg = "Group 1 Top N Not Defined."
    If pr_Cancel = False And mGrphFrm!adc("UoGrphSumm") = "" And _
      Not (mClsCn.RecSeek("Select * from Param Where PTyp = 'GRTOPN' And PMCD = '" + mGrphFrm!adc("UoGrphG2TopN") + "' And PNum = -1")) Then _
        pr_Cancel = True: pr_ErrMsg = "Group " + IIF(wiGrpNo = 1, "1", "2") + " Top N Should be 'GRTNNO'."
    If pr_Cancel = False And mGrphFrm!adc("UoGrphSumm") <> "" And _
      Not (mClsCn.RecSeek("Select * from Param Where PTyp = 'GRTOPN' And PMCD = '" + mGrphFrm!adc("UoGrphG2TopN") + "'")) Then _
        pr_Cancel = True: pr_ErrMsg = "Group " + IIF(wiGrpNo = 1, "1", "2") + " Top N Not Defined."
  
  End If
End Sub
Public Function SetGrphGrpNo() As Integer
  Dim wiGrpNo As Integer
  wiGrpNo = IIF(UCase$(mGrphFrm.GRP_REP.Value(gltGroup3)) <> UCase$("(None)"), 2, IIF(UCase$(mGrphFrm.GRP_REP.Value(gltGroup2)) <> UCase$("(None)"), 1, 0))
'  If wiGrpNo <> 0 Then gs_GrphGrpNo = wiGrpNo
  SetGrphGrpNo = wiGrpNo
End Function
Public Sub SetDefaVal()
  If UCase$(mGrphFrm.adc("UoShowGrph")) = "Y" Then
    If UCase$(mGrphFrm.GRP_REP.Value(gltGroup2)) = UCase$("(None)") Then
      mGrphFrm.GRP_REP.Value(gltGroup2) = "(Detail)"
    End If
    If UCase$(mGrphFrm.GRP_REP.Value(gltGroup3)) = UCase$("(None)") And UCase$(mGrphFrm.GRP_REP.Value(gltGroup2)) <> UCase$("(Detail)") Then
      mGrphFrm.GRP_REP.Value(gltGroup3) = "(Detail)"
    End If
    If mGrphFrm.adc("UoGrphTyp") = "" Then _
      mGrphFrm.adc("UoGrphTyp") = mClsCn.GetFldVal("Select PMCD From Param Where PTyp = 'GrTyp' And PNum = -1")
    If mGrphFrm.adc("UoGrphDir") = "" Then _
      mGrphFrm.adc("UoGrphDir") = mClsCn.GetFldVal("Select PMCD From Param Where PTyp = 'GRDIR' And PNum = -1")
    If mGrphFrm.adc("UoGrphDataLb") = "" Then _
      mGrphFrm.adc("UoGrphDataLb") = mClsCn.GetFldVal("Select PMCD From Param Where PTyp = 'GRDATALB' And PNum = -1")
    If mGrphFrm!adc("UoGrphVwAngl") = "" Then _
      mGrphFrm!adc("UoGrphVwAngl") = mClsCn.GetFldVal("Select PMCD From Param Where PTyp = 'GRVWANGL' And PNum = -1")
    If mGrphFrm!adc("UoGrphClr") = "" Then mGrphFrm!adc("UoGrphClr") = "Y"
    If mGrphFrm!adc("UoGrphPSZ") = "" Then _
      mGrphFrm!adc("UoGrphPSZ") = mClsCn.GetFldVal("Select PMCD From Param Where PTyp = 'CRPPRSZ' And PNum = -1")
    If mGrphFrm!adc("UoGrphOrient") = "" Then _
      mGrphFrm!adc("UoGrphOrient") = mClsCn.GetFldVal("Select PMCD From Param Where PTyp = 'CRORIENT' And PNum = -1")

    If mGrphFrm!adc("UoGrphG1TopN") = "" Then _
      mGrphFrm!adc("UoGrphG1TopN") = mClsCn.GetFldVal("Select PMCD From Param Where PTyp = 'GRTOPN' And PNum = -1")
    If mGrphFrm!adc("UoGrphG2TopN") = "" Then _
      mGrphFrm!adc("UoGrphG2TopN") = mClsCn.GetFldVal("Select PMCD From Param Where PTyp = 'GRTOPN' And PNum = -1")
    
    If mClsCn.RecSeek("Select PMCD From Param Where PTyp = 'GRTOPN' And PMCD ='" + mGrphFrm!adc("UoGrphG1TopN") + "' And PNum = -1") Then _
      mGrphFrm!adc("UoGrphG1NVal") = 0
    If mClsCn.RecSeek("Select PMCD From Param Where PTyp = 'GRTOPN' And PMCD ='" + mGrphFrm!adc("UoGrphG2TopN") + "' And PNum = -1") Then _
      mGrphFrm!adc("UoGrphG2NVal") = 0
    Call ShowFlds
  End If
End Sub

Public Function SetRepQuery(ps_QryStruct As String, ps_SqlStr As String, ps_MnuCd As String, ps_ValFld As String, ps_Filter As String, ps_Summ As String, ps_G1TopNType As String, pi_G1TopNVal As Integer, ps_G2TopNType As String, pi_G2TopNVal As Integer) As String
  
  Dim ws_TblNm As String, ws_FltTblNm As String, ws_QryStr As String, ws_ExecStr As String, ws_Summ As String
  Dim ws_G1TopN As String, wi_G1PerType As String, ws_G2TopN As String, wi_G2PerType As String
  
  'Temporary Table Names
  ws_TblNm = "#Tmp" + ps_MnuCd
  ws_FltTblNm = "#Tmp" + ps_MnuCd + "FLT"
  
  'Creating Temporary Tables
  '6.1 DMY changed in ws_QryStr - 2 place
  ws_QryStr = "  Set DateFormat " + DtFmtStr() + "" + vbCrLf + "SELECT * INTO " + ws_TblNm + " FROM (" + ps_QryStruct + ") as T"
  mClsCn.CreateTmpTable ws_TblNm, ws_QryStr
  ws_QryStr = "  Set DateFormat " + DtFmtStr() + "" + vbCrLf + " INSERT INTO " + ws_TblNm + " " + ps_SqlStr + ""
  mClsCn.Execute ws_QryStr
  mClsCn.CreateTmpTable ws_FltTblNm, "Select space(30) as Grp1, space(30) as Grp2 Into " + ws_FltTblNm + " From Param Where 1=2"
  
  'Setting Parameters
  Call SetTopNTyp(ps_G1TopNType, ws_G1TopN, wi_G1PerType)
  Call SetTopNTyp(ps_G2TopNType, ws_G2TopN, wi_G2PerType)
  
  ws_Summ = mClsCn.GetFldVal("Select PValue From Param Where PTYp = 'GRSUMM' AND PMCD = '" + ps_Summ + "'")
  
  'Calling Procedure
  ws_ExecStr = "dbo.MwSp_FillFltTbl '" + ws_TblNm + "', '" + ws_FltTblNm + "', '" + ps_ValFld + "', '" + ps_Filter + "', '" + ws_Summ + "', '" + ws_G1TopN + "', " + CStr(pi_G1TopNVal) + ", " + CStr(wi_G1PerType) + ", '" + ws_G2TopN + "', " + CStr(pi_G2TopNVal) + ", " + CStr(wi_G2PerType)
  mClsCn.Execute (ws_ExecStr)
  
  'Setting Report Query
  SetRepQuery = "Select * From " + ws_TblNm + " as Tbl  Where Exists (Select 1 From " + ws_FltTblNm + " Flt Where Tbl.Grp1 = Flt.Grp1 And Tbl.Grp2 = Flt.Grp2)"
End Function
Public Sub SetTopNTyp(ByVal ps_TopNType As String, ByRef ws_TopN As String, wi_PerTyp As Variant)
  Select Case UCase$(ps_TopNType)
  Case "GRBNPER"
    ws_TopN = "BOTTOM"
    wi_PerTyp = 1
    Case "GRBOTN"
    ws_TopN = "BOTTOM"
    wi_PerTyp = 0
  Case "GRTNALL"
    ws_TopN = "ALL"
    wi_PerTyp = 0
  Case "GRTNNO"
    ws_TopN = "NO"
    wi_PerTyp = 0
  Case "GRTNPER"
    ws_TopN = "TOP"
    wi_PerTyp = 1
  Case "GRTOPN"
    ws_TopN = "TOP"
    wi_PerTyp = 0
  End Select
End Sub

Public Sub SetGraphProp(pRep As Report, ps_QryStruct As String, ps_SqlStr As String)
On Error GoTo ErrHndlr
  Dim wsGrpSrt As String, wsGrp() As String
  Dim wo_Grph As GraphObject, wdHeight As Double, wdWidth As Double, wsTmp As String, wsATmp() As String, wiTmp As Integer
  Set mGrphRep = pRep
  Call SetSecVisible
  If mGrphFrm!ChkGrph.Value = vbChecked Then
    Dim wiGrphOptVal As Integer
    wsGrpSrt = mGrphFrm!GRP_REP.GrpOpts
    wsGrp = Split(wsGrpSrt, ",")
    If Trim(wsGrp(0)) <> "" Then
      Set wo_Grph = mGrphRep.Grph2
      mGrphRep.Grph1.Suppress = True
    ElseIf Trim(wsGrp(0)) = "" Then
      Set wo_Grph = mGrphRep.Grph1
      mGrphRep.Grph2.Suppress = True
    End If
    With wo_Grph
      
      wiGrphOptVal = mClsCn.GetFldVal("Select PNum From Param Where PTyp = 'CRPPRSZ' And PmCd = '" + mGrphFrm!adc("UoGrphPSZ") + "'")
      If wiGrphOptVal <> -1 Then mGrphRep.PaperSize = wiGrphOptVal
      If Not mClsCn.RecSeek("Select PValue From Param Where PTYP = 'CRPPRSZ' AND PNum =" + CStr(mGrphRep.PaperSize)) Then _
        mGrphRep.PaperSize = 9
      wiGrphOptVal = mClsCn.GetFldVal("Select PNum From Param Where PTyp = 'CRORIENT' And PmCd = '" + mGrphFrm!adc("UoGrphOrient") + "'")
      If wiGrphOptVal <> -1 Then mGrphRep.PaperOrientation = wiGrphOptVal
      
      wsTmp = mClsCn.GetFldVal("Select PValue From Param Where PTYP = 'CRPPRSZ' AND PNum =" + CStr(mGrphRep.PaperSize))
      wsATmp = Split(wsTmp, " X ")
      If UBound(wsATmp) - LBound(wsATmp) >= 1 Then
        If IsNumeric(wsATmp(0)) Then wdWidth = CDbl(wsATmp(0))
        If IsNumeric(wsATmp(1)) Then wdHeight = CDbl(wsATmp(1))
      End If
      
      .Left = 0
      .Top = 0
      If mGrphRep.PaperOrientation = crLandscape Then
        mGrphRep.LeftMargin = CLng(0.5 * INCHTWIPS)
        mGrphRep.TopMargin = CLng(0.3 * INCHTWIPS)
        mGrphRep.RightMargin = CLng(0.5 * INCHTWIPS)
        mGrphRep.BottomMargin = CLng(0.3 * INCHTWIPS)
        .Height = CLng(wdWidth * INCHTWIPS) - mGrphRep.TopMargin - mGrphRep.BottomMargin - mGrphRep.Sections("RH").Height
        .Width = CLng((wdHeight - 0.2) * INCHTWIPS) - mGrphRep.LeftMargin - mGrphRep.RightMargin
      ElseIf mGrphRep.PaperOrientation = crPortrait Then
        mGrphRep.LeftMargin = CLng(0.3 * INCHTWIPS)
        mGrphRep.TopMargin = CLng(0.5 * INCHTWIPS)
        mGrphRep.RightMargin = CLng(0.3 * INCHTWIPS)
        mGrphRep.BottomMargin = CLng(0.5 * INCHTWIPS)
        .Height = CLng((wdHeight - 0.2) * INCHTWIPS) - mGrphRep.TopMargin - mGrphRep.BottomMargin - mGrphRep.Sections("RH").Height
        .Width = CLng(wdWidth * INCHTWIPS) - mGrphRep.LeftMargin - mGrphRep.RightMargin
      End If
      
      wiGrphOptVal = mClsCn.GetFldVal("Select PNum From Param Where PTyp = 'GrTyp' And PmCd = '" + mGrphFrm!adc("UoGrphTyp") + "'")
      If wiGrphOptVal <> -1 Then .GraphType = wiGrphOptVal
      wiGrphOptVal = mClsCn.GetFldVal("Select PNum From Param Where PTyp = 'GrDir' And PmCd = '" + mGrphFrm!adc("UoGrphDir") + "'")
      If wiGrphOptVal <> -1 Then .GraphDirection = wiGrphOptVal
      wiGrphOptVal = mClsCn.GetFldVal("Select PNum From Param Where PTyp = 'GrDataLb' And PmCd = '" + mGrphFrm!adc("UoGrphDataLb") + "'")
      If wiGrphOptVal <> -1 Then .DataPoint = wiGrphOptVal
      wiGrphOptVal = mClsCn.GetFldVal("Select PNum From Param Where PTyp = 'GrVwAngl' And PmCd = '" + mGrphFrm!adc("UoGrphVwAngl") + "'")
      If wiGrphOptVal <> -1 Then .ViewingAngle = wiGrphOptVal
      
      If UCase$(mGrphFrm!adc("UoGrphClr")) = "Y" Then
        .GraphColor = crColorGraph
      ElseIf UCase$(mGrphFrm!adc("UoGrphClr")) = "N" Then
        .GraphColor = crBlackAndWhiteGraph
      End If
      
      If mGrphFrm!adc("UoGrphMTitle") <> "" Then .Title = mGrphFrm!adc("UoGrphMTitle")
      If mGrphFrm!adc("UoGrphSTitle") <> "" Then .SubTitle = mGrphFrm!adc("UoGrphSTitle")
      If mGrphFrm!adc("UoGrphFtNote") <> "" Then .FootNote = mGrphFrm!adc("UoGrphFtNote")
      If mGrphFrm!adc("UoGrphGTitle") <> "" Then .GroupsTitle = mGrphFrm!adc("UoGrphGTitle")
      If mGrphFrm!adc("UoGrphDTitle") <> "" Then .DataTitle = mGrphFrm!adc("UoGrphDTitle")
      If mGrphFrm!adc("UoGrphDTitle2") <> "" Then .Data2Title = mGrphFrm!adc("UoGrphDTitle2")
      
      If mGrphFrm!adc("UoGrphSumm") <> "" Then _
        mGrphFrm!adc.RepRecSource = SetRepQuery(ps_QryStruct, ps_SqlStr, mGrphFrm!adc.MenuCd, GetVal(mGrphFrm!cboGrphValue.text, DBFldNm), GetVal(mGrphFrm!cboGrphValue.text, TFilter), mGrphFrm!adc("UoGrphSumm"), mGrphFrm!adc("UoGrphG1TopN"), mGrphFrm!adc("UoGrphG1NVal"), mGrphFrm!adc("UoGrphG2TopN"), mGrphFrm!adc("UoGrphG2NVal"))
      
      mGrphFrm!GRP_REP.SetFormula mGrphRep, "wGrphVal", GetVal(mGrphFrm!cboGrphValue.text, RepFormula)
    End With
  End If
  Exit Sub
ErrHndlr:
  DispMsg Err.Description & ". Graph will be Shown Without Further Formmating", etError
  Err.Clear
  Resume Next
End Sub
Public Sub ShowFlds()
  Dim wi_GrpNo As Integer
  wi_GrpNo = SetGrphGrpNo
    If wi_GrpNo = 1 Then
      mGrphFrm!adc("UoGrphG1TopN") = mClsCn.GetFldVal("Select PMCD From Param Where PTyp = 'GRTOPN' And PNum = -1")
      mGrphFrm!adc("UoGrphG1NVal") = 0
      mGrphFrm!LblGrphGrp1.Visible = False
      mGrphFrm!LblGrphG1Srt.Visible = False
      mGrphFrm!LblGrphG1Val.Visible = False
      mGrphFrm!adc("UoGrphG1TopN").Visible = False
      mGrphFrm!adc("UoGrphG1NVal").Visible = False
      mGrphFrm!LblGrphGrp2.Caption = "Group 1"
    ElseIf wi_GrpNo = 2 Then
      mGrphFrm!LblGrphGrp1.Visible = True
      mGrphFrm!LblGrphG1Srt.Visible = True
      mGrphFrm!LblGrphG1Val.Visible = True
      mGrphFrm!adc("UoGrphG1TopN").Visible = True
      mGrphFrm!adc("UoGrphG1NVal").Visible = True
      mGrphFrm!LblGrphGrp2.Caption = "Group 2"
    End If
End Sub
