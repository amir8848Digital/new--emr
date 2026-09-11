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
Case Is = UCase("RepInvLst")  '6.1_38 invoice listing report
    Set wfrm = New EmrFrpInvLst
    
  Case Is = UCase("eInv")
    'Set wfrm = New EmrFrmeInvoice
  Case Is = UCase("CustMaint")
    Set wfrm = New EmrFrmCustMaint
  Case Is = UCase("Head")
    Set wfrm = New EmrFrmHead
  Case Is = UCase("CustMst")
    Set wfrm = New EmrFrmCustMst
  Case Is = UCase("CtlgMst")
    Set wfrm = New EmrFrmCustMst
  Case Is = UCase("Param")
    Set wfrm = New EmrFrmParam
    wfrm.Caption = "Parameters"
  Case Is = UCase("SalOrd")
    Set wfrm = New EmrFrmOrdMst
    
  'indigo.164 new menu for Sales Order Sketch
  Case Is = UCase("SalOrdSk")
    Set wfrm = New EmrFrmOrdMst
  ' Zubin 212P2
  Case Is = UCase("SalOrdAck")
    Set wfrm = New EmrFrmOrdAck
  ' Zubin 212P2
  Case Is = UCase("DM")
    Set wfrm = New EmrFrmDsgMst
  Case Is = UCase("PM")
    Set wfrm = New EmrFrmDsgMst
  Case Is = UCase("SM")
    Set wfrm = New EmrFrmDsgMst
  '*** Jay 3.4[DP]
  Case Is = UCase("DsgPrm")
    Set wfrm = New EmrFrmDsgPrm
  Case Is = UCase("DsgPrmCat")
    Set wfrm = New EmrFrmDsgPrm
  '*** Jay 3.4[DP]
  
  Case Is = UCase("VarParam")
    Set wfrm = New EmrFrmvParam
    wfrm.Caption = "Variable Parameters"
  Case Is = UCase("Bag")
    Set wfrm = New EmrFrmBag
  Case Is = UCase("RC")
    Set wfrm = New EmrFrmBagReco
    wfrm.Caption = "Bag Reconciliation"
  Case Is = UCase("Bnk")
    Set wfrm = New EmrFrmBnkAg
  Case Is = UCase("CpyRt")
    Set wfrm = New EmrFrmCpyRt
  Case Is = UCase("InCustmMtch")
    Set wfrm = New EmrFrmCustmMatch
    wfrm.Caption = "Invoice Custom Matching"
  Case Is = UCase("Tz")
    Set wfrm = New EmrFrmCustmPur
    wfrm.Caption = "Custom Purchase"
  Case Is = UCase("CustMst")
    Set wfrm = New EmrFrmCustMst
  Case Is = UCase("SuppMst")
    Set wfrm = New EmrFrmCustMst
  Case Is = UCase("PLMst")
    Set wfrm = New EmrFrmCustMst
  Case Is = UCase("CoMst")
    Set wfrm = New EmrFrmCustMst
    wfrm.Caption = "CoMst"
  Case Is = UCase("ContractorMst")
    Set wfrm = New EmrFrmCustMst
  Case Is = UCase("DtTbl")
    Set wfrm = New EmrFrmDtTbl
    
  ' ***** Manali 3.9.1 - 21/02/12 - RfId Module
  Case Is = UCase("RfIdMst")
    Set wfrm = New EmrFrmRfIdMst
  
  Case Is = UCase("FB")
    Set wfrm = New EmrFrmFg
    wfrm.Caption = "Finished Goods (Bag)"
  Case Is = UCase("FM")
    Set wfrm = New EmrFrmFgMvm
    wfrm.Caption = "Finished Goods (Bag Movement)"
  Case Is = UCase("FWB")
    Set wfrm = New EmrFrmFg
    wfrm.Caption = "Finished Goods (Bag Return)"
  Case Is = UCase("Head")
    Set wfrm = New EmrFrmHead
  Case Is = UCase("InAllBags")
    Set wfrm = New EmrFrmInvAllBags
    wfrm.Caption = "Invoice All Bags"
  Case Is = UCase("IN")
    Set wfrm = New EmrFrmInvHd
    wfrm.Caption = "Invoice Header"
  Case Is = UCase("LabRt")
    Set wfrm = New EmrFrmLabRt
  Case Is = UCase("RepLst")
'    *** Form with Menu Placed for all Listing Reports
    Set wfrm = New EmrFrmList
  Case Is = UCase("Loc")
    Set wfrm = New EmrFrmLoc
  Case Is = UCase("LossEntry")
    Set wfrm = New EmrFrmLossEntry
  Case Is = UCase("LsRec")
    Set wfrm = New EmrFrmLsRec
    
  Case Is = UCase("Zoom")
    Set wfrm = New EmrFrmOrdZoom
  Case Is = UCase("PpcCap")
    Set wfrm = New EmrFrmPpcCap
  Case Is = UCase("PpcDsgLd")
    Set wfrm = New EmrFrmPpcDsgLd
  Case Is = UCase("PpcOrdPln")
    Set wfrm = New EmrFrmPpcOrdPln
  Case Is = UCase("PrdPts")
    Set wfrm = New EmrFrmPrdPts
  Case Is = UCase("PrdPtsSketch")
    Set wfrm = New EmrFrmPrdPts
  
  Case Is = UCase("PrcLMst")  '3.11.0 '3.11.2 - Name change
    Set wfrm = New EmrFrmPrdPts
  
  Case Is = UCase("ResAuto")
    Set wfrm = New EmrFrmResAutogen
  Case Is = UCase("RmMst")
    Set wfrm = New EmrFrmRmMst
    wfrm.Caption = "Raw Material Master"
  Case Is = UCase("RmRt")
    Set wfrm = New EmrFrmRmRt
  '*** Jay 3.1.2 ***
  Case Is = UCase("ILC")
    Set wfrm = New EmrFrmRmRt
  '*** Jay 3.1.2 ***
  Case Is = UCase("RmRtHist")
    Set wfrm = New EmrFrmRmRt
  ' ****** Manali 3.03 - 04/07/08
  Case Is = UCase("RmIdSz")
    Set wfrm = New EmrFrmRmIdSz
  ' ****** Manali 3.03 - 04/07/08
  
  Case Is = UCase("RepTagPrn")
    Set wfrm = New EmrFrmTag
    wfrm.Caption = "Invoice Tag Printing"
' Zubin 212
  Case Is = UCase("RepTagPrnBagFg")
    Set wfrm = New EmrFrmFgTag
    wfrm.Caption = "Bag FG Tag Printing"
  Case Is = UCase("RepTagPrnFg")
    Set wfrm = New EmrFrmFgTag
    wfrm.Caption = "FG Tag Printing"
' Zubin 212

  Case Is = UCase("TagFmt")
    If GetMDIFrm.ADC.Connection.RecSeek("select 1 from dbo.sysobjects where id = object_id(N'[TagFmt]') and OBJECTPROPERTY(id, N'IsUserTable') = 1") _
      And _
      GetMDIFrm.ADC.Connection.RecSeek("select 1 from dbo.sysobjects where id = object_id(N'[TagFmtSr]') and OBJECTPROPERTY(id, N'IsUserTable') = 1") Then
        '  If gs_Ver = "2.6.0" Or gs_Ver = "2.7.0" Or gs_Ver = "2.8.0" Or gs_Ver = "2.9.0" Then
      Set wfrm = New EmrFrmTagFmt
    End If
  Case Is = UCase("Tree")
    Set wfrm = New EmrFrmTree
    ' Zubin 213
'  Case Is = UCase("XMLToEmr")
'    Set wfrm = New EmrFrmTrfIn
'    wfrm.Caption = "XML To Emr (Order)"
  Case Is = UCase("EmrToXML")
    Set wfrm = New EmrFrmTrfOut
' Zubin 213
'  Case Is = UCase("XMLMstToEmr")
'    Set wfrm = New xEmrFrmTrfIn
'    wfrm.Caption = "XML To Emr (Master)"
  Case Is = UCase("SA")
    Set wfrm = New EmrFrmTxnRm
    wfrm.Caption = "Sales"
  Case Is = UCase("CNV")
    Set wfrm = New EmrFrmTxnRm
    wfrm.Caption = "Conversion"
  Case Is = UCase("PR")
    Set wfrm = New EmrFrmTxnRm
    wfrm.Caption = "Purchase"
  Case Is = UCase("MV")
    Set wfrm = New EmrFrmTxnRm
    wfrm.Caption = "Raw Material Issue"
 Case Is = UCase("MA")
    Set wfrm = New EmrFrmTxnRm
    wfrm.Caption = "Raw Material Acknowledgement"
  Case Is = UCase("UsrAcss")
    Set wfrm = New EmrFrmUsrAcss
    wfrm.Caption = "User Access Master"
  Case Is = UCase("UsrAdm")
    Set wfrm = New EmrFrmUsrAdm
    wfrm.Caption = "User Administration"

  '*** (Jen 3.01)
  Case Is = UCase("IdeDfn")
    Set wfrm = New EmrFrmIDE
    wfrm.Caption = "IDE Definition"
  Case Is = UCase("IdeUsrRight")
    Set wfrm = New EmrFrmIDE
    wfrm.Caption = "IDE User Rights"
  '*** (Jen 3.01)

  Case Is = UCase("UsrMail")
    Set wfrm = New EmrFrmUsrMail
    wfrm.Caption = "User Mail"
  Case Is = UCase("VarParam")
    Set wfrm = New EmrFrmvParam
    wfrm.Caption = "Variable Parameters"
' Zubin 213
'  Case Is = UCase("XMLMstToEmr")
'    Set wfrm = New xEmrFrmTrfIn
'    wfrm.Caption = "XML To Emr (Master)"
  Case Is = UCase("EmrUpd")
    Set wfrm = New EmrSuppFrmUpd
  Case Is = UCase("RepDscRep")          '*** Jay 2.14 P2 *** [Dsc Rep Introduced]
    Set wfrm = New EmrSuppFrpDscRep
  Case Is = UCase("RepSOPnd")
    Set wfrm = New EmrFrpSOPnd
    wfrm.Caption = "Sales Order Pending"
    
     Case Is = UCase("RepSOPndRM")
    Set wfrm = New EmrFrpSOPnd
    wfrm.Caption = "Sales Order Pending with RM Detail"
    
  Case Is = UCase("RepBagHist")
    Set wfrm = New EmrFrpBagHist
    wfrm.Caption = "Bag History"
  Case Is = UCase("RmZMst")
    Set wfrm = New EmrFrmRmMst
    wfrm.Caption = "Customs Raw Material Master"
  Case Is = UCase("PrcLst")
    Set wfrm = New EmrFrmOrdMst
  '6.1-15 new menu for pdcm prominent menu
  Case Is = UCase("PDCMUPD")
    Set wfrm = New EmrFrmOrdMst
  Case Is = UCase("Quot")
   Set wfrm = New EmrFrmOrdMst
  Case Is = UCase("QB")
    Set wfrm = New EmrFrmOrdMst
  Case Is = UCase("SkQuot")
   Set wfrm = New EmrFrmOrdMst

  ' ***** Manali 3.10.0 - Multi Price Qt Module
  Case Is = UCase("MultiPrcQt")
    Set wfrm = New EmrFrmMultiPrcQt
    
  Case Is = UCase("RepMultiPrcQt")
    Set wfrm = New EmrFrpMultiPrcQt
    
  ' ***** Manali 3.10.0 - Multi Price Qt Module
  
  '*** (Jen 2.13)
  Case Is = UCase("RetMem")
    Set wfrm = New EmrFrmRetMemo
  '*** (Jen 2.13)
  
 Case Is = UCase("RepCustmStkLed")      '4.1.2.0
    Set wfrm = New EmrFrpCustmStkLed

' Geeta Emr 212
  Case Is = UCase("MF")
 '   Set wfrm = New EmrFrmOrdMst
' Geeta Emr 212

  Case Is = UCase("BM")
    Set wfrm = New EmrFrmTxn
    wfrm.Caption = "Bag Merge"
  Case Is = UCase("DT")
    Set wfrm = New EmrFrmTxn
    wfrm.Caption = "Daily Transactions"
  Case Is = UCase("BV")
    Set wfrm = New EmrFrmTxn
    wfrm.Caption = "Bag Movement"
  Case Is = UCase("BA")
    Set wfrm = New EmrFrmTxn
    wfrm.Caption = "Bag Acknowledgement"
  Case Is = UCase("BS")
    Set wfrm = New EmrFrmTxn
    wfrm.Caption = "Bag Split"
  Case Is = UCase("REJ")
    Set wfrm = New EmrFrmTxn
    wfrm.Caption = "Rejection"
  Case Is = UCase("MLT")
    Set wfrm = New EmrFrmTxn
    wfrm.Caption = "Melting"
  Case Is = UCase("CTBIn")
    Set wfrm = New EmrFrmInvHd
    wfrm.Caption = "CTB Header"
  Case Is = UCase("CTBInAllBags")
    Set wfrm = New EmrFrmInvAllBags
    wfrm.Caption = "CTB All Bags"
  Case Is = UCase("CTBInDsg")
    Set wfrm = New EmrFrmInvDsg
    wfrm.Caption = "CTB Design"
  Case Is = UCase("RepBagPrn")
    Set wfrm = New EmrFrpBagPrn
    wfrm.Caption = "Bag Printing"
  Case Is = UCase("RepBagReco")
    Set wfrm = New EmrFrpBagReco
    wfrm.Caption = "Bag Reconciliation Report"
  Case Is = UCase("RepBrkMsgRm")
    Set wfrm = New EmrFrpBrkMsg
    wfrm.Caption = "Breakage, Missing Report (Base= Rm)"
  Case Is = UCase("RepBrkMsgPrd")
    Set wfrm = New EmrFrpBrkMsg
    wfrm.Caption = "Breakage, Missing Report (Base= Prd)"
  Case Is = UCase("RepBv")
    Set wfrm = New EmrFrpBv
    wfrm.Caption = "Bag Movement (Bag Pieces)"
  
  '4.1.0.0
  Case Is = UCase("RepBrkMsgFg")
    Set wfrm = New EmrFrpBrkMsgFg
    wfrm.Caption = "Breakage Missing (FG Level)"
  
  Case Is = UCase("RepBvRm")
    Set wfrm = New EmrFrpBvRm
    wfrm.Caption = "Bag Movement (Raw Material)"
  Case Is = UCase("RepCustmStk")
    Set wfrm = New EmrFrpCustmStk
    wfrm.Caption = "Custom Stock"
  Case Is = UCase("RepDiaProfit")
    Set wfrm = New EmrFrpDiaProfit
    wfrm.Caption = "Diamond Profit"
  Case Is = UCase("InDsg")
    Set wfrm = New EmrFrmInvDsg
    wfrm.Caption = "Invoice Design"
  Case Is = UCase("RepFlashStk")
    Set wfrm = New EmrFrpFlashStk
    wfrm.Caption = "Flash Stock"
  Case Is = UCase("RepFltBagPrn")
    Set wfrm = New EmrFrpFltBagPrn
    wfrm.Caption = "Flute Bag Printing"
' ***** Manali 3.6.0 - 01/09/09 - Gold Loss Report Name and Menu Code Changed To Metal Loss For Users
'  Case Is = UCase("RepMetLs")
'    Set wfrm = New EmrFrpGldLs
'    wfrm.Caption = "Metal Loss"
  
  ' ***** Manali 3.8.0 - Metal Loss Report Name and Menu Code Changed To Metal Loss (WIP)
  Case Is = UCase("RepMetLsWip")
    Set wfrm = New EmrFrpGldLs
    wfrm.Caption = "Metal Loss (WIP)"
    
    
  Case Is = UCase("RepMetLsFg")
    Set wfrm = New EmrFrpGldLs
    wfrm.Caption = "Metal Loss (Fg)"
    
''' Bef 3.6.0
'''  Case Is = UCase("RepGldLs")
'''    Set wfrm = New EmrFrpGldLs
'''    wfrm.Caption = "Gold Loss"
' ***** Manali 3.6.0 - 01/09/09 - Gold Loss Report Name and Menu Code Changed To Metal Loss For Users
  Case Is = UCase("RepInv")
    Set wfrm = New EmrFrpInv
    wfrm.Caption = "Invoice Printing"
  Case Is = UCase("RepGrFormFrnt")
    Set wfrm = New EmrFrpInv
    wfrm.Caption = "GR Form (Front)"
    
  '4.1.0.0
  Case Is = UCase("RepEDFFormFrnt")
    Set wfrm = New EmrFrpEDF
    wfrm.Caption = "EDF Form Front"
    
  Case Is = UCase("RepEDFFormBck")
    Set wfrm = New EmrFrpEDF
    wfrm.Caption = "EDF Form Back"
    
  '*** Added Again In 3.2.0
    '*** JAY 2.14 P2 *** [DT:12/09/07]
    Case Is = UCase("RepGrFormFrntNew")
      Set wfrm = New EmrFrpInv
      wfrm.Caption = "GR Form (Front) - Preprinted"
    '*** JAY 2.14 P2 *** [DT:12/09/07]
    '*** Jay 3.1.1
    Case Is = UCase("RepGrFormFrntNewMul")
      Set wfrm = New EmrFrpInv
      wfrm.Caption = "GR Form(Front)-Preprinted (Multi Inv)"
    '*** Jay 3.1.1
    '*** Jay 3.1.1 (All Rates Rpt)
    Case Is = UCase("RepCustmStkLand")
      Set wfrm = New EmrFrpCustmStk
      wfrm.Caption = "Custom Stock (Landscape)"
    '*** Jay 3.1.1 (All Rates Rpt)
  '*** Added Again In 3.2.0
  Case Is = UCase("RepShpBill")
    Set wfrm = New EmrFrpInv
    wfrm.Caption = "Shipping Bill"
  
  ' 3.9.0
  Case Is = UCase("RepShpBillFile")
    Set wfrm = New EmrFrpShpBillFile
    wfrm.Caption = "Shipping Bill File"
  
  '*** Jay 2.14(LclInv) ***
  Case Is = UCase("RepLclInv")
    Set wfrm = New EmrFrpLclInv
    wfrm.Caption = "Local Invoice Printing"
  
  Case Is = UCase("RepInvLclRM")
    Set wfrm = New EmrFrpLclInv
    wfrm.Caption = "Local Invoice RM Printing"
  
  
  '*** Jay 2.14(LclInv) ***

  '*** (Jen 2.13)
  Case Is = UCase("RepGrFormFrntMulti")
    Set wfrm = New EmrFrpInv
    wfrm.Caption = "GR Form (Front) (Multi Invoice)"
  Case Is = UCase("RepShpBillMulti")
    Set wfrm = New EmrFrpInv
    wfrm.Caption = "Shipping Bill (Multi Invoice)"
  '*** (Jen 2.13)

  Case Is = UCase("RepGSPFORM")
    Set wfrm = New EmrFrpInv
    wfrm.Caption = "GSP Form"
  Case Is = UCase("RepInvBOELst")
    Set wfrm = New EmrFrpInvBoeLst
    wfrm.Caption = "Invoice BOE Listing"
  Case Is = UCase("RepDetInv")
    Set wfrm = New EmrFrpInvCost
    wfrm.Caption = "Invoice Costing"
  Case Is = UCase("RepAvgInv")
    Set wfrm = New EmrFrpInvCost
    wfrm.Caption = "Average Invoice Printing"
  Case Is = UCase("RepInvCat")
    Set wfrm = New EmrFrpInvCost
    wfrm.Caption = "Invoice Catalogue"
  Case Is = UCase("RepDetInvCTB")
    Set wfrm = New EmrFrpInvCost
    wfrm.Caption = "CTB Invoice Costing"
  Case Is = UCase("RepAvgInvCTB")
    Set wfrm = New EmrFrpInvCost
    wfrm.Caption = "Average CTB Invoice Printing"
  Case Is = UCase("RepInvCTBCat")
    Set wfrm = New EmrFrpInvCost
    wfrm.Caption = "CTB Invoice Catalogue"
' Zubin 213 (Before 2.13 this was missing)
  Case Is = UCase("RepCTBPckLst")
    Set wfrm = New EmrFrpPckLst
    wfrm.Caption = "Finished Goods Packing List For CTB"
' Zubin 213 (Before 2.13 this was missing)
  Case Is = UCase("RepInvCustStmt")
    Set wfrm = New EmrFrpInvCustStmt
  Case Is = UCase("RepGrFormBack")
    Set wfrm = New EmrFrpInvLet
    wfrm.Caption = "GR Form (Back)"

  '*** (Jen 2.14)
  Case Is = UCase("RepGrFormBackMulti")
    Set wfrm = New EmrFrpInvLet
    wfrm.Caption = "GR Form (Back) (Multi Invoice)"
  '*** (Jen 2.14)

  Case Is = UCase("RepGSPDetLet")
    Set wfrm = New EmrFrpInvLet
    wfrm.Caption = "GSP Detail Letter"
  Case Is = UCase("RepGSPCert")
    Set wfrm = New EmrFrpInvLet
    wfrm.Caption = "GSP Certification"
  Case Is = UCase("RepDeclnLet")
    Set wfrm = New EmrFrpInvLet
    wfrm.Caption = "Declaration Letter"

  '*** (Jen 3.01)
  Case Is = UCase("RepGSPFORMMulti")
    Set wfrm = New EmrFrpInv
    wfrm.Caption = "GSP Form (Multi Invoice)"
  Case Is = UCase("RepGSPDetLetMulti")
    Set wfrm = New EmrFrpInvLet
    wfrm.Caption = "GSP Detail Letter (Multi Invoice)"
  Case Is = UCase("RepGSPCertMulti")
    Set wfrm = New EmrFrpInvLet
    wfrm.Caption = "GSP Certification (Multi Invoice)"
  Case Is = UCase("RepDeclnLetMulti")
    Set wfrm = New EmrFrpInvLet
    wfrm.Caption = "Declaration Letter (Multi Invoice)"
  '*** (Jen 3.01)
  Case Is = UCase("RepRWrk")
    Set wfrm = New EmrFrpRWrk
    wfrm.Caption = "Rework Report"
  Case Is = UCase("RepInvLossSubRep")
    Set wfrm = New EmrFrpInvLossSubRep
    wfrm.Caption = "Invoice Metal Loss Sub Report"
  Case Is = UCase("RepInvProfitMrg")
    Set wfrm = New EmrFrpInvProfitMrg
    wfrm.Caption = "Profit Margin"
  Case Is = UCase("RepLsRec")
    Set wfrm = New EmrFrpLsRec
  Case Is = UCase("RepNetStk")
    Set wfrm = New EmrFrpNetStk
   Case Is = UCase("RepDetOrd")
    Set wfrm = New EmrFrpOrd
    wfrm.Caption = "Detail Order Printing"
    Case Is = UCase("RepAvgOrd")
    Set wfrm = New EmrFrpOrd
    wfrm.Caption = "Average Order Printing"
    Case Is = UCase("RepOrdCat")
    Set wfrm = New EmrFrpOrd
    wfrm.Caption = "Order Catalogue Printing"
    Case Is = UCase("RepOrdProfitMrg")
    Set wfrm = New EmrFrpOrdProfitMrg
    wfrm.Caption = "Order Profit Margin"
  ' Zubin (3.02) - 250308
  '****** Sachin 3.3.0 - Enabled Travel Excel in 3.3.0 ******
  Case Is = UCase("RepTravExcel")
    Set wfrm = New EmrFrpTravExcel
    wfrm.Caption = "Travel Excel Report"
  '****** Sachin 3.3.0 - Enabled Travel Excel in 3.3.0 ******
  Case Is = UCase("RepPckLst")
    Set wfrm = New EmrFrpPckLst
    wfrm.Caption = "Finished Goods Packing List Fr INV"
  Case Is = UCase("RepPckLstInv")
    Set wfrm = New EmrFrpPckLstInv
    wfrm.Caption = "Invoice Packing List"
  Case Is = UCase("RepCntrlCpy")
    Set wfrm = New EmrFrpPpcCntrlCpy
  Case Is = UCase("RepOrdPln")
    Set wfrm = New EmrFrpPpcOrdPln
  Case Is = UCase("RepDsgLd")
    Set wfrm = New EmrFrpPpcDsgLd
  Case Is = UCase("RepPrd")
    Set wfrm = New EmrFrpPrd
    wfrm.Caption = "Production Report"
  Case Is = UCase("RepPrdPts")
    Set wfrm = New EmrFrpPrdPts
    wfrm.Caption = "Production Points Report"
  Case Is = UCase("RepRejAna")
    Set wfrm = New EmrFrpRejAna
    wfrm.Caption = "Rejection Analysis"
  Case Is = UCase("RepRmMst")
    Set wfrm = New EmrFrpRmMst
    wfrm.Caption = "Raw Material/ Custom Master Listing"
  Case Is = UCase("RepRmReqStat")
    Set wfrm = New EmrFrpRmReq
    wfrm.Caption = "Raw Material Requirement (Static)"
  Case Is = UCase("RepRmReqDyn")
    Set wfrm = New EmrFrpRmReq
    wfrm.Caption = "Raw Material Requirement (Dynamic)"
  Case Is = UCase("RepRmReqLoc")
    Set wfrm = New EmrFrpRmReqLoc
    wfrm.Caption = "Raw Material Requirement For Loc"
  Case Is = UCase("RepRmReqSet")
    Set wfrm = New EmrFrpRmReqLoc
    wfrm.Caption = "Raw Material Requirement (Setting)"
  Case Is = UCase("RepRmRt")
    Set wfrm = New EmrFrpRmRt
    wfrm.Caption = "Raw Material Rate Master Listing"
  Case Is = UCase("RepLabRt")
    Set wfrm = New EmrFrpLabRt
    wfrm.Caption = "Labour Rate Master Listing"
  Case Is = UCase("RepRmRtHist")
    Set wfrm = New EmrFrpRmRt
    wfrm.Caption = "Raw Material Rate History Listing"
  Case Is = UCase("RepLoc")
    Set wfrm = New EmrFrpLoc
  Case Is = UCase("RepSalAnaDsg")
    Set wfrm = New EmrFrpSalAnaDsg
    wfrm.Caption = "Sales Analysis (Design Level)"
  Case Is = UCase("RepSalAnaRm")
    Set wfrm = New EmrFrpSalAnaRm
    wfrm.Caption = "Sales Analysis (Rm Level)"
  Case Is = UCase("RepSalAnaLab")
    Set wfrm = New EmrFrpSalAnaLab
    wfrm.Caption = "Sales Analysis (Lab Level)"
  Case Is = UCase("RepStkLed")
    Set wfrm = New EmrFrpStkLed
    wfrm.Caption = "Stock Ledger"
  Case Is = UCase("RepStkSumm")
    Set wfrm = New EmrFrpStkSumm
    wfrm.Caption = "Stock Summary"
  Case Is = UCase("RepAnnexValAddn")
    Set wfrm = New EmrFrpValAddnAnnex
    wfrm.Caption = "Annexure To Value Addition"
  Case Is = UCase("RepValAddn")
    Set wfrm = New EmrFrpValAddn
    wfrm.Caption = "Value Addition"
  ' ***** Manali 3.8.0
  Case Is = UCase("RepValAddnDta")
    Set wfrm = New EmrFrpValAddn
    wfrm.Caption = "Value Addition (DTA)"
  ' ***** Manali 3.8.0
  Case Is = UCase("RepWipBPcs")
    Set wfrm = New EmrFrpWipBPcs
    wfrm.Caption = "WIP Bag Pieces Report"
  Case Is = UCase("RepWipFgBPcs")
    Set wfrm = New EmrFrpWipBPcs
    wfrm.Caption = "Fg Bag Pieces Report"
  Case Is = UCase("RepWipBPcsBkDt")
    Set wfrm = New EmrFrpWipBPcsBkDt
    wfrm.Caption = "WIP Bag Pieces Report (Back Dated)"
  Case Is = UCase("RepWipFgBPcsBkDt")
    Set wfrm = New EmrFrpWipBPcsBkDt
    wfrm.Caption = "Fg Bag Pieces Report (Back Dated)"
  Case Is = UCase("RepWipBRm")
    Set wfrm = New EmrFrpWipBRm
    wfrm.Caption = "WIP Bag Raw Material Report"

  '*** (Jen 2.14)
  Case Is = UCase("RepWipBRmBkDt")
    Set wfrm = New EmrFrpWipBRm
    wfrm.Caption = "WIP Bag Raw Material Report (Back Dated)"
  '*** (Jen 2.14)

  Case Is = UCase("RepWipFgBRm")
    Set wfrm = New EmrFrpWipBRm
    wfrm.Caption = "Fg Bag Raw Material Report"
  Case Is = UCase("RepWIWPrd")
    Set wfrm = New EmrFrpWIW
    wfrm.Caption = "What Is Where (Production)"
  Case Is = UCase("RepWIWExp")
    Set wfrm = New EmrFrpWIW
    wfrm.Caption = "What Is Where (Export)"
  Case Is = UCase("RepWIWLongPrd")
    Set wfrm = New EmrFrpWIW
    wfrm.Caption = "What Is Where Long (Production)"
  Case Is = UCase("RepWIWLongExp")
    Set wfrm = New EmrFrpWIW
    wfrm.Caption = "What Is Where Long (Export)"

  Case Is = UCase("ClrDfn")
    Set wfrm = New EmrFrmClrDfn
  Case Is = UCase("ClrSel")
    Set wfrm = New EmrFrmClrPic

  Case Is = UCase("Grd")
    Set wfrm = New EmrFrmGrade
'' **** End Of MDI Menu Forms **** '
'
'
'
'' **** Report Listing (LST) **** '
  Case Is = UCase("RepPrm")
    Set wfrm = New EmrFrpParam
    wfrm.Caption = "Parameter Listing"
  Case Is = UCase("RepTxnBM")
    Set wfrm = New EmrFrpTxnLst
    wfrm.Caption = "Bag Merge Transaction Listing"
  Case Is = UCase("RepVarPrm")
    Set wfrm = New EmrFrpParam
    wfrm.Caption = "Variable Parameter Listing"
  Case Is = UCase("RepCustLst")
    Set wfrm = New EmrFrpCustMst
    wfrm.Caption = "Customer Listing"
  Case Is = UCase("RepUsrAccess")
    Set wfrm = New EmrFrpUsrAccess
    wfrm.Caption = "User Access Listing"
  Case Is = UCase("RepCustDsg")
    Set wfrm = New EmrFrpCustDsg
    wfrm.Caption = "Customer Design Listing"
  Case Is = UCase("RepCustAly")
    Set wfrm = New EmrFrpCustAly
    wfrm.Caption = "Customer Alloy Listing"
  Case Is = UCase("RepRmMst")
    Set wfrm = New EmrFrpRmMst
    wfrm.Caption = "Raw Material/ Custom Master Listing"
  Case Is = UCase("RepDsg")
    Set wfrm = New EmrFrpDsgLst
    wfrm.Caption = "Design Master Listing"
  Case Is = UCase("RepDsgCrd")
    Set wfrm = New EmrFrpDsgCrd
    wfrm.Caption = "Design Card Details"
  Case Is = UCase("RepDsgCat")
    Set wfrm = New EmrFrpDsgCat
    wfrm.Caption = "Design Catalogue"

  '*** (Jen 3.01)
  Case Is = UCase("DsgView")
    Set wfrm = New EmrFrmDsgView
  '*** (Jen 3.01)

  Case Is = UCase("RepRmRt")
    Set wfrm = New EmrFrpRmRt
    wfrm.Caption = "Raw Material Rate Master Listing"
  Case Is = UCase("RepLabRt")
    Set wfrm = New EmrFrpLabRt
    wfrm.Caption = "Labour Rate Master Listing"
  '*** Jay 3.01 ***
  Case Is = UCase("RepRtCmp")
    Set wfrm = New EmrFrpRtCmp
    wfrm.Caption = "Row Material Rate Chart Comparision"
  Case Is = UCase("RepTxnCustmPurLst")
    Set wfrm = New EmrFrpCustmPurLst
    wfrm.Caption = "Custom Purchase Listing"
  Case Is = UCase("RepVchPrnCustmPur")
    Set wfrm = New EmrFrpCustmPurVchPrn
    wfrm.Caption = "Custom Purchase Voucher"
  '*** Jay 3.01 ***
  Case Is = UCase("RepLoc")
    Set wfrm = New EmrFrpLoc
    wfrm.Caption = "Location Listing"
  Case Is = UCase("RepDtTbl")
    Set wfrm = New EmrFrpDtTbl
    wfrm.Caption = "Date Table Listing"
  Case Is = UCase("RepBnk")
    Set wfrm = New EmrFrpBnkAg
    wfrm.Caption = "Bank/Agent Listing"
  Case Is = UCase("RepBagLst")
    Set wfrm = New EmrFrpBagLst
    wfrm.Caption = "Bag Listing"
  Case Is = UCase("RepPrdPts")
    Set wfrm = New EmrFrpPrdPts
    wfrm.Caption = "Production Points Report"
  Case Is = UCase("RepLossMst")
    Set wfrm = New EmrFrpLossMstLst
    wfrm.Caption = "Loss Master Listing"
  Case Is = UCase("RepGrade")
    Set wfrm = New EmrFrpGrade
    wfrm.Caption = "Grade Listing"
  Case Is = UCase("RepTxnDT")
    Set wfrm = New EmrFrpTxnLst
    wfrm.Caption = "Detail Transaction Listing"
  Case Is = UCase("RepTxnBS")
    Set wfrm = New EmrFrpTxnLst
    wfrm.Caption = "Bag Split Transaction Listing"
  Case Is = UCase("RepTxnMLT")
    Set wfrm = New EmrFrpTxnLst
    wfrm.Caption = "Melting Transaction Listing"
  Case Is = UCase("RepTxnBV")
    Set wfrm = New EmrFrpTxnLst
    wfrm.Caption = "Bag Movement Transaction Listing"
  Case Is = UCase("RepTxnREJ")
    Set wfrm = New EmrFrpTxnLst
    wfrm.Caption = "Rejection Transaction Listing"
  Case Is = UCase("RepTxnPR")
    Set wfrm = New EmrFrpTxnLst
    wfrm.Caption = "Purchase/Sales Transaction Listing"
  Case Is = UCase("RepTxnMV")
    Set wfrm = New EmrFrpTxnLst
    wfrm.Caption = "RM Movement Transaction Listing"
  Case Is = UCase("RepTxnCNV")
    Set wfrm = New EmrFrpTxnLst
    wfrm.Caption = "Conversion Transaction Listing"
  Case Is = UCase("RepUnMtchCnv")
    Set wfrm = New EmrFrpUnMtchCnv
    wfrm.Caption = "Unmatched Conversion Listing"
'  Case Is = UCase("RepTxnREC")
'    Set wfrm = New EmrFrpTxnLst
'    wfrm.Caption = "Recovery Transaction Listing"
  Case Is = UCase("RepFbLst")
    Set wfrm = New EmrFrpFgList
    wfrm.Caption = "Finished Goods (Bag)"
  Case Is = UCase("RepFwbLst")
    Set wfrm = New EmrFrpFgList
    wfrm.Caption = "Finished Goods (Bag Return)"
  Case Is = UCase("RepFgMvm")
    Set wfrm = New EmrFrpFgMvm
    wfrm.Caption = "Finished Goods (Bag Movement)"
  Case Is = UCase("RepFgBagLed")
    Set wfrm = New EmrFrpFgBagLed
    wfrm.Caption = "Finished Goods Bag Ledger"
  Case Is = UCase("RepVchPrnBM")
    Set wfrm = New EmrFrpVchPrn
    wfrm.Caption = "Bag Merge Voucher Printing"
  Case Is = UCase("RepVchPrnDT")
    Set wfrm = New EmrFrpVchPrn
    wfrm.Caption = "Daily Voucher Printing"
  Case Is = UCase("RepVchPrnBS")
    Set wfrm = New EmrFrpVchPrn
    wfrm.Caption = "Bag Split Voucher Printing"
  Case Is = UCase("RepVchPrnMLT")
    Set wfrm = New EmrFrpVchPrn
    wfrm.Caption = "Melting Voucher Printing"
  Case Is = UCase("RepVchPrnBV")
    Set wfrm = New EmrFrpVchPrn
    wfrm.Caption = "Bag Movement Voucher Printing"
  Case Is = UCase("RepVchPrnREJ")
    Set wfrm = New EmrFrpVchPrn
    wfrm.Caption = "Rejection Voucher Printing"
  Case Is = UCase("RepVchPrnPR")
    Set wfrm = New EmrFrpVchPrn
    wfrm.Caption = "Purchase/Sales Voucher Printing"
  Case Is = UCase("RepVchPrnMV")
    Set wfrm = New EmrFrpVchPrn
    wfrm.Caption = "Rm Movement Voucher Printing"
  Case Is = UCase("RepVchPrnCNV")
    Set wfrm = New EmrFrpVchPrn
    wfrm.Caption = "Conversion Voucher Printing"
  Case Is = UCase("RepVchPrnREC")
    Set wfrm = New EmrFrpVchPrn
    wfrm.Caption = "Recovery Voucher Printing"
  Case Is = UCase("RepLsRec")
    Set wfrm = New EmrRepLsRec
  Case Is = UCase("RepCustRm")
    Set wfrm = New EmrFrpCustRm
    wfrm.Caption = "Customer RM Listing"
  Case Is = UCase("RepTreePerform")
    Set wfrm = New EmrFrpTreeLst
    wfrm.Caption = "Tree Performance Report"
  Case Is = UCase("RepTreeInvst")
    Set wfrm = New EmrFrpTreeLst
    wfrm.Caption = "Tree Investment Report"

  '*** (Jen 2.12)
  Case Is = UCase("RepInvLoose")
    Set wfrm = New EmrFrpInv
    wfrm.Caption = "Invoice Printing (Loose RM)"
  Case Is = UCase("RepShpBillLoose")
    Set wfrm = New EmrFrpInv
    wfrm.Caption = "Shipping Bill (Loose RM)"
  Case Is = UCase("RepGrFormLoose")
    Set wfrm = New EmrFrpInv
    wfrm.Caption = "GR Form (Loose RM)"
  Case Is = UCase("RepPckLstLoose")
    Set wfrm = New EmrFrpPckLstLooseExp
    wfrm.Caption = "Packing List (Loose RM)"
  '*** (Jen 2.12)
    
  ' ***** Manali 3.8.0
  Case Is = UCase("RepPckLstInvBag")
    Set wfrm = New EmrFrpPckLstInvBag   'EmrFrpInv   'EmrFrpPckLstInvBag
    wfrm.Caption = "Packing List (Invoice Bag)"
  ' ***** Manali 3.8.0

  '*** (Jen 2.14 Next)
  Case Is = UCase("RepTxnRetMem")
    Set wfrm = New EmrFrpRetMemTxnPrn
  Case Is = UCase("RepVchPrnRetMem")
    Set wfrm = New EmrFrpRetMemVchPrn
  '*** (Jen 2.14 Next)


  '****** Sachin 3.3.0 (28-06-08) - [27. Valid Rm Sz Master] ******
  Case Is = UCase("VldRmLn")
    Set wfrm = New EmrFrmVldRmLn
  '****** Sachin 3.3.0 (28-06-08) - [27. Valid Rm Sz Master] ******
  
  '***** Manali 3.03 - 11/07/08 - RepRmIdSz
  Case Is = UCase("RepRmIdSz")
    Set wfrm = New EmrFrpRmIdSz
  '***** Manali 3.03 - 11/07/08 - RepRmIdSz
  
  ' ****** Manali 3.9.1 - 23/02/12 - RfId Module
  Case Is = UCase("RepRfIdMst")
    Set wfrm = New EmrFrpRfIdMst
  ' ****** Manali 3.9.1 - 23/02/12 - RfId Module
  
  '*** Jay 3.3(FG) ***
  Case Is = UCase("Fin")
 '   Set wfrm = New EmrFrmOrdMst
  Case Is = UCase("Fot")
 '   Set wfrm = New EmrFrmOrdMst
  Case Is = UCase("InvHdFot")
    Set wfrm = New EmrFrmInvHd
    wfrm.Caption = "Invoice Header (FG)"
  Case Is = UCase("RepFgAvg")
    Set wfrm = New EmrFrpOrd
    wfrm.Caption = "Average Finish Goods Printing"
  Case Is = UCase("RepFgCat")
    Set wfrm = New EmrFrpOrd
    wfrm.Caption = "Finish Goods Catalogue"
  Case Is = UCase("RepFGDet")
    Set wfrm = New EmrFrpOrd
    wfrm.Caption = "Detail Finish Goods Printing"
  Case Is = UCase("RepFgInv")
    Set wfrm = New EmrFrpInv
    wfrm.Caption = "FG Invoice Printing"
  Case Is = UCase("RepFgGrFrnt")
    Set wfrm = New EmrFrpInv
    wfrm.Caption = "GR Form (Front)"
  Case Is = UCase("RepFgGrBk")
     Set wfrm = New EmrFrpInvLet
    wfrm.Caption = "GR Form (Back)"
  Case Is = UCase("RepFgGsp")
    Set wfrm = New EmrFrpInv
    wfrm.Caption = "GSP Form"
  Case Is = UCase("RepFgShpBill")
    Set wfrm = New EmrFrpInv
    wfrm.Caption = "Shipping Bill"
  Case Is = UCase("RepFgStkLed")
    Set wfrm = New EmrFrpFgStkLed
    wfrm.Caption = "FG Stock Ledger"
  Case Is = UCase("RepFgPckLst")
    Set wfrm = New EmrFrpPckLstInv
    wfrm.Caption = "Packing List (Fg Invoice)"
  Case Is = UCase("RepFgLclInv")
    Set wfrm = New EmrFrpLclInv
    wfrm.Caption = "Local Invoice Printing (FG)"
   '*** Jay 3.3(FG)
   
  '3.11.1.0
  Case Is = UCase("RepEPSPlainJwl")
    Set wfrm = New EmrFrpInvEPS
    wfrm.Caption = "EPS & Value Addition for Plain Jewellery"
   
  '3.11.1.0
  Case Is = UCase("RepEPSStudJwl")
    Set wfrm = New EmrFrpInvEPS
    wfrm.Caption = "Monthly EPS & Value Addition for Studded Jewellery"
   
  '3.11.1.0
  Case Is = UCase("RepEPSMetImpEx")
    Set wfrm = New EmrFrpInvEPS
    wfrm.Caption = "Monthly Stmt of Metal Procurement, Consumption & Export"
   
   '3.11.1.0
  Case Is = UCase("RepStudStmt")
    Set wfrm = New EmrFrpValAddn
    wfrm.Caption = "Statement of Studding Particulars"
   
 '4.1.2.0
  Case Is = UCase("RepFGCatalog")
    Set wfrm = New EmrFrpFgCat
    wfrm.Caption = "FG Catalog"

 'uma *** Repair Transaction
  Case Is = UCase("RDT")
    Set wfrm = New EmrFrmTxn
    wfrm.Caption = "Repair Bag Transaction"
    
 'uma *** Repair Movement
  Case Is = UCase("RBV")
    Set wfrm = New EmrFrmTxn
    wfrm.Caption = "Repair Bag Movement"
  
  Case Is = UCase("RprHis")
    Set wfrm = New EmrFrpBagHist
    wfrm.Caption = "Repair Bag History"
  Case Is = UCase("RprWIP")
    Set wfrm = New EmrFrpWipBPcs
    wfrm.Caption = "WIP Repair Bag Pieces Report"
  Case Is = UCase("RprWIPBv")
    Set wfrm = New EmrFrpBv
    wfrm.Caption = "Repair Bag Movement (Bag Pieces)"
  Case Is = UCase("REPLOTZOOM")
    Set wfrm = New EmrFrmLotZoom
 'VK.5-added menu for Certificate Entry
  Case Is = UCase("CertEntry")
    Set wfrm = New EmrFrmCertEntry
  'VK.5-added menu for Certificate Report
  Case Is = UCase("RepCertLst")
    Set wfrm = New EmrFrpCertLst
    wfrm.Caption = "Certificate Listing"
  Case Is = UCase("RepShpAna")
    Set wfrm = New EmrFrpShpAna
    wfrm.Caption = "Shipment Analysis"
  Case Is = UCase("SQT")  'SV.33 Supplier quotation added
    Set wfrm = New EmrFrmOrdMst
  '6.1.11 added menu for shipment note
  Case Is = UCase("ShpNote")
    Set wfrm = New EmrFrmShpNote
    
'''  ' **** Manali Trading Module
'''  Case Is = UCase("JSO")
'''    Set wfrm = New EmrFrmOrdMst
'''    'wfrm.Caption = "Trading Sales Order"
'''  Case Is = UCase("JSA")
'''    Set wfrm = New EmrFrmOrdMst
'''    'wfrm.Caption = "Trading Asset Sales"
'''  Case Is = UCase("JMI")
'''    Set wfrm = New EmrFrmOrdMst
'''    'wfrm.Caption = "Trading Memo Issue"
'''  Case Is = UCase("JMR")
'''    Set wfrm = New EmrFrmOrdMst
'''    'wfrm.Caption = "Trading Memo Return"
'''  Case Is = UCase("JMS")
'''    Set wfrm = New EmrFrmOrdMst
'''    'wfrm.Caption = "Trading Memo Sales"
'''  Case Is = UCase("JST")
'''    Set wfrm = New EmrFrmOrdMst
'''    'wfrm.Caption = "Trading Sales Return"
'''  Case Is = UCase("JPO")
'''    Set wfrm = New EmrFrmOrdMst
'''    'wfrm.Caption = "Trading Purchase Order"
'''  Case Is = UCase("JPR")
'''    Set wfrm = New EmrFrmOrdMst
'''    'wfrm.Caption = "Trading Purchase"
'''  Case Is = UCase("JCI")
'''    Set wfrm = New EmrFrmOrdMst
'''    'wfrm.Caption = "Trading Consignment Inward"
'''  Case Is = UCase("JCR")
'''    Set wfrm = New EmrFrmOrdMst
'''    'wfrm.Caption = "Trading Consignment Return"
'''  Case Is = UCase("JCP")
'''    Set wfrm = New EmrFrmOrdMst
'''    'wfrm.Caption = "Trading Consignment Purchase"
'''  Case Is = UCase("JPT")
'''    Set wfrm = New EmrFrmOrdMst
'''    'wfrm.Caption = "Trading Purchase Return"
'''  Case Is = UCase("JMV")
'''    Set wfrm = New EmrFrmOrdMst
'''    'wfrm.Caption = "Trading Design Movement"
'''  Case Is = UCase("JCN")
'''    Set wfrm = New EmrFrmOrdMst
'''    'wfrm.Caption = "Trading Design-Rm Conversion"
'''  Case Is = UCase("SPL")
'''    Set wfrm = New EmrFrmOrdMst
'''  Case Is = UCase("RepJtDet")
'''    Set wfrm = New EmrFrpOrd
'''    wfrm.Caption = "Trading Detail Printing"
'''  Case Is = UCase("RepJtAvg")
'''    Set wfrm = New EmrFrpOrd
'''    wfrm.Caption = "Trading Average Printing"
'''  Case Is = UCase("RepJtCat")
'''    Set wfrm = New EmrFrpOrd
'''    wfrm.Caption = "Trading Catalogue Printing"
'''  Case Is = UCase("RepJtStkLed")
'''    Set wfrm = New EmrFrpFgStkLed
'''    wfrm.Caption = "Jewellery Stock Ledger"
'''  Case Is = UCase("RepJtPnd")
'''    Set wfrm = New EmrFrpSOPnd
'''    wfrm.Caption = "Pending Voucher"
'''  Case Is = UCase("RepJtOpnOrd")
'''    Set wfrm = New EmrFrpSOPnd
'''    wfrm.Caption = "Open Order w.r.t JPO"
'''  Case Is = UCase("RepJtNetStk")
'''    Set wfrm = New EmrFrpJtNetStk
'''    wfrm.Caption = "Design Net Stock Requirement"
'''  Case Is = UCase("RepJtInv")
'''    Set wfrm = New EmrFrpInv
'''    wfrm.Caption = "Trading Invoice Printing"
'''  Case Is = UCase("RepJtGrFrnt")
'''    Set wfrm = New EmrFrpInv
'''    wfrm.Caption = "GR Form (Front)"
'''  Case Is = UCase("RepJtGrBk")
'''     Set wfrm = New EmrFrpInvLet
'''    wfrm.Caption = "GR Form (Back)"
'''  Case Is = UCase("RepJtGsp")
'''    Set wfrm = New EmrFrpInv
'''    wfrm.Caption = "GSP Form"
'''  Case Is = UCase("RepJtShpBill")
'''    Set wfrm = New EmrFrpInv
'''    wfrm.Caption = "Shipping Bill"
'''  Case Is = UCase("RepJtPckLst")
'''    Set wfrm = New EmrFrpPckLstInv
'''    wfrm.Caption = "Packing List (Trading Invoice)"
'''  Case Is = UCase("RepJtLclInv")
'''    Set wfrm = New EmrFrpLclInv
'''    wfrm.Caption = "Local Invoice Printing (Trading)"
    
  ' **** Manali Trading Module
   
  ' **** End Of Report Listing **** '


  '##################################################################################
  '##################################################################################
'  Case Is = UCase("RepWipFgBRm")
'    Set wfrm = New EmrFrpWipBRm
'    wfrm.Caption = "Fg Bag Raw Material Report"
'  Case Is = UCase("RepInvProfitMrg")
'    Set wfrm = New EmrFrpInvProfitMrg
'    wfrm.Caption = "Profit Margin"
'  Case Is = UCase("Zoom")
'    Set wfrm = New EmrFrmOrdZoom
  End Select

  If Not (wfrm Is Nothing) Then
    Dim ws_UaMnuForm As String
    ''ws_UaMnuForm = UCase(gocn.GetFldVal("Select UaMnuForm From UsrAccess where UaUsrCd= '" + gs_UsrCd + "' and UaMnuCd= '" + ps_Mnu + "' "))
    
    ws_UaMnuForm = UCase(GetMDIFrm.ADC.Connection.GetFldVal("Select UaMnuForm From UsrAccess where UaUsrCd= '" + gs_UsrCd + "' and UaMnuCd= '" + ps_Mnu + "' "))
    
    If ws_UaMnuForm = "MDI" Then
      ShowForm GetMDIFrm, wfrm, ps_Mnu
    Else
      '''Dim go_FrmLst As New EmrFrmList
      If (GetFrmLst Is Nothing) Then
        'Set go_FrmLst = New EmrFrmList
        Call SetFrmLst(New EmrFrmList)
        ShowForm GetMDIFrm, GetFrmLst, "RepLst"
      End If
      ShowForm GetFrmLst, wfrm, ps_Mnu
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

