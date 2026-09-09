Attribute VB_Name = "JMod"
Option Explicit
Public Const ctGenMenuUsr = "JMIC"
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
  If UCase(gs_UsrCd) = "MIC" And (UCase(ps_Mnu) <> UCase("UsrAcss") And UCase(ps_Mnu) <> UCase("Param") And UCase(ps_Mnu) <> UCase("UsrAdm") And UCase(ps_Mnu) <> UCase("RepLst") _
      And UCase(ps_Mnu) <> UCase("EmrUpd") And UCase(ps_Mnu) <> UCase("RepDscRep") And UCase(ps_Mnu) <> UCase("EmrToXML") And UCase(ps_Mnu) <> UCase("TagFmt")) Then _
      DispMsg "The User Has No Rights To View This Menu", etError: Exit Sub
  '****** User ADM Can Only View the Usr Administration Form ******
  If UCase(gs_UsrCd) = "ADM" And UCase(ps_Mnu) <> UCase("UsrAdm") Then DispMsg "The User Has No Rights To View This Menu", etError: Exit Sub
  '****** Sachin 2.14.0 - [Restrictions To Microway Users]

  Select Case UCase(ps_Mnu)
' **** MDI Menu Forms **** '
  Case Is = UCase("JSO")
    Set wfrm = New EmrFrmOrdMst
    'wfrm.Caption = "Trading Sales Order"
  ' **** Manali Trading Module
  Case Is = UCase("LSO")
    Set wfrm = New EmrFrmOrdMst
    'wfrm.Caption = "Trading Sales Order"
  Case Is = UCase("LSA")
    Set wfrm = New EmrFrmOrdMst
    'wfrm.Caption = "Trading Sales Order"
  Case Is = UCase("LST")
    Set wfrm = New EmrFrmOrdMst
    'wfrm.Caption = "Trading Sales Order"
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
   Case Is = UCase("RepJtStkReco")
    Set wfrm = New EmrFrpBagReco
    wfrm.Caption = "Stock Reconciliation Report"
  
  Case Is = UCase("VarParam")
    Set wfrm = New EmrFrmvParam
    wfrm.Caption = "Variable Parameters"
  Case Is = UCase("Bnk")
    Set wfrm = New EmrFrmBnkAg
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
  Case Is = UCase("Head")
    Set wfrm = New EmrFrmHead
  Case Is = UCase("LabRt")
    Set wfrm = New EmrFrmLabRt
  Case Is = UCase("RepLst")
'    *** Form with Menu Placed for all Listing Reports
    Set wfrm = New EmrFrmList
  Case Is = UCase("Loc")
    Set wfrm = New EmrFrmLoc
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

'  Case Is = UCase("TagFmt")
'    If GetMDIFrm.ADC.Connection.RecSeek("select 1 from dbo.sysobjects where id = object_id(N'[TagFmt]') and OBJECTPROPERTY(id, N'IsUserTable') = 1") _
'      And _
'      GetMDIFrm.ADC.Connection.RecSeek("select 1 from dbo.sysobjects where id = object_id(N'[TagFmtSr]') and OBJECTPROPERTY(id, N'IsUserTable') = 1") Then
'        '  If gs_Ver = "2.6.0" Or gs_Ver = "2.7.0" Or gs_Ver = "2.8.0" Or gs_Ver = "2.9.0" Then
'      Set wfrm = New EmrFrmTagFmt
'    End If
  Case Is = UCase("UsrAcss")
    Set wfrm = New EmrFrmUsrAcss
    wfrm.Caption = "User Access Master"
  Case Is = UCase("UsrAdm")
    Set wfrm = New EmrFrmUsrAdm
    wfrm.Caption = "User Administration"

  Case Is = UCase("UsrMail")
    Set wfrm = New EmrFrmUsrMail
    wfrm.Caption = "User Mail"
  Case Is = UCase("VarParam")
    Set wfrm = New EmrFrmvParam
    wfrm.Caption = "Variable Parameters"
  Case Is = UCase("EmrUpd")
    Set wfrm = New EmrSuppFrmUpd
  Case Is = UCase("RepDscRep")          '*** Jay 2.14 P2 *** [Dsc Rep Introduced]
    Set wfrm = New EmrSuppFrpDscRep
  Case Is = UCase("RepSOPnd")
    Set wfrm = New EmrFrpSOPnd
    wfrm.Caption = "Sales Order Pending"
  Case Is = UCase("RepBagHist")
    Set wfrm = New EmrFrpBagHist
    wfrm.Caption = "Bag History"
  Case Is = UCase("RmZMst")
    Set wfrm = New EmrFrmRmMst
    wfrm.Caption = "Customs Raw Material Master"
  Case Is = UCase("PrcLst")
    Set wfrm = New EmrFrmOrdMst
  Case Is = UCase("Quot")
    Set wfrm = New EmrFrmOrdMst
  Case Is = UCase("QB")
    Set wfrm = New EmrFrmOrdMst
  Case Is = UCase("SkQuot")
    Set wfrm = New EmrFrmOrdMst

' Geeta Emr 212
  Case Is = UCase("MF")
    Set wfrm = New EmrFrmOrdMst
' Geeta Emr 212

'  Case Is = UCase("RepDiaProfit")
'    Set wfrm = New EmrFrpDiaProfit
'    wfrm.Caption = "Diamond Profit"
'  Case Is = UCase("RepFlashStk")
'    Set wfrm = New EmrFrpFlashStk
'    wfrm.Caption = "Flash Stock"
  Case Is = UCase("RepInv")
    Set wfrm = New EmrFrpInv
    wfrm.Caption = "Invoice Printing"
  Case Is = UCase("RepGrFormFrnt")
    Set wfrm = New EmrFrpInv
    wfrm.Caption = "GR Form (Front)"
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
  Case Is = UCase("RepShpBill")
    Set wfrm = New EmrFrpInv
    wfrm.Caption = "Shipping Bill"
  '*** Jay 2.14(LclInv) ***
  Case Is = UCase("RepLclInv")
    Set wfrm = New EmrFrpLclInv
    wfrm.Caption = "Local Invoice Printing"
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

'  Case Is = UCase("RepInvProfitMrg")
'    Set wfrm = New EmrFrpInvProfitMrg
'    wfrm.Caption = "Profit Margin"
  Case Is = UCase("RepDetOrd")
    Set wfrm = New EmrFrpOrd
    wfrm.Caption = "Detail Order Printing"
  Case Is = UCase("RepAvgOrd")
    Set wfrm = New EmrFrpOrd
    wfrm.Caption = "Average Order Printing"
  Case Is = UCase("RepOrdCat")
    Set wfrm = New EmrFrpOrd
    wfrm.Caption = "Order Catalogue Printing"
'  Case Is = UCase("RepOrdProfitMrg")
'    Set wfrm = New EmrFrpOrdProfitMrg
'    wfrm.Caption = "Order Profit Margin"
'  Case Is = UCase("RepPckLst")
'    Set wfrm = New EmrFrpPckLst
'    wfrm.Caption = "Finished Goods Packing List Fr INV"
'  Case Is = UCase("RepPckLstInv")
'    Set wfrm = New EmrFrpPckLstInv
'    wfrm.Caption = "Invoice Packing List"
  Case Is = UCase("RepRmMst")
    Set wfrm = New EmrFrpRmMst
    wfrm.Caption = "Raw Material/ Custom Master Listing"
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
'  Case Is = UCase("RepSalAnaDsg")
'    Set wfrm = New EmrFrpSalAnaDsg
'    wfrm.Caption = "Sales Analysis (Design Level)"
'  Case Is = UCase("RepSalAnaRm")
'    Set wfrm = New EmrFrpSalAnaRm
'    wfrm.Caption = "Sales Analysis (Rm Level)"
'  Case Is = UCase("RepSalAnaLab")
'    Set wfrm = New EmrFrpSalAnaLab
'    wfrm.Caption = "Sales Analysis (Lab Level)"
  Case Is = UCase("RepStkLed")
    Set wfrm = New EmrFrpStkLed
    wfrm.Caption = "Stock Ledger"
'  Case Is = UCase("RepStkSumm")
'    Set wfrm = New EmrFrpStkSumm
'    wfrm.Caption = "Stock Summary"
  Case Is = UCase("RepWipFgBPcs")
    Set wfrm = New EmrFrpWipBPcs
    wfrm.Caption = "Fg Bag Pieces Report"
'  Case Is = UCase("RepWipFgBPcsBkDt")
'    Set wfrm = New EmrFrpWipBPcsBkDt
'    wfrm.Caption = "Fg Bag Pieces Report (Back Dated)"
  Case Is = UCase("RepWipFgBRm")
    Set wfrm = New EmrFrpWipBRm
    wfrm.Caption = "Fg Bag Raw Material Report"
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
  Case Is = UCase("RepGrade")
    Set wfrm = New EmrFrpGrade
    wfrm.Caption = "Grade Listing"
  Case Is = UCase("RepFbLst")
    Set wfrm = New EmrFrpFgList
    wfrm.Caption = "Finished Goods (Bag)"
  Case Is = UCase("RepFwbLst")
    Set wfrm = New EmrFrpFgList
    wfrm.Caption = "Finished Goods (Bag Return)"
'  Case Is = UCase("RepFgMvm")
'    Set wfrm = New EmrFrpFgMvm
'    wfrm.Caption = "Finished Goods (Bag Movement)"
  Case Is = UCase("RepFgBagLed")
    Set wfrm = New EmrFrpFgBagLed
    wfrm.Caption = "Finished Goods Bag Ledger"
  Case Is = UCase("RepCustRm")
    Set wfrm = New EmrFrpCustRm
    wfrm.Caption = "Customer RM Listing"
  Case Is = UCase("RepInvLoose")
    Set wfrm = New EmrFrpInv
    wfrm.Caption = "Invoice Printing (Loose RM)"
  Case Is = UCase("RepShpBillLoose")
    Set wfrm = New EmrFrpInv
    wfrm.Caption = "Shipping Bill (Loose RM)"
  Case Is = UCase("RepGrFormLoose")
    Set wfrm = New EmrFrpInv
    wfrm.Caption = "GR Form (Loose RM)"
  '****** Sachin 3.3.0 (28-06-08) - [27. Valid Rm Sz Master] ******
  Case Is = UCase("VldRmLn")
    Set wfrm = New EmrFrmVldRmLn
  '****** Sachin 3.3.0 (28-06-08) - [27. Valid Rm Sz Master] ******

'  '***** Manali 3.03 - 11/07/08 - RepRmIdSz
'  Case Is = UCase("RepRmIdSz")
'    Set wfrm = New EmrFrpRmIdSz
'  '***** Manali 3.03 - 11/07/08 - RepRmIdSz
  
  Case Is = UCase("JOT")
    Set wfrm = New EmrFrmOrdMst
  Case Is = UCase("JIN")
    Set wfrm = New EmrFrmOrdMst
  
  'uma *** Repair in & out
  Case Is = UCase("JRI")
    Set wfrm = New EmrFrmOrdMst
  Case Is = UCase("JRO")
    Set wfrm = New EmrFrmOrdMst
  
  '*** Jay 3.3(FG) ***
  Case Is = UCase("Fin")
    Set wfrm = New EmrFrmOrdMst
  Case Is = UCase("Fot")
    Set wfrm = New EmrFrmOrdMst
  Case Is = UCase("InvHdJT")
    Set wfrm = New EmrFrmInvHd
    wfrm.Caption = "Invoice Header "
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
'  Case Is = UCase("RepFgPckLst")
'    Set wfrm = New EmrFrpPckLstInv
'    wfrm.Caption = "Packing List (Fg Invoice)"
  Case Is = UCase("RepFgLclInv")
    Set wfrm = New EmrFrpLclInv
    wfrm.Caption = "Local Invoice Printing (FG)"
   '*** Jay 3.3(FG)
   
  ' **** Manali Trading Module
  Case Is = UCase("JSO")
    Set wfrm = New EmrFrmOrdMst
    'wfrm.Caption = "Trading Sales Order"
  ' **** Manali Trading Module
  Case Is = UCase("LSO")
    Set wfrm = New EmrFrmOrdMst
    'wfrm.Caption = "Trading Sales Order"
  Case Is = UCase("JSA")
    Set wfrm = New EmrFrmOrdMst
    'wfrm.Caption = "Trading Asset Sales"
  Case Is = UCase("JMI")
    Set wfrm = New EmrFrmOrdMst
    'wfrm.Caption = "Trading Memo Issue"
  Case Is = UCase("JMR")
    Set wfrm = New EmrFrmOrdMst
    'wfrm.Caption = "Trading Memo Return"
  Case Is = UCase("JMS")
    Set wfrm = New EmrFrmOrdMst
    'wfrm.Caption = "Trading Memo Sales"
  Case Is = UCase("JST")
    Set wfrm = New EmrFrmOrdMst
    'wfrm.Caption = "Trading Sales Return"
  Case Is = UCase("JPO")
    Set wfrm = New EmrFrmOrdMst
    'wfrm.Caption = "Trading Purchase Order"
  Case Is = UCase("JPR")
    Set wfrm = New EmrFrmOrdMst
    'wfrm.Caption = "Trading Purchase"
  Case Is = UCase("JCI")
    Set wfrm = New EmrFrmOrdMst
    'wfrm.Caption = "Trading Consignment Inward"
  Case Is = UCase("JCR")
    Set wfrm = New EmrFrmOrdMst
    'wfrm.Caption = "Trading Consignment Return"
  Case Is = UCase("JCP")
    Set wfrm = New EmrFrmOrdMst
    'wfrm.Caption = "Trading Consignment Purchase"
  Case Is = UCase("JPT")
    Set wfrm = New EmrFrmOrdMst
    'wfrm.Caption = "Trading Purchase Return"
  Case Is = UCase("JMV")
    Set wfrm = New EmrFrmOrdMst
    'wfrm.Caption = "Trading Design Movement"
  Case Is = UCase("JCN")
    Set wfrm = New EmrFrmOrdMst
    'wfrm.Caption = "Trading Design-Rm Conversion"
  Case Is = UCase("JMT")
    Set wfrm = New EmrFrmOrdMst
    'wfrm.Caption = "Trading Design-Rm Conversion"
  Case Is = UCase("SPL")
    Set wfrm = New EmrFrmOrdMst
  Case Is = UCase("SQT")  'SV.33 Supplier quotation added
    Set wfrm = New EmrFrmOrdMst
  Case Is = UCase("RepJtDet")
    Set wfrm = New EmrFrpOrd
    wfrm.Caption = "Detail Costing "
  Case Is = UCase("RepJtAvg")
    Set wfrm = New EmrFrpOrd
    wfrm.Caption = "Average Costing"
  Case Is = UCase("RepJtCat")
    Set wfrm = New EmrFrpOrd
    wfrm.Caption = "Costing Catalogue"
  Case Is = UCase("RepJtStkLed")
    Set wfrm = New EmrFrpFgStkLed
    wfrm.Caption = "Stock Ledger (Jwlry)"
  Case Is = UCase("RepJtPnd")
    Set wfrm = New EmrFrpSOPnd
    wfrm.Caption = "Pending Voucher (Jwlry)"
  Case Is = UCase("RepJtOpnOrd")
    Set wfrm = New EmrFrpSOPnd
    wfrm.Caption = "Open Reference Vchrs (Jwlry)"
  Case Is = UCase("RepJtNetStk")
    Set wfrm = New EmrFrpJtNetStk
    wfrm.Caption = "Net Stock Requirement (Jwlry)"
  Case Is = UCase("RepJtInv")
    Set wfrm = New EmrFrpInv
    wfrm.Caption = "Trading Invoice Printing"
  Case Is = UCase("RepJtGrFrnt")
    Set wfrm = New EmrFrpInv
    wfrm.Caption = "GR Form (Front)"
  Case Is = UCase("RepJtGrBk")
     Set wfrm = New EmrFrpInvLet
    wfrm.Caption = "GR Form (Back)"
  Case Is = UCase("RepJtGsp")
    Set wfrm = New EmrFrpInv
    wfrm.Caption = "GSP Form"
  Case Is = UCase("RepJtShpBill")
    Set wfrm = New EmrFrpInv
    wfrm.Caption = "Shipping Bill"
  Case Is = UCase("RepJtPckLst")
    Set wfrm = New EmrFrpPckLstInv
    wfrm.Caption = "Packing List (Trading Invoice)"
  Case Is = UCase("RepJtLclInv")
    Set wfrm = New EmrFrpLclInv
    wfrm.Caption = "Local Invoice Printing (Trading)"
  Case Is = UCase("RepJtVchAnaDsg")
    Set wfrm = New EmrFrpJtVchAna
    wfrm.Caption = "Voucher Analysis - Designwise (Jwlry)"
  Case Is = UCase("RepJtVchAnaRm")
    Set wfrm = New EmrFrpJtVchAna
    wfrm.Caption = "Voucher Analysis - RM (Jwlry)"
  Case Is = UCase("RepJtVchAnaLab")
    Set wfrm = New EmrFrpJtVchAna
    wfrm.Caption = "Voucher Analysis - Labour (Jwlry)"
  Case Is = UCase("RepJtStkRm")
    Set wfrm = New EmrFrpJtStkRm
    wfrm.Caption = "Stock Rm (Jwlry)"
  Case Is = UCase("RepJtPndRm")
    Set wfrm = New EmrFrpSOPnd
    wfrm.Caption = "Pending Voucher Rm (Rm Level) (Jwlry)"
    
  ' **** Manali Trading Module
  Case Is = UCase("LIN")
    Set wfrm = New EmrFrmOrdMst

  Case Is = UCase("LMI")
    Set wfrm = New EmrFrmOrdMst

  Case Is = UCase("LMR")
    Set wfrm = New EmrFrmOrdMst

  Case Is = UCase("LMS")
    Set wfrm = New EmrFrmOrdMst

  Case Is = UCase("LPO")
    Set wfrm = New EmrFrmOrdMst

  Case Is = UCase("LPR")
    Set wfrm = New EmrFrmOrdMst

  Case Is = UCase("LPT")
    Set wfrm = New EmrFrmOrdMst
  
  Case Is = UCase("LSA")
    Set wfrm = New EmrFrmOrdMst

  Case Is = UCase("LCI")
    Set wfrm = New EmrFrmOrdMst

Case Is = UCase("LSO")
    Set wfrm = New EmrFrmOrdMst
    'wfrm.Caption = "Trading Sales Order"
  Case Is = UCase("LSA")
    Set wfrm = New EmrFrmOrdMst
    'wfrm.Caption = "Trading Sales Order"
  Case Is = UCase("LST")
    Set wfrm = New EmrFrmOrdMst
    'wfrm.Caption = "Trading Sales Order"
  Case Is = UCase("InvHdJT")
    Set wfrm = New EmrFrmInvHd
    wfrm.Caption = "Invoice Header "

  Case Is = UCase("LIN")
    Set wfrm = New EmrFrmOrdMst

  Case Is = UCase("LOT")
    Set wfrm = New EmrFrmOrdMst
  Case Is = UCase("LMI")
    Set wfrm = New EmrFrmOrdMst

  Case Is = UCase("LMR")
    Set wfrm = New EmrFrmOrdMst
  Case Is = UCase("LCN")
    Set wfrm = New EmrFrmOrdMst

  Case Is = UCase("LMV")
    Set wfrm = New EmrFrmOrdMst
  Case Is = UCase("LMS")
    Set wfrm = New EmrFrmOrdMst

  Case Is = UCase("LPO")
    Set wfrm = New EmrFrmOrdMst

  Case Is = UCase("LPR")
    Set wfrm = New EmrFrmOrdMst

  Case Is = UCase("LPT")
    Set wfrm = New EmrFrmOrdMst
  Case Is = UCase("LSA")
    Set wfrm = New EmrFrmOrdMst
  Case Is = UCase("LCI")
    Set wfrm = New EmrFrmOrdMst
  Case Is = UCase("LCP")
    Set wfrm = New EmrFrmOrdMst
  Case Is = UCase("LCR")
    Set wfrm = New EmrFrmOrdMst
  Case Is = UCase("RLRMPO")
    Set wfrm = New EmrFrpLRmPndVch
    wfrm.Caption = "Pending Voucher (Loose Rm)    "
  Case Is = UCase("RLRMSTK")
    Set wfrm = New EmrFrpLRmStkLed
    wfrm.Caption = "Stock Ledger (Loose Rm)"
  Case Is = UCase("RLRMTL")
    Set wfrm = New EmrFrpLRmTxnLst
    wfrm.Caption = "Transaction Listing (Loose Rm)"
  Case Is = UCase("RLRMVPR")
    Set wfrm = New EmrFrpLRmVchPrn
    wfrm.Caption = "Voucher Printing (Loose Rm)"
  'U.jEmr.17 - New form for Voucher Profit margin added
  Case Is = UCase("JPRFMRGN")
    Set wfrm = New EmrFrpVchProfitMrg
    wfrm.Caption = "Profit Margin Report"
  'PG.7 new menu added For Design Flash Stock Report
  Case Is = UCase("RepJtDsgFlashStk")
    Set wfrm = New EmrFrpJtStkRm
    wfrm.Caption = "Design Flash Stock (Jwlry)"
  'PG.14
  Case Is = UCase("RepJtTxnLstRpt")
    Set wfrm = New EmrFrpJtTxnLst
    wfrm.Caption = "Transaction Listing - Jwlry"
  'PG.19
  Case Is = UCase("JtStkRecoEntry")
    Set wfrm = New EmrFrmStkReco
    wfrm.Caption = "Stock Reconciliation"
  'mw.118- Stock zoom
  Case Is = UCase("RepJtStkZoom")
    Set wfrm = New EmrFrmJtStkZoom
    wfrm.Caption = "Jewellery Trading Stock Zoom"
 'mw.119-jEmr zoom
  Case Is = UCase("RepJtZoom")
    Set wfrm = New EmrFrmJtZoom
    wfrm.Caption = "Jewellery Trading Zoom"
 'VK.27-New menu "Pending voucher rm(Dsg level)" added.
  Case Is = UCase("RepJtSOPndRm")
    Set wfrm = New EmrFrpSOPnd
    wfrm.Caption = "Pending voucher RM (Design Level) (Jwlry)"
  'VK.25 - New menu added
  Case Is = UCase("RepJtTagPrn")
    Set wfrm = New EmrFrmTag
    wfrm.Caption = "Tag Printing (Jwlry)"
  'vk.22 JQB added
  Case Is = UCase("JQB")
    Set wfrm = New EmrFrmOrdMst
    
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


