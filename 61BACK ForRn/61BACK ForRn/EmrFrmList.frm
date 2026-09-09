VERSION 5.00
Object = "{AB2CC32A-E652-42A2-B341-26ECE02F6386}#2.1#0"; "MwfCtlForSpeed610.ocx"
Begin VB.Form EmrFrmList 
   Caption         =   "Additional Entries & Reports"
   ClientHeight    =   9495
   ClientLeft      =   60
   ClientTop       =   630
   ClientWidth     =   15180
   DrawStyle       =   5  'Transparent
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MDIChild        =   -1  'True
   ScaleHeight     =   9495
   ScaleWidth      =   15180
   WindowState     =   2  'Maximized
   Begin VB.Frame FraAdc 
      BorderStyle     =   0  'None
      Height          =   675
      Left            =   0
      TabIndex        =   3
      Top             =   0
      Width           =   1035
      Begin MwfCtl.MWCTL_ADC ADC 
         CausesValidation=   0   'False
         Height          =   555
         Left            =   30
         TabIndex        =   4
         Top             =   120
         Width           =   1995
         _ExtentX        =   3519
         _ExtentY        =   979
         AllowAdd        =   0   'False
         AllowFind       =   0   'False
         AllowSave       =   0   'False
         AllowDelete     =   0   'False
         BtnWidth        =   885
         BtnHeight       =   435
         BtnMaskColor    =   8454143
         FontSz          =   8.25
      End
   End
   Begin VB.Frame Fra_NKeys 
      BorderStyle     =   0  'None
      Height          =   495
      Left            =   2430
      TabIndex        =   1
      Top             =   60
      Visible         =   0   'False
      Width           =   435
   End
   Begin VB.Frame FRA_KEYS 
      BorderStyle     =   0  'None
      Height          =   525
      Left            =   30
      TabIndex        =   0
      Top             =   675
      Width           =   5835
      Begin VB.Label LblInfo 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Close All Reports, Which Open From This Form Menu"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Left            =   90
         TabIndex        =   5
         Top             =   180
         Width           =   3810
      End
   End
   Begin MwfCtl.MWCTL_HLP MWCTL_HLP1 
      Height          =   405
      Left            =   1380
      TabIndex        =   2
      TabStop         =   0   'False
      Top             =   0
      Visible         =   0   'False
      Width           =   540
      _ExtentX        =   953
      _ExtentY        =   714
   End
   Begin VB.Menu MnuListMas 
      Caption         =   "&Master"
      Tag             =   "1"
      Begin VB.Menu RepPrm 
         Caption         =   "Parameter Listing"
         Tag             =   "11"
      End
      Begin VB.Menu RepVarPrm 
         Caption         =   "Variable Parameters Listing"
         Tag             =   "12"
      End
      Begin VB.Menu RepUsrAccess 
         Caption         =   "User Access List"
         Tag             =   "13"
      End
      Begin VB.Menu RepCust 
         Caption         =   "Customer/Supplier"
         Tag             =   "14"
         Begin VB.Menu RepCustLst 
            Caption         =   "Customer Listing"
            Tag             =   "141"
         End
         Begin VB.Menu RepCustDsg 
            Caption         =   "Customer Design Listing"
            Tag             =   "142"
         End
         Begin VB.Menu RepCustAly 
            Caption         =   "Customer Alloy Listing"
            Tag             =   "143"
         End
         Begin VB.Menu RepCustRm 
            Caption         =   "Customer Rm Listing"
            Tag             =   "144"
         End
         Begin VB.Menu D11 
            Caption         =   "DUM"
         End
      End
      Begin VB.Menu RepRmMst 
         Caption         =   "Raw Material/ Custom Listing"
         Tag             =   "15"
      End
      Begin VB.Menu RepDsg 
         Caption         =   "Design Listing"
         Tag             =   "16"
      End
      Begin VB.Menu RepDsgCrd 
         Caption         =   "Design Card"
         Tag             =   "17"
      End
      Begin VB.Menu RepDsgCat 
         Caption         =   "Design Catalogue"
         Tag             =   "18"
      End
      Begin VB.Menu RepRt 
         Caption         =   "Rate Master Listing"
         Tag             =   "19"
         Begin VB.Menu RepRmRt 
            Caption         =   "Raw Material Rate Master "
            Tag             =   "191"
         End
         Begin VB.Menu RepLabRt 
            Caption         =   "Labour Rate Master"
            Tag             =   "192"
         End
         Begin VB.Menu RepRtCmp 
            Caption         =   "Rate Comparision"
            Tag             =   "193"
         End
         Begin VB.Menu D12 
            Caption         =   "DUM"
         End
      End
      Begin VB.Menu RepLoc 
         Caption         =   "Location Listing"
         Tag             =   "1A"
      End
      Begin VB.Menu RepDtTbl 
         Caption         =   "Date Table Listing"
         Tag             =   "1B"
      End
      Begin VB.Menu RepBnk 
         Caption         =   "Bank/Agent Listing"
         Tag             =   "1C"
      End
      Begin VB.Menu RepBagLst 
         Caption         =   "Bag Listing"
         Tag             =   "1D"
      End
      Begin VB.Menu RepLsRec 
         Caption         =   "Loss Recovery Listing"
         Tag             =   "1E"
      End
      Begin VB.Menu RepPrdPts 
         Caption         =   "Production Points"
         Tag             =   "1F"
      End
      Begin VB.Menu RepLossMst 
         Caption         =   "Loss Master listing"
         Tag             =   "1G"
      End
      Begin VB.Menu RepRmIdSz 
         Caption         =   "RmId Size Desc Listing"
         Tag             =   "1H"
      End
      Begin VB.Menu RepRfIdMst 
         Caption         =   "Bag Alias No. Listing"
         Tag             =   "1I"
      End
      Begin VB.Menu D1 
         Caption         =   "DUM"
      End
   End
   Begin VB.Menu Rpr1 
      Caption         =   "Repair Bag"
      Tag             =   "8"
      Begin VB.Menu RDT 
         Caption         =   "Bag Transaction"
         Tag             =   "81"
      End
      Begin VB.Menu RBV 
         Caption         =   "Bag Movement"
         Tag             =   "82"
      End
      Begin VB.Menu RprRep 
         Caption         =   "Reports"
         Tag             =   "83"
         Begin VB.Menu RprWIP 
            Caption         =   "WIP Bag"
            Tag             =   "831"
         End
         Begin VB.Menu RprWIPBv 
            Caption         =   "WIP Bag Movement"
            Tag             =   "832"
         End
         Begin VB.Menu RprHis 
            Caption         =   "Bag History"
            Tag             =   "833"
         End
         Begin VB.Menu RprDUM 
            Caption         =   "DUM"
         End
      End
      Begin VB.Menu Dtst 
         Caption         =   "DUM"
      End
   End
   Begin VB.Menu Cert 
      Caption         =   "&Certificate"
      Tag             =   "A"
      Begin VB.Menu CertEntry 
         Caption         =   "Certificate Entry"
         Tag             =   "A1"
      End
      Begin VB.Menu RepCertLst 
         Caption         =   "Certificate Listing"
         Tag             =   "A2"
      End
      Begin VB.Menu DCert 
         Caption         =   "DUM"
      End
   End
   Begin VB.Menu Ppc 
      Caption         =   "&PPC"
      Tag             =   "B"
      Begin VB.Menu PpcCap 
         Caption         =   "Ppc Capacity"
         Tag             =   "B1"
      End
      Begin VB.Menu PpcDsgLd 
         Caption         =   "Design Load"
         Tag             =   "B2"
      End
      Begin VB.Menu PpcOrdPln 
         Caption         =   "Order Planning"
         Tag             =   "B3"
      End
      Begin VB.Menu DPpc 
         Caption         =   "DUM"
      End
   End
   Begin VB.Menu MnuLstDailyRep 
      Caption         =   "&Reports"
      Tag             =   "2"
      Begin VB.Menu MnuLstFg 
         Caption         =   "FG"
         Tag             =   "21"
         Begin VB.Menu RepFbLst 
            Caption         =   "Finished Goods (Bag)"
            Tag             =   "211"
         End
         Begin VB.Menu RepFwbLst 
            Caption         =   "Finished Goods (Bag Return)"
            Tag             =   "212"
         End
         Begin VB.Menu RepFgMvm 
            Caption         =   "Finished Goods (Bag Movement)"
            Tag             =   "213"
         End
         Begin VB.Menu RepBrkMsgFg 
            Caption         =   "Breakage Missing (FG)"
            Tag             =   "214"
         End
         Begin VB.Menu RepFGCatalog 
            Caption         =   "FG Catalog"
            Tag             =   "215"
         End
         Begin VB.Menu D21 
            Caption         =   "DUM"
         End
      End
      Begin VB.Menu REPORDLST 
         Caption         =   "Order\Quotation"
         Tag             =   "26"
         Begin VB.Menu REPTRAVEXCEL 
            Caption         =   "Travel Excel"
            Tag             =   "261"
         End
         Begin VB.Menu REPLOTZOOM 
            Caption         =   "Lot Zoom"
            Tag             =   "262"
         End
         Begin VB.Menu D21A 
            Caption         =   "DUM"
         End
      End
      Begin VB.Menu REPBR 
         Caption         =   "Bag Reconciliation"
         Tag             =   "914"
         Begin VB.Menu RC 
            Caption         =   "Bag Reconciliation Entry"
            Tag             =   "221"
         End
         Begin VB.Menu REPBAGRECO 
            Caption         =   "Bag Reconciliation Report"
            Tag             =   "222"
         End
         Begin VB.Menu D22 
            Caption         =   "DUM"
         End
      End
      Begin VB.Menu RepInCTB 
         Caption         =   "CTB"
         Tag             =   "23"
         Begin VB.Menu RepDetInvCTB 
            Caption         =   "Detail CTB Printing"
            Tag             =   "231"
         End
         Begin VB.Menu RepAvgInvCTB 
            Caption         =   "Average CTB Printing"
            Tag             =   "232"
         End
         Begin VB.Menu RepInvCTBCat 
            Caption         =   "CTB Catalogue"
            Tag             =   "233"
         End
         Begin VB.Menu RepCTBPckLst 
            Caption         =   "CTB Packing List"
            Tag             =   "234"
         End
         Begin VB.Menu D23 
            Caption         =   "DUM"
         End
      End
      Begin VB.Menu RepPpc 
         Caption         =   "PPC"
         Tag             =   "24"
         Begin VB.Menu RepDsgLd 
            Caption         =   "Design Load Listing"
            Tag             =   "241"
         End
         Begin VB.Menu RepCntrlCpy 
            Caption         =   "Controller Copy"
            Tag             =   "242"
         End
         Begin VB.Menu RepOrdPln 
            Caption         =   "Order Plan"
            Tag             =   "243"
         End
         Begin VB.Menu D24 
            Caption         =   "DUM"
         End
      End
      Begin VB.Menu RepRWrk 
         Caption         =   "Rework Report"
         Tag             =   "25"
      End
      Begin VB.Menu D2 
         Caption         =   "DUM"
      End
   End
   Begin VB.Menu MnuFGTxn 
      Caption         =   "&FG Transaction"
      Tag             =   "7"
      Begin VB.Menu MnuFgInOut 
         Caption         =   "FG IN/OUT"
         Tag             =   "71"
         Begin VB.Menu FIN 
            Caption         =   "FG In"
            Tag             =   "711"
         End
         Begin VB.Menu InvHdFot 
            Caption         =   "Invoice Header (FG)"
            Tag             =   "712"
         End
         Begin VB.Menu FOT 
            Caption         =   "FG Out"
            Tag             =   "713"
         End
         Begin VB.Menu D91 
            Caption         =   "DUM"
         End
      End
      Begin VB.Menu RepFgIO 
         Caption         =   "FG Rep"
         Tag             =   "72"
         Begin VB.Menu RepFGDet 
            Caption         =   "Detail FG Printing"
            Tag             =   "721"
         End
         Begin VB.Menu RepFgAvg 
            Caption         =   "Average FG Printing"
            Tag             =   "722"
         End
         Begin VB.Menu RepFGCat 
            Caption         =   "FG Catalogue"
            Tag             =   "723"
         End
         Begin VB.Menu RepFgStkLed 
            Caption         =   "Stock Ledger (FG)"
            Tag             =   "724"
         End
         Begin VB.Menu RepFgInv 
            Caption         =   "FG Invoice Printing"
            Tag             =   "725"
         End
         Begin VB.Menu RepFgPckLst 
            Caption         =   "Packing List (FG Invoice)"
            Tag             =   "726"
         End
         Begin VB.Menu RepFgShpBill 
            Caption         =   "Shipping Bill"
            Tag             =   "727"
         End
         Begin VB.Menu RepFgGrFrnt 
            Caption         =   "GR Form Front"
            Tag             =   "728"
         End
         Begin VB.Menu RepFgGrBk 
            Caption         =   "GR Form Back"
            Tag             =   "729"
         End
         Begin VB.Menu RepFgGsp 
            Caption         =   "GSP Form"
            Tag             =   "72A"
         End
         Begin VB.Menu RepFgLclInv 
            Caption         =   "Local Invoice"
            Tag             =   "72B"
         End
         Begin VB.Menu D82 
            Caption         =   "DUM"
         End
      End
      Begin VB.Menu D8 
         Caption         =   "DUM"
      End
   End
   Begin VB.Menu MnuLstTxn 
      Caption         =   "&Transactions"
      Tag             =   "3"
      Begin VB.Menu RepTxnDT 
         Caption         =   "Daily Transactions"
         Tag             =   "31"
      End
      Begin VB.Menu RepTxnBS 
         Caption         =   "Bag Split"
         Tag             =   "32"
      End
      Begin VB.Menu RepTxnBM 
         Caption         =   "Bag Merge"
         Tag             =   "33"
      End
      Begin VB.Menu RepTxnMLT 
         Caption         =   "Melting"
         Tag             =   "34"
      End
      Begin VB.Menu RepTxnBV 
         Caption         =   "Bag Movement"
         Tag             =   "35"
      End
      Begin VB.Menu RepTxnREJ 
         Caption         =   "Rejection"
         Tag             =   "36"
      End
      Begin VB.Menu RepTxnPR 
         Caption         =   "Purchase/Sales"
         Tag             =   "37"
      End
      Begin VB.Menu RepTxnMV 
         Caption         =   "RM Movement"
         Tag             =   "38"
      End
      Begin VB.Menu RepTxnCNV 
         Caption         =   "Conversion"
         Tag             =   "39"
      End
      Begin VB.Menu RepUnMtchCnv 
         Caption         =   "Unmatched Conversion"
         Tag             =   "3A"
      End
      Begin VB.Menu RepTxnRetMem 
         Caption         =   "Return Memo"
         Tag             =   "3B"
      End
      Begin VB.Menu RepTree 
         Caption         =   "Tree"
         Tag             =   "3C"
         Begin VB.Menu RepTreeInvst 
            Caption         =   "Tree Investment"
            Tag             =   "3C1"
         End
         Begin VB.Menu RepTreePerform 
            Caption         =   "Tree Performance"
            Tag             =   "3C2"
         End
         Begin VB.Menu D31 
            Caption         =   "DUM"
         End
      End
      Begin VB.Menu RepTxnCustmPurLst 
         Caption         =   "Custom Purchase"
         Tag             =   "3D"
      End
      Begin VB.Menu RepInvLst 
         Caption         =   "Invoice Amounts"
         Tag             =   "3E"
      End
      Begin VB.Menu RepOpnVch 
         Caption         =   "PO Pending Voucher"
         Tag             =   "3F"
      End
      Begin VB.Menu D3 
         Caption         =   "DUM"
      End
   End
   Begin VB.Menu MnuLstVchr 
      Caption         =   "&Voucher"
      Tag             =   "4"
      Begin VB.Menu RepVchPrnDT 
         Caption         =   "Daily Transactions"
         Tag             =   "41"
      End
      Begin VB.Menu RepVchPrnBS 
         Caption         =   "Bag Split"
         Tag             =   "42"
      End
      Begin VB.Menu RepVchPrnBM 
         Caption         =   "Bag Merge"
         Tag             =   "43"
      End
      Begin VB.Menu RepVchPrnMLT 
         Caption         =   "Melting"
         Tag             =   "44"
      End
      Begin VB.Menu RepVchPrnBV 
         Caption         =   "Bag Movement"
         Tag             =   "45"
      End
      Begin VB.Menu RepVchPrnREJ 
         Caption         =   "Rejection"
         Tag             =   "46"
      End
      Begin VB.Menu RepVchPrnPR 
         Caption         =   "Purchase/Sales"
         Tag             =   "47"
      End
      Begin VB.Menu RepVchPrnMV 
         Caption         =   "Raw Material Movement"
         Tag             =   "48"
      End
      Begin VB.Menu RepVchPrnCNV 
         Caption         =   "Conversion"
         Tag             =   "49"
      End
      Begin VB.Menu RepVchPrnRetMem 
         Caption         =   "Return Memo"
         Tag             =   "4A"
      End
      Begin VB.Menu RepVchPrnCustmPur 
         Caption         =   "Custom Purchase"
         Tag             =   "4B"
      End
      Begin VB.Menu D4 
         Caption         =   "DUM"
      End
   End
   Begin VB.Menu RepInvList 
      Caption         =   "&Invoice"
      Tag             =   "9"
      Begin VB.Menu EPS 
         Caption         =   "Export Performance"
         Tag             =   "91"
         Begin VB.Menu RepEPSPlainJwl 
            Caption         =   "Plain Jewellery"
            Tag             =   "911"
         End
         Begin VB.Menu RepEPSStudJwl 
            Caption         =   "Studded Jewellery (Monthly)"
            Tag             =   "912"
         End
         Begin VB.Menu RepEPSMetImpEx 
            Caption         =   "Metal Import Export (Monthly)"
            Tag             =   "913"
         End
         Begin VB.Menu D92 
            Caption         =   "DUM"
         End
      End
      Begin VB.Menu RepStudStmt 
         Caption         =   "Statement of Studding"
         Tag             =   "92"
      End
      Begin VB.Menu RepCustmStkLed 
         Caption         =   "Customs Stock Ledger By DATE"
         Tag             =   "98"
      End
      Begin VB.Menu RepEDFFormFrnt 
         Caption         =   "EDF Form Front"
         Tag             =   "93"
      End
      Begin VB.Menu RepEDFFormBck 
         Caption         =   "EDF Form Back"
         Tag             =   "94"
      End
      Begin VB.Menu RepInvLclRM 
         Caption         =   "Local Invoice RM"
         Tag             =   "95"
      End
      Begin VB.Menu RepGrFormFrntNew 
         Caption         =   "GR Form (Front)-Preprinted"
         Tag             =   "96"
      End
      Begin VB.Menu RepGrFormFrntNewMul 
         Caption         =   "GR Form(Frnt)-Preprint(Multi)"
         Tag             =   "97"
      End
      Begin VB.Menu D93 
         Caption         =   "DUM"
      End
   End
   Begin VB.Menu MTN 
      Caption         =   "Mai&ntenance"
      Tag             =   "5"
      Begin VB.Menu EmrUpd 
         Caption         =   "Emperor Updates"
         Tag             =   "51"
      End
      Begin VB.Menu RepDscRep 
         Caption         =   "Discrepancy Report"
         Tag             =   "54"
      End
      Begin VB.Menu Trf 
         Caption         =   "Transfers"
         Tag             =   "52"
         Begin VB.Menu EmrToXML 
            Caption         =   "Emr To XML Transfer"
            Tag             =   "521"
         End
         Begin VB.Menu D51 
            Caption         =   "DUM"
         End
      End
      Begin VB.Menu TagFmt 
         Caption         =   "Tag Format"
         Tag             =   "53"
      End
      Begin VB.Menu ClrDfn 
         Caption         =   "Color Definition"
         Tag             =   "55"
      End
      Begin VB.Menu ClrSel 
         Caption         =   "Color Selection"
         Tag             =   "56"
      End
      Begin VB.Menu RepSchedule 
         Caption         =   "Report Scheduling"
         Tag             =   "9957"
      End
      Begin VB.Menu ResAuto 
         Caption         =   "Reset Autogen"
         Tag             =   "58"
      End
      Begin VB.Menu IdeDfn 
         Caption         =   "IDE Definition"
         Tag             =   "59"
      End
      Begin VB.Menu IdeUsrRight 
         Caption         =   "IDE User Rights"
         Tag             =   "5A"
      End
      Begin VB.Menu CustMaint 
         Caption         =   "Customer Maintenance"
      End
      Begin VB.Menu D5 
         Caption         =   "DUM"
      End
   End
   Begin VB.Menu MnuTlBr 
      Caption         =   "T&oolBar"
      Tag             =   "6"
   End
   Begin VB.Menu WINDOW 
      Caption         =   "&Window"
      WindowList      =   -1  'True
      Begin VB.Menu CSC 
         Caption         =   "Cascade"
      End
      Begin VB.Menu TIL 
         Caption         =   "Tile"
      End
      Begin VB.Menu GENMNU 
         Caption         =   "Generate Menu"
      End
      Begin VB.Menu D7 
         Caption         =   "DUM"
      End
   End
End
Attribute VB_Name = "EmrFrmList"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim ms_MnuForm As String
Private Sub ADC_Load()
  ms_MnuForm = "LST"
  If (GetFrmLst Is Nothing) Then Call SetFrmLst(Me)   '*** (Jenny Color) (12/08/05)
End Sub

Private Sub CertEntry_Click()
'VK.5
  Call FormSelect("CertEntry")
End Sub

' ***** Manali 3.9.1 - 21/02/12 - Color Mnus shifted to AddonMnus
Private Sub ClrDfn_Click()
  Call FormSelect("CLRDFN")
End Sub

' ***** Manali 3.9.1 - 21/02/12 - Color Mnus shifted to AddonMnus
Private Sub ClrSel_Click()
  Call FormSelect("CLRSEL")
End Sub

Private Sub ContractorMst_Click()
 'Contractor Master shifted to Add-On Menu in 3.11.0
 Call FormSelect("ContractorMst")
End Sub



Private Sub CustMaint_Click()
    Call FormSelect("CustMaint")      '3.11.4

End Sub

Private Sub IdeUsrRight_Click()
    Call FormSelect("IdeUsrRight")      '3.11.4
End Sub

Private Sub IdeDFN_Click()
    Call FormSelect("IdeDFN")      '3.11.4
End Sub

Private Sub RBV_Click()
 'uma ** repair DT
  Call FormSelect("RBV")
End Sub

Private Sub RDT_Click()
 'uma ** repair DT
  Call FormSelect("RDT")
End Sub

Private Sub RepBrkMsgFg_Click()
    Call FormSelect("RepBrkMsgFg")
End Sub
'vk.5
Private Sub RepCertLst_Click()
  Call FormSelect("RepCertLst")
End Sub

Private Sub RepCustmStkLed_Click()
    Call FormSelect("RepCustmStkLed")
End Sub

Private Sub RepEDFFormBck_Click()
    Call FormSelect("RepEDFFormBck")    '4.1.0.0
End Sub

Private Sub RepEDFFormFrnt_Click()
    Call FormSelect("RepEDFFormFrnt")   '4.1.0.0
End Sub

Private Sub RepEPSMetImpEx_Click()
    '3.11.1.0
    Call FormSelect("RepEPSMetImpEx")
End Sub

Private Sub RepEPSPlainJwl_Click()
    '3.11.1.0
    Call FormSelect("RepEPSPlainJwl")
End Sub

Private Sub RepEPSStudJwl_Click()
    '3.11.1.0
    Call FormSelect("RepEPSStudJwl")
End Sub

Private Sub RepFGCatalog_Click()
    'geeta
    Call FormSelect("RepFgCatalog")
End Sub
Private Sub RepGrFormFrntNew_Click()
     Call FormSelect("RepGrFormFrntNew")    '3.11.4
End Sub
Private Sub RepGrFormFrntNewMul_Click()
      Call FormSelect("RepGrFormFrntNewMul")    '3.11.4
End Sub

Private Sub RepInvLclRM_Click()
    '4.1.2.0
    Call FormSelect("RepInvLclRM")
End Sub

Private Sub RepRWrk_Click()
Call FormSelect("RepRWrk")
End Sub
 
Private Sub RepSchedule_Click()
    'geeta
    Call FormSelect("RepSchedule")
End Sub

Private Sub RepLotZoom_Click()
  Call FormSelect("RepLOTZOOM")
End Sub

Private Sub RepStudStmt_Click()
    '3.11.1.0
    Call FormSelect("RepStudStmt")
End Sub

' ***** Manali 3.10.0 - 10/05/12 - Reset Autogen shifted to AddonMnus
Private Sub ResAuto_Click()
  Call FormSelect("ResAuto")
End Sub

'''  ***** Std Code not to be changed **********************************
Private Sub Form_KeyUp(KeyCode As Integer, Shift As Integer)
  If ADC.HotKeys(KeyCode, Shift) Then
          KeyCode = 0
          Shift = 0
  End If
End Sub
Private Sub Form_Activate()
  Call FrmActivate(Me)
' ########################  Manoj #### Ver: 2.0.6 #### Date: 03/06/2004  ########################
  Dim wc_Ctl As Control, mb_Flg As Boolean
  For Each wc_Ctl In Controls
    If TypeOf wc_Ctl Is Menu Then
      If wc_Ctl.Enabled = False Then
        mb_Flg = True
        Exit For
      End If
    End If
  Next
  If mb_Flg = False Then
    ADC.Visible = True
    FRA_KEYS.Visible = False
  Else
    FRA_KEYS.Visible = True
    DispMsg LblInfo.Caption, etInfo
  End If
' ###########################################  Manoj  ###########################################
End Sub
Private Sub Form_Deactivate()
  Call FrmDeActivate(Me)
  ADC.Visible = False   '*** Jenny Color (12/08/05)
End Sub
Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
  If UnloadMode <> vbFormCode Then Cancel = True: Exit Sub
  Call SetFrmLst(Nothing)    '*** Jenny Color (13/08/05)
  'Set go_FrmLst = Nothing '*** Jenny Color (12/08/05)
End Sub

Private Sub SPL_Click()
  Call FormSelect("SPL")
End Sub

Private Sub JCI_Click()
  Call FormSelect("JCI")
End Sub

Private Sub JCN_Click()
  Call FormSelect("JCN")
End Sub

Private Sub JCP_Click()
  Call FormSelect("JCP")
End Sub

Private Sub JCR_Click()
  Call FormSelect("JCR")
End Sub

Private Sub JMI_Click()
  Call FormSelect("JMI")
End Sub

Private Sub JMR_Click()
  Call FormSelect("JMR")
End Sub

Private Sub JMS_Click()
  Call FormSelect("JMS")
End Sub

Private Sub JMV_Click()
  Call FormSelect("JMV")
End Sub

Private Sub JPO_Click()
  Call FormSelect("JPO")
End Sub

Private Sub JPR_Click()
  Call FormSelect("JPR")
End Sub

Private Sub JPT_Click()
  Call FormSelect("JPT")
End Sub

Private Sub JSA_Click()
  Call FormSelect("JSA")
End Sub

Private Sub JSO_Click()
  Call FormSelect("JSO")
End Sub

Private Sub JST_Click()
  Call FormSelect("JST")
End Sub

' ########################  Manoj #### Ver: 2.0.6 #### Date: 03/06/2004  ########################
Private Sub MnuTlBr_Click()
  If GetMDIFrm.TlBar.Visible = True Then
    GetMDIFrm.TlBar.Visible = False
    'MnuTlBr.Caption = "&ToolBar"
  Else
    GetMDIFrm.TlBar.Top = 0
    If GetMDIFrm.CmdCtlToolBar.Count = 1 And GetMDIFrm.CmdCtlToolBar(0).Tag = "" Then Exit Sub
    GetMDIFrm.TlBar.Visible = True
    'MnuTlBr.Caption = "&Close ToolBar"
  End If
End Sub
Private Sub RepGrade_Click()
  Call FormSelect("RepGrade")
End Sub
' Zubin 213
Private Sub RC_Click()
  Call FormSelect("RC")
End Sub
' Zubin 213
Private Sub REPBAGRECO_Click()
  Call FormSelect("RepBagReco")
End Sub
Private Sub RepFgAvg_Click()
'*** Jay 3.4(FG)
  Call FormSelect("RepFgAvg")
End Sub

Private Sub Fin_Click()
  '*** Jay 3.3(FG) ***
  Call FormSelect("Fin")
End Sub
Private Sub Fot_Click()
  '*** Jay 3.3(FG) ***
  Call FormSelect("Fot")
End Sub
Private Sub InvHdFot_Click()
  '*** Jay 3.3(FG) ***
  Call FormSelect("InvHdFot")
End Sub
Private Sub RepFgCat_Click()
'*** Jay 3.4(FG)
  Call FormSelect("RepFgCat")
End Sub

Private Sub RepFGDet_Click()
  '*** Jay 3.4(FG)
  Call FormSelect("RepFGDet")
End Sub

Private Sub RepFgGrBk_Click()
'*** Jay 3.4(FG)
  Call FormSelect("RepFgGrBk")
End Sub

Private Sub RepFgGrFrnt_Click()
'*** Jay 3.4(FG)
  Call FormSelect("RepFgGrFrnt")
End Sub

Private Sub RepFgGsp_Click()
'*** Jay 3.4(FG)
  Call FormSelect("RepFgGsp")
End Sub

Private Sub RepFgInv_Click()
'*** Jay 3.4(FG)
  Call FormSelect("RepFgInv")
End Sub

Private Sub RepFgLclInv_Click()
'*** Jay 3.4(FG)
  Call FormSelect("RepFgLclInv")
End Sub

Private Sub RepFgPckLst_Click()
'*** Jay 3.4(FG)
  Call FormSelect("RepFgPckLst")
End Sub

Private Sub RepFgShpBill_Click()
'*** Jay 3.4(FG)
  Call FormSelect("RepFgShpBill")
End Sub

Private Sub RepFgStkLed_Click()
'*** Jay 3.4(FG)
  Call FormSelect("RepFgStkLed")
End Sub

Private Sub RepJtAvg_Click()
  ' **** Manali Trading Module
  Call FormSelect("RepJtAvg")
End Sub

Private Sub RepJTCat_Click()
  ' **** Manali Trading Module
  Call FormSelect("RepJtCat")
End Sub

Private Sub RepJtDet_Click()
  ' **** Manali Trading Module
  Call FormSelect("RepJtDet")
End Sub

Private Sub RepJtGrBk_Click()
  ' **** Manali Trading Module
  Call FormSelect("RepJtGrBk")
End Sub

Private Sub RepJtGrFrnt_Click()
  ' **** Manali Trading Module
  Call FormSelect("RepJtGrFrnt")
End Sub

Private Sub RepJtGsp_Click()
  Call FormSelect("RepJtGsp")
End Sub

Private Sub RepJtInv_Click()
  ' **** Manali Trading Module
  Call FormSelect("RepJtInv")
End Sub

Private Sub RepJtLclInv_Click()
  Call FormSelect("RepJtLclInv")
End Sub

Private Sub RepJtNetStk_Click()
  ' **** Manali Trading Module
  Call FormSelect("RepJtNetStk")
End Sub

Private Sub RepJtOpnOrd_Click()
  ' **** Manali Trading Module
  Call FormSelect("RepJtOpnOrd")
End Sub

Private Sub RepJtPckLst_Click()
  ' **** Manali Trading Module
  Call FormSelect("RepJtPckLst")
End Sub

Private Sub RepJtPnd_Click()
  ' **** Manali Trading Module
  Call FormSelect("RepJtPnd")
End Sub

Private Sub RepJtShpBill_Click()
  ' **** Manali Trading Module
  Call FormSelect("RepJtShpBill")
End Sub

Private Sub RepJTStkLed_Click()
  ' **** Manali Trading Module
  Call FormSelect("RepJtStkLed")
End Sub

Private Sub RepMetLsFg_Click()
  If WithinRTCRange = True Then
    Call FormSelect("RepMetLsFg")
  End If
End Sub

Private Sub RepPrm_Click()
  Call FormSelect("RepPrm")
End Sub

Private Sub RepRfIdMst_Click()
  ' ***** Manali 3.9.1 - 23/02/12 - RfId Module
  Call FormSelect("RepRfIdMst")
End Sub

' ****** Manali 3.3.0
Private Sub RepRmIdSz_Click()
  Call FormSelect("RepRmIdSz")
End Sub

Private Sub RepRtCmp_Click()
  Call FormSelect("RepRtCmp")
End Sub
Private Sub RepTxnBM_Click()
  Call FormSelect("RepTxnBM")
End Sub
'*** Jay 2.14 P2 ***
Private Sub RepDscRep_Click()
  Call FormSelect("RepDscRep")
End Sub
'*** Jay 3.01 *** [CustmPur Lst]
Private Sub RepTxnCustmPurLst_Click()
  Call FormSelect("RepTxnCustmPurLst")
End Sub
Private Sub RepVarPrm_Click()
  Call FormSelect("RepVarPrm")
End Sub
Private Sub RepCustLst_Click()
  Call FormSelect("RepCustLst")
End Sub
 Private Sub RepUsrAccess_Click()
  Call FormSelect("RepUsrAccess")
End Sub
Private Sub RepCustDsg_Click()
  Call FormSelect("RepCustDsg")
End Sub
Private Sub RepCustAly_Click()
  Call FormSelect("RepCustAly")
End Sub
Private Sub RepRmMst_Click()
  Call FormSelect("RepRmMst")
End Sub
Private Sub RepDsg_Click()
  Call FormSelect("RepDsg")
End Sub
Private Sub RepDsgCrd_Click()
  Call FormSelect("RepDsgCrd")
End Sub
Private Sub RepDsgCat_Click()
  Call FormSelect("RepDsgCat")
End Sub

'*** (Jen 3.01)
Private Sub DsgView_Click()
  Call FormSelect("DsgView")
End Sub
'*** (Jen 3.01)

Private Sub RepRmRt_Click()
  Call FormSelect("RepRmRt")
End Sub
Private Sub RepLabRt_Click()
  Call FormSelect("RepLabRt")
End Sub
Private Sub RepLoc_Click()
  Call FormSelect("RepLoc")
End Sub
Private Sub RepDtTbl_Click()
  Call FormSelect("RepDtTbl")
End Sub
Private Sub RepBnk_Click()
  Call FormSelect("RepBnk")
End Sub
Private Sub RepBagLst_Click()
  Call FormSelect("RepBagLst")
End Sub
' ###################################  Manoj #### Ver: 2.0.6  ###################################
' Date: 25/11/2003
' Change: Added for Production Points Master Listing
Private Sub RepPrdPts_Click()
  Call FormSelect("RepPrdPts")
End Sub
' ###########################################  Manoj  ###########################################
Private Sub RepLossMst_Click()
  Call FormSelect("RepLossMst")
End Sub
Private Sub RepTxnDT_Click()
  Call FormSelect("RepTxnDT")
End Sub
Private Sub RepTxnBS_Click()
  Call FormSelect("RepTxnBS")
End Sub
Private Sub RepTxnMLT_Click()
  Call FormSelect("RepTxnMLT")
End Sub
Private Sub RepTxnBV_Click()
  Call FormSelect("RepTxnBV")
End Sub
Private Sub RepTxnREJ_Click()
  Call FormSelect("RepTxnREJ")
End Sub
Private Sub RepTxnPR_Click()
  Call FormSelect("RepTxnPR")
End Sub
Private Sub RepTxnMV_Click()
  Call FormSelect("RepTxnMV")
End Sub
Private Sub RepTxnCNV_Click()
  Call FormSelect("RepTxnCNV")
End Sub
Private Sub RepUnMtchCNV_Click()
  Call FormSelect("RepUnMtchCNV")
End Sub
'Private Sub RepTxnREC_Click()
'  'Dim wfrm As New EmrFrpTxnLst
'  '  ShowForm me,  wfrm, "RepTxnREC"
'  '  wfrm.Caption = "Recovery Transaction Listing"
'  Call FormSelect("RepTxnREC")
'End Sub
Private Sub RepFbLst_Click()
  Call FormSelect("RepFbLst")
End Sub
Private Sub RepFwbLst_Click()
  Call FormSelect("RepFwbLst")
End Sub
Private Sub RepFgMvm_Click()
  Call FormSelect("RepFgMvm")
End Sub
Private Sub RepVchPrnBM_Click()
  Call FormSelect("RepVchPrnBM")
End Sub
Private Sub RepVchPrnCustmPur_Click()
  Call FormSelect("RepVchPrnCustmPur")
End Sub
Private Sub RepVchPrnDT_Click()
  Call FormSelect("RepVchPrnDT")
End Sub
Private Sub RepVchPrnBS_Click()
  Call FormSelect("RepVchPrnBS")
End Sub
Private Sub RepVchPrnMLT_Click()
  Call FormSelect("RepVchPrnMLT")
End Sub
Private Sub RepVchPrnBV_Click()
  Call FormSelect("RepVchPrnBV")
End Sub
Private Sub RepVchPrnREJ_Click()
  Call FormSelect("RepVchPrnREJ")
End Sub
Private Sub RepVchPrnPR_Click()
  Call FormSelect("RepVchPrnPR")
End Sub
Private Sub RepVchPrnMV_Click()
  Call FormSelect("RepVchPrnMV")
End Sub
Private Sub RepVchPrnCNV_Click()
  Call FormSelect("RepVchPrnCNV")
End Sub
Private Sub RepVchPrnREC_Click()
  Call FormSelect("RepVchPrnREC")
End Sub
Private Sub RepLsRec_Click()
  Call FormSelect("RepLsRec")
End Sub
Private Sub RepCustRm_Click()
  Call FormSelect("RepCustRm")
End Sub
Private Sub RepTreePerform_Click()
  Call FormSelect("RepTreePerform")
End Sub
Private Sub RepTreeInvst_Click()
'  Dim wfrm As New EmrFrpTreeLst
'  ShowForm Me, wfrm, "RepTreeInvst"
'  wfrm.Caption = "Tree Investment Report"
  '***************Geeta*****************
  Call FormSelect("RepTreeInvst")
End Sub
'*** (Jen 2.14) shifted from MDI Form
Private Sub RepDetInvCTB_Click()
  Call FormSelect("RepDetInvCTB")
End Sub
Private Sub RepAvgInvCTB_Click()
  Call FormSelect("RepAvgInvCTB")
End Sub
Private Sub RepInvCTBCat_Click()
  Call FormSelect("RepInvCTBCat")
End Sub
Private Sub RepCTBPckLst_Click()
  Call FormSelect("RepCTBPckLst")
End Sub
Private Sub RepDsgLd_Click()
  Call FormSelect("RepDsgLd")
End Sub
Private Sub RepCntrlCpy_Click()
  Call FormSelect("RepCntrlCpy")
End Sub
Private Sub RepOrdPln_Click()
  Call FormSelect("RepOrdPln")
End Sub
Private Sub EmrToXML_Click()
  Call FormSelect("EmrToXML")
End Sub
Private Sub EmrUpd_Click()
  Call FormSelect("EmrUpd")
End Sub

Private Sub RprBV_Click()

End Sub

Private Sub RprDT_Click()
 
End Sub

Private Sub Rptt_Click()

End Sub

Private Sub RprHis_Click()
  Call FormSelect("RprHis")
End Sub

Private Sub RprWIP_Click()
  Call FormSelect("RprWIP")
End Sub

Private Sub RprWIPBv_Click()
  Call FormSelect("RprWIPBv")
End Sub

Private Sub TagFmt_Click()
  Call FormSelect("TagFmt")
End Sub
'*** (Jen 2.14) shifted from MDI Form

'*** (Jen 2.14 Next)
Private Sub RepTxnRetMem_Click()
  Call FormSelect("RepTxnRetMem")
End Sub
Private Sub RepVchPrnRetMem_Click()
  Call FormSelect("RepVchPrnRetMem")
End Sub
'*** (Jen 2.14 Next)

'*****************  Std Code not to  be changed ****************************************'
Private Sub Form_Load()
  MWLib.BeginProcess Me, "Loading " + Me.Caption
  Dim wo_rsUsr As MwfLib.MDORowSet
  If Not ADC.InitRtn Then GoTo LstEnd
  On Error GoTo ErrH
  
  '*** Jenny Colour
  SetMDISysColorsAfterLogin   '*** Jenny Color (For Setting System Colors)
  
  Dim ws_BtnStr As String
  ws_BtnStr = ""
  Call SetObjColors(Me, ws_BtnStr)
  '*** Jenny Colour
  
  Call HideAllMnus

  '*** Jay 3.4(FG) *** [hFgIOYn]
  ' **** Manali 3.8.0 - CTB Menus shifted to Lst From MDI, HCtbMod added
  ' **** Manali 3.9.1 - RfId Module
  Dim wo_RsModules As MwfLib.MDORowSet, ws_Mod As String
  '6.1-15 HPPCYN added
  Set wo_RsModules = ADC.Connection.OpenRes("select HCoCd, HLclInvYn, hFgIOYn, HCtbMod, HRfIdMod,HTravelExlYN,HPPCYN from Head " + _
                     "where HCoCd ='" + UCase(gs_CoCd) + "' and HCd= '" + ctSelfCmCd + "'")
  If UCase(wo_RsModules.FldValue("hFgIOYn")) = "N" Then ws_Mod = ws_Mod + "'MnuFGTxn'"
  If UCase(wo_RsModules.FldValue("HLclInvYn")) = "N" Then ws_Mod = ws_Mod + IIF(ws_Mod <> "", ",", "") + "'RepFgLclInv'"
  If UCase(wo_RsModules.FldValue("HCtbMod")) = "N" Then ws_Mod = ws_Mod + IIF(ws_Mod <> "", ",", "") + "'CTB'"
  If UCase(wo_RsModules.FldValue("HRfIdMod")) = "N" Then ws_Mod = ws_Mod + IIF(ws_Mod <> "", ",", "") + "'RepRfIdMst'"
  'mw.104- Travel excel report moved from MDi to AddOns Menu. If Travel excel module is enabled then this report will display
  'otherwise report will not show in the menu list
  If UCase(wo_RsModules.FldValue("HTravelExlYN")) = "N" Then ws_Mod = ws_Mod + IIF(ws_Mod <> "", ",", "") + "'RepTravExcel'"
  '6.1-15 ppc menus will not display if PPC modelu is NO in head
  ' it is removed from mdi
  If UCase(wo_RsModules.FldValue("HPPCYN")) = "N" Then ws_Mod = ws_Mod + IIF(ws_Mod <> "", ",", "") + "'PPC','PPCCAP','PPCDSGLD','PPCORDPLN','REPPPC','REPDSGLD','REPCNTRLCPY','REPORDPLN'"
  
  ws_Mod = IIF(ws_Mod <> "", " and UaMnuCd Not In (" + ws_Mod + ")", "")
  Set wo_RsModules = Nothing
  '*** Jay 3.4(FG) *** [hFgIOYn]

  WindowState = 0
  Dim wo_rsusracc As MwfLib.MDORowSet
  
  '*** Jay 3.4(FG) [+ws_Mod added]
  Set wo_rsusracc = ADC.Connection.OpenRes("select * from usraccess where uausrcd='" + gs_UsrCd + "' And UaMnuForm ='" + ms_MnuForm + "'" + _
                      " And ',' + UaValidCoCd + ',' Like Case When (UaValidCoCd = '') Then ',' + UaValidCoCd + ',' Else '%," & gs_CoCd & ",%' End" + ws_Mod)
  '*** Jay 3.4(FG) [+ws_Mod added]
  Do While Not (wo_rsusracc.EOF Or wo_rsusracc.BOF)
    'Debug.Print wo_rsusracc.FldValue("uamnucd")
    Controls(wo_rsusracc.FldValue("uamnucd")).Visible = True
    'Controls(wo_rsusracc.FldValue("uamnucd")).Tag = IIF(IsNull(wo_rsusracc.FldValue("uaadcopt")), "", wo_rsusracc.FldValue("uaadcopt"))
MNext:
    wo_rsusracc.MoveNext
  Loop
  ' ***** Manali 3.8.0
  If UCase(GetMDIFrm.Name) = UCase("EmrFrmJMDI") Then
    MnuLstDailyRep.Visible = False
    MnuFGTxn.Visible = False
    MnuLstTxn.Visible = False
    MnuLstVchr.Visible = False
  End If
  ' ***** Manali 3.8.0
  WINDOW.Visible = True
  CSC.Visible = True
  TIL.Visible = True
  GENMNU.Visible = True
  MnuTlBr.Visible = True     '*** (Jenny Color)
  Call HideDumMnus
  WINDOW.Visible = True

  SendKeys "%"
' To Get Tag & Code of Menu
'  Call PrnMnus
LstEnd:
  MWLib.EndProcess Me
  Exit Sub
ErrH:
  If Err.Number = 423 Then
    Resume MNext
  
  ' ***** Manali 3.8.0
  ElseIf Err.Number = 730 Then
    Resume MNext
  ' ***** Manali 3.8.0
  
  Else
    MsgBox Err.Description
    Resume
  End If
End Sub
Private Sub Form_GotFocus()
  '*** Jenny Color (For Setting System Colors)
  '(Bef 13/08/05)
  'SetMDISysColorsAfterLogin
  '*** Jenny Color (For Setting System Colors)
End Sub
Private Sub CmdCan_Click()
    Unload Me
End Sub
Private Sub csc_Click()
    GetMDIFrm.Arrange vbCascade   ' ***** Manali 3.8.0 - New MDI and Prj
    'EmrFrmMDI.Arrange vbCascade
End Sub

'*** (Bef 12/08/05)
'Private Sub ShowForm(ByVal Frm As Form, MenuCd As String)
'    Dim wf_ExecGot  As Boolean
'    On Error GoTo errh
'
'' ########################  Manoj #### Ver: 2.0.6 #### Date: 16/06/2004  ########################
'' Check Login Session for Working or Not
'    If Not ChkLogin(ADC.Connection) Then End
'' ###########################################  Manoj  ###########################################
'
'    MWLib.BeginProcess Me, "Loading " + Controls(MenuCd).Caption
'    wf_ExecGot = False
'    If Not (EmrFrmMDI.ActiveForm Is Nothing) Then
'        EmrFrmMDI.ActiveForm.ADC.ExecGot = True
'        wf_ExecGot = True
'    End If
'    Load Frm
'    takecloseoff Frm.hwnd
'    If wf_ExecGot Then Frm.ADC.ExecGot = True
'    If Frm.WindowState = 0 Then
'      Frm.Left = 0
'      Frm.Top = 0
'    End If
'    Frm.ADC.MenuCd = MenuCd
'    Frm.ADC.UsrCd = gs_UsrCd
'' ########################  Manoj #### Ver: 2.0.6 #### Date: 03/06/2004  ########################
'    Frm.ADC.HostNm = gs_HostNm
'  ' set Lising Form AS MenuForm for Loading Form
'    Set Frm.ADC.MnuForm = Me
'    If Not Frm.ADC.InitRtn Then End
'    ADC.Visible = False
'' ###########################################  Manoj  ###########################################
'    Set Frm = Nothing
'    MWLib.EndProcess Me
'    Exit Sub
'errh:
'    If Err.Number = 40002 Then
'        Resume Next
'    Else
'        MsgBox Err.Description
'    End If
'End Sub
'*** (Bef 12/08/05)

Private Sub til_Click()
    GetMDIFrm.Arrange vbTileVertical    ' ***** Manali 3.8.0 - New MDI and Prj
    'EmrFrmMDI.Arrange vbTileVertical
End Sub
Public Sub takecloseoff(handle As Long)
    Dim SysMenHandle As Long, RetVal As Long
    SysMenHandle = GetSystemMenu(handle, 0)
    RetVal = RemoveMenu(SysMenHandle, 6, MF_BYPOSITION)
    RetVal = RemoveMenu(SysMenHandle, 5, MF_BYPOSITION)
End Sub
Private Sub GenMnu_click()
Dim W As Variant
  On Error GoTo GenMnuErr
    Dim rs_UsrAcc As MwfLib.MDORowSet
    Dim wo_Ctl As Control
    Dim a As Menu
'And UCase$(wo_Ctl.Name) <> "RESAUTO"
    MWLib.BeginProcess Me, "Menu Generation Is In Progress"
    ADC.Connection.Execute ("DELETE FROM USRACCess WHERE UAUSRCD= '" + ctGenMenuUsr + "' And UAMNUFORM = '" + ms_MnuForm + "'")
    Set rs_UsrAcc = ADC.Connection.OpenResultset("select * from usraccess where uausrcd= '" + ctGenMenuUsr + "' And UAMNUFORM = '" + ms_MnuForm + "'")
    For Each wo_Ctl In Me.Controls
        If TypeOf wo_Ctl Is Menu Then
            If wo_Ctl.Caption <> "DUM" And wo_Ctl.Name <> "WINDOW" And _
              wo_Ctl.Name <> "CSC" And wo_Ctl.Name <> "TIL" And wo_Ctl.Name <> "GENMNU" And UCase$(wo_Ctl.Name) <> "MNUTLBR" Then
                'Debug.Print wo_Ctl.Tag
                rs_UsrAcc.Addnew
                
                rs_UsrAcc!UAUSRCD = ctGenMenuUsr
                rs_UsrAcc!UaMnuCd = UCase(Trim(wo_Ctl.Name))
                W = rs_UsrAcc!UaMnuCd
                rs_UsrAcc!UaMnuDesc = Trim(wo_Ctl.Caption)
                rs_UsrAcc!UaAdcOpt = ""
                rs_UsrAcc!UaMnuForm = ms_MnuForm
                rs_UsrAcc!UaValidCoCd = ""
                rs_UsrAcc!UaTag = Trim(wo_Ctl.Tag)
                If (rs_UsrAcc!UaTag = "9957") Then
                  Dim wdde As String
                  wdde = "e"
                End If
                
                rs_UsrAcc!UaTlBarDesc = ""
                rs_UsrAcc!ModUsr = ""
                rs_UsrAcc!ModDt = MWLib.EmptyDate
                rs_UsrAcc!ModTime = 0
                rs_UsrAcc!UaSeoYN = ""
                rs_UsrAcc!UaReplNo = 0      '*** (Jen 2.14 Next Patch 1)
                'rs_UsrAcc!UaMetlrCd = ""   '***** Manali 3.8.0 - metlr Cd
                rs_UsrAcc.Update
            End If
        End If
    Next
    ADC.Connection.Execute ("Update UsrAccess Set UaTag = IsNull((Select isNull(UA1.UaTag , '') From UsrAccess UA1 " + _
        "Where UA1.UaMnuCd= UsrAccess.UaMnuCd And UA1.UaMnuForm = UsrAccess.UaMnuForm  And " + _
        "UA1.UAUSRCD= '" + ctGenMenuUsr + "'), '') Where UAMNUFORM = '" + ms_MnuForm + "'" + _
        " And exists (Select  'a' from UsrAccess UA2 where ua2.uamnucd=usraccess.uamnucd and ua2.uamnuform=usraccess.uamnuform and ua2.uausrcd='" + ctGenMenuUsr + "')")
    DispMsg "Menu on '" + ms_MnuForm + "' From Succesfully Generated for '" + ctGenMenuUsr + "'", etInfo
GenMnuErr:
    If Err.Number <> 0 Then
      DispMsg Err.Description + " Menu Can Not Generate", etError
      If Err.Number = -2147217900 Then rs_UsrAcc.CancelUpdate
    End If
    MWLib.EndProcess Me
End Sub
Private Sub HideAllMnus()
  Dim Ctl As Control
    For Each Ctl In Controls
        If TypeOf Ctl Is Menu Then
           If Ctl.Caption <> "DUM" Then
                Ctl.Visible = False
           End If
        End If
    Next
End Sub
Private Sub HideDumMnus()
  Dim Ctl As Control
  On Error GoTo ErrH
    For Each Ctl In Controls
        If TypeOf Ctl Is Menu Then
           If Ctl.Caption = "DUM" Then
                Ctl.Visible = False
           End If
        End If
    Next
    Exit Sub
ErrH:
    Resume Next
End Sub
' ########################  Manoj #### Ver: 2.0.6 #### Date: 03/06/2004  ########################
' ###########################################  Manoj  ###########################################

' Procedure To Get Tag & Code of Menu
'Private Sub PrnMnus()
'    Dim Ctl As Control
'    For Each Ctl In Controls
'        If TypeOf Ctl Is Menu Then
'           If Ctl.Caption <> "DUM" Then
'                Debug.Print Ctl.Tag + vbTab + vbTab + vbTab + Ctl.Name
'           End If
'        End If
'    Next
'End Sub

Private Sub CTBIn_Click()
  Call FormSelect("CTBIn")
End Sub
Private Sub CTBInAllBags_Click()
  Call FormSelect("CTBInAllBags")
End Sub
Private Sub CTBInDsg_Click()
  Call FormSelect("CTBInDsg")
End Sub

' Zubin 250308 (Emr 3.02.01)
'mw.104-Travel excel moved from MDi to AddOns menu
Private Sub RepTravExcel_Click()
  Call FormSelect("RepTravExcel")
End Sub
'6.1_38
Private Sub RepInvLst_Click()
  Call FormSelect("RepInvLst")
End Sub

'6.1-15
Private Sub PpcCap_Click()
  Call FormSelect("PpcCap")
End Sub
Private Sub PpcDsgLd_Click()
  Call FormSelect("PpcDsgLd")
End Sub
Private Sub PpcOrdPln_Click()
  Call FormSelect("PpcOrdPln")
End Sub
'6.1-124
Private Sub RepOpnVch_Click()
  Call FormSelect("RepOpnVch")
End Sub

