VERSION 5.00
Begin {BD4B4E61-F7B8-11D0-964D-00A0C9273C2A} EmrRepValAddnLocal 
   ClientHeight    =   9630
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   19140
   OleObjectBlob   =   "EmrRepValAddnLocal.dsx":0000
End
Attribute VB_Name = "EmrRepValAddnLocal"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False


Private Sub PageHeaderSection1_Format(ByVal pFormattingInfo As Object)
End Sub

Private Sub GroupHeaderSection1_Format(ByVal pFormattingInfo As Object)
End Sub

Private Sub GH2Blank_Format(ByVal pFormattingInfo As Object)
'  Dim wRowCnt As Integer
'  wRowCnt = 0
'  If Fldg2GWt.Value > 0 Then wRowCnt = wRowCnt + 1
'  If Fldg2PWt.Value > 0 Then wRowCnt = wRowCnt + 1
'  If Fldg2SWt.Value > 0 Then wRowCnt = wRowCnt + 1
'  If Fldg2LWt.Value > 0 Then wRowCnt = wRowCnt + 1
'  If Fldg2GWt.Value > 0 Then wRowCnt = wRowCnt + 1
  'g
  
  
End Sub

Private Sub G2HBlank_Format(ByVal pFormattingInfo As Object)
  Dim wMetSecHt As Single
  wMetSecHt = IIF(G2HGld.Suppress, G2HGld.Height, 0) + IIF(G2HPlt.Suppress, G2HPlt.Height, 0) + IIF(G2HSlv.Suppress, G2HSlv.Height, 0) + IIF(G2HPd.Suppress, G2HPd.Height, 0) + IIF(G2HM.Suppress, G2HM.Height, 0)
  
  G2HBlank.Height = IIF(wMetSecHt > SubRep2.Height, wMetSecHt, G2HStudSubrep.Height) ' + 100
End Sub

'*** Original FOrmula For wGoldBOE
'WhilePrintingRecords;
'stringVar a;
'a := a+
'(If {rdo.RmCtg}= 'G' AND {rdo.qIrSrNo}<> 0 AND
'    ((a<> '' and InStr (a+'  /  ', '  /  '+Trim({rdo.TBOENo})+' Dated: '+Trim(ToText ({rdo.TBOEDt}, 'dd-MM-yy'))+'  /  ')= .00) or a= '')
'Then
'    '  /  '+Trim({rdo.TBOENo}) +' Dated: '+ Trim(ToText ({rdo.TBOEDt}, 'dd-MM-yy'))
'Else '')
'*** Original FOrmula For wGoldBOE


'*** Original Formula For wPlBOE
'WhilePrintingRecords;
'stringVar b;
'b := b+
'(If {rdo.RmCtg}= 'P' AND {rdo.qIrSrNo}<> 0 AND
'    ((b<> '' and InStr (b+'  /  ', '  /  '+Trim({rdo.TBOENo})+' Dated: '+Trim(ToText ({rdo.TBOEDt}, 'dd-MM-yy'))+'  /  ')= .00) or b= '')
'Then
'    '  /  '+Trim({rdo.TBOENo}) +' Dated: '+ Trim(ToText ({rdo.TBOEDt}, 'dd-MM-yy'))
'Else '')
'*** Original Formula For wPlBOE
Private Sub Subreport1_DetailSection1_Format(ByVal pFormattingInfo As Object)

End Sub

