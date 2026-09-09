VERSION 5.00
Begin {BD4B4E61-F7B8-11D0-964D-00A0C9273C2A} EmrRepLclInv 
   ClientHeight    =   9450
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   15000
   OleObjectBlob   =   "EmrRepLclInv.dsx":0000
End
Attribute VB_Name = "EmrRepLclInv"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Dim mf_PicHt As Single
Dim mf_PicWdth As Single
Dim mf_PicDsgHt As Single
Dim mf_PicDsgWdth As Single
Dim mf_PicQRCodeHt As Single, mf_PicQRCodeWdth As Single

Private Sub GH2_Format(ByVal pFormattingInfo As Object)
'  If mf_PicHt = 0 Then mf_PicHt = PicLogo.Height
'  If mf_PicWdth = 0 Then mf_PicWdth = PicLogo.Width
'  Call SetRptLogoPic(mf_PicHt, mf_PicWdth, PicLogo, FldCmBillCo.Value, GH2)
End Sub

Private Sub GroupFooterSection14_Format(ByVal pFormattingInfo As Object)

End Sub

Private Sub GFAmtWord_Format(ByVal pFormattingInfo As Object)
' **** Manali Trading Module
TxtAmtInWrds.SetText FnNumToWords(FldInTotRs.Value, FldCurCd.Value)
End Sub

Private Sub GroupFooterSection7_Format(ByVal pFormattingInfo As Object)
  If mf_PicDTCHt = 0 Then mf_PicDTCHt = PicDTCLogo.Height
  If mf_PicDTCWdth = 0 Then mf_PicDTCWdth = PicDTCLogo.Width
  Call SetRptLogoPic(mf_PicDTCHt, mf_PicDTCWdth, PicDTCLogo, gs_CoCd, GroupFooterSection7, , True)
End Sub

Private Sub LinePicSec_Format(ByVal pFormattingInfo As Object)
If PicSec.Suppress = True Then
  LinePicSec.Suppress = True
Else
  If FldShowDemark.Value = "Y" Then LinePicSec.Suppress = False Else LinePicSec.Suppress = True
End If
End Sub

Private Sub LineSec_Format(ByVal pFormattingInfo As Object)
If PicSec.Suppress = True Then
  If FldShowDemark.Value = "Y" Then
    If DetDsg.Height >= 400 Or DetBagStkId.Suppress = False Then LineSec.Suppress = True Else LineSec.Suppress = False
  Else
    LineSec.Suppress = True
  End If
Else
  LineSec.Suppress = True
End If
End Sub

Private Sub Linesec1_Format(ByVal pFormattingInfo As Object)
If PicSec.Suppress = True Then
  If FldShowDemark.Value = "Y" Then   '62 DetDsg.Height >= 600 added
    If DetDsg.Height < 300 Or DetDsg.Height >= 580 Or DetBagStkId.Suppress = False Then Linesec1.Suppress = True Else Linesec1.Suppress = False
  Else
    Linesec1.Suppress = True
  End If
Else
  Linesec1.Suppress = True
End If
End Sub
'62 Linesec2_Format created. when show HUID is yes then line in this section will display
'38 DetBagStkId.Suppress = False added. Line in the DetBagStkId section will display
Private Sub Linesec2_Format(ByVal pFormattingInfo As Object)
If PicSec.Suppress = True Then
  If FldShowDemark.Value = "Y" Then
    If DetDsg.Height < 580 Or DetBagStkId.Suppress = False Then Linesec2.Suppress = True Else Linesec2.Suppress = False
  Else
    Linesec2.Suppress = True
  End If
Else
  Linesec2.Suppress = True
End If
End Sub
'38 LineSecBagStkId_Format created. when show Bag/stkid is yes then line in this section will display
Private Sub LineSecBagStkId_Format(ByVal pFormattingInfo As Object)
If PicSec.Suppress = True Then
  If FldShowDemark.Value = "Y" Then
    If DetBagStkId.Suppress = True Then LineSecBagStkId.Suppress = True Else LineSecBagStkId.Suppress = False
  Else
    LineSecBagStkId.Suppress = True
  End If
Else
  LineSecBagStkId.Suppress = True
End If

End Sub

'sv.63 showing QR code
Private Sub PageHeaderSection1_Format(ByVal pFormattingInfo As Object)
  If mf_PicQRCodeHt = 0 Then mf_PicQRCodeHt = PicQRCode.Height
  If mf_PicQRCodeWdth = 0 Then mf_PicQRCodeWdth = PicQRCode.Width
  Call SetRptQRCOde(mf_PicQRCodeHt, mf_PicQRCodeWdth, PicQRCode, gs_CoCd, FldInTc.Value, FldInYy.Value, FldInChr.Value, FldInNo.Value)
  
End Sub

Private Sub PH1_Format(ByVal pFormattingInfo As Object)
  If mf_PicHt = 0 Then mf_PicHt = PicLogo1.Height
  If mf_PicWdth = 0 Then mf_PicWdth = PicLogo1.Width
  Call SetRptLogoPic(mf_PicHt, mf_PicWdth, PicLogo1, gs_CoCd, PH1)
End Sub

Private Sub PicSec_Format(ByVal pFormattingInfo As Object)
  If FldShowPic.Value = "Y" Then
    If mf_PicDsgHt = 0 Then mf_PicDsgHt = PicDsg.Height
    If mf_PicDsgWdth = 0 Then mf_PicDsgWdth = PicDsg.Width
    Call SetRptPic(mf_PicDsgHt, mf_PicDsgWdth, PicDsg, "DM", False, True, FldDmCd.Value, FldDmSz.Value, PicSec, , , , , , , FldOdPicNm.Value)
  Else
    PicSec.Suppress = True
  End If
End Sub
