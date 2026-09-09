VERSION 5.00
Begin {BD4B4E61-F7B8-11D0-964D-00A0C9273C2A} EmrRepLclInv 
   ClientHeight    =   11190
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   14430
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
     If DetDsg.Height >= 400 Then LineSec.Suppress = True Else LineSec.Suppress = False
   Else
    LineSec.Suppress = True
   End If
Else
  LineSec.Suppress = True
End If
End Sub

Private Sub Linesec1_Format(ByVal pFormattingInfo As Object)
If PicSec.Suppress = True Then
  If FldShowDemark.Value = "Y" Then
     If DetDsg.Height < 300 Then Linesec1.Suppress = True Else Linesec1.Suppress = False
 Else
    Linesec1.Suppress = True
   End If
Else
  Linesec1.Suppress = True
End If

End Sub

Private Sub PH1_Format(ByVal pFormattingInfo As Object)
  If mf_PicHt = 0 Then mf_PicHt = PicLogo1.Height
  If mf_PicWdth = 0 Then mf_PicWdth = PicLogo1.Width
  Call SetRptLogoPic(mf_PicHt, mf_PicWdth, PicLogo1, FldCmBillCo1.Value, PH1)
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
