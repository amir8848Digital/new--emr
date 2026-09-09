VERSION 5.00
Begin {BD4B4E61-F7B8-11D0-964D-00A0C9273C2A} EmrRepBagHist 
   ClientHeight    =   9885
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   15495
   OleObjectBlob   =   "EmrRepBagHist.dsx":0000
End
Attribute VB_Name = "EmrRepBagHist"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Dim mf_PicHt As Single
Dim mf_PicWdth As Single

Private Sub DET_Format(ByVal pFormattingInfo As Object)

End Sub

Private Sub Section2_Format(ByVal pFormattingInfo As Object)
  If mf_PicHt = 0 Then mf_PicHt = Pic.Height
  If mf_PicWdth = 0 Then mf_PicWdth = Pic.Width
  If UCase(FldPicOpt.Value) = "3D" Then
    Call SetRptPic(mf_PicHt, mf_PicWdth, Pic, "DM", False, True, FldDmCd.Value, FldDmSz.Value, , , , , , , , FldOdPicNm.Value)
  ElseIf UCase(FldPicOpt.Value) = "HSK" Then
    Call SetRptPic(mf_PicHt, mf_PicWdth, Pic, "DM", True, True, FldDmCd.Value, FldDmSz.Value, , , , , , , , FldOdPicNm.Value)
  End If

End Sub
