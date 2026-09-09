VERSION 5.00
Begin {BD4B4E61-F7B8-11D0-964D-00A0C9273C2A} EmrRepMultiPrcQt 
   ClientHeight    =   9450
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   14985
   OleObjectBlob   =   "EmrRepMultiPrcQt.dsx":0000
End
Attribute VB_Name = "EmrRepMultiPrcQt"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Dim mf_PicHt As Single
Dim mf_PicWdth As Single

Dim mf_PicHtLogo As Single
Dim mf_PicWdthLogo As Single

Private Sub GH1BlankSec_Format(ByVal pFormattingInfo As Object)
End Sub

Private Sub GH2BlankSec_Format(ByVal pFormattingInfo As Object)
If PICSEC.Suppress = False Then GH2BlankSec.Suppress = True

End Sub

Private Sub PicSec_Format(ByVal pFormattingInfo As Object)
  If mf_PicHt = 0 Then mf_PicHt = picdesgn.Height
  If mf_PicWdth = 0 Then mf_PicWdth = picdesgn.Width

  Call SetRptPic(mf_PicHt, mf_PicWdth, picdesgn, FldDmTcTyp.Value, False, False, FldDmCd.Value, FldDmSz.Value, PICSEC)

End Sub


Private Sub RHLogo_Format(ByVal pFormattingInfo As Object)
  If mf_PicHtLogo = 0 Then mf_PicHtLogo = PicLogo.Height
  If mf_PicWdthLogo = 0 Then mf_PicWdthLogo = PicLogo.Width
  Call SetRptLogoPic(mf_PicHtLogo, mf_PicWdthLogo, PicLogo, gs_CoCd, RHLogo)
End Sub

