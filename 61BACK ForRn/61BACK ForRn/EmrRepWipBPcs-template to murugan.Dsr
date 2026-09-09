VERSION 5.00
Begin {BD4B4E61-F7B8-11D0-964D-00A0C9273C2A} EmrRepWipBPcs 
   ClientHeight    =   9390
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   16590
   OleObjectBlob   =   "EmrRepWipBPcs.dsx":0000
End
Attribute VB_Name = "EmrRepWipBPcs"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'*** 2.12 (Jen) (17/08/05)
Dim mf_PicHt As Single
Dim mf_PicWdth As Single
'*** 2.12 (Jen) (17/08/05)
Dim mf_PicHtLogo As Single
Dim mf_PicWdthLogo As Single

Private Sub RHLogo_Format(ByVal pFormattingInfo As Object)
  If mf_PicHtLogo = 0 Then mf_PicHtLogo = PicLogo.Height
  If mf_PicWdthLogo = 0 Then mf_PicWdthLogo = PicLogo.Width
  Call SetRptLogoPic(mf_PicHtLogo, mf_PicWdthLogo, PicLogo, gs_CoCd, RHLogo)
End Sub


