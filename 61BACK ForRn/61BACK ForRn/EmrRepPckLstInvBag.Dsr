VERSION 5.00
Begin {BD4B4E61-F7B8-11D0-964D-00A0C9273C2A} EmrRepPckLstInvBag 
   ClientHeight    =   9450
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   15645
   OleObjectBlob   =   "EmrRepPckLstInvBag.dsx":0000
End
Attribute VB_Name = "EmrRepPckLstInvBag"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' **** Zubin 211 **** '
'''''Dim wPth As String
' **** Zubin 211 **** '

'*** 2.12 (Jen) (17/08/05)
Dim mf_PicHt As Single
Dim mf_PicWdth As Single
'*** 2.12 (Jen) (17/08/05)


Private Sub PicSec_Format(ByVal pFormattingInfo As Object)
' **** Zubin 211 **** '
'''''  If PicSec.Suppress = True Then Exit Sub
'''''  wPth = GetPictPath("DM", False, FldDmCd.Value, FldDmSz.Value)
'''''  If Dir(wPth) <> "" And wPth <> "" Then
'''''    Set PicDsg.FormattedPicture = LoadPicture(wPth)
'''''  Else
'''''    'PicDsg.Suppress = True
'''''    PicSec.Suppress = True
'''''  End If

  '*** 2.12 (Jen) (17/08/05) (This code can be used to maintain the exact proportion of the pictures)
  If mf_PicHt = 0 Then mf_PicHt = PicDsg.Height
  If mf_PicWdth = 0 Then mf_PicWdth = PicDsg.Width
  '*** Jay 3.2.0 [OdPicNm]
  Call SetRptPic(mf_PicHt, mf_PicWdth, PicDsg, "DM", False, True, FldDmCd.Value, FldDmSz.Value, PICSEC, , , , , , , FldOdPicNm.Value)
  'Call SetRptPic(mf_PicHt, mf_PicWdth, PicDsg, "DM", False, True, FldDmCd.Value, FldDmSz.Value, PicSec)
  '*** Jay 3.2.0 [OdPicNm]
  '*** 2.12 (Jen) (17/08/05) (This code can be used to maintain the exact proportion of the pictures)


  '*** (Bef 2.12)
  'Call SetRptPic(PicDsg, "DM", False, True, FldDmCd.Value, FldDmSz.Value, PicSec)
  '*** (Bef 2.12)
' **** Zubin 211 **** '
End Sub


Private Sub PicSec1_Format(ByVal pFormattingInfo As Object)
  If mf_PicHt = 0 Then mf_PicHt = PicDsg.Height
  If mf_PicWdth = 0 Then mf_PicWdth = PicDsg.Width

  Call SetRptPic(mf_PicHt, mf_PicWdth, PicDsg1, "DM", False, True, FldDmCd.Value, FldDmSz.Value, PicSec1, , , , , , , FldOdPicNm.Value)

End Sub

