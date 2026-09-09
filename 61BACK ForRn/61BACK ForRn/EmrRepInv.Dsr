VERSION 5.00
Begin {BD4B4E61-F7B8-11D0-964D-00A0C9273C2A} EmrRepInv 
   ClientHeight    =   14130
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   22320
   OleObjectBlob   =   "EmrRepInv.dsx":0000
End
Attribute VB_Name = "EmrRepInv"
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
'new
Private Sub LinePicSec_Format(ByVal pFormattingInfo As Object)
  If PICSEC.Suppress = True Then
    LinePicSec.Suppress = True
  Else
    If wShowDemark1.Value = "Y" Then LinePicSec.Suppress = False Else LinePicSec.Suppress = True
  End If
End Sub

Private Sub LineSec_Format(ByVal pFormattingInfo As Object)
  'new
  If PICSEC.Suppress = True Then
    If wShowDemark1.Value = "Y" Then LineSec.Suppress = False Else LineSec.Suppress = True
  Else
    LineSec.Suppress = True
  End If
End Sub

Private Sub PH_Format(ByVal pFormattingInfo As Object)
  If mf_PicLogoHt = 0 Then mf_PicLogoHt = PicLogo.Height
  If mf_PicLogoWdth = 0 Then mf_PicLogoWdth = PicLogo.Width
  Call SetRptLogoPic(mf_PicLogoHt, mf_PicLogoWdth, PicLogo, gs_CoCd, PH)
End Sub

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

Private Sub SecFooter_Format(ByVal pFormattingInfo As Object)
  If mf_PicDTCHt = 0 Then mf_PicDTCHt = PicDTCLogo.Height
  If mf_PicDTCWdth = 0 Then mf_PicDTCWdth = PicDTCLogo.Width
  Call SetRptLogoPic(mf_PicDTCHt, mf_PicDTCWdth, PicDTCLogo, gs_CoCd, SecFooter, , True)
End Sub

