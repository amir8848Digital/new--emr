VERSION 5.00
Begin {BD4B4E61-F7B8-11D0-964D-00A0C9273C2A} EmrRepOrd 
   ClientHeight    =   11340
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   20565
   OleObjectBlob   =   "EmrRepOrd.dsx":0000
End
Attribute VB_Name = "EmrRepOrd"
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



Private Sub PicSec_Format(ByVal pFormattingInfo As Object)
' **** Zubin 211 **** '
'''''  wPth = GetPictPath(FldDmTcTyp.Value, False, FldDmCd.Value, FldDmSz.Value)
'''''  If Dir(wPth) <> "" And wPth <> "" Then
'''''    Set picdesgn.FormattedPicture = LoadPicture(wPth)
'''''  Else
'''''    'picdesgn.Suppress = True
'''''    PicSec.Suppress = True
'''''  End If
  
  '*** 2.12 (Jen) (17/08/05) (This code can be used to maintain the exact proportion of the pictures)
  If mf_PicHt = 0 Then mf_PicHt = picdesgn.Height
  If mf_PicWdth = 0 Then mf_PicWdth = picdesgn.Width
  
  ' **** Manali 3.6.1
  If UCase(FldPicOpt.Value) = "3D" Then
    Call SetRptPic(mf_PicHt, mf_PicWdth, picdesgn, FldDmTcTyp.Value, False, True, FldDmCd.Value, FldDmSz.Value, , , , , , , , FldOdPicNm.Value)
   ElseIf UCase(FldPicOpt.Value) = "HSK" Then
    Call SetRptPic(mf_PicHt, mf_PicWdth, picdesgn, FldDmTcTyp.Value, True, True, FldDmCd.Value, FldDmSz.Value, , , , , , , , FldOdPicNm.Value)
  End If
  
  
  
  '*** Jay 3.2.0 [PicNm]
  ' bEF  3.6.1
'  Call SetRptPic(mf_PicHt, mf_PicWdth, picdesgn, FldDmTcTyp.Value, False, True, FldDmCd.Value, FldDmSz.Value, PICSEC, , , , , , , FldOdPicNm.Value)
  'Call SetRptPic(mf_PicHt, mf_PicWdth, picdesgn, FldDmTcTyp.Value, False, True, FldDmCd.Value, FldDmSz.Value, PICSEC)
  '*** Jay 3.2.0 [PicNm]
  '*** 2.12 (Jen) (17/08/05) (This code can be used to maintain the exact proportion of the pictures)
  
  '*** (Bef 2.12)
  'Call SetRptPic(picdesgn, FldDmTcTyp.Value, False, True, FldDmCd.Value, FldDmSz.Value, PICSEC)
  '*** (Bef 2.12)
' **** Zubin 211 **** '
End Sub


Private Sub RHLogo_Format(ByVal pFormattingInfo As Object)
  If mf_PicHtLogo = 0 Then mf_PicHtLogo = PicLogo.Height
  If mf_PicWdthLogo = 0 Then mf_PicWdthLogo = PicLogo.Width
  Call SetRptLogoPic(mf_PicHtLogo, mf_PicWdthLogo, PicLogo, gs_CoCd, RHLogo)
End Sub

