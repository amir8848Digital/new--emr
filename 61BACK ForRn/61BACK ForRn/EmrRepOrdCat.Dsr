VERSION 5.00
Begin {BD4B4E61-F7B8-11D0-964D-00A0C9273C2A} EmrRepOrdCat 
   ClientHeight    =   14130
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   22155
   OleObjectBlob   =   "EmrRepOrdCat.dsx":0000
End
Attribute VB_Name = "EmrRepOrdCat"
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

Private Sub DET_Format(ByVal pFormattingInfo As Object)
    ' **** Zubin 211 **** '
    '''''  'wPth = GetPictPath(FldDmTcTyp.Value, False, FldDmCd.Value, FldDmSz.Value)
    '''''  If Dir(wPth) <> "" And wPth <> "" Then
    '''''    Set picdesgn.FormattedPicture = LoadPicture(wPth)
    '''''  'Else
    '''''    'picdesgn.Suppress = True
    '''''  End If
  
  '*** 2.12 (Jen) (17/08/05) (This code can be used to maintain the exact proportion of the pictures)
  If mf_PicHt = 0 Then mf_PicHt = picdesgn.Height
  If mf_PicWdth = 0 Then mf_PicWdth = picdesgn.Width
  'Call SetRptPic(mf_PicHt, mf_PicWdth, picdesgn, FldDmTcTyp.Value, False, True, FldDmCd.Value, FldDmSz.Value, DET, , , , , , , FldOdPicNm.Value)
  Call SetRptPic(mf_PicHt, mf_PicWdth, picdesgn, FldDmTcTyp.Value, False, True, FldDmCd.Value, FldDmSz.Value, , , , , , , , FldOdPicNm.Value)
  Call SetRptPic(mf_PicHt, mf_PicWdth, picdesgn2, FldDmTcTyp2.Value, False, True, FldDmCd2.Value, FldDmSz2.Value, , , , , , , , FldOdPicNm2.Value)
  Call SetRptPic(mf_PicHt, mf_PicWdth, picdesgn3, FldDmTcTyp3.Value, False, True, FldDmCd3.Value, FldDmSz3.Value, , , , , , , , FldOdPicNm3.Value)
  '*** 2.12 (Jen) (17/08/05) (This code can be used to maintain the exact proportion of the pictures)
  
  '*** (Bef 2.12)
  'Call SetRptPic(picdesgn, FldDmTcTyp.Value, False, True, FldDmCd.Value, FldDmSz.Value)
  '*** (Bef 2.12)
' **** Zubin 211 **** '
End Sub

Private Sub RHLogo_Format(ByVal pFormattingInfo As Object)
'  If mf_PicHtLogo = 0 Then mf_PicHtLogo = PicLogo.Height
'  If mf_PicWdthLogo = 0 Then mf_PicWdthLogo = PicLogo.Width
'  Call SetRptLogoPic(mf_PicHtLogo, mf_PicWdthLogo, PicLogo, gs_CoCd, RHLogo)
End Sub


