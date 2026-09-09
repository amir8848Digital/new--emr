VERSION 5.00
Begin {BD4B4E61-F7B8-11D0-964D-00A0C9273C2A} EmrRepBagPrn4 
   ClientHeight    =   9450
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   16635
   OleObjectBlob   =   "EmrRepBagPrn4.dsx":0000
End
Attribute VB_Name = "EmrRepBagPrn4"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'*** 2.12 (Jen) (17/08/05)
Dim mf_PicHt As Single
Dim mf_PicWdth As Single
'*** 2.12 (Jen) (17/08/05)


Dim mf_PicHtSub As Single
Dim mf_PicWdthSub As Single

Private Sub PicSec_Format(ByVal pFormattingInfo As Object)
' **** Zubin 211 **** '
'  If UCase(FldPicOpt.Value) = "3D" Then
'    wPth = GetPictPath("DM", False, FldDmCd.Value, FldDmSz.Value)
'  ElseIf UCase(FldPicOpt.Value) = "HSK" Then
'    wPth = GetPictPath("DM", True, FldDmCd.Value, FldDmSz.Value)
'  End If
'
'  If Dir(wPth) <> "" And wPth <> "" Then
'    Set Pic.FormattedPicture = LoadPicture(wPth)
'  Else
'    'Pic.Suppress = True
'    PicSec.Suppress = True
'  End If
  
  '*** 2.12 (Jen) (17/08/05) (This code can be used to maintain the exact proportion of the pictures)
  If mf_PicHt = 0 Then mf_PicHt = Pic.Height
  If mf_PicWdth = 0 Then mf_PicWdth = Pic.Width
  If UCase(FldPicOpt.Value) = "3D" Then
    '*** Jay 3.2.0 [PicNm]
    Call SetRptPic(mf_PicHt, mf_PicWdth, Pic, "DM", False, True, flddmcd.Value, flddmsz.Value, , , , , , , , FldOdPicNm.Value)
    '*** Jay 3.2.0 [PicNm]
  ElseIf UCase(FldPicOpt.Value) = "HSK" Then
    '*** Jay 3.2.0 [PicNm]
    Call SetRptPic(mf_PicHt, mf_PicWdth, Pic, "DM", True, True, flddmcd.Value, flddmsz.Value, , , , , , , , FldOdPicNm.Value)
    'Call SetRptPic(mf_PicHt, mf_PicWdth, Pic, "DM", True, True, FldDmCd.Value, FldDmSz.Value)
    '*** Jay 3.2.0 [PicNm]
  End If
  '*** 2.12 (Jen) (17/08/05) (This code can be used to maintain the exact proportion of the pictures)

  '*** (Bef 2.12)
  'If UCase(FldPicOpt.Value) = "3D" Then
  '  Call SetRptPic(Pic, "DM", False, True, FldDmCd.Value, FldDmSz.Value)
  'ElseIf UCase(FldPicOpt.Value) = "HSK" Then
  '  Call SetRptPic(Pic, "DM", True, True, FldDmCd.Value, FldDmSz.Value)
  'End If
  '*** (Bef 2.12)
' **** Zubin 211 **** '
End Sub

Private Sub Subreport2_RFPicSecHD_Format(ByVal pFormattingInfo As Object)

 If mf_PicHtSub = 0 Then mf_PicHtSub = Subreport2_PicSub.Height
  If mf_PicWdthSub = 0 Then mf_PicWdthSub = Subreport2_PicSub.Width
  If UCase(Subreport2_FldPicOpt.Value) = "3D" Then
    Call SetRptPic(mf_PicHtSub, mf_PicWdthSub, Subreport2_PicSub, "DM", False, True, Subreport2_FldDmCd.Value, Subreport2_FldDmSz.Value, , , , , , , , Subreport2_FldOdPicNm.Value)
  ElseIf UCase(Subreport2_FldPicOpt.Value) = "HSK" Then
    Call SetRptPic(mf_PicHtSub, mf_PicWdthSub, Subreport2_PicSub, "DM", True, True, Subreport2_FldDmCd.Value, Subreport2_FldDmSz.Value, , , , , , , , Subreport2_FldOdPicNm.Value)
  End If

End Sub

Private Sub Section2_Format(ByVal pFormattingInfo As Object)
'  If mf_PicHtSub = 0 Then mf_PicHtSub = Subreport2_PicSub.Height
'  If mf_PicWdthSub = 0 Then mf_PicWdthSub = Subreport2_PicSub.Width
'  If UCase(Subreport2_FldPicOpt.Value) = "3D" Then
'    Call SetRptPic(mf_PicHtSub, mf_PicWdthSub, Subreport2_PicSub, "DM", False, True, Subreport2_FldDmCd.Value, Subreport2_FldDmSz.Value, , , , , , , , Subreport2_FldOdPicNm.Value)
'  ElseIf UCase(Subreport2_FldPicOpt.Value) = "HSK" Then
'    Call SetRptPic(mf_PicHtSub, mf_PicWdthSub, Subreport2_PicSub, "DM", True, True, Subreport2_FldDmCd.Value, Subreport2_FldDmSz.Value, , , , , , , , Subreport2_FldOdPicNm.Value)
'  End If
  If mf_PicHt = 0 Then mf_PicHt = Picture1.Height
  If mf_PicWdth = 0 Then mf_PicWdth = Picture1.Width
  If UCase(FldPicOpt.Value) = "3D" Then
    '*** Jay 3.2.0 [PicNm]
    Call SetRptPic(mf_PicHt, mf_PicWdth, Picture1, "DM", False, True, flddmcd.Value, flddmsz.Value, , , , , , , , FldOdPicNm.Value)
    '*** Jay 3.2.0 [PicNm]
  ElseIf UCase(FldPicOpt.Value) = "HSK" Then
    '*** Jay 3.2.0 [PicNm]
    Call SetRptPic(mf_PicHt, mf_PicWdth, Picture1, "DM", True, True, flddmcd.Value, flddmsz.Value, , , , , , , , FldOdPicNm.Value)
    'Call SetRptPic(mf_PicHt, mf_PicWdth, Pic, "DM", True, True, FldDmCd.Value, FldDmSz.Value)
    '*** Jay 3.2.0 [PicNm]
  End If

End Sub
