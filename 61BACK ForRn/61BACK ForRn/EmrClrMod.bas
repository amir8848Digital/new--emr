Attribute VB_Name = "EmrClrMod"
'*** Jenny Color (For Setting System Colors)
    '*** MSDN Help Topic
    'HOWTO: Use Custom System Colors Only When Your App Has Focus
    '*** MSDN Help Topic

    Declare Function SetSysColors Lib "user32" (ByVal nChanges As _
       Long, lpSysColor As Long, lpColorValues As Long) As Long
    Declare Function GetSysColor Lib "user32" (ByVal nIndex As Long) _
       As Long
       
    '##$$ (Bef 09/08/05) This is necessary for Hook & UnHook to work. Uncomment it whenever required
    'Declare Function CallWindowProc Lib "user32" Alias _
    '   "CallWindowProcA" (ByVal lpPrevWndFunc As Long, _
    '   ByVal hWnd As Long, ByVal Msg As Long, ByVal wParam As Long, _
    '   ByVal lParam As Long) As Long
    'Declare Function SetWindowLong Lib "user32" Alias _
    '   "SetWindowLongA" (ByVal hWnd As Long, ByVal nIndex As Long, _
    '   ByVal dwNewLong As Long) As Long
    '
    'Public Const WM_ACTIVATEAPP = &H1C
    'Public Const GWL_WNDPROC = -4
    '##$$ (Bef 09/08/05) This is necessary for Hook & UnHook to work. Uncomment it whenever required
    

    Private Const COLOR_ACTIVECAPTION = 2        '#*** To Change: the Caption BackGround of the active form
    Private Const COLOR_MENU = 4                 '#*** To Change: Menu BackGround Color, Control Box BackGround Color
    Private Const COLOR_WINDOWFRAME = 6          '#*** To Change: Tab Page Border Colour, Selected Button full Highlight colour, PictureBox Frame colour
    Private Const COLOR_MENUTEXT = 7             '#*** To Change: Menu Text Color, Control Box Text Color
    Private Const COLOR_CAPTIONTEXT = 9          '#*** To Change: the Caption Text Color of the active form
    Private Const COLOR_HIGHLIGHT = 13           '#*** To Change: the BackGround Color of the Highlighted Menu
    Private Const COLOR_HIGHLIGHTTEXT = 14       '#*** To Change: the Text Color of the Highlighted Menu
    Private Const COLOR_BTNFACE = 15             '#####*** To Change: Form BackGround Color, TabPage BackGround Color, Frame  BackGround Color, Button BackGround Color, ScrollBar Arrow Container & Slider Color, statusbar color
    Private Const COLOR_BTNSHADOW = 16           '#*** To Change the Shadow color of CommandButtons, TabPage, ScrollBar, Flex Grid
    Private Const COLOR_GRAYTEXT = 17            '##*** To Change: Text color of disabled TextBox
    Private Const COLOR_BTNTEXT = 18             '#*** To Change: Button Text color, Label Text Color, Frame Text Color, ScrollBar Arrow triangle pic Color, (Min, Max, Normal button pic colors of a form)
    Private Const COLOR_BTNHIGHLIGHT = 20        '#*** To Change the Border Highlight color of Forms, CommandButtons, TabPage, ScrollBar, Flex Grid, TextBox etc.

    'Private Const COLOR_SCROLLBAR = 0            '#*** To Change: Blinking color of slider of ScrollBar
    'Private Const COLOR_BACKGROUND = 1
    'Private Const COLOR_INACTIVECAPTION = 3      '*** To Change: the Caption BackGround of the in-active form
    'Private Const COLOR_WINDOW = 5               '##*** To Change: TextBox BackGround Color, ScrollBar Scroll Area Color
    'Private Const COLOR_WINDOWTEXT = 8           '*** To Change: TextBox Text color
    'Private Const COLOR_ACTIVEBORDER = 10        '*** To Change: the Border Color of the active form
    Private Const COLOR_INACTIVEBORDER = 11      '*** To Change: the Border Color of the in-active form
    'Private Const COLOR_APPWORKSPACE = 12        '*** To Change: MDI Background Color
    'Private Const COLOR_INACTIVECAPTIONTEXT = 19 '*** To Change: the Caption Text Color of the in-active form

    'Global gArr_OriginalColors(20) As Long
    'Global gArr_NewColors(20) As Long
    'Global gArr_IndexArray(20) As Long
    Private gArr_OriginalColors(12) As Long
    Private gArr_NewColors(12) As Long
    Private gArr_IndexArray(12) As Long
    
    '##$$ (Bef 09/08/05) This is necessary for Hook & UnHook to work. Uncomment it whenever required
    'Global lpPrevWndProc As Long
    'Global gHW As Long
    '##$$ (Bef 09/08/05) This is necessary for Hook & UnHook to work. Uncomment it whenever required
    
'*** Jenny Color (For Setting System Colors)

'*** Jenny Color
Private go_RecClr As MwfLib.MDORowSet
'Public gs_ClrCd As String
Public Const ctDefClrCd = "DFCLR"

'Public go_RecFrmPic As MwfLib.MDORowSet
'Public go_FrmPic As Picture

'*** 16/05/05
Private go_MDIFrm As MDIForm
'*** 16/05/05

'*** (12/08/05)
Private go_FrmLst As Form
Private gs_ClrPth As String
'*** (12/08/05)

'*** 13/06/05
Private gs_MnuCd As String
'Public go_ADCPic As Picture
'*** 13/06/05

'*** (Jen 2.13)
Public wb_ApplyClr As Boolean
'*** (Jen 2.13)

'*** Jenny Color


'*** Jenny Color (For Setting System Colors)

'##$$ (Bef 09/08/05) This is necessary for Hook & UnHook to work. Uncomment it whenever required
'Public Sub Hook()
'  '*** Establish a hook to capture messages to this window.
'  lpPrevWndProc = SetWindowLong(gHW, GWL_WNDPROC, _
'     AddressOf WindowProc)
'End Sub
'Public Sub Unhook()
'  Dim temp As Long
'
'  '*** Reset the message handler for this window.
'  temp = SetWindowLong(gHW, GWL_WNDPROC, lpPrevWndProc)
'End Sub
'Function WindowProc(ByVal hw As Long, ByVal uMsg As Long, _
'  ByVal wParam As Long, ByVal lParam As Long) As Long
'  '*** Check for the ActivateApp message.
'  If uMsg = WM_ACTIVATEAPP Then
'    '*** Check to see if Activating the application.
'    If wParam <> 0 Then
'      '*** Check to see if application is minimized.
'      'If EmrSuppFrmMDI.WindowState <> vbMinimized Then   '*** (Bef 16/05/05)
'      If go_MDIFrm.WindowState <> vbMinimized Then
'        '*** Use custom colors.
'        SetNewColors
'      End If
'    Else
'      '*** Application is DeActivating so restore normal colors.
'      RestoreColors
'    End If
'  End If
'
'  '*** Pass message on to the original window message handler.
'  WindowProc = CallWindowProc(lpPrevWndProc, hw, uMsg, wParam, _
'    lParam)
'End Function
'##$$ (Bef 09/08/05) This is necessary for Hook & UnHook to work. Uncomment it whenever required


'*** (13/08/05)
Public Function GetMDIFrm() As MDIForm
  Set GetMDIFrm = go_MDIFrm
End Function
Public Sub SetMDIFrm(ByVal po_MDIFrm As MDIForm)
  Set go_MDIFrm = po_MDIFrm
End Sub
Public Function GetFrmLst() As Form
  Set GetFrmLst = go_FrmLst
End Function
Public Sub SetFrmLst(ByVal po_FrmLst As Form)
  Set go_FrmLst = po_FrmLst
End Sub
Public Function GetMnuCd() As String
  GetMnuCd = gs_MnuCd
End Function
Public Sub SetMnuCd(ByVal po_MnuCd As String)
  gs_MnuCd = po_MnuCd
End Sub
Public Function GetRecClrSet() As MwfLib.MDORowSet
  Set GetRecClrSet = go_RecClr
End Function
Public Sub SetRecClrSet(ByVal po_RecClrSet As MwfLib.MDORowSet)
  Set go_RecClr = po_RecClrSet
End Sub
Public Function GetClrPth() As String
  GetClrPth = gs_ClrPth
End Function
Public Sub SetClrPth(ByVal po_Pth As String)
  gs_ClrPth = po_Pth
End Sub
'*** (13/08/05)

Public Sub RestoreColors()
  '*** SetSysColors takes the following parameters:
  '    The first parameter indicates the total number of system colors you are attempting to change.
  '    The second parameter is an array of the numeric values for the display aspects you want to change.
  '    The third parameter is also an array whose elements are the new colors for the display aspects defined by the first array.

  SetSysColors 13, gArr_IndexArray(0), gArr_OriginalColors(0)
  'SetSysColors 21, gArr_IndexArray(0), gArr_OriginalColors(0)
End Sub
Public Sub SetNewColors()
  SetSysColors 13, gArr_IndexArray(0), gArr_NewColors(0)
  'SetSysColors 21, gArr_IndexArray(0), gArr_NewColors(0)
End Sub
Public Sub SaveOriginalColors()
  Dim i As Long
  
  '*** The GetSysColor API function is straightforward; you just pass a numeric value for
  '    the system setting whose color you seek, and the function returns the color.
  
  '*** Retrieve all current color settings.
  'For i = 0 To 20
  '  gArr_OriginalColors(i) = GetSysColor(i)
  '  gArr_NewColors(i) = gArr_OriginalColors(i)
  'Next i
  
  
  '*** Randomly pick and assign new colors.
  gArr_IndexArray(0) = COLOR_ACTIVECAPTION
  gArr_IndexArray(1) = COLOR_MENU
  gArr_IndexArray(2) = COLOR_WINDOWFRAME
  gArr_IndexArray(3) = COLOR_MENUTEXT
  gArr_IndexArray(4) = COLOR_CAPTIONTEXT
  gArr_IndexArray(5) = COLOR_HIGHLIGHT
  gArr_IndexArray(6) = COLOR_HIGHLIGHTTEXT
  gArr_IndexArray(7) = COLOR_BTNFACE
  gArr_IndexArray(8) = COLOR_BTNSHADOW
  gArr_IndexArray(9) = COLOR_GRAYTEXT
  gArr_IndexArray(10) = COLOR_BTNTEXT
  gArr_IndexArray(11) = COLOR_BTNHIGHLIGHT

  '*** Tmp
  gArr_IndexArray(12) = COLOR_INACTIVEBORDER
  'gArr_IndexArray(12) = COLOR_SCROLLBAR
  'gArr_IndexArray(13) = COLOR_BACKGROUND
  'gArr_IndexArray(14) = COLOR_INACTIVECAPTION
  'gArr_IndexArray(15) = COLOR_WINDOW
  'gArr_IndexArray(16) = COLOR_WINDOWTEXT
  'gArr_IndexArray(17) = COLOR_ACTIVEBORDER
  'gArr_IndexArray(18) = COLOR_INACTIVEBORDER
  'gArr_IndexArray(19) = COLOR_APPWORKSPACE
  'gArr_IndexArray(20) = COLOR_INACTIVECAPTIONTEXT
  '*** Tmp


  
'  '*** Randomly pick and assign new colors.
'  gArr_IndexArray(0) = COLOR_SCROLLBAR
'  gArr_IndexArray(1) = COLOR_BACKGROUND
'  gArr_IndexArray(2) = COLOR_ACTIVECAPTION
'  gArr_IndexArray(3) = COLOR_INACTIVECAPTION
'  gArr_IndexArray(4) = COLOR_MENU
'  gArr_IndexArray(5) = COLOR_WINDOW
'  gArr_IndexArray(6) = COLOR_WINDOWFRAME
'  gArr_IndexArray(7) = COLOR_MENUTEXT
'  gArr_IndexArray(8) = COLOR_WINDOWTEXT
'  gArr_IndexArray(9) = COLOR_CAPTIONTEXT
'  gArr_IndexArray(10) = COLOR_ACTIVEBORDER
'  gArr_IndexArray(11) = COLOR_INACTIVEBORDER
'  gArr_IndexArray(12) = COLOR_APPWORKSPACE
'  gArr_IndexArray(13) = COLOR_HIGHLIGHT
'  gArr_IndexArray(14) = COLOR_HIGHLIGHTTEXT
'  gArr_IndexArray(15) = COLOR_BTNFACE
'  gArr_IndexArray(16) = COLOR_BTNSHADOW
'  gArr_IndexArray(17) = COLOR_GRAYTEXT
'  gArr_IndexArray(18) = COLOR_BTNTEXT
'  gArr_IndexArray(19) = COLOR_INACTIVECAPTIONTEXT
'  gArr_IndexArray(20) = COLOR_BTNHIGHLIGHT
  
  For i = 0 To 12
    gArr_OriginalColors(i) = GetSysColor(gArr_IndexArray(i))
    gArr_NewColors(i) = gArr_OriginalColors(i)
  Next i
  
  
  '*** (Bef 09/08/05)
  'gArr_OriginalColors(0) = GetSysColor(COLOR_ACTIVECAPTION)
  'gArr_OriginalColors(1) = GetSysColor(COLOR_MENU)
  'gArr_OriginalColors(2) = GetSysColor(COLOR_WINDOWFRAME)
  'gArr_OriginalColors(3) = GetSysColor(COLOR_MENUTEXT)
  'gArr_OriginalColors(4) = GetSysColor(COLOR_CAPTIONTEXT)
  'gArr_OriginalColors(5) = GetSysColor(COLOR_HIGHLIGHT)
  'gArr_OriginalColors(6) = GetSysColor(COLOR_HIGHLIGHTTEXT)
  'gArr_OriginalColors(7) = GetSysColor(COLOR_BTNFACE)
  'gArr_OriginalColors(8) = GetSysColor(COLOR_BTNSHADOW)
  'gArr_OriginalColors(9) = GetSysColor(COLOR_GRAYTEXT)
  'gArr_OriginalColors(10) = GetSysColor(COLOR_BTNTEXT)
  'gArr_OriginalColors(11) = GetSysColor(COLOR_BTNHIGHLIGHT)
  '
  ''*** Tmp
  'gArr_OriginalColors(12) = GetSysColor(COLOR_INACTIVEBORDER)
  '
  ''gArr_OriginalColors(12) = GetSysColor(COLOR_SCROLLBAR)
  ''gArr_OriginalColors(13) = GetSysColor(COLOR_BACKGROUND)
  ''gArr_OriginalColors(14) = GetSysColor(COLOR_INACTIVECAPTION)
  ''gArr_OriginalColors(15) = GetSysColor(COLOR_WINDOW)
  ''gArr_OriginalColors(16) = GetSysColor(COLOR_WINDOWTEXT)
  ''gArr_OriginalColors(17) = GetSysColor(COLOR_ACTIVEBORDER)
  ''gArr_OriginalColors(18) = GetSysColor(COLOR_INACTIVEBORDER)
  ''gArr_OriginalColors(19) = GetSysColor(COLOR_APPWORKSPACE)
  ''gArr_OriginalColors(20) = GetSysColor(COLOR_INACTIVECAPTIONTEXT)
  ''*** Tmp
  '
  'For i = 0 To 12
  '  gArr_NewColors(i) = gArr_OriginalColors(i)
  'Next i
  
  '*** (Bef 09/08/05)
  
End Sub
Public Sub SetMDISysColorsBefLogin()
  gArr_NewColors(0) = &HE4BCC5       '&HFDE9BF    '&HFFC0C0          '*** Form Caption Bkg   '&H00000040& '&H00000000&  '&HFFC0C0     '&HFFFFC0             '&H0& '&H00FDE9BF&
  gArr_NewColors(1) = &HE0E0E0          '*** Menu Bkg '&H00D1E0FC&   '&H00E0E0E0&
  'gArr_NewColors(2) = gArr_OriginalColors(2)  '*** Border
  gArr_NewColors(2) = &HFDE9BF          '&HF9EAFF        '*** Border
  gArr_NewColors(3) = &H0&           '*** Menu Txt  '&H00000000&
  'gArr_NewColors(4) = gArr_OriginalColors(4)  '*** Caption Text Color of the active form '&H00000000&
  gArr_NewColors(4) = &H0&         '*** Caption Text Color of the active form '&H00000000&
  
  gArr_NewColors(5) = gArr_OriginalColors(5)  '*** BackGround Color of the Highlighted Menu
  gArr_NewColors(6) = gArr_OriginalColors(6)  '*** Text Color of the Highlighted Menu
  gArr_NewColors(7) = &HE0E0E0       '*** ##### Tab Page Bkg        '&H00000080& '&H8000000F&
  'gArr_NewColors(8) = gArr_OriginalColors(8) '*** Shadow
  gArr_NewColors(8) = &HC9B18F       '&H9B92C7  '&HC0C0FF '&H00C9B18F&  '*** Shadow
  gArr_NewColors(9) = gArr_OriginalColors(9)  '*** Text color of disabled TextBox
  gArr_NewColors(10) = &H0&                 '*** ##### Btn Txt       '&H00FFFFFF&  '&H00000000&
  'gArr_NewColors(11) = gArr_OriginalColors(11) '*** Border Highlight
  gArr_NewColors(11) = &HFFF2DF       '&HFDE9BF        '*** Border Highlight    '&HF9EAFF  '&H00FDE9BF&  '&H00F9EAFF& '&H00FFF2DF&
  
  
  '*** Tmp
  gArr_NewColors(12) = gArr_OriginalColors(12)
  'gArr_IndexArray(12) = COLOR_SCROLLBAR
  'gArr_IndexArray(13) = COLOR_BACKGROUND
  'gArr_IndexArray(14) = COLOR_INACTIVECAPTION
  'gArr_IndexArray(15) = COLOR_WINDOW
  'gArr_IndexArray(16) = COLOR_WINDOWTEXT
  'gArr_IndexArray(17) = COLOR_ACTIVEBORDER
  'gArr_IndexArray(18) = COLOR_INACTIVEBORDER
  'gArr_IndexArray(19) = COLOR_APPWORKSPACE
  'gArr_IndexArray(20) = COLOR_INACTIVECAPTIONTEXT
  '*** Tmp
  
  
  SetNewColors
End Sub
Public Sub SetMDISysColorsAfterLogin()
  If wb_ApplyClr = False Then Exit Sub '*** (Jen 2.13)
  If go_RecClr Is Nothing Then Exit Sub     '*** ###
  If go_RecClr.RecCount > 0 Then
    go_RecClr.MoveFirst

    gArr_NewColors(0) = Val(go_RecClr!CsSysActiveCaption)      '*** Form Caption Bkg
    gArr_NewColors(1) = Val(go_RecClr!CsSysMenu)               '*** Menu Bkg
    gArr_NewColors(2) = Val(go_RecClr!CsSysWindowFrame)        '*** Border
    gArr_NewColors(3) = Val(go_RecClr!CsSysMenuText)           '*** Menu Txt
    gArr_NewColors(4) = Val(go_RecClr!CsSysCaptionText)        '*** Caption Text Color of the active form
    gArr_NewColors(5) = Val(go_RecClr!CsSysHighLight)          '*** BackGround Color of the Highlighted Menu
    gArr_NewColors(6) = Val(go_RecClr!CsSysHighLightText)      '*** Text Color of the Highlighted Menu
    gArr_NewColors(7) = Val(go_RecClr!CsSysBtnFace)            '*** ##### Tab Page Bkg
    gArr_NewColors(8) = Val(go_RecClr!CsSysBtnShadow)          '*** Shadow
    gArr_NewColors(9) = Val(go_RecClr!CsSysGrayText)           '*** Text color of disabled TextBox
    gArr_NewColors(10) = Val(go_RecClr!CsSysBtnText)           '*** ##### Btn Txt
    gArr_NewColors(11) = Val(go_RecClr!CsSysBtnHighLight)      '*** Border Highlight
  
    '*** Tmp
    gArr_NewColors(12) = Val(go_RecClr!CsSysBtnFace)
'    If UCase(EmrFrmMDI.TXT_UCD) = "CLR4" Then
'      gArr_NewColors(12) = 6383772    '&H00001E6C&  '6383772
'    Else
'      gArr_NewColors(12) = Val(go_RecClr!CsSysBtnFace)     '&HC2A8DB     '*** Medium Pink
'    End If

    'gArr_NewColors(12) = &H0&          '&HFFFFFF    '&HC2A8DB
    'gArr_NewColors(13) = &H0&
    'gArr_NewColors(14) = &H0&
    'gArr_NewColors(15) = &H0&
    'gArr_NewColors(16) = &H0&
    'gArr_NewColors(17) = &HFF0000           '*** Bright Blue &H00FF0000&
    'gArr_NewColors(18) = &HFF00&      '**************   Light Green
    'gArr_NewColors(19) = &H0&
    'gArr_NewColors(20) = &H0&
    '*** Tmp
  
    SetNewColors
  End If
End Sub
Public Sub SetTabSysColors(ByVal po_BckClr As OLE_COLOR, ByVal po_ForeClr As OLE_COLOR, ByVal po_BorderClr As OLE_COLOR, ByVal po_ShadowClr As OLE_COLOR, ByVal po_HighlightClr As OLE_COLOR)
  '*** (IMP) For setting the Tab Color of TabPage (TabPage bahaves differently. it takes the previous setting)
  gArr_NewColors(2) = po_BorderClr                 '*** Border    '&HC0E0FF  '&H404080
  gArr_NewColors(7) = po_BckClr          '*** ##### Tab Page Bkg        '&H00000080& '&H8000000F&
  gArr_NewColors(8) = po_ShadowClr                  '*** Shadow '&H00800080&  '&H40C0&
  gArr_NewColors(10) = po_ForeClr         '*** ##### Btn Txt       '&H00000080& '&H8000000F&
  gArr_NewColors(11) = po_HighlightClr               '*** Border Highlight        '&H00F5E1FF&   '&H00FF80FF&  '&H404080  '&HC0E0FF
  
  SetNewColors
  'SetSysColors 1, gArr_IndexArray(7), gArr_NewColors(7)
End Sub

'*** (Bef 07/07/05)
'Public Sub AdjustBtnPicSz(ByVal po_Ctl As MWCTL_BTN1)
'    '*** 12/05/05
'    'Dim wi_StoreScaleMode As Integer
'    'wi_StoreScaleMode = Me.ScaleMode   'po_Ctl.parent.ScaleMode
'    'Me.ScaleMode = 3
'    'po_Ctl.parent.ScaleMode = 3
'
'    Dim PictureClip1 As PictureClip
'
'    'Set PictureClip1 = EmrSuppFrmMDI.PictureClip1   '*** (Bef 16/05/05)
'    Set PictureClip1 = go_MDIFrm.PictureClip1    '*** (16/05/05)
'    'Set EmrSuppFrmMDI.PictureClip1.Picture = po_Ctl.Pic
'    Set PictureClip1.Picture = po_Ctl.Pic
'
'
'    '*** Get X and Y coordinates of the clipping region.
'    PictureClip1.ClipX = 0
'    PictureClip1.ClipY = 0
'
'    '*** Set the area of the clipping region (in pixels).
'    PictureClip1.ClipHeight = PictureClip1.Height
'    PictureClip1.ClipWidth = PictureClip1.Width
'
'    '*** Set the destination area to fill the Button.
'    '*** For conversion from Twip to Pixel divide by 15.
'    If PictureClip1.StretchX <> po_Ctl.Width / 15 Then PictureClip1.StretchX = po_Ctl.Width / 15
'    If PictureClip1.StretchY <> po_Ctl.Height / 15 Then PictureClip1.StretchY = po_Ctl.Height / 15
'
'    '*** Assign the clipped bitmap to the Button.
'    po_Ctl.Pic = PictureClip1.Clip
'    'po_Ctl.parent.ScaleMode = wi_StoreScaleMode
'    'Me.ScaleMode = wi_StoreScaleMode
'
'    '*** 12/05/05
'End Sub
'Public Function AdjustBtnPicSize(ByVal fo_Pic As Picture, ByVal fi_Height As Integer, fi_Width As Integer) As Picture
'    '*** 12/05/05
'    Dim PictureClip1 As PictureClip
'
'    'Set PictureClip1 = EmrSuppFrmMDI.PictureClip1    '*** (Bef 16/05/05)
'    Set PictureClip1 = go_MDIFrm.PictureClip1     '*** (16/05/05)
'    Set PictureClip1.Picture = fo_Pic
'
'    '*** Get X and Y coordinates of the clipping region.
'    PictureClip1.ClipX = 0
'    PictureClip1.ClipY = 0
'
'    '*** Set the area of the clipping region (in pixels).
'    PictureClip1.ClipHeight = PictureClip1.Height
'    PictureClip1.ClipWidth = PictureClip1.Width
'
'    '*** Set the destination area to fill the Button.
'    '*** For conversion from Twip to Pixel divide by 15.
'    If PictureClip1.StretchX <> fi_Width / 15 Then PictureClip1.StretchX = fi_Width / 15
'    If PictureClip1.StretchY <> fi_Height / 15 Then PictureClip1.StretchY = fi_Height / 15
'
'    '*** Assign the clipped bitmap to the Button.
'    Set AdjustBtnPicSize = PictureClip1.Clip
'    '*** 12/05/05
'End Function
'*** (Bef 07/07/05)


Public Function CreateColorImage(mo_PicBox As PictureBox, ms_ColorCd As String, mi_Width As Integer, mi_Height As Integer) As Picture
  '*** Jenny Color
  '*** CreateColorImage is passed a picturebox which is
  '    invisible. It's also passed an image list which it will add
  '    a picture to, and sColor specifies that color in string representation.
  
  '*** The idea is, to Size the picturebox to the desired picture size,
  '    draw a border around it, and then set it's image property to
  '    it's picture property, which essentially takes a snapshot of the
  '    picture box and paints it. Then we can set this picture in out
  '    image list to be used in out combo box.
  
  mo_PicBox.Picture = LoadPicture()
 
  With mo_PicBox
    .AutoRedraw = True  '*** Ensure "True" painting
    
    '*** Set the width and height of the image in pixels. A height
    '    of 16 will leave no border between items in the combo box/ grid column.
    '    12 looks nive.
    '.Width = .ScaleX(32, vbPixels, vbTwips)
    '.Height = .ScaleY(12, vbPixels, vbTwips)
    .Width = .ScaleX(mi_Width, vbPixels, vbTwips)
    .Height = .ScaleY(mi_Height, vbPixels, vbTwips)
    
    
    '*** Flatten box to iliminate distortion of picture (ensures the image
    '    we create is in fact the dimensions we just specified.
    .BorderStyle = 0
    .Appearance = 0
    
    '*** By calling the GetColorFromString function on the color string
    '    passed to this sub, we can set the backcolor of the picture box
    '    to that color.
    .BackColor = IIF(Val(ms_ColorCd) < 0, 0, ms_ColorCd)
    
    '*** Now, using the Line function, we'll draw a 1 pixel black line around
    '    the box for a better look. (The white image won't bleed into background
    '    of the image combo.
    .ForeColor = vbBlack
    mo_PicBox.Line (0, 0)-(.Width - .ScaleX(1, vbPixels, vbTwips) _
                   , .Height - .ScaleY(1, vbPixels, vbTwips)), , B
    
    '*** Set the image (the way the picture box looks) to the box's picture
    '    property, and add it to the image list. Using the color string
    '    (all uppercase now) as the images key in the image list.
    .Picture = .Image
    Set CreateColorImage = .Picture
  End With
  '*** Jenny Color
End Function
Public Sub SetObjColors(ByVal po_Frm As Object, Optional ByVal ps_BtnStr As String)
  If (go_RecClr Is Nothing) Then Exit Sub
  
  po_Frm.Icon = GetMDIFrm.Icon
  If IsMissing(ps_BtnStr) Then
    ms_BtnStr = ""
  Else
    ms_BtnStr = ps_BtnStr
  End If
  
  If ms_BtnStr <> "" Then
    ms_BtnStr = UCase(ms_BtnStr)
    If InStr(1, Mid(ms_BtnStr, 1, 1), ",") <= 0 Then ms_BtnStr = "," + ms_BtnStr
    If InStr(1, Mid(ms_BtnStr, Len(ms_BtnStr), 1), ",") <= 0 Then ms_BtnStr = ms_BtnStr + ","
  End If
  
  'If TypeOf wo_Ctl Is Menu Then
  
  Dim wo_Ctl As Control
  Dim wi_ADCBtnWdthDiff As Integer
  
  Dim wo_CsFlxPic As Picture
  Dim wo_CsBtnPicSet1 As Picture
  Dim wo_CsADCBtnPic As Picture
  Dim wo_CsHlpPic As Picture
  
  '*** FraNKeyAll is not present in all the forms
  'po_Frm.FraNKeyAll.Height = po_Frm.FRA_NKEYS.Height + 500
  'po_Frm.FraNKeyAll.Width = po_Frm.FRA_NKEYS.Width + 500
  
  If go_RecClr.RecCount > 0 Then
    go_RecClr.MoveFirst
    
    'if gs_ClrCd <> ctDefClrCd Then
    If UCase(go_RecClr!CsCd) <> ctDefClrCd Then
      ''And Trim(go_RecClr!CsFormPic) <> "" And Dir(Trim(go_RecClr!CsFormPic)) <> ""
      ''po_Frm.Picture = go_RecClr.FldPicValue("CsFormPic")    '*** (Bef 23/05/05)
      
      '(Bef 15/06/05) po_Frm.Picture = LoadPicture(go_RecClr!CsFormPic)       '*** (23/05/05)
      'po_Frm.Picture = go_MDIFrm.Picture1.Picture   '*** 15/06/05    '*** (Bef 2.14 Next Patch 1)
      po_Frm.BackColor = Val(go_RecClr!CsFraBckClr)       '*** (Jen 2.14 Next Patch 1)
    Else
      po_Frm.Picture = LoadPicture("")
      po_Frm.BackColor = Val(go_RecClr!CsSysBtnFace)     '*** (Bef 2.14 Next Patch 1)
    End If
    
    po_Frm.ForeColor = Val(go_RecClr!CsLblForClr)
    
   
    '*** (Bef 18/06/05)
    ''*** 15/06/05
    'Set wo_CsFlxPic = LoadPicture(go_RecClr!CsFlxPic)
    'Set wo_CsBtnPicSet1 = go_MDIFrm.ImgDefBtnSet1
    'Set wo_CsADCBtnPic = go_ADCPic
    'Set wo_CsHlpPic = LoadPicture(go_RecClr!CsHlpPic)
    ''*** 15/06/05
    '*** (Bef 18/06/05)
    
    '*** (18/06/05)
    '*** (Bef 13/08/05)
    'Dim ws_Path As String
    'ws_Path = ""
    'If InStrRev(App.Path, "\") > 1 Then ws_Path = Mid(App.Path, 1, InStrRev(App.Path, "\") - 1)
    'If ws_Path <> "" Then ws_Path = ws_Path & "\MwClrSet\"
    'If Trim(go_RecClr!CsFlxPic) <> "" And Dir(Trim(GetClrPth & go_RecClr!CsFlxPic)) <> "" Then Set wo_CsFlxPic = LoadPicture(GetClrPth & go_RecClr!CsFlxPic)
    'Set wo_CsBtnPicSet1 = go_MDIFrm.CMDOK.Pic
    'Set wo_CsADCBtnPic = go_MDIFrm.ImgDefBtnSet
    'If Trim(go_RecClr!CsHlpPic) <> "" And Dir(Trim(GetClrPth & go_RecClr!CsHlpPic)) <> "" Then Set wo_CsHlpPic = LoadPicture(GetClrPth & go_RecClr!CsHlpPic)
    '*** (Bef 13/08/05)
    
    '*** (13/08/05)
    If Trim(go_RecClr!CsFlxPic) <> "" And IsValidPath(GetClrPth & Trim(go_RecClr!CsFlxPic)) = True Then Set wo_CsFlxPic = LoadPicture(GetClrPth & go_RecClr!CsFlxPic)
    Set wo_CsBtnPicSet1 = go_MDIFrm.CMDOK.Pic
    Set wo_CsADCBtnPic = go_MDIFrm.ImgDefBtnSet
    If Trim(go_RecClr!CsHlpPic) <> "" And IsValidPath(GetClrPth & Trim(go_RecClr!CsHlpPic)) = True Then Set wo_CsHlpPic = LoadPicture(GetClrPth & go_RecClr!CsHlpPic)
    '*** (13/08/05)
    '*** (18/06/05)
   
    '*** (Bef 15/06/05)
    ''*** (23/05/05)
    'Set wo_CsFlxPic = LoadPicture(go_RecClr!CsFlxPic)
    'If Trim(go_RecClr!CsBtnPicSet1) = "" Or Dir(Trim(go_RecClr!CsBtnPicSet1)) = "" Then
    '  Set wo_CsBtnPicSet1 = go_MDIFrm.ImgDefBtnSet1
    'Else
    '  Set wo_CsBtnPicSet1 = LoadPicture(go_RecClr!CsBtnPicSet1)
    'End If
    'If Trim(go_RecClr!CsADCBtnPic) = "" Or Dir(Trim(go_RecClr!CsADCBtnPic)) = "" Then
    '  Set wo_CsADCBtnPic = go_MDIFrm.ImgDefBtnSet1
    'Else
    '  Set wo_CsADCBtnPic = LoadPicture(go_RecClr!CsADCBtnPic)
    'End If
    'Set wo_CsHlpPic = LoadPicture(go_RecClr!CsHlpPic)
    ''*** (23/05/05)
    '*** (Bef 15/06/05)
    
    wi_ADCBtnWdthDiff = po_Frm.adc.Left + (3 * IIF(go_RecClr!CsADCBtnWidth > 0, go_RecClr!CsADCBtnWidth, po_Frm.adc.BtnWidth))
    
    For Each wo_Ctl In po_Frm.Controls
        If TypeOf wo_Ctl Is MWCTL_MED Then
          wo_Ctl.BckColor = Val(go_RecClr!CsMedBckClr)
          wo_Ctl.ForColor = Val(go_RecClr!CsMedForClr)
          wo_Ctl.BckColorDisable = Val(go_RecClr!CsMedBckClrDisa)
          ''wo_Ctl.ForColorDisable = &HFFFFFF
          wo_Ctl.BckColorSel = Val(go_RecClr!CsMedBckClrSel)
        ElseIf TypeOf wo_Ctl Is Label Then
'          If go_RecClr!Cs3D = "N" And wo_Ctl.BorderStyle = 1 Then
'            wo_Ctl.Appearance = 0
'          End If
'          wo_Ctl.ForeColor = Val(go_RecClr!CsLblForClr)
          
          wo_Ctl.ForeColor = Val(go_RecClr!CsLblForClr)
          If wo_Ctl.BorderStyle = 1 Then
            If go_RecClr!Cs3D = "N" Then
              wo_Ctl.Appearance = 0
              wo_Ctl.BorderStyle = 0
              wo_Ctl.Top = wo_Ctl.Top + 30
              wo_Ctl.Left = wo_Ctl.Left + 30
              'wo_Ctl.FontBold = True
              'wo_Ctl.FontSize = 9
            End If
            wo_Ctl.ForeColor = Val(go_RecClr!CsLblMainClr)  '*** (IMPORTANT) ForeColor should be set after setting the Appearance property as changing the Appearance property again changes the ForeColor to the default color.
          End If
        ElseIf TypeOf wo_Ctl Is MWCTL_FLX Then
          wo_Ctl.BckColor = Val(go_RecClr!CsFlxBckClr)
          wo_Ctl.BckColorBkg = Val(go_RecClr!CsFlxBckClrBkg)
          wo_Ctl.BckColorFixed = Val(go_RecClr!CsFlxBckClrFix)
          wo_Ctl.BckColorSel = Val(go_RecClr!CsFlxBckClrSel)
          wo_Ctl.BckColorSelRow = Val(go_RecClr!CsFlxBckClrSelRow)
          wo_Ctl.BckColorSelTxt = Val(go_RecClr!CsFlxBckClrSelTxt)
          
          wo_Ctl.ForColor = Val(go_RecClr!CsFlxForClr)
          wo_Ctl.ForColorFixed = Val(go_RecClr!CsFlxForClrFix)
          wo_Ctl.ForColorSel = Val(go_RecClr!CsFlxForClrSel)
          wo_Ctl.ForColorSelTxt = Val(go_RecClr!CsFlxForClrSelTxt)
          wo_Ctl.GrdColor = Val(go_RecClr!CsFlxGrdClr)
          If wo_Ctl.GrdLinesFixed = flxGrdFlat Then
            'Dim www As MWCTL_FLX
            wo_Ctl.GrdColorFixed = Val(go_RecClr!CsFlxGrdClrFix1)
          Else
            If go_RecClr!CsGrdInset = "Y" Then
              wo_Ctl.GrdColorFixed = Val(go_RecClr!CsFlxGrdClrFix)
            Else
              wo_Ctl.GrdLinesFixed = flxGrdFlat
              wo_Ctl.GrdColorFixed = Val(go_RecClr!CsFlxGrdClrFix1)
            End If
          End If
          
          If go_RecClr!Cs3D = "N" Then
            wo_Ctl.GrdAppearance = flxGrdAppFlat
          End If
          
          'wo_Ctl.Pic = go_RecClr.FldPicValue("CsFlxPic")
          wo_Ctl.Pic = wo_CsFlxPic
        ElseIf TypeOf wo_Ctl Is MWCTL_BTN1 Then
          If InStr(1, ms_BtnStr, "," + UCase(wo_Ctl.Name) + ",") <= 0 Then
            wo_Ctl.ForeColor = Val(go_RecClr!CsBtnForClrSet1)
            wo_Ctl.DisableColor = Val(go_RecClr!CsBtnDisaClrSet1)
            wo_Ctl.MaskColor = Val(go_RecClr!CsBtnMaskClrSet1)
            wo_Ctl.Pic = wo_CsBtnPicSet1
            
            '*** (Bef 20/05/05)
            'wo_Ctl.Height = go_RecClr!CsBtnHeightSet1
            'wo_Ctl.Width = go_RecClr!CsBtnWidthSet1
            'wo_Ctl.Top = wo_Ctl.Top + (go_RecClr!CsBtnTopDiffSet1 * Val(wo_Ctl.WhatsThisHelpID))
            'wo_Ctl.Left = IIF(wo_Ctl.Tag <> "", wi_ADCBtnWdthDiff + (Val(wo_Ctl.Tag) * wo_Ctl.Width), wo_Ctl.Left) + (go_RecClr!CsBtnLeftDiffSet1 * Val(wo_Ctl.Tag)) 'wo_Ctl.Left
            '*** (Bef 20/05/05)
            
            '*** (12/05/05)
            If go_RecClr!CsBtnHeightSet1 > 0 And go_RecClr!CsBtnWidthSet1 > 0 And go_RecClr!CsADCBtnHeight > 0 And go_RecClr!CsADCBtnWidth > 0 Then
              wo_Ctl.ReSizePic = False  '*** 07/07/05
              wo_Ctl.Height = go_RecClr!CsBtnHeightSet1
              wo_Ctl.Width = go_RecClr!CsBtnWidthSet1
              'wo_Ctl.Top = wo_Ctl.Top + (go_RecClr!CsBtnTopDiffSet1 * Val(wo_Ctl.BtnTopDiff))    'WhatsThisHelpID
              'wo_Ctl.Left = IIF(wo_Ctl.BtnLeftDiff <> "", wi_ADCBtnWdthDiff + (Val(wo_Ctl.BtnLeftDiff) * wo_Ctl.Width), wo_Ctl.Left) + (go_RecClr!CsBtnLeftDiffSet1 * Val(wo_Ctl.BtnLeftDiff)) 'wo_Ctl.Left
              If wo_Ctl.BtnTopDiff <> -1 Then wo_Ctl.Top = wo_Ctl.Top + (go_RecClr!CsBtnTopDiffSet1 * Val(wo_Ctl.BtnTopDiff))    'WhatsThisHelpID
              wo_Ctl.Left = IIF(wo_Ctl.BtnLeftDiff <> -1, wi_ADCBtnWdthDiff + (Val(wo_Ctl.BtnLeftDiff) * wo_Ctl.Width), wo_Ctl.Left) + (go_RecClr!CsBtnLeftDiffSet1 * Val(wo_Ctl.BtnLeftDiff)) 'wo_Ctl.Left
            Else
              '*** (Bef 07/07/05) Call AdjustBtnPicSz(wo_Ctl)
              wo_Ctl.ReSizePic = True '*** 07/07/05
            End If
            '*** (12/05/05)
          Else
            Call BtnSet2(wo_Ctl)
          End If
          
        ElseIf TypeOf wo_Ctl Is MWCTL_ADC Then
          wo_Ctl.BckColor = po_Frm.BackColor
          wo_Ctl.BtnPicture = wo_CsADCBtnPic
          wo_Ctl.BtnForeColor = Val(go_RecClr!CsADCBtnForClr)
          wo_Ctl.BtnDisableColor = Val(go_RecClr!CsADCBtnDisaClr)
          wo_Ctl.BtnMaskColor = Val(go_RecClr!CsADCBtnMaskClr)
          'wo_Ctl.Pic = po_Frm.Picture      '*** (Bef 2.14 Next Patch 1)
          '*** (Bef 20/05/05)
          'wo_Ctl.BtnHeight = go_RecClr!CsADCBtnHeight
          'wo_Ctl.BtnWidth = go_RecClr!CsADCBtnWidth
          'wo_Ctl.Top = ADC.Top + go_RecClr!CsADCDiffTop
          '*** (Bef 20/05/05)
          
          '*** (20/05/05)
          'If go_RecClr!CsADCBtnHeight > 0 And go_RecClr!CsADCBtnWidth > 0 Then
          If go_RecClr!CsBtnHeightSet1 > 0 And go_RecClr!CsBtnWidthSet1 > 0 And go_RecClr!CsADCBtnHeight > 0 And go_RecClr!CsADCBtnWidth > 0 Then
            wo_Ctl.ReSizeBtnPic = False   '*** (07/07/05)
            wo_Ctl.BtnHeight = go_RecClr!CsADCBtnHeight
            wo_Ctl.BtnWidth = go_RecClr!CsADCBtnWidth
            'wo_Ctl.Top = ADC.Top + go_RecClr!CsADCDiffTop
          Else
            wo_Ctl.ReSizeBtnPic = True   '*** (07/07/05)
            '*** (Bef 07/07/05) wo_Ctl.BtnPicture = AdjustBtnPicSize(wo_Ctl.BtnPicture, wo_Ctl.BtnHeight, wo_Ctl.BtnWidth)
          End If
          '*** (20/05/05)
        
        '*** (Jen 2.14 Next Patch 1)
        ElseIf TypeOf wo_Ctl Is Frame Then
          wo_Ctl.BackColor = Val(go_RecClr!CsFraBckClr)
          wo_Ctl.ForeColor = Val(go_RecClr!CsLblMainClr)
        '*** (Jen 2.14 Next Patch 1)
        
        ElseIf TypeOf wo_Ctl Is MWCTL_HLP Then
          wo_Ctl.BckColor = Val(go_RecClr!CsHlpBckClr)
          wo_Ctl.BckColorBkg = Val(go_RecClr!CsHlpBckClrBkg)
          wo_Ctl.BckColorFixed = Val(go_RecClr!CsHlpBckClrFix)
          wo_Ctl.BckColorSel = Val(go_RecClr!CsHlpBckClrSel)
          wo_Ctl.BckColorSelRow = Val(go_RecClr!CsHlpBckClrSelRow)
          
          wo_Ctl.ForColor = Val(go_RecClr!CsHlpForClr)
          wo_Ctl.ForColorFixed = Val(go_RecClr!CsHlpForClrFix)
          wo_Ctl.ForColorSel = Val(go_RecClr!CsHlpForClrSel)
          wo_Ctl.GrdColor = Val(go_RecClr!CsHlpGrdClr)
          wo_Ctl.GrdColorFixed = Val(go_RecClr!CsHlpGrdClrFix)
          
          If go_RecClr!CsGrdInset = "Y" Then
            wo_Ctl.GrdLinesFixed = flxGrdInset
          Else
            wo_Ctl.GrdLinesFixed = flxGrdFlat
          End If
          
          If go_RecClr!Cs3D = "Y" Then
            wo_Ctl.GrdAppearance = flxGrdApp3D
          Else
            wo_Ctl.GrdAppearance = flxGrdAppFlat
          End If
         
          wo_Ctl.Pic = wo_CsHlpPic
        ElseIf TypeOf wo_Ctl Is SSTab Then
          wo_Ctl.BackColor = Val(go_RecClr!CsTabBckClr)
          wo_Ctl.ForeColor = Val(go_RecClr!CsTabForClr)
        ElseIf TypeOf wo_Ctl Is TextBox Then
          wo_Ctl.BackColor = Val(go_RecClr!CsMedBckClr)
          wo_Ctl.ForeColor = Val(go_RecClr!CsMedForClr)
        ElseIf TypeOf wo_Ctl Is ImageCombo Then
          wo_Ctl.BackColor = Val(go_RecClr!CsMedBckClr)
          wo_Ctl.ForeColor = Val(go_RecClr!CsMedForClr)
        ElseIf TypeOf wo_Ctl Is ComboBox Then
          wo_Ctl.BackColor = Val(go_RecClr!CsMedBckClr)
          wo_Ctl.ForeColor = Val(go_RecClr!CsMedForClr)
        ElseIf TypeOf wo_Ctl Is Line Then
          wo_Ctl.BorderColor = Val(go_RecClr!CsLineBdrClr)
        ElseIf TypeOf wo_Ctl Is DriveListBox Then
          wo_Ctl.BackColor = Val(go_RecClr!CsMedBckClr)
          wo_Ctl.ForeColor = Val(go_RecClr!CsMedForClr)
        ElseIf TypeOf wo_Ctl Is MWCTL_GRP Then
          wo_Ctl.BckColor = Val(go_RecClr!CsMedBckClr)
          wo_Ctl.ForColor = Val(go_RecClr!CsMedForClr)
        ElseIf TypeOf wo_Ctl Is MWCTL_DST Then
          wo_Ctl.BckColor = Val(go_RecClr!CsDstBckClr)
          wo_Ctl.ForColor = Val(go_RecClr!CsDstForClr)
          wo_Ctl.ComboBckColor = Val(go_RecClr!CsMedBckClr)
          wo_Ctl.ComboForColor = Val(go_RecClr!CsMedForClr)
          'wo_Ctl.Pic = po_Frm.Picture
        
        ElseIf TypeOf wo_Ctl Is PictureBox Then
          If wo_Ctl.DataField = "" Then
            wo_Ctl.Picture = po_Frm.Picture
            wo_Ctl.BackColor = po_Frm.BackColor
            
            '*** (21/03/05)
            If go_RecClr!Cs3D = "N" And wo_Ctl.Appearance = 1 Then
              wo_Ctl.Appearance = 0
              
              If wo_Ctl.BorderStyle = 1 Then
                wo_Ctl.Top = wo_Ctl.Top + 15
                wo_Ctl.Left = wo_Ctl.Left + 15
              End If
            End If
            '*** (21/03/05)
          End If
          
          '*** (Bef 21/03/05)
          'If go_RecClr!Cs3D = "N" And wo_Ctl.Appearance = 1 Then
          '  wo_Ctl.Appearance = 0
          'End If
          '*** (Bef 21/03/05)
          
        ' ***** Manali 3.6.0 - Setting Colors for CheckBox and OptionButton
        ElseIf TypeOf wo_Ctl Is CheckBox Then
          wo_Ctl.BackColor = po_Frm.BackColor
          wo_Ctl.ForeColor = Val(go_RecClr!CsLblForClr)
        ElseIf TypeOf wo_Ctl Is OptionButton Then
          wo_Ctl.BackColor = po_Frm.BackColor
          wo_Ctl.ForeColor = Val(go_RecClr!CsLblForClr)
        ' ***** Manali 3.6.0 - Setting Colors for CheckBox and OptionButton
        End If
    Next
    
    ''PicBoxHsk.Visible = False: PicBox3D.Visible = False
    
  End If
  
  po_Frm.fra_keys.BorderStyle = 0
  
  Set wo_Ctl = Nothing
  Set wo_CsFlxPic = Nothing
  Set wo_CsBtnPicSet1 = Nothing
  Set wo_CsADCBtnPic = Nothing
  Set wo_CsHlpPic = Nothing

End Sub
Public Sub BtnSet2(ByVal po_Ctl As Control)
  If (go_RecClr Is Nothing) Then Exit Sub   '*** (Jen 2.13)
  
  ' , ByVal po_CsBtnPicSet2 As Picture      '(Bef 15/06/05)
  po_Ctl.ForeColor = Val(go_RecClr!CsBtnForClrSet2)
  po_Ctl.DisableColor = Val(go_RecClr!CsBtnDisaClrSet2)
  po_Ctl.MaskColor = Val(go_RecClr!CsBtnMaskClrSet2)
  'po_Ctl.Pic = po_CsBtnPicSet2   '*** (Bef 15/06/05)
  'po_Ctl.Pic = go_MDIFrm.ImgDefBtnSet2    '*** 15/06/05   (Bef 18/06/05)
  po_Ctl.Pic = go_MDIFrm.CMDCAN.Pic    '*** 18/06/05
  
  If go_RecClr!CsBtnHeightSet2 > 0 And go_RecClr!CsBtnWidthSet2 > 0 Then
    po_Ctl.ReSizePic = False  '*** 07/07/05
    po_Ctl.Height = go_RecClr!CsBtnHeightSet2
    po_Ctl.Width = go_RecClr!CsBtnWidthSet2
    
    po_Ctl.Top = po_Ctl.Top + (go_RecClr!CsBtnTopDiffSet2 * Val(po_Ctl.BtnTopDiff))   'WhatsThisHelpID
    po_Ctl.Left = po_Ctl.Left + (go_RecClr!CsBtnLeftDiffSet2 * Val(po_Ctl.BtnLeftDiff))
  Else
    po_Ctl.ReSizePic = True   '*** 07/07/05
    '*** (Bef 07/07/05) Call AdjustBtnPicSz(po_Ctl)
  End If
End Sub

Public Function IsValidPath(ByVal fs_Pth As String) As Boolean
  '*** (13/08/05)
  On Error GoTo FileErr
  'On Error Resume Next
  
  IsValidPath = False
  
  If Trim(fs_Pth) <> "" And Dir(fs_Pth) <> "" Then
    IsValidPath = True
  End If
  Exit Function

FileErr:
  ''If Err.Number = 52 Then
  ''    '*** Bad file name or number
  ''    GetValidPath = ""
  ''End If
  'GetValidPath = ""
  IsValidPath = False
  '*** (13/08/05)
End Function
'*** Jenny Color (For Setting System Colors)



