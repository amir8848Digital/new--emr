VERSION 5.00
Object = "{27395F88-0C0C-101B-A3C9-08002B2F49FB}#1.1#0"; "picclp32.ocx"
Begin VB.UserControl MWCTL_BTN1 
   AutoRedraw      =   -1  'True
   BackStyle       =   0  'Transparent
   ClientHeight    =   1560
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   3210
   DefaultCancel   =   -1  'True
   DrawStyle       =   2  'Dot
   FillColor       =   &H008080FF&
   KeyPreview      =   -1  'True
   LockControls    =   -1  'True
   MaskColor       =   &H0080FFFF&
   MaskPicture     =   "MWCTL_BTN1.ctx":0000
   ScaleHeight     =   1560
   ScaleWidth      =   3210
   Begin PicClip.PictureClip PicClip 
      Left            =   930
      Top             =   720
      _ExtentX        =   2143
      _ExtentY        =   1244
      _Version        =   393216
   End
   Begin VB.CommandButton Cmd 
      Height          =   255
      Left            =   2670
      MaskColor       =   &H00000000&
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   1200
      Width           =   285
   End
   Begin VB.Label Lbl 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Label"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   465
      Left            =   90
      TabIndex        =   0
      Top             =   60
      UseMnemonic     =   0   'False
      Width           =   885
   End
End
Attribute VB_Name = "MWCTL_BTN1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

'Dim m_AccessKey         As String * 1

'*** (Bef 14/09/05) (This is required to give a moving effect while clicking. but this stops the click event from being triggered while clicking at some points on the button (near the top left) )
'Dim mi_Top As Integer
'Dim mi_Left As Integer
'*** (Bef 14/09/05) (This is required to give a moving effect while clicking. but this stops the click event from being triggered while clicking at some points on the button (near the top left) )

'*** (14/09/05)
Dim mb_FntBld As Boolean
'*** (14/09/05)

Dim mo_ForeColor As OLE_COLOR
Dim mo_DisableColor As OLE_COLOR
'*** 07/07/05
Dim mb_ReSizePic As Boolean
'*** 07/07/05

'*** 13/06/05
Dim mi_TopDiff As Integer
Dim mi_LeftDiff As Integer
'*** 13/06/05

'*** (Bef 26/07/05)  Private WithEvents mFont As StdFont

Event Click()
Event CmdGotFocus()
Event CmdLostFocus()
Event CmdKeyDown(ByRef KeyCode As Integer, ByRef Shift As Integer)
Event CmdKeyUp(ByRef KeyCode As Integer, ByRef Shift As Integer)
Event CmdMouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
Event CmdMouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)


'***************** Props Begin ************************************************************'
Public Property Get Enabled() As Boolean
Attribute Enabled.VB_ProcData.VB_Invoke_Property = ";Additional"
    Enabled = UserControl.Enabled
End Property
Public Property Let Enabled(ByVal vNewValue As Boolean)
    UserControl.Enabled = vNewValue
    Cmd.Enabled = vNewValue   '*** (24/03/06)
    'Lbl.Enabled = vNewValue

    If vNewValue = False Then Lbl.ForeColor = mo_DisableColor      '&HB4B4B4
    If vNewValue = True Then Lbl.ForeColor = mo_ForeColor          '&HC0C0C0

    PropertyChanged "Enabled"
End Property
Public Property Get CausesValidation() As Boolean
    CausesValidation = Cmd.CausesValidation
End Property
Public Property Let CausesValidation(ByVal vNewValue As Boolean)
    Cmd.CausesValidation = vNewValue
    PropertyChanged "CausesValidation"
End Property
Public Property Get hwnd() As Boolean
    hwnd = UserControl.hwnd
End Property
Public Property Get Caption() As String
Attribute Caption.VB_ProcData.VB_Invoke_Property = ";Additional"
    'Caption = Cmd.Caption    '*** (Before 16/05/05)
    
    '*** (Bef 24/03/06)
    'Caption = Lbl.Caption    '*** 16/05/05
    '*** (Bef 24/03/06)
    
    '*** (24/03/06)
    If Cmd.Picture = 0 Then
      Caption = Cmd.Caption
    Else
      Caption = Lbl.Caption    '*** 16/05/05
    End If
    '*** (24/03/06)
End Property
Public Property Let Caption(ByVal vNewValue As String)
    '*** (Bef 16/05/05)
    'If Ambient.UserMode Then
    '  Lbl.UseMnemonic = False
    '  'Lbl.Caption = Replace(vNewValue, "&", "-")     (Code Bef 02/04/05)
    '
    '  '*** 02/04/05
    '  Dim wi_Pos As Integer
    '  wi_Pos = InStr(1, vNewValue, "&")
    '  Lbl.Caption = Replace(vNewValue, "&", "")
    '  If wi_Pos > 0 And wi_Pos <= Len(Lbl.Caption) Then
    '    Lbl.Caption = LCase(Lbl.Caption)
    '    Lbl.Caption = IIf(wi_Pos > 1, Mid(Lbl.Caption, 1, wi_Pos - 1), "") & UCase(Mid(Lbl.Caption, wi_Pos, 1)) & IIf(wi_Pos < Len(Lbl.Caption), Mid(Lbl.Caption, wi_Pos + 1, Len(Lbl.Caption) - wi_Pos), "") 'Replace(vNewValue, "&", "-")
    '  End If
    '  '*** 02/04/05
    'Else
    '  Lbl.UseMnemonic = True
    '  Lbl.Caption = vNewValue
    'End If
    'Cmd.Caption = vNewValue
    'PropertyChanged "Caption"
    ''### SetAccessKey
    '*** (Bef 16/05/05)
    
    '*** (16/05/05)
    Lbl.Left = 0
    Lbl.AutoSize = False
    Lbl.WordWrap = True
    Lbl.Width = UserControl.Width
    Lbl.Height = 0 'UserControl.Height
    
    Lbl.UseMnemonic = True
    Lbl.Caption = vNewValue
    
    Lbl.AutoSize = True
    Lbl.Top = (UserControl.Height - Lbl.Height) / 2
    If Lbl.Top < 0 Then Lbl.Top = 0
    
    If Ambient.UserMode Then
      Dim wi_Pos As Integer
      wi_Pos = InStr(1, vNewValue, "&")
      
      '*** (Bef 24/03/06)
      'If wi_Pos > 0 And wi_Pos < Len(Trim(vNewValue)) Then UserControl.AccessKeys = Mid(Lbl.Caption, wi_Pos + 1, 1)
      '*** (Bef 24/03/06)
      
      '*** (24/03/06)
      If Cmd.Picture <> 0 Then
        If wi_Pos > 0 And wi_Pos < Len(Trim(vNewValue)) Then UserControl.AccessKeys = Mid(Lbl.Caption, wi_Pos + 1, 1)
      End If
      '*** (24/03/06)
    End If
    
    '*** (24/03/06)
    If Cmd.Picture = 0 Then
      Lbl.Enabled = False
      Cmd.Caption = vNewValue
    Else
      Lbl.Enabled = True
      Cmd.Caption = ""
    End If
    Call LblPositioning
    '*** (24/03/06)
    
    PropertyChanged "Caption"
    '*** (16/05/05)
End Property

'*** 07/07/05
Public Property Get ReSizePic() As Boolean
Attribute ReSizePic.VB_ProcData.VB_Invoke_Property = ";Additional"
    ReSizePic = mb_ReSizePic
End Property
Public Property Let ReSizePic(ByVal vNewValue As Boolean)
    mb_ReSizePic = vNewValue
    Call SetCtrlPic
    PropertyChanged "ReSizePic"
End Property
'*** 07/07/05

Public Property Get Pic() As Picture
Attribute Pic.VB_ProcData.VB_Invoke_Property = ";Additional"
    Set Pic = Cmd.Picture
    'Set Pic = UserControl.Picture    '(Bef 07/07/05)
End Property
Public Property Let Pic(ByVal vNewValue As Picture)
    '*** 13/06/05
    If (vNewValue Is Nothing) Then
      UserControl.BackStyle = 1
      'UserControl.BorderStyle = 1
    Else
      If vNewValue = 0 Then
        UserControl.BackStyle = 1
      Else
        UserControl.BackStyle = 0
      End If
      'UserControl.BorderStyle = 0
    End If
    '*** 13/06/05

    '*** 07/07/05
    Set Cmd.Picture = vNewValue
    Call SetCtrlPic
    PropertyChanged "Pic"
    '*** 07/07/05

    '*** (Bef 07/07/05)
    'Set UserControl.Picture = vNewValue
    'Set UserControl.MaskPicture = vNewValue
    'PropertyChanged "Pic"
    '*** (Bef 07/07/05)
End Property
Public Property Set Pic(ByVal vNewValue As Picture)
    '*** 13/06/05
    If (vNewValue Is Nothing) Then
      UserControl.BackStyle = 1
      'UserControl.BorderStyle = 1
    Else
      If vNewValue = 0 Then
        UserControl.BackStyle = 1
      Else
        UserControl.BackStyle = 0
      End If
      'UserControl.BorderStyle = 0
    End If
    '*** 13/06/05
    
    '*** 07/07/05
    Set Cmd.Picture = vNewValue
    Call SetCtrlPic
    PropertyChanged "Pic"
    '*** 07/07/05
    
    '*** (Bef 07/07/05)
    'Set UserControl.Picture = vNewValue
    'Set UserControl.MaskPicture = vNewValue
    'PropertyChanged "Pic"
    '*** (Bef 07/07/05)
End Property
Public Property Get Visible() As Boolean
    Visible = UserControl.Extender.Visible
End Property
Public Property Let Visible(ByVal vNewValue As Boolean)
    UserControl.Extender.Visible = vNewValue
End Property
Public Property Let BackColor(ByVal vNewValue As OLE_COLOR)
    UserControl.BackColor = vNewValue
    Cmd.BackColor = vNewValue     '*** (24/03/06)
    PropertyChanged "BackColor"
End Property
Public Property Get BackColor() As OLE_COLOR
Attribute BackColor.VB_ProcData.VB_Invoke_Property = ";Additional"
    BackColor = UserControl.BackColor
End Property
Public Property Let MaskColor(ByVal vNewValue As OLE_COLOR)
    UserControl.MaskColor = vNewValue
    PropertyChanged "MaskColor"
End Property
Public Property Get MaskColor() As OLE_COLOR
Attribute MaskColor.VB_ProcData.VB_Invoke_Property = ";Additional"
    MaskColor = UserControl.MaskColor
End Property
Public Property Let ForeColor(ByVal vNewValue As OLE_COLOR)
    Lbl.ForeColor = vNewValue
    mo_ForeColor = vNewValue
    
    PropertyChanged "ForeColor"
End Property
Public Property Get ForeColor() As OLE_COLOR
Attribute ForeColor.VB_ProcData.VB_Invoke_Property = ";Additional"
    ForeColor = Lbl.ForeColor
End Property
Public Property Let DisableColor(ByVal vNewValue As OLE_COLOR)
    mo_DisableColor = vNewValue
    PropertyChanged "DisableColor"
End Property
Public Property Get DisableColor() As OLE_COLOR
Attribute DisableColor.VB_ProcData.VB_Invoke_Property = ";Additional"
    DisableColor = mo_DisableColor
End Property
Public Property Let LblFont(vNewValue As StdFont)
    ''ByVal vNewValue As Font

    '*** (Bef 26/07/05)
    ''Set mFont = vNewValue
    'With mFont
    '  .Bold = vNewValue.Bold
    '  .Italic = vNewValue.Italic
    '  .Name = vNewValue.Name
    '  '.Size = vNewValue.Size
    'End With
    '
    ''With Lbl
    ''  .Font.Bold = vNewValue.Bold
    ''  .Font.Italic = vNewValue.Italic
    ''  .Font.Name = vNewValue.Name
    ''  .Font.Size = vNewValue.Size
    ''End With
    '*** (Bef 26/07/05)
    
    '*** (26/07/05)
    Set Lbl.Font = vNewValue
    Set Cmd.Font = vNewValue     '*** (24/03/06)
    mb_FntBld = vNewValue.Bold   '*** 14/09/05
    
    Call LblPositioning
    '*** (26/07/05)
    
    PropertyChanged "LblFont"
End Property
Public Property Set LblFont(vNewValue As StdFont)
    '*** (26/07/05)
    Set Lbl.Font = vNewValue
    Set Cmd.Font = vNewValue     '*** (24/03/06)
    mb_FntBld = vNewValue.Bold   '*** 14/09/05
    
    Call LblPositioning
    
    PropertyChanged "LblFont"
    '*** (26/07/05)
End Property
Public Property Get LblFont() As StdFont
Attribute LblFont.VB_ProcData.VB_Invoke_Property = ";Additional"
    'Set LblFont = mFont      '*** (Bef 26/07/05)
    
    Set LblFont = Lbl.Font    '*** (26/07/05)
End Property
'*** (Bef 26/07/05)
'Public Property Let FontSz(ByVal vNewValue As Integer)
'    Lbl.Font.Size = vNewValue
'    PropertyChanged "FontSz"
'
'    '*** 16/05/05
'    Lbl.Left = 0
'    Lbl.AutoSize = False
'    Lbl.WordWrap = True
'    Lbl.Width = UserControl.Width
'    Lbl.Height = 0 'UserControl.Height
'    Lbl.AutoSize = True
'    Lbl.Top = (UserControl.Height - Lbl.Height) / 2
'    If Lbl.Top < 0 Then Lbl.Top = 0
'    '*** 16/05/05
'End Property
'Public Property Get FontSz() As Integer
'    FontSz = Lbl.Font.Size
'End Property
'Public Property Let FntBold(ByVal vNewValue As Boolean)
'    Lbl.Font.Bold = vNewValue
'    PropertyChanged "FntBold"
'End Property
'Public Property Get FntBold() As Boolean
'    FntBold = Lbl.Font.Bold
'End Property
'*** (Bef 26/07/05)
Public Property Let BtnCancel(ByVal vNewValue As Boolean)
    Cmd.Cancel = vNewValue
    PropertyChanged "BtnCancel"
End Property
Public Property Get BtnCancel() As Boolean
Attribute BtnCancel.VB_ProcData.VB_Invoke_Property = ";Additional"
    BtnCancel = Cmd.Cancel
End Property
Public Property Let BtnDefault(ByVal vNewValue As Boolean)
    Cmd.Default = vNewValue
    PropertyChanged "BtnDefault"
End Property
Public Property Get BtnDefault() As Boolean
Attribute BtnDefault.VB_ProcData.VB_Invoke_Property = ";Additional"
    BtnDefault = Cmd.Default
End Property
Public Property Get ToolTipTxt() As String
    ToolTipTxt = UserControl.Extender.ToolTipText
End Property
Public Property Let ToolTipTxt(ByVal vNewValue As String)
    UserControl.Extender.ToolTipText = vNewValue
    Lbl.ToolTipText = vNewValue
    Cmd.ToolTipText = vNewValue   '*** (24/03/06)
End Property

'*** 13/06/05
Public Property Let BtnLeftDiff(ByVal vNewValue As Integer)
    mi_LeftDiff = vNewValue
    PropertyChanged "BtnLeftDiff"
End Property
Public Property Get BtnLeftDiff() As Integer
Attribute BtnLeftDiff.VB_ProcData.VB_Invoke_Property = ";Additional"
    BtnLeftDiff = mi_LeftDiff
End Property
Public Property Let BtnTopDiff(ByVal vNewValue As Integer)
    mi_TopDiff = vNewValue
    PropertyChanged "BtnTopDiff"
End Property
Public Property Get BtnTopDiff() As Integer
Attribute BtnTopDiff.VB_ProcData.VB_Invoke_Property = ";Additional"
    BtnTopDiff = mi_TopDiff
End Property
'*** 13/06/05


'Property Get ParentRef() As Object
'    Set ParentRef = Parent
'
'    'If TypeOf Parent Is MWCTL_ADC Then
'    '    Set ParentRef = Parent.ParentRef
'    'Else
'    '    Set ParentRef = Parent
'    'End If
'End Property


'***************** Props End ************************************************************'
'***************** Events Begin ************************************************************'

Private Sub UserControl_Initialize()
    '*** (Bef 26/07/05)
    'Set mFont = New StdFont
    'With mFont
    '   .Bold = True
    '   .Name = "MS Sans Serif"
    '   .Size = 8
    'End With
    '
    'Set Lbl.Font = mFont
    '*** (Bef 26/07/05)
End Sub
'*** (Bef 26/07/05)
'Private Sub mFont_FontChanged(ByVal PropertyName As String)
'    Set Lbl.Font = mFont
'    Lbl.Refresh
'End Sub
'*** (Bef 26/07/05)
Private Sub UserControl_Click()
    If UserControl.Enabled = True Then
      RaiseEvent Click
    End If
End Sub
Private Sub UserControl_DblClick()
    If UCase(UserControl.ParentControls(0).Name) = "ADC" Then Exit Sub    '*** 01/07/06
    '***  23/08/05
    If UserControl.Enabled = True Then
      RaiseEvent Click
    End If
    '***  23/08/05
End Sub
Private Sub Lbl_Click()
    If UserControl.Enabled = True Then
      RaiseEvent Click
    End If
    
'    If TypeOf ParentRef Is MWCTL_ADC Then
'      'Debug.Print "In Lbl_Click: ParentRef.ValFailed= " + ParentRef.ValFailed   '*** Tmp By Jenny
'      Debug.Print "In Lbl_Click: ParentRef.OptSel= " + ParentRef.OptSel   '*** Tmp By Jenny
'
'      'If ParentRef.BtnValFailed = "" Then
'      If ParentRef.ValFailed = "" Or (ParentRef.OptSel = "X") Then
'        Debug.Print "Yes"
'        RaiseEvent Click
'      End If
'      '''If UCase(Parent.Name) = UCase("MWCTL_ADC") Then
'
'    'Else
'    '  RaiseEvent Click
'    End If
End Sub
Private Sub Lbl_DblClick()
    If UCase(UserControl.ParentControls(0).Name) = "ADC" Then Exit Sub    '*** 01/07/06
    '***  23/08/05
    If UserControl.Enabled = True Then
      RaiseEvent Click
    End If
    '***  23/08/05
End Sub
Private Sub Cmd_Click()
    If UserControl.Enabled = True Then
      RaiseEvent Click
    End If
End Sub
Private Sub Cmd_GotFocus()
    Call ShowFocusRectangle
    If UserControl.Enabled = True Then
      'Call SetFocus
      'UserControl.SetFocus
      RaiseEvent CmdGotFocus
    End If
End Sub
Private Sub Cmd_LostFocus()
    '####
    'UserControl.AutoRedraw = True      '*** For UserControl.Cls to work AutoRedraw Should Be True
    UserControl.Cls
    'UserControl.AutoRedraw = False
    '####
    
    ''UserControl.BorderStyle = 0
    RaiseEvent CmdLostFocus
End Sub
Private Sub UserControl_ExitFocus()
    'UserControl.Cls
End Sub
Private Sub UserControl_LostFocus()
    'UserControl.Cls
End Sub
Private Sub UserControl_Resize()
    '*** (Bef 16/05/05)
    'Lbl.Left = 0
    'Lbl.Top = (UserControl.Height / 2) / 4
    'Lbl.Height = UserControl.Height
    'Lbl.Width = UserControl.Width
    '*** (Bef 16/05/05)
    
    'Call LblPositioning     '*** (Bef 24/03/06)
    
    '*** 07/07/05
    Call SetCtrlPic
    'Cmd.Left = UserControl.Width + 1  '*** (Bef 24/03/06)
    '*** 07/07/05
End Sub
Private Sub UserControl_Show()
    Lbl.ZOrder (0)
End Sub
Private Sub Lbl_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    'If Ambient.UserMode Then
    '  Exit Property
    'End If
        
    '*** (Bef 14/09/05) (This is required to give a moving effect while clicking. but this stops the click event from being triggered while clicking at some points on the button (near the top left) )
    'If UserControl.Extender.Left <> mi_Left And UserControl.Extender.Left - 10 <> mi_Left Then mi_Left = UserControl.Extender.Left
    'If UserControl.Extender.Top <> mi_Top And UserControl.Extender.Top - 10 <> mi_Top Then mi_Top = UserControl.Extender.Top
    'UserControl.Extender.Left = mi_Left + 10
    'UserControl.Extender.Top = mi_Top + 10
    '*** (Bef 14/09/05) (This is required to give a moving effect while clicking. but this stops the click event from being triggered while clicking at some points on the button (near the top left) )
    
    
    '*** (14/09/05)
    'Lbl.ForeColor = vbRed    '*** (04/10/05) (Testing)
    'Lbl.FontBold = Not (mb_FntBld) 'False    '*** (Bef 04/10/05)    (enabling this line doesn't produce the click sometimes)
    
    'Call LblPositioning   '*** 04/10/05
    '*** (14/09/05)
    
    UserControl.Cls
    
    If UserControl.Enabled = True Then
      RaiseEvent CmdMouseDown(Button, Shift, X, Y)
    End If

End Sub
Private Sub Lbl_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    '*** (Bef 14/09/05)
    ''UserControl.Extender.Left = mi_Left
    ''UserControl.Extender.Top = mi_Top
    '*** (Bef 14/09/05)
    
    '*** 14/09/05 (This is required to give a moving effect while clicking. but this stops the click event from being triggered while clicking at some points on the button (near the top left, Y will be -15 in this case) )
    'If mi_Left <> -999 Then UserControl.Extender.Left = mi_Left
    'If mi_Top <> -999 Then UserControl.Extender.Top = mi_Top
    '*** 14/09/05 (This is required to give a moving effect while clicking. but this stops the click event from being triggered while clicking at some points on the button (near the top left, Y will be -15 in this case) )
    
    '*** (14/09/05)
    'Lbl.ForeColor = vbGreen   '*** (04/10/05) (Testing)
    'Lbl.FontBold = mb_FntBld    'True   '*** (Bef 04/10/05)    (enabling this line doesn't produce the click sometimes)
    
    'Call LblPositioning  '*** 04/10/05
    
    '*** 04/10/05 (Trial)
    ''Lbl.Width = UserControl.Width
    ''Lbl.Height = UserControl.Height

    ''If X < Lbl.Left Or Lbl.Left + Lbl.Width > X Or Y < Lbl.Top Or Lbl.Top + Lbl.Height > Y Then
    ''  RaiseEvent Click
    ''End If
    '*** 04/10/05 (Trial)
    
    '*** (14/09/05)
    
    Call ShowFocusRectangle

    If UserControl.Enabled = True Then
      RaiseEvent CmdMouseUp(Button, Shift, X, Y)
    End If
    
    'If TypeOf ParentRef Is MWCTL_ADC Then
    '  If ParentRef.ValFailed = "" Then
    '    RaiseEvent CmdMouseUp(Button, Shift, X, Y)
    '  End If
    'End If
    
End Sub

'*** (14/09/05)
Private Sub UserControl_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Lbl.FontBold = Not (mb_FntBld)
    Call LblPositioning  '*** 04/10/05
    
    UserControl.Cls
    
    If UserControl.Enabled = True Then
      RaiseEvent CmdMouseDown(Button, Shift, X, Y)
    End If
End Sub
Private Sub UserControl_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Lbl.FontBold = mb_FntBld
    Call LblPositioning  '*** 04/10/05
    
    Call ShowFocusRectangle
    If UserControl.Enabled = True Then
      RaiseEvent CmdMouseUp(Button, Shift, X, Y)
    End If
End Sub
'*** (14/09/05)

Private Sub Cmd_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    RaiseEvent CmdMouseDown(Button, Shift, X, Y)
End Sub
Private Sub Cmd_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    RaiseEvent CmdMouseUp(Button, Shift, X, Y)
End Sub
Private Sub Cmd_KeyDown(KeyCode As Integer, Shift As Integer)
    RaiseEvent CmdKeyDown(KeyCode, Shift)
End Sub
Private Sub Cmd_KeyUp(KeyCode As Integer, Shift As Integer)
    RaiseEvent CmdKeyUp(KeyCode, Shift)
End Sub
Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
    mo_DisableColor = PropBag.ReadProperty("DisableColor", &HB4B4B4)    '*** Should be read before 'Enabled' Property
    ForeColor = PropBag.ReadProperty("ForeColor", &H0&)                 '*** Should be read before 'Enabled' Property
    mo_ForeColor = ForeColor                                            '*** Should be read before 'Enabled' Property
    
    '*** (24/03/06)
    Pic = PropBag.ReadProperty("Pic", LoadPicture())                    '*** Should be read before 'Caption' Property
    '*** (24/03/06)
    
    Enabled = PropBag.ReadProperty("Enabled", True)
    CausesValidation = PropBag.ReadProperty("CausesValidation", True)
    BackColor = PropBag.ReadProperty("BackColor", &H80000005)
    MaskColor = PropBag.ReadProperty("MaskColor", &HFFFFFF)
    Caption = PropBag.ReadProperty("Caption", "")
        
    'UserControl.Picture = PropBag.ReadProperty("Pic", LoadPicture())
    '*** (Bef 13/06/05) Pic = PropBag.ReadProperty("Pic", UserControl.Picture)
    
    '*** (Bef 24/03/06) Pic = PropBag.ReadProperty("Pic", LoadPicture())
    
    '*** 07/07/05
    ReSizePic = PropBag.ReadProperty("ReSizePic", True)
    '*** 07/07/05
    
    LblFont = PropBag.ReadProperty("LblFont", Lbl.Font)
    
    '*** (Bef 26/07/05)
    'FontSz = PropBag.ReadProperty("FontSz", Lbl.Font.Size)
    'FntBold = PropBag.ReadProperty("FntBold", Lbl.Font.Bold)
    '*** (Bef 26/07/05)
    BtnCancel = PropBag.ReadProperty("BtnCancel", False)
    BtnDefault = PropBag.ReadProperty("BtnDefault", False)
    
    '*** 13/06/05
    mi_LeftDiff = PropBag.ReadProperty("BtnLeftDiff", -1)
    mi_TopDiff = PropBag.ReadProperty("BtnTopDiff", -1)
    '*** 13/06/05
    
    '*** (Bef 14/09/05)
    ''mi_Top = UserControl.Extender.Top
    ''mi_Left = UserControl.Extender.Left
    '*** (Bef 14/09/05)
    
    '*** 14/09/05 (This is required to give a moving effect while clicking. but this stops the click event from being triggered while clicking at some points on the button (near the top left) )
    'mi_Top = -999
    'mi_Left = -999
    '*** 14/09/05 (This is required to give a moving effect while clicking. but this stops the click event from being triggered while clicking at some points on the button (near the top left) )
    
    
    Call LblPositioning     '*** 14/06/05
    
    '### m_AccessKey = PropBag.ReadProperty("AccessKey", vbNullString)
    '### UserControl.AccessKeys = m_AccessKey
    'LblUnderLine.Caption = Caption
End Sub

'Write property values to storage
Private Sub UserControl_WriteProperties(PropBag As PropertyBag)
    Dim wo_LblFont As New StdFont
    With wo_LblFont
       .Bold = True
       .Name = "MS Sans Serif"
       .Size = 8
    End With

    Call PropBag.WriteProperty("Enabled", Enabled, True)
    Call PropBag.WriteProperty("CausesValidation", Cmd.CausesValidation, True)
    Call PropBag.WriteProperty("BackColor", UserControl.BackColor, &H80000005)
    Call PropBag.WriteProperty("MaskColor", UserControl.MaskColor, &HFFFFFF)
    Call PropBag.WriteProperty("ForeColor", Lbl.ForeColor, &H0&)
    Call PropBag.WriteProperty("DisableColor", mo_DisableColor, &HB4B4B4)
    
    'Call PropBag.WriteProperty("Caption", Cmd.Caption, "")    '*** (Before 16/05/05)
    Call PropBag.WriteProperty("Caption", Lbl.Caption, "")     '*** (16/05/05)

    '*** (Bef 07/07/05) Call PropBag.WriteProperty("Pic", UserControl.Picture, LoadPicture())
    '*** 07/07/05
    Call PropBag.WriteProperty("Pic", Cmd.Picture, LoadPicture())
    Call PropBag.WriteProperty("ReSizePic", mb_ReSizePic, True)
    '*** 07/07/05
    
    '*** (Bef 26/07/05)
    'Call PropBag.WriteProperty("LblFont", Lbl.Font, mFont)
    'Call PropBag.WriteProperty("FontSz", Lbl.Font.Size, 8)
    'Call PropBag.WriteProperty("FntBold", Lbl.Font.Bold, True)
    '*** (Bef 26/07/05)
    
    Call PropBag.WriteProperty("LblFont", Lbl.Font, wo_LblFont)   '*** (26/07/05)
    'Call PropBag.WriteProperty("LblFont", Lbl.Font)
    
    Call PropBag.WriteProperty("BtnCancel", Cmd.Cancel, False)
    Call PropBag.WriteProperty("BtnDefault", Cmd.Default, False)
    
    '*** 13/06/05
    Call PropBag.WriteProperty("BtnLeftDiff", mi_LeftDiff, -1)
    Call PropBag.WriteProperty("BtnTopDiff", mi_TopDiff, -1)
    '*** 13/06/05
    
    '### Call PropBag.WriteProperty("AccessKey", m_AccessKey, vbNullString)
End Sub

'************************** Events End ******************************************************'
'***************** Methods Begin ************************************************************'

Public Sub SetFocus()
    If UserControl.Extender.Enabled = True Then
      UserControl.Extender.SetFocus
    End If
End Sub
Public Sub Refresh()
    UserControl.Refresh
End Sub
Public Sub ClrFocus()
    UserControl.Cls
End Sub
'Public Sub SetToolTip()
'    '*** commented on 22/03/05
'    '*** sets the tooltip from Med to txt, Called by Adc.InitRtn
'    Cmd.ToolTipText = UserControl.Extender.ToolTipText
'
'    Lbl.ToolTipText = UserControl.Extender.ToolTipText
'    '*** commented on 22/03/05
'End Sub
Private Sub ShowFocusRectangle()
  ''UserControl.BorderStyle = 1
  
  Dim wi_OrigScaleMode As Integer
  wi_OrigScaleMode = ScaleMode
  Dim CX, CY                        '*** Declare variables
  ScaleMode = 3                     '*** Set ScaleMode to pixels.
  'DrawWidth = 1                    '*** Set DrawWidth.
  'ForeColor = QBColor(50 Mod 15)   '*** Set foreground color.
  'UserControl.FillColor = ForeColor    '*** This property is ignored if FillStyle= Transparent
    
  '*** (Bef 26/07/05)
  'CX = ScaleWidth / 2               '*** Get horizontal center.
  'CY = ScaleHeight / 2              '*** Get vertical center.
  'Line (4, 4)-(CX * 1.99, CY * 1.99), , B
  '*** (Bef 26/07/05)
  
  '*** (26/07/05)
  CX = ScaleWidth - 4              '*** Get horizontal center.
  CY = ScaleHeight - 4             '*** Get vertical center.
  Line (4, 4)-(CX, CY), , B
  '*** (26/07/05)
  
  'Line (4, 4)-(CX * 1.99, CY * 1.99), , BF
  ScaleMode = wi_OrigScaleMode
End Sub
Private Sub LblPositioning()
  '*** (Bef 24/03/06)
  ''*** 16/05/05
  'Lbl.Left = 0
  'Lbl.AutoSize = False
  'Lbl.WordWrap = True
  'Lbl.Width = UserControl.Width
  'Lbl.Height = 0 'UserControl.Height
  'Lbl.AutoSize = True
  'Lbl.Top = (UserControl.Height - Lbl.Height) / 2
  'If Lbl.Top < 0 Then Lbl.Top = 0
  ''*** 16/05/05
  '*** (Bef 24/03/06)
    
    
  '*** (24/03/06)
  If Cmd.Picture = 0 Then
    Lbl.Left = UserControl.Width + 1
    Cmd.Width = UserControl.Width
    Cmd.Height = UserControl.Height
    Cmd.Left = 0
    Cmd.Top = 0
  Else
    '*** 16/05/05
    Cmd.Left = UserControl.Width + 1
    
    Lbl.Left = 0
    Lbl.AutoSize = False
    Lbl.WordWrap = True
    Lbl.Width = UserControl.Width
    Lbl.Height = 0 'UserControl.Height
    Lbl.AutoSize = True
    Lbl.Top = (UserControl.Height - Lbl.Height) / 2
    If Lbl.Top < 0 Then Lbl.Top = 0
    '*** 16/05/05
  End If
  '*** (24/03/06)
    
End Sub
Private Sub SetCtrlPic()
  '*** (24/03/06)
  If Cmd.Picture = 0 Then
    Lbl.Enabled = False
    Cmd.Caption = Lbl.Caption
  Else
    Lbl.Enabled = True
    Cmd.Caption = ""
  End If
  Call LblPositioning
  '*** (24/03/06)
  
  '*** 07/07/05
  If Ambient.UserMode = False Or ReSizePic = True Then
    Set PicClip.Picture = Pic
    
    If PicClip.Picture = 0 Then
      Set UserControl.Picture = LoadPicture()
      Set UserControl.MaskPicture = LoadPicture()
      Exit Sub
    End If
    '*** Get X and Y coordinates of the clipping region.
    PicClip.ClipX = 0
    PicClip.ClipY = 0
    
    '*** Set the area of the clipping region (in pixels).
    PicClip.ClipHeight = PicClip.Height
    PicClip.ClipWidth = PicClip.Width
    
    '*** Set the destination area to fill the Button.
    '*** For conversion from Twip to Pixel divide by 15.
    If PicClip.StretchX <> UserControl.Width / 15 Then PicClip.StretchX = UserControl.Width / 15
    If PicClip.StretchY <> UserControl.Height / 15 Then PicClip.StretchY = UserControl.Height / 15
    
    '*** Assign the clipped bitmap to the Button.
    Set UserControl.Picture = PicClip.Clip
    Set UserControl.MaskPicture = PicClip.Clip
    Set PicClip.Picture = LoadPicture()
  Else
    Set UserControl.Picture = Pic
    Set UserControl.MaskPicture = Pic
  End If
  '*** 07/07/05
End Sub


'***************** Methods End ************************************************************'

