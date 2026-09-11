VERSION 5.00
Begin VB.UserControl MWCTL_REP 
   ClientHeight    =   3600
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4800
   PropertyPages   =   "MWCTL_REP.ctx":0000
   ScaleHeight     =   3600
   ScaleWidth      =   4800
   Begin MwfCtl.MWCTL_MED TXT_REP 
      Height          =   495
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   1605
      _ExtentX        =   2831
      _ExtentY        =   873
      MaxLength       =   0
   End
End
Attribute VB_Name = "MWCTL_REP"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Dim m_IdName As String
Dim m_cmpstr As String
Event FldWhen(ByVal IdName As String, Cancel As Boolean, ErrMsg As String)
Event FldValid(ByVal IdName As String, Cancel As Boolean, ErrMsg As String)
Public Property Get IdName() As String
    IdName = m_IdName
End Property
Public Property Let IdName(ByVal vNewValue As String)
    m_IdName = vNewValue
    TXT_REP.DataField = m_IdName
    PropertyChanged "optid"
End Property
Public Property Let text(ByVal vNewValue As Variant)
   TXT_REP.text = vNewValue
End Property

'Public Property Get DataField() As String
 '   DataField = TXT_REP.DataField
'End Property
Public Property Get text() As Variant
    text = TXT_REP.text
End Property
Public Property Get DataType() As String
    DataType = TXT_REP.DataType
End Property

Public Property Get CmpStr() As String
    CmpStr = TXT_REP.CmpStr
End Property
Public Property Let CmpStr(ByVal vNewValue As String)
    TXT_REP.CmpStr = vNewValue
    PropertyChanged "cmpstr"
End Property
Public Property Get ParentRef() As Object
    Set ParentRef = Parent
End Property
Public Property Get medref() As Object
    Set medref = TXT_REP
End Property

Private Sub TXT_REP_FldValid(ByVal IdName As String, Cancel As Boolean, ErrMsg As String)
    RaiseEvent FldValid(m_IdName, Cancel, ErrMsg)
End Sub

Private Sub TXT_REP_FldWhen(ByVal IdName As String, Cancel As Boolean, ErrMsg As String)
    RaiseEvent FldWhen(m_IdName, Cancel, ErrMsg)
End Sub
Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
    IdName = PropBag.ReadProperty("IdName", "")
    TXT_REP.CmpStr = PropBag.ReadProperty("CmpStr", "")
End Sub
Private Sub UserControl_Resize()
    TXT_REP.Top = 0
    TXT_REP.Left = 0
    TXT_REP.Height = UserControl.Height
    TXT_REP.Width = UserControl.Width
End Sub

Private Sub UserControl_WriteProperties(PropBag As PropertyBag)
    Call PropBag.WriteProperty("IdName", IdName, "")
    Call PropBag.WriteProperty("CmpStr", TXT_REP.CmpStr, "")
End Sub
Public Property Get Top() As Long
    Top = TXT_REP.Top
End Property
Public Property Get Left() As Long
    Left = TXT_REP.Left
End Property
Public Property Get Width() As Long
    Width = TXT_REP.Width
End Property
Public Property Get Height() As Long
    Height = TXT_REP.Height
End Property
Public Property Get Name() As String
    Name = UserControl.Extender.Name
End Property

Public Function EnAndVis()
    If Not (Enabled Or UserControl.Extender.Visible) Then
        EnAndVis = False
        Exit Function
    End If
    EnAndVis = True
    Set wobj = UserControl.Extender.Container
    Do While True
        If wobj.Name = Parent.Name Then
            Exit Do
        End If
        If (Not wobj.Enabled) Or (Not wobj.Visible) Then
            EnAndVis = False
            Exit Do
        Else
            Set wobj = wobj.Container
        End If
    Loop
End Function

Public Property Get IsBlank() As Boolean
    IsBlank = TXT_REP.IsBlank
End Property

