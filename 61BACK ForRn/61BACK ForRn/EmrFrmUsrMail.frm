VERSION 5.00
Object = "{6A2CE895-C1AC-48AF-B558-FB94BBE12499}#1.0#0"; "mwfctl.ocx"
Begin VB.Form EmrFrmUsrMail 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "User Administration"
   ClientHeight    =   9810
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   11925
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MDIChild        =   -1  'True
   ScaleHeight     =   9810
   ScaleMode       =   0  'User
   ScaleWidth      =   11925
   WindowState     =   2  'Maximized
   Begin MwfCtl.MWCTL_HLP MWCTL_HLP1 
      Height          =   405
      Left            =   8820
      TabIndex        =   5
      TabStop         =   0   'False
      Top             =   9300
      Visible         =   0   'False
      Width           =   555
      _ExtentX        =   979
      _ExtentY        =   714
   End
   Begin MwfCtl.MWCTL_ADC ADC 
      CausesValidation=   0   'False
      Height          =   555
      Left            =   90
      TabIndex        =   6
      Top             =   9270
      Width           =   11700
      _ExtentX        =   20638
      _ExtentY        =   979
      AllowAdd        =   0   'False
      AllowDelete     =   0   'False
      BtnWidth        =   885
      BtnHeight       =   435
      BtnMaskColor    =   8454143
      FontSz          =   8.25
   End
   Begin VB.Frame FRA_NKEYS 
      BorderStyle     =   0  'None
      Enabled         =   0   'False
      Height          =   9390
      Left            =   -120
      TabIndex        =   7
      Top             =   0
      Width           =   12255
      Begin VB.Frame FraNKeyAll 
         BorderStyle     =   0  'None
         Height          =   9195
         Left            =   150
         TabIndex        =   9
         Top             =   90
         Width           =   11850
         Begin VB.Frame FraDet 
            BorderStyle     =   0  'None
            Height          =   945
            Left            =   90
            TabIndex        =   10
            Top             =   8220
            Width           =   11715
            Begin MwfCtl.MWCTL_BTN1 CmdOk 
               Height          =   285
               Left            =   11100
               TabIndex        =   4
               Top             =   570
               Width           =   495
               _ExtentX        =   873
               _ExtentY        =   503
               BackColor       =   -2147483633
               MaskColor       =   8454143
               Caption         =   "&Ok"
               BeginProperty LblFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
            End
            Begin MwfCtl.MWCTL_MED ATXT 
               Height          =   285
               Index           =   1
               Left            =   1650
               TabIndex        =   3
               ToolTipText     =   "Enter Valid Machines for this User(No Spaces before or after Machine Name, Use comma("","") as a  seperator)"
               Top             =   570
               Width           =   9345
               _ExtentX        =   16484
               _ExtentY        =   503
               MaxLength       =   60
               DataType        =   4
               IdName          =   "WVALIDHOSTNM"
            End
            Begin MwfCtl.MWCTL_MED ATXT 
               Height          =   285
               Index           =   0
               Left            =   1650
               TabIndex        =   1
               ToolTipText     =   "Enter Valid Machines for this User"
               Top             =   0
               Width           =   9345
               _ExtentX        =   16484
               _ExtentY        =   503
               MaxLength       =   30
               IdName          =   "WLOGIN"
            End
            Begin MwfCtl.MWCTL_MED ATXT 
               Height          =   285
               Index           =   3
               Left            =   1650
               TabIndex        =   2
               ToolTipText     =   "Enter Valid Machines for this User"
               Top             =   285
               Width           =   9345
               _ExtentX        =   16484
               _ExtentY        =   503
               MaxLength       =   15
               IdName          =   "WPASS"
            End
            Begin VB.Label ALBL 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Password"
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   240
               Index           =   3
               Left            =   0
               TabIndex        =   13
               Top             =   285
               Width           =   915
            End
            Begin VB.Label ALBL 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Login Name"
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   240
               Index           =   0
               Left            =   0
               TabIndex        =   12
               Top             =   0
               Width           =   1155
            End
            Begin VB.Label ALBL 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Valid Machines"
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   240
               Index           =   2
               Left            =   0
               TabIndex        =   11
               Top             =   570
               Width           =   1455
            End
         End
         Begin MwfCtl.MWCTL_FLX GrdUsrMail 
            Height          =   8160
            Left            =   90
            TabIndex        =   0
            Top             =   0
            Width           =   11685
            _ExtentX        =   20611
            _ExtentY        =   14393
            Cols            =   9
            colname1        =   "UMUSRCD"
            heading1        =   "Usr Cd     "
            datafld1        =   "UmUsrCd"
            datatype1       =   4
            recalcparent1   =   "GrdUsrMail(WUSRNM)"
            maxlength1      =   5
            tooltiptext1    =   "Enter User Code"
            colname2        =   "WUSRNM"
            heading2        =   "User Name                                                              "
            recalcon2       =   "GrdUsrMail(UMUSRCD)"
            style2          =   2
            colname3        =   "UMMAILACT"
            heading3        =   "Active"
            datafld3        =   "UmMailAct"
            datatype3       =   4
            maxlength3      =   1
            tooltiptext3    =   "For This User Mail is Allowed or Not"
            colname4        =   "UMMAILADD"
            heading4        =   "Email Address                                                                                           "
            datafld4        =   "UmMailAdd"
            maxlength4      =   60
            tooltiptext4    =   "Email Address of User"
            colname5        =   "UMMAILLOGIN"
            heading5        =   "Login Name                              "
            datafld5        =   "UmMailLogin"
            maxlength5      =   30
            style5          =   0
            tooltiptext5    =   "Enter EMail Client Login Name"
            colname6        =   "UMMAILPASS"
            heading6        =   "Password                                "
            datafld6        =   "UmMailPass"
            maxlength6      =   15
            style6          =   0
            tooltiptext6    =   "Enter Email Client Pasword"
            colname7        =   "UMVALIDHOSTNM"
            heading7        =   "Valid Machins                                                          "
            datafld7        =   "UmValidHostNm"
            datatype7       =   4
            maxlength7      =   60
            style7          =   0
            tooltiptext7    =   "Enter Valid Host Names for This User"
            colname8        =   "WSETDET"
         End
      End
   End
   Begin VB.Frame FRA_KEYS 
      BorderStyle     =   0  'None
      Height          =   180
      Left            =   -36
      TabIndex        =   8
      Top             =   24
      Width           =   8190
   End
End
Attribute VB_Name = "EmrFrmUsrMail"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'--- General comments on the entry ----
'
'
'---
Option Explicit
Dim moCn As MwfLib.MDOConnection   '*** (Bef speed) bef 24/11/06

Private Sub ADC_Load()
  '*** Set the First Key Control property to adc
  '*** Set the First Non Key Control property to the grid GrdUsrMail
  '*** Set the Child Property Of ADC to GrdUsrMail
  '*** Set the Previous and Next Control Properties for the Grid GrdUsrMail
  '*** Set the Hot Key for the Grid GrdUsrMail as 'A'

  Set moCn = adc.Connection    '*** (Bef speed) bef 24/11/06

  Set adc.FirKeyCtl = adc
  Set adc.FirNKeyCtl = GrdUsrMail
  adc.Child = "GrdUsrMail"
  Set GrdUsrMail.PrevCtl = adc
  Set GrdUsrMail.NextCtl = adc
  GrdUsrMail.HotKey = "A"
End Sub
Private Sub ADC_SetRecSource()
  '*** Since this is an 'A' Entry, there is no RecSource associated with ADC. So set the
      'DataMode property of ADC control as xdatUnbound

  adc.DataMode = xdatUnbound
End Sub

Private Sub CmdOk_Click()
  With GrdUsrMail
    If .Row > 0 Then
      .Value(GrdUsrMail.Row, "UmMailLogin") = adc("WLOGIN")
      .Value(GrdUsrMail.Row, "UmMailPass") = adc("WPASS")
      .Value(GrdUsrMail.Row, "UmValidHostNm") = adc("WVALIDHOSTNM")
      Call SetDet(False)
      .Col = 1
      .SetFocus
    End If
  End With
End Sub

Private Sub ATXT_FldWhen(Index As Integer, ByVal IdName As String, Cancel As Boolean, ErrMsg As String)
  Call SetDet(True)
End Sub
Private Sub GrdUsrMail_When()
  Call SetDet(False)
End Sub
Private Sub GrdUsrMail_ColChange(ByVal RowNum As Integer, ByVal ColName As String, ByVal NewValue As Variant, Cancel As Boolean, ErrMsg As String)
  Select Case Trim$(UCase$(ColName))
    Case UCase$("UmUsrCd")
      Cancel = NewValue = "" Or Not moCn.RecSeek("Select 1 From Param Where PTyp ='USR' And PMCD = '" + NewValue + "'")
      If Cancel = True Then ErrMsg = "User Code Not Defined."
    Case UCase$("UmMailAct")
      Cancel = NewValue = "" Or Not moCn.RecSeek("Select 1 From Param Where PTyp ='YN' And PMCD = '" + NewValue + "'")
      If Cancel = True Then ErrMsg = "Mail Active(Y/N) Not Defined."
  End Select
End Sub
Private Sub GrdUsrMail_ColWhen(ByVal RowNum As Integer, ByVal ColName As String, Cancel As Boolean, ErrMsg As String, ByVal Action As MwfCtl.en_fgAction)
  Select Case Trim$(UCase$(ColName))
    Case UCase$("UmUsrCd")
      Call HlpList.PMCd("USR")
    Case UCase$("UmMailAct")
      Call HlpList.PMCd("YN")
    Case UCase$("wSetDet")
      adc("WLOGIN").SetFocus
  End Select
End Sub
Private Sub GrdUsrMail_ReCalc(ByVal RowNum As Integer, ByVal ColName As String, ByVal OldVal As Variant, ByVal SourceName As String, ByVal SourceColName As String)
  Select Case Trim$(UCase$(ColName))
    Case UCase$("wUsrNm")
      If RowNum > 0 And GrdUsrMail.Rows > RowNum Then _
        GrdUsrMail.Value(RowNum, ColName) = moCn.GetFldVal("Select PDesc From Param Where PTyp = 'USR' And PMCD = '" + GrdUsrMail.Value(RowNum, "UmUsrCd") + "'")
  End Select
End Sub
Private Sub GrdUsrMail_RowWhen(ByVal RowNum As Integer)
  adc("WLOGIN") = GrdUsrMail.Value(GrdUsrMail.Row, "UmMailLogin")
  adc("WPASS") = GrdUsrMail.Value(GrdUsrMail.Row, "UmMailPass")
  adc("WVALIDHOSTNM") = GrdUsrMail.Value(RowNum, "UmValidHostNm")
End Sub
Private Sub GrdUsrMail_SetRecSource()
  '*** Set the Record Source of the Grid GrdUsrMail
  GrdUsrMail.RecSource = "Select * from UsrMail Order By UmUsrCd"
End Sub


'''  ***** Std Code not to be changed **********************************
Private Sub Form_GotFocus()
  '*** Jenny Color (For Setting System Colors)
  SetMDISysColorsAfterLogin
  '*** Jenny Color (For Setting System Colors)
End Sub
Private Sub Form_Load()
  '*** Jenny Colour
  '(Bef 13/08/05)
  'SetMDISysColorsAfterLogin   '*** Jenny Color (For Setting System Colors)
  
  Dim ws_BtnStr As String
  ws_BtnStr = "CmdOk"
  Call SetObjColors(Me, ws_BtnStr)
  '*** Jenny Colour
End Sub
Private Sub Form_KeyUp(KeyCode As Integer, Shift As Integer)
  If adc.HotKeys(KeyCode, Shift) Then
    KeyCode = 0
    Shift = 0
  End If
End Sub
Private Sub Form_Activate()
  Call FrmActivate(Me)
End Sub
Private Sub Form_Deactivate()
  Call FrmDeActivate(Me)
End Sub
Private Sub ATXT_Validate(Index As Integer, Cancel As Boolean)
  Cancel = ATXT(Index).Validate
End Sub
Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    If UnloadMode <> vbFormCode Then Cancel = True: Exit Sub
    
    Set moCn = Nothing    '*** (09/08/05)    '*** (Bef speed) bef 24/11/06
End Sub
Private Sub GrdUsrMail_Validate(Cancel As Boolean)
  Cancel = GrdUsrMail.Validate
End Sub
Private Sub SetDet(ByVal pb_Set As Boolean)
  GrdUsrMail.Enabled = Not pb_Set
  adc.AllowSave = Not pb_Set
  CMDOK.Enabled = pb_Set
End Sub

