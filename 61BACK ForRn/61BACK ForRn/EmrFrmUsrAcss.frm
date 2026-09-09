VERSION 5.00
Object = "{6A2CE895-C1AC-48AF-B558-FB94BBE12499}#1.0#0"; "mwfctl.ocx"
Begin VB.Form EmrFrmUsrAcss 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "User Access Master"
   ClientHeight    =   9930
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   14580
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MDIChild        =   -1  'True
   ScaleHeight     =   9930
   ScaleWidth      =   14580
   WindowState     =   2  'Maximized
   Begin MwfCtl.MWCTL_BTN1 CmdPwd 
      Height          =   435
      Left            =   5100
      TabIndex        =   18
      ToolTipText     =   "Option To Copy The Menu Details From Another User"
      Top             =   9270
      Width           =   1035
      _ExtentX        =   1826
      _ExtentY        =   767
      MaskColor       =   8454143
      Caption         =   "&Password"
      Pic             =   "EmrFrmUsrAcss.frx":0000
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
   Begin MwfCtl.MWCTL_BTN1 CmdCP 
      Height          =   435
      Left            =   6135
      TabIndex        =   4
      ToolTipText     =   "Option To Copy The Menu Details From Another User"
      Top             =   9270
      Width           =   1035
      _ExtentX        =   1826
      _ExtentY        =   767
      BackColor       =   -2147483633
      MaskColor       =   8454143
      Caption         =   "&Copy"
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
   Begin MwfCtl.MWCTL_HLP MWCTL_HLP1 
      Height          =   405
      Left            =   9630
      TabIndex        =   6
      TabStop         =   0   'False
      Top             =   9420
      Visible         =   0   'False
      Width           =   555
      _ExtentX        =   979
      _ExtentY        =   714
   End
   Begin MwfCtl.MWCTL_ADC ADC 
      CausesValidation=   0   'False
      Height          =   555
      Left            =   60
      TabIndex        =   7
      Top             =   9270
      Width           =   12600
      _ExtentX        =   22225
      _ExtentY        =   979
      AllowAdd        =   0   'False
      AllowDelete     =   0   'False
      BtnWidth        =   885
      BtnHeight       =   435
      BtnMaskColor    =   8454143
      FontSz          =   8.25
   End
   Begin VB.Frame FRA_KEYS 
      BorderStyle     =   0  'None
      Height          =   420
      Left            =   -36
      TabIndex        =   9
      Top             =   24
      Width           =   9720
      Begin MwfCtl.MWCTL_MED ATXT 
         Height          =   285
         Index           =   0
         Left            =   1275
         TabIndex        =   0
         ToolTipText     =   "Enter User Code"
         Top             =   15
         Width           =   915
         _ExtentX        =   1614
         _ExtentY        =   503
         MaxLength       =   5
         DataType        =   4
         ReCalcParent    =   "WUAUSRDESC"
         IdName          =   "WUAUSRCD"
         CmpStr          =   "UaUsrCd= "
      End
      Begin MwfCtl.MWCTL_MED ATXT 
         Height          =   285
         Index           =   5
         Left            =   2235
         TabIndex        =   1
         ToolTipText     =   "User Description"
         Top             =   15
         Width           =   4680
         _ExtentX        =   8255
         _ExtentY        =   503
         Enabled         =   0   'False
         MaxLength       =   30
         ReCalcOn        =   "WUAUSRCD"
         IdName          =   "WUAUSRDESC"
      End
      Begin MwfCtl.MWCTL_MED ATXT 
         Height          =   285
         Index           =   3
         Left            =   8280
         TabIndex        =   2
         ToolTipText     =   "Enter User Code"
         Top             =   15
         Width           =   675
         _ExtentX        =   1191
         _ExtentY        =   503
         MaxLength       =   3
         DataType        =   4
         ReCalcParent    =   "WUAUSRDESC"
         IdName          =   "WMNUFORM"
      End
      Begin VB.Label ALBL 
         BackStyle       =   0  'Transparent
         Caption         =   "Menu Form"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Index           =   2
         Left            =   7080
         TabIndex        =   16
         Top             =   15
         Width           =   1305
      End
      Begin VB.Label ALBL 
         BackStyle       =   0  'Transparent
         Caption         =   "User Code"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Index           =   1
         Left            =   60
         TabIndex        =   11
         Top             =   15
         Width           =   1215
      End
   End
   Begin VB.Frame FRA_NKEYS 
      BorderStyle     =   0  'None
      Enabled         =   0   'False
      Height          =   9870
      Left            =   -120
      TabIndex        =   8
      Top             =   0
      Width           =   15315
      Begin VB.Frame FraCP 
         Height          =   975
         Left            =   2550
         TabIndex        =   12
         Top             =   8205
         Width           =   7800
         Begin MwfCtl.MWCTL_BTN1 CmdCpGo 
            Height          =   285
            Left            =   7020
            TabIndex        =   5
            ToolTipText     =   "Click Here To Copy From Another User"
            Top             =   420
            Width           =   555
            _ExtentX        =   0
            _ExtentY        =   0
            BackColor       =   -2147483633
            MaskColor       =   8454143
            Caption         =   "&Go"
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
            Left            =   1275
            TabIndex        =   13
            ToolTipText     =   "Enter Source User Code For Copying"
            Top             =   420
            Width           =   915
            _ExtentX        =   1614
            _ExtentY        =   503
            MaxLength       =   5
            DataType        =   4
            ReCalcParent    =   "WCPUSRDESC"
            IdName          =   "WCPUSR"
            CmpStr          =   "UaUsrCd= "
         End
         Begin MwfCtl.MWCTL_MED ATXT 
            Height          =   285
            Index           =   2
            Left            =   2235
            TabIndex        =   14
            ToolTipText     =   "User Description"
            Top             =   420
            Width           =   4680
            _ExtentX        =   8255
            _ExtentY        =   503
            Enabled         =   0   'False
            MaxLength       =   30
            ReCalcOn        =   "WCPUSR"
            IdName          =   "WCPUSRDESC"
         End
         Begin VB.Label LblMain 
            BackStyle       =   0  'Transparent
            BorderStyle     =   1  'Fixed Single
            Caption         =   "  Copy From"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   270
            Index           =   0
            Left            =   -30
            TabIndex        =   17
            Top             =   90
            Width           =   7800
         End
         Begin VB.Label ALBL 
            BackStyle       =   0  'Transparent
            Caption         =   "Copy From User"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   465
            Index           =   0
            Left            =   60
            TabIndex        =   15
            Top             =   420
            Width           =   1215
         End
      End
      Begin VB.Frame FraNKeyAll 
         BorderStyle     =   0  'None
         Height          =   9435
         Left            =   150
         TabIndex        =   10
         Top             =   90
         Width           =   15150
         Begin MwfCtl.MWCTL_FLX GrdUsrAcss 
            Height          =   8670
            Left            =   0
            TabIndex        =   3
            Top             =   420
            Width           =   14355
            _ExtentX        =   25321
            _ExtentY        =   15293
            Cols            =   9
            colname1        =   "UATAG"
            heading1        =   "Tag        "
            datafld1        =   "UaTag"
            datatype1       =   4
            recalcon1       =   "GrdUsrAcss(UAMNUCD)"
            recalcparent1   =   "GrdUsrAcss(UAMNUCD),GrdUsrAcss(UAMNUDESC),GrdUsrAcss(UAMNUFORM)"
            maxlength1      =   5
            tooltiptext1    =   "Menu Tag"
            colname2        =   "UAMNUCD"
            heading2        =   "Menu Code                                   "
            datafld2        =   "UaMnuCd"
            datatype2       =   4
            recalcon2       =   "GrdUsrAcss(UATAG)"
            recalcparent2   =   "GrdUsrAcss(UATAG),GrdUsrAcss(UAMNUDESC),GrdUsrAcss(UAMNUFORM)"
            maxlength2      =   20
            tooltiptext2    =   "Enter Menu Code"
            colname3        =   "UATLBARDESC"
            heading3        =   "ToolBar Caption"
            datafld3        =   "UaTlBarDesc"
            maxlength3      =   5
            tooltiptext3    =   "Enter Caption For Tool Bar Button"
            colname4        =   "UAMNUDESC"
            heading4        =   "Description                                                                  "
            datafld4        =   "UaMnuDesc"
            recalcon4       =   "GrdUsrAcss(UAMNUCD),GrdUsrAcss(UATAG)"
            maxlength4      =   30
            style4          =   2
            tooltiptext4    =   "Enter Menu Description"
            colname5        =   "UAADCOPT"
            heading5        =   "ADC (Y/N)"
            datafld5        =   "UaAdcOpt"
            datatype5       =   4
            maxlength5      =   1
            style5          =   0
            tooltiptext5    =   "Enter <Y>es/<N>o for Add/Delete/Change"
            colname6        =   "UAMNUFORM"
            heading6        =   "Menu Form"
            datafld6        =   "UaMnuForm"
            datatype6       =   4
            recalcon6       =   "GrdUsrAcss(UAMNUCD),GrdUsrAcss(UATAG)"
            maxlength6      =   3
            style6          =   0
            tooltiptext6    =   "Menu placed on which Form"
            colname7        =   "UAVALIDCOCD"
            heading7        =   "Valid Company Codes                 "
            datafld7        =   "UaValidCoCd"
            datatype7       =   4
            maxlength7      =   20
            tooltiptext7    =   "Enter Valid Company Codes"
            colname8        =   "UASEOYN"
            heading8        =   "SEO"
            datafld8        =   "UaSeoYn"
            datatype8       =   4
            maxlength8      =   1
            tooltiptext8    =   "Equivelant to SEO (Yes/No)"
         End
      End
      Begin VB.Frame FraPwd 
         Height          =   600
         Left            =   4140
         TabIndex        =   20
         Top             =   8580
         Width           =   4500
         Begin VB.TextBox TxtPwd 
            Height          =   285
            IMEMode         =   3  'DISABLE
            Left            =   2040
            MaxLength       =   30
            PasswordChar    =   "*"
            TabIndex        =   19
            ToolTipText     =   "Enter Seo Password"
            Top             =   105
            Width           =   2340
         End
         Begin VB.Label ALBL 
            BackStyle       =   0  'Transparent
            Caption         =   "SEO Pwd To Edit"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Index           =   51
            Left            =   60
            TabIndex        =   21
            Top             =   105
            Width           =   1935
         End
      End
   End
End
Attribute VB_Name = "EmrFrmUsrAcss"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'--- General comments on the entry ----
' FraCP, with CmdCP, CmdCpGo, wCpUsr
' X-'', A-UsrAccess
'
'---
Option Explicit
Dim moCn As MwfLib.MDOConnection   '*** (Bef speed) bef 24/11/06
Dim ms_wUaUsrCd As String
Dim mb_DelChild As Boolean
'*** Jay 2.14 *** (SEO Change)
Dim ms_PwdFlag As String
Enum en_UsrAcssFra
  Cpy = 0
  Pwd = 1
End Enum
'*** Jay 2.14 *** (SEO Change)

Private Sub ADC_Load()
  Set moCn = adc.Connection    '*** (Bef speed) bef 24/11/06

  Set adc.FirKeyCtl = adc("wUaUsrCd")
  Set adc.FirNKeyCtl = GrdUsrAcss
  adc.Child = "GrdUsrAcss"
  Set GrdUsrAcss.PrevCtl = adc
  Set GrdUsrAcss.NextCtl = adc
  GrdUsrAcss.HotKey = "A"
  ms_wUaUsrCd = ""
  '*** Jay 2.14 *** (SEO Change)
  ms_PwdFlag = moCn.RecSeek("Select PValue1 from Param where PTyp= 'SEOPWD' and PmCd= 'SEOPWD' and PValue1= '' ")
  '*** Jay 2.14 *** (SEO Change)
End Sub
Private Sub ADC_KeyWhen()
  Call CmpString("K")
  Call EnaDisaCmds(True)
  Call HideAllFras
  
  If ms_wUaUsrCd <> "" Then adc("wUaUsrCd") = ms_wUaUsrCd
End Sub
Private Sub ADC_KeyValid(Cancel As Boolean, ErrMsg As String)
  'Validating Usr Code, as it is used in Grd Init
  Call atxt_FldChange(0, UCase("wUaUsrCd"), Cancel, ErrMsg, adc("wUaUsrCd"))
  Call atxt_FldChange(2, UCase("wMnuForm"), Cancel, ErrMsg, adc("wMnuForm"))
  If Cancel = True Then Exit Sub
End Sub
Private Sub ADC_NonKeyWhen(pr_Cancel As Boolean, pr_ErrMsg As String)
  Call EnaDisaCmds(False)
  ms_wUaUsrCd = adc("wUaUsrCd")
  '*** Jay 2.14 *** (SEO Change)
  If FraPwd.Visible = True Then
    '*** Display or hide the Password frame depending on the toggle satate
    Call DispFra(Pwd)
  End If
  '*** Jay 2.14 *** (SEO Change)
End Sub
Private Sub ATXT_FldWhen(Index As Integer, ByVal IdName As String, Cancel As Boolean, ErrMsg As String)
  Select Case IdName
  Case Is = UCase("wUaUsrCd"), UCase("wCpUsr")
    Call HlpList.PMCd("USR", , , , False)
  Case Is = UCase("wMnuForm")
    Call HlpList.PMCd("MnuForm", "'" + ctProjectName + "'")
  End Select
End Sub
Private Sub atxt_ReCalculate(Index As Integer, ByVal IdName As String, ByVal OldVal As Variant, ByVal RowNum As Integer, ByVal SourceName As String, ByVal SourceColName As String)
  Select Case IdName
  Case Is = UCase("wUaUsrDesc")
    If adc.Mode = xNorm Then _
       adc("wUaUsrDesc") = moCn.GetFldVal("Select PDesc from Param where PTyp='USR' " + _
                                      "and PMCd='" + adc("wUaUsrCd") + "'")
  Case Is = UCase("wCpUsrDesc")
    If adc.Mode = xNorm Then _
       adc("wCpUsrDesc") = moCn.GetFldVal("Select PDesc from Param where PTyp='USR' " + _
                                      "and PMCd='" + adc("wCpUsr") + "'")
  End Select
End Sub
Private Sub atxt_FldChange(Index As Integer, ByVal IdName As String, Cancel As Boolean, ErrMsg As String, ByVal pv_NewValue As Variant)
  Select Case IdName
  Case Is = UCase("wUaUsrCd")
    Cancel = Not moCn.RecSeek("Select PMCd From Param Where PTyp = 'USR' " + _
                         "And PMCd = '" + pv_NewValue + "'")
    If Cancel = True Then ErrMsg = "Invalid User Code": Exit Sub
  Case Is = UCase("wMnuForm")
    Cancel = Not moCn.RecSeek("Select PMCd From Param Where PTyp = 'MnuForm' " + _
                         "And PMCd = '" + pv_NewValue + "' and PValue= '" + ctProjectName + "'")
    If Cancel = True Then ErrMsg = "Invalid Menu Form" ': Exit Sub
    Cancel = Not (UCase(pv_NewValue) = "MDI" Or UCase(pv_NewValue) = "LST")
    If Cancel = True Then ErrMsg = "Invalid Menu Form": Exit Sub
  Case Is = UCase("wCpUsr")
    Cancel = (pv_NewValue <> "") And (Not moCn.RecSeek("Select PMCd From Param Where PTyp = 'USR' " + _
                                      "And PMCd = '" + pv_NewValue + "'"))
    If Cancel = True Then ErrMsg = "Invalid Source User Code": Exit Sub
  End Select
End Sub
Private Sub GrdUsrAcss_ColWhen(ByVal RowNum As Integer, ByVal ColName As String, Cancel As Boolean, ErrMsg As String, ByVal Action As MwfCtl.en_fgAction)
  ' Manoj 2.10.0
  If Not moCn.RecSeek("Select 1 From Param Where PTyp = 'USR' And PmCd = '" & adc("WUAUSRCD") & "' And PValidYn = 'Y'") Then _
    Cancel = True: ErrMsg = "Invalid User Code, Cannot Edit Record.": Exit Sub
  ' Manoj 2.10.0

  Select Case ColName
  '****** Sachin 2.14.0 Next [03/02/2007] ******
  Case Is = UCase("UaTag")
    If GrdUsrAcss.IsNew(RowNum) Then
      Call HlpList.MenuTag(adc("WMNUFORM"))
    Else
      Cancel = True
      ErrMsg = "Cannot Edit Key Field."
    End If
   '****** Sachin 2.14.0 Next [03/02/2007] ******
  Case Is = UCase("UaMnuCd")
    If GrdUsrAcss.IsNew(RowNum) Then
      Call HlpList.MenuCode
    Else
      Cancel = True
      ErrMsg = "Cannot Edit Key Field."
    End If
  Case Is = UCase("UaAdcOpt")
    Call HlpList.PMCd("YN")
  ' Manoj 2.10.0
  Case Is = UCase("UaValidCoCd")
    HlpList.MultiSelect = True
    Call HlpList.hCoCd
  ' Manoj 2.10.0
  '*** Jay 2.14 *** (SEO Change)
  Case Is = UCase("UaSeoYn")
    If ms_PwdFlag = False Then Cancel = True: ErrMsg = "Enter SEO Password To Edit This Record": Exit Sub
    Call HlpList.PMCd("YN")
  '*** Jay 2.14 *** (SEO Change)
  
  End Select
End Sub

Private Sub GrdUsrAcss_OnDelete(ByVal RowNum As Integer, Cancel As Boolean, ErrMsg As String)
  Dim i As Integer, ws_Tag As String
  Dim wb_PrevDel As Boolean
  wb_PrevDel = mb_DelChild
  With GrdUsrAcss
    ws_Tag = .Value(RowNum, "UaTag")
    If ws_Tag <> "" Then
      If moCn.RecSeek("Select 'x' From UsrAccess Where UaUsrCd = '" + adc("wUaUsrCd") + "' " + _
                    "And UaTag Like '" + .Value(RowNum, "UaTag") + "%' " + _
                    "And UaMnuForm = '" + .Value(RowNum, "UaMnuForm") + "' " + _
                    "And UaMnuCd <> '" + .Value(RowNum, "UaMnuCd") + "'") Then
        If wb_PrevDel = False Then
          If MsgBox("Child Menus are Present for Current Menu. Do You Want to Delete All?", vbQuestion + vbYesNo, "Delete User Access") = vbYes Then _
            mb_DelChild = True
        End If
        If mb_DelChild Then
          i = 1
          Do While i <= .Rows - 1
            If .IsDel(i) = False And Left$(.Value(i, "UaTag"), Len(ws_Tag)) = ws_Tag And i <> RowNum Then
              .Del_Rec True, i
            Else
              i = i + 1
            End If
          Loop
        Else
          Cancel = True
          ErrMsg = "Deletion Cancelled."
        End If
      End If
    End If
  End With
  If wb_PrevDel = False Then mb_DelChild = False
End Sub

Private Sub GrdUsrAcss_ReCalc(ByVal RowNum As Integer, ByVal ColName As String, ByVal OldVal As Variant, ByVal SourceName As String, ByVal SourceColName As String)
  With GrdUsrAcss
    Select Case ColName
    Case Is = UCase("UaMnuCd")
      If (.Mode = fgmnorm) Then
        If (.Value(RowNum, "UaMnuCd") <> moCn.GetFldVal("Select UaMnuCd from UsrAccess where " + _
                                        "UaUsrCd= '" + ctGenMenuUsr + "' and UaTag= '" + .Value(RowNum, "UaTag") + "' And UaMnuForm='" + adc("wMnuForm") + "'")) Then _
          .Value(RowNum, "UaMnuCd") = moCn.GetFldVal("Select UaMnuCd from UsrAccess where " + _
                                        "UaUsrCd= '" + ctGenMenuUsr + "' and UaTag= '" + .Value(RowNum, "UaTag") + "' And UaMnuForm='" + adc("wMnuForm") + "'")
      End If
    Case Is = UCase("UaMnuDesc")
      If .Mode = fgmnorm Then
        .Value(RowNum, "UaMnuDesc") = moCn.GetFldVal("Select UaMnuDesc from UsrAccess where " + _
                                      "UaUsrCd= '" + ctGenMenuUsr + "' and " + IIF(UCase(SourceColName) = UCase("UaMnuCd"), "UaMnuCd= '" + .Value(RowNum, "UaMnuCd") + "'", "UaTag= '" + .Value(RowNum, "UaTag") + "' And UaMnuForm='" + adc("wMnuForm") + "'"))
      End If
    Case Is = UCase("UaMnuForm")
      If .Mode = fgmnorm Then
        .Value(RowNum, "UaMnuForm") = moCn.GetFldVal("Select UaMnuForm from UsrAccess where " + _
                                      "UaUsrCd= '" + ctGenMenuUsr + "' and UaMnuCd= '" + .Value(RowNum, "UaMnuCd") + "'")
      End If
    Case Is = UCase("UaTag")
      If (.Mode = fgmnorm) Then
        If (.Value(RowNum, "UaTag") <> moCn.GetFldVal("Select UaTag from UsrAccess where " + _
                                        "UaUsrCd= '" + ctGenMenuUsr + "' and UaMnuCd= '" + .Value(RowNum, "UaMnuCd") + "'")) Then _
          .Value(RowNum, "UaTag") = moCn.GetFldVal("Select UaTag from UsrAccess where " + _
                                        "UaUsrCd= '" + ctGenMenuUsr + "' and UaMnuCd= '" + .Value(RowNum, "UaMnuCd") + "'")
      End If
    End Select
  End With
End Sub

Private Sub GrdUsrAcss_ColChange(ByVal RowNum As Integer, ByVal ColName As String, ByVal NewValue As Variant, Cancel As Boolean, ErrMsg As String)
  Dim ws_CoCd() As String, i As Integer
  With GrdUsrAcss
    Select Case ColName
    Case Is = UCase("UaMnuCd")
      Cancel = (NewValue <> "") And Not moCn.RecSeek("Select UaMnuCd From UsrAccess where UaUsrCd= '" + ctGenMenuUsr + "' " + _
                           "and UaMnuCd = '" + Replace$(Replace$(NewValue, "'", "''"), """", """""") + "'")
      ErrMsg = "Invalid Menu Code": Exit Sub
    Case Is = UCase("UaAdcOpt")
      Cancel = (NewValue <> "") And (Not moCn.RecSeek("Select PMCd From Param where " + _
                                     "PTyp= 'YN' and PMCd = '" + NewValue + "'"))
      ErrMsg = "Enter 'Y' or 'N'": Exit Sub
    ' Manoj 2.10.0
    Case Is = UCase$("UaValidCoCd")
        ws_CoCd = Split(NewValue, ",")
        If IsArray(ws_CoCd) Then
          For i = 0 To UBound(ws_CoCd)
            Cancel = Not moCn.RecSeek("Select 1 From Head where " + _
                                      "HCd = 'ZSELF' and HCoCd = '" + Replace$(Replace$(ws_CoCd(i), "'", "''"), """", """""") + "'")
            If Cancel = True Then ErrMsg = "Value(" + CStr(i + 1) + ") Is An Invalid Company Code": Exit Sub
          Next i
        End If
    ' Manoj 2.10.0
    '*** Jay 2.14 *** (SEO Change)
    Case Is = UCase("UaSeoYn")
      If NewValue <> "" And NewValue <> "Y" Then
        Cancel = (Not moCn.RecSeek("Select PMCd From Param where PTyp= 'YN' and PMCd = '" + NewValue + "'"))
        If Cancel = True Then ErrMsg = "Enter 'Y' To Assign Equivalent SEO Rights and 'N'/'' Not To Assign": Exit Sub
      ElseIf NewValue = "Y" Then
         Cancel = moCn.RecSeek("Select 'x' from param where PTyp = 'USR' and PmCd = '" + adc("wUaUsrCd") + "' and Pvalue = ''")
         If Cancel = True Then ErrMsg = "User Passward Is Blank, Can Not Assign Seo Rights": Exit Sub
      End If
    '*** Jay 2.14 *** (SEO Change)
    
      
    End Select
  End With
End Sub
Private Sub GrdUsrAcss_InitKey(ByVal RowNum As Integer)
  With GrdUsrAcss
    .Store "UaUsrCd", adc("wUaUsrCd")
    
    '*** (Jen 2.14 Next Patch 1)
    If .IsNew(RowNum) Then
      .Store "UaReplNo", 0
    End If
    '*** (Jen 2.14 Next Patch 1)
  End With
End Sub
Private Sub ADC_SetRecSource()
  adc.DataMode = xdatUnbound
End Sub
Private Sub GrdUsrAcss_SetRecSource()
  Dim ws_Cnd As String, ws_OrdBy As String
  ws_Cnd = adc.RepCond
  ws_Cnd = IIF(ws_Cnd <> "", " where " + ws_Cnd, "")
  ws_Cnd = IIF(ws_Cnd = "", " where ", ws_Cnd + " And ") + " UaMnuForm = '" + adc("wMnuForm") + "'"
  ws_OrdBy = " Order By UaTag, UaMnuCd "
  GrdUsrAcss.RecSource = " Select * from UsrAccess " + ws_Cnd + ws_OrdBy
End Sub
Private Sub CmdCP_Click()
  Call DispFra(Cpy)
End Sub
Private Sub CmdCpGo_Click()
  If FraCP.Visible = True Then Call CpyMenuDet
  Call CmdCP_Click
End Sub
'*** Jay 2.14 *** (SEO Change)
Private Sub CmdPwd_Click()
  Call DispFra(Pwd)
End Sub
'*** Jay 2.14 *** (SEO Change)

Private Sub CpyMenuDet()
  Dim wCnd As String, wRow As Single, wOrdBy As String
  Call CmpString("C")
  
  If Not moCn.RecSeek("Select 1 From Param Where PTyp = 'USR' And PmCd = '" & adc("WUAUSRCD") & "' And PValidYn = 'Y'") Then _
    DispMsg "Invalid User Code, Cannot Copy Records.", etError: Exit Sub

  If GrdUsrAcss.Rows - 1 > 0 Then DispMsg "Record Already Exists", etError: Exit Sub
  If adc("wCpUsr") = "" Then DispMsg "Enter Source User Code For Copying", etError: Exit Sub

  wCnd = adc.RepCond
  wCnd = IIF(wCnd <> "", " where " + wCnd, "")
  wCnd = IIF(wCnd = "", " where ", wCnd + " And ") + " UaMnuForm = '" + adc("wMnuForm") + "'"
  wOrdBy = " Order By UaTag, UaMnuCd "
  
  '*** Beginning the copy process
  GrdUsrAcss.StartCopy
  MWLib.BeginProcess Me, "Copying Menu Options ..."
  Dim wRsUsrAcss As MwfLib.MDORowSet
  '*** Jay 2.14 *** (SEO Change)
  If ms_PwdFlag = True And moCn.RecSeek("Select 'x' From Param Where PTyp = 'USR' And PmCd = '" & adc("WUAUSRCD") & "' And PValue <> ''") Then
    Set wRsUsrAcss = moCn.OpenRes("Select UaSEOYn as qSeoYn, * from UsrAccess " + wCnd + wOrdBy)
  Else
    Set wRsUsrAcss = moCn.OpenRes("Select '' as qSeoYn, * from UsrAccess " + wCnd + wOrdBy)
  End If
  'Set wRsUsrAcss = moCn.OpenRes("Select * from UsrAccess where UaUsrCd= '" + ADC("wCpUsr") + "'")
  'Set wRsUsrAcss = moCn.OpenRes("Select * from UsrAccess " + wCnd + wOrdBy)
  '*** Jay 2.14 *** (SEO Change)
  With wRsUsrAcss
    Do While Not (.EOF Or .BOF)
      MWLib.RunningProcess "Copying Menu Code: " + !UaMnuCd
      GrdUsrAcss.AddItem
      wRow = GrdUsrAcss.Rows - 1
      GrdUsrAcss.Value(wRow, "UaMnuCd") = !UaMnuCd
      GrdUsrAcss.Value(wRow, "UaMnuDesc") = !UaMnuDesc
      GrdUsrAcss.Value(wRow, "UaAdcOpt") = !UaAdcOpt
      GrdUsrAcss.Value(wRow, "UaMnuForm") = !UaMnuForm
      GrdUsrAcss.Value(wRow, "UaTlBarDesc") = !UaTlBarDesc
      'Manoj 2.10.0
      GrdUsrAcss.Value(wRow, "UaValidCoCd") = !UaValidCoCd
      GrdUsrAcss.Value(wRow, "UaTag") = !UaTag
      'Manoj 2.10.0
      '*** Jay 2.14 *** (SEO Change)
      GrdUsrAcss.Value(wRow, "UaSeoYn") = !qSeoYn
      '*** Jay 2.14 *** (SEO Change)
      GrdUsrAcss.SaveRec (wRow)
      .MoveNext
    Loop
  End With
  MWLib.EndProcess Me
  GrdUsrAcss.EndCopy
  '*** End of the copy process
  Set wRsUsrAcss = Nothing
End Sub
Private Sub CmpString(ByVal mTyp As String)
  If mTyp = "K" Then
    adc("wUaUsrCd").CmpStr = "UaUsrCd= "
    adc("wCpUsr").CmpStr = ""
  ElseIf mTyp = "C" Then
    adc("wUaUsrCd").CmpStr = ""
    adc("wCpUsr").CmpStr = "UaUsrCd= "
  End If
End Sub
Private Sub DispFra(ByVal pv_UsrAcssFra As en_UsrAcssFra)
  Select Case pv_UsrAcssFra
  Case Is = Cpy
    If FraCP.Visible = True Then
      FraCP.Visible = False
      FraCP.Enabled = False
      Call EnaDisaCmds(False, CmdCP)
      CmdCP.SetFocus
    Else
      FraCP.Visible = True
      FraCP.Enabled = True
      adc("wCpUsr").SetFocus
      FraCP.ZOrder
      Call EnaDisaCmds(True, CmdCP)
    End If
  '*** Jay 2.14 *** (SEO Change)
  Case Is = Pwd
    If FraPwd.Visible = True Then
      FraPwd.Visible = False
      FraPwd.Enabled = False
      Call EnaDisaCmds(False, CmdPwd)
      CmdPwd.SetFocus
    Else
      FraPwd.Visible = True
      FraPwd.Enabled = True
      FraPwd.ZOrder
      Call EnaDisaCmds(True, CmdPwd)
    End If
  '*** Jay 2.14 *** (SEO Change)
  End Select
End Sub
Private Sub EnaDisaCmds(ByVal pv_ShowFra As Boolean, Optional ByVal pv_CmdBut As MWCTL_BTN1)
  If Not pv_CmdBut Is Nothing Then
    If pv_ShowFra = True Then
      '*** Jenny (Old Code Bef Color)
'      pv_CmdBut.BackColor = vbWhite
'      pv_CmdBut.FontBold = True
      '*** Jenny (Old Code Bef Color)
      If pv_CmdBut.Pic = 0 Then
        pv_CmdBut.BackColor = vbWhite
        'pv_CmdBut.FontBold = True
      End If
      pv_CmdBut.CausesValidation = True
    Else
      '*** Jenny (Old Code Bef Color)
'      pv_CmdBut.BackColor = &H8000000F
'      pv_CmdBut.FontBold = False
      '*** Jenny (Old Code Bef Color)
      If pv_CmdBut.Pic = 0 Then
        pv_CmdBut.BackColor = &H8000000F
        'pv_CmdBut.FontBold = False
      End If
      'pv_CmdBut.CausesValidation = False
    End If
  End If
  FraNKeyAll.Enabled = Not pv_ShowFra
  adc.AllowSave = Not pv_ShowFra
  CmdCP.Enabled = Not pv_ShowFra
  '*** Jay 2.14 *** (SEO Change)
  CmdPwd.Enabled = Not pv_ShowFra
  '*** Jay 2.14 *** (SEO Change)
  If Not pv_CmdBut Is Nothing Then pv_CmdBut.Enabled = True
End Sub
Private Sub HideAllFras()
  FraNKeyAll.Visible = True
  FraCP.Visible = False
  FraCP.Enabled = False
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
  ws_BtnStr = "CmdCpGo"
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
Private Sub GrdUsrAcss_Validate(Cancel As Boolean)
  Cancel = GrdUsrAcss.Validate
End Sub
'*** Jay 2.14 *** (SEO Change)
Private Sub TxtPwd_GotFocus()
  DispMsg "Enter Seo Password", etInfo
End Sub
'*** Jay 2.14 *** (SEO Change)
Private Sub TxtPwd_Validate(Cancel As Boolean)
  If adc.AddFndMode = xFndMode Then
    If Trim(TxtPwd.text) <> "" Then
      If Trim(TxtPwd.text) = moCn.GetFldVal("Select PValue1 from Param where Ptyp= 'SEOPWD' and PmCd= 'SEOPWD' ") Then
          ms_PwdFlag = True
      Else
          ms_PwdFlag = False: Cancel = True: DispMsg "Invalid SEO Password ", etError
      End If
    Else
      ms_PwdFlag = moCn.RecSeek("Select PValue1 from Param where PTyp= 'SEOPWD' and PmCd= 'SEOPWD' and PValue1= '' ")
    End If
  End If
End Sub

