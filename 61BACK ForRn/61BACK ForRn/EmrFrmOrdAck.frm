VERSION 5.00
Object = "{AB2CC32A-E652-42A2-B341-26ECE02F6386}#2.1#0"; "mwfctlforspeed9.ocx"
Begin VB.Form EmrFrmOrdAck 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Order Acknowledgment"
   ClientHeight    =   9810
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   15210
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MDIChild        =   -1  'True
   ScaleHeight     =   9810
   ScaleWidth      =   15210
   WindowState     =   2  'Maximized
   Begin MwfCtl.MWCTL_BTN1 CmdAck 
      CausesValidation=   0   'False
      Height          =   435
      Left            =   6555
      TabIndex        =   2
      ToolTipText     =   "Option To Acknowledge Order"
      Top             =   9195
      Width           =   1455
      _ExtentX        =   2566
      _ExtentY        =   767
      MaskColor       =   8454143
      Caption         =   "Ac&knowledge"
      Pic             =   "EmrFrmOrdAck.frx":0000
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
      Left            =   8370
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   9300
      Visible         =   0   'False
      Width           =   435
      _ExtentX        =   767
      _ExtentY        =   714
   End
   Begin MwfCtl.MWCTL_ADC ADC 
      CausesValidation=   0   'False
      Height          =   555
      Left            =   30
      TabIndex        =   1
      Top             =   9195
      Width           =   14505
      _ExtentX        =   25585
      _ExtentY        =   979
      AllowAdd        =   0   'False
      AllowDelete     =   0   'False
      BtnPicture      =   "EmrFrmOrdAck.frx":03B2
      BtnWidth        =   885
      BtnHeight       =   435
      BtnMaskColor    =   8454143
      FontSz          =   8.25
   End
   Begin VB.Frame FRA_NKEYS 
      BorderStyle     =   0  'None
      Enabled         =   0   'False
      Height          =   9405
      Left            =   -180
      TabIndex        =   6
      Top             =   -60
      Width           =   15345
      Begin VB.Frame FraAck 
         Height          =   1965
         Left            =   4950
         TabIndex        =   9
         Top             =   7200
         Width           =   5160
         Begin MwfCtl.MWCTL_BTN1 CmdAckGo 
            Height          =   495
            Left            =   4440
            TabIndex        =   5
            Top             =   1305
            Width           =   555
            _ExtentX        =   979
            _ExtentY        =   873
            MaskColor       =   8454143
            Caption         =   "&Go"
            Pic             =   "EmrFrmOrdAck.frx":0764
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
            Index           =   72
            Left            =   1620
            TabIndex        =   3
            ToolTipText     =   "Enter Voucher Character"
            Top             =   435
            Width           =   780
            _ExtentX        =   1376
            _ExtentY        =   503
            MaxLength       =   3
            DataType        =   4
            IdName          =   "WOMCHR"
         End
         Begin MwfCtl.MWCTL_MED ATXT 
            Height          =   285
            Index           =   42
            Left            =   2520
            TabIndex        =   4
            ToolTipText     =   "Enter Voucher No."
            Top             =   435
            Width           =   1080
            _ExtentX        =   1905
            _ExtentY        =   503
            Alignment       =   2
            Mask            =   "#####0"
            MaxLength       =   6
            DataType        =   1
            IdName          =   "WOMNO"
         End
         Begin VB.Label LblNewVchNo 
            BackStyle       =   0  'Transparent
            Caption         =   "New Order No Will Be KJ/SO/06/SPL/123456. Do You Wish To Continue?"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   525
            Left            =   90
            TabIndex        =   14
            Top             =   780
            Width           =   4905
         End
         Begin VB.Label ALBL 
            BackStyle       =   0  'Transparent
            Caption         =   "/"
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
            Index           =   40
            Left            =   2430
            TabIndex        =   15
            Top             =   450
            Width           =   105
         End
         Begin VB.Label ALBL 
            BackStyle       =   0  'Transparent
            Caption         =   "Vch Chr/No"
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
            Index           =   39
            Left            =   90
            TabIndex        =   13
            Top             =   435
            Width           =   1185
         End
         Begin VB.Label LblMain 
            BackStyle       =   0  'Transparent
            BorderStyle     =   1  'Fixed Single
            Caption         =   "  Order Acknowledgement"
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
            Index           =   1
            Left            =   -30
            TabIndex        =   10
            Top             =   90
            Width           =   5205
         End
      End
      Begin VB.Frame FraNKeyAll 
         BorderStyle     =   0  'None
         Height          =   9345
         Left            =   180
         TabIndex        =   8
         Top             =   30
         Width           =   14640
         Begin MwfCtl.MWCTL_FLX GrdOrdMst 
            Height          =   5550
            Left            =   30
            TabIndex        =   11
            Top             =   60
            Width           =   14340
            _ExtentX        =   25294
            _ExtentY        =   9790
            Cols            =   14
            colname1        =   "WSR"
            heading1        =   "Sr.       "
            datatype1       =   1
            mask1           =   "####0"
            maxlength1      =   5
            style1          =   2
            tooltiptext1    =   "Sr. No."
            colname2        =   "WOMCRMNO"
            heading2        =   "Web Order No       "
            datatype2       =   4
            maxlength2      =   13
            style2          =   2
            tooltiptext2    =   "CRM No"
            colname3        =   "WOMDT"
            heading3        =   "Date         "
            datatype3       =   3
            mask3           =   "##/##/##"
            maxlength3      =   8
            style3          =   2
            tooltiptext3    =   "Order Date"
            colname4        =   "WOMCMCD"
            heading4        =   "Customer     "
            datatype4       =   4
            maxlength4      =   8
            style4          =   2
            tooltiptext4    =   "Customer Code"
            colname5        =   "WOMPONO"
            heading5        =   "PO No                                           "
            datatype5       =   4
            maxlength5      =   20
            style5          =   2
            tooltiptext5    =   "Order PO No."
            colname6        =   "WOMACKYN"
            heading6        =   "Ack"
            datatype6       =   4
            maxlength6      =   1
            style6          =   2
            tooltiptext6    =   "Enter Order Acknowledged (Y/N)"
            colname7        =   "WOMCOCD"
            heading7        =   "CoCd"
            datatype7       =   4
            maxlength7      =   3
            style7          =   0
            colname8        =   "WOMTC"
            heading8        =   "Tc"
            datatype8       =   4
            maxlength8      =   3
            style8          =   0
            colname9        =   "WOMYY"
            heading9        =   "Yy"
            datatype9       =   4
            maxlength9      =   2
            style9          =   0
            colname10       =   "WOMCHR"
            heading10       =   "Chr"
            datatype10      =   4
            maxlength10     =   3
            style10         =   0
            colname11       =   "WOMNO"
            heading11       =   "No"
            datatype11      =   1
            mask11          =   "#####0"
            maxlength11     =   6
            style11         =   0
            colname12       =   "WOMNUMBER"
            heading12       =   "Current Order No                            "
            datatype12      =   4
            maxlength12     =   21
            style12         =   2
            tooltiptext12   =   "Current Order No"
            colname13       =   "WOMNONEW"
            heading13       =   "New Order No                               "
            datatype13      =   4
            maxlength13     =   21
            style13         =   2
            tooltiptext13   =   "New Order No"
         End
         Begin MwfCtl.MWCTL_FLX GrdOrdDsg 
            Height          =   3510
            Left            =   30
            TabIndex        =   12
            Top             =   5610
            Width           =   11610
            _ExtentX        =   20479
            _ExtentY        =   6191
            Cols            =   11
            colname1        =   "WODSR"
            heading1        =   "Sr.       "
            datatype1       =   1
            mask1           =   "####0"
            maxlength1      =   5
            style1          =   2
            tooltiptext1    =   "Order Design Serial"
            colname2        =   "WODDMCD"
            heading2        =   "Design Cd                       "
            datatype2       =   4
            maxlength2      =   15
            style2          =   2
            tooltiptext2    =   "Design Code"
            colname3        =   "WODSFX"
            heading3        =   "Suffix                 "
            datatype3       =   4
            maxlength3      =   15
            style3          =   2
            tooltiptext3    =   "Design Suffix"
            colname4        =   "WODDMSZ"
            heading4        =   "Size      "
            datatype4       =   4
            maxlength4      =   5
            style4          =   2
            tooltiptext4    =   "Design Size"
            colname5        =   "WODORDENT"
            heading5        =   "Qty           "
            datatype5       =   2
            mask5           =   "####0.0"
            maxlength5      =   7
            style5          =   2
            tooltiptext5    =   "Design Quantity"
            colname6        =   "WODCALCPRC"
            heading6        =   "Cost Prc      "
            datatype6       =   2
            mask6           =   "#####0.00"
            maxlength6      =   9
            style6          =   2
            tooltiptext6    =   "Calculated Price"
            colname7        =   "WODSALPRC"
            heading7        =   "Sales Prc    "
            datatype7       =   2
            mask7           =   "#####0.00"
            maxlength7      =   9
            style7          =   2
            tooltiptext7    =   "Sales Price"
            colname8        =   "WODEXPDELDT"
            heading8        =   "Exp Del Dt"
            datatype8       =   3
            maxlength8      =   8
            style8          =   2
            tooltiptext8    =   "Export Delivery Date"
            colname9        =   "WODKT"
            heading9        =   "Karat     "
            datatype9       =   4
            maxlength9      =   5
            style9          =   2
            tooltiptext9    =   "Design Karatage"
            colname10       =   "WODDMCOL"
            heading10       =   "Color         "
            datatype10      =   4
            maxlength10     =   5
            style10         =   2
            tooltiptext10   =   "Design Colour"
         End
      End
   End
   Begin VB.Frame FRA_KEYS 
      BorderStyle     =   0  'None
      Height          =   270
      Left            =   -36
      TabIndex        =   7
      Top             =   30
      Width           =   6780
   End
End
Attribute VB_Name = "EmrFrmOrdAck"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim moCn As MwfLib.MDOConnection   '*** (Bef speed) bef 24/11/06
Enum en_OrdAckFra
  OrdAck = 0
End Enum
'*** Arr_Cv stores the causes validation state, Arr_Ed the Ena/Disa state of Cmd Buttons
Dim mArr_Cv As Variant, mArr_Ed As Variant
Dim ws_NewVchNo As String

Private Sub ADC_Load()
  Set moCn = ADC.Connection  '*** (Bef speed) bef 24/11/06
  Set ADC.FirNKeyCtl = GrdOrdMst
  Set GrdOrdMst.PrevCtl = ADC
  Set GrdOrdMst.NextCtl = GrdOrdDsg
  Set GrdOrdDsg.PrevCtl = GrdOrdMst
  Set GrdOrdDsg.NextCtl = ADC
  GrdOrdMst.HotKey = "A"
  GrdOrdDsg.HotKey = "B"

  ReDim mArr_Cv(0), mArr_Ed(0)
  GrdOrdMst.AllowAdd = False: GrdOrdMst.AllowDelete = False
  GrdOrdDsg.AllowAdd = False: GrdOrdDsg.AllowDelete = False
End Sub
Private Sub ADC_KeyWhen()
  Call EnaDisaCmds(True)
  Call HideAllFras
  GrdOrdMst.Rows = 1
  GrdOrdDsg.Rows = 1
  
  '****** Sachin 2.14.0 - [24-11-2006] - [Multi-Company] ******
  Me.Caption = GetFrmCaption(gs_CoCd, "Order Acknowledgment")
  '****** Sachin 2.14.0 - [24-11-2006] - [Multi-Company] ******
  
End Sub
Private Sub ADC_NonKeyWhen(Cancel As Boolean, ErrMsg As String)
  Call EnaDisaCmds(False)
  CmdAck.Enabled = False
  Call FillOrdMst
End Sub
Private Sub ATXT_FldWhen(Index As Integer, ByVal IdName As String, Cancel As Boolean, ErrMsg As String)
  '*** Set Helps For Each Field
  Select Case UCase(IdName)
  Case Is = UCase("wOmChr")
    Call HlpList.vPSCd(gs_CoCd, "CHR", GrdOrdMst.Value(GrdOrdMst.Row, "wOmTc"), "Y")
  End Select
End Sub
Private Sub ATXT_FldChange(Index As Integer, ByVal IdName As String, Cancel As Boolean, ErrMsg As String, ByVal pv_NewValue As Variant)
  '***
  Select Case UCase(IdName)
  Case Is = UCase("wOmChr")
    Cancel = (pv_NewValue = "WEB") Or (pv_NewValue <> "" And (Not moCn.RecSeek("select vPSCd from vParam where " + _
                         "vPCoCd= '" + gs_CoCd + "' and vPTyp= 'CHR' and " + _
                         "vPMCd = '" + GrdOrdMst.Value(GrdOrdMst.Row, "wOmTc") + "' and vPSCd='" + pv_NewValue + "'")))
    'Cancel = pv_NewValue <> "" And (Not moCn.RecSeek("select vPSCd from vParam where " + _
                         "vPCoCd= '" + gs_CoCd + "' and vPTyp= 'CHR' and " + _
                         "vPMCd = '" + GrdOrdMst.Value(GrdOrdMst.Row, "wOmTc") + "' and vPSCd='" + pv_NewValue + "'"))
    If Cancel = True Then
      ErrMsg = "Invalid Voucher Character": Exit Sub
    Else
      '************Geeta********Emr207***
      Dim ws_UsrArr() As String, i As Integer
      ws_UsrArr = Split(moCn.GetFldVal("Select vPValue4 from vParam where vPCoCd= '" + gs_CoCd + "' and " + _
                                             "vPTyp='CHR' and vPMCd='" + GrdOrdMst.Value(GrdOrdMst.Row, "wOmTc") + "' and " + _
                                             "vPSCd= '" + pv_NewValue + "'"), ",")
      If IsArray(ws_UsrArr) Then
        For i = 0 To UBound(ws_UsrArr)
          Cancel = IIF(UCase(gs_UsrCd) = UCase(ws_UsrArr(i)), False, True)
          If Cancel = False Then: Exit For
        Next i
      End If
      If Cancel = True Then ErrMsg = "User Has No Right To Access This Voucher Character": Exit Sub
    End If
  End Select
  
  LblNewVchNo.Caption = ""
 
End Sub
Private Sub ADC_SetRecSource()
  '*** Set the RecSource for ADC
  ADC.DataMode = xdatUnbound
End Sub
Private Sub CmdAck_Click()
  '*** Display or hide the Copy frame depending on the toggle state
  If GrdOrdMst.Rows <= 1 Then Exit Sub
  
  If UCase(GrdOrdMst.Value(GrdOrdMst.Row, "wOmCoCd")) <> "WEB" Then
    ADC("wOmChr") = "": ADC("wOmChr").Enabled = False
    ADC("wOmNo") = "0": ADC("wOmNo").Enabled = False
  Else
    If ADC("wOmChr") = "" Then _
      ADC("wOmChr") = moCn.GetFldVal("Select vPValue from vParam where vPCoCd= '" + gs_CoCd + "' and vPTyp='DEFCHR' " + _
                                  "and vPMCd='" + GrdOrdMst.Value(GrdOrdMst.Row, "wOmTc") + "'")
      ADC("wOmChr").Enabled = True: ADC("wOmNo").Enabled = True
  End If
  
  LblNewVchNo.Caption = ""
  Call DispFra(OrdAck)
End Sub
Private Sub DispFra(ByVal pv_DsgFra As en_OrdAckFra)
  '*** This is the common procedure which is used to toggle between hiding and displaying
      'the frames (i.e. if the frame is visible then make it invisible and viseversa)
      'associated with a command button
  Select Case pv_DsgFra
  Case Is = OrdAck
    If FraAck.Visible = True Then
      FraAck.Visible = False
      FraAck.Enabled = False
      Call EnaDisaCmds(False, CmdAck, OrdAck)
      'CmdAck.SetFocus
      GrdOrdMst.SetFocus
    Else
      FraAck.Visible = True
      FraAck.Enabled = True
      If ADC("wOmChr").Enabled = True Then ADC("wOmChr").SetFocus
      FraAck.ZOrder
      Call EnaDisaCmds(True, CmdAck)
    End If
  End Select
End Sub
Private Sub EnaDisaCmds(ByVal pv_ShowFra As Boolean, Optional ByVal pv_CmdBut As MWCTL_BTN1, Optional ByVal pv_DispFra As en_OrdAckFra)
  '*** This procedure is used to give the toggle option (enable and disable) to
      'the command buttons
  
  '*** storing the state of all cmd buttons, when the frame is going to be visible.
  If pv_ShowFra = True Then
    Call StoreState
  End If
  
  If Not pv_CmdBut Is Nothing Then
    If pv_ShowFra = True Then
      '*** Old Code
      'pv_CmdBut.BackColor = vbWhite
      'pv_CmdBut.FontBold = True
      '*** Old Code
      pv_CmdBut.CausesValidation = True
    Else
      '*** Old Code
      'pv_CmdBut.BackColor = &H8000000F
      'pv_CmdBut.FontBold = False
      '*** Old Code
      
      'pv_CmdBut.CausesValidation = False
      If Not IsMissing(pv_DispFra) Then
        pv_CmdBut.CausesValidation = mArr_Cv(pv_DispFra)
      End If
    End If
  End If
  FraNKeyAll.Enabled = Not pv_ShowFra
  ADC.AllowSave = Not pv_ShowFra
  If pv_ShowFra = True Then
    '** All command buttons are disabled
    CmdAck.Enabled = Not pv_ShowFra
  Else
    CmdAck.Enabled = mArr_Ed(OrdAck)
  End If

  If Not pv_CmdBut Is Nothing Then pv_CmdBut.Enabled = True
End Sub
Private Sub HideAllFras()
  '*** This procedure hides all the frames other than FraNKeyAll
  FraNKeyAll.Visible = True
  FraAck.Visible = False
  FraAck.Enabled = False
End Sub
Private Sub CmdAckGo_Click()
  With GrdOrdMst
    If .Value(.Row, "wOmAckYn") = "Y" Then DispMsg "Order Already Acknowledged", etInfo: Call CmdAck_Click: Exit Sub

    If .Value(.Row, "wOmCoCd") <> "WEB" Then
      LblNewVchNo.Caption = "Order No " + .Value(.Row, "wOmNumber") + " Will Be Acknowleged. " + vbCrLf + "Click On 'GO' To Continue?"
      ws_NewVchNo = .Value(.Row, "wOmNumber")
    Else
      If ADC("wOmChr") = "" Then DispMsg "Voucher Character Cannot Be Blank", etInfo: Exit Sub
    End If
    
    ' Check if Order No already exists
    If ADC("wOmNo") <> 0 And moCn.RecSeek("Select 'x' from OrdMst where OmPrtKey='" + ctCurrPrtn + "' and OmCoCd= '" + gs_CoCd + "' and " + _
                      "OmTc= '" + .Value(.Row, "wOmTc") + "' and OmYy= '" + .Value(.Row, "wOmYy") + "' and " + _
                      "OmChr= '" + ADC("wOmChr") + "' and OmNo= " + CStr(ADC("wOmNo")) + " ") Then
      DispMsg "Voucher No Already Exists", etInfo: Exit Sub
    End If
    
    ' Check for 'LINE'
    If moCn.RecSeek("select 'x' from OrdDsg " + _
                      "where OdPrtKey='" + ctCurrPrtn + "' and OdCoCd= '" + .Value(.Row, "wOmCoCd") + "' and " + _
                      "OdTc= '" + .Value(.Row, "wOmTc") + "' and OdYy= '" + .Value(.Row, "wOmYy") + "' and " + _
                      "OdChr= '" + .Value(.Row, "wOmChr") + "' and OdNo= " + CStr(.Value(.Row, "wOmNo")) + " and " + _
                      "Not Exists (Select 'x' from vParam where vPCoCd= '" + gs_CoCd + "' and vPTyp= 'LINE' and " + _
                      "vPMCd= OdLine and vPMCd<> 'ZZZ' and vPSCd= '')") Then
      DispMsg "Line Not Defined", etInfo: Exit Sub
    End If
    
    ' Check for 'PRDSEQ'
    If moCn.RecSeek("select 'x' from OrdDsg " + _
                      "where OdPrtKey='" + ctCurrPrtn + "' and OdCoCd= '" + .Value(.Row, "wOmCoCd") + "' and " + _
                      "OdTc= '" + .Value(.Row, "wOmTc") + "' and OdYy= '" + .Value(.Row, "wOmYy") + "' and " + _
                      "OdChr= '" + .Value(.Row, "wOmChr") + "' and OdNo= " + CStr(.Value(.Row, "wOmNo")) + " and " + _
                      "Not Exists (Select 'x' from vParam where vPCoCd= '" + gs_CoCd + "' and vPTyp= 'PRDSEQ' and " + _
                      "vPMCd= OdPrdSeq and vPSCd= '')") Then
      DispMsg "Production Sequence Not Defined", etInfo: Exit Sub
    End If
    
    If LblNewVchNo.Caption = "" Then
      'Generate 'SO' No. from Autogen and display to the user
      If ADC("wOmNo") = 0 And ADC("wOmChr") <> "" Then
        ws_NewVchNo = CStr(moCn.GetFldVal("Select IsNull((select aNum  from Autogen where aTable= 'OrdMst' and " + _
                        "aFld= 'OmNo' and aPreKey= '" + gs_CoCd + .Value(.Row, "wOmTc") + .Value(.Row, "wOmYy") + _
                        ADC("wOmChr") + "'),0) + 1 "))
      Else
        ws_NewVchNo = CStr(ADC("wOmNo"))
      End If

      ws_NewVchNo = gs_CoCd + "/" + .Value(.Row, "wOmTc") + "/" + .Value(.Row, "wOmYy") + "/" + _
                      ADC("wOmChr") + "/" + ws_NewVchNo
      LblNewVchNo = "New Order No Will Be " + ws_NewVchNo + ". " + vbCrLf + "Click On 'GO' To Continue?"
    Else
      'Call Order Acknowledgement Procedure here
      If OrdAckow = True Then
        .Value(.Row, "wOmAckYn") = "Y"
        .Value(.Row, "wOmNoNew") = ws_NewVchNo
      End If

      ws_NewVchNo = 0
      Call CmdAck_Click
    End If
  End With
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
  ws_BtnStr = "CmdCpyGo"
  Call SetObjColors(Me, ws_BtnStr)
  '*** Jenny Colour
End Sub
Private Sub Form_KeyUp(KeyCode As Integer, Shift As Integer)
  If ADC.HotKeys(KeyCode, Shift) Then
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

Private Sub GrdOrdMst_RowWhen(ByVal RowNum As Integer)
  CmdAck.Enabled = True
  Call FillOrdDsg
End Sub
Private Sub GrdOrdMst_Valid()
  'CmdAck.Enabled = False
End Sub
Private Sub GrdOrdMst_Validate(Cancel As Boolean)
  Cancel = GrdOrdMst.Validate
End Sub
Private Sub GrdOrdDsg_Validate(Cancel As Boolean)
  Cancel = GrdOrdDsg.Validate
End Sub
Private Sub FillOrdMst()
  Dim woRecSet As MDORowSet, wSqlStrg As String, wi_Indx As Integer, wRow As Integer
  wSqlStrg = " Select OmCoCd, OmTc, OmCrmYy, OmCrmChr, OmCrmNo, OmDt, OmCmCd, OmPONo, OmPODt, OmAckYN, " + _
              " OmYy, OmChr, OmNo From OrdMst " + _
              " Where /*OmCoCd= 'WEB' and */ OmPrtKey='" + ctCurrPrtn + "' and OmTc= (select PMCd from Param where PTyp= 'TC' and PValue= 'SO') and " + _
              " OmSubmitYn= 'Y' and OmAckYn= 'N' /* and 1= 2 */ " + _
              " Order By OmCoCd, OmTc, OmCrmYy, OmCrmChr, OmCrmNo "

  Set woRecSet = moCn.OpenResultset(wSqlStrg)
  With GrdOrdMst
    .Rows = 1
    If woRecSet.RecCount > 0 Then woRecSet.MoveFirst
    For wi_Indx = 1 To woRecSet.RecCount
      .Rows = .Rows + 1
      wRow = .Rows - 1
      .Value(wRow, "wSr") = wRow
      .Value(wRow, "wOmCrmNo") = woRecSet.FldValue("OmCrmYy") + "/" + woRecSet.FldValue("OmCrmChr") + "/" + CStr(woRecSet.FldValue("OmCrmNo"))
      .Value(wRow, "wOmDt") = woRecSet.FldValue("OmDt")
      .Value(wRow, "wOmCmCd") = woRecSet.FldValue("OmCmCd")
      .Value(wRow, "wOmPONo") = woRecSet.FldValue("OmPONo")
      .Value(wRow, "wOmAckYn") = woRecSet.FldValue("OmAckYN")
      .Value(wRow, "wOmCoCd") = woRecSet.FldValue("OmCoCd")
      .Value(wRow, "wOmTc") = woRecSet.FldValue("OmTc")
      .Value(wRow, "wOmYy") = woRecSet.FldValue("OmYy")
      .Value(wRow, "wOmChr") = woRecSet.FldValue("OmChr")
      .Value(wRow, "wOmNo") = woRecSet.FldValue("OmNo")
      .Value(wRow, "WOMNumber") = woRecSet.FldValue("OmCoCd") + "/" + woRecSet.FldValue("OmTc") + "/" + woRecSet.FldValue("OmYy") + "/" + woRecSet.FldValue("OmChr") + "/" + CStr(woRecSet.FldValue("OmNo"))
      .Value(wRow, "wOmNoNew") = ""
      woRecSet.MoveNext
    Next
    If .Rows > 1 Then .Col = 1
  End With
End Sub
Private Sub FillOrdDsg()
  Dim woRecSet As MDORowSet, wSqlStrg As String, i As Integer, wRow As Integer, wmRow As Integer
  With GrdOrdMst
    wmRow = .Row
    wSqlStrg = " Select OdSr, OdDmCd, OdSfx, OdDmSz, OdOrdEnt, OdCalcPrc, OdSalPrc, OdExpDelDt, OdKT, OdDmCol " + _
                " From OrdDsg " + _
                " Where OdPrtKey='" + ctCurrPrtn + "' and OdCoCd = '" + .Value(wmRow, "wOmCoCd") + "' And OdTc = '" + .Value(wmRow, "wOmTc") + "' " + _
                " And OdYy = '" + .Value(wmRow, "wOmYy") + "' And OdChr = '" + .Value(wmRow, "wOmChr") + "' " + _
                " And OdNo = " + CStr(.Value(wmRow, "wOmNo")) + " " + _
                " Order By OdSr "
  End With
  Set woRecSet = moCn.OpenResultset(wSqlStrg)
  With GrdOrdDsg
    .Rows = 1
    If woRecSet.RecCount > 0 Then woRecSet.MoveFirst
    For i = 1 To woRecSet.RecCount
      .Rows = .Rows + 1
      wRow = .Rows - 1
      .Value(wRow, "wOdSr") = woRecSet.FldValue("OdSr")
      .Value(wRow, "wOdDmCd") = woRecSet.FldValue("OdDmCd")
      .Value(wRow, "wOdSfx") = woRecSet.FldValue("OdSfx")
      .Value(wRow, "wOdDmSz") = woRecSet.FldValue("OdDmSz")
      .Value(wRow, "wOdOrdEnt") = woRecSet.FldValue("OdOrdEnt")
      .Value(wRow, "wOdCalcPrc") = woRecSet.FldValue("OdCalcPrc")
      .Value(wRow, "wOdSalPrc") = woRecSet.FldValue("OdSalPrc")
      'Format(Math.Ceiling(.OdSalPrc * .OdOrdQty), "###0")
      .Value(wRow, "wOdExpDelDt") = woRecSet.FldValue("OdExpDelDt")
      .Value(wRow, "wOdKt") = woRecSet.FldValue("OdKt")
      .Value(wRow, "wOdDmCol") = woRecSet.FldValue("OdDmCol")
      woRecSet.MoveNext
    Next
  End With
End Sub
Private Sub StoreState()
  '*** This procedure is used to store the CausesValidation State and the Enabled/Disabled
      'states of all the Command Buttons
  '*** This is used for restoring the original states in the future
  mArr_Cv(OrdAck) = CmdAck.CausesValidation
  mArr_Ed(OrdAck) = CmdAck.Enabled
End Sub

Private Function OrdAckow() As Boolean
Dim wd_VchNo As Long
On Error GoTo ErrHndlr
Dim wPoNo As String, wIdNo As Double
  With GrdOrdMst
    If ADC("wOmNo") = 0 Then
      wd_VchNo = moCn.AutoGen("OrdMst", "OmNo", gs_CoCd + .Value(.Row, "wOmTc") + .Value(.Row, "wOmYy") + ADC("wOmChr"))
    Else
      wd_VchNo = ADC("wOmNo")
      Call moCn.AutoGen("OrdMst", "OmNo", gs_CoCd + .Value(.Row, "wOmTc") + .Value(.Row, "wOmYy") + ADC("wOmChr"), ADC("wOmNo"))
    End If
  
    moCn.Execute ("Update OrdMst Set OmAckYn= 'Y' where OmPrtKey='" + ctCurrPrtn + "' and OmCoCd= '" + .Value(.Row, "wOmCoCd") + "' and " + _
                  "OmTc= '" + .Value(.Row, "wOmTc") + "' and OmYy= '" + .Value(.Row, "wOmYy") + "' and " + _
                  "OmChr= '" + .Value(.Row, "wOmChr") + "' and OmNo= " + CStr(.Value(.Row, "wOmNo")))
    
    If UCase(.Value(.Row, "wOmCoCd")) = "WEB" Then
    '*** Geeta 16-12-06
    wPoNo = moCn.GetFldVal("Select OmPoNo From OrdMst " + _
        " Where OmPrtKey='" + ctCurrPrtn + "' and OmCoCd= '" + .Value(.Row, "wOmCoCd") + "' and " + _
                  "OmTc= '" + .Value(.Row, "wOmTc") + "' and OmYy= '" + .Value(.Row, "wOmYy") + "' and " + _
                  "OmChr= '" + .Value(.Row, "wOmChr") + "' and OmNo= " + CStr(.Value(.Row, "wOmNo")))
    

    
    moCn.Execute ("Insert Into OrdMst (OmTc,OmYy,OmChr,OmNo,OmCmCd,OmDt,OmLmgCst,OmLmgDt,OmLmgSal, " + _
        " OmDelDt,OmCmDelTerms,OmCmPayTerms, OmPerson,OmPrtCd,OmMulby,OmFixPrc,OmGldAs,OmIWtEqOrd,OmIWtFrOrd, " + _
        " OmIGldRtEqOrd,OmPoNo,OmPoDt,ModUsr,ModDt,ModTime,OmExpDelDt,OmLmpCst,OmLmpDt,OmLmpSal,OmCoCd, " + _
        " OmLine,OmWDiaAvlblDt,OmHDiaAvlblDt,OmFndAvlblDt,OmRefreshDt,OmOrgDelDt,OmOrgExpDelDt,OmHCtCd,OmSalRem,OmHistCmCtg," + _
        " OmHistCmCd,OmHistYyyyMm,OmLockYN,OmLockSalPrcYN,OmLmsCst,OmLmsSal,OmLmsDt,OmCrmYy,OmCrmChr,OmCrmNo,OmSubmitYN,OmAckYN, " + _
        " OmCnvFct,OmCnvFctDt,OmMulDiv,OmCmCurCd,OmLabAs,OmReYy,OmReChr,OmReNo,OmPoCnfDt,OmCustReqDt,OmRevDt1,OmRevDt2,OmPrtKey, " + _
        " OmILabWtFrOrd,OmChgPtrOnStwYN)  " + _
        " Select OmTc,OmYy,'" + ADC("wOmChr") + "'," + CStr(wd_VchNo) + ",OmCmCd,OmDt,OmLmgCst,OmLmgDt,OmLmgSal, " + _
        " OmDelDt,OmCmDelTerms,OmCmPayTerms, OmPerson,OmPrtCd,OmMulby,OmFixPrc,OmGldAs,OmIWtEqOrd,OmIWtFrOrd, " + _
        " OmIGldRtEqOrd,'WEB'+OmTc+OmYy+'" + ADC("wOmChr") + CStr(wd_VchNo) + "',OmPoDt,ModUsr,ModDt,ModTime,OmExpDelDt,OmLmpCst,OmLmpDt,OmLmpSal,'" + gs_CoCd + "', " + _
        " OmLine,OmWDiaAvlblDt,OmHDiaAvlblDt,OmFndAvlblDt,OmRefreshDt,OmOrgDelDt,OmOrgExpDelDt,OmHCtCd,OmSalRem,OmHistCmCtg," + _
        " OmHistCmCd,OmHistYyyyMm,OmLockYN,OmLockSalPrcYN,OmLmsCst,OmLmsSal,OmLmsDt,OmCrmYy,OmCrmChr,OmCrmNo,'','', " + _
        " OmCnvFct,OmCnvFctDt,OmMulDiv,OmCmCurCd,OmLabAs,OmReYy,OmReChr,OmReNo,OmPoCnfDt,OmCustReqDt,OmRevDt1,OmRevDt2,OmPrtKey,OmILabWtFrOrd,OmChgPtrOnStwYN " + _
        " From OrdMst " + _
        " Where OmPrtKey='" + ctCurrPrtn + "' and OmCoCd= '" + .Value(.Row, "wOmCoCd") + "' and " + _
                  "OmTc= '" + .Value(.Row, "wOmTc") + "' and OmYy= '" + .Value(.Row, "wOmYy") + "' and " + _
                  "OmChr= '" + .Value(.Row, "wOmChr") + "' and OmNo= " + CStr(.Value(.Row, "wOmNo")))
    
    
    
    wIdNo = moCn.GetFldVal("Select OmIdNo " + _
                  " From OrdMst " + _
                  " Where OmPrtKey='" + ctCurrPrtn + "' and OmCoCd= '" + gs_CoCd + "' and " + _
                            "OmTc= '" + .Value(.Row, "wOmTc") + "' and OmYy= '" + .Value(.Row, "wOmYy") + "' and " + _
                            "OmChr= '" + ADC("wOmChr") + "' and OmNo= " + CStr(wd_VchNo))


    moCn.Execute ("insert Into OrdDsg (OdTc,OdYy,OdChr,OdNo,OdSr,OdDmCd,OdSfx,OdDmSz,OdPrdSeq, " + _
        " OdDelDt,OdOrdEnt,OdOrdQty,OdExpQty,OdPrdEnt,OdPrdQty,OdFgQty,OdCalcPrc,OdSalPrc,OdCstPrc,OdDmPrdInst," + _
        " OdCmPrdInst,OdCmStmpInst,OdSzInst,OdPrtCd,OdHld,OdHldDesc,OdVaCtg,OdKt,OdMulby,OdFixPrc,OdGldAs," + _
        " OdIWtEqOrd,OdIWtFrOrd,OdIGldRtEqOrd,OdCls,OdBagPcs,OdDmCol,OdGldAsWt,ModUsr,ModDt,ModTime," + _
        " OdExpDelDt,OdSfxDesc,OdCoCd,OdWh,OdWDiaAvlblDt,OdHDiaAvlblDt,OdFndAvlblDt,OdLine,OdSalRem,OdOmCtCd,OdMinWt, " + _
        " OdMaxWt, OdBYy, OdBChr, OdBNo, OdGmChk, OdChkTol, OdGrMet, OdGrDia, OdGrCS, OdGrFin, OdCrmFixPrcYN, " + _
        " OdSubRem,OdLabAs,OdLabAsWt,OdOmCmCd,OdOmDt,OdOmIdNo,OdDmIdNo,OdPrtKey,OdPrtFgQty,OdPrtExpQty,OdPicNm, " + _
        " OdPoNo,OdILabWtFrOrd,OdChgPtrOnStwYN) " + _
        " Select OdTc,OdYy,'" + ADC("wOmChr") + "'," + CStr(wd_VchNo) + ",OdSr,OdDmCd,OdSfx,OdDmSz,OdPrdSeq, " + _
        " OdDelDt,OdOrdEnt,OdOrdQty,OdExpQty,OdPrdEnt,OdPrdQty,OdFgQty,OdCalcPrc,OdSalPrc,OdCstPrc,OdDmPrdInst," + _
        " OdCmPrdInst,OdCmStmpInst,OdSzInst,OdPrtCd,OdHld,OdHldDesc,OdVaCtg,OdKt,OdMulby,OdFixPrc,OdGldAs," + _
        " OdIWtEqOrd,OdIWtFrOrd,OdIGldRtEqOrd,OdCls,OdBagPcs,OdDmCol,OdGldAsWt,ModUsr,ModDt,ModTime," + _
        " OdExpDelDt,OdSfxDesc,'" + gs_CoCd + "',OdWh,OdWDiaAvlblDt,OdHDiaAvlblDt,OdFndAvlblDt,OdLine,OdSalRem,OdOmCtCd,OdMinWt, " + _
        " OdMaxWt, OdBYy, OdBChr, OdBNo, OdGmChk, OdChkTol, OdGrMet, OdGrDia, OdGrCS, OdGrFin, OdCrmFixPrcYN, " + _
        " OdSubRem,OdLabAs,OdLabAsWt,OdOmCmCd,OdOmDt," + CStr(wIdNo) + ",OdDmIdNo,'" + ctCurrPrtn + "',OdPrtFgQty,OdPrtExpQty,OdPicNm, " + _
        " OdPoNo,OdILabWtFrOrd,OdChgPtrOnStwYN " + _
        " From OrdDsg " + _
        " Where OdPrtKey='" + ctCurrPrtn + "' and OdCoCd= '" + .Value(.Row, "wOmCoCd") + "' and " + _
                  "OdTc= '" + .Value(.Row, "wOmTc") + "' and OdYy= '" + .Value(.Row, "wOmYy") + "' and " + _
                  "OdChr= '" + .Value(.Row, "wOmChr") + "' and OdNo= " + CStr(.Value(.Row, "wOmNo")))

    wIdNo = moCn.GetFldVal("Select OdIdNo " + _
                            " From OrdDsg " + _
                            " Where OdPrtKey='" + ctCurrPrtn + "' And OdCoCd= '" + gs_CoCd + "' and " + _
                                      "OdTc= '" + .Value(.Row, "wOmTc") + "' and OdYy= '" + .Value(.Row, "wOmYy") + "' and " + _
                                      "OdChr= '" + ADC("wOmChr") + "' and OdNo= " + CStr(wd_VchNo))


    moCn.Execute ("Insert Into OrdRm (OrTc,OrYy,OrChr,OrNo,OrSr,OrSrNo,OrRmCtg,OrRmSCtg,OrRmCd, " + _
        " OrLn1,OrLn2,OrQty,OrWt,OrSalRt,OrSalVal,OrCstRt,OrCstVal,OrSetScd,OrSetSalRt,OrSetSalVal, " + _
        " OrSetCstRt,OrSetCstVal,OrWsQty,OrHsQty,OrSubShp,OrAlyCd,OrAlySalRt,OrAlyCstRt,OrMainMet,ModUsr," + _
        " ModDt,ModTime,OrRmPtr,OrLmeSal,OrPrdQty,OrPrdWt,OrCoCd,OrOdIdNo,OrPrtKey) " + _
        " Select OrTc,OrYy,'" + ADC("wOmChr") + "'," + CStr(wd_VchNo) + ",OrSr,OrSrNo,OrRmCtg,OrRmSCtg,OrRmCd, " + _
        " OrLn1,OrLn2,OrQty,OrWt,OrSalRt,OrSalVal,OrCstRt,OrCstVal,OrSetScd,OrSetSalRt,OrSetSalVal, " + _
        " OrSetCstRt,OrSetCstVal,OrWsQty,OrHsQty,OrSubShp,OrAlyCd,OrAlySalRt,OrAlyCstRt,OrMainMet,ModUsr," + _
        " ModDt,ModTime,OrRmPtr,OrLmeSal,OrPrdQty,OrPrdWt,'" + gs_CoCd + "'," + CStr(wIdNo) + ",'" + ctCurrPrtn + "'" + _
        " From OrdRm " + _
        " Where OrPrtKey='" + ctCurrPrtn + "' and OrCoCd= '" + .Value(.Row, "wOmCoCd") + "' and " + _
                  "OrTc= '" + .Value(.Row, "wOmTc") + "' and OrYy= '" + .Value(.Row, "wOmYy") + "' and " + _
                  "OrChr= '" + .Value(.Row, "wOmChr") + "' and OrNo= " + CStr(.Value(.Row, "wOmNo")))

      moCn.Execute ("Insert Into OrdLab (OlTc,OlYy,OlChr,OlNo,OlSr,OlSrNo,OlMcd,OlScd,OlQty, " + _
          " OlSalRt,OlSalVal,OlCstRt,OlCstVal,OlQw,ModUsr,ModDt,ModTime,OlCstQw,OlCoCd,OlOdIdNo,OlPrtKey) " + _
          " Select OlTc,OlYy,'" + ADC("wOmChr") + "'," + CStr(wd_VchNo) + ",OlSr,OlSrNo,OlMcd,OlScd,OlQty, " + _
          " OlSalRt,OlSalVal,OlCstRt,OlCstVal,OlQw,ModUsr,ModDt,ModTime,OlCstQw,'" + gs_CoCd + "'," + CStr(wIdNo) + ",'" + ctCurrPrtn + "'" + _
      " From OrdLab " + _
        " Where OlPrtKey='" + ctCurrPrtn + "' and OlCoCd= '" + .Value(.Row, "wOmCoCd") + "' and " + _
                  "OlTc= '" + .Value(.Row, "wOmTc") + "' and OlYy= '" + .Value(.Row, "wOmYy") + "' and " + _
                  "OlChr= '" + .Value(.Row, "wOmChr") + "' and OlNo= " + CStr(.Value(.Row, "wOmNo")))

      moCn.Execute ("Delete From OrdDsg " + _
        " Where OdPrtKey='" + ctCurrPrtn + "' and OdCoCd= '" + .Value(.Row, "wOmCoCd") + "' and " + _
                  "OdTc= '" + .Value(.Row, "wOmTc") + "' and OdYy= '" + .Value(.Row, "wOmYy") + "' and " + _
                  "OdChr= '" + .Value(.Row, "wOmChr") + "' and OdNo= " + CStr(.Value(.Row, "wOmNo")))

      moCn.Execute ("Delete From OrdMst " + _
        " Where OmPrtKey='" + ctCurrPrtn + "' and OmCoCd= '" + .Value(.Row, "wOmCoCd") + "' and " + _
                  "OmTc= '" + .Value(.Row, "wOmTc") + "' and OmYy= '" + .Value(.Row, "wOmYy") + "' and " + _
                  "OmChr= '" + .Value(.Row, "wOmChr") + "' and OmNo= " + CStr(.Value(.Row, "wOmNo")))

      moCn.Execute ("Update OrdMst Set OmPoNo='" + wPoNo + "'" + _
        " Where OmPrtKey='" + ctCurrPrtn + "' and OmCoCd= '" + gs_CoCd + "' and " + _
                  "OmTc= '" + .Value(.Row, "wOmTc") + "' and OmYy= '" + .Value(.Row, "wOmYy") + "' and " + _
                  "OmChr= '" + ADC("wOmChr") + "' and OmNo= " + CStr(wd_VchNo))
          
    '****** Geeta 2.12 [16-12-06]
    '      moCn.Execute ("Update OrdMst Set OmAckYn= '', OmSubmitYn= '', OmCoCd= '" + gs_CoCd + "', " + _
    '                  "OmChr= '" + ADC("wOmChr") + "', OmNo= " + CStr(wd_VchNo) + " " + _
    '                  "where OmCoCd= '" + .Value(.Row, "wOmCoCd") + "' and " + _
    '                  "OmTc= '" + .Value(.Row, "wOmTc") + "' and OmYy= '" + .Value(.Row, "wOmYy") + "' and " + _
    '                  "OmChr= '" + .Value(.Row, "wOmChr") + "' and OmNo= " + CStr(.Value(.Row, "wOmNo")))
    '****** Geeta 16-12-06



      .Value(.Row, "wOmCoCd") = gs_CoCd
      .Value(.Row, "wOmChr") = ADC("wOmChr")
      .Value(.Row, "wOmNo") = wd_VchNo
    End If
  End With
  OrdAckow = True
  
Exit Function
ErrHndlr:
  OrdAckow = False
  If Err.Number = -2147217900 Then
    DispMsg Mid(Err.Description, InStr(1, Err.Description, "|", vbTextCompare) + 1), etError
  Else
    DispMsg Err.Description, etError
  End If
End Function
