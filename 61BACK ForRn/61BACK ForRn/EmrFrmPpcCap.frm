VERSION 5.00
Object = "{6A2CE895-C1AC-48AF-B558-FB94BBE12499}#1.0#0"; "mwfctl.ocx"
Begin VB.Form EmrFrmPpcCap 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Production  Capacity"
   ClientHeight    =   9810
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   15210
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MDIChild        =   -1  'True
   ScaleHeight     =   9810
   ScaleMode       =   0  'User
   ScaleWidth      =   15210
   WindowState     =   2  'Maximized
   Begin MwfCtl.MWCTL_HLP MWCTL_HLP1 
      Height          =   405
      Left            =   3090
      TabIndex        =   1
      TabStop         =   0   'False
      Top             =   9240
      Visible         =   0   'False
      Width           =   555
      _ExtentX        =   979
      _ExtentY        =   714
   End
   Begin MwfCtl.MWCTL_ADC ADC 
      CausesValidation=   0   'False
      Height          =   555
      Left            =   150
      TabIndex        =   2
      Top             =   9270
      Width           =   7590
      _ExtentX        =   13388
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
      Height          =   180
      Left            =   -36
      TabIndex        =   4
      Top             =   24
      Width           =   9060
      Begin MwfCtl.MWCTL_MED ATXT 
         Height          =   285
         Index           =   5
         Left            =   480
         TabIndex        =   6
         ToolTipText     =   "This field is put on the entry form because there should be atleast one atxt field in the form"
         Top             =   0
         Visible         =   0   'False
         Width           =   540
         _ExtentX        =   953
         _ExtentY        =   503
         Enabled         =   0   'False
         MaxLength       =   8
         DataType        =   4
         IdName          =   "WATXT"
      End
   End
   Begin VB.Frame FRA_NKEYS 
      BorderStyle     =   0  'None
      Enabled         =   0   'False
      Height          =   9330
      Left            =   0
      TabIndex        =   3
      Top             =   0
      Width           =   8565
      Begin VB.Frame FraNKeyAll 
         BorderStyle     =   0  'None
         Height          =   9105
         Left            =   150
         TabIndex        =   5
         Top             =   90
         Width           =   7800
         Begin MwfCtl.MWCTL_FLX GrdPpcCap 
            Height          =   8880
            Left            =   0
            TabIndex        =   0
            Top             =   210
            Width           =   7605
            _ExtentX        =   13414
            _ExtentY        =   15663
            Cols            =   7
            colname1        =   "PCLINE"
            heading1        =   "Line        "
            datafld1        =   "PcLine"
            datatype1       =   4
            maxlength1      =   5
            tooltiptext1    =   "Enter Production Line"
            colname2        =   "PCP1"
            heading2        =   "P1              "
            datafld2        =   "PcP1"
            datatype2       =   2
            mask2           =   "######0.0"
            maxlength2      =   9
            tooltiptext2    =   "Enter P1 Capacity Per Day For The Line"
            colname3        =   "PCP2"
            heading3        =   "P2              "
            datafld3        =   "PcP2"
            datatype3       =   2
            mask3           =   "######0.0"
            maxlength3      =   9
            tooltiptext3    =   "Enter P2 Capacity Per Day For The Line"
            colname4        =   "PCP3"
            heading4        =   "P3              "
            datafld4        =   "PcP3"
            datatype4       =   2
            mask4           =   "######0.0"
            maxlength4      =   9
            tooltiptext4    =   "Enter P3 Capacity Per Day For The Line"
            colname5        =   "PCP4"
            heading5        =   "P4              "
            datafld5        =   "PcP4"
            datatype5       =   2
            mask5           =   "######0.0"
            maxlength5      =   9
            tooltiptext5    =   "Enter P4 Capacity Per Day For The Line"
            colname6        =   "PCP5"
            heading6        =   "P5              "
            datafld6        =   "PcP5"
            datatype6       =   2
            mask6           =   "######0.0"
            maxlength6      =   9
            tooltiptext6    =   "Enter P5 Capacity Per Day For The Line"
         End
      End
   End
End
Attribute VB_Name = "EmrFrmPpcCap"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit


Dim moCn As MwfLib.MDOConnection   '*** (Bef speed) bef 24/11/06
Private Sub ADC_Load()
  '*** Set the First Key Control property to adc
  '*** Set the First Non Key Control property to the grid GrdPpcCap
  '*** Set the Child Property Of ADC to GrdPpcCap
  '*** Set the Previous and Next Control Properties for the Grid GrdPpcCap
  '*** Set the Hot Key for the Grid GrdPpcCap as 'A'

  Set moCn = adc.Connection  '*** (Bef speed) bef 24/11/06

  Set adc.FirKeyCtl = adc
  Set adc.FirNKeyCtl = GrdPpcCap
  adc.Child = "GrdPpcCap"
  Set GrdPpcCap.PrevCtl = adc
  Set GrdPpcCap.NextCtl = adc
  GrdPpcCap.HotKey = "A"
  
  '****** Sachin 2.14.0 - [24-11-2006] - [Multi-Company] ******
  Me.Caption = GetFrmCaption(gs_CoCd, "Production  Capacity")
  '****** Sachin 2.14.0 - [24-11-2006] - [Multi-Company] ******
  
  Call Form_Activate
  Call GrdPropSetPPC(GrdPpcCap, "Pc")   '*** (Jen 2.14 Next)
End Sub
Private Sub ADC_KeyValid(pr_Cancel As Boolean, pr_ErrMsg As String)
  If Not moCn.RecSeek("Select 'x' from Head where HCoCd= '" + gs_CoCd + "' and " + _
         "HCd= '" + ctSelfCmCd + "' and hPpcYN= 'Y'") Then
    pr_Cancel = True: pr_ErrMsg = "Cannot Edit This Entry As Permission For PPC Module Not Given": Exit Sub
  End If
End Sub
Private Sub ADC_SetRecSource()
  '*** Since this is an 'A' Entry, there is no RecSource associated with ADC. So set the
      'DataMode property of ADC control as xdatUnbound

  adc.DataMode = xdatUnbound
End Sub
Private Sub GrdPpcCap_ColChange(ByVal RowNum As Integer, ByVal ColName As String, ByVal NewValue As Variant, Cancel As Boolean, ErrMsg As String)
  'If NewValue = "" Then Cancel = True: ErrMsg = "" + Trim(GrdPpcCap.ColProp(ColName).Heading) + " Cannot Be Blank": Exit Sub
  
  Select Case UCase(ColName)
  Case Is = UCase("PcLine")
    If NewValue <> "" Then
      Cancel = Not moCn.RecSeek("Select vPMCd From vParam Where " + _
                           "vPCoCd='" + gs_CoCd + "' and vPTyp= 'Line' And vPMCd = '" + NewValue + "'")
      ErrMsg = "Invalid Line No": Exit Sub
    Else
      Cancel = True
      ErrMsg = "Invalid Line No": Exit Sub
    End If
    
  '*** (Jen 2.14 Next)
  Case Is = UCase("PcP1"), UCase("PcP2"), UCase("PcP3"), UCase("PcP4"), UCase("PcP5")
  '*** (Jen 2.14 Next)
  
  '*** (Bef 2.14 Next)
  'Case Is = UCase("PcCas"), UCase("PcFin"), UCase("PcEi")
  '  'UCase("PcWSet"), UCase("PcHSet")
  '*** (Bef 2.14 Next)
  
    If NewValue <= 0 Then
      Cancel = True
      ErrMsg = "Value of " + Trim(GrdPpcCap.ColProp(ColName).Heading) + " Should Be > 0": Exit Sub
    End If
  End Select
End Sub
Private Sub GrdPpcCap_ColWhen(ByVal RowNum As Integer, ByVal ColName As String, Cancel As Boolean, ErrMsg As String, ByVal Action As MwfCtl.en_fgAction)
  Select Case UCase(ColName)
  Case Is = UCase("PcLine")
    Call HlpList.vPMCd(gs_CoCd, "Line")
  End Select
End Sub

Private Sub GrdPpcCap_InitKey(ByVal RowNum As Integer)
  GrdPpcCap.Store "PcCoCd", gs_CoCd
End Sub

Private Sub GrdPpcCap_SetRecSource()
  '*** Set the Record Source of the Grid GrdPpcCap
      'Default sort option is PcLine
  GrdPpcCap.RecSource = " Select * from PpcCap  where PcCoCd= '" + gs_CoCd + "'Order By PcLine"
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
  ws_BtnStr = ""
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
Private Sub GrdPpcCap_Validate(Cancel As Boolean)
  Cancel = GrdPpcCap.Validate
End Sub



