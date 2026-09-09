VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Object = "{6A2CE895-C1AC-48AF-B558-FB94BBE12499}#1.0#0"; "mwfctl.ocx"
Begin VB.Form EmrFrmClrDfn 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Colour Master"
   ClientHeight    =   7620
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   11955
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MDIChild        =   -1  'True
   ScaleHeight     =   7620
   ScaleMode       =   0  'User
   ScaleWidth      =   11955
   WindowState     =   2  'Maximized
   Begin MwfCtl.MWCTL_BTN1 CmdSelClr 
      CausesValidation=   0   'False
      Height          =   465
      Left            =   4620
      TabIndex        =   3
      Top             =   7140
      Width           =   915
      _ExtentX        =   1614
      _ExtentY        =   820
      BackColor       =   -2147483633
      ForeColor       =   16777215
      Caption         =   "Sel C&olor"
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
   Begin VB.Frame FRA_KEYS 
      BorderStyle     =   0  'None
      Height          =   300
      Left            =   10380
      TabIndex        =   6
      Top             =   30
      Width           =   1575
      Begin MwfCtl.MWCTL_MED ATXT 
         Height          =   285
         Index           =   1
         Left            =   765
         TabIndex        =   4
         ToolTipText     =   "Enter Customer Code"
         Top             =   0
         Visible         =   0   'False
         Width           =   630
         _ExtentX        =   1111
         _ExtentY        =   503
         MaxLength       =   5
         DataType        =   4
         IdName          =   "CSCD"
      End
      Begin VB.Label ALBL 
         BackStyle       =   0  'Transparent
         Caption         =   "Range"
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
         Index           =   0
         Left            =   30
         TabIndex        =   7
         Top             =   15
         Visible         =   0   'False
         Width           =   945
      End
   End
   Begin MwfCtl.MWCTL_HLP MWCTL_HLP1 
      Height          =   405
      Left            =   6900
      TabIndex        =   1
      TabStop         =   0   'False
      Top             =   7215
      Visible         =   0   'False
      Width           =   555
      _ExtentX        =   979
      _ExtentY        =   714
   End
   Begin MwfCtl.MWCTL_ADC ADC 
      CausesValidation=   0   'False
      Height          =   555
      Left            =   30
      TabIndex        =   0
      Top             =   7080
      Width           =   10305
      _ExtentX        =   18177
      _ExtentY        =   979
      AllowAdd        =   0   'False
      AllowDelete     =   0   'False
      BtnWidth        =   900
      BtnHeight       =   465
      BtnMaskColor    =   8454143
      FontSz          =   8.25
   End
   Begin VB.Frame FRA_NKEYS 
      BorderStyle     =   0  'None
      Enabled         =   0   'False
      Height          =   7455
      Left            =   -120
      TabIndex        =   5
      Top             =   0
      Width           =   12045
      Begin VB.Frame FraNKeyAll 
         BorderStyle     =   0  'None
         Height          =   7035
         Left            =   150
         TabIndex        =   8
         Top             =   0
         Width           =   11865
         Begin VB.PictureBox PicClr 
            Height          =   705
            Left            =   6690
            ScaleHeight     =   645
            ScaleWidth      =   735
            TabIndex        =   10
            Top             =   5970
            Visible         =   0   'False
            Width           =   795
         End
         Begin VB.Frame Picture1 
            Height          =   1785
            Left            =   6810
            TabIndex        =   9
            Top             =   120
            Visible         =   0   'False
            Width           =   1095
            Begin MSComDlg.CommonDialog CommonDialog1 
               Left            =   0
               Top             =   0
               _ExtentX        =   847
               _ExtentY        =   847
               _Version        =   393216
            End
         End
         Begin MwfCtl.MWCTL_FLX GrdClr 
            Height          =   6900
            Left            =   0
            TabIndex        =   2
            Top             =   90
            Width           =   10305
            _ExtentX        =   18177
            _ExtentY        =   12171
            Cols            =   6
            Pic             =   "EmrFrmClrDfn.frx":0000
            colname1        =   "WCFSR"
            heading1        =   "Sr.      "
            datatype1       =   1
            mask1           =   "##0"
            recalcon1       =   "GrdClr(CFCD)"
            maxlength1      =   3
            style1          =   2
            tooltiptext1    =   "Sr No."
            colname2        =   "CFCD"
            heading2        =   "Color Cd              "
            datafld2        =   "CfCd"
            datatype2       =   4
            recalcparent2   =   "GrdClr(WCFCLRPIC),GrdClr(WCFSR)"
            maxlength2      =   12
            tooltiptext2    =   "Enter Colour Code"
            colname3        =   "WCFCLRPIC"
            heading3        =   "Color            "
            recalcon3       =   "GrdClr(CFCD)"
            style3          =   2
            tooltiptext3    =   "Color"
            colname4        =   "CFSHTNAME"
            heading4        =   "Short Name                                  "
            datafld4        =   "CfShtName"
            maxlength4      =   20
            tooltiptext4    =   "Enter Short Name"
            colname5        =   "CFNAME"
            heading5        =   "Name                                                                      "
            datafld5        =   "CfName"
            maxlength5      =   30
            tooltiptext5    =   "Enter Long Name Of The Color"
         End
      End
   End
End
Attribute VB_Name = "EmrFrmClrDfn"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'----   General Comments on the Entry Form  -------'
'
'
'----


Option Explicit
Dim moCn As MwfLib.MDOConnection   '*** (Bef speed) bef 24/11/06


Private Sub ADC_Load()
  '*** Set the First Key Control to Customer Code
  '*** Set the First Non Key Control to Customer Name
  '*** Set the Child Property Of ADC to GrdClr,GrdCustAly
      '(i.e. both GrdClr and GrdCustAly are children of ADC)
  '*** Set the Previous and Next Control Properties for ADC and the Grids GrdClr, GrdCustRm and GrdCustAly
  '*** Set the frame FraNKeyAll as the top most frame
  '*** Set the current tab as 0 (This is done so that whatever be the active tab during
      'design time, the first tab is set as the active tab at run time)
  
  Set moCn = adc.Connection    '*** (Bef speed) bef 24/11/06
  'Set ADC.FirKeyCtl = ADC("CsCd")
  Set adc.FirKeyCtl = adc
  Set adc.FirNKeyCtl = GrdClr
           
  adc.Child = "GrdClr"
    
  GrdClr.HotKey = "A"
  FraNKeyAll.ZOrder (0)
    
End Sub
Private Sub ADC_KeyWhen()
  '*** Set the caption for the Form based on the Menu Code selected
  '*** If Menu Code is CustMst the Caption= 'Customer Master'
  '*** If Menu Code is SuppMst the Caption= 'Supplier Master'
  '*** If Menu Code is PLMst the Caption= 'Price List Customer Master'
  '*** Disable all the command buttons (Because since the focus is on the key frame and
      'the user accidently clicks any of the command buttons then it might give a
      'run time error. So it is better to disable all the command buttons when the focus is
      'in the Key Frame.
  '*** Hide all frames except the frame FraNKeyAll
  
  CmdSelClr.Enabled = False
End Sub
Private Sub ADC_NonKeyWhen(Cancel As Boolean, ErrMsg As String)
  '*** Enable the command button CmdSelClr
  CmdSelClr.Enabled = True
End Sub
Private Sub ATXT_FldWhen(Index As Integer, ByVal IdName As String, Cancel As Boolean, ErrMsg As String)
'  '*** Set Helps For Each Field
'  Select Case UCase(IdName)
'  Case Is = UCase("CmCd")
'    Call HlpList.CustCd(ADC("CmCtg"))
'  End Select
End Sub
Private Sub atxt_ReCalculate(Index As Integer, ByVal IdName As String, ByVal OldVal As Variant, ByVal RowNum As Integer, ByVal SourceName As String, ByVal SourceColName As String)
'  '*** CmIWtFrOrd on CmIWtEqOrd         Normal
'      'If CmIWtEqOrd= 'Y' then ''
'  Select Case UCase(IdName)
'  Case Is = UCase("CmIWtFrOrd")
'    If ADC.Mode = xNorm Then
'      If ADC("CmIWtEqOrd") = "Y" Then ADC("CmIWtFrOrd") = ""
'    End If
'  End Select
End Sub
Private Sub atxt_FldChange(Index As Integer, ByVal IdName As String, Cancel As Boolean, ErrMsg As String, ByVal pv_NewValue As Variant)
'  '*** CmCd             Not Blank
'
'  Select Case UCase(IdName)
'  Case Is = UCase("CmCd")
'    If Trim(pv_NewValue) = "" Then Cancel = True: ErrMsg = "Customer Code Cannot Be Blank": Exit Sub
'    '*** For Ctg = "Z"
'    If ADC("CmCtg") = "Z" Then
'      Cancel = Not moCn.RecSeek("select HCoCd from Head where HCd = '" + ctSelfCmCd + "' and " + _
'                                "HCoCd = '" + pv_NewValue + "'")
'      ErrMsg = "Invalid Company Code": Exit Sub
'    End If
'  End Select
End Sub
Private Sub GrdClr_RowWhen(ByVal RowNum As Integer)
  If GrdClr.IsNew(RowNum) Then GrdClr.Value(RowNum, "wCfSr") = GrdClr.MaxVal("wCfSr") + 1
End Sub
Private Sub GrdClr_ColWhen(ByVal RowNum As Integer, ByVal ColName As String, Cancel As Boolean, ErrMsg As String, ByVal Action As MwfCtl.en_fgAction)
'  '*** Set Helps For Each Field
'  Select Case ColName
'  Case Is = UCase("CdOurDmCd")
'    Call HlpList.DmCd("DM")
'  End Select
End Sub
Private Sub GrdClr_ColChange(ByVal RowNum As Integer, ByVal ColName As String, ByVal NewValue As Variant, Cancel As Boolean, ErrMsg As String)
  '*** CfCd           valid value from 0 to 16777215 (&HFFFFFF)

  With GrdClr
    Select Case ColName
    Case Is = UCase("CfCd")
      Cancel = Val(NewValue) < 0 Or Val(NewValue) > 16777215
      'Cancel = Val(NewValue) > 16777215
      ErrMsg = "Invalid Color Code": Exit Sub
    End Select
  End With
End Sub
Private Sub GrdClr_ReCalc(ByVal RowNum As Integer, ByVal ColName As String, ByVal OldVal As Variant, ByVal SourceName As String, ByVal SourceColName As String)
  With GrdClr
    Select Case UCase(ColName)
    Case Is = UCase("wCfSr")
        If .Mode = fgminit Then
          .Value(RowNum, "wCfSr") = .MaxVal("wCfSr") + 1    '.Rows - 1
        End If
    Case Is = UCase("wCfClrPic")
        If .Mode = fgminit Or .Mode = fgmnorm Then
          If Trim(.Value(RowNum, "CfCd")) <> "" Then .SetCellPic RowNum, 3, CreateColorImage(PicClr, Val(.Value(RowNum, "CfCd")), 52, 12)
        End If
    End Select
  End With
End Sub
Private Sub GrdClr_InitKey(ByVal RowNum As Integer)
'  '*** Initialise Key Fields
'  With GrdClr
'    .Store "CdCtg", ADC("CmCtg")
'  End With
End Sub
Private Sub ADC_NonKeyValid(pr_Cancel As Boolean, pr_ErrMsg As String)
  '*** Set the current tab of the tabpage as 0
  'TabCustDet.Tab = 0
End Sub
Private Sub ADC_SetRecSource()
  '*** No RecSource for ADC
  adc.DataMode = xdatUnbound
End Sub
Private Sub GrdClr_SetRecSource()
  '*** Set the RecSource for GrdClr (Base Table is CustDsg)
  GrdClr.RecSource = "Select * from ClrDfn "
End Sub
Private Sub CmdSelClr_Click()
  If GrdClr.Rows - 1 > 0 Then
    Call GetColor
    GrdClr.SetFocus
  End If
End Sub
Private Sub DispFra(ByVal pv_CustFra As en_CustFra)
'  Select Case pv_CustFra
'  Case Is = Cpy
'    '*** If Frame PicBoxCpy is visible then make it invisible and enable all the command buttons
'    '*** If Frame PicBoxCpy is invisible then make it visible and disable all the
'        'command buttons except the one which is clicked. Change the background colour of
'        'the clicked command button to white and make the font bold
'    If PicBoxCpy.Visible = True Then
'      PicBoxCpy.Visible = False
'      PicBoxCpy.Enabled = False
'      Call EnaDisaCmds(False, CmdCpy)
'      CmdCpy.SetFocus
'    Else
'      PicBoxCpy.Visible = True
'      PicBoxCpy.Enabled = True
'      ADC("wCustFr").SetFocus
'      PicBoxCpy.ZOrder
'      Call EnaDisaCmds(True, CmdCpy)
'    End If
'  End Select
End Sub
Private Sub EnaDisaCmds(ByVal pv_ShowFra As Boolean, Optional ByVal pv_CmdBut As MWCTL_BTN1)
'  '*** If second parameter is not given and pv_ShowFra = True then the frame FraNKeyAll
'      'on Frame Fra_nkeys is disabled and allowadd and allowdelete options of adc are
'      'set to false
'
'  '*** If second parameter is not given and pv_ShowFra = False then the frame FraNKeyAll
'      'on Frame Fra_nkeys is enabled and allowadd and allowdelete options of adc are
'      'set to true
'
'  '*** If second parameter is given and pv_ShowFra = True then the frame FraNKeyAll
'      'on Frame Fra_nkeys is disabled and allowadd and allowdelete options of adc are
'      'set to false and for the command button specified in the second parameter,
'      'the BackColor is changed to white and the font is made Bold
'
'  '*** If second parameter is given and pv_ShowFra = False then the frame FraNKeyAll
'      'on Frame Fra_nkeys is enabled and allowadd and allowdelete options of adc are
'      'set to true and for the command button specified in the second parameter,
'      'the BackColor is changed to normal and the font is made Regular
'
'
'  If Not pv_CmdBut Is Nothing Then
'    If pv_ShowFra = True Then
'      '### pv_CmdBut.BackColor = vbWhite
'      '### pv_CmdBut.FontBold = True
'      pv_CmdBut.CausesValidation = True
'    Else
'      '### pv_CmdBut.BackColor = &H8000000F
'      '### pv_CmdBut.FontBold = False
'      'pv_CmdBut.CausesValidation = False
'    End If
'  End If
'  FraNKeyAll.Enabled = Not pv_ShowFra
'  ADC.AllowSave = Not pv_ShowFra
'  ADC.AllowDelete = Not pv_ShowFra
'  CmdCpy.Enabled = Not pv_ShowFra
'  If Not pv_CmdBut Is Nothing Then pv_CmdBut.Enabled = True
End Sub
Private Sub HideAllFras()
'  '*** This procedure hides all the frames other than FraNKeyAll
'  FraNKeyAll.Visible = True
'  PicBoxCpy.Visible = False
'  PicBoxCpy.Enabled = False
End Sub
Private Sub GetColor()
   '*** Set Cancel to True.
   CommonDialog1.CancelError = True
   On Error GoTo ErrHandler
   '*** Set the Flags property.
   CommonDialog1.Flags = cdlCCRGBInit
   '*** Display the Color dialog box.
   CommonDialog1.ShowColor
   '*** Set the form's background color to the selected color.
   
   GrdClr.Value(GrdClr.Row, "CfCd") = CommonDialog1.Color
   
   Exit Sub

ErrHandler:
   '*** User pressed Cancel button.
   Exit Sub
   
End Sub


'''  ***** Std Code not to be changed **********************************
Private Sub Form_GotFocus()
  '*** Jenny Color (For Setting System Colors)
  SetMDISysColorsAfterLogin
  '*** Jenny Color (For Setting System Colors)
End Sub
Private Sub Form_Load()
  '*** Jenny Colour
  FraNKeyAll.Height = fra_nkeys.Height + 500
  FraNKeyAll.Width = fra_nkeys.Width + 500
  
  '(Bef 13/08/05)
  'SetMDISysColorsAfterLogin   '*** Jenny Color (For Setting System Colors)
  Call SetObjColors(Me)
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
Private Sub GrdClr_Validate(Cancel As Boolean)
  Cancel = GrdClr.Validate
End Sub
