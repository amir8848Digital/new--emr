VERSION 5.00
Object = "{6A2CE895-C1AC-48AF-B558-FB94BBE12499}#1.0#0"; "mwfctl.ocx"
Begin VB.Form EmrFrmCpyRt 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Copy Rm & Lab Rates To History"
   ClientHeight    =   2625
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   8265
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MDIChild        =   -1  'True
   ScaleHeight     =   2625
   ScaleWidth      =   8265
   Begin MwfCtl.MWCTL_BTN1 CmdDelRt 
      CausesValidation=   0   'False
      Height          =   585
      Left            =   4110
      TabIndex        =   1
      ToolTipText     =   "Option To Reset The Autogen Table"
      Top             =   570
      Width           =   2145
      _ExtentX        =   3784
      _ExtentY        =   1032
      BackColor       =   -2147483633
      MaskColor       =   8454143
      Caption         =   "&Delete Rates From History"
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
   Begin MwfCtl.MWCTL_BTN1 CmdCpRt 
      CausesValidation=   0   'False
      Height          =   585
      Left            =   1515
      TabIndex        =   0
      ToolTipText     =   "Option To Reset The Autogen Table"
      Top             =   570
      Width           =   2145
      _ExtentX        =   0
      _ExtentY        =   0
      BackColor       =   -2147483633
      MaskColor       =   8454143
      Caption         =   "&Copy Rates To History"
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
      Height          =   435
      Left            =   6780
      TabIndex        =   6
      TabStop         =   0   'False
      Top             =   2190
      Visible         =   0   'False
      Width           =   435
      _ExtentX        =   767
      _ExtentY        =   767
   End
   Begin MwfCtl.MWCTL_ADC ADC 
      Height          =   555
      Left            =   7290
      TabIndex        =   7
      Top             =   2190
      Width           =   1950
      _ExtentX        =   3440
      _ExtentY        =   979
      AllowAdd        =   0   'False
      AllowFind       =   0   'False
      AllowSave       =   0   'False
      AllowDelete     =   0   'False
      BtnWidth        =   885
      BtnHeight       =   435
      BtnMaskColor    =   8454143
      FontSz          =   8.25
   End
   Begin VB.Frame FRA_KEYS 
      BorderStyle     =   0  'None
      Height          =   390
      Left            =   -36
      TabIndex        =   9
      Top             =   24
      Width           =   1350
   End
   Begin VB.Frame FRA_NKEYS 
      BorderStyle     =   0  'None
      Enabled         =   0   'False
      Height          =   390
      Left            =   -120
      TabIndex        =   8
      Top             =   510
      Width           =   1425
   End
   Begin VB.Frame FraDelRt 
      Height          =   1065
      Left            =   1920
      TabIndex        =   12
      Top             =   1320
      Width           =   3900
      Begin MwfCtl.MWCTL_BTN1 CmdDelGo 
         Height          =   495
         Left            =   2880
         TabIndex        =   5
         ToolTipText     =   "Press This Button To Delete"
         Top             =   495
         Width           =   555
         _ExtentX        =   979
         _ExtentY        =   873
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
         Index           =   0
         Left            =   1110
         TabIndex        =   4
         ToolTipText     =   "Enter Copy Year And Month"
         Top             =   495
         Width           =   1320
         _ExtentX        =   2328
         _ExtentY        =   503
         Alignment       =   2
         Mask            =   "#####0"
         MaxLength       =   6
         DataType        =   1
         IdName          =   "WDELYYYYMM"
      End
      Begin VB.Label LblMain 
         BackStyle       =   0  'Transparent
         BorderStyle     =   1  'Fixed Single
         Caption         =   "  Delete"
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
         Left            =   0
         TabIndex        =   14
         Top             =   90
         Width           =   3870
      End
      Begin VB.Label ALBL 
         BackStyle       =   0  'Transparent
         Caption         =   "YyyyMm"
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
         Left            =   150
         TabIndex        =   13
         Top             =   495
         Width           =   975
      End
   End
   Begin VB.Frame FraCpRt 
      Height          =   795
      Left            =   1920
      TabIndex        =   10
      Top             =   1320
      Width           =   3900
      Begin MwfCtl.MWCTL_BTN1 CmdCpGo 
         Height          =   285
         Left            =   2880
         TabIndex        =   3
         ToolTipText     =   "Press This Button To Copy"
         Top             =   285
         Width           =   555
         _ExtentX        =   0
         _ExtentY        =   0
         MaskColor       =   8454143
         Caption         =   "&Go"
         Pic             =   "EmrFrmCpyRt.frx":0000
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
         Index           =   11
         Left            =   1110
         TabIndex        =   2
         ToolTipText     =   "Enter Copy Year And Month"
         Top             =   285
         Width           =   1320
         _ExtentX        =   2328
         _ExtentY        =   503
         Alignment       =   2
         Mask            =   "#####0"
         MaxLength       =   6
         DataType        =   1
         IdName          =   "WCPYYYYMM"
      End
      Begin VB.Label LblMain 
         BackStyle       =   0  'Transparent
         BorderStyle     =   1  'Fixed Single
         Caption         =   "  Copy"
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
         TabIndex        =   15
         Top             =   -30
         Width           =   3900
      End
      Begin VB.Label ALBL 
         BackStyle       =   0  'Transparent
         Caption         =   "YyyyMm"
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
         Index           =   10
         Left            =   150
         TabIndex        =   11
         Top             =   285
         Width           =   975
      End
   End
End
Attribute VB_Name = "EmrFrmCpyRt"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim moCn As MDOConnection    '*** (Bef speed) bef 24/11/06

'*** For the Toggle Options of Command Buttons
Enum en_CpyRtFra
  Cpy = 0
  Del = 1
End Enum
Private Sub ATXT_FldWhen(Index As Integer, ByVal IdName As String, Cancel As Boolean, ErrMsg As String)
  Select Case UCase(IdName)
  Case Is = UCase("wCpYyyyMm"), UCase("wDelYyyyMm")
    HlpList.YyyyMm
  End Select
End Sub
Private Sub atxt_FldChange(Index As Integer, ByVal IdName As String, Cancel As Boolean, ErrMsg As String, ByVal pv_NewValue As Variant)
  '*** wCpYyyyMm       Valid Year & Month (Year should be between 1900 & 2099, Month between 1 & 12)
  '*** wDelYyyyMm      Valid Year & Month (Year should be between 1900 & 2099, Month between 1 & 12)
  
  Dim wYear As Integer, wMonth As Integer, wsql As String
  
  Select Case UCase(IdName)
  Case Is = UCase("wCpYyyyMm"), UCase("wDelYyyyMm")
    If pv_NewValue = 0 Then Exit Sub
    If Len(CStr(pv_NewValue)) <> 6 Then Cancel = True: ErrMsg = "Invalid Year & Month": Exit Sub
    wYear = Val(Mid(CStr(pv_NewValue), 1, 4))
    If wYear < 1900 Or wYear > 2099 Then Cancel = True: ErrMsg = "Invalid Year": Exit Sub
    wMonth = Val(Mid(CStr(pv_NewValue), 5, 2))
    If wMonth < 1 Or wMonth > 12 Then Cancel = True: ErrMsg = "Invalid Month": Exit Sub
  End Select
End Sub
Private Sub CmdCpRt_Click()
  '*** Display or hide the Copy frame depending on the toggle state
  Call DispFra(Cpy)
End Sub
Private Sub CmdCpGo_Click()
  '*** Call the Copy Routine and Close the Copy Frame
  Call CpyRatesToHist
End Sub
Private Sub CmdDelRt_Click()
  '*** Display or hide the Delete frame depending on the toggle state
  Call DispFra(Del)
End Sub
Private Sub CmdDelGo_Click()
  '*** Call the Delete Routine and Close the Delete Frame
  Call DelRatesFromHist
End Sub
Private Sub CpyRatesToHist()
  '*** Check if copy is possible
  '*** Check if record alrhheady exists for the Year & Month
  '*** The Rate TcTypes which are copied to history are: 'INC', 'INL', 'LMG', 'LMP', 'LS', 'RM'
  
  Dim wsql As String
  If adc("wCpYyyyMm") = 0 Then DispMsg "Enter Year && Month", etError: Exit Sub
  If moCn.RecSeek("Select 'x' from RmRtHist where RhYyyyMm= " + CStr(adc("wCpYyyyMm")) + " ") Or _
     moCn.RecSeek("Select 'x' from LabRtHist where LhYyyyMm= " + CStr(adc("wCpYyyyMm")) + " ") Then _
    DispMsg "Rates Already Exist In History For The Year && Month - " + CStr(adc("wCpYyyyMm")), etError: Exit Sub
  
  MWLib.BeginProcess Me, "Copying Rm && Lab Rates To History For the Year && Month - " + CStr(adc("wCpYyyyMm"))
  wsql = wsql + " Insert Into RmRtHist (RhYyyyMm, RhTcTyp, RhCmCtg, RhCmCd, RhCmCurCd, " + _
         " RhCtg, RhSCtg, RhSTWGrp, RhCd, RhDmCtg, RhLsCtg, RhLabMCd, RhRngSz, RhFrLn, RhToLn, " + _
         " RhCstRt, RhSalRt, ModUsr, ModDt, ModTime, RhXWtPer) " + _
         " Select " + CStr(adc("wCpYyyyMm")) + ", RrTcTyp, RrCmCtg, RrCmCd, RrCmCurCd, RrCtg, " + _
         " RrSCtg, RrSTWGrp, RrCd, RrDmCtg, RrLsCtg, RrLabMCd, RrRngSz, RrFrLn, RrToLn, " + _
         " RrCstRt, RrSalRt, ModUsr, ModDt, ModTime, RrXWtPer from RmRt " + _
         " where RrTcTyp in ('INC', 'INL', 'LMG', 'LMP', 'LS', 'RM') " + _
         " Order by RrTcTyp, RrCmCtg, RrCmCd, RrCmCurCd, RrCtg, RrSCtg, RrSTWGrp, RrCd, " + _
         " RrDmCtg, RrLsCtg, RrLabMCd, RrFrLn, RrToLn " + _
         " Insert Into LabRtHist (LhYyyyMm, LhCmCtg, LhCmCd, LhCmCurCd, LhMCd, LhSCd, LhQw, " + _
         " LhCstRt, LhSalRt, LhSalMin, LhSalMax, ModUsr , ModDt, ModTime, LhFrWt, LhToWt) " + _
         " Select " + CStr(adc("wCpYyyyMm")) + ", LrCmCtg, LrCmCd, LrCmCurCd, LrMCd, LrSCd, " + _
         " LrQw, LrCstRt, LrSalRt, LrSalMin, LrSalMax, ModUsr, ModDt, ModTime, LrFrWt, LrToWt  from LabRt " + _
         " Order by LrCmCtg, LrCmCd, LrCmCurCd, LrMCd, LrSCd "
    
  moCn.Execute (wsql)
  MWLib.EndProcess Me
  DispMsg "Completed Copying Rates To History", etInfo
  Call CmdCpRt_Click
End Sub
Private Sub DelRatesFromHist()
  
  '*** Check if delete is possible
  '*** Check if record exists for the Year & Month
  Dim wsql As String
  If adc("wDelYyyyMm") = 0 Then DispMsg "Enter Year && Month", etError: Exit Sub
  If Not (moCn.RecSeek("Select 'x' from RmRtHist where RhYyyyMm= " + CStr(adc("wDelYyyyMm")) + " ")) And _
     Not (moCn.RecSeek("Select 'x' from LabRtHist where LhYyyyMm= " + CStr(adc("wDelYyyyMm")) + " ")) Then _
    DispMsg "No Records Existing In History For The Year & Month - " + CStr(adc("wDelYyyyMm")), etError: Exit Sub
  
  MWLib.BeginProcess Me, "Deleting Rm && Lab Rates From History For the Year && Month - " + CStr(adc("wDelYyyyMm"))
  
  wsql = wsql + " Delete From RmRtHist WITH (TABLOCKX) where RhYyyyMm= " + CStr(adc("wDelYyyyMm")) + " " + _
         " Delete From LabRtHist WITH (TABLOCKX) where LhYyyyMm= " + CStr(adc("wDelYyyyMm")) + " "
                
  moCn.Execute (wsql)
  MWLib.EndProcess Me
  DispMsg "Completed Deleting Rates From History For " + CStr(adc("wDelYyyyMm")), etInfo
  Call CmdDelRt_Click
End Sub
Private Sub DispFra(ByVal pv_CpyRtFra As en_CpyRtFra)
  '*** This is the common procedure which is used to toggle between hiding and displaying
      'the frames (i.e. if the frame is visible then make it invisible and viseversa)
      'associated with a command button
  Select Case pv_CpyRtFra
  Case Is = Cpy
    '*** If Frame FraCpRt is visible then make it invisible and enable all the command buttons
    '*** If Frame FraCpRt is invisible then make it visible and disable all the
        'command buttons except the one which is clicked. Change the background colour of
        'the clicked command button to white and make the font bold
    If FraCpRt.Visible = True Then
      FraCpRt.Visible = False
      FraCpRt.Enabled = False
      Call EnaDisaCmds(False, CmdCpRt)
      CmdCpRt.SetFocus
    Else
      FraCpRt.Visible = True
      FraCpRt.Enabled = True
      adc("wCpYyyyMm").SetFocus
      FraCpRt.ZOrder
      Call EnaDisaCmds(True, CmdCpRt)
    End If
  Case Is = Del
    If FraDelRt.Visible = True Then
      FraDelRt.Visible = False
      FraDelRt.Enabled = False
      Call EnaDisaCmds(False, CmdDelRt)
      CmdDelRt.SetFocus
    Else
      FraDelRt.Visible = True
      FraDelRt.Enabled = True
      adc("wDelYyyyMm").SetFocus
      FraDelRt.ZOrder
      Call EnaDisaCmds(True, CmdDelRt)
    End If
  End Select
End Sub
Private Sub EnaDisaCmds(ByVal pv_ShowFra As Boolean, Optional ByVal pv_CmdBut As MWCTL_BTN1)
  '*** If second parameter is not given and pv_ShowFra = True then the frame FraNKeyAll
      'on Frame Fra_nkeys is disabled and allowadd and allowdelete options of adc are
      'set to false

  '*** If second parameter is not given and pv_ShowFra = False then the frame FraNKeyAll
      'on Frame Fra_nkeys is enabled and allowadd and allowdelete options of adc are
      'set to true

  '*** If second parameter is given and pv_ShowFra = True then the frame FraNKeyAll
      'on Frame Fra_nkeys is disabled and allowadd and allowdelete options of adc are
      'set to false and for the command button specified in the second parameter,
      'the BackColor is changed to white and the font is made Bold

  '*** If second parameter is given and pv_ShowFra = False then the frame FraNKeyAll
      'on Frame Fra_nkeys is enabled and allowadd and allowdelete options of adc are
      'set to true and for the command button specified in the second parameter,
      'the BackColor is changed to normal and the font is made Regular


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
  
  CmdCpRt.Enabled = Not pv_ShowFra
  CmdDelRt.Enabled = Not pv_ShowFra
  If Not pv_CmdBut Is Nothing Then pv_CmdBut.Enabled = True
End Sub
Private Sub HideAllFras()
  '*** This procedure hides all the frames
  FraCpRt.Visible = False
  FraCpRt.Enabled = False
  FraDelRt.Visible = False
  FraDelRt.Enabled = False
End Sub


'*** Std Code not to be changed
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
  ws_BtnStr = "CmdCpGo,CmdDelGo"
  Call SetObjColors(Me, ws_BtnStr)
  '*** Jenny Colour
End Sub
Private Sub ADC_Load()
  Set moCn = adc.Connection  '*** (Bef speed) bef 24/11/06
  Call HideAllFras
  Call EnaDisaCmds(False)
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

'*********** Note ******************


