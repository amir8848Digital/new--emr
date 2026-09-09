VERSION 5.00
Object = "{6A2CE895-C1AC-48AF-B558-FB94BBE12499}#1.0#0"; "mwfctl.ocx"
Begin VB.Form EmrFrmBnkAg 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Bank/Agent  Master"
   ClientHeight    =   5505
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   8955
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   5505
   ScaleMode       =   0  'User
   ScaleWidth      =   8955
   WindowState     =   2  'Maximized
   Begin MwfCtl.MWCTL_HLP MWCTL_HLP1 
      Height          =   405
      Left            =   4095
      TabIndex        =   7
      TabStop         =   0   'False
      Top             =   4890
      Visible         =   0   'False
      Width           =   555
      _ExtentX        =   979
      _ExtentY        =   714
   End
   Begin MwfCtl.MWCTL_ADC ADC 
      CausesValidation=   0   'False
      Height          =   555
      Left            =   120
      TabIndex        =   8
      Top             =   4920
      Width           =   8565
      _ExtentX        =   15108
      _ExtentY        =   979
      BtnWidth        =   885
      BtnHeight       =   435
      BtnMaskColor    =   8454143
      FontSz          =   8.25
   End
   Begin VB.Frame FRA_KEYS 
      BorderStyle     =   0  'None
      Height          =   675
      Left            =   -36
      TabIndex        =   10
      Top             =   120
      Width           =   2685
      Begin MwfCtl.MWCTL_MED ATXT 
         Height          =   285
         Index           =   1
         Left            =   1740
         TabIndex        =   1
         ToolTipText     =   "Enter Bank/Agent Code"
         Top             =   380
         Width           =   855
         _ExtentX        =   1508
         _ExtentY        =   503
         MaxLength       =   5
         DataType        =   4
         DataField       =   "BkCd"
         IdName          =   "BKCD"
      End
      Begin MwfCtl.MWCTL_MED ATXT 
         Height          =   285
         Index           =   5
         Left            =   1740
         TabIndex        =   0
         ToolTipText     =   "Enter Bank/Agent Category"
         Top             =   90
         Width           =   435
         _ExtentX        =   767
         _ExtentY        =   503
         MaxLength       =   2
         DataType        =   4
         DataField       =   "BkCtg"
         IdName          =   "BKCTG"
      End
      Begin VB.Label ALBL 
         BackStyle       =   0  'Transparent
         Caption         =   "Bank/Agnet Ctg"
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
         Left            =   60
         TabIndex        =   15
         Top             =   90
         Width           =   1635
      End
      Begin VB.Label ALBL 
         BackStyle       =   0  'Transparent
         Caption         =   "Bank/Agent Cd"
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
         Left            =   60
         TabIndex        =   11
         Top             =   375
         Width           =   1575
      End
   End
   Begin VB.Frame FRA_NKEYS 
      BorderStyle     =   0  'None
      Enabled         =   0   'False
      Height          =   5595
      Left            =   -120
      TabIndex        =   9
      Top             =   0
      Width           =   8970
      Begin VB.Frame FraNKeyAll 
         BorderStyle     =   0  'None
         Height          =   4725
         Left            =   150
         TabIndex        =   2
         Top             =   175
         Width           =   8670
         Begin MwfCtl.MWCTL_MED ATXT 
            Height          =   285
            Index           =   0
            Left            =   1680
            TabIndex        =   3
            ToolTipText     =   "Enter Bank/Agent Address"
            Top             =   900
            Width           =   6885
            _ExtentX        =   12144
            _ExtentY        =   503
            MaxLength       =   45
            DataField       =   "BkAdd1"
            IdName          =   "BKADD1"
         End
         Begin MwfCtl.MWCTL_MED ATXT 
            Height          =   285
            Index           =   2
            Left            =   1680
            TabIndex        =   4
            ToolTipText     =   "Enter Bank/Agent Address"
            Top             =   1185
            Width           =   6885
            _ExtentX        =   12144
            _ExtentY        =   503
            MaxLength       =   45
            DataField       =   "BkAdd2"
            IdName          =   "BKADD2"
         End
         Begin MwfCtl.MWCTL_MED ATXT 
            Height          =   285
            Index           =   4
            Left            =   1680
            TabIndex        =   5
            ToolTipText     =   "Enter Bank/Agent Address"
            Top             =   1470
            Width           =   6885
            _ExtentX        =   12144
            _ExtentY        =   503
            MaxLength       =   45
            DataField       =   "BkAdd3"
            IdName          =   "BKADD3"
         End
         Begin MwfCtl.MWCTL_MED ATXT 
            Height          =   2775
            Index           =   55
            Left            =   1680
            TabIndex        =   6
            ToolTipText     =   "Enter Bank/Agent Instructions"
            Top             =   1755
            Width           =   6885
            _ExtentX        =   12144
            _ExtentY        =   4895
            Alignment       =   3
            MaxLength       =   0
            DataField       =   "BkInstr"
            IdName          =   "BKINSTR"
         End
         Begin MwfCtl.MWCTL_MED ATXT 
            Height          =   285
            Index           =   3
            Left            =   1680
            TabIndex        =   17
            ToolTipText     =   "Enter Bank/Agent Name"
            Top             =   610
            Width           =   4725
            _ExtentX        =   8334
            _ExtentY        =   503
            MaxLength       =   30
            DataField       =   "BkName"
            IdName          =   "BKNAME"
         End
         Begin MwfCtl.MWCTL_MED ATXT 
            Height          =   285
            Index           =   21
            Left            =   3300
            TabIndex        =   18
            ToolTipText     =   "Enter Bank Account"
            Top             =   320
            Width           =   3105
            _ExtentX        =   5477
            _ExtentY        =   503
            MaxLength       =   20
            DataType        =   4
            DataField       =   "BkAc"
            IdName          =   "BKAC"
         End
         Begin VB.Label ALBL 
            BackStyle       =   0  'Transparent
            Caption         =   "A/c"
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
            Index           =   18
            Left            =   2760
            TabIndex        =   16
            Top             =   320
            Width           =   555
         End
         Begin VB.Label ALBL 
            BackStyle       =   0  'Transparent
            Caption         =   "Instruction"
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
            Index           =   54
            Left            =   0
            TabIndex        =   14
            Top             =   1755
            Width           =   1785
         End
         Begin VB.Label ALBL 
            BackStyle       =   0  'Transparent
            Caption         =   "Address"
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
            Left            =   0
            TabIndex        =   13
            Top             =   900
            Width           =   945
         End
         Begin VB.Label ALBL 
            BackStyle       =   0  'Transparent
            Caption         =   "Name"
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
            Index           =   3
            Left            =   0
            TabIndex        =   12
            Top             =   610
            Width           =   855
         End
      End
   End
End
Attribute VB_Name = "EmrFrmBnkAg"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'----   General Comments on the Entry Form  -------'
'Master Entry for entering the various Banks, their addresses and the corresponding bank
'   instructions
'X-Entry
'One Menu.
'
'----


Option Explicit
Dim moCn As MwfLib.MDOConnection   '*** (Bef speed) bef 24/11/06

Private Sub ADC_Load()
  '*** Set the First Key Control to Bank Category
  '*** Set the First Non Key Control to Bank Name
  '*** Set the Previous and Next Control Properties for ADC
  '*** Set the frame FraNKeyAll as the top most frame

  Set moCn = adc.Connection    '*** (Bef speed) bef 24/11/06
  Set adc.FirKeyCtl = adc("BkCtg")
  Set adc.FirNKeyCtl = adc("BkName")

  adc.Child = ""
  FraNKeyAll.ZOrder (0)
End Sub
Private Sub ADC_KeyWhen()
  '*** Hide all frames except the frame FraNKeyAll
  Call HideAllFras
End Sub
Private Sub ADC_KeyValid(pr_Cancel As Boolean, pr_ErrMsg As String)
    If adc("BkCd") = "" Then
        pr_Cancel = True: pr_ErrMsg = "Bank/Agent Code Can Not Be Left Blank": Exit Sub
    End If
End Sub

Private Sub ATXT_FldWhen(Index As Integer, ByVal IdName As String, Cancel As Boolean, ErrMsg As String)
  '*** Set Helps For Each Field
  Select Case UCase(IdName)
  Case Is = UCase("BkCd")
    Call HlpList.BkCd(adc("BkCtg"))
  Case Is = UCase("BkCtg")
    Call HlpList.PMCd("BKCTG")
  End Select
End Sub
Private Sub atxt_ReCalculate(Index As Integer, ByVal IdName As String, ByVal OldVal As Variant, ByVal RowNum As Integer, ByVal SourceName As String, ByVal SourceColName As String)
  Select Case UCase(IdName)
  End Select
End Sub
Private Sub atxt_FldChange(Index As Integer, ByVal IdName As String, Cancel As Boolean, ErrMsg As String, ByVal pv_NewValue As Variant)
  '*** BkCtg           Not Blank
  Select Case UCase(IdName)
  Case Is = UCase("BkCtg")
    If Trim(pv_NewValue) = "" Then
      Cancel = True: ErrMsg = "Bank/Agent Category Cannot Be Blank": Exit Sub
    ElseIf Trim(pv_NewValue) <> "" Then
        Cancel = pv_NewValue <> "" And (Not moCn.RecSeek("select PMCd from Param where " + _
                         "PTyp= 'BKCTG' and PMCd = '" + pv_NewValue + "'"))
        ErrMsg = "Invalid Bank/Agent Category": Exit Sub
    End If
  End Select
End Sub
Private Sub ADC_SetRecSource()
  '*** Set the RecSource for ADC  (Base Table is Bank)
  adc.RecSource = " Select * from BnkAgMst Where BkCtg = '" + adc("BkCtg") + "' and BkCd = '" + adc("BkCd") + "'"
End Sub
Private Sub HideAllFras()
  '*** This procedure hides all the frames other than FraNKeyAll
  FraNKeyAll.Visible = True
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
  ws_BtnStr = "CmdCpGo,CmdDelGo"
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

