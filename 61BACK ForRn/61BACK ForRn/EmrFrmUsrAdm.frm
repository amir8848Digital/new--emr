VERSION 5.00
Object = "{AB2CC32A-E652-42A2-B341-26ECE02F6386}#2.1#0"; "MwfCtlForSpeed61.ocx"
Begin VB.Form EmrFrmUsrAdm 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "User Administration"
   ClientHeight    =   9810
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   11865
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   9810
   ScaleMode       =   0  'User
   ScaleWidth      =   11865
   WindowState     =   2  'Maximized
   Begin MwfCtl.MWCTL_HLP MWCTL_HLP1 
      Height          =   405
      Left            =   3990
      TabIndex        =   1
      TabStop         =   0   'False
      Top             =   9270
      Visible         =   0   'False
      Width           =   555
      _ExtentX        =   979
      _ExtentY        =   714
   End
   Begin MwfCtl.MWCTL_ADC ADC 
      CausesValidation=   0   'False
      Height          =   555
      Left            =   90
      TabIndex        =   2
      Top             =   9300
      Width           =   11640
      _ExtentX        =   20532
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
      Height          =   9540
      Left            =   -120
      TabIndex        =   3
      Top             =   0
      Width           =   11985
      Begin VB.Frame FraNKeyAll 
         BorderStyle     =   0  'None
         Height          =   9165
         Left            =   150
         TabIndex        =   5
         Top             =   90
         Width           =   11790
         Begin MwfCtl.MWCTL_FLX GrdUsrLog 
            Height          =   9120
            Left            =   90
            TabIndex        =   0
            Top             =   0
            Width           =   11655
            _ExtentX        =   20558
            _ExtentY        =   16087
            Cols            =   8
            AllowAdd        =   0   'False
            colname1        =   "ULUSRCD"
            heading1        =   "Usr Cd      "
            datafld1        =   "UlUsrCd"
            datatype1       =   4
            recalcparent1   =   "GrdUsrLog(WUSRNM)"
            maxlength1      =   5
            style1          =   2
            tooltiptext1    =   "Enter User Code"
            colname2        =   "WUSRNM"
            heading2        =   "User Name                            "
            recalcon2       =   "GrdUsrLog(ULUSRCD)"
            maxlength2      =   30
            style2          =   2
            tooltiptext2    =   "User Name"
            colname3        =   "ULCOCD"
            heading3        =   "Co Cd"
            datafld3        =   "UlCoCd"
            datatype3       =   4
            recalcparent3   =   "GrdUsrLog(WCONM)"
            maxlength3      =   3
            style3          =   2
            tooltiptext3    =   "Company Code"
            colname4        =   "WCONM"
            heading4        =   "Company Name                              "
            recalcon4       =   "GrdUsrLog(ULCOCD)"
            maxlength4      =   30
            style4          =   2
            tooltiptext4    =   "Company Name"
            colname5        =   "ULHOSTNM"
            heading5        =   "Host Name                          "
            datafld5        =   "UlHostNm"
            maxlength5      =   30
            style5          =   2
            tooltiptext5    =   "Host Name"
            colname6        =   "ULDT"
            heading6        =   "Date            "
            datafld6        =   "UlDt"
            datatype6       =   3
            maxlength6      =   8
            tooltiptext6    =   "Login Date"
            colname7        =   "ULTIME"
            heading7        =   "Time     "
            datafld7        =   "UlTime"
            datatype7       =   2
            mask7           =   "##0.00"
            maxlength7      =   6
            tooltiptext7    =   "Login Type"
         End
      End
      Begin MwfCtl.MWCTL_MED ATXT 
         Height          =   285
         Index           =   0
         Left            =   9120
         TabIndex        =   6
         ToolTipText     =   "Number of Users Logged in Emperor"
         Top             =   9240
         Width           =   885
         _ExtentX        =   1561
         _ExtentY        =   503
         Enabled         =   0   'False
         Mask            =   "###0"
         MaxLength       =   4
         DataType        =   1
         IdName          =   "WUSRCOUNT"
      End
      Begin VB.Label ALBL 
         Caption         =   "Total Users :"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   0
         Left            =   7680
         TabIndex        =   7
         Top             =   9255
         Width           =   1455
      End
   End
   Begin VB.Frame FRA_KEYS 
      BorderStyle     =   0  'None
      Height          =   180
      Left            =   -36
      TabIndex        =   4
      Top             =   24
      Width           =   8190
   End
End
Attribute VB_Name = "EmrFrmUsrAdm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'--- General comments on the entry ----
'
'Master Entry for entering the various Banks, their addresses and the corresponding bank
'   instructions
'
'---

Option Explicit
Dim moCn As MwfLib.MDOConnection   '*** (Bef speed) bef 24/11/06

Private Sub ADC_Load()
  '*** Set the First Key Control property to adc
  '*** Set the First Non Key Control property to the grid GrdUsrLog
  '*** Set the Child Property Of ADC to GrdUsrLog
  '*** Set the Previous and Next Control Properties for the Grid GrdUsrLog
  '*** Set the Hot Key for the Grid GrdUsrLog as 'A'

  Set moCn = ADC.Connection    '*** (Bef speed) bef 24/11/06

  Set ADC.FirKeyCtl = ADC
  Set ADC.FirNKeyCtl = GrdUsrLog
  ADC.Child = "GrdUsrLog"
  Set GrdUsrLog.PrevCtl = ADC
  Set GrdUsrLog.NextCtl = ADC
  GrdUsrLog.HotKey = "A"
End Sub
Private Sub ADC_NonKeyWhen(pr_Cancel As Boolean, pr_ErrMsg As String)
    ADC("WUSRCOUNT") = moCn.GetFldVal("Select Count('x') from UsrLogin")    'Sachin - 4.1.0.0
End Sub
Private Sub ADC_SetRecSource()
  '*** Since this is an 'A' Entry, there is no RecSource associated with ADC. So set the
      'DataMode property of ADC control as xdatUnbound
  ADC.DataMode = xdatUnbound
End Sub
'
'Private Sub GrdUsrLog_OnDelete(ByVal RowNum As Integer, Cancel As Boolean, ErrMsg As String)
'  If UCase$(GrdUsrLog.Value(RowNum, "UlCoCd")) <> UCase$(gs_CoCd) Then
'    Cancel = True
'    ErrMsg = "You Cannot Delete this Record, User is in Different Company."
'  End If
'End Sub

Private Sub GrdUsrLog_OnDelete(ByVal RowNum As Integer, Cancel As Boolean, ErrMsg As String)
    'Reducing a user on delete
    ADC("WUSRCOUNT") = ADC("WUSRCOUNT") - 1
End Sub

Private Sub GrdUsrLog_ReCalc(ByVal RowNum As Integer, ByVal ColName As String, ByVal OldVal As Variant, ByVal SourceName As String, ByVal SourceColName As String)
  ' set then User & Company Name
  If ADC.Mode = xInit Or ADC.Mode = xNorm Then
    Select Case UCase$(Trim$(ColName))
      Case UCase$(Trim$("wUsrNm"))
        GrdUsrLog.Value(RowNum, "wUsrNm") = moCn.GetFldVal("Select PDesc From Param Where PTyp = 'USR' And PmCd = '" + GrdUsrLog.Value(RowNum, "UlUsrCd") + "'")
      Case UCase$(Trim$("wCoNm"))
        GrdUsrLog.Value(RowNum, "wCoNm") = moCn.GetFldVal("Select HName From Head Where HCoCd = '" + GrdUsrLog.Value(RowNum, "UlCoCd") + "'")
    End Select
  End If
End Sub

Private Sub GrdUsrLog_SetRecSource()
  '*** Set the Record Source of the Grid GrdUsrLog
      'Default sort option is BkCd
  '*** Jay 3.1.1 [UlCoCd Added]
  GrdUsrLog.RecSource = "Select * from UsrLogin Order By UlUsrCd, UlCoCd"
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

' ########################  Manoj #### Ver: 2.0.6 #### Date: 17/06/2004  ########################
' These Lines Commented b'coz no AtXt Field is required on this Fied.
'Private Sub ATXT_Validate(Index As Integer, Cancel As Boolean)
'  Cancel = ATXT(Index).Validate
'End Sub
' ###########################################  Manoj  ###########################################
Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    If UnloadMode <> vbFormCode Then Cancel = True: Exit Sub
    
    Set moCn = Nothing    '*** (09/08/05)    '*** (Bef speed) bef 24/11/06
End Sub
Private Sub GrdUsrLog_Validate(Cancel As Boolean)
  Cancel = GrdUsrLog.Validate
End Sub

