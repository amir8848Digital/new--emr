VERSION 5.00
Object = "{AB2CC32A-E652-42A2-B341-26ECE02F6386}#2.1#0"; "MwfCtl.ocx"
Begin VB.Form EmrFrmDtTbl 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Date Table"
   ClientHeight    =   7935
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6120
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MDIChild        =   -1  'True
   ScaleHeight     =   7935
   ScaleWidth      =   6120
   Begin MwfCtl.MWCTL_HLP MWCTL_HLP1 
      Height          =   435
      Left            =   2910
      TabIndex        =   6
      TabStop         =   0   'False
      Top             =   7440
      Visible         =   0   'False
      Width           =   435
      _ExtentX        =   767
      _ExtentY        =   767
   End
   Begin MwfCtl.MWCTL_ADC ADC 
      CausesValidation=   0   'False
      Height          =   555
      Left            =   225
      TabIndex        =   7
      Top             =   7350
      Width           =   5715
      _ExtentX        =   10081
      _ExtentY        =   979
      AllowAdd        =   0   'False
      AllowDelete     =   0   'False
      BtnWidth        =   885
      BtnHeight       =   435
      BtnMaskColor    =   12648447
      FontSz          =   8.25
   End
   Begin VB.Frame FRA_KEYS 
      BorderStyle     =   0  'None
      Height          =   690
      Left            =   -36
      TabIndex        =   9
      Top             =   24
      Width           =   5940
      Begin MwfCtl.MWCTL_MED ATXT 
         Height          =   285
         Index           =   7
         Left            =   765
         TabIndex        =   0
         ToolTipText     =   "Enter Date Tag"
         Top             =   0
         Width           =   540
         _ExtentX        =   953
         _ExtentY        =   503
         MaxLength       =   3
         DataType        =   4
         ReCalcParent    =   "WCURRCD2"
         IdName          =   "WDTTAG"
         CmpStr          =   "DtTag="
      End
      Begin MwfCtl.MWCTL_MED ATXT 
         Height          =   285
         Index           =   28
         Left            =   2745
         TabIndex        =   1
         ToolTipText     =   "Enter Currency Code 1"
         Top             =   0
         Width           =   900
         _ExtentX        =   1588
         _ExtentY        =   503
         MaxLength       =   5
         DataType        =   4
         ReCalcOn        =   "WDTTAG"
         IdName          =   "WCURRCD1"
      End
      Begin MwfCtl.MWCTL_MED ATXT 
         Height          =   285
         Index           =   29
         Left            =   5025
         TabIndex        =   2
         ToolTipText     =   "Enter Currency Code 2"
         Top             =   0
         Width           =   900
         _ExtentX        =   1588
         _ExtentY        =   503
         MaxLength       =   5
         DataType        =   4
         ReCalcOn        =   "WDTTAG"
         IdName          =   "WCURRCD2"
      End
      Begin MwfCtl.MWCTL_MED ATXT 
         Height          =   285
         Index           =   0
         Left            =   2745
         TabIndex        =   3
         ToolTipText     =   "Enter From Date"
         Top             =   285
         Width           =   900
         _ExtentX        =   1588
         _ExtentY        =   503
         MaxLength       =   8
         DataType        =   3
         IdName          =   "WFRDT"
         CmpStr          =   "DtDt >="
      End
      Begin MwfCtl.MWCTL_MED ATXT 
         Height          =   285
         Index           =   1
         Left            =   5025
         TabIndex        =   4
         ToolTipText     =   "Enter To Date"
         Top             =   285
         Width           =   900
         _ExtentX        =   1588
         _ExtentY        =   503
         MaxLength       =   8
         DataType        =   3
         IdName          =   "WTODT"
         CmpStr          =   "DtDt <="
      End
      Begin VB.Label ALBL 
         BackStyle       =   0  'Transparent
         Caption         =   "To Dt"
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
         Left            =   3810
         TabIndex        =   15
         Top             =   315
         Width           =   1125
      End
      Begin VB.Label ALBL 
         BackStyle       =   0  'Transparent
         Caption         =   "From Dt"
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
         Left            =   1560
         TabIndex        =   14
         Top             =   315
         Width           =   1125
      End
      Begin VB.Label ALBL 
         BackStyle       =   0  'Transparent
         Caption         =   "Currency 2"
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
         Index           =   19
         Left            =   3810
         TabIndex        =   13
         Top             =   0
         Width           =   1155
      End
      Begin VB.Label ALBL 
         BackStyle       =   0  'Transparent
         Caption         =   "Currency 1"
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
         Left            =   1560
         TabIndex        =   12
         Top             =   0
         Width           =   1125
      End
      Begin VB.Label ALBL 
         BackStyle       =   0  'Transparent
         Caption         =   "Tag"
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
         Index           =   7
         Left            =   270
         TabIndex        =   11
         Top             =   0
         Width           =   795
      End
   End
   Begin VB.Frame FRA_NKEYS 
      BorderStyle     =   0  'None
      Enabled         =   0   'False
      Height          =   7440
      Left            =   -120
      TabIndex        =   8
      Top             =   0
      Width           =   6255
      Begin VB.Frame FraNKeyAll 
         BorderStyle     =   0  'None
         Height          =   7305
         Left            =   150
         TabIndex        =   10
         Top             =   90
         Width           =   5910
         Begin MwfCtl.MWCTL_FLX GrdDt 
            Height          =   6540
            Left            =   210
            TabIndex        =   5
            Top             =   630
            Width           =   5700
            _ExtentX        =   10054
            _ExtentY        =   11536
            Cols            =   9
            colname1        =   "DTDT"
            heading1        =   "Date        "
            datafld1        =   "DtDt"
            datatype1       =   3
            recalcparent1   =   "GrdDt(DTDAY)"
            maxlength1      =   8
            tooltiptext1    =   "Enter Date"
            colname2        =   "DTCOCD"
            heading2        =   "CoCd "
            datafld2        =   "DtCoCd"
            datatype2       =   4
            maxlength2      =   3
            style2          =   0
            tooltiptext2    =   "Enter Company Code"
            colname3        =   "DTHOLYN"
            heading3        =   "Hol Y/N"
            datafld3        =   "DtHolYN"
            datatype3       =   4
            maxlength3      =   1
            colname4        =   "DTDAY"
            heading4        =   "Day                     "
            datafld4        =   "DtDay"
            recalcon4       =   "GrdDt(DTDT)"
            maxlength4      =   10
            style4          =   0
            colname5        =   "DTDESC"
            heading5        =   $"EmrFrmDtTbl.frx":0000
            datafld5        =   "DtDesc"
            maxlength5      =   120
            style5          =   0
            colname6        =   "DTCURCD1"
            heading6        =   "Fr Currency"
            datafld6        =   "DtCurCd1"
            datatype6       =   4
            maxlength6      =   5
            tooltiptext6    =   "Enter Currency 1"
            colname7        =   "DTVAL"
            heading7        =   "Value     "
            datafld7        =   "DtVal"
            datatype7       =   2
            mask7           =   "###0.00"
            maxlength7      =   7
            tooltiptext7    =   "Enter Value"
            colname8        =   "DTCURCD2"
            heading8        =   "To Currency"
            datafld8        =   "DtCurCd2"
            datatype8       =   4
            maxlength8      =   5
            tooltiptext8    =   "Enter Currency 2"
         End
      End
   End
End
Attribute VB_Name = "EmrFrmDtTbl"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
  '---- General Comments on the Entry --------'
  'Mana
  'This entry is used to enter the datewise London Metal Exchange rates for Gold & Platinum.
  'This entry is also used to enter the daily conversion Rates of Dollar to Rupee
  '
  ' X-'', A-DtTbl
  '
  '--------------------

  Option Explicit
  Dim moCn As MwfLib.MDOConnection   '*** (Bef speed) bef 24/11/06
  Private Sub ADC_Load()
    '*** Set the First Key Control property to the working field wDtTag
    '*** Set the First Non Key Control property to the grid GrdDt
    '*** Set the Child Property Of ADC to GrdDt
    '*** Set the Previous and Next Control Properties for the Grid GrdDt
    '*** Set the Hot Key for the Grid GrdDt as 'A'

    Set moCn = adc.Connection    '*** (Bef speed) bef 24/11/06
    Set adc.FirKeyCtl = adc("wDtTag")
    Set adc.FirNKeyCtl = GrdDt
    adc.Child = "GrdDt"
    Set GrdDt.PrevCtl = adc
    Set GrdDt.NextCtl = adc
    GrdDt.HotKey = "A"
  End Sub
  Private Sub ADC_KeyWhen()
    '*** Set the frame FraNKeyAll as the top most frame
    FraNKeyAll.ZOrder (0)
  End Sub
  Private Sub ADC_KeyValid(pr_Cancel As Boolean, pr_ErrMsg As String)
    '*** Check whether the Date Tag is valid
    Call atxt_FldChange(0, UCase("wDtTag"), pr_Cancel, pr_ErrMsg, adc("wDtTag"))
    If pr_Cancel = True Then Exit Sub
  End Sub
  Private Sub ADC_NonKeyWhen(pr_Cancel As Boolean, pr_ErrMsg As String)
   '*** Day and Description column makes visible, when wDtTag is 'HOL'
      
    If adc("wDtTag") = "HOL" Then
      GrdDt.ColProp("DtHolYN").Style = fgcNormal
      GrdDt.ColProp("DtVal").Style = fgcinvisible
      GrdDt.ColProp("DtDay").Style = fgcReadOnly
      GrdDt.ColProp("DtDesc").Style = fgcNormal
      GrdDt.ColProp("DtCoCd").Style = fgcNormal   '53 company code will visible only to HOL
    Else
      GrdDt.ColProp("DtHolYN").Style = fgcinvisible
      GrdDt.ColProp("DtVal").Style = fgcNormal
      GrdDt.ColProp("DtDay").Style = fgcinvisible
      GrdDt.ColProp("DtDesc").Style = fgcinvisible
      GrdDt.ColProp("DtCoCd").Style = fgcinvisible  '53 company code will visible only to HOL
    End If
  End Sub
  Private Sub ATXT_FldWhen(Index As Integer, ByVal IdName As String, Cancel As Boolean, ErrMsg As String)
    '*** Set Helps For Each Field
    Select Case IdName
    Case Is = UCase("wDtTag")
      Call HlpList.PMCd("DTTAG")
    ' Zubin 213
    '***** Manali 3.03 - 18/06/08 - 'LML' included
    Case Is = UCase("wCurrCd1")
      If UCase(adc("wDtTag")) <> "LMG" And UCase(adc("wDtTag")) <> "LMP" And UCase(adc("wDtTag")) <> "LMS" And _
        UCase(adc("wDtTag")) <> "LML" And UCase(adc("wDtTag")) <> "CNV" Then Cancel = True: Exit Sub
      Call HlpList.PMCd("CURNCY")
    Case Is = UCase("wCurrCd2")
      If UCase(adc("wDtTag")) <> "CNV" Then Cancel = True: Exit Sub
      Call HlpList.PMCd("CURNCY")
    ' Zubin 213
    End Select
  End Sub
  Private Sub atxt_FldChange(Index As Integer, ByVal IdName As String, Cancel As Boolean, ErrMsg As String, ByVal pv_NewValue As Variant)
    '*** wDtTag       valid PMCd('DTTAG')
    Select Case IdName
    Case Is = UCase("wDtTag")
      Cancel = Not moCn.RecSeek("Select PMCd From Param where PTyp= 'DTTAG' and " + _
              "PMCd = '" + pv_NewValue + "'")
      If Cancel = True Then ErrMsg = "Invalid Date Tag": Exit Sub
  ' Zubin 213
    Case Is = UCase("wCurrCd1"), UCase("wCurrCd2")
      Cancel = UCase(adc("wDtTag")) = "CNV" And pv_NewValue <> "" And _
      Not moCn.RecSeek("Select 'x' From Param where Ptyp= 'CURNCY' and PMCd= '" + pv_NewValue + "' ")
      If Cancel = True Then ErrMsg = "Invalid Currency Code": Exit Sub
  ' Zubin 213

    End Select
  End Sub

  Private Sub atxt_ReCalculate(Index As Integer, ByVal IdName As String, ByVal OldVal As Variant, ByVal RowNum As Integer, ByVal SourceName As String, ByVal SourceColName As String)
  ' Zubin 213
    Select Case IdName
    '***** Manali 3.03 - 18/06/08 - 'LML' included
    Case Is = UCase("wCurrCd1")
      If UCase(adc("wDtTag")) <> "LMG" And UCase(adc("wDtTag")) <> "LMP" And UCase(adc("wDtTag")) <> "LMS" And _
        UCase(adc("wDtTag")) <> "LML" And UCase(adc("wDtTag")) <> "CNV" Then adc("wCurrCd1") = ""
    Case Is = UCase("wCurrCd2")
      If UCase(adc("wDtTag")) <> "CNV" Then adc("wCurrCd2") = ""
    End Select
  ' Zubin 213
  End Sub

  Private Sub GrdDt_ColWhen(ByVal RowNum As Integer, ByVal ColName As String, Cancel As Boolean, ErrMsg As String, ByVal Action As MwfCtl.en_fgAction)
    Select Case ColName
      Case Is = UCase("DtHolYN")
        Call HlpList.PMCd("YN")
  ' Zubin 213
      Case Is = UCase("DtCurCd1"), UCase("DtCurCd2")
        Call HlpList.PMCd("CURNCY")
  ' Zubin 213
  '53 calling help
      Case Is = UCase("DtCoCd")
        Call HlpList.hCoCd
    End Select
  End Sub
  Private Sub GrdDt_ColChange(ByVal RowNum As Integer, ByVal ColName As String, ByVal NewValue As Variant, Cancel As Boolean, ErrMsg As String)
    '*** DtDt           Not Blank
    '*** DtVal          <> 0

    With GrdDt
      Select Case ColName
      Case Is = UCase("DtDt")
        If NewValue = "" Then Cancel = True: ErrMsg = "Enter Date": Exit Sub
      Case Is = UCase("DtVal")
        If NewValue = 0 Then Cancel = True: ErrMsg = "Enter Value": Exit Sub
      Case Is = UCase("DtHolYN")
        Cancel = NewValue <> "" And _
           (Not moCn.RecSeek("Select PMCd From Param where PTyp= 'YN' and " + _
          "PMCd = '" + NewValue + "'"))
        If Cancel = True Then ErrMsg = "Invalid Holiday Option": Exit Sub
  ' Zubin 213
      Case Is = UCase("DtCurCd1"), UCase("DtCurCd2")
        Cancel = Not moCn.RecSeek("Select PMCd From Param where PTyp= 'CURNCY' and " + _
          "PMCd = '" + NewValue + "' and PSCd= ''")
        If Cancel = True Then ErrMsg = "Invalid Currency Code": Exit Sub
  ' Zubin 213
  '53 validating company code
      Case Is = UCase("DtCoCd")
        Cancel = Not moCn.RecSeek("Select 'x' From Head Where HCd='" + ctSelfCmCd + "' And HCoCd='" + NewValue + "'")
        If Cancel = True Then ErrMsg = "Invalid Company Code": Exit Sub
     
      End Select
    End With
  End Sub
  Private Sub GrdDt_InitKey(ByVal RowNum As Integer)
    '*** Initialise Key Fields
        'The fields of the Base table which are not existing in the grid are stored in the
        'table using Grid_InitKey

    With GrdDt
      If .IsNew(RowNum) Then _
        .Store "DtTag", adc("wDtTag")
    End With
  End Sub
  Private Sub ADC_SetRecSource()
    '*** Since this is an 'A' Entry, there is no RecSource associated with ADC. So set the
        'DataMode property of ADC control as xdatUnbound
    adc.DataMode = xdatUnbound

    ' Zubin 213
    Call GrdPropSet
  End Sub
  Private Sub GrdDt_ReCalc(ByVal RowNum As Integer, ByVal ColName As String, ByVal OldVal As Variant, ByVal SourceName As String, ByVal SourceColName As String)
  '*** Find the value of the day field for the user given date
    If adc("wDtTag") = "HOL" Then
      With GrdDt
        Select Case ColName
          Case Is = UCase("DtDay")
            If (.Mode = fgmnorm) Then _
              .Value(RowNum, "DtDay") = WeekdayName(DatePart("w", .Value(RowNum, "DtDt")), False, vbSunday)
        End Select
      End With
    End If
  End Sub
  Private Sub GrdDt_SetRecSource()
    '*** Set the Record Source of the Grid GrdDt
        'The sort option DtDt

    Dim ws_Cnd As String, ws_OrdBy As String

    ' Zubin 213
    Dim ws_CurrCdFr As String, ws_CurrCdTo As String, ws_CurCnd As String
    
    '***** Manali 3.03 - 18/06/08 - 'LML' included
    If UCase(adc("wDtTag")) = "CNV" Then
      ws_CurrCdFr = IIF(adc("wCurrCd1") <> "", " (DtCurCd1 = '" + adc("wCurrCd1") + "' Or DtCurCd2 = '" + adc("wCurrCd1") + "') ", "")
      ws_CurrCdTo = IIF(adc("wCurrCd2") <> "", " (DtCurCd1 = '" + adc("wCurrCd2") + "' Or DtCurCd2 = '" + adc("wCurrCd2") + "') ", "")
      adc("wCurrCd1").CmpStr = ""
    ElseIf UCase(adc("wDtTag")) = "LMS" Or UCase(adc("wDtTag")) = "LMP" Or UCase(adc("wDtTag")) = "LMG" Or UCase(adc("wDtTag")) = "LML" Then
      adc("wCurrCd1").CmpStr = "DtCurCd1= "
    Else
      adc("wCurrCd1").CmpStr = ""
    End If

    ws_CurCnd = ws_CurrCdFr + IIF(adc("wCurrCd1") <> "", IIF(adc("wCurrCd2") <> "", " and ", ""), "") + ws_CurrCdTo
    ' Zubin 213

    ws_Cnd = adc.RepCond
    ws_Cnd = IIF(ws_Cnd <> "", " where " + ws_Cnd, "")
    ws_Cnd = ws_Cnd + IIF(ws_CurCnd <> "", " and " + ws_CurCnd, "") ' Zubin 213

    ws_OrdBy = " Order By DtDt"
    GrdDt.RecSource = " Select * from DtTbl " + ws_Cnd + ws_OrdBy
  End Sub
  Private Sub GrdPropSet()
    ' Zubin 213
    '***** Manali 3.03 - 18/06/08 - 'LML' included
    With GrdDt
      If UCase(adc("wDtTag")) = "LMG" Or UCase(adc("wDtTag")) = "LMP" Or UCase(adc("wDtTag")) = "LMS" Or _
        UCase(adc("wDtTag")) = "LML" Or UCase(adc("wDtTag")) = "USD" Then
          If UCase(adc("wDtTag")) <> "USD" Then .ColProp("DtCurCd1").Style = fgcNormal Else .ColProp("DtCurCd1").Style = fgcinvisible
          .ColProp("DtCurCd2").Style = fgcinvisible
          .ColProp("DtVal").Heading = "Value" + IIF((7 * 2.5) - (Len("Value")) > 0, Space(Abs((7 * 2.5) - (Len("Value")))), "")
          .ColProp("DtVal").Mask = "###0.00": .ColProp("DtVal").MaxLength = 7
          .ColProp("DtCurCd1").Heading = "Currency" + IIF((7 * 2.5) - (Len("Currency")) > 0, Space(Abs((7 * 2.5) - (Len("Currency")))), "")
      ElseIf UCase(adc("wDtTag")) = "CNV" Then
          .ColProp("DtCurCd1").Style = fgcNormal: .ColProp("DtCurCd2").Style = fgcNormal
          .ColProp("DtVal").Heading = "Conv Rate" + IIF((7 * 2.5) - (Len("Conv Rate")) > 0, Space(Abs((7 * 2.5) - (Len("Conv Rate")))), "")
          .ColProp("DtVal").Mask = "###0.0000": .ColProp("DtVal").MaxLength = 9
          .ColProp("DtCurCd1").Heading = "Fr Currency" + IIF((7 * 2.5) - (Len("Fr Currency")) > 0, Space(Abs((7 * 2.5) - (Len("Fr Currency")))), "")
      Else
          .ColProp("DtCurCd1").Style = fgcinvisible: .ColProp("DtCurCd2").Style = fgcinvisible
      End If
      GrdDt.SetColWidthToHd
    End With
    ' Zubin 213
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
    Cancel = atxt(Index).Validate
  End Sub
  Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
      If UnloadMode <> vbFormCode Then Cancel = True: Exit Sub

      Set moCn = Nothing    '*** (09/08/05)    '*** (Bef speed) bef 24/11/06
  End Sub
  Private Sub GrdDt_Validate(Cancel As Boolean)
    Cancel = GrdDt.Validate
  End Sub


