VERSION 5.00
Object = "{AB2CC32A-E652-42A2-B341-26ECE02F6386}#2.0#0"; "MwfCtl.ocx"
Begin VB.Form EmrFrmRfIdMst 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Bag Alias Master"
   ClientHeight    =   9135
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   10365
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   9135
   ScaleWidth      =   10365
   Begin MwfCtl.MWCTL_HLP MWCTL_HLP1 
      Height          =   435
      Left            =   2550
      TabIndex        =   11
      TabStop         =   0   'False
      Top             =   8460
      Visible         =   0   'False
      Width           =   435
      _ExtentX        =   767
      _ExtentY        =   767
   End
   Begin MwfCtl.MWCTL_ADC ADC 
      CausesValidation=   0   'False
      Height          =   555
      Left            =   105
      TabIndex        =   12
      Top             =   8520
      Width           =   10200
      _ExtentX        =   17992
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
      Left            =   0
      TabIndex        =   14
      Top             =   0
      Width           =   10335
      Begin MwfCtl.MWCTL_MED atxt 
         Height          =   285
         Index           =   28
         Left            =   570
         TabIndex        =   0
         ToolTipText     =   "Enter From Company Code"
         Top             =   60
         Width           =   840
         _ExtentX        =   1482
         _ExtentY        =   503
         MaxLength       =   3
         DataType        =   4
         IdName          =   "WRIMBCOCD"
         CmpStr          =   "RimBCoCd ="
      End
      Begin MwfCtl.MWCTL_MED atxt 
         Height          =   285
         Index           =   7
         Left            =   3075
         TabIndex        =   1
         ToolTipText     =   "Enter From Bag Alias No."
         Top             =   60
         Width           =   930
         _ExtentX        =   1640
         _ExtentY        =   503
         Alignment       =   2
         Mask            =   "#####0"
         MaxLength       =   6
         DataType        =   1
         IdName          =   "WRIMNOFR"
         CmpStr          =   "RimNo >="
      End
      Begin MwfCtl.MWCTL_MED atxt 
         Height          =   285
         Index           =   1
         Left            =   3075
         TabIndex        =   2
         ToolTipText     =   "Enter To Bag Alias No."
         Top             =   360
         Width           =   930
         _ExtentX        =   1640
         _ExtentY        =   503
         Alignment       =   2
         Mask            =   "#####0"
         MaxLength       =   6
         DataType        =   1
         IdName          =   "WRIMNOTO"
         CmpStr          =   "RimNo <="
      End
      Begin MwfCtl.MWCTL_MED atxt 
         Height          =   285
         Index           =   0
         Left            =   4950
         TabIndex        =   3
         ToolTipText     =   "Enter From Design Code"
         Top             =   60
         Width           =   2145
         _ExtentX        =   3784
         _ExtentY        =   503
         MaxLength       =   15
         DataType        =   4
         IdName          =   "WRIMDMCDFR"
         CmpStr          =   "RimDmCd >="
      End
      Begin MwfCtl.MWCTL_MED atxt 
         Height          =   285
         Index           =   2
         Left            =   4950
         TabIndex        =   4
         ToolTipText     =   "Enter To Rfid No"
         Top             =   360
         Width           =   2145
         _ExtentX        =   3784
         _ExtentY        =   503
         MaxLength       =   15
         DataType        =   4
         IdName          =   "WRIMDMCDTO"
         CmpStr          =   "RimDmCd <="
      End
      Begin MwfCtl.MWCTL_MED atxt 
         Height          =   285
         Index           =   3
         Left            =   8160
         TabIndex        =   5
         ToolTipText     =   "Enter From Bag Yy"
         Top             =   60
         Width           =   345
         _ExtentX        =   609
         _ExtentY        =   503
         Alignment       =   2
         Mask            =   "####0"
         MaxLength       =   2
         IdName          =   "WRIMBYYFR"
         CmpStr          =   "RimBYy >="
      End
      Begin MwfCtl.MWCTL_MED atxt 
         Height          =   285
         Index           =   4
         Left            =   8160
         TabIndex        =   8
         ToolTipText     =   "Enter From Bag YY"
         Top             =   360
         Width           =   345
         _ExtentX        =   609
         _ExtentY        =   503
         Alignment       =   2
         MaxLength       =   2
         IdName          =   "WRIMBYYTO"
         CmpStr          =   "RimBYy <="
      End
      Begin MwfCtl.MWCTL_MED atxt 
         Height          =   285
         Index           =   5
         Left            =   8600
         TabIndex        =   6
         ToolTipText     =   "Enter From Bag Chr"
         Top             =   60
         Width           =   600
         _ExtentX        =   1058
         _ExtentY        =   503
         Alignment       =   2
         MaxLength       =   3
         DataType        =   4
         IdName          =   "WRIMBCHRFR"
         CmpStr          =   "RimBChr >="
      End
      Begin MwfCtl.MWCTL_MED atxt 
         Height          =   285
         Index           =   6
         Left            =   8600
         TabIndex        =   9
         ToolTipText     =   "Enter From Bag Chr"
         Top             =   360
         Width           =   600
         _ExtentX        =   1058
         _ExtentY        =   503
         Alignment       =   2
         MaxLength       =   3
         DataType        =   4
         IdName          =   "WRIMBCHRTO"
         CmpStr          =   "RimBChr <="
      End
      Begin MwfCtl.MWCTL_MED atxt 
         Height          =   285
         Index           =   8
         Left            =   9300
         TabIndex        =   7
         ToolTipText     =   "Enter From Bag No"
         Top             =   60
         Width           =   950
         _ExtentX        =   1667
         _ExtentY        =   503
         Alignment       =   2
         Mask            =   "######0"
         MaxLength       =   7
         DataType        =   1
         IdName          =   "WRIMBNOFR"
         CmpStr          =   "RimBNo >="
      End
      Begin MwfCtl.MWCTL_MED atxt 
         Height          =   285
         Index           =   9
         Left            =   9300
         TabIndex        =   10
         ToolTipText     =   "Enter From Bag No"
         Top             =   360
         Width           =   950
         _ExtentX        =   1667
         _ExtentY        =   503
         Alignment       =   2
         Mask            =   "######0"
         MaxLength       =   7
         DataType        =   1
         IdName          =   "WRIMBNOTO"
         CmpStr          =   "RimBNo <="
      End
      Begin VB.Label ALBL 
         BackColor       =   &H00C0C0C0&
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
         ForeColor       =   &H00000000&
         Height          =   285
         Index           =   34
         Left            =   9200
         TabIndex        =   27
         Top             =   360
         Width           =   105
      End
      Begin VB.Label ALBL 
         BackColor       =   &H00C0C0C0&
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
         ForeColor       =   &H00000000&
         Height          =   285
         Index           =   35
         Left            =   9200
         TabIndex        =   26
         Top             =   60
         Width           =   105
      End
      Begin VB.Label ALBL 
         BackColor       =   &H00C0C0C0&
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
         ForeColor       =   &H00000000&
         Height          =   285
         Index           =   36
         Left            =   8520
         TabIndex        =   25
         Top             =   360
         Width           =   105
      End
      Begin VB.Label ALBL 
         BackColor       =   &H00C0C0C0&
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
         ForeColor       =   &H00000000&
         Height          =   285
         Index           =   37
         Left            =   8520
         TabIndex        =   24
         Top             =   60
         Width           =   105
      End
      Begin VB.Label ALBL 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "To"
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
         Index           =   5
         Left            =   7170
         TabIndex        =   23
         Top             =   360
         Width           =   945
      End
      Begin VB.Label ALBL 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Bag No. Fr"
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
         Index           =   4
         Left            =   6960
         TabIndex        =   22
         Top             =   60
         Width           =   1155
      End
      Begin VB.Label ALBL 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "DmCd Fr"
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
         Left            =   3960
         TabIndex        =   21
         Top             =   60
         Width           =   945
      End
      Begin VB.Label ALBL 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "To"
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
         Left            =   4170
         TabIndex        =   20
         Top             =   330
         Width           =   765
      End
      Begin VB.Label ALBL 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "To"
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
         Left            =   2280
         TabIndex        =   19
         Top             =   330
         Width           =   765
      End
      Begin VB.Label ALBL 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "CoCd"
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
         TabIndex        =   17
         Top             =   60
         Width           =   525
      End
      Begin VB.Label ALBL 
         BackStyle       =   0  'Transparent
         Caption         =   "Bag Alias No. Fr"
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
         Left            =   1515
         TabIndex        =   16
         Top             =   60
         Width           =   1545
      End
   End
   Begin VB.Frame FRA_NKEYS 
      BorderStyle     =   0  'None
      Enabled         =   0   'False
      Height          =   9150
      Left            =   0
      TabIndex        =   13
      Top             =   0
      Width           =   10425
      Begin VB.Frame FraNKeyAll 
         BorderStyle     =   0  'None
         Height          =   8745
         Left            =   30
         TabIndex        =   15
         Top             =   0
         Width           =   10215
         Begin MwfCtl.MWCTL_FLX GrdRfId 
            Height          =   7680
            Left            =   90
            TabIndex        =   18
            Top             =   720
            Width           =   10125
            _ExtentX        =   17859
            _ExtentY        =   13547
            Cols            =   8
            colname1        =   "RIMNO"
            heading1        =   "Bag Alias No.     "
            datafld1        =   "RimNo"
            datatype1       =   1
            mask1           =   "#####0"
            maxlength1      =   6
            colname2        =   "RIMBCOCD"
            datafld2        =   "RimBCoCd"
            datatype2       =   4
            maxlength2      =   3
            style2          =   0
            colname3        =   "WRIMBAG"
            datatype3       =   4
            maxlength3      =   18
            colname4        =   "RIMBYY"
            heading4        =   "Byy"
            datafld4        =   "RimBYy"
            datatype4       =   4
            recalcparent4   =   "GrdRfId(RIMDMCD)"
            maxlength4      =   2
            colname5        =   "RIMBCHR"
            heading5        =   "BChr"
            datafld5        =   "RimBChr"
            datatype5       =   4
            recalcparent5   =   "GrdRfId(RIMDMCD)"
            maxlength5      =   3
            colname6        =   "RIMBNO"
            heading6        =   "BNo                         "
            datafld6        =   "RimBNo"
            datatype6       =   1
            mask6           =   "######0"
            recalcparent6   =   "GrdRfId(RIMDMCD)"
            maxlength6      =   7
            colname7        =   "RIMDMCD"
            heading7        =   "Design Cd                        "
            datafld7        =   "RimDmCd"
            datatype7       =   4
            recalcon7       =   "GrdRfId(RIMBYY),GrdRfId(RIMBCHR),GrdRfId(RIMBNO)"
            maxlength7      =   15
            tooltiptext7    =   "Enter Design Code"
         End
      End
   End
End
Attribute VB_Name = "EmrFrmRfIdMst"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
  '---- General Comments on the Entry --------'
  Option Explicit
  Dim moCn As MwfLib.MDOConnection   '*** (Bef speed) bef 24/11/06
  Dim mb_AddRow As Boolean
  Dim mb_MultiCoCd As Boolean
  
  Private Sub ADC_Load()
    Set moCn = ADC.Connection
    mb_MultiCoCd = IIF(moCn.GetFldVal("Select hImpComLoc From Head Where HCoCd= '" + ctSelfCoCd + "' and HCd= '" + ctSelfCmCd + "'") = "Y", True, False)
    
    Set ADC.FirKeyCtl = IIF(mb_MultiCoCd = True, ADC("wRimBCoCd"), ADC("WRIMNOFR"))
    Set ADC.FirNKeyCtl = GrdRfId
    ADC.Child = "GrdRfId"
    Set GrdRfId.PrevCtl = ADC
    Set GrdRfId.NextCtl = ADC
    GrdRfId.HotKey = "A"
    mb_AddRow = True

  End Sub
  
  Private Sub ADC_KeyWhen()
    '*** Set the frame FraNKeyAll as the top most frame
    FraNKeyAll.ZOrder (0)
    '4.1.0.0
    ADC("WRIMBCOCD") = gs_CoCd
    If mb_MultiCoCd = False Then
      ADC("WRIMBCOCD").Enabled = False
    Else
      ADC("WRIMBCOCD").Enabled = True
    End If
  End Sub
  
  Private Sub ADC_KeyValid(pr_Cancel As Boolean, pr_ErrMsg As String)
    '*** Check whether the Date Tag is valid
    'Call atxt_FldChange(0, UCase("wDtTag"), pr_Cancel, pr_ErrMsg, ADC("wDtTag"))
    'If pr_Cancel = True Then Exit Sub
  End Sub
  
  Private Sub atxt_FldWhen(Index As Integer, ByVal IdName As String, Cancel As Boolean, ErrMsg As String)
    
    Select Case IdName
    
    Case Is = UCase("wRimNoFr"), UCase("wRimNoTo")
      Call HlpList.RimNo

    Case Is = UCase("WRIMDMCDFR"), UCase("WRIMDMCDTO")        '4.1.4
      Call HlpList.DmCd("DM")

    Case Is = UCase("WRIMBYYFR"), UCase("WRIMBYYTO")          '4.1.4
      Call HlpList.PMCd("YY")

    Case Is = UCase("WRIMBCHRFR"), UCase("WRIMBCHRTO")        '4.1.4
      Call HlpList.PMCd("BCHR")
      
    Case Is = UCase("WRIMBNOFR")                              '4.1.4
      Call HlpList.BNo(ADC("WRimBCoCd"), ADC("WRIMBYYFR"), ADC("WRimBChrFr"), "Y")

    Case Is = UCase("WRIMBNOTO")                              '4.1.4
      Call HlpList.BNo(ADC("WRimBCoCd"), ADC("WRIMBYYTO"), ADC("WRimBChrTo"), "Y")
      
    End Select
  
  End Sub
  
  Private Sub ATXT_FldChange(Index As Integer, ByVal IdName As String, Cancel As Boolean, ErrMsg As String, ByVal pv_NewValue As Variant)
    Select Case IdName
        '4.1.0.0
        Case Is = UCase("WRIMBCOCD")
            If pv_NewValue = "" Then Cancel = True: ErrMsg = "Company Code Cannot Be Blank": Exit Sub
            Cancel = Not moCn.RecSeek("Select 'x' from Head Where HCoCd='" + pv_NewValue + "'  and HCd='" + ctSelfCmCd + "' ")
            If Cancel = True Then ErrMsg = "Invalid Company Code": Exit Sub

    End Select
  End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
' ***** Manali 3.10.0 - 09/05/12 - Bar Code (Ctrl+N)
  If Not (KeyCode = vbKeyN And Shift = 2) Then Exit Sub
  If UCase(Me.ActiveControl.Name) = UCase("GrdRfId") Then
    With Me.ActiveControl
      If .Col = 3 Then
            mb_AddRow = False
        End If
    End With
  End If
' ***** Manali 3.10.0 - 09/05/12 - Bar Code (Ctrl+N)
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
' ***** Manali 3.10.0 - 12/04/12 - Bar Code
If KeyAscii <> vbKeyReturn Then Exit Sub
If UCase(Me.ActiveControl.Name) = UCase("GrdRfId") Then
  With Me.ActiveControl
      If .Col = 3 Then
          If .Rows - 1 = .Row Then
            .Col = 4
          End If
    End If
  End With
End If
' ***** Manali 3.10.0 - 12/04/12 - Bar Code
End Sub

  
  Private Sub ADC_SetRecSource()
    '*** Since this is an 'A' Entry, there is no RecSource associated with ADC. So set the
        'DataMode property of ADC control as xdatUnbound
    ADC.DataMode = xdatUnbound
  End Sub
  
Private Sub GrdRfId_ColChange(ByVal RowNum As Integer, ByVal ColName As String, ByVal NewValue As Variant, Cancel As Boolean, ErrMsg As String)
  With GrdRfId
    Select Case ColName
    Case Is = UCase("RimNo")
      'vk.4 RfId No changed as Bag Alias No.
      If NewValue = 0 Then Cancel = True: ErrMsg = "Bag Alias No. Cannot be 0": Exit Sub
    
    Case Is = UCase("RimBCoCd")
      Cancel = NewValue <> "" And Not moCn.RecSeek("Select 'x' from Head Where HCoCd='" + NewValue + "'  and HCd='" + ctSelfCmCd + "' ")
      If Cancel = True Then ErrMsg = "Invalid Company Code": Exit Sub
      
    Case Is = UCase("wRimBag")
      Dim ws_BagNo() As String, ws_BagNoBarCd As String
      If NewValue <> "" Then
        'uni.156 calling function to ignore CoCd if entered
        ws_BagNoBarCd = ParseBagNoBarCd(NewValue)
        ws_BagNo = Split(ws_BagNoBarCd, "/")

        If UBound(ws_BagNo) <> 2 Then Cancel = True: ErrMsg = "Invalid Bag Number Format": Exit Sub
        If Not IsNumeric(ws_BagNo(2)) Then Cancel = True: ErrMsg = "Invalid Bag No Format": Exit Sub
        Cancel = Not moCn.RecSeek("Select 'x' From Bag Where BCoCd= '" + .Value(RowNum, "RimBCoCd") + "' and " + _
                                  "BYy= '" + ws_BagNo(0) + "' And " + _
                                  "BChr = '" + ws_BagNo(1) + "' And " + _
                                  "BNo= " + ws_BagNo(2) + " /* and BLoc = '" + ctFgLoc + "' */ And BPrtKey='" + ctCurrPrtn + "' ")
        If Cancel = True Then ErrMsg = "Invalid Bag Number ": Exit Sub
      
      
        If .Mode = fgmnorm Then
          .Value(RowNum, "RimBYy") = ws_BagNo(0)
          .Value(RowNum, "RimBChr") = ws_BagNo(1)
          .Value(RowNum, "RimBNo") = ws_BagNo(2)
          .Value(RowNum, "RimDmCd") = moCn.GetFldVal("Select BOdDmCd From Bag Where BCoCd= '" + .Value(RowNum, "RimBCoCd") + "' and " + _
                                      "BYy= '" + ws_BagNo(0) + "' And " + _
                                      "BChr = '" + ws_BagNo(1) + "' And " + _
                                      "BNo= " + ws_BagNo(2) + " and /* BLoc = '" + ctFgLoc + "'  and */ BPrtKey='" + ctCurrPrtn + "' ")

        End If
      End If
      
      If .Rows - 1 = RowNum And mb_AddRow And .IsNew(RowNum) And NewValue <> "" Then
            .AddItem
            .Row = RowNum + 1
            .Col = 1
       End If
     
    Case Is = UCase("RimDmCd")
      If NewValue = "" Then Cancel = True: ErrMsg = "Design Code Cannot Be Blank": Exit Sub
      Cancel = .Value(RowNum, "wRimBag") <> "" And _
                  Not moCn.RecSeek("Select 'x' from Bag where BCoCd='" + .Value(RowNum, "RimBCoCd") + "' " + _
                                "and BYy='" + .Value(RowNum, "RimBYy") + "' and BChr='" + .Value(RowNum, "RimBChr") + "' and BNo=" + CStr(.Value(RowNum, "RimBNo")) + _
                                "and BOdDmCd='" + NewValue + "' and BLoc='" + ctFgLoc + "' and BPrtKey='" + ctCurrPrtn + "' ")
      If Cancel = True Then ErrMsg = "Design Code Not Of Bag in Fg": Exit Sub

      Cancel = .Value(RowNum, "wRimBag") = "" And _
                  Not moCn.RecSeek("Select 'x' from DsgMst where DmTcTyp='DM' and DmCd='" + NewValue + "' and DmSz='' and DmPrtKey='" + ctCurrPrtn + "' ")
      If Cancel = True Then ErrMsg = "Invalid Design Code": Exit Sub
      
      
    End Select
  End With
End Sub

Private Sub GrdRfId_ColWhen(ByVal RowNum As Integer, ByVal ColName As String, Cancel As Boolean, ErrMsg As String, ByVal Action As MwfCtl.en_fgAction)
  With GrdRfId
    Select Case ColName
      Case Is = UCase("RimBCoCd")
        Call HlpList.hCoCd
      Case Is = UCase("RimNo")  'vk.4 RimBCoCd changed as RimNo
          Cancel = Not GrdRfId.IsNew(RowNum)
          If Cancel Then ErrMsg = "Cannot Edit Bag Alias No.": Exit Sub 'vk.4 RfId No changed as Bag Alias No.
      Case Is = UCase("wRimBag")
         Call HlpList.BagNo(GrdRfId.Value(RowNum, "RimBCoCd"), "Y")
      Case Is = UCase("RimDmCd")
        If .Value(RowNum, "wRimBag") <> "" Then Cancel = True: ErrMsg = "Cannot Edit Value When Bag No Entered": Exit Sub
        Call HlpList.DmCd("DM")
      Case Is = UCase("RimBYy")
      'If Not .IsNew(RowNum) Then Cancel = True: ErrMsg = "Cannot Edit This Value": Exit Sub
      Call HlpList.PMCd("YY")
      
    Case Is = UCase("RimBChr")
      'If Not .IsNew(RowNum) Then Cancel = True: ErrMsg = "Cannot Edit This Value": Exit Sub
      Call HlpList.PMCd("BCHR")
      
    Case Is = UCase("RimBNo")
      'If Not .IsNew(RowNum) Then Cancel = True: ErrMsg = "Cannot Edit This Value": Exit Sub
        Call HlpList.BNo(.Value(RowNum, "RimBCoCd"), .Value(RowNum, "RimBYy"), .Value(RowNum, "RimBChr"), "Y")
        
    End Select
  End With
End Sub

Private Sub GrdRfId_InitKey(ByVal RowNum As Integer)
With GrdRfId
.Store "RimBCoCd", ADC("wRimBCoCd")
End With
End Sub

Private Sub GrdRfId_ReCalc(ByVal RowNum As Integer, ByVal ColName As String, ByVal OldVal As Variant, ByVal SourceName As String, ByVal SourceColName As String)
    With GrdRfId
    Select Case ColName
    Case Is = UCase("RimDmCd")
      If .Mode = fgmnorm Then
      If .Value(RowNum, "RimBYy") <> "" And .Value(RowNum, "RimBChr") <> "" And .Value(RowNum, "RimBNo") <> 0 Then
          .Value(RowNum, "RimDmCd") = moCn.GetFldVal("Select BOdDmCd From Bag Where BCoCd= '" + .Value(RowNum, "RimBCoCd") + "' and " + _
                                      "BYy= '" + .Value(RowNum, "RimBYy") + "' And " + _
                                      "BChr = '" + .Value(RowNum, "RimBChr") + "' And " + _
                                      "BNo= " + CStr(.Value(RowNum, "RimBNo")) + " And BPrtKey='" + ctCurrPrtn + "' ")

        End If
        End If
    
'    Case Is = UCase("wRimBag")
'      If .Mode = fgminit Then
'        If .Value(RowNum, "RimBNo") <> 0 Then
'          .Value(RowNum, "wRimBag") = .Value(RowNum, "RimBYy") + "/" + .Value(RowNum, "RimBChr") + "/" + Trim(CStr(.Value(RowNum, "RimBNo")))
'        End If
'      End If
'
    End Select
    End With
End Sub

Private Sub GrdRfId_RowWhen(ByVal RowNum As Integer)
  With GrdRfId
    If .IsNew(RowNum) Then
      .Value(RowNum, "RimNo") = GetMaxVal() + 1 '.MaxVal("RimNo") + 1
      If GrdRfId.Row = 1 Then
        .Value(RowNum, "RimBCoCd") = gs_CoCd
        .Value(RowNum, "RimBYy") = moCn.GetFldVal("Select hYy from Head where HCoCd= '" + ADC("wRimBCoCd") + "' and HCd= '" + ctSelfCmCd + "'")
        .Value(RowNum, "RimBChr") = moCn.GetFldVal("Select hBChr from Head where HCoCd= '" + ADC("wRimBCoCd") + "' and HCd= '" + ctSelfCmCd + "'")
      Else
        .Value(RowNum, "RimBCoCd") = .Value(RowNum - 1, "RimBCoCd")
        .Value(RowNum, "RimBYy") = .Value(RowNum - 1, "RimBYy")
        .Value(RowNum, "RimBChr") = .Value(RowNum - 1, "RimBChr")
      End If
    End If
      
  End With
   mb_AddRow = True
End Sub

Private Sub GrdRfId_SetRecSource()
    '*** Set the Record Source of the Grid GrdRfId
        
    Dim ws_Cnd As String            ', ws_OrdBy As String

    ws_Cnd = ADC.RepCond
    ws_Cnd = IIF(ws_Cnd <> "", " where " + ws_Cnd, "")

    'ws_OrdBy = " Order By RisRmCtg, RisRmId"
    GrdRfId.RecSource = " Select * from RfIdMst " + ws_Cnd + " Order By RimNo"

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
  Private Sub ATXT_Validate(Index As Integer, Cancel As Boolean)
    Cancel = ATXT(Index).Validate
  End Sub
  Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    If UnloadMode <> vbFormCode Then Cancel = True: Exit Sub
    Set moCn = Nothing    '*** (09/08/05)    '*** (Bef speed) bef 24/11/06
  End Sub


  Private Sub GrdRfId_Validate(Cancel As Boolean)
  Cancel = GrdRfId.Validate
  End Sub

Private Function GetMaxVal() As Long
    Dim mi_MaxVal As Long, i As Integer
    mi_MaxVal = 0
    With GrdRfId
    For i = 1 To GrdRfId.Rows - 1
        If .Value(i, "RimNo") > mi_MaxVal Then
            mi_MaxVal = .Value(i, "RimNo")
        End If
    Next i
    End With
    GetMaxVal = mi_MaxVal
End Function


