VERSION 5.00
Object = "{AB2CC32A-E652-42A2-B341-26ECE02F6386}#2.1#0"; "MwfCtl.ocx"
Begin VB.Form EmrFrmGrade 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Grade Master"
   ClientHeight    =   9810
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   10005
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   9810
   ScaleWidth      =   10005
   WindowState     =   2  'Maximized
   Begin MwfCtl.MWCTL_HLP MWCTL_HLP1 
      Height          =   405
      Left            =   6180
      TabIndex        =   8
      TabStop         =   0   'False
      Top             =   9360
      Visible         =   0   'False
      Width           =   525
      _ExtentX        =   926
      _ExtentY        =   714
   End
   Begin VB.Frame FRA_KEYS 
      BorderStyle     =   0  'None
      Height          =   570
      Left            =   -30
      TabIndex        =   10
      Top             =   24
      Width           =   8700
      Begin MwfCtl.MWCTL_MED ATXT 
         Height          =   285
         Index           =   4
         Left            =   1275
         TabIndex        =   0
         ToolTipText     =   "Enter Grade Type"
         Top             =   0
         Width           =   720
         _ExtentX        =   1270
         _ExtentY        =   503
         MaxLength       =   8
         DataType        =   4
         DataField       =   "GRTYP"
         IdName          =   "GRTYP"
      End
      Begin MwfCtl.MWCTL_MED ATXT 
         Height          =   285
         Index           =   1
         Left            =   2805
         TabIndex        =   1
         ToolTipText     =   "Enter Grade RM Category"
         Top             =   0
         Width           =   600
         _ExtentX        =   1058
         _ExtentY        =   503
         MaxLength       =   3
         DataType        =   4
         DataField       =   "GRRMCTG"
         IdName          =   "GRRMCTG"
      End
      Begin MwfCtl.MWCTL_MED ATXT 
         Height          =   285
         Index           =   2
         Left            =   4545
         TabIndex        =   2
         ToolTipText     =   "Enter From Customer Code"
         Top             =   0
         Width           =   1500
         _ExtentX        =   2646
         _ExtentY        =   503
         MaxLength       =   8
         DataType        =   4
         IdName          =   "WGRFRCMCD"
         CmpStr          =   "GrCmCd>="
      End
      Begin MwfCtl.MWCTL_MED ATXT 
         Height          =   285
         Index           =   0
         Left            =   4545
         TabIndex        =   3
         ToolTipText     =   "Enter To Customer Code"
         Top             =   285
         Width           =   1500
         _ExtentX        =   2646
         _ExtentY        =   503
         MaxLength       =   8
         DataType        =   4
         IdName          =   "WGRTOCMCD"
         CmpStr          =   "GrCmCd<="
      End
      Begin MwfCtl.MWCTL_MED ATXT 
         Height          =   285
         Index           =   3
         Left            =   7035
         TabIndex        =   4
         ToolTipText     =   "Enter From RM Id"
         Top             =   0
         Width           =   1500
         _ExtentX        =   2646
         _ExtentY        =   503
         MaxLength       =   8
         DataType        =   4
         IdName          =   "WGRFRRMID"
         CmpStr          =   "GrRmId>="
      End
      Begin MwfCtl.MWCTL_MED ATXT 
         Height          =   285
         Index           =   5
         Left            =   7035
         TabIndex        =   5
         ToolTipText     =   "Enter To Rm Id"
         Top             =   285
         Width           =   1500
         _ExtentX        =   2646
         _ExtentY        =   503
         MaxLength       =   8
         DataType        =   4
         IdName          =   "WGRTORMID"
         CmpStr          =   "GrRmId<="
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
         Left            =   6060
         TabIndex        =   26
         Top             =   285
         Width           =   945
      End
      Begin VB.Label ALBL 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "RM Id Fr"
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
         Left            =   6060
         TabIndex        =   25
         Top             =   0
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
         Left            =   3510
         TabIndex        =   24
         Top             =   285
         Width           =   1005
      End
      Begin VB.Label ALBL 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Cust Cd Fr"
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
         Left            =   3510
         TabIndex        =   14
         Top             =   0
         Width           =   1005
      End
      Begin VB.Label ALBL 
         BackStyle       =   0  'Transparent
         Caption         =   "RmCtg"
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
         Left            =   2100
         TabIndex        =   13
         Top             =   0
         Width           =   705
      End
      Begin VB.Label ALBL 
         BackStyle       =   0  'Transparent
         Caption         =   "Grade Type"
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
         Left            =   90
         TabIndex        =   11
         Top             =   0
         Width           =   1155
      End
   End
   Begin MwfCtl.MWCTL_BTN1 CmdCpy 
      CausesValidation=   0   'False
      Height          =   435
      Left            =   4830
      TabIndex        =   15
      ToolTipText     =   "Option To Copy Grade From Another Customer"
      Top             =   9300
      Width           =   945
      _ExtentX        =   1667
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
   Begin MwfCtl.MWCTL_ADC ADC 
      CausesValidation=   0   'False
      Height          =   555
      Left            =   60
      TabIndex        =   6
      Top             =   9300
      Width           =   9630
      _ExtentX        =   16986
      _ExtentY        =   979
      AllowDelete     =   0   'False
      BtnWidth        =   885
      BtnHeight       =   435
      BtnMaskColor    =   8454143
      FontSz          =   8.25
   End
   Begin VB.Frame FRA_NKEYS 
      BorderStyle     =   0  'None
      Enabled         =   0   'False
      Height          =   9495
      Left            =   -120
      TabIndex        =   9
      Top             =   0
      Width           =   10095
      Begin VB.Frame FraCpy 
         Height          =   1605
         Left            =   2190
         TabIndex        =   16
         Top             =   7515
         Visible         =   0   'False
         Width           =   5295
         Begin MwfCtl.MWCTL_MED ATXT 
            Height          =   285
            Index           =   11
            Left            =   1290
            TabIndex        =   17
            ToolTipText     =   "Enter Source Customer"
            Top             =   525
            Width           =   1320
            _ExtentX        =   2328
            _ExtentY        =   503
            MaxLength       =   8
            DataType        =   4
            IdName          =   "WCPYCMCDFR"
            CmpStr          =   "GrCmCd="
         End
         Begin MwfCtl.MWCTL_MED ATXT 
            Height          =   285
            Index           =   14
            Left            =   3675
            TabIndex        =   19
            ToolTipText     =   "Enter From Grade Code"
            Top             =   540
            Width           =   1320
            _ExtentX        =   2328
            _ExtentY        =   503
            MaxLength       =   12
            DataType        =   4
            IdName          =   "WCPYGRCDFR"
            CmpStr          =   "GrCd>="
         End
         Begin MwfCtl.MWCTL_BTN1 CmdCpyGo 
            Height          =   285
            Left            =   4410
            TabIndex        =   21
            ToolTipText     =   "Press This Button To Copy"
            Top             =   1170
            Width           =   585
            _ExtentX        =   1032
            _ExtentY        =   503
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
            Index           =   16
            Left            =   3675
            TabIndex        =   20
            ToolTipText     =   "Enter To Grade Code"
            Top             =   825
            Width           =   1320
            _ExtentX        =   2328
            _ExtentY        =   503
            MaxLength       =   12
            DataType        =   4
            IdName          =   "WCPYGRCDTO"
            CmpStr          =   "GrCd<="
         End
         Begin MwfCtl.MWCTL_MED ATXT 
            Height          =   285
            Index           =   6
            Left            =   1290
            TabIndex        =   18
            ToolTipText     =   "Enter Target Customer"
            Top             =   825
            Width           =   1320
            _ExtentX        =   2328
            _ExtentY        =   503
            MaxLength       =   8
            DataType        =   4
            IdName          =   "WCPYCMCDTO"
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
            Index           =   13
            Left            =   2700
            TabIndex        =   29
            Top             =   825
            Width           =   915
         End
         Begin VB.Label ALBL 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "Grade Fr"
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
            Index           =   15
            Left            =   2700
            TabIndex        =   28
            Top             =   540
            Width           =   915
         End
         Begin VB.Label ALBL 
            BackStyle       =   0  'Transparent
            Caption         =   "Target Cust"
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
            Index           =   6
            Left            =   90
            TabIndex        =   27
            Top             =   825
            Width           =   1305
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
            Height          =   300
            Index           =   1
            Left            =   -30
            TabIndex        =   23
            Top             =   90
            Width           =   5325
         End
         Begin VB.Label ALBL 
            BackStyle       =   0  'Transparent
            Caption         =   "Source Cust"
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
            Left            =   90
            TabIndex        =   22
            Top             =   525
            Width           =   1305
         End
      End
      Begin VB.Frame FraNKeyAll 
         BorderStyle     =   0  'None
         Height          =   9315
         Left            =   180
         TabIndex        =   12
         Top             =   0
         Width           =   9690
         Begin MwfCtl.MWCTL_FLX GrdGrade 
            Height          =   8430
            Left            =   0
            TabIndex        =   7
            Top             =   690
            Width           =   9645
            _ExtentX        =   17013
            _ExtentY        =   14870
            Cols            =   7
            colname1        =   "GRCMCD"
            heading1        =   "Cust Cd              "
            datafld1        =   "GRCMCD"
            datatype1       =   4
            maxlength1      =   8
            tooltiptext1    =   "Enter Customer Code"
            colname2        =   "GRCD"
            heading2        =   "Grade Cd                  "
            datafld2        =   "GrCd"
            datatype2       =   4
            maxlength2      =   8
            tooltiptext2    =   "Enter Grade Code"
            colname3        =   "GRRMID"
            heading3        =   "RM Id                   "
            datafld3        =   "GrRmId"
            datatype3       =   4
            maxlength3      =   8
            tooltiptext3    =   "Enter RM ID"
            colname4        =   "GRFRLN"
            heading4        =   "    Fr Size"
            datafld4        =   "GrFrLn"
            datatype4       =   2
            mask4           =   "####0.0000"
            maxlength4      =   10
            tooltiptext4    =   "Enter Fr Rm Szie"
            colname5        =   "GRTOLN"
            heading5        =   "    To Size"
            datafld5        =   "GrToLn"
            datatype5       =   2
            mask5           =   "####0.0000"
            maxlength5      =   10
            tooltiptext5    =   "Enter To Rm Size"
            colname6        =   "GRRMCD"
            heading6        =   "RM Cd                                      "
            datafld6        =   "GrRmCd"
            datatype6       =   4
            maxlength6      =   16
            tooltiptext6    =   "Enter RM Code"
         End
      End
   End
End
Attribute VB_Name = "EmrFrmGrade"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim moCn As MwfLib.MDOConnection   '*** (Bef speed) bef 24/11/06
Enum en_GradeFra
  Cpy = 0
End Enum
Private Sub ADC_KeyValid(pr_Cancel As Boolean, pr_ErrMsg As String)
  If ADC("GrTyp") = "" Then pr_Cancel = True: pr_ErrMsg = "Grade Type Cannot Be Blank": Exit Sub
  If ADC("GrTyp") <> "FIN" And ADC("GrRmCtg") = "" Then pr_Cancel = True: pr_ErrMsg = "RM Category Cannot Be Blank": Exit Sub
  If ADC("GrTyp") = "FIN" And ADC("GrRmCtg") <> "" Then pr_Cancel = True: pr_ErrMsg = "RM Category Should Be Blank": Exit Sub
End Sub
Private Sub ADC_KeyWhen()
  Call EnaDisaCmds(True)
' Zubin 212
  Call CmpString("F")
' Zubin 212
End Sub
Private Sub ADC_Load()
  '*** Set the First Key Control property to the field cuUsrCd
  '*** Set the First Non Key Control property to the grid cuUsrName
  '*** Set the Child Property Of ADC to GrdGrade
  '*** Set the Previous and Next Control Properties for the Grid GrdGrade as ADC
  '*** Set the Hot Key for the Grid GrdGrade as 'A'

  Set moCn = ADC.Connection    '*** (Bef speed) bef 24/11/06
  Set ADC.FirKeyCtl = ADC("GRTYP")
  Set ADC.FirNKeyCtl = GrdGrade
  ADC.Child = "GrdGrade"
  Set GrdGrade.PrevCtl = ADC
  Set GrdGrade.NextCtl = ADC
  GrdGrade.HotKey = "A"
  FraNKeyAll.ZOrder (0)
End Sub
Private Sub ADC_NonKeyValid(pr_Cancel As Boolean, pr_ErrMsg As String)
'
End Sub
Private Sub ADC_NonKeyWhen(pr_Cancel As Boolean, pr_ErrMsg As String)
  Call EnaDisaCmds(False)
End Sub
Private Sub atxt_FldChange(Index As Integer, ByVal IdName As String, Cancel As Boolean, ErrMsg As String, ByVal pv_NewValue As Variant)
  Select Case IdName
  Case Is = UCase("GrTyp")
    Cancel = Not moCn.RecSeek("Select 'x' From Param where PTyp= 'GRDTYP' " + _
                                "and PMCd = '" + pv_NewValue + "' and PSCd= ''")
    ErrMsg = "Invalid Grade Type": Exit Sub
  Case Is = UCase("GrRmCtg")
    Cancel = (ADC("GrTyp") <> "FIN" And Not moCn.RecSeek("Select 'x' From Param where PTyp= 'GRDRMCTG' " + _
                                "and PMCd = '" + ADC("GrTyp") + "' and PSCd='" + pv_NewValue + "'")) Or _
              (ADC("GrTyp") = "FIN" And pv_NewValue <> "")
                                
    ErrMsg = "Invalid RM Category": Exit Sub
  Case Is = UCase("wCpyCmCdFr")
    Cancel = Not moCn.RecSeek("Select 'x' From CustMst where CmCtg= 'C' and CmCd= '" + pv_NewValue + "' ")
    ErrMsg = "Invalid Source Customer Code": Exit Sub
  Case Is = UCase("wCpyCmCdTo")
    Cancel = Not moCn.RecSeek("Select 'x' From CustMst where CmCtg= 'C' and CmCd= '" + pv_NewValue + "' " + _
                                "and CmValidYN = 'Y' ")
    ErrMsg = "Invalid Target Customer Code": Exit Sub
  End Select
End Sub
Private Sub ATXT_FldWhen(Index As Integer, ByVal IdName As String, Cancel As Boolean, ErrMsg As String)
  '*** Set Helps For Each Field
  Select Case IdName
  Case Is = UCase("GrTyp")
    Call HlpList.PMCd("GRDTYP")
  Case Is = UCase("GrRmCtg")
    Call HlpList.PSCd("GRDRMCTG", ADC("GrTyp"), False)
  Case Is = UCase("wGrFrCmCd"), UCase("wGrToCmCd"), UCase("wCpyCmCdFr"), UCase("wCpyCmCdTo")
    Call HlpList.CustCd("C")
  Case Is = UCase("wGrFrRmId"), UCase("wGrToRmId")
    Call HlpList.PSCd("RMID", ADC("GRRmCtg"))
  Case Is = UCase("wCpyGrCdFr"), UCase("wCpyGrCdTo")
    Call HlpList.PSCd("GRDCD", ADC("GRTyp"))
  End Select
End Sub
Private Sub CmdCpy_Click()
Call DispFra(Cpy)
End Sub
Private Sub CmdCpyGo_Click()
  If ADC("WCPYCMCDFR") = "" Then DispMsg "Source Customer Cannot Be Blank", etError: Exit Sub
  If ADC("WCPYCMCDTo") = "" Then DispMsg "Target Customer Cannot Be Blank", etError: Exit Sub
  Call CpyGrades
  Call CmdCpy_Click
  End Sub
Private Sub GrdGrade_ColChange(ByVal RowNum As Integer, ByVal ColName As String, ByVal NewValue As Variant, Cancel As Boolean, ErrMsg As String)
  Select Case ColName
  Case Is = UCase("GrCmCd")
    Cancel = Not moCn.RecSeek("Select 'x' From CustMst where CmCtg= 'C' and CmCd= '" + NewValue + "' " + _
                                "and CmValidYN = 'Y' ")
    ErrMsg = "Invalid Customer Code": Exit Sub
  Case Is = UCase("GrCd")
    Cancel = Not moCn.RecSeek("select 'x' from Param where PTyp= 'GRDCD' and " + _
                                "PMCd= '" + ADC("GrTyp") + "' and PSCd= '" + NewValue + "' ")
    ErrMsg = "Invalid Grade Code": Exit Sub
  Case Is = UCase("GrRmId")
    Cancel = ADC("GrTyp") <> "FIN" And Not moCn.RecSeek("select 'x' from Param where PTyp= 'RMID' and " + _
                                "PMCd= '" + ADC("GrRmCtg") + "' and PSCd= '" + NewValue + "' ")
    ErrMsg = "Invalid Rm Id": Exit Sub
  Case Is = UCase("GrRmCd")
    ' ***** Manali 3.10.0 - 18/05/12 - 'S','L' added for 'MET'
    Cancel = ADC("GrTyp") <> "FIN" And Not moCn.RecSeek("select 'x' From RmMst where " + IIF(gs_Partition = ctCurrPrtn, " RmPrtKey='" + ctCurrPrtn + "' And ", "") + _
                                "(('" + ADC("GrTyp") + "'='MET' and " + _
                                " '" + ADC("GrRmCtg") + "' in ('G','P','X','S','L')) Or " + _
                                "RmCtg= '" + ADC("GrRmCtg") + "' ) and RmCd= '" + NewValue + "' ")
    ErrMsg = "Invalid RmCd For This Grade Type": Exit Sub
  End Select
End Sub
Private Sub GrdGrade_ColWhen(ByVal RowNum As Integer, ByVal ColName As String, Cancel As Boolean, ErrMsg As String, ByVal Action As MwfCtl.en_fgAction)
  '*** Set Helps For Each Field
  Select Case ColName
  Case Is = UCase("GrCmCd")
    Call HlpList.CustCd("C")
  Case Is = UCase("GrCd")
    Call HlpList.PSCd("GRDCD", ADC("GRTyp"))
  Case Is = UCase("GrRmId")
    If UCase(ADC("GrTyp")) = "FIN" Then Cancel = True: Exit Sub
    Call HlpList.PSCd("RMID", ADC("GRRmCtg"))
  Case Is = UCase("GrRmCd")
    If UCase(ADC("GrTyp")) = "FIN" Then Cancel = True: Exit Sub
    Call HlpList.GrRmCd(ADC("GrTyp"), GrdGrade.Value(RowNum, "GrRmId"))
  End Select
End Sub
Private Sub GrdGrade_InitKey(ByVal RowNum As Integer)
  '*** Initialise Key Fields
  With GrdGrade
    If .IsNew(RowNum) Then
      .Store "GrTyp", ADC("GrTyp")
      .Store "GrRmCtg", ADC("GrRmCtg")
    End If
  End With
End Sub
Private Sub ADC_SetRecSource()
  '*** Set the RecSource for ADC as xdatUnbound
  ADC.DataMode = xdatUnbound
  '***Bhavna added frln toln for D/C
  GrdGrade.ColProp("GrFrLn").Style = IIF(ADC("GrRmCtg") <> "D" And ADC("GrRmCtg") <> "C", fgcinvisible, fgcNormal)
  GrdGrade.ColProp("GrToLn").Style = IIF(ADC("GrRmCtg") <> "D" And ADC("GrRmCtg") <> "C", fgcinvisible, fgcNormal)
End Sub
Private Sub GrdGrade_SetRecSource()
  '*** Set the Record Source of the Grid GrdGrade
  Dim wCnd As String
  wCnd = ADC.RepCond
  wCnd = IIF(wCnd <> "", " and " + wCnd, "")

  GrdGrade.RecSource = " Select * from Grade " + _
                  " Where GrTyp='" + ADC("GrTyp") + "' and GrRmCtg='" + ADC("GrRmCtg") + "' " + wCnd + _
                  " order by GrCmCd, GrCd, GrRmId, GrFrLn, GrToLn, GrRmCd"
End Sub
Private Sub CpyGrades()
  Dim ws_GrdCnd As String
' Zubin 212
  Call CmpString("C")
' Zubin 212
  ws_GrdCnd = ADC.RepCond
  ws_GrdCnd = IIF(ws_GrdCnd <> "", " and " + ws_GrdCnd, "")

'*** Beginning the copy process
  GrdGrade.StartCopy
  MWLib.BeginProcess Me, "Copying Grades ..."
  Dim wRsCustGrade As MwfLib.MDORowSet, wi_Row As Integer
  'Set wRsCustGrade = moCn.OpenRes("Select * from Grade where GrTyp= '" + ADC("GrTyp") + "' and " + _
             "GrRmCtg= '" + ADC("GrRmCtg") + ws_GrdCnd + _
             " Order By GrCd,GrRmId")
             
' Zubin 212
'***********Bhavna added sz range
  Set wRsCustGrade = moCn.OpenRes("Select * from Grade Src where GrTyp= '" + ADC("GrTyp") + "' and " + _
             "GrRmCtg= '" + ADC("GrRmCtg") + "'" + ws_GrdCnd + " " + _
             "And Not Exists (Select 'x' from Grade Tgt where " + _
             "GrCmCd = '" + ADC("wCpyCmCdTo") + "' and Tgt.GrTyp=Src.GrTyp and  " + _
             "Tgt.GrCd=Src.GrCd and Tgt.GrRmCtg=Src.GrRmCtg and Tgt.GrFrLn=Src.GrFrLn and " + _
             "Tgt.GrToLn=Src.GrToLn and Tgt.GrRmId=Src.GrRmId ) " + _
             "Order By GrCd,GrRmId ")
' Zubin 212

  With wRsCustGrade
    Do While Not (.EOF Or .BOF)
      MWLib.RunningProcess "Copying Grades for RmId  " + !GrRmId
      GrdGrade.AddItem
      wi_Row = GrdGrade.Rows - 1
      GrdGrade.Value(wi_Row, "GrCmCd") = ADC("wCpyCmCdTo")
      GrdGrade.Value(wi_Row, "GrCd") = !GrCd
      GrdGrade.Value(wi_Row, "GrRmId") = !GrRmId
      GrdGrade.Value(wi_Row, "GrFrLn") = !GrFrLn            '**Bhavna added sz range
      GrdGrade.Value(wi_Row, "GrToLn") = !GrToLn            '**Bhavna added sz range
      GrdGrade.Value(wi_Row, "GrRmCd") = !GrRmCd
      GrdGrade.SaveRec (wi_Row)
      .MoveNext
    Loop
  End With
  MWLib.EndProcess Me
  GrdGrade.EndCopy
  Set wRsCustGrade = Nothing
  '*** End of the copy process
End Sub

' Zubin 212
Private Sub CmpString(ByVal mTyp As String)
  'mTyp = "F" (Filter)
  'mTyp = "C" (Copy)

  If UCase(mTyp) = UCase("F") Then
    ADC("wGrFrCmCd").CmpStr = "GrCmCd>="
    ADC("wGrToCmCd").CmpStr = "GrCmCd<="
    ADC("wGrFrRmId").CmpStr = "GrRmId>="
    ADC("wGrToRmId").CmpStr = "GrRmId<="
    ADC("wCpyCmCdFr").CmpStr = ""
    ADC("wCpyGrCdFr").CmpStr = ""
    ADC("wCpyGrCdTo").CmpStr = ""
  ElseIf UCase(mTyp) = UCase("C") Then
    ADC("wGrFrCmCd").CmpStr = ""
    ADC("wGrToCmCd").CmpStr = ""
    ADC("wGrFrRmId").CmpStr = ""
    ADC("wGrToRmId").CmpStr = ""
    ADC("wCpyCmCdFr").CmpStr = "GrCmCd="
    ADC("wCpyGrCdFr").CmpStr = "GrCd>="
    ADC("wCpyGrCdTo").CmpStr = "GrCd<="
  
  End If
End Sub

Private Sub DispFra(ByVal pv_GradeFra As en_GradeFra)
  '*** This is the common procedure which is used to toggle between hiding and displaying
      'the frames (i.e. if the frame is visible then make it invisible and viseversa)
      'associated with a command button
  
  Select Case pv_GradeFra
  Case Is = Cpy
    '*** If Frame FraCpy is visible then make it invisible and enable all the command buttons
    '*** If Frame FraCpy is invisible then make it visible and disable all the
        'command buttons except the one which is clicked. Change the background colour of
        'the clicked command button to white and make the font bold
    If FraCpy.Visible = True Then
      FraCpy.Visible = False
      FraCpy.Enabled = False
      Call EnaDisaCmds(False, CmdCpy)
      CmdCpy.SetFocus
    Else
      FraCpy.Visible = True
      FraCpy.Enabled = True
      ADC("WCPYCMCDFR").SetFocus
      FraCpy.ZOrder
      Call EnaDisaCmds(True, CmdCpy)
    End If
  End Select
End Sub
Private Sub EnaDisaCmds(ByVal pv_ShowFra As Boolean, Optional ByVal pv_CmdBut As MWCTL_BTN1)
  If Not pv_CmdBut Is Nothing Then
    If pv_ShowFra = True Then
      pv_CmdBut.CausesValidation = True
    End If
  End If
  FraNKeyAll.Enabled = Not pv_ShowFra
  ADC.AllowSave = Not pv_ShowFra
  CmdCpy.Enabled = Not pv_ShowFra
  If Not pv_CmdBut Is Nothing Then pv_CmdBut.Enabled = True
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
Private Sub GrdGrade_Validate(Cancel As Boolean)
  Cancel = GrdGrade.Validate
End Sub
