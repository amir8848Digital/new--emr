VERSION 5.00
Object = "{AB2CC32A-E652-42A2-B341-26ECE02F6386}#2.0#0"; "MwfCtl.ocx"
Begin VB.Form EmrFrmPpcDsgLd 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Design Production Load"
   ClientHeight    =   9810
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   15210
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   9810
   ScaleWidth      =   15210
   WindowState     =   2  'Maximized
   Begin MwfCtl.MWCTL_HLP MWCTL_HLP1 
      Height          =   435
      Left            =   5610
      TabIndex        =   6
      TabStop         =   0   'False
      Top             =   9360
      Visible         =   0   'False
      Width           =   435
      _ExtentX        =   767
      _ExtentY        =   767
   End
   Begin MwfCtl.MWCTL_ADC ADC 
      CausesValidation=   0   'False
      Height          =   555
      Left            =   180
      TabIndex        =   7
      Top             =   9285
      Width           =   10455
      _ExtentX        =   18441
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
      Height          =   735
      Left            =   0
      TabIndex        =   9
      Top             =   24
      Width           =   11130
      Begin MwfCtl.MWCTL_MED ATXT 
         Height          =   285
         Index           =   0
         Left            =   1350
         TabIndex        =   0
         ToolTipText     =   "Enter From Production Category"
         Top             =   60
         Width           =   1470
         _ExtentX        =   2593
         _ExtentY        =   503
         MaxLength       =   5
         DataType        =   4
         IdName          =   "WPRDCTGFR"
         CmpStr          =   "DldPrdCtg >="
      End
      Begin MwfCtl.MWCTL_MED ATXT 
         Height          =   285
         Index           =   1
         Left            =   1350
         TabIndex        =   1
         ToolTipText     =   "Enter To Production Category"
         Top             =   345
         Width           =   1470
         _ExtentX        =   2593
         _ExtentY        =   503
         MaxLength       =   5
         DataType        =   4
         IdName          =   "WPRDCTGTO"
         CmpStr          =   "DldPrdCtg<="
      End
      Begin MwfCtl.MWCTL_MED ATXT 
         Height          =   285
         Index           =   2
         Left            =   4920
         TabIndex        =   2
         ToolTipText     =   "Enter From Design Code"
         Top             =   60
         Width           =   2505
         _ExtentX        =   4419
         _ExtentY        =   503
         MaxLength       =   15
         DataType        =   4
         IdName          =   "WDMCDFR"
         CmpStr          =   "DldDmCd >="
      End
      Begin MwfCtl.MWCTL_MED ATXT 
         Height          =   285
         Index           =   3
         Left            =   4920
         TabIndex        =   3
         ToolTipText     =   "Enter To Design Code"
         Top             =   345
         Width           =   2505
         _ExtentX        =   4419
         _ExtentY        =   503
         MaxLength       =   15
         DataType        =   4
         IdName          =   "WDMCDTO"
         CmpStr          =   "DldDmCd <="
      End
      Begin MwfCtl.MWCTL_MED ATXT 
         Height          =   315
         Index           =   4
         Left            =   9000
         TabIndex        =   4
         ToolTipText     =   "Enter Sort Option"
         Top             =   60
         Width           =   1620
         _ExtentX        =   2858
         _ExtentY        =   556
         MaxLength       =   8
         DataType        =   4
         IdName          =   "WSRT"
      End
      Begin VB.Label ALBL 
         BackStyle       =   0  'Transparent
         Caption         =   "Prd Ctg Fr"
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
         Left            =   210
         TabIndex        =   15
         Top             =   60
         Width           =   1215
      End
      Begin VB.Label ALBL 
         BackStyle       =   0  'Transparent
         Caption         =   "             To"
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
         Left            =   210
         TabIndex        =   14
         Top             =   345
         Width           =   1215
      End
      Begin VB.Label ALBL 
         BackStyle       =   0  'Transparent
         Caption         =   "Dsg Cd Fr"
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
         Left            =   3780
         TabIndex        =   13
         Top             =   60
         Width           =   1215
      End
      Begin VB.Label ALBL 
         BackStyle       =   0  'Transparent
         Caption         =   "             To"
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
         Left            =   3780
         TabIndex        =   12
         Top             =   345
         Width           =   1215
      End
      Begin VB.Label ALBL 
         BackStyle       =   0  'Transparent
         Caption         =   "Sort"
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
         Left            =   8490
         TabIndex        =   11
         Top             =   60
         Width           =   795
      End
   End
   Begin VB.Frame FRA_NKEYS 
      BorderStyle     =   0  'None
      Enabled         =   0   'False
      Height          =   8610
      Left            =   -90
      TabIndex        =   8
      Top             =   750
      Width           =   11235
      Begin VB.Frame FraNKeyAll 
         BorderStyle     =   0  'None
         Height          =   8535
         Left            =   210
         TabIndex        =   10
         Top             =   -30
         Width           =   10710
         Begin MwfCtl.MWCTL_FLX GrdDld 
            Height          =   8430
            Left            =   60
            TabIndex        =   5
            Top             =   60
            Width           =   10470
            _ExtentX        =   18468
            _ExtentY        =   14870
            Cols            =   9
            colname1        =   "DLDPRDCTG"
            heading1        =   "Prd Ctg"
            datafld1        =   "DldPrdCtg"
            datatype1       =   4
            maxlength1      =   5
            tooltiptext1    =   "Enter Production Category"
            colname2        =   "DLDDMCD"
            heading2        =   "Design  Cd                   "
            datafld2        =   "DldDmCd"
            datatype2       =   4
            recalcparent2   =   "GrdDld(DLDP1),GrdDld(DLDP2),GrdDld(DLDP3),GrdDld(DLDP4),GrdDld(DLDP5)"
            maxlength2      =   15
            tooltiptext2    =   "Enter Design Code"
            colname3        =   "DLDWH"
            heading3        =   "W/ H Set"
            datafld3        =   "DldWH"
            datatype3       =   4
            recalcparent3   =   "GrdDld(DLDP1),GrdDld(DLDP2),GrdDld(DLDP3),GrdDld(DLDP4),GrdDld(DLDP5)"
            maxlength3      =   2
            tooltiptext3    =   "Enter Type  of Setting"
            colname4        =   "DLDP1"
            heading4        =   "P1              "
            datafld4        =   "DldP1"
            datatype4       =   2
            mask4           =   "######0.0"
            recalcon4       =   "GrdDld(DLDDMCD),GrdDld(DLDWH)"
            maxlength4      =   9
            tooltiptext4    =   "Enter P1 Load Per Piece"
            colname5        =   "DLDP2"
            heading5        =   "P2              "
            datafld5        =   "DldP2"
            datatype5       =   2
            mask5           =   "######0.0"
            recalcon5       =   "GrdDld(DLDDMCD),GrdDld(DLDWH)"
            maxlength5      =   9
            tooltiptext5    =   "Enter P2 Load Per Piece"
            colname6        =   "DLDP3"
            heading6        =   "P3              "
            datafld6        =   "DldP3"
            datatype6       =   2
            mask6           =   "######0.0"
            recalcon6       =   "GrdDld(DLDDMCD),GrdDld(DLDWH)"
            maxlength6      =   9
            tooltiptext6    =   "Enter P3 Load Per Piece"
            colname7        =   "DLDP4"
            heading7        =   "P4              "
            datafld7        =   "DldP4"
            datatype7       =   2
            mask7           =   "######0.0"
            recalcon7       =   "GrdDld(DLDDMCD),GrdDld(DLDWH)"
            maxlength7      =   9
            tooltiptext7    =   "Enter P4 Load Per Piece"
            colname8        =   "DLDP5"
            heading8        =   "P5              "
            datafld8        =   "DldP5"
            datatype8       =   2
            mask8           =   "######0.0"
            recalcon8       =   "GrdDld(DLDDMCD),GrdDld(DLDWH)"
            maxlength8      =   9
            tooltiptext8    =   "Enter P5 Load Per Piece"
         End
      End
   End
End
Attribute VB_Name = "EmrFrmPpcDsgLd"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim moCn As MwfLib.MDOConnection   '*** (Bef speed) bef 24/11/06
Private Sub ADC_Load()
  '*** Set the First Key Control property to the working field wDtTag
  '*** Set the First Non Key Control property to the grid GrdDld
  '*** Set the Child Property Of ADC to GrdDld
  '*** Set the Previous and Next Control Properties for the Grid GrdDld
  '*** Set the Hot Key for the Grid GrdDld as 'A'
  
  Set moCn = ADC.Connection    '*** (Bef speed) bef 24/11/06
  Set ADC.FirKeyCtl = ADC("wPrdCtgFr")
  Set ADC.FirNKeyCtl = GrdDld
  ADC.Child = "GrdDld"
  Set GrdDld.PrevCtl = ADC
  Set GrdDld.NextCtl = ADC
  GrdDld.HotKey = "A"
  
  Call Form_Activate
  Call GrdPropSetPPC(GrdDld, "Dld")   '*** (Jen 2.14 Next)
End Sub
Private Sub ADC_KeyWhen()
  '*** Set the frame FraNKeyAll as the top most frame
  FraNKeyAll.ZOrder (0)
End Sub
Private Sub atxt_FldChange(Index As Integer, ByVal IdName As String, Cancel As Boolean, ErrMsg As String, ByVal pv_NewValue As Variant)
  Select Case UCase(IdName)
    Case Is = UCase("wSrt")
    If pv_NewValue <> "" Then
      Cancel = Not moCn.RecSeek("Select PMCd From Param Where PTyp = 'SRTDLD' " + _
                                "And PMCd = '" + pv_NewValue + "'")
    End If
    If Cancel = True Then ErrMsg = "Invalid Sort Option": Exit Sub
  End Select
End Sub
Private Sub ATXT_FldWhen(Index As Integer, ByVal IdName As String, Cancel As Boolean, ErrMsg As String)
  '*** Set Helps For Each Field
  Select Case UCase(IdName)
  Case Is = UCase("wPrdCtgFr"), UCase("wPrdCtgTo")
    Call HlpList.PMCd("PRDCTG")
  Case Is = UCase("wDmCdFr"), UCase("wDmCdTo")
    Call HlpList.DmCd("DM")
  Case Is = UCase("wSrt")
    Call HlpList.PMCd("SRTDLD")
  End Select
End Sub

Private Sub GrdDld_ColWhen(ByVal RowNum As Integer, ByVal ColName As String, Cancel As Boolean, ErrMsg As String, ByVal Action As MwfCtl.en_fgAction)
  Select Case ColName
  Case Is = UCase("DldPrdCtg")
    Call HlpList.PMCd("PRDCTG")
  Case Is = UCase("DldDmCd")
    'Call HlpList.DmCd("DM")    '*** (Bef 2.14 Next)
    Call HlpList.DmCd("DM", , GrdDld.Value(RowNum, "DldPrdCtg"))    '*** (Jen 2.14 Next)
  Case Is = UCase("DldWH")
    Call HlpList.PMCd("WH")
  End Select
End Sub
Private Sub GrdDld_ColChange(ByVal RowNum As Integer, ByVal ColName As String, ByVal NewValue As Variant, Cancel As Boolean, ErrMsg As String)
  Select Case UCase(ColName)
  Case Is = UCase("DldPrdCtg")
    Cancel = Not moCn.RecSeek("Select PMCd From Param Where " + _
                              "PTyp= 'PRDCTG' And PMCd = '" + NewValue + "'")
    If Cancel = True Then ErrMsg = "Invalid Production Category No": Exit Sub
'  Case Is = UCase("DldCas"), UCase("DldFin"), UCase("DldEi")
'    'UCase("DldWSet"), UCase("DldHSet")
'    If NewValue <= 0 Then
'      Cancel = True
'      ErrMsg = "Value of " + Trim(GrdDld.ColProp(ColName).Heading) + " Should Be > 0"
'      Exit Sub
'    End If
  Case Is = UCase("DldDmCd")
    If NewValue <> "" Then
      Cancel = Not moCn.RecSeek("Select DmCd From DsgMst Where DmTcTyp in ('DM','SM') And " + _
                                "DmCd = '" + NewValue + "' And DmPrdCtg = '" + GrdDld.Value(RowNum, "DldPrdCtg") + "'")
    End If
    If Cancel = True Then ErrMsg = "Invalid DsgCd Or DsgCd Does Not Belong To The Selected PrdCtg": Exit Sub
  
  Case Is = UCase("DldWH")
    Cancel = Not moCn.RecSeek("Select PMCd From Param Where " + _
                              "PTyp= 'WH' And PMCd = '" + NewValue + "'")
    If Cancel = True Then ErrMsg = "Invalid Setting Option": Exit Sub
  End Select
End Sub
Private Sub GrdDld_ReCalc(ByVal RowNum As Integer, ByVal ColName As String, ByVal OldVal As Variant, ByVal SourceName As String, ByVal SourceColName As String)
  With GrdDld
    Select Case UCase(ColName)
    Case Is = UCase("DldP1"), UCase("DldP2"), UCase("DldP3"), UCase("DldP4"), UCase("DldP5")    '*** (Jen 2.14 Next)
    'Case Is = UCase("DldWSet"), UCase("DldCas"), UCase("DldFin"), UCase("DldHSet"), UCase("DldEi")   '*** (Bef 2.14 Next)
      
      If .Mode = fgmnorm Then
        If .Value(RowNum, ColName) = 0 And .Value(RowNum, "DldDmCd") <> "" And .Value(RowNum, "DldWH") <> "" Then
          .Value(RowNum, ColName) = moCn.GetFldVal("Select " + ColName + " from DsgLoad " + _
                  "where DldDmCd= '' and DldWH= '" + .Value(RowNum, "DldWH") + "' " + _
                  "and DldPrdCtg= (Select DmPrdCtg from DsgMst where DmTcTyp in ('DM') " + _
                  "and DmCd= '" + .Value(RowNum, "DldDmCd") + "' and DmSz= '') ")
        End If
      End If
    End Select
  End With
End Sub
Private Sub ADC_SetRecSource()
  ADC.DataMode = xdatUnbound
End Sub
Private Sub GrdDld_SetRecSource()
  Dim ws_Cnd As String, ws_OrdBy As String
  
  ws_Cnd = ADC.RepCond
  ws_Cnd = IIF(ws_Cnd <> "", " where " + ws_Cnd, "")
  
  If ADC("wSrt") = "" Then
    ws_OrdBy = " Order By DldPrdCtg,DldDmCd "
  Else
    ws_OrdBy = "Order By " + moCn.GetFldVal("Select PDesc225 from Param " + _
               "where PTyp='SRTDLD' and PMCd='" + ADC("wSrt") + "'")
  End If
  
  GrdDld.RecSource = " Select * from DsgLoad " + ws_Cnd + ws_OrdBy
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
Private Sub GrdDld_Validate(Cancel As Boolean)
  Cancel = GrdDld.Validate
End Sub



