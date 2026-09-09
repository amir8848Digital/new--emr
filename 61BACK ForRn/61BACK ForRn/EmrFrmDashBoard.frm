VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Object = "{AB2CC32A-E652-42A2-B341-26ECE02F6386}#2.1#0"; "MwfCtl.ocx"
Begin VB.Form EmrFrmDashBoard 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "DashBoard Report"
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
      Height          =   405
      Left            =   8910
      TabIndex        =   1
      TabStop         =   0   'False
      Top             =   9360
      Visible         =   0   'False
      Width           =   555
      _ExtentX        =   979
      _ExtentY        =   714
   End
   Begin VB.Frame FRA_KEYS 
      BorderStyle     =   0  'None
      Height          =   1410
      Left            =   0
      TabIndex        =   3
      Top             =   0
      Width           =   15135
      Begin MwfCtl.MWCTL_MED ATXT 
         Height          =   285
         Index           =   15
         Left            =   720
         TabIndex        =   0
         ToolTipText     =   "Enter Certificate Category"
         Top             =   120
         Visible         =   0   'False
         Width           =   420
         _ExtentX        =   741
         _ExtentY        =   503
         DataType        =   4
         IdName          =   "WCDCTG"
         CmpStr          =   "CdCtg="
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "DashBoard Report "
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Index           =   1
         Left            =   6840
         TabIndex        =   8
         Top             =   240
         Width           =   1785
      End
   End
   Begin VB.Frame FRA_NKEYS 
      BorderStyle     =   0  'None
      Enabled         =   0   'False
      Height          =   9300
      Left            =   -120
      TabIndex        =   2
      Top             =   0
      Width           =   15375
      Begin VB.Frame FraNKeyAll 
         BorderStyle     =   0  'None
         Height          =   9165
         Left            =   150
         TabIndex        =   6
         Top             =   90
         Width           =   14940
         Begin MSComDlg.CommonDialog Cd1 
            Left            =   11760
            Top             =   8160
            _ExtentX        =   847
            _ExtentY        =   847
            _Version        =   393216
         End
      End
   End
   Begin MwfCtl.MWCTL_ADC ADC 
      CausesValidation=   0   'False
      Height          =   555
      Left            =   0
      TabIndex        =   4
      Top             =   9360
      Width           =   15090
      _ExtentX        =   26617
      _ExtentY        =   979
      AllowAdd        =   0   'False
      AllowFind       =   0   'False
      AllowDelete     =   0   'False
      BtnWidth        =   885
      BtnHeight       =   435
      BtnMaskColor    =   8454143
      FontSz          =   8.25
   End
   Begin MwfCtl.MWCTL_BTN1 CmdExpExc 
      Height          =   465
      Left            =   6240
      TabIndex        =   5
      ToolTipText     =   "Export to Excel Option"
      Top             =   9360
      Width           =   960
      _ExtentX        =   1693
      _ExtentY        =   820
      BackColor       =   -2147483633
      MaskColor       =   8454143
      Caption         =   "&Export to Excel"
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
   Begin VB.OLE OLE1 
      Class           =   "AcroExch.Document.7"
      Enabled         =   0   'False
      Height          =   375
      Left            =   2640
      TabIndex        =   7
      Top             =   9600
      Visible         =   0   'False
      Width           =   975
   End
End
Attribute VB_Name = "EmrFrmDashBoard"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim moCn As MwfLib.MDOConnection

Private Sub ADC_Load()
  Set moCn = ADC.Connection
  'Set ADC.FirKeyCtl = ADC("wCdCtg")
     
End Sub

Private Sub ADC_SetRecSource()
  '*** Since this is an 'A' Entry, there is no RecSource associated with ADC. So set the
      'DataMode property of ADC control as xdatUnbound

  ADC.DataMode = xdatUnbound
End Sub

Private Sub Form_GotFocus()
  SetMDISysColorsAfterLogin

End Sub
Private Sub Form_Load()
  Dim ws_BtnStr As String
  Call SetObjColors(Me, ws_BtnStr)
  
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
    
    Set moCn = Nothing
End Sub

Private Sub HideAllFras()
  '*** This procedure hides all the frames other than FraNKeyAll
  FraNKeyAll.Visible = True
  
End Sub

Private Sub CmdExpExc_Click()
  'copy records from DashBoard table to excel file
  On Error GoTo Err_Hndlr
  Dim wExcVer As String, wExtn As String, sFilePath As String, wErrMsg As String
  Dim oExcel, oBook, oSheet
  Set oExcel = GetExcelObj()
  wExcVer = oExcel.version
  wExtn = IIF(wExcVer >= "12.0", ".xlsx", ".xls")
    
  Cd1.DialogTitle = "Save Excel File As"
  Cd1.FileName = "DashBoardReport" + wExtn
  Cd1.Filter = "Excel Files (*" + wExtn + ") |*" + wExtn
  Cd1.Flags = cdlOFNOverwritePrompt Or cdlOFNLongNames Or cdlOFNPathMustExist Or cdlOFNExtensionDifferent
  Cd1.CancelError = True
  
  If (SelectFileName(wExtn, sFilePath)) = False Then Exit Sub
  FileCopy App.Path + "\Template\DashBoard Template" + wExtn, sFilePath
  oExcel.Visible = False: oExcel.DisplayAlerts = False
  
  Set oBook = oExcel.Workbooks.Open(sFilePath)
  Set oSheet = oBook.worksheets(1)

  Dim wRsHd As MwfLib.MDORowSet, wSqlHd As String
  Dim wRsDBPrm As MwfLib.MDORowSet, wSqlDBPrm As String
  Dim wRsDBPrmRpt As MwfLib.MDORowSet, wSqlDBPrmRpt As String
  Dim wRsDB As MwfLib.MDORowSet, wSql As String
  
  wSqlHd = "select HBaseCurCd from Head where HCoCd ='ZZZ' "
  Set wRsHd = moCn.OpenRes(wSqlHd)
  oSheet.Cells(1, 4) = wRsHd!HBaseCurCd
  
  wSqlDBPrm = "Select PNum from Param where PTyp ='DBPARAM' and PMCd ='DBPARAM' "
  Set wRsDBPrm = moCn.OpenRes(wSqlDBPrm)
  
  oSheet.Range("A1:I12").Replace What:="Over Due > 30 Days", Replacement:="Over Due > " + CStr(wRsDBPrm!PNum) + " Days"
  oSheet.Range("A1:I12").Replace What:="Due In Next 30 Days", Replacement:="Due In Next " + CStr(wRsDBPrm!PNum) + " Days"
  
  
  wSqlDBPrmRpt = "Select PNum,PNum1 from Param where PTyp ='DBPRMRPT' and PMCd ='DBPRMRPT' "
  Set wRsDBPrmRpt = moCn.OpenRes(wSqlDBPrmRpt)
  oSheet.Cells(2, 4) = wRsDBPrmRpt!PNum
  oSheet.Cells(2, 5) = wRsDBPrmRpt!PNum1
  
  'Copying DashBoard report1
  wSql = "Select * From DashBoard where DbTyp ='SalOrd' "
  Set wRsDB = moCn.OpenRes(wSql)
  Do While Not (wRsDB.EOF Or wRsDB.BOF)
    oSheet.Cells(5, 4) = wRsDB!DbValToday / oSheet.Cells(2, 4)
    oSheet.Cells(5, 5) = wRsDB!DbQtyToday / oSheet.Cells(2, 5)
    oSheet.Cells(5, 6) = wRsDB!DbValYTD / oSheet.Cells(2, 4)
    oSheet.Cells(5, 7) = wRsDB!DbQtyYTD / oSheet.Cells(2, 5)
    oSheet.Cells(5, 8) = wRsDB!DbValPrevYTD / oSheet.Cells(2, 4)
    oSheet.Cells(5, 9) = wRsDB!DbQtyPrevYTD / oSheet.Cells(2, 5)
    'oSheet.Range("A1:I13").Replace What:="wval", Replacement:=wRsDB!DbValToday
    wRsDB.MoveNext
  Loop

  
  wSql = "Select * From DashBoard where DbTyp ='ExpOrd' "
  Set wRsDB = moCn.OpenRes(wSql)
  Do While Not (wRsDB.EOF Or wRsDB.BOF)
    oSheet.Cells(7, 4) = wRsDB!DbValToday / oSheet.Cells(2, 4)
    oSheet.Cells(7, 5) = wRsDB!DbQtyToday / oSheet.Cells(2, 5)
    oSheet.Cells(7, 6) = wRsDB!DbValYTD / oSheet.Cells(2, 4)
    oSheet.Cells(7, 7) = wRsDB!DbQtyYTD / oSheet.Cells(2, 5)
    oSheet.Cells(7, 8) = wRsDB!DbValPrevYTD / oSheet.Cells(2, 4)
    oSheet.Cells(7, 9) = wRsDB!DbQtyPrevYTD / oSheet.Cells(2, 5)
    wRsDB.MoveNext
  Loop
    
  wSql = "Select * From DashBoard where DbTyp ='OrdPnd' "
  Set wRsDB = moCn.OpenRes(wSql)
  Do While Not (wRsDB.EOF Or wRsDB.BOF)
    oSheet.Cells(11, 4) = wRsDB!DbValToday / oSheet.Cells(2, 4)
    oSheet.Cells(11, 5) = wRsDB!DbQtyToday / oSheet.Cells(2, 5)
    oSheet.Cells(11, 6) = wRsDB!DbValYTD / oSheet.Cells(2, 4)
    oSheet.Cells(11, 7) = wRsDB!DbQtyYTD / oSheet.Cells(2, 5)
    oSheet.Cells(11, 8) = wRsDB!DbValPrevYTD / oSheet.Cells(2, 4)
    oSheet.Cells(11, 9) = wRsDB!DbQtyPrevYTD / oSheet.Cells(2, 5)
    wRsDB.MoveNext
  Loop
  
  
  'Copying DashBoard report2
  Dim i As Integer, wRow As Integer
  wRow = 1
  wSql = "Select * From DashBoard where DbTyp ='Orders' "
  Set wRsDB = moCn.OpenRes(wSql)
  Do While Not (wRsDB.EOF Or wRsDB.BOF)
    If wRow = 1 Then
      i = 17
    Else
      oSheet.Rows(CStr(i) + ":" + CStr(i)).Insert
      oSheet.Range("B6" + ":I6").Copy
      oSheet.Range("B" + CStr(i) + ":I" + CStr(i)).Select
      oSheet.Paste
    End If
    oSheet.Cells(i, 2) = wRsDB!DbTyp
    oSheet.Cells(i, 3) = wRsDB!DbSubTyp
    oSheet.Cells(i, 4) = wRsDB!DbValToday
    oSheet.Cells(i, 5) = wRsDB!DbQtyToday / oSheet.Cells(2, 4)
    oSheet.Cells(i, 6) = wRsDB!DbValYTD
    oSheet.Cells(i, 7) = wRsDB!DbQtyYTD / oSheet.Cells(2, 4)
    oSheet.Cells(i, 8) = wRsDB!DbValPrevYTD
    oSheet.Cells(i, 9) = wRsDB!DbQtyPrevYTD / oSheet.Cells(2, 4)
    wRsDB.MoveNext
    wRow = wRow + 1
    i = i + 1
  Loop
      
  wRow = 1
  wSql = "Select * From DashBoard where DbTyp ='Shipment' "
  Set wRsDB = moCn.OpenRes(wSql)
  Do While Not (wRsDB.EOF Or wRsDB.BOF)
    If wRow = 1 Then
      i = i + 1
    Else
      oSheet.Rows(CStr(i) + ":" + CStr(i)).Insert
      oSheet.Range("B6" + ":I6").Copy
      oSheet.Range("B" + CStr(i) + ":I" + CStr(i)).Select
      oSheet.Paste
    End If
    oSheet.Cells(i, 2) = wRsDB!DbTyp
    oSheet.Cells(i, 3) = wRsDB!DbSubTyp
    oSheet.Cells(i, 4) = wRsDB!DbValToday
    oSheet.Cells(i, 5) = wRsDB!DbQtyToday / oSheet.Cells(2, 4)
    oSheet.Cells(i, 6) = wRsDB!DbValYTD
    oSheet.Cells(i, 7) = wRsDB!DbQtyYTD / oSheet.Cells(2, 4)
    oSheet.Cells(i, 8) = wRsDB!DbValPrevYTD
    oSheet.Cells(i, 9) = wRsDB!DbQtyPrevYTD / oSheet.Cells(2, 4)
    wRsDB.MoveNext
    i = i + 1
    wRow = wRow + 1
  Loop
    
  wRow = 1
  wSql = "Select * From DashBoard where DbTyp ='Pnd Order' "
  Set wRsDB = moCn.OpenRes(wSql)
  Do While Not (wRsDB.EOF Or wRsDB.BOF)
    If wRow = 1 Then
      i = i + 1
    Else
      oSheet.Rows(CStr(i) + ":" + CStr(i)).Insert
      oSheet.Range("B6" + ":I6").Copy
      oSheet.Range("B" + CStr(i) + ":I" + CStr(i)).Select
      oSheet.Paste
    End If
    oSheet.Cells(i, 2) = wRsDB!DbTyp
    oSheet.Cells(i, 3) = wRsDB!DbSubTyp
    oSheet.Cells(i, 4) = wRsDB!DbValToday
    oSheet.Cells(i, 5) = wRsDB!DbQtyToday / oSheet.Cells(2, 4)
    oSheet.Cells(i, 6) = wRsDB!DbValYTD
    oSheet.Cells(i, 7) = wRsDB!DbQtyYTD / oSheet.Cells(2, 4)
    oSheet.Cells(i, 8) = wRsDB!DbValPrevYTD
    oSheet.Cells(i, 9) = wRsDB!DbQtyPrevYTD / oSheet.Cells(2, 4)
    wRsDB.MoveNext
    i = i + 1
    wRow = wRow + 1
  Loop


  'Copying DashBoard report3A
  Dim wRowTot As Integer, wRowFr As Integer
  wRow = 1
  wSql = "Select * From DashBoard where DbTyp ='Stock3A' "
  Set wRsDB = moCn.OpenRes(wSql)
  Do While Not (wRsDB.EOF Or wRsDB.BOF)
    If wRow = 1 Then
      i = i + 5
      wRowFr = i
    Else
      oSheet.Rows(CStr(i) + ":" + CStr(i)).Insert
      oSheet.Range("B" + CStr(i - 1) + ":I" + CStr(i - 1)).Copy
      oSheet.Range("B" + CStr(i) + ":I" + CStr(i)).Select
      oSheet.Paste
    End If
    If wRsDB!DbSubTyp Like "X*" Then
      If wRowTot = 0 Then wRowTot = i
    End If
    oSheet.Cells(i, 2) = wRsDB!DbSubTyp
    oSheet.Cells(i, 3) = ""
    oSheet.Cells(i, 4) = wRsDB!DbValToday
    oSheet.Cells(i, 5) = wRsDB!DbQtyToday
    oSheet.Cells(i, 6) = wRsDB!DbValYTD
    oSheet.Cells(i, 7) = wRsDB!DbQtyYTD
    oSheet.Cells(i, 8) = wRsDB!DbValPrevYTD
    oSheet.Cells(i, 9) = wRsDB!DbQtyPrevYTD
    wRsDB.MoveNext
    i = i + 1
    wRow = wRow + 1
  Loop
  If wRsDB.RecCount > 0 Then
    oSheet.Rows(CStr(wRowTot) + ":" + CStr(wRowTot)).Insert
    oSheet.Cells(wRowTot, 2) = "Total"
    oSheet.Range("B" + CStr(wRowTot) + ":I" + CStr(wRowTot)).Font.Bold = True
    oSheet.Cells(wRowTot, 4).formula = "=Sum(D" + CStr(wRowFr) + ":D" + CStr(wRowTot - 1) + ")"
    oSheet.Cells(wRowTot, 5).formula = "=Sum(E" + CStr(wRowFr) + ":E" + CStr(wRowTot - 1) + ")"
    oSheet.Cells(wRowTot, 6).formula = "=Sum(F" + CStr(wRowFr) + ":F" + CStr(wRowTot - 1) + ")"
    oSheet.Cells(wRowTot, 7).formula = "=Sum(G" + CStr(wRowFr) + ":G" + CStr(wRowTot - 1) + ")"
    oSheet.Cells(wRowTot, 8).formula = "=Sum(H" + CStr(wRowFr) + ":H" + CStr(wRowTot - 1) + ")"
    oSheet.Cells(wRowTot, 9).formula = "=Sum(I" + CStr(wRowFr) + ":I" + CStr(wRowTot - 1) + ")"
    i = i + 1
  End If
  
  'Copying DashBoard report3B
  wRow = 1: wRowTot = 0: wRowFr = 0
  wSql = "Select * From DashBoard where DbTyp ='Stock3B' "
  Set wRsDB = moCn.OpenRes(wSql)
  Do While Not (wRsDB.EOF Or wRsDB.BOF)
    If wRow = 1 Then
      i = i + 5
      wRowFr = i
    Else
      oSheet.Rows(CStr(i) + ":" + CStr(i)).Insert
      oSheet.Range("B" + CStr(i - 1) + ":I" + CStr(i - 1)).Copy
      oSheet.Range("B" + CStr(i) + ":I" + CStr(i)).Select
      oSheet.Paste
    End If
    If wRsDB!DbSubTyp Like "X*" Then
      If wRowTot = 0 Then wRowTot = i
    End If
    oSheet.Cells(i, 2) = wRsDB!DbSubTyp
    oSheet.Cells(i, 3) = ""
    oSheet.Cells(i, 4) = wRsDB!DbValToday / oSheet.Cells(2, 4)
    oSheet.Cells(i, 5) = wRsDB!DbQtyToday / oSheet.Cells(2, 4)
    oSheet.Cells(i, 6) = wRsDB!DbValYTD / oSheet.Cells(2, 4)
    oSheet.Cells(i, 7) = wRsDB!DbQtyYTD / oSheet.Cells(2, 4)
    oSheet.Cells(i, 8) = wRsDB!DbValPrevYTD / oSheet.Cells(2, 4)
    oSheet.Cells(i, 9) = wRsDB!DbQtyPrevYTD / oSheet.Cells(2, 4)
    wRsDB.MoveNext
    i = i + 1
    wRow = wRow + 1
  Loop
  If wRsDB.RecCount > 0 Then
    oSheet.Rows(CStr(wRowTot) + ":" + CStr(wRowTot)).Insert
    oSheet.Cells(wRowTot, 2) = "Total"
    oSheet.Range("B" + CStr(wRowTot) + ":I" + CStr(wRowTot)).Font.Bold = True
    oSheet.Cells(wRowTot, 4).formula = "=Sum(D" + CStr(wRowFr) + ":D" + CStr(wRowTot - 1) + ")"
    oSheet.Cells(wRowTot, 5).formula = "=Sum(E" + CStr(wRowFr) + ":E" + CStr(wRowTot - 1) + ")"
    oSheet.Cells(wRowTot, 6).formula = "=Sum(F" + CStr(wRowFr) + ":F" + CStr(wRowTot - 1) + ")"
    oSheet.Cells(wRowTot, 7).formula = "=Sum(G" + CStr(wRowFr) + ":G" + CStr(wRowTot - 1) + ")"
    oSheet.Cells(wRowTot, 8).formula = "=Sum(H" + CStr(wRowFr) + ":H" + CStr(wRowTot - 1) + ")"
    oSheet.Cells(wRowTot, 9).formula = "=Sum(I" + CStr(wRowFr) + ":I" + CStr(wRowTot - 1) + ")"
    i = i + 1
  End If
  
  
  
  'Copying DashBoard report4
  wRow = 1
  wSql = "Select * From DashBoard where DbTyp ='FG (Quantity)'"
  Set wRsDB = moCn.OpenRes(wSql)
  Do While Not (wRsDB.EOF Or wRsDB.BOF)
    i = i + 5
    oSheet.Cells(i, 2) = wRsDB!DbTyp
    oSheet.Cells(i, 3) = wRsDB!DbSubTyp
    oSheet.Cells(i, 4) = wRsDB!DbValToday / oSheet.Cells(2, 5)
    oSheet.Cells(i, 5) = wRsDB!DbQtyToday / oSheet.Cells(2, 4)
    oSheet.Cells(i, 6) = wRsDB!DbValYTD / oSheet.Cells(2, 5)
    oSheet.Cells(i, 7) = wRsDB!DbQtyYTD / oSheet.Cells(2, 4)
    oSheet.Cells(i, 8) = wRsDB!DbValPrevYTD / oSheet.Cells(2, 5)
    oSheet.Cells(i, 9) = wRsDB!DbQtyPrevYTD / oSheet.Cells(2, 4)
    wRsDB.MoveNext
  Loop

  i = i + 1
  wSql = "Select * From DashBoard where DbTyp ='FG (Gold & Dia Wght)'"
  Set wRsDB = moCn.OpenRes(wSql)
  Do While Not (wRsDB.EOF Or wRsDB.BOF)
    oSheet.Cells(i, 2) = wRsDB!DbTyp
    oSheet.Cells(i, 3) = wRsDB!DbSubTyp
    oSheet.Cells(i, 4) = wRsDB!DbValToday
    oSheet.Cells(i, 5) = wRsDB!DbQtyToday
    oSheet.Cells(i, 6) = wRsDB!DbValYTD
    oSheet.Cells(i, 7) = wRsDB!DbQtyYTD
    oSheet.Cells(i, 8) = wRsDB!DbValPrevYTD
    oSheet.Cells(i, 9) = wRsDB!DbQtyPrevYTD
    wRsDB.MoveNext
  Loop

  i = i + 1
  wSql = "Select * From DashBoard where DbTyp ='Production Loss (Gold)'"
  Set wRsDB = moCn.OpenRes(wSql)
  Do While Not (wRsDB.EOF Or wRsDB.BOF)
    oSheet.Cells(i, 2) = wRsDB!DbTyp
    oSheet.Cells(i, 3) = wRsDB!DbSubTyp
    oSheet.Cells(i, 4) = wRsDB!DbValToday
    oSheet.Cells(i, 5).formula = "=ROUND(D" + CStr(i) + "/D" + CStr(i - 1) + "*100,2)"
    oSheet.Cells(i, 6) = wRsDB!DbValYTD
    oSheet.Cells(i, 7).formula = "=ROUND(F" + CStr(i) + "/F" + CStr(i - 1) + "*100,2)"
    oSheet.Cells(i, 8) = wRsDB!DbValPrevYTD
    oSheet.Cells(i, 9).formula = "=ROUND(H" + CStr(i) + "/H" + CStr(i - 1) + "*100,2)"
    wRsDB.MoveNext
  Loop

  i = i + 1
  wSql = "Select * From DashBoard where DbTyp ='Factory Loss (Gold)'"
  Set wRsDB = moCn.OpenRes(wSql)
  Do While Not (wRsDB.EOF Or wRsDB.BOF)
    oSheet.Cells(i, 2) = wRsDB!DbTyp
    oSheet.Cells(i, 3) = wRsDB!DbSubTyp
    oSheet.Cells(i, 4) = wRsDB!DbValToday
    oSheet.Cells(i, 5).formula = "=ROUND(D" + CStr(i) + "/D" + CStr(i - 2) + "*100,2)"
    oSheet.Cells(i, 6) = wRsDB!DbValYTD
    oSheet.Cells(i, 7).formula = "=ROUND(F" + CStr(i) + "/F" + CStr(i - 2) + "*100,2)"
    oSheet.Cells(i, 8) = wRsDB!DbValPrevYTD
    oSheet.Cells(i, 9).formula = "=ROUND(H" + CStr(i) + "/H" + CStr(i - 2) + "*100,2)"
    wRsDB.MoveNext
  Loop

  i = i + 1
  wSql = "Select * From DashBoard where DbTyp ='Broken Dia'"
  Set wRsDB = moCn.OpenRes(wSql)
  Do While Not (wRsDB.EOF Or wRsDB.BOF)
    oSheet.Cells(i, 2) = wRsDB!DbTyp
    oSheet.Cells(i, 3) = wRsDB!DbSubTyp
    oSheet.Cells(i, 4) = wRsDB!DbValToday
    oSheet.Cells(i, 5).formula = "=ROUND(D" + CStr(i) + "/E" + CStr(i - 3) + "*100,2)"
    oSheet.Cells(i, 6) = wRsDB!DbValYTD
    oSheet.Cells(i, 7).formula = "=ROUND(F" + CStr(i) + "/G" + CStr(i - 3) + "*100,2)"
    oSheet.Cells(i, 8) = wRsDB!DbValPrevYTD
    oSheet.Cells(i, 9).formula = "=ROUND(H" + CStr(i) + "/I" + CStr(i - 3) + "*100,2)"
    wRsDB.MoveNext
  Loop

  i = i + 2
  wSql = "Select * From DashBoard where DbTyp ='Missing Dia'"
  Set wRsDB = moCn.OpenRes(wSql)
  Do While Not (wRsDB.EOF Or wRsDB.BOF)
    oSheet.Cells(i, 2) = wRsDB!DbTyp
    oSheet.Cells(i, 3) = wRsDB!DbSubTyp
    oSheet.Cells(i, 4) = wRsDB!DbValToday
    oSheet.Cells(i, 5).formula = "=ROUND(D" + CStr(i) + "/E" + CStr(i - 5) + "*100,2)"
    oSheet.Cells(i, 6) = wRsDB!DbValYTD
    oSheet.Cells(i, 7).formula = "=ROUND(F" + CStr(i) + "/G" + CStr(i - 5) + "*100,2)"
    oSheet.Cells(i, 8) = wRsDB!DbValPrevYTD
    oSheet.Cells(i, 9).formula = "=ROUND(H" + CStr(i) + "/I" + CStr(i - 5) + "*100,2)"
    wRsDB.MoveNext
  Loop
    
  Set wRsDB = Nothing
  Set wRsDBPrm = Nothing
  Set oSheet = oBook.worksheets(1)
  DispMsg "Updated Excel Successfully !!", etInfo
  MsgBox "Updated Excel Successfully !!"
  oExcel.ActiveWorkbook.Close True
  oExcel.QUIT

  oExcel.DisplayAlerts = True
  Set oExcel = Nothing
  Set oBook = Nothing
  Set oSheet = Nothing
  Exit Sub
Err_Hndlr:
  If Err.Number = 32755 Then
    wErrMsg = "Operation Cancelled"
    oExcel.ActiveWorkbook.Close
    oExcel.QUIT
  ElseIf Err.Number = 429 Then
    DispMsg "MS Excel is either Corrupt Or Missing on this PC. Please Install/Re-Install MS Excel ! ", etError
    
  ElseIf Err.Number = 70 Then
    DispMsg "Cannot Access Template File. Close If Already Open and Continue.", etError
  Else
    wErrMsg = Err.Description
    DispMsg wErrMsg, etError
  End If


  
End Sub

Private Function SelectFileName(ByVal wExtn As String, ByRef ChosenFileName) As Boolean
  'checking selected excel file
  Do While True
    On Error GoTo ErrH
    Cd1.ShowSave
    If Len(Cd1.FileName) > 4 And ((UCase(Right(Cd1.FileName, 4)) <> UCase(wExtn) And UCase(Right(Cd1.FileName, 5)) <> UCase(wExtn))) Then
        MsgBox ("File Extension Is Wrong. Please re-select")
    Else
      If (CheckAndDeleteFile(Cd1.FileName)) Then
        ChosenFileName = Cd1.FileName
        SelectFileName = True
        Exit Function
      End If
    End If
  Loop
  Exit Function
ErrH:
  If Err.Number = cdlCancel Then
    SelectFileName = False
  End If
End Function
Private Function CheckAndDeleteFile(FileName As String) As Boolean
  If Dir(FileName, vbNormal) <> "" Then
    On Error GoTo errDeletingFile
    Kill FileName
    CheckAndDeleteFile = True
    Exit Function
errDeletingFile:
    MsgBox ("System is unable to delete file " + FileName + ". Please delete file or choose a different path or name")
    CheckAndDeleteFile = False
  Else
    CheckAndDeleteFile = True
  End If
End Function

