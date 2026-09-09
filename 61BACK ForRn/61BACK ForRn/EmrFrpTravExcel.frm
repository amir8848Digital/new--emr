VERSION 5.00
Object = "{8767A745-088E-4CA6-8594-073D6D2DE57A}#9.2#0"; "crviewer9.dll"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "tabctl32.ocx"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Object = "{AB2CC32A-E652-42A2-B341-26ECE02F6386}#2.1#0"; "mwfctlforspeed9.ocx"
Begin VB.Form EmrFrpTravExcel 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Order Printing"
   ClientHeight    =   10185
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   15270
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MDIChild        =   -1  'True
   ScaleHeight     =   10185
   ScaleWidth      =   15270
   WindowState     =   2  'Maximized
   Begin MSComDlg.CommonDialog CD1 
      Left            =   480
      Top             =   3900
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin MwfCtl.MWCTL_HLP MWCTL_HLP1 
      Height          =   495
      Left            =   10710
      TabIndex        =   8
      TabStop         =   0   'False
      Top             =   9570
      Visible         =   0   'False
      Width           =   495
      _ExtentX        =   873
      _ExtentY        =   873
   End
   Begin MwfCtl.MWCTL_DST DST_REP 
      Height          =   570
      Left            =   4440
      TabIndex        =   7
      TabStop         =   0   'False
      Top             =   9630
      Width           =   6360
      _ExtentX        =   11218
      _ExtentY        =   1005
   End
   Begin VB.Frame fra_keys 
      BackColor       =   &H00C0C0C0&
      BorderStyle     =   0  'None
      Height          =   465
      Left            =   4605
      TabIndex        =   17
      Top             =   9645
      Width           =   3555
      Begin MwfCtl.MWCTL_MED aktxt 
         Height          =   345
         Index           =   0
         Left            =   210
         TabIndex        =   18
         Top             =   135
         Visible         =   0   'False
         Width           =   1455
         _ExtentX        =   2566
         _ExtentY        =   609
         MaxLength       =   0
         DataField       =   "UOUSRCD"
         IdName          =   "UOUSRCD"
      End
      Begin MwfCtl.MWCTL_MED aktxt 
         Height          =   345
         Index           =   1
         Left            =   1860
         TabIndex        =   19
         Top             =   150
         Visible         =   0   'False
         Width           =   1455
         _ExtentX        =   2566
         _ExtentY        =   609
         MaxLength       =   0
         DataField       =   "UOMNUCD"
         IdName          =   "UOMNUCD"
      End
   End
   Begin VB.Frame FRA_REPOBJ 
      BorderStyle     =   0  'None
      Height          =   495
      Left            =   5820
      TabIndex        =   16
      Top             =   9600
      Visible         =   0   'False
      Width           =   855
      Begin CRVIEWER9LibCtl.CRViewer9 CRV_REP 
         Height          =   495
         Left            =   480
         TabIndex        =   13
         Top             =   -135
         Visible         =   0   'False
         Width           =   495
         lastProp        =   500
         _cx             =   5080
         _cy             =   5080
         DisplayGroupTree=   -1  'True
         DisplayToolbar  =   -1  'True
         EnableGroupTree =   -1  'True
         EnableNavigationControls=   -1  'True
         EnableStopButton=   -1  'True
         EnablePrintButton=   -1  'True
         EnableZoomControl=   -1  'True
         EnableCloseButton=   -1  'True
         EnableProgressControl=   -1  'True
         EnableSearchControl=   -1  'True
         EnableRefreshButton=   0   'False
         EnableDrillDown =   -1  'True
         EnableAnimationControl=   -1  'True
         EnableSelectExpertButton=   0   'False
         EnableToolbar   =   -1  'True
         DisplayBorder   =   -1  'True
         DisplayTabs     =   -1  'True
         DisplayBackgroundEdge=   -1  'True
         SelectionFormula=   ""
         EnablePopupMenu =   -1  'True
         EnableExportButton=   0   'False
         EnableSearchExpertButton=   0   'False
         EnableHelpButton=   0   'False
         LaunchHTTPHyperlinksInNewBrowser=   -1  'True
      End
   End
   Begin MwfCtl.MWCTL_ADC adc 
      Height          =   555
      Left            =   60
      TabIndex        =   9
      Top             =   9660
      Width           =   15120
      _ExtentX        =   26670
      _ExtentY        =   979
      AllowDelete     =   0   'False
      Automate        =   -1  'True
      EntType         =   1
      BtnWidth        =   885
      BtnHeight       =   435
      BtnMaskColor    =   8454143
      FontSz          =   8.25
   End
   Begin VB.Frame fra_nkeys 
      BorderStyle     =   0  'None
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   9780
      Left            =   0
      TabIndex        =   12
      Top             =   45
      Width           =   15255
      Begin TabDlg.SSTab TAB_REP 
         Height          =   9465
         Left            =   15
         TabIndex        =   11
         Top             =   45
         Width           =   15195
         _ExtentX        =   26802
         _ExtentY        =   16695
         _Version        =   393216
         Tabs            =   2
         TabsPerRow      =   2
         TabHeight       =   520
         TabCaption(0)   =   "&Data Range"
         Tab(0).ControlEnabled=   -1  'True
         Tab(0).Control(0)=   "fra_tabrep(0)"
         Tab(0).Control(0).Enabled=   0   'False
         Tab(0).ControlCount=   1
         TabCaption(1)   =   "&Group Sort"
         Tab(1).ControlEnabled=   0   'False
         Tab(1).Control(0)=   "fra_tabrep(1)"
         Tab(1).ControlCount=   1
         Begin VB.Frame fra_tabrep 
            BorderStyle     =   0  'None
            Enabled         =   0   'False
            Height          =   9070
            Index           =   1
            Left            =   -74940
            TabIndex        =   15
            Top             =   360
            Width           =   15105
            Begin MwfCtl.MWCTL_GRP GRP_REP 
               Height          =   360
               Left            =   240
               TabIndex        =   10
               Top             =   360
               Width           =   10335
               _ExtentX        =   18230
               _ExtentY        =   635
            End
         End
         Begin VB.Frame fra_tabrep 
            BorderStyle     =   0  'None
            Height          =   9070
            Index           =   0
            Left            =   60
            TabIndex        =   14
            Top             =   360
            Width           =   15105
            Begin VB.TextBox TxtTitle 
               Height          =   300
               Left            =   1900
               MaxLength       =   45
               TabIndex        =   6
               ToolTipText     =   "Enter Title"
               Top             =   1680
               Width           =   7095
            End
            Begin MwfCtl.MWCTL_MED ATXT 
               Height          =   285
               Index           =   20
               Left            =   2685
               TabIndex        =   2
               ToolTipText     =   "Enter From Quotation Year"
               Top             =   825
               Width           =   345
               _ExtentX        =   609
               _ExtentY        =   503
               MaxLength       =   2
               DataType        =   4
               DataField       =   "UOOMYYFR"
               IdName          =   "UOOMYYFR"
               CmpStr          =   "OdYy ="
            End
            Begin MwfCtl.MWCTL_MED ATXT 
               Height          =   285
               Index           =   43
               Left            =   1905
               TabIndex        =   1
               ToolTipText     =   "Enter From Quotation Tc"
               Top             =   825
               Width           =   705
               _ExtentX        =   1244
               _ExtentY        =   503
               MaxLength       =   3
               DataType        =   4
               DataField       =   "UOOMTCFR"
               IdName          =   "UOOMTCFR"
               CmpStr          =   "OdTc ="
            End
            Begin MwfCtl.MWCTL_MED ATXT 
               Height          =   285
               Index           =   45
               Left            =   3105
               TabIndex        =   3
               ToolTipText     =   "Enter From Quotation Character"
               Top             =   825
               Width           =   705
               _ExtentX        =   1244
               _ExtentY        =   503
               MaxLength       =   3
               DataType        =   4
               DataField       =   "UOOMCHRFR"
               IdName          =   "UOOMCHRFR"
               CmpStr          =   "OdChr ="
            End
            Begin MwfCtl.MWCTL_MED ATXT 
               Height          =   285
               Index           =   26
               Left            =   1905
               TabIndex        =   4
               ToolTipText     =   "Enter From Quotation Number"
               Top             =   1110
               Width           =   825
               _ExtentX        =   1455
               _ExtentY        =   503
               MaxLength       =   0
               DataType        =   1
               DataField       =   "UOOMNOFR"
               IdName          =   "UOOMNOFR"
               CmpStr          =   "OdNo ="
            End
            Begin MwfCtl.MWCTL_MED ATXT 
               Height          =   285
               Index           =   8
               Left            =   1905
               TabIndex        =   5
               ToolTipText     =   "Enter Diamond Grade Selection"
               Top             =   1395
               Width           =   8070
               _ExtentX        =   14235
               _ExtentY        =   503
               MaxLength       =   255
               DataType        =   4
               DataField       =   "UOPSCDSEL"
               IdName          =   "UOPSCDSEL"
            End
            Begin MwfCtl.MWCTL_MED ATXT 
               Height          =   285
               Index           =   35
               Left            =   1905
               TabIndex        =   0
               ToolTipText     =   "Enter Company Code"
               Top             =   540
               Width           =   705
               _ExtentX        =   1244
               _ExtentY        =   503
               Enabled         =   0   'False
               MaxLength       =   3
               DataType        =   4
               DataField       =   "UOCOCDFR"
               IdName          =   "UOCOCDFR"
               CmpStr          =   "OdCoCd = "
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
               Index           =   27
               Left            =   2625
               TabIndex        =   29
               Top             =   825
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
               Index           =   25
               Left            =   3045
               TabIndex        =   28
               Top             =   825
               Width           =   105
            End
            Begin VB.Label Label1 
               BackStyle       =   0  'Transparent
               Caption         =   "Diamond Grade"
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
               Index           =   1
               Left            =   75
               TabIndex        =   27
               ToolTipText     =   "Location"
               Top             =   1395
               Width           =   1875
            End
            Begin VB.Label LblGrpCombo 
               BackStyle       =   0  'Transparent
               Caption         =   "Title"
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
               Index           =   0
               Left            =   75
               TabIndex        =   26
               ToolTipText     =   "Location"
               Top             =   1695
               Width           =   1125
            End
            Begin VB.Label ALBL 
               BackStyle       =   0  'Transparent
               Caption         =   "Selection"
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
               Index           =   3
               Left            =   6765
               TabIndex        =   25
               Top             =   120
               Width           =   1575
            End
            Begin VB.Label ALBL 
               BackStyle       =   0  'Transparent
               Caption         =   "Company Cd"
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
               Index           =   15
               Left            =   75
               TabIndex        =   24
               Top             =   540
               Width           =   1845
            End
            Begin VB.Label ALBL 
               BackStyle       =   0  'Transparent
               Caption         =   "Quot. Tc/Yy/Chr"
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
               Index           =   9
               Left            =   75
               TabIndex        =   23
               Top             =   825
               Width           =   1845
            End
            Begin VB.Label ALBL 
               BackStyle       =   0  'Transparent
               Caption         =   "Quot. No"
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
               Index           =   10
               Left            =   75
               TabIndex        =   22
               Top             =   1110
               Width           =   1845
            End
            Begin VB.Label ALBL 
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
               ForeColor       =   &H00000000&
               Height          =   285
               Index           =   2
               Left            =   4290
               TabIndex        =   21
               Top             =   120
               Width           =   1275
            End
            Begin VB.Label ALBL 
               BackStyle       =   0  'Transparent
               Caption         =   "From"
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
               Index           =   1
               Left            =   1905
               TabIndex        =   20
               Top             =   120
               Width           =   1335
            End
         End
      End
   End
End
Attribute VB_Name = "EmrFrpTravExcel"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim moCn As MwfLib.MDOConnection   '*** (Bef speed) bef 24/11/06
Dim wErrMsg As String
Dim Rep As New EmrRepDummy
Dim wDiaGrd(15) As String
Dim oExcel, oBook, oSheet
Private Sub SetGroupSort()
  '
End Sub
Private Sub adc_setreprecsource()
On Error GoTo RepErr

Set Rep = Nothing
Set adc.RepSource = Rep
  
  wErrMsg = ""
  SetTravelExcel
  
  adc.RepRecSource = "Select 'x' from Param where 1= 2"

  If wErrMsg <> "" Then
    DispMsg wErrMsg, etError
  End If

Exit Sub
RepErr:
  'Call ChkSelScope(Err.Number, Err.Description)
  DispMsg wErrMsg, etError

End Sub

Private Sub SetTravelExcel() '(ByRef wFileAlias As String, Optional ByRef wErrMsg As String = "")
  
  On Error GoTo Err_Hndlr
  
  Dim wCnd As String, wFldLst As String, ws_SqlStr As String
  Dim wRsOrdMst As MDORowSet, wRsDsgDet As MDORowSet
  
  wCnd = IIF(adc.RepCond <> "", " where " + adc.RepCond, "")
  wFldLst = FldLst(15)
  
  Set wRsOrdMst = moCn.OpenRes("Select OmCmCd, OmLmgSal, OmLmpSal, OmLmsSal from OrdMst where OmCoCd='" + adc("UoCoCdFr") + "' and " + _
              "OmTc='" + adc("UoOmTcFr") + "' and OmYy='" + adc("UoOmYyFr") + "' and " + _
              "OmChr='" + adc("UoOmChrFr") + "' and OmNo='" + CStr(adc("UoOmNoFr")) + "'")
              
  If (wRsOrdMst.EOF And wRsOrdMst.BOF) Then
    wErrMsg = "Quotation Not Found"
    Exit Sub
  End If

  ws_SqlStr = " Select OdSr as qOdSr,PSCd as qGrdCd,OdDmCd as qOdDmCd, " + _
              " (Select Sum((case when dbo.MwFn_GetRmRate('" + wRsOrdMst!OmCmCd + "',GrRmCd,OrLn1,OrLn2,(Case when OrQty=0 Then 0 Else OrWt/OrQty End ),'S')=0 " + _
              "   then -99999 " + _
              " Else " + _
              "  dbo.MwFn_GetRmRate('" + wRsOrdMst!OmCmCd + "',GrRmCd,OrLn1,OrLn2,(Case when OrQty=0 Then 0 Else OrWt/OrQty End ),'S') " + _
              " End)* (Case when RmNew.RmQW='Q' then OrQty else OrWt End)) DiaVal " + _
              "    From OrdRm  join RmMst RMOld on RmOld.RmCd = OrRmCd " + _
              "    left join Grade Gr1 " + _
              "    on GrCmCd = '" + wRsOrdMst!OmCmCd + "' " + _
              "       and Gr1.GrRmCtg = OrRmCtg and Gr1.GrFrLn <= OrLn1 and Gr1.GrToLn >= OrLn1 and Gr1.GrRmId = RmId and Gr1.GrCd = PSCd " + _
              "    left join RmMst RmNew on RmNew.RmCd = Gr1.GrRmCd " + _
              "    where OrCoCd=OdCoCd and OrTc=OdTc and OrYy=OdYy and OrChr=OdChr and OrNo=OdNo and OrSr=OdSr and OrRmCtg='D' " + _
              "     Group By OrCoCd,OrTc,OrYy,OrChr,OrNo,OrSr) DiaVal Into #Tmp " + _
              " From OrdDsg, Param " + _
              wCnd + " and PTyp='GRDCD' " + _
              "    and PMCd='DIA' and '," + Replace(adc("UoPSCdSel"), "'", "") + ",' like '%,'+PSCd +',%' " + _
              " Order By OdSr "

    ' **** Manali 3.6.1 - 'L' added
    ' **** Manali 3.6.2 - colstn func added
    ws_SqlStr = ws_SqlStr + "select OdSr,max(OdDmCd) OdDmCd " + wFldLst + _
              "     ,max(qDiaWt) as qDiaWt,max(qDiaQty) as qDiaQty,max(qCsWt) as qCsWt,max(qCsQty) as qCsQty, " + _
              "     max(qMetWt) as qMetWt,max(qMetVal) as qMetVal,max(qCsVal) as qCsVal,max(qFndVal) as qFndVal, " + _
              "     max(qSetVal + qLabVal) as qLabVal, max(OdDmCol) as OdDmCol, " + _
              "     dbo.MwFn_GetColStnStr(max(OdCoCd), max(OdTc), max(OdYy), max(OdChr), max(OdNo), OdSr,'','N')  as qCs " + _
              " From OrdDsg " + _
              " Left Outer Join #Tmp On OdSr=qOdSr " + _
              " Left Outer Join (Select OrCoCd,OrTc,OrYy,OrChr,OrNo,OrSr, " + _
              "    Sum((Case when OrRmCtg in ('D') then OrWt else 0 End)) as qDiaWt, " + _
              "    Sum((Case when OrRmCtg in ('D') then OrQty else 0 End)) as qDiaQty, " + _
              "    Sum((Case when OrRmCtg in ('C') then OrWt else 0 End)) as qCsWt, " + _
              "    Sum((Case when OrRmCtg in ('C') then OrQty else 0 End)) as qCsQty, " + _
              "    Sum((Case when OrRmCtg in ('G','P','S','L','X') then OrWt else 0 End)) as qMetWt, " + _
              "    Sum((Case when OrRmCtg in ('G','P','S','L' /*'X'*/) then OrSalVal else 0 End)) as qMetVal, " + _
              "    Sum((Case when OrRmCtg = 'X' then OrSalVal else 0 End)) as qFndVal, " + _
              "    Sum((Case when OrRmCtg ='C' then OrSalVal else 0 End)) as qCsVal, " + _
              "    Sum(OrSetSalVal) As qSetVal " + _
              "    From OrdRm Group By OrCoCd,OrTc,OrYy,OrChr,OrNo,OrSr) ORm " + _
              "         On OrCoCd=OdCoCd and OrTc=OdTc and OrYy=OdYy and OrChr=OdChr and OrNo=OdNo and OrSr=OdSr " + _
              "    Left Outer Join (Select OlCoCd,OlTc,OlYy,OlChr,OlNo,OlSr,Sum(OlSalVal) as qLabVal " + _
              "       From OrdLab  Group By OlCoCd,OlTc,OlYy,OlChr,OlNo,OlSr) Ol " + _
              " On OlCoCd=OdCoCd and OlTc=OdTc and OlYy=OdYy and OlChr=OdChr and OlNo=OdNo and OlSr=OdSr " + _
              wCnd + _
              " Group By OdSr "

  Set wRsDsgDet = moCn.OpenRes(ws_SqlStr)
  
  '****** Sachin - Before 3.3.0 - Removed as gives Toolbar Error If Excel Reference is Given
  'Dim oExcel As New Excel.Application
  '****** Sachin - Before 3.3.0 - Removed as gives Toolbar Error If Excel Reference is Given
  
  ' ***** Manali 3.9.0 - Excel For 2007 n next
  Dim wExcVer As String, wExtn As String
  Set oExcel = GetExcelObj()                'CreateObject("Excel.Application")
  Dim sFile, sFilePath As String

  ' ***** Manali 3.9.0 - Excel 2007
  wExcVer = oExcel.version
  wExtn = IIF(wExcVer >= "12.0", ".xlsx", ".xls")
  
  sFile = gs_CoCd + "-" + gs_UsrCd + "-" + Format(Now, DtFullStr()) + wExtn   '6.1
  ' ***** Manali 3.9.0 - Excel 2007
  
  
  'Bef 3.9.0
  'sFile = gs_CoCd + "-" + gs_UsrCd + "-" + Format(Now, "ddMMyyHHmmss") + ".xls"
  'sFilePath = App.Path + "\Output\" + sFile
  
  ' ***** Manali 3.9.0 - Output folder Path From Param
  Dim wOutputPath As String
  
  wOutputPath = moCn.GetFldVal("Select PDesc225 from Param Where PTyp='OUTPUT' and PMCd='OUTPUT' ")
  sFilePath = wOutputPath + "\" + sFile
  ' ***** Manali 3.9.0 - Output folder Path From Param
  
  If Dir(sFilePath, vbNormal) <> "" Then
    Kill sFilePath
  End If
  
  ' ***** Manali 3.9.0 - Excel 2007
  FileCopy App.Path + "\Template\TravelExcel" + wExtn, sFilePath
  ' Bef  3.9.0
  'FileCopy App.Path + "\Template\TravelExcel.xls", sFilePath
  
  'Set oBook = oExcel.Workbooks.Open(sFilePath)
  '4.1.0.0
  Set oBook = oExcel.Workbooks.Open(sFilePath, Password:=ctXCelPwd, WriteResPassword:=ctXCelWritePwd)

  Set oSheet = oBook.worksheets(1)

  oExcel.Visible = False: oExcel.DisplayAlerts = False

  oSheet.Cells(1, 1) = TxtTitle.text
  oSheet.Cells(2, 3) = wRsOrdMst!OmLmgSal
  oSheet.Cells(2, 5) = wRsOrdMst!OmLmpSal
  oSheet.Cells(2, 7) = wRsOrdMst!OmLmsSal
  Dim wi_Ctr, wi_Indx As Integer
  wi_Ctr = 4
  For wi_Indx = LBound(wDiaGrd) To UBound(wDiaGrd) - 1
    oSheet.Cells(wi_Ctr, 8 + wi_Indx) = wDiaGrd(wi_Indx)
  Next wi_Indx
  wi_Ctr = 5
  Dim wd_OthVal As Double
  With wRsDsgDet
    Do While Not (.EOF Or .BOF)
      wd_OthVal = !qMetVal + !qFndVal + !qCsVal + !qLabVal
      oSheet.Cells(wi_Ctr, 1) = !OdDmCd
      oSheet.Cells(wi_Ctr, 2) = IIF(!qMetWt <> 0, !qMetWt, "N/A")
      oSheet.Cells(wi_Ctr, 3) = !OdDmCol
      oSheet.Cells(wi_Ctr, 4) = IIF(!qDiaQty <> 0, !qDiaQty, "N/A")
      oSheet.Cells(wi_Ctr, 5) = IIF(!qDiaWt <> 0, !qDiaWt, "N/A")
      oSheet.Cells(wi_Ctr, 6) = IIF(!qCsQty <> 0, !qCsQty, "N/A")
      oSheet.Cells(wi_Ctr, 7) = IIF(!qCsWt <> 0, !qCsWt, "N/A")
      oSheet.Cells(wi_Ctr, 8) = IIF(!qGdVal1 <> 0, !qGdVal1 + wd_OthVal, "N/A")
      oSheet.Cells(wi_Ctr, 9) = IIF(!qGdVal2 <> 0, !qGdVal2 + wd_OthVal, "N/A")
      oSheet.Cells(wi_Ctr, 10) = IIF(!qGdVal3 <> 0, !qGdVal3 + wd_OthVal, "N/A")
      oSheet.Cells(wi_Ctr, 11) = IIF(!qGdVal4 <> 0, !qGdVal4 + wd_OthVal, "N/A")
      oSheet.Cells(wi_Ctr, 12) = IIF(!qGdVal5 <> 0, !qGdVal5 + wd_OthVal, "N/A")
      oSheet.Cells(wi_Ctr, 13) = IIF(!qGdVal6 <> 0, !qGdVal6 + wd_OthVal, "N/A")
      oSheet.Cells(wi_Ctr, 14) = IIF(!qGdVal7 <> 0, !qGdVal7 + wd_OthVal, "N/A")
      oSheet.Cells(wi_Ctr, 15) = IIF(!qGdVal8 <> 0, !qGdVal8 + wd_OthVal, "N/A")
      oSheet.Cells(wi_Ctr, 16) = IIF(!qGdVal9 <> 0, !qGdVal9 + wd_OthVal, "N/A")
      oSheet.Cells(wi_Ctr, 17) = IIF(!qGdVal10 <> 0, !qGdVal10 + wd_OthVal, "N/A")
      oSheet.Cells(wi_Ctr, 18) = IIF(!qGdVal11 <> 0, !qGdVal11 + wd_OthVal, "N/A")
      oSheet.Cells(wi_Ctr, 19) = IIF(!qGdVal12 <> 0, !qGdVal12 + wd_OthVal, "N/A")
      oSheet.Cells(wi_Ctr, 20) = IIF(!qGdVal13 <> 0, !qGdVal13 + wd_OthVal, "N/A")
      oSheet.Cells(wi_Ctr, 21) = IIF(!qGdVal14 <> 0, !qGdVal14 + wd_OthVal, "N/A")
      oSheet.Cells(wi_Ctr, 22) = IIF(!qGdVal15 <> 0, !qGdVal15 + wd_OthVal, "N/A")
      oSheet.Cells(wi_Ctr, 23) = IIF(!qCsVal <> 0, !qCsVal, "N/A")
      oSheet.Cells(wi_Ctr, 24) = IIF(!qMetVal <> 0, !qMetVal, "N/A")
      oSheet.Cells(wi_Ctr, 25) = IIF(!qLabVal <> 0, !qLabVal, "N/A")
      oSheet.Cells(wi_Ctr, 26) = IIF(!qFndVal <> 0, !qFndVal, "N/A")
      oSheet.Cells(wi_Ctr, 27) = !qCs
      wi_Ctr = wi_Ctr + 1
      wRsDsgDet.MoveNext
    Loop
  End With
  Set wRsDsgDet = Nothing
  Set wRsOrdMst = Nothing
  Cd1.DialogTitle = "Save Travel Excel File As"
  ' ***** Manali 3.9.0 - Excel For 2007 n next
  Cd1.FileName = adc("UoOmTcFr") + "-" + adc("UoOmYyFr") + "-" + adc("UoOmChrFr") + "-" + CStr(adc("UoOmNoFr")) + wExtn
  Cd1.Filter = "Excel Files (*" + wExtn + ") |*" + wExtn
  ' Bef 3.9.0
  ''CD1.FileName = adc("UoOmTcFr") + "-" + adc("UoOmYyFr") + "-" + adc("UoOmChrFr") + "-" + CStr(adc("UoOmNoFr")) + ".xls"
  ''CD1.Filter = "Excel Files (*.xls) |*.xls"
  'CD1.FilterIndex = 1
  Cd1.Flags = cdlOFNOverwritePrompt Or cdlOFNLongNames Or cdlOFNPathMustExist Or cdlOFNExtensionDifferent
  'CD1.DefaultExt = "xls"
  
  Cd1.CancelError = True
  Cd1.ShowSave

  ' ***** Manali 3.9.0 - Excel For 2007 n next
  If Len(Cd1.FileName) > 5 Then
    If UCase(Right(Cd1.FileName, 4)) = UCase(wExtn) Or UCase(Right(Cd1.FileName, 5)) = UCase(wExtn) Then
    ' Bef  3.9.0
    'If UCase(Right(CD1.FileName, 4)) = UCase(".xls") Then
      'oExcel.ActiveWorkbook.SaveAs Cd1.FileName
      oExcel.ActiveWorkbook.SaveAs Cd1.FileName, , "", ""       '4.1.0.0 (Pwd Removed)
    Else
      wErrMsg = "File Extension Is Wrong": Exit Sub
    End If
  End If
  
  DispMsg "Excel Generated Successfully !!", etInfo
  oExcel.ActiveWorkbook.Close True
  oExcel.QUIT
  
Exit Sub
Err_Hndlr:
  If Err.Number = 32755 Then
    wErrMsg = "Operation Cancelled"
    oExcel.ActiveWorkbook.Close
    oExcel.QUIT
  Else
    wErrMsg = Err.Description
  End If
  Err.Clear

End Sub

Private Function FldLst(ByVal mCtr As Integer) As String
  Dim ws_Seq As String, ws_str As String, wi_cPos As Integer, wi_Ctr As Integer, wi_Pos As Integer
  Dim ws_fldlst As String, wReqStr As String
  
  ws_Seq = adc("UoPSCdSel")
  ws_str = "," + ws_Seq + ","
  
  wi_cPos = 1
  wi_Ctr = 1
  Do While wi_Ctr <= mCtr
    wi_Pos = InStr(wi_cPos + 1, ws_str, ",")
    If wi_Pos = 0 Then
      wDiaGrd(wi_Ctr - 1) = ""
      ws_fldlst = ws_fldlst + ",0 as qGdVal" + CStr(wi_Ctr)
    Else
      wReqStr = Mid(ws_str, wi_cPos + 1, wi_Pos - wi_cPos - 1)
      wDiaGrd(wi_Ctr - 1) = Replace(wReqStr, "'", "")
      wi_cPos = wi_Pos
      'ws_fldlst = ws_fldlst + ",Sum(Case When IsNull(BLoc, '')='" + wReqStr + "' Then " + _
                "IsNull(BQty, 0) Else 0 End) AS qQty" + CStr(wi_Ctr)
      ws_fldlst = ws_fldlst + ",max(case when qGrdCd=" + wReqStr + " and DiaVal>0 Then DiaVal else 0 End) as qGdVal" + CStr(wi_Ctr)
    End If
    wi_Ctr = wi_Ctr + 1
  Loop

  FldLst = ws_fldlst
End Function

'''' std code not to be changed
Private Sub Form_GotFocus()
  '*** Jenny Color (For Setting System Colors)
  SetMDISysColorsAfterLogin
  '*** Jenny Color (For Setting System Colors)
End Sub
Private Sub Form_Load()
  '*** Jenny Colour
  '(Bef 13/08/05)
  SetMDISysColorsAfterLogin   '*** Jenny Color (For Setting System Colors)
  
  Dim ws_BtnStr As String
  ws_BtnStr = ""
  Call SetObjColors(Me, ws_BtnStr)
  '*** Jenny Colour
End Sub
Private Sub ADC_Load()
    Set moCn = adc.Connection    '*** (Bef speed) bef 24/11/06
    Set adc.FirNKeyCtl = adc("UoOmTcFr")
    Call InitProp(Me)

    gs_PTyp = "GRDCD"
    gs_PMCd = "DIA"
    gs_OmTcTyp = "QT"
    
    Call SetGroupSort
End Sub
Private Sub ATXT_FldChange(Index As Integer, ByVal IdName As String, Cancel As Boolean, ErrMsg As String, ByVal pv_NewValue As Variant)
  
  Select Case UCase(IdName)
  Case Is = UCase("UoPSCdSel")
    If UBound(Split(pv_NewValue, ",")) >= 15 Then
      ErrMsg = "Maximum 15 Diamond Grades Can Be Selected "
      Cancel = True
      Exit Sub
    End If
  End Select
  
  Call SetProp(Me, IdName, Valid, Cancel, ErrMsg)
End Sub
Private Sub ATXT_FldWhen(Index As Integer, ByVal IdName As String, Cancel As Boolean, ErrMsg As String)
  SetProp Me, IdName, When
End Sub
Private Sub tab_REP_Click(PreviousTab As Integer)
    fra_tabrep(TAB_REP.Tab).Enabled = True
    fra_tabrep(PreviousTab).Enabled = False
End Sub
Private Sub ADC_KeyWhen()
    adc("UOUSRCD") = adc.UsrCd
    adc("UOMNUCD") = adc.MenuCd
End Sub
Private Sub ADC_NonKeyWhen(pr_Cancel As Boolean, pr_ErrMsg As String)
   adc("UoCoCdFr") = gs_CoCd
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
   
    'Set moCn = Nothing   '*** (Bef speed) bef 24/11/06
End Sub

