VERSION 5.00
Object = "{AB2CC32A-E652-42A2-B341-26ECE02F6386}#2.1#0"; "MwfCtl.ocx"
Begin VB.Form EmrFrmLsRec 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Loss Recovery"
   ClientHeight    =   9810
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   15210
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   9810
   ScaleWidth      =   15210
   Begin MwfCtl.MWCTL_BTN1 CmdCpy 
      Height          =   405
      Left            =   3120
      TabIndex        =   14
      ToolTipText     =   "Option To Copy The Rates From Another Customer"
      Top             =   9300
      Width           =   825
      _ExtentX        =   1455
      _ExtentY        =   714
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
   Begin MwfCtl.MWCTL_HLP MWCTL_HLP1 
      Height          =   405
      Left            =   4560
      TabIndex        =   15
      TabStop         =   0   'False
      Top             =   9420
      Visible         =   0   'False
      Width           =   525
      _ExtentX        =   926
      _ExtentY        =   714
   End
   Begin MwfCtl.MWCTL_ADC ADC 
      CausesValidation=   0   'False
      Height          =   555
      Left            =   90
      TabIndex        =   16
      Top             =   9300
      Width           =   7650
      _ExtentX        =   13494
      _ExtentY        =   979
      AllowAdd        =   0   'False
      BtnWidth        =   885
      BtnHeight       =   435
      BtnMaskColor    =   8454143
      FontSz          =   8.25
   End
   Begin VB.Frame FRA_KEYS 
      BorderStyle     =   0  'None
      Height          =   630
      Left            =   -36
      TabIndex        =   18
      Top             =   0
      Width           =   7890
      Begin MwfCtl.MWCTL_MED ATXT 
         Height          =   285
         Index           =   2
         Left            =   2895
         TabIndex        =   2
         ToolTipText     =   "Enter From Process Location"
         Top             =   15
         Width           =   1485
         _ExtentX        =   2619
         _ExtentY        =   503
         MaxLength       =   8
         DataType        =   4
         IdName          =   "WLRCLOCFR"
         CmpStr          =   "LrcLoc>="
      End
      Begin MwfCtl.MWCTL_MED ATXT 
         Height          =   285
         Index           =   3
         Left            =   2895
         TabIndex        =   3
         ToolTipText     =   "Enter To Process Location"
         Top             =   300
         Width           =   1485
         _ExtentX        =   2619
         _ExtentY        =   503
         MaxLength       =   8
         DataType        =   4
         IdName          =   "WLRCLOCTO"
         CmpStr          =   "LrcLoc<="
      End
      Begin MwfCtl.MWCTL_MED ATXT 
         Height          =   285
         Index           =   4
         Left            =   5415
         TabIndex        =   4
         ToolTipText     =   "Enter From Raw Material Code"
         Top             =   15
         Width           =   2460
         _ExtentX        =   4339
         _ExtentY        =   503
         MaxLength       =   16
         DataType        =   4
         IdName          =   "WLRCRMCDFR"
         CmpStr          =   "LrcRmCd>="
      End
      Begin MwfCtl.MWCTL_MED ATXT 
         Height          =   285
         Index           =   5
         Left            =   5415
         TabIndex        =   5
         ToolTipText     =   "Enter To Raw Material Code"
         Top             =   300
         Width           =   2460
         _ExtentX        =   4339
         _ExtentY        =   503
         MaxLength       =   16
         DataType        =   4
         IdName          =   "WLRCRMCDTO"
         CmpStr          =   "LrcRmCd<="
      End
      Begin MwfCtl.MWCTL_MED ATXT 
         Height          =   285
         Index           =   0
         Left            =   1125
         TabIndex        =   0
         ToolTipText     =   "Enter Year And Month"
         Top             =   15
         Width           =   885
         _ExtentX        =   1561
         _ExtentY        =   503
         Mask            =   "#####0"
         MaxLength       =   6
         DataType        =   1
         IdName          =   "WLRCYYMM"
         CmpStr          =   "LrcYyMM="
      End
      Begin MwfCtl.MWCTL_MED ATXT 
         Height          =   285
         Index           =   1
         Left            =   1125
         TabIndex        =   1
         ToolTipText     =   "Enter Raw Material Category"
         Top             =   300
         Width           =   480
         _ExtentX        =   847
         _ExtentY        =   503
         MaxLength       =   2
         DataType        =   4
         IdName          =   "WLRCRMCTG"
         CmpStr          =   "LrCRmCtg="
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
         Index           =   5
         Left            =   240
         TabIndex        =   29
         Top             =   15
         Width           =   915
      End
      Begin VB.Label ALBL 
         BackStyle       =   0  'Transparent
         Caption         =   "            To"
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
         Left            =   5100
         TabIndex        =   24
         Top             =   300
         Width           =   1035
      End
      Begin VB.Label ALBL 
         BackStyle       =   0  'Transparent
         Caption         =   "RmCd Fr"
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
         Left            =   4470
         TabIndex        =   23
         Top             =   15
         Width           =   1005
      End
      Begin VB.Label ALBL 
         BackStyle       =   0  'Transparent
         Caption         =   "               To"
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
         Left            =   1680
         TabIndex        =   22
         Top             =   300
         Width           =   1305
      End
      Begin VB.Label ALBL 
         BackStyle       =   0  'Transparent
         Caption         =   "Loc Fr"
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
         Left            =   2160
         TabIndex        =   21
         Top             =   15
         Width           =   705
      End
      Begin VB.Label ALBL 
         BackStyle       =   0  'Transparent
         Caption         =   "Rm Ctg"
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
         Left            =   240
         TabIndex        =   19
         Top             =   300
         Width           =   765
      End
   End
   Begin VB.Frame FRA_NKEYS 
      BorderStyle     =   0  'None
      Enabled         =   0   'False
      Height          =   9450
      Left            =   -120
      TabIndex        =   17
      Top             =   0
      Width           =   8175
      Begin VB.Frame FraCpy 
         Height          =   1695
         Left            =   1080
         TabIndex        =   7
         Top             =   7515
         Visible         =   0   'False
         Width           =   5865
         Begin MwfCtl.MWCTL_BTN1 CmdCpyGo 
            Height          =   285
            Left            =   5190
            TabIndex        =   13
            ToolTipText     =   "Press This Button To Copy"
            Top             =   1290
            Width           =   555
            _ExtentX        =   0
            _ExtentY        =   0
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
            Index           =   6
            Left            =   1770
            TabIndex        =   8
            ToolTipText     =   "Enter Source Year Month"
            Top             =   420
            Width           =   870
            _ExtentX        =   1535
            _ExtentY        =   503
            Mask            =   "#####0"
            MaxLength       =   6
            DataType        =   1
            IdName          =   "WCPYYYMMFR"
         End
         Begin MwfCtl.MWCTL_MED ATXT 
            Height          =   285
            Index           =   7
            Left            =   1770
            TabIndex        =   9
            ToolTipText     =   "Enter Source Raw Material Category"
            Top             =   705
            Width           =   450
            _ExtentX        =   794
            _ExtentY        =   503
            MaxLength       =   2
            DataType        =   4
            IdName          =   "WCPYRMCTGFR"
         End
         Begin MwfCtl.MWCTL_MED ATXT 
            Height          =   285
            Index           =   10
            Left            =   4860
            TabIndex        =   12
            ToolTipText     =   "Enter Target Year Month"
            Top             =   420
            Width           =   870
            _ExtentX        =   1535
            _ExtentY        =   503
            Mask            =   "#####0"
            MaxLength       =   6
            DataType        =   1
            IdName          =   "WCPYYYMMTO"
         End
         Begin MwfCtl.MWCTL_MED ATXT 
            Height          =   285
            Index           =   8
            Left            =   1770
            TabIndex        =   10
            ToolTipText     =   "Enter From Location"
            Top             =   990
            Width           =   1365
            _ExtentX        =   2408
            _ExtentY        =   503
            MaxLength       =   8
            DataType        =   4
            IdName          =   "WCPYLOCFR"
         End
         Begin MwfCtl.MWCTL_MED ATXT 
            Height          =   285
            Index           =   9
            Left            =   1770
            TabIndex        =   11
            ToolTipText     =   "Enter To Location"
            Top             =   1275
            Width           =   1365
            _ExtentX        =   2408
            _ExtentY        =   503
            MaxLength       =   8
            DataType        =   4
            IdName          =   "WCPYLOCTO"
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
            Height          =   270
            Index           =   0
            Left            =   -30
            TabIndex        =   31
            Top             =   90
            Width           =   5895
         End
         Begin VB.Label ALBL 
            BackStyle       =   0  'Transparent
            Caption         =   "Location  Fr"
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
            Index           =   14
            Left            =   60
            TabIndex        =   30
            Top             =   990
            Width           =   1725
         End
         Begin VB.Label ALBL 
            BackStyle       =   0  'Transparent
            Caption         =   "Target YyMm"
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
            Index           =   17
            Left            =   3360
            TabIndex        =   28
            Top             =   420
            Width           =   1425
         End
         Begin VB.Label ALBL 
            BackStyle       =   0  'Transparent
            Caption         =   "Source YyyyMm"
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
            Left            =   60
            TabIndex        =   27
            Top             =   420
            Width           =   1725
         End
         Begin VB.Label ALBL 
            BackStyle       =   0  'Transparent
            Caption         =   "                To"
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
            Index           =   12
            Left            =   60
            TabIndex        =   26
            Top             =   1275
            Width           =   1725
         End
         Begin VB.Label ALBL 
            BackStyle       =   0  'Transparent
            Caption         =   "Source Rm Ctg"
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
            Left            =   60
            TabIndex        =   25
            Top             =   705
            Width           =   1725
         End
      End
      Begin VB.Frame FraNKeyAll 
         BorderStyle     =   0  'None
         Height          =   9315
         Left            =   150
         TabIndex        =   20
         Top             =   90
         Width           =   8010
         Begin MwfCtl.MWCTL_FLX GrdLsRec 
            Height          =   8550
            Left            =   90
            TabIndex        =   6
            Top             =   570
            Width           =   7875
            _ExtentX        =   13891
            _ExtentY        =   15081
            Cols            =   9
            colname1        =   "LRCYYMM"
            heading1        =   "YyyyMm"
            datafld1        =   "LrcYyMm"
            datatype1       =   1
            mask1           =   "#####0"
            maxlength1      =   6
            tooltiptext1    =   "Enter Year And Month"
            colname2        =   "LRCLOC"
            heading2        =   "Loc                     "
            datafld2        =   "LrcLoc"
            datatype2       =   4
            maxlength2      =   8
            tooltiptext2    =   "Enter Process Location "
            colname3        =   "LRCRMCTG"
            heading3        =   "Ctg  "
            datafld3        =   "LrcRmCtg"
            datatype3       =   4
            maxlength3      =   2
            tooltiptext3    =   "Enter Raw Material Category"
            colname4        =   "LRCRMCD"
            heading4        =   "Rm Code                                  "
            datafld4        =   "LrcRmCd"
            datatype4       =   4
            maxlength4      =   16
            tooltiptext4    =   "Enter Raw Material Code"
            colname5        =   "LRCDUSTWT"
            heading5        =   "Dust Wt    "
            datafld5        =   "LrcDustWt"
            datatype5       =   2
            mask5           =   "####0.000"
            maxlength5      =   9
            tooltiptext5    =   "Enter Dust Weight"
            colname6        =   "LRCACTLS"
            heading6        =   "Act Loss    "
            datafld6        =   "LrcActLs"
            datatype6       =   2
            mask6           =   "####0.000"
            maxlength6      =   9
            tooltiptext6    =   "Enter Actual Loss"
            colname7        =   "LRCRECWT"
            heading7        =   "Rec Wt      "
            datafld7        =   "LrcRecWt"
            datatype7       =   2
            mask7           =   "####0.000"
            maxlength7      =   9
            tooltiptext7    =   "Enter Recovery"
            colname8        =   "LRCCOCD"
            datafld8        =   "LrcCoCd"
            maxlength8      =   3
            style8          =   0
         End
      End
   End
End
Attribute VB_Name = "EmrFrmLsRec"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'----- General comments on the Entry  ----
' A level Entry
' X-'', A-LsRec, GrdLsRec
' This entry is meant for the Loss Recovery Data entry
's
' The upper part of the screen specifies the scope to be given for the search purposes.
' The data in the grid goes in the table.
' The entry has a copy routine too which can be used to copy the previous month data into the current month data.
'-----
Option Explicit
Dim moCn As MwfLib.MDOConnection   '*** (Bef speed) bef 24/11/06

Private Sub ADC_Load()
  '*** Set the First Key Control property to the working field wLrcYyMm
  '*** Set the First Non Key Control property to the grid GrdLsRec
  '*** Set the Previous and Next Control Properties for the Grid GrdLsRec as ADC
  '*** Set the Hot Key for the Grid GrdLsRec as 'A'

  Set moCn = ADC.Connection    '*** (Bef speed) bef 24/11/06
  ADC.Child = "GrdLsRec"
  Set ADC.FirKeyCtl = ADC("wLrcYyMm")
  Set ADC.FirNKeyCtl = GrdLsRec
  Set GrdLsRec.PrevCtl = ADC
  Set GrdLsRec.NextCtl = ADC
  GrdLsRec.HotKey = "A"
End Sub
Private Sub ADC_KeyWhen()
  '*** Disable All command buttons
  '*** Hide all frames except Frame FraNKeyAll

  Call EnaDisaCmds(True)
  Call HideAllFras
End Sub
Private Sub ADC_NonKeyWhen(pr_Cancel As Boolean, pr_ErrMsg As String)
  '*** Enable All command buttons
  Call EnaDisaCmds(False)
End Sub
Private Sub ADC_KeyValid(pr_Cancel As Boolean, pr_ErrMsg As String)
'*** Check the compulsory fields i.e. Year-Month and the RmCtg
'*** when the control enters from Key Frame to the Non Key by calling the Atxt Field Change event

  Call atxt_FldChange(99, "wLrcYyMm", pr_Cancel, pr_ErrMsg, ADC("wLrcYyMm"))
  If pr_Cancel = True Then Exit Sub
  Call atxt_FldChange(99, "wLrcRmCtg", pr_Cancel, pr_ErrMsg, ADC("wLrcRmCtg"))
  If pr_Cancel = True Then Exit Sub
End Sub
Private Sub ATXT_FldWhen(Index As Integer, ByVal IdName As String, Cancel As Boolean, ErrMsg As String)
  '*** Set Helps For Each Field
  '*** For Working Fields in the Scope
  '*** For wLrcYyMm : LrcYyMm from the Previous Records from LsREc Table
  '*** For wLrcRmCtg : RmCtg from Param where ptyp is RmCtg
  '*** For wLrcLocFr and wLrcLocTo : LocCd from Loc where LocTyp = 'P'
  '*** For wLrcRmCdFr and wLrcRmCdTo : RmCd from RmMst where RmCtg = WLrcRmCtg
  '*** For Working Fields in the Copy Option
  '*** For wCpyYyMmFr : LrcYyMm from the Previous Records from LsRec Table
  '*** For wCpyRmCtgFr : RmCtg from Param where ptyp is RmCtg
  '*** For wCpyLocFr and wCpyLocTo : LocCd from Loc where LocTyp = 'P'
    
  Select Case IdName
  Case Is = UCase("wLrcYyMm"), UCase("wCpyYyMmFr")
    Call HlpList.LrcYyMm(gs_CoCd)
  Case Is = UCase("wLrcRmCtg"), UCase("wCpyRmCtgFr")
    Call HlpList.PMCd("RMCTG")
  Case Is = UCase("wLrcLocFr"), UCase("wLrcLocTo"), UCase("wCpyLocFr"), UCase("wCpyLocTo")
    Call HlpList.Loc(gs_CoCd, "'P'", , , False)
  Case Is = UCase("wLrcRmCdFr"), UCase("wLrcRmCdTo")
    Call HlpList.RmCd("N", "'" + ADC("WLRCRMCTG") + "'", , False)
  End Select
End Sub
Private Sub atxt_FldChange(Index As Integer, ByVal IdName As String, Cancel As Boolean, ErrMsg As String, ByVal pv_NewValue As Variant)
  '*** wLrcRmCtg     Only Category "G" or "P" or "S" is allowed
  '*** wLrcLocFr,wLrcLocTo,wCpyLocFr, wCpyLocTo     valid LocCd for Loctyp="P"
  '*** wLrcRmCd     valid RMCd for RmCtg =wLrcRmCtg
  
  Select Case UCase(IdName)
  Case Is = UCase("wLrcYyMm"), UCase("wCpyYyMmFr"), UCase("wCpyYyMmTo")
    Cancel = (pv_NewValue = "")
    ErrMsg = "Year Month Can not be Left Blank": Exit Sub
  Case Is = UCase("wLrcRmCtg"), UCase("wCpyRmCtgFr")
    ' ***** Manali - 3.03 - 20/06/08 - 'L' included
    Cancel = Not moCn.RecSeek("Select PmCd From Param where PTyp= 'RMCTG' " + _
             "and PMCd = '" + pv_NewValue + "' and PMCd in ('G','P','S','L')")   ' ****** Sachin 2.12.0 - Silver added (23/11/05)
    ErrMsg = "Invalid Rm Category Can not be Blank/(Only G/P/S/L Allowed)": Exit Sub
  Case Is = UCase("wLrcLocFr"), UCase("wLrcLocTo"), UCase("wCpyLocFr"), UCase("wCpyLocTo")
    Cancel = (pv_NewValue <> "") And (Not moCn.RecSeek("Select LocCd From Loc " + _
             "where LocCoCd ='" + gs_CoCd + "' and LocTyp= 'P' and " + _
             "LocCd = '" + pv_NewValue + "'"))
    ErrMsg = "Invalid Process Location": Exit Sub
  Case Is = UCase("wLrcRmCdFr")
    Cancel = (pv_NewValue <> "") And (Not moCn.RecSeek("Select RMCd From RmMst where " + IIF(gs_Partition = ctCurrPrtn, " RmPrtKey='" + ctCurrPrtn + "' And ", "") + _
             "RmCtg=  '" + ADC("wLrcRmctg") + "' and RMCd = '" + pv_NewValue + "'"))
    ErrMsg = "Invalid Rm Code Type": Exit Sub
  Case Is = UCase("wCpyRmCdTo")
    Cancel = (pv_NewValue <> "") And (Not moCn.RecSeek("Select RMCd From RmMst where " + IIF(gs_Partition = ctCurrPrtn, " RmPrtKey='" + ctCurrPrtn + "' And ", "") + " RmCtg=  " + ADC("wCpyRmctgFr") + _
                         "and RMCd = '" + pv_NewValue + "'"))
    ErrMsg = "Invalid Rm Code Type": Exit Sub
  End Select
End Sub
Private Sub GrdLsRec_ColChange(ByVal RowNum As Integer, ByVal ColName As String, ByVal NewValue As Variant, Cancel As Boolean, ErrMsg As String)
'*** Check for RmCtg in the Grid if not "G' or "P" then Restrict movement to the next column
  Select Case ColName
  Case Is = UCase("LrcRmCtg")
    ' ***** Manali - 3.03 -20/06/08 - 'L' included
    Cancel = Not moCn.RecSeek("Select PMCd From Param Where PTyp= 'RMCTG' " + _
             "and PMCd = '" + NewValue + "' and PMCd in ('G','P','S','L')")   ' ****** sachin 2.12 - silver added (23/11/05)
    ErrMsg = "Invalid Rm Category (Only G/P/S/L Allowed)": Exit Sub
  Case Is = UCase("LrcRecWt")
    '???
'    If GrdLsRec.Value(RowNum, "LrcActLs") = 0 And GrdLsRec.Value(RowNum, "LrcDustWt") = 0 Then
'      Cancel = True
'      ErrMsg = "Recovery Cannot be More than DustWt/ Actual Loss": Exit Sub
'    End If
'    If GrdLsRec.Value(RowNum, "LrcActLs") <> 0# And GrdLsRec.Value(RowNum, "LrcDustWt") = 0# And NewValue > GrdLsRec.Value(RowNum, "LrcActLs") Then
'      Cancel = True
'      ErrMsg = "Recovery Can not be More than DustWt/ Actual Loss": Exit Sub
'    End If
'    If GrdLsRec.Value(RowNum, "LrcDustWt") <> 0# And GrdLsRec.Value(RowNum, "LrcActLs") = 0# And NewValue > GrdLsRec.Value(RowNum, "LrcDustWt") Then
'      Cancel = True
'      ErrMsg = "Recovery Can not be More than DustWt/ Actual Loss": Exit Sub
'    End If
'    If GrdLsRec.Value(RowNum, "LrcDustWt") <> 0# And GrdLsRec.Value(RowNum, "LrcActLs") <> 0# And NewValue > GrdLsRec.Value(RowNum, "LrcDustWt") + GrdLsRec.Value(RowNum, "LrcActLs") Then
'      Cancel = True
'      ErrMsg = "Recovery Can not be More than DustWt/ Actual Loss": Exit Sub
'    End If
'???
  End Select
End Sub
Private Sub GrdLsRec_ColWhen(ByVal RowNum As Integer, ByVal ColName As String, Cancel As Boolean, ErrMsg As String, ByVal Action As MwfCtl.en_fgAction)
  '*** Set Helps For Each Field of the Grid
  '*** LrcYyMM : LrcYyMM of the previous months from the LsRec Table
  '*** LrcLoc : LocCd from the Loc Table Where LocTyp="P"
  '*** LrcRmCtg : PmCd from Param where ptyp="RmCtg"
  '*** LrcRmCd : RmCd from RmMst where RmZ = 'N'
  
  Select Case ColName
  Case Is = UCase("LrcYyMm")
    Call HlpList.LrcYyMm(UCase(gs_CoCd))
  Case Is = UCase("LrcLoc")
    Call HlpList.Loc(gs_CoCd, "'P'")
  Case Is = UCase("LrcRmCtg")
    Call HlpList.PMCd("RMCTG")
  Case Is = UCase("LrcRmCd")
    Call HlpList.RmCd("N", "'" + GrdLsRec.Value(RowNum, "LrcRmCtg") + "'")
  End Select
End Sub
Private Sub ADC_SetRecSource()
  '*** Since this is an 'A' Entry, there is no RecSource associated with ADC. So set the
   'DataMode property of ADC control as xdatUnbound
  ADC.DataMode = xdatUnbound
End Sub
Private Sub GrdLsRec_InitKey(ByVal RowNum As Integer)
  '*** Store the Company Code in the Grid column
  GrdLsRec.Store "LrcCoCd", gs_CoCd
End Sub
Private Sub GrdLsRec_RowWhen(ByVal RowNum As Integer)
  If GrdLsRec.IsNew(RowNum) Then
    GrdLsRec.Value(RowNum, "LrcYyMm") = ADC("wLrcYyMm")
    GrdLsRec.Value(RowNum, "LrcRmCtg") = ADC("wLrcRmCtg")
  End If
End Sub
Private Sub GrdLsRec_SetRecSource()
  '*** Set the Record Source of the Grid GrdLsRec
  Dim ws_Cnd As String
  
  ws_Cnd = ADC.RepCond
  ws_Cnd = IIF(ws_Cnd <> "", " where " + ws_Cnd, "")
  GrdLsRec.RecSource = "Select * from LsRec " + ws_Cnd + " and LrcCoCd = '" + gs_CoCd + "' Order By LrcYyMm, LrcLoc,LrcRmCtg,LrcRmCd "
End Sub
Private Sub CmdCpyGo_Click()
'*** Check the options to be Compulsorily filled like source and target Year Month
 If ADC("wCpyYyMmFr") = 0 Or ADC("wCpyYyMmTo") = 0 Then
    DispMsg "Source And Target Year Month Should Be Filled", etError
    Exit Sub
  End If
  If ADC("wCpyRmCtgFr") = "" Then
    DispMsg "Raw Material Category Should Be Filled", etError
    Exit Sub
  End If
  Call CpyLsRec
  Call CmdCpy_Click
End Sub
Private Sub CpyLsRec()
  '*** Copy Option will copy the previous any months Loss Records to the month specified.
  '*** Records can be copied only if already there are no records for that month

  Dim ws_RmCnd As String
  Call CmpString       ' to set the Compare String Property of the ADC controls when the Copy Option Is clicked

  '*** Checking if copy is possible
   
  ws_RmCnd = ADC.RepCond
  ws_RmCnd = IIF(ws_RmCnd <> "", " and " + ws_RmCnd, "")

  '*** Check if the Records exist for the given traget Year Month & Scope in the Table Already.
  If moCn.RecSeek("Select 'x' from LsRec where LrcCoCd='" + gs_CoCd + "' and " + _
                  " LrcYyMm = " & ADC("wCpyYyMmTo") & " and LrcRmCtg = '" + ADC("wCpyRmCtgFr") + "'") Then
    DispMsg "Cannot Copy as Loss Records Exist For The Target Month and Raw Material Category", etError
    Exit Sub
  End If

  '*** Beginning the copy process
  GrdLsRec.StartCopy
  MWLib.BeginProcess Me, "Copying Loss Recovery ..."
  Dim wRsLsRec As MwfLib.MDORowSet, wi_Row As Integer
  Set wRsLsRec = moCn.OpenRes("Select * from LsRec where LrcCoCd = '" + gs_CoCd + "' " + ws_RmCnd + _
                 " Order By LrcCoCd, LrcYyMm, LrcLoc,LrcRmCtg, LrcRmCd")
  With wRsLsRec
    Do While Not (.EOF Or .BOF)
      MWLib.RunningProcess "Copying Loss Recovery For RM Ctg " + !LrcRmCtg
      GrdLsRec.AddItem
      wi_Row = GrdLsRec.Rows - 1

      GrdLsRec.Value(wi_Row, "LrcCocd") = !LrcCoCd
      GrdLsRec.Value(wi_Row, "LrcYyMm") = ADC("wCpyYyMmTo")
      GrdLsRec.Value(wi_Row, "LrcLoc") = !LrcLoc
      GrdLsRec.Value(wi_Row, "LrcRmCtg") = !LrcRmCtg
      GrdLsRec.Value(wi_Row, "LrcRmCd") = !LrcRmCd
      GrdLsRec.Value(wi_Row, "LrcDustWt") = !LrcDustWt
      GrdLsRec.Value(wi_Row, "LrcActLs") = !LrcActLs
      GrdLsRec.Value(wi_Row, "LrcRecWt") = !LrcRecWt
      
      GrdLsRec.SaveRec (wi_Row)
      .MoveNext
    Loop
  End With
  MWLib.EndProcess Me
  GrdLsRec.EndCopy
  Set wRsLsRec = Nothing
  '*** End of the copy process
End Sub
Private Sub CmpString()
'*** Builds the compare string for the copy option.
  ADC("wCpyYyMmFr").CmpStr = "LrcYyMm = "
  ADC("wCpyRmCtgFr").CmpStr = "LrcRmCtg =  "
  ADC("wCpyLocFr").CmpStr = "LrcLoc >= "
  ADC("wCpyLocTo").CmpStr = "LrcLoc <= "
End Sub

Private Sub EnaDisaCmds(ByVal pv_ShowFra As Boolean, Optional ByVal pv_CmdBut As MWCTL_BTN1)
  '*** This procedure is used to give the toggle option (enable and disable) to
      'the command buttons

  If Not pv_CmdBut Is Nothing Then
    If pv_ShowFra = True Then
      '*** Jenny (Old Code Bef Color)
'      pv_CmdBut.BackColor = vbWhite
'      pv_CmdBut.FontBold = True
      '*** Jenny (Old Code Bef Color)
      pv_CmdBut.CausesValidation = True
    Else
      '*** Jenny (Old Code Bef Color)
'      pv_CmdBut.BackColor = &H8000000F
'      pv_CmdBut.FontBold = False
      '*** Jenny (Old Code Bef Color)
      'pv_CmdBut.CausesValidation = False
    End If
  End If
  FraNKeyAll.Enabled = Not pv_ShowFra
  ADC.AllowSave = Not pv_ShowFra
  CmdCpy.Enabled = Not pv_ShowFra
  If Not pv_CmdBut Is Nothing Then pv_CmdBut.Enabled = True
End Sub
Private Sub HideAllFras()
  '*** This procedure hides all the frames other than FraNKeyAll
  FraNKeyAll.Visible = True
  FraCpy.Visible = False
  FraCpy.Enabled = False
End Sub
Private Sub CmdCpy_Click()
   '*** Display or hide the Copy frame depending on the toggle satate
    If FraCpy.Visible = True Then
      FraCpy.Visible = False
      FraCpy.Enabled = False
      Call EnaDisaCmds(False, CmdCpy)
      CmdCpy.SetFocus
    Else
      FraCpy.Visible = True
      FraCpy.Enabled = True
      ADC("wCpyYyMmFr").SetFocus
      FraCpy.ZOrder
      Call EnaDisaCmds(True, CmdCpy)
    End If
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
    
    Set moCn = Nothing    '*** (09/08/05)  '*** (Bef speed) bef 24/11/06
End Sub
Private Sub GrdLsRec_Validate(Cancel As Boolean)
  Cancel = GrdLsRec.Validate
End Sub


  

