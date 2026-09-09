VERSION 5.00
Object = "{AB2CC32A-E652-42A2-B341-26ECE02F6386}#2.0#0"; "MwfCtl.ocx"
Begin VB.Form EmrFrmLossEntry 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Job Work  Loss Entry"
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
      Left            =   7560
      TabIndex        =   11
      TabStop         =   0   'False
      Top             =   9360
      Visible         =   0   'False
      Width           =   525
      _ExtentX        =   926
      _ExtentY        =   714
   End
   Begin MwfCtl.MWCTL_ADC ADC 
      CausesValidation=   0   'False
      Height          =   555
      Left            =   30
      TabIndex        =   12
      Top             =   9330
      Width           =   11250
      _ExtentX        =   19844
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
      Height          =   570
      Left            =   -36
      TabIndex        =   14
      Top             =   30
      Width           =   11670
      Begin MwfCtl.MWCTL_MED atxt 
         Height          =   285
         Index           =   0
         Left            =   1200
         TabIndex        =   0
         ToolTipText     =   "Enter From Process "
         Top             =   0
         Width           =   1250
         _ExtentX        =   2196
         _ExtentY        =   503
         MaxLength       =   8
         DataType        =   4
         IdName          =   "WLMLOCCDFR"
         CmpStr          =   "LmLocCd >= "
      End
      Begin MwfCtl.MWCTL_MED atxt 
         Height          =   285
         Index           =   3
         Left            =   1200
         TabIndex        =   1
         ToolTipText     =   "Enter To Process"
         Top             =   285
         Width           =   1250
         _ExtentX        =   2196
         _ExtentY        =   503
         MaxLength       =   8
         DataType        =   4
         IdName          =   "WLMLOCCDTO"
         CmpStr          =   "LmLocCd <= "
      End
      Begin MwfCtl.MWCTL_MED atxt 
         Height          =   285
         Index           =   4
         Left            =   4935
         TabIndex        =   4
         ToolTipText     =   "Enter From Design category"
         Top             =   0
         Width           =   1170
         _ExtentX        =   2064
         _ExtentY        =   503
         MaxLength       =   5
         DataType        =   4
         IdName          =   "WLMDMCTGFR"
         CmpStr          =   "LmDmCtg >= "
      End
      Begin MwfCtl.MWCTL_MED atxt 
         Height          =   285
         Index           =   2
         Left            =   4935
         TabIndex        =   5
         ToolTipText     =   "Enter To Design Catogery"
         Top             =   285
         Width           =   1170
         _ExtentX        =   2064
         _ExtentY        =   503
         MaxLength       =   5
         DataType        =   4
         IdName          =   "WLMDMCTGTO"
         CmpStr          =   "LmDmCtg <= "
      End
      Begin MwfCtl.MWCTL_MED atxt 
         Height          =   285
         Index           =   10
         Left            =   6810
         TabIndex        =   6
         ToolTipText     =   "Enter From Design"
         Top             =   0
         Width           =   2175
         _ExtentX        =   3836
         _ExtentY        =   503
         MaxLength       =   15
         DataType        =   4
         IdName          =   "WLMDMCDFR"
         CmpStr          =   "LmDmCtg>="
      End
      Begin MwfCtl.MWCTL_MED atxt 
         Height          =   285
         Index           =   17
         Left            =   6810
         TabIndex        =   7
         ToolTipText     =   "Enter To Design"
         Top             =   285
         Width           =   2175
         _ExtentX        =   3836
         _ExtentY        =   503
         MaxLength       =   15
         DataType        =   4
         IdName          =   "WLMDMCDTO"
         CmpStr          =   "LmDmCd<="
      End
      Begin MwfCtl.MWCTL_MED atxt 
         Height          =   285
         Index           =   87
         Left            =   11310
         TabIndex        =   22
         Top             =   0
         Visible         =   0   'False
         Width           =   270
         _ExtentX        =   476
         _ExtentY        =   503
         Enabled         =   0   'False
         MaxLength       =   3
         DataType        =   4
         IdName          =   "WLMCOCD"
      End
      Begin MwfCtl.MWCTL_MED atxt 
         Height          =   285
         Index           =   9
         Left            =   10005
         TabIndex        =   8
         ToolTipText     =   "Enter From Worker"
         Top             =   0
         Width           =   1290
         _ExtentX        =   2275
         _ExtentY        =   503
         MaxLength       =   8
         DataType        =   4
         IdName          =   "WLMWORKCDFR"
         CmpStr          =   "LmWorkCd >= "
      End
      Begin MwfCtl.MWCTL_MED atxt 
         Height          =   285
         Index           =   11
         Left            =   10005
         TabIndex        =   9
         ToolTipText     =   "Enter To Worker"
         Top             =   285
         Width           =   1290
         _ExtentX        =   2275
         _ExtentY        =   503
         MaxLength       =   8
         DataType        =   4
         IdName          =   "WLMWORKCDTO"
         CmpStr          =   "LmWorkCd <= "
      End
      Begin MwfCtl.MWCTL_MED atxt 
         Height          =   285
         Index           =   1
         Left            =   3495
         TabIndex        =   2
         ToolTipText     =   "Enter From Design category"
         Top             =   0
         Width           =   480
         _ExtentX        =   847
         _ExtentY        =   503
         MaxLength       =   5
         DataType        =   4
         IdName          =   "WLMRMCTGFR"
         CmpStr          =   "LmRmCtg >= "
      End
      Begin MwfCtl.MWCTL_MED atxt 
         Height          =   285
         Index           =   5
         Left            =   3495
         TabIndex        =   3
         ToolTipText     =   "Enter To Design Catogery"
         Top             =   285
         Width           =   480
         _ExtentX        =   847
         _ExtentY        =   503
         MaxLength       =   5
         DataType        =   4
         IdName          =   "WLMRMCTGTO"
         CmpStr          =   "LmRmCtg <= "
      End
      Begin VB.Label ALBL 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "RmCtg Fr"
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
         Index           =   5
         Left            =   2580
         TabIndex        =   28
         Top             =   0
         Width           =   885
      End
      Begin VB.Label ALBL 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
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
         Height          =   240
         Index           =   0
         Left            =   3210
         TabIndex        =   27
         Top             =   285
         Width           =   240
      End
      Begin VB.Label ALBL 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Worker Fr"
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
         Index           =   10
         Left            =   9000
         TabIndex        =   26
         Top             =   0
         Width           =   945
      End
      Begin VB.Label ALBL 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
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
         Height          =   240
         Index           =   9
         Left            =   9705
         TabIndex        =   25
         Top             =   285
         Width           =   240
      End
      Begin VB.Label ALBL 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
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
         Height          =   240
         Index           =   17
         Left            =   6525
         TabIndex        =   21
         Top             =   285
         Width           =   240
      End
      Begin VB.Label ALBL 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Dsg Fr"
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
         Index           =   16
         Left            =   6165
         TabIndex        =   20
         Top             =   0
         Width           =   600
      End
      Begin VB.Label ALBL 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
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
         Height          =   240
         Index           =   4
         Left            =   4665
         TabIndex        =   19
         Top             =   285
         Width           =   240
      End
      Begin VB.Label ALBL 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "DmCtg Fr"
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
         Index           =   3
         Left            =   4020
         TabIndex        =   18
         Top             =   0
         Width           =   885
      End
      Begin VB.Label ALBL 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
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
         Height          =   240
         Index           =   2
         Left            =   930
         TabIndex        =   17
         Top             =   285
         Width           =   240
      End
      Begin VB.Label ALBL 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Process Fr"
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
         Left            =   180
         TabIndex        =   16
         Top             =   0
         Width           =   990
      End
   End
   Begin VB.Frame FRA_NKEYS 
      BorderStyle     =   0  'None
      Enabled         =   0   'False
      Height          =   9300
      Left            =   -120
      TabIndex        =   13
      Top             =   0
      Width           =   11745
      Begin VB.Frame FraRmDesc 
         Height          =   495
         Left            =   120
         TabIndex        =   29
         Top             =   8820
         Width           =   11310
         Begin MwfCtl.MWCTL_MED TXTREMARKS 
            Height          =   285
            Index           =   0
            Left            =   930
            TabIndex        =   30
            ToolTipText     =   "Enter Raw Material Print Description"
            Top             =   150
            Width           =   9765
            _ExtentX        =   17224
            _ExtentY        =   503
            MaxLength       =   100
            IdName          =   "WLMREMARKS"
         End
         Begin MwfCtl.MWCTL_BTN1 cmdOK 
            Height          =   315
            Left            =   10710
            TabIndex        =   33
            Top             =   150
            Width           =   555
            _ExtentX        =   979
            _ExtentY        =   556
            BackColor       =   -2147483633
            MaskColor       =   8454143
            Caption         =   "&OK"
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
         Begin VB.Label ALBL 
            BackStyle       =   0  'Transparent
            Caption         =   "Remarks"
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
            TabIndex        =   32
            Top             =   150
            Width           =   855
         End
         Begin VB.Label ALBL 
            BackStyle       =   0  'Transparent
            Caption         =   "Valid (Y/N)"
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
            Index           =   34
            Left            =   11820
            TabIndex        =   31
            Top             =   465
            Width           =   1245
         End
      End
      Begin VB.Frame FraNKeyAll 
         BorderStyle     =   0  'None
         Height          =   9255
         Left            =   150
         TabIndex        =   15
         Top             =   90
         Width           =   11370
         Begin MwfCtl.MWCTL_FLX GrdLm 
            Height          =   8190
            Left            =   0
            TabIndex        =   10
            Top             =   570
            Width           =   11265
            _ExtentX        =   19870
            _ExtentY        =   14446
            Cols            =   12
            colname1        =   "LMLOCCD"
            heading1        =   "Process        "
            datafld1        =   "LmLocCd"
            datatype1       =   4
            maxlength1      =   8
            tooltiptext1    =   "Enter process"
            colname2        =   "LMRMCTG"
            heading2        =   "Rm Ctg"
            datafld2        =   "LmRmCtg"
            datatype2       =   4
            maxlength2      =   2
            tooltiptext2    =   "Enter Rm Category"
            colname3        =   "LMDMCTG"
            heading3        =   "Dsg Ctg "
            datafld3        =   "LmDmCtg"
            datatype3       =   4
            maxlength3      =   5
            tooltiptext3    =   "Enter Design Category"
            colname4        =   "LMDMCD"
            heading4        =   "Design Cd                       "
            datafld4        =   "LmDmCd"
            datatype4       =   4
            maxlength4      =   15
            tooltiptext4    =   "Enter Design Code"
            colname5        =   "LMWORKCD"
            heading5        =   "Worker            "
            datafld5        =   "LmWorkCd"
            datatype5       =   4
            maxlength5      =   8
            tooltiptext5    =   "Enter Worker"
            colname6        =   "LMFRRMWT"
            heading6        =   "From Rm Wt"
            datafld6        =   "LmFrRmWt"
            datatype6       =   2
            mask6           =   "####0.000"
            maxlength6      =   9
            tooltiptext6    =   "Enter From Rm Weight"
            colname7        =   "TEST"
            style7          =   0
            colname8        =   "LMTORMWT"
            heading8        =   "To Rm Wt     "
            datafld8        =   "LmToRmWt"
            datatype8       =   2
            mask8           =   "####0.000"
            maxlength8      =   9
            tooltiptext8    =   "Enter To Rm Weight"
            colname9        =   "LMLOSSPER"
            heading9        =   "Loss %     "
            datafld9        =   "LmLossPer"
            datatype9       =   2
            mask9           =   "##0.00"
            maxlength9      =   6
            tooltiptext9    =   "Enter Permitted Loss Percentage"
            colname10       =   "LMVALIDYN"
            heading10       =   "Valid"
            datafld10       =   "LMVALIDYN"
            datatype10      =   4
            maxlength10     =   1
            tooltiptext10   =   "Please enter Y for current N for history"
            colname11       =   "LMREMARKS"
            heading11       =   "                                                                             "
            datafld11       =   "LmRemarks"
            maxlength11     =   30
            style11         =   0
            tooltiptext11   =   "Enter Remarks"
         End
      End
   End
   Begin VB.Label ALBL 
      BackStyle       =   0  'Transparent
      Caption         =   "DmCtg Fr"
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
      Index           =   8
      Left            =   60
      TabIndex        =   24
      Top             =   0
      Width           =   1095
   End
   Begin VB.Label ALBL 
      BackStyle       =   0  'Transparent
      Caption         =   "              To"
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
      Left            =   0
      TabIndex        =   23
      Top             =   285
      Width           =   1155
   End
End
Attribute VB_Name = "EmrFrmLossEntry"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'----- General comments on the Entry  ----
' X-'', A-LossMst, GrdLm
' The grid properties are set by the TcStructure Table.
'
' private Sub CmpString is used to set the values for Compare String property,
'            passing 'K'ey Fields,
'
'
'-----
Option Explicit
Dim moCn As MwfLib.MDOConnection   '*** (Bef speed) bef 24/11/06
Dim mLmSrt As String

'*** For the Toggle Options of Command Buttons
'Enum en_LmFra1
'  Cpy = 0
'End Enum

Private Sub ADC_Load()
  '*** Set the First Key Control property to the working field wPpTyp
  '*** Set the First Non Key Control property to the grid GrdLm
  '*** Set the Child Property Of ADC to GrdLm
  '*** Set the Previous and Next Control Properties for the Grid GrdLm as ADC
  '*** Set the Hot Key for the Grid GrdLm as 'A'
  '*** initialise the default values for the working fields

  Set moCn = ADC.Connection    '*** (Bef speed) bef 24/11/06

  Set ADC.FirKeyCtl = ADC("wLmLocCdFr")
  Set ADC.FirNKeyCtl = GrdLm
  ADC.Child = "GrdLm"

  Set GrdLm.PrevCtl = ADC
  Set GrdLm.NextCtl = ADC
  GrdLm.HotKey = "A"
 
End Sub
Private Sub ADC_KeyWhen()
  '*** pass 'K' as the parameter for the CmpString procedure. This sets the CmpStr value of
      'the fields in the copy frame as blank and sets the proper cmpstr values for fileds in
      'the key frame (i.e. the range selection fields)
  '*** set the default values for PpTyp and Sort Option from the previous entry
  '*** Disable All command buttons
  '*** Hide all frames except Frame FraNKeyAll
    ADC("wLmCoCd") = gs_CoCd
End Sub
Private Sub ADC_KeyValid(Cancel As Boolean, ErrMsg As String)
  '*** Check whether the PpTyp and the Sort Option is valid
      '(Sort option can be blank in which case the default sorting is on PpLoc, PpSubPrc, PpPrdCtg, PpSetTyp, PpDmCd, PpPts)
 
  '*** Checking Key values in Key Valid as they are copied in Init
 ' Call atxt_FldChange(0, UCase("wPpTyp"), Cancel, ErrMsg, ADC("wPpTyp"))
End Sub
Private Sub ADC_NonKeyWhen(pr_Cancel As Boolean, pr_ErrMsg As String)
  '*** Store the PpTyp, Sort Option for the next entry
  '*** Enable all the command buttons
  
  
 ' mPpSrt = ADC("wPpSrt")
  
 ' Call EnaDisaCmds(False)
End Sub

Private Sub ATXT_FldWhen(Index As Integer, ByVal IdName As String, Cancel As Boolean, ErrMsg As String)
'*** Set Helps For Each Field
  
  Select Case IdName
'*** Fields of The Key Frame
'uma ****  P replaced by R
  Case Is = UCase("wLmLocCdFr"), UCase("wLmLocCdTo")
    Call HlpList.Loc(ADC("wLmCoCd"), "'R'", , , False)
  Case Is = UCase("wLmSrt")
    Call HlpList.PMCd("SRTPP")
  Case Is = UCase("wLmDmCtgFr"), UCase("wLmDmCtgTo")
    Call HlpList.PMCd("DMCTG", , , , False)
  Case Is = UCase("wLmDmCdFr"), UCase("wLmDmCdTo")
    Call HlpList.DmCd("DM")
  Case Is = UCase("wLmWorkCdFr"), UCase("wLmWorkCdTo")
    Call HlpList.vPMCd(ADC("wLmCoCd"), "WORK", , , False)
  Case Is = UCase("wLmRmCtgFr"), UCase("wLmRmCtgTo")
    Call HlpList.PMCd("RMCTG")
  End Select
End Sub
Private Sub atxt_FldChange(Index As Integer, ByVal IdName As String, Cancel As Boolean, ErrMsg As String, ByVal pv_NewValue As Variant)
  '*** wLmSrt       Blank or valid PMCd('SRTPP')
  
  Select Case UCase(IdName)
  Case Is = UCase("wLmSrt")
    If pv_NewValue <> "" Then
      Cancel = Not moCn.RecSeek("Select PMCd From Param where PTyp= 'SRTPP' " + _
                           "and PMCd = '" + pv_NewValue + "'")
      ErrMsg = "Invalid Sort Option": Exit Sub
    End If
  End Select
End Sub

Private Sub CmdOk_Click()
  Dim wRow
  With GrdLm
    wRow = IIF(.Row > 0, .Row, 0)
    '*** TXTDESC(0) is working field for RmDesc
    '*** TXTDESC(1) is working field for RmIntQly
    If wRow > 0 Then
      If .Value(wRow, "LmRemarks") <> TXTREMARKS(0).TxtText Then _
         .Value(wRow, "LmRemarks") = TXTREMARKS(0).TxtText
    End If
  End With
  
  FraNKeyAll.Enabled = True
  GrdLm.Col = 1
  GrdLm.SetFocus
  ADC.AllowSave = True

End Sub
Private Sub TXTRemarks_FldWhen(Index As Integer, ByVal IdName As String, Cancel As Boolean, ErrMsg As String)
  '*** When any of the working fields of FraRmDesc frame gets the focus then disable
      'all the other frames and also the ADC
  ADC.AllowSave = False
  FraNKeyAll.Enabled = False
End Sub
Private Sub GrdLm_ColWhen(ByVal RowNum As Integer, ByVal ColName As String, Cancel As Boolean, ErrMsg As String, ByVal Action As MwfCtl.en_fgAction)
  '*** Set Helps For Each Field
  With GrdLm
    Select Case ColName
    Case Is = UCase("LmLocCd")
    'uma **** P replaced by R
      Call HlpList.Loc(ADC("wLmCoCd"), "'R'")
    Case Is = UCase("LmDmCtg")
      Call HlpList.PMCd("DMCTG")
    Case Is = UCase("LmWorkCd")
      Call HlpList.vPMCd(ADC("wLmCoCd"), "WORK")
    Case Is = UCase("LmDmCd")
      Call HlpList.DmCd("DM")
    Case Is = UCase("LmRmCtg")
      Call HlpList.PMCd("RMCTG")
    Case Is = UCase("LmValidYn")
      Call HlpList.PMCd("YN")
      If GrdLm.Value(RowNum, "LmValidYn") = "" Then
        GrdLm.Value(RowNum, "LmValidYn") = "Y"
      End If
    End Select
  End With
End Sub
Private Sub GrdLm_ReCalc(ByVal RowNum As Integer, ByVal ColName As String, ByVal OldVal As Variant, ByVal SourceName As String, ByVal SourceColName As String)
 
'  Select Case ColName
'  Case Is = UCase("PpSubPrc")
'    If GrdLm.Mode = fgmnorm Then
'      GrdLm.Value(RowNum, "PpSubPrc") = ""
'    End If
'  End Select
End Sub
Private Sub GrdLm_ColChange(ByVal RowNum As Integer, ByVal ColName As String, ByVal NewValue As Variant, Cancel As Boolean, ErrMsg As String)
  '*** LmLoc     valid Loc(LocTyp= 'P',LoCCoCd=LmCoCd)
  '*** LmDmCtg   Blank or valid PMCD("DMCTG")
  '*** LmDmCd    Blank or  valid DmCd("DM")  From DsgMst
  '*** LmWorkCd  Blank or valid vPMCd(LmCoCd,'WORK')
  '*** LmLossPer     >0
  
  With GrdLm
    Select Case UCase(ColName)
    Case Is = UCase("LmLocCd")
    'uma **** P replaced by R
      Cancel = (Not moCn.RecSeek("Select LocCd From Loc where " + _
                                    "LocTyp= 'R' and LocCd = '" + NewValue + "' "))
      If Cancel = True Then ErrMsg = "Invalid Process": Exit Sub
    Case Is = UCase("LmDmCtg")
      Cancel = (NewValue <> "") And (Not moCn.RecSeek("Select PmCd From Param where PTyp= 'DMCTG' " + _
                                    "and PmCd = '" + NewValue + "'"))
      If Cancel = True Then ErrMsg = "Invalid Design Category": Exit Sub
    Case Is = UCase("LmDmCd")
      If GrdLm.Value(RowNum, "LmDmCtg") = "" And NewValue <> "" Then _
        Cancel = True: ErrMsg = "Design Code Cannot Be Enter when Dsg Ctg Blank": Exit Sub
        Cancel = (NewValue <> "") And (Not moCn.RecSeek("Select DmCd From DsgMst Where " + IIF(gs_Partition = ctCurrPrtn, " DmPrtKey='" + ctCurrPrtn + "' And ", "") + " DmTcTyp= 'DM' " + _
                           "and DmCd = '" + NewValue + "' and DmSz= '' and DmCtg='" + GrdLm.Value(RowNum, "LmDmCtg") + "'"))
      If Cancel = True Then ErrMsg = "Invalid Design Code": Exit Sub
    Case Is = UCase("LmWorkCd")
      If GrdLm.Value(RowNum, "LmDmCtg") = "" And NewValue <> "" Then _
        Cancel = True: ErrMsg = "Worker Cannot Be Enter when Dsg Category Blank": Exit Sub
      Cancel = (NewValue <> "") And (Not moCn.RecSeek("Select vPmCd From vParam  " + _
                                        "where vPTyp= 'WORK' and vPCoCd='" + ADC("wLmCoCd") + "' " + _
                                                       "and vPmCd = '" + NewValue + "'"))
      If Cancel = True Then ErrMsg = "Invalid Worker": Exit Sub
    Case Is = UCase("LmLossPer")
      If NewValue <= 0 Then Cancel = True: ErrMsg = "Loss Percentage Should Be Greater Than Zero": Exit Sub
    Case Is = UCase("LmRmCtg")
      Cancel = (Not moCn.RecSeek("Select PmCd From Param where PTyp= 'RMCTG' " + _
                                    "and PmCd = '" + NewValue + "'"))
      If Cancel = True Then ErrMsg = "Invalid Raw Material Category": Exit Sub
      ' ****** Manali 3.03 - 20/06/08 - 'L' included
      Cancel = UCase$(NewValue) <> "G" And UCase$(NewValue) <> "P" And UCase$(NewValue) <> "S" And UCase$(NewValue) <> "L"  ' ****** sachin 2.12 - Silver added
      If Cancel = True Then ErrMsg = "Only 'G','P', 'S' or 'L' RM Categories Are Allowed": Exit Sub
    Case Is = UCase("LmValidYn")
      Cancel = (Not moCn.RecSeek("Select PMCd From Param where " + _
                                      "PTyp= 'YN' and PMCd = '" + NewValue + "'"))
        If Cancel = True Then ErrMsg = "Invalid Option. Enter Y/N.": Exit Sub
    Case Is = UCase("LmToRmWt")
        If NewValue <> 0 And NewValue < GrdLm.Value(RowNum, "LmFrRmWt") Then
          Cancel = True
          ErrMsg = "To Rm Wt. cannot be less than Fr Rm Wt.": Exit Sub
        End If
    End Select
  End With
End Sub
Private Sub GrdLm_InitKey(ByVal RowNum As Integer)
  '*** Initialise Key Fields
  With GrdLm
    .Store "LmCoCd", ADC("wLmCoCd")
  End With
End Sub
Private Sub ADC_SetRecSource()
  '*** Since this is an 'A' Entry, there is no RecSource associated with ADC. So set the
      'DataMode property of ADC control as xdatUnbound
  ADC.DataMode = xdatUnbound
End Sub

Private Sub GrdLm_RowWhen(ByVal RowNum As Integer)
  If GrdLm.Value(RowNum, "LmRemarks") = "" Then
    GrdLm.Value(RowNum, "LmRemarks") = "-"
  End If
  If GrdLm.Value(RowNum, "LmValidYn") = "" Then
    GrdLm.Value(RowNum, "LmValidYn") = "Y"
  End If

  ADC("wLmRemarks") = GrdLm.Value(RowNum, "LmRemarks")
End Sub

Private Sub GrdLm_SetRecSource()
  '*** Set the Record Source of the Grid GrdLm
  Dim ws_Cnd As String, ws_OrdBy As String
  ws_Cnd = ADC.RepCond
  ws_Cnd = IIF(ws_Cnd <> "", " where " + ws_Cnd, "")
    ws_OrdBy = " Order By LmCoCd,LmLocCd, LmDmCtg,LmDmCd,LmWorkCd,LmLossPer,LmValidYn "
  GrdLm.RecSource = " Select * from LossMst " + ws_Cnd + ws_OrdBy
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
  Cancel = atxt(Index).Validate
End Sub
Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    If UnloadMode <> vbFormCode Then Cancel = True: Exit Sub
    
    Set moCn = Nothing    '*** (09/08/05)    '*** (Bef speed) bef 24/11/06
End Sub
Private Sub GrdLm_Validate(Cancel As Boolean)
  Cancel = GrdLm.Validate
End Sub
