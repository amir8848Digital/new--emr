VERSION 5.00
Object = "{AB2CC32A-E652-42A2-B341-26ECE02F6386}#2.0#0"; "MwfCtl.ocx"
Begin VB.Form EmrFrmIDE 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "IDE User Rights"
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
   Begin MwfCtl.MWCTL_BTN1 CmdCP 
      Height          =   465
      Left            =   6570
      TabIndex        =   14
      ToolTipText     =   "Copy Option"
      Top             =   9270
      Width           =   795
      _ExtentX        =   1402
      _ExtentY        =   820
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
   Begin VB.Frame FraCP 
      Height          =   1245
      Left            =   870
      TabIndex        =   23
      Top             =   7950
      Visible         =   0   'False
      Width           =   12180
      Begin MwfCtl.MWCTL_BTN1 CmdCPGo 
         Height          =   405
         Left            =   11505
         TabIndex        =   13
         Top             =   465
         Width           =   495
         _ExtentX        =   873
         _ExtentY        =   714
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
         Index           =   36
         Left            =   1635
         TabIndex        =   8
         ToolTipText     =   "Enter Source User Group"
         Top             =   465
         Width           =   1170
         _ExtentX        =   2064
         _ExtentY        =   503
         MaxLength       =   5
         DataType        =   4
         IdName          =   "WCPYSRCUSRGRP"
      End
      Begin MwfCtl.MWCTL_MED ATXT 
         Height          =   285
         Index           =   6
         Left            =   4605
         TabIndex        =   9
         ToolTipText     =   "Enter From Control Type"
         Top             =   465
         Width           =   885
         _ExtentX        =   1561
         _ExtentY        =   503
         MaxLength       =   5
         DataType        =   4
         IdName          =   "WCPYCTLTYPFR"
         CmpStr          =   "UrCtlTyp>="
      End
      Begin MwfCtl.MWCTL_MED ATXT 
         Height          =   285
         Index           =   8
         Left            =   4605
         TabIndex        =   10
         ToolTipText     =   "Enter To Control Type"
         Top             =   750
         Width           =   885
         _ExtentX        =   1561
         _ExtentY        =   503
         MaxLength       =   5
         DataType        =   4
         IdName          =   "WCPYCTLTYPTO"
         CmpStr          =   "UrCtlTyp<="
      End
      Begin MwfCtl.MWCTL_MED ATXT 
         Height          =   285
         Index           =   9
         Left            =   6855
         TabIndex        =   11
         ToolTipText     =   "Enter From Control Sub Type"
         Top             =   465
         Width           =   4485
         _ExtentX        =   7911
         _ExtentY        =   503
         MaxLength       =   30
         DataType        =   4
         IdName          =   "WCPYCTLSTYPFR"
         CmpStr          =   "UrCtlSTyp>="
      End
      Begin MwfCtl.MWCTL_MED ATXT 
         Height          =   285
         Index           =   10
         Left            =   6855
         TabIndex        =   12
         ToolTipText     =   "Enter To Control Sub Type"
         Top             =   750
         Width           =   4485
         _ExtentX        =   7911
         _ExtentY        =   503
         MaxLength       =   30
         DataType        =   4
         IdName          =   "WCPYCTLSTYPTO"
         CmpStr          =   "UrCtlSTyp<="
      End
      Begin VB.Label ALBL 
         BackStyle       =   0  'Transparent
         Caption         =   "SubTyp Fr"
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
         Index           =   9
         Left            =   5850
         TabIndex        =   33
         Top             =   465
         Width           =   1065
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
         Index           =   8
         Left            =   5850
         TabIndex        =   32
         Top             =   750
         Width           =   1065
      End
      Begin VB.Label ALBL 
         BackStyle       =   0  'Transparent
         Caption         =   "Control Typ Fr"
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
         Left            =   3180
         TabIndex        =   31
         Top             =   465
         Width           =   1395
      End
      Begin VB.Label ALBL 
         BackStyle       =   0  'Transparent
         Caption         =   "                   To"
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
         Left            =   3180
         TabIndex        =   30
         Top             =   750
         Width           =   1455
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
         Index           =   7
         Left            =   -30
         TabIndex        =   25
         Top             =   90
         Width           =   12240
      End
      Begin VB.Label LblGrpOpt 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Source Usr Grp"
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
         Left            =   60
         TabIndex        =   24
         Top             =   465
         Width           =   1500
      End
   End
   Begin MwfCtl.MWCTL_HLP MWCTL_HLP1 
      Height          =   435
      Left            =   4290
      TabIndex        =   15
      TabStop         =   0   'False
      Top             =   9300
      Visible         =   0   'False
      Width           =   435
      _ExtentX        =   767
      _ExtentY        =   767
   End
   Begin MwfCtl.MWCTL_ADC ADC 
      CausesValidation=   0   'False
      Height          =   555
      Left            =   45
      TabIndex        =   16
      Top             =   9285
      Width           =   13860
      _ExtentX        =   24448
      _ExtentY        =   979
      AllowAdd        =   0   'False
      AllowDelete     =   0   'False
      BtnWidth        =   885
      BtnHeight       =   435
      BtnMaskColor    =   65535
      FontSz          =   8.25
   End
   Begin VB.Frame FRA_KEYS 
      BorderStyle     =   0  'None
      Height          =   660
      Left            =   -36
      TabIndex        =   18
      Top             =   30
      Width           =   14280
      Begin MwfCtl.MWCTL_MED ATXT 
         Height          =   285
         Index           =   7
         Left            =   915
         TabIndex        =   0
         ToolTipText     =   "Enter TC Type"
         Top             =   0
         Width           =   750
         _ExtentX        =   1323
         _ExtentY        =   503
         MaxLength       =   3
         DataType        =   4
         IdName          =   "WTCTYP"
      End
      Begin MwfCtl.MWCTL_MED ATXT 
         Height          =   285
         Index           =   5
         Left            =   13950
         TabIndex        =   6
         ToolTipText     =   "Sort Records On"
         Top             =   0
         Visible         =   0   'False
         Width           =   300
         _ExtentX        =   529
         _ExtentY        =   503
         Enabled         =   0   'False
         MaxLength       =   8
         DataType        =   4
         IdName          =   "WIDESRT"
      End
      Begin MwfCtl.MWCTL_MED ATXT 
         Height          =   285
         Index           =   1
         Left            =   3195
         TabIndex        =   1
         ToolTipText     =   "Enter User Group"
         Top             =   0
         Width           =   1140
         _ExtentX        =   2011
         _ExtentY        =   503
         MaxLength       =   5
         DataType        =   4
         IdName          =   "WUSRGRP"
      End
      Begin MwfCtl.MWCTL_MED ATXT 
         Height          =   285
         Index           =   0
         Left            =   6135
         TabIndex        =   2
         ToolTipText     =   "Enter From Control Type"
         Top             =   0
         Width           =   885
         _ExtentX        =   1561
         _ExtentY        =   503
         MaxLength       =   5
         DataType        =   4
         IdName          =   "WCTLTYPFR"
         CmpStr          =   "UrCtlTyp>="
      End
      Begin MwfCtl.MWCTL_MED ATXT 
         Height          =   285
         Index           =   3
         Left            =   6135
         TabIndex        =   3
         ToolTipText     =   "Enter To Control Type"
         Top             =   285
         Width           =   885
         _ExtentX        =   1561
         _ExtentY        =   503
         MaxLength       =   5
         DataType        =   4
         IdName          =   "WCTLTYPTO"
         CmpStr          =   "UrCtlTyp<="
      End
      Begin MwfCtl.MWCTL_MED ATXT 
         Height          =   285
         Index           =   2
         Left            =   8385
         TabIndex        =   4
         ToolTipText     =   "Enter From Control Sub Type"
         Top             =   0
         Width           =   4485
         _ExtentX        =   7911
         _ExtentY        =   503
         MaxLength       =   30
         DataType        =   4
         IdName          =   "WCTLSTYPFR"
         CmpStr          =   "UrCtlSTyp>="
      End
      Begin MwfCtl.MWCTL_MED ATXT 
         Height          =   285
         Index           =   4
         Left            =   8385
         TabIndex        =   5
         ToolTipText     =   "Enter To Control Sub Type"
         Top             =   285
         Width           =   4485
         _ExtentX        =   7911
         _ExtentY        =   503
         MaxLength       =   30
         DataType        =   4
         IdName          =   "WCTLSTYPTO"
         CmpStr          =   "UrCtlSTyp<="
      End
      Begin VB.Label ALBL 
         BackStyle       =   0  'Transparent
         Caption         =   "SubTyp Fr"
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
         Left            =   7380
         TabIndex        =   29
         Top             =   0
         Width           =   1065
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
         Index           =   0
         Left            =   7380
         TabIndex        =   28
         Top             =   285
         Width           =   1065
      End
      Begin VB.Label ALBL 
         BackStyle       =   0  'Transparent
         Caption         =   "Control Typ Fr"
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
         Left            =   4710
         TabIndex        =   27
         Top             =   0
         Width           =   1395
      End
      Begin VB.Label ALBL 
         BackStyle       =   0  'Transparent
         Caption         =   "                   To"
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
         Left            =   4710
         TabIndex        =   26
         Top             =   285
         Width           =   1455
      End
      Begin VB.Label LblUsrGrp 
         BackStyle       =   0  'Transparent
         Caption         =   "User Group"
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
         Left            =   2070
         TabIndex        =   22
         Top             =   0
         Width           =   1125
      End
      Begin VB.Label ALBL 
         BackStyle       =   0  'Transparent
         Caption         =   "Sort On"
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
         Left            =   13170
         TabIndex        =   21
         Top             =   0
         Visible         =   0   'False
         Width           =   795
      End
      Begin VB.Label ALBL 
         BackStyle       =   0  'Transparent
         Caption         =   "Tc Type"
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
         Left            =   120
         TabIndex        =   20
         Top             =   0
         Width           =   795
      End
   End
   Begin VB.Frame FRA_NKEYS 
      BorderStyle     =   0  'None
      Enabled         =   0   'False
      Height          =   9570
      Left            =   -90
      TabIndex        =   17
      Top             =   0
      Width           =   15075
      Begin VB.Frame FraNKeyAll 
         BorderStyle     =   0  'None
         Height          =   9225
         Left            =   150
         TabIndex        =   19
         Top             =   90
         Width           =   14010
         Begin MwfCtl.MWCTL_FLX GrdIDE 
            Height          =   8460
            Left            =   0
            TabIndex        =   7
            Top             =   630
            Width           =   13845
            _ExtentX        =   24421
            _ExtentY        =   14923
            Cols            =   6
            colname1        =   "URCTLTYP"
            heading1        =   "Control Typ"
            datafld1        =   "UrCtlTyp"
            datatype1       =   4
            recalcparent1   =   "GrdIDE(URCTLSTYP),GrdIDE(URNAME),GrdIDE(URDESC)"
            maxlength1      =   5
            tooltiptext1    =   "Enter Control Type"
            colname2        =   "URCTLSTYP"
            heading2        =   "Sub Type                                       "
            datafld2        =   "UrCtlSTyp"
            datatype2       =   4
            recalcon2       =   "GrdIDE(URCTLTYP)"
            recalcparent2   =   "GrdIDE(URNAME),GrdIDE(URDESC)"
            maxlength2      =   30
            tooltiptext2    =   "Enter Control Sub Type (eg: Grid Name)"
            colname3        =   "URNAME"
            heading3        =   "Name                                                            "
            datafld3        =   "UrName"
            datatype3       =   4
            recalcon3       =   "GrdIDE(URCTLTYP),GrdIDE(URCTLSTYP)"
            recalcparent3   =   "GrdIDE(URDESC)"
            maxlength3      =   30
            tooltiptext3    =   "Enter Control Name"
            colname4        =   "URSTYLE"
            heading4        =   "Style"
            datafld4        =   "UrStyle"
            datatype4       =   4
            maxlength4      =   1
            tooltiptext4    =   "Enter Display Style"
            colname5        =   "URDESC"
            heading5        =   "Desc                                                                                      "
            datafld5        =   "UrDesc"
            recalcon5       =   "GrdIDE(URCTLTYP),GrdIDE(URCTLSTYP),GrdIDE(URNAME)"
            maxlength5      =   45
            style5          =   2
            tooltiptext5    =   "Control Description "
         End
      End
   End
End
Attribute VB_Name = "EmrFrmIDE"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'---- General Comments on the Entry --------'
' X-'', A-UsrRight
'--------------------
Option Explicit
Dim moCn As MwfLib.MDOConnection
Dim mb_ComLoc As Boolean

Enum en_IdeFra
  Cpy = 0
End Enum

Private Sub ADC_Load()
  '*** Set the First Key Control property to the working field wTcTyp
  '*** Set the First Non Key Control property to the grid GrdIDE
  '*** Set the Child Property Of ADC to GrdIDE
  '*** Set the Previous and Next Control Properties for the Grid GrdIDE
  '*** Set the Hot Key for the Grid GrdIDE as 'A'
  
  Set moCn = ADC.Connection
  Set ADC.FirKeyCtl = ADC("wTcTyp")
  Set ADC.FirNKeyCtl = GrdIDE
      ADC.Child = "GrdIDE"
  Set GrdIDE.PrevCtl = ADC
  Set GrdIDE.NextCtl = ADC
      GrdIDE.HotKey = "A"
 
  Select Case UCase$(ADC.MenuCd)
  Case Is = UCase("IdeDfn")
    LblUsrGrp.Visible = False
    ADC("wUsrGrp").Visible = False
    GrdIDE.ColProp("UrDesc").Style = fgcNormal
    GrdIDE.ColProp("UrStyle").Heading = "Default Style"
    CmdCP.Visible = False
  Case Is = UCase("IdeUsrRight")
    LblUsrGrp.Visible = True
    ADC("wUsrGrp").Visible = True
    GrdIDE.ColProp("UrDesc").Style = fgcReadOnly
    GrdIDE.ColProp("UrStyle").Heading = "Style"
  End Select
  GrdIDE.SetColWidthToHd
   
End Sub
Private Sub ADC_KeyWhen()
  '*** Bring the frame FraNKeyAll to the top
  FraNKeyAll.ZOrder (0)
  
  Call CmpString("K")     ' to set the Compare String Property of the ADC controls when the Focus is in the Key Frame
  Call EnaDisaCmds(True)
  Call HideAllFras
End Sub
Private Sub ADC_KeyValid(Cancel As Boolean, ErrMsg As String)
  '*** Check if the Tc Type, Usr Grp and the Sort Order selected is valid (Sort Option Can be '')
  
  Call atxt_FldChange(0, UCase("wTcTyp"), Cancel, ErrMsg, ADC("wTcTyp"))
  If Cancel = True Then Exit Sub
  
  Call atxt_FldChange(0, UCase("wUsrGrp"), Cancel, ErrMsg, ADC("wUsrGrp"))
  If Cancel = True Then Exit Sub
  
  'Call atxt_FldChange(0, UCase("wIdeSrt"), Cancel, ErrMsg, ADC("wIdeSrt"))
  'If Cancel = True Then Exit Sub
End Sub
Private Sub ADC_NonKeyWhen(pr_Cancel As Boolean, pr_ErrMsg As String)
  Call EnaDisaCmds(False)
End Sub
Private Sub ATXT_FldWhen(Index As Integer, ByVal IdName As String, Cancel As Boolean, ErrMsg As String)
  '*** Set Helps For Each Field
  
  Select Case IdName
  Case Is = UCase("wTcTyp")
    Call HlpList.PMCd("TCTYP", , "OM")
  Case Is = UCase("wUsrGrp"), UCase("wCpySrcUsrGrp")
    Call HlpList.PMCd("USRGRP")
  Case Is = UCase("wCtlTypFr"), UCase("wCtlTypTo"), UCase("wCpyCtlTypFr"), UCase("wCpyCtlTypTo")
    Call HlpList.PMCd("CTRLTYP")
  Case Is = UCase("wCtlSTypFr")
    Call HlpList.CtrlSubTyp(ADC("wTcTyp"), ADC("wCtlTypFr"))
  Case Is = UCase("wCtlSTypTo")
    Call HlpList.CtrlSubTyp(ADC("wTcTyp"), ADC("wCtlTypTo"))
  Case Is = UCase("wCpyCtlSTypFr")
    Call HlpList.CtrlSubTyp(ADC("wTcTyp"), ADC("wCpyCtlTypFr"))
  Case Is = UCase("wCpyCtlSTypTo")
    Call HlpList.CtrlSubTyp(ADC("wTcTyp"), ADC("wCpyCtlTypTo"))
  
  'Case Is = UCase("wIdeSrt")
  '  Call HlpList.PMCd("SRTIDE")
  End Select
End Sub
Private Sub atxt_FldChange(Index As Integer, ByVal IdName As String, Cancel As Boolean, ErrMsg As String, ByVal pv_NewValue As Variant)
  '*** wTcTyp                    valid PMCd('TCTYP') of PValue1= 'OM'
  '*** wUsrGrp, wCpySrcUsrGrp    valid PMCd('USRGRP')
  ''*** wIdeSrt                  valid PMCd('SRTIDE') or ''
  
  Select Case IdName
  Case Is = UCase("wTcTyp")
    Cancel = Not moCn.RecSeek("Select PMCd From Param where PTyp= 'TCTYP' and " + _
                              "PMCd = '" + pv_NewValue + "' and PValue1= 'OM'")
    If Cancel = True Then ErrMsg = "Invalid Tc Type": Exit Sub
  Case Is = UCase("wUsrGrp")
    Cancel = (pv_NewValue <> "") And (Not moCn.RecSeek("Select PMCd From Param where PTyp= 'USRGRP' and " + _
                         "PMCd = '" + pv_NewValue + "'"))
    
    If ADC("wUsrGrp").Visible = True And pv_NewValue = "" Then Cancel = True
    If Cancel = True Then ErrMsg = "Invalid User Group": Exit Sub

  Case Is = UCase("wCpySrcUsrGrp")
    Cancel = (pv_NewValue <> "") And (Not moCn.RecSeek("Select PMCd From Param where PTyp= 'USRGRP' and " + _
                         "PMCd = '" + pv_NewValue + "'"))
    
    If Cancel = True Then: ErrMsg = "Invalid Source User Group": Exit Sub


'  Case Is = UCase("wIdeSrt")
'    Cancel = (pv_NewValue <> "") And (Not moCn.RecSeek("Select PMCd From Param Where " + _
'                                        "PTyp = 'SRTIDE' And PMCd = '" + pv_NewValue + "'"))
'    If Cancel = True Then ErrMsg = "Invalid Sort Option": Exit Sub
  End Select
End Sub
Private Sub atxt_ReCalculate(Index As Integer, ByVal IdName As String, ByVal OldVal As Variant, ByVal RowNum As Integer, ByVal SourceName As String, ByVal SourceColName As String)
  'Select Case IdName
  'Case Is = UCase("wLocPrntYn")
  '  If ADC("wLocCoCd") <> ctSelfCoCd Then ADC("wLocPrntYN") = "N"
  'End Select
End Sub
Private Sub GrdIDE_RowWhen(ByVal RowNum As Integer)
  If GrdIDE.IsNew(RowNum) = True Then
    If RowNum > 1 Then
      GrdIDE.Value(RowNum, "UrCtlTyp") = GrdIDE.Value(RowNum - 1, "UrCtlTyp")
      GrdIDE.Value(RowNum, "UrCtlSTyp") = GrdIDE.Value(RowNum - 1, "UrCtlSTyp")
    End If
  End If
End Sub
Private Sub GrdIDE_ColWhen(ByVal RowNum As Integer, ByVal ColName As String, Cancel As Boolean, ErrMsg As String, ByVal Action As MwfCtl.en_fgAction)
  If (Not GrdIDE.IsNew(RowNum)) And _
    (ColName = UCase("UrCtlTyp") Or ColName = UCase("UrCtlSTyp") Or ColName = UCase("UrName")) Then
    Cancel = True: ErrMsg = "Cannot Edit Type/ Sub Type/ Name Of Exising Record"
  End If
  
  '*** Set Helps For Each Field
  Select Case ColName
  Case Is = UCase("UrCtlTyp")
    Call HlpList.PMCd("CTRLTYP")
  Case Is = UCase("UrCtlSTyp")
    'If GrdIDE.Value(RowNum, "UrCtlTyp") <> "GRD" Then Cancel = True: ErrMsg = "Cannot Edit This Field": Exit Sub
    If Not (GrdIDE.Value(RowNum, "UrCtlTyp") = "GRD" Or GrdIDE.Value(RowNum, "UrCtlTyp") = "X") Then Cancel = True: ErrMsg = "Cannot Edit This Field": Exit Sub
    'If ADC("wUsrGrp") <> "" Then
    Call HlpList.CtrlSubTyp(ADC("wTcTyp"), GrdIDE.Value(RowNum, "UrCtlTyp"))
  Case Is = UCase("UrName")
    If ADC("wUsrGrp") <> "" Then Call HlpList.CtrlName(ADC("wTcTyp"), GrdIDE.Value(RowNum, "UrCtlTyp"), _
                                      GrdIDE.Value(RowNum, "UrCtlSTyp"))
  Case Is = UCase("UrStyle")
    Call HlpList.PMCd("IDE")
End Select
End Sub
Private Sub GrdIDE_ColChange(ByVal RowNum As Integer, ByVal ColName As String, ByVal NewValue As Variant, Cancel As Boolean, ErrMsg As String)
  '*** UrCtlTyp
      'valid PMCd('CTRLTYP')
  '*** UrCtlSTyp
      'If UsrGrp= '' then <> Blank
      'If UsrGrp<> '' then Blank or valid UrCtlSTyp From UsrRight (UrTcTyp= TcTyp, UrGrp= '', UrCtlTyp= CtlTyp)
  '*** UrName
      'If UsrGrp= '' then <> Blank
      'If UsrGrp<> '' then Blank or valid UrName From UsrRight (UrTcTyp= TcTyp, UrGrp= '', UrCtlTyp= CtlTyp, UrCtlSTyp= CtlSubTyp)
  '*** UrStyle
      'valid PMCd('IDE')
  
      
  With GrdIDE
    Select Case ColName
    Case Is = UCase("UrCtlTyp")
      If Not moCn.RecSeek("Select 'x' From Param where PTyp= 'CTRLTYP' and PMCd= '" + NewValue + "' ") Then
        Cancel = True: ErrMsg = "Invalid Control Type": Exit Sub
      End If
    Case Is = UCase("UrCtlSTyp")
      If ADC("wUsrGrp") <> "" Then
        If (NewValue <> "") And (Not moCn.RecSeek("Select 'x' From UsrRight where UrTcTyp= '" + ADC("wTcTyp") + "' " + _
                            "and UrGrp= '' and UrCtlTyp= '" + .Value(RowNum, "UrCtlTyp") + "' " + _
                            "and UrCtlSTyp= '" + NewValue + "' ")) Then
          Cancel = True: ErrMsg = "Invalid Control Sub Type": Exit Sub
        End If
      End If
    Case Is = UCase("UrName")
      If ADC("wUsrGrp") <> "" Then
        If (NewValue <> "") And (Not moCn.RecSeek("Select 'x' From UsrRight where UrTcTyp= '" + ADC("wTcTyp") + "' " + _
                            "and UrGrp= '' and UrCtlTyp= '" + .Value(RowNum, "UrCtlTyp") + "' " + _
                            "and UrCtlSTyp= '" + .Value(RowNum, "UrCtlSTyp") + "' " + _
                            "and UrName= '" + NewValue + "' ")) Then
          Cancel = True: ErrMsg = "Invalid Control Name": Exit Sub
        End If
      End If
    Case Is = UCase("UrStyle")
      If .Value(RowNum, "UrName") <> "" And NewValue = "" Then Cancel = True: ErrMsg = "Style Cannot Be Blank": Exit Sub
      If (NewValue <> "") And Not moCn.RecSeek("Select 'x' From Param where PTyp= 'IDE' and PMCd= '" + NewValue + "' ") Then
        Cancel = True: ErrMsg = "Invalid IDE Style Option": Exit Sub
      End If
      If .Value(RowNum, "UrCtlTyp") = "CMD" And NewValue = "D" Then Cancel = True: ErrMsg = "IDE Style Option Cannot Be 'D'(Disable) For Button Control": Exit Sub
    End Select
  End With
End Sub
Private Sub GrdIDE_ReCalc(ByVal RowNum As Integer, ByVal ColName As String, ByVal OldVal As Variant, ByVal SourceName As String, ByVal SourceColName As String)
  With GrdIDE
    Select Case ColName
    Case Is = UCase("UrCtlSTyp")
      If .Mode = fgmnorm Then .Value(RowNum, "UrCtlSTyp") = ""
    Case Is = UCase("UrName")
      If .Mode = fgmnorm Then .Value(RowNum, "UrName") = ""
    Case Is = UCase("UrDesc")
      If .Mode = fgmnorm Then
        If .Value(RowNum, "UrName") <> "" Or UCase(.Value(RowNum, "UrCtlTyp")) = "GRD" Or _
           UCase(.Value(RowNum, "UrCtlTyp")) = "X" Then
          .Value(RowNum, "UrDesc") = moCn.GetFldVal("Select UrDesc From UsrRight Where UrTcTyp= '" + ADC("wTcTyp") + "' " + _
                                                    "and UrGrp= '' and UrCtlTyp= '" + .Value(RowNum, "UrCtlTyp") + "' " + _
                                                    "and UrCtlSTyp= '" + .Value(RowNum, "UrCtlSTyp") + "' " + _
                                                    "and UrName= '" + .Value(RowNum, "UrName") + "' ")
        Else
          .Value(RowNum, "UrDesc") = ""
        End If
      End If
    End Select
  End With
End Sub
Private Sub GrdIDE_InitKey(ByVal RowNum As Integer)
  '*** Initialise Key Fields
  With GrdIDE
    .Store "UrTcTyp", ADC("wTcTyp")
    .Store "UrGrp", ADC("wUsrGrp")
  End With
End Sub
Private Sub ADC_SetRecSource()
  '*** Since this is an 'A' Entry, there is no RecSource associated with ADC. So set the
      'DataMode property of ADC control as xdatUnbound
  ADC.DataMode = xdatUnbound
End Sub
Private Sub GrdIDE_SetRecSource()
  '*** Set the Record Source of the Grid GrdIDE
      'Default sort option is UrCtlTyp, UrCtlSTyp, UrName
  
  Dim wCnd  As String, wOrdBy As String
  wCnd = ADC.RepCond
  wCnd = IIF(wCnd <> "", " and " + wCnd, " ")
  
  wOrdBy = ""
  'wOrdBy = moCn.GetFldVal("Select PDesc225 from Param " + _
           "where PTyp= 'SRTIDE' and PMCd= '" + ADC("wIdeSrt") + "'")

  wOrdBy = IIF(wOrdBy = "", " Order By UrCtlTyp, UrCtlSTyp, UrName ", "Order By " + wOrdBy)

  GrdIDE.RecSource = "Select * from UsrRight where UrTcTyp= '" + ADC("wTcTyp") + "' " + _
                     "and UrGrp= '" + ADC("wUsrGrp") + "' " + wCnd + wOrdBy
End Sub
Private Sub CmdCP_Click()
  '*** Display or hide the Copy Frame depending on the toggle state
  Call DispFra(Cpy)
End Sub
Private Sub CmdCpGo_Click()
  Call CpyUsrRight
  Call CmdCP_Click
End Sub
Private Sub DispFra(ByVal pv_LocFra As en_LocFra)
  '*** This is the common procedure which is used to toggle between hiding and displaying
      'the frames (i.e. if the frame is visible then make it invisible and viseversa)
      'associated with a command button

  Select Case pv_LocFra
  Case Is = Cpy
    If FraCP.Visible = True Then
      FraCP.Visible = False
      FraCP.Enabled = False
      Call EnaDisaCmds(False, CmdCP)
      CmdCP.SetFocus
    Else
      FraCP.Visible = True
      FraCP.Enabled = True
      ADC("wCpySrcUsrGrp").SetFocus
      FraCP.ZOrder
      Call EnaDisaCmds(True, CmdCP)
    End If
  End Select
  
End Sub
Private Sub EnaDisaCmds(ByVal pv_ShowFra As Boolean, Optional ByVal pv_CmdBut As MWCTL_BTN1, Optional ByVal pv_DispFra As en_LocFra)
  
  If Not pv_CmdBut Is Nothing Then
    If pv_ShowFra = True Then pv_CmdBut.CausesValidation = True
  End If
  FraNKeyAll.Enabled = Not pv_ShowFra
  ADC.AllowSave = Not pv_ShowFra
  CmdCP.Enabled = Not pv_ShowFra
  If Not pv_CmdBut Is Nothing Then pv_CmdBut.Enabled = True
End Sub
Private Sub HideAllFras()
  '*** This procedure hides all the frames other than FraNKeyAll
  FraNKeyAll.Visible = True
  FraCP.Visible = False
  FraCP.Enabled = False
End Sub
Private Sub CmpString(ByVal mTyp As String)
  '*** Sets the CmpStr properties for the fields in the Key Frame and the Copy Frame
  '****** Seting values for CmpStr of 'K'ey fields, for Adc.RepCnd (Making Copy fields CmpStr = '')
  If mTyp = "K" Then
    'ADC("wTcTyp").CmpStr = "UrTcTyp="
    'ADC("wUsrGrp").CmpStr = "UrGrp="
    
    ADC("wCtlTypFr").CmpStr = "UrCtlTyp>= "
    ADC("wCtlTypTo").CmpStr = "UrCtlTyp<= "
    ADC("wCtlSTypFr").CmpStr = "UrCtlSTyp>="
    ADC("wCtlSTypTo").CmpStr = "UrCtlSTyp<="
   
    ADC("wCpyCtlTypFr").CmpStr = ""
    ADC("wCpyCtlTypTo").CmpStr = ""
    ADC("wCpyCtlSTypFr").CmpStr = ""
    ADC("wCpyCtlSTypTo").CmpStr = ""
    
  '****** Seting values for CmpStr of 'C'opy fields, for Adc.RepCnd (Making Key fields CmpStr = '')
  ElseIf mTyp = "C" Then
    'ADC("wTcTyp").CmpStr = ""
    'ADC("wUsrGrp").CmpStr = ""
    
    ADC("wCtlTypFr").CmpStr = ""
    ADC("wCtlTypTo").CmpStr = ""
    ADC("wCtlSTypFr").CmpStr = ""
    ADC("wCtlSTypTo").CmpStr = ""
   
    ADC("wCpyCtlTypFr").CmpStr = "UrCtlTyp>= "
    ADC("wCpyCtlTypTo").CmpStr = "UrCtlTyp<= "
    ADC("wCpyCtlSTypFr").CmpStr = "UrCtlSTyp>="
    ADC("wCpyCtlSTypTo").CmpStr = "UrCtlSTyp<="
  End If
End Sub
Private Sub CpyUsrRight()
  '*** Copy Option will copy the User Rights of one User Group to Another for the given Prd Points from one PrdCtg to another PrdCtg (in case of 'GN' and 'WP')
      'and from one Setting Type to another (in case of 'WS', 'GS' and 'GP')
      'for the given PrdTyp and for the given Design Code, Loc range

  Dim ws_CpyCnd As String
  
  Call CmpString("C")     ' to set the Compare String Property of the ADC controls when the Copy Option Is clicked
  
  ws_CpyCnd = ADC.RepCond
  ws_CpyCnd = IIF(ws_CpyCnd <> "", " and " + ws_CpyCnd, "")
  
  If ADC("wCpySrcUsrGrp") = "" Then DispMsg "Source User Group Should Be Filled Up", etError: Exit Sub
    
  If moCn.RecSeek("Select 'x' From UsrRight Where UrTcTyp= '" + ADC("wTcTyp") + "' and UrGrp= '" + ADC("wUsrGrp") + "' " + ws_CpyCnd) Then
    DispMsg "Cannot Copy as User Rights Records Already Exist For The Specified Scope", etError
    Exit Sub
  End If
  
  '*** Beginning the copy process
  GrdIDE.StartCopy
  MWLib.BeginProcess Me, "Copying IDE User Rights ..."
  Dim wRsUsrRight As MwfLib.MDORowSet, wi_Row As Integer
  
  Set wRsUsrRight = moCn.OpenRes("Select * From UsrRight Where UrTcTyp= '" + ADC("wTcTyp") + "' and " + _
                    "UrGrp= '" + ADC("wCpySrcUsrGrp") + "' " + ws_CpyCnd + " Order By UrCtlTyp, UrCtlSTyp, UrName")
  
  With wRsUsrRight
    Do While Not (.EOF Or .BOF)
      MWLib.RunningProcess "Copying Control Type: " + !UrCtlTyp + _
                           IIF(!UrCtlSTyp <> "", " Sub Type: ", "") + !UrCtlSTyp + _
                           " Name: " + !UrName
      GrdIDE.AddItem
      
      wi_Row = GrdIDE.Rows - 1
      GrdIDE.Value(wi_Row, "UrCtlTyp") = !UrCtlTyp
      GrdIDE.Value(wi_Row, "UrCtlSTyp") = !UrCtlSTyp
      GrdIDE.Value(wi_Row, "UrName") = !UrName
      GrdIDE.Value(wi_Row, "UrDesc") = !UrDesc
      GrdIDE.Value(wi_Row, "UrStyle") = !UrStyle
      GrdIDE.SaveRec (wi_Row)
      .MoveNext
    Loop
  End With
  MWLib.EndProcess Me
  GrdIDE.EndCopy
  Set wRsUsrRight = Nothing
  '*** End of the copy process
End Sub


'''  ***** Std Code not to be changed **********************************
Private Sub Form_GotFocus()
  '*** Jenny Color (For Setting System Colors)
  SetMDISysColorsAfterLogin
  '*** Jenny Color (For Setting System Colors)
End Sub
Private Sub Form_Load()
  '*** Jenny Colour
  Dim ws_BtnStr As String
  ws_BtnStr = "CmdCPGo"
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
    
    Set moCn = Nothing
End Sub
Private Sub GrdIDE_Validate(Cancel As Boolean)
  Cancel = GrdIDE.Validate
End Sub

