VERSION 5.00
Object = "{AB2CC32A-E652-42A2-B341-26ECE02F6386}#2.1#0"; "MwfCtl.ocx"
Begin VB.Form EmrFrmLoc 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Location Master"
   ClientHeight    =   9810
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   15210
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MDIChild        =   -1  'True
   ScaleHeight     =   9810
   ScaleWidth      =   15210
   WindowState     =   2  'Maximized
   Begin MwfCtl.MWCTL_BTN1 CmdCP 
      Height          =   465
      Left            =   5220
      TabIndex        =   21
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
      Height          =   975
      Left            =   3120
      TabIndex        =   14
      Top             =   8250
      Visible         =   0   'False
      Width           =   5565
      Begin MwfCtl.MWCTL_BTN1 CmdCPGo 
         Height          =   405
         Left            =   4905
         TabIndex        =   15
         Top             =   435
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
      Begin MwfCtl.MWCTL_MED Atxt 
         Height          =   285
         Index           =   36
         Left            =   1035
         TabIndex        =   16
         ToolTipText     =   "Enter Voucher Tc"
         Top             =   465
         Width           =   1485
         _ExtentX        =   2619
         _ExtentY        =   503
         MaxLength       =   8
         DataType        =   4
         IdName          =   "WFRLOC"
         CmpStr          =   "LocCd>="
      End
      Begin MwfCtl.MWCTL_MED Atxt 
         Height          =   285
         Index           =   50
         Left            =   3345
         TabIndex        =   17
         ToolTipText     =   "Enter Voucher Tc"
         Top             =   465
         Width           =   1485
         _ExtentX        =   2619
         _ExtentY        =   503
         MaxLength       =   8
         DataType        =   4
         IdName          =   "WTOLOC"
         CmpStr          =   "LocCd<="
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
         TabIndex        =   20
         Top             =   90
         Width           =   4410
      End
      Begin VB.Label LblGrpOpt 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "From Loc"
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
         TabIndex        =   19
         Top             =   465
         Width           =   900
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "To Loc"
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
         Left            =   2610
         TabIndex        =   18
         Top             =   480
         Width           =   645
      End
   End
   Begin MwfCtl.MWCTL_HLP MWCTL_HLP1 
      Height          =   435
      Left            =   4290
      TabIndex        =   5
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
      Left            =   15
      TabIndex        =   6
      Top             =   9285
      Width           =   11130
      _ExtentX        =   19632
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
      Height          =   330
      Left            =   -36
      TabIndex        =   8
      Top             =   30
      Width           =   9420
      Begin MwfCtl.MWCTL_MED Atxt 
         Height          =   285
         Index           =   7
         Left            =   675
         TabIndex        =   0
         ToolTipText     =   "Enter Location Type"
         Top             =   0
         Width           =   540
         _ExtentX        =   953
         _ExtentY        =   503
         MaxLength       =   2
         DataType        =   4
         ReCalcParent    =   "WLOCPRNTYN"
         IdName          =   "WLOCTYP"
         CmpStr          =   "LocTyp="
      End
      Begin MwfCtl.MWCTL_MED Atxt 
         Height          =   285
         Index           =   5
         Left            =   7980
         TabIndex        =   3
         ToolTipText     =   "Sort Records On"
         Top             =   0
         Visible         =   0   'False
         Width           =   690
         _ExtentX        =   1217
         _ExtentY        =   503
         Enabled         =   0   'False
         MaxLength       =   8
         DataType        =   4
         IdName          =   "WLOCSRT"
      End
      Begin MwfCtl.MWCTL_MED Atxt 
         Height          =   285
         Index           =   0
         Left            =   6285
         TabIndex        =   2
         ToolTipText     =   "Enter Whether Parent Code Or Normal Codes to Display"
         Top             =   0
         Width           =   450
         _ExtentX        =   794
         _ExtentY        =   503
         DataType        =   4
         ReCalcOn        =   "WLOCTYP,WLOCCOCD"
         IdName          =   "WLOCPRNTYN"
         CmpStr          =   "LocPrntYN="
      End
      Begin MwfCtl.MWCTL_MED Atxt 
         Height          =   285
         Index           =   1
         Left            =   3315
         TabIndex        =   1
         ToolTipText     =   "Enter Location Company Code"
         Top             =   0
         Width           =   840
         _ExtentX        =   1482
         _ExtentY        =   503
         MaxLength       =   3
         DataType        =   4
         ReCalcParent    =   "WLOCPRNTYN"
         IdName          =   "WLOCCOCD"
         CmpStr          =   "LocCoCd="
      End
      Begin VB.Label LblCoCd 
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
         Height          =   285
         Left            =   1980
         TabIndex        =   13
         Top             =   0
         Width           =   1485
      End
      Begin VB.Label LblPrntYn 
         BackStyle       =   0  'Transparent
         Caption         =   "Parent Cd (Y/N)"
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
         Left            =   4770
         TabIndex        =   12
         Top             =   0
         Width           =   1485
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
         Left            =   7170
         TabIndex        =   11
         Top             =   0
         Visible         =   0   'False
         Width           =   915
      End
      Begin VB.Label ALBL 
         BackStyle       =   0  'Transparent
         Caption         =   "Type"
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
         Left            =   60
         TabIndex        =   10
         Top             =   0
         Width           =   795
      End
   End
   Begin VB.Frame FRA_NKEYS 
      BorderStyle     =   0  'None
      Enabled         =   0   'False
      Height          =   9480
      Left            =   -90
      TabIndex        =   7
      Top             =   0
      Width           =   11985
      Begin VB.Frame FraNKeyAll 
         BorderStyle     =   0  'None
         Height          =   9315
         Left            =   150
         TabIndex        =   9
         Top             =   90
         Width           =   19000
         Begin MwfCtl.MWCTL_FLX GrdLoc 
            Height          =   8850
            Left            =   0
            TabIndex        =   4
            Top             =   270
            Width           =   11865
            _ExtentX        =   20929
            _ExtentY        =   15610
            Cols            =   16
            colname1        =   "LOCCD"
            heading1        =   "Loc Cd               "
            datafld1        =   "LocCd"
            datatype1       =   4
            recalcparent1   =   "GrdLoc(LOCPRNTCD)"
            maxlength1      =   8
            tooltiptext1    =   "Enter Location Code"
            colname2        =   "LOCPRNTCD"
            heading2        =   "Prnt Cd              "
            datafld2        =   "LocPrntCd"
            datatype2       =   4
            recalcon2       =   "GrdLoc(LOCCD)"
            maxlength2      =   8
            style2          =   0
            tooltiptext2    =   "Enter Location Parent Code"
            colname3        =   "LOCPRDTYP"
            heading3        =   "PrdTyp "
            datafld3        =   "LocPrdTyp"
            datatype3       =   4
            maxlength3      =   2
            style3          =   0
            tooltiptext3    =   "Enter Location Production Type"
            colname4        =   "LOCLOSS"
            heading4        =   "Loss Loc           "
            datafld4        =   "LocLoss"
            datatype4       =   4
            maxlength4      =   8
            style4          =   0
            tooltiptext4    =   "Enter Default Loss Location"
            colname5        =   "LOCBRK"
            heading5        =   "Brk Loc             "
            datafld5        =   "LocBrk"
            datatype5       =   4
            maxlength5      =   8
            style5          =   0
            tooltiptext5    =   "Enter Default Break Location"
            colname6        =   "LOCMSG"
            heading6        =   "Msg Loc            "
            datafld6        =   "LocMsg"
            datatype6       =   4
            maxlength6      =   8
            style6          =   0
            tooltiptext6    =   "Enter Default Missing Location"
            colname7        =   "LOCDESC"
            heading7        =   "Desc                                             "
            datafld7        =   "LocDesc"
            maxlength7      =   20
            tooltiptext7    =   "Enter Description"
            colname8        =   "LOCPRCS"
            heading8        =   "Process             "
            datafld8        =   "LocPrcs"
            datatype8       =   4
            maxlength8      =   8
            style8          =   0
            tooltiptext8    =   "Enter Process"
            colname9        =   "LOCUSECAPOF"
            heading9        =   "Phys. Loc          "
            datafld9        =   "LocUseCapOf"
            datatype9       =   4
            maxlength9      =   8
            style9          =   0
            tooltiptext9    =   "Enter Physical Location"
            colname10       =   "LOCPRDSTG"
            heading10       =   "Prc Stg"
            datafld10       =   "LocPrdStg"
            datatype10      =   4
            maxlength10     =   6
            style10         =   0
            colname11       =   "LOCLSTYP"
            heading11       =   "LsTyp"
            datafld11       =   "LocLsTyp"
            datatype11      =   4
            maxlength11     =   2
            style11         =   0
            tooltiptext11   =   "Enter Loss Type "
            colname12       =   "LOCALLOWLSPER"
            heading12       =   "Allowed Ls%"
            datafld12       =   "LocAllowLsPer"
            datatype12      =   2
            mask12          =   "#0.00"
            maxlength12     =   5
            style12         =   0
            tooltiptext12   =   "Enter Allowed Loss %"
            colname13       =   "LOCVALIDYN"
            heading13       =   "Vld"
            datafld13       =   "LocValidYN"
            datatype13      =   4
            maxlength13     =   1
            tooltiptext13   =   "Enter Location Valid (Y/N)"
            colname14       =   "LOCHIDEYN"
            heading14       =   "HideYN"
            datafld14       =   "LocHideYN"
            datatype14      =   4
            maxlength14     =   1
            style14         =   0
            tooltiptext14   =   "Location Hide YN"
            style15         =   0
         End
      End
   End
End
Attribute VB_Name = "EmrFrmLoc"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'---- General Comments on the Entry --------'
' X-'', A-Loc
'--------------------
Option Explicit
Dim moCn As MwfLib.MDOConnection   '*** (Bef speed) bef 24/11/06
Dim mb_ComLoc As Boolean

'****** Sachin 2.14.0 - 20/11/06 - [Multi-Company] ******
Enum en_LocFra
  Cpy = 0
End Enum
'****** Sachin 2.14.0 - 20/11/06 - [Multi-Company] ******

Private Sub ADC_Load()
  '*** Set the First Key Control property to the working field wLocTyp
  '*** Set the First Non Key Control property to the grid GrdLoc
  '*** Set the Child Property Of ADC to GrdLoc
  '*** Set the Previous and Next Control Properties for the Grid GrdLoc
  '*** Set the Hot Key for the Grid GrdLoc as 'A'
  
  Set moCn = adc.Connection    '*** (Bef speed) bef 24/11/06
  Set adc.FirKeyCtl = adc("wLocTyp")
  Set adc.FirNKeyCtl = GrdLoc
      adc.Child = "GrdLoc"
  Set GrdLoc.PrevCtl = adc
  Set GrdLoc.NextCtl = adc
      GrdLoc.HotKey = "A"
 
  '****** Sachin 2.14.0 - 27-10-2006 [Multi-Company] ******
  mb_ComLoc = IIF(moCn.GetFldVal("Select hImpComLoc From Head Where HCd='" + ctSelfCmCd + "' And HCoCd='" + gs_CoCd + "'") = "Y", True, False)
  adc("wLocPrntYN").Visible = mb_ComLoc
  adc("wLocCoCd").Visible = mb_ComLoc
  LblCoCd.Visible = mb_ComLoc
  LblPrntYN.Visible = mb_ComLoc
  CmdCP.Visible = mb_ComLoc
  '****** Sachin 2.14.0 - 27-10 -2006 [Multi-Company] ******
   
End Sub
Private Sub ADC_KeyWhen()
  '*** Bring the frame FraNKeyAll to the top
  FraNKeyAll.ZOrder (0)
  
  '****** Sachin 2.14.0 - 27-10-2006 [Multi-Company] ******
  Call CmpString("K")     ' to set the Compare String Property of the ADC controls when the Focus is in the Key Frame
  
  adc("wLocPrntYN") = "N"
  adc("wLocCoCd") = gs_CoCd
  '****** Sachin 2.14.0 - 27-10-2006 [Multi-Company] ******
  
  Call EnaDisaCmds(True)
  Call HideAllFras
    
End Sub
Private Sub ADC_KeyValid(Cancel As Boolean, ErrMsg As String)
  '*** Check if the Loc Type and the Sort Order selecter is valid (Sort Option Can be '')
  Call atxt_FldChange(0, UCase("wLocTyp"), Cancel, ErrMsg, adc("wLocTyp"))
  If Cancel = True Then Exit Sub
  
  '****** Sachin 2.14.0 - 27-10-2006 [Multi-Company] ******
  Call atxt_FldChange(0, UCase("wLocPrntYn"), Cancel, ErrMsg, adc("wLocPrntYn"))
  If Cancel = True Then Exit Sub
  
  Call atxt_FldChange(0, UCase("wLocCoCd"), Cancel, ErrMsg, adc("WLocCoCd"))
  If Cancel = True Then Exit Sub
  '****** Sachin 2.14.0 - 27-10-2006 [Multi-Company] ******
  
  Call atxt_FldChange(0, UCase("wLocSrt"), Cancel, ErrMsg, adc("wLocSrt"))
  If Cancel = True Then Exit Sub
  
  
End Sub

Private Sub ADC_NonKeyWhen(pr_Cancel As Boolean, pr_ErrMsg As String)
  '*** Show the columns for entering the default Loss Loc, Missing Loc and Broken Loc
       'only if the Loc Type = 'P', other wise hide them
    
    '****** Sachin 2.14.0 - 27-10-2006 [Multi-Company] ******
    'GrdLoc.ColProp("LocPrntYN").Style = fgcinvisible
    
    If mb_ComLoc = True Then
      GrdLoc.ColProp("LocPrntCd").Style = IIF(adc("wLocPrntYN") = "N", fgcNormal, fgcinvisible)
     Else
      GrdLoc.ColProp("LocPrntCd").Style = fgcinvisible
    End If
    '****** Sachin 2.14.0 - 27-10-2006 [Multi-Company] ******
  
  GrdLoc.ColProp("LocHideYN").Style = IIF(adc("wLocTyp") = "S", fgcNormal, fgcinvisible)
  If adc("wLocTyp") = "P" And adc("wLocPrntYN") = "N" Then
    'GrdLoc.ColProp("LOCPRNTCD").Style = fgcNormal
    GrdLoc.ColProp("LocPrdTyp").Style = fgcinvisible  'EmrSv
    GrdLoc.ColProp("LocLoss").Style = fgcNormal
    GrdLoc.ColProp("LocBrk").Style = fgcNormal
    GrdLoc.ColProp("LocMsg").Style = fgcNormal
    GrdLoc.ColProp("LocPrcs").Style = fgcNormal
    GrdLoc.ColProp("LocUseCapOf").Style = fgcNormal
    '    ** uma
    GrdLoc.ColProp("LocPrdStg").Style = fgcinvisible
    GrdLoc.ColProp("LocLsTyp").Style = fgcinvisible       'EmrSv
    GrdLoc.ColProp("LocAllowLsPer").Style = fgcinvisible  'EmrSv
    'GrdLoc.ColProp("LocPrdStg").Style = fgcNormal  ** uma
    '***Geeta***Emr208
    'GrdLoc.ColProp("LocLsTyp").Style = fgcNormal   ** uma
  'GrdLoc.ColProp("LocAllowLsPer").Style = fgcNormal     ' ***** Manali 3.10.0 - 30/03/12 - Allowed Ls% for PLoc
  '    ** uma
  ElseIf adc("wLocTyp") = "R" And adc("wLocPrntYN") = "N" Then
    'GrdLoc.ColProp("LOCPRNTCD").Style = fgcNormal
    GrdLoc.ColProp("LocPrdTyp").Heading = "PrdTyp"
    GrdLoc.ColProp("LocPrdTyp").Style = fgcNormal
    GrdLoc.ColProp("LocPrdStg").Style = fgcNormal
    GrdLoc.ColProp("LocLsTyp").Style = fgcNormal
    GrdLoc.ColProp("LocAllowLsPer").Style = fgcNormal
    GrdLoc.ColProp("LocLoss").Style = fgcinvisible
    GrdLoc.ColProp("LocBrk").Style = fgcinvisible
    GrdLoc.ColProp("LocMsg").Style = fgcinvisible
    GrdLoc.ColProp("LocPrcs").Style = fgcinvisible       'EmrSv
    GrdLoc.ColProp("LocUseCapOf").Style = fgcinvisible   'EmrSv
  Else
    'GrdLoc.ColProp("LOCPRNTCD").Style = fgcinvisible    'uma invisible -> visible in grid
    GrdLoc.ColProp("LocPrdTyp").Style = fgcinvisible
    GrdLoc.ColProp("LocLoss").Style = fgcinvisible
    GrdLoc.ColProp("LocBrk").Style = fgcinvisible
    GrdLoc.ColProp("LocMsg").Style = fgcinvisible
    GrdLoc.ColProp("LocPrdStg").Style = fgcinvisible
    '***Geeta***Emr208
    GrdLoc.ColProp("LocLsTyp").Style = fgcinvisible
    GrdLoc.ColProp("LocAllowLsPer").Style = fgcinvisible    ' ***** Manali 3.10.0 - 30/03/12 - Allowed Ls% for PLoc
    GrdLoc.ColProp("LocPrcs").Style = fgcinvisible       'EmrSv
    GrdLoc.ColProp("LocUseCapOf").Style = fgcinvisible   'EmrSv
  End If
  
  Call EnaDisaCmds(False)
  
End Sub

Private Sub ATXT_FldWhen(Index As Integer, ByVal IdName As String, Cancel As Boolean, ErrMsg As String)

  '*** Set Helps For Each Field
  Select Case IdName
  Case Is = UCase("wLocTyp")
    Call HlpList.PMCd("LOCTYP")
    
  '****** Sachin 2.14.0 - 27-10-2006 [Multi-Company] ******
  Case Is = UCase("wLocPrntYn")
    If adc("wLocCoCd") <> ctSelfCoCd Then adc("wLocPrntYN") = "N": Cancel = True: Exit Sub
    Call HlpList.PMCd("YN")
    
  Case Is = UCase("wLocCoCd")
    Call HlpList.hCoCd
  '****** Sachin 2.14.0 - 27-10-2006 [Multi-Company] ******
    
  Case Is = UCase("wLocSrt")
    Call HlpList.PMCd("SRTLOC")
    
  '****** Sachin 2.14.0 - 27-10-2006 [Multi-Company] ******
  Case Is = UCase("wFrLoc"), UCase("wToLoc")
    Call HlpList.Loc(ctSelfCoCd, "'" + adc("wLocTyp") + "'", , "'N'")
  '****** Sachin 2.14.0 - 27-10-2006 [Multi-Company] ******
    
  End Select
  
End Sub

Private Sub atxt_FldChange(Index As Integer, ByVal IdName As String, Cancel As Boolean, ErrMsg As String, ByVal pv_NewValue As Variant)
  '*** wLocTyp      valid PMCd('LOCTYP')
  '*** wLocSrt      valid PMCd('SRTLOC') or ''
  Select Case IdName
  Case Is = UCase("wLocTyp")
    Cancel = Not moCn.RecSeek("Select PMCd From Param where PTyp= 'LOCTYP' and " + _
                         "PMCd = '" + pv_NewValue + "'")
    If Cancel = True Then ErrMsg = "Invalid Location Type": Exit Sub
    
  '****** Sachin 2.14.0 - 27-10-2006 [Multi-Company] ******
  Case Is = UCase("wLocPrntYn")
    Cancel = Not moCn.RecSeek("Select PMCd From Param where PTyp= 'YN' and " + _
                         "PMCd = '" + pv_NewValue + "'")
    If Cancel = True Then ErrMsg = "Invalid Parent Code Option Selected": Exit Sub
    
    Cancel = pv_NewValue = "Y" And adc("wLocCoCd") <> ctSelfCoCd
    If Cancel = True Then ErrMsg = "Parent Codes Can Be Specified Only For Company 'ZZZ'": Exit Sub
    
  Case Is = UCase("wLocCoCd")
    Cancel = Not moCn.RecSeek("Select HCoCd From Head where HCoCd= '" + pv_NewValue + "' " + _
              " and HCd= '" + ctSelfCmCd + "'")
    If Cancel = True Then ErrMsg = "Invalid Company Code": Exit Sub
  '****** Sachin 2.14.0 - 27-10-2006 [Multi-Company] ******
        
  Case Is = UCase("wLocSrt")
    Cancel = (pv_NewValue <> "") And (Not moCn.RecSeek("Select PMCd From Param Where " + _
                                        "PTyp = 'SRTLOC' And PMCd = '" + pv_NewValue + "'"))
    If Cancel = True Then ErrMsg = "Invalid Sort Option": Exit Sub
  End Select
End Sub

Private Sub atxt_ReCalculate(Index As Integer, ByVal IdName As String, ByVal OldVal As Variant, ByVal RowNum As Integer, ByVal SourceName As String, ByVal SourceColName As String)

'****** Sachin 2.14.0 - 27-10-2006 [Multi-Company] ******
Select Case IdName
Case Is = UCase("wLocPrntYn")
  If adc("wLocCoCd") <> ctSelfCoCd Then adc("wLocPrntYN") = "N"
End Select
'****** Sachin 2.14.0 - 27-10-2006 [Multi-Company] ******

End Sub

Private Sub CmdCP_Click()

  '****** Sachin 2.14.0 - 20-11-2006 [Multi-Company] ******
  '*** Display or hide the Copy Frame depending on the toggle state
  If (adc("wLocCoCd") <> ctSelfCoCd) And adc("wLocPrntYN") <> "Y" Then
    Call DispFra(Cpy)
  Else
    DispMsg "Cannot Copy Parent Codes Or To Company Code Cannot be 'ZZZ'", etError
  End If
  '****** Sachin 2.14.0 - 20-11-2006 [Multi-Company] ******

End Sub

Private Sub CmdCpGo_Click()
  '****** Sachin 2.14.0 - 27-10-2006 [Multi-Company]
  Call CpyLoc
  Call CmdCP_Click
  '****** Sachin 2.14.0 - 27-10-2006 [Multi-Company]
End Sub

Private Sub GrdLoc_ColWhen(ByVal RowNum As Integer, ByVal ColName As String, Cancel As Boolean, ErrMsg As String, ByVal Action As MwfCtl.en_fgAction)
  
  '****** Sachin 2.14.0 - [06-12-2006] ******
  ' ***** Manali Trading Module - new location 'XD' added for Check
  If adc("wLocTyp") = "F" Or adc("wLocTyp") = "W" Or adc("wLocTyp") = "XD" Or adc("wLocTyp") = "XR" Then
     GrdLoc.AllowAdd = False: GrdLoc.AllowDelete = False
  Else
    GrdLoc.AllowAdd = True: GrdLoc.AllowDelete = True
  End If
  '****** Sachin 2.14.0 - [06-12-2006] ******
  
  '*** Set Helps For Each Field
  Select Case ColName
  Case Is = UCase("LocPrdTyp")
    If adc("WLOCTYP") = "R" Then Call HlpList.PMCd("PRDTYP")    '**** uma
    'If adc("WLOCTYP") = "P" Then Call HlpList.Loc(adc("wLocCoCd"), "'R'") '**** uma PrdTyp changed to PrcTyp
  Case Is = UCase("LocPrcs")
    Call HlpList.Loc(adc("wLocCoCd"), "'R'")  'EmrSv
  Case Is = UCase("LocUseCapOf")
    Call HlpList.Loc(adc("wLocCoCd"), "'P'")  'EmrSv
  Case Is = UCase("LocLoss")
    Call HlpList.Loc(adc("wLocCoCd"), "'L'")
  Case Is = UCase("LocBrk")
    Call HlpList.Loc(adc("wLocCoCd"), "'B'")
  Case Is = UCase("LocMsg")
    Call HlpList.Loc(adc("wLocCoCd"), "'M'")
  Case Is = UCase("LocPrdStg")
    Call HlpList.PMCd("PrdStg")
' ###################################  Manoj #### Ver: 2.0.6  ###################################
' Date:   20/01/2004
  Case Is = UCase("LocValidYN"), UCase("LocHideYN")
    Call HlpList.PMCd("YN")
' ###########################################  Manoj  ###########################################
'***Geeta***Emr208
 Case Is = UCase("LocLsTyp")
    Call HlpList.PMCd("LSTYP")
'***Geeta***Emr208
  '****** Sachin 2.14.0 - 27-10-2006 [Multi-Company]
  Case Is = UCase("LocPrntCd")
      If adc("wLocCoCd") <> ctSelfCoCd And GrdLoc.IsNew(RowNum) = False Then _
        Cancel = True: ErrMsg = "Parent Code Can Be Edited Only in 'ZZZ' Company"
        
      If adc("wLocCoCd") = ctSelfCoCd And (adc("wLocTyp") = "F" Or adc("wLocTyp") = "W" Or adc("wLocTyp") = "XR") Then _
        Cancel = True: ErrMsg = "Cannnot Edit Parent Codes When LocTyp = 'F' or 'W' or 'XR'"
        
      Call HlpList.Loc(ctSelfCoCd, "'" + adc("wLocTyp") + "'", , "'Y'")
      
  '****** Sachin 2.14.0 - 27-10-2006 [Multi-Company]
End Select
End Sub
Private Sub GrdLoc_ColChange(ByVal RowNum As Integer, ByVal ColName As String, ByVal NewValue As Variant, Cancel As Boolean, ErrMsg As String)
  '*** LocCd Not Blank
  '*** LocPrdTyp
      'If LocTyp= 'P' then valid PMCd('PRDTYP')
      'else ''
  '*** LocLoss
      'If LocTyp= 'P' then valid LocCd from Loc(LocTyp= 'L') or ''
      'else ''
  '*** LocBrk         valid LocCd from Loc(LocTyp= 'B') or ''
      'If LocTyp= 'P' then valid LocCd from Loc(LocTyp= 'B') or ''
      'else ''
  '*** LocMsg         valid LocCd from Loc(LocTyp= 'M') or ''
      'If LocTyp= 'P' then valid LocCd from Loc(LocTyp= 'M') or ''
      'else ''
  
  With GrdLoc
    Select Case ColName
    '****** Sachin 2.14.0 - 27-10-2006 [Multi-Company] ******
    Case Is = UCase("LocPrntCd")
      If (adc("wLocPrntYN") = "N") And _
          Not moCn.RecSeek("Select 'x' From Loc Where LocCoCd = '" + ctSelfCoCd + "' And LocCd = '" + NewValue + "' And LocPrntYN='Y' And LocValidYN='Y' And LocTyp='" + adc("wLocTyp") + "'") _
          And mb_ComLoc = True Then
        Cancel = True: ErrMsg = "Invalid Parent Code/Parent Code Not Defined": Exit Sub
      End If
      
      If (adc("wLocPrntYN") = "N") And adc("wLocCoCD") <> ctSelfCoCd And _
          moCn.RecSeek("Select 'x' From Loc Where LocCoCd = '" + ctSelfCoCd + "' And LocCd = '" + .Value(RowNum, "LocCd") + "' And LocPrntYN='N'") And _
          Not moCn.RecSeek("Select 'x' From Loc Where LocCoCd = '" + ctSelfCoCd + "' And LocCd = '" + .Value(RowNum, "LocCd") + "' And LocPrntYN='N' And LocPrntCd='" + NewValue + "' ") _
          And mb_ComLoc = True Then
      Cancel = True: ErrMsg = "Parent Code Should be Same as the Parent Code Specified For This Location In 'ZZZ' Company": Exit Sub
      End If
          
'    Case Is = UCase("LocCd")
'      If NewValue = "" Or (ADC("wLocCoCd") <> ctSelfCoCd And Not moCn.RecSeek("Select 'x' From Loc Where LocCoCd='" + ctSelfCoCd + "' And LocCd='" + NewValue + "'")) Then
'        Cancel = True: ErrMsg = "Not a Valid Location Code/'ZZZ' Record Not Found For This Location Code": Exit Sub
'      End If
    '****** Sachin 2.14.0 - 27-10-2006 [Multi-Company] ******
      
    Case Is = UCase("LocPrdTyp")
      'If adc("wLocTyp") = "R" Then
      Cancel = (adc("wLocTyp") = "R") And Not moCn.RecSeek("Select PMCd From Param Where " + _
                                          "PTyp = 'PRDTYP' And PMCd = '" + NewValue + "'")
      ErrMsg = "Invalid Production Type": Exit Sub
      'End If
      'If adc("wLocTyp") = "P" Then
      'Cancel = (adc("wLocTyp") = "P") And (Not moCn.RecSeek("Select LocCd From Loc Where " + _
                                          "LocCoCd= '" + adc("wLocCoCd") + "' " + _
                                          "and LocTyp= 'R' and LocCd = '" + NewValue + "'"))
      'ErrMsg = "Invalid Process Type": Exit Sub
      'End If
    'EmrSv
    Case Is = UCase("LocPrcs")
      Cancel = (adc("wLocTyp") = "P") And (Not moCn.RecSeek("Select LocCd From Loc Where " + _
                                          "LocCoCd= '" + adc("wLocCoCd") + "' " + _
                                          "and LocTyp= 'R' and LocCd = '" + NewValue + "'"))
      ErrMsg = "Invalid Process Type": Exit Sub
    Case Is = UCase("LocUseCapOf")
      Cancel = (GrdLoc.Value(RowNum, "LocCd") <> NewValue) And (Not moCn.RecSeek("Select LocCd From Loc Where " + _
                                          "LocCoCd= '" + adc("wLocCoCd") + "' " + _
                                          "and LocTyp= 'P' and LocCd = '" + NewValue + "'"))
      ErrMsg = "Invalid Physical Location": Exit Sub
  
    Case Is = UCase("LocLoss")
      Cancel = (adc("wLocTyp") = "P") And (NewValue <> "") And (Not moCn.RecSeek("Select LocCd From Loc where " + _
                                                      "LocCoCd= '" + adc("wLocCoCd") + "' " + _
                                                      "and LocTyp= 'L' " + _
                                                      "and LocCd = '" + NewValue + "'"))
      If Cancel = True Then ErrMsg = "Invalid Loss Location": Exit Sub
      
      If (adc("wLocTyp") <> "P") And (NewValue <> "") Then _
         Cancel = True: ErrMsg = "Loss Location Should Not Be Entered": Exit Sub
    Case Is = UCase("LocBrk")
      Cancel = (NewValue <> "") And (Not moCn.RecSeek("Select LocCd From Loc where " + _
                                    "LocCoCd= '" + adc("wLocCoCd") + "' and LocTyp= 'B' " + _
                                    "and LocCd = '" + NewValue + "'"))
      If Cancel = True Then ErrMsg = "Invalid Break Location": Exit Sub
      If (adc("wLocTyp") <> "P") And (NewValue <> "") Then _
         Cancel = True: ErrMsg = "Break Location Should Not Be Entered": Exit Sub
    Case Is = UCase("LocMsg")
      Cancel = (NewValue <> "") And (Not moCn.RecSeek("Select LocCd From Loc where " + _
                                    "LocCoCd= '" + adc("wLocCoCd") + "' and LocTyp= 'M' " + _
                                    "and LocCd = '" + NewValue + "'"))
      If Cancel = True Then ErrMsg = "Invalid Missing Location": Exit Sub
      If (adc("wLocTyp") <> "P") And (NewValue <> "") Then _
         Cancel = True: ErrMsg = "Missing Location Should Not Be Entered": Exit Sub
    Case Is = UCase("LocPrdStg")
      Cancel = (NewValue <> "") And Not moCn.RecSeek("Select PMCd From Param Where " + _
               "PTyp = 'PRDSTG' And PMCd = '" + NewValue + "'")
      If Cancel = True Then ErrMsg = "Invalid Production Stage": Exit Sub
' ###################################  Manoj #### Ver: 2.0.6  ###################################
' Date:   20/01/2004
    Case Is = UCase("LocValidYN")
      Cancel = (NewValue <> "") And Not moCn.RecSeek("Select PMCd From Param Where " + _
               "PTyp = 'YN' And PMCd = '" + NewValue + "'")
      If Cancel = True Then ErrMsg = "Invalid Location Valid (Y/N)": Exit Sub
      
      Cancel = (NewValue = "N") And moCn.RecSeek("Select 'x' From Loc where LocPrntYN='N' " + _
              " And LocCoCd<>'" + ctSelfCoCd + "' And LocPrntCd='" + GrdLoc.Value(RowNum, "LocCd") + "'")
      If Cancel = True Then ErrMsg = "Cannot Set ValidYN to 'N' As Parent Code already in Use": Exit Sub
' ###########################################  Manoj  ###########################################
'***Geeta***Emr208
    Case Is = UCase("LocLsTyp")
      Cancel = (NewValue <> "") And Not moCn.RecSeek("Select PMCd From Param Where " + _
               "PTyp = 'LSTYP' And PMCd = '" + NewValue + "'")
      If Cancel = True Then ErrMsg = "Invalid Loss Type": Exit Sub
   Case Is = UCase("LocHideYN")
      Cancel = (NewValue <> "") And Not moCn.RecSeek("Select PMCd From Param Where " + _
               "PTyp = 'YN' And PMCd = '" + NewValue + "'")
      If Cancel = True Then ErrMsg = "Invalid Hide Option.Enter 'Y' To Hide the Location in Stock Ledger,Flash stock Else 'N'. ": Exit Sub
    End Select
  End With
End Sub

Private Sub GrdLoc_ReCalc(ByVal RowNum As Integer, ByVal ColName As String, ByVal OldVal As Variant, ByVal SourceName As String, ByVal SourceColName As String)
'****** Sachin 2.14.0 - 31-10-2006 [Multi-Company] ******
With GrdLoc
    Select Case ColName
    Case Is = UCase("LocPrntCd")
    If (.Mode = fgmnorm Or .Mode = fgmCopy) Then _
      .Value(RowNum, "LocPrntCd") = moCn.GetFldVal("Select LocPrntCd From Loc Where LocPrntYN='N' And LocCoCd='" + ctSelfCoCd + "' And  LocCd='" + .Value(RowNum, "LocCd") + "'")
    End Select
End With
'****** Sachin 2.14.0 - 31-10-2006 [Multi-Company] ******
End Sub

' ########################  Manoj #### Ver: 2.0.6 #### Date: 02/02/2004  ########################
Private Sub GrdLoc_RowWhen(ByVal RowNum As Integer)
  If GrdLoc.IsNew(RowNum) = True Then
    GrdLoc.Value(RowNum, "LocValidYN") = "Y"
    GrdLoc.Value(RowNum, "LocHideYN") = "N"
  End If
End Sub
' ###########################################  Manoj  ###########################################

Private Sub GrdLoc_InitKey(ByVal RowNum As Integer)
  '*** Initialise Key Fields
  With GrdLoc
    .Store "LocCoCd", adc("wLocCoCd")
    .Store "LocTyp", adc("wLocTyp")
    '****** Sachin 2.14.0 - 27-10-2006 [Multi-Company] ******
    .Store "LocPrntYn", adc("wLocPrntYn")
    '****** Sachin 2.14.0 - 27-10-2006 [Multi-Company] ******
    .Store "LocPrtKey", ctCurrPrtn    '****** Sachin 3.02.0 - Default Value for Partition Key
  End With
End Sub
Private Sub ADC_SetRecSource()
  '*** Since this is an 'A' Entry, there is no RecSource associated with ADC. So set the
      'DataMode property of ADC control as xdatUnbound
  adc.DataMode = xdatUnbound
End Sub

Private Sub GrdLoc_SetRecSource()
  '*** Set the Record Source of the Grid GrdLoc
      'Default sort option is LocCd
''''?????? Check
  Dim wCnd  As String, wOrdBy As String
  wCnd = adc.RepCond
  wCnd = IIF(wCnd <> "", " and " + wCnd, "")
  
  wOrdBy = moCn.GetFldVal("Select PDesc225 from Param " + _
           "where PTyp='SRTLOC' and PMCd='" + adc("wLocSrt") + "'")

  wOrdBy = IIF(wOrdBy = "", " Order By LocCd ", "Order By " + wOrdBy)

  GrdLoc.RecSource = " Select * from Loc Where LocCoCd ='" + adc("wLocCoCd") + "' " + _
                                                        wCnd + wOrdBy
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
  Cancel = ATXT(Index).Validate
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    If UnloadMode <> vbFormCode Then Cancel = True: Exit Sub
    
    Set moCn = Nothing    '*** (09/08/05)    '*** (Bef speed) bef 24/11/06
End Sub

Private Sub GrdLoc_Validate(Cancel As Boolean)
  Cancel = GrdLoc.Validate
End Sub

'****** Sachin 2.14.0 - 20/11/06 - [Multi-Company] ******
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
      adc("wFrLoc").SetFocus
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
  adc.AllowSave = Not pv_ShowFra
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
    adc("wLocTyp").CmpStr = "LocTyp= "
    adc("wLocPrntYN").CmpStr = "LocPrntYN= "
    adc("wLocCoCd").CmpStr = "LocCoCd="
    adc("wFrLoc").CmpStr = ""
    adc("wToLoc").CmpStr = ""
    
  '****** Seting values for CmpStr of 'C'opy fields, for Adc.RepCnd (Making Key fields CmpStr = '')
  ElseIf mTyp = "C" Then
    adc("wLocTyp").CmpStr = "LocTyp= "
    adc("wLocPrntYN").CmpStr = ""
    adc("wLocCoCd").CmpStr = ""
    adc("wFrLoc").CmpStr = "LocCd>= "
    adc("wToLoc").CmpStr = "LocCd<= "
  End If
End Sub

Private Sub CpyLoc()
  '*** Copy Option will copy the Prd Points from one PrdCtg to another PrdCtg (in case of 'GN' and 'WP')
      'and from one Setting Type to another (in case of 'WS', 'GS' and 'GP')
      'for the given PrdTyp and for the given Design Code, Loc range

  Dim ws_LocCnd As String
  
  Call CmpString("C")     ' to set the Compare String Property of the ADC controls when the Copy Option Is clicked
  ws_LocCnd = adc.RepCond
  ws_LocCnd = IIF(ws_LocCnd <> "", " Where " + ws_LocCnd, "")
  
  If adc("wFrLoc") = "" Or adc("wToLoc") = "" Then DispMsg "All Fields In The Scope Frame Should Be Filled Up", etError: Exit Sub
    
'  If moCn.RecSeek("Select 'x' from Loc " + ws_LocCnd) Then
'    DispMsg "Cannot Copy as Location Already Exist In This Company", etError
'    Exit Sub
'  End If
  
  '*** Beginning the copy process
  GrdLoc.StartCopy
  MWLib.BeginProcess Me, "Copying Production Points ..."
  Dim wRsPrdPts As MwfLib.MDORowSet, wi_Row As Integer
  
  Set wRsPrdPts = moCn.OpenRes("Select * from Loc Lc " + ws_LocCnd + " And LocCoCd= '" + ctSelfCoCd + "' And LocPrntYN='N' And " + _
                                " Not Exists (Select 'x' From Loc Le Where Le.LocCd=Lc.LocCd And LocCoCd='" + adc("wLocCoCd") + "')")
  
  With wRsPrdPts
    Do While Not (.EOF Or .BOF)
      MWLib.RunningProcess "Copying Locations ..."
      GrdLoc.AddItem
      wi_Row = GrdLoc.Rows - 1
      GrdLoc.Value(wi_Row, "LocCd") = !LocCd
      GrdLoc.Value(wi_Row, "LocDesc") = !LocDesc
      GrdLoc.Value(wi_Row, "LocPrdTyp") = !LocPrdTyp
      GrdLoc.Value(wi_Row, "LocLoss") = !LocLoss
      GrdLoc.Value(wi_Row, "LocBrk") = !LocBrk
      GrdLoc.Value(wi_Row, "LocMsg") = !LocMsg
      GrdLoc.Value(wi_Row, "LocPrdStg") = !LocPrdStg
      GrdLoc.Value(wi_Row, "LocValidYN") = !LocValidYN
      GrdLoc.Value(wi_Row, "LocLsTyp") = !LocLsTyp
      GrdLoc.Value(wi_Row, "LocAllowLsPer") = !LocAllowLsPer    ' ***** Manali 3.10.0 - 30/03/12 - Allowed Ls% for PLoc
      GrdLoc.Value(wi_Row, "LocPrntCd") = !LOCPRNTCD
      GrdLoc.Value(wi_Row, "LocHideYN") = !LocHideYN
      GrdLoc.SaveRec (wi_Row)
      .MoveNext
    Loop
  End With
  MWLib.EndProcess Me
  GrdLoc.EndCopy
  Set wRsPrdPts = Nothing
  '*** End of the copy process
End Sub

'****** Sachin 2.14.0 - 20/11/06 - [Multi-Company] ******
