VERSION 5.00
Object = "{AB2CC32A-E652-42A2-B341-26ECE02F6386}#2.1#0"; "mwfctlforspeed9.ocx"
Begin VB.Form EmrFrmvParam 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Variable Parameter Entry"
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
   Begin MwfCtl.MWCTL_HLP MWCTL_HLP1 
      Height          =   435
      Left            =   9330
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
      Left            =   60
      TabIndex        =   6
      Top             =   9300
      Width           =   15060
      _ExtentX        =   26564
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
      Height          =   300
      Left            =   0
      TabIndex        =   8
      Top             =   24
      Width           =   9210
      Begin MwfCtl.MWCTL_MED ATXT 
         Height          =   285
         Index           =   7
         Left            =   3225
         TabIndex        =   1
         ToolTipText     =   "Enter Parameter Type"
         Top             =   0
         Width           =   1380
         _ExtentX        =   2434
         _ExtentY        =   503
         MaxLength       =   8
         DataType        =   4
         ReCalcParent    =   "WVPCOCD,WVPVALUE3"
         ReCalcOn        =   "WVENTRYTYP"
         IdName          =   "WVPTYP"
         CmpStr          =   "vPTyp ="
      End
      Begin MwfCtl.MWCTL_MED ATXT 
         Height          =   285
         Index           =   87
         Left            =   6240
         TabIndex        =   2
         ToolTipText     =   "Enter Company Code"
         Top             =   0
         Width           =   600
         _ExtentX        =   1058
         _ExtentY        =   503
         MaxLength       =   3
         DataType        =   4
         ReCalcOn        =   "WVPTYP"
         IdName          =   "WVPCOCD"
         CmpStr          =   "vPCoCd = "
      End
      Begin MwfCtl.MWCTL_MED ATXT 
         Height          =   285
         Index           =   0
         Left            =   1245
         TabIndex        =   0
         ToolTipText     =   "Specify Entry Type"
         Top             =   0
         Width           =   360
         _ExtentX        =   635
         _ExtentY        =   503
         DataType        =   4
         ReCalcParent    =   "WVPTYP"
         IdName          =   "WVENTRYTYP"
      End
      Begin MwfCtl.MWCTL_MED ATXT 
         Height          =   285
         Index           =   1
         Left            =   8010
         TabIndex        =   3
         ToolTipText     =   "Enter Parent Sequence (YN)"
         Top             =   0
         Width           =   390
         _ExtentX        =   688
         _ExtentY        =   503
         DataType        =   4
         ReCalcOn        =   "WVPTYP"
         IdName          =   "WVPVALUE3"
         CmpStr          =   "vPValue3 = "
      End
      Begin VB.Label LblPrntYN 
         BackStyle       =   0  'Transparent
         Caption         =   "Prnt YN"
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
         Left            =   7170
         TabIndex        =   13
         Top             =   0
         Width           =   1365
      End
      Begin VB.Label ALBL 
         BackStyle       =   0  'Transparent
         Caption         =   "Entry Type"
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
         Left            =   90
         TabIndex        =   12
         Top             =   0
         Width           =   1185
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
         Height          =   285
         Index           =   0
         Left            =   4890
         TabIndex        =   11
         Top             =   0
         Width           =   1365
      End
      Begin VB.Label ALBL 
         BackStyle       =   0  'Transparent
         Caption         =   "Param Type"
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
         Left            =   1920
         TabIndex        =   10
         Top             =   0
         Width           =   1335
      End
   End
   Begin VB.Frame FRA_NKEYS 
      BorderStyle     =   0  'None
      Enabled         =   0   'False
      Height          =   9360
      Left            =   -120
      TabIndex        =   7
      Top             =   0
      Width           =   15345
      Begin VB.Frame FraNKeyAll 
         BorderStyle     =   0  'None
         Height          =   9225
         Left            =   150
         TabIndex        =   9
         Top             =   90
         Width           =   15150
         Begin MwfCtl.MWCTL_FLX GrdvParam 
            Height          =   8880
            Left            =   30
            TabIndex        =   4
            Top             =   240
            Width           =   15105
            _ExtentX        =   26644
            _ExtentY        =   15663
            Cols            =   14
            colname1        =   "VPMCD"
            heading1        =   "Main Cd                 "
            datafld1        =   "vPMCd"
            datatype1       =   4
            maxlength1      =   8
            tooltiptext1    =   "Enter Param Main Code"
            colname2        =   "VPSCD"
            heading2        =   "Sub Cd           "
            datafld2        =   "vPSCd"
            datatype2       =   4
            maxlength2      =   8
            tooltiptext2    =   "Enter Param Sub Code"
            colname3        =   "VPDESC"
            heading3        =   "Desc                     "
            datafld3        =   "vPDesc"
            maxlength3      =   30
            tooltiptext3    =   "Enter Description"
            colname4        =   "VPDESC225"
            heading4        =   "Desc225                "
            datafld4        =   "vPDesc225"
            datatype4       =   4
            recalcparent4   =   "GrdvParam(VPVALUE3)"
            maxlength4      =   225
            tooltiptext4    =   "Enter PDesc225"
            colname5        =   "VPVALUE"
            heading5        =   "Value                     "
            datafld5        =   "vPValue"
            datatype5       =   4
            maxlength5      =   30
            tooltiptext5    =   "Enter PValue"
            colname6        =   "VPNUM"
            heading6        =   "Number            "
            datafld6        =   "vPNum"
            datatype6       =   2
            mask6           =   "######0.0000"
            recalcon6       =   "GrdvParam(WACKYN)"
            recalcparent6   =   "GrdvParam(WACKYN)"
            maxlength6      =   12
            tooltiptext6    =   "Enter PNum"
            colname7        =   "VPVALUE1"
            heading7        =   "Value1                   "
            datafld7        =   "vPValue1"
            datatype7       =   4
            maxlength7      =   60
            tooltiptext7    =   "Enter PValue1"
            colname8        =   "VPNUM1"
            heading8        =   "Number1           "
            datafld8        =   "vPNum1"
            datatype8       =   2
            mask8           =   "######0.0000"
            maxlength8      =   12
            tooltiptext8    =   "Enter PNum1"
            colname9        =   "VPVALUE2"
            heading9        =   "Value2                   "
            datafld9        =   "vPValue2"
            datatype9       =   4
            maxlength9      =   30
            tooltiptext9    =   "Enter PValue2"
            colname10       =   "VPVALUE3"
            heading10       =   "Value3                       "
            datafld10       =   "vPValue3"
            datatype10      =   4
            maxlength10     =   30
            tooltiptext10   =   "Enter PValue3"
            colname11       =   "VPVALIDYN"
            heading11       =   "Yn"
            datafld11       =   "vPValidYN"
            datatype11      =   4
            maxlength11     =   1
            style11         =   0
            tooltiptext11   =   "Enter Parameter Valid (Y/N)"
            colname12       =   "VPVALUE4"
            heading12       =   "Value4"
            datafld12       =   "VPVALUE4"
            maxlength12     =   120
            tooltiptext12   =   "Enter PValue4"
            colname13       =   "WACKYN"
            heading13       =   "AckYn"
            datatype13      =   4
            recalcon13      =   "GrdvParam(VPNUM)"
            recalcparent13  =   "GrdvParam(VPNUM)"
            maxlength13     =   1
            style13         =   0
            tooltiptext13   =   "Enter Acknowledge (Y/N)"
         End
      End
   End
End
Attribute VB_Name = "EmrFrmvParam"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'*** (A hyphen and three stars are used for general commenting
'****** (A hyphen and six stars are used for commenting difficult syntax)
'### (A hyphen and three hashes are used when there is some problem which has to be looked into)
'##### (A hyphen and five hashes are used for alternative codes)
'@@@@@ (A hyphen and five 'AtTheRates' are used for temporarily commenting code which is useful)


'--- General Comments ----
' There are two types of parameters under two different menu options :
'    1. System Parameters (These are high level parameters which the normal user has no rights to)
'    2. Parameters (These parameters can be changed by normal users)
'Whether a parameter is a System Parameter or not is specified in the field 'SysParam(Y/N)'
'    of the Parameter Type 'PTYP'
' A- Entry, GrdvParam
'TcStr table is used to set the properties of all the fields
'    in the grid GrdvParam
' fn GrdPropSet called from Adc_recsource, sets the properties like enable, visible,
'    heading, length and mask of each column by reading from TcStr table.
'
'Once the System Parameters are entered no further Add,Change or Delete Operations are allowed on them
'---

Option Explicit
Dim moCn As MwfLib.MDOConnection   '*** (Bef speed) bef 24/11/06
Dim ms_vEntryTyp As String

'****** Sachin 2.14.0 - [Multi-Company]
Dim mb_MultiCoCd As Boolean
'****** Sachin 2.14.0 - [Multi-Company]

'Dim wi_CharWdth As Integer
Private Sub ADC_Load()
  '*** Set the First Key Control to the field wvPTyp (working field for Parameter Type)
  '*** Set the First Non Key Control to GrdvParam
  '*** Set the Child Property Of ADC to GrdvParam
  '*** Set the Previous and Next Control Properties for Grid GrdvParam as ADC
  '*** Set the Hot Keys for the Grids GrdvParam as 'A'
  
  Set moCn = ADC.Connection    '*** (Bef speed) bef 24/11/06
  'Set ADC.FirKeyCtl = ADC("wvPTyp")
  Set ADC.FirKeyCtl = ADC("wvEntryTyp")
  Set ADC.FirNKeyCtl = GrdvParam
  ADC.Child = "GrdvParam"
  Set GrdvParam.PrevCtl = ADC
  Set GrdvParam.NextCtl = ADC
  GrdvParam.HotKey = "A"
  
  '****** Sachin 2.14.0 - [Multi-Company]
  mb_MultiCoCd = IIF(moCn.GetFldVal("Select hImpComLoc From Head Where HCd='" + ctSelfCmCd + "' And HCoCd='" + ctSelfCoCd + "'") = "Y", True, False)
  ADC("WVPVALUE3").Visible = mb_MultiCoCd
  LblPrntYN.Visible = mb_MultiCoCd
  '****** Sachin 2.14.0 - [Multi-Company]
  
End Sub
Private Sub ADC_KeyWhen()
  '*** Set the Frame FraNKeyAll as the first frame
  If ms_vEntryTyp <> "" Then ADC("wvEntryTyp") = ms_vEntryTyp
  FraNKeyAll.ZOrder (0)
  ADC("wvPCoCd") = gs_CoCd
    
End Sub

Private Sub ADC_KeyValid(pr_Cancel As Boolean, pr_ErrMsg As String)
  '*** Validate PTyp on Menu Cd (i.e. If The Parameter Is System Param Or Other Param) ***
  '*** If the Index of the field is not going to be used then pass 0 as the first
      'parameter of the atxt_FldChange Sub procedure, pass the IdName of the
      'field whose value has to be validated as the second parameter,
      'pass pr_Cancel and pr_ErrMsg as the 3rd and 4th parameters,
      'pass the value which has to be validated as the 5th parameter
  
  'Call atxt_FldChange(0, UCase("wvPTyp"), pr_Cancel, pr_ErrMsg, ADC("wvPTyp"))
  'If pr_Cancel = True Then Exit Sub
  If ADC("wvEntryTyp") = "" Then pr_Cancel = True: pr_ErrMsg = "Entry Type Cannot Be Blank": Exit Sub
  If ADC("wvPTyp") = "" Then pr_Cancel = True: pr_ErrMsg = "Param Type Cannot Be Blank": Exit Sub
  If ADC("wvPCoCd") = "" Then pr_Cancel = True: pr_ErrMsg = "Company Code Cannot Be Blank": Exit Sub
End Sub
Private Sub ADC_NonKeyWhen(pr_Cancel As Boolean, pr_ErrMsg As String)
  ms_vEntryTyp = ADC("wvEntryTyp")
  pr_Cancel = False
  '6.1_51 when ptyp=unlock then add or delete not allowed
  If UCase(ADC("wvPTyp")) = UCase("UnLock") Then
      GrdvParam.AllowAdd = False: GrdvParam.AllowDelete = False
  End If
End Sub
Private Sub ATXT_FldWhen(Index As Integer, ByVal IdName As String, Cancel As Boolean, ErrMsg As String)
  '*** Give help for the working field wvPTyp depending on the menucode (System Parameters or Other Parameters)
  Select Case IdName
  Case Is = UCase("wvEntryTyp")
    HlpList.PMCd ("SYSDFN")
  Case Is = UCase("wvPTyp")
    'If UCase(ADC.MenuCd) = UCase("VarSysParam") Then Call HlpList.vPMCd(ctSelfCoCd, "PTYP", "'Y'")
    'If UCase(ADC.MenuCd) = UCase("VarOthParam") Then Call HlpList.vPMCd(ctSelfCoCd, "PTYP", "'N'")
    'If UCase(ADC.MenuCd) = UCase("VarMicParam") Then Call HlpList.vPMCd(ctSelfCoCd, "PTYP", "'M'")
    Call HlpList.vPMCd(ctSelfCoCd, "PTYP", "'" + ADC("wvEntryTyp") + "'")
 'InUse Check- if Ptyp=WIWSEQ/WIWLONG/RWIWSEQ/RWIWLONG and CoCd=ZZZ then help of YN will come
  Case Is = UCase("wvPValue3")
    If (UCase(ADC("wvPTyp")) = "WIWSEQ" Or UCase(ADC("wvPTyp")) = "WIWLONG" Or UCase(ADC("wvPTyp")) = "RWIWSEQ" Or UCase(ADC("wvPTyp")) = "RWIWLONG") And UCase(ADC("wvPCoCd") = UCase(ctSelfCoCd)) Then
      HlpList.PMCd ("YN")
    Else
      Cancel = True: ErrMsg = "Cannot Edit If vPTyp Is Not 'WIWSEQ' Or 'WIWLONG' Or 'RWIWSEQ' Or 'RWIWLONG' And Company is 'ZZZ'"
    End If
  '****** Sachin 2.14.0 - 01-11-2006 [Multi-Company] ******
  Case Is = UCase("wvPCoCd")
    If UCase(ADC("wvPTyp")) = "PTYP" Then Cancel = True: ErrMsg = "Cannot Edit This Value": Exit Sub
    'If (UCase(ADC("wvPTyp")) = "WIWSEQ" Or UCase(ADC("wvPTyp")) = "WIWLONG") Then Cancel = True: ErrMsg = "Company Code Can Only Be 'ZZZ": Exit Sub
    HlpList.hCoCd
  '****** Sachin 2.14.0 - 01-11-2006 [Multi-Company] ******
  '****** Sachin 2.14.0 - 01-11-2006 [Multi-Company] ******
  End Select
End Sub
Private Sub ATXT_FldChange(Index As Integer, ByVal IdName As String, Cancel As Boolean, ErrMsg As String, ByVal pv_NewValue As Variant)
  '*** Validating PTyp on Menu Cd (i.e. If The Parameter Is System Param Or Simple Param) ***
  Dim ws_Sys As String
  Select Case IdName
  Case Is = UCase("wvEntryTyp")
    Cancel = Not moCn.RecSeek("Select PMCd From Param where PTyp= 'SYSDFN' and " + _
                              "PMCd = '" + pv_NewValue + "'")
    ErrMsg = "Invalid Entry Type": Exit Sub
  Case Is = UCase("wvPTyp")
    'If UCase(ADC.MenuCd) = UCase("VarSysParam") Then ws_Sys = "Y"
    'If UCase(ADC.MenuCd) = UCase("VarOthParam") Then ws_Sys = "N"
    'If UCase(ADC.MenuCd) = UCase("VarMicParam") Then ws_Sys = "M"
    Cancel = (pv_NewValue <> "") And (Not moCn.RecSeek("Select vPMCd From vParam where " + _
             "vPCoCd = '" + ctSelfCoCd + "' and vPTyp= 'PTYP' and " + _
             "vPValue= '" + ADC("wvEntryTyp") + "' and vPMCd = '" + pv_NewValue + "'"))
    If Cancel = True Then ErrMsg = "Invalid Parameter Type": Exit Sub
    
    '****** Sachin 2.14.0 - 03-11-2006 [Multi-Company] ******
    'InUse Check- PrntYN field will enable for Ptyp=WIWSEQ/WIWLONG/RWIWSEQ/RWIWLONG
    If UCase(pv_NewValue) <> "WIWSEQ" And UCase(pv_NewValue) <> "WIWLONG" And UCase(pv_NewValue) <> "RWIWSEQ" And UCase(pv_NewValue) <> "RWIWLONG" Then
      ADC("wvPValue3").Enabled = False
    Else
      ADC("wvPValue3").Enabled = True
    End If
    '****** Sachin 2.14.0 - 03-11-2006 [Multi-Company] ******
        
  Case Is = UCase("wvPCoCd")
    Cancel = Not moCn.RecSeek("Select HCoCd From Head where HCoCd= '" + pv_NewValue + "' and HCd= '" + ctSelfCmCd + "'")
    If Cancel = True Then ErrMsg = "Invalid Company Code": Exit Sub
    
    '****** Sachin 2.14.0 - 03-11-2006 [Multi-Company] ******
    'Cancel = (UCase(ADC("wvPTyp")) = "WIWSEQ" Or UCase(ADC("wvPTyp")) = "WIWLONG") And pv_NewValue <> ctSelfCoCd
    'If Cancel = True Then ErrMsg = "Company Code Should be 'ZZZ' For 'WIWSEQ' And 'WIWLONG'": Exit Sub
          
  Case Is = UCase("wvPValue3")
    Cancel = (pv_NewValue <> "") And Not moCn.RecSeek("Select PMCd From Param Where PTyp= 'YN' and PMCd= '" + pv_NewValue + "'")
    If Cancel = True Then ErrMsg = "Invalid Parent Code, Enter(Y/N)": Exit Sub
  '****** Sachin 2.14.0 - 03-11-2006 [Multi-Company] ******
    
  End Select
End Sub
Private Sub GrdVPARAM_RowSave(ByVal RowNum As Integer, Cancel As Boolean, ErrMsg As String)
  If ADC("WVPTYP") = "CHKTOL" Then
    If GrdvParam.Value(RowNum, "vpDESC") = "" Then
      Cancel = True
      ErrMsg = "BChr cannot be blank"
    End If
    If GrdvParam.Value(RowNum, "vPValue1") = "" Then
      Cancel = True
      ErrMsg = "RmCtg cannot be blank"
    End If

  End If
End Sub

Private Sub atxt_ReCalculate(Index As Integer, ByVal IdName As String, ByVal OldVal As Variant, ByVal RowNum As Integer, ByVal SourceName As String, ByVal SourceColName As String)

  Select Case IdName
  Case Is = UCase("wvPTyp")
    If ADC.Mode = xNorm Then
      ADC("wvPTyp") = ""
    End If
    
 Case Is = UCase("wvPCoCd")
    If ADC.Mode = xNorm Then
      If UCase(ADC("wvPTyp")) = "PTYP" Or UCase(ADC("wvPTyp")) = "WIWSEQ" Or UCase(ADC("wvPTyp")) = "WIWLONG" Or UCase(ADC("wvPTyp")) = "RWIWSEQ" Or UCase(ADC("wvPTyp")) = "RWIWLONG" Then
      '    ADC("wvPCoCd") = ctSelfCoCd
      'Else
          ADC("wvPCoCd") = gs_CoCd
      End If
    End If
    
  '****** Sachin 2.14.0 - 1-11-2006 [Multi-Company] ******
  Case Is = UCase("wvPValue3")
    If ADC.Mode = xNorm Then
    'InUse Check- if Ptyp=WIWSEQ/WIWLONG/RWIWSEQ/RWIWLONG the PrntYN flag is set as N
      If (UCase(ADC("wvPTyp")) = "WIWSEQ" Or UCase(ADC("wvPTyp")) = "WIWLONG" Or UCase(ADC("wvPTyp")) = "RWIWSEQ" Or UCase(ADC("wvPTyp")) = "RWIWLONG") Then
         ADC("wvPValue3") = "N"
      Else
        ADC("wvPValue3") = ""
      End If
    End If
  '****** Sachin 2.14.0 - 1-11-2006 [Multi-Company] ******
  End Select
  
End Sub

Private Sub GrdvParam_ColWhen(ByVal RowNum As Integer, ByVal ColName As String, Cancel As Boolean, ErrMsg As String, ByVal Action As MwfCtl.en_fgAction)
  '*** Set Helps For Each Field
  '*** PMCd
      'PTyp= 'CHR'
      '     Show help of all the Tc's
      'PTyp= 'DAANACD'
      '     Show help for Design Analysis Fields
      'PTyp= 'LABSCD'
      '     Show help for Labour Main Code
      'PTyp= 'RMSCTG' or 'RMCOl'
      '     Show all the Raw Material Categories available
      'PTyp= 'SUBPRC'
      '     Show help of all the Process Locations
  '*** PDesc225
      'PTyp= 'CHR'
      '     Show help of all the Locations which are not Process Locations
      'PTyp= 'GLDAS', 'IWTFRORD', 'CPYRT'
      '     MuliSelect= True    (This option allows the user to select multiple values from the help)
      '     Show help for Raw Material Category
      'PTyp= 'PRDSEQ', 'WIWSEQ', 'WIWLONG'
      '     MuliSelect= True
      '     Show help of all the Locations which are Process Locations
      'PTyp= 'RMSCTG'
      '     Show help for 'YN' (Yes/ No). Here the user has to specify whether the
      '     raw materail transactions are Size wise or not
      'PTyp= 'TC'
      '     Show help of all the Characters for the selected 'TC'. This field is used to
      '     store the dafault 'CHR' value for the current 'TC'
      'PTyp= 'INGR'
      '     MuliSelect= True
      '     Show help of all the Group Participants
      'PTyp= 'SRTBAG'
      '     MuliSelect= True
      '     Show help of all the fields in the Bag Table. This field stores the fields
      '     of Bag Table by which sorting has to be done in the Bag Generation Entry
      'PTyp= 'SRTINBAG'
      '     MuliSelect= True
      '     Show help of all the fields in the InvFgd Table. This field stores the fields
      '     of InvFgd Table by which sorting has to be done in the Invoice All Bags Entry
      'PTyp= 'SRTLR'
      '     MuliSelect= True
      '     Show help of all the fields in the LabRt Table. This field stores the fields
      '     of LabRt Table by which sorting has to be done in the Labour Rate Entry
      'PTyp= 'SRTPP'
      '     MuliSelect= True
      '     Show help of all the fields in the PrdPts Table. This field stores the fields
      '     of PrdPts Table by which sorting has to be done in the Production Points Master Entry
      'PTyp= 'SRTPRHLP'
      '     MuliSelect= True
      '     Show help of all the fields in the Txndz Table. This field stores the fields
      '     of Txndz Table by which sorting has to be done in the help of Base Voucher No
      '     field of Custom Matching Entry
      'PTyp= 'SRTRM'
      '     MuliSelect= True
      '     Show help of all the fields in the RmMst Table. This field stores the fields
      '     of RmMst Table by which sorting has to be done in the Raw Material Master Entry
      'PTyp= 'SRTRR'
      '     MuliSelect= True
      '     Show help of all the fields in the RmRt Table. This field stores the fields
      '     of RmRt Table by which sorting has to be done in the Raw Material Rate Master Entry
      'PTyp= 'USR'
      '     MuliSelect= True
      '     Show help of all the production Locs
  '*** PValue
      'PTyp= 'CHR'
      '     Show help of all the Locations which are not Process Locations
      '     This field is used to store the default 'From RmLoc' for that Character. It will
      '     be used in the Transaction Entry for the TC-CHR combination
      'PTyp= 'DAANAFLD'
      '     Show help for Param('YN')
      'PTyp= 'DMCTG'
      '     Show help for Design Unit Of Measurement
      'PTyp= 'LABSCD'
      '     Show help for LABQW. This field specifies whether the calculation of Labour
      '     Value is based on Qty Or Weight
      'PTyp= 'PARENTCD'
      '     Show help for RmCtg
      'PTyp= 'PRDCTG'
      '     Show help for PRDSEQ. This field stores the Sequence of locations the Designs
      '     of this Prd Category has to pass through while in production
      'PTyp= 'PRDSEQ'
      'PTyp= 'RMCTG'
      '     Show all the Raw Material Codes of the category which are not custom codes.
      '     This field is used to store the Pure Rm Code for this category. This RmCd is
      '     used to calculate the pure wt in many reports.
      'PTyp= 'RMSCTG'
      '     Show help for the 'RATEBY' option. This field specifies how the Rate for
      '     raw materials other than Gold & Platinum are calculated
      '     (i.e By Size, Pointer or Both Size)
      'PTyp= 'TC'
      '     Show the help for TC Type. This field specifies the Tc Type to which this
      '     this field belongs
      'PTyp= 'TCTYP'
      '     Show all the TC's where TcType= PMCd. This field is used to store the default
      '     TC for the current Tc Type
      'PTyp= 'WORK'
      '     Show help of all the Process Locations (Default Process Loc)
  '*** PValue1
      'PTyp= 'CHR'
      '     Show help of all the Process Locations
      'PTyp= 'DAANAFLD'
      '     Show help for Param('YN')
      'PTyp= 'RMSCTG'
      '     Show help for 'YN' (Yes/ No). Here the user has to specify whether the
      '     raw materail transactions are Stock Rate wise or not
      'PTyp= 'PARENTCD'
      '     Show help for RmSCtg
      'PTyp= 'TCTYP'
      '     This field is used to specify the Entry Type of the current Tc Type
      '     The Entry Types are hard coded in the System
      'PTyp= 'USR'
      '     MuliSelect= True
      '     Show help of all the non production Locs
  '*** PValue2
      'PTyp= 'CHR'
      '     Show help of 'DC' (Debit/Credit). This field is used to store the default
      '     Raw Material 'DC' to be used in the Transaction Entry for the TC-CHR combination
      'PTyp= 'RMSCTG'
      '     Show help for 'YN' (Yes/ No). Here the user has to specify whether the
      '     Rm Ptr field in the Design Entry is editable or not
  '*** PValue3
      'PTyp= 'RMSCTG'
      '     Show help for 'YN' (Yes/ No). Here the user has to specify whether the
      '     raw materail transactions are Pointer wise or not

  '****** Sachin 2.14.0 [22-11-2006] - [Multi-Company] ******
  'If (UCase(ADC("wvPTyp")) = "WIWSEQ" Or UCase(ADC("wvPTyp")) = "WIWLONG") And ADC("wvPCoCd") <> ctSelfCoCd Then
  '    Cancel = True: ErrMsg = "'WIWSEQ' And 'WIWLONG' Can Be Updated Only Through Company 'ZZZ'"
  '    GrdvParam.AllowAdd = False: GrdvParam.AllowDelete = False
  'Else
  '    GrdvParam.AllowAdd = True: GrdvParam.AllowDelete = True
  'End If
  '****** Sachin 2.14.0 [22-11-2006] - [Multi-Company] ******

  Select Case ColName
  Case Is = UCase("vPMCd")
    Select Case UCase(ADC("wvPTyp"))
    Case Is = "CHR"
      Call HlpList.PMCd("TC")
    Case Is = "DEFCHR"
      Call HlpList.PMCd("TC")
    Case Is = "CSEQ"
      Call HlpList.PMCd("CELL")
    Case Is = "CHKTOL"
      Call HlpList.Loc(ADC("wvPCoCd"), "'R'")
    'pg.11-help for TC
    Case Is = "GSTTCCHR"
      Call HlpList.PMCd("TC")
    Case Is = "UNLOCK"  '6.1_51 should not edit main code
      Cancel = True: ErrMsg = "Cannot Edit MainCode": Exit Sub
    End Select
  Case Is = UCase("vPSCd")
    Select Case UCase(ADC("wvPTyp"))
    Case Is = "CSEQ"
      Call HlpList.vPMCd(ADC("wvPCoCd"), "PRCSSEQ") 'EmrSv -
    'pg.11-help for Character
    Case Is = "GSTTCCHR"
      Call HlpList.vPSCd(ADC("wvPCoCd"), "CHR", GrdvParam.Value(RowNum, "vPMCd"))
    End Select
  
  '*** Jay 3.2.0 [AutoTree]
  Case Is = UCase("vPDesc")
    Select Case UCase(ADC("wvPTyp"))
    'If UCase(adc("wvPTyp")) = "TRDEF" Then
    Case Is = "TRDEF"
      Cancel = True: ErrMsg = "Cannot Edit Description": Exit Sub
  '*** Jay 3.2.0 [AutoTree]
    Case Is = "CHKTOL"
      HlpList.MultiSelect = True
      Call HlpList.PMCd("BCHR")
    Case Is = "UNLOCK"  '6.1_51 description should not edit
      Cancel = True: ErrMsg = "Cannot Edit Description": Exit Sub
    End Select
    
  Case Is = UCase("vPDesc225")
    Select Case UCase(ADC("wvPTyp"))
    Case Is = "CHR"
      Call HlpList.Loc(ADC("wvPCoCd"), , "'P'")
    Case Is = "PRDSEQ", "WIWSEQ", "WIWLONG"
      HlpList.MultiSelect = True
      Call HlpList.Loc(ADC("wvPCoCd"), "'P'", , IIF(ADC("wvPValue3") = "Y", "'Y'", "'N'"))
    ' ***
    Case Is = "PRCSSEQ", "RWIWSEQ", "RWIWLONG"
      HlpList.MultiSelect = True
      Call HlpList.Loc(ADC("wvPCoCd"), "'R'", , IIF(ADC("wvPValue3") = "Y", "'Y'", "'N'"))
    'mw.98-help for subshape
    Case Is = "CHKTOL"
      Call HlpList.PMCd("SUBSHP")
    End Select
  Case Is = UCase("vPValue")
    Select Case UCase(ADC("wvPTyp"))
    Case Is = "CHR"
      Call HlpList.Loc(ADC("wvPCoCd"), , "'P'")
    Case Is = "DEFCHR"
      Call HlpList.vPSCd(ADC("wvPCoCd"), "CHR", GrdvParam.Value(RowNum, "vPMCd"))
    Case Is = "PRDSEQ"
      '$$$ Call HlpList.PMCd("PLNSEQ")
'?????????
    Case Is = "PTYP"
      Call HlpList.PMCd("SYSDFN")
    Case Is = "WORK"
      Call HlpList.Loc(ADC("wvPCoCd"), "'P'")
    '*** Jay 3.2.0 [AutoTree]
    Case Is = "TRDEF"
      If UCase(GrdvParam.Value(RowNum, "vPMCd")) <> UCase("IssVch") And UCase(GrdvParam.Value(RowNum, "vPMCd")) <> UCase("RecVch") Then
        Call HlpList.Loc(ADC("wvPCoCd"), "'S','L'")
      Else
        Call HlpList.vPSCd(ADC("wvPCoCd"), "CHR", "MV")
      End If
    '*** Jay 3.2.0 [AutoTree]
    
    Case Is = "CSEQ"
      Call HlpList.vPMCd(ADC("wvPCoCd"), "PRDSEQ")
    'pg.11-help for Gst code
    Case Is = "GSTTCCHR"
      Call HlpList.PMCd("GSTCD")
    End Select
  Case Is = UCase("vPValue1")
    Select Case UCase(ADC("wvPTyp"))
    Case Is = "CHKTOL"
      HlpList.MultiSelect = True  'IG BCHR-Added MultiSelect for RmCtg
      Call HlpList.PMCd("RMCTG")

    Case Is = "CHR"
      '**********Bhavna 5.1.0
      If GrdvParam.Value(RowNum, "vPMCd") = "FB" Then
        HlpList.MultiSelect = True
      Else
        HlpList.MultiSelect = False
      End If
      '**********Bhavna 5.1.0
      Call HlpList.Loc(ADC("wvPCoCd"), "'P'")
    
    Case Is = "PRDSEQ"
      Call HlpList.vPMCd(ADC("wvPCoCd"), "PRCSSEQ")
    'EmrSv -
    Case Is = "PRCSSEQ"
      Call HlpList.PMCd("PRDLINE")
          
    '****Geeta***Emr208
    Case Is = "WORK"
      Call HlpList.Loc(ADC("wvPCoCd"), "'L'")
   '****Geeta***Emr208
    'mw.98-help for Rm category
    Case Is = "CHKTOL"
      HlpList.MultiSelect = True
      Call HlpList.PMCd("RMCTG")
    End Select
  Case Is = UCase("vPValue2")
    Select Case UCase(ADC("wvPTyp"))
    Case Is = "CHR"
      Call HlpList.PMCd("DC")
    'EmrSv -
    Case Is = "PRDSEQ"
      Call HlpList.PMCd("CELL")
    'mw.98-help for Rm Sub category
    Case Is = "CHKTOL"
      HlpList.MultiSelect = True
      Call HlpList.PSCd("RMSCTG", "", True)
    End Select
  'mw.98-help for production type
  Case Is = UCase("vPValue3")
    Select Case UCase(ADC("wvPTyp"))
    Case Is = "CHKTOL"
      ' SRI REMOVING below check as it is cornering the user into no way to get out of the row
      ' when RmCtg is removed but there is a value in this column
      'IG BCHR- The Value in this field will be allowed only if
      'RMCtg field will have the value as ‘D’ Or ‘C’ or both.
      'Cancel = Not (("," + GrdvParam.Value(RowNum, "vPValue1") + "," Like "*,C,*") Or ("," + GrdvParam.Value(RowNum, "vPValue1") + "," Like "*,D,*"))
      'If Cancel = True Then ErrMsg = "Cannot Enter Value if D Or C not in Rm Category": Exit Sub
      Call HlpList.PMCd("PRDTYP")
    End Select
' ###################################  Manoj #### Ver: 2.0.6  ###################################
' Date:   20/01/2004
  Case Is = UCase("vPValidYN")
    Select Case UCase(ADC("wvPTyp"))
    Case Is = "WORK"
      Call HlpList.PMCd("YN")
    End Select
' ###########################################  Manoj  ###########################################
'************Geeta*********Emr207*********
    Case Is = UCase("vPValue4")
      Select Case UCase(ADC("wvPTyp"))
       Case Is = "CHR", "FGSUBLOC"
        HlpList.MultiSelect = True
        Call HlpList.PMCd("USR")
       Case Is = "CHKTOL"
        HlpList.MultiSelect = True
        Call HlpList.PSCd("RMSCTG", "", True)
    End Select
'************Geeta*********Emr207*********
'MMAck-When vPtyp=Chr and Tc=MV then help for AckYN will work else set default value 'N' for other TC
  Case Is = UCase("wAckYn")
    Select Case UCase(ADC("wvPTyp"))
    Case Is = "CHR"
      If GrdvParam.Value(RowNum, "vPMCd") = "MV" Then Call HlpList.PMCd("YN")
      '5716 tagprc If GrdvParam.Value(RowNum, "vPMCd") <> "MV" Then GrdvParam.Value(RowNum, "WACKYN") = "N": Cancel = True
    End Select
  End Select
End Sub
Private Sub GrdvParam_ColChange(ByVal RowNum As Integer, ByVal ColName As String, ByVal NewValue As Variant, Cancel As Boolean, ErrMsg As String)
      '*** Check From TcStr If Blank Is Allowed ***
      'ws_BlnkYN = TsBlank from TcStr(TsTbl='Param' and TsTyp= PTyp and TsSTyp=''
      '                               and TsIdName= ColName)
  '*** PMCd
      'PTyp is 'CHR'
      '     If ws_BlnkYN= 'N' then valid PMCd('TC')
      '     If ws_BlnkYN= 'Y' then '' or valid PMCd('TC')
      'PTyp is 'DAANACD'
      '     If ws_BlnkYN= 'N' then valid PMCd('DAANAFLD')
      '     If ws_BlnkYN= 'Y' then '' or valid PMCd('DAANAFLD')
      'PTyp is 'DAANAFLD'
      '     Enter A Proper Integer > 0 & < 100
      'PTyp is 'LABSCD'
      '     If ws_BlnkYN= 'N' then valid PMCd('LABMCD')
      '     If ws_BlnkYN= 'Y' then '' or valid PMCd('LABMCD')
      'PTyp is 'RMSCTG', 'RMCOL'
      '     If ws_BlnkYN= 'N' then valid PMCd('RMCTG')
      '     If ws_BlnkYN= 'Y' then '' or valid PMCd('RMCTG')
      'PTyp is 'SUBPRC'
      '     If ws_BlnkYN= 'N' then valid LocCd from Loc(LocTyp= 'P')
      '     If ws_BlnkYN= 'Y' then '' or valid LocCd from Loc(LocTyp= 'P')
      'PTyp is 'PRTCD'
      '     If ws_BlnkYN= 'N' then valid two digit number
      '     If ws_BlnkYN= 'Y' then '' or valid two digit number
  '*** PSCd
      '     If ws_BlnkYN= 'N' then this field cannot be Blank
  '*** PDesc225
      'PTyp is 'CHR'
      '     If ws_BlnkYN= 'N' then valid LocCd from Loc(LocTyp<> 'P')
      '     If ws_BlnkYN= 'Y' then '' or valid LocCd from Loc(LocTyp<> 'P')
      'PTyp is 'GLDAS', 'IWTFRORD', 'CPYRT'
      '     valid sequence of PMCd('RMCTG') (separated by comma)
      'PTyp is 'PRDSEQ', 'WIWSEQ', 'WIWLONG'
      '     valid sequence of Process Locations i.e. valid LocCd from Loc(LocTyp= 'P')
      '     separated by commas
      'PTyp is 'RMSCTG'
      '     If ws_BlnkYN= 'N' then valid PMCd('YN')
      '     If ws_BlnkYN= 'Y' then '' or valid PMCd('YN')
      'PTyp is 'SRTBAG'
      '     If ws_BlnkYN= 'N' then valid sequence of fields from the Bag table
      '     If ws_BlnkYN= 'Y' then '' or valid sequence of fields from the Bag table
      'PTyp is 'SRTINBAG'
      '     If ws_BlnkYN= 'N' then valid sequence of fields from the InvFgd table
      '     If ws_BlnkYN= 'Y' then '' or valid sequence of fields from the InvFgd table
      'PTyp is 'SRTLR'
      '     If ws_BlnkYN= 'N' then valid sequence of fields from the LabRt table
      '     If ws_BlnkYN= 'Y' then '' or valid sequence of fields from the LabRt table
      'PTyp is 'SRTPP'
      '     If ws_BlnkYN= 'N' then valid sequence of fields from the PrdPts table
      '     If ws_BlnkYN= 'Y' then '' or valid sequence of fields from the PrdPts table
      'PTyp is 'SRTPRHLP'
      '     If ws_BlnkYN= 'N' then valid sequence of fields from the TxndZ table
      '     If ws_BlnkYN= 'Y' then '' or valid sequence of fields from the TxndZ table
      'PTyp is 'SRTRM'
      '     If ws_BlnkYN= 'N' then valid sequence of fields from the RmMst table
      '     If ws_BlnkYN= 'Y' then '' or valid sequence of fields from the RmMst table
      'PTyp is 'SRTRR'
      '     If ws_BlnkYN= 'N' then valid sequence of fields from the RmRt table
      '     If ws_BlnkYN= 'Y' then '' or valid sequence of fields from the RmRt table
      'PTyp is 'TC'
      '     If ws_BlnkYN= 'N' then valid PSCd('CHR', NewValue)
      '     If ws_BlnkYN= 'Y' then '' or valid PSCd('CHR', NewValue)
      'PTyp is 'INGR'
      '     If ws_BlnkYN= 'N' then valid PMCd('INGRPRCP', NewValue)
      '     If ws_BlnkYN= 'Y' then '' or valid PMCd('INGRPRCP', NewValue)
      'PTyp is 'USR'
      '     If ws_BlnkYN= 'N' then valid sequence of locations where LocTyp= 'P'
      '     If ws_BlnkYN= 'Y' then '' or valid sequence of locations where LocTyp= 'P'
  '*** PValue
      'PTyp is 'CHR'
      '     If ws_BlnkYN= 'N' then valid LocCd from Loc(LocTyp<> 'P')
      '     If ws_BlnkYN= 'Y' then '' or valid LocCd from Loc(LocTyp<> 'P')
      'PTyp is 'DAANAFLD'
      '     If ws_BlnkYN= 'N' then valid PMCd('YN')
      '     If ws_BlnkYN= 'Y' then '' or valid PMCd('YN')
      'PTyp is 'DMCTG'
      '     If ws_BlnkYN= 'N' then valid PMCd('DMUOM')
      '     If ws_BlnkYN= 'Y' then '' or valid PMCd('DMUOM')
      'PTyp is 'LABSCD'
      '     If ws_BlnkYN= 'N' then valid PMCd('LABQW')
      '     If ws_BlnkYN= 'Y' then '' or valid PMCd('LABQW')
      'PTyp is 'PARENTCD'
      '     If ws_BlnkYN= 'N' then valid PMCd('RMCTG')
      '     If ws_BlnkYN= 'Y' then '' or valid PMCd('RMCTG')
      'PTyp is 'PRDCTG'
      '     If ws_BlnkYN= 'N' then valid PMCd('PRDSEQ')
      '     If ws_BlnkYN= 'Y' then '' or valid PMCd('PRDSEQ')
      'PTyp is 'PRDSEQ'
      '     If ws_BlnkYN= 'N' then valid PMCd('PLNSEQ')
      '     If ws_BlnkYN= 'Y' then '' or valid PMCd('PLNSEQ')
      'PTyp is 'RMCTG'
      '     If ws_BlnkYN= 'N' then valid RmCd from RmMst(RmZ= 'N')
      '     If ws_BlnkYN= 'Y' then '' or valid RmCd from RmMst(RmZ= 'N')
      'PTyp is 'RMSCTG'
      '     If ws_BlnkYN= 'N' then valid PMCd('RATEBY')
      '     If ws_BlnkYN= 'Y' then '' or valid PMCd('RATEBY')
      'PTyp is 'TC'
      '     If ws_BlnkYN= 'N' then valid PMCd('TCTYP')
      '     If ws_BlnkYN= 'Y' then '' or valid PMCd('TCTYP')
      'PTyp is 'TCTYP'
      '     If ws_BlnkYN= 'N' then valid PMCd('TC')
      '     If ws_BlnkYN= 'Y' then '' or valid PMCd('TC')
      'PTyp is 'WORK'
      '     If ws_BlnkYN= 'N' then valid LocCd from Loc(LocTyp= 'P')
      '     If ws_BlnkYN= 'Y' then '' or valid LocCd from Loc(LocTyp= 'P')
  '*** PNum
      'PTyp is 'DMCTG'
      '     If ws_BlnkYN= 'N' then >0
      'PTyp is 'OUNCE'
      '     If ws_BlnkYN= 'N' then <> 0
      'PTyp is 'VACTG'
      '     If ws_BlnkYN= 'N' then <> 0
  '*** PValue1
      'PTyp is 'CHR'
      '     If ws_BlnkYN= 'N' then valid LocCd from Loc(LocTyp= 'P')
      '     If ws_BlnkYN= 'Y' then '' or valid LocCd from Loc(LocTyp= 'P')
      'PTyp is 'DAANAFLD'
      '     If ws_BlnkYN= 'N' then valid PMCd('YN')
      '     If ws_BlnkYN= 'Y' then '' or valid PMCd('YN')
      'PTyp is 'PARENTCD'
      '     If ws_BlnkYN= 'N' then valid PSCd('RMSCTG', PMCd)
      '     If ws_BlnkYN= 'Y' then '' or valid PMCd('RMSCTG', PMCd)
      'PTyp is 'RMSCTG'
      '     If ws_BlnkYN= 'N' then valid PMCd('YN')
      '     If ws_BlnkYN= 'Y' then '' or valid PMCd('YN')
      'PTyp is 'USR'
      '     If ws_BlnkYN= 'N' then valid sequence of locations where LocTyp<> 'P'
      '     If ws_BlnkYN= 'Y' then '' or valid sequence of locations where LocTyp<> 'P'
  '*** PValue2
      'PTyp is 'CHR'
      '     If ws_BlnkYN= 'N' then valid PMCd('DC')
      '     If ws_BlnkYN= 'Y' then '' or valid PMCd('DC')
      'PTyp is 'RMSCTG'
      '     If ws_BlnkYN= 'N' then valid PMCd('YN')
      '     If ws_BlnkYN= 'Y' then '' or valid PMCd('YN')
  '*** PValue3
      'PTyp is 'RMSCTG'
      '     If ws_BlnkYN= 'N' then valid PMCd('YN')
      '     If ws_BlnkYN= 'Y' then '' or valid PMCd('YN')
  
  Dim ws_BlnkYN As String, i As Integer, ws_Loc() As String
  With GrdvParam
    ws_BlnkYN = moCn.GetFldVal("Select TsBlank from TcStr where TsTbl= 'vParam' and " + _
                               "TsTyp='" + ADC("wvPTyp") + "' and TsSTyp= '' and " + _
                               "TsIdName='" + ColName + "'")
    Select Case ColName
    Case Is = UCase("vPMCd")
      Select Case UCase(ADC("wvPTyp"))
      Case Is = "CHR"
        If ws_BlnkYN = "N" And NewValue = "" Then _
           Cancel = True: ErrMsg = "Tc Should Be Entered": Exit Sub
        Cancel = (NewValue <> "") And (Not moCn.RecSeek("Select PMCd From Param where PTyp= 'TC' " + _
                                      "and PMCd = '" + NewValue + "'"))
        ErrMsg = "Invalid Tc"
        If Cancel Then Exit Sub
        
        If NewValue = "DT" Or NewValue = "BV" Then
          GrdvParam.ColProp("vPValue3").Heading = "RepairChr"
        Else
          If "Y" = moCn.GetFldVal("Select 'Y' from Param where Ptyp= 'TC' and PMCd='" + NewValue + "' and PValue='JT' and PValue1='N'") Then
            GrdvParam.ColProp("vPValue3").Heading = "MRP Chr"
          Else
            GrdvParam.ColProp("vPValue3").Heading = "TrialChr"
          End If
        End If
      Case Is = "DEFCHR"
        If ws_BlnkYN = "N" And NewValue = "" Then _
           Cancel = True: ErrMsg = "Tc Should Be Entered": Exit Sub
        Cancel = (NewValue <> "") And (Not moCn.RecSeek("Select PMCd From Param where PTyp= 'TC' " + _
                                      "and PMCd = '" + NewValue + "'"))
        ErrMsg = "Invalid Tc": Exit Sub
      ' ****
      Case Is = "CSEQ"
       If ws_BlnkYN = "N" And NewValue = "" Then _
           Cancel = True: ErrMsg = "Cell Name Should Be Entered": Exit Sub
        Cancel = (NewValue <> "") And (Not moCn.RecSeek("Select PMCd From Param where PTyp= 'CELL' " + _
                                      "and PMCd = '" + NewValue + "'"))
        ErrMsg = "Invalid Cell Name ": Exit Sub
      Case Is = "CHKTOL"
      If ws_BlnkYN = "N" And NewValue = "" Then _
           Cancel = True: ErrMsg = "Process Location Should Be Entered": Exit Sub
        Cancel = (NewValue <> "") And (Not moCn.RecSeek("Select LocCd From Loc where " + _
                                        "LocCoCd = '" + ADC("wvPCoCd") + "' and " + _
                                        "LocTyp= 'R' and LocCd = '" + NewValue + "'"))
        ErrMsg = "Invalid Process Location": Exit Sub
      'pg.11- checking valid Tc
      Case Is = "GSTTCCHR"
        If ws_BlnkYN = "N" And NewValue = "" Then _
           Cancel = True: ErrMsg = "Tc Should Be Entered": Exit Sub
        Cancel = (NewValue <> "") And (Not moCn.RecSeek("Select PMCd From Param where PTyp= 'TC' " + _
                                      "and PMCd = '" + NewValue + "'"))
        ErrMsg = "Invalid Tc": Exit Sub
      End Select
      If ws_BlnkYN = "N" And NewValue = "" Then _
         Cancel = True: ErrMsg = "Value Should Be Entered": Exit Sub
    Case Is = UCase("vPSCd")
    Select Case UCase(ADC("wvPTyp"))
      ' ***
      Case Is = "CSEQ" 'EmrSv -
        If ws_BlnkYN = "N" And NewValue = "" Then _
           Cancel = True: ErrMsg = "Process Sequence Should Be Entered": Exit Sub
        Cancel = (NewValue <> "") And Not moCn.RecSeek("Select vPMCd from vParam where " + _
                "vPTyp= 'PRCSSEQ' and vPMCd= '" + NewValue + "'")
        ErrMsg = "Invalid Process Sequence Code": Exit Sub
      'pg.11- checking valid Chr
      Case Is = "GSTTCCHR"
        If ws_BlnkYN = "N" And NewValue = "" Then _
           Cancel = True: ErrMsg = "Chr Should Be Entered": Exit Sub
        Cancel = (NewValue <> "") And Not moCn.RecSeek("Select vPMCd from vParam where vPCoCd='" + ADC("wvPCoCd") + "' and " + _
                "vPTyp= 'CHR' and vPMCd='" + GrdvParam.Value(RowNum, "vPMCd") + "' and vPSCd= '" + NewValue + "'")
        ErrMsg = "Invalid Chr": Exit Sub
      End Select
      If ws_BlnkYN = "N" And NewValue = "" Then _
         Cancel = True: ErrMsg = "Value Should Be Entered": Exit Sub
    
    
    Case Is = UCase("vPDesc")
      Select Case UCase(ADC("wvPTyp"))
      Case Is = "CHKTOL"
        Dim ws_BChrArr() As String
        ws_BChrArr = Split(NewValue, ",")
        If IsArray(ws_BChrArr) Then
          For i = 0 To UBound(ws_BChrArr)
            Cancel = Not moCn.RecSeek("Select PMCd From Param where PTyp= 'BCHR' and " + _
                                      "PMCd = '" + ws_BChrArr(i) + "'")
            If Cancel = True Then ErrMsg = "BChr (" + CStr(i + 1) + ") Value - " + ws_BChrArr(i) + " Is An Invalid ": Exit Sub
          Next i
        End If
      End Select
      If ws_BlnkYN = "N" And NewValue = "" Then _
         Cancel = True: ErrMsg = "Value Should Be Entered": Exit Sub
    
    Case Is = UCase("vPDesc225")
      Select Case UCase(ADC("wvPTyp"))
      Case Is = "CHR"
        If ws_BlnkYN = "N" And NewValue = "" Then _
           Cancel = True: ErrMsg = "To RM Location Should Be Entered": Exit Sub
        Cancel = (NewValue <> "") And (Not moCn.RecSeek("Select LocCd From Loc where " + _
                                      "LocCoCd = '" + ADC("wvPCoCd") + "' and " + _
                                      "LocTyp<> 'P' and LocCd = '" + NewValue + "'"))
        ErrMsg = "Invalid To RM Location": Exit Sub
      '*******Geeta***Crm**2.12**
      Case Is = "PRDSEQ", "WIWSEQ", "WIWLONG", "CRMWIW"
        Dim ws_LocArr() As String
        ws_LocArr = Split(NewValue, ",")
        If IsArray(ws_LocArr) Then
          For i = 0 To UBound(ws_LocArr)
            '****** Sachin 2.14.0 - 03-11-2006 [Multi-Company] - vPValue3 Check Added Below ******
            Cancel = Not moCn.RecSeek("Select LocCd From Loc where LocTyp= 'P' And LocPrntYN='" + IIF(UCase(ADC("wvPValue3")) = "Y", "Y", "N") + "' And " + _
                                      "LocCoCd = '" + ADC("wvPCoCd") + "' and " + _
                                      "LocCd = '" + ws_LocArr(i) + "'")
            If Cancel = True Then ErrMsg = "Loc(" + CStr(i + 1) + "), Value " + ws_LocArr(i) + " Is An Invalid" + " " + IIF(UCase(ADC("wvPValue3")) = "Y", "Parent Code", "Production Location"): Exit Sub
            '****** Sachin 2.14.0 - 03-11-2006 [Multi-Company] ******
          Next i
        End If
      ' ****
      Case Is = "PRCSSEQ", "RWIWSEQ", "RWIWLONG"
        Dim ws_PrcArr() As String
        ws_PrcArr = Split(NewValue, ",")
        If IsArray(ws_PrcArr) Then
          For i = 0 To UBound(ws_PrcArr)
            Cancel = Not moCn.RecSeek("Select LocCd From Loc where LocTyp= 'R' And LocPrntYN='" + IIF(UCase(ADC("wvPValue3")) = "Y", "Y", "N") + "' And " + _
                                      "LocCoCd = '" + ADC("wvPCoCd") + "' and " + _
                                      "LocCd = '" + ws_PrcArr(i) + "'")
            If Cancel = True Then ErrMsg = "Process(" + CStr(i + 1) + "), Value " + ws_PrcArr(i) + " Is An Invalid " + IIF(UCase(ADC("wvPValue3")) = "Y", "Parent Code", "Process"): Exit Sub
          Next i
        End If
      'mw.98-checking valid sub shape, blank allowed
      Case Is = "CHKTOL"
        Cancel = (NewValue <> "") And Not moCn.RecSeek("Select PMCd From Param where PTyp= 'SUBSHP' and " + _
                                      "PMCd = '" + NewValue + "'")
        ErrMsg = "Invalid Sub Shape": Exit Sub
      
      End Select
      If ws_BlnkYN = "N" And NewValue = "" Then _
         Cancel = True: ErrMsg = "Value Should Be Entered": Exit Sub
    Case Is = UCase("vPValue")
      Select Case UCase(ADC("wvPTyp"))
      Case Is = "CHR"
        If ws_BlnkYN = "N" And NewValue = "" Then _
           Cancel = True: ErrMsg = "From RM Location Should Be Entered": Exit Sub
        Cancel = (NewValue <> "") And (Not moCn.RecSeek("Select LocCd From Loc where " + _
                                      "LocCoCd = '" + ADC("wvPCoCd") + "' and " + _
                                      "LocTyp<> 'P' and LocCd = '" + NewValue + "'"))
        ErrMsg = "Invalid From RM Location": Exit Sub
      Case Is = "DEFCHR"
        If ws_BlnkYN = "N" And NewValue = "" Then _
           Cancel = True: ErrMsg = "Default Character Should Be Entered": Exit Sub
        Cancel = (NewValue <> "") And (Not moCn.RecSeek("Select vPSCd From vParam where " + _
                                      "vPCoCd= '" + ADC("wvPCoCd") + "' and vPTyp= 'CHR' and " + _
                                      "vPMCd ='" + GrdvParam.Value(RowNum, "vPMCd") + "' and " + _
                                      "vPSCd= '" + NewValue + "' and vpValidYN<>'N' "))
        ErrMsg = "Invalid Voucher Characher": Exit Sub
      Case Is = "PRDSEQ"
        '$$$
        'If ws_BlnkYN = "N" And NewValue = "" Then _
        '   Cancel = True: ErrMsg = "Plan Sequence Should Be Entered": Exit Sub
        'Cancel = (NewValue <> "") And (Not moCn.RecSeek("Select PMCd From Param where " + _
        '                              "PTyp= 'PLNSEQ' and PMCd = '" + NewValue + "'"))
        'ErrMsg = "Invalid Plan Sequence": Exit Sub
        '$$$
      Case Is = "PTYP"
        If ws_BlnkYN = "N" And NewValue = "" Then _
           Cancel = True: ErrMsg = "Specify Whether The Param Type Is Microway/System/Other": Exit Sub
        Cancel = (NewValue <> "") And (Not moCn.RecSeek("Select PMCd From Param where " + _
                                      "PTyp= 'SYSDFN' and PMCd = '" + NewValue + "'"))
        ErrMsg = "Invalid Patameter Category": Exit Sub
      Case Is = "WORK"
        If ws_BlnkYN = "N" And NewValue = "" Then _
           Cancel = True: ErrMsg = "Process Location Should Be Entered": Exit Sub
        Cancel = (NewValue <> "") And (Not moCn.RecSeek("Select LocCd From Loc where " + _
                                      "LocCoCd = '" + ADC("wvPCoCd") + "' and " + _
                                      "LocTyp= 'P' and LocCd = '" + NewValue + "'"))
        ErrMsg = "Invalid Process Location": Exit Sub
      ' ***
      Case Is = "CSEQ"
        If ws_BlnkYN = "N" And NewValue = "" Then _
           Cancel = True: ErrMsg = "Production Sequence Should Be Entered": Exit Sub
        Cancel = (NewValue <> "") And Not moCn.RecSeek("Select vPMCd from vParam where " + _
                "vPTyp= 'PRDSEQ' and vPMCd= '" + NewValue + "'")
        ErrMsg = "Invalid Production Sequence Code": Exit Sub
        
      '*** Jay 3.2.0 [AutoTree]
      Case Is = "TRDEF"
        If ws_BlnkYN = "N" And NewValue = "" Then _
           Cancel = True: ErrMsg = "Location Should Be Entered": Exit Sub
        If UCase(GrdvParam.Value(RowNum, "vPMCd")) <> UCase("IssVch") And UCase(GrdvParam.Value(RowNum, "vPMCd")) <> UCase("RecVch") Then
          Cancel = (NewValue <> "") And _
                    (Not moCn.RecSeek("Select LocCd From Loc where LocCoCd = '" + ADC("wvPCoCd") + "' and " + _
                      "LocTyp in " + IIF(UCase(GrdvParam.Value(RowNum, "vPMCd")) <> UCase("IssVch") And _
                                         UCase(GrdvParam.Value(RowNum, "vPMCd")) <> UCase("RecVch"), "('S','L')", "('P')") + _
                      "and LocCd = '" + NewValue + "' and LocValidYn= 'Y'"))
          ErrMsg = "Invalid Location": Exit Sub
        Else
          Cancel = (NewValue <> "") And (Not moCn.RecSeek("Select 'x' From vParam  where vPCoCd = '" + ADC("wvPCoCd") + "' and " + _
                          "vPTyp= 'CHR' and vPSCd = '" + NewValue + "' and vpValidYN<>'N' "))
          ErrMsg = "Invalid Voucher Character": Exit Sub
        End If
      '*** Jay 3.2.0 [AutoTree]
      'pg.11- checking valid Gst Code
      Case Is = "GSTTCCHR"
        If ws_BlnkYN = "N" And NewValue = "" Then _
           Cancel = True: ErrMsg = "Gst Code Should Be Entered": Exit Sub
        Cancel = (NewValue <> "") And (Not moCn.RecSeek("Select PMCd From Param where PTyp= 'GSTCD' " + _
                                      "and PMCd = '" + NewValue + "'"))
        ErrMsg = "Invalid Gst Code": Exit Sub
      End Select
      If ws_BlnkYN = "N" And NewValue = "" Then _
         Cancel = True: ErrMsg = "Value Should Be Entered": Exit Sub
    Case Is = UCase("vPValue1")
      Select Case UCase(ADC("wvPTyp"))
      Case Is = "CHR"
        If ws_BlnkYN = "N" And NewValue = "" Then _
           Cancel = True: ErrMsg = "Process Location Should Be Entered": Exit Sub
           
        If GrdvParam.Value(RowNum, "vPMCd") = "FB" Then
        
          '**********Bhavna 5.1.0  multiselect for 'FB' TC
          Dim ws_LocArr1() As String
          ws_LocArr1 = Split(NewValue, ",")
          If IsArray(ws_LocArr1) Then
            For i = 0 To UBound(ws_LocArr1)
              Cancel = (NewValue <> "") And (Not moCn.RecSeek("Select LocCd From Loc where " + _
                                        "LocCoCd = '" + ADC("wvPCoCd") + "' and " + _
                                        "LocTyp= 'P' and LocCd = '" + ws_LocArr1(i) + "'"))
              If Cancel = True Then ErrMsg = "Loc " + ws_LocArr1(i) + " Is An Invalid Process Location": Exit Sub
              '****** Sachin 2.14.0 - 03-11-2006 [Multi-Company] ******
            Next i
          End If
        Else
              Cancel = (NewValue <> "") And (Not moCn.RecSeek("Select LocCd From Loc where " + _
                                        "LocCoCd = '" + ADC("wvPCoCd") + "' and " + _
                                        "LocTyp= 'P' and LocCd = '" + NewValue + "'"))
              ErrMsg = "Invalid Process Location": Exit Sub
        End If
        '**********Bhavna 5.1.0  multiselect for 'FG' locns
        '***Geeta***Emr208
        Case Is = "WORK"
         Cancel = (NewValue <> "") And (Not moCn.RecSeek("Select LocCd From Loc where " + _
                                      "LocCoCd = '" + ADC("wvPCoCd") + "' and " + _
                                      "LocTyp= 'L' and LocCd = '" + NewValue + "'"))
        ErrMsg = "Invalid Loss Location": Exit Sub
        ' ****
        Case Is = "PRDSEQ"
           Cancel = (Not moCn.RecSeek("Select vPMCD From vParam where " + _
                                      "vPCoCd = '" + ADC("wvPCoCd") + "' and " + _
                                      "vPTyp= 'PRCSSEQ' and vPMCD = '" + NewValue + "'"))
        ErrMsg = "Invalid Process Sequence": Exit Sub
        'EmrSv -
        Case Is = "PRCSSEQ"
           Cancel = (Not moCn.RecSeek("Select PMCD From Param where " + _
                                      "PTyp= 'PRDLINE' and PMCD = '" + NewValue + "'"))
        ErrMsg = "Invalid Production Line": Exit Sub
        
        
        '***Geeta***Emr208
        'mw.98-checking valid Rm Category
        Case Is = "CHKTOL"
          Dim ws_RmCtgArr() As String
          ws_RmCtgArr = Split(NewValue, ",")
          If IsArray(ws_RmCtgArr) Then
            For i = 0 To UBound(ws_RmCtgArr)
              Cancel = Not moCn.RecSeek("Select PMCd From Param where PTyp= 'RMCTG' and " + _
                                      "PMCd = '" + ws_RmCtgArr(i) + "'")
              If Cancel = True Then ErrMsg = "RMCTG (" + CStr(i + 1) + ") Value - " + ws_RmCtgArr(i) + " Is An Invalid ": Exit Sub
            Next i
          End If
    End Select
      If ws_BlnkYN = "N" And NewValue = "" Then _
        Cancel = True: ErrMsg = "Value Should Be Entered": Exit Sub
    Case Is = UCase("vPValue2")
      Select Case UCase(ADC("wvPTyp"))
      Case Is = "CHR"
        If ws_BlnkYN = "N" And NewValue = "" Then _
           Cancel = True: ErrMsg = "Enter Debit/Credit": Exit Sub
        Cancel = (NewValue <> "") And (Not moCn.RecSeek("Select PMCd From Param where " + _
                                      "PTyp= 'DC' and PMCd = '" + NewValue + "'"))
        ErrMsg = "Invalid Debit/Credit Option": Exit Sub
      'EmrSv -
      Case Is = "PRDSEQ"
           Cancel = (Not moCn.RecSeek("Select PMCD From Param where " + _
                                      "PTyp= 'CELL' and PMCD = '" + NewValue + "'"))
        ErrMsg = "Invalid Cell": Exit Sub
      'mw.98-checking valid Rm Sub Category
      Case Is = "CHKTOL"
        Dim ws_RmSCtgArr() As String
        ws_RmSCtgArr = Split(NewValue, ",")
        If IsArray(ws_RmSCtgArr) Then
          For i = 0 To UBound(ws_RmSCtgArr)
            Cancel = Not moCn.RecSeek("Select PSCd From Param where PTyp= 'RMSCTG' and " + _
                                      "PSCd = '" + ws_RmSCtgArr(i) + "'")
            If Cancel = True Then ErrMsg = "RMSCTG (" + CStr(i + 1) + ") - Value " + ws_RmSCtgArr(i) + " Is An Invalid Value": Exit Sub
          Next i
        End If
      End Select
      If ws_BlnkYN = "N" And NewValue = "" Then _
         Cancel = True: ErrMsg = "Value Should Be Entered": Exit Sub
' ###################################  Manoj #### Ver: 2.0.6  ###################################
' Date:   20/01/2004
    Case Is = UCase("vPValidYN")
      Select Case UCase(ADC("wvPTyp"))
      Case Is = "WORK"
        If ws_BlnkYN = "N" And NewValue = "" Then _
           Cancel = True: ErrMsg = "Enter Debit/Credit": Exit Sub
        Cancel = (NewValue <> "") And (Not moCn.RecSeek("Select PMCd From Param where " + _
                                      "PTyp= 'YN' and PMCd = '" + NewValue + "'"))
        ErrMsg = "Worker Code Is Valid Or Not (Y/N)": Exit Sub
      End Select
' ###########################################  Manoj  ###########################################
    'mw.98-checking valid production type, blank allowed
    Case Is = UCase("vPValue3")
      Select Case UCase(ADC("wvPTyp"))
      Case Is = "CHKTOL"
        If ws_BlnkYN = "N" And NewValue = "" Then _
           Cancel = True: ErrMsg = "Production Type Should Be Entered": Exit Sub
        Cancel = (NewValue <> "") And Not ((NewValue = "GS" Or NewValue = "WS") And (("," + GrdvParam.Value(RowNum, "vPValue1") + "," Like "*,C,*") Or ("," + GrdvParam.Value(RowNum, "vPValue1") + "," Like "*,D,*")))
        ErrMsg = "Production Type should be 'GS' Or 'WS' Or Blank AND Rm Category should be D or C": Exit Sub
      End Select
    '************Geeta*************Emr207**********

    Case Is = UCase("vPValue4")
      If ADC("Wvptyp") <> "CHKTOL" Then
        Dim ws_UsrArr() As String
        ws_UsrArr = Split(NewValue, ",")
        If IsArray(ws_UsrArr) Then
          For i = 0 To UBound(ws_UsrArr)
            Cancel = Not moCn.RecSeek("Select PMCd From Param where PTyp= 'USR' and " + _
                                      "PMCd = '" + ws_UsrArr(i) + "'")
            If Cancel = True Then ErrMsg = "User(" + CStr(i + 1) + ") Is Not Defined": Exit Sub
          Next i
        End If
      Else
        Dim ws_RmSCtg() As String

      '************Geeta*************Emr207**********
        ws_RmSCtg = Split(NewValue, ",")
        If IsArray(ws_RmSCtg) Then
          For i = 0 To UBound(ws_RmSCtg)
            Cancel = Not moCn.RecSeek("Select 1 From Param where " + _
                                      "PTyp = 'RMSCTG'  And PSCd = '" + ws_RmSCtg(i) + "'")
            If Cancel = True Then ErrMsg = "Value ' " + ws_RmSCtg(i) + " ' Is An Invalid RmSCtg": Exit Sub
          Next i
        End If
      End If
    Case Is = UCase("wAckYn")
      Select Case UCase(ADC("wvPTyp"))
      Case Is = "CHR"
        Cancel = (Not moCn.RecSeek("Select PMCd From Param where " + _
                                      "PTyp= 'YN' and PMCd = '" + NewValue + "'"))
        ErrMsg = "InValid Acknowledge (Y/N)": Exit Sub
      End Select
      
    End Select
  End With
End Sub
Private Sub ADC_SetRecSource()
  '*** Since this is an 'A' Entry, there is no RecSource associated with ADC. So set the
      'DataMode property of ADC control as xdatUnbound
  ADC.DataMode = xdatUnbound
  'wi_CharWdth = 315
  Call GrdPropSet
End Sub

Private Sub GrdvParam_ReCalc(ByVal RowNum As Integer, ByVal ColName As String, ByVal OldVal As Variant, ByVal SourceName As String, ByVal SourceColName As String)
' *** Repair Trial CHR Header
Dim wRepr As String
Select Case ADC("WVPTYP")
Case Is = UCase("CHR")
With GrdvParam
  Select Case ColName
  'MMAck-if wAckYn value is Y then vpNum value is 1, if wAckYn is N then vpNum is 0 and wiseversa
  Case Is = UCase("vPNum")
    If .Mode = fgmnorm Then
      If .Value(RowNum, "wAckYn") = "N" And .Value(RowNum, "vPNum") <> 0 Then .Value(RowNum, "vPNum") = 0
      If .Value(RowNum, "wAckYn") = "Y" And .Value(RowNum, "vPNum") <> 1 Then .Value(RowNum, "vPNum") = 1
    End If
  Case Is = UCase("wAckYn")
      If .Value(RowNum, "vPNum") = 0 Then .Value(RowNum, "wAckYn") = "N"
      If .Value(RowNum, "vPNum") = 1 Then .Value(RowNum, "wAckYn") = "Y"
  End Select
End With
End Select

End Sub

Private Sub GrdvParam_RowValid(ByVal RowNum As Integer, Cancel As Boolean, ErrMsg As String)
  If ADC("WVPTYP") = "CHKTOL" Then
    If (GrdvParam.Value(RowNum, "vPDesc") = "") Then
      Cancel = True
      ErrMsg = "BCHR list cannot be empty"
    End If
    If (GrdvParam.Value(RowNum, "vPValue1") = "") Then
      Cancel = True
      ErrMsg = "RmCtg list cannot be empty"
    End If
  End If
End Sub


Private Sub GrdvParam_RowWhen(ByVal RowNum As Integer)
    If GrdvParam.ColProp("vPValidYn").Style = fgcNormal And GrdvParam.IsNew(RowNum) = True Then
      GrdvParam.Value(RowNum, "vPValidYn") = "Y"
    End If
    If GrdvParam.Value(RowNum, "vpMCD") = "DT" Or GrdvParam.Value(RowNum, "vpMCD") = "BV" Then
      GrdvParam.ColProp("vPValue3").Heading = "RepairChr"
    Else
      If "Y" = moCn.GetFldVal("Select 'Y' from Param where Ptyp= 'TC' and PMCd='" + GrdvParam.Value(RowNum, "vPMCD") + "' and PValue='JT' and PValue1='N'") Then
        GrdvParam.ColProp("vPValue3").Heading = "MRP Chr"
      Else
        'mw.98-if vptyp=CHKTOL then heading will be PrdTyp else Trial chr
        If ADC("wvPTyp") <> "CHKTOL" Then GrdvParam.ColProp("vPValue3").Heading = "Trial Chr"
      End If
    End If
End Sub

Private Sub GrdvParam_SetRecSource()
  '*** Set the Record Source of the Grid GrdvParam
      'Sort the Recs in the grid By PMCd, PSCd
  Dim ws_Cnd As String, ws_OrdBy As String
  ws_Cnd = IIF(ADC("wvPTyp") = "PTYP", " where vPCoCd= '" + ctSelfCoCd + "' ", " where vPCoCd= '" + ADC("wvPCoCd") + "' ")
  ws_Cnd = ws_Cnd + IIF(ADC.RepCond <> "", " and " + ADC.RepCond, "")
  ws_OrdBy = " Order By vPMCd, vPSCd "
  GrdvParam.RecSource = " Select * from vParam " + ws_Cnd + ws_OrdBy
  
End Sub
Private Sub GrdvParam_InitKey(ByVal RowNum As Integer)
  '*** Initialise Key Field PTyp
  With GrdvParam
    If .IsNew(RowNum) Then
      'If ADC("wvPTyp") = "PTYP" Then
      '  .Store "vPCoCd", ctSelfCoCd
      'Else
      .Store "vPCoCd", ADC("wvPCoCd")
      'End If
      .Store "vPTyp", ADC("wvPTyp")
      .Store "vPPrtKey", ctCurrPrtn   '****** Sachin 3.02.0 - Setting Default Partition Key
    End If
    


    '****** Sachin 2.14.0 - 01-11-2006 [Multi-Company] ******
   'InUse Check- if Ptyp=WIWSEQ/WIWLONG/RWIWSEQ/RWIWLONG then it store PrntYn value in database
    If ADC("wvPTyp") = "WIWSEQ" Or ADC("wvPTyp") = "WIWLONG" Or ADC("wvPTyp") = "RWIWSEQ" Or ADC("wvPTyp") = "RWIWLONG" Then _
      .Store "vPValue3", ADC("wvPValue3")
    '****** Sachin 2.14.0 - 01-11-2006 [Multi-Company] ******
  End With
End Sub

Private Sub GrdPropSet()
  '*** Write a procedure to set the properties like Style, Heading, DataType, Mask,
      'MaxLength of the Grid GrdvParam using the TcStr table
      'Note that the heading has to be appended by spaces when necessary if the
      'column width is not accomodating the length of the field

  Dim wo_RsTc As MwfLib.MDORowSet
  Dim wi_SpaceLen As Single
  Set wo_RsTc = moCn.OpenRes("Select * from TcStr where TsTbl='vParam' and " + _
                            "TsTyp='" + ADC("wvPTyp") + "'")
  With GrdvParam
    Do While Not (wo_RsTc.EOF Or wo_RsTc.BOF)
      .ColProp(wo_RsTc!TsIdName).Style = ColSty(wo_RsTc!TsStyle)
      '***** To Find Out The Space To Be Appended To The Heading
            '(This Is To Set The Proper Width For The Field)
      '***** 2.5 is taken as the approximate value which stands for the space occupied by one
            'character in Upper Case
            If wo_RsTc!TsMaxLen >= 225 Then
              wi_SpaceLen = (80 * 2.5) - (Len(wo_RsTc!TsLbl))
      
            Else
                wi_SpaceLen = (wo_RsTc!TsMaxLen * 2.5) - (Len(wo_RsTc!TsLbl))
      
            End If
      .ColProp(wo_RsTc!TsIdName).Heading = wo_RsTc!TsLbl + IIF(wi_SpaceLen > 0, Space(Abs(wi_SpaceLen)), "")
      
      .ColProp(wo_RsTc!TsIdName).DataType = DataTyp(wo_RsTc!TsDataTyp)
      .ColProp(wo_RsTc!TsIdName).Mask = wo_RsTc!TsMask
      .ColProp(wo_RsTc!TsIdName).MaxLength = wo_RsTc!TsMaxLen
      
      '***** Note: Width Property cannot be used to set the width
      '*****  .ColProp(wo_RsTc!TsIdName).Width = wi_CharWdth * wo_RsTc!TsMaxLen
      wo_RsTc.MoveNext
    Loop
  
  End With
    
  GrdvParam.SetColWidthToHd
  Set wo_RsTc = Nothing
End Sub
Private Function ColSty(ByVal mColSty As String) As en_fgColStyle
  '*** Write a function to return the Grid Column Style based on the value of the
      'TsStyle field of the TcStr table
  Select Case mColSty
    Case Is = "N"
      ColSty = fgcNormal
    Case Is = "I"
      ColSty = fgcinvisible
    Case Is = "R"
      ColSty = fgcReadOnly
  End Select
End Function
Private Function DataTyp(ByVal mDataTyp As String) As en_FldDataType
  '*** Write a function to return the datatype of a column based on the value of the
      'TsDataTyp field of the TcStr table
  Select Case mDataTyp
    Case Is = "C"
      DataTyp = fdtCharType
    Case Is = "N"
      DataTyp = fdtNumber
    Case Is = "F"
      DataTyp = fdtFloat
    Case Is = "D"
      DataTyp = fdtDateType
    Case Is = "U"
      DataTyp = fdtConvertToUpper
    Case Is = "T"
      DataTyp = fdtTimeType
  End Select
End Function

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
Private Sub GrdvParam_Validate(Cancel As Boolean)
  Cancel = GrdvParam.Validate
End Sub
