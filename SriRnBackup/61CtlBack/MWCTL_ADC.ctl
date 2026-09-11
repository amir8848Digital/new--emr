VERSION 5.00
Begin VB.UserControl MWCTL_ADC 
   BackStyle       =   0  'Transparent
   ClientHeight    =   570
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   10260
   LockControls    =   -1  'True
   ScaleHeight     =   570
   ScaleWidth      =   10260
   Begin VB.PictureBox PCT_SCD 
      Align           =   4  'Align Right
      BorderStyle     =   0  'None
      Height          =   570
      Left            =   5445
      ScaleHeight     =   570
      ScaleWidth      =   4815
      TabIndex        =   4
      TabStop         =   0   'False
      Top             =   0
      Width           =   4815
      Begin MwfCtl.MWCTL_BTN1 arr_cmd_svcndl 
         Height          =   435
         Index           =   0
         Left            =   2130
         TabIndex        =   5
         Top             =   0
         Width           =   885
         _extentx        =   1561
         _extenty        =   767
         disablecolor    =   8421504
         forecolor       =   8421504
         enabled         =   0
         backcolor       =   14737632
         maskcolor       =   8454143
         caption         =   "<S>ave"
         lblfont         =   "MWCTL_ADC.ctx":0000
      End
      Begin MwfCtl.MWCTL_BTN1 arr_cmd_svcndl 
         Height          =   435
         Index           =   1
         Left            =   3015
         TabIndex        =   6
         Top             =   0
         Visible         =   0   'False
         Width           =   885
         _extentx        =   1561
         _extenty        =   767
         disablecolor    =   8421504
         forecolor       =   8421504
         enabled         =   0
         backcolor       =   14737632
         maskcolor       =   8454143
         caption         =   "<C>lose"
         lblfont         =   "MWCTL_ADC.ctx":002E
      End
      Begin MwfCtl.MWCTL_BTN1 arr_cmd_svcndl 
         Height          =   435
         Index           =   2
         Left            =   3900
         TabIndex        =   7
         Top             =   0
         Width           =   885
         _extentx        =   1561
         _extenty        =   767
         disablecolor    =   8421504
         forecolor       =   8421504
         enabled         =   0
         backcolor       =   14737632
         maskcolor       =   8454143
         caption         =   "<D>elete"
         lblfont         =   "MWCTL_ADC.ctx":005C
      End
   End
   Begin VB.PictureBox PCT_AF 
      Align           =   3  'Align Left
      BorderStyle     =   0  'None
      Height          =   570
      Left            =   0
      ScaleHeight     =   570
      ScaleWidth      =   4815
      TabIndex        =   3
      TabStop         =   0   'False
      Top             =   0
      Width           =   4815
      Begin MwfCtl.MWCTL_BTN1 arr_cmd_adfnd 
         Height          =   435
         Index           =   0
         Left            =   0
         TabIndex        =   0
         Top             =   0
         Width           =   885
         _extentx        =   1561
         _extenty        =   767
         disablecolor    =   8421504
         backcolor       =   14737632
         maskcolor       =   8454143
         caption         =   "<A>dd"
         lblfont         =   "MWCTL_ADC.ctx":008A
      End
      Begin MwfCtl.MWCTL_BTN1 arr_cmd_adfnd 
         Height          =   435
         Index           =   1
         Left            =   885
         TabIndex        =   1
         Top             =   0
         Width           =   885
         _extentx        =   1561
         _extenty        =   767
         disablecolor    =   8421504
         backcolor       =   14737632
         maskcolor       =   8454143
         caption         =   "<F>ind"
         lblfont         =   "MWCTL_ADC.ctx":00B8
      End
      Begin MwfCtl.MWCTL_BTN1 cmd_exit 
         Height          =   435
         Left            =   1770
         TabIndex        =   2
         Top             =   0
         Width           =   885
         _extentx        =   1561
         _extenty        =   767
         disablecolor    =   8421504
         backcolor       =   14737632
         maskcolor       =   8454143
         caption         =   "e<X>it"
         lblfont         =   "MWCTL_ADC.ctx":00E6
      End
   End
End
Attribute VB_Name = "MWCTL_ADC"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
Dim me_DataMode As en_XDataMode
Dim me_EntType As en_XEntType
Dim ms_MenuCd As String
Dim ms_UsrCd As String
'*** (Jen 04/02/08)
Dim mo_Connection As MDOConnection     '*** (Bef speed) bef 24/11/06
'*** (Jen 04/02/08)
Dim mo_DataMode As en_XDataMode
Dim mf_ExcelReportMode As Boolean
Dim mo_FirKeyCtl As Object
Dim mo_FirNKeyCtl As Object
Dim ms_RecSource As String
Dim mo_RepSource As CRAXDRT.Report
Dim ms_RepRecSource As String
Dim ms_RepExportFileDirectory As String
Dim mf_AllowAdd As Boolean
Dim mf_AllowFind As Boolean
Dim mf_AllowSave As Boolean
Dim mf_AllowCancel As Boolean
Dim mf_AllowDelete As Boolean
Dim mf_AllowAddRt As Boolean
Dim mf_AllowFindRt As Boolean
Dim mf_AllowSaveRt As Boolean
Dim mf_AllowCancelRt As Boolean
Dim mf_AllowDeleteRt As Boolean
Dim mf_Automate As Boolean
Dim me_DeleteMode As en_XDeleteMode
Dim mf_ExecGot As Boolean
Dim mf_Hkey As Boolean
Dim mo_Rowset As MDORowSet
Dim ms_ValFailed As String
Dim ms_BtnValFailed As String         '*** Jenny
Dim mc_CtrlColln As Collection
Dim mc_FlxColln As Collection
Dim mc_AddSave As Collection
Private Const ctm_Read = "R"
Private Const ctm_Write = "W"
Private Const ctm_Init = "I"
Private Const ctm_InitAll = "A"
Public Event AutoGen()
Public Event SetRecSource()
Public Event SetRepRecSource()
Public Event NonKeyWhen(ByRef pr_Cancel As Boolean, ByRef pr_ErrMsg As String)
Public Event NonKeyValid(ByRef pr_Cancel As Boolean, ByRef pr_ErrMsg As String)
Public Event KeyValid(ByRef pr_Cancel As Boolean, ByRef pr_ErrMsg As String)
Public Event KeyWhen()
Public Event Load()
Const ct_SqlErrMsg = "[Microsoft][ODBC SQL Server Driver][SQL Server]"
Dim ms_child As String
Dim me_AddFndMode As en_XAddFndMode
Dim me_Mode As en_XMode
Dim mf_Opt As String
Dim mf_AddSave As Boolean
Dim md_BookMark As Double
Dim mf_DataChanged As Boolean
Dim mf_ChkData As Boolean
' ########################  Manoj #### Ver: 2.0.6 #### Date: 03/06/2004  ########################
Dim mo_MnuFrm As Form
Dim ms_HostNm As String
' ###########################################  Manoj  ###########################################

'*** Jenny Color (22/03/05)
Dim ms_ToolBarCtrl As String
Dim mi_ToolBarCtrlIndx As String
'*** Jenny Color (22/03/05)

Dim ms_CompanyCd As String  '*** (Jenny 11/08/07)
Dim ml_UlIdNo As Long       '****** Sachin 3.5.0 [11/11/08]

'***************** Props Begin ************************************************************'
                  ' xnorm, xinit, xclear, xcopy.
Public Property Get Mode() As en_XMode
    Mode = me_Mode
End Property
                  ' Bound, Unbound.
Public Property Get DataMode() As en_XDataMode
Attribute DataMode.VB_MemberFlags = "400"
    DataMode = me_DataMode
End Property
Public Property Let DataMode(ByVal vNewValue As en_XDataMode)
    me_DataMode = vNewValue
End Property
Public Property Get RecSource() As String
Attribute RecSource.VB_MemberFlags = "400"
    RecSource = ms_RecSource
End Property
Public Property Let RecSource(ByVal vNewValue As String)
    ms_RecSource = vNewValue
End Property
                  
' adc.repsource , sets the report object (dsr)
' reprecource is the string for the report.
Public Property Set RepSource(ByVal vNewValue As Object)
    Set mo_RepSource = vNewValue
End Property
Public Property Get RepSource() As CRAXDRT.Report
    Set RepSource = mo_RepSource
End Property
Public Property Set FirKeyCtl(ByVal vNewValue As Object)
Attribute FirKeyCtl.VB_MemberFlags = "400"
    Set mo_FirKeyCtl = vNewValue
End Property
Public Property Set FirNKeyCtl(ByVal vNewValue As Object)
Attribute FirNKeyCtl.VB_MemberFlags = "400"
    Set mo_FirNKeyCtl = vNewValue
End Property
                    
'*** Jenny
Public Property Let BckColor(ByVal vNewValue As OLE_COLOR)
    '*** This property is not actually required
    UserControl.BackColor = vNewValue
    '*** 21/05/05
    PCT_AF.BackColor = vNewValue
    PCT_SCD.BackColor = vNewValue
    '*** 21/05/05
    PropertyChanged "BckColor"
End Property
Public Property Get BckColor() As OLE_COLOR
    '*** This property is not actually required
    BckColor = UserControl.BackColor
End Property
Public Property Let BckColorButton(ByVal vNewValue As OLE_COLOR)
    '*** This property is not actually required
    arr_cmd_adfnd(0).BackColor = vNewValue
    arr_cmd_adfnd(1).BackColor = vNewValue
    cmd_exit.BackColor = vNewValue
    arr_cmd_svcndl(0).BackColor = vNewValue
    arr_cmd_svcndl(1).BackColor = vNewValue
    arr_cmd_svcndl(2).BackColor = vNewValue
    
    PropertyChanged "BckColorButton"
End Property
Public Property Get BckColorButton() As OLE_COLOR
    '*** This property is not actually required
    BckColorButton = arr_cmd_adfnd(0).BackColor
End Property
'Public Property Get BtnPicture() As Variant
'    BtnPicture = UserControl.Picture
'End Property
'Private Property Let BtnPicture(ByVal vNewValue As Variant)
'    If CStr(vNewValue) <> "" And Dir(CStr(vNewValue)) <> "" Then
'      UserControl.Picture = LoadPicture(CStr(vNewValue))
'    Else
'      UserControl.Picture = vNewValue
'      arr_cmd_adfnd(0).Pic = vNewValue
'      arr_cmd_adfnd(1).Pic = vNewValue
'      cmd_exit.Pic = vNewValue
'      arr_cmd_svcndl(0).Pic = vNewValue
'      arr_cmd_svcndl(1).Pic = vNewValue
'      arr_cmd_svcndl(2).Pic = vNewValue
'    End If
'    PropertyChanged "BtnPicture"
'End Property

Public Property Get BtnPicture() As Picture
Attribute BtnPicture.VB_ProcData.VB_Invoke_Property = ";Additional"
    Set BtnPicture = UserControl.Picture
End Property
Private Property Let BtnPicture(ByVal vNewValue As Picture)
    Set UserControl.Picture = vNewValue
    arr_cmd_adfnd(0).Pic = vNewValue
    arr_cmd_adfnd(1).Pic = vNewValue
    cmd_exit.Pic = vNewValue
    arr_cmd_svcndl(0).Pic = vNewValue
    arr_cmd_svcndl(1).Pic = vNewValue
    arr_cmd_svcndl(2).Pic = vNewValue

    PropertyChanged "BtnPicture"
End Property
Public Property Set BtnPicture(ByVal vNewValue As Picture)
    Set UserControl.Picture = vNewValue
    arr_cmd_adfnd(0).Pic = vNewValue
    arr_cmd_adfnd(1).Pic = vNewValue
    cmd_exit.Pic = vNewValue
    arr_cmd_svcndl(0).Pic = vNewValue
    arr_cmd_svcndl(1).Pic = vNewValue
    arr_cmd_svcndl(2).Pic = vNewValue
    PropertyChanged "BtnPicture"
End Property
'Public Property Get PicTxt() As String
'    PicTxt = ms_PicTxt
'End Property
'Public Property Let PicTxt(ByVal vNewValue As String)
'    'If GetUserMode = False Then
'    'End If
'    ms_PicTxt = vNewValue
'    BtnPicture = vNewValue
'
'    arr_cmd_adfnd(0).PicTxt = vNewValue
'    arr_cmd_adfnd(1).PicTxt = vNewValue
'    cmd_exit.PicTxt = vNewValue
'    arr_cmd_svcndl(0).PicTxt = vNewValue
'    arr_cmd_svcndl(1).PicTxt = vNewValue
'    arr_cmd_svcndl(2).PicTxt = vNewValue
'    PropertyChanged "PicTxt"
'End Property
Public Property Let BtnWidth(ByVal vNewValue As Integer)
    arr_cmd_adfnd(0).Width = vNewValue
    arr_cmd_adfnd(1).Width = vNewValue
    cmd_exit.Width = vNewValue
    arr_cmd_svcndl(0).Width = vNewValue
    arr_cmd_svcndl(1).Width = vNewValue
    arr_cmd_svcndl(2).Width = vNewValue
    
    Call SetAdFndPos
    Call SetSvCnDlPos
    
    PropertyChanged "BtnWidth"
End Property
Public Property Get BtnWidth() As Integer
Attribute BtnWidth.VB_ProcData.VB_Invoke_Property = ";Additional"
    BtnWidth = arr_cmd_adfnd(0).Width
End Property
Public Property Let BtnHeight(ByVal vNewValue As Integer)
    arr_cmd_adfnd(0).Height = vNewValue
    arr_cmd_adfnd(1).Height = vNewValue
    cmd_exit.Height = vNewValue
    arr_cmd_svcndl(0).Height = vNewValue
    arr_cmd_svcndl(1).Height = vNewValue
    arr_cmd_svcndl(2).Height = vNewValue
    
    Call SetAdFndPos
    Call SetSvCnDlPos
    
    PropertyChanged "BtnHeight"
End Property
Public Property Get BtnHeight() As Integer
Attribute BtnHeight.VB_ProcData.VB_Invoke_Property = ";Additional"
    BtnHeight = arr_cmd_adfnd(0).Height
End Property
Public Property Let BtnMaskColor(ByVal vNewValue As OLE_COLOR)
    arr_cmd_adfnd(0).MaskColor = vNewValue
    arr_cmd_adfnd(1).MaskColor = vNewValue
    cmd_exit.MaskColor = vNewValue
    arr_cmd_svcndl(0).MaskColor = vNewValue
    arr_cmd_svcndl(1).MaskColor = vNewValue
    arr_cmd_svcndl(2).MaskColor = vNewValue
    
    PropertyChanged "BtnMaskColor"
End Property
Public Property Get BtnMaskColor() As OLE_COLOR
Attribute BtnMaskColor.VB_ProcData.VB_Invoke_Property = ";Additional"
    BtnMaskColor = arr_cmd_adfnd(0).MaskColor
End Property
Public Property Let BtnForeColor(ByVal vNewValue As OLE_COLOR)
    arr_cmd_adfnd(0).ForeColor = vNewValue
    arr_cmd_adfnd(1).ForeColor = vNewValue
    cmd_exit.ForeColor = vNewValue
    arr_cmd_svcndl(0).ForeColor = vNewValue
    arr_cmd_svcndl(1).ForeColor = vNewValue
    arr_cmd_svcndl(2).ForeColor = vNewValue
    
    PropertyChanged "BtnForeColor"
End Property
Public Property Get BtnForeColor() As OLE_COLOR
Attribute BtnForeColor.VB_ProcData.VB_Invoke_Property = ";Additional"
    BtnForeColor = arr_cmd_adfnd(0).ForeColor
End Property
Public Property Let BtnDisableColor(ByVal vNewValue As OLE_COLOR)
    arr_cmd_adfnd(0).DisableColor = vNewValue
    arr_cmd_adfnd(1).DisableColor = vNewValue
    cmd_exit.DisableColor = vNewValue
    arr_cmd_svcndl(0).DisableColor = vNewValue
    arr_cmd_svcndl(1).DisableColor = vNewValue
    arr_cmd_svcndl(2).DisableColor = vNewValue

    PropertyChanged "BtnDisableColor"
End Property
Public Property Get BtnDisableColor() As OLE_COLOR
Attribute BtnDisableColor.VB_ProcData.VB_Invoke_Property = ";Additional"
    BtnDisableColor = arr_cmd_adfnd(0).DisableColor
End Property
Public Property Let FontSz(ByVal vNewValue As Integer)
    arr_cmd_adfnd(0).LblFont.Size = vNewValue
    arr_cmd_adfnd(1).LblFont.Size = vNewValue
    cmd_exit.LblFont.Size = vNewValue
    arr_cmd_svcndl(0).LblFont.Size = vNewValue
    arr_cmd_svcndl(1).LblFont.Size = vNewValue
    arr_cmd_svcndl(2).LblFont.Size = vNewValue
    
    PropertyChanged "FontSz"
End Property
Public Property Get FontSz() As Integer
Attribute FontSz.VB_ProcData.VB_Invoke_Property = ";Additional"
    FontSz = arr_cmd_adfnd(0).LblFont.Size
End Property
'*** 07/07/05
Public Property Get ReSizeBtnPic() As Boolean
Attribute ReSizeBtnPic.VB_ProcData.VB_Invoke_Property = ";Additional"
    ReSizeBtnPic = arr_cmd_adfnd(0).ReSizePic
End Property
Public Property Let ReSizeBtnPic(ByVal vNewValue As Boolean)
    arr_cmd_adfnd(0).ReSizePic = vNewValue
    arr_cmd_adfnd(1).ReSizePic = vNewValue
    cmd_exit.ReSizePic = vNewValue
    arr_cmd_svcndl(0).ReSizePic = vNewValue
    arr_cmd_svcndl(1).ReSizePic = vNewValue
    arr_cmd_svcndl(2).ReSizePic = vNewValue
    
    PropertyChanged "ReSizeBtnPic"
End Property
'*** 07/07/05
Public Property Get Pic() As Picture
Attribute Pic.VB_ProcData.VB_Invoke_Property = ";Additional"
    Set Pic = PCT_AF.Picture
End Property
Private Property Let Pic(ByVal vNewValue As Picture)
    Set PCT_AF.Picture = vNewValue
    Set PCT_SCD.Picture = vNewValue

    PropertyChanged "Pic"
End Property
Public Property Set Pic(ByVal vNewValue As Picture)
    Set PCT_AF.Picture = vNewValue
    Set PCT_SCD.Picture = vNewValue
    
    PropertyChanged "Pic"
End Property
'*** Jenny

'*** Jenny Color (22/03/05)
Public Property Let ToolBarCtrlName(ByVal vNewValue As String)
    ms_ToolBarCtrl = UCase(vNewValue)
End Property
Public Property Get ToolBarCtrlName() As String
    ToolBarCtrlName = ms_ToolBarCtrl
End Property
Public Property Let ToolBarCtrlIndx(ByVal vNewValue As Integer)
    mi_ToolBarCtrlIndx = UCase(vNewValue)
End Property
Public Property Get ToolBarCtrlIndx() As Integer
    ToolBarCtrlIndx = mi_ToolBarCtrlIndx
End Property
'*** Jenny Color (22/03/05)
                    
                    
                    
                    ' adc.child
Public Property Get Child() As String
    Child = ms_child
End Property
Public Property Let Child(ByVal vNewValue As String)
    ms_child = vNewValue
End Property
                    
                    ' Valfailed = '' if no errors, = Med.idname if validation failed.
                    ' lost focus exists it Valfailed <> ''
                    ' Got focus exists if Valfailed <> '' and <> its name
Public Property Get ValFailed() As String
    ValFailed = ms_ValFailed
End Property
                    ' current active control focus is Add / Find / Save / Cancel /Delete
Public Property Get OptSel() As String
    OptSel = ""
    If UserControl.ActiveControl Is Nothing Then Exit Property
    If UserControl.ActiveControl.Name = "arr_cmd_adfnd" Then
        Select Case UserControl.ActiveControl.Index
        Case Is = 0
            OptSel = "A"
        Case Is = 1
            OptSel = "F"
        End Select
    End If
    If UserControl.ActiveControl.Name = "arr_cmd_svcndl" Then
        Select Case UserControl.ActiveControl.Index
        Case Is = 0
            OptSel = "S"
        Case Is = 1
            OptSel = "C"
        Case Is = 2
            OptSel = "D"
        End Select
    End If
    If UserControl.ActiveControl.Name = "cmd_exit" Then
        OptSel = "X"
    End If
End Property
Public Property Let ValFailed(ByVal vNewValue As String)
    ms_ValFailed = vNewValue
    If vNewValue <> "" Then
     mf_Opt = ""
    End If
End Property
                        
'*** Jenny
Public Property Get BtnValFailed() As String
    BtnValFailed = ms_BtnValFailed
End Property
Public Property Let BtnValFailed(ByVal vNewValue As String)
    ms_BtnValFailed = vNewValue
End Property
'*** Jenny
                        
                        
                        ' AllowAdd returns true if both AllowAdd Design time & Run time are true
Public Property Get AllowAdd() As Boolean
    AllowAdd = (mf_AllowAdd And mf_AllowAddRt)
End Property

                        ' at run time Calls SetAddFndDis, Disables
                        ' at Design time Calls SetAdFndPos, Invisble
Public Property Let AllowAdd(ByVal vNewValue As Boolean)
    If GetUserMode Then
        mf_AllowAddRt = vNewValue
        If (Not mf_AllowAdd) And mf_AllowAddRt Then
            MsgBox "Cannot permit Add on entry with no Add option"
            mf_AllowAddRt = False
        End If
        Call SetAdFndEnDis
    Else
        mf_AllowAdd = vNewValue
        Call SetAdFndPos
    End If
End Property
Public Property Get AllowFind() As Boolean
    AllowFind = (mf_AllowFind And mf_AllowFindRt)
End Property
Public Property Let AllowFind(ByVal vNewValue As Boolean)
    If GetUserMode Then
        mf_AllowFindRt = vNewValue
        If (Not mf_AllowFind) And mf_AllowFindRt Then
            MsgBox "Cannot permit Find on entry with no Find option"
            mf_AllowFindRt = False
        End If
        Call SetAdFndEnDis
    Else
        mf_AllowFind = vNewValue
        Call SetAdFndPos
    End If
End Property
Public Property Get AllowSave() As Boolean
    AllowSave = (mf_AllowSave And mf_AllowSaveRt)
End Property
Public Property Let AllowSave(ByVal vNewValue As Boolean)
    If GetUserMode Then
        mf_AllowSaveRt = vNewValue
        If (Not mf_AllowSave) And mf_AllowSaveRt Then
            MsgBox "Cannot permit Save on entry with no Save option"
            mf_AllowSaveRt = False
        End If
        Call SetSvCnDlEnDis
    Else
        mf_AllowSave = vNewValue
        Call SetSvCnDlPos
    End If
End Property
Public Property Get AllowCancel() As Boolean
    AllowCancel = (mf_AllowCancel And mf_AllowCancelRt)
End Property
Public Property Let AllowCancel(ByVal vNewValue As Boolean)
    If GetUserMode Then
        mf_AllowCancelRt = vNewValue
        If (Not mf_AllowCancel) And mf_AllowCancelRt Then
            MsgBox "Cannot permit Cancel on entry with no Cancel option"
            mf_AllowCancelRt = False
        End If
        Call SetSvCnDlEnDis
    Else
        mf_AllowCancel = vNewValue
        Call SetSvCnDlPos
    End If
End Property
Public Property Get AllowDelete() As Boolean
    AllowDelete = (mf_AllowDelete And mf_AllowDeleteRt)
End Property
Public Property Let AllowDelete(ByVal vNewValue As Boolean)
    If GetUserMode Then
        mf_AllowDeleteRt = vNewValue
        If (Not mf_AllowDelete) And mf_AllowDeleteRt Then
            MsgBox "Cannot permit Delete on entry with no Delete option"
            mf_AllowDeleteRt = False
        End If
        Call SetSvCnDlEnDis
    Else
        mf_AllowDelete = vNewValue
        Call SetSvCnDlPos
    End If
End Property
Public Property Get ExcelReportMode() As Boolean
    ExcelReportMode = mf_ExcelReportMode
End Property
Public Property Let ExcelReportMode(ByVal vNewValue As Boolean)
  mf_ExcelReportMode = vNewValue
End Property

Public Property Get GetUserMode() As Boolean
    If mf_ExcelReportMode Then
      GetUserMode = False
      Exit Property
    End If
    GetUserMode = Ambient.UserMode

End Property

                        ' default property of Adc, contains the Med Controls
                        ' ??? Get as MWCTL_MED
Public Property Get Ctl(ByVal IdName As String) As Object
Attribute Ctl.VB_UserMemId = 0
    If GetUserMode Then
        Set Ctl = mc_CtrlColln(IdName)
    Else
        Set Ctl = FindObj(IdName)
    End If
End Property
                        ' Names of all the MwCtl_Flx in the form
Public Property Get FlxColln() As Collection
    Set FlxColln = mc_FlxColln
End Property
                          ' returns true if either a/f/x/s/c/d enabled
Public Property Get Enabled() As Boolean
    Enabled = arr_cmd_adfnd(0).Enabled Or arr_cmd_adfnd(1).Enabled Or cmd_exit.Enabled Or arr_cmd_svcndl(0).Enabled Or arr_cmd_svcndl(1).Enabled Or arr_cmd_svcndl(2).Enabled
End Property
                           ' where no add/find button is shown (mainly report)
Public Property Get Automate() As Boolean
    Automate = mf_Automate
End Property
Public Property Let Automate(ByVal vNewValue As Boolean)
    mf_Automate = vNewValue
    If mf_Automate Then
        PCT_AF.Align = 4
        PCT_SCD.Align = 3
        Call SetAdFndPos
    Else
        PCT_AF.Align = 3
        PCT_SCD.Align = 4
        Call SetAdFndPos
    End If
    PropertyChanged "Automate"
End Property
                            ' not used
Public Property Get ExecGot() As Boolean
    ExecGot = mf_ExecGot
End Property
Public Property Let ExecGot(ByVal vNewValue As Boolean)
    mf_ExecGot = vNewValue
End Property
                            ' delchildrecs, delnormal
Public Property Get DeleteMode() As en_XDeleteMode
    DeleteMode = me_DeleteMode
End Property
Public Property Let DeleteMode(ByVal vNewValue As en_XDeleteMode)
    me_DeleteMode = vNewValue
    PropertyChanged "DeleteMode"
End Property
Public Property Get Connection() As MDOConnection
    '*** (Jen 04/02/08)
    Set Connection = mo_Connection   '*** (Bef 24/11/06)
    '*** (Jen 04/02/08)
    
    '*** (Bef 04/02/08)
    ''*** (Jenny speed 24/11/06)
    'Set Connection = MWLib.GetConn
    ''*** (Jenny speed 24/11/06)
    '*** (Bef 04/02/08)
End Property
Public Property Let MenuCd(ByVal vNewValue As String)
    ms_MenuCd = UCase(vNewValue)
End Property
Public Property Get MenuCd() As String
    MenuCd = ms_MenuCd
End Property
Public Property Let UsrCd(ByVal vNewValue As String)
    ms_UsrCd = UCase(vNewValue)
End Property
Public Property Get UsrCd() As String
    UsrCd = ms_UsrCd
End Property
                          ' recsource for the report
Public Property Let RepRecSource(ByVal vNewValue As String)
    ms_RepRecSource = vNewValue
End Property
    ' recsource for the report
Public Property Let RepExportFileDirectory(ByVal vNewValue As String)
    ms_RepExportFileDirectory = vNewValue
End Property

Public Property Get RepExportFileDirectory() As String
    RepExportFileDirectory = ms_RepExportFileDirectory
End Property


' ########################  Manoj #################### Date: 30/04/2004  ########################
' RecRepSource Get Property Opened for Refreshing Report While Printing
Public Property Get RepRecSource() As String
    RepRecSource = ms_RepRecSource
End Property
' ###########################################  Manoj  ###########################################
                          
                          ' entry , report
Public Property Get EntType() As en_XEntType
    EntType = me_EntType
End Property
Public Property Let EntType(ByVal vNewValue As en_XEntType)
    me_EntType = vNewValue
    If me_EntType = xetReport Then
        arr_cmd_svcndl(0).Caption = "<S>how"
        arr_cmd_svcndl(1).Caption = "<C>lose"
    End If
    PropertyChanged "EntType"
End Property
                          ' Add, Find
Public Property Get AddFndMode() As en_XAddFndMode
    AddFndMode = me_AddFndMode
End Property
                          ' Report Condition based on CmpStr, 'SEL' as last 3 Charac
Public Property Get RepCond() As String
    Dim wo_ReqCtl As MWCTL_MED, ws_rcond As String, ws_cond As String, wo_Ctl As Control
    
    '*** (Bef 23/02/06) For Each wo_Ctl In Parent.Controls
    '*** (Jenny 23/02/06)
    For Each wo_Ctl In Parent.ATXT
    '*** (Jenny 23/02/06)
        '*** (Bef 23/02/06) If Not (TypeOf wo_Ctl Is MWCTL_MED) Then GoTo NextCtl
        Set wo_ReqCtl = wo_Ctl
       With wo_ReqCtl
            ws_rcond = ""
            If .CmpStr <> "" And (Not .IsBlank) Then
              If Not (.DataType = fdtDateType And CStr(.Value) = "01/01/80") Then
                ' Date check for 01/01/80 added by rn, 14/8/01
                ws_rcond = .CmpStr
                If UCase(Right(.IdName, 3)) = "SEL" Then
                    ws_rcond = ws_rcond + "(" + .Value + ")"
                Else
                    Select Case .DataType
                    Case Is = fdtCharType, fdtConvertToUpper
                        ws_rcond = ws_rcond + "'" + .Value + "'"
                    Case Is = fdtDateType
                        ws_rcond = ws_rcond + "'" + CStr(.Value) + "'"
                    Case Is = fdtNumber, fdtFloat, fdtTimeType
                        ws_rcond = ws_rcond + CStr(.Value)
                    End Select
                End If
                ws_cond = ws_cond + IIf(ws_cond <> "" And ws_rcond <> "", " and ", "") + ws_rcond
              End If
            End If
        End With
NextCtl:
    Next wo_Ctl
    RepCond = ws_cond
End Property
Public Property Get IsNew() As Boolean
    IsNew = IIf(md_BookMark = 0, True, False)
End Property
Public Property Get IsDel() As Boolean
    IsDel = IIf(md_BookMark = -1, True, False)
End Property
                  
' ########################  Manoj #################### Date: 03/06/2004  ########################
                  ' adc.MnuForm , sets the Menu Form
Public Property Set MnuForm(ByVal vNewValue As Object)
    Set mo_MnuFrm = vNewValue
End Property
' ###########################################  Manoj  ###########################################
' ########################  Manoj #################### Date: 16/06/2004  ########################
  ' adc.HostNm, sets the Host Machine Name
Public Property Get HostNm() As String
  HostNm = ms_HostNm
End Property
  ' adc.HostNm, Returns the Host Machine Name
Public Property Let HostNm(ByVal vNewValue As String)
  ms_HostNm = vNewValue
End Property
' ###########################################  Manoj  ###########################################

'*** (Jenny 11/08/07)
Public Property Get CompanyCd() As String
  CompanyCd = ms_CompanyCd
End Property
  ' adc.CompanyCd, Returns the Login Company Code
Public Property Let CompanyCd(ByVal vNewValue As String)
  ms_CompanyCd = vNewValue
End Property
'*** (Jenny 11/08/07)

'****** Sachin 3.5.0 [11/11/08] ******
Public Property Get UlIdNo() As Long
    UlIdNo = ml_UlIdNo
End Property
'ADC.UlIdNo returns the IdNo for respective Login assigned in UsrLogin Table
Public Property Let UlIdNo(ByVal vNewValue As Long)
    ml_UlIdNo = vNewValue
End Property
'****** Sachin 3.5.0 [11/11/08] ******

'***************** Props End ************************************************************'
'***************** Events Begin ************************************************************'
Private Sub arr_cmd_adfnd_Click(Index As Integer)
    If MWLib.GetDateFormat <> MWLib.DtSetting Then
'      Parent.adc.AllowSave = False
 '     Parent.adc.AllowClose = True
       MsgBox "Regional settings updated. This session will terminate. Please login again"

'      Unload Parent
'       MF_DTeRR = True
 
       DtFormatError
       'err.Raise ctErrDtFormat, , "Date format switched inbetween sessions"
       Exit Sub
    End If
    
    If mf_ExcelReportMode Then Exit Sub
    Dim pr_Cancel As Boolean, pr_ErrMsg As String, wl_reccount As Long, ws_rcnd As String, ws_rmsg As String
    On Error GoTo errh
    
    '*** Jenny (On 30/12/04)
    If BtnValFailed <> "" Or ms_ValFailed <> "" Then
        ms_ValFailed = ""
        BtnValFailed = ""
        Exit Sub
    End If
    '*** Jenny (On 30/12/04)
    
    '*** Original Code (Bef 30/12/04)
    'If ms_ValFailed <> "" Then
    '    ms_ValFailed = ""
    '    Exit Sub
    'End If
    '*** Original Code (Bef 30/12/04)
    
    
' ########################  Manoj #### Ver: 2.0.6 #### Date: 16/06/2004  ########################
  'Check Login Session is Working or Not
    If Not ChkLogin(Parent.adc) Then Exit Sub
' ###########################################  Manoj  ###########################################
    MWLib.BeginProcess Parent, "Fetching Record"
    pr_Cancel = False
    pr_ErrMsg = ""
    RaiseEvent KeyValid(pr_Cancel, pr_ErrMsg)
    
    If pr_Cancel = True Then
        err.Raise vbObjectError + 513, "", pr_ErrMsg
    End If
    If Index = 0 Then me_AddFndMode = xaddmode
    If Index = 1 Then me_AddFndMode = xFndMode
    
    'Call FetchxData
    wl_reccount = FetchXData
                         
    ' Called by Automate AddFnd.
    If Index = 99 And wl_reccount = 0 Then me_AddFndMode = xaddmode
    If Index = 99 And wl_reccount <> 0 Then me_AddFndMode = xFndMode
    If me_DataMode = xdatBound Then
        If me_AddFndMode = xaddmode And (Not AllowAdd) Then err.Raise vbObjectError + 513, "", "Add not permitted"
        If me_AddFndMode = xaddmode And wl_reccount <> 0 Then err.Raise vbObjectError + 513, "", "Record already exists"
        If me_AddFndMode = xFndMode And (Not AllowFind) Then err.Raise vbObjectError + 513, "", "Find not permitted"
        If me_AddFndMode = xFndMode And wl_reccount = 0 Then err.Raise vbObjectError + 513, "", "Record not found"
        If me_AddFndMode = xaddmode Then RaiseEvent AutoGen
    End If

    Call RefreshData
    MWLib.EndProcess Parent
    
    Call proc_KEYED("N")
    mf_Opt = ""
    Exit Sub
errh:
    Call HandleError
'    If Not (mo_FirKeyCtl Is Nothing) Then
 '       If mo_FirKeyCtl.Enabled And mo_FirKeyCtl.Visible Then mo_FirKeyCtl.SetFocus
  '  End If
    mf_Opt = ""
    Exit Sub
End Sub
Private Sub tp()
    Dim e As String
    On Error GoTo errh
    e = Parent.Name
    Exit Sub
errh:
    e = Parent.Name
    Exit Sub
End Sub
Private Sub arr_cmd_adfnd_GotFocus(Index As Integer)
    Call CmdClick
    arr_cmd_adfnd(0).TabIndex = 0
End Sub

'*** Jenny
Private Sub arr_cmd_adfnd_LostFocus(Index As Integer)
  'arr_cmd_adfnd(Index).ClrFocus
End Sub
Private Sub arr_cmd_svcndl_LostFocus(Index As Integer)
  'arr_cmd_svcndl(Index).ClrFocus
End Sub
'*** Jenny

Private Sub CMD_EXIT_gotfocus()
    Call CmdClick
End Sub
Private Sub arr_cmd_svcndl_Click(Index As Integer)
    Dim pr_Cancel As Boolean, pr_ErrMsg As String, wd_BookMark As Double
    On Error GoTo errh
    If MWLib.GetDateFormat <> MWLib.DtSetting Then
'      Parent.adc.AllowSave = False
 '     Parent.adc.AllowClose = True
       MsgBox "Regional settings updated. This session will terminate. Please login again"

'      Unload Parent
'       MF_DTeRR = True
 
       DtFormatError
'       err.Raise ctErrDtFormat, , "Date format switched inbetween sessions"
       Exit Sub
    End If
    
    '*** Jenny (On 30/12/04)
    If BtnValFailed <> "" Or ms_ValFailed <> "" Then
        ms_ValFailed = ""
        BtnValFailed = ""
        Exit Sub
    End If
    '*** Jenny (On 30/12/04)
    
    '*** Original Code (Bef 30/12/04)
    'If ms_ValFailed <> "" Then
    '    ms_ValFailed = ""
    '    Exit Sub
    'End If
    '*** Original Code (Bef 30/12/04)
    
    wd_BookMark = md_BookMark
    Select Case Index
    Case 0  'save option
' ########################  Manoj #### Ver: 2.0.6 #### Date: 16/06/2004  ########################
    ' Check Login Session is Working or Not
        If me_EntType = xetReport Then If Not ChkLogin(Parent.adc) Then Exit Sub
' ###########################################  Manoj  ###########################################
        
        MWLib.BeginProcess Parent, "Saving Changes"
        pr_Cancel = False: pr_ErrMsg = ""
        RaiseEvent NonKeyValid(pr_Cancel, pr_ErrMsg)
        If pr_Cancel = True Then err.Raise vbObjectError + 513, "", pr_ErrMsg
        If me_DataMode = xdatBound Then
            If Not RowSave Then err.Raise vbObjectError + 514, "", "X Save Failed"
        End If
        MWLib.EndProcess Parent
        If me_EntType <> xetReport Then
            Call proc_KEYED("K")
            Call CloseRes
        Else
            Call ShowRep
        End If
    Case 1  'close option
        If me_EntType <> xetReport Then
            Call proc_KEYED("K")
            Call CloseRes
        Else
            Call CloseRep
        End If
    Case 2 ' delete option
        Dim wyn As Long
        wyn = MsgBox("R U Sure                 ", vbYesNo + vbDefaultButton2, "Delete Confirmation")
' ########################  Manoj #### Ver: 2.0.6 #### Date: 16/06/2004  ########################
' Check Login Session is Working Or Not
'        If wyn = vbYes Then
        If wyn = vbYes And ChkLogin(Parent.adc) Then
' ###########################################  Manoj  ###########################################
            MWLib.BeginProcess Parent, "Deleting Data"
            md_BookMark = -1
            If me_DeleteMode = xdelChildRecs Then
                If Not DelChild Then err.Raise vbObjectError + 514, "", "Delete of ChildRecs failed"
            End If
            If me_DataMode = xdatBound Then mo_Rowset.Delete
            MWLib.EndProcess Parent
            Call proc_KEYED("K")
            Call CloseRes
        Else
            arr_cmd_svcndl(2).SetFocus
        End If
    End Select
    mf_Opt = ""
'    MWLib.EraseMsg
    Exit Sub
errh:
    Call HandleError
    md_BookMark = wd_BookMark
    If Index = 2 And Not (mo_Rowset Is Nothing) Then
        mo_Rowset.CancelUpdate
    End If
'    If Not (mo_FirNKeyCtl Is Nothing) Then
 '       If mo_FirNKeyCtl.Enabled And mo_FirNKeyCtl.Visible Then mo_FirNKeyCtl.SetFocus
  '  End If
    mf_Opt = ""
End Sub
Private Sub arr_cmd_svcndl_GotFocus(Index As Integer)
    Call CmdClick
End Sub
Private Sub cmd_exit_Click()
    Dim i As Integer
    DispMsg "", etInfo
    For i = 1 To mc_FlxColln.Count
        Call Parent.Controls(mc_FlxColln(i)).TermCtl
    Next i
    Set mc_CtrlColln = Nothing
    Set mc_FlxColln = Nothing
'    Unload Parent
 '   Set mo_Connection = Nothing
'   changing the order of above 2 lines, rn 25.7,

    '*** (Jen 04/02/08)
    Set mo_Connection = Nothing   '*** (Bef speed) bef 24/11/06
    '*** (Jen 04/02/08)
    
    Unload Parent
    
' ########################  Manoj #################### Date: 03/06/2004  ########################
  ' Enable the MenuItem which is Placed on MenuForm
'    If ms_MenuCd <> "" Then g_MDIApp.Controls(ms_MenuCd).Enabled = True
    If ms_MenuCd <> "" Then mo_MnuFrm.Controls(ms_MenuCd).Enabled = True
    mo_MnuFrm.ZOrder (0)
' ###########################################  Manoj  ###########################################

    '*** Jenny Color (21/03/05)
    If ms_ToolBarCtrl <> "" Then
      If mi_ToolBarCtrlIndx >= 0 Then
        g_MDIApp.Controls(ms_ToolBarCtrl).Item(mi_ToolBarCtrlIndx).Enabled = True
      Else
        g_MDIApp.Controls(ms_ToolBarCtrl).Enabled = True
      End If
      
    End If
    '*** Jenny Color (21/03/05)
End Sub
Private Sub cmd_exit_LostFocus()
    If me_EntType <> xetReport Then
        If mf_Automate Then
            If m_ActScrn.ActiveControl.Container.Name <> "FRA_KEYS" And m_ActScrn.ActiveControl.Name <> "ADC" Then
                Call AutomateAdFnd
            End If
        End If
    End If
End Sub
Private Sub UserControl_EnterFocus()
    Dim t As Long
    If Not mf_Hkey Then t = Beep(400, 200)
End Sub
Private Sub UserControl_ExitFocus()
    If me_AddFndMode = xaddmode Then
        arr_cmd_adfnd(0).TabIndex = 0
    Else
        arr_cmd_adfnd(1).TabIndex = 0
    End If
End Sub
Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
    me_AddFndMode = xaddmode
    me_DataMode = xdatBound
    Automate = PropBag.ReadProperty("Automate", False)
    EntType = PropBag.ReadProperty("EntType", 0)
    mf_AllowAdd = PropBag.ReadProperty("AllowAdd", True)
    mf_AllowFind = PropBag.ReadProperty("AllowFind", True)
    Call SetAdFndPos
    mf_AllowSave = PropBag.ReadProperty("AllowSave", True)
    mf_AllowCancel = True
    mf_AllowDelete = PropBag.ReadProperty("AllowDelete", True)
    Call SetSvCnDlPos
    me_DeleteMode = PropBag.ReadProperty("DeleteMode", 0)
    mf_AllowAddRt = True
    mf_AllowFindRt = True
    mf_AllowSaveRt = True
    mf_AllowCancelRt = True
    mf_AllowDeleteRt = True
    mf_AddSave = False
    md_BookMark = 0
    mf_DataChanged = False
    mf_ChkData = False
    
    
    '*** Jenny
    BckColor = PropBag.ReadProperty("BckColor", &H8000000F)
    BckColorButton = PropBag.ReadProperty("BckColorButton", &H8000000F)
    ReSizeBtnPic = PropBag.ReadProperty("ReSizeBtnPic", True)   '*** 07/07/05
    BtnPicture = PropBag.ReadProperty("BtnPicture", UserControl.Picture)
    BtnWidth = PropBag.ReadProperty("BtnWidth", arr_cmd_adfnd(0).Width)
    BtnHeight = PropBag.ReadProperty("BtnHeight", arr_cmd_adfnd(0).Height)
    BtnMaskColor = PropBag.ReadProperty("BtnMaskColor", &HFFFFFF)
    BtnForeColor = PropBag.ReadProperty("BtnForeColor", &H0&)
    BtnDisableColor = PropBag.ReadProperty("BtnDisableColor", &HB4B4B4)
    FontSz = PropBag.ReadProperty("FontSz", 8)
    Pic = PropBag.ReadProperty("Pic", LoadPicture())
    
      '*** Jenny1
      PCT_AF.BackColor = Parent.BackColor
      PCT_SCD.BackColor = Parent.BackColor
      '*** Jenny1
    '*** Jenny
    
End Sub
Private Sub UserControl_WriteProperties(PropBag As PropertyBag)
    Call PropBag.WriteProperty("AllowAdd", mf_AllowAdd, True)
    Call PropBag.WriteProperty("AllowFind", mf_AllowFind, True)
    Call PropBag.WriteProperty("AllowSave", mf_AllowSave, True)
    Call PropBag.WriteProperty("AllowDelete", mf_AllowDelete, True)
    Call PropBag.WriteProperty("Automate", mf_Automate, False)
    Call PropBag.WriteProperty("DeleteMode", me_DeleteMode, 0)
    Call PropBag.WriteProperty("EntType", me_EntType, 0)
    
    '*** Jenny
    Call PropBag.WriteProperty("BckColor", UserControl.BackColor, &H8000000F)
    Call PropBag.WriteProperty("BckColorButton", arr_cmd_adfnd(0).BackColor, &H8000000F)
    Call PropBag.WriteProperty("ReSizeBtnPic", arr_cmd_adfnd(0).ReSizePic, True) '*** 07/07/05
    Call PropBag.WriteProperty("BtnPicture", UserControl.Picture, LoadPicture())
    Call PropBag.WriteProperty("BtnWidth", arr_cmd_adfnd(0).Width, 1575)
    Call PropBag.WriteProperty("BtnHeight", arr_cmd_adfnd(0).Height, 555)
    Call PropBag.WriteProperty("BtnMaskColor", arr_cmd_adfnd(0).MaskColor, &HFFFFFF)
    Call PropBag.WriteProperty("BtnForeColor", arr_cmd_adfnd(0).ForeColor, &H0&)
    Call PropBag.WriteProperty("BtnDisableColor", arr_cmd_adfnd(0).DisableColor, &HB4B4B4)
    Call PropBag.WriteProperty("FontSz", arr_cmd_adfnd(0).LblFont.Size, 8)
    Call PropBag.WriteProperty("Pic", PCT_AF.Picture, LoadPicture())
    '*** Jenny
End Sub
Private Sub UserControl_Resize()
    'UserControl.Height = 495
    '*** Jenny
    UserControl.Height = 550
    'UserControl.Height = arr_cmd_adfnd(0).Height
    '*** Jenny
End Sub
'***************** Events End ***************************************************************'
'***************** Methods Begin ************************************************************'
Public Function HotKeys(KeyCode As Integer, Shift As Integer) As Boolean
    If ms_ValFailed <> "" Then
        ms_ValFailed = ""
        Exit Function
    End If
'??? not reqd if hotkeys of VB work
    'Dim wt_ReqCtl As CommandButton     '*** Old Code (Bef 29/12/2004)
    Dim wt_ReqCtl As MWCTL_BTN1         '*** Jenny (On 29/12/2004)
    
    Dim ws_Name As String               '*** Jenny (On 20/01/2005)
    
    mf_Hkey = False
    Select Case KeyCode
    Case vbKeyA And Shift = 2
        Set wt_ReqCtl = arr_cmd_adfnd(0)
        mf_Opt = "A"
        mf_Hkey = True
        ws_Name = "arr_cmd_adfnd"     '*** Jenny (On 20/01/2005)
    Case vbKeyF And Shift = 2
        Set wt_ReqCtl = arr_cmd_adfnd(1)
        mf_Opt = "F"
        mf_Hkey = True
        ws_Name = "arr_cmd_adfnd"     '*** Jenny (On 20/01/2005)
    Case vbKeyX And Shift = 2
        Set wt_ReqCtl = cmd_exit
        mf_Opt = "X"
        mf_Hkey = True
        ws_Name = "cmd_exit"          '*** Jenny (On 20/01/2005)
    Case vbKeyS And Shift = 2
        Set wt_ReqCtl = arr_cmd_svcndl(0)
        mf_Opt = "S"
        mf_Hkey = True
        ws_Name = "arr_cmd_svcndl"    '*** Jenny (On 20/01/2005)
    Case vbKeyC And Shift = 2
        Set wt_ReqCtl = arr_cmd_svcndl(1)
        mf_Opt = "C"
        mf_Hkey = True
        ws_Name = "arr_cmd_svcndl"    '*** Jenny (On 20/01/2005)
    Case vbKeyD And Shift = 2
        Set wt_ReqCtl = arr_cmd_svcndl(2)
        mf_Opt = "D"
        mf_Hkey = True
        ws_Name = "arr_cmd_svcndl"    '*** Jenny (On 20/01/2005)
    End Select
    
    
    '*** Jenny (On 20/01/2005)
    If mf_Hkey Then
        If wt_ReqCtl.Enabled And wt_ReqCtl.Visible Then
            'Debug.Print UserControl.ActiveControl.Name&; ": " & UserControl.ActiveControl.Index '*** Tmp By Jenny
            
            If Not (UserControl.ActiveControl Is Nothing) Then
              If UserControl.ActiveControl.Name = ws_Name Then
                 Call CmdClick
              Else
                  wt_ReqCtl.SetFocus
              End If
            Else
              wt_ReqCtl.SetFocus
            End If
        Else
            mf_Opt = ""
        End If
        HotKeys = True
    End If
    '*** Jenny (On 20/01/2005)

    '*** Old Code (Bef 20/01/2005)
    'If mf_Hkey Then
    '    If wt_ReqCtl.Enabled And wt_ReqCtl.Visible Then
    '        If UserControl.ActiveControl Is wt_ReqCtl Then
    '           Call CmdClick
    '        Else
    '            wt_ReqCtl.SetFocus
    '        End If
    '    Else
    '        mf_Opt = ""
    '    End If
    '    HotKeys = True
    'End If
    '*** Old Code (Bef 20/01/2005)
End Function
Public Function InitRtn() As Boolean
    On Error GoTo errh
    InitRtn = False
    Dim wo_Control As Control
    
    '*** (Bef speed) bef 24/11/06
    'Set mo_Connection = MWLib.SetConn
    '*** (Bef speed) bef 24/11/06
    
    '*** (Jenny speed) 04/02/08
    If me_EntType = xetEntry Then
      Set mo_Connection = MWLib.SetConn
    Else
      'if me_EntType = xetReport
      Set mo_Connection = MWLib.SetRepConn
    End If
    '*** (Jenny speed) 04/02/08
    
    ''*** (Jenny speed) 16/11/06 for a single connection
    ''If (MWLib.GetConn Is Nothing) Then
    ''  Set mo_Connection = MWLib.SetConn
    ''Else
    ''  Set mo_Connection = MWLib.GetConn
    ''End If
    ''*** (Jenny speed) 16/11/06 for a single connection
    
    Set mc_CtrlColln = New Collection
    For Each wo_Control In Parent.Controls
        If TypeOf wo_Control Is MWCTL_MED Then
            mc_CtrlColln.Add wo_Control, wo_Control.IdName
            wo_Control.SetToolTip
        End If
    Next
    Set mc_FlxColln = New Collection
    For Each wo_Control In Parent.Controls
        If TypeOf wo_Control Is MWCTL_FLX Then
            mc_FlxColln.Add wo_Control.Name
        End If
    Next
    RaiseEvent Load
    Call proc_KEYED("K")
    If mf_Automate And mo_FirKeyCtl Is Nothing Then Call AutomateAdFnd    ' mainly for report
    If me_EntType = xetReport Then AllowCancel = False
' ########################  Manoj #################### Date: 03/06/2004  ########################
  ' Disable the MenuItem Placed on MenuForm
'    If ms_MenuCd <> "" Then g_MDIApp.Controls(ms_MenuCd).Enabled = False
    If Not mo_MnuFrm Is Nothing And ms_MenuCd <> "" Then mo_MnuFrm.Controls(ms_MenuCd).Enabled = False
      'g_MDIApp.Controls(ms_MenuCd).Enabled = False
' ###########################################  Manoj  ###########################################

    '*** Jenny Color (21/03/05)
    If ms_ToolBarCtrl <> "" Then
      If mi_ToolBarCtrlIndx >= 0 Then
        g_MDIApp.Controls(ms_ToolBarCtrl).Item(mi_ToolBarCtrlIndx).Enabled = False
      Else
        g_MDIApp.Controls(ms_ToolBarCtrl).Enabled = False
      End If
      
    End If
    '*** Jenny Color (21/03/05)

    InitRtn = True
    Exit Function
errh:
    Call HandleError
End Function
                  ' Called by 1. InitRtn,
                  ' ??Got of Med (NonKey when prev is Key)
                  ' ?? Exit Lost Focus
Public Sub AutomateAdFnd()
    If ms_ValFailed <> "" Then
        ms_ValFailed = ""
        Exit Sub
    End If
    Call arr_cmd_adfnd_Click(99)
End Sub
Public Function XChild(ByVal pv_ChildNm As String) As Boolean
    XChild = False
    If pv_ChildNm <> "" And InStr(1, "," + UCase(ms_child) + ",", "," + UCase(pv_ChildNm) + ",") > 0 Then XChild = True
End Function
Public Function SaveRec() As Boolean
    SaveRec = RowSave
End Function
Public Function AddSaveRec() As Boolean
    mf_AddSave = True
    AddSaveRec = RowSave
    mf_AddSave = False
End Function
Public Sub DtFormatError()
  On Error GoTo errh
  Dim target_hwnd As Long
  Dim target_process_id As Long
  Dim target_process_handle As Long
  Unload Parent
  Dim C As Variant
  Set C = g_MDIApp
  g_MDIApp.DtFormatError
errh:
  Exit Sub
End Sub
Public Sub HandleError()
    Dim i As Integer, ws_errStr As String, wl_errNum As Long, wr_pos As Integer
    ws_errStr = err.Description
    wl_errNum = err.Number
    Call tp
    MWLib.EndProcess Parent
    wr_pos = InStr(1, ws_errStr, ct_SqlErrMsg)
    If wr_pos <> 0 Then ws_errStr = Mid(ws_errStr, Len(ct_SqlErrMsg) + 1)
    If wl_errNum = ctErrUpdConflict And me_EntType = xetEntry Then
        If Parent("FRA_NKEYS").Enabled = True Then AllowCancel = True
        MsgBox "Please logout"

    End If
    If wl_errNum = ctErrDtFormat And me_EntType = xetEntry Then
        If Parent("FRA_NKEYS").Enabled = True Then AllowCancel = True
        Unload Parent
    End If
    
    '*** (Jen 04/02/08)
    mo_Connection.ErrClear   '*** (Bef speed) bef 24/11/06
    '*** (Jen 04/02/08)
    
    '*** (Bef 04/02/08)
    'MWLib.GetConn.ErrClear    '*** (Jenny speed 24/11/06)
    '*** (Bef 04/02/08)
    
    If wl_errNum = vbObjectError + 514 Then
        DispMsg ws_errStr, en_ErrorType.etInfo
    ElseIf wl_errNum = vbObjectError + 513 Then
        DispMsg ws_errStr, en_ErrorType.etWarning
    Else
        DispMsg ws_errStr, en_ErrorType.etError
    End If

    err.Clear
End Sub
Public Sub Resync(ByVal pv_AddSaveOpt As en_AddSaveOpt)
    If pv_AddSaveOpt = asAll Then mo_Rowset.Resync mdoryAll
    If pv_AddSaveOpt = asCurrent Then mo_Rowset.Resync mdoryCurrent
End Sub
Public Sub AddSave(ByVal pv_ObjNm As String, ByVal pv_AddOpt As en_AddSaveOpt)
    Dim wa_arr As Variant
    On Error GoTo errh
    If mc_AddSave Is Nothing Then Set mc_AddSave = New Collection
    wa_arr = Array(pv_ObjNm, pv_AddOpt)
    mc_AddSave.Add wa_arr, pv_ObjNm
    Exit Sub
errh:
    If err.Number = 457 Then
        Exit Sub
    Else
        err.Raise err.Number, "", err.Description
    End If
End Sub
Public Sub StartCopy()
    me_Mode = xCopy
End Sub
Public Sub EndCopy()
    me_Mode = xNorm
End Sub
Public Sub CloseConn()
    Set mc_CtrlColln = Nothing
    Set mc_FlxColln = Nothing
    '*** (Jen 04/02/08)
    Set mo_Connection = Nothing     '*** (Bef speed) bef 24/11/06
    '*** (Jen 04/02/08)
End Sub
'***************** Methods End ************************************************************'
'***************** Subs Begin ************************************************************'
Private Sub CloseRep()
    If me_EntType = xetReport Then
        With Parent.Controls("FRA_REPOBJ")
            If .Visible = True Then
                .Visible = False
                AllowSave = True
                Parent.Controls("DST_REP").Enabled = True
                If arr_cmd_svcndl(0).Enabled And arr_cmd_svcndl(0).Visible Then arr_cmd_svcndl(0).SetFocus
                AllowCancel = False
            End If
        End With
    End If
End Sub

Private Sub ShowRep()
 '****** Sachin - As per Instructions From Sir ******
On Error GoTo errh
    If Parent.Controls("FRA_REPOBJ").Visible = False Then
        MWLib.BeginProcess Parent, "Processing Report SQL"
          RaiseEvent SetRepRecSource
        MWLib.EndProcess Parent
    End If
Exit Sub
errh:
    MWLib.EndProcess Parent
    Call HandleError
'****** Sachin - As per Instructions From Sir ******
End Sub

Public Sub ShowRepDone()
    On Error GoTo errh
    Dim wl_PaperOrientation As Long, wl_PaperSize As Long
    If Parent.Controls("FRA_REPOBJ").Visible = False Then
        MWLib.BeginProcess Parent, "Processing Report SQL"
        If ms_RepRecSource <> "" Then
            mo_RepSource.DiscardSavedData
            
            '*** (Jen 04/02/08)
            '*** (Bef speed) bef 24/11/06
            If mo_Connection.RepRes(mo_RepSource, ms_RepRecSource) = 0 Then
            '*** (Bef speed) bef 24/11/06
            '*** (Jen 04/02/08)
            
            '*** (Bef 04/02/08)
            '*** (Jenny speed 24/11/06)
            '****** Sachin 3.02 06/09/07 - Setting report connection
            'If MWLib.GetRepConn.RepRes(mo_RepSource, ms_RepRecSource) = 0 Then
            '****** Sachin 3.02 06/09/07 - Setting report connection
            '*** (Jenny speed 24/11/06)
            '*** (Bef 04/02/08)
                err.Raise vbObjectError + 512, "", "No Data Found for given scope"
            End If
            '****** Sachin 3.02 06/09/07  ******
        Else
          MWLib.EndProcess Parent
          Exit Sub
        End If
        '****** Sachin 3.02 06/09/07  ******
        
        '****** Sachin 3.2.2 - mo_repsource shifted below - as it gives error while viewing report on screen after exporting to file ******
        'Parent.Controls("CRV_REP").ReportSource = mo_RepSource

' ########################  Manoj #################### Date: 30/04/2004  ########################
  ' set Printer & Paper Settings
        If Not (Printer Is Nothing) And Printers.Count > 0 Then
          wl_PaperOrientation = mo_RepSource.PaperOrientation
          wl_PaperSize = mo_RepSource.PaperSize
          mo_RepSource.SelectPrinter Printer.DriverName, Printer.DeviceName, Printer.Port
          mo_RepSource.PaperOrientation = wl_PaperOrientation
          mo_RepSource.PaperSize = wl_PaperSize
        End If
' ###########################################  Manoj  ###########################################
        MWLib.EndProcess Parent
        If Parent.Controls("DST_REP").DstOpt = repOnScreen And mo_RepSource.ExportOptions.DiskFileName = "" Then        '3.11.4
            Parent.Controls("CRV_REP").ReportSource = mo_RepSource    '****** Sachin 3.2.2 - Line above shifted here
            MWLib.BeginProcess Parent, "Showing Report"
            With Parent.Controls("FRA_REPOBJ")
                .Left = 0
                .Top = 0
                .Height = Parent.Height - UserControl.Height - 300 'for caption of form
                .Width = Parent.Width
                .Visible = True
            End With
            With Parent.Controls("CRV_REP")
                .Left = 0
                .Top = 0
                .Height = Parent.Height - UserControl.Height - 300 'for caption of form
                .Width = Parent.Width
' ########################  Manoj #################### Date: 30/04/2004  ########################
' For Comuters having screen Resolution more than (800 * 600)
                Parent.Controls("ADC").Left = 0
                Parent.Controls("ADC").Top = IIf(Parent.ScaleHeight > 450, Parent.ScaleHeight - 450, Parent.ScaleHeight)
                Parent.Controls("ADC").Width = IIf(Parent.ScaleWidth > 50, Parent.ScaleWidth - 50, Parent.ScaleWidth)
                Parent.Controls("DST_REP").Top = Parent.Controls("ADC").Top
                Parent.Controls("FRA_KEYS").Top = Parent.Controls("ADC").Top
' ###########################################  Manoj  ###########################################
                .ZOrder
                MWLib.EndProcess Parent
                .ViewReport
'??? think of not using doevents
                DoEvents
                AllowCancel = True
                AllowSave = False
                'urmi 12/03/04 Undo disable DST_REP
                'Parent.Controls("DST_REP").Enabled = False
                If arr_cmd_svcndl(1).Enabled And arr_cmd_svcndl(1).Visible Then arr_cmd_svcndl(1).SetFocus
            End With
        Else
            MWLib.BeginProcess Parent, "Showing Report"
            Call Parent.Controls("DST_REP").ShowReport
            MWLib.EndProcess Parent
        End If
    End If
    Exit Sub
errh:
' ########################  Manoj #################### Date: 30/04/2004  ########################
    If err.Number = 484 Then Resume Next
' ###########################################  Manoj  ###########################################
    MWLib.EndProcess Parent
    Call HandleError
End Sub
Private Sub RefreshData()
    If me_DataMode = xdatBound Then
      If mo_Rowset.RecCount > 0 Then md_BookMark = mo_Rowset.Bookmark
    End If
    Dim wa_rArr() As String, i As Integer
    If me_DataMode = xdatBound Then
        me_Mode = xInit
        If me_AddFndMode = xFndMode Then
            Call MoveData(ctm_Read)
        Else
            Call MoveData(ctm_Init)
        End If
        me_Mode = xNorm
    End If
    wa_rArr = Split(ms_child, ",")
    If IsArray(wa_rArr) Then
        For i = 0 To UBound(wa_rArr)
            Call Parent.Controls(wa_rArr(i)).RefreshData
        Next i
    End If
End Sub
Private Function DelChild() As Boolean
                  ' called by SvCnDl_Click()
    Dim ws_rArr() As String, i As Integer
    DelChild = True
    ws_rArr = Split(ms_child, ",")
    If IsArray(ws_rArr) Then
        For i = 0 To UBound(ws_rArr)
            If Not Parent.Controls(ws_rArr(i)).DelAllRecs Then
                DelChild = False
                Exit Function
            End If
        Next i
    End If
End Function

                        'pv_Opt Read, Init, InitAll, Write.
                        ' mf_ChkData true then move value,
                        ' me_dataChanged = true if value is changed & ctm_Write
Private Sub MoveData(ByVal pv_Opt As String)
    On Error GoTo errh
    mf_DataChanged = False
    Dim wo_MED As MWCTL_MED, ws_rdat As String, wo_Ctl As Control
    Dim wo_Pic As PictureBox, wo_ImgCombo As ImageCombo   '*** Jenny
    
    For Each wo_Ctl In Parent.Controls
        If TypeOf wo_Ctl Is MWCTL_MED Then
            Set wo_MED = wo_Ctl
            With wo_MED
                ws_rdat = .DataField
                
                '*** (Bef 28/03/06)
                If pv_Opt = ctm_Read And ws_rdat <> "" Then .Value = mo_Rowset.FldValue(ws_rdat)
                '*** (Bef 28/03/06)
                '*** Jenny (28/03/06) This works too
                'Note: Collect is undocumented feature of ADO but it's faster than any other reference to recordset fields.
                '      Or why not call MS Access action query or SQL server stored procedure as connection method.
                
                'If pv_Opt = ctm_Read And ws_rdat <> "" Then .Value = mo_Rowset.CollectFldVal(ws_rdat)
                '*** Jenny (28/03/06) This works too
                
                If pv_Opt = ctm_Init And UCase(wo_Ctl.Container.Name) <> "FRA_KEYS" Then .Value = ""
                If pv_Opt = ctm_InitAll Then .Value = ""
                If pv_Opt = ctm_Write And ws_rdat <> "" Then
                    If IsNull(mo_Rowset.FldValue(ws_rdat)) Or IsEmpty(mo_Rowset.FldValue(ws_rdat)) Or (mo_Rowset.FldValue(ws_rdat) <> .Value) Then
                        If Not mf_ChkData Then mo_Rowset.FldValue(ws_rdat) = .Value
                        mf_DataChanged = True
                    End If
                End If
            End With
            
        '*** Jenny
        ElseIf TypeOf wo_Ctl Is PictureBox Then
            Set wo_Pic = wo_Ctl
            With wo_Pic
              ws_rdat = .DataField
              If ws_rdat <> "" Then
                If pv_Opt = ctm_Read Then .Picture = mo_Rowset.FldPicValue(ws_rdat)
                If pv_Opt = ctm_Init And UCase(wo_Ctl.Container.Name) <> "FRA_KEYS" Then .Picture = LoadPicture()
                If pv_Opt = ctm_InitAll Then .Picture = LoadPicture()
                If pv_Opt = ctm_Write And ws_rdat <> "" Then
                    'If IsNull(mo_Rowset.FldValue(ws_rdat)) Or IsEmpty(mo_Rowset.FldValue(ws_rdat)) Or (mo_Rowset.FldValue(ws_rdat) <> .Value) Then
                    'If mo_Rowset.FldPicValue(ws_rdat) <> .Picture Then
                        If Not mf_ChkData Then mo_Rowset.FldPicValue(ws_rdat) = .Picture
                        mf_DataChanged = True
                    'End If
                End If
              End If
            End With
        ElseIf TypeOf wo_Ctl Is ImageCombo Then
            '*** Note: The Component 'Microsoft Windows Common Controls 6.0' was added in this project just for the ImageCombo
            Set wo_ImgCombo = wo_Ctl
            With wo_ImgCombo
                ws_rdat = .DataField
                If pv_Opt = ctm_Read And ws_rdat <> "" Then .Text = mo_Rowset.FldValue(ws_rdat)
                If pv_Opt = ctm_Init And UCase(wo_Ctl.Container.Name) <> "FRA_KEYS" Then .Text = ""
                If pv_Opt = ctm_InitAll Then .Text = ""
                If pv_Opt = ctm_Write And ws_rdat <> "" Then
                    If IsNull(mo_Rowset.FldValue(ws_rdat)) Or IsEmpty(mo_Rowset.FldValue(ws_rdat)) Or (mo_Rowset.FldValue(ws_rdat) <> .Text) Then
                        If Not mf_ChkData Then mo_Rowset.FldValue(ws_rdat) = .Text
                        mf_DataChanged = True
                    End If
                End If
            End With
        '*** Jenny
        End If
    Next
    If me_EntType = xetReport Then Call MoveGrpSrt(pv_Opt)
    Exit Sub
errh:
    err.Raise vbError + 512, "", err.Description + " (" + ws_rdat + ") "
End Sub
                        ' called by MoveData
                        '?? err no 438
Private Sub MoveGrpSrt(ByVal pv_Opt As String)
    On Error GoTo errh
    With Parent.GRP_REP
       If pv_Opt = ctm_Read Then
            .Value(gltGroup1) = mo_Rowset!Group1
            .Value(gltGroup2) = mo_Rowset!Group2
            .Value(gltGroup3) = mo_Rowset!Group3
            .Value(gltsort1) = mo_Rowset!Sort1
            .Value(gltsort2) = mo_Rowset!Sort2
            .Value(gltsort3) = mo_Rowset!Sort3
        End If
        If pv_Opt = ctm_Write Then
            If mo_Rowset!Group1 <> .Value(gltGroup1) Then
                If Not mf_ChkData Then mo_Rowset!Group1 = .Value(gltGroup1)
                mf_DataChanged = True
            End If
            If mo_Rowset!Group2 <> .Value(gltGroup2) Then
                If Not mf_ChkData Then mo_Rowset!Group2 = .Value(gltGroup2)
                mf_DataChanged = True
            End If
            If mo_Rowset!Group3 <> .Value(gltGroup3) Then
                If Not mf_ChkData Then mo_Rowset!Group3 = .Value(gltGroup3)
                mf_DataChanged = True
            End If
            If mo_Rowset!Sort1 <> .Value(gltsort1) Then
                If Not mf_ChkData Then mo_Rowset!Sort1 = .Value(gltsort1)
                mf_DataChanged = True
            End If
            If mo_Rowset!Sort2 <> .Value(gltsort2) Then
                If Not mf_ChkData Then mo_Rowset!Sort2 = .Value(gltsort2)
                mf_DataChanged = True
            End If
            If mo_Rowset!Sort3 <> .Value(gltsort3) Then
                If Not mf_ChkData Then mo_Rowset!Sort3 = .Value(gltsort3)
                mf_DataChanged = True
            End If
        End If
    End With
    Exit Sub
errh:
    If err.Number = 438 Then
        Exit Sub
    Else
        Call HandleError
    End If
End Sub
                        ' called by rowsave
Private Function MoveUsrDets(ByVal pv_Opt As String) As Boolean
    On Error GoTo errh
    MoveUsrDets = False
    If pv_Opt = ctm_Write Then
        mo_Rowset!ModUsr = ms_UsrCd
        
        '*** (Jen 04/02/08)
        '*** (Bef speed) bef 24/11/06
        mo_Rowset!ModDt = mo_Connection.SrvrDate
        mo_Rowset!modTime = mo_Connection.SrvrTime
        '*** (Bef speed) bef 24/11/06
        '*** (Jen 04/02/08)
        
        '*** (Bef 04/02/08)
        ''*** (Jenny speed 24/11/06)
        'mo_Rowset!ModDt = MWLib.GetConn.SrvrDate
        'mo_Rowset!modTime = MWLib.GetConn.SrvrTime
        ''*** (Jenny speed 24/11/06)
        '*** (Bef 04/02/08)
    
    End If
    MoveUsrDets = True
    Exit Function
errh:
    err.Raise vbObjectError + 512, "", err.Description + " One of the User Dets flds missing in base table.can u pl  eklekrelkrlekfdf dfdjfj  df df end "
End Function
Private Sub proc_KEYED(ByVal pv_keyed As String)
    On Error GoTo errh
    If pv_keyed = "K" Then
        
        Parent.fra_keys.Enabled = True
        If Not mf_Automate Then
            Parent.fra_nkeys.Enabled = False
        Else
            Parent.fra_nkeys.Enabled = True
        End If

        arr_cmd_adfnd(0).Enabled = True
        arr_cmd_adfnd(1).Enabled = True
                
        cmd_exit.Enabled = True
        
        Call InitForm
        RaiseEvent KeyWhen
        
        If Not (mo_FirKeyCtl Is Nothing) Then
            If MWLib.CtlEnAndVis(mo_FirKeyCtl) Then mo_FirKeyCtl.SetFocus
        End If
        arr_cmd_svcndl(0).Enabled = False
        arr_cmd_svcndl(1).Enabled = False
        arr_cmd_svcndl(2).Enabled = False
    End If
    
    If pv_keyed = "N" Then
        Parent.fra_keys.Enabled = False
        Parent.fra_nkeys.Enabled = True
        
        arr_cmd_svcndl(0).Enabled = True
        
        If me_EntType = xetEntry Then
            If me_AddFndMode = xaddmode Then AllowCancel = True
            If me_AddFndMode = xFndMode Then AllowCancel = False
        End If

        If me_EntType = xetReport Then
            AllowDelete = False
        Else
            If me_AddFndMode = xaddmode And mf_AllowDelete Then AllowDelete = False
            If me_AddFndMode = xFndMode And mf_AllowDelete Then AllowDelete = True
        End If
        
        arr_cmd_adfnd(0).Enabled = False
        arr_cmd_adfnd(1).Enabled = False
        If me_EntType <> xetReport Then cmd_exit.Enabled = False
        
        Dim pr_Cancel As Boolean, pr_ErrMsg As String
        pr_Cancel = False
        RaiseEvent NonKeyWhen(pr_Cancel, pr_ErrMsg)
        If pr_Cancel = True Then
            err.Raise vbObjectError + 513, 0, pr_ErrMsg
        Else
            Call EnaDisaFlx
            If Not (mo_FirNKeyCtl Is Nothing) Then
                If MWLib.CtlEnAndVis(mo_FirNKeyCtl) Then
                    mo_FirNKeyCtl.SetFocus
                End If
            End If
        End If
    End If
    If Not mf_AllowAddRt Then arr_cmd_adfnd(0).Enabled = False
    If Not mf_AllowFindRt Then arr_cmd_adfnd(1).Enabled = False
    If Not mf_AllowDeleteRt Then arr_cmd_svcndl(2).Enabled = False
    If Not mf_AllowSaveRt Then arr_cmd_svcndl(0).Enabled = False
    Exit Sub
errh:
    Call HandleError
    If pv_keyed = "N" And pr_Cancel = True Then
        Parent.fra_nkeys.Enabled = False
        AllowSave = False
        AllowCancel = True
        If mf_AllowDelete Then AllowDelete = False
        Parent.adc.SetFocus
    End If
End Sub
                        ' called by Hotkeys, Gotfocus of ...
Private Sub CmdClick()
    If ms_ValFailed <> "" Then
        BtnValFailed = ms_ValFailed   '*** Jenny
        ms_ValFailed = ""
        Exit Sub
    End If
    If ms_ValFailed = "" Then
        Select Case mf_Opt
        Case Is = "A"
            Call arr_cmd_adfnd_Click(0)
        Case Is = "F"
            Call arr_cmd_adfnd_Click(1)
        Case Is = "S"
            Call arr_cmd_svcndl_Click(0)
        Case Is = "C"
            Call arr_cmd_svcndl_Click(1)
        Case Is = "D"
            Call arr_cmd_svcndl_Click(2)
        Case Is = "X"
            Call cmd_exit_Click
        End Select
    End If
End Sub
Private Sub SetAdFndPos()
    Dim wi_wdth As Integer, wi_Ctr As Integer
    'wi_wdth = 975     '*** Original Code (Bef 29/12/2004)
    wi_wdth = arr_cmd_adfnd(0).Width     '*** Jenny  (On 29/12/2004)
    'wi_wdth = BtnWidth     '*** Jenny  (On 29/12/2004)
    
    arr_cmd_adfnd(0).Left = 0
    arr_cmd_adfnd(1).Left = wi_wdth * 1
    cmd_exit.Left = wi_wdth * 2
    arr_cmd_adfnd(0).Visible = True
    arr_cmd_adfnd(1).Visible = True
    wi_Ctr = 3
    If (mf_Automate) Or (Not mf_AllowAdd) Then
        arr_cmd_adfnd(0).Visible = False
        arr_cmd_adfnd(1).Left = 0
        wi_Ctr = wi_Ctr - 1
    End If
    If (mf_Automate) Or (Not mf_AllowFind) Then
        arr_cmd_adfnd(1).Visible = False
        wi_Ctr = wi_Ctr - 1
    End If
    cmd_exit.Left = (wi_Ctr - 1) * wi_wdth
    PCT_AF.Width = wi_Ctr * wi_wdth + (PCT_AF.BorderStyle * 30 * 2)
End Sub
Private Sub SetSvCnDlPos()
    Dim wi_wdth As Integer, wi_Ctr As Integer
    
    'wi_wdth = 975                        '*** Original Code (Bef 29/12/2004)
    wi_wdth = arr_cmd_svcndl(0).Width     '*** Jenny  (On 29/12/2004)
    
    arr_cmd_svcndl(0).Left = 0
    arr_cmd_svcndl(1).Left = wi_wdth
    arr_cmd_svcndl(2).Left = wi_wdth * 2
    arr_cmd_svcndl(0).Visible = True
    arr_cmd_svcndl(1).Visible = True
    arr_cmd_svcndl(2).Visible = True
    wi_Ctr = 3
    If Not mf_AllowSave Then
        arr_cmd_svcndl(0).Visible = False
        arr_cmd_svcndl(1).Left = 0
        arr_cmd_svcndl(2).Left = wi_wdth
        wi_Ctr = wi_Ctr - 1
    End If
    If Not mf_AllowCancel Then
        arr_cmd_svcndl(1).Visible = False
        arr_cmd_svcndl(2).Left = wi_wdth
        wi_Ctr = wi_Ctr - 1
    End If
    If Not mf_AllowDelete Then
        arr_cmd_svcndl(2).Visible = False
        wi_Ctr = wi_Ctr - 1
    End If
    PCT_SCD.Width = (wi_Ctr * wi_wdth) + (PCT_SCD.BorderStyle * 30 * 2)
End Sub
Private Sub SetSvCnDlEnDis()
    If Parent.fra_nkeys.Enabled Then
        If mf_AllowSaveRt Then
            arr_cmd_svcndl(0).Enabled = True
        Else
            arr_cmd_svcndl(0).Enabled = False
        End If
        If mf_AllowCancelRt Then
        
            '*** Jenny (09/01/06)
            If me_AddFndMode = xaddmode Or me_EntType = xetReport Then
              arr_cmd_svcndl(1).Enabled = True
            Else
              arr_cmd_svcndl(1).Enabled = False
            End If
            '*** Jenny (09/01/06)
            
            '*** (Before 09/01/06)
            'arr_cmd_svcndl(1).Enabled = True
            '*** (Before 09/01/06)
        Else
            arr_cmd_svcndl(1).Enabled = False
        End If
        If mf_AllowDeleteRt Then
            '*** urmi 28/02/04 ******
            If me_AddFndMode <> xaddmode Then arr_cmd_svcndl(2).Enabled = True
        Else
            arr_cmd_svcndl(2).Enabled = False
        End If
    End If
End Sub
Private Sub SetAdFndEnDis()
    If Parent.fra_keys.Enabled Then
        If mf_AllowAddRt Then
            arr_cmd_adfnd(0).Enabled = True
        Else
            arr_cmd_adfnd(0).Enabled = False
        End If
        If mf_AllowFindRt Then
            arr_cmd_adfnd(1).Enabled = True
        Else
            arr_cmd_adfnd(1).Enabled = False
        End If
    End If
End Sub
                        ' called by Get Ctl
Private Function FindObj(ByVal pv_IdName As String) As MWCTL_MED
    Dim wo_Ctl As Control
    For Each wo_Ctl In Parent.Controls
        If TypeOf wo_Ctl Is MWCTL_MED Then
            If UCase(wo_Ctl.IdName) = UCase(pv_IdName) Then
                Set FindObj = wo_Ctl
                Exit For
            End If
        End If
    Next
End Function
                        ' called by KeyEd("K")
Private Sub InitForm()
    Dim i As Integer
    Dim ws_rArr() As String
    me_Mode = xClear
    md_BookMark = 0
    Call MoveData(ctm_InitAll)
    ws_rArr = Split(ms_child, ",")
    If IsArray(ws_rArr) Then
        For i = 0 To UBound(ws_rArr)
            Call Parent.Controls(ws_rArr(i)).Init
        Next i
    End If
    me_Mode = xNorm
End Sub
Private Sub EnaDisaFlx()
                            ' Calls .ChildsEnaDisa for Grids which have Childs.
    Dim i As Integer, wc_Colln As Collection, wo_Ctl As MWCTL_FLX
    Set wc_Colln = Parent.adc.FlxColln
    For i = 1 To wc_Colln.Count
        Set wo_Ctl = Parent.Controls(wc_Colln(i))
        If wo_Ctl.Child <> "" Then Call wo_Ctl.ChildsEnaDisa(False)
    Next i
End Sub
Private Sub CloseRes()
    Dim i As Integer
    Dim ws_rArr() As String
    Set mo_Rowset = Nothing
    ws_rArr = Split(ms_child, ",")
    If IsArray(ws_rArr) Then
        For i = 0 To UBound(ws_rArr)
            Call Parent.Controls(ws_rArr(i)).CloseRes
        Next i
    End If
End Sub
                                  ' Called by Sv_Click, .SaveRec
Private Function RowSave() As Boolean
    On Error GoTo errh
    If MWLib.GetDateFormat <> MWLib.DtSetting Then
'      Parent.adc.AllowSave = False
 '     Parent.adc.AllowClose = True
       MsgBox "Regional settings updated. This session will terminate. Please login again"

'      Unload Parent
'       MF_DTeRR = True
 
       DtFormatError
'       err.Raise ctErrDtFormat, , "Date format switched inbetween sessions"
       Exit Function
    End If
    
    
    RowSave = True
    If me_DataMode = xdatUnbound Then Exit Function
    If mf_AddSave = True And (IsNew Or IsDel) Then Exit Function
    
    If Not IsNew Then
        mf_ChkData = True
        Call MoveData(ctm_Write)
        mf_ChkData = False
    End If
    
    Parent.adc.Connection.BeginTrans (mdoCommitNone)
    If mf_DataChanged = True Or IsNew Then
        If Not AddSaveProc Then err.Raise vbObjectError + 514, 0, "AddSave Failed"
        If mo_Rowset.RecCount = 0 And me_AddFndMode = xaddmode Then
            mo_Rowset.Addnew
        Else
            mo_Rowset.edit
        End If
        If IsNew And me_EntType = xetReport Then Call InitAllFlds
        Call MoveData(ctm_Write)
        If Not MoveUsrDets(ctm_Write) Then err.Raise vbObjectError + 514, 0, "Error on Updating UserDets"
        mo_Rowset.Update
        If md_BookMark = 0 Then md_BookMark = mo_Rowset.Bookmark
        Set mc_AddSave = Nothing
    Else
        mo_Rowset.CancelUpdate
    End If
    
    Parent.adc.Connection.CommitTrans (mdoCommitNone)
    If me_EntType = xetEntry Then
        AllowCancel = False
' ########################  Manoj #### Ver: 2.0.6 #### Date: 03/06/2004  ########################
' After Saving Record in Add Mode, Delete Button Has To be Activate
        arr_cmd_svcndl(2).Enabled = True
' ###########################################  Manoj  ###########################################
    End If
    
    DispMsg "", etInfo
    Exit Function
errh:
    Call HandleError
    mf_Opt = ""
    If Not (mo_Rowset Is Nothing) Then
        mo_Rowset.CancelUpdate
    End If
    RowSave = False
End Function
                                ' Called by RowSave
Private Function AddSaveProc() As Boolean
    Dim i As Integer, ws_ObjNm As String, we_AddSaveOpt As en_AddSaveOpt, j As Integer, e As Integer
    AddSaveProc = True
    If Not (mc_AddSave Is Nothing) Then
        For i = 1 To mc_AddSave.Count
            ws_ObjNm = mc_AddSave(i)(0)
            we_AddSaveOpt = mc_AddSave(i)(1)
            If we_AddSaveOpt = asAll Then
                For j = 1 To Parent.Controls(ws_ObjNm).Rows - 1
                    If Not Parent.Controls(ws_ObjNm).AddSaveRec(j) Then
                        err.Raise vbObjectError + 514, "", "Save of " + ws_ObjNm + " failed "
                    End If
                Next j
            Else
                If Not Parent.Controls(ws_ObjNm).AddSaveRec Then
                    err.Raise vbObjectError + 514, "", "Save of " + ws_ObjNm + " failed "
                End If
            End If
'??? consider speed while calling saverec
'??? check for type before using .rows
        Next i
    End If
    Exit Function
errh:
    AddSaveProc = False
    If Not (mc_AddSave Is Nothing) Then
        For e = 1 To i - 1
            ws_ObjNm = mc_AddSave(e)(0)
            we_AddSaveOpt = mc_AddSave(e)(1)
            Parent.Controls(ws_ObjNm).Resync we_AddSaveOpt
'??? consider speed while calling resync
        Next e
    End If
End Function
                              ' Called by RowSave if Report & isNew
Private Sub InitAllFlds()
    Dim wv_InitVal As Variant, i As Integer
    For i = 0 To mo_Rowset.FldCount - 1
        Select Case mo_Rowset.FldType(i)
        Case Is = mdoftString
            wv_InitVal = ""
        Case Is = mdoftInt
            wv_InitVal = 0
        Case Is = mdoftFloat
            wv_InitVal = 0
        Case Is = mdoftNumeric
            wv_InitVal = 0
        Case Is = mdoftDateTime
            wv_InitVal = "01/01/80"
        End Select
        mo_Rowset.FldValue(i) = wv_InitVal
    Next i
End Sub

'***************** Subs End ****************************************************************'
                              ' Called by AddFnd_Click
Private Function FetchXData() As Long
   RaiseEvent SetRecSource
   If me_EntType = xetReport Then
       ms_RecSource = ""
        ms_RecSource = "select * from usropt where uousrcd='" + ms_UsrCd + "' and uomnucd='" + ms_MenuCd + "'"
    End If
    FetchXData = 0
    If me_DataMode = xdatBound Then
        '*** (Jen 04/02/08)
        '*** (Bef speed) bef 24/11/06
        Set mo_Rowset = mo_Connection.OpenResultset(ms_RecSource)
        '*** (Bef speed) bef 24/11/06
        '*** (Jen 04/02/08)
    
        '*** (Bef 04/02/08)
        ''*** (Jenny speed 24/11/06)
        ''****** Sachin 3.02 - 06/09/07 ******
        'If me_EntType = xetReport Then
        '  Set mo_Rowset = MWLib.GetRepConn.OpenResultset(ms_RecSource)
        'Else
        ''****** Sachin 3.02 - 06/09/07 ******
        '  Set mo_Rowset = MWLib.GetConn.OpenResultset(ms_RecSource)
        'End If
        ''*** (Jenny speed 24/11/06)
        '*** (Bef 04/02/08)
        
        FetchXData = mo_Rowset.RecCount
    End If
End Function
