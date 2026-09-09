VERSION 5.00
Object = "{AB2CC32A-E652-42A2-B341-26ECE02F6386}#2.1#0"; "mwfctlforspeed9.ocx"
Begin VB.Form EmrFrmResAutogen 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Reset Autogen"
   ClientHeight    =   2625
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   8265
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MDIChild        =   -1  'True
   ScaleHeight     =   2625
   ScaleWidth      =   8265
   Begin MwfCtl.MWCTL_BTN1 CmdResAuto 
      Height          =   585
      Left            =   3285
      TabIndex        =   0
      ToolTipText     =   "Option To Reset The Autogen Table"
      Top             =   570
      Width           =   1425
      _ExtentX        =   2514
      _ExtentY        =   1032
      BackColor       =   -2147483633
      MaskColor       =   8454143
      Caption         =   "&Reset Autogen"
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
      Height          =   435
      Left            =   5400
      TabIndex        =   1
      TabStop         =   0   'False
      Top             =   2130
      Visible         =   0   'False
      Width           =   435
      _ExtentX        =   767
      _ExtentY        =   767
   End
   Begin MwfCtl.MWCTL_ADC ADC 
      Height          =   555
      Left            =   7290
      TabIndex        =   2
      Top             =   2190
      Width           =   1950
      _ExtentX        =   3440
      _ExtentY        =   979
      AllowAdd        =   0   'False
      AllowFind       =   0   'False
      AllowSave       =   0   'False
      AllowDelete     =   0   'False
      BtnWidth        =   885
      BtnHeight       =   435
      BtnMaskColor    =   8454143
      FontSz          =   8.25
   End
   Begin VB.Frame FRA_KEYS 
      BorderStyle     =   0  'None
      Height          =   390
      Left            =   -36
      TabIndex        =   4
      Top             =   24
      Width           =   1350
   End
   Begin VB.Frame FRA_NKEYS 
      BorderStyle     =   0  'None
      Enabled         =   0   'False
      Height          =   390
      Left            =   -120
      TabIndex        =   3
      Top             =   510
      Width           =   1425
   End
End
Attribute VB_Name = "EmrFrmResAutogen"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim moCn As MDOConnection    '*** (Bef speed) bef 24/11/06

Private Sub CmdResAuto_Click()
  '*** Sometimes the user wants to reset the values in the Autogen table
      'This need arises when for example the user generates some bags (from 20 to 30)
      'and then deletes these bags (from 25 to 30). But now when he generates another bag,
      'all the numbering starts from 31 and not 25. It means that the numbers in between
      'are lost. So the need to Reset the Autogen
  '*** First show a Message Box to the user asking to make sure whether to reset the
      'records in the Autogen table.
      'Then all the records of the Autogen Table are deleted when the user says 'Yes'
      'New records are inserted into the table selecting the maximum value for a key
      'combination from the tables which makes use of Autogen
  '*** While resetting the values in the AutoGen table make sure that nobody will access
      'the records (insert, update or delete) in the autogen table during the period of
      'resetting. For this purpose set an exclusive lock on the Autogen table for the
      'transaction.
      
  Dim wsql As String
     
  ' ****** Sachin 2.13.0 [14-08-2006] ** Exclusive Access to the Database for Reset Autogen
  '*** Jay 3.1.1 [UlCoCd Added]
  '6.1-49 UlSysCd added
  If moCn.RecSeek("Select 'x' From UsrLogin Where (UlUsrCd <> '" + gs_UsrCd + "' and UlCoCd <> '" + gs_CoCd + "') " + _
    " or (UlUsrCd='" + gs_UsrCd + "' and UlCoCd='" + gs_CoCd + "' and UlSysCd<>'" + IIF(ctGenMenuUsr = "MIC", "EMR", "JEMR") + "')") Then _
    DispMsg "Cannot Reset Autogen As There are Other Users Accessing the Database", etError: Exit Sub
  ' ****** Sachin 2.13.0 [14-08-2006] ** Exclusive Access to the Database for Reset Autogen
  
  If MsgBox("User Should Have Exclusive Rights Of Database To Run This Option. Are U Sure U Want To Continue?", vbYesNo) = vbNo Then Exit Sub
  MWLib.BeginProcess Me, "Resetting Autogen Values"
  
  wsql = " Delete From AutoGen WITH (TABLOCKX) "
'Manoj 2.11.0 Corrections
'  wsql = wsql + " Insert Into AutoGen WITH (TABLOCKX) (ATABLE,AFLD,APREKEY,ANUM) " + _
'                " Select 'OrdMst', 'OmNo', OmCoCd+OmTc+OmYy+OmChr, max(OmNo) from OrdMst group by OmCoCd, OmTc, OmYy, OmChr union all " + _
'                " Select 'Txn', 'TNo', TCoCd+TTc+TYy+TChr, max(TNo) from Txn group by TCoCd, TTc, TYy, TChr union all " + _
'                " Select 'Fg', 'FgNo', FgCoCd+FgTc+FgYy+FgChr, max(FgNo) from Fg group by FgCoCd, FgTc, FgYy, FgChr union all " + _
'                " Select 'Fgd', 'FdKey', FdCoCd+FdYy, max(FdKey) from Fgd group by FdCoCd, FdYy union all " + _
'                " Select 'Fm', 'FmNo', FmCoCd+FmTc+FmYy+FmChr, max(FmNo) from Fm group by FmCoCd, FmTc, FmYy, FmChr union all " + _
'                " Select 'Fmd', 'FmdKey', FmdCoCd+FmdYy, max(FmdKey) from Fmd group by FmdCoCd, FmdYy union all " + _
'                " Select 'Bag', 'BNo', BCoCd+BYy+BChr, max(BNo) from Bag group by BCoCd, BYy, BChr union all " + _
'                " Select 'Txnd', 'TdKey', TdCoCd+TdYy, max(TdKey) from Txnd group by TdCoCd, TdYy union all " + _
'                " Select 'InvHd', 'InNo', InCoCd+InTc+InYy+InChr, max(InNo) from InvHd group by InCoCd, InTc, InYy, InChr Union All " + _
'                " Select 'TxndZ', 'TzKey', TzCoCd+TzYy, max(TzKey) from TxndZ group by TzCoCd, TzYy "
'                '''''" Select 'InvFgd', 'IfSrNo', IfTc+IfYy+IfChr+LTrim(IfNo), max(IfSrNo) from InvFgd group by IfTc, IfYy, IfChr, IfNo union all "
  ' Zubin 212 (TgNo for Tag added)
  
  '****** (Jen 2.13) added RetMem ******'
  '*** Jay 3.4[DP] Added DsgPrm
  ' ***** Manali 3.10.0 - 05/04/12 - MultiPrcQtMst added
  wsql = wsql + " Insert Into AutoGen WITH (TABLOCKX) (ATABLE,AFLD,APREKEY,ANUM) " + _
                " Select 'OrdMst', 'OmNo', OmCoCd+OmTc+OmYy+OmChr, max(OmNo) from OrdMst group by OmCoCd, OmTc, OmYy, OmChr union all " + _
                " Select 'Txn', 'TNo', TCoCd+TTc+TYy+TChr, max(TNo) from Txn group by TCoCd, TTc, TYy, TChr union all " + _
                " Select 'Fg', 'FgNo', FgCoCd+FgTc+FgYy+FgChr, max(FgNo) from Fg group by FgCoCd, FgTc, FgYy, FgChr union all " + _
                " Select 'Fgd', 'FdKey', FdCoCd+FdYy, max(FdKey) from Fgd group by FdCoCd, FdYy union all " + _
                " Select 'Fm', 'FmNo', FmCoCd+FmTc+FmYy+FmChr, max(FmNo) from Fm group by FmCoCd, FmTc, FmYy, FmChr union all " + _
                " Select 'Fmd', 'FmdKey', FmdCoCd+FmdYy, max(FmdKey) from Fmd group by FmdCoCd, FmdYy union all " + _
                " Select 'Bag', 'BNo', BCoCd+BYy+BChr, max(BNo) from Bag group by BCoCd, BYy, BChr union all " + _
                " Select 'Txnd', 'TdKey', TdCoCd+TdYy, max(TdKey) from Txnd group by TdCoCd, TdYy union all " + _
                " Select 'InvHd', 'InNo', InCoCd+InTc+InYy+InChr, max(InNo) from InvHd group by InCoCd, InTc, InYy, InChr Union All " + _
                " Select 'RetMem', 'ReNo', ReCoCd+ReYy+ReChr, max(ReNo) from RetMem group by ReCoCd, ReYy, ReChr Union All " + _
                " Select 'TxndZ', 'TzKey', TzCoCd+TzYy, max(TzKey) from TxndZ group by TzCoCd, TzYy Union All " + _
                " Select 'Tag', 'TgSr', TgCoCd+TgTc+TgYy+TgChr+LTrim(RTrim(Str(TgNo))), max(TgSr) from Tag group by TgCoCd, TgTc, TgYy, TgChr, TgNo Union All " + _
                " Select 'BgRecoMst', 'BmNo', BmCoCd+BmTc+BmYy+BmChr, max(BmNo) from BgRecoMst group by BmCoCd, BmTc, BmYy, BmChr Union All " + _
                " Select 'Tag', 'TgNo', TgCoCd+TgTc+TgYy+TgChr, max(TgNo) from Tag group by TgCoCd, TgTc, TgYy, TgChr Union All " + _
                " Select 'DsgPrm', 'DpCd', DpTyp, max(cast (DpCd as int)) from DsgPrm where DpTyp in ('SER', 'SET', 'VER') and DpDmCd='' group by DpTyp Union All " + _
                " Select 'MultiPrcQtMst', 'MqmNo', MqmCoCd+MqmTc+MqmYy+MqmChr, max(MqmNo) from MultiPrcQtMst Group by MqmCoCd, MqmTc, MqmYy, MqmChr Union All " + _
                " Select 'OrdDsg', 'OdDmStkNo',OdDmStkCoCd+OdDmStkYy+OdDmStkChr, max(OdDmStkNo) from OrdDsg where OdDmStkCoCd<>'' Group by OdDmStkCoCd, OdDmStkYy, OdDmStkChr "
                
'Manoj 2.11.0 Corrections
  
  moCn.Execute (wsql)
  
  If moCn.RecSeek("Select 'x' from Param where PTyp= 'SYS' and PMCd= 'XEMR' ") Then
    wsql = " Insert Into AutoGen WITH (TABLOCKX) (ATABLE,AFLD,APREKEY,ANUM) " + _
           " Select 'xTxn', 'xTNo', xTCoCd+xTTc+xTYy+xTChr, max(xTNo) from xTxn group by xTCoCd, xTTc, xTYy, xTChr Union all " + _
           " Select 'xTxn', 'xTKey', xTCoCd+xTYy, max(xTKey) from xTxn group by xTCoCd, xTYy Union all " + _
           " Select 'xTxnDsg', 'xTdKey', xTdCoCd+xTdYy, max(xTdKey) from xTxnDsg group by xTdCoCd, xTdYy "
    moCn.Execute (wsql)
  End If
  
  MWLib.EndProcess Me
  DispMsg "Reset AutoGen Values Complete", etInfo
End Sub

'*** Std Code not to be changed
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
Private Sub ADC_Load()
  Set moCn = adc.Connection    '*** (Bef speed) bef 24/11/06
End Sub
Private Sub Form_KeyUp(KeyCode As Integer, Shift As Integer)
  If adc.HotKeys(KeyCode, Shift) Then
    KeyCode = 0
    Shift = 0
  End If
End Sub


'*********** Note ******************


