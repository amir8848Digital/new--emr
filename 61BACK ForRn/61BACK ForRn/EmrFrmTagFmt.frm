VERSION 5.00
Object = "{6A2CE895-C1AC-48AF-B558-FB94BBE12499}#1.0#0"; "mwfctl.ocx"
Begin VB.Form EmrFrmTagFmt 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Tag Format"
   ClientHeight    =   9810
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   11505
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MDIChild        =   -1  'True
   ScaleHeight     =   9810
   ScaleMode       =   0  'User
   ScaleWidth      =   11505
   WindowState     =   2  'Maximized
   Begin MwfCtl.MWCTL_HLP MWCTL_HLP1 
      Height          =   405
      Left            =   6300
      TabIndex        =   13
      TabStop         =   0   'False
      Top             =   9540
      Visible         =   0   'False
      Width           =   540
      _ExtentX        =   953
      _ExtentY        =   714
   End
   Begin MwfCtl.MWCTL_ADC ADC 
      CausesValidation=   0   'False
      Height          =   555
      Left            =   60
      TabIndex        =   12
      Top             =   9300
      Width           =   11385
      _ExtentX        =   20082
      _ExtentY        =   979
      BtnWidth        =   885
      BtnHeight       =   435
      BtnMaskColor    =   8454143
      FontSz          =   8.25
   End
   Begin VB.Frame FRA_KEYS 
      BorderStyle     =   0  'None
      Height          =   300
      Left            =   0
      TabIndex        =   16
      Top             =   45
      Width           =   3330
      Begin MwfCtl.MWCTL_MED ATXT 
         Height          =   285
         Index           =   1
         Left            =   1305
         TabIndex        =   1
         ToolTipText     =   "Enter Tag Code"
         Top             =   15
         Width           =   1110
         _ExtentX        =   1958
         _ExtentY        =   503
         MaxLength       =   5
         DataType        =   4
         DataField       =   "TgfCd"
         IdName          =   "TGFCD"
      End
      Begin MwfCtl.MWCTL_MED ATXT 
         Height          =   285
         Index           =   6
         Left            =   2460
         TabIndex        =   0
         Top             =   30
         Visible         =   0   'False
         Width           =   120
         _ExtentX        =   212
         _ExtentY        =   503
         Enabled         =   0   'False
         MaxLength       =   3
         DataType        =   4
         DataField       =   "TgfCoCd"
         IdName          =   "TGFCOCD"
      End
      Begin VB.Label ALBL 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Tag Code"
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
         Left            =   60
         TabIndex        =   15
         Top             =   30
         Width           =   915
      End
   End
   Begin VB.Frame FRA_NKEYS 
      BorderStyle     =   0  'None
      Enabled         =   0   'False
      Height          =   9390
      Left            =   0
      TabIndex        =   14
      Top             =   45
      Width           =   12285
      Begin VB.Frame FraNKeyAll 
         BorderStyle     =   0  'None
         Height          =   9315
         Left            =   0
         TabIndex        =   17
         Top             =   0
         Width           =   11820
         Begin VB.TextBox TxtHlp 
            BackColor       =   &H80000018&
            BeginProperty Font 
               Name            =   "Courier"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000017&
            Height          =   3885
            Left            =   60
            Locked          =   -1  'True
            MultiLine       =   -1  'True
            ScrollBars      =   2  'Vertical
            TabIndex        =   23
            TabStop         =   0   'False
            Text            =   "EmrFrmTagFmt.frx":0000
            ToolTipText     =   "Help Text(Read Only)"
            Top             =   5280
            Width           =   11385
         End
         Begin MwfCtl.MWCTL_MED ATXT 
            Height          =   285
            Index           =   3
            Left            =   1305
            TabIndex        =   3
            ToolTipText     =   "Enter Starting String"
            Top             =   300
            Width           =   1725
            _ExtentX        =   3043
            _ExtentY        =   503
            MaxLength       =   8
            DataType        =   4
            DataField       =   "TgfStartStr"
            IdName          =   "TGFSTARTSTR"
         End
         Begin MwfCtl.MWCTL_MED ATXT 
            Height          =   285
            Index           =   0
            Left            =   4575
            TabIndex        =   4
            ToolTipText     =   "Enter Darkness String"
            Top             =   300
            Width           =   1725
            _ExtentX        =   3043
            _ExtentY        =   503
            MaxLength       =   8
            DataType        =   4
            DataField       =   "TgfDarkness"
            IdName          =   "TGFDARKNESS"
         End
         Begin MwfCtl.MWCTL_MED ATXT 
            Height          =   285
            Index           =   2
            Left            =   1305
            TabIndex        =   5
            ToolTipText     =   "Enter Heat String"
            Top             =   585
            Width           =   1725
            _ExtentX        =   3043
            _ExtentY        =   503
            MaxLength       =   8
            DataType        =   4
            DataField       =   "TgfHeat"
            IdName          =   "TGFHEAT"
         End
         Begin MwfCtl.MWCTL_MED ATXT 
            Height          =   285
            Index           =   4
            Left            =   4575
            TabIndex        =   6
            ToolTipText     =   "Enter Quantity String"
            Top             =   585
            Width           =   1725
            _ExtentX        =   3043
            _ExtentY        =   503
            MaxLength       =   8
            DataType        =   4
            DataField       =   "TgfQtyStr"
            IdName          =   "TGFQTYSTR"
         End
         Begin MwfCtl.MWCTL_MED ATXT 
            Height          =   285
            Index           =   5
            Left            =   1305
            TabIndex        =   7
            ToolTipText     =   "Enter End String"
            Top             =   870
            Width           =   1725
            _ExtentX        =   3043
            _ExtentY        =   503
            MaxLength       =   8
            DataType        =   4
            DataField       =   "TgfEndStr"
            IdName          =   "TGFENDSTR"
         End
         Begin MwfCtl.MWCTL_FLX GrdFmtSr 
            Height          =   3660
            Left            =   60
            TabIndex        =   11
            Top             =   1560
            Width           =   11385
            _ExtentX        =   20082
            _ExtentY        =   6456
            Cols            =   4
            colname1        =   "TGFSSR"
            heading1        =   "Sr. "
            datafld1        =   "TgfsSr"
            datatype1       =   1
            mask1           =   "##"
            maxlength1      =   2
            tooltiptext1    =   "Serial Number"
            colname2        =   "TGFSPRINTSTR"
            heading2        =   "Printing String                                                "
            datafld2        =   "TgfsPrintStr"
            maxlength2      =   30
            tooltiptext2    =   "Enter Printing String"
            colname3        =   "TGFSFIELD"
            heading3        =   $"EmrFrmTagFmt.frx":0006
            datafld3        =   "TgfsField"
            maxlength3      =   255
            tooltiptext3    =   "Enter Printing Field"
         End
         Begin MwfCtl.MWCTL_MED ATXT 
            Height          =   285
            Index           =   7
            Left            =   4575
            TabIndex        =   2
            ToolTipText     =   "Enter Description for Code"
            Top             =   0
            Width           =   6855
            _ExtentX        =   12091
            _ExtentY        =   503
            MaxLength       =   60
            DataField       =   "TgfDesc"
            IdName          =   "TGFDESC"
         End
         Begin MwfCtl.MWCTL_MED ATXT 
            Height          =   285
            Index           =   8
            Left            =   1305
            TabIndex        =   8
            ToolTipText     =   "Enter End String"
            Top             =   1170
            Width           =   1725
            _ExtentX        =   3043
            _ExtentY        =   503
            MaxLength       =   8
            DataType        =   4
            DataField       =   "TgfFldOrg"
            IdName          =   "TgfFldOrg"
         End
         Begin MwfCtl.MWCTL_MED ATXT 
            Height          =   285
            Index           =   9
            Left            =   4575
            TabIndex        =   9
            ToolTipText     =   "Enter End String"
            Top             =   1170
            Width           =   1725
            _ExtentX        =   3043
            _ExtentY        =   503
            MaxLength       =   8
            DataType        =   4
            DataField       =   "TgfFldData"
            IdName          =   "TgfFldData"
         End
         Begin MwfCtl.MWCTL_MED ATXT 
            Height          =   285
            Index           =   10
            Left            =   7935
            TabIndex        =   10
            ToolTipText     =   "Enter End String"
            Top             =   1170
            Width           =   1725
            _ExtentX        =   3043
            _ExtentY        =   503
            MaxLength       =   8
            DataType        =   4
            DataField       =   "TgfFldEnd"
            IdName          =   "TgfFldEnd"
         End
         Begin VB.Label ALBL 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Fld End Str"
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
            Left            =   6690
            TabIndex        =   27
            Top             =   1170
            Width           =   1050
         End
         Begin VB.Label ALBL 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Fld Data Str"
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
            Index           =   8
            Left            =   3330
            TabIndex        =   26
            Top             =   1170
            Width           =   1125
         End
         Begin VB.Label ALBL 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Fld Org Str"
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
            Index           =   7
            Left            =   60
            TabIndex        =   25
            Top             =   1170
            Width           =   1035
         End
         Begin VB.Label ALBL 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Desc"
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
            Index           =   6
            Left            =   3510
            TabIndex        =   24
            Top             =   15
            Width           =   450
         End
         Begin VB.Label ALBL 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "End String"
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
            Left            =   60
            TabIndex        =   22
            Top             =   870
            Width           =   990
         End
         Begin VB.Label ALBL 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Qty String"
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
            Left            =   3510
            TabIndex        =   21
            Top             =   585
            Width           =   945
         End
         Begin VB.Label ALBL 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Heat String"
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
            Left            =   60
            TabIndex        =   20
            Top             =   585
            Width           =   1065
         End
         Begin VB.Label ALBL 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Darkness"
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
            Left            =   3510
            TabIndex        =   19
            Top             =   300
            Width           =   855
         End
         Begin VB.Label ALBL 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Start String"
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
            Left            =   60
            TabIndex        =   18
            Top             =   300
            Width           =   1080
         End
      End
   End
End
Attribute VB_Name = "EmrFrmTagFmt"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'----   General Comments on the Entry Form  -------'
' ########################  Manoj #### Ver: 2.0.6 #### Date: 04/03/2003  ########################
' #                                                                                             #
' # In this Entry Form Format for Tag Printing is Entered.                                      #
' # In ColValid of Grid then Tag Fields are Checked for Syntax Errors & RunTime Errors          #
' # Which Evaluate the TagFileds with Default Values                                            #
' # For Synmtax Check & Invalid Tag Field Error Occurs, For Runtime Errors which may fire are   #
' # Trapped & Warning is Given                                                                  #
' #                                                                                             #
' # Special Help is Given for this Form in TextBox Placed Below Grid                            #
' # Which will Help to Set The Format of Printing String                                        #
' #                                                                                             #
' ###########################################  Manoj  ###########################################
Option Explicit
Dim moCn As MwfLib.MDOConnection   '*** (Bef speed) bef 24/11/06
Dim ScTgFmt As Object                       'Scripting Object
Dim mColTgFmtFlds As New Collection         'TagField Collection

Private Sub ADC_KeyWhen()
  adc("TgfCoCd") = gs_CoCd
End Sub

Private Sub ADC_Load()
  '*** Set the First Key Control to Customer Code
  '*** Set the First Non Key Control to Customer Name
  '*** Set the Previous and Next Control Properties for ADC
  ' Set GrdFmtSr as Child of ADC
  '*** Set the Previous and Next Control Properties for Grid GrdFmtSr
  ' Set Hotkey to Grid AS "A"
  ' Initialize Scripting Object
  
  Set moCn = adc.Connection    '*** (Bef speed) bef 24/11/06
  Set adc.FirKeyCtl = adc("TgfCd")
  Set adc.FirNKeyCtl = adc("TgfDesc")
  
  adc.Child = "GrdFmtSr"
  Set GrdFmtSr.PrevCtl = adc("TgfEndStr")
  Set GrdFmtSr.NextCtl = adc
  GrdFmtSr.HotKey = "A"
  
  Set ScTgFmt = CreateObject("MSScriptControl.ScriptControl")
  ScTgFmt.Language = "VBScript"
  ScTgFmt.Timeout = -1
  SetScriptCtl ScTgFmt, mColTgFmtFlds
  
  '****** Sachin 2.14.0 - [24-11-2006] - [Multi-Company] ******
  Me.Caption = GetFrmCaption(gs_CoCd, "Tag Format")
  '****** Sachin 2.14.0 - [24-11-2006] - [Multi-Company] ******
End Sub

Private Sub ADC_NonKeyWhen(pr_Cancel As Boolean, pr_ErrMsg As String)
  '*** Jay 3.01 *** [Tag Entry Lock for DFDMX, DFZEB]
  If adc("TgfCd") = "DFDMX" Or adc("TgfCd") = "DFZEB" Then
    GrdFmtSr.AllowAdd = False
    GrdFmtSr.AllowDelete = False
  Else
    GrdFmtSr.AllowAdd = True
    GrdFmtSr.AllowDelete = True
  End If
  '*** Jay 3.01 *** [Tag Entry Lock for DFDMX, DFZEB]
End Sub

Private Sub ATXT_FldWhen(Index As Integer, ByVal IdName As String, Cancel As Boolean, ErrMsg As String)
  Call SetHelpText(IdName)
  '*** Jay 3.01 *** [Tag Entry Lock for DFDMX, DFZEB]
  If adc("TgfCd") = "DFDMX" Or adc("TgfCd") = "DFZEB" Then Cancel = True: ErrMsg = "Can Not Change Default Settings": Exit Sub
  '*** Jay 3.01 *** [Tag Entry Lock for DFDMX, DFZEB]
  '*** Set Helps For Each Field
  Select Case UCase(IdName)
    Case Is = UCase("TgfCd")
      Call HlpList.TgfCd(adc("TgfCoCd"))
  End Select
End Sub
Private Sub atxt_FldChange(Index As Integer, ByVal IdName As String, Cancel As Boolean, ErrMsg As String, ByVal pv_NewValue As Variant)
  '*** Set Validations For Each Text Field
  
  Select Case UCase(IdName)
    Case Is = UCase("TgfCd")
      If Trim(pv_NewValue) = "" Then
        Cancel = True: ErrMsg = "Tag Format Code Cannot Be Blank": Exit Sub
      End If
    Case Is = UCase("TgfStartStr")
      If Trim(pv_NewValue) = "" Then
        Cancel = True: ErrMsg = "Tag Start String Cannot Be Blank": Exit Sub
      End If
'    Case Is = UCase("TgfDarkness")
'      If Trim(pv_NewValue) = "" Then
'        Cancel = True: ErrMsg = "Tag Darkness String Cannot Be Blank": Exit Sub
'      End If
'    Case Is = UCase("TgfHeat")
'      If Trim(pv_NewValue) = "" Then
'        Cancel = True: ErrMsg = "Tag Heat String Cannot Be Blank": Exit Sub
'      End If
    Case Is = UCase("TgfQtyStr")
      If Trim(pv_NewValue) = "" Then
        Cancel = True: ErrMsg = "Tag Quantity String Cannot Be Blank": Exit Sub
      End If
    Case Is = UCase("TgfEndStr")
      If Trim(pv_NewValue) = "" Then
        Cancel = True: ErrMsg = "Tag End String Cannot Be Blank": Exit Sub
      End If
  End Select
End Sub

Private Sub GrdFmtSr_ColChange(ByVal RowNum As Integer, ByVal ColName As String, ByVal NewValue As Variant, Cancel As Boolean, ErrMsg As String)
  '*** Set Validations For Each Column of Grid
  Dim wstVal As String          ', wiStStrt As Integer, wiStEnd As Integer
  Dim wstValSplit() As String, i As Integer, wTemp As Variant
  Dim wStCheck As String
  Select Case UCase(ColName)
    Case Is = UCase("TgfsSr")
      If Trim(NewValue) = "" Then
        Cancel = True: ErrMsg = "Tag Format Serial No. Cannot Be Blank": Exit Sub
      End If
    Case Is = UCase("TgfsPrintStr")
      If Trim(NewValue) = "" Then
        Cancel = True: ErrMsg = "Tag Print String Cannot Be Blank": Exit Sub
      End If
    Case Is = UCase("TgfsField")
      If Trim(NewValue) = "" Then
        Cancel = True: ErrMsg = "Tag Field Cannot Be Blank": Exit Sub
      Else
        wstVal = NewValue
        wstValSplit = Split(wstVal, "|")
        If UBound(wstValSplit) >= 0 Then
          For i = 0 To UBound(wstValSplit)
            If (i Mod 2) = 1 Then
              wStCheck = ""
            ' ########################  Manoj #### Ver: 2.0.6 #### Date: 25/12/2003  ########################
            ' # The String Value is Check For Syntax & Invalid TagFields                                    #
            ' ###########################################  Manoj  ###########################################
              If mEvalChk(wStCheck, ScTgFmt, mColTgFmtFlds, RemoveWhiteSpaces(wstValSplit(i), True)) = False Then
                Cancel = True: ErrMsg = wStCheck: Exit Sub
              ElseIf wStCheck <> "" Then
                Call DispMsg("Warning : " & wStCheck, etWarning)
              End If
            End If
          Next
        End If
      End If
  End Select
End Sub

Private Sub GrdFmtSr_ColWhen(ByVal RowNum As Integer, ByVal ColName As String, Cancel As Boolean, ErrMsg As String, ByVal Action As MwfCtl.en_fgAction)
  '*** Jay 3.01 *** [Tag Entry Lock for DFDMX, DFZEB]
  If adc("TgfCd") = "DFDMX" Or adc("TgfCd") = "DFZEB" Then Cancel = True: ErrMsg = "Can Not Change Default Settings": Exit Sub
  '*** Jay 3.01 *** [Tag Entry Lock for DFDMX, DFZEB]
  Call SetHelpText(ColName)
End Sub

Private Sub GrdFmtSr_RowWhen(ByVal RowNum As Integer)
  '*** If GrdDsgRm.IsNew Then DrSr = Max + 1
  With GrdFmtSr
    If .Value(RowNum, "TgfsSr") = 0 Then _
       .Value(RowNum, "TgfsSr") = .MaxVal("TgfsSr") + 1
  End With
End Sub

Private Sub GrdFmtSr_InitKey(ByVal RowNum As Integer)
  '*** Initialise Key Fields
  With GrdFmtSr
      .Store "TgfsCoCd", adc("TgfCoCd")
      .Store "TgfsCd", adc("TgfCd")
  End With
End Sub

Private Sub ADC_SetRecSource()
  '*** Set the RecSource for ADC  (Base Table is CustMst)
  adc.RecSource = " Select * from TagFmt " + _
                  " Where TgfCoCd ='" + adc("TgfCoCd") + "' And TgfCd ='" + adc("TgfCd") + "'"
End Sub

Private Sub GrdFmtSr_SetRecSource()
  GrdFmtSr.RecSource = " Select * from TagFmtSr " + _
                  " Where TgfsCoCd ='" + adc("TgfCoCd") + "' And TgfsCd ='" + adc("TgfCd") + "'"
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
' ########################  Manoj #### Ver: 2.0.6 #### Date: 25/12/2003  ########################
  Call SetTagFldColl(mColTgFmtFlds)
' ###########################################  Manoj  ###########################################
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
  Set mColTgFmtFlds = Nothing

  Set ScTgFmt = Nothing
End Sub
Private Sub GrdFmtSr_Validate(Cancel As Boolean)
  Cancel = GrdFmtSr.Validate
End Sub
Private Sub SetHelpText(IdName)
' ########################  Manoj #### Ver: 2.0.6 #### Date: 25/12/2003  ########################
' # This Function Sets Help For Each Field & Example for same.                                  #
' ###########################################  Manoj  ###########################################
  TxtHlp.Visible = False
  Select Case UCase$(IdName)
    Case Is = UCase$("TgfStartStr")
      TxtHlp.text = "Enter Starting Control Chracter" + vbCrLf + _
                    "This is Start for Printing of Tag." + vbCrLf + _
                    "  Eg. (Ctrl+B)L Or ~L"
    Case Is = UCase$("TgfDarkness")
      TxtHlp.text = "Enter Darkness Control Chracter along with Value of Darkness in Two Digits" + vbCrLf + _
                    "This will affect to Darkness of Printing." + vbCrLf + _
                    "  Eg. D11"
    Case Is = UCase$("TgfHeat")
      TxtHlp.text = "Enter Heat Control Chracter along with Value of Heat in Two Digits" + vbCrLf + _
                    "This will affect to Heat applying from Printer." + vbCrLf + _
                    "  Eg. H15"
    Case Is = UCase$("TgfQtyStr")
      TxtHlp.text = "Enter Quantity Control Chracter, " + vbCrLf + _
                    "This will define No of Times Tag goig to be Print." + vbCrLf + _
                    "  Eg. Q"
    Case Is = UCase$("TgfEndStr")
      TxtHlp.text = "Enter End Control Chracter" + vbCrLf + _
                    "This is End for Printing of Tag." + vbCrLf + _
                    "  Eg. E"
    Case Is = UCase$("TgfsPrintStr")
      TxtHlp.text = "111100100370003" + vbCrLf + _
                    "Explanation of the above line:" + vbCrLf + vbCrLf + _
                    "1          1          11         001        0037       0003" + vbCrLf + _
                    "Rotation   Font       Ratio      Height     Top        Left" + vbCrLf + _
                    "                  Width To Height" + vbCrLf + _
                    "Please note that the valid ranges for each characteristic can vary from " + vbCrLf + _
                    "printer to printer.  " + vbCrLf + vbCrLf + _
                    "The first character '1' signifies the degree of rotation of the characters " + vbCrLf + _
                    "to be printed. The valid values are 1, 2, 3, 4. " + vbCrLf + vbCrLf + _
                    "The second character '1' signifies the font size to be printed. The valid " + vbCrLf + _
                    "values are 1 to 9 and 'E'. " + vbCrLf + _
                    "If 'E' or any other character then it means that the content of the line has " + vbCrLf + _
                    "to be printed as a bar-code. For understanding of valid bar codes refer manual" + vbCrLf + _
                    " of the printer." + vbCrLf + vbCrLf + _
                    "When Not ‘E’" + vbCrLf + "The third and the fourth characters '1', '1' signifies the ratio between " + vbCrLf + _
                    "the width and height. It is normally kept 1:1. " + vbCrLf + vbCrLf + _
                    "The 5th to 7th characters signify the height of the content to be printed. " + vbCrLf + vbCrLf + _
                    "The 8th to 11th characters signify the start of the printing position from " + vbCrLf + _
                    "the bottom of the tag (row). (10 rows equals 1cm.)" + vbCrLf + vbCrLf + _
                    "The 12th to 15th characters signify the start of the printing position " + vbCrLf + _
                    "from the left of the tag (column). (10 means 1 cm, to be confirmed)" + _
                    ""
    Case Is = UCase$("TgfsField")
      TxtHlp.text = "This String is Split using ""|"" Character. Every Alternate String is Evaluated" + vbCrLf + _
                    "First String Always Print as It Is. Next String after ""|"" is Evaluated." + vbCrLf + _
                    "Eg. " + vbCrLf + "  ""Bag=|Fld(""BagNo"")"":->""Bag=04/REG/1"", ""Fld(""BagNo"")"":->""Fld(""BagNo"")""" + vbCrLf + _
                    "  ""|Fld(""BagNo"")"":->""04/REG/1"",         ""G.Wt.=|Fld(""BagGrsWt"")|Gms"":->""G.Wt.=5.953Gms""" + vbCrLf + _
                    "" + vbCrLf + "You Can Use Arithmatic Operations on Numeric Values, String Concatenation," + vbCrLf + _
                    "and Some Functions which can be run in vbSript Like Round(),Date(),etc." + vbCrLf + _
                    "" + vbCrLf + "Some Fileds are also Available to Use with Evaluating their Values." + vbCrLf + _
                    "These Fields are Listed in Microway Parameter with PTYP=""TAGFLDS""" + vbCrLf + _
                    "These Fields can be accessed Like Fld(""FilldName"")' Or ""Fld.Item(""FieldName"")'" + vbCrLf + _
                    "Eg." + vbCrLf + "  BChr:->Fld(""BChr""), BNo:->Fld(""BNo""), BagNo:->Fld(""BagNo"")" + vbCrLf + _
                    "" + vbCrLf + "Some Functions which are Not Available in VbScript are available here" + vbCrLf + _
                    "but those has to be called as 'Fun.FunctionName(Parameters)'" + vbCrLf + "Avalable Functions are as below" + vbCrLf + vbCrLf + _
                    "RndOff(Value, No of Digits, with Value, Type)" + vbCrLf + "  Value :-> which has to be rounded." + vbCrLf + _
                    "  No of Digits :-> Value rounded to these Many Digits " + vbCrLf + "  with Value :-> Value rounded to Nearest Multiple of this Value" + vbCrLf + _
                    "  Type :-> Which Type of Rounding to be Done." + vbCrLf + "    1 :-> Simple Rounding" + vbCrLf + _
                    "    2 :-> Rounding Up" + vbCrLf + "    3 :-> Rounding Down" + vbCrLf + "  Eg." + vbCrLf + _
                    "    Fun.RndOff(234.5678):->235,         Fun.RndOff(234.5678,3):->234.568" + vbCrLf + _
                    "    Fun.RndOff(234.5678,3,5):->234.570, Fun.RndOff(234.5678,3,5,3):->234.565" + vbCrLf + _
                    "" + vbCrLf + "StrFmt(Value,Format)" + vbCrLf + "  Value :-> Value to be Formated" + vbCrLf + _
                    "  Format:-> In Which Format Value has to be Return." + vbCrLf + "  Eg." + vbCrLf + _
                    "    StrFmt(123.456,""0.0000""):->""123"".4560, StrFmt(.456,""0.00""):->""0.46""" + vbCrLf + _
                    "    StrFmt(Date(),""dd/mm/yyyy""):->""01/01/2004""" + vbCrLf + vbCrLf + _
                    "IIF(Cond, TrueVal, FalseVal)" + vbCrLf + "  Cond :-> Conition to be Check" + vbCrLf + _
                    "  TrueVal :-> If Cond is True Then this Value Return" + vbCrLf + "  FalseVal :-> If Cond is False Then this Value Return" + _
                    vbCrLf + "  Eg." + vbCrLf + "    IIF(1>=2,1,2):-> 2, IIF(1<=2,1,2):-> 1"
    Case Else
      Exit Sub
  End Select
  TxtHlp.Visible = True
End Sub

Private Sub TxtHlp_GotFocus()
  DispMsg "This is Read Only Help Text.", etInfo
End Sub
