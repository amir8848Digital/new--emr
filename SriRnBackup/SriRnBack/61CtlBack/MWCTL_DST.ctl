VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Begin VB.UserControl MWCTL_DST 
   BackStyle       =   0  'Transparent
   ClientHeight    =   750
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   6345
   ScaleHeight     =   750
   ScaleWidth      =   6345
   Begin VB.PictureBox PicBox_SPF 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   455
      Left            =   30
      ScaleHeight     =   420
      ScaleWidth      =   6255
      TabIndex        =   0
      Top             =   30
      Width           =   6285
      Begin VB.OptionButton OPT_SPF 
         Caption         =   "E&xport"
         Height          =   375
         Index           =   3
         Left            =   3810
         TabIndex        =   11
         Top             =   60
         Width           =   1065
      End
      Begin MwfCtl.MWCTL_BTN1 cmd_pfset 
         Height          =   345
         Left            =   1695
         TabIndex        =   3
         Top             =   60
         Width           =   1080
         _ExtentX        =   1905
         _ExtentY        =   609
         BackColor       =   -2147483633
         DisableColor    =   0
         Caption         =   "S&elect Printer"
         BeginProperty LblFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin VB.OptionButton OPT_SPF 
         Caption         =   "&Screen"
         Height          =   375
         Index           =   0
         Left            =   60
         TabIndex        =   1
         Top             =   45
         Value           =   -1  'True
         Width           =   885
      End
      Begin VB.OptionButton OPT_SPF 
         Caption         =   "&File"
         Height          =   375
         Index           =   2
         Left            =   2880
         TabIndex        =   4
         Top             =   45
         Width           =   1065
      End
      Begin VB.CheckBox chk_datarng 
         Caption         =   "Data&Range"
         Height          =   375
         Left            =   4860
         TabIndex        =   5
         Top             =   45
         Width           =   1335
      End
      Begin VB.OptionButton OPT_SPF 
         Caption         =   "&Print"
         Height          =   375
         Index           =   1
         Left            =   1020
         TabIndex        =   2
         Top             =   45
         Width           =   735
      End
   End
   Begin MSComDlg.CommonDialog cmg_prn 
      Left            =   780
      Top             =   450
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
      FontSize        =   1.73992e-39
      Max             =   999
      PrinterDefault  =   0   'False
   End
   Begin VB.PictureBox PicBox_PrnSel 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   525
      Left            =   360
      ScaleHeight     =   495
      ScaleWidth      =   5595
      TabIndex        =   6
      Top             =   0
      Visible         =   0   'False
      Width           =   5625
      Begin VB.CommandButton cmd_OK 
         Caption         =   "&OK"
         Height          =   315
         Left            =   4050
         Style           =   1  'Graphical
         TabIndex        =   8
         Top             =   90
         Width           =   645
      End
      Begin VB.CommandButton cmd_Cancel 
         Caption         =   "&Cancel"
         Height          =   315
         Left            =   4695
         Style           =   1  'Graphical
         TabIndex        =   9
         Top             =   90
         Width           =   645
      End
      Begin VB.ComboBox cmb_prns 
         Height          =   315
         Left            =   1080
         Style           =   2  'Dropdown List
         TabIndex        =   7
         Top             =   90
         Width           =   2865
      End
      Begin VB.Label LblPrn 
         BackStyle       =   0  'Transparent
         Caption         =   "Select Printer"
         Height          =   315
         Left            =   30
         TabIndex        =   10
         Top             =   135
         Width           =   1005
      End
   End
End
Attribute VB_Name = "MWCTL_DST"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
Dim mo_RepSource As CRAXDRT.Report
Dim mo_DataRng As New MWREP_DATARNG
Dim mo_SelPrn As Printer

Private Sub cmd_OK_Click()
    Dim wprn As Printer
    If cmb_prns.Text <> "" Then
      For Each wprn In Printers
        If wprn.DeviceName = cmb_prns Then
          Set mo_SelPrn = wprn
        End If
      Next wprn
      PicBox_SPF.Visible = True
      PicBox_PrnSel.Visible = False
      If Parent.Controls("FRA_REPOBJ").Visible = True Then
        OPT_SPF(1) = True
        Call ShowReport
      End If
    Else
      If cmb_prns.Text = "" Then DispMsg "Printer is Not Selected Or Not Installed.", etError
    End If
End Sub
Private Sub cmd_cancel_Click()
     PicBox_SPF.Visible = True
     PicBox_PrnSel.Visible = False
End Sub
Private Sub cmd_pfset_click()
    If OPT_SPF(1).Value = True Then
        Dim wprn As Printer
        cmb_prns.Clear
        For Each wprn In Printers
            cmb_prns.Additem wprn.DeviceName
        Next wprn
        If Not (Printer Is Nothing) And Printers.Count > 0 Then cmb_prns = Printer.DeviceName
        PicBox_PrnSel.Visible = True
        cmb_prns.SetFocus
        PicBox_SPF.Visible = False
        PicBox_PrnSel.ZOrder
     End If
End Sub
Public Sub ShowReport()
    Dim wl_PaperOrientation As Long, wl_PaperSize As Long
    On Error GoTo errh
    Set mo_RepSource = Parent.adc.RepSource
    If OPT_SPF(1) = True And mo_RepSource.ExportOptions.DiskFileName = "" Then
        If (mo_SelPrn Is Nothing) And (Not (Printer Is Nothing)) Then
            Set mo_SelPrn = Printer
        End If
        If mo_SelPrn Is Nothing Then err.Raise vbObjectError + 512, "", "No Printer chosen"
' ########################  Manoj #################### Date: 30/04/2004  ########################
' If Reoport is Already Open Then by Canging Printer, Print Command is issued, set Report again
        If Parent.Controls("FRA_REPOBJ").Visible = True Then
           mo_RepSource.DiscardSavedData
           If Parent.Controls("Adc").Connection.RepRes(mo_RepSource, Parent.Controls("Adc").RepRecSource) = 0 Then
               err.Raise vbObjectError + 512, "", "No Data Found for given scope"
           End If
           Parent.Controls("CRV_REP").ReportSource = mo_RepSource
        End If
' ###########################################  Manoj  ###########################################
        wl_PaperOrientation = mo_RepSource.PaperOrientation
        wl_PaperSize = mo_RepSource.PaperSize
        mo_RepSource.SelectPrinter mo_SelPrn.DriverName, mo_SelPrn.DeviceName, mo_SelPrn.Port
        mo_RepSource.PaperOrientation = wl_PaperOrientation
        mo_RepSource.PaperSize = wl_PaperSize
        mo_RepSource.PrintOut True
        If chk_datarng.Value = 1 Then
            Call FormatDataRng
            mo_DataRng.SelectPrinter mo_SelPrn.DriverName, mo_SelPrn.DeviceName, mo_SelPrn.Port
            mo_DataRng.PrintOut True
        End If
' ########################  Manoj #################### Date: 30/04/2004  ########################
'  If Reoport is Already Open Then by Canging Printer, Print Command is issued, set Report again
        If Parent.Controls("FRA_REPOBJ").Visible = True Then
           Parent.Controls("CRV_REP").ViewReport
        End If
' ###########################################  Manoj  ###########################################
    ElseIf OPT_SPF(2) = True Or mo_RepSource.ExportOptions.DiskFileName <> "" Then
       ' mo_RepSource.Export True
    With mo_RepSource.ExportOptions
        mo_RepSource.ExportOptions.FormatType = crEFTText
'        mo_RepSource.ExportOptions.DiskFileName = "C:\Output\Abcd.txt"
        mo_RepSource.ExportOptions.DestinationType = crEDTDiskFile
        If .DiskFileName = "" Then
            mo_RepSource.Export True
        Else
            mo_RepSource.Export False
        End If
    End With
        If chk_datarng.Value = 1 Then
            Call FormatDataRng
            mo_DataRng.Export True
        End If
    ElseIf OPT_SPF(3) = True Then
        WriteOutputToText
    End If
    Exit Sub
errh:
    Call Parent.adc.HandleError
End Sub
Public Sub WriteQueryToCSVFile(ByVal SourceQuery As String, ByVal SourceFileName As String)
  Dim fFile As Integer
  fFile = FreeFile

  Dim i As Long
  Open SourceFileName For Output As #fFile
  Print #fFile, SourceQuery
  Close #fFile

'Destroy the instance of the recordset from memory
End Sub

Private Sub WriteOutputToText()
  Dim mo_Rowset As MDORowSet
  Set mo_Rowset = Parent.adc.Connection.OpenRes(Parent.adc.RepRecSource)
  Dim strFileName, strFileNameSource As String
  strFileName = Parent.adc.RepExportFileDirectory + "\" + Parent.adc.MenuCd + "_" + Parent.adc.UsrCd + "_" + Format(DateTime.Now, "yyyy_mm_dd_hh_mm_ss") + ".txt"
  strFileNameSource = Parent.adc.RepExportFileDirectory + "\" + Parent.adc.MenuCd + "_" + Parent.adc.UsrCd + "_Src_" + Format(DateTime.Now, "yyyy_mm_dd_hh_mm_ss") + ".txt"
  
  mo_Rowset.WriteToCSVFile strFileName
  WriteQueryToCSVFile Parent.adc.RepRecSource, strFileNameSource
  MsgBox "Export completed. File Name: " + strFileName
End Sub
Public Property Get DstOpt() As en_RepDstOpt
    If OPT_SPF(0).Value = True Then
        DstOpt = repOnScreen
    ElseIf OPT_SPF(1).Value = True Then
        DstOpt = repInPrint
    ElseIf OPT_SPF(2).Value = True Then
        DstOpt = repInFile
    ElseIf OPT_SPF(3).Value = True Then
        DstOpt = repInExport
    End If
End Property
Public Property Get Enabled() As Boolean
    Enabled = UserControl.Enabled
End Property
Public Property Let Enabled(ByVal vNewValue As Boolean)
    UserControl.Enabled = vNewValue
    PicBox_SPF.Enabled = vNewValue
    OPT_SPF(0).Enabled = vNewValue
    OPT_SPF(1).Enabled = vNewValue
    OPT_SPF(2).Enabled = vNewValue
    cmd_pfset.Enabled = vNewValue
    chk_datarng.Enabled = vNewValue
End Property
'*** Jenny (13/06/05)
Public Property Let BckColor(ByVal vNewValue As OLE_COLOR)
    Dim wi_i As Integer
    PicBox_SPF.BackColor = vNewValue
    cmd_pfset.BackColor = vNewValue
    'cmd_pfset1.BackColor = vNewValue
    chk_datarng.BackColor = vNewValue
    For wi_i = 0 To OPT_SPF.Count - 1
      OPT_SPF(wi_i).BackColor = vNewValue
    Next wi_i
    
    PicBox_PrnSel.BackColor = vNewValue
    cmd_OK.BackColor = vNewValue
    cmd_Cancel.BackColor = vNewValue
    
    PropertyChanged "BckColor"
End Property
Public Property Get BckColor() As OLE_COLOR
    BckColor = cmd_pfset.BackColor
End Property
Public Property Let ForColor(ByVal vNewValue As OLE_COLOR)
    Dim wi_i As Integer
    cmd_pfset.ForeColor = vNewValue
    chk_datarng.ForeColor = vNewValue
    For wi_i = 0 To OPT_SPF.Count - 1
      OPT_SPF(wi_i).ForeColor = vNewValue
    Next wi_i
    
    LblPrn.ForeColor = vNewValue
    
    PropertyChanged "ForColor"
End Property
Public Property Get ForColor() As OLE_COLOR
    ForColor = cmd_pfset.ForeColor
End Property
Public Property Let ComboBckColor(ByVal vNewValue As OLE_COLOR)
    cmb_prns.BackColor = vNewValue
    PropertyChanged "ComboBckColor"
End Property
Public Property Get ComboBckColor() As OLE_COLOR
    ComboBckColor = cmb_prns.BackColor
End Property
Public Property Let ComboForColor(ByVal vNewValue As OLE_COLOR)
    cmb_prns.ForeColor = vNewValue
    PropertyChanged "ComboForColor"
End Property
Public Property Get ComboForColor() As OLE_COLOR
    ComboForColor = cmb_prns.ForeColor
End Property
'Public Property Get Pic() As Picture
'    Set Pic = PicBox_PrnSel.Picture
'End Property
'Public Property Let Pic(ByVal vNewValue As Picture)
'    'Set PicBox_SPF.Picture = vNewValue
'    Set PicBox_PrnSel.Picture = vNewValue
'    PropertyChanged "Pic"
'End Property
'Public Property Set Pic(ByVal vNewValue As Picture)
'    'Set PicBox_SPF.Picture = vNewValue
'    Set PicBox_PrnSel.Picture = vNewValue
'    PropertyChanged "Pic"
'End Property
Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
    BckColor = PropBag.ReadProperty("BckColor", &H80000005)
    ForColor = PropBag.ReadProperty("ForColor", &H80000008)
    ComboBckColor = PropBag.ReadProperty("ComboBckColor", &H80000005)
    ComboForColor = PropBag.ReadProperty("ComboForColor", &H80000008)
    'Pic = PropBag.ReadProperty("Pic", LoadPicture())
End Sub
Private Sub UserControl_WriteProperties(PropBag As PropertyBag)
    Call PropBag.WriteProperty("ComboBckColor", cmb_prns.BackColor, &H80000005)
    Call PropBag.WriteProperty("ComboForColor", cmb_prns.ForeColor, &H80000008)
    'Call PropBag.WriteProperty("Pic", PicBox_SPF.Picture, LoadPicture())
End Sub
'*** Jenny (13/06/05)



Private Sub FormatDataRng()
    On Error GoTo errh
    Dim wf_Over As Boolean, wf_Empty As Boolean
    Dim wc_ObjColl As New Collection
    Dim wi_CntLbl As Integer, wi_CntMed As Integer, wi_Top As Integer, wi_Gap As Integer, i As Integer, wo_Ctl As Control, wo_RepObj As Object, j As Integer
    Dim Fnt As StdFont
    i = 1
    wi_Gap = 0
    Do While True
      wf_Over = True
      wi_Top = 9999
      For Each wo_Ctl In Parent.Controls
        If TypeOf wo_Ctl Is Label Or TypeOf wo_Ctl Is MWCTL_MED Then
            If wo_Ctl.Tag <> "C" Then
              wf_Over = False
              If wi_Top > wo_Ctl.Top Then
                wi_Top = wo_Ctl.Top
              End If
            End If
        End If
      Next wo_Ctl
      If wf_Over Then
          Exit Do
      End If
        
      wf_Empty = True
      wi_CntLbl = 0
      wi_CntMed = 0
      Set wc_ObjColl = Nothing
      For Each wo_Ctl In Parent.Controls
              If TypeOf wo_Ctl Is Label Or TypeOf wo_Ctl Is MWCTL_MED Then
                If wo_Ctl.Top = wi_Top Then
                  wc_ObjColl.Add wo_Ctl
                  wo_Ctl.Tag = "C"
                  If TypeOf wo_Ctl Is Label Then
                    wi_CntLbl = wi_CntLbl + 1
                  Else
' ########################  Manoj #################### Date: 30/04/2004  ########################
                    If wo_Ctl.IdName <> "UOUSRCD" And wo_Ctl.IdName <> "UOMNUCD" Then
                      If wo_Ctl.Value <> "" Then
' ###########################################  Manoj  ###########################################
                        wf_Empty = False
                      End If
                      wi_CntMed = wi_CntMed + 1
                    End If
                  End If
                End If
              End If
      Next wo_Ctl
            
            
      If wi_CntMed = 0 Or Not wf_Empty Then
              For j = 1 To wc_ObjColl.Count
                  Set wo_RepObj = mo_DataRng.Sections("DETAIL").ReportObjects(i)
                  
' ########################  Manoj #################### Date: 30/04/2004  ########################
                  Set Fnt = New StdFont
                  Fnt.Name = "Arial"
                  Fnt.Size = 10
' ###########################################  Manoj  ###########################################

                  If TypeOf wc_ObjColl(j) Is Label Then
                    wo_RepObj.SetText wc_ObjColl(j).Caption
                    Fnt.Bold = True
                    i = i + 1
                  End If
' ########################  Manoj #################### Date: 30/04/2004  ########################
'                  If TypeOf wc_ObjColl(j) Is MWCTL_MED Then
'                    wo_RepObj.SetText wc_ObjColl(j).Value
'                  End If
                  If TypeOf wc_ObjColl(j) Is MWCTL_MED Then
                    If wc_ObjColl(j).IdName = "UOUSRCD" Then
                      mo_DataRng.TxtUsr.SetText wc_ObjColl(j).Value
                    ElseIf wc_ObjColl(j).IdName = "UOMNUCD" Then
                      mo_DataRng.TxtMnuCd.SetText wc_ObjColl(j).Value
                    Else
                      Fnt.Bold = False
                      wo_RepObj.SetText wc_ObjColl(j).Value
                      i = i + 1
                    End If
                  End If
                  Set wo_RepObj.Font = Fnt
' ###########################################  Manoj  ###########################################
                  wo_RepObj.Left = wc_ObjColl(j).Left + 100
                  wo_RepObj.Top = wc_ObjColl(j).Top - wi_Gap
                  wo_RepObj.Height = wc_ObjColl(j).Height
                  wo_RepObj.Width = wc_ObjColl(j).Width + 100
                Next j
      Else
        wi_Gap = wi_Gap + wc_ObjColl(1).Height
      End If
  Loop
  mo_DataRng.TxtGrp1.SetText Parent.Controls("GRP_REP").Value(gltGroup1)
  mo_DataRng.TxtGrp2.SetText Parent.Controls("GRP_REP").Value(gltGroup2)
  mo_DataRng.TxtGrp3.SetText Parent.Controls("GRP_REP").Value(gltGroup3)
  mo_DataRng.TxtHdg.SetText Parent.Caption + " Data Range Options "
  Exit Sub
errh:
    If err.Number = 423 Then
        Exit Sub
    Else
        MsgBox err.Description
    End If

End Sub

