VERSION 5.00
Begin VB.UserControl MWCTL_GRP 
   BackStyle       =   0  'Transparent
   ClientHeight    =   735
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   5355
   LockControls    =   -1  'True
   ScaleHeight     =   735
   ScaleWidth      =   5355
   Begin VB.ComboBox ACMB 
      Height          =   315
      Index           =   4
      Left            =   1755
      Sorted          =   -1  'True
      Style           =   2  'Dropdown List
      TabIndex        =   4
      Top             =   360
      Width           =   1755
   End
   Begin VB.ComboBox ACMB 
      Height          =   315
      Index           =   5
      Left            =   3510
      Sorted          =   -1  'True
      Style           =   2  'Dropdown List
      TabIndex        =   5
      Top             =   360
      Width           =   1755
   End
   Begin VB.ComboBox ACMB 
      Height          =   315
      Index           =   3
      Left            =   0
      Sorted          =   -1  'True
      Style           =   2  'Dropdown List
      TabIndex        =   3
      Top             =   360
      Width           =   1755
   End
   Begin VB.ComboBox ACMB 
      Height          =   315
      Index           =   2
      Left            =   3510
      Sorted          =   -1  'True
      Style           =   2  'Dropdown List
      TabIndex        =   2
      Top             =   30
      Width           =   1755
   End
   Begin VB.ComboBox ACMB 
      Height          =   315
      Index           =   1
      Left            =   1755
      Sorted          =   -1  'True
      Style           =   2  'Dropdown List
      TabIndex        =   1
      Top             =   30
      Width           =   1755
   End
   Begin VB.ComboBox ACMB 
      Height          =   315
      Index           =   0
      Left            =   0
      Sorted          =   -1  'True
      Style           =   2  'Dropdown List
      TabIndex        =   0
      Top             =   30
      Width           =   1755
   End
End
Attribute VB_Name = "MWCTL_GRP"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
Dim mc_Groups As New Collection
Dim ms_GrpFld1 As String, ms_GrpFld2 As String, ms_GrpFld3 As String
Dim mo_RepSource As CRAXDRT.Report
Dim mi_supwdth As Integer

'*** Jenny (13/06/05)
'Dim mo_BckColor As OLE_COLOR
'Dim mo_ForColor As OLE_COLOR
'*** Jenny (13/06/05)

Const ctNone = "(None)"
Const ctDetail = "(Detail)"
Const ctGrand = "(Grand)"
Public Sub Add(ByVal DispStrg As String, ByVal DBFldNm As Variant, _
               ByVal DispFldNm As Variant, _
               ByVal RepFldNm As String, ByVal HdFldNm As String, _
               ByVal InvSubGrpLst As String, _
               ByVal SortLst As Variant, ByVal SortFldLst As Variant, _
               ByVal Desc As Variant)
                   
    Dim wv_ArrObj As Variant
    If Not IsArray(DBFldNm) Then
        If DBFldNm = "" Then DBFldNm = "' '"
    End If
    If Not IsArray(DispFldNm) Then
        If DispFldNm = "" Then DispFldNm = DBFldNm
    End If
    If Not IsArray(Desc) Then
        '(Bef 14/07/05) If Desc = "" Then Desc = "''"
        If Desc = "" Then Desc = "' '"
    End If
    If Not IsArray(DBFldNm) Then DBFldNm = Array(DBFldNm)
    If Not IsArray(DispFldNm) Then DispFldNm = Array(DispFldNm)
    If Not IsArray(Desc) Then Desc = Array(Desc)
    If Not IsArray(SortLst) Then SortLst = Array()
    If Not IsArray(SortFldLst) Then SortFldLst = Array()
    wv_ArrObj = Array(DispStrg, DBFldNm, DispFldNm, RepFldNm, HdFldNm, InvSubGrpLst, SortLst, SortFldLst, Desc)
    mc_Groups.Add wv_ArrObj, DispStrg
End Sub
Private Sub ACMB_GotFocus(Index As Integer)
    InitCmb (Index)
End Sub
Private Sub InitCmb(ByVal Index As Integer, Optional ByVal ReqVal As Variant)
    Dim ws_cVal As String, ws_InvSubGrp As String, wf_cVal As Boolean, ws_AddStrg As String, wv_rArr As Variant, i As Integer
    ws_cVal = IIf(IsMissing(ReqVal), ACMB(Index), ReqVal)
    ACMB(Index).Clear
    Select Case Index
    Case Is = 0, 1, 2
    ws_InvSubGrp = GetInvSubGrp(Index)
    wf_cVal = False
    For i = 1 To mc_Groups.Count
        ws_AddStrg = mc_Groups(i)(0)
        If InStr(1, ws_InvSubGrp, "," + ws_AddStrg + ",") = 0 Then
            If ws_AddStrg <> "" Then
                ACMB(Index).Additem ws_AddStrg
                If ws_cVal = ws_AddStrg Then wf_cVal = True
            End If
        End If
    Next i
    Case Is = 3, 4, 5
    wv_rArr = mc_Groups(ACMB(Index - 3))(6)
    For i = 0 To UBound(wv_rArr)
        If wv_rArr(i) <> "" Then
            ACMB(Index).Additem wv_rArr(i)
            If ws_cVal = wv_rArr(i) Then wf_cVal = True
        End If
    Next i
    ACMB(Index).Additem ctNone
    End Select
    If Not wf_cVal Then
        ACMB(Index) = GetDefaVal(Index)
    Else
        ACMB(Index) = ws_cVal
    End If
End Sub
Private Function GetDefaVal(ByVal wIndex As Long) As String
    GetDefaVal = IIf(wIndex = 0, ctDetail, ctNone)
End Function
Private Function GetInvSubGrp(ByVal wIndex As Long) As String
    Dim ws_InvSubGrp As String, ws_rStr As String, i As Integer
    Select Case wIndex
        Case Is = 0
            ws_InvSubGrp = "," + ctNone + ","
        Case Is = 1
            ws_InvSubGrp = IIf(ACMB(0) <> ctNone, "," + ACMB(0), "") + "," + mc_Groups(ACMB(0))(5) + "," + ctGrand + ","
        Case Is = 2
            ws_InvSubGrp = IIf(ACMB(0) <> ctNone, "," + ACMB(0), "") + IIf(ACMB(1) <> ctNone, "," + ACMB(1) + ",", ",") + mc_Groups(ACMB(0))(5) + "," + mc_Groups(ACMB(1))(5) + "," + ctGrand + ","
    End Select
    If InStr(1, ws_InvSubGrp, ",((All)),") <> 0 Then
        For i = 1 To mc_Groups.Count
            If mc_Groups(i)(0) <> ctNone Then
                ws_rStr = ws_rStr + mc_Groups(i)(0) + ","
            End If
        Next i
        ws_InvSubGrp = ws_InvSubGrp + ws_rStr
    End If
    GetInvSubGrp = ws_InvSubGrp
End Function
Private Sub ACMB_LostFocus(Index As Integer)
    Call ChkCmb
End Sub
Private Sub ChkCmb()
    Dim i As Integer
    For i = 0 To 2
        If InStr(1, GetInvSubGrp(i), "," + ACMB(i) + ",") <> 0 Then
            ACMB(i) = GetDefaVal(i)
        End If
    Next i
   For i = 3 To 5
        If GetPos(ACMB(i), i) < 0 Then
             ACMB(i) = GetDefaVal(i)
        End If
    Next i
End Sub
Private Function GetPos(ByVal wReqStr As String, ByVal wIndex As Integer) As Integer
    Dim wi_Pos As Integer, wv_rArr As Variant, i As Integer
    wi_Pos = -1
    wv_rArr = mc_Groups(ACMB(wIndex - 3))(6)
    For i = 0 To UBound(wv_rArr)
        If wReqStr = wv_rArr(i) Then
            wi_Pos = i
            Exit For
        End If
    Next i
    GetPos = wi_Pos
End Function
Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
    Dim i As Integer, ws_rStr As String
    If Ambient.UserMode Then
        Add ctNone, "", "", "", "", "((All))", "", "", ""
        Add ctDetail, "", "", "", "", "((All))", "", "", ""
        Add ctGrand, "", "", "", "", "((All))", "", "", ""
        For i = 0 To 5
            ws_rStr = GetDefaVal(i)
            ACMB(i).Additem ws_rStr
            ACMB(i) = ws_rStr
        Next i
    End If
    
    '*** Jenny (13/05/06)
    BckColor = PropBag.ReadProperty("BckColor", &H80000005)
    ForColor = PropBag.ReadProperty("ForColor", &H80000008)
    '*** Jenny (13/05/06)
End Sub
Private Sub UserControl_WriteProperties(PropBag As PropertyBag)
    '*** Jenny (13/05/06)
    Call PropBag.WriteProperty("BckColor", ACMB(0).BackColor, &H80000005)
    Call PropBag.WriteProperty("ForColor", ACMB(0).ForeColor, &H80000008)
    '*** Jenny (13/05/06)
End Sub

Public Sub Gen3LRep()
    Dim mf_OneLvl As Boolean, mf_TwoLvl As Boolean, mf_ThreeLvl As Boolean
    Set mo_RepSource = Parent.adc.RepSource
    mf_OneLvl = (ACMB(0) <> ctNone And ACMB(1) = ctNone And ACMB(2) = ctNone)
    mf_TwoLvl = (ACMB(0) <> ctNone And ACMB(1) <> ctNone And ACMB(2) = ctNone)
    mf_ThreeLvl = (ACMB(0) <> ctNone And ACMB(1) <> ctNone And ACMB(2) <> ctNone)
    If mf_OneLvl Then
        If ACMB(0) = ctDetail Then
            HideSec Array("G1H", "G2H", "G3F", "G2F", "G1F")
            ms_GrpFld1 = ctNone: ms_GrpFld2 = ctNone: ms_GrpFld3 = ctNone
        ElseIf ACMB(0) = ctGrand Then
            HideSec Array("G1H", "G2H", "DET", "G3F", "G2F", "G1F")
            ms_GrpFld1 = ctNone: ms_GrpFld2 = ctNone: ms_GrpFld3 = ctNone
        Else
            HideSec Array("G1H", "G2H", "DET", "G2F", "G1F")
            ms_GrpFld1 = ctNone: ms_GrpFld2 = ctNone: ms_GrpFld3 = ACMB(0)
       End If
    ElseIf mf_TwoLvl Then
        ms_GrpFld1 = ctNone: ms_GrpFld2 = ACMB(0)
        If ACMB(1) = ctDetail Then
            HideSec Array("G1H", "G3F", "G1F")
            ms_GrpFld3 = ctNone
        Else
            HideSec Array("G1H", "DET", "G1F")
            ms_GrpFld3 = ACMB(1)
        End If
    ElseIf mf_ThreeLvl Then
        ms_GrpFld1 = ACMB(0): ms_GrpFld2 = ACMB(1)
        If ACMB(2) = ctDetail Then
            HideSec Array("G3F")
            ms_GrpFld3 = ctNone
        Else
           HideSec Array("DET")
           ms_GrpFld3 = ACMB(2)
        End If
    End If
    mo_RepSource.Sections("G1H").ReportObjects("tGH1").SetText mc_Groups(ms_GrpFld1)(0)
    mo_RepSource.Sections("G2H").ReportObjects("tGH2").SetText mc_Groups(ms_GrpFld2)(0)
    mo_RepSource.Sections("PHGRP").ReportObjects("tGH3").SetText mc_Groups(ms_GrpFld3)(0)

    If ACMB(0) = ctDetail Or ACMB(1) = ctDetail Or ACMB(2) = ctDetail Then
        HideSec "PHGRP"
        ShiftHdFlds "PHDET", mc_Groups(ms_GrpFld1)(4) + "," + mc_Groups(ms_GrpFld2)(4) + "," + mc_Groups(ms_GrpFld3)(4)
        ShiftFlds "DET", mc_Groups(ms_GrpFld1)(3) + "," + mc_Groups(ms_GrpFld2)(3) + "," + mc_Groups(ms_GrpFld3)(3)
'        Call ShiftFldsTot
    Else
        HideSec "PHDET"
    End If
End Sub
Private Sub ShiftFlds(SecNm As String, FldLst As String)
    Dim wo_rFld As Object, ws_FldNm As String, wo_sFld As Object
    FldLst = "," + UCase(FldLst) + ","
    mi_supwdth = 0
    With mo_RepSource.Sections(SecNm)
        For Each wo_rFld In .ReportObjects
'??? consider speed
            If TypeOf wo_rFld Is CRAXDRT.FieldObject Then
                If TypeOf wo_rFld.Field Is CRAXDRT.DatabaseFieldDefinition Then ws_FldNm = UCase(wo_rFld.Field.DatabaseFieldName)
                If TypeOf wo_rFld.Field Is CRAXDRT.FormulaFieldDefinition Then ws_FldNm = UCase(wo_rFld.Field.FormulaFieldName)
                If InStr(1, FldLst, "," + ws_FldNm + ",") <> 0 Then
                    wo_rFld.Suppress = True
                    mi_supwdth = mi_supwdth + wo_rFld.Width
                    For Each wo_sFld In .ReportObjects
                        If wo_sFld.Left > wo_rFld.Left Then
                            wo_sFld.Left = wo_sFld.Left - wo_rFld.Width
                        End If
                    Next wo_sFld
                    Call ShiftTotFlds(wo_rFld.Left, wo_rFld.Width)
                End If
            End If
        Next wo_rFld
    End With
End Sub
Private Sub ShiftHdFlds(SecNm As String, FldLst As String)
    Dim wo_rFld As Object, ws_FldNm As String, wo_sFld As Object
    FldLst = "," + UCase(FldLst) + ","
    With mo_RepSource.Sections(SecNm)
        For Each wo_rFld In .ReportObjects
            ws_FldNm = UCase(wo_rFld.Name)
            If InStr(1, FldLst, "," + ws_FldNm + ",") <> 0 Then
                wo_rFld.Suppress = True
                For Each wo_sFld In .ReportObjects
                    If wo_sFld.Left > wo_rFld.Left Then
                        wo_sFld.Left = wo_sFld.Left - wo_rFld.Width
                    End If
                Next wo_sFld
            End If
        Next wo_rFld
    End With
End Sub
Private Sub SuppFlds(SecNm As String, HdFldLst As Variant)
    Dim i As Integer
   For i = 0 To UBound(HdFldLst)
        If HdFldLst(i) <> "" Then
            mo_RepSource.Sections(SecNm).ReportObjects(HdFldLst(i)).Suppress = True
        End If
    Next i
End Sub

'*** Jenny (13/06/05)
Public Property Let BckColor(ByVal vNewValue As OLE_COLOR)
    'mo_BckColor = vNewValue
    
    Dim wi_i As Integer
    For wi_i = 0 To ACMB.Count - 1
      ACMB(wi_i).BackColor = vNewValue
    Next wi_i
    
    PropertyChanged "BckColor"
End Property
Public Property Get BckColor() As OLE_COLOR
Attribute BckColor.VB_ProcData.VB_Invoke_Property = ";Additional"
    'BckColor = mo_BckColor
    BckColor = ACMB(0).BackColor
End Property
Public Property Let ForColor(ByVal vNewValue As OLE_COLOR)
    'mo_ForColor = vNewValue
    Dim wi_i As Integer
    For wi_i = 0 To ACMB.Count - 1
      ACMB(wi_i).ForeColor = vNewValue
    Next wi_i
    
    PropertyChanged "ForColor"
End Property
Public Property Get ForColor() As OLE_COLOR
Attribute ForColor.VB_ProcData.VB_Invoke_Property = ";Additional"
    'ForColor = mo_ForColor
    ForColor = ACMB(0).ForeColor
End Property
'*** Jenny (13/06/05)

Public Property Get GrpOpts() As String
    GrpOpts = IIf(ms_GrpFld1 = ctNone, "", ms_GrpFld1) + "," + _
               IIf(ms_GrpFld2 = ctNone, "", ms_GrpFld2) + "," + _
               IIf(ms_GrpFld3 = ctNone, "", ms_GrpFld3)
End Property
Public Property Get GrpFldLst(Optional OrdNum As Integer) As String
    Dim wi_Num As Integer, ws_Fld1 As String, ws_Fld2 As String, ws_Fld3 As String, ws_Fld4 As String, ws_Fld5 As String, ws_Fld6 As String, ws_Fld7 As String, ws_Fld8 As String, ws_Fld9 As String
    
    If IsMissing(OrdNum) Then OrdNum = 0
    
    wi_Num = OrdNum
    If OrdNum > UBound(mc_Groups(ms_GrpFld1)(1)) Then wi_Num = 0
    ws_Fld1 = mc_Groups(ms_GrpFld1)(1)(wi_Num)
    
    wi_Num = OrdNum
    If OrdNum > UBound(mc_Groups(ms_GrpFld2)(1)) Then wi_Num = 0
    ws_Fld2 = mc_Groups(ms_GrpFld2)(1)(wi_Num)
    
    wi_Num = OrdNum
    If OrdNum > UBound(mc_Groups(ms_GrpFld3)(1)) Then wi_Num = 0
    ws_Fld3 = mc_Groups(ms_GrpFld3)(1)(wi_Num)

    wi_Num = OrdNum
    If OrdNum > UBound(mc_Groups(ms_GrpFld1)(2)) Then wi_Num = 0
    ws_Fld4 = mc_Groups(ms_GrpFld1)(2)(wi_Num)

    wi_Num = OrdNum
    If OrdNum > UBound(mc_Groups(ms_GrpFld2)(2)) Then wi_Num = 0
    ws_Fld5 = mc_Groups(ms_GrpFld2)(2)(wi_Num)
    
    wi_Num = OrdNum
    If OrdNum > UBound(mc_Groups(ms_GrpFld3)(2)) Then wi_Num = 0
    ws_Fld6 = mc_Groups(ms_GrpFld3)(2)(wi_Num)
    
    wi_Num = OrdNum
    If OrdNum > UBound(mc_Groups(ms_GrpFld1)(8)) Then wi_Num = 0
    ws_Fld7 = mc_Groups(ms_GrpFld1)(8)(wi_Num)
    
    wi_Num = OrdNum
    If OrdNum > UBound(mc_Groups(ms_GrpFld2)(8)) Then wi_Num = 0
    ws_Fld8 = mc_Groups(ms_GrpFld2)(8)(wi_Num)
    
    wi_Num = OrdNum
    If OrdNum > UBound(mc_Groups(ms_GrpFld3)(8)) Then wi_Num = 0
    ws_Fld9 = mc_Groups(ms_GrpFld3)(8)(wi_Num)

    
    GrpFldLst = ws_Fld1 + " as Grp1," + ws_Fld2 + " as Grp2," + ws_Fld3 + " as Grp3," + _
                ws_Fld4 + " as DGrp1," + ws_Fld5 + " as DGrp2," + ws_Fld6 + " as DGrp3," + _
                ws_Fld7 + " as G1Desc," + ws_Fld8 + " as G2Desc," + ws_Fld9 + " as G3Desc"
              
End Property
Private Sub ShowSec(SecLst As Variant)
    Dim i As Integer
    With mo_RepSource
    If IsArray(SecLst) Then
        For i = 0 To UBound(SecLst)
            .Sections(SecLst(i)).Suppress = False
        Next i
    Else
        .Sections(SecLst).Suppress = False
    End If
    End With
End Sub
Private Sub HideSec(SecLst As Variant)
    Dim i As Integer
    With mo_RepSource
    If IsArray(SecLst) Then
        For i = 0 To UBound(SecLst)
            .Sections(SecLst(i)).Suppress = True
        Next i
    Else
        .Sections(SecLst).Suppress = True
    End If
    End With
End Sub
Private Sub ShiftFldsTot()
    Dim i, j, k As Integer
    Dim wo_Sects As Sections
    Dim wo_Sct As Section
    For i = 1 To mo_RepSource.Areas.Count
        If mo_RepSource.Areas.Item(i).Kind = crGroupFooter Or _
           mo_RepSource.Areas.Item(i).Kind = crPageFooter Or _
           mo_RepSource.Areas.Item(i).Kind = crReportFooter Then
                Set wo_Sects = mo_RepSource.Areas.Item(i).Sections
                For j = 1 To wo_Sects.Count
                    Set wo_Sct = wo_Sects.Item(j)
                    If wo_Sct.Name <> "DET" Then
                    '   when group footer 4 is treated as body in 3 lvl reports ex. totstk and wip challan
                        For k = 1 To wo_Sct.ReportObjects.Count
                            If wo_Sct.ReportObjects.Item(k).Kind = crFieldObject Then
                                If wo_Sct.ReportObjects.Item(k).Field.Kind = crFormulaField Then
                                    wo_Sct.ReportObjects.Item(k).Left = wo_Sct.ReportObjects.Item(k).Left - mi_supwdth
                                End If
                            End If
                        Next k
                    End If
                Next j
        End If
   Next i
End Sub
Private Sub ShiftTotFlds(ByVal pl_Left As Long, pl_Width As Long)
    Dim i, j, k As Integer
    Dim wo_Sects As Sections
    Dim wo_Sct As Section
    For i = 1 To mo_RepSource.Areas.Count
        If mo_RepSource.Areas.Item(i).Kind = crGroupFooter Or _
           mo_RepSource.Areas.Item(i).Kind = crPageFooter Or _
           mo_RepSource.Areas.Item(i).Kind = crReportFooter Then
                Set wo_Sects = mo_RepSource.Areas.Item(i).Sections
                For j = 1 To wo_Sects.Count
                    Set wo_Sct = wo_Sects.Item(j)
                    If wo_Sct.Name <> "DET" Then
                    '   when group footer 4 is treated as body in 3 lvl reports ex. totstk and wip challan
                        For k = 1 To wo_Sct.ReportObjects.Count
                            If wo_Sct.ReportObjects.Item(k).Kind = crFieldObject Then
                                If wo_Sct.ReportObjects.Item(k).Field.Kind = crFormulaField Or wo_Sct.ReportObjects.Item(k).Field.Kind = crRunningTotalField Then
                                    If wo_Sct.ReportObjects.Item(k).Left > pl_Left Then
                                      wo_Sct.ReportObjects.Item(k).Left = wo_Sct.ReportObjects.Item(k).Left - pl_Width
                                    End If
                                End If
                            End If
                        Next k
                    End If
                Next j
        End If
   Next i
End Sub
Public Property Let Value(ByVal Level As en_GrpLevelType, ByVal vNewValue As Variant)
    Call InitCmb(Level, vNewValue)
    Call ChkCmb
End Property
Public Property Get Value(ByVal Level As en_GrpLevelType)
    Value = ACMB(Level)
End Property
Public Sub SetFormula(ByVal ReqReport As CRAXDRT.Report, ByVal FormulaName As String, ByVal Expn As String)
'    FormulaName = "@" + Trim(FormulaName)
    FormulaName = Trim(FormulaName)
    Dim wForm As CRAXDRT.FormulaFieldDefinition
    For Each wForm In ReqReport.FormulaFields
        If UCase(wForm.FormulaFieldName) = UCase(FormulaName) Then
            wForm.Text = Expn
            Exit Sub
        End If
    Next wForm
End Sub

