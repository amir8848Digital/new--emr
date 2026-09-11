VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "msflxgrd.ocx"
Begin VB.UserControl MWCTL_HLP 
   ClientHeight    =   3600
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4800
   Enabled         =   0   'False
   KeyPreview      =   -1  'True
   ScaleHeight     =   3600
   ScaleWidth      =   4800
   Begin MSFlexGridLib.MSFlexGrid FLX_GRD 
      Height          =   1095
      Left            =   0
      TabIndex        =   0
      Top             =   30
      Width           =   1815
      _ExtentX        =   3201
      _ExtentY        =   1931
      _Version        =   393216
      ScrollBars      =   1
      MousePointer    =   11
   End
End
Attribute VB_Name = "MWCTL_HLP"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
Dim mo_RetInto As Object
Dim ms_SqlStrg As String
Dim ms_ColHeadg As String
Dim ms_ColFlds As String
Dim ms_ColMask As String
Dim ms_RetFldNm As String
Dim ma_RetOth As Variant
Dim mi_SearchColNum As Integer
Dim mf_MultiSelect As Boolean
Dim mf_stColorChng As Boolean
Dim mf_HlpTaken As Boolean
Enum en_HlpMode
    hmHlpAvlbl = 0
    hmHlpNotavlbl = 1
    hmHlpBrowsing = 2
End Enum
Dim me_Mode As en_HlpMode
Dim ma_FormatArr() As String
Dim ml_MatchRow As Long
Dim mo_rsHlp As MDORowSet

Private Sub FLX_GRD_GotFocus()
    If FLX_GRD.Cols > 1 Then FLX_GRD.COL = 1
End Sub
Private Sub FLX_GRD_KeyUp(KeyCode As Integer, Shift As Integer)
   If KeyCode = vbKeyEscape Then
        mo_RetInto.SetFocus
   End If
End Sub
Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
    me_Mode = hmHlpNotavlbl
End Sub
Private Sub UserControl_Resize()
    FLX_GRD.Height = UserControl.Height
    FLX_GRD.Width = UserControl.Width
End Sub
Private Sub UserControl_ExitFocus()
    me_Mode = hmHlpAvlbl
    If mo_RetInto.EnAndVis Then mo_RetInto.SetFocus
    UserControl.Extender.Visible = False
End Sub
Private Sub flx_grd_CLICK()
    me_Mode = hmHlpBrowsing
End Sub
Public Property Set RetInto(ByVal New_RetInto As Object)
    Set mo_RetInto = New_RetInto
End Property
Public Property Let Sqlstrg(ByVal vNewValue As String)
    ms_SqlStrg = vNewValue
    If vNewValue <> "" Then
        me_Mode = hmHlpAvlbl
    End If
End Property
Public Property Let ColHeadg(ByVal vNewValue As String)
    ms_ColHeadg = vNewValue
End Property
Public Property Let ColFlds(ByVal vNewValue As String)
    ms_ColFlds = vNewValue
End Property
Public Property Let ColMask(ByVal vNewValue As String)
    ms_ColMask = vNewValue
End Property
Public Property Let RetFldNm(ByVal vNewValue As String)
    ms_RetFldNm = vNewValue
End Property
Public Property Let RetOth(ByVal vNewValue As Variant)
    ma_RetOth = vNewValue
End Property
Public Property Let SearchColNum(ByVal vNewValue As Integer)
    mi_SearchColNum = vNewValue
End Property
Public Property Let Mode(ByVal vNewValue As en_HlpMode)
    me_Mode = vNewValue
    If me_Mode = hmHlpNotavlbl Then
        Set mo_RetInto = Nothing
        ms_SqlStrg = ""
        ms_ColHeadg = ""
        ms_ColFlds = ""
        ms_ColMask = ""
        ms_RetFldNm = ""
        ma_RetOth = ""
        mi_SearchColNum = 0
        mf_MultiSelect = False
        mf_HlpTaken = False
        Set mo_rsHlp = Nothing
    End If
End Property
Public Property Get Mode() As en_HlpMode
    Mode = me_Mode
End Property
Public Sub ReplSelVal()
    Dim i As Integer
    If FLX_GRD.Row = 0 Then Exit Sub
    Dim ws_AddTxt As String, ws_Txt As String, ws_GrdTxt As String, wi_rPos As Integer
    If mf_MultiSelect Then
        ws_AddTxt = GetAddTxt()
        If FLX_GRD.TextMatrix(FLX_GRD.Row, 0) <> "" Then
            FLX_GRD.TextMatrix(FLX_GRD.Row, 0) = ""
            FLX_GRD.COL = mi_SearchColNum
            ws_Txt = "," + mo_RetInto.Text + ","
            ws_GrdTxt = ws_AddTxt + FLX_GRD.TextMatrix(FLX_GRD.Row, mi_SearchColNum) + ws_AddTxt
            wi_rPos = InStr(1, ws_Txt, "," + ws_GrdTxt + ",")
            ws_Txt = Mid(ws_Txt, 1, wi_rPos - 1) + Mid(ws_Txt, wi_rPos + Len(ws_GrdTxt) + 1)
            If ws_Txt <> "" Then
                If Left(ws_Txt, 1) = "," Then ws_Txt = Mid(ws_Txt, 2)
                If Right(ws_Txt, 1) = "," Then ws_Txt = Left(ws_Txt, Len(ws_Txt) - 1)
            End If
            mo_RetInto.Text = ws_Txt
        Else
            FLX_GRD.TextMatrix(FLX_GRD.Row, 0) = "¶"
            FLX_GRD.COL = mi_SearchColNum
            ws_Txt = mo_RetInto.Text
            mo_RetInto.Text = Trim(ws_Txt) + IIf(Trim(ws_Txt) <> "", ",", "") + ws_AddTxt + FLX_GRD.TextMatrix(FLX_GRD.Row, mi_SearchColNum) + ws_AddTxt
        End If
    Else
        If FLX_GRD.Row > 0 Then
            mo_rsHlp.Bookmark = FLX_GRD.RowData(FLX_GRD.Row)
        End If
        If Not (mo_rsHlp.EOF Or mo_rsHlp.BOF) Then
            mo_RetInto.TxtText = mo_rsHlp.FldValue(ms_RetFldNm)
            If IsArray(ma_RetOth) Then
            For i = 0 To UBound(ma_RetOth) - 1 Step 2
                If ma_RetOth(i + 1) <> "" Then
                    If TypeOf mo_RetInto Is MWCTL_FLX Then
                        mo_RetInto.Value(mo_RetInto.Row, ma_RetOth(i + 1)) = mo_rsHlp.FldValue(ma_RetOth(i))
                    Else
                        Parent.adc.Ctl(ma_RetOth(i + 1)).Data = mo_rsHlp.FldValue(ma_RetOth(i))
                    End If
                End If
            Next i
            End If
        Else
            mo_RetInto.Text = ""
        End If
    End If
End Sub
Public Function Gen_Hlp() As Boolean
    If ms_SqlStrg = "" Then
        Gen_Hlp = False
        Exit Function
    End If
    Set mo_rsHlp = Parent.adc.Connection.OpenRes(ms_SqlStrg)
    mf_HlpTaken = True
    With FLX_GRD
        .Cols = 0
        .FormatString = IIf(mf_MultiSelect, "  ", "") + ms_ColHeadg
        .Cols = .Cols - 1
        .Rows = 1
        Call Format_Grd
        Call set_pos
        Call FndMatch
    End With
    Gen_Hlp = True
End Function
Private Sub set_pos()
    Dim wl_NoRows As Long, wl_reqH As Long, wo_Obj As Object, wl_cTop As Long, wl_cLeft As Long, wl_cWidth As Long, wl_cHeight As Long, wl_hTop As Long, wl_hHt As Long, wf_hMaxSet As Boolean, wl_reqW As Long, wl_hLeft As Long, wl_hWdth As Long, i As Integer
    With Parent
       If mo_rsHlp.RecCount = 0 Then
            wl_NoRows = 1
       Else
            wl_NoRows = mo_rsHlp.RecCount
       End If
       wl_reqH = FLX_GRD.RowHeight(0) * (wl_NoRows + 1)
       wl_reqH = wl_reqH + 100
       
       Set wo_Obj = mo_RetInto
       If TypeOf wo_Obj Is MWCTL_FLX Then
            wl_cTop = wo_Obj.Top + wo_Obj.object.Top
            wl_cLeft = wo_Obj.Left + wo_Obj.object.Left
            wl_cWidth = wo_Obj.object.Width
            wl_cHeight = wo_Obj.object.Height
       Else
            wl_cTop = wo_Obj.Top
            wl_cLeft = wo_Obj.Left
            wl_cWidth = wo_Obj.Width
            wl_cHeight = wo_Obj.Height
       End If
       Do While True
            If wo_Obj.Container.Name = Parent.Name Then
                Exit Do
            End If
            wl_cTop = wl_cTop + wo_Obj.Container.Top
            wl_cLeft = wl_cLeft + wo_Obj.Container.Left
            Set wo_Obj = wo_Obj.Container
       Loop
       If wl_cTop - (.Height - 405) / 2 <= 0 Then 'fld in upper half
            wl_hTop = wl_cTop + wl_cHeight  'start just below the fld
            If wl_reqH < (.Height - 405) - wl_hTop Then
                wl_hHt = wl_reqH
            Else
                wl_hHt = (.Height - 405) - wl_hTop 'form ht - startg pos of hlp box ie entire ht avlbl below the fld
                wf_hMaxSet = True
                wl_reqW = 250 'for scroll bar
            End If
       Else 'fld in lower half
           wl_hTop = IIf(wl_reqH < wl_cTop, wl_cTop - wl_reqH, 0) 'start at top of form
           If wl_reqH < wl_cTop Then
                wl_hHt = wl_reqH
            Else
                wl_hHt = wl_cTop
                wl_reqW = 250 'for scroll bar
            End If 'spc avlbl from begg of form to top of fld
       End If
       For i = 0 To FLX_GRD.Cols - 1
           wl_reqW = wl_reqW + FLX_GRD.ColWidth(i)
       Next i
       wl_reqW = wl_reqW + 120
       If wl_cLeft - .Width / 2 <= 0 Then 'fld in left half
           wl_hLeft = wl_cLeft 'start at pos of fld
           If wl_reqW < .Width - wl_hLeft Then
                wl_hWdth = wl_reqW
            Else
                wl_hWdth = .Width - wl_hLeft
                If Not wf_hMaxSet Then wl_hHt = wl_hHt + 250 'for scrol bar
            End If
       Else 'fld in right half
           wl_hLeft = IIf(wl_reqW < wl_cLeft + wl_cWidth, (wl_cLeft + wl_cWidth) - wl_reqW, .Left)
           If wl_reqW < wl_cLeft + wl_cWidth Then
                wl_hWdth = wl_reqW
            Else
                wl_hWdth = wl_cLeft + wl_cWidth
                If Not wf_hMaxSet Then wl_hHt = wl_hHt + 250 'for scrol bar
            End If
       End If
    End With
    UserControl.Extender.Top = wl_hTop
    UserControl.Extender.Height = wl_hHt
    UserControl.Extender.Left = wl_hLeft
    UserControl.Extender.Width = wl_hWdth
End Sub
Public Sub FndMatch(Optional KeyCode As Variant)
    Dim wv_GrdStartVal As Variant, wv_GrdEndVal As Variant, wv_reqVal As Variant, wi_CommPos As Integer, i As Integer
    If mo_rsHlp.RecCount = 0 Then Exit Sub
    wv_GrdStartVal = GrdVal(1)
    If FLX_GRD.Rows > 2 Then wv_GrdEndVal = GrdVal(FLX_GRD.Rows - 1)
    With FLX_GRD
        If IsMissing(KeyCode) Then
            wv_reqVal = mo_RetInto.TxtText
            If mf_MultiSelect Then
                wi_CommPos = InStr(1, wv_reqVal, ",")
                If wi_CommPos >= 1 Then wv_reqVal = Mid(wv_reqVal, 1, wi_CommPos - 1)
                If GetAddTxt <> "" And Len(wv_reqVal) >= 2 Then wv_reqVal = Mid(wv_reqVal, 2, Len(wv_reqVal) - 2)
            End If
            If (Not mo_RetInto.TxtIsBlank) And wv_reqVal >= wv_GrdStartVal And wv_reqVal <= wv_GrdEndVal Then
                ml_MatchRow = FndMatchRow()
            Else
                ml_MatchRow = 1
                If mo_RetInto.TxtIsBlank Then
                    If mo_rsHlp.RecCount > 0 Then mo_rsHlp.MoveFirst
                End If
                
                If wv_reqVal < wv_GrdStartVal Then
                    If .Rows > 1 Then
                        mo_rsHlp.Bookmark = .RowData(1)
                        mo_rsHlp.MovePrevious
                    End If
                    Do While Not (mo_rsHlp.EOF() Or mo_rsHlp.BOF())
                        If RSVal <= wv_reqVal Then
                            If RSVal <> wv_reqVal Then ml_MatchRow = ml_MatchRow + 1
                            Exit Do
                        End If
                        mo_rsHlp.MovePrevious
                    Loop
                End If
                
                If wv_reqVal > wv_GrdEndVal Then
                    If .Rows > 1 Then
                        mo_rsHlp.Bookmark = .RowData(.Rows - 1)
                        mo_rsHlp.MoveNext
                    End If
                    Do While Not (mo_rsHlp.EOF Or mo_rsHlp.BOF)
                        If RSVal >= wv_reqVal Then
                            If RSVal <> wv_reqVal Then
                                If Not mo_rsHlp.BOF Then
                                    mo_rsHlp.MovePrevious
                                    If Not mo_rsHlp.BOF Then ml_MatchRow = ml_MatchRow + 1
                                End If
                            End If
                            Exit Do
                        End If
                        mo_rsHlp.MoveNext
                    Loop
                End If
                
                Call Fil_Grd
            End If
        Else
            If KeyCode = vbKeyDown Then
                If .Row = .Rows - 1 Then
                    If .Rows = 2 Then
                      mo_rsHlp.Bookmark = .RowData(1)
                      ml_MatchRow = .Rows
                    Else
                      ml_MatchRow = .Rows - 1
                      mo_rsHlp.Bookmark = .RowData(2)
                    End If
                    Call Fil_Grd
                Else
                    ml_MatchRow = .Row + 1
                End If
            End If
'??? find easier using ados find function
'??? help in descending order
'??? anything better for help as required
            If KeyCode = vbKeyPageDown Then
                If .Row = .Rows - 1 Then
                    mo_rsHlp.Bookmark = .RowData(.Rows - 1)
                    mo_rsHlp.MoveNext
                    ml_MatchRow = .Rows - 1
                    Call Fil_Grd
                Else
                    ml_MatchRow = .Rows - 1
                End If
            End If
            
            
            If KeyCode = vbKeyUp Then
                If .Row = 1 Then
                    mo_rsHlp.Bookmark = .RowData(1)
                    mo_rsHlp.MovePrevious
                    ml_MatchRow = 1
                    Call Fil_Grd
                Else
                    ml_MatchRow = .Row - 1
                End If
            End If
            
            If KeyCode = vbKeyPageUp Then
                If .Row = 1 Then
                    mo_rsHlp.Bookmark = .RowData(1)
                    For i = 1 To MinRows
                        If mo_rsHlp.BOF Then Exit For
                        mo_rsHlp.MovePrevious
                    Next i
                    ml_MatchRow = 1
                    Call Fil_Grd
                Else
                    ml_MatchRow = 1
                End If
            End If
        End If
        
        .Row = IIf(ml_MatchRow < .Rows, ml_MatchRow, .Rows - 1)
        If Not .RowIsVisible(.Row) Then .TopRow = IIf(ml_MatchRow < .Rows, ml_MatchRow, .Rows - 1)
        .Row = IIf(ml_MatchRow < .Rows, ml_MatchRow, .Rows - 1)
        .COL = 0
        .ColSel = .Cols - 1
    End With
End Sub
Private Function FndMatchRow()
    Dim wv_reqVal As Variant, i As Long
    With FLX_GRD
        .Row = 1
        wv_reqVal = mo_RetInto.TxtText
        For i = 1 To .Rows - 1
            If GrdVal(i) >= wv_reqVal Then
                mo_rsHlp.Bookmark = .RowData(i)
                Exit For
            End If
        Next i
        FndMatchRow = i
    End With
End Function
Private Sub Format_Grd()
    Dim wi_cPos As Integer, wi_cfPos As Integer, wi_rCtr As Integer, wi_Pos As Integer, wi_fPos As Integer
    With FLX_GRD
        wi_cPos = 1
        wi_cfPos = 1
        ReDim ma_FormatArr(2, .Cols - 1)
        wi_rCtr = 1
        Do While True
            wi_Pos = InStr(wi_cPos + 1, ms_ColFlds, "|")
            wi_fPos = InStr(wi_cfPos + 1, ms_ColMask, "|")
            If wi_Pos = 0 Then Exit Do
            ma_FormatArr(0, wi_rCtr) = Mid(ms_ColFlds, wi_cPos + 1, wi_Pos - wi_cPos - 1)
            If ms_ColMask = "" Then
                ma_FormatArr(1, wi_rCtr) = ""
            Else
                ma_FormatArr(1, wi_rCtr) = Mid(ms_ColMask, wi_cfPos + 1, wi_fPos - wi_cfPos - 1)
            End If
            wi_cPos = wi_Pos
            wi_cfPos = wi_fPos
            wi_rCtr = wi_rCtr + 1
        Loop
    End With
End Sub
Private Sub Fil_Grd()
    Dim i As Long, wv_OrgRec As Variant, j As Long, wl_AddRows As Long, wv_Txt As Variant, wv_Filval As Variant, wv_AddTxt As Variant
    If mo_rsHlp.EOF Then
        wv_OrgRec = "EOF"
    ElseIf mo_rsHlp.BOF Then
        wv_OrgRec = "BOF"
    Else
        wv_OrgRec = mo_rsHlp.Bookmark
    End If
  
    i = 1
    Do While Not mo_rsHlp.EOF
       If i >= MinRows Then Exit Do
       mo_rsHlp.MoveNext
       i = i + 1
    Loop
    
    If wv_OrgRec = "BOF" Then
        If mo_rsHlp.RecCount <> 0 Then mo_rsHlp.MoveFirst
    ElseIf wv_OrgRec = "EOF" Then
        If mo_rsHlp.RecCount <> 0 Then mo_rsHlp.MoveLast
    Else
        mo_rsHlp.Bookmark = wv_OrgRec
    End If
    
    If i < MinRows Then
        wl_AddRows = MinRows - i
        For j = 1 To wl_AddRows
            If mo_rsHlp.BOF Then Exit For
            mo_rsHlp.MovePrevious
        Next j
        ml_MatchRow = ml_MatchRow + wl_AddRows
    End If
    
    mf_stColorChng = True
    wv_Txt = "," + Trim(mo_RetInto.TxtText) + ","
    With FLX_GRD
        .Redraw = False
        .Rows = 1
        If mo_rsHlp.BOF And mo_rsHlp.RecCount <> 0 Then mo_rsHlp.MoveFirst
        If mo_rsHlp.EOF And mo_rsHlp.RecCount <> 0 Then mo_rsHlp.MoveLast
         Do While Not (mo_rsHlp.EOF Or mo_rsHlp.BOF)
            If .Rows > MinRows() Then Exit Do
            .Additem ""
            For i = 1 To .Cols - 1
                wv_Filval = mo_rsHlp.FldValue(ma_FormatArr(0, i))
                If Not IsNull(wv_Filval) Then
                    .TextMatrix(.Rows - 1, i) = IIf(ma_FormatArr(1, i) <> "", Format(wv_Filval, ma_FormatArr(1, i)), wv_Filval)
                End If
            Next i
            .RowData(.Rows - 1) = mo_rsHlp.Bookmark
            
            If mf_MultiSelect Then
                wv_AddTxt = GetAddTxt
                If InStr(1, wv_Txt, "," + wv_AddTxt + .TextMatrix(.Rows - 1, mi_SearchColNum) + wv_AddTxt + ",") > 0 Then
                    .TextMatrix(.Rows - 1, 0) = "¶"
                    .Row = .Rows - 1
                    .COL = mi_SearchColNum
                 End If
            End If
            mo_rsHlp.MoveNext
        Loop
        
        If .Cols > 1 Then .COL = 1
        If .Rows > 1 Then .Row = 1
        
        mf_stColorChng = False
        FLX_GRD.Redraw = True
    End With
End Sub
Public Property Get MultiSelect() As Boolean
    MultiSelect = mf_MultiSelect
End Property
Public Property Let MultiSelect(ByVal vNewValue As Boolean)
    mf_MultiSelect = vNewValue
End Property
Public Property Get Text() As Variant
    If Not (mo_rsHlp.EOF Or mo_rsHlp.BOF) Then
        Text = mo_rsHlp.FldValue(ms_RetFldNm)
    Else
        Text = "((Null))"
    End If
End Property
Public Property Get HlpTaken() As Boolean
    HlpTaken = mf_HlpTaken
End Property
Private Function GrdVal(Optional ByVal pv_RowNum As Long) As Variant
    If IsMissing(pv_RowNum) Then pv_RowNum = FLX_GRD.Row
    If FLX_GRD.Rows > 1 And pv_RowNum >= 1 And pv_RowNum <= FLX_GRD.Rows - 1 Then
        If Not (mo_RetInto Is Nothing) Then
            If mo_RetInto.DataType = fdtNumber Or mo_RetInto.DataType = fdtFloat Or mo_RetInto.DataType = fdtTimeType Then
                GrdVal = Val(FLX_GRD.TextMatrix(pv_RowNum, mi_SearchColNum))
            Else
                GrdVal = FLX_GRD.TextMatrix(pv_RowNum, mi_SearchColNum)
            End If
        End If
    End If
End Function
Private Function RSVal() As Variant
    If Not (mo_rsHlp.EOF Or mo_rsHlp.BOF) Then
        RSVal = mo_rsHlp.FldValue(ms_RetFldNm)
    End If
End Function
Private Function MinRows() As Long
    MinRows = MWLib.Div(FLX_GRD.Height, FLX_GRD.RowHeight(0)) - 1
End Function
Private Function GetAddTxt() As String
    If Parent.adc.EntType = xetReport Then
          GetAddTxt = "'"
    Else
          GetAddTxt = ""
    End If
End Function

