VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.UserControl MWCTL_ZOOM 
   ClientHeight    =   3600
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4800
   ScaleHeight     =   3600
   ScaleWidth      =   4800
   Begin MSFlexGridLib.MSFlexGrid MSFg 
      Height          =   975
      Index           =   1
      Left            =   30
      TabIndex        =   0
      Top             =   30
      Width           =   3015
      _ExtentX        =   5318
      _ExtentY        =   1720
      _Version        =   393216
   End
   Begin MSFlexGridLib.MSFlexGrid MSFg 
      Height          =   975
      Index           =   2
      Left            =   150
      TabIndex        =   1
      Top             =   510
      Visible         =   0   'False
      Width           =   3015
      _ExtentX        =   5318
      _ExtentY        =   1720
      _Version        =   393216
   End
End
Attribute VB_Name = "MWCTL_ZOOM"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Dim wRs As rdoResultset
Public Sub FillZoom(wxRs As rdoResultset)
    Set wRs = wxRs
    wfld = ""
    warr = Parent.Controls("GRP_REP").zoomfldlst(Group1)
    If IsArray(warr) Then
        wtotcol = UBound(warr, 2) + 2
        waddflds = True
    Else
        wtotcol = 2
        waddflds = False
    End If
    MSFg(1).Cols = 1
    MSFg(1).Cols = wtotcol + 1
    wt_formatstring = "|"
    wt_formatstring = wt_formatstring + Parent.Controls("GRP_REP").Value(Group1) + "|"
    wt_formatstring = wt_formatstring + "Description    " + "|"
    
    
    For i = 0 To UBound(warr, 2) - 1
         wt_formatstring = wt_formatstring + warr(1, i) + "|"
    Next i
    wt_formatstring = Mid(wt_formatstring, 1, Len(wt_formatstring) - 1)
    MSFg(1).FormatString = wt_formatstring
    MSFg(1).Rows = 1
    MSFg(1).Redraw = False
    With MSFg(1)
    Do While Not (wRs.EOF Or wRs.BOF)
        If wRs.rdoColumns("Grp1") <> wfld Then
            .Additem ""
            wRow = .Rows - 1
            .TextMatrix(wRow, 1) = wRs.rdoColumns("dgrp1")
            .TextMatrix(wRow, 2) = wRs.rdoColumns("grp1desc")
            If waddflds Then
                For i = 0 To UBound(warr, 2) - 1
                    ColPos = 3 + i
                    If warr(3, i) = en_ZoomFldType.Plain Then
                        .TextMatrix(wRow, ColPos) = wRs.rdoColumns(warr(0, i))
                    End If
                    If warr(3, i) = en_ZoomFldType.NormalTotal Then
                        .TextMatrix(wRow, ColPos) = 0
                    End If
                Next i
            End If
        End If
        If waddflds Then
            For i = 0 To UBound(warr, 2) - 1
                ColPos = 3 + i
                If warr(3, i) = en_ZoomFldType.NormalTotal Then
                    .TextMatrix(wRow, ColPos) = .TextMatrix(wRow, ColPos) + wRs.rdoColumns(warr(0, i))
                End If
                ColPos = ColPos + 1
            Next i
        End If
        wfld = wRs.rdoColumns("Grp1")
        wRs.MoveNext
    Loop
    MSFg(1).Redraw = True
    End With
End Sub
Private Sub MSFg_DblClick(index As Integer)
If index = 1 Then
    wfld = ""
    warr = Parent.Controls("GRP_REP").zoomfldlst(Group2)
    If IsArray(warr) Then
        wtotcol = UBound(warr, 2) + 2
        waddflds = True
    Else
        wtotcol = 2
        waddflds = False
    End If
    MSFg(2).Cols = 1
    MSFg(2).Cols = wtotcol + 1
    wt_formatstring = "|"
    wt_formatstring = wt_formatstring + Parent.Controls("GRP_REP").Value(Group2) + "|"
    wt_formatstring = wt_formatstring + "Description    " + "|"
    
    
    For i = 0 To UBound(warr, 2) - 1
         wt_formatstring = wt_formatstring + warr(1, i) + "|"
    Next i
    wt_formatstring = Mid(wt_formatstring, 1, Len(wt_formatstring) - 1)
    MSFg(2).FormatString = wt_formatstring
    MSFg(2).Rows = 1
    MSFg(2).Redraw = False
    wRs.MoveFirst
    With MSFg(2)
    Do While Not (wRs.EOF Or wRs.BOF)
'        If wRs.rdoColumns("Grp1") <> wfld Then
            .Additem ""
            wRow = .Rows - 1
            .TextMatrix(wRow, 1) = wRs.rdoColumns("dgrp2")
            .TextMatrix(wRow, 2) = wRs.rdoColumns("grp2desc")
            If waddflds Then
                For i = 0 To UBound(warr, 2) - 1
                    ColPos = 3 + i
                    If warr(3, i) = en_ZoomFldType.Plain Then
                        .TextMatrix(wRow, ColPos) = wRs.rdoColumns(warr(0, i))
                    End If
                    If warr(3, i) = en_ZoomFldType.NormalTotal Then
                        .TextMatrix(wRow, ColPos) = 0
                    End If
                Next i
            End If
'        End If
        If waddflds Then
            For i = 0 To UBound(warr, 2) - 1
                ColPos = 3 + i
                If warr(3, i) = en_ZoomFldType.NormalTotal Then
                    .TextMatrix(wRow, ColPos) = .TextMatrix(wRow, ColPos) + wRs.rdoColumns(warr(0, i))
                End If
                ColPos = ColPos + 1
            Next i
        End If
        wfld = wRs.rdoColumns("Grp2")
        wRs.MoveNext
    Loop
    MSFg(2).Redraw = True
    MSFg(2).Visible = True
    MSFg(2).ZOrder
    End With
End If
End Sub

Private Sub UserControl_Resize()
    MSFg(1).Left = 0
    MSFg(1).Top = 0
    MSFg(1).Height = UserControl.Extender.Height
    MSFg(1).Width = UserControl.Extender.Width
    MSFg(2).Height = IIf(UserControl.Extender.Height - 600 < 0, 0, UserControl.Extender.Height - 600)
    MSFg(2).Width = IIf(UserControl.Extender.Width - 200 < 0, 0, UserControl.Extender.Width - 200)
End Sub
