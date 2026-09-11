VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "MWCLS_HLP"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Public Sub DiaCSGrd()
    With MWHlp
        .Sqlstrg = " select DCGCD,DCGDesc from DiaCSGrd where " + _
                    " DCGRmCode is null order by DCGCD "
        .ColHeadg = "|Code |Description |"
        .ColFlds = "|DCGCD|DCGDesc|"
        .ColMask = "|||"
        .RetFldNm = "DCGCD"
        .SearchColNum = 1
    End With
End Sub
'**** (Jen) DiaStk 26/08/00 ****
Public Sub DiaStkOrdNo()
    With MWHlp
        .Sqlstrg = " select ODOrdTC+'/'+ODOrdYY+'/'+ODOrdChr+'/'+ " + _
                   "convert(varchar,ODOrdNo)+'/'+convert(varchar,ODOrdSr)+ " + _
                   "'/'+ODDesgnCd as qOrd From OrdDesgn"
        .ColHeadg = "|Order+Design                                    |"
        .ColFlds = "|qOrd|"
        .ColMask = "||"
        .RetFldNm = "qOrd"
        .SearchColNum = 1
    End With
End Sub
'**** (Jen) DiaStk 26/08/00 ****

'**** (Jen) DiaStk 26/08/00 ****
Public Sub LaNo()
  wStrg = "Select SBOE, SDate from StkTxn where SVouchTc='PR' " + _
            " and SBOE<>'' "
  With MWHlp
    .Sqlstrg = wStrg
    .ColHeadg = "|La No                              |Date        |"
    .ColFlds = "|SBOE|SDate|"
    .ColMask = "|||"
    .RetFldNm = "SBOE"
    .SearchColNum = 1
  End With
End Sub
'**** (Jen) DiaStk 26/08/00 ****

'******** (Jen) 05/09/00 ********
Public Sub ExlInvNo(mYy As String)
  wsqlstrg = " SELECT ENo,EDt,ECustCd " + _
              " from ExlInv where EYy='" + mYy + "' Order By ENo"
  With MWHlp
          .Sqlstrg = wsqlstrg
          .ColHeadg = "|Number  |Date        |Customer Code   |"
          .ColFlds = "|ENo|EDt|ECustCd|"
          .ColMask = "||dd/mm/yy||"
          .SearchColNum = 1
          .RetFldNm = "ENo"
  End With
End Sub
'******** (Jen) 05/09/00 ********
Public Sub TagNo(ByVal YY As String, ByVal TC As String, ByVal CHR As String)
    With MWHlp
        .Sqlstrg = "SELECT TNO,TDt,TCd from Tag " + _
                " WHERE TYY='" + YY + "' AND TTC= '" + TC + _
                "' AND TCHR='" + CHR + "'" + _
                "ORDER BY TNO"
        .ColHeadg = "|VouchNo |Date        |Tag Code|"
        .ColFlds = "|TNO|TDT|TCd|"
        .ColMask = "||dd/mm/yy||"
        .RetFldNm = "TNO"
        .SearchColNum = 1
    End With
End Sub
Public Sub CustomerCode()
    With MWHlp
        .Sqlstrg = "SELECT CMCUSTCD,CMNAME,CMADD from CUSTMST where cmctg='C' order by CMCUSTCD"
        .ColHeadg = "|Code       |Name                    |Address                                |"
        .ColFlds = "|CMCUSTCD|CMNAME|CMADD|"
        .RetFldNm = "CMCUSTCD"
        .SearchColNum = 1
    End With
End Sub
Public Sub SupplierCode()
    With MWHlp
        .Sqlstrg = "SELECT CMCUSTCD,CMNAME,CMADD from CUSTMST where cmctg='S' order by CMCUSTCD"
        .ColHeadg = "|Code       |Name                    |Address                                |"
        .ColFlds = "|CMCUSTCD|CMNAME|CMADD|"
        .RetFldNm = "CMCUSTCD"
        .SearchColNum = 1
    End With
End Sub
Public Sub DiaStGrade(RMCtg As String)
    With MWHlp
        .Sqlstrg = "SELECT DCGCD,DCGDESC from DIACSGRD where DCGCTG='" + RMCtg + "' and dcgrmcode is null order by DCGCD"
        .ColHeadg = "|Code    |Desc                  |"
        .ColFlds = "|DCGCD|DCGDESC|"
        .RetFldNm = "DCGCD"
        .SearchColNum = 1
    End With
End Sub
Public Sub LabTypeSubType()
    With MWHlp
        .Sqlstrg = "SELECT labtype+','+labsubtype as Labour,labdesc from labmst " + _
                "where labsubtype is not null order by labtype,labsubtype"
        .ColHeadg = "|Type            |Description                                         |"
        .ColFlds = "|Labour|labdesc|"
        .RetFldNm = "Labour"
        .SearchColNum = 1
    End With
End Sub
Public Sub LabourType()
    With MWHlp
        .Sqlstrg = "SELECT labtype,labdesc from labmst where labsubtype is null order by labtype"
        .ColHeadg = "|Type     |Description                                                |"
        .ColFlds = "|labtype|labdesc|"
        .RetFldNm = "labtype"
        .SearchColNum = 1
    End With
End Sub
Public Sub LabSubType(Labtype As String)
    With MWHlp
        .Sqlstrg = "SELECT labsubtype, labdesc from labmst where labtype='" + Labtype + "' and labsubtype is not null order by labsubtype"
        .ColHeadg = "|Type     |Description                                                |"
        .ColFlds = "|labsubtype|labdesc|"
        .RetFldNm = "labsubtype"
        .SearchColNum = 1
    End With
End Sub
Public Sub Rmcode(Optional RMCtg As Variant, Optional RmShape As Variant)
    With MWHlp
        WREQCND = ""
        wreqcndshp = ""
        If Not IsMissing(RmShape) Then wreqcndshp = " and RMShape='" + RmShape + "'"
        If Not IsMissing(RMCtg) Then WREQCND = " and rmctg='" + RMCtg + "'" + wreqcndshp
        .Sqlstrg = "SELECT RMCODE,RMTYPE,RMSHAPE,RMCUSTCD,RMRMDEFCD from RMMST where RMType <>'ZZZZZ' " _
            + WREQCND + " order by RMCODE"
        .ColHeadg = "|Code                      |Type     |Shape    |Customer |Custom Cd|"
        .ColFlds = "|RMCODE|RMTYPE|RMSHAPE|RMCUSTCD|RMRMDEFCD|"
        .RetFldNm = "RMCODE"
        .SearchColNum = 1
    End With
End Sub
Public Sub GrdRMCode(Optional RMCtg As Variant, Optional RmShape As Variant)
With MWHlp
  WREQCND = ""
  wreqcndshp = ""
  If Not IsMissing(RmShape) Then wreqcndshp = " and RMShape='" + RmShape + "'"
  ' ***** (Jen) 10/06/00 *******
  If Not IsMissing(RMCtg) Then WREQCND = "where rmctg='" + RMCtg + "' and RMType <>'ZZZZZ'" + wreqcndshp
  ' ***** (Jen) 10/06/00 *******
  ' ***** Original Code (Before 10/06/00) *******
'  If Not IsMissing(RMCTG) Then WREQCND = "where rmctg='" + RMCTG + "' and RMType<>'GLD' and RMType <>'ZZZZZ'" + wreqcndshp
  ' ***** Original Code (Before 10/06/00) *******
  .Sqlstrg = "SELECT RMCODE,RMTYPE,RMSHAPE,RMCUSTCD,RMRMDEFCD from RMMST " _
      + WREQCND + " order by RMCODE"
  .ColHeadg = "|Code                      |Type     |Shape    |Customer |Custom Cd|"
  .ColFlds = "|RMCODE|RMTYPE|RMSHAPE|RMCUSTCD|RMRMDEFCD|"
  .RetFldNm = "RMCODE"
  .SearchColNum = 1
End With
End Sub
Public Sub CustomRMCd(Optional WCTG As Variant)
    With MWHlp
         WREQCND = ""
         If Not IsMissing(WCTG) Then WREQCND = " and rmctg='" + WCTG + "'"
        .Sqlstrg = "SELECT RMCODE,RMTYPE,RMSHAPE,RMCUSTCD from RMMST where rmtype='ZZZZZ'" + WREQCND + " order by RMCODE"
        .ColHeadg = "|RMCode                |Type    |Shape    |Customer   |"
        .ColFlds = "|RMCODE|RMTYPE|RMSHAPE|RMCUSTCD|"
        .RetFldNm = "RMCODE"
        .SearchColNum = 1
    End With
End Sub
'Public Sub PARAM(PTYPE As String, Optional PDESC30 As Variant, _
'                 Optional PDESC60 As Variant)
'Sqlstrg1 = "": Sqlstrg2 = ""
'whPDesc30 = MWLib.GetDefa("SELECT PDesc30 from Param where " + _
'           "PType='PTYPE' and PCode = '" + PTYPE + "'")
'whPDesc60 = MWLib.GetDefa("SELECT PDesc60 from Param where " + _
'           "PType='PTYPE' and PCode = '" + PTYPE + "'")
'With MWHlp
'  If Not IsMissing(PDESC30) Then
'    Sqlstrg1 = " and PDesc30= '" + PDESC30 + "'"
'    If Not IsMissing(PDESC60) Then
'      Sqlstrg2 = " and PDesc60= '" + PDESC60 + "'"
'  Else
'    If Not IsMissing(PDESC60) Then _
'      Sqlstrg2 = " and PDesc60= '" + PDESC60 + "'"
'  End If
'
'    .Sqlstrg = "SELECT PCODE,PDESC,PDESC30 from PARAM " + _
'              " where ptype = '" + PTYPE + Sqlstrg1 + Sqlstrg2 + _
'              " order by PCODE "
'    .ColHeadg = "|Code      |Description                    " + _
'                "                              |" + whPDesc30 + "|"
'    .ColFlds = "|PCODE|PDESC|PDESC30|"
'  Else
'    .Sqlstrg = "SELECT PCODE,PDESC,PDESC30 from PARAM " + _
'              " where ptype = '" + PTYPE + "' order by PCODE "
'    .ColHeadg = "|Code      |Description                    " + _
'                "                              |" + whPDesc30 + "|"
'    .ColFlds = "|PCODE|PDESC|PDESC30|"
'  End If
'    .RetFldNm = "PCODE"
'    .SearchColNum = 1
'End With
'End Sub

Public Sub PARAM(PTYPE As String, Optional PDESC30 As Variant, _
                 Optional PDesc60 As Variant)
  Sqlstrg1 = "": Sqlstrg2 = ""
  whPDesc30 = MWLib.GetDefa("SELECT PDesc30 from Param where " + _
             "PType='PTYPE' and PCode = '" + PTYPE + "'")
  
  With MWHlp
    If Not IsMissing(PDESC30) Then
      Sqlstrg1 = " and PDesc30= '" + PDESC30 + "' "
    End If
    If Not IsMissing(PDesc60) Then
      Sqlstrg2 = " and PDesc60 = '" + PDesc60 + "' "
    End If
    .Sqlstrg = "SELECT PCODE,PDESC,PDESC30 from PARAM " + _
             " where ptype = '" + PTYPE + "'" + Sqlstrg1 + Sqlstrg2 + _
             " order by PCODE "
  .ColHeadg = "|Code      |Description                    " + _
             "                              |" + whPDesc30 + "|"
  .ColFlds = "|PCODE|PDESC|PDESC30|"
  .RetFldNm = "PCODE"
  .SearchColNum = 1
End With
End Sub

Public Sub DesignCode(Optional ByVal mOldTag As Variant)
    With MWHlp
    If IsMissing(mOldTag) Then
        .Sqlstrg = "SELECT DMOldTag,DMDESGNCD,DMCUSTCD,DMDESGNDESC from DESGNMST ORDER BY DMDESGNCD"
        .ColHeadg = "|Tag|Code                            |Customer   |Desc                        |"
        .ColFlds = "|DMOldTag|DMDESGNCD|DMCUSTCD|DMDESGNDESC|"
        .SearchColNum = 2
    Else
        .Sqlstrg = "SELECT DMDESGNCD,DMCUSTCD,DMDESGNDESC from DESGNMST where DMOldTag='' ORDER BY DMDESGNCD"
        .ColHeadg = "|Code                          |Customer   |Desc                        |"
        .ColFlds = "|DMDESGNCD|DMCUSTCD|DMDESGNDESC|"
        .SearchColNum = 1
    End If
        .RetFldNm = "DMDESGNCD"
    End With
End Sub
Public Sub RMTYPE(PTYPE As String, Optional PDESC30 As Variant)
    WREQCND = ""
    If Not IsMissing(PDESC30) Then WREQCND = "and PDESC30 IN ('" + PDESC30 + "',' ')"
    With MWHlp
        .Sqlstrg = "SELECT PCODE,PDESC from PARAM where ptype='" + PTYPE + "'" + WREQCND + "  order by PCODE"
        .ColHeadg = "|Code      |Description                  |"
        .ColFlds = "|PCODE|PDESC|"
        .RetFldNm = "PCODE"
        .SearchColNum = 1
    End With
End Sub
Public Sub CombinationNo()
With MWHlp
        .Sqlstrg = "SELECT CCCOMBNO,CCCOMBCD,CCGLDKT,CCGLDCOL,CCDIAMGRD,CCCSQLY,CCFINISHTYPE from COMBCD order by CCCOMBNO"
        .ColHeadg = "|Number |Code          |Gld Kt |Gld Col|DiamGrd|CS Qly|Fin Typ|"
        .ColFlds = "|CCCOMBNO|CCCOMBCD|CCGLDKT|CCGLDCOL|CCDIAMGRD|CCCSQLY|CCFINISHTYPE|"
        .RetFldNm = "CCCOMBNO"
        .SearchColNum = 1
    End With
End Sub
Public Sub POVchNo(ByVal YY As String, ByVal TC As String, _
                    ByVal CHR As String, Optional INV As Variant)
  wsqlstrg = " SELECT POMVchNO,POMDT,POMShippingDt,POMCUSTCD " + _
                  " from POMST WHERE POMVchYY='" + YY + _
                  " ' AND POMVchTC= '" + TC + "' AND POMVchCHR='" + _
                  CHR + "' ORDER BY POMVchNO"
    With MWHlp
            .Sqlstrg = wsqlstrg
            .ColHeadg = "|Number  |Date        |Shipping Date |CustomerCd   |"
            .ColFlds = "|POMVchNO|POMDT|POMShippingDt|POMCUSTCD|"
            .ColMask = "||dd/mm/yy|dd/mm/yy||"
            .SearchColNum = 1
            .RetFldNm = "POMVchNO"
    End With
End Sub
Public Sub PODesignCd(POYy As String, POTc As String, _
                     POChr As String, PONo As Single)
  With MWHlp
    .Sqlstrg = "SELECT * from PODesgn where PODVchYY='" + POYy + "' and " + _
               " PODVchTC='" + POTc + "' and PODVchChr='" + POChr + "' and  " + _
               " PODVchNo=" + CStr(PONo) + " ORDER BY PODVchSr,PODDesgnCd"
    .ColHeadg = "|SNo.|Design Code        |No. of Flutes|"
    .ColFlds = "|PODVchSr|PODDesgnCd|PODDesgnQty|"
    .SearchColNum = 2
    .RetFldNm = "PODDesgnCd"
  End With
End Sub
Public Sub OrdNo(ByVal YY As String, ByVal TC As String, _
                ByVal CHR As String, Optional INV As Variant)
wAddCnd = ""
If Not IsMissing(INV) Then
    INV = INV + "/"
    For i = 1 To 4
        mStart = mPos + 1
        mPos = InStr(mStart, INV, "/")
        Select Case i
        Case 1
            InvYY = Mid(INV, mStart, mPos - mStart)
        Case 2
            InvTc = Mid(INV, mStart, mPos - mStart)
        Case 3
            InvChr = Mid(INV, mStart, mPos - mStart)
        Case 4
            InvNo = Val(Mid(INV, mStart, mPos - mStart))
        End Select
    Next i
    wAddCnd = " Select OMOrdNo,OMDt,OMOrdDelDt,OMCustCd,OMPriorLevel," + _
                "'I' as qOMClose from " + _
                "InvOrdDes left outer join OrdMst on " + _
                " IODOrdYy=OMOrdYy and IODOrdTc=OMOrdTc and " + _
                "IODOrdChr=OMOrdChr and IODOrdNo=OMOrdNo " + _
                "where IODInvYy='" + InvYY + "' and IODInvTc='" + _
                InvTc + "' and IODInvChr='" + InvChr + "' and IODInvNo=" + _
                CStr(InvNo) + " union "
End If
wsqlstrg = wAddCnd
wsqlstrg = wsqlstrg + " SELECT OMORDNO,OMDT,OMOrdDelDt,OMCUSTCD, " + _
                       " OMPRIORLEVEL, 'O' as qOMClose " + _
                       " from ORDMST WHERE OMORDYY='" + YY + _
                       " ' AND OMORDTC= '" + TC + "' AND OMORDCHR='" + _
                       CHR + "' ORDER BY qOMClose,OMORDNO"
    With MWHlp
        .Sqlstrg = wsqlstrg
        If Not IsMissing(INV) Then
            .ColHeadg = "|Status|Number  |Date        |Delivery Date |CustomerCd|Priority|"
            .ColFlds = "|qOMClose|OMORDNO|OMDT|OMOrdDelDt|OMCUSTCD|OMPRIORLEVEL|"
            .ColMask = "|||dd/mm/yy|dd/mm/yy|||"
            .SearchColNum = 2
        Else
            .ColHeadg = "|Number  |Date        |Delivery Date |CustomerCd|Priority|"
            .ColFlds = "|OMORDNO|OMDT|OMOrdDelDt|OMCUSTCD|OMPRIORLEVEL|"
            .ColMask = "||dd/mm/yy|dd/mm/yy|||"
            .SearchColNum = 1
        End If
        .RetFldNm = "OMORDNO"
    End With
End Sub
Public Sub VouchNo(ByVal YY As String, ByVal TC As String, ByVal CHR As String)
    With MWHlp
        wreqwhere = " WHERE CHTVOUCHYY='" + YY + "' AND CHTVOUCHTC= '" + TC + "' AND CHTVOUCHCHR='" + CHR + "'"
        .Sqlstrg = "SELECT CHTVOUCHNO,CHTDATE from CHLNTXN " + wreqwhere + " ORDER BY CHTVOUCHNO"
        .ColHeadg = "|VouchNo|Date        |"
        .ColFlds = "|CHTVOUCHNO|CHTDATE|"
        .ColMask = "||dd/mm/yy|"
        .RetFldNm = "CHTVOUCHNO"
        .SearchColNum = 1
    End With
End Sub
Public Sub StkVouchNo(ByVal YY As String, ByVal TC As String, ByVal CHR As String)
  With MWHlp
    wreqwhere = " WHERE SVOUCHYY='" + YY + "' AND SVOUCHTC= '" + TC + "' AND SVOUCHCHR='" + CHR + "'"
    .Sqlstrg = "SELECT SVOUCHNO,SDATE from STKTXN " + wreqwhere + " ORDER BY SVOUCHNO"
    .ColHeadg = "|VouchNo |Date        |"
    .ColFlds = "|SVOUCHNO|SDATE|"
    .ColMask = "||dd/mm/yy|"
    .RetFldNm = "SVOUCHNO"
    .SearchColNum = 1
  End With
End Sub
Public Sub StkVouchNoBag(ByVal YY As String, ByVal TC As String, ByVal CHR As String)
  With MWHlp
    wreqwhere = " WHERE SVOUCHYY='" + YY + "' AND SVOUCHTC= '" + TC + "' AND SVOUCHCHR='" + CHR + "'"
    .Sqlstrg = "SELECT SVOUCHNO,SDATE,(SFrChlnYY+'/'+SFrChlnChr+'/'+ ltrim(Str(SFrChlnNo))) as ChallanNo from STKTXN " + wreqwhere + " ORDER BY SVOUCHNO"
    .ColHeadg = "|VouchNo |Date        |Challan No.  |"
    .ColFlds = "|SVOUCHNO|SDATE|ChallanNo|"
    .ColMask = "||dd/mm/yy||"
    .RetFldNm = "SVOUCHNO"
    .SearchColNum = 1
  End With
End Sub
Public Sub ChallanRM(ChlnYY As String, ChlnChr As String, ChlnNo As Long)
    With MWHlp
        .Sqlstrg = "(SELECT ORMTYPE,ORMRMCODE,ORMSIEVESZ,ORMLENGTH,ORMBREADTH,ORMQTY,ORMWT from Challan " + _
                   " join OrdRM (index=OrdDesgn) on  " + _
                   " ormordyy=chorderyy and ormordtc=chordertc and ormordchr=chorderchr and ormordno=chorderno and ormordsrno=chordersrno where " + _
                   " challanyy='" + ChlnYY + "' and challanchr='" + ChlnChr + "' and challanno=" + CStr(ChlnNo) + ")" + _
                   " union (select 'M',odgldcd as ormrmcode,0 as ormsievesz, 0 as ormlength,0 as ormbreadth,0 as ormqty,odnetwt as ormwt from Challan " + _
                   " join OrdDesgn (index=Uniq) on " + _
                   " odordyy=chorderyy and odordtc=chordertc and odordchr=chorderchr and odordno=chorderno and odordsr=chordersrno where " + _
                   " challanyy='" + ChlnYY + "' and challanchr='" + ChlnChr + "' and challanno=" + CStr(ChlnNo) + ")" + _
                   " order by ormrmcode "
        .ColHeadg = "|RM Code                     |SieveSz|Length|Breadth|Quantity|Weight  |"
        .ColFlds = "|ORMRMCODE|ORMSIEVESZ|ORMLENGTH|ORMBREADTH|ORMQTY|ORMWT|"
        .RetFldNm = "ORMRMCODE"
'        .RetOth = Array("ormrmszidchpt", chpt_into, "iduom", uom_into)
        .SearchColNum = 1
    End With
End Sub
Public Sub OrdDesign(OrdYy As String, OrdTc As String, _
                     OrdChr As String, OrdNo As Single, _
                     Optional Category As Variant, _
                     Optional INV As Variant)
wAddCnd = ""
If Not IsMissing(Category) Then
  Select Case Category
  Case "O"
    wAddCnd = " and ODDesgnQty>ODExpQty "
  Case "C"
    wAddCnd = " and ODDesgnQty=ODExpQty "
  Case "A"
    wAddCnd = ""
  End Select
End If
wInvStrg = ""
wCnd = ""
If Not IsMissing(INV) Then
  INV = INV + "/"
  For i = 1 To 4
    mStart = mPos + 1
    mPos = InStr(mStart, INV, "/")
    Select Case i
    Case 1
        InvYY = Mid(INV, mStart, mPos - mStart)
    Case 2
        InvTc = Mid(INV, mStart, mPos - mStart)
    Case 3
        InvChr = Mid(INV, mStart, mPos - mStart)
    Case 4
        InvNo = Val(Mid(INV, mStart, mPos - mStart))
    End Select
  Next i
    wInvStrg = " Select ODOrdSr,ODDesgnCd,ODGldCd,ODGldCol,ODDesgnQty,ODDelDt," + _
                "'I' as qODClose from " + _
                "InvOrdDes left outer join OrdDesgn on " + _
                " IODOrdYy=ODOrdYy and IODOrdTc=ODOrdTc and " + _
                "IODOrdChr=ODOrdChr and IODOrdNo=ODOrdNo and " + _
                "IODOrdSrNo=ODOrdSr " + _
                "where IODInvYy='" + InvYY + "' and IODInvTc='" + _
                InvTc + "' and IODInvChr='" + InvChr + "' and IODInvNo=" + _
                CStr(InvNo) + " and IODOrdYy='" + OrdYy + _
                "' and ODOrdTC='" + OrdTc + "' and ODOrdChr='" + _
                OrdChr + "' AND ODOrdNo=" + CStr(OrdNo) + " union all "
    wCnd = " and not exists (Select IODDesgn from InvOrdDes where " + _
                "IODInvYy='" + InvYY + "' and IODInvTc='" + _
                InvTc + "' and IODInvChr='" + InvChr + "' and IODInvNo=" + _
                CStr(InvNo) + " and IODOrdYy='" + OrdYy + _
                "' and IODOrdTc='" + OrdTc + "' and IODOrdChr='" + _
                OrdChr + "' and IODOrdNo=" + CStr(OrdNo) + _
                " and IODDesgn=ODDesgnCd and IODOrdSrNo=ODOrdSr)"
End If
wsqlstrg = "SELECT ODORDSR,ODDESGNCD,ODGLDCD, ODGLDCOL, ODDESGNQTY,ODDELDT,'O' as qODClose from " + _
            "ORDDESGN WHERE ODORDYY='" + OrdYy + "' AND ODORDTC='" + _
            OrdTc + "' AND ODORDCHR='" + OrdChr + _
            "' AND ODORDNO=" + CStr(OrdNo) + wCnd
wsqlstrg = wInvStrg + wsqlstrg + wAddCnd + " ORDER BY qODClose,ODORDSR,ODDesgnCd"
With MWHlp
'*************************Savita(06/06/00)*****************************
  .Sqlstrg = wsqlstrg
  If Not IsMissing(INV) Then
    .ColHeadg = "|Status|Sr. No|DesignCode                 |Kt.   |Colour|Qty   |DeliveryDt|"
    .ColFlds = "|qODClose|ODORDSR|ODDESGNCD|ODGLDCD|ODGLDCOL|ODDESGNQTY|ODDELDT|"
    .ColMask = "|||||||dd/mm/yy|"
'    .ColHeadg = "|Status|Sr. No|DesignCode                 |Qty   |DeliveryDt|"
'    .ColFlds = "|qODClose|ODORDSR|ODDESGNCD|ODDESGNQTY|ODDELDT|"
'    .ColMask = "|||||dd/mm/yy|"
     .SearchColNum = 3
    Else
      .ColHeadg = "|Sr. No|DesignCode                |Kt.   |Colour|Qty   |DeliveryDt|"
      .ColFlds = "|ODORDSR|ODDESGNCD|ODGLDCD|ODGLDCOL|ODDESGNQTY|ODDELDT|"
      .ColMask = "||||||dd/mm/yy|"
'      .ColHeadg = "|Sr. No|DesignCode                |Qty   |DeliveryDt|"
'      .ColFlds = "|ODORDSR|ODDESGNCD|ODDESGNQTY|ODDELDT|"
'      .ColMask = "||||dd/mm/yy|"
'*************************Savita(06/06/00)********************************
      .SearchColNum = 2
    End If
    .RetFldNm = "ODDESGNCD"
End With
End Sub
Public Sub DesignSrNo(OrdYy As String, OrdTc As String, _
                      OrdChr As String, OrdNo As Integer, _
                      DESIGN As String, _
                      Optional ByVal Category As Variant, _
                      Optional ByVal INV As Variant)
wAddCnd = ""
If Not IsMissing(Category) Then
    Select Case Category
    Case "O"
        wAddCnd = " and ODDesgnQty>ODExpQty "
    Case "C"
        wAddCnd = " and ODDesgnQty=ODExpQty "
    Case "A"
        wAddCnd = ""
    End Select
End If
wInvStrg = ""
If Not IsMissing(INV) Then
    INV = INV + "/"
    For i = 1 To 4
        mStart = mPos + 1
        mPos = InStr(mStart, INV, "/")
        Select Case i
        Case 1
            InvYY = Mid(INV, mStart, mPos - mStart)
        Case 2
            InvTc = Mid(INV, mStart, mPos - mStart)
        Case 3
            InvChr = Mid(INV, mStart, mPos - mStart)
        Case 4
            InvNo = Val(Mid(INV, mStart, mPos - mStart))
        End Select
    Next i
    wInvStrg = " Select IODOrdSrNo as ODOrdSr,'I' as qODClose from " + _
                "InvOrdDes where IODInvYy='" + InvYY + _
                "' and IODInvTc='" + InvTc + "' and IODInvChr='" + _
                InvChr + "' and IODInvNo=" + CStr(InvNo) + _
                " and IODOrdYy='" + OrdYy + "' and IODOrdTC='" + _
                OrdTc + "' and IODOrdChr='" + OrdChr + _
                "' and IODOrdNo=" + CStr(OrdNo) + " AND IODDesgn='" + _
                DESIGN + "' union "
    wCnd = " and not exists (Select IODOrdSrNo from InvOrdDes where " + _
                "IODInvYy='" + InvYY + "' and IODInvTc='" + _
                InvTc + "' and IODInvChr='" + InvChr + "' and IODInvNo=" + _
                CStr(InvNo) + " and IODOrdYy='" + OrdYy + _
                "' and IODOrdTc='" + OrdTc + "' and IODOrdChr='" + _
                OrdChr + "' and IODOrdNo=" + CStr(OrdNo) + _
                " and IODDesgn='" + DESIGN + "' and IODOrdSrNo=ODOrdSr)"
End If
wsqlstrg = "SELECT ODORDSR,'O' as qODClose from ORDDESGN WHERE " + _
           "ODORDYY='" + OrdYy + "' AND ODORDTC='" + OrdTc + _
           "' AND ODORDCHR='" + OrdChr + "' AND ODORDNO=" + _
           CStr(OrdNo) + " AND ODDESGNCD='" + _
           DESIGN + "'" + wCnd
wsqlstrg = wInvStrg + wsqlstrg + wAddCnd + " ORDER BY qODClose,ODORDSR"
    With MWHlp
        .Sqlstrg = wsqlstrg
        If Not IsMissing(INV) Then
            .ColHeadg = "|Status|SrNo   |"
            .ColFlds = "|qODClose|ODORDSR|"
            .SearchColNum = 2
        Else
            .ColHeadg = "|SrNo   |"
            .ColFlds = "|ODORDSR|"
            .SearchColNum = 1
        End If
        .RetFldNm = "ODORDSR"
    End With
End Sub
Public Sub ChallanNo(ByVal YY As String, ByVal CHR As String, _
                    Optional Category As Variant, _
                    Optional ORD As Variant)
If IsMissing(Category) Then Category = "A"
Select Case Category
Case "O"
    wAddCnd = " and CHCloseTag=' ' "
Case "C"
    wAddCnd = " and CHCloseTag='C' "
Case "A"
    wAddCnd = ""
End Select
If Not IsMissing(ORD) Then
    ' **** (Jen) 30/06/00 *****
    OrdYy = "": OrdTc = "": OrdChr = "": mOrdNo = 0: mOrdSrNo = 0
    ' **** (Jen) 30/06/00 *****
    ORD = ORD + "/"
    For i = 1 To 5
        mStart = mPos + 1
        mPos = InStr(mStart, ORD, "/")
        Select Case i
        Case 1
            If mPos > 0 Then OrdYy = Mid(ORD, mStart, mPos - mStart)
        Case 2
            If mPos > 0 Then OrdTc = Mid(ORD, mStart, mPos - mStart)
        Case 3
            If mPos > 0 Then OrdChr = Mid(ORD, mStart, mPos - mStart)
        Case 4
            If mPos > 0 Then mOrdNo = Val(Mid(ORD, mStart, mPos - mStart))
        Case 5
            If mPos > 0 Then mOrdSrNo = Val(Mid(ORD, mStart, mPos - mStart))
        End Select
    Next i
    ' **** (Jen) 30/06/00 *****
    If OrdYy <> "" And OrdTc <> "" And OrdChr <> "" And mOrdNo <> 0 And mOrdSrNo = 0 Then
      wAddCnd = wAddCnd + _
              " and CHORDERYY='" + OrdYy + "' and CHORDERTC='" + OrdTc + _
              "' and CHORDERCHR='" + OrdChr + "' and CHORDERNO=" + _
              CStr(mOrdNo)
    End If
    If OrdYy <> "" And OrdTc <> "" And OrdChr <> "" And mOrdNo <> 0 And mOrdSrNo <> 0 Then
      wAddCnd = wAddCnd + _
              " and CHORDERYY='" + OrdYy + "' and CHORDERTC='" + OrdTc + _
              "' and CHORDERCHR='" + OrdChr + "' and CHORDERNO=" + _
              CStr(mOrdNo) + " and CHORDERSRNO=" + CStr(mOrdSrNo)
    End If
    ' **** (Jen) 30/06/00 *****
    
    ' **** Original Code (Before 30/06/00) *****
'    wAddCnd = wAddCnd + _
'        " and CHORDERYY='" + OrdYY + "' and CHORDERTC='" + OrdTC + _
'        "' and CHORDERCHR='" + OrdChr + "' and CHORDERNO=" + _
'        CStr(mOrdNo) + " and CHORDERSRNO=" + CStr(mOrdSrNo)
    ' **** Original Code (Before 30/06/00) *****
    
End If
wsqlstrg = "SELECT CHALLANNO,CHCURRQTY,CHDESIGN,CHORDERYY+'/'+" + _
           "CHORDERTC+'/'+CHORDERCHR+'/'+LTRIM(STR(CHORDERNO)) AS " + _
           "CHORDER,CHCURRLOC,CHCURRWORK from CHALLAN " + _
           "WHERE CHALLANYY='" + YY + "' AND CHALLANCHR='" + CHR + "'"
wsqlstrg = wsqlstrg + wAddCnd + " ORDER BY CHALLANNO"
    With MWHlp
        .Sqlstrg = wsqlstrg
        .ColHeadg = "|ChallanNo |Quantity|Design              |OrderNumber          |Location|Associate|"
        .ColFlds = "|CHALLANNO|CHCURRQTY|CHDESIGN|CHORDER|CHCURRLOC|CHCURRWORK|"
        .RetFldNm = "CHALLANNO"
        .SearchColNum = 1
    End With
End Sub
Public Sub Worker(Optional Location As String)
    With MWHlp
        .Sqlstrg = "SELECT PCODE,PDESC from PARAM where ptype='" + PTYPE + "' AND PDESC30='" + Location + "' order by PCODE"
        .ColHeadg = "|Code      |Description                                      |"
        .ColFlds = "|PCODE|PDESC|"
        .RetFldNm = "PCODE"
        .SearchColNum = 1
    End With
End Sub
Public Sub CustomCtg()
    With MWHlp
        '***** (Jen) 10/06/00 *******
        .Sqlstrg = "SELECT PCODE,PDESC from PARAM where ptype='RMCTG'  " + _
                   "and pcode <>'A' and pcode <>'O'  " + _
                   "order by PCODE"
        '***** (Jen) 10/06/00 *******
                   
'***** Original Code (Before 10/06/00) *******
'        .Sqlstrg = "SELECT PCODE,PDESC from PARAM where ptype='RMCTG'  " + _
'                   "and pcode <>'M' and pcode <>'A' and pcode <>'O'  " + _
'                   "order by PCODE"
'***** Original Code (Before 10/06/00) *******
        .ColHeadg = "|Code      |Description                                      |"
        .ColFlds = "|PCODE|PDESC|"
        .RetFldNm = "PCODE"
        .SearchColNum = 1
    End With
End Sub
Public Sub LossLoc()
    With MWHlp
        .Sqlstrg = "SELECT PDESC30,PCODE from PARAM where ptype='LOC' order by PDESC30"
        .ColHeadg = "|LossLocation|Location|"
        .ColFlds = "|PDESC30|PCODE|"
        .RetFldNm = "PDESC30"
        .SearchColNum = 1
    End With
End Sub
Public Sub StoneCtg()
    With MWHlp
        .Sqlstrg = "SELECT PCODE,PDESC from PARAM where ptype='RMCTG'  and pcode <>'M' and pcode <>'A' and pcode <>'O' and pcode <>'X'  order by PCODE"
        .ColHeadg = "|Code      |Description                                      |"
        .ColFlds = "|PCODE|PDESC|"
        .RetFldNm = "PCODE"
        .SearchColNum = 1
    End With
End Sub
Public Sub InvoiceNo(ByVal YY As String, ByVal TC As String, ByVal CHR As String)
    With MWHlp
        wreqwhere = " WHERE IHDInvYY='" + YY + "' AND IHDInvTC= '" + TC + "' AND IHDInvChr='" + CHR + "'"
        .Sqlstrg = "SELECT IHDInvNo,IHDInvDt,IHDCustCd from INVHD" + wreqwhere + " ORDER BY IHDInvNo"
        .ColHeadg = "|Number|Date        |CustomerCd|"
        .ColFlds = "|IHDINVNO|IHDINVDT|IHDCUSTCD|"
        .ColMask = "||dd/mm/yy||"
        .RetFldNm = "IHDINVNO"
        .SearchColNum = 1
    End With
End Sub
Public Sub CostingNo(ByVal YY As String, ByVal TC As String, ByVal CHR As String)
    With MWHlp
        wreqwhere = " WHERE CMCOSTYY='" + YY + "' AND CMCOSTTC= '" + TC + "' AND CMCOSTCHR='" + CHR + "'"
        .Sqlstrg = "SELECT CMCOSTNO,CMDT,CMCUSTCD from COSTMST" + wreqwhere + " ORDER BY CMCOSTNO"
        .ColHeadg = "|Number|Date        |CustomerCd|"
        .ColFlds = "|CMCOSTNO|CMDT|CMCUSTCD|"
        .ColMask = "||dd/mm/yy||"
        .RetFldNm = "CMCOSTNO"
        .SearchColNum = 1
    End With
End Sub
'**** Old Code 08/08/00 *****
'Public Sub CstmVouch(ByVal CstmCd As String)
'wSqlStr = "Select CSTKVouchYy+'/'+CSTKVouchTc+'/'+CSTKVouchChr+'/'+" + _
'        "ltrim(str(CSTKVouchNo))+'/'+ltrim(str(CSTKSrNo)) as " + _
'        "qVouch,CSTKBOE,CSTKBOEDt,CSTKCustRmCd,CSTKRate,(CSTKWt-CSTKDespWt) as " + _
'        "qBalWt,CSTKUseBef from CustMStk where CSTKClose<>'Y' and " + _
'        "CSTKBM='B' and CSTKCustRmCd='" + CstmCd + _
'        "' order by CSTKUseBef,qVouch"
'With MWHlp
'    .Sqlstrg = wSqlStr
'    .ColHeadg = "|Voucher             |Custom Bill           |Bill Date   |CstmRmCd|Rate   |Bal Wt     |Use Before  |"
'    .ColFlds = "|qVouch|CSTKBOE|CSTKBOEDt|CSTKCustRmCd|CSTKRate|qBalWt|CSTKUseBef|"
'    .ColMask = "|||dd/mm/yyyy||###0.00|#####0.000|dd/mm/yyyy|"
'    .RetFldNm = "qVouch"
'    .SearchColNum = 1
'End With
'End Sub
'**** Old Code 08/08/00 *****
'**** New Code 08/08/00 *****
Public Sub CstmVouch(ByVal CstmCd As String, Optional Ctg As Variant)
If Not IsMissing(Ctg) Then
  If Ctg = "O" Then wAddCnd = "and CSTKWt>CSTKDespWt" Else wAddCnd = ""
Else
  wAddCnd = ""
End If
wSqlStr = "Select CSTKVouchYy+'/'+CSTKVouchTc+'/'+CSTKVouchChr+'/'+" + _
        "ltrim(str(CSTKVouchNo))+'/'+ltrim(str(CSTKSrNo)) as " + _
        "qVouch,CSTKBOE,CSTKBOEDt,CSTKCustRmCd,CSTKQty,CSTKRate," + _
        "(CSTKWt-CSTKDespWt) as qBalWt,CSTKUseBef " + _
        "from CustMStk where CSTKClose<>'Y' and " + _
        "CSTKBM='B' " + wAddCnd + " and CSTKCustRmCd='" + CstmCd + _
        "' order by CSTKUseBef, CSTKBOEDt, qVouch "
wRMCtg = MWLib.GetDefa("Select RmCtg from RmMst where RmCode='" + _
                        CstmCd + "' and RmType='ZZZZZ'")
With MWHlp
    .Sqlstrg = wSqlStr
    Select Case wRMCtg
    Case Is = "D", "C", "Z", "X"
        If wRMCtg = "X" Then wPcsHdg = "Qty  " Else wPcsHdg = "Avg Pcs"
        .ColHeadg = "|Voucher             |Custom Bill           |Bill Date       |CstmRmCd|" + wPcsHdg + "|Rate   |Bal Wt     |Use Before   |"
        .ColFlds = "|qVouch|CSTKBOE|CSTKBOEDt|CSTKCustRmCd|CSTKQty|CSTKRate|qBalWt|CSTKUseBef|"
        .ColMask = "|||dd/mm/yyyy||####0|###0.00|#####0.000|dd/mm/yyyy|"
    Case Else
        .ColHeadg = "|Voucher             |Custom Bill           |Bill Date       |CstmRmCd|Rate   |Bal Wt     |Use Before   |"
        .ColFlds = "|qVouch|CSTKBOE|CSTKBOEDt|CSTKCustRmCd|CSTKRate|qBalWt|CSTKUseBef|"
        .ColMask = "|||dd/mm/yyyy||###0.00|#####0.000|dd/mm/yyyy|"
    End Select
    .RetFldNm = "qVouch"
    .SearchColNum = 1
End With
End Sub
'**** New Code 08/08/00 *****

Public Sub InvKeyVal(ByVal INV As String)
INV = INV + "/"
For i = 1 To 4
    mStart = mPos + 1
    mPos = InStr(mStart, INV, "/")
    Select Case i
    Case 1
        InvYY = Mid(INV, mStart, mPos - mStart)
    Case 2
        InvTc = Mid(INV, mStart, mPos - mStart)
    Case 3
        InvChr = Mid(INV, mStart, mPos - mStart)
    Case 4
        InvNo = Val(Mid(INV, mStart, mPos - mStart))
    End Select
Next i
wSqlStr = "Select distinct IRMKeyVal from InvODRm where IRMInvYy='" + _
        InvYY + "' and IRMInvTc='" + InvTc + "' and IRMInvChr='" + _
        InvChr + "' and IRMInvNo=" + CStr(InvNo)
With MWHlp
    .Sqlstrg = wSqlStr
    .ColHeadg = "|Key Value             |"
    .ColFlds = "|IRMKeyVal|"
    .RetFldNm = "IRMKeyVal"
    .SearchColNum = 1
End With
End Sub
Public Sub MenuCode()
    With MWHlp
        .Sqlstrg = "SELECT UAMNUCD,UAMNUDESC from USRACCESS where UAUSRCD='MIC' order by UAMNUCD"
        .ColHeadg = "|MenuCode       |Description                            |"
        .ColFlds = "|UAMNUCD|UAMNUDESC|"
        .RetFldNm = "UAMNUCD"
        .SearchColNum = 1
    End With
End Sub
Public Sub BOENo()
wSqlStr = "Select distinct CSTKBOE, CSTKBOEDt, CSTKCustRmCd, CSTKRate, " + _
        "CSTKUseBef from CustMStk where CSTKBM='B' " + _
        " order by CSTKUseBef, CSTKBOE"
With MWHlp
    .Sqlstrg = wSqlStr
    .ColHeadg = "|Custom Bill           |Bill Date   |CstmRmCd|Rate   |Use Before  |"
    .ColFlds = "|CSTKBOE|CSTKBOEDt|CSTKCustRmCd|CSTKRate|CSTKUseBef|"
    .ColMask = "||dd/mm/yyyy||###0.00|dd/mm/yyyy|"
    .RetFldNm = "CSTKBOE"
    .SearchColNum = 1
End With
End Sub
Public Sub SRStmntNo(ByVal YY As String, ByVal TC As String, ByVal CHR As String)
    With MWHlp
        wreqwhere = " WHERE SRSTMNTYY='" + YY + "' AND SRSTMNTTC= '" + TC + "' AND SRSTMNTCHR='" + CHR + "'"
        .Sqlstrg = "SELECT SRSTMNTNO,SRDT,SRLOC from stkreco " + wreqwhere + " ORDER BY SRSTMNTNO"
        .ColHeadg = "|Number|Date        |CustomerCd|"
        .ColFlds = "|SRSTMNTNO|SRDT|SRLOC|"
        .ColMask = "||dd/mm/yy||"
        .RetFldNm = "SRSTMNTNO"
        .SearchColNum = 1
    End With
End Sub
Public Property Let MultiSelect(ByVal vNewValue As Boolean)
    MWHlp.MultiSelect = vNewValue
End Property


