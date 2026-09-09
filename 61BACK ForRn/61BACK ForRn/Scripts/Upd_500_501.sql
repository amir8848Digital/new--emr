/****** Object: Trigger dbo.InvFgd_NR_INUPDL Script Date: 03/10/05 ******/   
/****** Object: Trigger dbo.InvFgd_NR_INUPDL Script Date: 03/10/05 ******/   
ALTER Trigger
[dbo].[InvFgd_INUPDL] ON [dbo].[InvFgd]  
  FOR INSERT,UPDATE,DELETE   NOT FOR REPLICATION
AS   
BEGIN   
 /* Trigger Version 3.8.0 */  
 DECLARE @ROW INT   
 DECLARE @mode char(1)   
 DECLARE @errmsg varchar(255)   
 DECLARE @wTcTyp varchar(3)  
 DECLARE @wEntTyp varchar(8)  
 DECLARE @ReplAgt BIT    
 DECLARE @wCurrPrtn VarChar(1)   
   
 IF @@ROWCOUNT=0 RETURN   
   
 If Exists(Select 'x' From TrgDisa Where TdTblNm = 'InvFgd' )   
  RETURN   
   
 select @mode=case (select count('x') from inserted) when 0 then 'D'   
 else   
  case (select count('x') from deleted) when 0 then 'A' else 'C' end   
 end   
   
 /*** Storing Tc Type In The Variable @wTcTyp ***/  
 if @mode in ('A','C')  
 begin  
  select @wTcTyp= max(PValue) from inserted join Param on PTyp= 'TC' and PMCd= IfTc and PSCd= ''  
 end  
  
 if @mode in ('D')  
 begin  
  select @wTcTyp= max(PValue) from deleted join Param on PTyp= 'TC' and PMCd= IfTc and PSCd= ''  
 end  
   
 /*** Storing Entry Type In The Variable @wEntTyp ***/  
 select @wEntTyp= PValue1 from Param  where PTyp= 'TCTYP' and PMCd= @wTcTyp and PSCd= ''  
 Select @wCurrPrtn='C'  
  
 /* **** Checking Whether a Normal Transaction or Replicated One **** */  
 Set @ReplAgt=Convert(BIT,SessionProperty('Replication_Agent'))  

 If @ReplAgt = 1         Return
      
 
 /******* Sachin 3.02.0 [30/11/07] - Purge Partition (Speed Optimization) */  
 if @mode='C'  
 begin  
  Select @row=count('x') From Inserted Where IfPrtKey<>@wCurrPrtn  
  if @row> 0   
  begin  
   select @errmsg= 'Cannot Edit Records of Previous Partition'  
   goto ErrHndlr  
  end  
 End  
  
 if @mode='A'  
 Begin  
  Select @row=count('x') From Inserted Where IfPrtKey <> @wCurrPrtn   
  if @row> 0   
  begin  
   select @errmsg= 'Cannot Insert Partition Key Other Than ''C'''  
   goto ErrHndlr  
  end  
 end   
  
 If @Mode='D'   
 Begin   
  select @row= count('x') from deleted del where del.IfPrtKey<>@wCurrPrtn  
  if @row> 0   
  begin  
   select @errmsg= 'Cannot Delete Records From Previous Partition'  
   goto ErrHndlr  
  end  
 End  
 /******* Sachin 3.02.0 [30/11/07] - Purge Partition (Speed Optimization) */  
   
 If  (@ReplAgt<>1)  
 Begin  
  /******Cannot Insert/Update/Delete  if Crm database /Emr database with TrdCocd not in sdCoCd ****/  
  /***if SysId=Emr Record not found then normal behaviour i.e can insert/update/delete****/  
  if @mode in ('A','C')  
  Begin   
  /* Current Partition */  
  select @ROW =Count('x')  From inserted where IfPrtKey=@wCurrPrtn and   
  Exists ( select 'x' From SysDb where (SysId='EMR' and (SysDb='LTP' Or ','+sdCocd+',' Not  like '%,'+ IfCoCd +',%'))   
   or (SysId='CRM' and SysDb='WEB'))  
  if @Row > 0    
   Begin   
    select @errmsg= 'Cannot Insert /Update '   
    goto ErrHndlr   
   end   
  End   
  
  if @mode in ('D')  
  Begin   
  /* Current Partition */  
  select @ROW =Count('x')  From Deleted where IfPrtKEy=@wCurrPrtn and   
  Exists ( select 'x' From SysDb where (SysId='EMR' and (SysDb='LTP' Or ','+sdCocd+',' Not  like '%,'+ IfCoCd +',%'))  
   or (SysId='CRM' and SysDb='WEB'))  
  if @Row > 0    
   Begin   
    select @errmsg= 'Cannot Delete '   
    goto ErrHndlr   
   end   
  End   
   
  if @mode = 'D'    Begin  
   /******* (BCls= 'N') Bag is again opened for transaction ******/  
   /* Current Partition */  
  /* update Bag set  
   BCls= 'N', BClsDt= '01/01/80'    
   from deleted where @wTcTyp= 'IN' and IfPrtKey=@wCurrPrtn   
    and BIdNo=IfBIdNo and BPrtKey=IfPrtKey and BCoCd= IfCoCd   
    and BYy= IfFdBYy and BChr= IfFdBChr and BNo= IfFdBNo    
   */
   /******* If @wTcTyp= 'IN' then update FdDespQty of Fgd ******/  
   /* Current Partition */  
   update Fgd set   
   FdDespQty= FdDespQty- IsNull((Select sum(del.IfQty) from deleted del where   
    @wTcTyp= 'IN' and del.IfPrtKey=FdPrtKey and del.IfFdIdNo=FdIdNo and del.IfCoCd= FdCoCd and del.IfFdTc= FdTc   
    and del.IfFdYy= FdYy and del.IfFdChr= FdChr and del.IfFdNo= FdNo   
    and del.IfFdSr= FdSr), 0)   
   from deleted where IfPrtkey=@wCurrPrtn and @wTcTyp= 'IN' and FdPrtKey=IfPrtKey   
    and FdIdNo=IfFdIdNo and FdCoCd= IfCoCd   
    and FdTc= IfFdTc and FdYy= IfFdYy and FdChr= IfFdChr and FdNo= IfFdNo and FdSr= IfFdSr  
   
   /******* If @wTcTyp= 'CTB' then update FdCTBDespQty of Fgd ******/  
   /* Current Partition */  
   update Fgd set   
   FdCTBDespQty= FdCTBDespQty- IsNull((Select sum(del.IfQty) from deleted del where   
    @wTcTyp= 'CTB' and del.IfPrtKey=FdPrtKey and del.IfFdIdNo=FdIdNo and del.IfCoCd= FdCoCd and del.IfFdTc= FdTc   
    and del.IfFdYy= FdYy and del.IfFdChr= FdChr and del.IfFdNo= FdNo   
    and del.IfFdSr= FdSr), 0)   
   from deleted where @wTcTyp= 'CTB' and IfPrtKey=@wCurrPrtn and FdIdNo=IfFdIdNo and FdCoCd= IfCoCd and FdTc= IfFdTc   
    and FdYy= IfFdYy and FdChr= IfFdChr and FdNo= IfFdNo and FdSr= IfFdSr  
   
  End   

	update Bag set   
	   BCls= (Case when IsNull(ins.IfSr,0)>0 Then 'Y' Else 'N' End),   
	   BClsDt= IsNull((Select InDt From InvHd Where InPrtKey=ins.IfPrtKey and InIdNo=ins.IfInIdNo and InCoCd= ins.IfCoCd And InTc = ins.IfTc   
		 And InYy = ins.IfYy And InChr= ins.IfChr And InNo= ins.IfNo),'01/01/80')   
	   from  Bag b
		Left Outer Join inserted ins On	b.BIdNo=ins.IfBIdNo and b.BCoCd= ins.IfCoCd and b.BYy= ins.IfFdBYy   and b.BChr= ins.IfFdBChr and b.BNo= ins.IfFdBNo and ins.IfPrtKey=@wCurrPrtn 
		Left Outer Join Deleted del On	b.BIdNo=del.IfBIdNo and b.BCoCd= del.IfCoCd and b.BYy= del.IfFdBYy   and b.BChr= del.IfFdBChr and b.BNo= del.IfFdBNo and del.IfPrtKey=@wCurrPrtn 
	 where @wTcTyp= 'IN' /*and Bag.BIdNo=b.BIdNo and Bag.BCoCd= b.BCoCd and Bag.BYy= b.BYy   and Bag.BChr= b.Chr and Bag.BNo= b.BNo */
						and (IsNull(ins.IfSr,-1)<>IsNull(del.IfSr,-1) and (@Mode<>'A' Or IsNull(ins.IfSr,-1)>0))
						OR (IsNull(ins.IfSr,0)>0 and IsNull(ins.IfSr,-1)=IsNull(del.IfSr,-1))

   
  if @mode in ('A')  
  Begin  
   /******* update FdDespQty of Fgd ******/  
   /* Current Partition */  
   update Fgd set   
   FdDespQty= FdDespQty+ IsNull((Select sum(ins.IfQty) from inserted ins where   
    @wTcTyp= 'IN' and ins.IfPrtKey=FdPrtKey and ins.IfFdIdNo=FdIdNo and ins.IfCoCd= FdCoCd and ins.IfFdTc= FdTc   
    and ins.IfFdYy= FdYy and ins.IfFdChr= FdChr and ins.IfFdNo= FdNo and ins.IfFdSr= FdSr), 0)   
   from inserted where @wTcTyp= 'IN' and IfPrtKey=@wCurrPrtn and FdPrtKey=IfPrtKey and FdIdNo=IfFdIdNo  
    and FdCoCd= IfCoCd and FdTc= IfFdTc and FdYy= IfFdYy and FdChr= IfFdChr and FdNo= IfFdNo and FdSr= IfFdSr  
   
   /******* update FdDespQty of Fgd ******/  
   /* Current Partition */  
   update Fgd set   
   FdCTBDespQty= FdCTBDespQty+ IsNull((Select sum(ins.IfQty) from inserted ins where   
    @wTcTyp= 'CTB' and ins.IfPrtKey=FdPrtKey and ins.IfFdIdNo=FdIdNo and ins.IfCoCd= FdCoCd and ins.IfFdTc= FdTc   
    and ins.IfFdYy= FdYy and ins.IfFdChr= FdChr and ins.IfFdNo= FdNo   
    and ins.IfFdSr= FdSr), 0)   
   from inserted where @wTcTyp= 'CTB' and IfPrtKey=@wCurrPrtn and FdPrtKey=IfPrtKey and FdIdNo=IfFdIdNo  
    and FdCoCd= IfCoCd and FdTc= IfFdTc and FdYy= IfFdYy and FdChr= IfFdChr and FdNo= IfFdNo and FdSr= IfFdSr  
   
   /******* (BCls= 'Y') Bag is closed for transaction ******/  
   /* Current Partition */  
  /*  update Bag set   
   BCls= 'Y',   
   BClsDt= (Select InDt From InvHd Where InPrtKey=IfPrtKey and InIdNo=IfInIdNo and InCoCd= IfCoCd And InTc = IfTc   
     And InYy = IfYy And InChr= IfChr And InNo= IfNo)   
   from inserted where IfPrtKey=@wCurrPrtn and @wTcTyp= 'IN' and BIdNo=IfBIdNo and BCoCd= IfCoCd and BYy= IfFdBYy   
    and BChr= IfFdBChr and BNo= IfFdBNo    */
  End  
 End  
  
 /******** Only One TcTyp records can be Inserted at a time.********/  
 /* Current Partition */   
 select @Row= Count(distinct(PValue)) from inserted Join param On  
  PTyp= 'TC' and PMCd= IfTc and PSCd= '' Where IfPrtKey=@wCurrPrtn   
 If @Row> 1  
 begin  
  Select @ErrMsg= 'Transactions Of Only One TcType Can Be Acted On'  
    GoTo ErrHndlr  
 end  
  
 /******** Only One TcTyp records can be Deleted at a time. *********/  
 /* Current Partition */  
 select @Row= Count(distinct(PValue)) from deleted Join param on  
  PTyp= 'TC' and PMcd= IfTc and PSCd= '' Where IfPrtKEy=@wCurrPrtn   
 If @Row> 1  
 begin  
  Select @ErrMsg= 'Transactions Of Only One TcType Can Be Acted On'  
  GoTo ErrHndlr  
 end  
  
  /******** Cannot Add/ Delete/ Change If Invoice Is Closed ********/   
 /* Current Partition */  
  if @mode in('A', 'C', 'D')  
  Begin  
  select @Row=Count('x') from deleted join InvHd on InPrtKEy=IfPrtKey and InIdNo=IfInIdNo   
   and InCoCd= IfCoCd and InTc= IfTc and InYy= IfYy and InChr= IfChr and InNo= IfNo   
  where IfPrtKEy=@wCurrPrtn and InCls= 'Y'   
  If @Row > 0  
  begin   
   Select @ErrMsg='Cannot Delete/ Change As Invoice Is Closed'  
   GoTo ErrHndlr  
  end  
  
  /* Current Partition */  
  select @Row=Count('x') from inserted   
   join InvHd on InPrtKey=IfPrtKey and InIdNo=IfInIdNo and InCoCd= IfCoCd and InTc= IfTc and InYy= IfYy and InChr= IfChr and InNo= IfNo   
  where IfPrtKey=@wCurrPrtn and InCls= 'Y'   
  If @Row > 0  
  begin   
   Select @ErrMsg='Cannot Add/ Change As Invoice Is Closed'  
   GoTo ErrHndlr  
  end  
 End  
  
  if @mode = 'C'  
  Begin  
  /******** Key Cannot Be Changed (Key without IfSr) ********/   
  /* Current Partition */  
  select @Row=Count('x') from deleted del where del.IfPrtKey=@wCurrPrtn and  
    not exists (Select 'x' from inserted ins where ins.IfPrtKey=del.IfPrtKey and ins.IfIdNo=del.IfIdNo   
     and ins.IfCoCd= del.IfCoCd and ins.IfTc= del.IfTc and ins.IfYy= del.IfYy and ins.IfChr= del.IfChr   
     and ins.IfNo= del.IfNo and ins.IfSrNo= del.IfSrNo)  
  If @Row > 0  
  begin   
   Select @ErrMsg='Cannot Change Company Code/ Tc/Yy/Chr/No/SrNo'  
   GoTo ErrHndlr  
  end  
  
  /******** Cannot change IfSr if record exists in InvDsg ********/   
  /* Current Partition */  
	/**** Manali 3.8.0 - InvDsg Check was commented ealier - Now uncommented *****/
  select @Row= Count('x') from deleted del where del.IfPrtKEy=@wCurrPrtn and del.IfSr<> 0 and   
    not exists (Select 'x' from inserted ins   
      where ins.IfPrtKey=del.IfPrtKey and ins.IfIdNo=del.IfIdNo and ins.IfCoCd= del.IfCoCd   
       and ins.IfTc= del.IfTc and ins.IfYy= del.IfYy   
       and ins.IfChr= del.IfChr and ins.IfNo= del.IfNo   
       and ins.IfSr= del.IfSr and ins.IfSrNo= del.IfSrNo)   
     and exists (Select 'x' from InvDsg where IdCoCd= del.IfCoCd   
     and IdTc= del.IfTc and IdYy= del.IfYy and IdChr= del.IfChr   
     and IdNo= del.IfNo and IdSr= del.IfSr) 
  If @Row > 0  
  begin   
   Select @ErrMsg= 'Cannot Change Inv Voucher Sr'  
   GoTo ErrHndlr  
  end  
  
  /******** Cannot Change Fg Vouch, Bag No, Qty, Grs Wt ********/   
  select @Row=Count('x') from deleted del where del.IfPrtKey=@wCurrPrtn and  
    not exists (Select 'x' from inserted ins where ins.IfPrtKey=del.IfPrtKey and ins.IfIdNo=del.IfIdNo  
     and ins.IfCoCd= del.IfCoCd and ins.IfTc= del.IfTc and ins.IfYy= del.IfYy and ins.IfChr= del.IfChr   
     and ins.IfNo= del.IfNo and ins.IfSrNo= del.IfSrNo and ins.IfFdTc= del.IfFdTc   
     and ins.IfFdYy= del.IfFdYy and ins.IfFdChr= del.IfFdChr and ins.IfFdNo= del.IfFdNo   
     and ins.IfFdSr= del.IfFdSr and ins.IfFdBYy= del.IfFdBYy and ins.IfFdBChr= del.IfFdBChr   
     and ins.IfFdBNo= del.IfFdBNo and ins.IfQty= del.IfQty and ins.IfGrsWt= del.IfGrsWt)  
  If @Row > 0  
  begin   
   Select @ErrMsg='Cannot Change Bag Number/ Finished Goods Vouch/ Qty/ Grs Wt'  
   GoTo ErrHndlr  
  end  
 End  
  
 if @mode= 'A'  
 Begin  
		/* Sachin 3.2.2 - Bag Opn Dt should be Less than Inv Date */
		/* *** Check Below Commented as it was not required [Discussion with Poonam][18/10/08] ***
		select @Row= Count('x') from inserted
			Join InvHd On InPrtKey=IfPrtKey and InIdNo=IfInIdNo and InCoCd=IfCoCd and InTc=IfTc and InYy=IfYy and InChr=IfChr and InNo=IfNo
			join Bag on IfBIdNo=BIdNo and IfCoCd= BCoCd and IfFdBYy= BYy and IfFdBChr= BChr and IfFdBNo= BNo 
					and IfPrtKey=BPrtKey
			where IfPrtKey=@wCurrPrtn and InDt<BOpnDt
		If @Row > 0
		begin
			Select @ErrMsg= 'Bag Cannot Be Invoiced Before Bag Opening Date'
			GoTo ErrHndlr
		end
		*/		

  /****** If record exists in InvDsg, no records with Sr=0 will be allowed to be added into table InvFgd ******/  
  /* Current Partition */  
  select @row=count('x') from inserted where IfPrtKey=@wCurrPrtn and IfSr= 0 and   
    exists (select 'x' from InvDsg where IdPrtKey=IfPrtKey and IdInIdNo=IfInIdNo and IdCoCd= IfCoCd   
    and IdTc= IfTc and IdYy= IfYy and IdChr= IfChr and IdNo= IfNo)  
  if @row>0   
  begin  
   select @errmsg= 'Cannot Enter Record With Sr = 0 (i.e. When Inv Design Record Exists)'  
   goto ErrHndlr  
  end  
  
  /****** Cannot add Inv Bag record if Inv Rm or Inv Lab record exists ******/  
  /* Current Partition */  
  select @row= count('x') from inserted where IfPrtKey=@wCurrPrtn and  
   (exists(Select 'x' from InvRm where IrPrtKey=IfPrtKey  
     and IrCoCd= IfCoCd and IrTc= IfTc and IrYy= IfYy and IrChr= IfChr and IrNo= IfNo and IrSr= IfSr) or  
   exists(Select 'x' from InvLab where IlPrtKey=IfPRtKey and iLCoCd= IfCoCd and iLTc= IfTc and iLYy= IfYy and   
     iLChr= IfChr and iLNo= IfNo and iLSr= IfSr))   
  if @row> 0   
  begin  
   select @errmsg= 'Cannot Add As Invoice Raw Material/ Labour Record Exists'  
   goto ErrHndlr  
  end  
  
  /****** For TcTyp= 'IN', Total Bag Qty in the InvFgd table should not exceed the OdOrdQty of the corresponding Export Order Design ******/  
  /* Current Partition */  
  select @row= count('x') from inserted ins join Fgd Fd1 on   
    FdPrtKey=IfPrtKey and FdIdNo=IfFdIdNo and FdCoCd= IfCoCd and FdTc= IfFdTc and FdYy= IfFdYy   
    and FdChr= IfFdChr and FdNo= IfFdNo and FdSr= IfFdSr   
   where @wTcTyp= 'IN' and IfPrtKey=@wCurrPrtn and IfSr= 0 and IfPrtKey=@wCurrPrtn and  
    Round(IsNull((select OdOrdQty- OdExpQty from OrdDsg where OdPrtKey=FdPrtKey and  
     OdCoCd= FdCoCd and OdTc= FdExpOdTc and OdYy= FdExpOdYy   
     and OdChr= FdExpOdChr and OdNo= FdExpOdNo and   
     OdSr= FdExpOdSr), 0) -   
    (Select sum(IfQty) from InvFgd Ifg join Fgd Fd2 on Fd2.FdPrtKey=IfPrtKey and Fd2.FdIdNo=IfFdIdNo and  
     Fd2.FdCoCd= IfCoCd and Fd2.FdTc= IfFdTc and Fd2.FdYy= IfFdYy   
     and Fd2.FdChr= IfFdChr and Fd2.FdNo= IfFdNo and Fd2.FdSr= IfFdSr   
     where Fd2.FdCoCd= Fd1.FdCoCd and Fd2.FdExpOdTc= Fd1.FdExpOdTc and   
     Fd2.FdExpOdYy= Fd1.FdExpOdYy and Fd2.FdExpOdChr= Fd1.FdExpOdChr and   
     Fd2.FdExpOdNo= Fd1.FdExpOdNo and Fd2.FdExpOdSr= Fd1.FdExpOdSr and   
     Ifg.IfSr= 0 and   
     (Select PValue From Param where PTyp= 'TC' and PMCd= Ifg.IfTc)= 'IN'), 1)< 0   
  if @row> 0   
  begin  
   select @errmsg= 'Cannot Enter This Bag As Export Qty Going Over Order Qty'  
   goto ErrHndlr  
  end  
  
  /****** For TcTyp= 'CTB', Total Bag Qty in the InvFgd table should not exceed the OdPrdQty of the corresponding Production Order Design ******/  
  /* Current Partition */  
  select @row= count('x') from inserted ins join Fgd Fd1 on FdPrtKey=IfPrtKey and FdIdNo=IfFdIdNo and  
    FdCoCd= IfCoCd and FdTc= IfFdTc and FdYy= IfFdYy   
    and FdChr= IfFdChr and FdNo= IfFdNo and FdSr= IfFdSr   
   where @wTcTyp= 'CTB' and IfPrtKey=@wCurrPrtn and  
    Round(IsNull((select OdPrdQty from OrdDsg where OdPrtKey=FdPrtKey and   
     OdCoCd= FdCoCd and OdTc= FdPrdOdTc   
     and OdYy= FdPrdOdYy and OdChr= FdPrdOdChr   
     and OdNo= FdPrdOdNo and OdSr= FdPrdOdSr), 0) -   
    (Select sum(IfQty) from InvFgd Ifg join Fgd Fd2 on Fd2.FdPrtKey=IfPrtKey and Fd2.FdIdNo=IfFdIdNo and  
     Fd2.FdCoCd= IfCoCd and Fd2.FdTc= IfFdTc and Fd2.FdYy= IfFdYy   
     and Fd2.FdChr= IfFdChr and Fd2.FdNo= IfFdNo and Fd2.FdSr= IfFdSr   
     where Fd2.FdCoCd= Fd1.FdCoCd and Fd2.FdPrdOdTc= Fd1.FdPrdOdTc and   
     Fd2.FdPrdOdYy= Fd1.FdPrdOdYy and Fd2.FdPrdOdChr= Fd1.FdPrdOdChr and   
     Fd2.FdPrdOdNo= Fd1.FdPrdOdNo and Fd2.FdPrdOdSr= Fd1.FdPrdOdSr and   
     (Select PValue From Param where PTyp= 'TC' and PMCd= Ifg.IfTc)= 'CTB'), 1)< 0   
  if @row> 0   
  begin  
   select @errmsg= 'Cannot Enter This Bag As CTB Qty Going Over Prd Qty'  
   goto ErrHndlr  
  end  
  
  /********************Geeta******Emr206***********09/04/04*********/  
  /**************Insertion not allowed If Design is on hold *************/  
  /* Current Partition */  
  select @row= count('x') from inserted join Fgd  on FdPRtKey=IfPrtKey and FdIdNo=IfFdIdNo and  
    FdCoCd= IfCoCd and FdTc= IfFdTc and FdYy= IfFdYy   
    and FdChr= IfFdChr and FdNo= IfFdNo and FdSr= IfFdSr  
   where @wTcTyp= 'IN' and IfPrtKey=@wCurrPrtn and  
   not exists (select * From OrdDsg where OdPRtKey=FdPrtKey and OdCoCd= FdCoCd and OdTc= FdExpOdTc and   
     OdYy=FdExpOdYy and OdChr= FdExpOdChr and   
     OdNo=FdExpOdNo and OdSr=FdExpOdSr and OdHld= 'N')  
  if @row> 0   
  begin  
   select @errmsg= 'Order Design is on Hold'  
   goto ErrHndlr  
  end  
  /**************************Geeta*************************/  
  
 End  
  
 if @mode in ('A','C')  
 Begin  
  /****** Voucher (Tc/Yy/Chr/No) should exist in InvHd ******/  
  /* Current Partition */  
  select @row= count('x') from inserted where IfPRtKey=@wCurrPrtn   
   and not exists (select 'x' from InvHd where InPrtKey=IfPrtKey and InIdNo=IfInIdNo  
    and InCoCd= IfCoCd and InTc= IfTc and InYy= IfYy and InChr= IfChr and InNo= IfNo)  
  if @row> 0   
  begin  
   select @errmsg= 'Voucher Not Defined In Invoice Header'  
   goto ErrHndlr  
  end  
  
  /* ### Not There In Specs ### */  
  /****** If IfSr<>0 then Tc + Yy + Chr + No + Sr should exist in InvDsg ******/  
  /* Current Partition */   
  select @row=count('x') from inserted where IfPrtKey=@wCurrPrtn and IfSr<> 0   
    and not exists (select 'x' from InvDsg where   
         IdPrtKey=IfPrtKey and IdCoCd= IfCoCd and IdTc= IfTc and IdYy= IfYy   
         and IdChr= IfChr and IdNo= IfNo and IdSr= IfSr)  
  if @row>0   
  begin  
   select @errmsg= 'Voucher Sr Not Defined In Invoice Design'  
   goto ErrHndlr  
  end  
  
  /****** IfSr >= 0 ******/  
  /* Current Partition */  
  select @row=count('x') from inserted where IfPrtKey=@wCurrPrtn and IfSr< 0   
  if @row>0   
  begin  
   select @errmsg='Invoice Bag Serial Cannot Be < Zero'  
   goto ErrHndlr  
  end  
  
  /****** IfSrNo > 0 ******/    
  /* Current Partition */  
  select @row=count('x') from inserted where IfPrtKey=@wCurrPrtn and IfSrNo<= 0   
  if @row>0   
  begin  
   select @errmsg='Invoice Bag Serial Number Cannot Be <= Zero'  
   goto ErrHndlr  
  end  
  
  /****** Quantity should be > 0 ******/  
  /* Current Partition */  
  select @row=count('x') from inserted where IfPrtKey=@wCurrPrtn and IfQty<= 0   
  if @row>0   
  begin    select @errmsg= 'Inv Bag Qty Should Be > Zero'  
   goto ErrHndlr  
  end  
  
  /****** The Fg Voucher should be of Type 'FB' or 'FR' ******/  
  /* Current Partition */  
  select @row= count('x') from inserted where IfPrtKey=@wCurrPrtn and  
   not exists (select 'x' from Param where PTyp='TC' and PMCd= IfFdTc and   
        PValue in ('FB', 'FR')    )  
  if @row>0   
  begin  
   select @errmsg= 'Fg Voucher Tc Not Defined'  
   goto ErrHndlr  
  end  
  
  /****** If @wTcTyp= 'CTB' then Bag Number Should Be Entered ******/  
  /* Current Partition */   
  select @row= count('x') from inserted where IfPrtKey=@wCurrPrtn and @wTcTyp= 'CTB' and IfFdBNo<= 0   
  if @row> 0   
  begin  
   select @errmsg= 'Bag Number Should Be Entered For CTB'  
   goto ErrHndlr  
  end  
  
  /****** 'FG' voucher/ Bag Number should be valid ******/  
  /* Current Partition */  
  select @row= count('x') from inserted where IfPrtKey=@wCurrPrtn and not exists (select 'x' from Fgd where   
    FdPrtKey=IfPrtKey and FdIdNo=IfFdIdNo and FdCoCd= IfCoCd and FdTc= IfFdTc and FdYy= IfFdYy and   
    FdChr= IfFdChr and FdNo= IfFdNo and FdSr= IfFdSr and   
    FdBYy= IfFdBYy and FdBChr= IfFdBChr and FdBNo= IfFdBNo)   
  if @row>0   
  begin  
   select @errmsg= 'Finished Goods Voucher/ Bag Number Not Defined'  
   goto ErrHndlr  
  end  
  
  /*** (Jen 2.13) ***/  
  /****** Currency of the Order and Invoice should match ******/  
  /* Current Partition */  
  select @row= count('x') from inserted   
    join Fgd on FdPrtKey=IfPrtKey and FdIdNo=IfFdIdNo and FdCoCd= IfCoCd and FdTc= IfFdTc and FdYy= IfFdYy and FdChr= IfFdChr   
     and FdNo= IfFdNo and FdSr= IfFdSr   
    join OrdMst on OmPrtKey=FdPrtKey and OmCoCd= FdCoCd and OmTc= FdExpOdTc and OmYy= FdExpOdYy   
     and OmChr= FdExpOdChr and OmNo= FdExpOdNo   
    where IfPrtKey=@wCurrPrtn and OmCmCurCd <>   
     (select CmCurCd from InvHd   
     join CustMst on CmCtg= (case when @wTcTyp= 'IN' then 'C'   
      when @wTcTyp= 'CTB' then 'T' else '' end) and CmCd= InCmCd   
     where InPrtKey=IfPrtKey and InIdNo=IfInIdNo and InCoCd= IfCoCd and InTc= IfTc and InYy= IfYy and InChr= IfChr   
      and InNo= IfNo   
     )  
  if @row> 0   
  begin  
   select @errmsg= 'Currency Of The Export Order And Invoice Not Matching'  
   goto ErrHndlr  
  end  
  /*** (Jen 2.13) ***/  
  
  /****** If TcType= 'CTB' then check if latest 'FG' ******/  
  /* Current Partition */  
  select @row= count('x') from inserted where IfPrtKey=@wCurrPrtn and @wTcTyp= 'CTB' and   
    not exists (select 'x' from Fgd Fb where FdPrtKey=IfPrtKey and FdIdNo=IfFdIdNo and  
    FdCoCd= IfCoCd and FdTc= IfFdTc and FdYy= IfFdYy and   
    FdChr= IfFdChr and FdNo= IfFdNo and FdSr= IfFdSr and FdBIdNo=IfBIdNo and  
    FdBYy= IfFdBYy and FdBChr= IfFdBChr and FdBNo= IfFdBNo   
    and not exists(Select 'x' from Fgd Fwb where   
      (Select PValue from Param where PTyp= 'TC'   
      and PMCd= Fwb.FdTc)= 'FWB'   
      and Fwb.FdPRtKey=Fb.FdPrtKey  
      and Fwb.FdCoCd= Fb.FdCoCd   
      and Fwb.FdRefYy= Fb.FdYy   
      and Fwb.FdRefKey= Fb.FdKey) )   
  if @row> 0   
  begin  
   select @errmsg= 'Not The Latest Finished Goods Voucher'  
   goto ErrHndlr  
  end  
  
  /****** IfQty & IfGrsWt should be equal to the actual Qty & Weight of the bag ******/  
  /* Current Partition */  
  select @row=count('x') from inserted where IfPrtKey=@wCurrPrtn and not exists (select 'x' from Fgd where   
    FdPRtKey=IfPrtKey and FdIdNo=IfFdIdNo and FdCoCd= IfCoCd and FdTc= IfFdTc and FdYy= IfFdYy and   
    FdChr= IfFdChr and FdNo= IfFdNo and FdSr= IfFdSr and FdBIdNo=IfBIdNo and FdBYy= IfFdBYy   
    and FdBChr= IfFdBChr and FdBNo= IfFdBNo and FdQty= IfQty and FdGrWt= IfGrsWt)   
  if @row>0   
  begin  
   select @errmsg= 'Qty/ Weight Should Be The Actual Qty/ Weight In The Bag'  
   goto ErrHndlr  
  end  
  
  /*** (Jen 3.01) ***/  
  /****** 'FG' voucher Export Order should match with that of InvDsg ******/  
  /* Current Partition */  
  select @row= count('x') from inserted   
   join InvDsg on IdPrtKey=IfPrtKey and IdCoCd= IfCoCd and IdTc= IfTc and IdYy= IfYy and IdChr= IfChr   
    and IdNo= IfNo and IdSr= IfSr   
   where IfPrtKey=@wCurrPrtn and not exists (select 'x' from Fgd where   
    FdPrtKey=IfPrtKey and FdIdNo=IfFdIdNo and FdCoCd= IfCoCd and FdTc= IfFdTc and FdYy= IfFdYy and   
    FdChr= IfFdChr and FdNo= IfFdNo and FdSr= IfFdSr and   
    FdBYy= IfFdBYy and FdBChr= IfFdBChr and FdBNo= IfFdBNo and   
    FdExpOdTc= IdExpOdTc and FdExpOdYy= IdExpOdYy and   
    FdExpOdChr= IdExpOdChr and FdExpOdNo= IdExpOdNo and   
    FdExpOdSr= IdExpOdSr)   
  if @row>0   
  begin  
   select @errmsg= 'Finished Goods Export Order Does Not Match With Invoice Design Export Voucher'  
   goto ErrHndlr  
  end  
  /*** (Jen 3.01) ***/  
  
  
  /*** ###@@@ ***/  
  /*** This check is not required in this trigger as it already exists in the 'FB' part of Fgd trigger ***/  
  /***** If FdDespQty of Base > 0 then 'CTB' rec should not exist for the Bag *****/  
  /* Select @Row= Count('x') From Inserted Ins Where @wTcTyp= 'CTB' and   
   Exists(Select 'x' from Fgd B where FdCoCd= IfCoCd and FdTc= IfFdTc   
    and FdYy= IfFdYy and FdChr= IfFdChr and FdNo= IfFdNo   
    and FdSr= IfFdSr and B.FdBM= 'B' and B.FdBNo<> 0   
    and exists(select 'x' from Fgd M where M.FdCoCd= B.FdCoCd   
     and M.FdRefYy= B.FdYy and M.FdRefKey= B.FdKey   
     and (Select PValue from Param where PTyp= 'TC' and PMCd= M.FdTc)= 'FWB')   
    )  
  If @Row> 0  
  begin  
     Select @ErrMsg= 'Contractor Bill Cannot Be Generated For This Bag As It Is In Production'  
     GoTo ErrHndlr  
  end  
  */   
  
  /*** ###@@@ ***/  
  /*** This check is not required in this trigger as it already exists in the 'FB' part of Fgd trigger ***/  
  /***** FdCTBDespQty of Base <= FdQty of Base *****/  
  /* Select @Row= Count('x') From Inserted Ins Where @wTcTyp= 'CTB' and   
   Exists(Select 'x' from Fgd B where FdCoCd= Ins.IfFdCoCd   
    and FdTc= Ins.IfFdTc and FdYy= Ins.IfFdYy   
    and FdChr= Ins.IfFdChr and FdNo= Ins.IfFdNo   
    and FdSr= Ins.IfFdSr and B.FdBM= 'B' and B.FdBNo<> 0   
    and B.FdQty< IsNull((Select sum(M.IfQty)   
      from InvFgd M where M.IfFdCoCd= FdCoCd   
      and M.IfFdTc= FdTc and M.IfFdYy= FdYy   
      and M.IfFdChr= FdChr and M.IfFdNo= FdNo   
      and M.IfFdSr= FdSr), 0)   
    )  
  If @Row> 0  
  begin  
     Select @ErrMsg= 'CTB Desp Qty Should Be<= Quantity'  
     GoTo ErrHndlr  
  end  
  */   
 End  
  
  
 if @mode = 'D'   Begin  
  /****** Cannot delete Inv Bag record if Inv Rm or Inv Lab record exists ******/  
  /* Current Partition */  
  select @row=count('x') from deleted where IfPrtKey=@wCurrPrtn and  
   (exists(Select 'x' from InvRm where IrPrtKey=IfPrtKey and IrCoCd= IfCoCd and IrTc= IfTc and IrYy= IfYy and   
     IrChr= IfChr and IrNo= IfNo and IrSr= IfSr) Or  
   exists(Select 'x' from InvLab where IlPrtKey=IfPrtKey and iLCoCd= IfCoCd and iLTc= IfTc and iLYy= IfYy and   
     iLChr= IfChr and iLNo= IfNo and iLSr= IfSr))   
  if @row>0   
  begin  
   select @errmsg='Cannot Delete As Invoice Raw Material/ Labour Record Exists'  
   goto ErrHndlr  
  end  
  
/* Zubin 212 (Shifted To Not For Repl Trigger [Manoj] ) */   
  /******* (BCls= 'N') Bag is again opened for transaction ******/  
/*  update Bag set  
  BCls= 'N'   
  from deleted where @wTcTyp= 'IN' and BCoCd= IfCoCd and BYy= IfFdBYy and BChr= IfFdBChr   
   and BNo= IfFdBNo    
*/  
/* Zubin 212 (Shifted To Not For Repl Trigger [Manoj] ) */   
  
/* Zubin 212 (Shifted To Not For Repl Trigger [Manoj] ) */   
  /******* If @wTcTyp= 'IN' then update FdDespQty of Fgd ******/  
/*  update Fgd set   
  FdDespQty= FdDespQty- IsNull((Select sum(del.IfQty) from deleted del where   
   @wTcTyp= 'IN' and del.IfCoCd= FdCoCd and del.IfFdTc= FdTc   
   and del.IfFdYy= FdYy and del.IfFdChr= FdChr and del.IfFdNo= FdNo   
   and del.IfFdSr= FdSr), 0)   
  from deleted where @wTcTyp= 'IN' and FdCoCd= IfCoCd and FdTc= IfFdTc   
   and FdYy= IfFdYy and FdChr= IfFdChr and FdNo= IfFdNo and FdSr= IfFdSr  
*/  
/* Zubin 212 (Shifted To Not For Repl Trigger [Manoj] ) */   
  
/* Zubin 212 (Shifted To Not For Repl Trigger [Manoj] ) */   
  /******* If @wTcTyp= 'CTB' then update FdCTBDespQty of Fgd ******/  
/*  update Fgd set   
  FdCTBDespQty= FdCTBDespQty- IsNull((Select sum(del.IfQty) from deleted del where   
   @wTcTyp= 'CTB' and del.IfCoCd= FdCoCd and del.IfFdTc= FdTc   
   and del.IfFdYy= FdYy and del.IfFdChr= FdChr and del.IfFdNo= FdNo   
   and del.IfFdSr= FdSr), 0)   
  from deleted where @wTcTyp= 'CTB' and FdCoCd= IfCoCd and FdTc= IfFdTc   
   and FdYy= IfFdYy and FdChr= IfFdChr and FdNo= IfFdNo and FdSr= IfFdSr  
*/  
/* Zubin 212 (Shifted To Not For Repl Trigger [Manoj] ) */   
 End  
  
  
/* if @mode in ('A')  
 Begin  
*/  
/* Zubin 212 (Shifted To Not For Repl Trigger [Manoj] ) */   
  /******* update FdDespQty of Fgd ******/  
/*  update Fgd set   
  FdDespQty= FdDespQty+ IsNull((Select sum(ins.IfQty) from inserted ins where   
   @wTcTyp= 'IN' and ins.IfCoCd= FdCoCd and ins.IfFdTc= FdTc   
   and ins.IfFdYy= FdYy and ins.IfFdChr= FdChr and ins.IfFdNo= FdNo   
   and ins.IfFdSr= FdSr), 0)   
  from inserted where @wTcTyp= 'IN' and FdCoCd= IfCoCd and FdTc= IfFdTc   
   and FdYy= IfFdYy and FdChr= IfFdChr and FdNo= IfFdNo and FdSr= IfFdSr  
*/  
  /******* update FdDespQty of Fgd ******/  
/* Zubin 212 (Shifted To Not For Repl Trigger [Manoj] ) */   
/*  update Fgd set   
  FdCTBDespQty= FdCTBDespQty+ IsNull((Select sum(ins.IfQty) from inserted ins where   
   @wTcTyp= 'CTB' and ins.IfCoCd= FdCoCd and ins.IfFdTc= FdTc   
   and ins.IfFdYy= FdYy and ins.IfFdChr= FdChr and ins.IfFdNo= FdNo   
   and ins.IfFdSr= FdSr), 0)   
  from inserted where @wTcTyp= 'CTB' and FdCoCd= IfCoCd and FdTc= IfFdTc   
   and FdYy= IfFdYy and FdChr= IfFdChr and FdNo= IfFdNo and FdSr= IfFdSr  
*/  
  
/* Zubin 212 (Shifted To Not For Repl Trigger [Manoj] ) */   
  /******* (BCls= 'Y') Bag is closed for transaction ******/  
/*  update Bag set   
  BCls= 'Y'  
  from inserted where @wTcTyp= 'IN' and BCoCd= IfCoCd and BYy= IfFdBYy   
   and BChr= IfFdBChr and BNo= IfFdBNo   
 End  
*/  
  
 Return  
ErrHndlr:  
 Begin   select @errmsg='|'+@errmsg  
  raiserror(@errmsg,16,1)   rollback transaction   
 end  
end  


GOGO

/****** Object:  Trigger [dbo].[InvHd_INUPDL]    Script Date: 5/12/2015 11:26:16 PM ******/
    
ALTER TRIGGER [dbo].[InvHd_INUPDL] ON [dbo].[InvHd]    
    FOR INSERT,UPDATE,DELETE    NOT FOR REPLICATION
AS    
BEGIN    
 /* Trigger Version 3.9.0 - Speed Optimization - If Update */    
 DECLARE @ROW INT    
 DECLARE @mode char(1)    
 DECLARE @errmsg varchar(255)    
 Declare @wHCd Varchar(8)     
 DECLARE @wTcTyp varchar(3)    
 DECLARE @wEntTyp varchar(8)    
 DECLARE @ReplAgt BIT      
 DECLARE @wCurrPrtn VarChar(1)     
     
 IF @@ROWCOUNT=0   RETURN    
 If Exists(Select 'x' From TrgDisa Where TdTblNm = 'InvHd')    
  RETURN    
    
 select @mode=case (select count('x') from inserted) when 0 then 'D'    
       else    
   case (select count('x') from deleted) when 0 then 'A' else 'C' end    
       end     
      
 Select @wHCd= 'ZSELF'    
 Select @wCurrPrtn='C'    
    
 /*** Storing Tc Type In The Variable @wTcTyp ***/    
 if @mode in ('A','C')    
 begin    
  select @wTcTyp= max(PValue) from inserted join Param on PTyp= 'TC' and PMCd= InTc and PSCd= ''    
 end    
    
 if @mode in ('D')    
 begin    
  select @wTcTyp= max(PValue) from deleted join Param on PTyp= 'TC' and PMCd= InTc and PSCd= ''    
 end    
    
 /* **** Checking Whether a Normal Transaction or Replicated One **** */    
 Set @ReplAgt=Convert(BIT,SessionProperty('Replication_Agent'))    

 If @ReplAgt = 1         Return
     
 /******* Sachin 3.02.0 [30/11/07] - Purge Partition (Speed Optimization) */    
 if @mode='C'    
 begin    
  Select @row=count('x') From Inserted Where InPrtKey<>@wCurrPrtn    
  if @row> 0     
  begin    
   select @errmsg= 'Cannot Edit Records of Previous Partition'    
   goto ErrHndlr    
  end    
 End    
    
 if @mode='A'    
 Begin    
  Select @row=count('x') From Inserted Where InPrtKey <> @wCurrPrtn     
  if @row> 0     
  begin    
   select @errmsg= 'Cannot Insert Partition Key Other Than ''C'''    
   goto ErrHndlr    
  end    
 end     
    
 If @Mode='D'     
 Begin     
  select @row= count('x') from deleted del where del.InPrtKey<>@wCurrPrtn    
  if @row> 0     
  begin    
   select @errmsg= 'Cannot Delete Records From Previous Partition'    
   goto ErrHndlr    
  end    
 End    
 /******* Sachin 3.02.0 [30/11/07] - Purge Partition (Speed Optimization) */    
    
 If  (@ReplAgt<>1)    
 Begin    
  /******Cannot Insert/Update/Delete  if Crm database /Emr database with TrdCocd not in sdCoCd ****/    
  /***if SysId=Emr Record not found then normal behaviour i.e can insert/update/delete****/    
  if @mode in ('A','C')    
  Begin     
  /* Current Partition */    
  select @ROW =Count('x')  From inserted where InPrtKey=@wCurrPrtn and     
  Exists ( select 'x' From SysDb where (SysId='EMR' and (SysDb='LTP' Or ','+sdCocd+',' Not  like '%,'+ InCoCd +',%'))     
   or (SysId='CRM' and SysDb='WEB'))    
  if @Row > 0      
   Begin     
    select @errmsg= 'Cannot Insert /Update '     
    goto ErrHndlr     
   end     
  End     
    
  if @mode in ('D')    
  Begin     
  /* Current Partition */    
  select @ROW =Count('x')  From Deleted where InPrtKey=@wCurrPrtn and     
  Exists ( select 'x' From SysDb where (SysId='EMR' and (SysDb='LTP' Or ','+sdCocd+',' Not  like '%,'+ InCoCd +',%'))    
   or (SysId='CRM' and SysDb='WEB'))    
  if @Row > 0      
   Begin     
    select @errmsg= 'Cannot Delete '     
    goto ErrHndlr     
   end     
  End     
     
  if @mode in ('A','C')    
  Begin     
   /*** (Bef 2.13)     
   select @row= count('x') from inserted where @wTcTyp= 'IN'     
    and Str(InFOBRs, 16, 0)<> Str(Round(InFOB* inCnvRt, 2), 16, 0)     
   (Bef 2.13) ***/    
   /*** (Jen 2.13) ***/    
   /* (Bef 22/08/06)    
   select @row= count('x') from inserted where @wTcTyp= 'IN'     
    and convert (decimal (16, 0), InFOBRs)<>     
        convert( decimal (16, 0), InFOB* (case when InMulDiv= 'D' then InCnvFct else 1 end)    
           / (case when InMulDiv= 'M' then InCnvFct else 1 end)     
           * (case when InRsMulDiv= 'M' then inCnvRt else 1 end)    
           / (case when InRsMulDiv= 'D' then inCnvRt else 1 end))   
   (Bef 22/08/06) */    
      
   /*****4.1.2********/     
   Select @Row=Count('x') From Inserted ins where     
    (Select hHolChkYN From Head where HCd='ZSELF' and HCoCd=InCoCd)='Y' and    
    (@Mode='A' or (@Mode='C' and  ins.InDt<>IsNull((Select InDt From Deleted del where     
          ins.InPrtKey=del.InPrtKey and ins.InIdNo=del.InIdNo     
          and ins.InCoCd= del.InCoCd and ins.InTc= del.InTc and ins.InYy= del.InYy     
           and ins.InChr= del.InChr and ins.InNo= del.InNo),ins.InDt))) and      
         exists(Select 'x' From DtTbl Where DtTag='HOL' and InDt=DtDt and DtHolYN='Y')    
   If @Row> 0    
   begin      
    select @errmsg= 'Cannot Add Voucher on Holidays.'      
    goto ErrHndlr      
   end      
       
     
   /* Current Partition */    
   /*** Jay 3.3(FG) ***/    
   /***** Manali trading Module *****/    
   If @mode='A' or (@mode='C' and (Update(InFOBRs) or Update(InFOB) or Update(InRsMulDiv) or Update(inCnvRt)))    
   Begin     
     select @row= count('x') from inserted where @wTcTyp in ('IN','FG', 'JT') and InPrtKey=@wCurrPrtn     
      and convert (decimal (16, 0), InFOBRs)<>     
          convert( decimal (16, 0), InFOB* (case when InRsMulDiv= 'M' then inCnvRt else 1 end)    
             / (case when InRsMulDiv= 'D' then inCnvRt else 1 end))     
     /*** (Jen 2.13) ***/    
     if @row> 0     
     begin    
      select @errmsg= 'FOB In Rs Is Not Proper'    
      goto ErrHndlr    
     end    
   End      
  End    
 End    
    
 /******** Only One TcTyp records can be Inserted at a time.********/    
 /* Current Partition */    
 select @Row= Count(distinct(PValue)) from inserted Join param On    
  PTyp= 'TC' and PMCd= InTc and PScd= '' Where InPrtKey=@wCurrPrtn     
 If @Row> 1    
 begin    
  Select @ErrMsg= 'Transactions of only One TcType can be Acted On'    
    GoTo ErrHndlr    
 end    
    
 /******** Only One TcTyp records can be Deleted at a time. *********/    
 /* Current Partition */    
 select @Row= Count(distinct(PValue)) from deleted Join param on    
  PTyp= 'TC' and PMcd= InTc and PSCd='' Where InPrtKey=@wCurrPrtn    
 If @Row> 1    
 begin    
  Select @ErrMsg= 'Transactions of only One TcType can be Acted on'    
  GoTo ErrHndlr    
 end    
    
 /*** Storing Entry Type In The Variable @wEntTyp ***/    
 select @wEntTyp= PValue1 from Param  where PTyp= 'TCTYP' and PMCd= @wTcTyp and PSCd= ''    
     
  if @mode = 'C'    
  Begin    
  /******** Key Cannot Be Changed ********/     
  /* Current Partition */    
  If Update(InPrtKey) or Update(InIdNo) or Update(InCoCd) or Update(InTc) or Update(InYy) or Update(InChr) or Update(InNo)    
  Begin     
    select @Row=Count('x') from deleted del where del.InPrtKey=@wCurrPrtn and    
      not exists (Select 'x' from inserted ins where ins.InPrtKey=del.InPrtKey     
       and ins.InIdNo=del.InIdNo and ins.InCoCd= del.InCoCd     
       and ins.InTc= del.InTc and ins.InYy= del.InYy and ins.InChr= del.InChr     
       and ins.InNo= del.InNo)    
    If @Row > 0    
    begin     
     Select @ErrMsg='Cannot Change Company Code/ Tc/Yy/Chr/No'    
     GoTo ErrHndlr    
    end    
    End    
        
  /****** No Change when InCls= 'Y' ******/    
  /* Current Partition */    
  If Update(InCls) or Update(InExpNo) or Update(InCmCd) or Update(InDt) or    
     Update(InLmgCst) or Update(InLmgDt) or Update(InLmgSal) Or    
     Update(InLmpCst) or Update(InLmpDt) or Update(InLmpSal) Or    
     Update(InLmsCst) or Update(InLmsDt) or Update(InLmsSal) Or    
     Update(InLmlCst) or Update(InLmlDt) or Update(InLmlSal) Or    
     Update(InVaDc) or Update(InVaX) or Update(InCnvRt) or Update(InFOB) or Update(InFOBRs) Or    
     Update(InComm) or Update(InCommRs) Or Update(InFrt) Or Update(InFrtRs) Or Update(InIns) Or Update(InInsRs) Or    
     Update(InDsc) Or Update(InDscRs) Or Update(InDeduct) Or Update(InDeductRs) Or     
     Update(InHHideCommYn) or Update(InHSelfDscYN) or Update(InHSelfDeductYN) or    
     Update(InTot) Or Update(InTotRs) Or Update(InTotAs) Or Update(InDiaCost) Or    
     Update(InCSCost) Or Update(InAdv) Or Update(InPrnDt) Or Update(InhInvTyp)    
  Begin     
    select @Row=Count('x') from deleted del where del.InPrtKey=@wCurrPrtn and del.InCls= 'Y' and     
    /*** (Jen 2.13) 28/11/06 ***/    
    /* *** Jay 3.1.1 [Discount] *** */    
    /* **** Manali 3.03 - 21/06/08 - Palladium - InLmlCst, InLmlDt, InLmlSal added */    
    /* **** Manali 3.8.1 - Self Payable Options Added in Invoice - InDeduct, InDeductRs, InHHideCommYn, InHSelfDscYN, InHSelfDeductYN */     
      not exists (Select 'x' from inserted ins where ins.InPrtKey=del.InPrtKey     
       and ins.InIdNo=del.InIdNo and ins.InCoCd= del.InCoCd     
       and ins.InTc= del.InTc and ins.InYy= del.InYy and ins.InChr= del.InChr     
       and ins.InNo= del.InNo and ins.InExpNo= del.InExpNo and ins.InCmCd= del.InCmCd     
       and ins.InDt= del.InDt and del.InLmgCst= ins.InLmgCst and del.InLmgDt= ins.InLmgDt     
       and del.InLmgSal= ins.InLmgSal and del.InLmpCst= ins.InLmpCst and del.InLmpDt= ins.InLmpDt     
       and del.InLmsSal= ins.InLmsSal and del.InLmsCst= ins.InLmsCst and del.InLmsDt= ins.InLmsDt    
       and del.InLmlSal= ins.InLmlSal and del.InLmlCst= ins.InLmlCst and del.InLmlDt= ins.InLmlDt    
       and del.InLmpSal= ins.InLmpSal and del.InVaDC= ins.InVaDC and del.InVaX= ins.InVaX     
       and convert (decimal (16, 4), del.InCnvRt)= convert (decimal (16, 4), ins.InCnvRt)     
       and convert (decimal (16, 2), del.InFOB)= convert (decimal (16, 2), ins.InFOB)     
       and convert (decimal (16, 2), del.InFOBRs)= convert (decimal (16, 2), ins.InFOBRs)     
       and convert (decimal (16, 2), del.InComm)= convert (decimal (16, 2), ins.InComm)     
       and convert (decimal (16, 2), del.InCommRs)= convert (decimal (16, 2), ins.InCommRs)     
       and convert (decimal (16, 2), del.InFrt)= convert (decimal (16, 2), ins.InFrt)     
       and convert (decimal (16, 2), del.InFrtRs)= convert (decimal (16, 2), ins.InFrtRs)     
       and convert (decimal (16, 2), del.InIns)= convert (decimal (16, 2), ins.InIns)     
       and convert (decimal (16, 2), del.InInsRs)= convert (decimal (16, 2), ins.InInsRs)     
       and convert (decimal (16, 2), del.InDsc)= convert (decimal (16, 2), ins.InDsc)     
       and convert (decimal (16, 2), del.InDscRs)= convert (decimal (16, 2), ins.InDscRs)     
       and convert (decimal (16, 2), del.InDeduct)= convert (decimal (16, 2), ins.InDeduct)     
       and convert (decimal (16, 2), del.InDeductRs)= convert (decimal (16, 2), ins.InDeductRs)     
       and del.InHHideCommYn=ins.InHHideCommYn and del.InHSelfDscYN=ins.InHSelfDscYN and del.InHSelfDeductYN=ins.InHSelfDeductYN    
       and convert (decimal (16, 2), del.InTot)= convert (decimal (16, 2), ins.InTot)     
       and convert (decimal (16, 2), del.InTotRs)= convert (decimal (16, 2), ins.InTotRs)     
       and del.InTotAs= ins.InTotAs     
       and convert (decimal (16, 2), del.InDiaCost)= convert (decimal (16, 2), ins.InDiaCost)     
       and convert (decimal (16, 2), del.InCSCost)= convert (decimal (16, 2), ins.InCSCost)     
       and convert (decimal (16, 2), del.InAdv)= convert (decimal (16, 2), ins.InAdv)     
       and ins.InPrnDt= del.InPrnDt and del.InhInvTyp= ins.InhInvTyp)    
    /*** (Jen 2.13) 28/11/06 ***/    
    
    /* (Bef 2.13)    
      not exists (Select 'x' from inserted ins where ins.InCoCd= del.InCoCd     
       and ins.InTc= del.InTc and ins.InYy= del.InYy and ins.InChr= del.InChr     
       and ins.InNo= del.InNo and ins.InExpNo= del.InExpNo and ins.InCmCd= del.InCmCd     
       and ins.InDt= del.InDt and del.InLmgCst= ins.InLmgCst and del.InLmgDt= ins.InLmgDt     
       and del.InLmgSal= ins.InLmgSal and del.InLmpCst= ins.InLmpCst and del.InLmpDt= ins.InLmpDt     
       and del.InLmsSal= ins.InLmsSal and del.InLmsCst= ins.InLmsCst and del.InLmsDt= ins.InLmsDt    
       and del.InLmpSal= ins.InLmpSal and del.InVaDC= ins.InVaDC and del.InVaX= ins.InVaX     
       and convert (decimal (16, 4), del.InCnvRt)= convert (decimal (16, 4), ins.InCnvRt)     
       and del.InFOB= ins.InFOB and del.InFOBRs= ins.InFOBRs     
       and del.InComm= ins.InComm and del.InCommRs= ins.InCommRs and del.InFrt= ins.InFrt     
       and del.InFrtRs= ins.InFrtRs and del.InIns= ins.InIns and del.InInsRs= ins.InInsRs and del.InTot= ins.InTot     
       and del.InTotRs= ins.InTotRs and del.InTotAs= ins.InTotAs and del.InDiaCost= ins.InDiaCost     
       and del.InCSCost= ins.InCSCost and del.InAdv= ins.InAdv and ins.InPrnDt= del.InPrnDt    
       and del.InhInvTyp= ins.InhInvTyp)    
    (Bef 2.13) */    
    
    If @Row > 0    
    begin     
     Select @ErrMsg= 'Cannot Change InvExpNo,Cust/Contractor,Dt,LMG, LMP, LMS and LML Cst/Dt/Sal,CnvRt,FOB,Comm,Frt,Ins,Dsc,Deduct,Self Payable,TotVal,Dia/CS Cst,Adv Remit,Prn Dt, Inv Typ As Inv Is Clsd'    
     GoTo ErrHndlr    
    end    
    End    
        
  /*** Jay 3.3(FG)***/    
  /****** (Jen 3.1.0_1) included InRemakingCtgs, InInvGrsWtEqActual in the list ******/     
  /****** (Jen 2.13) included InFobPer in the list ******/     
  /******** Cannot Change InDt, InLmgCst, InLmgDt, InLmgSal, InLmpCst, InLmpDt, InLmpSal, InVaDc, InVaX if Record exists in InvDsg ********/     
                /*****Cannot Change InHGldFiness,InHPlFiness ***/     
  /* Current Partition */    
  /* **** Manali 3.03 - 21/06/08 - Palladium - InLmlCst, InLmlDt, InLmlSal, InHPdFiness added*/    
  /* ***** Manali 3.6.0 - 29/10/09 - Lme Fields Remove from Following Check (Refresh Metal Rate in Invoice)   ***** */    
  /***** Manali trading Module *****/    
  If Update(InDt) or Update(InVaDc) or Update(InVaX) or Update(InHRndOffFOB) or Update(InFrRmLoc) or Update(InToRmLoc)     
     or Update(InFrRmDc) or Update(InToRmDc) or Update(InHGldFiness) or Update(InHPlFiness)    
     or Update(InHSlFiness) or Update(InHPdFiness) or Update(InFobPer) or Update(InRemakingCtgs)    
      or Update(InInvGrsWtEqActual)    
  Begin     
    select @Row=Count('x') from deleted del where del.InPrtKey=@wCurrPrtn and    
     ((@wTcTyp Not In ('FG', 'JT') and exists(Select 'x' from InvDsg where IdPRtKey=del.InPrtKey and IdInIdNo=del.InIdNo and IdCoCd= del.InCoCd and IdTc= del.InTc     
       and IdYy= del.InYy and IdChr= del.InChr and IdNo= del.InNo)) Or    
      (@wTcTyp In ('FG', 'JT') and exists(Select 'x' from OrdMst where OmPRtKey=del.InPrtKey and OmCoCd= del.InCoCd and OmTc= del.InTc     
       and OmYy= del.InYy and OmChr= del.InChr and OmNo= del.InNo))    
     ) and    
     /*exists(Select 'x' from InvDsg where IdPRtKey=del.InPrtKey and IdInIdNo=del.InIdNo and IdCoCd= del.InCoCd and IdTc= del.InTc     
       and IdYy= del.InYy and IdChr= del.InChr and IdNo= del.InNo) and */    
     /*** (Jen 2.13) 28/11/06 ***/    
      not exists (Select 'x' from inserted ins where ins.InPrtKey=del.InPrtKey     
       and ins.InIdNo=del.InIdNo and ins.InCoCd= del.InCoCd     
       and ins.InTc= del.InTc and ins.InYy= del.InYy and ins.InChr= del.InChr     
       and ins.InNo= del.InNo     
      /*  and ins.InDt= del.InDt     
       and ins.InLmgDt= del.InLmgDt and ins.InLmsDt= del.InLmsDt and ins.InLmpDt= del.InLmpDt and ins.InLmlDt= del.InLmlDt      
       and convert (decimal (16, 2), ins.InLmgCst)= convert (decimal (16, 2), del.InLmgCst)     
       and convert (decimal (16, 2), ins.InLmgSal)= convert (decimal (16, 2), del.InLmgSal)     
       and convert (decimal (16, 2), ins.InLmpCst)= convert (decimal (16, 2), del.InLmpCst)     
       and convert (decimal (16, 2), ins.InLmsSal)= convert (decimal (16, 2), del.InLmsSal)     
       and convert (decimal (16, 2), Ins.InLmsCst)= convert (decimal (16, 2), del.InLmsCst)     
       and convert (decimal (16, 2), ins.InLmpSal)= convert (decimal (16, 2), del.InLmpSal)     
       and convert (decimal (16, 2), ins.InLmlSal)= convert (decimal (16, 2), del.InLmlSal)     
       and convert (decimal (16, 2), Ins.InLmlCst)= convert (decimal (16, 2), del.InLmlCst)  */    
       and convert (decimal (16, 2), ins.InVaDc)= convert (decimal (16, 2), del.InVaDc)     
       and convert (decimal (16, 2), ins.InVaX)= convert (decimal (16, 2), del.InVaX)     
       and ins.InHRndOffFOB= del.InHRndOffFOB     
       and ins.InFrRmLoc= del.InFrRmLoc and ins.InFrRmDc= del.InFrRmDc     
       and ins.InToRmLoc= del.InToRmLoc and ins.InToRmDc= del.InToRmDc    
       and convert (decimal (16, 3), ins.InHGldFiness)= convert (decimal (16, 3), del.InHGldFiness)     
       and convert (decimal (16, 3), ins.InHPlFiness)= convert (decimal (16, 3), del.InHPlFiness)     
       and convert (decimal (16, 3), ins.InHSlFiness)= convert (decimal (16, 3), del.InHSlFiness)     
       and convert (decimal (16, 3), ins.InHPdFiness)= convert (decimal (16, 3), del.InHPdFiness)     
       and ins.InFobPer= del.InFobPer and ins.InRemakingCtgs= del.InRemakingCtgs     
       and ins.InInvGrsWtEqActual= del.InInvGrsWtEqActual     
       )     
     /*** (Jen 2.13) 28/11/06 ***/    
    
     /* (Bef 2.13)    
      not exists (Select 'x' from inserted ins where ins.InCoCd= del.InCoCd     
       and ins.InTc= del.InTc and ins.InYy= del.InYy and ins.InChr= del.InChr     
       and ins.InNo= del.InNo and ins.InDt= del.InDt and ins.InLmgCst= del.InLmgCst     
       and ins.InLmgDt= del.InLmgDt and ins.InLmgSal= del.InLmgSal and ins.InLmpCst= del.InLmpCst     
       and ins.InLmsSal= del.InLmsSal and Ins.InLmsCst= del.InLmsCst and ins.InLmsDt= del.InLmsDt    
       and ins.InLmpDt= del.InLmpDt and ins.InLmpSal= del.InLmpSal and ins.InVaDc= del.InVaDc     
       and ins.InVaX= del.InVaX and ins.InHRndOffFOB= del.InHRndOffFOB     
       and ins.InFrRmLoc= del.InFrRmLoc and ins.InFrRmDc= del.InFrRmDc     
       and ins.InToRmLoc= del.InToRmLoc and ins.InToRmDc= del.InToRmDc    
       and ins.InHGldFiness=del.InHGldFiness and ins.InHPlFiness=del.InHPlFiness     
       and ins.InHSlFiness=del.InHSlFiness and ins.InFobPer= del.InFobPer     
       )     
     (Bef 2.13) */    
    If @Row > 0    
    begin     
     Select @ErrMsg= 'Cannot Chng Inv Dt, Gld, Plt, Slv And Pd Fineness, Value Addn For Dia, ColStn & Acc, FOB Rnd Opt, FOB %, Rm Locs, IR As ' + case When @wTcTyp In ('FG', 'JT') Then 'Ord' else 'Inv' End + 'Dsg Rec Exists'    
     /*Select @ErrMsg= 'Cannot Chng Inv Dt, LMG, LMP, LMS And LML Cst/Dt/Sal, Gld, Plt, Slv And Pd Fineness, Value Addn For Dia, ColStn & Acc, FOB Rnd Opt, FOB %, Rm Locs, IR As Inv Dsg Rec Exists'*/    
     GoTo ErrHndlr    
    end    
    End    
        
  /* ***** Manali 3.6.0 - 29/10/09 - Lme Fields Check for TcTyp<>'IN' (Refresh Metal Rate in Invoice)   ***** */    
  /***** Manali trading Module *****/    
  If @wTcTyp In ('CTB', 'FG', 'JT')    
  Begin    
    If Update(InLmgCst) or Update(InLmgDt) or Update(InLmgSal) Or    
       Update(InLmpCst) or Update(InLmpDt) or Update(InLmpSal) Or    
       Update(InLmsCst) or Update(InLmsDt) or Update(InLmsSal) Or    
       Update(InLmlCst) or Update(InLmlDt) or Update(InLmlSal)    
    Begin     
      select @Row=Count('x') from deleted del where del.InPrtKey=@wCurrPrtn and    
       ((@wTcTyp = 'CTB' and exists(Select 'x' from InvDsg where IdPRtKey=del.InPrtKey and IdInIdNo=del.InIdNo and IdCoCd= del.InCoCd and IdTc= del.InTc     
         and IdYy= del.InYy and IdChr= del.InChr and IdNo= del.InNo)) Or    
        (@wTcTyp In ('FG', 'JT') and exists(Select 'x' from OrdMst where OmPRtKey=del.InPrtKey and OmCoCd= del.InCoCd and OmTc= del.InTc     
         and OmYy= del.InYy and OmChr= del.InChr and OmNo= del.InNo))    
       ) and    
       /*exists(Select 'x' from InvDsg where IdPRtKey=del.InPrtKey and IdInIdNo=del.InIdNo and IdCoCd= del.InCoCd and IdTc= del.InTc     
         and IdYy= del.InYy and IdChr= del.InChr and IdNo= del.InNo) and */    
        not exists (Select 'x' from inserted ins where ins.InPrtKey=del.InPrtKey     
         and ins.InIdNo=del.InIdNo and ins.InCoCd= del.InCoCd     
         and ins.InTc= del.InTc and ins.InYy= del.InYy and ins.InChr= del.InChr     
         and ins.InNo= del.InNo     
         and ins.InLmgDt= del.InLmgDt and ins.InLmsDt= del.InLmsDt and ins.InLmpDt= del.InLmpDt and ins.InLmlDt= del.InLmlDt      
         and convert (decimal (16, 2), ins.InLmgCst)= convert (decimal (16, 2), del.InLmgCst)     
         and convert (decimal (16, 2), ins.InLmgSal)= convert (decimal (16, 2), del.InLmgSal)     
         and convert (decimal (16, 2), ins.InLmpCst)= convert (decimal (16, 2), del.InLmpCst)     
         and convert (decimal (16, 2), ins.InLmsSal)= convert (decimal (16, 2), del.InLmsSal)     
         and convert (decimal (16, 2), Ins.InLmsCst)= convert (decimal (16, 2), del.InLmsCst)     
         and convert (decimal (16, 2), ins.InLmpSal)= convert (decimal (16, 2), del.InLmpSal)     
         and convert (decimal (16, 2), ins.InLmlSal)= convert (decimal (16, 2), del.InLmlSal)     
         and convert (decimal (16, 2), Ins.InLmlCst)= convert (decimal (16, 2), del.InLmlCst)       
         )     
      If @Row > 0    
      begin     
       Select @ErrMsg= 'Cannot Chng LMG, LMP, LMS And LML Cst/Dt/Sal As ' + case When @wTcTyp In ('FG', 'JT') Then 'Ord' else 'Inv' End + ' Dsg Rec Exists'    
       GoTo ErrHndlr    
      End    
      End    
    End    
            
  /* ***** Manali 3.6.0 - 29/10/09 - For TcTyp='IN' - Lme fields - Custom Matching Check (Refresh Metal Rate in Invoice)   ***** */    
   If @wTcTyp = 'IN'     
   Begin    
     If (Update(InLmgCst) or Update(InLmgDt) or Update(InLmgSal) Or    
         Update(InLmpCst) or Update(InLmpDt) or Update(InLmpSal) Or    
         Update(InLmsCst) or Update(InLmsDt) or Update(InLmsSal) Or    
         Update(InLmlCst) or Update(InLmlDt) or Update(InLmlSal) Or Update(InNotionRtYn))    
     Begin     
      select @Row=Count('x') from deleted del where del.InPrtKey=@wCurrPrtn and @wTcTyp = 'IN'     
       and exists(Select 'x' from InvZRm where IzPRtKey=del.InPrtKey and IzInIdNo=del.InIdNo and IzCoCd= del.InCoCd and IzTc= del.InTc     
          and IzYy= del.InYy and IzChr= del.InChr and IzNo= del.InNo) and     
        not exists (Select 'x' from inserted ins where ins.InPrtKey=del.InPrtKey     
         and ins.InIdNo=del.InIdNo and ins.InCoCd= del.InCoCd     
         and ins.InTc= del.InTc and ins.InYy= del.InYy and ins.InChr= del.InChr     
         and ins.InNo= del.InNo     
         and ins.InLmgDt= del.InLmgDt and ins.InLmsDt= del.InLmsDt and ins.InLmpDt= del.InLmpDt and ins.InLmlDt= del.InLmlDt      
         and convert (decimal (16, 2), ins.InLmgCst)= convert (decimal (16, 2), del.InLmgCst)     
         and convert (decimal (16, 2), ins.InLmgSal)= convert (decimal (16, 2), del.InLmgSal)     
         and convert (decimal (16, 2), ins.InLmpCst)= convert (decimal (16, 2), del.InLmpCst)     
         and convert (decimal (16, 2), ins.InLmsSal)= convert (decimal (16, 2), del.InLmsSal)     
         and convert (decimal (16, 2), Ins.InLmsCst)= convert (decimal (16, 2), del.InLmsCst)     
         and convert (decimal (16, 2), ins.InLmpSal)= convert (decimal (16, 2), del.InLmpSal)     
         and convert (decimal (16, 2), ins.InLmlSal)= convert (decimal (16, 2), del.InLmlSal)     
         and convert (decimal (16, 2), Ins.InLmlCst)= convert (decimal (16, 2), del.InLmlCst)       
         and ins.InNotionRtYn= del.InNotionRtYn    
         )     
      If @Row > 0    
      begin     
       Select @ErrMsg= 'Cannot Chng LMG, LMP, LMS And LML Cst/Dt/Sal, Notional Rate As Custom Matching Done'    
       GoTo ErrHndlr    
      End    
    End    
  End    
  /*** (Bef 2.13)    
  --******** If @wTcTyp= 'CTB' then Cannot Change InCmCd, InCnvRt if Record exists in InvDsg ********     
  select @Row= Count('x') from deleted del where @wTcTyp= 'CTB' and     
   exists(Select 'x' from InvDsg where IdTc= del.InTc and IdYy= del.InYy and IdChr= del.InChr     
     and IdNo= del.InNo) and     
    not exists (Select 'x' from inserted ins where ins.InTc= del.InTc and ins.InYy= del.InYy and ins.InChr= del.InChr     
     and ins.InNo= del.InNo and ins.InCmCd= del.InCmCd and Str(Round(ins.InCnvRt, 3), 16, 3)= Str(Round(del.InCnvRt, 3), 16, 3))     
  If @Row> 0    
  begin     
   Select @ErrMsg= 'Cannot Change Contractor Code, Currency Conversion Rate As Inv Dsg Rec Exists'    
   GoTo ErrHndlr    
  end    
  (Bef 2.13) ****/    
    
  /*** (Jen 2.13) ***/      
  /******** If @wTcTyp= 'CTB' then Cannot Change InCmCd, InRsMulDiv, InCnvRt if Record exists in InvDsg ********/     
  /* Current Partition */    
  If @wTcTyp= 'CTB'     
  Begin    
     If Update(InCmCd) or Update(InRsMulDiv) or Update(InCnvRt)    
     Begin     
      select @Row= Count('x') from deleted del where @wTcTyp= 'CTB' and del.InPrtKey=@wCurrPrtn and    
       exists(Select 'x' from InvDsg where IdInIdNo=del.InIdNo and IdTc= del.InTc and IdYy= del.InYy and IdChr= del.InChr     
         and IdNo= del.InNo and IdPrtKey=del.InPrtKey) and     
        not exists (Select 'x' from inserted ins where ins.InIdNo=del.InIdNo and ins.InTc= del.InTc and ins.InYy= del.InYy and ins.InChr= del.InChr     
         and ins.InNo= del.InNo and ins.InPrtKey=del.InPrtKey and ins.InCmCd= del.InCmCd and ins.InRsMulDiv= del.InRsMulDiv     
         and convert (decimal (16, 4), ins.InCnvRt)= convert (decimal (16, 4), del.InCnvRt))     
      If @Row> 0    
      begin     
       Select @ErrMsg= 'Cannot Change Contractor Code, Currency Conversion Rate As Inv Dsg Rec Exists'    
       GoTo ErrHndlr    
      end    
      End    
    End    
        
  /*** Jay 3.3(FG) ***/    
  /******** If @wTcTyp= 'IN' then Cannot Change InCmCd, InCnvFct, InMulDiv if Record exists in InvDsg, InvFgd ********/     
  /* Current Partition */    
  If @wTcTyp= 'IN'     
  Begin    
     If Update(InCmCd) or Update(InRsMulDiv) or Update(InCnvRt)    
     Begin     
      select @Row= Count('x') from deleted del where del.InPrtKey=@wCurrPrtn and @wTcTyp= 'IN' and    
       (exists(Select 'x' from InvDsg where IdInIdNo=del.InIdNo and IdTc= del.InTc and IdYy= del.InYy and IdChr= del.InChr     
         and IdNo= del.InNo and IdPrtKEy=del.InPrtKEy)     
       or exists(Select 'x' from InvFgd where IfInIdNo=del.InIdNo and IfTc= del.InTc and IfYy= del.InYy and IfChr= del.InChr     
         and IfNo= del.InNo and IfPrtKey=del.InPrtKey)     
       )    
       And     
       not exists (Select 'x' from inserted ins where ins.InIdNo=del.InIdNo and ins.InTc= del.InTc and ins.InYy= del.InYy and ins.InChr= del.InChr     
         and ins.InNo= del.InNo and ins.InPrtKey=del.InPrtKey and ins.InCmCd= del.InCmCd and ins.InMulDiv= del.InMulDiv     
         and convert (decimal (16, 4), ins.InCnvFct)= convert (decimal (16, 4), del.InCnvFct))     
      If @Row> 0    
      begin     
       Select @ErrMsg= 'Cannot Change Customer Code, Currency Conversion Rate As Inv Dsg/ Bag Rec Exists'     
       GoTo ErrHndlr    
      end    
      End    
    End    
  /*** (Jen 2.13) ***/     
      
  /*** Jay 3.3(FG) ***/    
  /******** If @wTcTyp 'FG' then Cannot Change InCmCd, InCnvFct, InMulDiv if Record exists in InvDsg, InvFgd ********/     
  /* Current Partition */    
  /* Manali Trading Module */    
  If @wTcTyp In ('FG', 'JT')    
  Begin    
     If Update(InCmCd) or Update(InRsMulDiv) or Update(InCnvRt)    
     Begin     
      select @Row= Count('x') from deleted del where del.InPrtKey=@wCurrPrtn and @wTcTyp In ('FG', 'JT') and     
       exists(Select 'x' from OrdMst where OmTc= del.InTc and OmYy= del.InYy and OmChr= del.InChr     
         and OmNo= del.InNo and OmPrtKEy=del.InPrtKEy) And     
       not exists (Select 'x' from inserted ins where ins.InIdNo=del.InIdNo and ins.InTc= del.InTc and ins.InYy= del.InYy and ins.InChr= del.InChr     
         and ins.InNo= del.InNo and ins.InPrtKey=del.InPrtKey and ins.InCmCd= del.InCmCd and ins.InMulDiv= del.InMulDiv     
         and convert (decimal (16, 4), ins.InCnvFct)= convert (decimal (16, 4), del.InCnvFct))     
      If @Row> 0    
      begin     
       Select @ErrMsg= 'Cannot Change Customer Code, Currency Conversion Rate As FG Out Rec Exists'     
       GoTo ErrHndlr    
      end    
    End      
   End    
 End    
    
    
 if @mode in ('A','C')    
 Begin    
    
  /*      
  ******  InHBnkInstr  =  Not Blank ******    
  select @row= count('x') from inserted where InHBnkInstr= ''     
  if @row>0     
  Begin    
   select @errmsg= 'Bank Instructions Cannot Be Blank'     
   goto ErrHndlr    
  end    
  */    
  /******  InCoCd = HCoCd   *******/    
 /* Current Partition */    
  If @mode='A' or (@mode='C' and (Update(InCoCd)))    
  Begin    
    select @row= count('x') from inserted where InPrtKey=@wCurrPrtn     
      and not exists (select 'x' from Head where HCoCd= InCoCd and HCd= @wHCd)    
    if @row>0     
    begin     
     select @errMsg= 'Company Code Not Valid'    
     goto ErrHndlr    
    end    
    End    
        
  /****** The Entry Type should Be 'IN' ******/    
  /* Current Partition */    
  /*** Jay 3.3(FG) ***/    
  /***** Manali trading Module *****/    
  If @mode='A' or (@mode='C' and (Update(InTc)))    
  Begin    
    select @row= count('x') from inserted where InPrtKey=@wCurrPrtn and    
     not exists (select 'x' from Param where PTyp= 'TC' and PMCd= InTc and     
          PValue in (Select PMCd from Param where PTyp='TCTYP' and PValue1 in ('IN','FG', 'JT')))     
    /*select @row= count('x') from inserted where InPrtKey=@wCurrPrtn and    
     not exists (select 'x' from Param where PTyp= 'TC' and PMCd= InTc and     
          PValue in (Select PMCd from Param where     
          PTyp='TCTYP' and PValue1= 'IN'))    
    */    
    if @row>0     
    begin    
     select @errmsg= 'Voucher Tc Not Defined'    
     goto ErrHndlr    
    end    
    End    
        
  /* Current Partition */    
  If @mode='A' or (@mode='C' and (Update(InYy)))    
  Begin    
    select @row=count('x') from inserted where InPrtKey=@wCurrPrtn and    
     not exists (select 'x' from Param where     
          PTyp='YY' and PMCd= InYy)    
    if @row>0     
    begin    
     select @errmsg='Voucher Year Not Defined'    
     goto ErrHndlr    
    end    
    End    
        
  /* Current Partition */    
  If @mode='A' or (@mode='C' and (Update(InTc) Or Update(InChr)))    
  Begin    
    select @row=count('x') from inserted where InPrtKey=@wCurrPrtn and    
     not exists (select 'x' from vParam where vPCoCd = InCoCd and     
       vPTyp= 'CHR' and vPMCd= InTc and vPSCd= InChr)    
    if @row>0     
    begin    
     select @errmsg='Voucher Character Not Defined'    
     goto ErrHndlr    
    end    
    End    
        
  /* Current Partition */    
  If @mode='A' or (@mode='C' and (Update(InNo)))    
  Begin    
    select @row=count('x') from inserted where InPrtKey=@wCurrPrtn and InNo<= 0     
    if @row>0     
    begin    
     select @errmsg='Voucher Number Cannot Be <= Zero'    
     goto ErrHndlr    
    end    
    End    
        
  /* Current Partition */    
  If @mode='A' or (@mode='C' and (Update(InExpNo)))    
  Begin    
    select @row=count('x') from inserted where InPrtKey=@wCurrPrtn and InExpNo= ''    
    if @row>0     
    begin    
     select @errmsg='Invoice Export Number Cannot Be Blank'    
     goto ErrHndlr    
    end    
    End    
        
  /* Emr 211 (InExpNo Key shld be unique) */     
  /****Crm Geeta****InExpNo should be unique***/      
  /* Complete Database */    
  If @mode='A' or (@mode='C' and (Update(InExpNo)))    
  Begin    
    select @row= count('x') from inserted ins where ins.InPrtKey=@wCurrPrtn and    
     (Select count('x') from InvHd InHd where /* InHd.InCoCd= ins.InCoCd and */     
      InHd.InExpNo=ins.InExpNo)> 1     
     And Exists (Select 'x' From SysDb Where SysId='CRM')     
      if @row> 0     
      begin    
        select @errmsg= 'InExpNo Key Should Be Unique'    
        goto ErrHndlr    
      end    
    
  /*** (Jen 2.14 Next) ***/    
  /**** InExpNo should be unique within a company ***/      
  /* Complete Database */    
    select @row= count('x') from inserted ins where ins.InPrtKey=@wCurrPrtn and    
     (Select count('x') from InvHd InHd where InHd.InCoCd= ins.InCoCd and     
      InHd.InExpNo=ins.InExpNo)> 1     
      if @row> 0     
      begin    
        select @errmsg= 'InExpNo Key Should Be Unique Within A Company'    
        goto ErrHndlr    
      end    
   End       
  /*** (Jen 2.14 Next) ***/    
    
    
  /****** InHRndOffFOB = PMCd('YN') ******/      
  /* Current Partition */    
  If @mode='A' or (@mode='C' and (Update(InHRndOffFOB)))    
  Begin    
    select @row=count('x') from inserted where InPrtKey=@wCurrPrtn and    
     not exists (select 'x' from Param where PTyp= 'YN' and PMCd= InHRndOffFOB)    
    if @row>0     
    begin    
     select @errmsg= 'Specify Whether To Round Off The Inv & Custom FOB Vals At The Dsg Level In Inv (Enter Y/N)'    
     goto ErrHndlr    
    end    
    End    
        
  /****** Sachin 3.3.0 (23-06-08) - [42. Make Invoice By Bag] ******/    
  If @mode='A' or (@mode='C' and (Update(InByBagYN)))    
  Begin    
    select @row=count('x') from inserted where InPrtKey=@wCurrPrtn and    
     not exists (select 'x' from Param where PTyp= 'YN' and PMCd=InByBagYN)    
    if @row>0     
    begin    
     select @errmsg= 'Specify Whether Invoicing to be done by Bag (Enter Y/N)'    
     goto ErrHndlr    
    end    
  End    
  /****** Sachin 3.3.0 (23-06-08) - [42. Make Invoice By Bag] ******/    
    
  /****** Manali 3.5.0 - 11/12/08 - Hide Commission YN ******/    
  If @mode='A' or (@mode='C' and (Update(InHHideCommYN)))    
  Begin    
    select @row=count('x') from inserted where InPrtKey=@wCurrPrtn and    
     not exists (select 'x' from Param where PTyp= 'YN' and PMCd=InHHideCommYN)    
    if @row>0     
    begin    
     select @errmsg= 'Specify Whether Commission Is Self Payable in Invoice(Enter Y/N) '    
     /*select @errmsg= 'Specify Whether Hide Commission in Invoice Report (Enter Y/N) '*/    
     goto ErrHndlr    
    end    
    End    
        
  /****** Manali 3.8.1 - Self Payble Discount YN ******/    
  If @mode='A' or (@mode='C' and (Update(InHSelfDscYN)))    
  Begin    
    select @row=count('x') from inserted where InPrtKey=@wCurrPrtn and    
     not exists (select 'x' from Param where PTyp= 'YN' and PMCd=InHSelfDscYN)    
    if @row>0     
    begin    
     select @errmsg= 'Specify Whether Discount Is Self Payable in Invoice(Enter Y/N) '    
     goto ErrHndlr    
    end    
    End    
        
  /****** Manali 3.8.1 - Self Payable Deduction YN ******/    
  If @mode='A' or (@mode='C' and (Update(InHSelfDeductYN)))    
  Begin    
    select @row=count('x') from inserted where InPrtKey=@wCurrPrtn and    
     not exists (select 'x' from Param where PTyp= 'YN' and PMCd=InHSelfDeductYN)    
    if @row>0     
    begin    
     select @errmsg= 'Specify Whether Deduction Is Self Payable in Invoice(Enter Y/N) '    
     goto ErrHndlr    
    end    
    End    
        
  /****** Sachin 3.3.0 (23-06-08) - [42. Make Invoice By Bag] ******/    
    
  /*** jAY 3.3(FG) ***/    
  /* Current Partition */    
  /***** Manali trading Module *****/    
  If @mode='A' or (@mode='C' and (Update(InCmCd)))    
  Begin    
    select @row= count('x') from inserted where InPrtKey=@wCurrPrtn and    
     not exists (select 'x' from CustMst where     
     CmCtg = (case when @wTcTyp in ('IN','FG') then 'C' when @wTcTyp= 'CTB' then 'T'     
            When @wTcTyp= 'JT' Then IsNull((Select OmCmCtg from OrdMst Where OmCoCd=InCoCd and OmTc=inTc and OmYy=InYy and OmChr=inChr and OmNo=InNo and OmPrtKey=InPrtKey), '')     
            else '' end) and CmCd= InCmCd)     
    if @row> 0     
    begin    
     select @errmsg= 'Customer/Contractor Code Not Defined'    
     goto ErrHndlr    
    end    
    End    
        
  /* Current Partition */    
  If @mode='A' or (@mode='C' and (Update(InDt)))    
  Begin    
    select @row= count('x') from inserted where InPRtKey=@wCurrPrtn and InDt in('01/01/1980', '01/01/1900')    
    if @row> 0     
    begin    
     select @errmsg= 'Date Cannot Be Blank'    
     goto ErrHndlr    
    end    
    End    
        
  /* Current Partition */    
  If @mode='A' or (@mode='C' and (Update(InPrnDt)))    
  Begin    
    select @row= count('x') from inserted where InPrtKey=@wCurrPrtn and InPrnDt in('01/01/1980', '01/01/1900')    
    if @row> 0     
    begin    
     select @errmsg= 'Invoice Printing Date Cannot Be Blank'    
     goto ErrHndlr    
    end    
    End    
        
  /* Current Partition */    
  /* **** Manali 3.03 - 21/06/08 - Palladium - LML added*/    
  /*** Jay 3.3(FG) ***/    
  /***** Manali trading Module *****/    
  If @mode='A' or (@mode='C' and (Update(InLmgCst) or Update(InLmpCst) or Update(InLmsCst) or Update(InLmlCst)))    
  Begin    
    select @row= count('x') from inserted     
     where @wTcTyp in ('IN','FG', 'JT') and InPrtKey=@wCurrPrtn and (InLmgCst<= 0 or InLmpCst<= 0 or (Exists    
     (Select 'x' From Head Where HCoCd=InCoCd And HSlvModYN='Y') And InLmsCst<= 0)    
     or (Exists (Select 'x' From Head Where HCoCd=InCoCd And HPdModYN='Y') And InLmlCst<= 0))     
    if @row> 0     
    begin    
     select @errmsg= 'LMG, LMP, LMS, LML Cost Cannot Be <= Zero'    
     goto ErrHndlr    
    end    
    End    
  /*** (Bef 2.13)     
  select @row= count('x') from inserted where InLmgDt in ('01/01/1980', '01/01/1900')     
  -- or InLmpDt in('01/01/1980', '01/01/1900')     
  if @row> 0     
  begin    
   select @errmsg= 'LMG Date Cannot Be Blank'    
   goto ErrHndlr    
  end    
  ***/    
     
  /* Current Partition */    
  /* **** Manali 3.03 - 21/06/08 - Palladium - LML added*/    
  /*** Jay 3.3(FG) ***/    
  /***** Manali trading Module *****/    
  If @mode='A' or (@mode='C' and (Update(InLmgSal) or Update(InLmpSal) or Update(InLmsSal) or Update(InLmlSal)))    
  Begin    
    select @row= count('x') from inserted where @wTcTyp in ('IN','FG', 'JT') and InPrtKey=@wCurrPrtn     
     and (InLmgSal<= 0 or InLmpSal<= 0 or     
     (Exists (Select 'x' From Head Where HCoCd=InCoCd And HSlvModYN='Y') And InLmsSal<= 0) or     
     (Exists (Select 'x' From Head Where HCoCd=InCoCd And HPdModYN='Y') And InLmlSal<= 0))     
    if @row> 0     
    begin    
     select @errmsg= 'LMG, LMP, LMS, LML Sales Value Cannot Be <= Zero'    
     goto ErrHndlr    
    end    
    End    
        
  /*select @row= count('x') from inserted where InVaDc<= 0     
  if @row> 0     
  begin    
   select @errmsg= 'Value Addition For Dia& ColStn Cannot Be <= Zero'    
   goto ErrHndlr    
  end    
    
  select @row= count('x') from inserted where InVaX< 0     
  if @row> 0     
  begin    
   select @errmsg= 'Value Addition For Accessories Cannot Be < Zero'    
   goto ErrHndlr    
  end    
  */    
      
  /* Current Partition */    
  If @mode='A' or (@mode='C' and (Update(InCnvRt)))    
  Begin    
    select @row= count('x') from inserted where InPrtKey=@wCurrPrtn and  convert( decimal (16, 4), InCnvRt)= convert( decimal (16, 4), 0)     
    if @row> 0     
    begin    
     select @errmsg= 'Rupee Conversion Rate Cannot Be Zero'    
     goto ErrHndlr    
    end    
    End    
        
  /*** (Jen 2.13) ***/     
  /* Current Partition */    
  If @mode='A' or (@mode='C' and (Update(InCnvFct)))    
  Begin    
    select @row= count('x') from inserted where InPrtKey=@wCurrPrtn and convert( decimal (16, 4), InCnvFct)= convert( decimal (16, 4), 0)     
    if @row> 0     
    begin    
     select @errmsg= 'Currency Conversion Rate Cannot Be Zero'    
     goto ErrHndlr    
    end    
    End    
        
  /* Current Partition */    
  If @mode='A' or (@mode='C' and (Update(InMulDiv)))    
  Begin    
    select @row= count('x') from inserted where InPrtKey=@wCurrPRtn and    
     not exists (select 'x' from Param where PTyp= 'MULDIV' and PMCd= InMulDiv)    
    if @row> 0     
    begin    
     select @errmsg=' Multiplication Factor For Customer Currency Not Defined'    
     goto ErrHndlr    
    end    
    End    
      
  /* Current Partition */    
  If @mode='A' or (@mode='C' and (Update(InRsMulDiv)))    
  Begin    
    select @row= count('x') from inserted where InPrtKey=@wCurrPrtn and    
     not exists (select 'x' from Param where PTyp= 'MULDIV' and PMCd= InRsMulDiv)    
    if @row> 0     
    begin    
     select @errmsg=' Multiplication Factor For Rupee Not Defined'    
     goto ErrHndlr    
    end    
  End    
  /*** (Jen 2.13) ***/    
    
  /* Current Partition */    
  If @mode='A' or (@mode='C' and (Update(InTotAs)))    
  Begin    
    select @row= count('x') from inserted where InPrtKey=@wCurrPrtn and    
     not exists (select 'x' from Param where PTyp= 'INVTOTAS' and PMCd= InTotAs)    
    if @row> 0     
    begin    
     select @errmsg=' ''Invoice Total As'' Option Not Defined'    
     goto ErrHndlr    
    end    
  End    
    
  /* Sachin 3.11.0 - InTotAsTxt Cannot be Blank */    
  If @mode='A' or (@mode='C' and (Update(InTotAsTxt)))    
  Begin    
    select @row= count('x') from inserted where InPrtKey=@wCurrPrtn and LTrim(RTrim(InTotAsTxt))=''    
    if @row> 0     
    begin    
     select @errmsg=' ''Invoice Total As Text'' Cannot be Blank'    
     goto ErrHndlr    
    end    
  End    
        
  /* Current Partition */    
  If @mode='A' or (@mode='C' and (Update(InCls)))    
  Begin    
    select @row=count('x') from inserted where InPrtKey=@wCurrPrtn and    
     not exists (select 'x' from Param where PTyp= 'YN' and PMCd= InCls)    
    if @row> 0     
    begin    
     select @errmsg= ' ''Invoice Closed'' Option Not Defined (Enter Y/ N)'    
     goto ErrHndlr    
    end    
    End    
        
  /* Current Partition */    
  If @mode='A' or (@mode='C' and (Update(InPrdDesc)))    
  Begin    
    select @row=count('x') from inserted where InPrtKey=@wCurrPrtn and InPrdDesc= ''    
    if @row>0     
    begin    
     select @errmsg= 'Product Description Cannot Be Blank'    
     goto ErrHndlr    
    end    
    End    
        
  /*** (Jen 3.1.0_1) ***/    
  /* Current Partition */    
  If @mode='A' or (@mode='C' and (Update(InDsgOnHld)))    
  Begin    
    select @row= count('x') from inserted where InPrtKey=@wCurrPrtn and    
     not exists (select 'x' from Param where PTyp= 'YN' and PMCd= InDsgOnHld)     
    if @row> 0     
    begin    
     select @errmsg= ' ''Design On Hold'' Option Not Defined (Enter Y/ N)'    
     goto ErrHndlr    
    end    
    End    
        
  /* Current Partition */    
  If @mode='A' or (@mode='C' and (Update(InRemakingCtgs)))    
  Begin    
    select @row= count('x') from inserted where InPrtKey=@wCurrPrtn and InRemakingCtgs<> '' and     
     not exists (select 'x' from Param where PTyp= 'REMAKE' and PMCd= InRemakingCtgs)     
    if @row> 0     
    begin    
     select @errmsg= 'Remaking Rm Category Option Not Defined'    
     goto ErrHndlr    
    end    
    End    
        
  /* Current Partition */    
  If @mode='A' or (@mode='C' and (Update(InInvGrsWtEqActual)))    
  Begin    
    Select @row= count('x') from inserted where InPrtKey=@wCurrPrtn and    
     not exists (select 'x' from Param where PTyp= 'YN' and PMCd= InInvGrsWtEqActual)     
    if @row> 0     
    begin     
     select @errmsg= 'Specify Whether To Keep The Invoice Gross Weight Equal To Actual Gross Weight During Copy (Y/N)'     
     goto ErrHndlr     
    end    
  End    
  /*** (Jen 3.1.0_1) ***/     
    
    
--  Manoj FOB     
/* (Shifted to Not For Repl Trg)    
  select @row= count('x') from inserted where @wTcTyp= 'IN'     
   and Str(InFOBRs, 16, 0)<> Str(Round(InFOB* inCnvRt, 2), 16, 0)     
  if @row> 0     
  begin    
   select @errmsg= 'FOB In Rs Is Not Proper'    
   goto ErrHndlr    
  end    
*/    
--  Manoj FOB    
    
/*   select @row= count('x') from inserted where InComm<> 0 and Round(InCommRs, 2)<> Round(InComm* inCnvRt, 2)    
  if @row> 0     
  begin    
   select @errmsg= 'Commision In Rs Is Not Proper'    
   goto ErrHndlr    
  end    
    
  select @row= count('x') from inserted where InFrt<> 0 and Round(InFrtRs, 2)<> Round(InFrt* inCnvRt, 2)    
  if @row> 0     
  begin    
   select @errmsg= 'Freight In Rs Is Not Proper'    
   goto ErrHndlr    
  end    
    
  select @row= count('x') from inserted where InIns<> 0 and Round(InInsRs, 2)<> Round(InIns* inCnvRt, 2)    
  if @row> 0     
  begin    
   select @errmsg= 'Insurance In Rs Is Not Proper'    
   goto ErrHndlr    
  end    
*/   
    
/* Bef Jay 2.14 (LCLINV)    
  select @row= count('x') from inserted where     
   Str(InTot, 16, 2)<> Str((case InTotAs when 'FOB' then InFOB when 'CIF' then InFOB+ InComm+ InFrt+ InIns     
    when 'C&F' then InFOB+ InComm+ InFrt when 'C&I' then InFOB+ InComm+ InIns else 0 end), 16, 2)    
  if @row>0     
  begin    
   select @errmsg= 'Total Value Is Not Proper'    
   goto ErrHndlr    
  end    
Bef Jay 2.14 (LCLINV) */    
    
/* *** Jay 3.1.1 [Discount] *** */    
  /* Current Partition */    
  /*** Jay 3.3(FG) ***/    
  /* ***** Manali 3.5.0 - 11/12/08 - INHHideCommYN Case added ***** */    
  /***** Manali trading Module *****/    
  /***** Manali 3.8.1 - Self Payable Options Added *****/    
  If @mode='A' or (@mode='C' and (Update(InTot) or Update(InTotAs) or Update(InFOB) or Update(InHHideCommYN) or Update(InComm)    
                  or Update(InFrt) or Update(InIns) or Update(InHSelfDscYN) or Update(InDsc)    
                   or Update(InHSelfDeductYN) or Update(InDeduct)))    
  Begin    
    Select @row= count('x') from inserted where @wTcTyp in ('IN','FG', 'JT')  and InPrtKey=@wCurrPrtn and     
     convert (decimal (16, 2), InTot)<>     
      (Case InLclInvYn when 'N' then     
        convert (decimal (16, 2), (case InTotAs when 'FOB' then InFOB +(case when InHSelfDscYN ='Y' then 0 Else InDsc End) +(case when InHSelfDeductYN ='Y' then 0 Else InDeduct End)    
        when 'CIF' then InFOB+ (case when InHHideCommYN ='Y' then 0 Else InComm End)+ InFrt+ InIns +(case when InHSelfDscYN ='Y' then 0 Else InDsc End) +(case when InHSelfDeductYN ='Y' then 0 Else InDeduct End)    
        when 'C&F' then InFOB+ (case when InHHideCommYN ='Y' then 0 Else InComm End)+ InFrt +(case when InHSelfDscYN ='Y' then 0 Else InDsc End) + (case when InHSelfDeductYN ='Y' then 0 Else InDeduct End)    
        when 'C&I' then InFOB+ (case when InHHideCommYN ='Y' then 0 Else InComm End)+ InIns +(case when InHSelfDscYN ='Y' then 0 Else InDsc End) + (case when InHSelfDeductYN ='Y' then 0 Else InDeduct End)    
        else 0 end))     
      when 'Y' then     
        convert (decimal (16, 2), InFob+ InLclVatRs+ InLclCstRs+ InLclOctroiRs +InLclOthrValRs)    
      else 0 end)    
    if @row>0     
    begin    
     select @errmsg= 'Total Value Is Not Proper'    
     goto ErrHndlr    
    end    
    End    
  /**** Bef 3.8.1 ****/    
  /*Select @row= count('x') from inserted where @wTcTyp in ('IN','FG', 'JT')  and InPrtKey=@wCurrPrtn and     
   convert (decimal (16, 2), InTot)<>     
     (Case InLclInvYn when 'N' then     
      (case when InHHideCommYN ='Y' then    
       convert (decimal (16, 2), (case InTotAs when 'FOB' then InFOB +InDsc    
                when 'CIF' then InFOB+ InFrt+ InIns +InDsc    
                when 'C&F' then InFOB+ InFrt +InDsc    
                when 'C&I' then InFOB+ InIns +InDsc    
                else 0 end))     
      Else    
       convert (decimal (16, 2), (case InTotAs when 'FOB' then InFOB +InDsc    
                when 'CIF' then InFOB+ InComm+ InFrt+ InIns +InDsc    
                when 'C&F' then InFOB+ InComm+ InFrt +InDsc    
                when 'C&I' then InFOB+ InComm+ InIns +InDsc    
                else 0 end))     
      End)    
     when 'Y' then convert (decimal (16, 2), InFob+ InLclVatRs+ InLclCstRs+ InLclOctroiRs +InLclOthrValRs)     
     else 0 end)    
  if @row>0     
  begin    
   select @errmsg= 'Total Value Is Not Proper'    
   goto ErrHndlr    
  end*/    
    
/* *** Jay 3.1.1 [Discount] *** */    
    
    
/*    
  select @row=count('x') from inserted where Round(InTotRs, 2)<> Round(InTot* inCnvRt, 2)    
  if @row>0     
  begin    
   select @errmsg= 'Total Value In Rs Is Not Proper'    
   goto ErrHndlr    
  end    
*/    
    
  /*** (Bef 2.13)     
  select @row= count('x') from inserted where @wTcTyp= 'IN' and     
   Str(InTotRs, 16, 2)<> Str((case InTotAs when 'FOB' then InFOBRs when 'CIF' then InFOBRs+ InCommRs+ InFrtRs+ InInsRs     
    when 'C&F' then InFOBRs+ InCommRs+ InFrtRs when 'C&I' then InFOBRs+ InCommRs+ InInsRs else 0 end), 16, 2)    
  (Bef 2.13) ***/    
    
  /*** (Jen 2.13) ***/    
  /* *** Jay 2.14(LclInv) *** */    
  /* Current Partition */    
  /*** Jay 3.3(FG) ***/    
  /***** Manali trading Module *****/    
  /***** Manali 3.8.1 - Self Payable Options Added *****/    
  If @mode='A' or (@mode='C' and (Update(InTotRs) or Update(InTotAs) or Update(InFOBRs) or Update(InHHideCommYN) or Update(InCommRs)    
                  or Update(InFrtRs) or Update(InInsRs) or Update(InHSelfDscYN) or Update(InDscRs)    
                   or Update(InHSelfDeductYN) or Update(InDeductRs)))    
  Begin    
    select @row= count('x') from inserted where @wTcTyp in ('IN','FG', 'JT') and InLclInvYn = 'N' and InPrtKEy=@wCurrPrtn and    
     (convert (decimal (16, 2), InTotRs))<>     
     convert (decimal (16, 2), (case InTotAs when 'FOB' then InFOBRs +(case when InHSelfDscYN ='Y' then 0 Else InDscRs End) +(case when InHSelfDeductYN ='Y' then 0 Else InDeductRs End)    
      when 'CIF' then InFOBRs+ (case when InHHideCommYN ='Y' then 0 Else InCommRs End)+ InFrtRs+ InInsRs +(case when InHSelfDscYN ='Y' then 0 Else InDscRs End) +(case when InHSelfDeductYN ='Y' then 0 Else InDeductRs End)    
      when 'C&F' then InFOBRs+ (case when InHHideCommYN ='Y' then 0 Else InCommRs End)+ InFrtRs +(case when InHSelfDscYN ='Y' then 0 Else InDscRs End) + (case when InHSelfDeductYN ='Y' then 0 Else InDeductRs End)    
      when 'C&I' then InFOBRs+ (case when InHHideCommYN ='Y' then 0 Else InCommRs End)+ InInsRs +(case when InHSelfDscYN ='Y' then 0 Else InDscRs End) + (case when InHSelfDeductYN ='Y' then 0 Else InDeductRs End)    
      else 0 end))    
    if @row> 0     
    begin    
     select @errmsg= 'Total Value In Rs Is Not Proper'    
     goto ErrHndlr    
    end    
    End    
        
  /*** Bef 3.8.1 ***/    
  /*select @row= count('x') from inserted where @wTcTyp in ('IN','FG', 'JT') and InLclInvYn = 'N' and InPrtKEy=@wCurrPrtn and    
   (convert (decimal (16, 2), InTotRs))<>     
     (case when InHHideCommYN='Y' then (convert (decimal (16, 2), (case InTotAs when 'FOB' then InFOBRs +InDscRs    
               when 'CIF' then InFOBRs+ InFrtRs+ InInsRs +InDscRs     
               when 'C&F' then InFOBRs+ InFrtRs +InDscRs    
               when 'C&I' then InFOBRs+ InInsRs +InDscRs    
               else 0 end)))    
     Else (convert (decimal (16, 2), (case InTotAs when 'FOB' then InFOBRs +InDscRs    
               when 'CIF' then InFOBRs+ InCommRs+ InFrtRs+ InInsRs +InDscRs     
               when 'C&F' then InFOBRs+ InCommRs+ InFrtRs +InDscRs    
               when 'C&I' then InFOBRs+ InCommRs+ InInsRs +InDscRs    
               else 0 end)))    
     End)    
  if @row> 0     
  begin    
   select @errmsg= 'Total Value In Rs Is Not Proper'    
   goto ErrHndlr    
  end*/     
    
  /* (Bef)    
  select @row= count('x') from inserted where @wTcTyp= 'IN' and InLclInvYn = 'Y' and     
   convert (decimal (16, 0), InTotRs)<> convert (decimal (16, 0),     
      (((InFob * (1 + (InCmOctroiPer / 100))) + InLclOthrValRs) * (((InHVatPer +InHCstPer )/100 ) + 1))     
       * (case when InRsMulDiv= 'M' then inCnvRt else 1 end)     
       / (case when InRsMulDiv= 'D' then inCnvRt else 1 end) )    
  (Bef) */    
    
  /*** Jen 2.14(LclInv) ***/     
  /* Current Partition */    
  /*** Jay 3.3(FG) ***/    
  /***** Manali trading Module *****/    
  If @mode='A' or (@mode='C' and (Update(InLclInvYn) or Update(InTotRs) or Update(InFOB) or Update(InLclOctroiRs)     
                  or Update(InLclVatRs) or Update(InLclCstRs) or Update(InLclOthrValRs) or Update(InRsMulDiv)    
                   or Update(inCnvRt)))    
  Begin    
    select @row= count('x') from inserted where @wTcTyp in ('IN','FG', 'JT') and InLclInvYn = 'Y' and InPrtKey=@wCurrPrtn and    
     convert (decimal (16, 0), InTotRs)<> convert (decimal (16, 0),     
              (InFob+ InLclOctroiRs+ InLclVatRs+ InLclCstRs+ InLclOthrValRs)     
              * (case when InRsMulDiv= 'M' then inCnvRt else 1 end)     
            / (case when InRsMulDiv= 'D' then inCnvRt else 1 end) )     
    /*** Jen 2.14(LclInv) ***/     
    /* *** Jay 2.14(LclInv) *** */    
    /*** (Jen 2.13) ***/    
    
    if @row> 0     
    begin    
     select @errmsg= 'Total Value In Rs Is Not Proper'    
     goto ErrHndlr    
    end     
    End    
        
  /********** If @wTcTyp= 'CTB' then IrFrRmLoc, IrFrRmDc, IrTORmLoc, IrToRmDc     
       should be blank *************/    
  /* Current Partition */    
  If @wTcTyp= 'CTB'     
  Begin    
    If @mode='A' or (@mode='C' and (Update(InFrRmLoc) or Update(InFrRmDc) or Update(InToRmLoc) or Update(InToRmDc)))    
    Begin    
      Select @Row= Count('x') From Inserted Ins Where @wTcTyp= 'CTB' and InPrtKEy=@wCurrPrtn and    
       (InFrRmLoc<> '' or InFrRmDc<> '' or InToRmLoc<> '' or InToRmDc<> '')    
        If @Row> 0    
      begin    
       Select @ErrMsg= 'From/To Rm Loc, From/To Rm DC Should Be Blank For CTB'    
         GoTo ErrHndlr    
      end    
      End    
    End        
    
  /********** InFrRmDc  Valid 'DC' *************/    
  /* Current Partition */    
  /*** Jay 3.3(FG) ***/    
  /***** Manali trading Module *****/    
   If @mode='A' or (@mode='C' and (Update(InFrRmDc) or Update(InToRmDc)))    
   Begin    
    Select @Row= Count('x') From Inserted Ins     
     Where @wTcTyp in ('IN','FG','JT') and InPrtKey=@wCurrPrtn and    
      (Not Exists (Select 'x' From Param where PTyp= 'DC'    
       and PMCd = Ins.InFrRmDc and PSCd= '')    
      or Not Exists (Select 'x' From Param where PTyp= 'DC'    
       and PMCd = Ins.InToRmDc and PSCd= ''))    
      If @Row> 0    
    begin    
     Select @ErrMsg= 'Invalid From/ To Rm DC'    
       GoTo ErrHndlr    
    end    
        
        
    /******* InToRmDc should be Opp of  InFrRmDc *******/    
    /* Current Partition */    
      Select @Row=Count('x') From Inserted Ins    
     Where InPrtKey=@wCurrPrtn and ((InToRmDc= InFrRmDc And InFrRmDc<> '') Or     
      (InFrRmDc<> '' and InToRmDc= '') Or     
      (InFrRmDc= '' and InToRmDc<> ''))     
      If @Row>0    
    begin    
     Select @ErrMsg= 'Invalid To Rm Dc'    
     GoTo ErrHndlr    
    end    
    End    
        
  /****** InFrRmLoc or InToRmLoc Should Have LocTyp= 'XS' ********/    
  /* Current Partition */    
  /*** Jay 3.3(FG) ***/    
   If @mode='A' or (@mode='C' and (Update(InFrRmLoc) or Update(InToRmLoc)))    
   Begin    
    Select @Row= Count('x') From Inserted Ins where @wTcTyp in ('IN','FG')  and InPrtKey=@wCurrPrtn     
     and Not Exists (Select 'x' From Loc where LocCoCd = InCoCd     
      and LocTyp= 'XS' and (LocCd= InFrRmLoc or LocCd = InToRmLoc))    
      If @Row> 0    
    begin    
     Select @ErrMsg= 'Atleast One RM Loc Must Be Of Type ''XS'' '    
       GoTo ErrHndlr    
    end    
    
    /****** InFrRmLoc or IrToRmLoc Should Have LocTyp= 'F' ********/    
    /* Current Partition */    
    /*** Jay 3.3(FG) ***/    
    Select @Row= Count('x') From Inserted Ins where @wTcTyp in ('IN','FG')  and InPrtKey=@wCurrPrtn     
     and Not Exists (Select 'x' From Loc where LocCoCd = InCoCd     
      and LocTyp= 'F' and (LocCd= InFrRmLoc or LocCd = InToRmLoc))    
        If @Row>0    
    begin    
     Select @ErrMsg= 'Atleast One RM Loc Must Be Of Type ''F'' '    
       GoTo ErrHndlr    
    end    
    End    
        
  /***** Manali Trading Module *****/    
  If @wTcTyp= 'JT'     
  Begin    
    If @mode='A' or (@mode='C' and (Update(InFrRmLoc) or Update(InFrRmDc) or Update(InToRmLoc) or Update(InToRmDc)))    
    Begin    
      Select @Row= Count('x') From Inserted Ins where @wTcTyp = 'JT'  and InPrtKey=@wCurrPrtn     
       and not exists(Select 'x' from OrdMst Where OmCoCd=InCoCd and OmTc=InTc and OmYy=InYy and OmChr=InChr and OmNo=InNo and OmPrtKey=InPrtKey     
              and OmFrDmLoc=InFrRmLoc and OmToDmLoc=InToRmLoc and OmFrDmDc=InFrRmDc and OmToDmDc=InToRmDc)    
        If @Row> 0    
      begin    
       Select @ErrMsg= 'From Loc/DC and To Loc/Dc should be same as Order Voucher  '    
         GoTo ErrHndlr    
      end    
      End    
    End    
          
  /****** RM Wt/ Qty Can Be Issued From Locations Of Type 'F' Only ********/    
  /* Current Partition */    
  /*** Jay 3.3(FG) ***/    
  If @wTcTyp in ('IN','FG')    
  Begin    
    If @mode='A' or (@mode='C' and (Update(InFrRmLoc) or Update(InFrRmDc) or Update(InToRmLoc) or Update(InToRmDc)))    
    Begin    
      Select @Row= Count('x') From Inserted Ins where @wTcTyp in ('IN','FG') and InPrtKey=@wCurrPrtn and    
       Not Exists (Select 'x' From Loc where LocCoCd = InCoCd and     
        LocTyp= 'F' and ( (LocCd= InFrRmLoc and InFrRmDc= 'C') or     
           (LocCd= InToRmLoc and InToRmDc= 'C')                   ))    
          If @Row> 0    
      begin    
       Select @ErrMsg= 'RM Wt/ Qty Can Be Issued From Locations Of Type ''F'' Only'    
         GoTo ErrHndlr    
      end    
      End    
    End    
          
  /*** Shilpa Emr 2.04 ***/    
  /* Current Partition */    
   If @mode='A' or (@mode='C' and (Update(InhInvTyp)))    
   Begin    
    select @row= count('x') from inserted where InPRtKey=@wCurrPrtn and Not Exists (select 'x'     
     from Param where PTyp= 'INVTYP' and PMCd= InhInvTyp)     
    if @row> 0     
    begin    
     select @errmsg= 'Invoice Type Not Defined'    
     goto ErrHndlr    
    end    
   End    
    
   If @mode='A' or (@mode='C' and (Update(InRtFrInvCust)))    
   Begin    
    select @row= count('x') from inserted where InPRtKey=@wCurrPrtn and Not Exists (select 'x' from Param     
      where PTyp= 'YN' and PMCd= InRtFrInvCust)     
    if @row> 0     
    begin    
     select @errmsg= 'Invoice Rate from Invoice Customer Not Defined'    
     goto ErrHndlr    
    end    
   End    
        
  /* Current Partition */    
   If @mode='A' or (@mode='C' and (Update(InHValAddMtlColYN)))    
   Begin    
    select @row= count('x') from inserted where InPrtKey=@wCurrPrtn and Not Exists (select 'x' from Param     
      where PTyp= 'YN' and PMCd= InHValAddMtlColYN)     
    if @row> 0     
    begin    
     select @errmsg= 'Show Value Addition On Metal Column In Value Addition Report? Enter Yes Or No (Y/N)?'    
     goto ErrHndlr    
    end    
    End    
        
  /* Current Partition */    
   If @mode='A' or (@mode='C' and (Update(InHValBOEDetYN)))    
   Begin    
    select @row= count('x') from inserted where InPrtKey=@wCurrPrtn and    
     Not Exists (select 'x' from Param where PTyp= 'YN' and PMCd= InHValBOEDetYN)     
    if @row> 0     
    begin    
     select @errmsg= 'Show BOE Details In Value Addition Annexur Report? Enter Yes Or No (Y/N)?'    
     goto ErrHndlr    
    end    
  End    
  /*** Shilpa Emr 2.04***/    
    
  /*** (Jen 2.13) ***/    
  /* Current Partition */    
   If @mode='A' or (@mode='C' and (Update(InShowGldLs)))    
   Begin    
    select @row= count('x') from inserted where InPrtKey=@wCurrPrtn and    
     Not Exists (select 'x' from Param where PTyp= 'YN' and PMCd= InShowGldLs)     
    if @row> 0     
    begin    
     select @errmsg= 'Show Gold Loss % In Invoice Reports? Enter Yes Or No (Y/N)?'    
     goto ErrHndlr    
    end    
  End    
  /*** (Jen 2.13) ***/    
    
  /* *** Jay 2.14(LclInv) *** */    
  /* Current Partition */    
   If @mode='A' or (@mode='C' and (Update(InLclInvYn)))    
   Begin    
    select @row= count('x') from inserted where InPrtKey=@wCurrPrtn and    
     Not Exists (select 'x' from Param where PTyp= 'YN' and PMCd= InLclInvYn)     
    if @row> 0     
    begin    
     select @errmsg= 'Invalid Local Invoice (Y/N)?'    
     goto ErrHndlr    
    end    
        
  /* Current Partition */    
    Select @row= count('x') from inserted ins where ins.InLclInvYn = 'Y' and ins.InPrtKey=@wCurrPRtn and    
      Not Exists (Select 'x' from Head where HCocd = ins.InCoCd and HLclInvYn= 'Y')    
    if @row> 0     
    begin    
     select @errmsg= 'Local Invoice Module Is Not Available'    
     goto ErrHndlr    
    End      
    End    
        
  /* Current Partition */    
   If @mode='A' or (@mode='C' and (Update(InHVatPer) or Update(InHCstPer) or Update(InLclInvYn)))    
   Begin    
    Select @row= count('x') from inserted where InPrtKey=@wCurrPrtn and InHVatPer <> 0 and InHCstPer <> 0 and InLclInvYn = 'Y'     
    if @row> 0     
    begin    
     select @errmsg= 'VAT and CST, Both Percentage Can Not Be Aplicable'    
     goto ErrHndlr    
    end    
  End      
        
  /* Current Partition */    
   If @mode='A' or (@mode='C' and (Update(InLclVatRs) or Update(InLclCstRs) or Update(InLclInvYn)))    
   Begin    
    Select @row= count('x') from inserted where InPrtKey=@wCurrPrtn and InLclVatRs <> 0 and InLclCstRs <> 0 and InLclInvYn = 'Y'     
    if @row> 0     
    begin    
     select @errmsg= 'VAT and CST, Both Values Can Not Be Aplicable'    
     goto ErrHndlr    
    end    
  End    
  /* *** Jay 2.14(LclInv) *** */    
    
  /* ***** Manali 3.6.0 - 30/10/09 - Use Notional Rate in Custom Matching ****** */    
   If @mode='A' or (@mode='C' and (Update(InNotionRtYn)))    
   Begin    
    select @row= count('x') from inserted where InPrtKey=@wCurrPrtn and    
     Not Exists (select 'x' from Param where PTyp= 'YN' and PMCd= InNotionRtYn)     
    if @row> 0     
    begin    
     select @errmsg= 'Use Notional Rate in Custom Matching (Y/N)?'    
     goto ErrHndlr    
    end    
    
/*    select @row= count('x') from inserted where InPrtKey=@wCurrPrtn and    
     InNotionRtYn='Y' and not Exists (select 'x' from Head where HCoCd= InCoCd and HCd= @wHCd and HExpLocal='Y')     
    if @row> 0     
    begin    
     select @errmsg= 'Use Notional Rate Option Cannot Be ''Y''es When DTA is No'    
     goto ErrHndlr    
    end */    
  End      
  /* ***** Manali 3.6.0 - 30/10/09 - Use Notional Rate in Custom Matching ****** */    
      
  /* ***** Sachin 3.11.0 - Use Markup for Gold on Stock Rate ****** */    
   If @mode='A' or (@mode='C' and (Update(InMetIRtwrtStkYN)))    
   Begin    
    Select @row= count('x') from inserted where InPrtKey=@wCurrPrtn and    
     Not Exists (select 'x' from Param where PTyp= 'YN' and PMCd=InMetIRtwrtStkYN)     
    if @row> 0     
    begin    
     select @errmsg= 'Use Metal Invoice Rate wrt Stock Rate should be (Y/N)?'    
     goto ErrHndlr    
    end    
    
      /* ***** Sachin 3.11.0 - Invoice RM Rate from Rate Chart ****** */    
    Select @row= count('x') from inserted where InPrtKey=@wCurrPrtn and    
     Not Exists (select 'x' from Param where PTyp= 'YN' and PMCd=InRmRtfrRtChrt)     
    if @row> 0     
    begin    
     select @errmsg= 'Use Invoice RM Rate from Rate Chart should be (Y/N)?'    
     goto ErrHndlr    
    end    
    
      /* ***** Sachin 3.11.0 - Invoice Labor Rate from Rate Chart ****** */    
    Select @row= count('x') from inserted where InPrtKey=@wCurrPrtn and    
     Not Exists (select 'x' from Param where PTyp= 'YN' and PMCd=InLabRtfrRtChrt)     
    if @row> 0     
    begin    
     select @errmsg= 'Use Invoice Labor Rate from Rate Chart should be (Y/N)?'    
     goto ErrHndlr    
    end    
          
    Select @row= count('x') from inserted where InPrtKey=@wCurrPrtn and    
     Not Exists (select 'x' from Param where PTyp= 'YN' and PMCd=InXIRtwrtStkYN)     
    if @row> 0     
    begin    
     select @errmsg= 'Use Accessories Invoice Rate wrt Stock Rate should be (Y/N)?'    
     goto ErrHndlr    
    end    
    
    Select @row= count('x') from inserted where InPrtKey=@wCurrPrtn and    
     Not Exists (select 'x' from Param where PTyp= 'YN' and PMCd=InIRtwrtStkYN)     
    if @row> 0     
    begin    
     select @errmsg= 'Use Diamond & Color Stones Invoice Rate wrt Stock Rate should be (Y/N)?'    
     goto ErrHndlr    
    end    
  End      
  /* ***** Sachin 3.11.0 - Use Markup for Gold on Stock Rate ****** */    
 End    
    
 -- 4.3.0.0    
 If @wTcTyp ='IN' And @mode = 'C' and Update(InDt)    
 Begin     
  Update InvRm Set IrDt=InDt From Inserted Where InCoCd=IrCoCd And InTc=IrTc And InYy=IrYy And InChr=IrChr And     
     InNo=IrNo And InDt<>IrDt    

  Update InvFgd Set ModUsr='MIC' From Inserted where Inserted.InCoCd=IfCoCd And Inserted.InTc=IfTc And Inserted.InYy=IfYy And Inserted.InChr=IfChr And     
     Inserted.InNo=IfNo and Inserted.InDt <> (select InDt from Deleted where Deleted.InCoCd = inserted.InCoCd and Deleted.InYy = Deleted.InYy and
											Deleted.InTc = Inserted.InTc and Deleted.InNo = Inserted.InNo)
 End     
    
 if @mode = 'D'     
 Begin    
  /****** Cannot Delete If Record Exists In InvDsg Table ******/    
  /*** Jay 3.3(FG) ***/      
  /* Current Partition */    
  select @row= count('x') from deleted, InvDsg where @wTcTyp= 'IN' and InPrtKey=@wCurrPrtn and InIdNo=IdInIdNo and InPrtKey=IdPRtKey and InCoCd= IdCoCd and InTc= IdTc     
   and InYy= IdYy and InChr= IdChr and InNo= IdNo     
  if @row> 0     
  begin    
   select @errmsg= 'Cannot Delete Invoice As Invoice Design Record Exists'    
   goto ErrHndlr    
  end    
      
  /****** Cannot Delete If Record Exists In OrdMst Table ******/    
  /*** Jay 3.3(FG) ***/      
  /* Current Partition */    
  select @row= count('x') from deleted, OrdMst where @wTcTyp= 'FG' and InPrtKey=@wCurrPrtn and InPrtKey=OmPRtKey and InCoCd= OmCoCd and InTc= OmTc     
   and InYy= OmYy and InChr= OmChr and InNo= OmNo     
  if @row> 0     
  begin    
   select @errmsg= 'Cannot Delete Invoice As Fg Out Transaction Record Exists'    
   goto ErrHndlr    
  end    
    
  /***** Manali Trading Module *****/    
  select @row= count('x') from deleted, OrdMst where @wTcTyp= 'JT' and InPrtKey=@wCurrPrtn and InPrtKey=OmPRtKey and InCoCd= OmCoCd and InTc= OmTc     
   and InYy= OmYy and InChr= OmChr and InNo= OmNo     
  if @row> 0     
  begin    
   select @errmsg= 'Cannot Delete Invoice As Master Record Exists'    
   goto ErrHndlr    
  end    
    
  /****** Cannot Delete If Record Exists In InvFgd Table ******/    
  /* Current Partition */    
  select @row= count('x') from deleted, InvFgd where InPrtKey=@wCurrPrtn and InIdNo=IfInIdNo and InPrtKey=IfPrtKey and InCoCd= IfCoCd and InTc= IfTc     
   and InYy= IfYy and InChr= IfChr and InNo= IfNo     
  if @row> 0     
  begin    
   select @errmsg= 'Cannot Delete Invoice As Invoice Bag Record Exists'    
   goto ErrHndlr    
  end    
 End    
    
 return    
ErrHndlr:    
 begin    
  select @errmsg='|'+@errmsg   raiserror(@errmsg,16,1)    
  rollback transaction     
 end    
end

GOGO


ALTER TRIGGER [dbo].[UsrOpt_INUPDL] ON [dbo].[UsrOpt]
               FOR INSERT,UPDATE,DELETE NOT FOR REPLICATION
AS
BEGIN
	/****** Trigger Version - 3.11.1 ******/
	DECLARE @ROW INT
	declare @mode char(1)
	declare @errmsg varchar(255)
	DECLARE @SYS VARCHAR(5)
	DECLARE @wSelfCd VARCHAR(8)
	DECLARE @wSelfCoCd varchar(3)
	 DECLARE @ReplAgt BIT    

	set @wSelfCoCd= 'ZZZ'
	set @wSelfCd = 'ZSELF'	

	IF @@ROWCOUNT=0
		RETURN
	If Exists(Select 'x' From TrgDisa Where TdTblNm = 'UsrOpt' )
		RETURN

	Set @ReplAgt=Convert(BIT,SessionProperty('Replication_Agent'))  

	If @ReplAgt = 1         Return

	Set @SYS= IsNull((Select PMCd from Param where PTyp= 'SYS' and PMCd= 'XEMR' and PSCd= ''), '') 

	select @mode=case (select count('x') from inserted) when 0 then 'D'
	      else
		case (select count('x') from deleted) when 0 then 'A' else 'C' end
	     end

	
	/****** added RepTxnRetMem, RepVchPrnRetMem (Jen 2.14 Next) ******/
	/****** added RepGRFormBackMulti, RepWipBRmBkDt (Jen 2.14) ******/
	/****** added RepGSPFORMMulti, RepGSPDetLetMulti, RepGSPCertMulti, RepDeclnLetMulti (Jen 3.01) ******/
	/* **** Zubin 211 'RepFltBagPrn' added **** */ 
	/* ****** Sachin 2.14.0 'RepWIWExp', 'RepWIWLongExp', 'RepWIWPrd', 'RepWIWLongPrd', 'RepWipBPcs','RepWipFgBPcs'  Removed ****** */	
	/* 'RepWipBRm', 'RepWipFgBRm', 'RepWipBPcsBkDt', 'RepWipFgBPcsBkDt','RepStkLed','RepFlashStk','RepSalAnaDsg', 'RepSalAnaRm','RepSalAnaLab', 'RepStkSumm',  */
	/* 'RepGldLs', 'RepBrkMsg', 'RepRejAna', RepWipBRmBkDt  */
	/* (Jen Bef 3.01) 'RepLoc', */
	/* Zubin 3.01.01 (260308) 'RepTravExcel' */
	/*** Jay 3.4(FG)***/
	/* ***** Manali Trading Module ***** */
	/**** Manali 3.7.1 - RepJtPndRm added ****/
	/**** Manali 3.8.0 - 'RepValAddnDta' added ****/
	/***** Manali 3.10.0 - 'RepSoPnd', 'RepJtPnd', 'RepJtOpnOrd', 'RepJtPndRm', 'RepPrd' from Following List to allow Multi Companyu in reports ****/
	select @row= count('x') from inserted where UoMnuCd in (
	'RepAnnexValAddn', 'RepValAddnDta', 'RepAvgOrd', 'RepBagHist', 'RepBagLst', 'RepBagPrn', 'RepFltBagPrn', 
	'RepBv','RepCustmStk', 'RepDeclnLet', 'RepDetInv', 'RepDetInvCTB', 
	'RepDetOrd', 'RepFbLst', 'RepFrLst', 'RepGrFormBack',
	'RepGrFormFrnt', 'RepGspCert', 'RepGspDetLet', 'RepGspForm', 'RepInv','RepInvCustStmt',
	'RepInvCat', 'RepInvCTBCat', 'RepInvLossSubRep', 'RepOrdCat', 
	'RepOrdProfitMrg', 'RepPckLst', 'RepCTBPckLst', /*'RepPrd',*/ 'RepRmReq', 
	'RepRmReqDyn', 'RepRmReqLoc', 'RepRmReqStat', 'RepShpBill', /*'RepSoPnd', 'RepJtPnd', 'RepJtOpnOrd', 'RepJtPndRm',*/
	'RepTxnBS', 'RepTxnBV', 'RepTxnCnv', 'RepTxnDT', 'RepTxnMLT', 
	'RepTxnMV', 'RepTxnPR', 'RepTxnREC', 'RepTxnREJ', 'RepValAddn', 'RepVchPrnBS', 
	'RepVchPrnBV', 'RepVchPrnCNV', 'RepVchPrnDT', 'RepVchPrnMLT', 'RepVchPrnMV', 
	'RepVchPrnPR', 'RepVchPrnREC', 'RepVchPrnREJ',   'RepBvRm', 
	'RepVarPrm', 'RepPckLstInv', 'RepCntrlCpy', 'RepOrdPln', 'RepFgMvm',
	'RepTreeInvst','RepTreePerform','RepFgBagLed', 
	'RepInvLoose', 'RepShpBillLoose', 'RepGrFormLoose', 'RepPckLstLoose', 'RepShpBillMulti', 
	'RepGrFormFrntMulti', 'RepGRFormBackMulti', 'RepTxnRetMem', 'RepVchPrnRetMem', 
	'RepGSPFORMMulti', 'RepGSPDetLetMulti', 'RepGSPCertMulti', 'RepDeclnLetMulti', 'RepTravExcel',
	'REPFGDET','REPFGAVG','REPFGCAT','RepFgInv','RepFgPckLst','RepFgShpBill','RepFgGrFrnt','RepFgGrBk','RepFgGsp','RepFgLclInv',
	'REPJTDET','REPJTAVG','REPJTCAT','RepJTInv','RepJTPckLst','RepJTShpBill','RepJTGrFrnt','RepJTGrBk','RepJTGsp','RepJTLclInv', 'RepShpBillFile',
	'REPMULTIPRCQT') 
		and not exists (select 'x' from Head where HCoCd = UoCoCdFr) 
	if @row> 0 
	Begin
		select @errmsg= 'Company Code Not Defined'
		goto ErrHndlr
	end
	
	/* Manali 3.3.0 - RepInvBOELst removed */
	/* ***** Manali 3.6.0 - 19/09/09 - MenuCd RepGldLs Changed to RepMetLs ***** */
	/**** Manali 3.8.0 - RepMetLsFg added ****/
	select @row= count('x') from inserted where UoMnuCd in 
		('RepCustAly', 'RepRmMst' /*, 'RepMetLsWip'*/ , 'RepMetLsFg' /*'RepGldLs'*/, 'RepBrkMsgRm', 'RepBrkMsgPrd', 
		 'RepStkSumm', /*'RepInvBOELst',*/ 'RepNetStk') and /* 'RepRmZ', */
		not exists (select 'x' from param where PTyp ='RMCTG' and UoRmCtgFr= PMCd) 
	if @row> 0 
	begin
		select @errmsg= 'Invalid Raw Material Category'
		goto ErrHndlr
	end

	/* **** Manali 3.5.0 - 05/11/08 - ErrMsg changed for Radio buttons   */
	select @row= count('x') from inserted where UoMnuCd in ('RepRmMst') and 
		not exists (select 'x' from param where PTyp ='YN' and PMCd= UOYN and PSCd= '') 
	if @row> 0 
	begin
		/* select @errmsg= 'Invalid ''Show Custm Cd'' Option ' */
		select @errmsg= 'Select if to Show Customs Codes Or Rm Codes '
		goto ErrHndlr
	end

	select @row= count('x') from inserted where UoMnuCd in ('RepRmReqDyn', 'RepRmReqStat', 'RepRmReqLoc', 'RepStkLed', 'RepFlashStk') and UoRmCtgFr<> '' and 
		not exists (select 'x' from param where PTyp ='RMCTG' and UoRmCtgFr= PMCd)
	if @row> 0 
	begin
		select @errmsg= 'Enter Blank Or Valid Raw Material Category'
		goto ErrHndlr
	end

/* **** Zubin 211 **** */ 
	select @row= count('x') from inserted where UoMnuCd in ('RepRmStkReq') and UoRmCtgFr<> 'D' and UoRmCtgFr<> 'C'
	if @row> 0 
	begin
		select @errmsg= 'Raw Material Category Can Be ''D'' Or ''C'' Only'
		goto ErrHndlr
	end 

/* **** Sachin 3.11.1.0 **** */ 
	select @row= count('x') from inserted where UoMnuCd in ('RepEPSMetImpEx') 
			and ((UoRmCtgFr<>'' and (UoRmCtgFr<> 'G' and UoRmCtgFr<> 'P' and UoRmCtgFr<> 'S' and UoRmCtgFr<> 'L')) Or 
				 (UoRmCtgTo<>'' and (UoRmCtgTo<> 'G' and UoRmCtgTo<> 'P' and UoRmCtgTo<> 'S' and UoRmCtgTo<> 'L')) Or 
				 (UoRmCtgSel<>'' And (UoRmCtgSel Like '%''D''%' Or UoRmCtgSel Like '%''C''%' Or UoRmCtgSel Like '%''X''%')))
	if @row> 0 
	begin
		select @errmsg= 'Only Metal Raw Material Categories can be Entered'
		goto ErrHndlr
	end 


	/* ****** Manali 3.5.0 - 17/11/08 - ErrMsg Changed For Check Boxes */
	select @row= count('x') from inserted where UoMnuCd in ('RepRmStkReq') and 
		Not Exists (select 'x' from Param where PTyp= 'YN' and PMCd= UOYN2 and PSCd= '') 
	if @row> 0 
	begin
		select @errmsg= 'Check To Show Order Box Detail?'
		/* select @errmsg= 'Show Order Box Detail?   (Enter ''Y'' / ''N'')' */
		goto ErrHndlr
	end 

	/* ***** Manali Trading Module ***** */
	Select @row= count('x') from inserted Where UoMnuCd In ('REPJTDET','REPJTAVG') and 
		not exists(Select 'x' from Param Where PTyp='TCTYP' and PMCd =UoTcTypFr and PMCd='JT')
	if @row> 0 
	begin
		select @errmsg= 'TcTyp Should Be ''JT'''
		goto ErrHndlr
	end

	Select @row= count('x') from inserted Where UoMnuCd In ('REPJTDET','REPJTAVG') and 
		not exists(Select 'x' from Param Where PTyp='TC' and PMCd In (UoOmTcFr, UoOmTcTo) and PValue=UoTcTypFr)
	if @row> 0 
	begin
		select @errmsg= 'Tc Should Be of TcTyp ''JT'''
		goto ErrHndlr
	end
	/* ***** Manali Trading Module ***** */
	
	/*** Jay 3.4(FG)***/
	/* ***** Manali Trading Module ***** */
	select @row= count('x') from inserted where UoMnuCd in ('RepDetInv', 'RepDetInvCTB', 
			'RepAvgInv', 'RepAvgInvCTB', 'RepDetOrd', 'RepAvgOrd', 'REPFGDET','REPFGAVG', 'REPJTDET','REPJTAVG', 'REPMULTIPRCQT') and 
		not exists (select 'x' from param where PTyp ='YN' and UOYN2= PMCd) 
	if @row> 0 
	begin
		/* select @errmsg= 'Show All Value? (Enter ''Y'' or ''N'')' */
		select @errmsg= 'Select If To Show All Value? '
		goto ErrHndlr
	end

  /***** Manali 3.9.0 - LME Selection *****/
	select @row= count('x') from inserted where UoMnuCd in ('RepDetInv', 'RepDetInvCTB', 
			'RepAvgInv', 'RepAvgInvCTB', 'RepDetOrd', 'RepAvgOrd', 'REPFGDET','REPFGAVG', 'REPJTDET','REPJTAVG') and 
			(UoLme1 Not In ('LMG', 'LMP', 'LMS', 'LML', '') or UoLme2 Not In ('LMG', 'LMP', 'LMS', 'LML', '')) 
		  
	if @row> 0 
	begin
		select @errmsg= 'Invalid Lme Selected'
		goto ErrHndlr
	end

	select @row= count('x') from inserted where UoMnuCd in ('RepDetInv', 'RepDetInvCTB', 
			'RepDetOrd', 'REPFGDET','REPJTDET') and UoYn23='Y'
			 and (UoLme1 <>  '' or UoLme2 <> '') 
	if @row> 0 
	begin
		select @errmsg= 'Lme Cannot Be Selected when Excel Format is Selected'
		goto ErrHndlr
	end

	select @row= count('x') from inserted where UoMnuCd in ('RepRmStkReq') and 
		Not Exists (select 'x' from Param where PTyp= 'YN' and PMCd= UOYN4 and PSCd= '') 
	if @row> 0 
	begin
		/* select @errmsg= 'Show Stock Box Detail?   (Enter ''Y'' / ''N'')' */
		select @errmsg= 'Show Stock Box Detail?  '
		goto ErrHndlr
	end 

	/* **** Zubin 211 **** */ 
	/*** Jay 3.4(FG) ***/
	/**** Manali Trading Module *****/
	select @row= count('x') from inserted where UoMnuCd in ('RepStkLed', 'RepStkSumm','RepFgStkLed', 'RepJtStkLed') 
		and UoOpnYYMm  > convert(int,convert(varchar(6),UoDtFr,112))

	if @row> 0 
	begin
		select @errmsg= 'Open Year/Month Should Be <= From Voucher Date'
		goto ErrHndlr
	end

	/*** Jay 3.4(FG) ***/
	/**** Manali Trading Module *****/
	select @row= count('x') from inserted where UoMnuCd in ('RepStkLed','RepFgStkLed', 'RepJtStkLed') and UoDtTo< UoDtFr 
	if @row> 0 
	begin
		select @errmsg= 'To Date Should Be Greater Than From Date'
		goto ErrHndlr
	end

	select @row= count('x') from inserted where UoMnuCd in ('RepFlashStk') and 
		UoOpnYyMM > UoYyMmFr
	if @row> 0 
	begin
		select @errmsg= 'Open Year/Month Should Be Less Than From Year/Month'
		goto ErrHndlr
	end

	/* **** Zubin 211 (RepRmStkReq added) **** */ 
	select @row= count('x') from inserted where UoMnuCd in ('RepFlashStk', 'RepRmStkReq') and UoYyMmTo< UoYyMmFr 
	if @row> 0 
	begin
		select @errmsg= 'To YyyyMm Should Be Greater Than From YyyyMm'
		goto ErrHndlr
	end

	select @row= count('x') from inserted where UoMnuCd in ('RepDsg', 'RepFbLst', 'RepFwbLst', 'RepWIWPrd', 'RepWIWExp', 'RepWIWLongPrd', 'RepWIWLongExp') and 
		not exists (select 'x' from param where PTyp ='YN' and UoYN= PMCd) 
	if @row> 0 
	begin
		select @errmsg= 'Show The Second Detail Line Of Report Yes/No?   (Enter ''Y'' / ''N'')'
		goto ErrHndlr
	end

	/* ***** Manali 3.5.0 - 10/11/08 - Show Special Remark in wiw ***** */
	/* ***** Manali Trading Module ***** */
	select @row= count('x') from inserted where UoMnuCd in ('RepWIWPrd', 'RepWIWExp', 'RepWIWLongPrd', 'RepWIWLongExp', 'RepSoPnd', 'RepJtPnd', 'RepJtOpnOrd') and 
		not exists (select 'x' from param where PTyp ='YN' and UoYN4= PMCd) 
	if @row> 0 
	begin
		select @errmsg= 'Check if To Show Special Remark ? '
		goto ErrHndlr
	end

	/* ***** Manali 3.5.0 - 12/11/08 - Apply Colours in wiw ***** */
	select @row= count('x') from inserted where UoMnuCd in ('RepWIWPrd', 'RepWIWExp', 'RepWIWLongPrd', 'RepWIWLongExp') and 
		not exists (select 'x' from param where PTyp ='YN' and UoYN5= PMCd) 
	if @row> 0 
	begin
		select @errmsg= 'Check if To Apply Colors ?'
		goto ErrHndlr
	end

	/*** Jay 3.4(FG) ***/ 
	/* ***** Manali Trading Module ***** */
	/***** Manali 3.8.0 - RepPckLstInvBag added  ****/
	select @row= count('x') from inserted where UoMnuCd in ('RepDsgCrd', 
		'RepDetOrd', 'RepAvgOrd', 'RepInvCustStmt', 'RepDetInv', 'RepDetInvCTB', 'RepAvgInv', 
		'RepAvgInvCTB', 'RepInv','REPFGDET','REPFGAVG','RepFgInv','REPJTDET','REPJTAVG','RepJtInv', 'RepPckLstInvBag',
		'REPMULTIPRCQT') and 
		not exists (select 'x' from param where PTyp ='YN' and UoYN= PMCd) 
	if @row> 0 
	begin
		/* select @errmsg= 'Show Picture Yes/No?    (Enter ''Y'' / ''N'')' */
		select @errmsg= 'Show Picture ?'
		goto ErrHndlr
	end

	/* **** Manali 3.5.0 - 06/11/08 - ErrMsg changed for Radio Buttons  */
	select @row= count('x') from inserted where UoMnuCd in ('RepBagLst', 'RepWipBPcs', 'RepWipFgBPcs') and UoYN<> '' and 
		not exists (select 'x' from param where PTyp ='YN' and UoYN= PMCd) 
	if @row> 0 
	begin
		/* select @errmsg= 'Show Open/ Closed/ All Bags?   (Enter ''N'' / ''Y'' / Blank)' */
		select @errmsg= 'Select if to Show Open/ Closed/ All Bags'
		goto ErrHndlr
	end

	/* ****** Sachin 2.14.0 [23-11-2006] - [Multi-Company] - Changed Company to Self CoCd  ****** */
	select @row= count('x') from inserted where UoMnuCd in ('RepWIWPrd', 'RepWIWExp') and UoYn2='N'
		And not exists (select 'x' from vparam where vPCoCd = @wSelfCoCd 
				and vPTyp ='WIWSEQ' and UoSeqCdFr= vPMCd) 
	if @row> 0 
	begin
		select @errmsg= 'Enter Valid WIW Sequence Code'
		goto ErrHndlr
	end

	select @row= count('x') from inserted where UoMnuCd in ('RepWIWLongPrd', 'RepWIWLongExp') and UoYn2='N'
		And not exists (select 'x' from vparam where vPCoCd = @wSelfCoCd and 
				vPTyp ='WIWLONG' and UoSeqCdFr= vPMCd) 
	if @row> 0 
	begin
		select @errmsg= 'Enter Valid WIW Long Sequence Code'
		goto ErrHndlr
	end

	select @row= count('x') from inserted where UoMnuCd in ('RepWIWPrd', 'RepWIWExp') and UoYn2='Y'
		And not exists (select 'x' from vparam where vPCoCd = @wSelfCoCd 
				and vPTyp ='WIWSEQ' and UoPrntSeqCdFr= vPMCd) 
	if @row> 0 
	begin
		select @errmsg= 'Enter Valid WIW Parent Sequence Code'
		goto ErrHndlr
	end

	select @row= count('x') from inserted where UoMnuCd in ('RepWIWLongPrd', 'RepWIWLongExp') and UoYn2='Y'
		And not exists (select 'x' from vparam where vPCoCd = @wSelfCoCd and 
				vPTyp ='WIWLONG' and UoPrntSeqCdFr= vPMCd) 
	if @row> 0 
	begin
		select @errmsg= 'Enter Valid WIW Long Parent Sequence Code'
		goto ErrHndlr
	end

	/* ****** Sachin 2.14.0 [23-11-2006] - [Multi-Company] - Changed Company to Self CoCd  ****** */
	select @row= count('x') from inserted where UoMnuCd in ('RepWIWPrd', 'RepWIWExp', 
			'RepWIWLongPrd', 'RepWIWLongExp') and 
		not exists (select 'x' from param where PTyp ='YN' and UOYN2= PMCd) 
	if @row> 0 
	begin
		/* select @errmsg= 'Show Parent Sequence (Y/N)? (Enter ''Y'' or ''N'')' */
		select @errmsg= 'Show Parent Sequence ?'
		goto ErrHndlr
	end
	/* ****** Sachin 2.14.0 [23-11-2006] - [Multi-Company] - Changed Company to Self CoCd  ****** */
	
       /****Geeta***Emr208****/
	select @row= count('x') from inserted where UoMnuCd in ('RepWIWPrd', 'RepWIWExp','RepWIWLongPrd', 'RepWIWLongExp',
								'RepWipBPcs', 'RepWipFgBPcs', 'RepWipBPcsBkDt', 'RepWipFgBPcsBkDt') and 
		not exists (select 'x' from Param where PTyp ='YN' and UoYN1= PMCd) 
	if @row> 0 
	begin
		/* select @errmsg= 'Show Picture Yes/No?    (Enter ''Y'' / ''N'')' */
		select @errmsg= 'Show Picture ? '
		goto ErrHndlr
	end
/* *** JAY 2.13(CT) *** */
/* **** Zubin 211 **** */ 
	/****** added RepWipBRmBkDt (Jen 2.14) ******/
	/* ***** Manali 3.6.0 - 19/09/09 - MenuCd RepGldLs Changed to RepMetLs ***** */
	select @row= count('x') from inserted where 
		(UoMnuCd In ('RepWipBPcs', 'RepWipBPcsBkDt', 'RepPrd', 'RepBv', 'RepBagHist', 'RepRejAna', 
			'RepBrkMsgRm', 'RepBrkMsgPrd', /* 'RepTxnDT', 'RepTxnBS', 'RepTxnMLT', 'RepTxnBV', 
			'RepTxnREJ', 'RepTxnBM', */ 'RepBagLst', 'REPBAGRECO','RepWipBRm', 'RepWipBRmBkDt', 'RepBvRm', 'RepMetLsWip' /*'RepGldLs'*/)
			and UoBagTyp<> '' and 
			not exists (select 'X' from param where PTyp= 'BTYP' and UoBagTyp= PMCd and PSCd= '')) 
	if @row> 0 
	Begin
		select @errmsg= 'Invalid Bag Type Option. (''P'' For Primary, ''F'' For Flute, ''C'' For Component Bags, '''' For All)' 
		goto ErrHndlr
	end

/* *** BEF 2.13(CT) *** */
/*	select @row= count('x') from inserted where 
		(UoMnuCd In ('RepWipBPcs', 'RepWipBPcsBkDt', 'RepPrd', 'RepBv', 'RepBagHist', 'RepRejAna', 
			'RepBrkMsgRm', 'RepBrkMsgPrd', */ 
			/* 'RepTxnDT', 'RepTxnBS', 'RepTxnMLT', 'RepTxnBV', 'RepTxnREJ', 'RepTxnBM', */ 
			/*'RepBagLst', 'REPBAGRECO') and UoYN2<> '' and 
			not exists (select 'X' from param where PTyp= 'YN' and UoYN2= PMCd and PSCd= '')) 
		OR 
		(UoMnuCd In ('RepWipBRm', 'RepBvRm', 'RepGldLs') and UoYN3<> '' and 
			not exists (select 'x' from param where PTyp= 'YN' and UoYN3= PMCd and PSCd= '') )
	if @row> 0 
	Begin
		select @errmsg= 'Invalid ''Show Flute/Primary Bags'' Option. (''Y'' For Flute, ''N'' For Primary Bags, '''' For Both)' 
		goto ErrHndlr
	end
*/

	select @row= count('x') from inserted where UoMnuCd in ('RepWIWPrd','RepWIWExp','RepWIWLongExp','RepWIWLongPrd') 
		and not exists (select 'x' from param where PTyp= 'BTYP' and UoBagTyp= PMCd and PSCd= '') 
	if @row> 0 
	Begin
		select @errmsg= 'Show Flute/ Component Or Primary Bags? (''F'' For Flute, ''C'' For Component, ''P'' For Primary Bags)'
		goto ErrHndlr
	end 

	/****** (added RepWipBRmBkDt) (Jen 2.14) ******/
	/* ***** Manali 3.6.0 - 19/09/09 - MenuCd RepGldLs Changed to RepMetLs ***** */
	select @row= count('x') from inserted where UoBChrFr<> '' and 
		UoMnuCd In ('RepWIWPrd', 'RepWIWExp', 'RepWIWLongExp', 'RepWIWLongPrd', 'RepWipBPcs', 
				'RepWipBPcsBkDt', 'RepPrd', 'RepBrkMsgRm', 'RepBrkMsgPrd', 'REPBAGRECO', 'RepWipBRm', 
				'RepWipBRmBkDt', 'RepBvRm', 'RepMetLsWip' /*'RepGldLs'*/) and UoBagTyp= '' 
	if @row> 0 
	Begin
		select @errmsg= 'Bag Character Should Not Be Entered' 
		goto ErrHndlr
	end

/* *** BEF 2.13(CT) *** */
/*	select @row= count('x') from inserted where UoBChrFr<> '' and 
		(	(	UoMnuCd In ('RepWIWPrd', 'RepWIWExp', 'RepWIWLongExp', 'RepWIWLongPrd', 'RepWipBPcs', 
				'RepWipBPcsBkDt', 'RepPrd', 'RepBrkMsgRm', 'RepBrkMsgPrd', 'REPBAGRECO') and 
				UoYN2 = '' 
			)
			Or
			(	UoMnuCd In ('RepWipBRm', 'RepBvRm', 'RepGldLs') and UoYN3 = '' 
			)
		)
	if @row> 0 
	Begin
		select @errmsg= 'Bag Character Should Not Be Entered' 
		goto ErrHndlr
	end
*/
	
	/****** added RepWipBRmBkDt (Jen 2.14) ******/
	/* ***** Manali 3.6.0 - 19/09/09 - MenuCd RepGldLs Changed to RepMetLs ***** */
	select @row= count('x') from inserted where UoBChrFr<> '' and 
	UoMnuCd In ('RepWIWPrd', 'RepWIWExp', 'RepWIWLongExp', 'RepWIWLongPrd', 'RepWipBPcs', 
	'RepWipBPcsBkDt', 'RepPrd', 'RepBrkMsgRm', 'RepBrkMsgPrd', 'REPBAGRECO','RepWipBRm', 
	'RepWipBRmBkDt', 'RepBvRm', 'RepMetLsWip' /*'RepGldLs'*/) and 
	Not Exists (select 'x' from param where PTyp= 'BCHR' and UoBChrFr= PMCd and PSCd= '' and PValue= UoBagTyp) 
	if @row> 0 
	Begin
		select @errmsg= 'Invalid Bag Character' 
		goto ErrHndlr
	end

/* *** BEF 2.13(CT) *** */
/*	select @row= count('x') from inserted where UoBChrFr<> '' and 
			(	(	UoMnuCd In ('RepWIWPrd', 'RepWIWExp', 'RepWIWLongExp', 'RepWIWLongPrd', 'RepWipBPcs', 
					'RepWipBPcsBkDt', 'RepPrd', 'RepBrkMsgRm', 'RepBrkMsgPrd', 'REPBAGRECO') and 
					Not Exists (select 'x' from param where PTyp= 'BCHR' and UoBChrFr= PMCd and PSCd= '' and PValue= UoYN2) 
				)
				Or
				(	UoMnuCd In ('RepWipBRm', 'RepBvRm', 'RepGldLs') and 
					Not Exists (select 'x' from param where PTyp= 'BCHR' and UoBChrFr= PMCd and PSCd= '' and PValue= UoYN3) 
				)
			)
*/			/* 'RepBagHist', 'RepBagLst', 'RepRejAna', 'RepTxnDT', 'RepTxnBS', 'RepTxnMLT', 'RepTxnBV', 'RepTxnREJ', 'RepTxnBM', */ 
/*	if @row> 0 
	Begin
		select @errmsg= 'Invalid Bag Character' 
		goto ErrHndlr
	end
*/

	select @row= count('x') from inserted where UoBagTyp in ('C','F') and UOBCHRFR= '' and 
			UoMnuCd In ('RepWIWPrd', 'RepWIWExp', 'RepWIWLongExp', 'RepWIWLongPrd') 
	if @row> 0 
	Begin
		select @errmsg= 'Bag Character Cannot Be Blank For Flute/ Component Bags' 
		goto ErrHndlr
	end

	select @row= count('x') from inserted where UoBagTyp= 'P' and UOBCHRFR<> '' and 
			UoMnuCd In ('RepWIWPrd', 'RepWIWExp', 'RepWIWLongExp', 'RepWIWLongPrd') 
	if @row> 0 
	Begin
		select @errmsg= 'Bag Character Should Be Blank For Primary Bags' 
		goto ErrHndlr
	end

	/*** (Bef 2.13 CT)
	select @row= count('x') from inserted where 
		(UoMnuCd In ('RepWipBPcs') and (UoYN= '' or UoYN= 'N') and UoBagTyp IN ('','C','F') )
		OR (UoMnuCd In ('RepWipBRm', 'RepBvRm') and (UoYN1= '' or UoYN1= 'N') and UoBagTyp IN ('','C','F') )
		OR (UoMnuCd In ('RepBagLst') and (UoYN= '' or UoYN= 'Y') and UoBagTyp IN ('','C','F') )
			 -- In RepBagLst UoYN= 'Y' means BCls= 'Y'
	(Bef 2.13 CT) ***/
	

	/*** (Jen 2.13 CT) ***/
	/* **** Manali 3.5.0 - 06/11/08 - ErrMsg changed for Radio Buttons ***** */
	/* **** Manali 3.6.0 - 01/11/09 - Following check commented as Flute/Comp Bags can be closed ***** */
/*	select @row= count('x') from inserted where 
		(UoMnuCd In ('RepWipBPcs') and (UoYN= '' or UoYN= 'N') and UoBagTyp IN ('','C','F') )
		OR (UoMnuCd In ('RepWipBRm', 'RepBvRm') and (UoYN1= '' or UoYN1= 'N') and UoBagTyp IN ('','C','F') )
		OR (UoMnuCd In ('RepBagLst') and (UoYN= '' or UoYN= 'N') and UoBagTyp IN ('','C','F') )
	
	/*** (Jen 2.13 CT) ***/
	if @row> 0 
	Begin 
		/* select @errmsg= ' ''Open Bags'' Option Should Be ''Y'' For ''Flute Bags'' / ''Component Bags'''  */
		select @errmsg= ' ''Open Bags'' Option Should Be Selected For ''Flute Bags'' / ''Component Bags''' 
		goto ErrHndlr 
	end 
*/

/* **** Zubin 211 **** */ 
	
	/****** added RepWipBRmBkDt (Jen 2.14) ******/
	/* ******** Manali 3.5.0 - 11/11/08 - RepSOPnd, RepSalAnaDsg added */ 
	/* ***** Manali Trading Module ***** */
	select @row= count('x') from inserted where UoMnuCd in ('RepWipBRm','RepWipFgBRm', 'RepWipBRmBkDt', 'RepSOPnd', 'RepJtPnd', 'RepJtOpnOrd', 'RepSalAnaDsg') and 
		not exists (select 'x' from Param where PTyp ='YN' and UoYN2= PMCd) 
		/* , 'RepBvRm' */
	if @row> 0 
	begin
		/* select @errmsg= 'Show Picture Yes/No?    (Enter ''Y'' / ''N'')' */
		select @errmsg= 'Show Picture ?'
		goto ErrHndlr
	end
       /****Geeta***Emr208****/
	
	/*******	Manoj 2.10.0 (NOT REQUIRED)	*******/
/*
	select @row= count('x') from inserted where UoMnuCd in ('RepWipFgBPcs') and UoYN2 <>'' and 
		not exists (select 'x' from Param where PTyp ='YN' and UoYN2= PMCd) and
		(Select HCtbMod From Head Where HCd = 'ZSELF' ANd HCoCd = UOCoCdFr) = 'Y'
	if @row> 0 
	begin
		select @errmsg= 'Show CTB Open/ Closed/ All Bags?   (Enter ''N'' / ''Y'' / Blank)'
		goto ErrHndlr
	end

	select @row= count('x') from inserted where UoMnuCd in ('RepWipFgBPcs') and UoYN2 <>'' and 
		(Select HCtbMod From Head Where HCd = 'ZSELF' ANd HCoCd = UOCoCdFr) = 'N'
	if @row> 0 
	begin
		select @errmsg= 'Show CTB Open/ Closed/ All Bags? Should be Blank.'
		goto ErrHndlr
	end
*/
	/*******	Manoj 2.10.0	*******/
	/****** added RepWipBRmBkDt (Jen 2.14) ******/
	/* *** Jay 2.14Next *** [removed 'RepBvRm'] */
	select @row= count('x') from inserted where UoMnuCd in ('RepWipBPcs', 'RepWipBRm', 'RepWipBRmBkDt', 
		'RepWipBPcsBkDt' ) and 
		not exists (select 'x' from param where PTyp ='BAGPCS' and UoBagPcs= PMCd) 
	if @row> 0 
	begin
		/* select @errmsg= 'Show Normal/ Rejected Pieces?   (Enter ''N'' / ''R'')' */
		select @errmsg= 'Select If To Show Normal/ Rejected Pieces?  '
		goto ErrHndlr
	end
	/* *** Jay 2.14Next *** [removed 'RepBvRm'] */

	/***** Manali 3.7.0 - Added Order Det and avg Reports added ****/
	select @row= count('x') from inserted where UoMnuCd in ('RepBagPrn', 'RepDetOrd', 'RepAvgOrd', 'RepFgDet', 'RepFgAvg', 'RepJtDet', 'RepJtAvg') and 
		not exists (select 'x' from param where PTyp ='PCT' and UoPMCdFr= PMCd) 
	if @row> 0 
	begin
		select @errmsg= 'Show 3D Picture Or Line Drawing? (Enter Valid Option)'
		goto ErrHndlr
	end
	
	/*** Jay 3.4(FG) ***/ 
	/* ***** Manali Trading Module ***** */	
	select @row= count('x') from inserted where UoMnuCd in ('RepDetOrd', 'RepDetInv', 'RepDetInvCTB','REPFGDET', 'REPJTDET',
	'REPMULTIPRCQT') and 
		not exists (select 'x' from param where PTyp= 'RMGRP' and PMCd= 'OrdDet' and PSCD= UoRmGrp)
	if @row> 0 
	begin
		select @errmsg= 'Invalid Raw Material Group Option'
		goto ErrHndlr
	end

/*	select @row=count('x') from inserted where UoMnuCd= 'RepAvgOrd' and 
		not exists (select 'x' from param where PTyp= 'RMGRP' and PMCd= 'OrdAvg' and PSCD= UoRmGrp)
	if @row> 0 
	begin
		select @errmsg= 'Invalid Raw Material Group Option'
		goto ErrHndlr
	end
*/
	
	/*** Jay 3.4(FG) ***/ 
	/* ***** Manali Trading Module ***** */
	select @row= count('x') from inserted where UoMnuCd in ('RepDetOrd', 'RepDetInv', 'RepDetInvCTB','REPFGDET', 'REPJTDET') and 
		not exists (select 'x' from param where PTyp= 'LABGRP' and PMCd= 'OrdDet' and PSCD= UoLabGrp)
	if @row> 0 
	begin
		select @errmsg= 'Invalid Labour Group Option'
		goto ErrHndlr
	end

/*	select @row=count('x') from inserted where UoMnuCd= 'RepAvgOrd' and 
		not exists (select 'x' from param where PTyp= 'LABGRP' and PMCd= 'OrdAvg' and PSCD= UoLabGrp)
	if @row>0 
	Begin		select @errmsg= 'Invalid Labour Group Option'
		goto ErrHndlr
	end
*/

	select @row= count('x') from inserted where UoMnuCd in ('RepDetOrd', 'RepAvgOrd', 'RepOrdCat', 'RepOrdProfitMrg') and 
		not exists (select 'x' from param where PTyp= 'TCTYP' and PMCd= UoTcTypFr and PValue1= 'OM')
	if @row> 0 
	begin
		select @errmsg= 'Invalid Tc Type'
		goto ErrHndlr
	end

	/*** Jay 3.4(FG) ***/ 
	select @row= count('x') from inserted where UoMnuCd in ('REPFGDET','REPFGAVG','REPFGCAT') and 
		not exists (select 'x' from param where PTyp= 'TCTYP' and PMCd= UoTcTypFr and PValue1= 'FG')
	if @row> 0 
	begin
		select @errmsg= 'Invalid Tc Type'
		goto ErrHndlr
	end

	/* ***** Manali Trading Module ***** */
	select @row= count('x') from inserted where UoMnuCd in ('REPJTDET','REPJTAVG','REPJTCAT') and 
		not exists (select 'x' from param where PTyp= 'TCTYP' and PMCd= UoTcTypFr and PValue1= 'JT')
	if @row> 0 
	begin
		select @errmsg= 'Invalid Tc Type'
		goto ErrHndlr
	end

	/************************* Geeta 2.06 **************************/
        select @row= count('x') from inserted where UoMnuCd In ('RepInvCuststmt', 'REPMULTIPRCQT') and
		(UoLMCd1='SET'  or UoLMCd2='SET'  or UoLMCd3='SET'  )
	if @row> 0
	begin
		select @errmsg= 'Setting is not a Valid Option '
		goto ErrHndlr
	end
	select @row= count('x') from inserted where UoMnuCd In ('RepInvCustStmt', 'REPMULTIPRCQT') and
		( (UoLMCd1<>'' and ((UoLMcd1=UoLMCd2)  or (UoLMCd1=UoLMCd3))) or
		 (UoLMCd2<>'' and (UoLMCd2=UoLMCd3)))
	if @row> 0
	begin
		select @errmsg= 'Two or More Labour Codes Cannot Be The Same'
		goto ErrHndlr
	end

	select @row= count('x') from inserted where UoMnuCd In ('RepInvCustStmt', 'REPMULTIPRCQT') and UoLMCd1<>'' and
		not exists (select 'x' from param where PTyp= 'LabMCD' and PMCd= UoLMCd1 and PSCD= '')
	if @row> 0
	begin
		select @errmsg= 'Invalid Labour Code '
		goto ErrHndlr
	end
        select @row= count('x') from inserted where UoMnuCd In ('RepInvCustStmt', 'REPMULTIPRCQT') and UoLMCd2<>'' and
		not exists (select 'x' from param where PTyp= 'LabMCD' and PMCd= UoLMCd2 and PSCD= '')
	if @row> 0
	begin
		select @errmsg= 'Invalid Labour Code '
		goto ErrHndlr
	end
        select @row= count('x') from inserted where UoMnuCd In ('RepInvCustStmt', 'REPMULTIPRCQT') and UoLMCd3<>'' and
		not exists (select 'x' from param where PTyp= 'LabMCD' and PMCd= UoLMCd3 and PSCD= '')
	if @row> 0
	begin
		select @errmsg= 'Invalid Labour Code '
		goto ErrHndlr
	end
	select @row= count('x') from inserted where UoMnuCd in ('RepInvCustStmt') and
		not exists (select 'x' from param where PTyp ='YN' and UoYN5= PMCd)
	if @row> 0
	begin
		/* select @errmsg= 'Show Diamond And ColStn Value Per Ct (Enter ''Y'' / ''N'')' */
		select @errmsg= 'Show Diamond And ColStn Value Per Ct ?'
		goto ErrHndlr
	end
        select @row= count('x') from inserted where UoMnuCd in ('RepInvCustStmt') and
		not exists (select 'x' from param where PTyp ='YN' and UoYN6= PMCd)
	if @row> 0
	begin
		/* select @errmsg= 'Show PO Number (Enter ''Y'' / ''N'')' */
		select @errmsg= 'Show PO Number ?'
		goto ErrHndlr
	end
/* **** Zubin 211 Not Required **** */ 
/*	select @row= count('x') from inserted where UoMnuCd in ('RepInvCustStmt') and
		not exists (select 'x' from param where PTyp ='YN' and UoYN3= PMCd)
	if @row> 0
	begin
		select @errmsg= 'Show Running Serial (Enter ''Y'' / ''N'')'
		goto ErrHndlr
	end
*/
/* **** Zubin 211 Not Required **** */ 

	select @row= count('x') from inserted where UoMnuCd in
		('RepInvCustStmt') and
		not exists (select 'x' from param where PTyp ='YN' and UOYN7= PMCd)
	if @row> 0
	begin
		/* select @errmsg= 'Show Customer Design Code? (Enter ''Y'' or ''N'')' */
		select @errmsg= 'Show Customer Design Code? '
		goto ErrHndlr
	end

	/*** Jay 3.4(FG) ***/ 
	/* ***** Manali Trading Module ***** */
	select @row= count('x') from inserted where UoMnuCd in ('RepInvCustStmt', 'RepDetOrd','REPFGDET','REPJTDET',
	  'REPMULTIPRCQT') and
	    	not exists (select 'x' from param where PTyp ='YN' and UOYN8= PMCd)
	if @row> 0
	begin
  		/* select @errmsg= 'Show Customer RM Code? (Enter ''Y'' or ''N'')' */
		select @errmsg= 'Show Customer RM Code? '
		goto ErrHndlr
	end
	select @row= count('x') from inserted where UoMnuCd in ('RepInvCustStmt') and
		not exists (select 'x' from param where PTyp ='INVAL' and UoInVal= PMCd )
	if @row> 0
	begin
		select @errmsg= 'Show Actual, Invoice Or Customs Values?   (Enter ''A'' / ''I''/''Z''/''S'')'
		goto ErrHndlr
	end

	select @row= count('x') from inserted where UoMnuCd in ('RepInvCuststmt') and
		not exists (select 'x' from param where PTyp= 'RMGRP' and PMCd= 'OrdDet' and PSCD= UoRmGrp)
	if @row> 0
	begin
		select @errmsg= 'Invalid Raw Material Group Option'
		goto ErrHndlr
	end

	--	Manoj 2.10.0
	/*** Jay 3.4(FG) ***/ 
	/* ***** Manali Trading module ***** */
	select @row= count('x') from inserted where UoMnuCd in ('RepDetOrd', 'RepDetInv', 'RepDetInvCTB','RepAvgOrd', 'RepAvgInv', 'RepAvgInvCTB', 
			'REPFGDET','REPFGAVG','REPJTDET','REPJTAVG') and 
		not exists (select 'x' from param where PTyp ='YN' and UoYN9= PMCd) 
	if @row> 0 
	Begin 
		/* select @errmsg= 'Show Design & RM Summary (Yes / No)?   (Enter ''Y'' / ''N'')' */
		select @errmsg= 'Show Design & RM Summary ?'
		goto ErrHndlr
	end
	
	/***** Manali 3.10.0 - 14/03/12 - Show GldAs Wt or Net Wt ****/
	select @row= count('x') from inserted where UoMnuCd in ('RepDetOrd', 'RepDetInv', 'RepDetInvCTB','RepAvgOrd', 'RepAvgInv', 'RepAvgInvCTB', 
			'REPFGDET','REPFGAVG','REPJTDET','REPJTAVG')
		and not exists (select 'x' from param where PTyp ='YN' and UoYN28= PMCd) 
	if @row> 0 
	Begin
		select @errmsg= 'Select if to Show GldAs Wt or Net Wt ?'
		goto ErrHndlr
	end


	
	/* ***** Manali 3.5.0 - 07/11/08 - ErrMsg changed for Radio Buttons */
	select @row= count('x') from inserted where UoMnuCd in ('RepDetOrd') and 
		not exists (select 'x' from param where PTyp ='YN' and UoYN10= PMCd) 
	if @row> 0 
	begin
		/* select @errmsg= 'Show Calculated Pointer (Yes / No)?   (Enter ''Y'' / ''N'')' */
		select @errmsg= 'Select if to Show Calculated Pointer Or Actual Stored Pointer'
		goto ErrHndlr
	end

/* Sachin 211 */ 
	select @row= count('x') from inserted where UoMnuCd in ('RepBagReco') and 
		not exists (select 'x' from param where PTyp ='YN' and UOYN= PMCd and PSCd= '') 
	if @row> 0 
	begin
		/* select @errmsg= 'Show All Or UnMatched Records (Yes / No)?   (Enter ''Y'' / ''N'')' */
		select @errmsg= 'Select If To Show All Or UnMatched Records '
		goto ErrHndlr
	end
/* Sachin 211 */ 

/* **** Zubin 211 'RepRmStkReq' added **** */ 
	/* ***** Manali3.5.0 - 05/11/08 - ErrMsg changed for Radio buttons */
	select @row= count('x') from inserted where UoMnuCd in ('RepRmReqStat', 'RepRmReqDyn', 'RepRmReqLoc', 
		'RepRmStkReq') and not exists (select 'x' from param where PTyp ='YN' and UoYN1= PMCd) 
	if @row> 0 
	begin	
		/* select @errmsg= 'Show Calculated Pointer (Yes / No)?   (Enter ''Y'' / ''N'')' */
		select @errmsg= 'Select if to display Calculated Pointer Or Actual Pointer'
		goto ErrHndlr
	end
	
	/* ***** Manali3.5.0 - 05/11/08 - ErrMsg changed for Radio buttons */
	select @row= count('x') from inserted where UoMnuCd in ('RepDsgCrd') and 
		not exists (select 'x' from param where PTyp ='YN' and UoYN4= PMCd) 
	if @row> 0 
	begin
		/* select @errmsg= 'Show Calculated Pointer (Yes / No)?   (Enter ''Y'' / ''N'')' */
		select @errmsg= 'Select if to display Calculated Pointer Or Actual Pointer'
		goto ErrHndlr
	end

	/* **** Manali 3.6.0 -  Pointer Option Can be Selected Only form Order Rm Details ***** */
	select @row= count('x') from inserted where UoMnuCd in ('RepBagPrn') 
		and not exists (select 'x' from param where PTyp ='YN' and UoYN= PMCd) 
	if @row> 0 
	begin
		select @errmsg= 'Select Actual Wt or Order Wt '
		goto ErrHndlr
	end

	/* ***** Manali3.5.0 - 05/11/08 - ErrMsg changed for Radio buttons */
	/* **** Manali 3.6.0 -  Pointer Option Can be Selected Only form Order Rm Details ***** */
	select @row= count('x') from inserted where UoMnuCd in ('RepBagPrn') and UoYn<>'Y'
		and not exists (select 'x' from param where PTyp ='YN' and UoYN1= PMCd) 
	if @row> 0 
	begin
		/* select @errmsg= 'Show Production / Sales Pointer (Yes / No)?   (Enter ''Y'' / ''N'')'*/
		select @errmsg= 'Select Production / Sales Pointer '
		goto ErrHndlr
	end

	/* **** Manali 3.6.0 -  Pointer Option Can be Selected Only form Order Rm Details ***** */
	select @row= count('x') from inserted where UoMnuCd in ('RepBagPrn') and UoYn='Y'
		and UoYn1<>''
	if @row> 0 
	begin
		select @errmsg= 'Pointer Option Should not be Selected When Actual Rm Details are to be Displayed '
		goto ErrHndlr
	end

/* Sachin 211 */ 
	select @row= count('x') from inserted where UoMnuCd in ('RepBagPrn') and 
		not exists (select 'x' from param where PTyp ='BAGFMT' and UoPMCd= PMCd) 
	if @row> 0 
	begin
		select @errmsg= 'Bag Format Invalid'
		goto ErrHndlr
	end
/* Sachin 211 */ 

	select @row= count('x') from inserted where UoMnuCd in ('RepBagPrn') and 
		not exists (select 'x' from param where PTyp ='YN' and UoYn2= PMCd) 
	if @row> 0 
	begin
		select @errmsg= 'Show Single Bag Per Page?'
		goto ErrHndlr
	end

	select @row= count('x') from inserted where UoMnuCd in ('RepBagPrn') and 
		UoPMCd='BAG6' AND UoYn2='Y'
	if @row> 0 
	begin
		select @errmsg= 'Single Bag Per Page Option Cannot Be Selected For Bag Format- Bag6'
		goto ErrHndlr
	end

	/* **** Zubin 211 **** */ 
	/*** Jay 3.4(FG) ***/ 
	/* ***** Manali Trading module ***** */
	select @row= count('x') from inserted where UoMnuCd in ('RepDetOrd', 'RepAvgOrd', 'RepOrdCat',
	  'REPFGDET','REPFGAVG','REPFGCAT', 'REPJTDET','REPJTAVG','REPJTCAT', 'REPMULTIPRCQT') and 
		Not Exists (select 'x' from param where PTyp ='YN' and UoYN11= PMCd) 
/*	select @row= count('x') from inserted where UoMnuCd in ('RepDetOrd', 'RepAvgOrd', 'RepOrdCat', 'RepInvOrd', 'RepInvCat') and 
		not exists (select 'x' from param where PTyp ='YN' and UoYN11= PMCd) 
*/	/* **** Zubin 211 **** */ 
	if @row> 0 
	begin
		/* select @errmsg= 'Final Group as Design Code (Yes / No)?   (Enter ''Y'' / ''N'')' */
		select @errmsg= 'Final Group as Design Code ?'
		goto ErrHndlr
	end

	/* **** Zubin 211 (Not Required) **** */ 
/*	select @row= count('x') from inserted where UoMnuCd in ('RepInvOrd') and 
		not exists (select 'x' from param where PTyp ='YN' and UoYN8= PMCd) 
	if @row> 0 
	begin
		select @errmsg= 'Show Running Serial (Enter ''Y'' / ''N'')'
		goto ErrHndlr
	end
*/
	/* **** Zubin 211 **** */ 

	select @row= count('x') from inserted where UoMnuCd in ('RepInvLossSubRep') and 
		not exists (select 'x' from param where PTyp ='INVAL' and UOINVAL= PMCd) 
	if @row> 0 
	begin
		select @errmsg= 'Value Not Defined'
		goto ErrHndlr
	end
	--	Manoj 2.10.0

	/*** Jay 3.4(FG) ***/
	/* ***** Manali Trading module ***** */ 
	select @row= count('x') from inserted where UoMnuCd in ('RepDetOrd', 'RepAvgOrd', 'RepOrdCat',
	'RepAvgInv', 'RepDetInv','RepInvCat','REPFGDET','REPFGAVG','REPFGCAT', 'REPJTDET','REPJTAVG','REPJTCAT','REPMULTIPRCQT') and 
		not exists (select 'x' from param where PTyp ='YN' and UoYN7= PMCd) 
	if @row> 0 
	begin
		/* select @errmsg= 'Show Customer Design Code  (Yes / No)?   (Enter ''Y'' / ''N'')' */
		select @errmsg= 'Show Customer Design Code  ?'
		goto ErrHndlr
	end

	select @row= count('x') from inserted where UoMnuCd in ('RepDetInv') and 
		not exists (select 'x' from param where PTyp ='YN' and UoYN6= PMCd) 
	if @row> 0 
	begin
		/* select @errmsg= 'Show Customer RmCd (Yes / No)?   (Enter ''Y'' / ''N'')' */
		select @errmsg= 'Show Customer RmCd ?'
		goto ErrHndlr
	end
	select @row= count('x') from inserted where UoMnuCd in ('RepTreeInvst','RepTreePerform') and 
		not exists (select 'x' from param where PTyp ='YN' and UoYN= PMCd) 
	if @row> 0 
	begin
		/* select @errmsg= 'Show Bag Details (Yes / No)?   (Enter ''Y'' / ''N'')' */
		select @errmsg= 'Show Bag Details ?'
		goto ErrHndlr
	end




	/************************* Geeta 2.06 **************************/

/*  #################################  Manoj #### Ver: 2.0.6  #################################  */
/*	Date: 08/12/2003																			 */
/*	Cnange:	Condition UoMenuCd in ('RepStkLed','RepFlshLed') To UoMenuCd = 'RepStkLed'			 */
	select @row= count('x') from inserted where UoMnuCd = 'RepStkLed' and
		not exists (select 'x' from param where PTyp= 'LOCTYP' and PMCd= UoLocTypFr)
	if @row> 0
	begin
		select @errmsg= 'Invalid Location Type'
		goto ErrHndlr
	end
/* (Bef 2.06)
	select @row= count('x') from inserted where UoMnuCd in ('RepStkLed', 'RepFlashStk') and
		not exists (select 'x' from param where PTyp= 'LOCTYP' and PMCd= UoLocTypFr)
	if @row> 0
	begin
		select @errmsg= 'Invalid Location Type'
		goto ErrHndlr
	end
*/
/*  #########################################  Manoj  #########################################  */

	select @row= count('x') from inserted where UoMnuCd in ('RepStkLed') and 
		not exists (select 'x' from param where PTyp= 'VCHOPT' and PMCd= UoVchOpt)
	if @row> 0 
	begin
		select @errmsg= 'Invalid Voucher Option'
		goto ErrHndlr
	end

	select @row= count('x') from inserted where UoMnuCd in ('RepFlashStk') and 
		not exists (select 'x' from param where PTyp ='YN' and UoYN= PMCd) 
	if @row> 0 
	begin
		/* select @errmsg= 'Show Raw Material Break-Up?   (Enter ''Y'' / ''N'')' */
		select @errmsg= 'Show Raw Material Break-Up ?'
		goto ErrHndlr
	end

	/****** added RepWipBRmBkDt (Jen 2.14) ******/
	/* ****** Manali 3.5.0 - 17/11/08 - ErrMsg Changed For Check Boxes */
	/* ***** Manali 3.6.0 - 19/09/09 - MenuCd RepGldLs Changed to RepMetLs ***** */
	select @row= count('x') from inserted where UoMnuCd in ('RepMetLsWip' /*'RepGldLs'*/, 'RepRejAna', 'RepBrkMsgRm', 'RepBrkMsgPrd', 
		'RepWipBRm', 'RepWipFgBRm', 'RepWipBRmBkDt', 'RepRmReqLoc', 'RepPckLst', 'RepCTBPckLst', 'RepInvBOELst', 
		'RepBvRm') and 
		not exists (select 'x' from param where PTyp ='YN' and UoYN= PMCd) 
	if @row> 0 
	begin
		select @errmsg= 'Show Or Hide Box Detail?'
		/* select @errmsg= 'Show Box Detail?   (Enter ''Y'' / ''N'')' */
		goto ErrHndlr
	end

	select @row= count('x') from inserted where UoMnuCd in ('RepMetLsFg') and UoYN4<> '' and 
		not exists (select 'x' from param where PTyp ='YN' and UoYN4= PMCd) 
	if @row> 0 
	begin
		select @errmsg= 'Select if to Show Open/ Closed/ All Bags'
		goto ErrHndlr
	end

	select @row= count('x') from inserted where UoMnuCd in ('RepMetLsWip', 'RepMetLsFg') and 
		not exists (select 'x' from param where PTyp ='YN' and UoYN2= PMCd) 
	if @row> 0 
	begin
		select @errmsg= 'Select In Wt Or Out Wt'
		goto ErrHndlr
	end

	select @row= count('x') from inserted where UoMnuCd in ('RepMetLsWip') and 
		not exists (select 'x' from param where PTyp ='YN' and UoYN5= PMCd) 
	if @row> 0 
	begin
		select @errmsg= 'Select if Show Allowed Loss Percentage For Given Process'
		goto ErrHndlr
	end

	select @row= count('x') from inserted where UoMnuCd in ('RepMetLsWip') and 
		UoYN5= 'Y' and UoYN1<>'' 
	if @row> 0 
	begin
		select @errmsg= 'Recovery Option and Allowed Loss Percentage Option Cannnot Be Selected at a time'
		goto ErrHndlr
	end

/* **** Zubin 211 **** */ 
	/* ***** Manali 3.5.0 - 05/11/08 - ErrMsg changed for Radio Buttons ***** */
	select @row= count('x') from inserted where UoMnuCd In ('RepRejAna') and 
		not exists (select 'x' from param where PTyp ='YN' and UoYN1= PMCd) 
	if @row> 0 
	begin 
		/* select @errmsg= 'Invalid ''Show Qty/Wt'' Option'  */
		select @errmsg= 'Select either Quantity Or Weight' 
		goto ErrHndlr 
	end 
/* **** Zubin 211 **** */ 

	/* ***** Manali 3.5.0 - 17/11/08 - ErrMsg changed for Radio Buttons ***** */
	select @row= count('x') from inserted where UoMnuCd in ('RepBrkMsgRm', 'RepBrkMsgPrd') and UoBrkMsg<> '' and 
		not exists (select 'x' from param where PTyp ='BRKMSG' and UoBrkMsg= PMCd) 
	if @row> 0 
	begin
		select @errmsg= 'Select If To Show Broken/ Missing/ Both Details?'
		/* select @errmsg= 'Show Broken/ Missing/ Both Details?   (Enter ''B'' / ''M'' / Blank)' */
		goto ErrHndlr
	end

	/* ***** Manali 3.5.0 - 17/11/08 - ErrMsg changed for Radio Buttons ***** */
	select @row= count('x') from inserted where UoMnuCd in ('RepBrkMsgRm') and 
		not exists (select 'x' from param where PTyp ='SHPCDGR' and UoShpCd= PMCd) 
	if @row> 0 
	begin
		select @errmsg= 'Select If To Show Raw Material Shape/ Code As The Detail?  '
		/* select @errmsg= 'Show Raw Material Shape/ Code As The Detail?   (Enter ''S'' / ''R'')' */
		goto ErrHndlr
	end

	/*** (Jen 2.13) ***/
	/* ***** Manali 3.5.0 - 17/11/08 - ErrMsg changed for Radio Buttons ***** */
	select @row= count('x') from inserted where UoMnuCd in ('RepRejAna', 'RepBrkMsgRm', 'RepBrkMsgPrd') and 
		not exists (select 'x' from param where PTyp ='YN' and UoYN3= PMCd) 
	if @row> 0 
	begin
		select @errmsg= 'Show All Production Lines?   '
		/* select @errmsg= 'Show All Production Lines?   (Enter ''Y'' / ''N'')' */
		goto ErrHndlr
	end
	/*** (Jen 2.13) ***/

	/* ***** Manali Trading Module ***** */
	select @row= count('x') from inserted where UoMnuCd in ('RepFbLst', 'RepFRLst', 'RepFwbLst', 'RepFwrLst', 'RepCustmStk', 'RepSOPnd', 'RepJtPnd', 'RepJtOpnOrd') and 
		not exists (select 'x' from param where PTyp ='YN' and UoYN1= PMCd) 
	if @row> 0 
	begin
		select @errmsg= 'Show Box Detail ?'
		/* select @errmsg= 'Show Box Detail?   (Enter ''Y'' / ''N'')' */
		goto ErrHndlr
	end

	select @row= count('x') from inserted where UoMnuCd in ('RepFbLst') And UoYN3 <>'' And 
		not exists (select 'x' from param where PTyp ='YN' and UoYN3= PMCd) 
	if @row> 0 
	begin
		select @errmsg= 'Show CTB Pending / CTB Closed / All (Enter ''Y'' / ''N'' / Blank )'
		goto ErrHndlr
	end

	select @row= count('x') from inserted where UoMnuCd in ('RepStkSumm') and 
		not exists (select 'x' from param where PTyp ='YN' and UoYN= PMCd) 
	if @row> 0 
	begin
		select @errmsg= 'Show Pure Weight ?'
		/* select @errmsg= 'Show Pure Weight (Yes / No)' */
		goto ErrHndlr
	end

/*
	select @row= count('x') from inserted where UoMnuCd in ('RepBagHist') and 
		not exists (select 'x' from param where PTyp ='YN' and UoYN= PMCd) 
	if @row> 0 
	begin
		select @errmsg= 'Show Rm Stock Rate Instead Of Rm Size?   (Enter 'Y' / 'N')'
		goto ErrHndlr
	end
*/

	select @row= count('x') from inserted where UoMnuCd in ('RepBagHist') and 
		not exists (select 'x' from param where PTyp ='BHISTOPT' and UoPMCdFr= PMCd) 
	if @row> 0 
	begin
		select @errmsg= 'Specify The Group Option For The Rm Summary Sub Report'
		goto ErrHndlr
	end

	select @row= count('x') from inserted where UoMnuCd in ('RepDetOrd', 'RepAvgOrd', 'RepOrdCat') and 
		not exists (select 'x' from param where PTyp ='YN' and UoYN1= PMCd) 
	if @row> 0 
	begin
		select @errmsg= 'Select If To Show Amended Qty or Entered Qty '
		/* select @errmsg= 'Show Amended Qty or Entered Qty?   (Enter ''Y'' / ''N'')' */
		goto ErrHndlr
	end

	/************************ urmi 2.06 Show Remark fields **********************************/
	/*** Jay 3.4(FG) ***/ 
	/* ***** Manali Trading module ***** */
	/* ***** Manali 3.9.0 - Different Option asked For Order Master Level Remarks */
	select @row= count('x') from inserted where UoMnuCd in ('RepDetOrd','REPFGDET', 'REPJTDET') and 
		not exists (select 'x' from param where PTyp ='YN' and UoYN6= PMCd) 
	if @row> 0 
	begin
		select @errmsg= 'Show Order Master Remarks/instructions in Report ?'
		/* select @errmsg= 'Show Remarks/instructions in Report? (Enter ''Y'' / ''N'')' */
		goto ErrHndlr
	end
	
	/* ***** Manali 3.9.0 - Different Option asked For Order Design Level Remarks */
	select @row= count('x') from inserted where UoMnuCd in ('RepDetOrd','REPFGDET', 'REPJTDET') and 
		not exists (select 'x' from param where PTyp ='YN' and UoYN26= PMCd) 
	if @row> 0 
	begin
		select @errmsg= 'Show Order Design Remarks/instructions in Report ?'
		goto ErrHndlr
	end	
	/************************ urmi 2.06 Show Remark fields **********************************/
	/*** Jay 3.4(FG) ***/
	/* ***** Manali Trading module ***** */	
	select @row= count('x') from inserted where UoMnuCd in ('RepInvCustStmt','RepAvgOrd','REPFGAVG', 'REPJTAVG') and 
		not exists (select 'x' from param where PTyp ='YN' and UoYN2= PMCd) 
	if @row> 0 
	Begin
		select @errmsg= 'Show All Value ?'
		/* select @errmsg= 'Show All Value? (Enter ''Y'' / ''N'')' */
		goto ErrHndlr
	end

	select @row= count('x') from inserted where UoMnuCd in ('RepDetInv', 'RepDetInvCTB', 
		'RepAvgInv', 'RepAvgInvCTB', 'RepInvCat', 'RepInvCTBCat', 
		'RepSalAnaRm') and 
		not exists (select 'x' from param where PTyp ='INVAL' and UoInVal= PMCd) 
	if @row> 0 
	begin
		select @errmsg= 'Show Actual, Invoice, Stock Or Customs Values?   (Enter ''A'' / ''I''/ ''S''/ ''Z'')'
		goto ErrHndlr
	end
	/************Emr208******************/
	
	select @row= count('x') from inserted where UoMnuCd in ('RepSalAnaDsg','RepSalAnaLab') and 
		not exists (select 'x' from param where PTyp ='INVAL' and UoInVal= PMCd and PMCd<>'S') 
	if @row> 0 
	begin
		select @errmsg= 'Show Actual, Invoice Or Customs Values?   (Enter ''A'' / ''I''/ ''Z'')'
		goto ErrHndlr
	end
	/************Emr208******************/

	select @row= count('x') from inserted where UoMnuCd in ('RepSalAnaDsg','RepSalAnaLab','RepSalAnaRm') and 
		not exists (select 'x' from param where PTyp ='YN' and UoYN3= PMCd) 
	if @row> 0 
	begin
		select @errmsg= 'Consider Date as Invoice Date Or Printing Date?'
		goto ErrHndlr
	end


	/*** Jay 3.4(FG) ***/
	/* ***** Manali Trading Module ***** */
	/*** Manali 3.8.0 - Jt Lcl Rep removed from Following Check ****/
	/***** Manali 3.8.0 - RepPckLstInvBag added ****/
    /* 3.11.0 - RepInvBOELst removed */
	select @row= count('x') from inserted where UoMnuCd in ('RepInv', 'RepGRFormFrnt', 
		'RepShpBill', 'RepGSPForm', 'RepGSPCert', 'RepGRFormBack', 'RepInvLossSubRep', 
		'RepGSPDetLet', 'RepDeclnLet', 'RepPckLst', 'RepCTBPckLst',/* 'RepInvBOELst', */
		'RepPckLstInv', 'RepInvLoose', 'RepShpBillLoose', 'RepGrFormLoose', 
		'RepPckLstLoose','RepFgInv','RepFgPckLst','RepFgShpBill','RepFgGrFrnt',
		'RepFgGrBk','RepFgGsp','RepFgLclInv', 'RepJTPckLst', 'RepJtGrBk',
		'RepJTInv','RepJTShpBill','RepJtGrFrnt', 'RepJtGsp'/*,'RepJtLclInv'*/, 'RepPckLstInvBag') and 
		(	not exists (select 'x' from InvHd where InCoCd= UoCoCdFr and InTc= UoInTcFr and 
				InYy= UoInYyFr and InChr= UoInChrFr and InNo= UoInNoFr) 
			AND 
			not exists (select 'x' from InvHd where InCoCd= UoCoCdFr and InExpNo= UoInExpNoFr) 
		)			
	if @row> 0 
	begin
		select @errmsg= 'Invoice Voucher Number Not Defined'
		goto ErrHndlr
	end
/*
	select @row= count('x') from inserted where UoMnuCd in ('RepJTInv','RepJTShpBill','RepJtGrFrnt','RepJtGsp') and 
		(	not exists (select 'x' from OrdMst where OmCoCd= UoCoCdFr and OmTc= UoInTcFr and 
				OmYy= UoInYyFr and OmChr= UoInChrFr and OmNo= UoInNoFr) 
		)			
	if @row> 0 
	begin
		select @errmsg= 'Invoice Voucher Number Not Defined'
		goto ErrHndlr
	end
*/
	/****** added RepGRFormBackMulti (Jen 2.14) ******/
	/****** added RepGSPFORMMulti, RepGSPDetLetMulti, RepGSPCertMulti, RepDeclnLetMulti (Jen 3.01) ******/
	/*** (Jen 2.13) ***/
	select @row= count('x') from inserted where UoMnuCd in ('RepGrFormFrntMulti', 'RepShpBillMulti', 
			'RepGRFormBackMulti', 'RepGSPFORMMulti', 'RepGSPDetLetMulti', 'RepGSPCertMulti', 'RepDeclnLetMulti') and 
			UoInExpNoSel= '' 
	if @row> 0 
	begin
		select @errmsg= 'Invoice Export Number(s) Cannot Be Blank'
		goto ErrHndlr
	end

	/****** added RepGSPFORMMulti, RepGSPDetLetMulti, RepGSPCertMulti, RepDeclnLetMulti (Jen 3.01) ******/
	select @row= count('x') from inserted where UoMnuCd in ('RepGrFormFrntMulti', 'RepShpBillMulti', 
			'RepGRFormBackMulti', 'RepGSPFORMMulti', 'RepGSPDetLetMulti', 'RepGSPCertMulti', 'RepDeclnLetMulti') and 
			not exists (select 'x' from InvHd where InCoCd= UoCoCdFr and 
					UoInExpNoSel+ ',' like '%''' + InExpNo + ''',%') 
	if @row> 0 
	begin
		select @errmsg= 'Invoice Export Number(s) Not Defined'
		goto ErrHndlr
	end 
	/*** (Jen 2.13) ***/


	/*** (Jen 2.12) ***/
	/* 'RepInv', 'RepGRFormFrnt', 'RepShpBill', 'RepShpBillLoose', 'RepGrFormLoose' */
	select @row= count('x') from inserted where UoMnuCd in ('RepInvLoose') and 
		not exists (select 'x' from param where PTyp= 'RMGRP' and PMCd= 'INV' and PSCD= UoRmGrp)
	if @row> 0 
	begin
		select @errmsg= 'Invalid Raw Material Group Option'
		goto ErrHndlr
	end

	/* (Bef 2.12 Patch)
	select @row= count('x') from inserted where 
		(Select hExpLocal from Head where HCd='ZSELF' and HCoCd= UoCoCdFr)= 'Y' and 
		UoMnuCd in ('RepShpBill', 'RepShpBillLoose', 'RepGRFormFrnt', 'RepGrFormLoose') and 
		not exists (select 'x' from param where PTyp= 'GRLSSUMM' and PMCd= UoPmCdFr)
	if @row> 0 
	begin
		select @errmsg= 'Invalid Metal Loss Group Option'
		goto ErrHndlr
	end
	(Bef 2.12 Patch) */
	/*** (Jen 2.12) ***/

	/****** (Jen 3.01) added RepGSPFORMMulti ******/
	/*** Jay 3.4(FG) ***/
	/* ***** Manali Trading Module ***** */
	/**** Manali 3.8.0 - RepPckLstInvBag added ****/
	select @row= count('x') from inserted where UoMnuCd in ('RepInv', 'RepGRFormFrnt', 
		'RepShpBill', 'RepGSPForm', 'RepPckLstInv', 'RepGrFormFrntMulti', 'RepShpBillMulti', 'RepGSPFORMMulti',
		'RepFgInv','RepFgPckLst','RepFgShpBill','RepFgGrFrnt','RepFgGsp',
		'RepJtInv','RepJtPckLst','RepJtShpBill','RepJtGrFrnt','RepJtGsp', 'RepPckLstInvBag') and 
		not exists (select 'x' from param where PTyp= 'INGR' and PMCd= UoInGrp)
	if @row> 0 
	begin
		select @errmsg= 'Invalid Invoice Group Option'
		goto ErrHndlr
	end

	/* Zubin 212 */ 
	/****** (Jen 3.01) added RepGSPFORMMulti ******/
	/*** Jay 3.4(FG) ***/ 
	/***** Manali Tradiing Module *****/
	select @row= count('x') from inserted where UoMnuCd in ('RepGSPForm', 'RepGSPFORMMulti','RepFgGsp', 'RepJtGsp') and 
		not exists (select 'x' from param where PTyp= 'YN' and PMCd= UoYN4 and PSCd= '') and 
		(Select hExpLocal from Head where HCd='ZSELF' and HCoCd =UoCoCdFr)= 'N' 
	if @row> 0 
	begin
		select @errmsg= 'Invalid ''Show Net Weight'' Option'
		goto ErrHndlr
	end
	
	/*** Jay 3.4 ***/
	/* ***** Manali Trading Module ***** */
	select @row= count('x') from inserted where UoMnuCd in ('RepInv', 'RepInvLoose','RepFgInv','RepJtInv') and 
		not exists (select 'x' from param where PTyp= 'YN' and PMCd= UoYN5 and PSCd= '')
	if @row> 0 
	begin
		select @errmsg= 'Invalid ''Show Declaration For Customer'' Option'
		goto ErrHndlr
	end
	/* Zubin 212 */ 

        /***Geeta**Emr208***Blank allowed for local export Only*/
	/* (Bef 2.13)
	select @row= count('x') from inserted where UoMnuCd in ('RepInv', 'RepGRFormFrnt', 
		'RepShpBill') and 
		(Select hExpLocal from Head where HCd='ZSELF'  and HCoCd =UoCoCdFr)='N' and 
		not exists (select 'x' from param where PTyp= 'INDET' and PMCd= UoInDet)
	(Bef 2.13) */

	/*** (Jen 2.13) ***/
	/*** Jay 3.4(FG) ***/
	/* ***** Manali Trading Module ***** */
	/* ***** Manali 3.8.0 - RepPckLstInvBag added ***/
	select @row= count('x') from inserted where UoMnuCd in ('RepInv', 'RepGRFormFrnt', 'RepPckLstInvBag',
		'RepShpBill', 'RepGrFormFrntMulti', 'RepShpBillMulti','RepFgInv','RepFgShpBill','RepFgGrFrnt', 'RepJtInv','RepJtShpBill','RepJtGrFrnt') 
		and not exists (select 'x' from param where PTyp= 'INDET' and PMCd= UoInDet)
	/*** (Jen 2.13) ***/
	if @row> 0 
	begin
		select @errmsg= 'Invalid Invoice ''Detail Sort'' Option'
		goto ErrHndlr
	end

	/*** (Jen 2.14 Next) ***/
	/*** Jay 3.4(FG)***/
	/***** Manali Trading Module *****/
	select @row= count('x') from inserted where UoMnuCd in ('RepGRFormFrnt', 'RepGrFormLoose', 'RepGrFormFrntMulti','RepFgGrFrnt', 'RepJtGrFrnt') and 
		not exists (select 'x' from param where PTyp= 'YN' and PMCd= UoYN7 and PSCd= '')
	if @row> 0 
	begin
		select @errmsg= 'Invalid ''Show Payment Terms'' Option'
		goto ErrHndlr
	end
	/*** (Jen 2.14 Next) ***/

	/* (Bef 2.13)
        Select @row= count('x') from inserted where UoMnuCd in ('RepInv', 'RepGRFormFrnt', 
		'RepShpBill') and 
		((Select hExpLocal from Head where HCd='ZSELF'  and HCoCd =UoCoCdFr)='Y' and UoInDet<>'') and
		not exists (select 'x' from param where PTyp= 'INDET' and PMCd= UoInDet)
	if @row> 0 
	begin
		select @errmsg= 'Invalid Invoice ''Detail Sort'' Option'
		goto ErrHndlr
	end
	(Bef 2.13) */

/* **** Zubin 211 **** */ 
        Select @row= count('x') from inserted where UoMnuCd in ('RepInvCustStmt', 'RepDetInv', 'RepAvgInv', 
		'RepInvCat', 'RepPckLstInv', 'RepAnnexValAddn') and 
		Not Exists (select 'x' from param where PTyp= 'INDET' and PMCd= UoInDet and PSCd= '')
	if @row> 0 
	begin
		select @errmsg= 'Invalid Invoice ''Detail Sort'' Option'
		goto ErrHndlr
	end

	/***** Manali 3.8.0 - 'RepValAddnDTA' added *****/ 
	Select @row= count('x') from inserted where UoMnuCd in ('RepValAddn', 'RepValAddnDta') and UoInDet<> '' and 
		Not Exists (select 'x' from param where PTyp= 'INDET' and PMCd= UoInDet and PSCd= '')
	if @row> 0 
	begin
		select @errmsg= 'Invalid Invoice ''Detail Sort'' Option'
		goto ErrHndlr
	end

  Select @row= count('x') from inserted where UoMnuCd in ('RepValAddn', 'RepValAddnDta') and (UoInGrp1= '' and UoInDet= '' ) 
	if @row> 0 
	begin
		select @errmsg= ' ''Detail Sort'' Option Cannot Be Blank When ''Grp Combination2'' Is Blank'
		goto ErrHndlr
	end

        Select @row= count('x') from inserted where UoMnuCd in ('RepValAddn', 'RepValAddnDta') and UoInGrp1<> '' and UoInDet<> ''  
	if @row> 0 
	begin
		select @errmsg= ' ''Detail Sort'' Option Should Be Blank When ''Grp Combination2'' Is Not Blank'
		goto ErrHndlr
	end

    Select @row= count('x') from inserted where UoMnuCd = 'RepInvCat' and UoInDet Not In ('DMCD', 'INVSR', 'TRNO/SR')
	if @row> 0 
	begin
		select @errmsg= 'Only ''DMCD'', ''INVSR'' And ''TRNO/SR'' Detail Sort Options Are Applicable For Invoice Catalogue '
		goto ErrHndlr
	end

  /* ***** Manali Trading Module ***** */
	/*****  Manali 3.8.0 - RepPckLstInvBag added ****/
	/**** Following Check Commented to allow Use of 'TRNO/SR'  ***/
	/*** Select @row= count('x') from inserted where UoMnuCd in ('RepInv', 'RepGRFormFrnt', 'RepShpBill', 
		'RepGrFormFrntMulti', 'RepShpBillMulti','RepFgInv','RepFgShpBill','RepFgGrFrnt', 'RepValAddn', 'RepValAddnDta',
		'RepJtInv','RepJtShpBill','RepJtGrFrnt', 'RepPckLstInvBag') 
		and UoInDet = 'TRNO/SR'
	if @row> 0 
	begin
		select @errmsg= '''TRNO/SR'''' Detail Sort Option Is Not Applicable For This Report'
		goto ErrHndlr
	end ****/

/* **** Zubin 211 **** */ 

	/********************Geeta************Emr206**************/
	/* **** Zubin 211 (UoYN3 not required) **** */ 
/*	select @row= count('x') from inserted where UoMnuCd = 'RepInv' and UoInDet='INVSR'  
		and not exists (select 'x' from param where PTyp= 'YN' and PMCd= UoYN3)
	if @row> 0 
	begin
		select @errmsg= 'Invalid Running Serial Option (Enter Y/N )'
		goto ErrHndlr
	end
	
	select @row= count('x') from inserted where UoMnuCd = 'RepInv' and (( UoInDet='INVSR'  
		and UoYN3='') or ( UoInDet<>'INVSR' and UoYN3<>''))
	if @row> 0 
	begin
		select @errmsg= 'Running Serial Should Entered Only for Det Level='INVSR''
		goto ErrHndlr
	end
*/
	/* **** Zubin 211 (UoYN3 not required) **** */ 	

	select @row= count('x') from inserted where UoMnuCd in ('RepDetInv', 'RepDetInvCTB', 
		'RepAvgInv', 'RepAvgInvCTB', 'RepInvCuststmt') and 
		not exists (select 'x' from param where PTyp ='YN' and UoYN1= PMCd) 
	if @row> 0 
	begin
		select @errmsg= 'Show Average Weights ?'
		/* select @errmsg= 'Show Average Weights (Enter ''Y'' / ''N'')' */
		goto ErrHndlr
	end

/*	select @row= count('x') from inserted where UoMnuCd in ('RepDetInv', 'RepAvgInv') 
		and UoRmGrp = 'RMRSZRt'
	if @row> 0 
	begin
		select @errmsg= 'Range Size Option Not Available With Invoice Reports'
		goto ErrHndlr
	end
*/
	select @row= count('x') from inserted where UoMnuCd in ('RepCustmStk') and 
		not exists (select 'x' from param where PTyp ='YN' and UoYN= PMCd) 
	if @row> 0 
	begin
		select @errmsg= 'Show Pending Stock ?'
		/* select @errmsg= 'Show Pending Stock (Enter ''Y'' / ''N'')' */
		goto ErrHndlr
	end

	select @row= count('x') from inserted where UoMnuCd in ('RepInvProfitMrg') and 
		not exists (select 'x' from param where PTyp ='INVAL' and UoInVal= PMCd) 
	if @row> 0 
	begin
		select @errmsg= 'Enter Base Value As Actual/ Invoice/ Customs/ Stock'
		goto ErrHndlr
	end

	select @row= count('x') from inserted where UoMnuCd in ('RepInvProfitMrg') and 
		not exists (select 'x' from param where PTyp ='INVAL' and UoInVal1= PMCd) 
	if @row> 0 
	begin
		select @errmsg= 'Enter Value Other Than Base As Actual/ Invoice/ Customs/ Stock '
		goto ErrHndlr
	end

	select @row= count('x') from inserted where UoMnuCd in ('RepOrdProfitMrg') and 
		not exists (select 'x' from param where PTyp= 'ORDVAL' and UoOrdVal= PMCd) 
	if @row> 0 
	begin
		select @errmsg= 'Enter Base Value As Cost/ Sale'
		goto ErrHndlr
	end

	select @row= count('x') from inserted where UoMnuCd in ('RepOrdProfitMrg') and 
		not exists (select 'x' from param where PTyp= 'ORDVAL' and UoOrdVal1= PMCd) 
	if @row> 0 
	begin
		select @errmsg= 'Enter Value Other Than Base As Cost/ Sale'
		goto ErrHndlr
	end

	/* **** Manali 3.5.0 - 05/11/08 - ErrMsg Changed for Radio Buttons */
	select @row= count('x') from inserted where UoMnuCd in ('RepInvProfitMrg', 'RepSalAnaDsg') and 
		not exists (select 'x' from param where PTyp ='YN' and UoYN= PMCd) 
	if @row> 0 
	begin
		/* select @errmsg= 'Show Value/ Weight (Enter ''Y'' / ''N'')' */
		select @errmsg= 'Select if to show Value or Weight'
		goto ErrHndlr
	end

	/* **** Manali 3.5.0 - 05/11/08 - ErrMsg Changed for Radio Buttons */
	select @row= count('x') from inserted where UoMnuCd in ('RepInvProfitMrg') and 
		not exists (select 'x' from param where PTyp ='YN' and UoYN1= PMCd) 
	if @row> 0 
	begin
 		/* select @errmsg= 'Get LME Cost Rates From Invoice (Enter ''Y'' / ''N'')' */
		select @errmsg= 'Select if to Get LME Cost Rates From Invoice Or Order'
		goto ErrHndlr
	end
	/* **** Manali 3.5.0 - 05/11/08 - ErrMsg Changed for Radio Buttons */

	select @row= count('x') from inserted where UoMnuCd in ('RepValAddn','RepValAddnDta', 'RepAnnexValAddn') and 
		UoInGrp= '' and UoInGrp1<> '' 
	if @row> 0 
	begin
		select @errmsg= 'Cannot Enter Value In Group2 When Group1 Is Blank'
		goto ErrHndlr
	end

	select @row= count('x') from inserted where UoMnuCd in ('RepValAddn','RepValAddnDta', 'RepInvCuststmt', 'RepAnnexValAddn') and 
		UoInGrp<> '' and not exists (select 'x' from param where PTyp ='INGR' and UoInGrp= PMCd) 
	if @row> 0 
	begin
		select @errmsg= 'Group1 Not Defined'
		goto ErrHndlr
	end

	select @row= count('x') from inserted where UoMnuCd in ('RepValAddn','RepValAddnDta', 'RepInvcustStmt', 'RepAnnexValAddn') and 
		UoInGrp1<> '' and not exists (select 'x' from param where PTyp ='INGR' and UoInGrp1= PMCd) 
	if @row> 0 
	begin
		select @errmsg= 'Group2 Not Defined'
		goto ErrHndlr
	end

	select @row= count('x') from inserted where UoMnuCd in ('RepValAddn','RepValAddnDta') and 
		not exists (select 'x' from param where PTyp ='FNDOPT' and PMCd= UoPMCdFr) 
	if @row> 0 
	begin
		select @errmsg= 'Specify Whether Finding Cost To Be Included Under Total Metal Cost Or Studding Cost'
		goto ErrHndlr
	end

/* **** Zubin 211 **** */ 
	select @row= count('x') from inserted where UoMnuCd in ('RepValAddn','RepValAddnDta') and 
		Not Exists (select 'x' from param where PTyp= 'YN' and PMCd= UOYN1) 
	if @row> 0 
	begin
		select @errmsg= 'Invalid ''Annexure Attached'' Option'
		goto ErrHndlr
	end

	/*** Bef 2.12 Patch 2 *** 
	select @row= count('x') from inserted where UoMnuCd in ('RepValAddn') and 
		Not Exists (select 'x' from param where PTyp= 'YN' and PMCd= UOYN2) 
	*** Bef 2.12 Patch 2 ***/	

	/*** Jen (2.12 Patch 2) ***/
	select @row= count('x') from inserted where UoMnuCd in ('RepValAddn') and 
		exists(Select 'x' from Head where HCoCd= UoCoCdFr and hExpLocal= 'N') and 
		Not Exists (select 'x' from param where PTyp= 'YN' and PMCd= UOYN2) 
	/*** Jen (2.12 Patch 2) ***/
	if @row> 0 
	begin
		select @errmsg= 'Hide Cols 13-16 ?' 
		/* select @errmsg= 'Invalid ''Hide Cols 13-16'' Option'  */
		goto ErrHndlr
	end
/* **** Zubin 211 **** */ 


	/*** Jen (2.14 Next Patch 1) ***/
	select @row= count('x') from inserted where UoMnuCd in ('RepValAddn') and 
		exists(Select 'x' from Head where HCoCd= UoCoCdFr and hExpLocal= 'N') and 
		Not Exists (select 'x' from param where PTyp= 'YN' and PMCd= UOYN4) 
	if @row> 0 
	begin
		select @errmsg= 'Hide Column 12 ?' 
		/* select @errmsg= 'Invalid ''Hide Column 12'' Option'  */
		goto ErrHndlr
	end
	/*** Jen (2.14 Next Patch 1) ***/


	/*** (Jen 2.13) ***/
	select @row= count('x') from inserted where UoMnuCd in ('RepValAddn') and 
		not exists (select 'x' from param where PTyp= 'YN' and PMCd= UOYN3 and PSCd= '') 
	if @row> 0 
	begin
		select @errmsg= 'Show Loss Percentage ?'
		/* select @errmsg= 'Invalid ''Show Loss Percentage'' Option ' */
		goto ErrHndlr
	end
	/*** (Jen 2.13) ***/

	/****** Sachin 3.5.0 ******/	
	select @row= count('x') from inserted where UoMnuCd in ('RepValAddn') and 		
		Not Exists (select 'x' from param where PTyp= 'YN' and PMCd= UOYN5) 
	if @row> 0 
	begin
		select @errmsg= 'Show Profit Percentage ?'
		/* select @errmsg= 'Invalid ''Show Profit Percentage'' Option'  */
		goto ErrHndlr
	end	
	/****** Sachin 3.5.0 ******/

	/****** Manali 3.9.0 ******/	
	select @row= count('x') from inserted where UoMnuCd in ('RepValAddn') and 		
		UOYN5 = 'Y' and (UOYN2='Y' or UOYN4='Y')  
	if @row> 0 
	begin
		select @errmsg= 'Cannot Hide Column 12 and 13-15 When Show Profit Percentage Option Selected'
		goto ErrHndlr
	end	
	/****** Manali 3.9.0 ******/

/* **** Zubin 211 (Not required) **** */ 
/*	select @row= count('x') from inserted where UoMnuCd in ('RepAnnexValAddn') and 
		not exists (select 'x' from param where PTyp = 'YN' and UoYN= PMCd) 
	if @row> 0 
	begin
		select @errmsg= 'Show Running Serial or Invoice Serial In The Report Detail?   (Enter 'Y' / 'N')'
		goto ErrHndlr
	end
*/
/* **** Zubin 211 (Not required) **** */ 
	select @row= count('x') from inserted where UoMnuCd in ('RepInv', 'RepPckLstInv', 'RepInvLoose') and 
		not exists (select 'x' from param where PTyp ='INWT' and UoInWt= PMCd) 
	if @row> 0 
	begin
		select @errmsg= 'Show Invoice Or Customs Weights?   (Enter ''I''/ ''Z'')'
		goto ErrHndlr
	end

	
	/* 	*** (Bef 2.13) ***
		select @row= count('x') from inserted where UoMnuCd in ('RepInv', 'RepInvLoose') and 
		not exists (select 'x' from param where PTyp ='YN' and UoYN2= PMCd) 
		if @row> 0 
		begin
			select @errmsg= 'Show Metal Loss Summary? (Enter ''Y''/ ''N'')'
			goto ErrHndlr
		end
	*/

	/*** (Jen 2.13) ***/
	select @row= count('x') from inserted where UoMnuCd in ('RepInv') and 
		not exists (select 'x' from param where PTyp ='YN' and UoYN2= PMCd) 
	if @row> 0 
	begin
		select @errmsg= 'Show Metal Loss Summary? '
		/* select @errmsg= 'Show Metal Loss Summary? (Enter ''Y''/ ''N'')' */
		goto ErrHndlr
	end
	
	/*** Jay 3.4(FG) ***/
	/* ***** Manali Trading Module ***** */
	select @row= count('x') from inserted where UoMnuCd in ('RepInv','RepFgInv', 'RepJtInv') and 
		not exists (select 'x' from param where PTyp ='YN' and UoYN6= PMCd) 
	if @row> 0 
	begin
		select @errmsg= 'Show Dia/ ColStn Weight? '
		/* select @errmsg= 'Show Dia/ ColStn Weight? (Enter ''Y''/ ''N'')' */
		goto ErrHndlr
	end
	/*** (Jen 2.13) ***/
/* *** Jay 2.14(LclInv) *** */
/* **** Zubin 211 **** */
	/*** Jay 3.4(FG) ***/
	/* ***** Manali Trading Module ***** */
	select @row= count('x') from inserted where UoMnuCd in ('RepInv', 'RepLclInv','RepFgInv','RepFgLclInv','RepJtInv','RepJtLclInv') and 
		Not Exists (select 'x' from param where PTyp ='YN' and UoYN3= PMCd and PSCd= '') 
	if @row> 0 
	begin	
		select @errmsg= 'Show Detail Line Demarcation? '
		/* select @errmsg= 'Show Detail Line Demarcation? (Enter ''Y''/''N'')' */
		goto ErrHndlr
	end
/* **** Zubin 211 **** */

	/* ***** Manali 3.6.0 - Show Running Sr in LclInv ***** */
	select @row= count('x') from inserted where UoMnuCd ='RepLclInv' and 
		Not Exists (select 'x' from param where PTyp ='YN' and UoYN= PMCd and PSCd= '') 
	if @row> 0 
	begin	
		select @errmsg= 'Show Running Sr? '
		goto ErrHndlr
	end

	select @row= count('x') from inserted where UoMnuCd in ('RepLclInv','RepFgLclInv','RepJtLclInv') and 
		Not Exists (select 'x' from param where PTyp ='YN' and UoYN4= PMCd and PSCd= '') 
	if @row> 0 
	begin	
		select @errmsg= 'Show Rm Summary? '
		goto ErrHndlr
	end


	/* ***** Manali 3.6.0 Show Net wt or DC Wt ***** */
/*	select @row= count('x') from inserted where UoMnuCd = 'RepInv' and 
		UoYn4='Y' and UoYn6='Y'
	if @row> 0 
	begin
		select @errmsg= 'Select either  Show Net Wt or Show DC Wt'
		goto ErrHndlr
	end
*/

	select @row= count('x') from inserted where UoMnuCd in ('RepBnk') and 
		not exists (select 'x' from Param where PTyp ='BKCTG' and UoPmCdFr= PMCd) 
	if @row> 0 
	begin
		select @errmsg= 'Invalid Bank/Agent Category'
		goto ErrHndlr
	end
/*
	select @row=count('x') from inserted where UoMnuCd in ('RepInv') and 
		not exists (select 'x' from param where PTyp = 'YN' and UoYN1= PMCd) 
	if @row>0 
	Begin
		select @errmsg= 'Show Gross Weight Field?   (Enter 'Y' / 'N')'
		goto ErrHndlr
	end
	select @row=count('x') from inserted where UoMnuCd in ('RepInv') and 
		not exists (select 'x' from param where PTyp = 'YN' and UoYN2= PMCd) 
	if @row>0 
	Begin
		select @errmsg= 'Show Metal Loss Summary Sub Report?   (Enter 'Y' / 'N')'
		goto ErrHndlr
	end
	select @row=count('x') from inserted where UoMnuCd in ('RepInv') and 
		not exists (select 'x' from param where PTyp = 'YN' and UoYN3= PMCd) 
	if @row>0 
	Begin
		select @errmsg= 'Show Conversion Rate/ Rupee Equivalent?   (Enter 'Y' / 'N')'
		goto ErrHndlr
	end
*/

	select @row= count('x') from inserted where UoMnuCd in ('RepFbLst','RepFrLst') and 
		not exists (select 'x' from param where PTyp = 'YN' and UoYN2= PMCd) 
	if @row> 0 
	begin
		select @errmsg= 'Show Pending/ All Fg?  '
		/* select @errmsg= 'Show Pending/ All Fg?   (Enter ''Y'' / ''N'')' */
		goto ErrHndlr
	end

	/* **** Manali 3.5.0 - 05/11/08 - ErrMsg changed for Radio Buttons **** */
	select @row= count('x') from inserted where UoMnuCd in ('RepSalAnaDsg') and 
		not exists (select 'x' from param where PTyp ='YN' and UoYN1= PMCd) 
	if @row> 0 
	begin
		/* select @errmsg= 'Show Inv Or Ord Customer (Enter ''Y'' for Inv / ''N'' For Ord)'*/
		select @errmsg= 'Select either to Show Invoice Customer or Order Customer in the report'
		goto ErrHndlr
	end
	/* **** Manali 3.5.0 - 05/11/08 - ErrMsg changed for Radio Buttons **** */

	/* **** Manali 3.5.0 - 05/11/08 - ErrMsg changed for Radio Buttons **** */
	select @row= count('x') from inserted where UoMnuCd in ('RepSalAnaRm', 'RepSalAnaLab') and 
		not exists (select 'x' from param where PTyp ='YN' and UoYN= PMCd) 
	if @row> 0 
	begin
		/* select @errmsg= 'Show Inv Or Ord Customer (Enter ''Y'' for Inv / ''N'' For Ord)' */
		select @errmsg= 'Select either to Show Invoice Customer or Order Customer'
		goto ErrHndlr
	end
	/* **** Manali 3.5.0 - 05/11/08 - ErrMsg changed for Radio Buttons **** */

	select @row= count('x') from inserted where UoMnuCd in ('RepSalAnaRm') and 
		not exists (select 'x' from param where PTyp ='SARMGR' and UoSaRmGrp= PMCd) 
	if @row> 0 
	begin
		select @errmsg= 'Enter Valid Report Detail Level'
		goto ErrHndlr
	end

	select @row= count('x') from inserted where UoMnuCd in ('RepSalAnaLab') and 
		not exists (select 'x' from param where PTyp ='SALABGR' and UoSaLabGrp= PMCd) 
	if @row> 0 
	begin
		select @errmsg= 'Enter Valid Report Detail Level'
		goto ErrHndlr
	end

	/* *** Jay 3.01 *** [RepRtCmp added] */
	select @row= count('x') from inserted where UoMnuCd in ('RepRmRt', 'RepLabRt', 'RepRtCmp') and 
		not exists (select 'x' from param where PTyp= 'CMCTG' and UoCmCtgFr= PMCd) 
	if @row> 0 
	begin
		select @errmsg= 'Enter Valid Customer Category'
		goto ErrHndlr
	end

	/* *** Jay 3.01 *** [RepRtCmp] */
	select @row= count('x') from inserted where UoMnuCd in ('RepRtCmp') and 
		not exists (select 'x' from param where PTyp= 'CMCTG' and UoCmCtgTo= PMCd) 
	if @row> 0 
	begin
		select @errmsg= 'Enter Valid Source Customer Category'
		goto ErrHndlr
	end
	/* *** Jay 3.01 *** [RepRtCmp] */

	select @row= count('x') from inserted where UoMnuCd in ('RepRmRt', 'RepLabRt') and UoYyMmFr<> 0 and 
		(	len(convert (varchar(6), UoYyMmFr)) <> 6 or 
			(substring(convert (varchar(6), UoYyMmFr), 1, 4) not between 1900 and 2099) or 
			(substring(convert (varchar(6), UoYyMmFr), 5, 2) not between 1 and 12) 
		)
 	if @row> 0
		begin
  		select @errmsg= 'Year And Month Not Proper'
			goto ErrHndlr
 	end

	select @row= count('x') from inserted where UoMnuCd in ('RepSOPnd') and 
		not exists (select 'x' from param where PTyp ='YN' and UoYN= PMCd) 
	if @row> 0 
	begin
		select @errmsg= 'Show Pending Orders Or All Orders ?'
		/* select @errmsg= 'Show Pending Orders Or All Orders (Enter ''Y'' / ''N'')' */
		goto ErrHndlr
	end

	/* ***** Manali Trading Module ***** */
	/**** Manali 3.7.1 - RepJtPndRm added */
	select @row= count('x') from inserted where UoMnuCd in ('RepJtPnd', 'RepJtOpnOrd', 'RepJtPndRm') and 
		not exists (select 'x' from param where PTyp ='YN' and UoYN= PMCd) 
	if @row> 0 
	begin
		select @errmsg= 'Show Pending Vouchers Or All Vouchers?'
		goto ErrHndlr
	end

	/* ***** Manali Trading Module ***** */
	/***** Manali 3.7.1 - RepJtPndRm added **** */
	select @row= count('x') from inserted where UoMnuCd in ('RepJtPnd', 'RepJtPndRm') and 
		UoOmTc not in (Select distinct TmBaseTcs from TcMst Where TmCoCd=UoCoCdFr)
	if @row> 0 
	begin
		select @errmsg= 'Report Cannot Be Viewed For This TC'
		goto ErrHndlr
	end

	/* ***** Manali Trading Module ***** */
	select @row= count('x') from inserted where UoMnuCd in ('RepJtOpnOrd') and UoOmTc <>'JSO'
	if @row> 0 
	begin
		select @errmsg= 'Report Can Be Viewed Only for ''JSO'''
		goto ErrHndlr
	end

	/* ***** Manali Trading Module ***** */
	select @row= count('x') from inserted where UoMnuCd in ('REPJTNETSTK') 
			and (UoOmTcFr<>'JSO' Or UoOmTcTo<>'JSO')
	if @row> 0 
	begin
		select @errmsg= 'Report Can Be Viewed Only for ''JSO'''
		goto ErrHndlr
	end

	/**** Manali 3.7.1 ****/
	select @row= count('x') from inserted where UoMnuCd in ('RepJtVchAnaDsg', 'RepJtVchAnaRm') 
			and UoCmCtgFr Not In ('C', 'S')
	if @row> 0 
	begin
		select @errmsg= 'Customer Category ''C'' or ''S'' Allowed '
		goto ErrHndlr
	end

	/**** Manali 3.7.1 ****/
	select @row= count('x') from inserted where UoMnuCd in ('RepJtVchAnaDsg') 
			and not exists(Select 'x' from Param Where PTyp='YN' and PMCd=UoYn4)
	if @row> 0 
	begin
		select @errmsg= 'Select to Show Profit Percentage On Cost Or Sales Price'
		goto ErrHndlr
	end

	/**** Manali 3.7.1 ****/
	select @row= count('x') from inserted where UoMnuCd in ('RepJtVchAnaDsg') 
			and not exists(Select 'x' from Param Where PTyp='YN' and PMCd=UoYn3)
	if @row> 0 
	begin
		select @errmsg= 'Select to If Show Picture'
		goto ErrHndlr
	end

	/**** Manali 3.7.1 ****/
	select @row= count('x') from inserted where UoMnuCd in ('RepJtVchAnaDsg', 'RepJtVchAnaRm') 
			and not exists(Select 'x' from Param Where PTyp='YN' and PMCd=UoYn1)
	if @row> 0 
	begin
		select @errmsg= 'Select to If To Add Value for 2nd Tc'
		goto ErrHndlr
	end

	/**** Manali 3.7.1 ****/
	select @row= count('x') from inserted where UoMnuCd in ('RepJtVchAnaDsg', 'RepJtVchAnaRm') 
			and not exists(Select 'x' from Param Where PTyp='YN' and PMCd=UoYn2)
	if @row> 0 
	begin
		select @errmsg= 'Select to If To Add Value for 3rd Tc'
		goto ErrHndlr
	end

	/**** Manali 3.7.1 ****/
	select @row= count('x') from inserted where UoMnuCd in ('RepJtVchAnaDsg', 'RepJtVchAnaRm') 
			and ((UoOmTc=UoOmTc1 and UoOmTc<>'' and UoOmTc1<>'') 
					or (UoOmTc1=UoOmTc2  and UoOmTc1<>'' and UoOmTc2<>'')
					or (UoOmTc=UoOmTc2  and UoOmTc<>'' and UoOmTc2<>''))
	if @row> 0 
	begin
		select @errmsg= 'Cannot Repeat Same Tc in Selection'
		goto ErrHndlr
	end



	/**** Manali 3.7.1 ****/
	select @row= count('x') from inserted where UoMnuCd in ('RepJtVchAnaDsg', 'RepJtVchAnaRm') 
		and (
					(not exists(Select 'x' from TcMst Where TmCoCd=UoCoCdFr and TmTc=UoOmTc and TmCmCtg=UoCmCtgFr ) 
						and UoOmTc<>'' )
									or 
					(not exists(Select 'x' from TcMst Where TmCoCd=UoCoCdFr and TmTc=UoOmTc1 and TmCmCtg=UoCmCtgFr) 
						and UoOmTc1<>'')
									or
					(not exists(Select 'x' from TcMst Where TmCoCd=UoCoCdFr and TmTc=UoOmTc2 and TmCmCtg=UoCmCtgFr) 
						and UoOmTc2<>'')
						)			
	if @row> 0 
	begin
		select @errmsg= 'Tcs Should Belong To Customer Category Selected'
		goto ErrHndlr
	end




/* **** Zubin 211 'RepRmStkReq' added **** */ 
	/* ***** Manali3.5.0 - 05/11/08 - ErrMsg changed for Radio buttons */
	select @row= count('x') from inserted where UoMnuCd in ('RepDetOrd', 'RepOrdCat', 
			'RepAvgOrd', 'RepRmReqDyn', 'RepRmReqLoc', 'RepRmReqStat', 
			'RepDsg', 'RepDsgCrd', 'RepRmStkReq') 
		and not exists (select 'x' from param where PTyp ='YN' and UoYN3= PMCd) 
	if @row> 0 
	Begin
		/* select @errmsg= 'Show Prd Rm Qty & Wt Instead Of Rm Qty & Wt?   (Enter ''Y'' / ''N'')' */
		select @errmsg= 'Select Prd Rm Qty & Wt Or Rm Qty & Wt'
		goto ErrHndlr
	end

	/* **** Manali 3.5.0 - Design Catalogue New Fields For Display/Hide */
	select @row= count('x') from inserted where UoMnuCd in ('RepDsgCat') 
		and not exists (select 'x' from param where PTyp ='YN' and UoYN4= PMCd) 
	if @row> 0 
	Begin
		select @errmsg= 'Show Design Category ?'
		goto ErrHndlr
	end

	select @row= count('x') from inserted where UoMnuCd in ('RepDsgCat') 
		and not exists (select 'x' from param where PTyp ='YN' and UoYN5= PMCd) 
	if @row> 0 
	Begin
		select @errmsg= 'Show Design Size ?'
		goto ErrHndlr
	end

	select @row= count('x') from inserted where UoMnuCd in ('RepDsgCat') 
		and not exists (select 'x' from param where PTyp ='YN' and UoYN6= PMCd) 
	if @row> 0 
	Begin
		select @errmsg= 'Show Design Colour?'
		goto ErrHndlr
	end

	select @row= count('x') from inserted where UoMnuCd in ('RepDsgCat') 
		and not exists (select 'x' from param where PTyp ='YN' and UoYN7= PMCd) 
	if @row> 0 
	Begin
		select @errmsg= 'Show Design Gross Wt?'
		goto ErrHndlr
	end

	select @row= count('x') from inserted where UoMnuCd in ('RepDsgCat') 
		and not exists (select 'x' from param where PTyp ='YN' and UoYN8= PMCd) 
	if @row> 0 
	Begin
		select @errmsg= 'Show Design Prd/Sal Ctg ?'
		goto ErrHndlr
	end


	select @row= count('x') from inserted where UoMnuCd in ('RepDsgCat') 
		and not exists (select 'x' from param where PTyp ='YN' and UoYN9= PMCd) 
	if @row> 0 
	Begin
		select @errmsg= 'Show Rm Sub Category ?'
		goto ErrHndlr
	end

	select @row= count('x') from inserted where UoMnuCd in ('RepDsgCat') 
		and not exists (select 'x' from param where PTyp ='YN' and UoYN10= PMCd) 
	if @row> 0 
	Begin
		select @errmsg=  'Show Rm Weight Details ?'
		goto ErrHndlr
	end

	select @row= count('x') from inserted where UoMnuCd in ('RepDsgCat') 
		and not exists (select 'x' from param where PTyp ='YN' and UoYN11= PMCd) 
	if @row> 0 
	Begin
		select @errmsg= 'Show Rm Quantity Details  ?'
		goto ErrHndlr
	end

	select @row= count('x') from inserted where UoMnuCd in ('RepDsgCat') 
		and not exists (select 'x' from param where PTyp ='YN' and UoYN12= PMCd) 
	if @row> 0 
	Begin
		select @errmsg= 'Show Rm Code Details ?'
		goto ErrHndlr
	end

	select @row= count('x') from inserted where UoMnuCd in ('RepDsgCat') 
		and not exists (select 'x' from param where PTyp ='YN' and UoYN13= PMCd) 
	if @row> 0 
	Begin
		select @errmsg= 'Check To Print Design Catalogue '
		goto ErrHndlr
	end

/*  ***** Manali 3.6.0 - Catlogue entry removed from Design Catalogue Module ***** */
/* 	select @row= count('x') from inserted where UoMnuCd in ('RepDsgCat') 
		and (Select HDsgCatYN from Head where HCoCd=@wSelfCoCd and HCd=@wSelfCd) ='N' 
		and (UoYN13='Y' Or UoDpCatFr<>'' Or UoDpCatTo<>'' Or UoDpCatSel<>'')
	if @row> 0 
	Begin
		select @errmsg= 'Cat Code Scope fields Should be Blank and Print Catalogue Should be Unchecked when ''Design Catalogue Module'' Is ''N'''
		goto ErrHndlr
	end
*/	
	/* **** Manali 3.5.0 - Design Catalogue New Fields For Display/Hide */

	/* **** Manali 3.5.0 - Order and Invoice Catalogue New Fields For Display/Hide */
	/* **** Manali Trading Module - REPJTCAT added below */
	select @row= count('x') from inserted where UoMnuCd in ('RepOrdCat', 'RepInvCat', 'RepJtCat') 
		and not exists (select 'x' from param where PTyp ='YN' and UoYN15= PMCd) 
	if @row> 0 
	Begin
		select @errmsg= 'Show Design Colour ?'
		goto ErrHndlr
	end

	select @row= count('x') from inserted where UoMnuCd in ('RepOrdCat', 'RepInvCat', 'RepJtCat') 
		and not exists (select 'x' from param where PTyp ='YN' and UoYN16= PMCd) 
	if @row> 0 
	Begin
		select @errmsg= 'Show Design Gross Weight ?'
		goto ErrHndlr
	end

	select @row= count('x') from inserted where UoMnuCd in ('RepOrdCat', 'RepInvCat', 'RepJtCat') 
		and not exists (select 'x' from param where PTyp ='YN' and UoYN17= PMCd) 
	if @row> 0 
	Begin
		select @errmsg= 'Show Rm Code Details ?'
		goto ErrHndlr
	end

	select @row= count('x') from inserted where UoMnuCd in ('RepOrdCat', 'RepInvCat', 'RepJtCat') 
		and not exists (select 'x' from param where PTyp ='YN' and UoYN18= PMCd) 
	if @row> 0 
	Begin
		select @errmsg= 'Show Rm Sub Category ?'
		goto ErrHndlr
	end

	select @row= count('x') from inserted where UoMnuCd in ('RepOrdCat', 'RepInvCat', 'RepJtCat') 
		and not exists (select 'x' from param where PTyp ='YN' and UoYN19= PMCd) 
	if @row> 0 
	Begin
		select @errmsg= 'Show Rm Weight Details ?'
		goto ErrHndlr
	end

	select @row= count('x') from inserted where UoMnuCd in ('RepOrdCat', 'RepInvCat', 'RepJtCat') 
		and not exists (select 'x' from param where PTyp ='YN' and UoYN20= PMCd) 
	if @row> 0 
	Begin
		select @errmsg= 'Show Rm Quantity Details  ?'
		goto ErrHndlr
	end

	select @row= count('x') from inserted where UoMnuCd in ('RepOrdCat', 'RepInvCat', 'RepJtCat') 
		and not exists (select 'x' from param where PTyp ='YN' and UoYN21= PMCd) 
	if @row> 0 
	Begin
		select @errmsg= 'Show Unit Price ?'
		goto ErrHndlr
	end

  /*** Manali 3.10.0 - 09/03/12 - Different Suppress options for Unit Pricce and Qty ***/
	select @row= count('x') from inserted where UoMnuCd in ('RepOrdCat', 'RepInvCat', 'RepJtCat') 
		and not exists (select 'x' from param where PTyp ='YN' and UoYN27= PMCd) 
	if @row> 0 
	Begin
		select @errmsg= 'Show Design Qty ?'
		goto ErrHndlr
	end

	select @row= count('x') from inserted where UoMnuCd in ('RepOrdCat', 'RepInvCat', 'RepJtCat') 
		and not exists (select 'x' from param where PTyp ='YN' and UoYN22= PMCd) 
	if @row> 0 
	Begin
		select @errmsg= 'Show Sub PO No ?'
		goto ErrHndlr
	end
	/* **** Manali 3.5.0 - Order and Invoice Catalogue New Fields For Display/Hide */

	
	select @row= count('x') from inserted where UoMnuCd in ('RepInvCat') 
		and UoYn3='Y' and UoYn24='Y'
	if @row> 0 
	Begin
		select @errmsg= 'Show Voucher Sr and Show Distinct Design Options Cannot Be Selected Same Time'
		goto ErrHndlr
	end

	select @row= count('x') from inserted where UoMnuCd in ('RepInvCat', 'RepOrdCat') 
		and not exists (select 'x' from param where PTyp ='YN' and UoYN24= PMCd)
	if @row> 0 
	Begin
		select @errmsg= 'Show Voucher Sr? '
		goto ErrHndlr
	end

	/* Zubin 212 */ 
	select @row= count('x') from inserted where UoMnuCd= 'RepDsg' 
		and not exists (select 'x' from param where PTyp ='YN' and PMCd= UoYN4 and PSCd= '') 
	if @row> 0 
	Begin
		select @errmsg= 'Show Designs With Below ''D'' SCtg Only ?' 
		/* select @errmsg= 'Show Designs With Below ''D'' SCtg Only (Enter ''Y'' / ''N'')'  */
		goto ErrHndlr
	end
	/* Zubin 212 */ 
	
	/*** Manali 3.10.0 - 10/03/12 - Show Designs With Below ''D'' SCtg Only in Design Catalogue ***/
	select @row= count('x') from inserted where UoMnuCd= 'RepDsgCat' 
		and not exists (select 'x' from param where PTyp ='YN' and PMCd= UoYN and PSCd= '') 
	if @row> 0 
	Begin
		select @errmsg= 'Show Designs With Below ''D'' SCtg Only ?' 
		goto ErrHndlr
	end
	

/* **** Zubin 211 **** */
	/* ***** Manali3.5.0 - 05/11/08 - ErrMsg changed for Radio buttons */
	select @row= count('x') from inserted where UoMnuCd in ('RepRmReqStat', 'RepRmReqDyn') and UoYN<> '' and 
		Not Exists (select 'x' from param where PTyp ='YN' and UoYN= PMCd and PSCd= '') 
	if @row> 0 
	Begin
		/* select @errmsg= 'Show Cost Value, Sale Value Or No Value?   (Enter ''Y'' / ''N''/ '''')' */
		select @errmsg= 'Select Cost Value Or Sale Value Or No Value? '
		goto ErrHndlr
	end

	select @row= count('x') from inserted where UoMnuCd in ('RepRmReqStat', 'RepRmReqDyn') and 
		(UoYN3= 'Y' and UoYN= 'N') 
	if @row> 0 
	Begin
		select @errmsg= 'Cannot Check ''Sale Value'' When ''Production Qty/Wt'' Is Selected ' 
		goto ErrHndlr
	end

	select @row= count('x') from inserted where UoMnuCd in ('RepRmReqStat', 'RepRmReqDyn') and 
		(UoYN3= 'N' and UoYN= 'Y') 
	if @row> 0 
	Begin
		select @errmsg= 'Cannot Check ''Cost Value'' When ''Order Qty/Wt'' Is Selected '
		goto ErrHndlr
	end

/* **** Zubin 211 **** */

	select @row= count('x') from inserted where UoMnuCd in ('RepDsgCrd') and 
		not exists (select 'x' from param where PTyp= 'RMGRP' and PMCd= 'DSGCRD' and PSCD= UoRmGrp)
	if @row> 0 
	begin
		select @errmsg= 'Invalid Raw Material Group Option'
		goto ErrHndlr
	end
/* Zubin 212 */ 
	select @row= count('x') from inserted where UoMnuCd in ('RepDsgCrd') and 
		not exists (select 'x' from param where PTyp ='YN' and UoYN5= PMCd and PSCd= '') 
	if @row> 0 
	begin
		select @errmsg= 'Invalid ''Show Design Analysis'' Option (Enter Y/N)'
		goto ErrHndlr
	end
/* Zubin 212 */ 
	/**********Geeta******Emr207*********/
	select @row= count('x') from inserted where UoMnuCd in ('RepDsgCrd') and 
		not exists (select 'x' from param where PTyp ='YN' and UoYN2= PMCd) 
	if @row> 0 
	begin
		select @errmsg= 'Show Design History? '
		/* select @errmsg= 'Show Design History? (Enter Y/N)' */
		goto ErrHndlr
	end
	/**********Geeta******Emr207*********/
	select @row= count('x') from inserted where UoMnuCd in ('RepDsgCrd') and 
		not exists (select 'x' from param where PTyp= 'LABGRP' and PMCd= 'DSGCRD' and PSCD= UoLabGrp)
	if @row> 0 
	begin
		select @errmsg= 'Invalid Labour Group Option'
		goto ErrHndlr
	end

	select @row= count('x') from inserted where UoMnuCd in ('RepPrm', 'RepVarPrm') and 
		not exists (select 'x' from param where PTyp= 'SYSDFN' and PMCd= UoPmCd)
	if @row> 0 
	begin
		select @errmsg= 'Invalid Entry Type'
		goto ErrHndlr
	end

	select @row= count('x') from inserted where UoMnuCd in ('RepCustLst') and 
		not exists (select 'x' from param where PTyp= 'CMCTG' and PMCd= UoPmCd)
	if @row> 0 
	begin
		select @errmsg= 'Invalid Customer Category'
		goto ErrHndlr
	end

	select @row= count('x') from inserted where UoMnuCd in ('RepDsg', 'RepDsgCrd') and 
		not exists (select 'x' from param where PTyp= 'DMTCTYP' and PMCd= UoPmCd)
	if @row> 0 
	begin
		select @errmsg= 'Invalid Design Type'
		goto ErrHndlr
	end

	select @row= count('x') from inserted where UoMnuCd in ('RepOrdPln') and 
		not exists (select 'x' from param where PTyp= 'PRDSTG' and PMCd= UoPmCdFr)
	if @row> 0 
	begin
		select @errmsg= 'Invalid Production Stage'
		goto ErrHndlr
	end

	select @row= count('x') from inserted where UoMnuCd in ('RepOrdPln') and 
		not exists (select 'x' from param where PTyp= 'YN' and PMCd= UoYN)
	if @row> 0 
	begin
		select @errmsg= 'Specify Whether To Show Plan Pcs Instead Of Load ?'
		/* select @errmsg= 'Specify Whether To Show Plan Pcs Instead Of Load (Enter ''Y'' / ''N'')' */
		goto ErrHndlr
	end
	
	/* **** Manali 3.5.0 - 06/11/08 - ErrMsg changed for Radio buttons */
	select @row= count('x') from inserted where UoMnuCd in ('RepBV') and UoYN<> '' and
		not exists (select 'x' from param where PTyp ='YN' and UoYN= PMCd) 
	if @row> 0 
	begin
		/* select @errmsg= '(Enter ''Y'' / ''N''/ Blank)' */
		select @errmsg= 'Select if to show Acknowledged Bags / Unacknowledged bags/ all the Bags'
		goto ErrHndlr
	end

/********** urmi 2.04 ********************/
/*  ######################  Manoj #### Ver: 2.0.6 #### Date: 18/12/2003  ######################  */
/* (Bef 2.06)
	select @row= count('x') from inserted where UoMnuCd in ('RepStkLed', 'RepFlashStk') and
		not exists (select 'x' from param where PTyp ='YN' and UOYN1= PMCd)
	if @row> 0
	begin
		select @errmsg= 'Specify Whether To Show Value or Not (Enter 'Y'/'N')'
		goto ErrHndlr
	end
*/
	select @row= count('x') from inserted where UoMnuCd in ('RepStkLed', 'RepFlashStk') and
			UOPMCDFR <> '' And
			not exists(select 'x' from param where PTyp ='WIPVAL' and UOPMCDFR= PMCd)
	if @row> 0
	begin
		select @errmsg= 'Value Type Option Not Defined'
		goto ErrHndlr
	end
/*
	select @row= count('x') from inserted where UoMnuCd in ('RepStkLed', 'RepFlashStk') and
			UoYN1 = 'Y' and
			not exists(select 'x' from param where PTyp ='WIPVAL' and UOPMCDFR= PMCd)
	if @row> 0
	begin
		select @errmsg= 'Value Type Option Not Defined'
		goto ErrHndlr
	end
	*/
	/*
	select @row= count('x') from inserted where UoMnuCd in ('RepStkLed', 'RepFlashStk') and
			UoYN1 = 'N' and UOPMCDFR<> ''
	if @row> 0
	begin
		select @errmsg= 'Value Type should be Blank'
		goto ErrHndlr
	end
	*/
	/*  #########################################  Manoj  #########################################  */

	/*** Jay 3.4(FG) ***/
	/***** Manali Trading Module *****/
	select @row= count('x') from inserted where UoMnuCd in ('RepStkLed','RepFgStkLed', 'RepJtStkLed') and 
		not exists (select 'x' from param where PTyp ='YN' and UOYN2= PMCd) 
	if @row> 0 
	begin
		select @errmsg= 'Specify Whether To apply Reverse Logic or Not'
		/* select @errmsg= 'Specify Whether To apply Reverse Logic or Not (Enter ''Y''/''N'')' */
		goto ErrHndlr
	end
	/********** urmi 2.04 ********************/

	/* Zubin 213 (Below not req) */
	/*	select @row= count('x') from inserted where UoMnuCd in ('RepCustmStk') and 
			UoYN2 <> '' and not exists (select 'x' from param where PTyp ='YN' and UoYN2= PMCd) 
		if @row> 0 
		begin
 			select @errmsg= 'Show Deemed/Actual Import ?(Enter ''Y'' / ''N''/ Blank)'
			goto ErrHndlr
		end
	*/ 
	/* Zubin 213 */ 

	/* ***** Manali Trading Module ***** */
	/***** Manali 3.7.1 - RepJtPndRm added ***** */
	Select @row= count('x') From Inserted Where UoMnuCd in ('RepCustmStk', 'RepVchPrnPR', 'RepTxnPR', 'RepDetInv', 
		'RepAvgInvCTB', 'RepDetInvCTB', 'RepInvCat', 'RepInvCTBCat', 'RepOrdProfitMrg', 'RepRmReqStat', 'RepAvgInv', 
		'RepRmReqDyn', 'RepRmReqLoc', 'RepInvCustStmt', 'RepSOPnd', 'RepDiaProfit', 'RepPckLstInv', 'RepInvProfitMrg',
		'RepSalAnaDsg', 'RepSalAnaRm', 'RepSalAnaLab', 'RepWipFgBPcs', 'RepJtPnd', 'RepJtOpnOrd', 'RepJtPndRm', 'REPMULTIPRCQT') 
		And UoCurCdFr <> '' And 		
		Not Exists (select 'x' from param where PTyp= 'CURNCY' and PMCd= UoCurCdFr and PSCd= '') 
	If @row> 0 
	Begin
 		select @errmsg= 'Invalid ''Currency'' Option'
		goto ErrHndlr
	End

	/* Manali 3.3.0 Stock Amt in the Purchase Listing cannot be blank */
	/* ***** Manali 3.5.0 - 06/11/08 - ErrMsg changed for radio buttons */	
	select @row= count('x') from inserted where UoMnuCd ='RepTxnPR' and 
		not exists (select 'x' from Param where PTyp ='YN' and UoYN1= PMCd) 
	if @row> 0 
	begin
		/* select @errmsg= 'Enter ''Y'' to display Stk Amt, ''N'' to display Purchase Amt' */
		select @errmsg= 'Select if to display Stock Amt Or Purchase Amt'
		goto ErrHndlr
	end

	/*		
	Select @row= count('x') From Inserted Where UoMnuCd in ('RepSalAnaDsg', 'RepSalAnaRm', 'RepSalAnaLab', 
		 'RepWipFgBPcs') And 
		((Exists (Select 'x' From Head Where HCd=@wSelfCd And HCoCd=@wSelfCoCd And HImpComLoc='N') And UoCurCdFr<> '') Or 		
		 (Exists (Select 'x' From Head Where HCd=@wSelfCd And HCoCd=@wSelfCoCd And HImpComLoc='Y') And UoCurCdFr = '')) And 
		Not Exists (select 'x' from param where PTyp= 'CURNCY' and PMCd= UoCurCdFr and PSCd= '') 
	If @row> 0 
	Begin
 		Select @errmsg= 'Invalid ''Currency'' Option'
		GoTo ErrHndlr
	End
	*/

	/*** Sachin, Jen (2.14 Next) ***/
	Select @row= count(distinct HBaseCurCd) From Inserted, Head 
		Where UoMnuCd in ('RepSalAnaDsg', 'RepSalAnaRm', 'RepSalAnaLab', 'RepWipFgBPcs') and UoCurCdFr= '' 
		and HCoCd>= UoCoCdFr and (UoCoCdTo= '' or HCoCd<= UoCoCdTo)  
		and (UoCoCdSel= '' or UoCoCdSel like '%'''+ HCoCd + '''%')
	If @row> 1 
	Begin
 		Select @errmsg= '''Currency'' Cannot Be Blank if Base Currency is different for Companies in the Scope'
		GoTo ErrHndlr
	End
	/*** Sachin, Jen (2.14 Next) ***/

	/*	select @row= count('x') from inserted where UoMnuCd in ('RepInvBOELst') and 
			Not Exists (select 'x' from param where PTyp= 'YN' and PMCd= UoYn1 and PSCd= '') 
		if @row> 0 
		begin
 			select @errmsg= 'Invalid ''In Base Currency'' Option'
			goto ErrHndlr
		end
	*/

	/* Zubin 213 */ 
	/*** Jay 3.4(FG) ***/
	/* ***** Manali Trading Module ***** */
	/* ***** Manali 3.8.0 - RepPckLstInvBag added ****/
	select @row= count('x') from inserted where UoMnuCd in ('RepInv','RepFgInv', 'RepJtInv', 'RepPckLstInvBag') and 
        	not exists (select 'x' from param where PTyp ='YN' and UoYN1= PMCd) 
	if @row> 0 
	begin
		select @errmsg= 'Show Customer PO No, Customer Design?'
		/* select @errmsg= 'Show Customer PO No, Customer Design? (Enter ''Y'' / ''N'')' */
		goto ErrHndlr
	end


	/* ***** Manali 3.8.0 - RepPckLstInvBag added ****/
	select @row= count('x') from inserted where UoMnuCd in ('RepPckLstInvBag') and 
        	not exists (select 'x' from param where PTyp ='YN' and UoYN9= PMCd) 
	if @row> 0 
	begin
		select @errmsg= 'Show Bag Details?'
		goto ErrHndlr
	end

	/*** Jay 3.4(FG) ***/ 
	/* ***** Manali Trading module ***** */
	select @row= count('x') from inserted where UoMnuCd in ('RepDetInv', 'RepDetInvCTB', 
			'RepAvgInv', 'RepAvgInvCTB', 'RepDetOrd', 'RepAvgOrd','REPFGDET','REPFGAVG', 'REPJTDET','REPJTAVG') and 
		not exists (select 'x' from param where PTyp ='YN' and UOYN2= PMCd) 
	if @row> 0 
	begin
		select @errmsg= 'Show Main Value? ' 
		/* select @errmsg= 'Show Main Value? (Enter ''Y'' or ''N'')' */
		goto ErrHndlr
	end
		
	select @row= count('x') from inserted where 
	  UoMnuCd in ('RepDetOrd', 'RepAvgOrd','REPFGDET','REPFGAVG', 'REPJTDET','REPJTAVG', 'REPMULTIPRCQT') and 
		not exists (select 'x' from param where PTyp ='YN' and UOYN25= PMCd) 
	if @row> 0 
	begin
		select @errmsg= 'Show Run Sr? ' 
		goto ErrHndlr
	end
		
	/* *** JAY 2.14Next *** */
	/*** Jay 3.4(FG) ***/
	/* ***** Manali Trading module ***** */ 
	select @row= count('x') from inserted where UoMnuCd in ('RepDetInv', 'RepDetInvCTB', 
			'RepAvgInv', 'RepAvgInvCTB', 'RepDetOrd', 'RepAvgOrd','REPFGDET','REPFGAVG', 'REPJTDET','REPJTAVG', 'REPMULTIPRCQT') and 
		not exists (select 'x' from param where PTyp ='YN' and UOYN14= PMCd) 
	if @row> 0 
	begin
		select @errmsg= 'Show Rm Value? '
		/* select @errmsg= 'Show Rm Value? (Enter ''Y'' or ''N'')' */
		goto ErrHndlr
	end


/* Zubin 212 (Before 011205) */ 
/*	select @row= count('x') from inserted where UoMnuCd in ('RepDetInv', 'RepInvCustStmt', 'RepDetInvCTB', 
			'RepAvgInv', 'RepAvgInvCTB', 'RepDetOrd', 'RepAvgOrd') and 
		UOYN4<> '' And UOYN2 = 'Y' 
	if @row> 0 
	begin
		select @errmsg= 'Labour Value Option Should Be Blank'
		goto ErrHndlr
	end 
*/ 
/* Zubin 212 */ 
	/*** Jay 3.4(FG) ***/ 
	/* ***** Manali Trading module ***** */
	select @row= count('x') from inserted where UoMnuCd in ('RepDetInv', 'RepDetInvCTB', 'RepAvgInv', 
			'RepAvgInvCTB', 'RepInvCustStmt', 'RepDetOrd', 'RepAvgOrd','REPFGDET','REPFGAVG', 'REPJTDET','REPJTAVG', 'REPMULTIPRCQT') and 
		not exists (select 'x' from param where PTyp ='YN' and UOYN4= PMCd) 
	if @row> 0 
	begin
		select @errmsg= 'Show Labour Value? '
		/* select @errmsg= 'Show Labour Value? (Enter ''Y'' or ''N'')' */
		goto ErrHndlr

	end
	
	/*** Jay 3.4(FG) ***/
	/* ***** Manali Trading module ***** */
	select @row= count('x') from inserted where UoMnuCd in ('RepDetInv', 'RepDetInvCTB', 
		'RepAvgInv', 'RepInvCustStmt', 'RepAvgInvCTB', 'RepDetOrd', 'RepAvgOrd','REPFGDET','REPFGAVG','REPJTDET','REPJTAVG', 'REPMULTIPRCQT') and 
		not exists (select 'x' from param where PTyp ='YN' and UOYN4= PMCd) and UOYN2 = 'N'
	if @row> 0 
	begin
		select @errmsg= 'Show Labour Value? '
		/* select @errmsg= 'Show Labour Value? (Enter ''Y'' or ''N'')' */
		goto ErrHndlr
	end

	/* ***** Manali 3.6.0 - Use Level above 5 cant view values */
	/* ***** Manali Trading module ***** */
	select @row= count('x') from inserted where UoMnuCd in ('RepDetOrd', 'RepAvgOrd','REPFGDET','REPFGAVG', 'REPJTDET','REPJTAVG',
	'REPMULTIPRCQT') and 
		(UOYN2='Y' or UOYN14='Y' Or UOYN4= 'Y') and (Select PNum from Param Where Ptyp='USR' and PMCd=UoUsrCd)>5
	if @row> 0 
	begin
		select @errmsg= 'User Cannot View Rm/Lab/All Values'
		goto ErrHndlr
	end

	select @row= count('x') from inserted where UoMnuCd in ('RepOrdCat') and 
		UOYN21='Y'  and (Select PNum from Param Where Ptyp='USR' and PMCd=UoUsrCd)>5
	if @row> 0 
	begin
		select @errmsg= 'User Cannot View Price'
		goto ErrHndlr
	end
	/* ***** Manali 3.6.0 - Use Level above 5 cant view values */

	/* ***** Manali 3.6.0 - Excel Report ***** */
	/* ***** Manali Trading Module ***** */
	select @row= count('x') from inserted where UOYN23 = 'Y' and
		((UoMnuCd in ('RepDetInv', 'RepDetInvCTB') and (UoInTcFr='' Or UoInTcTo='' or UoInYyFr='' Or UoInYyTo='' or UoInChrFr='' Or UoInChrTo='' or UoInNoFr=0 Or UoInNoTo=0)) Or
		(UoMnuCd In ('RepDetOrd', 'RepJtDet') and (UoOmTcFr='' Or UoOmTcTo='' or UoOmYyFr='' Or UoOmYyTo='' or UoOmChrFr='' Or UoOmChrTo='' or UoOmNoFr=0 Or UoOmNoTo=0)))
	if @row> 0 
	begin
		select @errmsg= '''Tc, Yy, Chr, No'' Cannot be blank when Excel Format is Selected'
		goto ErrHndlr
	end

	select @row= count('x') from inserted where UOYN23 = 'Y' and 
		((UoMnuCd in ('RepDetInv', 'RepDetInvCTB', 'RepJtDet') and (UoInTcFr<>UoInTcTo or UoInYyFr<>UoInYyTo or UoInChrFr<>UoInChrTo or UoInNoFr<>UoInNoTo)) Or
		(UoMnuCd In ('RepDetOrd', 'RepJtDet') and (UoOmTcFr<>UoOmTcTo or UoOmYyFr<>UoOmYyTo or UoOmChrFr<>UoOmChrTo or UoOmNoFr<>UoOmNoTo)))
	if @row> 0 
	begin
		select @errmsg= 'From and To Voucher Scope Should be same when Excel Format is Selected'
		goto ErrHndlr
	end

	select @row= count('x') from inserted where UoMnuCd in ('RepDetOrd', 'RepDetInv', 'RepJtDet') and 
		UOYN23 = 'Y' and Not(UoRmGrp='SR' And UoLabgrp='SR')
		
	if @row> 0 
	begin
		select @errmsg= 'Rm and Labour Group Should be ''SR'' same when Excel Format is Selected'
		goto ErrHndlr
	end

	select @row= count('x') from inserted where UoMnuCd = 'RepDetOrd' and 
		UOYN23 = 'Y' and UoYn12='Y'		
	if @row> 0 
	begin
		select @errmsg= 'Show Components Option Cannot Be Selected when Excel Format is Selected'
		goto ErrHndlr
	end

	select @row= count('x') from inserted where UoMnuCd in ('RepDetOrd', 'RepDetInv', 'RepDetInvCTB', 'RepJtDet') and 
		UOYN23 = 'Y' and Not(UOYN2='Y' And UOYN14='Y' And UOYN4= 'Y')
	if @row> 0 
	begin
		select @errmsg= 'Show Value Options Sholud Be Selected when Excel Format is Selected'
		goto ErrHndlr
	end
	/* ***** Manali 3.6.0 - Excel Report ***** */

	/* Zubin 213 */
	select @row= count('x') from inserted where UoMnuCd in ('RepInvCustStmt') and 
		not exists (select 'x' from param where PTyp ='YN' and UOYN9= PMCd and PSCd= '') 
	if @row> 0 
	begin
		select @errmsg= '''In Base Currency'' ?'
		/* select @errmsg= '''In Base Currency'' (Enter ''Y'' or ''N'')' */ 
		goto ErrHndlr
	end
	/* Zubin 213 */
	select @row= count('x') from inserted where UoMnuCd in ('RepInvCustStmt') and UoInVal= 'S' and UOYN9<> 'Y' 
	if @row> 0 
	begin
		select @errmsg= '''In Base currency'' Should Be Checked When Value As Is ''S'' '
		/* select @errmsg= '''In Base currency'' Should Be ''Y'' When Value As Is ''S'' ' */
		goto ErrHndlr
	end

	/* Zubin 213 */
/*	select @row= count('x') from inserted where UoMnuCd in ('RepGrFormBack', 'RepGSPCert', 'RepDeclnLet') and 
		not exists (select 'x' from param where PTyp ='YN' and UOYN1= PMCd and PSCd= '') 
	if @row> 0 
	begin
		select @errmsg= '''In Base currency'' (Enter ''Y'' or ''N'')'
		goto ErrHndlr
	end
*/

	/* ***** Manali 3.5.0 - 06/11/08 - ErrMsg changed for Radio Buttons  ***** */
	/***** Manali 3.8.0 - 'RepNetStk' added *****/
	select @row= count('x') from inserted where UoMnuCd in ('RepStkLed', 'RepFlashStk', 'RepNetStk') and 
		not exists (select 'x' from param where PTyp ='YN' and UOYN3= PMCd) 
	if @row> 0 
	begin
		/* select @errmsg= 'Specify Whether To Show Quantity/Weight (Enter ''Y'' for Qty, ''N'' for Wt)' */
		select @errmsg= 'Select Whether To Show Quantity/Weight'
		goto ErrHndlr
	end

	/*** (Jen 2.13) ***/
	select @row= count('x') from inserted where UoMnuCd in ('RepStkLed') and 
		not exists (select 'x' from param where PTyp ='YN' and UOYN= PMCd) 
	if @row> 0 
	begin
		select @errmsg= 'Specify Whether To Show Quantity/Wt Or Value Or blank in Last Column'
		/* select @errmsg= 'Specify Whether To Show Quantity If Value Is Not Selected (Enter ''Y'' to Show Qty, ''N'' To Hide The Column)' */
		goto ErrHndlr
	end
	/*** (Jen 2.13) ***/

/**** Shilpa Emr 206 **/
	/****** added RepWipBRmBkDt (Jen 2.14) ******/
	select @row= count('x') from inserted where UoMnuCd in ('RepWipBRm', 'RepWipFgBRm', 'RepWipBRmBkDt', 'RepBvRm') and 
		not exists (select 'x' from param where PTyp= 'WIPVAL' and UoPMCdFr= PMCd) and UoPmCdFr <> ''
	if @row> 0 
	Begin
		select @errmsg= 'Show Stock Value Or Cost Value Or Sale Value Or No Value?'
		goto ErrHndlr
	end

	select @row= count('x') from inserted where UoMnuCd in ('RepWipBRm') and 
		not exists (select 'x' from param where PTyp= 'WIPVAL' and UoPMCdTo= PMCd) and UoPmCdTo <> ''
	if @row> 0 
	Begin
		select @errmsg= 'Show Stock Value Or Cost Value Or Sale Value Or No Value?'
		goto ErrHndlr
	end

	select @row= count('x') from inserted where UoMnuCd in ('RepDsgCat') and 
		not exists (select 'x' from param where PTyp= 'DMTcTyp' and PMCd= Inserted.UoPmCd)
	if @row> 0 
	begin
		select @errmsg= 'Invalid Design Type'
		goto ErrHndlr
	end

	/* ****** Manali 3.5.0 - 05/11/08 - ErrMsg changed for Radio buttons **** */
	select @row= count('x') from inserted where UoMnuCd in ('RepDsgCat') and 
		not exists (select 'x' from param where PTyp= 'YN' and PMCd= Inserted.UoYN3)
	if @row> 0 
	begin
		 /* select @errmsg= 'Invalid Show Rm Option' */
		select @errmsg= 'Select if to show Prouction Rm Qty and Wt Or Sales Rm Qty and Wt'
		goto ErrHndlr
	end
	/* ****** Manali 3.5.0 - 05/11/08 - ErrMsg changed for Radio buttons **** */

	/*********************Geeta*******Emr206************10/04/04***/

	select @row= count('x') from inserted ins where UoMnuCd in ('RepDsgCat') and 
		not exists (select 'x' from param where PTyp= 'YN' and PMCd= Ins.UoYN1)
	if @row> 0 
	begin
		select @errmsg= 'Enter Design On Hold ?'
		/* select @errmsg= 'Enter Design On Hold (Y/N)?' */
		goto ErrHndlr
	end

	/* **** Manali 3.6.0 - 19/09/09 -  For RepDsgCat, UoYN2 replaced by UoDsgExc so following changed commented  **** */	
	/* ****** Manali 3.5.0 - 05/11/08 - ErrMsg changed for Radio buttons **** */
/*	select @row= count('x') from inserted ins where UoMnuCd in ('RepDsgCat') and 
	 ins.UoYn2<>'' and not exists (select 'x' from param where PTyp= 'YN' and PMCd= Ins.UoYN2)
	if @row> 0 
	begin
		 /* select @errmsg= 'Enter Free Design /Region Exclusive/Not Exclusive (''/Y/N)?' */
		select @errmsg= 'Select if to Show Free Design /Region Exclusive/Not Exclusive'
		goto ErrHndlr
	end
	*/
	
	/* **** Manali 3.6.0 - 19/09/09 - For RepDsgCat, UoYN2 replaced by UoDsgExc **** */
	select @row= count('x') from inserted ins where UoMnuCd in ('RepDsgCat') 
	 and not exists (select 'x' from param where PTyp= 'DSGEXC' and PMCd= Ins.UoDsgExc)
	if @row> 0 
	begin
		select @errmsg= 'Select if to Show All Design/Free Design /Region Exclusive/Not Exclusive'
		goto ErrHndlr
	end

	select @row= count('x') from inserted ins where UoMnuCd in ('RepDsgCat') and 
	 /*ins.UoYn2=''*/ Ins.UoDsgExc In ('F', 'A') and (ins.UoCmCdFr<>'' or ins.UoCmCdTo<>''  or ins.UoCmCdSel<>'' )
	if @row> 0 
	begin
		select @errmsg= 'Customer Code should be blank for Free/All Design'
		goto ErrHndlr
	end

	select @row= count('x') from inserted ins where UoMnuCd in ('RepDsgCat') and 
	 /*ins.UoYn2<>''*/ Ins.UoDsgExc In ('C', 'R') and (ins.UoCmCdFr='' and ins.UoCmCdTo=''  and ins.UoCmCdSel='' )
	if @row> 0 
	begin
		select @errmsg= 'Customer Code Cannot be blank for Region Exclusive/Not Exclusive Design'
		goto ErrHndlr
	end

	/****************************Geeta**********************************/	
	/*********Geeta***********Emr207**************/
	select @row= count('x') from inserted where UoMnuCd in ('RepFgBagLed') and 
		not exists (select 'x' from param where PTyp ='YN' and UOYN= PMCd) 
	if @row> 0 
	begin
		select @errmsg= 'Show Box Details? '
		/* select @errmsg= 'Show Box Details? (Enter ''Y'' or ''N'')' */
		goto ErrHndlr
	end

	select @row= count('x') from inserted where UoMnuCd in ('RepFgBagLed') and 
		UOPMCdFr<> '' And UOYN = 'N' 
	if @row> 0 
	begin
		select @errmsg= 'Box Sort Option  Should Be Blank'
		goto ErrHndlr
	end

	select @row= count('x') from inserted where UoMnuCd in ('RepFgBagLed') and UOYN = 'Y'
	and not exists (select 'x' from param where PTyp ='SRTFGBAG' and UOPMCdFr= PMCd) 
	if @row> 0 
	begin
		select @errmsg= 'Enter Box Sort Option '
		goto ErrHndlr
	end
	/*********Geeta***********Emr207**************/
	select @row= count('x') from inserted where UoMnuCd in ('RepDiaProfit') and 
		not exists (select 'x' from param where PTyp= 'InVal' and PMCd= Inserted.UoInVal)
	if @row> 0 
	begin
		select @errmsg= 'Invalid Base Value Option'
		goto ErrHndlr
	end
	select @row= count('x') from inserted where UoMnuCd in ('RepDiaProfit') and 
		not exists (select 'x' from param where PTyp= 'INVAL' and PMCd= Inserted.UoInVal1)
	if @row> 0 
	begin
		select @errmsg= 'Invalid Other Value Option'
		goto ErrHndlr
	end

	/* ***** Manali 3.5.0 - 07/11/08 - ErrMsg changed for Radio Buttons */
	select @row= count('x') from inserted where UoMnuCd in ('RepDetOrd', 'RepAvgOrd') and 
		not exists (select 'x' from param where PTyp ='YN' and UOYN5= PMCd) 
	if @row> 0 
	begin
		/* select @errmsg= 'Specify Whether To Show Order Delivery Date/Production Delivery Date (Enter ''Y'' for Order, ''N'' for Production)' */
		select @errmsg= 'Select Whether To Show Order Delivery Date/Production Delivery Date '
		goto ErrHndlr
	end

	select @row= count('x') from inserted where UoMnuCd in ('RepInvCat', 'RepInvCTBCat') and 
		not exists (select 'x' from param where PTyp ='YN' and UOYN5= PMCd) 
	if @row> 0 
	begin
		select @errmsg= 'Sort Option ?'
		/* select @errmsg= 'Invalid Sort Option' */
		goto ErrHndlr
	end

/*	Manoj 2.10.0	*/
	select @row= count('x') from inserted where UoMnuCd in ('RepInvCat') and 
		not exists (select 'x' from param where PTyp ='YN' and UOYN3= PMCd) 
	if @row> 0 
	begin
		select @errmsg= 'Distinct Design for Scope ?'
		/* select @errmsg= 'Invalid Value of Distinct Design for Scope' */
		goto ErrHndlr
	end
/*	Manoj 2.10.0	*/

	select @row= count('x') from inserted where UoMnuCd in ('RepDiaProfit') and 
		not exists (select 'x' from param where PTyp= 'RMCTG' and UORMCTGFR = PmCd)
	if @row> 0 
	begin
		select @errmsg= 'Invalid Raw Material Category'
		goto ErrHndlr
	end

	/* **** Manali 3.5.0 - 12/11/08 - ErrMsg changed for Radio Buttons ***** */
	select @row= count('x') from inserted where UoMnuCd in ('RepWIWPrd','RepWIWExp','RepWIWLongExp','RepWIWLongPrd') 
		and not exists (select 'x' from param where PTyp ='YN' and UoYN3= PMCd) 
	if @row> 0 
	Begin
		/* select @errmsg= 'Show Exp or Prd Del Date? (Enter ''Y'' / ''N'')'	*/
		select @errmsg= 'Select Exp or Prd Del Date'	
		goto ErrHndlr
	end


	/*********geeta*******Emr206*************/	
	/* **** Manali 3.5.0 - 05/11/08 - Errmsg changed for Radio buttons */
	select @row= count('x') from inserted where UoMnuCd = 'RepSOPnd' and UoYN3<>'' 
		and not exists (select 'x' from param where PTyp ='YN' and UoYN3= PMCd) 
	if @row> 0 
	Begin
		/* select @errmsg= 'Show Exp or Prd or Both Del Dates? (Enter ''Y'' / ''N''/'''')' */
		select @errmsg= 'Select if to Show Exp or Prd or Both Del Dates '
		goto ErrHndlr
	end
	/* **** Manali 3.5.0 - 05/11/08 - Errmsg changed for Radio buttons */
	/*********geeta*******Emr206*************/	

	/* ***** Manali Trading Module ***** */
/*	select @row= count('x') from inserted where UoMnuCd In ('RepJtPnd', 'RepJtOpnOrd') and UoYN3<>'Y'
	if @row> 0 
	Begin
		select @errmsg= 'Select if to Show Exp Date'
		goto ErrHndlr
	end
*/

/*** Shilpa Emr 206 ***/	
/**** Shilpa Emr 2.04 *****/

	select @row= count('x') from inserted where UoMnuCd in ('RepFlashStk') and 
		UoOpnYyMM > UoYyMmFr
	if @row> 0 
	begin
		select @errmsg= 'Open Year Month Should be Less than From Year Month'
		goto ErrHndlr
	end

	/*** Jay 3.4(FG) ***/
	/***** Manali Trading Module *****/
	select @row= count('x') from inserted where UoMnuCd in ('RepStkLed', 'RepStkSumm','RepFgStkLed', 'RepJtStkLed') and 
		UoOpnYYMm  > convert(int,convert(varchar(6),UoDtFr,112))

	if @row> 0 
	begin
		select @errmsg= 'Open Year Month Should be Less than From Voucher Date'
		goto ErrHndlr
	end
	
	/* select @row= count('x') from inserted where UoMnuCd in ('RepDetInv', 'RepAvgInv') and UoRmGrp = 'RMRSZRt'
	
	if @row> 0 
	begin
		select @errmsg= 'Range Size Option Not Available With Invoice Reports'
		goto ErrHndlr
	end*/
/**** Shilpa Emr 2.04 *****/

	/*** (Jen 2.11) ***/
	select @row= count('x') from inserted where UoMnuCd in ('RepNetStk') and 
		not exists (select 'x' from param where PTyp ='YN' and UoYN= PMCd) 
	if @row> 0 
	begin
		select @errmsg= 'Show Loc And Order Detail?  '
		/* select @errmsg= 'Show Loc And Order Detail?  (Enter ''Y'' / ''N'')' */
		goto ErrHndlr
	end
	/*** (Jen 2.11) ***/

	/* ****** Sachin 2.12 ****** */
	select @row= count('x') from inserted where UoMnuCd in ('RepNetStk') and 
		not exists (select 'x' from param where PTyp ='YN' and UoYN1= PMCd) 
	if @row> 0 
	begin
		select @errmsg= 'Show All?  (Enter ''Y'' / ''N'')'
		goto ErrHndlr
	end
	/* ****** Sachin 2.12 ****** */

	/* *** Jay 2.13(CT) *** */ 
	select @row= count('x') from inserted where UoMnuCd = 'RepDetOrd' and 
		not exists (select 'x' from param where PTyp ='YN' and UOYN12= PMCd) 
	if @row> 0 
	begin
		select @errmsg= 'Specify Whether To Show Component Code '
		/* select @errmsg= 'Specify Whether To Show Component Code - Enter (Yes/No)' */
		goto ErrHndlr
	end
	/* *** Jay 2.13(CT) *** */ 

	/* *** Jay 2.13(CT) *** */ 
	select @row= count('x') from inserted where UoMnuCd = 'RepDsgCrd' and 
		not exists (select 'x' from param where PTyp ='YN' and UOYN6= PMCd) 
	if @row> 0 
	begin
		select @errmsg= 'Specify Whether To Show Component Code '
		/* select @errmsg= 'Specify Whether To Show Component Details - Enter (Yes/No)'*/
		goto ErrHndlr
	end
	/* *** Jay 2.13(CT) *** */ 

	/* ***** Manali 3.5.0 - 03/11/08 - Err Msg Changed for Radio buttons ***** */
	/* *** Jay 2.14 *** */ 	
	/* **** Manali 3.8.0 - Commented ****/
	/*Select @row= count('x') from inserted ins where UoMnuCd iN ('RepPrd') and 
			Not exists (Select 'x' from Param where PTyp= 'YN' and PMCd= ins.UOYN3)
	if @row> 0 	
	Begin
		/* select @errmsg= 'Enter ''Y'' for Production Points and ''N'' for Rupees Value' */
		select @errmsg= 'Select if To Show eihter Production Points or Rupees Value '
		goto ErrHndlr
	end*/
	/* *** Jay 2.14 *** */ 	
	/* ***** Manali 3.5.0 - 03/11/08 - Err Msg Changed for Radio buttons ***** */

	/* *** Jay 2.14 *** */ 	
	select @row= count('x') from inserted where UoMnuCd= 'RepStkLed' and 
		not exists (select 'x' from param where PTyp ='YN' and UOYN4= PMCd) 
	if @row> 0 
	begin
		select @errmsg= 'Specify Whether To Show Description '
		/* select @errmsg= 'Specify Whether To Show Description (Enter Yes/No)' */
		goto ErrHndlr
	end
	/* *** Jay 2.14 *** */ 	
	
	/* *** Jay 2.14 *** */ 	
	/*** Jay 3.4(FG) ***/
	/* ***** Manali Trading module ***** */
	select @row= count('x') from inserted where UoMnuCd in ('RepDetOrd', 'RepAvgOrd', 'RepDetInv', 
								'RepDetInvCTB', 'RepAvgInv', 'RepAvgInvCTB','REPFGDET','REPFGAVG','REPJTDET','REPJTAVG') and 
		Not Exists (select 'x' from param where PTyp ='YN' and PMCd= UoYN13 and PSCd= '') 
	if @row> 0 
	begin
		select @errmsg= 'Show Group Combination Description In Report ?'
		/* select @errmsg= 'Show Group Combination Description In Report(Yes/ No)? (Enter ''Y'' / ''N'')' */
		goto ErrHndlr
	end
	/* *** Jay 2.14 *** */ 	

	/* *** Jay 2.14 *** */ 	
	select @row= count('x') from inserted where UoMnuCd in ('RepWipFgBPcs') and 
		(UoPmCdFr<> '' and Not Exists (Select 'x' from param where PTyp ='ORDVAL' and PMCd= UoPmCdFr and PSCd= ''))
	if @row> 0 
	begin
		select @errmsg= 'Value Type Option Is Invalid (Enter ''C'' for Cost, ''S'' for Sale or Blank for None)'
		goto ErrHndlr
	end
	/* *** Jay 2.14 *** */ 	

	/*** (Jen 3.01) ***/
	/* ***** Manali 3.5.0 - 06/11/08 - ErrMsg changed for Radio buttons   */
	select @row= count('x') from inserted where UoMnuCd in ('RepLoc') and 
		Not Exists (select 'x' from param where PTyp ='YN' and PMCd= UoYN and PSCd= '') 
	if @row> 0 
	begin
		/* select @errmsg= 'Show Parent Location Or Other Location In Report(Yes/ No)? (Enter ''Y'' / ''N'')'  */
		select @errmsg= 'Select if to show Parent Location Or Other Location In Report'
		goto ErrHndlr
	end
	/*** (Jen 3.01) ***/

	/* *** Jay 3.01 *** [RepRtCmp] */
	select @row= count('x') from inserted where UoMnuCd in ('RepRtCmp') and 
		not exists (select 'x' from param where PTyp ='RMCTG' and UoRmCtgFr= PMCD and UoRmCtgFr in ('C','D')) 
	if @row> 0 
	begin
		select @errmsg= 'Rm Catagory Should Be In ''C'' Or ''D'''
		goto ErrHndlr
	end
	/* *** Jay 3.01 *** [RepRtCmp] */

	/* *** Jay 3.01 *** [RepRtCmp] */
	select @row= Count('x') from inserted where UoMnuCd in ('RepRtCmp') and 
		(Not Exists (select 'x' from CustMst where CmCtg= UoCmCtgFr and CmCd= UoCmCdFr) Or  
		Not Exists (select 'x' from CustMst where CmCtg= UoCmCtgTo and CmCd= UoCmCdTo))
	if @row> 0 
	begin
		select @errmsg= 'Invalid Customer /Source Customer Code'
		goto ErrHndlr
	end
	/* *** Jay 3.01 *** [RepRtCmp] */
	
	/* *** Jay 3.01 *** [RepRtCmp] */
	select @row= count('x') from inserted where UoMnuCd in ('RepRtCmp') and 
			UoCmCtgFr= UoCmCtgTo and UoCmCdFr= UoCmCdTo and UoYyMmFr= 0 and UoCmCdFr<>'ZSelf'
	if @row> 0 
	begin
		select @errmsg= 'Customer And Source Customer Can Not Be Same When Source YyyyMm Is Blank'
		goto ErrHndlr
	end
	/* *** Jay 3.01 *** [RepRtCmp] */

	/* *** Jay 3.01 *** [RepRtCmp] */
	/* ***** Manali 3.5.0 - 06/11/08 - ErrMsg chaged for radio buttons ***** */
	select @row= count('x') from inserted where UoMnuCd in ('RepRtCmp') and 
		not exists (select 'x' from param where PTyp ='YN' and UoYn= PMCd) 
	if @row> 0 
	begin
		/* select @errmsg= 'Show Sales Prc/Cost Prc ? (Enter ''N'' / ''Y'')' */
		select @errmsg= 'Select Sales Prc Or Cost Prc '
		goto ErrHndlr
	end
	/* *** Jay 3.01 *** [RepRtCmp] */

	/* *** Jay 3.01 *** [RepRtCmp] */
	/* **** Manali 3.3.0 - 'ZVALUE' included below */
	/* ***** Manali 3.5.0 - 06/11/08 - ErrMsg chaged for radio buttons ***** */
	select @row= count('x') from inserted where UoMnuCd in ('RepRtCmp') and UoYn= 'Y' and UoCmCdTo not in ('ZSelf', 'ZValue')
	if @row> 0 
	begin
		/* select @errmsg= 'CostRt Y/N Should Be ''N'' When Source Customer Is Not ZSELF or ZVALUE' */
		select @errmsg= 'Cost Rt cannot be selected When Source Customer Is Not ZSELF or ZVALUE'
		goto ErrHndlr
	end
	/* *** Jay 3.01 *** [RepRtCmp] */

	/* *** Jay 3.01 *** [RepRtCmp] */
	/* **** Manali 3.3.0 - 'ZVALUE' included below */
	/* ***** Manali 3.5.0 - 06/11/08 - ErrMsg chaged for radio buttons ***** */
	select @row= count('x') from inserted where UoMnuCd in ('RepRtCmp') and 
		UoCmCtgFr= UoCmCtgTo and UoCmCdFr= UoCmCdTo and UoYyMmFr= 0 and UoYn<> 'Y' and UoCmCdTo in ('ZSelf', 'ZValue') 
	if @row> 0 
	begin
		select @errmsg= 'CostRt Y/N Should Be ''Y'' When Source YyyyMm = 0 and Source Customer = ''ZSELF'' or ''ZVALUE'''
		goto ErrHndlr
	end
	/* *** Jay 3.01 *** [RepRtCmp] */

	
	/* *** Manali 3.03 *** [RepRtCmp] */
	select @row= count('x') from inserted where UoMnuCd in ('RepRtCmp') and 
		not exists (select 'x' from param where PTyp ='YN' and UoYn1= PMCd) 
	if @row> 0 
	begin
		select @errmsg= 'Select If To Show Sales Prc/Cost Prc '
		/* select @errmsg= 'Show Sales Prc/Cost Prc ? (Enter ''N'' / ''Y'')' */
		goto ErrHndlr
	end
	/* *** Manali 3.03 *** [RepRtCmp] */

	/* *** Manali 3.03 *** [RepRtCmp] */
	select @row= count('x') from inserted where UoMnuCd in ('RepRtCmp') and UoYn1= 'Y' and UoCmCdFr not in ('ZSELF', 'ZVALUE')
	if @row> 0 
	begin
		select @errmsg= 'CostRt Should Not Be Selected When Customer Is Not ''ZSELF'' or ''ZVALUE'''
		/* select @errmsg= 'CostRt Y/N Should Be ''N'' When Customer Is Not ''ZSELF'' or ''ZVALUE'''  */
		goto ErrHndlr
	end
	/* *** Manali 3.03 *** [RepRtCmp] */

	/* *** Jay 3.01 *** [RepTxnCustmPurLst] */
	/* ***** Manali 3.5.0 - 06/11/08 - Err Msg changed for Radio Buttons */
	select @row= count('x') from inserted where UoMnuCd in ('RepTxnCustmPurLst') and UoYn <> '' and 
		not exists (select 'x' from param where PTyp= 'YN' and PMCd= UoYn) 
	if @row> 0 
	begin
		/* select @errmsg= 'Invalid Show Closed Purchase (Yes/ No)' */
		select @errmsg= 'Select if to Show Closed / Open / Both Purchases'
		goto ErrHndlr
	end
	/* *** Jay 3.01 *** [RepRtCmp] */


	/* *** Jay 3.1.0_1 *** [Derived Value] */
	/* ***** Manali 3.5.0 - 03/11/08 - Err Msg changed for Radio Buttons */
	select @row= count('x') from inserted  where UoMnuCd in ('RepCustmStk', 'RepCustmStkLand') And 
		not exists (select 'x' from param where PTyp ='YN' and UoYN3= PMCd) 
	if @row> 0 
	begin
		/* select @errmsg= 'Show Derived Value / Actual Value (Enter ''Y'' / ''N'' )' */
		select @errmsg= 'Select either Derived Value Or Actual Value '
		goto ErrHndlr
	end
	/* *** Jay 3.1.0_1 *** [Derived Value] */
	/* *** Jay 3.1.0_1 *** [Derived Value] */
	/* ***** Manali 3.5.0 - 03/11/08 - Err Msg changed for Radio Buttons */


	/* ***** Manali 3.5.0 - 03/11/08 - Err Msg changed for Radio Buttons */
	/* **** Manali 3.03 - RepCustmStk - Derived Rate */
	select @row= count('x') from inserted where UoMnuCd in ('RepCustmStk', 'RepCustmStkLand') And 
		not exists (select 'x' from param where PTyp ='YN' and UoYN6= PMCd) 
	if @row> 0 
	begin
		/* select @errmsg= 'Show Derived Rate / Actual Rate (Enter ''Y'' / ''N'' )' */
		select @errmsg= 'Select either Derived Rate Or Actual Rate'
		goto ErrHndlr
	end
	/* **** Manali 3.03 - RepCustmStk - Derived Rate */
	/* ***** Manali 3.5.0 - 03/11/08 - Err Msg changed for Radio Buttons */

	select @row= count('x') from inserted where UoMnuCd in ('RepPckLst', 'RepCustmStkLand') And 
		not exists (select 'x' from param where PTyp ='YN' and UoYN1= PMCd) 
	if @row> 0 
	begin
		select @errmsg= 'Show Summary ?'
		/* select @errmsg= 'Show Summary (Enter ''Y'' / ''N'' )' */
		goto ErrHndlr
	end
	/* *** Jay 3.1.0_1 *** [Derived Value] */

/* Zubin 3.01.01 (260308) */
	select @row= count('x') from inserted where UoMnuCd = 'RepTravExcel' and 
		not exists (select 'x' from OrdMst where OmCoCd =UoCoCdFr and OmTc=UoOmTcFr and OmYy =UoOmYyFr and 
					OmChr =UoOmChrFr and OmNo =UoOmNoFr) 
	if @row> 0 
	begin
		select @errmsg= 'Quotation Number Does Not Exist'
		goto ErrHndlr
	end

	select @row= count('x') from inserted where UoMnuCd = 'RepTravExcel' and UoPSCdSel = '' 
	if @row> 0 
	begin
		select @errmsg= 'Diamond Grade Cannot Be Blank'
		goto ErrHndlr
	end

/* Zubin 3.01.01 (260308) */


	/* ***** Manali 3.5.0 - YN fields added in Dsg Cat Rep to hide and Display fields in reports */	
	 /*Select @row = count('x') from UsrOpt where UaMnuCd='REPDSGCAT' */
	/* ***** Manali 3.5.0 - YN fields added in Dsg Cat Rep to hide and Display fields in reports */	

	if @SYS= 'XEMR' 
	BEGIN
		select @row= count('x') from inserted where UoMnuCd in ('RepRmStkLed', 
			'RepRmFlashStk', 'RepDsgStkLed', 'RepDsgFlashStk', 'RepVchDet', 
			'RepVchAvg', 'RepVchCat', 'RepTxnPnd', 'RepProfitMargin', 
			'RepAnaDsg', 'RepAnaRm', 'RepAnaLab') and 			
			not exists (select 'x' from xHead where xHcd = UoxCoCdFr) 
		if @row> 0 
		begin
			select @errmsg= 'Company Code Not Defined'
			goto ErrHndlr
		end

		select @row= count('x') from inserted where UoMnuCd in 
			('RepRmStkLed', 'RepRmFlashStk') and 
			not exists (select 'x' from param where PTyp= 'LOCTYP' and PMCd= UoLocTypFr)
		if @row> 0 
		begin
			select @errmsg= 'Invalid Location Type'
			goto ErrHndlr
		end

		select @row= count('x') from inserted where UoMnuCd in 
			('RepRmStkLed', 'RepRmFlashStk') and 
			not exists (select 'x' from param where PTyp ='RMCTG' and 
			UoRmCtgFr= PMCd) 
		if @row> 0 
		begin
			select @errmsg= 'Invalid Raw Material Category'
			goto ErrHndlr
		end

		select @row= count('x') from inserted where UoMnuCd in ('RepRmStkLed') and 
			not exists (select 'x' from param where PTyp ='VchOpt' and 
			UoVchOpt= PMCd) 
		if @row> 0 
		begin
			select @errmsg= 'Invalid Option for Type Of Detail (Enter VCH, VCHSR or VCHSRNO)'
			goto ErrHndlr
		end

		/* **** Zubin 211 **** */ 
		select @row= count('x') from inserted where UoMnuCd in ('RepBvRm') and UoDtTo< UoDtFr 
		if @row> 0 
		begin
			select @errmsg= 'To Date Should Be Greater Than Or Equal To From Date'
			goto ErrHndlr
		end
		/* **** Zubin 211 **** */ 

		select @row= count('x') from inserted where UoMnuCd in (
			'RepRmStkLed', 'RepDsgStkLed', 'RepVchDet', 'RepVchAvg', 'RepVchCat', 
			'RepTxnPnd', 'RepProfitMargin', 'RepAnaDsg', 'RepAnaRm', 'RepAnaLab') and 
			UoDtTo< UoDtFr 
		if @row> 0 
		begin
			select @errmsg= 'To Date Should Be Greater Than From Date'
			goto ErrHndlr
		end

		select @row= count('x') from inserted where UoMnuCd in ('RepVchDet', 
			'RepVchAvg', 'RepVchCat', 'RepTxnPnd', 'RepProfitMargin','RepAnaDsg', 
			'RepAnaRm', 'RepAnaLab') and 
			UoDocDtTo< UoDocDtFr 
		if @row> 0 
		begin
			select @errmsg= 'To Doc Date Should Be Greater Than From Doc Date'
			goto ErrHndlr
		end

		select @row= count('x') from inserted where UoMnuCd in (
			'RepRmFlashStk', 'RepDsgFlashStk') and UoYyMmTo< UoYyMmFr 
		if @row> 0 
		begin
			select @errmsg= 'To YyyyMm Should Be Greater Than From YyyyMm'
			goto ErrHndlr
		end

		select @row= count('x') from inserted where UoMnuCd in ('RepRmFlashStk', 
			'RepDsgFlashStk', 'RepTxnPnd') and 
			not exists (select 'x' from param where PTyp ='YN' and UoYN= PMCd) 
		if @row> 0 
		begin
			select @errmsg= 'Show The Break Up Detail Box ?'
			/* select @errmsg= 'Show The Break Up Detail Box Yes/No?   (Enter ''Y'' / ''N'')' */
			goto ErrHndlr
		end

		select @row= count('x') from inserted where UoMnuCd in ('RepRmFlashStk', 
			'RepTxnPnd') and 
			not exists (select 'x' from param where PTyp ='ORDVAL' and UoOrdVal= PMCd) 
		if @row> 0 
		begin
			select @errmsg= 'Show Cost Or Sales Rate? (''Y'' for Cost/''N'' For Sale)'
			goto ErrHndlr
		end

		select @row= count('x') from inserted where UoMnuCd in ('RepTxnPnd') and 
			not exists (select 'x' from param where PTyp ='YN' and UoYN2= PMCd) 
		if @row> 0 
		begin
			select @errmsg= 'Show Only Pending/Matched Records ?'
			/* select @errmsg= 'Show Only Pending/Matched Records(''Y'' to show Only Pending, ''N'' to show all)' */
			goto ErrHndlr
		end

		select @row= count('x') from inserted where UoMnuCd in ('RepTxnPnd', 'RepVchDet', 
			'RepVchAvg', 'RepVchCat', 'RepProfitMargin', 'RepAnaDsg', 'RepAnaRm',
			'RepAnaLab') and 
			not exists (select 'x' from xparam where xPTyp ='CurTyp' and 
			UoCurTyp= xPMCd) 
		if @row> 0 
		begin
			select @errmsg= 'Select Txn Currency/ Base Currency ? (Enter ''T'' for ''Txn'' or ''B'' for ''Base'')'
			goto ErrHndlr
		end

		select @row= count('x') from inserted where UoMnuCd in ('RepProfitMargin') 
			and not exists (select 'x' from param where PTyp ='OrdVal' and 
			UoOrdVal= PMCd) 
		if @row> 0 
		begin
			select @errmsg= 'Select Cost(C)/Sales(S) as Base Value'
			goto ErrHndlr
		end

		select @row= count('x') from inserted where UoMnuCd in ('RepProfitMargin') 
			and not exists (select 'x' from param where PTyp ='OrdVal' and 
			UoOrdVal1= PMCd) 
		if @row> 0 
		begin
			select @errmsg= 'Select Cost(C)/Sales(S) as Other Value'
			goto ErrHndlr
		end

		select @row= count('x') from inserted where UoMnuCd in ('RepTxnPnd', 'RepVchDet', 
			'RepVchAvg', 'RepVchCat', 'RepProfitMargin', 'RepAnaDsg', 'RepAnaRm', 'RepVchPrnPR', 'RepTxnPR', 
			'RepAnaLab') and 
			(UoCurTyp = 'B' And UoCurCdFr<>'') or (UoCurTyp = 'T' and
			not exists (select 'x' from param where PTyp ='Curncy' and UoCurCdFr= PMCd)) 
		if @row> 0 
		begin
			select @errmsg= 'Enter Valid Currency Code'
			goto ErrHndlr
		end

		select @row= count('x') from inserted where 
		  UoMnuCd in ('RepVchPrnPR', 'RepVchPrnCNV', 'RepVchPrnMV', 'RepVchPrnREC') 
			and not exists (select 'x' from param where PTyp ='YN' and UoYN1= PMCd) 
		if @row> 0 
		begin
			select @errmsg= 'Select if To Show Description'
			goto ErrHndlr
		end


		select @row= count('x') from inserted where UoMnuCd in ('RepTxnPnd') and 
			not exists (select 'x' from xparam where xPTyp ='MatchTyp' and UoMatchTyp= xPMCd) 
		if @row> 0 
		begin
			select @errmsg= 'Select Valid MatchType (Enter ''T'' for ''Txn'' and ''M'' for ''Memo'')'
			goto ErrHndlr
		end

		select @row= count('x') from inserted where UoMnuCd in ('RepVchDet', 'RepVchAvg', 'RepVchCat') and 
			not exists (select 'x' from xTcTyp where xTpTcTyp= UoxTcTypFr) 
		if @row> 0 
		begin
			select @errmsg= 'InValid Tc Type'
			goto ErrHndlr
		end

/*** Bef(2.05)	select @row= count('x') from inserted where UoMnuCd in ('RepAnaDsg', 
			'RepAnaRm', 'RepAnaLab') and 
			not exists (select 'x' from xTcTyp where xTpVchTyp= UoVchTypFr)  
		if @row> 0 
		begin
			select @errmsg= 'InValid Voucher Type'
			goto ErrHndlr
		end
***/

/*** Zubin xEmr2.05 ********************************************************************/
		select @row= count('x') from inserted where (UoMnuCd in ('RepAnaRm', 
			'RepAnaLab') and 
			not exists (select 'x' from xTcTyp where xTpVchTyp= UoVchTypFr) )
			OR UoMnuCd in ('RepAnaDsg') and UoVchTypFr <> '' and 
			not exists (select 'x' from xTcTyp where xTpVchTyp= UoVchTypFr)
		if @row> 0 
		begin
			select @errmsg= 'Invalid Voucher Type'
			goto ErrHndlr
		end

		Select @Row = count('x') from Inserted 
			where UoMnuCd = 'REPANADSG' and UoxRepGrp <> ''
				and Not Exists (Select 'x' from xParam where xPTyp = 'REPGRP'
				and xPmCd = UoxRepGrp)
		If @Row > 0 
		Begin
			Select @ErrMsg = 'Invalid Report Summary Group Option'
			Goto ErrHndlr
		End
/*** Zubin xEmr2.05 ********************************************************************/

		select @row= count('x') from inserted where UoMnuCd in ('RepAnaDsg') and 
			not exists (select 'x' from param where PTyp= 'DMTCTYP' and 
			UoDmTypFr= PMCd) 
		if @row> 0 
		begin
			select @errmsg= 'Invalid Design Type'
			goto ErrHndlr
		end

		select @row= count('x') from inserted where UoMnuCd in ('RepAnaDsg') and 
			not exists (select 'x' from param where PTyp ='YN' and UoYN= PMCd) 
		if @row> 0 
		begin
			select @errmsg= 'Show Val/Wt ? (''Y'' for Value/''N'' For Weight)' 
			goto ErrHndlr
		end

		select @row= count('x') from inserted where UoMnuCd in ('RepAnaDsg') and 
			UoYN= 'Y' and not exists (select 'x' from param 
						where PTyp ='OrdVal' and UoOrdVal= PMCd) 
		if @row> 0 
		begin
			select @errmsg= 'Show Cost/Sales Value ?(Enter ''C'' or ''S'')'
			goto ErrHndlr
		end

/*** (Bef 2.05)	select @row= count('x') from inserted where UoMnuCd in ('RepAnaDsg') and 
			UoYN= 'N' and UoOrdVal<> ''  
		if @row> 0 
		begin
			select @errmsg= 'Cost/Sales Value Field Should Be Blank'
			goto ErrHndlr
		end
***/

/*** Zubin xEmr2.05 *****************************************************************/
		select @row= count('x') from inserted where UoMnuCd in ('RepAnaDsg') and 
			UoYN= 'N' and (UoOrdVal <> '' Or UoxDc <> '')
		if @row> 0 
		begin
			select @errmsg= 'Cost/Sales Value and Show Positive Val As (D/C) Fields Should Be Blank'
			goto ErrHndlr
		end
/*** Zubin xEmr2.05 *****************************************************************/

		select @row= count('x') from inserted where UoMnuCd in ('RepAnaRm', 'RepAnaLab') and 
			not exists (select 'x' from param where PTyp ='OrdVal' and 
			UoOrdVal= PMCd) 
		if @row> 0 
		begin
			select @errmsg= 'Show Cost/Sales Value ?(Enter ''C'' or ''S'')'
			goto ErrHndlr
		end

		select @row= count('x') from inserted where UoMnuCd in ('RepAnaLab') and 
			not exists (select 'x' from param where PTyp= 'SALABGR' and 
			UoSaLabGrp= PMCd) 
		if @row> 0 
		begin
			select @errmsg= 'Invalid Labour Grouping Option'
			goto ErrHndlr
		end

		select @row= count('x') from inserted where UoMnuCd in ('RepAnaRm') and 
			not exists (select 'x' from param where PTyp= 'SARMGR' and 
			UoSaRmGrp= PMCd) 
		if @row> 0 
		begin
			select @errmsg= 'Invalid RM Grouping Option'
			goto ErrHndlr
		end

		select @row= count('x') from inserted where UoMnuCd in ('RepTxnPnd') and 
			not exists (select 'x' from xTcTyp where xTpTcTyp= UoxTcTypFr and 
			(xTpTxnBM = 'B' or xTpMemBM = 'B')) 
		if @row> 0 
		begin
			select @errmsg= 'InValid Voucher(Tc) Type(Select Only TxnBase/MemoBase TcTyp)'
			goto ErrHndlr
		end

		/*('Select xTpVchTyp from xTcTyp where xTpCoCd= '' + _
              	adc('UoCoCdFr') + '' and xTpTcTyp= '' + adc('UoTcTypFr') + ''')*/

		select @row= count('x') from inserted where UoMnuCd in ('RepProfitMargin') 
			and not exists (select 'x' from xTcTyp where xTpTcTyp= UoxTcTypFr and 
			xTpVchTyp not in('PUR', 'PORD', 'MVM') and xTpVchTyp in ('SAL', 'SORD')) 
		if @row> 0 
		begin
			select @errmsg= 'InValid Voucher(Tc) Type (Select Only ''SAL''/''SORD'' TcTypes)'
			goto ErrHndlr
		end

		select @row= count('x') from inserted where UoMnuCd in ('RepVchDet', 'RepVchAvg') and 
			not exists (select 'x' from param where PTyp ='YN' and UoYN= PMCd) 
		if @row> 0 
		begin
			select @errmsg= 'Show Picture ? '
			/* select @errmsg= 'Show Picture ? (Enter ''Y'' / ''N'')' */
			goto ErrHndlr
		end

		select @row= count('x') from inserted where UoMnuCd in ('RepVchDet', 'RepVchAvg') and 
			not exists (select 'x' from param where PTyp ='ORDVAL' and UoOrdVal= PMCd) 
		if @row> 0 
		begin
			select @errmsg= 'Show Sales Prc/Cost Prc ? (Enter ''Y'' / ''N'')'
			goto ErrHndlr
		end
		select @row= count('x') from inserted where UoMnuCd in ('RepVchDet', 'RepVchAvg') and 
			not exists (select 'x' from param where PTyp ='YN' and UoYN4= PMCd) 
		if @row> 0 
		begin
			select @errmsg= 'Show Unit Val/Total Val ? (Enter ''Y'' / ''N'')'
			goto ErrHndlr
		end

		select @row= count('x') from inserted where UoMnuCd in ('RepVchDet', 'RepVchAvg') and 
			UoYN1<> '' and not exists (select 'x' from param where PTyp ='YN' and UoYN1= PMCd) 
		if @row> 0 
		begin
			select @errmsg= 'Entered + Amendment Qty ? (Enter ''Y'' / ''N'')'
			goto ErrHndlr
		end

		select @row= count('x') from inserted where UoMnuCd in ('RepVchAvg') and 
			not exists (select 'x' from param where PTyp ='YN' and UoYN2= PMCd) 
		if @row> 0 
		begin
			select @errmsg= 'Show Value ? (Enter ''Y'' / ''N'')'
			goto ErrHndlr
		end

		select @row= count('x') from inserted where UoMnuCd in ('RepVchDet') and 
			not exists (select 'x' from xparam where xPTyp= 'RmGrp' and 'OrdDet'= xPMCd and UoRmGrp= xPSCd) 
		if @row> 0 
		begin
			select @errmsg= 'Invalid RmGrp Option '
			goto ErrHndlr
		end

		select @row= count('x') from inserted where UoMnuCd in ('RepVchDet') and 
			not exists (select 'x' from xparam where xPTyp= 'LabGrp' and 'OrdDet'= xPMCd and UoLabGrp= xPSCd) 
		if @row> 0 
		begin
			select @errmsg= 'Invalid LabGrp Option '
			goto ErrHndlr
		end

    /*** Manali 3.9.0 - UoRepSbFile ***/
		select @row= count('x') from inserted where UoMnuCd in ('RepShpBillFile') and 
			UoSbICC='' or UoSbPOD='' or UoSbCOD='' or UoSbAF=''
		if @row> 0 
		begin
			select @errmsg= 'Fields Cannot be Blank'
			goto ErrHndlr
		end


/*** Zubin xEmr2.05 *****************************************************************/
		Select @Row= count('x') from Inserted 
			where (UoMnuCd = 'REPANADSG' And UoYN= 'Y' 
				and UoxRepGrp <> '') 
				and Not Exists (Select 'x' from Param where PTyp = 'DC' 
					and PmCd = UoxDc)
		If @Row> 0 
		Begin
			Select @ErrMsg = 'Invalid Positive Value As Option'
			Goto ErrHndlr
		End

		Select @Row= count('x') from Inserted 
			where UoMnuCd = 'REPANADSG' and UoxRepGrp = '' and UoxDc <> '' 
		If @Row> 0 
		Begin
			Select @ErrMsg= ' ''Positive Value As'' Option Should Be Blank '
			Goto ErrHndlr
		End
/*** Zubin xEmr2.05 *****************************************************************/

	END

	return
ErrHndlr:
	begin
		select @errmsg='|'+@errmsg
		raiserror(@errmsg,16,1)
		rollback transaction 
	end
end

GOGO

SET TRANSACTION ISOLATION LEVEL READ COMMITTED

GOGO

/************************ Add New Columns  *******************************************************************************/
if not exists(select * from sys.columns 
            where Name = N'OdDmStkYy' and Object_ID = Object_ID(N'OrdDsg'))
	Alter Table OrdDsg add OdDmStkYy varchar(2) not null default ''

if not exists(select * from sys.columns 
            where Name = N'OdDmStkChr' and Object_ID = Object_ID(N'OrdDsg'))
	Alter Table OrdDsg add OdDmStkChr varchar(3) not null default ''
 
if not exists(select * from sys.columns 
            where Name = N'OdDmStkNo' and Object_ID = Object_ID(N'OrdDsg'))
	Alter Table OrdDsg add OdDmStkNo int not null default 0
 
if not exists(select * from sys.columns 
            where Name = N'HjStkNumjFrBagYN' and Object_ID = Object_ID(N'Head'))
		Alter Table Head add HjStkNumjFrBagYN char(1) not null default 'N'

if not exists(select * from sys.columns 
            where Name = N'hjStkDefChr' and Object_ID = Object_ID(N'Head'))
	Alter table Head add hjStkDefChr varchar(3) not null default ''


if not exists(select * from sys.columns 
            where Name = N'DDmStkYy' and Object_ID = Object_ID(N'DmYyMm'))
	Alter Table dmyymm add ddmstkyy varchar(2) not null default ''

if not exists(select * from sys.columns 
            where Name = N'DDmStkChr' and Object_ID = Object_ID(N'DmYyMm'))
	Alter table DmYyMm add ddmstkchr varchar(3) not null default ''

if not exists(select * from sys.columns 
            where Name = N'DDmStkNo' and Object_ID = Object_ID(N'DmYyMm'))
	Alter table DmYyMm add ddmstkno int not null default 0

/************************* Trigger Change ****************************************************************************/
SET TRANSACTION ISOLATION LEVEL READ UNCOMMITTED


/****** Object:  Trigger [dbo].[OrdDsg_INUPDL]    Script Date: 6/9/2015 5:35:02 PM ******/
GOGO
/****** Object:  Trigger dbo.OrdDsg_INUPDL    Script Date: 26/02/01 10:00:34 AM ******/    

ALTER TRIGGER [dbo].[OrdDsg_INUPDL] ON [dbo].[OrdDsg]     
   FOR INSERT,UPDATE,DELETE     NOT FOR REPLICATION
As    
BEGIN    
 /* Trigger Version 3.9.1  */    
 DECLARE @ROW INT    
 DECLARE @mode char(1)    
 DECLARE @errmsg varchar(255)    
 Declare @wSelf VarChar(8)    
 Declare @wSysDb Varchar(3)    
 DECLARE @ReplAgt BIT    
 DECLARE @wCurrPrtn VarChar(1)    
 /*** Jay 3.3(FG) ***/  
 Declare @wInFob Float  
 Declare @wInFobRs Float  
 Declare @wInLclOctroiRs Float  
 Declare @wInLclVatRs Float  
 Declare @wInLclCstRs Float  
 /*** Jay 3.3(FG) ***/     
 /*** Manali - UniDesign ****/  
 DECLARE @wTcTyp varchar(3)   
 /*** Manali - UniDesign ****/  
   
 IF @@ROWCOUNT=0   RETURN    
 If Exists(Select 'x' From TrgDisa Where TdTblNm = 'OrdDsg' )    
  RETURN    
    
 select @mode= case (select count('x') from inserted) when 0 then 'D'    
        else    
    case (select count('x') from deleted) when 0 then 'A' else 'C' end    
        end     
    
 Select @wSysDb=SysDb From SysDb where SysId='CRM'     
 Select @wCurrPrtn = 'C'       
 Set @wSysDb=IsNull(@wSysDb,'')     
    
 /* **** Checking Whether a Normal Transaction or Replicated One **** */    
 Set @ReplAgt=Convert(BIT,SessionProperty('Replication_Agent'))    
 If @ReplAgt = 1         Return
 

Select @wSelf='ZSELF'     
    
 /******* Sachin 3.02.0 [30/11/07] - Purge Partition (Speed Optimization) */    
 if @mode='C' And @ReplAgt<>1   
 Begin    
  Select @row=count('x') From Inserted Where OdPrtKey<>@wCurrPrtn    
    if @row> 0     
  Begin    
     select @errmsg= 'Cannot Edit Records of Previous Partition'    
     goto ErrHndlr    
  End   
  
  Select @row=count('x') From Deleted del  
      Where (Select PValue From Param Where PTyp='TCTYP' and PMCd=del.OdTc)<>'PL' and   
        Exists(Select 'x' From Inserted ins Where ins.OdIdNo=del.OdIdNo and ins.OdCoCd=del.OdCoCd   
            and ins.OdTc=del.OdTc and ins.OdYy=del.OdYy and ins.OdChr=del.OdChr   
            and ins.OdNo=del.OdNo and ins.OdSr=del.OdSr and ins.OdPrtKey<>del.OdPrtKey)   
  if @row> 0   
  begin  
   select @errmsg= 'Cannot Edit Partition Key Field'  
   goto ErrHndlr  
  end  
   
 End    
   
 /**** Manali - (UniDesign Prob) ****/  
 /******** Only One TcTyp records can be Inserted at a time.********/  
 /* Current Partition */   
 select @Row= Count(distinct(PValue)) from inserted Join param On  
  PTyp= 'TC' and PMCd= OdTc and PScd= '' Where OdPrtKey=@wCurrPrtn   
 If @Row> 1  
 begin  
  Select @ErrMsg= 'Transactions of only One TcType can be Acted On'  
  GoTo ErrHndlr  
 end  
  
 /**** Manali - (UniDesign Prob) ****/  
 /******** Only One TcTyp records can be Deleted at a time. *********/  
 /* Current Partition */  
 select @Row= Count(distinct(PValue)) from deleted Join param on  
  PTyp= 'TC' and PMcd= OdTc and PSCd='' Where OdPrtKey=@wCurrPrtn  
 If @Row> 1  
 begin  
  Select @ErrMsg= 'Transactions of only One TcType can be Acted on'  
  GoTo ErrHndlr  
 end  
  
 /*** Manali 3.8.2 Storing Tc Type In The Variable @wTcTyp ***/  
 if @mode in ('A','C')  
 begin  
  select @wTcTyp= max(PValue) from inserted join Param on PTyp= 'TC' and PMCd= OdTc and PSCd= ''  
 end  
  
 if @mode in ('D')  
 begin  
  select @wTcTyp=max(PValue) from deleted join Param on PTyp= 'TC' and PMCd= OdTc and PSCd= ''   
 end  
  
   
 if @mode='A'  And @ReplAgt<>1    
 Begin    
  Select @row=count('x') From Inserted Where OdPrtKey <> @wCurrPrtn     
  if @row> 0     
  begin    
   select @errmsg= 'Cannot Insert Partition Key Other Than ''C'''    
   goto ErrHndlr    
  end    
 end     
    
 If @Mode='D'  And @ReplAgt<>1    
 Begin     
  select @row= count('x') from deleted del where del.OdPrtKey<>@wCurrPrtn    
  if @row> 0     
  begin    
   select @errmsg= 'Cannot Delete Records From Previous Partition'    
   goto ErrHndlr    
  end    
 End    
 /******* Sachin 3.02.0 [30/11/07] - Purge Partition (Speed Optimization) */    
    
    If (@ReplAgt<>1)    
  Begin    
   /*db='WEB' then cannot  Insert /Update/Delete  where OmChr<> 'WEB'     
   Cannot edit/Delete where OmSubmitYN=’Y’ i.e Check deleted.OmSubmitYN=’Y’      
   Cannot edit OmAckYN      
   */    
  /*******No Need to Check  Sysdb='' beceuse Chr='WEB' insert is not allowed in vParam where not exists SysId='CRM'*************************/    
  If @Mode in ('A','C' )     
  Begin      
   /******Cannot Insert/Update/Delete  if Crm database /Emr database with TrdCocd not in sdCoCd ****/    
   /***if SysId=Emr Record not found then normal behaviour i.e can insert/update/delete****/    
   Select @ROW =Count('x')  From inserted where      
   Exists (Select 'x' From SysDb where (SysId='EMR' and ',' + SdCoCd + ',' Not  like '%,'+ OdCoCd +',%'))    
   If @Row > 0      
   Begin     
    select @errmsg= 'Cannot Insert /Update '     
    goto ErrHndlr     
   end     
       
   /*****If  db= '-' then Insert  Not allowed in OrdDsg  for OmChr='WEB' if OmAckYN<>'Y' ******/    
   /* Current Partition */    
   Select @Row= Count('x') From Inserted where OdPrtKey=@wCurrPrtn and OdChr= 'WEB' and @wSysDb= '-' and     
    (Select OmAckYN From OrdMst where OmIdNo=OdOmIdNo and OmCoCd=OdCoCd and OmTc=OdTc and OmYy=OdYy and OmChr=OdChr and OmNo=OdNo and OmPrtKey=OdPrtKey)= 'N'      
   If @Row> 0    
   Begin      
    select @errmsg= 'Cannot Insert/Update Order Design of an UnAcknowledged Web Orders'      
    goto ErrHndlr      
   end      
     
   /*****If  db= 'WEB' then Insert/Update  Not allowed in OrdMst  for OdChr<>'WEB'/OdChr='WEB' and OmSubmitYN='Y' ******/    
   /* Current Partition */    
   Select @Row= Count('x') From Inserted where OdPrtKey=@wCurrPrtn and @wSysDb= 'WEB' and     
    (OdChr<> 'WEB' Or (OdChr= 'WEB' and (Select OmSubmitYN From OrdMst where OmIdNo=OdOmIdNo and OmCoCd=OdCoCd and OmTc=OdTc and     
        OmYy=OdYy and OmChr=OdChr and OmNo=OdNo and OmPrtKey=OdPrtKey)= 'Y'))     
   If @Row> 0    
   Begin      
    select @errmsg= 'Cannot Insert/Update Orders Other than Chr=''WEB'' and Submitted=''N'' '      
    goto ErrHndlr   
   End      
 /*** Jay 3.3(FG) ***/  
 select @Row= Count('x') from inserted join InvHd on InCoCd= OdCoCd and InTc= OdTc and   
   InYy= OdYy and InChr= OdChr and InNo= OdNo where InCls= 'Y' and OdTc='FOT'  
 If @Row > 0  
 begin   
  Select @ErrMsg= 'Cannot Add/Change As Invoice Is Closed'  
  GoTo ErrHndlr  
 end     
 End    
     
  If @Mode ='D'     
  Begin    
   /* Current Partition */    
   select @ROW =Count('x')  From Deleted where  OdPrtKey=@wCurrPrtn and    
    Exists (select 'x' From SysDb where (SysId='EMR' and ',' + SdCoCd + ',' Not like '%,'+ OdCoCd +',%'))    
   if @Row > 0      
   Begin     
    select @errmsg= 'Cannot Delete '     
    goto ErrHndlr     
   end     
    
   /*****If  db= '-' then Insert  Not allowed in OrdDsg  for OmChr='WEB' if OmAckYN<>'Y' ******/    
   /* Current Partition */    
   Select @Row=Count('x') From Deleted where OdPrtKey=@wCurrPrtn and OdChr='WEB' and     
   (Select OmAckYN From OrdMst where OmIdNo=OdOmIdNo and OmCoCd=OdCoCd and OmTc=OdTc and OmYy=OdYy and OmChr=OdChr and OmNo=OdNo and OmPrtKey=OdPrtKey)='N'    
    and @wSysDb='-'    
   If @Row> 0    
   Begin      
    select @errmsg= 'Cannot Delete Order Design of an UnAcknowledged Web Orders'      
    goto ErrHndlr      
   end      
     
   /*****If  db= 'WEB' then Insert/Update  Not allowed in OrdMst  for OdChr<>'WEB'/OdChr='WEB' and OmSubmitYN='Y' ******/    
   /* Current Partition */    
   Select @Row=Count('x') From Deleted where OdPrtKey=@wCurrPrtn and @wSysDb='WEB' and     
   (OdChr<>'WEB' Or (OdChr='WEB' and (Select OmSubmitYN From OrdMst where OmIdNo=OdOmIdNo and OmCoCd=OdCoCd and OmTc=OdTc and     
        OmYy=OdYy and OmChr=OdChr and OmNo=OdNo and OmPrtKey=OdPrtKey)='Y'))     
   If @Row> 0    
   Begin      
    select @errmsg= 'Cannot Delete Ord Designs from Orders Other than Chr=''WEB'' and Submitted=''N'' '      
    goto ErrHndlr      
   end      
    /*** Jay 3.3(FG) ***/  
 select @Row= Count('x') from deleted join InvHd on InCoCd= OdCoCd and InTc= OdTc and   
   InYy= OdYy and InChr= OdChr and InNo= OdNo where InCls= 'Y' and OdTc='FOT'  
 If @Row > 0  
 begin   
  Select @ErrMsg= 'Cannot Delete As Invoice Is Closed'  
  GoTo ErrHndlr  
 end      
   /****** Whenever the OrdDsg record is deleted, delete records from the child tables (OrdRm, OrdLab, OrdAmd, OrdTgt) ******/    
   /* *** Jay 2.13(CT) ***  OrdCT Table Added */    
   /* Sachin 3.02 - Current Partition - Id Nos added */    
   Delete OrdRm From deleted, OrdRm where OdIdNo=OrOdIdNo and OdCoCd = OrCoCd and OdTc = OrTc and     
    OdYy = OrYy and OdChr = OrChr and OdNo = OrNo and OdSr= OrSr and OdPrtKey=OrPrtKey    
     
   Delete OrdLab From deleted, OrdLab where OdIdNo=OlOdIdNo and OdCoCd = OlCoCd and OdTc = OlTc and     
    OdYy = OlYy and OdChr = OlChr and OdNo = OlNo and OdSr= OlSr and OdPrtKey=OlPrtKey    
     
   Delete OrdAmd From deleted, OrdAmd where OdIdNo=OaOdIdNo and OdCoCd = OaCoCd and OdTc = OaTc and     
    OdYy = OaYy and OdChr = OaChr and OdNo = OaNo and OdSr= OaSr and OdPrtKey=OaPrtKey    
     
   Delete OrdTgt From deleted, OrdTgt where OdCoCd = OtCoCd and OdTc = OtTc and     
    OdYy = OtYy and OdChr = OtChr and OdNo = OtNo and OdSr = OtSr and OdPrtKey=OtPrtKey    
     
   Delete OrdPln From deleted, OrdPln where OdCoCd = OpCoCd and OdTc = OpTc and     
    OdYy = OpYy and OdChr = OpChr and OdNo = OpNo and OdSr = OpSr     
    
   Delete OrdCT From deleted, OrdCT where OdIdNo=OctOdIdNo and OdCoCd = OctCoCd and OdTc = OctTc and     
    OdYy = OctYy and OdChr = OctChr and OdNo = OctNo and OdSr = OctSr and OdPrtKey=OctPrtKey    
    End      
 End    
    
/* If Not Exists(Select 'x' From Inserted Join Head On OdCoCd=HCoCd )    
   and  Not Exists(Select 'x' From deleted Join Head On OdCoCd=HCoCd )     
 Return    
*/    
  /******** Key Cannot Be Changed ********/     
  if @mode = 'C'    
  Begin      
  /*** (should thi check be kept??? was there in 2.13)    
   ******** Key Cannot Be Changed when bag exist ********     
  select @Row= Count('x') from deleted del where     
    exists(Select 'x' from Bag where BCoCd= del.OdCoCd and BOdTc= del.OdTc     
    and BOdYy= del.OdYy and BOdChr= del.OdChr and BOdNo= del.OdNo and BOdSr= del.OdSr)     
  If @Row > 0    
  begin     
   Select @ErrMsg= 'Cannot Change Key As Bag exists.'    
   GoTo ErrHndlr     
  end    
  ***/    
    
  /* Prt Key Not Required - Id No added */    
  select @Row= Count('x') from deleted del where     
    not exists (Select 'x' from inserted ins where ins.OdIdNo=del.OdIdNo and ins.OdCoCd= del.OdCoCd     
    and ins.OdTc= del.OdTc and ins.OdYy= del.OdYy and ins.OdChr= del.OdChr     
    and ins.OdNo= del.OdNo and ins.OdSr= del.OdSr)    
  If @Row > 0    
  begin     
   Select @ErrMsg= 'Cannot Change Company Code/Tc/Yy/Chr/No/Sr'    
   GoTo ErrHndlr    
  end    
  /*********************Geeta***************Emr206************************/    
  /*********************On Hold can be changed ***************************/    
  /****** Cannot change On Hold option if Order= Export Order Of Fgd Table  ******/    
/*  select @row= count('x') from deleted del where     
   not exists (select 'x' from inserted ins where ins.OdCoCd= del.OdCoCd and     
     ins.OdTc= del.OdTc and ins.OdYy= del.OdYy and ins.OdChr= del.OdChr and    
     ins.OdNo= del.OdNo and ins.OdSr= del.OdSr and ins.OdHld= del.OdHld) and     
   exists(Select 'x' from Fgd where FdCoCd= del.OdCoCd and FdExpOdTc= del.OdTc and     
    FdExpOdYy = del.OdYy and FdExpOdChr = del.OdChr and FdExpOdNo = del.OdNo and     
    FdExpOdSr = del.OdSr)    
  if @row> 0     
  begin    
   select @errmsg= 'Cannot Change On Hold Option As Export Order Exists In Fgd Table'    
   goto ErrHndlr    
  end   */    
    
  /*** (Jen 2.14 Next) ***/    
  /****** Cannot change On Hold option to 'Y' if rec exists in InvFgd Table  ******/    
  /* Complete Database  ****Geeta */    
  /****** Manali 3.9.1 - Check skipped for CTB *****/  
  select @row= count('x') from inserted ins where ins.OdHld= 'Y'     
   and (ins.OdPrtFGQty<>0 or exists(Select 'x' from Fgd     
    join InvFgd on IfPrtKey=FdPrtKey and IfFdIdNo=FdIdNo and IfCoCd= FdCoCd and IfFdTc= FdTc   
    and IfFdYy= FdYy and IfFdChr= FdChr and IfFdNo= FdNo and IfFdSr= FdSr    
    and (Select PValue From Param Where PTyp='TC' and PMCd=IfTc)='IN'  
    where FdCoCd= ins.OdCoCd and FdExpOdTc= ins.OdTc and FdExpOdYy = ins.OdYy     
    and FdExpOdChr = ins.OdChr and FdExpOdNo = ins.OdNo and FdExpOdSr = ins.OdSr))    
  if @row> 0     
  begin    
   select @errmsg= 'Cannot Change On Hold Option To ''Y'' As Bag Exists For The Export Order In InvFgd Table'    
   goto ErrHndlr    
  end       
  /*** (Jen 2.14 Next) ***/    
    
    
  /* Geeta CRM Emr 213 @wPkChg=0 added */     
  /****** Cannot Change OdDmCol code when bag exists For Ord Design ******/    
  /* Current Partition */    
  select @row= count('x') from deleted del where del.OdPrtKey=@wCurrPrtn and    
   not exists(Select 'x' from inserted ins where ins.OdIdNo=del.OdIdNo and ins.OdCoCd= del.OdCoCd     
    and ins.OdTc= del.OdTc and ins.OdYy= del.OdYy     
    and ins.OdChr= del.OdChr and ins.OdNo= del.OdNo     
    and ins.OdSr= del.OdSr and ins.OdDmCol= del.OdDmCol)     
   and (OdPrtFgQty >0 or exists(Select 'x' from Bag where BOdIdNo=OdIdNo and BCoCd= OdCoCd     
    and BOdTc= del.OdTc and BOdYy= del.OdYy     
    and BOdChr= del.OdChr and BOdNo= del.OdNo     
    and BOdSr= del.OdSr and BPrtKey=del.OdPrtKey))     
  if @row> 0     
  begin    
   select @errmsg= 'Cannot Change Design Color As Bag Exists'    
   goto ErrHndlr    
  end    
  
  
 /****** Zubin 3.02.01 (270308) OdChgPtrOnStwYN =  PMCd ('YN')  ******/  
 /* Current Partition */  
 select @row= count('x') from inserted where OdPrtKey=@wCurrPrtn and   
  not exists (select 'x' from Param where PTyp= 'YN' and PMCd= OdChgPtrOnStwYN and PSCd= '')  
 if @row> 0     
 begin  
  select @errmsg= 'Change Pointer On STW (Y/N)'  
  goto ErrHndlr  
 end   
  
 /****** Zubin 3.02.01 (270308) OdChgPtrOnStwYN cannot be changed when OrdRm exists ******/  
 select @row= count('x') from deleted del where del.OdPrtKey=@wCurrPrtn and    
  not exists(Select 'x' from inserted ins where ins.OdIdNo=del.OdIdNo and ins.OdCoCd= del.OdCoCd     
   and ins.OdTc= del.OdTc and ins.OdYy= del.OdYy and ins.OdChr= del.OdChr and ins.OdNo= del.OdNo     
   and ins.OdSr= del.OdSr and ins.OdChgPtrOnStwYN= del.OdChgPtrOnStwYN)     
  and exists(Select 'x' from OrdRm where OdPrtKey=OrPrtKey and OdIdNo=OrOdIdNo and OdCoCd= OrCoCd   
   and OdTc= OrTc and OdYy= OrYy and OdChr= OrChr and OdNo= OrNo and OdSr= OrSr)    
 if @row> 0     
 begin    
  Select @ErrMsg= 'Cannot Change ''Change Pointer On STW (Y/N)'' When Rm Exists'  
  goto ErrHndlr    
 end    
  
 /****** Zubin 3.02.01 (270308) OdILabWtFrOrd =  PMCd ('YN')  ******/  
 select @row= count('x') from inserted where OdPrtKey=@wCurrPrtn and  
  not exists (select 'x' from Param where PTyp= 'YN' and PMCd= OdILabWtFrOrd and PSCd= '')  
 if @row> 0     
 begin  
  select @errmsg= 'Labor Weight From Order (Y/N)'  
  goto ErrHndlr  
 end   
   
 /*** Jay 3.3(FG) - dmid/ sfx cannot change for 'FG' tctyp ***/  
 select @row= count('x') from deleted del where del.OdPrtKey=@wCurrPrtn and    
  not exists(Select 'x' from inserted ins where ins.OdIdNo=del.OdIdNo and ins.OdCoCd= del.OdCoCd     
   and ins.OdTc= del.OdTc and ins.OdYy= del.OdYy and ins.OdChr= del.OdChr and ins.OdNo= del.OdNo     
   and ins.OdSr= del.OdSr and (ins.OdDmStkNo= del.OdDmStkNo or ins.OdSfx= del.OdSfx ))  
  and exists (Select 'x' from Param where Ptyp = 'TC' and PValue= 'FG' and PMCd= del.OdTc)  
 if @row> 0     
 begin    
  Select @ErrMsg= 'Cannot Change DmId/ Suffix'  
  goto ErrHndlr    
 end   
  
 /* *** manali 3.5.0 Diamond tolerance at Order Level *** */  
 Select @row=count('x') from inserted Where OdPrtKey=@wCurrPrtn and OdMinDiaTolWt > OdMaxDiaTolWt  
 If @row > 0  
 Begin  
  Select @Errmsg='Minimunm Diamond Tolerance cannot be Greater Than Maximum Diamond Tolerance'  
  Goto ErrHndlr  
 end   
    
  /****** Cannot Change OdOmCtCd code when bag exists For Ord Design ******/    
/*  ################### Check Remove Now onwards     
  Contractor cn be changed any Time in Order Master.    
 select @row= count('x') from deleted del where     
   not exists(Select 'x' from inserted ins where ins.OdCoCd= del.OdCoCd     
    and ins.OdTc= del.OdTc and ins.OdYy= del.OdYy     
    and ins.OdChr= del.OdChr and ins.OdNo= del.OdNo     
    and ins.OdSr= del.OdSr and ins.OdOmCtCd= del.OdOmCtCd)     
   and exists(Select 'x' from Bag where BCoCd= OdCoCd     
    and BOdTc= del.OdTc and BOdYy= del.OdYy     
    and BOdChr= del.OdChr and BOdNo= del.OdNo     
    and BOdSr= del.OdSr)     
  if @row> 0     
  begin    
   select @errmsg= 'Cannot Change Contractor Cd As Bag Exists'    
   goto ErrHndlr    
  end    
*/    
    
  /*** Jen Try 25/06/07 ***/    
  /*** OdOrdQty Cannot Be Changed when OdCls= 'Y' ***/    
  /* Select @row= count('x') from inserted where     
   OdCls= 'Y' and convert (decimal (14, 1), OdOrdQty)<> convert (decimal (14, 1), OdExpQty)     
  if @row> 0     
  Begin    
   select @errmsg= 'Order Quantity Cannot Be Changed When Order Is Closed.'     
   goto ErrHndlr    
  end    
  */     
  /*** Jen Try 25/06/07 ***/    
    
 End    
    
  if @mode = 'A'     
  Begin    
/* **** Zubin 211 TCTYP= 'QF' added **** */     
  /****** Check for valid Design or whether the Design is on hold for Order Tc Types 'SO', 'PL', 'QT', 'QF' ******/    
  /* Current Partition */    
 /***** Manali Trading Module - 'JT' added *****/  
  -- 4.1.3.0 - DmValid  
  select @row= count('x') from inserted     
    Join OrdMst On OmIdNo=OdOmIdNo and OmCoCd=OdCoCd And OmTc=OdTc And OmYy=OdYy And OmChr=OdChr And OmNo=OdNo and OmPrtKey=OdPrtKey    
    where OdPrtKey=@wCurrPrtn and OmReYy= '' And OmReChr= '' And OmReNo=0 And      
    (Select PValue from Param where PTyp= 'TC' and PMCd= OdTc) in ('PL', 'QT', 'SO', 'QB', 'JT') and     
    not exists (select 'x' from DsgMst where DmIdNo=OdDmIdNo and DmTcTyp= 'DM'     
        and DmHld= 'N' And DmValidYn='Y' and DmCd= OdDmCd and DmSz= ''  and DmPrtKey=OdPrtKey)    
  if @row> 0     
  begin    
   select @errmsg= 'Design Code Not Defined Or Is On Hold'    
   goto ErrHndlr    
  end    
    
  /* Current Partition */    
  /***** Manali Trading Module - 'JT' added *****/  
  -- 4.1.3.0 - DmValid  
  select @row= count('x') from inserted     
    Join OrdMst On OmIdNo=OdOmIdNo and OmCoCd=OdCoCd And OmTc=OdTc And OmYy=OdYy And OmChr=OdChr And OmNo=OdNo and OmPrtKey=OdPrtKey    
   where OdPrtKey=@wCurrPrtn and OmReYy= '' And OmReChr= '' And OmReNo=0 And      
    (Select PValue from Param where PTyp= 'TC' and PMCd= OdTc) in ('PL', 'QT', 'SO', 'QB', 'JT') and     
    Exists (select 'x' from DsgMst where DmIdNo=OdDmIdNo and DmTcTyp= 'DM'     
     and (DmHld= 'Y' Or DmValidYn='N') and DmCd= OdDmCd and DmSz= OdDmSz and DmPrtKey=OdPrtKey)    
  if @row> 0     
  begin    
   select @errmsg= 'This Design Size Is On Hold/Invalid'    
   goto ErrHndlr    
  end    
    
  /*** (Jen 2.11) ***/    
  /* Current Partition */      
 /* Manali Trading Module */  
  select @row= count('x') from inserted where OdPrtKey=@wCurrPrtn and    
   ((Select PValue from Param where PTyp= 'TC' and PMCd= OdTc) <> 'QB' And  
  (Select PMCd from Param where PTyp= 'TC' and PMCd= OdTc and PValue='JT') <> 'JCN')   
  and (OdBYy <> '' or OdBChr <> '' or OdBNo <> 0)    
  if @row> 0     
  begin    
   select @errmsg= 'Bag Number Should Be Blank For This Voucher Type'    
   goto ErrHndlr    
  end    
    
  /*** Check For validity of Bag Number ***/    
  select @row= count('x') from inserted Ins where Ins.OdPrtKey=@wCurrPrtn and    
   ((Select PValue from Param where PTyp= 'TC' and PMCd= OdTc) in ('QB')   
  Or ((Select PValue from Param where PTyp= 'TC' and PMCd= OdTc) in ('JT') and OdTc ='JCN'))  
 and not exists (Select 'x' From Bag Where BCoCd= OdCoCd and BYy= OdBYy     
      and BChr = OdBChr And BNo= OdBNo and BLoc = 'PFG'     
      and BOdDmCd= OdDmCd and BPrtKey=OdPrtKey)    
   And Not Exists (Select 'x' From OrdDsg Od Where Od.OdCoCd=ins.OdCoCd And Od.OdTc In ('QB', 'JCN') And     
      Od.OdBYy=Ins.OdBYy And Od.OdBChr=Ins.OdBChr And Od.OdBNo=Ins.OdBNo And Od.OdPrtKey=Ins.OdPrtKey)    
  if @row> 0     
  begin    
   select @errmsg= 'Bag Number Not Defined For the Design'    
   goto ErrHndlr    
  end    
  /*** (Jen 2.11) ***/    
    
  
  /* Manali Trading Module */  
  if @wTcTyp = 'JT'   
  Begin  
  select @row= count('x') from inserted Ins where Ins.OdPrtKey=@wCurrPrtn and    
   (Select PValue from Param where PTyp= 'TC' and PMCd= OdTc) in ('JT') and OdTc ='JCN' and  
   not exists (Select 'x' From Bag Where BCoCd= OdCoCd and BYy= OdBYy     
      and BChr = OdBChr And BNo= OdBNo and BLoc = 'PFG' and BCls='N'  
      and BOdDmCd= OdDmCd and BPrtKey=OdPrtKey)    
  if @row> 0     
  begin    
   select @errmsg= 'Bag Closed'    
   goto ErrHndlr    
  end  
  End  
    
 End    
/* **** Zubin 211 TCTYP= 'QF' added **** */     
 if @mode in ('A','C')    
 Begin    
    
  /******  OdCoCd, OdTc, OdYy, OdChr, OdNo combination does not exist in OrdMst  ******/    
  /* Current Partition */    
  select @row= count('x') from inserted where OdPrtKey=@wCurrPrtn and not exists (select 'x' from OrdMst where     
     OmIdNo=OdOmIdNo and OmCoCd= OdCoCd and OmTc= OdTc and OmYy= OdYy and     
     OmChr= OdChr and OmNo= OdNo and OmPrtKey=OdPrtKey)    
  if @row> 0     
  begin    
   select @errmsg= 'Voucher Not Defined In Order Master'    
   goto ErrHndlr    
  end    
    
  /******  OdSr  >  0  ******/    
  /* Current Partition */    
  select @row= count('x') from inserted where OdPrtKey=@wCurrPrtn and OdSr<= 0     
  if @row> 0     
  begin    
   select @errmsg= 'Design Serial Number Cannot Be <= Zero'    
   goto ErrHndlr    
  end    
    
  /*** (Jen 2.13) ***/    
  /******  If Voucher copied from 'Return Memo' voucher, OdSr = RedSr of RetMemDet  ******/    
  /* Current Partition****Geeta */    
  
  /*** (Jen 3.1.4) ***/     
  select @row= count('x') from inserted ins join OrdMst on OdOmIdNo=OmIdNo and OdCoCd= OmCoCd and OdTc= OmTc and OdYy= OmYy     
   and OdChr= OmChr and OdNo= OmNo and OdPRtKey=OmPRtKey where OdPRtKey=@wCurrPrtn and OmReNo> 0     
   and not exists (Select 'x' from RetMemDet join InvDsg on IdCoCd= RedCoCd and IdTc= RedIdTc     
      and IdYy= RedIdYy and IdChr= RedIdChr and IdNo= RedIdNo and IdSr= RedIdSr     
     where RedCoCd= OmCoCd and RedYy= OmReYy and RedChr= OmReChr and RedNo= OmReNo     
      and RedSr= OdSr /* and RedQty= OdOrdQty */ and IdDmCd= OdDmCd     
      and IdSfx= OdSfx and IdDmSz= OdDmSz     
     ) and @ReplAgt<>1    
     and (@Mode='A' or   
          (@Mode='C' and   
            exists(Select 'x' from Deleted Del Where   
            del.OdIdNo=ins.OdIdNo and del.OdCoCd= ins.OdCoCd and del.OdTc= ins.OdTc and del.OdYy= ins.OdYy     
            and del.OdChr= ins.OdChr and del.OdNo= ins.OdNo and del.OdPRtKey=ins.OdPRtKey  
            and (del.OdDmCd<> ins.OdDmCd or del.OdSfx<> ins.OdSfx or del.OdDmSz<> ins.OdDmSz or  del.OdSr<> ins.OdSr)  
            )))  
  if @row> 0     
  begin    
   select @errmsg= 'Design Serial Number/ Design Code/Sfx/Sz/ Qty Not As Per Return Memo'    
   goto ErrHndlr    
  end    
  /*** (Jen 3.1.4) ***/  
  
  /*** (Bef 3.1.4)       
  select @row= count('x') from inserted join OrdMst on OdOmIdNo=OmIdNo and OdCoCd= OmCoCd and OdTc= OmTc and OdYy= OmYy     
   and OdChr= OmChr and OdNo= OmNo and OdPRtKey=OmPRtKey where OdPRtKey=@wCurrPrtn and OmReNo> 0     
   and not exists (Select 'x' from RetMemDet join InvDsg on IdCoCd= RedCoCd and IdTc= RedIdTc     
      and IdYy= RedIdYy and IdChr= RedIdChr and IdNo= RedIdNo and IdSr= RedIdSr     
     where RedCoCd= OmCoCd and RedYy= OmReYy and RedChr= OmReChr and RedNo= OmReNo     
      and RedSr= OdSr and RedQty= OdOrdQty and IdDmCd= OdDmCd     
      and IdSfx= OdSfx and IdDmSz= OdDmSz     
     ) and @ReplAgt<>1    
  if @row> 0     
  begin    
   select @errmsg= 'Design Serial Number/ Design Code/Sfx/Sz/ Qty Not As Per Return Memo'    
   goto ErrHndlr    
  end    
  (Bef 3.1.4) ***/    
   
  /*** (Jen 2.13) ***/    
    
    
    
/*  Manoj 2.10.0  */    
  /******  OdDmSz  =  Blank  or  PMCd ('DMSZ')  ******/    
/*  select @row= count('x') from inserted where OdDmSz<> '' and      
   not exists (select 'x' from Param where PTyp= 'DMSZ' and PMCd= OdDmSz)    
  if @row> 0     
  begin    
   select @errmsg= 'Design Size Not Defined'    
   goto ErrHndlr    
  end    
*/    
  /* Current Partition */    
  select @row= count('x') from inserted Ins where Ins.OdPrtKey=@wCurrPrtn and OdDmSz<> '' and     
   Ins.OdDmSz <> IsNull((Select Del.OdDmSz From Deleted Del Where Ins.OdIdNo=Del.OdIdNo and     
    Ins.OdCoCd = Del.OdCoCd And Ins.OdTc = Del.OdTc And Ins.OdYy = Del.OdYy     
    And Ins.OdChr = Del.OdChr And Ins.OdNo = Del.OdNo And Ins.OdSr = Del.OdSr And Ins.OdPrtKey=Del.OdPrtKey), '') And     
   not exists (select 'x' from Param where PTyp= 'DMSZ' and PMCd= OdDmSz And PValidYn = 'Y')    
  if @row> 0     
  begin    
   select @errmsg= 'Design Size Not Defined / Invalid.'    
   goto ErrHndlr    
  end    
/*  Manoj 2.10.0  */    
    
    
/* **** Zubin 211 (Include below checks if TcTyp= 'QF' included) **** */     
  /****** Check for valid BYy, BChr, BNo (shld be in 'PFG') ******/     
/*    
  select @row= count('x') from inserted where     
   (Select PValue from Param where PTyp= 'TC' and PMCd= OdTc) = 'QF' and     
   Not Exists (select 'x' from Bag Where BCoCd= OdCoCd and BYy= OdBYy And BChr= OdBChr And     
    BNo= OdBNo and BLoc = 'PFG')    
  if @row> 0     
  begin    
   select @errmsg= 'Bag Number Invalid Or Bag Not In PFG'    
   goto ErrHndlr    
  end    
*/    
  /* Design Cd shld belong to the Bag */     
/*  select @row= count('x') from inserted where     
   (Select PValue from Param where PTyp= 'TC' and PMCd= OdTc) = 'QF' and     
   Not Exists (select 'x' from Bag Where BCoCd= OdCoCd and BYy= OdBYy And BChr= OdBChr And     
    BNo= OdBNo and BOdDmCd= OdDmCd)    
  if @row> 0     
  begin    
   select @errmsg= 'Design Code Invalid '    
   goto ErrHndlr    
  end    
*/    
  /* Design Sfx shld belong to the Bag */     
/*  select @row= count('x') from inserted where     
   (Select PValue from Param where PTyp= 'TC' and PMCd= OdTc) = 'QF' and     
   Not Exists (select 'x' from Bag Where BCoCd= OdCoCd and BYy= OdBYy And BChr= OdBChr And     
    BNo= OdBNo and BOdSfx= OdSfx)    
  if @row> 0     
  begin    
   select @errmsg= 'Design Suffix Invalid '    
   goto ErrHndlr    
  end    
*/    
  /* Design Sz shld belong to the Bag */     
/*  select @row= count('x') from inserted where     
   (Select PValue from Param where PTyp= 'TC' and PMCd= OdTc) = 'QF' and     
   Not Exists (select 'x' from Bag Where BCoCd= OdCoCd and BYy= OdBYy And BChr= OdBChr And     
    BNo= OdBNo and BOdDmSz= OdDmSz)    
  if @row> 0     
  begin    
   select @errmsg= 'Design Size Invalid '    
   goto ErrHndlr    
  end     
*/    
  /* OdQfRmFrFg shld be 'YN' */     
/*  select @row= count('x') from inserted where     
   (Select PValue from Param where PTyp= 'TC' and PMCd= OdTc) = 'QF' and     
   Not Exists (select 'x' from Param Where PTyp= 'YN' And PMCd= OdQfRmFrFg and PSCd= '')    
  if @row> 0     
  begin     
   select @errmsg= 'Get Raw Material From FG Invalid '     
   goto ErrHndlr     
  end     
*/    
/* **** Zubin 211 (Include above checks if TcTyp= 'QF' included) **** */     
    
    
  /******  OdPrdSeq  =  PMCd ('PRDSEQ')  ******/    
  /* Current Partition */     
  select @row= count('x') from inserted where OdPrtKey=@wCurrPrtn and not exists (select 'x' from vParam where vPCoCd= OdCoCd and     
     vPTyp= 'PRDSEQ' and vPMCd= OdPrdSeq)    
  if @row> 0     
  begin    
   select @errmsg= 'Production Sequence Not Defined'    
   goto ErrHndlr    
  end    
    
  /****** Production Quantity for the OrdSr should be >= Finished Goods Quantity for that OrdSr ******/    
  /* Current Partition */    
  select @row= count('x') from inserted where OdPrtKey=@wCurrPRtn and OdPrdQty<  OdFgQty     
  if @row> 0     
  begin    select @errmsg= 'Production Qty Should Be >= Finished Goods Qty'    
   goto ErrHndlr    
  end    
    
  /******  OdPrtCd  =  PMCd ('PRTCD')  ******/     
/* Current Partition */     
select @row= count('x') from inserted where OdPrtKey=@wCurrPrtn and    
   not exists (select 'x' from Param where PTyp= 'PRTCD' and PMCd= OdPrtCd)    
  if @row> 0     
  begin    select @errmsg= 'Priority Code Not Defined'    
   goto ErrHndlr    
  end    
   /******  OdHld  =  PMCd ('YN')  ******/    
  /* Current Partition */    
  select @row= count('x') from inserted where OdPrtKey=@wCurrPRtn and    
   not exists (select 'x' from Param where PTyp= 'YN' and PMCd= OdHld)    
  if @row> 0     
  begin    
   select @errmsg= 'Enter On Hold (Y/N)'    
   goto ErrHndlr    
  end    
   /******  If OdHld = 'N'  then  OdHldDesc  =  Blank  ******/    
  /* Current Partition */    
  select @row= count('x') from inserted where OdPRtKey=@wCurrPRtn and OdHld= 'N' and OdHldDesc<> ''     
  if @row> 0     
  begin    
   select @errmsg= 'Holding Desc Should Be Entered Only When Design Is On Hold'    
   goto ErrHndlr    
  end    
    
  /******  OdVaCtg  =  PMCd ('VACTG')  ******/     
  /* Current Partition */    
  select @row= count('x') from inserted where OdPrtKey=@wCurrPrtn and    
   not exists (select 'x' from Param where PTyp= 'VACTG' and PMCd= OdVaCtg)    
  if @row> 0     
  begin    
   select @errmsg= 'Value Addition Category Not Defined'    goto ErrHndlr    
  end    
    
  /******  OdKt  =  PMCd ('KT')  ******/    
  /* Current Partition */    
  select @row= count('x') from inserted where OdPrtKEy=@wCurrPrtn and    
   not exists (select 'x' from Param where PTyp= 'KT' and PMCd= OdKt)    
  if @row> 0     
  begin    
   select @errmsg= 'Karatage Not Defined'    
   goto ErrHndlr    
  end    
    
  /******  OdMulBy  >  0  ******/    
  /* Current Partition */     
  select @row= count('x') from inserted where OdPrtKey=@wCurrPrtn and OdMulBy<= 0     
  if @row> 0     
  begin    
   select @errmsg= 'Multiplying Factor Cannot Be <= Zero'    
   goto ErrHndlr    
  end    
   /******  OdFixPrc  =  PMCd ('YN')  ******/    
  /* Current Partition */    
  select @row= count('x') from inserted where OdPrtKEy=@wCurrPrtn and    
   not exists (select 'x' from Param where PTyp= 'YN' and PMCd= OdFixPrc)    
  if @row> 0     
  begin    select @errmsg= 'Enter Fix Price (Y/N)'    
   goto ErrHndlr    
  end    
    
  /******  OdGldAs  =  PMCd ('GLDAS')  ******/    
  /*** (Jen 2.14) ***/     
  /* select @row= count('x') from inserted where OdGldAs<> '' and     
   not exists (select 'x' from Param where PTyp= 'GLDAS' and PMCd= OdGldAs    
   and not (',' + PDesc225 + ',' like '%,G,%' Or ',' + PDesc225 + ',' like '%,P,%' Or ',' + PDesc225 + ',' like '%,S,%')     
   )    
  */    
  /* Current Partition */    
  select @row= count('x') from inserted where OdPrtKey=@wCurrPRtn and OdGldAs<> '' and     
   not exists (select 'x' from Param where PTyp= 'GLDAS' and PMCd= OdGldAs)     
  /*** (Jen 2.14) ***/     
    
  /*** (Bef 2.14)     
  select @row= count('x') from inserted where     
   not exists (select 'x' from Param where PTyp= 'GLDAS' and PMCd= OdGldAs)    
  (Bef 2.14) ***/    
  if @row> 0     
  begin    
   select @errmsg= 'Gold As Value Not Defined'    
   goto ErrHndlr    
  end    
    
  /******  OdLabAs  =  PMCd ('GLDAS') - Sachin 2.13.0 - 25-05-2006  ******/    
  /* Current Partition */    
  select @row= count('x') from inserted where OdPrtKEy=@wCurrPrtn and    
   not exists (select 'x' from Param where PTyp= 'GLDAS' and PMCd= OdLabAs)    
  if @row> 0     
  begin    
   select @errmsg= 'Labour As Value Not Defined'    
   goto ErrHndlr    
  end    
    
  /******  OdIWtEqOrd  =  PMCd ('YN')  ******/    
  /* Current Partition */    
  select @row= count('x') from inserted where OdPrtKey=@wCurrPrtn and    
   not exists (select 'x' from Param where PTyp= 'YN' and PMCd= OdIWtEqOrd)    
  if @row> 0     
  begin    
   select @errmsg= 'Invoice Wt. Equal To Order Wt. (Y/N)'    
   goto ErrHndlr    
  end    
    
  /******  If OdIWtEqOrd  =  'Y'  then  OdIWtFrOrd  =  Blank      
   If OdIWtEqOrd  =  'N'  then  OdIWtFrOrd  =  Blank  or  PMCd ('IWTFRORD')  ******/    
  /* Current Partition */    
  select @row= count('x') from inserted where OdIWtFrOrd<> '' and OdPrtKEy=@wCurrPrtn and    
   not exists (select 'x' from Param where PTyp= 'IWTFRORD' and PMCd= OdIWtFrOrd)    
  if @row> 0     
  begin    
   select @errmsg= 'Invoice Wt. From Order Not Defined'    
   goto ErrHndlr    
  end     
  /* Current Partition */    
  select @row= count('x') from inserted where OdPrtKey=@wCurrPrtn and (OdIWtFrOrd<> '' and OdIWtEqOrd= 'Y')     
  if @row> 0     
  begin    
   select @errmsg= 'Mismatch in values InvWtEqOrd and InvWtFrOrd'    
   goto ErrHndlr    
  end     
    
  /******  OdIGldRtEqOrd  =  PMCd ('YN')  ******/    
  /* Current Partition */    
  select @row= count('x') from inserted where OdPrtKey=@wCurrPrtn and    
   not exists (select 'x' from Param where PTyp= 'YN' and PMCd= OdIGldRtEqOrd)    
  if @row> 0     
  begin    
   select @errmsg= 'Invoice Gold Rate Equal To Order Rate (Y/N)'    
   goto ErrHndlr    
  end    
    
  /******  OdCls  =  PMCd ('YN')  ******/    
  /* Current Partition */    
  select @row= count('x') from inserted where OdPrtKey=@wCurrPrtn and     
     not exists (select 'x' from Param where PTyp='YN' and PMCd= OdCls)    
  if @row> 0     
  begin    
   select @errmsg= 'Order Closed (Y/N)'    
   goto ErrHndlr    
  end    
    
/*  Manoj 2.10.0  */    
  /******  OdDmCol  =  PMCd ('DMCOL')  ******/    
/*  select @row= count('x') from inserted where     
   not exists (select 'x' from Param where PTyp= 'DMCOL' and PMCd= OdDmCol)    
  if @row> 0     
  begin    
   select @errmsg= 'Design Colour Not Defined'    
   goto ErrHndlr    
  end    
*/    
    
/*** (Bef 2.14 Next Patch 2)     
  select @row= count('x') from inserted Ins where     
   Ins.OdDmCol <> IsNull((Select Del.OdDmCol From Deleted Del Where     
    Ins.OdCoCd = Del.OdCoCd And Ins.OdTc = Del.OdTc And Ins.OdYy = Del.OdYy     
    And Ins.OdChr = Del.OdChr And Ins.OdNo = Del.OdNo And Ins.OdSr = Del.OdSr), '') And     
   not exists (select 'x' from Param where PTyp= 'DMCOL' and PMCd= OdDmCol And PValidYn = 'Y')    
(Bef 2.14 Next Patch 2) ***/    
    
  /*** (Jen 2.14 Next Patch 2) ***/     
  /* Current Partition */    
  select @row= count('x') from inserted Ins where Ins.OdPrtKEy=@wCurrPRtn and    
   Ins.OdDmCol= '' or     
   ( Ins.OdDmCol <> IsNull((Select Del.OdDmCol From Deleted Del Where Ins.OdIdNo=Del.OdIdNo and     
      Ins.OdCoCd = Del.OdCoCd And Ins.OdTc = Del.OdTc And Ins.OdYy = Del.OdYy And     
      Ins.OdChr = Del.OdChr And Ins.OdNo = Del.OdNo And Ins.OdSr = Del.OdSr and Ins.OdPRtKey=Del.OdPrtKey), '') And     
    not exists (select 'x' from Param where PTyp= 'DMCOL' and PMCd= OdDmCol And PValidYn = 'Y')     
   )     
  /*** (Jen 2.14 Next Patch 2) ***/     
  if @row> 0     
  begin     
   select @errmsg= 'Design Colour Not Defined / Invalid.'     
   goto ErrHndlr     
  end     
/*  Manoj 2.10.0  */    
    
  /******  OdWh  =  PMCd ('WH')  ******/     
  /* Current Partition */    
  select @row= count('x') from inserted where OdPRtKey=@wCurrPrtn and    
   not exists (select 'x' from Param where PTyp= 'WH' and PMCd= OdWh)    
  if @row> 0      
  begin    
   select @errmsg= 'W/H Set Option Not Defined'    
   goto ErrHndlr    
  end    
    
  /******  OdLine  =  vPMCd ('LINE')  ******/     
  /* Current Partition */    
  select @row= count('x') from inserted where OdPrtKey=@wCurrPrtn and    
   not exists (select 'x' from vParam where vPCoCd= OdCoCd and vPTyp= 'LINE' and vPMCd= OdLine and vPMCd<> 'ZZZ')    
  if @row> 0    begin    
   select @errmsg= 'Line Not Defined/ Line Cannot Be ''ZZZ'' '    
   goto ErrHndlr    
  end    
    
  /******  OdWDiaAvlblDt And OdHDiaAvlblDt <>  '01/01/1980'  or  '01/01/1900'  (The dates '01/01/1980'  and  '01/01/1900'  are considered as blank) ******/    
  /* Current Partition */    
  select @row= count('x') from inserted where OdPrtKey=@wCurrPrtn and (OdWDiaAvlblDt in('01/01/1980', '01/01/1900') or     
   OdHDiaAvlblDt in('01/01/1980', '01/01/1900'))    
  if @row> 0     
  begin    
   select @errmsg= 'Wax/ Hand Set Diamond Availability Date Cannot Be Blank'    
   goto ErrHndlr    
  end    
    
  /****** OdGmChk = PMCd('CPYRT') ******/     
  /* Current Partition */    
  select @row= count('x') from inserted where OdPRtKey=@wCurrPRtn and    
    not exists (select 'x' from Param where PTyp= 'CPYRT' and PMCd= OdGmChk)    
  if @row> 0     
  begin    
   select @errmsg= 'Gram Check Option Not Defined'    
   goto ErrHndlr    
  end    
    
  /****** OdChkTol = PMCd('YN') ******/     
  /* Current Partition */    
  select @row= count('x') from inserted where OdPrtKey=@wCurrPrtn and    
    not exists (select 'x' from Param where PTyp= 'YN' and PMCd= OdChkTol)    
  if @row> 0     
  begin    
   select @errmsg= 'Should Tolerence Be Checked? (Specify Y/ N)'    
   goto ErrHndlr    
  end    
    
  /******  Cannot Change OdDmCd  ******/    
  /* Current Partition */    
  select @row= count('x') from  deleted del where del.OdPRtKey=@wCurrPrtn and    
   not exists (select 'x' from inserted ins where ins.OdIdNo=del.OdIdNo and ins.OdCoCd= del.OdCoCd and     
    ins.OdTc= del.OdTc and ins.OdYy= del.OdYy and ins.OdChr= del.OdChr and     
    ins.OdNo= del.OdNo and ins.OdSr= del.OdSr and ins.OdDmCd= del.OdDmCd)     
  if @row> 0     
  begin    
   select @errmsg= 'Cannot Change Design'    
   goto ErrHndlr    
  end    
    
  /*** (Jen 2.11) ***/    
  /******  Cannot Change BagNo  ******/    
  /* Current Partition */    
      
  select @row= count('x') from  deleted del where del.OdPRtKEy=@wCurrPrtn and     
   not exists (select 'x' from inserted ins where ins.OdIdNo=del.OdIdNo and ins.OdCoCd= del.OdCoCd and     
    ins.OdTc= del.OdTc and ins.OdYy= del.OdYy and ins.OdChr= del.OdChr and     
    ins.OdNo= del.OdNo and ins.OdSr= del.OdSr and ins.OdPRtKey=del.OdPRtKey and ins.OdBYy= del.OdBYy and     
    ins.OdBChr= del.OdBChr and ins.OdBNo= del.OdBNo)     
  if @row> 0     
  begin    
   select @errmsg= 'Cannot Change Bag Number'    
   goto ErrHndlr    
  end    
  /*** (Jen 2.11) ***/    
    
  /******  Cannot change OdDmSz, OdSfx when Bag exists for the Order Sr  ******/    
  /* Current Partition - Check */    
  select @row= count('x') from  deleted del where del.OdPrtKey=@wCurrPrtn and     
   not exists (select 'x' from inserted ins where ins.OdIdNo=del.OdIdNo and ins.OdCoCd= del.OdCoCd and     
     ins.OdTc= del.OdTc and ins.OdYy= del.OdYy and ins.OdChr= del.OdChr and    
     ins.OdNo= del.OdNo and ins.OdSr= del.OdSr and ins.OdPrtKey=del.OdPrtKey and     
     ins.OdDmSz= del.OdDmSz and ins.OdSfx= del.OdSfx)     
   and (exists(Select 'x' from Bag where BOdIdNo=del.OdIdNo and BCoCd= del.OdCoCd and BOdTc= del.OdTc and     
     BOdYy= del.OdYy and BOdChr= del.OdChr and BOdNo= del.OdNo and     
     BOdSr= del.OdSr and BPrtKey=del.OdPRtKey) Or OdPRtFGQty<>0 Or OdPrtExpQty<>0)    
  if @row> 0     
  begin    
   select @errmsg= 'Cannot Change Design Size/ Suffix When Bag Exists'    
   goto ErrHndlr    
  end    
    
  /* Zubin 213 (CurCd added) */     
  /****** Only One Record Allowed For A Customer, Design, Size, Suffix Combination In Party-Design Combination Master ******/    
  /*** The company code for a 'PL' TcType is always 'ZZZ' ***/    
  /* Complete Database*****Geeta */    
 /* ***** Manali Trading Module **** */  
  select @row= count('x') from inserted ins     
   join OrdMst insOm On insOm.OmIdNo=ins.OdOmIdNo and insOm.OmCoCd= ins.OdCoCd and insOm.OmTc= ins.OdTc and insOm.OmYy= ins.OdYy     
     and insOm.OmChr= ins.OdChr and insOm.OmNo= ins.OdNo      
  where OdPrtKey=@wCurrPrtn and    
     (Select PValue from Param where PTyp='TC' and PMCd= ins.OdTc)= 'PL' and     
     (Select count('x') from OrdDsg Od join OrdMst Om on  Om.OmIdNo=Od.OdOmIdNo and Om.OmCoCd= Od.OdCoCd     
    and Om.OmTc= Od.OdTc and Om.OmYy= Od.OdYy and Om.OmChr= Od.OdChr     
    and Om.OmNo= Od.OdNo where Om.OmCoCd= insOm.OmCoCd and Om.OmTc= (Case When insOm.OmCmCtg = 'S' Then 'SPL' Else 'PL' End)     
    and Om.OmCmCtg= insOm.OmCmCtg and Om.OmCmCd= insOm.OmCmCd and Od.OdDmCd= ins.OdDmCd     
    and Od.OdSfx= ins.OdSfx and Od.OdDmSz= ins.OdDmSz and Om.OmCmCurCd= insOm.OmCmCurCd    and Om.OmPrtKey=ins.OdPrtKey and Om.OmPrtKey=Od.OdPrtKey) > 1     
  if @row> 0     
   begin    
     select @errmsg= 'Price List Already Defined for this Cust, Curr, Dsg, Sz, Sfx combination'    
     goto ErrHndlr    
   end    
  
/* Bef Trading Module */  
/*  select @row= count('x') from inserted ins     
   join OrdMst insOm On insOm.OmIdNo=ins.OdOmIdNo and insOm.OmCoCd= ins.OdCoCd and insOm.OmTc= ins.OdTc and insOm.OmYy= ins.OdYy     
     and insOm.OmChr= ins.OdChr and insOm.OmNo= ins.OdNo      
  where OdPrtKey=@wCurrPrtn and    
     (Select PValue from Param where PTyp='TC' and PMCd= ins.OdTc)= 'PL' and     
     (Select count('x') from OrdDsg Od join OrdMst Om on  Om.OmIdNo=Od.OdOmIdNo and Om.OmCoCd= Od.OdCoCd     
    and Om.OmTc= Od.OdTc and Om.OmYy= Od.OdYy and Om.OmChr= Od.OdChr     
    and Om.OmNo= Od.OdNo where Om.OmCoCd= insOm.OmCoCd and Om.OmTc= 'PL'     
    and Om.OmCmCd= insOm.OmCmCd and Od.OdDmCd= ins.OdDmCd     
    and Od.OdSfx= ins.OdSfx and Od.OdDmSz= ins.OdDmSz and Om.OmCmCurCd= insOm.OmCmCurCd   
 and Om.OmPrtKey=ins.OdPrtKey and Om.OmPrtKey=Od.OdPrtKey) > 1     
  if @row> 0     
   begin    
     select @errmsg= 'Price List Already Defined for this Cust, Curr, Dsg, Sz, Sfx combination'    
     goto ErrHndlr    
   end    
 */  
   
  /****** Check for valid Sketch Design Codes for Order Tc Type 'QS' ******/    
  /* Current Partition */    
  select @row= count('x') from inserted where OdPrtKey=@wCurrPrtn and    
   (Select PValue from Param where PTyp= 'TC' and PMCd= OdTc) in ('QS') and     
   not exists (select 'x' from DsgMst where DmIdNo=OdDmIdNo and DmTcTyp= 'SM'     
        and DmCd= OdDmCd and DmSz= '' and DmPrtKey=OdPrtKey )    
  if @row> 0     
  begin    
   select @errmsg= 'Sketch Design Code Not Defined'    
   goto ErrHndlr    
  end    
    
/* **** Zubin 211 **** */     
  /****** Prd Delivery date cannot be null or less than Order Date ******/    
/*    
  select @row= count('x') from inserted where OdDelDt= '01/01/1900' or OdDelDt= '01/01/1980' or     
   OdDelDt < (Select OmDt from OrdMst where OmCoCd= OdCoCd and OmTc= OdTc     
     and OmYy= OdYy and OmChr= OdChr and OmNo= OdNo)       
*/    
  /* Current Partition */    
  select @row= count('x') from inserted where OdPrtKey=@wCurrPrtn and OdDelDt= '01/01/1900' or (OdDelDt<> '01/01/1980' and     
   OdDelDt < (Select OmDt from OrdMst where OmIdNo=OdOmIdNo and OmCoCd= OdCoCd and OmTc= OdTc     
     and OmYy= OdYy and OmChr= OdChr and OmNo= OdNo and OmPrtKey=OdPrtKey))     
/* **** Zubin 211 **** */     
  if @row> 0     
  begin    
   select @errmsg= 'Production Delivery Date Cannot Be Less Than Order Date'    
   goto ErrHndlr     
  end     
/* **** Zubin 211 **** */     
  /****** Export Delivery date cannot be null or less than Order Date ******/    
/*  select @row= count('x') from inserted where OdExpDelDt= '01/01/1900' or OdExpDelDt= '01/01/1980' or     
   OdExpDelDt < (Select OmDt from OrdMst where OmCoCd= OdCoCd     
    and OmTc= OdTc and OmYy= OdYy and OmChr= OdChr and OmNo= OdNo)     
*/    
  /* Current Partition */    
  select @row= count('x') from inserted where OdPrtKey=@wCurrPrtn and OdExpDelDt= '01/01/1900' or (OdExpDelDt<> '01/01/1980' and     
   OdExpDelDt < (Select OmDt from OrdMst where OmIdNo=OdOmIdNo and OmCoCd= OdCoCd     
    and OmTc= OdTc and OmYy= OdYy and OmChr= OdChr and OmNo= OdNo and OmPrtKey=OdPrtKey))     
    
/* **** Zubin 211 **** */     
  if @row> 0     
  begin    
   select @errmsg= 'Export Delivery Date Cannot Be Less Than Order Date'    
   goto ErrHndlr   end    
  /****** Order Qty, Prd Qty should be >= 0 ******/    
    
  /* Current Partition */    
  select @row= count('x') from inserted where OdPrtKey=@wCurrPrtn and OdOrdQty < 0 or OdPrdQty< 0    
  if @row> 0     
  begin    select @errmsg= 'Order/ Prd Qty Should Be >= 0 '    
   goto ErrHndlr    
  end    
    
  /****** Order Quantity should be >= Export Quantity ******/    
  /* Current Partition */    
 select @row= count('x') from inserted where OdPrtKey=@wCurrPrtn and    
   Round(OdOrdQty, 1) < Round(OdExpQty, 1)     
  
  if @row> 0     
  begin    select @errmsg= 'Order Qty Should Be >= Export Qty '    
   goto ErrHndlr    
  end    
    
  /*** (Jen 2.14 Next Trial)     
  ****** Ord Qty should be = (Entered Qty + Amd Qty) ******    
  select @row= count('x') from inserted where     
   convert (decimal (16, 1), OdOrdQty) <> convert (decimal (16, 1), OdOrdEnt     
     + IsNull((Select sum (OaOrdQty) From OrdAmd where OaCoCd= OdCoCd     
      and OaTc= OdTc and OaYy= OdYy and OaChr= OdChr and OaNo= OdNo     
      and OaSr= OdSr), 0))    
  if @row> 0     
  begin    select @errmsg= 'Order Qty Should Be = Entered Qty + Amendment Qty'    
   goto ErrHndlr    
  end    
    
  ****** Prd Qty should be = (Entered Qty + Amd Qty) ******    
  select @row= count('x') from inserted where     
   convert (decimal (16, 1), OdPrdQty) <> convert (decimal (16, 1), OdPrdEnt     
     + IsNull((Select sum (OaPrdQty) From OrdAmd where OaCoCd= OdCoCd     
      and OaTc= OdTc and OaYy= OdYy and OaChr= OdChr and OaNo= OdNo     
      and OaSr= OdSr), 0))    
  if @row> 0     
  begin    select @errmsg= 'Production Qty Should Be = Entered Qty + Amendment Qty'    
   goto ErrHndlr    
  end    
  (Jen 2.14 Next Trial) ***/    
    
    
  /****** Production Quantity for the OrdSr should be >= the total Bag Quantity for that OrdSr ******/    
/* **** Zubin 211 **** */     
/*  select @row= count('x') from inserted where     
   OdPrdQty <  IsNull((Select sum(BQty) from Bag where     
                   BCoCd = OdCoCd and BOdTc = OdTc and     
            BOdYy = OdYy and BOdChr = OdChr and     
                   BOdNo = OdNo and BOdSr = OdSr), 0)    
*/    
/* *** Jay 2.13(CT) *** */     
  /* Current Partition */    
  select @row= count('x') from inserted where OdPrtKEy=@wCurrPrtn and    
   OdPrdQty-OdPrtFgQty <  IsNull((Select sum(BQty) from Bag Where BOdIdNo=OdIdNo and    
             BCoCd = OdCoCd and BOdTc = OdTc and BOdYy = OdYy and BOdChr = OdChr     
       and BOdNo = OdNo and BOdSr = OdSr and BPrtKey=OdPrtKey and    
      (select PValue from Param where PTyp= 'BCHR' and PMCd= BCHR and     
       PSCd= '')= 'P'), 0)    
/* **** Zubin 211 **** */     
  if @row> 0     
  begin    select @errmsg= 'Production Qty Should Be >=  Bag Qty'    
   goto ErrHndlr    
  end     
    
/* **** Zubin 212 **** */     
    
/* *** Jay 2.13(CT) *** */     
  /* Current Partition */    
  select @row= count('x') from inserted     
   where OdPrtKey=@wCurrPrtn     
    and Exists (Select 'x' from Bag where BOdIdNo=OdIdNo and BCoCd = OdCoCd and BOdTc = OdTc and BOdYy = OdYy and     
      BOdChr = OdChr and BOdNo = OdNo and BOdSr = OdSr and BPrtKey=OdPrtKey and     
      (Select PValue from Param where PTyp= 'BCHR' and PMCd= BCHR and PSCd= '')= 'F'     
       Group By BChr having OdPrdQty-OdPrtFgQty< sum(BQty) )     
  if @row> 0     
  begin     select @errmsg= 'Production Qty Should Be >=  Bag Qty'     
   goto ErrHndlr    
  end    
    
  /* **** Zubin 212 **** */     
  /* *** Jay 2.13(CT) *** */ /*(DT: 18/12/06)*/    
  Select @row= count('x') from inserted     
   Where OdPrtKey=@wCurrPrtn and     
    Exists (Select 'x' from Bag where BOdIdNo=OdIdNo and BCoCd = OdCoCd and BOdTc = OdTc and BOdYy = OdYy and     
      BOdChr = OdChr and BOdNo = OdNo and BOdSr = OdSr and BPrtKey=OdPrtKey and     
      (select PValue from Param where PTyp= 'BCHR' and PMCd= BCHR and     
       PSCd= '')= 'C'     
      Group By BChr having OdPrdQty-OdPrtFgQty< sum(BQty))     
  if @row> 0     
  begin     select @errmsg= 'Production Qty Should Be >=  Bag Qty'     
   goto ErrHndlr    
  end    
  /* *** Jay 2.13(CT) *** */     
    
  /***********Geeta********Emr206**********************************/    
  /**********OnHld Check not required****** removed** and ins.OdHld= del.OdHld********************/    
  /*** Cannot change OnHld, PrdSeq, OrdEnt, PrdEnt when Bag Exists ***/    
  /* Current Partition */    
  select @row= count('x') from  deleted del where del.OdPRtKey=@wCurrPrtn and    
   not exists (select 'x' from inserted ins where ins.OdIdNo=del.OdIdNo and ins.OdCoCd= del.OdCoCd and ins.OdTc= del.OdTc and ins.OdYy= del.OdYy and ins.OdChr= del.OdChr and    
     ins.OdNo= del.OdNo and ins.OdSr= del.OdSr and ins.OdPrtKey=del.OdPrtKey and ins.OdOrdEnt= del.OdOrdEnt and ins.OdPrdEnt= del.OdPrdEnt    
      and ins.OdPrdSeq= del.OdPrdSeq) and     
   (exists (Select 'x' from Bag where BOdIdNo=OdIdNo and BCoCd= del.OdCoCd and BOdTc= del.OdTc     
     and BOdYy= del.OdYy and BOdChr= del.OdChr and BOdNo= del.OdNo     
     and BOdSr= del.OdSr and BPrtKey=del.OdPrtKey) Or OdPrtFgQty<>0 or OdPrtExpQty<>0)    
  if @row> 0     
  begin    
   select @errmsg= 'Cannot Change OrdEnt/ PrdEnt / PrdSeq When Bag Exists'    
   goto ErrHndlr    
  end    
    
  /******  OdOmCtCd  =  Blank or CmCd from CustMst (CmCtg= 'T')  ******/    
  /* Current Partition */    
  select @row= count('x') from inserted where OdPrtKey=@wCurrPrtn and OdOmCtCd<> '' and     
   not exists (select 'x' from CustMst where CmCtg= 'T' and CmCd= OdOmCtCd)    
  if @row> 0     
  begin     
   select @errmsg= 'Contractor Code Not Defined'    
   goto ErrHndlr    
  end    
    
/* ### Doubt ful   ( it is advisable to avoid Self updating triggers )   */    
/*** (Bef 2.12) (has been shifted to InvDsg_NR_InUpDl)    
  update OrdDsg set OdCls = (case when ins.OdOrdQty= ins.OdExpQty and Round(ins.OdOrdQty, 1)<> 0 then 'Y' else 'N' end)     
  from inserted ins, OrdDsg Od where ins.OdCoCd= Od.OdCoCd and ins.OdTc= Od.OdTc     
   and ins.OdYy= Od.OdYy and ins.OdChr= Od.OdChr and ins.OdNo= Od.OdNo and ins.OdSr= Od.OdSr    
 (Bef 2.12) ***/    
    
  /* Current Partition */    
  select @row= count('x') from inserted where OdPrtKey=@wCurrPrtn and OdBagPcs<= 0     
  if @row> 0     
  begin    
   select @errmsg= 'Default Bag Opening Pieces Cannot Be <= Zero'    
   goto ErrHndlr    
  end    
/* ### Doubt ful */    
 /*********Geeta****Crm**2.12****************************************/    
 /********If OmChr<>'WEB'  Then OdGrMet,OdGrDia,OdGrCs,OdGrFin should be Blank *******************************/    
/*  select @row= count('x') from inserted where (OdChr<>'WEB' and     
  (select PValue  from Param where PTyp='TC' and PMCd=OdTc and PSCd='')<>'PL' )  and     
  (OdGrMet<>'' Or OdGrDia<>'' Or OdGrCs<>'' Or OdGrFin<>'' )    
  if @row> 0     
  begin     
   select @errmsg= 'Grade Fields should be Blank for Chr<>''WEB''/OmTc<>''PL'' '    
   goto ErrHndlr    
  end    
*/      
 /********If OmChr='WEB' Then OdGrMet,OdGrDia,OdGrCs,OdGrFin should be  Valid     
 GrCd From Grade where GrCmCd in (OmCmcd,'ZSELF') and Blank Allowed*******************************/    
  /* Current Partition */    
  select @row= count('x') from inserted ins     
  Join OrdMst On OdOmIdNo=OmIdNo and OdCoCd=OmCoCd and OdTc=OmTc and OdYy=OmYy and OdChr=OmChr and OdNo=OmNo and OdPrtKey=OmPrtKey    
  where OdPrtKey=@wCurrPrtn and OdGrMet<>''     
   and Not Exists(Select 'x' From Param where PTyp= 'GRDCD' and PMCd= 'MET' and PSCd= OdGrMet)     
  if @row> 0     
  begin     
   select @errmsg= 'Metal Grade not Defined'    
   goto ErrHndlr    
  end    
    
  /* Current Partition */    
  select @row= count('x') from inserted ins    
  Join OrdMst On OdOmIdNo=OmIdNo and OdCoCd=OmCoCd and OdTc=OmTc and OdYy=OmYy and OdChr=OmChr and OdNo=OmNo and OdPrtKey=OmPrtKey    
  where OdPrtKey=@wCurrPrtn and OdGrDia<>'' and Not Exists(Select 'x' From Param where PTyp= 'GRDCD' and PMCd= 'DIA' and     
   PSCd= OdGrDia)     
  if @row> 0     
  begin     
   select @errmsg= 'Diamond Grade not Defined'    
   goto ErrHndlr    
  end    
    
  /********Geeta*Crm212**30/12/05**OmTCTyp='PL' and 'Crm' record exists in SysDb then OmCrmFixPrcYNYN=Valid PMCd ('YN') Else ''**********************/    
  /* Current Partition */    
  select @row= count('x') from inserted where OdPrtKey=@wCurrPrtn and    
  (select PValue  from Param where PTyp='TC' and PMCd=OdTc and PSCd='') ='PL'    
/*  and Exists (Select 'x' from SysDb where SysId='CRM') */     
   and Not Exists (Select 'x' From Param where PTyp='YN' and PMCd=OdCrmFixPrcYN and PSCd='')    
  if @row> 0    
  Begin    
   select @errmsg= 'Invalid Crm Fix Price (Enter Y/N?) '    
   goto ErrHndlr    
  end    
    
  /*****Geeta***Crm**212*CrmFixPrc Should be Blank if OmTcTyp<>'PL' or Not Exists 'CRM' Record in SysDb************/    
  /* Current Partition */    
  select @row= count('x') from inserted where OdPrtKey=@wCurrPrtn and OdCrmFixPrcYN<>'' and     
  (select PValue  from Param where PTyp='TC' and PMCd=OdTc and PSCd='') <>'PL' /* Or    
  not Exists (Select 'x' from SysDb where SysId='CRM') */     
  if @row> 0    
  Begin    
   select @errmsg= 'Crm Fix Price Should Be Blank '    
   goto ErrHndlr    
  end    
  
 /* Manali Trading Module */  
  if @wTcTyp = 'JT'   
  Begin   
   Select @Row=Count('x') from Inserted   
    Join Bag on OdPrtKey=BPrtKey and OdCoCd= BCoCd and OdByy= BYy and OdBChr= BChr and OdBNo= BNo  
    Where OdOmDt<BMaxTxnDt and   
    exists(Select 'x' from Param where PTyp= 'TC' and PMCd= OdTc and PValue='JT' and PMCd='JCN')  
   If @Row > 0  
   Begin  
    Select @ErrMsg= 'A Transaction of Later Date Exists For this Bag'  
    GoTo ErrHndlr  
   end  
      
   /* Manali Trading Module */  
   Select @Row=Count('x') from Inserted Ins  
    Join OrdMst Om On Om.OmCoCd=Ins.OdCoCd and Om.OmTc=Ins.OdTc and Om.OmChr=Ins.OdChr and Om.OmNo=Ins.OdNo and Om.OmIdNo=Ins.OdOmIdNo and Om.OmPrtKey=Ins.OdPrtKey     
    Where exists(Select 'x' from Param where PTyp= 'TC' and PMCd= OdTc and PValue='JT')  
    And exists(Select 'x' from TcMst Where TmCoCd=Ins.OdCoCd and TmTc=Ins.OdTc and TmBaseTcs<>'' and   
       (   
       (TmBaseMustYN='Y' and (Ins.OdBaseIdKey=0  Or    
       TmBaseTcs<>IsNull((Select Od.OdTc from OrdDsg Od   
            Join OrdMst Om1 On Om1.OmCoCd=Od.OdCoCd and Om1.OmTc=Od.OdTc and Om1.OmChr=Od.OdChr and Om1.OmNo=Od.OdNo and Om1.OmIdNo=Od.OdOmIdNo and Om1.OmPrtKey=Od.OdPrtKey    
            Where Od.OdPrtKey=Ins.OdPrtKey and Od.OdCoCd=(Case when ins.OdTc='JIN'  Then Om.OmBaseCoCd Else Ins.OdCoCd End)  
            and Od.OdDmCd=Ins.OdDmCd and Od.OdSfx=Ins.OdSfx and Od.OdDmSz=Ins.OdDmSz   
            and (Od.OdDmStkNo=Ins.OdDmStkNo/*Or (Od.OdDmStkNo=0 Or Ins.OdDmStkNo=0)*/)  
            and Om1.OmCmCtg=Om.OmCmCtg and Om1.OmCmCd=Om.OmCmCd  
            and Od.OdIdNo=Ins.OdBaseIdKey), '')))   
       Or  
       (TmBaseMustYN='N' and Ins.OdBaseIdKey<>0  AND    
       TmBaseTcs<>IsNull((Select Od.OdTc from OrdDsg Od   
            Join OrdMst Om1 On Om1.OmCoCd=Od.OdCoCd and Om1.OmTc=Od.OdTc and Om1.OmChr=Od.OdChr and Om1.OmNo=Od.OdNo and Om1.OmIdNo=Od.OdOmIdNo and Om1.OmPrtKey=Od.OdPrtKey    
            Where Od.OdPrtKey=Ins.OdPrtKey and Od.OdCoCd=Ins.OdCoCd  
            and Od.OdDmCd=Ins.OdDmCd and Od.OdSfx=Ins.OdSfx and Od.OdDmSz=Ins.OdDmSz   
            and (Od.OdDmStkNo=Ins.OdDmStkNo Or Od.OdDmStkNo=0)  
            and Om1.OmCmCtg=Om.OmCmCtg and Om1.OmCmCd=Om.OmCmCd  
            and Od.OdIdNo=Ins.OdBaseIdKey), ''))  
       )  
        )  
   If @Row > 0  
   Begin  
    Select @ErrMsg= 'Invalid Base Voucher'  
    GoTo ErrHndlr  
   end  
  
   Select @Row=Count('x') from Inserted Ins Where OdPrtKey=@wCurrPrtn  
    And exists(Select 'x' from Param where PTyp= 'TC' and PMCd= OdTc and PValue='JT')  
    And exists(Select 'x' from TcMst Where TmCoCd=Ins.OdCoCd and TmTc=Ins.OdTc and TmRefTcs<>'' and   
         Ins.OdRefIdKey<>0 And  
         ','+TmRefTcs+',' Not Like   
          '%,'+IsNull((Select OdTc from OrdDsg Od  
                Where OdPrtKey=Ins.OdPrtKey and OdCoCd=Ins.OdCoCd and OdIdNo=Ins.OdRefIdKey   
                and Od.OdDmCd=Ins.OdDmCd and Od.OdSfx=Ins.OdSfx and Od.OdDmSz=Ins.OdDmSz   
                /*and (Od.OdDmId=Ins.OdDmId Or Od.OdDmId=0 Or Ins.OdDmId=0)*/  
           ), '')+',%')  
   If @Row > 0  
   Begin  
    Select @ErrMsg= 'Invalid Reference Voucher'  
    GoTo ErrHndlr  
   end  
  
    select @row= count('x') from inserted where OdPrtKey=@wCurrPrtn and  
    exists(Select 'x' from Param Where PTyp='TC' and PMCd=OdTc and PValue='JT') and Round(OdPrdQty, 1) > 0    
    if @row> 0     
    begin    select @errmsg= 'Prd Qty should be 0 for Trading voucher'    
     goto ErrHndlr    
    end    
  End  
  
End    
    
    
 if @mode = 'D'   Begin    
  /*** (Jen 2.13) ***/    
  /****** Cannot delete OrdDsg record if ******/    
  /* select @row= count('x') from deleted where     
   exists(Select 'x' from OrdMst where OmCoCd = OdCoCd and OmTc = OdTc and     
    OmYy = OdYy and OmChr = OdChr and OmNo = OdNo and OmReNo> 0)    
  if @row> 0     
  begin    
   select @errmsg= 'Cannot Delete Selected Row(s) As Order Linked To Return Memo'     
   goto ErrHndlr     
  end */    
  /*** (Jen 2.13) ***/    
    
  /****** Cannot delete OrdDsg record if Export has taken place for that OrdSr or bags have been generated for the order ******/    
  /* Current Partition */    
  select @row= count('x') from deleted where OdPrtKey=@wCurrPrtn and    
   (OdExpQty<> 0 Or OdPrtExpQty<>0 Or OdPrtFgQty<>0 Or exists(Select 'x' from Bag where     
                BOdIdNo=OdIdNo and BCoCd = OdCoCd and BOdTc = OdTc and     
                BOdYy = OdYy and BOdChr = OdChr and     
                BOdNo = OdNo and BOdSr = OdSr and BPRtKey=OdPrtKey))    
  if @row> 0     
  begin    
   select @errmsg= 'Cannot Delete As Export Has Taken Place / Design Under Production'    
   goto ErrHndlr    
  end    
    
  /****** Cannot delete OrdDsg record if Order= Export Order Of Fgd Table  ******/    
  /* Current Partition */    
  select @row= count('x') from deleted where OdPrtKey=@wCurrPrtn and    
   (OdPrtExpQty<>0 Or exists(Select 'x' from Fgd where FdCoCd= OdCoCd and FdExpOdTc= OdTc and     
         FdExpOdYy = OdYy and FdExpOdChr = OdChr and FdExpOdNo = OdNo and     
         FdExpOdSr = OdSr and FdPrtKey=OdPrtKey))    
  if @row> 0     
  begin    
   select @errmsg= 'Cannot Delete As Record Exists As Export Order In Fgd Table'    
   goto ErrHndlr    
  end    
    
 /* ***** Manali Trding Module ***** */  
  if @wTcTyp = 'JT'   
  Begin   
   Select @row=Count('x') from deleted Del Where OdPrtKey=@wCurrPrtn and  
    exists(Select 'x' from OrdDsg Od Where  (Del.OdTc='JIN' Or Od.OdCoCd=Del.OdCoCd) and Od.OdPrtKey=Del.OdPrtKey and   
       Od.OdBaseIdKey=Del.OdIdNo)  
   If @row > 0  
   Begin  
    Select @errmsg='Cannot Delete as Order Serial Is Used as Base Voucher'  
    goto ErrHndlr  
   End  
    
   Select @row=Count('x') from deleted Del Where OdPrtKey=@wCurrPrtn and  
    exists(Select 'x' from OrdDsg Od Where Od.OdCoCd=Del.OdCoCd and Od.OdPrtKey=Del.OdPrtKey and   
       Od.OdRefIdKey=Del.OdIdNo)  
   If @row > 0  
   Begin  
    Select @errmsg='Cannot Delete as Order Serial Is Used as Refernce Voucher'  
    goto ErrHndlr  
   End  
  End   
 /* ***** Manali Trding Module ***** */  
  
  
  
  /********Geeta***Crm 2.12*****Moved to Not For Replication Trigger*****************************/    
  /****** Whenever the OrdDsg record is deleted, delete records from the child tables (OrdRm, OrdLab, OrdAmd, OrdTgt) ******/    
  /*    
  Delete OrdRm From deleted, OrdRm where OdCoCd = OrCoCd and OdTc = OrTc and     
   OdYy = OrYy and OdChr = OrChr and OdNo = OrNo and OdSr= OrSr     
    
  Delete OrdLab From deleted, OrdLab where OdCoCd = OlCoCd and OdTc = OlTc and     
   OdYy = OlYy and OdChr = OlChr and OdNo = OlNo and OdSr= OlSr     
    
  Delete OrdAmd From deleted, OrdAmd where OdCoCd = OaCoCd and OdTc = OaTc and     
   OdYy = OaYy and OdChr = OaChr and OdNo = OaNo and OdSr= OaSr     
    
  Delete OrdTgt From deleted, OrdTgt where OdCoCd = OtCoCd and OdTc = OtTc and     
   OdYy = OtYy and OdChr = OtChr and OdNo = OtNo and OdSr = OtSr     
    
  Delete OrdPln From deleted, OrdPln where OdCoCd = OpCoCd and OdTc = OpTc and     
   OdYy = OpYy and OdChr = OpChr and OdNo = OpNo and OdSr = OpSr     
  */    
 End    
   
 /* **** Manali Trading Module **** */  
 if @mode ='C'  
 Begin  
 if @wTcTyp = 'JT'   
  Begin  
  Select @row=Count('x') from deleted Del   
   Join TcMst On TmCoCd=Del.OdCoCd and TmTc=Del.OdTc  
   Where OdPrtKey=@wCurrPrtn and  
   exists(Select 'x' from OrdDsg Od Where  (Del.OdTc='JIN' Or  Od.OdCoCd=Del.OdCoCd) and Od.OdPrtKey=Del.OdPrtKey and   
      Od.OdBaseIdKey=Del.OdIdNo)  
   and not exists (Select 'x' from inserted ins where ins.OdIdNo=del.OdIdNo and ins.OdCoCd= del.OdCoCd     
    and ins.OdTc= del.OdTc and ins.OdYy= del.OdYy and ins.OdChr= del.OdChr     
    and ins.OdNo= del.OdNo and ins.OdSr= del.OdSr   
    and ins.OdDmCd=del.OdDmCd and ins.OdSfx=del.OdSfx and ins.OdDmSz=del.OdDmSz   and ins.OdDmStkYy = del.oddmstkyy and ins.oddmstkchr = del.oddmstkchr 
    and ins.OdDmStkNo=del.OdDmStkNo
    /*((TmBaseMustYn='Y' and ins.OdDmId=del.OdDmId) or (TmBaseMustYn='N' and (ins.OdDmId=0 or del.OdDmId=0) ))*/  
      )    
  If @row > 0  
  Begin  
   Select @errmsg='Cannot Change Design/Size/Suffix/DmId as Order Serial Is Used as Base Voucher'  
   goto ErrHndlr  
  End  
    
  Select @row=Count('x') from deleted Del Where OdPrtKey=@wCurrPrtn and  
   exists(Select 'x' from OrdDsg Od Where Od.OdCoCd=Del.OdCoCd and Od.OdPrtKey=Del.OdPrtKey and   
      Od.OdRefIdKey=Del.OdIdNo)  
   and not exists (Select 'x' from inserted ins where ins.OdIdNo=del.OdIdNo and ins.OdCoCd= del.OdCoCd     
     and ins.OdTc= del.OdTc and ins.OdYy= del.OdYy and ins.OdChr= del.OdChr     
     and ins.OdNo= del.OdNo and ins.OdSr= del.OdSr   
     and ins.OdDmCd=del.OdDmCd and ins.OdSfx=del.OdSfx and ins.OdDmSz=del.OdDmSz   
     and ((ins.OdDmStkyy = del.oddmstkyy and ins.oddmstkchr = del.oddmstkchr and ins.oddmstkno = del.oddmstkno and ins.OdDmStkNo=del.OdDmStkNo) or ins.OdDmStkNo=0 or del.OdDmStkNo=0))    
  If @row > 0  
  Begin  
   Select @errmsg='Cannot Change Design/Size/Suffix/DmId as Order Serial Is Used as Reference Voucher'  
   goto ErrHndlr  
  End  
 End  
 End  
 /* **** Manali Trading Module **** */  
  
 if @mode in ('C','D')  
 begin  
 /*** Jay 3.3(FG) ***/  
 /*** UPDATES FG In-Out ***/  
 /* **** Manali Trading Module **** */  
 /**** Manali 3.8.1 - Self Payable Considered In  Calculation *****/  
 /*If exists (Select 'x' from Deleted   
   where exists (Select 'x' from Param where PTyp= 'TC'   
      and (PValue = 'FG' Or (PValue='JT' and exists(Select 'x' from TcMst Where TmCoCd=OdCoCd and TmTc=OdTc and TmLocTypFr<>'' and TmLocTypTo<>'' )))   
      and PMCd= OdTc))*/  
 If @wTcTyp = 'FG' Or @wTcTyp = 'JT'         
 BEgin  
  /*** Update InvHd ***/  
  update InvHd set   
  @wInFob= convert (decimal (16, 2), InFOB- IsNull((Select sum(del.OdSalPrc*del.OdOrdQty) from deleted del where InPrtKey=del.OdPrtKey and del.OdCoCd= InCoCd and del.OdTc= InTc and del.OdYy= InYy   
   and del.OdChr= InChr and del.OdNo= InNo), 0)),  
  @wInFobRs= convert (decimal (16, 0), (@wInFob)  
    * (case when InRsMulDiv= 'M' then inCnvRt else 1 end)   
    / (case when InRsMulDiv= 'D' then inCnvRt else 1 end) ),   
  InFOB= @wInFob,   
  InFOBRs= @wInFobRs,   
  @wInLclOctroiRs= convert (decimal (16, 2), (Case when InLclInvYn = 'Y' then   
       (@wInFob * (InCmOctroiPer / 100)) else 0 end)),  
  @wInLclVatRs= convert (decimal (16, 2), (Case when InLclInvYn = 'Y' then   
       ((@wInFOB * (1 + (InCmOctroiPer/100 ))+InLclOthrValRs)* (InHVatPer /100 )) else 0 end)),  
  @wInLclCstRs= convert (decimal (16, 2), (Case when InLclInvYn = 'Y' then   
       ((@wInFOB * (1 + (InCmOctroiPer/100 ))+InLclOthrValRs)* (InHCstPer /100 )) else 0 end)),  
  InLclOctroiRs= @wInLclOctroiRs,   
  InLclVatRs= @wInLclVatRs,   
  InLclCstRs= @wInLclCstRs,   
  InTot= convert (decimal (16, 2), (Case InLclInvYn   
     when 'N' then @wInFOB + (case InTotAs when 'FOB' then (case when InHSelfDscYN ='Y' then 0 Else InDsc End) +(case when InHSelfDeductYN ='Y' then 0 Else InDeduct End)  
           when 'CIF' then (case when InHHideCommYN ='Y' then 0 Else InComm End)+ InFrt+ InIns +(case when InHSelfDscYN ='Y' then 0 Else InDsc End) +(case when InHSelfDeductYN ='Y' then 0 Else InDeduct End)  
           when 'C&F' then (case when InHHideCommYN ='Y' then 0 Else InComm End)+ InFrt +(case when InHSelfDscYN ='Y' then 0 Else InDsc End) + (case when InHSelfDeductYN ='Y' then 0 Else InDeduct End)  
           when 'C&I' then (case when InHHideCommYN ='Y' then 0 Else InComm End)+ InIns +(case when InHSelfDscYN ='Y' then 0 Else InDsc End) + (case when InHSelfDeductYN ='Y' then 0 Else InDeduct End)  
           else 0 end)  
     when 'Y' then  @wInFob+ @wInLclOctroiRs+ @wInLclVatRs+ @wInLclCstRs+ InLclOthrValRs   
     else 0 end)),    InTotRs= (Case InLclInvYn   
     when 'N' then @wInFOBRs + (case InTotAs when 'FOB' then (case when InHSelfDscYN ='Y' then 0 Else InDscRs End) +(case when InHSelfDeductYN ='Y' then 0 Else InDeductRs End)  
            when 'CIF' then (case when InHHideCommYN ='Y' then 0 Else InCommRs End)+ InFrtRs+ InInsRs +(case when InHSelfDscYN ='Y' then 0 Else InDscRs End) +(case when InHSelfDeductYN ='Y' then 0 Else InDeductRs End)  
            when 'C&F' then (case when InHHideCommYN ='Y' then 0 Else InCommRs End)+ InFrtRs +(case when InHSelfDscYN ='Y' then 0 Else InDscRs End) + (case when InHSelfDeductYN ='Y' then 0 Else InDeductRs End)  
            when 'C&I' then (case when InHHideCommYN ='Y' then 0 Else InCommRs End)+ InInsRs +(case when InHSelfDscYN ='Y' then 0 Else InDscRs End) + (case when InHSelfDeductYN ='Y' then 0 Else InDeductRs End)  
            else 0 end)  
     when 'Y' then convert (decimal (16, 0),   
      (@wInFob+ @wInLclOctroiRs+ @wInLclVatRs+ @wInLclCstRs+ InLclOthrValRs)   
      * (case when InRsMulDiv= 'M' then inCnvRt else 1 end)   
      / (case when InRsMulDiv= 'D' then inCnvRt else 1 end) )  
     else 0 end)       
  from deleted   
  Join OrdMst On OmCoCd=OdCoCd and OmTc=OdTc and OmChr=odChr and OmNo=OdNo and OmIdNo=OdOmIdNo and OmPrtKey=OdPrtKey  
  where InPrtKey=OdPrtKey and InCoCd= OdCoCd and InTc= OdTc and InYy= OdYy and InChr= OdChr and InNo= OdNo   
      and (exists (Select 'x' from Param where PTyp= 'TCTYP' and PMCD In ('FG', 'JT') and PValue= OdTc)  
         or   
        exists (Select 'x' from Param where PTyp= 'TCTYP' and   
            PMCD = (Select PValue from Param Where PTyp='TC' and PMCd= OdTc)))  
  /*and OmFrDmLoc <>'' and OmToDmLoc<>''*/  
  
  
  /*** Update DmYyMm from OmFrDmLoc ***/  
  /* Current Partition */  
  update DmYyMm set  
  DDmCrQty= DDmCrQty- (select isnull(sum(Del.OdOrdQty), 0) from Deleted Del  
   Join OrdMst on OmPrtKey=Del.OdPrtKey and OmIdNo=Del.OdOmIdNo and OmCoCd= Del.OdCoCd and OmTc= Del.OdTc and OmYy= Del.OdYy and OmChr= Del.OdChr  
   and OmNo= Del.ODNo where OmFrDmDc='C' and Del.OdPrtKey=DPrtKey and Del.OdCoCd= DCoCd and Del.OdDmCd= DDmCd   
   and Del.OdSfx= DOdSfx and Del.OdDmSz= DDmSz and Del.OdDmStkYy = DDmStkYy and Del.OdDmStkChr = DDmStkChr and Del.OdDmStkNo = DDmStkNo
   and OmFrDmLoc=DLoc  
   and convert(int,convert(varchar(6),OmDt,112))= DYyMm),  
  DDmDrQty= DDmDrQty-(select isnull(sum(Del.OdOrdQty),0) from Deleted Del  
   Join OrdMst on OmPrtKey=Del.OdPrtKey and OmIdNo=Del.OdOmIdNo and OmCoCd= Del.OdCoCd and OmTc= Del.OdTc and OmYy= Del.OdYy and OmChr= Del.OdChr  
   and OmNo= Del.OdNo where OmFrDmDc='D' and Del.OdPrtKey=DPrtKey and Del.OdCoCd= DCoCd and Del.OdDmCd=DDmCd   
   and Del.OdSfx= DOdSfx and Del.OdDmSz= DDmSz and Del.OdDmStkNo= DDmStkNo and Del.OdDmStkyy = DDmStkYy and del.OdDmStkChr = 
			DDmStkChr and OmFrDmLoc=DLoc  
   and convert(int,convert(varchar(6),OmDt,112))=DYyMm)  
  from Deleted Join OrdMst on OmPrtKey=Deleted.OdPrtKey and OmIdNo=Deleted.OdOmIdNo and OmCoCd= Deleted.OdCoCd and OmTc= Deleted.OdTc and OmYy= Deleted.OdYy  
   and OmChr= Deleted.OdChr and OmNo= Deleted.OdNo  
   where Deleted.OdPrtKey=DPrtKey and Deleted.OdCoCd= DCoCd and Deleted.OdDmCd= DDmCd    
   and Deleted.OdSfx= DOdSfx and Deleted.OdDmSz= DDmSz and Deleted.OdDmStkNo= DDmStkNo and Deleted.OdDmStkYy= DDmStkYy and Deleted.OdDmStkChr = DDmStkChr and OmFrDmLoc= DLoc  
   and convert(int,convert(varchar(6),OmDt,112))=DYyMm  
   and (OmFrDmDc='D' or OmFrDmDc='C')  
   and OmFrDmLoc <>'' and OmToDmLoc<>''  
 End  
  
 /*If exists (Select 'x' from Deleted   
   where exists (Select 'x' from Param where PTyp= 'TC' and PMCd= OdTc and PValue='JT' and PMCd='JCN'))*/  
 if @wTcTyp = 'JT'   
 Begin    
  update SYyMm set  
  SRmCrQty= SRmCrQty- (select isnull(sum(FrRmQty * (Case When FrFrRmLoc='FG' and FrFrRmDc= 'D' Then 1 When FrFrRmLoc='FG' and FrFrRmDc= 'C' Then -1 When FrToRmLoc='FG' and FrToRmDc= 'D' Then 1 When FrToRmLoc='FG' and FrToRmDc= 'C' Then -1 Else 0 End)), 0
)   
   From FgRm   
   Join Fgd on FdPrtKey=FrPrtKey and FdCoCd= FrCoCd and FdTc= FrTc and FdYy= FrYy and FdChr= FrChr   
   and FdNo= FrNo and FdSr=FrSr    
   Join Deleted Del On Del.OdPrtKey=FdPrtKey and Del.OdCoCd=FdCoCd and Del.OdBYy=FdBYy and Del.OdBChr=FdBChr and Del.OdBNo=FdBNo  
   Join OrdMst On OmPrtKey=Del.OdPrtKey and OmCoCd=Del.OdCoCd and OmTc=Del.OdTc and OmYy=Del.OdYy and OmChr=Del.OdChr and OmNo=Del.OdNo  
   where OmFrRmDc= 'C' and FrPrtKey=SPrtKey and FrCoCd= SCoCd and FrRmCd=SRmCd   
   and FrLotNo=SLotNo and FrRmSz=SRmSz and FrRmStkRt=SRmStkRt and OmFrRmLoc=SLoc   
   and convert(int,convert(varchar(6),OmDt,112))= SYyMm and OdTc= 'JCN' ),  
  SRmCrWt= SRmCrWt-(select isnull(sum(FrRmWt *(Case When FrFrRmLoc='FG' and FrFrRmDc= 'D' Then 1 When FrFrRmLoc='FG' and FrFrRmDc= 'C' Then -1 When FrToRmLoc='FG' and FrToRmDc= 'D' Then 1 When FrToRmLoc='FG' and FrToRmDc= 'C' Then -1 Else 0 End)), 0)   
   from FgRm  
   Join Fgd on FdPrtKey=FrPrtKey and FdCoCd= FrCoCd and FdTc= FrTc and FdYy= FrYy and FdChr= FrChr   
   and FdNo= FrNo and FdSr=FrSr    
   Join Deleted Del On Del.OdPrtKey=FdPrtKey and Del.OdCoCd=FdCoCd and Del.OdBYy=FdBYy and Del.OdBChr=FdBChr and Del.OdBNo=FdBNo  
   Join OrdMst On OmPrtKey=Del.OdPrtKey and OmCoCd=Del.OdCoCd and OmTc=Del.OdTc and OmYy=Del.OdYy and OmChr=Del.OdChr and OmNo=Del.OdNo  
   where OmFrRmDc= 'C' and FrPrtKey=SPrtKey and FrCoCd= SCoCd and FrRmCd=SRmCd   
   and FrLotNo=SLotNo and FrRmSz=SRmSz and FrRmStkRt=SRmStkRt and OmFrRmLoc=SLoc   
   and convert(int,convert(varchar(6),OmDt,112))= SYyMm and OdTc= 'JCN') ,  
  SRmDrQty= SRmDrQty-(select isnull(sum(FrRmQty * (Case When FrFrRmLoc='FG' and FrFrRmDc= 'D' Then 1 When FrFrRmLoc='FG' and FrFrRmDc= 'C' Then -1 When FrToRmLoc='FG' and FrToRmDc= 'D' Then 1 When FrToRmLoc='FG' and FrToRmDc= 'C' Then -1 Else 0 End)), 0) 
  
   From FgRm   
   Join Fgd on FdPrtKey=FrPrtKey and FdCoCd= FrCoCd and FdTc= FrTc and FdYy= FrYy and FdChr= FrChr   
   and FdNo= FrNo and FdSr=FrSr    
   Join Deleted Del On Del.OdPrtKey=FdPrtKey and Del.OdCoCd=FdCoCd and Del.OdBYy=FdBYy and Del.OdBChr=FdBChr and Del.OdBNo=FdBNo  
   Join OrdMst On OmPrtKey=Del.OdPrtKey and OmCoCd=Del.OdCoCd and OmTc=Del.OdTc and OmYy=Del.OdYy and OmChr=Del.OdChr and OmNo=Del.OdNo  
   where OmFrRmDc= 'D' and FrPrtKey=SPrtKey and FrCoCd= SCoCd and FrRmCd=SRmCd   
   and FrLotNo=SLotNo and FrRmSz=SRmSz and FrRmStkRt=SRmStkRt and OmFrRmLoc=SLoc   
   and convert(int,convert(varchar(6),OmDt,112))= SYyMm and OdTc= 'JCN') ,  
  SRmDrWt= SRmDrWt-(select isnull(sum(FrRmWt * (Case When FrFrRmLoc='FG' and FrFrRmDc= 'D' Then 1 When FrFrRmLoc='FG' and FrFrRmDc= 'C' Then -1 When FrToRmLoc='FG' and FrToRmDc= 'D' Then 1 When FrToRmLoc='FG' and FrToRmDc= 'C' Then -1 Else 0 End)), 0)   
   from FgRm   
   Join Fgd on FdPrtKey=FrPrtKey and FdCoCd= FrCoCd and FdTc= FrTc and FdYy= FrYy and FdChr= FrChr   
   and FdNo= FrNo and FdSr=FrSr    
   Join Deleted Del On Del.OdPrtKey=FdPrtKey and Del.OdCoCd=FdCoCd and Del.OdBYy=FdBYy and Del.OdBChr=FdBChr and Del.OdBNo=FdBNo  
   Join OrdMst On OmPrtKey=Del.OdPrtKey and OmCoCd=Del.OdCoCd and OmTc=Del.OdTc and OmYy=Del.OdYy and OmChr=Del.OdChr and OmNo=Del.OdNo  
   where OmFrRmDc= 'D' and FrPrtKey=SPrtKey and FrCoCd= SCoCd and FrRmCd=SRmCd   
   and FrLotNo=SLotNo and FrRmSz=SRmSz and FrRmStkRt=SRmStkRt and OmFrRmLoc=SLoc   
   and convert(int,convert(varchar(6),OmDt,112))= SYyMm and OdTc= 'JCN')   
  from FgRm   
  Join Fgd on FdPrtKey=FrPrtKey and FdCoCd= FrCoCd and FdTc= FrTc and FdYy= FrYy and FdChr= FrChr and FdNo= FrNo and FdSr=FrSr    
  Join Deleted On Deleted.OdPrtKey=FdPrtKey and Deleted.OdCoCd=FdCoCd and Deleted.OdBYy=FdBYy and Deleted.OdBChr=FdBChr and Deleted.OdBNo=FdBNo  
  Join OrdMst On OmPrtKey=Deleted.OdPrtKey and OmCoCd=Deleted.OdCoCd and OmTc=Deleted.OdTc and OmYy=Deleted.OdYy and OmChr=Deleted.OdChr and OmNo=Deleted.OdNo  
  where FrPrtKey=SPrtKey and FrCoCd= SCoCd and FrRmCd= SRmCd and FrLotNo= SLotNo   
  and FrRmSz= SRmSz and FrRmStkRt= SRmStkRt   
  and OmFrRmLoc= SLoc and convert(int,convert(varchar(6),OmDt,112))=SYyMm  
  and (OmFrRmDc='D'  or OmFrRmDc='C')  
  and OdTc= 'JCN'  
     
  if @mode = 'D'    Begin  
   /******* (BCls= 'N') Bag is again opened for transaction ******/  
   update Bag set  
   BCls= 'N', BClsDt= '01/01/80', BCnvJtYn='N'  
   from deleted where BPrtKey=@wCurrPrtn and BCoCd= OdCoCd and BYy= OdBYy and BChr= OdBChr and BNo= OdBNo    
   and OdTc= 'JCN'   
  End    
 End  
 End   
   
 If @mode in ('A','C')  
 begin  
 /*** Jay 3.3(FG) ***/  
 /*** UPDATES FG In-Out ***/  
 /* **** Manali Trading Module **** */  
 /**** Manali 3.8.1 - Self Payable Considered In  Calculation *****/  
 /*If exists (Select 'x' from Inserted   
   where exists (Select 'x' from Param where PTyp= 'TC'   
       and (PValue = 'FG' Or (PValue='JT' and exists(Select 'x' from TcMst Where TmCoCd=OdCoCd and TmTc=OdTc and TmLocTypFr<>'' and TmLocTypTo<>'' )))   
       and PMCd= OdTc))*/  
 If @wTcTyp = 'FG' or @wTcTyp = 'JT'   
 Begin  
  /*** Update InvHd ***/  
  update InvHd set  
  @wInFob = convert (decimal (16, 2), InFOB+ IsNull((Select sum(ins.OdSalPrc*ins.OdOrdQty) from inserted ins where   
   ins.OdPrtKey=InPrtKey and ins.OdCoCd= InCoCd and ins.OdTc= InTc and ins.OdYy= InYy   
   and ins.OdChr= InChr and ins.OdNo= InNo), 0)),   
  @wInFobRs= convert (decimal (16, 0), (@wInFob)  
    * (case when InRsMulDiv= 'M' then inCnvRt else 1 end)   
    / (case when InRsMulDiv= 'D' then inCnvRt else 1 end) ),   
  InFOB= @wInFob,   
  InFOBRs= @wInFobRs,   
  @wInLclOctroiRs= convert (decimal (16, 2), (Case when InLclInvYn = 'Y' then   
       (@wInFob * (InCmOctroiPer / 100)) else 0 end)),  
  @wInLclVatRs= convert (decimal (16, 2), (Case when InLclInvYn = 'Y' then   
       ((@wInFOB * (1 + (InCmOctroiPer/100 ))+InLclOthrValRs)* (InHVatPer /100 )) else 0 end)),  
  @wInLclCstRs= convert (decimal (16, 2), (Case when InLclInvYn = 'Y' then   
       ((@wInFOB * (1 + (InCmOctroiPer/100 ))+InLclOthrValRs)* (InHCstPer /100 )) else 0 end)),  
  InLclOctroiRs= @wInLclOctroiRs,   
  InLclVatRs= @wInLclVatRs,   
  InLclCstRs= @wInLclCstRs,   
  InTot= convert (decimal (16, 2),   
     (Case InLclInvYn   
     when 'N' then @wInFOB + (case InTotAs when 'FOB' then (case when InHSelfDscYN ='Y' then 0 Else InDsc End) +(case when InHSelfDeductYN ='Y' then 0 Else InDeduct End)  
           when 'CIF' then (case when InHHideCommYN ='Y' then 0 Else InComm End)+ InFrt+ InIns +(case when InHSelfDscYN ='Y' then 0 Else InDsc End) +(case when InHSelfDeductYN ='Y' then 0 Else InDeduct End)  
           when 'C&F' then (case when InHHideCommYN ='Y' then 0 Else InComm End)+ InFrt +(case when InHSelfDscYN ='Y' then 0 Else InDsc End) + (case when InHSelfDeductYN ='Y' then 0 Else InDeduct End)  
           when 'C&I' then (case when InHHideCommYN ='Y' then 0 Else InComm End)+ InIns +(case when InHSelfDscYN ='Y' then 0 Else InDsc End) + (case when InHSelfDeductYN ='Y' then 0 Else InDeduct End)  
           else 0 end)  
     when 'Y' then @wInFob+ @wInLclOctroiRs+ @wInLclVatRs+ @wInLclCstRs+ InLclOthrValRs   
     else 0 end)),  
  InTotRs= (Case InLclInvYn   
     when 'N' then @wInFOBRs + (case InTotAs when 'FOB' then (case when InHSelfDscYN ='Y' then 0 Else InDscRs End) +(case when InHSelfDeductYN ='Y' then 0 Else InDeductRs End)  
            when 'CIF' then (case when InHHideCommYN ='Y' then 0 Else InCommRs End)+ InFrtRs+ InInsRs +(case when InHSelfDscYN ='Y' then 0 Else InDscRs End) +(case when InHSelfDeductYN ='Y' then 0 Else InDeductRs End)  
            when 'C&F' then (case when InHHideCommYN ='Y' then 0 Else InCommRs End)+ InFrtRs +(case when InHSelfDscYN ='Y' then 0 Else InDscRs End) + (case when InHSelfDeductYN ='Y' then 0 Else InDeductRs End)  
            when 'C&I' then (case when InHHideCommYN ='Y' then 0 Else InCommRs End)+ InInsRs +(case when InHSelfDscYN ='Y' then 0 Else InDscRs End) + (case when InHSelfDeductYN ='Y' then 0 Else InDeductRs End)  
            else 0 end)  
     when 'Y' then convert (decimal (16, 0),   
      (@wInFob+ @wInLclOctroiRs+ @wInLclVatRs+ @wInLclCstRs+ InLclOthrValRs)   
      * (case when InRsMulDiv= 'M' then inCnvRt else 1 end)   
      / (case when InRsMulDiv= 'D' then inCnvRt else 1 end) )  
     else 0 end)  
  from inserted   
  Join OrdMst On OmCoCd=OdCoCd and OmTc=OdTc and OmChr=odChr and OmNo=OdNo and OmIdNo=OdOmIdNo and OmPrtKey=OdPrtKey  
  where InPrtKey=OdPrtKey and InCoCd= OdCoCd and InTc= OdTc and InYy= OdYy and InChr= OdChr and InNo= OdNo   
      and (exists (Select 'x' from Param where PTyp= 'TCTYP' and PMCD ='FG' and PValue= OdTc)  
        or   
        exists (Select 'x' from Param where PTyp= 'TCTYP' and   
            PMCD = (Select PValue from Param Where PTyp='TC' and PMCd= OdTc)))  
  /*and OmFrDmLoc <>'' and OmToDmLoc<>''*/  
  
  
  /* Current Partition */  
  insert into DmYyMm(DPrtKey,DCoCd,DDmCd,DOdSfx,DDmSz,DDmStkYy,DDmStkChr, DDmStkNo, DLoc,DYyMm,DDmDrQty,DDmCrQty)  
   select distinct OdPrtKey, OdCoCd, OdDmCd, OdSfx, OdDmSz,OdDmStkYy,OdDmStkChr,OdDmStkNo,OmFrDmLoc,convert(int,convert(varchar(6),OmDt,112)),0,0  
   from inserted join OrdMst on OmPrtKey=OdPrtKey and OmIdNo=OdOmIdNo and OmCoCd=OdCoCd and OmTc=OdTc and OmYy=OdYy and OmChr=OdChr and OmNo=OdNo  
   where not exists(select 'x' from  DmYyMm where DPrtKey=OdPrtKey and DCoCd= OdCoCd and DDmCd=OdDmCd   
        and OdSfx= DOdSfx and OdDmSz= DDmSz and OdDmStkYy= DDmStkYy and OdDmStkChr = DDmStkChr and OdDmStkNo = DDmStkNo and DLoc=OmFrDmLoc and  
        DYyMm= convert(int,convert(varchar(6),OmDt,112)))  
    and OmFrDmLoc <>'' and OmToDmLoc<>''  
      
  /* Current Partition */   
  update DmYyMm set  
  DDmCrQty=DDmCrQty+(select isnull(sum(Ins.OdOrdQty),0) from Inserted Ins Join OrdMst   
   on OmPrtKey=Ins.OdPrtKey and OmIdNo=Ins.OdOmIdNo and OmCoCd= Ins.OdCoCd and OmTc= Ins.OdTc and OmYy= Ins.OdYy and OmChr= Ins.OdChr and OmNo= Ins.OdNo  
   where OmFrDmDc='C' and Ins.OdPrtKey=DPrtKey and Ins.OdCoCd= DCoCd and Ins.OdDmCd=DDmCd   
   and Ins.OdSfx= DOdSfx and Ins.OdDmSz= DDmSz and Ins.OdDmStkNo = DDmStkNo and Ins.OdDmStkYy = DDmStkYy and Ins.OdDmStkChr = DDmStkChr 
   and OmFrDmLoc=DLoc  
   and convert(int,convert(varchar(6),OmDt,112))=DYyMm),  
  DDmDrQty=DDmDrQty+(select isnull(sum(Ins.OdOrdQty),0) from Inserted Ins Join OrdMst   
   on OmPrtKey=Ins.OdPrtKey and OmIdNo=Ins.OdOmIdNo and OmCoCd= Ins.OdCoCd and OmTc= Ins.OdTc and OmYy= Ins.OdYy and OmChr= Ins.OdChr and OmNo= Ins.OdNo  
   where OmFrDmDc='D'  and Ins.OdPrtKey=DPrtKey and Ins.OdCoCd= DCoCd and Ins.OdDmCd=DDmCd   
   and Ins.OdSfx= DOdSfx and Ins.OdDmSz= DDmSz and Ins.OdDmStkNo= DDmStkNo and Ins.OdDmStkYy = DDmStkYy and Ins.OdDmStkChr = DDmStkChr 
   and OmFrDmLoc=DLoc  
   and convert(int,convert(varchar(6),OmDt,112))=DYyMm)  
  from Inserted Join OrdMst on OmPrtKey=Inserted.OdPrtKey and OmIdNo=Inserted.OdOmIdNo and OmCoCd= Inserted.OdCoCd and OmTc= Inserted.OdTc  
   and OmYy= Inserted.OdYy and OmChr= Inserted.OdChr and OmNo= Inserted.OdNo  
   where Inserted.OdPrtKey=DPrtKey and Inserted.OdCoCd= DCoCd and Inserted.OdDmCd= DDmCd   
   and Inserted.OdSfx= DOdSfx and Inserted.OdDmSz= DDmSz and Inserted.OdDmStkNo= DDmStkNo 
   and Inserted.OdDmStkYy = DDmStkYy and Inserted.OdDmStkChr = DDmStkChr and OmFrDmLoc= DLoc  
   and convert(int,convert(varchar(6),OmDt,112))=DYyMm  
   and (OmFrDmDc='D' or OmFrDmDc='C')  
   and OmFrDmLoc <>'' and OmToDmLoc<>''  
 end  
  
 /*If exists (Select 'x' from Inserted   
   where exists (Select 'x' from Param where PTyp= 'TC' and PMCd= OdTc and PValue='JT' and PMCd='JCN'))*/  
 If @wTcTyp = 'JT'   
 Begin  
  Insert into SYyMm(SPrtKey,SCoCd,SRmCd,SLotNo,SRmSz,SRmStkRt,SLoc,SYyMm,SRmDrqty,SRmDrWt,SRmCrQty,SRmCrWt)   
   select distinct FrPrtKey,FrCoCd,FrRmCd,FrLotNo,FrRmSz,FrRmStkRt,OmFrRmLoc,convert(int,convert(varchar(6),OmDt,112)),0,0,0,0  
   From FgRm  
   Join Fgd on FdPrtKey=FrPrtKey and FdCoCd= FrCoCd and FdTc= FrTc and FdYy= FrYy and FdChr= FrChr   
   and FdNo= FrNo and FdSr=FrSr    
   Join Inserted On OdPrtKey=FdPrtKey and OdCoCd=FdCoCd and OdBYy=FdBYy and OdBChr=FdBChr and OdBNo=FdBNo  
   Join OrdMst On OmPrtKey=OdPrtKey and OmCoCd=OdCoCd and OmTc=OdTc and OmYy=OdYy and OmChr=OdChr and OmNo=OdNo  
   where not exists(select 'x' from  SYyMm where SPrtKey=FrPrtKey and SCoCd= FrCoCd and SRmCd=FrRmCd and SLotNo=FrLotNo   
   and SRmSZ=FrRmSz and SRmStkRt=FrRmStkRt and SLoc=OmFrRmLoc and  
   SYyMm= convert(int,convert(varchar(6),OmDt,112))) and OdTc= 'JCN'   
     
  
  /* Current Partition */  
  Update SYyMm set  
  SRmCrQty= SRmCrQty+(select isnull(sum(FrRmQty * (Case When FrFrRmLoc='FG' and FrFrRmDc= 'D' Then 1 When FrFrRmLoc='FG' and FrFrRmDc= 'C' Then -1 When FrToRmLoc='FG' and FrToRmDc= 'D' Then 1 When FrToRmLoc='FG' and FrToRmDc= 'C' Then -1 Else 0 End)), 0) 
  
   from FgRm  
   Join Fgd on FdPrtKey=FrPrtKey and FdCoCd= FrCoCd and FdTc= FrTc and FdYy= FrYy and FdChr= FrChr   
   and FdNo= FrNo and FdSr=FrSr    
   Join Inserted Ins On Ins.OdPrtKey=FdPrtKey and Ins.OdCoCd=FdCoCd and Ins.OdBYy=FdBYy and Ins.OdBChr=FdBChr and Ins.OdBNo=FdBNo  
   Join OrdMst On OmPrtKey=Ins.OdPrtKey and OmCoCd=Ins.OdCoCd and OmTc=Ins.OdTc and OmYy=Ins.OdYy and OmChr=Ins.OdChr and OmNo=Ins.OdNo  
   where OmFrRmDc= 'C' and FrPrtKey=SPrtKey and FrCoCd= SCoCd and FrRmCd=SRmCd   
   and FrLotNo=SLotNo and FrRmSz=SRmSz and FrRmStkRt=SRmStkRt and OmFrRmLoc=SLoc   
   and convert(int,convert(varchar(6),OmDt,112))= SYyMm and OdTc= 'JCN') ,  
  SRmCrWt=SRmCrWt+(select isnull(sum(FrRmWt * (Case When FrFrRmLoc='FG' and FrFrRmDc= 'D' Then 1 When FrFrRmLoc='FG' and FrFrRmDc= 'C' Then -1 When FrToRmLoc='FG' and FrToRmDc= 'D' Then 1 When FrToRmLoc='FG' and FrToRmDc= 'C' Then -1 Else 0 End)), 0)   
   from FgRm   
   Join Fgd on FdPrtKey=FrPrtKey and FdCoCd= FrCoCd and FdTc= FrTc and FdYy= FrYy and FdChr= FrChr   
   and FdNo= FrNo and FdSr=FrSr    
   Join Inserted Ins On Ins.OdPrtKey=FdPrtKey and Ins.OdCoCd=FdCoCd and Ins.OdBYy=FdBYy and Ins.OdBChr=FdBChr and Ins.OdBNo=FdBNo  
   Join OrdMst On OmPrtKey=Ins.OdPrtKey and OmCoCd=Ins.OdCoCd and OmTc=Ins.OdTc and OmYy=Ins.OdYy and OmChr=Ins.OdChr and OmNo=Ins.OdNo  
   where OmFrRmDc= 'C' and FrPrtKey=SPrtKey and FrCoCd= SCoCd and FrRmCd=SRmCd   
   and FrLotNo=SLotNo and FrRmSz=SRmSz and FrRmStkRt=SRmStkRt and OmFrRmLoc=SLoc   
   and convert(int,convert(varchar(6),OmDt,112))= SYyMm and OdTc= 'JCN') ,  
  SRmDrQty=SRmDrQty+(select isnull(sum(FrRmQty * (Case When FrFrRmLoc='FG' and FrFrRmDc= 'D' Then 1 When FrFrRmLoc='FG' and FrFrRmDc= 'C' Then -1 When FrToRmLoc='FG' and FrToRmDc= 'D' Then 1 When FrToRmLoc='FG' and FrToRmDc= 'C' Then -1 Else 0 End)), 0)  
 
   from FgRm  
   Join Fgd on FdPrtKey=FrPrtKey and FdCoCd= FrCoCd and FdTc= FrTc and FdYy= FrYy and FdChr= FrChr   
   and FdNo= FrNo and FdSr=FrSr    
   Join Inserted Ins On Ins.OdPrtKey=FdPrtKey and Ins.OdCoCd=FdCoCd and Ins.OdBYy=FdBYy and Ins.OdBChr=FdBChr and Ins.OdBNo=FdBNo  
   Join OrdMst On OmPrtKey=Ins.OdPrtKey and OmCoCd=Ins.OdCoCd and OmTc=Ins.OdTc and OmYy=Ins.OdYy and OmChr=Ins.OdChr and OmNo=Ins.OdNo  
   where OmFrRmDc= 'D' and FrPrtKey=SPrtKey and FrCoCd= SCoCd and FrRmCd=SRmCd   
   and FrLotNo=SLotNo and FrRmSz=SRmSz and FrRmStkRt=SRmStkRt and OmFrRmLoc=SLoc   
   and convert(int,convert(varchar(6),OmDt,112))= SYyMm and OdTc= 'JCN') ,  
  SRmDrWt=SRmDrWt+(select isnull(sum(FrRmWt * (Case When FrFrRmLoc='FG' and FrFrRmDc= 'D' Then 1 When FrFrRmLoc='FG' and FrFrRmDc= 'C' Then -1 When FrToRmLoc='FG' and FrToRmDc= 'D' Then 1 When FrToRmLoc='FG' and FrToRmDc= 'C' Then -1 Else 0 End)), 0)   
   from FgRm   
   Join Fgd on FdPrtKey=FrPrtKey and FdCoCd= FrCoCd and FdTc= FrTc and FdYy= FrYy and FdChr= FrChr   
   and FdNo= FrNo and FdSr=FrSr    
   Join Inserted Ins On Ins.OdPrtKey=FdPrtKey and Ins.OdCoCd=FdCoCd and Ins.OdBYy=FdBYy and Ins.OdBChr=FdBChr and Ins.OdBNo=FdBNo  
   Join OrdMst On OmPrtKey=Ins.OdPrtKey and OmCoCd=Ins.OdCoCd and OmTc=Ins.OdTc and OmYy=Ins.OdYy and OmChr=Ins.OdChr and OmNo=Ins.OdNo  
   where OmFrRmDc= 'D' and FrPrtKey=SPrtKey and FrCoCd= SCoCd and FrRmCd=SRmCd   
   and FrLotNo=SLotNo and FrRmSz=SRmSz and FrRmStkRt=SRmStkRt and OmFrRmLoc=SLoc   
   and convert(int,convert(varchar(6),OmDt,112))= SYyMm and OdTc= 'JCN')   
  from FgRm   
  Join Fgd on FdPrtKey=FrPrtKey and FdCoCd= FrCoCd and FdTc= FrTc and FdYy= FrYy and FdChr= FrChr and FdNo= FrNo and FdSr=FrSr    
  Join Inserted On Inserted.OdPrtKey=FdPrtKey and Inserted.OdCoCd=FdCoCd and Inserted.OdBYy=FdBYy and Inserted.OdBChr=FdBChr and Inserted.OdBNo=FdBNo  
  Join OrdMst On OmPrtKey=Inserted.OdPrtKey and OmCoCd=Inserted.OdCoCd and OmTc=Inserted.OdTc and OmYy=Inserted.OdYy and OmChr=Inserted.OdChr and OmNo=Inserted.OdNo  
  where FrPrtKey=SPrtKey and FrCoCd= SCoCd and FrRmCd= SRmCd and FrLotNo= SLotNo   
  and FrRmSz= SRmSz and FrRmStkRt= SRmStkRt   
  and OmFrRmLoc= SLoc and convert(int,convert(varchar(6),OmDt,112))=SYyMm  
  and (OmFrRmDc='D'  or OmFrRmDc='C')  
  and OdTc= 'JCN'  
    
  if @mode in ('A')  
  Begin   
   update Bag set   
   BCls= 'Y',   
   BCnvJtYn ='Y',  
   BClsDt= (Select OmDt From OrdMst Where OmCoCd= OdCoCd And OmTc = OdTc   
     And OmYy = OdYy And OmChr= OdChr And OmNo= OdNo And OmPrtKey=OdPrtKey and OmIdNo=OdOmIdNo)   
   from inserted where BCoCd= OdCoCd and BYy= OdBYy and BChr= OdBChr and BNo= OdBNo and OdTc= 'JCN'    
  End  
  
  
 End  
 End  
  
 /*** Update DmYyMm from OmToDmLoc ***/  
 if @mode in ('C','D')  
 begin  
 /*** Jay 3.3(FG) ***/  
 /*** UPDATES FG In-Out ***/  
 /* **** Manali Trading Module **** */  
 /*If exists (Select 'x' from Deleted   
   where exists (Select 'x' from Param where PTyp= 'TC'   
      and (PValue = 'FG' Or (PValue='JT' and exists(Select 'x' from TcMst Where TmCoCd=OdCoCd and TmTc=OdTc and TmLocTypFr<>'' and TmLocTypTo<>'' )))   
      and PMCd= OdTc))*/  
 If @wTcTyp = 'FG' or @wTcTyp = 'JT'     
 Begin  
  /* Current Partition */  
  update DmYyMm set  
  DDmCrQty=DDmCrQty-(select isnull(sum(Del.OdOrdQty),0) from Deleted Del Join OrdMst  
   on OmPrtKey=Del.OdPrtKey and OmIdNo=Del.OdOmIdNo and OmCoCd= Del.OdCoCd and OmTc= Del.OdTc and OmYy= Del.OdYy and OmChr= Del.OdChr and OmNo= Del.OdNo  
   where OmToDmDc='C' and Del.OdPrtKey=DPrtKey and Del.OdCoCd= DCoCd and Del.OdDmCd=DDmCd   
   and Del.OdSfx= DOdSfx and Del.OdDmSz= DDmSz and Del.OdDmStkNo= DDmStkNo and Del.OdDmStkYy = DDmStkYy and Del.OdDmStkChr = DDmStkChr and OmToDmLoc=DLoc  
   and convert(int,convert(varchar(6),OmDt,112))=DYyMm),  
  DDmDrQty=DDmDrQty-(select isnull(sum(Del.OdOrdQty),0) from Deleted Del Join OrdMst   
   on OmPrtKey=Del.OdPrtKey and OmIdNo=Del.OdOmIdNo and OmCoCd= Del.OdCoCd and OmTc= Del.OdTc and OmYy= Del.OdYy and OmChr= Del.OdChr and OmNo= Del.OdNo  
   where OmToDmDc='D' and Del.OdPrtKey=DPrtKey and Del.OdCoCd= DCoCd and Del.OdDmCd=DDmCd   
   and Del.OdSfx= DOdSfx and Del.OdDmSz= DDmSz and Del.OdDmStkNo= DDmStkNo and Del.OdDmStkYy = DDmStkYy and Del.OdDmStkChr = DDmStkChr
   and OmToDmLoc=DLoc  
   and convert(int,convert(varchar(6),OmDt,112))=DYyMm)  
  from Deleted Join OrdMst on OmPrtKey=Deleted.OdPrtKey and OmIdNo=Deleted.OdOmIdNo and OmCoCd= Deleted.OdCoCd and OmTc= Deleted.OdTc  
   and OmYy= Deleted.OdYy and OmChr= Deleted.OdChr and OmNo= Deleted.OdNo  
   where Deleted.OdPrtKey=DPrtKey and Deleted.OdCoCd= DCoCd and Deleted.OdDmCd= DDmCd   
   and Deleted.OdSfx= DOdSfx and Deleted.OdDmSz= DDmSz and Deleted.OdDmStkNo = DDmStkNo and Deleted.OdDmStkYy = DDmStkYy and 
   Deleted.OdDmStkChr = OdDmStkChr and OmToDmLoc=DLoc  
   and convert(int,convert(varchar(6),OmDt,112))= DYyMm  
   and (OmToDmDc='D' or OmToDmDc='C')  
   and OmFrDmLoc <>'' and OmToDmLoc<>''  
 End  
  
 /*If exists (Select 'x' from Deleted   
   where exists (Select 'x' from Param where PTyp= 'TC' and PMCd= OdTc and PValue='JT' and PMCd='JCN'))*/  
  If @wTcTyp = 'JT'      
 Begin  
  update SYyMm set  
  SRmCrQty= SRmCrQty- (select isnull(sum(FrRmQty * (Case When FrFrRmLoc='FG' and FrFrRmDc= 'D' Then 1 When FrFrRmLoc='FG' and FrFrRmDc= 'C' Then -1 When FrToRmLoc='FG' and FrToRmDc= 'D' Then 1 When FrToRmLoc='FG' and FrToRmDc= 'C' Then -1 Else 0 End)), 0
)   
   from FgRm   
   Join Fgd on FdPrtKey=FrPrtKey and FdCoCd= FrCoCd and FdTc= FrTc and FdYy= FrYy and FdChr= FrChr   
   and FdNo= FrNo and FdSr=FrSr    
   Join Deleted Del On Del.OdPrtKey=FdPrtKey and Del.OdCoCd=FdCoCd and Del.OdBYy=FdBYy and Del.OdBChr=FdBChr and Del.OdBNo=FdBNo  
   Join OrdMst On OmPrtKey=Del.OdPrtKey and OmCoCd=Del.OdCoCd and OmTc=Del.OdTc and OmYy=Del.OdYy and OmChr=Del.OdChr and OmNo=Del.OdNo  
   where OmToRmDc= 'C' and FrPrtKey=SPrtKey and FrCoCd= SCoCd and FrRmCd=SRmCd   
   and FrLotNo=SLotNo and FrRmSz=SRmSz and FrRmStkRt=SRmStkRt and OmToRmLoc=SLoc   
   and convert(int,convert(varchar(6),OmDt,112))= SYyMm and OdTc= 'JCN') ,  
  SRmCrWt= SRmCrWt-(select isnull(sum(FrRmWt * (Case When FrFrRmLoc='FG' and FrFrRmDc= 'D' Then 1 When FrFrRmLoc='FG' and FrFrRmDc= 'C' Then -1 When FrToRmLoc='FG' and FrToRmDc= 'D' Then 1 When FrToRmLoc='FG' and FrToRmDc= 'C' Then -1 Else 0 End)), 0)   
   from FgRm  
   Join Fgd on FdPrtKey=FrPrtKey and FdCoCd= FrCoCd and FdTc= FrTc and FdYy= FrYy and FdChr= FrChr   
   and FdNo= FrNo and FdSr=FrSr    
   Join Deleted Del On Del.OdPrtKey=FdPrtKey and Del.OdCoCd=FdCoCd and Del.OdBYy=FdBYy and Del.OdBChr=FdBChr and Del.OdBNo=FdBNo  
   Join OrdMst On OmPrtKey=Del.OdPrtKey and OmCoCd=Del.OdCoCd and OmTc=Del.OdTc and OmYy=Del.OdYy and OmChr=Del.OdChr and OmNo=Del.OdNo  
   where OmToRmDc= 'C' and FrPrtKey=SPrtKey and FrCoCd= SCoCd and FrRmCd=SRmCd   
   and FrLotNo=SLotNo and FrRmSz=SRmSz and FrRmStkRt=SRmStkRt and OmToRmLoc=SLoc   
   and convert(int,convert(varchar(6),OmDt,112))= SYyMm and OdTc= 'JCN') ,  
  SRmDrQty= SRmDrQty-(select isnull(sum(FrRmQty * (Case When FrFrRmLoc='FG' and FrFrRmDc= 'D' Then 1 When FrFrRmLoc='FG' and FrFrRmDc= 'C' Then -1 When FrToRmLoc='FG' and FrToRmDc= 'D' Then 1 When FrToRmLoc='FG' and FrToRmDc= 'C' Then -1 Else 0 End)), 0) 
  
   from FgRm   
   Join Fgd on FdPrtKey=FrPrtKey and FdCoCd= FrCoCd and FdTc= FrTc and FdYy= FrYy and FdChr= FrChr   
   and FdNo= FrNo and FdSr=FrSr    
   Join Deleted Del On Del.OdPrtKey=FdPrtKey and Del.OdCoCd=FdCoCd and Del.OdBYy=FdBYy and Del.OdBChr=FdBChr and Del.OdBNo=FdBNo  
   Join OrdMst On OmPrtKey=Del.OdPrtKey and OmCoCd=Del.OdCoCd and OmTc=Del.OdTc and OmYy=Del.OdYy and OmChr=Del.OdChr and OmNo=Del.OdNo  
   where OmToRmDc= 'D' and FrPrtKey=SPrtKey and FrCoCd= SCoCd and FrRmCd=SRmCd   
   and FrLotNo=SLotNo and FrRmSz=SRmSz and FrRmStkRt=SRmStkRt and OmToRmLoc=SLoc   
   and convert(int,convert(varchar(6),OmDt,112))= SYyMm and OdTc= 'JCN') ,  
  SRmDrWt= SRmDrWt-(select isnull(sum(FrRmWt * (Case When FrFrRmLoc='FG' and FrFrRmDc= 'D' Then 1 When FrFrRmLoc='FG' and FrFrRmDc= 'C' Then -1 When FrToRmLoc='FG' and FrToRmDc= 'D' Then 1 When FrToRmLoc='FG' and FrToRmDc= 'C' Then -1 Else 0 End)), 0)   
   from FgRm   
   Join Fgd on FdPrtKey=FrPrtKey and FdCoCd= FrCoCd and FdTc= FrTc and FdYy= FrYy and FdChr= FrChr   
   and FdNo= FrNo and FdSr=FrSr    
   Join Deleted Del On Del.OdPrtKey=FdPrtKey and Del.OdCoCd=FdCoCd and Del.OdBYy=FdBYy and Del.OdBChr=FdBChr and Del.OdBNo=FdBNo  
   Join OrdMst On OmPrtKey=Del.OdPrtKey and OmCoCd=Del.OdCoCd and OmTc=Del.OdTc and OmYy=Del.OdYy and OmChr=Del.OdChr and OmNo=Del.OdNo  
   where OmToRmDc= 'D' and FrPrtKey=SPrtKey and FrCoCd= SCoCd and FrRmCd=SRmCd   
   and FrLotNo=SLotNo and FrRmSz=SRmSz and FrRmStkRt=SRmStkRt and OmToRmLoc=SLoc   
   and convert(int,convert(varchar(6),OmDt,112))= SYyMm and OdTc= 'JCN')   
  from FgRm   
  Join Fgd on FdPrtKey=FrPrtKey and FdCoCd= FrCoCd and FdTc= FrTc and FdYy= FrYy and FdChr= FrChr and FdNo= FrNo and FdSr=FrSr    
  Join Deleted On Deleted.OdPrtKey=FdPrtKey and Deleted.OdCoCd=FdCoCd and Deleted.OdBYy=FdBYy and Deleted.OdBChr=FdBChr and Deleted.OdBNo=FdBNo  
  Join OrdMst On OmPrtKey=Deleted.OdPrtKey and OmCoCd=Deleted.OdCoCd and OmTc=Deleted.OdTc and OmYy=Deleted.OdYy and OmChr=Deleted.OdChr and OmNo=Deleted.OdNo  
  where FrPrtKey=SPrtKey and FrCoCd= SCoCd and FrRmCd= SRmCd and FrLotNo= SLotNo   
  and FrRmSz= SRmSz and FrRmStkRt= SRmStkRt   
  and OmToRmLoc= SLoc and convert(int,convert(varchar(6),OmDt,112))=SYyMm  
  and (OmToRmDc='D'  or OmToRmDc='C')  
  and OdTc= 'JCN'  
 End  
  
 /*If exists (Select 'x' from Deleted   
   where exists (Select 'x' from Param where PTyp= 'TC' and PMCd= OdTc and PValue='JT')  
   and Exists (Select 'x' from TcMst Where TmCoCd=OdCoCd and TmTc=OdTC and TmBaseTCs<>''))*/  
 If @wTcTyp = 'JT'     
 Begin   
  Update OrdDsg   
  Set OrdDsg.OdExpQty= OrdDsg.OdExpQty- IsNull((Select sum(del.OdOrdQty) from deleted del where   
   del.OdPrtKey=OrdDsg.OdPrtKey and  (del.OdTc='JIN' Or del.OdCoCd= OrdDsg.OdCoCd) and del.OdBaseIdKey= OrdDsg.OdIdNo), 0)  
   from deleted where Deleted.OdPRtKey=@wCurrPrtn and OrdDsg.OdPrtKey=Deleted.OdPrtKey and  
    OrdDsg.OdCoCd= deleted.OdCoCd and OrdDsg.OdIdNo=deleted.OdBaseIdkey  
   and Exists (Select 'x' from TcMst Where TmCoCd=deleted.OdCoCd and TmTc=deleted.OdTC and TmBaseTCs<>'')  
  
  Update OrdDsg   
  Set OdCls = (case when OrdDsg.OdOrdQty= OrdDsg.OdExpQty and Round(OrdDsg.OdOrdQty, 1)<> 0 then 'Y' else 'N' end)   
   from deleted del where OrdDsg.OdPRtKEy=Del.OdPRtKey and OrdDsg.OdCoCd= Del.OdCoCd and OrdDsg.OdIdNo= Del.OdBaseIdKey  
   and Del.OdPrtKey=@wCurrPrtn   
   and Exists (Select 'x' from TcMst Where TmCoCd=del.OdCoCd and TmTc=del.OdTC and TmBaseTCs<>'')  
 End   
  
 End   
  
 if @mode in ('A','C')  
 begin  
 /*** Jay 3.3(FG) ***/  
 /*** UPDATES FG In-Out ***/  
 /* **** Manali Trading Module **** */  
 /*If exists (Select 'x' from Inserted   
   where exists (Select 'x' from Param where PTyp= 'TC'   
      and (PValue = 'FG' Or (PValue='JT' and exists(Select 'x' from TcMst Where TmCoCd=OdCoCd and TmTc=OdTc and TmLocTypFr<>'' and TmLocTypTo<>'' )))   
      and PMCd= OdTc))*/  
 If @wTcTyp = 'FG' or @wTcTyp = 'JT'    
 Begin  
  /* Current Partition */  
  insert into DmYyMm(DPrtKey,DCoCd,DDmCd,DOdSfx,DDmSz,DDmStkYy,DDmStkChr,DDmStkNo,DLoc,DYyMm,DDmDrQty,DDmCrQty)  
   select distinct OdPrtKey, OdCoCd, OdDmCd, OdSfx, OdDmSz,OdDmStkYy, OdDmStkChr, OdDmStkNo,OmToDmLoc,convert(int,convert(varchar(6),OmDt,112)),0,0  
   from inserted join OrdMst on OmPrtKey=OdPrtKey and OmIdNo=OdOmIdNo and OmCoCd=OdCoCd and OmTc=OdTc and OmYy=OdYy and OmChr=OdChr and OmNo=OdNo  
   where not exists(select 'x' from  DmYyMm where DPrtKey=OdPrtKey and DCoCd= OdCoCd and DDmCd=OdDmCd   
       and OdSfx= DOdSfx and OdDmSz= DDmSz and OdDmStkYy = DDmStkYy and OdDmStkChr = DDmStkChr and OdDmStkNo = DDmStkNo and DLoc=OmToDmLoc and  
       DYyMm= convert(int,convert(varchar(6),OmDt,112)))  
    and OmFrDmLoc <>'' and OmToDmLoc<>''  
  
    
  /* Current Partition */  
  update DmYyMm set  
  DDmCrQty=DDmCrQty+(select isnull(sum(Ins.OdOrdQty),0) from Inserted Ins Join OrdMst   
   on OmPrtKey=Ins.OdPrtKey and OmIdNo=Ins.OdOmIdNo and OmCoCd= Ins.OdCoCd and OmTc= Ins.OdTc and OmYy= Ins.OdYy and OmChr= Ins.OdChr and OmNo= Ins.OdNo  
   where OmToDmDc='C' and Ins.OdPrtKey=DPrtKey and Ins.OdCoCd= DCoCd and Ins.OdDmCd=DDmCd   
   and Ins.OdSfx= DOdSfx and Ins.OdDmSz= DDmSz and Ins.OdDmStkNo= DDmStkNo and Ins.OdDmStkYy = DDmStkYy and Ins.OdDmStkChr = DDmStkChr and OmToDmLoc=DLoc  
   and convert(int,convert(varchar(6),OmDt,112))=DYyMm),  
  DDmDrQty=DDmDrQty+(select isnull(sum(Ins.OdOrdQty),0) from Inserted Ins Join OrdMst   
   on OmPrtKey=Ins.OdPrtKey and OmIdNo=Ins.OdOmIdNo and OmCoCd= Ins.OdCoCd and OmTc= Ins.OdTc and OmYy= Ins.OdYy and OmChr= Ins.OdChr and OmNo= Ins.OdNo  
   where OmToDmDc='D' and Ins.OdPrtKey=DPrtKey and Ins.OdCoCd= DCoCd and Ins.OdDmCd=DDmCd   
   and Ins.OdSfx= DOdSfx and Ins.OdDmSz= DDmSz and Ins.OdDmStkNo= DDmStkNo and Ins.OdDmStkChr = DDmStkChr and Ins.OdDmStkChr = DDmStkChr and OmToDmLoc=DLoc  
   and convert(int,convert(varchar(6),OmDt,112))=DYyMm)  
  from Inserted Join OrdMst on OmPrtKey=Inserted.OdPrtKey and OmIdNo=Inserted.OdOmIdNo and OmCoCd= Inserted.OdCoCd and OmTc= Inserted.OdTc  
   and OmYy= Inserted.OdYy and OmChr= Inserted.OdChr and OmNo= Inserted.OdNo  
   where Inserted.OdPrtKey=DPrtKey and Inserted.OdCoCd= DCoCd and Inserted.OdDmCd=DDmCd   
   and Inserted.OdSfx= DOdSfx and Inserted.OdDmSz= DDmSz and Inserted.OdDmStkNo= DDmStkNo and Inserted.OdDmStkChr = DDmStkChr and Inserted.OdDmStkYy = DDmStkYy 
   and OmToDmLoc= DLoc  
   and convert(int,convert(varchar(6),OmDt,112))=DYyMm  
   and (OmToDmDc='D' or OmToDmDc='C')  
   and OmFrDmLoc <>'' and OmToDmLoc<>''  
 End  
  
 /*If exists (Select 'x' from Inserted   
   where exists (Select 'x' from Param where PTyp= 'TC' and PMCd= OdTc and PValue='JT' and PMCd='JCN'))*/  
 If @wTcTyp = 'JT'     
 Begin   
  insert into SYyMm(SPrtKey,SCoCd,SRmCd,SLotNo,SRmSz,SRmStkRt,SLoc,SYyMm,SRmDrqty,SRmDrWt,SRmCrQty,SRmCrWt)   
   select distinct FrPrtKey,FrCoCd,FrRmCd,FrLotNo,FrRmSz,FrRmStkRt,OmToRmLoc,convert(int,convert(varchar(6),OmDt,112)),0,0,0,0  
   From FgRm  
   Join Fgd on FdPrtKey=FrPrtKey and FdCoCd= FrCoCd and FdTc= FrTc and FdYy= FrYy and FdChr= FrChr   
   and FdNo= FrNo and FdSr=FrSr    
   Join Inserted On OdPrtKey=FdPrtKey and OdCoCd=FdCoCd and OdBYy=FdBYy and OdBChr=FdBChr and OdBNo=FdBNo  
   Join OrdMst On OmPrtKey=OdPrtKey and OmCoCd=OdCoCd and OmTc=OdTc and OmYy=OdYy and OmChr=OdChr and OmNo=OdNo  
   where not exists(select 'x' from  SYyMm where SPrtKey=FrPrtKey and SCoCd= FrCoCd and SRmCd=FrRmCd and SLotNo=FrLotNo   
   and SRmSZ=FrRmSz and SRmStkRt=FrRmStkRt and SLoc=OmToRmLoc and  
   SYyMm= convert(int,convert(varchar(6),OmDt,112))) and OdTc= 'JCN'   
  update SYyMm set  
  SRmCrQty= SRmCrQty+(select isnull(sum(FrRmQty * (Case When FrFrRmLoc='FG' and FrFrRmDc= 'D' Then 1 When FrFrRmLoc='FG' and FrFrRmDc= 'C' Then -1 When FrToRmLoc='FG' and FrToRmDc= 'D' Then 1 When FrToRmLoc='FG' and FrToRmDc= 'C' Then -1 Else 0 End)), 0) 
  
   from FgRm  
   Join Fgd on FdPrtKey=FrPrtKey and FdCoCd= FrCoCd and FdTc= FrTc and FdYy= FrYy and FdChr= FrChr   
   and FdNo= FrNo and FdSr=FrSr    
   Join Inserted Ins On Ins.OdPrtKey=FdPrtKey and Ins.OdCoCd=FdCoCd and Ins.OdBYy=FdBYy and Ins.OdBChr=FdBChr and Ins.OdBNo=FdBNo  
   Join OrdMst On OmPrtKey=Ins.OdPrtKey and OmCoCd=Ins.OdCoCd and OmTc=Ins.OdTc and OmYy=Ins.OdYy and OmChr=Ins.OdChr and OmNo=Ins.OdNo  
   where OmToRmDc= 'C' and FrPrtKey=SPrtKey and FrCoCd= SCoCd and FrRmCd=SRmCd   
   and FrLotNo=SLotNo and FrRmSz=SRmSz and FrRmStkRt=SRmStkRt and OmToRmLoc=SLoc   
   and convert(int,convert(varchar(6),OmDt,112))= SYyMm and OdTc= 'JCN') ,  
  SRmCrWt=SRmCrWt+(select isnull(sum(FrRmWt * (Case When FrFrRmLoc='FG' and FrFrRmDc= 'D' Then 1 When FrFrRmLoc='FG' and FrFrRmDc= 'C' Then -1 When FrToRmLoc='FG' and FrToRmDc= 'D' Then 1 When FrToRmLoc='FG' and FrToRmDc= 'C' Then -1 Else 0 End)), 0)   
   from FgRm   
   Join Fgd on FdPrtKey=FrPrtKey and FdCoCd= FrCoCd and FdTc= FrTc and FdYy= FrYy and FdChr= FrChr   
   and FdNo= FrNo and FdSr=FrSr    
   Join Inserted Ins On Ins.OdPrtKey=FdPrtKey and Ins.OdCoCd=FdCoCd and Ins.OdBYy=FdBYy and Ins.OdBChr=FdBChr and Ins.OdBNo=FdBNo  
   Join OrdMst On OmPrtKey=Ins.OdPrtKey and OmCoCd=Ins.OdCoCd and OmTc=Ins.OdTc and OmYy=Ins.OdYy and OmChr=Ins.OdChr and OmNo=Ins.OdNo  
   where OmToRmDc= 'C' and FrPrtKey=SPrtKey and FrCoCd= SCoCd and FrRmCd=SRmCd   
   and FrLotNo=SLotNo and FrRmSz=SRmSz and FrRmStkRt=SRmStkRt and OmToRmLoc=SLoc   
   and convert(int,convert(varchar(6),OmDt,112))= SYyMm and OdTc= 'JCN') ,  
  SRmDrQty=SRmDrQty+(select isnull(sum(FrRmQty * (Case When FrFrRmLoc='FG' and FrFrRmDc= 'D' Then 1 When FrFrRmLoc='FG' and FrFrRmDc= 'C' Then -1 When FrToRmLoc='FG' and FrToRmDc= 'D' Then 1 When FrToRmLoc='FG' and FrToRmDc= 'C' Then -1 Else 0 End)), 0)  
 
   from FgRm   
   Join Fgd on FdPrtKey=FrPrtKey and FdCoCd= FrCoCd and FdTc= FrTc and FdYy= FrYy and FdChr= FrChr   
   and FdNo= FrNo and FdSr=FrSr    
   Join Inserted Ins On Ins.OdPrtKey=FdPrtKey and Ins.OdCoCd=FdCoCd and Ins.OdBYy=FdBYy and Ins.OdBChr=FdBChr and Ins.OdBNo=FdBNo  
   Join OrdMst On OmPrtKey=Ins.OdPrtKey and OmCoCd=Ins.OdCoCd and OmTc=Ins.OdTc and OmYy=Ins.OdYy and OmChr=Ins.OdChr and OmNo=Ins.OdNo  
   where OmToRmDc= 'D' and FrPrtKey=SPrtKey and FrCoCd= SCoCd and FrRmCd=SRmCd   
   and FrLotNo=SLotNo and FrRmSz=SRmSz and FrRmStkRt=SRmStkRt and OmToRmLoc=SLoc   
   and convert(int,convert(varchar(6),OmDt,112))= SYyMm and OdTc= 'JCN') ,  
  SRmDrWt=SRmDrWt+(select isnull(sum(FrRmWt * (Case When FrFrRmLoc='FG' and FrFrRmDc= 'D' Then 1 When FrFrRmLoc='FG' and FrFrRmDc= 'C' Then -1 When FrToRmLoc='FG' and FrToRmDc= 'D' Then 1 When FrToRmLoc='FG' and FrToRmDc= 'C' Then -1 Else 0 End)), 0)   
   from FgRm   
   Join Fgd on FdPrtKey=FrPrtKey and FdCoCd= FrCoCd and FdTc= FrTc and FdYy= FrYy and FdChr= FrChr   
   and FdNo= FrNo and FdSr=FrSr    
   Join Inserted Ins On Ins.OdPrtKey=FdPrtKey and Ins.OdCoCd=FdCoCd and Ins.OdBYy=FdBYy and Ins.OdBChr=FdBChr and Ins.OdBNo=FdBNo  
   Join OrdMst On OmPrtKey=Ins.OdPrtKey and OmCoCd=Ins.OdCoCd and OmTc=Ins.OdTc and OmYy=Ins.OdYy and OmChr=Ins.OdChr and OmNo=Ins.OdNo  
   where OmToRmDc= 'D' and FrPrtKey=SPrtKey and FrCoCd= SCoCd and FrRmCd=SRmCd   
   and FrLotNo=SLotNo and FrRmSz=SRmSz and FrRmStkRt=SRmStkRt and OmToRmLoc=SLoc   
   and convert(int,convert(varchar(6),OmDt,112))= SYyMm and OdTc= 'JCN')   
  from FgRm   
  Join Fgd on FdPrtKey=FrPrtKey and FdCoCd= FrCoCd and FdTc= FrTc and FdYy= FrYy and FdChr= FrChr and FdNo= FrNo and FdSr=FrSr    
  Join Inserted On Inserted.OdPrtKey=FdPrtKey and Inserted.OdCoCd=FdCoCd and Inserted.OdBYy=FdBYy and Inserted.OdBChr=FdBChr and Inserted.OdBNo=FdBNo  
  Join OrdMst On OmPrtKey=Inserted.OdPrtKey and OmCoCd=Inserted.OdCoCd and OmTc=Inserted.OdTc and OmYy=Inserted.OdYy and OmChr=Inserted.OdChr and OmNo=Inserted.OdNo  
  where FrPrtKey=SPrtKey and FrCoCd= SCoCd and FrRmCd= SRmCd and FrLotNo= SLotNo   
  and FrRmSz= SRmSz and FrRmStkRt= SRmStkRt   
  and OmToRmLoc= SLoc and convert(int,convert(varchar(6),OmDt,112))=SYyMm  
  and (OmToRmDc='D'  or OmToRmDc='C')  
  and OdTc= 'JCN'  
 End  
  
 /*If exists (Select 'x' from Inserted   
   where exists (Select 'x' from Param where PTyp= 'TC' and PMCd= OdTc and PValue='JT')  
   and Exists (Select 'x' from TcMst Where TmCoCd=OdCoCd and TmTc=OdTC and TmBaseTCs<>''))*/  
 If @wTcTyp = 'JT'   
 Begin   
/*  
  select @row= count('x') from Inserted   
   Join OrdDsg On OrdDsg.OdPrtKey=Inserted.OdPrtKey and OrdDsg.OdCoCd= Inserted.OdCoCd   
         and OrdDsg.OdIdNo=Inserted.OdBaseIdkey  
   where Inserted.OdPRtKey=@wCurrPrtn   
   and OrdDsg.OdOrdQty < OrdDsg.OdExpQty + Inserted.OdOrdQty  
  if @row> 0     
  begin select @errmsg= 'Order Qty Should Be >= Export Qty 1'    
   goto ErrHndlr    
  end    
*/  
  Update OrdDsg   
  Set OrdDsg.OdExpQty= OrdDsg.OdExpQty + IsNull((Select sum(Ins.OdOrdQty) from Inserted Ins where   
   Ins.OdPrtKey=OrdDsg.OdPrtKey and  (ins.OdTc='JIN' Or Ins.OdCoCd= OrdDsg.OdCoCd)  and Ins.OdBaseIdKey= OrdDsg.OdIdNo), 0)  
   from Inserted where Inserted.OdPRtKey=@wCurrPrtn and OrdDsg.OdPrtKey=Inserted.OdPrtKey and  
    OrdDsg.OdCoCd= Inserted.OdCoCd and OrdDsg.OdIdNo=Inserted.OdBaseIdkey  
    and Exists (Select 'x' from TcMst Where TmCoCd=Inserted.OdCoCd and TmTc=Inserted.OdTC and TmBaseTCs<>'')  
  
  Update OrdDsg   
  Set OdCls = (case when OrdDsg.OdOrdQty= OrdDsg.OdExpQty and Round(OrdDsg.OdOrdQty, 1)<> 0 then 'Y' else 'N' end)   
   from Inserted Ins where OrdDsg.OdPRtKEy=Ins.OdPRtKey and (ins.OdTc='JIN' Or  OrdDsg.OdCoCd= Ins.OdCoCd)   
   and OrdDsg.OdIdNo= Ins.OdBaseIdKey and Ins.OdPrtKey=@wCurrPrtn   
   and Exists (Select 'x' from TcMst Where TmCoCd=Ins.OdCoCd and TmTc=Ins.OdTC and TmBaseTCs<>'')  
 End  
  
 End    
  
  
/*if @mode in ('C','A') and exists (Select 'x' from Inserted   
   Join Param On PTyp= 'TC' and PValue='JT'   
   and exists(Select 'x' from TcMst Where TmCoCd=OdCoCd and TmTc=OdTc and TmbaseTcs<>'')  
   and PMCd= OdTc)*/  
if @mode in ('C','A') and @wTcTyp = 'JT'   
Begin   
 Select @Row=Count('x') from Inserted Ins   
   /*Join OrdDsg B On B.OdPrtKey=Ins.OdPrtKey and B.OdCoCd= Ins.OdCoCd and B.OdIdNo=Ins.OdBaseIdkey */  
   where Ins.OdPRtKey=@wCurrPrtn   
    AND exists(Select 'x' from OrdDsg B Where  
      B.OdPrtKey=Ins.OdPrtKey and (Ins.OdTc='JIN' Or  B.OdCoCd= Ins.OdCoCd) and B.OdIdNo=Ins.OdBaseIdkey and B.OdOrdQty < B.OdExpQty)  
    and exists(Select 'x' from TcMst Where TmCoCd=OdCoCd and TmTc=OdTc and TmbaseTcs<>'')  
  If @Row> 0  
  begin   
   select @errmsg= 'Order Qty Should Be >= Matched Qty for Base Voucher'    
   goto ErrHndlr    
  end  
end    
  
/*if @mode in ('C','A') and exists (Select 'x' from Inserted Ins  
   Join Param On PTyp= 'TC' and (PValue = 'FG' Or (PValue='JT' and exists(Select 'x' from TcMst Where TmCoCd=OdCoCd and TmTc=OdTc and TmLocTypFr<>'' and TmLocTypTo<>'' )))   
   and PMCd= OdTc)*/  
If @mode in ('C','A') and  (@wTcTyp = 'FG' or @wTcTyp = 'JT')  
Begin   
 Select @Row=Count('x') From Inserted Ins   
   Join OrdMst on OmPrtKey=Ins.OdPrtKey and OmIdNo=Ins.OdOmIdNo and OmCoCd= Ins.OdCoCd and OmTc= Ins.OdTc  
       and OmYy= Ins.OdYy and OmChr= Ins.OdChr and OmNo= Ins.OdNo  
   Join Param On PTyp='TC' and PMCd=OdTc  
   Join Head On HCoCd=OdCoCd And HNegStkYn= 'N'  
   Join Loc On LocCoCd= Ins.OdCoCd and LocTyp =(Case PValue When 'FG' Then 'F' When 'JT' Then 'S' End) and   
      LocCd in (case when OmFrDmDC = 'C' then OmFrDmLoc when OmFrDmDC = 'D' then OmToDmLoc else '' end )  
      /*LocCd in (OmFrDmLoc,OmToDmLoc)*/  
  Where   
    not Exists (select 'x' from Param Join UsrAccess on UaUsrCd = PMCd where PTyp= 'USR' and           
     PMCd= Ins.ModUsr and UaMnuCd = Ins.OdTc and UaSeoYn= 'Y' and PValue= OmSeoPwd)           
    and not Exists (Select 'x' from Param where PTyp= 'SEOPWD' and PMCD= 'SEOPWD' and OmSeoPwd= PValue1)               
    AND (Select CAST(SUM(DDmDrQty)-SUM(DDmCrQty) As Decimal(10,3)) from DmYyMm  
    where Ins.OdPrtKEy=DPrtKey and Ins.OdCoCd= DCoCd and Ins.OdDmCd= DDmCd and Ins.OdSfx= DOdSfx and Ins.OdDmSz= DDmSz and   
     (case when OmFrDmDC = 'C' then OmFrDmLoc when OmFrDmDC = 'D' then OmToDmLoc else '' end )= DLoc   
    /*Ins.OdDmId= DDmId and DLoc=LocCd*/  
      )< 0  
    and OmFrDmLoc <>'' and OmToDmLoc<>''   
  
  If @Row> 0  
  begin  
   Select @ErrMsg= 'Only Authorized User Can Enter Negative Stock For This Desing Code, Suffix, ID'  
     GoTo ErrHndlr  
     end  
end  
  
  
  
/*if @mode in ('D') and exists (Select 'x' from Deleted   
     Join Param On PTyp= 'TC' and (PValue = 'FG' Or (PValue='JT' and exists(Select 'x' from TcMst Where TmCoCd=OdCoCd and TmTc=OdTc and TmLocTypFr<>'' and TmLocTypTo<>'' )))   
     and PMCd= OdTc)*/  
If @mode in ('D') and  (@wTcTyp = 'FG' or @wTcTyp = 'JT')  
Begin   
 Select @Row=Count('x') From Deleted Del   
   Join OrdMst on OmPrtKey=Del.OdPrtKey and OmIdNo=Del.OdOmIdNo and OmCoCd= Del.OdCoCd and OmTc= Del.OdTc  
       and OmYy= Del.OdYy and OmChr= Del.OdChr and OmNo= Del.OdNo  
   Join Param On PTyp='TC' and PMCd=OdTc  
   Join Head On HCoCd=OdCoCd And HNegStkYn= 'N'  
   Join Loc On LocCoCd= Del.OdCoCd and LocTyp =(Case PValue When 'FG' Then 'F' When 'JT' Then 'S' End)  and   
      LocCd in (case when OmFrDmDC = 'D' then OmFrDmLoc when OmFrDmDC = 'C' then OmToDmLoc else '' end )  
      /*LocCd in (OmFrDmLoc,OmToDmLoc)*/  
  Where   
    not Exists (select 'x' from Param Join UsrAccess on UaUsrCd = PMCd where PTyp= 'USR' and           
     PMCd= Del.ModUsr and UaMnuCd = Del.OdTc and UaSeoYn= 'Y' and PValue= OmSeoPwd)           
    and not Exists (Select 'x' from Param where PTyp= 'SEOPWD' and PMCD= 'SEOPWD' and OmSeoPwd= PValue1)               
    AND (Select CAST(SUM(DDmDrQty)-SUM(DDmCrQty) As Decimal(10,3)) from DmYyMm  
    where Del.OdPrtKEy=DPrtKey and Del.OdCoCd= DCoCd and Del.OdDmCd= DDmCd and Del.OdSfx= DOdSfx and Del.OdDmSz= DDmSz and   
     (case when OmFrDmDC = 'D' then OmFrDmLoc when OmFrDmDC = 'C' then OmToDmLoc else '' end )= DLoc  
     /*Del.OdDmId= DDmId and DLoc=LocCd*/  
      )< 0  
    and OmFrDmLoc <>'' and OmToDmLoc<>''  
  If @Row> 0  
  begin  
   Select @ErrMsg= 'Only Authorized User Can Enter Negative Stock For This Desing Code, Suffix, ID'  
     GoTo ErrHndlr  
     end  
end  
/*** Jay 3.3(FG) ***/  
return    
ErrHndlr:    
 begin   select @errmsg='|'+@errmsg    
  raiserror(@errmsg,16,1)   rollback transaction     
 end    
end   
GOGO

ALTER TRIGGER [dbo].[Head_INUPDL] ON [dbo].[Head]
    FOR INSERT,UPDATE,DELETE NOT FOR REPLICATION
AS
BEGIN
	/* *** Trigger Version 3.11.0 *** */
	DECLARE @ROW INT
	DECLARE @mode char(1)
	DECLARE @errmsg varchar(255)
	DECLARE @wSelfCd VARCHAR(8)
	DECLARE @wSelfCoCd varchar(3)
	DECLARE @ReplAgt BIT
	
	set @wSelfCoCd= 'ZZZ'
	set @wSelfCd = 'ZSELF'	

	IF @@ROWCOUNT=0
		RETURN
	If Exists(Select 'x' From TrgDisa Where TdTblNm = 'Head')
		RETURN

	select @mode=case (select count('x') from inserted) when 0 then 'D'
			else
			case (select count('x') from deleted) when 0 then 'A' else 'C' end
			end 

	/* **** Checking Whether a Normal Transaction or Replicated One **** */
	Set @ReplAgt=Convert(BIT,SessionProperty('Replication_Agent'))
	
	If @ReplAgt = 1         Return

	If (@ReplAgt<>1)
	Begin
		/******Cannot Insert/Update/Delete  if Crm database /Emr database with TrdCocd not in sdCoCd ****/
		/***if SysId=Emr Record not found then normal behaviour i.e can insert/update/delete****/
		if @mode in ('A','C')
		Begin 
		select @ROW =Count('x')  From inserted where  
		Exists ( select 'x' From SysDb where (SysId='EMR' and ','+sdCocd+',' Not  like '%,'+ HCoCd +',%') 
			or (SysId='CRM' and SysDb='WEB'))
		if @Row > 0		
			Begin 
				select @errmsg= 'Cannot Insert /Update ' 
				goto ErrHndlr 
			end 
		End 
		if @mode in ('D')
		Begin 
		select @ROW =Count('x')  From Deleted where  
		Exists ( select 'x' From SysDb where (SysId='EMR' and ','+sdCocd+',' Not  like '%,'+ HCoCd +',%')
			or (SysId='CRM' and SysDb='WEB'))
		if @Row > 0		
			Begin 
				select @errmsg= 'Cannot Delete ' 
				goto ErrHndlr 
			end 
		End 
	End	

	If @mode in ('A' , 'C')
	Begin
	/* Zubin 211 (hFluteBagYN= ins.hFluteBagYN and HDTPrdPtr= ins.HDTPrdPtr added) */ 
		/*** introduced hLooseInv (Jen 2.12 Patch 3)***/
		/*** introduced HConsolidatedInv, HAutomateReturns (Jen 2.13)***/
		/****** Sachin 2.13.0 - introduced HQWModYN And HModAvlYN ******/ 
		/****** Sachin 2.14.0 - introduced HImpComLoc ******/ 
		/* Zubin 213 (HMultiCurYn added) */ 
		/* *** JAY 2.13(CT) *** (hCompBagsYn ADDED)*/
		/* *** Jay 2.14(LclInv) *** [HLclInvYn]*/
		/* *** Jay 3.01 *** [FgTag Module] */
		/*** (Jen 3.1.0_1 added HInInvGrsWtEqActual) ***/
		/*** Jay 3.1.0_1 [HValidDsgSetCd] ***/ 
		/****** Jen (3.1.2) added hBestUse ******/
		/****** Manali 3.03 added HPdModYN */
		/****** Sachin 3.3.0 (30-06-08) - [hChkVldRmLnYN added for 28. Valid RmSz Master] */
		/*** Jay 3.3(FG) [hFgIOYn] ***/
		/***** Manali 3.4.1 - 22/10/08 - Design Catalogue Module [added HDsgCatYN] *****/
		/* **** Manali 3.5.0 -  hCustomizeRmCtgExp, HMaxDsgViewRec added */
		/* **** Manali 3.6.0 - HOrdExcelRepYn added */
		/* **** Manali 3.9.1 - RfId Module ****/
		/* **** Manali 3.10.0 - Multi Price Quotation Module */
        /* 4.1.4 */
		Update Head 
		Set hCmFixPrc= ins.hCmFixPrc, hCmGldAs= ins.hCmGldAs, hCmLabAs=ins.hCmLabAs, 
			hCmIWtEqOrd = ins.hCmIWtEqOrd, hCmIWtFrOrd= ins.hCmIWtFrOrd, 
			hCmIGldRtEqOrd= ins.hCmIGldRtEqOrd, hCmDiaRtFA= ins.hCmDiaRtFA, 
			hCmCsRtFA= ins.hCmCsRtFA, HPtrYn= ins.HPtrYn, 
			hPPCYN= ins.hPPCYN, HPDCSfx= ins.HPDCSfx,
			HRndOrdSalPrc= ins.HRndOrdSalPrc, HShowPrdRmDets= ins.HShowPrdRmDets, 
			HCtbMod= ins.HCtbMod, hFluteBagYN= ins.hFluteBagYN, 
			HDTPrdPtr= ins.HDTPrdPtr, HSlvModYN = ins.HSlvModYN, HPdModYN = ins.HPdModYN,
			hLooseInv= ins.hLooseInv, hConsolidatedInv= ins.hConsolidatedInv, 
			HAutomateReturns= ins.HAutomateReturns, HQWModYN = Ins.HQWModYN, 
			HModAvlYN = Ins.HModAvlYN, HMultiCurYn= Ins.HMultiCurYn, hCompBagsYn= Ins.hCompBagsYn,
			HImpComLoc = Ins.HImpComLoc, HInRepMetLsRmCd= Ins.HInRepMetLsRmCd, HLclInvYn= Ins.HLclInvYn,
			HFgTagMod = Ins.HFgTagMod, HCustomiseOrd= Ins.HCustomiseOrd, HAutoCustmMtch= Ins.HAutoCustmMtch, 
			HInInvGrsWtEqActual= Ins.HInInvGrsWtEqActual, HValidDsgSetCd = Ins.HValidDsgSetCd, 
			hBestUse= Ins.hBestUse, hChkVldRmLnYN= Ins.hChkVldRmLnYN , HExcelImpInOrdYN=Ins.HExcelImpInOrdYN,
			HTravelExlYN=Ins.HTravelExlYN,hFgIOYn= Ins.hFgIOYn, HDsgCatYN= Ins.HDsgCatYN, 
			hCustomizeRmCtgExp=Ins.hCustomizeRmCtgExp, HMaxDsgViewRec=ins.HMaxDsgViewRec, HOrdExcelRepYn=Ins.HOrdExcelRepYn,
			HRfIdMod=ins.HRfIdMod, HMultiPrcQtMod=ins.HMultiPrcQtMod, hRtfrRtChrtYN=ins.hRtfrRtChrtYN,
			hTolQtyPctDC=ins.hTolQtyPctDC, hTolWtPctDC=ins.hTolWtPctDC  
		From Inserted ins where ins.HCoCd= @wSelfCoCd and Head.HCd= ins.HCd 
			And Head.HCoCd<> ins.HCoCd 
			And (IsNull((Select ''','+ SdCoCd + ',''' From SysDb Where SysId='EMR'),',' + Head.HCoCd + ',')) Like
					'%,' + Head.HCoCd + ',%'  
	/* Zubin 211 (hFluteBagYN= ins.hFluteBagYN added) */ 

/*      urmi 2.04-3 maintain HPtrYn same for other cocds as that of 'ZZZ'*/
/**  	Update the HPtrYn for other  Compny Codes when changed for 'ZZZ'    **/
/*		select @row=count('x') from inserted ins where ins.HCoCd= @wSelfCoCd and
			Exists(Select 'x' from Deleted del where 
				ins.HCoCd= del.HCoCd and ins.HCd= del.HCd and ins.hPtrYN<> del.hPtrYN)
		if @row> 0 
		begin
			Update Head set HPtrYn = (select H1.hPtrYn from inserted H1 where H1.HCoCd= @wSelfCoCd)
			where HCoCd<> @wSelfCoCd			
		end
*/
	End

	if @mode in ('A','C')
	Begin
		/****** Sachin 3.3.0 - Cannot Update the Check Integrity Field ******/
		Select @row=count('x') from inserted ins Where 
			Not Exists(Select 'x' From Deleted del Where del.HChkIntgty=ins.HChkIntgty)
		if @row>0 
		begin
			select @errmsg= ' Cannot Update Check Integrity Field...'
			goto ErrHndlr
		end

		/****** The HCd field in Head will always have the value 'ZSELF' ******/
		select @row=count('x') from inserted where HCd<> 'ZSELF' 
		if @row>0 
		begin
			select @errmsg= ' In Head Master Head Code Should Always Be ''ZSELF'' '
			goto ErrHndlr
		end

		/****** HCoCd = Not Blank ******/ 
		select @row=count('x') from inserted where HCoCd= ''
		if @row>0 
		begin
			select @errmsg= 'Company Code Cannot Be Blank'
			goto ErrHndlr
		end

		/****** HName = Not Blank ******/		
		select @row=count('x') from inserted where HName= ''
		if @row>0 
		begin
			select @errmsg= 'Company Name Cannot Be Blank'
			goto ErrHndlr
		end

		/*
		******  HBnkInstr  =  Not Blank ******
		select @row=count('x') from inserted where HBnkInstr= '' 
		if @row>0 
		Begin
			select @errmsg= 'Bank Instructions Cannot Be Blank'
			goto ErrHndlr
		end
		*/

		/****** HTotAs = PMCd('INVTOTAS') ******/	
		select @row= count('x') from inserted where 
			not exists (select 'x' from Param where 
				    PTyp='INVTOTAS' and PMCd= HTotAs)
		if @row>0 
		begin
			select @errmsg=' ''Invoice Total As'' Option Not Defined'
			goto ErrHndlr
		end

		/****** HCmGldAs = PMCd('GLDAS') ******/	
		/*** (Jen 2.14) ***/ 
		select @row= count('x') from inserted where hCmGldAs<> '' and 
			not exists (select 'x' from Param where PTyp='GLDAS' and PMCd= hCmGldAs)

		/* select @row= count('x') from inserted where hCmGldAs<> '' and 
			not exists (select 'x' from Param where PTyp='GLDAS' and PMCd= hCmGldAs 
			and not (',' + PDesc225 + ',' like '%,G,%' Or ',' + PDesc225 + ',' like '%,P,%' Or ',' + PDesc225 + ',' like '%,S,%') 
		) */
		/*** (Jen 2.14) ***/ 

		/*** (Bef 2.14) 
		select @row= count('x') from inserted where 
				not exists (select 'x' from Param where PTyp='GLDAS' and PMCd= hCmGldAs)
		(Bef 2.14) ***/
		if @row> 0 
		begin
			select @errmsg= 'Gold As Value Not Defined'
			goto ErrHndlr
		end



		/* ****** Manali 3.9.1 - Rate Round Off Option For Order ****** */
		select @row= count('x') from inserted Ins where HOrdRtRndOffRmCtg <> '' 
			and	not exists (select 'x' from Param where PTyp='CPYRT' and PMCd= HOrdRtRndOffRmCtg)
		if @row> 0 
		begin
			select @errmsg= 'RmCtg Combination For Order Rate Round Off Not Defined'
			goto ErrHndlr
		end

		/* ****** Manali 3.9.1 - Rate Round Off Option For Order ****** */
		select @row= count('x') from inserted Ins where HOrdRtRndOffRmCtg <> '' 
			and	exists (select 'x' from Param where PTyp='CPYRT' and PMCd= HOrdRtRndOffRmCtg
			  and (','+PDesc225+',' LIKE '%,G,%' or ','+PDesc225+',' LIKE '%,P,%'  
              or ','+PDesc225+',' LIKE '%,S,%' or ','+PDesc225+',' LIKE '%,L,%' ))			
		if @row> 0 
		begin
			select @errmsg= '''G'', ''P'', ''S'', ''L'' Not Allowed in RmCtg Combination For Order Rate Round Off'
			goto ErrHndlr
		end
		
		/* ****** Manali 3.9.1 - RfId Module in Emr - Y/N ****** */
		select @row= count('x') from inserted where 
				not exists (select 'x' from Param where PTyp= 'YN' and PMCd= HRfIdMod)
		if @row> 0 
		begin
			select @errmsg= 'RfId Module in Emr (Enter Y/N)'
			goto ErrHndlr
		end


		/* ****** Sachin 3.5.0 ****** */
		select @row= count('x') from inserted Ins where hCustomizeRmCtgExp <> '' and 
			exists(select 'x' from Head where HCoCd= Ins.HcoCd and HCd=Ins.HCd and HCustomiseOrd ='Y') 
			and	not exists (select 'x' from Param where PTyp='GLDAS' and PMCd= hCustomizeRmCtgExp)
		if @row> 0 
		begin
			select @errmsg= 'Customize RmCtg Expception Value Not Defined'
			goto ErrHndlr
		end

		select @row= count('x') from inserted Ins where hCustomizeRmCtgExp <> '' and 
			not exists(select 'x' from Head where HCoCd= Ins.HcoCd and HCd=Ins.HCd and HCustomiseOrd ='Y')
		if @row> 0 
		begin
			select @errmsg= 'Customize RmCtg Expception Value Should Be Blank when HCustomiseOrd in Head is Set to ''N'''
			goto ErrHndlr
		end

		/****** HCmLabAs = PMCd('GLDAS') ******/	
		select @row= count('x') from inserted where 
				not exists (select 'x' from Param where PTyp='GLDAS' and PMCd= hCmLabAs)
		if @row> 0 
		begin
			select @errmsg= 'Labour As Value Not Defined'
			goto ErrHndlr
		end

		/****** HCmIWtEqOrd = PMCd('YN') ******/		
		select @row= count('x') from inserted where 
				not exists (select 'x' from Param where PTyp='YN' and PMCd= hCmIWtEqOrd)
		if @row> 0 
		begin
			select @errmsg= 'Invoice Wts From Order (Enter Y/N)'
			goto ErrHndlr
		end

		/****** If hCmIWtEqOrd= 'Y' then hCmIWtFrOrd = Blank 
			else hCmIWtFrOrd = Blank or valid PMCd('IWTFRORD') ******/		
		select @row= count('x') from inserted where hCmIWtEqOrd='N' and hCmIWtFrOrd <> '' and 
				not exists (select 'x' from Param where PTyp='IWTFRORD' and PMCd= hCmIWtFrOrd)
		if @row> 0 
		begin
			select @errmsg= 'Wts From Order Not Defined'
			goto ErrHndlr
		end

		select @row= count('x') from inserted where hCmIWtEqOrd= 'Y' and hCmIWtFrOrd<> '' 
		if @row> 0 
		begin
			select @errmsg= 'Wts From Order Should Not Be Entered'
			goto ErrHndlr
		end

		/****** hCmIGldRtEqOrd = PMCd('YN') ******/		
		select @row= count('x') from inserted where 
				not exists (select 'x' from Param where PTyp= 'YN' and PMCd= hCmIGldRtEqOrd)
		if @row> 0 
		begin
			select @errmsg= 'Invoice Gold Rate From Order (Enter Y/N)'
			goto ErrHndlr
		end

		/****** hCmDiaRtFA = PMCd('FA') ******/		
		/*** (Jen 2.14) ***/
		select @row= count('x') from inserted where hCmDiaRtFA<> 'A'
		/*** (Jen 2.14) ***/

		/*** (Bef 2.14) 
		select @row= count('x') from inserted where 
			not exists (select 'x' from Param where PTyp= 'FA' and PMCd= hCmDiaRtFA)
		(Bef 2.14) ***/
		if @row> 0 
		begin
			select @errmsg= 'Enter Diamond Rate Favourable/Actual'
			goto ErrHndlr
		end

		/****** hCmCSRtFA = PMCd('FA') ******/		
		/*** (Jen 2.14) ***/
		select @row= count('x') from inserted where hCmCSRtFA<> 'A' 
		/*** (Jen 2.14) ***/

		/*** (Bef 2.14) 
		select @row= count('x') from inserted where 
			not exists (select 'x' from Param where PTyp='FA' and PMCd= hCmCSRtFA)
		(Bef 2.14) ***/
		if @row> 0 
		begin
			select @errmsg= 'Enter Colour Stone Rate Favourable/Actual'
			goto ErrHndlr
		end

		/****** hCmFixPrc = PMCd('YN') ******/		
		select @row=count('x') from inserted where 
			not exists (select 'x' from Param where PTyp='YN' and PMCd= hCmFixPrc)
		if @row>0 
		begin
			select @errmsg='Fixed Price (Enter Y/N)'
			goto ErrHndlr
		end

		/****** hDmPrdSeq = PMCd('PRDSEQ') ******/		
		select @row= count('x') from inserted where 
			not exists (select 'x' from vParam where 
				vPCoCd= HCoCd and vPTyp= 'PRDSEQ' and vPMCd= hDmPrdSeq)
		if @row> 0 
		begin
			select @errmsg='Production Sequence Not Defined'
			goto ErrHndlr
		end

		/****** hDmVaCtg = PMCd('VACTG') ******/	
		select @row=count('x') from inserted where 
			not exists (select 'x' from Param where PTyp='VACTG' and PMCd= hDmVaCtg)
		if @row>0 
		begin
			select @errmsg='Value Addition Category Not Defined'
			goto ErrHndlr
		end

		/****** hOmPrtCd = PMCd('PRTCD') ******/		
		select @row=count('x') from inserted where 
			not exists (select 'x' from Param where PTyp='PRTCD' and PMCd= hOmPrtCd)
		if @row>0 
		begin
			select @errmsg='Order Priority Code Not Defined'
			goto ErrHndlr
		end

		/****** hBOpnLoc = LocCd from Loc(LocTyp= 'P') ******/		
		select @row=count('x') from inserted where hBOpnLoc<> '' and 
		not exists (select 'x' from Loc where LocCoCd= HCoCd and LocTyp= 'P' and LocCd= hBOpnLoc)
		if @row>0 
		begin
			select @errmsg='Bag Opening Location Not Defined'
			goto ErrHndlr
		end

		/****** Manali 3.5.0 - 02/04/09 - HMinDiaProcess = LocCd from Loc(LocTyp= 'P') ******/		
/*		select @row=count('x') from inserted where HMinDiaProcess<> '' and 
		not exists (select 'x' from Loc where LocCoCd= HCoCd and LocTyp= 'P' and LocCd= HMinDiaProcess)
		if @row>0 
		begin
			select @errmsg='Process Location for applying minimun Diamond Tolerance Not Defined'
			goto ErrHndlr
		end
*/
		/****** hBOpnLoc = LocCd from Loc(LocTyp= 'P') ******/		
		/*select @row=count('x') from inserted where HCoCd<> @wSelfCoCd and hBOpnLoc= '' 
		if @row> 0 
		begin
			select @errmsg= 'Bag Opening Location Should Be Entered'
			goto ErrHndlr
		end
		*/

		/****** hBChr = PMCd('BCHR') ******/		
		select @row=count('x') from inserted where 
			not exists (select 'x' from Param where PTyp= 'BCHR' and PMCd= hBChr)
		if @row>0 
		begin
			select @errmsg= 'Bag Character Not Defined'
			goto ErrHndlr
		end

		/****** hTdAck = PMCd('YN') ******/		
		select @row=count('x') from inserted where 
			not exists (select 'x' from Param where PTyp= 'YN' and PMCd= hTdAck)
		if @row>0 
		begin
			select @errmsg= 'Acknowledgement Required (Enter Y/N)'
			goto ErrHndlr
		end

		/****** hFgPlChr = PMCd('PLCHR') ******/		
		select @row=count('x') from inserted where 
			not exists (select 'x' from Param where PTyp= 'PLCHR' and PMCd= hFgPlChr)
		if @row>0 
		begin
			select @errmsg= 'Packing List Character Not Defined'
			goto ErrHndlr
		end

		/****** hYy = PMCd('YY') ******/		
		select @row=count('x') from inserted where 
			not exists (select 'x' from Param where PTyp= 'YY' and PMCd= hYy)
		if @row>0 
		begin
			select @errmsg= 'Year Not Defined'
			goto ErrHndlr
		end

		/****** hZWtEq = PMCd('ZWTEQ') ******/		
		select @row=count('x') from inserted where 
			not exists (select 'x' from Param where PTyp= 'ZWTEQ' and PMCd= hZWtEq)
		if @row>0 
		begin
			select @errmsg= 'Customs Weight Option Not Defined'
			goto ErrHndlr
		end

		/****** hOrdDsgGr = PMCd('ORDDSGGR') ******/		
		select @row=count('x') from inserted where 
			not exists (select 'x' from Param where PTyp= 'ORDDSGGR' and PMCd= hOrdDsgGr)
		if @row>0 
		begin
			select @errmsg= 'Order Design Group Option Not Defined'
			goto ErrHndlr
		end

		/****** hOrdRmGr = PMCd('ORDRMGR') ******/		
		select @row=count('x') from inserted where 
			not exists (select 'x' from Param where PTyp= 'ORDRMGR' and PMCd= hOrdRmGr)
		if @row>0 
		begin
			select @errmsg= 'Order Raw Material Group Option Not Defined'
			goto ErrHndlr
		end

		/****** hOrdLabGr = PMCd('ORDLABGR') ******/		
		select @row=count('x') from inserted where 
			not exists (select 'x' from Param where PTyp= 'ORDLABGR' and PMCd= hOrdLabGr)
		if @row>0 
		begin
			select @errmsg= 'Order Labour Group Option Not Defined'
			goto ErrHndlr
		end

		/****** hOthLocBag = PMCd('YN') ******/		
		select @row=count('x') from inserted where 
			not exists (select 'x' from Param where PTyp= 'YN' and PMCd= hOthLocBag)
		if @row>0 
		begin
			select @errmsg= 'Show Other Loc Bags (Enter Y/N)'
			goto ErrHndlr
		end

		/****** hOdAmdQty = PMCd('YN') ******/		
		select @row=count('x') from inserted where 
			not exists (select 'x' from Param where PTyp= 'YN' and PMCd= hOdAmdQty)
		if @row>0 
		begin
			select @errmsg= 'Show Amended Qty Or Entered Qty In Ord Entry Summary Option (Enter Y/N)'
			goto ErrHndlr
		end

		/****** hZValEqInv = PMCd('YN') ******/		
		select @row=count('x') from inserted where 
			not exists (select 'x' from Param where PTyp= 'YN' and PMCd= hZValEqInv)
		if @row>0 
		begin
			select @errmsg= 'Specify Whether Customs Value Should Be Equal To Invoice Value (Enter Y/N)'
			goto ErrHndlr
		end

		/****** hCmCdHlp = PMCd('YN') ******/		
		select @row=count('x') from inserted where hCmCdHlp<> '' and 
			not exists (select 'x' from Param where PTyp= 'YN' and PMCd= hCmCdHlp)
		if @row>0 
		begin
			select @errmsg= 'Sort Customer Code Help By Name? (Enter Y/N)'
			goto ErrHndlr
		end

		/****** HRndOffFOB = PMCd('YN') ******/		
		select @row=count('x') from inserted where 
			not exists (select 'x' from Param where PTyp= 'YN' and PMCd= HRndOffFOB)
		if @row>0 
		begin
			select @errmsg= 'Specify Whether To Round Off The Inv & Custom FOB Values At The Dsg Level In Invoicing (Enter Y/N)'
			goto ErrHndlr
		end

		/****** HShowActWt = PMCd('YN') ******/		
		select @row=count('x') from inserted where 
			not exists (select 'x' from Param where PTyp= 'YN' and PMCd= HShowActWt)
		if @row>0 
		begin
			select @errmsg= 'Specify Whether To Show The Actual Weight Field In The Rm Grid In Inv Dsg Entry (Enter Y/ N)'
			goto ErrHndlr
		end

		/****** HSrtBasePurVch = PMCd('SRTPRHLP') ******/		
		select @row= count('x') from inserted where 
			not exists (select 'x' from Param where PTyp= 'SRTPRHLP' and PMCd= HSrtBasePurVch)
		if @row> 0 
		begin
			select @errmsg= 'Base Custom Purchase Voucher Help Sort Option Not Defined'
			goto ErrHndlr
		end

		/* ***** Manali 3.5.0 - 26/11/08 - HHideCommYN = PMCd('YN') */
		select @row= count('x') from inserted where 
				not exists (select 'x' from Param where PTyp='YN' and PMCd= HHideCommYN)
		if @row> 0 
		begin
			select @errmsg= 'Commission Self Payable in Invoice? Enter ''Y''es / ''N''o'
			/*select @errmsg= 'Hide Commission in Invoice Printing ? Enter ''Y''es / ''N''o'*/
			goto ErrHndlr
		end
		
		/* ***** Manali 3.8.1 - HSelfDscYN = PMCd('YN') */
		select @row= count('x') from inserted where 
				not exists (select 'x' from Param where PTyp='YN' and PMCd= HSelfDscYN)
		if @row> 0 
		begin
			select @errmsg= 'Discount Self Payable in Invoice? Enter ''Y''es / ''N''o'
			goto ErrHndlr
		end

		/* ***** Manali 3.5.0 - 26/11/08 - HSelfDeductYN = PMCd('YN') */
		select @row= count('x') from inserted where 
				not exists (select 'x' from Param where PTyp='YN' and PMCd= HSelfDeductYN)
		if @row> 0 
		begin
			select @errmsg= 'Deduction Self Payable in Invoice? Enter ''Y''es / ''N''o'
			goto ErrHndlr
		end

		select @row= count('x') from inserted ins where ins.hCoCD <> @wSelfCoCd and 
			ins.HMaxDsgViewRec<> (select HMaxDsgViewRec from Head where HCoCd=  @wSelfCoCd) 
		if @row> 0 
		begin
			select @errmsg= 'The ''HMaxDsgViewRec'' value For This Company Should Be Same As '+ @wSelfCoCd
			goto ErrHndlr
		end

		Select @row=count('x') from Inserted where HMaxDsgViewRec not between 1 and 9999
		if @row > 0
		Begin 
			Select @errmsg = 'Maximum No. Of Records allowed for Desgin View should be Greater than 0 and Less than 10000'
			goto ErrHndlr
		End

		/***** Manali 3.8.0 - Invoice Design Sort By *****/
		select @row=count('x') from inserted where HInSrt<>'' and 
			not exists (select 'x' from Param where PTyp= 'INSRT' and PMCd= HInSrt)
		if @row>0 
		begin
			select @errmsg= 'Enter Valid Invoice Design Sort By Option or Keep It Blank'
			goto ErrHndlr
		end

		/****** hShowRepLogo = PMCd('YN') ******/		
		select @row= count('x') from inserted where 
				not exists (select 'x' from Param where PTyp='YN' and PMCd= hShowRepLogo)
		if @row> 0 
		begin
			select @errmsg= 'Show Company Logo in Reports (Enter Y/N)'
			goto ErrHndlr
		end

		/****** hInRepWtValAs = PMCd('INWT') ******/		
		select @row=count('x') from inserted where 
			not exists (select 'x' from Param where PTyp= 'INWT' and PMCd= hInRepWtValAs)
		if @row>0 
		begin
			select @errmsg= 'Show Invoice Or Customs Wt & Amt In Inv Printing Rep (Enter Valid Value)'
			goto ErrHndlr
		end

		/****** hInRepWtFld = PMCd('YN') ******/		
		select @row=count('x') from inserted where 
			not exists (select 'x' from Param where PTyp= 'YN' and PMCd= hInRepWtFld)
		if @row>0 
		begin
			select @errmsg= 'Show Wt Column In Inv Printing Rep (Enter Y/ N)'
			goto ErrHndlr
		end

		/****** hInRepMetLs = PMCd('YN') ******/		
		select @row=count('x') from inserted where 
			not exists (select 'x' from Param where PTyp= 'YN' and PMCd= hInRepMetLs)
		if @row>0 
		begin
			select @errmsg= 'Show Metal Loss Summary In Inv Printing Rep (Enter Y/ N)'
			goto ErrHndlr
		end

		/****** hInRepCnvRt = PMCd('YN') ******/		
		select @row=count('x') from inserted where 
			not exists (select 'x' from Param where PTyp= 'YN' and PMCd= hInRepCnvRt)
		if @row>0 
		begin
			select @errmsg= 'Show Conv Rt & Rupee Equivalent In Inv Printing Rep (Enter Y/ N)'
			goto ErrHndlr
		end

		/****** hInRepRtFld = PMCd('YN') ******/		
		select @row= count('x') from inserted where 
			not exists (select 'x' from Param where PTyp= 'YN' and PMCd= hInRepRtFld)
		if @row> 0 
		begin
			select @errmsg= 'Show Rate Column In Inv Printing Rep (Enter Y/ N)'
			goto ErrHndlr
		end

		/****** hPtrYN = PMCd('YN') ******/		
		select @row= count('x') from inserted where 
			not exists (select 'x' from Param where PTyp= 'YN' and PMCd= hPtrYN)
		if @row> 0 
		begin
			select @errmsg= 'Allow The User To Change The ''Ptr Y/N'' Option In The Param Of Type ''RMSCTG'' (Enter Y/ N)'
			goto ErrHndlr
		end

		/****** hPdcSfx = PMCd('PDCSFX') ******/		
		select @row= count('x') from inserted where 
			not exists (select 'x' from Param where PTyp= 'PDCSFX' and PMCd= hPdcSfx)
		if @row> 0 
		begin
			select @errmsg= 'PDC Suffix Generation Method Not Defined'
			goto ErrHndlr
		end

			/****** HRndOrdSalPrc = PMCd('RNDOPT') ******/	
		select @row= count('x') from inserted where 
				not exists (select 'x' from Param where PTyp= 'RNDOPT' and PMCd= HRndOrdSalPrc)
		if @row> 0 
		begin
			select @errmsg= 'Round Off Option For Sales Price Not Defined'
			goto ErrHndlr
		end

		/****** HOrdCpyRtOpt = PMCd('CPYRT') ******/	
		select @row= count('x') from inserted where 
				not exists (select 'x' from Param where PTyp= 'CPYRT' and PMCd= HOrdCpyRtOpt)
		if @row> 0 
		begin
			select @errmsg= 'Copy Rate Option For Order/ Design Copy Not Defined'
			goto ErrHndlr
		end

		/****** HShowPrdRmDets = PMCd('YN') ******/	
		select @row= count('x') from inserted where 
				not exists (select 'x' from Param where PTyp= 'YN' and PMCd= HShowPrdRmDets)
		if @row> 0 
		begin
			select @errmsg= 'Show Production Rm Qty && Wt In Design And Order Entries (Enter Y/ N)'
			goto ErrHndlr
		end

		/****** HShowPrdRmDets = 'N' for CoCd = 'ZZZ' ******/	
		/*### select @row= count('x') from inserted where HCoCd = @wSelfCoCd and HShowPrdRmDets <> 'N'
		if @row> 0 
		begin
			select @errmsg= 'Show Production Rm Qty && Wt In Design And Order Entries Should Be 'N''
			goto ErrHndlr
		end ###*/

		/****** hGrsWtInBsYN = PMCd('YN') ******/		
		select @row= count('x') from inserted where 
				not exists (select 'x' from Param where PTyp='YN' and PMCd= hGrsWtInBsYN)
		if @row> 0 
		begin
			select @errmsg= 'Show Gross Weight in Bag Split (Enter Y/N)'
			goto ErrHndlr
		end

		/****** hExpLocal = PMCd('YN') ******/		
		select @row= count('x') from inserted where 
			not exists (select 'x' from Param where PTyp='YN' and PMCd= hExpLocal)
		if @row> 0 
		begin
			select @errmsg= 'Specify Whether Export Is From SEEPZ Or Outside? (Enter Y/N)'
			goto ErrHndlr
		end

		/******	Manali 3.6.0 - 30/10/09 -		hInNotionRtYn = PMCd('YN') ******/		
		select @row= count('x') from inserted where 
			not exists (select 'x' from Param where PTyp='YN' and PMCd= hInNotionRtYn)
		if @row> 0 
		begin
			select @errmsg= 'Specify Use Notional Rate in Custom Matching? (Enter Y/N)'
			goto ErrHndlr
		end

		/* Sachin 4.1.0.0
		select @row= count('x') from inserted where 
			hExpLocal='N' and hInNotionRtYn='Y'
		if @row> 0 
		begin
			select @errmsg= 'Use Notional Rate Cannot Be ''Yes'' When DTA is No'
			goto ErrHndlr
		end
		*/

		--4.1.2 - Geeta
		select @row= count('x') from inserted where 
			not exists (select 'x' from Param where PTyp='YN' and PMCd= HEnfFrBLocInFgYN)
		if @row> 0 
		begin
			select @errmsg= 'Enforce Bag Location In Finished Goods? (Enter Y/N)'
			goto ErrHndlr
		end

		/****** hPpcYN = PMCd('YN') ******/		
		select @row= count('x') from inserted where 
			not exists (select 'x' from Param where PTyp='YN' and PMCd= hPpcYN)
		if @row> 0 
		begin
			select @errmsg= 'Specify Whether To Show The PPC Module Or Not'
			goto ErrHndlr
		end

		/****** hMaxLineUtil > 0, hMaxOrdQty > 0, hMinOrdPrcnt > 0, hMinOrdQty > 0 ******/	
		select @row= count('x') from inserted where 
			Round(hMaxLineUtil, 3) <= 0 or Round(hMaxOrdQty, 1) <= 0 
			/*or Round(hMinOrdPrcnt, 3) <= 0 or Round(hMinOrdQty, 1) <= 0 */
		if @row> 0 
		begin
			select @errmsg= 'Max Line Utilization %/ Max Ord Qty Should Be > 0'
			goto ErrHndlr
		end

		/****** hMaxOrdQty >= hMinOrdQty ******/	
		select @row= count('x') from inserted where hMaxOrdQty< hMinOrdQty 
		if @row> 0 
		begin
			select @errmsg= 'Max Ord Qty Should Be >= Min Ord Qty'
			goto ErrHndlr
		end

		/****** hLine = vPMCd('LINE') ******/ 
		select @row= count('x') from inserted where 
			not exists (select 'x' from vParam where vPCoCd= HCoCd and vPTyp= 'LINE' and vPMCd= hLine and vPMCd<> 'ZZZ')
		if @row> 0 
		begin
			select @errmsg= 'Line Not Defined/ Line Cannot Be ''ZZZ'' '
			goto ErrHndlr
		end

		/****** hWH = vPMCd('WH') ******/ 
		select @row= count('x') from inserted where 
			not exists (select 'x' from Param where PTyp= 'WH' and PMCd= hWH)
		if @row> 0 
		begin
			select @errmsg= 'W-H Set Not Defined'
			goto ErrHndlr
		end

		/*** HFgSubLoc  =  vPMCd('FGSUBLOC') ***/
		Select @Row= Count('x') From Inserted Ins Where 
  			Not Exists (Select vPMCd From vParam Where vPCoCd= HCoCd and 
					vPTyp= 'FGSUBLOC' And vPMCd = HFgSubLoc)
	  	If @Row> 0
 		begin
			Select @ErrMsg= 'Default Fg Sub Location Not Defined'
       			GoTo ErrHndlr
  		end
		
		/*** urmi Ech Start***/
		/****** hNegStkYn = PMCd('YN') ******/		
		select @row= count('x') from inserted where 
				not exists (select 'x' from Param where PTyp='YN' and PMCd= hNegStkYN)
		if @row> 0 
		begin
			select @errmsg= 'Allow Negative stocks without SEO PAssword (Enter Y/N)'
			goto ErrHndlr
		end

		/****** hAvgSRtYn = PMCd('YN') ******/		
		select @row= count('x') from inserted where 
				not exists (select 'x' from Param where PTyp='YN' and PMCd= hAvgSRtYn)
		if @row> 0 
		begin
			select @errmsg= 'Implement Average Stock Rate (Enter Y/N)'
			goto ErrHndlr
		end

		/****** hSizeStkRt = PMCd('YN') ******/		
		select @row= count('x') from inserted where 
				not exists (select 'x' from Param where PTyp='YN' and PMCd= hSizeStkRt)
		if @row> 0 
		begin
			select @errmsg= 'Enable RmSz and Stk Rt in Invoicing (Enter Y/N)'
			goto ErrHndlr
		end

		/*** if HAvgSRtYn = 'Y' then hSizeStkrt should be 'Y' ***/
		select @row=count('x') from inserted where hAvgSRtYn='Y' and hSizeStkRt <> 'Y' 
		if @row>0 
		begin
			select @errmsg='Size and Stock Rt in Invoicing should be enabled to implement Average Stock Rate'
			goto ErrHndlr
		end

		/**** Shilpa Emr 2.04 ****/
		/*** Invoice type = valid PMCd('INVTYP') ***/
		select @row= count('x') from inserted ins where HInvTyp Not In (Select PMCd 
				From Param where PTyp= 'INVTYP')
		if @row> 0 
		begin
			select @errmsg= 'Invoice Type Not Defined'
			goto ErrHndlr
		end

		/*** Invalid Value Addition Option ***/
		select @row= count('x') from inserted ins where Not Exists(Select 'x'  
				From Param where PTyp= 'YN' and PMCd= hValAddMtlColYN)
		if @row> 0 
		begin
			select @errmsg= 'Show Value Addition On Metal Column In Value Addition Report? Enter Yes Or No (Y/N)?'
			goto ErrHndlr
		end

		/***Invalid BOE Details Option ***/
		select @row= count('x') from inserted ins where 
			hValBOEDetYN Not In (Select PmCd From Param where Ptyp= 'YN')
		if @row> 0 
		begin
			select @errmsg= 'Show BOE Details In Value Addition Annexur Report? Enter Yes Or No (Y/N)?'
			goto ErrHndlr
		end

		/***Invalid BOE Details Option ***/ 
		select @row= count('x') from inserted ins where 
			hBagNoEditYN Not In (Select PMCd From Param where Ptyp= 'YN')
		if @row> 0 
		begin
			select @errmsg= 'Allow Bag No. Edit? Enter Yes Or No (Y/N)?'
			goto ErrHndlr
		end
		/**** Invalid Allow Production Points Editing option ***/
		select @row=count('x') from inserted where 
			not exists (select 'x' from Param where PTyp='YN' and PMCd= hPrdPtsYN)
		if @row>0 
		begin
			select @errmsg='Allow Editing of Production Points? (Enter Y/N)'
			goto ErrHndlr
		end

		/****** HGenExpNoYN = PMCd('YN') ******/		
		select @row=count('x') from inserted where 
				not exists (select 'x' from Param where PTyp='YN' and PMCd= hGenExpNoYN)
		if @row>0 
		begin
			select @errmsg='Generate Export Number(Enter Y/N)'
			goto ErrHndlr
		end

		/****** 4.1.0.0 hShowBgHistDetsYN = PMCd('YN') ******/		
		select @row=count('x') from inserted where hShowBgHistDetsYN<>'' And
				not exists (select 'x' from Param where PTyp='YN' and PMCd= hShowBgHistDetsYN)
		if @row>0 
		begin
			select @errmsg='Show Details in Bag History Report (Enter Y/N)'
			goto ErrHndlr
		end

		/****** 4.1.2.0 HHolChkYn = PMCd('YN') ******/		
		select @row=count('x') from inserted where 
				not exists (select 'x' from Param where PTyp='YN' and PMCd= HHolChkYn)
		if @row>0 
		begin
			select @errmsg='Apply Holiday Checks in Entries (Enter Y/N)'
			goto ErrHndlr
		end


		/****** HCtbMod = PMCd('YN') ******/		
		select @row= count('x') from inserted where 
			not exists (select 'x' from Param where PTyp= 'YN' and PMCd= HCtbMod)
		if @row> 0 
		begin
			select @errmsg= 'Show Contractor Bill Module In The System? (Enter Y/ N)'
			goto ErrHndlr
		end

		/****** Sachin 3.3.0 (23-06-08) - HInIRtWrtStkYN = PMCd('YN') ******/		
		select @row= count('x') from inserted where 
				not exists (select 'x' from Param where PTyp='YN' and PMCd= HInIRtWrtStkYN)
		if @row> 0 
		begin
			select @errmsg= 'Invoice Rate of Diamond and Color Stones With Respect to Stock Rate YN (Enter Y/N)'
			goto ErrHndlr
		end

		/****** Manali 3.6.0 - 27/10/09 - HInXIRtWrtStkYN = PMCd('YN') ******/		
		select @row= count('x') from inserted where 
				not exists (select 'x' from Param where PTyp='YN' and PMCd= HInXIRtWrtStkYN)
		if @row> 0 
		begin
			select @errmsg= 'Invoice Rate of Accessories With Respect to Stock Rate YN (Enter Y/N)'
			goto ErrHndlr
		end

		/****** Sachin 3.11.0 - HMetIRtWrtStkYN = PMCd('YN') ******/		
		select @row= count('x') from inserted where 
				not exists (select 'x' from Param where PTyp='YN' and PMCd= HMetIRtWrtStkYN)
		if @row> 0 
		begin
			select @errmsg= 'Invoice Rate of Metal(s) With Respect to Stock Rate YN (Enter Y/N)'
			goto ErrHndlr
		end

		select @row= count('x') from inserted where 
				not exists (select 'x' from Param where PTyp='YN' and PMCd= hSetTypOnlyFrOrdYN)
		if @row> 0 
		begin
			select @errmsg= 'Setting type only from Order YN (Enter Y/N)'
			goto ErrHndlr
		end

		select @row= count('x') from inserted where 
				not exists (select 'x' from Param where PTyp='YN' and PMCd= hCpyActCalcWtYN)
		if @row> 0 
		begin
			select @errmsg= 'Copy Calculated Weight in Design Master (Enter Y/N)'
			goto ErrHndlr
		end

		/****** Sachin 3.3.0 (23-06-08) - HChkVldRmLnYN = PMCd('YN') ******/		
		select @row= count('x') from inserted where 
				not exists (select 'x' from Param where PTyp='YN' and PMCd= HChkVldRmLnYN)
		if @row> 0 
		begin
			select @errmsg= 'Check Valid RM Ln From Master(Enter Y/N)'
			goto ErrHndlr
		end

		/****** Sachin 3.3.0 (23-06-08) - HExcelImpInOrdYN = PMCd('YN') ******/		
		select @row= count('x') from inserted where 
				not exists (select 'x' from Param where PTyp='YN' and PMCd= HExcelImpInOrdYN)
		if @row> 0 
		begin
			select @errmsg= 'Enable Excel Import In Order Entry (Enter Y/N)'
			goto ErrHndlr
		end

		/****** Sachin 3.3.0 (19-07-08) - Travel Excel Option = PMCd('YN') ******/		
		select @row= count('x') from inserted where 
				not exists (select 'x' from Param where PTyp='YN' and PMCd= HTravelExlYN)
		if @row> 0 
		begin
			select @errmsg= 'Enable Travel Excel Module (Enter Y/N)'
			goto ErrHndlr
		end

		/****** Manali 3.8.0 - Metlr Cd ******/		
		select @row= count('x') from inserted where HMetlrCd<>'' and
				not exists (select 'x' from Param where PTyp='METLR' and PMCd= HMetlrCd)
		if @row> 0 
		begin
			select @errmsg= 'Enter Valid Metler Code As Default'
			goto ErrHndlr
		end

		/****** HCtCd = '' If HCtbMod= 'N' ******/	
		select @row= count('x') from inserted where HCtCd<> '' and HCtbMod= 'N'  
		if @row> 0 
		begin
			select @errmsg= 'Default Contractor Cannot Be Entered If CTB Module Is Not To Be Included In The System'
			goto ErrHndlr
		end

		/****** HCtCd = CmCd from CustMst(CmCtg= 'T') ******/	
		select @row= count('x') from inserted where HCtCd<> '' and 
			not exists (select 'x' from CustMst where CmCtg= 'T' and CmCd= HCtCd) 
		if @row> 0 
		begin
			select @errmsg= 'Default Contractor Code Not Defined'
			goto ErrHndlr
		end

		/****** HOdGmChk = PMCd('CPYRT') ******/	
		select @row= count('x') from inserted where 
				not exists (select 'x' from Param where PTyp= 'CPYRT' and PMCd= HOdGmChk)
		if @row> 0 
		begin
			select @errmsg= 'Gram Check Option Not Defined'
			goto ErrHndlr
		end

		/*********      urmi  2.04-3 Option defined for zzz r not to changed for other comp codes*/
		select @row= count('x') from inserted ins where ins.hCoCD <> @wSelfCoCd and 
			ins.hPtrYn<> (select hPtrYn from Head where HCoCd=  @wSelfCoCd) 
		if @row> 0 
		begin
			select @errmsg= 'The ''Ptr Y/N'' Option For This Company Should Be Same As '+ @wSelfCoCd
			goto ErrHndlr
		end
		/*********      urmi  2.04-3 Option defined for zzz r not to changed for other comp codes*/

		/******  FixPrc, GldAs, IWtEqOrd, IWtFrOrd, IGldRtEqOrd, DiaRtFA, CSRtFA of all rec should Be Equal to that of @wSelfCoCd ******/	
		select @row= count('x') from inserted ins 
			join Head z on z.HCoCd= @wSelfCoCd and z.HCd= @wSelfCd 
			where ins.HCoCd<> @wSelfCoCd and 
				(ins.hCmFixPrc<> z.hCmFixPrc or ins.hCmGldAs<> z.hCmGldAs Or Ins.hCmLabAs<>z.hCmLabAs 
				or ins.hCmIWtEqOrd<> z.hCmIWtEqOrd or ins.hCmIWtFrOrd<> z.hCmIWtFrOrd 
				or ins.hCmIGldRtEqOrd<> z.hCmIGldRtEqOrd or ins.hCmDiaRtFA<> z.hCmDiaRtFA
				or ins.hCmCsRtFA<> z.hCmCsRtFA or ins.hPPCYN<> z.hPPCYN 
				or ins.HPDCSfx<> z.HPDCSfx or ins.HRndOrdSalPrc<> z.HRndOrdSalPrc 
				or ins.HCtbMod<> z.HCtbMod 
				or ins.HShowPrdRmDets<> z.HShowPrdRmDets) 
		if @row> 0 
		begin
			select @errmsg= 'For This Company Options Such As FixPrc, GldAs, InvWtEqualsOrd, InvWtFromOrd, InvGoldRtEqualsOrd, Dia/CS Rt Favourable, SfxGenMethod, RndOrdSalPrc, ShowCTB, ShowPrdRmDets Should Be Same As '+ @wSelfCoCd
			goto ErrHndlr
		end

		/****** HPDCSfx of all rec should Be equal to HPDCSfx of @wSelfCoCd ******/	
		select @row= count('x') from inserted where HCoCd<> @wSelfCoCd and 
			HPDCSfx<> IsNull((select HPDCSfx from Head where HCoCd= @wSelfCoCd 
					and HCd= @wSelfCd) , '')
		if @row> 0 
		begin
			select @errmsg= 'Suffix Generation Option For This Company Should Be Same As '+ @wSelfCoCd
			goto ErrHndlr
		end

		/****** HRndOrdSalPrc of all rec should Be equal to HRndOrdSalPrc of @wSelfCoCd ******/	
		select @row= count('x') from inserted where HCoCd<> @wSelfCoCd and 
			HRndOrdSalPrc<> IsNull((select HRndOrdSalPrc from Head where HCoCd= @wSelfCoCd 
					and HCd= @wSelfCd) , '')
		if @row> 0 
		begin
			select @errmsg= 'Sales Price Rounding Option For This Company Should Be Same As '+ @wSelfCoCd
			goto ErrHndlr
		end

		/****** HShowPrdRmDets of all rec should Be equal to HShowPrdRmDets of @wSelfCoCd ******/	
		select @row= count('x') from inserted where HCoCd<> @wSelfCoCd and 
			HShowPrdRmDets<> IsNull((select HShowPrdRmDets from Head where HCoCd= @wSelfCoCd 
					and HCd= @wSelfCd) , '')
		if @row> 0 
		begin
			select @errmsg= 'Show Production Rm Dets (Wt And Qty) Option For This Company Should Be Same As '+ @wSelfCoCd
			goto ErrHndlr
		end

		/****** HCtbMod of all rec should Be equal to HCtbMod of @wSelfCoCd ******/	
		select @row= count('x') from inserted where HCoCd<> @wSelfCoCd and 
			HCtbMod<> IsNull((select HCtbMod from Head where HCoCd= @wSelfCoCd 
					and HCd= @wSelfCd) , '')
		if @row> 0 
		begin
			select @errmsg= ' ''Show Contractor Module'' Option For This Company Should Be Same As '+ @wSelfCoCd
			goto ErrHndlr
		end

		/* ****** Sachin ****** */
		/****** HSlvModYN of all rec should Be equal to HSlvModYN of @wSelfCoCd ******/	
		select @row= count('x') from inserted where HCoCd<> @wSelfCoCd and 
			HSlvModYN <> IsNull((select HSlvModYN from Head where HCoCd= @wSelfCoCd 
					and HCd= @wSelfCd) , '')
		if @row> 0 
		begin
			select @errmsg= ' ''Silver Module'' Option For This Company Should Be Same As '+ @wSelfCoCd
			goto ErrHndlr
		end

		/* Manali 3.03 */
		/****** HPdModYN of all rec should Be equal to HPdModYN of @wSelfCoCd ******/	
		select @row= count('x') from inserted where HCoCd<> @wSelfCoCd and 
			HPdModYN <> IsNull((select HPdModYN from Head where HCoCd= @wSelfCoCd 
					and HCd= @wSelfCd) , '')
		if @row> 0 
		begin
			select @errmsg= ' ''Palaadium Module'' Option For This Company Should Be Same As '+ @wSelfCoCd
			goto ErrHndlr
		end

		/****** HQWModYN of all rec should Be equal to HQWModYN of @wSelfCoCd ******/	
		select @row= count('x') from inserted where HCoCd<> @wSelfCoCd and 
			HQWModYN <> IsNull((select HQWModYN from Head where HCoCd= @wSelfCoCd 
					and HCd= @wSelfCd) , '')
		if @row> 0 
		begin
			select @errmsg= ' ''QW Module'' Option For This Company Should Be Same As '+ @wSelfCoCd
			goto ErrHndlr
		end

		/****** HModAvlYN of all rec should Be equal to HModAvlYN of @wSelfCoCd ******/	
		select @row= count('x') from inserted where HCoCd<> @wSelfCoCd and 
			HModAvlYN <> IsNull((select HModAvlYN from Head where HCoCd= @wSelfCoCd 
					and HCd= @wSelfCd) , '')
		if @row> 0 
		begin
			select @errmsg= ' ''Model Availability'' Option For This Company Should Be Same As '+ @wSelfCoCd
			goto ErrHndlr
		end
		/* ****** Sachin ****** */

		/* Sachin 2.14.0 - [27-10-2006] - Multi-Company */
		/****** HImpComLoc of all rec should Be equal to HImpComLoc of @wSelfCoCd ******/	
		select @row= count('x') from inserted where HCoCd<> @wSelfCoCd and 
			HImpComLoc <> IsNull((select HImpComLoc from Head where HCoCd= @wSelfCoCd 
					and HCd= @wSelfCd) , '')
		if @row> 0 
		begin
			select @errmsg= ' ''Implement Common Locations Module'' Option For This Company Should Be Same As '+ @wSelfCoCd
			goto ErrHndlr
		end	

		/****** hInRepMetLsRmCd of all rec should Be equal to hInRepMetLsRmCd of @wSelfCoCd - Sachin 2.14.0 Next ******/	
		select @row= count('x') from inserted where HCoCd<> @wSelfCoCd and 
			hInRepMetLsRmCd <> IsNull((select hInRepMetLsRmCd from Head where HCoCd= @wSelfCoCd 
					and HCd= @wSelfCd) , '')
		if @row> 0 
		begin
			select @errmsg= ' ''Inv Rep Metal Loss RmCd'' Option All Companies Should Be Same As '+ @wSelfCoCd
			goto ErrHndlr
		end	

		/* Sachin 2.14.0 - [27-10-2006] - Multi-Company */

		
		/****** Manali 3.4.1 - 22/10/08 - HDsgCatYN of all rec should Be equal to HDsgCatYN of @wSelfCoCd ******/	
		select @row= count('x') from inserted where HCoCd<> @wSelfCoCd and 
			HDsgCatYN <> IsNull((select HDsgCatYN from Head where HCoCd= @wSelfCoCd 
					and HCd= @wSelfCd) , '')
		if @row> 0 
		begin
			select @errmsg= ' ''Design Catalogue Module'' Option For This Company Should Be Same As '+ @wSelfCoCd
			goto ErrHndlr
		end

		/****** Manali 3.6.0 - HOrdExcelRepYn of all rec should Be equal to HOrdExcelRepYn of @wSelfCoCd ******/	
		select @row= count('x') from inserted where HCoCd<> @wSelfCoCd and 
			HOrdExcelRepYn <> IsNull((select HOrdExcelRepYn from Head where HCoCd= @wSelfCoCd 
					and HCd= @wSelfCd) , '')
		if @row> 0 
		begin
			select @errmsg= ' ''Order Printing in Excel Format'' Option For This Company Should Be Same As '+ @wSelfCoCd
			goto ErrHndlr
		end

    /* **** Manali 3.10.0 - Multi Price Quotation Module */
		select @row= count('x') from inserted where HCoCd<> @wSelfCoCd and 
			HMultiPrcQtMod <> IsNull((select HMultiPrcQtMod from Head where HCoCd= @wSelfCoCd 
					and HCd= @wSelfCd) , '')
		if @row> 0 
		begin
			select @errmsg= ' ''Multi Price Quotation'' Option For This Company Should Be Same As '+ @wSelfCoCd
			goto ErrHndlr
		end

		-- 3.11.0
		select @row= count('x') from inserted where HCoCd<> @wSelfCoCd and 
			hRtfrRtChrtYN <> IsNull((select hRtfrRtChrtYN from Head where HCoCd= @wSelfCoCd 
					and HCd= @wSelfCd) , '')
		if @row> 0 
		begin
			select @errmsg= ' ''Refresh from Rate Chart'' Option For This Company Should Be Same As '+ @wSelfCoCd
			goto ErrHndlr
		end

		/****** hMMTc ******/		
		select @row= count('x') from inserted where 
			inserted.hMMTc<> '' and inserted.HCoCd= @wSelfCoCd
		if @row> 0 
		begin
			select @errmsg= 'MMTC Code not to be defined for ''ZZZ'' '
			goto ErrHndlr
		end

		/****************************** Geeta ***** Emr 2.06 ***************************/
		/****** HPrtAllotOnPrdYN = PMCd('YN') ******/		
		select @row= count('x') from inserted where HPrtAllotOnPrdYN<> '' and 
			not exists (select 'x' from Param where PTyp= 'YN' and PMCd= HPrtAllotOnPrdYN)
		if @row> 0 
		begin
			select @errmsg= 'Specify Whether The Priority Allotment Is Based On Production /Export/No Priority Allotment? (Y/N/'')'
			goto ErrHndlr
		end
		/****** HChkGPWtAtIssYN = PMCd('YN') ******/		
		select @row= count('x') from inserted where 
			not exists (select 'x' from Param where PTyp= 'YN' and PMCd= HChkGPWtAtIssYN)
		if @row> 0 
		begin
			select @errmsg= 'Specify Whether to check GP Wt at Issue? (Y/N/'')'
			goto ErrHndlr
		end
		/****** HChkDCWtAtIssYN = PMCd('YN') ******/		
		select @row= count('x') from inserted where 
			not exists (select 'x' from Param where PTyp= 'YN' and PMCd= HChkDCWtAtIssYN)
		if @row> 0 
		begin
			select @errmsg= 'Specify Whether to check Diamond/Color Stone Wt at Issue? (Y/N/'')'
			goto ErrHndlr
		end
		/****** HChkDCQtyAtIssYN = PMCd('YN') ******/		
		select @row= count('x') from inserted where 
			not exists (select 'x' from Param where PTyp= 'YN' and PMCd= HChkDCQtyAtIssYN)
		if @row> 0 
		begin
			select @errmsg= 'Specify Whether to check Diamond/Color Stone Qty at Issue? (Y/N/'')'
			goto ErrHndlr
		end
		/****************************** Geeta ***** Emr 2.06 ***************************/
		/****************************** Geeta ***** Emr 2.07 ***************************/
		/****** HPrtAllotOnPrdYN = PMCd('YN') ******/		
		select @row= count('x') from inserted where  
			not exists (select 'x' from Param where PTyp= 'YN' and PMCd= hFgRmEqOrdRm)
		if @row> 0 
		begin
			select @errmsg= 'Specify Whether to check All RawMaterials Specified in Ord Are there in FgBag?(Enter Y/N)'
			goto ErrHndlr
		end
		/****************************** Geeta ***** Emr 208 ***************************/
		/*******HGldFiness>0***********/
		select @row= count('x') from inserted where  HGldFiness<=0 
				if @row> 0 
		begin
			select @errmsg= 'Gold Fineness Should be >0'
			goto ErrHndlr
		end
		/*******HPlFiness>0***********/
		select @row= count('x') from inserted where  HPlFiness<=0 
				if @row> 0 
		begin
			select @errmsg= 'Platinum Fineness Should be >0'
			goto ErrHndlr
		end
		/* ****** Sachin 2.12 - 06/12/05 - SJM ****** */
		/*******HSlFiness>0***********/
		select @row= count('x') from inserted where  HSlFiness<=0 
				if @row> 0 
		begin
			select @errmsg= 'Silver Fineness Should be >0'
			goto ErrHndlr
		end
		/* ****** Sachin 2.12 - 06/12/05 - SJM ****** */
		/* **** Manali 3.03 Palladium **** */
		/*******HPdFiness>0***********/
		select @row= count('x') from inserted where  HPdFiness<=0 
				if @row> 0 
		begin
			select @errmsg= 'Palladium Fineness Should be >0'
			goto ErrHndlr
		end
		/* **** Manali 3.03 Palladium **** */

		/****************************** Geeta ***** Emr 208 ***************************/
		/* **** Zubin 211 **** */ 
		select @row= count('x') from inserted where 
			Not Exists (select 'x' from Param where PTyp= 'YN' and PMCd= hFluteBagYN and PSCd= '') 
		if @row> 0 
		begin 
			select @errmsg= 'Specify Whether to use Flute Bags?(Enter Y/N)' 
			goto ErrHndlr 
		end 

		/* hFluteBagYN cannot be changed to 'N' if Flute bag exists in BChr */ 
		select @row= count('x') from inserted ins where ins.hFluteBagYN= 'N' and 
			Not Exists (select 'x' from deleted del where ins.hCoCd= del.hCoCd and ins.hCd= del.hCd and 
				ins.hFluteBagYN= del.hFluteBagYN) and 
			Exists (select 'x' from Param where PTyp= 'BCHR' and PSCd= '' and PValue= 'Y') 
		if @row> 0 
		begin 
			select @errmsg= ' ''Allow Flute Bags'' Option Cannot Be ''N'' When Flute Bag Character(s) Exist' 
			goto ErrHndlr 
		end 

		select @row=count('x') from inserted where 
				Not Exists (select 'x' from Param where PTyp='YN' and PMCd= HOrdPrdPts and PSCd= '')
		if @row>0 
		begin
			select @errmsg= 'Restrict Production Points On Order RmQty (Enter Y/N)' 
			goto ErrHndlr
		End 

		select @row=count('x') from inserted where HAlyCnvDiff< 0 or HAlyCnvDiff> 100 
		if @row>0 
		begin
			select @errmsg= 'Permissible Weight Difference Percentage Should Be Between 0 and 100 ' 
			goto ErrHndlr
		end
		/**Geeta***2.11**/
		/* HCnvDiffEffDt shld be '01/01/80' when HAlyCnvDiff= 0 */ 
		/* 	select @row=count('x') from inserted where Convert(Decimal(16, 2), HAlyCnvDiff)= 0 and 
				HCnvDiffEffDt<> '01/01/1980' and HCnvDiffEffDt <> '01/01/1900'
			if @row>0 
			begin
				select @errmsg= ''Effective Date For Permissible Weight Difference' Should Be '01/01/80' When 'Permissible Wt. Difference % In Cnv' Is Zero        (Tab 5)' 
				goto ErrHndlr
			end
		*/
		/* **** Zubin 211 **** */ 

		/* Zubin 212 */ 
		/* HLotNoYN Should be Y/N */
		select @row=count('x') from inserted where 
			Not Exists (select 'x' from Param where PTyp='YN' and PMCd= HLotNoYN and PSCd= '')
		if @row>0 
		begin
			select @errmsg= '''Implement Lot No For Dia/Cs'' (Enter Y/N)' 
			goto ErrHndlr
		End 
		/* HLotNoYN Cannot be 'N' if SLotNo <> '' in SYyMm */ 
		select @row=count('x') from inserted where HLotNoYN= 'N' and 
			Exists (select 'x' from SYyMm where SCoCd= HCoCd and SLotNo<> '')
		if @row>0 
		begin
			select @errmsg= '''Implement Lot No For Dia/Cs'' Cannot Be Changed To ''N'' When Lot No. Already Exists' 
			goto ErrHndlr
		End 
		/* HFgQuotYN Should be Y/N */
		select @row=count('x') from inserted where 
			Not Exists (select 'x' from Param where PTyp='YN' and PMCd= HFgQuotYN and PSCd= '')
		if @row>0 
		begin
			select @errmsg= '''Show Quotation For FG Bag'' (Enter Y/N)' 
			goto ErrHndlr
		End 
		/* HHGldRtDAYN Should be Y/N */
		select @row=count('x') from inserted where 
			Not Exists (select 'x' from Param where PTyp='YN' and PMCd= HGldRtDAYN and PSCd= '')
		if @row>0 
		begin
			select @errmsg= '''Derived/Actual GldRt In Ord/Inv Prnt'' (Enter ''Y'' For Derived Or ''N'' For Actual)' 
			goto ErrHndlr
		End 
		/* Zubin 212 */ 

		/*** (Jen 2.12 Patch 2) ***/
		select @row= count('x') from inserted where 
			Not Exists (select 'x' from Param where PTyp= 'YN' and PMCd= hLooseInv and PSCd= '') 
		if @row> 0 
		begin 
			select @errmsg= 'Specify Whether to Show Loose RM Invoice Menus? (Enter Y/N)' 
			goto ErrHndlr 
		end 
		/*** (Jen 2.12 Patch 2) ***/

		/*** (Jen 2.13) ***/
		select @row= count('x') from inserted where 
			Not Exists (select 'x' from Param where PTyp= 'YN' and PMCd= hConsolidatedInv and PSCd= '') 
		if @row> 0 
		begin 
			select @errmsg= 'Specify Whether to Show Consolidated Invoice Report Menus? (Enter Y/N)' 
			goto ErrHndlr 
		end 

		select @row= count('x') from inserted where 
			Not Exists (select 'x' from Param where PTyp= 'YN' and PMCd= HAutomateReturns and PSCd= '') 
		if @row> 0 
		begin 
			select @errmsg= 'Specify Whether to Automate Invoice Returns? (Enter Y/N)' 
			goto ErrHndlr 
		end 
		/*** (Jen 2.13) ***/

		/* Zubin 213 */ 
		select @row= count('x') from inserted where HCoCd<> @wSelfCoCd and 
			HMultiCurYn<> IsNull((select HMultiCurYn from Head where HCoCd= @wSelfCoCd 
					and HCd= @wSelfCd) , '')
		if @row> 0 
		begin
			select @errmsg= ' ''Implement Multi Currency Module'' Option For This Company Should Be Same As '+ @wSelfCoCd
			goto ErrHndlr
		end
		/* Zubin 213 (HMultiCurYn= Valid (Y/N)) */ 
		select @row= count('x') from inserted where 
			Not Exists (select 'x' from Param where PTyp= 'YN' and PMCd= HMultiCurYn and PSCd= '') 
		if @row> 0 
		begin 
			select @errmsg= 'Specify Whether to Implement Multi Currency? (Enter Y/N)' 
			goto ErrHndlr 
		end 
		/* Zubin 213 (HMultiCurYn cannot be changed to 'N' if there are more than 1 curr cd in Head) */ 
		select @row= count('x') from inserted ins where 
			Not Exists (Select 'x' from deleted del where ins.HCoCd= del.HCoCd and ins.HCd= del.HCd and 
				ins.HMultiCurYn= del.HMultiCurYn) and HMultiCurYn= 'N' and 
			(select count('x') from (select distinct HBaseCurCd from Head) Hd )> 1 
		if @row> 0 
		begin 
			select @errmsg= 'Cannot Change ''Implement Multi Currency'' To ''N'' As Companies With Different Currencies Exist In Head Table' 
			goto ErrHndlr 
		end 
		/* Zubin 213 (HMultiCurYn cannot be changed to 'N' if recs exist in CustMst) */ 
		select @row= count('x') from inserted ins where 
			Not Exists (Select 'x' from deleted del where ins.HCoCd= del.HCoCd and ins.HCd= del.HCd and 
				ins.HMultiCurYn= del.HMultiCurYn) and HMultiCurYn= 'N' and 
			Exists (select 'x' from CustMst where CmCtg= 'C' and 
				CmCurCd In (select PMCd from Param where PTyp= 'CURNCY' and PSCd= '' and 
					PmCd Not In (select hBaseCurCd from Head where HCd= @wSelfCd)) ) 
		if @row> 0 
		begin 
			select @errmsg= 'Cannot Change ''Implement Multi Currency'' To ''N'' As Customers Exist With Multiple Currencies' 
			goto ErrHndlr 
		end 
		/* Zubin 213 (HMultiCurYn cannot be changed to 'N' if recs exist in OrdMst (For PDC)) */ 
		select @Row= Count('x') from inserted ins where 
			Not Exists (Select 'x' from deleted del where ins.HCoCd= del.HCoCd and ins.HCd= del.HCd and 
				ins.HMultiCurYn= del.HMultiCurYn) and ins.HMultiCurYn= 'N' and 
			Exists (select 'x' from OrdMst where 
				(Select PValue from Param where PTyp= 'TC' and PMCd= OmTc)= 'PL' and 
				OmCmCurCd In (select PMCd from Param where PTyp= 'CURNCY' and PSCd= '' and 
					PmCd Not In (select hBaseCurCd from Head where HCd= @wSelfCd) ) ) 
		if @row> 0 
		begin 
			select @errmsg= 'Cannot Change ''Implement Multi Currency'' To ''N'' As PDC-Master Exists With Multiple Currencies' 
			goto ErrHndlr 
		end 
		/* Zubin 213 */ 

		/* ****** Sachin 2.12 - 05/12/05 - SJM ****** */
		Select @row=count('x') from inserted where 
			Not Exists (select 'x' from Param where PTyp='YN' and PMCd= HSlvModYN and PSCd= '')
		if @row>0 
		begin
			select @errmsg= '''Implement Silver Module'' (Enter Y/N)' 
			goto ErrHndlr
		End 
		/* ****** Sachin 2.12 - 05/12/05 - SJM ****** */
	
		/* Manali 3.03 Palladium */
		Select @row=count('x') from inserted where 
			Not Exists (select 'x' from Param where PTyp='YN' and PMCd= HPdModYN and PSCd= '')
		if @row>0 
		begin
			select @errmsg= '''Implement Palladium Module'' (Enter Y/N)' 
			goto ErrHndlr
		End 
		
		/* ****** Sachin 2.13.0 - 05/06/2006 ****** */
		Select @row=count('x') from inserted where 
			Not Exists (select 'x' from Param where PTyp='YN' and PMCd= HQWModYN and PSCd= '')
		if @row>0 
		begin
			select @errmsg= '''Implement Q/W Module'' (Enter Y/N)' 
			goto ErrHndlr
		End 
	
		/* ****** Sachin 2.14.0 - 26/10/2006 - [Multi-Company] ****** */
		Select @row=count('x') from Inserted where 
			Not Exists (select 'x' from Param where PTyp='YN' and PMCd= HImpComLoc and PSCd= '')
		if @row>0 
		begin
			select @errmsg= '''Implement Common Locations'' (Enter Y/N)' 
			goto ErrHndlr
		End 

		Select @row=count('x') from Inserted where 
			Not Exists (select 'x' from Param where PTyp='YN' and PMCd= hInRepMetLsRmCd and PSCd= '')
		if @row>0 
		begin
			select @errmsg= '''Inv Rep Metal Loss RmCd'' (Enter Y/N)' 
			goto ErrHndlr
		End 
		/* ****** Sachin 2.14.0 - 26/10/2006 - [Multi-Company]  ****** */

		/* ****** Whether to Enable the Model Availability Option ****** */
		Select @row=count('x') from inserted where 
			Not Exists (select 'x' from Param where PTyp='YN' and PMCd= hModAvlYN and PSCd= '')
		if @row>0 
		begin
			select @errmsg= '''Display Model Availability in Design Master/Order Entry'' (Enter Y/N)' 
			goto ErrHndlr
		End 
		
 /* **** Manali 3.10.0 - Multi Price Quotation Module */		
		Select @row=count('x') from inserted where 
			Not Exists (select 'x' from Param where PTyp='YN' and PMCd= HMultiPrcQtMod and PSCd= '')
		if @row>0 
		begin
			select @errmsg= '''Enable Multi Price Quotation Module'' (Enter Y/N)' 
			goto ErrHndlr
		End 

		/* ****** Cannot Enable Q/W When the SysId=CRM in the SysDb Table ****** */
		Select @row=Count('x') From inserted ins where 
			Not Exists (Select 'x' From deleted del Where Ins.hCoCd=del.hCoCd And Ins.hCd=del.hCd And 
				Ins.hQwModYn=del.hQwModYn) And ins.hQwModYn='Y' And
			Exists (Select 'x' From SysDb Where SysId='CRM')
		if @row> 0 
		begin
			select @errmsg= 'Cannot Enable ''Q/W For All RM Categories'' on a CRM Database' 
			goto ErrHndlr
		end
		
		/* ****** Cannot Change Q/W Option When If there is RmCd with RmQw And RmQwCst MisMatch ****** */		
		select @row= count('x') from inserted ins where 
			Not Exists (Select 'x' from deleted del where ins.HCoCd= del.HCoCd and ins.HCd= del.HCd and 
				ins.HQWModYN= del.HQWModYN) and ins.HQWModYN= 'N' and 
			Exists (select 'x' from RmMst where RmQw<> RmQwCst ) 
		if @row> 0 
		begin 
			select @errmsg= 'Cannot Change ''Q/W For All RM Categories'' To ''N'' As Records In RmMst Exist With Different QW Cost And Sale' 
			goto ErrHndlr 
		end 
		/* ****** Sachin 2.13 - 05/06/2006 ****** */
		
		/* *** Jay 2.13(CT) *** */ 
		/* ***** hCompBagsYn = PMCd('YN') ***** */		

		Select @row= count('x') from inserted where 
			not exists (select 'x' from Param where PTyp='YN' and PMCd= hCompBagsYn)
		if @row> 0 
		begin
			select @errmsg= 'Implement Component Bags Module (Y/N)'
			goto ErrHndlr
		end
		
		Select @row= count('x') from inserted where hCompBagsYn = 'N' and 
			exists (select 'x' from Param where PTyp= 'BCHR' and PValue = 'C')
		if @row> 0 
		begin
			select @errmsg= 'Component Bags Module Can Not Be Disabled as Bag Char Exists In Param Table'
			goto ErrHndlr
		end
		/* *** Jay 2.13(CT) *** */ 

		/* *** Jay 2.14Next *** */ 
		Select @row= count('x') from inserted where hCompBagsYn = 'Y' 
		select @row= count('x') from inserted ins where hCompBagsYn = 'Y' and 
			Not Exists (select 'x' from deleted del where ins.hCoCd= del.hCoCd and ins.hCd= del.hCd and 
				ins.hCompBagsYn = del.hCompBagsYn) 
		If @row> 0 
		Begin
			Select @errmsg= 'Component Bags Module Can Not Be Enabled, Please Update From Maintenance.'
			Goto ErrHndlr
		End
		/* *** Jay 2.14Next *** */ 		


		/* *** Jay 2.14(LclInv) *** */ 
		/* ***** hLclInvYn = PMCd('YN') ***** */		
		Select @row= count('x') from inserted where 
			not exists (select 'x' from Param where PTyp='YN' and PMCd= hLclInvYn)
		if @row> 0 
		begin
			select @errmsg= 'Implement Local Invoice Module (Y/N)'
			goto ErrHndlr
		end
		/* *** Jay 2.14(LclInv) *** */ 
		
		/* *** Jay 3.01 *** [FgTag Module] */ 
		Select @row= count('x') from inserted where 
			not exists (select 'x' from Param where PTyp='YN' and PMCd= HFgTagMod)
		if @row> 0 
		begin
			select @errmsg= 'Implement FG Tag Module (Y/N)'
			goto ErrHndlr
		end
		/* *** Jay 3.01 *** [FgTag Module] */

		/*** (Jen 3.01) ***/ 
		Select @row= count('x') from inserted where 
			not exists (select 'x' from Param where PTyp= 'YN' and PMCd= HCustomiseOrd) 
		if @row> 0 
		begin
			select @errmsg= 'Customise Order Entry Screen (Y/N)' 
			goto ErrHndlr 
		end

		Select @row= count('x') from inserted where 
			not exists (select 'x' from Param where PTyp= 'YN' and PMCd= HAutoCustmMtch) 
		if @row> 0 
		begin
			select @errmsg= 'Enable Auto Custom Matching (Y/N)' 
			goto ErrHndlr 
		end

		/* Enforce From Bag Location in Fg Entry - 4.1.2 */
        Select @row= count('x') from inserted where 
			not exists (select 'x' from Param where PTyp= 'YN' and PMCd=HEnfFrBLocInFgYN) 
		if @row> 0 
		begin
			select @errmsg= 'Enforce From Bag Location in Fg Entry (Y/N)' 
			goto ErrHndlr 
		end

		/*** Jay 3.3(FG) ***/ 
		Select @row= count('x') from inserted where 
			not exists (select 'x' from Param where PTyp= 'YN' and PMCd= hFgIOYn) 
		if @row> 0 
		begin
			select @errmsg= 'Enable Finish Goods Transaction (Y/N)' 
			goto ErrHndlr 
		end
		
		/* ***** Manali 3.4.1 - 22/10/08 - HDsgCatYN */
		Select @row=count('x') from inserted where 
			Not Exists (select 'x' from Param where PTyp='YN' and PMCd= HDsgCatYN and PSCd= '')
		if @row>0 
		begin
			select @errmsg= '''Implement Design Catalogue Module'' (Enter Y/N)' 
			goto ErrHndlr
		End 

		/* ***** Manali 3.6.0 - HOrdExcelRepYn */
		Select @row=count('x') from inserted where 
			Not Exists (select 'x' from Param where PTyp='YN' and PMCd= HOrdExcelRepYn and PSCd= '')
		if @row>0 
		begin
			select @errmsg= '''Order Printing in Excel Format'' (Enter Y/N)' 
			goto ErrHndlr
		End 

		/* **** Manali 3.03 - 18/06/08 -  fields for Palladium ('L') included */
		Select @row= count('x') from inserted where 
			(convert (decimal (14, 2), HRtMinusD)<> convert (decimal (14, 2), 0) and convert (decimal (14, 2), HRtPlusD)<> convert (decimal (14, 2), 0) and convert (decimal (14, 2), HRtMinusD)> convert (decimal (14, 2), HRtPlusD)) or 
			(convert (decimal (14, 2), HRtMinusC)<> convert (decimal (14, 2), 0) and convert (decimal (14, 2), HRtPlusC)<> convert (decimal (14, 2), 0) and convert (decimal (14, 2), HRtMinusC)> convert (decimal (14, 2), HRtPlusC)) or 
			(convert (decimal (14, 2), HRtMinusG)<> convert (decimal (14, 2), 0) and convert (decimal (14, 2), HRtPlusG)<> convert (decimal (14, 2), 0) and convert (decimal (14, 2), HRtMinusG)> convert (decimal (14, 2), HRtPlusG)) or 
			(convert (decimal (14, 2), HRtMinusP)<> convert (decimal (14, 2), 0) and convert (decimal (14, 2), HRtPlusP)<> convert (decimal (14, 2), 0) and convert (decimal (14, 2), HRtMinusP)> convert (decimal (14, 2), HRtPlusP)) or 
			(convert (decimal (14, 2), HRtMinusS)<> convert (decimal (14, 2), 0) and convert (decimal (14, 2), HRtPlusS)<> convert (decimal (14, 2), 0) and convert (decimal (14, 2), HRtMinusS)> convert (decimal (14, 2), HRtPlusS)) or 
			(convert (decimal (14, 2), HRtMinusL)<> convert (decimal (14, 2), 0) and convert (decimal (14, 2), HRtPlusL)<> convert (decimal (14, 2), 0) and convert (decimal (14, 2), HRtMinusL)> convert (decimal (14, 2), HRtPlusL)) or 
			(convert (decimal (14, 2), HRtMinusX)<> convert (decimal (14, 2), 0) and convert (decimal (14, 2), HRtPlusX)<> convert (decimal (14, 2), 0) and convert (decimal (14, 2), HRtMinusX)> convert (decimal (14, 2), HRtPlusX)) 
		if @row> 0 
		begin
			select @errmsg= 'From Rate Variation Percentage Cannot Be Greater Than To Rate Variation Percentage' 
			goto ErrHndlr 
		end

		/* **** Manali 3.03 - 18/06/08 -  fields for Palladium ('L') included */
		Select @row= count('x') from inserted where 
			(convert (decimal (14, 2), HPtrMinusD)<> convert (decimal (14, 2), 0) and convert (decimal (14, 2), HPtrPlusD)<> convert (decimal (14, 2), 0) and convert (decimal (14, 2), HPtrMinusD)> convert (decimal (14, 2), HPtrPlusD)) or 
			(convert (decimal (14, 2), HPtrMinusC)<> convert (decimal (14, 2), 0) and convert (decimal (14, 2), HPtrPlusC)<> convert (decimal (14, 2), 0) and convert (decimal (14, 2), HPtrMinusC)> convert (decimal (14, 2), HPtrPlusC)) or 
			(convert (decimal (14, 2), HPtrMinusG)<> convert (decimal (14, 2), 0) and convert (decimal (14, 2), HPtrPlusG)<> convert (decimal (14, 2), 0) and convert (decimal (14, 2), HPtrMinusG)> convert (decimal (14, 2), HPtrPlusG)) or 
			(convert (decimal (14, 2), HPtrMinusP)<> convert (decimal (14, 2), 0) and convert (decimal (14, 2), HPtrPlusP)<> convert (decimal (14, 2), 0) and convert (decimal (14, 2), HPtrMinusP)> convert (decimal (14, 2), HPtrPlusP)) or 
			(convert (decimal (14, 2), HPtrMinusS)<> convert (decimal (14, 2), 0) and convert (decimal (14, 2), HPtrPlusS)<> convert (decimal (14, 2), 0) and convert (decimal (14, 2), HPtrMinusS)> convert (decimal (14, 2), HPtrPlusS)) or 
			(convert (decimal (14, 2), HPtrMinusL)<> convert (decimal (14, 2), 0) and convert (decimal (14, 2), HPtrPlusL)<> convert (decimal (14, 2), 0) and convert (decimal (14, 2), HPtrMinusL)> convert (decimal (14, 2), HPtrPlusL)) or 		
			(convert (decimal (14, 2), HPtrMinusX)<> convert (decimal (14, 2), 0) and convert (decimal (14, 2), HPtrPlusX)<> convert (decimal (14, 2), 0) and convert (decimal (14, 2), HPtrMinusX)> convert (decimal (14, 2), HPtrPlusX)) 
		if @row> 0 
		begin
			select @errmsg= 'From Pc/Ct Variation Percentage Cannot Be Greater Than To Pc/Ct Variation Percentage' 
			goto ErrHndlr 
		end
		/*** (Jen 3.01) ***/ 

		/* ***** Manali 3.03 - Check on bef and aft dates for Auto-Custom matching ***** */
		Select @row= count('x') from Inserted where
			(HImpAftD <> '01/01/80' and HImpBefD <> '01/01/80' and HImpAftD > HImpBefD ) Or
			(HImpAftC <> '01/01/80' and HImpBefC <> '01/01/80' and HImpAftC > HImpBefC ) Or
			(HImpAftG <> '01/01/80' and HImpBefG <> '01/01/80' and HImpAftG > HImpBefG ) Or
			(HImpAftP <> '01/01/80' and HImpBefP <> '01/01/80' and HImpAftP > HImpBefP ) Or
			(HImpAftS <> '01/01/80' and HImpBefS <> '01/01/80' and HImpAftS > HImpBefS ) Or
			(HImpAftL <> '01/01/80' and HImpBefL <> '01/01/80' and HImpAftL > HImpBefL ) Or
			(HImpAftX <> '01/01/80' and HImpBefX <> '01/01/80' and HImpAftX > HImpBefX ) 
		If @row> 0 
		Begin
			Select @errmsg= 'Import After Date(s) Cannot Be Greater Than Import Before Date(s)' 
			Goto ErrHndlr 
		End
		/* ***** Manali 3.03 - Check on bef and aft dates for Auto-Custom matching ***** */


		/*** (Jen 3.1.0_1) ***/ 
		Select @row= count('x') from inserted where 
			not exists (select 'x' from Param where PTyp= 'YN' and PMCd= HInInvGrsWtEqActual) 
		if @row> 0 
		begin 
			select @errmsg= 'Specify Whether To Keep The Invoice Gross Weight Equal To Actual Gross Weight During Copy (Y/N)' 
			goto ErrHndlr 
		end
		/*** (Jen 3.1.0_1) ***/ 
		/*** Jay 3.1.0_1 ***/ 
		Select @row= count('x') from inserted where 
			not exists (select 'x' from Param where PTyp= 'YN' and PMCd= HValidDsgSetCd) 
		if @row> 0 
		begin 
			select @errmsg= 'Design Set Code Validate from Param (Yes/ No) Invalid.' 
			goto ErrHndlr 
		end
		/*** Jay 3.1.0_1 ***/ 

		/*** (Jen 3.1.2) ***/
		Select @row= count('x') from inserted where 
			not exists (select 'x' from Param where PTyp= 'YN' and PMCd= hBestUse) 
		if @row> 0 
		begin
			select @errmsg= 'Enable Best Use In Auto Custom Matching (Y/N)' 
			goto ErrHndlr 
		end
		/*** (Jen 3.1.2) ***/

		/* Sachin 3.11.0 - Invoice rate from Rate Chart */
		select @row=count('x') from inserted where 
			Not Exists (select 'x' from Param where PTyp='YN' and PMCd= hRtfrRtChrtYN and PSCd= '')
		if @row>0 
		begin
			select @errmsg= '''Invoice Rates from Rate Chart'' (Enter Y/N)' 
			goto ErrHndlr
		End 

	End

	if @mode = 'C' 
	Begin
		/****** Key Cannot Be Changed ******/
		select @Row= Count('x') from deleted del where 
			 not exists (Select 'x' from inserted ins where ins.HCoCd= del.HCoCd and ins.HCd= del.HCd) 
		If @Row > 0
		begin 
			Select @ErrMsg= 'Cannot Change Company Code'
			GoTo ErrHndlr
		end
		/****** Cannot Change HName ******/	
		select @row=count('x') from inserted ins where 
			not Exists(Select 'x' from Deleted del where 
				ins.HCoCd= del.HCoCd and ins.HCd= del.HCd and ins.HName= del.HName)
		if @row>0 
		begin
			select @errmsg= 'Cannot Change Company Name'
			goto ErrHndlr
		end
		/****** Cannot change hPtrYN if 'Ptr Y/N' is 'Y' for any RmSCtg ******/	
		select @row=count('x') from inserted ins where 
			not Exists(Select 'x' from Deleted del where 
				ins.HCoCd= del.HCoCd and ins.HCd= del.HCd and ins.hPtrYN= del.hPtrYN)
			and exists (select 'x' from Param where PTyp= 'RMSCTG' and PValue2= 'Y')
		if @row> 0 
		begin
			select @errmsg= 'Cannot Change The ''Ptr Y/N'' Option When ''Ptr Y/N'' option of RMSCTG is ''Y'' '
			goto ErrHndlr
		end

		/****** Cannot Change ExpLocal Option ******/	
		select @row=count('x') from inserted ins where 
			not Exists(Select 'x' from Deleted del where 
				ins.HCoCd= del.HCoCd and ins.HCd= del.HCd and ins.hExpLocal= del.hExpLocal)
		if @row>0 
		begin
			select @errmsg= 'Cannot Change The Local Export Option'
			goto ErrHndlr
		end

		/****** Cannot Change 'Show PPC Module' option if rec exists in PpcCap, OrdPln, OrdPlnDt ******/	
		select @row= count('x') from inserted ins where 
			not Exists(Select 'x' from Deleted del where ins.HCoCd= del.HCoCd 
					and ins.HCd= del.HCd and ins.hPpcYN= del.hPpcYN) 
			and (	Exists(Select 'x' from PpcCap where PcCoCd= HCoCd) 
				Or Exists(Select 'x' from OrdPln where OpCoCd= HCoCd) 
			    )
		if @row> 0 
		begin
			select @errmsg= 'Cannot Change ''Show PPC Module'' Option As Plan Rec Exists'
			goto ErrHndlr
		end

		/*** Check For HShowPrdRmDets ***/
		/*select @row= count('x') from inserted ins where 
			ins.HShowPrdRmDets= 'N' and 
			Exists(Select 'x' from Deleted del where ins.HCoCd= del.HCoCd and 
				ins.HCd= del.HCd and del.HShowPrdRmDets= 'Y') 
			and Exists(Select 'x' from OrdRm where OrCoCd= HCoCd 
				   and (	Str(OrPrdWt, 16, 3)<> Str(OrWt, 16, 3) 
						or OrQty<> OrPrdQty
					) 
				   )
		if @row> 0 
		begin
			select @errmsg= 'Cannot Change 'Show Prd Rm Dets' Option To 'N' As OrdRm Rec Exists Where OrdWt<> PrdWt or OrdQty<> PrdQty'
			goto ErrHndlr
		end*/

		/*** Check For HShowPrdRmDets ***/
		select @row= count('x') from inserted ins where ins.hCoCd= @wSelfCoCd and
			ins.HShowPrdRmDets= 'N' and 
			Exists(Select 'x' from Deleted del where ins.HCoCd= del.HCoCd and 
				ins.HCd= del.HCd and del.HShowPrdRmDets= 'Y') 
			and (Exists(Select 'x' from OrdRm where 
					Str(OrPrdWt, 16, 3)<> Str(OrWt, 16, 3) 
					or OrQty<> OrPrdQty)   
				or Exists(Select 'x' from DsgRm where 
					Str(DrPrdWt, 16, 3)<> Str(DrWt, 16, 3) 
					or DrQty<> DrPrdQty)  
			    )
		if @row> 0 
		begin
			select @errmsg= 'Cannot Change ''Show Prd Rm Dets'' Option To ''N'' As OrdRm Rec Exists Where OrdWt<> PrdWt or OrdQty<> PrdQty'
			goto ErrHndlr
		end

		/*** Check For Bag Acknowledgement ***/
		select @row= count('x') from inserted ins where ins.hTdAck= 'N' and 
			Exists(Select 'x' from Deleted del where ins.HCoCd= del.HCoCd and 
				ins.HCd= del.HCd and del.hTdAck= 'Y')
			and Exists(Select 'x' from Txnd join Param on PTyp = 'TC' and 
				PMCd = TdTc where TdCoCd= HCoCd and PValue = 'BV' and 
				TdAck = 'N')
		if @row>0 
		begin
			select @errmsg= 'Cannot Change ''Acknowledgement Reqd'' Option If Any Bag Movement Is UnAcknowledged'
			goto ErrHndlr
		end

		/****** HCtbMod cannot be changed to 'N' if record exists in CustMst(CmCtg= 'T') ******/	
		select @row= count('x') from inserted where HCtbMod= 'N' 
			and exists (select 'x' from CustMst where CmCtg= 'T') 
		if @row> 0 
		begin
			select @errmsg= 'Cannot Change The ''Show Contractor Module'' Option To ''N'' When Contractor Code Exists'
			goto ErrHndlr
		end
		/* Zubin 213 - HBaseCurCd shld be valid */ 
		select @row= count('x') from inserted ins where 
			Not Exists (select 'x' from Param where PTyp= 'CURNCY' and PMCd= HBaseCurCd and PSCd= '') 
		if @row> 0 
		begin
			select @errmsg= 'Invalid Currency Code' 
			goto ErrHndlr
		end

		/* Zubin 213 - Cannot change Base Cur if rec exists in DtTbl table */ 
		select @row= count('x') from deleted del where 
			Not Exists (select 'x' from inserted ins where ins.hCd= del.hCd and ins.hCoCd= del.hCoCd and 
				ins.HBaseCurCd= del.HBaseCurCd) and 
			Exists (select 'x' from DtTbl where DtTag= 'CNV' and DtDt= '01/01/80' and 
				(DtCurCd1= HBaseCurCd or DtCurCd2= HBaseCurCd) ) 
		if @row> 0 
		begin
			select @errmsg= 'Cannot Change ''Base Currency Code'' As Record(s) Exist In Date Table For Currency Conversion' 
			goto ErrHndlr
		end
		/* Zubin 213 (250606) - Cannot change Base Cur if rec exists in OrdMst table */ 
		select @row= count('x') from deleted del where 
			Not Exists (select 'x' from inserted ins where ins.hCoCd= del.hCoCd and ins.hCd= del.hCd and 
				ins.HBaseCurCd= del.HBaseCurCd) and 
			Exists (select 'x' from OrdMst where OmCoCd= HCoCd) 
		if @row> 0 
		begin
			select @errmsg= 'Cannot Change ''Base Currency Code'' As Record(s) Exist In Order Master Table' 
			goto ErrHndlr
		end
		/* Zubin 213 - Cannot change Base Cur if rec exists in Txn table */ 
		select @row= count('x') from deleted del where 
			Not Exists (select 'x' from inserted ins where ins.hCoCd= del.hCoCd and ins.hCd= del.hCd and 
				ins.HBaseCurCd= del.HBaseCurCd) and 
			Exists (select 'x' from Txn where TCoCd= HCoCd) 
		if @row> 0 
		begin
			select @errmsg= 'Cannot Change ''Base Currency Code'' As Record(s) Exist In Txn Table' 
			goto ErrHndlr
		end
		/* Zubin 213 (250606) */ 
		
		/* ****** Sachin 2.14.0 - 27-10-2006 [Multi-Company] ****** */
		Select @row= Count('x') From Deleted Del Where Del.HImpComLoc='Y' And
			Not Exists (Select 'x' From Inserted Ins Where Ins.hCoCd= Del.hCoCd And Ins.hCd= Del.hCd and 
				Ins.HImpComLoc= del.HImpComLoc) And 
			Exists (Select 'x' From Loc Where LocPrntCd<>'' Or LocPrntYN='Y') 
		If @row> 0 
		Begin
			Select @errmsg= 'Cannot Change ''Implement Common Locations'' As Record(s) Exist With Parent Location Codes in Loc Master Table' 
			GoTo ErrHndlr
		End
		/* ****** Sachin 2.14.0 - 27-10-2006 [Multi-Company] ****** */

		/*** (Jen 3.01) ***/
		Select @row= Count('x') From Deleted Del Where Del.HCustomiseOrd= 'Y' And 
			Not Exists (Select 'x' From Inserted Ins Where Ins.hCoCd= Del.hCoCd And Ins.hCd= Del.hCd and 
				Ins.HCustomiseOrd= del.HCustomiseOrd) And 
			Exists (Select 'x' From UsrRight where not(UrGrp= '')) 
		If @row> 0 
		Begin
			Select @errmsg= 'Cannot Change ''Customise Order Entry Screen'' As User Right Record(s) Exists' 
			GoTo ErrHndlr
		End
		/*** (Jen 3.01) ***/

		Select @row=count('x') from Inserted where 
			Not Exists (select 'x' from Param where PTyp='YN' and PMCd= HjStkNumjFrBagYN and PSCd= '')
		if @row>0 
		begin
			select @errmsg= 'Invalid ''Copy jEmr DmStkNum from Bag'' value. (Enter Y/N)' 
			goto ErrHndlr
		End 


	End

	If @mode = 'D' 
	Begin
		/****** Cannot Delete Any Record ******/	
		Select @row= count('x') from deleted 
		If @row> 0 
		Begin
			Select @errmsg= 'Cannot Delete Any Record From Head file'
			GoTo ErrHndlr
		End
	End

Return
ErrHndlr:
	begin
		select @errmsg='|'+@errmsg
		raiserror(@errmsg,16,1)
		rollback transaction 
	end
END

GOGO

/*************** Remove Existing Index ****************************************************/

If exists (select 'x' from sys.indexes where  name = 'PK_DmYyMm' and upper(object_name(object_id)) ='DMYYMM') 
Begin
	ALTER TABLE [dbo].[DmYyMm] DROP CONSTRAINT [PK_DmYyMm]
End
GOGO

If exists (select 'x' from sys.objects where type = 'UQ' and upper(name) = 'UK_DMYYMM' and upper(object_name(parent_object_id)) ='DMYYMM') 
Begin
	alter table [dbo].[DmYyMm] drop constraint Uk_DmYyMm 
End
GOGO


If exists (select 'x' from sys.indexes where  name = 'UK_DmYyMm' and upper(object_name(object_id)) ='DMYYMM') 
Begin
	drop index Uk_DmYyMm on dbo.DmYyMm 
End
GOGO


/***************************** Remove Columns ***************/
if exists(select * from sys.columns 
            where Name = N'OdDmId' and Object_ID = Object_ID(N'OrdDsg'))
	Alter Table OrdDsg drop column OdDmId 

if exists(select * from sys.columns 
            where Name = N'DDmId' and Object_ID = Object_ID(N'DmYyMm'))
	Alter Table DmYyMM drop column DDmId


/************ Update DmYyMm *********************************************************************/



Delete from DmYyMm

Insert into DmYyMm(DPrtKey,DCoCd,DDmCd,DOdSfx,DDmSz,DDmStkYy,DDmStkChr, DDmStkNo, DLoc,DYyMm,DDmDrQty,DDmCrQty)  
   select distinct OdPrtKey, OdCoCd, OdDmCd, OdSfx, OdDmSz,OdDmStkYy,OdDmStkChr,OdDmStkNo,OmFrDmLoc,convert(int,convert(varchar(6),OmDt,112)),0,0  
   from orddsg join OrdMst on OmPrtKey=OdPrtKey and OmIdNo=OdOmIdNo and OmCoCd=OdCoCd and OmTc=OdTc and OmYy=OdYy and OmChr=OdChr and OmNo=OdNo  
   where not exists(select 'x' from  DmYyMm where DPrtKey=OdPrtKey and DCoCd= OdCoCd and DDmCd=OdDmCd   
        and OdSfx= DOdSfx and OdDmSz= DDmSz and OdDmStkYy= DDmStkYy and OdDmStkChr = DDmStkChr and OdDmStkNo = DDmStkNo and DLoc=OmFrDmLoc and  
        DYyMm= convert(int,convert(varchar(6),OmDt,112)))  
    and OmFrDmLoc <>'' and OmToDmLoc<>''  and
	(select pvalue from param 
		where PTyp= 'TC' and PMCd= OdTc and PSCd= '')  in ('FG','JT')
      
  /* Current Partition */   
  update DmYyMm set  
  DDmCrQty=DDmCrQty+(select isnull(sum(Ins.OdOrdQty),0) from OrdDsg Ins Join OrdMst   
   on OmPrtKey=Ins.OdPrtKey and OmIdNo=Ins.OdOmIdNo and OmCoCd= Ins.OdCoCd and OmTc= Ins.OdTc and OmYy= Ins.OdYy and OmChr= Ins.OdChr and OmNo= Ins.OdNo  
   where OmFrDmDc='C' and Ins.OdPrtKey=DPrtKey and Ins.OdCoCd= DCoCd and Ins.OdDmCd=DDmCd   
   and Ins.OdSfx= DOdSfx and Ins.OdDmSz= DDmSz and Ins.OdDmStkNo = DDmStkNo and Ins.OdDmStkYy = DDmStkYy and Ins.OdDmStkChr = DDmStkChr 
   and OmFrDmLoc=DLoc  	and (select pvalue from param 
		where PTyp= 'TC' and PMCd= OdTc and PSCd= '')  in ('FG','JT')
   and convert(int,convert(varchar(6),OmDt,112))=DYyMm),  
  DDmDrQty=DDmDrQty+(select isnull(sum(Ins.OdOrdQty),0) from OrdDsg Ins Join OrdMst   
   on OmPrtKey=Ins.OdPrtKey and OmIdNo=Ins.OdOmIdNo and OmCoCd= Ins.OdCoCd and OmTc= Ins.OdTc and OmYy= Ins.OdYy and OmChr= Ins.OdChr and OmNo= Ins.OdNo  
   where OmFrDmDc='D'  and Ins.OdPrtKey=DPrtKey and Ins.OdCoCd= DCoCd and Ins.OdDmCd=DDmCd   
   and Ins.OdSfx= DOdSfx and Ins.OdDmSz= DDmSz and Ins.OdDmStkNo= DDmStkNo and Ins.OdDmStkYy = DDmStkYy and Ins.OdDmStkChr = DDmStkChr 
   and OmFrDmLoc=DLoc  
   and convert(int,convert(varchar(6),OmDt,112))=DYyMm)  
  from OrdDsg Inserted Join OrdMst on OmPrtKey=Inserted.OdPrtKey and OmIdNo=Inserted.OdOmIdNo and OmCoCd= Inserted.OdCoCd and OmTc= Inserted.OdTc  
   and OmYy= Inserted.OdYy and OmChr= Inserted.OdChr and OmNo= Inserted.OdNo  
   where Inserted.OdPrtKey=DPrtKey and Inserted.OdCoCd= DCoCd and Inserted.OdDmCd= DDmCd   
   and Inserted.OdSfx= DOdSfx and Inserted.OdDmSz= DDmSz and Inserted.OdDmStkNo= DDmStkNo 
   and Inserted.OdDmStkYy = DDmStkYy and Inserted.OdDmStkChr = DDmStkChr and OmFrDmLoc= DLoc  
   and convert(int,convert(varchar(6),OmDt,112))=DYyMm  
   and (OmFrDmDc='D' or OmFrDmDc='C')  
   and OmFrDmLoc <>'' and OmToDmLoc<>''  
	and (select pvalue from param 
		where PTyp= 'TC' and PMCd= OdTc and PSCd= '')  in ('FG','JT')
  


  /* Current Partition */  
  insert into DmYyMm(DPrtKey,DCoCd,DDmCd,DOdSfx,DDmSz,DDmStkYy,DDmStkChr,DDmStkNo,DLoc,DYyMm,DDmDrQty,DDmCrQty)  
   select distinct OdPrtKey, OdCoCd, OdDmCd, OdSfx, OdDmSz,OdDmStkYy, OdDmStkChr, OdDmStkNo,OmToDmLoc,convert(int,convert(varchar(6),OmDt,112)),0,0  
   from OrdDsg inserted join OrdMst on OmPrtKey=OdPrtKey and OmIdNo=OdOmIdNo and OmCoCd=OdCoCd and OmTc=OdTc and OmYy=OdYy and OmChr=OdChr and OmNo=OdNo  
   where not exists(select 'x' from  DmYyMm where DPrtKey=OdPrtKey and DCoCd= OdCoCd and DDmCd=OdDmCd   
       and OdSfx= DOdSfx and OdDmSz= DDmSz and OdDmStkYy = DDmStkYy and OdDmStkChr = DDmStkChr and OdDmStkNo = DDmStkNo and DLoc=OmToDmLoc and  
       DYyMm= convert(int,convert(varchar(6),OmDt,112)))  
    and OmFrDmLoc <>'' and OmToDmLoc<>''  and
	(select pvalue from param 
		where PTyp= 'TC' and PMCd= OdTc and PSCd= '')  in ('FG','JT')
  
    
  /* Current Partition */  
  update DmYyMm set  
  DDmCrQty=DDmCrQty+(select isnull(sum(Ins.OdOrdQty),0) from OrdDsg Ins Join OrdMst   
   on OmPrtKey=Ins.OdPrtKey and OmIdNo=Ins.OdOmIdNo and OmCoCd= Ins.OdCoCd and OmTc= Ins.OdTc and OmYy= Ins.OdYy and OmChr= Ins.OdChr and OmNo= Ins.OdNo  
   where OmToDmDc='C' and Ins.OdPrtKey=DPrtKey and Ins.OdCoCd= DCoCd and Ins.OdDmCd=DDmCd   
   and Ins.OdSfx= DOdSfx and Ins.OdDmSz= DDmSz and Ins.OdDmStkNo= DDmStkNo and Ins.OdDmStkYy = DDmStkYy and Ins.OdDmStkChr = DDmStkChr and OmToDmLoc=DLoc  
   and convert(int,convert(varchar(6),OmDt,112))=DYyMm and (select pvalue from param 
		where PTyp= 'TC' and PMCd= OdTc and PSCd= '')  in ('FG','JT')) ,  
  DDmDrQty=DDmDrQty+(select isnull(sum(Ins.OdOrdQty),0) from OrdDsg Ins Join OrdMst   
   on OmPrtKey=Ins.OdPrtKey and OmIdNo=Ins.OdOmIdNo and OmCoCd= Ins.OdCoCd and OmTc= Ins.OdTc and OmYy= Ins.OdYy and OmChr= Ins.OdChr and OmNo= Ins.OdNo  
   where OmToDmDc='D' and Ins.OdPrtKey=DPrtKey and Ins.OdCoCd= DCoCd and Ins.OdDmCd=DDmCd   
   and Ins.OdSfx= DOdSfx and Ins.OdDmSz= DDmSz and Ins.OdDmStkNo= DDmStkNo and Ins.OdDmStkChr = DDmStkChr and Ins.OdDmStkChr = DDmStkChr and OmToDmLoc=DLoc  
   and convert(int,convert(varchar(6),OmDt,112))=DYyMm)  
  from OrdDsg Inserted Join OrdMst on OmPrtKey=Inserted.OdPrtKey and OmIdNo=Inserted.OdOmIdNo and OmCoCd= Inserted.OdCoCd and OmTc= Inserted.OdTc  
   and OmYy= Inserted.OdYy and OmChr= Inserted.OdChr and OmNo= Inserted.OdNo  
   where Inserted.OdPrtKey=DPrtKey and Inserted.OdCoCd= DCoCd and Inserted.OdDmCd=DDmCd   
   and Inserted.OdSfx= DOdSfx and Inserted.OdDmSz= DDmSz and Inserted.OdDmStkNo= DDmStkNo and Inserted.OdDmStkChr = DDmStkChr and Inserted.OdDmStkYy = DDmStkYy 
   and OmToDmLoc= DLoc  
   and convert(int,convert(varchar(6),OmDt,112))=DYyMm  
   and (OmToDmDc='D' or OmToDmDc='C')  
   and OmFrDmLoc <>'' and OmToDmLoc<>''  and
	(select pvalue from param 
		where PTyp= 'TC' and PMCd= OdTc and PSCd= '')  in ('FG','JT')

/***********************************************************************/
GOGO
/****** Object:  Index [PK_Bag]    Script Date: 22/04/15 1:33:45 AM ******/
ALTER TABLE [dbo].[DmYyMm] ADD  CONSTRAINT [PK_DmYyMm] PRIMARY KEY CLUSTERED 
(
	[DPrtKey] ASC,
	[DCoCd] ASC,
	[DYyMm] ASC,
	[DLoc] ASC,
	[DDmCd] ASC,
	[DOdSfx] ASC,
	[DDmSz] ASC,
	[DDmStkYy] ASC,
	[DDmStkChr] ASC,
	[DDmStkNo] ASC
)
GOGO
CREATE unique NONCLUSTERED INDEX [UK_DmYyMm] ON [dbo].[DmYyMm]
(
	[DCoCd] ASC,
	[DYyMm] ASC,
	[DLoc] ASC,
	[DDmCd] ASC,
	[DOdSfx] ASC,
	[DDmSz] ASC,
	[DDmStkYy] ASC,
	[DDmStkChr] ASC,
	[DDmStkNo] ASC
)
GOGO

/**************************Tc Str ************************************************************************************/
begin try
insert into TcStr(TsTbl, TsTyp, TsSTyp,TsIdName, TsStyle, TsLbl, TsDataTyp,TsMask,TsMaxLen,TsBlank)
values ('Ord','DSG','JSA','OdDmStkChr','N','','U','',         2,'Y')
insert into TcStr(TsTbl, TsTyp, TsSTyp,TsIdName, TsStyle, TsLbl, TsDataTyp,TsMask,TsMaxLen,TsBlank)
values ('Ord','DSG','JSA','OdDmStkYy','N','','U','',         2,'Y')
insert into TcStr(TsTbl, TsTyp, TsSTyp,TsIdName, TsStyle, TsLbl, TsDataTyp,TsMask,TsMaxLen,TsBlank)
values ('Ord','JCI','DSG','OdDmStkChr','N','','U','',         2,'Y')
insert into TcStr(TsTbl, TsTyp, TsSTyp,TsIdName, TsStyle, TsLbl, TsDataTyp,TsMask,TsMaxLen,TsBlank)
values ('Ord','JCI','DSG','OdDmStkNo','N','','N','###0',         4,'Y')
insert into TcStr(TsTbl, TsTyp, TsSTyp,TsIdName, TsStyle, TsLbl, TsDataTyp,TsMask,TsMaxLen,TsBlank)
values ('Ord','JCI','DSG','OdDmStkYy','N','','U','',         2,'Y')
insert into TcStr(TsTbl, TsTyp, TsSTyp,TsIdName, TsStyle, TsLbl, TsDataTyp,TsMask,TsMaxLen,TsBlank)
values ('Ord','JCN','DSG','OdDmStkChr','N','','U','',         2,'Y')
insert into TcStr(TsTbl, TsTyp, TsSTyp,TsIdName, TsStyle, TsLbl, TsDataTyp,TsMask,TsMaxLen,TsBlank)
values ('Ord','JCN','DSG','OdDmStkNo','N','','N','###0',         4,'Y')
insert into TcStr(TsTbl, TsTyp, TsSTyp,TsIdName, TsStyle, TsLbl, TsDataTyp,TsMask,TsMaxLen,TsBlank)
values ('Ord','JCN','DSG','OdDmStkYy','N','','U','',         2,'Y')
insert into TcStr(TsTbl, TsTyp, TsSTyp,TsIdName, TsStyle, TsLbl, TsDataTyp,TsMask,TsMaxLen,TsBlank)
values ('Ord','JCP','DSG','OdDmStkChr','N','','U','',         2,'Y')
insert into TcStr(TsTbl, TsTyp, TsSTyp,TsIdName, TsStyle, TsLbl, TsDataTyp,TsMask,TsMaxLen,TsBlank)
values ('Ord','JCP','DSG','OdDmStkNo','N','','N','###0',         4,'Y')
insert into TcStr(TsTbl, TsTyp, TsSTyp,TsIdName, TsStyle, TsLbl, TsDataTyp,TsMask,TsMaxLen,TsBlank)
values ('Ord','JCP','DSG','OdDmStkYy','N','','U','',         2,'Y')
insert into TcStr(TsTbl, TsTyp, TsSTyp,TsIdName, TsStyle, TsLbl, TsDataTyp,TsMask,TsMaxLen,TsBlank)
values ('Ord','JCR','DSG','OdDmStkChr','N','','U','',         2,'Y')
insert into TcStr(TsTbl, TsTyp, TsSTyp,TsIdName, TsStyle, TsLbl, TsDataTyp,TsMask,TsMaxLen,TsBlank)
values ('Ord','JCR','DSG','OdDmStkNo','N','','N','###0',         4,'Y')
insert into TcStr(TsTbl, TsTyp, TsSTyp,TsIdName, TsStyle, TsLbl, TsDataTyp,TsMask,TsMaxLen,TsBlank)
values ('Ord','JCR','DSG','OdDmStkYy','N','','U','',         2,'Y')
insert into TcStr(TsTbl, TsTyp, TsSTyp,TsIdName, TsStyle, TsLbl, TsDataTyp,TsMask,TsMaxLen,TsBlank)
values ('Ord','JIN','DSG','OdDmStkChr','N','','U','',         2,'Y')
insert into TcStr(TsTbl, TsTyp, TsSTyp,TsIdName, TsStyle, TsLbl, TsDataTyp,TsMask,TsMaxLen,TsBlank)
values ('Ord','JIN','DSG','OdDmStkNo','N','','N','###0',         4,'Y')
insert into TcStr(TsTbl, TsTyp, TsSTyp,TsIdName, TsStyle, TsLbl, TsDataTyp,TsMask,TsMaxLen,TsBlank)
values ('Ord','JIN','DSG','OdDmStkYy','N','','U','',         2,'Y')
insert into TcStr(TsTbl, TsTyp, TsSTyp,TsIdName, TsStyle, TsLbl, TsDataTyp,TsMask,TsMaxLen,TsBlank)
values ('Ord','JMI','DSG','OdDmStkChr','N','','U','',         2,'Y')
insert into TcStr(TsTbl, TsTyp, TsSTyp,TsIdName, TsStyle, TsLbl, TsDataTyp,TsMask,TsMaxLen,TsBlank)
values ('Ord','JMI','DSG','OdDmStkNo','N','','N','###0',         4,'Y')
insert into TcStr(TsTbl, TsTyp, TsSTyp,TsIdName, TsStyle, TsLbl, TsDataTyp,TsMask,TsMaxLen,TsBlank)
values ('Ord','JMI','DSG','OdDmStkYy','N','','U','',         2,'Y')
insert into TcStr(TsTbl, TsTyp, TsSTyp,TsIdName, TsStyle, TsLbl, TsDataTyp,TsMask,TsMaxLen,TsBlank)
values ('Ord','JMR','DSG','OdDmStkChr','N','','U','',         2,'Y')
insert into TcStr(TsTbl, TsTyp, TsSTyp,TsIdName, TsStyle, TsLbl, TsDataTyp,TsMask,TsMaxLen,TsBlank)
values ('Ord','JMR','DSG','OdDmStkNo','N','','N','###0',         4,'Y')
insert into TcStr(TsTbl, TsTyp, TsSTyp,TsIdName, TsStyle, TsLbl, TsDataTyp,TsMask,TsMaxLen,TsBlank)
values ('Ord','JMR','DSG','OdDmStkYy','N','','U','',         2,'Y')
insert into TcStr(TsTbl, TsTyp, TsSTyp,TsIdName, TsStyle, TsLbl, TsDataTyp,TsMask,TsMaxLen,TsBlank)
values ('Ord','JMS','DSG','OdDmStkChr','N','','U','',         2,'Y')
insert into TcStr(TsTbl, TsTyp, TsSTyp,TsIdName, TsStyle, TsLbl, TsDataTyp,TsMask,TsMaxLen,TsBlank)
values ('Ord','JMS','DSG','OdDmStkNo','N','','N','###0',         4,'Y')
insert into TcStr(TsTbl, TsTyp, TsSTyp,TsIdName, TsStyle, TsLbl, TsDataTyp,TsMask,TsMaxLen,TsBlank)
values ('Ord','JMS','DSG','OdDmStkYy','N','','U','',         2,'Y')
insert into TcStr(TsTbl, TsTyp, TsSTyp,TsIdName, TsStyle, TsLbl, TsDataTyp,TsMask,TsMaxLen,TsBlank)
values ('Ord','JMT','DSG','OdDmStkChr','N','','U','',         2,'Y')
insert into TcStr(TsTbl, TsTyp, TsSTyp,TsIdName, TsStyle, TsLbl, TsDataTyp,TsMask,TsMaxLen,TsBlank)
values ('Ord','JMT','DSG','OdDmStkNo','N','','N','###0',         4,'Y')
insert into TcStr(TsTbl, TsTyp, TsSTyp,TsIdName, TsStyle, TsLbl, TsDataTyp,TsMask,TsMaxLen,TsBlank)
values ('Ord','JMT','DSG','OdDmStkYy','N','','U','',         2,'Y')
insert into TcStr(TsTbl, TsTyp, TsSTyp,TsIdName, TsStyle, TsLbl, TsDataTyp,TsMask,TsMaxLen,TsBlank)
values ('Ord','JMV','DSG','OdDmStkChr','N','','U','',         2,'Y')
insert into TcStr(TsTbl, TsTyp, TsSTyp,TsIdName, TsStyle, TsLbl, TsDataTyp,TsMask,TsMaxLen,TsBlank)
values ('Ord','JMV','DSG','OdDmStkNo','N','','N','###0',         4,'Y')
insert into TcStr(TsTbl, TsTyp, TsSTyp,TsIdName, TsStyle, TsLbl, TsDataTyp,TsMask,TsMaxLen,TsBlank)
values ('Ord','JMV','DSG','OdDmStkYy','N','','U','',         2,'Y')
insert into TcStr(TsTbl, TsTyp, TsSTyp,TsIdName, TsStyle, TsLbl, TsDataTyp,TsMask,TsMaxLen,TsBlank)
values ('Ord','JOT','DSG','OdDmStkChr','N','','U','',         2,'Y')
insert into TcStr(TsTbl, TsTyp, TsSTyp,TsIdName, TsStyle, TsLbl, TsDataTyp,TsMask,TsMaxLen,TsBlank)
values ('Ord','JOT','DSG','OdDmStkNo','N','','N','###0',         4,'Y')
insert into TcStr(TsTbl, TsTyp, TsSTyp,TsIdName, TsStyle, TsLbl, TsDataTyp,TsMask,TsMaxLen,TsBlank)
values ('Ord','JOT','DSG','OdDmStkYy','N','','U','',         2,'Y')
insert into TcStr(TsTbl, TsTyp, TsSTyp,TsIdName, TsStyle, TsLbl, TsDataTyp,TsMask,TsMaxLen,TsBlank)
values ('Ord','JPO','DSG','OdDmStkChr','I','','U','',         2,'Y')
insert into TcStr(TsTbl, TsTyp, TsSTyp,TsIdName, TsStyle, TsLbl, TsDataTyp,TsMask,TsMaxLen,TsBlank)
values ('Ord','JPO','DSG','OdDmStkNo','I','','N','###0',         4,'Y')
insert into TcStr(TsTbl, TsTyp, TsSTyp,TsIdName, TsStyle, TsLbl, TsDataTyp,TsMask,TsMaxLen,TsBlank)
values ('Ord','JPO','DSG','OdDmStkYy','I','','U','',         2,'Y')
insert into TcStr(TsTbl, TsTyp, TsSTyp,TsIdName, TsStyle, TsLbl, TsDataTyp,TsMask,TsMaxLen,TsBlank)
values ('Ord','JPR','DSG','OdDmStkChr','N','','U','',         2,'Y')
insert into TcStr(TsTbl, TsTyp, TsSTyp,TsIdName, TsStyle, TsLbl, TsDataTyp,TsMask,TsMaxLen,TsBlank)
values ('Ord','JPR','DSG','OdDmStkNo','N','','N','###0',         4,'Y')
insert into TcStr(TsTbl, TsTyp, TsSTyp,TsIdName, TsStyle, TsLbl, TsDataTyp,TsMask,TsMaxLen,TsBlank)
values ('Ord','JPR','DSG','OdDmStkYy','N','','U','',         2,'Y')
insert into TcStr(TsTbl, TsTyp, TsSTyp,TsIdName, TsStyle, TsLbl, TsDataTyp,TsMask,TsMaxLen,TsBlank)
values ('Ord','JPT','DSG','OdDmStkChr','N','','U','',         2,'Y')
insert into TcStr(TsTbl, TsTyp, TsSTyp,TsIdName, TsStyle, TsLbl, TsDataTyp,TsMask,TsMaxLen,TsBlank)
values ('Ord','JPT','DSG','OdDmStkNo','N','','N','###0',         4,'Y')
insert into TcStr(TsTbl, TsTyp, TsSTyp,TsIdName, TsStyle, TsLbl, TsDataTyp,TsMask,TsMaxLen,TsBlank)
values ('Ord','JPT','DSG','OdDmStkYy','N','','U','',         2,'Y')
insert into TcStr(TsTbl, TsTyp, TsSTyp,TsIdName, TsStyle, TsLbl, TsDataTyp,TsMask,TsMaxLen,TsBlank)
values ('Ord','JSA','DSG','OdDmStkChr','N','','U','',         2,'Y')
insert into TcStr(TsTbl, TsTyp, TsSTyp,TsIdName, TsStyle, TsLbl, TsDataTyp,TsMask,TsMaxLen,TsBlank)
values ('Ord','JSA','DSG','OdDmStkNo','N','','N','###0',         4,'Y')
insert into TcStr(TsTbl, TsTyp, TsSTyp,TsIdName, TsStyle, TsLbl, TsDataTyp,TsMask,TsMaxLen,TsBlank)
values ('Ord','JSA','DSG','OdDmStkYy','N','','U','',         2,'Y')
insert into TcStr(TsTbl, TsTyp, TsSTyp,TsIdName, TsStyle, TsLbl, TsDataTyp,TsMask,TsMaxLen,TsBlank)
values ('Ord','JSO','DSG','OdDmStkChr','I','','U','',         2,'Y')
insert into TcStr(TsTbl, TsTyp, TsSTyp,TsIdName, TsStyle, TsLbl, TsDataTyp,TsMask,TsMaxLen,TsBlank)
values ('Ord','JSO','DSG','OdDmStkNo','I','','N','###0',         4,'Y')
insert into TcStr(TsTbl, TsTyp, TsSTyp,TsIdName, TsStyle, TsLbl, TsDataTyp,TsMask,TsMaxLen,TsBlank)
values ('Ord','JSO','DSG','OdDmStkYy','I','','U','',         2,'Y')
insert into TcStr(TsTbl, TsTyp, TsSTyp,TsIdName, TsStyle, TsLbl, TsDataTyp,TsMask,TsMaxLen,TsBlank)
values ('Ord','JST','DSG','OdDmStkChr','N','','U','',         2,'Y')
insert into TcStr(TsTbl, TsTyp, TsSTyp,TsIdName, TsStyle, TsLbl, TsDataTyp,TsMask,TsMaxLen,TsBlank)
values ('Ord','JST','DSG','OdDmStkNo','N','','N','###0',         4,'Y')
insert into TcStr(TsTbl, TsTyp, TsSTyp,TsIdName, TsStyle, TsLbl, TsDataTyp,TsMask,TsMaxLen,TsBlank)
values ('Ord','JST','DSG','OdDmStkYy','N','','U','',         2,'Y')
insert into TcStr(TsTbl, TsTyp, TsSTyp,TsIdName, TsStyle, TsLbl, TsDataTyp,TsMask,TsMaxLen,TsBlank)
values ('Ord','TIN','DSG','OdDmStkChr','N','','U','',         2,'Y')
insert into TcStr(TsTbl, TsTyp, TsSTyp,TsIdName, TsStyle, TsLbl, TsDataTyp,TsMask,TsMaxLen,TsBlank)
values ('Ord','TIN','DSG','OdDmStkNo','N','','N','###0',         4,'Y')
insert into TcStr(TsTbl, TsTyp, TsSTyp,TsIdName, TsStyle, TsLbl, TsDataTyp,TsMask,TsMaxLen,TsBlank)
values ('Ord','TIN','DSG','OdDmStkYy','N','','U','',         2,'Y')
insert into TcStr(TsTbl, TsTyp, TsSTyp,TsIdName, TsStyle, TsLbl, TsDataTyp,TsMask,TsMaxLen,TsBlank)
values ('Ord','TOT','DSG','OdDmStkChr','N','','U','',         2,'Y')
insert into TcStr(TsTbl, TsTyp, TsSTyp,TsIdName, TsStyle, TsLbl, TsDataTyp,TsMask,TsMaxLen,TsBlank)
values ('Ord','TOT','DSG','OdDmStkNo','N','','N','###0',         4,'Y')
insert into TcStr(TsTbl, TsTyp, TsSTyp,TsIdName, TsStyle, TsLbl, TsDataTyp,TsMask,TsMaxLen,TsBlank)
values ('Ord','TOT','DSG','OdDmStkYy','N','','U','',         2,'Y')
end try
begin catch
end catch
GOGO
Delete from TcStr where TsIdName= 'OdDmId'
Delete from TcStr where TsIdName= 'DmId'
GOGO

insert into trgdisa values ('param');
update param set pdesc225 = '5.0.1' where ptyp = 'VER' and pmcd = 'DATVER';
delete from trgdisa;
GOGO



