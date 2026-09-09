SET TRANSACTION ISOLATION LEVEL READ COMMITTED
GOGO

ALTER TRIGGER [dbo].[Bag_INUPDL] ON [dbo].[Bag] 
		FOR INSERT,UPDATE,DELETE NOT FOR REPLICATION
	AS
	BEGIN		
	DECLARE @ROW INT
	DECLARE @mode char(1)
	DECLARE @errmsg varchar(255)
	DECLARE @ReplAgt BIT
	DECLARE @wCurrPrtn VarChar(1)
	DECLARE @wSysDbUni varchar(1)

	IF @@ROWCOUNT=0
		RETURN

	If Exists(Select 'x' From TrgDisa Where TdTblNm = 'Bag' )
		RETURN

	select @mode=case (select count('x') from inserted) when 0 then 'D'
		else
			case (select count('x') from deleted) when 0 then 'A' else 'C' end
		end 

	Select @wCurrPrtn='C'

	Select @wSysDbUni = 'N'
      
	Set @ReplAgt=Convert(BIT,SessionProperty('Replication_Agent'))
	
	 if @mode ='C' 
		 Begin  
		/****** BLstYy and BLstKey should contain the TdYy and TdKey values of the Last transaction of the bag (index= xTdBag) ******/
				/* Current Partition */
				select @row= count('x') from inserted ins where BPrtKey=@wCurrPrtn and BLstYy<>'' and 
					(Select del.BLstYyKey from deleted del where del.BIdNo=ins.BIdNo and del.BCoCd= ins.BCoCd and
						del.BYy= ins.BYy and del.BChr= ins.BChr and del.BNo= ins.BNo and del.BPrtKey=ins.BPRtKey)
						<> ins.BLstYyKey and 
					IsNull((Select max(TdYyKey) from Txnd where TdBIdNo=BIdNo and TdCoCd=BCoCd and TdBYy= BYy 
					and TdBChr= BChr and TdBNo= BNo and TdPrtKey=BPrtKey and TdSrNo= 0), '')<> BLstYyKey 
				if @row> 0 
				begin
					select @errmsg= 'This Is Not The Last YyKey For This Bag'
					goto ErrHndlr
				end
		
	/******  Cannot change BCoCd, BYy, BChr, BNo, BOdTc, BOdYy, BOdChr, BOdNo, BOdSr, BOdDmCd, BOdSfx, BOdDmSz  ******/
		/* Current Partition */
		select @row= count('x') from deleted del where del.BPrtKey=@wCurrPrtn and
			not exists(Select 'x' from inserted ins where ins.BPrtKey=del.BPrtKey and ins.BIdNo=del.BIdNo 
			and ins.BCoCd= del.BCoCd and ins.BYy  = del.BYy and ins.BChr = del.BChr and ins.BNo = del.BNo 
			and ins.BOdIdNo=del.BOdIdNo and ins.BOdTc= del.BOdTc and ins.BOdYy= del.BOdYy and ins.BOdChr= del.BOdChr 
			and ins.BOdNo= del.BOdNo and ins.BOdSr= del.BOdSr and ins.BOdDmCd= del.BOdDmCd 
			and ins.BOdSfx = del.BOdSfx and ins.BOdDmSz = del.BOdDmSz 
			)
		if @row> 0 
		begin
			select @errmsg= 'Contact 9820490905 Cannot Change Company Code/ Bag No/ Order/ Design, Sfx And Size'
			goto ErrHndlr
		end
		End
	
	Select @row=count('x') from inserted where 
			not exists (select 'x' from Loc where 
						LocCoCd= BCoCd and LocTyp= 'P' and LocCd= BLoc)
	if @row>0 
	Begin
		select @errmsg= 'Bag Location Not Defined'
		goto ErrHndlr
	end

	if @mode in ('A')
	Begin
		select @row=count('x') from inserted where (BLstYy<> '' or BLstKey<> 0)
		if @row>0 
		Begin
			select @errmsg= 'Last Yy/ Key Should Be Blank'
			goto ErrHndlr
		end

		select @row=count('x') from inserted where 
			Exists (Select 'x' from Param where PTyp= 'BCHR' and PMCd= BChr and PSCd= '' and PValue= 'F') 
			And 
			Exists (Select 'x' from Head where HCoCd= 'ZZZ' and HCd= 'ZSELF' and hFluteBagYN= 'N') 
		if @row> 0 
		Begin 
			select @errmsg= 'Cannot Add A Flute Bag Chr As The ''Allow Flute Bags'' = ''N'' In Head Table' 
			goto ErrHndlr 
		end 

		select @row=count('x') from inserted where 
			Exists (Select 'x' from Param where PTyp= 'BCHR' and PMCd= BChr and PSCd= '' and PValue= 'C') 
			And 
			Exists (Select 'x' from Head where HCoCd= 'ZZZ' and HCd= 'ZSELF' and hCompBagsYn= 'N') 
		if @row> 0 
		Begin 
			select @errmsg= 'Cannot Add A Component Bag Chr As The ''Allow Component Bags'' = ''N'' In Head Table' 
			goto ErrHndlr 
		end 

		select @row=count('x') from inserted where 
			not Exists (Select 'x' from OrdCT where OctPrtKey=BPrtKey and OctOdIdNo=BOdIdNo and OctCoCd= BCoCd 
						and OctTc= BOdTc and OctYy= BOdYy and OctChr= BOdChr and OctNo= BOdNo and OctSr= BOdSr and OctCtChr= BChr)
			and Exists (Select 'x' from Param where PTyp= 'BCHR' and PMCd= BChr and PSCd= '' and PValue= 'C') 
		if @row> 0 
		Begin 
			select @errmsg= 'Cannot Open This Component Bag, Component Is Not Present In Order'
			goto ErrHndlr 
		end 

			Select @Row=Count('x') from Inserted 
					Join OrdMst On OmPrtKey=BPrtKey And OmIdNo=BOmIdNo And OmCoCd=BCoCd And OmTc=BOdTc And OmYy=BOdYy
									And OmChr=BOdChr And OmNo=BOdNo
					Where BPrtKey=@wCurrPrtn and BMaxTxnDt<OmDt
			If @Row > 0
			Begin
				Select @ErrMsg= 'Bag Opening Date Cannot Be Less Than Order Date'
				GoTo ErrHndlr
			end			

	End

	if @mode in ('D')
	Begin
		select @row=count('x') from deleted del where BLstYyKey<> ''
		if @row>0 
		Begin
			select @errmsg= 'Cannot Delete This Record as Transaction Already Exists'
			goto ErrHndlr
		end

		Select @row= count('x') from deleted del where 
			exists (select 'x' from TreeDet where TrdCoCd= BCoCd and TrdBYy= BYy and TrdBChr= BChr and TrdBNo= BNo)
		if @row>0 
		Begin
			select @errmsg= 'Cannot Delete This Record as Tree Detail Rec Exists'
			goto ErrHndlr
		end

		select @row= count('x') from deleted del where 
			exists (select 'x' from OrdDsg where OdTc='QB' and OdCoCd= BCoCd and 
					OdBYy= BYy and OdBChr= BChr and OdBNo= BNo)
		if @row> 0 
		Begin
			select @errmsg= 'Cannot Delete This Record as FG Bag Quotation Rec Exists'
			goto ErrHndlr
		end
		/******5717 checks missed are taken from 577 and added to 5716 trigger -- start *****/		
		select @row= count('x') from deleted del where 
			exists (select 'x' from Fgd where FdPrtKey=BPrtKey and FdBIdNo=BIdNo and FdCoCd= BCoCd and 
					FdBYy= BYy and FdBChr= BChr and FdBNo= BNo)
		if @row>0 
		Begin
			select @errmsg= 'Cannot Delete This Record as Fin Goods Rec Exists'
			goto ErrHndlr
		end
		select @row= count('x') from deleted del where 
			exists (select 'x' from RfIdMst where RimBCoCd= BCoCd and 
					RimBYy= BYy and RimBChr= BChr and RimBNo= BNo)
		if @row> 0 
		Begin
			select @errmsg= 'Cannot Delete This Record as Record Of This Bag Exists in RfIdMst'
			goto ErrHndlr
		end		
		select @row= count('x') from deleted del where 
			exists (select 'x' from CertDet where CdBCoCd = BCoCd and CdBYy = BYy and CdBChr = BChr and CdBNo = BNo)
		if @row>0 
		Begin
			select @errmsg= 'Cannot Delete This Record as Certificate Detail Rec Exists'
			goto ErrHndlr
		end
		/******5717 checks missed are taken from 577 and added to 5716 trigger -- end *****/
	End

	/******5717 checks missed are taken from 577 and added to 5716 trigger -- start *****/
	select @row=count('x') from inserted where 
			not exists (select 'x' from Param where PTyp= 'TC' and PMCd= BOdTc and PValue In('SO','OS'))
	if @row>0 
	Begin
		select @errmsg= 'Sales Order Tc Not Defined'
		goto ErrHndlr
	end	
	Select @row=count('x') from inserted where BOdlSrNO <> 0 and
			not exists (select 'x' from OrdLn where OdlCoCd= BCoCd and 
		OdlTc= BOdTc and OdlYy= BOdYy and OdlChr= BOdChr and OdlNo= BOdNo and OdlSr= BOdSr and OdlSrNo = BOdlSrNo)
	if @row>0 
	Begin
		select @errmsg='Given Order Line does Not Exist'
		goto ErrHndlr
	end	
	select @row=count('x') from inserted where BRepYN='Y' and
			not exists (select 'x' from OrdDsg where OdPRtKey=BPrtKey and OdCoCd= BCoCd and 
		OdTc= BRepOdTc and OdYy= BRepOdYy and OdChr= BRepOdChr and OdNo= BRepOdNo and OdSr= BRepOdSr)
	if @row>0 
	Begin
		select @errmsg='Repair Order Does Not Exist'
		goto ErrHndlr
	end
	select @row=count('x') from inserted where not exists (select 'x' from OrdDsg where 
		OdPrtKEy=BPrtKey and OdIdNo=BOdIdNo and OdCoCd= BCoCd and OdTc= BOdTc and OdYy= BOdYy and OdChr= BOdChr and OdNo= BOdNo 
		and OdSr= BOdSr and OdDmCd= BOdDmCd and OdSfx= BOdSfx and OdDmSz= BOdDmSz )
	if @row>0 
	Begin
		select @errmsg='Design/ Suffix/ Size Does Not Exist For This Order Sr.'
		goto ErrHndlr
	end
	select @row=count('x') from inserted where BOpnQty< 0 
	if @row>0 
	Begin
		select @errmsg= 'Bag Opening Qty Should Be Greater Than Or Equal To Zero'
		goto ErrHndlr
	end
	select @row=count('x') from inserted where BOpnPtQty< 0 	
	if @row> 0 
	Begin
		select @errmsg= 'Opening Part Qty Should Be Greater Than Or Equal To Zero For Any Bag'		
		goto ErrHndlr
	end
	select @row=count('x') from inserted where BOpnDt in ('01/01/1980', '01/01/1900')  
	if @row>0 
	Begin
		select @errmsg= 'Bag Opening Date Should Be Entered'
		goto ErrHndlr
	end
	select @row= count('x') from inserted where BRecvDt in ('01/01/1980', '01/01/1900')  
	if @row> 0        
	Begin            
		select @errmsg= 'Bag Receiving Date Should Be Entered'
		goto ErrHndlr
	end  
	select @row=count('x') from inserted where 
				not exists (select 'x' from Loc where 
						LocCoCd= BCoCd and LocTyp= 'P' and LocCd= BOpnLoc and (@mode <> 'A' or LocValidYN='Y') and LocCd<> 'PFG') 
	if @row>0 
	Begin
		select @errmsg= 'Bag Opening Location Not Defined'
		goto ErrHndlr
	end
	select @row=count('x') from inserted where BRepLoc <>'' and
			not exists (select 'x' from Loc where 
						LocCoCd= BCoCd and LocTyp= 'P' and LocCd= BRepLoc) 
	if @row>0 
	Begin
		select @errmsg= 'Bag Repair Location Not Defined'
		goto ErrHndlr
	end
	select @row= count('x') from inserted 
			where BLstYyKey= '' and BLoc<> BOpnLoc 
	if @row> 0 
	Begin
		select @errmsg= 'Bag Location Should Be Equal To Bag Opening Location'
		goto ErrHndlr
	end
	select @row=count('x') from inserted where 
				not exists (select 'x' from Param where PTyp= 'YN' and PMCd= BCls)
	if @row>0 
	Begin
		select @errmsg= 'Bag Closed (Enter Y/N)'
		goto ErrHndlr
	end
	select @row=count('x') from inserted where 
				not exists (select 'x' from Param where PTyp= 'YN' and PMCd= BRepYN)
	if @row>0 
	Begin
		select @errmsg= 'Bag Repair (Enter Y/N)'
		goto ErrHndlr
	end	
	select @row=count('x') from inserted where not exists (select 'x' from vParam where 
			vPCoCd= BCoCd and vPTyp= 'WORK' and vPMCd= BWrk)
	if @row>0 
	Begin
		select @errmsg= 'Worker Code Not Defined'
		goto ErrHndlr
	end
	select @row=count('x') from inserted where BLstYy<> '' 
		and not exists (select 'x' from Param where PTyp= 'YY' and PMCd= BLstYy)
	if @row>0 
	Begin
		select @errmsg= 'Last Yy Not Defined'
		goto ErrHndlr
	end
	--5717 SubString(str(BLstKey), 4, 7) giving error so changed to SubString(str(BLstKey), 3, 8)
	select @row=count('x') from inserted where 			
			BLstYyKey<> BLstYy+ (case when BLstKey= 0 then '' else SubString(str(BLstKey), 3, 8) end)
	if @row>0 
	Begin
		select @errmsg= 'Last YyKey Not Defined'
		goto ErrHndlr
	end
	select @row= count('x') from inserted where 
			BPrdStg<>'' and not exists (select 'x' from Param where PTyp= 'PRDSTG' and PMCd= BPrdStg)
	if @row> 0 
	Begin
		select @errmsg= 'Invalid Production Stage'
		goto ErrHndlr
	end
	Select @Row= Count('x') From Inserted Ins Where BFgSubLoc<> '' and 
			Not Exists (Select vPMCd From vParam Where vPCoCd= BCoCd and 
				    vPTyp= 'FGSUBLOC' And vPMCd = BFgSubLoc)
  	If @Row> 0
	begin
		Select @ErrMsg= 'Fg Sub Location Not Defined'
		GoTo ErrHndlr
	end
	select @row= count('x') from inserted ins 
		Join OrdDsg on OdPrtKey=ins.BPrtKey and OdIdNo=ins.BOdIdNo and OdCoCd= ins.BCoCd and OdTc= ins.BOdTc 
			and OdYy= ins.BOdYy and OdChr= ins.BOdChr and OdNo= ins.BOdNo and OdSr= ins.BOdSr 
		Join Param On PTyp= 'BCHR' and PMCd= ins.BChr and PSCd= '' and PValue = 'C' 
		where 
			(Select convert(decimal(16,1), sum(Bg.BQty)) from Bag Bg 
				where Bg.BPrtKey=OdPrtKey and Bg.BOdIdNo=OdIdNo and Bg.BCoCd= OdCoCd and 
				Bg.BOdTc= OdTc and Bg.BOdYy= OdYy and Bg.BOdChr= OdChr and Bg.BOdNo= OdNo and 
				Bg.BOdSr= OdSr and Bg.BChr= ins.BChr and 
			(Select PValue from Param where PTyp= 'BCHR' and PMCd= BChr and PSCd= '')= 'C') > 
				convert(decimal(16,1), OdPrdQty, 1) 
	if @row> 0 
	Begin
		select @errmsg= 'Component Bag Qty For This Bag Character Cannot Not Be Greater Than Prd Qty'
		goto ErrHndlr
	end
	/******5717 checks missed are taken from 577 and added to 5716 trigger -- end *****/

	
	if @mode='A'
	begin
	Select @row= count('x') from inserted where 
						not exists (select 'x' from Head where HCoCd= BCoCd)
	if @row> 0 
	begin 
		select @errMsg= 'Company Code Not Valid'
		goto ErrHndlr
	end


	select @row=count('x') from inserted where 
			not exists (select 'x' from Param where PTyp= 'YY' and PMCd= BYy)
	if @row>0 
	Begin
		select @errmsg= 'Bag Year Not Defined'
		goto ErrHndlr
	end

	select @row=count('x') from inserted where 
			not exists (select 'x' from Param where PTyp= 'BCHR' and PMCd= BChr)
	if @row>0 
	Begin
		select @errmsg= 'Bag Character Not Defined'
		goto ErrHndlr
	end

	select @row=count('x') from inserted where BNo<= 0 
	if @row> 0 
	Begin
		select @errmsg= 'Bag Number Should Be Entered'
		goto ErrHndlr
	end

	Select @row=count('x') from inserted where BRepYN='Y' and BRepQty>BQty 
	if @row>0 
	Begin
		select @errmsg= 'Repair Qty must be less than or equal To BQty'
		goto ErrHndlr
	end
 End
	--577chk
	select @row=count('x') from inserted where 
			not exists (select 'x' from OrdDsg where OdPRtKey=BPrtKey and OdIdNo=BOdIdNo and OdCoCd= BCoCd and 
		OdTc= BOdTc and OdYy= BOdYy and OdChr= BOdChr and OdNo= BOdNo and OdSr= BOdSr)
	if @row>0 
	Begin
		select @errmsg='Order Does Not Exist'
		goto ErrHndlr
	end
	
	if @mode in ('A','C')
	begin
	select @row=count('x') from inserted where BQty< 0 
	if @row>0 
	Begin
		select @errmsg= 'Bag Qty Should Be Greater Than Or Equal To Zero'
		goto ErrHndlr
	end

	select @row=count('x') from inserted where BRjQty< 0 
	if @row>0 
	Begin
		select @errmsg= 'Rejection Qty Should Be Greater Than Or Equal To Zero'
		goto ErrHndlr
	end
	
	End
	/******5717 checks missed are taken from 577 and added to 5716 trigger -- start *****/
	if @mode in ('C')
	Begin
		/****** Cannot modify record if BLstYyKey<> '' ******/		
		select @row= count('x') from deleted del where del.BPrtKey=@wCurrPrtn and del.BLstYyKey<> '' and 
			not exists(Select 'x' from inserted ins where ins.BPrtKey=del.BPrtKey and ins.BIdNo=del.BIdNo 
			and ins.BCoCd= del.BCoCd and ins.BYy  = del.BYy and ins.BChr = del.BChr and ins.BNo = del.BNo 
			and ins.BOpnQty = del.BOpnQty and ins.BOpnPtQty =del.BOpnPtQty 
			and ins.BOpnDt =del.BOpnDt and ins.BOpnLoc =del.BOpnLoc 
			and ins.BOpnDesc =del.BOpnDesc 
			)
		if @row> 0 
		begin
			select @errmsg= 'Cannot Modify This Record as Transaction Already Exists'
			goto ErrHndlr
		end
		/****** Cannot modify a closed Bag i.e. when BCls= 'Y' ******/				
/*		select @row= count('x') from deleted del where del.BPrtKey=@wCurrPrtn and del.BCls= 'Y' and del.BRepYN ='N' and
			not exists(Select 'x' from inserted ins where ins.BPrtKey=del.BPrtKey and ins.BIdNo=del.BIdNo 
				and ins.BCoCd= del.BCoCd and ins.BYy  = del.BYy and ins.BChr = del.BChr and ins.BNo = del.BNo and 
				(ins.BCls= 'N'  or (ins.BLoc= del.BLoc 
					and ins.BQty= del.BQty and ins.BPtQty= del.BPtQty and ins.BWrk= del.BWrk 
					and ins.BRjQty= del.BRjQty and ins.BRjPtQty= del.BRjPtQty 
					and ins.BLstYy= del.BLstYy and ins.BLstKey= del.BLstKey 
					and ins.BLstYyKey= del.BLstYyKey)
				) 
			)
		if @row> 0 
		begin
			select @errmsg= 'Cannot Modify This Record as Bag Is Closed'
			goto ErrHndlr
		end*/
	end
	/******5717 checks missed are taken from 577 and added to 5716 trigger -- end *****/

	if @Mode = 'A'
	begin
		Select @row= count('x') from inserted Join Param on PTyp= 'BCHR' and PMCd = BChr 
						where (PValue1<> '' Or PValue2<> '' or PValue3 <> '') and 
							not exists(Select 'x' From OrdRm where OrPrtKey=BPrtKey and OrOdIdNo=BOdIdNo and OrCoCd= BCoCd and OrTc= BOdTc and OrYy= BOdYy 
								and OrChr= BOdChr and OrNo= BOdNo and OrSr= BOdSr 
								and (PValue1= '' Or ','+ PValue1+ ',' like '%,'+ OrRmCtg + ',%' )
								and (PValue2= '' Or ','+ PValue2+ ',' like '%,'+ OrRmSCtg + ',%' ) 
								and (PDesc225='' Or PDesc225=OrSubShp) 
								and (PValue3= '' Or (case when PValue3='WS' then OrWsQty else (case when PValue3='GS' then OrHsQty else 1 end) end)>0)) 										
		if @row> 0 
		begin		
			select @errmsg= 'Order Design Of Bag Does Not Contain Rm Ctg/ SubCtg/ PrdTyp/ SubShp Of Flute Bag Character'
			goto ErrHndlr
		end	
	End

	select @row= count('x') from inserted ins 
		Join OrdDsg on OdPRtKey=ins.BPrtKey and OdIdNo=ins.BOdIdNo and OdCoCd= ins.BCoCd and OdTc= ins.BOdTc 
			and OdYy= ins.BOdYy and OdChr= ins.BOdChr and OdNo= ins.BOdNo and OdSr= ins.BOdSr 
		Join Param On PTyp= 'BCHR' and PMCd= ins.BChr and PSCd= '' and PValue = 'P' 
		where 
			(Select convert(decimal(16,1), sum(Bg.BQty)) from Bag Bg 
				where Bg.BCoCd= OdCoCd and Bg.BOdIdNo=OdIdNo and 
				Bg.BOdTc= OdTc and Bg.BOdYy= OdYy and Bg.BOdChr= OdChr and Bg.BOdNo= OdNo and 
				Bg.BOdSr= OdSr and 
			(Select PValue from Param where PTyp= 'BCHR' and PMCd= BChr and PSCd= '')= 'P') > 
				convert(decimal(16,1), OdPrdQty, 1) 
	if @row> 0 
	Begin
		select @errmsg= 'Primary Bag Qty Cannot Not Be Greater Than Prd Qty' 
		goto ErrHndlr 
	end 

	Select @row= count('x') from inserted ins 
		Join OrdDsg on OdPRtKey=ins.BPrtKey and OdIdNo=ins.BOdIdNo and OdCoCd= ins.BCoCd and OdTc= ins.BOdTc and OdYy= ins.BOdYy 
			and OdChr= ins.BOdChr and OdNo= ins.BOdNo and OdSr= ins.BOdSr 
		Join Param On PTyp= 'BCHR' and PMCd= ins.BChr and PSCd= '' and PValue = 'F' 
		where 
			(Select convert(decimal(16,1), sum(Bg.BQty)) from Bag Bg where Bg.BPrtKey=OdPrtKey 
				and Bg.BOdIdNo=OdIdNo and Bg.BCoCd= OdCoCd and 
				Bg.BOdTc= OdTc and Bg.BOdYy= OdYy and Bg.BOdChr= OdChr and Bg.BOdNo= OdNo and 
				Bg.BOdSr= OdSr and Bg.BChr= ins.BChr and 
			(Select PValue from Param where PTyp= 'BCHR' and PMCd= BChr and PSCd= '')= 'F') > 
				convert(decimal(16,1), OdPrdQty, 1) 
	if @row> 0 
	Begin
		select @errmsg= 'Flute Bag Qty For This Bag Character Cannot Not Be Greater Than Prd Qty'
		goto ErrHndlr
	End
return
ErrHndlr:
	begin
		select @errmsg='|'+@errmsg
		raiserror(@errmsg,16,1)
		rollback transaction 
	end
end


GOGO

SET TRANSACTION ISOLATION LEVEL READ UNCOMMITTED

